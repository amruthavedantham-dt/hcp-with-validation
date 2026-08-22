function runGeminiStructure(companyName, runId) {
  runId = runId || newManualRunId_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName(RAW_EVIDENCE_SHEET_NAME);
  if (!rawSheet) throw new Error('RAW_EVIDENCE sheet not found. Run setupHVTSheet() first.');

  const snippets = getRawEvidenceForCompany_(rawSheet, companyName);
  if (snippets.length === 0) {
    throw new Error('No staged evidence for "' + companyName + '". Run Serper Grounding first.');
  }

  EventLog.info(runId, companyName, '', 'gemini_structure', 'start', 'Gemini structuring started');
  checkForUnknownRegistryDomain_(snippets, companyName, runId);
  const prompt = buildGeminiPrompt_(companyName, snippets);
  const rawDraft = callGemini_(prompt, 'gemini_structure', runId, companyName);
  let draft = validateAndSanitizeDraft_(rawDraft, companyName, runId);
  draft = upgradeToStrongestSourceForPillars_(draft, snippets, companyName, runId);
  draft = checkPillarSourceFidelity_(draft, snippets, companyName, runId);
  draft = resolveAndFilterAlignment_(draft, companyName, runId);
  draft = verifyGapAgainstEvidence_(draft, snippets, companyName, runId);
  draft = checkForMissedMoatSignal_(draft, snippets, companyName, runId);
  draft = checkGapSourceFidelity_(draft, snippets, companyName, runId);

  const rowRange = findOrCreateCompanyRow_(companyName);
  const currentValues = rowRange.getValues()[0];
  const fields = {};

  if (draft.differentiation) {
    fields.differentiation = draft.differentiation.claim || '';
    fields.differentiation_source = draft.differentiation.source_url || '';
    fields.differentiation_credibility = draft.differentiation.credibility || '';
    fields.differentiation_needs_verify = draft.differentiation.needs_verify || '';
  }
  if (draft.moat) {
    fields.moat_type = draft.moat.type || '';
    fields.moat_durability = draft.moat.durability || '';
    fields.moat = draft.moat.claim || '';
    fields.moat_source = draft.moat.source_url || '';
    fields.moat_credibility = draft.moat.credibility || '';
    fields.moat_needs_verify = draft.moat.needs_verify || '';
  }
  if (draft.product_improvement) {
    fields.product_improvement = draft.product_improvement.claim || '';
    fields.product_improvement_source = draft.product_improvement.source_url || '';
    fields.product_improvement_credibility = draft.product_improvement.credibility || '';
    fields.product_improvement_needs_verify = draft.product_improvement.needs_verify || '';
  }

  // unresolved_pillars is computed, never LLM-drafted — a pure function of
  // which pillars this run flagged needs_verify: yes. Zero new failure surface.
  const unresolved = GEMINI_DRAFT_PILLARS.filter(function (p) { return fields[p + '_needs_verify'] === 'yes'; });
  fields.unresolved_pillars = unresolved.join(', ');
  fields.triage_status = unresolved.length === 0 ? TRIAGE_STATUS.AUTO_CONFIRMED : TRIAGE_STATUS.NEEDS_REVIEW;

  draft = resolveProofPillar_(draft, unresolved, companyName, runId);

  // Clear all 3 alignment slots before refilling — otherwise a slot from a
  // PREVIOUS run (when more entries existed) silently persists as stale,
  // orphaned data that no longer traces to anything in the current draft.
  // Found on Eldorado Agritech: a re-run with only 2 valid alignment entries
  // left align3 untouched, still showing a capability from an earlier run's
  // differentiation claim that no longer exists.
  [1, 2, 3].forEach(function (n) {
    fields['align' + n + '_capability'] = '';
    fields['align' + n + '_need'] = '';
    fields['align' + n + '_pull'] = '';
    fields['align' + n + '_source'] = '';
  });

  if (draft.alignment && Array.isArray(draft.alignment)) {
    draft.alignment.slice(0, 3).forEach(function (a, idx) {
      const n = idx + 1;
      fields['align' + n + '_capability'] = a.capability || '';
      fields['align' + n + '_need'] = a.need || '';
      fields['align' + n + '_pull'] = a.pull || '';
      fields['align' + n + '_source'] = a.source || '';
    });
  }

  if (draft.gap) {
    fields.gap_latent = draft.gap.latent || '';
    fields.gap_currently_monetized = draft.gap.currently_monetized || '';
    fields.gap_summary = draft.gap.summary || '';
    fields.gap_activation_path = draft.gap.activation_path || '';
    fields.gap_source = draft.gap.source || '';
  }

  if (draft.proof) {
    fields.proof_type = draft.proof.type || '';
    fields.proof_status = draft.proof.status || '';
    fields.proof_source = draft.proof.source || '';
    fields.proof_pillar = draft.proof.pillar || '';
  }

  // A prior human review of a pillar applied to the OLD claim text. If the
  // new draft changed that text, the old Verified/Rejected status no longer
  // means anything — reset it to Pending rather than let it silently point
  // at content nobody actually checked.
  const PILLAR_CLAIM_COLUMNS = { differentiation: 'differentiation', moat: 'moat', product_improvement: 'product_improvement' };
  Object.keys(PILLAR_CLAIM_COLUMNS).forEach(function (pillarKey) {
    const claimCol = PILLAR_CLAIM_COLUMNS[pillarKey];
    if (fields[claimCol] === undefined) return;
    const oldClaim = String(currentValues[HVT_HEADERS.indexOf(claimCol)] || '').trim();
    const newClaim = String(fields[claimCol] || '').trim();
    if (oldClaim !== newClaim) {
      fields[pillarKey + '_review'] = 'Pending';
      Logger.log('INFO: ' + pillarKey + ' claim changed for "' + companyName + '" — resetting ' + pillarKey + '_review to Pending');
    }
  });

  setRowFields_(rowRange, fields);
  EventLog.info(runId, companyName, '', 'gemini_structure', 'done', 'triage: ' + fields.triage_status + (unresolved.length ? ' (' + fields.unresolved_pillars + ')' : ''));
  return draft;
}

// 5th pipeline step — a one-paragraph prose recap of the already-drafted
// row, restatement not new analysis, so a reviewer working the REVIEW_QUEUE
// doesn't have to read 15+ columns to get oriented. Leads with the triage
// status so the summary doubles as the queue's row description.
function runCompanySummary(companyName, runId) {
  runId = runId || newManualRunId_();
  const rowRange = findOrCreateCompanyRow_(companyName);
  const values = rowRange.getValues()[0];
  const get = function (colName) { return String(values[HVT_HEADERS.indexOf(colName)] || '').trim(); };

  const triageStatus = get('triage_status') || TRIAGE_STATUS.NEEDS_REVIEW;
  const unresolved = get('unresolved_pillars');

  const rowText = [
    'differentiation: ' + (get('differentiation') || '(none)'),
    'moat: ' + (get('moat') || '(none)'),
    'product_improvement: ' + (get('product_improvement') || '(none)'),
    'alignment 1: ' + (get('align1_capability') || '(none)') + ' -> ' + (get('align1_need') || ''),
    'alignment 2: ' + (get('align2_capability') || '(none)') + ' -> ' + (get('align2_need') || ''),
    'alignment 3: ' + (get('align3_capability') || '(none)') + ' -> ' + (get('align3_need') || ''),
    'gap: ' + (get('gap_summary') || '(none)')
  ].join('\n');

  const prompt = [
    'Write ONE short prose paragraph (3-5 sentences, no markdown, no headers) summarizing this company\'s wealth-creation-engine characterization below. ' +
    'This is a RESTATEMENT of what is already drafted, not new analysis — do not add any fact not present in the row content below.',
    triageStatus === TRIAGE_STATUS.NEEDS_REVIEW
      ? 'Start the paragraph with "NEEDS REVIEW (' + unresolved + '): " then continue the summary.'
      : 'Start the paragraph with "AUTO-CONFIRMED: " then continue the summary.',
    '',
    'ROW CONTENT:',
    rowText,
    '',
    'Respond with ONLY the paragraph text, no JSON, no quotes.'
  ].join('\n');

  const apiKey = getScriptProp_('GEMINI_API_KEY');
  const model = PropertiesService.getScriptProperties().getProperty('GEMINI_MODEL') || 'gemini-2.5-flash';
  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + model + ':generateContent?key=' + apiKey;

  if (CircuitBreaker.isHalted('GEMINI')) {
    EventLog.warn(runId, companyName, '', 'company_summary', 'skipped', 'GEMINI circuit breaker halted — summary skipped');
    return '';
  }
  rateLimit('GEMINI');

  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({
      contents: [{ parts: [{ text: prompt }] }],
      generationConfig: { temperature: 0, thinkingConfig: { thinkingBudget: 0 } }
    }),
    muteHttpExceptions: true
  });

  const data = JSON.parse(response.getContentText());
  if (response.getResponseCode() !== 200) {
    CircuitBreaker.recordFailure('GEMINI');
    rateLimitBackoff('GEMINI');
    if (data.usageMetadata) TokenLog.logGemini(runId, companyName, 'company_summary', model, data.usageMetadata);
    EventLog.error(runId, companyName, '', 'company_summary', 'error', 'Gemini call failed: ' + response.getContentText());
    return '';
  }
  CircuitBreaker.recordSuccess('GEMINI');
  rateLimitSuccess('GEMINI');
  TokenLog.logGemini(runId, companyName, 'company_summary', model, data.usageMetadata);

  const text = (data.candidates && data.candidates[0] && data.candidates[0].content.parts[0].text || '').trim();
  setRowFields_(rowRange, { company_summary: text });
  EventLog.info(runId, companyName, '', 'company_summary', 'done', 'summary written (' + text.length + ' chars)');
  return text;
}

function getRawEvidenceForCompany_(rawSheet, companyName) {
  const lastRow = rawSheet.getLastRow();
  if (lastRow <= 1) return [];
  const values = rawSheet.getRange(2, 1, lastRow - 1, RAW_EVIDENCE_HEADERS.length).getValues();
  return values
    .filter(function (row) { return String(row[0]).trim().toLowerCase() === companyName.trim().toLowerCase(); })
    .map(function (row) {
      return { pillar: row[1], query: row[2], source_url: row[3], title: row[4], snippet: row[5], full_text: row[6] };
    });
}

// Only these three pillars are drafted by Gemini. `identity` and
// `alignment_gap` snippets stay staged in RAW_EVIDENCE for manual use.
const GEMINI_DRAFT_PILLARS = ['differentiation', 'moat', 'product_improvement'];

function buildGeminiPrompt_(companyName, snippets) {
  const relevant = snippets.filter(function (s) { return EVIDENCE_INPUT_TAGS.indexOf(s.pillar) !== -1; });
  const lines = relevant.map(function (s, i) {
    const body = s.full_text ? s.full_text : s.snippet;
    // Serper's /patents endpoint is unreliable in testing — timeouts, empty
    // results, and confident-looking but unrelated patents (matching only on
    // keyword overlap, not actual assignee) have all been observed. A plain
    // "informational only" tag wasn't enough context for this one specifically.
    const tagNote = s.pillar === 'patents'
      ? ' — informational only, NOT authoritative, see rule 3. WARNING: this patents search endpoint has returned unrelated results in testing (keyword-matched, not assignee-matched) — only use this snippet if its content explicitly names ' + companyName + ' or one of its known products, not just topical overlap'
      : ' — informational only, NOT authoritative, see rule 3';
    return (i + 1) + '. [search tag: ' + s.pillar + tagNote + '] URL: ' + s.source_url + '\n   Title: ' + s.title + '\n   Content: ' + body;
  }).join('\n');

  return [
    'You are structuring evidence for "' + companyName + '" into a wealth-creation-engine characterization: three pillar claims, up to 3 alignment entries, an execution gap, and a proof pointer.',
    '',
    'STRICT RULES:',
    '1. Only use facts present in the snippets below. Never add a fact that is not in a snippet.',
    '2. Every claim you propose must carry the exact source_url copied verbatim from the snippet that supports it. If no snippet supports a field, leave it as an empty string (or omit the alignment entry / leave gap fields empty) rather than inventing one.',
    '3. Each snippet carries a "search tag" showing which query surfaced it — that tag reflects search strategy only and is NOT authoritative for classification. Classify strictly by these definitions instead:',
    '   - differentiation: what makes the OFFERING itself non-commodity — cornered resource, niche/scale leadership, manufacturing flexibility (HMLV-at-scale), a named proprietary technology/process, design library (switching cost), integration, installed base, brand.',
    '   - moat: a barrier that protects differentiation over time — see rule 4 for the exact type to pick.',
    '   - product_improvement: evidence of the R&D ENGINE itself — DSIR/registry R&D recognition, patents, NPD pipeline, history of product evolution. Patents and DSIR recognition belong here, not in differentiation or moat, even if their search tag says otherwise. Example: "has 100+ patents" alone is product_improvement evidence, NOT a moat entry — do not create a moat.claim for a bare patent/technology count. Only populate moat.claim/type/durability if the evidence explicitly describes resulting lock-in (e.g. "sole licensee of X", "competitors cannot replicate the process"). If that specific lock-in language is absent, leave moat.claim, moat.type, and moat.durability all as empty strings. IMPORTANT: simply restating the range of products a company manufactures (e.g. "manufactures a wide range of X, Y, and Z equipment") is NOT product_improvement evidence — that is a product catalog, not evidence of an R&D engine. Only draft this pillar if a snippet explicitly shows an R&D-specific signal: a named R&D facility/center, DSIR/registry recognition, a patent count, or a described history of new product development. If no snippet shows this, leave product_improvement.claim as an empty string rather than repackaging the product list from differentiation.',
    '4. moat.type must be exactly one of these seven — choose whichever definition actually fits the evidence, do not default to qualification-lock-in just because it is the most commonly seen: switching-cost (cost/effort a customer would incur to switch away, beyond any certification — e.g. embedded tooling, long-term integration); qualification-lock-in (customer-facing certifications/approvals that gate who can even supply — JSS, AS9100D, IATF, sole-supplier approval, GOTS; generic export certs like BRC/FSSC/Kosher are table-stakes, NOT this); IP (patents/trade secrets that legally or practically block replication — only when the evidence shows resulting exclusivity, not just "has patents", which is product_improvement per rule 3); brand (reputation/trust that itself commands preference or pricing power); integration (vertical/horizontal integration making the full value chain hard to replicate); scale (advantage from being uniquely large, first, or backed at a scale competitors cannot quickly replicate — e.g. a sole large government-backed facility); network (value increases as more participants join — rare in manufacturing, mostly platforms). moat.durability must be exactly one of: ' + MOAT_DURABILITY.join(', ') + '. If unclear from the snippets, leave both type and durability as empty strings rather than guessing.',
    '5. credibility must be exactly one of: ' + CREDIBILITY_TIERS.join(', ') + '. Default to "self-claimed". A source hosted on YouTube, LinkedIn, Facebook, X/Twitter, Instagram, Medium, or a wire service (PR Newswire, Business Wire, GlobeNewswire) is NOT automatically third-party — it is still the company\'s own words unless the content shows a named independent journalist, analyst, or publication doing their own reporting or investigation (not just interviewing/quoting the company). A snippet tagged "news" is more likely to be genuine third-party reporting — check it first for that tier. "qualification-gated" or "registry-confirmed" require the source itself to BE the certifying/registry body (e.g. a dsir.gov.in or ipindia.gov.in page) — a company\'s own website describing a certification it holds is "self-claimed", not "qualification-gated".',
    '6. If a claim would need a registry to confirm (DSIR, patents, MCA, sector certification bodies), set needs_verify to "yes" and do not claim it is confirmed. Otherwise needs_verify is "no".',
    '7. Never output "commodity" or any absence-of-differentiation finding for differentiation unless a snippet explicitly states an absence of differentiation. Silence in the snippets is not evidence of commodity — if unclear, leave the claim empty rather than guessing.',
    '8. Write every claim, alignment capability/need, and gap field as a short, analytical, fact-first phrase — aim for under 15 words. Never restate the company name (redundant with the row it\'s written into). Never use marketing language ("showcases", "designed to help", "redefining", "state-of-the-art"). Write like this: "India #1 in electromechanical relays/switches/EV-contactors; custom assemblies" or "design-in + qualification lock-in (JSS/AS9100D/IATF)" — not like a press release sentence.',
    '9. alignment: up to 3 entries, each {capability, need, pull}. capability MUST be the actual text (or a close paraphrase) of a claim you already drafted above in differentiation, moat, or product_improvement — never introduce a capability that was not drafted into one of those three; entries that do not match will be discarded automatically, so there is no benefit to inventing one. If a fact seems alignment-worthy but you did not draft it into any of the three, go back and check whether it should have been first — in particular, a first-mover, government-backed, or uniquely-scaled barrier is a "scale"-type moat (rule 4). need is a market/demand trend found in "alignment_gap" or "news"-tagged snippets. pull must be exactly one of: ' + PULL_STRENGTHS.join(', ') + '. Do not include a source field — the source for each entry is taken automatically from whichever pillar its capability matches. Only include an entry where the snippets actually connect a capability to a demand signal — do not invent generic alignment. If none is supported, return an empty array.',
    '10. gap: {latent, currently_monetized, summary, activation_path, source} — latent is what the capability could do per the evidence, currently_monetized is what is actually being captured today, summary is one line, activation_path is one line only if the evidence suggests a concrete next step, source is the exact source_url backing this analysis. Leave any field empty if unsupported — do NOT infer financial details (cash flow, profitability, revenue) unless a snippet explicitly states them. Example: receiving a government grant or financial assistance for a project is NOT evidence of cash flow pressure or financial distress — do not draw that inference. Only state a financial detail if a snippet directly says so in those terms. Before writing activation_path, check whether the underlying fact should have been drafted as a moat entry first (rule 4) — a government approval/grant/financial-assistance fact for a unique or first-of-its-kind facility is "scale"-type moat evidence, not a gap fact. If it qualifies as a moat, draft it in moat, and activation_path should describe what happens NEXT with that already-drafted moat (e.g. "scale up the [moat] to capture more value"), not restate the barrier itself as if it only belonged in gap. A gap requires genuine tension between what the evidence shows is POSSIBLE and what it shows is ACTUALLY captured — latent and currently_monetized must describe two different things, not the same fact worded three ways. Example of what NOT to do: if a snippet only says "is a net contributor of green energy to the grid," that describes current state only — it is not a gap, because there is no accompanying evidence of untapped capacity, limited scope, or unrealized potential (e.g. "exports only 10% of generated capacity" or "capability exists but is not yet offered commercially" would be genuine gap evidence). If the snippets do not show that contrast, leave latent, currently_monetized, summary, and activation_path all empty rather than restating the one fact you do have across all four fields.',
    '11. proof: {type, status, source, pillar} summarizes the single most important unresolved pillar. status must be exactly one of: ' + PROOF_STATUS_DRAFTABLE.join(', ') + ' — NEVER "confirmed" or "registry-access-blocked" (those are only set by a human after an actual registry check, never by you). type is the registry that would resolve it (e.g. "DSIR", "Patent", "MCA"). source is the strongest source_url backing the claimed status. pillar must be exactly one of: ' + GEMINI_DRAFT_PILLARS.join(', ') + ' — whichever pillar this proof entry is actually about. If only one pillar above has needs_verify "yes", pillar must be that one.',
    '',
    'SNIPPETS:',
    lines,
    '',
    'Respond with ONLY this JSON shape (no markdown, no commentary):',
    '{',
    '  "differentiation": {"claim": "", "source_url": "", "credibility": "", "needs_verify": ""},',
    '  "moat": {"type": "", "durability": "", "claim": "", "source_url": "", "credibility": "", "needs_verify": ""},',
    '  "product_improvement": {"claim": "", "source_url": "", "credibility": "", "needs_verify": ""},',
    '  "alignment": [{"capability": "", "need": "", "pull": ""}],',
    '  "gap": {"latent": "", "currently_monetized": "", "summary": "", "activation_path": "", "source": ""},',
    '  "proof": {"type": "", "status": "", "source": "", "pillar": ""}',
    '}'
  ].join('\n');
}

// Logs a warning to both Logger.log (script-editor-only) and EventLog.warn
// (user-visible in the sheet) — every diagnostic in this file routes through
// this instead of calling Logger.log directly, so nothing important is only
// visible to someone who opens the script editor. runId/companyName may be
// undefined (e.g. Test.js's direct calls) — EventLog handles that quietly.
function logWarning_(runId, companyName, stage, message) {
  Logger.log('WARNING: ' + message);
  if (typeof EventLog !== 'undefined') EventLog.warn(runId, companyName, '', stage, 'warning', message);
}

function logInfo_(runId, companyName, stage, message) {
  Logger.log('INFO: ' + message);
  if (typeof EventLog !== 'undefined') EventLog.info(runId, companyName, '', stage, 'info', message);
}

// Deterministic safety net, applied regardless of how well the prompt was
// followed. This is what makes the fix permanent across model versions and
// companies — it doesn't rely on Gemini remembering the rules every time.
function validateAndSanitizeDraft_(draft, companyName, runId) {
  function sanitizePillar(obj, label) {
    if (!obj) return obj;

    const claimText = String(obj.claim || '').trim();

    // No claim → nothing to carry a credibility tier, a verify flag, or a
    // source. Without this, a pillar can end up with needs_verify: "yes" and
    // no claim at all — internally inconsistent, and it wrongly shows up in
    // unresolved_pillars as if there were something to check.
    if (claimText === '') {
      obj.credibility = '';
      obj.needs_verify = '';
      obj.source_url = '';
      return obj;
    }

    if (obj.credibility && CREDIBILITY_TIERS.indexOf(obj.credibility) === -1) {
      logWarning_(runId, companyName, 'gemini_structure', label + '.credibility "' + obj.credibility + '" not in enum — defaulting to self-claimed');
      obj.credibility = 'self-claimed';
    }

    if (obj.source_url && isSelfPublishedDomain_(obj.source_url) && obj.credibility !== 'self-claimed') {
      logWarning_(runId, companyName, 'gemini_structure', label + ' source ' + obj.source_url + ' is a self-published platform — forcing credibility to self-claimed (model said "' + obj.credibility + '")');
      obj.credibility = 'self-claimed';
      obj.needs_verify = 'yes';
    }

    if ((obj.credibility === 'qualification-gated' || obj.credibility === 'registry-confirmed') &&
        !(obj.source_url && isOfficialRegistryDomain_(obj.source_url))) {
      logWarning_(runId, companyName, 'gemini_structure', label + ' claimed "' + obj.credibility + '" but source ' + obj.source_url + ' is not an official registry domain — forcing to self-claimed');
      obj.credibility = 'self-claimed';
      obj.needs_verify = 'yes';
    }

    // Deterministic backstop for rule 6 — catches a registry-triggering word
    // (patent, DSIR, MCA, named certs) sitting inside the claim text even
    // when the model didn't flag needs_verify for it.
    const lowerClaim = claimText.toLowerCase();
    const matchedKeywords = REGISTRY_TRIGGER_KEYWORDS.filter(function (kw) { return lowerClaim.indexOf(kw) !== -1; });
    if (matchedKeywords.length > 0 && obj.needs_verify !== 'yes') {
      logWarning_(runId, companyName, 'gemini_structure', label + ' claim mentions registry keyword(s) [' + matchedKeywords.join(', ') + '] but needs_verify was "' + obj.needs_verify + '" — forcing to yes');
      obj.needs_verify = 'yes';
    }

    // Same backstop for market-exclusivity/superlative claims ("sole
    // supplier", "world's largest", etc.) — these need independent
    // verification via industry reports/competitor data, not a registry, but
    // they're still an unconfirmed self-claim until checked.
    const matchedExclusivity = EXCLUSIVITY_TRIGGER_KEYWORDS.filter(function (kw) { return lowerClaim.indexOf(kw) !== -1; });
    if (matchedExclusivity.length > 0 && obj.needs_verify !== 'yes') {
      logWarning_(runId, companyName, 'gemini_structure', label + ' claim mentions exclusivity keyword(s) [' + matchedExclusivity.join(', ') + '] but needs_verify was "' + obj.needs_verify + '" — forcing to yes');
      obj.needs_verify = 'yes';
    }

    if (obj.needs_verify !== 'yes' && obj.needs_verify !== 'no') {
      logWarning_(runId, companyName, 'gemini_structure', label + '.needs_verify "' + obj.needs_verify + '" invalid — defaulting to yes');
      obj.needs_verify = 'yes';
    }

    // If credibility genuinely IS qualification-gated/registry-confirmed at
    // this point, the source already passed the official-registry-domain
    // check above — meaning the registry check rule 6 asks for has, by
    // definition, already happened via that source. needs_verify: yes here
    // would be self-contradictory (found on Eldorado: product_improvement
    // came back registry-confirmed + needs_verify: yes together, once its
    // source domain was added to the allowlist). This runs LAST, after the
    // keyword backstops above, so a genuine registry source overrides any
    // keyword-based suspicion rather than the other way around.
    if ((obj.credibility === 'qualification-gated' || obj.credibility === 'registry-confirmed') && obj.needs_verify !== 'no') {
      logInfo_(runId, companyName, 'gemini_structure', label + ' credibility is ' + obj.credibility + ' but needs_verify was "' + obj.needs_verify + '" — forcing to no, contradictory otherwise');
      obj.needs_verify = 'no';
    }

    return obj;
  }

  draft.differentiation = sanitizePillar(draft.differentiation, 'differentiation');
  draft.product_improvement = sanitizePillar(draft.product_improvement, 'product_improvement');
  draft.moat = sanitizePillar(draft.moat, 'moat');

  if (draft.moat) {
    if (String(draft.moat.claim || '').trim() === '') {
      draft.moat.type = '';
      draft.moat.durability = '';
    } else {
      if (draft.moat.type && MOAT_TYPES.indexOf(draft.moat.type) === -1) {
        logWarning_(runId, companyName, 'gemini_structure', 'moat.type "' + draft.moat.type + '" not in enum — clearing');
        draft.moat.type = '';
      }
      if (draft.moat.durability && MOAT_DURABILITY.indexOf(draft.moat.durability) === -1) {
        logWarning_(runId, companyName, 'gemini_structure', 'moat.durability "' + draft.moat.durability + '" not in enum — clearing');
        draft.moat.durability = '';
      }
    }
  }

  if (draft.alignment && Array.isArray(draft.alignment)) {
    draft.alignment.forEach(function (a) {
      if (a.pull && PULL_STRENGTHS.indexOf(a.pull) === -1) {
        logWarning_(runId, companyName, 'gemini_structure', 'alignment.pull "' + a.pull + '" not in enum — clearing');
        a.pull = '';
      }
    });
  }

  if (draft.proof && draft.proof.status && PROOF_STATUS_DRAFTABLE.indexOf(draft.proof.status) === -1) {
    logWarning_(runId, companyName, 'gemini_structure', 'proof.status "' + draft.proof.status + '" is not a draftable state (only not-found/claimed-verify — confirmed/registry-access-blocked are human-only) — forcing to not-found');
    draft.proof.status = 'not-found';
  }

  return draft;
}

function normalizeForMatch_(text) {
  return String(text || '').toLowerCase().replace(/[^a-z0-9\s]/g, '').replace(/\s+/g, ' ').trim();
}

function wordOverlapRatio_(a, b) {
  const wordsA = a.split(' ').filter(function (w) { return w.length > 2; });
  const wordsB = b.split(' ').filter(function (w) { return w.length > 2; });
  if (wordsA.length === 0 || wordsB.length === 0) return 0;
  const setB = {};
  wordsB.forEach(function (w) { setB[w] = true; });
  const shared = wordsA.filter(function (w) { return setB[w]; }).length;
  return shared / Math.min(wordsA.length, wordsB.length);
}

// Returns which pillar (if any) an alignment capability actually traces to —
// substring containment catches truncation/expansion, word-overlap catches
// paraphrasing. Returns null if it matches none of the three.
function matchAlignmentCapabilityToPillar_(capabilityText, draft) {
  const normalizedCap = normalizeForMatch_(capabilityText);
  if (!normalizedCap) return null;

  const pillars = [
    { key: 'differentiation', obj: draft.differentiation },
    { key: 'moat', obj: draft.moat },
    { key: 'product_improvement', obj: draft.product_improvement }
  ];

  for (let i = 0; i < pillars.length; i++) {
    const obj = pillars[i].obj;
    if (!obj || !obj.claim) continue;
    const normalizedClaim = normalizeForMatch_(obj.claim);
    if (!normalizedClaim) continue;
    if (normalizedClaim.indexOf(normalizedCap) !== -1 || normalizedCap.indexOf(normalizedClaim) !== -1) {
      return pillars[i].key;
    }
    if (wordOverlapRatio_(normalizedCap, normalizedClaim) >= 0.5) {
      return pillars[i].key;
    }
  }
  return null;
}

// A capability that isn't traceable to a drafted pillar claim gets dropped
// rather than written — three straight rounds of "capability must reference
// a drafted pillar" as a plain instruction failed to hold. This makes it
// unconditional. The matched pillar's OWN source_url is then forced onto the
// entry (never the model's own guess) — this is also what stops two
// unrelated entries from sharing a generic fallback source: they can only
// share a source if they genuinely trace to the same pillar.
function resolveAndFilterAlignment_(draft, companyName, runId) {
  if (!draft.alignment || !Array.isArray(draft.alignment)) return draft;

  // Drop entries whose capability is a near-duplicate of an already-kept
  // entry's capability — found on Eldorado Agritech: two entries repeated
  // the IDENTICAL capability text ("Offers hybrid seeds and crop care
  // products; wide network of over 2500 dealers.") word-for-word. Dedup by
  // capability TEXT, not by which pillar it traces to — a single pillar's
  // claim can legitimately bundle several distinct facts, each deserving its
  // own alignment entry (Jaishil: "OMRI listed organic products" and
  // "NutriZin 700 zinc fertilizer" are two different facts, both substrings
  // of one combined differentiation claim, and both correctly kept — word
  // overlap between them is low, ~0%, unlike Eldorado's identical repeat).
  // Garware's two entries (differentiation-derived, ~33% overlap) are the
  // same legitimate pattern as Jaishil's, not a duplicate — deduping by
  // pillar alone would have wrongly dropped one of them.
  const keptCapabilities = [];

  draft.alignment = draft.alignment.filter(function (a) {
    const matchedPillar = matchAlignmentCapabilityToPillar_(a.capability, draft);
    if (!matchedPillar) {
      logWarning_(runId, companyName, 'gemini_structure', 'alignment entry dropped for "' + companyName + '" — capability "' + a.capability + '" does not match any drafted pillar claim');
      return false;
    }

    const normalizedCap = normalizeForMatch_(a.capability);
    const isDuplicate = keptCapabilities.some(function (kept) { return wordOverlapRatio_(normalizedCap, kept) >= 0.7; });
    if (isDuplicate) {
      logWarning_(runId, companyName, 'gemini_structure', 'alignment entry dropped for "' + companyName + '" — capability "' + a.capability + '" duplicates an already-kept entry\'s capability');
      return false;
    }

    keptCapabilities.push(normalizedCap);
    a.source = draft[matchedPillar].source_url || '';
    return true;
  });

  return draft;
}

// proof describes exactly one pillar (rule 11), but nothing on the row said
// WHICH one once unresolved_pillars lists more than one. Can't derive this
// from proof.source alone — pillars often share the same source_url (e.g.
// Natesan's differentiation_source and moat_source are both natesan.co.in),
// so matching by URL would silently pick the wrong pillar whenever two
// pillars cite the same page. Instead: when there's exactly one unresolved
// pillar, there's no ambiguity to leave to the model — force proof.pillar to
// that one regardless of what was drafted. Only trust the model's own
// proof.pillar when there's real ambiguity (0 or 2+ unresolved), and even
// then validate it's actually a member of that list.
function resolveProofPillar_(draft, unresolvedPillars, companyName, runId) {
  if (!draft.proof) return draft;

  if (draft.proof.pillar && GEMINI_DRAFT_PILLARS.indexOf(draft.proof.pillar) === -1) {
    logWarning_(runId, companyName, 'gemini_structure', 'proof.pillar "' + draft.proof.pillar + '" for "' + companyName + '" not in enum — clearing');
    draft.proof.pillar = '';
  }

  if (unresolvedPillars.length === 1) {
    if (draft.proof.pillar && draft.proof.pillar !== unresolvedPillars[0]) {
      logInfo_(runId, companyName, 'gemini_structure', 'proof.pillar for "' + companyName + '" overridden from "' + draft.proof.pillar + '" to "' + unresolvedPillars[0] + '" — only one pillar is unresolved, no ambiguity to leave to the model');
    }
    draft.proof.pillar = unresolvedPillars[0];
  } else if (unresolvedPillars.length > 1 && draft.proof.pillar && unresolvedPillars.indexOf(draft.proof.pillar) === -1) {
    logWarning_(runId, companyName, 'gemini_structure', 'proof.pillar "' + draft.proof.pillar + '" for "' + companyName + '" is not among the unresolved pillars [' + unresolvedPillars.join(', ') + '] — clearing, check manually');
    draft.proof.pillar = '';
  }

  return draft;
}

// If the evidence pool for a pillar contains a snippet from an actual
// official registry domain, prefer that as the pillar's source over whatever
// Gemini picked (usually the company's own site) — found on Natesan: the
// real DSIR "Directory of Recognised In-house R&D Units" PDF (a .gov.in
// domain) was sitting in the product_improvement evidence pool the whole
// time, but the claim cited natesan.co.in instead, leaving it self-claimed
// when it could have been registry-confirmed. Only triggers on evidence
// already tagged to the SAME pillar's search group.
//
// Requires a content-overlap check before swapping (added after Garware):
// a registry hit tagged to the right pillar is NOT automatically evidence
// FOR this specific claim — Garware's differentiation claim ("leading
// player... aquaculture, fishing, sports, safety, agricultural nets, coated
// fabrics, ropes, geosynthetics") got upgraded to
// indiascienceandtechnology.gov.in, a bare directory listing ("Institute
// Category: Private Sector - DSIR Registered") that says nothing about any
// of that. The registry domain was real, the pillar tag matched, but the
// content didn't — the claim was actually built from garwarefibres.com. A
// mismatched "registry-confirmed" is worse than self-claimed: it looks more
// trustworthy while being just as unverified for the specific claim text,
// and combined with auto-triage (needs_verify forced to "no" here), a wrong
// upgrade could mark a row auto-confirmed when nothing about the claim was
// actually checked.
function upgradeToStrongestSourceForPillars_(draft, snippets, companyName, runId) {
  GEMINI_DRAFT_PILLARS.forEach(function (pillarKey) {
    const obj = draft[pillarKey];
    if (!obj || !obj.claim) return;
    if (obj.source_url && isOfficialRegistryDomain_(obj.source_url)) return;

    const claimText = normalizeForMatch_(obj.claim);
    const registryHit = snippets.filter(function (s) {
      if (s.pillar !== pillarKey || !s.source_url || !isOfficialRegistryDomain_(s.source_url)) return false;
      const registryText = normalizeForMatch_(s.full_text || s.snippet || '');
      return wordOverlapRatio_(claimText, registryText) >= 0.2;
    })[0];

    if (registryHit) {
      logInfo_(runId, companyName, 'gemini_structure', pillarKey + ' for "' + companyName + '" upgraded from ' + (obj.source_url || '(none)') + ' to registry source ' + registryHit.source_url);
      obj.source_url = registryHit.source_url;
      obj.credibility = 'registry-confirmed';
      obj.needs_verify = 'no';
    } else {
      const hadCandidate = snippets.some(function (s) { return s.pillar === pillarKey && s.source_url && isOfficialRegistryDomain_(s.source_url); });
      if (hadCandidate) {
        logWarning_(runId, companyName, 'gemini_structure', pillarKey + ' for "' + companyName + '" had a registry-domain snippet available but its content did not actually overlap with the claim — kept the original source instead of a mismatched upgrade. Check manually whether the registry evidence should be used differently.');
      }
    }
  });

  return draft;
}

// Diagnostic only, same posture and mechanism as checkGapSourceFidelity_ — a
// pillar claim can quietly combine facts from two different sources while
// citing only one (found on Eldorado Agritech: product_improvement's
// "188.16 acres" detail actually came from the company's own site, but the
// claim cited only the separate ICRA PDF that supplied the DSIR-recognition
// half of the same sentence). Runs AFTER upgradeToStrongestSourceForPillars_
// so it checks the final source_url that will actually be written to the
// sheet, not a pre-upgrade one. Threshold is lower than gap's (20% vs 30%)
// because rule 8 asks for short, compressed claims (under 15 words) — a
// terse paraphrase naturally shares fewer words with its source than a
// longer gap sentence would, so gap's threshold would over-fire here.
// Warning-only: never clears a claim, so this cannot change what ends up in
// the sheet, only what shows up in the log for manual review.
function checkPillarSourceFidelity_(draft, snippets, companyName, runId) {
  GEMINI_DRAFT_PILLARS.forEach(function (pillarKey) {
    const obj = draft[pillarKey];
    if (!obj || !obj.claim || !obj.source_url) return;

    const claimText = normalizeForMatch_(obj.claim);
    if (!claimText) return;

    const sourceEvidence = snippets.filter(function (s) { return s.source_url === obj.source_url; })[0];
    const sourceText = normalizeForMatch_(sourceEvidence ? (sourceEvidence.full_text || sourceEvidence.snippet || '') : '');

    const overlap = wordOverlapRatio_(claimText, sourceText);
    if (overlap < 0.2) {
      logWarning_(runId, companyName, 'gemini_structure', pillarKey + ' for "' + companyName + '" has low word overlap (' +
        Math.round(overlap * 100) + '%) with its cited source (' + obj.source_url + ') — claim may combine ' +
        'details from more than one source. Check manually before trusting this as fully backed by the cited URL.');
    }
  });

  return draft;
}

// Cross-checks gap's financial-sounding language against what its own cited
// source actually says. A targeted "don't infer distress from a grant" rule
// didn't stop this fabrication across four straight runs — this doesn't rely
// on the model's reasoning at all, it just checks the source's own text.
function verifyGapAgainstEvidence_(draft, snippets, companyName, runId) {
  if (!draft.gap || !draft.gap.source) return draft;

  const gapText = (String(draft.gap.currently_monetized || '') + ' ' + String(draft.gap.summary || '')).toLowerCase();
  const mentionsFinancial = GAP_FINANCIAL_KEYWORDS.some(function (kw) { return gapText.indexOf(kw) !== -1; });
  if (!mentionsFinancial) return draft;

  const sourceEvidence = snippets.filter(function (s) { return s.source_url === draft.gap.source; })[0];
  const evidenceText = sourceEvidence ? String(sourceEvidence.full_text || sourceEvidence.snippet || '').toLowerCase() : '';
  const evidenceSupportsFinancial = GAP_FINANCIAL_KEYWORDS.some(function (kw) { return evidenceText.indexOf(kw) !== -1; });

  if (!evidenceSupportsFinancial) {
    logWarning_(runId, companyName, 'gemini_structure', 'gap for "' + companyName + '" uses financial-distress language not found in its cited source (' + draft.gap.source + ') — clearing gap_currently_monetized/gap_summary as likely fabricated');
    draft.gap.currently_monetized = '';
    draft.gap.summary = '';
  }

  return draft;
}

// Diagnostic only — never auto-fills moat, that judgment call stays with the
// analyst. But when moat comes up empty, this makes the reason explicit
// instead of leaving you guessing: it scans gap's own text and the
// moat-tagged evidence pool for scale/government-backing signals that
// suggest a moat entry was available but got used elsewhere (e.g. Gemini
// wrote gap_activation_path from a fact instead of drafting it as moat).
function checkForMissedMoatSignal_(draft, snippets, companyName, runId) {
  if (!draft.moat || String(draft.moat.claim || '').trim() !== '') return draft;

  const gapText = draft.gap
    ? [draft.gap.latent, draft.gap.currently_monetized, draft.gap.summary, draft.gap.activation_path].join(' ').toLowerCase()
    : '';
  const moatTaggedText = snippets
    .filter(function (s) { return s.pillar === 'moat'; })
    .map(function (s) { return (s.title || '') + ' ' + (s.full_text || s.snippet || ''); })
    .join(' ')
    .toLowerCase();

  const hitInGap = MOAT_SCALE_SIGNAL_KEYWORDS.filter(function (kw) { return gapText.indexOf(kw) !== -1; });
  const hitInMoatEvidence = MOAT_SCALE_SIGNAL_KEYWORDS.filter(function (kw) { return moatTaggedText.indexOf(kw) !== -1; });

  if (hitInGap.length > 0) {
    logWarning_(runId, companyName, 'gemini_structure', 'moat is empty for "' + companyName + '", but gap text contains scale-moat signal(s) [' + hitInGap.join(', ') + '] — this fact likely should have been drafted as a "scale"-type moat entry instead. Check gap_activation_path/gap_latent manually.');
  } else if (hitInMoatEvidence.length > 0) {
    logWarning_(runId, companyName, 'gemini_structure', 'moat is empty for "' + companyName + '", but moat-tagged evidence contains scale-moat signal(s) [' + hitInMoatEvidence.join(', ') + '] that were not drafted into any field — check RAW_EVIDENCE rows tagged "moat" manually.');
  }

  return draft;
}

// Diagnostic only, same posture as checkForMissedMoatSignal_ — never clears
// or edits the gap fields. Sandstorm's gap blended details drawn from
// multiple pillars/sources into one sentence while gap_source cited only
// one URL, so nothing verified the WHOLE stated gap against that single
// source. Unlike verifyGapAgainstEvidence_'s narrow financial-keyword check,
// there's no fixed vocabulary for "this detail came from somewhere else" —
// reuses the existing word-overlap helper (built for alignment-capability
// matching) as a general fidelity signal instead. Low overlap is a prompt to
// go check manually, not proof of fabrication — a legitimate, well-
// paraphrased finding can also score low, which is exactly why this only
// logs a warning rather than auto-clearing like the financial-keyword check
// does.
function checkGapSourceFidelity_(draft, snippets, companyName, runId) {
  if (!draft.gap || !draft.gap.source) return draft;

  const gapText = normalizeForMatch_(
    String(draft.gap.currently_monetized || '') + ' ' + String(draft.gap.summary || '')
  );
  if (!gapText) return draft;

  const sourceEvidence = snippets.filter(function (s) { return s.source_url === draft.gap.source; })[0];
  const sourceText = normalizeForMatch_(sourceEvidence ? (sourceEvidence.full_text || sourceEvidence.snippet || '') : '');

  const overlap = wordOverlapRatio_(gapText, sourceText);
  if (overlap < 0.3) {
    logWarning_(runId, companyName, 'gemini_structure', 'gap for "' + companyName + '" has low word overlap (' + Math.round(overlap * 100) + '%) with its cited source (' + draft.gap.source + ') — currently_monetized/summary may synthesize details not actually in that source. Check manually before trusting this as fully backed by the cited URL.');
  }

  return draft;
}

// Found on a 100-company run: a bare "log and skip to next step" on a
// Gemini 503 ("high demand, try again later" — Google's own wording says
// this is transient) permanently left several companies with empty fields,
// even though the very next attempt (sometimes hours later, once Apps
// Script's own trigger got around to firing again) would likely have
// succeeded. These codes are worth a few short retries before giving up.
const GEMINI_RETRYABLE_CODES = [429, 500, 502, 503, 504];
const GEMINI_MAX_RETRIES = 3;
const GEMINI_RETRY_DELAY_MS = 15000;

// stage/runId/companyName are all optional (Test.js's testApiKeys() calls
// this with just a prompt) — purely for EventLog/TokenLog attribution, never
// required for the call itself to work.
function callGemini_(prompt, stage, runId, companyName) {
  const apiKey = getScriptProp_('GEMINI_API_KEY');
  const model = PropertiesService.getScriptProperties().getProperty('GEMINI_MODEL') || 'gemini-2.5-flash';
  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + model + ':generateContent?key=' + apiKey;

  if (CircuitBreaker.isHalted('GEMINI')) {
    throw new Error('Gemini circuit breaker is halted — too many consecutive failures, cooling down.');
  }

  for (let attempt = 1; attempt <= GEMINI_MAX_RETRIES; attempt++) {
    rateLimit('GEMINI');

    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig: {
          responseMimeType: 'application/json',
          // temperature: 0 — this is classification/extraction, not creative
          // writing. Identical evidence was visibly reclassifying differently
          // between runs (patents moved pillars 3 times; a moat-worthy fact
          // moved to alignment and back) with no code change in between.
          temperature: 0,
          // No benefit from open-ended "thinking" on a temperature-0
          // structured-extraction task — disabling saves latency/cost.
          thinkingConfig: { thinkingBudget: 0 }
        }
      }),
      muteHttpExceptions: true
    });

    const code = response.getResponseCode();
    const data = JSON.parse(response.getContentText());

    if (code === 200) {
      CircuitBreaker.recordSuccess('GEMINI');
      rateLimitSuccess('GEMINI');
      TokenLog.logGemini(runId, companyName, stage, model, data.usageMetadata);
      const text = data.candidates && data.candidates[0] && data.candidates[0].content.parts[0].text;
      if (!text) throw new Error('Gemini returned no content: ' + response.getContentText());
      return JSON.parse(text);
    }

    const body = response.getContentText();
    // Token usage may still be present even on a non-200 (e.g. a content
    // filter block) — a rejected call still spent tokens, log it if we can.
    if (data.usageMetadata) TokenLog.logGemini(runId, companyName, stage, model, data.usageMetadata);

    if (isExhaustedResponse_(code, body)) {
      CircuitBreaker.recordFailure('GEMINI');
      rateLimitBackoff('GEMINI');
      haltBatchForExhaustion_('GEMINI', 'HTTP ' + code + ': ' + body, runId);
      throw new Error('Gemini API error (' + code + '): ' + body);
    }

    const canRetry = GEMINI_RETRYABLE_CODES.indexOf(code) !== -1 && attempt < GEMINI_MAX_RETRIES;
    if (!canRetry) {
      CircuitBreaker.recordFailure('GEMINI');
      rateLimitBackoff('GEMINI');
      throw new Error('Gemini API error (' + code + '): ' + body);
    }

    // Retryable and attempts remain — deliberately does NOT count this
    // attempt toward the circuit breaker's consecutive-failure counter; a
    // transient blip that resolves within a couple retries shouldn't look
    // like a sustained outage.
    const message = 'Gemini HTTP ' + code + ' (attempt ' + attempt + '/' + GEMINI_MAX_RETRIES + ') — retrying in ' + (GEMINI_RETRY_DELAY_MS / 1000) + 's: ' + body.slice(0, 200);
    Logger.log('WARNING: ' + message);
    if (typeof EventLog !== 'undefined') EventLog.warn(runId, companyName, '', stage, 'retry', message);
    Utilities.sleep(GEMINI_RETRY_DELAY_MS);
  }
}
