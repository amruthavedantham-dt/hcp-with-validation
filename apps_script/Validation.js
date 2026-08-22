// 5th pipeline step — an independent Gemini re-check of the 3 pillar claims
// already drafted by runGeminiStructure (Gemini.js), reproducing the manual
// verification procedure (identity anchoring, per-claim re-sourcing,
// corroboration-tier discipline, contradiction flagging) as code instead of
// a human doing it by hand every time. At 50-70 companies/day, a human can
// realistically spot-check only a handful of cells, not every row — so this
// file's job is to make "needs-review" a SIGNAL that's actually worth
// spending that scarce attention on, not noise.
//
// needs-review fires ONLY for genuine unresolved doubt: an identity-
// coherence problem, or a real contradiction between sources. A claim
// backed by a plausible single source (even self-claimed) that isn't
// contradicted and isn't identity-suspect is allowed to be Verified — the
// old "Verified requires 2+ independent sources" gate produced a review
// queue full of ordinary small companies with nothing wrong, which drowned
// out the rows that actually needed a look.
//
// Two further asks (given the scale — a human cannot verify most companies,
// so empty/uncertain fields are worse than a clearly-labeled best guess):
//  - Corroboration search: before judging an existing claim, if nothing
//    already staged in RAW_EVIDENCE for that pillar comes from an
//    independent (different-domain) source, fire ONE fresh Serper search
//    looking specifically for outside confirmation — not just re-fetching
//    the same cited page.
//  - Empty-pillar fill: a pillar with NO claim at all, but at least some
//    real evidence (existing or freshly searched), gets a best-guess claim
//    filled in and labeled 'Inferred — Low Confidence' rather than staying
//    blank. A pillar with truly zero evidence stays empty — there is
//    nothing to infer from, and inventing one would be fabrication, not a
//    guess. This is NOT the same failure mode rule 1 of buildGeminiPrompt_
//    guards against (never claim a fact absent from ANY snippet) — an
//    inferred claim must still be grounded in real evidence, just weaker
//    evidence than Verified normally requires, and it is always visibly
//    labeled as such rather than presented as equally reliable.
//
// Hard boundaries, enforced throughout this file:
// - NEVER writes differentiation_review/moat_review/product_improvement_review
//   or their _notes siblings — those are human-only forever (see SheetIO.js).
//   This file only writes the *_auto_review / *_auto_notes columns (and, for
//   an empty-pillar fill only, the claim/source/credibility/needs_verify
//   columns that would otherwise sit blank).
// - NEVER writes cin/ownership_type — identity stays a manual registry check
//   per the project's existing discipline (see Identity.js).
// - Fails closed: an unfetchable source, a malformed verdict, or missing
//   evidence all resolve to "Needs Verify", never silently to "Verified".
// - 'Inferred — Low Confidence' never counts toward auto_triage_status —
//   it is an accepted, visible judgment call, not an unresolved doubt.

const VALIDATION_PILLARS = ['differentiation', 'moat', 'product_improvement'];
const INFERRED_LOW_CONFIDENCE = 'Inferred — Low Confidence';

// Corroboration-search query templates — deliberately generic/topical
// (not just re-searching the company name, which is what Serper Grounding
// already did) so this has a real chance of surfacing something the
// original grounding pass didn't. {company} substituted at call time.
const VALIDATION_CORROBORATION_QUERY = '"{company}" review OR news OR profile OR directory -site:{domain}';

// forceRerun: when true, re-checks every populated pillar regardless of
// whether its claim text is unchanged since the last validation pass.
// Needed because the staleness check below only detects "the CLAIM changed"
// — it has no way to know "the validation LOGIC changed" (a prompt/code fix
// to Validation.js itself). Without this, re-running validation after fixing
// a bug in this file silently returns the same stale, already-wrong verdict
// forever, since the claim text it's keyed on never moved. The single-row
// menu item (Menu.js) defaults this to true — a deliberate manual click is
// an explicit "check this now" request, not a batch-cost-saving pass. Batch
// runs (Batch.js) pass false, keeping the normal skip-unchanged behavior.
function runGeminiValidation(companyName, runId, forceRerun) {
  runId = runId || newManualRunId_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName(RAW_EVIDENCE_SHEET_NAME);
  if (!rawSheet) throw new Error('RAW_EVIDENCE sheet not found. Run setupHVTSheet() first.');

  const rowRange = findOrCreateCompanyRow_(companyName);
  const values = rowRange.getValues()[0];
  const get = function (colName) { return String(values[HVT_HEADERS.indexOf(colName)] || '').trim(); };

  const snippets = getRawEvidenceForCompany_(rawSheet, companyName);
  const website = get('website');

  const populatedPillars = VALIDATION_PILLARS.filter(function (p) { return get(p) !== ''; });
  const emptyPillars = VALIDATION_PILLARS.filter(function (p) { return get(p) === ''; });

  // Re-check: skip a pillar whose claim text is unchanged since its last
  // validation pass — a fresh Gemini call on identical input is pure waste,
  // UNLESS forceRerun is set (see comment above function).
  const pillarsNeedingRecheck = populatedPillars.filter(function (p) {
    if (forceRerun) return true;
    const autoReview = get(p + '_auto_review');
    if (!autoReview || autoReview === 'Not Run') return true;
    const autoNotes = get(p + '_auto_notes');
    const claimText = get(p);
    return !autoNotes || autoNotes.indexOf(claimStamp_(claimText)) === -1;
  });

  // Fill attempt: an empty pillar only needs to be attempted once — if a
  // prior pass already tried and found truly nothing, auto_review will be
  // 'Not Run' still (attemptEmptyPillarFill_ never writes anything when
  // there's no evidence), so re-attempting on every run is cheap enough not
  // to bother gating further; a genuinely thin/empty company won't suddenly
  // start returning results from the same cached RAW_EVIDENCE. Only new
  // evidence (a fresh Serper Grounding re-run) would change the outcome,
  // which already resets auto_review back to 'Not Run' via schema/RAW_EVIDENCE
  // clearing — so no separate staleness check is needed here.
  const pillarsToFill = emptyPillars;

  if (pillarsNeedingRecheck.length === 0 && pillarsToFill.length === 0) {
    EventLog.info(runId, companyName, '', 'gemini_validation', 'skipped', 'Nothing to validate or fill — all populated claims unchanged, and no empty pillars to attempt.');
    return null;
  }

  EventLog.info(runId, companyName, '', 'gemini_validation', 'start',
    'Re-checking: ' + (pillarsNeedingRecheck.join(', ') || '(none)') + '. Attempting fill: ' + (pillarsToFill.join(', ') || '(none)'));

  const identityWarnings = checkIdentityCoherence_(companyName, get, snippets, pillarsNeedingRecheck, runId);
  const refetched = refetchPillarSources_(companyName, get, pillarsNeedingRecheck, runId);
  const corroboration = runCorroborationSearches_(companyName, website, get, pillarsNeedingRecheck, snippets, runId);

  const fields = {};

  if (pillarsNeedingRecheck.length > 0) {
    const prompt = buildValidationPrompt_(companyName, get, pillarsNeedingRecheck, refetched, snippets, corroboration, identityWarnings);
    const rawVerdicts = callGemini_(prompt, 'gemini_validation', runId, companyName);
    const verdicts = validateAndSanitizeValidation_(rawVerdicts, pillarsNeedingRecheck, identityWarnings, refetched, companyName, runId);

    pillarsNeedingRecheck.forEach(function (p) {
      const v = verdicts[p] || {};
      fields[p + '_auto_review'] = v.verdict;
      fields[p + '_auto_notes'] = buildAutoNotes_(get(p), v, refetched[p], corroboration[p]);
    });
  }

  if (pillarsToFill.length > 0) {
    pillarsToFill.forEach(function (p) {
      const fill = attemptEmptyPillarFill_(companyName, p, snippets, get, runId);
      if (fill) {
        fields[p] = fill.claim;
        fields[p + '_source'] = fill.source_url;
        fields[p + '_credibility'] = 'self-claimed';
        fields[p + '_needs_verify'] = 'yes';
        fields[p + '_auto_review'] = INFERRED_LOW_CONFIDENCE;
        fields[p + '_auto_notes'] = claimStamp_(fill.claim) + '\nInferred fill: ' + fill.rationale + '\nSource: ' + fill.source_url;
        if (p === 'moat' && fill.moat_type) fields.moat_type = fill.moat_type;
        if (p === 'moat' && fill.moat_durability) fields.moat_durability = fill.moat_durability;
      }
    });
  }

  if (Object.keys(fields).length > 0) setRowFields_(rowRange, fields);

  // auto_triage_status reflects ALL populated pillars (including ones
  // skipped this run because unchanged) — read back whatever is now on the
  // row, not just this run's subset. Inferred-low-confidence fills do NOT
  // count as needs-review (see file header) — only genuine doubt does.
  const finalValues = rowRange.getValues()[0];
  const allPillarsNow = VALIDATION_PILLARS.filter(function (p) { return String(finalValues[HVT_HEADERS.indexOf(p)] || '').trim() !== ''; });
  const anyNeedsReview = allPillarsNow.some(function (p) {
    const v = String(finalValues[HVT_HEADERS.indexOf(p + '_auto_review')] || '');
    return v === 'Needs Verify' || v === 'Rejected' || v === 'Contradiction Found';
  });
  setRowFields_(rowRange, { auto_triage_status: anyNeedsReview ? TRIAGE_STATUS.NEEDS_REVIEW : TRIAGE_STATUS.AUTO_CONFIRMED });

  EventLog.info(runId, companyName, '', 'gemini_validation', 'done', 'auto_triage: ' + (anyNeedsReview ? TRIAGE_STATUS.NEEDS_REVIEW : TRIAGE_STATUS.AUTO_CONFIRMED));

  return fields;
}

// Short, stable fingerprint of claim text embedded in auto_notes so a later
// run can tell "this exact claim was already validated" without a separate
// hidden column — cheap and avoids adding schema just to track staleness.
function claimStamp_(claimText) {
  return '[[claim:' + Utilities.base64Encode(Utilities.newBlob(claimText).getBytes()).slice(0, 16) + ']]';
}

function buildAutoNotes_(claimText, verdict, refetchInfo, corroborationInfo) {
  const lines = [
    claimStamp_(claimText),
    'Verdict: ' + (verdict.verdict || '') + ' (corroboration: ' + (verdict.corroboration_tier || 'unknown') + ')',
    'Rationale: ' + (verdict.rationale || '(none given)')
  ];
  if (verdict.contradiction_note) lines.push('Contradiction: ' + verdict.contradiction_note);
  if (refetchInfo && refetchInfo.url) {
    lines.push('Re-fetched source: ' + refetchInfo.url + (refetchInfo.fetchFailed ? ' (re-fetch FAILED — verdict based on original cached snippet only)' : ''));
  }
  if (corroborationInfo && corroborationInfo.searched) {
    lines.push('Corroboration search run: "' + corroborationInfo.query + '"' + (corroborationInfo.found ? ' — found ' + corroborationInfo.resultUrl : ' — no independent result found'));
  }
  if (verdict.identityWarning) lines.push('Identity warning: ' + verdict.identityWarning);
  return lines.join('\n');
}

// Re-fetches each rechecked pillar's cited *_source URL fresh, going through
// the same cache-first / free-direct-then-paid-scrape path Serper.js and
// Identity.js already use, so this costs nothing extra on a re-run unless
// the cache was explicitly cleared for the row. Falls back to whatever
// full_text/snippet is already staged in RAW_EVIDENCE for that exact URL if
// the live re-fetch fails — never blocks validation entirely just because
// one source is unreachable (ZaubaCorp/LinkedIn/Facebook-style 403s are
// common and expected, not exceptional).
function refetchPillarSources_(companyName, get, pillars, runId) {
  const result = {};
  pillars.forEach(function (p) {
    const url = get(p + '_source');
    if (!url) { result[p] = { url: '', text: '', fetchFailed: false }; return; }

    const text = scrapeWithGuards_(url, companyName, 'validation_refetch', runId);
    result[p] = { url: url, text: text, fetchFailed: !text };
    if (!text) {
      logWarning_(runId, companyName, 'gemini_validation', p + ' source ' + url + ' could not be re-fetched — falling back to originally staged evidence text for this pillar.');
    }
  });
  return result;
}

// Fires ONE fresh Serper search per pillar being rechecked, but ONLY when
// nothing already staged in RAW_EVIDENCE for that pillar comes from a
// domain independent of the cited source (i.e. no existing corroboration
// candidate) — keeps this from doubling Serper cost on every validation run
// when the structuring pass already turned up independent evidence. Result
// is staged into RAW_EVIDENCE (tagged so it doesn't collide with the normal
// pillar tags) so it also becomes available to the NEXT validation pass and
// to a human reviewing RAW_EVIDENCE directly, not thrown away after this call.
function runCorroborationSearches_(companyName, website, get, pillars, snippets, runId) {
  const result = {};
  const domain = website ? extractDomain_(website) : '';

  pillars.forEach(function (p) {
    const sourceUrl = get(p + '_source');
    const sourceDomain = sourceUrl ? extractDomain_(sourceUrl) : '';

    const hasIndependentEvidence = snippets.some(function (s) {
      return s.pillar === p && s.source_url && extractDomain_(s.source_url) !== sourceDomain;
    });

    if (hasIndependentEvidence || !sourceUrl) {
      result[p] = { searched: false };
      return;
    }

    const query = VALIDATION_CORROBORATION_QUERY
      .replace('{company}', companyName)
      .replace('{domain}', domain || sourceDomain);

    const rawSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RAW_EVIDENCE_SHEET_NAME);
    const results = fetchSerperResults_(companyName, p + '_corroboration', query, runId, function () { return callSerper_(query); });

    if (results.length > 0 && results[0].link && extractDomain_(results[0].link) !== sourceDomain) {
      const now = new Date();
      rawSheet.getRange(rawSheet.getLastRow() + 1, 1, 1, RAW_EVIDENCE_HEADERS.length).setValues([[
        companyName, p, query, results[0].link || '', results[0].title || '', results[0].snippet || '', results[0].full_text || '', now
      ]]);
      snippets.push({ pillar: p, query: query, source_url: results[0].link, title: results[0].title, snippet: results[0].snippet, full_text: results[0].full_text });
      result[p] = { searched: true, query: query, found: true, resultUrl: results[0].link };
    } else {
      result[p] = { searched: true, query: query, found: false };
    }
  });

  return result;
}

// Pure set-membership check against this company's own staged RAW_EVIDENCE —
// no Gemini call, no live registry lookup. Targets the exact contamination
// pattern found in production (Vikalp/Nostrain — a pillar claim sourced from
// an unrelated same-name entity): a claim whose cited source domain never
// appeared anywhere in search results scoped to this company's own name is a
// same-name-collision red flag, independent of whether the claim otherwise
// reads as plausible.
function checkIdentityCoherence_(companyName, get, snippets, pillars, runId) {
  const knownDomains = {};
  snippets.forEach(function (s) { if (s.source_url) knownDomains[extractDomain_(s.source_url)] = true; });
  const website = get('website');
  if (website) knownDomains[extractDomain_(website)] = true;

  const warnings = {};
  pillars.forEach(function (p) {
    const sourceUrl = get(p + '_source');
    if (!sourceUrl) return;
    const domain = extractDomain_(sourceUrl);
    if (!knownDomains[domain]) {
      const msg = p + ' source domain "' + domain + '" was not found among any search result scoped to "' + companyName +
        '" — possible same-name-different-entity contamination (cf. Vikalp/Nostrain). Forcing extra scrutiny.';
      logWarning_(runId, companyName, 'gemini_validation', msg);
      warnings[p] = msg;
    }
  });
  return warnings;
}

function buildValidationPrompt_(companyName, get, pillars, refetched, snippets, corroboration, identityWarnings) {
  const cin = get('cin');
  const sector = get('sector');
  const website = get('website');

  const pillarBlocks = pillars.map(function (p) {
    const claim = get(p);
    const sourceUrl = get(p + '_source');
    const refetch = refetched[p];
    const liveText = refetch.fetchFailed
      ? '(RE-FETCH FAILED — this source could not be reached again. Use the ORIGINAL staged evidence below instead, and note in your rationale that the source was not independently re-verifiable this pass.)'
      : refetch.text;

    const originalStaged = snippets.filter(function (s) { return s.source_url === sourceUrl; })[0];
    const originalText = originalStaged ? (originalStaged.full_text || originalStaged.snippet || '') : '(not found in RAW_EVIDENCE)';

    const otherEvidence = snippets
      .filter(function (s) { return s.pillar === p && s.source_url !== sourceUrl; })
      .map(function (s, i) { return '   ' + (i + 1) + '. URL: ' + s.source_url + ' | Title: ' + s.title + ' | Content: ' + (s.full_text || s.snippet); })
      .join('\n');

    const corrobLine = corroboration[p] && corroboration[p].searched
      ? (corroboration[p].found
          ? 'A fresh independent search for outside confirmation found: ' + corroboration[p].resultUrl + ' (included above in "other evidence" if relevant to this pillar).'
          : 'A fresh independent search for outside confirmation of this claim found nothing usable.')
      : '';

    return [
      'PILLAR: ' + p,
      'Claim as currently written: "' + claim + '"',
      'Cited source URL: ' + (sourceUrl || '(none)'),
      'Cited source — LIVE re-fetched text just now: ' + liveText,
      'Cited source — originally staged text (for comparison only, may be stale): ' + originalText,
      corrobLine,
      identityWarnings[p] ? 'IDENTITY COHERENCE WARNING: ' + identityWarnings[p] : '',
      'Other ' + p + '-tagged evidence NOT currently cited (candidate corroboration or contradiction):\n' + (otherEvidence || '   (none)')
    ].filter(function (l) { return l; }).join('\n');
  }).join('\n\n');

  return [
    'You are independently RE-VERIFYING claims already drafted for "' + companyName + '" (CIN: ' + (cin || 'unknown') +
      ', sector: ' + (sector || 'unknown') + ', website: ' + (website || 'unknown') + '). ' +
      'You did not write these claims — your job is adversarial verification, not confirmation. Actively look for the reason a claim might be wrong (wrong entity, self-serving source, stale data, unsupported inference) rather than a reason to accept it.',
    '',
    'STRICT RULES:',
    '1. Judge each claim ONLY against the LIVE re-fetched source text (or, if re-fetch failed, the originally staged text) and the other evidence provided. Do not use outside knowledge about the company.',
    '2. corroboration_tier must be exactly one of: self-claimed (only the source itself, or another page on the same domain, supports the claim), third-party-unaudited (an independent-looking source — registry aggregator, directory, single news article — supports it, but it is not a government/registry primary source and there is no second independent source), independently-corroborated (two or more genuinely unrelated sources — e.g. the cited source AND a different-domain piece of "other evidence" above — describe the same fact consistently). This is recorded for reference on every verdict — it does NOT by itself decide the verdict (see rule 3). Do not upgrade a tier just because a claim sounds specific or confident — specificity is not corroboration.',
    '3. verdict must be exactly one of: Verified, Needs Verify, Rejected, Contradiction Found. Pick based on genuine doubt, not source count — a self-claimed source is perfectly fine evidence for "Verified" as long as it actually says what the claim says, the company\'s identity is not in question, and nothing contradicts it. Specifically: use "Verified" when the live source text genuinely supports the claim AND there is no identity-coherence warning for this pillar AND no contradiction with other evidence. Use "Needs Verify" ONLY for genuine unresolved doubt: an identity-coherence warning is present (see rule 4), OR the source could not be re-fetched and you cannot independently confirm it this pass (see rule 5), OR the live text is ambiguous/thin such that you genuinely cannot tell if it supports the claim. Use "Rejected" when the live re-fetched (or staged) source text clearly does NOT support the claim as written — e.g. the page no longer says what the claim asserts, or never did; this is a confident negative finding, not "insufficient evidence" (that is Needs Verify instead). Use "Contradiction Found" when the cited source and another piece of evidence for the SAME pillar state genuinely conflicting facts (different founding years, different capability scope, different certification status) — do not silently pick whichever one sounds more credible.',
    '4. If an IDENTITY COHERENCE WARNING is present for a pillar, treat it as a serious reason for suspicion — a claim whose source was never found under this company\'s own name search is a real risk of describing a different, unrelated company that happens to share a name. Do not set Verified for such a pillar; use Needs Verify or Rejected depending on whether the live text is even about a plausible match for ' + companyName + '.',
    '5. If the source re-fetch failed for a pillar, you MUST NOT set verdict to Rejected on that basis alone (an unreachable page is not evidence the claim is false). If the ORIGINAL staged evidence text (also shown above) clearly and specifically supports the claim, you may still use "Verified" based on that original text — say so explicitly in rationale. Only fall back to "Needs Verify" if the original staged text is also too thin/ambiguous to independently confirm the claim.',
    '6. contradiction_note: empty string unless verdict is "Contradiction Found" — then state both conflicting facts and both source URLs plainly. Never resolve the conflict yourself by picking one.',
    '7. rationale: one short sentence explaining the verdict — what the live text does or does not actually support.',
    '',
    pillarBlocks,
    '',
    'Respond with ONLY this JSON shape (no markdown, no commentary), one entry per pillar listed above:',
    '{',
    pillars.map(function (p) {
      return '  "' + p + '": {"verdict": "", "corroboration_tier": "", "contradiction_note": "", "rationale": ""}';
    }).join(',\n'),
    '}'
  ].join('\n');
}

// Deterministic backstop, same posture as validateAndSanitizeDraft_ in
// Gemini.js. Needs-review is meant to fire ONLY for genuine doubt — an
// unresolved identity question or an actual contradiction — not merely
// because a claim is self-claimed/single-sourced (see rule 3 in the prompt:
// self-claimed evidence that plainly supports the claim, with no identity
// risk and no contradiction, is allowed to be Verified). What IS still
// enforced here in code, because a soft prompt rule alone has repeatedly
// failed to hold elsewhere in this project even when reworded:
// - an identity-coherence warning must never coexist with Verified
// - a re-fetch failure must never justify Rejected on its own
function validateAndSanitizeValidation_(rawVerdicts, pillars, identityWarnings, refetched, companyName, runId) {
  const out = {};
  pillars.forEach(function (p) {
    const v = (rawVerdicts && rawVerdicts[p]) || {};

    if (AUTO_REVIEW_STATUS_OPTIONS.indexOf(v.verdict) === -1 || v.verdict === 'Not Run' || v.verdict === INFERRED_LOW_CONFIDENCE) {
      logWarning_(runId, companyName, 'gemini_validation', p + '.verdict "' + v.verdict + '" invalid — defaulting to Needs Verify (fail closed)');
      v.verdict = 'Needs Verify';
    }

    const validTiers = ['self-claimed', 'third-party-unaudited', 'independently-corroborated'];
    if (validTiers.indexOf(v.corroboration_tier) === -1) {
      logWarning_(runId, companyName, 'gemini_validation', p + '.corroboration_tier "' + v.corroboration_tier + '" invalid — defaulting to self-claimed');
      v.corroboration_tier = 'self-claimed';
    }

    // Identity-coherence warning overrides the model's own verdict if it
    // still came back Verified despite the warning — fail-closed, same
    // posture as rule 4 in the prompt, enforced rather than trusted.
    if (identityWarnings[p] && v.verdict === 'Verified') {
      logWarning_(runId, companyName, 'gemini_validation', p + ' had an identity coherence warning but verdict was still "Verified" — forcing to Needs Verify.');
      v.verdict = 'Needs Verify';
    }
    if (identityWarnings[p]) v.identityWarning = identityWarnings[p];

    // A re-fetch failure means the source is merely unconfirmable this pass,
    // never disproven — "Rejected" claims a confident negative finding that
    // an unreachable page cannot support.
    if (refetched[p] && refetched[p].fetchFailed && v.verdict === 'Rejected') {
      logWarning_(runId, companyName, 'gemini_validation', p + ' verdict was "Rejected" but its source re-fetch failed — an unreachable page is not evidence of a false claim. Downgrading to Needs Verify.');
      v.verdict = 'Needs Verify';
    }

    if (v.verdict !== 'Contradiction Found') v.contradiction_note = '';

    out[p] = v;
  });
  return out;
}

// Attempts to fill a pillar that currently has NO claim at all, using
// whatever real evidence is already staged for it in RAW_EVIDENCE (this
// intentionally does NOT trigger its own corroboration search — that search
// exists to find a SECOND source for an EXISTING claim, a different job;
// an empty pillar either has evidence from the original Serper Grounding
// pass or it doesn't). Returns null (fill nothing, stays empty) when there
// is truly no evidence — an empty pillar with zero evidence has nothing to
// infer from, and inventing a claim in that case would be exactly the
// fabrication rule 1 of buildGeminiPrompt_ forbids, not a low-confidence
// guess. Only ever writes when at least one real snippet exists.
function attemptEmptyPillarFill_(companyName, pillar, snippets, get, runId) {
  const pillarEvidence = snippets.filter(function (s) { return s.pillar === pillar || s.pillar === pillar + '_corroboration'; });
  if (pillarEvidence.length === 0) return null;

  const evidenceLines = pillarEvidence.map(function (s, i) {
    return (i + 1) + '. URL: ' + s.source_url + ' | Title: ' + s.title + ' | Content: ' + (s.full_text || s.snippet);
  }).join('\n');

  const moatGuidance = pillar === 'moat'
    ? ' Also propose "moat_type" (one of: ' + MOAT_TYPES.join(', ') + ') and "moat_durability" (one of: ' + MOAT_DURABILITY.join(', ') + ') if the evidence supports a specific type/durability — leave both as empty strings if unclear.'
    : '';

  const prompt = [
    'The "' + pillar + '" pillar for "' + companyName + '" currently has NO claim drafted, but the following evidence was found. ' +
      'Propose a best-guess claim ONLY if the evidence below actually supports one, even weakly — do not invent a claim with no basis in this evidence.',
    '',
    'Evidence:',
    evidenceLines,
    '',
    'Definition of "' + pillar + '": ' + (
      pillar === 'differentiation' ? 'what makes the OFFERING itself non-commodity — cornered resource, niche/scale leadership, manufacturing flexibility, a named proprietary technology/process, design library, integration, installed base, brand.' :
      pillar === 'moat' ? 'a barrier that protects differentiation over time — switching-cost, qualification-lock-in, IP, brand, integration, scale, or network. Only propose this if the evidence explicitly describes resulting lock-in, not just a bare capability.' :
      'evidence of the R&D ENGINE itself — DSIR/registry R&D recognition, patents, NPD pipeline, history of product evolution. A product catalog listing is NOT this.'
    ),
    '',
    'Write the claim as a short, analytical, fact-first phrase under 15 words, same style as the rest of this pipeline — no marketing language, never restate the company name.' + moatGuidance,
    '',
    'Respond with ONLY this JSON shape: {"claim": "", "source_url": "", "moat_type": "", "moat_durability": "", "rationale": ""}',
    'If the evidence does NOT actually support any claim for this pillar, respond with {"claim": "", "source_url": "", "moat_type": "", "moat_durability": "", "rationale": ""} (empty claim) rather than guessing.'
  ].join('\n');

  const raw = callGemini_(prompt, 'gemini_validation_fill', runId, companyName);
  const claim = String((raw && raw.claim) || '').trim();
  if (!claim) {
    logInfo_(runId, companyName, 'gemini_validation', pillar + ' fill attempted but evidence did not support any claim — left empty.');
    return null;
  }

  const sourceUrl = String((raw && raw.source_url) || '').trim() || pillarEvidence[0].source_url;
  if (raw.moat_type && MOAT_TYPES.indexOf(raw.moat_type) === -1) raw.moat_type = '';
  if (raw.moat_durability && MOAT_DURABILITY.indexOf(raw.moat_durability) === -1) raw.moat_durability = '';

  logInfo_(runId, companyName, 'gemini_validation', pillar + ' filled from empty via low-confidence inference: "' + claim + '"');
  return {
    claim: claim,
    source_url: sourceUrl,
    moat_type: raw.moat_type || '',
    moat_durability: raw.moat_durability || '',
    rationale: String((raw && raw.rationale) || '(no rationale given)').trim()
  };
}
