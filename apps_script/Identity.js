function runIdentityResolution(companyName, website, runId) {
  runId = runId || newManualRunId_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName(RAW_EVIDENCE_SHEET_NAME);
  if (!rawSheet) throw new Error('RAW_EVIDENCE sheet not found. Run setupHVTSheet() first.');

  const allSnippets = getRawEvidenceForCompany_(rawSheet, companyName);
  const identitySnippets = allSnippets.filter(function (s) { return s.pillar === 'identity'; });
  if (identitySnippets.length === 0) {
    throw new Error('No staged identity evidence for "' + companyName + '". Run Serper Grounding first.');
  }

  EventLog.info(runId, companyName, website, 'identity_resolution', 'start', 'Identity resolution started');
  const websiteText = scrapeWithGuards_(website, companyName, 'identity_website', runId);
  const existingSegmentRefs = getExistingSegmentRefs_(companyName);

  const prompt = buildIdentityPrompt_(companyName, website, identitySnippets, websiteText, existingSegmentRefs);
  const rawDraft = callGemini_(prompt, 'identity_resolution', runId, companyName);
  const draft = validateIdentityDraft_(rawDraft, companyName, runId);

  const rowRange = findOrCreateCompanyRow_(companyName);
  const fields = {};
  if (draft.cin) fields.cin = draft.cin;
  if (draft.ownership_type) fields.ownership_type = draft.ownership_type;
  if (draft.sector) fields.sector = draft.sector;
  if (draft.segment_ref) fields.segment_ref = draft.segment_ref;
  setRowFields_(rowRange, fields);

  if (draft.conflict_note) {
    appendNote_(rowRange, '[identity ' + new Date().toISOString().slice(0, 10) + '] ' + draft.conflict_note);
  }

  EventLog.info(runId, companyName, website, 'identity_resolution', 'done',
    'cin: ' + (draft.cin || '(none)') + ', ownership_type: ' + (draft.ownership_type || '(none)') + ', sector: ' + (draft.sector || '(none)'));
  return draft;
}

function buildIdentityPrompt_(companyName, website, identitySnippets, websiteText, existingSegmentRefs) {
  const snippetLines = identitySnippets.map(function (s, i) {
    const body = s.full_text ? s.full_text : s.snippet;
    return (i + 1) + '. URL: ' + s.source_url + '\n   Title: ' + s.title + '\n   Content: ' + body;
  }).join('\n');

  return [
    'You are resolving company identity fields for "' + companyName + '" (official website: ' + website + ') from search evidence and the company\'s own website content.',
    '',
    'STRICT RULES:',
    '1. Only use facts present in the evidence below. Never invent a CIN, ownership type, or sector.',
    '2. cin must be a 21-character Indian Corporate Identification Number if one is present in the evidence (format like U00000CH2004PLC027625). If none is present, leave cin as an empty string.',
    '3. ownership_type must be exactly one of: ' + OWNERSHIP_TYPES.join(', ') + '. Both "sole-proprietorship" and "partnership" share the same signal (a GST/Udyam registration, no CIN) — the difference is who owns it. Use "sole-proprietorship" ONLY when evidence explicitly names a single individual owner (e.g. "Legal Status of Firm: Proprietorship", "owner: [person name]"). Use "partnership" when evidence shows a GST/Udyam registration with no CIN but does NOT name a single individual owner — do not default to "sole-proprietorship" just because a CIN is absent; that only tells you it is not a Companies-Act entity, not that it has exactly one owner. Do not force either into "private", which implies a Companies-Act-registered private limited company. If sources disagree, pick the value from the more authoritative-looking source (a registry aggregator like Zauba/Tofler over a generic SEO/content site) and explain the disagreement in conflict_note. If one source shows a CIN for this company while another describes it as a partnership or proprietorship, this is not a "pick the more authoritative source" situation — a CIN and proprietorship/partnership status are mutually exclusive under Indian company law, so this is a genuine unresolved contradiction. In that case leave ownership_type as an empty string (do not guess either way), still report cin if you found one, and explain the contradiction plainly in conflict_note (e.g. "source X shows a CIN under company law; source Y describes it as a proprietorship — status unresolved, needs manual registry check"). If nothing supports any value, leave it as an empty string.',
    '4. sector should be a short, specific industry CATEGORY label, up to 4 words (e.g. "Electrical Components", "Process Equipment", "Bioenergy Equipment") — not a descriptive sentence or a phrase about what the product specifically does. Prefer the most specific accurate category over a generic one. Base it on what the company actually makes, drawn from the website content below.',
    '5. segment_ref should be a short kebab-case slug, 2-4 words, lowercase with hyphens (e.g. "blast-room-systems", "transmission-synchro-rings") identifying the company\'s SPECIFIC product niche — one level narrower than sector, never just a restatement of the sector itself (e.g. "specialty-chemical-intermediates" is NOT specific enough for a company in the "Specialty Chemicals" sector — that is the sector wearing hyphens, not a niche). Base it on what the evidence/website shows the company actually makes. EXISTING segment_ref VALUES ALREADY USED FOR OTHER COMPANIES: ' + (existingSegmentRefs.length > 0 ? existingSegmentRefs.join(', ') : '(none yet)') + '. Reuse an existing value ONLY if this company makes the SAME specific type of product/process as whichever company that value was assigned to — sharing a broad sector is NOT enough to justify reuse. Two companies in the same sector making genuinely different things must get different values: e.g. a company making agricultural sulphur/zinc formulations and a company making pharma structure-directing-agents are both loosely "specialty chemicals," but are not the same niche and must NOT share a segment_ref. When unsure, invent a new specific slug rather than reuse a generic-sounding existing one.',
    '6. conflict_note should be an empty string if sources agree or if there is nothing to resolve.',
    '',
    'IDENTITY SEARCH EVIDENCE:',
    snippetLines,
    '',
    'COMPANY WEBSITE CONTENT (' + website + '):',
    (websiteText || '(could not be fetched)'),
    '',
    'Respond with ONLY this JSON shape (no markdown, no commentary):',
    '{',
    '  "cin": "",',
    '  "ownership_type": "",',
    '  "sector": "",',
    '  "segment_ref": "",',
    '  "conflict_note": ""',
    '}'
  ].join('\n');
}

// Existing segment_ref values across all other companies already in the
// sheet — passed into the prompt so the model has a chance to reuse a slug
// for a matching niche instead of minting a near-duplicate. Not a substitute
// for a real curated taxonomy at production scale (there's no dedup/merge
// step here, just a nudge), but a cheap step in the right direction while
// the sheet is still small enough to pass every existing value in directly.
function getExistingSegmentRefs_(excludeCompanyName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HVT_SHEET_NAME);
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  const nameCol = HVT_HEADERS.indexOf('company_name') + 1;
  const segmentCol = HVT_HEADERS.indexOf('segment_ref') + 1;
  const values = sheet.getRange(2, 1, lastRow - 1, HVT_HEADERS.length).getValues();

  const distinct = {};
  values.forEach(function (row) {
    const name = String(row[nameCol - 1] || '').trim();
    if (name.toLowerCase() === excludeCompanyName.trim().toLowerCase()) return;
    const segment = String(row[segmentCol - 1] || '').trim();
    if (segment) distinct[segment] = true;
  });
  return Object.keys(distinct);
}

function validateIdentityDraft_(draft, companyName, runId) {
  if (draft.cin && !CIN_FORMAT_REGEX.test(draft.cin)) {
    logWarning_(runId, companyName, 'identity_resolution', 'cin "' + draft.cin + '" does not match CIN format — clearing');
    draft.cin = '';
  }
  if (draft.ownership_type && OWNERSHIP_TYPES.indexOf(draft.ownership_type) === -1) {
    logWarning_(runId, companyName, 'identity_resolution', 'ownership_type "' + draft.ownership_type + '" not in enum — clearing');
    draft.ownership_type = '';
  }

  // Force kebab-case regardless of what the model actually output — a
  // formatting instruction alone isn't reliable, and segment_ref only works
  // as a grouping key if its format is consistent even when its wording isn't.
  if (draft.segment_ref) {
    const normalized = String(draft.segment_ref).toLowerCase().trim().replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, '');
    if (normalized !== draft.segment_ref) {
      logInfo_(runId, companyName, 'identity_resolution', 'segment_ref "' + draft.segment_ref + '" normalized to "' + normalized + '"');
    }
    draft.segment_ref = normalized;
  }

  // A CIN is mutually exclusive with BOTH sole-proprietorship and
  // partnership under Indian company law — a CIN is only issued to
  // Companies-Act-registered entities, regardless of how many owners a
  // non-CIN entity has. Deterministic backstop for rule 3's contradiction
  // guidance, in case the model resolves the conflict by picking a side
  // instead of flagging it.
  if (draft.cin && (draft.ownership_type === 'sole-proprietorship' || draft.ownership_type === 'partnership')) {
    logWarning_(runId, companyName, 'identity_resolution', 'cin "' + draft.cin + '" and ownership_type "' + draft.ownership_type + '" are mutually exclusive — clearing ownership_type, leaving cin for manual check');
    const conflictingType = draft.ownership_type;
    draft.ownership_type = '';
    if (!draft.conflict_note) {
      draft.conflict_note = 'cin "' + draft.cin + '" was found but ownership_type was drafted as ' + conflictingType + ' — these are mutually exclusive under Indian company law (a CIN implies Companies-Act registration). Verify manually against MCA/GST records before setting either field.';
    }
  }

  return draft;
}

function appendNote_(rowRange, note) {
  const sheet = rowRange.getSheet();
  const row = rowRange.getRow();
  const notesCol = HVT_HEADERS.indexOf('notes') + 1;
  const cell = sheet.getRange(row, notesCol);
  const existing = cell.getValue();
  cell.setValue(existing ? existing + '\n' + note : note);
}
