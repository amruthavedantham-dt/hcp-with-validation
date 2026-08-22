// CRM export seam — serializes HVT rows into the profiler-import v1 envelope
// and hands it to the operator as a downloadable file or a Google Drive file.
// It NEVER talks to the CRM: no URL, no token, no network call. The operator
// takes the JSON and uploads it themselves through the CRM's "Import from
// profiler" screen, where their own login and the manager-only gate apply.
//
// Read-only against the HVT sheet: nothing here writes back to any pipeline
// column, so it can never clobber batch/serper/gemini output.
//
// Config: NONE. The offering is NOT part of the export at all — the operator
// picks it from a dropdown in the CRM's import screen at upload time, so there
// is nothing to configure and nothing to ask here.

// The CRM import endpoint accepts at most 200 companies per call, so a Drive
// export of a large set is split into 200-company files.
const CRM_EXPORT_MAX_PER_FILE = 200;
// Evidence snippets shipped per company (from the durable SEARCH_CACHE, the
// research corpus behind the row's claims). Caps mirror the CRM's Zod contract.
const CRM_EXPORT_EVIDENCE_PER_COMPANY = 25;
const CRM_EXPORT_EVIDENCE_TEXT_CHARS = 400;
const CRM_EXPORT_FOLDER_NAME = 'CRM Profiler Exports';
const CRM_EXPORT_SEQ_PROP = 'CRM_EXPORT_SEQ';
// The CRM contract's review enum. The sheet's 4th option ('Needs registry
// check') is sheet-workflow-only — it maps to 'Pending' on export, since both
// mean "a human still has to look at this".
const CRM_REVIEW_VALUES = ['Pending', 'Verified', 'Rejected'];
const CRM_EXPORT_PILLARS = ['differentiation', 'moat', 'product_improvement'];

// ---------------------------------------------------------------------------
// Menu entry points (wired in Menu.gs onOpen)
// ---------------------------------------------------------------------------

function menuCrmExportSelectedDownload() {
  const picked = getSelectedHvtCompanies_();
  if (!picked) return;
  if (picked.companies.length === 0) {
    SpreadsheetApp.getUi().alert('None of the selected row(s) have a company_name — nothing to export.');
    return;
  }
  if (picked.skippedRows.length > 0) {
    SpreadsheetApp.getUi().alert('Row(s) ' + picked.skippedRows.join(', ') + ' have no company_name and were skipped. Exporting the other ' + picked.companies.length + '.');
  }
  const envelope = buildExportEnvelope_(picked.companies, newExportRunId_());
  showExportDownloadDialog_(envelope);
}

function menuCrmExportSelectedToDrive() {
  const picked = getSelectedHvtCompanies_();
  if (!picked) return;
  if (picked.companies.length === 0) {
    SpreadsheetApp.getUi().alert('None of the selected row(s) have a company_name — nothing to export.');
    return;
  }
  if (picked.skippedRows.length > 0) {
    SpreadsheetApp.getUi().alert('Row(s) ' + picked.skippedRows.join(', ') + ' have no company_name and were skipped. Exporting the other ' + picked.companies.length + '.');
  }
  exportCompaniesToDrive_(picked.companies, 'selected rows');
}

function menuCrmExportAllAutoConfirmedToDrive() {
  const ui = SpreadsheetApp.getUi();
  const companies = getAllAutoConfirmedCompanies_();
  if (companies.length === 0) {
    ui.alert('No auto-confirmed companies found — every row is either needs-review or not yet triaged (run Gemini Structure first).');
    return;
  }
  const parts = Math.ceil(companies.length / CRM_EXPORT_MAX_PER_FILE);
  const response = ui.alert(
    'Export ALL ' + companies.length + ' auto-confirmed compan' + (companies.length === 1 ? 'y' : 'ies') + ' to Google Drive?\n\n' +
    (parts > 1 ? 'They will be written as ' + parts + ' file(s) of up to ' + CRM_EXPORT_MAX_PER_FILE + ' each (the CRM imports at most ' + CRM_EXPORT_MAX_PER_FILE + ' per file).' : 'One .json file will be written.'),
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;
  exportCompaniesToDrive_(companies, 'all auto-confirmed');
}

// ---------------------------------------------------------------------------
// Row selection / collection
// ---------------------------------------------------------------------------

// Same posture as getActiveRowInput_ (Menu.gs), extended to multi-row: the
// operator selects one or more ranges in the HVT sheet (any columns — only
// the row numbers matter), and every distinct data row in the selection is
// serialized. Rows without a company_name are reported back, not silently
// dropped.
function getSelectedHvtCompanies_() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();

  if (sheet.getName() !== HVT_SHEET_NAME) {
    ui.alert('Select row(s) in the "' + HVT_SHEET_NAME + '" sheet first, then try again.');
    return null;
  }

  const rangeList = sheet.getActiveRangeList();
  if (!rangeList) {
    ui.alert('Select at least one company row first.');
    return null;
  }

  const rowSet = {};
  rangeList.getRanges().forEach(function (range) {
    for (let r = range.getRow(); r <= range.getLastRow(); r++) {
      if (r >= 2) rowSet[r] = true;
    }
  });

  const lastRow = sheet.getLastRow();
  const rowNumbers = Object.keys(rowSet)
    .map(Number)
    .filter(function (r) { return r <= lastRow; })
    .sort(function (a, b) { return a - b; });

  if (rowNumbers.length === 0) {
    ui.alert('Select at least one company row (not the header row).');
    return null;
  }

  const data = sheet.getRange(2, 1, lastRow - 1, HVT_HEADERS.length).getValues();
  const evidenceByCompany = loadEvidenceByCompany_();
  const companies = [];
  const skippedRows = [];

  rowNumbers.forEach(function (r) {
    const company = serializeCompanyRow_(hvtRowToMap_(data[r - 2]));
    if (company) {
      const evidence = evidenceForCompany_(evidenceByCompany, company.identity.name);
      if (evidence.length > 0) company.evidence = evidence;
      companies.push(company);
    } else skippedRows.push(r);
  });

  return { companies: companies, skippedRows: skippedRows };
}

function getAllAutoConfirmedCompanies_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HVT_SHEET_NAME);
  if (!sheet) throw new Error('HVT sheet not found. Run setupHVTSheet() first.');

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  const triageIdx = HVT_HEADERS.indexOf('triage_status');
  const data = sheet.getRange(2, 1, lastRow - 1, HVT_HEADERS.length).getValues();
  const evidenceByCompany = loadEvidenceByCompany_();

  const companies = [];
  data.forEach(function (row) {
    if (String(row[triageIdx] || '').trim() !== TRIAGE_STATUS.AUTO_CONFIRMED) return;
    const company = serializeCompanyRow_(hvtRowToMap_(row));
    if (company) {
      const evidence = evidenceForCompany_(evidenceByCompany, company.identity.name);
      if (evidence.length > 0) company.evidence = evidence;
      companies.push(company);
    }
  });
  return companies;
}

function hvtRowToMap_(rowValues) {
  const rowMap = {};
  HVT_HEADERS.forEach(function (colName, idx) { rowMap[colName] = rowValues[idx]; });
  return rowMap;
}

// ---------------------------------------------------------------------------
// Evidence — the research corpus behind the row, from the durable SEARCH_CACHE
// (RAW_EVIDENCE is cleared per run; the cache persists and carries the same
// snippet + scraped text). Shipping it lets spotlight generation write from
// the evidence itself, not just the three distilled claims.
// ---------------------------------------------------------------------------

// One bulk read for the whole export: lowercase company_name -> cache rows.
function loadEvidenceByCompany_() {
  const byCompany = {};
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SEARCH_CACHE_SHEET_NAME);
  if (!sheet) return byCompany;
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return byCompany;
  const values = sheet.getRange(2, 1, lastRow - 1, RAW_EVIDENCE_HEADERS.length).getValues();
  values.forEach(function (row) {
    const key = String(row[0] || '').trim().toLowerCase();
    if (!key) return;
    if (!byCompany[key]) byCompany[key] = [];
    byCompany[key].push({ pillar: row[1], source: row[3], title: row[4], snippet: row[5], fullText: row[6] });
  });
  return byCompany;
}

// Up to CRM_EXPORT_EVIDENCE_PER_COMPANY deduped entries {text, source?, pillar?}.
// text = "title — snippet" (or a slice of the scraped page when the snippet is
// blank), truncated to the contract cap. source only when it is a real URL
// (the CRM validates strictly and a bad value would reject the whole file).
function evidenceForCompany_(evidenceByCompany, companyName) {
  const rows = evidenceByCompany[String(companyName || '').trim().toLowerCase()] || [];
  const seen = {};
  const out = [];
  for (let i = 0; i < rows.length && out.length < CRM_EXPORT_EVIDENCE_PER_COMPANY; i++) {
    const r = rows[i];
    const title = String(r.title || '').trim();
    const snippet = String(r.snippet || '').trim();
    const fallback = String(r.fullText || '').trim().slice(0, CRM_EXPORT_EVIDENCE_TEXT_CHARS);
    let text = title && snippet ? title + ' — ' + snippet : (snippet || title || fallback);
    text = text.slice(0, CRM_EXPORT_EVIDENCE_TEXT_CHARS).trim();
    if (!text || text.length < 20) continue;
    const dedupeKey = text.toLowerCase();
    if (seen[dedupeKey]) continue;
    seen[dedupeKey] = true;
    const entry = { text: text };
    const source = String(r.source || '').trim();
    if (/^https?:\/\/\S+$/.test(source)) entry.source = source;
    const pillar = String(r.pillar || '').trim();
    if (pillar) entry.pillar = pillar.slice(0, 120);
    out.push(entry);
  }
  return out;
}

// ---------------------------------------------------------------------------
// Serialization — HVT columns -> profiler-import v1 company object
// ---------------------------------------------------------------------------

// Maps one HVT row to the contract's companies[] element. Empty optional
// fields are omitted entirely (the CRM's Zod schema is strict). Returns null
// when the row has no company_name (the only required field).
function serializeCompanyRow_(rowMap) {
  function val(colName) {
    const v = rowMap[colName];
    return String(v === undefined || v === null ? '' : v).trim();
  }

  const name = val('company_name');
  if (!name) return null;

  const identity = { name: name };
  if (val('website')) identity.website = val('website');
  if (val('cin')) identity.cin = val('cin');
  if (val('ownership_type')) identity.ownershipType = val('ownership_type');
  if (val('sector')) identity.sector = val('sector');
  if (val('segment_ref')) identity.segmentRef = val('segment_ref');

  const company = { identity: identity };

  if (val('company_summary')) company.summary = val('company_summary');

  const triageStatus = val('triage_status');
  if (triageStatus === TRIAGE_STATUS.AUTO_CONFIRMED || triageStatus === TRIAGE_STATUS.NEEDS_REVIEW) {
    company.triage = {
      status: triageStatus,
      unresolvedPillars: val('unresolved_pillars')
        .split(',')
        .map(function (p) { return p.trim(); })
        .filter(function (p) { return p; })
    };
  }

  const claims = [];
  CRM_EXPORT_PILLARS.forEach(function (pillar) {
    const claimText = val(pillar);
    if (!claimText) return;

    const claim = { pillar: pillar, claim: claimText };

    // The CRM requires claim.source to be a URL (Zod z.string().url()) and a
    // strict-schema failure rejects the WHOLE file — a hand-typed non-URL
    // cell ("company website", "brochure") must be omitted (source is
    // optional in the contract), never sent.
    const source = val(pillar + '_source');
    if (/^https?:\/\/\S+$/.test(source)) claim.source = source;

    // credibility/needsVerify are REQUIRED on every claim in the contract.
    // A hand-typed row can have them blank — default to the conservative end
    // (self-claimed, needs verification), same posture as
    // validateAndSanitizeDraft_'s invalid-value fallbacks.
    const credibility = val(pillar + '_credibility');
    claim.credibility = CREDIBILITY_TIERS.indexOf(credibility) !== -1 ? credibility : 'self-claimed';

    const needsVerify = val(pillar + '_needs_verify');
    claim.needsVerify = (needsVerify === 'yes' || needsVerify === 'no') ? needsVerify : 'yes';

    const review = val(pillar + '_review');
    if (review === 'Needs registry check') claim.review = 'Pending';
    else if (CRM_REVIEW_VALUES.indexOf(review) !== -1) claim.review = review;

    if (pillar === 'moat') {
      if (val('moat_type')) claim.moatType = val('moat_type');
      if (val('moat_durability')) claim.moatDurability = val('moat_durability');
    }

    claims.push(claim);
  });
  company.claims = claims;

  const alignment = [];
  [1, 2, 3].forEach(function (n) {
    const capability = val('align' + n + '_capability');
    const need = val('align' + n + '_need');
    if (!capability || !need) return;
    const entry = { capability: capability, need: need };
    if (val('align' + n + '_pull')) entry.pull = val('align' + n + '_pull');
    if (val('align' + n + '_source')) entry.source = val('align' + n + '_source');
    alignment.push(entry);
  });
  if (alignment.length > 0) company.alignment = alignment;

  const gap = {};
  if (val('gap_latent')) gap.latent = val('gap_latent');
  if (val('gap_currently_monetized')) gap.currentlyMonetized = val('gap_currently_monetized');
  if (val('gap_summary')) gap.summary = val('gap_summary');
  if (val('gap_activation_path')) gap.activationPath = val('gap_activation_path');
  if (val('gap_source')) gap.source = val('gap_source');
  if (Object.keys(gap).length > 0) company.gap = gap;

  const proof = {};
  if (val('proof_type')) proof.type = val('proof_type');
  if (val('proof_status')) proof.status = val('proof_status');
  if (val('proof_source')) proof.source = val('proof_source');
  if (val('proof_pillar')) proof.pillar = val('proof_pillar');
  if (Object.keys(proof).length > 0) company.proof = proof;

  // contacts: the HVT sheet carries no contact columns — omitted entirely
  // (optional in the contract; the CRM creates none when absent).

  return company;
}

// No offering in the file: the operator picks it from the dropdown in the
// CRM import screen at upload time.
function buildExportEnvelope_(companies, runId) {
  return {
    schemaVersion: 1,
    source: 'hcp-profiler',
    runId: runId,
    exportedAt: new Date().toISOString(),
    companies: companies
  };
}

// Persistent counter + timestamp, so run ids stay unique even across a copy
// of the spreadsheet (which would copy the counter but not the clock).
function newExportRunId_() {
  const props = PropertiesService.getScriptProperties();
  const seq = Number(props.getProperty(CRM_EXPORT_SEQ_PROP) || '0') + 1;
  props.setProperty(CRM_EXPORT_SEQ_PROP, String(seq));
  return 'export-' + seq + '-' + Date.now();
}

// ---------------------------------------------------------------------------
// Drive export
// ---------------------------------------------------------------------------

// Writes one or more .json files (split at CRM_EXPORT_MAX_PER_FILE) into a
// "CRM Profiler Exports" folder in the operator's Drive, then shows a dialog
// with clickable links to the folder and each file.
function exportCompaniesToDrive_(companies, label) {
  const runId = newExportRunId_();
  const folder = getExportFolder_();

  const totalParts = Math.ceil(companies.length / CRM_EXPORT_MAX_PER_FILE);
  const files = [];
  for (let start = 0; start < companies.length; start += CRM_EXPORT_MAX_PER_FILE) {
    const chunk = companies.slice(start, start + CRM_EXPORT_MAX_PER_FILE);
    const partNum = Math.floor(start / CRM_EXPORT_MAX_PER_FILE) + 1;
    const suffix = totalParts > 1 ? '-part' + partNum + '-of-' + totalParts : '';
    const envelope = buildExportEnvelope_(chunk, runId);
    const file = folder.createFile(runId + suffix + '.json', JSON.stringify(envelope, null, 2), 'application/json');
    files.push(file);
  }

  showExportDriveResultDialog_(folder, files, companies.length, label);
}

function getExportFolder_() {
  const it = DriveApp.getFoldersByName(CRM_EXPORT_FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(CRM_EXPORT_FOLDER_NAME);
}

// ---------------------------------------------------------------------------
// Dialogs
// ---------------------------------------------------------------------------

// Download / copy the exact envelope for the selected rows, then upload it in
// the CRM's "Import from profiler" screen. No config needed.
function showExportDownloadDialog_(envelope) {
  const json = JSON.stringify(envelope, null, 2);
  const oversize = envelope.companies.length > CRM_EXPORT_MAX_PER_FILE;
  const html =
    '<div style="font-family:sans-serif;padding:12px;">' +
    (oversize
      ? '<p style="color:#b00020;">' + envelope.companies.length + ' companies — the CRM imports at most ' + CRM_EXPORT_MAX_PER_FILE + ' per file. Export to Drive instead (it splits automatically), or select fewer rows.</p>'
      : '') +
    '<p>' + envelope.companies.length + ' compan' + (envelope.companies.length === 1 ? 'y' : 'ies') +
    ' (runId ' + escapeHtml_(envelope.runId) + '). Download it or copy it, then upload in the CRM\'s "Import from profiler" screen.</p>' +
    '<textarea id="json" readonly style="width:100%;height:320px;font-family:monospace;font-size:11px;white-space:pre;box-sizing:border-box;">' +
    escapeHtml_(json) +
    '</textarea>' +
    '<p><a id="dl" download="' + escapeHtml_(envelope.runId) + '.json">Download .json</a> &nbsp; ' +
    '<button onclick="copyJson()">Copy to clipboard</button></p>' +
    '<script>' +
    'var ta = document.getElementById("json");' +
    'document.getElementById("dl").href = "data:application/json;charset=utf-8," + encodeURIComponent(ta.value);' +
    'function copyJson() { ta.select(); document.execCommand("copy"); }' +
    '</scr' + 'ipt></div>';

  const output = HtmlService.createHtmlOutput(html).setWidth(760).setHeight(560);
  SpreadsheetApp.getUi().showModalDialog(output, 'CRM Export (download)');
}

function showExportDriveResultDialog_(folder, files, companyCount, label) {
  const links = files.map(function (f) {
    return '<li><a href="' + f.getUrl() + '" target="_blank">' + escapeHtml_(f.getName()) + '</a></li>';
  }).join('');
  const html =
    '<div style="font-family:sans-serif;padding:12px;">' +
    '<p>Exported ' + companyCount + ' compan' + (companyCount === 1 ? 'y' : 'ies') + ' (' + escapeHtml_(label) + ') to Drive folder ' +
    '<a href="' + folder.getUrl() + '" target="_blank">' + escapeHtml_(CRM_EXPORT_FOLDER_NAME) + '</a>:</p>' +
    '<ul>' + links + '</ul>' +
    '<p>Download each file from Drive, then upload it in the CRM\'s "Import from profiler" screen.</p>' +
    '</div>';

  const output = HtmlService.createHtmlOutput(html).setWidth(620).setHeight(360);
  SpreadsheetApp.getUi().showModalDialog(output, 'CRM Export to Drive');
}
