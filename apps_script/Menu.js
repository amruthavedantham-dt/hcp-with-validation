function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Hidden Champions')
    .addItem('1. Run Serper Grounding', 'menuRunSerperGrounding')
    .addItem('2. Resolve Identity (CIN/Ownership/Sector)', 'menuRunIdentityResolution')
    .addItem('3. Run Sector-Context Grounding', 'menuRunSectorContextGrounding')
    .addItem('4. Run Gemini Structure', 'menuRunGeminiStructure')
    .addItem('5. Run Gemini Validation', 'menuRunGeminiValidation')
    .addItem('6. Run Company Summary', 'menuRunCompanySummary')
    .addSeparator()
    .addItem('Start Batch (All Companies, All Steps)', 'startBatchProcessing')
    .addItem('Start Batch (New/Incomplete Only)', 'startBatchForIncomplete')
    .addItem('Resume Batch', 'resumeBatchProcessing')
    .addItem('Stop Batch (Pause)', 'stopBatchProcessing')
    .addItem('Retry Failed Companies', 'retryFailedCompanies_')
    .addSeparator()
    .addItem('Build Review Queue', 'buildReviewQueue')
    .addItem('Show Pipeline Summary', 'buildPipelineSummaryAndThinRows_')
    .addItem('Show Cost Summary', 'buildCostSummary')
    .addItem('Reconcile Historical Serper Cost...', 'menuReconcileHistoricalSerperCost')
    .addItem('Estimate Batch Cost...', 'menuEstimateBatchCost')
    .addItem('Show Batch Failures', 'showBatchFailures')
    .addSeparator()
    .addItem('Run Calibration', 'runCalibration')
    .addItem('Snapshot Calibration Evidence', 'snapshotCalibrationEvidence_')
    .addSeparator()
    .addItem('Clear Search Cache (This Row)', 'menuClearCacheForRow')
    .addItem('Clear Search Cache (All)', 'menuClearCacheAll')
    .addItem('Repair Review Column Defaults (Run Once)', 'repairReviewDefaults_')
    .addSeparator()
    .addItem('⚠ Wipe ALL Company Data (Testing Reset)', 'wipeAllCompanyData_')
    .addSeparator()
    .addItem('CRM: Export selected rows (download)', 'menuCrmExportSelectedDownload')
    .addItem('CRM: Export selected rows to Drive', 'menuCrmExportSelectedToDrive')
    .addItem('CRM: Export ALL auto-confirmed to Drive', 'menuCrmExportAllAutoConfirmedToDrive')
    .addToUi();
}

function menuClearCacheForRow() {
  const input = getActiveRowInput_();
  if (!input) return;
  const removed = clearCacheForCompany_(input.name);
  SpreadsheetApp.getUi().alert('Cleared ' + removed + ' cached evidence row(s) for "' + input.name + '". Next Serper Grounding run will fetch fresh.');
}

function menuClearCacheAll() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('Clear the ENTIRE search cache for every company? This cannot be undone.', ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) return;
  clearCacheAll_();
  ui.alert('Search cache cleared for all companies.');
}

// Reads company_name (col A) and website (col B) from whichever row the
// cursor is on in the HVT sheet — no more prompt dialogs. Click any cell in
// the company's row, then run a menu item.
function getActiveRowInput_() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();

  if (sheet.getName() !== HVT_SHEET_NAME) {
    ui.alert('Click a cell in the "' + HVT_SHEET_NAME + '" sheet, on the row for the company you want to run, then try again.');
    return null;
  }

  const row = sheet.getActiveCell().getRow();
  if (row < 2) {
    ui.alert('Click a company row (not the header row) first.');
    return null;
  }

  const nameCol = HVT_HEADERS.indexOf('company_name') + 1;
  const websiteCol = HVT_HEADERS.indexOf('website') + 1;
  const name = String(sheet.getRange(row, nameCol).getValue() || '').trim();
  const website = String(sheet.getRange(row, websiteCol).getValue() || '').trim();

  if (!name) {
    ui.alert('Row ' + row + ' has no company_name (column A). Fill it in first.');
    return null;
  }

  return { name: name, website: website, row: row };
}

function menuRunSerperGrounding() {
  const input = getActiveRowInput_();
  if (!input) return;
  if (!input.website) {
    SpreadsheetApp.getUi().alert('Row ' + input.row + ' has no website (column B) — site-restricted queries will be skipped for "' + input.name + '".');
  }
  const count = runSerperGrounding(input.name, input.website, newManualRunId_());
  SpreadsheetApp.getUi().alert(
    'Serper grounding done for "' + input.name + '". ' + count + ' evidence rows written to RAW_EVIDENCE.'
  );
}

function menuRunGeminiStructure() {
  const input = getActiveRowInput_();
  if (!input) return;
  runGeminiStructure(input.name, newManualRunId_());
  SpreadsheetApp.getUi().alert(
    'Gemini structure draft written for "' + input.name + '" in the HVT sheet.\n\n' +
    'Before trusting anything: open each *_source URL and confirm the page actually says what the claim says.'
  );
}

// Manual click = deliberate "check this now" request, so forceRerun is
// always true here regardless of whether the claim text changed since a
// prior pass — see the forceRerun comment on runGeminiValidation in
// Validation.js for why this matters (a code/prompt fix to validation
// itself wouldn't otherwise trigger a fresh check on unchanged claims).
// Batch runs (Batch.js) pass forceRerun=false to keep the cost-saving
// skip-unchanged behavior for the 50-70/day case.
function menuRunGeminiValidation() {
  const input = getActiveRowInput_();
  if (!input) return;
  const fields = runGeminiValidation(input.name, newManualRunId_(), true);
  if (!fields) {
    SpreadsheetApp.getUi().alert(
      'Nothing to validate for "' + input.name + '" — either no pillar claims/evidence exist yet (run "1. Run Serper Grounding" and "4. Run Gemini Structure" first), ' +
      'or no populated/fillable pillars were found. Check EVENT_LOG for details.'
    );
    return;
  }
  const lines = VALIDATION_PILLARS
    .filter(function (p) { return fields[p + '_auto_review'] !== undefined; })
    .map(function (p) {
      const filled = fields[p] !== undefined ? ' (filled from empty)' : '';
      return p + ': ' + fields[p + '_auto_review'] + filled;
    });
  SpreadsheetApp.getUi().alert(
    'Gemini validation done for "' + input.name + '":\n\n' + (lines.join('\n') || '(no pillars changed)') +
    '\n\nSee the *_auto_review / *_auto_notes columns for details. This never touches the human *_review columns.'
  );
}

function menuRunCompanySummary() {
  const input = getActiveRowInput_();
  if (!input) return;
  const summary = runCompanySummary(input.name, newManualRunId_());
  SpreadsheetApp.getUi().alert('Company summary written for "' + input.name + '":\n\n' + (summary || '(empty — check EVENT_LOG for why)'));
}

function menuRunIdentityResolution() {
  const input = getActiveRowInput_();
  if (!input) return;
  if (!input.website) {
    SpreadsheetApp.getUi().alert('Row ' + input.row + ' has no website (column B) — identity resolution needs it to scrape the company\'s own site. Fill column B first.');
    return;
  }
  const draft = runIdentityResolution(input.name, input.website, newManualRunId_());
  SpreadsheetApp.getUi().alert(
    'Identity draft written for "' + input.name + '":\n\n' +
    'cin: ' + (draft.cin || '(not found)') + '\n' +
    'ownership_type: ' + (draft.ownership_type || '(not found)') + '\n' +
    'sector: ' + (draft.sector || '(not found)') + '\n\n' +
    'These are drafts from search + the company website — confirm cin/ownership_type against MCA master data before treating as fact. Check the notes column for any source conflicts found.'
  );
}

function menuRunSectorContextGrounding() {
  const input = getActiveRowInput_();
  if (!input) return;

  const sheet = SpreadsheetApp.getActiveSheet();
  const sectorCol = HVT_HEADERS.indexOf('sector') + 1;
  const sector = String(sheet.getRange(input.row, sectorCol).getValue() || '').trim();

  if (!sector) {
    SpreadsheetApp.getUi().alert('Row ' + input.row + ' has no sector filled in yet — run "2. Resolve Identity" first (or fill in the sector column manually), then try this again.');
    return;
  }

  const count = runSectorContextGrounding(input.name, sector, newManualRunId_());
  SpreadsheetApp.getUi().alert(
    'Sector-context grounding done for "' + input.name + '" (sector: "' + sector + '"). ' + count + ' evidence rows added to RAW_EVIDENCE.\n\n' +
    'Run "4. Run Gemini Structure" next to see if it picks up better alignment/moat evidence.'
  );
}
