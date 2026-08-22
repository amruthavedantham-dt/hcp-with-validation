// Rebuilt-on-demand review queue, same pattern as COST_SUMMARY — lists rows
// flagged by EITHER honesty signal: triage_status (computed in Gemini.js's
// runGeminiStructure from needs_verify — "did the structuring pass have
// doubts") OR auto_triage_status (computed in Validation.js from the
// *_auto_review verdicts — "did an independent re-check find problems").
// These are deliberately two separate signals, not merged into one, since a
// row can be auto-confirmed at structuring time but still get flagged by
// validation (or vice versa) — one combined queue lets a human work off
// either without needing two separate sheets, while the "Why" column keeps
// the two reasons distinguishable.
const REVIEW_QUEUE_SHEET_NAME = 'REVIEW_QUEUE';

function buildReviewQueue() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hvtSheet = ss.getSheetByName(HVT_SHEET_NAME);
  if (!hvtSheet) throw new Error('HVT sheet not found.');

  const lastRow = hvtSheet.getLastRow();
  let queueSheet = ss.getSheetByName(REVIEW_QUEUE_SHEET_NAME);
  if (queueSheet) queueSheet.clear();
  else queueSheet = ss.insertSheet(REVIEW_QUEUE_SHEET_NAME);

  const headers = ['Company', 'Why', 'Unresolved Pillars', 'Auto-Review Flags', 'Sector', 'Go to Row'];
  queueSheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  queueSheet.setFrozenRows(1);

  if (lastRow <= 1) {
    SpreadsheetApp.getUi().alert('No company rows in HVT yet.');
    return;
  }

  const values = hvtSheet.getRange(2, 1, lastRow - 1, HVT_HEADERS.length).getValues();
  const nameCol = HVT_HEADERS.indexOf('company_name');
  const sectorCol = HVT_HEADERS.indexOf('sector');
  const unresolvedCol = HVT_HEADERS.indexOf('unresolved_pillars');
  const triageCol = HVT_HEADERS.indexOf('triage_status');
  const autoTriageCol = HVT_HEADERS.indexOf('auto_triage_status');

  const spreadsheetUrl = ss.getUrl();
  const hvtGid = hvtSheet.getSheetId();

  const rows = [];
  values.forEach(function (row, idx) {
    const structuringFlagged = row[triageCol] === TRIAGE_STATUS.NEEDS_REVIEW;
    const validationFlagged = row[autoTriageCol] === TRIAGE_STATUS.NEEDS_REVIEW;
    if (!structuringFlagged && !validationFlagged) return;

    const why = [];
    if (structuringFlagged) why.push('structuring');
    if (validationFlagged) why.push('validation');

    const autoFlags = AUTO_REVIEW_COLUMNS
      .map(function (col) { return { pillar: col.replace('_auto_review', ''), value: row[HVT_HEADERS.indexOf(col)] }; })
      .filter(function (f) { return f.value === 'Needs Verify' || f.value === 'Rejected' || f.value === 'Contradiction Found'; })
      .map(function (f) { return f.pillar + ': ' + f.value; })
      .join(', ');

    const companyName = row[nameCol];
    const sheetRowNum = idx + 2;
    const link = '=HYPERLINK("' + spreadsheetUrl + '#gid=' + hvtGid + '&range=A' + sheetRowNum + '","Row ' + sheetRowNum + '")';
    rows.push([companyName, why.join(' + '), row[unresolvedCol], autoFlags, row[sectorCol], link]);
  });

  if (rows.length > 0) {
    queueSheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
  queueSheet.autoResizeColumns(1, headers.length);

  SpreadsheetApp.getUi().alert('Review Queue rebuilt: ' + rows.length + ' row(s) need review out of ' + values.length + ' total.');
}
