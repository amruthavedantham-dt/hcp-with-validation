function findOrCreateCompanyRow_(companyName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HVT_SHEET_NAME);
  if (!sheet) throw new Error('HVT sheet not found. Run setupHVTSheet() first.');

  const lastRow = sheet.getLastRow();
  const nameColIndex = HVT_HEADERS.indexOf('company_name') + 1;

  if (lastRow > 1) {
    const names = sheet.getRange(2, nameColIndex, lastRow - 1, 1).getValues();
    for (let i = 0; i < names.length; i++) {
      if (String(names[i][0]).trim().toLowerCase() === companyName.trim().toLowerCase()) {
        const row = i + 2;
        ensureReviewDefaults_(sheet, row);
        return sheet.getRange(row, 1, 1, HVT_HEADERS.length);
      }
    }
  }

  const newRow = lastRow + 1;
  const rowValues = new Array(HVT_HEADERS.length).fill('');
  rowValues[nameColIndex - 1] = companyName;
  sheet.getRange(newRow, 1, 1, HVT_HEADERS.length).setValues([rowValues]);
  ensureReviewDefaults_(sheet, newRow);

  return sheet.getRange(newRow, 1, 1, HVT_HEADERS.length);
}

// Backfills the 'Pending' default + dropdown validation on a row's _review
// columns whenever they're missing — covers both a genuinely new row AND a
// row a human created by typing company_name directly into the sheet (the
// documented normal workflow), which used to skip this entirely because
// findOrCreateCompanyRow_ found it as "already existing" on the very first
// pipeline call and never ran the default-fill branch. Never overwrites a
// column that already has a real value (e.g. a human-set 'Verified'/
// 'Rejected'), and skips re-applying validation that's already there.
// Returns the number of cells actually touched (value set and/or validation
// added) — repairReviewDefaults_ needs a real count, not just "how many rows
// did we look at," since most rows may already be fine.
function ensureReviewDefaults_(sheet, row) {
  const reviewRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(REVIEW_STATUS_OPTIONS, true)
    .setAllowInvalid(false)
    .build();
  let touched = 0;
  REVIEW_COLUMNS.forEach(function (colName) {
    const colIndex = HVT_HEADERS.indexOf(colName) + 1;
    const cell = sheet.getRange(row, colIndex);
    let cellTouched = false;
    if (!cell.getDataValidation()) { cell.setDataValidation(reviewRule); cellTouched = true; }
    if (String(cell.getValue() || '').trim() === '') { cell.setValue('Pending'); cellTouched = true; }
    if (cellTouched) touched++;
  });

  // Same never-overwrite backfill, for the machine-verdict columns —
  // defaults to 'Not Run' rather than 'Pending' so the two column families
  // stay visually distinguishable even before either pipeline stage has run.
  const autoReviewRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(AUTO_REVIEW_STATUS_OPTIONS, true)
    .setAllowInvalid(false)
    .build();
  AUTO_REVIEW_COLUMNS.forEach(function (colName) {
    const colIndex = HVT_HEADERS.indexOf(colName) + 1;
    const cell = sheet.getRange(row, colIndex);
    let cellTouched = false;
    if (!cell.getDataValidation()) { cell.setDataValidation(autoReviewRule); cellTouched = true; }
    if (String(cell.getValue() || '').trim() === '') { cell.setValue('Not Run'); cellTouched = true; }
    if (cellTouched) touched++;
  });

  return touched;
}

// One-time repair for rows that were created before ensureReviewDefaults_
// existed — walks every existing HVT row and backfills it immediately,
// instead of waiting for some future pipeline call to touch that row again.
// Reports the real number of cells changed, not the row count — a prior
// version of this alert always printed lastRow-1 regardless of whether
// anything actually needed fixing, which made "100 rows backfilled" true
// even when zero cells changed.
function repairReviewDefaults_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HVT_SHEET_NAME);
  if (!sheet) throw new Error('HVT sheet not found. Run setupHVTSheet() first.');
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    SpreadsheetApp.getUi().alert('No company rows in HVT yet.');
    return;
  }
  let rowsTouched = 0;
  let cellsTouched = 0;
  for (let row = 2; row <= lastRow; row++) {
    const touched = ensureReviewDefaults_(sheet, row);
    if (touched > 0) { rowsTouched++; cellsTouched += touched; }
  }
  SpreadsheetApp.getUi().alert(
    cellsTouched > 0
      ? 'Backfilled ' + cellsTouched + ' cell(s) (Pending/Not Run default and/or dropdown validation) across ' + rowsTouched + ' of ' + (lastRow - 1) + ' row(s).'
      : 'Checked all ' + (lastRow - 1) + ' row(s) — every review column already had a value and validation. Nothing to backfill.'
  );
}

function setRowFields_(rowRange, fieldValues) {
  const sheet = rowRange.getSheet();
  const row = rowRange.getRow();
  Object.keys(fieldValues).forEach(function (colName) {
    const colIndex = HVT_HEADERS.indexOf(colName);
    if (colIndex === -1) throw new Error('Unknown column: ' + colName);
    sheet.getRange(row, colIndex + 1).setValue(fieldValues[colName]);
  });
}
