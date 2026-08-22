// PIPELINE_SUMMARY (high-level counts + condensed cost) and THIN_ROWS
// (detail list of companies with too many empty substantive columns) are
// deliberately built together from ONE read of HVT via computeHvtStats_() —
// this is what guarantees the "thin rows" count on PIPELINE_SUMMARY always
// matches the row count on THIN_ROWS. There is no menu path that rebuilds
// one without the other, so the two sheets can never drift apart.
const PIPELINE_SUMMARY_SHEET_NAME = 'PIPELINE_SUMMARY';
const THIN_ROWS_SHEET_NAME = 'THIN_ROWS';
const THIN_ROW_EMPTY_THRESHOLD = 5;

// Columns excluded from the "empty column" count: _review defaults to
// 'Pending' (not pipeline output), _notes/notes is human free-text and
// legitimately often empty, unresolved_pillars being empty is a GOOD sign
// (nothing needs review), triage_status is a status flag not evidence, and
// company_summary is deliberately NOT part of default BATCH_STEPS (see
// Batch.js) so most companies never get it even when fully processed.
//
// align2_*/align3_*, gap_*, and proof_* are ALSO excluded (added 2026-07-25
// after a real 100-company run showed 53/100 flagged as "thin" purely
// because of these) — all three are opportunistic, evidence-gated fields by
// design, not required output: align2/align3 only populate if a 2nd/3rd
// genuinely-supported alignment entry exists (most companies legitimately
// have 0-1), gap_* only populates when evidence shows real latent-vs-
// monetized tension (most legitimately have none), and proof_* only
// populates when a pillar is unresolved (Gemini.js rule 11) — an
// auto-confirmed company is SUPPOSED to have empty proof fields, that's not
// a gap. align1_* stays in the check: a company with literally zero
// alignment evidence at all is still worth a glance.
const THIN_ROW_EXCLUDED_COLUMNS = [
  'company_name', 'website',
  'differentiation_review', 'differentiation_notes',
  'moat_review', 'moat_notes',
  'product_improvement_review', 'product_improvement_notes',
  'differentiation_auto_review', 'differentiation_auto_notes',
  'moat_auto_review', 'moat_auto_notes',
  'product_improvement_auto_review', 'product_improvement_auto_notes',
  'notes', 'unresolved_pillars', 'triage_status', 'auto_triage_status', 'company_summary',
  'align2_capability', 'align2_need', 'align2_pull', 'align2_source',
  'align3_capability', 'align3_need', 'align3_pull', 'align3_source',
  'gap_latent', 'gap_currently_monetized', 'gap_summary', 'gap_activation_path', 'gap_source',
  'proof_type', 'proof_status', 'proof_source', 'proof_pillar'
];
const THIN_ROW_CHECK_COLUMNS = HVT_HEADERS.filter(function (h) { return THIN_ROW_EXCLUDED_COLUMNS.indexOf(h) === -1; });

function computeHvtStats_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hvtSheet = ss.getSheetByName(HVT_SHEET_NAME);
  if (!hvtSheet) throw new Error('HVT sheet not found. Run setupHVTSheet() first.');

  const stats = { totalCompanies: 0, processed: 0, notProcessed: 0, needsReview: 0, autoConfirmed: 0, thinRows: [] };

  const lastRow = hvtSheet.getLastRow();
  if (lastRow <= 1) return stats;

  const values = hvtSheet.getRange(2, 1, lastRow - 1, HVT_HEADERS.length).getValues();
  const nameCol = HVT_HEADERS.indexOf('company_name');
  const sectorCol = HVT_HEADERS.indexOf('sector');
  const triageCol = HVT_HEADERS.indexOf('triage_status');
  const spreadsheetUrl = ss.getUrl();
  const hvtGid = hvtSheet.getSheetId();

  values.forEach(function (row, idx) {
    const companyName = String(row[nameCol] || '').trim();
    if (!companyName) return;
    stats.totalCompanies++;

    // "Processed" = triage_status is set, i.e. Gemini Structure actually
    // completed for this company — not just that some earlier step (Serper
    // Grounding, Identity) touched the row.
    const triageStatus = String(row[triageCol] || '').trim();
    if (triageStatus === TRIAGE_STATUS.NEEDS_REVIEW) {
      stats.processed++;
      stats.needsReview++;
    } else if (triageStatus === TRIAGE_STATUS.AUTO_CONFIRMED) {
      stats.processed++;
      stats.autoConfirmed++;
    } else {
      stats.notProcessed++;
    }

    const missing = THIN_ROW_CHECK_COLUMNS.filter(function (colName) {
      return String(row[HVT_HEADERS.indexOf(colName)] || '').trim() === '';
    });
    if (missing.length > THIN_ROW_EMPTY_THRESHOLD) {
      const sheetRowNum = idx + 2;
      stats.thinRows.push({
        company: companyName,
        sector: row[sectorCol],
        emptyCount: missing.length,
        missingColumns: missing,
        link: '=HYPERLINK("' + spreadsheetUrl + '#gid=' + hvtGid + '&range=A' + sheetRowNum + '","Row ' + sheetRowNum + '")'
      });
    }
  });

  return stats;
}

function buildPipelineSummaryAndThinRows_() {
  const stats = computeHvtStats_();
  const costTotals = computeCostTotals_();
  writePipelineSummarySheet_(stats, costTotals);
  writeThinRowsSheet_(stats);
  SpreadsheetApp.getUi().alert(
    'Pipeline Summary rebuilt.\n\n' +
    stats.totalCompanies + ' companies total, ' + stats.processed + ' processed, ' +
    stats.thinRows.length + ' thin (>' + THIN_ROW_EMPTY_THRESHOLD + ' empty columns — see THIN_ROWS).'
  );
}

function writePipelineSummarySheet_(stats, costTotals) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(PIPELINE_SUMMARY_SHEET_NAME);
  if (sheet) sheet.clear();
  else sheet = ss.insertSheet(PIPELINE_SUMMARY_SHEET_NAME);

  const pct = function (n) { return stats.totalCompanies > 0 ? (Math.round((n / stats.totalCompanies) * 1000) / 10) + '%' : '—'; };

  const rows = [
    ['PIPELINE SUMMARY', '', ''],
    ['Generated', new Date(), ''],
    ['', '', ''],
    ['Total companies', stats.totalCompanies, ''],
    ['Processed (triage_status set)', stats.processed, pct(stats.processed)],
    ['   — auto-confirmed', stats.autoConfirmed, pct(stats.autoConfirmed)],
    ['   — needs review', stats.needsReview, pct(stats.needsReview)],
    ['Not yet processed', stats.notProcessed, pct(stats.notProcessed)],
    ['', '', ''],
    ['Review queue (needs-review)', stats.needsReview, 'see REVIEW_QUEUE sheet'],
    ['Thin rows (>' + THIN_ROW_EMPTY_THRESHOLD + ' empty columns)', stats.thinRows.length, 'see THIN_ROWS sheet'],
    ['', '', ''],
    ['API USAGE & COST (totals — see COST_SUMMARY for breakdown)', '', ''],
    ['Gemini calls', costTotals.geminiCalls, ''],
    ['Serper calls', costTotals.serperCalls, ''],
    ['Gemini cost', '$' + r6_(costTotals.geminiUsd), '₹' + r4_(costTotals.geminiUsd * INR_PER_USD)],
    ['Serper cost', '$' + r6_(costTotals.serperUsd), '₹' + r4_(costTotals.serperUsd * INR_PER_USD)],
    ['TOTAL cost', '$' + r6_(costTotals.totalUsd), '₹' + r4_(costTotals.totalInr)]
  ];

  sheet.getRange(1, 1, rows.length, 3).setValues(rows);
  sheet.getRange(1, 1).setFontWeight('bold').setFontSize(12);
  sheet.getRange(13, 1).setFontWeight('bold');
  sheet.getRange(18, 1, 1, 3).setFontWeight('bold').setBackground('#fff2cc');
  sheet.autoResizeColumns(1, 3);
}

function writeThinRowsSheet_(stats) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(THIN_ROWS_SHEET_NAME);
  if (sheet) sheet.clear();
  else sheet = ss.insertSheet(THIN_ROWS_SHEET_NAME);

  const headers = ['Company', 'Sector', 'Empty Column Count', 'Missing Columns', 'Go to Row'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  sheet.setFrozenRows(1);

  if (stats.thinRows.length === 0) return;

  const rows = stats.thinRows
    .slice()
    .sort(function (a, b) { return b.emptyCount - a.emptyCount; })
    .map(function (t) { return [t.company, t.sector, t.emptyCount, t.missingColumns.join(', '), t.link]; });

  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  sheet.autoResizeColumns(1, 3);
}
