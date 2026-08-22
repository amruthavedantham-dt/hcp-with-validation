// Reads GEMINI_LOG/SERPER_LOG (TokenLog.js) and rolls them into 3 stacked
// sections in a rebuilt-on-demand COST_SUMMARY sheet. Menu-triggered, not
// automatic — cheap to rebuild, no reason to run it on every single call.
const COST_SUMMARY_SHEET_NAME = 'COST_SUMMARY';

// Optional, manually-entered Script Property — this account's real Serper
// spend attributable to THIS pipeline (from the Serper dashboard/invoice,
// not a guess). Set it via Project Settings > Script Properties, key
// KNOWN_SERPER_INVOICE_USD, when you have a real figure to compare SERPER_LOG
// against. Left unset, the Reconciled Historical section simply omits the
// comparison — it never assumes a value.
const KNOWN_SERPER_INVOICE_PROP = 'KNOWN_SERPER_INVOICE_USD';

function buildCostSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const geminiRows = readLogSheet_(GEMINI_LOG_SHEET_NAME, GEMINI_LOG_HEADERS.length);
  const serperRows = readLogSheet_(SERPER_LOG_SHEET_NAME, SERPER_LOG_HEADERS.length);
  const knownInvoiceUsd = Number(PropertiesService.getScriptProperties().getProperty(KNOWN_SERPER_INVOICE_PROP)) || null;

  let sheet = ss.getSheetByName(COST_SUMMARY_SHEET_NAME);
  if (sheet) sheet.clear();
  else sheet = ss.insertSheet(COST_SUMMARY_SHEET_NAME);

  let row = 1;
  row = writeRunTotals_(sheet, row, geminiRows, serperRows);
  row += 2;
  row = writeGeminiByStage_(sheet, row, geminiRows);
  row += 2;
  row = writeSerperByType_(sheet, row, serperRows);
  row += 2;
  writeReconciledHistoricalTotals_(sheet, row, serperRows, knownInvoiceUsd);

  sheet.autoResizeColumns(1, 8);
  SpreadsheetApp.getUi().alert('Cost Summary rebuilt.');
}

function readLogSheet_(sheetName, numCols) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  return sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
}

function writeSectionHeader_(sheet, row, title) {
  sheet.getRange(row, 1).setValue(title).setFontWeight('bold').setFontSize(12);
  return row + 1;
}

function writeRunTotals_(sheet, startRow, geminiRows, serperRows) {
  let row = writeSectionHeader_(sheet, startRow, 'PER-RUN TOTALS');
  const headers = ['Run_ID', 'Gemini Calls', 'Gemini Cost USD', 'Gemini Cost INR', 'Serper Calls', 'Serper Cost USD', 'Serper Cost INR', 'Total Cost USD'];
  sheet.getRange(row, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  row++;

  const byRun = {};
  geminiRows.forEach(function (r) {
    const runId = r[1];
    if (!byRun[runId]) byRun[runId] = { geminiCalls: 0, geminiUsd: 0, geminiInr: 0, serperCalls: 0, serperUsd: 0, serperInr: 0 };
    byRun[runId].geminiCalls++;
    byRun[runId].geminiUsd += Number(r[8]) || 0;
    byRun[runId].geminiInr += Number(r[9]) || 0;
  });
  serperRows.forEach(function (r) {
    const runId = r[1];
    if (!byRun[runId]) byRun[runId] = { geminiCalls: 0, geminiUsd: 0, geminiInr: 0, serperCalls: 0, serperUsd: 0, serperInr: 0 };
    byRun[runId].serperCalls++;
    byRun[runId].serperUsd += Number(r[5]) || 0;
    byRun[runId].serperInr += Number(r[6]) || 0;
  });

  let totalUsd = 0;
  Object.keys(byRun).sort().forEach(function (runId) {
    const b = byRun[runId];
    const total = b.geminiUsd + b.serperUsd;
    totalUsd += total;
    sheet.getRange(row, 1, 1, headers.length).setValues([[runId, b.geminiCalls, r6_(b.geminiUsd), r4_(b.geminiInr), b.serperCalls, r6_(b.serperUsd), r4_(b.serperInr), r6_(total)]]);
    row++;
  });

  sheet.getRange(row, 1, 1, headers.length).setValues([['TOTAL', '', '', '', '', '', '', r6_(totalUsd)]]).setFontWeight('bold').setBackground('#fff2cc');
  return row + 1;
}

function writeGeminiByStage_(sheet, startRow, geminiRows) {
  let row = writeSectionHeader_(sheet, startRow, 'GEMINI — PER-STAGE BREAKDOWN');
  const headers = ['Stage', 'Calls', 'Input Tokens', 'Output Tokens', 'Thinking Tokens', 'Total Cost USD', 'Avg Cost/Call USD'];
  sheet.getRange(row, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  row++;

  const byStage = {};
  geminiRows.forEach(function (r) {
    const stage = r[3] || '(unspecified)';
    if (!byStage[stage]) byStage[stage] = { calls: 0, input: 0, output: 0, thinking: 0, usd: 0 };
    byStage[stage].calls++;
    byStage[stage].input += Number(r[5]) || 0;
    byStage[stage].output += Number(r[6]) || 0;
    byStage[stage].thinking += Number(r[7]) || 0;
    byStage[stage].usd += Number(r[8]) || 0;
  });

  const stages = Object.keys(byStage).sort(function (a, b) { return byStage[b].usd - byStage[a].usd; });
  let totalUsd = 0;
  stages.forEach(function (stage) {
    const s = byStage[stage];
    totalUsd += s.usd;
    sheet.getRange(row, 1, 1, headers.length).setValues([[stage, s.calls, s.input, s.output, s.thinking, r6_(s.usd), r6_(s.usd / s.calls)]]);
    row++;
  });
  sheet.getRange(row, 1, 1, headers.length).setValues([['TOTAL', '', '', '', '', r6_(totalUsd), '']]).setFontWeight('bold').setBackground('#fff2cc');
  return row + 1;
}

function writeSerperByType_(sheet, startRow, serperRows) {
  let row = writeSectionHeader_(sheet, startRow, 'SERPER — PER-QUERY-TYPE BREAKDOWN');
  const headers = ['Query_Type', 'Calls', 'Total Results', 'Avg Results/Call', 'Total Cost USD', 'Avg Cost/Call USD'];
  sheet.getRange(row, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  row++;

  const byType = {};
  serperRows.forEach(function (r) {
    const type = r[3] || '(unspecified)';
    if (!byType[type]) byType[type] = { calls: 0, results: 0, usd: 0 };
    byType[type].calls++;
    byType[type].results += Number(r[4]) || 0;
    byType[type].usd += Number(r[5]) || 0;
  });

  const types = Object.keys(byType).sort(function (a, b) { return byType[b].usd - byType[a].usd; });
  let totalUsd = 0;
  let searchUsd = 0, scrapeUsd = 0, searchCalls = 0, scrapeCalls = 0;
  types.forEach(function (type) {
    const t = byType[type];
    totalUsd += t.usd;
    if (/_SCRAPE$/.test(type)) { scrapeUsd += t.usd; scrapeCalls += t.calls; }
    else { searchUsd += t.usd; searchCalls += t.calls; }
    sheet.getRange(row, 1, 1, headers.length).setValues([[type, t.calls, t.results, r4_(t.results / t.calls), r6_(t.usd), r6_(t.usd / t.calls)]]);
    row++;
  });
  sheet.getRange(row, 1, 1, headers.length).setValues([['TOTAL', '', '', '', r6_(totalUsd), '']]).setFontWeight('bold').setBackground('#fff2cc');
  row++;

  // Search vs. scrape split — makes the "paid scrape doubled this query's
  // cost" effect visible at a glance instead of requiring a manual sum of
  // every *_SCRAPE row. scrapeCalls / searchCalls is the actual doubling
  // rate: 0 means fetchPageDirect_ is catching everything for free, 1.0
  // means every query is paying for both /search and scrape.serper.dev.
  row++;
  sheet.getRange(row, 1).setValue('Search calls (1 credit each)');
  sheet.getRange(row, 2).setValue(searchCalls);
  sheet.getRange(row, 5).setValue(r6_(searchUsd));
  row++;
  sheet.getRange(row, 1).setValue('Paid scrape calls (fetchPageDirect_ failed, 1 credit each)');
  sheet.getRange(row, 2).setValue(scrapeCalls);
  sheet.getRange(row, 5).setValue(r6_(scrapeUsd));
  row++;
  sheet.getRange(row, 1).setValue('Scrape rate (scrape calls / search calls)');
  sheet.getRange(row, 2).setValue(searchCalls > 0 ? r4_(scrapeCalls / searchCalls) : 0);
  row++;

  return row + 1;
}

function r6_(n) { return Math.round((n || 0) * 1e6) / 1e6; }
function r4_(n) { return Math.round((n || 0) * 1e4) / 1e4; }

// Grand-totals-only view for PIPELINE_SUMMARY — deliberately not broken
// down by stage/query-type (that detail already lives in COST_SUMMARY).
function computeCostTotals_() {
  const geminiRows = readLogSheet_(GEMINI_LOG_SHEET_NAME, GEMINI_LOG_HEADERS.length);
  const serperRows = readLogSheet_(SERPER_LOG_SHEET_NAME, SERPER_LOG_HEADERS.length);
  let geminiUsd = 0, serperUsd = 0;
  geminiRows.forEach(function (r) { geminiUsd += Number(r[8]) || 0; });
  serperRows.forEach(function (r) { serperUsd += Number(r[5]) || 0; });
  const totalUsd = geminiUsd + serperUsd;
  return {
    geminiCalls: geminiRows.length, geminiUsd: geminiUsd,
    serperCalls: serperRows.length, serperUsd: serperUsd,
    totalUsd: totalUsd, totalInr: totalUsd * INR_PER_USD
  };
}

// Projects a full-scale cost from the actual per-company average seen so
// far in GEMINI_LOG/SERPER_LOG — check this before committing spend at
// 11x the previously-tested scale, not after.
function estimateBatchCost_(companyCount) {
  const geminiRows = readLogSheet_(GEMINI_LOG_SHEET_NAME, GEMINI_LOG_HEADERS.length);
  const serperRows = readLogSheet_(SERPER_LOG_SHEET_NAME, SERPER_LOG_HEADERS.length);

  const companiesSeen = {};
  geminiRows.forEach(function (r) { if (r[2]) companiesSeen[r[2]] = true; });
  serperRows.forEach(function (r) { if (r[2]) companiesSeen[r[2]] = true; });
  const seenCount = Object.keys(companiesSeen).length;

  if (seenCount === 0) {
    SpreadsheetApp.getUi().alert('No data in GEMINI_LOG/SERPER_LOG yet — run at least one company through the pipeline first so there is a real per-company average to project from.');
    return null;
  }

  let geminiUsd = 0, serperUsd = 0;
  geminiRows.forEach(function (r) { geminiUsd += Number(r[8]) || 0; });
  serperRows.forEach(function (r) { serperUsd += Number(r[5]) || 0; });

  const perCompanyUsd = (geminiUsd + serperUsd) / seenCount;
  const projectedUsd = perCompanyUsd * companyCount;

  SpreadsheetApp.getUi().alert(
    'Cost estimate for ' + companyCount + ' companies:\n\n' +
    'Based on ' + seenCount + ' companies seen so far (avg $' + r6_(perCompanyUsd) + '/company).\n\n' +
    'Projected total: $' + r6_(projectedUsd) + ' USD (~₹' + r4_(projectedUsd * INR_PER_USD) + ')\n\n' +
    'This is a projection from actual logged usage, not a guess — but confirm cost rate tables in TokenLog.js are accurate for your real billing before trusting it for a budget decision.'
  );
  return projectedUsd;
}

// Historical SERPER_LOG rows from before scrape logging existed have no
// matching *_SCRAPE row at all (not zero-cost — simply never written), so
// every section above this one (PER-RUN TOTALS, SERPER — PER-QUERY-TYPE)
// understates any run that predates this fix. There's no way to tell, after
// the fact, which of those old full_text values in RAW_EVIDENCE came from
// the free direct fetch vs. the paid scraper — that distinction was never
// recorded per-row — so this does NOT rewrite SERPER_LOG with invented rows,
// and it does NOT touch the sections above. It writes a clearly-labeled
// correction section using the scrape rate measured from calls logged AFTER
// the fix, so the correction is based on your own observed data rather than
// an assumed "cost is always double".
//
// knownInvoiceUsd (optional): if you have a real number from the Serper
// dashboard/invoice for this account's spend attributable to this pipeline,
// pass it here to surface the gap against SERPER_LOG's logged total. Left
// undefined/null when you don't have one — the section still renders, just
// without that comparison line. This never multiplies SERPER_LOG by a
// guessed factor derived from the gap; it only reports the gap as unexplained
// so it stays visible instead of being silently absorbed into a "corrected"
// number nobody can trace back to evidence.
function writeReconciledHistoricalTotals_(sheet, startRow, serperRows, knownInvoiceUsd) {
  let row = writeSectionHeader_(sheet, startRow, 'RECONCILED HISTORICAL SERPER COST (ESTIMATE)');

  if (serperRows.length === 0) {
    sheet.getRange(row, 1).setValue('(SERPER_LOG is empty — nothing to reconcile)').setFontColor('#888888');
    return row + 1;
  }

  let searchCalls = 0, searchUsd = 0, scrapeCalls = 0, scrapeUsd = 0;
  serperRows.forEach(function (r) {
    const type = String(r[3] || '');
    const usd = Number(r[5]) || 0;
    if (/_SCRAPE$/.test(type)) { scrapeCalls++; scrapeUsd += usd; }
    else { searchCalls++; searchUsd += usd; }
  });

  if (scrapeCalls === 0) {
    sheet.getRange(row, 1).setValue(
      '(No *_SCRAPE rows in SERPER_LOG yet — every row logged so far predates scrape-call logging, ' +
      'so the observed scrape rate is unknown. Run a few companies through the pipeline, then rebuild ' +
      'this sheet to get a real measured correction instead of a guess.)'
    ).setFontColor('#888888');
    row++;

    if (typeof knownInvoiceUsd === 'number' && knownInvoiceUsd > 0) {
      const loggedTotalUsd = searchUsd; // scrapeUsd is 0 in this branch
      const gapUsd = knownInvoiceUsd - loggedTotalUsd;
      row++;
      sheet.getRange(row, 1).setValue('Logged total USD (SERPER_LOG, search calls only)').setFontWeight('bold');
      sheet.getRange(row, 2).setValue(r6_(loggedTotalUsd));
      row++;
      sheet.getRange(row, 1).setValue('Known real spend USD (from account/invoice, provided manually)').setFontWeight('bold');
      sheet.getRange(row, 2).setValue(r6_(knownInvoiceUsd));
      row++;
      sheet.getRange(row, 1).setValue('UNEXPLAINED GAP USD (logged total vs. known real spend)').setFontWeight('bold').setBackground('#f4cccc');
      sheet.getRange(row, 2).setValue(r6_(gapUsd)).setBackground('#f4cccc');
      row++;
      sheet.getRange(row, 1).setValue(
        'This gap is NOT attributed to scraping or any other specific cause — there is no *_SCRAPE data yet to ' +
        'measure a rate from, and no other confirmed explanation. Possible causes include: paid scrapes that ' +
        'happened before scrape-call logging existed and were never recorded, retried/failed calls that still ' +
        'consumed a credit, or the known-spend figure itself being an estimate rather than a reconciled invoice ' +
        'line. Do NOT multiply SERPER_LOG or this total by (real spend / logged total) to "correct" it — that ' +
        'would assert a specific cause for the gap that has not been confirmed. Investigate the gap directly ' +
        '(Serper dashboard usage history, RAW_EVIDENCE/SEARCH_CACHE full_text presence) before treating any ' +
        'multiplier as fact.'
      ).setFontColor('#b00020');
      row++;
    }

    return row;
  }

  const observedScrapeRate = scrapeCalls / searchCalls; // scrape calls per search call, measured post-fix
  const loggedTotalUsd = searchUsd + scrapeUsd;
  // Apply the observed rate to the search-call volume to estimate the scrape
  // cost that would have been logged had this fix existed from the start.
  const estimatedTrueScrapeUsd = searchUsd * observedScrapeRate * (SERPER_SCRAPE_COST_PER_CALL / SERPER_COST_PER_CALL);
  const estimatedTrueTotalUsd = searchUsd + estimatedTrueScrapeUsd;
  const estimatedMissingUsd = estimatedTrueTotalUsd - loggedTotalUsd;

  const headers = ['Metric', 'Value'];
  sheet.getRange(row, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  row++;

  const lines = [
    ['Logged search calls', searchCalls],
    ['Logged search cost USD', r6_(searchUsd)],
    ['Logged scrape calls (post-fix only)', scrapeCalls],
    ['Logged scrape cost USD (post-fix only)', r6_(scrapeUsd)],
    ['Logged total USD (as currently in SERPER_LOG)', r6_(loggedTotalUsd)],
    ['Observed scrape rate (scrape calls / search calls)', r4_(observedScrapeRate)],
    ['Estimated true scrape cost USD (search volume x observed rate)', r6_(estimatedTrueScrapeUsd)],
    ['Estimated true total USD', r6_(estimatedTrueTotalUsd)],
    ['Estimated under-logged amount USD', r6_(estimatedMissingUsd)],
  ];
  lines.forEach(function (pair) {
    sheet.getRange(row, 1, 1, 2).setValues([pair]);
    row++;
  });

  sheet.getRange(row, 1).setValue(
    'Estimate only, from your own measured scrape rate — not a blanket "cost is always double" assumption. ' +
    'SERPER_LOG itself is never modified by this.'
  ).setFontColor('#888888');
  row++;

  return row;
}

// Thin menu wrapper — the reconciled section is now part of the regular
// Cost Summary rebuild, so this just re-runs the full build and points the
// user at the new section rather than duplicating a separate dialog.
function menuReconcileHistoricalSerperCost() {
  buildCostSummary();
  SpreadsheetApp.getUi().alert('See the "RECONCILED HISTORICAL SERPER COST (ESTIMATE)" section at the bottom of COST_SUMMARY.');
}

function menuEstimateBatchCost() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Estimate Batch Cost', 'How many companies do you want to project the cost for?', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;
  const count = parseInt(response.getResponseText(), 10);
  if (!count || count <= 0) {
    ui.alert('Enter a valid positive number.');
    return;
  }
  estimateBatchCost_(count);
}
