function testApiKeys() {
  const results = [];

  try {
    const res = callSerper_('test');
    results.push('Serper: OK (' + res.length + ' result(s) returned)');
  } catch (e) {
    results.push('Serper: FAILED — ' + e.message);
  }

  try {
    const draft = callGemini_('Reply with exactly this JSON and nothing else: {"status": "ok"}');
    results.push('Gemini: OK (parsed response: ' + JSON.stringify(draft) + ')');
  } catch (e) {
    results.push('Gemini: FAILED — ' + e.message);
  }

  Logger.log(results.join('\n\n'));
}

// Log-only diagnostic for the Patents-noise question (Serper.js's
// callSerperTyped_ comment: the /patents endpoint was never verified against
// Serper's live docs, and returned unrelated results for Sandstorm — e.g. an
// unrelated fire-alarm IoT patent). Run this directly from the script editor
// (select it in the function dropdown, click Run) with a company row
// selected in the HVT sheet first, then check View > Executions for the raw
// response — that tells us whether the endpoint itself is fundamentally
// noisy or just needs a different query shape, before deciding between a
// prompt-skepticism fix and something more involved.
function testSerperPatentsEndpoint() {
  const input = getActiveRowInput_();
  if (!input) return;

  const apiKey = getScriptProp_('SERPER_API_KEY');
  const response = UrlFetchApp.fetch('https://google.serper.dev/patents', {
    method: 'post',
    contentType: 'application/json',
    headers: { 'X-API-KEY': apiKey },
    payload: JSON.stringify({ q: input.name }),
    muteHttpExceptions: true
  });

  Logger.log('Query: ' + input.name);
  Logger.log('Status: ' + response.getResponseCode());
  Logger.log('Body: ' + response.getContentText());
}

// Verifies the search/scrape logging fix without spending any real Serper
// credits: writes a handful of fake SERPER_LOG rows through the same
// TokenLog.logSerper() path the pipeline uses, covering every callType
// combination, then asserts the rates and rollup math came out right. Run
// this from the script editor (select it in the function dropdown, click
// Run), then check View > Executions or the Logger.log output for PASS/FAIL.
//
// This does NOT touch live Serper/Gemini APIs and does NOT need a company
// row selected — it only proves the accounting is correct, not that the
// live callSerper_/callSerperScrape_ calls behave a given way. Pair it with
// testSerperLiveScrapeFallback_ for that.
function testSerperCostLogging() {
  const failures = [];
  function check(label, condition) {
    if (!condition) failures.push(label);
    Logger.log((condition ? 'PASS' : 'FAIL') + ' — ' + label);
  }

  const runId = 'TEST-COST-LOGGING-' + Date.now();
  const company = '__test_company__';

  // 1. A plain search call — should log at SERPER_COST_PER_CALL.
  TokenLog.logSerper(runId, company, 'differentiation', 5, 'search');
  // 2. A scrape call (the paid fallback after fetchPageDirect_ failed) —
  //    should log at SERPER_SCRAPE_COST_PER_CALL under the *_SCRAPE tag,
  //    matching what scrapeWithGuards_ actually does in Serper.js.
  TokenLog.logSerper(runId, company, 'differentiation_SCRAPE', 1, 'scrape');
  // 3. Omitted callType should default to 'search' (backward compatible with
  //    any caller that hasn't been updated to pass callType explicitly).
  TokenLog.logSerper(runId, company, 'moat', 3);
  // 4. A scrape call that failed (fetchPageDirect_ AND callSerperScrape_ both
  //    came back empty) still spent the credit — resultsCount 0, cost > 0.
  TokenLog.logSerper(runId, company, 'product_improvement_SCRAPE', 0, 'scrape');

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SERPER_LOG_SHEET_NAME);
  const lastRow = sheet.getLastRow();
  const allRows = sheet.getRange(2, 1, lastRow - 1, SERPER_LOG_HEADERS.length).getValues();
  const rows = allRows.filter(function (r) { return r[1] === runId; });

  check('wrote exactly 4 rows for this test run', rows.length === 4);
  if (rows.length !== 4) {
    Logger.log('Aborting further checks — row count mismatch means the rest would be meaningless.');
    reportTestResult_('testSerperCostLogging', failures);
    return;
  }

  const byType = {};
  rows.forEach(function (r) { byType[r[3]] = r; });

  check('search row (differentiation) costs SERPER_COST_PER_CALL',
    byType['differentiation'][5] === SERPER_COST_PER_CALL);
  check('scrape row (differentiation_SCRAPE) costs SERPER_SCRAPE_COST_PER_CALL',
    byType['differentiation_SCRAPE'][5] === SERPER_SCRAPE_COST_PER_CALL);
  check('scrape row cost is present even though Results_Count is 1 (flat per-call rate, not per-result)',
    byType['differentiation_SCRAPE'][5] > 0);
  check('omitted callType defaults to the search rate',
    byType['moat'][5] === SERPER_COST_PER_CALL);
  check('a failed scrape (0 results) still logs the credit cost — a blocked scrape is not free',
    byType['product_improvement_SCRAPE'][4] === 0 && byType['product_improvement_SCRAPE'][5] === SERPER_SCRAPE_COST_PER_CALL);

  // INR conversion sanity check — catches a stale INR_PER_USD or a rate
  // change in one currency but not the other.
  check('INR cost matches USD cost x INR_PER_USD for the search row',
    Math.abs(byType['differentiation'][6] - byType['differentiation'][5] * INR_PER_USD) < 1e-9);

  // This is the actual "is the doubling visible now" check: a query that
  // fell through to the paid scraper must show up as 2 billed rows (search +
  // scrape), each with its own nonzero cost — not 1 row, and not a single
  // row silently doubled.
  const queryFellThroughToScrape = byType['differentiation'] && byType['differentiation_SCRAPE'];
  const totalForThatQuery = byType['differentiation'][5] + byType['differentiation_SCRAPE'][5];
  check('a query + its paid scrape together log as 2 separate billed rows summing to 2 credits worth of cost',
    queryFellThroughToScrape && Math.abs(totalForThatQuery - (SERPER_COST_PER_CALL + SERPER_SCRAPE_COST_PER_CALL)) < 1e-9);

  // Clean up the fake rows so this test doesn't pollute COST_SUMMARY.
  // Delete bottom-up so row indices don't shift under us.
  const rowIndexesToDelete = [];
  allRows.forEach(function (r, i) { if (r[1] === runId) rowIndexesToDelete.push(i + 2); });
  rowIndexesToDelete.sort(function (a, b) { return b - a; }).forEach(function (rowIdx) { sheet.deleteRow(rowIdx); });
  Logger.log('Cleaned up ' + rowIndexesToDelete.length + ' test row(s) from ' + SERPER_LOG_SHEET_NAME + '.');

  reportTestResult_('testSerperCostLogging', failures);
}

// Verifies writeReconciledHistoricalTotals_ (CostSummary.js) computes the
// right numbers, without touching the real COST_SUMMARY sheet — writes into
// a throwaway temp sheet instead, reads the printed metric rows back out,
// and checks the arithmetic against hand-computed expected values. Run this
// from the script editor; check Logger.log for PASS/FAIL. Deletes its temp
// sheet when done, pass or fail.
function testReconciledHistoricalTotals() {
  const failures = [];
  function check(label, condition) {
    if (!condition) failures.push(label);
    Logger.log((condition ? 'PASS' : 'FAIL') + ' — ' + label);
  }

  // 10 search calls @ SERPER_COST_PER_CALL, 4 of them followed by a paid
  // scrape @ SERPER_SCRAPE_COST_PER_CALL -> observed scrape rate 4/10 = 0.4.
  const fakeRows = [];
  for (let i = 0; i < 10; i++) {
    fakeRows.push(['', 'r', 'c', 'differentiation', 5, SERPER_COST_PER_CALL, SERPER_COST_PER_CALL * INR_PER_USD]);
  }
  for (let i = 0; i < 4; i++) {
    fakeRows.push(['', 'r', 'c', 'differentiation_SCRAPE', 1, SERPER_SCRAPE_COST_PER_CALL, SERPER_SCRAPE_COST_PER_CALL * INR_PER_USD]);
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tempName = '__test_reconcile_temp__';
  let tempSheet = ss.getSheetByName(tempName);
  if (tempSheet) ss.deleteSheet(tempSheet);
  tempSheet = ss.insertSheet(tempName);

  try {
    writeReconciledHistoricalTotals_(tempSheet, 1, fakeRows);

    const values = tempSheet.getRange(1, 1, tempSheet.getLastRow(), 2).getValues();
    const metric = {};
    values.forEach(function (r) { if (r[0]) metric[r[0]] = r[1]; });

    check('section header written', values[0][0] === 'RECONCILED HISTORICAL SERPER COST (ESTIMATE)');
    check('logged search calls = 10', metric['Logged search calls'] === 10);
    check('logged scrape calls = 4', metric['Logged scrape calls (post-fix only)'] === 4);
    check('observed scrape rate = 0.4', Math.abs(metric['Observed scrape rate (scrape calls / search calls)'] - 0.4) < 1e-9);

    const expectedSearchUsd = 10 * SERPER_COST_PER_CALL;
    const expectedLoggedScrapeUsd = 4 * SERPER_SCRAPE_COST_PER_CALL;
    const expectedLoggedTotal = expectedSearchUsd + expectedLoggedScrapeUsd;
    // observedScrapeRate(0.4) applied to search volume, at the scrape rate —
    // here SERPER_SCRAPE_COST_PER_CALL == SERPER_COST_PER_CALL so this
    // collapses to searchUsd * 0.4, same as the 4 rows already logged.
    const expectedTrueScrapeUsd = expectedSearchUsd * 0.4 * (SERPER_SCRAPE_COST_PER_CALL / SERPER_COST_PER_CALL);
    const expectedTrueTotal = expectedSearchUsd + expectedTrueScrapeUsd;
    const expectedMissing = expectedTrueTotal - expectedLoggedTotal;

    check('logged total USD matches hand-computed value',
      Math.abs(metric['Logged total USD (as currently in SERPER_LOG)'] - r6_(expectedLoggedTotal)) < 1e-9);
    check('estimated true total USD matches hand-computed value',
      Math.abs(metric['Estimated true total USD'] - r6_(expectedTrueTotal)) < 1e-9);
    check('estimated under-logged amount matches hand-computed value (0 here, since rate is self-consistent)',
      Math.abs(metric['Estimated under-logged amount USD'] - r6_(expectedMissing)) < 1e-9);

    // Empty-input edge case: must not throw, must say there's nothing to reconcile.
    const ss2 = SpreadsheetApp.getActiveSpreadsheet();
    const tempName2 = '__test_reconcile_empty__';
    let tempSheet2 = ss2.getSheetByName(tempName2);
    if (tempSheet2) ss2.deleteSheet(tempSheet2);
    tempSheet2 = ss2.insertSheet(tempName2);
    try {
      writeReconciledHistoricalTotals_(tempSheet2, 1, []);
      const emptyText = String(tempSheet2.getRange(2, 1).getValue());
      check('empty SERPER_LOG input reports nothing-to-reconcile without throwing',
        /nothing to reconcile/i.test(emptyText));
    } finally {
      ss2.deleteSheet(tempSheet2);
    }

    // No-scrape-rows-yet edge case: must say the rate is unknown, not divide by zero.
    const tempName3 = '__test_reconcile_noscrapes__';
    let tempSheet3 = ss2.getSheetByName(tempName3);
    if (tempSheet3) ss2.deleteSheet(tempSheet3);
    tempSheet3 = ss2.insertSheet(tempName3);
    try {
      const searchOnlyRows = fakeRows.slice(0, 10);
      writeReconciledHistoricalTotals_(tempSheet3, 1, searchOnlyRows);
      const noScrapeText = String(tempSheet3.getRange(2, 1).getValue());
      check('search-only input (no scrape rows yet) reports the rate as unknown, not a divide-by-zero',
        /scrape rate is unknown/i.test(noScrapeText));
    } finally {
      ss2.deleteSheet(tempSheet3);
    }
  } finally {
    ss.deleteSheet(tempSheet);
  }

  reportTestResult_('testReconciledHistoricalTotals', failures);
}

// Live, credit-spending check that the fetchPageDirect_ -> callSerperScrape_
// fallback actually wires up end to end, and that CostSummary's scrape-rate
// line reflects it. Costs 1 search credit, plus 1 scrape credit IF the
// direct fetch fails for the returned #1 result (most of the time it won't,
// since most pages are ordinary HTML). Run manually, not part of routine
// testing — opt-in because it spends real money and writes real SERPER_LOG
// rows tagged with a throwaway runId (not cleaned up automatically, so you
// can inspect them in the sheet afterward; delete manually when done).
function testSerperLiveScrapeFallback_() {
  const runId = 'TEST-LIVE-SCRAPE-' + Date.now();
  const company = '__test_live_scrape__';
  const query = '"Tata Motors" annual report';

  const before = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SERPER_LOG_SHEET_NAME).getLastRow();

  const results = fetchSerperResults_(company, 'differentiation', query, runId, function () { return callSerper_(query); });

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SERPER_LOG_SHEET_NAME);
  const after = sheet.getLastRow();
  const newRows = sheet.getRange(before + 1, 1, after - before, SERPER_LOG_HEADERS.length).getValues()
    .filter(function (r) { return r[1] === runId; });

  Logger.log('Live query returned ' + results.length + ' result(s).');
  Logger.log('SERPER_LOG rows written for this test run: ' + newRows.length);
  newRows.forEach(function (r) {
    Logger.log('  ' + r[3] + ' | results=' + r[4] + ' | $' + r[5]);
  });

  const hadTopResult = results.length > 0 && results[0].link;
  const scrapeRowLogged = newRows.some(function (r) { return /_SCRAPE$/.test(r[3]); });

  if (!hadTopResult) {
    Logger.log('No #1 result with a link — scrapeWithGuards_ never runs for this query, nothing more to check. Try a different query.');
  } else if (results[0].full_text) {
    Logger.log(scrapeRowLogged
      ? 'Full text was captured via the PAID scrape.serper.dev fallback (fetchPageDirect_ failed) — correctly logged as a _SCRAPE row costing SERPER_SCRAPE_COST_PER_CALL.'
      : 'Full text was captured via the FREE fetchPageDirect_ path — correctly logged NO additional cost (no _SCRAPE row).');
  } else {
    Logger.log('Neither the free direct fetch nor the paid scraper got usable text for this page (paywall/blocked). Check whether a _SCRAPE row was still logged — it should be, since the credit was still spent even on failure.');
  }

  Logger.log('Test rows are tagged runId="' + runId + '" in ' + SERPER_LOG_SHEET_NAME + ' — delete them manually if you don\'t want this test run counted in Cost Summary.');
}

function reportTestResult_(testName, failures) {
  if (failures.length === 0) {
    Logger.log('\n' + testName + ': ALL CHECKS PASSED');
  } else {
    Logger.log('\n' + testName + ': ' + failures.length + ' CHECK(S) FAILED:\n  - ' + failures.join('\n  - '));
  }
}

function listGeminiModels() {
  const apiKey = getScriptProp_('GEMINI_API_KEY');
  const url = 'https://generativelanguage.googleapis.com/v1beta/models?key=' + apiKey;

  const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (response.getResponseCode() !== 200) {
    throw new Error('ListModels error (' + response.getResponseCode() + '): ' + response.getContentText());
  }

  const data = JSON.parse(response.getContentText());
  const names = (data.models || [])
    .filter(function (m) { return (m.supportedGenerationMethods || []).indexOf('generateContent') !== -1; })
    .map(function (m) { return m.name.replace('models/', ''); });

  Logger.log(names.join('\n'));
  return names;
}
