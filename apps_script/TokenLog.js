// Ledgers for every Gemini and Serper call, feeding CostSummary.js. Every
// call is logged even on failure/empty response — a failed Gemini call
// still spent tokens, and that cost is real.
const GEMINI_LOG_SHEET_NAME = 'GEMINI_LOG';
const GEMINI_LOG_HEADERS = ['Timestamp', 'Run_ID', 'Company', 'Stage', 'Model', 'Input_Tokens', 'Output_Tokens', 'Thinking_Tokens', 'Cost_USD', 'Cost_INR'];
const SERPER_LOG_SHEET_NAME = 'SERPER_LOG';
const SERPER_LOG_HEADERS = ['Timestamp', 'Run_ID', 'Company', 'Query_Type', 'Results_Count', 'Cost_USD', 'Cost_INR'];

// PLACEHOLDER RATES — copied from publicly listed pricing at the time this
// was written, NOT verified against your actual billing. Confirm these
// against your real Gemini/Serper invoices before trusting COST_SUMMARY
// numbers for a real budget decision (e.g. before the 1,100-company run).
const INR_PER_USD = 84;
const GEMINI_COST_RATES = {
  'gemini-2.5-flash': { inputPerMillion: 0.30, outputPerMillion: 2.50, thinkingPerMillion: 2.50 },
  'gemini-2.0-flash': { inputPerMillion: 0.10, outputPerMillion: 0.40, thinkingPerMillion: 0 }
};
// Confirmed against the account's actual recharge: $50 for 50,000 credits =
// $0.001/credit (2026-08-20). Both endpoints bill 1 credit per call. They
// are SEPARATE billable calls: a query whose #1 result falls through to the
// paid scraper (fetchPageDirect_ failed) spends 2 credits total, not 1 — one
// for /search, one for scrape.serper.dev. Keeping the rates as two constants
// (even though they're equal today) means a future Serper repricing of one
// endpoint but not the other doesn't silently mis-cost SERPER_LOG.
const SERPER_COST_PER_CALL = 0.001;
const SERPER_SCRAPE_COST_PER_CALL = 0.001;

function ensureLogSheet_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (sheet) return sheet;
  sheet = ss.insertSheet(name);
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  sheet.setFrozenRows(1);
  return sheet;
}

function calcGeminiCost_(model, inputTokens, outputTokens, thinkingTokens) {
  const rates = GEMINI_COST_RATES[model] || GEMINI_COST_RATES['gemini-2.5-flash'];
  const usd = (inputTokens / 1e6) * rates.inputPerMillion +
              (outputTokens / 1e6) * rates.outputPerMillion +
              (thinkingTokens / 1e6) * rates.thinkingPerMillion;
  return { usd: usd, inr: usd * INR_PER_USD };
}

const TokenLog = {
  logGemini: function (runId, company, stage, model, usage) {
    try {
      const sheet = ensureLogSheet_(GEMINI_LOG_SHEET_NAME, GEMINI_LOG_HEADERS);
      const inputTokens = (usage && usage.promptTokenCount) || 0;
      const outputTokens = (usage && usage.candidatesTokenCount) || 0;
      const thinkingTokens = (usage && usage.thoughtsTokenCount) || 0;
      const cost = calcGeminiCost_(model, inputTokens, outputTokens, thinkingTokens);
      sheet.appendRow([new Date(), runId || '', company || '', stage || '', model, inputTokens, outputTokens, thinkingTokens, cost.usd, cost.inr]);
    } catch (e) {
      Logger.log('WARNING: TokenLog.logGemini failed — ' + e.message);
    }
  },

  // callType: 'search' (default) for /search, /patents, /news — or 'scrape'
  // for a paid scrape.serper.dev call. Each is its own billed credit, so
  // logging must use the matching rate rather than assuming every row is a
  // search call.
  logSerper: function (runId, company, queryType, resultsCount, callType) {
    try {
      const sheet = ensureLogSheet_(SERPER_LOG_SHEET_NAME, SERPER_LOG_HEADERS);
      const rate = callType === 'scrape' ? SERPER_SCRAPE_COST_PER_CALL : SERPER_COST_PER_CALL;
      sheet.appendRow([new Date(), runId || '', company || '', queryType || '', resultsCount || 0, rate, rate * INR_PER_USD]);
    } catch (e) {
      Logger.log('WARNING: TokenLog.logSerper failed — ' + e.message);
    }
  }
};
