function runSerperGrounding(companyName, website, runId) {
  runId = runId || newManualRunId_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName(RAW_EVIDENCE_SHEET_NAME);
  if (!rawSheet) throw new Error('RAW_EVIDENCE sheet not found. Run setupHVTSheet() first.');

  clearRawEvidenceForCompany_(rawSheet, companyName);
  EventLog.info(runId, companyName, website, 'serper_grounding', 'start', 'Serper grounding started');

  const rowsToWrite = [];
  const now = new Date();

  function fireQuery(pillar, query) {
    const results = fetchSerperResults_(companyName, pillar, query, runId, function () { return callSerper_(query); });
    results.forEach(function (r) {
      rowsToWrite.push([companyName, pillar, query, r.link || '', r.title || '', r.snippet || '', r.full_text || '', now]);
    });
  }

  Object.keys(SERPER_QUERIES_BY_PILLAR).forEach(function (pillar) {
    SERPER_QUERIES_BY_PILLAR[pillar].forEach(function (template) {
      fireQuery(pillar, template.replace('{company}', companyName));
    });
  });

  // Site-restricted queries — only possible once a website is supplied.
  // Pushes the company's own deeper pages ahead of thin social posts.
  if (website) {
    const domain = extractDomain_(website);
    Object.keys(SERPER_SITE_QUERIES_BY_PILLAR).forEach(function (pillar) {
      SERPER_SITE_QUERIES_BY_PILLAR[pillar].forEach(function (template) {
        fireQuery(pillar, template.replace('{domain}', domain));
      });
    });
  }

  // Patents/News search types — one call each per company, gives claims a
  // real path to something other than "self-claimed" (an actual patent
  // record or actual journalism, instead of the company's own website/social).
  [
    { endpoint: 'patents', tag: 'patents' },
    { endpoint: 'news', tag: 'news' }
  ].forEach(function (typed) {
    const results = fetchSerperResults_(companyName, typed.tag, companyName, runId, function () { return callSerperTyped_(typed.endpoint, companyName); });
    results.forEach(function (r) {
      rowsToWrite.push([companyName, typed.tag, companyName, r.link || '', r.title || '', r.snippet || '', r.full_text || '', now]);
    });
  });

  if (rowsToWrite.length > 0) {
    rawSheet.getRange(rawSheet.getLastRow() + 1, 1, rowsToWrite.length, RAW_EVIDENCE_HEADERS.length).setValues(rowsToWrite);
  }

  EventLog.info(runId, companyName, website, 'serper_grounding', 'done', rowsToWrite.length + ' evidence rows written');
  return rowsToWrite.length;
}

// Fired as its own step, after `sector` is known (Identity Resolution or
// manual entry) — see SERPER_SECTOR_QUERIES in Config.js for why this exists.
// Appends to RAW_EVIDENCE (does NOT clear first) since it's meant to add to
// whatever runSerperGrounding already staged for this company, tagged with
// the same pillar names (alignment_gap, moat) so Gemini Structure picks it
// up automatically without any prompt changes.
function runSectorContextGrounding(companyName, sector, runId) {
  runId = runId || newManualRunId_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName(RAW_EVIDENCE_SHEET_NAME);
  if (!rawSheet) throw new Error('RAW_EVIDENCE sheet not found. Run setupHVTSheet() first.');

  const rowsToWrite = [];
  const now = new Date();

  Object.keys(SERPER_SECTOR_QUERIES).forEach(function (pillar) {
    SERPER_SECTOR_QUERIES[pillar].forEach(function (template) {
      const query = template.replace('{sector}', sector);
      const results = fetchSerperResults_(companyName, pillar, query, runId, function () { return callSerper_(query); });
      results.forEach(function (r) {
        rowsToWrite.push([companyName, pillar, query, r.link || '', r.title || '', r.snippet || '', r.full_text || '', now]);
      });
    });
  });

  if (rowsToWrite.length > 0) {
    rawSheet.getRange(rawSheet.getLastRow() + 1, 1, rowsToWrite.length, RAW_EVIDENCE_HEADERS.length).setValues(rowsToWrite);
  }

  EventLog.info(runId, companyName, '', 'sector_context_grounding', 'done', rowsToWrite.length + ' evidence rows added (sector: "' + sector + '")');
  return rowsToWrite.length;
}

// The single entry point every Serper /search-shaped query goes through:
// cache check first, then circuit breaker, then rate limiter, then the live
// call, then scrape the #1 result's full page text, then cache write +
// token/serper logging. queryType is the cache key dimension (pillar name,
// or 'patents'/'news') — companyName+queryType+query together form the
// cache key. Scraping happens HERE (not at the caller) specifically so the
// full_text ends up in the cache too — a future cache hit needs the same
// full page text a live run would have gotten, not just the bare snippet.
function fetchSerperResults_(companyName, queryType, query, runId, liveCallFn) {
  const cached = getCachedSearch_(companyName, queryType, query);
  if (cached) return cached;

  if (CircuitBreaker.isHalted('SERPER')) {
    EventLog.warn(runId, companyName, '', 'serper_call', 'skipped', 'SERPER circuit breaker is halted — query skipped: "' + query + '"');
    return [];
  }

  rateLimit('SERPER');
  let results;
  try {
    results = liveCallFn();
    CircuitBreaker.recordSuccess('SERPER');
    rateLimitSuccess('SERPER');
  } catch (e) {
    CircuitBreaker.recordFailure('SERPER');
    rateLimitBackoff('SERPER');
    EventLog.error(runId, companyName, '', 'serper_call', 'error', 'Serper call failed for "' + query + '": ' + e.message);
    return [];
  }

  TokenLog.logSerper(runId, companyName, queryType, results.length, 'search');

  // Only scrape the #1 result per query (cost control) — this is what makes
  // it real "search+fetch" per §4 instead of snippet-only, which was too
  // thin to carry specifics like patent counts or product names.
  if (results.length > 0 && results[0].link) {
    results[0].full_text = scrapeWithGuards_(results[0].link, companyName, queryType, runId);
  }

  writeCacheEntries_(companyName, queryType, query, results);
  return results;
}

// Best-effort full-page fetch. Tries a FREE direct fetch first and only pays
// for scrape.serper.dev when that fails. The sibling pipelines (Manufacturing
// c4.js, RocketReach techLeadership.js) have always scraped Tofler/Zauba with
// a bare UrlFetchApp request at zero API cost, and most pages surfaced here
// are ordinary HTML that yields to the same treatment — this pipeline was the
// only one routing every page fetch through a paid endpoint, at up to one
// extra billable call per query. Serper's scraper stays as the fallback
// because it genuinely wins on JS-rendered and bot-blocked pages a direct
// request can't reach.
//
// Scrape failures (paywalls, timeouts — common and expected) still do NOT
// count toward the circuit breaker the way a hard Serper /search API error
// does; a blocked scrape isn't evidence Serper itself is unreliable.
//
// The paid fallback is now logged as queryType + '_SCRAPE', which it never
// was before: scrape.serper.dev calls were invisible to BOTH SERPER_LOG and
// the circuit breaker, so no amount of reading the cost summary could show
// what they cost or whether they were silently failing.
function scrapeWithGuards_(url, companyName, queryType, runId) {
  const direct = fetchPageDirect_(url);
  if (direct) return direct;

  rateLimit('SERPER');
  const text = callSerperScrape_(url);
  TokenLog.logSerper(runId, companyName, String(queryType || '') + '_SCRAPE', text ? 1 : 0, 'scrape');
  return text;
}

// Free page fetch — no API key, no credits, bounded only by the Apps Script
// UrlFetchApp quota. Returns '' whenever the page can't be read usefully, so
// the caller knows to fall back to the paid scraper. Same browser User-Agent
// and paywall detection the sibling pipelines already use.
function fetchPageDirect_(url) {
  try {
    const response = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      followRedirects:    true,
      headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120 Safari/537.36' }
    });
    if (response.getResponseCode() !== 200) return '';

    const html = response.getContentText();
    if (!html || html.length < 500) return '';
    if (/get pro|subscribe|unlock|sign up to view|login to view|please login/i.test(html)) return '';

    const text = htmlToText_(html);
    // A page that strips down to almost nothing is usually JS-rendered — the
    // markup arrived but the content didn't. That is exactly the case Serper's
    // scraper handles, so treat it as a miss rather than returning a shell and
    // letting thin evidence through as if it were a real page read.
    return text.length >= 200 ? text.slice(0, 3000) : '';
  } catch (e) {
    return '';
  }
}

// Matches what callSerperScrape_ returns (plain text, 3000-char cap) so both
// paths hand the same shape of evidence to RAW_EVIDENCE and the cache.
function htmlToText_(html) {
  return String(html)
    .replace(/<script[\s\S]*?<\/script>/gi, ' ')
    .replace(/<style[\s\S]*?<\/style>/gi, ' ')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&quot;/gi, '"')
    .replace(/&#39;/gi, "'")
    .replace(/\s+/g, ' ')
    .trim();
}

// Was previously a full read-all/clearContent-all/write-back-the-rest of
// RAW_EVIDENCE on EVERY company's Serper Grounding call — safe only if the
// execution never gets interrupted between the clear and the write-back. In
// a batch run, an interruption there (6-min hard kill, quota error) wiped
// every OTHER company's staged evidence, not just this one's — and the
// operation got more expensive as the batch progressed, since RAW_EVIDENCE
// is never pruned after a company's Gemini step consumes it. Fixed by only
// touching THIS company's own rows in place — other companies' rows are
// never read or rewritten, so an interruption mid-loop can, at worst, leave
// a few of this company's own stale rows uncleared (mixed with fresh ones),
// never affect anyone else's data.
function clearRawEvidenceForCompany_(rawSheet, companyName) {
  const lastRow = rawSheet.getLastRow();
  if (lastRow <= 1) return;
  const targetLower = companyName.trim().toLowerCase();
  const names = rawSheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < names.length; i++) {
    if (String(names[i][0]).trim().toLowerCase() === targetLower) {
      rawSheet.getRange(i + 2, 1, 1, RAW_EVIDENCE_HEADERS.length).clearContent();
    }
  }
}

// Bumped from 3 to 5 (Phase E) — Natesan's real CIN was sitting in a page
// that ranked #4 in Serper's own results for the exact right query; the old
// top-3 cutoff meant it was silently never even staged as evidence. Measure
// the added cost via SERPER_LOG before assuming this is free.
const SERPER_RESULTS_PER_QUERY = 5;

function callSerper_(query) {
  const apiKey = getScriptProp_('SERPER_API_KEY');
  const response = UrlFetchApp.fetch('https://google.serper.dev/search', {
    method: 'post',
    contentType: 'application/json',
    headers: { 'X-API-KEY': apiKey },
    payload: JSON.stringify({ q: query }),
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    const code = response.getResponseCode();
    const body = response.getContentText();
    if (isExhaustedResponse_(code, body)) {
      haltBatchForExhaustion_('SERPER', 'HTTP ' + code + ' on query "' + query + '": ' + body);
    }
    throw new Error('Serper API error (' + code + ') for query "' + query + '": ' + body);
  }

  const data = JSON.parse(response.getContentText());
  const organic = data.organic || [];
  return organic.slice(0, SERPER_RESULTS_PER_QUERY).map(function (r) {
    return { title: r.title, link: r.link, snippet: r.snippet };
  });
}

// Patents/News use different Serper endpoints than /search. News is a
// well-established part of Serper's API; Patents is less certain since it
// wasn't verified against their live docs — this degrades quietly (empty
// array + a log line) rather than breaking the whole grounding run if the
// endpoint or response shape turns out to differ. Errors here are NOT
// thrown (unlike callSerper_) so they don't trip the circuit breaker the
// same way a real /search failure does — this endpoint's own unreliability
// (confirmed via testSerperPatentsEndpoint) is already a known, accepted
// limitation, not a signal Serper itself is down.
function callSerperTyped_(endpoint, query) {
  const apiKey = getScriptProp_('SERPER_API_KEY');
  try {
    const response = UrlFetchApp.fetch('https://google.serper.dev/' + endpoint, {
      method: 'post',
      contentType: 'application/json',
      headers: { 'X-API-KEY': apiKey },
      payload: JSON.stringify({ q: query }),
      muteHttpExceptions: true
    });
    if (response.getResponseCode() !== 200) {
      const code = response.getResponseCode();
      const body = response.getContentText();
      // This endpoint normally degrades quietly (see comment above) — but an
      // exhausted key/credits is a real problem regardless of which endpoint
      // surfaced it, so that specific case still triggers a hard stop.
      if (isExhaustedResponse_(code, body)) {
        haltBatchForExhaustion_('SERPER', 'HTTP ' + code + ' on /' + endpoint + ' query "' + query + '": ' + body);
      }
      Logger.log('WARNING: Serper /' + endpoint + ' returned ' + code + ' for query "' + query + '" — skipping');
      return [];
    }
    const data = JSON.parse(response.getContentText());
    const results = data.organic || data.news || data.patents || [];
    return results.slice(0, SERPER_RESULTS_PER_QUERY).map(function (r) {
      return { title: r.title || '', link: r.link || '', snippet: r.snippet || '' };
    });
  } catch (e) {
    Logger.log('WARNING: Serper /' + endpoint + ' call failed for query "' + query + '": ' + e.message);
    return [];
  }
}

// Best-effort full-page fetch. Some pages block scraping or time out — that's
// not fatal, we just fall back to the snippet for that row.
function callSerperScrape_(url) {
  const apiKey = getScriptProp_('SERPER_API_KEY');
  try {
    const response = UrlFetchApp.fetch('https://scrape.serper.dev', {
      method: 'post',
      contentType: 'application/json',
      headers: { 'X-API-KEY': apiKey },
      payload: JSON.stringify({ url: url }),
      muteHttpExceptions: true
    });
    if (response.getResponseCode() !== 200) return '';
    const data = JSON.parse(response.getContentText());
    return (data.text || '').slice(0, 3000);
  } catch (e) {
    return '';
  }
}

// ---- SEARCH_CACHE: check-before-fetch cache, separate from RAW_EVIDENCE ----
// See the comment on SEARCH_CACHE_SHEET_NAME in Config.js for why this is a
// separate sheet. In-memory Map loaded once per execution (GAS resets
// module state on every trigger fire, so this only helps within one tick,
// not across ticks — the sheet itself is the durable cache).
let _searchCacheMap = null;

function loadSearchCacheMap_() {
  if (_searchCacheMap) return _searchCacheMap;
  _searchCacheMap = {};
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SEARCH_CACHE_SHEET_NAME);
  if (!sheet) return _searchCacheMap;
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return _searchCacheMap;

  const values = sheet.getRange(2, 1, lastRow - 1, RAW_EVIDENCE_HEADERS.length).getValues();
  values.forEach(function (row) {
    const key = cacheKey_(row[0], row[1], row[2]);
    if (!_searchCacheMap[key]) _searchCacheMap[key] = [];
    _searchCacheMap[key].push({ title: row[4], link: row[3], snippet: row[5], full_text: row[6] });
  });
  return _searchCacheMap;
}

function cacheKey_(companyName, queryType, query) {
  return (String(companyName).trim() + '|' + String(queryType).trim() + '|' + String(query).trim()).toLowerCase();
}

function getCachedSearch_(companyName, queryType, query) {
  const map = loadSearchCacheMap_();
  const hit = map[cacheKey_(companyName, queryType, query)];
  if (!hit) return null;
  return hit.map(function (r) { return { title: r.title, link: r.link, snippet: r.snippet, full_text: r.full_text, fromCache: true }; });
}

function writeCacheEntries_(companyName, queryType, query, results) {
  if (!results || results.length === 0) return;
  const sheet = ensureSearchCacheSheet_();
  const now = new Date();
  const rows = results.map(function (r) {
    return [companyName, queryType, query, r.link || '', r.title || '', r.snippet || '', r.full_text || '', now];
  });
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, RAW_EVIDENCE_HEADERS.length).setValues(rows);

  const map = loadSearchCacheMap_();
  const key = cacheKey_(companyName, queryType, query);
  map[key] = results.map(function (r) { return { title: r.title, link: r.link, snippet: r.snippet, full_text: r.full_text || '' }; });
}

function ensureSearchCacheSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SEARCH_CACHE_SHEET_NAME);
  if (sheet) return sheet;
  sheet = ss.insertSheet(SEARCH_CACHE_SHEET_NAME);
  sheet.getRange(1, 1, 1, RAW_EVIDENCE_HEADERS.length).setValues([RAW_EVIDENCE_HEADERS]).setFontWeight('bold');
  sheet.setFrozenRows(1);
  return sheet;
}

// Menu-triggered, deliberate cache refresh — not automatic, since the whole
// point of no-TTL is to stay stable unless you explicitly ask otherwise.
// Same targeted-in-place fix as clearRawEvidenceForCompany_ — only touches
// this company's own rows, never reads/rewrites everyone else's cache.
function clearCacheForCompany_(companyName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SEARCH_CACHE_SHEET_NAME);
  if (!sheet) return 0;
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 0;
  const targetLower = companyName.trim().toLowerCase();
  const names = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  let removed = 0;
  for (let i = 0; i < names.length; i++) {
    if (String(names[i][0]).trim().toLowerCase() === targetLower) {
      sheet.getRange(i + 2, 1, 1, RAW_EVIDENCE_HEADERS.length).clearContent();
      removed++;
    }
  }
  _searchCacheMap = null;
  return removed;
}

function clearCacheAll_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SEARCH_CACHE_SHEET_NAME);
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, RAW_EVIDENCE_HEADERS.length).clearContent();
  _searchCacheMap = null;
}
