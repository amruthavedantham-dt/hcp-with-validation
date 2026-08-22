// A real orchestrator across every company in the HVT sheet — in contrast
// to Menu.js's one-company-at-a-time items. Explicitly out of scope for the
// original 5-company method test (see project non-goals), built once that
// test was validated and requested directly.
//
// Apps Script kills any single execution — including trigger-invoked ones —
// at 6 minutes. This can't loop through every company in one call, so it
// processes one (company, step) unit at a time, persists progress to Script
// Properties after every single unit, and re-schedules itself via a
// short-delay time-based trigger to continue exactly where it left off. A
// crash, quota error, or manual stop mid-run loses at most one unit of work,
// never the whole batch.
//
// A single bad company (API timeout, missing data) is logged and skipped
// rather than halting the whole batch — see the try/catch in
// processBatchTick_. Check View > Executions in the Apps Script editor, or
// the EVENT_LOG sheet, to watch progress — there is no UI popup for
// tick-to-tick progress since trigger-invoked executions have no UI context
// to alert into.

const BATCH_STATE_PROP = 'BATCH_STATE';
const BATCH_TRIGGER_HANDLER = 'processBatchTick_';
const BATCH_TIME_BUDGET_MS = 5 * 60 * 1000; // 5 min — buffer under the 6-min execution cap
// company_summary deliberately NOT included as a default step — it's a 5th
// Gemini call per company (real added cost at 1,100-company scale) for a
// paragraph that's largely redundant with what REVIEW_QUEUE already shows.
// Still available as its own manual menu item ("5. Run Company Summary")
// for anyone who wants it selectively.
const BATCH_STEPS = ['serper', 'identity', 'sector', 'gemini', 'validation'];

function startBatchProcessing() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HVT_SHEET_NAME);
  if (!sheet) throw new Error('HVT sheet not found. Run setupHVTSheet() first.');

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    ui.alert('No company rows found in the HVT sheet.');
    return;
  }

  const nameCol = HVT_HEADERS.indexOf('company_name') + 1;
  const values = sheet.getRange(2, nameCol, lastRow - 1, 1).getValues();
  const companies = values.map(function (r) { return String(r[0] || '').trim(); }).filter(function (n) { return n; });

  if (companies.length === 0) {
    ui.alert('No company names found in column A.');
    return;
  }

  startBatchForCompanies_(companies, 'Batch started for ' + companies.length + ' companies (' + BATCH_STEPS.length + ' steps each). ' +
    'It will keep running in the background across multiple executions — check View > Executions or EVENT_LOG to watch progress. ' +
    'Run "Stop Batch" from this menu to pause at any time.');
}

// Filters the company list down to rows that haven't completed Gemini
// Structure yet (triage_status still empty) before starting — lets you add
// new companies to an already-processed sheet and only pay for the new
// ones, instead of "Start Batch" re-running everyone from scratch every
// time. This is a checkpoint that survives across full batch completions
// (unlike BATCH_STATE, which only helps mid-run) — the sheet itself is the
// source of truth for "what's already done." Combined with
// resumeBatchProcessing() for the mid-run case, this covers both senses of
// "continue from where it stopped."
function startBatchForIncomplete() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HVT_SHEET_NAME);
  if (!sheet) throw new Error('HVT sheet not found. Run setupHVTSheet() first.');

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    ui.alert('No company rows found in the HVT sheet.');
    return;
  }

  const nameCol = HVT_HEADERS.indexOf('company_name') + 1;
  const triageCol = HVT_HEADERS.indexOf('triage_status') + 1;
  const values = sheet.getRange(2, 1, lastRow - 1, HVT_HEADERS.length).getValues();

  const companies = values
    .filter(function (row) { return String(row[triageCol - 1] || '').trim() === ''; })
    .map(function (row) { return String(row[nameCol - 1] || '').trim(); })
    .filter(function (n) { return n; });

  if (companies.length === 0) {
    ui.alert('No incomplete companies found — every row already has a triage_status. Use "Start Batch" if you want to reprocess everyone from scratch.');
    return;
  }

  startBatchForCompanies_(companies, 'Batch started for ' + companies.length + ' NEW/INCOMPLETE compan' + (companies.length === 1 ? 'y' : 'ies') +
    ' (skipping rows that already have a triage_status): ' + companies.join(', '));
}

// Shared by startBatchProcessing(), startBatchForIncomplete(), and
// retryFailedCompanies_() — all three just build a company list and hand it
// to the same tick loop.
//
// Guards against a race found in testing: clicking "Start Batch" twice in
// quick succession fires two concurrent executions that both try to
// initialize BATCH_STATE, the second silently clobbering the first's
// progress. Refusing to start a second batch while one is already active
// (not paused) closes that — the fix is "refuse," not "queue," since there's
// no safe way to merge two different company lists into one run.
function startBatchForCompanies_(companies, alertMessage) {
  const props = PropertiesService.getScriptProperties();
  const existingRaw = props.getProperty(BATCH_STATE_PROP);
  if (existingRaw) {
    const existing = JSON.parse(existingRaw);
    if (!existing.paused) {
      SpreadsheetApp.getUi().alert(
        'A batch is already running (company ' + (existing.index + 1) + '/' + existing.companies.length + '). ' +
        'Wait for it to finish, or run "Stop Batch (Pause)" first if you want to start a different one.'
      );
      return;
    }
  }

  const runId = Utilities.getUuid();
  const state = { runId: runId, companies: companies, index: 0, step: 0, paused: false };
  props.setProperty(BATCH_STATE_PROP, JSON.stringify(state));
  deleteBatchTriggers_();
  batchFailuresInit_(runId);

  EventLog.info(runId, '', '', 'batch', 'start', 'Batch started for ' + companies.length + ' companies: ' + companies.join(', '));
  Logger.log('INFO: batch started (run ' + runId + ') for ' + companies.length + ' companies: ' + companies.join(', '));
  SpreadsheetApp.getUi().alert(alertMessage);

  processBatchTick_();
}

// Pauses rather than cancels — sets a flag and removes the continuation
// trigger, but keeps BATCH_STATE intact so resumeBatchProcessing() can pick
// up exactly where it left off. Use "Start Batch" if you actually want to
// discard progress and begin a fresh company list.
function stopBatchProcessing() {
  const raw = PropertiesService.getScriptProperties().getProperty(BATCH_STATE_PROP);
  if (!raw) {
    SpreadsheetApp.getUi().alert('No batch was running.');
    return;
  }
  const state = JSON.parse(raw);
  state.paused = true;
  PropertiesService.getScriptProperties().setProperty(BATCH_STATE_PROP, JSON.stringify(state));
  deleteBatchTriggers_();
  EventLog.info(state.runId, '', '', 'batch', 'paused', 'Batch paused at company ' + (state.index + 1) + '/' + state.companies.length);
  Logger.log('INFO: batch paused.');
  SpreadsheetApp.getUi().alert(
    'Batch paused at company ' + (state.index + 1) + '/' + state.companies.length + '. ' +
    'Run "Resume Batch" to continue exactly where it left off, or "Start Batch" to discard this progress and begin fresh.'
  );
}

function resumeBatchProcessing() {
  const raw = PropertiesService.getScriptProperties().getProperty(BATCH_STATE_PROP);
  if (!raw) {
    SpreadsheetApp.getUi().alert('No paused batch to resume — use "Start Batch" instead.');
    return;
  }
  const state = JSON.parse(raw);
  state.paused = false;
  PropertiesService.getScriptProperties().setProperty(BATCH_STATE_PROP, JSON.stringify(state));
  EventLog.info(state.runId, '', '', 'batch', 'resumed', 'Batch resumed at company ' + (state.index + 1) + '/' + state.companies.length);
  Logger.log('INFO: batch resumed.');
  SpreadsheetApp.getUi().alert('Batch resumed from company ' + (state.index + 1) + '/' + state.companies.length + '.');
  processBatchTick_();
}

// Re-runs only the companies that had at least one failed step in the most
// recent batch run (per the BATCH_FAILURES buffer) — the highest-leverage
// addition for 1,100-company operation, since partial per-company failures
// are expected at that volume and re-running everything to fix a handful is
// wasteful.
function retryFailedCompanies_() {
  const failed = getFailedCompanyNames_();
  if (failed.length === 0) {
    SpreadsheetApp.getUi().alert('No failed companies recorded from the last run.');
    return;
  }
  startBatchForCompanies_(failed, 'Retrying ' + failed.length + ' failed compan' + (failed.length === 1 ? 'y' : 'ies') + ' from the last run: ' + failed.join(', '));
}

function deleteBatchTriggers_() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === BATCH_TRIGGER_HANDLER) {
      ScriptApp.deleteTrigger(t);
    }
  });
}

// The actual work loop — runs once when the batch starts, and again every
// time its own continuation trigger fires. Processes (company, step) units
// sequentially, saving progress after each one, until either the whole
// batch is done, this execution's time budget runs out, the batch was
// paused, or the circuit breaker halts an API mid-run.
function processBatchTick_() {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty(BATCH_STATE_PROP);
  if (!raw) {
    Logger.log('INFO: no batch in progress — nothing to do.');
    return;
  }

  const state = JSON.parse(raw);
  if (state.paused) {
    Logger.log('INFO: batch is paused — not processing. Run "Resume Batch" to continue.');
    return;
  }

  const runId = state.runId;
  const startTime = Date.now();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HVT_SHEET_NAME);
  const websiteCol = HVT_HEADERS.indexOf('website') + 1;
  const sectorCol = HVT_HEADERS.indexOf('sector') + 1;

  while (state.index < state.companies.length && (Date.now() - startTime) < BATCH_TIME_BUDGET_MS) {
    // Re-read paused fresh from Script Properties every iteration, not just
    // once at the top of this function — an exhausted-API detection deep
    // inside a Serper/Gemini call (haltBatchForExhaustion_) sets this
    // mid-loop, and without this check the loop would keep burning through
    // more companies that will fail the exact same way for the rest of this
    // execution's 5-minute budget before ever noticing.
    const freshRaw = props.getProperty(BATCH_STATE_PROP);
    if (!freshRaw || JSON.parse(freshRaw).paused) {
      Logger.log('INFO: batch — stopping mid-loop, likely paused by an exhausted-API halt. Check EVENT_LOG.');
      return;
    }

    if (CircuitBreaker.isHalted('SERPER') && CircuitBreaker.isHalted('GEMINI')) {
      EventLog.warn(runId, '', '', 'batch', 'halted', 'Both circuit breakers are halted — pausing batch until cooldown clears them.');
      Logger.log('WARNING: batch — both circuit breakers halted, stopping this execution without scheduling a retry (the breaker\'s own resume trigger will let a future tick pick this back up).');
      props.setProperty(BATCH_STATE_PROP, JSON.stringify(state));
      return;
    }

    const companyName = state.companies[state.index];
    const stepName = BATCH_STEPS[state.step];
    const row = findOrCreateCompanyRow_(companyName).getRow();
    const website = String(sheet.getRange(row, websiteCol).getValue() || '').trim();

    try {
      if (stepName === 'serper') {
        runSerperGrounding(companyName, website, runId);
      } else if (stepName === 'identity') {
        if (!website) {
          EventLog.warn(runId, companyName, website, 'identity', 'skipped', 'No website — identity resolution skipped');
        } else {
          runIdentityResolution(companyName, website, runId);
        }
      } else if (stepName === 'sector') {
        const sector = String(sheet.getRange(row, sectorCol).getValue() || '').trim();
        if (!sector) {
          EventLog.warn(runId, companyName, website, 'sector_context', 'skipped', 'No sector — sector-context grounding skipped (identity resolution may have failed or found nothing)');
        } else {
          runSectorContextGrounding(companyName, sector, runId);
        }
      } else if (stepName === 'gemini') {
        runGeminiStructure(companyName, runId);
      } else if (stepName === 'validation') {
        runGeminiValidation(companyName, runId);
      } else if (stepName === 'company_summary') {
        runCompanySummary(companyName, runId);
      }
    } catch (e) {
      // One bad company must never take down the whole batch — log loudly
      // and move on. Whatever step failed simply leaves that field blank or
      // stale for this company, visible on manual review like everything
      // else in this tool.
      EventLog.error(runId, companyName, website, stepName, 'error', e.message);
      Logger.log('ERROR: batch — "' + stepName + '" failed for "' + companyName + '": ' + e.message + ' — skipping to next step');
    }

    state.step += 1;
    if (state.step >= BATCH_STEPS.length) {
      state.step = 0;
      state.index += 1;
    }
    props.setProperty(BATCH_STATE_PROP, JSON.stringify(state));
  }

  if (state.index >= state.companies.length) {
    EventLog.info(runId, '', '', 'batch', 'done', 'Batch complete — all ' + state.companies.length + ' companies processed.');
    Logger.log('INFO: batch complete — all ' + state.companies.length + ' companies processed.');
    props.deleteProperty(BATCH_STATE_PROP);
    deleteBatchTriggers_();
    return;
  }

  // Not done yet — schedule the next tick and let this execution end
  // cleanly well before the 6-minute hard limit, rather than risk a
  // mid-step kill that could leave a company's row half-written.
  deleteBatchTriggers_();
  ScriptApp.newTrigger(BATCH_TRIGGER_HANDLER)
    .timeBased()
    .after(10 * 1000)
    .create();
  Logger.log('INFO: batch paused for this execution — next tick in ~10s. Progress: company ' +
    (state.index + 1) + '/' + state.companies.length + ', next step "' + BATCH_STEPS[state.step] + '".');
}
