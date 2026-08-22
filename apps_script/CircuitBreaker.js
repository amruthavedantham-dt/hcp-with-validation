// Halts an API after too many consecutive failures, so one sustained outage
// (Serper down, Gemini quota exhausted) doesn't burn through the rest of an
// unattended 1,100-company batch retrying a call that will keep failing.
// State lives in Script Properties (survives across the 10-second
// re-trigger cadence Batch.js uses), with an in-memory cache to avoid
// re-reading the property on every single call within one execution.
const CB_STATE_PROP = 'CB_STATE';
const CB_THRESHOLDS = { SERPER: 8, GEMINI: 12 };
const CB_COOLDOWN_MINUTES = 15;
const CB_TRIGGER_HANDLER = 'resumeAfterCircuitBreaker';

const CircuitBreaker = (function () {
  let _state = null;

  function load_() {
    if (_state) return _state;
    const raw = PropertiesService.getScriptProperties().getProperty(CB_STATE_PROP);
    _state = raw ? JSON.parse(raw) : { SERPER: { consecutive: 0, halted: false, haltedAt: '' }, GEMINI: { consecutive: 0, halted: false, haltedAt: '' } };
    return _state;
  }

  function save_() {
    PropertiesService.getScriptProperties().setProperty(CB_STATE_PROP, JSON.stringify(_state));
  }

  return {
    isHalted: function (api) {
      return !!load_()[api].halted;
    },
    recordFailure: function (api) {
      const state = load_();
      state[api].consecutive += 1;
      if (state[api].consecutive >= CB_THRESHOLDS[api] && !state[api].halted) {
        state[api].halted = true;
        state[api].haltedAt = new Date().toISOString();
        save_();
        scheduleResumeTrigger_();
        Logger.log('WARNING: circuit breaker HALTED ' + api + ' after ' + state[api].consecutive + ' consecutive failures — cooling down ' + CB_COOLDOWN_MINUTES + ' min');
        return;
      }
      save_();
    },
    recordSuccess: function (api) {
      const state = load_();
      if (state[api].consecutive === 0 && !state[api].halted) return; // avoid a property write on every single successful call
      state[api].consecutive = 0;
      save_();
    },
    reset: function (api) {
      const state = load_();
      state[api] = { consecutive: 0, halted: false, haltedAt: '' };
      save_();
    }
  };
})();

function scheduleResumeTrigger_() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === CB_TRIGGER_HANDLER) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger(CB_TRIGGER_HANDLER)
    .timeBased()
    .after(CB_COOLDOWN_MINUTES * 60 * 1000)
    .create();
}

// Only clears the halt flags — deliberately does NOT re-invoke the batch
// pipeline itself. Batch.js's own 10-second continuation trigger (if a
// batch is still in progress) will naturally pick the halt-clear up on its
// next tick. Having this handler ALSO call processBatchTick_ would risk two
// overlapping executions hammering the same API and re-tripping the breaker
// immediately.
function resumeAfterCircuitBreaker() {
  CircuitBreaker.reset('SERPER');
  CircuitBreaker.reset('GEMINI');
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === CB_TRIGGER_HANDLER) ScriptApp.deleteTrigger(t);
  });
  Logger.log('INFO: circuit breaker cooldown complete — halts cleared for SERPER and GEMINI');
}

// Distinguishes a TRANSIENT failure (429 rate limit, 5xx server hiccup —
// waiting and retrying is the right response, which is what recordFailure's
// backoff-then-15-min-cooldown already does) from an EXHAUSTED one (out of
// credits, invalid/revoked key, billing problem — waiting 15 minutes and
// retrying is pointless, it'll just fail the same way forever and silently
// eat the rest of an unattended batch run doing nothing). 401/403 are
// always treated as exhausted (auth-level, never self-resolves). Otherwise
// checks the response body for common "you're out" phrasing — deliberately
// broad/case-insensitive since different providers word this differently
// and a false positive here (stopping when it wasn't truly exhausted) is
// far cheaper than a false negative (looping forever on a dead key).
const EXHAUSTION_INDICATORS = [
  'insufficient credit', 'out of credit', 'quota exceeded', 'exceeded your quota',
  'exhausted', 'invalid api key', 'invalid_api_key', 'unauthorized', 'billing',
  'account suspended', 'payment required'
];

function isExhaustedResponse_(responseCode, responseText) {
  if (responseCode === 401 || responseCode === 403 || responseCode === 402) return true;
  const lower = String(responseText || '').toLowerCase();
  return EXHAUSTION_INDICATORS.some(function (phrase) { return lower.indexOf(phrase) !== -1; });
}

// Hard stop, distinct from the circuit breaker's temporary halt — pauses
// the batch (same paused:true flag "Stop Batch (Pause)" uses) and does NOT
// schedule any auto-resume, unlike a normal circuit-breaker cooldown. A
// human needs to actually fix the underlying problem (add credits, replace
// the key) before "Resume Batch" makes sense to click.
function haltBatchForExhaustion_(api, reason, runId) {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty(BATCH_STATE_PROP);
  if (raw) {
    const state = JSON.parse(raw);
    if (!state.paused) {
      state.paused = true;
      props.setProperty(BATCH_STATE_PROP, JSON.stringify(state));
      deleteBatchTriggers_();
    }
  }
  const message = api + ' appears EXHAUSTED (not just rate-limited) — ' + reason +
    '. Batch STOPPED and will NOT auto-resume. Fix the underlying issue (credits/billing/API key), then run "Resume Batch" manually.';
  Logger.log('ERROR: ' + message);
  if (typeof EventLog !== 'undefined') EventLog.error(runId, '', '', 'batch', 'stopped', message);
}
