// Adaptive per-API sleep: ratchets UP on 429/503 (multiplicative backoff),
// ratchets DOWN slowly on success (multiplicative decay), independently for
// Serper and Gemini. Not a token bucket — just a scalar sleep duration that
// self-tunes to how much the API is actually tolerating right now, instead
// of a fixed sleep that's either too conservative (wastes time when the API
// is happy) or too aggressive (trips rate limits under real load).
//
// Persisted in a Script Property, NOT module memory — Batch.js reschedules
// every ~10 seconds via a fresh one-shot trigger, and each trigger fire is a
// new execution with reset module state. An in-memory-only sleep value
// would reset to the base every tick and backoff would never actually
// accumulate across a sustained problem.
const RATE_LIMIT_STATE_PROP = 'RATE_LIMIT_STATE';
const RATE_LIMIT_CONFIG = {
  SERPER: { base: 1200, floor: 800, ceiling: 10000 },
  GEMINI: { base: 8000, floor: 4000, ceiling: 30000 }
};
const RATE_LIMIT_BACKOFF_FACTOR = 1.5;
const RATE_LIMIT_DECAY_FACTOR = 0.95;

function loadRateLimitState_() {
  const raw = PropertiesService.getScriptProperties().getProperty(RATE_LIMIT_STATE_PROP);
  if (raw) return JSON.parse(raw);
  return { SERPER: RATE_LIMIT_CONFIG.SERPER.base, GEMINI: RATE_LIMIT_CONFIG.GEMINI.base };
}

function saveRateLimitState_(state) {
  PropertiesService.getScriptProperties().setProperty(RATE_LIMIT_STATE_PROP, JSON.stringify(state));
}

// Called immediately before every Serper/Gemini fetch — never call
// Utilities.sleep() directly anywhere else in the project.
function rateLimit(api) {
  const state = loadRateLimitState_();
  const ms = state[api] || RATE_LIMIT_CONFIG[api].base;
  Utilities.sleep(ms);
}

function rateLimitBackoff(api) {
  const state = loadRateLimitState_();
  const current = state[api] || RATE_LIMIT_CONFIG[api].base;
  state[api] = Math.min(Math.round(current * RATE_LIMIT_BACKOFF_FACTOR), RATE_LIMIT_CONFIG[api].ceiling);
  saveRateLimitState_(state);
}

function rateLimitSuccess(api) {
  const state = loadRateLimitState_();
  const current = state[api] || RATE_LIMIT_CONFIG[api].base;
  const next = Math.max(Math.round(current * RATE_LIMIT_DECAY_FACTOR), RATE_LIMIT_CONFIG[api].floor);
  if (next === current) return; // avoid a property write when already at floor
  state[api] = next;
  saveRateLimitState_(state);
}
