// In-run failure buffer, backed by a Script Property (not a sheet) — same
// pattern as CB_STATE/BATCH_STATE. Populated exclusively from
// EventLog.error(), capped at 200 entries to stay well under the ~9KB
// per-property size limit. Feeds retryFailedCompanies_() in Batch.js.
const BATCH_FAILURES_PROP = 'BATCH_FAILURES';
const BATCH_FAILURES_CAP = 200;

function batchFailuresInit_(runId) {
  PropertiesService.getScriptProperties().setProperty(BATCH_FAILURES_PROP, JSON.stringify({
    runId: runId,
    startedAt: new Date().toISOString(),
    entries: []
  }));
}

function batchFailuresPush_(company, website, stage, message) {
  try {
    const raw = PropertiesService.getScriptProperties().getProperty(BATCH_FAILURES_PROP);
    if (!raw) return; // no batch in progress — nothing to accumulate into
    const state = JSON.parse(raw);
    if (state.entries.length >= BATCH_FAILURES_CAP) return;
    state.entries.push({ company: company || '', website: website || '', stage: stage || '', message: message || '', time: new Date().toISOString() });
    PropertiesService.getScriptProperties().setProperty(BATCH_FAILURES_PROP, JSON.stringify(state));
  } catch (e) {
    Logger.log('WARNING: batchFailuresPush_ failed — ' + e.message);
  }
}

function getFailedCompanyNames_() {
  const raw = PropertiesService.getScriptProperties().getProperty(BATCH_FAILURES_PROP);
  if (!raw) return [];
  const state = JSON.parse(raw);
  const distinct = {};
  state.entries.forEach(function (e) { if (e.company) distinct[e.company] = true; });
  return Object.keys(distinct);
}

function showBatchFailures() {
  const raw = PropertiesService.getScriptProperties().getProperty(BATCH_FAILURES_PROP);
  let html;
  if (!raw) {
    html = '<p style="font-family:sans-serif;padding:16px;">No run recorded yet.</p>';
  } else {
    const state = JSON.parse(raw);
    if (state.entries.length === 0) {
      html = '<p style="font-family:sans-serif;padding:16px;color:green;">&#10003; No failures in this run (' + state.runId + ').</p>';
    } else {
      const rows = state.entries.map(function (e) {
        return '<tr><td>' + escapeHtml_(e.company) + '</td><td>' + escapeHtml_(e.website) + '</td><td>' + escapeHtml_(e.stage) + '</td><td>' + escapeHtml_(e.message) + '</td><td>' + escapeHtml_(e.time) + '</td></tr>';
      }).join('');
      html = '<div style="font-family:sans-serif;padding:12px;">' +
        '<p>Run: ' + escapeHtml_(state.runId) + ' — ' + state.entries.length + ' failure(s)</p>' +
        '<table border="1" cellpadding="6" style="border-collapse:collapse;font-size:12px;width:100%;">' +
        '<tr style="background:#f4c7c3;font-weight:bold;"><td>Company</td><td>Website</td><td>Stage</td><td>Error</td><td>Time</td></tr>' +
        rows + '</table></div>';
    }
  }
  const output = HtmlService.createHtmlOutput(html).setWidth(820).setHeight(460);
  SpreadsheetApp.getUi().showModalDialog(output, 'Batch Failures');
}

function escapeHtml_(text) {
  return String(text || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}
