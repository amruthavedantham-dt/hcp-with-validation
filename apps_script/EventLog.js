// User-visible run log — separate from Logger.log (Apps Script's own log,
// which only the developer sees in the editor). EVENT_LOG is a real sheet
// so anyone with the spreadsheet open can see what a batch run actually did,
// without opening the script editor. Never throws — a logging failure must
// never take down the pipeline it's trying to observe.
const EVENT_LOG_SHEET_NAME = 'EVENT_LOG';
const EVENT_LOG_HEADERS = ['Timestamp', 'Run_ID', 'Company', 'Website', 'Stage', 'Level', 'Outcome', 'Message'];

const EventLog = (function () {
  function ensureSheet_() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(EVENT_LOG_SHEET_NAME);
    if (sheet) return sheet;

    sheet = ss.insertSheet(EVENT_LOG_SHEET_NAME);
    sheet.getRange(1, 1, 1, EVENT_LOG_HEADERS.length).setValues([EVENT_LOG_HEADERS]).setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, EVENT_LOG_HEADERS.length).createFilter();
    return sheet;
  }

  function write_(level, runId, company, website, stage, outcome, message) {
    try {
      const sheet = ensureSheet_();
      const row = [new Date(), runId || '', company || '', website || '', stage || '', level, outcome || '', message || ''];
      sheet.appendRow(row);

      const rowNum = sheet.getLastRow();
      const levelCell = sheet.getRange(rowNum, EVENT_LOG_HEADERS.indexOf('Level') + 1);
      const colors = { ERROR: '#f4c7c3', WARN: '#fce8b2', INFO: '#c9daf8' };
      levelCell.setBackground(colors[level] || null);

      if (level === 'ERROR' && typeof batchFailuresPush_ === 'function') {
        batchFailuresPush_(company, website, stage, message);
      }
    } catch (e) {
      // Logging must never break the pipeline it's observing.
      Logger.log('WARNING: EventLog failed to write — ' + e.message);
    }
  }

  return {
    info: function (runId, company, website, stage, outcome, message) {
      write_('INFO', runId, company, website, stage, outcome, message);
    },
    warn: function (runId, company, website, stage, outcome, message) {
      write_('WARN', runId, company, website, stage, outcome, message);
    },
    error: function (runId, company, website, stage, outcome, message) {
      write_('ERROR', runId, company, website, stage, outcome, message);
    }
  };
})();

// Called from Menu.js's manual single-company items, which don't have a
// batch runId — gives every log line a distinguishable, timestamped ID
// without requiring the user to pick one.
function newManualRunId_() {
  return 'manual-' + Date.now();
}
