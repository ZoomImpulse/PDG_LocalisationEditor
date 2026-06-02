// ─────────────────────────────────────────────────────────────────────────────
// ProgressCache.gs — Pre-computed progress statistics for instant sheet picker
// Updated automatically via time-driven trigger (Option A: fully automatic)
// ─────────────────────────────────────────────────────────────────────────────

var PROGRESS_CACHE_SHEET_NAME = '📊 ProgressCache';
var CACHE_UPDATE_INTERVAL_MINUTES = 10; // How often to refresh in background

/**
 * Gets or creates the progress cache sheet.
 * Structure: SheetName | Language | Total | Translated | Reviewed | Disputed | Untranslated | LastUpdated
 *
 * @returns {Sheet} The progress cache sheet
 */
function getProgressCacheSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(PROGRESS_CACHE_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(PROGRESS_CACHE_SHEET_NAME);
    sheet.hideSheet();
    // Add headers
    sheet.getRange(1, 1, 1, 8).setValues([['SheetName', 'Language', 'Total', 'Translated', 'Reviewed', 'Disputed', 'Untranslated', 'LastUpdated']]);
    sheet.getRange(1, 1, 1, 8).setFontWeight('bold');
  }
  return sheet;
}

/**
 * Main function to refresh the entire progress cache.
 * Call this from a time-driven trigger (e.g., every 10 minutes).
 * Scans all sheets and pre-computes progress for all languages.
 */
function refreshProgressCache() {
  var startTime = new Date();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();

  // Build all data in memory FIRST (don't touch cache yet)
  var cacheRows = [];

  sheets.forEach(function(sheet) {
    var sheetName = sheet.getName();
    if (isIgnoredSheet(sheetName)) return;
    if (sheetName === PROGRESS_CACHE_SHEET_NAME) return;
    if (sheetName === STATUS_SHEET_NAME) return;

    var languages = detectLanguageColumns(sheet);
    if (!languages.length) return;

    var lastRow = sheet.getLastRow();
    if (lastRow < 3) return;

    var numRows = lastRow - 2;

    languages.forEach(function(lang) {
      var colIndices = lang.columns.map(function(c) { return c.colIndex; });

      // Read statuses from background colors (fast batch read)
      var statusMap = getStatusFromColors(sheet, 3, numRows, colIndices);

      var stats = {
        total: 0,
        translated: 0,
        reviewed: 0,
        disputed: 0,
        untranslated: 0
      };

      colIndices.forEach(function(colIndex) {
        for (var r = 3; r <= lastRow; r++) {
          stats.total++;

          var status = '';
          if (statusMap[r] && statusMap[r][colIndex]) {
            status = statusMap[r][colIndex];
          }

          if (!status || status === STATUS.UNTRANSLATED) {
            stats.untranslated++;
          } else if (status === STATUS.TRANSLATED) {
            stats.translated++;
          } else if (status === STATUS.REVIEWED) {
            stats.reviewed++;
          } else if (status === STATUS.DISPUTED) {
            stats.disputed++;
          } else if (status === STATUS.STATIC) {
            // Static doesn't count toward completion metrics
            stats.total--; // Don't count static entries
          }
        }
      });

      // Only add if there are actual translation entries
      if (stats.total > 0) {
        cacheRows.push([
          sheetName,
          lang.languageName,
          stats.total,
          stats.translated,
          stats.reviewed,
          stats.disputed,
          stats.untranslated,
          startTime
        ]);
      }
    });
  });

  // Now atomically update the cache: clear old + write new in one operation
  var cacheSheet = getProgressCacheSheet();
  var lastRow = cacheSheet.getLastRow();

  // Clear existing data (keep header row)
  if (lastRow > 1) {
    cacheSheet.deleteRows(2, lastRow - 1);
  }

  // Write new data immediately (no gap where cache is empty for readers)
  if (cacheRows.length > 0) {
    // Ensure we have enough rows
    var rowsNeeded = cacheRows.length;
    var currentRows = cacheSheet.getMaxRows();
    if (currentRows < rowsNeeded + 1) {
      cacheSheet.insertRowsAfter(currentRows, rowsNeeded + 1 - currentRows);
    }
    cacheSheet.getRange(2, 1, cacheRows.length, 8).setValues(cacheRows);
  }

  var duration = (new Date() - startTime) / 1000;
  Logger.log('Progress cache refreshed: ' + cacheRows.length + ' entries in ' + duration + 's');

  // Rebuild the overview sheet from the freshly written cache
  try {
    refreshOverviewSheet();
  } catch (e) {
    Logger.log('Overview sheet refresh failed: ' + e);
  }
}

/**
 * Gets cached progress data for the sheet picker.
 * Returns instantly - no sheet scanning required.
 *
 * @returns {Array} [
 *   {
 *     sheetName,
 *     language,
 *     total,
 *     translated,
 *     reviewed,
 *     disputed,
 *     untranslated,
 *     percentTranslated,
 *     percentReviewed,
 *     lastUpdated
 *   }
 * ]
 */
function getCachedProgress() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(PROGRESS_CACHE_SHEET_NAME);
  if (!sheet) {
    return [];
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }

  var numRows = lastRow - 1;
  var data = sheet.getRange(2, 1, numRows, 8).getValues();
  var result = [];

  // Get list of actual sheet names to filter out deleted sheets
  var actualSheets = {};
  ss.getSheets().forEach(function(s) {
    actualSheets[s.getName()] = true;
  });

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var sheetName = String(row[0]);
    var total = row[2];

    // Only include if sheet still exists and has data
    if (actualSheets[sheetName] && total > 0) {
      result.push({
        sheetName: sheetName,
        language: String(row[1]),
        total: Number(total),
        translated: Number(row[3]),
        reviewed: Number(row[4]),
        disputed: Number(row[5]),
        untranslated: Number(row[6])
      });
    }
  }

  return result;
}

/**
 * Gets cached progress for a specific language.
 * Even faster when you only need one language.
 *
 * @param {string} targetLanguage
 * @returns {Array} Same structure as getCachedProgress but filtered
 */
function getCachedProgressForLanguage(targetLanguage) {
  var allProgress = getCachedProgress();
  return allProgress.filter(function(p) {
    return p.language === targetLanguage;
  });
}

/**
 * Checks if progress cache trigger is installed for this spreadsheet.
 * Auto-installs if missing. Safe to call on every modal open.
 *
 * @returns {boolean} true if trigger exists or was just installed
 */
function ensureProgressCacheTrigger() {
  // Check if trigger already exists
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'refreshProgressCache') {
      return true;
    }
  }

  // Trigger not found - install it
  try {
    ScriptApp.newTrigger('refreshProgressCache')
      .timeBased()
      .everyMinutes(CACHE_UPDATE_INTERVAL_MINUTES)
      .create();

    Logger.log('Progress cache trigger auto-installed for this spreadsheet');

    // Run initial refresh in background (don't block UI)
    try {
      refreshProgressCache();
    } catch (e) {
      Logger.log('Initial cache refresh failed (will retry on next trigger): ' + e);
    }

    return true;
  } catch (e) {
    Logger.log('Failed to auto-install progress cache trigger: ' + e);
    return false;
  }
}

/**
 * Legacy function - redirects to ensureProgressCacheTrigger.
 * Kept for backwards compatibility if anyone calls it manually.
 */
function installProgressCacheTrigger() {
  ensureProgressCacheTrigger();
}

/**
 * Removes all progress cache triggers.
 * Use this if you want to disable automatic updates.
 */
function uninstallProgressCacheTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  var removed = 0;

  for (var i = triggers.length - 1; i >= 0; i--) {
    if (triggers[i].getHandlerFunction() === 'refreshProgressCache') {
      ScriptApp.deleteTrigger(triggers[i]);
      removed++;
    }
  }

  Logger.log('Removed ' + removed + ' progress cache triggers');
}

/**
 * Gets the last cache update timestamp.
 * Useful for showing "last updated" info in the UI.
 *
 * @returns {Date|null}
 */
function getCacheLastUpdated() {
  var cacheSheet = getProgressCacheSheet();
  var lastRow = cacheSheet.getLastRow();

  if (lastRow < 2) {
    return null;
  }

  // Get the most recent timestamp from any row
  var timestamps = cacheSheet.getRange(2, 8, lastRow - 1, 1).getValues();
  var latest = null;

  for (var i = 0; i < timestamps.length; i++) {
    var ts = timestamps[i][0];
    if (ts && (!latest || ts > latest)) {
      latest = ts;
    }
  }

  return latest;
}

/**
 * Forces an immediate cache refresh.
 * Call this when you want fresh data on demand.
 */
function forceRefreshProgressCache() {
  refreshProgressCache();
}
