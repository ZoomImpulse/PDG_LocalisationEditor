// ─────────────────────────────────────────────────────────────────────────────
// StatusStorage.gs — Alternative status storage using a separate sheet
// instead of inserted columns. Less disruptive to the original data structure.
// ─────────────────────────────────────────────────────────────────────────────

var STATUS_SHEET_NAME = '📋 Status';

// Cache for status data to minimize sheet reads
var statusCache = {};

/**
 * Gets or creates the status tracking sheet.
 * Structure: SheetName | RowIndex | ColIndex | Status | Timestamp
 *
 * @returns {Sheet} The status sheet
 */
function getStatusSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(STATUS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(STATUS_SHEET_NAME);
    // Hide the sheet - users don't need to see it
    sheet.hideSheet();
    // Add header row
    sheet.getRange(1, 1, 1, 5).setValues([['SheetName', 'RowIndex', 'ColIndex', 'Status', 'LastModified']]);
    sheet.getRange(1, 1, 1, 5).setFontWeight('bold');

    // Protect the sheet from manual edits - only script can modify
    var protection = sheet.protect().setDescription('Status Storage - Auto-managed by script');
    var editors = protection.getEditors();
    editors.forEach(function(editor) {
      protection.removeEditor(editor);
    });
    protection.setWarningOnly(false);
  }
  return sheet;
}

/**
 * Returns a cache key for a given sheet
 */
function getCacheKey(sheetName) {
  return 'status_' + sheetName;
}

/**
 * Clears the status cache for a sheet (call when data might have changed)
 */
function clearStatusCache(sheetName) {
  var key = getCacheKey(sheetName);
  delete statusCache[key];
}

/**
 * Reads status from cell background colors instead of the status sheet.
 * Much faster because it's from the same sheet as data - no separate sheet read.
 * Reads ALL language columns in a SINGLE getBackgrounds() call.
 *
 * @param {Sheet} sheet - The data sheet (not status sheet)
 * @param {number} startRow - 1-based start row
 * @param {number} numRows - Number of rows to read
 * @param {Array} colIndices - Array of column indices to check
 * @returns {Object} statusMap[rowIndex][colIndex] = status
 */
function getStatusFromColors(sheet, startRow, numRows, colIndices) {
  var statusMap = {};

  if (colIndices.length === 0 || numRows <= 0) {
    return statusMap;
  }

  // Find min/max column for efficient batch reading
  var minCol = Math.min.apply(Math, colIndices);
  var maxCol = Math.max.apply(Math, colIndices);
  var numCols = maxCol - minCol + 1;

  // Read ALL backgrounds in ONE API call for the entire column range
  var allColors = sheet.getRange(startRow, minCol, numRows, numCols).getBackgrounds();

  // Build a map of column index to position in the results
  var colPos = {};
  colIndices.forEach(function(colIndex) {
    colPos[colIndex] = colIndex - minCol;
  });

  // Extract statuses from the batch results
  for (var r = 0; r < allColors.length; r++) {
    var rowIndex = startRow + r;
    var rowColors = allColors[r];

    colIndices.forEach(function(colIndex) {
      var pos = colPos[colIndex];
      var hex = String(rowColors[pos] || '').toLowerCase();
      var status = COLOR_TO_STATUS[hex];

      if (status) {
        if (!statusMap[rowIndex]) statusMap[rowIndex] = {};
        statusMap[rowIndex][colIndex] = status;
      }
    });
  }

  return statusMap;
}

/**
 * Loads all status entries for a given sheet into memory.
 * Returns a nested object: statusMap[rowIndex][colIndex] = { status, sheetRow }
 * where sheetRow is the 1-based row index in the status sheet (for O(1) updates).
 */
function loadStatusMap(sheetName) {
  return loadStatusMapPaginated(sheetName, null, null);
}

/**
 * Loads status entries for a specific row range (paginated).
 * Much faster for large sheets - only loads statuses for the rows being fetched.
 *
 * @param {string} sheetName
 * @param {number|null} startRow - 1-based start row (null for all)
 * @param {number|null} endRow - 1-based end row (null for all)
 * @returns {Object} statusMap[rowIndex][colIndex] = { status, sheetRow }
 */
function loadStatusMapPaginated(sheetName, startRow, endRow) {
  var cacheKey = getCacheKey(sheetName);
  // Only use cache if loading all rows (startRow and endRow are null)
  var useCache = !startRow && !endRow;
  if (useCache && statusCache[cacheKey]) {
    return statusCache[cacheKey];
  }

  var statusSheet = getStatusSheet();
  var lastRow = statusSheet.getLastRow();

  var statusMap = {};

  if (lastRow < 2) {
    if (useCache) statusCache[cacheKey] = statusMap;
    return statusMap;
  }

  // Read all data at once - we filter by row range in memory
  // (more efficient than multiple sheet queries)
  var data = statusSheet.getRange(2, 1, lastRow - 1, 4).getValues();

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var sheet = String(row[0] || '');
    if (sheet !== sheetName) continue;

    var rowIdx = parseInt(row[1], 10);

    // Skip if outside requested range
    if (startRow && rowIdx < startRow) continue;
    if (endRow && rowIdx > endRow) continue;

    var colIdx = parseInt(row[2], 10);
    var status = String(row[3] || '').trim();

    if (!status) continue;

    if (!statusMap[rowIdx]) statusMap[rowIdx] = {};
    // Store both status and the sheet row index for fast updates
    statusMap[rowIdx][colIdx] = { status: status, sheetRow: i + 2 };
  }

  if (useCache) statusCache[cacheKey] = statusMap;
  return statusMap;
}

/**
 * Gets the status for a specific cell.
 *
 * @param {string} sheetName
 * @param {number} rowIndex
 * @param {number} colIndex
 * @returns {string} Status value or '' if not found
 */
function getCellStatus(sheetName, rowIndex, colIndex) {
  var statusMap = loadStatusMap(sheetName);
  if (statusMap[rowIndex] && statusMap[rowIndex][colIndex]) {
    return statusMap[rowIndex][colIndex].status;
  }
  return '';
}

/**
 * Sets the status for a specific cell.
 * Uses O(1) update via cached sheet row index - no linear scan needed.
 *
 * @param {string} sheetName
 * @param {number} rowIndex
 * @param {number} colIndex
 * @param {string} status
 */
function setCellStatus(sheetName, rowIndex, colIndex, status) {
  var statusSheet = getStatusSheet();

  // Load cache BEFORE modifying - we need to check if entry exists
  var cacheKey = getCacheKey(sheetName);
  var statusMap = loadStatusMap(sheetName);
  var existingEntry = statusMap[rowIndex] && statusMap[rowIndex][colIndex];

  if (existingEntry && existingEntry.sheetRow) {
    // O(1) update: we know exactly which row to update
    var sheetRow = existingEntry.sheetRow;
    // Batch update both status and timestamp in one call
    statusSheet.getRange(sheetRow, 4, 1, 2).setValues([[status, new Date()]]);
    // Update cache in place (faster than clearing and reloading)
    existingEntry.status = status;
  } else {
    // New entry - append and store the new row index
    var newRow = statusSheet.getLastRow() + 1;
    statusSheet.appendRow([sheetName, rowIndex, colIndex, status, new Date()]);
    // Update cache with new entry
    if (!statusMap[rowIndex]) statusMap[rowIndex] = {};
    statusMap[rowIndex][colIndex] = { status: status, sheetRow: newRow };
  }
}

/**
 * Batch gets statuses for multiple cells at once.
 * More efficient than individual getCellStatus calls.
 *
 * @param {string} sheetName
 * @param {Array} cells Array of {rowIndex, colIndex} objects
 * @returns {Object} Map of "row,col" -> status
 */
function batchGetStatuses(sheetName, cells) {
  var statusMap = loadStatusMap(sheetName);
  var result = {};

  for (var i = 0; i < cells.length; i++) {
    var cell = cells[i];
    var key = cell.rowIndex + ',' + cell.colIndex;
    var status = '';
    if (statusMap[cell.rowIndex] && statusMap[cell.rowIndex][cell.colIndex]) {
      status = statusMap[cell.rowIndex][cell.colIndex].status;
    }
    result[key] = status;
  }

  return result;
}

/**
 * Seeds status values from existing cell background colors.
 * Call this during initial setup to migrate from color-based tracking.
 */
function seedStatusesFromColors() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var statusSheet = getStatusSheet();
  var ui = SpreadsheetApp.getUi();

  var totalSeeded = 0;
  var rowsToAdd = [];

  sheets.forEach(function(sheet) {
    var sheetName = sheet.getName();
    if (isIgnoredSheet(sheetName)) return;
    if (sheetName === STATUS_SHEET_NAME) return;

    var languages = detectLanguageColumns(sheet);
    if (!languages.length) return;

    var lastRow = sheet.getLastRow();
    if (lastRow < 3) return;

    var numRows = lastRow - 2;

    languages.forEach(function(lang) {
      lang.columns.forEach(function(col) {
        var colIndex = col.colIndex;

        // Read background colors
        var translationRange = sheet.getRange(3, colIndex, numRows, 1);
        var bgColors = translationRange.getBackgrounds();

        for (var r = 0; r < bgColors.length; r++) {
          var hex = String(bgColors[r][0] || '').toLowerCase();
          var status = COLOR_TO_STATUS[hex];

          // Only store non-default statuses to keep sheet smaller
          // UNTRANSLATED is default, so we skip it unless you want explicit tracking
          if (status && status !== STATUS.UNTRANSLATED) {
            var rowIndex = r + 3;
            rowsToAdd.push([sheetName, rowIndex, colIndex, status, new Date()]);
            totalSeeded++;
          }
        }
      });
    });
  });

  // Batch write all rows
  if (rowsToAdd.length > 0) {
    var startRow = statusSheet.getLastRow() + 1;
    statusSheet.getRange(startRow, 1, rowsToAdd.length, 5).setValues(rowsToAdd);
  }

  ui.alert(
    'Status Storage Setup',
    'Done! Seeded ' + totalSeeded + ' status entries from cell colors.\n\n' +
    'Run "Refresh Dashboard" to update the progress overview.',
    ui.ButtonSet.OK
  );
}

/**
 * Clears all status data. Use with caution!
 */
function clearAllStatuses() {
  var statusSheet = getStatusSheet();
  var lastRow = statusSheet.getLastRow();
  if (lastRow > 1) {
    statusSheet.deleteRows(2, lastRow - 1);
  }
  statusCache = {};
}
