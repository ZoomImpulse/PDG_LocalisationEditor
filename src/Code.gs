// ─────────────────────────────────────────────────────────────────────────────
// Code.gs — Core logic
// ─────────────────────────────────────────────────────────────────────────────

// ── Menu ─────────────────────────────────────────────────────────────────────

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('PDG Localisation')
    .addItem('Open Editor', 'openTranslatorModal')
    .addSeparator()
    .addItem('Refresh Overview', 'refreshOverviewSheet')
    .addSeparator()
    .addItem('Setup: Seed Statuses', 'insertStatusColumns')
    .addItem('Setup: Initialise Roles', 'initialiseRoles')
    .addToUi();
}

// ── Sidebar (legacy, kept for compatibility) ───────────────────────────────────

function showSidebar() {
  openTranslatorModal();
}

// ── Modal ─────────────────────────────────────────────────────────────────────

function openTranslatorModal() {
  var template = HtmlService.createTemplateFromFile('Modal');
  template.userRole = getCurrentUserRole();
  template.userLastLang = getUserLastLang();
  var html = template.evaluate().setWidth(1200).setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, ' ');  // Empty title, no header
}

// ── Column detection ──────────────────────────────────────────────────────────

/**
 * Scans row 1 and row 2 of a sheet and returns a structured map of all
 * translatable language columns, grouped by language name.
 *
 * Rules applied:
 *  - Columns where row1 === row2 (export formula columns) are skipped entirely.
 *  - All other columns are candidate translation columns; their language name
 *    is extracted from row2 via extractLanguageName().
 *
 * Returns an array of language groups:
 * [
 *   {
 *     languageName: 'English',        // extracted, used as grouping key
 *     columns: [
 *       { colIndex: 2, header: 'Localisation (English)' },
 *       { colIndex: 4, header: 'Name Localisation (English)' },
 *     ]
 *   }, ...
 * ]
 *
 * colIndex is 1-based (Google Sheets column numbering).
 *
 * @param {Sheet} sheet
 * @returns {Array}
 */
/**
 * Finds the 1-based column index of the ID/key column by scanning row 2
 * for a header that contains "ID" (case-insensitive). Falls back to column 1.
 *
 * @param {Sheet} sheet
 * @returns {number}
 */
function detectKeyColumn(sheet) {
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) return 1;
  var row2 = sheet.getRange(2, 1, 1, lastCol).getValues()[0];
  for (var i = 0; i < row2.length; i++) {
    var h = String(row2[i] || '').trim().toUpperCase();
    if (h === 'ID' || h.indexOf('ID') !== -1) return i + 1;
  }
  return 1; // default
}

function detectLanguageColumns(sheet) {
  var lastCol = sheet.getLastColumn();
  if (lastCol < 2) return [];

  var keyCol = detectKeyColumn(sheet);

  // Read both header rows in one batch call
  var headerRows = sheet.getRange(1, 1, 2, lastCol).getValues();
  var row1Values = headerRows[0];
  var row2Values = headerRows[1];

  // Build raw column info first
  var cols = [];
  for (var i = 0; i < lastCol; i++) {
    var r1 = String(row1Values[i] || '').trim();
    var r2 = String(row2Values[i] || '').trim();
    var colIndex = i + 1; // 1-based

    if (colIndex === keyCol) {
      // Key column — always skip
      continue;
    }

    if (!r2) {
      // Empty header — skip
      continue;
    }

    if (r1 === r2) {
      // Export formula column — skip entirely
      continue;
    }

    var langName = extractLanguageName(r2);
    if (!langName) continue;

    cols.push({
      type: 'translation',
      colIndex: colIndex,
      header: r2,
      languageName: langName,
    });
  }

  // Group by language name
  var groups = {};
  var order = [];
  cols.forEach(function(c) {
    if (c.type !== 'translation') return;
    if (!groups[c.languageName]) {
      groups[c.languageName] = { languageName: c.languageName, columns: [] };
      order.push(c.languageName);
    }
    groups[c.languageName].columns.push({
      colIndex: c.colIndex,
      header: c.header,
    });
  });

  return order.map(function(name) { return groups[name]; });
}

// ── getSheetData (legacy, loads all rows) ────────────────────────────────────

/**
 * Returns all row data needed by the sidebar for a given sheet.
 * NOTE: For large sheets, use getSheetDataPaginated instead.
 *
 * @param {string} sheetName
 * @returns {Object} { sheetName, languages, rows: [...] }
 */
function getSheetData(sheetName) {
  return getSheetDataPaginated(sheetName, 3, Infinity);
}

// ── getSheetInfo (fast metadata) ───────────────────────────────────────────

/**
 * Returns sheet metadata quickly without loading all row data.
 * Used for initial modal load to show UI immediately.
 *
 * @param {string} sheetName
 * @returns {Object} { sheetName, languages, totalRows, hasMoreData }
 */
function getSheetInfo(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('Sheet not found: ' + sheetName);

  var languages = detectLanguageColumns(sheet);
  var lastRow = sheet.getLastRow();
  var totalDataRows = Math.max(0, lastRow - 2);

  return {
    sheetName: sheetName,
    languages: languages,
    totalRows: totalDataRows,
    hasMoreData: totalDataRows > 100
  };
}

// ── getSheetDataPaginated (fast batch loading) ───────────────────────────────

/**
 * Returns a batch of row data for paginated loading.
 * Much faster for large sheets - loads only requested rows.
 *
 * @param {string} sheetName
 * @param {number} startRow    1-based row index to start from (typically 3)
 * @param {number} limit       Max rows to return (use Infinity for all)
 * @param {Array} languages    Optional pre-detected languages (skips re-detection)
 * @returns {Object} { sheetName, languages, rows: [...], hasMore, nextStartRow }
 */
function getSheetDataPaginated(sheetName, startRow, limit, languages) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('Sheet not found: ' + sheetName);

  // Use provided languages or detect them
  languages = languages || detectLanguageColumns(sheet);

  var lastRow = sheet.getLastRow();
  if (lastRow < 3 || startRow > lastRow) {
    return { sheetName: sheetName, languages: languages, rows: [], hasMore: false };
  }

  // Calculate range to fetch
  var startDataRow = Math.max(3, startRow);
  var maxRows = lastRow - startDataRow + 1;
  var numRows = (limit === Infinity) ? maxRows : Math.min(limit, maxRows);

  // Collect all column indices for status reading
  var colIndices = [];
  languages.forEach(function(lang) {
    lang.columns.forEach(function(c) {
      colIndices.push(c.colIndex);
    });
  });

  var keyCol = detectKeyColumn(sheet);

  // Find the column range we actually need (key col + all language columns)
  var maxLangCol = Math.max.apply(Math, colIndices.concat([keyCol]));
  var numColsToRead = maxLangCol;

  // Load statuses from cell colors (fast - same sheet, no status sheet read)
  var statusMap = getStatusFromColors(sheet, startDataRow, numRows, colIndices);

  // Fetch only the rows and columns we need (not getLastColumn which could be much larger)
  var dataRange = sheet.getRange(startDataRow, 1, numRows, numColsToRead);
  var values = dataRange.getValues();

  var rows = [];
  for (var r = 0; r < values.length; r++) {
    var rowIndex = startDataRow + r;
    var key = String(values[r][keyCol - 1] || '').trim();
    if (!key) continue;

    var cells = {};
    languages.forEach(function(lang) {
      lang.columns.forEach(function(c) {
        var val = String(values[r][c.colIndex - 1] || '');
        var status = '';
        if (statusMap[rowIndex] && statusMap[rowIndex][c.colIndex]) {
          status = statusMap[rowIndex][c.colIndex];
        }
        cells[String(c.colIndex)] = { value: val, status: status };
      });
    });

    rows.push({ rowIndex: rowIndex, key: key, cells: cells });
  }

  var nextStartRow = startDataRow + numRows;
  var hasMore = nextStartRow <= lastRow;

  return {
    sheetName: sheetName,
    languages: languages,
    keyColIndex: keyCol,
    rows: rows,
    hasMore: hasMore,
    nextStartRow: hasMore ? nextStartRow : null
  };
}

// ── getFullSheetData (load entire sheet at once) ────────────────────────────────

/**
 * Returns all row data for a sheet in a single call.
 * Used when full sheet loading is preferred over pagination.
 *
 * @param {string} sheetName
 * @returns {Object} { sheetName, languages, rows: [...] }
 */
function getFullSheetData(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('Sheet not found: ' + sheetName);

  var languages = detectLanguageColumns(sheet);
  var lastRow = sheet.getLastRow();

  if (lastRow < 3) {
    return { sheetName: sheetName, languages: languages, rows: [] };
  }

  var numRows = lastRow - 2; // Subtract header rows

  // Collect all column indices for status reading
  var colIndices = [];
  languages.forEach(function(lang) {
    lang.columns.forEach(function(c) {
      colIndices.push(c.colIndex);
    });
  });

  var keyCol = detectKeyColumn(sheet);

  // Find the column range we actually need
  var maxLangCol = Math.max.apply(Math, colIndices.concat([keyCol]));
  var numColsToRead = maxLangCol;

  // Load statuses from cell colors
  var statusMap = getStatusFromColors(sheet, 3, numRows, colIndices);

  // Fetch all data
  var dataRange = sheet.getRange(3, 1, numRows, numColsToRead);
  var values = dataRange.getValues();

  var rows = [];
  for (var r = 0; r < values.length; r++) {
    var rowIndex = r + 3;
    var key = String(values[r][keyCol - 1] || '').trim();
    if (!key) continue;

    var cells = {};
    languages.forEach(function(lang) {
      lang.columns.forEach(function(c) {
        var val = String(values[r][c.colIndex - 1] || '');
        var status = '';
        if (statusMap[rowIndex] && statusMap[rowIndex][c.colIndex]) {
          status = statusMap[rowIndex][c.colIndex];
        }
        cells[String(c.colIndex)] = { value: val, status: status };
      });
    });

    rows.push({ rowIndex: rowIndex, key: key, cells: cells });
  }

  return {
    sheetName: sheetName,
    languages: languages,
    keyColIndex: keyCol,
    rows: rows
  };
}

// ── searchAllSheets (global search across all sheets) ─────────────────────────

/**
 * Searches across all non-ignored sheets for keys or translations matching the query.
 *
 * @param {string} query - Search string (case-insensitive)
 * @param {number} maxResults - Maximum results to return (default 100)
 * @returns {Array} [{ sheetName, rowIndex, key, preview, matchType, language, columnIndex, fieldName }]
 */
function searchAllSheets(query, maxResults) {
  maxResults = maxResults || 100;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var results = [];
  var queryLower = String(query).toLowerCase();

  for (var s = 0; s < sheets.length && results.length < maxResults; s++) {
    var sheet = sheets[s];
    var sheetName = sheet.getName();

    if (isIgnoredSheet(sheetName)) continue;

    var languages = detectLanguageColumns(sheet);
    if (!languages.length) continue;

    var lastRow = sheet.getLastRow();
    if (lastRow < 3) continue;

    var numRows = lastRow - 2;

    // Build column-to-language map and collect indices
    var colToLang = {};
    var colToField = {};
    var colIndices = [];

    languages.forEach(function(lang) {
      lang.columns.forEach(function(c) {
        colIndices.push(c.colIndex);
        colToLang[c.colIndex] = lang.languageName;
        colToField[c.colIndex] = c.fieldName || 'Text';
      });
    });

    if (colIndices.length === 0) continue;

    var maxLangCol = Math.max.apply(Math, colIndices);

    var keyCol = detectKeyColumn(sheet);
    var maxDataCol = Math.max(maxLangCol, keyCol);

    // Read key column and all language columns
    var dataRange = sheet.getRange(3, 1, numRows, maxDataCol);
    var values = dataRange.getValues();

    for (var r = 0; r < values.length && results.length < maxResults; r++) {
      var rowIndex = r + 3;
      var key = String(values[r][keyCol - 1] || '').trim();
      if (!key) continue;

      // Check if key matches
      var keyMatches = key.toLowerCase().includes(queryLower);
      var matchColIndex = null;
      var matchLanguage = null;
      var matchField = null;

      // Check all language values
      var valueMatches = false;
      var preview = '';

      for (var i = 0; i < colIndices.length; i++) {
        var colIdx = colIndices[i];
        var val = String(values[r][colIdx - 1] || '');

        // Use first non-empty value as preview (prefer non-English)
        if (!preview && val && colToLang[colIdx] !== 'English') {
          preview = val.substring(0, 100);
        }

        if (val.toLowerCase().includes(queryLower)) {
          valueMatches = true;
          matchColIndex = colIdx;
          matchLanguage = colToLang[colIdx];
          matchField = colToField[colIdx];
          preview = val.substring(0, 100);
          break;
        }
      }

      // If only key matched, pick first non-English column as default
      if (keyMatches && !matchColIndex) {
        for (var i = 0; i < colIndices.length; i++) {
          var colIdx = colIndices[i];
          if (colToLang[colIdx] !== 'English') {
            matchColIndex = colIdx;
            matchLanguage = colToLang[colIdx];
            matchField = colToField[colIdx];
            // Use preview from this column
            var val = String(values[r][colIdx - 1] || '');
            if (val) preview = val.substring(0, 100);
            break;
          }
        }
      }

      if (keyMatches || valueMatches) {
        results.push({
          sheetName: sheetName,
          rowIndex: rowIndex,
          key: key,
          preview: preview,
          matchType: keyMatches ? 'key' : 'value',
          language: matchLanguage,
          columnIndex: matchColIndex,
          fieldName: matchField
        });
      }
    }
  }

  return results;
}

// ── saveTranslation ───────────────────────────────────────────────────────────

/**
 * Writes a translation value and status back to the sheet.
 *
 * @param {string} sheetName
 * @param {number} rowIndex   1-based row number
 * @param {number} colIndex   1-based column number of the translation cell
 * @param {number} statusCol  Unused (legacy param for compatibility)
 * @param {string} value      Translation text
 * @param {string} status     One of STATUS values
 */
function saveTranslation(sheetName, rowIndex, colIndex, statusCol, value, status, cascadeColIndices) {
  // Server-side role enforcement
  var callerRole = getCurrentUserRole();
  if (callerRole === ROLE.NONE) throw new Error('Access denied: you do not have a translator role.');
  if (callerRole === ROLE.TRANSLATOR && status === 'REVIEWED') {
    throw new Error('Access denied: only Reviewers and Admins can mark entries as Reviewed.');
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('Sheet not found: ' + sheetName);

  var color = STATUS_COLOR[status] || STATUS_COLOR.UNTRANSLATED;
  sheet.getRange(rowIndex, colIndex).setValue(value).setBackground(color);

  // Cascade DISPUTED to other language columns supplied by the client
  if (status === STATUS.DISPUTED && cascadeColIndices && cascadeColIndices.length > 0) {
    var a1Notations = cascadeColIndices.map(function(ci) {
      return sheet.getRange(rowIndex, ci).getA1Notation();
    });
    sheet.getRangeList(a1Notations).setBackground(STATUS_COLOR.DISPUTED);
  }
}

// ── getProgress ───────────────────────────────────────────────────────────────

/**
 * Returns translation progress for all non-ignored sheets.
 *
 * @returns {Array} [
 *   {
 *     sheetName,
 *     languages: [
 *       { languageName, translated, total }
 *     ]
 *   }
 * ]
 */
function getProgress() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var result = [];

  sheets.forEach(function(sheet) {
    var name = sheet.getName();
    if (isIgnoredSheet(name)) return;

    var languages = detectLanguageColumns(sheet);
    if (!languages.length) return;

    var lastRow = sheet.getLastRow();
    if (lastRow < 3) {
      result.push({
        sheetName: name,
        languages: languages.map(function(l) {
          return { languageName: l.languageName, translated: 0, total: 0 };
        })
      });
      return;
    }

    var numRows = lastRow - 2;
    var keyCol = detectKeyColumn(sheet);

    // Collect all column indices we need to check
    var colIndices = [];
    var colToLang = {};
    languages.forEach(function(lang) {
      lang.columns.forEach(function(c) {
        colIndices.push(c.colIndex);
        colToLang[c.colIndex] = lang.languageName;
      });
    });

    // Read keys and all language backgrounds in just 2 batch calls
    var maxLangCol = Math.max.apply(Math, colIndices);
    var numColsToRead = maxLangCol;

    // Single batch read for all backgrounds
    var allBackgrounds = sheet.getRange(3, 1, numRows, numColsToRead).getBackgrounds();
    var allKeys = sheet.getRange(3, keyCol, numRows, 1).getValues();

    // Build column index to array index map
    var colToArrayIdx = {};
    colIndices.forEach(function(ci) {
      colToArrayIdx[ci] = ci - 1; // 1-based to 0-based
    });

    // Count stats per language
    var langStats = {};
    languages.forEach(function(lang) {
      langStats[lang.languageName] = { translated: 0, total: 0 };
    });

    for (var r = 0; r < numRows; r++) {
      var key = String(allKeys[r][0] || '').trim();
      if (!key) continue;

      var rowBackgrounds = allBackgrounds[r];

      colIndices.forEach(function(colIndex) {
        var langName = colToLang[colIndex];
        var hex = String(rowBackgrounds[colToArrayIdx[colIndex]] || '').toLowerCase();
        var status = COLOR_TO_STATUS[hex] || '';

        // Skip cells with no color (white/default) or explicitly static
        if (!status || status === STATUS.STATIC) return;

        langStats[langName].total++;
        if (status === STATUS.TRANSLATED || status === STATUS.REVIEWED) {
          langStats[langName].translated++;
        }
      });
    }

    var langArray = Object.keys(langStats).map(function(name) {
      return {
        languageName: name,
        translated: langStats[name].translated,
        total: langStats[name].total,
      };
    });

    result.push({ sheetName: name, languages: langArray });
  });

  return result;
}

// ── getProgressForLanguage ───────────────────────────────────────────────────

/**
 * Returns translation progress for a specific language across all non-ignored sheets.
 * Much faster than getProgress() when only one language is needed.
 *
 * @param {string} targetLanguage - The language name to get progress for
 * @returns {Array} [
 *   {
 *     sheetName,
 *     languages: [
 *       { languageName, translated, total }
 *     ]
 *   }
 * ]
 */
function getProgressForLanguage(targetLanguage) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var result = [];

  sheets.forEach(function(sheet) {
    var name = sheet.getName();
    if (isIgnoredSheet(name)) return;

    var languages = detectLanguageColumns(sheet);
    if (!languages.length) return;

    // Find the target language
    var targetLang = languages.find(function(l) {
      return l.languageName === targetLanguage;
    });

    if (!targetLang) {
      // Language not in this sheet
      result.push({
        sheetName: name,
        languages: []
      });
      return;
    }

    var lastRow = sheet.getLastRow();
    if (lastRow < 3) {
      result.push({
        sheetName: name,
        languages: [{ languageName: targetLanguage, translated: 0, total: 0 }]
      });
      return;
    }

    var numRows = lastRow - 2;
    var keyCol = detectKeyColumn(sheet);

    // Collect only the columns for this specific language
    var colIndices = targetLang.columns.map(function(c) { return c.colIndex; });

    // Read all relevant data in 2 batch calls
    var maxLangCol = Math.max.apply(Math, colIndices);
    var numColsToRead = maxLangCol;

    var allBackgrounds = sheet.getRange(3, 1, numRows, numColsToRead).getBackgrounds();
    var allKeys = sheet.getRange(3, keyCol, numRows, 1).getValues();

    // Build column index to array index map
    var colToArrayIdx = {};
    colIndices.forEach(function(ci) {
      colToArrayIdx[ci] = ci - 1;
    });

    var translated = 0;
    var total = 0;

    for (var r = 0; r < numRows; r++) {
      var key = String(allKeys[r][0] || '').trim();
      if (!key) continue;

      var rowBackgrounds = allBackgrounds[r];

      colIndices.forEach(function(colIndex) {
        var hex = String(rowBackgrounds[colToArrayIdx[colIndex]] || '').toLowerCase();
        var status = COLOR_TO_STATUS[hex] || '';

        // Skip cells with no color (white/default) or explicitly static
        if (!status || status === STATUS.STATIC) return;

        total++;
        if (status === STATUS.TRANSLATED || status === STATUS.REVIEWED) {
          translated++;
        }
      });
    }

    result.push({
      sheetName: name,
      languages: [{ languageName: targetLanguage, translated: translated, total: total }]
    });
  });

  return result;
}

// ── getSheetNames ─────────────────────────────────────────────────────────────

/**
 * Returns all non-ignored sheet names for the sheet selector in the sidebar.
 *
 * @returns {string[]}
 */
function getSheetNames() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheets()
    .filter(function(s) {
      if (isIgnoredSheet(s.getName())) return false;
      // Exclude sheets with a non-warning-only sheet-level protection
      var protections = s.getProtections(SpreadsheetApp.ProtectionType.SHEET);
      return !protections.some(function(p) { return !p.isWarningOnly(); });
    })
    .map(function(s) { return s.getName(); });
}

// ── insertStatusColumns (setup) ───────────────────────────────────────────────

/**
 * One-time setup function. Creates a separate status tracking sheet and
 * seeds it with status values based on existing cell background colors.
 *
 * Status is now stored in a hidden '📋 Status' sheet instead of inserted
 * columns — much less disruptive to the original data structure.
 *
 * Safe to re-run — will not duplicate existing status entries.
 */
function insertStatusColumns() {
  // Delegate to StatusStorage.gs function
  seedStatusesFromColors();
}

// ── protectKeyColumn ──────────────────────────────────────────────────────────

/**
 * Adds a protection on column A of every non-ignored sheet so that
 * non-owner editors cannot accidentally modify translation keys.
 * Existing protections on column A are left untouched.
 */
function protectKeyColumn() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var ui = SpreadsheetApp.getUi();
  var count = 0;

  sheets.forEach(function(sheet) {
    var name = sheet.getName();
    if (isIgnoredSheet(name)) return;

    // Check if column A is already protected
    var existing = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    var alreadyProtected = existing.some(function(p) {
      var range = p.getRange();
      return range.getColumn() === 1 && range.getNumColumns() === 1;
    });

    if (alreadyProtected) return;

    var protection = sheet.getRange(1, 1, sheet.getMaxRows(), 1)
      .protect()
      .setDescription('PDG Localisation: Key column (do not edit)');

    // Only the spreadsheet owner can edit; show a warning to others
    protection.setWarningOnly(true);
    count++;
  });

  ui.alert(
    'PDG Localisation',
    'Protected column A on ' + count + ' sheet(s).',
    ui.ButtonSet.OK
  );
}
