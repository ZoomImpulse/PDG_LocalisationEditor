// ─────────────────────────────────────────────────────────────────────────────
// Overview.gs — Auto-generated translation overview sheet
// Called at the end of refreshProgressCache() — updates every 10 minutes.
// ─────────────────────────────────────────────────────────────────────────────

var OVERVIEW_SHEET_NAME = '📊 Overview';

// ── Colour palette ────────────────────────────────────────────────────────────

var OV_COLORS = {
  // ── Page chrome ────────────────────────────────────────────────────────────
  TITLE_BG:      '#f8f9fa',  // near-white page bg
  TITLE_FG:      '#111827',  // near-black heading text
  SECTION_BG:    '#f1f3f4',  // subtle grey section divider
  SECTION_FG:    '#374151',  // dark-grey section label
  SUBHEADER_BG:  '#e8eaed',  // column header row
  SUBHEADER_FG:  '#1f2937',  // column header text
  META_BG:       '#f8f9fa',
  META_FG:       '#6b7280',  // muted timestamp

  // ── Heat-map (muted, readable) ─────────────────────────────────────────────
  PCT_100:       '#16a34a',  // green-600  – 100%
  PCT_75:        '#4ade80',  // green-400  – ≥75%
  PCT_50:        '#fde68a',  // amber-200  – ≥50%
  PCT_25:        '#fca5a5',  // red-300    – ≥25%
  PCT_0:         '#ef4444',  // red-500    – <25%
  PCT_NONE:      '#e5e7eb',  // grey-200   – no data
  PCT_FG_DARK:   '#111827',
  PCT_FG_LIGHT:  '#ffffff',

  // ── Legend (matching Config.gs STATUS_COLOR exactly) ──────────────────────
  LEG_STATIC:       '#d9d9d9',
  LEG_UNTRANSLATED: '#e06666',
  LEG_TRANSLATED:   '#ffff00',
  LEG_REVIEWED:     '#00ff00',
  LEG_DISPUTED:     '#ff0000',

  // ── Neutrals ───────────────────────────────────────────────────────────────
  WHITE:         '#ffffff',
  LIGHT_GREY:    '#f9fafb',
  ROW_ALT:       '#f3f4f6',  // alternating row tint
  BORDER:        '#e5e7eb',
  TEXT_PRIMARY:  '#111827',
  TEXT_MUTED:    '#6b7280',
};

// Fixed language display order
var OV_LANG_ORDER = ['English', 'German', 'French', 'Russian', 'Polish', 'Spanish', 'Italian', 'Braz_Por'];

// ── Main entry point ──────────────────────────────────────────────────────────

/**
 * Builds / refreshes the 📊 Overview sheet from the cached progress data.
 * Called automatically at the end of refreshProgressCache().
 */
function refreshOverviewSheet() {
  var cachedProgress = getCachedProgress();

  // ── Collect cross-spreadsheet progress ──────────────────────────────────────
  // Build: globalData[ssName][language] = { total, reviewed, translated, disputed, untranslated }
  var globalData = {};
  var globalLangSet = {};

  // Include this spreadsheet itself
  var thisSsName = SpreadsheetApp.getActiveSpreadsheet().getName();
  globalData[thisSsName] = {};
  cachedProgress.forEach(function(entry) {
    if (!globalData[thisSsName][entry.language]) {
      globalData[thisSsName][entry.language] = { total: 0, translated: 0, reviewed: 0, disputed: 0, untranslated: 0 };
    }
    var t = globalData[thisSsName][entry.language];
    t.total       += entry.total;
    t.translated  += entry.translated;
    t.reviewed    += entry.reviewed;
    t.disputed    += entry.disputed;
    t.untranslated += entry.untranslated;
    globalLangSet[entry.language] = true;
  });

  // Include linked spreadsheets
  if (LINKED_SPREADSHEETS && LINKED_SPREADSHEETS.length > 0) {
    LINKED_SPREADSHEETS.forEach(function(ssEntry) {
      var linkedProgress = readLinkedProgress(ssEntry);
      globalData[ssEntry.name] = {};
      linkedProgress.forEach(function(entry) {
        if (!globalData[ssEntry.name][entry.language]) {
          globalData[ssEntry.name][entry.language] = { total: 0, translated: 0, reviewed: 0, disputed: 0, untranslated: 0 };
        }
        var t = globalData[ssEntry.name][entry.language];
        t.total       += entry.total;
        t.translated  += entry.translated;
        t.reviewed    += entry.reviewed;
        t.disputed    += entry.disputed;
        t.untranslated += entry.untranslated;
        globalLangSet[entry.language] = true;
      });
    });
  }

  var globalSsNames = Object.keys(globalData);
  var globalLangs = OV_LANG_ORDER.filter(function(l) { return globalLangSet[l]; });
  Object.keys(globalLangSet).forEach(function(l) {
    if (globalLangs.indexOf(l) === -1) globalLangs.push(l);
  });
  var hasGlobal = (typeof MAIN_SPREADSHEET_ID !== 'undefined'
                   && SpreadsheetApp.getActiveSpreadsheet().getId() === MAIN_SPREADSHEET_ID)
                  && globalSsNames.length > 1;

  var sheet = getOrCreateOverviewSheet();

  // Unprotect for writing
  var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  protections.forEach(function(p) { p.remove(); });

  sheet.clearContents();
  sheet.clearFormats();

  // Provision generous rows/cols upfront; excess will be deleted after writing.
  var neededRows = 200;
  var neededCols = Math.max(10, 2 + OV_LANG_ORDER.length); // recalculated below once allLangs is known
  if (sheet.getMaxRows() < neededRows) sheet.insertRowsAfter(sheet.getMaxRows(), neededRows - sheet.getMaxRows());
  if (sheet.getMaxColumns() < neededCols) sheet.insertColumnsAfter(sheet.getMaxColumns(), neededCols - sheet.getMaxColumns());

  // ── Build data structures ──────────────────────────────────────────────────

  // Collect all languages present, in fixed order
  var langSet = {};
  cachedProgress.forEach(function(entry) { langSet[entry.language] = true; });
  var allLangs = OV_LANG_ORDER.filter(function(l) { return langSet[l]; });
  // Append any unknown languages at the end
  Object.keys(langSet).forEach(function(l) {
    if (allLangs.indexOf(l) === -1) allLangs.push(l);
  });

  // Build lookup: sheetName → language → stats
  var lookup = {};
  cachedProgress.forEach(function(entry) {
    if (!lookup[entry.sheetName]) lookup[entry.sheetName] = {};
    lookup[entry.sheetName][entry.language] = entry;
  });

  // Recalculate neededCols now that allLangs is known
  neededCols = Math.max(9, 1 + allLangs.length);
  if (sheet.getMaxColumns() < neededCols) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), neededCols - sheet.getMaxColumns());
  }

  // Collect sheet names (preserve order from cache)
  var sheetOrder = [];
  var seenSheets = {};
  cachedProgress.forEach(function(entry) {
    if (!seenSheets[entry.sheetName]) {
      sheetOrder.push(entry.sheetName);
      seenSheets[entry.sheetName] = true;
    }
  });

  // ── Section 1: Title & meta ────────────────────────────────────────────────

  var cursor = 1; // current write row (1-based)

  // Row 1: Title
  var titleRange = sheet.getRange(cursor, 1, 1, neededCols);
  titleRange.merge()
    .setValue(SpreadsheetApp.getActiveSpreadsheet().getName() + '  ·  Overview')
    .setBackground(OV_COLORS.WHITE)
    .setFontColor(OV_COLORS.TEXT_PRIMARY)
    .setFontSize(20)
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(cursor, 52);
  cursor++;

  // Row 2: Timestamp
  var now = new Date();
  var tsRange = sheet.getRange(cursor, 1, 1, neededCols);
  tsRange.merge()
    .setValue('Auto-refreshed every 10 minutes  ·  Last update: ' + now.toLocaleString())
    .setBackground(OV_COLORS.WHITE)
    .setFontColor(OV_COLORS.TEXT_MUTED)
    .setFontSize(11)
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(cursor, 24);
  cursor++;

  // Thin separator row
  sheet.getRange(cursor, 1, 1, neededCols).merge().setBackground(OV_COLORS.BORDER);
  sheet.setRowHeight(cursor, 2);
  cursor++;

  // Blank spacer
  sheet.getRange(cursor, 1, 1, neededCols).merge().setBackground(OV_COLORS.WHITE);
  sheet.setRowHeight(cursor, 12);
  cursor++;

  // ── Section: Global Overview (only when linked spreadsheets are configured) ──

  if (hasGlobal) {
    sheet.getRange(cursor, 1, 1, neededCols).merge()
      .setValue('  GLOBAL OVERVIEW  (reviewed %  ·  all spreadsheets)')
      .setBackground(OV_COLORS.SECTION_BG)
      .setFontColor(OV_COLORS.SECTION_FG)
      .setFontSize(10)
      .setFontWeight('bold')
      .setHorizontalAlignment('left')
      .setVerticalAlignment('middle');
    sheet.setRowHeights(cursor, 1, 28);
    cursor++;

    // Column headers: Spreadsheet | lang1 | lang2 | ...
    sheet.getRange(cursor, 1)
      .setValue('Spreadsheet')
      .setBackground(OV_COLORS.SUBHEADER_BG)
      .setFontColor(OV_COLORS.SUBHEADER_FG)
      .setFontSize(10).setFontWeight('bold')
      .setHorizontalAlignment('left').setVerticalAlignment('middle');
    globalLangs.forEach(function(lang, i) {
      sheet.getRange(cursor, i + 2)
        .setValue(lang)
        .setBackground(OV_COLORS.SUBHEADER_BG)
        .setFontColor(OV_COLORS.SUBHEADER_FG)
        .setFontSize(10).setFontWeight('bold')
        .setHorizontalAlignment('center').setVerticalAlignment('middle');
    });
    sheet.setRowHeights(cursor, 1, 30);
    cursor++;

    // One row per spreadsheet
    globalSsNames.forEach(function(ssName, rowIdx) {
      var rowBg = (rowIdx % 2 === 0) ? OV_COLORS.WHITE : OV_COLORS.LIGHT_GREY;
      var isThisSs = (ssName === thisSsName);

      sheet.getRange(cursor, 1)
        .setValue((isThisSs ? 'Main' : ssName) + (isThisSs ? '  ★' : ''))
        .setBackground(rowBg)
        .setFontSize(11).setFontWeight('bold')
        .setFontColor(OV_COLORS.TEXT_PRIMARY)
        .setHorizontalAlignment('left').setVerticalAlignment('middle');

      globalLangs.forEach(function(lang, i) {
        var cell = sheet.getRange(cursor, i + 2);
        var s = globalData[ssName] && globalData[ssName][lang];
        if (!s || s.total === 0) {
          cell.setValue('—').setBackground(OV_COLORS.PCT_NONE)
            .setFontColor(OV_COLORS.TEXT_MUTED).setFontSize(11)
            .setHorizontalAlignment('center').setVerticalAlignment('middle');
        } else {
          var pct = Math.round((s.reviewed / s.total) * 100);
          cell.setValue(pct + '%')
            .setBackground(pctToColor(pct)).setFontColor(pctToFontColor(pct))
            .setFontSize(11).setFontWeight('bold')
            .setHorizontalAlignment('center').setVerticalAlignment('middle')
            .setNote('Reviewed: ' + s.reviewed + ' / ' + s.total + '\nTranslated: ' + s.translated + '\nDisputed: ' + s.disputed + '\nUntranslated: ' + s.untranslated);
        }
      });

      sheet.setRowHeights(cursor, 1, 28);
      cursor++;
    });

    // Spacer + separator
    sheet.getRange(cursor, 1, 1, neededCols).merge().setBackground(OV_COLORS.WHITE);
    sheet.setRowHeight(cursor, 12); cursor++;
    sheet.getRange(cursor, 1, 1, neededCols).merge().setBackground(OV_COLORS.BORDER);
    sheet.setRowHeight(cursor, 2); cursor++;
    sheet.getRange(cursor, 1, 1, neededCols).merge().setBackground(OV_COLORS.WHITE);
    sheet.setRowHeight(cursor, 12); cursor++;
  }

  // Legend section header
  var legHeaderRange = sheet.getRange(cursor, 1, 1, neededCols);
  legHeaderRange.merge()
    .setValue('  STATUS LEGEND')
    .setBackground(OV_COLORS.SECTION_BG)
    .setFontColor(OV_COLORS.SECTION_FG)
    .setFontSize(10)
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');
  sheet.setRowHeights(cursor, 1, 28);
  cursor++;

  // Rows 5–6: Legend — coloured label + description beneath
  var legendItems = [
    {
      label: 'Reviewed',
      bg: OV_COLORS.LEG_REVIEWED,
      fg: '#000000',
      desc: 'Cell has been proofread for grammar and content.'
    },
    {
      label: 'Translated',
      bg: OV_COLORS.LEG_TRANSLATED,
      fg: '#000000',
      desc: 'Cell needs to be proofread for grammar and content. If the English cell is this colour, do not translate yet.'
    },
    {
      label: 'Untranslated',
      bg: OV_COLORS.LEG_UNTRANSLATED,
      fg: '#ffffff',
      desc: 'Cell needs a translation (or description is WIP).'
    },
    {
      label: 'Disputed',
      bg: OV_COLORS.LEG_DISPUTED,
      fg: '#ffffff',
      desc: 'Cell contains false or missing information and needs to be rewritten/corrected. If the English cell is this colour, do not translate yet.'
    },
    {
      label: 'Static',
      bg: OV_COLORS.LEG_STATIC,
      fg: '#000000',
      desc: 'Cell content is static. Do not change it.'
    },
    {
      label: '%C/O%',
      bg: '#b4a7d6',
      fg: '#111827',
      desc: 'Cell is currently empty. Content is optional — used for empire-specific localisation.'
    },
  ];

  legendItems.forEach(function(item, i) {
    var col = i + 1; // cols 1,2,3,4,5 — one per column, no skipping
    // Label row (coloured background)
    var labelCell = sheet.getRange(cursor, col);
    labelCell
      .setValue(item.label)
      .setBackground(item.bg)
      .setFontColor(item.fg)
      .setFontSize(10)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    // Add a visible border for white/light bg cells so they don't disappear
    if (item.bg === OV_COLORS.WHITE || item.bg === OV_COLORS.LIGHT_GREY) {
      labelCell.setBorder(true, true, true, true, false, false, OV_COLORS.BORDER, SpreadsheetApp.BorderStyle.SOLID);
    }
    // Description row (light background)
    sheet.getRange(cursor + 1, col)
      .setValue(item.desc)
      .setBackground('#f8f9fa')
      .setFontColor('#3c4043')
      .setFontSize(9)
      .setFontWeight('normal')
      .setHorizontalAlignment('left')
      .setVerticalAlignment('top')
      .setWrap(true);
  });

  sheet.setRowHeights(cursor,     1, 28);
  sheet.setRowHeights(cursor + 1, 1, 56);
  cursor += 2;

  // Blank spacer
  sheet.getRange(cursor, 1, 1, neededCols).merge().setBackground(OV_COLORS.WHITE);
  sheet.setRowHeight(cursor, 12);
  cursor++;

  // Thin separator
  sheet.getRange(cursor, 1, 1, neededCols).merge().setBackground(OV_COLORS.BORDER);
  sheet.setRowHeight(cursor, 2);
  cursor++;

  // Blank spacer
  sheet.getRange(cursor, 1, 1, neededCols).merge().setBackground(OV_COLORS.WHITE);
  sheet.setRowHeight(cursor, 12);
  cursor++;

  // ── Section 2: Overall progress per language ───────────────────────────────

  var secHeaderRange = sheet.getRange(cursor, 1, 1, neededCols);
  secHeaderRange.merge()
    .setValue('  OVERALL PROGRESS PER LANGUAGE')
    .setBackground(OV_COLORS.SECTION_BG)
    .setFontColor(OV_COLORS.SECTION_FG)
    .setFontSize(10)
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');
  sheet.setRowHeights(cursor, 1, 28);
  cursor++;

  // Column headers for summary table
  var summaryHeaders = ['Language', 'Reviewed %', 'Translated %', 'Total Keys', 'Reviewed', 'Translated', 'Disputed', 'Untranslated', 'Progress Bar'];
  summaryHeaders.forEach(function(h, i) {
    sheet.getRange(cursor, i + 1)
      .setValue(h)
      .setBackground(OV_COLORS.SUBHEADER_BG)
      .setFontColor(OV_COLORS.SUBHEADER_FG)
      .setFontSize(10)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
  });
  sheet.setRowHeights(cursor, 1, 30);
  cursor++;

  // Aggregate totals per language across all sheets
  var langTotals = {};
  allLangs.forEach(function(lang) {
    langTotals[lang] = { total: 0, translated: 0, reviewed: 0, disputed: 0, untranslated: 0 };
  });
  cachedProgress.forEach(function(entry) {
    if (!langTotals[entry.language]) return;
    langTotals[entry.language].total       += entry.total;
    langTotals[entry.language].translated  += entry.translated;
    langTotals[entry.language].reviewed    += entry.reviewed;
    langTotals[entry.language].disputed    += entry.disputed;
    langTotals[entry.language].untranslated += entry.untranslated;
  });

  // Summary rows — store row numbers for sparkline data references
  var summaryDataStartRow = cursor;
  allLangs.forEach(function(lang) {
    var s = langTotals[lang];
    var revPct  = s.total > 0 ? Math.round((s.reviewed  / s.total) * 100) : 0;
    var traPct  = s.total > 0 ? Math.round((s.translated / s.total) * 100) : 0;

    var rowBg = (cursor % 2 === 0) ? OV_COLORS.WHITE : OV_COLORS.LIGHT_GREY;
    var rowData = [lang, revPct + '%', traPct + '%', s.total, s.reviewed, s.translated, s.disputed, s.untranslated];
    rowData.forEach(function(val, i) {
      var cell = sheet.getRange(cursor, i + 1);
      cell.setValue(val).setFontSize(11).setHorizontalAlignment('center')
        .setVerticalAlignment('middle').setBackground(rowBg).setFontColor(OV_COLORS.TEXT_PRIMARY);
      if (i === 0) { cell.setFontWeight('bold').setHorizontalAlignment('left'); }
    });

    // Color the Reviewed% cell (col 2)
    sheet.getRange(cursor, 2)
      .setBackground(pctToColor(revPct)).setFontColor(pctToFontColor(revPct)).setFontWeight('bold').setFontSize(11);

    // Progress bar sparkline in col 9
    var reviewedVal  = s.reviewed;
    var remainingVal = Math.max(0, s.total - s.reviewed);
    if (s.total > 0) {
      sheet.getRange(cursor, 9).setBackground(rowBg).setFormula(
        '=SPARKLINE({' + reviewedVal + ',' + remainingVal + '},{"charttype","bar";"max",' + s.total + ';"color1","' + OV_COLORS.PCT_100 + '";"color2","' + OV_COLORS.PCT_NONE + '"})'
      );
    } else {
      sheet.getRange(cursor, 9).setValue('—').setFontColor(OV_COLORS.TEXT_MUTED).setBackground(rowBg);
    }

    sheet.setRowHeights(cursor, 1, 30);
    cursor++;
  });

  // Spacer + separator before heat-map
  sheet.getRange(cursor, 1, 1, neededCols).merge().setBackground(OV_COLORS.WHITE);
  sheet.setRowHeight(cursor, 12); cursor++;
  sheet.getRange(cursor, 1, 1, neededCols).merge().setBackground(OV_COLORS.BORDER);
  sheet.setRowHeight(cursor, 2); cursor++;
  sheet.getRange(cursor, 1, 1, neededCols).merge().setBackground(OV_COLORS.WHITE);
  sheet.setRowHeight(cursor, 12); cursor++;

  // ── Section 3: Heat-map — Sheet × Language ────────────────────────────────

  var hmHeaderRange = sheet.getRange(cursor, 1, 1, neededCols);
  hmHeaderRange.merge()
    .setValue('  PROGRESS BY SHEET  (reviewed %)')
    .setBackground(OV_COLORS.SECTION_BG)
    .setFontColor(OV_COLORS.SECTION_FG)
    .setFontSize(10)
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');
  sheet.setRowHeights(cursor, 1, 28);
  cursor++;

  // Column headers: Sheet | lang1 | lang2 | ...
  sheet.getRange(cursor, 1)
    .setValue('Sheet')
    .setBackground(OV_COLORS.SUBHEADER_BG)
    .setFontColor(OV_COLORS.SUBHEADER_FG)
    .setFontSize(10)
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');

  allLangs.forEach(function(lang, i) {
    sheet.getRange(cursor, i + 2)
      .setValue(lang)
      .setBackground(OV_COLORS.SUBHEADER_BG)
      .setFontColor(OV_COLORS.SUBHEADER_FG)
      .setFontSize(10)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
  });
  sheet.setRowHeights(cursor, 1, 30);
  cursor++;

  // Data rows
  sheetOrder.forEach(function(sheetName, rowIdx) {
    var rowBg = (rowIdx % 2 === 0) ? OV_COLORS.WHITE : OV_COLORS.LIGHT_GREY;

    sheet.getRange(cursor, 1)
      .setValue(sheetName)
      .setBackground(rowBg)
      .setFontSize(11)
      .setFontWeight('bold')
      .setFontColor(OV_COLORS.TEXT_PRIMARY)
      .setHorizontalAlignment('left')
      .setVerticalAlignment('middle');

    allLangs.forEach(function(lang, i) {
      var cell = sheet.getRange(cursor, i + 2);
      var entry = lookup[sheetName] && lookup[sheetName][lang];

      if (!entry || entry.total === 0) {
        cell.setValue('—')
          .setBackground(OV_COLORS.PCT_NONE)
          .setFontColor(OV_COLORS.TEXT_MUTED)
          .setFontSize(11)
          .setHorizontalAlignment('center')
          .setVerticalAlignment('middle');
      } else {
        var pct = Math.round((entry.reviewed / entry.total) * 100);
        cell.setValue(pct + '%')
          .setBackground(pctToColor(pct))
          .setFontColor(pctToFontColor(pct))
          .setFontSize(11)
          .setFontWeight('bold')
          .setHorizontalAlignment('center')
          .setVerticalAlignment('middle')
          .setNote('Reviewed: ' + entry.reviewed + ' / ' + entry.total + '\nTranslated: ' + entry.translated + '\nDisputed: ' + entry.disputed + '\nUntranslated: ' + entry.untranslated);
      }
    });

    sheet.setRowHeights(cursor, 1, 28);
    cursor++;
  });

  // ── Column widths (single pass, screen-fit) ───────────────────────────────
  //
  // Target: fill a 1920×1080 screen.
  // Google Sheets chrome (row headers, sidebar, scrollbar) ≈ 220 px.
  // Usable width ≈ 1700 px.
  //
  // Layout:
  //   Col 1        : 220 px  (sheet name / language label)
  //   Last col     : 260 px  (sparkline progress bar)
  //   Middle cols  : remaining width shared equally, min 90 px each
  //
  // With 8 languages (totalDataCols=9):
  //   220 + 7×middleCols + 260 = 1700  →  middleCols = (1700-480)/7 ≈ 174 px ✓
  //
  var totalDataCols = Math.max(9, 1 + allLangs.length);
  var COL1_WIDTH    = 220;
  var LAST_WIDTH    = 260;
  var TOTAL_WIDTH   = 1700;

  var middleCount = totalDataCols - 2; // cols between col1 and last
  var midWidth = middleCount > 0
    ? Math.floor((TOTAL_WIDTH - COL1_WIDTH - LAST_WIDTH) / middleCount)
    : 120;
  midWidth = Math.max(midWidth, 90); // never narrower than 90 px

  sheet.setColumnWidth(1, COL1_WIDTH);
  for (var wc = 2; wc <= totalDataCols - 1; wc++) {
    sheet.setColumnWidth(wc, midWidth);
  }
  sheet.setColumnWidth(totalDataCols, LAST_WIDTH);

  // Delete excess rows (cursor now points to first unused row)
  var usedRows = cursor - 1;
  var currentMaxRows = sheet.getMaxRows();
  if (currentMaxRows > usedRows) {
    sheet.deleteRows(usedRows + 1, currentMaxRows - usedRows);
  }

  // Delete excess columns beyond what we use
  var currentMaxCols = sheet.getMaxColumns();
  if (currentMaxCols > totalDataCols) {
    sheet.deleteColumns(totalDataCols + 1, currentMaxCols - totalDataCols);
  }

  // Freeze just the title + timestamp rows so the page header stays visible.
  try { sheet.setFrozenRows(2); } catch(e) { /* ignore */ }

  // Hard-lock the sheet: remove all editors so only the spreadsheet owner can edit.
  // The script itself bypasses protection when it runs as the owner.
  var protection = sheet.protect().setDescription('PDG Localisation: auto-generated, do not edit manually');
  protection.setWarningOnly(false);
  // Remove every editor except the owner (owner cannot be removed)
  var editors = protection.getEditors();
  if (editors.length > 0) {
    protection.removeEditors(editors);
  }

  Logger.log('Overview sheet refreshed: ' + sheetOrder.length + ' sheets, ' + allLangs.length + ' languages');
}

// ── Helpers ───────────────────────────────────────────────────────────────────

/**
 * Returns a background hex colour for a reviewed percentage.
 * @param {number} pct  0–100
 * @returns {string}
 */
function pctToColor(pct) {
  if (pct === 100) return OV_COLORS.PCT_100;
  if (pct >= 75)   return OV_COLORS.PCT_75;
  if (pct >= 50)   return OV_COLORS.PCT_50;
  if (pct >= 25)   return OV_COLORS.PCT_25;
  return OV_COLORS.PCT_0;
}

/**
 * Returns a contrasting font colour (dark on light cells, white on dark cells).
 * @param {number} pct
 * @returns {string}
 */
function pctToFontColor(pct) {
  if (pct >= 50) return OV_COLORS.PCT_FG_DARK;
  return OV_COLORS.PCT_FG_LIGHT;
}

/**
 * Opens a linked spreadsheet by ID and reads its 📊 ProgressCache sheet.
 * Returns an array in the same format as getCachedProgress(), or [] on error.
 *
 * @param {{ name: string, id: string }} ssEntry
 * @returns {Array}
 */
function readLinkedProgress(ssEntry) {
  try {
    var ss = SpreadsheetApp.openById(ssEntry.id);
    var cacheSheet = ss.getSheetByName(PROGRESS_CACHE_SHEET_NAME);
    if (!cacheSheet) return [];

    var lastRow = cacheSheet.getLastRow();
    if (lastRow < 2) return [];

    var data = cacheSheet.getRange(2, 1, lastRow - 1, 8).getValues();
    var result = [];
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var total = Number(row[2]);
      if (total > 0) {
        result.push({
          sheetName: String(row[0]),
          language:  String(row[1]),
          total:       total,
          translated:  Number(row[3]),
          reviewed:    Number(row[4]),
          disputed:    Number(row[5]),
          untranslated: Number(row[6])
        });
      }
    }
    return result;
  } catch (e) {
    Logger.log('readLinkedProgress failed for "' + ssEntry.name + '": ' + e);
    return [];
  }
}

/**
 * Gets or creates the 📊 Overview sheet, positioned as the first tab.
 * @returns {Sheet}
 */
function getOrCreateOverviewSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(OVERVIEW_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(OVERVIEW_SHEET_NAME, 0); // insert at position 0 = first tab
  } else {
    // Move to first position if not already there
    ss.setActiveSheet(sheet);
    ss.moveActiveSheet(1);
  }

  return sheet;
}
