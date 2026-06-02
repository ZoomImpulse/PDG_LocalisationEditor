// ─────────────────────────────────────────────────────────────────────────────
// Config.gs — Constants and shared helpers
// ─────────────────────────────────────────────────────────────────────────────

// Spreadsheet ID of the main spreadsheet that should show the Global Overview.
// Only this spreadsheet will render the cross-spreadsheet summary section.
// Copy from the URL: /spreadsheets/d/<ID>/edit
var MAIN_SPREADSHEET_ID = '1jQOrWJpAF_9TQVyrrOfxinyTTxvoDJg_E7BHUNEkoio';

// Linked spreadsheets shown in the Global Overview section of 📊 Overview.
// Add one entry per spreadsheet. The 'id' is the Google Sheets spreadsheet ID
// (from the URL: /spreadsheets/d/<ID>/edit). Leave the array empty to disable.
var LINKED_SPREADSHEETS = [
   { name: 'Events',      id: '1YNdrUt0Ro1w6aiVZR0uSJnnulpzhh4thvy3K1-fJ_qA' },
   { name: 'Tech',      id: '15QcA1M4dX455UYD2GEv3tDJ3P4z3jhK7p5qPMTDFS60' },
   { name: 'Ships', id: '19z068O5ARdrXLyswqTeDqcQdhAwA39kI8Gx_nhZPL3I' },
   { name: 'Modifiers',    id: '1TZylnt8An15CLYlQmy1tjUYvHgMQoosh_x1jC35HOck' },
   { name: 'Synced',   id: '1MgcmiOr8OMqD6qo5EMwk3ymVenSqAS8MWdo33hKjIPk' },
];

// ID of the single spreadsheet that hosts the shared 🔐 Roles sheet.
// Set this to the spreadsheet ID of your designated roles spreadsheet.
// Leave as '' to fall back to the active spreadsheet (per-spreadsheet roles).
var ROLES_SPREADSHEET_ID = '1jQOrWJpAF_9TQVyrrOfxinyTTxvoDJg_E7BHUNEkoio';

// Status values stored in the separate status tracking sheet
var STATUS = {
  STATIC:       'STATIC',
  UNTRANSLATED: 'UNTRANSLATED',
  TRANSLATED:   'TRANSLATED',
  REVIEWED:     'REVIEWED',
  DISPUTED:     'DISPUTED',
};

// Cell background colors matching the existing workflow
var STATUS_COLOR = {
  STATIC:       '#d9d9d9',
  UNTRANSLATED: '#e06666',
  TRANSLATED:   '#ffff00',
  REVIEWED:     '#00ff00',
  DISPUTED:     '#ff0000',
};

// Reverse map: hex color → status value (for seeding from existing colors)
var COLOR_TO_STATUS = {
  '#d9d9d9': STATUS.STATIC,
  '#e06666': STATUS.UNTRANSLATED,
  '#ffff00': STATUS.TRANSLATED,
  '#00ff00': STATUS.REVIEWED,
  '#ff0000': STATUS.DISPUTED,
};

// Statuses the translator can manually set (STATIC is read-only)
var SELECTABLE_STATUSES = [
  STATUS.UNTRANSLATED,
  STATUS.TRANSLATED,
  STATUS.REVIEWED,
  STATUS.DISPUTED,
];

// Statuses that are shown when "Untranslated only" filter is active
var UNTRANSLATED_FILTER_STATUSES = [STATUS.UNTRANSLATED, ''];

/**
 * Returns true if the sheet should be completely ignored by all script
 * functions. A sheet is ignored when its name is entirely uppercase
 * (e.g. STATS, VALUES, DATA) — these are stats/value sheets, not
 * localisation sheets.
 *
 * The special dashboard and status sheets are also excluded here.
 *
 * @param {string} name  Sheet tab name
 * @returns {boolean}
 */
function isIgnoredSheet(name) {
  if (!name) return true;
  // System sheets - never show in picker
  if (name === '📊 Overview') return true;
  if (name === '📊 Progress') return true;
  if (name === '📊 ProgressCache') return true;
  if (name === '📋 Status') return true;
  if (name === '🔐 Roles') return true;
  if (name === 'Debugging') return true;
  if (name === 'Name List') return true;
  // A sheet is ignored when its name contains only uppercase letters, digits,
  // underscores and spaces — i.e. no lowercase letters at all.
  // Examples: STATS, VALUES, DATA, ARMY_VALUES → ignored
  // Examples: Armies, Empire, Governments      → kept
  return /^[^a-z]+$/.test(name);
}

/**
 * Extracts a plain language name from a header cell that may contain
 * extra descriptive text.
 *
 * Only headers containing parentheses are considered language columns.
 * This prevents non-language columns (e.g., "Description", "Notes")
 * from being treated as translation columns.
 *
 * Examples:
 *   "Localisation (English)"        → "English"
 *   "Name Localisation (English)"   → "English"
 *   "German"                        → "" (no parens, not a language col)
 *   "Description"                   → "" (no parens, not a language col)
 *
 * Strategy:
 *   1. If the value contains parentheses, return the content inside the
 *      last pair of parentheses.
 *   2. Otherwise return empty string (not a language column).
 *
 * @param {string} raw  Raw cell value from row 2
 * @returns {string}    Extracted language name, or empty string if not a language column
 */
function extractLanguageName(raw) {
  if (!raw) return '';
  var s = String(raw).trim();
  var match = s.match(/\(([^)]+)\)\s*$/);
  if (match) return match[1].trim();
  return '';
}

/**
 * Returns the background color of a cell as a lowercase hex string,
 * or '' if the cell has no background / white background.
 *
 * @param {Range} cell  A single-cell Range object
 * @returns {string}    e.g. '#e06666' or ''
 */
function getCellColor(cell) {
  var bg = cell.getBackground();
  if (!bg || bg === '#ffffff' || bg === 'white') return '';
  return bg.toLowerCase();
}
