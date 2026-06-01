// ─────────────────────────────────────────────────────────────────────────────
// Roles.gs — Role-based access control
// ─────────────────────────────────────────────────────────────────────────────

var ROLES_SHEET_NAME = '🔐 Roles';

var ROLE = {
  ADMIN:      'ADMIN',
  REVIEWER:   'REVIEWER',
  TRANSLATOR: 'TRANSLATOR',
  NONE:       'NONE',
};

// ── Sheet helpers ─────────────────────────────────────────────────────────────

/**
 * Returns the Spreadsheet object that hosts the 🔐 Roles sheet.
 * Uses ROLES_SPREADSHEET_ID when set; falls back to the active spreadsheet.
 *
 * @returns {Spreadsheet}
 */
function getRolesSpreadsheet() {
  if (ROLES_SPREADSHEET_ID) {
    return SpreadsheetApp.openById(ROLES_SPREADSHEET_ID);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * Gets or creates the roles sheet.
 * Columns: Email | Role | AddedBy | Timestamp
 *
 * @returns {Sheet}
 */
function getRoleSheet() {
  var ss = getRolesSpreadsheet();
  var sheet = ss.getSheetByName(ROLES_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(ROLES_SHEET_NAME);
    sheet.hideSheet();
    sheet.getRange(1, 1, 1, 4).setValues([['Email', 'Role', 'AddedBy', 'Timestamp']]);
    sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
    sheet.setColumnWidth(1, 240);
    sheet.setColumnWidth(2, 120);
    sheet.setColumnWidth(3, 240);

    // Protect: show warning to non-owners
    var protection = sheet.protect().setDescription('PDG Translate: Roles (do not edit manually)');
    protection.setWarningOnly(true);
  }
  return sheet;
}

/**
 * Returns the email of the spreadsheet owner.
 *
 * @returns {string}
 */
function getOwnerEmail() {
  return SpreadsheetApp.getActiveSpreadsheet().getOwner().getEmail().toLowerCase();
}

// ── Role resolution ───────────────────────────────────────────────────────────

/**
 * Returns the role for the currently active user.
 * Owner always returns ADMIN regardless of the sheet.
 *
 * @returns {string} One of ROLE values
 */
function getCurrentUserRole() {
  var email = Session.getActiveUser().getEmail().toLowerCase();
  if (!email) return ROLE.TRANSLATOR;

  // Owner is always ADMIN — no sheet read needed
  if (email === getOwnerEmail()) return ROLE.ADMIN;

  // Check CacheService first (60 s TTL) to avoid sheet scan on every save
  var cache = CacheService.getUserCache();
  var cacheKey = 'pdg_role_' + email;
  var cached = cache.get(cacheKey);
  if (cached && ROLE[cached]) return cached;

  var sheet = getRoleSheet();
  var lastRow = sheet.getLastRow();
  var resolved = ROLE.TRANSLATOR;

  if (lastRow >= 2) {
    var data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    for (var i = 0; i < data.length; i++) {
      var rowEmail = String(data[i][0] || '').toLowerCase().trim();
      var rowRole  = String(data[i][1] || '').toUpperCase().trim();
      if (rowEmail === email && ROLE[rowRole]) {
        resolved = rowRole;
        break;
      }
    }
  }

  cache.put(cacheKey, resolved, 60);
  return resolved;
}

// ── Admin functions ───────────────────────────────────────────────────────────

/**
 * Returns all role assignments (Admin use only).
 *
 * @returns {Array} [{ email, role }]
 */
function getUserRoles() {
  var callerRole = getCurrentUserRole();
  if (callerRole !== ROLE.ADMIN) throw new Error('Access denied: Admin only.');

  var sheet = getRoleSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  var result = [];
  for (var i = 0; i < data.length; i++) {
    var email = String(data[i][0] || '').trim();
    var role  = String(data[i][1] || '').trim();
    if (email && role) {
      result.push({ email: email, role: role });
    }
  }
  return result;
}

/**
 * Adds or updates a user's role.
 * Only callable by ADMIN users.
 *
 * @param {string} email
 * @param {string} role  One of ADMIN / REVIEWER / TRANSLATOR
 */
function setUserRole(email, role) {
  var callerRole = getCurrentUserRole();
  if (callerRole !== ROLE.ADMIN) throw new Error('Access denied: Admin only.');

  email = String(email || '').toLowerCase().trim();
  role  = String(role  || '').toUpperCase().trim();

  if (!email) throw new Error('Email is required.');
  if (!ROLE[role] || role === ROLE.NONE) throw new Error('Invalid role: ' + role);

  var caller = Session.getActiveUser().getEmail().toLowerCase();
  var sheet = getRoleSheet();
  var lastRow = sheet.getLastRow();

  // Check if the email already exists → update in place
  if (lastRow >= 2) {
    var data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < data.length; i++) {
      var existing = String(data[i][0] || '').toLowerCase().trim();
      if (existing === email) {
        var sheetRow = i + 2;
        sheet.getRange(sheetRow, 1, 1, 4).setValues([[email, role, caller, new Date()]]);
        return;
      }
    }
  }

  // New entry
  sheet.appendRow([email, role, caller, new Date()]);
}

/**
 * Removes a user's role assignment.
 * Only callable by ADMIN users.
 *
 * @param {string} email
 */
function removeUserRole(email) {
  var callerRole = getCurrentUserRole();
  if (callerRole !== ROLE.ADMIN) throw new Error('Access denied: Admin only.');

  email = String(email || '').toLowerCase().trim();
  if (!email) throw new Error('Email is required.');

  var sheet = getRoleSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  var data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = data.length - 1; i >= 0; i--) {
    var existing = String(data[i][0] || '').toLowerCase().trim();
    if (existing === email) {
      sheet.deleteRow(i + 2);
      return;
    }
  }
}

// ── User preferences ──────────────────────────────────────────────────────────

/**
 * Persists the user's last selected translation language.
 *
 * @param {string} lang  Language name (e.g. 'German')
 */
function saveUserLastLang(lang) {
  PropertiesService.getUserProperties().setProperty('pdg_last_lang', String(lang || ''));
}

/**
 * Returns the user's last selected translation language, or '' if not set.
 *
 * @returns {string}
 */
function getUserLastLang() {
  return PropertiesService.getUserProperties().getProperty('pdg_last_lang') || '';
}

// ── Setup ─────────────────────────────────────────────────────────────────────

/**
 * One-time setup: creates the Roles sheet and ensures the owner has ADMIN.
 * Safe to re-run.
 */
function initialiseRoles() {
  var ui = SpreadsheetApp.getUi();
  var sheet = getRoleSheet(); // creates if not exists

  var ownerEmail = getOwnerEmail();
  var lastRow = sheet.getLastRow();
  var alreadyHasOwner = false;

  if (lastRow >= 2) {
    var data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    for (var i = 0; i < data.length; i++) {
      var e = String(data[i][0] || '').toLowerCase().trim();
      var r = String(data[i][1] || '').toUpperCase().trim();
      if (e === ownerEmail && r === ROLE.ADMIN) {
        alreadyHasOwner = true;
        break;
      }
    }
  }

  if (!alreadyHasOwner) {
    sheet.appendRow([ownerEmail, ROLE.ADMIN, 'setup', new Date()]);
  }

  ui.alert(
    'PDG Translate — Roles',
    'Roles sheet ready.\n\nOwner (' + ownerEmail + ') has been granted ADMIN.\n\n' +
    'Open the translator and use the 👤 button to manage user roles.',
    ui.ButtonSet.OK
  );
}
