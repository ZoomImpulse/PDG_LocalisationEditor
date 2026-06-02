// ─────────────────────────────────────────────────────────────────────────────
// Stub.gs — Thin wrapper bound to each spreadsheet.
//
// HOW THIS WORKS
// ──────────────
// All real logic lives in the shared "LocalisationMasterTool" Apps Script Library.
// This file is copy-pasted ONCE into each of the 5 other spreadsheets and then
// NEVER needs to be touched again — update the library instead.
//
// SETUP STEPS (one-time per spreadsheet)
// ───────────────────────────────────────
// 1. Open the MAIN spreadsheet → Extensions → Apps Script
//    (project is called "LocalisationMasterTool").
//    Deploy → New deployment → Type: Library.
//    Copy the Script ID from Project Settings (⚙).
// 2. In the Apps Script editor for THIS spreadsheet, add the library:
//    Click ⊕ next to "Libraries" in the left panel.
//    Script ID : <paste the Script ID from step 1>
//    Version   : pick the latest numbered version (or Development for testing)
//    Identifier: PDGLib   ← must match exactly
// 3. Delete every other .gs file and Modal.html from this project — they
//    now live in the library.
// 4. Keep only this Stub.gs file.
// 5. Save and reload the spreadsheet.
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

// ── Menu action delegates ─────────────────────────────────────────────────────

function openTranslatorModal()  {
  // Auto-installs progress cache trigger on first use per spreadsheet
  PDGLib.ensureProgressCacheTrigger();
  PDGLib.openTranslatorModal();
}
function refreshOverviewSheet() { PDGLib.refreshOverviewSheet(); }
function insertStatusColumns()  { PDGLib.insertStatusColumns(); }
function initialiseRoles()      { PDGLib.initialiseRoles(); }

// ── google.script.run delegates ───────────────────────────────────────────────
// Modal.html calls these by name; they must exist in the bound script.
// Each one simply forwards to the library.

function getSheetNames()                                          { return PDGLib.getSheetNames(); }
function getSheetInfo(sheetName)                                  { return PDGLib.getSheetInfo(sheetName); }
function getFullSheetData(sheetName)                              { return PDGLib.getFullSheetData(sheetName); }
function getProgressForLanguage(targetLanguage)                   { return PDGLib.getProgressForLanguage(targetLanguage); }
function getCachedProgress()                                      { return PDGLib.getCachedProgress(); }
function getCacheLastUpdated()                                    { return PDGLib.getCacheLastUpdated(); }
function forceRefreshProgressCache()                              { return PDGLib.forceRefreshProgressCache(); }
function searchAllSheets(query, maxResults)                       { return PDGLib.searchAllSheets(query, maxResults); }
function saveUserLastLang(lang)                                   { return PDGLib.saveUserLastLang(lang); }
function saveTranslation(sheetName, rowIndex, colIndex, statusCol, value, status, cascadeColIndices) {
  return PDGLib.saveTranslation(sheetName, rowIndex, colIndex, statusCol, value, status, cascadeColIndices);
}
function getUserRoles()                                           { return PDGLib.getUserRoles(); }
function setUserRole(email, role)                                 { return PDGLib.setUserRole(email, role); }
function removeUserRole(email)                                    { return PDGLib.removeUserRole(email); }
