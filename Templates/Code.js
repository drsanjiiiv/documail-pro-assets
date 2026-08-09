// ==========================================
// DOCUMENT ADD-ON - ONOPEN (SINGLE ENTRY POINT)
// ==========================================
function onOpen(e) {
  createMenu();
  if (e && e.authMode === ScriptApp.AuthMode.NONE) {
    return;
  }
  try {
    // Fire the automatic guiding toast on-screen instantly when opened
    DocumentApp.getActiveDocument().toast(
      "To start creating your template rules, use: Extensions > DocuMail Pro Template > Start Dynamic Doc Template from the menu above.",
      "🚀 DocuMail Pro Engine",
      12
    );
  } catch (err) {
    Logger.log("Error loading document menu layout: " + err.toString());
  }
}

function onInstall() {
  onOpen();
}

/**
 * Creates the custom add-on menu in Google Docs (under Extensions)
 */
function createMenu() {
  var ui = DocumentApp.getUi();
  ui.createAddonMenu()
    .addItem('📄 Initialize DocuMail Pro Template', 'INITIALIZE_DOC_DESIGNER_SIDEBAR')
    .addItem('📝 Open Smart Variable Window', 'OPEN_SMART_VARIABLE_WINDOW')
    .addItem('🔗 Link to Spreadsheet', 'OPEN_SMART_VARIABLE_WINDOW_FROM_LINK_MENU')
    .addSeparator()
    .addItem('❓ Help', 'showHelp')
    .addToUi();
}
