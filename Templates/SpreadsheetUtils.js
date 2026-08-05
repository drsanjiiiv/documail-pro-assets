/**
 * SpreadsheetUtils.gs (Inside your Document Template Project)
 * Fetches headers up to the 'Recipient Email' baseline column boundary.
 * Uses the Sheets Advanced Service so this add-on stays on the
 * spreadsheets.readonly scope - no full SpreadsheetApp access required.
 */
function getSheetHeaders() {
  try {
    var activeDoc = DocumentApp.getActiveDocument();

    // The Sheets add-on stores the linked spreadsheet ID as a document
    // property at init time (footer is read once and cleared). Property
    // first, footer as fallback for templates linked before this change.
    var sourceSheetId = PropertiesService.getDocumentProperties().getProperty('DOCUMAIL_SOURCE_SHEET_ID');
    if (!sourceSheetId) {
      var footer = activeDoc.getFooter();
      if (footer) {
        var footerText = footer.getText();
        var match = footerText.match(/DOCUMAIL_SOURCE_SHEET_ID=([\w-]+)/);
        if (match) sourceSheetId = match[1];
      }
    }

    if (!sourceSheetId) {
      throw new Error("No linked sheet found. This template was not created from a DocuMail Pro sheet. Please use 'Create Dynamic Doc Template From Sheet' in your spreadsheet first.");
    }

    var result = Sheets.Spreadsheets.Values.get(sourceSheetId, '1:1');
    var rawHeaders = result.values ? result.values[0] : [];
    var filteredHeaders = [];

    // --- BOUNDARY TRUNCATION ENGINE ---
    for (var i = 0; i < rawHeaders.length; i++) {
      var headerVal = rawHeaders[i];
      if (!headerVal) continue;

      var cleanHeader = headerVal.toString().trim();
      if (cleanHeader === "") continue;

      var hStr = cleanHeader.toLowerCase();
      if (hStr.indexOf("merged doc status") !== -1) break;
      if (hStr.indexOf("merged doc id") !== -1) continue;
      if (hStr.indexOf("merged doc url") !== -1) continue;
      if (hStr.indexOf("sent mail status") !== -1) continue;

      filteredHeaders.push(cleanHeader);
    }

    return filteredHeaders;

  } catch (error) {
    Logger.log("Error fetching automated sheet layout: " + error.toString());
    throw new Error("Could not automatically connect to your data sheet: " + error.message);
  }
}