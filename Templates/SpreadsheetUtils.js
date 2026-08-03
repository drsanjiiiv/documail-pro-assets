/**
 * SpreadsheetUtils.gs (Inside your Document Template Project)
 * Fetches headers up to the 'Recipient Email' baseline column boundary.
 */
function getSheetHeaders() {
  try {
    var activeDoc = DocumentApp.getActiveDocument();
    var docId = activeDoc.getId();
    
    var docFile = DriveApp.getFileById(docId);
    var sourceSheetId = docFile.getDescription();
    
    if (!sourceSheetId) {
      throw new Error("No linked sheet found. This template was not created from a DocuMail Pro sheet. Please use 'Create Dynamic Doc Template From Sheet' in your spreadsheet first.");
    }
    
    var ss = SpreadsheetApp.openById(sourceSheetId);
    var sheet = ss.getSheets()[0]; 
    var lastColumn = sheet.getLastColumn();
    
    if (lastColumn === 0) return [];
    
    var rawHeaders = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
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