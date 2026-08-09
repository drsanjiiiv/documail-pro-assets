/**
 * SpreadsheetUtils.gs (Inside your Document Template Project)
 * Fetches headers up to the 'Recipient Email' baseline column boundary.
 * Uses the Sheets Advanced Service (read-only calls) under the approved
 * spreadsheets scope - no full SpreadsheetApp access required.
 */

/**
 * Resolves the linked source spreadsheet ID from document properties first,
 * then the footer (legacy templates linked before the property handoff).
 * Returns null when no sheet is linked.
 */
function RESOLVE_SOURCE_SHEET_ID() {
  var sourceSheetId = PropertiesService.getDocumentProperties().getProperty('DOCUMAIL_SOURCE_SHEET_ID');
  if (!sourceSheetId) {
    var activeDoc = DocumentApp.getActiveDocument();
    var footer = activeDoc.getFooter();
    if (footer) {
      var footerText = footer.getText();
      var match = footerText.match(/DOCUMAIL_SOURCE_SHEET_ID=([\w-]+)/);
      if (match) sourceSheetId = match[1];
    }
  }
  return sourceSheetId || null;
}

/**
 * Returns the current linking status for this template document.
 * @return {Object} { linked: boolean, sheetId: string|null, sheetName: string|null }
 */
function GET_LINK_STATUS() {
  var sourceSheetId = RESOLVE_SOURCE_SHEET_ID();
  if (!sourceSheetId) {
    return { linked: false, sheetId: null, sheetName: null };
  }
  var sheetName = null;
  try {
    var sheetInfo = Sheets.Spreadsheets.get(sourceSheetId);
    sheetName = sheetInfo.properties.title;
  } catch (e) {
    Logger.log("Could not fetch linked sheet name: " + e.toString());
  }
  return { linked: true, sheetId: sourceSheetId, sheetName: sheetName };
}

/**
 * Returns the OAuth token used by the Google Picker (same pattern as the
 * Sheets add-on). No new scopes required - uses the add-on's current token.
 */
function GET_SYSTEM_OAUTH_TOKEN() {
  try {
    return ScriptApp.getOAuthToken();
  } catch (e) {
    return "";
  }
}

/**
 * Links this template document to a spreadsheet, then returns the filtered
 * headers so the sidebar can refresh its variable dropdown.
 * @param {string} spreadsheetId - ID of the spreadsheet to link.
 * @return {Object} { success, headers, sheetName, matched, unmatched } or { success:false, error }
 */
function LINK_SHEET_TO_TEMPLATE(spreadsheetId) {
  try {
    if (!spreadsheetId || String(spreadsheetId).trim() === "") {
      return { success: false, error: "No spreadsheet was selected." };
    }
    var sheetId = String(spreadsheetId).trim();

    // Validate access by reading row 1 via the Sheets API (existing readonly scope).
    var sheetInfo = Sheets.Spreadsheets.get(sheetId);
    var sheetName = sheetInfo.properties.title;
    var result = Sheets.Spreadsheets.Values.get(sheetId, '1:1');
    var rawHeaders = result.values ? result.values[0] : [];

    // Persist the link (property first, footer as resilience).
    PropertiesService.getDocumentProperties().setProperty('DOCUMAIL_SOURCE_SHEET_ID', sheetId);
    try {
      var activeDoc = DocumentApp.getActiveDocument();
      var footer = activeDoc.getFooter() || activeDoc.addFooter();
      footer.setText('DOCUMAIL_SOURCE_SHEET_ID=' + sheetId);
    } catch (e) {
      Logger.log("Could not write linked sheet ID to footer: " + e.toString());
    }

    var filteredHeaders = FILTER_HEADERS(rawHeaders);

    // Compare the document's existing {tags} against the new headers.
    var tagCheck = EXTRACT_DOC_TAGS_AND_MATCH(filteredHeaders);

    // Linking a sheet makes this document a working DocuMail Pro template.
    PropertiesService.getDocumentProperties().setProperty('DOCUMAIL_TEMPLATE_INITIALIZED', 'true');

    return {
      success: true,
      headers: filteredHeaders,
      sheetName: sheetName,
      matched: tagCheck.matched,
      unmatched: tagCheck.unmatched
    };

  } catch (error) {
    var message = String(error.message || error).toLowerCase();
    if (message.indexOf("cannot find") !== -1 || message.indexOf("spreadsheet not found") !== -1 || message.indexOf("400") !== -1) {
      return { success: false, error: "That doesn't look like a valid spreadsheet ID. Please pick a Google Sheet." };
    }
    if (message.indexOf("permission") !== -1 || message.indexOf("403") !== -1 || message.indexOf("401") !== -1 || message.indexOf("access") !== -1) {
      return { success: false, error: "You don't have access to that spreadsheet. Open it once, then try again." };
    }
    return { success: false, error: "Could not link that spreadsheet: " + message };
  }
}

/**
 * Applies the boundary truncation engine to a raw header row.
 */
function FILTER_HEADERS(rawHeaders) {
  var filteredHeaders = [];
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
}

/**
 * Extracts {variable} tags from the active document body (Docs-side, no sheet
 * dependency) and diffs them against the given headers.
 * @return {Object} { matched: string[], unmatched: string[] }
 */
function EXTRACT_DOC_TAGS_AND_MATCH(headers) {
  var docTags = [];
  try {
    var activeDoc = DocumentApp.getActiveDocument();
    var bodyText = activeDoc.getBody().getText();
    var tagMatchRegex = /\{([^}]+)\}/g;
    var matchItem;
    while ((matchItem = tagMatchRegex.exec(bodyText)) !== null) {
      var cleanToken = matchItem[1].trim();
      if (cleanToken && docTags.indexOf(cleanToken) === -1) {
        docTags.push(cleanToken);
      }
    }
  } catch (e) {
    Logger.log("Could not extract document tags: " + e.toString());
  }

  var headerSet = {};
  for (var h = 0; h < headers.length; h++) {
    headerSet[String(headers[h]).toLowerCase().trim()] = true;
  }

  var matched = [];
  var unmatched = [];
  for (var t = 0; t < docTags.length; t++) {
    if (headerSet[docTags[t].toLowerCase().trim()]) {
      matched.push(docTags[t]);
    } else {
      unmatched.push(docTags[t]);
    }
  }
  return { matched: matched, unmatched: unmatched };
}

function getSheetHeaders() {
  try {
    var sourceSheetId = RESOLVE_SOURCE_SHEET_ID();

    if (!sourceSheetId) {
      throw new Error("No linked sheet found. This template was not created from a DocuMail Pro sheet. Please use 'Create Dynamic Doc Template From Sheet' in your spreadsheet first.");
    }

    var result = Sheets.Spreadsheets.Values.get(sourceSheetId, '1:1');
    var rawHeaders = result.values ? result.values[0] : [];

    return FILTER_HEADERS(rawHeaders);

  } catch (error) {
    Logger.log("Error fetching automated sheet layout: " + error.toString());
    throw new Error("Could not automatically connect to your data sheet: " + error.message);
  }
}
