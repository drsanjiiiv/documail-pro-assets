/**file name: email.gs
/**
 * ============================================================================
 * DOCUMAIL PRO COMPLETE MASTER CORE SCRIPT
 * ============================================================================
 * FILE: email.gs - Email Engine and Quota Management
 * ============================================================================
 */

function executeEmailSend(selectedRows, emailConfig, templateName, tagMappings, optSheet) {
  var sheet = optSheet || SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var allHeaders = GET_ALL_RAW_HEADERS(sheet);
  var emailColIdx = -1;

  for (var c = 0; c < allHeaders.length; c++) {
    var hName = String(allHeaders[c]).toLowerCase().trim();
    if (hName.indexOf("recipient email") !== -1) emailColIdx = c;
  }
  if (emailColIdx === -1) emailColIdx = 4;

  var statusColIdx = GET_OR_CREATE_STATUS_COLUMN(sheet, templateName);

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, allHeaders.length).getValues();
  var displayData = sheet.getRange(2, 1, sheet.getLastRow() - 1, allHeaders.length).getDisplayValues();
  var sent = 0;
  var errors = [];
  
  // Use manual mappings if provided, fallback to standard header-matching
  var mappings = tagMappings || {};

  for (var i = 0; i < data.length; i++) {
    var rowNum = i + 2;
    var found = false;
    for (var s = 0; s < selectedRows.length; s++) {
      if (selectedRows[s] === rowNum) {
        found = true;
        break;
      }
    }
    if (!found) continue;

    var currentStatus = data[i][statusColIdx];
    if (currentStatus && currentStatus.toString().trim() !== "" && currentStatus.toString().indexOf("Failed") === -1) {
      continue;
    }

    var recipient = RESOLVE_RECIPIENT(emailConfig, data[i], allHeaders, emailColIdx);
    if (!recipient || recipient.indexOf("@") === -1) continue;

    var subject = emailConfig.subject || "";
    var body = emailConfig.body || "";

    // =======================================================
    // ✅ FIX: ITERATE OVER THE DICTIONARY MAPPINGS LIVE
    // =======================================================
    if (Object.keys(mappings).length > 0) {
      for (var docTag in mappings) {
        if (mappings.hasOwnProperty(docTag)) {
          var targetHeader = mappings[docTag];
          var headerIdx = allHeaders.indexOf(targetHeader);
          
          if (headerIdx !== -1) {
            var rawValue = data[i][headerIdx];
            var val = "";
            
            if (rawValue instanceof Date) {
              val = FORMAT_DATE_FOR_DISPLAY(rawValue);
            } else if (typeof rawValue === 'number') {
              val = displayData[i][headerIdx] || FORMAT_NUMBER_FOR_DISPLAY(rawValue);
            } else {
              val = String(rawValue || "");
            }
            
            var regex = new RegExp("\\{" + escapeRegex(docTag) + "\\}", "g");
            subject = subject.replace(regex, val);
            body = body.replace(regex, val);
          }
        }
      }
    } else {
      // Fallback behavior if no template mapping configurations exist
      for (var h = 0; h < allHeaders.length; h++) {
        var header = allHeaders[h];
        if (header) {
          var rawValue = data[i][h];
          var val = "";
          if (rawValue instanceof Date) {
            val = FORMAT_DATE_FOR_DISPLAY(rawValue);
          } else if (typeof rawValue === 'number') {
            val = displayData[i][h] || FORMAT_NUMBER_FOR_DISPLAY(rawValue);
          } else {
            val = String(rawValue || "");
          }
          var regex = new RegExp("\\{" + escapeRegex(header) + "\\}", "g");
          subject = subject.replace(regex, val);
          body = body.replace(regex, val);
        }
      }
    }

    var mailOptions = { htmlBody: body };
    if (emailConfig.replyTo) mailOptions.replyTo = emailConfig.replyTo;
    if (emailConfig.cc) mailOptions.cc = emailConfig.cc;
    if (emailConfig.bcc) mailOptions.bcc = emailConfig.bcc;

    try {
      GmailApp.sendEmail(recipient, subject, "", mailOptions);
      var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
      var ccSuffix = emailConfig.cc ? " (CC: " + emailConfig.cc + ")" : "";
      var bccSuffix = emailConfig.bcc ? " (BCC: " + emailConfig.bcc + ")" : "";
      sheet.getRange(rowNum, statusColIdx + 1).setValue("Sent to " + recipient + " on " + timestamp + ccSuffix + bccSuffix);
      sent++;
    } catch (e) {
      errors.push("Row " + rowNum + ": " + e.message);
      sheet.getRange(rowNum, statusColIdx + 1).setValue("Failed: " + e.message);
    }
  }

  if (errors.length > 0) {
    return "✅ Sent " + sent + " emails.\n⚠️ Errors: " + errors.join(", ");
  }
  return "✅ Successfully sent " + sent + " emails!";
}

// =======================================================
// ✅ FIX: PASS REFRESHED PARAM MAPPINGS VIA CONTROLLER
// =======================================================
function EXECUTE_TEMPLATE_ACTION(params) {
    var mappings = (params.config && params.config.tagMappings) ? params.config.tagMappings : {};
    
    if (params.type === "EMAIL_ONLY") {
        return executeEmailSend(params.selectedRows, params.emailConfig, params.templateName, mappings);
    }
    return "PDF generation coming soon";
}

function CHECK_SIDEBAR_REFRESH() {
  try {
    var props = PropertiesService.getDocumentProperties();
    var sheetName = SpreadsheetApp.getActiveSheet().getName();

    // Signal includes the active sheet name so ANY tab move/return changes
    // the value and the sidebar immediately reloads that tab's own templates.
    var key = 'SIDEBAR_REFRESH_SIGNAL_KEY_' + sheetName;
    return sheetName + "|" + (props.getProperty(key) || "");
  } catch (e) {
    console.log("Error reading layout sync properties: " + e.message);
    return "";
  }
}

function SIGNAL_SIDEBAR_REFRESH() {
  try {
    var sheetName = SpreadsheetApp.getActiveSheet().getName();
    var key = 'SIDEBAR_REFRESH_SIGNAL_KEY_' + sheetName;
    PropertiesService.getDocumentProperties().setProperty(key, "REFRESH_" + new Date().getTime());
    return true;
  } catch (e) {
    return false;
  }
}