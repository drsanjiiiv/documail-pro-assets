/**
 * ============================================================================
 * DOCUMAIL PRO COMPLETE MASTER CORE SCRIPT
 * ============================================================================
 * FILE: templates.gs - Template CRUD Operations
 * ============================================================================
 */

// ==========================================
// BLOCK: HELPER - GET STORAGE KEY PER TAB
// ==========================================

function GET_TEMPLATES_KEY(optSheetName) {
  var sheetName = optSheetName || SpreadsheetApp.getActiveSheet().getName();
  return 'documail_templates_' + sheetName;
}

function GET_REFRESH_KEY() {
  var sheetName = SpreadsheetApp.getActiveSheet().getName();
  return 'SIDEBAR_REFRESH_SIGNAL_KEY_' + sheetName;
}

// ==========================================
// BLOCK: TEMPLATE SYSTEM - CRUD OPERATIONS
// ==========================================

function GET_ALL_TEMPLATES(optSheetName) {
  try {
    var key = GET_TEMPLATES_KEY(optSheetName);
    var templates = PropertiesService.getDocumentProperties().getProperty(key);
    return templates ? JSON.parse(templates) : [];
  } catch (e) {
    console.error("GET_ALL_TEMPLATES error: " + e.message);
    return [];
  }
}

function GET_SIDEBAR_TEMPLATES() {
  try {
    // The sidebar always reflects the currently-ACTIVE sheet so each tab
    // shows its own templates ("No templates" when the tab has none).
    return GET_ALL_TEMPLATES(SpreadsheetApp.getActiveSheet().getName());
  } catch (e) {
    console.error("GET_SIDEBAR_TEMPLATES error: " + e.message);
    return [];
  }
}

function SAVE_TEMPLATE(template) {
  var key = GET_TEMPLATES_KEY();
  var templates = GET_ALL_TEMPLATES();
  var existingIndex = -1;
  var isNewTemplate = false;
  
  for (var i = 0; i < templates.length; i++) {
    if (templates[i].id === template.id) {
      existingIndex = i;
      break;
    }
  }
  
  if (existingIndex !== -1) {
    templates[existingIndex] = template;
  } else {
    template.id = generateUUID();
    templates.push(template);
    isNewTemplate = true;
  }
  
  PropertiesService.getDocumentProperties().setProperty(key, JSON.stringify(templates));

  // Check/create status column
  var columnCreated = false;
  var columnExists = false;

  if (template.type === "EMAIL_ONLY" || template.type === "BOTH") {
    try {
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
      var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      var statusColumnName = "Sent Mail Status - " + template.name;
      var found = false;

      for (var i = 0; i < headers.length; i++) {
        if (String(headers[i]).toLowerCase().trim() === statusColumnName.toLowerCase()) {
          columnExists = true;
          found = true;
          break;
        }
      }

      if (!found) {
        GET_OR_CREATE_STATUS_COLUMN(sheet, template.name);
        columnCreated = true;
      }
    } catch (e) {
      console.log("Could not create status column: " + e.message);
    }
  }

  // Return both template and column status
  return {
    template: template,
    columnCreated: columnCreated,
    columnExists: columnExists,
    isNewTemplate: isNewTemplate
  };
}

function DELETE_TEMPLATE(templateId) {
  var key = GET_TEMPLATES_KEY();
  var templates = GET_ALL_TEMPLATES();
  var targetTemplate = null;

  for (var i = 0; i < templates.length; i++) {
    if (templates[i].id === templateId) {
      targetTemplate = templates[i];
      break;
    }
  }

  if (!targetTemplate) {
    throw new Error("Template not found.");
  }

  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastColumn = sheet.getLastColumn();
  var headers = [];
  var emailColumnIndex = -1;
  var hasEmailData = false;
  var hasDocData = false;
  var docStatusColName = "";

  if (lastColumn > 0) {
    headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  }

  // =======================================================
  // CHECK EMAIL STATUS COLUMN (For EMAIL_ONLY and BOTH)
  // =======================================================
  if (targetTemplate.type === "EMAIL_ONLY" || targetTemplate.type === "BOTH") {
    var statusColumnName = ("Sent Mail Status - " + targetTemplate.name).toLowerCase().trim();
    
    for (var c = 0; c < headers.length; c++) {
      if (String(headers[c]).toLowerCase().trim() === statusColumnName) {
        emailColumnIndex = c;
        break;
      }
    }

    if (emailColumnIndex !== -1) {
      var lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        var dataRange = sheet.getRange(2, emailColumnIndex + 1, lastRow - 1, 1).getValues();
        for (var r = 0; r < dataRange.length; r++) {
          if (String(dataRange[r][0]).trim() !== "") {
            hasEmailData = true;
            break;
          }
        }
      }
    }
  }

  // =======================================================
  // CHECK MERGED DOC STATUS COLUMN (For PDF_ONLY and BOTH)
  // =======================================================
  if (targetTemplate.type === "PDF_ONLY" || targetTemplate.type === "BOTH") {
    for (var c = 0; c < headers.length; c++) {
      if (String(headers[c]).toLowerCase().indexOf("merged doc status") !== -1) {
        docStatusColName = headers[c];
        var lastRow = sheet.getLastRow();
        if (lastRow > 1) {
          var dataRange = sheet.getRange(2, c + 1, lastRow - 1, 1).getValues();
          for (var r = 0; r < dataRange.length; r++) {
            if (String(dataRange[r][0]).trim() !== "") {
              hasDocData = true;
              break;
            }
          }
        }
        break;
      }
    }
  }

  // =======================================================
  // IF EMAIL DATA EXISTS - BLOCK DELETION
  // =======================================================
  if (hasEmailData) {
    ui.alert(
      "❌ Cannot Delete Template",
      "This template has existing email status data in the column:\n\n" +
      "📊 " + headers[emailColumnIndex] + "\n\n" +
      "Please clear the data in this column first before deleting this template.",
      ui.ButtonSet.OK
    );
    return { success: false, message: "Template has email data. Cannot delete." };
  }

  // =======================================================
  // IF DOC DATA EXISTS - BLOCK DELETION
  // =======================================================
  if (hasDocData && targetTemplate.type !== "BOTH") {
    ui.alert(
      "❌ Cannot Delete Template",
      "This template has existing merged document records in the column:\n\n" +
      "📊 " + docStatusColName + "\n\n" +
      "Please clear the data in this column first before deleting this template.",
      ui.ButtonSet.OK
    );
    return { success: false, message: "Template has merged document data. Cannot delete." };
  }

  // =======================================================
  // ASK CONFIRMATION
  // =======================================================
  var confirmMessage = "";
  var deleteEmailColumn = false;

  if (targetTemplate.type === "PDF_ONLY") {
    confirmMessage = "No columns will be deleted.\n\n";
  } else if (emailColumnIndex !== -1) {
    confirmMessage = "The email status column will be deleted:\n" +
      "📊 " + headers[emailColumnIndex] + "\n\n";
    deleteEmailColumn = true;
  }

  confirmMessage += "Merge Doc columns of DocuMail Pro Engine will NOT be deleted.\n" +
    "Generated PDFs in Drive will NOT be deleted.\n\n" +
    "Are you sure you want to delete this template?";

  var response = ui.alert(
    "⚠️ Confirm Delete",
    confirmMessage,
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    return { success: false, message: "Deletion cancelled by user." };
  }

  // =======================================================
  // DELETE EMAIL STATUS COLUMN (ONLY IF EXISTS)
  // =======================================================
  if (deleteEmailColumn && emailColumnIndex !== -1) {
    try {
      var colIndex = emailColumnIndex + 1;
      
      // Remove protection
      var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
      var columnDescription = 'DocuMail Pro Status Column - ' + targetTemplate.name;
      for (var p = 0; p < protections.length; p++) {
        if (protections[p].getDescription() === columnDescription) {
          protections[p].remove();
          break;
        }
      }
      
      SpreadsheetApp.flush();
      sheet.deleteColumn(colIndex);
      console.log("🗑️ Email status column deleted: " + headers[emailColumnIndex]);
    } catch (e) {
      console.log("Could not delete email column: " + e.message);
    }
  }

  // =======================================================
  // REMOVE TIME-DRIVEN TRIGGER FOR THIS TEMPLATE
  // =======================================================
  try {
    REMOVE_TEMPLATE_TRIGGER(templateId);
  } catch (e) {
    console.log("Could not remove trigger: " + e.message);
  }

  // =======================================================
  // REMOVE TEMPLATE FROM STORAGE (PER TAB)
  // =======================================================
  var filtered = [];
  for (var i = 0; i < templates.length; i++) {
    if (templates[i].id !== templateId) {
      filtered.push(templates[i]);
    }
  }
  PropertiesService.getDocumentProperties().setProperty(key, JSON.stringify(filtered));

  // =======================================================
  // SIDEBAR REFRESH (PER TAB)
  // =======================================================
  try {
    var refreshKey = GET_REFRESH_KEY();
    PropertiesService.getDocumentProperties().setProperty(refreshKey, "REFRESH_" + new Date().getTime());
  } catch (err) {}

  ui.alert("✅ Template Deleted", "Template '" + targetTemplate.name + "' has been deleted successfully.", ui.ButtonSet.OK);

  return { success: true, message: "Template deleted successfully." };
}

function GET_TEMPLATE_BY_ID(templateId, optSheetName) {
  var templates = GET_ALL_TEMPLATES(optSheetName);

  // Debug: Log all template IDs
  var ids = [];
  for (var i = 0; i < templates.length; i++) {
    ids.push(templates[i].id);
  }
  Logger.log("Looking for ID: " + templateId);
  Logger.log("Available IDs: " + ids.join(", "));

  for (var i = 0; i < templates.length; i++) {
    if (templates[i].id === templateId) {
      return templates[i];
    }
  }
  return null;
}

function GET_TEMPLATE_BY_ID_FOR_EDIT(templateId, optSheetName) {
  return GET_TEMPLATE_BY_ID(templateId, optSheetName);
}

function SAVE_TEMPLATE_SCHEDULE(params) {
  var key = GET_TEMPLATES_KEY();
  var templates = GET_ALL_TEMPLATES();
  for (var i = 0; i < templates.length; i++) {
    if (templates[i].id === params.templateId) {
      templates[i].schedule = params.schedule;
      break;
    }
  }
  PropertiesService.getDocumentProperties().setProperty(key, JSON.stringify(templates));

  try {
    MANAGE_TRIGGER_FOR_TEMPLATE(params.templateId);
  } catch (e) {
    console.log("Trigger management error: " + e.message);
  }

  return "Schedule saved";
}

function generateUUID() {
  return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function (c) {
    var r = Math.random() * 16 | 0;
    var v = c == 'x' ? r : (r & 0x3 | 0x8);
    return v.toString(16);
  });
}

function GET_TEMPLATE_BY_NAME(templateName, optSheetName) {
  var templates = GET_ALL_TEMPLATES(optSheetName);
  for (var i = 0; i < templates.length; i++) {
    if (templates[i].name === templateName) {
      return templates[i];
    }
  }
  return null;
}

function DUPLICATE_TEMPLATE(templateId) {
  var original = GET_TEMPLATE_BY_ID(templateId);
  if (!original) throw new Error("Template not found");

  var copy = JSON.parse(JSON.stringify(original));
  copy.id = generateUUID();

  // Generate unique name
  var allTemplates = GET_ALL_TEMPLATES();
  var existingNames = {};
  for (var i = 0; i < allTemplates.length; i++) {
    existingNames[allTemplates[i].name.toLowerCase().trim()] = true;
  }

  var baseName = original.name + " (Copy)";
  var newName = baseName;
  var counter = 2;
  while (existingNames[newName.toLowerCase().trim()]) {
    newName = original.name + " (Copy " + counter + ")";
    counter++;
  }
  copy.name = newName;

  // Copy the Google Doc template if exists
  if (copy.config && copy.config.templateUrl) {
    try {
      var docId = copy.config.templateUrl.split("/d/")[1].split("/")[0];
      var newDocFile = Drive.Files.copy({ name: "Copy of " + copy.name + " template" }, docId);
      var newDocUrl = "https://docs.google.com/document/d/" + newDocFile.id + "/edit";
      copy.config.templateUrl = newDocUrl;
    } catch (e) {
      console.log("Could not copy document template: " + e.message);
    }
  }

  // Set schedule to MANUAL for the copy
  copy.schedule = "MANUAL";

  // Save the duplicated template
  SAVE_TEMPLATE(copy);

  // Trigger sidebar refresh
  try {
    var refreshKey = GET_REFRESH_KEY();
    PropertiesService.getDocumentProperties().setProperty(refreshKey, "REFRESH_" + new Date().getTime());
  } catch (err) {}

  return { success: true, newId: copy.id, newName: copy.name };
}

function SAVE_TEMPLATE_WITH_SCHEDULE(templateData, schedule, templateId, sendEmail) {
  var templates = GET_ALL_TEMPLATES();
  var existingIndex = -1;
  var template = templateData;
  var isNewTemplate = false;

  // 👇 SAFETY ADDITION: Fallback to internal object ID if parameter is missed
  var effectiveId = templateId || templateData.id || null;

  if (effectiveId) {
    template.id = effectiveId;
    for (var i = 0; i < templates.length; i++) {
      if (templates[i].id === effectiveId) {
        existingIndex = i;
        break;
      }
    }
  }

  // Handle status column rename if template name changed
  if (existingIndex !== -1) {
    var oldName = templates[existingIndex].name;
    if (oldName !== template.name) {
      try {
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
        var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        var oldStatusColName = "Sent Mail Status - " + oldName;
        for (var c = 0; c < headers.length; c++) {
          if (String(headers[c]).toLowerCase().trim() === oldStatusColName.toLowerCase()) {
            sheet.getRange(1, c + 1).setValue("Sent Mail Status - " + template.name);
            break;
          }
        }
      } catch (e) {
        console.log("Could not rename status column: " + e.message);
      }
    }
  }

  if (existingIndex !== -1) {
    templates[existingIndex] = template;
  } else {
    template.id = generateUUID();
    templates.push(template);
    isNewTemplate = true;
  }

  // Add schedule
  template.schedule = schedule;

  // Save to properties
  var key = GET_TEMPLATES_KEY();
  PropertiesService.getDocumentProperties().setProperty(key, JSON.stringify(templates));

  // Update time-driven trigger based on schedule
  try {
    MANAGE_TRIGGER_FOR_TEMPLATE(template.id);
  } catch (e) {
    console.log("Trigger management error: " + e.message);
  }

  // Check/create status column
  var columnCreated = false;
  var columnExists = false;

  if (template.type === "EMAIL_ONLY" || template.type === "BOTH") {
    try {
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
      var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      var statusColumnName = "Sent Mail Status - " + template.name;
      var found = false;

      for (var i = 0; i < headers.length; i++) {
        if (String(headers[i]).toLowerCase().trim() === statusColumnName.toLowerCase()) {
          columnExists = true;
          found = true;
          break;
        }
      }

      if (!found) {
        GET_OR_CREATE_STATUS_COLUMN(sheet, template.name);
        columnCreated = true;
      }
    } catch (e) {
      console.log("Could not create status column: " + e.message);
    }
  }

  // =======================================================
  // AUTO-REFRESH TOKEN
  // =======================================================
  try {
    var refreshKey = GET_REFRESH_KEY();
    PropertiesService.getDocumentProperties().setProperty(refreshKey, "REFRESH_" + new Date().getTime());
  } catch (err) {
    console.log("Error updating sidebar sync token: " + err.message);
  }

  // =======================================================
  // NOTIFICATION EMAIL - ONLY IF sendEmail IS TRUE
  // =======================================================
  if (sendEmail === true) {
    try {
      var userEmail = Session.getActiveUser().getEmail();
      var userName = Session.getActiveUser().getUsername() || userEmail;
      
      if (userEmail) {
        var templateName = template.name || "Unnamed Template";
        var typeDisplayMsg = "";
        if (template.type === "PDF_ONLY") {
          typeDisplayMsg = "📄 PDF Only";
        } else if (template.type === "EMAIL_ONLY") {
          typeDisplayMsg = "✉️ Email Only";
        } else if (template.type === "BOTH") {
          typeDisplayMsg = "🔄 PDF & Email";
        } else {
          typeDisplayMsg = template.type || "Standard Layout";
        }

        var runModeMsg = "";
        if (schedule === "MANUAL") {
          runModeMsg = 'Will run "Manually" only';
        } else {
          runModeMsg = 'Will run "Time Trigger" (Schedule: ' + schedule + ')';
        }

        var actionType = isNewTemplate ? "Created" : "Updated";

        var emailSubject = "📄 DocuMail Pro: Template [" + templateName + "] " + actionType + " Successfully";
        var emailBody = "Hello,\n\n" +
          "This is to confirm that a template configuration has been successfully " + actionType.toLowerCase() + ".\n\n" +
          "Details:\n" +
          "▪️ Template Name: " + templateName + "\n" +
          "▪️ Template Type: " + typeDisplayMsg + "\n" +
          "▪️ Execution Mode: " + runModeMsg + "\n" +
          "▪️ " + actionType + " By: " + userName + " (" + userEmail + ")\n" +
          "▪️ Date/Time: " + new Date().toLocaleString() + "\n\n" +
          "Regards,\n" +
          "DocuMail Pro System Engine";

        GmailApp.sendEmail(userEmail, emailSubject, emailBody);
        console.log("📨 " + actionType + " confirmation email sent to: " + userEmail);
      }
    } catch (err) {
      console.log("Could not send template configuration email: " + err.message);
    }
  } else {
    console.log("⏭️ Email skipped (auto-save or no sendEmail flag)");
  }

  return {
    template: template,
    columnCreated: columnCreated,
    columnExists: columnExists,
    isNewTemplate: isNewTemplate
  };
}