  /**
   * ============================================================================
   * DOCUMAIL PRO COMPLETE MASTER CORE SCRIPT
   * ============================================================================
   */

  // ==========================================
  // TEMPLATE SYSTEM - CRUD OPERATIONS
  // ==========================================

  function GET_ALL_TEMPLATES() {
    var templates = PropertiesService.getDocumentProperties().getProperty('documail_templates');
    return templates ? JSON.parse(templates) : [];
  }

  function SAVE_TEMPLATE(template) {
    var templates = GET_ALL_TEMPLATES();
    var existingIndex = -1;
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
    }
    PropertiesService.getDocumentProperties().setProperty('documail_templates', JSON.stringify(templates));
    return template;
  }

  function DELETE_TEMPLATE(templateId) {
    var templates = GET_ALL_TEMPLATES();
    var filtered = [];
    for (var i = 0; i < templates.length; i++) {
      if (templates[i].id !== templateId) {
        filtered.push(templates[i]);
      }
    }
    PropertiesService.getDocumentProperties().setProperty('documail_templates', JSON.stringify(filtered));
    return true;
  }

  function GET_TEMPLATE_BY_ID(templateId) {
    var templates = GET_ALL_TEMPLATES();
    for (var i = 0; i < templates.length; i++) {
      if (templates[i].id === templateId) {
        return templates[i];
      }
    }
    return null;
  }

  function GET_TEMPLATE_BY_ID_FOR_EDIT(templateId) {
    return GET_TEMPLATE_BY_ID(templateId);
  }

  function SAVE_TEMPLATE_SCHEDULE(params) {
    var templates = GET_ALL_TEMPLATES();
    for (var i = 0; i < templates.length; i++) {
      if (templates[i].id === params.templateId) {
        templates[i].schedule = params.schedule;
        break;
      }
    }
    PropertiesService.getDocumentProperties().setProperty('documail_templates', JSON.stringify(templates));
    return "Schedule saved";
  }

  function generateUUID() {
    return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function(c) {
      var r = Math.random() * 16 | 0;
      var v = c == 'x' ? r : (r & 0x3 | 0x8);
      return v.toString(16);
    });
  }

  // ==========================================
  // Function (onOpen) Starts
  // ==========================================
  function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('DocuMail Pro Platform')
      .addItem('Initialize Sheet Structural Layout', 'GENERATE_DOCUMAIL_TEMPLATE')
      .addItem('Open Engine Workspace Control Sidebar', 'INITIALIZE_ADDON_SIDEBAR')
      .addToUi();
  }

  // ==========================================
  // Function (INITIALIZE_ADDON_SIDEBAR) Starts
  // ==========================================
  function INITIALIZE_ADDON_SIDEBAR() {
    var html = HtmlService.createTemplateFromFile('SidebarView');
    var sidebarUi = html.evaluate().setTitle("DocuMail Pro").setSandboxMode(HtmlService.SandboxMode.IFRAME);
    SpreadsheetApp.getUi().showSidebar(sidebarUi);
  }

  // ==========================================
  // FUNCTION: GENERATE_DOCUMAIL_TEMPLATE Starts
  // ==========================================
  function GENERATE_DOCUMAIL_TEMPLATE() {
    var appName = "DocuMail Pro";
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();
    
    var existingProtections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    for (var i = 0; i < existingProtections.length; i++) {
      existingProtections[i].remove();
    }
    SpreadsheetApp.flush();
    
    sheet.clear();
    sheet.clearFormats();
    sheet.clearContents();
    sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).clearDataValidations();
    
    var headers = [
      "Name", "Company", "Designation", "Department",
      "Recipient Email",
      "Merged Doc Status - " + appName,
      "Merged Doc ID - " + appName,
      "Merged Doc URL - " + appName,
      "Sent Mail Status - " + appName
    ];
    
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    var headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight("bold")
      .setBackground("#E8F0FE")
      .setFontColor("#1A73E8")
      .setWrap(true)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");
    
    sheet.setFrozenRows(1);
    
    sheet.getRange("E1").setFontColor("#137333").setBackground("#E6F4EA");
    var emailRange = sheet.getRange("E2:E1000");
    emailRange.setDataValidation(SpreadsheetApp.newDataValidation().requireTextIsEmail().setAllowInvalid(false).build());
    
    var systemRange = sheet.getRange("F2:I1000");
    systemRange.setBackground("#f3f3f3");
    systemRange.setFontColor("#888888");
    
    sheet.getRange("F1").setNote("🔒 SYSTEM COLUMN - Auto-generated. Do not edit.");
    sheet.getRange("G1").setNote("🔒 SYSTEM COLUMN - Auto-generated Document ID");
    sheet.getRange("H1").setNote("🔒 SYSTEM COLUMN - Auto-generated Document URL");
    sheet.getRange("I1").setNote("🔒 SYSTEM COLUMN - Auto-generated Email Status");
    
    var protection = sheet.getRange("F1:I1000").protect();
    protection.setDescription('DocuMail Pro System Columns - Locked');
    protection.removeEditors(protection.getEditors());
    
    for (var col = 1; col <= headers.length; col++) {
      sheet.autoResizeColumn(col);
    }
    sheet.setRowHeight(1, 40);
    sheet.setColumnWidth(5, 200);
    sheet.setColumnWidth(8, 300);
    
    SpreadsheetApp.getUi().alert("✅ Sheet structure initialized successfully!");
  }

  // ==========================================
  // Helper Functions
  // ==========================================
  function GET_ALL_RAW_HEADERS() {
    try {
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
      var lastCol = sheet.getLastColumn();
      if (lastCol < 1) lastCol = 1;
      return sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    } catch (e) {
      return ["Name", "Company", "Designation", "Department", "Recipient Email"];
    }
  }

  function GET_LIVE_SHEET_HEADERS() {
    try {
      var rawHeaders = GET_ALL_RAW_HEADERS();
      var result = [];
      for (var i = 0; i < rawHeaders.length; i++) {
        var hStr = String(rawHeaders[i]).toLowerCase().trim();
        if (hStr === "") continue;
        if (hStr.indexOf("recipient email") !== -1) continue;
        if (hStr.indexOf("merged doc") !== -1) continue;
        if (hStr.indexOf("sent mail status") !== -1) continue;
        result.push(rawHeaders[i]);
      }
      return result;
    } catch (e) {
      return ["Name", "Company", "Designation", "Department"];
    }
  }

  function GET_WIZARD_INITIALIZATION_PAYLOAD() {
    return { userHeaders: GET_LIVE_SHEET_HEADERS() };
  }

  function GET_SYSTEM_OAUTH_TOKEN() {
    try {
      DriveApp.getRootFolder();
      return ScriptApp.getOAuthToken();
    } catch (e) {
      return "";
    }
  }

  function EXTRACT_TEMPLATE_TAGS_STREAM(docUrl) {
    try {
      var fileId = docUrl.split("/d/")[1].split("/")[0];
      var docInstance = DocumentApp.openById(fileId);
      var bodyText = docInstance.getBody().getText();
      var tagMatchRegex = /\{([^}]+)\}/g;
      var matches = [];
      var matchItem;
      while ((matchItem = tagMatchRegex.exec(bodyText)) !== null) {
        var cleanToken = matchItem[1].trim();
        if (matches.indexOf(cleanToken) === -1) {
          matches.push(cleanToken);
        }
      }
      return matches;
    } catch (err) {
      return [];
    }
  }

  function GET_ELIGIBLE_RECORDS_FOR_TEMPLATE(params) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var allHeaders = GET_ALL_RAW_HEADERS();
    var emailColIdx = -1;
    var mailStatusColIdx = -1;
    
    for (var c = 0; c < allHeaders.length; c++) {
      var hName = String(allHeaders[c]).toLowerCase().trim();
      if (hName.indexOf("recipient email") !== -1) emailColIdx = c;
      if (hName.indexOf("sent mail status") !== -1) mailStatusColIdx = c;
    }
    if (emailColIdx === -1) emailColIdx = 4;
    if (mailStatusColIdx === -1) mailStatusColIdx = 8;
    
    var criteriaColIdx = -1;
    for (var c = 0; c < allHeaders.length; c++) {
      if (allHeaders[c] === params.condField) {
        criteriaColIdx = c;
        break;
      }
    }
    if (criteriaColIdx === -1) criteriaColIdx = 0;
    
    var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, allHeaders.length).getValues();
    var eligible = [];
    
    for (var i = 0; i < values.length; i++) {
      var rowNum = i + 2;
      var mailStatus = String(values[i][mailStatusColIdx] || "").trim();
      if (params.type === "EMAIL_ONLY" && mailStatus !== "") continue;
      
      var cell = String(values[i][criteriaColIdx] || "").trim();
      var match = false;
      if (params.condOperator === "NOT_EMPTY") {
        match = (cell !== "");
      } else if (params.condOperator === "CONTAINS") {
        match = (cell.toLowerCase().indexOf(params.condValue.toLowerCase()) !== -1);
      }
      if (match) {
        eligible.push({ 
          rowNum: rowNum, 
          email: values[i][emailColIdx] || "", 
          refData: values[i][0] || "" 
        });
      }
    }
    return { eligibleRows: eligible };
  }

  function executeEmailSend(selectedRows, emailConfig) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var allHeaders = GET_ALL_RAW_HEADERS();
    var emailColIdx = -1;
    var mailStatusColIdx = -1;
    
    for (var c = 0; c < allHeaders.length; c++) {
      var hName = String(allHeaders[c]).toLowerCase().trim();
      if (hName.indexOf("recipient email") !== -1) emailColIdx = c;
      if (hName.indexOf("sent mail status") !== -1) mailStatusColIdx = c;
    }
    if (emailColIdx === -1) emailColIdx = 4;
    if (mailStatusColIdx === -1) mailStatusColIdx = 8;
    
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, allHeaders.length).getValues();
    var sent = 0;
    var errors = [];
    
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
      
      var recipient = data[i][emailColIdx];
      if (!recipient || recipient.indexOf("@") === -1) continue;
      
      var subject = emailConfig.subject || "";
      var body = emailConfig.body || "";
      
      for (var h = 0; h < allHeaders.length; h++) {
        var header = allHeaders[h];
        if (header) {
          var val = String(data[i][h] || "");
          var regex = new RegExp("\\{" + escapeRegex(header) + "\\}", "g");
          subject = subject.replace(regex, val);
          body = body.replace(regex, val);
        }
      }
      
      try {
        GmailApp.sendEmail(recipient, subject, "", { htmlBody: body });
        var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
        sheet.getRange(rowNum, mailStatusColIdx + 1).setValue("Sent to " + recipient + " on " + new Date());
        sent++;
      } catch (e) {
        errors.push("Row " + rowNum + ": " + e.message);
        sheet.getRange(rowNum, mailStatusColIdx + 1).setValue("Failed: " + e.message);
      }
    }
    
    if (errors.length > 0) {
      return "✅ Sent " + sent + " emails.\n⚠️ Errors: " + errors.join(", ");
    }
    return "✅ Successfully sent " + sent + " emails!";
  }

  function EXECUTE_TEMPLATE_ACTION(params) {
    if (params.type === "EMAIL_ONLY") {
      return executeEmailSend(params.selectedRows, params.emailConfig);
    }
    return "PDF generation coming soon";
  }

  function RUN_TEMPLATE(templateId) {
    var template = GET_TEMPLATE_BY_ID(templateId);
    if (!template) return "Template not found";
    
    if (template.type === "PDF_ONLY") {
      var config = JSON.parse(JSON.stringify(template.config));
      config.isPreview = false;
      var result = EXECUTE_DOCUMENT_MERGE_ENGINE(config);
      return result;
    }
    
    if (template.type === "EMAIL_ONLY") {
      var allHeaders = GET_ALL_RAW_HEADERS();
      var emailColIdx = -1;
      var mailStatusColIdx = -1;
      for (var c = 0; c < allHeaders.length; c++) {
        var hName = String(allHeaders[c]).toLowerCase().trim();
        if (hName.indexOf("recipient email") !== -1) emailColIdx = c;
        if (hName.indexOf("sent mail status") !== -1) mailStatusColIdx = c;
      }
      if (emailColIdx === -1) emailColIdx = 4;
      if (mailStatusColIdx === -1) mailStatusColIdx = 8;
      
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
      var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, allHeaders.length).getValues();
      var selectedRows = [];
      
      for (var i = 0; i < data.length; i++) {
        var mailStatus = String(data[i][mailStatusColIdx] || "").trim();
        if (mailStatus === "") {
          selectedRows.push(i + 2);
        }
      }
      
      if (selectedRows.length === 0) return "No eligible records found";
      return executeEmailSend(selectedRows, template.emailConfig);
    }
    
    if (template.type === "BOTH") {
      // First generate PDF (this will also update columns F, G, H)
      var config = JSON.parse(JSON.stringify(template.config));
      config.isPreview = false;
      var docResult = EXECUTE_DOCUMENT_MERGE_ENGINE(config);
      
      // Get the folder where PDF was saved
      var folderId = config.folderDestination;
      var folderIdExtracted = folderId ? (folderId.split("/folders/")[1] || folderId.split("id=")[1]) : null;
      
      // Then send emails with PDF attachments
      var emailResult = executeEmailSendWithAttachments(template.emailConfig, folderIdExtracted);
      
      return docResult + "\n\n" + emailResult;
    }
    
    return "Unknown template type";
  }

  // ==========================================
  // UI Functions
  // ==========================================
  function OPEN_DOCUMENT_WIZARD_MODAL(templateId) {
    var htmlLayout = HtmlService.createTemplateFromFile('DocConfigView');
    
    // CRITICAL FIX: Pass templateId to the HTML template
    if (templateId) {
      htmlLayout.existingTemplateId = templateId;
    } else {
      htmlLayout.existingTemplateId = null;
    }
    
    var dialogWindow = htmlLayout.evaluate().setWidth(960).setHeight(680);
    SpreadsheetApp.getUi().showModalDialog(dialogWindow, templateId ? "Edit Template" : "Create New Template");
  }

  function OPEN_CENTER_TEMPLATE_MODAL() {
    var htmlLayout = HtmlService.createTemplateFromFile('EmailComposerView');
    var dialogWindow = htmlLayout.evaluate().setWidth(980).setHeight(650);
    SpreadsheetApp.getUi().showModalDialog(dialogWindow, "Build & Automate Email Templates");
  }

  function OPEN_RECORDS_PREVIEW_WINDOW() {
    var htmlLayout = HtmlService.createTemplateFromFile('RecordsPreviewView');
    var dialogWindow = htmlLayout.evaluate().setWidth(1000).setHeight(700);
    SpreadsheetApp.getUi().showModalDialog(dialogWindow, "Confirm Send Records & Daily Limits Check");
  }

  function UPLOAD_LOCAL_FILE_STREAM(fileObject) {
    try {
      var folderName = "DocuMail Pro Local Uploads";
      var folders = DriveApp.getFoldersByName(folderName);
      var targetFolder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
      var contentSplit = fileObject.data.split(',');
      var base64Data = contentSplit.length > 1 ? contentSplit[1] : contentSplit[0];
      var fileBlob = Utilities.newBlob(Utilities.base64Decode(base64Data), fileObject.mimeType, fileObject.name);
      var driveFile = targetFolder.createFile(fileBlob);
      return { success: true, id: driveFile.getId(), name: driveFile.getName() };
    } catch (err) {
      return { success: false, error: err.message };
    }
  }

  function SAVE_TEMPLATE_STATE(configPayload) {
    CacheService.getDocumentCache().put('email_config_state', JSON.stringify(configPayload), 21600);
  }

  function GET_STORED_TEMPLATE_STATE() {
    var savedState = CacheService.getDocumentCache().get('email_config_state');
    return savedState ? JSON.parse(savedState) : null;
  }

  function SAVE_DOC_WIZARD_STATE(configPayload) {
    CacheService.getDocumentCache().put('doc_config_state_memory', JSON.stringify(configPayload), 21600);
  }

  function GET_SAVED_DOC_WIZARD_STATE() {
    var cachePayload = CacheService.getDocumentCache().get('doc_config_state_memory');
    return cachePayload ? JSON.parse(cachePayload) : null;
  }

  function RECEPTACLE_SIDEBAR_STATE_PASS(sidebarOptionsPayload) {
    CacheService.getDocumentCache().put('sidebar_channel_options', JSON.stringify(sidebarOptionsPayload), 21600);
  }

  // ==========================================
  // PREVIEW TEMPLATE FUNCTION - Shows both PDF and Email preview for BOTH type
  // ==========================================
  function PREVIEW_TEMPLATE(templateId) {
    var template = GET_TEMPLATE_BY_ID(templateId);
    if (!template) return { name: "Not Found", previewHtml: "<p>Template not found</p>" };
    
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var allHeaders = GET_ALL_RAW_HEADERS();
    
    // Get first eligible row for preview
    var criteriaColIdx = allHeaders.indexOf(template.config?.condField);
    if (criteriaColIdx === -1) criteriaColIdx = 0;
    
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, allHeaders.length).getValues();
    var previewRow = null;
    
    // Find first row that matches filter condition
    for (var i = 0; i < data.length; i++) {
      var evalCellText = String(data[i][criteriaColIdx] || "").trim();
      var isMatch = false;
      if (template.config?.condOperator === "NOT_EMPTY" && evalCellText !== "") {
        isMatch = true;
      } else if (template.config?.condOperator === "CONTAINS" && evalCellText.toLowerCase().indexOf(String(template.config?.condValue).toLowerCase().trim()) !== -1) {
        isMatch = true;
      }
      if (isMatch) {
        previewRow = data[i];
        break;
      }
    }
    
    if (!previewRow) {
      return { name: template.name, previewHtml: "<p>No eligible rows found for preview</p>" };
    }
    
    var previewHtml = '<div style="font-family: Roboto, sans-serif; padding: 20px;">';
    previewHtml += '<h2 style="color:#1a73e8;">✅ Preview Generated Successfully!</h2>';
    previewHtml += '<hr>';
    previewHtml += '<p><strong>📄 Template Name:</strong> ' + escapeHtml(template.name) + '</p>';
    previewHtml += '<p><strong>📋 Template Type:</strong> ' + (template.type === "PDF_ONLY" ? "PDF Only" : template.type === "EMAIL_ONLY" ? "Email Only" : "PDF & Email") + '</p>';
    
    // PDF Section (for PDF_ONLY or BOTH)
    if (template.type === "PDF_ONLY" || template.type === "BOTH") {
      var previewConfig = JSON.parse(JSON.stringify(template.config));
      previewConfig.isPreview = true;
      
      var fileName = previewConfig.namePattern || "Document";
      for (var h = 0; h < allHeaders.length; h++) {
        var header = allHeaders[h];
        if (header) {
          var regex = new RegExp("\\{" + escapeRegex(header) + "\\}", "g");
          fileName = fileName.replace(regex, String(previewRow[h] || ""));
        }
      }
      fileName = fileName.replace(/[\\/:*?"<>|]/g, "");
      var expectedFileName = "PROV-" + fileName + ".pdf";
      
      var result = EXECUTE_DOCUMENT_MERGE_ENGINE(previewConfig);
      
      var fileUrl = "";
      var folderId = template.config?.folderDestination;
      if (folderId) {
        var folderIdExtracted = folderId.split("/folders/")[1] || folderId.split("id=")[1];
        if (folderIdExtracted) {
          var folder = DriveApp.getFolderById(folderIdExtracted);
          var files = folder.getFiles();
          var latestFile = null;
          var latestTime = 0;
          while (files.hasNext()) {
            var file = files.next();
            var fileCreated = file.getDateCreated().getTime();
            if (file.getName().indexOf("PROV-") === 0 && fileCreated > latestTime) {
              latestFile = file;
              latestTime = fileCreated;
            }
          }
          if (latestFile) {
            fileUrl = latestFile.getUrl();
          }
        }
      }
      
      previewHtml += '<div style="background:#e8f0fe; padding:12px; border-radius:6px; margin:10px 0;">';
      previewHtml += '<h3 style="color:#1a73e8; margin:0 0 10px 0;">📄 Document Preview</h3>';
      previewHtml += '<p><strong>Document Name:</strong> ' + escapeHtml(expectedFileName) + '</p>';
      previewHtml += '<p><strong>Output Format:</strong> ' + (template.config?.format === "PDF" ? "PDF Document" : "Editable Google Doc") + '</p>';
      if (fileUrl) {
        previewHtml += '<p><strong>File Link:</strong> <a href="' + fileUrl + '" target="_blank" style="color:#1a73e8;">Click here to open</a></p>';
      }
      previewHtml += '</div>';
    }
    
    // Email Section (for EMAIL_ONLY or BOTH)
    if (template.type === "EMAIL_ONLY" || template.type === "BOTH") {
      var emailColIdx = -1;
      for (var c = 0; c < allHeaders.length; c++) {
        var hName = String(allHeaders[c]).toLowerCase().trim();
        if (hName.indexOf("recipient email") !== -1) emailColIdx = c;
      }
      if (emailColIdx === -1) emailColIdx = 4;
      
      var recipientEmail = previewRow[emailColIdx] || "example@email.com";
      var emailSubject = template.emailConfig?.subject || "";
      var emailBody = template.emailConfig?.body || "";
      
      for (var h = 0; h < allHeaders.length; h++) {
        var header = allHeaders[h];
        if (header) {
          var val = String(previewRow[h] || "");
          var regex = new RegExp("\\{" + escapeRegex(header) + "\\}", "g");
          emailSubject = emailSubject.replace(regex, val);
          emailBody = emailBody.replace(regex, val);
        }
      }
      
      previewHtml += '<div style="background:#e6f4ea; padding:12px; border-radius:6px; margin:10px 0;">';
      previewHtml += '<h3 style="color:#137333; margin:0 0 10px 0;">✉️ Email Preview</h3>';
      previewHtml += '<p><strong>To:</strong> ' + escapeHtml(recipientEmail) + '</p>';
      previewHtml += '<p><strong>Subject:</strong> ' + escapeHtml(emailSubject) + '</p>';
      previewHtml += '<div style="border:1px solid #dadce0; padding:12px; border-radius:6px; margin-top:12px; background:#ffffff;">';
      previewHtml += '<strong>Email Body:</strong><br><br>' + emailBody;
      previewHtml += '</div>';
      previewHtml += '<p style="margin-top:12px; color:#5f6368;"><em>Note: This is a preview only. No email was sent.</em></p>';
      previewHtml += '</div>';
    }
    
    previewHtml += '<hr>';
    previewHtml += '<p style="color:#5f6368; font-size:12px;"><em>Note: This is a preview. No spreadsheet data was updated.</em></p>';
    previewHtml += '<div style="margin-top:20px; text-align:center;">';
    previewHtml += '<button onclick="google.script.host.close()" style="background:#1a73e8; color:white; border:none; padding:8px 24px; border-radius:4px; cursor:pointer;">Close</button>';
    previewHtml += '</div>';
    previewHtml += '</div>';
    
    return { name: template.name, previewHtml: previewHtml };
  }

  function escapeHtml(str) {
    if (!str) return '';
    return String(str).replace(/[&<>]/g, function(m) {
      if (m === '&') return '&amp;';
      if (m === '<') return '&lt;';
      if (m === '>') return '&gt;';
      return m;
    });
  }

  function escapeRegex(str) {
    return String(str).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }

  function SIGNAL_SIDEBAR_REFRESH() {
    // Store a timestamp in cache that sidebar can check
    CacheService.getScriptCache().put('sidebar_refresh_signal', Date.now().toString(), 60);
  }

  function CHECK_SIDEBAR_REFRESH() {
    return CacheService.getScriptCache().get('sidebar_refresh_signal');
  }

  // ==========================================
  // FUNCTION: EXECUTE_DOCUMENT_MERGE_ENGINE Starts
  // ==========================================
  function EXECUTE_DOCUMENT_MERGE_ENGINE(payload) {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getActiveSheet();
      var allHeaders = GET_ALL_RAW_HEADERS();
      
      // Find column indices
      var docStatusColIdx = -1;
      var docIdColIdx = -1;
      var docUrlColIdx = -1;
      
      for (var c = 0; c < allHeaders.length; c++) {
        var hName = String(allHeaders[c]).toLowerCase();
        if (hName.indexOf("merged doc status") !== -1) docStatusColIdx = c;
        if (hName.indexOf("merged doc id") !== -1) docIdColIdx = c;
        if (hName.indexOf("merged doc url") !== -1) docUrlColIdx = c;
      }
      
      if (docStatusColIdx === -1) docStatusColIdx = 5;
      if (docIdColIdx === -1) docIdColIdx = 6;
      if (docUrlColIdx === -1) docUrlColIdx = 7;
      
      // TEMPORARILY REMOVE PROTECTION TO ALLOW WRITING (for non-preview)
      var tempRemovedProtections = [];
      if (!payload.isPreview) {
        var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
        for (var p = 0; p < protections.length; p++) {
          var protection = protections[p];
          var range = protection.getRange();
          var colStart = range.getColumn();
          var colEnd = colStart + range.getWidth() - 1;
          if (colStart <= docUrlColIdx + 1 && colEnd >= docStatusColIdx + 1) {
            tempRemovedProtections.push({
              protection: protection,
              rangeA1: range.getA1Notation()
            });
            protection.remove();
          }
        }
      }
      
      var lastRow = sheet.getLastRow();
      if (lastRow < 2) return "No data rows found.";
      
      var dataValues = sheet.getRange(2, 1, lastRow - 1, allHeaders.length).getValues();
      var processedCount = 0;
      
      // Find criteria column
      var criteriaColIdx = allHeaders.indexOf(payload.condField);
      if (criteriaColIdx === -1) criteriaColIdx = 0;
      
      // Get template file ID
      var templateUrl = payload.templateUrl;
      var templateId = templateUrl.split("/d/")[1].split("/")[0];
      var templateFile = DriveApp.getFileById(templateId);
      
      // Get destination folder
      var folderId = "";
      if (payload.folderDestination) {
        folderId = payload.folderDestination.split("/folders/")[1] || payload.folderDestination.split("id=")[1];
        if (!folderId) folderId = payload.folderDestination;
      }
      var destinationFolder = folderId ? DriveApp.getFolderById(folderId) : DriveApp.getRootFolder();
      
      // Process each row
      for (var i = 0; i < dataValues.length; i++) {
        var rowNum = i + 2;
        var rowData = dataValues[i];
        
        // For preview mode, process ONLY the first eligible row
        // For production, process all eligible rows
        
        // Check filter condition
        var evalCellText = String(rowData[criteriaColIdx] || "").trim();
        var isTargetEligible = false;
        
        if (payload.condOperator === "NOT_EMPTY" && evalCellText !== "") {
          isTargetEligible = true;
        } else if (payload.condOperator === "CONTAINS" && evalCellText.toLowerCase().indexOf(String(payload.condValue).toLowerCase().trim()) !== -1) {
          isTargetEligible = true;
        }
        
        if (!isTargetEligible) continue;
        
        // For preview mode, only process one row
        if (payload.isPreview && processedCount >= 1) break;
        
        // Create filename from pattern
        var fileName = payload.namePattern || "Generated Document";
        for (var h = 0; h < allHeaders.length; h++) {
          var header = allHeaders[h];
          if (header) {
            var regex = new RegExp("\\{" + escapeRegex(header) + "\\}", "g");
            fileName = fileName.replace(regex, String(rowData[h] || ""));
          }
        }
        
        fileName = fileName.replace(/[\\/:*?"<>|]/g, "");
        
        if (payload.isPreview) {
          fileName = "PROV-" + fileName;
        }
        
        // Copy the template file
        var newDocFile = templateFile.makeCopy(fileName, destinationFolder);
        var newDocId = newDocFile.getId();
        
        // Open and replace tags
        var newDoc = DocumentApp.openById(newDocId);
        var body = newDoc.getBody();
        
        // Replace tags using mapped values
        if (payload.tagMappings) {
          for (var docTag in payload.tagMappings) {
            if (payload.tagMappings.hasOwnProperty(docTag)) {
              var sheetColumn = payload.tagMappings[docTag];
              var replacementValue = "";
              var mappedColIdx = allHeaders.indexOf(sheetColumn);
              if (mappedColIdx !== -1) {
                replacementValue = String(rowData[mappedColIdx] || "");
              }
              
              // Search for the exact tag pattern
              var searchPattern = "{" + docTag + "}";
              var foundElement = body.findText(searchPattern);
              
              while (foundElement) {
                var element = foundElement.getElement();
                if (element.editAsText) {
                  var startOffset = foundElement.getStartOffset();
                  var endOffsetInclusive = foundElement.getEndOffsetInclusive();
                  element.asText().deleteText(startOffset, endOffsetInclusive);
                  element.asText().insertText(startOffset, replacementValue);
                }
                foundElement = body.findText(searchPattern, foundElement);
              }
            }
          }
        }
        
        newDoc.saveAndClose();
        
        var finalFileId = newDocId;
        var finalFileUrl = newDocFile.getUrl();
        
        // Convert to PDF if format is PDF
        if (payload.format === "PDF") {
          var pdfBlob = newDocFile.getAs('application/pdf');
          var pdfFileName = fileName + ".pdf";
          newDocFile.setTrashed(true);
          var pdfFile = destinationFolder.createFile(pdfBlob).setName(pdfFileName);
          finalFileId = pdfFile.getId();
          finalFileUrl = pdfFile.getUrl();
        }
        
        // Update sheet (only for non-preview)
        if (!payload.isPreview) {
          sheet.getRange(rowNum, docStatusColIdx + 1).setValue("Success");
          sheet.getRange(rowNum, docIdColIdx + 1).setValue(finalFileId);
          sheet.getRange(rowNum, docUrlColIdx + 1).setValue(finalFileUrl);
        }
        
        processedCount++;
      }
      
      // RESTORE PROTECTIONS (for non-preview)
      if (!payload.isPreview) {
        for (var r = 0; r < tempRemovedProtections.length; r++) {
          var prot = tempRemovedProtections[r];
          var newProtection = sheet.getRange(prot.rangeA1).protect();
          newProtection.setDescription(prot.protection.getDescription());
          var editors = prot.protection.getEditors();
          newProtection.addEditors(editors);
          if (prot.protection.canDomainEdit()) {
            newProtection.setDomainEdit(true);
          }
        }
      }
      
      var mode = payload.isPreview ? "PREVIEW" : "PRODUCTION";
      return mode + " completed! Processed " + processedCount + " record(s). " + (payload.format === "PDF" ? "PDF files" : "Google Docs") + " saved to: " + destinationFolder.getName();
      
    } catch (err) {
      return "Error: " + err.toString();
    }
  }
  // ==========================================
  // FUNCTION: EXECUTE_DOCUMENT_MERGE_ENGINE Ends
  // ==========================================

  function SHOW_PREVIEW_DIALOG(templateId) {
    var result = PREVIEW_TEMPLATE(templateId);
    var htmlOutput = HtmlService.createHtmlOutput(result.previewHtml)
      .setWidth(650)
      .setHeight(550)
      .setTitle('Preview: ' + result.name);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Preview: ' + result.name);
  }

  function SHOW_RUN_DIALOG(templateId) {
    var template = GET_TEMPLATE_BY_ID(templateId);
    var remainingQuota = MailApp.getRemainingDailyQuota();
    var totalQuota = 100; // Gmail daily quota is 100 for consumer accounts
    var usedQuota = totalQuota - remainingQuota;
    
    var result = RUN_TEMPLATE(templateId);
    
    // Get file URLs for PDF Only or BOTH templates
    var fileUrls = [];
    if (template.type === "PDF_ONLY" || template.type === "BOTH") {
      var folderId = template.config?.folderDestination;
      if (folderId) {
        var folderIdExtracted = folderId.split("/folders/")[1] || folderId.split("id=")[1];
        if (folderIdExtracted) {
          var folder = DriveApp.getFolderById(folderIdExtracted);
          var files = folder.getFiles();
          var latestFiles = [];
          var currentTime = new Date().getTime();
          
          while (files.hasNext()) {
            var file = files.next();
            var fileCreated = file.getDateCreated().getTime();
            if (fileCreated > currentTime - 120000) {
              latestFiles.push(file);
            }
          }
          latestFiles.sort(function(a, b) {
            return b.getDateCreated().getTime() - a.getDateCreated().getTime();
          });
          for (var i = 0; i < latestFiles.length; i++) {
            fileUrls.push({ name: latestFiles[i].getName(), url: latestFiles[i].getUrl() });
          }
        }
      }
    }
    
    var html = '<div style="font-family: Roboto, sans-serif; padding: 20px;">';
    html += '<h2 style="color:#1a73e8;">✅ Execution Completed!</h2>';
    html += '<hr>';
    html += '<p><strong>📄 Template:</strong> ' + escapeHtml(template.name) + '</p>';
    html += '<p><strong>📋 Type:</strong> ' + (template.type === "PDF_ONLY" ? "PDF Only" : template.type === "EMAIL_ONLY" ? "Email Only" : "PDF & Email") + '</p>';
    
    if (template.type === "EMAIL_ONLY" || template.type === "BOTH") {
      html += '<div style="background:#fce8e6; padding:10px; border-radius:6px; margin:10px 0;">';
      html += '<p style="margin:0;"><strong>📊 Daily Email Quota:</strong></p>';
      html += '<p style="margin:5px 0;"><strong>✅ Remaining:</strong> ' + remainingQuota + ' emails</p>';
      html += '<p style="margin:5px 0;"><strong>📤 Used Today:</strong> ' + usedQuota + ' emails</p>';
      html += '<p style="margin:5px 0; color:#5f6368; font-size:12px;">📅 Resets every 24 hours</p>';
      html += '</div>';
    }
    
    // Show file links for PDF Only or BOTH
    if (fileUrls.length > 0) {
      html += '<div style="background:#e8f0fe; padding:12px; border-radius:6px; margin:10px 0;">';
      html += '<h3 style="color:#1a73e8; margin:0 0 10px 0;">📄 Generated Files</h3>';
      for (var i = 0; i < fileUrls.length; i++) {
        html += '<p><strong>File ' + (i+1) + ':</strong> <a href="' + fileUrls[i].url + '" target="_blank" style="color:#1a73e8;">' + escapeHtml(fileUrls[i].name) + '</a></p>';
      }
      html += '</div>';
    }
    
    html += '<div style="background:#e6f4ea; padding:12px; border-radius:6px; margin:10px 0;">';
    html += '<strong>📝 Result:</strong><br>' + result.replace(/\n/g, '<br>');
    html += '</div>';
    
    html += '<hr>';
    html += '<div style="margin-top:20px; text-align:center;">';
    html += '<button onclick="google.script.host.close()" style="background:#1a73e8; color:white; border:none; padding:8px 24px; border-radius:4px; cursor:pointer;">Close</button>';
    html += '</div>';
    html += '</div>';
    
    var htmlOutput = HtmlService.createHtmlOutput(html)
      .setWidth(550)
      .setHeight(550)
      .setTitle('Execution Result: ' + template.name);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Execution Result');
  }


  function executeEmailSendWithAttachments(emailConfig, folderId) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var allHeaders = GET_ALL_RAW_HEADERS();
    
    var emailColIdx = -1;
    var mailStatusColIdx = -1;
    
    for (var c = 0; c < allHeaders.length; c++) {
      var hName = String(allHeaders[c]).toLowerCase().trim();
      if (hName.indexOf("recipient email") !== -1) emailColIdx = c;
      if (hName.indexOf("sent mail status") !== -1) mailStatusColIdx = c;
    }
    if (emailColIdx === -1) emailColIdx = 4;
    if (mailStatusColIdx === -1) mailStatusColIdx = 8;
    
    // Get all PDF files from the folder (generated in the last 2 minutes)
    var pdfFiles = [];
    if (folderId) {
      var folder = DriveApp.getFolderById(folderId);
      var files = folder.getFiles();
      var currentTime = new Date().getTime();
      while (files.hasNext()) {
        var file = files.next();
        var fileCreated = file.getDateCreated().getTime();
        if (file.getMimeType() === 'application/pdf' && fileCreated > currentTime - 120000) {
          pdfFiles.push(file);
        }
      }
    }
    
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, allHeaders.length).getValues();
    var sent = 0;
    var errors = [];
    
    for (var i = 0; i < data.length; i++) {
      var rowNum = i + 2;
      var existingStatus = String(data[i][mailStatusColIdx] || "").trim();
      
      if (existingStatus !== "") continue;
      
      var recipient = data[i][emailColIdx];
      if (!recipient || recipient.indexOf("@") === -1) continue;
      
      var subject = emailConfig.subject || "";
      var body = emailConfig.body || "";
      
      for (var h = 0; h < allHeaders.length; h++) {
        var header = allHeaders[h];
        if (header) {
          var val = String(data[i][h] || "");
          var regex = new RegExp("\\{" + escapeRegex(header) + "\\}", "g");
          subject = subject.replace(regex, val);
          body = body.replace(regex, val);
        }
      }
      
      var mailOptions = { htmlBody: body };
      if (emailConfig.replyTo) mailOptions.replyTo = emailConfig.replyTo;
      if (emailConfig.cc) mailOptions.cc = emailConfig.cc;
      if (emailConfig.bcc) mailOptions.bcc = emailConfig.bcc;
      
      // Attach PDF files
      var attachments = [];
      for (var p = 0; p < pdfFiles.length; p++) {
        attachments.push(pdfFiles[p].getBlob());
      }
      if (attachments.length > 0) {
        mailOptions.attachments = attachments;
      }
      
      try {
        GmailApp.sendEmail(recipient, subject, "", mailOptions);
        var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
        sheet.getRange(rowNum, mailStatusColIdx + 1).setValue("Sent to " + recipient + " on " + timestamp + " (with PDF)");
        sent++;
      } catch (e) {
        errors.push("Row " + rowNum + ": " + e.message);
        sheet.getRange(rowNum, mailStatusColIdx + 1).setValue("Failed: " + e.message);
      }
    }
    
    if (errors.length > 0) {
      return "✅ Sent " + sent + " emails (with PDF attachments).\n⚠️ Errors: " + errors.join(", ");
    }
    return "✅ Successfully sent " + sent + " emails (with PDF attachments)!";
  }