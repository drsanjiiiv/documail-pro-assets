/** file name: code.gs
/**
 * ============================================================================
 * DOCUMAIL PRO COMPLETE MASTER CORE SCRIPT
 * ============================================================================
 * FILE: code.gs - Main UI, Menu, Sidebar, and Dialog Functions
 * ============================================================================
 */

// ==========================================
// TEMPLATE CREATION CONFIGURATION
// ==========================================
// Dynamic template documents are created fresh (app-owned) per sheet via
// DocumentApp.create(). The linked spreadsheet ID is written into each
// document's footer (DOCUMAIL_SOURCE_SHEET_ID=...) so the Docs add-on can read
// the sheet headers as variables without any Drive API access. No master
// template, no bound scripts, no script binding steps are involved.

// ==========================================
// SPREADSHEET PLATFORM - ONOPEN (CLEAN)
// ==========================================
function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  
  ui.createAddonMenu()
    .addItem('📄 Initialize DocuMail Pro Sheet Structure', 'GENERATE_DOCUMAIL_TEMPLATE')
    .addItem('📝 Create Dynamic Doc Template From Sheet', 'CREATE_DOCUMENT_TEMPLATE_MENU')
    .addItem('🧠 Open Smart Template Engine', 'INITIALIZE_ADDON_SIDEBAR')
    .addSeparator()
    .addItem('❓ Help', 'showSheetHelp')
    .addToUi();
}

function onInstall(e) {
  onOpen(e);
}

// ==========================================
// AUTO-CLOSE SIDEBAR ON SHEET CHANGE
// ==========================================
function onSelectionChange(e) {
  var props = PropertiesService.getDocumentProperties();
  var anchorSheet = props.getProperty('SIDEBAR_ANCHOR_SHEET');
  if (!anchorSheet) return;

  var currentSheet = e.range.getSheet().getName();
  if (currentSheet !== anchorSheet) {
    props.deleteProperty('SIDEBAR_ANCHOR_SHEET');
    SpreadsheetApp.getUi().showSidebar(null);
  }
}

// ==========================================
// Function (INITIALIZE_ADDON_SIDEBAR) Starts
// ==========================================
function INITIALIZE_ADDON_SIDEBAR() {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSheet();

  // =======================================================
  // CHECK IF SYSTEM COLUMNS EXIST
  // =======================================================
  var lastCol = sheet.getLastColumn();
  var hasSystemColumns = false;

  if (lastCol > 0) {
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    for (var i = 0; i < headers.length; i++) {
      var hName = String(headers[i]).toLowerCase().trim();
      if (hName.indexOf("merged doc status") !== -1 ||
        hName.indexOf("recipient email") !== -1) {
        hasSystemColumns = true;
        break;
      }
    }
  }

  // =======================================================
  // IF SYSTEM COLUMNS DON'T EXIST, PROMPT USER
  // =======================================================
  if (!hasSystemColumns) {
    var response = ui.alert(
      "⚠️ DocuMail Pro Not Initialized",
      "DocuMail Pro system columns don't exist in this sheet.\n\n" +
      "Please initialize the sheet structure first to continue.\n\n" +
      "Do you want to initialize it now?",
      ui.ButtonSet.YES_NO
    );

    if (response === ui.Button.YES) {
      // Call the initialization function
      GENERATE_DOCUMAIL_TEMPLATE();
    }
    return;
  }

  // =======================================================
  // SYSTEM COLUMNS EXIST - PROCEED WITH SIDEBAR
  // =======================================================
  var html = HtmlService.createTemplateFromFile('SidebarView');
  var sidebarUi = html.evaluate()
    .setTitle("DocuMail Pro")
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showSidebar(sidebarUi);
  PropertiesService.getDocumentProperties().setProperty('SIDEBAR_ANCHOR_SHEET', sheet.getName());
}

  // ==========================================
// HELP FUNCTION FOR SHEETS MENU
// ==========================================
function showSheetHelp() {
  var ui = SpreadsheetApp.getUi();
  ui.alert(
    '📚 DocuMail Pro - Help',
    '📄 Initialize DocuMail Pro Sheet Structure\n' +
    'Adds all required system columns to your sheet for tracking:\n' +
    '• Merged Doc Status\n' +
    '• Merged Doc ID\n' +
    '• Merged Doc URL\n' +
    '• Recipient Email (if not exists)\n' +
    'These columns are essential for document generation & tracking.\n\n' +
    '─────────────────────────────\n\n' +
    '📝 Create Dynamic Doc Template From Sheet\n' +
    'Creates a new Google Doc template linked to your current sheet:\n' +
    '• Creates a fresh template document in your Drive\n' +
    '• Automatically links it to this sheet\n' +
    '• Opens the template for you to design\n' +
    '• Variables from your sheet headers can be inserted\n\n' +
    '─────────────────────────────\n\n' +
    '🧠 Open Smart Template Engine\n' +
    'Opens the main sidebar where you can:\n' +
    '• View all your templates\n' +
    '• Create new templates (PDF / Email / Both)\n' +
    '• Edit existing templates\n' +
    '• Preview templates with real data\n' +
    '• Run templates to generate documents & send emails\n\n' +
    '─────────────────────────────\n\n' +
    '💡 Pro Tip: Always initialize your sheet structure first\n' +
    'before creating templates for best results!',
    ui.ButtonSet.OK
  );
}

// ==========================================
// FUNCTION: CREATE_DOCUMENT_TEMPLATE_MENU
// Shows step-by-step animation with dummy progression
// ==========================================
function CREATE_DOCUMENT_TEMPLATE_MENU() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();
    var sheetName = sheet.getName();
    var sheetId = ss.getId();

    // Show the step-by-step animation dialog IMMEDIATELY
    var htmlContent = `
<!DOCTYPE html>
<html>
<head>
  <style>
    body {
      margin: 0;
      padding: 0;
      display: flex;
      justify-content: center;
      align-items: center;
      height: 100vh;
      background: #f8f9fa;
      font-family: 'Roboto', sans-serif;
    }
    .loader-container {
      text-align: left;
      padding: 40px 50px;
      background: white;
      border-radius: 12px;
      box-shadow: 0 2px 10px rgba(0,0,0,0.1);
      min-width: 420px;
    }
    .step {
      padding: 10px 0;
      border-bottom: 1px solid #f1f3f4;
      display: flex;
      align-items: center;
      font-size: 14px;
      color: #3c4043;
    }
    .step:last-child {
      border-bottom: none;
    }
    .step .icon {
      width: 28px;
      text-align: center;
      margin-right: 12px;
      font-size: 16px;
    }
    .step.done .icon {
      color: #137333;
    }
    .step.active .icon {
      color: #1a73e8;
    }
    .step.pending .icon {
      color: #dadce0;
    }
    .step.done .text {
      color: #137333;
    }
    .step.active .text {
      color: #1a73e8;
      font-weight: 500;
    }
    .step.pending .text {
      color: #9aa0a6;
    }
    .title {
      font-size: 16px;
      font-weight: 600;
      color: #1a73e8;
      margin-bottom: 16px;
      text-align: center;
    }
    .error-text {
      color: #c5221f;
      font-size: 14px;
      margin-top: 12px;
      text-align: center;
    }
  </style>
</head>
<body>
  <div class="loader-container" id="loader">
    <div class="title">📄 Creating Your Template</div>
    
    <div class="step active" id="step1">
      <span class="icon">⏳</span>
      <span class="text">📊 Fetching headers from your sheet...</span>
    </div>
    
    <div class="step pending" id="step2">
      <span class="icon">⏳</span>
      <span class="text">📝 Creating dynamic template document...</span>
    </div>
    
    <div class="step pending" id="step3">
      <span class="icon">⏳</span>
      <span class="text">🔗 Linking template to your sheet...</span>
    </div>
    
    <div class="step pending" id="step4">
      <span class="icon">⏳</span>
      <span class="text">📄 Creating your dynamic template document...</span>
    </div>
    
    <div class="step pending" id="step5">
      <span class="icon">⏳</span>
      <span class="text">⚙️ Configuring template settings...</span>
    </div>
    
    <div class="step pending" id="step6">
      <span class="icon">⏳</span>
      <span class="text">🚀 Launching template canvas...</span>
    </div>
  </div>
  
  <script>
    var steps = [
      { id: 'step1', done: false },
      { id: 'step2', done: false },
      { id: 'step3', done: false },
      { id: 'step4', done: false },
      { id: 'step5', done: false },
      { id: 'step6', done: false }
    ];
    
    function updateStep(stepIndex, status) {
      var step = document.getElementById(steps[stepIndex].id);
      var icon = step.querySelector('.icon');
      var text = step.querySelector('.text');
      
      step.className = 'step';
      if (status === 'done') {
        step.classList.add('done');
        icon.textContent = '✅';
        steps[stepIndex].done = true;
        if (stepIndex < steps.length - 1) {
          var nextStep = document.getElementById(steps[stepIndex + 1].id);
          nextStep.className = 'step active';
          nextStep.querySelector('.icon').textContent = '⏳';
        }
      } else if (status === 'active') {
        step.classList.add('active');
        icon.textContent = '⏳';
      } else {
        step.classList.add('pending');
        icon.textContent = '⏳';
      }
    }
    
    // DUMMY ANIMATION - Steps complete one by one
    setTimeout(function() { updateStep(0, 'done'); }, 300);
    setTimeout(function() { updateStep(1, 'done'); }, 600);
    setTimeout(function() { updateStep(2, 'done'); }, 900);
    setTimeout(function() { updateStep(3, 'done'); }, 1200);
    setTimeout(function() { updateStep(4, 'done'); }, 1500);
    
    // Open blank popup BEFORE async call (user gesture still active)
    var templateWindow = window.open('', '_blank');
    var popupBlocked = !templateWindow || templateWindow.closed;

    // Start the actual work in background
    google.script.run
      .withSuccessHandler(function(result) {
        if (result.success) {
          updateStep(5, 'done');

          if (!popupBlocked && templateWindow && !templateWindow.closed) {
            templateWindow.location.href = result.url;
          } else {
            document.getElementById('loader').innerHTML =
              '<div style="text-align:center;padding:20px;font-family:sans-serif;">' +
              '<h2 style="color:#137333;">✅ Template Created!</h2>' +
              '<p style="color:#5f6368;margin:16px 0;">Click to open your template:</p>' +
              '<a href="' + result.url + '" target="_blank" style="display:inline-block;padding:12px 32px;background:#1a73e8;color:#fff;text-decoration:none;border-radius:6px;font-size:15px;">📄 Open Template</a>' +
              '<p style="color:#9aa0a6;font-size:12px;margin-top:16px;">(Enable popups for auto-open next time)</p>' +
              '</div>';
            return;
          }

          setTimeout(function() {
            google.script.host.close();
          }, 600);
        } else {
          document.getElementById('loader').innerHTML = 
            '<div class="title" style="color:#c5221f;">❌ Error</div>' +
            '<div class="error-text">' + result.error + '</div>';
        }
      })
      .withFailureHandler(function(error) {
        document.getElementById('loader').innerHTML = 
          '<div class="title" style="color:#c5221f;">❌ Error</div>' +
          '<div class="error-text">' + error.message + '</div>';
      })
      .CREATE_TEMPLATE_IN_BACKGROUND("` + sheetName + `", "` + sheetId + `");
  </script>
</body>
</html>
`;

    var interfaceOutput = HtmlService.createHtmlOutput(htmlContent)
        .setWidth(520)
        .setHeight(400)
        .setTitle("Creating Template...");
        
    SpreadsheetApp.getUi().showModalDialog(interfaceOutput, "Creating Template...");

  } catch (e) {
    SpreadsheetApp.getUi().alert("Error launching template: " + e.message);
  }
}

// ==========================================
// FUNCTION: CREATE_TEMPLATE_IN_BACKGROUND
// All your existing logic - runs in background while animation shows
// ==========================================
function CREATE_TEMPLATE_IN_BACKGROUND(sheetName, sheetId) {
  try {
    // 1. Create the template document directly in the root of My Drive
    // (no folder - users move it to their preferred location for organization)
    var docFile = Drive.Files.create({
      name: "DocTemplate for " + sheetName,
      mimeType: 'application/vnd.google-apps.document'
    });
    var docId = docFile.id;

    // 2. Open it with the Document service to write canvas + footer
    var doc = DocumentApp.openById(docId);

    // 3. Write onboarding text + the sheet link into the footer
    WRITE_TEMPLATE_ONBOARDING(doc, sheetId);

    doc.saveAndClose();
    var docUrl = doc.getUrl();

    return { success: true, url: docUrl };
    
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ==========================================
// FUNCTION: WRITE_TEMPLATE_ONBOARDING
// Writes the designer onboarding text on the canvas and the linked source
// sheet ID into the document footer. The Docs add-on reads the footer to
// expose the sheet's headers as variables (no Drive API involved).
// ==========================================
function WRITE_TEMPLATE_ONBOARDING(doc, sheetId) {
  var body = doc.getBody();
  body.clear();

  var titleParagraph = body.appendParagraph("📄 DocuMail Pro Template Canvas");
  titleParagraph.setHeading(DocumentApp.ParagraphHeading.HEADING1);

  var descParagraph = body.appendParagraph("\n👉 Go to: Extensions > DocuMail Pro Template > Start Dynamic Doc Template\n\nThis will clear the canvas and link sheet headers, and a DocuMail Pro Template Engine will open on the side, with all the headers available as Variables.\n\nYou are free to insert any variable, any number of times. You can also use Variables with conditions, like when a variable shall be visible.\n\nMake sure to choose Paragraph Text for Paragraph & Table Row for Table, if using conditional insert");
  descParagraph.setHeading(DocumentApp.ParagraphHeading.NORMAL);

  body.appendParagraph("📁 Your template «" + doc.getName() + "» (auto-named from your sheet) is saved in the root of your My Drive. For better organization, right-click it in Drive, select Move to, and pick your preferred folder before generating documents.\n\n" +
    "Generated documents and email attachments are saved to the destination folder you choose in Step 6 of the template wizard (a destination folder is required). No folders are auto-created in your Drive.")
    .setHeading(DocumentApp.ParagraphHeading.NORMAL);

  var footer = doc.getFooter() || doc.addFooter();
  footer.setText('DOCUMAIL_SOURCE_SHEET_ID=' + sheetId);
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

function OPEN_RECORDS_PREVIEW_WINDOW(templateName) {
  var htmlLayout = HtmlService.createTemplateFromFile('RecordsPreviewView');
  htmlLayout.templateName = templateName;
  var dialogWindow = htmlLayout.evaluate().setWidth(1000).setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(dialogWindow, "Confirm Send Records & Daily Limits Check");
}

function SHOW_PREVIEW_DIALOG(templateId) {
  // =======================================================
  // SHOW STEP-BY-STEP ANIMATION DIALOG FIRST
  // =======================================================
  var template = GET_TEMPLATE_BY_ID(templateId);
  var templateName = template ? template.name : "Template";
  
  var loadingHtml = `
<!DOCTYPE html>
<html>
<head>
  <style>
    body {
      margin: 0;
      padding: 0;
      display: flex;
      justify-content: center;
      align-items: center;
      height: 100vh;
      background: #f8f9fa;
      font-family: 'Roboto', sans-serif;
    }
    .loader-container {
      text-align: left;
      padding: 40px 50px;
      background: white;
      border-radius: 12px;
      box-shadow: 0 2px 10px rgba(0,0,0,0.1);
      min-width: 420px;
    }
    .step {
      padding: 10px 0;
      border-bottom: 1px solid #f1f3f4;
      display: flex;
      align-items: center;
      font-size: 14px;
      color: #3c4043;
    }
    .step:last-child {
      border-bottom: none;
    }
    .step .icon {
      width: 28px;
      text-align: center;
      margin-right: 12px;
      font-size: 16px;
    }
    .step.done .icon {
      color: #137333;
    }
    .step.active .icon {
      color: #1a73e8;
    }
    .step.pending .icon {
      color: #dadce0;
    }
    .step.done .text {
      color: #137333;
    }
    .step.active .text {
      color: #1a73e8;
      font-weight: 500;
    }
    .step.pending .text {
      color: #9aa0a6;
    }
    .title {
      font-size: 16px;
      font-weight: 600;
      color: #1a73e8;
      margin-bottom: 16px;
      text-align: center;
    }
    .subtitle {
      font-size: 12px;
      color: #5f6368;
      text-align: center;
      margin-bottom: 12px;
    }
    .error-text {
      color: #c5221f;
      font-size: 14px;
      margin-top: 12px;
      text-align: center;
    }
  </style>
</head>
<body>
  <div class="loader-container" id="loader">
    <div class="title">👁️ Generating Preview</div>
    <div class="subtitle">` + templateName + `</div>
    
    <div class="step active" id="step1">
      <span class="icon">⏳</span>
      <span class="text">📊 Fetching headers from your sheet...</span>
    </div>
    
    <div class="step pending" id="step2">
      <span class="icon">⏳</span>
      <span class="text">📋 Fetching row data for preview...</span>
    </div>
    
    <div class="step pending" id="step3">
      <span class="icon">⏳</span>
      <span class="text">🔍 Mapping template tags to sheet columns...</span>
    </div>
    
    <div class="step pending" id="step4">
      <span class="icon">⏳</span>
      <span class="text">📄 Processing conditional blocks...</span>
    </div>
    
    <div class="step pending" id="step5">
      <span class="icon">⏳</span>
      <span class="text">⚙️ Generating preview document...</span>
    </div>
    
    <div class="step pending" id="step6">
      <span class="icon">⏳</span>
      <span class="text">🚀 Launching preview window...</span>
    </div>
  </div>
  
  <script>
    var steps = [
      { id: 'step1', done: false },
      { id: 'step2', done: false },
      { id: 'step3', done: false },
      { id: 'step4', done: false },
      { id: 'step5', done: false },
      { id: 'step6', done: false }
    ];
    
    function updateStep(stepIndex, status) {
      var step = document.getElementById(steps[stepIndex].id);
      var icon = step.querySelector('.icon');
      var text = step.querySelector('.text');
      
      step.className = 'step';
      if (status === 'done') {
        step.classList.add('done');
        icon.textContent = '✅';
        steps[stepIndex].done = true;
        if (stepIndex < steps.length - 1) {
          var nextStep = document.getElementById(steps[stepIndex + 1].id);
          nextStep.className = 'step active';
          nextStep.querySelector('.icon').textContent = '⏳';
        }
      } else if (status === 'active') {
        step.classList.add('active');
        icon.textContent = '⏳';
      } else {
        step.classList.add('pending');
        icon.textContent = '⏳';
      }
    }
    
    // DUMMY ANIMATION
    setTimeout(function() { updateStep(0, 'done'); }, 300);
    setTimeout(function() { updateStep(1, 'done'); }, 600);
    setTimeout(function() { updateStep(2, 'done'); }, 900);
    setTimeout(function() { updateStep(3, 'done'); }, 1200);
    setTimeout(function() { updateStep(4, 'done'); }, 1500);
    
    // Start the actual work in background
    google.script.run
      .withSuccessHandler(function(result) {
        updateStep(5, 'done');
        // Show the preview result
        var previewHtml = result.previewHtml;
        document.body.innerHTML = previewHtml;
      })
      .withFailureHandler(function(error) {
        document.getElementById('loader').innerHTML = 
          '<div class="title" style="color:#c5221f;">❌ Error</div>' +
          '<div class="error-text">' + error.message + '</div>';
      })
      .GENERATE_PREVIEW_IN_BACKGROUND("` + templateId + `");
  </script>
</body>
</html>
`;

  var loadingDialog = HtmlService.createHtmlOutput(loadingHtml)
    .setWidth(650)
    .setHeight(520)
    .setTitle('Generating Preview...');
  SpreadsheetApp.getUi().showModalDialog(loadingDialog, 'Generating Preview...');
}

// =======================================================
// SHOW RUN DIALOG - WITH STEP-BY-STEP ANIMATION
// =======================================================

function SHOW_RUN_DIALOG(templateId) {
  // =======================================================
  // SHOW STEP-BY-STEP ANIMATION DIALOG FIRST
  // =======================================================
  var template = GET_TEMPLATE_BY_ID(templateId);
  var templateName = template ? template.name : "Template";
  var templateType = template ? template.type : "UNKNOWN";
  var typeLabel = templateType === "PDF_ONLY" ? "📄 PDF" : templateType === "EMAIL_ONLY" ? "✉️ Email" : "📄✉️ Both";
  
  var loadingHtml = `
<!DOCTYPE html>
<html>
<head>
  <style>
    body {
      margin: 0;
      padding: 0;
      display: flex;
      justify-content: center;
      align-items: center;
      height: 100vh;
      background: #f8f9fa;
      font-family: 'Roboto', sans-serif;
    }
    .loader-container {
      text-align: left;
      padding: 40px 50px;
      background: white;
      border-radius: 12px;
      box-shadow: 0 2px 10px rgba(0,0,0,0.1);
      min-width: 420px;
    }
    .step {
      padding: 10px 0;
      border-bottom: 1px solid #f1f3f4;
      display: flex;
      align-items: center;
      font-size: 14px;
      color: #3c4043;
    }
    .step:last-child {
      border-bottom: none;
    }
    .step .icon {
      width: 28px;
      text-align: center;
      margin-right: 12px;
      font-size: 16px;
    }
    .step.done .icon {
      color: #137333;
    }
    .step.active .icon {
      color: #1a73e8;
    }
    .step.pending .icon {
      color: #dadce0;
    }
    .step.done .text {
      color: #137333;
    }
    .step.active .text {
      color: #1a73e8;
      font-weight: 500;
    }
    .step.pending .text {
      color: #9aa0a6;
    }
    .title {
      font-size: 16px;
      font-weight: 600;
      color: #1a73e8;
      margin-bottom: 4px;
      text-align: center;
    }
    .subtitle {
      font-size: 12px;
      color: #5f6368;
      text-align: center;
      margin-bottom: 12px;
    }
    .badge {
      display: inline-block;
      padding: 2px 12px;
      border-radius: 12px;
      font-size: 11px;
      font-weight: 500;
      background: #e8f0fe;
      color: #1a73e8;
    }
    .error-text {
      color: #c5221f;
      font-size: 14px;
      margin-top: 12px;
      text-align: center;
    }
  </style>
</head>
<body>
  <div class="loader-container" id="loader">
    <div class="title">▶️ Running Template</div>
    <div class="subtitle">
      <span class="badge">` + typeLabel + `</span>
      &nbsp; ` + templateName + `
    </div>
    
    <div class="step active" id="step1">
      <span class="icon">⏳</span>
      <span class="text">📊 Fetching headers from your sheet...</span>
    </div>
    
    <div class="step pending" id="step2">
      <span class="icon">⏳</span>
      <span class="text">📋 Scanning for eligible rows...</span>
    </div>
    
    <div class="step pending" id="step3">
      <span class="icon">⏳</span>
      <span class="text">🔍 Mapping template tags to sheet columns...</span>
    </div>
    
    <div class="step pending" id="step4">
      <span class="icon">⏳</span>
      <span class="text">📄 Processing conditional blocks...</span>
    </div>
    
    <div class="step pending" id="step5">
      <span class="icon">⏳</span>
      <span class="text">⚙️ Generating documents & sending emails...</span>
    </div>
    
    <div class="step pending" id="step6">
      <span class="icon">⏳</span>
      <span class="text">✅ Finalizing and updating status...</span>
    </div>
  </div>
  
  <script>
    var steps = [
      { id: 'step1', done: false },
      { id: 'step2', done: false },
      { id: 'step3', done: false },
      { id: 'step4', done: false },
      { id: 'step5', done: false },
      { id: 'step6', done: false }
    ];
    
    function updateStep(stepIndex, status) {
      var step = document.getElementById(steps[stepIndex].id);
      var icon = step.querySelector('.icon');
      var text = step.querySelector('.text');
      
      step.className = 'step';
      if (status === 'done') {
        step.classList.add('done');
        icon.textContent = '✅';
        steps[stepIndex].done = true;
        if (stepIndex < steps.length - 1) {
          var nextStep = document.getElementById(steps[stepIndex + 1].id);
          nextStep.className = 'step active';
          nextStep.querySelector('.icon').textContent = '⏳';
        }
      } else if (status === 'active') {
        step.classList.add('active');
        icon.textContent = '⏳';
      } else {
        step.classList.add('pending');
        icon.textContent = '⏳';
      }
    }
    
    // DUMMY ANIMATION - Steps complete one by one
    setTimeout(function() { updateStep(0, 'done'); }, 300);
    setTimeout(function() { updateStep(1, 'done'); }, 600);
    setTimeout(function() { updateStep(2, 'done'); }, 900);
    setTimeout(function() { updateStep(3, 'done'); }, 1200);
    setTimeout(function() { updateStep(4, 'done'); }, 1500);
    
    // Start the actual work in background
    google.script.run
      .withSuccessHandler(function(result) {
        updateStep(5, 'done');
        // Show the result
        document.body.innerHTML = result;
      })
      .withFailureHandler(function(error) {
        document.getElementById('loader').innerHTML = 
          '<div class="title" style="color:#c5221f;">❌ Error</div>' +
          '<div class="error-text">' + error.message + '</div>';
      })
      .RUN_TEMPLATE_IN_BACKGROUND("` + templateId + `");
  </script>
</body>
</html>
`;

  var loadingDialog = HtmlService.createHtmlOutput(loadingHtml)
    .setWidth(720)
    .setHeight(600)
    .setTitle('Running Template...');
  SpreadsheetApp.getUi().showModalDialog(loadingDialog, 'Running Template...');
}

// ==========================================
// FUNCTION: GENERATE_PREVIEW_IN_BACKGROUND
// Runs preview in background
// ==========================================
function GENERATE_PREVIEW_IN_BACKGROUND(templateId) {
  try {
    var previewFileUrl = '';
    var previewFileDisplayName = '';
    var previewError = '';

    // First, process the preview (generate the document)
    var template = GET_TEMPLATE_BY_ID(templateId);
    if (template && (template.type === "PDF_ONLY" || template.type === "BOTH")) {
      var config = JSON.parse(JSON.stringify(template.config));
      config.isPreview = true;

      REFRESH_TAG_MAPPINGS_IF_DOC_CHANGED(config);

      if (!config.tagMappings || Object.keys(config.tagMappings).length === 0) {
        var allHeaders = GET_ALL_RAW_HEADERS();
        var tags = EXTRACT_TEMPLATE_TAGS_STREAM(config.templateUrl);
        config.tagMappings = {};
        for (var i = 0; i < tags.length; i++) {
          if (allHeaders.indexOf(tags[i]) !== -1) {
            config.tagMappings[tags[i]] = tags[i];
          }
        }
      }

      var mergeResult = EXECUTE_DOCUMENT_MERGE_ENGINE(config);
      if (mergeResult && typeof mergeResult === 'object') {
        if (mergeResult.fileUrl) {
          previewFileUrl = mergeResult.fileUrl;
          previewFileDisplayName = mergeResult.fileDisplayName || '';
        } else if (mergeResult.error) {
          previewError = String(mergeResult.error);
        }
      } else if (typeof mergeResult === 'string') {
        if (mergeResult.indexOf("Error") === 0 || mergeResult.indexOf("No destination folder") === 0) {
          previewError = mergeResult;
        }
      }
    }
    
    // Then get the preview HTML
    var result = PREVIEW_TEMPLATE(templateId, previewFileUrl, previewFileDisplayName, previewError);
    return result;
    
  } catch (e) {
    return { name: "Error", previewHtml: "<p>ERROR: " + e.toString() + "</p>" };
  }
}

// ==========================================
// FUNCTION: RUN_TEMPLATE_IN_BACKGROUND
// Runs template in background with file links - MATCHES PREVIEW STYLE
// ==========================================
function RUN_TEMPLATE_IN_BACKGROUND(templateId) {
  try {
    var template = GET_TEMPLATE_BY_ID(templateId);
    
    if (!template) {
      return '<div style="font-family: Roboto, sans-serif; padding: 24px;">' +
             '<h2 style="color: #c5221f; margin: 0 0 12px 0;">❌ Template Not Found</h2>' +
             '<hr style="border: none; border-top: 1px solid #e8eaed; margin: 12px 0;">' +
             '<p style="color: #5f6368; margin: 8px 0;">Template ID: <strong>' + templateId + '</strong></p>' +
             '<p style="color: #5f6368; margin: 8px 0;">Please delete this template and recreate it.</p>' +
             '<hr style="border: none; border-top: 1px solid #e8eaed; margin: 16px 0;">' +
             '<div style="margin-top: 20px; text-align: center;">' +
             '<button onclick="google.script.host.close()" style="background: #1a73e8; color: white; border: none; padding: 10px 28px; border-radius: 4px; cursor: pointer; font-size: 14px; font-weight: 500;">Close</button>' +
             '</div></div>';
    }
    
    // Process mapping configurations
    if (template.type === "PDF_ONLY" || template.type === "BOTH") {
      var config = JSON.parse(JSON.stringify(template.config));
      config.isPreview = false;
      
      if (!config.tagMappings) {
        config.tagMappings = {};
      }

      REFRESH_TAG_MAPPINGS_IF_DOC_CHANGED(config);

      if (Object.keys(config.tagMappings).length === 0) {
        var allHeaders = GET_ALL_RAW_HEADERS();
        var tags = EXTRACT_TEMPLATE_TAGS_STREAM(config.templateUrl);
        for (var i = 0; i < tags.length; i++) {
          if (allHeaders.indexOf(tags[i]) !== -1) {
            config.tagMappings[tags[i]] = tags[i];
          }
        }
        template.config.tagMappings = config.tagMappings;
        if (typeof SAVE_TEMPLATE === 'function') {
          SAVE_TEMPLATE(template);
        }
      }
    }
    
    // Capture rows already "Success" BEFORE the run so reused files are not listed as newly generated
    var reusedRowSet = {};
    if (template.type === "PDF_ONLY" || template.type === "BOTH") {
      try {
        var preSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
        var preHeaders = GET_ALL_RAW_HEADERS(preSheet);
        var preStatusColIdx = -1;
        for (var c = 0; c < preHeaders.length; c++) {
          if (String(preHeaders[c]).toLowerCase().indexOf("merged doc status") !== -1) {
            preStatusColIdx = c;
            break;
          }
        }
        if (preStatusColIdx !== -1) {
          var preLastRow = preSheet.getLastRow();
          var preStatusRange = preSheet.getRange(2, preStatusColIdx + 1, Math.max(preLastRow - 1, 1), 1).getValues();
          for (var r = 0; r < preStatusRange.length; r++) {
            if (String(preStatusRange[r][0] || "").trim() === "Success") {
              reusedRowSet[r + 2] = true;
            }
          }
        }
      } catch (e) {
        // Sheet read error - ignore, all rows will be considered new
      }
    }

    var result = RUN_TEMPLATE(templateId);

    // Get file URLs from the sheet's "Merged Doc URL" column (drive.file compatible - no Drive scanning)
    var fileUrls = [];
    if ((template.type === "PDF_ONLY" || template.type === "BOTH") && result !== "NO_ROWS_ELIGIBLE" && result && result.indexOf("Error") === -1) {
      try {
        var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
        var allHeaders = GET_ALL_RAW_HEADERS(activeSheet);
        var mergedUrlColIdx = -1;
        for (var c = 0; c < allHeaders.length; c++) {
          if (String(allHeaders[c]).toLowerCase().indexOf("merged doc url") !== -1) {
            mergedUrlColIdx = c;
            break;
          }
        }
        if (mergedUrlColIdx !== -1) {
          var lastRow = activeSheet.getLastRow();
          var urlRange = activeSheet.getRange(2, mergedUrlColIdx + 1, Math.max(lastRow - 1, 1), 1).getValues();
          for (var r = urlRange.length - 1; r >= 0; r--) {
            if (reusedRowSet[r + 2]) continue;
            var cellUrl = String(urlRange[r][0] || "").trim();
            if (cellUrl !== "" && cellUrl.indexOf("http") === 0) {
              fileUrls.push({ name: "Generated Document (Row " + (r + 2) + ")", url: cellUrl });
              if (fileUrls.length >= 10) break;
            }
          }
        }
      } catch (e) {
        // Sheet read error - ignore, file links section will be skipped
      }
    }
    var html = '<style>body{display:block!important;align-items:initial!important;justify-content:initial!important;overflow-y:auto;margin:0;padding:0;}</style><div style="font-family: Roboto, sans-serif; padding: 24px;">';
    
    if (result && result.indexOf("Error") !== -1) {
      html += '<h2 style="color: #c5221f; margin: 0 0 12px 0;">❌ Execution Failed</h2>';
      html += '<hr style="border: none; border-top: 1px solid #e8eaed; margin: 12px 0;">';
      html += '<div style="background: #fce8e6; padding: 16px; border-radius: 6px; margin: 12px 0; border-left: 4px solid #c5221f; line-height: 1.6;">';
      html += '<span style="color: #c5221f;">' + result.replace(/\n/g, '<br>') + '</span>';
      html += '</div>';
      
    } else if (result === "NO_ROWS_ELIGIBLE") {
      html += '<h2 style="color: #f2994a; margin: 0 0 12px 0;">ℹ️ No Rows Eligible</h2>';
      html += '<hr style="border: none; border-top: 1px solid #e8eaed; margin: 12px 0;">';
      html += '<div style="background: #fff3cd; padding: 16px; border-radius: 6px; margin: 12px 0; border-left: 4px solid #f2994a; color: #856404; line-height: 1.6;">';
      html += 'All rows matching your filter conditions have already been merged successfully with a status of <strong>Success</strong>.<br><br>';
      html += 'There are no fresh pending records available to process.';
      html += '</div>';
      
    } else {
      html += '<h2 style="color: #1a73e8; margin: 0 0 12px 0;">✅ Execution Completed!</h2>';
      html += '<hr style="border: none; border-top: 1px solid #e8eaed; margin: 12px 0;">';
      
      if (template.type === "EMAIL_ONLY") {
        html += '<div style="background: #e8f0fe; padding: 16px; border-radius: 6px; margin: 12px 0;">';
        html += '<h3 style="color: #1a73e8; margin: 0 0 10px 0;">📧 Email Summary</h3>';
        html += '<p style="margin: 0 0 8px 0;"><strong>Template:</strong> ' + escapeHtml(template.name) + '</p>';
        html += '<p style="margin: 0;"><strong>Type:</strong> Email Only</p>';
        html += '</div>';
      } else {
        html += '<div style="background: #e8f0fe; padding: 16px; border-radius: 6px; margin: 12px 0;">';
        html += '<h3 style="color: #1a73e8; margin: 0 0 10px 0;">📄 Document Preview</h3>';
        html += '<p style="margin: 0 0 8px 0;"><strong>Template:</strong> ' + escapeHtml(template.name) + '</p>';
        html += '<p style="margin: 0 0 8px 0;"><strong>Type:</strong> ' + (template.type === "PDF_ONLY" ? "PDF Only" : "PDF & Email") + '</p>';
        html += '<p style="margin: 0;"><strong>Destination:</strong> ' + escapeHtml(template.config?.folderDestination || "Default Folder") + '</p>';
        html += '</div>';
      }
      
      // Result section
      html += '<div style="background: #e6f4ea; padding: 16px; border-radius: 6px; margin: 12px 0; border-left: 4px solid #137333;">';
      var successMsg = template.type === "EMAIL_ONLY" ? "✅ Emails Sent Successfully!" : "✅ Documents Generated Successfully!";
      html += '<p style="margin: 0; color: #137333;"><strong>' + successMsg + '</strong></p>';
      html += '<p style="margin: 5px 0;"><strong>Result:</strong> ' + result.replace(/\n/g, '<br>') + '</p>';

      // Email quota note (static - no live quota API available in this scope)
      if (template.type !== "PDF_ONLY") {
        html += '<p style="margin: 8px 0 0 0; color: #c5221f;"><strong>📊 Email Quota:</strong> Check your automated email quota — 100 emails/day for a normal Google account, 1500/day for Google Workspace.</p>';
      }

      html += '</div>';
      
      // Email details section - MERGED with a real processed row (like Preview)
      if (template.type === "EMAIL_ONLY" || template.type === "BOTH") {
        var emailPreview = BUILD_RUN_EMAIL_PREVIEW(template);
        if (emailPreview) {
          html += '<div style="background:#e6f4ea; padding:12px; border-radius:6px; margin:10px 0;">';
          html += '<h3 style="color:#137333; margin:0 0 10px 0;">✉️ Email Details</h3>';
          html += '<p><strong>To:</strong> ' + escapeHtml(emailPreview.recipientEmail) + '</p>';
          html += '<p><strong>Subject:</strong> ' + escapeHtml(emailPreview.subject) + '</p>';
          html += '<div style="border:1px solid #dadce0; padding:12px; border-radius:6px; margin-top:12px; background:#ffffff;"><strong>Email Body:</strong><br><br>' + emailPreview.body + '</div>';
          html += '</div>';
        }
      }
      
      // File links section - MATCHES PREVIEW STYLE
      if (fileUrls.length > 0) {
        html += '<div style="background: #e6f4ea; padding: 16px; border-radius: 6px; margin: 12px 0; border-left: 4px solid #137333;">';
        html += '<p style="margin: 0; color: #137333;"><strong>📂 Generated Files:</strong></p>';
        for (var i = 0; i < fileUrls.length; i++) {
          html += '<p style="margin: 5px 0;"><a href="' + fileUrls[i].url + '" target="_blank" style="color: #1a73e8; text-decoration: none;">📄 ' + escapeHtml(fileUrls[i].name) + '</a></p>';
        }
        html += '</div>';
      }
    }
    
    html += '<hr style="border: none; border-top: 1px solid #e8eaed; margin: 16px 0;">';
    html += '<div style="margin-top: 20px; text-align: center;">';
    html += '<button onclick="google.script.host.close()" style="background: #1a73e8; color: white; border: none; padding: 10px 28px; border-radius: 4px; cursor: pointer; font-size: 14px; font-weight: 500;">Close</button>';
    html += '</div>';
    html += '</div>';
    
    return html;
    
  } catch (e) {
    return '<div style="font-family: Roboto, sans-serif; padding: 24px;">' +
           '<h2 style="color: #c5221f; margin: 0 0 12px 0;">❌ Error</h2>' +
           '<hr style="border: none; border-top: 1px solid #e8eaed; margin: 12px 0;">' +
           '<p style="color: #5f6368; margin: 8px 0;">' + e.message + '</p>' +
           '<hr style="border: none; border-top: 1px solid #e8eaed; margin: 16px 0;">' +
           '<div style="margin-top: 20px; text-align: center;">' +
           '<button onclick="google.script.host.close()" style="background: #1a73e8; color: white; border: none; padding: 10px 28px; border-radius: 4px; cursor: pointer; font-size: 14px; font-weight: 500;">Close</button>' +
           '</div></div>';
  }
}

// ==========================================
// FUNCTION: BUILD_RUN_EMAIL_PREVIEW
// Merges To/Subject/Body against the first actually-processed row (like Preview)
// ==========================================
function BUILD_RUN_EMAIL_PREVIEW(template) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var allHeaders = GET_ALL_RAW_HEADERS(sheet);

    var emailColIdx = -1;
    for (var c = 0; c < allHeaders.length; c++) {
      var hName = String(allHeaders[c]).toLowerCase().trim();
      if (hName.indexOf("recipient email") !== -1) emailColIdx = c;
    }
    if (emailColIdx === -1) emailColIdx = 4;

    var emailStatusColName = "Sent Mail Status - " + template.name;
    var emailStatusColIdx = -1;
    for (var c = 0; c < allHeaders.length; c++) {
      if (String(allHeaders[c]).toLowerCase().trim() === emailStatusColName.toLowerCase()) {
        emailStatusColIdx = c;
        break;
      }
    }

    var criteriaColIdx = allHeaders.indexOf(template.config.condField);
    if (criteriaColIdx === -1) criteriaColIdx = 0;

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return null;

    var data = sheet.getRange(2, 1, lastRow - 1, allHeaders.length).getValues();
    var processedRow = null;

    for (var i = 0; i < data.length; i++) {
      var evalCellText = String(data[i][criteriaColIdx] || "").trim();
      var isCriteriaMatch = false;
      if (template.config.condOperator === "NOT_EMPTY" && evalCellText !== "") {
        isCriteriaMatch = true;
      } else if (template.config.condOperator === "CONTAINS" && evalCellText.toLowerCase().indexOf(String(template.config.condValue).toLowerCase().trim()) !== -1) {
        isCriteriaMatch = true;
      }
      if (!isCriteriaMatch) continue;

      var wasSent = emailStatusColIdx !== -1 && String(data[i][emailStatusColIdx] || "").trim() !== "";
      if (!wasSent) continue;

      processedRow = data[i];
      break;
    }

    if (!processedRow) return null;

    var recipientEmail = "example@email.com";
    var emailToConfig = template.emailConfig?.to || "";
    if (emailToConfig) {
      var toTagMatch = emailToConfig.match(/^\{(.+)\}$/);
      if (toTagMatch) {
        var toColName = toTagMatch[1];
        var toColIdx = allHeaders.indexOf(toColName);
        if (toColIdx !== -1) {
          recipientEmail = processedRow[toColIdx] || recipientEmail;
        }
      } else {
        recipientEmail = emailToConfig;
      }
    } else {
      recipientEmail = processedRow[emailColIdx] || recipientEmail;
    }

    var emailSubject = template.emailConfig?.subject || "";
    var emailBody = template.emailConfig?.body || "";

    for (var h = 0; h < allHeaders.length; h++) {
      var header = allHeaders[h];
      if (header) {
        var val = processedRow[h];
        if (val instanceof Date) {
          val = FORMAT_DATE_FOR_DISPLAY(val);
        } else if (typeof val === 'number') {
          val = FORMAT_NUMBER_FOR_DISPLAY(val);
        } else {
          val = String(val || "");
        }
        var regex = new RegExp("\\{" + escapeRegex(header) + "\\}", "g");
        emailSubject = emailSubject.replace(regex, val);
        emailBody = emailBody.replace(regex, val);
      }
    }

    return { recipientEmail: recipientEmail, subject: emailSubject, body: emailBody };
  } catch (e) {
    return null;
  }
}


// =======================================================
// SHOW NOTIFICATION FUNCTIONS
// =======================================================

function SHOW_NOTIFICATION(title, message, type) {
  var icon = type === "success" ? "✅" : (type === "warning" ? "⚠️" : "ℹ️");
  var bgColor = type === "success" ? "#e6f4ea" : (type === "warning" ? "#fce8e6" : "#e8f0fe");
  var borderColor = type === "success" ? "#137333" : (type === "warning" ? "#c5221f" : "#1a73e8");

  var html = '<div style="font-family: Roboto, sans-serif; padding: 20px;">' +
    '<div style="display: flex; align-items: center; margin-bottom: 16px;">' +
    '<span style="font-size: 28px; margin-right: 12px;">' + icon + '</span>' +
    '<h2 style="color: #1a73e8; margin: 0;">' + title + '</h2>' +
    '</div>' +
    '<div style="background: ' + bgColor + '; padding: 16px; border-radius: 6px; margin: 10px 0; border-left: 4px solid ' + borderColor + ';">' +
    '<p style="margin: 0; color: #3c4043; line-height: 1.5;">' + message + '</p>' +
    '</div>' +
    '<div style="margin-top: 24px; text-align: center;">' +
    '<button onclick="google.script.host.close()" style="background: #1a73e8; color: white; border: none; padding: 8px 24px; border-radius: 4px; font-size: 14px; font-family: Roboto, sans-serif; cursor: pointer;">OK</button>' +
    '</div>' +
    '</div>';

  return HtmlService.createHtmlOutput(html)
    .setWidth(450)
    .setHeight(220)
    .setTitle(title);
}

function SHOW_COLUMN_CREATED_NOTIFICATION(templateName) {
  return SHOW_NOTIFICATION(
    "Status Column Created",
    "✅ Column to track Status \"" + templateName + "\" Template has been created in your Google Sheet. It is protected with system styling.",
    "success"
  );
}

function SHOW_COLUMN_EXISTS_NOTIFICATION(templateName) {
  return SHOW_NOTIFICATION(
    "Status Column Already Exists",
    "ℹ️ Column to track Status \"" + templateName + "\" Template already exists in your Google Sheet. No changes were made.",
    "info"
  );
}

function SHOW_TEMPLATE_SAVED_NOTIFICATION(templateName, columnCreated, columnExists) {
  if (columnCreated) {
    return SHOW_COLUMN_CREATED_NOTIFICATION(templateName);
  } else if (columnExists) {
    return SHOW_COLUMN_EXISTS_NOTIFICATION(templateName);
  } else {
    return SHOW_NOTIFICATION(
      "Template Saved",
      "✅ Template \"" + templateName + "\" has been saved successfully.",
      "success"
    );
  }
}

// =======================================================
// AUTO MATCH TAGS FUNCTION 
// =======================================================

function AUTO_MATCH_TAGS(docUrl) {
  try {
    var tags = EXTRACT_TEMPLATE_TAGS_STREAM(docUrl);
    var headers = GET_LIVE_SHEET_HEADERS();
    var matched = [];
    var unmatched = [];

    // Build a lookup map for headers
    var headerMap = {};
    for (var i = 0; i < headers.length; i++) {
      headerMap[headers[i].toLowerCase().trim()] = headers[i];
    }

    for (var i = 0; i < tags.length; i++) {
      var tag = tags[i];
      var matchedHeader = null;
      for (var h = 0; h < headers.length; h++) {
        if (headers[h].toLowerCase().trim() === tag.toLowerCase().trim()) {
          matchedHeader = headers[h];
          break;
        }
      }
      if (matchedHeader) {
        matched.push({ tag: tag, header: matchedHeader });
      } else {
        unmatched.push(tag);
      }
    }
    return { success: true, matched: matched, unmatched: unmatched };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ==========================================
// FUNCTION: CREATE_DOCUMENT_TEMPLATE
// Creates a clean, blank Google Doc template
// ==========================================
function CREATE_DOCUMENT_TEMPLATE() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();
    var sheetName = sheet.getName();

    // Create a beautifully clean, completely blank document in the root of My Drive
    // (no folder - users move it to their preferred location for organization)
    var docName = "DocTemplate for " + sheetName;
    var docFile = Drive.Files.create({
      name: docName,
      mimeType: 'application/vnd.google-apps.document'
    });
    var docId = docFile.id;
    var doc = DocumentApp.openById(docId);

    // Write onboarding text + the sheet link into the footer
    // (replaces the old userProperties/Drive-description linking mechanism)
    WRITE_TEMPLATE_ONBOARDING(doc, ss.getId());

    doc.saveAndClose();
    var docUrl = doc.getUrl();

    console.log("✅ Blank Template created: " + docName);
    return {
      success: true,
      url: docUrl,
      id: docId,
      name: docName,
      folder: "My Drive"
    };

  } catch (e) {
    console.error("Error creating document template: " + e.message);
    return { success: false, error: e.message };
  }
}



//file content end