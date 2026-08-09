/**
 * DocuMain.gs (Inside your Document Add-on)
 * Main Controller for the Document Template Designer.
 * The add-on menu (onOpen / createMenu) lives in Code.js.
 */

/**
 * Shows help information to the user
 */
function showHelp() {
  var ui = DocumentApp.getUi();
  ui.alert(
    '📚 DocuMail Pro - Help',
    'How to use Smart Variables:\n\n' +
    '1. Click "Start Dynamic Template" to clear the canvas and start fresh\n' +
    '2. Click "Open Smart Variable Window" to open the sidebar\n' +
    '3. Select a variable from the dropdown\n' +
    '4. Choose to insert it "Always" or "Only when..." with conditions\n' +
    '5. Click "Insert" to add it to your template\n\n' +
    'Variables will be replaced with actual data when generating documents.',
    ui.ButtonSet.OK
  );
}

/**
 * Classifies the active document into one of:
 *   'initialized' - marker set or a source sheet is already linked
 *   'footer'      - created from a DocuMail Pro sheet (footer link) but not yet
 *                   initialized
 *   'blank'       - empty body, no footer, no link
 *   'custom'      - an existing document with its own content (no footer/link)
 */
function GET_TEMPLATE_STATE() {
  try {
    var props = PropertiesService.getDocumentProperties();
    if (props.getProperty('DOCUMAIL_TEMPLATE_INITIALIZED') === 'true') return 'initialized';
    if (props.getProperty('DOCUMAIL_SOURCE_SHEET_ID')) return 'initialized';

    var doc = DocumentApp.getActiveDocument();
    var footer = doc.getFooter();
    if (footer && footer.getText().indexOf('DOCUMAIL_SOURCE_SHEET_ID=') !== -1) return 'footer';

    var bodyText = doc.getBody().getText().trim();
    if (bodyText.length === 0) return 'blank';
    return 'custom';
  } catch (e) {
    return 'custom';
  }
}

/**
 * Renders the Smart Variable sidebar panel.
 */
function showTemplateSidebar(title) {
  var htmlOutput = HtmlService.createTemplateFromFile('SidebarDocView')
    .evaluate()
    .setTitle(title || 'DocuMail Pro - Smart Variables')
    .setWidth(300);
  DocumentApp.getUi().showSidebar(htmlOutput);
}

/**
 * Shared gated sidebar opener.
 * @param {boolean} showBlankHint - only the "Open Smart Variable Window" menu
 *        shows the informational hint on blank documents; the Link menu skips it.
 */
function openSidebarGate(showBlankHint) {
  try {
    var ui = DocumentApp.getUi();
    var state = GET_TEMPLATE_STATE();

    // Sheet-created templates must be initialized before the sidebar can open.
    if (state === 'footer') {
      var initResponse = ui.alert(
        '⚠️ DocuMail Pro Template Not Initiated',
        'Template is not initiated yet. Do you want to initialize it now?',
        ui.ButtonSet.YES_NO
      );
      if (initResponse !== ui.Button.YES) {
        DocumentApp.getActiveDocument().toast(
          'Initialization cancelled. Run "Initialize DocuMail Pro Template" when you are ready.',
          "DocuMail Pro", 4
        );
        return;
      }
      INITIALIZE_DOC_DESIGNER_SIDEBAR();
      return;
    }

    // Blank documents: explain that they need to be linked first (only for the
    // "Open Smart Variable Window" entry point).
    if (state === 'blank' && showBlankHint) {
      ui.alert(
        '💡 DocuMail Pro Template',
        'This template was not created from a spreadsheet using the DocuMail Pro Sheets add-on.\n\n' +
        'Link the document to an existing spreadsheet via "Link Spreadsheet" in the sidebar or from the menu, to fetch the sheet headers as variables of this document.',
        ui.ButtonSet.OK
      );
    }

    // initialized / custom / blank: open the sidebar so the user can link or
    // insert variables.
    showTemplateSidebar('DocuMail Pro - Smart Variables');

  } catch (error) {
    Logger.log("Error opening sidebar: " + error.toString());
    DocumentApp.getUi().alert("Could not open sidebar: " + error.message);
  }
}

/**
 * Open Smart Variable Window - Opens the sidebar WITHOUT resetting the canvas.
 * Shows the informational hint on blank documents.
 */
function OPEN_SMART_VARIABLE_WINDOW() {
  openSidebarGate(true);
}

/**
 * Link to Spreadsheet menu item - Opens the sidebar WITHOUT the blank-document
 * hint (linking is the direct action here).
 */
function OPEN_SMART_VARIABLE_WINDOW_FROM_LINK_MENU() {
  openSidebarGate(false);
}

/**
 * Homepage auto-open: only shows the sidebar for already-initialized/linked
 * templates. For footer/blank/custom documents it does nothing, so the canvas
 * guidance stays visible (no auto-opening that contradicts the Step-1 warning).
 */
function OPEN_SMART_VARIABLE_WINDOW_FROM_HOMEPAGE() {
  try {
    if (GET_TEMPLATE_STATE() !== 'initialized') return;
    showTemplateSidebar('DocuMail Pro - Smart Variables');
  } catch (error) {
    Logger.log("Could not open sidebar from homepage: " + error.toString());
  }
}

/**
 * Checks for existing custom layout text, prompts user, and opens the sidebar interface.
 */
function INITIALIZE_DOC_DESIGNER_SIDEBAR() {
  try {
    var doc = DocumentApp.getActiveDocument();
    var body = doc.getBody();
    var ui = DocumentApp.getUi();

    // 1. Get all text currently sitting on the canvas
    var currentText = body.getText().trim();

    // 1b. One-time footer handoff: the Sheets add-on writes the linked source
    // sheet ID into this document's footer. Capture it into document
    // properties (fast, scope-safe) then clear the footer so it never
    // re-triggers. Fallback reads remain in getSheetHeaders().
    var sourceSheetId = null;
    var footer = doc.getFooter();
    if (footer) {
      var fMatch = footer.getText().match(/DOCUMAIL_SOURCE_SHEET_ID=([\w-]+)/);
      if (fMatch) sourceSheetId = fMatch[1];
    }
    if (sourceSheetId) {
      PropertiesService.getDocumentProperties().setProperty('DOCUMAIL_SOURCE_SHEET_ID', sourceSheetId);
      footer.setText("");
    }

    // 2. If the page has any text, warn before wiping it
    if (currentText.length > 0) {
      var response = ui.alert(
        '⚠️ Confirm Initialization',
        'Initializing will clear the content of the page. Make sure you want to proceed before initializing.',
        ui.ButtonSet.YES_NO
      );

      // If user clicks "NO", stop execution instantly
      if (response !== ui.Button.YES) {
        doc.toast("Initialization cancelled. Your existing template was saved.", "🚀 DocuMail Pro");
        return;
      }
    }

    // 4. Wipe the canvas completely clean (keep last paragraph to avoid section error)
    var pars = body.getParagraphs();
    while (pars.length > 1) {
      body.removeChild(pars[0]);
      pars = body.getParagraphs();
    }
    pars[0].setText(" ");

    // 5. Mark as initialized so future sidebar opens skip the gate
    PropertiesService.getDocumentProperties().setProperty('DOCUMAIL_TEMPLATE_INITIALIZED', 'true');

    // 6. Evaluate and render the HTML sidebar panel
    var htmlOutput = HtmlService.createTemplateFromFile('SidebarDocView')
      .evaluate()
      .setTitle('DocuMail Pro - Template Designer')
      .setWidth(300);

    ui.showSidebar(htmlOutput);

  } catch (error) {
    Logger.log("Error initializing document canvas workspace: " + error.toString());
    DocumentApp.getUi().alert("Could not load designer panel: " + error.message);
  }
}

/**
 * Universal token injector engine.
 * Added 'content' parameter to allow custom template content insertion.
 */
function injectTagAtCursor(openingTag, closingTag, content) {
  try {
    var doc = DocumentApp.getActiveDocument();
    var body = doc.getBody();
    var cursor = doc.getCursor();

    if (!cursor) {
      throw new Error("Please click your mouse cursor inside the document body.");
    }

    var element = cursor.getElement();
    var offset = cursor.getOffset();
    var textElement = null;
    var parentParagraph = null;

    // Get the current text element and its parent paragraph
    var elementType = element.getType();
    if (elementType === DocumentApp.ElementType.TEXT) {
      textElement = element.asText();
      parentParagraph = textElement.getParent();
      if (offset === 0 && textElement.getText().length > 0) offset = textElement.getText().length;
    } else if (elementType === DocumentApp.ElementType.PARAGRAPH) {
      var paragraph = element.asParagraph();
      parentParagraph = paragraph;
      if (paragraph.getText() === "") {
        textElement = paragraph.appendText(" ");
        offset = 1;
      } else {
        textElement = paragraph.getChild(0).asText();
        offset = textElement.getText().length;
      }
    } else if (elementType === DocumentApp.ElementType.BODY_SECTION || elementType === DocumentApp.ElementType.DOCUMENT) {
      var targetPara = body.getParagraphs()[0] || body.appendParagraph("");
      parentParagraph = targetPara;
      if (targetPara.getText() === "") {
        textElement = targetPara.appendText(" ");
        offset = 1;
      } else {
        textElement = targetPara.getChild(0).asText();
        offset = textElement.getText().length;
      }
    } else {
      var parent = element.getParent();
      while (parent && parent.getType() !== DocumentApp.ElementType.PARAGRAPH) parent = parent.getParent();
      if (parent) {
        parentParagraph = parent.asParagraph();
        if (parentParagraph.getText() === "") {
          textElement = parentParagraph.appendText(" ");
          offset = 1;
        } else {
          textElement = parentParagraph.getChild(0).asText();
          offset = textElement.getText().length;
        }
      } else {
        throw new Error("Please click inside an active text line.");
      }
    }

    var closeStr = closingTag ? closingTag.trim() : "";
    var isContentEmpty = (content === undefined || content === null || content.toString().trim() === "");

    // ============================================================
    // EARLY RETURN: Simple variable insert (no tags, no content)
    // ============================================================
    if (closeStr === "" && isContentEmpty) {
      var simpleText = parentParagraph.appendText(openingTag);
      simpleText.setForegroundColor(defaultForegroundColor);
      simpleText.setFontSize(defaultFontSize);
      simpleText.setFontFamily(defaultFontFamily);
      simpleText.setItalic(false);
      simpleText.setBold(false);
      simpleText.setUnderline(false);
      var cursorPos = doc.newPosition(simpleText, simpleText.getText().length);
      doc.setCursor(cursorPos);
      return { success: true };
    }

    // Determine what content to insert
    var contentToInsert;
    if (content !== undefined && content !== null && content.toString().trim() !== "") {
      contentToInsert = content;
    } else {
      contentToInsert = "[ Enter your conditional template content here ]";
    }

    // ============================================================
    // Get the full text and split
    // ============================================================
    var fullText = textElement.getText();
    var beforeText = fullText.substring(0, offset);
    var afterText = fullText.substring(offset);
    
    // Clear the current text element
    textElement.setText("");
    
    // Insert the text before the cursor
    textElement.appendText(beforeText);
    
    // ============================================================
    // GET DEFAULT FORMATTING FROM THE PARAGRAPH
    // ============================================================
    // Get the paragraph's default formatting
    var defaultFontSize = parentParagraph.getFontSize();
    var defaultForegroundColor = parentParagraph.getForegroundColor();
    var defaultFontFamily = parentParagraph.getFontFamily();
    
    // If paragraph has no explicit formatting, use body defaults
    if (!defaultFontSize) defaultFontSize = 11;
    if (!defaultForegroundColor) defaultForegroundColor = "#000000";
    if (!defaultFontFamily) defaultFontFamily = "Arial";
    
    // ============================================================
    // CREATE SEPARATE TEXT ELEMENTS FOR EACH PART
    // ============================================================
    
    // 1. Opening tag - STYLED
    var openingText = parentParagraph.appendText(openingTag);
    openingText.setForegroundColor("#b0b0b0");
    openingText.setFontSize(9);
    openingText.setItalic(true);
    
    // 2. Content - APPLY DEFAULT FORMATTING EXPLICITLY
    var contentText = parentParagraph.appendText(contentToInsert);
    // Explicitly set to default values (not null)
    contentText.setForegroundColor(defaultForegroundColor);
    contentText.setFontSize(defaultFontSize);
    contentText.setFontFamily(defaultFontFamily);
    contentText.setItalic(false);
    contentText.setBold(false);
    contentText.setUnderline(false);
    
    // 3. Closing tag (if exists) - STYLED
    var closingText = null;
    if (closeStr !== "") {
      closingText = parentParagraph.appendText(closeStr);
      closingText.setForegroundColor("#b0b0b0");
      closingText.setFontSize(9);
      closingText.setItalic(true);
    }
    
    // 4. Remaining text after cursor
    // Ensure there's always a default-formatted run at the end, so grey/italic
    // from the closing tag never becomes the "last format" a new line inherits.
    var afterTextNode = parentParagraph.appendText(afterText.length > 0 ? afterText : " ");
    afterTextNode.setForegroundColor(defaultForegroundColor);
    afterTextNode.setFontSize(defaultFontSize);
    afterTextNode.setFontFamily(defaultFontFamily);
    afterTextNode.setItalic(false);
    afterTextNode.setBold(false);
    afterTextNode.setUnderline(false);
    
    // ============================================================
    // SET CURSOR POSITION - At the end of content
    // ============================================================
    
    var cursorPos = doc.newPosition(contentText, contentText.getText().length);
    doc.setCursor(cursorPos);

    return { success: true };

  } catch (error) {
    Logger.log("Core insertion crash tracker: " + error.toString());
    throw new Error(error.message);
  }
}

// ======================================================================
// 📦 Styling of tags - For existing documents (legacy support)
// ======================================================================
function styleLogicTags() {
  var body = DocumentApp.getActiveDocument().getBody();
  
  var tagDefinitions = [
    { tag: "<<If:", length: 5 },
    { tag: "<<EndIf>>", length: 9 },
    { tag: "<<RowIf:", length: 8 }
  ];
  
  tagDefinitions.forEach(function(item) {
    var tag = item.tag;
    var tagLength = item.length;
    var found = body.findText(tag);
    
    while (found) {
      var text = found.getElement().asText();
      var start = found.getStartOffset();
      var end = found.getEndOffsetInclusive();
      var fullText = text.getText();
      
      // Style the tag only
      var tagEnd = Math.min(start + tagLength - 1, fullText.length - 1);
      text.setForegroundColor(start, tagEnd, "#b0b0b0");
      text.setFontSize(start, tagEnd, 9);
      text.setItalic(start, tagEnd, true);
      
      // Reset formatting for content after tag (prevents bleed)
      var contentStart = tagEnd + 1;
      var contentEnd = fullText.length - 1;
      if (contentStart <= contentEnd) {
        text.setForegroundColor(contentStart, contentEnd, null);
        text.setFontSize(contentStart, contentEnd, null);
        text.setItalic(contentStart, contentEnd, false);
        text.setBold(contentStart, contentEnd, false);
        text.setUnderline(contentStart, contentEnd, false);
      }
      
      found = body.findText(tag, found);
    }
  });
}

// ======================================================================
// 📢 Toast notification for sidebar feedback
// ======================================================================
function showToast(message) {
  try {
    DocumentApp.getActiveDocument().toast(message, "DocuMail Pro", 3);
  } catch (e) {
    Logger.log("Toast error: " + e.toString());
  }
}