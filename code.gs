/**
 * Indents = Headings - Document outline creation.
 *
 * Use for outline preperation. For document set up with many headings.
 *
 * Applies heading styles based on paragraph indentation (levels Subtitle-H6).
 * 
 * @OnlyCurrentDoc
 * Converts indented paragraphs into headings
 * - 3 indents → Subtitle
 * - 4 indents → Heading 1
 * - ...
 * - 9+ indents → Heading 6
 * 
 * Resets indentation to zero for paragraphs that become headings.
 */

function onOpen(e) {
  if (e && e.authMode !== ScriptApp.AuthMode.NONE) {
    DocumentApp.getUi()
      .createMenu('Doc Outliner')
      .addItem('Indents = Headings', 'convertIndentsToHeadings')
      .addToUi();
  }
}

function convertIndentsToHeadings() {
  const doc = DocumentApp.getActiveDocument();
  const paragraphs = doc.getSelection()
    ? getValidParagraphsFromSelection(doc.getSelection())
    : getValidParagraphs(doc.getBody());
  
  let changesMade = false;
  
  paragraphs.forEach(paragraph => {
    const indent = getFirstLineIndent(paragraph);
    if (indent >= 3) {
      if (applyHeading(paragraph, indent)) {
        changesMade = true;
      }
    }
  });
  
  showCompletionAlert(changesMade);
}

// ========== IMPROVED HELPER FUNCTIONS ========== //

function getValidParagraphsFromSelection(selection) {
  return selection.getRangeElements().reduce((acc, element) => {
    try {
      const elem = element.getElement();
      if (elem.getType() === DocumentApp.ElementType.PARAGRAPH) {
        const para = elem.asParagraph();
        if (para.getText().trim().length > 0) acc.push(para);
      }
    } catch (e) {
      console.warn("Skipped non-paragraph element");
    }
    return acc;
  }, []);
}

function getValidParagraphs(body) {
  return body.getParagraphs().filter(para => {
    return para.getText().trim().length > 0;
  });
}

function getFirstLineIndent(paragraph) {
  try {
    const indent = paragraph.getIndentFirstLine();
    return Math.max(0, Math.round(indent / 36)); // 36pt = 1 indent level
  } catch (e) {
    return 0;
  }
}

function applyHeading(paragraph, indentLevel) {
  try {
    const headingType = getHeadingType(indentLevel);
    if (!headingType) return false;
    
    paragraph.setHeading(headingType);
    paragraph.setIndentFirstLine(0);
    paragraph.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
    return true;
  } catch (e) {
    console.warn("Couldn't apply heading to:", paragraph.getText());
    return false;
  }
}

function getHeadingType(indentLevel) {
  const clampedLevel = Math.min(Math.max(indentLevel, 3), 9);
  return [
    null, null, null,
    DocumentApp.ParagraphHeading.SUBTITLE,    // 3
    DocumentApp.ParagraphHeading.HEADING1,   // 4
    DocumentApp.ParagraphHeading.HEADING2,   // 5
    DocumentApp.ParagraphHeading.HEADING3,   // 6
    DocumentApp.ParagraphHeading.HEADING4,   // 7
    DocumentApp.ParagraphHeading.HEADING5,   // 8
    DocumentApp.ParagraphHeading.HEADING6    // 9+
  ][clampedLevel];
}

function showCompletionAlert(changesMade) {
  const ui = DocumentApp.getUi();
  if (changesMade) {
    ui.alert('✅ Success', 'Converted indented text to headings!', ui.ButtonSet.OK);
  } else {
    ui.alert('ℹ️ Info', 'No qualifying indented text found (needs 3+ indents on non-empty paragraphs).', ui.ButtonSet.OK);
  }
}

/**
 * Adds a required doGet.
 */
function doGet() {
  return ContentService.createTextOutput('I just successfully handled your GET request.');
}
