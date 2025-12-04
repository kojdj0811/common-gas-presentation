/**
 * @OnlyCurrentDoc Limits the script to only accessing the current presentation.
 */

/**
 * Create a open translate menu item.
 * @param {Event} event The open event.
 */
function onOpen(event) {
  SlidesApp.getUi()
    .createAddonMenu()
    .addItem("Open Translate", "showSidebar")
    .addToUi();
}

/**
 * Open the Add-on upon install.
 * @param {Event} event The install event.
 */
function onInstall(event) {
  onOpen(event);
}

/**
 * Opens a sidebar in the document containing the add-on's user interface.
 */
function showSidebar() {
  const ui =
    HtmlService.createHtmlOutputFromFile("sidebar").setTitle("Translate");
  SlidesApp.getUi().showSidebar(ui);
}

/**
 * Recursively gets child text elements a list of elements.
 * @param {PageElement[]} elements The elements to get text from.
 * @return {Text[]} An array of text elements.
 */
function getElementTexts(elements) {
  let texts = [];
  for (const element of elements) {
    switch (element.getPageElementType()) {
      case SlidesApp.PageElementType.GROUP:
        for (const child of element.asGroup().getChildren()) {
          texts = texts.concat(getElementTexts(child));
        }
        break;
      case SlidesApp.PageElementType.TABLE: {
        const table = element.asTable();
        for (let r = 0; r < table.getNumRows(); ++r) {
          for (let c = 0; c < table.getNumColumns(); ++c) {
            texts.push(table.getCell(r, c).getText());
          }
        }
        break;
      }
      case SlidesApp.PageElementType.SHAPE:
        texts.push(element.asShape().getText());
        break;
    }
  }
  return texts;
}

/**
 * Translates selected slide elements to the target language using Apps Script's Language service.
 *
 * @param {string} targetLanguage The two-letter short form for the target language. (ISO 639-1)
 * @return {number} The number of elements translated.
 */
function translateSelectedElements(targetLanguage) {
  // Get selected elements.
  const selection = SlidesApp.getActivePresentation().getSelection();
  const selectionType = selection.getSelectionType();
  let texts = [];
  switch (selectionType) {
    case SlidesApp.SelectionType.PAGE:
      for (const page of selection.getPageRange().getPages()) {
        texts = texts.concat(getElementTexts(page.getPageElements()));
      }
      break;
    case SlidesApp.SelectionType.PAGE_ELEMENT: {
      const pageElements = selection.getPageElementRange().getPageElements();
      texts = texts.concat(getElementTexts(pageElements));
      break;
    }
    case SlidesApp.SelectionType.TABLE_CELL:
      for (const cell of selection.getTableCellRange().getTableCells()) {
        texts.push(cell.getText());
      }
      break;
    case SlidesApp.SelectionType.TEXT:
      for (const element of selection.getPageElementRange().getPageElements()) {
        texts.push(element.asShape().getText());
      }
      break;
  }

  // Translate all elements in-place.
  for (const text of texts) {
    text.setText(
      LanguageApp.translate(text.asRenderedString(), "", targetLanguage),
    );
  }

  return texts.length;
}