/**
 * @file server_ai_controller.js
 * @description Controller for handling UI interactions related to AI features.
 * This file contains functions that are directly called by the client-side JavaScript.
 */

/**
 * Displays a modal dialog to the user with a list of patients and AI prompts.
 * @param {Array<Object>} prompts - The list of available AI prompts.
 */
function showModalDocumentStructure(prompts) {
  const allPairs = getDocumentHeadlinePairs();
  const template = HtmlService.createTemplateFromFile('dialog_ai_magic');

  template.promptData = JSON.stringify(prompts);
  template.allPairsData = JSON.stringify(allPairs);

  const title = prompts.map(p => p.optionText).join(' & ');
  const html = template.evaluate().setWidth(500).setHeight(600);

  DocumentApp.getUi().showModalDialog(html, title);
}

/**
 * Retrieves all H1/H2 pairs and their content from the document using the DocumentService.
 * This function is called by the UI to populate the patient/section selection dialog.
 * @returns {Array<Object>} An array of objects, each containing headline1, headline2, and content.
 */
function getDocumentHeadlinePairs() {
  try {
    // Delegate directly to the DocumentService to fetch data.
    const pairs = DocumentService.getDocumentHeadlinePairsAndContent();
    if (!pairs || pairs.length === 0) {
      console.warn('[AIController] No H1/H2 pairs were returned by DocumentService.');
    }
    return pairs;
  } catch (error) {
    console.error('[AIController] Fatal error in getDocumentHeadlinePairs:', error);
    throw error; // Re-throw to be caught by client's onFailure handler.
  }
}

/**
 * Main entry point for processing AI tasks selected by the user in the dialog.
 * It delegates the heavy lifting to AIService and then updates the document.
 * @param {Object} processingData - The data submitted from the modal dialog.
 * @returns {Object} A success or failure object.
 */
function processSequentialPromptsForPairs(processingData) {
  try {
    // 1. Delegate the core AI processing to the AIService.
    // AIService will return the complete, updated text for the document body.
    const newBodyText = AIService.processPrompts(processingData);

    // 2. Update the document with the new text.
    const doc = DocumentApp.getActiveDocument();
    doc.getBody().setText(newBodyText);

    // 3. Apply standard formatting to the entire document.
    DocumentService.executeFormat();

    return { success: true };

  } catch (error) {
    console.error("❌ Fatal error during AI processing:", error.message, error.stack);
    return { success: false, message: error.message };
  }
}
