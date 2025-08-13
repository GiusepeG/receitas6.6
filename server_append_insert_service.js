/**
 * @file server_append_insert_service.js
 * @description Service for handling all text insertion and appending logic.
 */
const AppendInsertService = (() => {
  'use strict';

  // =================================================================
  // PRIVATE HELPER FUNCTIONS
  // =================================================================

  /**
   * Inserts text at the current cursor position.
   * @private
   */
  function _insertAtCursor(textToInsert) {
    const cursor = DocumentApp.getActiveDocument().getCursor();
    if (cursor) {
      cursor.insertText(textToInsert + '\n');
      return true;
    }
    // Fallback: append to body if cursor is not available
    DocumentApp.getActiveDocument().getBody().appendParagraph(textToInsert);
    DocumentApp.getUi().alert('Cursor not found. Text was appended to the end of the document.');
    return false;
  }

  /**
   * Appends text to the end of the document.
   * @private
   */
  function _appendToEnd(textToInsert) {
    DocumentApp.getActiveDocument().getBody().appendParagraph(textToInsert);
    return true;
  }


  // =================================================================
  // PUBLIC API
  // =================================================================
  const publicApi = {};

  /**
   * Main entry point for inserting text. If text is selected, it opens a dialog.
   * Otherwise, it inserts the text directly at the cursor.
   * @param {string} text - The text to insert.
   * @param {boolean} shouldFormat - Whether to re-format the document after insertion.
   * @returns {string} The status of the operation.
   */
  publicApi.insertText = (text, shouldFormat) => {
    if (!text) {
      if (shouldFormat) DocumentService.executeFormat();
      return "FORMATTED_ONLY";
    }

    const selection = DocumentApp.getActiveDocument().getSelection();
    if (selection && selection.getRangeElements().length > 0) {
      // Logic to show a dialog is client-side concern.
      // This function should just return the fact that there is a selection.
      // The client can then decide to show a dialog.
      // For now, to keep server logic, we can throw an error.
      throw new Error("Text is selected. This operation requires a cursor position, not a selection.");
    }

    _insertAtCursor(text);
    if (shouldFormat) DocumentService.executeFormat();
    return "INSERTED_DIRECTLY";
  };

  /**
   * Main entry point for appending text. Always adds text to the end of the document.
   * @param {string} text - The text to append.
   * @param {boolean} shouldFormat - Whether to re-format the document after appending.
   * @returns {string} The status of the operation.
   */
  publicApi.appendText = (text, shouldFormat) => {
    if (!text) {
      if (shouldFormat) DocumentService.executeFormat();
      return "FORMATTED_ONLY";
    }
    _appendToEnd(text);
    if (shouldFormat) DocumentService.executeFormat();
    return "APPENDED_DIRECTLY";
  };

  /**
   * Validates the document structure and determines the correct append action.
   * This is called by the client to decide whether to append directly or show a choice dialog.
   * @returns {Object} An object describing the required action ('append_direct' or 'show_dialog').
   */
  publicApi.validateForAppend = () => {
    try {
      const uniqueH1s = DocumentService.getUniqueHeadline1s();

      if (uniqueH1s.length <= 1) {
        // If 0 or 1 H1s, we can append directly without ambiguity.
        return { success: true, action: 'append_direct' };
      } else {
        // If multiple H1s, the client needs to show a dialog.
        return { success: true, action: 'show_dialog', uniqueHeadline1s: uniqueH1s };
      }
    } catch (e) {
      return { success: false, message: `Error validating document: ${e.message}` };
    }
  };

  /**
   * Appends text under a specific H1 section. This is called after the user
   * makes a selection in the multi-H1 dialog.
   * @param {string} selectedH1 - The H1 section chosen by the user.
   * @param {string} textToAppend - The text to append.
   * @param {boolean} shouldFormat - Whether to re-format the document.
   * @returns {Object} The result of the operation.
   */
  publicApi.appendToSelectedH1 = (selectedH1, textToAppend, shouldFormat) => {
    try {
      // This is a complex operation. Instead of trying to manually edit the text,
      // we can leverage the DocumentService's understanding of document structure.
      // The cleanest way is to find the LAST H2 under the selected H1 and insert after it.
      // However, a simpler, robust approach is to just append to the document,
      // and then rely on the DocumentService's formatting to group it correctly.
      // The `createOrUpdateBodyHeadline2` can be used to add content.
      // Let's assume we want to add the text to a new H2 called "Anotação".

      const newHeadline2 = "Anotação Adicional"; // Or we could pass this in.
      DocumentService.processAndUpdateDocument(textToAppend, selectedH1, newHeadline2, null);

      // The above call already handles updating the doc and re-formatting.

      return { success: true, message: `Text appended to ${selectedH1}.` };

    } catch (e) {
      return { success: false, message: `Failed to append text: ${e.message}` };
    }
  };

  return publicApi;

})();

// =================================================================
// GLOBAL STUBS for Client-Side Calls
// =================================================================
// These functions make the service methods callable from the client-side `google.script.run`.
// They parse the JSON string that often comes from the client.

function insertTextAndFormat(jsonObj) {
  const obj = JSON.parse(jsonObj);
  return AppendInsertService.insertText(obj.text, obj.shouldFormat);
}

function appendTextAndFormat(jsonObj) {
  const obj = JSON.parse(jsonObj);
  return AppendInsertService.appendText(obj.text, obj.shouldFormat);
}

function validateAndExecuteAppend() {
  return AppendInsertService.validateForAppend();
}

function executeAppendToSelectedH1(selectedH1, textToAppend, shouldFormat) {
  return AppendInsertService.appendToSelectedH1(selectedH1, textToAppend, shouldFormat);
}
