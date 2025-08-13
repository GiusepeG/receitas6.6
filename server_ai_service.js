/**
 * @file server_ai_service.js
 * @description Encapsulated service for handling all AI model interactions, including prompt construction,
 * API calls, and processing responses.
 */
const AIService = (() => {
  'use strict';

  // =================================================================
  // PRIVATE HELPER FUNCTIONS
  // =================================================================

  /**
   * Calls the configured Generative AI model via REST API.
   * @private
   * @param {string} finalPrompt - The complete prompt to send to the model.
   * @param {boolean} [isTest=false] - If true, returns a stubbed response without calling the API.
   * @returns {string} The text response from the AI model.
   * @throws {Error} If the API key is not configured or if the API call fails.
   */
  function _callAIModel(finalPrompt, isTest = false) {
    if (isTest) return "<texto da IA (stub)>";

    let apiKey;
    try {
      apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
      if (!apiKey) {
        throw new Error("Chave da API Gemini não encontrada nas Script Properties.");
      }
    } catch (e) {
      throw new Error("Falha ao acessar Script Properties para a chave da API Gemini.");
    }

    const modelId = "gemini-1.5-flash-latest"; // Recommended Flash model
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${modelId}:generateContent?key=${apiKey}`;

    const requestBody = {
      contents: [{ parts: [{ text: finalPrompt }] }],
      generationConfig: {
        temperature: 0,
        maxOutputTokens: 8192,
      }
    };

    const options = {
      method: "POST",
      contentType: "application/json",
      payload: JSON.stringify(requestBody),
      muteHttpExceptions: true // Important for parsing error responses
    };

    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    if (responseCode !== 200) {
      throw new Error(`API Call Failed with code ${responseCode}: ${responseText}`);
    }

    const json = JSON.parse(responseText);

    if (!json.candidates || !json.candidates[0] || !json.candidates[0].content || !json.candidates[0].content.parts || !json.candidates[0].content.parts[0].text) {
        throw new Error(`Resposta da API inválida: ${responseText}`);
    }

    return json.candidates[0].content.parts[0].text;
  }

  /**
   * Logs the details of an AI interaction to a specified Google Doc for auditing.
   * @private
   */
  function _logInteraction(promptContent, headline1, fromHeadline2, toHeadline2, clinicalContent, newContent, counter) {
    try {
      const logDocId = PropertiesService.getScriptProperties().getProperty('logDocFileId');
      if (!logDocId) {
        console.warn("Property 'logDocFileId' is not set. Skipping interaction log.");
        return;
      }

      const logDoc = DocumentApp.openById(logDocId);
      const body = logDoc.getBody();
      const iterationStr = String(counter).padStart(2, '0');

      body.appendParagraph(`DateTime: ${new Date().toLocaleString()}, ${headline1}`).setHeading(DocumentApp.ParagraphHeading.HEADING2);
      body.appendParagraph(`Iteration: ${iterationStr}, From: ${fromHeadline2}, To: ${toHeadline2}`).setHeading(DocumentApp.ParagraphHeading.HEADING3);
      body.appendParagraph('--- Clinical Content ---').setHeading(DocumentApp.ParagraphHeading.HEADING4);
      body.appendParagraph(clinicalContent);
      body.appendParagraph('--- AI Generated Content ---').setHeading(DocumentApp.ParagraphHeading.HEADING4);
      body.appendParagraph(newContent);
      body.appendParagraph('--- Prompt ---').setHeading(DocumentApp.ParagraphHeading.HEADING4);
      body.appendParagraph(promptContent);

    } catch (e) {
      console.error(`Failed to log interaction to Google Doc. Error: ${e.toString()}`);
    }
  }


  // =================================================================
  // PUBLIC API
  // =================================================================
  const publicApi = {};

  /**
   * Processes a batch of prompts for a list of patients (H1s), updates the document structure in memory,
   * and returns the final generated text for the document body.
   * @param {Object} processingData - The data object from the client.
   * @param {Array} processingData.headline1s - List of patient H1s to process.
   * @param {Array} processingData.prompts - List of AI prompts/tasks to perform.
   * @param {boolean} processingData.requireAugmentedContext - Whether to use full patient context.
   * @returns {string} The complete, updated text for the document body.
   * @throws {Error} If input data is invalid.
   */
  publicApi.processPrompts = (processingData) => {
    const { headline1s, prompts, requireAugmentedContext } = processingData;

    if (!headline1s || headline1s.length === 0 || !prompts || prompts.length === 0) {
      throw new Error("Dados de processamento inválidos. Headline1s ou prompts ausentes.");
    }

    // Use the DocumentService.Manager for stateful batch processing.
    const docManager = new DocumentService.Manager();
    let iterationCounter = 0;

    prompts.forEach(prompt => {
      const { fromHeadline2, toHeadline2, promptContent, optionText } = prompt;

      headline1s.forEach(headline1Item => {
        const { headline1 } = headline1Item;
        let clinicalContent;

        if (requireAugmentedContext) {
          clinicalContent = docManager.getAugmentedContentForH1(headline1);
        } else {
          clinicalContent = docManager.getContentForPair(headline1, fromHeadline2);
        }

        if (clinicalContent === null || clinicalContent.trim() === '') {
          console.warn(`Skipping patient "${headline1}" for prompt "${optionText}" due to missing content.`);
          return;
        }

        const finalPrompt = `${promptContent}\n\n${headline1}\n\n${clinicalContent}`;

        const newContent = _callAIModel(finalPrompt);

        iterationCounter++;
        _logInteraction(promptContent, headline1, fromHeadline2, toHeadline2, clinicalContent, newContent, iterationCounter);

        // Update the document structure in memory.
        docManager.createOrUpdateBodyHeadline2(newContent, headline1, toHeadline2, fromHeadline2);
      });
    });

    // Return the final text to be written by the controller.
    docManager._sortAndGroupHeadlines();
    return docManager.getText();
  };

  /**
   * Creates a new Google Doc to be used for logging AI interactions.
   * @returns {string|null} The ID of the new document, or null on failure.
   */
  publicApi.createLogDocument = () => {
    try {
      const newDoc = DocumentApp.create(`AI Interaction Log - ${new Date().toLocaleDateString()}`);
      const newDocId = newDoc.getId();

      const body = newDoc.getBody();
      body.appendParagraph('AI Interaction Log').setHeading(DocumentApp.ParagraphHeading.TITLE);
      body.appendParagraph(`Created on: ${new Date().toLocaleString()}`);
      body.appendParagraph('');

      // Automatically set this new doc ID as the active log file
      PropertiesService.getScriptProperties().setProperty('logDocFileId', newDocId);

      console.log(`✅ New log document created with ID: ${newDocId}`);
      DocumentApp.getUi().alert(`New log document created. ID: ${newDocId}. This has been set as the active log file.`);

      return newDocId;
    } catch (e) {
      console.error(`❌ Failed to create log document: ${e.message}`);
      DocumentApp.getUi().alert(`Error creating log document: ${e.message}`);
      return null;
    }
  };

  return publicApi;

})();

function createLogDocument() {
    return AIService.createLogDocument();
}
