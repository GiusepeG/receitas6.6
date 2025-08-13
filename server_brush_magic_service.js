/**
 * @file server_brush_magic_service.js
 * @description Service for orchestrating the 'Brush' and 'Magic' AI features.
 */
const BrushMagicService = (() => {
  'use strict';

  // =================================================================
  // PRIVATE HELPER FUNCTIONS
  // =================================================================

  /**
   * Retrieves the configured pairs for the brush feature from script properties.
   * @private
   * @returns {Array<Object>} An array of {fromHeadline2, toHeadline2} pairs.
   */
  function _getBrushPairs() {
    const properties = PropertiesService.getScriptProperties();
    let pairsString = properties.getProperty('brushButtonPairs');
    if (!pairsString) {
      // If not configured, set up a default and use that.
      const defaultPairs = "Prontuário Médico, Prontuário Médico; Prontuário Médico, Prescrição de Óculos; Laudo de Mapeamento de Retina, Laudo de Mapeamento de Retina";
      properties.setProperty('brushButtonPairs', defaultPairs);
      pairsString = defaultPairs;
      console.log('✅ Default brush pairs have been configured.');
    }

    return pairsString
      .split(';')
      .map(pair => pair.trim())
      .filter(Boolean)
      .map(pair => {
        const [from, to] = pair.split(',').map(item => item.trim());
        return { fromHeadline2: from, toHeadline2: to };
      });
  }

  // =================================================================
  // PUBLIC API
  // =================================================================
  const publicApi = {};

  /**
   * Executes the entire "Brush" workflow. It validates the document, finds applicable
   * AI prompts based on configured pairs, and then calls the AI service to process them.
   * @returns {Object} A result object indicating success or failure.
   */
  publicApi.executeBrushFlow = () => {
    try {
      // Step 1: Validate document state using DocumentService
      const uniqueH1s = DocumentService.getUniqueHeadline1s();
      if (uniqueH1s.length === 0) {
        return { success: false, message: 'No patient (H1) found in the document.' };
      }
      if (uniqueH1s.length > 1) {
        return { success: false, message: `Multiple patients (${uniqueH1s.length}) found. The Brush feature currently supports only one patient per document.` };
      }
      const headline1 = uniqueH1s[0];

      // Step 2: Get configured brush pairs and all available prompts
      const brushPairs = _getBrushPairs();
      const allPrompts = JSON.parse(DataService.getSidebarData()).allPrompts;

      if (!allPrompts || Object.keys(allPrompts).length === 0) {
        return { success: false, message: 'No AI prompts found in the system.' };
      }

      // Step 3: Find which of the configured brush pairs are valid for the current document
      const docContent = DocumentService.getDocumentHeadlinePairsAndContent();
      const validPrompts = [];

      brushPairs.forEach(pair => {
        const hasFromH2 = docContent.some(item => item.headline1 === headline1 && item.headline2 === pair.fromHeadline2);
        if (!hasFromH2) return; // Skip if the source H2 doesn't exist for this patient

        const promptKey = Object.keys(allPrompts).find(key => {
            const p = allPrompts[key];
            return p.fromHeadline2 === pair.fromHeadline2 && p.toHeadline2 === pair.toHeadline2;
        });

        if (promptKey) {
            validPrompts.push({
                fromHeadline2: pair.fromHeadline2,
                toHeadline2: pair.toHeadline2,
                promptContent: allPrompts[promptKey].content,
                optionText: `Improve ${pair.toHeadline2}` // Placeholder text
            });
        }
      });

      if (validPrompts.length === 0) {
        return { success: false, message: 'No valid actions found for the current document state. Check H2 sections and prompt configurations.' };
      }

      // Step 4: Execute AI processing by calling AIService
      const processingData = {
        headline1s: [{ headline1: headline1 }],
        prompts: validPrompts,
        requireAugmentedContext: true // Brush feature always uses augmented context
      };

      const newBodyText = AIService.processPrompts(processingData);

      // Step 5: Update the document
      DocumentApp.getActiveDocument().getBody().setText(newBodyText);
      DocumentService.executeFormat();

      return { success: true, message: `${validPrompts.length} section(s) were successfully improved.` };

    } catch (e) {
      console.error(`❌ Brush Flow Error: ${e.message}`, e.stack);
      return { success: false, message: `An unexpected error occurred: ${e.message}` };
    }
  };

  /**
   * Displays the "Magic Modal" for more complex AI interactions.
   */
  publicApi.showMagicModal = () => {
    try {
        const allPrompts = JSON.parse(DataService.getSidebarData()).allPrompts;
        const promptsArray = Object.keys(allPrompts).map(key => ({
            optionText: key,
            ...allPrompts[key]
        }));

        const docContent = DocumentService.getDocumentHeadlinePairsAndContent();
        const eligibleH1s = {};
        docContent.forEach(item => {
            if (!eligibleH1s[item.headline1]) eligibleH1s[item.headline1] = { headline1: item.headline1, eligibleH2s: [] };
            eligibleH1s[item.headline1].eligibleH2s.push(item.headline2);
        });

        const template = HtmlService.createTemplateFromFile('dialog_ai_magic');
        template.promptData = JSON.stringify(promptsArray);
        template.uniqueHeadline1sData = JSON.stringify(Object.values(eligibleH1s));

        const html = template.evaluate().setWidth(900).setHeight(700);
        DocumentApp.getUi().showModalDialog(html, 'AI Magic Processing');

    } catch (e) {
        DocumentApp.getUi().alert(`Error showing Magic Modal: ${e.message}`);
    }
  }

  return publicApi;

})();


// =================================================================
// GLOBAL STUBS for Client-Side Calls
// =================================================================

function executeCompleteBrushFlowWithPairs() {
    return BrushMagicService.executeBrushFlow();
}

function callMagicModal() {
    return BrushMagicService.showMagicModal();
}