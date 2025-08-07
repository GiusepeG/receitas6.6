const AIService = {
  /**
   * Chama o modelo Gemini 2.0 Flash-Lite (Google) via REST.
   * @param {string} finalPrompt – prompt completo montado
   * @param {boolean} isTest – modo stub (devolve texto fixo) - opcional
   * @return {string} – resposta do modelo
   */
  callAIModel: function(finalPrompt, isTest = false) {
    if (isTest) return "<texto da IA (stub)>";

    let apiKey;
    try {
      apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
      if (!apiKey) {
        try {
          apiKey = DriveApp.getFilesByName('GEMINI_KEY.txt').next().getBlob().getDataAsString().trim();
        } catch (e) {
          Logger.log("❌ Erro ao ler arquivo GEMINI_KEY.txt:", e.message);
          throw new Error("Chave da API Gemini não encontrada. Configure a propriedade GEMINI_API_KEY nas Script Properties ou crie um arquivo GEMINI_KEY.txt no Drive.");
        }
      }
    } catch (e) {
      Logger.log("❌ Erro ao obter chave da API:", e.message);
      throw new Error("Chave da API Gemini não configurada.");
    }

    const modelId  = "gemini-2.5-flash-lite";
    const url      = `https://generativelanguage.googleapis.com/v1beta/models/${modelId}:generateContent?key=${apiKey}`;

    const requestBody = {
      contents: [{ parts: [{ text: finalPrompt }] }],
      generationConfig: {
        temperature: 0,
        maxOutputTokens: 3500
      }
    };

    const response = UrlFetchApp.fetch(url, {
      method: "POST",
      contentType: "application/json",
      payload: JSON.stringify(requestBody)
    });

    const json = JSON.parse(response.getContentText());
    return json.candidates[0].content.parts[0].text;
  },

  /**
   * Processa pares selecionados com IA.
   * @param {Array<Object>} selectedPairs - Pares de H1/H2 selecionados.
   * @returns {Object} - Resultado da operação.
   */
  processAISelectedPairsWithPrompt: function(selectedPairs) {
    Logger.log(`[AI] ✅ Processando pares selecionados com IA: ${selectedPairs.length}`);
    const doc = DocumentApp.getActiveDocument();
    const documentManager = new DocumentService.DocumentHeadlineManager();
    let processedCount = 0;
    const totalPairs = selectedPairs.length;

    selectedPairs.forEach((pair, index) => {
      try {
        const finalPrompt = `${pair.promptContent}\n\n${pair.headline1}\n${pair.toHeadline2}\n${pair.clinicalContent}`;
        const aiResponse = this.callAIModel(finalPrompt);
        documentManager.createOrUpdateBodyHeadline2(aiResponse, pair.headline1, pair.toHeadline2, pair.fromHeadline2);
        processedCount++;
      } catch (error) {
        console.error(`[AI] ❌ Erro ao processar par ${index + 1}:`, error);
      }
    });

    if (processedCount > 0) {
      try {
        const newBodyText = documentManager.getText();
        doc.getBody().setText(newBodyText);
        DocumentService.executeFormat();
      } catch (error) {
        console.error('[AI] ❌ Erro ao atualizar documento ou formatar:', error);
      }
    }
    return { success: true };
  }
};
