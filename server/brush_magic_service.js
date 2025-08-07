const BrushMagicService = {
  buildDocumentStructureWithH1Only: function(text, headline1Rules, headline2Rules) {
    const lines = text.split('\n');
    let currentHeadline1 = null;
    let documentStructure = [];
    let foundFirstStructure = false;
    let awaitingHeadline2 = false;

    lines.forEach((line, idx) => {
      if (line.trim() === "") return;

      if (headline1Rules.some(rule => rule.condition(line))) {
        if (awaitingHeadline2 && currentHeadline1) {
          documentStructure.push({
            headline1: currentHeadline1,
            headline2: "Documento Indefinido",
            text: ''
          });
        }

        currentHeadline1 = line;
        awaitingHeadline2 = true;
      }
      else if (headline2Rules.some(rule => rule.condition(line))) {
        documentStructure.push({
          headline1: currentHeadline1,
          headline2: line,
          text: ''
        });
        foundFirstStructure = true;
        awaitingHeadline2 = false;
      }
      else if (awaitingHeadline2) {
        documentStructure.push({
          headline1: currentHeadline1,
          headline2: "Título",
          text: line + '\n'
        });
        foundFirstStructure = true;
        awaitingHeadline2 = false;
      }
      else if (documentStructure.length > 0) {
        documentStructure[documentStructure.length - 1].text += line + '\n';
      }
    });

    if (awaitingHeadline2 && currentHeadline1) {
      documentStructure.push({
        headline1: currentHeadline1,
        headline2: "Documento Indefinido",
        text: ''
      });
    }

    return documentStructure;
  }
};

  validateSingleHeadline1: function() {
    try {
      const doc = DocumentApp.getActiveDocument();
      const bodyText = doc.getBody().getText();
      const formattingRules = DocumentService.rules.getFormattingRules();
      const placeholders = DocumentService.rules.getRulesForPlaceholders();
      const heading1Rules = formattingRules.filter(rule => rule.heading === DocumentApp.ParagraphHeading.HEADING1);
      const heading2Rules = formattingRules.filter(rule => rule.heading === DocumentApp.ParagraphHeading.HEADING2);
      const textManipulator = new DocumentService.TextManipulator(bodyText, placeholders, heading1Rules, heading2Rules);
      const processedText = textManipulator.splitInLines().removeEmptyLines().trimLeadingSpaces().checkAndModifyFirstLine().processLines().structureH2Blocks().getResult();
      const headlineManager = new DocumentService.DocumentHeadlineManager();
      headlineManager.bodyText = processedText;
      headlineManager.bodyStructure = this.buildDocumentStructureWithH1Only(processedText, heading1Rules, heading2Rules);
      const uniqueHeadline1s = headlineManager.getUniqueHeadline1s();
      if (uniqueHeadline1s.length === 0) {
        return { isValid: false, message: 'Nenhum Headline1 encontrado no documento. É necessário ter pelo menos um Headline1 para usar esta função.' };
      }
      if (uniqueHeadline1s.length === 1) {
        return { isValid: true, message: 'Documento válido com um único Headline1.' };
      }
      return { isValid: false, showDialog: true, uniqueHeadline1s: uniqueHeadline1s, message: `Foram encontrados ${uniqueHeadline1s.length} Headline1 diferentes no documento.` };
    } catch (error) {
      return { isValid: false, message: 'Erro ao analisar a estrutura do documento: ' + error.message };
    }
  }
};

  openAppendChoiceDialog: function(uniqueHeadline1s, textToAppend, shouldFormat) {
    try {
      const htmlTemplate = HtmlService.createTemplateFromFile('dialogs/append_multi_h1_dialog.html');
      htmlTemplate.uniqueHeadline1s = uniqueHeadline1s;
      htmlTemplate.textToAppend = textToAppend;
      htmlTemplate.shouldFormat = shouldFormat;

      const ui = DocumentApp.getUi();
      const dialog = htmlTemplate.evaluate()
        .setWidth(400)
        .setHeight(500);

      ui.showModalDialog(dialog, 'Escolher Paciente');

    } catch (error) {
      DocumentApp.getUi().alert('Erro ao abrir diálogo de seleção: ' + error.message);
    }
  }
};


  executeCompleteBrushFlow: function(brushData) {
    const { fixedFromHeadline2 } = brushData;
    try {
      const properties = PropertiesService.getScriptProperties();
      const toHeadline2sString = properties.getProperty('BRUSH_TO_HEADLINE2S');
      if (!toHeadline2sString) {
        return { success: false, action: 'no_config', message: 'Nenhum toHeadline2 configurado. Configure os toHeadline2s nas propriedades do script primeiro.' };
      }
      const toHeadline2sArray = toHeadline2sString.split(',').map(item => item.trim()).filter(item => item.length > 0);
      if (toHeadline2sArray.length === 0) {
        return { success: false, action: 'no_config', message: 'Configuração de toHeadline2s está vazia.' };
      }
      const documentId = properties.getProperty('docData');
      const sheetId = properties.getProperty('sheetData');
      if (!documentId || !sheetId) {
        return { success: false, action: 'no_config', message: 'Configuração de documentId ou sheetId não encontrada.' };
      }
      const dataFetcher = new DataService.GoogleDriveDataFetcher(documentId, sheetId);
      const data = dataFetcher.extractAllDataFromSources();
      const allPrompts = data.allPrompts;
      if (!allPrompts || Object.keys(allPrompts).length === 0) {
        return { success: false, action: 'no_prompts', message: 'Nenhum prompt encontrado no sistema.' };
      }
      const doc = DocumentApp.getActiveDocument();
      const bodyText = doc.getBody().getText();
      const formattingRules = DocumentService.rules.getFormattingRules();
      const placeholders = DocumentService.rules.getRulesForPlaceholders();
      const heading1Rules = formattingRules.filter(rule => rule.heading === DocumentApp.ParagraphHeading.HEADING1);
      const heading2Rules = formattingRules.filter(rule => rule.heading === DocumentApp.ParagraphHeading.HEADING2);
      const textManipulator = new DocumentService.TextManipulator(bodyText, placeholders, heading1Rules, heading2Rules);
      const processedText = textManipulator.splitInLines().removeEmptyLines().trimLeadingSpaces().checkAndModifyFirstLine().processLines().structureH2Blocks().getResult();
      const headlineManager = new DocumentService.DocumentHeadlineManager();
      headlineManager.bodyText = processedText;
      headlineManager.bodyStructure = this.buildDocumentStructureWithH1Only(processedText, heading1Rules, heading2Rules);
      const uniqueHeadline1s = headlineManager.getUniqueHeadline1s();
      if (uniqueHeadline1s.length === 0) {
        return { success: false, action: 'no_h1', message: 'Nenhum Headline1 encontrado no documento. É necessário ter pelo menos um paciente para usar esta função.' };
      }
      if (uniqueHeadline1s.length > 1) {
        return { success: false, action: 'multiple_h1', message: `Encontrados ${uniqueHeadline1s.length} Headline1s no documento. O processamento de múltiplos pacientes ainda não está implementado para o botão Brush.` };
      }
      const headline1 = uniqueHeadline1s[0];
      const hasProtuarioPair = headlineManager.bodyStructure.some(item => item.headline1 === headline1 && item.headline2 === "Prontuário Médico");
      if (!hasProtuarioPair) {
        return { success: false, action: 'no_prontuario', message: 'Não foi encontrado o Headline2 "Prontuário Médico" no documento. Esta função requer a seção "Prontuário Médico" para funcionar.' };
      }
      const validPrompts = [];
      const pairs = [];
      toHeadline2sArray.forEach(toHeadline2 => {
        const promptKey = Object.keys(allPrompts).find(key => {
          const prompt = allPrompts[key];
          const normalizedFromPrompt = (prompt.fromHeadline2 || '').normalize('NFC');
          const normalizedToPrompt = (prompt.toHeadline2 || '').normalize('NFC');
          const normalizedFromFixed = fixedFromHeadline2.normalize('NFC');
          const normalizedToSearch = toHeadline2.normalize('NFC');
          return normalizedFromPrompt === normalizedFromFixed && normalizedToPrompt === normalizedToSearch;
        });
        if (promptKey) {
          const promptContent = allPrompts[promptKey].content;
          validPrompts.push({ fromHeadline2: fixedFromHeadline2, toHeadline2: toHeadline2, promptContent: promptContent, optionText: `Melhorar ${toHeadline2}` });
          pairs.push({ headline1: headline1, headline2: toHeadline2 });
        }
      });
      if (validPrompts.length === 0) {
        return { success: false, action: 'no_prompts', message: 'Nenhum prompt válido encontrado para os toHeadline2s configurados. Verifique a configuração dos prompts.' };
      }
      const headline1s = [{ headline1: headline1 }];
      const processingData = { headline1s: headline1s, prompts: validPrompts, requireAugmentedContext: true };
      let result;
      try {
        result = AIService.processSequentialPromptsForPairs(processingData);
      } catch (error) {
        return { success: false, action: 'processing_error', message: 'Erro ao chamar função de processamento: ' + error.message };
      }
      if (result.success) {
        return { success: true, action: 'processed', processedCount: validPrompts.length, message: `Processamento concluído com sucesso! ${validPrompts.length} seção(ões) melhorada(s).` };
      } else {
        return { success: false, action: 'processing_error', message: `Erro no processamento: ${result.message || 'Erro desconhecido'}` };
      }
    } catch (error) {
      return { success: false, action: 'error', message: 'Erro durante o processamento brush: ' + error.message };
    }
  }
};


  validateHeadline1sForBrush: function() {
    try {
      const doc = DocumentApp.getActiveDocument();
      const bodyText = doc.getBody().getText();
      const formattingRules = DocumentService.rules.getFormattingRules();
      const placeholders = DocumentService.rules.getRulesForPlaceholders();
      const heading1Rules = formattingRules.filter(rule => rule.heading === DocumentApp.ParagraphHeading.HEADING1);
      const heading2Rules = formattingRules.filter(rule => rule.heading === DocumentApp.ParagraphHeading.HEADING2);
      const textManipulator = new DocumentService.TextManipulator(bodyText, placeholders, heading1Rules, heading2Rules);
      const processedText = textManipulator.splitInLines().removeEmptyLines().trimLeadingSpaces().checkAndModifyFirstLine().processLines().structureH2Blocks().getResult();
      const headlineManager = new DocumentService.DocumentHeadlineManager();
      headlineManager.bodyText = processedText;
      headlineManager.bodyStructure = this.buildDocumentStructureWithH1Only(processedText, heading1Rules, heading2Rules);
      const uniqueHeadline1s = headlineManager.getUniqueHeadline1s();
      let hasProntuario = false;
      if (uniqueHeadline1s.length === 1) {
        hasProntuario = headlineManager.bodyStructure.some(item => item.headline1 === uniqueHeadline1s[0] && item.headline2 === "Prontuário Médico");
      }
      return { count: uniqueHeadline1s.length, uniqueHeadline1s: uniqueHeadline1s, hasProntuario: hasProntuario, message: `${uniqueHeadline1s.length} Headline1(s) encontrado(s) no documento.` };
    } catch (error) {
      throw new Error('Erro ao validar Headline1s do documento: ' + error.message);
    }
  }
};

  getBrushToHeadline2s: function() {
    try {
      const properties = PropertiesService.getScriptProperties();
      const toHeadline2sString = properties.getProperty('BRUSH_TO_HEADLINE2S');

      if (!toHeadline2sString) {
        console.log('📝 Nenhum toHeadline2 configurado no PropertiesService');
        return [];
      }

      const toHeadline2sArray = toHeadline2sString
        .split(',')
        .map(item => item.trim())
        .filter(item => item.length > 0);

      console.log('📝 toHeadline2s obtidos do PropertiesService:', toHeadline2sArray);
      return toHeadline2sArray;

    } catch (error) {
      console.error('❌ Erro ao obter toHeadline2s do PropertiesService:', error);
      throw new Error('Erro ao obter configuração de toHeadline2s: ' + error.message);
    }
  },

  saveBrushToHeadline2s: function(toHeadline2sArray) {
    try {
      const properties = PropertiesService.getScriptProperties();
      const toHeadline2sString = toHeadline2sArray.join(', ');

      properties.setProperty('BRUSH_TO_HEADLINE2S', toHeadline2sString);
      console.log('✅ toHeadline2s salvos no PropertiesService:', toHeadline2sString);

      return true;
    } catch (error) {
      console.error('❌ Erro ao salvar toHeadline2s no PropertiesService:', error);
      throw new Error('Erro ao salvar configuração de toHeadline2s: ' + error.message);
    }
  }
};

  executeCompleteBrushFlowWithPairs: function() {
    try {
      const properties = PropertiesService.getScriptProperties();
      let brushButtonPairsString = properties.getProperty('brushButtonPairs');
      if (!brushButtonPairsString) {
        const setupResult = this.setupBrushButtonPairs();
        if (setupResult.success) {
          brushButtonPairsString = properties.getProperty('brushButtonPairs');
        } else {
          return { success: false, action: 'no_config', message: 'Não foi possível configurar os pares automaticamente: ' + setupResult.message };
        }
      }
      const pairsArray = brushButtonPairsString.split(';').map(pair => pair.trim()).filter(pair => pair.length > 0).map(pair => {
        const [fromHeadline2, toHeadline2] = pair.split(',').map(item => item.trim());
        return { fromHeadline2, toHeadline2 };
      });
      if (pairsArray.length === 0) {
        return { success: false, action: 'no_config', message: 'Configuração de pares está vazia ou mal formatada.' };
      }
      const documentId = properties.getProperty('docData');
      const sheetId = properties.getProperty('sheetData');
      if (!documentId || !sheetId) {
        return { success: false, action: 'no_config', message: 'Configuração de documentId ou sheetId não encontrada.' };
      }
      const dataFetcher = new DataService.GoogleDriveDataFetcher(documentId, sheetId);
      const data = dataFetcher.extractAllDataFromSources();
      const allPrompts = data.allPrompts;
      if (!allPrompts || Object.keys(allPrompts).length === 0) {
        return { success: false, action: 'no_prompts', message: 'Nenhum prompt encontrado no sistema.' };
      }
      const doc = DocumentApp.getActiveDocument();
      const bodyText = doc.getBody().getText();
      const formattingRules = DocumentService.rules.getFormattingRules();
      const placeholders = DocumentService.rules.getRulesForPlaceholders();
      const heading1Rules = formattingRules.filter(rule => rule.heading === DocumentApp.ParagraphHeading.HEADING1);
      const heading2Rules = formattingRules.filter(rule => rule.heading === DocumentApp.ParagraphHeading.HEADING2);
      const textManipulator = new DocumentService.TextManipulator(bodyText, placeholders, heading1Rules, heading2Rules);
      const processedText = textManipulator.splitInLines().removeEmptyLines().trimLeadingSpaces().checkAndModifyFirstLine().processLines().structureH2Blocks().getResult();
      const headlineManager = new DocumentService.DocumentHeadlineManager();
      headlineManager.bodyText = processedText;
      headlineManager.bodyStructure = this.buildDocumentStructureWithH1Only(processedText, heading1Rules, heading2Rules);
      const uniqueHeadline1s = headlineManager.getUniqueHeadline1s();
      if (uniqueHeadline1s.length === 0) {
        return { success: false, action: 'no_h1', message: 'Nenhum Headline1 encontrado no documento. É necessário ter pelo menos um paciente para usar esta função.' };
      }
      if (uniqueHeadline1s.length > 1) {
        return { success: false, action: 'multiple_h1', message: `Encontrados ${uniqueHeadline1s.length} Headline1s no documento. O processamento de múltiplos pacientes ainda não está implementado para o botão Brush.` };
      }
      const headline1 = uniqueHeadline1s[0];
      const validPrompts = [];
      pairsArray.forEach(pair => {
        const { fromHeadline2, toHeadline2 } = pair;
        const hasFromHeadline2 = headlineManager.bodyStructure.some(item => item.headline1 === headline1 && item.headline2 === fromHeadline2);
        if (!hasFromHeadline2) return;
        const promptKey = Object.keys(allPrompts).find(key => {
          const prompt = allPrompts[key];
          const normalizedFromPrompt = (prompt.fromHeadline2 || '').normalize('NFC');
          const normalizedToPrompt = (prompt.toHeadline2 || '').normalize('NFC');
          const normalizedFromSearch = fromHeadline2.normalize('NFC');
          const normalizedToSearch = toHeadline2.normalize('NFC');
          return normalizedFromPrompt === normalizedFromSearch && normalizedToPrompt === normalizedToSearch;
        });
        if (promptKey) {
          const promptContent = allPrompts[promptKey].content;
          validPrompts.push({ fromHeadline2: fromHeadline2, toHeadline2: toHeadline2, promptContent: promptContent, optionText: `Melhorar ${toHeadline2}` });
        }
      });
      if (validPrompts.length === 0) {
        return { success: false, action: 'no_valid_pairs', message: 'Nenhum par válido encontrado. Verifique se existem H2s compatíveis no documento e se os prompts estão configurados corretamente.' };
      }
      const headline1s = [{ headline1: headline1 }];
      const processingData = { headline1s: headline1s, prompts: validPrompts, requireAugmentedContext: true };
      let result;
      try {
        result = AIService.processSequentialPromptsForPairs(processingData);
      } catch (error) {
        return { success: false, action: 'processing_error', message: 'Erro ao chamar função de processamento: ' + error.message };
      }
      if (result.success) {
        return { success: true, action: 'processed', processedCount: validPrompts.length, message: `Processamento concluído com sucesso! ${validPrompts.length} seção(ões) melhorada(s).` };
      } else {
        return { success: false, action: 'processing_error', message: `Erro no processamento: ${result.message || 'Erro desconhecido'}` };
      }
    } catch (error) {
      return { success: false, action: 'error', message: 'Erro durante o processamento brush com pares: ' + error.message };
    }
  }
};

  setupBrushButtonPairs: function() {
    try {
      const properties = PropertiesService.getScriptProperties();
      const existingPairs = properties.getProperty('brushButtonPairs');

      if (!existingPairs) {
        const defaultPairs = "Prontuário Médico, Prontuário Médico; Prontuário Médico, Prescrição de Óculos; Laudo de Mapeamento de Retina, Laudo de Mapeamento de Retina";

        properties.setProperty('brushButtonPairs', defaultPairs);
        console.log('✅ Pares padrão do botão brush configurados:', defaultPairs);

        return {
          success: true,
          message: 'Pares padrão configurados com sucesso.'
        };
      } else {
        console.log('📝 Pares do botão brush já existem:', existingPairs);
        return {
          success: true,
          message: 'Pares já configurados.'
        };
      }
    } catch (error) {
      console.error('❌ Erro ao configurar pares do botão brush:', error);
      return {
        success: false,
        message: 'Erro ao configurar pares: ' + error.message
      };
    }
  }
};

  callMagicModal: function(promptsData) {
    console.log('[SB_server_buttonMagic] Iniciando modal mágico com', promptsData ? promptsData.length : 0, 'prompts');

    try {
      const prompts = promptsData && promptsData.length > 0 ? promptsData : this.getAllPromptsForMagic();

      const uniqueHeadline1sWithEligibleH2s = this.getUniqueDocumentHeadline1sWithEligibleH2s(prompts);

      this.showMagicModal(prompts, uniqueHeadline1sWithEligibleH2s);
    } catch (error) {
      console.error('[SB_server_buttonMagic] Erro ao abrir modal mágico:', error);
      throw error;
    }
  }
};

  getUniqueDocumentHeadline1sWithEligibleH2s: function(promptsArray) {
    try {
      const documentManager = new DocumentService.DocumentHeadlineManager();
      const uniqueHeadline1s = documentManager.getUniqueHeadline1s();

      const fromH2Elegíveis = [...new Set(promptsArray.map(prompt => prompt.fromHeadline2))];

      const uniqueHeadline1sWithEligibleH2s = uniqueHeadline1s.map(headline1 => {
        const allH2sForThisH1 = documentManager.getHeadline2sForHeadline1(headline1);
        const eligibleH2s = allH2sForThisH1.filter(h2 => fromH2Elegíveis.includes(h2));

        return {
          headline1: headline1,
          eligibleH2s: eligibleH2s
        };
      }).filter(item => item.eligibleH2s.length > 0);

      console.log('[SB_server_buttonMagic] H1s com H2s elegíveis:', uniqueHeadline1sWithEligibleH2s.length);

      return uniqueHeadline1sWithEligibleH2s;
    } catch (error) {
      console.error('[SB_server_buttonMagic] Erro ao obter H1s com H2s elegíveis:', error);
      throw error;
    }
  }
};

  getAllPromptsForMagic: function() {
    try {
      const sheetId = ConfigService.getSheetId();
      const documentId = ConfigService.getDocumentId();

      if (!sheetId || !documentId) {
        throw new Error('IDs de planilha ou documento não encontrados');
      }

      const dataFetcher = new DataService.GoogleDriveDataFetcher(documentId, sheetId);
      const data = dataFetcher.extractAllDataFromSources();

      const promptsArray = [];

      if (data.allPrompts) {
        Object.keys(data.allPrompts).forEach(key => {
          const prompt = data.allPrompts[key];
          promptsArray.push({
            optionText: key,
            fromHeadline2: prompt.fromHeadline2,
            toHeadline2: prompt.toHeadline2,
            promptContent: prompt.content
          });
        });
      }

      console.log('[SB_server_buttonMagic] Prompts carregados:', promptsArray.length);
      return promptsArray;
    } catch (error) {
      console.error('[SB_server_buttonMagic] Erro ao obter prompts:', error);
      throw error;
    }
  }
};

  showMagicModal: function(prompts, uniqueHeadline1s) {
    try {
      console.log('[SB_server_buttonMagic] Preparando modal com', prompts.length, 'prompts e', uniqueHeadline1s.length, 'H1 únicos');

      const template = HtmlService.createTemplateFromFile('client/dialogs/ai_magic_dialog.html');

      template.promptData = JSON.stringify(prompts);
      template.uniqueHeadline1sData = JSON.stringify(uniqueHeadline1s);

      const html = template.evaluate()
        .setWidth(900)
        .setHeight(700);

      DocumentApp.getUi().showModalDialog(html, 'Processamento de IA');

      console.log('[SB_server_buttonMagic] Modal exibido com sucesso');
    } catch (error) {
      console.error('[SB_server_buttonMagic] Erro ao exibir modal:', error);
      throw error;
    }
  }
};