
/**
 * Obtém o ID do documento das propriedades
 * @return {string|null} ID do documento ou null
 */
function getDocumentId() {
    const SP = PropertiesService.getScriptProperties();
    const documentId = SP.getProperty('docData');
    
    if (!documentId) {
      console.error('❌ ID do documento não encontrado');
      return null;
    }
    
    return documentId;
  }
  
  /**
   * Obtém o ID da planilha das propriedades
   * @return {string|null} ID da planilha ou null
   */
  function getSheetId() {
    const SP = PropertiesService.getScriptProperties();
    const sheetId = SP.getProperty('sheetData');
    
    if (!sheetId) {
      console.error('❌ ID da planilha não encontrado');
      return null;
    }
    
    return sheetId;
  }
  
  /**
   * Valida se as propriedades necessárias estão configuradas
   * @return {boolean} true se todas configuradas
   */
  function validateScriptProperties() {
      const SP = PropertiesService.getScriptProperties();
      const sheetDataId = SP.getProperty('sheetData');
      const docDataId = SP.getProperty('docData');
      const logDocFileId = SP.getProperty('logDocFileId');
      
      if (!sheetDataId || !docDataId) {
        console.error('❌ Propriedades principais não configuradas (sheetData ou docData)');
        return false;
      }
      
      if (!logDocFileId) {
        console.warn('⚠️ Propriedade logDocFileId não configurada - logs de interação não funcionarão');
        // Não retorna false pois o log não é crítico para o funcionamento principal
      }
      
      return true;
  }
  
  /**
   * Configura propriedades com valores padrão
   */
  function setupScriptProperties() {
      const SP = PropertiesService.getScriptProperties();
      
      // IDs padrão - você pode alterar estes valores
      const defaultSheetData = "1uaStYZhSDK7tXtBsJlLJrCm59RQwBjKV7keLVFME2nM";
      const defaultDocData = "1wA4vP5qoUnNpvR7YxrV3nCGEwwKGjJ6gzgZfgLQJnrM";
      // ID padrão para documento de log de interações da IA - você pode alterar este valor
      const defaultLogDocFileId = "1wA4vP5qoUnNpvR7YxrV3nCGEwwKGjJ6gzgZfgLQJnrM"; // Mesmo que docData por padrão
      
      if (!SP.getProperty('sheetData')) {
        SP.setProperty('sheetData', defaultSheetData);
      }
      
      if (!SP.getProperty('docData')) {
        SP.setProperty('docData', defaultDocData);
      }
      
      // Configura o ID do documento de log se não existir
      if (!SP.getProperty('logDocFileId')) {
        SP.setProperty('logDocFileId', defaultLogDocFileId);
        console.log("✅ Propriedade logDocFileId configurada com valor padrão:", defaultLogDocFileId);
      }
      
      validateScriptProperties();
  }
  
  /**
   * Configura especificamente o ID do documento de log
   * @param {string} logDocFileId - ID do documento para log (opcional)
   * @return {boolean} true se configurado com sucesso
   */
  function setupLogDocumentId(logDocFileId = null) {
      const SP = PropertiesService.getScriptProperties();
      
      if (!logDocFileId) {
          // Se não fornecido, usa o mesmo ID do documento principal por padrão
          logDocFileId = SP.getProperty('docData');
          
          if (!logDocFileId) {
              console.error('❌ Não foi possível configurar logDocFileId: docData não encontrado');
              return false;
          }
      }
      
      try {
          // Testa se o documento existe e se temos acesso
          const testDoc = DocumentApp.openById(logDocFileId);
          testDoc.getName(); // Testa acesso
          
          SP.setProperty('logDocFileId', logDocFileId);
          console.log('✅ logDocFileId configurado com sucesso:', logDocFileId);
          return true;
          
      } catch (error) {
          console.error('❌ Erro ao configurar logDocFileId:', error.message);
          console.error('Verifique se o ID do documento está correto e se você tem acesso a ele');
          return false;
      }
  }
  
  /**
   * Cria um novo documento especificamente para logs de interação da IA
   * @return {string|null} ID do documento criado ou null se falhou
   */
  function createLogDocument() {
      try {
          const newDoc = DocumentApp.create('Log de Interações IA - ' + new Date().toLocaleDateString());
          const newDocId = newDoc.getId();
          
          // Adiciona conteúdo inicial ao documento
          const body = newDoc.getBody();
          body.appendParagraph('Log de Interações da IA').setHeading(DocumentApp.ParagraphHeading.TITLE);
          body.appendParagraph('Este documento registra todas as interações entre os prompts e a IA.')
              .setHeading(DocumentApp.ParagraphHeading.SUBTITLE);
          body.appendParagraph('Criado automaticamente em: ' + new Date().toLocaleString());
          body.appendParagraph(''); // Linha em branco
          
          console.log('✅ Novo documento de log criado:', newDocId);
          
          // Configura automaticamente como logDocFileId
          if (setupLogDocumentId(newDocId)) {
              return newDocId;
          } else {
              return null;
          }
          
      } catch (error) {
          console.error('❌ Erro ao criar documento de log:', error.message);
          return null;
      }
  }
  
  /**
   * Obtém prompts filtrados por fromHeadline2 (função de servidor para fallback)
   * @param {string} fromHeadline2 - Valor do fromHeadline2 para filtrar
   * @return {Array} Array com prompts filtrados
   */
  function getFilteredPromptsFromServer(fromHeadline2) {
    try {
      // Obtém os IDs necessários das propriedades (mesmo padrão de getSidebarData)
      const sheetId = getSheetId();
      const documentId = getDocumentId();
  
      if (!sheetId) {
        const errorMsg = '🛑 FALHA CRÍTICA: ID da planilha (sheetData) não encontrado nas propriedades do script.';
        console.error(errorMsg);
        return [];
      }
  
      if (!documentId) {
        const errorMsg = '🛑 FALHA CRÍTICA: ID do Documento (docData) não encontrado nas propriedades do script.';
        console.error(errorMsg);
        return [];
      }
  
      // Cria DataFetcher com os IDs corretos (mesmo padrão de getSidebarData)
      const dataFetcher = new GoogleDriveDataFetcher(documentId, sheetId);
      const sidebarData = dataFetcher.extractAllDataFromSources();
      const allPrompts = sidebarData.allPrompts || {};
      
      const filteredPrompts = [];
      
      Object.keys(allPrompts).forEach(promptKey => {
        const prompt = allPrompts[promptKey];
        if (prompt.fromHeadline2 === fromHeadline2) {
          filteredPrompts.push({
            key: promptKey,
            title: promptKey,
            fromHeadline2: prompt.fromHeadline2,
            toHeadline2: prompt.toHeadline2,
            content: prompt.content
          });
        }
      });
      
      console.log('✅ Prompts filtrados obtidos do servidor:', filteredPrompts.length);
      return filteredPrompts;
      
    } catch (error) {
      console.error('❌ Erro ao obter prompts filtrados do servidor:', error);
      return [];
    }
  }

function includeServerSide(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function validateSingleHeadline1() {
  try {
    const doc = DocumentApp.getActiveDocument();
    const bodyText = doc.getBody().getText();

    const h1Pattern = /^#\s+/gm;
    const h1Matches = bodyText.match(h1Pattern);

    if (!h1Matches || h1Matches.length === 0) {
      return {
        isValid: false,
        message: 'Nenhum Headline1 encontrado no documento.'
      };
    }

    if (h1Matches.length === 1) {
      return {
        isValid: true,
        message: 'Documento válido com um único Headline1.'
      };
    }

    return {
      isValid: false,
      showDialog: true,
      uniqueHeadline1s: h1Matches.map((match, index) => `H1 ${index + 1}`),
      message: `Foram encontrados ${h1Matches.length} Headline1 diferentes no documento.`
    };

  } catch (error) {
    return {
      isValid: false,
      message: 'Erro ao analisar a estrutura do documento: ' + error.message
    };
  }
}

function buildDocumentStructureWithH1Only(text, headline1Rules, headline2Rules) {
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

function openAppendChoiceDialog(uniqueHeadline1s, textToAppend, shouldFormat) {
  const html = HtmlService.createTemplateFromFile('dialogs/append_choice_dialog.html');
  html.uniqueHeadline1s = uniqueHeadline1s;
  html.textToAppend = textToAppend;
  html.shouldFormat = shouldFormat;

  const dialog = html.evaluate()
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setWidth(400)
    .setHeight(300);

  DocumentApp.getUi().showModalDialog(dialog, 'Escolha o Headline1');
}

function validateHeadline1sForBrush() {
  try {
    const doc = DocumentApp.getActiveDocument();
    const bodyText = doc.getBody().getText();

    const validation = validateSingleHeadline1();

    if (!validation.isValid && validation.showDialog) {
      return {
        isValid: false,
        showDialog: true,
        uniqueHeadline1s: validation.uniqueHeadline1s,
        message: validation.message
      };
    }

    return validation;

  } catch (error) {
    return {
      isValid: false,
      message: 'Erro ao validar estrutura para brush: ' + error.message
    };
  }
}

function getBrushToHeadline2s() {
  try {
    const properties = PropertiesService.getScriptProperties();
    const stored = properties.getProperty('brushToHeadline2s');

    if (stored) {
      return JSON.parse(stored);
    }

    return [];

  } catch (error) {
    console.error('❌ Erro ao obter brush mapping:', error);
    return [];
  }
}

function saveBrushToHeadline2s(toHeadline2sArray) {
  try {
    const properties = PropertiesService.getScriptProperties();
    properties.setProperty('brushToHeadline2s', JSON.stringify(toHeadline2sArray));

    console.log('✅ Brush mapping salvo');

  } catch (error) {
    console.error('❌ Erro ao salvar brush mapping:', error);
  }
}
  