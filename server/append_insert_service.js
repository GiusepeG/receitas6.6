const AppendInsertService = {
  insertTextAndFormat: function(jsonObj) {
    const obj = JSON.parse(jsonObj);
    const textToInsert = obj.text;
    const shouldFormat = obj.shouldFormat;
    if (!textToInsert) {
      if (shouldFormat) DocumentService.executeFormat();
      return "FORMATTED_ONLY";
    }
    const selection = DocumentApp.getActiveDocument().getSelection();
    if (selection && selection.getRangeElements().length > 0) {
      const htmlTemplate = HtmlService.createTemplateFromFile('dialogs/append_choice_dialog.html');
      htmlTemplate.textToInsert = textToInsert;
      htmlTemplate.shouldFormat = shouldFormat;
      DocumentApp.getUi().showModelessDialog(htmlTemplate.evaluate().setWidth(300).setHeight(200), 'Posicione o Cursor');
      return "DIALOG_OPENED";
    } else {
      const success = this.insertTextAtCursor(jsonObj);
      return success ? "INSERTED_DIRECTLY" : "INSERTION_FAILED";
    }
  },

  appendTextAndFormat: function(jsonObj) {
    const obj = JSON.parse(jsonObj);
    const textToInsert = obj.text;
    const shouldFormat = obj.shouldFormat;
    if (!textToInsert) {
      if (shouldFormat) DocumentService.executeFormat();
      return "FORMATTED_ONLY";
    }
    const success = this.appendTextAtEnd(jsonObj);
    return success ? "INSERTED_DIRECTLY" : "INSERTION_FAILED";
  },

  appendTextAtEnd: function(jsonObj) {
    const obj = JSON.parse(jsonObj);
    const textToInsert = obj.text;
    const shouldFormat = obj.shouldFormat;
    const body = DocumentApp.getActiveDocument().getBody();
    try {
      body.appendParagraph(textToInsert);
    } catch (e) {
      DocumentApp.getUi().alert('Erro ao adicionar texto ao final do documento: ' + e.message);
      return false;
    }
    if (shouldFormat) DocumentService.executeFormat();
    return true;
  },

  insertTextAtCursor: function(jsonObj) {
    const obj = JSON.parse(jsonObj);
    const textToInsert = obj.text;
    const shouldFormat = obj.shouldFormat;
    const doc = DocumentApp.getActiveDocument();
    const selection = doc.getSelection();
    if (selection && selection.getRangeElements().length > 0) {
      throw new Error('Ainda há texto selecionado. Por favor, posicione o cursor sem selecionar nenhum texto e tente novamente.');
    }
    const cursor = doc.getCursor();
    if (cursor) {
      try {
        const element = cursor.insertText(textToInsert + '\n');
        if (!element) {
          doc.getBody().appendParagraph(textToInsert);
        }
      } catch (e) {
        doc.getBody().appendParagraph(textToInsert);
      }
    } else {
      doc.getBody().appendParagraph(textToInsert);
    }
    if (shouldFormat) DocumentService.executeFormat();
    return true;
  },

  insertTextAtCursorOnly: function(jsonObj) {
    const obj = JSON.parse(jsonObj);
    const textToInsert = obj.text;
    const shouldFormat = obj.shouldFormat;
    const doc = DocumentApp.getActiveDocument();
    const selection = doc.getSelection();
    if (selection && selection.getRangeElements().length > 0) {
      throw new Error('Ainda há texto selecionado. Por favor, posicione o cursor sem selecionar nenhum texto e tente novamente.');
    }
    const cursor = doc.getCursor();
    if (!cursor) {
      throw new Error('Cursor não disponível. Posicione o cursor no documento e tente novamente.');
    }
    try {
      const element = cursor.insertText(textToInsert + '\n');
      if (!element) {
        throw new Error('Não foi possível inserir o texto na posição do cursor. Tente posicionar o cursor em outro local.');
      }
    } catch (e) {
      throw new Error('Erro ao inserir texto na posição do cursor: ' + e.message);
    }
    if (shouldFormat) DocumentService.executeFormat();
    return true;
  },

  appendTextDirectlyToEnd: function(jsonObj) {
    const obj = JSON.parse(jsonObj);
    const textToInsert = obj.text;
    const shouldFormat = obj.shouldFormat;
    const body = DocumentApp.getActiveDocument().getBody();
    try {
      body.appendParagraph(textToInsert);
    } catch (e) {
      throw new Error('Erro ao adicionar texto ao final do documento: ' + e.message);
    }
    if (shouldFormat) DocumentService.executeFormat();
    return true;
  },

  executeAppendToSelectedH1: function(selectedH1, textToAppend, shouldFormat) {
    try {
      const doc = DocumentApp.getActiveDocument();
      const body = doc.getBody();
      const bodyText = body.getText();
      const formattingRules = DocumentService.rules.getFormattingRules();
      const placeholders = DocumentService.rules.getRulesForPlaceholders();
      const heading1Rules = formattingRules.filter(rule => rule.heading === DocumentApp.ParagraphHeading.HEADING1);
      const heading2Rules = formattingRules.filter(rule => rule.heading === DocumentApp.ParagraphHeading.HEADING2);
      const textManipulator = new DocumentService.TextManipulator(bodyText, placeholders, heading1Rules, heading2Rules);
      const processedText = textManipulator.splitInLines().removeEmptyLines().trimLeadingSpaces().checkAndModifyFirstLine().processLines().structureH2Blocks().getResult();
      const documentStructure = Utils.buildDocumentStructureWithH1Only(processedText, heading1Rules, heading2Rules);
      const uniqueH1s = [...new Set(documentStructure.map(item => item.headline1).filter(Boolean))];
      if (!uniqueH1s.includes(selectedH1)) {
        return { success: false, message: `O Headline1 "${selectedH1}" não foi encontrado no documento.` };
      }
      const modifiedText = this.insertTextIntoSpecificH1(processedText, selectedH1, textToAppend, heading1Rules, heading2Rules);
      body.setText(modifiedText);
      if (shouldFormat) DocumentService.executeFormat();
      return { success: true, message: `Texto adicionado com sucesso ao paciente "${selectedH1}".` };
    } catch (error) {
      return { success: false, message: 'Erro ao adicionar texto: ' + error.message };
    }
  },

  insertTextIntoSpecificH1: function(processedText, targetH1, textToInsert, heading1Rules, heading2Rules) {
    const lines = processedText.split('\n');
    const modifiedLines = [];
    let foundTargetH1 = false;
    let targetH1EndIndex = -1;
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i];
      if (heading1Rules.some(rule => rule.condition(line))) {
        if (foundTargetH1) {
          targetH1EndIndex = i - 1;
          break;
        }
        if (line === targetH1) {
          foundTargetH1 = true;
        }
      }
    }
    if (foundTargetH1 && targetH1EndIndex === -1) {
      targetH1EndIndex = lines.length - 1;
    }
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i];
      if (i === targetH1EndIndex) {
        modifiedLines.push(line);
        if (line.trim() !== '') {
          modifiedLines.push('');
        }
        modifiedLines.push(textToInsert);
        continue;
      }
      modifiedLines.push(line);
    }
    if (foundTargetH1 && targetH1EndIndex === -1) {
      modifiedLines.push('');
      modifiedLines.push(textToInsert);
    }
    return modifiedLines.join('\n');
  },

  checkSelectionAndInsertAtCursor: function(jsonObj) {
    const obj = JSON.parse(jsonObj);
    const textToInsert = obj.text;
    const shouldFormat = obj.shouldFormat;
    const doc = DocumentApp.getActiveDocument();
    const selection = doc.getSelection();
    if (selection && selection.getRangeElements().length > 0) {
      return { success: false, hasSelection: true, message: 'Há texto selecionado no documento. Posicione o cursor sem selecionar texto e tente novamente.' };
    }
    const cursor = doc.getCursor();
    if (!cursor) {
      return { success: false, hasSelection: false, message: 'Cursor não disponível. Posicione o cursor no documento e tente novamente.' };
    }
    try {
      const element = cursor.insertText(textToInsert + '\n');
      if (!element) {
        return { success: false, hasSelection: false, message: 'Não foi possível inserir o texto na posição do cursor. Tente posicionar o cursor em outro local.' };
      }
    } catch (e) {
      return { success: false, hasSelection: false, message: 'Erro ao inserir texto na posição do cursor: ' + e.message };
    }
    if (shouldFormat) {
      try {
        DocumentService.executeFormat();
      } catch (e) {
        // ignore
      }
    }
    return { success: true, hasSelection: false, message: 'Texto inserido com sucesso na posição do cursor.' };
  },

  validateAndExecuteAppend: function(jsonObj) {
    const obj = JSON.parse(jsonObj);
    const textToInsert = obj.text;
    const shouldFormat = obj.shouldFormat;
    try {
      const doc = DocumentApp.getActiveDocument();
      const body = doc.getBody();
      const bodyText = body.getText();
      const formattingRules = DocumentService.rules.getFormattingRules();
      const placeholders = DocumentService.rules.getRulesForPlaceholders();
      const heading1Rules = formattingRules.filter(rule => rule.heading === DocumentApp.ParagraphHeading.HEADING1);
      const heading2Rules = formattingRules.filter(rule => rule.heading === DocumentApp.ParagraphHeading.HEADING2);
      const textManipulator = new DocumentService.TextManipulator(bodyText, placeholders, heading1Rules, heading2Rules);
      const processedText = textManipulator.splitInLines().removeEmptyLines().trimLeadingSpaces().checkAndModifyFirstLine().processLines().structureH2Blocks().getResult();
      const documentStructure = Utils.buildDocumentStructureWithH1Only(processedText, heading1Rules, heading2Rules);
      const uniqueH1s = [...new Set(documentStructure.map(item => item.headline1).filter(Boolean))];
      if (uniqueH1s.length === 0) {
        try {
          body.appendParagraph(textToInsert);
          if (shouldFormat) DocumentService.executeFormat();
          return { success: true, action: 'append_direct', message: 'Texto adicionado ao final do documento (sem H1s).' };
        } catch (e) {
          return { success: false, action: 'append_direct', message: 'Erro ao adicionar texto ao final do documento: ' + e.message };
        }
      }
      if (uniqueH1s.length === 1) {
        try {
          const success = this.appendTextAtEnd(JSON.stringify(obj));
          return { success: success, action: 'append_normal', message: success ? 'Texto adicionado com sucesso.' : 'Falha ao adicionar texto.' };
        } catch (e) {
          return { success: false, action: 'append_normal', message: 'Erro ao adicionar texto: ' + e.message };
        }
      }
      return { success: false, action: 'show_dialog', uniqueHeadline1s: uniqueH1s, textToInsert: textToInsert, shouldFormat: shouldFormat, message: `Foram encontrados ${uniqueH1s.length} Headline1 diferentes no documento.` };
    } catch (error) {
      return { success: false, action: 'error', message: 'Erro ao analisar a estrutura do documento: ' + error.message };
    }
  }
};
