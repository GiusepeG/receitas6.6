  DocumentHeadlineManager: class {
    constructor() {
      // Obtém automaticamente os dados necessários do documento ativo
      const doc = DocumentApp.getActiveDocument();
      const bodyText = doc.getBody().getText();

      // Obtém o texto selecionado (pode ser vazio se nada estiver selecionado)
      let selectionText = '';
      try {
        const selection = doc.getSelection();
        if (selection) {
          selectionText = selection.getText();
        }
      } catch (e) {
        // Se não houver seleção, selectionText permanece vazio
      }

      // Obtém as regras de formatação
      const headline1Conditions = DocumentService.rules.getFormattingRules().filter(rule => rule.heading === DocumentApp.ParagraphHeading.HEADING1);
      const headline2Conditions = DocumentService.rules.getFormattingRules().filter(rule => rule.heading === DocumentApp.ParagraphHeading.HEADING2);

      this.bodyText = bodyText;
      this.selectionText = selectionText;

      this.bodyStructure = this.buildDocumentStructure(this.bodyText, headline1Conditions, headline2Conditions);
      this.selectionStructure = selectionText
        ? this.buildDocumentStructure(this.selectionText, headline1Conditions, headline2Conditions, true)
        : null;

    }

    buildDocumentStructure(text, headline1Conditions, headline2Conditions, limitToOneHeadline2 = false) {
      const lines = text.split('\n');
      let currentHeadline1 = null;
      let documentStructure = [];
      let foundFirstStructure = false;
      let awaitingHeadline2 = false;

      lines.forEach((line, idx) => {

        if (headline1Conditions.some(rule => rule.condition(line))) {
          currentHeadline1 = line;
          awaitingHeadline2 = true;
        }
        else if (headline2Conditions.some(rule => rule.condition(line))) {
          if (limitToOneHeadline2 && foundFirstStructure) return;
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

      return documentStructure;
    }

    GroupedHeadlines() {
      const groupedStructure = [];
      const headline1Map = new Map();

      this.bodyStructure.forEach(item => {
        if (!headline1Map.has(item.headline1)) {
          headline1Map.set(item.headline1, []);
        }
        headline1Map.get(item.headline1).push(item);
      });

      headline1Map.forEach((items, headline1) => {
        // Ordena os itens para colocar "Prontuário Médico" no topo de cada grupo
        const sortedItems = items.sort((a, b) => {
          if (a.headline2 === "Prontuário Médico" && b.headline2 !== "Prontuário Médico") {
            return -1; // a vem antes
          }
          if (a.headline2 !== "Prontuário Médico" && b.headline2 === "Prontuário Médico") {
            return 1; // b vem antes
          }
          return 0; // mantém ordem original para outros casos
        });

        groupedStructure.push(...sortedItems);
      });

      this.bodyStructure = groupedStructure;

      return this;
    }

    hasHeadline1AndHeadline2(structure, headline2) {
      return structure && structure.some(doc => doc.headline2 === headline2);
    }

    hasHeadline1AndHeadline2InSelection(headline2) {
      return this.hasHeadline1AndHeadline2(this.selectionStructure, headline2);
    }

    hasHeadline1AndHeadline2InBody(headline2) {
      return this.hasHeadline1AndHeadline2(this.bodyStructure, headline2);
    }

    getHeadline1fromSelection() {
      if (this.selectionStructure && this.selectionStructure.length > 0) {
        return this.selectionStructure[0].headline1;
      } else {
        throw new Error("Nenhum headline1 encontrado na seleção.");
      }
    }

    getHeadline1fromBody(headline2) {
      if (this.bodyStructure && this.bodyStructure.length > 0) {
        for (const item of this.bodyStructure) {
          if (item.headline2 === headline2) {
            return item.headline1;
          }
        }
        throw new Error(`Nenhum headline1 encontrado para o headline2: ${headline2}`);
      } else {
        throw new Error("Nenhuma estrutura encontrada no corpo do documento.");
      }
    }

    getFirstHeadline2fromBody(headline1) {
      if (this.bodyStructure && this.bodyStructure.length > 0) {
        for (const item of this.bodyStructure) {
          if (item.headline1 === headline1) {
            return item.headline2;
          }
        }
        throw new Error(`Nenhum headline2 encontrado para o headline1: ${headline1}`);
      } else {
        throw new Error("Nenhum headline1 encontrado no corpo do documento.");
      }
    }

    getFirstTextfromBody(headline2) {
      if (this.bodyStructure && this.bodyStructure.length > 0) {
        for (const item of this.bodyStructure) {
          if (item.headline2 === headline2) {
            return item.text;
          }
        }
        throw new Error(`Nenhum texto encontrado para o headline2: ${headline2}`);
      } else {
        throw new Error("Nenhuma estrutura encontrada no corpo do documento.");
      }
    }

    /**
     * Obtém apenas os Headline1 únicos do documento
     * @return {Array<string>} Array com os Headline1 únicos
     */
    getUniqueHeadline1s() {
      if (!this.bodyStructure || this.bodyStructure.length === 0) {
        return [];
      }

      // Extrai apenas os Headline1 únicos e não vazios
      const uniqueHeadlines = [...new Set(
        this.bodyStructure
          .map(item => item.headline1)
          .filter(Boolean) // Remove valores null/undefined/vazios
      )];

      return uniqueHeadlines;
    }

    /**
     * Obtém todos os Headline2 associados a um Headline1 específico
     * @param {string} headline1 - O Headline1 para buscar
     * @return {Array<string>} Array com os Headline2 associados
     */
    getHeadline2sForHeadline1(headline1) {
      if (!this.bodyStructure || this.bodyStructure.length === 0) {
        return [];
      }

      const headline2s = this.bodyStructure
        .filter(item => item.headline1 === headline1)
        .map(item => item.headline2)
        .filter(Boolean);

      return [...new Set(headline2s)]; // Remove duplicatas
    }

    /**
     * Reúne todo o conteúdo de todos os H2 sob um H1 específico.
     * @param {string} headline1 - O Headline1 (paciente) para o qual o conteúdo será reunido.
     * @return {string|null} Uma string única com todo o conteúdo ou null se o H1 não for encontrado.
     */
    getAugmentedContentForH1(headline1) {
      if (!this.bodyStructure || this.bodyStructure.length === 0) {
        console.warn(`[getAugmentedContentForH1] A estrutura do corpo está vazia. Não é possível buscar por "${headline1}".`);
        return null;
      }

      // Filtra todos os itens que pertencem ao headline1 especificado
      const patientSections = this.bodyStructure.filter(item => item.headline1 === headline1);

      if (patientSections.length === 0) {
        console.warn(`[getAugmentedContentForH1] Nenhum conteúdo encontrado para o paciente: "${headline1}".`);
        return null; // Retorna nulo se nenhum H2 for encontrado para este H1
      }
      // Concatena o texto de todas as seções, cada uma com seu H2
      const combinedContent = patientSections
        .map(section => {
          // Garante que o texto não seja nulo ou indefinido antes de juntar
          const textContent = section.text || '';
          // Remove a quebra de linha final para evitar espaçamento duplo
          const cleanedContent = textContent.endsWith('\n') ? textContent.slice(0, -1) : textContent;
          return `${section.headline2}\n${cleanedContent}`;
        })
        .join('\n\n'); // Separa as seções com uma linha em branco


      return combinedContent;
    }

         /**
      * Cria ou atualiza uma estrutura H1/H2 no documento
      * @param {string} newText - Novo texto a ser inserido
      * @param {string} headline1 - Headline1 (paciente)
      * @param {string} newHeadline2 - Headline2 de destino
      * @param {string} fromHeadline2 - Headline2 de origem (opcional)
      *
      * REGRAS DE COMPORTAMENTO:
      * - Se fromHeadline2 === newHeadline2: ATUALIZAÇÃO - Remove todas as estruturas antigas {H1, H2} e substitui pela nova
      * - Se fromHeadline2 = "Prontuário Médico" e newHeadline2 = "Prescrição de Óculos": ATUALIZAÇÃO - Mantém "Prontuário Médico" e atualiza "Prescrição de Óculos"
      * - Se fromHeadline2 !== newHeadline2: ADIÇÃO - Adiciona nova estrutura, mantendo todas as existentes
      * - Se fromHeadline2 for null: ADIÇÃO - Comportamento padrão de adicionar nova estrutura
      * - Evita duplicatas exatas (mesmo H1, H2 e texto)
      */
    createOrUpdateBodyHeadline2(newText, headline1, newHeadline2, fromHeadline2 = null) {
      let cleanedText = newText;
      // Remove o headline2 do início do texto, permitindo espaços em branco antes
      if (cleanedText.trim().startsWith(newHeadline2)) {
        cleanedText = cleanedText.substring(cleanedText.indexOf(newHeadline2) + newHeadline2.length);
        // Se o restante for apenas espaços/quebras de linha, remove-os
        if (cleanedText.trim().length === 0) {
          cleanedText = '';
        } else {
          // Remove apenas a quebra de linha inicial, se houver
          cleanedText = cleanedText.startsWith('\n') ? cleanedText.substring(1) : cleanedText;
        }
      }

      // ✅ REGRA ESPECÍFICA BRUSH&MAGIC: Se fromHeadline2 === toHeadline2, é uma ATUALIZAÇÃO
      if (fromHeadline2 && fromHeadline2 === newHeadline2) {
        // Remove todas as estruturas antigas com o mesmo H1 e H2
        this.bodyStructure = this.bodyStructure.filter(doc =>
          !(doc.headline1 === headline1 && doc.headline2 === fromHeadline2)
        );

        // Adiciona a nova estrutura atualizada
        this.bodyStructure.push({ headline1, headline2: newHeadline2, text: cleanedText });
        return;
      }

      // ✅ REGRA ESPECÍFICA: Prontuário Médico → Prescrição de Óculos (mantém prontuário, atualiza prescrição)
      if (fromHeadline2 === "Prontuário Médico" && newHeadline2 === "Prescrição de Óculos") {
        // Remove todas as estruturas antigas de "Prescrição de Óculos" para este H1
        this.bodyStructure = this.bodyStructure.filter(doc =>
          !(doc.headline1 === headline1 && doc.headline2 === "Prescrição de Óculos")
        );

        // Adiciona a nova estrutura de "Prescrição de Óculos" (mantém o Prontuário Médico)
        this.bodyStructure.push({ headline1, headline2: newHeadline2, text: cleanedText });
        return;
      }

      // ✅ Verifica se já existe um item EXATAMENTE igual (mesmo H1, H2 e texto)
      const exactMatch = this.bodyStructure.find(doc =>
        doc.headline1 === headline1 &&
        doc.headline2 === newHeadline2 &&
        doc.text === cleanedText
      );

      if (exactMatch) {
        // Se o item é exatamente igual, não faz nada para evitar duplicatas totais
        return;
      }

      // ✅ Permite múltiplos H2s iguais para o mesmo H1 - sempre adiciona nova entrada
      this.bodyStructure.push({ headline1, headline2: newHeadline2, text: cleanedText });
    }

    /**
     * Obtém os pares de Headline1 e Headline2 do documento
     * @return {Array} Array de objetos com headline1 e headline2
     */
    getDocumentHeadlinePairs() {
      if (!this.bodyStructure || this.bodyStructure.length === 0) {
        return [];
      }

      return this.bodyStructure.map(item => ({
        headline1: item.headline1,
        headline2: item.headline2
      }));
    }

    /**
     * Obtém os pares de Headline1 e Headline2 do documento junto com seu conteúdo
     * @return {Array} Array de objetos com headline1, headline2 e content
     */
    getDocumentHeadlinePairsAndContent() {
      if (!this.bodyStructure || this.bodyStructure.length === 0) {
        return [];
      }

      return this.bodyStructure.map(item => ({
        headline1: item.headline1,
        headline2: item.headline2,
        content: item.text ? item.text.trim() : ''
      }));
    }

    /**
     * Obtém o conteúdo de texto para um par H1/H2 específico.
     * @param {string} headline1 - O Headline1 a ser buscado.
     * @param {string} headline2 - O Headline2 a ser buscado.
     * @return {string|null} O conteúdo de texto, ou null se não encontrado.
     */
    getContentForPair(headline1, headline2) {
      const item = this.bodyStructure.find(doc => doc.headline1 === headline1 && doc.headline2 === headline2);
      if (item && typeof item.text !== 'undefined' && item.text !== null) {
        return item.text;
      }
      // Retorna string vazia se o par existe mas não tem texto,
      // e null se o par não foi encontrado de todo.
      return item ? '' : null;
    }

    /**
     * Agrupa os Headline2 conforme seu Headline1 e retorna o texto organizado
     * @return {string} Texto do documento com Headline2 agrupados por Headline1
     */
    getGroupedText() {
      if (!this.bodyStructure || this.bodyStructure.length === 0) {
        return '';
      }

      // Agrupa os itens por Headline1
      const groupedByHeadline1 = new Map();

      this.bodyStructure.forEach(item => {
        if (!groupedByHeadline1.has(item.headline1)) {
          groupedByHeadline1.set(item.headline1, []);
        }
        groupedByHeadline1.get(item.headline1).push(item);
      });

      // Constrói o texto agrupado
      let groupedText = '';

      groupedByHeadline1.forEach((items, headline1) => {
        // Adiciona o Headline1 uma vez
        groupedText += `${headline1}\n`;

        // Adiciona todos os Headline2 e seus textos associados
        items.forEach(item => {
          groupedText += `${item.headline2}\n`;
          if (item.text) { // Garante que a propriedade text exista
            // Remove o último \n para evitar linhas duplas, mas mantém o resto
            const textContent = item.text.endsWith('\n') ? item.text.slice(0, -1) : item.text;
            groupedText += `${textContent}\n`;
          }
        });

        // Adiciona uma linha em branco entre grupos
        groupedText += '\n';
      });

      return groupedText.trim();
    }

    getText() {
      return this.bodyStructure
        .map(item => {
          const textContent = item.text || '';
          const cleanedContent = textContent.endsWith('\n') ? textContent.slice(0, -1) : textContent;
          return `${item.headline1}\n${item.headline2}\n${cleanedContent}`;
        })
        .join("\n\n");
    }

    logDocumentStructure(title = "Document Structure") {
      Logger.log(`=== ${title} ===`);
      if (!this.bodyStructure || this.bodyStructure.length === 0) {
        Logger.log("Estrutura do documento vazia.");
        return;
      }
      this.bodyStructure.forEach((item, idx) => {
        Logger.log(`[${idx}] Headline1: "${item.headline1}" | Headline2: "${item.headline2}" | Texto: "${(item.text || '')}"`);
      });
      Logger.log(`=== Fim de ${title} ===`);
    }


    // ✅ Novo método embutido para processar e atualizar o documento
    processAndUpdateDocument(newText, headline1, toHeadline2, doc, fromHeadline2 = null) {

      this.createOrUpdateBodyHeadline2(newText, headline1, toHeadline2, fromHeadline2);

      this.GroupedHeadlines();
      const newBodyText = this.getText();
      doc.getBody().setText(newBodyText);
    }
  }

  ParagraphManipulator: class {
    /**
     * @param {Array<Object>} processedLines Array of processed lines with text and heading properties.
     * @param {Array<Object>} formattingRules Array of formatting rules for paragraphs.
     * @param {Array<Object>} insideParagraphRules Array of formatting rules for text inside paragraphs.
     */
    constructor(processedLines, formattingRules, insideParagraphRules) {
      this.processedLines = processedLines || [];
      this.formattingRules = formattingRules || [];
      this.insideParagraphRules = insideParagraphRules || [];
      this.shouldLog = false;
      this.document = DocumentApp.getActiveDocument();
      this.body = this.document.getBody();
    }

    /**
     * Enables or disables logging for debugging purposes.
     * @param {boolean} [shouldLog=true]
     * @returns {ParagraphManipulator}
     */
    withLogging(shouldLog = true) {
      this.shouldLog = shouldLog;
      return this;
    }

    /**
     * Clears the document body and prepares it for new content.
     * @returns {ParagraphManipulator} The instance for chaining.
     */
    clearDocument() {
      this._log('Clearing document body...');
      this.body.clear();
      this._log(`Document cleared. Number of children: ${this.body.getNumChildren()}`);
      return this;
    }

    /**
     * Formats all processed lines into the document with proper styling and page breaks.
     * @returns {ParagraphManipulator} The instance for chaining.
     */
    formatParagraphs() {
      if (this.processedLines.length === 0) {
        this._log('No processed lines available for formatting.');
        return this;
      }

      this._log('Starting paragraph formatting...');

      const processedH1Texts = new Set();
      const h1Rule = this._getH1Rule();
      const defaultRule = this._getDefaultRule();

      let firstMeaningfulParagraphAdded = false;

      this.processedLines.forEach((item, index) => {
        const line = item.text;
        const heading = item.heading;

        this._log(`Processing line #${index}: "${line}" with heading: ${heading}`);

        // Check if line is completely empty (no characters at all)
        const isLineCompletelyEmpty = line === '';

        if (!isLineCompletelyEmpty) {
          this._log(`Line #${index}: Adding paragraph with text: "${line}"`);
          const paragraph = this.body.appendParagraph(line);
          const paragraphNativeIndex = this.body.getChildIndex(paragraph);

          firstMeaningfulParagraphAdded = true;

          // Apply formatting based on heading or rules
          this._applyFormatting(paragraph, line, heading, h1Rule, defaultRule, processedH1Texts);

          // Handle page breaks
          this._handlePageBreak(paragraph, paragraphNativeIndex, line, h1Rule);

        } else {
          this._log(`Line #${index}: Completely empty line detected`);
          if (firstMeaningfulParagraphAdded) {
            const emptyPara = this.body.appendParagraph(line);
            this._log(`Line #${index}: Empty paragraph added at index: ${this.body.getChildIndex(emptyPara)}`);
          } else {
            this._log(`Line #${index}: Initial empty line ignored`);
          }
        }
      });

      this._log('Paragraph formatting completed.');
      return this;
    }

    /**
     * Performs final cleanup of the document.
     * @returns {ParagraphManipulator} The instance for chaining.
     */
    cleanupDocument() {
      this._log('Starting document cleanup...');

      const paragraphs = this.body.getParagraphs();
      this._log(`Total paragraphs before cleanup: ${paragraphs.length}`);

      if (paragraphs.length > 0) {
        const firstPara = paragraphs[0];
        const firstParaText = firstPara.getText();

        this._log(`First paragraph text: "${firstParaText}"`);

        if (firstParaText.trim() === '') {
          this._log('First paragraph is empty, removing...');
          try {
            firstPara.removeFromParent();
            this._log('First empty paragraph removed successfully.');
          } catch (e) {
            this._log(`Error removing first paragraph: ${e.message}`);
          }
        } else {
          this._log('First paragraph is not empty, no cleanup needed.');
        }
      }

      this._log('Document cleanup completed.');
      return this;
    }

    /**
     * Internal log helper.
     * @private
     */
    _log(message) {
      if (this.shouldLog) {
        if (typeof Logger !== 'undefined') {
          Logger.log(message);
        } else {
          console.log(message);
        }
      }
    }

    /**
     * Gets the H1 rule from formatting rules.
     * @private
     */
    _getH1Rule() {
      return this.formattingRules.find(rule =>
        rule && rule.heading === DocumentApp.ParagraphHeading.HEADING1 && rule.matchType !== 'default'
      );
    }

    /**
     * Gets the default rule from formatting rules.
     * @private
     */
    _getDefaultRule() {
      return this.formattingRules.find(rule => rule.matchType === 'default');
    }

    /**
     * Applies formatting to a paragraph based on its heading and rules.
     * @private
     */
    _applyFormatting(paragraph, line, heading, h1Rule, defaultRule, processedH1Texts) {
      if (heading === DocumentApp.ParagraphHeading.HEADING1) {
        if (processedH1Texts.has(line)) {
          this._log(`H1 repeated: "${line}". Applying NORMAL heading with H1 style.`);
          paragraph.setHeading(DocumentApp.ParagraphHeading.NORMAL);
          if (h1Rule && h1Rule.style) {
            paragraph.setAttributes(h1Rule.style);
          }
        } else {
          this._log(`New H1: "${line}". Applying HEADING1.`);
          paragraph.setHeading(DocumentApp.ParagraphHeading.HEADING1);
          if (h1Rule && h1Rule.style) {
            paragraph.setAttributes(h1Rule.style);
          }
          processedH1Texts.add(line);
        }
      } else if (heading === DocumentApp.ParagraphHeading.HEADING2) {
        this._log(`H2: "${line}". Applying HEADING2.`);
        paragraph.setHeading(DocumentApp.ParagraphHeading.HEADING2);
        const h2Rule = this.formattingRules.find(rule =>
          rule && rule.heading === DocumentApp.ParagraphHeading.HEADING2
        );
        if (h2Rule && h2Rule.style) {
          paragraph.setAttributes(h2Rule.style);
        }
      } else {
        // Apply formatting rules for regular text
        this._applyFormattingRules(paragraph, line, defaultRule);
      }
    }

    /**
     * Applies formatting rules to a paragraph.
     * @private
     */
    _applyFormattingRules(paragraph, line, defaultRule) {
      let ruleApplied = null;

      for (const rule of this.formattingRules) {
        if (!rule || rule.matchType === 'default') continue;

        if (rule.condition && typeof rule.condition === 'function' && rule.condition(line)) {
          this._log(`Rule applied: "${line}"`);
          if (rule.heading) paragraph.setHeading(rule.heading);
          if (rule.style) paragraph.setAttributes(rule.style);
          ruleApplied = rule;
          break;
        }
      }

      if (!ruleApplied && defaultRule) {
        this._log(`Default rule applied: "${line}"`);
        if (defaultRule.heading) paragraph.setHeading(defaultRule.heading);
        if (defaultRule.style) paragraph.setAttributes(defaultRule.style);
      }
    }

    /**
     * Handles page breaks for paragraphs.
     * @private
     */
    _handlePageBreak(paragraph, paragraphNativeIndex, line, h1Rule) {
      // Check if page break is required
      let requiresPageBreak = false;

      if (h1Rule && h1Rule.condition && h1Rule.condition(line)) {
        requiresPageBreak = h1Rule.requiresPageBreak || false;
      }

      if (requiresPageBreak) {
        this._log(`Checking page break for paragraph at index ${paragraphNativeIndex}`);

        let suppressPageBreak = false;

        if (paragraphNativeIndex === 0) {
          suppressPageBreak = true;
          this._log('Page break suppressed (first paragraph)');
        } else if (paragraphNativeIndex === 1) {
          const firstChild = this.body.getChild(0);
          if (firstChild.getType() === DocumentApp.ElementType.PARAGRAPH &&
              firstChild.asParagraph().getText().trim() === '') {
            suppressPageBreak = true;
            this._log('Page break suppressed (second paragraph with ghost first)');
          }
        }

        if (!suppressPageBreak) {
          this._log(`Inserting page break before paragraph at index ${paragraphNativeIndex}`);
          this.body.insertPageBreak(paragraphNativeIndex);
        } else {
          this._log('Page break suppressed');
        }
      }
    }

    /**
     * Formats specific texts inside paragraphs based on formatting rules.
     * @returns {ParagraphManipulator} The instance for chaining.
     */
    formatTextInsideParagraphs() {
      if (!this.insideParagraphRules || this.insideParagraphRules.length === 0) {
        this._log('No inside paragraph formatting rules available.');
        return this;
      }

      this._log('Starting formatting inside paragraphs...');

      this.insideParagraphRules.forEach((rule, ruleIndex) => {
        this._log(`Processing rule ${ruleIndex + 1}: ${rule.styleName}`);

        if (!rule.targets || !Array.isArray(rule.targets)) {
          this._log(`Warning: Rule ${ruleIndex + 1} has no valid targets.`);
          return;
        }

        rule.targets.forEach((target, targetIndex) => {
          this._log(`Processing target ${targetIndex + 1}: "${target}"`);

          let searchResult = this.body.findText(target);
          let matchCount = 0;

          while (searchResult !== null) {
            const element = searchResult.getElement();

            // Check if the element is of text type
            if (element.getType() === DocumentApp.ElementType.TEXT) {
              const textElement = element.asText();
              const startOffset = searchResult.getStartOffset();
              const endOffset = searchResult.getEndOffsetInclusive();

              this._log(`Found match ${matchCount + 1} for "${target}" at positions ${startOffset}-${endOffset}`);

              // Apply the formatting based on the styleName
              switch (rule.styleName) {
                case 'bold':
                  textElement.setBold(startOffset, endOffset, true);
                  this._log(`Applied bold formatting to "${target}"`);
                  break;
                case 'italic':
                  textElement.setItalic(startOffset, endOffset, true);
                  this._log(`Applied italic formatting to "${target}"`);
                  break;
                case 'underline':
                  textElement.setUnderline(startOffset, endOffset, true);
                  this._log(`Applied underline formatting to "${target}"`);
                  break;
                default:
                  this._log(`Warning: Unknown style name "${rule.styleName}" for target "${target}"`);
                  break;
              }

              matchCount++;
            } else {
              this._log(`Element is not of text type: ${element.getType()}`);
            }

            // Find the next match
            searchResult = this.body.findText(target, searchResult);
          }

          this._log(`Total matches found for "${target}": ${matchCount}`);
        });
      });

      this._log('Inside paragraph formatting completed.');
      return this;
    }
  },

  TextManipulator: class {
    /**
     * @param {string} text The initial text to be manipulated.
     * @param {Array<Object>} placeholders
     * @param {Array<Object>} heading1Rules Rules for HEADING1 paragraphs.
     * @param {Array<Object>} heading2Rules Rules for HEADING2 paragraphs.
     */
    constructor(text, placeholders, heading1Rules, heading2Rules) {
      this.text = text;
      this.placeHolders = placeholders || [];
      this.heading1Rules = heading1Rules || [];
      this.heading2Rules = heading2Rules || [];

      this.lines = null;
      this.lastHeading1 = null;
      this.currentPatientName = null;
      this.processedLines = []; // Array of dictionaries with {text, heading}
      this.shouldLog = false;
    }

    /**
     * Enables or disables logging for debugging purposes.
     * @param {boolean} [shouldLog=true]
     * @returns {TextManipulator}
     */
    withLogging(shouldLog = true) {
      this.shouldLog = shouldLog;
      return this;
    }

    /**
     * Splits the initial text into an array of lines.
     * @returns {TextManipulator} The instance for chaining.
     */
    splitInLines() {
      this.lines = this.text.split(/\r\n|\r|\n/g);
      this._log('Text split into ' + this.lines.length + ' lines.');
      return this;
    }

    /**
     * Removes empty lines from the current array of lines.
     * @returns {TextManipulator} The instance for chaining.
     */
    removeEmptyLines() {
      if (this.lines === null) {
        this.splitInLines();
      }
      const originalCount = this.lines.length;
      this.lines = this.lines.filter(line => line !== "");
      this._log(`${originalCount - this.lines.length} empty lines removed.`);
      return this;
    }

    /**
     * Removes unnecessary leading spaces from lines while preserving intentional indentation.
     * @returns {TextManipulator} The instance for chaining.
     */
    trimLeadingSpaces() {
      if (this.lines === null) {
        this.splitInLines();
      }

      const originalCount = this.lines.length;
      this.lines = this.lines.map(line => {
        // Check if line starts with spaces followed by non-space characters
        const regexLeadingSpaces = /^\s+\S+/;
        if (regexLeadingSpaces.test(line)) {
          return line.trim();
        }
        return line;
      });

      this._log('Leading spaces trimmed from lines.');
      return this;
    }

    /**
     * Removes escape characters from all lines.
     * Replaces "\" with "" and "*" with " ".
     * @returns {TextManipulator} The instance for chaining.
     */
    removeEscapeCharacters() {
      if (this.lines === null) {
        this.splitInLines();
      }

      this.lines = this.lines.map(line => {
        return line.replace(/\\/g, '').replace(/\*/g, ' ');
      });

      this._log('Escape characters removed from lines.');
      return this;
    }

    /**
     * Checks the first line and modifies it if it appears to be a person's name.
     * Prefixes with "Nome: " if it matches the criteria.
     * @returns {TextManipulator} The instance for chaining.
     */
    checkAndModifyFirstLine() {
      if (this.lines === null) {
        this.splitInLines();
      }
      if (this.lines.length === 0) return this;

      const firstLine = this.lines[0];
      if (typeof firstLine !== 'string') {
        this.lines[0] = "";
        return this;
      }

      const trimmedText = firstLine.trim();
      if (trimmedText === "") {
          this.lines[0] = "";
          return this;
      }

      const regexUppercaseWithAccents = /^[A-ZÇÃÕÁÉÍÓÚÂÊÎÔÛÀÈÌÒÙÄËÏÖÜŸŠŽČĐÑ ]+$/;
      const words = trimmedText.split(" ").filter(word => word.length > 0);

      if (regexUppercaseWithAccents.test(trimmedText) && words.length >= 2) {
        this.lines[0] = "Nome: " + trimmedText;
        this._log('First line modified to: ' + this.lines[0]);
      } else {
        this.lines[0] = trimmedText;
      }
      return this;
    }

    /**
     * Processes all lines according to the configured formatting rules and placeholders.
     * This is the main processing method that orchestrates the manipulation.
     * @returns {TextManipulator} The instance for chaining.
     */
    processLines() {
      if (this.lines === null) {
        this._log('Lines not available. Run splitInLines() first.');
        return this;
      }

      this.lines.forEach(line => {
        let currentLine = this._trimSpacesAndTest(line);

        // Check if the line is a name declaration (e.g., "Nome: John Doe")
        if (currentLine.trim().toUpperCase().startsWith("NOME:")) {
            const extractedName = this._extractName(currentLine);
            if (extractedName) {
              this.currentPatientName = capitalizeFirstLetter(extractedName);
            }
        }

        currentLine = this._replacePlaceholders(currentLine);

        // Check HEADING1 rules first
        const heading1Rule = this.heading1Rules.find(rule => rule.condition(currentLine));
        if (heading1Rule) {
          this._handleAppliedRule(heading1Rule, currentLine);
          return;
        }

        // Check HEADING2 rules
        const heading2Rule = this.heading2Rules.find(rule => rule.condition(currentLine));
        if (heading2Rule) {
          this._handleAppliedRule(heading2Rule, currentLine);
          return;
        }

        // If no rules match, add as regular line
        this.processedLines.push({
          text: currentLine,
          heading: null
        });
      });

      this._log('Finished processing all lines.');
      return this;
    }

    /**
     * Structures the document by grouping H2 sections under their parent H1,
     * optionally prioritizing specific H2s, and ensuring the H1 is repeated for each section.
     * @param {Array<string>|null} [priorityHeadlines=null] - A list of H2 texts to place first, in order.
     * @returns {TextManipulator} The instance for chaining.
     */
    structureH2Blocks(priorityHeadlines = null) {
      if (this.processedLines.length === 0) {
        this._log('No lines to structure.');
        return this;
      }

      const h1Groups = new Map();
      let currentH1 = null;
      let currentH2Block = null;

      // Etapa 1: Agrupar linhas, criando H1s e H2s padrão conforme necessário.
      this.processedLines.forEach(item => {
        if (item.heading === DocumentApp.ParagraphHeading.HEADING1) {
          currentH1 = item;
          currentH2Block = null; // Redefinir para um novo H1
          if (!h1Groups.has(currentH1.text)) {
            h1Groups.set(currentH1.text, { h1: currentH1, h2Blocks: [] });
          }
        } else if (item.heading === DocumentApp.ParagraphHeading.HEADING2) {
          if (!currentH1) {
            currentH1 = {
              text: "Nome: PACIENTE INDEFINIDO",
              heading: DocumentApp.ParagraphHeading.HEADING1
            };
            h1Groups.set(currentH1.text, { h1: currentH1, h2Blocks: [] });
          }
          currentH2Block = [item];
          h1Groups.get(currentH1.text).h2Blocks.push(currentH2Block);
        } else { // Parágrafo de texto regular
          if (!currentH1) {
            currentH1 = {
              text: "Nome: PACIENTE INDEFINIDO",
              heading: DocumentApp.ParagraphHeading.HEADING1
            };
            h1Groups.set(currentH1.text, { h1: currentH1, h2Blocks: [] });
          }
          if (!currentH2Block) {
            const defaultH2 = {
              text: "Documento Indefinido",
              heading: DocumentApp.ParagraphHeading.HEADING2
            };
            currentH2Block = [defaultH2];
            h1Groups.get(currentH1.text).h2Blocks.push(currentH2Block);
          }
          currentH2Block.push(item);
        }
      });
      this._log(`Finished grouping. H1 groups count: ${h1Groups.size}`);

      // Etapa 3: Reordenar blocos H2 dentro de cada grupo H1 por prioridade.
      if (priorityHeadlines && priorityHeadlines.length > 0) {
        const priorityTexts = priorityHeadlines.map(h => h.trim().toUpperCase());

        h1Groups.forEach(group => {
          group.h2Blocks.sort((a, b) => {
            const aText = a[0].text.trim().toUpperCase();
            const bText = b[0].text.trim().toUpperCase();

            const aIndex = priorityTexts.indexOf(aText);
            const bIndex = priorityTexts.indexOf(bText);

            if (aIndex !== -1 && bIndex !== -1) {
              return aIndex - bIndex; // Both are priority, sort by index
            }
            if (aIndex !== -1) {
              return -1; // a is priority, b is not
            }
            if (bIndex !== -1) {
              return 1; // b is priority, a is not
            }
            return 0; // neither is priority
          });
        });
        this._log(`H2 blocks reordered based on priorities: "${priorityHeadlines.join(', ')}"`);
      }

      // Etapa 4: Planificar a estrutura agrupada de volta para uma lista de linhas.
      const finalStructuredLines = [];
      h1Groups.forEach(group => {
        if (group.h2Blocks.length === 0) {
          finalStructuredLines.push(group.h1);
        } else {
          group.h2Blocks.forEach(block => {
            finalStructuredLines.push(group.h1);
            finalStructuredLines.push(...block);
          });
        }
      });

      this.processedLines = finalStructuredLines;
      this._log(`Final line structure rebuilt. Total lines: ${this.processedLines.length}`);
      return this;
    }

    /**
     * Returns the final processed text.
     * @param {string} [joiner='\n'] - The character to join lines with.
     * @returns {string} The final text.
     */
    getResult(joiner = '\n') {
      return this.processedLines.map(item => item.text).join(joiner);
    }

    /**
     * Logs the array of dictionaries for debugging purposes.
     * @returns {TextManipulator} The instance for chaining.
     */
    logDictionaries() {
      if (this.shouldLog) {
        this._log('=== Processed Lines Dictionaries ===');
        this.processedLines.forEach((item, index) => {
          this._log(`[${index}] text: "${item.text}", heading: ${item.heading}`);
        });
        this._log('=== End of Dictionaries ===');
      }
      return this;
    }

    /**
     * Internal log helper.
     * @private
     */
    _log(message) {
      if (this.shouldLog) {
        // In a Google Apps Script environment, Logger.log is preferred.
        if (typeof Logger !== 'undefined') {
          Logger.log(message);
        } else {
          console.log(message);
        }
      }
    }

    /**
     * Trims leading spaces if the line contains a mix of spaces and characters.
     * @private
     */
    _trimSpacesAndTest(text) {
      // If the line is only spaces, preserve it as is
      if (text.trim() === '') {
        return text;
      }

      // If the line starts with spaces followed by non-space characters, trim leading spaces
      const regex = /^\s+\S+/;
      return regex.test(text) ? text.trim() : text;
    }

    /**
     * Replaces placeholders in the text.
     * @param {string} text The text to process.
     * @returns {string} The text with placeholders replaced.
     */
    _replacePlaceholders(text) {
      if (!this.placeHolders || this.placeHolders.length === 0) {
        return text;
      }
      let modifiedText = text;
      this.placeHolders.forEach(rule => {
        const placeholderRegex = new RegExp(this._escapeRegExp(rule.text), 'gi'); // Case-insensitive
        if (placeholderRegex.test(modifiedText)) {
          const replacementValue = typeof rule.replacement === 'function'
            ? rule.replacement(this.currentPatientName)
            : rule.replacement;

          modifiedText = modifiedText.replace(placeholderRegex, replacementValue);
        }
      });
      return modifiedText;
    }

    /**
     * Handles a line where a formatting rule was applied.
     * @private
     */
    _handleAppliedRule(appliedRule, text) {
      if (this._isHeading1(appliedRule)) {
        let heading1Text = text;
        const regexNomeVariations = /^(NOME:|Nome:|nome:|nOME:)/i;
        if (regexNomeVariations.test(heading1Text)) {
          heading1Text = this._formatNomeLine(heading1Text);
        }
        this._storeHeading1(heading1Text);
      } else if (this._isHeading2(appliedRule)) {
        this.processedLines.push({
          text: text,
          heading: DocumentApp.ParagraphHeading.HEADING2
        });
      } else {
        this.processedLines.push({
          text: text,
          heading: null
        });
      }
    }

    /**
     * Formats a "Nome:" line to a consistent case.
     * @private
     */
    _formatNomeLine(text) {
      const regexNomeVariations = /^(NOME:|Nome:|nome:|nOME:)\s*/i;
      const processedText = text.replace(regexNomeVariations, "Nome: ");
      return processedText.replace(/Nome: (.+)/, (_, namePart) => "Nome: " + namePart.toUpperCase());
    }

    /**
     * Stores a HEADING1 line and updates the lastHeading1 reference.
     * @private
     */
    _storeHeading1(text) {
      this.processedLines.push({
        text: text,
        heading: DocumentApp.ParagraphHeading.HEADING1
      });
      this.lastHeading1 = text;
    }

    /**
     * Extracts the first and full name from a "Nome:" heading.
     * @private
     */
    _extractName(heading1Text) {
      const parts = heading1Text.split(":");
      if (parts.length > 1) {
        return parts[1].trim();
      }
      return null;
    }

    /**
     * Checks if a rule is for HEADING1.
     * @private
     */
    _isHeading1(appliedRule) {
      return appliedRule.heading === DocumentApp.ParagraphHeading.HEADING1;
    }

    /**
     * Checks if a rule is for HEADING2.
     * @private
     */
    _isHeading2(appliedRule) {
      return appliedRule.heading === DocumentApp.ParagraphHeading.HEADING2;
    }

    /**
     * Escapes a string for use in a regular expression.
     * @private
     */
    _escapeRegExp(string) {
      return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); // $& means the whole matched string
    }
  },

/**
 * Capitalizes the first letter of a string.
 * @param {string} str - The string to capitalize.
 * @returns {string} The string with the first letter capitalized.
 */
const DocumentService = {
  executeFormat: function() {
    const headlineManager = new this.DocumentHeadlineManager();
    let patientName = null;

    if (headlineManager.bodyStructure && headlineManager.bodyStructure.length > 0) {
      const firstHeadline1 = headlineManager.bodyStructure[0].headline1;
      if (firstHeadline1) {
        const parts = firstHeadline1.split(':');
        if (parts.length > 1) {
          const extractedName = parts[1].trim();
          patientName = this.capitalizeFirstLetter(extractedName);
        }
      }
    }

    const bodyText = DocumentApp.getActiveDocument().getBody().getText();
    const rulesForPlaceholders = this.rules.getRulesForPlaceholders();
    const rulesForWholeParagraphs = this.rules.getRulesForWholeParagraphs();
    const rulesForHeading1 = this.filterRulesByHeading(rulesForWholeParagraphs, DocumentApp.ParagraphHeading.HEADING1);
    const rulesForHeading2 = this.filterRulesByHeading(rulesForWholeParagraphs, DocumentApp.ParagraphHeading.HEADING2);
    const rulesForInsideParagraphs = this.rules.getRulesForInsideParagraphs();
    const priorityHeadlines = ["Prontuário Médico", "Laudo de Mapeamento de Retina"];
    const textManipulator = new this.TextManipulator(bodyText, rulesForPlaceholders, rulesForHeading1, rulesForHeading2);
    textManipulator.currentPatientName = patientName;

    const processedLines = textManipulator
      .withLogging(false)
      .splitInLines()
      .removeEmptyLines()
      .trimLeadingSpaces()
      .removeEscapeCharacters()
      .checkAndModifyFirstLine()
      .processLines()
      .structureH2Blocks(priorityHeadlines)
      .logDictionaries()
      .processedLines;

    const paragraphManipulator = new this.ParagraphManipulator(processedLines, rulesForWholeParagraphs, rulesForInsideParagraphs);
    paragraphManipulator
      .withLogging(false)
      .clearDocument()
      .formatParagraphs()
      .formatTextInsideParagraphs()
      .cleanupDocument();
  }
};

  filterRulesByHeading: function(rules, heading) {
    return rules.filter(function(rule) {
      return rule.heading === heading;
    });
  },

  capitalizeFirstLetter: function(str) {
    if (!str || typeof str !== 'string') return str;
    return str.split(' ').map(word => word.charAt(0).toUpperCase() + word.slice(1).toLowerCase()).join(' ');
  },

  rules: {
    getRulesForInsideParagraphs: function() {
      return [
        {
          styleName: 'italic',
          targets: ["screening", "floaters", "flashes", "lattice", "leak point"],
        },
        {
          styleName: 'bold',
          targets: ["No olho direito", "no olho direito", "No olho esquerdo", "no olho esquerdo", "Em ambos os olhos", "em ambos os olhos", "com urgência", "063 3219 9700"],
        },
        {
          styleName: 'underline',
          targets: ["OD:", "OE:", "AO:", "ORIENTAÇÕES", "CID 10:", "composição:"],
        },
      ];
    },

    getRulesForPlaceholders: function() {
      return [
        { text: "{Nome}", replacement: (patientName) => { if (!patientName || typeof patientName !== 'string') return "Nome"; const firstName = patientName.split(' ')[0]; return firstName; } },
        { text: "{Nome Completo}", replacement: (patientName) => patientName || "Nome Completo" },
        { text: "{Secretária do consultório 13}", replacement: PropertiesService.getScriptProperties().getProperty('secretary13') },
        { text: "{Secretária do consultório 12}", replacement: PropertiesService.getScriptProperties().getProperty('secretary12') },
        { text: "{HH:MM}", replacement: () => { const now = new Date(); let hours = now.getHours(); let minutes = now.getMinutes(); let roundedMinutes = Math.ceil(minutes / 15) * 15; if (roundedMinutes === 60) { roundedMinutes = 0; hours = (hours + 1) % 24; } const hoursStr = hours.toString().padStart(2, '0'); const minutesStr = roundedMinutes.toString().padStart(2, '0'); return `${hoursStr}:${minutesStr}`; } },
        { text: "{À}", replacement: "À disposição para quaisquer outros esclarecimentos." },
        { text: "{prontuário médico}", replacement: "Prontuário Médico" },
        { text: "{encaminhamento}", replacement: "Encaminhamento" },
        { text: "{laudo de mapeamento de retina}", replacement: "Laudo de Mapeamento de Retina" },
        { text: "{relatório médico}", replacement: "Relatório Médico" },
        { text: "{laudo de tomografia de coerência óptica}", replacement: "Laudo de Tomografia de Coerência Óptica" },
        { text: "{laudo de angiografia fluoresceínica digital}", replacement: "Laudo de Angiografia Fluoresceínica Digital" },
        { text: "{laudo de retinografia colorida digital}", replacement: "Laudo de Retinografia Colorida Digital" }
      ];
    },

    getRulesForWholeParagraphs: function() {
      return [
        {
        keywords: ['OD:', 'OE:', 'AO:', 'ADIÇÃO:'], // e também "Papila normocorada;" -> linhas mais comuns antes!
        heading: DocumentApp.ParagraphHeading.NORMAL,
        requiresPageBreak: false,
        condition: function(line) {
          var text = line.toUpperCase().trim();
          return this.keywords.some(keyword => text.startsWith(keyword)) || text.endsWith(";"); // Verifica se a linha começa com um dos keywords ou termina com ";"
        },
        style: {
          [DocumentApp.Attribute.BACKGROUND_COLOR]: null,
          [DocumentApp.Attribute.BOLD]: null,
          [DocumentApp.Attribute.FONT_FAMILY]: 'IBM Plex Serif',
          [DocumentApp.Attribute.FONT_SIZE]: 10,
          [DocumentApp.Attribute.FOREGROUND_COLOR]: null,
          [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.LEFT,
          [DocumentApp.Attribute.INDENT_END]: 0,
          [DocumentApp.Attribute.INDENT_FIRST_LINE]: 50,
          [DocumentApp.Attribute.INDENT_START]: 0,
          [DocumentApp.Attribute.ITALIC]: null,
          [DocumentApp.Attribute.LEFT_TO_RIGHT]: true,
          [DocumentApp.Attribute.LINE_SPACING]: 1,
          [DocumentApp.Attribute.SPACING_AFTER]: 0,
          [DocumentApp.Attribute.SPACING_BEFORE]: 0,
          [DocumentApp.Attribute.STRIKETHROUGH]: null,
          [DocumentApp.Attribute.UNDERLINE]: null
        }
      },
      {
        keywords: ['NOME:', 'REF:', 'PACIENTE:'],
        heading: DocumentApp.ParagraphHeading.HEADING1,
        requiresPageBreak: true,
        condition: function(line) {
          var text = line.toUpperCase();
          return this.keywords.some(keyword => text.startsWith(keyword.toUpperCase()));
        },
        style: {
          [DocumentApp.Attribute.BACKGROUND_COLOR]: null,
          [DocumentApp.Attribute.BOLD]: true,
          [DocumentApp.Attribute.FONT_FAMILY]: 'IBM Plex Sans',
          [DocumentApp.Attribute.FONT_SIZE]: 16,
          [DocumentApp.Attribute.FOREGROUND_COLOR]: null,
          [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.LEFT,
          [DocumentApp.Attribute.INDENT_END]: 0,
          [DocumentApp.Attribute.INDENT_FIRST_LINE]: 0,
          [DocumentApp.Attribute.INDENT_START]: 0,
          [DocumentApp.Attribute.ITALIC]: null,
          [DocumentApp.Attribute.LEFT_TO_RIGHT]: true,
          [DocumentApp.Attribute.LINE_SPACING]: 1,
          [DocumentApp.Attribute.SPACING_AFTER]: 0,
          [DocumentApp.Attribute.SPACING_BEFORE]: 0,
          [DocumentApp.Attribute.STRIKETHROUGH]: null,
          [DocumentApp.Attribute.UNDERLINE]: null
        }
      },
      {
        keywords: [
          'PRONTUÁRIO', 'LAUDO', 'RECEITUÁRIO', 'PRESCRIÇÃO', 'RECEITA',
          'RELATÓRIO', 'ATESTADO', 'SOLICITAÇÃO', 'ORIENTAÇÃO', 'ENCAMINHAMENTO',
          'RECOMENDAÇÕES', 'DESCRIÇÃO', 'GUIA','JUSTIFICATIVA', 'TRANSCRIÇÃO', 'DOCUMENTO INDEFINIDO'
        ],
        heading: DocumentApp.ParagraphHeading.HEADING2,
        requiresPageBreak: false,
        condition: function(line) {
          var text = line.toUpperCase(); // Não sensível a maiúsculas e minúsculas
          return this.keywords.some(keyword => text.startsWith(keyword.toUpperCase()));
        },
        style: {
          [DocumentApp.Attribute.BACKGROUND_COLOR]: null,
          [DocumentApp.Attribute.BOLD]: true,
          [DocumentApp.Attribute.FONT_FAMILY]: 'IBM Plex Sans',
          [DocumentApp.Attribute.FONT_SIZE]: 18,
          [DocumentApp.Attribute.FOREGROUND_COLOR]: null,
          [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.CENTER,
          [DocumentApp.Attribute.INDENT_END]: 0,
          [DocumentApp.Attribute.INDENT_FIRST_LINE]: 0,
          [DocumentApp.Attribute.INDENT_START]: 0,
          [DocumentApp.Attribute.ITALIC]: null,
          [DocumentApp.Attribute.LEFT_TO_RIGHT]: true,
          [DocumentApp.Attribute.LINE_SPACING]: 1,
          [DocumentApp.Attribute.SPACING_AFTER]: 0,
          [DocumentApp.Attribute.SPACING_BEFORE]: 20,
          [DocumentApp.Attribute.STRIKETHROUGH]: null,
          [DocumentApp.Attribute.UNDERLINE]: null
        }
      },
      {
        keywords: ['História', 'Exame Físico', 'Exames Complementares', 'Hipótese Diagnóstica', 'Conclusão', 'Conduta', 'Prezad', 'Uso', 'Orientações' ],
        heading: DocumentApp.ParagraphHeading.NORMAL,
        requiresPageBreak: false,
        condition: function(line) {
          // Verifica se a linha começa com uma palavra-chave que inicia com maiúscula
          return this.keywords.some(keyword => line.startsWith(keyword));
        },
        style: {
          [DocumentApp.Attribute.BACKGROUND_COLOR]: null,
          [DocumentApp.Attribute.BOLD]: true,
          [DocumentApp.Attribute.FONT_FAMILY]: 'IBM Plex Mono',
          [DocumentApp.Attribute.FONT_SIZE]: 14,
          [DocumentApp.Attribute.FOREGROUND_COLOR]: null,
          [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.LEFT,
          [DocumentApp.Attribute.INDENT_END]: 0,
          [DocumentApp.Attribute.INDENT_FIRST_LINE]: 50,
          [DocumentApp.Attribute.INDENT_START]: 0,
          [DocumentApp.Attribute.ITALIC]: null,
          [DocumentApp.Attribute.LEFT_TO_RIGHT]: true,
          [DocumentApp.Attribute.LINE_SPACING]: 1,
          [DocumentApp.Attribute.SPACING_AFTER]: 0,
          [DocumentApp.Attribute.SPACING_BEFORE]: 10,
          [DocumentApp.Attribute.STRIKETHROUGH]: null,
          [DocumentApp.Attribute.UNDERLINE]: null
        }
      },
      {
        keywords: [],
        heading: DocumentApp.ParagraphHeading.NORMAL,
        requiresPageBreak: false,
        condition: function(line) {
          var text = line.trim();
          var regex = /^[A-Z][0-9]{2}[.][0-9]{1}\s*\(.*\)$/; // H10.0 (Conjuntivite mucopurulenta)
          return regex.test(text);
        },
        style: {
          [DocumentApp.Attribute.BACKGROUND_COLOR]: null,
          [DocumentApp.Attribute.BOLD]: true,
          [DocumentApp.Attribute.FONT_FAMILY]: 'IBM Plex Serif',
          [DocumentApp.Attribute.FONT_SIZE]: 12,
          [DocumentApp.Attribute.FOREGROUND_COLOR]: null,
          [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.LEFT,
          [DocumentApp.Attribute.INDENT_END]: 0,
          [DocumentApp.Attribute.INDENT_FIRST_LINE]: 50,
          [DocumentApp.Attribute.INDENT_START]: 0,
          [DocumentApp.Attribute.ITALIC]: null,
          [DocumentApp.Attribute.LEFT_TO_RIGHT]: true,
          [DocumentApp.Attribute.LINE_SPACING]: 1,
          [DocumentApp.Attribute.SPACING_AFTER]: 0,
          [DocumentApp.Attribute.SPACING_BEFORE]: 0,
          [DocumentApp.Attribute.STRIKETHROUGH]: null,
          [DocumentApp.Attribute.UNDERLINE]: null
        }
      },
      {
        keywords: [],
        heading: DocumentApp.ParagraphHeading.NORMAL,
        requiresPageBreak: false,
        condition: function(line) { // Simbrinza (colírio de uso contínuo)
          var text = line.trim();
          var startsWithLetter = /^[a-zA-ZçÇãÃõÕáÁéÉíÍóÓúÚâÂêÊîÎôÔûÛàÀèÈìÌòÒùÙ]/.test(text.charAt(0));
          var endsWithParenthesis = text.endsWith(")");
          return startsWithLetter && endsWithParenthesis;
        },
        style: {
          [DocumentApp.Attribute.BACKGROUND_COLOR]: null,
          [DocumentApp.Attribute.BOLD]: true,
          [DocumentApp.Attribute.FONT_FAMILY]: 'IBM Plex Sans',
          [DocumentApp.Attribute.FONT_SIZE]: 12,
          [DocumentApp.Attribute.FOREGROUND_COLOR]: null,
          [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.LEFT,
          [DocumentApp.Attribute.INDENT_END]: 0,
          [DocumentApp.Attribute.INDENT_FIRST_LINE]: 50,
          [DocumentApp.Attribute.INDENT_START]: 0,
          [DocumentApp.Attribute.ITALIC]: null,
          [DocumentApp.Attribute.LEFT_TO_RIGHT]: true,
          [DocumentApp.Attribute.LINE_SPACING]: 1,
          [DocumentApp.Attribute.SPACING_AFTER]: 0,
          [DocumentApp.Attribute.SPACING_BEFORE]: 10,
          [DocumentApp.Attribute.STRIKETHROUGH]: null,
          [DocumentApp.Attribute.UNDERLINE]: null
        }
      },
      {
        keywords: ['OLHO DIREITO', 'OLHO ESQUERDO'],
        heading: DocumentApp.ParagraphHeading.NORMAL,
        requiresPageBreak: false,
        condition: function(line) {
          var text = line.trim();
          var regex = /^[A-ZÀ-Ÿ0-9\s:/çÇãÃõÕáÁéÉíÍóÓúÚâÂêÊîÎôÔûÛàÀèÈìÌòÒùÙ]+:$/;  // "ACUIDADE VISUAL SEM CORREÇÃO:" e "REFRAÇÃO:"
          return !text.startsWith('CID') && (regex.test(text) || this.keywords.some(keyword => text.includes(keyword)));
        },
        style: {
          [DocumentApp.Attribute.BACKGROUND_COLOR]: null,
          [DocumentApp.Attribute.BOLD]: null,
          [DocumentApp.Attribute.FONT_FAMILY]: 'IBM Plex Serif',
          [DocumentApp.Attribute.FONT_SIZE]: 12,
          [DocumentApp.Attribute.FOREGROUND_COLOR]: null,
          [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.LEFT,
          [DocumentApp.Attribute.INDENT_END]: 0,
          [DocumentApp.Attribute.INDENT_FIRST_LINE]: 50,
          [DocumentApp.Attribute.INDENT_START]: 0,
          [DocumentApp.Attribute.ITALIC]: null,
          [DocumentApp.Attribute.LEFT_TO_RIGHT]: true,
          [DocumentApp.Attribute.LINE_SPACING]: 1,
          [DocumentApp.Attribute.SPACING_AFTER]: 0,
          [DocumentApp.Attribute.SPACING_BEFORE]: 5,
          [DocumentApp.Attribute.STRIKETHROUGH]: null,
          [DocumentApp.Attribute.UNDERLINE]: null
        }
      },
      {
        keywords: [], // composição: carmelose 5mg/mL
        heading: DocumentApp.ParagraphHeading.NORMAL,
        requiresPageBreak: false,
        condition: function(line) {
          return line.startsWith("composição:");
        },
        style: {
          [DocumentApp.Attribute.BACKGROUND_COLOR]: null,
          [DocumentApp.Attribute.BOLD]: null,
          [DocumentApp.Attribute.FONT_FAMILY]: 'IBM Plex Sans Condensed',
          [DocumentApp.Attribute.FONT_SIZE]: 8,
          [DocumentApp.Attribute.FOREGROUND_COLOR]: null,
          [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.RIGHT,
          [DocumentApp.Attribute.INDENT_END]: 0,
          [DocumentApp.Attribute.INDENT_FIRST_LINE]: 0,
          [DocumentApp.Attribute.INDENT_START]: 0,
          [DocumentApp.Attribute.ITALIC]: null,
          [DocumentApp.Attribute.LEFT_TO_RIGHT]: true,
          [DocumentApp.Attribute.LINE_SPACING]: 1,
          [DocumentApp.Attribute.SPACING_AFTER]: 0,
          [DocumentApp.Attribute.SPACING_BEFORE]: 0,
          [DocumentApp.Attribute.STRIKETHROUGH]: null,
          [DocumentApp.Attribute.UNDERLINE]: null
        }
      },
      {
        keywords: ['CID'],
        heading: DocumentApp.ParagraphHeading.NORMAL,
        requiresPageBreak: false,
        condition: function(line) {
          var text = line.trim();
          return this.keywords.some(keyword => text.startsWith(keyword)) && text.endsWith(":");
        },
        style: {
          [DocumentApp.Attribute.BACKGROUND_COLOR]: null,
          [DocumentApp.Attribute.BOLD]: null,
          [DocumentApp.Attribute.FONT_FAMILY]: 'IBM Plex Serif',
          [DocumentApp.Attribute.FONT_SIZE]: 12,
          [DocumentApp.Attribute.FOREGROUND_COLOR]: null,
          [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.LEFT,
          [DocumentApp.Attribute.INDENT_END]: 0,
          [DocumentApp.Attribute.INDENT_FIRST_LINE]: 50,
          [DocumentApp.Attribute.INDENT_START]: 0,
          [DocumentApp.Attribute.ITALIC]: null,
          [DocumentApp.Attribute.LEFT_TO_RIGHT]: true,
          [DocumentApp.Attribute.LINE_SPACING]: 1,
          [DocumentApp.Attribute.SPACING_AFTER]: 0,
          [DocumentApp.Attribute.SPACING_BEFORE]: 0,
          [DocumentApp.Attribute.STRIKETHROUGH]: null,
          [DocumentApp.Attribute.UNDERLINE]: null
        }
      },
      {
        keywords: ["À disposição", "Auxiliar"],
        heading: DocumentApp.ParagraphHeading.NORMAL,
        requiresPageBreak: false,
        condition: function(line) {
          var text = line.trim();
          return this.keywords.some(keyword => text.startsWith(keyword));
        },
        style: {
          [DocumentApp.Attribute.BACKGROUND_COLOR]: null,
          [DocumentApp.Attribute.BOLD]: null,
          [DocumentApp.Attribute.FONT_FAMILY]: 'IBM Plex Serif',
          [DocumentApp.Attribute.FONT_SIZE]: 12,
          [DocumentApp.Attribute.FOREGROUND_COLOR]: null,
          [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.RIGHT,
          [DocumentApp.Attribute.INDENT_END]: 0,
          [DocumentApp.Attribute.INDENT_FIRST_LINE]: 50,
          [DocumentApp.Attribute.INDENT_START]: 0,
          [DocumentApp.Attribute.ITALIC]: null,
          [DocumentApp.Attribute.LEFT_TO_RIGHT]: true,
          [DocumentApp.Attribute.LINE_SPACING]: 1.5,
          [DocumentApp.Attribute.SPACING_AFTER]: 0,
          [DocumentApp.Attribute.SPACING_BEFORE]: 15,
          [DocumentApp.Attribute.STRIKETHROUGH]: null,
          [DocumentApp.Attribute.UNDERLINE]: null
        }
      },
      {
        keywords: [],
        heading: DocumentApp.ParagraphHeading.NORMAL,
        requiresPageBreak: false,
        condition: function(line) {
          return true; // Aplica-se a todos os parágrafos que não correspondem a outras regras
        },
        style: {
          [DocumentApp.Attribute.BACKGROUND_COLOR]: null,
          [DocumentApp.Attribute.BOLD]: null,
          [DocumentApp.Attribute.FONT_FAMILY]: 'IBM Plex Serif',
          [DocumentApp.Attribute.FONT_SIZE]: 12,
          [DocumentApp.Attribute.FOREGROUND_COLOR]: null,
          [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.LEFT,
          [DocumentApp.Attribute.INDENT_END]: 0,
          [DocumentApp.Attribute.INDENT_FIRST_LINE]: 50,
          [DocumentApp.Attribute.INDENT_START]: 0,
          [DocumentApp.Attribute.ITALIC]: null,
          [DocumentApp.Attribute.LEFT_TO_RIGHT]: true,
          [DocumentApp.Attribute.LINE_SPACING]: 1,
          [DocumentApp.Attribute.SPACING_AFTER]: 0,
          [DocumentApp.Attribute.SPACING_BEFORE]: 0,
          [DocumentApp.Attribute.STRIKETHROUGH]: null,
          [DocumentApp.Attribute.UNDERLINE]: null
        }
      }
    ];
    },

    getFormattingRules: function() {
      return this.getRulesForWholeParagraphs();
    }
  }
};

function getRulesForInsideParagraphs() {
    return [
      {
        styleName: 'italic',
        targets: [
          "screening",
          "floaters",
          "flashes",
          "lattice",
          "leak point"
        ],
      },
      {
        styleName: 'bold',
        targets: [
          "No olho direito",
          "no olho direito",
          "No olho esquerdo",
          "no olho esquerdo",
          "Em ambos os olhos",
          "em ambos os olhos",
          "com urgência",
          "063 3219 9700"
        ],
      },
      {
        styleName: 'underline',
        targets: [
          "OD:",
          "OE:",
          "AO:",
          "ORIENTAÇÕES",
          "CID 10:",
          "composição:"
        ],
      },
    ];
  }

function getRulesForPlaceholders() {
    return [
      {
        text: "{Nome}",
        replacement: (patientName) => {
          if (!patientName || typeof patientName !== 'string') return "Nome";
          const firstName = patientName.split(' ')[0];
          return firstName;
        }
      },
      {
        text: "{Nome Completo}",
        replacement: (patientName) => patientName || "Nome Completo"
      },
      {
        text: "{Secretária do consultório 13}",
        replacement: PropertiesService.getScriptProperties().getProperty('secretary13')
      },
      {
        text: "{Secretária do consultório 12}",
        replacement: PropertiesService.getScriptProperties().getProperty('secretary12')
      },
      {
        text: "{HH:MM}",
        replacement: (patientName) => {
          const now = new Date();
          let hours = now.getHours();
          let minutes = now.getMinutes();

          // Arredonda para o próximo múltiplo de 15 minutos
          let roundedMinutes = Math.ceil(minutes / 15) * 15;
          if (roundedMinutes === 60) {
            roundedMinutes = 0;
            hours = (hours + 1) % 24;
          }

          // Formata com zero à esquerda se necessário
          const hoursStr = hours.toString().padStart(2, '0');
          const minutesStr = roundedMinutes.toString().padStart(2, '0');
          return `${hoursStr}:${minutesStr}`;
        }
      },
      {
        text: "{À}",
        replacement: "À disposição para quaisquer outros esclarecimentos."
      },
      {
        text: "{prontuário médico}",
        replacement: "Prontuário Médico"
      },
      {
        text: "{encaminhamento}",
        replacement: "Encaminhamento"
      },
      {
        text: "{laudo de mapeamento de retina}",
        replacement: "Laudo de Mapeamento de Retina"
      },
      {
        text: "{relatório médico}",
        replacement: "Relatório Médico"
      },
      {
        text: "{laudo de tomografia de coerência óptica}",
        replacement: "Laudo de Tomografia de Coerência Óptica"
      },
      {
        text: "{laudo de angiografia fluoresceínica digital}",
        replacement: "Laudo de Angiografia Fluoresceínica Digital"
      },
      {
        text: "{laudo de retinografia colorida digital}",
        replacement: "Laudo de Retinografia Colorida Digital"
      }
      // Adicione mais placeholders conforme necessário
    ];
  }

function getRulesForWholeParagraphs() {
    return [
        {
        keywords: ['OD:', 'OE:', 'AO:', 'ADIÇÃO:'], // e também "Papila normocorada;" -> linhas mais comuns antes!
        heading: DocumentApp.ParagraphHeading.NORMAL,
        requiresPageBreak: false,
        condition: function(line) {
          var text = line.toUpperCase().trim();
          return this.keywords.some(keyword => text.startsWith(keyword)) || text.endsWith(";"); // Verifica se a linha começa com um dos keywords ou termina com ";"
        },
        style: {
          [DocumentApp.Attribute.BACKGROUND_COLOR]: null,
          [DocumentApp.Attribute.BOLD]: null,
          [DocumentApp.Attribute.FONT_FAMILY]: 'IBM Plex Serif',
          [DocumentApp.Attribute.FONT_SIZE]: 10,
          [DocumentApp.Attribute.FOREGROUND_COLOR]: null,
          [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.LEFT,
          [DocumentApp.Attribute.INDENT_END]: 0,
          [DocumentApp.Attribute.INDENT_FIRST_LINE]: 50,
          [DocumentApp.Attribute.INDENT_START]: 0,
          [DocumentApp.Attribute.ITALIC]: null,
          [DocumentApp.Attribute.LEFT_TO_RIGHT]: true,
          [DocumentApp.Attribute.LINE_SPACING]: 1,
          [DocumentApp.Attribute.SPACING_AFTER]: 0,
          [DocumentApp.Attribute.SPACING_BEFORE]: 0,
          [DocumentApp.Attribute.STRIKETHROUGH]: null,
          [DocumentApp.Attribute.UNDERLINE]: null
        }
      },
      {
        keywords: ['NOME:', 'REF:', 'PACIENTE:'],
        heading: DocumentApp.ParagraphHeading.HEADING1,
        requiresPageBreak: true,
        condition: function(line) {
          var text = line.toUpperCase();
          return this.keywords.some(keyword => text.startsWith(keyword.toUpperCase()));
        },
        style: {
          [DocumentApp.Attribute.BACKGROUND_COLOR]: null,
          [DocumentApp.Attribute.BOLD]: true,
          [DocumentApp.Attribute.FONT_FAMILY]: 'IBM Plex Sans',
          [DocumentApp.Attribute.FONT_SIZE]: 16,
          [DocumentApp.Attribute.FOREGROUND_COLOR]: null,
          [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.LEFT,
          [DocumentApp.Attribute.INDENT_END]: 0,
          [DocumentApp.Attribute.INDENT_FIRST_LINE]: 0,
          [DocumentApp.Attribute.INDENT_START]: 0,
          [DocumentApp.Attribute.ITALIC]: null,
          [DocumentApp.Attribute.LEFT_TO_RIGHT]: true,
          [DocumentApp.Attribute.LINE_SPACING]: 1,
          [DocumentApp.Attribute.SPACING_AFTER]: 0,
          [DocumentApp.Attribute.SPACING_BEFORE]: 0,
          [DocumentApp.Attribute.STRIKETHROUGH]: null,
          [DocumentApp.Attribute.UNDERLINE]: null
        }
      },
      {
        keywords: [
          'PRONTUÁRIO', 'LAUDO', 'RECEITUÁRIO', 'PRESCRIÇÃO', 'RECEITA',
          'RELATÓRIO', 'ATESTADO', 'SOLICITAÇÃO', 'ORIENTAÇÃO', 'ENCAMINHAMENTO',
          'RECOMENDAÇÕES', 'DESCRIÇÃO', 'GUIA','JUSTIFICATIVA', 'TRANSCRIÇÃO', 'DOCUMENTO INDEFINIDO'
        ],
        heading: DocumentApp.ParagraphHeading.HEADING2,
        requiresPageBreak: false,
        condition: function(line) {
          var text = line.toUpperCase(); // Não sensível a maiúsculas e minúsculas
          return this.keywords.some(keyword => text.startsWith(keyword.toUpperCase()));
        },
        style: {
          [DocumentApp.Attribute.BACKGROUND_COLOR]: null,
          [DocumentApp.Attribute.BOLD]: true,
          [DocumentApp.Attribute.FONT_FAMILY]: 'IBM Plex Sans',
          [DocumentApp.Attribute.FONT_SIZE]: 18,
          [DocumentApp.Attribute.FOREGROUND_COLOR]: null,
          [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.CENTER,
          [DocumentApp.Attribute.INDENT_END]: 0,
          [DocumentApp.Attribute.INDENT_FIRST_LINE]: 0,
          [DocumentApp.Attribute.INDENT_START]: 0,
          [DocumentApp.Attribute.ITALIC]: null,
          [DocumentApp.Attribute.LEFT_TO_RIGHT]: true,
          [DocumentApp.Attribute.LINE_SPACING]: 1,
          [DocumentApp.Attribute.SPACING_AFTER]: 0,
          [DocumentApp.Attribute.SPACING_BEFORE]: 20,
          [DocumentApp.Attribute.STRIKETHROUGH]: null,
          [DocumentApp.Attribute.UNDERLINE]: null
        }
      },
      {
        keywords: ['História', 'Exame Físico', 'Exames Complementares', 'Hipótese Diagnóstica', 'Conclusão', 'Conduta', 'Prezad', 'Uso', 'Orientações' ],
        heading: DocumentApp.ParagraphHeading.NORMAL,
        requiresPageBreak: false,
        condition: function(line) {
          // Verifica se a linha começa com uma palavra-chave que inicia com maiúscula
          return this.keywords.some(keyword => line.startsWith(keyword));
        },
        style: {
          [DocumentApp.Attribute.BACKGROUND_COLOR]: null,
          [DocumentApp.Attribute.BOLD]: true,
          [DocumentApp.Attribute.FONT_FAMILY]: 'IBM Plex Mono',
          [DocumentApp.Attribute.FONT_SIZE]: 14,
          [DocumentApp.Attribute.FOREGROUND_COLOR]: null,
          [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.LEFT,
          [DocumentApp.Attribute.INDENT_END]: 0,
          [DocumentApp.Attribute.INDENT_FIRST_LINE]: 50,
          [DocumentApp.Attribute.INDENT_START]: 0,
          [DocumentApp.Attribute.ITALIC]: null,
          [DocumentApp.Attribute.LEFT_TO_RIGHT]: true,
          [DocumentApp.Attribute.LINE_SPACING]: 1,
          [DocumentApp.Attribute.SPACING_AFTER]: 0,
          [DocumentApp.Attribute.SPACING_BEFORE]: 10,
          [DocumentApp.Attribute.STRIKETHROUGH]: null,
          [DocumentApp.Attribute.UNDERLINE]: null
        }
      },
      {
        keywords: [],
        heading: DocumentApp.ParagraphHeading.NORMAL,
        requiresPageBreak: false,
        condition: function(line) {
          var text = line.trim();
          var regex = /^[A-Z][0-9]{2}[.][0-9]{1}\s*\(.*\)$/; // H10.0 (Conjuntivite mucopurulenta)
          return regex.test(text);
        },
        style: {
          [DocumentApp.Attribute.BACKGROUND_COLOR]: null,
          [DocumentApp.Attribute.BOLD]: true,
          [DocumentApp.Attribute.FONT_FAMILY]: 'IBM Plex Serif',
          [DocumentApp.Attribute.FONT_SIZE]: 12,
          [DocumentApp.Attribute.FOREGROUND_COLOR]: null,
          [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.LEFT,
          [DocumentApp.Attribute.INDENT_END]: 0,
          [DocumentApp.Attribute.INDENT_FIRST_LINE]: 50,
          [DocumentApp.Attribute.INDENT_START]: 0,
          [DocumentApp.Attribute.ITALIC]: null,
          [DocumentApp.Attribute.LEFT_TO_RIGHT]: true,
          [DocumentApp.Attribute.LINE_SPACING]: 1,
          [DocumentApp.Attribute.SPACING_AFTER]: 0,
          [DocumentApp.Attribute.SPACING_BEFORE]: 0,
          [DocumentApp.Attribute.STRIKETHROUGH]: null,
          [DocumentApp.Attribute.UNDERLINE]: null
        }
      },
      {
        keywords: [],
        heading: DocumentApp.ParagraphHeading.NORMAL,
        requiresPageBreak: false,
        condition: function(line) { // Simbrinza (colírio de uso contínuo)
          var text = line.trim();
          var startsWithLetter = /^[a-zA-ZçÇãÃõÕáÁéÉíÍóÓúÚâÂêÊîÎôÔûÛàÀèÈìÌòÒùÙ]/.test(text.charAt(0));
          var endsWithParenthesis = text.endsWith(")");
          return startsWithLetter && endsWithParenthesis;
        },
        style: {
          [DocumentApp.Attribute.BACKGROUND_COLOR]: null,
          [DocumentApp.Attribute.BOLD]: true,
          [DocumentApp.Attribute.FONT_FAMILY]: 'IBM Plex Sans',
          [DocumentApp.Attribute.FONT_SIZE]: 12,
          [DocumentApp.Attribute.FOREGROUND_COLOR]: null,
          [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.LEFT,
          [DocumentApp.Attribute.INDENT_END]: 0,
          [DocumentApp.Attribute.INDENT_FIRST_LINE]: 50,
          [DocumentApp.Attribute.INDENT_START]: 0,
          [DocumentApp.Attribute.ITALIC]: null,
          [DocumentApp.Attribute.LEFT_TO_RIGHT]: true,
          [DocumentApp.Attribute.LINE_SPACING]: 1,
          [DocumentApp.Attribute.SPACING_AFTER]: 0,
          [DocumentApp.Attribute.SPACING_BEFORE]: 10,
          [DocumentApp.Attribute.STRIKETHROUGH]: null,
          [DocumentApp.Attribute.UNDERLINE]: null
        }
      },
      {
        keywords: ['OLHO DIREITO', 'OLHO ESQUERDO'],
        heading: DocumentApp.ParagraphHeading.NORMAL,
        requiresPageBreak: false,
        condition: function(line) {
          var text = line.trim();
          var regex = /^[A-ZÀ-Ÿ0-9\s:/çÇãÃõÕáÁéÉíÍóÓúÚâÂêÊîÎôÔûÛàÀèÈìÌòÒùÙ]+:$/;  // "ACUIDADE VISUAL SEM CORREÇÃO:" e "REFRAÇÃO:"
          return !text.startsWith('CID') && (regex.test(text) || this.keywords.some(keyword => text.includes(keyword)));
        },
        style: {
          [DocumentApp.Attribute.BACKGROUND_COLOR]: null,
          [DocumentApp.Attribute.BOLD]: null,
          [DocumentApp.Attribute.FONT_FAMILY]: 'IBM Plex Serif',
          [DocumentApp.Attribute.FONT_SIZE]: 12,
          [DocumentApp.Attribute.FOREGROUND_COLOR]: null,
          [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.LEFT,
          [DocumentApp.Attribute.INDENT_END]: 0,
          [DocumentApp.Attribute.INDENT_FIRST_LINE]: 50,
          [DocumentApp.Attribute.INDENT_START]: 0,
          [DocumentApp.Attribute.ITALIC]: null,
          [DocumentApp.Attribute.LEFT_TO_RIGHT]: true,
          [DocumentApp.Attribute.LINE_SPACING]: 1,
          [DocumentApp.Attribute.SPACING_AFTER]: 0,
          [DocumentApp.Attribute.SPACING_BEFORE]: 5,
          [DocumentApp.Attribute.STRIKETHROUGH]: null,
          [DocumentApp.Attribute.UNDERLINE]: null
        }
      },
      {
        keywords: [], // composição: carmelose 5mg/mL
        heading: DocumentApp.ParagraphHeading.NORMAL,
        requiresPageBreak: false,
        condition: function(line) {
          return line.startsWith("composição:");
        },
        style: {
          [DocumentApp.Attribute.BACKGROUND_COLOR]: null,
          [DocumentApp.Attribute.BOLD]: null,
          [DocumentApp.Attribute.FONT_FAMILY]: 'IBM Plex Sans Condensed',
          [DocumentApp.Attribute.FONT_SIZE]: 8,
          [DocumentApp.Attribute.FOREGROUND_COLOR]: null,
          [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.RIGHT,
          [DocumentApp.Attribute.INDENT_END]: 0,
          [DocumentApp.Attribute.INDENT_FIRST_LINE]: 0,
          [DocumentApp.Attribute.INDENT_START]: 0,
          [DocumentApp.Attribute.ITALIC]: null,
          [DocumentApp.Attribute.LEFT_TO_RIGHT]: true,
          [DocumentApp.Attribute.LINE_SPACING]: 1,
          [DocumentApp.Attribute.SPACING_AFTER]: 0,
          [DocumentApp.Attribute.SPACING_BEFORE]: 0,
          [DocumentApp.Attribute.STRIKETHROUGH]: null,
          [DocumentApp.Attribute.UNDERLINE]: null
        }
      },
      {
        keywords: ['CID'],
        heading: DocumentApp.ParagraphHeading.NORMAL,
        requiresPageBreak: false,
        condition: function(line) {
          var text = line.trim();
          return this.keywords.some(keyword => text.startsWith(keyword)) && text.endsWith(":");
        },
        style: {
          [DocumentApp.Attribute.BACKGROUND_COLOR]: null,
          [DocumentApp.Attribute.BOLD]: null,
          [DocumentApp.Attribute.FONT_FAMILY]: 'IBM Plex Serif',
          [DocumentApp.Attribute.FONT_SIZE]: 12,
          [DocumentApp.Attribute.FOREGROUND_COLOR]: null,
          [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.LEFT,
          [DocumentApp.Attribute.INDENT_END]: 0,
          [DocumentApp.Attribute.INDENT_FIRST_LINE]: 50,
          [DocumentApp.Attribute.INDENT_START]: 0,
          [DocumentApp.Attribute.ITALIC]: null,
          [DocumentApp.Attribute.LEFT_TO_RIGHT]: true,
          [DocumentApp.Attribute.LINE_SPACING]: 1,
          [DocumentApp.Attribute.SPACING_AFTER]: 0,
          [DocumentApp.Attribute.SPACING_BEFORE]: 0,
          [DocumentApp.Attribute.STRIKETHROUGH]: null,
          [DocumentApp.Attribute.UNDERLINE]: null
        }
      },
      {
        keywords: ["À disposição", "Auxiliar"],
        heading: DocumentApp.ParagraphHeading.NORMAL,
        requiresPageBreak: false,
        condition: function(line) {
          var text = line.trim();
          return this.keywords.some(keyword => text.startsWith(keyword));
        },
        style: {
          [DocumentApp.Attribute.BACKGROUND_COLOR]: null,
          [DocumentApp.Attribute.BOLD]: null,
          [DocumentApp.Attribute.FONT_FAMILY]: 'IBM Plex Serif',
          [DocumentApp.Attribute.FONT_SIZE]: 12,
          [DocumentApp.Attribute.FOREGROUND_COLOR]: null,
          [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.RIGHT,
          [DocumentApp.Attribute.INDENT_END]: 0,
          [DocumentApp.Attribute.INDENT_FIRST_LINE]: 50,
          [DocumentApp.Attribute.INDENT_START]: 0,
          [DocumentApp.Attribute.ITALIC]: null,
          [DocumentApp.Attribute.LEFT_TO_RIGHT]: true,
          [DocumentApp.Attribute.LINE_SPACING]: 1.5,
          [DocumentApp.Attribute.SPACING_AFTER]: 0,
          [DocumentApp.Attribute.SPACING_BEFORE]: 15,
          [DocumentApp.Attribute.STRIKETHROUGH]: null,
          [DocumentApp.Attribute.UNDERLINE]: null
        }
      },
      {
        keywords: [],
        heading: DocumentApp.ParagraphHeading.NORMAL,
        requiresPageBreak: false,
        condition: function(line) {
          return true; // Aplica-se a todos os parágrafos que não correspondem a outras regras
        },
        style: {
          [DocumentApp.Attribute.BACKGROUND_COLOR]: null,
          [DocumentApp.Attribute.BOLD]: null,
          [DocumentApp.Attribute.FONT_FAMILY]: 'IBM Plex Serif',
          [DocumentApp.Attribute.FONT_SIZE]: 12,
          [DocumentApp.Attribute.FOREGROUND_COLOR]: null,
          [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.LEFT,
          [DocumentApp.Attribute.INDENT_END]: 0,
          [DocumentApp.Attribute.INDENT_FIRST_LINE]: 50,
          [DocumentApp.Attribute.INDENT_START]: 0,
          [DocumentApp.Attribute.ITALIC]: null,
          [DocumentApp.Attribute.LEFT_TO_RIGHT]: true,
          [DocumentApp.Attribute.LINE_SPACING]: 1,
          [DocumentApp.Attribute.SPACING_AFTER]: 0,
          [DocumentApp.Attribute.SPACING_BEFORE]: 0,
          [DocumentApp.Attribute.STRIKETHROUGH]: null,
          [DocumentApp.Attribute.UNDERLINE]: null
        }
      }
    ];
  }

/**
 * Retorna as regras de formatação para os títulos.
 * Esta função atua como um alias para getRulesForWholeParagraphs,
 * garantindo compatibilidade com o DocumentHeadlineManager.
 * @returns {Array} As regras de formatação.
 */
function getFormattingRules() {
  return getRulesForWholeParagraphs();
}
