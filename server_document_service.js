/**
 * @file server_document_service.js
 * @description Service for all document-related manipulations, including text processing, formatting, and structure management.
 * This service is implemented using the Module pattern (IIFE) to encapsulate private functions and classes,
 * exposing only a public API. This prevents global scope pollution and improves maintainability.
 */

const DocumentService = (() => {
  'use strict';

  // =================================================================
  // PRIVATE HELPER FUNCTIONS
  // =================================================================

  /**
   * Capitalizes the first letter of each word in a string.
   * @param {string} str The input string.
   * @returns {string} The capitalized string.
   */
  function _capitalizeFirstLetter(str) {
    if (!str || typeof str !== 'string') return str;
    return str.split(' ').map(word => word.charAt(0).toUpperCase() + word.slice(1).toLowerCase()).join(' ');
  }

  /**
   * Filters formatting rules by a specific heading type.
   * @param {Array<Object>} rules - The list of rule objects.
   * @param {DocumentApp.ParagraphHeading} heading - The heading type to filter by.
   * @returns {Array<Object>} The filtered list of rules.
   */
  function _filterRulesByHeading(rules, heading) {
    return rules.filter(rule => rule.heading === heading);
  }

  /**
   * Defines rules for applying specific styles (bold, italic, etc.) to text within paragraphs.
   * @returns {Array<Object>} A list of rules for inline text formatting.
   */
  function _getRulesForInsideParagraphs() {
    return [
      { styleName: 'italic', targets: ["screening", "floaters", "flashes", "lattice", "leak point"] },
      { styleName: 'bold', targets: ["No olho direito", "no olho direito", "No olho esquerdo", "no olho esquerdo", "Em ambos os olhos", "em ambos os olhos", "com urgência", "063 3219 9700"] },
      { styleName: 'underline', targets: ["OD:", "OE:", "AO:", "ORIENTAÇÕES", "CID 10:", "composição:"] },
    ];
  }

  /**
   * Defines rules for replacing placeholder text (e.g., "{Nome}") with dynamic content.
   * @returns {Array<Object>} A list of placeholder replacement rules.
   */
  function _getRulesForPlaceholders() {
    return [
      { text: "{Nome}", replacement: (patientName) => patientName ? patientName.split(' ')[0] : "Nome" },
      { text: "{Nome Completo}", replacement: (patientName) => patientName || "Nome Completo" },
      { text: "{Secretária do consultório 13}", replacement: PropertiesService.getScriptProperties().getProperty('secretary13') },
      { text: "{Secretária do consultório 12}", replacement: PropertiesService.getScriptProperties().getProperty('secretary12') },
      {
        text: "{HH:MM}",
        replacement: () => {
          const now = new Date();
          let hours = now.getHours();
          let minutes = now.getMinutes();
          let roundedMinutes = Math.ceil(minutes / 15) * 15;
          if (roundedMinutes === 60) {
            roundedMinutes = 0;
            hours = (hours + 1) % 24;
          }
          return `${hours.toString().padStart(2, '0')}:${roundedMinutes.toString().padStart(2, '0')}`;
        }
      },
      { text: "{À}", replacement: "À disposição para quaisquer outros esclarecimentos." },
      { text: "{prontuário médico}", replacement: "Prontuário Médico" },
      { text: "{encaminhamento}", replacement: "Encaminhamento" },
      { text: "{laudo de mapeamento de retina}", replacement: "Laudo de Mapeamento de Retina" },
      { text: "{relatório médico}", replacement: "Relatório Médico" },
      { text: "{laudo de tomografia de coerência óptica}", replacement: "Laudo de Tomografia de Coerência Óptica" },
      { text: "{laudo de angiografia fluoresceínica digital}", replacement: "Laudo de Angiografia Fluoresceínica Digital" },
      { text: "{laudo de retinografia colorida digital}", replacement: "Laudo de Retinografia Colorida Digital" }
    ];
  }

  /**
   * Defines a comprehensive set of rules for formatting entire paragraphs based on their content.
   * @returns {Array<Object>} A list of paragraph formatting rules.
   */
  function _getRulesForWholeParagraphs() {
    // This function remains large due to the declarative nature of styling rules.
    // Each object represents a distinct styling condition for a paragraph.
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
          [DocumentApp.Attribute.BACKGROUND_COLOR]: null, [DocumentApp.Attribute.BOLD]: null, [DocumentApp.Attribute.FONT_FAMILY]: 'IBM Plex Serif',
          [DocumentApp.Attribute.FONT_SIZE]: 10, [DocumentApp.Attribute.FOREGROUND_COLOR]: null, [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.LEFT,
          [DocumentApp.Attribute.INDENT_END]: 0, [DocumentApp.Attribute.INDENT_FIRST_LINE]: 50, [DocumentApp.Attribute.INDENT_START]: 0,
          [DocumentApp.Attribute.ITALIC]: null, [DocumentApp.Attribute.LEFT_TO_RIGHT]: true, [DocumentApp.Attribute.LINE_SPACING]: 1,
          [DocumentApp.Attribute.SPACING_AFTER]: 0, [DocumentApp.Attribute.SPACING_BEFORE]: 0, [DocumentApp.Attribute.STRIKETHROUGH]: null, [DocumentApp.Attribute.UNDERLINE]: null
        }
      },
      {
        keywords: ['NOME:', 'REF:', 'PACIENTE:'], heading: DocumentApp.ParagraphHeading.HEADING1, requiresPageBreak: true,
        condition: function(line) { return this.keywords.some(keyword => line.toUpperCase().startsWith(keyword.toUpperCase())); },
        style: {
          [DocumentApp.Attribute.BOLD]: true, [DocumentApp.Attribute.FONT_FAMILY]: 'IBM Plex Sans', [DocumentApp.Attribute.FONT_SIZE]: 16,
          [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.LEFT, [DocumentApp.Attribute.INDENT_FIRST_LINE]: 0,
          [DocumentApp.Attribute.SPACING_BEFORE]: 0
        }
      },
      {
        keywords: [
          'PRONTUÁRIO', 'LAUDO', 'RECEITUÁRIO', 'PRESCRIÇÃO', 'RECEITA', 'RELATÓRIO', 'ATESTADO', 'SOLICITAÇÃO',
          'ORIENTAÇÃO', 'ENCAMINHAMENTO', 'RECOMENDAÇÕES', 'DESCRIÇÃO', 'GUIA','JUSTIFICATIVA', 'TRANSCRIÇÃO', 'DOCUMENTO INDEFINIDO'
        ],
        heading: DocumentApp.ParagraphHeading.HEADING2, requiresPageBreak: false,
        condition: function(line) { return this.keywords.some(keyword => line.toUpperCase().startsWith(keyword.toUpperCase())); },
        style: {
          [DocumentApp.Attribute.BOLD]: true, [DocumentApp.Attribute.FONT_FAMILY]: 'IBM Plex Sans', [DocumentApp.Attribute.FONT_SIZE]: 18,
          [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.CENTER, [DocumentApp.Attribute.INDENT_FIRST_LINE]: 0,
          [DocumentApp.Attribute.SPACING_BEFORE]: 20
        }
      },
      {
        keywords: ['História', 'Exame Físico', 'Exames Complementares', 'Hipótese Diagnóstica', 'Conclusão', 'Conduta', 'Prezad', 'Uso', 'Orientações' ],
        heading: DocumentApp.ParagraphHeading.NORMAL, requiresPageBreak: false,
        condition: function(line) { return this.keywords.some(keyword => line.startsWith(keyword)); },
        style: {
          [DocumentApp.Attribute.BOLD]: true, [DocumentApp.Attribute.FONT_FAMILY]: 'IBM Plex Mono', [DocumentApp.Attribute.FONT_SIZE]: 14,
          [DocumentApp.Attribute.INDENT_FIRST_LINE]: 50, [DocumentApp.Attribute.SPACING_BEFORE]: 10,
        }
      },
      {
        keywords: [], heading: DocumentApp.ParagraphHeading.NORMAL, requiresPageBreak: false,
        condition: function(line) { return /^[A-Z][0-9]{2}[.][0-9]{1}\s*\(.*\)$/.test(line.trim()); },
        style: { [DocumentApp.Attribute.BOLD]: true, [DocumentApp.Attribute.FONT_FAMILY]: 'IBM Plex Serif', [DocumentApp.Attribute.FONT_SIZE]: 12, [DocumentApp.Attribute.INDENT_FIRST_LINE]: 50, }
      },
      {
        keywords: [], heading: DocumentApp.ParagraphHeading.NORMAL, requiresPageBreak: false,
        condition: function(line) {
          const text = line.trim();
          return /^[a-zA-ZçÇãÃõÕáÁéÉíÍóÓúÚâÂêÊîÎôÔûÛàÀèÈìÌòÒùÙ]/.test(text.charAt(0)) && text.endsWith(")");
        },
        style: {
          [DocumentApp.Attribute.BOLD]: true, [DocumentApp.Attribute.FONT_FAMILY]: 'IBM Plex Sans', [DocumentApp.Attribute.FONT_SIZE]: 12,
          [DocumentApp.Attribute.INDENT_FIRST_LINE]: 50, [DocumentApp.Attribute.SPACING_BEFORE]: 10,
        }
      },
      {
        keywords: ['OLHO DIREITO', 'OLHO ESQUERDO'], heading: DocumentApp.ParagraphHeading.NORMAL, requiresPageBreak: false,
        condition: function(line) {
          const text = line.trim();
          return !text.startsWith('CID') && (/^[A-ZÀ-Ÿ0-9\s:/çÇãÃõÕáÁéÉíÍóÓúÚâÂêÊîÎôÔûÛàÀèÈìÌòÒùÙ]+:$/.test(text) || this.keywords.some(keyword => text.includes(keyword)));
        },
        style: { [DocumentApp.Attribute.FONT_FAMILY]: 'IBM Plex Serif', [DocumentApp.Attribute.FONT_SIZE]: 12, [DocumentApp.Attribute.INDENT_FIRST_LINE]: 50, [DocumentApp.Attribute.SPACING_BEFORE]: 5, }
      },
      {
        keywords: [], heading: DocumentApp.ParagraphHeading.NORMAL, requiresPageBreak: false,
        condition: function(line) { return line.startsWith("composição:"); },
        style: {
          [DocumentApp.Attribute.FONT_FAMILY]: 'IBM Plex Sans Condensed', [DocumentApp.Attribute.FONT_SIZE]: 8,
          [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.RIGHT, [DocumentApp.Attribute.INDENT_FIRST_LINE]: 0,
        }
      },
      {
        keywords: ['CID'], heading: DocumentApp.ParagraphHeading.NORMAL, requiresPageBreak: false,
        condition: function(line) { return this.keywords.some(keyword => line.trim().startsWith(keyword)) && line.trim().endsWith(":"); },
        style: { [DocumentApp.Attribute.FONT_FAMILY]: 'IBM Plex Serif', [DocumentApp.Attribute.FONT_SIZE]: 12, [DocumentApp.Attribute.INDENT_FIRST_LINE]: 50, }
      },
      {
        keywords: ["À disposição", "Auxiliar"], heading: DocumentApp.ParagraphHeading.NORMAL, requiresPageBreak: false,
        condition: function(line) { return this.keywords.some(keyword => line.trim().startsWith(keyword)); },
        style: {
          [DocumentApp.Attribute.FONT_FAMILY]: 'IBM Plex Serif', [DocumentApp.Attribute.FONT_SIZE]: 12,
          [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.RIGHT, [DocumentApp.Attribute.INDENT_FIRST_LINE]: 50,
          [DocumentApp.Attribute.LINE_SPACING]: 1.5, [DocumentApp.Attribute.SPACING_BEFORE]: 15,
        }
      },
      {
        keywords: [], heading: DocumentApp.ParagraphHeading.NORMAL, requiresPageBreak: false,
        condition: function(line) { return true; }, // Default rule
        style: {
          [DocumentApp.Attribute.FONT_FAMILY]: 'IBM Plex Serif', [DocumentApp.Attribute.FONT_SIZE]: 12,
          [DocumentApp.Attribute.INDENT_FIRST_LINE]: 50,
        }
      }
    ];
  }

  /**
   * Alias for _getRulesForWholeParagraphs. Provides compatibility for _DocumentHeadlineManager.
   * @returns {Array} The formatting rules.
   */
  function _getFormattingRules() {
    return _getRulesForWholeParagraphs();
  }


  // =================================================================
  // PRIVATE CLASSES
  // =================================================================

  /**
   * @class _TextManipulator
   * @classdesc Processes and formats raw text using a fluent, chainable interface.
   */
  class _TextManipulator {
    constructor(text, placeholders, heading1Rules, heading2Rules) {
      this.text = text;
      this.placeHolders = placeholders || [];
      this.heading1Rules = heading1Rules || [];
      this.heading2Rules = heading2Rules || [];
      this.lines = null;
      this.currentPatientName = null;
      this.processedLines = [];
      this.shouldLog = false;
    }

    withLogging(shouldLog = true) { this.shouldLog = shouldLog; return this; }
    splitInLines() { this.lines = this.text.split(/\r\n|\r|\n/g); this._log(`Split into ${this.lines.length} lines.`); return this; }
    removeEmptyLines() { if (!this.lines) this.splitInLines(); this.lines = this.lines.filter(line => line.trim() !== ""); return this; }
    trimLeadingSpaces() { if (!this.lines) this.splitInLines(); this.lines = this.lines.map(line => /^\s+\S+/.test(line) ? line.trim() : line); return this; }
    removeEscapeCharacters() { if (!this.lines) this.splitInLines(); this.lines = this.lines.map(line => line.replace(/\\/g, '').replace(/\*/g, ' ')); return this; }
    getResult(joiner = '\n') { return this.processedLines.map(item => item.text).join(joiner); }

    checkAndModifyFirstLine() {
      if (!this.lines) this.splitInLines();
      if (this.lines.length === 0) return this;
      const firstLine = this.lines[0].trim();
      const words = firstLine.split(" ").filter(word => word.length > 0);
      if (/^[A-ZÇÃÕÁÉÍÓÚÂÊÎÔÛÀÈÌÒÙÄËÏÖÜŸŠŽČĐÑ ]+$/.test(firstLine) && words.length >= 2) {
        this.lines[0] = "Nome: " + firstLine;
      }
      return this;
    }

    processLines() {
      if (!this.lines) this.splitInLines();
      this.lines.forEach(line => {
        let currentLine = this._trimSpacesAndTest(line);
        if (currentLine.trim().toUpperCase().startsWith("NOME:")) {
          const extractedName = this._extractName(currentLine);
          if (extractedName) this.currentPatientName = _capitalizeFirstLetter(extractedName);
        }
        currentLine = this._replacePlaceholders(currentLine);

        const heading1Rule = this.heading1Rules.find(rule => rule.condition(currentLine));
        if (heading1Rule) { this._handleAppliedRule(heading1Rule, currentLine); return; }

        const heading2Rule = this.heading2Rules.find(rule => rule.condition(currentLine));
        if (heading2Rule) { this._handleAppliedRule(heading2Rule, currentLine); return; }

        this.processedLines.push({ text: currentLine, heading: null });
      });
      return this;
    }

    structureH2Blocks(priorityHeadlines = null) {
      if (this.processedLines.length === 0) return this;
      const h1Groups = new Map();
      let currentH1 = null, currentH2Block = null;

      this.processedLines.forEach(item => {
        if (item.heading === DocumentApp.ParagraphHeading.HEADING1) {
          currentH1 = item;
          if (!h1Groups.has(currentH1.text)) h1Groups.set(currentH1.text, { h1: currentH1, h2Blocks: [] });
        } else if (item.heading === DocumentApp.ParagraphHeading.HEADING2) {
          if (!currentH1) this._ensureDefaultH1(h1Groups);
          currentH2Block = [item];
          h1Groups.get(currentH1.text).h2Blocks.push(currentH2Block);
        } else if (item.text.trim() !== '') {
          if (!currentH1) this._ensureDefaultH1(h1Groups);
          if (!currentH2Block) {
            const defaultH2 = { text: "Documento Indefinido", heading: DocumentApp.ParagraphHeading.HEADING2 };
            currentH2Block = [defaultH2];
            h1Groups.get(currentH1.text).h2Blocks.push(currentH2Block);
          }
          currentH2Block.push(item);
        }
      });

      if (priorityHeadlines) this._sortH2BlocksByPriority(h1Groups, priorityHeadlines);
      this._flattenStructure(h1Groups);
      return this;
    }

    _ensureDefaultH1(h1Groups) {
        currentH1 = { text: "Nome: PACIENTE INDEFINIDO", heading: DocumentApp.ParagraphHeading.HEADING1 };
        if (!h1Groups.has(currentH1.text)) h1Groups.set(currentH1.text, { h1: currentH1, h2Blocks: [] });
    }

    _sortH2BlocksByPriority(h1Groups, priorityHeadlines) {
      const priorityTexts = priorityHeadlines.map(h => h.trim().toUpperCase());
      h1Groups.forEach(group => {
        group.h2Blocks.sort((a, b) => {
          const aIndex = priorityTexts.indexOf(a[0].text.trim().toUpperCase());
          const bIndex = priorityTexts.indexOf(b[0].text.trim().toUpperCase());
          if (aIndex !== -1 && bIndex !== -1) return aIndex - bIndex;
          if (aIndex !== -1) return -1;
          if (bIndex !== -1) return 1;
          return 0;
        });
      });
    }

    _flattenStructure(h1Groups) {
      const finalStructuredLines = [];
      h1Groups.forEach(group => {
        if (group.h2Blocks.length === 0) {
          finalStructuredLines.push(group.h1);
        } else {
          group.h2Blocks.forEach(block => {
            finalStructuredLines.push(group.h1, ...block);
          });
        }
      });
      this.processedLines = finalStructuredLines;
    }

    _log(message) { if (this.shouldLog) Logger.log(message); }
    _trimSpacesAndTest(text) { return text.trim() === '' ? text : ( /^\s+\S+/.test(text) ? text.trim() : text ); }
    _extractName(h1Text) { const parts = h1Text.split(":"); return parts.length > 1 ? parts[1].trim() : null; }
    _isHeading1(rule) { return rule.heading === DocumentApp.ParagraphHeading.HEADING1; }
    _isHeading2(rule) { return rule.heading === DocumentApp.ParagraphHeading.HEADING2; }
    _escapeRegExp(string) { return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }

    _replacePlaceholders(text) {
      if (!this.placeHolders) return text;
      return this.placeHolders.reduce((accText, rule) => {
        const regex = new RegExp(this._escapeRegExp(rule.text), 'gi');
        const replacement = typeof rule.replacement === 'function' ? rule.replacement(this.currentPatientName) : rule.replacement;
        return accText.replace(regex, replacement);
      }, text);
    }

    _handleAppliedRule(rule, text) {
      if (this._isHeading1(rule)) {
        const formattedText = /^(NOME:|Nome:|nome:|nOME:)/i.test(text) ? text.replace(/^(NOME:|Nome:|nome:|nOME:)\s*/i, "Nome: ").replace(/Nome: (.+)/, (_, name) => "Nome: " + name.toUpperCase()) : text;
        this.processedLines.push({ text: formattedText, heading: DocumentApp.ParagraphHeading.HEADING1 });
      } else {
        this.processedLines.push({ text: text, heading: rule.heading });
      }
    }
  }

  /**
   * @class _ParagraphManipulator
   * @classdesc Applies formatting to paragraphs in the document body.
   */
  class _ParagraphManipulator {
    constructor(processedLines, formattingRules, insideParagraphRules) {
      this.processedLines = processedLines || [];
      this.formattingRules = formattingRules || [];
      this.insideParagraphRules = insideParagraphRules || [];
      this.shouldLog = false;
      this.body = DocumentApp.getActiveDocument().getBody();
    }

    withLogging(shouldLog = true) { this.shouldLog = shouldLog; return this; }
    clearDocument() { this.body.clear(); return this; }

    formatParagraphs() {
        if (!this.processedLines.length) return this;
        this.body.clear(); // Ensure clean slate
        const processedH1Texts = new Set();
        const h1Rule = this._getH1Rule();

        this.processedLines.forEach((item, index) => {
            if (item.text.trim() === '' && index === 0) return; // Skip leading empty lines

            const paragraph = this.body.appendParagraph(item.text);
            this._applyFormatting(paragraph, item, h1Rule, processedH1Texts);

            if (item.heading === DocumentApp.ParagraphHeading.HEADING1 && index > 0) {
                 const prevElement = paragraph.getPreviousSibling();
                 // Ensure page break is not inserted at the very top or after an empty paragraph
                 if (prevElement && prevElement.asText().getText().trim() !== '') {
                    paragraph.insertPageBreakBefore();
                 }
            }
        });
        return this;
    }

    formatTextInsideParagraphs() {
      if (!this.insideParagraphRules || this.insideParagraphRules.length === 0) return this;
      this.insideParagraphRules.forEach(rule => {
        rule.targets.forEach(target => {
          let searchResult = this.body.findText(target);
          while (searchResult) {
            const textElement = searchResult.getElement().asText();
            textElement.setAttributes(searchResult.getStartOffset(), searchResult.getEndOffsetInclusive(), { [DocumentApp.Attribute[rule.styleName.toUpperCase()]]: true });
            searchResult = this.body.findText(target, searchResult);
          }
        });
      });
      return this;
    }

    cleanupDocument() {
      const firstPara = this.body.getParagraphs()[0];
      if (firstPara && firstPara.getText().trim() === '') {
        firstPara.removeFromParent();
      }
      return this;
    }

    _log(message) { if (this.shouldLog) Logger.log(message); }
    _getH1Rule() { return this.formattingRules.find(rule => rule.heading === DocumentApp.ParagraphHeading.HEADING1 && rule.matchType !== 'default'); }

    _applyFormatting(paragraph, item, h1Rule, processedH1Texts) {
        let appliedStyle = {};
        const rule = this.formattingRules.find(r => r.condition(item.text));
        if (rule) appliedStyle = rule.style || {};

        if (item.heading) {
            if (item.heading === DocumentApp.ParagraphHeading.HEADING1) {
                if (processedH1Texts.has(item.text)) {
                    paragraph.setHeading(DocumentApp.ParagraphHeading.NORMAL);
                } else {
                    paragraph.setHeading(DocumentApp.ParagraphHeading.HEADING1);
                    processedH1Texts.add(item.text);
                }
            } else {
                paragraph.setHeading(item.heading);
            }
        }
        paragraph.setAttributes(appliedStyle);
    }
  }

  /**
   * @class _DocumentHeadlineManager
   * @classdesc Manages the structure and content of the document based on H1/H2 headings.
   */
  class _DocumentHeadlineManager {
    constructor() {
      const doc = DocumentApp.getActiveDocument();
      this.bodyText = doc.getBody().getText();
      this.selectionText = (doc.getSelection() && doc.getSelection().getRangeElements().map(e => e.getElement().asText().getText().substring(e.getStartOffset(), e.getEndOffsetInclusive() + 1)).join('')) || '';

      const rules = _getFormattingRules();
      const h1Conditions = rules.filter(r => r.heading === DocumentApp.ParagraphHeading.HEADING1);
      const h2Conditions = rules.filter(r => r.heading === DocumentApp.ParagraphHeading.HEADING2);

      this.bodyStructure = this._buildDocumentStructure(this.bodyText, h1Conditions, h2Conditions);
      this.selectionStructure = this.selectionText ? this._buildDocumentStructure(this.selectionText, h1Conditions, h2Conditions, true) : null;
    }

    _buildDocumentStructure(text, h1Conditions, h2Conditions, limitToOneH2 = false) {
      const lines = text.split('\n');
      let currentH1 = null;
      let docStructure = [];
      let foundFirstH2 = false;

      lines.forEach(line => {
        if (h1Conditions.some(rule => rule.condition(line))) {
          currentH1 = line;
        } else if (h2Conditions.some(rule => rule.condition(line))) {
          if (limitToOneH2 && foundFirstH2) return;
          docStructure.push({ headline1: currentH1, headline2: line, text: '' });
          foundFirstH2 = true;
        } else if (docStructure.length > 0) {
          docStructure[docStructure.length - 1].text += line + '\n';
        }
      });
      return docStructure;
    }

    _sortAndGroupHeadlines() {
      const headline1Map = new Map();
      this.bodyStructure.forEach(item => {
        if (!headline1Map.has(item.headline1)) headline1Map.set(item.headline1, []);
        headline1Map.get(item.headline1).push(item);
      });

      const grouped = [];
      headline1Map.forEach(items => {
        items.sort((a, b) => {
          if (a.headline2 === "Prontuário Médico") return -1;
          if (b.headline2 === "Prontuário Médico") return 1;
          return 0;
        });
        grouped.push(...items);
      });
      this.bodyStructure = grouped;
    }

    getText() {
      return this.bodyStructure
        .map(item => `${item.headline1}\n${item.headline2}\n${(item.text || '').trim()}`)
        .join("\n\n");
    }

    createOrUpdateBodyHeadline2(newText, headline1, newHeadline2, fromHeadline2 = null) {
      let cleanedText = newText;
      if (cleanedText.trim().startsWith(newHeadline2)) {
        cleanedText = cleanedText.substring(cleanedText.indexOf(newHeadline2) + newHeadline2.length).trim();
      }

      const isUpdate = fromHeadline2 && fromHeadline2 === newHeadline2;
      const isSpecialCase = fromHeadline2 === "Prontuário Médico" && newHeadline2 === "Prescrição de Óculos";

      if (isUpdate) {
        this.bodyStructure = this.bodyStructure.filter(doc => !(doc.headline1 === headline1 && doc.headline2 === fromHeadline2));
      } else if (isSpecialCase) {
        this.bodyStructure = this.bodyStructure.filter(doc => !(doc.headline1 === headline1 && doc.headline2 === "Prescrição de Óculos"));
      }

      const exactMatch = this.bodyStructure.find(doc => doc.headline1 === headline1 && doc.headline2 === newHeadline2 && doc.text.trim() === cleanedText);
      if (!exactMatch) {
        this.bodyStructure.push({ headline1, headline2: newHeadline2, text: cleanedText + '\n' });
      }
    }

    // Public-facing methods for the manager
    getUniqueHeadline1s() { return [...new Set(this.bodyStructure.map(item => item.headline1).filter(Boolean))]; }
    getHeadline2sForHeadline1(h1) { return [...new Set(this.bodyStructure.filter(item => item.headline1 === h1).map(item => item.headline2).filter(Boolean))]; }
    getContentForPair(h1, h2) { const item = this.bodyStructure.find(doc => doc.headline1 === h1 && doc.headline2 === h2); return item ? item.text : null; }
    getDocumentHeadlinePairsAndContent() { return this.bodyStructure.map(item => ({ headline1: item.headline1, headline2: item.headline2, content: (item.text || '').trim() })); }
    getAugmentedContentForH1(h1) {
        const sections = this.bodyStructure.filter(item => item.headline1 === h1);
        if (sections.length === 0) return null;
        return sections.map(s => `${s.headline2}\n${(s.text || '').trim()}`).join('\n\n');
    }
    hasHeadline1AndHeadline2InSelection(h2) { return this.selectionStructure && this.selectionStructure.some(doc => doc.headline2 === h2); }
    getHeadline1fromSelection() { return this.selectionStructure && this.selectionStructure.length > 0 ? this.selectionStructure[0].headline1 : null; }
    getFirstTextfromBody(h2) { const item = this.bodyStructure.find(doc => doc.headline2 === h2); return item ? item.text : null; }
  }


  // =================================================================
  // PUBLIC API
  // =================================================================
  const publicApi = {};

  /**
   * Fully formats the active Google Document based on predefined rules.
   * This is the primary entry point for the auto-formatting feature.
   */
  publicApi.executeFormat = () => {
    const doc = DocumentApp.getActiveDocument();
    const bodyText = doc.getBody().getText();
    const manager = new _DocumentHeadlineManager(); // To get patient name
    const patientName = manager.getUniqueHeadline1s()[0]?.split(':')[1]?.trim() || null;

    const placeholders = _getRulesForPlaceholders();
    const wholeParaRules = _getRulesForWholeParagraphs();
    const h1Rules = _filterRulesByHeading(wholeParaRules, DocumentApp.ParagraphHeading.HEADING1);
    const h2Rules = _filterRulesByHeading(wholeParaRules, DocumentApp.ParagraphHeading.HEADING2);
    const insideParaRules = _getRulesForInsideParagraphs();
    const priorityH2s = ["Prontuário Médico", "Laudo de Mapeamento de Retina"];

    const textManipulator = new _TextManipulator(bodyText, placeholders, h1Rules, h2Rules);
    textManipulator.currentPatientName = patientName ? _capitalizeFirstLetter(patientName) : null;

    const processedLines = textManipulator
      .splitInLines()
      .removeEmptyLines()
      .trimLeadingSpaces()
      .removeEscapeCharacters()
      .checkAndModifyFirstLine()
      .processLines()
      .structureH2Blocks(priorityH2s)
      .processedLines;

    new _ParagraphManipulator(processedLines, wholeParaRules, insideParaRules)
      .formatParagraphs()
      .formatTextInsideParagraphs()
      .cleanupDocument();
  };

  /**
   * Creates or updates a section (H2) in the document for a given patient (H1).
   * @param {string} newText - The new content for the section.
   * @param {string} headline1 - The patient's H1.
   * @param {string} newHeadline2 - The target H2.
   * @param {string} [fromHeadline2=null] - The source H2, if applicable.
   */
  publicApi.processAndUpdateDocument = (newText, headline1, newHeadline2, fromHeadline2 = null) => {
    const manager = new _DocumentHeadlineManager();
    manager.createOrUpdateBodyHeadline2(newText, headline1, newHeadline2, fromHeadline2);
    manager._sortAndGroupHeadlines();
    const newBodyText = manager.getText();
    DocumentApp.getActiveDocument().getBody().setText(newBodyText);
    publicApi.executeFormat(); // Re-format to apply styles correctly
  };

  // Exposing headline manager functionalities through the service
  publicApi.getUniqueHeadline1s = () => new _DocumentHeadlineManager().getUniqueHeadline1s();
  publicApi.getHeadline2sForHeadline1 = (h1) => new _DocumentHeadlineManager().getHeadline2sForHeadline1(h1);
  publicApi.getDocumentHeadlinePairsAndContent = () => new _DocumentHeadlineManager().getDocumentHeadlinePairsAndContent();
  publicApi.getContentForPair = (h1, h2) => new _DocumentHeadlineManager().getContentForPair(h1, h2);
  publicApi.getAugmentedContentForH1 = (h1) => new _DocumentHeadlineManager().getAugmentedContentForH1(h1);
  publicApi.hasHeadline1AndHeadline2InSelection = (h2) => new _DocumentHeadlineManager().hasHeadline1AndHeadline2InSelection(h2);
  publicApi.getHeadline1fromSelection = () => new _DocumentHeadlineManager().getHeadline1fromSelection();
  publicApi.getFirstTextfromBody = (h2) => new _DocumentHeadlineManager().getFirstTextfromBody(h2);

  /**
   * Exposes the internal manager class for stateful batch operations.
   * @type {_DocumentHeadlineManager}
   */
  publicApi.Manager = _DocumentHeadlineManager;

  return publicApi;

})();
