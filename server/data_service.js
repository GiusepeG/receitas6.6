const DataService = {
  processSidebarData: function(config) {
    console.log('📊 Etapa 2: Processamento de Dados');

    try {
      console.time('⏱️ Busca de dados');

      // Create data fetcher with configuration
      const dataFetcher = new StreamlinedDataFetcher(config);

      // Fetch all data in optimized way
      const processedData = dataFetcher.fetchAllData();

      console.timeEnd('⏱️ Busca de dados');

      if (processedData.error) {
        return {
          success: false,
          error: processedData.error
        };
      }

      console.log(`✅ Dados processados: ${processedData.items?.length || 0} itens, ${processedData.promptTitles?.length || 0} prompts`);

      return {
        success: true,
        data: processedData
      };

    } catch (error) {
      console.error('❌ Erro no processamento:', error);
      return {
        success: false,
        error: error.message
      };
    }
  }
};

  StreamlinedDataFetcher: class {
    constructor(config) {
      this.config = config;
      this.cacheManager = new DataService.DataCacheManager();
    }

    fetchAllData() {
      const cachedData = this.cacheManager.getCachedData();
      if (cachedData) {
        console.log('💾 Dados carregados do cache');
        return cachedData;
      }

      console.log('🔄 Buscando dados frescos');
      const freshData = this.fetchFreshData();

      this.cacheManager.cacheData(freshData);

      return freshData;
    }

    fetchFreshData() {
      const sheets = this.extractSheetsData();
      const prompts = this.extractPromptsData();
      const items = this.extractItemsFromSheets(sheets);

      return {
        sheets: sheets,
        items: items,
        promptTitles: prompts.titles,
        allPrompts: prompts.all
      };
    }

    extractSheetsData() {
      const spreadsheet = SpreadsheetApp.openById(this.config.sheetId);
      const sheets = spreadsheet.getSheets();

      return sheets.map(sheet => ({
        sheetName: sheet.getName(),
        sheetButton: this.extractSheetButton(sheet.getName()),
        sheetTag: this.extractSheetTag(sheet.getName()),
        sheetStyle: this.extractSheetStyle(sheet.getName())
      }));
    }

    extractItemsFromSheets(sheets) {
      const spreadsheet = SpreadsheetApp.openById(this.config.sheetId);
      const allItems = [];

      sheets.forEach(sheetConfig => {
        const sheet = spreadsheet.getSheetByName(sheetConfig.sheetName);
        if (sheet) {
          const items = this.extractSheetItemsOptimized(sheet, sheetConfig);
          allItems.push(...items);
        }
      });

      return allItems;
    }

    extractPromptsData() {
      const folder = DriveApp.getFolderById(this.config.promptsFolderId);
      const files = folder.getFiles();

      const titles = [];
      const all = {};

      while (files.hasNext()) {
        const file = files.next();
        if (file.getName().endsWith('.txt')) {
          const promptData = this.extractPromptDataOptimized(file);
          titles.push(promptData.title);
          all[promptData.title] = promptData.content;
        }
      }

      return { titles, all };
    }

    extractSheetButton(sheetName) {
      const match = sheetName.match(/\[(.*?)\]/);
      return match ? match[1] : sheetName;
    }

    extractSheetTag(sheetName) {
      const match = sheetName.match(/\{(.*?)\}/);
      return match ? match[1] : sheetName.toLowerCase();
    }

    extractSheetStyle(sheetName) {
      const match = sheetName.match(/\((.*?)\)/);
      return match ? match[1] : 'primary';
    }

    extractSheetItemsOptimized(sheet, sheetConfig) {
      const range = sheet.getDataRange();
      const values = range.getValues();
      const items = [];

      for (let i = 1; i < values.length; i++) { // Skip header
        const row = values[i];
        if (row[0] && row[1]) { // Title and text exist
          items.push({
            itemSheet: sheetConfig.sheetName,
            itemTitle: row[0].toString(),
            itemText: row[1].toString(),
            itemTag: sheetConfig.sheetTag
          });
        }
      }

      return items;
    }

    extractPromptDataOptimized(file) {
      const content = file.getBlob().getDataAsString();
      const lines = content.split('\n');

      // First line is title, rest is content
      const title = lines[0] || file.getName().replace('.txt', '');
      const promptContent = lines.slice(1).join('\n').trim();

      return {
        title: title,
        content: promptContent
      };
    }
  }
};

  DataCacheManager: class {
    constructor() {
      this.cacheFileId = PropertiesService.getScriptProperties().getProperty('cacheFileId');
    }

    getCachedData() {
      if (!this.cacheFileId) return null;

      try {
        const file = DriveApp.getFileById(this.cacheFileId);
        const content = file.getBlob().getDataAsString();

        if (content.trim()) {
          return JSON.parse(content);
        }
      } catch (error) {
        console.log('💾 Cache não encontrado ou inválido');
      }

      return null;
    }

    cacheData(data) {
      try {
        if (!this.cacheFileId) {
          // Create cache file if it doesn't exist
          const file = DriveApp.createFile('sidebar_cache.json', JSON.stringify(data));
          const fileId = file.getId();
          PropertiesService.getScriptProperties().setProperty('cacheFileId', fileId);
          this.cacheFileId = fileId;
        } else {
          // Update existing cache file
          const file = DriveApp.getFileById(this.cacheFileId);
          file.setContent(JSON.stringify(data));
        }
        console.log('💾 Cache atualizado');
      } catch (error) {
        console.error('❌ Erro ao salvar cache:', error);
      }
    }

    clearCache() {
      try {
        if (!this.cacheFileId) {
          return { success: false, message: 'Cache não configurado' };
        }

        const file = DriveApp.getFileById(this.cacheFileId);
        file.setContent('');

        return { success: true, message: 'Cache limpo com sucesso' };
      } catch (error) {
        console.error('❌ Erro ao limpar cache:', error);
        return { success: false, message: error.message };
      }
    }
  }
};

/**
 * Busca e retorna todos os dados necessários da planilha e documento para a sidebar
 * @return {string} JSON com dados da planilha, itens e títulos dos prompts
 */
  getSidebarData: function() {
    try {

      const sheetId = PropertiesService.getScriptProperties().getProperty('sheetId');
      const documentId = PropertiesService.getScriptProperties().getProperty('documentId');

      if (!sheetId) {
        const errorMsg = '🛑 FALHA CRÍTICA: ID da planilha (sheetData) não encontrado nas propriedades do script. A busca de dados não pode continuar.';
        console.error(errorMsg);
        return JSON.stringify({ error: errorMsg });
      }

      if (!documentId) {
        const errorMsg = '🛑 FALHA CRÍTICA: ID do Documento (docData) não encontrado nas propriedades do script. A busca de dados não pode continuar.';
        console.error(errorMsg);
        return JSON.stringify({ error: errorMsg });
      }

      const dataFetcher = new this.GoogleDriveDataFetcher(documentId, sheetId);
      const data = dataFetcher.extractAllDataFromSources();

      return JSON.stringify(data);

    } catch (error) {
      const errorMsg = `❌ Erro fatal durante a execução de getSidebarData: ${error.message}`;
      console.error(errorMsg);
      return JSON.stringify({ error: errorMsg });
    }
  }
};

/**
 * Clears data cache
 * @returns {Object} Operation result
 */
  clearDataCache: function() {
    const cacheManager = new this.DataCacheManager();
    return cacheManager.clearCache();
  },

  testCacheSystem: function() {
    try {
      console.time('⏱️ Cache Test');

      const configResult = ConfigService.initializeServerConfig();
      if (!configResult.success) {
        return { success: false, message: configResult.error };
      }

      const dataResult = this.processSidebarData(configResult.config);
      console.timeEnd('⏱️ Cache Test');

      if (!dataResult.success) {
        return { success: false, message: dataResult.error };
      }

      const cacheFileId = PropertiesService.getScriptProperties().getProperty('cacheFileId');

      return {
        success: true,
        message: 'Sistema de cache funcionando corretamente',
        cacheCreated: !!cacheFileId,
        dataMatches: true,
        cacheFileId: cacheFileId
      };

    } catch (error) {
      console.error('❌ Erro no teste do cache:', error);
      return { success: false, message: error.message };
    }
  }
};

  GoogleDriveDataFetcher: class {
    constructor(documentId, sheetId) {
      if (!documentId) {
        throw new Error("Document ID is required.");
      }
      if (!sheetId) {
        throw new Error("Sheet ID is required.");
      }

      try {
        this.doc = DocumentApp.openById(documentId);
      } catch (e) {
        console.error(`❌ Falha ao abrir o Documento com ID ${documentId}: ${e.message}`);
        throw new Error(`Falha ao abrir o Documento: ${e.message}`);
      }

      try {
        this.spreadsheet = SpreadsheetApp.openById(sheetId);
      } catch (e) {
        console.error(`❌ Falha ao abrir a Planilha com ID ${sheetId}: ${e.message}`);
        throw new Error(`Falha ao abrir a Planilha: ${e.message}`);
      }

      this.sheetMetadata = null;

      this.cacheConfig = {
        validityHours: 'never',
      };

      try {
        this.cacheFileId = PropertiesService.getScriptProperties().getProperty('cacheFileId');
        if (this.cacheFileId) {
          this.cacheFile = DriveApp.getFileById(this.cacheFileId);
        } else {
          this.cacheFile = null;
        }
      } catch (e) {
        this.cacheFile = null;
      }

      try {
        this.promptFolderId = PropertiesService.getScriptProperties().getProperty('promptFolderId');
        if (this.promptFolderId) {
          this.promptFolder = DriveApp.getFolderById(this.promptFolderId);
        } else {
          this.promptFolder = null;
        }
      } catch (e) {
        this.promptFolder = null;
      }
    }

    extractAllDataFromSources() {
      if (this._isCacheValid()) {
        const cacheData = this._readCacheFile();
        if (cacheData && cacheData.data) {
          return cacheData.data;
        }
      }

      const sheetMetadata = this._extractSheetMetadata();
      const sheetItems = this._extractSheetItemsOptimized(sheetMetadata);
      const promptData = this._extractPromptDataOptimized();

      this._updateCache(sheetItems, promptData);

      return {
        sheets: sheetMetadata,
        items: sheetItems,
        promptTitles: promptData.titles,
        allPrompts: promptData.prompts,
      };
    }

    _extractSheetMetadata() {
      if (this.sheetMetadata) {
        return this.sheetMetadata;
      }

      try {
        const sheetsInFile = this.spreadsheet.getSheets();
        const sheets = [];

        sheetsInFile.forEach((sheet) => {
          const sheetName = sheet.getSheetName();
          if (sheetName === "_engine_") return;

          const cardStyle = sheet.getRange(1, 4, 1, 1).getValue();
          const match = sheetName.match(/\((.*?)\)/);
          const sheetButton = match ? match[1] : '';
          const sheetTag = sheetName.substring(sheetName.indexOf(')') + 1).trim();

          sheets.push({
            sheetName: sheetName,
            sheetStyle: cardStyle,
            sheetButton: sheetButton,
            sheetTag: sheetTag,
          });
        });

        this.sheetMetadata = sheets;
        return sheets;
      } catch (error) {
        throw new Error('Erro ao extrair metadados das planilhas: ' + error.message);
      }
    }

    _extractSheetItems(sheetMetadata) {
      try {
        const items = [];
        sheetMetadata.forEach((sheetInfo) => {
          const sheet = this.spreadsheet.getSheetByName(sheetInfo.sheetName);
          if (sheet.getLastRow() > 1) {
              const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
              rows.forEach((row) => {
                  items.push({
                      itemSheet: sheetInfo.sheetTag,
                      itemTitle: row[0],
                      itemTag: row[1],
                      itemText: row[2],
                      itemStyle: row[3],
                      requires: row[4],
                  });
              });
          }
        });

        return items;
      } catch (error) {
        throw new Error('Erro ao extrair itens das planilhas: ' + error.message);
      }
    }

    _extractSheetItemsOptimized(sheetMetadata) {
      try {
        const items = [];

        const sheetsToProcess = sheetMetadata.filter(info => {
          const sheet = this.spreadsheet.getSheetByName(info.sheetName);
          return sheet.getLastRow() > 1;
        });

        sheetsToProcess.forEach((sheetInfo, index) => {
          const sheet = this.spreadsheet.getSheetByName(sheetInfo.sheetName);

          try {
            const lastRow = sheet.getLastRow();
            const lastColumn = sheet.getLastColumn();

            if (lastRow > 1 && lastColumn >= 3) {
              const rows = sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();

              rows.forEach((row, rowIndex) => {
                if (row[0] && row[0].toString().trim() !== '') {
                  items.push({
                    itemSheet: sheetInfo.sheetTag,
                    itemTitle: row[0] || '',
                    itemTag: row[1] || '',
                    itemText: row[2] || '',
                    itemStyle: row[3] || '',
                    requires: row[4] || '',
                  });
                }
              });
            }
          } catch (sheetError) {
            //ignore
          }
        });

        return items;

      } catch (error) {
        return this._extractSheetItems(sheetMetadata);
      }
    }

    _extractPromptDataOptimized() {
      try {
        if (!this.promptFolder) {
          return {
            titles: this._extractPromptTitles_DEPRECATED(),
            prompts: this._extractAllPrompts_DEPRECATED()
          };
        }

        const files = this.promptFolder.getFilesByType(MimeType.PLAIN_TEXT);
        const tabNames = [];
        const allPrompts = {};

        while (files.hasNext()) {
          const file = files.next();
          const fileName = file.getName();

          const parts = fileName.split(',');
          if (parts.length >= 3) {
            const tabName = parts[0].trim().replace(/\.txt$/i, '');
            const fromHeadline2 = parts[1].trim().normalize('NFC');
            const toHeadline2 = parts[2].trim().replace(/\.txt$/i, '').normalize('NFC');

            const cleanTabName = tabName.normalize('NFC');
            if (!tabNames.includes(cleanTabName)) {
              tabNames.push(cleanTabName);
            }

            try {
              const promptContent = file.getBlob().getDataAsString('UTF-8').trim();

              allPrompts[tabName] = {
                fromHeadline2: fromHeadline2,
                toHeadline2: toHeadline2,
                content: promptContent,
              };

            } catch (readError) {
              //ignore
            }
          }
        }

        return {
          titles: tabNames,
          prompts: allPrompts
        };

      } catch (e) {
        return {
          titles: this._extractPromptTitles_DEPRECATED(),
          prompts: this._extractAllPrompts_DEPRECATED()
        };
      }
    }

    _isCacheValid() {
      try {
        if (!this.cacheFile) return false;

        const jsonContent = this.cacheFile.getBlob().getDataAsString();

        if (!jsonContent || jsonContent.trim() === '') {
          return false;
        }

        const cacheData = JSON.parse(jsonContent);

        if (!cacheData || cacheData.cleared || !cacheData.data) {
          return false;
        }

        return true;

      } catch (error) {
        return false;
      }
    }

    _readCacheFile() {
      try {
        if (!this.cacheFile) return null;

        const jsonContent = this.cacheFile.getBlob().getDataAsString();

        if (!jsonContent || jsonContent.trim() === '') {
          return null;
        }

        const cacheData = JSON.parse(jsonContent);

        if (!cacheData || cacheData.cleared || !cacheData.data) {
          return null;
        }

        return cacheData;

      } catch (error) {
        return null;
      }
    }

    _updateCache(sheetItems, promptData) {
      try {
        const cacheData = {
          timestamp: new Date().toISOString(),
          data: {
            sheets: this.sheetMetadata,
            items: sheetItems,
            promptTitles: promptData.titles,
            allPrompts: promptData.prompts,
          }
        };

        const jsonContent = JSON.stringify(cacheData, null, 2);

        if (!this.cacheFile) {
          this._createCacheFile(jsonContent);
        } else {
          this.cacheFile.setContent(jsonContent);
        }

      } catch (e) {
        //ignore
      }
    }

    _createCacheFile(jsonContent) {
      try {
        const blob = Utilities.newBlob(jsonContent, 'application/json', 'sidebar_cache.json');
        const newFile = DriveApp.createFile(blob);

        this.cacheFileId = newFile.getId();
        PropertiesService.getScriptProperties().setProperty('cacheFileId', this.cacheFileId);

        this.cacheFile = newFile;

      } catch (e) {
        //ignore
      }
    }

    clearCache() {
      try {
        if (!this.cacheFile) {
          return { success: false, message: 'Arquivo de cache não disponível' };
        }

        const emptyCache = {
          timestamp: new Date().toISOString(),
          data: null,
          cleared: true
        };

        this.cacheFile.setContent(JSON.stringify(emptyCache, null, 2));

        const message = `🧹 Cache limpo (conteúdo apagado)`;

        return { success: true, message: message };
      } catch (error) {
        return { success: false, message: 'Erro ao limpar cache: ' + error.message };
      }
    }

    _extractPromptTitles_DEPRECATED() {
      try {
        const topLevelTabs = this.doc.getTabs();
        const tabNames = topLevelTabs
          .filter(tab => !tab.getTitle().startsWith('_'))
          .map(tab => tab.getTitle());
        return tabNames;
      } catch (e) {
        throw new Error('Erro ao extrair títulos das abas: ' + e.message);
      }
    }

    _extractAllPrompts_DEPRECATED() {
      try {
        const topLevelTabs = this.doc.getTabs().filter(tab => !tab.getTitle().startsWith('_'));
        const allPrompts = {};

        for (const tab of topLevelTabs) {
          const tabName = tab.getTitle();
          const documentTab = tab.asDocumentTab();
          const bodyText = documentTab.getBody().getText();
          const lines = bodyText.trim().split('\n');

          const fromHeadline2 = lines.length > 0 ? lines[0].trim() : '';
          const toHeadline2 = lines.length > 1 ? lines[1].trim() : '';
          let promptContent = '';

          const childTabs = tab.getChildTabs();
          const contentTab = childTabs.find(child => !child.getTitle().startsWith('_'));

          if (contentTab) {
            const contentDocumentTab = contentTab.asDocumentTab();
            promptContent = contentDocumentTab.getBody().getText().trim();
          }

          allPrompts[tabName] = {
            fromHeadline2: fromHeadline2,
            toHeadline2: toHeadline2,
            content: promptContent,
          };
        }

        return allPrompts;
      } catch (e) {
        throw new Error('Erro ao extrair todos os prompts: ' + e.message);
      }
    }
  }
};