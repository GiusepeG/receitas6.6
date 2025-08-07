const ConfigService = {
  initializeServerConfig: function() {
    console.log('📋 Etapa 1: Configuração do Servidor');

    try {
      // Validate existing properties
      if (!this.validateScriptProperties()) {
        console.log('⚙️ Configurando propriedades...');
        this.setupScriptProperties();

        if (!this.validateScriptProperties()) {
          return {
            success: false,
            error: 'Não foi possível configurar as Propriedades do Script automaticamente'
          };
        }
      }

      // Get and validate all required IDs
      const config = {
        sheetId: this.getSheetId(),
        documentId: this.getDocumentId(),
        promptsFolderId: this.getPromptsFolderId()
      };

      // Validate all IDs are present
      if (!config.sheetId || !config.documentId || !config.promptsFolderId) {
        return {
          success: false,
          error: 'IDs de configuração não encontrados'
        };
      }

      console.log('✅ Configuração validada');
      return {
        success: true,
        config: config
      };

    } catch (error) {
      console.error('❌ Erro na configuração:', error);
      return {
        success: false,
        error: error.message
      };
    }
  }
};

  getSheetId: function() {
    return PropertiesService.getScriptProperties().getProperty('sheetId');
  },

  getDocumentId: function() {
    return PropertiesService.getScriptProperties().getProperty('documentId');
  },

  getPromptsFolderId: function() {
    return PropertiesService.getScriptProperties().getProperty('promptsFolderId');
  }
};

  validateScriptProperties: function() {
    const properties = PropertiesService.getScriptProperties();
    const requiredProps = ['sheetId', 'documentId', 'promptsFolderId'];

    return requiredProps.every(prop => {
      const value = properties.getProperty(prop);
      return value && value.trim() !== '';
    });
  },

  setupScriptProperties: function() {
    const properties = PropertiesService.getScriptProperties();

    // Set default values if not exist
    if (!properties.getProperty('sheetId')) {
      properties.setProperty('sheetId', 'YOUR_SHEET_ID_HERE');
    }

    if (!properties.getProperty('documentId')) {
      properties.setProperty('documentId', 'YOUR_DOCUMENT_ID_HERE');
    }

    if (!properties.getProperty('promptsFolderId')) {
      properties.setProperty('promptsFolderId', 'YOUR_PROMPTS_FOLDER_ID_HERE');
    }
  }
};