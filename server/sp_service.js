const SPService = {
  openPropertiesModal: function() {
    if (!ConfigService.validateScriptProperties()) {
        ConfigService.setupScriptProperties();

        if (!ConfigService.validateScriptProperties()) {
            DocumentApp.getUi().alert('❌ Erro: Não foi possível configurar as Propriedades do Script automaticamente. Execute a função setupScriptProperties() manualmente.');
            return;
        }
    }

    const scriptProperties = PropertiesService.getScriptProperties();

    const currentSecretary13 = scriptProperties.getProperty('secretary13') || '';
    const currentSecretary12 = scriptProperties.getProperty('secretary12') || '';

    const template = HtmlService.createTemplateFromFile('dialogs/sp_client_dialog.html');

    template.currentSecretary13 = currentSecretary13;
    template.currentSecretary12 = currentSecretary12;

    const html = template.evaluate()
        .setWidth(600)
        .setHeight(300);
    DocumentApp.getUi().showModalDialog(html, 'Configurar Propriedades');
  }
};

  savePropertiesData: function(formData) {
    const scriptProperties = PropertiesService.getScriptProperties();

    try {
      scriptProperties.setProperties({
        'secretary13': formData.secretary13 || '',
        'secretary12': formData.secretary12 || ''
      });

    } catch (e) {
      throw new Error('Não foi possível salvar as configurações. Erro: ' + e.message);
    }
  }
};