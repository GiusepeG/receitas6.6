function onOpen(e) {
  DocumentApp.getUi()
    .createMenu('Custom Menu')
    .addItem('Abrir Sidebar', 'openSidebarServerSide')
    .addItem('Cabeçalhos e Rodapés', 'chooseHeaderFooter')
    .addSeparator()
    .addItem('Extrair `Nome:` da agenda', 'extractHeadlineOneFromText')
    .addSeparator()
    .addItem('Configurar Propriedades', 'openPropertiesModal')
    .addItem('Limpar Cache', 'clearDataCacheFromMenu')
    .addToUi();
}

/**
 * Função para limpar cache via menu
 */
function clearDataCacheFromMenu() {
  try {
    const result = DataService.clearDataCache();

    if (result.success) {
      DocumentApp.getUi().alert('✅ Cache limpo com sucesso!\n\n' + result.message);
    } else {
      DocumentApp.getUi().alert('❌ Erro ao limpar cache:\n\n' + result.message);
    }
  } catch (error) {
    DocumentApp.getUi().alert('❌ Erro inesperado:\n\n' + error.message);
  }
}

function openSidebarServerSide() {
  const htmlFileName = 'client/sidebar_view.html';
  const title = 'Sidebar';

  console.log('📋 Etapa 1: Validação de Propriedades');
  if (!ConfigService.validateScriptProperties()) {
    ConfigService.setupScriptProperties();

    if (!ConfigService.validateScriptProperties()) {
      DocumentApp.getUi().alert('❌ Erro: Não foi possível configurar as Propriedades do Script manualmente. Execute a função setupScriptProperties() manualmente.');
      return;
    }
  }
  console.log('✅ Propriedades validadas');

  console.log('📊 Etapa 2: Busca de Dados');
  let sidebarData;
  try {
    console.time('⏱️ Busca de dados');

    const jsonResponse = DataService.getSidebarData();
    sidebarData = JSON.parse(jsonResponse);

    console.timeEnd('⏱️ Busca de dados');

    if (sidebarData.error) {
      DocumentApp.getUi().alert('❌ Erro ao buscar dados: ' + sidebarData.error);
      return;
    }

    console.log(`✅ Dados obtidos: ${sidebarData.items?.length || 0} itens, ${sidebarData.promptTitles?.length || 0} prompts`);

  } catch (error) {
    console.error('❌ Erro ao buscar dados da sidebar:', error);
    DocumentApp.getUi().alert('❌ Erro ao buscar dados: ' + error.message);
    return;
  }

  console.log('🎨 Etapa 3: Criação da Sidebar');
  const html = HtmlService.createTemplateFromFile(htmlFileName);
  html.initialSearchInput = "Anamnese";
  html.preloadedData = JSON.stringify(sidebarData);

  const sidebar = html.evaluate()
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setTitle(title)
    .setWidth(300);

  console.log('✅ Sidebar criada');

  console.log('📱 Exibindo Sidebar');
  DocumentApp.getUi().showSidebar(sidebar);
  console.log('✅ Sidebar ativa');
}

function createOptimizedSidebar(data) {
  console.log('🎨 Etapa 3: Criação da Sidebar');

  try {
    const html = HtmlService.createTemplateFromFile('client/sidebar_view.html');
    html.initialSearchInput = "Anamnese";
    html.preloadedData = JSON.stringify(data);

    const sidebar = html.evaluate()
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setTitle('Sidebar')
      .setWidth(300);

    DocumentApp.getUi().showSidebar(sidebar);

    console.log('✅ Sidebar ativa');

    return { success: true };

  } catch (error) {
    console.error('❌ Erro na criação da sidebar:', error);
    return {
      success: false,
      error: error.message
    };
  }
}