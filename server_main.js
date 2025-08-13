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
    const result = DataService.clearCache();

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
  const htmlFileName = 'client_sidebar_view';
  const title = 'Sidebar';

  // Etapa 1: Validação de Propriedades
  if (!validateScriptProperties()) {
    setupScriptProperties();
    if (!validateScriptProperties()) {
      DocumentApp.getUi().alert('❌ Erro: Não foi possível configurar as Propriedades do Script automaticamente.');
      return;
    }
  }

  // Etapa 2: Busca de Dados
  let sidebarData;
  try {
    const jsonResponse = getSidebarData(); // This now calls DataService.getSidebarData
    sidebarData = JSON.parse(jsonResponse);

    if (sidebarData.error) {
      DocumentApp.getUi().alert('❌ Erro ao buscar dados: ' + sidebarData.error);
      return;
    }
  } catch (error) {
    DocumentApp.getUi().alert('❌ Erro ao buscar dados: ' + error.message);
    return;
  }

  // Etapa 3: Criação e Exibição da Sidebar
  const html = HtmlService.createTemplateFromFile(htmlFileName);
  html.initialSearchInput = "Anamnese";
  html.preloadedData = JSON.stringify(sidebarData);

  const sidebar = html.evaluate()
    .setSandboxMode(HtmlService.SandboxMode.IFRAME) // IFRAME is recommended for security
    .setTitle(title)
    .setWidth(300);

  DocumentApp.getUi().showSidebar(sidebar);
}

function includeServerSide(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}