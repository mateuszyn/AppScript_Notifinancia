// =================== SERVIDOR WEB ===================

/**
 * Lida com as requisições GET. Serve a página principal do Web App.
 * @param {Object} e O objeto de evento da requisição.
 * @returns {HtmlOutput} A página HTML renderizada.
 */

function doGet(e) {
  // Captura o e-mail do usuário logado
  const userEmail = Session.getActiveUser().getEmail() || "Visitante";
  
  // LOGS DE DEPURAÇÃO: Aparecerão no menu 'Execuções' do Apps Script
  console.log("--- Iniciando doGet Multi-Usuário ---");
  console.log("Usuário detectado: " + userEmail);

  // 1. REGISTRO DE ANALYTICS: Log de acesso ao sistema
  registrarLog("ACESSO", "Usuário abriu o Dashboard");

  // 2. GESTÃO DE USUÁRIO: Garante que o usuário tenha sua própria aba
  garantirUsuario(userEmail);

  const htmlTemplate = HtmlService.createTemplateFromFile('WebApp');
  
  // 3. BUSCA DE DADOS: Agora usamos 'userEmail' diretamente como o nome da aba
  // Garante que o carregamento inicial use a ordenação da célula E2 do usuário
  htmlTemplate.ativos = buscarDadosAtivos(null, null, userEmail);

  // Carrega configurações globais e do usuário
  htmlTemplate.allTickers = getAllTickers(); 
  
  // Abre a planilha para buscar a corretora salva na célula E1 da aba do usuário
  const planilha = SpreadsheetApp.openById(ID_PLANILHA);
  const abaUsuario = planilha.getSheetByName(userEmail);
  htmlTemplate.defaultBroker = abaUsuario ? abaUsuario.getRange('E1').getValue() : "";
  
  htmlTemplate.brokers = JSON.stringify(getBrokers());

// 3. BUSCA DE CONFIGURAÇÕES PERSONALIZADAS (V152)
  // Busca da aba do usuário em vez da aba global
  htmlTemplate.currentSort = getConfigUsuario('OrdenacaoPadrao') || 'ticker_asc'; 
  
  const notificacoesConfig = getConfigUsuario('ReceberNotificacoesDiarias') || 'NAO';
  htmlTemplate.notificacoesAtivas = (notificacoesConfig.trim().toUpperCase() === 'SIM');

  // Define o título da aba do navegador com o e-mail do usuário para facilitar a identificação
  const html = htmlTemplate.evaluate()
    .setTitle('Dashboard de Operações - ' + userEmail)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  return html;
}

/**
 * Lida com as requisições POST. Usado para adicionar um novo ativo à planilha.
 * @param {Object} e O objeto de evento da requisição, contendo os parâmetros.
 * @returns {ContentOutput} Uma resposta em formato JSON com a mensagem de status.
 */
function doPost(e) {
  try {
    const params = e.parameter ? e.parameter : e;
    const planilha = SpreadsheetApp.openById(ID_PLANILHA).getSheetByName(NOME_DA_ABA);
    let newAssetAdded = false;

    // Verifica se os parâmetros necessários foram enviados
    if (params.novoAtivoTicker && params.novoAtivoQtd && params.novoAtivoPM) {
      const novaLinha = [
        params.novoAtivoTicker.toUpperCase(), 
        params.novoAtivoQtd, 
        String(params.novoAtivoPM).replace(',', '.')
      ];
      planilha.appendRow(novaLinha);
      newAssetAdded = true;
    }
    
    const mensagem = newAssetAdded 
      ? `Ativo ${params.novoAtivoTicker.toUpperCase()} salvo com sucesso!` 
      : 'Nenhuma nova operação foi registrada.';
    
    const response = { 
      message: mensagem, 
      newAssetAdded: newAssetAdded
    };
    
    // Retorna a resposta como texto JSON
    return ContentService.createTextOutput(JSON.stringify(response)).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    const errorResponse = { message: `Erro no servidor: ${err.message}` };
    return ContentService.createTextOutput(JSON.stringify(errorResponse)).setMimeType(ContentService.MimeType.JSON);
  }
}