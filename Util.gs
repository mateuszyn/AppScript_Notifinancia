// Arquivo: Util.gs

function getBrokers() {
  return [
      { name: 'Nubank', type: 'app', appUrl: 'https://play.google.com/store/apps/details?id=com.nu.production', color: '#820AD1', textColor: '#FFFFFF' },
      { name: 'Inter', type: 'app', appUrl: 'https://play.google.com/store/apps/details?id=br.com.intermedium', color: '#FF7A00', textColor: '#FFFFFF' },
      { name: 'XP', type: 'web', webUrl: 'https://www.xpi.com.br/login/', appUrl: 'https://play.google.com/store/apps/details?id=br.com.xp.investimentos', color: '#212529', textColor: '#FFFFFF' },
      { name: 'Rico', type: 'web', webUrl: 'https://www.rico.com.vc/login/', appUrl: 'https://play.google.com/store/apps/details?id=com.rico.fox', color: '#005AAA', textColor: '#FF8A00' }
  ];
}

function testarFuncaoBrokers() {
  try {
    const listaDeBrokers = getBrokers();
    console.log("Resultado da função getBrokers():");
    console.log(listaDeBrokers);
    console.log("Resultado após JSON.stringify:");
    console.log(JSON.stringify(listaDeBrokers));
  } catch (e) {
    console.error("Ocorreu um erro ao testar getBrokers(): " + e.message);
  }
}

/**
 * Registra uma ação no sistema na aba 'Logs'.
 * @param {string} acao O nome da ação (Ex: ACESSO, EDITAR, EXCLUIR).
 * @param {string} detalhes Informações adicionais da ação.
 */
function registrarLog(acao, detalhes) {
  try {
    const planilha = SpreadsheetApp.openById(ID_PLANILHA);
    const abaLogs = planilha.getSheetByName(ABA_LOGS);
    
    // Tenta pegar o e-mail; se falhar/estiver vazio, usa "Desconhecido"
    let userEmail = Session.getActiveUser().getEmail();
    if (!userEmail || userEmail === "") userEmail = "Usuario_Nao_Identificado";
    
    const agora = new Date();
    abaLogs.appendRow([agora, userEmail, acao, detalhes]);
  } catch (e) {
    console.error("Erro ao registrar log: " + e.message);
  }
}

function garantirUsuario(email) {
  if (!email || email === "") {
    registrarLog("ALERTA", "Tentativa de acesso sem e-mail identificado.");
    return; 
  }

  const planilha = SpreadsheetApp.openById(ID_PLANILHA);
  const abaUsuarios = planilha.getSheetByName(ABA_USUARIOS);
  const dadosUsuarios = abaUsuarios.getDataRange().getValues();
  
  // Verifica se o usuário existe (ignora linhas vazias)
  const usuarioExiste = dadosUsuarios.some(linha => linha[0] === email);
  
  if (!usuarioExiste) {
    abaUsuarios.appendRow([email, new Date(), "Ativo"]);
    registrarLog("NOVO_USUARIO", "Criando ambiente para: " + email);
    
    const abaTemplate = planilha.getSheetByName(ABA_TEMPLATE);
    // Cria a aba apenas se ela realmente não existir
    if (!planilha.getSheetByName(email)) {
      abaTemplate.copyTo(planilha).setName(email);
      // Limpa possíveis resíduos que o copyTo possa levar (opcional)
      registrarLog("SISTEMA", "Aba criada com sucesso para " + email);
    }
  }
}

/**
 * Busca uma configuração específica na aba do usuário logado.
 * @param {string} chave Nome da configuração (ex: 'Ordenacao', 'Notificacao').
 */
function getConfigUsuario(chave) {
  const userEmail = Session.getActiveUser().getEmail();
  const planilha = SpreadsheetApp.openById(ID_PLANILHA).getSheetByName(userEmail);
  if (!planilha) return null;

  // Mapeamento de células para cada configuração
  const mapa = {
    'OrdenacaoPadrao': 'E2',
    'ReceberNotificacoesDiarias': 'E3'
  };

  const celula = mapa[chave];
  return celula ? planilha.getRange(celula).getValue() : null;
}

/**
 * Salva uma configuração específica na aba do usuário logado.
 */
function salvarConfigUsuario(chave, valor) {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    const planilha = SpreadsheetApp.openById(ID_PLANILHA).getSheetByName(userEmail);
    
    const mapa = {
      'OrdenacaoPadrao': 'E2',
      'ReceberNotificacoesDiarias': 'E3'
    };

    const celula = mapa[chave];
    if (celula) {
      planilha.getRange(celula).setValue(valor);
      registrarLog("CONFIG_ALTERADA", `Config ${chave} mudou para ${valor}`);
      return { success: true, message: "Configuração salva com sucesso!" };
    }
  } catch (e) {
    return { success: false, message: "Erro ao salvar: " + e.message };
  }
}