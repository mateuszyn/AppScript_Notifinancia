// =================== CONFIGURAÇÕES GLOBAIS ===================

/**
 * Busca um valor de configuração na aba 'Configurações'.
 * @param {string} key O nome da configuração (na coluna A).
 * @returns {string} O valor encontrado (na coluna B) ou null se não encontrar.
 */
function getConfigValue(key) {
  try {
    const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Configurações');
    const data = configSheet.getRange('A2:B' + configSheet.getLastRow()).getValues();
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === key) {
        return data[i][1];
      }
    }
    return null; // Retorna null se a chave não for encontrada
  } catch (e) {
    // Se a aba 'Configurações' não existir, ou outro erro, retorna null.
    return null;
  }
}

// --- CONFIGURAÇÕES DINÂMICAS ---

// Pega o ID da planilha atual (já estava correto)
const ID_PLANILHA = SpreadsheetApp.getActiveSpreadsheet().getId();

// --- CONFIGURAÇÕES DE MULTI-USUÁRIO E LOGS (V148) ---
const ABA_LOGS = 'Logs';
const ABA_USUARIOS = 'Usuarios';
const ABA_TEMPLATE = 'Template'; // Aba modelo para novos amigos

// --- OPÇÕES PARA O E-MAIL DO USUÁRIO ---

// Abordagem 1 (RECOMENDADA): Pega o e-mail do usuário logado AUTOMATICAMENTE
const SEU_EMAIL = Session.getActiveUser().getEmail();

/*
// Abordagem 2 (Alternativa): Permite ao usuário definir um e-mail na planilha
// Para usar esta, adicione a linha 'EmailRelatorio' na sua aba de Configurações
// e comente a linha de cima.
const SEU_EMAIL = getConfigValue('EmailRelatorio') || Session.getActiveUser().getEmail();
*/


// --- CONFIGURAÇÕES TÉCNICAS ---
const CACHE_EXPIRATION_SECONDS = 900; // 15 minutos