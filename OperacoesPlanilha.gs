/**
 * Salva a corretora padrão selecionada pelo usuário na célula E1 da aba principal.
 * @param {string} brokerName O nome da corretora a ser salva.
 * @returns {string} Uma mensagem de sucesso ou erro.
 */
function salvarCorretoraPadrao(brokerName) {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    const planilha = SpreadsheetApp.openById(ID_PLANILHA).getSheetByName(userEmail);
    planilha.getRange('E1').setValue(brokerName);
    
    registrarLog("CONFIG_CORRETORA", `Alterou corretora padrão para: ${brokerName}`);
    return `Corretora padrão atualizada para ${brokerName}.`;
  } catch (e) {
    return `Erro ao salvar corretora: ${e.message}`;
  }
}

/**
 * Exclui a linha correspondente a um ativo da planilha.
 * @param {string} ticker O ticker do ativo a ser excluído.
 * @returns {{success: boolean, message: string}} Objeto com status e mensagem.
 */
function excluirAtivo(ticker) {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    const planilha = SpreadsheetApp.openById(ID_PLANILHA).getSheetByName(userEmail);
    const range = planilha.getDataRange();
    const dados = range.getValues();
    
    for (let i = 1; i < dados.length; i++) {
      if (dados[i][0].toUpperCase() === ticker.toUpperCase()) {
        planilha.deleteRow(i + 1);
        const cache = CacheService.getScriptCache();
        cache.removeAll([`price_${ticker}`, `yesterday_price_${ticker}`]);
        
        registrarLog("EXCLUIR_ATIVO", `Excluiu o ativo: ${ticker}`);
        return { success: true, message: `Ativo ${ticker} excluído com sucesso.` };
      }
    }
    return { success: false, message: `Ativo ${ticker} não encontrado.` };
  } catch (e) {
    return { success: false, message: `Erro ao excluir: ${e.message}` };
  }
}

/**
 * Busca a lista de todos os tickers disponíveis na aba 'Tickers'.
 * Usa cache para evitar leituras repetidas da planilha.
 * @returns {string} Uma string JSON contendo um array de tickers.
 */
function getAllTickers() {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'all_tickers_list';
  const cachedValue = cache.get(cacheKey);
  if (cachedValue != null) { return cachedValue; }
  try {
    const tickerSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tickers');
    if (!tickerSheet) return '[]';
    const tickers = tickerSheet.getRange('A2:A').getValues().flat().filter(String);
    const tickersJSON = JSON.stringify(tickers);
    cache.put(cacheKey, tickersJSON, 21600); // Cache por 6 horas
    return tickersJSON;
  } catch (e) { return '[]'; }
}

/**
 * Atualiza a quantidade, o preço médio ou o total de um ativo existente.
 * @param {string} ticker O ticker do ativo a ser atualizado.
 * @param {number|null} novaQuantidade A nova quantidade do ativo.
 * @param {number|null} novoPrecoMedio O novo preço médio do ativo.
 * @param {number|null} novoTotal O novo valor total investido (para recalcular o PM).
 * @returns {string} Uma mensagem de sucesso ou erro.
 */
function atualizarAtivoManualmente(ticker, novaQuantidade, novoPrecoMedio, novoTotal) {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    const planilha = SpreadsheetApp.openById(ID_PLANILHA).getSheetByName(userEmail);
    const range = planilha.getDataRange();
    const dados = range.getValues();
    let logDetalhes = `Ativo: ${ticker}`;

    for (let i = 1; i < dados.length; i++) {
      if (dados[i][0].toUpperCase() === ticker.toUpperCase()) {
        if (novaQuantidade) { 
          dados[i][1] = parseFloat(novaQuantidade); 
          logDetalhes += ` | Nova Qtd: ${novaQuantidade}`;
        }
        if (novoPrecoMedio) { 
          dados[i][2] = parseFloat(String(novoPrecoMedio).replace(',', '.')); 
          logDetalhes += ` | Novo PM: ${novoPrecoMedio}`;
        }
        break;
      }
    }
    range.setValues(dados);
    registrarLog("EDITAR_ATIVO", logDetalhes);
    return `Ativo ${ticker} atualizado com sucesso!`;
  } catch (e) {
    return `Erro ao atualizar: ${e.message}`;
  }
}

/**
 * Busca todos os dados dos ativos da carteira, enriquece com preços atuais,
 * calcula as variações e ordena os resultados. DEBUG VERSION.
 * @param {string} [sortBy='ticker'] O campo pelo qual ordenar.
 * @param {string} [sortOrder='asc'] A ordem da classificação.
 * @returns {Array<Object>} Um array ORDENADO de objetos.
 */
function buscarDadosAtivos(sortByInput = null, sortOrderInput = null, abaNome = null) {
  const userEmail = abaNome || Session.getActiveUser().getEmail();
  const planilha = SpreadsheetApp.openById(ID_PLANILHA).getSheetByName(userEmail);
  
  if (!planilha) return [];

  // Define a ordenação baseada na preferência do usuário (E2)
  let sortBy = sortByInput;
  let sortOrder = sortOrderInput;
  if (!sortBy) {
    const configSalva = planilha.getRange('E2').getValue() || 'ticker_asc';
    const partes = configSalva.split('_');
    sortBy = partes[0];
    sortOrder = partes[1] || 'asc';
  }

  const dados = planilha.getDataRange().getValues();
  dados.shift(); // Remove cabeçalho

  const ativosProcessados = dados.map(linha => {
    // Limpa espaços e garante que o ticker existe
    const ativo = (linha[0] && typeof linha[0] === 'string') ? linha[0].trim().toUpperCase() : null;
    
    // Se a linha estiver vazia ou o ticker for inválido, ignora completamente
    if (!ativo || ativo === "" || ativo === "UNDEFINED") return null;

    const quantidade = parseFloat(linha[1]) || 0;
    const precoMedio = parseFloat(linha[2]) || 0;
    
    // Agora o getPreco recebe um nome limpo e real
    const precoAtualNum = parseFloat(getPrecoAtual(ativo)) || 0;
    const precoAnteriorNum = parseFloat(getPrecoFechamentoAnterior(ativo)) || 0;

    return {
      ticker: ativo, 
      quantidade: quantidade, 
      precoMedio: precoMedio,
      investimentoTotal: (quantidade * precoMedio),
      precoAtual: precoAtualNum > 0 ? precoAtualNum.toFixed(2) : "0.00",
      variacaoDia: (precoAnteriorNum > 0) ? ((precoAtualNum / precoAnteriorNum) - 1) * 100 : 0,
      variacaoPM: (precoMedio > 0) ? ((precoAtualNum / precoMedio) - 1) * 100 : 0
    };
  }).filter(Boolean); // Remove os 'null' (linhas vazias) da lista final

  // Ordenação (Mantida a lógica da V152)
  ativosProcessados.sort((a, b) => {
    let valA = a[sortBy];
    let valB = b[sortBy];
    let comparison = (sortBy === 'ticker') ? (valA || '').localeCompare(valB || '') : (parseFloat(valA) || 0) - (parseFloat(valB) || 0);
    return (sortOrder === 'desc') ? (comparison * -1) : comparison;
  });

  return ativosProcessados;
}

/**
 * Função de conveniência para o ServidorWeb carregar dados por aba específica.
 */
function buscarDadosAtivosPorAba(nomeAba) {
  return buscarDadosAtivos('ticker', 'asc', nomeAba);
}

/**
 * Salva a preferência de ordenação do usuário na aba Configurações.
 * @param {string} sortValue O valor da ordenação (ex: "varDia_desc").
 * @returns {string} Uma mensagem de sucesso.
 */
function salvarOrdenacaoPadrao(sortValue) {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    const planilha = SpreadsheetApp.openById(ID_PLANILHA).getSheetByName(userEmail);
    
    if (!planilha) throw new Error("Aba do usuário não encontrada.");

    // Salva na célula E2 da aba privada do usuário
    planilha.getRange('E2').setValue(sortValue);
    
    registrarLog("CONFIG_ORDENACAO", `Usuário definiu ordenação padrão como: ${sortValue}`);
    return `Ordenação salva como ${sortValue}.`;
  } catch (e) {
    return `Erro ao salvar ordenação: ${e.message}`;
  }
}

/**
 * Salva a preferência de recebimento de notificações diárias (VERSÃO CORRETA).
 * @param {boolean | string} receber Pode chegar como boolean ou string "true"/"false".
 * @returns {string} O novo estado salvo ('SIM' ou 'NAO').
 */
function salvarPreferenciaNotificacao(receber) {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    const planilha = SpreadsheetApp.openById(ID_PLANILHA).getSheetByName(userEmail);
    
    if (!planilha) throw new Error("Aba do usuário não encontrada.");

    // Converte o booleano do checkbox em SIM/NAO
    const novoValor = (receber === true || String(receber).toLowerCase() === 'true') ? 'SIM' : 'NAO';

    // Salva na célula E3 da aba privada do usuário
    planilha.getRange('E3').setValue(novoValor);
    
    registrarLog("CONFIG_NOTIFICACAO", `Usuário mudou notificações para: ${novoValor}`);
    return novoValor; 
  } catch (e) {
    console.error("Erro ao salvar preferência: " + e.message);
    // Em caso de erro, retorna o contrário do que foi tentado para o checkbox "voltar" no site
    return (receber === true || String(receber).toLowerCase() === 'true') ? 'NAO' : 'SIM';
  }
}

/**
 * Adiciona um novo ativo na aba específica do usuário logado.
 */
function adicionarAtivo(ticker, quantidade, precoMedio) {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    const planilha = SpreadsheetApp.openById(ID_PLANILHA);
    const abaUsuario = planilha.getSheetByName(userEmail);
    
    if (!abaUsuario) {
      throw new Error("Sua aba pessoal não foi encontrada. Tente atualizar a página.");
    }

    // Adiciona os dados na próxima linha disponível: [Ticker, Quantidade, Preço Médio]
    // Forçamos o Ticker para maiúsculo para manter a organização
    abaUsuario.appendRow([
      ticker.toUpperCase().trim(), 
      parseFloat(quantidade), 
      parseFloat(String(precoMedio).replace(',', '.'))
    ]);
    
    // Registra a ação no Analytics para seu TCC
    registrarLog("ADICIONAR_ATIVO", `Adicionou: ${ticker} | Qtd: ${quantidade} | PM: ${precoMedio}`);
    
    return { success: true, message: `Ativo ${ticker} adicionado com sucesso!` };
  } catch (e) {
    console.error("Erro ao adicionar ativo: " + e.message);
    return { success: false, message: `Erro ao salvar: ${e.message}` };
  }
}