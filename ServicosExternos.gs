/**
 * Valida um ticker e busca seu preço atual.
 * Usado pelo modal "Adicionar Ativo" para verificar se um ticker é válido.
 * @param {string} ticker O ticker do ativo a ser verificado.
 * @returns {{success: boolean, price: number|null, error: string|null}} Objeto com o resultado da operação.
 */
function getSingleTickerPrice(ticker) {
  if (!ticker || typeof ticker !== 'string') {
    return { success: false, error: "Ticker inválido." };
  }
  const upperTicker = ticker.toUpperCase();

  // Valida se o ticker existe na lista da aba 'Tickers'
  const allTickers = JSON.parse(getAllTickers());
  if (!allTickers.includes(upperTicker)) {
    return { success: false, error: `Ticker "${upperTicker}" não encontrado na base de dados.` };
  }

  // Tenta buscar o preço
  const preco = getPrecoAtual(upperTicker);
  if (typeof preco === 'string' && (preco === "N/D" || preco === "Erro")) {
    return { success: false, error: `Não foi possível obter o preço para "${upperTicker}".` };
  }

  return { success: true, price: preco };
}

/**
 * Busca o preço de mercado atual de um único ticker usando GOOGLEFINANCE.
 * Usa um sistema de cache para evitar chamadas repetidas.
 * @param {string} ticker O ticker do ativo.
 * @returns {string} O preço formatado como string ("123.45"), ou "N/D", ou "Erro".
 */
/**
 * Busca o preço de mercado atual.
 */
function getPrecoAtual(ticker) {
  // NOVA TRAVA: Se o ticker estiver vazio ou for undefined, retorna "N/D" imediatamente
  if (!ticker || ticker === "undefined" || ticker === "") return "0.00";

  const cache = CacheService.getScriptCache();
  const cacheKey = `price_${ticker}`;
  const cachedValue = cache.get(cacheKey);
  if (cachedValue != null) return cachedValue;

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    // Usa uma aba fixa de processamento em vez de criar/deletar
    let helperSheet = ss.getSheetByName("_Calculos_");
    if (!helperSheet) helperSheet = ss.insertSheet("_Calculos_").hideSheet();

    const cell = helperSheet.getRange("A1");
    cell.setFormula(`=GOOGLEFINANCE("${ticker}"; "price")`);
    
    // Aumentamos levemente o tempo para garantir que o Google retorne o dado
    SpreadsheetApp.flush(); 
    Utilities.sleep(2000); 

    const preco = cell.getValue();
    if (typeof preco === 'number' && preco > 0) {
      const precoFormatado = preco.toFixed(2);
      cache.put(cacheKey, precoFormatado, 1800); // 30 min de cache
      return precoFormatado;
    }
    return "N/D";
  } catch (e) {
    return "Erro";
  }
}

/**
 * Busca o preço de fechamento do dia anterior.
 */
function getPrecoFechamentoAnterior(ticker) {
  if (!ticker || ticker === "undefined" || ticker === "") return "0.00";

  const cache = CacheService.getScriptCache();
  const cacheKey = `yesterday_price_${ticker}`;
  const cachedValue = cache.get(cacheKey);
  if (cachedValue != null) return cachedValue;

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let helperSheet = ss.getSheetByName("_Calculos_");
    if (!helperSheet) helperSheet = ss.insertSheet("_Calculos_").hideSheet();

    const cell = helperSheet.getRange("B1");
    // Usamos a fórmula de fechamento anterior que é mais rápida que a busca histórica por data
    cell.setFormula(`=IFERROR(GOOGLEFINANCE("${ticker}"; "closeyest"); 0)`);
    
    // LOOP DE PERSISTÊNCIA: Tenta ler o valor até 5 vezes se ele estiver carregando
    let tentativas = 0;
    let preco = 0;
    
    while (tentativas < 5) {
      SpreadsheetApp.flush(); // Força o Sheets a calcular
      Utilities.sleep(1500);  // Espera 1.5s por tentativa
      preco = cell.getValue();
      
      // Se já for um número válido, sai do loop
      if (typeof preco === 'number' && preco > 0) break;
      
      tentativas++;
      console.warn(`Tentativa ${tentativas} para ${ticker}: Google Finance ainda carregando...`);
    }

    if (typeof preco === 'number' && preco > 0) {
      const precoFormatado = preco.toFixed(2);
      cache.put(cacheKey, precoFormatado, 1800);
      return precoFormatado;
    }
    
    registrarLog("ERRO_FINANCE", `Não foi possível obter preço anterior de ${ticker} após 5 tentativas.`);
    return "N/D";
  } catch (e) {
    return "N/D";
  }
}