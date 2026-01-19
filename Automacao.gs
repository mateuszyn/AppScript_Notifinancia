/**
 * Função principal que será agendada no gatilho (Trigger).
 * Ela percorre todos os usuários cadastrados e envia o relatório individual.
 */
function dispararRelatorioParaTodos() {
  const planilha = SpreadsheetApp.openById(ID_PLANILHA);
  const abaUsuarios = planilha.getSheetByName(ABA_USUARIOS);
  
  if (!abaUsuarios) return;

  const dadosUsuarios = abaUsuarios.getDataRange().getValues();
  dadosUsuarios.shift(); // Remove o cabeçalho [Email, Data, Status]

  dadosUsuarios.forEach(linha => {
    const email = linha[0];
    const status = linha[2];

    // Só envia se houver um e-mail válido e o status for "Ativo"
    if (email && status === "Ativo") {
      try {
        processarEnvioIndividual(email);
        registrarLog("EMAIL_DIARIO", `Relatório enviado com sucesso para: ${email}`);
      } catch (e) {
        registrarLog("ERRO_EMAIL", `Falha ao enviar para ${email}: ${e.message}`);
      }
    }
  });
}

/**
 * Prepara os dados e envia o e-mail para um único usuário.
 */
function processarEnvioIndividual(email) {
  const dadosAtivos = buscarDadosAtivos(null, null, email);
  
  if (dadosAtivos.length === 0) return; 

  const ativosAbaixoDoPM = dadosAtivos.filter(ativo => ativo.variacaoPM < 0);
  let mensagemDeAlerta = '';

  // 1. LÓGICA DE ALERTA (Vermelho) - Quando há prejuízo
  if (ativosAbaixoDoPM.length > 0) {
    let tickersArray = ativosAbaixoDoPM.map(a => `<b>${a.ticker} (${a.variacaoPM.toFixed(2)}%)</b>`);
    let listaDeTickers = tickersArray.length === 1 ? tickersArray[0] : tickersArray.slice(0, -1).join(', ') + ' e ' + tickersArray.slice(-1);

    mensagemDeAlerta = `
      <div style="padding: 15px; border: 1px solid #dc3545; background-color: #f8d7da; border-radius: 8px; color: #721c24; text-align: center; margin-bottom: 25px; font-size: 14px;">
        <strong style="font-size: 16px;">Atenção!</strong><br>
        Sua carteira tem ativos abaixo do Preço Médio: ${listaDeTickers}
      </div>`;
  } 
  // 2. LÓGICA DE PARABÉNS (Verde) - Quando tudo está valorizado
  else {
    // Encontra o ativo com a menor valorização positiva (o "menos valorizado")
    const menorValorizacao = dadosAtivos.reduce((min, atual) => (atual.variacaoPM < min.variacaoPM ? atual : min), dadosAtivos[0]);

    mensagemDeAlerta = `
      <div style="padding: 15px; border: 1px solid #28a745; background-color: #d4edda; border-radius: 8px; color: #155724; text-align: center; margin-bottom: 25px; font-size: 14px;">
        <strong style="font-size: 16px;">Parabéns!</strong><br>
        Sua carteira está toda valorizada! Sua menor valorização é <b>${menorValorizacao.ticker} (${menorValorizacao.variacaoPM.toFixed(2)}%)</b>
      </div>`;
  }

  const corpoHtml = gerarCorpoEmail(email, dadosAtivos, mensagemDeAlerta);
  
  MailApp.sendEmail({
    to: email,
    subject: `Resumo Diário de Investimentos - ${email}`,
    htmlBody: corpoHtml
  });
}
/**
 * Constrói o HTML do e-mail com as novas cores solicitadas (Verde do site e cores das corretoras).
 */
function gerarCorpoEmail(email, dadosAtivos, mensagemDeAlerta) {
  let tabelaHtml = '';
  const verdeSite = '#28a745'; // Cor verde padrão de sucesso/adicionar do seu site

  dadosAtivos.forEach(ativo => {
      // 1. GARANTIA DE VALORES: Se o valor não existir, vira 0 para não dar erro de 'undefined'
      const variacao = ativo.variacaoPM !== undefined ? ativo.variacaoPM : 0;
      const preco = ativo.precoAtual !== undefined ? ativo.precoAtual : "0.00";
      
      const corVariacao = variacao >= 0 ? '#28a745' : '#dc3545';
      const seta = variacao >= 0 ? '▲' : '▼';
      
      tabelaHtml += `
        <tr>
          <td style="padding: 12px 8px; border-bottom: 1px solid #eee; font-weight: bold; color: #333;">${ativo.ticker}</td>
          <td style="padding: 12px 8px; border-bottom: 1px solid #eee; color: #666;">${ativo.quantidade}</td>
          <td style="padding: 12px 8px; border-bottom: 1px solid #eee; color: #333;">R$ ${preco}</td>
          <td style="padding: 12px 8px; border-bottom: 1px solid #eee; color: ${corVariacao}; font-weight: bold;">
            ${seta} ${Math.abs(Number(variacao)).toFixed(2)}%
          </td>
        </tr>`;
  });

  const urlSistema = ScriptApp.getService().getUrl();

  return `
    <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; max-width: 600px; margin: auto; background-color: #f4f7f6; padding: 20px; border-radius: 10px;">
      <div style="background-color: #ffffff; padding: 30px; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.05);">
        <h2 style="color: ${verdeSite}; text-align: center; margin-top: 0; font-size: 24px;">Resumo da Carteira</h2>
        <p style="text-align: center; color: #666;">Olá, <strong>${email.split('@')[0]}</strong>! Veja o status dos seus investimentos:</p>
        
        ${mensagemDeAlerta}

        <table style="width: 100%; border-collapse: collapse; margin-top: 20px;">
          <thead>
            <tr style="background-color: #f8f9fa; color: #777; text-transform: uppercase; font-size: 11px; letter-spacing: 1px;">
              <th style="padding: 10px; border-bottom: 2px solid #eee; text-align: left;">Ativo</th>
              <th style="padding: 10px; border-bottom: 2px solid #eee; text-align: left;">Qtd</th>
              <th style="padding: 10px; border-bottom: 2px solid #eee; text-align: left;">Preço</th>
              <th style="padding: 10px; border-bottom: 2px solid #eee; text-align: left;">P.M.</th>
            </tr>
          </thead>
          <tbody>
            ${tabelaHtml}
          </tbody>
        </table>

        <div style="text-align: center; margin-top: 35px;">
          <a href="${urlSistema}" style="background-color: #ffffff; color: ${verdeSite}; border: 2px solid ${verdeSite}; padding: 14px 28px; text-decoration: none; border-radius: 6px; font-weight: bold; display: inline-block; transition: all 0.3s;">
            Acessar Meu Dashboard Completo
          </a>
        </div>

        <hr style="border: 0; border-top: 1px solid #eee; margin: 40px 0 20px 0;">
        
        <p style="font-size: 13px; font-weight: bold; color: #555; text-align: center; margin-bottom: 15px;">Operar em sua corretora:</p>
        <div style="text-align: center;">
          ${gerarLinksCorretorasHtml()}
        </div>

        <p style="font-size: 11px; color: #aaa; text-align: center; margin-top: 40px; line-height: 1.5;">
          Acesse sempre o último E-Mail para a versão mais recente do NotiFinancia. Mude preferências de notificação e ordenação acessando o Dashboard Completo.
        </p>
      </div>
    </div>`;
}

/**
 * Gera os botões de links para as corretoras com as URLs e cores exatas fornecidas.
 * Atualizado com: Nubank, Inter, XP e Rico.
 */
function gerarLinksCorretorasHtml() {
  const corretoras = [
    { name: 'Nubank', appUrl: 'https://play.google.com/store/apps/details?id=com.nu.production', color: '#820AD1', textColor: '#FFFFFF' },
    { name: 'Inter', appUrl: 'https://play.google.com/store/apps/details?id=br.com.intermedium', color: '#FF7A00', textColor: '#FFFFFF' },
    { name: 'XP', appUrl: 'https://play.google.com/store/apps/details?id=br.com.xp.investimentos', color: '#212529', textColor: '#FFFFFF' },
    { name: 'Rico', appUrl: 'https://play.google.com/store/apps/details?id=com.rico.fox', color: '#005AAA', textColor: '#FF8A00' }
  ];

  return corretoras.map(c => `
    <a href="${c.appUrl}" style="display: inline-block; padding: 10px 16px; margin: 5px; background-color: ${c.color}; color: ${c.textColor}; text-decoration: none; border-radius: 6px; font-size: 11px; font-weight: bold; min-width: 90px; border: 1px solid ${c.color === '#FFFFFF' ? '#ddd' : c.color};">
      ${c.name}
    </a>
  `).join('');
}

/**
 * Envia um e-mail de teste exclusivamente para o desenvolvedor.
 * Útil para validar o layout e os links das corretoras.
 */
function enviarEmailTeste() {
  const emailTeste = "";
  
  try {
    console.log("--- Iniciando Envio de E-mail Teste ---");
    
    // Chama a função de processamento individual definida na V148
    // Ela buscará os dados da aba que tem o nome do seu e-mail
    processarEnvioIndividual(emailTeste);
    
    // Registra no sistema de auditoria do seu TCC
    registrarLog("TESTE_EMAIL", `E-mail de teste enviado com sucesso para: ${emailTeste}`);
    
    console.log("Teste concluído com sucesso!");
  } catch (e) {
    console.error("Falha no e-mail teste: " + e.message);
    registrarLog("ERRO_TESTE", `Falha ao enviar teste para ${emailTeste}: ${e.message}`);
  }
}