// ====================================================================================
// 🧪 AMBIENTE DE TESTES ISOLADO E BLINDADO (SANDBOX)
// ====================================================================================


/**
 * Lê a planilha de testes usando a EXATA mesma lógica da produção, 
 * mas sem reaproveitar funções para não gerar dependência cruzada.
 */
function web_obterFilaTesteIsolada() {
  const ssTeste = SpreadsheetApp.openById(PLANILHA_TESTE_ID);
  // Reutiliza a lógica bruta de varredura, mas apontando apenas para o ssTeste
  const abas = ssTeste.getSheets().filter(s => s.getName().includes("1 -") || s.getName().includes("2 -") || s.getName().includes("3 -"));
  const filaTeste = [];
  
  abas.forEach(aba => {
    const nomeAba = aba.getName();
    let numEtapa = nomeAba.includes("2 -") ? 2 : nomeAba.includes("3 -") ? 3 : 1;
    const dados = aba.getDataRange().getValues();
    
    for (let i = 1; i < dados.length; i++) {
      const l = dados[i];
      const nome = l[MAPA_COLUNAS.NOME] ? String(l[MAPA_COLUNAS.NOME]).trim() : "";
      const placa = l[MAPA_COLUNAS.PLACA] ? String(l[MAPA_COLUNAS.PLACA]).trim() : "";
      const chassi = l[MAPA_COLUNAS.CHASSI] ? String(l[MAPA_COLUNAS.CHASSI]).trim() : "";
      
      if (!placa && !chassi && !nome) continue;

      let dataEntradaStr = "Data não registrada";
      if (l[MAPA_COLUNAS.DATA] instanceof Date) {
         dataEntradaStr = Utilities.formatDate(l[MAPA_COLUNAS.DATA], Session.getScriptTimeZone(), "dd/MM/yyyy");
      } else {
         const strData = String(l[MAPA_COLUNAS.DATA] || "").split(" ")[0];
         if (strData && strData.includes("/")) dataEntradaStr = strData;
      }

      filaTeste.push({
        idUnico: nomeAba + "-" + (i + 1),
        etapaNum: numEtapa, 
        linhaOriginal: i + 1, 
        abaNome: nomeAba, 
        nome: nome, 
        placa: placa, 
        chassi: chassi,
        // O e-mail original vai para o Front, mas o Backend vai ignorá-lo no disparo
        emailOriginal: l[MAPA_COLUNAS.EMAIL] ? String(l[MAPA_COLUNAS.EMAIL]).trim() : "Sem E-mail",
        dataPlanilha: dataEntradaStr,
        isEnviado: (l[MAPA_COLUNAS.CHECK_EMAIL] === true || l[MAPA_COLUNAS.CHECK_EMAIL] === "TRUE" || l[MAPA_COLUNAS.CHECK_EMAIL] === 1)
      });
    }
  });
  
  return filaTeste;
}

/**
 * Dispara os e-mails com TRAVA DE SEGURANÇA. 
 * O destino real é ignorado e sobrescrito para rastreamentozen.03@gmail.com
 */
function dispararEmailTesteIsolado(grupos, responsavel) {
  const ssTeste = SpreadsheetApp.openById(PLANILHA_TESTE_ID);
  const dt = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
  const templatesDict = getTemplatesDict(ssTeste); // Pega templates da planilha de teste
  
  let enviados = 0;
  
  grupos.forEach(g => {
    let txt = "", tituloHeader = "", ass = "";
    const isPlural = g.veiculosStr.includes(",");
    const lblVeic = isPlural ? "Veículos" : "Veículo";
    
    if (g.etapaNum === 1) {
      tituloHeader = "BEM-VINDO À ZEN SEGUROS (TESTE)";
      ass = `Bem-vindo à ZEN Seguros - Orientações - ${lblVeic}: ${g.veiculosStr}`;
      txt = aplicarTemplate(templatesDict, "BOAS_VINDAS_NORMAL", g.nome, g.veiculosStr, isPlural, 0, 10, g.dataEntrada);
    } else if (g.etapaNum === 2) {
      tituloHeader = "LEMBRETE: INSTALAÇÃO PENDENTE (TESTE)";
      ass = `Lembrete: Instalação Pendente - ${lblVeic}: ${g.veiculosStr}`;
      txt = aplicarTemplate(templatesDict, "LEMBRETE_5_DIAS", g.nome, g.veiculosStr, isPlural, 0, 10, g.dataEntrada);
    } else {
      tituloHeader = "URGENTE: PRAZO EXPIRADO (TESTE)";
      ass = `[URGENTE] Prazo Expirado! ${lblVeic}: ${g.veiculosStr}`;
      txt = aplicarTemplate(templatesDict, "PRAZO_EXPIRADO", g.nome, g.veiculosStr, isPlural, 0, 10, g.dataEntrada);
    }

    // ALERTA VISUAL NO CORPO DO EMAIL PARA IDENTIFICAR QUE É TESTE
    const alertaTesteHtml = `<div style="background-color: #fca5a5; color: #991b1b; padding: 10px; font-weight: bold; text-align: center; border-radius: 5px; margin-bottom: 15px;">⚠️ AMBIENTE DE TESTES DO SGCW ⚠️<br>E-mail original do cliente: ${g.emailOriginal}</div>`;
    const htmlFinal = formatarComoEmail(alertaTesteHtml + txt, tituloHeader);
    
    // =========================================================
    // TRAVA DE SEGURANÇA: OVERRIDE DO DESTINATÁRIO
    // =========================================================
    MailApp.sendEmail({
      to: EMAIL_TRAVA_TESTE, // SEMPRE manda para rastreamentozen.03@gmail.com
      subject: "[TESTE] " + ass,
      body: txt,
      htmlBody: htmlFinal, 
      name: "Laboratório ZEN"
    });
    
    // Atualiza APENAS a planilha de teste
    g.linhas.forEach(cli => {
      const aba = ssTeste.getSheetByName(cli.abaNome);
      if (aba) {
        aba.getRange(cli.linhaOriginal, MAPA_COLUNAS.CHECK_EMAIL + 1).setValue(true);
        aba.getRange(cli.linhaOriginal, MAPA_COLUNAS.DATA_EMAIL + 1).setValue(dt);
        aba.getRange(cli.linhaOriginal, MAPA_COLUNAS.RESPONSAVEL + 1).setValue(responsavel + " (Modo Teste)");
      }
    });
    enviados++;
  });
  
  return `Disparo de Teste Seguro Concluído!\n${enviados} e-mails redirecionados com sucesso para ${EMAIL_TRAVA_TESTE}.`;
}
/**
 * 🧪 AMBIENTE DE LABORATÓRIO: Remove todas as marcações de Check e Respostas de uma lista de clientes.
 * Esta função é blindada e SÓ atua na planilha de testes, ignorando a produção.
 */
function resetarMarcacoesTesteBackend(idsUnicos) {
  // ID fixo para impedir a destruição da base principal por acidente
  const PLANILHA_TESTE_ID = "1c2_JOsPRbVttvfhxi8qEhpK2xkvrnwY2TbqwRp_xEjo";
  const ssTeste = SpreadsheetApp.openById(PLANILHA_TESTE_ID);
  
  let cont = 0;
  const operacoes = {};
  
  // Organiza os IDs por aba para otimizar as chamadas
  idsUnicos.forEach(id => {
      const partes = id.lastIndexOf('-');
      const abaNome = id.substring(0, partes);
      const linha = parseInt(id.substring(partes + 1));
      if (!operacoes[abaNome]) operacoes[abaNome] = [];
      operacoes[abaNome].push(linha);
  });
  
  // Limpa as colunas baseando-se no índice exato
  for (const abaNome in operacoes) {
      const aba = ssTeste.getSheetByName(abaNome);
      if (!aba) continue;
      
      operacoes[abaNome].forEach(linha => {
          // Os "+1" servem para alinhar a contagem de array do JS com as células do Google Sheets
          aba.getRange(linha, 8 + 1).setValue(false); // Coluna I: Check E-mail
          aba.getRange(linha, 9 + 1).setValue("");    // Coluna J: Data E-mail
          aba.getRange(linha, 10 + 1).setValue(false); // Coluna K: Respondeu E-mail
          aba.getRange(linha, 11 + 1).setValue(false); // Coluna L: Check Whats
          aba.getRange(linha, 12 + 1).setValue("");    // Coluna M: Data Whats
          aba.getRange(linha, 13 + 1).setValue(false); // Coluna N: Respondeu Whats
          aba.getRange(linha, 14 + 1).setValue("");    // Coluna O: Responsável
          
          cont++;
      });
  }
  return `O histórico de envios de ${cont} clientes foi zerado no Laboratório.`;
}