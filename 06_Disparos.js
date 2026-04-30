// ====================================================================================
// OBTENÇÃO DE PRÉVIA E DISPARO CANAL A CANAL
// ====================================================================================

// Variáveis Globais de Teste (Segurança)


function obterPreviewDisparoAgrupadoWeb(grupos, isTeste) {
  const idPlanilhaParaUso = isTeste ? PLANILHA_TESTE_ID : PLANILHA_ID;
  const ss = SpreadsheetApp.openById(idPlanilhaParaUso);
  const templatesDict = getTemplatesDict(ss);

  let feriadosTime = [];
  try {
    const abaFeriados = ss.getSheetByName("Feriados");
    if (abaFeriados) {
      feriadosTime = abaFeriados.getRange("A2:A").getValues().map(r => r[0] instanceof Date ? r[0].getTime() : null).filter(r => r);
    }
  } catch(e) {}
  const dtHoje = new Date();

  return grupos.map(g => {
    let ass = "", txt = "", tituloHeader = "";
    
    const isPlural = g.veiculosStr.includes(",");
    const lblVeic = isPlural ? "Veículos" : "Veículo";

    let diasDecorridosParaSLA = 0;
    let limiteBaseSLA = 10; 
    
    let dtEntrada = null;
    let dataEntradaFormatada = "Data não registrada";
    const dPlanilhaStr = g.dataEntrada || (g.linhas && g.linhas.length > 0 ? g.linhas[0].dataPlanilha : "") || "";
    
    if (dPlanilhaStr) {
       dataEntradaFormatada = String(dPlanilhaStr).split(" ")[0];
       const partes = dataEntradaFormatada.split("/");
       if (partes.length === 3) {
          dtEntrada = new Date(partes[2], partes[1] - 1, partes[0]);
       }
    }

    if (g.etapaNum === 1 || g.etapaNum === 2) {
        if (dtEntrada) {
            try { diasDecorridosParaSLA = calcularDiasUteis(dtEntrada, dtHoje, feriadosTime); } catch(e) {}
        }
    } else if (g.etapaNum === 3) {
        const dEmailStr = (g.linhas && g.linhas.length > 0 ? g.linhas[0].dataEmail : "") || "";
        let dEmailObj = null;
        if (dEmailStr && dEmailStr !== "Aguardando...") {
           const partes = String(dEmailStr).split(" ")[0].split("/");
           if (partes.length === 3) {
              dEmailObj = new Date(partes[2], partes[1] - 1, partes[0]);
           }
        }
        if (dtEntrada && dEmailObj) {
            try { diasDecorridosParaSLA = calcularDiasUteis(dtEntrada, dEmailObj, feriadosTime); } catch(e) {}
        } else {
            diasDecorridosParaSLA = 5;
        }
    }

    if (g.etapaNum === 1) {
      tituloHeader = "BEM-VINDO À ZEN SEGUROS";
      ass = `Bem-vindo à ZEN Seguros - Orientações - ${lblVeic}: ${g.veiculosStr}`;
      txt = aplicarTemplate(templatesDict, g.isFipeBaixa ? "BOAS_VINDAS_FIPE_BAIXA" : "BOAS_VINDAS_NORMAL", g.nome, g.veiculosStr, isPlural, diasDecorridosParaSLA, limiteBaseSLA, dataEntradaFormatada);
    } else if (g.etapaNum === 2) {
      tituloHeader = "LEMBRETE: INSTALAÇÃO PENDENTE";
      ass = `Lembrete: Instalação Pendente - ${lblVeic}: ${g.veiculosStr}`;
      txt = aplicarTemplate(templatesDict, "LEMBRETE_5_DIAS", g.nome, g.veiculosStr, isPlural, diasDecorridosParaSLA, limiteBaseSLA, dataEntradaFormatada);
    } else {
      tituloHeader = "URGENTE: PRAZO EXPIRADO";
      ass = `[URGENTE] Prazo Expirado! ${lblVeic}: ${g.veiculosStr}`;
      txt = aplicarTemplate(templatesDict, "PRAZO_EXPIRADO", g.nome, g.veiculosStr, isPlural, diasDecorridosParaSLA, limiteBaseSLA, dataEntradaFormatada);
    }

    let htmlBodyFormatado = formatarComoEmail(txt, tituloHeader);
    
    // Alerta visual no preview se estiver em modo teste
    if (isTeste) {
        const alertaTesteHtml = `<div style="background-color: #fca5a5; color: #991b1b; padding: 10px; font-weight: bold; text-align: center; border-radius: 5px; margin-bottom: 15px;">⚠️ AMBIENTE DE TESTES DO SGCW ⚠️<br>O e-mail original (${g.email}) será ignorado e enviado para a caixa de testes.</div>`;
        htmlBodyFormatado = alertaTesteHtml + htmlBodyFormatado;
    }

    let disclaimerFormatado = aplicarTemplate(templatesDict, "WHATSAPP_DISCLAIMER", g.nome, g.veiculosStr, isPlural, diasDecorridosParaSLA, limiteBaseSLA, dataEntradaFormatada);
    let msgWhats = "";
    
    if (disclaimerFormatado && disclaimerFormatado.indexOf("⚠️ Erro:") === -1) {
      msgWhats = disclaimerFormatado + "\n\n" + txt;
    } else {
      msgWhats = "> *MENSAGEM AUTOMÁTICA*\n> _Esse WhatsApp é utilizado apenas para envio de recados_\n> _Nossos contatos estarão disponíveis no final da mensagem_\n\n" + txt;
    }

    let telefoneBase = (g.linhas && g.linhas.length > 0) ? (g.linhas[0].telefone || "") : "";
    let numeroLimpo = telefoneBase.toString().replace(/\D/g, "");
    if (numeroLimpo.length >= 10 && !numeroLimpo.startsWith("55")) numeroLimpo = "55" + numeroLimpo;
    
    return { 
      email: g.email, nome: g.nome, veiculosStr: g.veiculosStr, 
      etapaNum: g.etapaNum, assunto: ass, emailHtml: htmlBodyFormatado, 
      whatsText: msgWhats, telefoneLimpo: numeroLimpo,
      isErroEmail: g.isErroEmail, isEnviado: g.isEnviado, isInativo: g.isInativo 
    };
  });
}

function dispararEmailAgrupadoWeb(grupos, responsavel, isTeste) {
  const idPlanilhaParaUso = isTeste ? PLANILHA_TESTE_ID : PLANILHA_ID;
  const ss = SpreadsheetApp.openById(idPlanilhaParaUso);
  const dt = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
  const templatesDict = getTemplatesDict(ss);
  let errosCriticos = [];

  let feriadosTime = [];
  try {
    const abaFeriados = ss.getSheetByName("Feriados");
    if (abaFeriados) {
      feriadosTime = abaFeriados.getRange("A2:A").getValues().map(r => r[0] instanceof Date ? r[0].getTime() : null).filter(r => r);
    }
  } catch(e) {}
  const dtHoje = new Date();

  grupos.forEach(g => {
    let ass = "", txt = "", ac = "", tituloHeader = "";
    const isPlural = g.veiculosStr.includes(",");
    const lblVeic = isPlural ? "Veículos" : "Veículo";

    let diasDecorridosParaSLA = 0;
    let limiteBaseSLA = 10; 
    
    let dtEntrada = null;
    let dataEntradaFormatada = "Data não registrada";
    const dPlanilhaStr = g.dataEntrada || (g.linhas && g.linhas.length > 0 ? g.linhas[0].dataPlanilha : "") || "";
    if (dPlanilhaStr) {
       dataEntradaFormatada = String(dPlanilhaStr).split(" ")[0];
       const partes = dataEntradaFormatada.split("/");
       if (partes.length === 3) {
          dtEntrada = new Date(partes[2], partes[1] - 1, partes[0]);
       }
    }

    if (g.etapaNum === 1 || g.etapaNum === 2) {
        if (dtEntrada) {
            try { diasDecorridosParaSLA = calcularDiasUteis(dtEntrada, dtHoje, feriadosTime); } catch(e) {}
        }
    } else if (g.etapaNum === 3) {
        const dEmailStr = (g.linhas && g.linhas.length > 0 ? g.linhas[0].dataEmail : "") || "";
        let dEmailObj = null;
        if (dEmailStr && dEmailStr !== "Aguardando...") {
           const partes = String(dEmailStr).split(" ")[0].split("/");
           if (partes.length === 3) {
              dEmailObj = new Date(partes[2], partes[1] - 1, partes[0]);
           }
        }
        if (dtEntrada && dEmailObj) {
            try { diasDecorridosParaSLA = calcularDiasUteis(dtEntrada, dEmailObj, feriadosTime); } catch(e) {}
        } else {
            diasDecorridosParaSLA = 5;
        }
    }

    if (g.etapaNum === 1) {
      tituloHeader = "BEM-VINDO À ZEN SEGUROS";
      ass = `Bem-vindo à ZEN Seguros - Orientações - ${lblVeic}: ${g.veiculosStr}`;
      txt = aplicarTemplate(templatesDict, g.isFipeBaixa ? "BOAS_VINDAS_FIPE_BAIXA" : "BOAS_VINDAS_NORMAL", g.nome, g.veiculosStr, isPlural, diasDecorridosParaSLA, limiteBaseSLA, dataEntradaFormatada);
      ac = "1_EMAIL";
    } else if (g.etapaNum === 2) {
      tituloHeader = "LEMBRETE: INSTALAÇÃO PENDENTE";
      ass = `Lembrete: Instalação Pendente - ${lblVeic}: ${g.veiculosStr}`;
      txt = aplicarTemplate(templatesDict, "LEMBRETE_5_DIAS", g.nome, g.veiculosStr, isPlural, diasDecorridosParaSLA, limiteBaseSLA, dataEntradaFormatada);
      ac = "2_EMAIL";
    } else {
      tituloHeader = "URGENTE: PRAZO EXPIRADO";
      ass = `[URGENTE] Prazo Expirado! ${lblVeic}: ${g.veiculosStr}`;
      txt = aplicarTemplate(templatesDict, "PRAZO_EXPIRADO", g.nome, g.veiculosStr, isPlural, diasDecorridosParaSLA, limiteBaseSLA, dataEntradaFormatada);
      ac = "3_EMAIL";
    }

    try {
      let htmlPadrao = formatarComoEmail(txt, tituloHeader);
      
      // Adiciona o selo de teste no corpo do e-mail, caso esteja no Laboratório
      if (isTeste) {
          const alertaTesteHtml = `<div style="background-color: #fca5a5; color: #991b1b; padding: 10px; font-weight: bold; text-align: center; border-radius: 5px; margin-bottom: 15px;">⚠️ AMBIENTE DE TESTES DO SGCW ⚠️<br>E-mail original do cliente: ${g.email}</div>`;
          htmlPadrao = alertaTesteHtml + htmlPadrao;
      }

      const assuntoFinal = isTeste ? "[MODO TESTE] " + (g.customAssunto ? g.customAssunto : ass) : (g.customAssunto ? g.customAssunto : ass);
      const htmlFinal = g.customEmailHtml ? g.customEmailHtml : htmlPadrao;
      const textoPuroFallback = txt + "\n\nAtenciosamente,\nSetor de Rastreamento\nZEN Seguros";
      
      // TRAVA DE DESTINATÁRIO (Se for teste, sobrepõe o e-mail do cliente pelo e-mail da Rastreamentozen)
      const emailAlvo = isTeste ? EMAIL_TESTE_TRAVA : g.email;

      MailApp.sendEmail({
        to: emailAlvo, 
        subject: assuntoFinal,
        body: textoPuroFallback,
        htmlBody: htmlFinal, 
        name: "Setor de Rastreamento - ZEN Seguros"
      });
      
      g.linhas.forEach(cli => {
        const aba = ss.getSheetByName(cli.abaNome);
        if (aba) {
          const nomeResponsavelReal = isTeste ? responsavel + " (Modo Teste)" : responsavel;
          aba.getRange(cli.linhaOriginal, MAPA_COLUNAS.CHECK_EMAIL + 1).setValue(true);
          aba.getRange(cli.linhaOriginal, MAPA_COLUNAS.DATA_EMAIL + 1).setValue(dt);
          aba.getRange(cli.linhaOriginal, MAPA_COLUNAS.RESPONSAVEL + 1).setValue(nomeResponsavelReal);
          registrarAuditoriaExata(cli.nome, cli.placa, cli.chassi, cli.email, cli.telefone, ac, dt, nomeResponsavelReal);
        }
      });
    } catch (e) {
      errosCriticos.push(`Destino [${g.email}]\nErro Final: ${e.message}`);
      g.linhas.forEach(cli => {
        const aba = ss.getSheetByName(cli.abaNome);
        if (aba) sinalizarErroEmail(aba, cli.linhaOriginal, "Falha Envio: " + e.message, dt);
      });
    }
  });

  if (errosCriticos.length > 0) throw new Error("\n" + errosCriticos.join("\n\n"));
  
  if (isTeste) {
     return `✅ LABORATÓRIO: Disparo Simulado Concluído!\nE-mails redirecionados com sucesso para ${EMAIL_TESTE_TRAVA}.`;
  }
  return `E-mail enviado com sucesso!`;
}

function marcarWhatsAgrupadoWeb(grupos, responsavel, isTeste) {
  const idPlanilhaParaUso = isTeste ? PLANILHA_TESTE_ID : PLANILHA_ID;
  const ss = SpreadsheetApp.openById(idPlanilhaParaUso);
  const dt = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
  
  grupos.forEach(g => {
    let ac = g.etapaNum === 1 ? "1_WHATS" : g.etapaNum === 2 ? "2_WHATS" : "3_WHATS";
    g.linhas.forEach(cli => {
      const aba = ss.getSheetByName(cli.abaNome);
      if (aba) {
        const nomeResponsavelReal = isTeste ? responsavel + " (Modo Teste)" : responsavel;
        aba.getRange(cli.linhaOriginal, MAPA_COLUNAS.CHECK_WHATS + 1).setValue(true);
        aba.getRange(cli.linhaOriginal, MAPA_COLUNAS.DATA_WHATS + 1).setValue(dt);
        aba.getRange(cli.linhaOriginal, MAPA_COLUNAS.RESPONSAVEL + 1).setValue(nomeResponsavelReal);
        registrarAuditoriaExata(cli.nome, cli.placa, cli.chassi, cli.email, cli.telefone, ac, dt, nomeResponsavelReal);
      }
    });
  });
  return "WhatsApp marcado na Planilha e Auditoria!";
}

function formatarComoEmail(textoHtmlOriginal, tituloEmail) {
  var textoHTML = textoHtmlOriginal.replace(/\n/g, '<br>');
  var htmlFinal = `
    <div style="font-family: Arial, sans-serif; font-size: 14px; color: #333333; max-width: 600px; margin: 0; line-height: 1.6;">
      <h3 style="color: #333333; margin-bottom: 20px; font-weight: bold; text-transform: uppercase;">${tituloEmail}</h3>
      <div style="margin-bottom: 20px;">${textoHTML}</div>
      <div style="border-top: 1px solid #dddddd; padding-top: 15px; margin-top: 20px; font-size: 13px; line-height: 1.5; color: #666666;">
        <img src="https://www.zensegurosbr.com/uploads/images/configuracoes/redimencionar-230-78-logo.png" width="160" alt="ZEN Seguros" style="display: block; margin-bottom: 8px; border: none; outline: none; text-decoration: none;">
        Atenciosamente,<br><strong style="color: #444444; font-size: 14px;">Setor de Rastreamento</strong><br>ZEN Seguros
      </div>
      <div style="display:none; color:transparent; font-size:1px;">Anti-Spam ID: ${new Date().getTime()}</div>
    </div>
  `;
  return htmlFinal;
}

// ====================================================================================
// FUNÇÃO DE RESET DO LABORATÓRIO (LIMPAR STATUS)
// ====================================================================================
function resetarMarcacoesTesteBackend(idsUnicos) {
  // Trava rígida: Esta função NUNCA deve ler PLANILHA_ID (Produção)
  const ssTeste = SpreadsheetApp.openById(PLANILHA_TESTE_ID);
  
  let cont = 0;
  const operacoes = {};
  
  idsUnicos.forEach(id => {
      const partes = id.lastIndexOf('-');
      const abaNome = id.substring(0, partes);
      const linha = parseInt(id.substring(partes + 1));
      if (!operacoes[abaNome]) operacoes[abaNome] = [];
      operacoes[abaNome].push(linha);
  });
  
  for (const abaNome in operacoes) {
      const aba = ssTeste.getSheetByName(abaNome);
      if (!aba) continue;
      
      operacoes[abaNome].forEach(linha => {
          aba.getRange(linha, MAPA_COLUNAS.CHECK_EMAIL + 1).setValue(false); 
          aba.getRange(linha, MAPA_COLUNAS.DATA_EMAIL + 1).setValue("");    
          aba.getRange(linha, MAPA_COLUNAS.RESPONDEU_EMAIL + 1).setValue(false); 
          aba.getRange(linha, MAPA_COLUNAS.CHECK_WHATS + 1).setValue(false); 
          aba.getRange(linha, MAPA_COLUNAS.DATA_WHATS + 1).setValue("");    
          aba.getRange(linha, MAPA_COLUNAS.RESPONDEU_WHATS + 1).setValue(false); 
          aba.getRange(linha, MAPA_COLUNAS.RESPONSAVEL + 1).setValue("");    
          cont++;
      });
  }
  return `O histórico de envios de ${cont} clientes foi zerado no Laboratório.`;
}