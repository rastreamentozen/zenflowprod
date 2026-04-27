// ====================================================================================
// OBTENÇÃO DE PRÉVIA E DISPARO CANAL A CANAL
// ====================================================================================
function obterPreviewDisparoAgrupadoWeb(grupos) {
  const ss = SpreadsheetApp.openById(PLANILHA_ID);
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
            // [SÊNIOR FIX]: Fallback - Se não houver envio prévio de 5 dias, mantém o padrão fixo de 5 dias restantes (10 base - 5 passados)
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

    const htmlBodyFormatado = formatarComoEmail(txt, tituloHeader);
    
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

function dispararEmailAgrupadoWeb(grupos, responsavel) {
  const ss = SpreadsheetApp.openById(PLANILHA_ID);
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
            // [SÊNIOR FIX]: Fallback para o Motor Efetivo também
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
      const htmlPadrao = formatarComoEmail(txt, tituloHeader);
      
      const assuntoFinal = g.customAssunto ? g.customAssunto : ass;
      const htmlFinal = g.customEmailHtml ? g.customEmailHtml : htmlPadrao;
      const textoPuroFallback = txt + "\n\nAtenciosamente,\nSetor de Rastreamento\nZEN Seguros";

      MailApp.sendEmail({
        to: g.email, 
        subject: assuntoFinal,
        body: textoPuroFallback,
        htmlBody: htmlFinal, 
        name: "Setor de Rastreamento - ZEN Seguros"
      });
      
      g.linhas.forEach(cli => {
        const aba = ss.getSheetByName(cli.abaNome);
        if (aba) {
          aba.getRange(cli.linhaOriginal, MAPA_COLUNAS.CHECK_EMAIL + 1).setValue(true);
          aba.getRange(cli.linhaOriginal, MAPA_COLUNAS.DATA_EMAIL + 1).setValue(dt);
          aba.getRange(cli.linhaOriginal, MAPA_COLUNAS.RESPONSAVEL + 1).setValue(responsavel);
          registrarAuditoriaExata(cli.nome, cli.placa, cli.chassi, cli.email, cli.telefone, ac, dt, responsavel);
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
  return `E-mail enviado com sucesso!`;
}

function marcarWhatsAgrupadoWeb(grupos, responsavel) {
  const ss = SpreadsheetApp.openById(PLANILHA_ID);
  const dt = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
  
  grupos.forEach(g => {
    let ac = g.etapaNum === 1 ? "1_WHATS" : g.etapaNum === 2 ? "2_WHATS" : "3_WHATS";
    g.linhas.forEach(cli => {
      const aba = ss.getSheetByName(cli.abaNome);
      if (aba) {
        aba.getRange(cli.linhaOriginal, MAPA_COLUNAS.CHECK_WHATS + 1).setValue(true);
        aba.getRange(cli.linhaOriginal, MAPA_COLUNAS.DATA_WHATS + 1).setValue(dt);
        aba.getRange(cli.linhaOriginal, MAPA_COLUNAS.RESPONSAVEL + 1).setValue(responsavel);
        registrarAuditoriaExata(cli.nome, cli.placa, cli.chassi, cli.email, cli.telefone, ac, dt, responsavel);
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