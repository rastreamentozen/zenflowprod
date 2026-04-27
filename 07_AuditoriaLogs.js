// ====================================================================================
// AUDITORIA E LOGS (COM CENTRAL UNIFICADA E UX COMPACTA)
// ====================================================================================

function sinalizarErroEmail(aba, numeroLinha, motivo, dataHora) {
  var ss = SpreadsheetApp.openById(PLANILHA_ID);
  var abaErro = ss.getSheetByName("Erro") || ss.insertSheet("Erro");
  var dados = aba.getRange(numeroLinha, 1, 1, aba.getLastColumn()).getValues()[0];
  var emailAtual = dados[MAPA_COLUNAS.EMAIL] ? String(dados[MAPA_COLUNAS.EMAIL]).toLowerCase().trim() : "";
  
  if (emailAtual !== "") {
    var linhaParaErro = dados.slice(); 
    linhaParaErro.push("FALHA: " + motivo); 
    linhaParaErro.push(dataHora);
    abaErro.appendRow(linhaParaErro);
  }
  aba.getRange(numeroLinha, MAPA_COLUNAS.EMAIL + 1).setFontColor("#FF0000").setFontWeight("bold").setNote("⚠️ Erro: " + motivo);
}

function registrarAuditoriaExata(nome, placa, chassi, email, telefone, chaveAcao, dataHora, responsavel) {
  var abaAuditoria = SpreadsheetApp.openById(PLANILHA_ID).getSheetByName("4 -Registro - NÃO ALTERAR");
  if (!abaAuditoria) return;

  const MAPA_AUDITORIA = { 
    "1_EMAIL": { check: "1- E-mail (boas vindas)", data: "1- Enviado e-mail em:", resp: "1-Responsável (boas vindas)" }, 
    "1_WHATS": { check: "1-Whatsapp (boas vindas)", data: "1 -Enviado whats em:", resp: "1-Responsável (boas vindas)" }, 
    "2_EMAIL": { check: "2- E-mail (5 dias)", data: "2- Enviado e-mail em:", resp: "2-Responsável (5 dias)" }, 
    "2_WHATS": { check: "2-Whatsapp (5 dias)", data: "2 -Enviado whats em:", resp: "2-Responsável (5 dias)" }, 
    "3_EMAIL": { check: "3- E-mail (prazo)", data: "3- Enviado e-mail em:", resp: "3-Responsável (prazo)" }, 
    "3_WHATS": { check: "3-Whatsapp (prazo)", data: "3 -Enviado whats em:", resp: "3-Responsável (prazo)" } 
  };
  
  if (!MAPA_AUDITORIA[chaveAcao]) return;

  var cabecalho = abaAuditoria.getRange(1, 1, 1, abaAuditoria.getLastColumn()).getValues()[0].map(x => x ? String(x).trim() : "");
  var cChk = cabecalho.indexOf(MAPA_AUDITORIA[chaveAcao].check.trim()), 
      cDat = cabecalho.indexOf(MAPA_AUDITORIA[chaveAcao].data.trim()), 
      cRes = cabecalho.indexOf(MAPA_AUDITORIA[chaveAcao].resp.trim());
      
  var cNom = cabecalho.indexOf("Nome"), 
      cPla = cabecalho.indexOf("Placa"), 
      cCha = cabecalho.indexOf("Chassi"), 
      cEma = cabecalho.indexOf("E-mail") > -1 ? cabecalho.indexOf("E-mail") : cabecalho.indexOf("Email"), 
      cTel = cabecalho.indexOf("Telefone");
      
  var uLinha = abaAuditoria.getLastRow(), 
      dados = uLinha > 1 ? abaAuditoria.getRange(2, 1, uLinha - 1, abaAuditoria.getLastColumn()).getValues() : [];
  var lAlvo = -1;
  
  for (var i = 0; i < dados.length; i++) { 
    if ((placa && String(dados[i][cPla]).trim() === placa) || (chassi && String(dados[i][cCha]).trim() === chassi)) { 
      lAlvo = i + 2;
      break; 
    } 
  }

  if (lAlvo === -1) { 
    lAlvo = uLinha + 1;
    if (cNom > -1) abaAuditoria.getRange(lAlvo, cNom + 1).setValue(nome); 
    if (cPla > -1) abaAuditoria.getRange(lAlvo, cPla + 1).setValue(placa);
    if (cCha > -1) abaAuditoria.getRange(lAlvo, cCha + 1).setValue(chassi); 
    if (cEma > -1) abaAuditoria.getRange(lAlvo, cEma + 1).setValue(email);
    if (cTel > -1) abaAuditoria.getRange(lAlvo, cTel + 1).setValue(telefone); 
  }

  if (cChk > -1) abaAuditoria.getRange(lAlvo, cChk + 1).setValue(true);
  if (cDat > -1) abaAuditoria.getRange(lAlvo, cDat + 1).setValue(dataHora); 
  if (cRes > -1 && responsavel) abaAuditoria.getRange(lAlvo, cRes + 1).setValue(responsavel);
}

function conciliarErrosMailerDaemon() {
  const ss = SpreadsheetApp.openById(PLANILHA_ID); 
  const abaErro = ss.getSheetByName("Erro") || ss.insertSheet("Erro"); 
  let threads;
  
  try { 
    threads = GmailApp.search('(from:mailer-daemon OR from:postmaster) subject:("Delivery" OR "Failure" OR "Falha" OR "Undeliverable" OR "Returned" OR "Undelivered") newer_than:2d', 0, 50);
  } catch (e) { return; }
  
  if (threads.length === 0) return;
  const errosDet = {};
  
  threads.forEach(t => { 
    t.getMessages().forEach(m => { 
      const c = m.getPlainBody(); 
      const match = c.match(/([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9_-]+)/g); 
      if (match) { 
        match.forEach(em => { 
          let e = em.toLowerCase().trim(); 
          if (!e.includes("postmaster") && !e.includes("mailer-daemon")) errosDet[e] = "Falha na entrega"; 
        }); 
      } 
    }); 
  });

  const abas = ss.getSheets();
  for (let a = 0; a < abas.length; a++) {
    if (abas[a].getName().includes("1 -") || abas[a].getName().includes("2 -") || abas[a].getName().includes("3 -")) {
      const d = abas[a].getDataRange().getValues();
      for (let j = d.length - 1; j >= 1; j--) { 
        const em = d[j][MAPA_COLUNAS.EMAIL] ? String(d[j][MAPA_COLUNAS.EMAIL]).toLowerCase().trim() : "";
        if (em && errosDet[em]) { 
          abas[a].getRange(j + 1, MAPA_COLUNAS.EMAIL + 1).setFontColor("#FF0000").setFontWeight("bold").setNote("⚠️ Erro: " + errosDet[em]);
        } 
      }
    }
  }
}

// ------------------------------------------------------------------------------------
// MOTOR UNIFICADO DA CENTRAL DE LOGS (UX COMPACTA)
// ------------------------------------------------------------------------------------
function web_obterDadosLogsUnificado() {
  const ss = SpreadsheetApp.openById(PLANILHA_ID); 
  const a4 = ss.getSheetByName("4 -Registro - NÃO ALTERAR");
  if (!a4) return { colunas: [], linhas: [] };
  
  const dados = a4.getDataRange().getValues();
  if (dados.length < 2) return { colunas: [], linhas: [] };
  
  const aConc = ss.getSheetByName("Log Concluídos");
  const aSit = ss.getSheetByName("6 -Situação");
  const aErr = ss.getSheetByName("Erro");
  
  const sC = new Set();
  const mapI = new Map();
  const mapE = new Map();
  
  // Mapear Concluídos
  if (aConc) { 
    const d = aConc.getDataRange().getValues();
    for (let i = 1; i < d.length; i++) { 
      if (d[i][2]) sC.add(String(d[i][2]).trim().toUpperCase());
      if (d[i][3]) sC.add(String(d[i][3]).trim().toUpperCase());
    } 
  }
  
  // Mapear Situações SGA
  if (aSit) { 
    const d = aSit.getDataRange().getValues();
    for (let i = 1; i < d.length; i++) { 
      const cod = String(d[i][6]).trim();
      if (cod && cod !== "1" && cod !== "14") { 
        if (d[i][2]) mapI.set(String(d[i][2]).trim().toUpperCase(), cod);
        if (d[i][3]) mapI.set(String(d[i][3]).trim().toUpperCase(), cod);
      } 
    } 
  }
  
  // Mapear Erros
  if (aErr) { 
    const d = aErr.getDataRange().getValues();
    for (let i = 1; i < d.length; i++) { 
      const e = String(d[i][MAPA_COLUNAS.EMAIL]).trim().toLowerCase();
      let erroMsg = "Erro de entrega";
      for(let j = d[i].length - 1; j >= 0; j--) {
         if(String(d[i][j]).includes("FALHA:")) { erroMsg = String(d[i][j]).replace("FALHA:", "").trim(); break; }
      }
      if (e) mapE.set(e, erroMsg);
    } 
  }
  
  const cab = dados[0].map(c => String(c).trim());
  const iN = cab.indexOf("Nome"), iP = cab.indexOf("Placa"), iC = cab.indexOf("Chassi"), iE = cab.findIndex(c => c === "E-mail" || c === "Email");
  
  // Nova estrutura de colunas (Status ao lado do Cliente)
  const colunas = ["Etapa", "Dados do Cliente", "Status", "Identificação", "Histórico de Disparos", "Responsáveis"];
  const linhasMatriz = [];
  let inicio = dados.length > 1001 ? dados.length - 1000 : 1; 
  
  const mapEtapas = [ 
    { emj: "🔰", chkE: "1- E-mail (boas vindas)", datE: "1- Enviado e-mail em:", chkW: "1-Whatsapp (boas vindas)", datW: "1 -Enviado whats em:", resp: "1-Responsável (boas vindas)" }, 
    { emj: "⚠️", chkE: "2- E-mail (5 dias)", datE: "2- Enviado e-mail em:", chkW: "2-Whatsapp (5 dias)", datW: "2 -Enviado whats em:", resp: "2-Responsável (5 dias)" }, 
    { emj: "⛔", chkE: "3- E-mail (prazo)", datE: "3- Enviado e-mail em:", chkW: "3-Whatsapp (prazo)", datW: "3 -Enviado whats em:", resp: "3-Responsável (prazo)" } 
  ];
  
  const dataVarredura = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");

  for (let i = dados.length - 1; i >= inicio; i--) {
    const p = iP > -1 && dados[i][iP] ? String(dados[i][iP]).trim().toUpperCase() : "";
    const c = iC > -1 && dados[i][iC] ? String(dados[i][iC]).trim().toUpperCase() : "";
    const e = iE > -1 && dados[i][iE] ? String(dados[i][iE]).trim().toLowerCase() : "";
    const nome = iN > -1 ? dados[i][iN] : "";
    
    const ident = p ? p : (c ? c : "---");
    
    let envios = [], resps = new Set(), highestStage = 0;
    
    mapEtapas.forEach((etp, index) => {
      let idxCE = cab.indexOf(etp.chkE), idxDE = cab.indexOf(etp.datE), idxR = cab.indexOf(etp.resp), idxCW = cab.indexOf(etp.chkW), idxDW = cab.indexOf(etp.datW);
      let dE = idxDE > -1 && dados[i][idxDE] ? web_formatarDataSegura(dados[i][idxDE]).split(" ")[0] : "";
      let dW = idxDW > -1 && dados[i][idxDW] ? web_formatarDataSegura(dados[i][idxDW]).split(" ")[0] : "";
      let responsavel = idxR > -1 && dados[i][idxR] ? String(dados[i][idxR]).trim() : "";
      let hasEnvio = false;
      
      if ((idxCE > -1 && dados[i][idxCE] === true) || dE) { 
        envios.push(`<div class="mb-1 text-[12px] flex items-center gap-1.5 whitespace-nowrap"><span class="text-base">${etp.emj}</span> <span class="text-indigo-600 dark:text-indigo-400 font-bold bg-indigo-50 dark:bg-indigo-900/30 px-1.5 py-0.5 rounded border border-indigo-100 dark:border-indigo-800/50 shadow-sm">E-mail</span> <span class="text-slate-500 font-mono">${dE || '✔️'}</span></div>`); 
        if (responsavel && responsavel !== "Sistema") resps.add(`<div class="mb-1 text-[12px] flex items-center gap-1.5 whitespace-nowrap"><span class="text-base">${etp.emj}</span> <span class="font-black text-slate-700 dark:text-slate-300">${responsavel}</span></div>`);
        hasEnvio = true; 
      }
      if ((idxCW > -1 && dados[i][idxCW] === true) || dW) { 
        envios.push(`<div class="mb-1 text-[12px] flex items-center gap-1.5 whitespace-nowrap"><span class="text-base">${etp.emj}</span> <span class="text-emerald-600 dark:text-emerald-400 font-bold bg-emerald-50 dark:bg-emerald-900/30 px-1.5 py-0.5 rounded border border-emerald-100 dark:border-emerald-800/50 shadow-sm">Whats</span> <span class="text-slate-500 font-mono">${dW || '✔️'}</span></div>`);
        if (responsavel && responsavel !== "Sistema") resps.add(`<div class="mb-1 text-[12px] flex items-center gap-1.5 whitespace-nowrap"><span class="text-base">${etp.emj}</span> <span class="font-black text-slate-700 dark:text-slate-300">${responsavel}</span></div>`); 
        hasEnvio = true;
      }
      if (hasEnvio) highestStage = index + 1;
    });
    
    let isConcluido = (p && sC.has(p)) || (c && sC.has(c));
    let codSGA = (p && mapI.has(p)) ? mapI.get(p) : (c && mapI.has(c)) ? mapI.get(c) : null;
    let isInativo = codSGA !== null;
    let isErro = e && mapE.has(e);
    let erroMsg = isErro ? mapE.get(e) : "";
    let txtSGA = codSGA ? (MAPA_SITUACAO_SGA[codSGA] || "Inativo") : "";

    let htmlStatus = `<div class="flex items-center gap-1.5 w-max">`;
    let temStatus = false;
    
    if (isConcluido) { 
        htmlStatus += `<span class="text-lg cursor-help hover:scale-125 transition drop-shadow-sm" title="✅ STATUS: Concluído / Fechado&#10;🕒 Varredura: ${dataVarredura}">✅</span>`; 
        temStatus = true; 
    }
    if (isInativo) { 
        htmlStatus += `<span class="text-lg cursor-help hover:scale-125 transition drop-shadow-sm" title="🟣 SGA: ${txtSGA}&#10;🕒 Varredura: ${dataVarredura}">🟣</span>`; 
        temStatus = true; 
    }
    if (isErro) { 
        htmlStatus += `<span class="text-lg cursor-help hover:scale-125 transition drop-shadow-sm" title="❌ ERRO DE ENVIO: ${erroMsg}&#10;🕒 Varredura: ${dataVarredura}">❌</span>`; 
        temStatus = true; 
    }
    if (!temStatus) { 
        htmlStatus += `<span class="text-lg cursor-help hover:scale-125 transition opacity-50 drop-shadow-sm grayscale" title="⏳ STATUS: Ativo / Pendente de Ação&#10;🕒 Varredura: ${dataVarredura}">⏳</span>`; 
    }
    htmlStatus += `</div>`;

    let htmlEtapa = highestStage === 1 ? `<span class="text-lg cursor-help drop-shadow-sm" title="Etapa Atual: Boas Vindas">🔰</span>` :
                    highestStage === 2 ? `<span class="text-lg cursor-help drop-shadow-sm" title="Etapa Atual: Alerta 5 Dias">⚠️</span>` :
                    highestStage === 3 ? `<span class="text-lg cursor-help drop-shadow-sm" title="Etapa Atual: Prazo Expirado">⛔</span>` :
                    `<span class="text-lg cursor-help opacity-40 drop-shadow-sm" title="Nenhum disparo na esteira">➖</span>`;

    linhasMatriz.push({
      etapa: highestStage,
      isConcluido: isConcluido,
      isInativo: isInativo,
      isErro: isErro,
      htmlArr: [ 
        `<div class="text-center">${htmlEtapa}</div>`,
        `<div class="font-black text-slate-800 dark:text-white text-sm mb-0.5 whitespace-nowrap truncate max-w-[220px]" title="${nome}">${nome}</div><div class="text-[11px] text-slate-500 font-medium whitespace-nowrap truncate max-w-[220px]" title="${e}">${e}</div>`, 
        htmlStatus,
        `<div class="font-mono text-slate-600 dark:text-slate-400 font-bold bg-slate-100 dark:bg-slate-800 px-2.5 py-1.5 rounded w-max border border-slate-200 dark:border-slate-700 shadow-sm tracking-wide text-xs">${ident}</div>`, 
        envios.length > 0 ? envios.join("") : `<span class="text-slate-400 dark:text-slate-500 italic text-xs font-medium">Nenhum disparo registrado</span>`, 
        resps.size > 0 ? Array.from(resps).join("") : `<span class="text-slate-400 dark:text-slate-500 text-center block w-full">-</span>` 
      ]
    });
  }
  
  return JSON.parse(JSON.stringify({ colunas, linhas: linhasMatriz }));
}