// ====================================================================================
// AUTOMAÇÕES DE LOTE (REVISADAS E SINCRONIZADAS COM GOOGLE SHEETS)
// ====================================================================================

function processarItemLoteWeb(cli, comando) {
  const token = autenticarHINOVA();
  if (!token) return { status: 'erro', msg: 'Falha na autenticação Hinova.' };
  
  const ss = SpreadsheetApp.openById(PLANILHA_ID);
  const aba = ss.getSheetByName(cli.abaNome);
  if (!aba) return { status: 'erro', msg: 'Aba não encontrada.' };
  
  // Lógica corrigida: Prioridade estrita para PLACA
  const vb = String(cli.placa || cli.chassi).replace(/[^A-Za-z0-9]/g, '');
  const pb = cli.placa ? "placa" : "chassi";
  
  const linha = cli.linhaOriginal;
  const dt = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
  const codMoto = ["3", "126", "115", "100", "105", "127", "116", "32", "33", "34", "35", "95", "96", "97"];
  
  const optionsGet = { 
    "method": "get", 
    "headers": { "Authorization": "Bearer " + token, "Content-Type": "application/json" }, 
    "muteHttpExceptions": true 
  };

  // [SÊNIOR FIX]: Ferramentas Implacáveis de Extração de Endereço
  function buscarViaCep(cepStr) {
    if (!cepStr) return null;
    const c = String(cepStr).replace(/\D/g, '');
    if (c.length === 8) {
      try {
        const r = UrlFetchApp.fetch(`https://viacep.com.br/ws/${c}/json/`, { muteHttpExceptions: true });
        if (r.getResponseCode() === 200) {
          const j = JSON.parse(r.getContentText());
          if (!j.erro) return { estado: j.uf, cidade: j.localidade, bairro: j.bairro };
        }
      } catch(e) {}
    }
    return null;
  }

  function extrairEndereco(obj) {
    if (!obj || typeof obj !== 'object') return { cep: "", uf: "", cid: "", bai: "" };
    let cep = obj.cep || obj.cep_residencial || obj.cep_comercial || "";
    let uf = obj.estado || obj.uf || obj.estado_residencial || obj.uf_residencial || "";
    let cid = obj.cidade || obj.cidade_residencial || obj.localidade || "";
    let bai = obj.bairro || obj.bairro_residencial || "";
    
    if (Array.isArray(obj.enderecos) && obj.enderecos.length > 0) {
      let e = obj.enderecos[0];
      cep = cep || e.cep; uf = uf || e.estado || e.uf; cid = cid || e.cidade; bai = bai || e.bairro;
    }
    if (Array.isArray(obj.endereco) && obj.endereco.length > 0) {
      let e = obj.endereco[0];
      cep = cep || e.cep; uf = uf || e.estado || e.uf; cid = cid || e.cidade; bai = bai || e.bairro;
    }
    if (obj.endereco && typeof obj.endereco === 'object' && !Array.isArray(obj.endereco)) {
      let e = obj.endereco;
      cep = cep || e.cep; uf = uf || e.estado || e.uf; cid = cid || e.cidade; bai = bai || e.bairro;
    }
    return { cep: cep, uf: uf, cid: cid, bai: bai };
  }

  try {
    // 1. LÓGICA DE LOGÍSTICA
    if (comando === "logistica") {
      let est = String(aba.getRange(linha, MAPA_COLUNAS.ESTADO + 1).getValue()).trim().toUpperCase();
      let celulaEstado = aba.getRange(linha, MAPA_COLUNAS.ESTADO + 1);
      let notaEndereço = celulaEstado.getNote() || "";
      
      let cidadeCli = "", bairroCli = "";
      if (notaEndereço.includes("Cidade:")) {
        const cMatch = notaEndereço.match(/Cidade:\s*([^\n]*)/);
        if (cMatch) cidadeCli = cMatch[1].trim();
      }
      if (notaEndereço.includes("Bairro:")) {
        const bMatch = notaEndereço.match(/Bairro:\s*([^\n]*)/);
        if (bMatch) bairroCli = bMatch[1].trim();
      }

      if (!cidadeCli || !est || est === "N/A" || est === "") {
          const resLocal = processarItemLoteWeb(cli, "estados"); 
          if (resLocal.status === 'ok') {
            return processarItemLoteWeb(cli, "logistica"); 
          }
      }

      if (cidadeCli && est && est !== "N/A") {
        const ssTec = SpreadsheetApp.openById(ID_PLANILHA_TECNICOS);
        const listaTecnicos = ssTec.getSheets()[0].getDataRange().getValues().slice(1).filter(t => t[0]); 
        const enderecoDestino = `${bairroCli ? bairroCli + ', ' : ''}${cidadeCli} - ${est}, Brasil`;
        
        let melhorDistancia = Infinity, melhorTecnico = null, melhorTempoSeg = 0, melhorTipo = "Volante";
        let diagnosticoErros = [];

        let notaLimpa = notaEndereço.split("--- 🛰️ LOGÍSTICA ---")[0].trim();
        celulaEstado.setNote(notaLimpa + "\n\n--- 🛰️ LOGÍSTICA ---\n⏳ Calculando rotas...");
        SpreadsheetApp.flush(); 
        
        listaTecnicos.forEach(tecnico => {
          try {
            let partesEnd = [String(tecnico[1]||"").trim() + (tecnico[2] ? ", " + tecnico[2] : ""), String(tecnico[3]||"").trim(), String(tecnico[4]||"").trim() + " - " + String(tecnico[5]||"").trim(), String(tecnico[6]||"").trim(), "Brasil"];
            const origemCompleta = partesEnd.filter(p => p && p !== " - ").join(", ");
            const direcoes = Maps.newDirectionFinder().setOrigin(origemCompleta).setDestination(enderecoDestino).setMode(Maps.DirectionFinder.Mode.DRIVING).getDirections();

            if (direcoes && direcoes.routes && direcoes.routes.length > 0) {
              const rota = direcoes.routes[0].legs[0];
              if (rota.distance.value < melhorDistancia) {
                melhorDistancia = rota.distance.value; melhorTecnico = tecnico[0]; melhorTempoSeg = rota.duration.value; melhorTipo = tecnico[8] || "Volante";
              }
            }
          } catch (e) { diagnosticoErros.push(`[${tecnico[0]}: Erro]`); }
          Utilities.sleep(400); 
        });
        
        if (melhorTecnico) {
          const distKm = (melhorDistancia / 1000).toFixed(1);
          const h = Math.floor(melhorTempoSeg / 3600), m = Math.floor((melhorTempoSeg % 3600) / 60);
          celulaEstado.setNote(notaLimpa + "\n\n--- 🛰️ LOGÍSTICA ---\n" + `Atendimento: [${melhorTipo}] "${melhorTecnico}" - ${distKm} Km / ${String(h).padStart(2,'0')}:${String(m).padStart(2,'0')} de distância`);
        } else {
          celulaEstado.setNote(notaLimpa + "\n\n--- 🛰️ LOGÍSTICA ---\n⚠️ Rota inviável no estado.");
        }
        SpreadsheetApp.flush();
        return { status: 'ok', msg: 'Rota calculada' };
      }
      return { status: 'ignorado', msg: 'Endereço incompleto' };
    }

    // 2. BUSCA DE DADOS NA HINOVA
    const resp = UrlFetchApp.fetch(`${SGA_CONFIG.URL_CONSULTA_BASE}${vb}/${pb}`, optionsGet);
    
    if (resp.getResponseCode() !== 200) {
      aba.getRange(linha, MAPA_COLUNAS.NOME + 1).setNote(`⚠️ Falha na rede SGA.\nStatus HTTP: ${resp.getResponseCode()}\nTentado: ${pb.toUpperCase()} = ${vb}`);
      return { status: 'ignorado', msg: `Erro HTTP ${resp.getResponseCode()}` };
    }

    const jsonRaw = JSON.parse(resp.getContentText());
    const v = Array.isArray(jsonRaw) ? jsonRaw[0] : jsonRaw;
    
    if (!v || Object.keys(v).length === 0) {
      aba.getRange(linha, MAPA_COLUNAS.NOME + 1).setNote(`⚠️ Veículo não localizado na API da Hinova.\nTentado: ${pb.toUpperCase()} = ${vb}\n(Pode ser atraso de integração do SGA).`);
      return { status: 'ignorado', msg: 'Sem dados na Hinova' };
    }

    let alterado = false;

    if (comando === "estados") {
      let estFinal = "N/A", cidFinal = "Não informada", baiFinal = "Não informado";
      let endVeiculo = extrairEndereco(v);
      if (endVeiculo.cep) {
        let vCep = buscarViaCep(endVeiculo.cep);
        if (vCep) { endVeiculo.uf = vCep.estado; endVeiculo.cid = vCep.cidade; endVeiculo.bai = vCep.bairro; }
      }
      if (endVeiculo.uf && endVeiculo.cid) {
        estFinal = String(endVeiculo.uf).trim().toUpperCase();
        cidFinal = String(endVeiculo.cid).trim();
        baiFinal = String(endVeiculo.bai || "Não informado").trim();
      } else {
        let rotaBusca = "";
        if (v.codigo_associado && v.codigo_associado !== "0") rotaBusca = `${v.codigo_associado}/codigo`;
        else if (v.cpf_associado) rotaBusca = `${String(v.cpf_associado).replace(/\D/g, '')}/cpf`;
        else if (v.cnpj_associado) rotaBusca = `${String(v.cnpj_associado).replace(/\D/g, '')}/cnpj`;

        if (rotaBusca) {
          const rA = UrlFetchApp.fetch(`https://api.hinova.com.br/api/sga/v2/associado/buscar/${rotaBusca}`, optionsGet);
          if (rA.getResponseCode() === 200) {
            const jA = JSON.parse(rA.getContentText());
            const aD = Array.isArray(jA) ? jA[0] : jA;
            if (aD) {
              let endAssoc = extrairEndereco(aD);
              if (endAssoc.cep) {
                let aCep = buscarViaCep(endAssoc.cep);
                if (aCep) { endAssoc.uf = aCep.estado; endAssoc.cid = aCep.cidade; endAssoc.bai = aCep.bairro; }
              }
              if (endAssoc.uf) estFinal = String(endAssoc.uf).trim().toUpperCase();
              if (endAssoc.cid) cidFinal = String(endAssoc.cid).trim();
              if (endAssoc.bai) baiFinal = String(endAssoc.bai).trim();
            }
          }
        }
      }

      aba.getRange(linha, MAPA_COLUNAS.ESTADO + 1).setValue(estFinal).setNote(`📍 Cidade: ${cidFinal}\n🏘️ Bairro: ${baiFinal}`);
      if (estFinal !== "RJ" && estFinal !== "N/A") {
        aba.getRange(linha, MAPA_COLUNAS.TECNICO_INDISPONIVEL + 1).setValue(true);
      }
      alterado = true;
    }
    else if (comando === "fipe") {
      let valorFipeNum = 0;
      let strFipe = String(v.valor_fipe || "").trim();
      if (strFipe) {
        if (strFipe.indexOf(',') > -1 && strFipe.indexOf('.') > -1) strFipe = strFipe.replace(/\./g, '').replace(',', '.');
        else if (strFipe.indexOf(',') > -1) strFipe = strFipe.replace(',', '.');
        else if (strFipe.indexOf('.') > -1) { if (strFipe.split('.').pop().length === 3) strFipe = strFipe.replace(/\./g, ''); }
        valorFipeNum = parseFloat(strFipe.replace(/[^\d.-]/g, '')) || 0;
      }

      if (strFipe) aba.getRange(linha, MAPA_COLUNAS.FIPE + 1).setValue(v.valor_fipe);
      const isMoto = codMoto.indexOf(String(v.codigo_tipo_veiculo)) > -1;
      const isBaixa = (isMoto && valorFipeNum > 0 && valorFipeNum < 20000) || (!isMoto && valorFipeNum > 0 && valorFipeNum < 30000);
      
      aba.getRange(linha, MAPA_COLUNAS.FIPE_BAIXA + 1).setValue(isBaixa);
      alterado = true;
    }
    else if (comando === "motos") {
      if (codMoto.includes(String(v.codigo_tipo_veiculo))) {
        aba.getRange(linha, 1, 1, aba.getLastColumn()).setBackground("#d1fae5");
        aba.getRange(linha, MAPA_COLUNAS.PLACA + 1).setNote("🏍️ MOTO (SGA)");
        alterado = true;
      } else {
        aba.getRange(linha, 1, 1, aba.getLastColumn()).setBackground(null);
        aba.getRange(linha, MAPA_COLUNAS.PLACA + 1).clearNote();
      }
    } 
    else if (comando === "inativos") {
      const cSit = String(v.codigo_situacao), cClass = String(v.codigo_classificacao || ""), cNome = aba.getRange(linha, MAPA_COLUNAS.NOME + 1);
      if (cClass === "1") {
        cNome.setFontColor("#16a34a").setFontWeight("bold").setNote(`✅ CONCLUÍDO SGA em: ${dt}`);
      } else if (cSit !== "1" && cSit !== "14") {
        cNome.setFontColor("#9C27B0").setFontWeight("bold").setNote(`⚠️ Situação SGA: ${MAPA_SITUACAO_SGA[cSit] || "Desconhecida"}\nVerificado: ${dt}`);
      } else {
        cNome.setFontColor("#000000").setFontWeight("normal").clearNote();
      }
      alterado = true;
    }
    else if (comando === "completar") {
      if (!cli.nome && v.nome) { aba.getRange(linha, MAPA_COLUNAS.NOME + 1).setValue(String(v.nome).toUpperCase()); alterado = true; }
      if (!cli.email && v.email) { aba.getRange(linha, MAPA_COLUNAS.EMAIL + 1).setValue(String(v.email).toLowerCase()); alterado = true; }
      if (!cli.telefone) {
        let fT = v.ddd_celular && v.telefone_celular ? `(${v.ddd_celular}) ${v.telefone_celular}` : (v.telefone_celular || "");
        if (fT) { aba.getRange(linha, MAPA_COLUNAS.TELEFONE + 1).setValue(fT); alterado = true; }
      }
    }

    if (alterado) SpreadsheetApp.flush();
    return { status: 'ok', msg: 'Processado' };
    
  } catch (e) { 
    return { status: 'erro', msg: e.message }; 
  }
}

// ====================================================================================
// FUNÇÕES DE SUPORTE (AUXILIARES)
// ====================================================================================

function web_atualizarDadosCliente(abaNome, linha, dados) {
  try {
    const ss = SpreadsheetApp.openById(PLANILHA_ID), aba = ss.getSheetByName(abaNome);
    if (!aba) return "❌ Aba não encontrada.";
    if (dados.nome) aba.getRange(linha, MAPA_COLUNAS.NOME + 1).setValue(String(dados.nome).toUpperCase());
    if (dados.placa) aba.getRange(linha, MAPA_COLUNAS.PLACA + 1).setValue(String(dados.placa).toUpperCase());
    if (dados.chassi) aba.getRange(linha, MAPA_COLUNAS.CHASSI + 1).setValue(String(dados.chassi).toUpperCase());
    if (dados.email) aba.getRange(linha, MAPA_COLUNAS.EMAIL + 1).setValue(String(dados.email).toLowerCase());
    if (dados.telefone) aba.getRange(linha, MAPA_COLUNAS.TELEFONE + 1).setValue(String(dados.telefone));
    if (dados.fipe) aba.getRange(linha, MAPA_COLUNAS.FIPE + 1).setValue(String(dados.fipe));
    return "✅ Dados atualizados!";
  } catch (e) { return "❌ Erro: " + e.message; }
}

function executarFerramentaWebGlobal(comando) {
  try {
    if (comando === "sincronizar_erros") { conciliarErrosMailerDaemon(); return "✅ Erros sincronizados!"; }
    if (comando === "varrer_concluidos") { return varrerConcluidosGlobalWeb(); }
    return "⚠️ Comando inválido.";
  } catch (e) { return "❌ Erro: " + e.message; }
}

function autenticarHINOVA() {
  const cache = CacheService.getScriptCache(), tC = cache.get("HINOVA_TOKEN");
  if (tC) return tC;
  try {
    const opt = { "method": "post", "headers": { "Authorization": "Bearer " + SGA_CONFIG.TOKEN_ASSOCIACAO, "Content-Type": "application/json" }, "payload": JSON.stringify({ "usuario": SGA_CONFIG.USUARIO, "senha": SGA_CONFIG.SENHA }), "muteHttpExceptions": true };
    const r = UrlFetchApp.fetch(SGA_CONFIG.URL_AUTH, opt);
    const tok = JSON.parse(r.getContentText()).token_usuario || null;
    if (tok) cache.put("HINOVA_TOKEN", tok, 3000);
    return tok;
  } catch (e) { return null; }
}

function varrerConcluidosGlobalWeb() {
  const token = autenticarHINOVA(); if (!token) return "❌ Falha Login Hinova.";
  const ss = SpreadsheetApp.openById(PLANILHA_ID), dt = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
  let cConc = 0, logS = ss.getSheetByName("Log Concluídos") || ss.insertSheet("Log Concluídos");
  
  ["1 -", "2 -", "3 -"].forEach(f => {
    const aba = ss.getSheets().find(s => s.getName().includes(f)); if (!aba) return;
    const d = aba.getDataRange().getValues();
    for (let i = d.length - 1; i >= 1; i--) {
      // Lógica corrigida: Prioridade estrita para PLACA
      const vb = d[i][MAPA_COLUNAS.PLACA] || d[i][MAPA_COLUNAS.CHASSI]; if (!vb) continue;
      const pb = d[i][MAPA_COLUNAS.PLACA] ? "placa" : "chassi";
      try {
        const r = UrlFetchApp.fetch(`${SGA_CONFIG.URL_CONSULTA_BASE}${encodeURIComponent(vb)}/${pb}`, { "method": "get", "headers": { "Authorization": "Bearer " + token }, "muteHttpExceptions": true });
        const j = JSON.parse(r.getContentText()); const v = Array.isArray(j) ? j[0] : j;
        if (v && String(v.codigo_classificacao) === "1") {
          logS.appendRow([dt, d[i][MAPA_COLUNAS.NOME], d[i][MAPA_COLUNAS.PLACA], d[i][MAPA_COLUNAS.CHASSI], d[i][MAPA_COLUNAS.FIPE], d[i][MAPA_COLUNAS.EMAIL], d[i][MAPA_COLUNAS.TELEFONE], aba.getName()]);
          aba.deleteRow(i + 1); cConc++;
        }
      } catch (e) {}
    }
  });
  return `✅ Varredura Concluída! ${cConc} movidos para Log.`;
}

function varrerConcluidosSelecionadosWeb(sel) {
  if (!sel || sel.length === 0) return "Vazio.";
  const tok = autenticarHINOVA(); if (!tok) return "❌ Falha Login.";
  const ss = SpreadsheetApp.openById(PLANILHA_ID), dt = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
  let cC = 0, logS = ss.getSheetByName("Log Concluídos");
  
  const linhasParaDeletar = {};
  sel.forEach(cli => {
    // Lógica corrigida: Prioridade estrita para PLACA
    const vb = cli.placa || cli.chassi, pb = cli.placa ? "placa" : "chassi";
    try {
      const r = UrlFetchApp.fetch(`${SGA_CONFIG.URL_CONSULTA_BASE}${encodeURIComponent(vb)}/${pb}`, { "method": "get", "headers": { "Authorization": "Bearer " + tok }, "muteHttpExceptions": true });
      const j = JSON.parse(r.getContentText()); const v = Array.isArray(j) ? j[0] : j;
      if (v && String(v.codigo_classificacao) === "1") {
        logS.appendRow([dt, cli.nome, cli.placa, cli.chassi, cli.fipe, cli.email, cli.telefone, cli.abaNome]);
        if (!linhasParaDeletar[cli.abaNome]) linhasParaDeletar[cli.abaNome] = [];
        linhasParaDeletar[cli.abaNome].push(cli.linhaOriginal);
        cC++;
      }
    } catch(e) {}
  });

  for (const abaN in linhasParaDeletar) {
    const aba = ss.getSheetByName(abaN);
    linhasParaDeletar[abaN].sort((a,b) => b-a).forEach(l => aba.deleteRow(l));
  }
  return `✅ ${cC} concluídos removidos!`;
}

function atualizarMarcacaoWeb(abaN, l, c, v) {
  try {
    const aba = SpreadsheetApp.openById(PLANILHA_ID).getSheetByName(abaN);
    let col = c === 'fipeBaixa' ? MAPA_COLUNAS.FIPE_BAIXA+1 : c === 'tecnicoIndisp' ? MAPA_COLUNAS.TECNICO_INDISPONIVEL+1 : c === 'respEmail' ? MAPA_COLUNAS.RESPONDEU_EMAIL+1 : c === 'respWhats' ? MAPA_COLUNAS.RESPONDEU_WHATS+1 : 0;
    if (col) { aba.getRange(l, col).setValue(v); return "✅ Salvo!"; }
  } catch(e) { return "❌ Erro: " + e.message; }
}

function marcarComoEnviadoWeb(sel, resp) {
  const ss = SpreadsheetApp.openById(PLANILHA_ID), dt = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
  sel.forEach(cli => {
    const aba = ss.getSheetByName(cli.abaNome);
    aba.getRange(cli.linhaOriginal, MAPA_COLUNAS.CHECK_EMAIL+1).setValue(true);
    aba.getRange(cli.linhaOriginal, MAPA_COLUNAS.DATA_EMAIL+1).setValue(dt);
    aba.getRange(cli.linhaOriginal, MAPA_COLUNAS.RESPONSAVEL+1).setValue(resp);
    registrarAuditoriaExata(cli.nome, cli.placa, cli.chassi, cli.email, cli.telefone, cli.etapaNum === 1 ? "1_EMAIL" : cli.etapaNum === 2 ? "2_EMAIL" : "3_EMAIL", dt, resp);
  });
  return `✅ ${sel.length} marcados!`;
}

function web_obterConfiguracoes() {
  const aba = SpreadsheetApp.openById(PLANILHA_ID).getSheetByName("⚙️ Configurações");
  return aba.getDataRange().getValues().slice(1).map((r,i) => ({ linhaOriginal: i+2, chave: String(r[0]), texto: String(r[1]||"") }));
}

function salvarConfiguracaoWeb(l, t) {
  try { SpreadsheetApp.openById(PLANILHA_ID).getSheetByName("⚙️ Configurações").getRange(l, 2).setValue(t); return "✅ Atualizado!"; }
  catch(e) { return "❌ Erro: " + e.message; }
}