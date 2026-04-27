// ====================================================================================
// MOTOR DE IMPORTAÇÃO E FILA GERAL - OMNI-SCAN (BLINDADO)
// ====================================================================================

function cadastrarLoteWeb(loteDeClientes) {
  try {
    const ss = SpreadsheetApp.openById(PLANILHA_ID);
    const aba1 = ss.getSheets().find(s => s.getName().includes("1 -"));
    
    if (!aba1) return "❌ Erro: Aba da Etapa 1 não encontrada.";

    const chassisNoSistema = new Set();
    const placasNoSistema = new Set();
    const nomesNoSistema = new Set(); // Fallback para clientes sem veículo
    
    // Regex tolerante para absorver placas sujas do SGA (com espaços, traços, mercosul e antiga)
    const regexPlaca = /^[A-Z]{3}[\s-]?[0-9][A-Z0-9][0-9]{2}$|^[A-Z]{3}[\s-]?[0-9]{4}$/i;
    const regexChassi = /^[A-HJ-NPR-Z0-9]{17}$/i;

    // 1. OMNI-SCAN: Varredura 2D em 100% da Planilha
    ss.getSheets().forEach(aba => {
        const nomeAba = aba.getName().toLowerCase();
        if (nomeAba === "feriados" || nomeAba === "dashboard") return;

        const ultimaLinha = aba.getLastRow();
        const ultimaColuna = aba.getLastColumn();
        if (ultimaLinha < 2 || ultimaColuna < 1) return;

        const dados = aba.getRange(1, 1, ultimaLinha, ultimaColuna).getValues();

        for (let r = 1; r < dados.length; r++) { // Inicia na linha 1 ignorando o cabeçalho
            let linhaTemVeiculo = false;
            
            // Abordagem A: Busca por Padrão (Independe de qual coluna o usuário colocou o dado)
            for (let c = 0; c < dados[r].length; c++) {
                let val = String(dados[r][c]).trim().toUpperCase();
                if (!val) continue;

                let limpo = val.replace(/[^A-Z0-9]/g, '');

                if ((limpo.length >= 7 && limpo.length <= 8) && regexPlaca.test(val)) {
                    placasNoSistema.add(limpo);
                    linhaTemVeiculo = true;
                } 
                else if (val.length === 17 && regexChassi.test(val)) {
                    chassisNoSistema.add(val);
                    linhaTemVeiculo = true;
                }
            }

            // Abordagem B: Busca Mapeada (Garante o bloqueio caso o regex falhe bizarramente mas a coluna esteja correta)
            if (nomeAba.includes("1") || nomeAba.includes("2") || nomeAba.includes("3") || nomeAba.includes("4") || nomeAba.includes("6") || nomeAba.includes("auditoria") || nomeAba.includes("cancelado")) {
                let possivelNome = String(dados[r][MAPA_COLUNAS.NOME - 1] || "").trim().toUpperCase();
                let possivelPlaca = String(dados[r][MAPA_COLUNAS.PLACA - 1] || "").replace(/[^A-Z0-9]/g, '');
                let possivelChassi = String(dados[r][MAPA_COLUNAS.CHASSI - 1] || "").trim().toUpperCase();

                if (possivelPlaca && possivelPlaca.length >= 6) { placasNoSistema.add(possivelPlaca); linhaTemVeiculo = true; }
                if (possivelChassi && possivelChassi.length >= 6) { chassisNoSistema.add(possivelChassi); linhaTemVeiculo = true; }

                // Fallback: Se não detectamos nenhum veículo na linha inteira, travamos o cadastro pelo NOME para evitar fantasmas
                if (!linhaTemVeiculo && possivelNome) {
                    nomesNoSistema.add(possivelNome);
                }
            }
        }
    });

    const qtdColunasParaInserir = Math.max(aba1.getLastColumn(), 20) - 1;
    const dtHoje = new Date();
    const dtHojeStr = Utilities.formatDate(dtHoje, Session.getScriptTimeZone(), "dd/MM/yyyy");
    
    let contInseridos = 0, contDuplicados = 0;
    const lotesPorAba = [];
    const notasPorAba = [];
    
    loteDeClientes.forEach((cliente) => {
      const chassiCli = String(cliente.chassi || "").trim().toUpperCase();
      const placaCli = String(cliente.placa || "").trim().toUpperCase().replace(/[^A-Z0-9]/g, '');
      const nomeCli = String(cliente.nome || "").trim().toUpperCase();
      
      let isDuplicate = false;
      
      if (chassiCli && chassisNoSistema.has(chassiCli)) {
          isDuplicate = true;
      } else if (placaCli && placasNoSistema.has(placaCli)) {
          isDuplicate = true;
      } else if (!chassiCli && !placaCli && nomeCli && nomesNoSistema.has(nomeCli)) {
          isDuplicate = true;
      }

      if (isDuplicate) { 
        contDuplicados++; 
        return; 
      }

      const novaLinha = new Array(qtdColunasParaInserir).fill("");
      const novaNota = new Array(qtdColunasParaInserir).fill("");
      
      novaLinha[0] = cliente.data || dtHojeStr;
      novaLinha[MAPA_COLUNAS.NOME - 1] = nomeCli;
      novaLinha[MAPA_COLUNAS.PLACA - 1] = placaCli;
      novaLinha[MAPA_COLUNAS.CHASSI - 1] = chassiCli;
      novaLinha[MAPA_COLUNAS.FIPE - 1] = String(cliente.fipe || "").trim();
      novaLinha[MAPA_COLUNAS.EMAIL - 1] = String(cliente.email || "").trim().toLowerCase();
      novaLinha[MAPA_COLUNAS.TELEFONE - 1] = String(cliente.telefone || "").trim();
      
      if (cliente.estado) {
         novaLinha[MAPA_COLUNAS.ESTADO - 1] = String(cliente.estado).trim().toUpperCase();
         let cid = cliente.cidade ? String(cliente.cidade).trim() : "Não informada";
         let bai = cliente.bairro ? String(cliente.bairro).trim() : "Não informado";
         novaNota[MAPA_COLUNAS.ESTADO - 1] = `📍 Cidade: ${cid}\n🏘️ Bairro: ${bai}`;
      }
      
      lotesPorAba.push(novaLinha);
      notasPorAba.push(novaNota);
      contInseridos++;
      
      // Realimentação em Tempo Real (Evita que o usuário copie a mesma linha no SGA e cole 2x no lote atual)
      if (chassiCli) chassisNoSistema.add(chassiCli);
      if (placaCli) placasNoSistema.add(placaCli);
      if (!chassiCli && !placaCli && nomeCli) nomesNoSistema.add(nomeCli);
    });
    
    // Inserção Otimizada em Batch
    if (lotesPorAba.length > 0) {
        const nomes = aba1.getRange("C1:C").getValues();
        let ultimaLinhaReal = 1;
        for (let j = nomes.length - 1; j >= 0; j--) {
          if (String(nomes[j][0]).trim() !== "") { ultimaLinhaReal = j + 1; break; }
        }
        const range = aba1.getRange(ultimaLinhaReal + 1, 2, lotesPorAba.length, qtdColunasParaInserir);
        range.setValues(lotesPorAba);
        range.setNotes(notasPorAba); 
    }

    let msg = `✅ Lote Processado com Sucesso!\n📥 ${contInseridos} roteados para a Etapa 1.`;
    if (contDuplicados > 0) msg += `\n⚠️ ${contDuplicados} clientes ignorados (já existiam nas outras Etapas ou Históricos).`;
    return msg;
  } catch (e) { 
    return "❌ Erro Crítico no Motor de Lote: " + e.message;
  }
}

function web_obterFilaGeral() {
  const ss = SpreadsheetApp.openById(PLANILHA_ID);
  const abas = ss.getSheets().filter(s => s.getName().includes("1 -") || s.getName().includes("2 -") || s.getName().includes("3 -"));
  const fila = [];
  const templatesDict = getTemplatesDict(ss);

  let feriadosTime = [];
  try {
    const abaFeriados = ss.getSheetByName("Feriados");
    if (abaFeriados) {
      feriadosTime = abaFeriados.getRange("A2:A").getValues().map(r => r[0] instanceof Date ? r[0].getTime() : null).filter(r => r);
    }
  } catch(e) {}
  
  const mapHist = {};
  try {
    const abasAud = ss.getSheets().filter(s => s.getName().includes("4 -") || s.getName().toLowerCase().includes("auditoria"));
    
    abasAud.forEach(abaAud => {
      const dadosAud = abaAud.getDataRange().getValues();
      if (dadosAud.length < 2) return;
      
      const headAud = dadosAud[0];
      let idxPlaca = headAud.findIndex(c => String(c).toLowerCase().includes("placa"));
      let idxChassi = headAud.findIndex(c => String(c).toLowerCase().includes("chassi"));
      let idxDataE1 = headAud.findIndex(c => String(c).includes("1- Enviado e-mail"));
      let idxDataW1 = headAud.findIndex(c => String(c).includes("1 -Enviado whats"));
      let idxDataE2 = headAud.findIndex(c => String(c).includes("2- Enviado e-mail"));
      let idxDataW2 = headAud.findIndex(c => String(c).includes("2 -Enviado whats"));
      let idxDataE3 = headAud.findIndex(c => String(c).includes("3- Enviado e-mail"));
      let idxDataW3 = headAud.findIndex(c => String(c).includes("3 -Enviado whats"));

      for (let i = 1; i < dadosAud.length; i++) {
        let l = dadosAud[i];
        let placaStr = idxPlaca > -1 ? String(l[idxPlaca]).toUpperCase().replace(/[^A-Z0-9]/g, '') : "";
        let chassiStr = idxChassi > -1 ? String(l[idxChassi]).toUpperCase().trim() : "";
        
        let key = placaStr || chassiStr;
        if (!key) continue;

        if (!mapHist[key]) mapHist[key] = { e1: "", e2: "", e3: "" };

        let dE1 = idxDataE1 > -1 ? web_formatarDataSegura(l[idxDataE1]) : "";
        let dW1 = idxDataW1 > -1 ? web_formatarDataSegura(l[idxDataW1]) : "";
        if (dE1 && dE1 !== "Aguardando...") mapHist[key].e1 = dE1;
        else if (dW1 && dW1 !== "Aguardando...") mapHist[key].e1 = dW1;

        let dE2 = idxDataE2 > -1 ? web_formatarDataSegura(l[idxDataE2]) : "";
        let dW2 = idxDataW2 > -1 ? web_formatarDataSegura(l[idxDataW2]) : "";
        if (dE2 && dE2 !== "Aguardando...") mapHist[key].e2 = dE2;
        else if (dW2 && dW2 !== "Aguardando...") mapHist[key].e2 = dW2;

        let dE3 = idxDataE3 > -1 ? web_formatarDataSegura(l[idxDataE3]) : "";
        let dW3 = idxDataW3 > -1 ? web_formatarDataSegura(l[idxDataW3]) : "";
        if (dE3 && dE3 !== "Aguardando...") mapHist[key].e3 = dE3;
        else if (dW3 && dW3 !== "Aguardando...") mapHist[key].e3 = dW3;
      }
    });
  } catch(e) { console.log("Erro ao mapear auditoria: " + e.message); }

  const dtHoje = new Date();

  abas.forEach(aba => {
    const nomeAba = aba.getName();
    let numEtapa = nomeAba.includes("2 -") ? 2 : nomeAba.includes("3 -") ? 3 : 1;
    const ultimaLinha = aba.getLastRow();
    const ultimaColuna = aba.getLastColumn();
    if (ultimaLinha < 2 || ultimaColuna < 1) return;

    const range = aba.getRange(1, 1, ultimaLinha, ultimaColuna);
    const dados = range.getValues();
    const notas = range.getNotes();

    for (let i = 1; i < dados.length; i++) {
      const l = dados[i];
      const nome = l[MAPA_COLUNAS.NOME] ? String(l[MAPA_COLUNAS.NOME]).trim() : "";
      const placa = l[MAPA_COLUNAS.PLACA] ? String(l[MAPA_COLUNAS.PLACA]).trim() : "";
      const chassi = l[MAPA_COLUNAS.CHASSI] ? String(l[MAPA_COLUNAS.CHASSI]).trim() : "";
      
      if (!placa && !chassi && !nome) continue;

      const notaNome = (MAPA_COLUNAS.NOME < ultimaColuna && notas[i][MAPA_COLUNAS.NOME]) ? String(notas[i][MAPA_COLUNAS.NOME]) : "";
      const notaPlaca = (MAPA_COLUNAS.PLACA < ultimaColuna && notas[i][MAPA_COLUNAS.PLACA]) ? String(notas[i][MAPA_COLUNAS.PLACA]).toUpperCase() : "";
      const notaEmail = (MAPA_COLUNAS.EMAIL < ultimaColuna && notas[i][MAPA_COLUNAS.EMAIL]) ? String(notas[i][MAPA_COLUNAS.EMAIL]) : "";
      const notaEstado = (MAPA_COLUNAS.ESTADO < ultimaColuna && notas[i][MAPA_COLUNAS.ESTADO]) ? String(notas[i][MAPA_COLUNAS.ESTADO]) : "";

      let cidade = "", bairro = "";
      let tecnicoDisp = "", tecnicoDist = "", tecnicoTempo = "", tecnicoTipo = "Volante";
      
      if (notaEstado.includes("Cidade:")) {
        const parts = notaEstado.split("\n");
        cidade = parts[0] ? parts[0].replace("📍 Cidade:", "").trim() : "";
        bairro = parts[1] ? parts[1].replace("🏘️ Bairro:", "").trim() : "";
      }

      if (notaEstado.includes("🛰️ LOGÍSTICA")) {
        let logMatchNovo = notaEstado.match(/Atendimento: \[(.*?)\] "(.*?)" - (.*?) \/ (.*?) de distância/);
        if (logMatchNovo) {
          tecnicoTipo = logMatchNovo[1]; tecnicoDisp = logMatchNovo[2]; tecnicoDist = logMatchNovo[3]; tecnicoTempo = logMatchNovo[4];
        } else {
          let logMatch = notaEstado.match(/Técnico Disponível: "(.*?)" - (.*?) \/ (.*?) de distância/);
          if (logMatch) { tecnicoDisp = logMatch[1]; tecnicoDist = logMatch[2]; tecnicoTempo = logMatch[3]; }
        }
      }

      let dPlanilha = null;
      let dataEntradaStr = "Data não registrada";
      if (l[MAPA_COLUNAS.DATA] instanceof Date) {
         dPlanilha = l[MAPA_COLUNAS.DATA];
         dataEntradaStr = Utilities.formatDate(dPlanilha, Session.getScriptTimeZone(), "dd/MM/yyyy");
      }

      let dEmail = null;
      const valDataEmail = l[MAPA_COLUNAS.DATA_EMAIL];
      if (valDataEmail instanceof Date) dEmail = valDataEmail;

      let diasUteisParaUI = 0; 
      if (dPlanilha && !isNaN(dPlanilha)) {
         try { diasUteisParaUI = calcularDiasUteis(dPlanilha, dtHoje, feriadosTime); } catch(e) {}
      }

      const telefone = l[MAPA_COLUNAS.TELEFONE] ? String(l[MAPA_COLUNAS.TELEFONE]).trim() : "";
      let msgWhats = "";
      if (telefone) {
        const idVeic = placa || chassi;
        let chaveCorpo = numEtapa === 1 ? (l[MAPA_COLUNAS.FIPE_BAIXA] === true ? "BOAS_VINDAS_FIPE_BAIXA" : "BOAS_VINDAS_NORMAL") : numEtapa === 2 ? "LEMBRETE_5_DIAS" : "PRAZO_EXPIRADO";
        msgWhats = aplicarTemplate(templatesDict, chaveCorpo, nome || "Cliente", idVeic, false, diasUteisParaUI, 10, dataEntradaStr);
      }

      const keyBusca = placa.replace(/[^A-Z0-9]/g, '') || chassi.toUpperCase();
      const hist = mapHist[keyBusca] ? { ...mapHist[keyBusca] } : { e1: "", e2: "", e3: "" };

      fila.push({
        idUnico: nomeAba + "-" + (i + 1), 
        etapaNum: numEtapa, linhaOriginal: i + 1, abaNome: nomeAba, nome: nome, placa: placa, chassi: chassi,
        fipe: l[MAPA_COLUNAS.FIPE] ? String(l[MAPA_COLUNAS.FIPE]).trim() : "",
        email: l[MAPA_COLUNAS.EMAIL] ? String(l[MAPA_COLUNAS.EMAIL]).trim() : "",
        telefone: telefone, estado: l[MAPA_COLUNAS.ESTADO] ? String(l[MAPA_COLUNAS.ESTADO]).trim() : "",
        cidade: cidade, bairro: bairro,
        tecnicoDisp: tecnicoDisp, tecnicoDist: tecnicoDist, tecnicoTempo: tecnicoTempo, tecnicoTipo: tecnicoTipo, 
        dataPlanilha: dataEntradaStr,
        dataEmail: web_formatarDataSegura(l[MAPA_COLUNAS.DATA_EMAIL]),
        dataWhats: web_formatarDataSegura(l[MAPA_COLUNAS.DATA_WHATS]),
        isEnviado: (l[MAPA_COLUNAS.CHECK_EMAIL] === true),
        isWhatsEnviado: (l[MAPA_COLUNAS.CHECK_WHATS] === true),
        isRespondeuEmail: (l[MAPA_COLUNAS.RESPONDEU_EMAIL] === true),
        isRespondeuWhats: (l[MAPA_COLUNAS.RESPONDEU_WHATS] === true),
        isFipeBaixa: (l[MAPA_COLUNAS.FIPE_BAIXA] === true),
        isTecnicoIndisp: (l[MAPA_COLUNAS.TECNICO_INDISPONIVEL] === true),
        isMoto: notaPlaca.includes("MOTO"),
        isInativo: notaNome.includes("Situação SGA"),
        isErroEmail: notaEmail.includes("Erro:"),
        notaNome: notaNome, notaEmail: notaEmail, mensagemWhatsApp: msgWhats,
        diasUteisSLA: diasUteisParaUI, etapaSugerida: numEtapa,
        histE1: hist.e1, histE2: hist.e2, histE3: hist.e3 
      });
    }
  });
  return fila;
}

function web_migrarClientesEtapa(movimentacoes) {
  try {
    const ss = SpreadsheetApp.openById(PLANILHA_ID);
    const abasInfo = {
      1: ss.getSheets().find(s => s.getName().includes("1 -")),
      2: ss.getSheets().find(s => s.getName().includes("2 -")),
      3: ss.getSheets().find(s => s.getName().includes("3 -"))
    };

    const operacoesOrigem = {}; let sucesso = 0;

    movimentacoes.forEach(mov => {
      const partes = mov.idUnico.lastIndexOf('-');
      const abaNome = mov.idUnico.substring(0, partes);
      const linha = parseInt(mov.idUnico.substring(partes + 1));
      if (!operacoesOrigem[abaNome]) operacoesOrigem[abaNome] = [];
      const abaOrigem = ss.getSheetByName(abaNome);
      if (abaOrigem) {
        const r = abaOrigem.getRange(linha, 1, 1, abaOrigem.getLastColumn());
        operacoesOrigem[abaNome].push({ linha: linha, novaEtapa: mov.novaEtapa, rowData: r.getValues()[0], rowNotes: r.getNotes()[0] });
      }
    });

    for (const [abaNome, tarefas] of Object.entries(operacoesOrigem)) {
      const abaOrigem = ss.getSheetByName(abaNome);
      tarefas.sort((a, b) => b.linha - a.linha).forEach(t => {
        const abaDestino = abasInfo[t.novaEtapa];
        if (abaDestino) {
          const last = abaDestino.getLastRow() + 1;
          abaDestino.getRange(last, 1, 1, t.rowData.length).setValues([t.rowData]).setNotes([t.rowNotes]);
          abaOrigem.deleteRow(t.linha);
          sucesso++;
        }
      });
    }
    return `✅ Migração efetuada! ${sucesso} clientes movidos.`;
  } catch(e) { return "❌ Erro na migração: " + e.message; }
}

function gatilho_migracaoAutomaticaSLA() {
  const ss = SpreadsheetApp.openById(PLANILHA_ID);
  const abasInfo = {
    1: ss.getSheets().find(s => s.getName().includes("1 -")),
    2: ss.getSheets().find(s => s.getName().includes("2 -")),
    3: ss.getSheets().find(s => s.getName().includes("3 -"))
  };
  
  let feriadosTime = [];
  try {
    const abaFeriados = ss.getSheetByName("Feriados");
    if (abaFeriados) {
      feriadosTime = abaFeriados.getRange("A2:A").getValues().map(r => r[0] instanceof Date ? r[0].getTime() : null).filter(r => r);
    }
  } catch(e) {}

  const dtHoje = new Date();
  const movimentacoes = [];

  [1, 2].forEach(numEtapa => {
    const aba = abasInfo[numEtapa];
    if (!aba) return;
    const dados = aba.getDataRange().getValues();
    
    for (let i = 1; i < dados.length; i++) {
      const l = dados[i];
      const placa = String(l[MAPA_COLUNAS.PLACA] || "").trim();
      const chassi = String(l[MAPA_COLUNAS.CHASSI] || "").trim();
      if (!placa && !chassi) continue;

      let dBaseSLA = null;

      if (numEtapa === 1) {
         if (l[MAPA_COLUNAS.DATA] instanceof Date) {
            dBaseSLA = l[MAPA_COLUNAS.DATA];
         } else {
            const strData = String(l[MAPA_COLUNAS.DATA] || "").split(" ")[0];
            if (strData && strData.includes("/")) {
                const partes = strData.split("/");
                if (partes.length === 3) dBaseSLA = new Date(partes[2], partes[1] - 1, partes[0]);
            }
         }
      } else if (numEtapa === 2) {
         const valDataEmail = l[MAPA_COLUNAS.DATA_EMAIL];
         if (valDataEmail instanceof Date) {
            dBaseSLA = valDataEmail;
         } else {
            const strData = String(valDataEmail || "").split(" ")[0];
            if (strData && strData !== "Aguardando..." && strData.includes("/")) {
                const partes = strData.split("/");
                if (partes.length === 3) dBaseSLA = new Date(partes[2], partes[1] - 1, partes[0]);
            }
         }
      }

      if (dBaseSLA && !isNaN(dBaseSLA)) {
         try {
             const diasUteis = calcularDiasUteis(dBaseSLA, dtHoje, feriadosTime);
             if (diasUteis >= 5) {
                 const idUnico = aba.getName() + "-" + (i + 1);
                 movimentacoes.push({ idUnico: idUnico, novaEtapa: numEtapa + 1 });
             }
         } catch(e) {}
      }
    }
  });

  if (movimentacoes.length > 0) {
    web_migrarClientesEtapa(movimentacoes);
    return `✅ Migração Automática de SLA concluída! ${movimentacoes.length} clientes movidos de etapa devido à expiração dos 5 dias úteis.`;
  }
  return "✅ Varredura concluída. Nenhum cliente com SLA expirado hoje.";
}