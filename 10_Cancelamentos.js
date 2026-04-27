// ====================================================================================
// SISTEMA DE CANCELAMENTOS E MIGRAÇÃO INTERNA (TUDO NA PLANILHA PRINCIPAL)
// ====================================================================================

const NOME_ABA_CANC = "🟣CancelamentoWeb";

// Mapeamento das colunas da aba 🟣CancelamentoWeb (0-indexado)
const COL_CANC = {
  NOME: 0, LOCAL: 1, PLACA: 2, CHASSI: 3, FIPE: 4, 
  EMAIL: 5, TELEFONE: 6, EMAIL_ENV: 7, DATA_EMAIL: 8, 
  WHATS_ENV: 9, DATA_WHATS: 10, RESP: 11,
  SITUACAO_REAL: 12, ETAPA: 13
};

function web_obterFilaCancelamento() {
  try {
    const ss = SpreadsheetApp.openById(PLANILHA_ID);
    const aba = ss.getSheetByName(NOME_ABA_CANC);
    if (!aba) return []; 

    const dados = aba.getDataRange().getValues();
    const notas = aba.getDataRange().getNotes();
    if (dados.length < 2) return [];

    const templatesDict = getTemplatesDict(ss);
    const fila = [];

    for (let i = 1; i < dados.length; i++) {
      const l = dados[i];
      const n = notas[i];
      const placa = String(l[COL_CANC.PLACA] || "").trim().toUpperCase();
      const chassi = String(l[COL_CANC.CHASSI] || "").trim().toUpperCase();
      const nome = String(l[COL_CANC.NOME] || "").trim().toUpperCase();
      const idVeic = placa || chassi;
      
      if (!placa && !chassi && !nome) continue;

      const localStr = String(l[COL_CANC.LOCAL] || "");
      const localNota = String(n[COL_CANC.LOCAL] || "");
      let cidade = "", bairro = "";
      if (localNota.includes("Cidade:")) {
         const parts = localNota.split("\n");
         cidade = parts[0] ? parts[0].replace("Cidade:", "").trim() : "";
         bairro = parts[1] ? parts[1].replace("Bairro:", "").trim() : "";
      }

      let telefone = String(l[COL_CANC.TELEFONE] || "").trim();
      let msgWhats = "";
      if (telefone) {
        let txtCorpo = aplicarTemplate(templatesDict, "CANC_WHATSAPP", nome || "Cliente", idVeic, false);
        let disclaimer = aplicarTemplate(templatesDict, "WHATSAPP_DISCLAIMER", nome || "Cliente", idVeic, false);
        msgWhats = (disclaimer && !disclaimer.includes("⚠️")) ? disclaimer + "\n\n" + txtCorpo : txtCorpo;
      }

      let etapaStr = String(l[COL_CANC.ETAPA] || "1");
      let numEtapa = parseInt(etapaStr.replace(/\D/g, '')) || 1;

      fila.push({
        idUnico: "CANC-" + (i + 1), linhaOriginal: i + 1, abaNome: NOME_ABA_CANC, 
        nome: nome, estado: localStr, cidade: cidade, bairro: bairro,
        placa: placa, chassi: chassi, fipe: String(l[COL_CANC.FIPE] || "").trim(), 
        email: String(l[COL_CANC.EMAIL] || "").trim().toLowerCase(),
        telefone: telefone, 
        isEnviado: (l[COL_CANC.EMAIL_ENV] === true || l[COL_CANC.EMAIL_ENV] === "TRUE" || l[COL_CANC.EMAIL_ENV] === 1),
        dataEmail: web_formatarDataSegura(l[COL_CANC.DATA_EMAIL]), 
        isWhatsEnviado: (l[COL_CANC.WHATS_ENV] === true || l[COL_CANC.WHATS_ENV] === "TRUE" || l[COL_CANC.WHATS_ENV] === 1),
        dataWhats: web_formatarDataSegura(l[COL_CANC.DATA_WHATS]), 
        responsavel: String(l[COL_CANC.RESP] || ""), 
        codSituacao: String(l[COL_CANC.SITUACAO_REAL] || "PENDENTE DE ANÁLISE"),
        etapaNum: numEtapa,
        mensagemWhatsApp: msgWhats
      });
    }
    return fila;
  } catch (e) {
    return { erro: "Falha Crítica ao ler aba de Cancelamentos: " + e.message };
  }
}

// ====================================================================================
// AUTOMAÇÃO LOTE: COMPLETAR DADOS SGA (FIPE E ENDEREÇO)
// ====================================================================================
function web_completarItemCancelamento(idUnico) {
  try {
    const linha = parseInt(idUnico.split('-')[1]);
    const ss = SpreadsheetApp.openById(PLANILHA_ID);
    const aba = ss.getSheetByName(NOME_ABA_CANC);
    if(!aba) return { status: 'erro', msg: 'Aba não encontrada' };
    
    const placa = String(aba.getRange(linha, COL_CANC.PLACA + 1).getValue()).trim().replace(/[^A-Z0-9]/gi, '');
    const chassi = String(aba.getRange(linha, COL_CANC.CHASSI + 1).getValue()).trim().toUpperCase();
    const veic = placa || chassi;
    const pb = placa ? "placa" : "chassi";
    
    if (!veic) return { status: 'erro', msg: 'Sem veículo para buscar' };
    
    const token = autenticarHINOVA();
    if (!token) return { status: 'erro', msg: 'Falha na autenticação' };
    
    const baseUrl = SGA_CONFIG.URL_CONSULTA_BASE.endsWith('/') ? SGA_CONFIG.URL_CONSULTA_BASE : SGA_CONFIG.URL_CONSULTA_BASE + '/';
    const resp = UrlFetchApp.fetch(`${baseUrl}${encodeURIComponent(veic)}/${pb}`, { headers: { "Authorization": "Bearer " + token }, muteHttpExceptions: true });
    
    if (resp.getResponseCode() === 200) {
       const j = JSON.parse(resp.getContentText());
       const arr = Array.isArray(j) ? j : [j];
       
       if (arr.length > 0 && arr[0]) {
           const d = arr[0];
           let alterado = false;
           
           // Processa FIPE
           if (d.valor_fipe) {
               aba.getRange(linha, COL_CANC.FIPE + 1).setValue(d.valor_fipe);
               alterado = true;
           }
           
           // Processa Endereço
           let valEstado = aba.getRange(linha, COL_CANC.LOCAL + 1).getValue();
           if (d.estado || d.cidade) {
               if (!valEstado && d.estado) aba.getRange(linha, COL_CANC.LOCAL + 1).setValue(String(d.estado).toUpperCase());
               if (d.cidade) {
                   const novaNota = `Cidade: ${d.cidade}\nBairro: ${d.bairro || 'N/A'}`;
                   aba.getRange(linha, COL_CANC.LOCAL + 1).setNote(novaNota);
               }
               alterado = true;
           }
           
           return alterado ? { status: 'ok', msg: 'Dados completados' } : { status: 'ok', msg: 'Nada novo a adicionar' };
       }
    }
    return { status: 'erro', msg: 'Veículo não localizado no SGA' };
  } catch(e) {
    return { status: 'erro', msg: e.message };
  }
}

// ====================================================================================
// MOTOR DE MIGRAÇÃO: LÊ ABA "6-SITUAÇÃO" -> COPIA PARA "🟣CancelamentoWeb"
// ====================================================================================
function web_migrarInativosParaCancelamento() {
  const ss = SpreadsheetApp.openById(PLANILHA_ID);
  
  const abaOrigem = ss.getSheets().find(s => s.getName().includes("6") && s.getName().toLowerCase().includes("situa"));
  if (!abaOrigem) return "❌ Erro: Aba '6-Situação' não encontrada na planilha principal.";
  
  let abaCanc = ss.getSheetByName(NOME_ABA_CANC);
  if (!abaCanc) {
    abaCanc = ss.insertSheet(NOME_ABA_CANC);
    abaCanc.appendRow(["Nome", "Local", "Placa", "Chassi", "Fipe", "E-mail", "Telefone", "E-mail Enviado", "Enviado e-mail em:", "Whatsapp", "Enviado whats em:", "Responsável", "Situação Real", "Etapa Original"]);
    abaCanc.getRange("A1:N1").setFontWeight("bold").setBackground("#f1f5f9");
    abaCanc.setFrozenRows(1);
  }

  const dadosOrigem = abaOrigem.getDataRange().getValues();
  if (dadosOrigem.length < 2) return "⚠️ A aba '6-Situação' está vazia. Não há clientes para copiar.";
  
  const dadosDestino = abaCanc.getDataRange().getValues();
  const placasExistentes = new Set();
  const chassisExistentes = new Set();
  
  for (let i = 1; i < dadosDestino.length; i++) {
    if (dadosDestino[i][COL_CANC.PLACA]) placasExistentes.add(String(dadosDestino[i][COL_CANC.PLACA]).trim().toUpperCase().replace(/[^A-Z0-9]/g, ''));
    if (dadosDestino[i][COL_CANC.CHASSI]) chassisExistentes.add(String(dadosDestino[i][COL_CANC.CHASSI]).trim().toUpperCase());
  }
  
  const headerOrigem = dadosOrigem[0].map(h => String(h).toLowerCase().trim());
  const getIdx = (keywords) => headerOrigem.findIndex(h => keywords.some(kw => h.includes(kw)));
  
  const idxNome = getIdx(["nome", "cliente"]);
  const idxPlaca = getIdx(["placa"]);
  const idxChassi = getIdx(["chassi"]);
  const idxFipe = getIdx(["fipe", "valor"]);
  const idxEmail = getIdx(["email", "e-mail"]);
  const idxTelefone = getIdx(["telefone", "celular", "whats", "contato"]);
  const idxLocal = getIdx(["estado", "uf", "local", "cidade"]);
  const idxSituacaoReal = getIdx(["situação real", "situacao real", "situação", "situacao", "status"]);
  
  // [SÊNIOR FIX]: Adicionado os termos para encontrar a coluna I (Aba de Origem) mapeada por você
  const idxEtapa = getIdx(["etapa", "aba de origem", "origem"]);
  
  const linhasAppend = [];
  let contCopiados = 0;
  let contIgnorados = 0;
  
  for (let i = 1; i < dadosOrigem.length; i++) {
    const row = dadosOrigem[i];
    
    let nome = idxNome >= 0 ? String(row[idxNome]).trim().toUpperCase() : String(row[MAPA_COLUNAS.NOME - 1] || "").trim().toUpperCase();
    let placa = idxPlaca >= 0 ? String(row[idxPlaca]).trim().toUpperCase() : String(row[MAPA_COLUNAS.PLACA - 1] || "").trim().toUpperCase();
    let chassi = idxChassi >= 0 ? String(row[idxChassi]).trim().toUpperCase() : String(row[MAPA_COLUNAS.CHASSI - 1] || "").trim().toUpperCase();
    let fipe = idxFipe >= 0 ? String(row[idxFipe]).trim() : String(row[MAPA_COLUNAS.FIPE - 1] || "").trim();
    let email = idxEmail >= 0 ? String(row[idxEmail]).trim().toLowerCase() : String(row[MAPA_COLUNAS.EMAIL - 1] || "").trim().toLowerCase();
    let telefone = idxTelefone >= 0 ? String(row[idxTelefone]).trim() : String(row[MAPA_COLUNAS.TELEFONE - 1] || "").trim();
    let local = idxLocal >= 0 ? String(row[idxLocal]).trim() : String(row[MAPA_COLUNAS.ESTADO - 1] || "").trim();
    let situacaoReal = idxSituacaoReal >= 0 ? String(row[idxSituacaoReal]).trim().toUpperCase() : "DESCONHECIDA";
    
    if (!placa && !chassi && !nome) continue;

    if (situacaoReal === "ATIVO" || situacaoReal === "ATIVO COM ADESIVO" || situacaoReal === "") continue;

    const placaLimpa = placa.replace(/[^A-Z0-9]/g, '');
    
    if ((chassi && chassisExistentes.has(chassi)) || (placaLimpa && placasExistentes.has(placaLimpa))) {
      contIgnorados++;
      continue;
    }
    
    // [SÊNIOR FIX]: Inteligência Regex. Lê "Aba de Origem" (Ex: "2 -Comunicação 5 Dias") e extrai apenas o dígito da etapa.
    let etapaReal = "1";
    if (idxEtapa >= 0) {
      let valOrigem = String(row[idxEtapa]).trim();
      let match = valOrigem.match(/(\d)/); // Captura o primeiro dígito que encontrar na string
      if (match) {
        etapaReal = match[1];
      }
    }
    
    linhasAppend.push([
      nome, local, placa, chassi, fipe, email, telefone, 
      false, "", false, "", "", situacaoReal, etapaReal
    ]);
    
    if (chassi) chassisExistentes.add(chassi);
    if (placaLimpa) placasExistentes.add(placaLimpa);
    
    contCopiados++;
  }
  
  if (linhasAppend.length > 0) {
    const startRow = abaCanc.getLastRow() + 1;
    abaCanc.getRange(startRow, 1, linhasAppend.length, 14).setValues(linhasAppend);
  }
  
  let msg = `✅ Cópia Concluída!`;
  if (contCopiados > 0) msg += `\n📥 ${contCopiados} clientes copiados da aba '6-Situação' para '🟣CancelamentoWeb'.`;
  if (contIgnorados > 0) msg += `\n⚠️ ${contIgnorados} já existiam em '🟣CancelamentoWeb' e foram ignorados para não duplicar.`;
  if (contCopiados === 0 && contIgnorados === 0) msg = `⚠️ Nenhum cliente inativo novo foi encontrado na aba '6-Situação'.`;
  
  return msg;
}

function web_marcarComoEnviadoCancelamento(clientesSelecionados, responsavel) {
  if (!clientesSelecionados || clientesSelecionados.length === 0) return "Vazio.";
  const ss = SpreadsheetApp.openById(PLANILHA_ID);
  const aba = ss.getSheetByName(NOME_ABA_CANC);
  const dt = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
  let cont = 0;
  
  clientesSelecionados.forEach(cli => {
    aba.getRange(cli.linhaOriginal, COL_CANC.EMAIL_ENV + 1).setValue(true);
    aba.getRange(cli.linhaOriginal, COL_CANC.DATA_EMAIL + 1).setValue(dt);
    aba.getRange(cli.linhaOriginal, COL_CANC.RESP + 1).setValue(responsavel);
    cont++;
  });
  return `✅ ${cont} e-mails marcados na aba de Cancelamentos!`;
}

function web_marcarWhatsCancelamentoWeb(grupos, responsavel) {
  const ss = SpreadsheetApp.openById(PLANILHA_ID);
  const aba = ss.getSheetByName(NOME_ABA_CANC);
  const dt = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
  
  grupos.forEach(grupo => {
    grupo.linhas.forEach(cli => {
      if (cli.idUnico && cli.idUnico.includes("CANC-")) {
        aba.getRange(cli.linhaOriginal, COL_CANC.WHATS_ENV + 1).setValue(true);
        aba.getRange(cli.linhaOriginal, COL_CANC.DATA_WHATS + 1).setValue(dt);
        aba.getRange(cli.linhaOriginal, COL_CANC.RESP + 1).setValue(responsavel);
      }
    });
  });
  return "✅ WhatsApp registrado no Cancelamento!";
}