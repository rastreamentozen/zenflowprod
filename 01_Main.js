// ====================================================================================
// 🧠 ARQUIVO: 01_Main.gs (Entrada do App e Funções Base)
// ====================================================================================

function Z_AUTORIZAR_SCRIPT() {
  const usuario = Session.getEffectiveUser().getEmail();
  MailApp.sendEmail({ to: usuario, subject: "SGCW - Autorização de BI e Slides", body: "Permissões de E-mail concedidas!" });
  try { SlidesApp.create("SGCW_Auth").setTrashed(true); } catch (e) { }
  console.log("✅ Permissões de envio de E-mail e Google Slides concedidas com sucesso.");
}

function doGet(e) {
  const t = HtmlService.createTemplateFromFile('Index');
  t.urlApp = ScriptApp.getService().getUrl();
  t.viewParam = e.parameter.view || '';
  return t.evaluate()
          .setTitle('SGCW - Portal Operacional')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
          .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// Utilitário para importar módulos HTML fatiados (Obrigatório para arquitetura limpa no GAS)
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function solicitarDadosWeb(tela, parametro) {
  try {
    if (tela === 'dashboard') return web_obterDadosDashboard();
    if (tela === 'filaGeral') return web_obterFilaGeral();
    if (tela === 'logs') return web_obterDadosLogsUnificado(); // Rota atualizada para a Central Unificada
    if (tela === 'config') return web_obterConfiguracoes();
    if (tela === 'tecnicos') return web_obterTecnicos();
    if (tela === 'statusAPI') return web_testarAPIs(); 
    return null;
  } catch (erro) { 
    return { erro: "Erro Backend: " + erro.message };
  }
}

function web_formatarDataSegura(v) { 
  return (v && v instanceof Date) ? Utilities.formatDate(v, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm") : String(v || "").trim();
}

function web_converterBooleano(v) { 
  return (v === true || v === "TRUE" || v === 1) ? "✔️" : (v === false || v === "FALSE" || v === 0) ? "❌" : String(v || "");
}

function web_extrairDataParaCalculo(valor) {
  if (!valor) return null;
  if (valor instanceof Date) return new Date(valor.getTime());
  var str = String(valor).trim().split(" ")[0];
  var p = str.split("/");
  if (p.length === 3) return new Date(p[2], p[1] - 1, p[0]);
  return null;
}

function getTemplatesDict(ss) {
  const aba = ss.getSheetByName("⚙️ Configurações");
  const dict = {};
  if (aba) {
    const dados = aba.getDataRange().getValues();
    for (let i = 1; i < dados.length; i++) {
      if (dados[i][0]) dict[String(dados[i][0]).trim()] = String(dados[i][1] || "");
    }
  }
  return dict;
}

function web_testarAPIs() {
  const status = { hinova: false, tempo: 0, erro: null };
  const inicio = new Date().getTime();
  
  try {
    const options = {
      "method": "post",
      "headers": { 
        "Authorization": "Bearer " + SGA_CONFIG.TOKEN_ASSOCIACAO, 
        "Content-Type": "application/json" 
      },
      "payload": JSON.stringify({ 
        "usuario": SGA_CONFIG.USUARIO, 
        "senha": SGA_CONFIG.SENHA 
      }),
      "muteHttpExceptions": true
    };
    
    const resp = UrlFetchApp.fetch(SGA_CONFIG.URL_AUTH, options);
    status.tempo = new Date().getTime() - inicio;
    if (resp.getResponseCode() === 200) {
      const json = JSON.parse(resp.getContentText());
      if (json.token_usuario) {
        status.hinova = true;
      } else {
        status.erro = "A API Hinova respondeu, mas o token de autenticação veio vazio.";
      }
    } else {
      status.erro = `Falha de Comunicação (HTTP ${resp.getResponseCode()})`;
    }
  } catch (e) {
    status.erro = "Erro interno do servidor Google: " + e.message;
  }
  
  return status;
}