// ====================================================================================
// 🧠 ARQUIVO: 09_Auth.gs (Gestão de Identidade, Sessões e Permissões)
// ====================================================================================

function web_validarLoginInterno(login, senha) {
  try {
    const ss = SpreadsheetApp.openById(PLANILHA_ID);
    const aba = ss.getSheetByName("🔐 Usuários");
    if (!aba) return { erro: "Aba '🔐 Usuários' não configurada no banco." };

    const dados = aba.getDataRange().getValues();
    const senhaBase64 = Utilities.base64Encode(senha, Utilities.Charset.UTF_8);
    const senhaPadraoBase64 = Utilities.base64Encode("ZEN0102*", Utilities.Charset.UTF_8);
    
    for (let i = 1; i < dados.length; i++) {
      const loginBanco = String(dados[i][1] || "").trim();
      const senhaBanco = String(dados[i][2] || "").trim();
      
      if (loginBanco === login && (senhaBanco === senhaBase64 || senhaBanco === senha)) {
        return {
          sucesso: true,
          login: login,
          nome: String(dados[i][0]).trim(),
          nivel: String(dados[i][3] || "OPERADOR").trim().toUpperCase(), // Lendo a Coluna D (Índice 3)
          isDefault: (senhaBanco === senhaPadraoBase64 || senhaBanco === "ZEN0102*")
        };
      }
    }
    
    return { erro: "Login ou senha incorretos." };
  } catch (e) {
    return { erro: "Erro de servidor: " + e.message };
  }
}

function web_alterarSenha(login, novaSenha) {
  try {
    const aba = SpreadsheetApp.openById(PLANILHA_ID).getSheetByName("🔐 Usuários");
    const dados = aba.getDataRange().getValues();
    
    for (let i = 1; i < dados.length; i++) {
      if (String(dados[i][1] || "").trim() === login) {
        aba.getRange(i + 1, 3).setValue(Utilities.base64Encode(novaSenha, Utilities.Charset.UTF_8));
        return { sucesso: true };
      }
    }
    return { erro: "Usuário não encontrado." };
  } catch (e) { return { erro: e.message }; }
}

function web_resetarSenha(linha) {
  try {
    SpreadsheetApp.openById(PLANILHA_ID).getSheetByName("🔐 Usuários").getRange(linha, 3).setValue(Utilities.base64Encode("ZEN0102*", Utilities.Charset.UTF_8));
    return "✅ Senha resetada para o padrão (ZEN0102*).";
  } catch (e) { return "❌ Erro ao resetar: " + e.message; }
}

function web_obterUsuarios() {
  const aba = SpreadsheetApp.openById(PLANILHA_ID).getSheetByName("🔐 Usuários");
  if (!aba) return [];
  const dados = aba.getDataRange().getValues();
  const usuarios = [];
  
  for (let i = 1; i < dados.length; i++) {
    if (dados[i][1]) {
      usuarios.push({
        linha: i + 1,
        nome: String(dados[i][0] || "").trim(),
        login: String(dados[i][1] || "").trim(),
        nivel: String(dados[i][3] || "OPERADOR").trim().toUpperCase() // Lendo a Coluna D (Índice 3)
      });
    }
  }
  return usuarios;
}

function web_adicionarUsuario(obj) {
  try {
    const aba = SpreadsheetApp.openById(PLANILHA_ID).getSheetByName("🔐 Usuários");
    const senhaPadraoBase64 = Utilities.base64Encode("ZEN0102*", Utilities.Charset.UTF_8);
    // Insere apenas as 4 colunas (A, B, C, D)
    aba.appendRow([ obj.nome, obj.login, senhaPadraoBase64, obj.nivel || "Operador" ]);
    return "✅ Usuário cadastrado! A senha padrão é ZEN0102*";
  } catch (e) { return "❌ Erro ao cadastrar: " + e.message; }
}

function web_removerUsuario(linha) {
  try {
    SpreadsheetApp.openById(PLANILHA_ID).getSheetByName("🔐 Usuários").deleteRow(linha);
    return "✅ Usuário removido do sistema!";
  } catch (e) { return "❌ Erro: " + e.message; }
}