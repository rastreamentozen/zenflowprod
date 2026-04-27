// ====================================================================================
// GESTÃO DE TÉCNICOS (CRUD PLANILHA TÉCNICOS LENDO ATÉ COLUNA I)
// ====================================================================================
function web_obterTecnicos() {
  try {
    const ss = SpreadsheetApp.openById(ID_PLANILHA_TECNICOS);
    const aba = ss.getSheets()[0];
    const dados = aba.getDataRange().getValues();
    const tecnicos = [];
    
    for (let i = 1; i < dados.length; i++) {
      if (dados[i][0]) {
        tecnicos.push({
          linha: i + 1,
          nome: String(dados[i][0] || "").trim(),
          endereco: String(dados[i][1] || "").trim(),
          numero: String(dados[i][2] || "").trim(),
          bairro: String(dados[i][3] || "").trim(),
          cidade: String(dados[i][4] || "").trim(),
          estado: String(dados[i][5] || "").trim(),
          cep: String(dados[i][6] || "").trim(),
          telefone: String(dados[i][7] || "").trim(),
          tipo: String(dados[i][8] || "Volante").trim() // Lê a coluna I (Índice 8)
        });
      }
    }
    return tecnicos;
  } catch (e) { return { erro: "Erro ao ler a planilha de técnicos: " + e.message }; }
}

function web_adicionarTecnico(dadosObj) {
  try {
    const ss = SpreadsheetApp.openById(ID_PLANILHA_TECNICOS);
    const aba = ss.getSheets()[0];
    // Grava as 9 colunas
    aba.appendRow([ dadosObj.nome, dadosObj.endereco, dadosObj.numero, dadosObj.bairro, dadosObj.cidade, dadosObj.estado, dadosObj.cep, dadosObj.telefone, dadosObj.tipo || 'Volante' ]);
    return "✅ Técnico cadastrado com sucesso!";
  } catch (e) { return "❌ Erro ao salvar técnico: " + e.message; }
}

function web_atualizarTecnico(linha, dadosObj) {
  try {
    const ss = SpreadsheetApp.openById(ID_PLANILHA_TECNICOS);
    const aba = ss.getSheets()[0];
    // Atualiza as 9 colunas
    aba.getRange(linha, 1, 1, 9).setValues([[ dadosObj.nome, dadosObj.endereco, dadosObj.numero, dadosObj.bairro, dadosObj.cidade, dadosObj.estado, dadosObj.cep, dadosObj.telefone, dadosObj.tipo || 'Volante' ]]);
    return "✅ Dados do técnico atualizados com sucesso!";
  } catch (e) { return "❌ Erro ao atualizar técnico: " + e.message; }
}

function web_removerTecnico(linha) {
  try {
    const ss = SpreadsheetApp.openById(ID_PLANILHA_TECNICOS);
    const aba = ss.getSheets()[0];
    aba.deleteRow(linha);
    return "✅ Técnico removido com sucesso!";
  } catch (e) { return "❌ Erro ao remover técnico: " + e.message; }
}