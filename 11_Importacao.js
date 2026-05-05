/**
 * FASE 1: Compara os dados contra as 3 abas de Comunicação E a aba de Registro (Double Check Global).
 * @param {Array} dadosImportados - JSON padronizado do frontend.
 */
function validarLoteImportacao(dadosImportados) {
  if (!dadosImportados || !Array.isArray(dadosImportados) || dadosImportados.length === 0) {
    return { sucesso: false, mensagem: "Nenhum dado válido fornecido para validação." };
  }

  const ss = SpreadsheetApp.openById(PLANILHA_ID);

  try {
    const normalizar = (str) => {
      if (!str) return "";
      return str.toString().toLowerCase().trim().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
    };

    const placasExistentes = new Set();
    const chassisExistentes = new Set();
    
    // Mapeamento Cirúrgico: Abas 1, 2, 3 (Placa na D, Chassi na E) e Aba 4 (Placa na C, Chassi na D)
    const abasParaValidar = [
      { nome: "1 - Comunicação de Boas Vindas", idxPlaca: 3, idxChassi: 4 }, 
      { nome: "2 -Comunicação 5 Dias", idxPlaca: 3, idxChassi: 4 },          
      { nome: "3 - Prazo Expirado", idxPlaca: 3, idxChassi: 4 },             
      { nome: "4 -Registro - NÃO ALTERAR", idxPlaca: 2, idxChassi: 3 }       
    ];

    // Varredura Omnisciente de Segurança
    abasParaValidar.forEach(abaConfig => {
      const sheet = ss.getSheetByName(abaConfig.nome);
      if (sheet) {
        const values = sheet.getDataRange().getValues();
        if (values.length > 1) { 
          for (let i = 1; i < values.length; i++) {
            if (values[i][abaConfig.idxPlaca]) placasExistentes.add(values[i][abaConfig.idxPlaca].toString().toUpperCase().trim());
            if (values[i][abaConfig.idxChassi]) chassisExistentes.add(values[i][abaConfig.idxChassi].toString().toUpperCase().trim());
          }
        }
      }
    });

    const validos = [];
    const duplicados = [];

    // Triagem do payload colado
    dadosImportados.forEach(item => {
      const valorPlaca = item.placa ? item.placa.toString().toUpperCase().trim() : "";
      const valorChassi = item.chassi ? item.chassi.toString().toUpperCase().trim() : "";

      let isDuplicado = false;
      if (valorPlaca && placasExistentes.has(valorPlaca)) isDuplicado = true;
      if (valorChassi && chassisExistentes.has(valorChassi)) isDuplicado = true;

      if (isDuplicado) {
        duplicados.push(item);
      } else {
        validos.push(item);
        // Trava para evitar que ele importe o mesmo chassi/placa duas vezes no mesmo CTRL+V
        if (valorPlaca) placasExistentes.add(valorPlaca);
        if (valorChassi) chassisExistentes.add(valorChassi);
      }
    });

    return { sucesso: true, validos: validos, duplicados: duplicados };

  } catch (error) {
    return { sucesso: false, mensagem: `Erro na validação interna: ${error.message}` };
  }
}

/**
 * FASE 2: Efetiva a inserção na Etapa 1 APÓS a aprovação humana.
 * Resolve a matriz exata da coluna B até a H, e o Estado/Comentário na coluna R.
 * @param {Array} dadosValidos - Apenas os JSONs que passaram na validação do frontend.
 */
function efetivarImportacaoBanco(dadosValidos) {
  if (!dadosValidos || dadosValidos.length === 0) return { sucesso: false, mensagem: "Lote vazio aprovado." };

  const ss = SpreadsheetApp.openById(PLANILHA_ID);
  const abaDestino = "1 - Comunicação de Boas Vindas"; 
  const sheetDestino = ss.getSheetByName(abaDestino);

  if (!sheetDestino) return { sucesso: false, mensagem: `A aba '${abaDestino}' não foi encontrada.` };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); 
  } catch (e) {
    return { sucesso: false, mensagem: "O banco de dados está ocupado no momento. Tente novamente em 10 segundos." };
  }

  try {
    const linhasParaInserir = [];
    const notasParaInserir = []; 

    dadosValidos.forEach(item => {
      // Cria um array de 18 posições (0 a 17, onde 0 é a Coluna A e 17 é a Coluna R)
      const row = new Array(18).fill("");
      const rowNotes = new Array(18).fill("");

      row[1]  = item.dataPlanilha || ""; // B: Data
      row[2]  = item.nome || "";         // C: Nome
      row[3]  = item.placa || "";        // D: Placa
      row[4]  = item.chassi || "";       // E: Chassi
      row[5]  = item.fipe || "";         // F: Fipe
      row[6]  = item.email || "";        // G: Email
      row[7]  = item.telefone || "";     // H: Telefone
      row[17] = item.estado || "";       // R: Estado

      // Injeta Cidade e Bairro como comentário invisível na célula de Estado
      let notaR = "";
      if (item.cidade) notaR += `Cidade: ${item.cidade}`;
      if (item.bairro) notaR += (notaR ? '\n' : '') + `Bairro: ${item.bairro}`;
      rowNotes[17] = notaR; 

      linhasParaInserir.push(row);
      notasParaInserir.push(rowNotes);
    });

    if (linhasParaInserir.length > 0) {
      const startRow = sheetDestino.getLastRow() + 1;
      const range = sheetDestino.getRange(startRow, 1, linhasParaInserir.length, 18);
      range.setValues(linhasParaInserir);
      range.setNotes(notasParaInserir); 
    }

    return { sucesso: true, mensagem: `${linhasParaInserir.length} clientes inseridos com sucesso na Etapa 1!` };
  } catch (error) {
    return { sucesso: false, mensagem: `Erro no banco: ${error.message}` };
  } finally {
    lock.releaseLock();
  }
}