// ====================================================================================
// MOTOR GRAMATICAL E SLA
// ====================================================================================
function aplicarTemplate(dict, chave, nomeCliente, identificadorVeiculo, isPlural, diasDecorridosSLA, limiteBaseSLA, dataEntradaStr) {
  let txt = dict[chave] || "⚠️ Erro: Template não encontrado.";

  // [SÊNIOR FIX]: Lógica Matemática do SLA dinâmico (Contagem Regressiva)
  let base = limiteBaseSLA !== undefined ? limiteBaseSLA : (chave.includes("5_DIAS") ? 5 : 10);
  let diasRestantes = base;

  if (diasDecorridosSLA !== undefined && diasDecorridosSLA !== null && !isNaN(diasDecorridosSLA)) {
     // Subtrai da base (10) e usa Math.max para impedir números negativos (ex: -2 dias)
     diasRestantes = Math.max(0, base - parseInt(diasDecorridosSLA));
  }

  let textoFinal = txt
    .replace(/{{NOME}}/g, nomeCliente)
    .replace(/{{VEICULO}}/g, identificadorVeiculo)
    .replace(/{{DIAS_RESTANTES}}/g, diasRestantes)
    .replace(/{{DATA_ENTRADA}}/g, dataEntradaStr || "Data não registrada");
  
  if (isPlural) {
    const mapaPlural = [
      [/do seu veículo/gi, "dos seus veículos"], [/o seu veículo/gi, "os seus veículos"],
      [/em seu veículo/gi, "em seus veículos"], [/seu veículo/gi, "seus veículos"],
      [/a instalação do rastreador/gi, "a instalação dos rastreadores"], [/do rastreador/gi, "dos rastreadores"],
      [/do equipamento/gi, "dos equipamentos"], [/o equipamento não for instalado/gi, "os equipamentos não forem instalados"],
      [/o rastreador ainda não/gi, "os rastreadores ainda não"], [/esteja instalado/gi, "estejam instalados"],
      [/um rastreador instalado/gi, "rastreadores instalados"], [/o veículo não contará/gi, "os veículos não contarão"],
      [/o veículo não estará/gi, "os veículos não estarão"], [/o veículo não se encontra/gi, "os veículos não se encontram"],
      [/permaneça assegurado/gi, "permaneçam assegurados"], [/o veículo/gi, "os veículos"]
    ];
    mapaPlural.forEach(p => textoFinal = textoFinal.replace(p[0], p[1]));
  }
  return textoFinal;
}

function obterFeriadosDoAno(ano) {
  const feriados = [];
  const calcularPascoa = (year) => {
    const a = year % 19, b = Math.floor(year / 100), c = year % 100;
    const d = Math.floor(b / 4), e = b % 4, f = Math.floor((b + 8) / 25);
    const g = Math.floor((b - f + 1) / 3), h = (19 * a + b - d - g + 15) % 30;
    const i = Math.floor(c / 4), k = c % 4, l = (32 + 2 * e + 2 * i - h - k) % 7;
    const m = Math.floor((a + 11 * h + 22 * l) / 451);
    const mes = Math.floor((h + l - 7 * m + 114) / 31);
    const dia = ((h + l - 7 * m + 114) % 31) + 1;
    return new Date(year, mes - 1, dia);
  };
  const pascoa = calcularPascoa(ano);
  const addDias = (data, dias) => { const nd = new Date(data.getTime()); nd.setDate(nd.getDate() + dias); return nd; };
  
  feriados.push(formatar(addDias(pascoa, -48)));
  feriados.push(formatar(addDias(pascoa, -47)));
  feriados.push(formatar(addDias(pascoa, -2)));
  feriados.push(formatar(addDias(pascoa, 60)));
  
  function formatar(d) { 
    const dd = String(d.getDate()).padStart(2, '0');
    const mm = String(d.getMonth() + 1).padStart(2, '0'); 
    return `${dd}/${mm}/${ano}`; 
  }
  
  const fixos = [ `01/01/${ano}`, `21/04/${ano}`, `23/04/${ano}`, `01/05/${ano}`, `24/06/${ano}`, `07/09/${ano}`, `12/10/${ano}`, `02/11/${ano}`, `15/11/${ano}`, `20/11/${ano}`, `22/11/${ano}`, `25/12/${ano}` ];
  return [...feriados, ...fixos];
}

function calcularDiasUteis(dataInicial, dataFinal, arrayFeriadosPersonalizadosTime) {
  let diasUteis = 0;
  let dataAtual = new Date(dataInicial.getTime());
  dataAtual.setHours(0, 0, 0, 0);
  let dFinal = new Date(dataFinal.getTime());
  dFinal.setHours(0, 0, 0, 0);
  const cacheFeriados = {};
  
  while (dataAtual < dFinal) {
    dataAtual.setDate(dataAtual.getDate() + 1);
    let diaSemana = dataAtual.getDay();
    if (diaSemana !== 0 && diaSemana !== 6) {
      const anoAtual = dataAtual.getFullYear();
      if (!cacheFeriados[anoAtual]) cacheFeriados[anoAtual] = obterFeriadosDoAno(anoAtual);
      const dd = String(dataAtual.getDate()).padStart(2, '0');
      const mm = String(dataAtual.getMonth() + 1).padStart(2, '0');
      const strDataAtual = `${dd}/${mm}/${anoAtual}`;
      
      let ehFeriado = cacheFeriados[anoAtual].includes(strDataAtual);
      if (!ehFeriado && arrayFeriadosPersonalizadosTime && arrayFeriadosPersonalizadosTime.includes(dataAtual.getTime())) ehFeriado = true;
      if (!ehFeriado) diasUteis++;
    }
  }
  return diasUteis;
}