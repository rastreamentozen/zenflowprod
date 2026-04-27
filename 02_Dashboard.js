// ====================================================================================
// NOVO MOTOR DE DASHBOARD (BI COMPLETO) E SLIDES
// ====================================================================================
function web_obterDadosDashboard() {
  const ss = SpreadsheetApp.openById(PLANILHA_ID);
  const hoje = new Date();
  hoje.setHours(0, 0, 0, 0);
  const mesAtual = hoje.getMonth();
  const anoAtual = hoje.getFullYear();
  const stats = {
    kpis: {
      pendentesTotal: 0, pendentesEtapa: { '1': 0, '2': 0, '3': 0 },
      conclusaoTotal: 0, conclusaoEtapa: { '1': 0, '2': 0, '3': 0 },
      emails: { hoje: 0, semana: 0, mes: 0, total: 0, etapa: { '1': 0, '2': 0, '3': 0 } },
      whats: { hoje: 0, semana: 0, mes: 0, total: 0, etapa: { '1': 0, '2': 0, '3': 0 } }
    },
    graficos: { 
      historicoMensal: {}, equipeMes: {} 
    }
  };

  ["1 -", "2 -", "3 -"].forEach(nomeFrag => {
    const aba = ss.getSheets().find(s => s.getName().includes(nomeFrag));
    if (aba) {
      const etapaStr = nomeFrag.includes("1") ? '1' : nomeFrag.includes("2") ? '2' : '3';
      const d = aba.getDataRange().getValues();
      for (let i = 1; i < d.length; i++) {
        if (!d[i][MAPA_COLUNAS.PLACA] && !d[i][MAPA_COLUNAS.CHASSI]) continue;
        const eE = d[i][MAPA_COLUNAS.CHECK_EMAIL] === true || d[i][MAPA_COLUNAS.CHECK_EMAIL] === "TRUE" || d[i][MAPA_COLUNAS.CHECK_EMAIL] === 1;
        const eW = d[i][MAPA_COLUNAS.CHECK_WHATS] === true || d[i][MAPA_COLUNAS.CHECK_WHATS] === "TRUE" || d[i][MAPA_COLUNAS.CHECK_WHATS] === 1;
        if (!eE && !eW) { stats.kpis.pendentesTotal++; stats.kpis.pendentesEtapa[etapaStr]++; }
      }
    }
  });

  const aud = ss.getSheetByName("4 -Registro - NÃO ALTERAR");
  if (aud) {
    const dAud = aud.getDataRange().getValues();
    if (dAud.length > 0) {
      const cab = dAud[0].map(c => String(c).trim());
      const mapasEnvio = [
        { canal: 'emails', etapa: '1', idxDat: cab.findIndex(c => c.includes("1- Enviado e-mail")), idxResp: cab.findIndex(c => c.includes("1-Responsável")) },
        { canal: 'whats', etapa: '1', idxDat: cab.findIndex(c => c.includes("1 -Enviado whats")), idxResp: -1 },
        { canal: 'emails', etapa: '2', idxDat: cab.findIndex(c => c.includes("2- Enviado e-mail")), idxResp: cab.findIndex(c => c.includes("2-Responsável")) },
        { canal: 'whats', etapa: '2', idxDat: cab.findIndex(c => c.includes("2 -Enviado whats")), idxResp: -1 },
        { canal: 'emails', etapa: '3', idxDat: cab.findIndex(c => c.includes("3- Enviado e-mail")), idxResp: cab.findIndex(c => c.includes("3-Responsável")) },
        { canal: 'whats', etapa: '3', idxDat: cab.findIndex(c => c.includes("3 -Enviado whats")), idxResp: -1 }
      ];

      for (let i = 1; i < dAud.length; i++) {
        mapasEnvio.forEach(m => {
          if (m.idxDat > -1 && dAud[i][m.idxDat]) {
            const dataCalc = web_extrairDataParaCalculo(dAud[i][m.idxDat]);
            if (dataCalc) {
              stats.kpis[m.canal].total++;
              stats.kpis[m.canal].etapa[m.etapa]++;
        
              const diffDias = Math.floor((hoje.getTime() - dataCalc.getTime()) / 86400000);
              if (diffDias === 0) stats.kpis[m.canal].hoje++;
              if (diffDias >= 0 && diffDias <= 7) stats.kpis[m.canal].semana++;

              if (dataCalc.getMonth() === mesAtual && dataCalc.getFullYear() === anoAtual) {
                stats.kpis[m.canal].mes++;
         
                if (m.canal === 'emails' && m.idxResp > -1) {
                  const resp = String(dAud[i][m.idxResp]).trim();
                  if (resp && resp !== "Sistema" && !resp.includes("Gatilho")) {
                    stats.graficos.equipeMes[resp] = (stats.graficos.equipeMes[resp] || 0) + 1;
                  }
                }
              }

              const chaveMes = ("0" + (dataCalc.getMonth() + 1)).slice(-2) + "/" + dataCalc.getFullYear();
              if (!stats.graficos.historicoMensal[chaveMes]) {
                stats.graficos.historicoMensal[chaveMes] = { email: 0, whats: 0, sortKey: (dataCalc.getFullYear() * 100) + dataCalc.getMonth() };
              }
              if (m.canal === 'emails') stats.graficos.historicoMensal[chaveMes].email++;
              else stats.graficos.historicoMensal[chaveMes].whats++;
            }
          }
        });
      }
    }
  }

  const conc = ss.getSheetByName("Log Concluídos");
  if (conc) {
    const dConc = conc.getDataRange().getValues();
    for (let i = 1; i < dConc.length; i++) {
      stats.kpis.conclusaoTotal++;
      const abaOrigem = String(dConc[i][7] || "");
      if (abaOrigem.includes("1 -")) stats.kpis.conclusaoEtapa['1']++;
      else if (abaOrigem.includes("2 -")) stats.kpis.conclusaoEtapa['2']++;
      else if (abaOrigem.includes("3 -")) stats.kpis.conclusaoEtapa['3']++;
    }
  }
  return JSON.parse(JSON.stringify(stats));
}

function exportarDashboardParaSlidesWeb(graficos, statsObj) {
  try {
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
    const presentation = SlidesApp.create("Relatório Executivo SGCW - " + timestamp);
    const slideTitulo = presentation.getSlides()[0];

    slideTitulo.insertTextBox("Relatório Executivo - BI\nSGCW OPERACIONAL", 50, 150, 600, 100).getText().getTextStyle().setFontSize(32).setBold(true).setForegroundColor("#4f46e5");
    slideTitulo.insertTextBox("Gerado automaticamente em: " + timestamp, 50, 260, 600, 40).getText().getTextStyle().setFontSize(14).setForegroundColor("#64748b");
    
    if (statsObj && statsObj.kpis) {
      const slideKpi = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
      slideKpi.insertTextBox("Resumo Operacional", 50, 30, 620, 50).getText().getTextStyle().setFontSize(22).setBold(true).setForegroundColor("#1e293b");
      
      let textoKpi = "📊 NÚMEROS GERAIS DA OPERAÇÃO:\n\n";
      textoKpi += "• Veículos Pendentes de Ação: " + statsObj.kpis.pendentesTotal + "\n";
      textoKpi += "   (Etapa 1: " + statsObj.kpis.pendentesEtapa['1'] + " | Etapa 2: " + statsObj.kpis.pendentesEtapa['2'] + " | Etapa 3: " + statsObj.kpis.pendentesEtapa['3'] + ")\n\n";
      textoKpi += "• Total de E-mails Enviados: " + statsObj.kpis.emails.total + "\n";
      textoKpi += "   (Hoje: " + statsObj.kpis.emails.hoje + " | Semana: " + statsObj.kpis.emails.semana + " | Mês Atual: " + statsObj.kpis.emails.mes + ")\n\n";
      textoKpi += "• Total de WhatsApps Marcados: " + statsObj.kpis.whats.total + "\n";
      textoKpi += "   (Hoje: " + statsObj.kpis.whats.hoje + " | Semana: " + statsObj.kpis.whats.semana + " | Mês Atual: " + statsObj.kpis.whats.mes + ")\n\n";
      textoKpi += "• Conclusões (Instalações Finalizadas): " + statsObj.kpis.conclusaoTotal + "\n";
      textoKpi += "   (Etapa 1: " + statsObj.kpis.conclusaoEtapa['1'] + " | Etapa 2: " + statsObj.kpis.conclusaoEtapa['2'] + " | Etapa 3: " + statsObj.kpis.conclusaoEtapa['3'] + ")\n";
      
      slideKpi.insertTextBox(textoKpi, 50, 100, 620, 350).getText().getTextStyle().setFontSize(14).setForegroundColor("#334155");
    }

    graficos.forEach(graf => {
      const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
      slide.insertTextBox(graf.titulo, 50, 30, 620, 50).getText().getTextStyle().setFontSize(22).setBold(true).setForegroundColor("#1e293b");
      const base64Data = graf.base64.split(',')[1];
      const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), 'image/png', graf.titulo + ".png");
      slide.insertImage(blob, 50, 100, 600, 300);
    });

    return { url: presentation.getUrl(), erro: null };
  } catch (e) { 
    return { url: null, erro: e.message };
  }
}