function gerarPDF() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dash = ss.getSheetByName("Dashboard");

  var scoreFinalBruto = dash.getRange("G11").getDisplayValue();
  if (!scoreFinalBruto || scoreFinalBruto === "") {
    SpreadsheetApp.getUi().alert("Calcule o score antes de gerar o PDF.");
    return;
  }

  var agora = new Date();
  var dataFormatada = Utilities.formatDate(agora, Session.getScriptTimeZone(), "dd/MM/yyyy");
  var horaFormatada = Utilities.formatDate(agora, Session.getScriptTimeZone(), "HH:mm");

  // Inputs
  var dados = {
    emerg1_vm:     dash.getRange("C3").getDisplayValue(),
    emerg1_semvm:  dash.getRange("C4").getDisplayValue(),
    emerg2_vm:     dash.getRange("C5").getDisplayValue(),
    emerg2_semvm:  dash.getRange("C6").getDisplayValue(),
    s4_vm:         dash.getRange("C7").getDisplayValue(),
    s4_semvm:      dash.getRange("C8").getDisplayValue(),
    s4_iso:        dash.getRange("C9").getDisplayValue(),
    s4_hemo:       dash.getRange("C10").getDisplayValue(),
    s3_o2:         dash.getRange("C11").getDisplayValue(),
    s3_dva:        dash.getRange("C12").getDisplayValue(),
    s3_total:      dash.getRange("C13").getDisplayValue(),
    s2_o2:         dash.getRange("C14").getDisplayValue(),
    s2_monit:      dash.getRange("C15").getDisplayValue(),
    s2_total:      dash.getRange("C16").getDisplayValue(),
    s1_macas:      dash.getRange("C17").getDisplayValue(),
    s1_poltronas:  dash.getRange("C18").getDisplayValue(),
    s1_consult:    dash.getRange("C19").getDisplayValue(),
    porta_24h:     dash.getRange("C20").getDisplayValue(),
    porta_menos24: dash.getRange("C21").getDisplayValue(),
    porta_acam:    dash.getRange("C22").getDisplayValue(),
    cross:         dash.getRange("C23").getDisplayValue(),
    sem_leito:     dash.getRange("C24").getDisplayValue()
  };

  // Percepcao
  var percepcao = dash.getRange("C25").getDisplayValue() || "Não informada";

  // Dados complementares
  var compl = {
    pediatria:    dash.getRange("C27").getDisplayValue(),
    cardiologia:  dash.getRange("C28").getDisplayValue(),
    marcapasso:   dash.getRange("C29").getDisplayValue(),
    cate:         dash.getRange("C30").getDisplayValue(),
    eda_bronco:   dash.getRange("C31").getDisplayValue(),
    total_hemo:   dash.getRange("C32").getDisplayValue(),
    o2_excEmerg:  dash.getRange("C33").getDisplayValue(),
    total_vm:     dash.getRange("C34").getDisplayValue(),
    sem_solic:    dash.getRange("C35").getDisplayValue(),
    perm48h:      dash.getRange("C36").getDisplayValue(),
    leitoSolic:   dash.getRange("C37").getDisplayValue()
  };

  // Scores
  var scores = {
    bruto:       dash.getRange("G5").getDisplayValue(),
    brutoPond:   dash.getRange("G6").getDisplayValue(),
    linear:      dash.getRange("G7").getDisplayValue(),
    linearPond:  dash.getRange("G8").getDisplayValue(),
    logistico:   dash.getRange("G9").getDisplayValue(),
    critico:     dash.getRange("G10").getDisplayValue(),
    final:       dash.getRange("G11").getDisplayValue(),
    classif:     dash.getRange("G12").getDisplayValue()
  };

  // ════════════════════════════════════════════════════════════
  // INDICADORES AVANCADOS DE OCUPACAO (G27:G36)
  // ════════════════════════════════════════════════════════════
  var indOcup = {
    totalAdultoSemEmerg: dash.getRange("G27").getDisplayValue(),
    totalAdultoGeral:    dash.getRange("G28").getDisplayValue(),
    taxaOcupAdulto:      dash.getRange("G29").getDisplayValue() + "%",
    taxaOcupEmerg1:      dash.getRange("G30").getDisplayValue() + "%",
    taxaOcupEmerg2:      dash.getRange("G31").getDisplayValue() + "%",
    taxaOcupPorta:       dash.getRange("G32").getDisplayValue() + "%",
    taxaOcupPediatria:   dash.getRange("G33").getDisplayValue() + "%",
    taxaVM:              dash.getRange("G34").getDisplayValue() + "%",
    totalAguardaInter:   dash.getRange("G35").getDisplayValue(),
    perc48hAguarda:      dash.getRange("G36").getDisplayValue() + "%"
  };

  // Indicadores (coluna D)
  var indicadores = [];
  for (var row = 3; row <= 24; row++) {
    var val = dash.getRange("D" + row).getDisplayValue();
    if (val && val !== "") indicadores.push(val);
  }

  // Flags
  var flags = [];
  var flagHeader = dash.getRange("E14").getDisplayValue();
  if (flagHeader && flagHeader.indexOf("ALERTAS") >= 0) {
    for (var f = 15; f <= 20; f++) {
      var flagE = dash.getRange("E" + f).getDisplayValue();
      var flagG = dash.getRange("G" + f).getDisplayValue();
      if (flagE && flagE !== "") flags.push({ alerta: flagE, detalhe: flagG || "" });
    }
  }

  // Total geral (linha 39)
  var totalGeral = dash.getRange("C39").getDisplayValue();
  var totalGeralInfo = dash.getRange("D39").getDisplayValue();

  // Cores do nivel
  var nivelCores = {
    "ROTINA":  { bg: "#1565C0", fg: "#ffffff" },
    "NÍVEL 1": { bg: "#2E7D32", fg: "#ffffff" },
    "NÍVEL 2": { bg: "#F9A825", fg: "#000000" },
    "NÍVEL 3": { bg: "#EF6C00", fg: "#ffffff" },
    "NÍVEL 4": { bg: "#C62828", fg: "#ffffff" },
    "NÍVEL 5": { bg: "#000000", fg: "#ffffff" }
  };
  var corNivel = nivelCores[scores.classif] || nivelCores["ROTINA"];
  var scoreFinal = parseFloat(scores.final.replace("%", "").replace(",", ".")) || 0;

  // ══════════════════════════════════════
  // HTML DO RELATORIO
  // ══════════════════════════════════════

  var html = '<!DOCTYPE html><html><head><meta charset="utf-8">';
  html += '<style>';
  html += 'body{font-family:Arial,sans-serif;margin:0;padding:20px 30px;color:#333;background:#fff;font-size:11px}';
  html += '.header{background:#1a1a2e;color:#fff;padding:20px 30px;border-radius:8px;margin-bottom:20px}';
  html += '.header h1{margin:0;font-size:20px;letter-spacing:1px}';
  html += '.header .subtitle{color:#aaa;font-size:11px;margin-top:5px}';
  html += '.nivel-badge{display:inline-block;padding:10px 30px;border-radius:6px;font-size:18px;font-weight:bold;margin:10px 0}';
  html += '.score-final-box{text-align:center;padding:15px;background:#f8f8f8;border-radius:8px;border:2px solid ' + corNivel.bg + ';margin-bottom:15px}';
  html += '.score-final-num{font-size:32px;font-weight:bold;color:' + corNivel.bg + '}';
  html += '.barra-container{width:100%;height:18px;background:#e0e0e0;border-radius:10px;overflow:hidden;margin:8px 0}';
  html += '.barra-fill{height:100%;border-radius:10px;background:' + corNivel.bg + '}';
  html += '.percepcao-box{text-align:center;padding:8px;background:#f3e5f5;border-radius:6px;margin-bottom:15px;border:1px solid #ce93d8}';
  html += '.section{margin-bottom:15px}';
  html += '.section-title{font-size:12px;font-weight:bold;color:#1a1a2e;border-bottom:2px solid #1a1a2e;padding-bottom:3px;margin-bottom:6px;text-transform:uppercase;letter-spacing:1px}';
  html += 'table{width:100%;border-collapse:collapse;font-size:10px;margin-bottom:10px}';
  html += 'th{background:#1a1a2e;color:#fff;padding:5px 7px;text-align:left;font-size:9px;text-transform:uppercase}';
  html += 'td{padding:4px 7px;border-bottom:1px solid #e0e0e0}';
  html += 'tr:nth-child(even){background:#f5f5f5}';
  html += '.flag-red{color:#C62828;font-weight:bold}';
  html += '.flag-yellow{color:#EF6C00;font-weight:bold}';
  html += '.indicador{background:#fff3e0;padding:2px 6px;border-radius:3px;display:inline-block;margin:2px;font-size:9px}';
  html += '.two-col{display:flex;gap:12px}';
  html += '.two-col>div{flex:1}';
  html += '.compl-table td:last-child{text-align:center;font-weight:bold}';
  html += '.ocup-val{text-align:center;font-weight:bold}';
  html += '.footer{text-align:center;color:#999;font-size:8px;margin-top:25px;border-top:1px solid #e0e0e0;padding-top:8px}';
  html += '</style></head><body>';

  // HEADER
  html += '<div class="header">';
  html += '<h1>SCORE DE CONTINGENCIAMENTO - UER UNICAMP</h1>';
  html += '<div class="subtitle">Relatorio gerado em ' + dataFormatada + ' as ' + horaFormatada + '</div>';
  html += '</div>';

  // SCORE FINAL + CLASSIFICACAO
  html += '<div class="score-final-box">';
  html += '<div class="score-final-num">' + scores.final + '</div>';
  html += '<div class="barra-container"><div class="barra-fill" style="width:' + Math.min(scoreFinal, 100) + '%"></div></div>';
  html += '<div class="nivel-badge" style="background:' + corNivel.bg + ';color:' + corNivel.fg + '">' + scores.classif + '</div>';
  html += '</div>';

  // PERCEPCAO DO AVALIADOR
  html += '<div class="percepcao-box">';
  html += '<b>Percepção do Avaliador:</b> ' + percepcao;
  html += '</div>';

  // FLAGS
  if (flags.length > 0) {
    html += '<div class="section">';
    html += '<div class="section-title">⚠ Alertas Criticos</div>';
    html += '<table><tr><th>Indicador</th><th>Detalhe</th></tr>';
    for (var fi = 0; fi < flags.length; fi++) {
      var cssClass = flags[fi].alerta.indexOf("🔴") >= 0 ? "flag-red" : "flag-yellow";
      html += '<tr><td class="' + cssClass + '">' + flags[fi].alerta + '</td><td>' + flags[fi].detalhe + '</td></tr>';
    }
    html += '</table></div>';
  }

  // DADOS DE ENTRADA
  html += '<div class="section">';
  html += '<div class="section-title">Dados de Entrada</div>';
  html += '<div class="two-col">';

  // Coluna esquerda
  html += '<div><table>';
  html += '<tr><th colspan="2">EMERGENCIAS</th></tr>';
  html += '<tr><td>Emerg 1 - VM</td><td><b>' + dados.emerg1_vm + '</b></td></tr>';
  html += '<tr><td>Emerg 1 - Sem VM</td><td><b>' + dados.emerg1_semvm + '</b></td></tr>';
  html += '<tr><td>Emerg 2 - VM</td><td><b>' + dados.emerg2_vm + '</b></td></tr>';
  html += '<tr><td>Emerg 2 - Sem VM</td><td><b>' + dados.emerg2_semvm + '</b></td></tr>';
  html += '<tr><th colspan="2">SETOR 4</th></tr>';
  html += '<tr><td>VM</td><td><b>' + dados.s4_vm + '</b></td></tr>';
  html += '<tr><td>Sem VM</td><td><b>' + dados.s4_semvm + '</b></td></tr>';
  html += '<tr><td>Isolamento Resp.</td><td><b>' + dados.s4_iso + '</b></td></tr>';
  html += '<tr><td>Hemodialise</td><td><b>' + dados.s4_hemo + '</b></td></tr>';
  html += '<tr><th colspan="2">SETOR 3</th></tr>';
  html += '<tr><td>Em uso de O2</td><td><b>' + dados.s3_o2 + '</b></td></tr>';
  html += '<tr><td>Com DVA</td><td><b>' + dados.s3_dva + '</b></td></tr>';
  html += '<tr><td>Total pacientes</td><td><b>' + dados.s3_total + '</b></td></tr>';
  html += '</table></div>';

  // Coluna direita
  html += '<div><table>';
  html += '<tr><th colspan="2">SETOR 2</th></tr>';
  html += '<tr><td>Em uso de O2</td><td><b>' + dados.s2_o2 + '</b></td></tr>';
  html += '<tr><td>Monitorizacao/BI</td><td><b>' + dados.s2_monit + '</b></td></tr>';
  html += '<tr><td>Total pacientes</td><td><b>' + dados.s2_total + '</b></td></tr>';
  html += '<tr><th colspan="2">SETOR 1</th></tr>';
  html += '<tr><td>Internados em macas</td><td><b>' + dados.s1_macas + '</b></td></tr>';
  html += '<tr><td>Internados em poltronas</td><td><b>' + dados.s1_poltronas + '</b></td></tr>';
  html += '<tr><td>Internados em consultorios</td><td><b>' + dados.s1_consult + '</b></td></tr>';
  html += '<tr><th colspan="2">PORTA</th></tr>';
  html += '<tr><td>Permanencia > 24h</td><td><b>' + dados.porta_24h + '</b></td></tr>';
  html += '<tr><td>Permanencia < 24h</td><td><b>' + dados.porta_menos24 + '</b></td></tr>';
  html += '<tr><td>Pacientes acamados</td><td><b>' + dados.porta_acam + '</b></td></tr>';
  html += '<tr><th colspan="2">CROSS E LEITOS</th></tr>';
  html += '<tr><td>Cross proximas 6h</td><td><b>' + dados.cross + '</b></td></tr>';
  html += '<tr><td>Sem leito > 24h</td><td><b>' + dados.sem_leito + '</b></td></tr>';
  html += '</table></div>';

  html += '</div>'; // fecha two-col

  // Total geral
  if (totalGeral) {
    html += '<table><tr style="background:#333;color:#fff">';
    html += '<td><b>TOTAL GERAL UER</b></td>';
    html += '<td style="text-align:center"><b>' + totalGeral + '</b></td>';
    html += '<td>' + (totalGeralInfo || '') + '</td>';
    html += '</tr></table>';
  }
  html += '</div>';

  // DADOS COMPLEMENTARES
  html += '<div class="section">';
  html += '<div class="section-title">Dados Complementares</div>';
  html += '<table class="compl-table">';
  html += '<tr><th>Indicador</th><th>Valor</th></tr>';
  html += '<tr><td>Pacientes Pediatria</td><td>' + (compl.pediatria || "0") + '</td></tr>';
  html += '<tr><td>Pacientes Cardiologia</td><td>' + (compl.cardiologia || "0") + '</td></tr>';
  html += '<tr><td>Em uso de Marcapasso</td><td>' + (compl.marcapasso || "0") + '</td></tr>';
  html += '<tr><td>Aguarda CATE</td><td>' + (compl.cate || "0") + '</td></tr>';
  html += '<tr><td>Aguarda EDA/Broncoscopia</td><td>' + (compl.eda_bronco || "0") + '</td></tr>';
  html += '<tr><td>Total em Hemodialise</td><td>' + (compl.total_hemo || "0") + '</td></tr>';
  html += '<tr><td>Uso de O₂ (exceto Emerg)</td><td>' + (compl.o2_excEmerg || "0") + '</td></tr>';
  html += '<tr><td>Total em Vent. Mecanica</td><td>' + (compl.total_vm || "0") + '</td></tr>';
  html += '<tr><td>Sem solicitação de leito &gt;24h</td><td>' + (compl.sem_solic || "0") + '</td></tr>';
  html += '<tr><td>Permanência sem leito &gt;48h</td><td>' + (compl.perm48h || "0") + '</td></tr>';
  html += '<tr><td>Total com leito solicitado</td><td>' + (compl.leitoSolic || "0") + '</td></tr>';
  html += '</table></div>';

  // ════════════════════════════════════════════════════════════
  // INDICADORES DE OCUPACAO (NOVO)
  // ════════════════════════════════════════════════════════════
  html += '<div class="section">';
  html += '<div class="section-title">Indicadores de Ocupação</div>';
  html += '<table>';
  html += '<tr><th>Indicador</th><th style="text-align:center">Valor</th></tr>';

  var tabelaOcup = [
    ["Total Adulto (sem Emerg.)",              indOcup.totalAdultoSemEmerg],
    ["Total Adulto Geral",                     indOcup.totalAdultoGeral],
    ["Taxa Ocupação Adulto (sem Emerg.)",      indOcup.taxaOcupAdulto],
    ["Taxa Ocupação Emerg. 1",                 indOcup.taxaOcupEmerg1],
    ["Taxa Ocupação Emerg. 2",                 indOcup.taxaOcupEmerg2],
    ["Taxa Ocupação Porta",                    indOcup.taxaOcupPorta],
    ["Taxa Ocupação Pediatria",                indOcup.taxaOcupPediatria],
    ["Taxa Pacientes em VM",                   indOcup.taxaVM],
    ["Total Pacientes Aguardando Internação",  indOcup.totalAguardaInter],
    ["% Pacientes &gt;48h Aguardando Int.",    indOcup.perc48hAguarda]
  ];

  for (var oc = 0; oc < tabelaOcup.length; oc++) {
    var bgOcup = (oc % 2 === 0) ? "#f0f4ff" : "#ffffff";
    html += '<tr style="background:' + bgOcup + '">';
    html += '<td>' + tabelaOcup[oc][0] + '</td>';
    html += '<td class="ocup-val">' + tabelaOcup[oc][1] + '</td>';
    html += '</tr>';
  }
  html += '</table></div>';

  // INDICADORES (coluna D)
  if (indicadores.length > 0) {
    html += '<div class="section">';
    html += '<div class="section-title">Indicadores</div>';
    for (var ind = 0; ind < indicadores.length; ind++) {
      html += '<span class="indicador">' + indicadores[ind] + '</span>';
    }
    html += '</div>';
  }

  // TABELA DE SCORES
  html += '<div class="section">';
  html += '<div class="section-title">Detalhamento dos Scores</div>';
  html += '<table>';
  html += '<tr><th>Metrica</th><th>Valor</th></tr>';
  html += '<tr><td>Score Bruto</td><td>' + scores.bruto + '</td></tr>';
  html += '<tr><td>Score Bruto Ponderado</td><td>' + scores.brutoPond + '</td></tr>';
  html += '<tr><td>Score Linear</td><td>' + scores.linear + '</td></tr>';
  html += '<tr><td>Score Linear Ponderado</td><td>' + scores.linearPond + '</td></tr>';
  html += '<tr><td>Score Logistico</td><td>' + scores.logistico + '</td></tr>';
  html += '<tr style="background:#fff3e0"><td><b>Score Critico</b></td><td><b>' + scores.critico + '</b></td></tr>';
  html += '<tr style="background:' + corNivel.bg + ';color:' + corNivel.fg + '"><td><b>SCORE FINAL</b></td><td><b>' + scores.final + '</b></td></tr>';
  html += '</table></div>';

  // FOOTER
  html += '<div class="footer">';
  html += 'Score de Contingenciamento UER - Hospital de Clinicas UNICAMP<br>';
  html += 'Documento gerado automaticamente em ' + dataFormatada + ' ' + horaFormatada;
  html += '</div>';

  html += '</body></html>';

  // CONVERTER E SALVAR
  var blob = HtmlService.createHtmlOutput(html).getContent();
  var pdfBlob = Utilities.newBlob(blob, MimeType.HTML, "relatorio_temp").getAs(MimeType.PDF);
  var nomeArquivo = "Score_UER_" + Utilities.formatDate(agora, Session.getScriptTimeZone(), "yyyyMMdd_HHmm") + ".pdf";
  pdfBlob.setName(nomeArquivo);

  var pasta;
  try {
    var pastas = DriveApp.getFoldersByName("Relatorios UER");
    if (pastas.hasNext()) { pasta = pastas.next(); }
    else { pasta = DriveApp.createFolder("Relatorios UER"); }
  } catch(e) { pasta = DriveApp.getRootFolder(); }

  var arquivo = pasta.createFile(pdfBlob);
  var url = arquivo.getUrl();

  var htmlDialog = '<div style="font-family:Arial;padding:10px">'
    + '<p style="font-size:14px">PDF gerado com sucesso!</p>'
    + '<p><b>' + nomeArquivo + '</b></p>'
    + '<p>Salvo na pasta: <b>Relatorios UER</b></p>'
    + '<a href="' + url + '" target="_blank" '
    + 'style="display:inline-block;padding:10px 20px;background:#1565C0;color:#fff;'
    + 'text-decoration:none;border-radius:5px;font-weight:bold;margin-top:10px">'
    + 'ABRIR PDF</a></div>';

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(htmlDialog).setWidth(350).setHeight(250),
    "Relatorio PDF"
  );

  Logger.log("PDF gerado: " + nomeArquivo + " -> " + url);
}
