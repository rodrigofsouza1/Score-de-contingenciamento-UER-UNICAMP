function classificarScore(score) {
  if (score <= 15) return "ROTINA";
  if (score <= 30) return "NÍVEL 1";
  if (score <= 50) return "NÍVEL 2";
  if (score <= 70) return "NÍVEL 3";
  if (score <= 85) return "NÍVEL 4";
  return "NÍVEL 5";
}

function gravarEExcluir() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dashboard = ss.getSheetByName("Dashboard");
  var historico = ss.getSheetByName("HISTORICO");

  if (!historico) {
    throw new Error("Aba HISTORICO nao encontrada.");
  }

  var finalBruto = dashboard.getRange("G11").getDisplayValue();
  var final_score = parseFloat(finalBruto.replace("%", "").replace(",", "."));

  if (isNaN(final_score)) {
    SpreadsheetApp.getUi().alert("Score final invalido: " + finalBruto);
    return;
  }

  var bruto = dashboard.getRange("G5").getValue();
  if (bruto === "" || isNaN(bruto)) {
    SpreadsheetApp.getUi().alert("Calcule o score antes de gravar.");
    return;
  }

  var brutoPond  = dashboard.getRange("G6").getDisplayValue();
  var linear     = dashboard.getRange("G7").getDisplayValue();
  var linearPond = dashboard.getRange("G8").getDisplayValue();
  var logistico  = dashboard.getRange("G9").getDisplayValue();
  var critico    = dashboard.getRange("G10").getDisplayValue();
  var inputs     = dashboard.getRange("C3:C24").getValues().flat();

  // Percepcao do avaliador (C25)
  var percepcao  = dashboard.getRange("C25").getDisplayValue();

  // Capturar flags
  var flagsTexto = "";
  var flagHeader = dashboard.getRange("E14").getDisplayValue();
  if (flagHeader && flagHeader.indexOf("ALERTAS") >= 0) {
    var flagsList = [];
    for (var f = 15; f <= 20; f++) {
      var fval = dashboard.getRange("E" + f).getDisplayValue();
      if (fval !== "") flagsList.push(fval);
    }
    flagsTexto = flagsList.join(" | ");
  }

  // Dados complementares (C27:C35)
  var dadosCompl = dashboard.getRange("C27:C35").getDisplayValues().flat();

  // ════════════════════════════════════════════════════════════
  // NOVOS INDICADORES (C36:C37 + G27:G36)
  // ════════════════════════════════════════════════════════════
  var permSemLeito48h      = dashboard.getRange("C36").getDisplayValue();
  var totalLeitoSolicitado = dashboard.getRange("C37").getDisplayValue();
  var totalAdultoSemEmerg  = dashboard.getRange("G27").getDisplayValue();
  var totalAdultoGeral     = dashboard.getRange("G28").getDisplayValue();
  var taxaOcupAdulto       = dashboard.getRange("G29").getDisplayValue() + "%";
  var taxaOcupEmerg1       = dashboard.getRange("G30").getDisplayValue() + "%";
  var taxaOcupEmerg2       = dashboard.getRange("G31").getDisplayValue() + "%";
  var taxaOcupPorta        = dashboard.getRange("G32").getDisplayValue() + "%";
  var taxaOcupPediatria    = dashboard.getRange("G33").getDisplayValue() + "%";
  var taxaVM               = dashboard.getRange("G34").getDisplayValue() + "%";
  var totalAguardaInter    = dashboard.getRange("G35").getDisplayValue();
  var perc48hAguarda       = dashboard.getRange("G36").getDisplayValue() + "%";

  var indicadoresAvancados = [
    permSemLeito48h, totalLeitoSolicitado,
    totalAdultoSemEmerg, totalAdultoGeral,
    taxaOcupAdulto, taxaOcupEmerg1, taxaOcupEmerg2,
    taxaOcupPorta, taxaOcupPediatria, taxaVM,
    totalAguardaInter, perc48hAguarda
  ];

  var agora = new Date();
  var classificacao = classificarScore(final_score);

  // Montar linha do historico:
  // A=data, B=hora, C=classificacao, D=percepcao,
  // E=bruto, F=brutoPond, G=linear, H=linearPond, I=logistico, J=critico,
  // K=scoreFinal, L=flags,
  // M..AH = inputs C3:C24 (22 colunas)
  // AI..AQ = dados complementares C27:C35 (9 colunas)
  // AR..BC = indicadores avancados (12 colunas)

  var linha = [
    agora,              // A - data
    agora,              // B - hora
    classificacao,      // C - classificacao
    percepcao,          // D - percepcao do avaliador
    bruto,              // E - score bruto
    brutoPond,          // F - bruto ponderado
    linear,             // G - linear
    linearPond,         // H - linear ponderado
    logistico,          // I - logistico
    critico,            // J - critico
    final_score,        // K - score final
    flagsTexto          // L - flags
  ].concat(inputs)      // M..AH - 22 inputs
   .concat(dadosCompl)  // AI..AQ - 9 dados complementares
   .concat(indicadoresAvancados); // AR..BC - 12 indicadores avancados

  var ultimaLinha = historico.getLastRow();
  var proximaLinha = Math.max(ultimaLinha + 1, 3);

  historico.getRange(proximaLinha, 1, 1, linha.length).setValues([linha]);
  SpreadsheetApp.flush();

  historico.getRange(proximaLinha, 1).setNumberFormat("dd/MM/yyyy");
  historico.getRange(proximaLinha, 2).setNumberFormat("HH:mm:ss");

  // LIMPAR DASHBOARD (conteudo, preservar formato)
  dashboard.getRange("C3:C24").clearContent();
  dashboard.getRange("C25").clearContent();         // Percepcao
  dashboard.getRange("C27:C31").clearContent();     // Dados compl. manuais
  // C32:C34 sao formulas, nao limpar
  dashboard.getRange("C35").clearContent();         // Sem solicitacao leito
  dashboard.getRange("C36:C37").clearContent();     // Perm sem leito >48h + Total leito solicitado
  dashboard.getRange("G5:G12").clearContent();
  dashboard.getRange("E4").clearContent();
  dashboard.getRange("E12:G12").clearContent();  // ← Limpa nível de lotação
  dashboard.getRange("E13").clearContent();
  // NAO limpar G27:G36 → formulas permanecem ativas!
  dashboard.getRange("G35:G36").clearContent();  // ← Limpa Total Aguarda Inter. e % >48h
  // Flags: limpar conteudo e reaplicar fundo dark
  dashboard.getRange("E14:G20").clearContent();
  dashboard.getRange("E14:G20")
      .setBackground("#0f0f1a")
      .setFontFamily("Arial")
      .setBorder(false, true, true, true, false, false, "#4a4a6a", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  try {
    montarDashboardBI();
  } catch (e) {
    Logger.log("Erro ao montar dashboard BI: " + e.message);
  }

  Logger.log("Dados gravados na linha " + proximaLinha + " com percepcao: " + percepcao);
}
