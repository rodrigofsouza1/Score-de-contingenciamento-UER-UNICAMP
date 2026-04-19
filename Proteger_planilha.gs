function protegerAbas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // ════════════════════════════════════════════════════════════
  // DASHBOARD
  // ═══════════════════════════════════════════════════════════
  var dash = ss.getSheetByName("Dashboard");
  if (dash) {
    var pDash = dash.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    for (var i = 0; i < pDash.length; i++) pDash[i].remove();
    var pDashR = dash.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    for (var j = 0; j < pDashR.length; j++) pDashR[j].remove();

    var protDash = dash.protect().setDescription("Dashboard - Use apenas campos de entrada");
    protDash.setUnprotectedRanges([
      dash.getRange("C3:C24"),
      dash.getRange("C25"),
      dash.getRange("C27:C31"),
      dash.getRange("C35"),
      dash.getRange("C36"),
      dash.getRange("C37")
    ]);
    protDash.setWarningOnly(true);
    Logger.log("Dashboard protegido.");
  }

  // ════════════════════════════════════════════════════════════
  // VARIAVEIS
  // ════════════════════════════════════════════════════════════
  var vars = ss.getSheetByName("Variaveis");
  if (vars) {
    var pVars = vars.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    for (var k = 0; k < pVars.length; k++) pVars[k].remove();

    var protVars = vars.protect().setDescription("Variaveis - Nao editar manualmente");
    protVars.setWarningOnly(true);
    Logger.log("Variaveis protegido.");
  }

  // ════════════════════════════════════════════════════════════
  // HISTORICO
  // ════════════════════════════════════════════════════════════
  var hist = ss.getSheetByName("HISTORICO");
  if (hist) {
    var pHist = hist.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    for (var m = 0; m < pHist.length; m++) pHist[m].remove();

    var protHist = hist.protect().setDescription("Historico - Nao editar manualmente");
    protHist.setWarningOnly(true);
    Logger.log("HISTORICO protegido.");
  }

  // ════════════════════════════════════════════════════════════
  // SIMULACOES
  // ════════════════════════════════════════════════════════════
  var sim = ss.getSheetByName("SIMULACOES");
  if (sim) {
    var pSim = sim.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    for (var n = 0; n < pSim.length; n++) pSim[n].remove();

    var protSim = sim.protect().setDescription("Simulacoes - Nao editar manualmente");
    protSim.setWarningOnly(true);
    Logger.log("SIMULACOES protegido.");
  }

  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert(
    "Proteção aplicada!\n\nAgora rode 'instalarTriggerProtecao' para ativar a reversão automática."
  );
}

function removerProtecoes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var abas = ss.getSheets();
  for (var i = 0; i < abas.length; i++) {
    var ps = abas[i].getProtections(SpreadsheetApp.ProtectionType.SHEET);
    for (var j = 0; j < ps.length; j++) ps[j].remove();
    var pr = abas[i].getProtections(SpreadsheetApp.ProtectionType.RANGE);
    for (var k = 0; k < pr.length; k++) pr[k].remove();
  }
  SpreadsheetApp.getUi().alert("Protecoes removidas!");
}
