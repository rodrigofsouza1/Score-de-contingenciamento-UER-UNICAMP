// ════════════════════════════════════════════════════════════
// CONFIGURAÇÃO: ID da planilha de origem
// Pegue o ID na URL da outra planilha:
// https://docs.google.com/spreadsheets/d/ <<< ESTE ID AQUI >>> /edit
// ════════════════════════════════════════════════════════════

var ID_PLANILHA_PACIENTES = "1RthkjF_kxah7I-j8mvutD8p3fusIvCxCTJ3gU1-Tf-E";
var NOME_ABA_ORIGEM = "ATUAL";

// ════════════════════════════════════════════════════════════
// FUNÇÃO PRINCIPAL: IMPORTAR E CALCULAR INDICADORES
// ════════════════════════════════════════════════════════════

function importarIndicadores() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dashboard = ss.getSheetByName("Dashboard");

  // ── Abrir planilha de origem ──
  var planOrigem;
  try {
    planOrigem = SpreadsheetApp.openById(ID_PLANILHA_PACIENTES);
  } catch (e) {
    SpreadsheetApp.getUi().alert(
      "Erro ao acessar a planilha de pacientes.\n\n"
      + "Verifique:\n"
      + "1. O ID está correto\n"
      + "2. Você tem acesso de editor/visualizador\n\n"
      + "Erro: " + e.message
    );
    return;
  }

  var abaOrigem = planOrigem.getSheetByName(NOME_ABA_ORIGEM);
  if (!abaOrigem) {
    SpreadsheetApp.getUi().alert("Aba '" + NOME_ABA_ORIGEM + "' não encontrada na planilha de pacientes.");
    return;
  }

  // ── Pegar dados das colunas D (admissão) e R (leito) ──
  var ultimaLinha = abaOrigem.getLastRow();
  if (ultimaLinha < 2) {
    SpreadsheetApp.getUi().alert("Planilha de pacientes está vazia.");
    return;
  }

  // Coluna D = 4, Coluna R = 18
  var colD = abaOrigem.getRange(2, 4, ultimaLinha - 1, 1).getValues().flat();
  var colR = abaOrigem.getRange(2, 18, ultimaLinha - 1, 1).getValues().flat();

  var agora = new Date();
  var ms24h = 24 * 60 * 60 * 1000;
  var ms48h = 48 * 60 * 60 * 1000;

  // ════════════════════════════════════════════════════════════
  // CALCULAR INDICADORES
  // ════════════════════════════════════════════════════════════

  var permMaior24h = 0;
  var permMaior48h = 0;
  var semSolicLeito = 0;
  var totalLeitoSolic = 0;

  for (var i = 0; i < colD.length; i++) {
    var dataAdm = colD[i];
    var statusLeito = (colR[i] || "").toString().trim().toUpperCase();

    if (!dataAdm || dataAdm === "") continue;
    if (!(dataAdm instanceof Date)) dataAdm = new Date(dataAdm);
    if (isNaN(dataAdm.getTime())) continue;

    var diffMs = agora.getTime() - dataAdm.getTime();

    if (diffMs > ms24h) permMaior24h++;
    if (diffMs > ms48h) permMaior48h++;

    if (statusLeito === "SOLICITAR INTERNAÇÃO" || statusLeito === "SOLICITAR INTERNACAO") {
      semSolicLeito++;
    }

    if (statusLeito !== ""
        && statusLeito !== "SOLICITAR INTERNAÇÃO"
        && statusLeito !== "SOLICITAR INTERNACAO"
        && statusLeito !== "AVALIAR INTERNAÇÃO"
        && statusLeito !== "AVALIAR INTERNACAO") {
      totalLeitoSolic++;
    }
  }

  // ════════════════════════════════════════════════════════════
  // GRAVAR INDICADORES PRIMÁRIOS NO DASHBOARD
  // ═══════════════════════════════════════════════════════════

  dashboard.getRange("C24").setValue(permMaior24h);
  dashboard.getRange("C35").setValue(semSolicLeito);
  dashboard.getRange("C36").setValue(permMaior48h);
  dashboard.getRange("C37").setValue(totalLeitoSolic);

  SpreadsheetApp.flush(); // Garante que C36, C37 estejam gravados antes de calcular derivados

  // ════════════════════════════════════════════════════════════
  // CALCULAR E GRAVAR INDICADORES DERIVADOS
  // ════════════════════════════════════════════════════════════

  // G35: TOTAL PACIENTES AGUAR. INTER. = C36 + C37
  var totalAguardaInter = semSolicLeito + totalLeitoSolic;
  dashboard.getRange("G35").setValue(totalAguardaInter);

  // G36: % PACIENTES >48H AG. INTERNAÇÃO = C36 / G28 * 100
  var totalAdultoGeral = parseFloat(dashboard.getRange("G28").getValue()) || 0;
  var perc48h = 0;
  if (totalAdultoGeral > 0) {
    perc48h = Math.round((permMaior48h / totalAdultoGeral) * 10000) / 100;
  }
  dashboard.getRange("G36").setValue(perc48h);

  SpreadsheetApp.flush();

  // ════════════════════════════════════════════════════════════
  // LOG E CONFIRMAÇÃO
  // ═══════════════════════════════════════════════════════════

  var resumo = "Indicadores importados com sucesso!\n\n"
    + "Fonte: " + planOrigem.getName() + " (aba " + NOME_ABA_ORIGEM + ")\n"
    + "Registros analisados: " + colD.length + "\n\n"
    + "═══ PRIMÁRIOS ═══\n"
    + "C24 - Permanência >24h: " + permMaior24h + "\n"
    + "C35 - Sem solicitação leito: " + semSolicLeito + "\n"
    + "C36 - Permanência >48h: " + permMaior48h + "\n"
    + "C37 - Total leito solicitado: " + totalLeitoSolic + "\n\n"
    + "═══ DERIVADOS ═══\n"
    + "G35 - Total aguarda internação (C35+C37): " + totalAguardaInter + "\n"
    + "G36 - % >48h aguardando (C36/G28×100): " + perc48h + "%\n\n"
    + "Atualizado em: " + Utilities.formatDate(agora, "America/Sao_Paulo", "dd/MM/yyyy HH:mm:ss");

  Logger.log(resumo);
  SpreadsheetApp.getUi().alert(resumo);
}

// ════════════════════════════════════════════════════════════
// ATUALIZAÇÃO AUTOMÁTICA (TRIGGER DE TEMPO)
// Roda a cada X minutos para manter dados atualizados
// ════════════════════════════════════════════════════════════

function instalarTriggerImportacao() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "importarIndicadoresAuto") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  ScriptApp.newTrigger("importarIndicadoresAuto")
    .timeBased()
    .everyMinutes(15)
    .create();

  SpreadsheetApp.getUi().alert(
    "Trigger de importação automática instalado!\n\n"
    + "Os indicadores serão atualizados a cada 15 minutos automaticamente."
  );
}

function removerTriggerImportacao() {
  var triggers = ScriptApp.getProjectTriggers();
  var removidos = 0;
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "importarIndicadoresAuto") {
      ScriptApp.deleteTrigger(triggers[i]);
      removidos++;
    }
  }
  SpreadsheetApp.getUi().alert("Triggers de importação removidos: " + removidos);
}

// Versão silenciosa (sem alert) para o trigger automático
function importarIndicadoresAuto() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dashboard = ss.getSheetByName("Dashboard");

  var planOrigem;
  try {
    planOrigem = SpreadsheetApp.openById(ID_PLANILHA_PACIENTES);
  } catch (e) {
    Logger.log("Erro ao acessar planilha de pacientes: " + e.message);
    return;
  }

  var abaOrigem = planOrigem.getSheetByName(NOME_ABA_ORIGEM);
  if (!abaOrigem) {
    Logger.log("Aba '" + NOME_ABA_ORIGEM + "' não encontrada.");
    return;
  }

  var ultimaLinha = abaOrigem.getLastRow();
  if (ultimaLinha < 2) return;

  var colD = abaOrigem.getRange(2, 4, ultimaLinha - 1, 1).getValues().flat();
  var colR = abaOrigem.getRange(2, 18, ultimaLinha - 1, 1).getValues().flat();

  var agora = new Date();
  var ms24h = 24 * 60 * 60 * 1000;
  var ms48h = 48 * 60 * 60 * 1000;

  var permMaior24h = 0;
  var permMaior48h = 0;
  var semSolicLeito = 0;
  var totalLeitoSolic = 0;

  for (var i = 0; i < colD.length; i++) {
    var dataAdm = colD[i];
    var statusLeito = (colR[i] || "").toString().trim().toUpperCase();

    if (!dataAdm || dataAdm === "") continue;
    if (!(dataAdm instanceof Date)) dataAdm = new Date(dataAdm);
    if (isNaN(dataAdm.getTime())) continue;

    var diffMs = agora.getTime() - dataAdm.getTime();

    if (diffMs > ms24h) permMaior24h++;
    if (diffMs > ms48h) permMaior48h++;

    if (statusLeito === "SOLICITAR INTERNAÇÃO" || statusLeito === "SOLICITAR INTERNACAO") {
      semSolicLeito++;
    }

    if (statusLeito !== ""
        && statusLeito !== "SOLICITAR INTERNAÇÃO"
        && statusLeito !== "SOLICITAR INTERNACAO"
        && statusLeito !== "AVALIAR INTERNAÇÃO"
        && statusLeito !== "AVALIAR INTERNACAO") {
      totalLeitoSolic++;
    }
  }

  // Gravar primários
  dashboard.getRange("C24").setValue(permMaior24h);
  dashboard.getRange("C35").setValue(semSolicLeito);
  dashboard.getRange("C36").setValue(permMaior48h);
  dashboard.getRange("C37").setValue(totalLeitoSolic);

  SpreadsheetApp.flush();

  // Gravar derivados
  var totalAguardaInter = semSolicLeito + totalLeitoSolic;
  dashboard.getRange("G35").setValue(totalAguardaInter);

  var totalAdultoGeral = parseFloat(dashboard.getRange("G28").getValue()) || 0;
  var perc48h = 0;
  if (totalAdultoGeral > 0) {
    perc48h = Math.round((permMaior48h / totalAdultoGeral) * 10000) / 100;
  }
  dashboard.getRange("G36").setValue(perc48h);

  SpreadsheetApp.flush();

  Logger.log("Importação automática: C24=" + permMaior24h
    + " C35=" + semSolicLeito
    + " C36=" + permMaior48h
    + " C37=" + totalLeitoSolic
    + " G35=" + totalAguardaInter
    + " G36=" + perc48h + "%"
    + " em " + Utilities.formatDate(agora, "America/Sao_Paulo", "dd/MM/yyyy HH:mm:ss"));
}
