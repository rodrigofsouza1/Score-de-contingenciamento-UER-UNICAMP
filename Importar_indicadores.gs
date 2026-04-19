// ════════════════════════════════════════════════════════════
// CONFIGURAÇÃO: IDs das planilhas de origem
// ════════════════════════════════════════════════════════════

var ID_PLANILHA_PACIENTES = "1RthkjF_kxah7I-j8mvutD8p3fusIvCxCTJ3gU1-Tf-E";
var NOME_ABA_ORIGEM = "ATUAL";

var ID_PLANILHA_REGULADOS = "1XlkypNaGgNbCtbW80s3bIBerwb8T7ajaiPkdP0J8g4o";
var NOME_ABA_REGULADOS = "Controle pacientes regulados";

// ════════════════════════════════════════════════════════════
// MAPEAMENTO DE ESPECIALIDADES (coluna C da origem → célula destino)
// ════════════════════════════════════════════════════════════

var ESPECIALIDADES = [
  { codigos: ["UCA"],        destino: "C31" },  // Cardiologia
  { codigos: ["UCM"],        destino: "C32" },  // Clínica Médica
  { codigos: ["UCT"],        destino: "C33" },  // Cirurgia do Trauma
  { codigos: ["UNC"],        destino: "C34" },  // Neurocirurgia
  { codigos: ["UNL","UNA"],  destino: "C35" },  // Neuroclínica
  { codigos: ["UOR"],        destino: "C36" },  // Ortopedia
  { codigos: ["UPS"],        destino: "C37" },  // Psiquiatria
  { codigos: ["UOF"],        destino: "C38" }   // Oftalmologia
];

// ════════════════════════════════════════════════════════════
// FUNÇÃO AUXILIAR: contar especialidades e indicadores
// ════════════════════════════════════════════════════════════

function processarDadosOrigem(abaOrigem) {
  var ultimaLinha = abaOrigem.getLastRow();
  if (ultimaLinha < 2) return null;

  var colC = abaOrigem.getRange(2, 3, ultimaLinha - 1, 1).getValues().flat();  // Coluna C (especialidade)
  var colD = abaOrigem.getRange(2, 4, ultimaLinha - 1, 1).getValues().flat();  // Coluna D (data admissão)
  var colF = abaOrigem.getRange(2, 6, ultimaLinha - 1, 1).getValues().flat();  // Coluna F (setor/local)
  var colR = abaOrigem.getRange(2, 18, ultimaLinha - 1, 1).getValues().flat(); // Coluna R (status leito)

  var agora = new Date();
  var ms24h = 24 * 60 * 60 * 1000;
  var ms48h = 48 * 60 * 60 * 1000;

  var permMaior24h = 0;
  var permMaior48h = 0;
  var semSolicLeito = 0;
  var totalLeitoSolic = 0;
  var totalSetor3 = 0;
  var totalSetor2 = 0;
  var portaMaior24h = 0;
  var portaMenor24h = 0;

  // Contagem de especialidades
  var contEsp = {};
  for (var e = 0; e < ESPECIALIDADES.length; e++) {
    for (var cc = 0; cc < ESPECIALIDADES[e].codigos.length; cc++) {
      contEsp[ESPECIALIDADES[e].codigos[cc]] = 0;
    }
  }

  for (var i = 0; i < colD.length; i++) {
    var dataAdm = colD[i];
    var statusLeito = (colR[i] || "").toString().trim().toUpperCase();
    var especialidade = (colC[i] || "").toString().trim().toUpperCase();
    var setor = (colF[i] || "").toString().trim().toUpperCase();

    // Contagem de setores
    if (setor.indexOf("SETOR 3") >= 0) totalSetor3++;
    if (setor.indexOf("SETOR 2") >= 0) totalSetor2++;

    // Contagem de especialidades (excluir PORTA)
    if (especialidade !== "" && setor.indexOf("PORTA") === -1) {
      if (contEsp.hasOwnProperty(especialidade)) {
        contEsp[especialidade]++;
      }
    }

    // Indicadores de tempo/leito
    if (!dataAdm || dataAdm === "") continue;
    if (!(dataAdm instanceof Date)) dataAdm = new Date(dataAdm);
    if (isNaN(dataAdm.getTime())) continue;

    var diffMs = agora.getTime() - dataAdm.getTime();

    // Contagem Porta por tempo de permanência
    if (setor.indexOf("PORTA") >= 0) {
      if (diffMs > ms24h) {
        portaMaior24h++;
      } else {
        portaMenor24h++;
      }
    }

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

  // Agrupar contagens por destino
  var especialidadesResult = [];
  for (var e = 0; e < ESPECIALIDADES.length; e++) {
    var total = 0;
    for (var cc = 0; cc < ESPECIALIDADES[e].codigos.length; cc++) {
      total += contEsp[ESPECIALIDADES[e].codigos[cc]];
    }
    especialidadesResult.push({ destino: ESPECIALIDADES[e].destino, total: total });
  }

  return {
    permMaior24h: permMaior24h,
    permMaior48h: permMaior48h,
    semSolicLeito: semSolicLeito,
    totalLeitoSolic: totalLeitoSolic,
    totalSetor3: totalSetor3,
    totalSetor2: totalSetor2,
    portaMaior24h: portaMaior24h,
    portaMenor24h: portaMenor24h,
    especialidades: especialidadesResult,
    totalRegistros: colD.length
  };
}

// ════════════════════════════════════════════════════════════
// FUNÇÃO AUXILIAR: contar Cross "AGUARDANDO" (planilha regulados)
// Lê coluna I da aba "Controle pacientes regulados"
// ════════════════════════════════════════════════════════════

function contarCrossAguardando() {
  var planReg;
  try {
    planReg = SpreadsheetApp.openById(ID_PLANILHA_REGULADOS);
  } catch (e) {
    Logger.log("Erro ao acessar planilha de regulados: " + e.message);
    return 0;
  }

  // Buscar aba pelo nome (ignorando espaços extras)
  var abaReg = null;
  var abas = planReg.getSheets();
  for (var i = 0; i < abas.length; i++) {
    if (abas[i].getName().trim() === NOME_ABA_REGULADOS.trim()) {
      abaReg = abas[i];
      break;
    }
  }

  if (!abaReg) {
    Logger.log("Aba '" + NOME_ABA_REGULADOS + "' não encontrada na planilha de regulados.");
    return 0;
  }

  var ultLinha = abaReg.getLastRow();
  if (ultLinha < 2) return 0;

  // Coluna I = 9
  var colI = abaReg.getRange(2, 9, ultLinha - 1, 1).getDisplayValues().flat();
  var count = 0;
  var valoresEncontrados = {};

  for (var i = 0; i < colI.length; i++) {
    var raw = (colI[i] || "").toString().trim();
    var status = raw.toUpperCase();

    if (raw !== "") {
      valoresEncontrados[raw] = (valoresEncontrados[raw] || 0) + 1;
    }

    if (status.indexOf("AGUARDANDO") >= 0) count++;
  }

  Logger.log("Cross - Valores encontrados na coluna I: " + JSON.stringify(valoresEncontrados));
  Logger.log("Cross - Total AGUARDANDO: " + count + " de " + colI.length + " linhas");

  return count;
}
// ════════════════════════════════════════════════════════════
// FUNÇÃO AUXILIAR: gravar resultados no dashboard
// ════════════════════════════════════════════════════════════

function gravarResultados(dashboard, resultado, crossAguardando) {
  // Setores automáticos (do CENSO)
  dashboard.getRange("C13").setValue(resultado.totalSetor3);      // Total Setor 3
  dashboard.getRange("C16").setValue(resultado.totalSetor2);      // Total Setor 2
  dashboard.getRange("C20").setValue(resultado.portaMaior24h);    // Porta >24h
  dashboard.getRange("C21").setValue(resultado.portaMenor24h);    // Porta <24h

  // Cross nas próximas 6h (planilha de regulados)
  dashboard.getRange("C23").setValue(crossAguardando);            // Cross 6h

  // Indicadores primários
  dashboard.getRange("C24").setValue(resultado.permMaior24h);     // Sem leito >24h
  dashboard.getRange("C42").setValue(resultado.semSolicLeito);    // Sem solicitação de leito
  dashboard.getRange("C43").setValue(resultado.permMaior48h);     // Permanência >48h
  dashboard.getRange("C44").setValue(resultado.totalLeitoSolic);  // Total leito solicitado

  // Especialidades
  for (var e = 0; e < resultado.especialidades.length; e++) {
    dashboard.getRange(resultado.especialidades[e].destino).setValue(resultado.especialidades[e].total);
  }

  SpreadsheetApp.flush();

  // Indicadores derivados
  // G35: TOTAL PACIENTES AGUAR. INTER. = C42 + C44
  var totalAguardaInter = resultado.semSolicLeito + resultado.totalLeitoSolic;
  dashboard.getRange("G35").setValue(totalAguardaInter);

  // G36: % PACIENTES >48H AG. INTERNAÇÃO = C43 / G28 * 100
  var totalAdultoGeral = parseFloat(dashboard.getRange("G28").getValue()) || 0;
  var perc48h = 0;
  if (totalAdultoGeral > 0) {
    perc48h = Math.round((resultado.permMaior48h / totalAdultoGeral) * 10000) / 100;
  }
  dashboard.getRange("G36").setValue(perc48h);

  SpreadsheetApp.flush();

  return { totalAguardaInter: totalAguardaInter, perc48h: perc48h };
}

// ════════════════════════════════════════════════════════════
// FUNÇÃO PRINCIPAL: IMPORTAR E CALCULAR INDICADORES
// ════════════════════════════════════════════════════════════

function importarIndicadores() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dashboard = ss.getSheetByName("Dashboard");

  // ── Abrir planilha de origem (CENSO) ──
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

  var resultado = processarDadosOrigem(abaOrigem);
  if (!resultado) {
    SpreadsheetApp.getUi().alert("Planilha de pacientes está vazia.");
    return;
  }

  // ── Cross da planilha de regulados ──
  var crossAguardando = contarCrossAguardando();

  var derivados = gravarResultados(dashboard, resultado, crossAguardando);

  var agora = new Date();
  var resumo = "Indicadores importados com sucesso!\n\n"
    + "Fonte: " + planOrigem.getName() + " (aba " + NOME_ABA_ORIGEM + ")\n"
    + "Registros analisados: " + resultado.totalRegistros + "\n\n"
    + "═══ SETORES / PORTA ═══\n"
    + "C13 - Total Setor 3: " + resultado.totalSetor3 + "\n"
    + "C16 - Total Setor 2: " + resultado.totalSetor2 + "\n"
    + "C20 - Porta >24h: " + resultado.portaMaior24h + "\n"
    + "C21 - Porta <24h: " + resultado.portaMenor24h + "\n\n"
    + "═══ REGULADOS ═══\n"
    + "C23 - Cross 6h (Aguardando): " + crossAguardando + "\n\n"
    + "═══ PRIMÁRIOS ═══\n"
    + "C24 - Permanência >24h: " + resultado.permMaior24h + "\n"
    + "C42 - Sem solicitação leito: " + resultado.semSolicLeito + "\n"
    + "C43 - Permanência >48h: " + resultado.permMaior48h + "\n"
    + "C44 - Total leito solicitado: " + resultado.totalLeitoSolic + "\n\n"
    + "═══ ESPECIALIDADES ═══\n";

  var nomesEsp = ["Cardiologia","Clín.Médica","Cir.Trauma","Neurocirurgia","Neuroclínica","Ortopedia","Psiquiatria","Oftalmologia"];
  for (var e = 0; e < resultado.especialidades.length; e++) {
    resumo += resultado.especialidades[e].destino + " - " + nomesEsp[e] + ": " + resultado.especialidades[e].total + "\n";
  }

  resumo += "\n═══ DERIVADOS ═══\n"
    + "G35 - Total aguarda internação: " + derivados.totalAguardaInter + "\n"
    + "G36 - % >48h aguardando: " + derivados.perc48h + "%\n\n"
    + "Atualizado em: " + Utilities.formatDate(agora, "America/Sao_Paulo", "dd/MM/yyyy HH:mm:ss");

  Logger.log(resumo);
  SpreadsheetApp.getUi().alert(resumo);
}

// ════════════════════════════════════════════════════════════
// ATUALIZAÇÃO AUTOMÁTICA (TRIGGER DE TEMPO)
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

  var resultado = processarDadosOrigem(abaOrigem);
  if (!resultado) return;

  var crossAguardando = contarCrossAguardando();
  var derivados = gravarResultados(dashboard, resultado, crossAguardando);

  var agora = new Date();
  Logger.log("Importação automática:"
    + " C13=" + resultado.totalSetor3
    + " C16=" + resultado.totalSetor2
    + " C20=" + resultado.portaMaior24h
    + " C21=" + resultado.portaMenor24h
    + " C23=" + crossAguardando
    + " C24=" + resultado.permMaior24h
    + " C42=" + resultado.semSolicLeito
    + " C43=" + resultado.permMaior48h
    + " C44=" + resultado.totalLeitoSolic
    + " G35=" + derivados.totalAguardaInter
    + " G36=" + derivados.perc48h + "%"
    + " Especialidades importadas"
    + " em " + Utilities.formatDate(agora, "America/Sao_Paulo", "dd/MM/yyyy HH:mm:ss"));
}
