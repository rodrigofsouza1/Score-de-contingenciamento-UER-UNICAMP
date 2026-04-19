// ════════════════════════════════════════════════════════════
// CONFIGURACAO
// ════════════════════════════════════════════════════════════

var ADMIN_EMAILS = [
  "r157205@dac.unicamp.br"
];

var CELULAS_LIVRES = [
  "C3","C4","C5","C6","C7","C8","C9","C10","C11","C12",
  "C13","C14","C15","C16","C17","C18","C19","C20","C21","C22",
  "C23","C24","C25","C27","C28","C29","C30","C31","C35",
  "C36","C37"
];

// ════════════════════════════════════════════════════════════
// INSTALAR TRIGGER (rodar UMA VEZ como admin)
// ════════════════════════════════════════════════════════════

function instalarTriggerProtecao() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "onEditProtecao") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  ScriptApp.newTrigger("onEditProtecao")
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();

  SpreadsheetApp.getUi().alert("Trigger de proteção instalado com sucesso!");
}

function removerTriggerProtecao() {
  var triggers = ScriptApp.getProjectTriggers();
  var removidos = 0;
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "onEditProtecao") {
      ScriptApp.deleteTrigger(triggers[i]);
      removidos++;
    }
  }
  SpreadsheetApp.getUi().alert("Triggers removidos: " + removidos);
}

// ════════════════════════════════════════════════════════════
// TRIGGER: REVERTE EDICOES NAO AUTORIZADAS
// ════════════════════════════════════════════════════════════

function onEditProtecao(e) {
  try {
    if (!e || !e.range) return;

    var sheet = e.range.getSheet();
    var nomeAba = sheet.getName();
    var usuario = Session.getActiveUser().getEmail();

    // Admin pode tudo
    if (ADMIN_EMAILS.indexOf(usuario) >= 0) return;

    // DASHBOARD: verifica se celula e permitida
    if (nomeAba === "Dashboard") {
      var startRow = e.range.getRow();
      var startCol = e.range.getColumn();
      var numRows = e.range.getNumRows();
      var numCols = e.range.getNumColumns();
      var editouProtegida = false;

      for (var r = 0; r < numRows; r++) {
        for (var c = 0; c < numCols; c++) {
          var cellRef = sheet.getRange(startRow + r, startCol + c).getA1Notation();
          if (CELULAS_LIVRES.indexOf(cellRef) === -1) {
            editouProtegida = true;
            break;
          }
        }
        if (editouProtegida) break;
      }

      if (editouProtegida) {
        if (e.oldValue !== undefined) {
          e.range.setValue(e.oldValue);
        } else {
          e.range.clearContent();
        }
        Logger.log("REVERTIDO: " + usuario + " em " + e.range.getA1Notation());
      }
      return;
    }

    // DEMAIS ABAS: reverte tudo
    if (nomeAba === "Variaveis" || nomeAba === "HISTORICO" || nomeAba === "SIMULACOES") {
      if (e.oldValue !== undefined) {
        e.range.setValue(e.oldValue);
      } else {
        e.range.clearContent();
      }
      Logger.log("REVERTIDO: " + usuario + " em " + nomeAba + "!" + e.range.getA1Notation());
    }

  } catch(err) {
    Logger.log("Erro no trigger: " + err.message);
  }
}

// ════════════════════════════════════════════════════════════
// FUNCOES SEGURAS PARA OS BOTOES
// ════════════════════════════════════════════════════════════

function calcularScoreSeguro() {
  importarIndicadoresAuto();  // ← Atualiza indicadores da planilha externa antes de calcular
  calcularScore();
}
function gravarEExcluirSeguro() { gravarEExcluir(); }
function gerarPDFSeguro() { gerarPDF(); }
function aplicarMelhoriasSeguro() { aplicarMelhorias(); }
function simularCenariosSeguro() { simularCenarios(); }
