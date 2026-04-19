function calcularScore() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dash = ss.getSheetByName("Dashboard");
  var vars = ss.getSheetByName("Variaveis");

  var val = dash.getRange("C3:C24").getValues().flat().map(Number);
  var min = vars.getRange("C3:C24").getValues().flat().map(Number);
  var rotina = vars.getRange("D3:D24").getValues().flat().map(Number);
  var max = vars.getRange("E3:E24").getValues().flat().map(Number);

  var CONFIG = {
    capacidadeTotal: 100,
    capacidadeS4: 9,
    maxVMS4: 4,
    capacidadeS2: 18,
    maxO2S2: 6,
    maxMonitS2: 6,
    capacidadeEmerg1: 6,
    capacidadeEmerg2: 6,
    pesoAssistencial: 0.50,
    pesoSistemico: 0.25,
    pesoEstrutural: 0.25,
    rotinaVM: 4,
    rotinaCross: 5,
    limiarAzul: 15,
    limiarVerde: 30,
    limiarAmarelo: 50,
    limiarLaranja: 70,
    limiarVermelho: 85
  };

  function pesoVM(vmEmerg, vmSetor4) {
    var qtdEfetiva = vmEmerg + vmSetor4 * 0.7;
    if (qtdEfetiva <= 0) return 0;
    if (qtdEfetiva <= 2) return qtdEfetiva * 12;
    if (qtdEfetiva <= 5) return 24 + Math.pow(qtdEfetiva - 2, 1.6) * 10;
    if (qtdEfetiva <= 10) return 24 + Math.pow(3, 1.6) * 10 + (qtdEfetiva - 5) * 15;
    if (qtdEfetiva <= 16) return 24 + Math.pow(3, 1.6) * 10 + 75 + (qtdEfetiva - 10) * 20;
    return 24 + Math.pow(3, 1.6) * 10 + 75 + 120 + (qtdEfetiva - 16) * 25;
  }

  function pesoLotacaoEmerg(semVMEmerg, totalS4, totalS3) {
    if (semVMEmerg <= 0) return 0;
    var base = semVMEmerg * 5;
    var ocupS4 = totalS4 / 9;
    var ocupS3 = totalS3 / 18;
    var fatorContexto = 0.4 + (ocupS4 * 0.6) + (ocupS3 * 0.5);
    fatorContexto = Math.min(fatorContexto, 1.8);
    return base * fatorContexto;
  }

  function pesoSemLeito(qtd) {
    if (qtd <= 0) return 0;
    if (qtd <= 10) return qtd * 5;
    if (qtd <= 30) return 50 + (qtd - 10) * 3;
    return 110 + Math.sqrt(qtd - 30) * 10;
  }

  function pesoVolume(totalPac, capacidade) {
    var ratio = totalPac / capacidade;
    if (ratio <= 0.5) return 0;
    if (ratio <= 0.8) return (ratio - 0.5) * 50;
    if (ratio <= 1.0) return 15 + (ratio - 0.8) * 150;
    return 45 + (ratio - 1.0) * 200;
  }

  function calcularComponentes(v) {
    var vmEmerg = v[0] + v[2];
    var vmSetor4 = v[4];
    var semVMEmerg = v[1] + v[3];
    var totalS4 = v[4] + v[5];
    var totalS3 = v[10];
    var totalPacientes = v[0] + v[1] + v[2] + v[3] + totalS4 + v[10] + v[13] + v[14] + v[15] + v[16];

    var assist = 0;
    assist += pesoVM(vmEmerg, vmSetor4);
    assist += pesoLotacaoEmerg(semVMEmerg, totalS4, totalS3);
    assist += (v[8] + v[11]) * 2;
    assist += v[9] * 7;
    assist += v[7] * 6;
    assist += v[6] * 3;
    assist += v[12] * 3;

    var estrut = 0;
    estrut += pesoVolume(totalPacientes, CONFIG.capacidadeTotal);
    estrut += v[14] * 3;
    estrut += v[15] * 5;
    estrut += v[16] * 7;
    estrut += v[19] * 4;

    var sist = 0;
    sist += v[17] * 4;
    sist += v[18] * 1;
    sist += v[20] * 5;
    sist += pesoSemLeito(v[21]);

    return { assist: assist, estrut: estrut, sist: sist, totalPac: totalPacientes };
  }

  var minCalc = calcularComponentes(min);
  var rotinaCalc = calcularComponentes(rotina);
  var maxCalc = calcularComponentes(max);
  var atualCalc = calcularComponentes(val);

  function bruto(obj) { return obj.assist + obj.estrut + obj.sist; }
  function ponderado(obj) {
    return (obj.assist * CONFIG.pesoAssistencial) + (obj.estrut * CONFIG.pesoEstrutural) + (obj.sist * CONFIG.pesoSistemico);
  }

  var brutoMin = bruto(minCalc);
  var brutoMax = bruto(maxCalc);
  var pondMin = ponderado(minCalc);
  var pondMax = ponderado(maxCalc);
  var brutoAtual = bruto(atualCalc);
  var pondAtual = ponderado(atualCalc);

  function norm(v, minVal, maxVal) {
    if (maxVal <= minVal) return 50;
    return Math.max(0, Math.min(100, ((v - minVal) / (maxVal - minVal)) * 100));
  }

  var linear = norm(brutoAtual, brutoMin, brutoMax);
  var linearPond = norm(pondAtual, pondMin, pondMax);

  var mid = (pondMax + pondMin) / 2;
  var rangeVal = pondMax - pondMin;
  var k = (rangeVal > 0) ? 6 / rangeVal : 0.05;
  var logistico = 100 / (1 + Math.exp(-k * (pondAtual - mid)));

  // SCORES CRITICOS

  function scoreVMCritico(vmEmerg, vmSetor4) {
    var sEmerg = 0;
    if (vmEmerg <= 2) sEmerg = vmEmerg * 15;
    else if (vmEmerg <= 6) sEmerg = 30 + (vmEmerg - 2) * 12;
    else if (vmEmerg <= 12) sEmerg = 78 + (vmEmerg - 6) * 3.7;
    else sEmerg = 100;
    var sS4 = 0;
    if (vmSetor4 <= 2) sS4 = vmSetor4 * 12;
    else if (vmSetor4 <= 4) sS4 = 24 + (vmSetor4 - 2) * 18;
    else sS4 = Math.min(60 + (vmSetor4 - 4) * 10, 100);
    var combinado = Math.max(sEmerg, sS4 * 0.7);
    if (sEmerg > 30 && sS4 > 20) combinado += Math.min((sEmerg + sS4) * 0.1, 20);
    return Math.min(combinado, 100);
  }

  function scoreLotacaoEmergCritico(semVMEmerg, totalS4, totalS3, capS4) {
    if (semVMEmerg <= 0) return 0;
    var capEmerg = CONFIG.capacidadeEmerg1 + CONFIG.capacidadeEmerg2;
    var ratioEmerg = semVMEmerg / capEmerg;
    var sBase = 0;
    if (ratioEmerg <= 0.5) sBase = ratioEmerg * 30;
    else if (ratioEmerg <= 0.8) sBase = 15 + (ratioEmerg - 0.5) * 80;
    else if (ratioEmerg <= 1.0) sBase = 39 + (ratioEmerg - 0.8) * 155;
    else sBase = Math.min(70 + (ratioEmerg - 1.0) * 100, 100);
    var ocupS4 = totalS4 / capS4;
    var ocupS3 = (totalS3 > 0) ? Math.min(totalS3 / 15, 1.5) : 0;
    var fator = 0.4 + (ocupS4 * 0.4) + (ocupS3 * 0.3);
    fator = Math.max(0.4, Math.min(fator, 1.3));
    return Math.min(sBase * fator, 100);
  }

  function scoreOcupacao(total, capacidade) {
    if (capacidade <= 0) return 100;
    var ratio = total / capacidade;
    if (ratio <= 0.5) return ratio * 40;
    if (ratio <= 0.8) return 20 + ((ratio - 0.5) / 0.3) * 30;
    if (ratio <= 1.0) return 50 + ((ratio - 0.8) / 0.2) * 30;
    return Math.min(80 + ((ratio - 1.0) / 0.3) * 20, 100);
  }

  function scoreOcupacaoS4(vmS4, semVMS4, capTotal, capVM) {
    var totalS4 = vmS4 + semVMS4;
    var ratioTotal = totalS4 / capTotal;
    var ratioVM = (capVM > 0) ? vmS4 / capVM : 0;
    var sTotal = 0;
    if (ratioTotal <= 0.7) sTotal = ratioTotal * 30;
    else if (ratioTotal <= 1.0) sTotal = 21 + ((ratioTotal - 0.7) / 0.3) * 49;
    else sTotal = Math.min(70 + (ratioTotal - 1.0) * 100, 100);
    var sVM = 0;
    if (ratioVM <= 0.75) sVM = ratioVM * 30;
    else if (ratioVM <= 1.0) sVM = 22.5 + ((ratioVM - 0.75) / 0.25) * 37.5;
    else sVM = Math.min(60 + (ratioVM - 1.0) * 100, 100);
    return (sTotal * 0.6) + (sVM * 0.4);
  }

  function scoreBoarding(semLeito24h, perm24h) {
    var s = 0;
    if (semLeito24h <= 5) s += semLeito24h * 8;
    else if (semLeito24h <= 20) s += 40 + (semLeito24h - 5) * 3;
    else s += Math.min(85 + (semLeito24h - 20) * 0.2, 100);
    s += Math.min(perm24h * 5, 40);
    return Math.min(s, 100);
  }

  function scoreAcomodacao(macas, poltronas, consultorios) {
    if (macas + poltronas + consultorios <= 0) return 0;
    return Math.min((macas * 2 + poltronas * 4 + consultorios * 6), 100);
  }

  function scoreCross(qtd, rotinaRef) {
    if (rotinaRef <= 0) return Math.min(qtd * 10, 100);
    var ratio = qtd / rotinaRef;
    if (ratio <= 1.0) return ratio * 15;
    if (ratio <= 2.0) return 15 + (ratio - 1) * 45;
    return Math.min(60 + (ratio - 2) * 40, 100);
  }

  var vmEmerg = val[0] + val[2];
  var vmSetor4 = val[4];
  var semVMEmerg = val[1] + val[3];
  var semVMSetor4 = val[5];
  var totalS4 = vmSetor4 + semVMSetor4;
  var totalS3 = val[10];
  var totalPacientesReal = atualCalc.totalPac;

  var critVM       = scoreVMCritico(vmEmerg, vmSetor4);
  var critLotEmerg = scoreLotacaoEmergCritico(semVMEmerg, totalS4, totalS3, CONFIG.capacidadeS4);
  var critOcup     = scoreOcupacao(totalPacientesReal, CONFIG.capacidadeTotal);
  var critS4       = scoreOcupacaoS4(vmSetor4, semVMSetor4, CONFIG.capacidadeS4, CONFIG.maxVMS4);
  var critBoard    = scoreBoarding(val[21], val[17]);
  var critAcom     = scoreAcomodacao(val[14], val[15], val[16]);
  var critCross    = scoreCross(val[20], CONFIG.rotinaCross);

  var todosCriticos = [critVM, critLotEmerg, critOcup, critS4, critBoard, critAcom, critCross];
  var sorted = todosCriticos.slice().sort(function(a, b) { return b - a; });
  var maxCritico = sorted[0];
  var p90Critico = sorted[1] || maxCritico;
  var mediaCritico = todosCriticos.reduce(function(a, b) { return a + b; }, 0) / todosCriticos.length;
  var scoreCritico = (maxCritico * 0.50) + (p90Critico * 0.30) + (mediaCritico * 0.20);

  var pesoFinalCritico, pesoFinalLinear;
  if (scoreCritico > 70) { pesoFinalCritico = 0.70; pesoFinalLinear = 0.30; }
  else if (scoreCritico > 50) { pesoFinalCritico = 0.55; pesoFinalLinear = 0.45; }
  else { pesoFinalCritico = 0.45; pesoFinalLinear = 0.55; }

  var scoreFinal = Math.max(0, Math.min(100, (scoreCritico * pesoFinalCritico) + (linearPond * pesoFinalLinear)));

  var classificacao;
  if (scoreFinal <= CONFIG.limiarAzul) classificacao = "ROTINA";
  else if (scoreFinal <= CONFIG.limiarVerde) classificacao = "NÍVEL 1";
  else if (scoreFinal <= CONFIG.limiarAmarelo) classificacao = "NÍVEL 2";
  else if (scoreFinal <= CONFIG.limiarLaranja) classificacao = "NÍVEL 3";
  else if (scoreFinal <= CONFIG.limiarVermelho) classificacao = "NÍVEL 4";
  else classificacao = "NÍVEL 5";

  // FLAGS
  var flags = [];
  if (scoreFinal > CONFIG.limiarVerde) {
    var indicadores = [
      { nome: "VM",         score: critVM,       valor: "Emerg:" + vmEmerg + " S4:" + vmSetor4 },
      { nome: "Lot.Emerg",  score: critLotEmerg, valor: "SemVM:" + semVMEmerg + " (S4:" + totalS4 + " S3:" + totalS3 + ")" },
      { nome: "Ocupacao",   score: critOcup,     valor: totalPacientesReal + "/" + CONFIG.capacidadeTotal },
      { nome: "Setor 4",    score: critS4,       valor: totalS4 + "/9 VM:" + vmSetor4 + "/4" },
      { nome: "Boarding",   score: critBoard,    valor: "SLei:" + val[21] + " P24h:" + val[17] },
      { nome: "Acomodacao", score: critAcom,     valor: "Mac:" + val[14] + " Pol:" + val[15] + " Con:" + val[16] },
      { nome: "Cross",      score: critCross,    valor: "6h:" + val[20] }
    ];
    indicadores.sort(function(a, b) { return b.score - a.score; });
    for (var i = 0; i < indicadores.length; i++) {
      if (indicadores[i].score >= 20) flags.push(indicadores[i]);
    }
  }

  // OUTPUT VARIAVEIS
  vars.getRange("C31").setValue(brutoMin);
  vars.getRange("D31").setValue(bruto(rotinaCalc));
  vars.getRange("E31").setValue(brutoMax);
  vars.getRange("C32").setValue(pondMin);
  vars.getRange("D32").setValue(ponderado(rotinaCalc));
  vars.getRange("E32").setValue(pondMax);

  // OUTPUT DASHBOARD
  dash.getRange("G5").setValue(brutoAtual.toFixed(1));
  dash.getRange("G6").setValue(pondAtual.toFixed(1));
  dash.getRange("G7").setValue(linear.toFixed(1) + "%");
  dash.getRange("G8").setValue(linearPond.toFixed(1) + "%");
  dash.getRange("G9").setValue(logistico.toFixed(1) + "%");
  dash.getRange("G10").setValue(scoreCritico.toFixed(1) + "%");
  dash.getRange("G11").setValue(scoreFinal.toFixed(1) + "%");
  dash.getRange("E12").setValue(classificacao);

  // FLAGS NO DASHBOARD
  dash.getRange("E14:G20").clearContent();
  dash.getRange("E14:G20")
      .setBackground("#0f0f1a")
      .setFontFamily("Arial")
      .setBorder(false, true, true, true, false, false, "#4a4a6a", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  if (flags.length > 0) {
    dash.getRange("E14").setValue("⚠ ALERTAS CRITICOS")
        .setFontWeight("bold").setFontColor("#ef5350").setFontSize(9);
    for (var f = 0; f < flags.length && f < 6; f++) {
      var linhaFlag = 15 + f;
      var icone = flags[f].score >= 50 ? "🔴" : "🟡";
      var corTexto = flags[f].score >= 50 ? "#ef5350" : "#ffa726";
      dash.getRange("E" + linhaFlag).setValue(icone + " " + flags[f].nome + " (" + flags[f].score.toFixed(0) + ")")
          .setFontSize(8).setFontWeight("bold").setFontColor(corTexto);
      dash.getRange("G" + linhaFlag).setValue(flags[f].valor)
          .setFontSize(8).setFontColor("#888888");
    }
  }

  atualizarIndicadoresVisuais();

  // LOG
  Logger.log("══════ SCORE DE CONTINGENCIA UER ══════");
  Logger.log("Total pacientes: " + totalPacientesReal + "/" + CONFIG.capacidadeTotal);
  Logger.log("VM Emerg: " + vmEmerg + " | VM S4: " + vmSetor4 + " | critVM: " + critVM.toFixed(1));
  Logger.log("Sem VM Emerg: " + semVMEmerg + " | critLotEmerg: " + critLotEmerg.toFixed(1));
  Logger.log("Percepcao avaliador: " + dash.getRange("C25").getDisplayValue());
  Logger.log("Score Critico: " + scoreCritico.toFixed(1) + "%");
  Logger.log("Linear Pond: " + linearPond.toFixed(1) + "%");
  if (flags.length > 0) {
    Logger.log("FLAGS:");
    for (var fl = 0; fl < flags.length; fl++) {
      Logger.log("  " + flags[fl].nome + "=" + flags[fl].score.toFixed(1) + " | " + flags[fl].valor);
    }
  }
  Logger.log("══ SCORE FINAL: " + scoreFinal.toFixed(1) + "% -> " + classificacao + " ══");
}
