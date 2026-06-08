// --- MAPEAMENTO DE COLUNAS (Índices começam em 0) ---
const COL_DATA_INICIO = 1;  // B
const COL_ALUNO       = 3;  // D
const COL_PLANO       = 4;  // E
const COL_FREQ        = 5;  // F
const COL_CONTRATADAS = 10; // K
const COL_RESTANTES   = 27; // AB
const COL_FEITAS_AD   = 29; // AD 
const COL_STATUS      = 11; // L

/**
 * Função Auxiliar de Cálculos
 */
function calcularAulasContratadas(plano, freq) {
  const p = String(plano || "").toUpperCase();
  const f = parseInt(freq) || 1;

  if (p.includes("5 AULAS")) return 5;
  if (p.includes("12 AULAS")) return 12;
  if (p.includes("6 MESES")) return 24 * f;
  if (p.includes("12 MESES")) return 48 * f;
  if (p.includes("MENSALIDADE")) return ""; 
  
  return ""; 
}

/**
 * 1. ADICIONAR ALUNO (Web App)
 */
function CONTRATOS_adicionarDaAgenda(nomeDoAluno, planoEscolhido, freqEscolhida) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName("Contratos");
    const hoje = new Date();
    
    const plano = String(planoEscolhido || "").toUpperCase();
    const freq = parseInt(freqEscolhida) || 1;
    const contratadas = calcularAulasContratadas(plano, freq);
    
    const aulasFeitas = 1; 
    const restantes = (contratadas !== "") ? contratadas - aulasFeitas : "";

    let novaLinha = [];
    novaLinha[0] = "";                        // A
    novaLinha[1] = hoje;                      // B: COL_DATA_INICIO
    novaLinha[2] = "";                        // C
    novaLinha[3] = nomeDoAluno.toUpperCase(); // D: COL_ALUNO
    novaLinha[4] = plano;                     // E: COL_PLANO
    novaLinha[5] = freq;                      // F: COL_FREQ

    sh.appendRow(novaLinha);
    let lastRow = sh.getLastRow();

    sh.getRange(lastRow, 11).setValue(contratadas); // K
    sh.getRange(lastRow, 12).setValue("ativo");      // L
    sh.getRange(lastRow, 28).setValue(restantes);   // AB
    sh.getRange(lastRow, 30).setValue(aulasFeitas);  // AD

    return "✅ " + nomeDoAluno.toUpperCase() + " ADICIONADO!";
  } catch (e) { return "❌ Erro ao adicionar: " + e.toString(); }
}

/**
 * 2. RENOVAR PLANO (Web App)
 */
function CONTRATOS_renovacaoManual(nomeDoAluno, planoEscolhido, freqEscolhida) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName("Contratos");
    const data = sh.getDataRange().getValues();
    
    const busca = nomeDoAluno.trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
    let linha = -1;

    for (let i = 1; i < data.length; i++) {
      let nomePlanilha = String(data[i][3] || "").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
      if (nomePlanilha === busca) {
        linha = i + 1;
        break;
      }
    }

    if (linha === -1) return "❌ Aluno não encontrado.";

    const hoje = new Date();
    const plano = String(planoEscolhido || "").toUpperCase();
    const freq = parseInt(freqEscolhida) || 1;
    const contratadas = calcularAulasContratadas(plano, freq);
    
    const aulasFeitas = 1; 
    const restantes = (contratadas !== "") ? contratadas - aulasFeitas : "";

    sh.getRange(linha, 2).setValue(hoje);          
    sh.getRange(linha, 5).setValue(plano);         
    sh.getRange(linha, 6).setValue(freq);          
    sh.getRange(linha, 11).setValue(contratadas);  
    sh.getRange(linha, 12).setValue("ativo");      
    sh.getRange(linha, 28).setValue(restantes);    
    sh.getRange(linha, 30).setValue(aulasFeitas);  

    return "🔄 RENOVAÇÃO OK: " + nomeDoAluno.toUpperCase();
  } catch (e) { return "❌ Erro na renovação: " + e.toString(); }
}

/**
 * 3. EVENTOS FUTUROS (Web App)
 */
function APP_verEventosFuturos(nomeDoAluno) {
  try {
    const nomeBusca = nomeDoAluno.trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
    if (!nomeBusca) return { erro: "Nome inválido" };

    const agenda = CalendarApp.getDefaultCalendar();
    const hoje = new Date();
    const trintaDias = new Date(Date.now() + 30 * 24 * 60 * 60 * 1000);
    const eventos = agenda.getEvents(hoje, trintaDias);
    
    const regexBusca = new RegExp("\\b" + nomeBusca + "\\b", "i");
    let listaParaApp = [];

    eventos.forEach(e => {
      let tituloOriginal = e.getTitle();
      let tNorm = tituloOriginal.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
      
      if (regexBusca.test(tNorm)) {
        let inicio = e.getStartTime();
        listaParaApp.push({
          data: Utilities.formatDate(inicio, "GMT-3", "dd/MM/yy"),
          hora: Utilities.formatDate(inicio, "GMT-3", "HH:mm"),
          titulo: tituloOriginal
        });
      }
    });

    return { 
      aluno: nomeDoAluno.toUpperCase(), 
      eventos: listaParaApp 
    };
  } catch (e) { return { erro: e.toString() }; }
}
/**
 * BUSCA HISTÓRICO COMPLETO (Soma da Planilha + Agenda)
 */
function APP_verHistoricoAulas(nomeDoAluno) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName("Contratos");
    const dataC = sh.getDataRange().getValues();
    const busca = nomeDoAluno.trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
    
    let dadosAluno = null;
    let dataInicioPlano = null;

    for (let i = 1; i < dataC.length; i++) {
      let nomePlanilha = String(dataC[i][3] || "").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
      if (nomePlanilha.includes(busca)) { 
        dataInicioPlano = dataC[i][1]; // Coluna B (Data de Início)
        dadosAluno = {
          nome: dataC[i][3],
          plano: dataC[i][4],
          restantes: dataC[i][27] || 0,
          feitas: dataC[i][29] || 0,
          status: String(dataC[i][11]).toUpperCase(),
          inicio: Utilities.formatDate(new Date(dataInicioPlano), "GMT-3", "dd/MM/yy")
        };
        break;
      }
    }

    if (!dadosAluno) return { erro: "Aluno não encontrado." };

    // Busca aulas com Check na Agenda APENAS desde a data de início
    const agenda = CalendarApp.getDefaultCalendar();
    const eventos = agenda.getEvents(new Date(dataInicioPlano), new Date()); 
    const regexBusca = new RegExp("\\b" + busca + "\\b", "i");
    const checkPat = /✅|✔|☑|✓/;
    let historicoAgenda = [];

    eventos.forEach(e => {
      let tit = e.getTitle();
      let titNorm = tit.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
      if (regexBusca.test(titNorm) && checkPat.test(tit)) {
        historicoAgenda.push({
          data: Utilities.formatDate(e.getStartTime(), "GMT-3", "dd/MM"),
          hora: Utilities.formatDate(e.getStartTime(), "GMT-3", "HH:mm")
        });
      }
    });

    return { resumo: dadosAluno, aulasFeitas: historicoAgenda };
  } catch (e) { return { erro: e.toString() }; }
}