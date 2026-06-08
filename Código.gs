/****************************************************************
 * Studio GS — GESTÃO BLINDADA v6.0
 * Correção: Sincronização de nomes de funções e fechamento de chaves
 ****************************************************************/

const CALENDAR_ID = 'primary';
const SHEET_NAME  = 'Contratos';

/**
 * MENU DO STUDIO GS — GESTÃO BLINDADA
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('Financeiro do Studio')
    // --- ATUALIZAÇÕES AUTOMÁTICAS ---
    .addItem('⭐ Atualizar MENSALISTAS (Agenda -> AD/AB)', 'atualizar_MENSALISTAS_Agenda_Saldo')
    .addItem('⭐ Atualizar PLANOS LONGOS (Z + Agenda)', 'atualizarPlanos_BRUTO')
    .addSeparator()
    
    // --- GESTÃO DE DIAS (EXCLUSIVO MENSALISTAS) ---
    .addItem('📅 Sincronizar Dias da Agenda (Coluna AL)', 'configurarMensalistasPelaAgenda')
    .addSeparator()
    
    // --- FERRAMENTAS DE ALUNOS ---
    .addItem('➕ Adicionar Aluno Novo (Agenda)', 'CONTRATOS_adicionarDaAgenda')
    .addItem('🔍 Verificar Novos na Agenda', 'GS_verificarNovosAlunosComPresenca')
    .addItem('🔄 Renovar Linha Selecionada', 'CONTRATOS_renovacaoManual')
    .addItem('🔍 Gerar Relatório Individual', 'verHistoricoAulasFeitas')
    .addSeparator()
    
    // --- CONSULTAS E CONFIG ---
    .addItem('🔍 Ver Futuro (Aba Futuros)', 'verEventosFuturosAluno')
    .addItem('🚨 Verificar Alunos Zerados', 'verificarRenovacoesZero')
    .addItem('⚙️ Configurar Google Tasks', 'setupGoogleTasks')
    .addToUi();
}
function atualizarPlanos_BRUTO() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Contratos");
  const data = sheet.getDataRange().getValues();
  
  // --- MAPEAMENTO DE COLUNAS ---
  const COL_DATA_INICIO = 1;  // B
  const COL_ALUNO       = 3;  // D
  const COL_PLANO       = 4;  // E
  const COL_CONTRATADAS = 10; // K
  const COL_RESTANTES   = 27; // AB
  const COL_FEITAS_AD   = 29; // AD 
  const COL_STATUS      = 11; // L
  const COL_PAPEL_Z     = 25; // Z 

  const agenda = CalendarApp.getDefaultCalendar();
  const checkPat = /✅|✔|☑|✓/;
  const hoje = new Date();
  const dataCutover = new Date(2025, 9, 1); 

  for (let i = 1; i < data.length; i++) {
    const nomeOriginal = String(data[i][COL_ALUNO] || "").trim();
    if (!nomeOriginal) continue;

    const nomeBusca = nomeOriginal.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
    const planoTexto = String(data[i][COL_PLANO] || "").toUpperCase();
    const status = String(data[i][COL_STATUS] || "").toLowerCase();
    const dataInicioPlano = data[i][COL_DATA_INICIO];

    const ehPlanoLongo = !planoTexto.includes("MENSALIDADE") && !planoTexto.includes("AVULSA");

    if (status === "ativo" && ehPlanoLongo && dataInicioPlano instanceof Date) {
      let contagemAgenda = 0;
      let dataBuscaInicio = (dataInicioPlano < dataCutover) ? dataCutover : dataInicioPlano;
      
      const eventos = agenda.getEvents(dataBuscaInicio, hoje);
      
      eventos.forEach(e => {
        let tituloOriginal = e.getTitle();
        let t = tituloOriginal.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
        
        if (checkPat.test(tituloOriginal)) {
           let regexNome = new RegExp("\\b" + nomeBusca + "\\b", "i");
           if (regexNome.test(t)) {
             contagemAgenda++;
           }
        }
      });

      let aulasPapel = (dataInicioPlano < dataCutover) ? (Number(data[i][COL_PAPEL_Z]) || 0) : 0;
      let totalFeitas = contagemAgenda + aulasPapel;
      let contratadas = Number(data[i][COL_CONTRATADAS]) || 0;
      let restantes = contratadas - totalFeitas;

      // 1. Atualiza a Planilha
      sheet.getRange(i + 1, COL_FEITAS_AD + 1).setValue(totalFeitas); 
      sheet.getRange(i + 1, COL_RESTANTES + 1).setValue(restantes);
      sheet.getRange(i + 1, COL_RESTANTES + 1).setBackground(restantes <= 0 ? "#ea9999" : "#cfe2f3");

      // 2. 🚀 CRIAÇÃO DA TASK (Se faltar 2 aulas ou menos)
      if (restantes <= 2 && restantes > -5) { // -5 para não criar task de contratos muito antigos/vencidos
        criarTaskRenovacao(nomeOriginal, planoTexto, restantes);
      }
    }
  }
}

// Função de apoio para criar a Task
function criarTaskRenovacao(aluno, plano, saldo) {
  try {
    const tituloTask = "⚠️ RENOVAR: " + aluno + " (Plano: " + plano + ")";
    const notas = "O aluno(a) " + aluno + " está com apenas " + saldo + " aulas restantes no plano " + plano + ". Favor entrar em contato para renovação.";
    
    // Verifica se o serviço de Tasks está ativo
    const taskLists = Tasks.Tasklists.list();
    if (taskLists.items && taskLists.items.length > 0) {
      const listId = taskLists.items[0].id; // Pega a sua lista principal de tarefas
      
      // Cria o objeto da tarefa
      const newTask = {
        title: tituloTask,
        notes: notas
      };
      
      Tasks.Tasks.insert(newTask, listId);
      console.log("Task criada para: " + aluno);
    }
  } catch (e) {
    console.log("Erro ao criar task: " + e.message + ". Certifique-se de ativar o Google Tasks no menu '+' do Apps Script.");
  }
}
/**
 * ATUALIZAÇÃO DE MENSALISTAS (Sincronização com Agenda + Abatimento de AG)
 * Esta função deve ser rodada para atualizar os saldos baseados nos checks da agenda.
 */
function atualizar_MENSALISTAS_Agenda_Saldo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Contratos");
  const data = sheet.getDataRange().getValues();
  
  const COL_INICIO      = 1;  // B
  const COL_ALUNO       = 3;  // d
  const COL_PLANO       = 4;  // E
  const COL_K           = 10; // K (Contratadas)
  const COL_RESTANTES   = 27; // AB (Saldo Mês)
  const COL_FEITAS_AD   = 29; // AD (Feitas)
  const COL_RECUPERACAO = 32; // AG (Saldo Acumulado)

  const agenda = CalendarApp.getDefaultCalendar();
  const checkPat = /✅|✔|☑|✓/;
  const hoje = new Date();

  for (let i = 1; i < data.length; i++) {
    const plano = String(data[i][COL_PLANO] || "").toLowerCase();
    
    if (plano.includes("mensalidade")) {
      const nomeOriginal = String(data[i][COL_ALUNO] || "").trim();
      const dataInicio = data[i][COL_INICIO];
      const contratadas = Number(data[i][COL_K]) || 0;
      let saldoRecAtual = Number(data[i][COL_RECUPERACAO]) || 0;

      if (nomeOriginal !== "" && dataInicio instanceof Date) {
        const nomeBusca = nomeOriginal.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
        let contagemCheck = 0;

        // Busca checks na agenda desde o início do plano até hoje
        const eventos = agenda.getEvents(dataInicio, hoje);
        eventos.forEach(e => {
          let t = e.getTitle().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
          if (t.includes(nomeBusca) && checkPat.test(e.getTitle())) contagemCheck++;
        });

        // --- LÓGICA DE RECUPERAÇÃO INTELIGENTE ---
        let saldoMes = 0;
        let novoSaldoRec = saldoRecAtual;

        if (contagemCheck > contratadas) {
          // Se o aluno fez aulas extras, abate do saldo acumulado (AG)
          let excedente = contagemCheck - contratadas;
          saldoMes = 0; 
          novoSaldoRec = Math.max(0, saldoRecAtual - excedente);
          sheet.getRange(i + 1, COL_RECUPERACAO + 1).setBackground("#d9ead3"); // Verde se usou rec
        } else {
          saldoMes = contratadas - contagemCheck;
          sheet.getRange(i + 1, COL_RECUPERACAO + 1).setBackground(null);
        }

        // Gravando os dados na planilha
        sheet.getRange(i + 1, COL_FEITAS_AD + 1).setValue(contagemCheck);
        sheet.getRange(i + 1, COL_RESTANTES + 1).setValue(saldoMes);
        sheet.getRange(i + 1, COL_RESTANTES + 1).setBackground(saldoMes <= 0 ? "#ea9999" : "#cfe2f3");
        sheet.getRange(i + 1, COL_RECUPERACAO + 1).setValue(novoSaldoRec);
      }
    }
  }
}


// --- RESTANTE DAS FUNÇÕES AUXILIARES ---

function fin_relatorioAluno() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAME);
  const ui = SpreadsheetApp.getUi();

  const prompt = ui.prompt('Gerar Relatório', 'Digite o nome do aluno:', ui.ButtonSet.OK_CANCEL);
  if (prompt.getSelectedButton() !== ui.Button.OK) return;
  
  const busca = prompt.getResponseText().trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  if (busca === "") return;

  const data = sh.getDataRange().getValues();
  let linhaEncontrada = -1;

  for (let i = 1; i < data.length; i++) {
    let nomePlanilha = String(data[i][2]).toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
    if (nomePlanilha.includes(busca)) { linhaEncontrada = i; break; }
  }

  if (linhaEncontrada === -1) { ui.alert("❌ Aluno não encontrado."); return; }

  const alunoData = data[linhaEncontrada];
  let relatorio = "📊 *RELATÓRIO - STUDIO GS*\n\n👤 *Aluno:* " + alunoData[2] + "\n📋 *Plano:* " + alunoData[4] + "\n------------------\n✅ *Aulas Feitas:* " + (alunoData[29] || 0) + "\n⏳ *Restantes:* " + (alunoData[27] || 0) + " de " + alunoData[10];
  ui.alert(relatorio);
}

/**
 * RENOVAÇÃO MANUAL POR NOME (BUSCA AUTOMÁTICA)
 * Conectada ao Monitor do Painel
 */
function CONTRATOS_renovacaoManual() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Contratos");
  const ui = SpreadsheetApp.getUi();

  // 1. PERGUNTA O NOME DO ALUNO
  const promptBusca = ui.prompt('Renovação de Contrato', 'Digite o nome do aluno que deseja renovar:', ui.ButtonSet.OK_CANCEL);
  if (promptBusca.getSelectedButton() !== ui.Button.OK) return;
  
  const busca = promptBusca.getResponseText().trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  if (busca === "") return;

  // 2. BUSCA A LINHA DO ALUNO NA ABA CONTRATOS
  const data = sh.getDataRange().getValues();
  let linha = -1;

  for (let i = 1; i < data.length; i++) {
    // Busca na Coluna C (Índice 2)
    let nomePlanilha = String(data[i][2]).toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
    if (nomePlanilha.includes(busca)) { 
      linha = i + 1; // +1 porque as linhas da planilha começam em 1
      break; 
    }
  }

  if (linha === -1) {
    gs_escreverNoMonitor("❌ ERRO: Aluno '" + busca.toUpperCase() + "' não encontrado para renovação.");
    return;
  }

  const alunoNomeCompleto = data[linha - 1][2];
  gs_escreverNoMonitor("🔄 Renovando agora: " + alunoNomeCompleto.toUpperCase());

  // 3. PERGUNTAS DO NOVO PLANO
  const promptPlano = ui.prompt('Renovando: ' + alunoNomeCompleto, 'Tipo (5 aulas, 12 aulas, Mensalidade, 6 meses, 12 meses):', ui.ButtonSet.OK_CANCEL);
  if (promptPlano.getSelectedButton() !== ui.Button.OK) return;
  const plano = promptPlano.getResponseText().toLowerCase().trim();

  let freq = 1;
  if (!plano.includes("aulas")) {
    const promptFreq = ui.prompt('Frequência', 'Vezes por semana (1, 2 ou 3):', ui.ButtonSet.OK_CANCEL);
    freq = parseInt(promptFreq.getResponseText()) || 1;
  }

  // Lógica de cálculo (Valor base: 130)
  const valorBaseAula = 130; 
  let aulasContratadas = 0;
  let meses = 1;
  let desconto = 0;

  if (plano.includes("5 aulas")) { 
    aulasContratadas = 5; 
    desconto = 0.05; 
  } else if (plano.includes("12 aulas")) { 
    aulasContratadas = 12; 
    meses = 2; 
    desconto = 0.15; 
  } else if (plano.includes("6 meses")) { 
    meses = 6; 
    aulasContratadas = freq * 4 * 6; 
    desconto = 0.25; 
  } else if (plano.includes("12 meses")) { 
    meses = 12; 
    aulasContratadas = freq * 4 * 12; 
    desconto = 0.35; 
  } else if (plano.includes("mensal")) {
    meses = 1;
    aulasContratadas = freq * 4; 
    if (freq == 1) desconto = 0.05;
    else if (freq == 2) desconto = 0.10;
    else if (freq == 3) desconto = 0.15;
  }

  const precoUnitario = valorBaseAula * (1 - desconto);

  // 4. GRAVAÇÃO NA PLANILHA
  sh.getRange(linha, 2).setValue(new Date());          // B: DATA_INICIO
  sh.getRange(linha, 5).setValue(plano.toUpperCase()); // E: PLANO
  sh.getRange(linha, 6).setValue(freq);                // F: FREQUENCIA
  sh.getRange(linha, 7).setValue(meses);               // G: DURACAO
  sh.getRange(linha, 9).setValue(desconto);            // I: DESCONTO_%
  sh.getRange(linha, 10).setValue(precoUnitario);      // J: PRECO_UNIT
  sh.getRange(linha, 11).setValue(aulasContratadas);   // K: QTDE_AULAS_CONTR
  sh.getRange(linha, 12).setValue("ativo");            // L: STATUS
  
  // Limpa progresso para o novo ciclo
  sh.getRange(linha, 28).setValue(aulasContratadas);   // AB: SALDO (Inicia cheio)
  sh.getRange(linha, 30).setValue(0);                  // AD: FEITAS (Zera)
  
  // 5. FEEDBACK FINAL NO MONITOR
  let msgFinal = "✅ RENOVAÇÃO CONCLUÍDA!\n\n" +
                 "👤 Aluno: " + alunoNomeCompleto + "\n" +
                 "📋 Plano: " + plano.toUpperCase() + "\n" +
                 "🔢 Aulas: " + aulasContratadas + "\n" +
                 "💰 Preço Unit: R$ " + precoUnitario.toFixed(2);
  
  gs_escreverNoMonitor(msgFinal);
}
function setupGoogleTasks() { SpreadsheetApp.getUi().alert("Serviço de Tasks deve ser ativado no menu '+' à esquerda."); }

/**
 * CONSULTA EVENTOS FUTUROS (ABA ESPECÍFICA + PAINEL ESTILO SITE)
 * Fonte: Montserrat, Tamanho 14
 */
/**
 * CONSULTA EVENTOS FUTUROS (CORRIGIDA)
 * Fonte: Montserrat, Tamanho 14 | Saída: Aba Futuros + Painel
 */
function verEventosFuturosAluno() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const shPainel = ss.getSheetByName("PAINEL");
  
  // 1. Prepara a aba de destino específica
  let sheetFuturo = ss.getSheetByName("Futuros_Aluno_Pilates") || ss.insertSheet("Futuros_Aluno_Pilates");
  
  // 2. Pergunta o nome do aluno
  const resposta = ui.prompt('Consultar Agenda Futura', 'Digite o nome do aluno:', ui.ButtonSet.OK_CANCEL);
  if (resposta.getSelectedButton() !== ui.Button.OK) return;
  const nome = resposta.getResponseText().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  if (!nome) return;

  // 3. Limpeza do Monitor no Painel (D4:H40)
  const areaMonitor = shPainel.getRange("D4:H40");
  areaMonitor.clearContent()
             .setBackground(null)
             .setFontWeight("normal")
             .setFontStyle("normal") // Reseta o itálico de buscas anteriores
             .setFontFamily("Montserrat")
             .setFontSize(14)
             .setVerticalAlignment("middle");

  // 4. Busca eventos na Agenda (Próximos 30 dias)
  const agenda = CalendarApp.getDefaultCalendar();
  const hoje = new Date();
  const trintaDias = new Date(Date.now() + 30 * 24 * 60 * 60 * 1000);
  const eventos = agenda.getEvents(hoje, trintaDias);
  
  let dadosAba = [["ALUNO", "DATA", "HORA", "TITULO"]]; 
  let dadosPainel = [["DATA", "HORA", "EVENTO"]];      

  eventos.forEach(e => {
    let tituloOriginal = e.getTitle();
    let tBusca = tituloOriginal.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
    
    if (tBusca.includes(nome)) {
      let inicio = e.getStartTime();
      let dataF = Utilities.formatDate(inicio, "GMT-3", "dd/MM/yy");
      let horaF = Utilities.formatDate(inicio, "GMT-3", "HH:mm");
      
      dadosAba.push([nome.toUpperCase(), inicio, inicio.getHours(), tituloOriginal]);
      dadosPainel.push([dataF, horaF, tituloOriginal]);
    }
  });

  // 5. OUTPUT NA ABA ESPECÍFICA
  sheetFuturo.clear();
  if (dadosAba.length > 1) {
    sheetFuturo.getRange(1, 1, dadosAba.length, 4).setValues(dadosAba);
  }

  // 6. OUTPUT NO PAINEL (Estilo Site)
  if (dadosPainel.length > 1) {
    shPainel.getRange("D4:H4").merge()
            .setValue("📅 PRÓXIMAS AULAS: " + nome.toUpperCase())
            .setFontWeight("bold")
            .setBackground("#e2efda") 
            .setFontColor("#2e7d32");

    const rangeTabela = shPainel.getRange(6, 4, dadosPainel.length, 3);
    rangeTabela.setValues(dadosPainel);
    
    shPainel.getRange(6, 4, 1, 3).setBackground("#d9d9d9").setFontWeight("bold");
    shPainel.setRowHeights(6, dadosPainel.length, 28);
    
    // CORREÇÃO AQUI: de .setFontItalic para .setFontStyle("italic")
    shPainel.getRange(6 + dadosPainel.length + 1, 4)
            .setValue("Lista atualizada com sucesso.")
            .setFontSize(10)
            .setFontStyle("italic");
            
  } else {
    shPainel.getRange("D4").setValue("❌ Nenhum evento encontrado para '" + nome.toUpperCase() + "' nos próximos 30 dias.")
            .setFontColor("red").setFontWeight("bold");
  }
}

function verificarRenovacoesZero() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][27] === 0 && data[i][11] === "ativo") sheet.getRange(i+1, 28).setBackground("#ea9999");
  }
}

/**
 * OUTPUT ESTILO "SITE": GERA RESUMO E LISTA DIRETAMENTE NO PAINEL
 */
/**
 * OUTPUT ESTILO "SITE": FONTE MONTSERRAT TAMANHO 14
 */
/**
 * VERSÃO CORRIGIDA: Limpa mesclagens antes de escrever para evitar o erro do Sheets
 */
function verHistoricoAulasFeitas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shPainel = ss.getSheetByName("PAINEL");
  const shContratos = ss.getSheetByName("Contratos");
  const ui = SpreadsheetApp.getUi();

  const resp = ui.prompt('Histórico do Aluno', 'Digite o nome do aluno:', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  
  const busca = resp.getResponseText().trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  if (!busca) return;

  // --- SOLUÇÃO PARA O ERRO DE MESCLAGEM ---
  const areaMonitor = shPainel.getRange("D4:H40");
  areaMonitor.breakApart(); // <--- Comando mágico que desfaz mesclagens problemáticas
  areaMonitor.clearContent()
             .setBackground(null)
             .setFontWeight("normal")
             .setFontFamily("Montserrat")
             .setFontSize(14)
             .setVerticalAlignment("middle")
             .setHorizontalAlignment("left");

  const dataC = shContratos.getDataRange().getValues();
  let dadosAluno = null;
  for (let i = 1; i < dataC.length; i++) {
    let nomeC = String(dataC[i][2] || "").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
    if (nomeC.includes(busca)) {
      dadosAluno = dataC[i];
      break;
    }
  }

  if (!dadosAluno) {
    shPainel.getRange("D4").setValue("❌ Aluno não encontrado.").setFontColor("red").setFontWeight("bold");
    return;
  }

  // Escrita segura com merge controlado pelo código
  shPainel.getRange("D4:H4").breakApart().merge().setValue("📊 HISTÓRICO: " + dadosAluno[2].toUpperCase())
          .setFontWeight("bold")
          .setBackground("#f3f3f3");
  
  const cabecalhoResumo = [["PLANO", "FEITAS", "RESTANTES", "STATUS"]];
  const valoresResumo = [[dadosAluno[4], dadosAluno[29] || 0, dadosAluno[27] || 0, String(dadosAluno[11]).toUpperCase()]];
  
  shPainel.getRange("D5:G5").setValues(cabecalhoResumo).setBackground("#eeeeee").setFontWeight("bold");
  shPainel.getRange("D6:G6").setValues(valoresResumo);

  // Busca na Agenda
  const agenda = CalendarApp.getDefaultCalendar();
  const eventos = agenda.getEvents(new Date(2025, 0, 1), new Date()); 
  let listaEventos = [["DATA", "HORA", "EVENTO"]];

  eventos.forEach(e => {
    let t = e.getTitle().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
    if (t.includes(busca) && (/✅|✔|☑|✓/.test(e.getTitle()))) {
      listaEventos.push([
        Utilities.formatDate(e.getStartTime(), "GMT-3", "dd/MM/yy"),
        Utilities.formatDate(e.getStartTime(), "GMT-3", "HH:mm"),
        e.getTitle()
      ]);
    }
  });

  if (listaEventos.length > 1) {
    shPainel.getRange(8, 4).setValue("📅 AULAS REALIZADAS:").setFontWeight("bold");
    const rangeLista = shPainel.getRange(9, 4, listaEventos.length, 3);
    rangeLista.setValues(listaEventos);
    shPainel.getRange(9, 4, 1, 3).setBackground("#cfe2f3").setFontWeight("bold");
  } else {
    shPainel.getRange("D8").setValue("Nenhuma aula encontrada.");
  }
}
/**
 * Função Auxiliar para manter a aba de histórico atualizada também
 */
function atualizarAbaHistoricoSilencioso(dados, lista) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName("Historico_Aluno_Pilates") || ss.insertSheet("Historico_Aluno_Pilates");
  sh.clear();
  sh.getRange(1, 1).setValue("HISTÓRICO COMPLETO: " + dados[2]);
  sh.getRange(3, 1, lista.length, 3).setValues(lista);
}
function CONTRATOS_adicionarDaAgenda() { SpreadsheetApp.getUi().alert("Função de adicionar alunos novos."); }

/**
 * VIRADA DE MÊS AUTOMÁTICA (Reseta o mês e gera as cobranças)
 * Índice 38 é usado para ler a coluna AM (Telefone)
 */
function viradaDeMesAutomatica() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Contratos");
  const data = sheet.getDataRange().getValues();
  
  // MAPEAMENTO DE COLUNAS (v6.1 - Blindada)
  const COL_INICIO      = 1;  // B
  const COL_NOME        = 3;  // D
  const COL_PLANO       = 4;  // E
  const COL_FREQ        = 5;  // F (Frequência semanal)
  const COL_VALOR_BASE  = 9;  // J (Preço unitário)
  const COL_K           = 10; // K (Contratadas)
  const COL_STATUS      = 11; // L (Status)
  const COL_RESTANTES   = 27; // AB (Saldo Mês)
  const COL_FEITAS_AD   = 29; // AD (Feitas)
  const COL_RECUPERACAO = 32; // AG (Saldo Acumulado)
  const COL_TELEFONE_AM = 38; // AM (Índice 38 no array data)

  const hoje = new Date();
  const primeiroDiaMes = new Date(hoje.getFullYear(), hoje.getMonth(), 1);
  const pixChave = "20028886000199";

  console.log("--- INICIANDO VIRADA DE MÊS ---");

  for (let i = 1; i < data.length; i++) {
    const nomeCompleto = String(data[i][COL_NOME]).trim();
    const plano = String(data[i][COL_PLANO] || "").toLowerCase();
    const status = String(data[i][COL_STATUS] || "").toLowerCase();
    const telefoneRaw = String(data[i][COL_TELEFONE_AM] || "");
    const telefoneLimpo = telefoneRaw.replace(/\D/g, "");

    if (plano.includes("mensal") && status === "ativo" && nomeCompleto !== "") {
      
      try {
        if (telefoneLimpo.length >= 10) {
          const primeiroNome = nomeCompleto.split(" ")[0];
          const freq = Number(data[i][COL_FREQ]) || 0;
          const valorBase = Number(data[i][COL_VALOR_BASE]) || 0;
          
          // Cálculo: (Valor J * Frequência F) * 4 semanas
          const totalAcertar = (valorBase * freq) * 4;
          const valorFormatado = totalAcertar.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });

          const textoMensagem = `Oi ${primeiroNome}, tudo bem? Você acabou de completar mais um mês no nosso espaço, estamos muito felizes de te ter conosco! 

Esta é uma mensagem automática para lhe avisar da renovação da sua mensalidade (${freq}x na semana). 

Segue as informações da sua renovação para você continuar cuidando de você e mantendo seus resultados lindos:

📍 Código PIX (CNPJ): ${pixChave}
💰 Valor: R$ ${valorFormatado}

Posso emitir o próximo período? 😊`;

          const linkFinal = `https://wa.me/55${telefoneLimpo}?text=${encodeURIComponent(textoMensagem)}`;
          
          // CRIAÇÃO DA TASK NO GOOGLE TASKS
          Tasks.Tasks.insert({
            title: `🚨 RENOVAR MENSALIDADE: ${nomeCompleto.toUpperCase()}`,
            notes: `Valor Total: R$ ${valorFormatado}\nLink para enviar:\n${linkFinal}`
          }, "@default");
        }
      } catch (e) {
        console.error("Erro na task de: " + nomeCompleto);
      }

      // RESET DA PLANILHA PARA O NOVO MÊS
      let saldoMesAnterior = Number(data[i][COL_RESTANTES]) || 0;
      let recAtual = Number(data[i][COL_RECUPERACAO]) || 0;
      
      // 1. Salva o que sobrou no acumulado (AG)
      sheet.getRange(i + 1, COL_RECUPERACAO + 1).setValue(recAtual + saldoMesAnterior);
      
      // 2. Atualiza data de início (B)
      sheet.getRange(i + 1, COL_INICIO + 1).setValue(primeiroDiaMes); 
      
      // 3. Zera as aulas feitas (AD)
      sheet.getRange(i + 1, COL_FEITAS_AD + 1).setValue(0);          
      
      // 4. Reinicia o saldo (AB) com o valor total de contratadas (K)
      let contratadasK = Number(data[i][COL_K]) || 0;
      sheet.getRange(i + 1, COL_RESTANTES + 1).setValue(contratadasK).setBackground("#cfe2f3");
    }
  }

  // CHAMA AUTOMATICAMENTE A CONFIGURAÇÃO DE DIAS E QUANTIDADE (K)
  configurarMensalistasPelaAgenda(); 
}

/**
 * CONFIGURAR MENSALISTAS PELA AGENDA
 * Escreve os dias da semana na Coluna AL (38) e calcula o total de aulas do mês (K).
 */
function configurarMensalistasPelaAgenda() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Contratos");
  const data = sheet.getDataRange().getValues();
  const agenda = CalendarApp.getDefaultCalendar();
  const hoje = new Date();
  
  for (let i = 1; i < data.length; i++) {
    const nomeOriginal = String(data[i][2]).trim(); // Nome na Coluna C
    const plano = String(data[i][4] || "").toLowerCase(); // Plano na Coluna E

    if (plano.includes("mensal") && nomeOriginal !== "") {
      const nomeBusca = nomeOriginal.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
      
      // Procura horários nos próximos 14 dias para identificar o padrão (ex: seg e quarta)
      const eventos = agenda.getEvents(hoje, new Date(hoje.getTime() + 14 * 24 * 60 * 60 * 1000));
      let diasIdentificados = [];
      
      eventos.forEach(e => {
        let t = e.getTitle().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
        if (t.includes(nomeBusca)) {
          let dia = e.getStartTime().getDay(); // 0 = Domingo, 1 = Segunda...
          if (!diasIdentificados.includes(dia)) diasIdentificados.push(dia);
        }
      });

      if (diasIdentificados.length > 0) {
        diasIdentificados.sort();
        
        // 1. ESCREVE OS DIAS NA COLUNA AL (38)
        sheet.getRange(i + 1, 38).setValue(diasIdentificados.join(",")); 
        
        // 2. CALCULA QUANTAS VEZES ESSES DIAS OCORREM NO MÊS ATUAL
        let cont = 0;
        let ano = hoje.getFullYear();
        let mes = hoje.getMonth();
        let fimMes = new Date(ano, mes + 1, 0).getDate();
        
        for (let d = 1; d <= fimMes; d++) {
          let dataTeste = new Date(ano, mes, d);
          if (diasIdentificados.includes(dataTeste.getDay())) {
            cont++;
          }
        }
        
        // 3. ATUALIZA A COLUNA K (11) COM O TOTAL DE AULAS DO MÊS
        sheet.getRange(i + 1, 11).setValue(cont);
      }
    }
  }
}

/**
 * VERSÃO FINAL: FISCAL DE ALUNOS (Limpeza de texto riscado e emojis)
 */
function GS_verificarNovosAlunosComPresenca() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Contratos");
  const data = sheet.getDataRange().getValues();
  
  // Padrão que aceita vários tipos de check
  const checkPat = /[\u2705\u2714\u2611\u2713]|✅|✔|☑|✓/; 
  
  // 1. Mapeia nomes da planilha (Coluna D e C) com a nova limpeza
  const alunosExistentes = new Set();
  for (let i = 1; i < data.length; i++) {
    let nomeD = _gs_limparNomeComparacao(String(data[i][3] || ""));
    let nomeC = _gs_limparNomeComparacao(String(data[i][2] || ""));
    if (nomeD) alunosExistentes.add(nomeD);
    if (nomeC) alunosExistentes.add(nomeC);
  }

  const agendas = CalendarApp.getAllCalendars();
  const hoje = new Date();
  hoje.setHours(0,0,0,0); // Garante que pega desde o início do dia
  const proximoMes = new Date();
  proximoMes.setDate(hoje.getDate() + 30);
  
  let novosAchei = 0;
  let nomesJaProcessadosNestaRodada = new Set();

  agendas.forEach(agenda => {
    try {
      const eventos = agenda.getEvents(hoje, proximoMes);
      
      eventos.forEach(ev => {
        const tituloOriginal = ev.getTitle() || "";
        
        // Critério: Tem que ter o CHECK
        if (checkPat.test(tituloOriginal)) {
          let nomeAgendaLimpo = _gs_limparNomeComparacao(tituloOriginal);
          
          // Se não está na planilha e não foi processado agora
          if (nomeAgendaLimpo && !alunosExistentes.has(nomeAgendaLimpo) && !nomesJaProcessadosNestaRodada.has(nomeAgendaLimpo)) {
            
            // Limpa o nome para ficar bonito na Task (tira o riscado e o check)
            const nomeExibicao = _gs_limparNomeParaVisualizar(tituloOriginal);
            
            try {
              Tasks.Tasks.insert({
                title: "📝 NOVO ALUNO: " + nomeExibicao,
                notes: "Identificado com ✅ na agenda '" + agenda.getName() + "', mas não consta na aba Contratos.\n\nVerifique o cadastro."
              }, "@default");
              
              novosAchei++;
              nomesJaProcessadosNestaRodada.add(nomeAgendaLimpo);
            } catch (e) {
              console.error("Erro Tasks: " + e.message);
            }
          }
        }
      });
    } catch (e) { }
  });

  let msg = novosAchei > 0 
    ? `🔍 FISCAL: Encontrei ${novosAchei} aluno(s) pendente(s). Verifique seu Google Tasks!`
    : "🔍 FISCAL: Tudo em dia. Todos com ✅ estão na planilha.";
    
  gs_escreverNoMonitor(msg);
}

/**
 * FUNÇÃO DE LIMPEZA MÁXIMA: Remove acentos, emojis, espaços E TEXTO RISCADO
 */
function _gs_limparNomeComparacao(texto) {
  if (!texto) return "";
  return texto
    .toString()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "") // Remove acentos
    .replace(/[\u0335\u0336\u0337\u0338]/g, "") // REMOVE O RISCADO (Overlay characters)
    .replace(/[^\x00-\x7F]/g, "")    // Remove emojis e símbolos especiais
    .replace(/✅|✔|☑|✓/g, "")       // Remove checks remanescentes
    .replace(/\s+/g, "")            // Remove espaços
    .toLowerCase()
    .trim();
}

/**
 * AUXILIAR: Deixa o nome legível para a Task
 */
function _gs_limparNomeParaVisualizar(texto) {
  const checkPat = /[\u2705\u2714\u2611\u2713]|✅|✔|☑|✓/g;
  return texto
    .replace(checkPat, "")
    .replace(/[\u0335\u0336\u0337\u0338]/g, "") // Tira o riscado para ler na Task
    .trim()
    .toUpperCase();
}

/**
 * LIMPEZA TOTAL PARA COMPARAÇÃO
 */
function _gs_limparNomeComparacao(texto) {
  if (!texto) return "";
  return texto
    .toString()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "") // Remove acentos
    .replace(/[^\x00-\x7F]/g, "")    // Remove emojis
    .replace(/✅|✔|☑|✓/g, "")       // Remove checks
    .replace(/\s+/g, "")            // Remove TODOS os espaços
    .toLowerCase()
    .trim();
}

// 1. Botão: ATUALIZAR PLANOS (Ignora mensalidade)
function btn_atualizarPlanosLongos() {
  gs_escreverNoMonitor("⏳ Atualizando saldos dos PLANOS...\n(Aguarde, processando apenas planos longos)");
  try {
    atualizarPlanos_BRUTO(); // Sua função que já filtra !plano.includes("mensalidade")
    gs_escreverNoMonitor("✅ PLANOS ATUALIZADOS!\n\nOs saldos dos pacotes de aulas foram sincronizados com a agenda.");
  } catch(e) { 
    gs_escreverNoMonitor("❌ ERRO:\n" + e.message); 
  }
}

// 2. Botão: RELATÓRIO + LISTA DE EVENTOS
function btn_relatorio() {
  const ui = SpreadsheetApp.getUi();
  const prompt = ui.prompt('Gerar Relatório', 'Nome do aluno:', ui.ButtonSet.OK_CANCEL);
  if (prompt.getSelectedButton() !== ui.Button.OK) return;
  
  const busca = prompt.getResponseText().trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Contratos");
  const data = sh.getDataRange().getValues();
  
  let linha = -1;
  for (let i = 1; i < data.length; i++) {
    let nome = String(data[i][2]).toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
    if (nome.includes(busca)) { linha = i; break; }
  }

  if (linha === -1) {
    gs_escreverNoMonitor("❌ Aluno não encontrado.");
  } else {
    const d = data[linha];
    let relatorio = "📊 RELATÓRIO: " + d[2].toUpperCase() + "\n" +
                    "----------------------------------\n" +
                    "📋 Plano: " + d[4] + "\n" +
                    "✅ Feitas: " + (d[29] || 0) + " | ⏳ Restantes: " + (d[27] || 0) + "\n" +
                    "Status: " + String(d[11]).toUpperCase() + "\n" +
                    "----------------------------------\n" +
                    "📅 PRÓXIMAS AULAS:\n";

    // Busca eventos futuros (30 dias)
    const agenda = CalendarApp.getDefaultCalendar();
    const eventos = agenda.getEvents(new Date(), new Date(Date.now() + 30 * 24 * 60 * 60 * 1000));
    let encontrou = false;
    
    eventos.forEach(ev => {
      if (ev.getTitle().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").includes(busca)) {
        let dataF = Utilities.formatDate(ev.getStartTime(), "GMT-3", "dd/MM (EEE) - HH:mm");
        relatorio += "🔹 " + dataF + "\n";
        encontrou = true;
      }
    });
    if (!encontrou) relatorio += "Nenhum horário futuro na agenda.";
    
    gs_escreverNoMonitor(relatorio);
  }
}

// 3. Botão: EVENTOS FUTUROS (Apenas Lista)
function btn_eventosFuturos() {
  const ui = SpreadsheetApp.getUi();
  const prompt = ui.prompt('Consultar Agenda', 'Nome do aluno:', ui.ButtonSet.OK_CANCEL);
  if (prompt.getSelectedButton() !== ui.Button.OK) return;
  
  const busca = prompt.getResponseText().trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  gs_escreverNoMonitor("🔍 Consultando agenda...");

  const agenda = CalendarApp.getDefaultCalendar();
  const eventos = agenda.getEvents(new Date(), new Date(Date.now() + 60 * 24 * 60 * 60 * 1000));
  let lista = "📅 AGENDA FUTURA (60 dias):\n----------------------------------\n";
  let cont = 0;

  eventos.forEach(ev => {
    if (ev.getTitle().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").includes(busca)) {
      let dataF = Utilities.formatDate(ev.getStartTime(), "GMT-3", "dd/MM/yy - HH:mm");
      lista += "📌 " + dataF + " (" + ev.getTitle() + ")\n";
      cont++;
    }
  });
  
  lista += cont === 0 ? "❌ Nada encontrado." : "\nTotal: " + cont + " agendamentos.";
  gs_escreverNoMonitor(lista);
}

// Função para escrever no monitor (Célula D4)
function gs_escreverNoMonitor(texto) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PAINEL");
  sh.getRange("D4:H15").clearContent();
  sh.getRange("D4").setValue(texto);
}
/**
 * FUNÇÃO PARA CRIAR TASK COM LINK MÁGICO
 * Mapeamento: D (Nome), E (Serviço), AL (Telefone)
 */
function _gs_criarTaskComLink(nome, servico, saldo, telefoneBruto) {
  let linkWhats = "";
  
  // Limpa o telefone: remove espaços e símbolos (Ex: 55 51... -> 5551...)
  let telefone = String(telefoneBruto || "").replace(/\D/g, "");

  if (telefone && telefone.length >= 10) {
    // A MENSAGEM:
    let msg = `Olá ${nome}! Vi aqui no Studio GS que seu plano de ${servico} está no fim (resta apenas ${saldo} aula). Vamos garantir seu próximo horário? 😊`;
    let msgEncoded = encodeURIComponent(msg);
    linkWhats = `https://wa.me/${telefone}?text=${msgEncoded}`;
  }

  const tituloTask = `🚨 RENOVAR: ${nome}`;
  const notas = (linkWhats !== "") 
    ? `Serviço: ${servico}\nSaldo: ${saldo} aulas.\n\nClique para abrir WhatsApp:\n${linkWhats}`
    : `Serviço: ${servico}\nSaldo: ${saldo} aulas.\n(Telefone não cadastrado na Coluna AL)`;

  try {
    Tasks.Tasks.insert({
      title: tituloTask,
      notes: notas
    }, "@default");
  } catch (e) {
    console.error("Erro ao criar Task: " + e.message);
  }
}
/**
 * 1. FUNÇÃO PRINCIPAL: VERIFICA QUEM ZEROU
 */
function verificarRenovacoesZero() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Contratos");
  const data = sheet.getDataRange().getValues();
  
  // MAPEAMENTO
  const COL_ALUNO = 3;      // D
  const COL_SERVICO = 4;    // E
  const COL_STATUS = 11;    // L
  const COL_RESTANTES = 27; // AB
  const COL_TELEFONE = 38;  // AL

  let listaAlertas = "🚨 ALERTAS DE RENOVAÇÃO:\n--------------------------\n";
  let encontrou = false;

  for (let i = 1; i < data.length; i++) {
    const nome = data[i][COL_ALUNO];
    const servico = data[i][COL_SERVICO];
    const saldo = data[i][COL_RESTANTES];
    const status = String(data[i][COL_STATUS]).toLowerCase();
    const telefoneBruto = data[i][COL_TELEFONE];

    if (nome && status === "ativo" && (saldo === 0)) {
      encontrou = true;
      listaAlertas += `🔹 ${nome.toUpperCase()} (${saldo} aulas)\n`;
      
      // 🔥 AQUI ESTÁ A CHAVE: CHAMA A FUNÇÃO ABAIXO PASSANDO TUDO
      _gs_criarTaskInteligente(nome, servico, saldo, telefoneBruto);
      
      sheet.getRange(i + 1, COL_RESTANTES + 1).setBackground("#ea9999");
    }
  }

  if (!encontrou) listaAlertas = "✅ TUDO EM DIA!\nNenhum saldo crítico (0 ou 1).";
  gs_escreverNoMonitor(listaAlertas);
}



function _gs_criarTaskInteligente(nome, servico, saldo, telefoneBruto) {
  let linkWhats = "";
  
  // Limpa o telefone (Coluna AL)
  let telefone = String(telefoneBruto || "").replace(/\D/g, "");

  // Se tiver telefone, cria o link mágico
  if (telefone && telefone.length >= 10) {
    let msg = `Olá ${nome}! Vi aqui no Espaço Giulia Schifino que seu plano de ${servico} está no fim (resta apenas ${saldo} aula). Vamos garantir seu próximo horário?`;
    let msgEncoded = encodeURIComponent(msg);
    linkWhats = `https://wa.me/${telefone}?text=${msgEncoded}`;
  }

  const tituloTask = `🚨 RENOVAR: ${nome.toUpperCase()}`;
  const notas = (linkWhats !== "") 
    ? `Serviço: ${servico}\nSaldo: ${saldo} aulas.\n\n📲 CLIQUE PARA AVISAR NO WHATS:\n${linkWhats}`
    : `Serviço: ${servico}\nSaldo: ${saldo} aulas.\n(Telefone não cadastrado na Coluna AL)`;

  try {
    // ESSA LINHA É A QUE FINALMENTE CRIA NO GOOGLE TASKS
    Tasks.Tasks.insert({
      title: tituloTask,
      notes: notas
    }, "@default");
  } catch (e) {
    console.error("Erro ao criar task: " + e.message);
  }
}

/**
 * ESCREVER NO MONITOR - ESTILO PAINEL
 */
function gs_escreverNoMonitor(texto) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("PAINEL");
  if (!sh) return;
  
  const areaMonitor = sh.getRange("D4:H15");
  areaMonitor.breakApart().clearContent().setBackground(null);
  areaMonitor.merge();
  
  areaMonitor.setFontFamily("Montserrat")
             .setFontSize(14)
             .setVerticalAlignment("middle")
             .setHorizontalAlignment("center")
             .setWrap(true)
             .setValue(texto);
}


/**
 * ASSISTENTE DE CADASTRO PASSO A PASSO (V2)
 * Considera apenas alunos que NÃO possuem contrato com status "ativo"
 */
function CONTRATOS_adicionarDaAgenda() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Contratos");
  const ui = SpreadsheetApp.getUi();
  
  // 1. MAPEIA APENAS QUEM JÁ ESTÁ ATIVO
  const data = sh.getDataRange().getValues();
  const alunosAtivos = new Set();
  
  for (let i = 1; i < data.length; i++) {
    let nome = _gs_limparNomeComparacao(String(data[i][3] || "")); // Coluna D
    let status = String(data[i][11] || "").toLowerCase().trim(); // Coluna L
    
    if (nome && status === "ativo") {
      alunosAtivos.add(nome);
    }
  }

  // 2. BUSCA NA AGENDA QUEM NÃO ESTÁ ATIVO
  const agenda = CalendarApp.getDefaultCalendar();
  const hoje = new Date();
  const eventos = agenda.getEventsForDay(hoje);
  let nomesParaAdicionar = [];
  let infoEventos = {};

  eventos.forEach(ev => {
    const titulo = ev.getTitle();
    if (/✅|✔|☑|✓/.test(titulo)) {
      let nomeLimpo = _gs_limparNomeComparacao(titulo);
      
      // Se o aluno NÃO está na lista de ATIVOS, ele é candidato ao cadastro
      if (!alunosAtivos.has(nomeLimpo)) {
        let nomeBonito = _gs_limparNomeParaVisualizar(titulo);
        // Evita duplicar o mesmo nome se ele aparecer 2x na agenda do dia
        if (!nomesParaAdicionar.includes(nomeBonito)) {
          nomesParaAdicionar.push(nomeBonito);
          infoEventos[nomeBonito] = {
            cor: ev.getColor(),
            data: ev.getStartTime()
          };
        }
      }
    }
  });

  if (nomesParaAdicionar.length === 0) {
    ui.alert("⚠️ Studio GS: Nenhum aluno novo ou com contrato expirado encontrado com ✅ hoje.");
    return;
  }

  // 3. PASSO A PASSO: SELEÇÃO DO ALUNO
  const listaNomes = nomesParaAdicionar.join("\n- ");
  const respAluno = ui.prompt("PASSO 1: Selecionar Aluno", 
    "Alunos sem contrato ativo hoje. Digite o nome de quem deseja cadastrar:\n\n- " + listaNomes, ui.ButtonSet.OK_CANCEL);
  
  if (respAluno.getSelectedButton() !== ui.Button.OK) return;
  const aluno = respAluno.getResponseText().trim().toUpperCase();

  if (!nomesParaAdicionar.includes(aluno)) {
    ui.alert("❌ Erro: O nome digitado não está na lista ou já possui um plano ativo.");
    return;
  }

  // 4. PASSO A PASSO: PLANO
  const respPlano = ui.prompt("PASSO 2: Qual o Plano?", 
    "Digite o plano para " + aluno + ":\n(Avulsa, 5 aulas, 12 aulas, 6 meses, 12 meses)", ui.ButtonSet.OK_CANCEL);
  if (respPlano.getSelectedButton() !== ui.Button.OK) return;
  const plano = respPlano.getResponseText().toLowerCase();

  // 5. PASSO A PASSO: FREQUÊNCIA
  const respFreq = ui.prompt("PASSO 3: Frequência Semanal", 
    "Vezes por semana? (1, 2 ou 3):", ui.ButtonSet.OK_CANCEL);
  if (respFreq.getSelectedButton() !== ui.Button.OK) return;
  const freq = parseInt(respFreq.getResponseText()) || 1;

  // 6. CÁLCULO E MODALIDADE
  let aulasContratadas = 0;
  if (plano.includes("5")) aulasContratadas = 5;
  else if (plano.includes("12 aulas")) aulasContratadas = 12;
  else if (plano.includes("6 meses")) aulasContratadas = freq * 4 * 6;
  else if (plano.includes("12 meses")) aulasContratadas = freq * 4 * 12;
  else if (plano.includes("avulsa")) aulasContratadas = 1;

  const modalidade = (infoEventos[aluno].cor === "1") ? "DUPLA" : "INDIVIDUAL";

  // 7. CONFIRMAÇÃO
  const confirmacao = ui.alert("CONFIRMAÇÃO", 
    "Aluno: " + aluno + "\nPlano: " + plano.toUpperCase() + "\nFreq: " + freq + "x\nModalidade: " + modalidade + "\nAulas: " + aulasContratadas + "\n\nSalvar novo contrato?", ui.ButtonSet.YES_NO);

  if (confirmacao !== ui.Button.YES) return;

  // 8. SALVANDO
  const novaLinha = sh.getLastRow() + 1;
  sh.getRange(novaLinha, 2).setValue(infoEventos[aluno].data); 
  sh.getRange(novaLinha, 4).setValue(aluno);                  
  sh.getRange(novaLinha, 5).setValue(plano.toUpperCase());
  sh.getRange(novaLinha, 6).setValue(freq);                   
  sh.getRange(novaLinha, 8).setValue(modalidade);             
  sh.getRange(novaLinha, 11).setValue(aulasContratadas);      
  sh.getRange(novaLinha, 12).setValue("ativo");               
  sh.getRange(novaLinha, 30).setValue(1);                     
  sh.getRange(novaLinha, 28).setValue(aulasContratadas - 1);  

  ui.alert("✅ Sucesso! Plano ATIVO registrado.");
}
// --- FUNÇÃO PARA ABRIR O APP ---
function doGet() {
  return HtmlService.createTemplateFromFile('Interface').evaluate()
      .setTitle('Gestão Studio GS')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, user-scalable=no, viewport-fit=cover')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// --- PONTES DE EXECUÇÃO ---

function ponte_atualizarPlanos() {
  try {
    atualizarPlanos_BRUTO(); // Sua função original
    return "✅ Planos atualizados!";
  } catch(e) { return "❌ Erro: " + e.message; }
}

function ponte_historico(nome) {
  try {
    // Aqui injetamos o nome direto na sua lógica de busca para não abrir prompt
    _gs_gerarHistoricoSemPrompt(nome); 
    return "📊 Histórico de " + nome + " gerado no Painel!";
  } catch(e) { return "❌ Erro: " + e.message; }
}

function ponte_renovacaoApp(dados) {
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Contratos");
    const data = sh.getDataRange().getValues();
    let linha = -1;
    let busca = dados.aluno.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][3]).toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").includes(busca)) {
        linha = i + 1;
        break;
      }
    }

    if (linha === -1) return "❌ Aluno não encontrado!";

    // Executa a renovação nas colunas exatas da v6.0
    sh.getRange(linha, 2).setValue(new Date()); // B
    sh.getRange(linha, 5).setValue(dados.plano.toUpperCase()); // E
    sh.getRange(linha, 6).setValue(dados.freq); // F
    sh.getRange(linha, 12).setValue("ativo"); // L
    sh.getRange(linha, 30).setValue(0); // AD
    
    return "✅ " + dados.aluno + " renovado!";
  } catch(e) { return "❌ Erro: " + e.message; }
}
/**
 * PONTE DE BUSCA PARA O APP (PILATES)
 */
function GS_MOTOR_BUSCA_TUDO(nomeBusca) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName("Contratos");
    const data = sh.getDataRange().getValues();
    const buscaLimpa = _gs_limparNomeComparacao(nomeBusca);
    
    if (!buscaLimpa) return {erro: "Digite um nome válido."};

    let alunoData = null;
    for (let i = 1; i < data.length; i++) {
      let nomeC = _gs_limparNomeComparacao(data[i][2]); // Coluna C
      let nomeD = _gs_limparNomeComparacao(data[i][3]); // Coluna D
      if (nomeC.includes(buscaLimpa) || nomeD.includes(buscaLimpa)) {
        alunoData = {
          nome: data[i][2],
          plano: data[i][4],
          restantes: data[i][27] || 0,
          feitas: data[i][29] || 0,
          status: data[i][11]
        };
        break;
      }
    }

    if (!alunoData) return {erro: "Aluno não encontrado."};

    // Busca Histórico e Futuros na Agenda
    const agenda = CalendarApp.getDefaultCalendar();
    const hoje = new Date();
    const eventos = agenda.getEvents(new Date(2025, 0, 1), new Date(hoje.getTime() + 30 * 24 * 60 * 60 * 1000));
    
    let h = []; let f = [];
    eventos.forEach(e => {
      let tOriginal = e.getTitle();
      let tLimpo = _gs_limparNomeComparacao(tOriginal);
      
      if (tLimpo.includes(buscaLimpa)) {
        let txt = Utilities.formatDate(e.getStartTime(), "GMT-3", "dd/MM HH:mm") + " - " + tOriginal;
        if (/✅|✔|☑|✓/.test(tOriginal)) h.push(txt);
        else if (e.getStartTime() >= hoje) f.push(txt);
      }
    });

    return {
      nome: alunoData.nome,
      restantes: alunoData.restantes,
      feitas: alunoData.feitas,
      historico: h.sort().reverse(),
      futuros: f.sort()
    };
  } catch (e) {
    return {erro: "Erro no motor: " + e.message};
  }
}