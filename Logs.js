

/*function testeEnvioDireto() {
  GmailApp.sendEmail(
    'lucas.leite@charth.com.br',
    'Teste direto Apps Script',
    'Se voce recebeu este email, o GmailApp esta funcionando.'
  );
}

/*
function testeEnvioEmailAvancado() {
  const ARQUIVO = 'TesteEmail.gs';
  const FUNCAO = 'testeEnvioEmailAvancado';

  const destinatario = Session.getActiveUser().getEmail();
  const cc = getEmailCcDefault_ ? getEmailCcDefault_() : '';

  const variaveis = {
    destinatario: destinatario,
    titulo_vistoria: 'Revisao Preventiva',
    local: 'Unidade Centro',
    proxima_revisao: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy'),
    dias_para_revisao: '7'
  };

  let assunto = 'Teste de lembrete: {{titulo_vistoria}}';
  let corpoTexto =
    'Ola {{destinatario}},\n\n' +
    'A vistoria "{{titulo_vistoria}}" no local "{{local}}" esta prevista para {{proxima_revisao}}.\n' +
    'Dias para revisao: {{dias_para_revisao}}.\n\n' +
    'Este e um teste automatico do Apps Script.';

  let corpoHtml =
    '<p>Ola <strong>{{destinatario}}</strong>,</p>' +
    '<p>A vistoria "<strong>{{titulo_vistoria}}</strong>" no local "<strong>{{local}}</strong>" esta prevista para <strong>{{proxima_revisao}}</strong>.</p>' +
    '<p>Dias para revisao: <strong>{{dias_para_revisao}}</strong>.</p>' +
    '<p><em>Este e um teste automatico do Apps Script.</em></p>';

  assunto = substituirVariaveisTeste_(assunto, variaveis);
  corpoTexto = substituirVariaveisTeste_(corpoTexto, variaveis);
  corpoHtml = substituirVariaveisTeste_(corpoHtml, variaveis);

  try {
    const opcoes = {
      htmlBody: corpoHtml
    };

    if (cc) {
      opcoes.cc = cc;
    }

    GmailApp.sendEmail(destinatario, assunto, corpoTexto, opcoes);

    if (typeof logInfo_ === 'function') {
      logInfo_(ARQUIVO, FUNCAO, 'Email de teste enviado com sucesso', {
        destinatario: destinatario,
        cc: cc,
        assunto: assunto
      });
    }

    Logger.log('Email de teste enviado para: ' + destinatario);

  } catch (e) {
    if (typeof logErro_ === 'function') {
      logErro_(ARQUIVO, FUNCAO, 'Erro ao enviar email de teste', {
        destinatario: destinatario,
        cc: cc,
        assunto: assunto,
        mensagemErro: e.message,
        stack: e.stack
      });
    }

    throw e;
  }
}

function substituirVariaveisTeste_(textoBase, variaveis) {
  if (!textoBase) return '';

  return textoBase
    .replace(/\{\{destinatario\}\}/gi, variaveis.destinatario || '')
    .replace(/\{\{titulo_vistoria\}\}/gi, variaveis.titulo_vistoria || '')
    .replace(/\{\{local\}\}/gi, variaveis.local || '')
    .replace(/\{\{proxima_revisao\}\}/gi, variaveis.proxima_revisao || '')
    .replace(/\{\{dias_para_revisao\}\}/gi, variaveis.dias_para_revisao || '');
}

/*
//validar o envio de e-mail
function testeTemplateEmail() {
  const destinatario = Session.getActiveUser().getEmail();

  const template = 'Olá {{destinatario}}, sua vistoria "{{titulo_vistoria}}" será em {{proxima_revisao}} no local {{local}}.';

  const variaveis = {
    destinatario: destinatario,
    titulo_vistoria: 'Revisão Preventiva',
    local: 'Unidade Centro',
    proxima_revisao: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy')
  };

  const corpo = substituirVariaveis(template, variaveis);

  GmailApp.sendEmail(destinatario, 'Teste Template', corpo);

  Logger.log('Email template enviado');
}
*/  
  
  /*function testeCabecalho() {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Emails_enviar');
  const ultimaColuna = aba.getLastColumn();
  const cabecalho = aba.getRange(1, 1, 1, ultimaColuna).getDisplayValues()[0];

  Logger.log(JSON.stringify(cabecalho));
}
*/

/*function testeCriarEventoManual() {
  const config = getConfig_();
  const calendarId = config.calendarId || 'primary';
  const calendar = CalendarApp.getCalendarById(calendarId);

  const inicio = new Date();
  inicio.setDate(inicio.getDate() + 1);
  inicio.setHours(10, 0, 0, 0);

  const fim = new Date(inicio);
  fim.setHours(12, 0, 0, 0);

  const evento = calendar.createEvent(
    'TESTE - Evento Calendar',
    inicio,
    fim,
    { description: 'Teste manual de criação de evento' }
  );

  Logger.log('Evento criado: ' + evento.getId());
}
*/

/*function testarCalendar() {
  const config = getConfig_();
  const calendarId = config.calendarId || 'primary';
  const calendar = CalendarApp.getCalendarById(calendarId);

  Logger.log('calendarId: ' + calendarId);
  Logger.log('calendar encontrado? ' + !!calendar);

  if (calendar) {
    Logger.log('nome do calendário: ' + calendar.getName());
  }
}
*/









/*function forcarLembreteTeste() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const vt = ss.getSheetByName('Vistorias_Tratadas');
  
  // Pega 1ª linha com proxima_revisao hoje ou atrasada (assume col T=20)
  const dados = vt.getRange(2, 1, vt.getLastRow()-1, 25).getValues();
  for (let i = 0; i < dados.length; i++) {
    const proximaRevisao = dados[i][19]; // Col T
    if (proximaRevisao && new Date(proximaRevisao) <= new Date('2026-04-15')) {
      vt.getRange(i+2, 19).setValue('Enviar hoje'); // Col S
      console.log('✅ Linha', i+2, 'marcada "Enviar hoje"');
      break;
    }
  }
}
*/


/*function testarCabecalhos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  //const abaAtividades = ss.getSheetByName('Base_Atividades');
  const abaAtividades = ss.getSheetByName('Base_Formularios');
  const dados = abaAtividades.getDataRange().getValues();
  const cabAt = dados[0];

  Logger.log(JSON.stringify(cabAt));
}
*/

/*
function debugProcessar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1️⃣ Verifica se as abas existem
  const abas = ['Base_Atividades', 'Base_Formularios', 'Vistorias_Tratadas'];
  abas.forEach(nome => {
    const aba = ss.getSheetByName(nome);
    Logger.log(`${nome}: ${aba ? '✅ EXISTE (' + aba.getLastRow() + ' linhas)' : '❌ NÃO EXISTE'}`);
  });

  // 2️⃣ Verifica cabeçalhos
  const abaAt = ss.getSheetByName('Base_Atividades');
  const abaFo = ss.getSheetByName('Base_Formularios');
  
  if (abaAt) {
    const cabAt = abaAt.getRange(1, 1, 1, abaAt.getLastColumn()).getValues()[0];
    Logger.log('🔍 Base_Atividades cabeçalhos:', cabAt.slice(0, 10).join(', '));
  }
  
  if (abaFo) {
    const cabFo = abaFo.getRange(1, 1, 1, abaFo.getLastColumn()).getValues()[0];
    Logger.log('🔍 Base_Formularios cabeçalhos:', cabFo.slice(0, 10).join(', '));
  }

  // 3️⃣ Primeiras 3 linhas de cada
  if (abaAt && abaAt.getLastRow() > 1) {
    const dadosAt = abaAt.getRange(2, 1, 3, 10).getValues();
    Logger.log('📊 Base_Atividades (linhas 2-4):', dadosAt);
  }
  
  if (abaFo && abaFo.getLastRow() > 1) {
    const dadosFo = abaFo.getRange(2, 1, 3, 10).getValues();
    Logger.log('📊 Base_Formularios (linhas 2-4):', dadosFo);
  }

  Logger.log('🔍 DEBUG FINALIZADO - Verifique os logs acima!');
} 

*/


/*
function garantirAbaLogs_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let aba = ss.getSheetByName('Logs');

  if (!aba) {
    aba = ss.insertSheet('Logs');
    aba.appendRow([
      'data_hora',
      'nivel',
      'arquivo',
      'funcao',
      'mensagem',
      'detalhes'
    ]);
  }

  return aba;
}

function logPlanilha_(nivel, arquivo, funcao, mensagem, detalhes) {
  const aba = garantirAbaLogs_();
  const agora = new Date();

  let detalhesTexto = '';
  if (detalhes !== null && detalhes !== undefined) {
    if (typeof detalhes === 'object') {
      try {
        detalhesTexto = JSON.stringify(detalhes);
      } catch (e) {
        detalhesTexto = String(detalhes);
      }
    } else {
      detalhesTexto = String(detalhes);
    }
  }

  aba.appendRow([
    agora,
    nivel,
    arquivo || '',
    funcao || '',
    mensagem || '',
    detalhesTexto
  ]);
}

function logInfo_(arquivo, funcao, mensagem, detalhes) {
  logPlanilha_('INFO', arquivo, funcao, mensagem, detalhes);
}

function logWarn_(arquivo, funcao, mensagem, detalhes) {
  logPlanilha_('WARN', arquivo, funcao, mensagem, detalhes);
}

function logErro_(arquivo, funcao, mensagem, detalhes) {
  logPlanilha_('ERROR', arquivo, funcao, mensagem, detalhes);
}
*/