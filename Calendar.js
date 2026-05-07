// Calendar.gs

function criarEventosCalendar() {
  const sh = getSheet_('Vistorias_Tratadas');

  if (!sh || sh.getLastRow() < 2) {
    Logger.log('Nenhuma vistoria para processar no Calendar');
    return { criados: [] };
  }

  const config = getConfig_();
  Logger.log('CONFIG LIDO: ' + JSON.stringify(config));

  const calendarId = config.calendarId || config.calendarid || config.calendar_id;

  if (!calendarId) {
    throw new Error('calendarId não configurado na aba Config.');
  }

  const calendar = CalendarApp.getCalendarById(calendarId);

  if (!calendar) {
    throw new Error('Calendário não encontrado: ' + calendarId);
  }

  const dados = sh.getDataRange().getValues();
  const cab = dados[0];

  const idx = {
    id: cab.indexOf('id_atividade'),
    titulo: cab.indexOf('titulo_vistoria'),
    status: cab.indexOf('status_conformidade'),
    criticidade: cab.indexOf('criticidade'),
    percentual: cab.indexOf('percentual_conformidade'),
    naoConf: cab.indexOf('qtde_nao_conformes'),
    criticos: cab.indexOf('qtde_campos_criticos_irregulares'),
    dataFim: cab.indexOf('data_finalizacao'),
    dataSync: cab.indexOf('data_sincronizacao'),
    revisao: cab.indexOf('proxima_revisao'),
    l15: cab.indexOf('lembrete_15_dias'),
    l28: cab.indexOf('lembrete_28_dias')
  };

  const colEventRev = getOrCreateColumn_(sh, cab, 'event_id_revisao');
  const colEvent15 = getOrCreateColumn_(sh, cab, 'event_id_lembrete_15');
  const colEvent28 = getOrCreateColumn_(sh, cab, 'event_id_lembrete_28');

  const criados = [];

  for (let i = 1; i < dados.length; i++) {
    const linha = dados[i];

    const id = linha[idx.id];
    const titulo = linha[idx.titulo];

    if (!id || !titulo) continue;

    const status = normalizarTexto_(linha[idx.status]);

    if (status === 'conforme') {
      const idEvento = criarEventoSeNaoExiste_(calendar, sh, i + 1, colEventRev + 1, linha[idx.revisao], '🔄 Revisão 120 dias', linha, idx);
      if (idEvento) criados.push(idEvento);
    } else {
      const id15 = criarEventoSeNaoExiste_(calendar, sh, i + 1, colEvent15 + 1, linha[idx.l15], '⚠️ Lembrete 15 dias', linha, idx);
      const id28 = criarEventoSeNaoExiste_(calendar, sh, i + 1, colEvent28 + 1, linha[idx.l28], '🚨 Lembrete 28 dias', linha, idx);

      if (id15) criados.push(id15);
      if (id28) criados.push(id28);
    }
  }

  Logger.log('Eventos criados: ' + criados.length);

  return { criados: criados };
}

function criarEventoSeNaoExiste_(calendar, sh, linhaPlanilha, colunaEventId, dataEvento, tipo, linha, idx) {
  if (!dataEvento) return '';

  const data = new Date(dataEvento);
  if (isNaN(data)) return '';

  const agora = new Date();
  if (data <= agora) return '';

  const eventIdExistente = sh.getRange(linhaPlanilha, colunaEventId).getValue();
  if (eventIdExistente) return '';

  const titulo = tipo + ' - ' + linha[idx.titulo];

  const inicio = new Date(data);
  const fim = new Date(data);
  fim.setHours(fim.getHours() + 1);

  const descricao = [
    'ID: ' + linha[idx.id],
    'Vistoria: ' + linha[idx.titulo],
    'Status: ' + linha[idx.status],
    'Data Sincronização: ' + formatarData_(linha[idx.dataSync]),
    'Data Finalização: ' + formatarData_(linha[idx.dataFim]),
    'Conformidade: ' + Math.round((Number(linha[idx.percentual]) || 0) * 100) + '%',
    'Não conformes: ' + (linha[idx.naoConf] || 0),
    'Críticos: ' + (linha[idx.criticos] || 0)
  ].join('\n');

  const evento = calendar.createEvent(titulo, inicio, fim, {
    description: descricao,
    location: 'Revisão de Vistoria - Produttivo'
  });

  aplicarCorEvento_(evento, linha[idx.criticidade]);
  configurarLembretes_(evento, linha[idx.criticidade]);

  sh.getRange(linhaPlanilha, colunaEventId).setValue(evento.getId());

  return evento.getId();
}

function limparEventosAntigos() {
  const sh = getSheet_('Vistorias_Tratadas');

  if (!sh || sh.getLastRow() < 2) {
    Logger.log('Nenhuma linha para limpar.');
    return 0;
  }

  const config = getConfig_();
  const calendarId = config.calendarId || config.calendarid || config.calendar_id;

  if (!calendarId) {
    throw new Error('calendarId não configurado na aba Config.');
  }

  const calendar = CalendarApp.getCalendarById(calendarId);

  if (!calendar) {
    throw new Error('Calendário não encontrado: ' + calendarId);
  }

  const dados = sh.getDataRange().getValues();
  const cab = dados[0];

  const pares = [
    { ev: cab.indexOf('event_id_revisao'), dt: cab.indexOf('proxima_revisao') },
    { ev: cab.indexOf('event_id_lembrete_15'), dt: cab.indexOf('lembrete_15_dias') },
    { ev: cab.indexOf('event_id_lembrete_28'), dt: cab.indexOf('lembrete_28_dias') }
  ];

  const agora = new Date();
  let removidos = 0;

  for (let i = 1; i < dados.length; i++) {
    pares.forEach(function(par) {
      if (par.ev < 0 || par.dt < 0) return;

      const eventId = dados[i][par.ev];
      const dataEvento = new Date(dados[i][par.dt]);

      if (!eventId || isNaN(dataEvento)) return;
      if (dataEvento >= agora) return;

      const evento = calendar.getEventById(eventId);

      if (evento) {
        evento.deleteEvent();
        removidos++;
      }

      sh.getRange(i + 1, par.ev + 1).clearContent();
    });
  }

  Logger.log('Eventos antigos removidos: ' + removidos);
  return removidos;
}

function sincronizarCalendarCompleto() {
  const removidos = limparEventosAntigos();
  const resultado = criarEventosCalendar();

  Logger.log(
    'Sincronização concluída. Removidos: ' + removidos +
    ', Criados: ' + resultado.criados.length
  );

  return {
    removidos: removidos,
    criados: resultado.criados
  };
}

function aplicarCorEvento_(evento, criticidade) {
  const crit = normalizarTexto_(criticidade);

  if (crit === 'alta') {
    evento.setColor(CalendarApp.EventColor.RED);
  } else if (crit === 'media' || crit === 'média') {
    evento.setColor(CalendarApp.EventColor.YELLOW);
  } else {
    evento.setColor(CalendarApp.EventColor.GREEN);
  }
}

function configurarLembretes_(evento, criticidade) {
  evento.removeAllReminders();

  const crit = normalizarTexto_(criticidade);

  if (crit === 'alta') {
    evento.addPopupReminder(1440);
    evento.addEmailReminder(1440);
  } else {
    evento.addPopupReminder(1440);
  }
}

function setupTriggerCalendar() {
  const triggers = ScriptApp.getProjectTriggers();

  const jaExiste = triggers.some(function(trigger) {
    return trigger.getHandlerFunction() === 'sincronizarCalendarCompleto';
  });

  if (jaExiste) {
    Logger.log('Trigger já existe para sincronizarCalendarCompleto');
    return;
  }

  ScriptApp.newTrigger('sincronizarCalendarCompleto')
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();

  Logger.log('Trigger criado com sucesso.');
}
