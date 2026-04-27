//Calendar
function criarEventosCalendar() {
  const shVit = getSheet_('Vistorias_Tratadas');

  if (!shVit || shVit.getLastRow() < 2) {
    Logger.log('Nenhuma vistoria para processar no Calendar');
    return { criados: [], atualizados: [] };
  }
  log_('INFO', 'Iniciando criação de eventos no Calendar');
  const config = getConfig_();
  const calendarId = config.calendarId || 'primary';
  const diasAntecedencia = parseInt(config.dias_aviso, 10) || 7;
  log_('INFO', 'calendarId: ' + calendarId);
  log_('INFO', 'diasAntecedencia: ' + diasAntecedencia);

  const calendar = CalendarApp.getCalendarById(calendarId);
  if (!calendar) {
    throw new Error('Calendário não encontrado: ' + calendarId);
  }

  const dados = shVit.getDataRange().getValues();
  const cabecalho = dados[0];

  const idxIdAtv = getColIndex_(cabecalho, 'id_atividade');
  const idxTitulo = getColIndex_(cabecalho, 'titulo_vistoria');
  const idxDataSync = getColIndex_(cabecalho, 'data_sincronizacao');
  const idxDataFim = getColIndex_(cabecalho, 'data_finalizacao');
  const idxCriticidade = getColIndex_(cabecalho, 'criticidade');
  const idxProxRevisao = getColIndex_(cabecalho, 'proxima_revisao');
  const idxPontPerc = getColIndex_(cabecalho, 'percentual_conformidade');
  const idxHash = getColIndex_(cabecalho, 'hash_evento');
  const idxEventId = getOrCreateColumn_(shVit, cabecalho, 'event_id');

  const agora = new Date();
  const eventosCriados = [];
  const eventosAtualizados = [];
  const LIMITE_EVENTOS_POR_EXECUCAO = 50;
  let processados = 0;

  log_('INFO', 'Iniciando criação de eventos no Calendar');

  for (let i = 1; i < dados.length; i++) {
    if (processados >= LIMITE_EVENTOS_POR_EXECUCAO) {
      Logger.log('Limite de eventos por execução atingido.');
      break;
    }
    log_('INFO', 'Processando linha ' + i);

    const linha = dados[i];

    const idAtv = linha[idxIdAtv];
    const titulo = linha[idxTitulo] || 'Revisão de Vistoria';
    const dataSync = linha[idxDataSync];
    const dataFim = linha[idxDataFim];
    const criticidade = linha[idxCriticidade] || 'Baixa';
    const proxRevisao = linha[idxProxRevisao];
    const percentual = Number(linha[idxPontPerc]) || 0;
    const hashAtual = linha[idxHash];
    const eventIdSalvo = idxEventId >= 0 ? linha[idxEventId] : '';

    if (!idAtv || !proxRevisao){log_('WARN', 'Linha ignorada: sem id ou data'); 
    continue;
    }
    const dataRevisao = normalizarData_(proxRevisao);
    if (!dataRevisao) continue;
    if (dataRevisao <= agora){log_('INFO', 'Linha ignorada: data já passou');
    continue;
    }
    const diasAteRevisao = Math.ceil((dataRevisao - agora) / (1000 * 60 * 60 * 24));
    if (diasAteRevisao > diasAntecedencia) {
    log_('INFO', 'Linha ignorada: fora do prazo');
    continue;
    }
    const novoHash = gerarHashEvento_({
      idAtividade: idAtv,
      titulo: titulo,
      dataFim: dataFim,
      criticidade: criticidade,
      dataRevisao: dataRevisao,
      percentual: percentual
    });

    if (eventIdSalvo) {
      const eventoExistente = calendar.getEventById(eventIdSalvo);

      if (eventoExistente) {
        log_('INFO', 'Atualizando evento: ' + idAtv);
        if (hashAtual === novoHash) {
          continue;
        }

        atualizarEventoVistoria_(eventoExistente, {
          idAtividade: idAtv,
          titulo: titulo,
          dataSync: dataSync,
          dataFim: dataFim,
          criticidade: criticidade,
          dataRevisao: dataRevisao,
          percentual: percentual,
          hash: novoHash
        });

        shVit.getRange(i + 1, idxHash + 1).setValue(novoHash);
        eventosAtualizados.push(eventoExistente.getId());
        processados++;
        continue;
      }
    }
    log_('INFO', 'Criando evento: ' + idAtv);
      const evento = criarEventoVistoria_({
      idAtividade: idAtv,
      titulo: titulo,
      dataSync: dataSync,
      dataFim: dataFim,
      criticidade: criticidade,
      dataRevisao: dataRevisao,
      percentual: percentual,
      hash: novoHash
    }, calendar);

    if (evento) {
      log_('INFO', 'Evento criado: ' + evento.id);
      shVit.getRange(i + 1, idxEventId + 1).setValue(evento.id);
      shVit.getRange(i + 1, idxHash + 1).setValue(novoHash);
      eventosCriados.push(evento.id);
      processados++;
    }
  }

  Logger.log(
    'Calendar: ' +
    eventosCriados.length + ' criados, ' +
    eventosAtualizados.length + ' atualizados.'
  );
  log_(
  'INFO',
  'Finalizado: ' +
  eventosCriados.length + ' criados, ' +
  eventosAtualizados.length + ' atualizados'
);
  return {
    criados: eventosCriados,
    atualizados: eventosAtualizados
  };
}

function criarEventoVistoria_(dados, calendar) {
  const ARQUIVO = 'Calendar.gs';
  const FUNCAO = 'criarEventoVistoria_';

  const tituloEvento = montarTituloEvento_(dados);
  const descricao = montarDescricaoEvento_(dados);

  const inicio = new Date(dados.dataRevisao);
  const fim = new Date(inicio);
  fim.setHours(fim.getHours() + 2);

  const maxTentativas = 5;
  let tentativa = 0;

  while (tentativa < maxTentativas) {
    try {
      const evento = calendar.createEvent(tituloEvento, inicio, fim, {
        description: descricao,
        location: 'Revisao de Vistoria - Produttivo'
      });

      aplicarCorEvento_(evento, dados.criticidade);
      configurarLembretes_(evento, dados.criticidade);

      Utilities.sleep(500);

      logInfo_(ARQUIVO, FUNCAO, 'Evento criado com sucesso', {
        idAtividade: dados.idAtividade,
        eventId: evento.getId(),
        titulo: tituloEvento
      });

      return {
        id: evento.getId(),
        titulo: tituloEvento,
        data: inicio
      };
    } catch (error) {
      const msg = String(error);

      if (
        msg.includes('too many calendars or calendar events in a short time') ||
        msg.includes('Service invoked too many times')
      ) {
        const esperaMs = Math.pow(2, tentativa) * 1000 + Math.floor(Math.random() * 500);

        logWarn_(ARQUIVO, FUNCAO, 'Rate limit no Calendar, aguardando nova tentativa', {
          idAtividade: dados.idAtividade,
          tentativa: tentativa + 1,
          maxTentativas: maxTentativas,
          esperaMs: esperaMs
        });

        Utilities.sleep(esperaMs);
        tentativa++;
        continue;
      }

      logErro_(ARQUIVO, FUNCAO, 'Erro ao criar evento', {
        idAtividade: dados.idAtividade,
        mensagemErro: error.message || String(error),
        stack: error.stack || ''
      });

      return null;
    }
  }

  logErro_(ARQUIVO, FUNCAO, 'Falha apos varias tentativas para criar evento', {
    idAtividade: dados.idAtividade,
    maxTentativas: maxTentativas
  });

  return null;
}

function limparEventosAntigos() {
  const shVit = getSheet_('Vistorias_Tratadas');
  if (!shVit || shVit.getLastRow() < 2) {
    Logger.log('Nenhuma linha para limpar.');
    return 0;
  }

  const config = getConfig_();
  const calendarId = config.calendarId || 'primary';

  const calendar = CalendarApp.getCalendarById(calendarId);
  if (!calendar) {
    throw new Error('Calendário não encontrado: ' + calendarId);
  }

  const dados = shVit.getDataRange().getValues();
  const cabecalho = dados[0];

  const idxEventId = cabecalho.indexOf('event_id');
  const idxProxRevisao = cabecalho.indexOf('proxima_revisao');

  if (idxEventId === -1 || idxProxRevisao === -1) {
    Logger.log('Colunas event_id ou proxima_revisao não encontradas.');
    return 0;
  }

  const agora = new Date();
  let removidos = 0;

  for (let i = 1; i < dados.length; i++) {
    const linha = dados[i];
    const eventId = linha[idxEventId];
    const proxRevisao = normalizarData_(linha[idxProxRevisao]);

    if (!eventId || !proxRevisao) continue;
    if (proxRevisao >= agora) continue;

    const evento = calendar.getEventById(eventId);
    if (evento) {
      evento.deleteEvent();
      removidos++;
    }

    shVit.getRange(i + 1, idxEventId + 1).clearContent();
  }

  Logger.log('Removidos ' + removidos + ' eventos antigos do Calendar');
  return removidos;
}

function sincronizarCalendarCompleto() {
  Logger.log('Iniciando sincronização completa Calendar...');
  const removidos = limparEventosAntigos();
  const resultado = criarEventosCalendar();

  Logger.log(
    'Sincronização concluída. Removidos: ' + removidos +
    ', Criados: ' + resultado.criados.length +
    ', Atualizados: ' + resultado.atualizados.length
  );

  return {
    removidos: removidos,
    criados: resultado.criados,
    atualizados: resultado.atualizados
  };
}

function aplicarCorEvento_(evento, criticidade) {
  const crit = normalizarTexto_(criticidade);

  if (crit === 'alta') {
    evento.setColor(CalendarApp.EventColor.RED);
    return;
  }

  if (crit === 'media' || crit === 'média') {
    evento.setColor(CalendarApp.EventColor.YELLOW);
    return;
  }

  evento.setColor(CalendarApp.EventColor.GREEN);
}

function configurarLembretes_(evento, criticidade) {
  const crit = normalizarTexto_(criticidade);
  if (crit === 'baixa') return;

  evento.addPopupReminder(1440);
}

function montarTituloEvento_(dados) {
  return '🔄 ' + dados.criticidade + ': ' + dados.titulo + ' (' + dados.idAtividade + ')';
}

function montarDescricaoEvento_(dados) {
  return [
    '📋 Vistoria: ' + dados.titulo,
    '🆔 ID: ' + dados.idAtividade,
    '📅 Data Sincronização: ' + formatarData_(dados.dataSync),
    '📅 Data Finalização: ' + formatarData_(dados.dataFim),
    '📊 Conformidade: ' + Math.round((Number(dados.percentual) || 0) * 100) + '%',
    '🎯 Data Revisão: ' + formatarData_(dados.dataRevisao),
    '🔗 Hash: ' + dados.hash,
    '',
    '⚠️ ITENS CRÍTICOS PENDENTES - REVISAR!'
  ].join('\n');
}
function log_(nivel, mensagem) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let aba = ss.getSheetByName('Logs');

  if (!aba) {
    aba = ss.insertSheet('Logs');
    aba.appendRow(['Data/Hora', 'Nível', 'Mensagem']);
  }

  const dataFormatada = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    'dd/MM/yyyy, HH:mm:ss'
  );

  aba.appendRow([dataFormatada, nivel, mensagem]);

  
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

