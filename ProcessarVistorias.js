function processarVistorias() {
  const ARQUIVO = 'ProcessarVistorias.gs';
  const FUNCAO = 'processarVistorias';

  logInfo_(ARQUIVO, FUNCAO, 'Iniciando processamento de vistorias');

  try {
    const shAt = getSheet_('Base_Atividades');
    const shFo = getSheet_('Base_Formularios');
    const shVit = getSheet_('Vistorias_Tratadas');
    const shPen = getSheet_('Pendencias');

    const atVals = shAt.getDataRange().getValues();
    const foVals = shFo.getDataRange().getValues();

    if (atVals.length < 2) {
      limparAba_(shVit);
      limparAba_(shPen);
      logWarn_(ARQUIVO, FUNCAO, 'Base_Atividades sem dados.');
      return;
    }

    const cabAt = atVals[0];
    const cabFo = foVals.length ? foVals[0] : [];

    const idxTitulo = encontrarColuna_(cabAt, ['Título', 'Titulo']);
    const idxForm = encontrarColuna_(cabAt, ['Formulário', 'Formulario']);
    const idxExecutores = encontrarColuna_(cabAt, ['Executores']);
    const idxStatus = encontrarColuna_(cabAt, ['Status']);
    const idxDataIni = encontrarColuna_(cabAt, ['Data e hora inicial']);
    const idxDataFim = encontrarColuna_(cabAt, ['Data e hora final']);
    const idxIdAtividade = encontrarColuna_(cabAt, ['ID da Atividade', 'ID Atividade', 'id_atividade']);

    const idxIdFo = encontrarColuna_(cabFo, ['ID da Atividade', 'ID Atividade', 'id_atividade']);
    const idxAtivoFo = encontrarColuna_(cabFo, ['Atividade']);
    const idxLocalFo = encontrarColuna_(cabFo, ['Local/Cliente', 'Local Cliente']);
    const idxDataSync = encontrarColuna_(cabFo, [
      'Data de sincronização',
      'Data de sincronizacao',
      'Data Sincronização',
      'Data Sincronizacao',
      'Sincronização',
      'Sincronizacao',
      'data_sincronizacao'
    ]);
    const idxPont = encontrarColuna_(cabFo, ['Pontuação', 'Pontuacao']);
    const idxPontMax = encontrarColuna_(cabFo, ['Pontuação máxima', 'Pontuacao maxima']);

    validarIndice_(idxTitulo, 'Título', 'Base_Atividades');
    validarIndice_(idxStatus, 'Status', 'Base_Atividades');
    validarIndice_(idxIdAtividade, 'ID da Atividade', 'Base_Atividades');
    validarIndice_(idxDataSync, 'Data de sincronizacao', 'Base_Formularios');
    validarIndice_(idxPont, 'Pontuação', 'Base_Formularios');
    validarIndice_(idxPontMax, 'Pontuação máxima', 'Base_Formularios');

    if (idxIdFo === -1 && idxAtivoFo === -1) {
      throw new Error('Nao encontrei "ID da Atividade" nem "Atividade" em Base_Formularios.');
    }

    const criticosLabels = [
      'HÁ MENORES TRABALHANDO? SE SIM, QUANTOS E QUAIS FUNÇÕES.',
      'CONDIÇÕES SATISFATÓRIAS DE ORDEM, ARRUMAÇÃO E LIMPEZA',
      'EQUIPAMENTO DE ÁGUA POTÁVEL',
      'INSTALAÇÕES ELÉTRICAS',
      'CONDIÇÃO DOS BANHEIROS (FUNCIONAMENTO E HIGIENE)'
    ];

    const idxCriticos = criticosLabels.map(function(lbl) {
      return encontrarColuna_(cabFo, [lbl]);
    });

    const mapaForm = {};

    for (let i = 1; i < foVals.length; i++) {
      const linhaFo = foVals[i];

      const chaveFormulario = idxIdFo >= 0
        ? obterTextoProcessar_(linhaFo[idxIdFo])
        : obterTextoProcessar_(linhaFo[idxAtivoFo]);

      if (!chaveFormulario) continue;

      const pontFo = parseNumber_(linhaFo[idxPont]);
      const pontMaxFo = parseNumber_(linhaFo[idxPontMax]);
      let naoConformesFo = 0;
      let criticosIrregularesFo = 0;

      idxCriticos.forEach(function(idx) {
        if (idx < 0) return;

        const val = String(linhaFo[idx] || '').toLowerCase();
        if (!val) return;

        if (
          val.indexOf('não conforme') !== -1 ||
          val.indexOf('nao conforme') !== -1 ||
          val.indexOf('irregular') !== -1
        ) {
          naoConformesFo++;
          criticosIrregularesFo++;
        }
      });

      mapaForm[chaveFormulario] = {
        pontuacao: pontFo,
        pontuacaoMaxima: pontMaxFo,
        naoConformes: naoConformesFo,
        criticosIrregulares: criticosIrregularesFo,
        dataSincronizacao: parseDate_(linhaFo[idxDataSync]),
        atividade: idxAtivoFo >= 0 ? obterTextoProcessar_(linhaFo[idxAtivoFo]) : '',
        local: idxLocalFo >= 0 ? obterTextoProcessar_(linhaFo[idxLocalFo]) : ''
      };
    }

    logInfo_(ARQUIVO, FUNCAO, 'Mapa de formularios montado', {
      totalChavesFormulario: Object.keys(mapaForm).length,
      chaveUsada: idxIdFo >= 0 ? 'ID da Atividade' : 'Atividade'
    });

    const headerVit = [
      'id_atividade',
      'titulo_vistoria',
      'formulario',
      'data_sincronizacao',
      'executor',
      'status',
      'data_programada',
      'data_finalizacao',
      'chave_local',
      'pontuacao',
      'pontuacao_maxima',
      'percentual_conformidade',
      'qtde_nao_conformes',
      'qtde_campos_criticos_irregulares',
      'criticidade',
      'proxima_revisao',
      'responsavel_email',
      'observacoes',
      'hash_evento'
    ];

    const dadosVit = [];
    const dadosPen = [];
    const resumoCriticidade = { alta: 0, media: 0, baixa: 0 };

    for (let i = 1; i < atVals.length; i++) {
      const linhaAt = atVals[i];

      const titulo = obterTextoProcessar_(linhaAt[idxTitulo]);
      const formulario = idxForm >= 0 ? obterTextoProcessar_(linhaAt[idxForm]) : '';
      const executor = idxExecutores >= 0 ? obterTextoProcessar_(linhaAt[idxExecutores]) : '';
      const status = obterTextoProcessar_(linhaAt[idxStatus]);
      const statusNorm = normalizarProcessar_(status);
      const dataProg = idxDataIni >= 0 ? parseDate_(linhaAt[idxDataIni]) : '';
      const dataFim = idxDataFim >= 0 ? parseDate_(linhaAt[idxDataFim]) : '';
      const idAtv = obterTextoProcessar_(linhaAt[idxIdAtividade]);

      const chaveForm = idxIdFo >= 0 ? idAtv : titulo;

      const form = mapaForm[chaveForm] || {
        pontuacao: 0,
        pontuacaoMaxima: 0,
        naoConformes: 0,
        criticosIrregulares: 0,
        dataSincronizacao: '',
        atividade: '',
        local: ''
      };

      if (!mapaForm[chaveForm]) {
        logWarn_(ARQUIVO, FUNCAO, 'Formulario nao encontrado para atividade', {
          linha: i + 1,
          idAtividade: idAtv,
          titulo: titulo,
          chaveForm: chaveForm
        });
      }

      const dataSync = form.dataSincronizacao || '';
      const pont = form.pontuacao;
      const pontMax = form.pontuacaoMaxima;
      const perc = pontMax > 0 ? pont / pontMax : 0;
      const naoConf = form.naoConformes;
      const critIrreg = form.criticosIrregulares;

      const dataChave = dataSync || dataFim || dataProg || new Date();

      const chaveLocal = titulo + ' | ' + Utilities.formatDate(
        dataChave,
        Session.getScriptTimeZone(),
        'dd/MM/yyyy'
      );

      let criticidade = 'Baixa';
      if (perc < 0.7 || critIrreg > 0) {
        criticidade = 'Alta';
      } else if (perc < 0.9) {
        criticidade = 'Média';
      }

      if (criticidade === 'Alta') resumoCriticidade.alta++;
      else if (criticidade === 'Média') resumoCriticidade.media++;
      else resumoCriticidade.baixa++;

      const diasPrazo = getPrazoRevisaoPorCriticidade_(criticidade);
      const baseData = dataFim || dataSync || dataProg || new Date();
      const proxima = new Date(baseData);
      proxima.setDate(proxima.getDate() + diasPrazo);

      const responsavelEmail = getEmailParaDefault_();
      const hashEvento = gerarHashEvento_(idAtv, proxima);

      dadosVit.push([
        idAtv,
        titulo,
        formulario,
        dataSync,
        executor,
        status,
        dataProg,
        dataFim,
        chaveLocal,
        pont,
        pontMax,
        perc,
        naoConf,
        critIrreg,
        criticidade,
        proxima,
        responsavelEmail,
        '',
        hashEvento
      ]);

      if (statusNorm === 'concluida') {
        if (perc < 1 || naoConf > 0 || critIrreg > 0) {
          dadosPen.push([
            idAtv,
            titulo,
            dataSync ? Utilities.formatDate(dataSync, Session.getScriptTimeZone(), 'dd/MM/yyyy') : '',
            criticidade,
            perc,
            naoConf,
            proxima,
            'Revisar itens não conformes',
            ''
          ]);
        }
      }
    }

    setHeaderAndData_(shVit, headerVit, dadosVit);

    setHeaderAndData_(shPen, [
      'id_atividade',
      'titulo_vistoria',
      'data_sincronizacao',
      'criticidade',
      'percentual_conformidade',
      'qtde_nao_conformes',
      'proxima_revisao',
      'acao_recomendada',
      'detalhes'
    ], dadosPen);

    formatarAbaVistoriasTratadas_(shVit);

    logInfo_(ARQUIVO, FUNCAO, 'Resumo de criticidade', resumoCriticidade);

    logInfo_(ARQUIVO, FUNCAO, 'Processamento concluido com sucesso', {
      totalVistoriasTratadas: dadosVit.length,
      totalPendencias: dadosPen.length
    });

  } catch (e) {
    logErro_(ARQUIVO, FUNCAO, 'Erro geral ao processar vistorias', {
      mensagemErro: e.message,
      stack: e.stack
    });
    throw e;
  }
}

function validarIndice_(idx, coluna, aba) {
  if (idx === -1) {
    throw new Error('Coluna "' + coluna + '" nao encontrada em ' + aba + '.');
  }
}

function encontrarColuna_(cabecalho, nomesPossiveis) {
  const normalizar = function(valor) {
    return String(valor || '')
      .trim()
      .toLowerCase()
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '');
  };

  const cabNormalizado = cabecalho.map(normalizar);

  for (let i = 0; i < nomesPossiveis.length; i++) {
    const idx = cabNormalizado.indexOf(normalizar(nomesPossiveis[i]));
    if (idx >= 0) return idx;
  }

  return -1;
}

function obterTextoProcessar_(valor) {
  if (valor === null || valor === undefined) return '';
  return String(valor).trim();
}

function normalizarProcessar_(valor) {
  return String(valor || '')
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');
}

function formatarAbaVistoriasTratadas_(sh) {
  const ultimaLinha = sh.getLastRow();
  if (ultimaLinha < 2) return;

  const cab = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];

  const colDataSync = cab.indexOf('data_sincronizacao') + 1;
  const colDataProg = cab.indexOf('data_programada') + 1;
  const colDataFim = cab.indexOf('data_finalizacao') + 1;
  const colProxima = cab.indexOf('proxima_revisao') + 1;

  const colPont = cab.indexOf('pontuacao') + 1;
  const colPontMax = cab.indexOf('pontuacao_maxima') + 1;
  const colPercentual = cab.indexOf('percentual_conformidade') + 1;

  if (colDataSync > 0) sh.getRange(2, colDataSync, ultimaLinha - 1).setNumberFormat('dd/MM/yyyy HH:mm');
  if (colDataProg > 0) sh.getRange(2, colDataProg, ultimaLinha - 1).setNumberFormat('dd/MM/yyyy HH:mm');
  if (colDataFim > 0) sh.getRange(2, colDataFim, ultimaLinha - 1).setNumberFormat('dd/MM/yyyy HH:mm');
  if (colProxima > 0) sh.getRange(2, colProxima, ultimaLinha - 1).setNumberFormat('dd/MM/yyyy HH:mm');

  if (colPont > 0) sh.getRange(2, colPont, ultimaLinha - 1).setNumberFormat('0.00');
  if (colPontMax > 0) sh.getRange(2, colPontMax, ultimaLinha - 1).setNumberFormat('0.00');
  if (colPercentual > 0) sh.getRange(2, colPercentual, ultimaLinha - 1).setNumberFormat('0.00%');
}
