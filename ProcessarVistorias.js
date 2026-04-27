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

    logInfo_(ARQUIVO, FUNCAO, 'Dados carregados', {
      linhasBaseAtividades: atVals.length,
      linhasBaseFormularios: foVals.length
    });

    if (atVals.length < 2) {
      limparAba_(shVit);
      limparAba_(shPen);

      logWarn_(ARQUIVO, FUNCAO, 'Base_Atividades sem dados. Abas de saida foram limpas.');
      return;
    }

    const cabAt = atVals[0];
    const cabFo = foVals.length ? foVals[0] : [];

    const idxTitulo = cabAt.indexOf('Título');
    const idxForm = cabAt.indexOf('Formulário');
    const idxDataSync = cabAt.indexOf('Cliente ou Local');
    const idxExecutores = cabAt.indexOf('Executores');
    const idxStatus = cabAt.indexOf('Status');
    const idxDataIni = cabAt.indexOf('Data e hora inicial');
    const idxDataFim = cabAt.indexOf('Data e hora final');
    const idxIdAtividade = cabAt.indexOf('ID da Atividade');

    const idxAtivoFo = cabFo.indexOf('Atividade');
    const idxLocalFo = cabFo.indexOf('Local/Cliente');
    const idxPont = cabFo.indexOf('Pontuação');
    const idxPontMax = cabFo.indexOf('Pontuação máxima');

    if (idxTitulo === -1) {
      logErro_(ARQUIVO, FUNCAO, 'Coluna obrigatoria ausente em Base_Atividades', {
        coluna: 'Título'
      });
      throw new Error('Coluna "Título" não encontrada em Base_Atividades.');
    }

    if (idxIdAtividade === -1) {
      logErro_(ARQUIVO, FUNCAO, 'Coluna obrigatoria ausente em Base_Atividades', {
        coluna: 'ID da Atividade'
      });
      throw new Error('Coluna "ID da Atividade" não encontrada em Base_Atividades.');
    }

    if (idxDataSync === -1) {
      logErro_(ARQUIVO, FUNCAO, 'Coluna obrigatoria ausente em Base_Atividades', {
        coluna: 'Cliente ou Local'
      });
      throw new Error('Coluna "Cliente ou Local" nao encontrada em Base_Atividades.');
    }

    logWarn_(ARQUIVO, FUNCAO, 'Coluna "Cliente ou Local" esta sendo usada como data_sincronizacao. Validar se este nome de coluna esta correto.', {
      colunaMapeada: 'Cliente ou Local'
    });

    if (idxAtivoFo === -1) {
      logErro_(ARQUIVO, FUNCAO, 'Coluna obrigatoria ausente em Base_Formularios', {
        coluna: 'Atividade'
      });
      throw new Error('Coluna "Atividade" não encontrada em Base_Formularios.');
    }

    if (idxPont === -1) {
      logErro_(ARQUIVO, FUNCAO, 'Coluna obrigatoria ausente em Base_Formularios', {
        coluna: 'Pontuação'
      });
      throw new Error('Coluna "Pontuação" não encontrada em Base_Formularios.');
    }

    if (idxPontMax === -1) {
      logErro_(ARQUIVO, FUNCAO, 'Coluna obrigatoria ausente em Base_Formularios', {
        coluna: 'Pontuação máxima'
      });
      throw new Error('Coluna "Pontuação máxima" não encontrada em Base_Formularios.');
    }

    const criticosLabels = [
      'HÁ MENORES TRABALHANDO? SE SIM, QUANTOS E QUAIS FUNÇÕES.',
      'CONDIÇÕES SATISFATÓRIAS DE ORDEM, ARRUMAÇÃO E LIMPEZA',
      'EQUIPAMENTO DE ÁGUA POTÁVEL',
      'INSTALAÇÕES ELÉTRICAS',
      'CONDIÇÃO DOS BANHEIROS (FUNCIONAMENTO E HIGIENE)'
    ];

    const idxCriticos = criticosLabels.map(function(lbl) {
      return cabFo.indexOf(lbl);
    });

    logInfo_(ARQUIVO, FUNCAO, 'Campos criticos mapeados', {
      totalCamposCriticos: criticosLabels.length,
      encontrados: idxCriticos.filter(function(i) { return i >= 0; }).length
    });

    const mapaForm = {};

    for (let i = 1; i < foVals.length; i++) {
      const linha = foVals[i];
      const chaveAtv = linha[idxAtivoFo] || linha[idxLocalFo] || '';
      if (!chaveAtv) continue;

      const pont = parseNumber_(linha[idxPont]);
      const pontMax = parseNumber_(linha[idxPontMax]);
      let naoConformes = 0;
      let criticosIrregulares = 0;

      idxCriticos.forEach(function(idx) {
        if (idx < 0) return;
        const val = String(linha[idx] || '').toLowerCase();
        if (!val) return;

        if (
          val.indexOf('não conforme') !== -1 ||
          val.indexOf('nao conforme') !== -1 ||
          val.indexOf('irregular') !== -1
        ) {
          naoConformes++;
          criticosIrregulares++;
        }
      });

      mapaForm[chaveAtv] = {
        pontuacao: pont,
        pontuacaoMaxima: pontMax,
        naoConformes: naoConformes,
        criticosIrregulares: criticosIrregulares
      };
    }

    logInfo_(ARQUIVO, FUNCAO, 'Mapa de formularios montado', {
      totalChavesFormulario: Object.keys(mapaForm).length
    });

    const headerVit = [
      'id_atividade', 'titulo_vistoria', 'formulario', 'data_sincronizacao', 'executor',
      'status', 'data_programada', 'data_finalizacao', 'chave_local',
      'pontuacao', 'pontuacao_maxima', 'percentual_conformidade',
      'qtde_nao_conformes', 'qtde_campos_criticos_irregulares',
      'criticidade', 'proxima_revisao', 'responsavel_email',
      'observacoes', 'hash_evento'
    ];

    const dadosVit = [];
    const dadosPen = [];
    const resumoCriticidade = { alta: 0, media: 0, baixa: 0 };

    for (let i = 1; i < atVals.length; i++) {
      if (i % 50 === 0) {
        logInfo_(ARQUIVO, FUNCAO, 'Processando lote de atividades', {
          linhaAtual: i + 1,
          totalLinhas: atVals.length - 1
        });
      }

      const linha = atVals[i];

      const titulo = linha[idxTitulo] || '';
      const formulario = linha[idxForm] || '';
      const dataSync = parseDate_(linha[idxDataSync]);
      const executor = linha[idxExecutores] || '';
      const status = linha[idxStatus] || '';
      const statusNorm = String(status || '').toLowerCase();
      const dataProg = parseDate_(linha[idxDataIni]);
      const dataFim = parseDate_(linha[idxDataFim]);
      const idAtv = linha[idxIdAtividade] || '';

      const chaveLocal = titulo + ' | ' + Utilities.formatDate(
        dataSync || new Date(),
        Session.getScriptTimeZone(),
        'dd/MM/yyyy'
      );

      const chaveForm = titulo || chaveLocal;

      if (!mapaForm[chaveForm]) {
        logWarn_(ARQUIVO, FUNCAO, 'Formulario nao encontrado para atividade', {
          linha: i + 1,
          idAtividade: idAtv,
          titulo: titulo,
          chaveForm: chaveForm
        });
      }

      const form = mapaForm[chaveForm] || {
        pontuacao: 0,
        pontuacaoMaxima: 0,
        naoConformes: 0,
        criticosIrregulares: 0
      };

      const pont = form.pontuacao;
      const pontMax = form.pontuacaoMaxima;
      const perc = pontMax > 0 ? pont / pontMax : 0;
      const naoConf = form.naoConformes;
      const critIrreg = form.criticosIrregulares;

      let criticidade = 'Baixa';
      if (perc < 0.7 || critIrreg > 0) criticidade = 'Alta';
      else if (perc < 0.9) criticidade = 'Média';

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

      if (statusNorm === 'concluída' || statusNorm === 'concluida') {
        if (perc < 1 || naoConf > 0 || critIrreg > 0) {
          dadosPen.push([
            idAtv,
            titulo,
            Utilities.formatDate(
              dataSync || new Date(),
              Session.getScriptTimeZone(),
              'dd/MM/yyyy'
            ),
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
      'id_atividade', 'titulo_vistoria', 'data_sincronizacao', 'criticidade',
      'percentual_conformidade', 'qtde_nao_conformes', 'proxima_revisao',
      'acao_recomendada', 'detalhes'
    ], dadosPen);

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