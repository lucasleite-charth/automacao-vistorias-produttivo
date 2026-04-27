var REGRAS_FORMULARIO = {
  camposCriticos: [
    'HÁ MENORES TRABALHANDO? SE SIM, QUANTOS E QUAIS FUNÇÕES.',
    'CONDIÇÕES SATISFATÓRIAS DE ORDEM, ARRUMAÇÃO E LIMPEZA',
    'EQUIPAMENTO DE ÁGUA POTÁVEL',
    'INSTALAÇÕES ELÉTRICAS',
    'CONDIÇÃO DO REFEITÓRIO',
    'CONDIÇÃO DOS BANHEIROS (FUNCIONAMENTO E HIGIENE)'
  ],
  termosNaoConforme: [
    'não conforme',
    'nao conforme',
    'irregular',
    'inadequado',
    'inadequada',
    'faltando',
    'falta',
    'sem',
    'ausente',
    'improvisada',
    'improvisado'
  ],
  termosNeutros: [
    'conforme',
    'n/a',
    'na',
    'não se aplica',
    'nao se aplica',
    'ok'
  ]
};

function obterMapaCabecalho_(headerRow) {
  var mapa = {};
  for (var i = 0; i < headerRow.length; i++) {
    mapa[String(headerRow[i] || '').trim()] = i;
  }
  return mapa;
}

function normalizarTexto_(valor) {
  if (valor === null || typeof valor === 'undefined') return '';
  return String(valor)
    .toLowerCase()
    .trim()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');
}

function textoIndicaNaoConformidade_(valor) {
  var txt = normalizarTexto_(valor);
  if (!txt) return false;

  for (var i = 0; i < REGRAS_FORMULARIO.termosNeutros.length; i++) {
    if (txt === normalizarTexto_(REGRAS_FORMULARIO.termosNeutros[i])) return false;
  }

  for (var j = 0; j < REGRAS_FORMULARIO.termosNaoConforme.length; j++) {
    if (txt.indexOf(normalizarTexto_(REGRAS_FORMULARIO.termosNaoConforme[j])) !== -1) return true;
  }

  return false;
}

function textoIndicaMenorIrregular_(valor, notas) {
  var resposta = normalizarTexto_(valor);
  var obs = normalizarTexto_(notas);

  if (!resposta && !obs) return false;
  if (resposta === 'nao' || resposta === 'não') return false;
  if (resposta === 'sim') return true;

  if (obs.indexOf('menor') !== -1) return true;
  if (obs.indexOf('crianca') !== -1 || obs.indexOf('criança') !== -1) return true;

  return false;
}

function calcularPercentualConformidade_(pontuacao, pontuacaoMaxima) {
  var pont = parseNumber_(pontuacao);
  var max = parseNumber_(pontuacaoMaxima);
  if (!max || max <= 0) return 0;
  return pont / max;
}

function avaliarCamposCriticosFormulario_(headerFormulario, linhaFormulario) {
  var mapa = obterMapaCabecalho_(headerFormulario);
  var totalNaoConformes = 0;
  var totalCriticosIrregulares = 0;
  var detalhes = [];

  for (var i = 0; i < REGRAS_FORMULARIO.camposCriticos.length; i++) {
    var campo = REGRAS_FORMULARIO.camposCriticos[i];
    var idxCampo = mapa[campo];
    var idxNotas = mapa[campo + ' - Notas'];

    if (typeof idxCampo === 'undefined') continue;

    var valor = linhaFormulario[idxCampo];
    var notas = typeof idxNotas !== 'undefined' ? linhaFormulario[idxNotas] : '';
    var irregular = false;

    if (campo === 'HÁ MENORES TRABALHANDO? SE SIM, QUANTOS E QUAIS FUNÇÕES.') {
      irregular = textoIndicaMenorIrregular_(valor, notas);
    } else {
      irregular = textoIndicaNaoConformidade_(valor) || textoIndicaNaoConformidade_(notas);
    }

    if (irregular) {
      totalNaoConformes++;
      totalCriticosIrregulares++;
      detalhes.push({
        campo: campo,
        valor: valor || '',
        notas: notas || ''
      });
    }
  }

  return {
    qtdeNaoConformes: totalNaoConformes,
    qtdeCamposCriticosIrregulares: totalCriticosIrregulares,
    detalhes: detalhes
  };
}

function definirCriticidade_(percentualConformidade, qtdeCriticosIrregulares, qtdeNaoConformes) {
  var perc = Number(percentualConformidade || 0);
  var criticos = Number(qtdeCriticosIrregulares || 0);
  var naoConformes = Number(qtdeNaoConformes || 0);

  if (criticos > 0) return 'Alta';
  if (perc < 0.7) return 'Alta';
  if (naoConformes >= 3) return 'Alta';

  if (perc < 0.9) return 'Média';
  if (naoConformes >= 1) return 'Média';

  return 'Baixa';
}

function calcularProximaRevisao_(dataBase, criticidade) {
  var data = parseDate_(dataBase);
  if (!data) data = new Date();

  var dias = getPrazoRevisaoPorCriticidade_(criticidade);
  var proxima = new Date(data);
  proxima.setDate(proxima.getDate() + dias);
  return proxima;
}

function definirAcaoRecomendada_(criticidade, qtdeCriticosIrregulares, qtdeNaoConformes) {
  if (criticidade === 'Alta' && qtdeCriticosIrregulares > 0) {
    return 'Agendar revisão prioritária e cobrar regularização imediata';
  }
  if (criticidade === 'Alta') {
    return 'Agendar revisão em curto prazo e acompanhar plano de ação';
  }
  if (criticidade === 'Média') {
    return 'Enviar lembrete e acompanhar correções pendentes';
  }
  if (qtdeNaoConformes > 0) {
    return 'Monitorar pendências e validar no próximo ciclo';
  }
  return 'Sem ação imediata';
}

function montarDetalhesPendencia_(avaliacao) {
  if (!avaliacao || !avaliacao.detalhes || !avaliacao.detalhes.length) return '';
  return avaliacao.detalhes.map(function (item) {
    var partes = [item.campo];
    if (item.valor) partes.push('Valor: ' + item.valor);
    if (item.notas) partes.push('Notas: ' + item.notas);
    return partes.join(' | ');
  }).join('\n');
}

function montarObjetoFormulario_(headerFormulario, linhaFormulario) {
  var mapa = obterMapaCabecalho_(headerFormulario);

  var pontuacao = linhaFormulario[mapa['Pontuação']];
  var pontuacaoMaxima = linhaFormulario[mapa['Pontuação máxima']];
  var atividade = linhaFormulario[mapa['Atividade']];
  var localCliente = linhaFormulario[mapa['Local/Cliente']];
  var ativo = linhaFormulario[mapa['Ativo']];
  var dataHora = linhaFormulario[mapa['DATA E HORA']];

  var avaliacao = avaliarCamposCriticosFormulario_(headerFormulario, linhaFormulario);
  var percentual = calcularPercentualConformidade_(pontuacao, pontuacaoMaxima);
  var criticidade = definirCriticidade_(
    percentual,
    avaliacao.qtdeCamposCriticosIrregulares,
    avaliacao.qtdeNaoConformes
  );

  return {
    atividade: atividade || '',
    localCliente: localCliente || '',
    ativo: ativo || '',
    dataHora: dataHora || '',
    pontuacao: parseNumber_(pontuacao),
    pontuacaoMaxima: parseNumber_(pontuacaoMaxima),
    percentualConformidade: percentual,
    qtdeNaoConformes: avaliacao.qtdeNaoConformes,
    qtdeCamposCriticosIrregulares: avaliacao.qtdeCamposCriticosIrregulares,
    criticidade: criticidade,
    detalhesPendencia: montarDetalhesPendencia_(avaliacao)
  };
}

function montarIndiceFormulariosPorAtividade_(headerFormulario, linhasFormulario) {
  var indice = {};

  for (var i = 0; i < linhasFormulario.length; i++) {
    var obj = montarObjetoFormulario_(headerFormulario, linhasFormulario[i]);
    var chavePrincipal = String(obj.atividade || '').trim();
    var chaveSecundaria = String(obj.localCliente || '').trim();

    if (chavePrincipal) indice[chavePrincipal] = obj;
    if (chaveSecundaria && !indice[chaveSecundaria]) indice[chaveSecundaria] = obj;
  }

  return indice;
}

function obterFormularioRelacionado_(tituloAtividade, localAtividade, indiceFormularios) {
  var titulo = String(tituloAtividade || '').trim();
  var local = String(localAtividade || '').trim();

  if (titulo && indiceFormularios[titulo]) return indiceFormularios[titulo];
  if (local && indiceFormularios[local]) return indiceFormularios[local];

  var tituloNorm = normalizarTexto_(titulo);
  var localNorm = normalizarTexto_(local);
  var chaves = Object.keys(indiceFormularios);

  for (var i = 0; i < chaves.length; i++) {
    var chave = chaves[i];
    var chaveNorm = normalizarTexto_(chave);

    if (tituloNorm && chaveNorm.indexOf(tituloNorm) !== -1) return indiceFormularios[chave];
    if (tituloNorm && tituloNorm.indexOf(chaveNorm) !== -1) return indiceFormularios[chave];
    if (localNorm && chaveNorm.indexOf(localNorm) !== -1) return indiceFormularios[chave];
    if (localNorm && localNorm.indexOf(chaveNorm) !== -1) return indiceFormularios[chave];
  }

  return null;
}

function statusEhConcluido_(status) {
  var s = normalizarTexto_(status);
  return s === 'concluida' || s === 'concluída' || s === 'finalizada' || s === 'finalizado';
}

function deveGerarPendencia_(status, percentualConformidade, qtdeNaoConformes, qtdeCriticosIrregulares) {
  if (!statusEhConcluido_(status)) return false;
  if (qtdeCriticosIrregulares > 0) return true;
  if (qtdeNaoConformes > 0) return true;
  if (Number(percentualConformidade || 0) < 1) return true;
  return false;
}