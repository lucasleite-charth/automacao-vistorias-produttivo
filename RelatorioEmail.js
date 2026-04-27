//RelatorioEmail
function enviarRelatorioConsolidadoVistorias() {
  const ARQUIVO = 'RelatorioEmail.gs';
  const FUNCAO = 'enviarRelatorioConsolidadoVistorias';
  const NOME_ABA = 'Vistorias_Tratadas';

  logInfo_(ARQUIVO, FUNCAO, 'Iniciando envio de relatorio consolidado');

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const aba = ss.getSheetByName(NOME_ABA);

    if (!aba) {
      logErro_(ARQUIVO, FUNCAO, 'Aba nao encontrada', { nomeAba: NOME_ABA });
      throw new Error('A aba "' + NOME_ABA + '" nao foi encontrada.');
    }

    const dados = aba.getDataRange().getValues();
    if (!dados || dados.length < 2) {
      logWarn_(ARQUIVO, FUNCAO, 'Aba sem dados para envio');
      return;
    }

    const cab = dados[0];

    const idx = {
      id_atividade: cab.indexOf('id_atividade'),
      titulo_vistoria: cab.indexOf('titulo_vistoria'),
      executor: cab.indexOf('executor'),
      data_sincronizacao: cab.indexOf('data_sincronizacao'),
      percentual_conformidade: cab.indexOf('percentual_conformidade'),
      qtde_nao_conformes: cab.indexOf('qtde_nao_conformes'),
      qtde_campos_criticos_irregulares: cab.indexOf('qtde_campos_criticos_irregulares'),
      criticidade: cab.indexOf('criticidade'),
      proxima_revisao: cab.indexOf('proxima_revisao')
    };

    const obrigatorias = [
      'id_atividade',
      'titulo_vistoria',
      'executor',
      'data_sincronizacao',
      'percentual_conformidade',
      'qtde_nao_conformes',
      'qtde_campos_criticos_irregulares',
      'criticidade',
      'proxima_revisao'
    ];

    const faltando = obrigatorias.filter(function(nome) {
      return idx[nome] < 0;
    });

    if (faltando.length > 0) {
      logErro_(ARQUIVO, FUNCAO, 'Colunas obrigatorias ausentes', {
        faltando: faltando
      });
      throw new Error('Colunas obrigatorias nao encontradas: ' + faltando.join(', '));
    }

    const destinatario = getEmailParaDefault_();
    const cc = getEmailCcDefault_();

    if (!destinatario) {
      logErro_(ARQUIVO, FUNCAO, 'email_para nao configurado na aba Config');
      throw new Error('Configure "email_para" na aba Config.');
    }

    const hoje = new Date();
    const linhasRelatorio = [];
    const resumo = {
      alta: 0,
      media: 0,
      baixa: 0
    };

    for (let i = 1; i < dados.length; i++) {
      const linha = dados[i];

      const idAtividade = obterTextoRel_(linha[idx.id_atividade]);
      const titulo = obterTextoRel_(linha[idx.titulo_vistoria]);
      const executor = obterTextoRel_(linha[idx.executor]);
      const dataSync = linha[idx.data_sincronizacao];
      const percentual = Number(linha[idx.percentual_conformidade]) || 0;
      const naoConf = Number(linha[idx.qtde_nao_conformes]) || 0;
      const criticos = Number(linha[idx.qtde_campos_criticos_irregulares]) || 0;
      const criticidade = obterTextoRel_(linha[idx.criticidade]) || 'Baixa';
      const vencimento = linha[idx.proxima_revisao];

      if (!idAtividade || !titulo || !vencimento) {
        continue;
      }

      const criticidadeNorm = normalizarTextoRel_(criticidade);
      if (criticidadeNorm === 'alta') resumo.alta++;
      else if (criticidadeNorm === 'media') resumo.media++;
      else resumo.baixa++;

      const diasRestantes = calcularDiasRestantesRel_(vencimento, hoje);

      linhasRelatorio.push({
        idAtividade: idAtividade,
        titulo: titulo,
        executor: executor,
        dataSync: formatarDataRel_(dataSync),
        percentualTexto: Math.round(percentual * 100) + '%',
        percentualNumero: percentual,
        naoConf: naoConf,
        criticos: criticos,
        criticidade: criticidade,
        vencimento: formatarDataHoraRel_(vencimento),
        diasRestantes: diasRestantes
      });
    }

    if (linhasRelatorio.length === 0) {
      logWarn_(ARQUIVO, FUNCAO, 'Nenhuma vistoria valida encontrada para montar o relatorio');
      return;
    }

    linhasRelatorio.sort(function(a, b) {
      const ordem = { alta: 0, media: 1, baixa: 2 };
      const oa = ordem[normalizarTextoRel_(a.criticidade)] ?? 9;
      const ob = ordem[normalizarTextoRel_(b.criticidade)] ?? 9;
      if (oa !== ob) return oa - ob;
      return a.diasRestantes - b.diasRestantes;
    });

    const assunto = 'VISTORIAS PARA REVISÃO (' + linhasRelatorio.length + ') - ' + formatarDataHoraRel_(new Date());
    const corpoTexto = montarTextoRelatorioVistorias_(linhasRelatorio, resumo);
    const corpoHtml = montarHtmlRelatorioVistorias_(linhasRelatorio, resumo);

    const opcoes = {
      htmlBody: corpoHtml
    };
    if (cc) opcoes.cc = cc;

    GmailApp.sendEmail(destinatario, assunto, corpoTexto, opcoes);

    logInfo_(ARQUIVO, FUNCAO, 'Relatorio consolidado enviado com sucesso', {
      destinatario: destinatario,
      cc: cc,
      totalVistorias: linhasRelatorio.length,
      resumo: resumo
    });

  } catch (e) {
    logErro_(ARQUIVO, FUNCAO, 'Erro ao enviar relatorio consolidado', {
      mensagemErro: e.message,
      stack: e.stack
    });
    throw e;
  }
}

function montarHtmlRelatorioVistorias_(linhas, resumo) {
  const dataRelatorio = formatarDataHoraRel_(new Date());

  const linhasTabela = linhas.map(function(item) {
    const corPercentual = item.percentualNumero < 0.7 ? '#d32f2f' : (item.percentualNumero < 0.9 ? '#f9a825' : '#2e7d32');
    const corDias = item.diasRestantes <= 2 ? '#2e7d32' : '#555';

    return `
      <tr style="background:#f9f9f9;">
        <td style="padding:8px; border:1px solid #ccc;">${escapeHtmlRel_(item.idAtividade)}</td>
        <td style="padding:8px; border:1px solid #ccc;">${escapeHtmlRel_(item.titulo)}</td>
        <td style="padding:8px; border:1px solid #ccc;">${escapeHtmlRel_(item.executor)}</td>
        <td style="padding:8px; border:1px solid #ccc;">${escapeHtmlRel_(item.dataSync)}</td>
        <td style="padding:8px; border:1px solid #ccc; color:${corPercentual}; font-weight:bold;">${escapeHtmlRel_(item.percentualTexto)}</td>
        <td style="padding:8px; border:1px solid #ccc; text-align:center;">${item.naoConf}</td>
        <td style="padding:8px; border:1px solid #ccc; text-align:center; color:#d32f2f;">${item.criticos}</td>
        <td style="padding:8px; border:1px solid #ccc;">
          ${escapeHtmlRel_(item.vencimento)}<br>
          <span style="color:${corDias}; font-weight:bold;">(${item.diasRestantes} dias)</span>
        </td>
      </tr>
    `;
  }).join('');

  return `
  <div style="font-family: Arial, sans-serif; background:#f4f6f8; padding:20px;">
    <div style="max-width:800px; margin:auto; background:#ffffff; border-radius:10px; overflow:hidden;">

      <div style="padding:20px; text-align:center;">
        <h2 style="color:#d32f2f; margin:0;"><strong> VISTORIAS PARA REVISAO</strong></h2>
      </div>

      <div style="padding:0 24px 20px 24px; color:#333;">
        <p><strong>Data do relatório:</strong> ${escapeHtmlRel_(dataRelatorio)}</p>
        <p><strong>Total de vistorias:</strong> ${linhas.length}</p>
      </div>

      <div style="padding:0 24px;">
        <h3 style="color:#444;"><b> RESUMO POR CRITICIDADE:</b></h3>

        <table width="100%" style="border-collapse:collapse; margin-top:10px;">
          <tr style="background:#e0e0e0;">
            <th style="padding:10px; text-align:left; border:1px solid #ccc;">Criticidade</th>
            <th style="padding:10px; border:1px solid #ccc;">Quantidade</th>
          </tr>
          <tr>
            <td style="padding:10px; color:#d32f2f; border:1px solid #ccc;"><strong>Alta</strong></td>
            <td style="text-align:center; border:1px solid #ccc;">${resumo.alta}</td>
          </tr>
          <tr>
            <td style="padding:10px; color:#f9a825; border:1px solid #ccc;"><strong>Média</strong></td>
            <td style="text-align:center; border:1px solid #ccc;">${resumo.media}</td>
          </tr>
          <tr>
            <td style="padding:10px; color:#2e7d32; border:1px solid #ccc;"><strong>Baixa</strong></td>
            <td style="text-align:center; border:1px solid #ccc;">${resumo.baixa}</td>
          </tr>
        </table>
      </div>

      <div style="padding:20px 24px;">
        <h3 style="color:#444;">⚠️ DETALHAMENTO DAS VISTORIAS:</h3>

        <table width="100%" style="border-collapse:collapse; margin-top:10px; font-size:13px;">
          <tr style="background:#333; color:#fff;">
            <th style="padding:8px; border:1px solid #555;">ID</th>
            <th style="padding:8px; border:1px solid #555;">Vistoria</th>
            <th style="padding:8px; border:1px solid #555;">Executor</th>
            <th style="padding:8px; border:1px solid #555;">Data Sync</th>
            <th style="padding:8px; border:1px solid #555;">% Conf.</th>
            <th style="padding:8px; border:1px solid #555;">Não Conf.</th>
            <th style="padding:8px; border:1px solid #555;">Críticos</th>
            <th style="padding:8px; border:1px solid #555;">Vencimento</th>
          </tr>
          ${linhasTabela}
        </table>
      </div>

      <div style="background:#fff3cd; border-left:5px solid #f9a825; margin:20px; padding:15px;">
        <h4 style="margin-top:0;">⚠️ AÇÃO REQUERIDA:</h4>
        <ul style="margin:0; padding-left:20px;">
          <li><strong>Contatar executor para revisão dos itens não conformes</strong></li>
          <li><strong>Agendar nova vistoria para as pendências críticas</strong></li>
          <li><strong>Documentar ações corretivas implementadas</strong></li>
          <li><strong>✅ Atualizar status no Produttivo após correções</strong></li>
        </ul>
      </div>

      <div style="text-align:center; font-size:12px; color:#777; padding:15px;">
        Relatório automático do sistema Vistorias Produttivo | ${escapeHtmlRel_(dataRelatorio)}
      </div>
    </div>
  </div>
  `;
}

function montarTextoRelatorioVistorias_(linhas, resumo) {
  const cabecalho = [
    'VISTORIAS PARA REVISAO',
    '',
    'Data do relatorio: ' + formatarDataHoraRel_(new Date()),
    'Total de vistorias: ' + linhas.length,
    '',
    'RESUMO POR CRITICIDADE',
    'Alta: ' + resumo.alta,
    'Media: ' + resumo.media,
    'Baixa: ' + resumo.baixa,
    '',
    'DETALHAMENTO'
  ];

  const detalhes = linhas.map(function(item) {
    return [
      'ID: ' + item.idAtividade,
      'Vistoria: ' + item.titulo,
      'Executor: ' + item.executor,
      'Data Sync: ' + item.dataSync,
      'Conformidade: ' + item.percentualTexto,
      'Nao conformes: ' + item.naoConf,
      'Criticos: ' + item.criticos,
      'Vencimento: ' + item.vencimento + ' (' + item.diasRestantes + ' dias)',
      '---'
    ].join('\n');
  });

  return cabecalho.concat(detalhes).join('\n');
}

function formatarDataRel_(valor) {
  if (!valor) return '';
  if (Object.prototype.toString.call(valor) === '[object Date]' && !isNaN(valor)) {
    return Utilities.formatDate(valor, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  }
  return String(valor).trim();
}

function formatarDataHoraRel_(valor) {
  if (!valor) return '';
  if (Object.prototype.toString.call(valor) === '[object Date]' && !isNaN(valor)) {
    return Utilities.formatDate(valor, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
  }
  return String(valor).trim();
}

function obterTextoRel_(valor) {
  if (valor === null || valor === undefined) return '';
  return String(valor).trim();
}

function normalizarTextoRel_(valor) {
  return String(valor || '')
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');
}

function calcularDiasRestantesRel_(dataFutura, dataBase) {
  const futura = new Date(dataFutura);
  futura.setHours(0, 0, 0, 0);

  const base = new Date(dataBase);
  base.setHours(0, 0, 0, 0);

  return Math.ceil((futura - base) / (1000 * 60 * 60 * 24));
}

function escapeHtmlRel_(texto) {
  return String(texto || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}
