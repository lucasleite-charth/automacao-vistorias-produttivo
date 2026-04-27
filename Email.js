function prepararEmailsLembrete() {
  const ARQUIVO = 'Email.gs';
  const FUNCAO = 'prepararEmailsLembrete';
  const NOME_ABA = 'Emails_Enviar';

  logInfo_(ARQUIVO, FUNCAO, 'Iniciando preparacao de emails lembrete');

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const aba = ss.getSheetByName(NOME_ABA);

    if (!aba) {
      logErro_(ARQUIVO, FUNCAO, 'A aba de emails nao foi encontrada', {
        nomeAba: NOME_ABA
      });
      throw new Error('A aba "' + NOME_ABA + '" nao foi encontrada.');
    }

    const ultimaLinha = aba.getLastRow();
    const ultimaColuna = aba.getLastColumn();

    if (ultimaLinha < 2) {
      logWarn_(ARQUIVO, FUNCAO, 'A aba nao possui linhas de dados', {
        nomeAba: NOME_ABA,
        ultimaLinha: ultimaLinha
      });
      throw new Error('A aba precisa ter cabecalho na linha 1 e pelo menos 1 linha de dados.');
    }

    const dados = aba.getRange(1, 1, ultimaLinha, ultimaColuna).getValues();
    const col = mapearColunas(aba);

    logInfo_(ARQUIVO, FUNCAO, 'Dados e colunas carregados', {
      totalLinhas: dados.length - 1,
      totalColunas: ultimaColuna
    });

    const hoje = new Date();
    hoje.setHours(0, 0, 0, 0);

    let preparados = 0;
    let ignorados = 0;

    for (let i = 1; i < dados.length; i++) {
      if (i % 50 === 0) {
        logInfo_(ARQUIVO, FUNCAO, 'Processando lote de linhas', {
          linhaAtual: i + 1,
          totalLinhas: dados.length - 1
        });
      }

      const linha = dados[i];

      const destinatario = obterTexto(linha[col.destinatario]) || getEmailParaDefault_();
      const ccAtual = col.cc >= 0 ? obterTexto(linha[col.cc]) : '';
      const ccFinal = ccAtual || getEmailCcDefault_();

      const tituloVistoria = obterTexto(linha[col.titulo_vistoria]);
      const local = obterTexto(linha[col.local]);
      const statusEnvio = obterTexto(linha[col.status_envio]);
      const diasParaRevisao = Number(linha[col.dias_para_revisao]);
      const dataRevisao = linha[col.proxima_revisao];

      if (!destinatario || !tituloVistoria || !local || !dataRevisao || isNaN(diasParaRevisao)) {
        ignorados++;
        logWarn_(ARQUIVO, FUNCAO, 'Linha ignorada por dados obrigatorios ausentes ou invalidos', {
          linha: i + 1,
          destinatario: destinatario,
          tituloVistoria: tituloVistoria,
          local: local,
          diasParaRevisao: diasParaRevisao,
          possuiDataRevisao: !!dataRevisao
        });
        continue;
      }

      if (!(dataRevisao instanceof Date) || isNaN(dataRevisao)) {
        ignorados++;
        logWarn_(ARQUIVO, FUNCAO, 'Linha ignorada por data de revisao invalida', {
          linha: i + 1,
          valorData: String(dataRevisao)
        });
        continue;
      }

      const revisao = new Date(dataRevisao);
      revisao.setHours(0, 0, 0, 0);

      const diferencaMs = revisao - hoje;
      const diferencaDias = Math.ceil(diferencaMs / (1000 * 60 * 60 * 24));

      if (normalizar(statusEnvio) === 'enviado') {
        ignorados++;
        logInfo_(ARQUIVO, FUNCAO, 'Linha ignorada porque ja foi enviada', {
          linha: i + 1,
          destinatario: destinatario,
          statusEnvio: statusEnvio
        });
        continue;
      }

      if (diferencaDias <= diasParaRevisao && diferencaDias >= 0) {
        const assunto = '🚨 VISTORIAS PARA REVISAO: ' + tituloVistoria;
        const corpo = 'Ola, a vistoria "' + tituloVistoria + '" no local "' + local + '" esta prevista para ' + formatarData(revisao) + '.';

        aba.getRange(i + 1, col.destinatario + 1).setValue(destinatario);
        if (col.cc >= 0) aba.getRange(i + 1, col.cc + 1).setValue(ccFinal);
        aba.getRange(i + 1, col.assunto + 1).setValue(assunto);
        aba.getRange(i + 1, col.corpo + 1).setValue(corpo);
        aba.getRange(i + 1, col.status_envio + 1).setValue('Pendente');

        preparados++;

        logInfo_(ARQUIVO, FUNCAO, 'Email preparado com sucesso', {
          linha: i + 1,
          destinatario: destinatario,
          cc: ccFinal,
          tituloVistoria: tituloVistoria,
          local: local,
          dataRevisao: formatarData(revisao),
          diferencaDias: diferencaDias
        });
      } else {
        ignorados++;
        logInfo_(ARQUIVO, FUNCAO, 'Linha fora da janela de envio', {
          linha: i + 1,
          destinatario: destinatario,
          tituloVistoria: tituloVistoria,
          diferencaDias: diferencaDias,
          diasParaRevisao: diasParaRevisao
        });
      }
    }

    logInfo_(ARQUIVO, FUNCAO, 'Preparacao de emails concluida', {
      totalPreparados: preparados,
      totalIgnorados: ignorados
    });

  } catch (e) {
    logErro_(ARQUIVO, FUNCAO, 'Erro geral ao preparar emails lembrete', {
      mensagemErro: e.message,
      stack: e.stack
    });
    throw e;
  }
}

function prepararEEnviarEmails() {
  const ARQUIVO = 'Email.gs';
  const FUNCAO = 'prepararEEnviarEmails';

  logInfo_(ARQUIVO, FUNCAO, 'Iniciando fluxo preparar e enviar emails');

  try {
    prepararEmailsLembrete();
    enviarEmailsVistoria();

    logInfo_(ARQUIVO, FUNCAO, 'Fluxo preparar e enviar concluido com sucesso');
  } catch (e) {
    logErro_(ARQUIVO, FUNCAO, 'Erro no fluxo preparar e enviar emails', {
      mensagemErro: e.message,
      stack: e.stack
    });
    throw e;
  }
}

function enviarEmailsVistoria() {
  const ARQUIVO = 'Email.gs';
  const FUNCAO = 'enviarEmailsVistoria';
  const NOME_ABA = 'Emails_Enviar';
  const STATUS_ENVIADO = 'Enviado';

  logInfo_(ARQUIVO, FUNCAO, 'Iniciando envio de emails');

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const aba = ss.getSheetByName(NOME_ABA);

    if (!aba) {
      logErro_(ARQUIVO, FUNCAO, 'A aba de emails nao foi encontrada', {
        nomeAba: NOME_ABA
      });
      throw new Error('A aba "' + NOME_ABA + '" nao foi encontrada.');
    }

    const ultimaLinha = aba.getLastRow();
    const ultimaColuna = aba.getLastColumn();

    if (ultimaLinha < 2) {
      logWarn_(ARQUIVO, FUNCAO, 'A aba nao possui linhas de dados', {
        nomeAba: NOME_ABA,
        ultimaLinha: ultimaLinha
      });
      throw new Error('A aba precisa ter cabecalho na linha 1 e pelo menos 1 linha de dados.');
    }

    const dados = aba.getRange(1, 1, ultimaLinha, ultimaColuna).getValues();
    const col = mapearColunas(aba);

    logInfo_(ARQUIVO, FUNCAO, 'Dados e colunas carregados para envio', {
      totalLinhas: dados.length - 1,
      totalColunas: ultimaColuna
    });

    let enviados = 0;
    let ignorados = 0;
    let erros = 0;

    for (let i = 1; i < dados.length; i++) {
      if (i % 50 === 0) {
        logInfo_(ARQUIVO, FUNCAO, 'Processando lote de envios', {
          linhaAtual: i + 1,
          totalLinhas: dados.length - 1
        });
      }

      const linha = dados[i];

      const destinatario = obterTexto(linha[col.destinatario]) || getEmailParaDefault_();
      const cc = col.cc >= 0 ? (obterTexto(linha[col.cc]) || getEmailCcDefault_()) : getEmailCcDefault_();
      let assunto = obterTexto(linha[col.assunto]);
      let corpo = obterTexto(linha[col.corpo]);
      const diasParaRevisao = col.dias_para_revisao >= 0 ? obterTexto(linha[col.dias_para_revisao]) : '';
      const tituloVistoria = obterTexto(linha[col.titulo_vistoria]);
      const local = obterTexto(linha[col.local]);
      const proximaRevisao = formatarData(linha[col.proxima_revisao]);
      const statusEnvio = obterTexto(linha[col.status_envio]);

      if (!destinatario) {
        aba.getRange(i + 1, col.status_envio + 1).setValue('Erro: destinatario vazio');
        erros++;

        logErro_(ARQUIVO, FUNCAO, 'Linha com destinatario vazio', {
          linha: i + 1,
          tituloVistoria: tituloVistoria,
          local: local
        });
        continue;
      }

      if (normalizar(statusEnvio) !== 'pendente') {
        ignorados++;
        logWarn_(ARQUIVO, FUNCAO, 'Linha ignorada no envio por status diferente de Pendente', {
          linha: i + 1,
          statusEnvio: statusEnvio,
          destinatario: destinatario,
          assunto: assunto
        });
        continue;
      }

      const variaveis = {
        destinatario: destinatario,
        cc: cc,
        assunto: assunto,
        corpo: corpo,
        dias_para_revisao: diasParaRevisao,
        titulo_vistoria: tituloVistoria,
        local: local,
        proxima_revisao: proximaRevisao
      };

      assunto = substituirVariaveis(assunto, variaveis);
      corpo = substituirVariaveis(corpo, variaveis);

      try {
        const htmlBody = montarHtmlEmailVistoria_({
          tituloVistoria: tituloVistoria,
          local: local,
          dataRevisao: proximaRevisao,
          destinatario: destinatario,
          diasParaRevisao: diasParaRevisao
        });

        const opcoes = {
          htmlBody: htmlBody
        };

        if (cc) opcoes.cc = cc;

        GmailApp.sendEmail(destinatario, assunto, corpo, opcoes);
        aba.getRange(i + 1, col.status_envio + 1).setValue(STATUS_ENVIADO);

        enviados++;

        logInfo_(ARQUIVO, FUNCAO, 'Email enviado com sucesso', {
          linha: i + 1,
          destinatario: destinatario,
          cc: cc,
          assunto: assunto
        });
      } catch (e) {
        aba.getRange(i + 1, col.status_envio + 1).setValue('Erro: ' + e.message);
        erros++;

        logErro_(ARQUIVO, FUNCAO, 'Erro ao enviar email', {
          linha: i + 1,
          destinatario: destinatario,
          cc: cc,
          assunto: assunto,
          mensagemErro: e.message,
          stack: e.stack
        });
      }
    }

    logInfo_(ARQUIVO, FUNCAO, 'Envio de emails concluido', {
      totalEnviados: enviados,
      totalIgnorados: ignorados,
      totalErros: erros
    });

  } catch (e) {
    logErro_(ARQUIVO, FUNCAO, 'Erro geral ao enviar emails', {
      mensagemErro: e.message,
      stack: e.stack
    });
    throw e;
  }
}

/* =========================
   TEMPLATE HTML
========================= */

function montarHtmlEmailVistoria_(dados) {
  return `
  <div style="font-family: Arial, sans-serif; background:#f4f6f8; padding:20px;">
    <div style="max-width:800px; margin:auto; background:#ffffff; border-radius:10px; overflow:hidden;">
      <div style="padding:20px; text-align:center;">
        <h2 style="color:#d32f2f; margin:0;">📋 VISTORIAS PARA REVISAO</h2>
      </div>

      <div style="padding:0 24px 20px 24px; color:#333;">
        <p><strong>Data do relatorio:</strong> ${formatarDataHoraEmail_(new Date())}</p>
        <p><strong>Total de vistorias:</strong> 1</p>
      </div>

      <div style="padding:0 24px;">
        <h3 style="color:#444;">📊 RESUMO POR CRITICIDADE:</h3>

        <table width="100%" style="border-collapse:collapse; margin-top:10px;">
          <tr style="background:#e0e0e0;">
            <th style="padding:10px; text-align:left; border:1px solid #ccc;">Criticidade</th>
            <th style="padding:10px; border:1px solid #ccc;">Quantidade</th>
          </tr>
          <tr>
            <td style="padding:10px; color:#d32f2f; border:1px solid #ccc;"><strong>Alta</strong></td>
            <td style="text-align:center; border:1px solid #ccc;">1</td>
          </tr>
          <tr>
            <td style="padding:10px; color:#f9a825; border:1px solid #ccc;"><strong>Media</strong></td>
            <td style="text-align:center; border:1px solid #ccc;">0</td>
          </tr>
          <tr>
            <td style="padding:10px; color:#2e7d32; border:1px solid #ccc;"><strong>Baixa</strong></td>
            <td style="text-align:center; border:1px solid #ccc;">0</td>
          </tr>
        </table>
      </div>

      <div style="padding:20px 24px;">
        <h3 style="color:#444;">⚠️ DETALHAMENTO DAS VISTORIAS:</h3>

        <table width="100%" style="border-collapse:collapse; margin-top:10px; font-size:13px;">
          <tr style="background:#333; color:#fff;">
            <th style="padding:8px; border:1px solid #555;">Vistoria</th>
            <th style="padding:8px; border:1px solid #555;">Local</th>
            <th style="padding:8px; border:1px solid #555;">Vencimento</th>
            <th style="padding:8px; border:1px solid #555;">Prazo</th>
          </tr>
          <tr style="background:#f9f9f9;">
            <td style="padding:8px; border:1px solid #ccc;">${escapeHtml_(dados.tituloVistoria)}</td>
            <td style="padding:8px; border:1px solid #ccc;">${escapeHtml_(dados.local)}</td>
            <td style="padding:8px; border:1px solid #ccc;">${escapeHtml_(dados.dataRevisao)}</td>
            <td style="padding:8px; border:1px solid #ccc;">${escapeHtml_(String(dados.diasParaRevisao || ''))} dias</td>
          </tr>
        </table>
      </div>

      <div style="background:#fff3cd; border-left:5px solid #f9a825; margin:20px; padding:15px;">
        <h4 style="margin-top:0;">⚠️ ACAO REQUERIDA:</h4>
        <ul style="margin:0; padding-left:20px;">
          <li>📞 Contatar responsavel para revisao dos itens nao conformes</li>
          <li>📅 Agendar nova vistoria para pendencias criticas</li>
          <li>📝 Documentar acoes corretivas implementadas</li>
          <li>✅ Atualizar status no sistema apos as correcoes</li>
        </ul>
      </div>

      <div style="text-align:center; font-size:12px; color:#777; padding:15px;">
        Relatorio automatico do sistema Vistorias Produttivo
      </div>
    </div>
  </div>
  `;
}

/* =========================
   TESTES
========================= */

function testeEnvioDireto() {
  const ARQUIVO = 'Email.gs';
  const FUNCAO = 'testeEnvioDireto';

  try {
    const destino = 'lucas.leite@charth.com.br';

    GmailApp.sendEmail(
      destino,
      'Teste direto Apps Script',
      'Se voce recebeu este email, o GmailApp esta funcionando.'
    );

    logInfo_(ARQUIVO, FUNCAO, 'Teste direto enviado com sucesso', {
      destinatario: destino
    });
  } catch (e) {
    logErro_(ARQUIVO, FUNCAO, 'Erro no teste direto de email', {
      mensagemErro: e.message,
      stack: e.stack
    });
    throw e;
  }
}

function forcarEnvioLinhaTeste_(numeroLinha) {
  const ARQUIVO = 'Email.gs';
  const FUNCAO = 'forcarEnvioLinhaTeste_';
  const NOME_ABA = 'Emails_Enviar';

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const aba = ss.getSheetByName(NOME_ABA);
    const col = mapearColunas(aba);

    aba.getRange(numeroLinha, col.status_envio + 1).setValue('Pendente');

    logInfo_(ARQUIVO, FUNCAO, 'Linha marcada como Pendente para teste', {
      linha: numeroLinha
    });

    enviarEmailsVistoria();
  } catch (e) {
    logErro_(ARQUIVO, FUNCAO, 'Erro ao forcar envio de linha', {
      linha: numeroLinha,
      mensagemErro: e.message,
      stack: e.stack
    });
    throw e;
  }
}

/* =========================
   AUXILIARES
========================= */

function mapearColunas(aba) {
  if (!aba) {
    throw new Error('Nao execute mapearColunas diretamente. Use o menu "Emails" na planilha.');
  }

  const cabecalho = aba.getRange(1, 1, 1, aba.getLastColumn()).getDisplayValues()[0];

  const col = {
    destinatario: cabecalho.indexOf('destinatario'),
    cc: cabecalho.indexOf('cc'),
    assunto: cabecalho.indexOf('assunto'),
    corpo: cabecalho.indexOf('corpo'),
    dias_para_revisao: cabecalho.indexOf('dias_para_revisao'),
    titulo_vistoria: cabecalho.indexOf('titulo_vistoria'),
    local: cabecalho.indexOf('local'),
    proxima_revisao: cabecalho.indexOf('proxima_revisao'),
    status_envio: cabecalho.indexOf('status_envio')
  };

  const obrigatorias = [
    'destinatario',
    'assunto',
    'corpo',
    'titulo_vistoria',
    'local',
    'proxima_revisao',
    'status_envio'
  ];

  const faltando = obrigatorias.filter(function(nome) {
    return col[nome] < 0;
  });

  if (faltando.length > 0) {
    throw new Error('Colunas obrigatorias nao encontradas: ' + faltando.join(', '));
  }

  return col;
}

function substituirVariaveis(textoBase, variaveis) {
  if (!textoBase) return '';

  return textoBase
    .replace(/\{\{destinatario\}\}/gi, variaveis.destinatario || '')
    .replace(/\{\{cc\}\}/gi, variaveis.cc || '')
    .replace(/\{\{assunto\}\}/gi, variaveis.assunto || '')
    .replace(/\{\{corpo\}\}/gi, variaveis.corpo || '')
    .replace(/\{\{dias_para_revisao\}\}/gi, variaveis.dias_para_revisao || '')
    .replace(/\{\{titulo_vistoria\}\}/gi, variaveis.titulo_vistoria || '')
    .replace(/\{\{local\}\}/gi, variaveis.local || '')
    .replace(/\{\{proxima_revisao\}\}/gi, variaveis.proxima_revisao || '');
}

function obterTexto(valor) {
  if (valor === null || valor === undefined) return '';
  return String(valor).trim();
}

function formatarData(valor) {
  if (!valor) return '';

  if (Object.prototype.toString.call(valor) === '[object Date]' && !isNaN(valor)) {
    return Utilities.formatDate(valor, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  }

  return String(valor).trim();
}

function formatarDataHoraEmail_(valor) {
  return Utilities.formatDate(valor, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
}

function normalizar(valor) {
  return obterTexto(valor).toLowerCase().trim();
}

function escapeHtml_(texto) {
  return String(texto || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}