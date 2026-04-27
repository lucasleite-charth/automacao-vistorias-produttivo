function importarRelatorios() {
  const ARQUIVO = 'Importador.gs';
  const FUNCAO = 'importarRelatorios';

  logInfo_('Importador.gs', 'importarRelatorios', 'Iniciando importacao de relatorios');

  try {
    const pastaId = '1WAwOZV-msPVqojhcHkVFDtm8NrT6biEF'.trim();
    logInfo_(ARQUIVO, FUNCAO, 'Buscando pasta no Drive', { pastaId: pastaId });

    const pasta = DriveApp.getFolderById(pastaId);
    const arquivos = pasta.getFiles();

    let importados = 0;
    let ignorados = 0;

    while (arquivos.hasNext()) {
      const arquivo = arquivos.next();
      const nomeOriginal = arquivo.getName();
      const nome = nomeOriginal.toLowerCase();

      logInfo_(ARQUIVO, FUNCAO, 'Processando arquivo', {
        nome: nomeOriginal,
        id: arquivo.getId(),
        mimeType: arquivo.getMimeType()
      });

      if (nome.includes('atividades')) {
        importarArquivoParaAba_(arquivo, 'Base_Atividades');
        importados++;
        continue;
      }

      if (nome.includes('relatorio_consolidado') || nome.includes('consolidado')) {
        importarArquivoParaAba_(arquivo, 'Base_Formularios');
        importados++;
        continue;
      }

      ignorados++;
      logWarn_(ARQUIVO, FUNCAO, 'Arquivo ignorado por nao corresponder ao filtro', {
        nome: nomeOriginal
      });
    }

    logInfo_(ARQUIVO, FUNCAO, 'Importacao concluida', {
      totalImportados: importados,
      totalIgnorados: ignorados
    });
  } catch (e) {
    logErro_(ARQUIVO, FUNCAO, 'Erro geral na importacao de relatorios', {
      mensagemErro: e.message,
      stack: e.stack
    });
    throw e;
  }
}

function importarArquivoParaAba_(arquivo, nomeAbaDestino) {
  const ARQUIVO = 'Importador.gs';
  const FUNCAO = 'importarArquivoParaAba_';

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaDestino = ss.getSheetByName(nomeAbaDestino);

  if (!abaDestino) {
    logErro_(ARQUIVO, FUNCAO, 'A aba de destino nao foi encontrada', {
      nomeAbaDestino: nomeAbaDestino
    });
    throw new Error('A aba "' + nomeAbaDestino + '" nao foi encontrada.');
  }

  const blob = arquivo.getBlob();

  logInfo_(ARQUIVO, FUNCAO, 'Preparando conversao do arquivo', {
    nome: arquivo.getName(),
    mimeTypeOriginal: arquivo.getMimeType(),
    destino: nomeAbaDestino
  });

  const convertido = Drive.Files.create(
    {
      name: arquivo.getName(),
      mimeType: MimeType.GOOGLE_SHEETS
    },
    blob
  );

  try {
    logInfo_(ARQUIVO, FUNCAO, 'Arquivo convertido temporariamente', {
      nome: arquivo.getName(),
      idConvertido: convertido.id
    });

    const planilhaConvertida = SpreadsheetApp.openById(convertido.id);
    const abaOrigem = planilhaConvertida.getSheets()[0];
    const dados = abaOrigem.getDataRange().getValues();

    if (!dados || dados.length === 0 || dados[0].length === 0) {
      logWarn_(ARQUIVO, FUNCAO, 'Arquivo sem dados para importacao', {
        nome: arquivo.getName(),
        destino: nomeAbaDestino
      });
      return;
    }

    abaDestino.clearContents();
    abaDestino.getRange(1, 1, dados.length, dados[0].length).setValues(dados);
    SpreadsheetApp.flush();

    logInfo_(ARQUIVO, FUNCAO, 'Arquivo importado com sucesso', {
      nome: arquivo.getName(),
      destino: nomeAbaDestino,
      linhas: dados.length,
      colunas: dados[0].length
    });
  } catch (e) {
    logErro_(ARQUIVO, FUNCAO, 'Erro ao importar arquivo', {
      nome: arquivo.getName(),
      destino: nomeAbaDestino,
      mensagemErro: e.message,
      stack: e.stack
    });
    throw e;
  } finally {
    try {
      DriveApp.getFileById(convertido.id).setTrashed(true);
      logInfo_(ARQUIVO, FUNCAO, 'Arquivo temporario removido', {
        nome: arquivo.getName(),
        idConvertido: convertido.id
      });
    } catch (erroRemocao) {
      logWarn_(ARQUIVO, FUNCAO, 'Falha ao remover arquivo temporario', {
        nome: arquivo.getName(),
        idConvertido: convertido.id,
        mensagemErro: erroRemocao.message
      });
    }
  }
}

/* =========================
   LOGS
========================= */

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
    arquivo,
    funcao,
    mensagem,
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