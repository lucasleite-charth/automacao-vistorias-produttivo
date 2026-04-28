function limparLogsSeExceder_(limite) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName('Logs');

  if (!aba) return;

  const ultimaLinha = aba.getLastRow();

  if (ultimaLinha <= limite) return;

  const excesso = ultimaLinha - limite;

  // mantém cabeçalho (linha 1)
  aba.deleteRows(2, excesso);
}

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
  const dataHora = new Date();

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
    dataHora,
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
  logPlanilha_('ERRO', arquivo, funcao, mensagem, detalhes);
}

