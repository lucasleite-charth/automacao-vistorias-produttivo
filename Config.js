var CONFIG_CACHE = null;

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Emails')
    .addItem('Preparar lembretes', 'prepararEmailsLembrete')
    .addItem('Enviar emails', 'enviarEmailsVistoria')
    .addItem('Preparar e enviar', 'prepararEEnviarEmails')
    .addToUi();
}

function prepararEEnviarEmails() {
  limparConfigCache_();
  prepararEmailsLembrete();
  enviarEmailsVistoria();
}

/* =========================
   CONFIG
========================= */

function getConfig_() {
  if (CONFIG_CACHE) return CONFIG_CACHE;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Config');

  if (!sh) {
    throw new Error('A aba "Config" nao foi encontrada.');
  }

  const ultimaLinha = sh.getLastRow();
  const ultimaColuna = sh.getLastColumn();

  if (ultimaLinha < 2 || ultimaColuna < 2) {
    throw new Error('A aba "Config" precisa ter cabecalho e pelo menos uma linha com chave e valor.');
  }

  const values = sh.getRange(1, 1, ultimaLinha, Math.max(2, ultimaColuna)).getValues();
  const cfg = {};

  for (let i = 1; i < values.length; i++) {
    const chave = normalizarChaveConfig_(values[i][0]);
    const valor = values[i][1];

    if (chave) {
      cfg[chave] = valor;
    }
  }

  CONFIG_CACHE = cfg;
  return cfg;
}

function limparConfigCache_() {
  CONFIG_CACHE = null;
}

function getPastaDriveId_() {
  return getConfigTexto_('pastadriveid', '');
}

function getCalendarId_() {
  return getConfigTexto_('calendarid', 'primary');
}

function getEmailParaDefault_() {
  return getConfigTexto_('email_para', '');
}

function getEmailCcDefault_() {
  return getConfigTexto_('email_cc', '');
}

function getDiasLembrete_() {
  const valor = getConfigTexto_('dias_lembrete', '');
  const padrao = [15, 7, 1, 0];

  if (!valor) return padrao;

  const lista = valor
    .split(',')
    .map(function(s) {
      return parseInt(String(s).trim(), 10);
    })
    .filter(function(n) {
      return !isNaN(n);
    });

  return lista.length ? lista : padrao;
}

function getPrazoRevisaoPorCriticidade_(criticidade) {
  const c = normalizarTexto_(criticidade);

  if (c === 'alta') {
    return getConfigNumero_('dias_revisao_alta', 15);
  }

  if (c === 'media') {
    return getConfigNumero_('dias_revisao_media', 30);
  }

  return getConfigNumero_('dias_revisao_baixa', 60);
}

function getConfigTexto_(chave, valorPadrao) {
  const cfg = getConfig_();
  const valor = cfg[normalizarChaveConfig_(chave)];

  if (valor === null || valor === undefined || String(valor).trim() === '') {
    return valorPadrao;
  }

  return String(valor).trim();
}

function getConfigNumero_(chave, valorPadrao) {
  const texto = getConfigTexto_(chave, '');
  const numero = parseInt(texto, 10);
  return isNaN(numero) ? valorPadrao : numero;
}

function normalizarChaveConfig_(valor) {
  return String(valor || '')
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, '')
    .replace(/[^\w]/g, '');
}

function normalizarTexto_(valor) {
  return String(valor || '')
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');
}
