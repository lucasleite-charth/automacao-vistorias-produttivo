function getSheet_(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

function limparAba_(sh) {
  sh.clearContents();
}

function setHeaderAndData_(sh, header, data) {
  limparAba_(sh);
  if (header && header.length) {
    sh.getRange(1, 1, 1, header.length).setValues([header]);
  }
  if (data && data.length) {
    sh.getRange(2, 1, data.length, data[0].length).setValues(data);
  }
}

function parseNumber_(val) {
  if (val === null || val === '' || typeof val === 'undefined') return 0;
  var str = String(val).replace('.', '').replace(',', '.');
  var n = parseFloat(str);
  return isNaN(n) ? 0 : n;
}

function parseDate_(val) {
  if (Object.prototype.toString.call(val) === '[object Date]') return val;
  if (val === null || val === '') return null;
  var tryDate = new Date(val);
  if (!isNaN(tryDate.getTime())) return tryDate;
  return null;
}

function formatDate_(date) {
  if (!date) return '';
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd/MM/yyyy');
}

function gerarHashEvento_(obj) {
  const payload = JSON.stringify({
    idAtividade: obj.idAtividade,
    titulo: obj.titulo,
    dataFim: formatarIso_(obj.dataFim),
    criticidade: obj.criticidade,
    dataRevisao: formatarIso_(obj.dataRevisao),
    percentual: Number(obj.percentual) || 0
  });

  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.MD5,
    payload
  );

  return digest.map(function(b) {
    const v = (b < 0 ? b + 256 : b).toString(16);
    return v.length === 1 ? '0' + v : v;
  }).join('');
}

function getOrCreateColumn_(sheet, cabecalho, nomeColuna) {
  let idx = cabecalho.indexOf(nomeColuna);
  if (idx !== -1) return idx;

  idx = cabecalho.length;
  sheet.getRange(1, idx + 1).setValue(nomeColuna);
  return idx;
}

function normalizarData_(valor) {
  if (!valor) return null;

  const data = valor instanceof Date ? valor : new Date(valor);
  if (isNaN(data.getTime())) return null;

  return data;
}

function normalizarTexto_(texto) {
  return String(texto || '')
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');
}

function formatarIso_(data) {
  const d = normalizarData_(data);
  return d ? d.toISOString() : '';
}

function formatarData_(data) {
  const d = normalizarData_(data);
  if (!d) return 'N/D';
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
}

function existeHashEventoNaPlanilha_(hash) {
  var sh = getSheet_('Vistorias_Tratadas');
  var values = sh.getRange(2, 19, sh.getLastRow() - 1, 1).getValues(); // col 19 = hash_evento
  for (var i = 0; i < values.length; i++) {
    if (values[i][0] === hash) return true;
  }
  return false;
}
function getColIndex_(cabecalho, nomeColuna) {
  const idx = cabecalho.indexOf(nomeColuna);
  if (idx === -1) {
    throw new Error('Coluna obrigatória não encontrada: ' + nomeColuna);
  }
  return idx;
}

function getOrCreateColumn_(sheet, cabecalho, nomeColuna) {
  let idx = cabecalho.indexOf(nomeColuna);
  if (idx !== -1) return idx;

  idx = cabecalho.length;
  sheet.getRange(1, idx + 1).setValue(nomeColuna);
  return idx;
}

function normalizarData_(valor) {
  if (!valor) return null;

  const data = valor instanceof Date ? valor : new Date(valor);
  if (isNaN(data.getTime())) return null;

  return data;
}

function normalizarTexto_(texto) {
  return String(texto || '')
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');
}

function formatarIso_(data) {
  const d = normalizarData_(data);
  return d ? d.toISOString() : '';
}