const SHEET_NAME = 'Respuestas de formulario 1';
const SUMMARY_SHEET_NAME = 'Resumen';
const INVITATION_BASE_URL = 'https://xvmarianarojas.com/';
const MAX_TOTAL_INVITADOS = 180;

const HEADERS = {
  name: 'Nombre completo del invitado',
  invitedBy: 'Invitado por',
  phone: 'Celular del invitado',
  allowed: 'Cantidad de acompañantes',
  token: 'ID Invitación',
  link: 'Enlace de invitación',
  confirmed: 'Confirmado',
  attendance: 'Asistencia confirmada',
  companionCount: 'Acompañantes confirmados',
  companionNames: 'Nombres de acompañantes',
  confirmedAt: 'Fecha de confirmación'
};

function doGet(e) {
  try {
    const token = cleanToken_(e.parameter.i);
    if (!token) return json_({ok: false, error: 'Código de invitación requerido'});

    const record = findInvitation_(token);
    if (!record) return json_({ok: false, error: 'Invitación no encontrada'});

    return json_({
      ok: true,
      name: record.name,
      maxCompanions: record.maxCompanions,
      confirmed: record.confirmed,
      attendance: record.attendance,
      companionCount: record.companionCount
    });
  } catch (error) {
    return json_({ok: false, error: 'No fue posible consultar la invitación'});
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const token = cleanToken_(e.parameter.i);
    const attendance = e.parameter.attendance === 'Sí' ? 'Sí' : e.parameter.attendance === 'No' ? 'No' : '';
    const companionCount = Number(e.parameter.companionCount || 0);
    let companionNames = [];
    try { companionNames = JSON.parse(e.parameter.companionNames || '[]'); } catch (ignore) {}
    companionNames = Array.isArray(companionNames) ? companionNames.map(function(name) { return String(name || '').trim(); }) : [];
    if (!token || !attendance || !Number.isInteger(companionCount)) {
      return json_({ok: false, error: 'Datos de confirmación incompletos'});
    }

    const record = findInvitation_(token);
    if (!record) return json_({ok: false, error: 'Invitación no encontrada'});
    if (record.confirmed) return json_({ok: false, duplicate: true, error: 'Esta invitación ya fue confirmada'});
    if (companionCount < 0 || companionCount > record.maxCompanions) {
      return json_({ok: false, error: 'La cantidad supera el cupo autorizado'});
    }
    if (attendance === 'Sí' && (companionNames.length !== companionCount || companionNames.some(function(name) { return !name; }))) {
      return json_({ok: false, error: 'Debes registrar el nombre de cada acompañante'});
    }

    const count = attendance === 'Sí' ? companionCount : 0;
    const sheet = getSheet_();
    const columns = ensureColumns_(sheet);
    sheet.getRange(record.row, columns.confirmed).setValue('Sí');
    sheet.getRange(record.row, columns.attendance).setValue(attendance);
    sheet.getRange(record.row, columns.companionCount).setValue(count);
    sheet.getRange(record.row, columns.companionNames).setValue(attendance === 'Sí' ? companionNames.join('\n') : '');
    sheet.getRange(record.row, columns.confirmedAt).setValue(new Date());
    actualizarResumen_();
    SpreadsheetApp.flush();

    return json_({ok: true, name: record.name, attendance: attendance, companionCount: count});
  } catch (error) {
    return json_({ok: false, error: 'No fue posible guardar la confirmación'});
  } finally {
    lock.releaseLock();
  }
}

function prepararInvitaciones() {
  const sheet = getSheet_();
  const columns = ensureColumns_(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const values = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  values.forEach(function(row, index) {
    const sheetRow = index + 2;
    const name = String(row[columns.name - 1] || '').trim();
    if (!name) return;
    let token = cleanToken_(row[columns.token - 1]);
    if (!token) token = Utilities.getUuid().replace(/-/g, '').slice(0, 24);
    sheet.getRange(sheetRow, columns.token).setValue(token);
    sheet.getRange(sheetRow, columns.link).setValue(INVITATION_BASE_URL + '?i=' + token);
  });
  sheet.autoResizeColumns(1, sheet.getLastColumn());
  actualizarResumen_();
  SpreadsheetApp.flush();
}

function actualizarResumen() {
  actualizarResumen_();
}

function onEdit(e) {
  if (!e || !e.range || e.range.getSheet().getName() !== SHEET_NAME) return;
  actualizarResumen_();
  actualizarCupoFormulario_();
}

function onFormSubmit(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    prepararInvitaciones();
    actualizarCupoFormulario_();
  } finally {
    lock.releaseLock();
  }
}

function actualizarResumen_() {
  const source = getSheet_();
  const columns = ensureColumns_(source);
  const lastRow = source.getLastRow();
  const rows = lastRow < 2 ? [] : source.getRange(2, 1, lastRow - 1, source.getLastColumn()).getValues();

  let registered = 0;
  let authorizedCompanions = 0;
  let responses = 0;
  let attending = 0;
  let notAttending = 0;
  let companions = 0;
  const groups = {};

  rows.forEach(function(row) {
    if (!String(row[columns.name - 1] || '').trim()) return;
    registered++;
    const allowed = Math.max(0, Math.min(5, Number(row[columns.allowed - 1]) || 0));
    authorizedCompanions += allowed;
    const invitedBy = String(row[columns.invitedBy - 1] || '').trim() || 'Sin asignar';
    if (!groups[invitedBy]) groups[invitedBy] = {guests: 0, companions: 0};
    groups[invitedBy].guests++;
    groups[invitedBy].companions += allowed;
    const confirmed = String(row[columns.confirmed - 1] || '').trim().toLowerCase() === 'sí';
    const attendance = String(row[columns.attendance - 1] || '').trim();
    if (confirmed) responses++;
    if (attendance === 'Sí') {
      attending++;
      companions += Math.max(0, Number(row[columns.companionCount - 1]) || 0);
    } else if (attendance === 'No') {
      notAttending++;
    }
  });

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let summary = spreadsheet.getSheetByName(SUMMARY_SHEET_NAME);
  if (!summary) summary = spreadsheet.insertSheet(SUMMARY_SHEET_NAME, 0);
  summary.clear();
  summary.getRange('A1:B1').merge().setValue('Resumen de confirmaciones');
  summary.getRange('A3:B12').setValues([
    ['Indicador', 'Cantidad'],
    ['Invitados principales registrados', registered],
    ['Acompañantes autorizados', authorizedCompanions],
    ['Total de personas invitadas', registered + authorizedCompanions],
    ['Confirmaciones recibidas', responses],
    ['Invitados que asistirán', attending],
    ['Acompañantes confirmados', companions],
    ['Total de personas que asistirán', attending + companions],
    ['Invitados que no asistirán', notAttending],
    ['Invitados sin responder', Math.max(0, registered - responses)]
  ]);
  summary.getRange('A1:B1').setBackground('#07162f').setFontColor('#f4d98b').setFontWeight('bold').setFontSize(16).setHorizontalAlignment('center');
  summary.getRange('A3:B3').setBackground('#173f75').setFontColor('#ffffff').setFontWeight('bold');
  summary.getRange('A4:A12').setFontWeight('bold');
  summary.getRange('B4:B12').setNumberFormat('0').setHorizontalAlignment('center');
  summary.getRange('A6:B6').setBackground('#e8eef7').setFontColor('#07162f').setFontWeight('bold');
  summary.getRange('A10:B10').setBackground('#f4d98b').setFontColor('#07162f').setFontWeight('bold');
  summary.getRange('A12:B12').setBackground('#e8eef7');
  const groupRows = Object.keys(groups).sort().map(function(invitedBy) {
    return [invitedBy, groups[invitedBy].guests, groups[invitedBy].companions, groups[invitedBy].guests + groups[invitedBy].companions];
  });
  summary.getRange('D1:G1').merge().setValue('Invitados por familiar');
  summary.getRange('D1:G1').setBackground('#07162f').setFontColor('#f4d98b').setFontWeight('bold').setFontSize(16).setHorizontalAlignment('center');
  summary.getRange('D3:G3').setValues([['Invitado por', 'Invitados principales', 'Acompañantes autorizados', 'Total de personas invitadas']]);
  summary.getRange('D3:G3').setBackground('#173f75').setFontColor('#ffffff').setFontWeight('bold').setWrap(true);
  if (groupRows.length) {
    summary.getRange(4, 4, groupRows.length, 4).setValues(groupRows);
    summary.getRange(4, 5, groupRows.length, 3).setNumberFormat('0').setHorizontalAlignment('center');
  }
  summary.setColumnWidth(1, 260);
  summary.setColumnWidth(2, 110);
  summary.setColumnWidth(4, 180);
  summary.setColumnWidth(5, 130);
  summary.setColumnWidth(6, 150);
  summary.setColumnWidth(7, 150);
  summary.setFrozenRows(3);
}

function configurarCampoInvitadoPor() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const formUrl = spreadsheet.getFormUrl();
  if (!formUrl) throw new Error('La hoja no tiene un formulario vinculado');
  const form = FormApp.openByUrl(formUrl);
  const items = form.getItems();
  const current = items.find(function(item) {
    return item.getTitle().trim().toLowerCase() === 'invitado por';
  });
  const duplicate = items.find(function(item) {
    return item.getTitle().trim().toLowerCase() === 'invitado de';
  });
  if (current && duplicate) {
    form.deleteItem(duplicate);
  } else if (duplicate) {
    duplicate.asTextItem().setTitle(HEADERS.invitedBy).setRequired(true);
  } else if (!current) {
    form.addTextItem().setTitle(HEADERS.invitedBy).setRequired(true);
  }
}

function configurarControlCupo() {
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'onFormSubmit') ScriptApp.deleteTrigger(trigger);
  });
  ScriptApp.newTrigger('onFormSubmit')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onFormSubmit()
    .create();
  actualizarCupoFormulario_();
}

function actualizarCupoFormulario_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const formUrl = spreadsheet.getFormUrl();
  if (!formUrl) throw new Error('La hoja no tiene un formulario vinculado');

  const total = calcularTotalInvitados_();
  const disponibles = Math.max(0, MAX_TOTAL_INVITADOS - total);
  const form = FormApp.openByUrl(formUrl);

  if (disponibles === 0) {
    form.setCustomClosedFormMessage('Hemos completado el cupo de 180 personas. Gracias.');
    form.setAcceptingResponses(false);
    return;
  }

  form.setAcceptingResponses(true);
  const maxAcompanantes = Math.min(5, Math.max(0, disponibles - 1));
  const item = form.getItems().find(function(formItem) {
    return formItem.getTitle().trim().toLowerCase() === HEADERS.allowed.toLowerCase();
  });
  if (!item) throw new Error('No se encontró la pregunta de cantidad de acompañantes');

  const helpText = 'Cupo disponible: ' + disponibles + ' persona(s). Máximo ' + maxAcompanantes + ' acompañante(s) para este registro.';
  const options = [];
  for (let value = 0; value <= maxAcompanantes; value++) options.push(String(value));
  if (item.getType() === FormApp.ItemType.MULTIPLE_CHOICE) {
    item.asMultipleChoiceItem().setChoiceValues(options).setHelpText(helpText).setRequired(true);
  } else if (item.getType() === FormApp.ItemType.LIST) {
    item.asListItem().setChoiceValues(options).setHelpText(helpText).setRequired(true);
  } else if (item.getType() === FormApp.ItemType.TEXT) {
    const validationBuilder = FormApp.createTextValidation().setHelpText(helpText);
    if (maxAcompanantes === 0) {
      validationBuilder.requireNumberEqualTo(0);
    } else {
      validationBuilder.requireNumberBetween(0, maxAcompanantes);
    }
    item.asTextItem().setValidation(validationBuilder.build()).setHelpText(helpText).setRequired(true);
  } else {
    throw new Error('El tipo de pregunta de acompañantes no es compatible');
  }
}

function calcularTotalInvitados_() {
  const sheet = getSheet_();
  const columns = ensureColumns_(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  const rows = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  return rows.reduce(function(total, row) {
    if (!String(row[columns.name - 1] || '').trim()) return total;
    const allowed = Math.max(0, Math.min(5, Number(row[columns.allowed - 1]) || 0));
    return total + 1 + allowed;
  }, 0);
}

function findInvitation_(token) {
  const sheet = getSheet_();
  const columns = ensureColumns_(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  const values = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  for (let index = 0; index < values.length; index++) {
    const row = values[index];
    if (cleanToken_(row[columns.token - 1]) !== token) continue;
    return {
      row: index + 2,
      name: String(row[columns.name - 1] || '').trim(),
      maxCompanions: Math.max(0, Math.min(5, Number(row[columns.allowed - 1]) || 0)),
      confirmed: String(row[columns.confirmed - 1] || '').trim().toLowerCase() === 'sí',
      attendance: String(row[columns.attendance - 1] || '').trim(),
      companionCount: Number(row[columns.companionCount - 1]) || 0
    };
  }
  return null;
}

function ensureColumns_(sheet) {
  const required = Object.values(HEADERS);
  const current = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0];
  required.forEach(function(header) {
    if (current.indexOf(header) === -1) {
      current.push(header);
      sheet.getRange(1, current.length).setValue(header);
    }
  });
  const columns = {};
  Object.keys(HEADERS).forEach(function(key) {
    columns[key] = current.indexOf(HEADERS[key]) + 1;
  });
  return columns;
}

function getSheet_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('No se encontró la hoja de invitados');
  return sheet;
}

function cleanToken_(value) {
  const token = String(value || '').trim();
  return /^[a-f0-9]{20,40}$/i.test(token) ? token : '';
}

function json_(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload)).setMimeType(ContentService.MimeType.JSON);
}
