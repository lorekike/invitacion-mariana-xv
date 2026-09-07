const SHEET_NAME = 'Respuestas de formulario 1';
const INVITATION_BASE_URL = 'https://xvmarianarojas.com/';

const HEADERS = {
  name: 'Nombre completo del invitado',
  phone: 'Celular del invitado',
  allowed: 'Cantidad de acompañantes',
  token: 'ID Invitación',
  link: 'Enlace de invitación',
  confirmed: 'Confirmado',
  attendance: 'Asistencia confirmada',
  companionCount: 'Acompañantes confirmados',
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
    if (!token || !attendance || !Number.isInteger(companionCount)) {
      return json_({ok: false, error: 'Datos de confirmación incompletos'});
    }

    const record = findInvitation_(token);
    if (!record) return json_({ok: false, error: 'Invitación no encontrada'});
    if (record.confirmed) return json_({ok: false, duplicate: true, error: 'Esta invitación ya fue confirmada'});
    if (companionCount < 0 || companionCount > record.maxCompanions) {
      return json_({ok: false, error: 'La cantidad supera el cupo autorizado'});
    }

    const count = attendance === 'Sí' ? companionCount : 0;
    const sheet = getSheet_();
    const columns = ensureColumns_(sheet);
    sheet.getRange(record.row, columns.confirmed).setValue('Sí');
    sheet.getRange(record.row, columns.attendance).setValue(attendance);
    sheet.getRange(record.row, columns.companionCount).setValue(count);
    sheet.getRange(record.row, columns.confirmedAt).setValue(new Date());
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
  SpreadsheetApp.flush();
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
