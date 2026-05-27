/**
 * RACS Minera Casma — Apps Script v3
 * Fix: firmaRegistro y firmaEdicion nunca se guardan como dataURL
 *      (evita que Sheets trunce la fila y desplace Mes/Año)
 */

const CARPETA_DRIVE_ID = '1OuP4ySfsrKpwhxgVxsAR6tdGi2fJYbNs';
const SHEET_ID_MINA    = '141L31UcPGvrKEPGD9bZjczUpftNByUrYs315-KKJJww';
const HOJA_REPORTES    = 'Reportes';
const HOJA_PERSONAL    = 'Personal';
const HOJA_RESPONSABLES= 'Responsables';

const HEADERS_REPORTES = [
  'ID','CreadoEn','Estado','Tipo','NivelRiesgo','Categoría','CausaProbable','Descripción',
  'Responsable','Ubicación','PersonaObservada','DniObservado',
  'ÁreaObservado','CargoObservado','Fecha','Hora',
  'Reportador','DniReportador','ÁreaReportador','CargoReportador',
  'FotosHallazgo','FotosLevantamiento','MedidasAcciones','FechaCierre','ActualizadoEn',
  'FirmaRegistro','FirmaEdicion','Mes','Año'
];
const HEADERS_PERSONAL = ['DNI','Nombre','Área','Cargo','AgregadoEn'];

// ──────────── ENTRY POINTS ────────────
function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || '';
    if (action === 'personal')     return jsonOk({ personal: leerPersonal() });
    if (action === 'reportes')     return jsonOk({ reportes: leerReportes() });
    if (action === 'responsables') return jsonOk({ responsables: leerResponsables() });
    return jsonOk({ message: 'RACS Minera Casma Apps Script v3 listo' });
  } catch (err) { return jsonErr(err); }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.type === 'foto')     return procesarFoto(body);
    if (body.type === 'firma')    return procesarFirma(body);
    if (body.type === 'reporte')  return procesarReportes(body.data || []);
    if (body.type === 'update')   return procesarUpdate(body.id, body.data);
    if (body.type === 'personal') return procesarPersonal(body.data);
    return jsonErr('Tipo no reconocido: ' + body.type);
  } catch (err) { return jsonErr(err); }
}

// ──────────── FOTO ────────────
function procesarFoto(body) {
  const TIPOS = ['image/jpeg','image/png','image/gif','image/webp'];
  if (!TIPOS.includes(body.mime)) return jsonErr('Tipo no permitido: ' + body.mime);
  const carpeta = DriveApp.getFolderById(CARPETA_DRIVE_ID);
  const blob = Utilities.newBlob(Utilities.base64Decode(body.data), body.mime, body.nombre);
  const archivo = carpeta.createFile(blob);
  archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return jsonOk({ url: 'https://drive.google.com/uc?id=' + archivo.getId(), id: archivo.getId() });
}

// ──────────── FIRMA ────────────
function procesarFirma(body) {
  const carpetaRaiz = DriveApp.getFolderById(CARPETA_DRIVE_ID);
  const firmasIter  = carpetaRaiz.getFoldersByName('Firmas');
  const carpetaFirmas = firmasIter.hasNext() ? firmasIter.next() : carpetaRaiz.createFolder('Firmas');
  const blob = Utilities.newBlob(Utilities.base64Decode(body.data), 'image/png', body.nombre);
  const archivo = carpetaFirmas.createFile(blob);
  archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return jsonOk({ url: 'https://drive.google.com/uc?id=' + archivo.getId(), id: archivo.getId() });
}

// ──────────── REPORTE NUEVO ────────────
function procesarReportes(reportes) {
  const sheet = obtenerHoja(HOJA_REPORTES, HEADERS_REPORTES);
  reportes.forEach(r => {
    const fecha = r.fecha ? new Date(r.fecha + 'T12:00:00') : null;

    // Protección: si llega un dataURL en vez de un link de Drive, guardar vacío
    // Esto evita que Sheets trunque la fila y desplace Mes/Año
    const firmaReg = (r.firmaRegistro && !r.firmaRegistro.startsWith('data:')) ? r.firmaRegistro : '';
    const firmaEd  = (r.firmaEdicion  && !r.firmaEdicion.startsWith('data:'))  ? r.firmaEdicion  : '';

    sheet.appendRow([
      r.id || '',
      r.createdAt || '',
      r.estado || 'Abierto',
      r.tipo || '',
      r.nivelRiesgo || '',
      r.categoria || '',
      r.causaProbable || '',
      r.descripcion || '',
      r.responsable || '',
      r.ubicacion || '',
      r.persona || '',
      r.dniObservado || '',
      r.areaReportado || '',
      r.cargoReportado || '',
      r.fecha || '',
      r.hora || '',
      r.reportador || '',
      r.dniReportador || '',
      r.areaReportador || '',
      r.cargoReportador || '',
      (r.linksFotos || []).join(' | '),
      (r.linksFotosLevantamiento || []).join(' | '),
      r.medidasAcciones || '',
      r.fechaCierre || '',
      r.updatedAt || '',
      firmaReg,
      firmaEd,
      fecha ? fecha.getMonth() + 1 : '',
      fecha ? fecha.getFullYear() : ''
    ]);
  });
  return jsonOk({ inserted: reportes.length });
}

// ──────────── UPDATE REPORTE ────────────
function procesarUpdate(id, data) {
  if (!id) return jsonErr('Falta ID');
  const sheet = obtenerHoja(HOJA_REPORTES, HEADERS_REPORTES);
  const rangeData = sheet.getDataRange().getValues();
  const headers = rangeData[0];
  const idCol = headers.indexOf('ID');
  let rowNum = -1;
  for (let i = 1; i < rangeData.length; i++) {
    if (String(rangeData[i][idCol]) === String(id)) { rowNum = i + 1; break; }
  }
  if (rowNum < 0) return jsonErr('No encontrado: ' + id);

  const set = (col, val) => {
    const idx = headers.indexOf(col);
    if (idx >= 0) sheet.getRange(rowNum, idx + 1).setValue(val);
  };

  set('Estado',          data.estado || 'Abierto');
  set('MedidasAcciones', data.medidasAcciones || '');
  set('FechaCierre',     data.fechaCierre || '');
  set('ActualizadoEn',   data.updatedAt || new Date().toISOString());

  // Protección: solo guardar firma de edición si es un link de Drive
  if (data.firmaEdicion && !data.firmaEdicion.startsWith('data:')) {
    set('FirmaEdicion', data.firmaEdicion);
  }

  if (data.linksFotosLevantamiento && data.linksFotosLevantamiento.length) {
    set('FotosLevantamiento', data.linksFotosLevantamiento.join(' | '));
  }

  return jsonOk({ updated: id });
}

// ──────────── LEER REPORTES ────────────
function leerReportes() {
  const ss = SpreadsheetApp.openById(SHEET_ID_MINA);
  const sheet = ss.getSheetByName(HOJA_REPORTES) || ss.getSheets()[0];
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0].map(h => String(h).trim());
  const MAPA = {
    'ID':'id','CreadoEn':'createdAt','Estado':'estado','Tipo':'tipo',
    'NivelRiesgo':'nivelRiesgo','Categoría':'categoria','CausaProbable':'causaProbable',
    'Descripción':'descripcion','Responsable':'responsable','Ubicación':'ubicacion',
    'PersonaObservada':'persona','DniObservado':'dniObservado',
    'ÁreaObservado':'areaReportado','CargoObservado':'cargoReportado',
    'Fecha':'fecha','Hora':'hora','Reportador':'reportador',
    'DniReportador':'dniReportador','ÁreaReportador':'areaReportador',
    'CargoReportador':'cargoReportador','FotosHallazgo':'linksFotos',
    'FotosLevantamiento':'linksFotosLevantamiento','MedidasAcciones':'medidasAcciones',
    'FechaCierre':'fechaCierre','ActualizadoEn':'updatedAt',
    'FirmaRegistro':'firmaRegistro','FirmaEdicion':'firmaEdicion'
  };
  const colMap = {};
  headers.forEach((h, i) => { if (MAPA[h]) colMap[i] = MAPA[h]; });
  const reportes = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row.every(c => c === '' || c === null || c === undefined)) continue;
    const obj = { synced: true };
    Object.entries(colMap).forEach(([ci, campo]) => {
      let val = row[ci];
      if (val instanceof Date) val = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");
      val = String(val === null || val === undefined ? '' : val).trim();
      if (val) obj[campo] = val;
    });
    if (!obj.id) obj.id = 'sheet_row_' + i;
    if (!obj.estado) obj.estado = 'Abierto';
    if (obj.fecha && /^\d{2}\/\d{2}\/\d{4}$/.test(obj.fecha)) {
      const p = obj.fecha.split('/');
      obj.fecha = p[2] + '-' + p[1] + '-' + p[0];
    }
    reportes.push(obj);
  }
  return reportes.reverse();
}

// ──────────── PERSONAL ────────────
function procesarPersonal(persona) {
  if (!persona || !persona.dni || !persona.nombre) return jsonErr('Falta DNI o nombre');
  const sheet = obtenerHoja(HOJA_PERSONAL, HEADERS_PERSONAL);
  const data = sheet.getDataRange().getValues();
  const dniCol = data[0].indexOf('DNI');
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][dniCol]).trim() === String(persona.dni).trim()) {
      return jsonOk({ existed: true, dni: persona.dni });
    }
  }
  sheet.appendRow([
    persona.dni, persona.nombre, persona.area || '',
    persona.cargo || '', persona.addedAt || new Date().toISOString()
  ]);
  return jsonOk({ added: persona.dni });
}

function leerPersonal() {
  const sheet = obtenerHoja(HOJA_PERSONAL, HEADERS_PERSONAL);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const h = data[0];
  const idx = { dni: h.indexOf('DNI'), nombre: h.indexOf('Nombre'), area: h.indexOf('Área'), cargo: h.indexOf('Cargo') };
  return data.slice(1).filter(r => r[idx.dni]).map(r => ({
    dni:    String(r[idx.dni]).trim(),
    nombre: String(r[idx.nombre] || '').trim(),
    area:   String(r[idx.area] || '').trim(),
    cargo:  String(r[idx.cargo] || '').trim()
  }));
}

// ──────────── RESPONSABLES ────────────
function leerResponsables() {
  const ss = SpreadsheetApp.openById(SHEET_ID_MINA);
  const sheet = ss.getSheetByName(HOJA_RESPONSABLES);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0].map(h => String(h).trim().toLowerCase());
  const iN = headers.indexOf('nombre');
  const iT = headers.indexOf('telefono');
  const iD = headers.indexOf('dni');
  const iC = headers.indexOf('cargo');
  return data.slice(1).filter(r => r[iN]).map(r => ({
    nombre:   String(r[iN] || '').trim(),
    telefono: String(r[iT] || '').trim(),
    dni:      String(r[iD] || '').trim(),
    cargo:    String(r[iC] || '').trim()
  }));
}

// ──────────── HELPERS ────────────
function obtenerHoja(nombre, headersEsperados) {
  const ss = SpreadsheetApp.openById(SHEET_ID_MINA);
  let sheet = ss.getSheetByName(nombre);
  if (!sheet) {
    sheet = ss.insertSheet(nombre);
    sheet.appendRow(headersEsperados);
    sheet.getRange(1, 1, 1, headersEsperados.length)
      .setFontWeight('bold').setBackground('#162236').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    return sheet;
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headersEsperados);
    sheet.getRange(1, 1, 1, headersEsperados.length)
      .setFontWeight('bold').setBackground('#162236').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    return sheet;
  }
  // Agregar columnas faltantes al final (sin tocar las existentes)
  const actuales = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const faltantes = headersEsperados.filter(h => !actuales.includes(h));
  if (faltantes.length) {
    const startCol = sheet.getLastColumn() + 1;
    sheet.getRange(1, startCol, 1, faltantes.length).setValues([faltantes]).setFontWeight('bold');
  }
  return sheet;
}

function jsonOk(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(Object.assign({ ok: true }, payload || {})))
    .setMimeType(ContentService.MimeType.JSON);
}

function jsonErr(err) {
  const msg = err && err.message ? err.message : String(err);
  return ContentService
    .createTextOutput(JSON.stringify({ ok: false, error: msg }))
    .setMimeType(ContentService.MimeType.JSON);
}
