// ═══════════════════════════════════════════════════════════════════════
// ALMA SALUD — Apps Script para Viajes de Egresados
// ───────────────────────────────────────────────────────────────────────
// INSTRUCCIONES DE DEPLOY:
//   1. Abrí Google Apps Script en la Spreadsheet de Alma Salud
//   2. Pegá este código completo
//   3. Implementar > Nueva implementación > Aplicación web
//   4. Ejecutar como: "Yo (tu cuenta)"
//   5. Quién tiene acceso: "Cualquier persona"
//   6. Copiá la URL de implementación y reemplazá REEMPLAZAR_CON_URL_APPS_SCRIPT en los 3 HTML
//
// HOJAS QUE SE CREAN AUTOMÁTICAMENTE:
//   · Viajes        — registro de viajes creados desde el panel admin
//   · Participantes — lista de egresados con sus fichas médicas
//   · Atenciones    — registros de atención médica durante el viaje
// ═══════════════════════════════════════════════════════════════════════

// ─── Encabezados de cada hoja ─────────────────────────────────────────
const VIAJES_HEADERS = ['id', 'nombre', 'organizadora', 'fechaCreacion'];

const PART_HEADERS = [
  'viajeId', 'id', 'nombre', 'contacto', 'fichaCompletada', 'fechaFicha',
  'edad', 'dni', 'sexo', 'grupoSanguineo', 'enfermedadesCronicas',
  'medicacion', 'alergias', 'cirugias',
  'contactoEmergenciaNombre', 'contactoEmergenciaTel',
  'obraSocial', 'numAfiliado', 'tutorNombre', 'tutorTel', 'obs', 'colegio'
];

const ATENCION_HEADERS = [
  'viajeId', 'id', 'fecha', 'hora', 'sede', 'profesional',
  'participanteId', 'participante', 'edad', 'colegio',
  'motivo', 'zona', 'diagnostico', 'atencion',
  'derivacion', 'derDetalle', 'obs', 'fechaRegistro'
];

// ─── doPost: recibe datos desde los paneles HTML ───────────────────────
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let result;
    switch (data.tipo) {
      case 'crear_viaje':           result = crearViaje(ss, data);          break;
      case 'agregar_participantes': result = agregarParticipantes(ss, data); break;
      case 'ficha_medica':          result = guardarFicha(ss, data);         break;
      case 'registrar_atencion':    result = registrarAtencion(ss, data);    break;
      case 'eliminar_participante': result = eliminarParticipante(ss, data); break;
      default: result = { ok: false, error: 'Tipo desconocido: ' + data.tipo };
    }
    return jsonOut(result);
  } catch (err) {
    return jsonOut({ ok: false, error: err.message });
  }
}

// ─── doGet: sirve datos a los paneles HTML ─────────────────────────────
function doGet(e) {
  try {
    const p = e.parameter;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let result;
    switch (p.tipo) {
      case 'get_all_viajes':    result = getAllViajes(ss);                            break;
      case 'get_viaje':        result = getViaje(ss, p.id);                          break;
      case 'get_participantes': result = getParticipantes(ss, p.viajeId);            break;
      case 'get_participante':  result = getParticipante(ss, p.viajeId, p.participanteId); break;
      case 'get_atenciones':   result = getAtenciones(ss, p.viajeId);               break;
      default: result = { ok: false, error: 'Tipo desconocido: ' + p.tipo };
    }
    return jsonOut(result);
  } catch (err) {
    return jsonOut({ ok: false, error: err.message });
  }
}

// ─── Helpers comunes ──────────────────────────────────────────────────
function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreate(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#003D6B')
      .setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function sheetToObjects(sheet, filterFn) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0].map(String);
  return data.slice(1)
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = row[i]; });
      return obj;
    })
    .filter(filterFn || (() => true));
}

function rowToValues(headers, data) {
  return headers.map(h => (data[h] !== undefined && data[h] !== null) ? data[h] : '');
}

// ─── Escritura: crear_viaje ────────────────────────────────────────────
function crearViaje(ss, data) {
  const sheet = getOrCreate(ss, 'Viajes', VIAJES_HEADERS);
  const existing = sheetToObjects(sheet, r => String(r.id) === String(data.id));
  if (existing.length > 0) return { ok: true, msg: 'El viaje ya existía' };
  sheet.appendRow(rowToValues(VIAJES_HEADERS, data));
  return { ok: true };
}

// ─── Escritura: agregar_participantes (acepta array) ──────────────────
function agregarParticipantes(ss, data) {
  const sheet = getOrCreate(ss, 'Participantes', PART_HEADERS);
  const lista = Array.isArray(data.participantes) ? data.participantes : [data];
  const viajeId = data.viajeId;

  // Obtener IDs ya existentes para no duplicar
  const existentes = sheetToObjects(sheet, r => String(r.viajeId) === String(viajeId))
    .map(r => String(r.id));

  let added = 0;
  lista.forEach(p => {
    if (existentes.includes(String(p.id))) return;
    sheet.appendRow(rowToValues(PART_HEADERS, {
      viajeId, id: p.id, nombre: p.nombre, contacto: p.contacto || '',
      fichaCompletada: false, fechaFicha: '', colegio: p.colegio || ''
    }));
    added++;
  });
  return { ok: true, added };
}

// ─── Escritura: ficha_medica — actualiza fila del participante ─────────
function guardarFicha(ss, data) {
  const sheet = getOrCreate(ss, 'Participantes', PART_HEADERS);
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(String);
  const viajeIdIdx = headers.indexOf('viajeId');
  const idIdx      = headers.indexOf('id');

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][viajeIdIdx]) === String(data.viajeId) &&
        String(values[i][idIdx])      === String(data.participanteId)) {

      const updated = [...values[i]];
      headers.forEach((h, col) => {
        if (h === 'fichaCompletada') { updated[col] = true; return; }
        if (h === 'fechaFicha')      { updated[col] = data.fechaEnvio || new Date().toISOString(); return; }
        // Para el campo 'id' del participante mapeamos desde participanteId
        if (h === 'id')              { return; } // no sobreescribir
        const val = data[h];
        if (val !== undefined && val !== null && val !== '') updated[col] = val;
      });
      sheet.getRange(i + 1, 1, 1, updated.length).setValues([updated]);
      return { ok: true };
    }
  }

  // No se encontró: puede pasar si el participante se agregó desde otro dispositivo
  // antes de que Sheets sincronizara. Insertamos la ficha como nueva fila.
  sheet.appendRow(rowToValues(PART_HEADERS, {
    viajeId:          data.viajeId,
    id:               data.participanteId,
    nombre:           data.nombre || '',
    contacto:         '',
    fichaCompletada:  true,
    fechaFicha:       data.fechaEnvio || new Date().toISOString(),
    edad:             data.edad || '',
    dni:              data.dni || '',
    sexo:             data.sexo || '',
    grupoSanguineo:   data.grupoSanguineo || '',
    enfermedadesCronicas: data.enfermedadesCronicas || '',
    medicacion:       data.medicacion || '',
    alergias:         data.alergias || '',
    cirugias:         data.cirugias || '',
    contactoEmergenciaNombre: data.contactoEmergenciaNombre || '',
    contactoEmergenciaTel:    data.contactoEmergenciaTel || '',
    obraSocial:       data.obraSocial || '',
    numAfiliado:      data.numAfiliado || '',
    tutorNombre:      data.tutorNombre || '',
    tutorTel:         data.tutorTel || '',
    obs:              data.obs || '',
    colegio:          data.colegio || ''
  }));
  return { ok: true, msg: 'Participante creado con ficha' };
}

// ─── Escritura: registrar_atencion ────────────────────────────────────
function registrarAtencion(ss, data) {
  const sheet = getOrCreate(ss, 'Atenciones', ATENCION_HEADERS);
  const row = rowToValues(ATENCION_HEADERS, {
    ...data,
    fechaRegistro: data.fechaRegistro || new Date().toISOString()
  });
  sheet.appendRow(row);
  return { ok: true };
}

// ─── Lectura: get_all_viajes ──────────────────────────────────────────
function getAllViajes(ss) {
  const sheet = ss.getSheetByName('Viajes');
  if (!sheet) return { ok: true, viajes: [] };
  return { ok: true, viajes: sheetToObjects(sheet) };
}

// ─── Lectura: get_viaje ───────────────────────────────────────────────
function getViaje(ss, id) {
  const sheet = ss.getSheetByName('Viajes');
  if (!sheet) return { ok: true, viaje: null };
  const rows = sheetToObjects(sheet, r => String(r.id) === String(id));
  return { ok: true, viaje: rows[0] || null };
}

// ─── Lectura: get_participantes ───────────────────────────────────────
function getParticipantes(ss, viajeId) {
  const sheet = ss.getSheetByName('Participantes');
  if (!sheet) return { ok: true, participantes: [] };
  const rows = sheetToObjects(sheet, r => String(r.viajeId) === String(viajeId));
  rows.forEach(r => {
    r.fichaCompletada = r.fichaCompletada === true ||
      String(r.fichaCompletada).toUpperCase() === 'TRUE';
  });
  return { ok: true, participantes: rows };
}

// ─── Lectura: get_participante (individual) ───────────────────────────
function getParticipante(ss, viajeId, participanteId) {
  const sheet = ss.getSheetByName('Participantes');
  if (!sheet) return { ok: true, participante: null };
  const rows = sheetToObjects(sheet, r =>
    String(r.viajeId) === String(viajeId) &&
    String(r.id)      === String(participanteId)
  );
  if (rows.length > 0) {
    rows[0].fichaCompletada = rows[0].fichaCompletada === true ||
      String(rows[0].fichaCompletada).toUpperCase() === 'TRUE';
  }
  return { ok: true, participante: rows[0] || null };
}

// ─── Escritura: eliminar_participante ────────────────────────────────
function eliminarParticipante(ss, data) {
  const sheet = ss.getSheetByName('Participantes');
  if (!sheet) return { ok: false, error: 'Hoja no encontrada' };
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(String);
  const viajeIdIdx = headers.indexOf('viajeId');
  const idIdx      = headers.indexOf('id');
  for (let i = values.length - 1; i >= 1; i--) {
    if (String(values[i][viajeIdIdx]) === String(data.viajeId) &&
        String(values[i][idIdx])      === String(data.participanteId)) {
      sheet.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { ok: true, msg: 'No encontrado en Sheets' };
}

// ─── Lectura: get_atenciones ──────────────────────────────────────────
function getAtenciones(ss, viajeId) {
  const sheet = ss.getSheetByName('Atenciones');
  if (!sheet) return { ok: true, atenciones: [] };
  const rows = sheetToObjects(sheet, r => String(r.viajeId) === String(viajeId));
  return { ok: true, atenciones: rows };
}
