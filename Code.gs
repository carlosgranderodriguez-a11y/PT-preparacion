/**
 * BACKEND — Oposición PT2028 (Pedagogía Terapéutica, oposición Comunidad de Madrid)
 * Conecta profesor.html y alumno.html con un Google Sheet.
 *
 * INSTALACIÓN:
 * 1. Crea un Google Sheet nuevo (vacío, las pestañas se crean solas).
 * 2. Copia su ID (la parte de la URL entre /d/ y /edit) y pégalo abajo en SHEET_ID.
 * 3. Extensiones > Apps Script, borra el contenido de Code.gs y pega todo este archivo.
 * 4. Implementar > Nueva implementación > Tipo: Aplicación web.
 *    - Ejecutar como: Yo (tu cuenta)
 *    - Quién tiene acceso: Cualquier usuario
 * 5. Copia la URL que termina en /exec. Esa es tu API_URL para profesor.html y alumno.html.
 * 6. Autoriza los permisos la primera vez que se ejecute (Drive + Sheets).
 */

const SHEET_ID = '1Z5-EHnE2ZALflhpWRUXH3srSgUvlJU-YL1M-ZuhkkZg';
const DRIVE_FOLDER_NAME = 'PT_Practicos_Fotos';

const HEADERS = {
  Alumnos: ['id', 'nombre', 'email', 'clave', 'etapa', 'discapacidad', 'notas'],
  Calendario: ['id', 'fecha', 'hora', 'alumnoId', 'tipo', 'tema', 'notas'],
  Practicos: ['id', 'alumnoId', 'fecha', 'texto', 'fotoUrl', 'estado', 'feedback', 'fechaFeedback'],
  Temas: ['id', 'numero', 'titulo', 'archivoUrl'],
  TemasProgreso: ['alumnoId', 'temaId', 'estado', 'pct'],
  CG: ['id', 'numero', 'titulo', 'archivoUrl'],
  CGProgreso: ['alumnoId', 'temaId', 'estado', 'pct'],
  Estudio: ['id', 'alumnoId', 'fecha', 'horas', 'bloque', 'notas'],
  // Planificador personal del alumno (privado: solo se envía a su propia cuenta, nunca al panel del profesor). tipo: 'tarea' | 'nota'. dia: 0=lunes … 6=domingo, 7=cualquier día
  Planificador: ['id', 'alumnoId', 'semana', 'dia', 'tipo', 'texto', 'hecho'],
  Objetivos: ['id', 'alumnoId', 'semana', 'texto', 'cumplido'],
  Dafo: ['alumnoId', 'fortalezas', 'debilidades', 'oportunidades', 'amenazas'],
  Archivos: ['id', 'categoria', 'nombre', 'titulo', 'url', 'fecha'],
  Recursos: ['id', 'fecha', 'titulo', 'texto', 'url', 'archivoUrl'],
  Ajustes: ['clave', 'valor'],
  // Cultura General: cuadernos de teoría (bloque = 'lengua' | 'matematicas')
  Cuadernos: ['id', 'bloque', 'titulo', 'descripcion', 'archivoUrl', 'archivoNombre', 'orden', 'fecha'],
  // Cultura General: exámenes con solucionario y puntuación máxima
  // notaMax es una columna heredada que ya no se usa. maxLengua / maxMates / maxCG: puntuación máxima de cada parte (vacío = el examen no incluye esa parte)
  Examenes: ['id', 'titulo', 'descripcion', 'examenUrl', 'examenNombre', 'solucionarioUrl', 'solucionarioNombre', 'notaMax', 'orden', 'fecha', 'maxLengua', 'maxMates', 'maxCG'],
  // Notas que registra cada alumno (un único registro por alumno y examen)
  // nota es una columna heredada que ya no se usa. notaLengua / notaMates / notaCG: nota del alumno en cada parte
  Notas: ['id', 'alumnoId', 'examenId', 'nota', 'fechaRegistro', 'fechaActualizacion', 'notaLengua', 'notaMates', 'notaCG']
};

let _ss_ = null;
function getSpreadsheet_() {
  if (!_ss_) _ss_ = SpreadsheetApp.openById(SHEET_ID);
  return _ss_;
}

function getSheet_(name) {
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

function ensureHeaders_() {
  const cache = CacheService.getScriptCache();
  if (cache.get('headers_ok_v4')) return;
  Object.keys(HEADERS).forEach(name => {
    const sh = getSheet_(name);
    if (sh.getLastRow() === 0) {
      sh.appendRow(HEADERS[name]);
    } else {
      const cur = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
      if (cur.length < HEADERS[name].length) {
        sh.getRange(1, 1, 1, HEADERS[name].length).setValues([HEADERS[name]]);
      }
    }
  });
  try { cache.put('headers_ok_v4', '1', 21600); } catch (err) { /* no pasa nada si falla */ }
}

const DATE_FIELDS = { Calendario: ['fecha'], Practicos: ['fecha', 'fechaFeedback'], Archivos: ['fecha'], Estudio: ['fecha'], Planificador: ['semana'], Objetivos: ['semana'], Recursos: ['fecha'], Cuadernos: ['fecha'], Examenes: ['fecha'], Notas: ['fechaRegistro', 'fechaActualizacion'] };
const TIME_FIELDS = { Calendario: ['hora'] };

function normalizeRow_(sheetName, obj) {
  const tz = Session.getScriptTimeZone();
  (DATE_FIELDS[sheetName] || []).forEach(f => {
    if (obj[f] instanceof Date) obj[f] = Utilities.formatDate(obj[f], tz, 'yyyy-MM-dd');
  });
  (TIME_FIELDS[sheetName] || []).forEach(f => {
    if (obj[f] instanceof Date) obj[f] = Utilities.formatDate(obj[f], tz, 'HH:mm');
  });
  return obj;
}

function sheetToObjects_(name) {
  const sh = getSheet_(name);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0];
  return values.slice(1)
    .filter(r => r.some(c => c !== ''))
    .map(r => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = r[i]);
      return normalizeRow_(name, obj);
    });
}

function appendObject_(name, obj) {
  const sh = getSheet_(name);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const row = headers.map(h => obj[h] !== undefined ? obj[h] : '');
  sh.appendRow(row);
}

// Evita que un texto escrito por un alumno se interprete como fórmula de la hoja (empieza por = + - @)
function safeText_(t) { t = String(t); return /^[=+\-@]/.test(t) ? "'" + t : t; }

function updateObjectById_(name, idField, id, updates) {
  const sh = getSheet_(name);
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  const idCol = headers.indexOf(idField);
  for (let r = 1; r < values.length; r++) {
    if (String(values[r][idCol]) === String(id)) {
      headers.forEach((h, i) => {
        if (updates[h] !== undefined) sh.getRange(r + 1, i + 1).setValue(updates[h]);
      });
      return true;
    }
  }
  return false;
}

function deleteObjectById_(name, idField, id) {
  const sh = getSheet_(name);
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  const idCol = headers.indexOf(idField);
  for (let r = values.length - 1; r >= 1; r--) {
    if (String(values[r][idCol]) === String(id)) {
      sh.deleteRow(r + 1);
      return true;
    }
  }
  return false;
}

function setKeyValue_(sheetName, keyCols, keyVals, valueCol, value) {
  const sh = getSheet_(sheetName);
  const values = sh.getDataRange().getValues();
  for (let r = 1; r < values.length; r++) {
    let match = true;
    keyCols.forEach((col, i) => { if (String(values[r][col]) !== String(keyVals[i])) match = false; });
    if (match) { sh.getRange(r + 1, valueCol + 1).setValue(value); return; }
  }
  const row = [];
  keyCols.forEach((col, i) => row[col] = keyVals[i]);
  row[valueCol] = value;
  sh.appendRow(row);
}

function upsertProgreso_(sheetName, alumnoId, temaId, estado, pct) {
  const sh = getSheet_(sheetName);
  const values = sh.getDataRange().getValues();
  for (let r = 1; r < values.length; r++) {
    if (String(values[r][0]) === String(alumnoId) && String(values[r][1]) === String(temaId)) {
      if (estado !== undefined) sh.getRange(r + 1, 3).setValue(estado);
      if (pct !== undefined) sh.getRange(r + 1, 4).setValue(pct);
      return;
    }
  }
  sh.appendRow([alumnoId, temaId, estado || '', pct !== undefined ? pct : '']);
}

function upsertByKey_(sheetName, keyField, keyValue, updates) {
  const sh = getSheet_(sheetName);
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  const keyCol = headers.indexOf(keyField);
  for (let r = 1; r < values.length; r++) {
    if (String(values[r][keyCol]) === String(keyValue)) {
      headers.forEach((h, i) => { if (updates[h] !== undefined) sh.getRange(r + 1, i + 1).setValue(updates[h]); });
      return;
    }
  }
  const row = headers.map(h => h === keyField ? keyValue : (updates[h] !== undefined ? updates[h] : ''));
  sh.appendRow(row);
}

function guardarEnDrive_(base64Data, nombre, folderName) {
  const folders = DriveApp.getFoldersByName(folderName);
  const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
  const matches = base64Data.match(/^data:(.+);base64,(.*)$/);
  const contentType = matches ? matches[1] : 'application/octet-stream';
  const data = matches ? matches[2] : base64Data;
  const blob = Utilities.newBlob(Utilities.base64Decode(data), contentType, nombre);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return 'https://drive.google.com/file/d/' + file.getId() + '/view';
}

function guardarFoto_(base64Data, nombre) {
  const folders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(DRIVE_FOLDER_NAME);
  const matches = base64Data.match(/^data:(.+);base64,(.*)$/);
  const contentType = matches ? matches[1] : 'image/jpeg';
  const data = matches ? matches[2] : base64Data;
  const blob = Utilities.newBlob(Utilities.base64Decode(data), contentType, nombre);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  if (contentType.indexOf('image/') === 0) {
    return 'https://drive.google.com/uc?id=' + file.getId();
  }
  return 'https://drive.google.com/file/d/' + file.getId() + '/view';
}

/* ============================================================
 * CULTURA GENERAL: cuadernos (Lengua / Matemáticas), exámenes con
 * solucionario y notas autoevaluadas por los alumnos.
 * ============================================================ */
const CG_BLOQUES = ['lengua', 'matematicas'];
const NOTA_MAX_LIMITE = 100;

function hoyISO_() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
function fmtFecha_(v) { return v instanceof Date ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(v || ''); }
function porOrden_(a, b) { return (Number(a.orden) || 0) - (Number(b.orden) || 0); }
function siguienteOrden_(items) { return items.reduce(function (m, x) { return Math.max(m, Number(x.orden) || 0); }, 0) + 1; }
function esPdf_(base64) { return /^data:application\/pdf;base64,/.test(String(base64 || '')); }

function guardarPdf_(base64Data, nombre) {
  if (!esPdf_(base64Data)) throw new Error('Solo se admiten archivos PDF.');
  return guardarEnDrive_(base64Data, nombre || 'documento.pdf', 'PT_Materiales');
}

function papeleraDrive_(url) {
  try {
    const s = String(url || '');
    const m = s.match(/\/d\/([A-Za-z0-9_-]+)/) || s.match(/[?&]id=([A-Za-z0-9_-]+)/);
    if (m) DriveApp.getFileById(m[1]).setTrashed(true);
  } catch (err) { /* si el archivo ya no existe, seguimos */ }
}

// Partes puntuables de un examen: [nombre, columna del máximo, columna de la nota, clave en la petición]
const PARTES = [
  ['Lengua', 'maxLengua', 'notaLengua', 'lengua'],
  ['Matemáticas', 'maxMates', 'notaMates', 'mates'],
  ['Cultura General', 'maxCG', 'notaCG', 'cg']
];
function vacio_(v) { return v === '' || v === null || v === undefined; }

// Devuelve el máximo validado, o '' si la parte no se puntúa (vacío o 0).
function parseMaxParte_(v, nombre) {
  if (vacio_(v)) return '';
  const n = Number(String(v).replace(',', '.'));
  if (!isFinite(n) || n < 0 || n > NOTA_MAX_LIMITE) {
    throw new Error('La puntuación máxima de ' + nombre + ' debe ser un número entre 0 y ' + NOTA_MAX_LIMITE + '.');
  }
  return n === 0 ? '' : Math.round(n * 100) / 100;
}

function reordenar_(nombreHoja, ids) {
  const sh = getSheet_(nombreHoja);
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  const iId = headers.indexOf('id'), iOrden = headers.indexOf('orden');
  (ids || []).forEach(function (id, pos) {
    for (let r = 1; r < values.length; r++) {
      if (String(values[r][iId]) === String(id)) { sh.getRange(r + 1, iOrden + 1).setValue(pos + 1); break; }
    }
  });
}

function addCuaderno_(p) {
  if (CG_BLOQUES.indexOf(p.bloque) < 0) throw new Error('El cuaderno debe asignarse a Lengua o Matemáticas.');
  if (!p.titulo || !String(p.titulo).trim()) throw new Error('El cuaderno necesita un título.');
  if (!esPdf_(p.base64)) throw new Error('Sube el cuaderno en formato PDF.');
  const url = guardarPdf_(p.base64, p.nombre);
  const mismos = sheetToObjects_('Cuadernos').filter(function (c) { return c.bloque === p.bloque; });
  const fila = { id: p.id, bloque: p.bloque, titulo: String(p.titulo).trim(), descripcion: String(p.descripcion || '').trim(), archivoUrl: url, archivoNombre: p.nombre || '', orden: siguienteOrden_(mismos), fecha: hoyISO_() };
  appendObject_('Cuadernos', fila);
  return { ok: true, cuaderno: fila };
}

function updateCuaderno_(p) {
  const actual = sheetToObjects_('Cuadernos').find(function (c) { return String(c.id) === String(p.id); });
  if (!actual) throw new Error('No se encuentra el cuaderno.');
  const upd = {};
  if (p.titulo !== undefined) {
    if (!String(p.titulo).trim()) throw new Error('El cuaderno necesita un título.');
    upd.titulo = String(p.titulo).trim();
  }
  if (p.descripcion !== undefined) upd.descripcion = String(p.descripcion).trim();
  if (p.bloque !== undefined && p.bloque !== actual.bloque) {
    if (CG_BLOQUES.indexOf(p.bloque) < 0) throw new Error('El cuaderno debe asignarse a Lengua o Matemáticas.');
    upd.bloque = p.bloque;
    upd.orden = siguienteOrden_(sheetToObjects_('Cuadernos').filter(function (c) { return c.bloque === p.bloque; }));
  }
  if (p.base64) {
    if (!esPdf_(p.base64)) throw new Error('Sube el cuaderno en formato PDF.');
    upd.archivoUrl = guardarPdf_(p.base64, p.nombre);
    upd.archivoNombre = p.nombre || '';
  }
  updateObjectById_('Cuadernos', 'id', p.id, upd);
  if (upd.archivoUrl) papeleraDrive_(actual.archivoUrl);
  return { ok: true, cuaderno: Object.assign({}, actual, upd) };
}

function deleteCuaderno_(id) {
  const actual = sheetToObjects_('Cuadernos').find(function (c) { return String(c.id) === String(id); });
  if (!actual) return;
  deleteObjectById_('Cuadernos', 'id', id);
  papeleraDrive_(actual.archivoUrl);
}

function maximosDeExamen_(p, actual) {
  const out = {};
  let alguna = false;
  PARTES.forEach(function (pt) {
    const origen = (actual && p[pt[1]] === undefined) ? actual[pt[1]] : p[pt[1]];
    out[pt[1]] = parseMaxParte_(origen, pt[0]);
    if (out[pt[1]] !== '') alguna = true;
  });
  if (!alguna) throw new Error('Activa al menos una parte (Lengua, Matemáticas o Cultura General) indicando su puntuación máxima.');
  return out;
}

function addExamen_(p) {
  if (!p.titulo || !String(p.titulo).trim()) throw new Error('El examen necesita un título.');
  const maximos = maximosDeExamen_(p, null);
  if (!esPdf_(p.examenBase64) || !esPdf_(p.solucionarioBase64)) throw new Error('Sube el PDF del examen y el PDF del solucionario.');
  const examenUrl = guardarPdf_(p.examenBase64, p.examenNombre);
  let solUrl;
  try { solUrl = guardarPdf_(p.solucionarioBase64, p.solucionarioNombre); }
  catch (err) { papeleraDrive_(examenUrl); throw err; }
  const fila = {
    id: p.id, titulo: String(p.titulo).trim(), descripcion: String(p.descripcion || '').trim(),
    examenUrl: examenUrl, examenNombre: p.examenNombre || '',
    solucionarioUrl: solUrl, solucionarioNombre: p.solucionarioNombre || '',
    orden: siguienteOrden_(sheetToObjects_('Examenes')), fecha: hoyISO_()
  };
  Object.keys(maximos).forEach(function (k) { fila[k] = maximos[k]; });
  appendObject_('Examenes', fila);
  return { ok: true, examen: fila };
}

function updateExamen_(p) {
  const actual = sheetToObjects_('Examenes').find(function (e) { return String(e.id) === String(p.id); });
  if (!actual) throw new Error('No se encuentra el examen.');
  const upd = {};
  if (p.titulo !== undefined) {
    if (!String(p.titulo).trim()) throw new Error('El examen necesita un título.');
    upd.titulo = String(p.titulo).trim();
  }
  if (p.descripcion !== undefined) upd.descripcion = String(p.descripcion).trim();

  const hayMax = PARTES.some(function (pt) { return p[pt[1]] !== undefined; });
  if (hayMax) {
    const maximos = maximosDeExamen_(p, actual);
    const notas = sheetToObjects_('Notas').filter(function (n) { return String(n.examenId) === String(p.id); });
    PARTES.forEach(function (pt) {
      const nuevo = maximos[pt[1]];
      notas.forEach(function (n) {
        if (vacio_(n[pt[2]])) return;
        if (nuevo === '') throw new Error('No se puede quitar ' + pt[0] + ': hay alumnos con nota registrada en esa parte.');
        if (Number(n[pt[2]]) > nuevo) throw new Error('No se puede bajar el máximo de ' + pt[0] + ' a ' + nuevo + ': hay alumnos con una nota superior.');
      });
      upd[pt[1]] = nuevo;
    });
  }
  if (p.examenBase64) {
    if (!esPdf_(p.examenBase64)) throw new Error('El examen debe ser un PDF.');
    upd.examenUrl = guardarPdf_(p.examenBase64, p.examenNombre);
    upd.examenNombre = p.examenNombre || '';
  }
  if (p.solucionarioBase64) {
    if (!esPdf_(p.solucionarioBase64)) throw new Error('El solucionario debe ser un PDF.');
    upd.solucionarioUrl = guardarPdf_(p.solucionarioBase64, p.solucionarioNombre);
    upd.solucionarioNombre = p.solucionarioNombre || '';
  }
  // updateObjectById_ ignora los valores undefined pero sí escribe '' (parte desactivada)
  updateObjectById_('Examenes', 'id', p.id, upd);
  if (upd.examenUrl) papeleraDrive_(actual.examenUrl);
  if (upd.solucionarioUrl) papeleraDrive_(actual.solucionarioUrl);
  return { ok: true, examen: Object.assign({}, actual, upd) };
}

function deleteExamen_(id) {
  const actual = sheetToObjects_('Examenes').find(function (e) { return String(e.id) === String(id); });
  if (!actual) return;
  // Borra también las notas registradas de ese examen para no dejar datos huérfanos.
  const sh = getSheet_('Notas');
  const values = sh.getDataRange().getValues();
  const iE = values[0].indexOf('examenId');
  for (let r = values.length - 1; r >= 1; r--) {
    if (String(values[r][iE]) === String(id)) sh.deleteRow(r + 1);
  }
  deleteObjectById_('Examenes', 'id', id);
  papeleraDrive_(actual.examenUrl);
  papeleraDrive_(actual.solucionarioUrl);
}

/**
 * Guarda (o actualiza) las notas de un alumno en un examen: Lengua, Matemáticas y Cultura General.
 * Cada parte es opcional, pero hace falta al menos una. Un único registro por alumno y examen.
 * Valida en el servidor que cada nota esté entre 0 y el máximo de esa parte del examen.
 */
function guardarNota_(p) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const alumno = sheetToObjects_('Alumnos').find(function (a) { return String(a.id) === String(p.alumnoId); });
    if (!alumno) throw new Error('Alumno no encontrado.');
    const examen = sheetToObjects_('Examenes').find(function (e) { return String(e.id) === String(p.examenId); });
    if (!examen) throw new Error('Este examen ya no está disponible.');

    const vals = {};
    let alguna = false;
    PARTES.forEach(function (pt) {
      const max = examen[pt[1]];
      const raw = p[pt[3]];
      if (vacio_(max) || !(Number(max) > 0)) {
        if (!vacio_(raw)) throw new Error(pt[0] + ' no se puntúa en este examen.');
        vals[pt[2]] = '';
        return;
      }
      if (vacio_(raw)) { vals[pt[2]] = ''; return; }
      const n = Number(String(raw).replace(',', '.'));
      if (!isFinite(n)) throw new Error('Introduce una nota válida en ' + pt[0] + '.');
      if (n < 0 || n > Number(max)) throw new Error('La nota de ' + pt[0] + ' debe estar entre 0 y ' + Number(max) + '.');
      vals[pt[2]] = Math.round(n * 100) / 100;
      alguna = true;
    });
    if (!alguna) throw new Error('Introduce al menos una nota.');

    const hoy = hoyISO_();
    const sh = getSheet_('Notas');
    const values = sh.getDataRange().getValues();
    const h = values[0];
    const col = function (name) { return h.indexOf(name); };
    const iA = col('alumnoId'), iE = col('examenId'), iFR = col('fechaRegistro'), iFA = col('fechaActualizacion');
    for (let r = values.length - 1; r >= 1; r--) {
      if (String(values[r][iA]) === String(p.alumnoId) && String(values[r][iE]) === String(p.examenId)) {
        let cambio = false;
        PARTES.forEach(function (pt) {
          const c = col(pt[2]);
          const antes = vacio_(values[r][c]) ? '' : Number(values[r][c]);
          if (antes !== vals[pt[2]]) { sh.getRange(r + 1, c + 1).setValue(vals[pt[2]]); cambio = true; }
        });
        let fa = fmtFecha_(values[r][iFA]);
        if (cambio) { sh.getRange(r + 1, iFA + 1).setValue(hoy); fa = hoy; }
        const reg = { id: values[r][col('id')], alumnoId: p.alumnoId, examenId: p.examenId, fechaRegistro: fmtFecha_(values[r][iFR]), fechaActualizacion: fa };
        Object.keys(vals).forEach(function (k) { reg[k] = vals[k]; });
        return { ok: true, registro: reg };
      }
    }
    const registro = { id: 'nt_' + Date.now() + '_' + Math.random().toString(36).slice(2, 7), alumnoId: p.alumnoId, examenId: p.examenId, fechaRegistro: hoy, fechaActualizacion: '' };
    Object.keys(vals).forEach(function (k) { registro[k] = vals[k]; });
    appendObject_('Notas', registro);
    return { ok: true, registro: registro };
  } finally {
    lock.releaseLock();
  }
}

/**
 * RECORDATORIOS AUTOMÁTICOS POR EMAIL.
 * Para activarlos: en el editor de Apps Script, icono del reloj (Activadores) >
 * "Añadir activador" > función: enviarRecordatorios > Basado en tiempo >
 * Temporizador diario > entre 19:00 y 20:00. Guardar.
 * Cada tarde enviará un email a los alumnos que tengan clase al día siguiente.
 */
function enviarRecordatorios() {
  const tz = Session.getScriptTimeZone();
  const manana = new Date();
  manana.setDate(manana.getDate() + 1);
  const fManana = Utilities.formatDate(manana, tz, 'yyyy-MM-dd');
  const alumnos = sheetToObjects_('Alumnos');
  const eventos = sheetToObjects_('Calendario');
  const ajustes = sheetToObjects_('Ajustes');
  const academiaReg = ajustes.find(function(a){ return a.clave === 'academia'; });
  const academia = academiaReg ? academiaReg.valor : 'Oposición PT2028';
  alumnos.forEach(function(al) {
    if (!al.email) return;
    const evs = eventos.filter(function(ev) {
      return (ev.alumnoId === al.id || ev.alumnoId === 'ALL') && String(ev.fecha) === fManana;
    });
    if (!evs.length) return;
    let cuerpo = 'Hola ' + al.nombre + ',\n\nMañana tienes programado:\n\n';
    evs.forEach(function(ev) {
      cuerpo += '• ' + (ev.hora ? ev.hora + ' — ' : '') + (ev.tipo || '') + (ev.tema ? ': ' + ev.tema : '') + (ev.notas ? ' (' + ev.notas + ')' : '') + '\n';
    });
    cuerpo += '\n¡A por ello!\n\n' + academia;
    MailApp.sendEmail(al.email, '📅 Recordatorio: mañana tienes clase', cuerpo);
  });
}

function enviarBienvenida_(alumno) {
  try {
    if (!alumno.email) return;
    const ajustes = sheetToObjects_('Ajustes');
    const urlReg = ajustes.find(function(a){ return a.clave === 'urlAlumno'; });
    const academiaReg = ajustes.find(function(a){ return a.clave === 'academia'; });
    const url = (urlReg && urlReg.valor) ? urlReg.valor : 'https://carlosgranderodriguez-a11y.github.io/PT-preparacion/alumno.html';
    const academia = academiaReg && academiaReg.valor ? academiaReg.valor : 'Oposición PT2028';
    let cuerpo = 'Hola, me llamo Sara y quiero darte la enhorabuena por la decisión que has tomado. Esto es el principio del camino. ¡Bienvenido/a a tu preparación para las oposiciones de PT! 💪\n\n';
    cuerpo += 'Accede a tu espacio personal aquí:\n' + url + '\n\n';
    cuerpo += 'La primera vez, pulsa "¿Primera vez? Crea tu clave aquí", escribe este mismo correo (' + alumno.email + ') y elige tu clave personal. A partir de entonces entrarás con tu correo y tu clave.\n\n';
    cuerpo += 'En tu espacio verás tu calendario de clases, enviarás tus prácticos, descargarás materiales y llevarás el control de tu estudio.\n\n';
    cuerpo += '¡Mucho ánimo!\n\nSara — ' + academia;
    MailApp.sendEmail(alumno.email, '👋 Bienvenido/a a ' + academia + ' — tu acceso', cuerpo);
  } catch (err) {
    // No bloquear el alta si el correo falla
  }
}

function buildFullData_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('getAll_v2');
  if (cached) {
    try { return JSON.parse(cached); } catch (err) { /* si el caché está corrupto, recalculamos */ }
  }
  const result = {
    alumnos: sheetToObjects_('Alumnos').map(function(a){
      var tiene = !!a.clave;
      delete a.clave;
      a.tieneClave = tiene;
      return a;
    }),
    calendario: sheetToObjects_('Calendario'),
    practicos: sheetToObjects_('Practicos'),
    temas: sheetToObjects_('Temas'),
    temasProgreso: sheetToObjects_('TemasProgreso'),
    cg: sheetToObjects_('CG'),
    cgProgreso: sheetToObjects_('CGProgreso'),
    dafo: sheetToObjects_('Dafo'),
    archivos: sheetToObjects_('Archivos'),
    recursos: sheetToObjects_('Recursos'),
    estudio: sheetToObjects_('Estudio'),
    objetivos: sheetToObjects_('Objetivos'),
    ajustes: sheetToObjects_('Ajustes'),
    cuadernos: sheetToObjects_('Cuadernos').sort(porOrden_),
    examenes: sheetToObjects_('Examenes').sort(porOrden_),
    notas: sheetToObjects_('Notas')
  };
  try { cache.put('getAll_v2', JSON.stringify(result), 300); } catch (err) { /* si supera 100KB, seguimos sin caché */ }
  return result;
}

function buildAlumnoScoped_(id) {
  const full = buildFullData_();
  const alumno = full.alumnos.find(function(a){ return a.id === id; }) || null;
  return {
    alumno: alumno,
    calendario: full.calendario.filter(function(ev){ return ev.alumnoId === id || ev.alumnoId === 'ALL'; }),
    practicos: full.practicos.filter(function(p){ return p.alumnoId === id; }),
    temas: full.temas,
    temasProgreso: full.temasProgreso.filter(function(x){ return x.alumnoId === id; }),
    cg: full.cg,
    cgProgreso: full.cgProgreso.filter(function(x){ return x.alumnoId === id; }),
    archivos: full.archivos,
    recursos: full.recursos,
    estudio: full.estudio.filter(function(x){ return x.alumnoId === id; }),
    objetivos: full.objetivos.filter(function(o){ return o.alumnoId === id || o.alumnoId === 'ALL'; }),
    ajustes: full.ajustes,
    cuadernos: full.cuadernos,
    examenes: full.examenes,
    notas: full.notas.filter(function(n){ return String(n.alumnoId) === String(id); }),
    // Privado: solo las filas de este alumno; no forma parte de getAll (panel del profesor)
    planificador: sheetToObjects_('Planificador').filter(function(x){ return String(x.alumnoId) === String(id); })
  };
}

function doGet(e) {
  ensureHeaders_();
  const action = e.parameter.action;

  if (action === 'getAll') {
    return ContentService.createTextOutput(JSON.stringify(buildFullData_())).setMimeType(ContentService.MimeType.JSON);
  }

  if (action === 'getAlumnoData') {
    const id = e.parameter.id;
    return ContentService.createTextOutput(JSON.stringify(buildAlumnoScoped_(id))).setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService.createTextOutput(JSON.stringify({ error: 'acción desconocida' })).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  ensureHeaders_();
  let result = { ok: true };
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    const p = body.payload || {};

    if (action !== 'loginAlumno') {
      try { CacheService.getScriptCache().remove('getAll_v2'); } catch (err) { /* no pasa nada si falla */ }
    }

    switch (action) {
      case 'addAlumno':
        appendObject_('Alumnos', p);
        enviarBienvenida_(p);
        break;
      case 'updateAlumno':
        updateObjectById_('Alumnos', 'id', p.id, p);
        break;
      case 'deleteAlumno':
        deleteObjectById_('Alumnos', 'id', p.id);
        break;
      case 'addEvento':
        appendObject_('Calendario', p);
        break;
      case 'deleteEvento':
        deleteObjectById_('Calendario', 'id', p.id);
        break;
      case 'addPractico':
        if (p.fotoBase64) {
          p.fotoUrl = guardarFoto_(p.fotoBase64, p.fotoNombre || ('practico_' + p.id + '.jpg'));
          delete p.fotoBase64;
        }
        p.estado = 'pendiente';
        appendObject_('Practicos', p);
        break;
      case 'darFeedback':
        updateObjectById_('Practicos', 'id', p.id, {
          estado: 'corregido',
          feedback: p.feedback,
          fechaFeedback: p.fechaFeedback
        });
        break;
      case 'addTema':
        if (p.base64) {
          p.archivoUrl = guardarEnDrive_(p.base64, p.nombre || ('tema_' + p.id), 'PT_Materiales');
          delete p.base64;
          delete p.nombre;
        }
        appendObject_(p.tipo === 'cg' ? 'CG' : 'Temas', { id: p.id, numero: p.numero, titulo: p.titulo, archivoUrl: p.archivoUrl || '' });
        result = { ok: true, url: p.archivoUrl || '' };
        break;
      case 'deleteTema':
        deleteObjectById_(p.tipo === 'cg' ? 'CG' : 'Temas', 'id', p.id);
        break;
      case 'setArchivoTema': {
        var urlTema = '';
        if (p.base64) {
          urlTema = guardarEnDrive_(p.base64, p.nombre || ('tema_' + p.id), 'PT_Materiales');
        }
        updateObjectById_(p.tipo === 'cg' ? 'CG' : 'Temas', 'id', p.id, { archivoUrl: urlTema });
        result = { ok: true, url: urlTema };
        break;
      }
      case 'setProgreso':
        upsertProgreso_(p.tipo === 'cg' ? 'CGProgreso' : 'TemasProgreso', p.alumnoId, p.temaId, p.estado, p.pct);
        break;
      case 'addPlan': {
        var texto = String(p.texto || '').trim();
        if (!p.alumnoId || !p.id) throw new Error('Datos incompletos.');
        if (!texto) throw new Error('Escribe la tarea.');
        if (texto.length > 500) throw new Error('La tarea es demasiado larga (máximo 500 caracteres).');
        var dia = Number(p.dia);
        if (!(dia >= 0 && dia <= 7)) throw new Error('Día no válido.');
        appendObject_('Planificador', { id: p.id, alumnoId: p.alumnoId, semana: p.semana, dia: dia, tipo: 'tarea', texto: safeText_(texto), hecho: false });
        break;
      }
      case 'updatePlan': {
        var fila = sheetToObjects_('Planificador').find(function (x) { return String(x.id) === String(p.id); });
        if (!fila || String(fila.alumnoId) !== String(p.alumnoId)) throw new Error('No se encuentra la tarea.');
        var cambios = {};
        if (p.hecho !== undefined) cambios.hecho = (p.hecho === true || p.hecho === 'true');
        if (p.texto !== undefined) {
          var t = String(p.texto).trim();
          if (!t) throw new Error('La tarea no puede quedar vacía.');
          if (t.length > 500) throw new Error('La tarea es demasiado larga (máximo 500 caracteres).');
          cambios.texto = safeText_(t);
        }
        if (p.dia !== undefined) {
          var d2 = Number(p.dia);
          if (!(d2 >= 0 && d2 <= 7)) throw new Error('Día no válido.');
          cambios.dia = d2;
        }
        updateObjectById_('Planificador', 'id', p.id, cambios);
        break;
      }
      case 'deletePlan': {
        var f2 = sheetToObjects_('Planificador').find(function (x) { return String(x.id) === String(p.id); });
        if (f2 && String(f2.alumnoId) === String(p.alumnoId)) deleteObjectById_('Planificador', 'id', p.id);
        break;
      }
      case 'saveNotaSemana': {
        var txtN = String(p.texto || '');
        if (!p.alumnoId || !p.semana) throw new Error('Datos incompletos.');
        if (txtN.length > 5000) throw new Error('La nota es demasiado larga (máximo 5000 caracteres).');
        var lockN = LockService.getScriptLock();
        lockN.waitLock(20000);
        try {
          var ex = sheetToObjects_('Planificador').find(function (x) {
            return x.tipo === 'nota' && String(x.alumnoId) === String(p.alumnoId) && String(x.semana) === String(p.semana);
          });
          if (ex) updateObjectById_('Planificador', 'id', ex.id, { texto: safeText_(txtN) });
          else appendObject_('Planificador', { id: p.id, alumnoId: p.alumnoId, semana: p.semana, dia: '', tipo: 'nota', texto: safeText_(txtN), hecho: false });
        } finally { lockN.releaseLock(); }
        break;
      }
      case 'addEstudio':
        appendObject_('Estudio', p);
        break;
      case 'deleteEstudio':
        deleteObjectById_('Estudio', 'id', p.id);
        break;
      case 'addObjetivo':
        appendObject_('Objetivos', p);
        break;
      case 'deleteObjetivo':
        deleteObjectById_('Objetivos', 'id', p.id);
        break;
      case 'setObjetivoCumplido':
        updateObjectById_('Objetivos', 'id', p.id, { cumplido: p.cumplido });
        break;
      case 'subirArchivo':
        var urlArchivo = guardarEnDrive_(p.base64, p.nombre, 'PT_Materiales');
        appendObject_('Archivos', { id: p.id, categoria: p.categoria, nombre: p.nombre, titulo: p.titulo || '', url: urlArchivo, fecha: p.fecha });
        result = { ok: true, url: urlArchivo };
        break;
      case 'deleteArchivo':
        sheetToObjects_('Archivos').filter(function (x) { return String(x.id) === String(p.id); }).forEach(function (x) { papeleraDrive_(x.url); });
        deleteObjectById_('Archivos', 'id', p.id);
        break;
      case 'deletePractico':
        sheetToObjects_('Practicos').filter(function (x) { return String(x.id) === String(p.id); }).forEach(function (x) { papeleraDrive_(x.fotoUrl); });
        deleteObjectById_('Practicos', 'id', p.id);
        break;
      case 'addRecurso':
        if (p.base64) {
          p.archivoUrl = guardarEnDrive_(p.base64, p.nombre || ('recurso_' + p.id), 'PT_Materiales');
          delete p.base64;
          delete p.nombre;
        }
        appendObject_('Recursos', p);
        result = { ok: true, url: p.archivoUrl || '' };
        break;
      case 'deleteRecurso':
        deleteObjectById_('Recursos', 'id', p.id);
        break;
      case 'saveDafo':
        upsertByKey_('Dafo', 'alumnoId', p.alumnoId, p);
        break;
      case 'setClaveAlumno': {
        var alumnos1 = sheetToObjects_('Alumnos');
        var al1 = alumnos1.find(function(a){ return String(a.email).toLowerCase().trim() === String(p.email).toLowerCase().trim(); });
        if (!al1) { result = { error: 'No hay ningún alumno registrado con ese correo. Habla con tu preparador.' }; break; }
        if (al1.clave) { result = { error: 'Este correo ya tiene una clave creada. Si la has olvidado, pide a tu preparador que la restablezca.' }; break; }
        updateObjectById_('Alumnos', 'id', al1.id, { clave: p.clave });
        try { CacheService.getScriptCache().remove('getAll_v2'); } catch (errc) {}
        result = { ok: true, alumnoId: al1.id, nombre: al1.nombre, data: buildAlumnoScoped_(al1.id) };
        break;
      }
      case 'loginAlumno': {
        var alumnos2 = sheetToObjects_('Alumnos');
        var al2 = alumnos2.find(function(a){ return String(a.email).toLowerCase().trim() === String(p.email).toLowerCase().trim(); });
        if (!al2) { result = { error: 'Correo o clave incorrectos.' }; break; }
        if (!al2.clave) { result = { error: 'Todavía no has creado tu clave. Usa la opción "Primera vez".' }; break; }
        if (String(al2.clave) !== String(p.clave)) { result = { error: 'Correo o clave incorrectos.' }; break; }
        result = { ok: true, alumnoId: al2.id, nombre: al2.nombre, data: buildAlumnoScoped_(al2.id) };
        break;
      }
      case 'resetClaveAlumno':
        updateObjectById_('Alumnos', 'id', p.id, { clave: '' });
        break;
      case 'addCuaderno':
        result = addCuaderno_(p);
        break;
      case 'updateCuaderno':
        result = updateCuaderno_(p);
        break;
      case 'deleteCuaderno':
        deleteCuaderno_(p.id);
        break;
      case 'reorderCuadernos':
        reordenar_('Cuadernos', p.ids);
        break;
      case 'addExamen':
        result = addExamen_(p);
        break;
      case 'updateExamen':
        result = updateExamen_(p);
        break;
      case 'deleteExamen':
        deleteExamen_(p.id);
        break;
      case 'reorderExamenes':
        reordenar_('Examenes', p.ids);
        break;
      case 'guardarNota':
        result = guardarNota_(p);
        break;
      case 'saveAjuste':
        setKeyValue_('Ajustes', [0], [p.clave], 1, p.valor);
        break;
      default:
        result = { error: 'acción desconocida: ' + action };
    }
  } catch (err) {
    result = { error: err.message };
  }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}
