/**
 * BACKEND — Academia PT (Pedagogía Terapéutica, oposición Comunidad de Madrid)
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
  Temas: ['id', 'numero', 'titulo'],
  TemasProgreso: ['alumnoId', 'temaId', 'estado'],
  CG: ['id', 'numero', 'titulo'],
  CGProgreso: ['alumnoId', 'temaId', 'estado'],
  Dafo: ['alumnoId', 'fortalezas', 'debilidades', 'oportunidades', 'amenazas'],
  Archivos: ['id', 'categoria', 'nombre', 'url', 'fecha'],
  Ajustes: ['clave', 'valor']
};

function getSheet_(name) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

function ensureHeaders_() {
  Object.keys(HEADERS).forEach(name => {
    const sh = getSheet_(name);
    if (sh.getLastRow() === 0) sh.appendRow(HEADERS[name]);
  });
}

const DATE_FIELDS = { Calendario: ['fecha'], Practicos: ['fecha', 'fechaFeedback'], Archivos: ['fecha'] };
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
  return 'https://drive.google.com/uc?id=' + file.getId();
}

function doGet(e) {
  ensureHeaders_();
  const action = e.parameter.action;
  let result;
  if (action === 'getAll') {
    result = {
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
      ajustes: sheetToObjects_('Ajustes')
    };
  } else {
    result = { error: 'acción desconocida' };
  }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  ensureHeaders_();
  let result = { ok: true };
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    const p = body.payload || {};

    switch (action) {
      case 'addAlumno':
        appendObject_('Alumnos', p);
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
        appendObject_(p.tipo === 'cg' ? 'CG' : 'Temas', { id: p.id, numero: p.numero, titulo: p.titulo });
        break;
      case 'deleteTema':
        deleteObjectById_(p.tipo === 'cg' ? 'CG' : 'Temas', 'id', p.id);
        break;
      case 'setProgreso':
        setKeyValue_(p.tipo === 'cg' ? 'CGProgreso' : 'TemasProgreso', [0, 1], [p.alumnoId, p.temaId], 2, p.estado);
        break;
      case 'subirArchivo':
        var urlArchivo = guardarEnDrive_(p.base64, p.nombre, 'PT_Materiales');
        appendObject_('Archivos', { id: p.id, categoria: p.categoria, nombre: p.nombre, url: urlArchivo, fecha: p.fecha });
        result = { ok: true, url: urlArchivo };
        break;
      case 'deleteArchivo':
        deleteObjectById_('Archivos', 'id', p.id);
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
        result = { ok: true, alumnoId: al1.id, nombre: al1.nombre };
        break;
      }
      case 'loginAlumno': {
        var alumnos2 = sheetToObjects_('Alumnos');
        var al2 = alumnos2.find(function(a){ return String(a.email).toLowerCase().trim() === String(p.email).toLowerCase().trim(); });
        if (!al2) { result = { error: 'Correo o clave incorrectos.' }; break; }
        if (!al2.clave) { result = { error: 'Todavía no has creado tu clave. Usa la opción "Primera vez".' }; break; }
        if (String(al2.clave) !== String(p.clave)) { result = { error: 'Correo o clave incorrectos.' }; break; }
        result = { ok: true, alumnoId: al2.id, nombre: al2.nombre };
        break;
      }
      case 'resetClaveAlumno':
        updateObjectById_('Alumnos', 'id', p.id, { clave: '' });
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
