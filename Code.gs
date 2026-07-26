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
  Objetivos: ['id', 'alumnoId', 'semana', 'texto', 'cumplido'],
  Dafo: ['alumnoId', 'fortalezas', 'debilidades', 'oportunidades', 'amenazas'],
  Archivos: ['id', 'categoria', 'nombre', 'titulo', 'url', 'fecha'],
  Recursos: ['id', 'fecha', 'titulo', 'texto', 'url', 'archivoUrl'],
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
    if (sh.getLastRow() === 0) {
      sh.appendRow(HEADERS[name]);
    } else {
      const cur = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
      if (cur.length < HEADERS[name].length) {
        sh.getRange(1, 1, 1, HEADERS[name].length).setValues([HEADERS[name]]);
      }
    }
  });
}

const DATE_FIELDS = { Calendario: ['fecha'], Practicos: ['fecha', 'fechaFeedback'], Archivos: ['fecha'], Estudio: ['fecha'], Objetivos: ['semana'], Recursos: ['fecha'] };
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
  const cached = cache.get('getAll_v1');
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
    ajustes: sheetToObjects_('Ajustes')
  };
  try { cache.put('getAll_v1', JSON.stringify(result), 30); } catch (err) { /* si supera 100KB, seguimos sin caché */ }
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
    ajustes: full.ajustes
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
      try { CacheService.getScriptCache().remove('getAll_v1'); } catch (err) { /* no pasa nada si falla */ }
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
        deleteObjectById_('Archivos', 'id', p.id);
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
        try { CacheService.getScriptCache().remove('getAll_v1'); } catch (errc) {}
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
