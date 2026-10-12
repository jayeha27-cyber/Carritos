// Carritos de publicaciones: recibe los avisos de la página, los guarda en esta hoja
// y envía un correo a las direcciones de la pestaña "Correos".
// Si un carrito lleva más de 24 horas fuera, pinta la fila de rojo y avisa por correo.
const ZONA = 'America/Puerto_Rico';
const HORAS_LIMITE = 24;
const ENCABEZADOS = ['Fecha y hora', 'Carrito', 'Tipo', 'Nombre', 'Localidad', 'Revisión',
  'Pendientes en la revisión', 'Faltan publicaciones', 'Problema', 'Detalles',
  'Estado', 'Marca de tiempo'];
const C_TIPO = 3, C_CARRITO = 2, C_ESTADO = 11, C_MARCA = 12;
const FUERA = 'Fuera', TARDE = '⚠️ Más de 24 h';
const ROJO = '#f4c7c3', VERDE = '#d9ead3';
const VERBO = { Recogido: 'recogido por', Devuelto: 'devuelto por', Problema: 'problema reportado por' };

function preparar() {
  const ss = SpreadsheetApp.getActive();
  hojaAvisos_(ss);
  correos_(ss);
  MailApp.sendEmail(Session.getEffectiveUser().getEmail(), 'Carritos: listo',
    'La hoja de avisos de los carritos ya está preparada.');
}

// Ejecutar UNA vez: revisa cada hora si algún carrito lleva más de 24 h fuera.
function activarAviso24h() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'revisarPendientes')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('revisarPendientes').timeBased().everyHours(1).create();
  revisarPendientes();
}

function doGet() {
  return salida_({ ok: true });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const d = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const tipo = VERBO[d.tipo] ? d.tipo : null;
    const carrito = /^Carrito [1-9]\d?$/.test(d.carrito || '') ? d.carrito : null;
    const nombre = t_(d.nombre, 80);
    if (!tipo || !carrito || nombre.split(' ').length < 2) {
      return salida_({ ok: false, error: 'Datos incompletos' });
    }
    const ahora = new Date();
    const fila = [
      fecha_(ahora), carrito, tipo, nombre,
      t_(d.localidad, 120), t_(d.revision, 300), t_(d.pendientes, 300),
      t_(d.faltan, 300), t_(d.problema, 80), t_(d.detalles, 1000),
      tipo === 'Recogido' ? FUERA : '', ahora
    ];
    const ss = SpreadsheetApp.getActive();
    const h = hojaAvisos_(ss);
    if (tipo === 'Devuelto') cerrarRecogido_(h, carrito, ahora);
    h.appendRow(fila);

    const alerta = tipo === 'Problema' || fila[6] || fila[7];
    const asunto = (alerta ? '⚠️ ' : '') + carrito + ': ' + VERBO[tipo] + ' ' + nombre;
    const cuerpo = ENCABEZADOS.slice(0, 10).map((t, i) => (fila[i] ? t + ': ' + fila[i] : ''))
      .filter(Boolean).join('\n');
    const para = correos_(ss);
    if (para.length) MailApp.sendEmail(para.join(','), asunto, cuerpo);
    return salida_({ ok: true });
  } catch (err) {
    return salida_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// Marca como devuelto el último "Recogido" abierto de ese carrito.
function cerrarRecogido_(h, carrito, ahora) {
  const n = h.getLastRow() - 1;
  if (n < 1) return;
  const datos = h.getRange(2, 1, n, ENCABEZADOS.length).getValues();
  for (let i = datos.length - 1; i >= 0; i--) {
    const r = datos[i];
    if (r[C_CARRITO - 1] === carrito && r[C_TIPO - 1] === 'Recogido' &&
        (r[C_ESTADO - 1] === FUERA || r[C_ESTADO - 1] === TARDE)) {
      h.getRange(i + 2, C_ESTADO).setValue('Devuelto ' + fecha_(ahora));
      h.getRange(i + 2, 1, 1, ENCABEZADOS.length).setBackground(VERDE);
      return;
    }
  }
}

// Se ejecuta cada hora: pinta de rojo y avisa los carritos con más de 24 h fuera.
function revisarPendientes() {
  const ss = SpreadsheetApp.getActive();
  const h = hojaAvisos_(ss);
  const n = h.getLastRow() - 1;
  if (n < 1) return;
  const datos = h.getRange(2, 1, n, ENCABEZADOS.length).getValues();
  const ahora = new Date();
  const tarde = [];
  datos.forEach((r, i) => {
    const marca = r[C_MARCA - 1];
    if (r[C_TIPO - 1] !== 'Recogido' || r[C_ESTADO - 1] !== FUERA || !(marca instanceof Date)) return;
    const horas = (ahora - marca) / 36e5;
    if (horas < HORAS_LIMITE) return;
    h.getRange(i + 2, C_ESTADO).setValue(TARDE);
    h.getRange(i + 2, 1, 1, ENCABEZADOS.length).setBackground(ROJO);
    tarde.push('• ' + r[1] + ': ' + r[3] + (r[4] ? ' (' + r[4] + ')' : '') +
      '. Lo recogió el ' + r[0] + ', hace ' + Math.floor(horas) + ' horas.');
  });
  const para = correos_(ss);
  if (tarde.length && para.length) {
    MailApp.sendEmail(para.join(','),
      '⚠️ ' + (tarde.length === 1 ? 'Un carrito lleva' : tarde.length + ' carritos llevan') +
      ' más de 24 horas sin devolver',
      'No se ha registrado la devolución de:\n\n' + tarde.join('\n') +
      '\n\nConfirmen con la persona. La fila está en rojo en la hoja "Avisos de carritos".');
  }
}

function fecha_(d) {
  const meses = ['ene', 'feb', 'mar', 'abr', 'may', 'jun', 'jul', 'ago', 'sep', 'oct', 'nov', 'dic'];
  const f = k => Utilities.formatDate(d, ZONA, k);
  const ampm = f('a') === 'AM' ? 'a. m.' : 'p. m.';
  return f('d') + ' ' + meses[Number(f('M')) - 1] + ' ' + f('yyyy') + ', ' + f('h:mm') + ' ' + ampm;
}

function hojaAvisos_(ss) {
  let h = ss.getSheetByName('Avisos');
  if (!h) {
    h = ss.insertSheet('Avisos', 0);
    h.appendRow(ENCABEZADOS);
    h.setFrozenRows(1);
  } else if (h.getRange(1, C_MARCA).getValue() !== ENCABEZADOS[C_MARCA - 1]) {
    h.getRange(1, 1, 1, ENCABEZADOS.length).setValues([ENCABEZADOS]);
  }
  h.getRange(1, 1, 1, ENCABEZADOS.length).setFontWeight('bold');
  h.hideColumns(C_MARCA);
  return h;
}

function correos_(ss) {
  let h = ss.getSheetByName('Correos');
  if (!h) {
    h = ss.insertSheet('Correos');
    h.appendRow(['Correos que reciben los avisos (uno por fila)']);
    h.appendRow([Session.getEffectiveUser().getEmail()]);
  }
  return h.getRange(1, 1, h.getLastRow(), 1).getValues().flat()
    .map(c => String(c).trim())
    .filter(c => /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(c));
}

function t_(v, max) {
  const s = String(v == null ? '' : v).replace(/\s+/g, ' ').trim().slice(0, max);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function salida_(o) {
  return ContentService.createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}
