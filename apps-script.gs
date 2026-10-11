// Carritos de publicaciones: recibe los avisos de la página, los guarda en esta hoja
// y envía un correo a las direcciones de la pestaña "Correos".
const ZONA = 'America/Puerto_Rico';
const ENCABEZADOS = ['Fecha y hora', 'Carrito', 'Tipo', 'Nombre', 'Localidad', 'Revisión',
  'Pendientes en la revisión', 'Faltan publicaciones', 'Problema', 'Detalles'];
const VERBO = { Recogido: 'recogido por', Devuelto: 'devuelto por', Problema: 'problema reportado por' };

function preparar() {
  const ss = SpreadsheetApp.getActive();
  hojaAvisos_(ss);
  correos_(ss);
  MailApp.sendEmail(Session.getEffectiveUser().getEmail(), 'Carritos: listo',
    'La hoja de avisos de los carritos ya está preparada.');
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
    const fila = [
      Utilities.formatDate(new Date(), ZONA, 'yyyy-MM-dd HH:mm'),
      carrito, tipo, nombre,
      t_(d.localidad, 120), t_(d.revision, 300), t_(d.pendientes, 300),
      t_(d.faltan, 300), t_(d.problema, 80), t_(d.detalles, 1000)
    ];
    const ss = SpreadsheetApp.getActive();
    hojaAvisos_(ss).appendRow(fila);

    const alerta = tipo === 'Problema' || fila[6] || fila[7];
    const asunto = (alerta ? '⚠️ ' : '') + carrito + ': ' + VERBO[tipo] + ' ' + nombre;
    const cuerpo = ENCABEZADOS.map((h, i) => (fila[i] ? h + ': ' + fila[i] : ''))
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

function hojaAvisos_(ss) {
  let h = ss.getSheetByName('Avisos');
  if (!h) {
    h = ss.insertSheet('Avisos', 0);
    h.appendRow(ENCABEZADOS);
    h.setFrozenRows(1);
    h.getRange(1, 1, 1, ENCABEZADOS.length).setFontWeight('bold');
  }
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
