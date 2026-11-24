/*  GRAN PRIX SANTIAGO #4 — v3
 *  (ranking en vivo + familias únicas + ingreso con código desde otro celular)
 *
 *  CÓMO ACTUALIZAR:
 *  1. Pega TODO esto en tu Apps Script (reemplaza lo anterior) y guarda (💾).
 *  2. IMPORTANTE: en la barra de arriba elige la función  probarAviso  y presiona
 *     ▶ Ejecutar. Google te pedirá permisos: "Revisar permisos" → tu cuenta →
 *     "Configuración avanzada" → "Ir a ... (no seguro)" → Permitir.
 *     (Llega un correo de prueba a los dos correos.)
 *  3. Implementar → Gestionar implementaciones → lápiz (editar) →
 *     Versión: "Nueva versión" → Implementar.   (La URL /exec NO cambia.)
 *  4. Para borrar los datos de prueba: elige la función  limpiarTodo  y ▶ Ejecutar.
 */

var CORREOS = ['wil1509@live.com', 'Paozuso2282@gmail.com'];

/* WhatsApp en tiempo real (gratis con CallMeBot, a TU propio número).
   Cuando CallMeBot te responda "Your APIKEY is 123456", pega ese número aquí. */
var WA_TEL    = '+573182840198';
var WA_APIKEY = '';            // ej: '123456'  (vacío = WhatsApp apagado, solo correo)
var TZ = 'America/Bogota';

function instalar() {
  hoja_('Inscritos', ['Fecha/Hora', 'Familia', 'WhatsApp', 'Cupos', 'Asistencia', 'Mensaje', 'N° Largada']);
  hoja_('Puntajes',  ['Fecha/Hora', 'Juego', 'Familia', 'N° Largada', 'Puntaje/Tiempo']);
  var ya = ScriptApp.getProjectTriggers();
  for (var i = 0; i < ya.length; i++) {
    if (ya[i].getHandlerFunction() === 'resumenDiario') ScriptApp.deleteTrigger(ya[i]);
  }
  ScriptApp.newTrigger('resumenDiario').timeBased().everyDays(1).atHour(20).inTimezone(TZ).create();
  SpreadsheetApp.getActiveSpreadsheet().toast('Instalado. Vuelve a desplegar como aplicación web.');
}

/* Borra TODAS las filas de datos (deja los títulos). Úsalo para quitar las pruebas. */
function limpiarTodo() {
  ['Inscritos', 'Puntajes'].forEach(function (n) {
    var sh = hoja_(n), last = sh.getLastRow();
    if (last > 1) sh.deleteRows(2, last - 1);
  });
  try { SpreadsheetApp.getActiveSpreadsheet().toast('Datos de prueba borrados. ¡Listo para la fiesta!'); } catch (e) {}
}

/* ---------- Utilidades de nombres y códigos ---------- */
function norm_(s) {
  s = String(s || '').toLowerCase();
  if (s.normalize) s = s.normalize('NFD').replace(/[̀-ͯ]/g, '');
  s = s.replace(/[^a-z0-9ñ ]/g, ' ').replace(/\s+/g, ' ').trim();
  s = s.replace(/^(la )?(familia|flia|fam)\.? /, '');
  return s.trim();
}
function digitos_(s) { return String(s || '').replace(/\D/g, ''); }
function ult10_(s) { var d = digitos_(s); return d.length > 10 ? d.slice(-10) : d; }

function inscritos_() {
  return datos_('Inscritos').filter(function (r) { return String(r[1] || '').trim(); });
}
function buscarFamilia_(fam) {
  var k = norm_(fam);
  if (!k) return null;
  var rows = inscritos_();
  for (var i = 0; i < rows.length; i++) if (norm_(rows[i][1]) === k) return rows[i];
  return null;
}
function largadaUsada_(cod) {
  cod = digitos_(cod);
  if (!cod) return false;
  return inscritos_().some(function (r) { return digitos_(r[6]) === cod; });
}

/* ---------- Escrituras ---------- */
function doPost(e) {
  try {
    var p = (e && e.parameter) ? e.parameter : {};
    var ahora = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm');
    if (p.tipo === 'inscripcion') {
      if (buscarFamilia_(p.familia)) return ContentService.createTextOutput('duplicado');
      var fila = [ahora, p.familia || '', p.whatsapp || '', p.cupos || '', p.asistencia || '', p.mensaje || '', p.largada || ''];
      hoja_('Inscritos').appendRow(fila);
      try { avisoInscripcion_(fila); } catch (e) {}   // nunca bloquea el registro
    } else if (p.tipo === 'puntaje') {
      // usa el nombre oficial de la familia si existe
      var f = buscarFamilia_(p.familia);
      hoja_('Puntajes').appendRow([ahora, p.juego || '', f ? f[1] : (p.familia || ''), p.largada || '', p.puntaje || p.tiempo || '']);
    }
    return ContentService.createTextOutput('ok');
  } catch (err) {
    return ContentService.createTextOutput('error');
  }
}

/* ---------- Lecturas (JSONP) ---------- */
function doGet(e) {
  var p = (e && e.parameter) ? e.parameter : {};
  var cb = p.callback || '';
  var res = {};
  try {
    if (p.accion === 'existe') {
      // ¿ya hay una familia con ese apellido?  ¿el número propuesto ya está tomado?
      var f = buscarFamilia_(p.fam);
      res = { existe: !!f, familia: f ? f[1] : '', largadaUsada: largadaUsada_(p.cod) };
    } else if (p.accion === 'validar') {
      // ingreso desde otro celular con #código o número de celular
      var cod = digitos_(p.cod), hit = null, rows = inscritos_();
      if (cod.length >= 7) {
        var c10 = ult10_(cod);
        for (var i = 0; i < rows.length; i++) if (ult10_(rows[i][2]) && ult10_(rows[i][2]) === c10) { hit = rows[i]; break; }
      } else if (cod) {
        for (var j = 0; j < rows.length; j++) if (digitos_(rows[j][6]) === cod) { hit = rows[j]; break; }
      }
      res = hit ? { ok: true, familia: hit[1], largada: String(hit[6]) } : { ok: false };
    } else {
      res = { top: top5_() };
    }
  } catch (err) { res = { error: true, top: [] }; }
  var payload = JSON.stringify(res);
  if (cb) return ContentService.createTextOutput(cb + '(' + payload + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(payload).setMimeType(ContentService.MimeType.JSON);
}

/* Top 5 familias (mejor puntaje de carrera). Une variantes del mismo nombre. */
function top5_() {
  var oficial = {};
  inscritos_().forEach(function (r) { var k = norm_(r[1]); if (k && !oficial[k]) oficial[k] = r[1]; });
  var mejor = {};
  datos_('Puntajes').forEach(function (r) {
    if (String(r[1]).toLowerCase().indexOf('carrera') < 0) return;
    var k = norm_(r[2]); if (!k) return;
    var pts = parseInt(r[4], 10) || 0;
    if (!mejor[k] || pts > mejor[k].puntaje) mejor[k] = { familia: oficial[k] || r[2], puntaje: pts };
  });
  return Object.keys(mejor).map(function (k) { return mejor[k]; })
    .sort(function (a, b) { return b.puntaje - a.puntaje; }).slice(0, 5);
}

/* ---------- Aviso INMEDIATO por cada familia inscrita ---------- */
function avisoInscripcion_(r) {
  var ins = inscritos_();
  var va = r[4] === 'no' ? '❌ NO asistirá' : '✅ Asistirá';
  var lista = ins.map(function (x, i) {
    return (i + 1) + '. ' + x[1] + ' — ' + (String(x[4]) === 'no' ? 'no asiste' : (x[3] + ' cupos'));
  });
  var c = [];
  c.push('🏁 ¡Nueva familia inscrita en el Gran Prix de Santiago!'); c.push('');
  c.push('Familia: ' + r[1]);
  c.push('Asistencia: ' + va);
  c.push('Cupos: ' + (r[3] || '-'));
  c.push('WhatsApp: ' + (r[2] || 'no lo dejó'));
  c.push('N° de largada: #' + (r[6] || '-'));
  if (r[5]) c.push('Mensaje: ' + r[5]);
  c.push('');
  c.push('📋 FAMILIAS INSCRITAS AL MOMENTO (' + ins.length + ') · asistentes confirmados: ' + totalAsistentes_(ins));
  c = c.concat(lista);
  try {
    MailApp.sendEmail({ to: CORREOS.join(','), subject: '🏁 Nueva inscripción: ' + r[1] + ' — van ' + ins.length + ' familias', body: c.join('\n') });
  } catch (e) {}
  whatsapp_(c.join('\n'));
}
function whatsapp_(texto) {
  if (!WA_APIKEY) return;
  try {
    if (texto.length > 1500) texto = texto.slice(0, 1480) + '\n…(ver correo)';
    var url = 'https://api.callmebot.com/whatsapp.php?phone=' + encodeURIComponent(WA_TEL) +
              '&apikey=' + encodeURIComponent(WA_APIKEY) + '&text=' + encodeURIComponent(texto);
    UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  } catch (e) {}
}
function totalAsistentes_(ins) {
  return ins.reduce(function (t, r) { return t + (String(r[4]) === 'no' ? 0 : (parseInt(r[3], 10) || 0)); }, 0);
}
/* Ejecútala una vez para probar que el aviso llega a los correos */
function probarAviso() {
  avisoInscripcion_([Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'), 'Familia PRUEBA (ignorar)', '3000000000', '2', 'si', 'Esto es una prueba', '00000']);
}

/* ---------- Correo diario 8:00 PM (con TODAS las familias) ---------- */
function resumenDiario() {
  var hoy = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  var ins = inscritos_(), pun = datos_('Puntajes');
  var insHoy = ins.filter(function (r) { return String(r[0]).indexOf(hoy) === 0; });
  var punHoy = pun.filter(function (r) { return String(r[0]).indexOf(hoy) === 0; });
  var noVan = ins.filter(function (r) { return String(r[4]) === 'no'; }).length;
  var podio = top5_();
  var c = [];
  c.push('RESUMEN DEL DÍA — Gran Prix Santiago #4'); c.push(hoy); c.push('');
  c.push('Familias inscritas (total): ' + ins.length + '   ·   nuevas hoy: ' + insHoy.length);
  c.push('Asistentes confirmados (suma de cupos): ' + totalAsistentes_(ins));
  c.push('Familias que NO asistirán: ' + noVan);
  c.push('Partidas/puntajes hoy: ' + punHoy.length); c.push('');
  c.push('PODIO PRELIMINAR (mejor puntaje por familia):');
  if (podio.length === 0) c.push('  (aún sin puntajes)');
  podio.forEach(function (f, i) { c.push('  ' + (i + 1) + '. ' + f.familia + ' — ' + f.puntaje + ' pts'); });
  c.push(''); c.push('TODAS LAS FAMILIAS INSCRITAS:');
  if (ins.length === 0) c.push('  (aún ninguna)');
  ins.forEach(function (r, i) {
    var nuevo = String(r[0]).indexOf(hoy) === 0 ? '  🆕' : '';
    c.push('  ' + (i + 1) + '. ' + r[1] + ' — ' + (String(r[4]) === 'no' ? 'NO asiste' : (r[3] + ' cupos')) + ' — ' + (r[2] || 'sin WhatsApp') + ' — #' + r[6] + nuevo);
  });
  MailApp.sendEmail({ to: CORREOS.join(','), subject: 'Gran Prix Santiago #4 — Resumen ' + hoy + ' (' + ins.length + ' familias)', body: c.join('\n') });
}

function hoja_(nombre, encabezados) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(nombre);
  if (!sh) { sh = ss.insertSheet(nombre); if (encabezados) { sh.appendRow(encabezados); sh.getRange(1, 1, 1, encabezados.length).setFontWeight('bold'); } }
  return sh;
}
function datos_(nombre) {
  var sh = hoja_(nombre);
  var v = sh.getDataRange().getValues();
  return v.slice(1);
}
