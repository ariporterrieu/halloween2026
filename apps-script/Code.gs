/**
 * Invitación Halloween · Registro anónimo de disfraces
 *
 * Guarda las temáticas en la planilla "Halloween 2026 - Disfraces" (ID abajo). Usa la hoja "Disfraces":
 *   Columna A: Temática   Columna B: Fecha de registro
 *
 * GET  → devuelve la lista:            { ok: true, disfraces: [{ tema }] }
 * POST → registra { tema } y devuelve: { ok: true|false, error?, disfraces }
 *        Si la temática ya existe (sin importar mayúsculas, tildes o signos) responde error "tomado".
 */

const PLANILLA_ID = '1-CvVaPrSKIA1iI0LXtpesWniQalxwhkj3HunBMMHqZY';
const HOJA = 'Disfraces';
const MAX_LARGO = 60;

function doGet() {
  return responder({ ok: true, disfraces: leerDisfraces() });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000); // evita que dos personas registren lo mismo al mismo tiempo
  try {
    const datos = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const tema = String(datos.tema || '').replace(/\s+/g, ' ').trim().slice(0, MAX_LARGO);
    if (!tema) return responder({ ok: false, error: 'vacio', disfraces: leerDisfraces() });

    const actuales = leerDisfraces();
    const clave = normalizar(tema);
    if (actuales.some(function (d) { return normalizar(d.tema) === clave; })) {
      return responder({ ok: false, error: 'tomado', disfraces: actuales });
    }

    // El apóstrofo inicial evita que Sheets interprete el texto como fórmula
    hoja().appendRow([(/^[=+\-@]/.test(tema) ? "'" : '') + tema, new Date()]);
    actuales.unshift({ tema: tema });
    return responder({ ok: true, disfraces: actuales });
  } catch (err) {
    return responder({ ok: false, error: 'servidor', disfraces: leerDisfraces() });
  } finally {
    lock.releaseLock();
  }
}

/** Lista de temáticas, la más reciente primero. */
function leerDisfraces() {
  const h = hoja();
  const filas = h.getLastRow() - 1;
  if (filas < 1) return [];
  return h.getRange(2, 1, filas, 1).getDisplayValues()
    .map(function (f) { return String(f[0]).trim(); })
    .filter(String)
    .reverse()
    .map(function (t) { return { tema: t }; });
}

function hoja() {
  const libro = SpreadsheetApp.openById(PLANILLA_ID);
  let h = libro.getSheetByName(HOJA);
  if (!h) {
    h = libro.insertSheet(HOJA);
    h.appendRow(['Temática', 'Registrado']);
    h.setFrozenRows(1);
  }
  return h;
}

function normalizar(t) {
  return String(t).toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '')
    .replace(/[^a-z0-9ñ ]/g, ' ').replace(/\s+/g, ' ').trim();
}

function responder(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
