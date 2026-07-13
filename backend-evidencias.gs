/* =========================================================================
 * EVIDENCIAS DEL EXPEDIENTE · Backend Apps Script
 * Sistema de Investigación de Casos · RL Intelligence
 * -------------------------------------------------------------------------
 * INSTALACIÓN (una sola vez):
 * 1. Abre script.google.com → el proyecto del Sistema de Investigación.
 * 2. Crea un archivo nuevo llamado "evidencias" y pega TODO este código.
 * 3. En tu doPost, donde enrutas las acciones autenticadas (después de
 *    validar el token, junto a 'crearCaso', 'analizarCaso', etc.), agrega:
 *
 *      if (accion === 'subirEvidencia')   return respuestaOk(subirEvidencia(body));
 *      if (accion === 'listarEvidencias') return respuestaOk(listarEvidencias(body));
 *
 *    (usa el MISMO helper de respuesta que tus otras acciones: el que
 *     devuelve {ok:true, data:...})
 * 4. Implementar → Administrar implementaciones → ✏️ → Nueva versión → Implementar.
 *
 * Los archivos se guardan en la carpeta de Drive del expediente (la busca
 * por el código; si no la encuentra, crea RL_Evidencias/<código>).
 * Se registran en la hoja "Evidencias" del mismo spreadsheet (se crea sola).
 * ========================================================================= */

var EV_MAX_BYTES = 15 * 1024 * 1024; // 15 MB

function subirEvidencia(d) {
  if (!d.codigo || !d.base64 || !d.nombre) throw new Error('Faltan datos (codigo, nombre, base64).');
  var bytes = Utilities.base64Decode(d.base64);
  if (bytes.length > EV_MAX_BYTES) throw new Error('El archivo supera 15 MB.');
  var mime = d.mime || 'application/octet-stream';
  if (!/^image\/|^application\/pdf$|word|officedocument/.test(mime)) {
    throw new Error('Solo se permiten imágenes, PDF o Word.');
  }

  var carpeta = _carpetaDelCaso_(String(d.codigo).trim());
  var nombre = String(d.nombre).replace(/[\\/:*?"<>|]/g, '_');
  var file = carpeta.createFile(Utilities.newBlob(bytes, mime, nombre));
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  var sh = _hojaEvidencias_();
  sh.appendRow([
    String(d.codigo).trim(),
    nombre,
    file.getUrl(),
    mime,
    d.usuario || '',
    Utilities.formatDate(new Date(), 'America/Lima', 'dd/MM/yyyy HH:mm')
  ]);
  return { nombre: nombre, url: file.getUrl() };
}

function listarEvidencias(d) {
  if (!d.codigo) throw new Error('Falta el código del expediente.');
  var sh = _hojaEvidencias_();
  var last = sh.getLastRow();
  if (last < 2) return [];
  var vals = sh.getRange(2, 1, last - 1, 6).getValues();
  var cod = String(d.codigo).trim();
  var out = [];
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === cod) {
      out.push({ nombre: vals[i][1], url: vals[i][2], mime: vals[i][3], subido_por: vals[i][4], fecha: vals[i][5] });
    }
  }
  return out;
}

/* ---------- helpers ---------- */
function _hojaEvidencias_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('Evidencias');
  if (!sh) {
    sh = ss.insertSheet('Evidencias');
    sh.appendRow(['codigo', 'nombre', 'url', 'mime', 'subido_por', 'fecha']);
    sh.getRange(1, 1, 1, 6).setFontWeight('bold');
  }
  return sh;
}

// Busca la carpeta de Drive del expediente por su código (la que crea crearCaso).
// Si no existe, crea RL_Evidencias/<código>.
function _carpetaDelCaso_(codigo) {
  var it = DriveApp.searchFolders('title contains "' + codigo + '"');
  if (it.hasNext()) return it.next();
  var rootIt = DriveApp.getRootFolder().getFoldersByName('RL_Evidencias');
  var root = rootIt.hasNext() ? rootIt.next() : DriveApp.getRootFolder().createFolder('RL_Evidencias');
  var subIt = root.getFoldersByName(codigo);
  return subIt.hasNext() ? subIt.next() : root.createFolder(codigo);
}
