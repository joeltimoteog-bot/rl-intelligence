// ============ EVIDENCIAS DEL EXPEDIENTE (v3.1) ============
// Complemento para RL INTELLIGENCE AI - BACKEND v3.0
// Sube capturas/declaraciones/informes a la carpeta Drive DEL CASO
// (caso.drive_folder_id) y las registra en la hoja 'Documentos' existente
// con tipo 'EVIDENCIA'. listarEvidencias devuelve TODOS los documentos del
// caso (evidencias + informes ya analizados).
//
// INSTALACIÓN:
// 1) Pega este archivo completo en tu proyecto GAS (archivo nuevo "evidencias").
// 2) En doPost, dentro del switch, agrega estas 2 líneas junto a las demás:
//
//    case 'subirEvidencia':   _checkCaso(ses, req.codigo); out = subirEvidencia(req, quien); break;
//    case 'listarEvidencias': _checkCaso(ses, req.codigo); out = listarEvidencias(req.codigo); break;
//
// 3) Implementar → Administrar implementaciones → ✏️ → Nueva versión.

const EV_MAX_BYTES = 15 * 1024 * 1024; // 15 MB

function subirEvidencia(req, quien) {
  if (!req.codigo || !req.base64 || !req.nombre) throw new Error('Faltan datos (codigo, nombre, base64)');
  const caso = _buscarCaso(req.codigo);
  if (!caso) throw new Error('Caso no encontrado: ' + req.codigo);

  const bytes = Utilities.base64Decode(req.base64);
  if (bytes.length > EV_MAX_BYTES) throw new Error('El archivo supera 15 MB');
  const mime = req.mime || 'application/octet-stream';
  if (!/^image\/|^application\/pdf$|word|officedocument/.test(mime)) {
    throw new Error('Solo se permiten imágenes, PDF o Word');
  }

  const nombre = String(req.nombre).replace(/[\\/:*?"<>|]/g, '_');
  const blob = Utilities.newBlob(bytes, mime, nombre);
  const archivo = DriveApp.getFolderById(caso.drive_folder_id).createFile(blob);

  _hoja('Documentos').appendRow([_id(), req.codigo, 'EVIDENCIA', nombre,
    archivo.getId(), _ahora(), quien || 'web', '', 'PENDIENTE']);
  _auditar(quien, 'Documento', req.codigo, 'EVIDENCIA', nombre);
  return { nombre: nombre, url: 'https://drive.google.com/file/d/' + archivo.getId() + '/view' };
}

function listarEvidencias(codigo) {
  const datos = _hoja('Documentos').getDataRange().getValues();
  const out = [];
  for (let i = datos.length - 1; i >= 1; i--) {
    if (String(datos[i][1]) !== String(codigo)) continue;
    const nombre = String(datos[i][3] || '');
    out.push({
      tipo: datos[i][2] || 'DOCUMENTO',
      nombre: nombre,
      url: 'https://drive.google.com/file/d/' + datos[i][4] + '/view',
      mime: _mimePorNombre(nombre),
      subido_por: datos[i][6] || '',
      fecha: datos[i][5] || ''
    });
  }
  return out;
}

function _mimePorNombre(n) {
  n = String(n).toLowerCase();
  if (/\.pdf$/.test(n)) return 'application/pdf';
  if (/\.(doc|docx)$/.test(n)) return 'application/msword';
  if (/\.(png|jpe?g|gif|webp|heic)$/.test(n)) return 'image/*';
  return '';
}
