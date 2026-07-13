// ============ RL INTELLIGENCE AI - BACKEND v3.1 ============
// v3.1: TODO lo de v3.0 + EVIDENCIAS del expediente (subir capturas,
// declaraciones e informes a la carpeta Drive del caso, con registro
// en la hoja Documentos y auditoría).
// v3.0: empresas configurables, usuarios con periodo de prueba, instalador
// para venta (setupCompleto), extracción de reglamentos por empresa desde
// la web, caché de casos, sesiones con roles. Todo lo anterior integrado.

const CONFIG = {
  SPREADSHEET_ID: '1n_9OwSVfhGK4A5vfjrN09ppmnWPKxVdFwJysg9YI1oM', // fallback (instancia de Joel)
  MODELO: 'gemini-2.5-flash'
};

function _ss() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID') || CONFIG.SPREADSHEET_ID;
  return SpreadsheetApp.openById(id);
}
function _hoja(n) { return _ss().getSheetByName(n); }
function _ahora() { return Utilities.formatDate(new Date(), 'America/Lima', 'yyyy-MM-dd HH:mm'); }
function _hoy() { return Utilities.formatDate(new Date(), 'America/Lima', 'yyyy-MM-dd'); }
function _id() { return Utilities.getUuid().slice(0, 8); }
function _emailAlertas() { return Session.getActiveUser().getEmail(); }

// ============ AUDITORÍA ============
function _auditar(usuario, entidad, entidadId, accion, detalle) {
  _hoja('Auditoria').appendRow([_id(), _ahora(), usuario || 'sistema', entidad, entidadId, accion, detalle || '']);
}

// ============ EMPRESAS (configurables) ============
function _hojaEmpresas() {
  let h = _ss().getSheetByName('Empresas');
  if (!h) {
    h = _ss().insertSheet('Empresas');
    h.appendRow(['nombre', 'prefijo', 'activo']);
    h.setFrozenRows(1);
  }
  return h;
}

function listarEmpresas() {
  const cache = CacheService.getScriptCache();
  const c = cache.get('empresas');
  if (c) return JSON.parse(c);
  const d = _hojaEmpresas().getDataRange().getValues();
  const out = [];
  for (let i = 1; i < d.length; i++)
    if (String(d[i][2]).toUpperCase() === 'TRUE')
      out.push({ nombre: String(d[i][0]).toUpperCase(), prefijo: String(d[i][1]).toUpperCase() });
  cache.put('empresas', JSON.stringify(out), 300);
  return out;
}

function agregarEmpresa(nombre, prefijo) {
  if (!nombre || !prefijo) throw new Error('Completa nombre y prefijo de la empresa');
  nombre = String(nombre).toUpperCase().trim();
  prefijo = String(prefijo).toUpperCase().trim().slice(0, 4);
  const d = _hojaEmpresas().getDataRange().getValues();
  for (let i = 1; i < d.length; i++)
    if (String(d[i][0]).toUpperCase() === nombre) throw new Error('La empresa ya existe');
  _hojaEmpresas().appendRow([nombre, prefijo, 'TRUE']);
  CacheService.getScriptCache().remove('empresas');
  return { nombre: nombre, prefijo: prefijo };
}

function _prefijoEmpresa(nombre) {
  const es = listarEmpresas();
  for (let i = 0; i < es.length; i++)
    if (es[i].nombre === String(nombre).toUpperCase()) return es[i].prefijo;
  return String(nombre || 'GEN').toUpperCase().slice(0, 3);
}

// ============ CÓDIGO DE EXPEDIENTE ============
function _siguienteCodigo(empresa) {
  const base = _prefijoEmpresa(empresa) + '-' + new Date().getFullYear() + '-';
  const datos = _hoja('Casos').getDataRange().getValues();
  let max = 0;
  for (let i = 1; i < datos.length; i++) {
    const c = String(datos[i][0]);
    if (c.indexOf(base) === 0) {
      const n = parseInt(c.slice(base.length), 10);
      if (n > max) max = n;
    }
  }
  return base + ('0000' + (max + 1)).slice(-4);
}

// ============ CREAR CASO ============
function crearCaso(d) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000); // evita códigos duplicados con varios usuarios a la vez
  try {
    const codigo = _siguienteCodigo(d.empresa || (listarEmpresas()[0] || {}).nombre || 'GEN');
    const carpeta = DriveApp.getFolderById(_carpetaCasosId()).createFolder(codigo);
    _hoja('Casos').appendRow([
      codigo, 'ABIERTO', d.empresa || '', d.regimen || 'agrario', d.trabajador || '', d.dni || '',
      d.cargo || '', d.sector || '', d.ruta || '', d.supervisor || '', d.administrador || '',
      d.tipo_falta || '', d.fecha_hechos || '', d.lugar_hechos || '', d.descripcion_corta || '',
      '', '', _ahora(), '', d.creado_por || 'web', carpeta.getId()
    ]);
    _auditar(d.creado_por, 'Caso', codigo, 'CREAR', d.trabajador || '');
    CacheService.getScriptCache().remove('casos_recientes');
    _alertaCorreo('🆕 Nuevo caso ' + codigo,
      'Empresa: ' + d.empresa + '\nTrabajador: ' + (d.trabajador || '-') + '\nFalta: ' + (d.tipo_falta || '-') +
      '\nDescripción: ' + (d.descripcion_corta || '-') + '\n\nCarpeta: ' + carpeta.getUrl());
    return { codigo: codigo, drive_folder_id: carpeta.getId() };
  } finally {
    lock.releaseLock();
  }
}

function _carpetaCasosId() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('CARPETA_CASOS');
  if (!id) {
    const f = DriveApp.createFolder('RL_INTELLIGENCE_CASOS');
    id = f.getId();
    props.setProperty('CARPETA_CASOS', id);
  }
  return id;
}

// ============ REGLAMENTOS ============
function _hojaReglamentos() {
  let h = _ss().getSheetByName('Reglamentos');
  if (!h) {
    h = _ss().insertSheet('Reglamentos');
    h.appendRow(['clave', 'parte', 'texto']);
    h.setFrozenRows(1);
  }
  return h;
}

// Extrae (o re-extrae) el RIT y RISST de UNA empresa desde Drive.
// Los PDF deben llamarse RIT-NOMBRE y RISST-NOMBRE (ej. RIT-RAPEL).
function extraerReglamentosEmpresa(nombreEmpresa) {
  nombreEmpresa = String(nombreEmpresa).toUpperCase().trim();
  const hoja = _hojaReglamentos();
  const claves = ['RIT-' + nombreEmpresa, 'RISST-' + nombreEmpresa];
  const resultado = [];

  // conservar filas de otras claves
  const datos = hoja.getDataRange().getValues();
  const conservar = [datos[0]];
  for (let i = 1; i < datos.length; i++)
    if (claves.indexOf(datos[i][0]) === -1) conservar.push(datos[i]);
  hoja.clearContents();
  conservar.forEach(function(f) { hoja.appendRow(f); });

  claves.forEach(function(clave) {
    const archivos = DriveApp.searchFiles('title contains "' + clave + '"');
    if (!archivos.hasNext()) { resultado.push('⚠️ No encontrado en Drive: ' + clave); return; }
    const pdf = archivos.next();
    const doc = Drive.Files.create(
      { name: 'tmp_' + clave, mimeType: 'application/vnd.google-apps.document' },
      pdf.getBlob(), { ocrLanguage: 'es' }
    );
    const texto = DocumentApp.openById(doc.id).getBody().getText();
    Drive.Files.remove(doc.id);
    for (let i = 0; i < texto.length; i += 45000)
      hoja.appendRow([clave, Math.floor(i / 45000) + 1, texto.slice(i, i + 45000)]);
    resultado.push('✅ ' + clave + ': ' + texto.length + ' caracteres');
  });
  return { resultado: resultado };
}

// Legacy: extrae los 4 de Joel de una vez (sigue funcionando)
function extraerReglamentos() {
  ['RAPEL', 'VERFRUT'].forEach(function(e) {
    Logger.log(JSON.stringify(extraerReglamentosEmpresa(e)));
  });
}

function _reglamento(empresa, tipo) { // tipo: 'RIT' o 'RISST'
  const clave = tipo + '-' + String(empresa).toUpperCase();
  const datos = _hojaReglamentos().getDataRange().getValues();
  let texto = '';
  for (let i = 1; i < datos.length; i++) if (datos[i][0] === clave) texto += datos[i][2];
  return texto;
}

// ============ MARCO LEGAL ============
const MARCO_LEGAL = 'Eres un abogado laboralista peruano senior que asesora al área de Relaciones Laborales de la empresa del usuario. Dominas y aplicas SIEMPRE este marco actualizado:\n' +
'== REGÍMENES ==\n' +
'1) RÉGIMEN AGRARIO (Ley 31110 y D.S. 005-2021-MIDAGRI): RB no menor a RMV; BETA 30% RMV (no remunerativa); gratificaciones 16.66% RB; CTS 9.72% RB; vacaciones 30 días; indemnización por despido arbitrario = 45 remuneraciones diarias por año completo, tope 360 RD.\n' +
'2) RÉGIMEN GENERAL (D.S. 003-97-TR, TUO LPCL): indemnización por despido arbitrario = 1.5 remuneraciones mensuales por año, tope 12.\n' +
'La VALORACIÓN de responsabilidad, gravedad, prueba y debido proceso es IDÉNTICA en ambos regímenes; el régimen solo cambia la consecuencia económica.\n' +
'== PROCEDIMIENTO DISCIPLINARIO ==\n' +
'- Sanciones progresivas: amonestación verbal → amonestación escrita → suspensión sin goce → despido. Principios: legalidad, tipicidad, inmediatez, proporcionalidad, razonabilidad, non bis in idem.\n' +
'- Despido: SOLO por falta grave del art. 25 D.S. 003-97-TR y con debido proceso (preaviso con hechos + base legal + pruebas + plazo NO MENOR de 6 días naturales para descargos; 30 si es capacidad; excepción flagrancia). Despido sin causa o sin proceso = ARBITRARIO; despido NULO (embarazo, sindicalización, discriminación, represalia) = reposición.\n' +
'== SST (Ley 29783) ==: faltas de seguridad se tipifican contra el RISST; el empleador investiga accidentes e incidentes. Una falta de SST solo constituye falta grave laboral si por su entidad quebranta la buena fe o encaja en el art. 25.\n' +
'== HOSTIGAMIENTO (Ley 27942) ==: procedimiento especial con su comité y plazos; no lo resuelvas por la vía disciplinaria común, adviértelo.\n' +
'== TEST OBLIGATORIO ANTES DE SUGERIR CUALQUIER MEDIDA (aplícalo SIEMPRE, en orden, y muéstralo en fundamentos) ==\n' +
'1. TIPICIDAD ESTRICTA: cita TEXTUALMENTE cada inciso del RIT/RISST que invoques y verifica ELEMENTO POR ELEMENTO que los hechos encajen. Si un solo elemento no se cumple (ej.: el inciso exige "no encontrarse en labores" y el trabajador SÍ estaba laborando), el inciso NO aplica: descártalo expresamente y dilo. PROHIBIDO citar artículos cuyo texto no esté en el reglamento proporcionado.\n' +
'2. CULPABILIDAD: clasifica la conducta como DOLO (intención), CULPA GRAVE (negligencia inexcusable), CULPA LEVE (descuido) o ERROR/OLVIDO EXCUSABLE (equivocación humana de buena fe). Un olvido o error de buena fe, en primera falta, rara vez quebranta la buena fe laboral que exige el art. 25.\n' +
'3. DAÑO: distingue daño REAL consumado (personas afectadas, intoxicaciones, pérdida material) de riesgo POTENCIAL. El riesgo potencial AGRAVA pero NO equivale al daño consumado. Sin daño real y sin dolo → la medida debe ser moderada aunque el riesgo fuera serio.\n' +
'4. ANTECEDENTES: récord limpio → gradualidad (amonestación escrita o suspensión corta de 1 a 3 días según entidad de la falta); reincidencia en faltas similares ya sancionadas → escalamiento citando cada antecedente.\n' +
'5. RAZONABILIDAD: pondera atenuantes (confesión, colaboración, antigüedad sin sanciones, estaba cumpliendo sus labores) y agravantes. Ante duda entre dos medidas, sugiere SIEMPRE la MENOR, acompañada de advertencia formal y capacitación. El despido es ÚLTIMA RATIO: solo con tipicidad plena del art. 25 + dolo o culpa grave con daño relevante, o reincidencia acreditada.\n' +
'== ESCALA INTERNA DE GRAVEDAD (gestión) ==\n' +
'- LEVE: error/descuido sin daño, primera vez → amonestación verbal o escrita.\n' +
'- GRAVE: culpa con afectación real u operativa, o reincidencia de leves → amonestación severa o suspensión 1-3 días.\n' +
'- MUY GRAVE: encuadra plenamente en art. 25 → suspensión ejemplar o procedimiento de despido.\n' +
'- CRÍTICO: dolo, flagrancia, violencia, daño a la vida o integridad, daño consumado grave → acción inmediata + procedimiento de despido; evaluar denuncia penal.\n' +
'La escala es de GESTIÓN INTERNA; legalmente el despido solo procede si la conducta encaja en el art. 25: mapea siempre y dilo expresamente.\n' +
'== REGLAS DE TRABAJO ==\n' +
'- Funda todo en: (a) RIT o RISST de la empresa con citas textuales, (b) la ley del régimen del trabajador, (c) D.S. 003-97-TR para procedimiento.\n' +
'- MODO ASESOR: si falta información relevante (récord, versión del trabajador, testigos, si estaba o no en labores, magnitud del daño), FORMULA preguntas concretas antes de endurecer la medida. Toda medida es SUGERENCIA sujeta a validación humana.\n' +
'- Datos sensibles (Ley 29733): confidencialidad. Español formal, claro y accionable.';

// ============ MOTOR GEMINI ============
function _openai(messages, esTexto) { // esTexto=true → conversacional (no JSON)
  const key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!key) throw new Error('Falta GEMINI_API_KEY en Propiedades del script');
  const sistema = messages.filter(function(m){ return m.role === 'system'; })
                          .map(function(m){ return m.content; }).join('\n');
  const chat = messages.filter(function(m){ return m.role !== 'system'; })
                       .map(function(m){
                         return { role: m.role === 'assistant' ? 'model' : 'user',
                                  parts: [{ text: m.content }] };
                       });
  const cfg = { temperature: 0.2 };
  if (!esTexto) cfg.responseMimeType = 'application/json';
  const res = UrlFetchApp.fetch(
    'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=' + key, {
    method: 'post', contentType: 'application/json',
    payload: JSON.stringify({ systemInstruction: { parts: [{ text: sistema }] }, contents: chat, generationConfig: cfg }),
    muteHttpExceptions: true
  });
  const j = JSON.parse(res.getContentText());
  if (j.error) throw new Error('Gemini: ' + j.error.message);
  if (!j.candidates || !j.candidates.length) throw new Error('Gemini no devolvió respuesta');
  return j.candidates[0].content.parts[0].text;
}

// ============ ANÁLISIS JURÍDICO (Fase 2) ============
function analizarCaso(casoCodigo, textoInforme, usuario, antecedentes) {
  const caso = _buscarCaso(casoCodigo);
  if (!caso) throw new Error('Caso no encontrado: ' + casoCodigo);
  const esSST = /accidente|epp|seguridad|salud|incidente/i.test(caso.tipo_falta + ' ' + textoInforme.slice(0, 500));
  const reglamento = _reglamento(caso.empresa, esSST ? 'RISST' : 'RIT');
  const inv = obtenerInvestigacion(casoCodigo);

  const prompt = 'EMPRESA: ' + caso.empresa + ' | RÉGIMEN: ' + caso.regimen +
    '\n\nREGLAMENTO APLICABLE (' + (esSST ? 'RISST' : 'RIT') + '):\n' + reglamento.slice(0, 60000) +
    '\n\nINFORME A ANALIZAR:\n' + textoInforme +
    (inv ? '\n\nEXPEDIENTE YA INVESTIGADO (Fase 1 — úsalo, no re-extraigas): Hallazgos: ' + inv.hallazgos + ' | Auditoría: ' + inv.auditoria + ' | Confianza: ' + inv.nivel_confianza : '') +
    (antecedentes ? '\n\nRÉCORD DE SANCIONES / ANTECEDENTES DEL(LOS) TRABAJADOR(ES):\n' + antecedentes +
      '\nVALIDA la medida contra este récord: si hay reincidencia en faltas similares ya sancionadas, sustenta el escalamiento progresivo (amonestación → suspensión → despido) citando cada antecedente; si el récord está limpio o las faltas previas no guardan relación, aplica gradualidad y modera la medida. Indica expresamente cómo el récord influyó en la sugerencia.'
      : '\n\nNO se proporcionó récord de sanciones: asume primera falta salvo que el informe indique lo contrario, y aplica gradualidad.') +
    '\n\nAnaliza como abogado laboralista y responde SOLO un JSON válido con estas claves exactas: ' +
    'resumen_ejecutivo, cronologia, responsabilidades (analiza a CADA involucrado por separado: grado de responsabilidad, gravedad de su falta con artículo citado, medida individual), ' +
    'agravantes, atenuantes, riesgo_legal, fundamentos (artículos del reglamento y de la ley), conclusion, ' +
    'medida_recomendada (preséntala como SUGERENCIA sujeta a validación humana, indicando si el récord de sanciones la agrava o la modera), ' +
    'clasificacion_gravedad (LEVE, GRAVE, MUY GRAVE o CRITICO, con su equivalencia legal), ' +
    'nivel_riesgo (BAJO, MEDIO o ALTO), observaciones, ' +
    'preguntas_al_usuario (array con 2-4 preguntas concretas que un abogado te haría para afinar la medida; si no falta nada, array vacío).';

  const r = _openai([{ role: 'system', content: MARCO_LEGAL }, { role: 'user', content: prompt }]);
  const ev = JSON.parse(r.replace(/```json|```/g, '').trim());

  const version = _versionSiguiente('Evaluaciones', casoCodigo);
  _hoja('Evaluaciones').appendRow([
    _id(), casoCodigo, version, ev.resumen_ejecutivo, ev.cronologia, JSON.stringify(ev.responsabilidades),
    ev.agravantes, ev.atenuantes, ev.riesgo_legal, ev.fundamentos, ev.conclusion,
    ev.medida_recomendada, ev.nivel_riesgo, ev.observaciones, CONFIG.MODELO, _ahora(), usuario || 'sistema'
  ]);
  _actualizarCaso(casoCodigo, { nivel_riesgo: ev.nivel_riesgo, medida_recomendada: ev.medida_recomendada });
  CacheService.getScriptCache().remove('casos_recientes');
  _auditar(usuario, 'Evaluacion', casoCodigo, 'ANALIZAR', 'v' + version + ' riesgo ' + ev.nivel_riesgo + (antecedentes ? ' (con récord)' : ''));
  _alertaCorreo('⚖️ Análisis listo: ' + casoCodigo + ' — riesgo ' + ev.nivel_riesgo,
    ev.resumen_ejecutivo + '\n\nMedida sugerida: ' + ev.medida_recomendada);
  return ev;
}

// ============ HELPERS DE CASOS ============
function _buscarCaso(codigo) {
  const h = _hoja('Casos'), datos = h.getDataRange().getValues(), head = datos[0];
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0] === codigo) {
      const o = { _fila: i + 1 };
      head.forEach(function(c, j) { o[c] = datos[i][j]; });
      return o;
    }
  }
  return null;
}

function _actualizarCaso(codigo, cambios) {
  const h = _hoja('Casos'), caso = _buscarCaso(codigo);
  if (!caso) return;
  const head = h.getRange(1, 1, 1, h.getLastColumn()).getValues()[0];
  Object.keys(cambios).forEach(function(campo) {
    const col = head.indexOf(campo) + 1;
    if (col > 0) h.getRange(caso._fila, col).setValue(cambios[campo]);
  });
}

function _versionSiguiente(hoja, casoCodigo) {
  const datos = _hoja(hoja).getDataRange().getValues();
  let v = 0;
  for (let i = 1; i < datos.length; i++) if (datos[i][1] === casoCodigo && datos[i][2] > v) v = datos[i][2];
  return v + 1;
}

// ============ ALERTAS ============
function _alertaCorreo(asunto, cuerpo) {
  try { MailApp.sendEmail(_emailAlertas(), '[RL Intelligence] ' + asunto, cuerpo); } catch (e) {}
}

// ============ LISTAR CASOS (con caché) ============
function listarCasos(filtro, ses) {
  const alcance = _empresasDe(ses);
  const permitidas = alcance === 'TODAS' ? null :
    alcance.split(',').map(function(x){ return x.trim().toUpperCase(); });
  const cache = CacheService.getScriptCache();
  if (!filtro && !permitidas) {
    const c = cache.get('casos_recientes');
    if (c) return JSON.parse(c);
  }
  const datos = _hoja('Casos').getDataRange().getValues();
  const head = datos[0];
  const out = [];
  for (let i = datos.length - 1; i >= 1; i--) {
    const o = {};
    head.forEach(function(c, j) { o[c] = datos[i][j]; });
    if (permitidas && permitidas.indexOf(String(o.empresa).toUpperCase()) === -1) continue;
    if (filtro) {
      const q = String(filtro).toLowerCase();
      const blob = (o.codigo + ' ' + o.trabajador + ' ' + o.dni + ' ' + o.supervisor + ' ' + o.empresa + ' ' + o.sector + ' ' + o.ruta).toLowerCase();
      if (blob.indexOf(q) === -1) continue;
    }
    out.push(o);
    if (out.length >= 60) break;
  }
  if (!filtro && !permitidas) cache.put('casos_recientes', JSON.stringify(out), 60);
  return out;
}

function obtenerEvaluacion(casoCodigo) {
  const datos = _hoja('Evaluaciones').getDataRange().getValues();
  const head = datos[0];
  for (let i = datos.length - 1; i >= 1; i--) {
    if (datos[i][1] === casoCodigo) {
      const o = {};
      head.forEach(function(c, j) { o[c] = datos[i][j]; });
      return o;
    }
  }
  return null;
}

// ============ ANALIZAR DESDE ARCHIVO ============
function analizarArchivo(casoCodigo, nombreArchivo, base64, mimeType, usuario, antecedentes) {
  const caso = _buscarCaso(casoCodigo);
  if (!caso) throw new Error('Caso no encontrado: ' + casoCodigo);
  const blob = Utilities.newBlob(Utilities.base64Decode(base64), mimeType, nombreArchivo);
  const archivo = DriveApp.getFolderById(caso.drive_folder_id).createFile(blob);
  _hoja('Documentos').appendRow([_id(), casoCodigo, 'INFORME', nombreArchivo,
    archivo.getId(), _ahora(), usuario || 'web', '', 'ANALIZADO']);
  const doc = Drive.Files.create(
    { name: 'tmp_' + nombreArchivo, mimeType: 'application/vnd.google-apps.document' },
    blob, { ocrLanguage: 'es' }
  );
  const texto = DocumentApp.openById(doc.id).getBody().getText();
  Drive.Files.remove(doc.id);
  if (!texto || texto.length < 30) throw new Error('No se pudo extraer texto del archivo');
  return analizarCaso(casoCodigo, texto, usuario, antecedentes);
}

// ============ EVIDENCIAS DEL EXPEDIENTE (v3.1) ============
// Sube capturas/declaraciones/informes a la carpeta Drive DEL CASO y las
// registra en la hoja 'Documentos' con tipo 'EVIDENCIA'.
// listarEvidencias devuelve TODOS los documentos del caso.
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

// ============ ASESOR CONVERSACIONAL ============
const PERSONA_ASESOR = '\n== TU IDENTIDAD EN LA CONVERSACIÓN ==\n' +
  'Eres el asesor legal personal del usuario, su MANO DERECHA y colega de despacho: juntos son dos abogados laboralistas (régimen agrario y general) que analizan casos en equipo.\n' +
  '- Trátalo por su nombre, con calidez de colega y respeto profesional. Si es el inicio de la conversación, salúdalo breve y natural antes de entrar al caso.\n' +
  '- Habla COMO PERSONA, no como chatbot: frases naturales, cero muletillas de asistente ("como modelo de IA...", "estoy aquí para ayudarte" están PROHIBIDAS).\n' +
  '- Sé FRANCO y directo: si el expediente está flojo, dilo sin rodeos ("Mira, aquí estamos débiles: ..."). Si el usuario propone algo arriesgado, adviérteselo con argumentos. Si está bien encaminado, confírmaselo con seguridad.\n' +
  '- Piensa en voz alta como en una reunión de trabajo: "yo aquí veo dos caminos...", "te soy honesto, esto no resiste una impugnación...", "¿tú qué información tienes de...?".\n' +
  '- Debate con altura: puedes discrepar y sostener tu posición con la norma, pero la decisión final es del usuario.\n' +
  '- PRIMERO LA EMPRESA: los reglamentos (RIT y RISST) de cada empresa son DIFERENTES. Si no tienes certeza de a qué empresa pertenece el trabajador y bajo qué régimen está, PREGÚNTALO ANTES de citar cualquier artículo. NUNCA mezcles reglamentos de empresas distintas; di siempre de cuál reglamento estás citando.\n' +
  '- GESTIONA EL CIERRE DEL EXPEDIENTE: actúa como el abogado responsable del caso. En cada revisión identifica y pide expresamente lo que falta ("Para cerrar este caso necesitamos: 1) el descargo del trabajador, 2) la constancia de capacitación, 3) ..."). Lleva la cuenta de los pendientes ya cubiertos y los que quedan; cuando todo esté completo, propón el levantamiento de observaciones, la medida final y ofrece preparar la exposición para jefaturas y gerencia.\n' +
  '- Breve y al grano: respuestas de conversación, no ensayos, salvo que pida un documento.';

function consultarCaso(casoCodigo, mensajes, usuario) {
  let contexto = '';
  let caso = null;
  if (casoCodigo) caso = _buscarCaso(casoCodigo);

  if (caso) {
    const ev = obtenerEvaluacion(casoCodigo);
    const inv = obtenerInvestigacion(casoCodigo);
    const esSST = /accidente|epp|seguridad|salud|incidente/i.test(String(caso.tipo_falta));
    const reglamento = _reglamento(caso.empresa, esSST ? 'RISST' : 'RIT');

    contexto = 'CONTEXTO DEL EXPEDIENTE ' + casoCodigo + ':\n' +
      'Empresa: ' + caso.empresa + ' | Régimen: ' + caso.regimen + ' | Trabajador: ' + caso.trabajador +
      ' | Falta: ' + caso.tipo_falta + ' | Estado: ' + caso.estado + '\n' +
      'Descripción: ' + caso.descripcion_corta + '\n' +
      'Usuario que consulta: ' + (usuario || '-') + '\n';
    if (ev) contexto += '\nÚLTIMA EVALUACIÓN (v' + ev.version + '):\nResumen: ' + ev.resumen_ejecutivo +
      '\nConclusión: ' + ev.conclusion + '\nMedida sugerida: ' + ev.medida_recomendada +
      '\nRiesgo: ' + ev.nivel_riesgo + '\nFundamentos: ' + ev.fundamentos + '\n';
    if (inv) contexto += '\nINVESTIGACIÓN v' + inv.version + ' — Hallazgos: ' + inv.hallazgos +
      '\nAuditoría del expediente: ' + inv.auditoria +
      '\nConclusión preliminar: ' + inv.conclusion_preliminar +
      '\nSi es el inicio de la conversación, tras el saludo dile cuántos hallazgos podrían modificar la decisión y ofrece revisarlos UNO POR UNO, como colega que le reporta.\n';
    contexto += '\nREGLAMENTO APLICABLE (extracto):\n' + reglamento.slice(0, 40000);
  } else {
    contexto = 'MODO CONSULTA GENERAL: no hay expediente cargado. El usuario (' + (usuario || '-') + ') quiere conversar sobre sus casos, dudas laborales o estrategia. ' +
      'Orienta con el marco legal (agrario y general) y, si un tema amerita expediente formal, sugiérele crearlo o indicar el código para revisarlo a fondo.';
  }

  contexto += '\n\nOrienta las medidas con la escala LEVE/GRAVE/MUY GRAVE/CRÍTICO y pide la información que te falte. Texto claro (no JSON).';

  const msgs = [{ role: 'system', content: MARCO_LEGAL + PERSONA_ASESOR + '\n\n' + contexto }];
  (mensajes || []).slice(-10).forEach(function(m) { msgs.push(m); });

  const respuesta = _openai(msgs, true);
  _auditar(usuario, 'Consulta', casoCodigo || 'GENERAL', 'ASESOR', String((mensajes || []).length) + ' turnos');
  return { respuesta: respuesta };
}

// ============ EXPOSICIÓN PARA GERENCIA ============
function exposicionGerencia(casoCodigo, usuario) {
  const caso = _buscarCaso(casoCodigo);
  if (!caso) throw new Error('Caso no encontrado: ' + casoCodigo);
  const ev = obtenerEvaluacion(casoCodigo);
  const inv = obtenerInvestigacion(casoCodigo);
  if (!ev && !inv) throw new Error('Este caso aún no tiene investigación ni análisis. Corre primero la Fase 1 (Investigar) o el análisis jurídico.');

  const prompt = 'Prepara una EXPOSICIÓN EJECUTIVA para las jefaturas y la gerencia sobre el expediente ' + casoCodigo +
    ' (' + caso.empresa + ' | régimen ' + caso.regimen + ' | trabajador: ' + caso.trabajador + ' | falta: ' + caso.tipo_falta + ').\n' +
    (inv ? '\nINVESTIGACIÓN — Expediente: ' + inv.expediente + '\nAuditoría: ' + inv.auditoria + '\nHallazgos: ' + inv.hallazgos + '\n' : '') +
    (ev ? '\nANÁLISIS JURÍDICO — Resumen: ' + ev.resumen_ejecutivo + '\nConclusión: ' + ev.conclusion +
      '\nMedida sugerida: ' + ev.medida_recomendada + '\nRiesgo: ' + ev.nivel_riesgo + '\nFundamentos: ' + ev.fundamentos + '\n' : '') +
    '\nFormato para exponer en 3 minutos, lenguaje ejecutivo claro (gerencia no es abogada):\n' +
    '1. EL CASO EN DOS FRASES.\n2. HECHOS CLAVE (viñetas cortas).\n3. LO ACREDITADO vs LO NO ACREDITADO.\n' +
    '4. EVALUACIÓN LEGAL EN SIMPLE (qué falta es, del reglamento de qué empresa, y por qué).\n' +
    '5. MEDIDA SUGERIDA Y SU PROPORCIONALIDAD (por qué esta y no una mayor o menor).\n' +
    '6. RIESGO SI SE IMPUGNA (y costo según el régimen).\n7. PENDIENTES PARA CERRAR EL EXPEDIENTE.\n' +
    'Texto claro, no JSON.';

  const r = _openai([{ role: 'system', content: MARCO_LEGAL + PERSONA_ASESOR }, { role: 'user', content: prompt }], true);
  _auditar(usuario, 'Exposicion', casoCodigo, 'GERENCIA', '');
  return { respuesta: r };
}

// ============ MOTOR DE INVESTIGACIÓN (FASE 1) ============
function _hojaInvestigaciones() {
  let h = _ss().getSheetByName('Investigaciones');
  if (!h) {
    h = _ss().insertSheet('Investigaciones');
    h.appendRow(['id','caso_codigo','version','expediente','auditoria','hallazgos','conclusion_preliminar','nivel_confianza','fecha','generado_por']);
    h.setFrozenRows(1);
  }
  return h;
}

function investigarCaso(casoCodigo, textoInforme, usuario, antecedentes) {
  const caso = _buscarCaso(casoCodigo);
  if (!caso) throw new Error('Caso no encontrado: ' + casoCodigo);
  _hojaInvestigaciones(); // asegura la hoja ANTES de versionar

  const prompt = 'FASE 1 — INVESTIGACIÓN. Actúa como Coordinador Senior de Relaciones Laborales con 20+ años investigando casos. ' +
    'NO emitas opinión jurídica NI medida disciplinaria todavía. Lee COMPLETO el informe y construye el expediente. ' +
    'PROHIBIDO inventar: si un dato no está, regístralo como vacío. PROHIBIDO preguntar lo que ya está en el informe.\n\n' +
    'EXPEDIENTE ' + casoCodigo + ' | EMPRESA: ' + caso.empresa + ' | RÉGIMEN: ' + caso.regimen +
    '\n\nINFORME:\n' + textoInforme +
    (antecedentes ? '\n\nRÉCORD DE SANCIONES:\n' + antecedentes : '') +
    '\n\nResponde SOLO un JSON válido con estas claves exactas:\n' +
    '"expediente": {"personas":[{"nombre","dni","cargo","rol" (TRABAJADOR INVESTIGADO|SUPERVISOR|ADMINISTRADOR|JEFE|TESTIGO|AGRAVIADO|OTRO),"declaracion" (resumen FIEL de lo que declaró, o null)}], ' +
    '"cronologia":[{"fecha","hora","hecho","fuente" (quién lo afirma)}], ' +
    '"evidencias":[{"tipo" (FOTO|VIDEO|CORREO|WHATSAPP|DOCUMENTO|TESTIMONIO|OTRO),"descripcion","estado" (ADJUNTA|SOLO MENCIONADA)}], ' +
    '"normas_citadas":[{"norma","articulo","como_la_cita_el_informe"}], "lugares":[], "fechas_clave":[]}, ' +
    '"auditoria": {"contradicciones":[{"descripcion","entre_quienes","impacto"}], "vacios":[], "documentos_faltantes":[], "entrevistas_faltantes":[], ' +
    '"errores_del_supervisor":[], "errores_de_procedimiento":[] (debido proceso, inmediatez, non bis in idem), ' +
    '"calidad_informe":{"nota_0_a_100","sustento"}, "nivel_confianza" (ALTO|MEDIO|BAJO: cuán confiable es el expediente para decidir)}, ' +
    '"hallazgos":[{"titulo","detalle","podria_cambiar_la_decision" (true/false)}], ' +
    '"conclusion_preliminar" (2-3 frases, SIN medida), ' +
    '"preguntas_al_usuario":[SOLO lo que NO esté en el informe].';

  const r = _openai([{ role: 'system', content: MARCO_LEGAL }, { role: 'user', content: prompt }]);
  const inv = JSON.parse(r.replace(/```json|```/g, '').trim());

  const version = _versionSiguiente('Investigaciones', casoCodigo);
  _hojaInvestigaciones().appendRow([_id(), casoCodigo, version,
    JSON.stringify(inv.expediente), JSON.stringify(inv.auditoria), JSON.stringify(inv.hallazgos),
    inv.conclusion_preliminar, (inv.auditoria && inv.auditoria.nivel_confianza) || '',
    _ahora(), usuario || 'web']);
  _actualizarCaso(casoCodigo, { estado: 'EN ANALISIS' });
  CacheService.getScriptCache().remove('casos_recientes');
  _auditar(usuario, 'Investigacion', casoCodigo, 'INVESTIGAR', 'v' + version);
  return inv;
}

function obtenerInvestigacion(casoCodigo) {
  const datos = _hojaInvestigaciones().getDataRange().getValues();
  const head = datos[0];
  for (let i = datos.length - 1; i >= 1; i--) {
    if (datos[i][1] === casoCodigo) {
      const o = {};
      head.forEach(function(c, j) { o[c] = datos[i][j]; });
      return o;
    }
  }
  return null;
}

function investigarArchivo(casoCodigo, nombreArchivo, base64, mimeType, usuario, antecedentes) {
  const caso = _buscarCaso(casoCodigo);
  if (!caso) throw new Error('Caso no encontrado: ' + casoCodigo);
  const blob = Utilities.newBlob(Utilities.base64Decode(base64), mimeType, nombreArchivo);
  const archivo = DriveApp.getFolderById(caso.drive_folder_id).createFile(blob);
  _hoja('Documentos').appendRow([_id(), casoCodigo, 'INFORME', nombreArchivo,
    archivo.getId(), _ahora(), usuario || 'web', '', 'INVESTIGADO']);
  const doc = Drive.Files.create(
    { name: 'tmp_' + nombreArchivo, mimeType: 'application/vnd.google-apps.document' },
    blob, { ocrLanguage: 'es' }
  );
  const texto = DocumentApp.openById(doc.id).getBody().getText();
  Drive.Files.remove(doc.id);
  if (!texto || texto.length < 30) throw new Error('No se pudo extraer texto del archivo');
  return investigarCaso(casoCodigo, texto, usuario, antecedentes);
}

// ============ CASO AUTOMÁTICO (n8n / buzón) ============
function casoAutomatico(nombreArchivo, base64, mimeType, empresa, origen) {
  const blob = Utilities.newBlob(Utilities.base64Decode(base64), mimeType || 'application/pdf', nombreArchivo);
  const doc = Drive.Files.create(
    { name: 'tmp_' + nombreArchivo, mimeType: 'application/vnd.google-apps.document' },
    blob, { ocrLanguage: 'es' }
  );
  const texto = DocumentApp.openById(doc.id).getBody().getText();
  Drive.Files.remove(doc.id);
  if (!texto || texto.length < 30) throw new Error('No se pudo extraer texto de ' + nombreArchivo);

  const r = _openai([{ role: 'system', content: MARCO_LEGAL }, { role: 'user', content:
    'Extrae del informe los datos del caso. Responde SOLO JSON con estas claves exactas: ' +
    'trabajador, dni, cargo, sector, ruta, supervisor, administrador, ' +
    'tipo_falta (Tardanza|Inasistencia|Abandono de puesto|Indisciplina|Daño a propiedad|Seguridad y salud|Otro), ' +
    'fecha_hechos (AAAA-MM-DD), lugar_hechos, descripcion_corta (máximo 2 frases). ' +
    'Si un dato no aparece, usa "".\n\nINFORME:\n' + texto.slice(0, 20000) }]);
  const d = JSON.parse(r.replace(/```json|```/g, '').trim());

  d.empresa = String(empresa || (listarEmpresas()[0] || {}).nombre || 'GEN').toUpperCase();
  d.regimen = 'agrario';
  d.creado_por = origen || 'n8n';
  const caso = crearCaso(d);

  const archivo = DriveApp.getFolderById(caso.drive_folder_id).createFile(blob);
  _hoja('Documentos').appendRow([_id(), caso.codigo, 'INFORME', nombreArchivo,
    archivo.getId(), _ahora(), origen || 'n8n', '', 'AUTO']);

  const inv = investigarCaso(caso.codigo, texto, origen || 'n8n');
  const criticos = (inv.hallazgos || []).filter(function(h){ return h.podria_cambiar_la_decision; }).length;
  _alertaCorreo('🤖 Caso automático ' + caso.codigo + ' — ya investigado',
    'Informe detectado: ' + nombreArchivo + '\nExpediente: ' + caso.codigo + ' (' + d.empresa + ')\n' +
    'Trabajador: ' + (d.trabajador || '-') + ' | Falta: ' + (d.tipo_falta || '-') + '\n' +
    'Confianza del expediente: ' + ((inv.auditoria && inv.auditoria.nivel_confianza) || '-') +
    ' | Hallazgos críticos: ' + criticos + '\n' +
    'Conclusión preliminar: ' + (inv.conclusion_preliminar || '-') +
    '\n\nEntra al sistema para revisarlo y correr el análisis jurídico.');
  return { codigo: caso.codigo, hallazgos_criticos: criticos,
    confianza: (inv.auditoria && inv.auditoria.nivel_confianza) || '' };
}

// ============ USUARIOS Y SESIONES (con periodo de prueba) ============
function _hash(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8)
    .map(function(b){ b = (b + 256) % 256; return ('0' + b.toString(16)).slice(-2); }).join('');
}

function _hojaUsuarios() {
  let h = _ss().getSheetByName('Usuarios');
  if (!h) {
    h = _ss().insertSheet('Usuarios');
    h.appendRow(['usuario','hash','nombre','rol','activo','vence','empresas']);
    h.setFrozenRows(1);
  } else {
    if (h.getLastColumn() < 6) h.getRange(1, 6).setValue('vence');
    if (h.getLastColumn() < 7) h.getRange(1, 7).setValue('empresas'); // migración
  }
  return h;
}

function setupUsuarios() { // correr UNA vez
  const h = _hojaUsuarios();
  if (h.getLastRow() < 2) {
    h.appendRow(['jtimoteo', _hash('marciatimoteo2026'), 'Joel Timoteo', 'ADMIN', 'TRUE', '', 'TODAS']);
  }
  Logger.log('✅ Hoja Usuarios lista. Usuario inicial: jtimoteo');
}

function crearUsuario(usuario, clave, nombre, rol) {
  _hojaUsuarios().appendRow([usuario, _hash(clave), nombre, rol || 'USUARIO', 'TRUE', '']);
}

function loginUsuario(usuario, clave) {
  const datos = _hojaUsuarios().getDataRange().getValues();
  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][0]).toLowerCase().trim() === String(usuario).toLowerCase().trim() &&
        datos[i][1] === _hash(String(clave))) {
      if (String(datos[i][4]).toUpperCase() !== 'TRUE')
        throw new Error('Tu acceso está desactivado. Contacta al administrador.');
      const vence = datos[i][5];
      if (vence) {
        const v = new Date(vence);
        if (!isNaN(v.getTime()) && new Date() > v)
          throw new Error('Tu periodo de prueba venció. Contacta al administrador para activar tu licencia.');
      }
      let acceso = String((datos[i].length > 6 ? datos[i][6] : '') || '').toUpperCase().trim();
      if (datos[i][3] === 'ADMIN' || !acceso) acceso = 'TODAS';
      const token = Utilities.getUuid();
      CacheService.getScriptCache().put('sess_' + token,
        JSON.stringify({ u: datos[i][0], n: datos[i][2], r: datos[i][3], e: acceso }), 21600); // 6 horas
      _auditar(datos[i][0], 'Sesion', datos[i][0], 'LOGIN', acceso);
      return { token: token, nombre: datos[i][2], rol: datos[i][3], empresas: _empresasPorAlcance(acceso) };
    }
  }
  throw new Error('Usuario o contraseña incorrectos');
}

function _sesion(token) {
  if (!token) return null;
  const s = CacheService.getScriptCache().get('sess_' + token);
  return s ? JSON.parse(s) : null;
}

// ============ ALCANCE POR EMPRESA (privacidad entre usuarios) ============
function _empresasDe(ses) {
  if (!ses) return 'TODAS';               // token maestro (n8n)
  if (ses.r === 'ADMIN') return 'TODAS';  // los admin ven todo
  return ses.e || 'TODAS';
}

function _puedeVer(ses, empresaCaso) {
  const alc = _empresasDe(ses);
  if (alc === 'TODAS') return true;
  return alc.split(',').map(function(x){ return x.trim().toUpperCase(); })
    .indexOf(String(empresaCaso).toUpperCase()) !== -1;
}

function _checkCaso(ses, codigo) {
  const c = _buscarCaso(codigo);
  if (!c) throw new Error('Caso no encontrado: ' + codigo);
  if (!_puedeVer(ses, c.empresa)) throw new Error('No tienes acceso a este expediente');
}

function _empresasPorAlcance(alcance) {
  let e = listarEmpresas();
  if (alcance && alcance !== 'TODAS') {
    const s = alcance.split(',').map(function(x){ return x.trim().toUpperCase(); });
    e = e.filter(function(x){ return s.indexOf(x.nombre) !== -1; });
  }
  return e;
}

// ============ GESTIÓN DE USUARIOS (solo ADMIN) ============
function _soloAdmin(ses) { if (!ses || ses.r !== 'ADMIN') throw new Error('Solo el administrador puede hacer esto'); }

function listarUsuarios() {
  const d = _hojaUsuarios().getDataRange().getValues();
  const out = [];
  for (let i = 1; i < d.length; i++) {
    let vence = d[i][5];
    if (vence && vence.getTime) vence = Utilities.formatDate(vence, 'America/Lima', 'yyyy-MM-dd');
    out.push({ usuario: d[i][0], nombre: d[i][2], rol: d[i][3],
      activo: String(d[i][4]).toUpperCase() === 'TRUE', vence: vence || '',
      empresas: String((d[i].length > 6 ? d[i][6] : '') || 'TODAS') });
  }
  return out;
}

function adminCrearUsuario(usuario, clave, nombre, rol, diasPrueba, empresasAcceso) {
  if (!usuario || !clave || !nombre) throw new Error('Completa usuario, contraseña y nombre');
  if (String(clave).length < 8) throw new Error('La contraseña debe tener al menos 8 caracteres');
  const d = _hojaUsuarios().getDataRange().getValues();
  for (let i = 1; i < d.length; i++)
    if (String(d[i][0]).toLowerCase() === String(usuario).toLowerCase()) throw new Error('El usuario ya existe');
  let vence = '';
  const dias = parseInt(diasPrueba, 10);
  if (dias > 0) {
    const f = new Date(); f.setDate(f.getDate() + dias);
    vence = Utilities.formatDate(f, 'America/Lima', 'yyyy-MM-dd');
  }
  const acceso = String(empresasAcceso || 'TODAS').toUpperCase().trim();
  _hojaUsuarios().appendRow([String(usuario).toLowerCase().trim(), _hash(String(clave)), nombre, rol || 'USUARIO', 'TRUE', vence, acceso]);
  return { creado: usuario, vence: vence || 'sin límite', empresas: acceso };
}

function adminEstadoUsuario(usuario, activo) {
  const h = _hojaUsuarios(), d = h.getDataRange().getValues();
  for (let i = 1; i < d.length; i++)
    if (String(d[i][0]).toLowerCase() === String(usuario).toLowerCase()) {
      h.getRange(i + 1, 5).setValue(activo ? 'TRUE' : 'FALSE');
      return { usuario: usuario, activo: activo };
    }
  throw new Error('Usuario no encontrado');
}

function adminRenovarUsuario(usuario, dias) {
  const h = _hojaUsuarios(), d = h.getDataRange().getValues();
  for (let i = 1; i < d.length; i++)
    if (String(d[i][0]).toLowerCase() === String(usuario).toLowerCase()) {
      let vence = '';
      const n = parseInt(dias, 10);
      if (n > 0) {
        const f = new Date(); f.setDate(f.getDate() + n);
        vence = Utilities.formatDate(f, 'America/Lima', 'yyyy-MM-dd');
      }
      h.getRange(i + 1, 6).setValue(vence);
      h.getRange(i + 1, 5).setValue('TRUE');
      return { usuario: usuario, vence: vence || 'sin límite' };
    }
  throw new Error('Usuario no encontrado');
}

// ============ API WEB ============
function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);

    if (req.accion === 'login') return _json({ ok: true, data: loginUsuario(req.usuario, req.clave) });

    const maestro = PropertiesService.getScriptProperties().getProperty('API_TOKEN');
    const ses = _sesion(req.token);
    if (!ses && req.token !== maestro) return _json({ ok: false, error: 'Sesión expirada' });
    const quien = ses ? ses.u : (req.usuario || 'n8n');

    let out;
    switch (req.accion) {
      case 'crearCaso':          if (!_puedeVer(ses, (req.datos || {}).empresa)) throw new Error('No tienes acceso a esa empresa'); out = crearCaso(Object.assign({}, req.datos, { creado_por: quien })); break;
      case 'analizarCaso':       _checkCaso(ses, req.codigo); out = analizarCaso(req.codigo, req.texto, quien, req.antecedentes); break;
      case 'analizarArchivo':    _checkCaso(ses, req.codigo); out = analizarArchivo(req.codigo, req.nombre, req.base64, req.mime, quien, req.antecedentes); break;
      case 'investigar':         _checkCaso(ses, req.codigo); out = investigarCaso(req.codigo, req.texto, quien, req.antecedentes); break;
      case 'investigarArchivo':  _checkCaso(ses, req.codigo); out = investigarArchivo(req.codigo, req.nombre, req.base64, req.mime, quien, req.antecedentes); break;
      case 'casoAutomatico':     if (!_puedeVer(ses, req.empresa)) throw new Error('No tienes acceso a esa empresa'); out = casoAutomatico(req.nombre, req.base64, req.mime, req.empresa, quien); break;
      case 'buscarCaso':         _checkCaso(ses, req.codigo); out = _buscarCaso(req.codigo); break;
      case 'listarCasos':        out = listarCasos(req.filtro, ses); break;
      case 'obtenerEvaluacion':  _checkCaso(ses, req.codigo); out = obtenerEvaluacion(req.codigo); break;
      case 'subirEvidencia':     _checkCaso(ses, req.codigo); out = subirEvidencia(req, quien); break;
      case 'listarEvidencias':   _checkCaso(ses, req.codigo); out = listarEvidencias(req.codigo); break;
      case 'consultar':          if (req.codigo) _checkCaso(ses, req.codigo); out = consultarCaso(req.codigo, req.mensajes, quien); break;
      case 'exposicion':         _checkCaso(ses, req.codigo); out = exposicionGerencia(req.codigo, quien); break;
      case 'listarEmpresas':     out = _empresasPorAlcance(_empresasDe(ses)); break;
      case 'listarUsuarios':     _soloAdmin(ses); out = listarUsuarios(); break;
      case 'crearUsuarioAdmin':  _soloAdmin(ses); out = adminCrearUsuario(req.usuario_nuevo, req.clave, req.nombre, req.rol, req.dias_prueba, req.empresas_acceso); _auditar(quien, 'Usuario', req.usuario_nuevo, 'CREAR', (req.rol || '') + ' ' + (req.dias_prueba || 'sin límite') + ' → ' + (req.empresas_acceso || 'TODAS')); break;
      case 'estadoUsuario':      _soloAdmin(ses); out = adminEstadoUsuario(req.usuario_obj, req.activo); _auditar(quien, 'Usuario', req.usuario_obj, req.activo ? 'ACTIVAR' : 'DESACTIVAR', ''); break;
      case 'renovarUsuario':     _soloAdmin(ses); out = adminRenovarUsuario(req.usuario_obj, req.dias); _auditar(quien, 'Usuario', req.usuario_obj, 'RENOVAR', String(req.dias || 'sin límite')); break;
      case 'agregarEmpresa':     _soloAdmin(ses); out = agregarEmpresa(req.empresa_nombre, req.empresa_prefijo); _auditar(quien, 'Empresa', req.empresa_nombre, 'CREAR', ''); break;
      case 'extraerReglamentosEmpresa': _soloAdmin(ses); out = extraerReglamentosEmpresa(req.empresa_nombre); _auditar(quien, 'Reglamento', req.empresa_nombre, 'EXTRAER', ''); break;
      default: throw new Error('Acción desconocida: ' + req.accion);
    }
    return _json({ ok: true, data: out });
  } catch (err) {
    return _json({ ok: false, error: String(err) });
  }
}
function _json(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

// ============ SETUP INICIAL (instancia de Joel) ============
function setupInicial() {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('API_TOKEN')) props.setProperty('API_TOKEN', Utilities.getUuid());
  const cat = _hoja('Catalogos');
  if (cat && cat.getLastRow() < 2) {
    [['Regimen','agrario'],['Regimen','general'],
     ['TipoFalta','Tardanza'],['TipoFalta','Inasistencia'],['TipoFalta','Abandono de puesto'],
     ['TipoFalta','Indisciplina'],['TipoFalta','Daño a propiedad'],['TipoFalta','Seguridad y salud'],
     ['Estado','ABIERTO'],['Estado','EN ANALISIS'],['Estado','CERRADO']
    ].forEach(function(f) { cat.appendRow([f[0], f[1], 'TRUE']); });
  }
  // sembrar empresas de Joel si no existen
  try { agregarEmpresa('RAPEL', 'RAP'); } catch (e) {}
  try { agregarEmpresa('VERFRUT', 'VER'); } catch (e) {}
  Logger.log('✅ Setup listo. API_TOKEN: ' + props.getProperty('API_TOKEN'));
  Logger.log('Carpeta de casos: ' + DriveApp.getFolderById(_carpetaCasosId()).getUrl());
}

// ============ INSTALADOR PARA VENTA (cuenta del comprador) ============
// 1) El comprador crea un proyecto GAS en SU cuenta y pega este archivo completo.
// 2) Habilita el servicio "Drive API" y guarda su GEMINI_API_KEY en Propiedades.
// 3) Edita y ejecuta instalarParaCliente() UNA vez.
// 4) Implementar → Aplicación web (Ejecutar como: yo / Acceso: cualquiera) → copiar URL.
// 5) En index.html cambiar API_URL por la URL nueva → subir a su GitHub Pages.
function instalarParaCliente() {
  setupCompleto('admin', 'CambiarEstaClave123', 'Administrador', 'MI EMPRESA', 'EMP');
}

function setupCompleto(adminUsuario, adminClave, adminNombre, empresaNombre, empresaPrefijo) {
  const ESQUEMA = {
    'Casos': ['codigo','estado','empresa','regimen','trabajador','dni','cargo','sector','ruta','supervisor','administrador','tipo_falta','fecha_hechos','lugar_hechos','descripcion_corta','nivel_riesgo','medida_recomendada','fecha_creacion','fecha_cierre','creado_por','drive_folder_id'],
    'Personas': ['id','caso_codigo','rol','nombre','dni','cargo','regimen','grado_responsabilidad','gravedad_falta','medida_individual'],
    'Documentos': ['id','caso_codigo','tipo','nombre_archivo','drive_file_id','fecha_carga','cargado_por','hash','estado_analisis'],
    'Evaluaciones': ['id','caso_codigo','version','resumen_ejecutivo','cronologia','responsabilidades','agravantes','atenuantes','riesgo_legal','fundamentos','conclusion','medida_recomendada','nivel_riesgo','observaciones','modelo_ia','fecha','generado_por'],
    'Generados': ['id','caso_codigo','tipo_documento','version','drive_file_id','estado','generado_por','fecha'],
    'Auditoria': ['id','fecha_hora','usuario','entidad','entidad_id','accion','detalle'],
    'Catalogos': ['categoria','valor','activo'],
    'Reglamentos': ['clave','parte','texto'],
    'Usuarios': ['usuario','hash','nombre','rol','activo','vence'],
    'Empresas': ['nombre','prefijo','activo'],
    'Investigaciones': ['id','caso_codigo','version','expediente','auditoria','hallazgos','conclusion_preliminar','nivel_confianza','fecha','generado_por']
  };
  const ss = SpreadsheetApp.create('RL_Intelligence_DB');
  Object.keys(ESQUEMA).forEach(function(nombre, i) {
    const hoja = (i === 0) ? ss.getSheets()[0].setName(nombre) : ss.insertSheet(nombre);
    hoja.getRange(1, 1, 1, ESQUEMA[nombre].length).setValues([ESQUEMA[nombre]]).setFontWeight('bold');
    hoja.setFrozenRows(1);
  });

  const props = PropertiesService.getScriptProperties();
  props.setProperty('SPREADSHEET_ID', ss.getId());
  props.setProperty('API_TOKEN', Utilities.getUuid());
  props.deleteProperty('CARPETA_CASOS');

  ss.getSheetByName('Usuarios').appendRow([
    String(adminUsuario).toLowerCase().trim(), _hash(String(adminClave)), adminNombre, 'ADMIN', 'TRUE', '', 'TODAS'
  ]);
  ss.getSheetByName('Empresas').appendRow([
    String(empresaNombre).toUpperCase().trim(), String(empresaPrefijo).toUpperCase().trim().slice(0, 4), 'TRUE'
  ]);
  [['Regimen','agrario'],['Regimen','general'],
   ['TipoFalta','Tardanza'],['TipoFalta','Inasistencia'],['TipoFalta','Abandono de puesto'],
   ['TipoFalta','Indisciplina'],['TipoFalta','Daño a propiedad'],['TipoFalta','Seguridad y salud'],
   ['Estado','ABIERTO'],['Estado','EN ANALISIS'],['Estado','CERRADO']
  ].forEach(function(f) { ss.getSheetByName('Catalogos').appendRow([f[0], f[1], 'TRUE']); });

  Logger.log('✅ INSTALACIÓN COMPLETA');
  Logger.log('Base de datos: ' + ss.getUrl());
  Logger.log('Usuario administrador: ' + adminUsuario);
  Logger.log('API_TOKEN (para automatizaciones): ' + props.getProperty('API_TOKEN'));
  Logger.log('Siguiente paso: Implementar como Aplicación web y pegar la URL en index.html');
}

// ============ PRUEBA RÁPIDA ============
function testCrearCaso() {
  const r = crearCaso({ empresa: (listarEmpresas()[0] || {}).nombre || 'GEN', regimen: 'agrario',
    trabajador: 'PRUEBA SISTEMA', tipo_falta: 'Tardanza', descripcion_corta: 'Caso de prueba', creado_por: 'admin' });
  Logger.log(r);
}
