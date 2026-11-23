/**
 * Tecno Log IA · Catálogo App — Inventario y relación de material (v2, por tandas)
 *
 * Revisa las carpetas de Dakota, Ventus e Ilumi Home (y todas sus subcarpetas)
 * en tandas de 4,5 minutos. Si no termina, se programa solo para continuar al minuto
 * siguiente, hasta acabar. Al final relaciona fotos, fichas, manuales y videos con
 * cada producto de «Productos» por su código.
 *
 * Uso: ejecutar «inventariar» UNA vez. Luego todo sigue solo.
 * El avance se ve en la hoja, pestaña Config (fila «Inventario»).
 */
const ID_HOJA = '1_XtTBAGbIJFcbP6tLNTU4cADwF1eKN59WaX1olJkvOE';
const CARPETAS = [
  { marca: 'Dakota', id: '1mBRuedD2DCBgMm3hA0HHwH1ATQq63V7d' },
  { marca: 'Ventus', id: '1isrj-_gpln9VQE0DXuookwjVGTVGrd5S' },
  { marca: 'Ilumi',  id: '16WbcpYYwgnsVgCEwHIIDdkOCTZ7DHYqT' },
];
const LIMITE_MS = 4.5 * 60 * 1000;

function hoja_() { return SpreadsheetApp.openById(ID_HOJA); }

function estadoSheet_(ss) {
  let sh = ss.getSheetByName('_estado');
  if (!sh) { sh = ss.insertSheet('_estado'); sh.hideSheet(); }
  return sh;
}
// Estado: A1 = datos pequeños (JSON). Cola de carpetas: filas desde A3 (id, ruta, marca).
function leerEstado_(ss) {
  const sh = estadoSheet_(ss);
  const v = sh.getRange('A1').getValue();
  if (!v) return null;
  const e = JSON.parse(v);
  const n = sh.getLastRow() - 2;
  e.cola = n > 0 ? sh.getRange(3, 1, n, 3).getValues().filter(r => r[0]).map(r => ({ id: r[0], ruta: r[1], marca: r[2] })) : [];
  return e;
}
function guardarEstado_(ss, e) {
  const sh = estadoSheet_(ss);
  sh.clearContents();
  if (!e) return;
  const meta = Object.assign({}, e); delete meta.cola;
  sh.getRange('A1').setValue(JSON.stringify(meta));
  if (e.cola.length) {
    if (sh.getMaxRows() < e.cola.length + 3) sh.insertRowsAfter(sh.getMaxRows(), e.cola.length + 3 - sh.getMaxRows());
    sh.getRange(3, 1, e.cola.length, 3).setValues(e.cola.map(c => [c.id, c.ruta, c.marca]));
  }
}
function avance_(ss, texto) {
  const c = ss.getSheetByName('Config');
  c.getRange('A14:B14').setValues([['Inventario de material', texto]]);
}

function borrarDisparadores_() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'continuar')
    .forEach(t => ScriptApp.deleteTrigger(t));
}

/** Ejecutar esta función una sola vez. */
function inventariar() {
  const ss = hoja_();
  borrarDisparadores_();
  const shM = ss.getSheetByName('Material');
  shM.getRange(2, 1, Math.max(shM.getMaxRows() - 1, 1), 8).clearContent();
  guardarEstado_(ss, {
    cola: CARPETAS.map(c => ({ id: c.id, ruta: '', marca: c.marca })),
    actual: null, tokenArchivos: null, fila: 2, archivos: 0, tandas: 0,
  });
  avance_(ss, 'En curso… (iniciado ' + new Date().toLocaleString() + ')');
  continuar();
}

function continuar() {
  const ss = hoja_();
  const e = leerEstado_(ss);
  if (!e) return;
  const t0 = Date.now();
  const shM = ss.getSheetByName('Material');
  const buffer = [];
  e.tandas++;

  while (Date.now() - t0 < LIMITE_MS) {
    if (!e.actual) {
      if (!e.cola.length) break;
      e.actual = e.cola.shift();
      e.tokenArchivos = null;
      // subcarpetas a la cola
      const carpeta = DriveApp.getFolderById(e.actual.id);
      if (!e.actual.ruta) e.actual.ruta = carpeta.getName().trim();
      const subs = carpeta.getFolders();
      while (subs.hasNext()) {
        const s = subs.next();
        e.cola.push({ id: s.getId(), ruta: e.actual.ruta + ' / ' + s.getName().trim(), marca: e.actual.marca });
      }
    }
    const it = e.tokenArchivos
      ? DriveApp.continueFileIterator(e.tokenArchivos)
      : DriveApp.getFolderById(e.actual.id).getFiles();
    let corto = false;
    while (it.hasNext()) {
      const f = it.next();
      const nombre = f.getName(), mime = f.getMimeType(), id = f.getId();
      buffer.push([e.actual.marca, e.actual.ruta, nombre, tipo_(nombre, mime, e.actual.ruta), '', id,
                   'https://drive.google.com/file/d/' + id + '/view', '']);
      if (Date.now() - t0 > LIMITE_MS) { e.tokenArchivos = it.getContinuationToken(); corto = true; break; }
    }
    if (corto) break;
    e.actual = null; e.tokenArchivos = null;
  }

  if (buffer.length) {
    const falta = e.fila + buffer.length - 1 - shM.getMaxRows();
    if (falta > 0) shM.insertRowsAfter(shM.getMaxRows(), falta + 100);
    shM.getRange(e.fila, 1, buffer.length, 8).setValues(buffer);
    e.fila += buffer.length;
    e.archivos += buffer.length;
  }

  const terminado = !e.actual && !e.cola.length;
  if (terminado) {
    guardarEstado_(ss, null);
    borrarDisparadores_();
    const msg = relacionar();
    avance_(ss, 'LISTO · ' + e.archivos + ' archivos en ' + e.tandas + ' tanda(s) · ' + msg + ' · Copiando fotos…');
    Logger.log('LISTO · ' + msg);
    copiarFotos();
  } else {
    guardarEstado_(ss, e);
    borrarDisparadores_();
    ScriptApp.newTrigger('continuar').timeBased().after(60 * 1000).create();
    const m = 'En curso… tanda ' + e.tandas + ' · ' + e.archivos + ' archivos · ' + e.cola.length + ' carpetas por revisar. Continúa solo en 1 minuto.';
    avance_(ss, m);
    Logger.log(m);
  }
}

function norm_(s) {
  return String(s || '').toUpperCase()
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .replace(/[\s\-_.\/]/g, '');
}

// Posibles códigos dentro de un nombre de archivo: cada palabra y pares de palabras seguidas
function claves_(nombre) {
  const base = String(nombre).replace(/\.[A-Za-z0-9]{2,4}$/, '');
  const t = base.split(/[\s_(),+]+/).filter(Boolean);
  const out = new Set();
  for (let i = 0; i < t.length; i++) {
    out.add(norm_(t[i]));
    if (i + 1 < t.length) out.add(norm_(t[i] + t[i + 1]));
  }
  out.add(norm_(base));
  return out;
}

function tipo_(nombre, mime, ruta) {
  const r = (ruta + ' ' + nombre).toUpperCase();
  if (mime.indexOf('image/') === 0) return 'Foto';
  if (mime.indexOf('video/') === 0) return 'Video';
  if (mime === 'application/pdf') {
    if (/MANUAL|INSTRUCT|USUARIO/.test(r)) return 'Manual';
    if (/FICHA|TECNIC|SPEC/.test(r)) return 'Ficha técnica';
    if (/CATALOG/.test(r)) return 'Catálogo';
    return 'PDF';
  }
  return 'Otro';
}

/** Relaciona el material con los productos (se llama sola al terminar; también se puede ejecutar a mano). */
function relacionar() {
  const ss = hoja_();
  const shM = ss.getSheetByName('Material');
  const shP = ss.getSheetByName('Productos');
  const nM = shM.getLastRow() - 1;
  const M = nM > 0 ? shM.getRange(2, 1, nM, 8).getValues() : [];
  const P = shP.getRange(2, 1, shP.getLastRow() - 1, 18).getValues();

  const porCodigo = {};
  P.forEach((p, i) => {
    const k = norm_(p[1]);
    if (k.length >= 3) (porCodigo[k] = porCodigo[k] || []).push(i);
  });

  const mat = P.map(() => ({ fotos: [], ficha: '', manual: '', video: '', marca: '' }));
  M.forEach(f => {
    const hits = new Set();
    claves_(f[2]).forEach(k => (porCodigo[k] || []).forEach(i => hits.add(i)));
    if (!hits.size) { f[4] = ''; f[7] = ''; return; }
    const codigos = [];
    hits.forEach(i => {
      codigos.push(P[i][1]);
      const m = mat[i];
      m.marca = m.marca || f[0];
      if (f[3] === 'Foto') m.fotos.push(f[5]);
      else if (f[3] === 'Ficha técnica') m.ficha = m.ficha || f[6];
      else if (f[3] === 'Manual') m.manual = m.manual || f[6];
      else if (f[3] === 'Video') m.video = m.video || f[6];
      else if (f[3] === 'PDF' && !m.ficha) m.ficha = f[6];
    });
    f[4] = codigos[0];
    f[7] = codigos.join(', ');
  });
  if (M.length) shM.getRange(2, 1, M.length, 8).setValues(M);

  const url = id => 'https://lh3.googleusercontent.com/d/' + id;
  const marcas = P.map((p, i) => [p[2] || mat[i].marca]);
  const resto = mat.map((m, i) => {
    const res = [];
    m.fotos = Array.from(new Set(m.fotos));
    if (m.fotos.length) res.push(m.fotos.length + ' foto(s)');
    if (m.ficha) res.push('ficha');
    if (m.manual) res.push('manual');
    if (m.video) res.push('video');
    // no borra lo que ya estuviera escrito a mano
    const p = P[i];
    return [p[12] || (m.fotos[0] ? url(m.fotos[0]) : ''),
            p[13] || m.fotos.slice(1).map(url).join(' | '),
            p[14] || m.ficha, p[15] || m.manual, p[16] || m.video,
            res.length ? res.join(' · ') : (p[12] ? 'foto propia' : 'SIN MATERIAL')];
  });
  shP.getRange(2, 3, marcas.length, 1).setValues(marcas);
  shP.getRange(2, 13, resto.length, 6).setValues(resto);

  const con = resto.filter(r => r[5] !== 'SIN MATERIAL').length;
  return 'Productos con material: ' + con + ' de ' + P.length;
}


// ======================= Fotos públicas para la app =======================
// Copia la foto principal y hasta 3 de galería de cada producto a TU Drive
// (carpeta «Tecno Log IA · Catálogo App · Fotos»), las comparte «cualquiera con el enlace»
// y reemplaza los enlaces en Productos (columnas M y N). Por tandas, sigue sola.
const MAX_FOTOS = 4;
function carpetaFotos_() {
  const it = DriveApp.getFoldersByName('Tecno Log IA · Catálogo App · Fotos');
  return it.hasNext() ? it.next() : DriveApp.createFolder('Tecno Log IA · Catálogo App · Fotos');
}
function mapaFotos_(ss) {
  let sh = ss.getSheetByName('_fotos');
  if (!sh) { sh = ss.insertSheet('_fotos'); sh.hideSheet(); sh.getRange('A1:B1').setValues([['origen', 'copia']]); }
  const n = sh.getLastRow() - 1;
  const m = {};
  if (n > 0) sh.getRange(2, 1, n, 2).getValues().forEach(r => { if (r[0]) m[r[0]] = r[1]; });
  return { sh, m };
}
const idDe_ = u => { const x = String(u || '').match(/\/d\/([\w-]{20,})/); return x ? x[1] : ''; };
const urlFoto_ = id => 'https://lh3.googleusercontent.com/d/' + id + '=w1000';

/** Ejecutar a mano si se quiere repetir solo la copia de fotos. */
function copiarFotos() {
  const ss = hoja_();
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'continuarFotos').forEach(t => ScriptApp.deleteTrigger(t));
  PropertiesService.getScriptProperties().setProperty('fotoFila', '2');
  continuarFotos();
}

function continuarFotos() {
  const ss = hoja_();
  const t0 = Date.now();
  const props = PropertiesService.getScriptProperties();
  let fila = Number(props.getProperty('fotoFila') || 2);
  const shP = ss.getSheetByName('Productos');
  const last = shP.getLastRow();
  const { sh: shF, m } = mapaFotos_(ss);
  const nuevos = [];
  const carpeta = carpetaFotos_();
  let copiadas = 0;

  while (fila <= last && Date.now() - t0 < LIMITE_MS) {
    const r = shP.getRange(fila, 13, 1, 2).getValues()[0]; // M foto principal, N galería
    const vistos = new Set();
    const urls = [r[0]].concat(String(r[1] || '').split('|')).map(s => String(s).trim()).filter(Boolean)
      .filter(u => { const k = idDe_(u) || u; if (vistos.has(k)) return false; vistos.add(k); return true; }).slice(0, MAX_FOTOS);
    const salida = [];
    urls.forEach(u => {
      const id = idDe_(u);
      if (!id || u.indexOf('github.io') >= 0) { salida.push(u); return; }
      if (m[id]) { salida.push(urlFoto_(m[id])); return; }
      if (Object.values(m).indexOf(id) >= 0) { salida.push(urlFoto_(id)); return; } // ya es copia propia
      try {
        const copia = DriveApp.getFileById(id).makeCopy(DriveApp.getFileById(id).getName(), carpeta);
        copia.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        m[id] = copia.getId(); nuevos.push([id, copia.getId()]); copiadas++;
        salida.push(urlFoto_(copia.getId()));
      } catch (err) { /* sin permiso o archivo borrado: se omite */ }
    });
    if (urls.length) shP.getRange(fila, 13, 1, 2).setValues([[salida[0] || '', salida.slice(1).join(' | ')]]);
    fila++;
  }
  if (nuevos.length) shF.getRange(shF.getLastRow() + 1, 1, nuevos.length, 2).setValues(nuevos);

  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'continuarFotos').forEach(t => ScriptApp.deleteTrigger(t));
  if (fila <= last) {
    props.setProperty('fotoFila', String(fila));
    ScriptApp.newTrigger('continuarFotos').timeBased().after(60 * 1000).create();
    avance_(ss, 'Copiando fotos… fila ' + fila + ' de ' + last + ' · ' + Object.keys(m).length + ' fotos copiadas. Continúa solo en 1 minuto.');
  } else {
    props.deleteProperty('fotoFila');
    avance_(ss, 'LISTO · material relacionado y ' + Object.keys(m).length + ' fotos publicadas para la app (' + new Date().toLocaleString() + ')');
  }
}
