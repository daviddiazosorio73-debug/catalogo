/**
 * Tecno Log IA · Catálogo App — API v2 «IA 1» (Google Apps Script)
 *
 * Se instala en script.google.com con la cuenta tecnologiapperu@gmail.com y se publica como
 * «Aplicación web» (Ejecutar como: Yo · Quién tiene acceso: Cualquier usuario).
 *
 *  GET  ?accion=productos        → configuración + productos visibles (lee la base maestra)
 *  GET  ?accion=ruc&ruc=20xxxxxx → datos públicos del cliente (razón social, dirección, ciudad)
 *  POST {accion:'cotizacion', ...} → cliente nuevo a «Clientes» (Etapa «Nuevo») o nota al existente + lista «Hoy»
 *  POST {accion:'orden', ...}    → registra la orden, genera el PDF con la firma, lo envía por correo y lo pone en tu lista «Hoy»
 *  POST {accion:'carrito', ...}  → guarda el carrito del cliente identificado (para carritos abandonados)
 *  POST {accion:'eventos', ...}  → búsquedas, productos vistos y agregados (métricas)
 *  POST {accion:'asesor', ...}   → Asesor IA (Google Gemini) que arma pedidos con productos reales del catálogo
 *
 *  Instalación única: ejecutar «instalar» (crea los avisos por hora y el resumen diario de las 7 a. m.).
 *  Asesor IA: en ⚙️ Configuración del proyecto → Propiedades del script → GEMINI_API_KEY = tu clave de Google AI Studio.
 *
 * Los precios de la orden SIEMPRE se recalculan aquí con la base maestra (no se confía en el navegador).
 */
const BASE_MAESTRA = '1_XtTBAGbIJFcbP6tLNTU4cADwF1eKN59WaX1olJkvOE';      // Tecno Log IA · Catálogo App (base maestra)
const HOJA_CLIENTES = '12gaQxik4JuXN6cXUf7AwoQEWBSoEMwf6Drj8sAywh-s';    // Tecno Log IA — Hoja maestra (distribuidores)
const CARPETA_ORDENES = 'Tecno Log IA · Órdenes de compra';
const ZONA = 'America/Lima';

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  try {
    if (p.accion === 'ruc') return json_(buscarRuc_(p.ruc));
    return json_(catalogo_());
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.accion === 'orden') return json_(registrarOrden_(d));
    if (d.accion === 'carrito') return json_(registrarCarrito_(d));
    if (d.accion === 'cotizacion') return json_(registrarCotizacion_(d));
    if (d.accion === 'eventos') return json_(registrarEventos_(d));
    if (d.accion === 'asesor') return json_(asesor_(d));
    return json_({ ok: false, error: 'Acción no válida' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

// ---------- Configuración ----------
function config_() {
  const v = SpreadsheetApp.openById(BASE_MAESTRA).getSheetByName('Config').getRange('A3:B30').getValues();
  const m = {};
  v.forEach(r => { if (r[0]) m[String(r[0]).trim()] = r[1]; });
  return {
    nombre: m['Nombre comercial'] || 'Tecno Log IA',
    correo: m['Correo de ventas (envía y recibe copia)'] || 'tecnologiapperu@gmail.com',
    copia: m['Correo de copia adicional'] || '',
    whatsapp: String(m['WhatsApp ventas'] || '51928146619'),
    moneda: m['Moneda'] || 'PEN',
    precios: m['Precios mostrados'] || 'PVP',
    vigencia: m['Vigencia de la lista'] instanceof Date
      ? Utilities.formatDate(m['Vigencia de la lista'], ZONA, 'MMMM yyyy').replace(/^./, c => c.toUpperCase())
      : String(m['Vigencia de la lista'] || ''),
    validez: Number(m['Validez de la orden (días)'] || 7),
    asesor: String(m['Asesor IA activo'] || 'Sí') === 'Sí' && !!PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY'),
    modelo: String(m['Modelo IA'] || 'gemini-flash-latest'),
    horasCarrito: Number(m['Horas para carrito abandonado'] || 2),
    hojaVentas: String(m['Hoja de ventas (lista Hoy)'] || HOJA_CLIENTES),
    correoResumen: String(m['Correo del resumen diario'] || ''),
  };
}

// ---------- Catálogo ----------
function productos_() {
  const sh = SpreadsheetApp.openById(BASE_MAESTRA).getSheetByName('Productos');
  const v = sh.getRange(2, 1, sh.getLastRow() - 1, 22).getValues();
  return v.filter(r => r[0] && r[9] !== 'Oculto').map(r => ({
    id: String(r[0]), codigo: String(r[1]), marca: String(r[2] || ''), lista: String(r[3] || ''),
    familia: String(r[4] || ''), categoria: String(r[5] || ''), nombre: String(r[6] || ''),
    pvp: Number(r[7]) || 0, nuevo: r[8] === 'Sí', estado: String(r[9] || 'Activo'),
    oferta: String(r[10] || ''), orden: Number(r[11]) || 0,
    foto: String(r[12] || ''), fotos: String(r[13] || '').split('|').map(s => s.trim()).filter(Boolean),
    ficha: String(r[14] || ''), manual: String(r[15] || ''), video: String(r[16] || ''),
    zonas: String(r[18] || '').split(',').map(s => s.trim()).filter(Boolean),
    descripcion: String(r[19] || ''), ideal: String(r[20] || ''),
    complementos: String(r[21] || '').split('|').map(s => s.trim()).filter(Boolean),
  }));
}

function catalogo_() {
  const cfg = config_();
  // datos internos que no van al navegador
  delete cfg.hojaVentas; delete cfg.correoResumen; delete cfg.copia; delete cfg.modelo;
  return { ok: true, actualizado: new Date().toISOString(), config: cfg, productos: productos_() };
}

// ---------- Clientes por RUC ----------
function buscarRuc_(ruc) {
  ruc = String(ruc || '').replace(/\D/g, '');
  if (!/^(10|15|17|20)\d{9}$/.test(ruc)) return { ok: false, error: 'RUC no válido' };
  const sh = SpreadsheetApp.openById(HOJA_CLIENTES).getSheetByName('Clientes');
  const h = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(x => String(x).trim());
  const c = n => h.indexOf(n);
  const v = sh.getRange(2, 1, sh.getLastRow() - 1, h.length).getValues();
  const r = v.find(x => String(x[c('RUC')]).replace(/\D/g, '') === ruc);
  if (!r) return { ok: true, encontrado: false };
  // Solo datos públicos (los mismos que figuran en SUNAT). Correo y celular los escribe el cliente.
  const val = n => (c(n) >= 0 ? r[c(n)] : '');
  return {
    ok: true, encontrado: true, idCliente: r[0], razon: val('Razón social'), direccion: val('Dirección'),
    ciudad: [val('Distrito'), val('Provincia'), val('Departamento')].filter(Boolean).join(' - '), zona: val('Zona') || '',
  };
}

// ---------- Órdenes de compra ----------
function carpeta_() {
  const it = DriveApp.getFoldersByName(CARPETA_ORDENES);
  return it.hasNext() ? it.next() : DriveApp.createFolder(CARPETA_ORDENES);
}

function siguienteNumero_(sh) {
  const n = sh.getLastRow(); // fila 1 = encabezados
  const anio = Utilities.formatDate(new Date(), ZONA, 'yyyy');
  return 'TLI-' + anio + '-' + String(n).padStart(4, '0');
}

// Logo de la marca incrustado en el PDF (se descarga una vez y queda guardado 6 horas)
function logoData_() {
  const cache = CacheService.getScriptCache();
  let b64 = cache.get('logo');
  if (!b64) {
    try {
      b64 = Utilities.base64Encode(UrlFetchApp.fetch('https://daviddiazosorio73-debug.github.io/catalogo/assets/logo-pdf.png').getContent());
      if (b64.length < 95000) cache.put('logo', b64, 21600);
    } catch (e) { return ''; }
  }
  return 'data:image/png;base64,' + b64;
}

function esc_(s) {
  return String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
}

function registrarOrden_(d) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    // Validación de datos
    const ruc = String(d.ruc || '').replace(/\D/g, '');
    if (!/^(10|15|17|20)\d{9}$/.test(ruc)) return { ok: false, error: 'RUC no válido' };
    if (!d.razon || !d.email || !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(d.email)) return { ok: false, error: 'Faltan datos del cliente' };
    if (!d.firma || String(d.firma).indexOf('data:image/png;base64,') !== 0) return { ok: false, error: 'Falta la firma' };
    if (!Array.isArray(d.items) || !d.items.length) return { ok: false, error: 'El carrito está vacío' };

    // Precios desde la base maestra
    const cat = {};
    productos_().forEach(p => cat[p.id] = p);
    const lineas = [];
    let total = 0;
    d.items.forEach(it => {
      const p = cat[it.id];
      const q = Math.max(1, Math.min(999, parseInt(it.cantidad, 10) || 0));
      if (!p || !p.pvp) return;
      const sub = Math.round(p.pvp * q * 100) / 100;
      total += sub;
      lineas.push({ id: p.id, codigo: p.codigo, marca: p.marca, nombre: p.nombre, estado: p.estado, cantidad: q, precio: p.pvp, subtotal: sub });
    });
    if (!lineas.length) return { ok: false, error: 'Ningún producto válido en el carrito' };
    total = Math.round(total * 100) / 100;

    const cfg = config_();
    const ss = SpreadsheetApp.openById(BASE_MAESTRA);
    const shPed = ss.getSheetByName('Pedidos');
    const numero = siguienteNumero_(shPed);
    const ahora = new Date();
    const fecha = Utilities.formatDate(ahora, ZONA, 'dd/MM/yyyy HH:mm');
    const vence = Utilities.formatDate(new Date(ahora.getTime() + cfg.validez * 864e5), ZONA, 'dd/MM/yyyy');
    const cli = buscarRuc_(ruc);
    const igv = Math.round((total - total / 1.18) * 100) / 100;

    // Firma
    const carpeta = carpeta_();
    const firmaBlob = Utilities.newBlob(Utilities.base64Decode(d.firma.split(',')[1]), 'image/png', numero + '_firma.png');
    const firmaFile = carpeta.createFile(firmaBlob);
    const firmaReg = registrarFirma_(ss, ruc, d.razon, d.contacto, numero, firmaFile.getUrl(), d.firma);

    // PDF de la orden
    const money = n => 'S/ ' + Number(n).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    const filas = lineas.map((l, i) => `<tr><td>${i + 1}</td><td class="n">${esc_(l.codigo)}</td><td>${esc_(l.marca)}</td><td>${esc_(l.nombre)}${l.estado === 'Próximo ingreso' ? ' <i>(próximo ingreso)</i>' : ''}</td><td class="r">${l.cantidad}</td><td class="r">${money(l.precio)}</td><td class="r">${money(l.subtotal)}</td></tr>`).join('');
    const html = `<html><head><meta charset="utf-8"><style>
      body{font-family:Arial,Helvetica,sans-serif;color:#1b1d26;font-size:11px;margin:28px}
      h1{font-size:20px;margin:0;color:#14161f} .muted{color:#666} .r{text-align:right}
      .top{width:100%;border-bottom:3px solid #FF2D55;margin-bottom:14px}.top td{border:none;padding:0 0 10px 0}
      .tag{display:inline-block;background:#FFD24A;color:#3A2600;font-weight:bold;padding:3px 8px;border-radius:4px;font-size:10px}
      table{width:100%;border-collapse:collapse;margin-top:10px} th{background:#14161f;color:#fff;text-align:left;padding:6px;font-size:10px}
      td{border-bottom:1px solid #ddd;padding:6px;vertical-align:top} .tot td{border:none;font-size:12px}
      td.r,th.r,td.n{white-space:nowrap} .top h1{white-space:nowrap} .logo{height:46px}
      .box{border:1px solid #ccc;border-radius:6px;padding:10px;margin-top:6px} .firma{margin-top:28px;width:300px;text-align:center;page-break-inside:avoid} .firma img{height:90px;display:block;margin:0 auto} .linea{border-top:1px solid #1b1d26;margin-top:2px;padding-top:4px}
    </style></head><body>
      <table class="top"><tr><td><img class="logo" src="${logoData_()}"><br><div class="muted">Tecnología e IA para el Perú · WhatsApp ${esc_(cfg.whatsapp.replace(/^51/, ''))} · ${esc_(cfg.correo)}</div>
      <div class="tag">SOLO DISTRIBUIDORES Y MAYORISTAS</div></td>
      <td class="r"><h1>ORDEN DE COMPRA</h1><div><b>N° ${numero}</b></div><div>Fecha: ${fecha}</div><div>Válida hasta: ${vence}</div></td></tr></table>
      <div class="box"><b>Cliente:</b> ${esc_(d.razon)} &nbsp; <b>RUC:</b> ${ruc}<br>
      <b>Contacto:</b> ${esc_(d.contacto)} &nbsp; <b>Correo:</b> ${esc_(d.email)}<br>
      <b>Dirección de entrega:</b> ${esc_(d.direccion)} · ${esc_(d.ciudad)}${d.notas ? '<br><b>Notas:</b> ' + esc_(d.notas) : ''}</div>
      <table><tr><th>#</th><th>Código</th><th>Marca</th><th>Producto</th><th class="r">Cant.</th><th class="r">P. unit.</th><th class="r">Subtotal</th></tr>${filas}
      <tr class="tot"><td colspan="6" class="r">Op. gravada</td><td class="r">${money(total - igv)}</td></tr>
      <tr class="tot"><td colspan="6" class="r">IGV (18%)</td><td class="r">${money(igv)}</td></tr>
      <tr class="tot"><td colspan="6" class="r"><b>TOTAL</b></td><td class="r"><b>${money(total)}</b></td></tr></table>
      <p class="muted">Precios en soles, incluyen IGV (${esc_(cfg.precios)} · ${esc_(cfg.vigencia)}). Sujetos a stock y confirmación de ${esc_(cfg.nombre)}. Los productos marcados como «próximo ingreso» se despachan a su llegada.</p>
      <table style="width:100%;margin-top:30px"><tr><td style="border:none;vertical-align:bottom" class="muted">Firmado electrónicamente en ${esc_(cfg.nombre)} el ${fecha}.<br>Registro de firma: ${firmaReg.version}</td>
      <td style="border:none;width:320px"><div class="firma"><img src="${d.firma}"><div class="linea"><b>${esc_(d.contacto || d.razon)}</b><br>${esc_(d.razon)} · RUC ${ruc}<br>Firma del cliente · conformidad de la orden</div></div></td></tr></table>
    </body></html>`;
    const pdf = Utilities.newBlob(html, 'text/html', numero + '.html').getAs('application/pdf').setName(numero + ' - ' + d.razon + '.pdf');
    const pdfFile = carpeta.createFile(pdf);

    // Registro
    shPed.appendRow([numero, fecha, d.razon, "'" + ruc, d.email, "'" + d.celular, d.direccion, d.ciudad,
      lineas.map(l => l.codigo + ' × ' + l.cantidad).join(' | '), total, firmaFile.getUrl(), pdfFile.getUrl(), 'Nuevo',
      [cli.encontrado ? 'Cliente ' + cli.idCliente : 'CLIENTE NUEVO', d.origen ? 'origen: ' + d.origen : '', d.notas || ''].filter(Boolean).join(' · ')]);

    // A tu lista «Hoy» (prioridad 0) y el carrito queda como convertido
    try { avisarHoy_(ruc, 'Orden web ' + numero + ' · ' + money(total) + ': confirmar'); } catch (e) { /* sin acceso */ }
    // También a «Pedidos» de la hoja de ventas (cuenta en la venta del mes y en la revisión del viernes)
    try {
      const pv = SpreadsheetApp.openById(cfg.hojaVentas).getSheetByName('Pedidos');
      const zonaCli = (cli && cli.zona) || '';
      const filasPv = lineas.map(l => [ahora, numero, (cli && cli.idCliente) || '', d.razon, zonaCli, "'" + l.codigo, l.nombre, l.cantidad, l.precio, l.subtotal, 'Nuevo', 'App', 'Orden web · RUC ' + ruc, l.marca || '']);
      pv.getRange(pv.getLastRow() + 1, 1, filasPv.length, 14).setValues(filasPv);
    } catch (e) { /* sin acceso a la hoja de ventas */ }
    try { cerrarCarrito_(d.sesion, ruc, numero); } catch (e) { /* sin carrito */ }

    // Correo al comprador con copia
    const cuerpo = `<p>Hola ${esc_(d.contacto || d.razon)},</p>
      <p>Gracias por tu pedido. Adjuntamos tu <b>orden de compra N° ${numero}</b> por un total de <b>${money(total)}</b> (incluye IGV).</p>
      <p>Un asesor de ${esc_(cfg.nombre)} te confirmará el stock, la forma de pago y la fecha de despacho.</p>
      <p>Consultas: WhatsApp <a href="https://wa.me/${esc_(cfg.whatsapp)}">${esc_(cfg.whatsapp.replace(/^51/, ''))}</a></p>
      <p>— ${esc_(cfg.nombre)} · Solo distribuidores y mayoristas</p>`;
    MailApp.sendEmail({
      to: d.email, cc: cfg.correo, bcc: cfg.copia || undefined,
      subject: 'Orden de compra ' + numero + ' · ' + cfg.nombre,
      htmlBody: cuerpo, attachments: [pdf], name: cfg.nombre, replyTo: cfg.correo,
    });

    return { ok: true, numero: numero, total: total, cliente: cli.encontrado ? 'existente' : 'nuevo' };
  } finally {
    lock.releaseLock();
  }
}

/** Prueba rápida desde el editor: Ejecutar ▶ «probar». */
function probar() {
  const c = catalogo_();
  Logger.log('Productos visibles: ' + c.productos.length + ' · Correo: ' + c.config.correo);
}

// =====================================================================
//  IA 1 · Ventas: lista «Hoy», carritos abandonados, métricas, asesor IA
// =====================================================================

/** Pone al cliente en tu lista «Hoy» (hoja de ventas) con prioridad 0. */
function avisarHoy_(ruc, motivo) {
  const cfg = config_();
  const sh = SpreadsheetApp.openById(cfg.hojaVentas).getSheetByName('Clientes');
  const h = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(x => String(x).trim());
  const cRuc = h.indexOf('RUC') + 1, cAsig = h.indexOf('Asignado el') + 1, cMot = h.indexOf('Motivo del día') + 1;
  if (!cRuc || !cAsig || !cMot) return false;
  const n = sh.getLastRow() - 1;
  const rucs = sh.getRange(2, cRuc, n, 1).getValues();
  const i = rucs.findIndex(r => String(r[0]).replace(/\D/g, '') === String(ruc));
  if (i < 0) return false; // cliente nuevo: queda en Pedidos/Carritos y en el resumen diario
  // Fecha y hora actuales: así aparece aunque ya lo hayas llamado hoy más temprano
  sh.getRange(i + 2, cAsig).setValue(new Date());
  sh.getRange(i + 2, cMot).setValue('0 ' + motivo);
  return true;
}

// ---------- Firmas: se guarda cada firma; si es distinta a la anterior del mismo RUC, es una nueva versión ----------
function registrarFirma_(ss, ruc, razon, firmante, doc, url, dataUrl) {
  const sh = ss.getSheetByName('Firmas');
  if (!sh) return { version: doc };
  const huella = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(dataUrl))).slice(0, 16);
  const prev = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 8).getValues().filter(r => String(r[1]).replace(/\D/g, '') === ruc) : [];
  const igual = prev.filter(r => r[6] === huella)[0];
  const version = igual ? String(igual[7]).split(' ')[0] : 'F' + (prev.length ? Math.max.apply(null, prev.map(r => Number(String(r[7]).replace(/\D/g, '')) || 0)) + 1 : 1);
  sh.appendRow([new Date(), "'" + ruc, razon, firmante || '', doc, url, huella, version + (igual ? ' (misma firma)' : ' (nueva)')]);
  return { version: ruc + '-' + version };
}

// ---------- Cotizaciones web: el cliente nuevo entra a «Clientes» como «Nuevo» ----------
function registrarCotizacion_(d) {
  const ruc = String(d.ruc || '').replace(/\D/g, '');
  if (!/^(10|15|17|20)\d{9}$/.test(ruc) || !Array.isArray(d.items)) return { ok: false, error: 'Datos incompletos' };
  const cat = {}; productos_().forEach(p => cat[p.id] = p);
  let total = 0;
  const lineas = d.items.slice(0, 60).map(it => {
    const p = cat[it.id]; const q = Math.max(1, Math.min(999, parseInt(it.cantidad, 10) || 1));
    if (!p) return null; total += p.pvp * q; return p.codigo + ' × ' + q;
  }).filter(Boolean);
  total = Math.round(total * 100) / 100;
  const numero = String(d.numero || '').slice(0, 30);
  const nota = 'Cotización web ' + numero + ' · S/ ' + total.toFixed(2) + ': ' + lineas.join(', ');
  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  let r;
  try {
    r = altaCliente_(SpreadsheetApp.openById(config_().hojaVentas).getSheetByName('Clientes'), {
      ruc, razon: d.razon, contacto: d.contacto, celular: d.celular, email: d.email, ciudad: d.ciudad, notas: d.notas,
      accion: 'Llamar: cotización web ' + numero,
    }, nota.slice(0, 400), 'Cotización web ' + numero + ' · S/ ' + total.toFixed(2) + ': llamar y cerrar');
  } finally { lock.releaseLock(); }
  try { SpreadsheetApp.openById(BASE_MAESTRA).getSheetByName('Cotizaciones').appendRow([new Date(), numero, 'Web', "'" + ruc, String(d.razon || ''), lineas.join(' | '), total, 'Abierta', String(d.origen || '').slice(0, 20), r.nuevo ? 'Cliente nuevo ' + r.id : r.id]); } catch (e) { /* */ }
  SpreadsheetApp.openById(BASE_MAESTRA).getSheetByName('Eventos').appendRow([new Date(), String(d.sesion || '').slice(0, 40), 'cotizacion',
    (numero + ' · ' + (r.nuevo ? 'cliente nuevo ' : '') + r.id).slice(0, 120), total, "'" + ruc, String(d.origen || '').slice(0, 20)]);
  return { ok: true, nuevo: r.nuevo, id: r.id };
}

// ---------- Carritos ----------
function hojaCarritos_() { return SpreadsheetApp.openById(BASE_MAESTRA).getSheetByName('Carritos'); }

function registrarCarrito_(d) {
  const ruc = String(d.ruc || '').replace(/\D/g, '');
  const sesion = String(d.sesion || '').slice(0, 40);
  if (!sesion || !/^(10|15|17|20)\d{9}$/.test(ruc) || !Array.isArray(d.items)) return { ok: false };
  const cat = {}; productos_().forEach(p => cat[p.id] = p);
  let total = 0;
  const lineas = d.items.slice(0, 60).map(it => {
    const p = cat[it.id]; const q = Math.max(1, Math.min(999, parseInt(it.cantidad, 10) || 1));
    if (!p) return null; total += p.pvp * q; return p.codigo + ' × ' + q;
  }).filter(Boolean);
  if (!lineas.length) return { ok: false };
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const sh = hojaCarritos_();
    const n = sh.getLastRow() - 1;
    const v = n > 0 ? sh.getRange(2, 1, n, 8).getValues() : [];
    const i = v.findIndex(r => r[1] === sesion && r[7] === 'Abierto');
    const fila = [new Date(), sesion, "'" + ruc, String(d.razon || '').slice(0, 120), "'" + String(d.celular || '').slice(0, 20),
      lineas.join(' | '), Math.round(total * 100) / 100, 'Abierto', '', '', String(d.origen || '').slice(0, 20)];
    if (i >= 0) sh.getRange(i + 2, 1, 1, 11).setValues([fila]); else sh.appendRow(fila);
  } finally { lock.releaseLock(); }
  return { ok: true };
}

function cerrarCarrito_(sesion, ruc, numero) {
  const sh = hojaCarritos_();
  const n = sh.getLastRow() - 1; if (n < 1) return;
  const v = sh.getRange(2, 1, n, 8).getValues();
  v.forEach((r, i) => {
    if (r[7] === 'Abierto' && ((sesion && r[1] === sesion) || String(r[2]).replace(/\D/g, '') === String(ruc))) {
      sh.getRange(i + 2, 8, 1, 2).setValues([['Convertido', numero]]);
    }
  });
}

/** Cada hora: los carritos sin orden después de N horas pasan a tu lista «Hoy». */
function revisarCarritos() {
  const cfg = config_();
  const sh = hojaCarritos_();
  const n = sh.getLastRow() - 1; if (n < 1) return;
  const v = sh.getRange(2, 1, n, 10).getValues();
  const limite = Date.now() - cfg.horasCarrito * 3600e3;
  v.forEach((r, i) => {
    if (r[7] !== 'Abierto' || !(r[0] instanceof Date) || r[0].getTime() > limite) return;
    const ruc = String(r[2]).replace(/\D/g, '');
    const enHoy = avisarHoy_(ruc, 'Carrito abandonado S/ ' + Number(r[6]).toFixed(2) + ': llamar y cerrar');
    sh.getRange(i + 2, 8, 1, 3).setValues([['Abandonado', '', enHoy ? 'Sí' : 'Cliente nuevo (ver resumen)']]);
  });
}

// ---------- Métricas ----------
function registrarEventos_(d) {
  if (!Array.isArray(d.eventos) || !d.eventos.length) return { ok: true };
  const sesion = String(d.sesion || '').slice(0, 40), ruc = String(d.ruc || '').replace(/\D/g, '').slice(0, 11);
  const origen = String(d.origen || '').slice(0, 20);
  const filas = d.eventos.slice(0, 50).map(e => [
    new Date(Number(e.t) || Date.now()), sesion, String(e.tipo || '').slice(0, 20),
    String(e.dato || '').slice(0, 120), e.n === undefined ? '' : Number(e.n) || 0, ruc ? "'" + ruc : '', origen]);
  const sh = SpreadsheetApp.openById(BASE_MAESTRA).getSheetByName('Eventos');
  sh.getRange(sh.getLastRow() + 1, 1, filas.length, 7).setValues(filas);
  return { ok: true };
}

// ---------- Asesor IA (Google Gemini) ----------
function catalogoTexto_() {
  const cache = CacheService.getScriptCache();
  const partes = [];
  for (let k = 0; k < 8; k++) { const c = cache.get('cat_' + k); if (c === null) { partes.length = 0; break; } partes.push(c); }
  if (partes.length) return partes.join('');
  const txt = productos_().filter(p => p.estado !== 'Agotado').map(p =>
    [p.id, p.marca, p.categoria, p.nombre, 'S/' + p.pvp, p.estado === 'Próximo ingreso' ? 'próximo ingreso' : ''].join(' | ')).join('\n');
  const trozo = Math.ceil(txt.length / 8);
  try { for (let k = 0; k < 8; k++) cache.put('cat_' + k, txt.slice(k * trozo, (k + 1) * trozo), 1800); } catch (e) { /* caché llena */ }
  return txt;
}

function asesor_(d) {
  const cfg = config_();
  const clave = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!clave || !cfg.asesor) return { ok: false, error: 'El asesor IA no está activo.' };
  // Límite de uso: 25 consultas por sesión al día y 400 en total al día
  const cache = CacheService.getScriptCache();
  const dia = Utilities.formatDate(new Date(), ZONA, 'yyyyMMdd');
  const kS = 'as_' + dia + '_' + String(d.sesion || 'x').slice(0, 40), kT = 'as_' + dia;
  const nS = Number(cache.get(kS) || 0), nT = Number(cache.get(kT) || 0);
  if (nS >= 25 || nT >= 400) return { ok: false, error: 'Llegaste al límite de consultas de hoy. Escríbenos por WhatsApp.' };
  cache.put(kS, String(nS + 1), 21600); cache.put(kT, String(nT + 1), 21600);

  const historial = (Array.isArray(d.historial) ? d.historial : []).slice(-6).map(m => ({
    role: m.rol === 'asesor' ? 'model' : 'user', parts: [{ text: String(m.texto || '').slice(0, 1500) }] }));
  const pregunta = String(d.pregunta || '').slice(0, 1500);
  if (!pregunta) return { ok: false, error: 'Escribe tu consulta.' };

  const sistema = `Eres el asesor comercial de ${cfg.nombre} (Tecnología e IA para el Perú). Atiendes SOLO a distribuidores, mayoristas y negocios.
Recomiendas equipos y productos EXCLUSIVAMENTE del catálogo de abajo (formato: id | marca | categoría | producto | PVP con IGV | estado).
Reglas:
- Usa solo productos del catálogo, con su id exacto. Nunca inventes productos, precios, medidas, potencias ni garantías.
- Si piden datos técnicos que no están en el catálogo, di que están en la ficha técnica del producto o que un asesor los confirma por WhatsApp.
- Arma soluciones completas para el negocio del cliente (por ejemplo, una heladería, pollería, panadería o bodega), con cantidades razonables.
- Respuestas cortas, claras y amables, en español del Perú. Precios en soles con IGV (PVP).
- Responde SIEMPRE en JSON: {"mensaje": "texto para el cliente", "productos": [{"id": "...", "cantidad": 1, "motivo": "por qué"}]} (máximo 8 productos; lista vacía si no aplica).
CATÁLOGO:
${catalogoTexto_()}`;

  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(cfg.modelo) + ':generateContent';
  const cuerpo = {
    systemInstruction: { parts: [{ text: sistema }] },
    contents: historial.concat([{ role: 'user', parts: [{ text: pregunta }] }]),
    generationConfig: { temperature: 0.3, responseMimeType: 'application/json', maxOutputTokens: 4096, thinkingConfig: { thinkingBudget: 0 } },
  };
  const r = UrlFetchApp.fetch(url, { method: 'post', contentType: 'application/json', headers: { 'x-goog-api-key': clave }, payload: JSON.stringify(cuerpo), muteHttpExceptions: true });
  if (r.getResponseCode() !== 200) return { ok: false, error: 'El asesor no está disponible en este momento (' + r.getResponseCode() + ').' };
  let out;
  try {
    const txt = JSON.parse(r.getContentText()).candidates[0].content.parts.map(p => p.text || '').join('');
    out = JSON.parse(txt);
  } catch (e) { return { ok: false, error: 'No entendí la respuesta del asesor. Intenta de nuevo.' }; }
  // Solo productos que existen y están a la venta
  const cat = {}; productos_().forEach(p => cat[p.id] = p);
  const productos = (Array.isArray(out.productos) ? out.productos : [])
    .filter(x => cat[x.id] && cat[x.id].estado !== 'Agotado').slice(0, 8)
    .map(x => ({ id: x.id, cantidad: Math.max(1, Math.min(99, parseInt(x.cantidad, 10) || 1)), motivo: String(x.motivo || '').slice(0, 200) }));
  try {
    SpreadsheetApp.openById(BASE_MAESTRA).getSheetByName('Eventos').appendRow([new Date(), String(d.sesion || '').slice(0, 40), 'asesor', pregunta.slice(0, 120), productos.length, '', String(d.origen || '')]);
  } catch (e) { /* sin registro */ }
  return { ok: true, mensaje: String(out.mensaje || '').slice(0, 2000), productos: productos };
}

// ---------- Resumen diario (7 a. m.) ----------
function resumenDiario() {
  const cfg = config_();
  const para = cfg.correoResumen || cfg.correo;
  const ss = SpreadsheetApp.openById(BASE_MAESTRA);
  const desde = Date.now() - 24 * 3600e3;
  const filas = (nombre, cols) => { const sh = ss.getSheetByName(nombre); const n = sh.getLastRow() - 1; return n > 0 ? sh.getRange(2, 1, n, cols).getValues() : []; };
  const money = n => 'S/ ' + Number(n || 0).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const fecha = v => { if (v instanceof Date) return v.getTime(); const m = String(v).match(/(\d{2})\/(\d{2})\/(\d{4}) (\d{2}):(\d{2})/); return m ? new Date(+m[3], +m[2] - 1, +m[1], +m[4], +m[5]).getTime() : 0; };

  const ped = filas('Pedidos', 14).filter(r => fecha(r[1]) >= desde);
  const car = filas('Carritos', 11).filter(r => fecha(r[0]) >= desde - 24 * 3600e3 && r[7] !== 'Convertido');
  const ev = filas('Eventos', 7).filter(r => fecha(r[0]) >= desde);
  const cuenta = (lista, f) => { const m = {}; lista.forEach(x => { const k = f(x); if (k) m[k] = (m[k] || 0) + 1; }); return Object.entries(m).sort((a, b) => b[1] - a[1]).slice(0, 10); };
  const busq = cuenta(ev.filter(e => e[2] === 'busqueda'), e => String(e[3]).toLowerCase());
  const sinRes = cuenta(ev.filter(e => e[2] === 'sin_resultado'), e => String(e[3]).toLowerCase());
  const vistos = cuenta(ev.filter(e => e[2] === 'vista'), e => e[3]);
  const sesiones = new Set(ev.map(e => e[1])).size;
  const origenes = cuenta(ev, e => e[6]);
  const asesor = ev.filter(e => e[2] === 'asesor');
  const li = arr => arr.length ? '<ul>' + arr.map(([k, n]) => `<li>${esc_(k)} <b>(${n})</b></li>`).join('') + '</ul>' : '<p>—</p>';

  const html = `<h2>${esc_(cfg.nombre)} · Resumen del ${Utilities.formatDate(new Date(), ZONA, 'dd/MM/yyyy')}</h2>
    <p><b>${ped.length}</b> orden(es) web por <b>${money(ped.reduce((a, r) => a + Number(r[9] || 0), 0))}</b> ·
    <b>${car.length}</b> carrito(s) sin cerrar · <b>${sesiones}</b> visita(s) al catálogo · <b>${asesor.length}</b> consulta(s) al asesor IA.</p>
    <h3>Órdenes de las últimas 24 h (confirmar hoy)</h3>${ped.length ? '<ul>' + ped.map(r => `<li>${esc_(r[0])} · ${esc_(r[2])} · ${money(r[9])} · ${esc_(r[12])}</li>`).join('') + '</ul>' : '<p>—</p>'}
    <h3>Carritos abandonados (llamar)</h3>${car.length ? '<ul>' + car.map(r => `<li>${esc_(r[3])} (RUC ${esc_(r[2])}) · cel. ${esc_(r[4])} · ${money(r[6])} · ${esc_(r[5])}</li>`).join('') + '</ul>' : '<p>—</p>'}
    <h3>Lo que más buscaron</h3>${li(busq)}
    <h3>Buscaron y NO encontraron (oportunidad para nuevos proveedores)</h3>${li(sinRes)}
    <h3>Productos más vistos</h3>${li(vistos)}
    <h3>De dónde llegaron</h3>${li(origenes)}
    ${asesor.length ? '<h3>Qué le preguntaron al asesor IA</h3><ul>' + asesor.slice(-10).map(e => `<li>${esc_(e[3])}</li>`).join('') + '</ul>' : ''}
    <p style="color:#666">Los clientes con orden o carrito ya están en tu lista «Hoy» con prioridad 0.</p>`;
  MailApp.sendEmail({ to: para, subject: cfg.nombre + ' · Resumen del catálogo ' + Utilities.formatDate(new Date(), ZONA, 'dd/MM'), htmlBody: html, name: cfg.nombre });
}

/** Ejecutar UNA vez: crea el aviso de carritos (cada hora) y el resumen diario (7 a. m.). */
function instalar() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (['revisarCarritos', 'resumenDiario'].indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('revisarCarritos').timeBased().everyHours(1).create();
  ScriptApp.newTrigger('resumenDiario').timeBased().atHour(7).nearMinute(0).inTimezone(ZONA).everyDays(1).create();
  Logger.log('Listo: avisos de carritos cada hora y resumen diario a las 7 a. m. Asesor IA: ' + (PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY') ? 'con clave' : 'falta GEMINI_API_KEY'));
}

// =====================================================================
//  Alta de clientes nuevos (Panel y cotizaciones web). Mismo código en el Motor y en la API.
// =====================================================================
const normT_ = s => String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').trim();
/** Zona más frecuente entre los clientes del mismo distrito / provincia / departamento. */
function zonaDe_(h, v, x) {
  const iz = h.indexOf('Zona'), cols = [['Distrito', 3], ['Provincia', 2], ['Departamento', 1]];
  const lugares = [x.distrito, x.provincia, x.departamento].concat(String(x.ciudad || '').split(/[-,/·]/)).map(normT_).filter(Boolean);
  if (iz < 0 || !lugares.length) return '';
  const votos = {};
  v.forEach(r => cols.forEach(([n, peso]) => {
    const i = h.indexOf(n); if (i < 0 || !r[iz]) return;
    if (lugares.indexOf(normT_(r[i])) >= 0) votos[r[iz]] = (votos[r[iz]] || 0) + peso;
  }));
  return Object.keys(votos).sort((a, b) => votos[b] - votos[a])[0] || '';
}
/** Crea el cliente (Etapa «Nuevo») o, si el RUC ya existe, solo agrega la nota. No toca columnas con fórmulas. */
function altaCliente_(sh, x, nota, llamarHoy) {
  const ruc = String(x.ruc || '').replace(/\D/g, '');
  if (!/^(10|15|17|20)\d{9}$/.test(ruc)) throw new Error('RUC no válido (11 dígitos)');
  if (!String(x.razon || '').trim()) throw new Error('Falta la razón social');
  const h = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(s => String(s).trim());
  const ids = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 1).getValues();
  let ultima = 0; ids.forEach((r, i) => { if (String(r[0]).trim()) ultima = i + 1; });
  const v = ultima ? sh.getRange(2, 1, ultima, h.length).getValues() : [];
  const col = n => h.indexOf(n) + 1, iRuc = h.indexOf('RUC'), iNotas = h.indexOf('Notas');
  const hoyTxt = Utilities.formatDate(new Date(), 'America/Lima', 'dd/MM');
  const k = v.findIndex(r => String(r[iRuc]).replace(/\D/g, '') === ruc);
  const marcarHoy = (fila, motivo) => {
    if (col('Asignado el')) sh.getRange(fila, col('Asignado el')).setValue(new Date());
    if (col('Motivo del día')) sh.getRange(fila, col('Motivo del día')).setValue('0 ' + motivo);
  };
  if (k >= 0) {
    if (nota && iNotas >= 0) sh.getRange(k + 2, iNotas + 1).setValue([String(v[k][iNotas] || ''), hoyTxt + ' ' + nota].filter(Boolean).join(' | ').slice(0, 1500));
    if (llamarHoy) marcarHoy(k + 2, llamarHoy);
    return { ok: true, nuevo: false, id: v[k][0], razon: v[k][h.indexOf('Razón social')] };
  }
  const num = v.reduce((m, r) => Math.max(m, Number((String(r[0]).match(/^C-(\d+)/) || [0, 0])[1])), 0) + 1;
  const id = 'C-' + String(num).padStart(4, '0');
  const cel = String(x.celular || '').replace(/\D/g, '');
  const wa = cel.length === 9 ? '51' + cel : cel;
  const zona = String(x.zona || '').trim() || zonaDe_(h, v, x) || 'Por asignar';
  const datos = {
    'ID cliente': id, 'RUC': "'" + ruc, 'Razón social': String(x.razon).trim().toUpperCase(), 'Contacto': x.contacto,
    'WhatsApp': wa ? "'" + wa : '', 'Celular 2': x.celular2 ? "'" + String(x.celular2).replace(/\D/g, '') : '', 'Teléfono fijo': x.fijo, 'Email': x.email,
    'Dirección': x.direccion, 'Distrito': x.distrito, 'Provincia': x.provincia, 'Departamento': x.departamento, 'Zona': zona,
    'Condición de pago': x.pago, 'Líneas de interés': x.lineas, 'Validación': wa ? 'Pendiente' : 'Sin celular',
    'Notas': [hoyTxt + ' ' + (nota || 'Alta de cliente'), x.notas].filter(Boolean).join(' | '),
    'Etapa': 'Nuevo', 'Próxima acción': x.accion || 'Primera llamada', 'Fecha próx. acción': new Date(), 'Prioridad (1-5)': 2,
  };
  const fila = ultima + 2;
  Object.keys(datos).forEach(n => { const c = col(n); if (c && datos[n] !== undefined && datos[n] !== '') sh.getRange(fila, c).setValue(datos[n]); });
  if (col('Fecha próx. acción')) sh.getRange(fila, col('Fecha próx. acción')).setNumberFormat('dd/mm/yyyy');
  if (llamarHoy) marcarHoy(fila, llamarHoy);
  return { ok: true, nuevo: true, id, zona };
}
