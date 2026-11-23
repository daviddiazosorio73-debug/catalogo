/**
 * Tecno Log IA · Catálogo App — API (Google Apps Script)
 *
 * Se instala en script.google.com con la cuenta tecnologiapperu@gmail.com y se publica como
 * «Aplicación web» (Ejecutar como: Yo · Quién tiene acceso: Cualquier usuario).
 *
 *  GET  ?accion=productos        → configuración + productos visibles (lee la base maestra)
 *  GET  ?accion=ruc&ruc=20xxxxxx → datos públicos del cliente (razón social, dirección, ciudad)
 *  POST {accion:'orden', ...}    → registra la orden, genera el PDF con la firma y lo envía por correo
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
    return json_({ ok: false, error: 'Acción no válida' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

// ---------- Configuración ----------
function config_() {
  const v = SpreadsheetApp.openById(BASE_MAESTRA).getSheetByName('Config').getRange('A3:B13').getValues();
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
  };
}

// ---------- Catálogo ----------
function productos_() {
  const sh = SpreadsheetApp.openById(BASE_MAESTRA).getSheetByName('Productos');
  const v = sh.getRange(2, 1, sh.getLastRow() - 1, 18).getValues();
  return v.filter(r => r[0] && r[9] !== 'Oculto').map(r => ({
    id: String(r[0]), codigo: String(r[1]), marca: String(r[2] || ''), lista: String(r[3] || ''),
    familia: String(r[4] || ''), categoria: String(r[5] || ''), nombre: String(r[6] || ''),
    pvp: Number(r[7]) || 0, nuevo: r[8] === 'Sí', estado: String(r[9] || 'Activo'),
    oferta: String(r[10] || ''), orden: Number(r[11]) || 0,
    foto: String(r[12] || ''), fotos: String(r[13] || '').split('|').map(s => s.trim()).filter(Boolean),
    ficha: String(r[14] || ''), manual: String(r[15] || ''), video: String(r[16] || ''),
  }));
}

function catalogo_() {
  return { ok: true, actualizado: new Date().toISOString(), config: config_(), productos: productos_() };
}

// ---------- Clientes por RUC ----------
function buscarRuc_(ruc) {
  ruc = String(ruc || '').replace(/\D/g, '');
  if (!/^(10|15|17|20)\d{9}$/.test(ruc)) return { ok: false, error: 'RUC no válido' };
  const sh = SpreadsheetApp.openById(HOJA_CLIENTES).getSheetByName('Clientes');
  const v = sh.getRange(2, 1, sh.getLastRow() - 1, 14).getValues();
  const r = v.find(x => String(x[1]).replace(/\D/g, '') === ruc);
  if (!r) return { ok: true, encontrado: false };
  // Solo datos públicos (los mismos que figuran en SUNAT). Correo y celular los escribe el cliente.
  return {
    ok: true, encontrado: true, idCliente: r[0], razon: r[2],
    direccion: r[9], ciudad: [r[10], r[11], r[12]].filter(Boolean).join(' - '),
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
      .box{border:1px solid #ccc;border-radius:6px;padding:10px;margin-top:6px} .firma img{height:90px}
    </style></head><body>
      <table class="top"><tr><td><img class="logo" src="${logoData_()}"><br><div class="muted">Tecnología e IA para el Perú · WhatsApp ${esc_(cfg.whatsapp.replace(/^51/, ''))} · ${esc_(cfg.correo)}</div>
      <div class="tag">SOLO DISTRIBUIDORES Y MAYORISTAS</div></td>
      <td class="r"><h1>ORDEN DE COMPRA</h1><div><b>N° ${numero}</b></div><div>Fecha: ${fecha}</div><div>Válida hasta: ${vence}</div></td></tr></table>
      <div class="box"><b>Cliente:</b> ${esc_(d.razon)} &nbsp; <b>RUC:</b> ${ruc}<br>
      <b>Contacto:</b> ${esc_(d.contacto)} &nbsp; <b>Celular:</b> ${esc_(d.celular)} &nbsp; <b>Correo:</b> ${esc_(d.email)}<br>
      <b>Dirección de entrega:</b> ${esc_(d.direccion)} · ${esc_(d.ciudad)}${d.notas ? '<br><b>Notas:</b> ' + esc_(d.notas) : ''}</div>
      <table><tr><th>#</th><th>Código</th><th>Marca</th><th>Producto</th><th class="r">Cant.</th><th class="r">P. unit.</th><th class="r">Subtotal</th></tr>${filas}
      <tr class="tot"><td colspan="6" class="r">Op. gravada</td><td class="r">${money(total - igv)}</td></tr>
      <tr class="tot"><td colspan="6" class="r">IGV (18%)</td><td class="r">${money(igv)}</td></tr>
      <tr class="tot"><td colspan="6" class="r"><b>TOTAL</b></td><td class="r"><b>${money(total)}</b></td></tr></table>
      <p class="muted">Precios en soles, incluyen IGV (${esc_(cfg.precios)} · ${esc_(cfg.vigencia)}). Sujetos a stock y confirmación de ${esc_(cfg.nombre)}. Los productos marcados como «próximo ingreso» se despachan a su llegada.</p>
      <div class="box firma"><b>Conformidad del cliente</b><br><img src="${d.firma}"><br>${esc_(d.contacto || d.razon)} · RUC ${ruc} · ${fecha}</div>
    </body></html>`;
    const pdf = Utilities.newBlob(html, 'text/html', numero + '.html').getAs('application/pdf').setName(numero + ' - ' + d.razon + '.pdf');
    const pdfFile = carpeta.createFile(pdf);

    // Registro
    shPed.appendRow([numero, fecha, d.razon, "'" + ruc, d.email, "'" + d.celular, d.direccion, d.ciudad,
      lineas.map(l => l.codigo + ' × ' + l.cantidad).join(' | '), total, firmaFile.getUrl(), pdfFile.getUrl(), 'Nuevo',
      [cli.encontrado ? 'Cliente ' + cli.idCliente : 'CLIENTE NUEVO', d.origen ? 'origen: ' + d.origen : '', d.notas || ''].filter(Boolean).join(' · ')]);

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
