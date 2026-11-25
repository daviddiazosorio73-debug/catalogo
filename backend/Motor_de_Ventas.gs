/**
 * MOTOR DE VENTAS v8 · Tecno Log IA (hoja maestra de distribuidores) + PANEL DE VENTAS (app)
 *
 * Instalación (una vez):
 * 1) script.google.com → tu proyecto del Motor → borra todo y pega este archivo (Motor.gs).
 * 2) Archivo + → HTML → nómbralo «Panel» (sin .html) → pega el archivo Panel.html.
 * 3) ⚙️ Configuración del proyecto → Propiedades del script → GEMINI_API_KEY = tu clave (la misma del catálogo).
 * 4) Elige la función «instalar» → Ejecutar (repara columnas, crea avisos y te manda un correo de prueba).
 * 5) Implementar → Nueva implementación → Aplicación web → Ejecutar como: Yo · Acceso: Solo yo → Implementar.
 *    Abre la URL en tu celular → menú ⋮ → «Agregar a la pantalla principal».
 *
 * Calendario: llamadas de lunes a jueves · viernes revisión (pedidos, envíos, facturas) · sábado y domingo descanso.
 */

const ID_HOJA = '12gaQxik4JuXN6cXUf7AwoQEWBSoEMwf6Drj8sAywh-s';
const BASE_MAESTRA = '1_XtTBAGbIJFcbP6tLNTU4cADwF1eKN59WaX1olJkvOE'; // órdenes web (si tu cuenta tiene acceso)
const CATALOGO_URL = 'https://daviddiazosorio73-debug.github.io/catalogo/';
const CATALOGO_API = 'https://script.google.com/macros/s/AKfycby_n087GKueloyFmtqNp7u-LcsRmKk9m3fZ70deCms2PhRr8vL4qsZL0oG9Jw5xOUs9/exec';
const PANEL_URL = CATALOGO_URL + 'panel/'; // ícono y acceso directo del Panel
const DIAS_PARA_SEGUIMIENTO = 3;
const ZONA = 'America/Lima';
const DIAS_ES = ['', 'Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado', 'Domingo'];

function hoja_() { return SpreadsheetApp.openById(ID_HOJA); }
/** Correo para los avisos: Config «Correo de avisos» o la cuenta donde está instalado el motor. */
function correo_() { try { const c = String(cfgMap_(hoja_())['Correo de avisos'] || '').trim(); if (/@/.test(c)) return c; } catch (e) { /* */ } return Session.getEffectiveUser().getEmail(); }
function diaSemana_() { return Number(Utilities.formatDate(new Date(), ZONA, 'u')); } // 1 = lunes … 7 = domingo
function hoyCero_() { const h = new Date(); h.setHours(0, 0, 0, 0); return h; }
function cfgMap_(ss) {
  const v = ss.getSheetByName('Config').getRange('A1:B40').getValues();
  const m = {}; v.forEach(r => { if (r[0]) m[String(r[0]).trim()] = r[1]; }); return m;
}

/** Fechas: acepta Date, «dd/mm/aaaa hh:mm», «aaaa-mm-dd hh:mm» y «mm/dd/aaaa» no ambiguas. */
function aFecha_(x) {
  if (x instanceof Date && !isNaN(x)) return x;
  const s = String(x || '').trim(); let m;
  if ((m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{2}))?/))) return new Date(+m[1], +m[2] - 1, +m[3], +(m[4] || 0), +(m[5] || 0));
  if ((m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?: (\d{1,2}):(\d{2}))?/))) return new Date(+m[3], +m[2] - 1, +m[1], +(m[4] || 0), +(m[5] || 0));
  return null;
}
const dia0_ = d => { const x = new Date(d); x.setHours(0, 0, 0, 0); return x; };
const fmt_ = (d, p) => d ? Utilities.formatDate(d, ZONA, p || 'dd/MM/yyyy') : '';

/** Columnas de «Clientes» por su encabezado (no se rompe si agregas o mueves columnas). */
function colsClientes_(cl) {
  const h = cl.getRange(1, 1, 1, cl.getLastColumn()).getValues()[0].map(x => String(x).trim());
  const c = n => { const i = h.indexOf(n); if (i < 0) throw new Error('Falta la columna «' + n + '» en Clientes'); return i + 1; };
  const o = n => h.indexOf(n) + 1; // opcional (0 si no existe)
  return {
    h, id: 1, ruc: c('RUC'), razon: c('Razón social'), contacto: o('Contacto'), wa: c('WhatsApp'), cel2: o('Celular 2'), fijo: o('Teléfono fijo'),
    email: o('Email'), dir: o('Dirección'), distrito: o('Distrito'), provincia: o('Provincia'), depto: o('Departamento'), zona: c('Zona'),
    pago: o('Condición de pago'), linea: o('Línea de crédito'), moneda: o('Moneda'), lineas: o('Líneas de interés'),
    validacion: o('Validación'), notas: o('Notas') || c('Validación'), etapa: c('Etapa'), ultCompra: o('Última compra'), dias: c('Días sin compra'),
    accion: c('Próxima acción'), fecha: c('Fecha próx. acción'), pri: c('Prioridad (1-5)'),
    ult: c('Fecha último contacto'), mes: c('Contactos este mes'), asig: c('Asignado el'), motivo: c('Motivo del día'),
  };
}

// =====================================================================
//  Instalación y reparación
// =====================================================================
function instalar() {
  repararClientes();
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('resumenDiario').timeBased().everyDays(1).atHour(7).inTimezone(ZONA).create();
  ScriptApp.newTrigger('alEditar').forSpreadsheet(ID_HOJA).onEdit().create();
  Logger.log('Listo: lun-jue llamadas 7 a. m., viernes revisión, registro automático. IA: ' + (clave_() ? 'con clave' : 'falta GEMINI_API_KEY'));
  resumenDiario(true); // correo de prueba
}

/**
 * Repara las filas que la versión anterior escribió una columna a la izquierda
 * (pasó cuando se insertó una columna en Clientes): «Próxima acción» quedó en «Días sin compra»,
 * la fecha en «Próxima acción» y la etapa en «Notas». También restaura la fórmula de «Días sin compra».
 */
function repararClientes() {
  const ss = hoja_(), cl = ss.getSheetByName('Clientes'), C = colsClientes_(cl);
  const n = cl.getLastRow() - 1; if (n < 1) return;
  const ancho = cl.getLastColumn();
  const v = cl.getRange(2, 1, n, ancho).getValues();
  const ETAPAS = ['Cotizado', 'Validado', 'Cliente activo', 'Inactivo', 'Por validar', 'No interesado'];
  const orden = e => ['Por validar', 'Validado', 'Cotizado', 'Cliente activo'].indexOf(e);
  let arreglados = 0;
  const notas = [], etapas = [], acc = [], fec = [];
  v.forEach(r => {
    let nota = r[C.notas - 1], etapa = r[C.etapa - 1], a = r[C.accion - 1], f = r[C.fecha - 1];
    const dias = r[C.dias - 1];
    if (typeof dias === 'string' && dias.trim() && isNaN(Number(dias))) { // acción escrita en «Días sin compra»
      a = dias; f = aFecha_(r[C.accion - 1]) || r[C.accion - 1]; arreglados++;
    }
    if (ETAPAS.indexOf(String(nota).trim()) >= 0) { // etapa escrita en «Notas»
      if (orden(String(nota).trim()) > orden(etapa)) etapa = String(nota).trim();
      nota = '';
    }
    notas.push([nota]); etapas.push([etapa]); acc.push([a]); fec.push([f]);
  });
  cl.getRange(2, C.notas, n, 1).setValues(notas);
  cl.getRange(2, C.etapa, n, 1).setValues(etapas);
  cl.getRange(2, C.accion, n, 1).setValues(acc);
  cl.getRange(2, C.fecha, n, 1).setValues(fec).setNumberFormat('dd/mm/yyyy');
  // Fórmula de «Días sin compra»
  const L = c => { let s = ''; while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); } return s; };
  if (C.ultCompra) {
    cl.getRange(2, C.dias, n, 1).clearContent();
    const u = L(C.ultCompra);
    cl.getRange(2, C.dias).setFormula('=ARRAYFORMULA(IF(A2:A="",,IF(' + u + '2:' + u + '="",,TODAY()-' + u + '2:' + u + ')))');
  }
  Logger.log('Clientes reparados: ' + arreglados + ' filas con la acción corrida.');
}

// =====================================================================
//  Registro desde la hoja (onEdit) y desde el Panel
// =====================================================================
function alEditar(e) {
  const h = e.range.getSheet(), fila = e.range.getRow(), col = e.range.getColumn();
  if (fila < 2) return;
  if (h.getName() === 'Pedidos' && col <= 8) {
    const idCliente = h.getRange(fila, 3).getValue();
    if (!idCliente) return;
    if (!h.getRange(fila, 1).getValue()) h.getRange(fila, 1).setValue(new Date());
    if (!h.getRange(fila, 11).getValue()) h.getRange(fila, 11).setValue('Nuevo');
    if (!h.getRange(fila, 2).getValue()) h.getRange(fila, 2).setValue(numeroPedido_(h, idCliente, fila));
    marcarClienteActivo_(e.source, idCliente);
  }
  if (h.getName() === 'Llamadas' && col >= 2 && col <= 9) registrarLlamadaHoja_(e.source, h, fila, col);
  if (h.getName() === 'Clientes') {
    const C = colsClientes_(h);
    if (col !== C.accion) return;
    const accion = String(e.range.getValue()).trim();
    if (!accion) { h.getRange(fila, C.fecha).clearContent(); return; }
    const prox = new Date(); prox.setDate(prox.getDate() + DIAS_PARA_SEGUIMIENTO);
    if (!h.getRange(fila, C.fecha).getValue()) h.getRange(fila, C.fecha).setValue(prox);
    if (h.getRange(fila, C.etapa).getValue() === 'Por validar') h.getRange(fila, C.etapa).setValue('Validado');
    const nota = h.getRange(fila, C.notas).getValue();
    h.getRange(fila, C.notas).setValue((nota ? nota + ' | ' : '') + fmt_(new Date(), 'dd/MM') + ' contacto: ' + accion);
  }
}

/** Resultado → [etapa nueva, próxima acción sugerida, días hasta ella, canal] */
const REGLAS = {
  'Pedido tomado':               ['Cliente activo', 'Confirmar despacho y pago', 2, 'Llamada'],
  'Interesado - cotizar':        ['Cotizado', 'Enviar cotización', 1, 'Llamada'],
  'Interesado - pidió catálogo': ['Cotizado', 'Enviar catálogo', 0, 'Llamada'],
  'Catálogo enviado':            ['Cotizado', 'Confirmar que revisó el catálogo y tomar pedido', 2, 'Llamada'],
  'Volver a llamar':             [null, 'Volver a llamar', 2, 'Llamada'],
  'Ocupad@':                     [null, 'Llamar de nuevo (estaba ocupad@)', 1, 'Llamada'],
  'WhatsApp':                    [null, 'Seguir la conversación por WhatsApp', 1, 'WhatsApp'],
  'Difusión':                    [null, 'Confirmar que vio la difusión y tomar pedido', 2, 'WhatsApp'],
  'No por ahora':                ['Validado', 'Ofrecer novedades del mes', 30, 'Llamada'],
  'No contesta':                 [null, 'Reintentar llamada', 1, 'Llamada'],
  'Número equivocado':           [null, 'Buscar número correcto', 3, 'Llamada'],
  'Cerró / cambió de rubro':     ['Inactivo', '', 0, 'Llamada'],
};
const REINTENTO = ['Reintentar llamada', 'Llamar de nuevo (estaba ocupad@)'];

/** Fecha sugerida: salta viernes, sábado y domingo (las llamadas son de lunes a jueves). */
function diaHabil_(dias) {
  const f = hoyCero_(); f.setDate(f.getDate() + dias);
  while ([5, 6, 0].indexOf(f.getDay()) >= 0 && dias > 0) f.setDate(f.getDate() + 1);
  return f;
}

function registrarLlamadaHoja_(ss, h, fila, col) {
  const v = h.getRange(fila, 1, 1, 9).getValues()[0];
  const id = v[1], res = v[5];
  if (!id) return;
  if (!v[0]) h.getRange(fila, 1).setValue(new Date());
  if (!v[4]) h.getRange(fila, 5).setValue(REGLAS[res] ? REGLAS[res][3] : 'Llamada');
  if (!res || !REGLAS[res]) return;
  const [, accionSug, dias] = REGLAS[res];
  let accion = v[7], fecha = v[8];
  if (!accion && accionSug) { accion = accionSug; h.getRange(fila, 8).setValue(accion); }
  if (!fecha && accionSug) { fecha = diaHabil_(dias); h.getRange(fila, 9).setValue(fecha); }
  aplicarResultado_(ss, id, res, col === 7 ? v[6] : '', accion, fecha);
}

/** Actualiza la ficha del cliente con el resultado de un contacto. */
function aplicarResultado_(ss, id, res, comentario, accion, fecha) {
  const cl = ss.getSheetByName('Clientes'), C = colsClientes_(cl);
  const ids = cl.getRange(2, 1, cl.getLastRow() - 1, 1).getValues().flat();
  const i = ids.indexOf(id); if (i < 0) return;
  const f = i + 2;
  const etapa = REGLAS[res] ? REGLAS[res][0] : null;
  const actual = cl.getRange(f, C.etapa).getValue();
  let nueva = etapa || (actual === 'Por validar' && ['Volver a llamar', 'WhatsApp', 'Difusión', 'Ocupad@'].indexOf(res) >= 0 ? 'Validado' : actual);
  if (actual === 'Cliente activo' && (nueva === 'Validado' || nueva === 'Cotizado')) nueva = actual; // no bajar a un cliente que ya compra
  cl.getRange(f, C.etapa).setValue(nueva);
  cl.getRange(f, C.accion).setValue(accion || '');
  cl.getRange(f, C.fecha).setValue(fecha || '');
  if (comentario) {
    const nota = cl.getRange(f, C.notas).getValue();
    cl.getRange(f, C.notas).setValue((nota ? nota + ' | ' : '') + fmt_(new Date(), 'dd/MM') + ' ' + res + ': ' + comentario);
  }
}

function numeroPedido_(h, idCliente, fila) {
  if (fila > 2) {
    const ant = h.getRange(fila - 1, 1, 1, 3).getValues()[0];
    const hoy = fmt_(new Date(), 'yyyyMMdd');
    if (ant[2] === idCliente && ant[0] && fmt_(new Date(ant[0]), 'yyyyMMdd') === hoy) return ant[1];
  }
  return 'P-' + fmt_(new Date(), 'yyMMdd-HHmmss');
}

function marcarClienteActivo_(ss, idCliente) {
  const cl = ss.getSheetByName('Clientes');
  const ids = cl.getRange(2, 1, cl.getLastRow() - 1, 1).getValues().flat();
  const i = ids.indexOf(idCliente);
  if (i >= 0) cl.getRange(i + 2, colsClientes_(cl).etapa).setValue('Cliente activo');
}

// =====================================================================
//  Lectura de historial: llamadas, pedidos, órdenes web
// =====================================================================
function leerLlamadas_(ss) {
  const sh = ss.getSheetByName('Llamadas'); const n = sh.getLastRow() - 1;
  const m = {}; if (n < 1) return m;
  sh.getRange(2, 1, n, 9).getValues().forEach(r => {
    if (!r[1]) return; const f = aFecha_(r[0]); if (!f) return;
    (m[r[1]] = m[r[1]] || []).push({ f, canal: r[4], res: r[5], com: r[6], acc: r[7], prox: aFecha_(r[8]) });
  });
  Object.values(m).forEach(l => l.sort((a, b) => a.f - b.f));
  return m;
}

function leerPedidos_(ss) {
  const sh = ss.getSheetByName('Pedidos'); const n = sh.getLastRow() - 1;
  const m = {}; if (n < 1) return m;
  const porNum = {};
  sh.getRange(2, 1, n, 14).getValues().forEach((r, i) => {
    if (!r[2] || String(r[10]) === 'Anulado') return;
    const k = r[1] || ('fila' + i);
    const p = porNum[k] = porNum[k] || { id: r[2], num: r[1], fecha: aFecha_(r[0]), estado: r[10], canal: r[11], items: [], total: 0 };
    p.items.push((r[7] || 1) + ' × ' + (r[6] || r[5])); p.total += Number(r[9]) || 0;
  });
  Object.values(porNum).forEach(p => (m[p.id] = m[p.id] || []).push(p));
  Object.values(m).forEach(l => l.sort((a, b) => (b.fecha || 0) - (a.fecha || 0)));
  return m;
}

/** Órdenes y carritos del catálogo web por RUC (solo si tu cuenta puede abrir la base maestra). */
function leerWeb_() {
  const out = {};
  try {
    const bm = SpreadsheetApp.openById(BASE_MAESTRA);
    const p = bm.getSheetByName('Pedidos'); const n = p.getLastRow() - 1;
    if (n > 0) p.getRange(2, 1, n, 13).getValues().forEach(r => {
      const ruc = String(r[3]).replace(/\D/g, ''); if (!ruc || r[12] === 'Anulado') return;
      (out[ruc] = out[ruc] || []).push({ tipo: 'Orden web', num: r[0], fecha: aFecha_(r[1]), items: String(r[8]), total: Number(String(r[9]).replace(/[^\d.]/g, '')) || 0, estado: r[12] });
    });
    const c = bm.getSheetByName('Carritos'); const m = c ? c.getLastRow() - 1 : 0;
    if (m > 0) c.getRange(2, 1, m, 8).getValues().forEach(r => {
      const ruc = String(r[2]).replace(/\D/g, ''); if (!ruc || r[7] === 'Convertido') return;
      (out[ruc] = out[ruc] || []).push({ tipo: 'Carrito ' + String(r[7]).toLowerCase(), fecha: aFecha_(r[0]), items: String(r[5]), total: Number(r[6]) || 0 });
    });
  } catch (e) { /* sin acceso a la base del catálogo */ }
  return out;
}

/** Ciclo de recompra en días (promedio entre pedidos) o null si tiene menos de 2. */
function cicloRecompra_(peds) {
  const f = (peds || []).map(p => p.fecha).filter(Boolean).map(d => dia0_(d).getTime());
  const u = [...new Set(f)].sort((a, b) => a - b);
  if (u.length < 2) return null;
  return Math.round((u[u.length - 1] - u[0]) / 86400000 / (u.length - 1));
}

/** Intento N°: llamadas seguidas sin respuesta al final del historial + 1. */
function intento_(hist) {
  let k = 0;
  for (let i = (hist || []).length - 1; i >= 0; i--) { if (['No contesta', 'Ocupad@'].indexOf(hist[i].res) >= 0) k++; else break; }
  return k + 1;
}

/** Puntaje 0-100: probabilidad de compra (etapa, compras, respuesta, actividad web, crédito). */
function puntaje_(r, C, hist, peds, web) {
  let s = 20;
  const etapa = String(r[C.etapa - 1]);
  s += { 'Cliente activo': 30, 'Cotizado': 20, 'Validado': 10 }[etapa] || 0;
  const np = (peds || []).length; if (np) s += 15 + Math.min(10, (np - 1) * 5);
  if ((web || []).length) s += 15;
  const ult = (hist || [])[hist && hist.length - 1];
  if (ult) {
    if (['Pedido tomado', 'Interesado - cotizar', 'Interesado - pidió catálogo', 'Catálogo enviado', 'WhatsApp'].indexOf(ult.res) >= 0) s += 15;
    if (ult.res === 'Número equivocado') s -= 20;
  }
  s -= Math.min(15, (intento_(hist) - 1) * 5);
  const nota = String(r[C.notas - 1] || '').toLowerCase();
  if (/activo/.test(nota)) s += 5;
  if (/debe|deuda|juicio|protesto|no pag|sobregiro/.test(nota)) s -= 15;
  if (C.linea && String(r[C.linea - 1]).replace(/\D/g, '')) s += 5;
  return Math.max(0, Math.min(100, s));
}

// =====================================================================
//  Lista del día (lunes a jueves)
// =====================================================================
/**
 * Orden: 0) órdenes web, carritos y no contestaron / ocupad@ (reintento), 1) agendadas para hoy o vencidas,
 * 2) 1ª llamada del mes, 3) recompra (según su ciclo de compra), 4) 2ª llamada del mes (14+ días).
 * Dentro de cada grupo: la zona del día primero (Config «Zona lunes»… si la llenas), luego prioridad y puntaje.
 */
function armarDia() {
  if (diaSemana_() > 4) return; // viernes, sábado y domingo no se asignan llamadas
  const ss = hoja_(), cl = ss.getSheetByName('Clientes'), C = colsClientes_(cl), cfg = cfgMap_(ss);
  const meta = Number(cfg['Meta diaria de contactos']) || 30;
  const diasRecompra = Number(cfg['Días para recompra']) || 30;
  const zonaDia = String(cfg['Zona ' + DIAS_ES[diaSemana_()].toLowerCase()] || '').split(',').map(s => s.trim()).filter(Boolean);
  const n = cl.getLastRow() - 1;
  const d = cl.getRange(2, 1, n, cl.getLastColumn()).getValues();
  const llam = leerLlamadas_(ss), peds = leerPedidos_(ss), web = leerWeb_();
  const hoy = hoyCero_(), DIA = 86400000, mesIni = new Date(hoy.getFullYear(), hoy.getMonth(), 1);
  const activo = r => r[0] && ['Inactivo', 'No interesado'].indexOf(r[C.etapa - 1]) < 0;
  const salida = d.map(r => [r[C.asig - 1], r[C.motivo - 1]]);
  const tomados = new Set(); let usados = 0;
  // Pendientes: asignados antes de hoy y sin contacto desde entonces
  d.forEach((r, i) => {
    const asig = aFecha_(r[C.asig - 1]), h = llam[r[0]] || [], ult = h.length ? h[h.length - 1].f : null;
    if (activo(r) && asig && asig < hoy && (!ult || ult < asig)) { tomados.add(i); usados++; }
  });
  const cand = [];
  d.forEach((r, i) => {
    if (tomados.has(i) || !activo(r)) return;
    const h = llam[r[0]] || [], ult = h.length ? h[h.length - 1].f : null;
    if (ult && ult >= hoy) return; // ya contactado hoy
    const fAcc = aFecha_(r[C.fecha - 1]), acc = String(r[C.accion - 1] || '');
    const mes = h.filter(x => x.f >= mesIni).length, pri = Number(r[C.pri - 1]) || 2;
    const desde = ult ? (hoy - dia0_(ult)) / DIA : 9999;
    const zona = zonaDia.length && zonaDia.indexOf(String(r[C.zona - 1])) >= 0 ? 0 : 1;
    const sc = puntaje_(r, C, h, peds[r[0]], web[String(r[C.ruc - 1]).replace(/\D/g, '')]);
    const k = (g, mot) => cand.push([g, zona, pri, -sc, i, mot]);
    if (fAcc && fAcc <= hoy && REINTENTO.indexOf(acc) >= 0) return k(0, '0 ' + (acc.indexOf('ocupad') >= 0 ? 'Estaba ocupad@' : 'No contestó') + ' · REINTENTAR (intento ' + intento_(h) + ')');
    if (fAcc && fAcc <= hoy) return k(1, '1 Agendada: ' + (acc || 'seguimiento'));
    if (fAcc && fAcc > hoy) return;
    if (mes === 0) return k(2, '2 1ª llamada del mes');
    const ciclo = cicloRecompra_(peds[r[0]]);
    const uc = (peds[r[0]] || [])[0];
    const sinCompra = uc && uc.fecha ? (hoy - dia0_(uc.fecha)) / DIA : Number(r[C.dias - 1]) || 0;
    if (r[C.etapa - 1] === 'Cliente activo' && sinCompra >= (ciclo ? Math.max(7, ciclo - 3) : diasRecompra))
      return k(3, '3 Recompra' + (ciclo ? ' (compra cada ' + ciclo + ' días)' : ''));
    if (mes === 1 && desde >= 14) k(4, '4 2ª llamada del mes');
  });
  cand.sort((a, b) => a[0] - b[0] || a[1] - b[1] || a[2] - b[2] || a[3] - b[3] || a[4] - b[4]);
  cand.forEach(c => { if (c[0] <= 1 || usados < meta) { salida[c[4]] = [hoy, c[5]]; usados++; } });
  cl.getRange(2, C.asig, n, 1).setValues(salida.map(x => [x[0]]));
  cl.getRange(2, C.motivo, n, 1).setValues(salida.map(x => [x[1]]));
}

// =====================================================================
//  Correos: día de ventas (lun-jue) y revisión del viernes
// =====================================================================
function resumenDiario(forzar) {
  const dia = diaSemana_();
  if (dia === 5 && forzar !== true) { revisionSemanal(); return; }
  if (dia > 4 && forzar !== true) return;
  armarDia();
  SpreadsheetApp.flush();
  const p = panelDatos();
  const td = 'style="padding:6px;border-bottom:1px solid #ddd;vertical-align:top"';
  const filas = p.cola.map(c => '<tr><td ' + td + '>' + c.motivo + '</td><td ' + td + '><b>' + c.razon + '</b><br><span style="color:#666">' + c.zona + ' · puntaje ' + c.puntaje + '</span></td></tr>').join('');
  const url = PANEL_URL;
  const html = '<div style="font-family:Arial;font-size:14px"><h2 style="color:#002172">Tu día de ventas · ' + p.cola.length + ' contactos</h2>' +
    '<p>Venta del mes: <b>' + p.kpi.venta + '</b> · Meta: <b>' + p.kpi.meta + '</b> · Avance: <b>' + p.kpi.avance + '</b></p>' +
    (url ? '<p><a href="' + url + '" style="background:#e8264f;color:#fff;padding:10px 16px;border-radius:8px;text-decoration:none">Abrir mi Panel de ventas</a></p>' : '') +
    '<table style="border-collapse:collapse">' + filas + '</table></div>';
  MailApp.sendEmail({ to: correo_(), subject: 'Tu día de ventas · ' + p.cola.length + ' contactos', htmlBody: html });
}

/** Viernes: revisión de la semana + plan IA de la próxima semana (sin llamadas). */
function revisionSemanal() {
  const r = panelRevision();
  const li = arr => arr.length ? '<ul>' + arr.map(p => '<li>' + [p.num, p.razon, p.items, 'S/ ' + p.total.toFixed(2)].join(' · ') + ' · <b>' + p.estado + '</b></li>').join('') + '</ul>' : '<p>—</p>';
  const url = PANEL_URL;
  const html = '<div style="font-family:Arial;font-size:14px"><h2 style="color:#002172">Revisión de la semana</h2>' +
    '<p>Contactos de la semana: <b>' + r.semana.contactos + '</b> · Pedidos: <b>' + r.semana.pedidos + '</b> · Interesados: <b>' + r.semana.interesados + '</b> · Pendientes que pasan al lunes: <b>' + r.semana.pendientes + '</b></p>' +
    '<h3>Pedidos por cerrar (pago, despacho, factura)</h3>' + li(r.abiertos) +
    '<h3>Plan para la próxima semana (IA)</h3><div>' + (r.plan || 'Activa la clave de IA para recibir el plan.').replace(/\n/g, '<br>') + '</div>' +
    (url ? '<p><a href="' + url + '">Abrir mi Panel de ventas</a></p>' : '') + '</div>';
  MailApp.sendEmail({ to: correo_(), subject: 'Revisión de la semana · ' + r.abiertos.length + ' pedidos por cerrar', htmlBody: html });
}

// =====================================================================
//  PANEL DE VENTAS (aplicación web)
// =====================================================================
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Panel')
    .setTitle('Panel de ventas · Tecno Log IA')
    .setFaviconUrl(PANEL_URL + 'icon-32.png')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

/** Datos del día: modo, indicadores y la cola de clientes en el orden de «Hoy». */
function panelDatos() {
  const ss = hoja_(), cl = ss.getSheetByName('Clientes'), C = colsClientes_(cl), cfg = cfgMap_(ss);
  const dia = diaSemana_(), hoy = hoyCero_();
  const n = cl.getLastRow() - 1;
  const d = cl.getRange(2, 1, n, cl.getLastColumn()).getValues();
  const llam = leerLlamadas_(ss), peds = leerPedidos_(ss), web = leerWeb_();
  const hd = ss.getSheetByName('Hoy').getRange('A1:H2').getDisplayValues();
  let hechasHoy = 0; Object.values(llam).forEach(l => l.forEach(x => { if (x.f >= hoy) hechasHoy++; }));
  const g = (r, k) => (C[k] ? r[C[k] - 1] : '');
  const cola = [];
  d.forEach(r => {
    const asig = aFecha_(g(r, 'asig')); if (!r[0] || !asig || dia0_(asig) > hoy) return;
    const h = llam[r[0]] || [], ult = h.length ? h[h.length - 1].f : null;
    if (ult && ult >= asig) return; // ya contactado después de asignarlo
    const ruc = String(g(r, 'ruc')).replace(/\D/g, '');
    const p = peds[r[0]] || [], w = web[ruc] || [];
    const ciclo = cicloRecompra_(p);
    const mot = String(g(r, 'motivo') || '');
    cola.push({
      id: r[0], ruc, razon: g(r, 'razon'), contacto: g(r, 'contacto'), wa: String(g(r, 'wa')).replace(/\D/g, ''),
      cel2: String(g(r, 'cel2')).replace(/\D/g, ''), fijo: String(g(r, 'fijo')), email: g(r, 'email'),
      zona: g(r, 'zona'), ciudad: [g(r, 'distrito'), g(r, 'provincia'), g(r, 'depto')].filter(Boolean).join(' - '), direccion: g(r, 'dir'),
      pago: g(r, 'pago'), linea: [g(r, 'moneda'), g(r, 'linea')].filter(String).join(' '), lineas: g(r, 'lineas'),
      validacion: g(r, 'validacion'), notas: String(g(r, 'notas') || ''), etapa: g(r, 'etapa'),
      accion: g(r, 'accion'), fechaAccion: fmt_(aFecha_(g(r, 'fecha'))), prioridad: Number(g(r, 'pri')) || 2,
      motivo: mot.replace(/^\d\s*/, ''), rango: Number(mot.charAt(0)) || 0,
      pendienteDe: dia0_(asig) < hoy ? fmt_(asig, 'dd/MM') : '',
      intento: intento_(h), contactos: h.length,
      puntaje: puntaje_(r, C, h, p, w),
      historial: h.slice(-5).reverse().map(x => ({ f: fmt_(x.f, 'dd/MM HH:mm'), res: x.res, com: x.com || '', canal: x.canal })),
      pedidos: p.slice(0, 3).map(x => ({ f: fmt_(x.fecha), num: x.num, items: x.items.join(', '), total: Math.round(x.total * 100) / 100, estado: x.estado })),
      web: w.slice(-3).reverse().map(x => ({ tipo: x.tipo, f: fmt_(x.fecha), num: x.num || '', items: x.items, total: x.total, estado: x.estado || '' })),
      recompra: ciclo && p[0] && p[0].fecha ? { ciclo, proxima: fmt_(new Date(dia0_(p[0].fecha).getTime() + ciclo * 86400000)) } : null,
      _asig: dia0_(asig).getTime(),
    });
  });
  cola.sort((a, b) => a.rango - b.rango || a._asig - b._asig || a.prioridad - b.prioridad || b.puntaje - a.puntaje);
  return {
    hoy: DIAS_ES[dia] + ' ' + fmt_(new Date(), 'dd/MM/yyyy'), dia,
    modo: dia <= 4 ? 'llamadas' : dia === 5 ? 'revision' : 'descanso',
    ia: !!clave_(), catalogo: CATALOGO_URL, resultados: Object.keys(REGLAS), reglas: REGLAS,
    kpi: { meta: hd[1][3], venta: hd[1][1], avance: hd[1][5], pedidosPorAtender: hd[0][7], porValidar: hd[1][7],
      metaDia: Number(cfg['Meta diaria de contactos']) || 30, hechasHoy },
    mensajeVenta: String(cfg['Mensaje de venta'] || ''),
    cola,
  };
}

/** Guarda el contacto en «Llamadas», actualiza la ficha del cliente y devuelve OK. */
function panelRegistrar(x) {
  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    const ss = hoja_(), cl = ss.getSheetByName('Clientes'), C = colsClientes_(cl);
    if (!REGLAS[x.resultado]) throw new Error('Resultado no válido');
    const ids = cl.getRange(2, 1, cl.getLastRow() - 1, 1).getValues().flat();
    const i = ids.indexOf(x.id); if (i < 0) throw new Error('Cliente no encontrado');
    const r = cl.getRange(i + 2, 1, 1, cl.getLastColumn()).getValues()[0];
    const [, accSug, dias, canal] = REGLAS[x.resultado];
    const accion = String(x.accion || accSug || '').trim();
    const fecha = x.fecha ? aFecha_(x.fecha) : (accion ? diaHabil_(x.dias != null ? Number(x.dias) : dias) : '');
    const ll = ss.getSheetByName('Llamadas');
    const fila = ll.getLastRow() + 1;
    ll.getRange(fila, 1, 1, 9).setValues([[new Date(), x.id, r[C.razon - 1], r[C.zona - 1], x.canal || canal, x.resultado,
      String(x.comentario || '').slice(0, 500), accion, fecha || '']]);
    ll.getRange(fila, 1).setNumberFormat('dd/mm/yyyy hh:mm'); ll.getRange(fila, 9).setNumberFormat('dd/mm/yyyy');
    aplicarResultado_(ss, x.id, x.resultado, String(x.comentario || '').slice(0, 300), accion, fecha);
    if (x.resultado === 'Pedido tomado') marcarClienteActivo_(ss, x.id);
    return { ok: true, accion, fecha: fmt_(fecha || null) };
  } finally { lock.releaseLock(); }
}

/** Viernes: pedidos abiertos y resumen de la semana (también lo usa el correo). */
function panelRevision() {
  const ss = hoja_(), hoy = hoyCero_();
  const lunes = new Date(hoy); lunes.setDate(hoy.getDate() - ((diaSemana_() + 6) % 7));
  const sh = ss.getSheetByName('Pedidos'); const n = sh.getLastRow() - 1;
  const v = n > 0 ? sh.getRange(2, 1, n, 14).getValues() : [];
  const porNum = {};
  v.forEach((r, i) => {
    if (!r[2] && !r[3]) return;
    const k = r[1] || ('fila' + (i + 2));
    const p = porNum[k] = porNum[k] || { num: r[1] || '(sin N°)', fecha: fmt_(aFecha_(r[0])), id: r[2], razon: r[3], estado: r[10] || 'Nuevo', canal: r[11], filas: [], items: [], total: 0 };
    p.filas.push(i + 2); p.items.push((r[7] || 1) + ' × ' + (r[6] || r[5])); p.total += Number(r[9]) || 0;
  });
  const todos = Object.values(porNum);
  const abiertos = todos.filter(p => ['Entregado', 'Anulado'].indexOf(String(p.estado)) < 0).map(p => Object.assign(p, { items: p.items.join(', ') }));
  const llam = leerLlamadas_(ss);
  let contactos = 0, interesados = 0; const porRes = {}, porZona = {};
  const cl = ss.getSheetByName('Clientes'), C = colsClientes_(cl);
  const d = cl.getRange(2, 1, cl.getLastRow() - 1, cl.getLastColumn()).getValues();
  const zonaDe = {}; d.forEach(r => zonaDe[r[0]] = r[C.zona - 1]);
  Object.entries(llam).forEach(([id, l]) => l.forEach(x => {
    if (x.f < lunes) return; contactos++; porRes[x.res] = (porRes[x.res] || 0) + 1;
    const z = zonaDe[id] || 'Sin zona'; porZona[z] = (porZona[z] || 0) + 1;
    if (/Interesado|Catálogo|Pedido/.test(x.res)) interesados++;
  }));
  const pendientes = d.filter(r => { const a = aFecha_(r[C.asig - 1]); const l = llam[r[0]] || []; const u = l.length ? l[l.length - 1].f : null; return r[0] && a && a >= lunes && (!u || u < a); }).length;
  const semana = { contactos, interesados, pedidos: todos.filter(p => { const f = aFecha_(p.fecha); return f && f >= lunes; }).length, pendientes, porRes, porZona };
  let plan = '';
  if (clave_()) {
    try {
      plan = gemini_('Eres gerente comercial de Tecno Log IA (venta multimarca SOLO a distribuidores y mayoristas en el Perú; llamadas de lunes a jueves, viernes revisión). ' +
        'Con estos datos de la semana escribe un plan breve para la próxima semana en 5 viñetas: zonas y tipo de clientes a priorizar, qué reintentar, qué ofrecer y una meta de contactos por día. ' +
        'Sin inventar cifras que no estén en los datos. Español del Perú, directo.\nDATOS: ' + JSON.stringify({ semana, pedidosAbiertos: abiertos.length }), false);
    } catch (e) { plan = ''; }
  }
  return { abiertos: abiertos.map(p => ({ num: p.num, fecha: p.fecha, razon: p.razon, items: p.items, total: Math.round(p.total * 100) / 100, estado: p.estado, filas: p.filas })), semana, plan };
}

/** Cambia el estado de un pedido (todas sus filas). */
function panelEstadoPedido(filas, estado) {
  const ok = ['Nuevo', 'Confirmado', 'Facturado', 'Despachado', 'Entregado', 'Pagado', 'Anulado'];
  if (ok.indexOf(estado) < 0) throw new Error('Estado no válido');
  const sh = hoja_().getSheetByName('Pedidos');
  (filas || []).forEach(f => { if (f >= 2) sh.getRange(f, 11).setValue(estado); });
  return { ok: true };
}

/** Productos del catálogo (para la cotización en 1 clic). */
function panelProductos() {
  const r = UrlFetchApp.fetch(CATALOGO_API + '?accion=productos', { muteHttpExceptions: true });
  const d = JSON.parse(r.getContentText());
  return (d.productos || []).filter(p => p.estado !== 'Agotado').map(p => ({ id: p.id, c: p.codigo, n: p.nombre, m: p.marca, k: p.categoria, p: p.pvp, o: p.oferta || '' }));
}

// =====================================================================
//  IA (Google Gemini) — guion, mensaje de WhatsApp y notas libres
//  Privacidad: a la IA no se envían RUC, teléfonos, correos ni direcciones.
// =====================================================================
function clave_() { return PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY'); }

function probarIA() { const t = Date.now(); Logger.log(gemini_('Responde solo: «IA conectada a Tecno Log IA».', false) + ' · ' + Math.round((Date.now() - t) / 1000) + ' s'); }

function gemini_(prompt, comoJson) {
  const ss = hoja_();
  const principal = String(cfgMap_(ss)['Modelo IA'] || 'gemini-flash-latest');
  // Si el modelo principal está saturado (503/429), reintenta y pasa a uno más liviano
  const modelos = [principal, 'gemini-flash-lite-latest', 'gemini-flash-latest'].filter((m, i, a) => a.indexOf(m) === i);
  let ultimo = '';
  for (const modelo of modelos) {
    // La clave va en el encabezado (funciona con las claves nuevas «AQ.» y las antiguas «AIza»)
    const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(modelo) + ':generateContent';
    const llamar = conPensamiento => {
      const gc = { temperature: 0.4, maxOutputTokens: 4096 };
      if (comoJson) gc.responseMimeType = 'application/json';
      if (!conPensamiento) gc.thinkingConfig = { thinkingBudget: 0 }; // respuesta rápida
      return UrlFetchApp.fetch(url, { method: 'post', contentType: 'application/json', headers: { 'x-goog-api-key': clave_() },
        payload: JSON.stringify({ contents: [{ role: 'user', parts: [{ text: prompt }] }], generationConfig: gc }), muteHttpExceptions: true });
    };
    for (let intento = 0; intento < 2; intento++) {
      let r = llamar(false);
      if (r.getResponseCode() === 400) r = llamar(true);
      const code = r.getResponseCode();
      if (code === 200) {
        const c = (JSON.parse(r.getContentText()).candidates || [])[0];
        const t = c && c.content ? (c.content.parts || []).filter(p => !p.thought).map(p => p.text || '').join('') : '';
        if (!t) { ultimo = 'sin respuesta' + (c && c.finishReason ? ' (' + c.finishReason + ')' : ''); break; }
        if (!comoJson) return t.trim();
        return JSON.parse(t.slice(t.indexOf('{'), t.lastIndexOf('}') + 1));
      }
      ultimo = code + ': ' + r.getContentText().slice(0, 120);
      if (code !== 503 && code !== 429 && code !== 500) break; // otros errores: probar el siguiente modelo
      Utilities.sleep(1500 * (intento + 1));
    }
  }
  throw new Error('IA ocupada, intenta en un momento (' + ultimo + ')');
}

function ofertas_() {
  const cache = CacheService.getScriptCache(); const c = cache.get('ofertas');
  if (c) return JSON.parse(c);
  let o = [];
  try { o = panelProductos().filter(p => p.o).slice(0, 25).map(p => p.n + ' (' + (p.m || '') + ', S/ ' + p.p + ', oferta ' + p.o + ')'); } catch (e) { /* sin catálogo */ }
  cache.put('ofertas', JSON.stringify(o), 3600);
  return o;
}

/** Guion de 3 líneas + mensaje de WhatsApp personalizado para el cliente (se guarda en caché 6 h). */
function panelIA(c) {
  if (!clave_()) return { ok: false, error: 'Falta GEMINI_API_KEY' };
  const cache = CacheService.getUserCache(), k = 'ia_' + c.id + '_' + fmt_(new Date(), 'yyyyMMdd');
  const h = cache.get(k); if (h) return JSON.parse(h);
  const datos = {
    negocio: c.razon, contacto: c.contacto || '', zona: c.zona, etapa: c.etapa, motivo_de_hoy: c.motivo, intento: c.intento,
    condicion_pago: c.pago, lineas_de_interes: c.lineas, notas: String(c.notas || '').replace(/\d{6,}/g, '[número]').slice(0, 600),
    proxima_accion: c.accion, historial: c.historial, compras: c.pedidos, catalogo_web: c.web, recompra: c.recompra,
  };
  const prompt = 'Eres el mejor asesor de ventas B2B del Perú. Vendes para Tecno Log IA (multimarca: Oster y equipos de cocina, frío, panadería, etc.) SOLO a distribuidores y mayoristas. ' +
    'El vendedor se llama David Díaz. Prepara la próxima llamada con estos datos del cliente (no inventes compras, precios ni datos que no estén):\n' + JSON.stringify(datos) +
    '\nOFERTAS VIGENTES DEL CATÁLOGO: ' + JSON.stringify(ofertas_()) +
    '\nResponde en JSON: {"guion":["apertura (1 frase)","propuesta concreta según su historial y ofertas (1-2 frases)","cierre con pregunta para tomar pedido o siguiente paso (1 frase)"],' +
    '"objecion":"la objeción más probable y cómo responderla (1 frase)","whatsapp":"mensaje de WhatsApp de máx. 300 caracteres, cordial, que mencione el catálogo digital y termine con una pregunta; firma David Díaz · Tecno Log IA"}';
  const out = gemini_(prompt, true);
  const res = { ok: true, guion: (out.guion || []).slice(0, 3), objecion: out.objecion || '', whatsapp: String(out.whatsapp || '').slice(0, 400) };
  cache.put(k, JSON.stringify(res), 21600);
  return res;
}

/** Convierte lo que escribes (o dictas) en resultado, comentario limpio, próxima acción y días. */
function panelNotasIA(texto, c) {
  if (!clave_()) return { ok: false, error: 'Falta GEMINI_API_KEY' };
  const prompt = 'Clasifica el resultado de una llamada de ventas B2B. Notas del vendedor: «' + String(texto).slice(0, 800) + '». ' +
    'Cliente: ' + JSON.stringify({ negocio: c.razon, etapa: c.etapa, motivo: c.motivo }) + '. Hoy es ' + DIAS_ES[diaSemana_()] + ' ' + fmt_(new Date()) + '. Las llamadas son de lunes a jueves.\n' +
    'Resultados permitidos (elige UNO exacto): ' + JSON.stringify(Object.keys(REGLAS)) +
    '\nResponde en JSON: {"resultado":"...","comentario":"resumen claro en 1 frase con nombres, productos y cantidades mencionados","accion":"próxima acción concreta","fecha":"dd/mm/aaaa (lunes a jueves; si dijo un día, úsalo)","productos":["productos mencionados"]}';
  const o = gemini_(prompt, true);
  if (!REGLAS[o.resultado]) o.resultado = 'Volver a llamar';
  return { ok: true, resultado: o.resultado, comentario: String(o.comentario || texto).slice(0, 500), accion: String(o.accion || '').slice(0, 120), fecha: String(o.fecha || ''), productos: o.productos || [] };
}
