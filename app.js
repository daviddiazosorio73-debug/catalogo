/* Tecno Log IA · Catálogo digital
 * Datos: API de Google Apps Script (config.js → api) o, si no hay API, data/productos.json.
 * Vistas: inicio (ofertas, marcas, categorías), resultados, producto (uno a la vez),
 * presentación (play), carrito, cotización (PDF + WhatsApp) y orden de compra con firma.
 */
(() => {
  'use strict';
  const CFG = window.CATALOGO_CONFIG || {};
  const $ = (s, el = document) => el.querySelector(s);
  const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const money = n => 'S/ ' + Number(n || 0).toLocaleString('es-PE', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const norm = s => String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
  const normCod = s => String(s || '').toUpperCase().replace(/[\s\-_.\/]/g, '');
  const store = {
    get: (k, d) => { try { return JSON.parse(localStorage.getItem(k)) ?? d; } catch { return d; } },
    set: (k, v) => { try { localStorage.setItem(k, JSON.stringify(v)); } catch { /* sin almacenamiento */ } },
  };

  const S = {
    config: {}, productos: [], porId: {}, marcas: [], categorias: [],
    cart: store.get('tli_cart', {}),   // {id: cantidad}
    lista: [],                         // lista actual (para siguiente/anterior y presentación)
    filtro: { q: '', marca: '', cat: '' },
    oferta: '',
  };
  const params = new URLSearchParams(location.search);
  if (params.get('d')) sessionStorage.setItem('tli_origen', params.get('d'));
  const origen = () => sessionStorage.getItem('tli_origen') || 'directo';
  // Sesión anónima del equipo (para métricas y carritos; no guarda datos personales)
  let SES = store.get('tli_ses', '');
  if (!SES) { SES = Date.now().toString(36) + Math.random().toString(36).slice(2, 8); store.set('tli_ses', SES); }
  const cliente = () => store.get('tli_cliente', {});

  // ---------- Métricas (búsquedas, vistas, agregados) ----------
  let colaEv = [];
  function track(tipo, dato, n) {
    if (!CFG.api) return;
    colaEv.push({ t: Date.now(), tipo, dato: String(dato || '').slice(0, 120), n });
    if (colaEv.length >= 20) enviarEventos();
  }
  function enviarEventos() {
    if (!colaEv.length || !CFG.api) return;
    const body = JSON.stringify({ accion: 'eventos', sesion: SES, ruc: cliente().ruc || '', origen: origen(), eventos: colaEv.splice(0, 50) });
    try {
      if (!navigator.sendBeacon || !navigator.sendBeacon(CFG.api, new Blob([body], { type: 'text/plain;charset=UTF-8' })))
        fetch(CFG.api, { method: 'POST', mode: 'no-cors', keepalive: true, headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body });
    } catch { /* sin conexión */ }
  }
  setInterval(enviarEventos, 15000);
  document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'hidden') { enviarEventos(); sincronizarCarrito(true); } });

  // ---------- Carrito en la nube (solo si el cliente ya se identificó con su RUC) ----------
  let syncT = null, ultimoSync = '';
  function sincronizarCarrito(ya) {
    clearTimeout(syncT);
    const c = cliente();
    if (!CFG.api || !/^(10|15|17|20)\d{9}$/.test(c.ruc || '')) return;
    const items = Object.entries(S.cart).map(([id, cantidad]) => ({ id, cantidad }));
    const firma = JSON.stringify(items);
    if (!items.length || firma === ultimoSync) return;
    const enviar = () => {
      ultimoSync = firma;
      const body = JSON.stringify({ accion: 'carrito', sesion: SES, ruc: c.ruc, razon: c.razon || '', celular: c.celular || '', origen: origen(), items });
      try { fetch(CFG.api, { method: 'POST', mode: 'no-cors', keepalive: true, headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body }); } catch { /* sin conexión */ }
    };
    if (ya) enviar(); else syncT = setTimeout(enviar, 4000);
  }

  // ---------- Carga ----------
  // Muestra al instante lo guardado en el equipo y actualiza en segundo plano desde la hoja.
  async function cargar() {
    // Índice de fotos copiadas al sitio (no bloquea más de 1,5 s)
    try {
      const r = await Promise.race([fetch('assets/d/index.json', { cache: 'no-cache' }), new Promise((_, ko) => setTimeout(ko, 1500))]);
      if (r && r.ok) S.espejo = new Set(await r.json());
    } catch { /* sin espejo: se usa Drive */ }
    const guardado = store.get('tli_cache', null);
    let mostrado = false;
    if (guardado && guardado.productos && guardado.productos.length) { aplicar(guardado); mostrado = true; }
    let data = null;
    if (CFG.api) {
      try {
        const r = await fetch(CFG.api + '?accion=productos', { cache: 'no-store' });
        data = await r.json();
        if (!data.ok) data = null; else store.set('tli_cache', data);
      } catch { data = null; }
    }
    if (!data && !mostrado) {
      try { data = await (await fetch('data/productos.json', { cache: 'no-cache' })).json(); } catch { /* sin copia local */ }
    }
    if (data) aplicar(data, mostrado);
    else if (!mostrado) $('#main').innerHTML = '<div class="vacio">No se pudo cargar el catálogo. Revisa tu conexión.</div>';
  }

  function aplicar(data, refresco) {
    S.config = data.config || {};
    S.productos = (data.productos || []).filter(p => p.estado !== 'Oculto');
    S.porId = {};
    S.productos.forEach(p => {
      p._t = norm([p.codigo, p.nombre, p.marca, p.categoria, p.familia, p.lista].join(' '));
      p._c = normCod(p.codigo);
      S.porId[p.id] = p;
    });
    const cuenta = (key) => {
      const m = {};
      S.productos.forEach(p => { const k = p[key] || 'Otras'; m[k] = (m[k] || 0) + 1; });
      return Object.entries(m).sort((a, b) => b[1] - a[1]);
    };
    S.marcas = cuenta('marca');
    S.categorias = cuenta('categoria');
    if (S.config.vigencia) $('#vig').textContent = '· PVP con IGV · ' + S.config.vigencia;
    $('#wa').href = 'https://wa.me/' + (S.config.whatsapp || '51928146619') + '?text=' + encodeURIComponent('Hola Tecno Log IA, tengo una consulta del catálogo digital.');
    mostrarAsesor();
    Object.keys(S.cart).forEach(id => { if (!S.porId[id]) delete S.cart[id]; });
    // Cotización enviada por el asesor: ?pedido=ID*2,ID*1&ruc=20xxxxxxxxx (se aplica una sola vez)
    const ped = params.get('pedido');
    if (ped && !aplicar.cot) {
      aplicar.cot = true;
      const nuevo = {};
      ped.split(',').forEach(x => { const [id, q] = x.split('*'); const idd = decodeURIComponent(id || ''); if (S.porId[idd] && S.porId[idd].estado !== 'Agotado') nuevo[idd] = Math.max(1, Math.min(999, parseInt(q, 10) || 1)); });
      if (Object.keys(nuevo).length) { S.cart = nuevo; guardarCart(); setTimeout(() => { abrirCart(); toast('Tu cotización está lista: revisa y genera tu orden'); }, 300); }
      const rucP = (params.get('ruc') || '').replace(/\D/g, '');
      if (/^(10|15|17|20)\d{9}$/.test(rucP) && !cliente().ruc) store.set('tli_cliente', Object.assign(cliente(), { ruc: rucP }));
      history.replaceState(null, '', location.pathname + location.hash);
    }
    guardarCart();
    // En el refresco no se interrumpe al usuario si está viendo un producto, el carrito o la orden
    if (refresco && (!$('#pv').hidden || !$('#co').hidden || !$('#cart').hidden || !$('#play').hidden)) return;
    router();
  }

  // ---------- Utilidades de producto ----------
  // Fotos: copia propia en el sitio (assets/d, rápida y sin límites de Google Drive) y, si falla, Drive en 2 formatos
  const idDrive = u => { const m = String(u).match(/\/d\/([\w-]{20,})|[?&]id=([\w-]{20,})/); return m ? (m[1] || m[2]) : ''; };
  const variantes = u => {
    const id = idDrive(u); if (!id) return [u];
    const v = [];
    if (S.espejo && S.espejo.has(id)) v.push('assets/d/' + id + '.webp');
    v.push('https://lh3.googleusercontent.com/d/' + id + '=w1000', 'https://drive.google.com/thumbnail?id=' + id + '&sz=w1000');
    return v;
  };
  const fotosDe = p => { const vis = new Set(); return [p.foto, ...(p.fotos || [])].filter(Boolean).filter(f => { const k = idDrive(f) || f; if (vis.has(k)) return false; vis.add(k); return true; }).map(f => variantes(f)[0]); };
  // Cadena de respaldo: prueba la siguiente versión de la misma foto, luego la siguiente foto del producto
  const altsDe = p => [p.foto, ...(p.fotos || [])].filter(Boolean).flatMap(variantes).filter((x, i, a) => a.indexOf(x) === i);
  window.tliErr = function (img) {
    const alts = (img.dataset.alts || '').split('|').filter(Boolean);
    const k = Number(img.dataset.k || 0) + 1;
    if (k < alts.length) { img.dataset.k = k; img.src = alts[k]; return; }
    img.onerror = null;
    if (img.dataset.ph === 'no') { img.style.visibility = 'hidden'; return; }
    img.replaceWith(Object.assign(document.createElement('div'), { className: 'noimg', innerHTML: '<div><img src="assets/icon-192.png" alt="">Foto pronto</div>' }));
  };
  const imgTag = (src, alts, extra = '') => `<img src="${esc(src)}" data-alts="${esc(alts.join('|'))}" data-k="${Math.max(0, alts.indexOf(src))}" onerror="tliErr(this)" ${extra}>`;
  const imgHTML = (p, cls = '') => {
    const f = fotosDe(p)[0];
    return f
      ? imgTag(f, altsDe(p), `class="${cls}" alt="${esc(p.nombre)}" loading="lazy"`)
      : `<div class="noimg"><div><img src="assets/icon-192.png" alt="">Foto pronto</div></div>`;
  };
  const badges = p => {
    const b = [];
    if (p.oferta) b.push(`<span class="badge of">Oferta ${esc(p.oferta.toLowerCase() === 'día' ? 'del día' : 'de la ' + p.oferta.toLowerCase()).replace('de la mes', 'del mes')}</span>`);
    if (p.nuevo) b.push('<span class="badge nv">Nuevo</span>');
    if (p.estado === 'Próximo ingreso') b.push('<span class="badge pi">Próximo ingreso</span>');
    if (p.estado === 'Agotado') b.push('<span class="badge ag">Agotado</span>');
    return `<div class="badges">${b.join('')}</div>`;
  };
  const tarjeta = p => `
    <article class="card" data-id="${esc(p.id)}">
      <div class="ph">${imgHTML(p)}${badges(p)}${p.estado !== 'Agotado' ? `<button class="add" data-add="${esc(p.id)}" aria-label="Agregar al carrito">+</button>` : ''}</div>
      <div class="info"><span class="cod">${esc(p.marca ? p.marca + ' · ' : '')}${esc(p.codigo)}</span>
      <span class="nom">${esc(p.nombre)}</span>
      <span class="pre">${money(p.pvp)} <small>inc. IGV</small></span></div>
    </article>`;

  // ---------- Búsqueda ----------
  function buscar() {
    const { q, marca, cat } = S.filtro;
    const nq = norm(q).trim(), cq = normCod(q);
    const palabras = nq.split(/\s+/).filter(Boolean);
    let r = S.productos.filter(p =>
      (!marca || (p.marca || 'Otras') === marca) &&
      (!cat || (p.categoria || 'Otras') === cat) &&
      (!palabras.length || palabras.every(w => p._t.includes(w)) || (cq.length >= 3 && p._c.includes(cq))));
    if (cq.length >= 3) r.sort((a, b) => (b._c === cq) - (a._c === cq) || (b._c.startsWith(cq)) - (a._c.startsWith(cq)));
    // disponibles primero
    const rank = p => (p.estado === 'Activo' ? 0 : p.estado === 'Próximo ingreso' ? 1 : 2);
    r = r.map((p, i) => [p, i]).sort((a, b) => rank(a[0]) - rank(b[0]) || a[1] - b[1]).map(x => x[0]);
    return r;
  }

  // ---------- Vistas ----------
  function router() {
    const id = new URLSearchParams(location.search).get('p');
    const h = location.hash.slice(1);
    const hp = new URLSearchParams(h);
    S.filtro = { q: hp.get('q') || '', marca: hp.get('m') || '', cat: hp.get('c') || '' };
    $('#q').value = S.filtro.q;
    if (S.filtro.q || S.filtro.marca || S.filtro.cat) vistaResultados(); else vistaInicio();
    if (id && S.porId[id]) abrirProducto(id);
  }

  function irA(f) {
    const p = new URLSearchParams();
    if (f.q) p.set('q', f.q); if (f.m) p.set('m', f.m); if (f.c) p.set('c', f.c);
    location.hash = p.toString();
  }

  function vistaInicio() {
    const ofertas = { 'Día': [], 'Semana': [], 'Mes': [] };
    const zona = cliente().zona || '';
    // Ofertas de toda la red + las de la zona del cliente (si ya se identificó con su RUC)
    const paraMi = p => !(p.zonas || []).length || (zona && p.zonas.indexOf(zona) >= 0);
    S.productos.forEach(p => { if (ofertas[p.oferta] && paraMi(p)) ofertas[p.oferta].push(p); });
    const ult = store.get('tli_ultimo', null);
    const ultDisp = ult ? ult.items.filter(x => S.porId[x.id] && S.porId[x.id].estado !== 'Agotado') : [];
    Object.values(ofertas).forEach(l => l.sort((a, b) => (a.orden || 999) - (b.orden || 999)));
    const tabs = Object.keys(ofertas).filter(k => ofertas[k].length);
    if (!S.oferta || !ofertas[S.oferta] || !ofertas[S.oferta].length) S.oferta = tabs[0] || '';
    const lista = S.oferta ? ofertas[S.oferta] : [];
    const nuevos = S.productos.filter(p => p.nuevo).slice(0, 24);
    $('#main').innerHTML = `
      ${ultDisp.length ? `<section class="repetir"><div><b>Tu último pedido</b> · N° ${esc(ult.numero)} · ${esc(ult.fecha)}<br>
        <span class="meta">${ultDisp.length} producto${ultDisp.length === 1 ? '' : 's'} · ${money(ultDisp.reduce((t, x) => t + S.porId[x.id].pvp * x.cantidad, 0))} a precio de hoy</span></div>
        <button class="btn red" data-repetir>↻ Volver a pedir</button></section>` : ''}
      ${tabs.length ? `<section class="hero">
        <h2>🔥 Ofertas</h2>
        <div class="tabs">${tabs.map(t => `<button class="pill ${t === S.oferta ? 'on' : ''}" data-of="${t}">${t === 'Día' ? 'Del día' : t === 'Semana' ? 'De la semana' : 'Del mes'} <b>${ofertas[t].length}</b></button>`).join('')}
          <button class="pill" data-play="ofertas">▶ Ver en presentación</button></div>
        <div class="carrusel">${lista.map(tarjeta).join('')}</div></section>` : ''}
      <h2>Marcas</h2>
      <div class="row">${S.marcas.map(([m, n]) => `<button class="pill" data-m="${esc(m)}">${esc(m)} <b>${n}</b></button>`).join('')}</div>
      <h2>Categorías <small>${S.categorias.length}</small></h2>
      <div class="row" style="flex-wrap:wrap">${S.categorias.map(([c, n]) => `<button class="pill" data-c="${esc(c)}">${esc(c)} <b>${n}</b></button>`).join('')}</div>
      ${nuevos.length ? `<h2>Novedades</h2><div class="grid">${nuevos.map(tarjeta).join('')}</div>` : ''}
      <footer class="pie">${esc(S.config.nombre || 'Tecno Log IA')} · ${S.productos.length} productos · Solo distribuidores y mayoristas · PVP con IGV</footer>`;
    S.lista = lista.length ? lista : S.productos;
  }

  let paginaActual = 1;
  let ultimaBusq = '';
  function vistaResultados() {
    const r = buscar();
    S.lista = r;
    const qq = norm(S.filtro.q).trim();
    if (qq.length >= 2 && qq !== ultimaBusq) { ultimaBusq = qq; track(r.length ? 'busqueda' : 'sin_resultado', qq, r.length); }
    paginaActual = 1;
    const { q, marca, cat } = S.filtro;
    const chips = [];
    if (marca) chips.push(`<button class="pill on" data-quitar="m">${esc(marca)} ✕</button>`);
    if (cat) chips.push(`<button class="pill on" data-quitar="c">${esc(cat)} ✕</button>`);
    if (q) chips.push(`<button class="pill on" data-quitar="q">“${esc(q)}” ✕</button>`);
    $('#main').innerHTML = `
      <div class="row">${chips.join('')}<button class="pill" data-inicio>← Inicio</button>
        ${r.length ? '<button class="pill" data-play="lista">▶ Presentación</button>' : ''}</div>
      <h2>${r.length} producto${r.length === 1 ? '' : 's'}</h2>
      ${r.length ? `<div class="grid" id="grid"></div><div style="text-align:center;padding:20px"><button class="btn ghost" id="mas" hidden>Ver más</button></div>`
        : `<div class="vacio">No encontramos productos con esa búsqueda.<br><br><button class="btn" data-inicio>Ver todo el catálogo</button></div>`}`;
    pintarPagina();
  }
  function pintarPagina() {
    const g = $('#grid'); if (!g) return;
    const N = 48, hasta = paginaActual * N;
    g.innerHTML = S.lista.slice(0, hasta).map(tarjeta).join('');
    $('#mas').hidden = S.lista.length <= hasta;
  }

  // ---------- Producto (uno a la vez) ----------
  let pvEstado = { id: null, i: 0, zoom: false };
  function abrirProducto(id, idx) {
    const p = S.porId[id]; if (!p) return;
    if (pvEstado.id !== id) track('vista', id);
    pvEstado = { id, i: 0, zoom: false };
    const media = fotosDe(p).map(f => ({ t: 'img', src: f }));
    const vid = driveId(p.video);
    if (vid) media.push({ t: 'vid', src: 'https://drive.google.com/file/d/' + vid + '/preview' });
    const variantes = S.productos.filter(x => x.codigo === p.codigo && x.id !== p.id);
    const pos = S.lista.findIndex(x => x.id === id);
    const enCarrito = S.cart[id] || 0;
    const url = location.origin + location.pathname + '?p=' + encodeURIComponent(id);
    $('#pv').innerHTML = `<div class="pvbox">
      <button class="pvclose" data-cerrar aria-label="Cerrar">✕</button>
      <div class="gal">
        <div class="stage" id="stage"></div>
        ${media.length > 1 ? `<div class="thumbs">${media.map((m, i) => m.t === 'img'
          ? `<button data-th="${i}" class="${i ? '' : 'on'}">${imgTag(m.src, altsDe(p), 'alt="" loading="lazy" data-ph="no"')}</button>`
          : `<button data-th="${i}" class="vid">▶</button>`).join('')}</div>` : ''}
      </div>
      <div class="det">
        <span class="marca">${esc(p.marca || 'Tecno Log IA')}</span>
        <h1>${esc(p.nombre)}</h1>
        <div class="meta">Código <b>${esc(p.codigo)}</b> · ${esc(p.categoria)}${p.familia ? ' · ' + esc(p.familia) : ''}</div>
        <div>${badges(p).replace('class="badges"', 'class="row"')}</div>
        <div class="precio">${money(p.pvp)}<small>Precio de venta al público (PVP) · incluye IGV</small></div>
        ${p.descripcion ? `<p class="desc">${esc(p.descripcion)}</p>` : ''}
        ${p.ideal ? `<p class="ideal"><b>Ideal para:</b> ${esc(p.ideal)}</p>` : ''}
        ${variantes.length ? `<div><b>Otras presentaciones</b><div class="var">${variantes.map(v => `<button class="pill" data-ver="${esc(v.id)}">${esc(v.nombre.replace(p.nombre, '').trim() || v.id)}</button>`).join('')}</div></div>` : ''}
        ${p.estado === 'Agotado' ? `<p class="err"><b>Agotado por ahora.</b> Escríbenos y te avisamos cuando llegue.</p>
          <div class="acc"><a class="btn green" target="_blank" rel="noopener" href="https://wa.me/${esc(S.config.whatsapp || '51928146619')}?text=${encodeURIComponent('Hola, quiero que me avisen cuando llegue: ' + p.codigo + ' · ' + p.nombre)}">Avísame por WhatsApp</a></div>`
        : `<div class="acc"><div class="qty"><button data-q="-1">−</button><input id="pvq" type="number" min="1" value="${enCarrito || 1}"><button data-q="1">+</button></div>
          <button class="btn red" data-addq="${esc(p.id)}">${enCarrito ? 'Actualizar carrito' : 'Agregar al carrito'}</button></div>
          ${p.estado === 'Próximo ingreso' ? '<p class="meta">Producto en camino: puedes pedirlo y se despacha a su llegada.</p>' : ''}`}
        ${(p.complementos || []).length ? `<div><b>Se complementa con</b><div class="var">${p.complementos.filter(c => S.categorias.some(x => x[0] === c)).map(c => `<button class="pill" data-c="${esc(c)}">${esc(c)}</button>`).join('')}</div></div>` : ''}
        <div class="docs">
          ${p.ficha ? `<a class="btn ghost" target="_blank" rel="noopener" href="${esc(p.ficha)}">📄 Ficha técnica</a>` : ''}
          ${p.manual ? `<a class="btn ghost" target="_blank" rel="noopener" href="${esc(p.manual)}">📘 Manual</a>` : ''}
          <a class="btn ghost" target="_blank" rel="noopener" href="https://wa.me/?text=${encodeURIComponent(p.nombre + ' · ' + url)}">Compartir</a>
        </div>
        <div class="pvbar">
          <button class="btn ghost" data-nav="-1" ${pos > 0 ? '' : 'disabled'}>← Anterior</button>
          <button class="btn ghost" data-nav="1" ${pos >= 0 && pos < S.lista.length - 1 ? '' : 'disabled'}>Siguiente →</button>
        </div>
      </div></div>`;
    $('#pv').hidden = false;
    document.body.style.overflow = 'hidden';
    pvEstado.media = media;
    pintarMedia();
    history.replaceState(null, '', '?p=' + encodeURIComponent(id) + location.hash);
  }
  function pintarMedia() {
    const m = pvEstado.media, st = $('#stage');
    if (!st) return;
    if (!m.length) { st.innerHTML = imgHTML(S.porId[pvEstado.id]); return; }
    const it = m[pvEstado.i];
    st.className = 'stage' + (pvEstado.zoom ? ' zoom' : '');
    st.innerHTML = (it.t === 'img' ? imgTag(it.src, altsDe(S.porId[pvEstado.id]), 'alt="" data-zoom') : `<iframe src="${esc(it.src)}" allow="autoplay; fullscreen" allowfullscreen></iframe>`)
      + (m.length > 1 ? '<button class="nav l" data-m="-1">‹</button><button class="nav r" data-m="1">›</button>' : '');
    document.querySelectorAll('.thumbs button').forEach((b, i) => b.classList.toggle('on', i === pvEstado.i));
  }
  function cerrarProducto() {
    $('#pv').hidden = true; $('#pv').innerHTML = '';
    document.body.style.overflow = '';
    history.replaceState(null, '', location.pathname + location.hash);
  }
  function driveId(u) {
    const m = String(u || '').match(/\/d\/([\w-]{20,})|[?&]id=([\w-]{20,})/);
    return m ? (m[1] || m[2]) : '';
  }

  // ---------- Presentación ----------
  let playT = null, playI = 0, playL = [], pausado = false;
  function abrirPlay(lista) {
    playL = lista.filter(p => p.estado !== 'Agotado');
    if (!playL.length) return;
    playI = 0; pausado = false;
    $('#play').hidden = false; document.body.style.overflow = 'hidden';
    pintarPlay();
  }
  function pintarPlay() {
    clearTimeout(playT);
    const p = playL[playI];
    $('#play').innerHTML = `
      <div class="bar"><i id="pbar"></i></div>
      <div class="ptop"><img src="assets/logo-wordmark.png" alt="Tecno Log IA"><span>${playI + 1} / ${playL.length}</span></div>
      <div class="pstage">${fotosDe(p)[0] ? imgTag(fotosDe(p)[0], altsDe(p), 'alt=""') : `<div class="noimg" style="width:min(60vh,80vw);height:min(60vh,80vw);border-radius:18px"><div><img src="assets/icon-192.png" alt="">${esc(p.codigo)}</div></div>`}</div>
      <div class="pinfo"><div>${esc(p.marca || '')} · ${esc(p.codigo)}</div><h3>${esc(p.nombre)}</h3><div class="pp">${money(p.pvp)}</div></div>
      <div class="pctl">
        <button class="btn ghost" data-pl="-1">‹ Anterior</button>
        <button class="btn" data-pl="pausa">${pausado ? '▶ Seguir' : '❚❚ Pausa'}</button>
        <button class="btn ghost" data-pl="1">Siguiente ›</button>
        <button class="btn red" data-add="${esc(p.id)}">+ Agregar</button>
        <button class="btn gold" data-pl="ver">Ver detalle</button>
        <button class="btn ghost" data-pl="cerrar">✕ Salir</button>
      </div>`;
    const bar = $('#pbar');
    if (!pausado) {
      bar.animate([{ width: '0%' }, { width: '100%' }], { duration: 5000, fill: 'forwards' });
      playT = setTimeout(() => { playI = (playI + 1) % playL.length; pintarPlay(); }, 5000);
    }
  }
  function cerrarPlay() { clearTimeout(playT); $('#play').hidden = true; $('#play').innerHTML = ''; document.body.style.overflow = ''; }

  // ---------- Carrito ----------
  function guardarCart() {
    store.set('tli_cart', S.cart);
    sincronizarCarrito();
    const n = Object.values(S.cart).reduce((a, b) => a + b, 0);
    $('#cartCount').textContent = n;
  }
  function agregar(id, cant, reemplazar) {
    const p = S.porId[id]; if (!p || p.estado === 'Agotado') return;
    cant = Math.max(1, Math.min(999, parseInt(cant, 10) || 1));
    S.cart[id] = reemplazar ? cant : (S.cart[id] || 0) + cant;
    guardarCart();
    track('agregado', id, S.cart[id]);
    toast('✓ ' + p.codigo + ' en el carrito (' + S.cart[id] + ')');
  }
  function totalCart() {
    return Object.entries(S.cart).reduce((t, [id, q]) => t + (S.porId[id] ? S.porId[id].pvp * q : 0), 0);
  }
  function abrirCart() {
    const ids = Object.keys(S.cart);
    $('#cart').innerHTML = `
      <header><b>Tu pedido (${ids.length})</b><button data-ccerrar aria-label="Cerrar">✕</button></header>
      <div class="items">${ids.length ? ids.map(id => {
        const p = S.porId[id];
        return `<div class="it">${imgHTML(p)}<div><div class="cod">${esc(p.codigo)}</div><div class="nom" style="font-size:13px">${esc(p.nombre)}</div>
          <div><b>${money(p.pvp * S.cart[id])}</b> <button class="x" data-del="${esc(id)}">Quitar</button></div></div>
          <div class="qty"><button data-cq="${esc(id)}" data-d="-1">−</button><input value="${S.cart[id]}" data-ci="${esc(id)}" type="number" min="1"><button data-cq="${esc(id)}" data-d="1">+</button></div></div>`;
      }).join('') : '<div class="vacio">Tu carrito está vacío.<br>Toca <b>+</b> en cualquier producto para agregarlo.</div>'}</div>
      <div class="sum"><div class="t"><span>Total</span><span>${money(totalCart())}</span></div>
        <small class="meta">PVP · incluye IGV · sujeto a stock</small>
        ${completa(ids)}
        <button class="btn red" data-orden ${ids.length ? '' : 'disabled'}>Generar orden de compra</button>
        <button class="btn gold" data-cotizar ${ids.length ? '' : 'disabled'}>🧾 Solicitar cotización</button>
        <button class="btn ghost" data-ccerrar>Seguir comprando</button></div>`;
    $('#cart').hidden = false;
  }
  function cerrarCart() { $('#cart').hidden = true; }
  // «Completa tu pedido»: productos de las categorías que complementan lo que ya lleva
  function completa(ids) {
    if (!ids.length) return '';
    const enCart = new Set(ids.map(id => S.porId[id].categoria));
    const cats = [];
    ids.forEach(id => (S.porId[id].complementos || []).forEach(c => { if (!enCart.has(c) && cats.indexOf(c) < 0) cats.push(c); }));
    if (!cats.length) return '';
    const puntaje = p => (p.oferta ? 0 : 2) + (fotosDe(p).length ? 0 : 1);
    const sug = [];
    cats.slice(0, 6).forEach(c => {
      const op = S.productos.filter(p => p.categoria === c && p.estado === 'Activo' && !S.cart[p.id]).sort((a, b) => puntaje(a) - puntaje(b));
      if (op[0]) sug.push(op[0]);
    });
    if (!sug.length) return '';
    return `<div class="completa"><b>Completa tu pedido</b><div class="cmini">${sug.map(p => `
      <div class="mini" data-ver="${esc(p.id)}">${imgHTML(p)}<span>${esc(p.categoria)}</span><b>${money(p.pvp)}</b>
      <button class="pill" data-addc="${esc(p.id)}">+ Agregar</button></div>`).join('')}</div></div>`;
  }

  // ---------- Orden de compra ----------
  let firma = { ctx: null, dibujo: false, vacia: true };
  function abrirOrden() {
    cerrarCart();
    const c = store.get('tli_cliente', {});
    const ids = Object.keys(S.cart);
    $('#co').innerHTML = `<div class="cobox">
      <button class="pvclose" data-cocerrar aria-label="Cerrar">✕</button>
      <h1 style="margin:0 0 4px">Orden de compra</h1>
      <p class="meta" style="margin:0 0 14px">Solo distribuidores y mayoristas · Recibirás la orden en tu correo.</p>
      <div class="resumen">${ids.map(id => `<div><span>${S.cart[id]} × ${esc(S.porId[id].codigo)} · ${esc(S.porId[id].nombre)}</span><b>${money(S.porId[id].pvp * S.cart[id])}</b></div>`).join('')}
        <div><span><b>TOTAL (inc. IGV)</b></span><b>${money(totalCart())}</b></div></div>
      <h2>Tus datos</h2>
      <form class="form" id="fo" novalidate>
        <label>RUC *<input name="ruc" inputmode="numeric" maxlength="11" required value="${esc(c.ruc || '')}" placeholder="20xxxxxxxxx"><span id="rucmsg" class="meta"></span></label>
        <label>Razón social *<input name="razon" required value="${esc(c.razon || '')}"></label>
        <label>Persona de contacto *<input name="contacto" required value="${esc(c.contacto || '')}"></label>
        <label>Celular / WhatsApp *<input name="celular" inputmode="tel" required value="${esc(c.celular || '')}"></label>
        <label class="full">Correo (aquí llegará tu orden) *<input name="email" type="email" required value="${esc(c.email || '')}"></label>
        <label class="full">Dirección de entrega *<input name="direccion" required value="${esc(c.direccion || '')}"></label>
        <label>Ciudad / distrito *<input name="ciudad" required value="${esc(c.ciudad || '')}"></label>
        <label>Notas (opcional)<input name="notas" placeholder="Horario, referencia, forma de pago…"></label>
        <label class="full">Firma de conformidad * <canvas class="firma" id="firma"></canvas>
          <span><button type="button" class="pill" id="limpiar">Borrar firma</button></span></label>
        <label class="full" style="flex-direction:row;gap:10px;align-items:flex-start;color:var(--ink);font-weight:500">
          <input type="checkbox" name="acepto" style="width:22px;height:22px" required>
          Confirmo mi pedido. Los precios incluyen IGV y están sujetos a stock y a la confirmación de Tecno Log IA.</label>
        <div class="full"><button class="btn red" style="width:100%;font-size:18px" id="enviar">Enviar orden de compra</button>
        <p id="comsg" class="meta"></p></div>
      </form></div>`;
    $('#co').hidden = false; document.body.style.overflow = 'hidden';
    prepararFirma();
    const ruc = $('#fo [name=ruc]');
    ruc.addEventListener('input', () => { ruc.value = ruc.value.replace(/\D/g, ''); if (ruc.value.length === 11) consultarRuc(ruc.value); });
    $('#limpiar').onclick = () => { limpiarFirma(); };
    $('#fo').onsubmit = enviarOrden;
  }
  async function consultarRuc(ruc) {
    const msg = $('#rucmsg');
    if (!/^(10|15|17|20)\d{9}$/.test(ruc)) { msg.textContent = 'Revisa el RUC (11 dígitos).'; msg.className = 'err'; return; }
    if (!CFG.api) { msg.textContent = ''; return; }
    msg.textContent = 'Buscando…'; msg.className = 'meta';
    try {
      const r = await (await fetch(CFG.api + '?accion=ruc&ruc=' + ruc)).json();
      const f = $('#fo');
      if (r.ok && r.encontrado) {
        if (!f.razon.value) f.razon.value = r.razon || '';
        if (!f.direccion.value) f.direccion.value = r.direccion || '';
        if (!f.ciudad.value) f.ciudad.value = r.ciudad || '';
        msg.textContent = '✓ Cliente registrado'; msg.className = 'ok';
        store.set('tli_cliente', Object.assign(cliente(), { ruc, razon: f.razon.value, zona: r.zona || '' }));
        sincronizarCarrito(true);
      } else { msg.textContent = 'Cliente nuevo: completa tus datos.'; msg.className = 'meta'; }
    } catch { msg.textContent = ''; }
  }
  function prepararFirma() {
    const c = $('#firma');
    const r = c.getBoundingClientRect(), dpr = window.devicePixelRatio || 1;
    c.width = r.width * dpr; c.height = r.height * dpr;
    const ctx = c.getContext('2d'); ctx.scale(dpr, dpr);
    ctx.lineWidth = 2.4; ctx.lineCap = 'round'; ctx.lineJoin = 'round'; ctx.strokeStyle = '#151823';
    firma = { ctx, dibujo: false, vacia: true };
    const pos = e => { const b = c.getBoundingClientRect(); return [e.clientX - b.left, e.clientY - b.top]; };
    c.onpointerdown = e => { try { c.setPointerCapture(e.pointerId); } catch { /* sin captura */ } firma.dibujo = true; const [x, y] = pos(e); ctx.beginPath(); ctx.moveTo(x, y); };
    c.onpointermove = e => { if (!firma.dibujo) return; const [x, y] = pos(e); ctx.lineTo(x, y); ctx.stroke(); firma.vacia = false; };
    c.onpointerup = c.onpointercancel = () => { firma.dibujo = false; };
  }
  function limpiarFirma() { const c = $('#firma'); firma.ctx.clearRect(0, 0, c.width, c.height); firma.vacia = true; }
  async function enviarOrden(ev) {
    ev.preventDefault();
    const f = $('#fo'), msg = $('#comsg');
    const d = Object.fromEntries(new FormData(f).entries());
    const faltan = ['ruc', 'razon', 'contacto', 'celular', 'email', 'direccion', 'ciudad'].filter(k => !String(d[k] || '').trim());
    if (faltan.length) { msg.textContent = 'Completa los campos marcados con *.'; msg.className = 'err'; return; }
    if (!/^(10|15|17|20)\d{9}$/.test(d.ruc)) { msg.textContent = 'El RUC debe tener 11 dígitos.'; msg.className = 'err'; return; }
    if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(d.email)) { msg.textContent = 'Revisa tu correo.'; msg.className = 'err'; return; }
    if (firma.vacia) { msg.textContent = 'Firma en el recuadro para confirmar.'; msg.className = 'err'; return; }
    if (!d.acepto) { msg.textContent = 'Marca la casilla de confirmación.'; msg.className = 'err'; return; }
    store.set('tli_cliente', Object.assign(cliente(), { ruc: d.ruc, razon: d.razon, contacto: d.contacto, celular: d.celular, email: d.email, direccion: d.direccion, ciudad: d.ciudad }));
    if (!CFG.api) { msg.textContent = 'La app aún no está conectada al sistema de órdenes. Escríbenos por WhatsApp.'; msg.className = 'err'; return; }
    const btn = $('#enviar'); btn.disabled = true; btn.textContent = 'Enviando…';
    const payload = {
      accion: 'orden', ...d, origen: origen(), sesion: SES,
      items: Object.entries(S.cart).map(([id, cantidad]) => ({ id, cantidad })),
      firma: $('#firma').toDataURL('image/png'),
    };
    try {
      const r = await (await fetch(CFG.api, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify(payload) })).json();
      if (!r.ok) throw new Error(r.error || 'Error');
      store.set('tli_ultimo', { numero: r.numero, fecha: new Date().toLocaleDateString('es-PE'), items: payload.items });
      ultimoSync = '[]'; S.cart = {}; guardarCart();
      $('#co').innerHTML = `<div class="cobox" style="text-align:center;padding-top:60px">
        <div style="font-size:64px">✅</div><h1>¡Orden enviada!</h1>
        <p style="font-size:18px">Tu orden <b>N° ${esc(r.numero)}</b> por <b>${money(r.total)}</b> llegó a <b>${esc(d.email)}</b>.</p>
        <p class="meta">Un asesor de Tecno Log IA te confirmará stock, pago y despacho.</p>
        <p><a class="btn green" target="_blank" rel="noopener" href="https://wa.me/${esc(S.config.whatsapp || '51928146619')}?text=${encodeURIComponent('Hola, acabo de enviar la orden ' + r.numero + ' desde el catálogo digital.')}">Avisar por WhatsApp</a></p>
        <p><button class="btn ghost" data-cocerrar>Volver al catálogo</button></p></div>`;
    } catch (e) {
      btn.disabled = false; btn.textContent = 'Enviar orden de compra';
      msg.textContent = 'No se pudo enviar (' + e.message + '). Intenta de nuevo o escríbenos por WhatsApp.'; msg.className = 'err';
    }
  }

  // ---------- Cotización (sin firma; no compromete la compra) ----------
  function abrirCotizar() {
    cerrarCart();
    const c = store.get('tli_cliente', {});
    $('#co').innerHTML = `<div class="cobox">
      <button class="pvclose" data-cocerrar aria-label="Cerrar">✕</button>
      <h1 style="margin:0 0 4px">Solicitar cotización</h1>
      <p class="meta" style="margin:0 0 14px">Solo distribuidores y mayoristas · Recibe tu cotización al instante, descárgala en PDF y envíala por WhatsApp. No es una orden de compra.</p>
      <form class="form" id="fc" novalidate>
        <label>RUC *<input name="ruc" inputmode="numeric" maxlength="11" required value="${esc(c.ruc || '')}" placeholder="20xxxxxxxxx"><span id="rucmsg" class="meta"></span></label>
        <label>Razón social *<input name="razon" required value="${esc(c.razon || '')}"></label>
        <label>Persona de contacto *<input name="contacto" required value="${esc(c.contacto || '')}"></label>
        <label>Celular / WhatsApp *<input name="celular" inputmode="tel" required value="${esc(c.celular || '')}"></label>
        <label>Correo (opcional)<input name="email" type="email" value="${esc(c.email || '')}"></label>
        <label>Ciudad / distrito<input name="ciudad" value="${esc(c.ciudad || '')}"></label>
        <label class="full">Notas (opcional)<input name="notas" placeholder="Ej.: para licitación, plazo de entrega, forma de pago…"></label>
        <div class="full"><button class="btn gold" style="width:100%;font-size:18px">Ver mi cotización</button><p id="comsg" class="meta"></p></div>
      </form></div>`;
    $('#co').hidden = false; document.body.style.overflow = 'hidden';
    const f = $('#fc'), ruc = f.ruc;
    ruc.addEventListener('input', () => { ruc.value = ruc.value.replace(/\D/g, ''); if (ruc.value.length === 11) consultarRucEn(f, ruc.value); });
    f.onsubmit = generarCotizacion;
  }
  async function consultarRucEn(f, ruc) {
    const msg = $('#rucmsg');
    if (!/^(10|15|17|20)\d{9}$/.test(ruc)) { msg.textContent = 'Revisa el RUC (11 dígitos).'; msg.className = 'err'; return; }
    if (!CFG.api) return;
    msg.textContent = 'Buscando…'; msg.className = 'meta';
    try {
      const r = await (await fetch(CFG.api + '?accion=ruc&ruc=' + ruc)).json();
      if (r.ok && r.encontrado) {
        if (!f.razon.value) f.razon.value = r.razon || '';
        if (!f.ciudad.value) f.ciudad.value = r.ciudad || '';
        msg.textContent = '✓ Cliente registrado'; msg.className = 'ok';
      } else { msg.textContent = 'Cliente nuevo: completa tus datos.'; msg.className = 'meta'; }
    } catch { msg.textContent = ''; }
  }
  function generarCotizacion(ev) {
    ev.preventDefault();
    const f = $('#fc'), msg = $('#comsg');
    const d = Object.fromEntries(new FormData(f).entries());
    if (['ruc', 'razon', 'contacto', 'celular'].some(k => !String(d[k] || '').trim())) { msg.textContent = 'Completa los campos marcados con *.'; msg.className = 'err'; return; }
    if (!/^(10|15|17|20)\d{9}$/.test(d.ruc)) { msg.textContent = 'El RUC debe tener 11 dígitos.'; msg.className = 'err'; return; }
    if (d.email && !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(d.email)) { msg.textContent = 'Revisa tu correo.'; msg.className = 'err'; return; }
    store.set('tli_cliente', Object.assign(cliente(), { ruc: d.ruc, razon: d.razon, contacto: d.contacto, celular: d.celular, email: d.email || cliente().email || '', ciudad: d.ciudad || cliente().ciudad || '' }));
    const hoy = new Date(), dias = Number(S.config.validez || 7);
    const vence = new Date(hoy.getTime() + dias * 864e5);
    const p2 = n => String(n).padStart(2, '0');
    const numero = 'COT-' + String(hoy.getFullYear()).slice(2) + p2(hoy.getMonth() + 1) + p2(hoy.getDate()) + '-' + Math.random().toString(36).slice(2, 6).toUpperCase();
    const items = Object.entries(S.cart).filter(([id]) => S.porId[id]).map(([id, cantidad]) => ({ id, cantidad }));
    const total = totalCart();
    const cots = store.get('tli_cots', []); cots.unshift({ numero, fecha: hoy.toISOString(), total, items }); store.set('tli_cots', cots.slice(0, 20));
    track('cotizacion', numero + ' · ' + d.ruc, Math.round(total));
    enviarEventos(); sincronizarCarrito(true);
    // Registra la cotización: si el RUC es nuevo, el cliente entra a «Clientes» como «Nuevo» y a la lista «Hoy» del asesor
    if (CFG.api) try { fetch(CFG.api, { method: 'POST', mode: 'no-cors', keepalive: true, headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ accion: 'cotizacion', numero, sesion: SES, origen: origen(), ruc: d.ruc, razon: d.razon, contacto: d.contacto, celular: d.celular, email: d.email || '', ciudad: d.ciudad || '', notas: d.notas || '', items }) }); } catch { /* sin conexión */ }
    const enlace = location.origin + location.pathname + '?d=cot&ruc=' + d.ruc + '&pedido=' + encodeURIComponent(items.map(x => x.id + '*' + x.cantidad).join(','));
    const wa = 'Hola Tecno Log IA, solicito la cotización ' + numero + ' (' + d.razon + ', RUC ' + d.ruc + '):\n' +
      items.map(x => '• ' + x.cantidad + ' × ' + S.porId[x.id].codigo + ' · ' + S.porId[x.id].nombre).join('\n') +
      '\nTotal referencial: ' + money(total) + ' (inc. IGV)' + (d.notas ? '\nNotas: ' + d.notas : '') + '\nVer: ' + enlace;
    const fmt = x => x.toLocaleDateString('es-PE', { day: '2-digit', month: '2-digit', year: 'numeric' });
    $('#co').innerHTML = `<div class="cobox">
      <div class="noprint" style="display:flex;justify-content:flex-end"><button class="pvclose" style="position:static" data-cocerrar aria-label="Cerrar">✕</button></div>
      <article class="cotdoc" id="cotdoc">
        <header><img src="assets/logo-wordmark.png" alt="Tecno Log IA"><div><b>COTIZACIÓN</b><br>N° ${esc(numero)}</div></header>
        <p class="rib">SOLO PARA DISTRIBUIDORES Y MAYORISTAS</p>
        <div class="cgrid">
          <div><span>Cliente</span><b>${esc(d.razon)}</b></div><div><span>RUC</span><b>${esc(d.ruc)}</b></div>
          <div><span>Contacto</span><b>${esc(d.contacto)}</b></div><div><span>Ciudad</span><b>${esc(d.ciudad || '—')}</b></div>
          <div><span>Fecha</span><b>${fmt(hoy)}</b></div><div><span>Válida hasta</span><b>${fmt(vence)}</b></div>
        </div>
        <table><thead><tr><th>Código</th><th>Descripción</th><th>Cant.</th><th>P. unit.</th><th>Subtotal</th></tr></thead><tbody>
          ${items.map(x => { const p = S.porId[x.id]; return `<tr><td>${esc(p.codigo)}</td><td>${esc(p.nombre)}${p.marca ? ' · ' + esc(p.marca) : ''}</td><td>${x.cantidad}</td><td>${money(p.pvp)}</td><td>${money(p.pvp * x.cantidad)}</td></tr>`; }).join('')}
        </tbody><tfoot><tr><td colspan="4">TOTAL (incluye IGV)</td><td>${money(total)}</td></tr></tfoot></table>
        ${d.notas ? `<p><b>Notas:</b> ${esc(d.notas)}</p>` : ''}
        <p class="cond">Precios en soles, incluyen IGV. Válida ${dias} días o hasta agotar stock. Sujeta a confirmación de stock, condiciones de pago y despacho por Tecno Log IA. Esta cotización no es una orden de compra.</p>
        <p class="cond">Tecno Log IA · WhatsApp ${esc((S.config.whatsapp || '51928146619').replace(/^51/, ''))} · ${esc(S.config.correo || 'tecnologiapperu@gmail.com')}</p>
      </article>
      <div class="noprint cotacc">
        <a class="btn green" target="_blank" rel="noopener" href="https://wa.me/${esc(S.config.whatsapp || '51928146619')}?text=${encodeURIComponent(wa)}">Enviar por WhatsApp</a>
        <button class="btn gold" data-cotpdf>Descargar PDF</button>
        <button class="btn red" data-orden>Convertir en orden de compra</button>
        <button class="btn ghost" data-cocerrar>Seguir viendo el catálogo</button>
      </div></div>`;
    $('#co').scrollTop = 0;
  }
  function cerrarOrden() { $('#co').hidden = true; $('#co').innerHTML = ''; document.body.style.overflow = ''; }

  window.addEventListener('afterprint', () => document.body.classList.remove('pcot'));

  // ---------- Aviso ----------
  let toastT;
  function toast(t) { const el = $('#toast'); el.textContent = t; el.hidden = false; clearTimeout(toastT); toastT = setTimeout(() => el.hidden = true, 1800); }

  // ---------- Eventos ----------
  let qT;
  $('#q').addEventListener('input', e => {
    clearTimeout(qT);
    qT = setTimeout(() => irA({ q: e.target.value.trim(), m: S.filtro.marca, c: S.filtro.cat }), 250);
  });
  $('#btnCart').onclick = abrirCart;
  window.addEventListener('hashchange', router);

  document.addEventListener('click', e => {
    const t = e.target.closest('button,a') || e.target.closest('article,[data-zoom],[data-ver]');
    if (!t) return;
    const ds = t.dataset;
    if (ds.add) { e.stopPropagation(); agregar(ds.add, 1); return; }
    if (ds.addc) { e.stopPropagation(); agregar(ds.addc, 1); abrirCart(); return; }
    if ('repetir' in ds) { const u = store.get('tli_ultimo', null); if (u) { u.items.forEach(x => { if (S.porId[x.id] && S.porId[x.id].estado !== 'Agotado') S.cart[x.id] = (S.cart[x.id] || 0) + x.cantidad; }); guardarCart(); abrirCart(); } return; }
    if (ds.asesor !== undefined || ds.as) { asesorClick(t, ds); return; }
    if (t.matches('article.card')) { abrirProducto(t.dataset.id); return; }
    if (ds.of) { S.oferta = ds.of; vistaInicio(); return; }
    if (ds.m && !t.closest('#pv')) { irA({ m: ds.m }); return; }
    if (ds.c) { cerrarProducto(); cerrarCart(); irA({ c: ds.c }); return; }
    if ('inicio' in ds) { irA({}); return; }
    if (ds.quitar) { const f = { q: S.filtro.q, m: S.filtro.marca, c: S.filtro.cat }; f[ds.quitar] = ''; irA(f); return; }
    if (ds.play) { abrirPlay(ds.play === 'lista' ? S.lista : S.lista); return; }
    if (t.id === 'mas') { paginaActual++; pintarPagina(); return; }
    // producto
    if ('cerrar' in ds) { cerrarProducto(); return; }
    if (ds.th) { pvEstado.i = +ds.th; pvEstado.zoom = false; pintarMedia(); return; }
    if (ds.m && t.closest('#pv')) { const n = pvEstado.media.length; pvEstado.i = (pvEstado.i + +ds.m + n) % n; pvEstado.zoom = false; pintarMedia(); return; }
    if ('zoom' in ds) { pvEstado.zoom = !pvEstado.zoom; pintarMedia(); return; }
    if (ds.q) { const i = $('#pvq'); i.value = Math.max(1, (+i.value || 1) + +ds.q); return; }
    if (ds.addq) { agregar(ds.addq, $('#pvq').value, true); abrirProducto(ds.addq); return; }
    if (ds.ver) { if (t.closest('#cart')) cerrarCart(); if (t.closest('#asesor')) $('#asesor').hidden = true; abrirProducto(ds.ver); return; }
    if (ds.nav) { const pos = S.lista.findIndex(x => x.id === pvEstado.id) + +ds.nav; if (S.lista[pos]) abrirProducto(S.lista[pos].id); return; }
    // presentación
    if (ds.pl) {
      if (ds.pl === 'cerrar') cerrarPlay();
      else if (ds.pl === 'pausa') { pausado = !pausado; pintarPlay(); }
      else if (ds.pl === 'ver') { const id = playL[playI].id; pausado = true; clearTimeout(playT); cerrarPlay(); abrirProducto(id); }
      else { playI = (playI + +ds.pl + playL.length) % playL.length; pintarPlay(); }
      return;
    }
    // carrito
    if ('ccerrar' in ds) { cerrarCart(); return; }
    if (ds.del) { delete S.cart[ds.del]; guardarCart(); abrirCart(); return; }
    if (ds.cq) { S.cart[ds.cq] = Math.max(1, (S.cart[ds.cq] || 1) + +ds.d); guardarCart(); abrirCart(); return; }
    if ('orden' in ds) { abrirOrden(); return; }
    if ('cotizar' in ds) { abrirCotizar(); return; }
    if ('cotpdf' in ds) { document.body.classList.add('pcot'); track('cotizacion_pdf', ''); setTimeout(() => { window.print(); }, 50); return; }
    if ('cocerrar' in ds) { cerrarOrden(); return; }
  });
  document.addEventListener('change', e => {
    const id = e.target.dataset && e.target.dataset.ci;
    if (id) { S.cart[id] = Math.max(1, parseInt(e.target.value, 10) || 1); guardarCart(); abrirCart(); }
  });
  document.addEventListener('keydown', e => {
    if (!$('#play').hidden) { if (e.key === 'Escape') cerrarPlay(); if (e.key === 'ArrowRight') { playI = (playI + 1) % playL.length; pintarPlay(); } if (e.key === 'ArrowLeft') { playI = (playI - 1 + playL.length) % playL.length; pintarPlay(); } return; }
    if (!$('#pv').hidden) {
      if (e.key === 'Escape') cerrarProducto();
      if (e.key === 'ArrowRight' && pvEstado.media.length > 1) { pvEstado.i = (pvEstado.i + 1) % pvEstado.media.length; pintarMedia(); }
      if (e.key === 'ArrowLeft' && pvEstado.media.length > 1) { pvEstado.i = (pvEstado.i - 1 + pvEstado.media.length) % pvEstado.media.length; pintarMedia(); }
    }
    if (e.key === 'Escape') { cerrarCart(); }
  });
  // deslizar en la galería (celular)
  let x0 = null;
  document.addEventListener('touchstart', e => { if (e.target.closest('#stage')) x0 = e.touches[0].clientX; }, { passive: true });
  document.addEventListener('touchend', e => {
    if (x0 == null) return;
    const dx = e.changedTouches[0].clientX - x0; x0 = null;
    if (Math.abs(dx) > 50 && pvEstado.media && pvEstado.media.length > 1) {
      const n = pvEstado.media.length; pvEstado.i = (pvEstado.i + (dx < 0 ? 1 : -1) + n) % n; pvEstado.zoom = false; pintarMedia();
    }
  });

  // ---------- Instalar como app (ícono en el celular) ----------
  let promptInst = null;
  const instalada = () => window.matchMedia('(display-mode: standalone)').matches || navigator.standalone;
  const esIOS = /iphone|ipad|ipod/i.test(navigator.userAgent);
  const mostrarInst = () => { if (!instalada() && !store.get('tli_inst_no', false)) $('#inst').hidden = false; };
  window.addEventListener('beforeinstallprompt', e => { e.preventDefault(); promptInst = e; mostrarInst(); });
  if (esIOS) mostrarInst();
  $('#btnInst').onclick = async () => {
    if (promptInst) { promptInst.prompt(); const r = await promptInst.userChoice; if (r.outcome === 'accepted') $('#inst').hidden = true; promptInst = null; }
    else if (esIOS) $('#ios').hidden = false;
    else toast('Abre el menú ⋮ del navegador y elige «Agregar a la pantalla principal»');
  };
  $('#instNo').onclick = () => { $('#inst').hidden = true; store.set('tli_inst_no', true); };
  $('#iosOk').onclick = () => { $('#ios').hidden = true; };
  window.addEventListener('appinstalled', () => { $('#inst').hidden = true; toast('✓ Catálogo instalado'); });

  // ---------- Asesor IA (Gemini, desde la API) ----------
  let asH = (() => { try { return JSON.parse(sessionStorage.getItem('tli_asesor')) || []; } catch { return []; } })();
  const guardarAs = () => { try { sessionStorage.setItem('tli_asesor', JSON.stringify(asH.slice(-20))); } catch { /* */ } };
  function mostrarAsesor() { $('#btnAs').hidden = !(CFG.api && S.config.asesor); }
  function pintarAsesor() {
    const box = $('#asesor');
    const msgs = asH.length ? asH : [{ rol: 'asesor', texto: '¡Hola! Soy el asesor de Tecno Log IA. Cuéntame qué negocio tienes o qué necesitas (por ejemplo: «voy a abrir una heladería» o «necesito equipos para una pollería») y te armo un pedido con productos del catálogo.' }];
    box.innerHTML = `<header><b>✨ Asesor IA</b><button data-as="cerrar" aria-label="Cerrar">✕</button></header>
      <div class="asmsgs" id="asmsgs">${msgs.map((m, k) => `<div class="am ${m.rol}">${esc(m.texto).replace(/\n/g, '<br>')}
        ${(m.productos || []).length ? `<div class="asprod">${m.productos.filter(x => S.porId[x.id]).map(x => { const p = S.porId[x.id]; return `
          <div class="aspi" data-ver="${esc(p.id)}">${imgHTML(p)}<div><b>${x.cantidad} × ${esc(p.nombre)}</b><br><span class="meta">${esc(p.codigo)} · ${money(p.pvp)} c/u</span>${x.motivo ? `<br><small>${esc(x.motivo)}</small>` : ''}</div></div>`; }).join('')}
          <button class="btn red" data-as="todo" data-k="${k}">Agregar todo al carrito</button></div>` : ''}</div>`).join('')}
        ${asOcupado ? '<div class="am asesor">Pensando…</div>' : ''}</div>
      <form class="asform" id="asform"><input id="asq" placeholder="Escribe tu consulta…" autocomplete="off" maxlength="500"><button class="btn red" ${asOcupado ? 'disabled' : ''}>Enviar</button></form>
      <small class="meta" style="padding:0 12px 10px">Respuestas generadas con IA a partir del catálogo; confirma stock y detalles con un asesor.</small>`;
    $('#asform').onsubmit = ev => { ev.preventDefault(); preguntar($('#asq').value); };
    const m = $('#asmsgs'); m.scrollTop = m.scrollHeight;
  }
  let asOcupado = false;
  async function preguntar(q) {
    q = String(q || '').trim(); if (!q || asOcupado) return;
    asH.push({ rol: 'cliente', texto: q }); asOcupado = true; pintarAsesor();
    try {
      const r = await (await fetch(CFG.api, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify({ accion: 'asesor', sesion: SES, origen: origen(), pregunta: q, historial: asH.slice(-7, -1).map(m => ({ rol: m.rol, texto: m.texto })) }) })).json();
      asH.push(r.ok ? { rol: 'asesor', texto: r.mensaje, productos: r.productos } : { rol: 'asesor', texto: r.error || 'No pude responder ahora. Escríbenos por WhatsApp.' });
    } catch { asH.push({ rol: 'asesor', texto: 'Sin conexión. Intenta de nuevo o escríbenos por WhatsApp.' }); }
    asOcupado = false; guardarAs(); pintarAsesor();
  }
  function asesorClick(t, ds) {
    if (t.id === 'btnAs') { $('#asesor').hidden = false; pintarAsesor(); setTimeout(() => $('#asq') && $('#asq').focus(), 50); return; }
    if (ds.as === 'cerrar') { $('#asesor').hidden = true; return; }
    if (ds.as === 'todo') { (asH[+ds.k].productos || []).forEach(x => { if (S.porId[x.id]) S.cart[x.id] = Math.max(S.cart[x.id] || 0, x.cantidad); }); guardarCart(); toast('✓ Productos agregados al carrito'); $('#asesor').hidden = true; abrirCart(); }
  }

  if ('serviceWorker' in navigator) navigator.serviceWorker.register('sw.js').catch(() => {});
  cargar();
})();
