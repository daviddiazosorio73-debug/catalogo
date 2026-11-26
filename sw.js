// Caché básico: la app abre rápido y funciona aunque la señal sea débil.
const V = 'tli-v8';
const BASE = ['./', 'index.html', 'styles.css', 'app.js', 'config.js', 'assets/logo-wordmark.png', 'assets/icon-192.png'];
self.addEventListener('install', e => { e.waitUntil(caches.open(V).then(c => c.addAll(BASE))); self.skipWaiting(); });
self.addEventListener('activate', e => { e.waitUntil(caches.keys().then(k => Promise.all(k.filter(x => x !== V).map(x => caches.delete(x))))); self.clients.claim(); });
self.addEventListener('fetch', e => {
  const u = new URL(e.request.url);
  if (e.request.method !== 'GET' || u.origin !== location.origin) return; // API y fotos externas: siempre en línea
  // Red primero (precios y archivos al día); si no hay red, lo guardado.
  e.respondWith(fetch(e.request).then(r => { const c = r.clone(); caches.open(V).then(x => x.put(e.request, c)); return r; })
    .catch(() => caches.match(e.request)));
});
