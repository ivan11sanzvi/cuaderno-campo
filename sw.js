// Guarda la app en el teléfono para que funcione sin cobertura.
const CACHE = 'cuaderno-campo-v21';
const TESELAS = 'cuaderno-campo-teselas';
const MAX_TESELAS = 3000;
const ASSETS = ['./', './index.html', './styles.css', './app.js', './manifest.webmanifest', './icons/icon-192.png', './icons/icon-512.png',
  './vendor/leaflet.js', './vendor/leaflet.css', './vendor/xlsx.full.min.js', './fitos.json'];

// Al instalar, todo se pide al servidor (cache: 'reload'), nunca a la copia vieja del navegador.
self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(ASSETS.map(u => new Request(u, { cache: 'reload' })))).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE && k !== TESELAS).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});

async function recortar() {
  const c = await caches.open(TESELAS); const ks = await c.keys();
  if (ks.length > MAX_TESELAS) await Promise.all(ks.slice(0, ks.length - MAX_TESELAS + 300).map(k => c.delete(k)));
}
let n = 0;
self.addEventListener('fetch', e => {
  if (e.request.method !== 'GET') return;
  const url = new URL(e.request.url);
  // Ortofoto y linderos SIGPAC ya vistos: se guardan para verlos sin cobertura.
  if (url.hostname === 'tms-pnoa-ma.idee.es' || (url.hostname === 'sigpac-hubcloud.es' && url.pathname.startsWith('/wms'))) {
    e.respondWith(caches.open(TESELAS).then(async c => {
      const hit = await c.match(e.request);
      if (hit) return hit;
      const r = await fetch(e.request);
      if (r.ok) { c.put(e.request, r.clone()); if (++n % 100 === 0) recortar(); }
      return r;
    }));
    return;
  }
  if (url.origin !== location.origin) return;
  // App: primero la copia guardada; si hay red, se actualiza para la próxima vez.
  e.respondWith(caches.open(CACHE).then(async cache => {
    const hit = await cache.match(e.request, { ignoreSearch: true });
    const red = fetch(e.request).then(r => { if (r.ok) cache.put(e.request, r.clone()); return r; }).catch(() => hit);
    return hit || red;
  }));
});
