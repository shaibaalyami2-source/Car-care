/* CRM Clean — service worker: the app works offline, map tiles and fonts are kept for later */
const V = 'crm-clean-3.0-eaa5ff856d';
const CORE = ['./', './index.html', './manifest.webmanifest', './config.js', './icons/icon-192.png', './icons/icon-512.png'];
const TILES = 'crm-tiles', FONTS = 'crm-fonts', MAX_TILES = 900;

self.addEventListener('install', (e) => { e.waitUntil(caches.open(V).then((c) => c.addAll(CORE))); });
self.addEventListener('activate', (e) => {
  e.waitUntil(caches.keys().then((ks) => Promise.all(ks.filter((k) => k.startsWith('crm-clean-') && k !== V).map((k) => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('message', (e) => { if (e.data === 'skip') self.skipWaiting(); });

async function trim(name, max) { const c = await caches.open(name); const ks = await c.keys(); for (let i = 0; i < ks.length - max; i++) await c.delete(ks[i]); }
async function cacheFirst(req, name, limit) {
  const c = await caches.open(name); const hit = await c.match(req); if (hit) return hit;
  try { const res = await fetch(req); if (res && (res.ok || res.type === 'opaque')) { c.put(req, res.clone()); if (limit) trim(name, limit); } return res; }
  catch (err) { return hit || Response.error(); }
}
async function staleWhileRevalidate(req, name) {
  const c = await caches.open(name); const hit = await c.match(req);
  const net = fetch(req).then((res) => { if (res && (res.ok || res.type === 'opaque')) c.put(req, res.clone()); return res; }).catch(() => null);
  return hit || (await net) || Response.error();
}
self.addEventListener('fetch', (e) => {
  const req = e.request; if (req.method !== 'GET') return;
  const url = new URL(req.url);
  if (/script\.google(usercontent)?\.com$|mymemory\.translated\.net$|googleapis\.com$/.test(url.hostname) && !/fonts\.googleapis\.com$/.test(url.hostname)) return; // live data: always the network
  if (/tile\.openstreetmap\.org$|basemaps\.cartocdn\.com$|tile\./.test(url.hostname)) { e.respondWith(cacheFirst(req, TILES, MAX_TILES)); return; }
  if (/fonts\.(googleapis|gstatic)\.com$/.test(url.hostname)) { e.respondWith(staleWhileRevalidate(req, FONTS)); return; }
  if (url.origin !== location.origin) return;
  if (req.mode === 'navigate') { e.respondWith(caches.open(V).then((c) => c.match('./index.html')).then((hit) => hit || fetch(req)).catch(() => caches.match('./index.html'))); return; }
  if (url.pathname.endsWith('/config.js')) { e.respondWith(fetch(req).then((res) => { const copy = res.clone(); caches.open(V).then((c) => c.put(req, copy)); return res; }).catch(() => caches.match(req))); return; }
  e.respondWith(cacheFirst(req, V));
});
