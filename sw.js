// Offline: the app's files come from the cache and are refreshed in the background.
const CACHE = 'study-pacer-v4';
const SHELL = [
  './', 'index.html', 'css/app.css', 'js/app.js', 'js/pacer.js', 'js/viewer.js', 'js/importers.js', 'js/pptx.js', 'js/store.js', 'js/themes.js',
  'vendor/pdf.min.js', 'vendor/pdf.worker.min.js', 'vendor/jszip.min.js', 'vendor/mammoth.browser.min.js',
  'vendor/paper/fine.webp', 'vendor/paper/book.webp', 'vendor/paper/laid.webp', 'vendor/paper/rice.webp', 'vendor/paper/handmade.webp', 'vendor/paper/recycled.webp', 'vendor/paper/natural.webp', 'vendor/paper/linen.webp', 'vendor/paper/felt.webp', 'vendor/paper/card.webp',
  'manifest.webmanifest', 'favicon.ico', 'icons/icon-192x192.png',
];
self.addEventListener('install', (e) => { e.waitUntil(caches.open(CACHE).then((c) => c.addAll(SHELL)).then(() => self.skipWaiting())); });
self.addEventListener('activate', (e) => { e.waitUntil(caches.keys().then((ks) => Promise.all(ks.filter((k) => k !== CACHE).map((k) => caches.delete(k)))).then(() => self.clients.claim())); });
self.addEventListener('fetch', (e) => {
  const req = e.request;
  if (req.method !== 'GET' || new URL(req.url).origin !== location.origin) return;
  const vendor = req.url.includes('/vendor/');
  e.respondWith(caches.open(CACHE).then(async (cache) => {
    const cached = await cache.match(req, { ignoreSearch: true });
    if (vendor && cached) return cached;
    const net = fetch(req, { cache: 'no-cache' }).then((r) => { if (r.ok) cache.put(req, r.clone()); return r; }).catch(() => cached);
    return navigator.onLine === false && cached ? cached : net.then((r) => r || cached);
  }));
});
