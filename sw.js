/* Modo sin conexión: guarda la app completa en el dispositivo. */
const VERSION = 'folio-v2';
const FILES = [
  './', 'index.html', 'app.css', 'manifest.webmanifest',
  'js/core.js', 'js/word.js', 'js/excel.js', 'js/ppt.js', 'js/app.js',
  'vendor/jszip.min.js', 'vendor/docx-preview.min.js', 'vendor/ssf.js', 'vendor/hyperformula.full.min.js',
  'assets/nuevo.docx', 'assets/nuevo.xlsx', 'assets/nuevo.pptx',
  'icons/icon-192.png', 'icons/icon-512.png', 'icons/maskable-512.png',
];
self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(VERSION).then((c) => c.addAll(FILES)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', (e) => {
  e.waitUntil(caches.keys().then((keys) => Promise.all(keys.filter((k) => k !== VERSION).map((k) => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', (e) => {
  if (e.request.method !== 'GET') return;
  e.respondWith(
    caches.match(e.request, { ignoreSearch: true }).then((hit) => hit || fetch(e.request).then((res) => {
      if (res.ok && new URL(e.request.url).origin === location.origin) { const copy = res.clone(); caches.open(VERSION).then((c) => c.put(e.request, copy)); }
      return res;
    }).catch(() => caches.match('index.html')))
  );
});
