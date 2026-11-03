/* Modo sin conexión: guarda la app completa en el dispositivo. */
const VERSION = 'folio-v8';
const FILES = [
  './', 'index.html', 'app.css', 'manifest.webmanifest',
  'js/core.js', 'js/common-plus.js', 'js/chart.js', 'js/word.js', 'js/word-plus.js', 'js/word-plus2.js', 'js/excel.js', 'js/excel-plus.js', 'js/excel-cf.js', 'js/ppt.js', 'js/ppt-plus.js', 'js/ppt-anim.js', 'js/app.js', 'js/convert-worker.js',
  'vendor/jszip.min.js', 'vendor/docx-preview.min.js', 'vendor/ssf.js', 'vendor/hyperformula.full.min.js', 'vendor/xlsx.full.min.js',
  'assets/nuevo.docx', 'assets/nuevo.xlsx', 'assets/nuevo.pptx',
  'assets/plantillas/factura.xlsx', 'assets/plantillas/cotizacion.docx', 'assets/plantillas/presupuesto.xlsx', 'assets/plantillas/inventario.xlsx', 'assets/plantillas/informe.docx', 'assets/plantillas/presentacion.pptx',
  'icons/icon-192.png', 'icons/icon-512.png', 'icons/maskable-512.png',
];
self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(VERSION).then((c) => c.addAll(FILES)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', (e) => {
  e.waitUntil(caches.keys().then((keys) => Promise.all(keys.filter((k) => k !== VERSION && k !== 'folio-compartido').map((k) => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', (e) => {
  // Compartir a Folio (Android): el archivo llega por POST, se guarda y se abre en la app
  const url = new URL(e.request.url);
  if (e.request.method === 'POST' && url.pathname.endsWith('/compartir')) {
    e.respondWith((async () => {
      try {
        const data = await e.request.formData();
        const files = data.getAll('archivos').filter((f) => f && f.size);
        const cache = await caches.open('folio-compartido');
        if (files[0]) await cache.put('compartido', new Response(files[0], { headers: { 'x-nombre': encodeURIComponent(files[0].name || 'archivo') } }));
      } catch (err) {}
      return Response.redirect(new URL('./?compartido=1', self.registration.scope).href, 303);
    })());
    return;
  }
  if (e.request.method !== 'GET') return;
  e.respondWith(
    caches.match(e.request, { ignoreSearch: true }).then((hit) => hit || fetch(e.request).then((res) => {
      if (res.ok && new URL(e.request.url).origin === location.origin) { const copy = res.clone(); caches.open(VERSION).then((c) => c.put(e.request, copy)); }
      return res;
    }).catch(() => caches.match('index.html')))
  );
});
