// ================================================================
// Research Assessor — Service Worker (app instalable / PWA)
// • La app (index.html) se pide primero a la red: siempre abre la última
//   versión publicada; sin conexión usa la copia guardada.
// • Íconos, manifiesto y JSZip se sirven desde caché.
// • Las llamadas al Apps Script (datos) NUNCA se guardan aquí.
// Al publicar cambios importantes, sube el número de VERSION.
// ================================================================
const VERSION = 'ra-v7.6.1';
const APP_SHELL = [
  './',
  './index.html',
  './manifest.webmanifest'
];
// Íconos (en la raíz, junto a index.html). Si alguno falta, el servicio se instala igual.
const ICONOS = ['./icon-192.png', './icon-512.png', './apple-touch-icon.png'];
const CDN_JSZIP = 'https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js';
// Librerías de los contratos (QR y lectura de firmas electrónicas): se guardan al usarlas
const CDN_LIBS = [CDN_JSZIP,
  'https://cdnjs.cloudflare.com/ajax/libs/qrcode-generator/1.4.4/qrcode.min.js',
  'https://cdnjs.cloudflare.com/ajax/libs/forge/1.3.1/forge.min.js',
  'https://cdnjs.cloudflare.com/ajax/libs/pdf-lib/1.17.1/pdf-lib.min.js',
  'https://cdn.jsdelivr.net/npm/@pdf-lib/fontkit@1.1.1/dist/fontkit.umd.min.js',
  'https://cdn.jsdelivr.net/gh/google/fonts@main/ofl/crimsontext/CrimsonText-Regular.ttf',
  'https://cdn.jsdelivr.net/gh/google/fonts@main/ofl/crimsontext/CrimsonText-Bold.ttf'];

self.addEventListener('install', function (e) {
  e.waitUntil(
    caches.open(VERSION).then(function (c) {
      return c.addAll(APP_SHELL).then(function () {
        return Promise.all(ICONOS.map(function (u) { return c.add(u).catch(function () {}); }));
      }).then(function () {
        // JSZip es opcional: si falla, la app funciona igual con conexión
        return c.add(CDN_JSZIP).catch(function () {});
      });
    })
  );
});

self.addEventListener('activate', function (e) {
  e.waitUntil(
    caches.keys().then(function (keys) {
      return Promise.all(keys.filter(function (k) { return k !== VERSION; }).map(function (k) { return caches.delete(k); }));
    }).then(function () { return self.clients.claim(); })
  );
});

// La página pide activar la versión nueva inmediatamente
self.addEventListener('message', function (e) {
  if (e.data === 'activar') self.skipWaiting();
});

self.addEventListener('fetch', function (e) {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);

  // Datos del Apps Script / Google: siempre directo a la red
  if (/(^|\.)google(usercontent)?\.com$/.test(url.hostname)) return;

  // Navegación (abrir la app): red primero, caché si no hay conexión
  if (req.mode === 'navigate') {
    e.respondWith(
      fetch(req).then(function (res) {
        const copia = res.clone();
        caches.open(VERSION).then(function (c) { c.put('./index.html', copia); });
        return res;
      }).catch(function () {
        return caches.match('./index.html').then(function (r) { return r || caches.match('./'); });
      })
    );
    return;
  }

  // Archivos propios y JSZip: caché primero, y se actualiza en segundo plano
  if (url.origin === self.location.origin || CDN_LIBS.indexOf(req.url) > -1) {
    e.respondWith(
      caches.match(req).then(function (hit) {
        const red = fetch(req).then(function (res) {
          if (res && res.ok) {
            const copia = res.clone();
            caches.open(VERSION).then(function (c) { c.put(req, copia); });
          }
          return res;
        }).catch(function () { return hit; });
        return hit || red;
      })
    );
  }
});
