// ================================================================
// Research Assessor — Service Worker (app instalable / PWA)
// • La app (index.html) se pide primero a la red: siempre abre la última
//   versión publicada; sin conexión usa la copia guardada.
// • Íconos, manifiesto y JSZip se sirven desde caché.
// • Las llamadas al Apps Script (datos) NUNCA se guardan aquí.
// Al publicar cambios importantes, sube el número de VERSION.
// ================================================================
const VERSION = 'ra-v7.2';
const APP_SHELL = [
  './',
  './index.html',
  './manifest.webmanifest',
  './icons/icon-192.png',
  './icons/icon-512.png',
  './icons/apple-touch-icon.png'
];
const CDN_JSZIP = 'https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js';

self.addEventListener('install', function (e) {
  e.waitUntil(
    caches.open(VERSION).then(function (c) {
      return c.addAll(APP_SHELL).then(function () {
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
  if (url.origin === self.location.origin || req.url === CDN_JSZIP) {
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
