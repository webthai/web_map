// ===== Service Worker: caches the app shell so it loads with no internet =====
// NOTE: map tiles and routing graph data are cached separately in IndexedDB
// (handled inside index.html), not here — this only caches the app shell.

var SHELL_CACHE = 'map-app-shell-v1';
var SHELL_FILES = [
  './',
  './index.html',
  './1.png',
  'https://unpkg.com/leaflet@1.9.4/dist/leaflet.css',
  'https://unpkg.com/leaflet@1.9.4/dist/leaflet.js'
];

self.addEventListener('install', function (event) {
  self.skipWaiting();
  event.waitUntil(
    caches.open(SHELL_CACHE).then(function (cache) {
      return cache.addAll(SHELL_FILES).catch(function () {
        // ignore individual failures (e.g. offline on first install)
      });
    })
  );
});

self.addEventListener('activate', function (event) {
  event.waitUntil(
    caches.keys().then(function (keys) {
      return Promise.all(
        keys.filter(function (k) { return k !== SHELL_CACHE; })
            .map(function (k) { return caches.delete(k); })
      );
    })
  );
  self.clients.claim();
});

// Cache-first for the app shell; network-first fallback for everything else
// (Google Apps Script API calls always go to network — never cached here).
self.addEventListener('fetch', function (event) {
  var url = event.request.url;
  if (url.indexOf('script.google.com') !== -1 || url.indexOf('script.googleusercontent.com') !== -1) {
    return; // never intercept API calls
  }
  event.respondWith(
    caches.match(event.request).then(function (cached) {
      if (cached) return cached;
      return fetch(event.request).then(function (response) {
        if (response && response.ok && (event.request.method === 'GET')) {
          var copy = response.clone();
          caches.open(SHELL_CACHE).then(function (cache) { cache.put(event.request, copy); });
        }
        return response;
      }).catch(function () {
        if (event.request.mode === 'navigate') return caches.match('./index.html');
      });
    })
  );
});
