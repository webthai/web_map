// ===== Service Worker: caches the app shell so it loads with no internet =====
// NOTE: map tiles and routing graph data are cached separately in IndexedDB
// (handled inside index.html), not here — this only caches the app shell.

var SHELL_CACHE = 'map-app-shell-v2';
var SHELL_FILES = [
  './',
  './index.html',
  './1.png',
  'https://unpkg.com/leaflet@1.9.4/dist/leaflet.css',
  'https://unpkg.com/leaflet@1.9.4/dist/leaflet.js'
];

self.addEventListener('install', function (event) {
  // NOTE: no self.skipWaiting() here on purpose — the new worker waits until
  // the page (via applyUpdate() in index.html) tells it to take over, so the
  // person gets a chance to see the "update available" banner first.
  event.waitUntil(
    caches.open(SHELL_CACHE).then(function (cache) {
      // cache.addAll() is all-or-nothing — if even one URL fails (e.g. a
      // brief CDN hiccup), NOTHING gets cached. Cache each file separately
      // instead so a single failure doesn't cost us the whole app shell.
      return Promise.all(
        SHELL_FILES.map(function (url) {
          return cache.add(url).catch(function () { /* skip this one, keep the rest */ });
        })
      );
    })
  );
});

self.addEventListener('message', function (event) {
  if (event.data && event.data.type === 'SKIP_WAITING') {
    self.skipWaiting();
  }
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
