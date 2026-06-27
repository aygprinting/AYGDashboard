// AYG Printing — Service Worker
// Caches only the static app shell (HTML/icons/manifest) so the app installs
// and launches instantly offline. Deliberately does NOT cache anything from
// Microsoft Graph / SharePoint / login.microsoftonline.com — those must
// always go to the network so the dashboard never shows stale business data.

const CACHE_NAME = 'ayg-printing-shell-v1';
const SHELL_FILES = [
  './',
  './index.html',
  './manifest.json',
  './icon-192.png',
  './icon-512.png',
  './icon-192-maskable.png',
  './icon-512-maskable.png'
];

self.addEventListener('install', function (event) {
  event.waitUntil(
    caches.open(CACHE_NAME).then(function (cache) {
      return cache.addAll(SHELL_FILES);
    })
  );
  self.skipWaiting();
});

self.addEventListener('activate', function (event) {
  event.waitUntil(
    caches.keys().then(function (keys) {
      return Promise.all(
        keys.filter(function (k) { return k !== CACHE_NAME; })
            .map(function (k) { return caches.delete(k); })
      );
    })
  );
  self.clients.claim();
});

self.addEventListener('fetch', function (event) {
  var url = event.request.url;

  // Never intercept anything that isn't a simple same-origin GET for our own
  // shell files. This excludes Microsoft login/Graph/SharePoint calls, the
  // Excel data fetches, and any POST/OAuth redirect traffic automatically.
  if (event.request.method !== 'GET') return;
  if (url.indexOf(self.location.origin) !== 0) return;
  if (url.indexOf('login.microsoftonline.com') !== -1) return;
  if (url.indexOf('graph.microsoft.com') !== -1) return;
  if (url.indexOf('sharepoint.com') !== -1) return;

  var isShellFile = SHELL_FILES.some(function (f) {
    return url.indexOf(f.replace('./', '')) !== -1;
  });
  if (!isShellFile) return;

  // Shell files: try cache first for instant load, fall back to network,
  // and refresh the cache in the background so updates still arrive.
  event.respondWith(
    caches.match(event.request).then(function (cached) {
      var networkFetch = fetch(event.request).then(function (response) {
        if (response && response.status === 200) {
          var copy = response.clone();
          caches.open(CACHE_NAME).then(function (cache) {
            cache.put(event.request, copy);
          });
        }
        return response;
      }).catch(function () {
        return cached;
      });
      return cached || networkFetch;
    })
  );
});
