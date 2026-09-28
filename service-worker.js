// AYG Printing — Service Worker
// Caches only the static app shell (HTML/icons/manifest) so the app installs
// and launches instantly offline. Deliberately does NOT cache anything from
// Microsoft Graph / SharePoint / login.microsoftonline.com — those must
// always go to the network so the dashboard never shows stale business data.
//
// UPDATE STRATEGY (20260928-RECON2):
//   • The page itself (index.html / the app URL) is NETWORK-FIRST: every load asks
//     GitHub Pages for the latest build and only falls back to the cached copy when
//     offline. A new index.html therefore shows up on the next load — no need to
//     bump anything here for ordinary dashboard updates.
//   • Icons and manifest stay cache-first (they almost never change).
//   • Bump CACHE_NAME only if you change this file's logic or the icon set.

const CACHE_NAME = 'ayg-printing-shell-20260928-RECON2';
const SHELL_FILES = [
  './',
  './index.html',
  './manifest.json',
  './icon-192.png',
  './icon-512.png',
  './icon-192-maskable.png',
  './icon-512-maskable.png'
];

// Path of a URL relative to this worker's folder, e.g. '' for the app root,
// 'index.html', 'icon-192.png'. Used instead of a substring test — the old
// check matched './' → '' against EVERY same-origin request.
function relPath(url) {
  var scope = new URL(self.registration.scope).pathname;   // e.g. /AYGDashboard/
  var path = new URL(url).pathname;
  return path.indexOf(scope) === 0 ? path.slice(scope.length) : null;
}
var SHELL_REL = SHELL_FILES.map(function (f) { return f.replace('./', ''); });

self.addEventListener('install', function (event) {
  event.waitUntil(
    caches.open(CACHE_NAME).then(function (cache) {
      // cache:'reload' bypasses the browser's HTTP cache, so the install always
      // stores the freshly deployed files rather than a copy up to 10 min old.
      return cache.addAll(SHELL_FILES.map(function (f) {
        return new Request(f, { cache: 'reload' });
      }));
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
    }).then(function () { return self.clients.claim(); })
  );
});

self.addEventListener('fetch', function (event) {
  var req = event.request;
  var url = req.url;

  // Never intercept anything that isn't a simple same-origin GET for our own
  // shell files. This excludes Microsoft login/Graph/SharePoint calls, the
  // Excel data fetches, and any POST/OAuth redirect traffic automatically.
  if (req.method !== 'GET') return;
  if (url.indexOf(self.location.origin) !== 0) return;
  if (url.indexOf('login.microsoftonline.com') !== -1) return;
  if (url.indexOf('graph.microsoft.com') !== -1) return;
  if (url.indexOf('sharepoint.com') !== -1) return;

  var rel = relPath(url);
  if (rel === null) return;
  var isPage = req.mode === 'navigate' || rel === '' || rel === 'index.html';
  if (!isPage && SHELL_REL.indexOf(rel) === -1) return;

  if (isPage) {
    // NETWORK-FIRST for the dashboard page. The OAuth redirect lands on the app
    // URL with ?code=… — that must reach the page untouched, so we always fetch
    // the real request, and only store the clean app URL in the cache.
    event.respondWith(
      fetch(url, { cache: 'no-store', credentials: 'same-origin' }).then(function (response) {
        if (response && response.status === 200) {
          var copy = response.clone();
          caches.open(CACHE_NAME).then(function (cache) {
            cache.put('./index.html', copy);
          });
        }
        return response;
      }).catch(function () {
        // Offline: serve the last good copy of the app.
        return caches.match('./index.html').then(function (c) {
          return c || caches.match('./');
        });
      })
    );
    return;
  }

  // Icons / manifest: cache first, refresh in the background.
  event.respondWith(
    caches.match(req).then(function (cached) {
      var networkFetch = fetch(req).then(function (response) {
        if (response && response.status === 200) {
          var copy = response.clone();
          caches.open(CACHE_NAME).then(function (cache) {
            cache.put(req, copy);
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
