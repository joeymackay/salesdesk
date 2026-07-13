/* SalesDesk service worker — caches the app shell so it opens offline.
   It deliberately NEVER touches Microsoft login or Graph API traffic. */

const CACHE = "salesdesk-v1";
const SHELL = [
  "./",
  "./index.html",
  "./app.js",
  "./react.js",
  "./react-dom.js",
  "./msal.js",
  "./manifest.webmanifest",
  "./icon-192.png",
  "./icon-512.png",
];

self.addEventListener("install", (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(SHELL)).then(() => self.skipWaiting()));
});

self.addEventListener("activate", (e) => {
  e.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", (e) => {
  const url = new URL(e.request.url);

  // Never cache or intercept auth / data traffic — always go straight to the network.
  if (
    url.hostname.endsWith("login.microsoftonline.com") ||
    url.hostname.endsWith("graph.microsoft.com") ||
    url.hostname.endsWith("login.live.com")
  ) {
    return; // browser handles it normally
  }

  if (e.request.method !== "GET") return;
  if (url.origin !== self.location.origin) return;

  // App shell: serve from cache, refresh in the background.
  e.respondWith(
    caches.match(e.request).then((hit) => {
      const net = fetch(e.request)
        .then((res) => {
          if (res && res.ok) {
            const copy = res.clone();
            caches.open(CACHE).then((c) => c.put(e.request, copy));
          }
          return res;
        })
        .catch(() => hit);
      return hit || net;
    })
  );
});
