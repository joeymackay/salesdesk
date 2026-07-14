/* SalesDesk service worker — caches the app shell so it opens offline.
   Never touches GitHub API traffic. */
const CACHE = "salesdesk-gh-v1";
const SHELL = ["./","./index.html","./app.js","./react.js","./react-dom.js","./manifest.webmanifest","./icon-192.png","./icon-512.png"];

self.addEventListener("install", (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(SHELL)).then(() => self.skipWaiting()));
});
self.addEventListener("activate", (e) => {
  e.waitUntil(caches.keys()
    .then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
    .then(() => self.clients.claim()));
});
self.addEventListener("fetch", (e) => {
  const url = new URL(e.request.url);
  // Data + auth traffic always goes straight to the network.
  if (url.hostname === "api.github.com" || url.hostname.endsWith("github.com")) return;
  if (e.request.method !== "GET") return;
  if (url.origin !== self.location.origin) return;
  e.respondWith(
    caches.match(e.request).then((hit) => {
      const net = fetch(e.request).then((res) => {
        if (res && res.ok) { const copy = res.clone(); caches.open(CACHE).then((c) => c.put(e.request, copy)); }
        return res;
      }).catch(() => hit);
      return hit || net;
    })
  );
});
