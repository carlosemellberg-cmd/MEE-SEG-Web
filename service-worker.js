const CACHE = "mee-seg-web-v1.0.2";
const ASSETS = ["./", "./index.html", "./styles.css", "./config.js", "./data-adapter.js", "./mee-bridge.js", "./file-sync.js", "./app.js", "./manifest.webmanifest", "./assets/icon-192.png", "./assets/icon-512.png", "./assets/logo-loma-negra.png"];
self.addEventListener("install", event => event.waitUntil(caches.open(CACHE).then(cache => cache.addAll(ASSETS)).then(() => self.skipWaiting())));
self.addEventListener("activate", event => event.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(key => key !== CACHE).map(key => caches.delete(key)))).then(() => self.clients.claim())));
self.addEventListener("fetch", event => {
  if (event.request.method !== "GET") return;
  event.respondWith(caches.match(event.request).then(cached => cached || fetch(event.request).then(response => {
    const copy = response.clone(); caches.open(CACHE).then(cache => cache.put(event.request, copy)); return response;
  }).catch(() => caches.match("./index.html"))));
});
