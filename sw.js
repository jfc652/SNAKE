/* Service worker du Générateur Snake Excel
   - Incrémentez VERSION à chaque mise à jour pour forcer le rafraîchissement du cache. */
const VERSION = "snake-v3";
const CACHE = VERSION;

const FICHIERS_LOCAUX = [
  "./",
  "./index.html",
  "./app.js",
  "./modele-b64.js",
  "./manifest.json",
  "./icon-192.png",
  "./icon-512.png",
];

const LIBS_CDN = [
  "https://cdn.jsdelivr.net/npm/papaparse@5.4.1/papaparse.min.js",
  "https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js",
  "https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js",
];

self.addEventListener("install", event => {
  event.waitUntil((async () => {
    const cache = await caches.open(CACHE);
    // Fichiers locaux : on n'échoue pas si l'un manque
    await Promise.all(FICHIERS_LOCAUX.map(u => cache.add(u).catch(() => {})));
    // Librairies externes (réponses opaques acceptées)
    await Promise.all(LIBS_CDN.map(u =>
      fetch(u, { mode: "no-cors" }).then(r => cache.put(u, r)).catch(() => {})
    ));
    await self.skipWaiting();
  })());
});

self.addEventListener("activate", event => {
  event.waitUntil((async () => {
    const noms = await caches.keys();
    await Promise.all(noms.filter(n => n !== CACHE).map(n => caches.delete(n)));
    await self.clients.claim();
  })());
});

self.addEventListener("fetch", event => {
  const req = event.request;
  if (req.method !== "GET") return;

  const url = new URL(req.url);
  const memeOrigine = url.origin === self.location.origin;

  // Fichiers de l'appli : réseau d'abord (mises à jour immédiates), cache si hors-ligne
  if (memeOrigine) {
    event.respondWith((async () => {
      try {
        const rep = await fetch(req);
        if (rep && rep.ok) {
          const cache = await caches.open(CACHE);
          cache.put(req, rep.clone());
        }
        return rep;
      } catch (e) {
        const enCache = await caches.match(req, { ignoreSearch: true });
        if (enCache) return enCache;
        if (req.mode === "navigate") return caches.match("./index.html");
        throw e;
      }
    })());
    return;
  }

  // Librairies externes : cache d'abord
  event.respondWith((async () => {
    const enCache = await caches.match(req);
    if (enCache) return enCache;
    try {
      const rep = await fetch(req);
      const cache = await caches.open(CACHE);
      cache.put(req, rep.clone());
      return rep;
    } catch (e) {
      return new Response("", { status: 504 });
    }
  })());
});
