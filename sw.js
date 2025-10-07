// Service worker for Evidence Collector PWA - simple precache + runtime cache (stale-while-revalidate)
const CACHE_NAME = 'evidencecam-v1';
const PRECACHE_ASSETS = [
  './',
  './evidence_pwa_v1_2_sharepoint.html',
  './manifest.webmanifest',
  './sw.js',
  'https://cdn.jsdelivr.net/npm/pptxgenjs@3.12.0/dist/pptxgen.bundle.js',
  'https://cdn.jsdelivr.net/npm/jspdf@2.5.1/dist/jspdf.umd.min.js',
  './icon-192.png',
  './icon-512.png',
  './icon-maskable.png'
];

self.addEventListener('install', event => {
  self.skipWaiting();
  event.waitUntil(
    caches.open(CACHE_NAME).then(cache => cache.addAll(PRECACHE_ASSETS))
  );
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(
        keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k))
      )
    ).then(() => self.clients.claim())
  );
});

// fetch: try cache first, then network; update cache in background
self.addEventListener('fetch', event => {
  const req = event.request;

  // Only handle GET requests
  if (req.method !== 'GET') return;

  // Navigation requests fallback to the app shell (HTML)
  if (req.mode === 'navigate') {
    event.respondWith(
      fetch(req).then(res => {
        const copy = res.clone();
        caches.open(CACHE_NAME).then(cache => cache.put(req, copy));
        return res;
      }).catch(() => caches.match('./evidence_pwa_v1_2_sharepoint.html'))
    );
    return;
  }

  // For other requests: stale-while-revalidate
  event.respondWith(
    caches.match(req).then(cached => {
      const networkFetch = fetch(req).then(networkRes => {
        // only cache successful responses
        if (networkRes && networkRes.status === 200 && networkRes.type !== 'opaque') {
          caches.open(CACHE_NAME).then(cache => cache.put(req, networkRes.clone()));
        }
        return networkRes;
      }).catch(() => null);
      return cached || networkFetch;
    })
  );
});