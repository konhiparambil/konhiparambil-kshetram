const CACHE = 'kshetram-v1';
const ASSETS = [
  '/konhiparambil-kshetram/',
  '/konhiparambil-kshetram/index.html',
  '/konhiparambil-kshetram/icon-192.png',
  '/konhiparambil-kshetram/icon-512.png',
  'https://fonts.googleapis.com/css2?family=Noto+Serif+Malayalam:wght@400;600&family=Noto+Sans+Malayalam:wght@400;500&family=Crimson+Pro:ital,wght@0,400;0,600;1,400&family=DM+Sans:wght@400;500&display=swap'
];

self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(CACHE).then(cache => cache.addAll(ASSETS)).then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))
    ).then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', e => {
  // For Google Sheets API calls — always go to network
  if (e.request.url.includes('script.google.com') || e.request.url.includes('open-meteo') || e.request.url.includes('sunrise-sunset')) {
    e.respondWith(fetch(e.request).catch(() => new Response('{}', {headers: {'Content-Type': 'application/json'}})));
    return;
  }
  // For everything else — cache first, network fallback
  e.respondWith(
    caches.match(e.request).then(cached => {
      if (cached) return cached;
      return fetch(e.request).then(response => {
        const clone = response.clone();
        caches.open(CACHE).then(cache => cache.put(e.request, clone));
        return response;
      });
    })
  );
});
