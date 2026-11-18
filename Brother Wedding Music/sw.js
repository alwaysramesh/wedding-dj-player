// Service Worker for Ramesh Wedding DJ Pro Console
const CACHE_NAME = 'wedding-dj-cache-v1';
const ASSETS_TO_CACHE = [
  './',
  './index.html',
  './manifest.json'
];

self.addEventListener('install', (e) => {
  e.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      return cache.addAll(ASSETS_TO_CACHE);
    })
  );
  self.skipWaiting();
});

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys().then((keys) => {
      return Promise.all(
        keys.map((k) => {
          if (k !== CACHE_NAME) return caches.delete(k);
        })
      );
    })
  );
  self.clients.claim();
});

self.addEventListener('fetch', (e) => {
  // Never intercept API calls or media streaming range requests
  const url = new URL(e.request.url);
  if (url.pathname.startsWith('/api/') || e.request.headers.has('range') || e.request.url.includes('.webm') || e.request.url.includes('.m4a')) {
    return;
  }
  e.respondWith(
    caches.match(e.request).then((res) => {
      return res || fetch(e.request);
    })
  );
});
