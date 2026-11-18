// Service Worker for Ramesh Wedding DJ Pro Console (Network-First Strategy)
const CACHE_NAME = 'wedding-dj-cache-v3';
const ASSETS_TO_CACHE = [
  './',
  './index.html',
  './manifest.json',
  './qrcode.min.js'
];

self.addEventListener('install', (e) => {
  self.skipWaiting();
  e.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      return cache.addAll(ASSETS_TO_CACHE);
    })
  );
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
  const url = new URL(e.request.url);
  // Never intercept API calls, media files, or byte ranges
  if (url.pathname.startsWith('/api/') || e.request.headers.has('range') || e.request.url.includes('.webm') || e.request.url.includes('.m4a')) {
    return;
  }

  // Network-First: Always fetch the latest version from GitHub/Server, fallback to cache if offline
  e.respondWith(
    fetch(e.request)
      .then((response) => {
        if (response && response.status === 200 && response.type === 'basic') {
          const resClone = response.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put(e.request, resClone));
        }
        return response;
      })
      .catch(() => {
        return caches.match(e.request);
      })
  );
});
