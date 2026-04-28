const CACHE_NAME = 'erp-cache-v1';
const ASSETS = [
  '/d.html',
  '/manifest.json'
];

self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME).then(cache => cache.addAll(ASSETS))
  );
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys => Promise.all(
      keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k))
    ))
  );
});

self.addEventListener('fetch', event => {
  // Only cache GET requests
  if(event.request.method !== 'GET') return;
  // Only cache same-origin requests
  if(!event.request.url.startsWith(self.location.origin)) return;
  // Don't cache /api calls
  if(event.request.url.includes('/api/')) return;
  
  event.respondWith(
    caches.match(event.request).then(response => {
      return response || fetch(event.request).then(fetchRes => {
        if (!fetchRes || fetchRes.status !== 200 || fetchRes.type !== 'basic') return fetchRes;
        const responseToCache = fetchRes.clone();
        caches.open(CACHE_NAME).then(cache => cache.put(event.request, responseToCache));
        return fetchRes;
      });
    }).catch((err) => {
      // Fallback for HTML pages when offline
      const accept = event.request.headers.get('accept');
      if (accept && accept.includes('text/html')) {
        return caches.match('/d.html');
      }
      throw err;
    })
  );
});
