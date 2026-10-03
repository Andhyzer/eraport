const CACHE = 'erapor-v1.0.0';
const ASSETS = [
  './',
  './index.html',
  './verify.html',
  './manifest.json',
  './assets/style.css',
  './assets/app.js',
  './assets/cache.js',
  './assets/notif.js',
  './assets/qrcode-sign.js',
  './assets/excel.js',
  './assets/multitenant.js',
  './assets/backup.js',
  './assets/dapodik.js',
  './assets/gemini.js',
  './assets/template-upload.js',
  './assets/rapor-template.js',
  './assets/analytics.js',
  './assets/analytics-ui.js'
];

self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(CACHE).then(c => c.addAll(ASSETS).catch(err => console.warn('SW asset cache:', err)))
      .then(() => self.skipWaiting())
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
  if (e.request.method !== 'GET') return;
  e.respondWith(
    caches.match(e.request).then(cached => {
      if (cached) return cached;
      return fetch(e.request).then(res => {
        if (res.ok && res.type === 'basic') {
          const clone = res.clone();
          caches.open(CACHE).then(c => c.put(e.request, clone));
        }
        return res;
      }).catch(() => caches.match('./index.html'));
    })
  );
});