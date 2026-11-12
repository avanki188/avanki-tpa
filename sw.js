/* Avanki TPA – lưu sẵn app để mở được khi mất mạng. Ưu tiên bản mới trên mạng, mất mạng thì dùng bản đã lưu.
   Chỉ lưu file của app (index.html, manifest, icon); dữ liệu hồ sơ nằm trên OneDrive, không đi qua đây. */
const CACHE = 'avanki-tpa-v1';
const SHELL = ['./', './index.html', './manifest.webmanifest', './icon-192.png', './icon-512.png'];
self.addEventListener('install', e => { e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).then(() => self.skipWaiting())); });
self.addEventListener('activate', e => { e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim())); });
self.addEventListener('fetch', e => {
  const u = new URL(e.request.url);
  if (e.request.method !== 'GET' || u.origin !== location.origin) return;
  if (u.pathname.endsWith('/version.json')) return; /* luôn hỏi mạng */
  e.respondWith(fetch(e.request, { cache: 'no-store' }).then(r => { if (r.ok) { const c = r.clone(); caches.open(CACHE).then(x => x.put(e.request, c)); } return r; }).catch(() => caches.match(e.request).then(r => r || caches.match('./index.html'))));
});
