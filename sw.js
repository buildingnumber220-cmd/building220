// يخلّي الصفحات تفتح حتى لو النت فاصل (آخر نسخة اتفتحت). لا يلمس طلبات الـ API ولا يحفظ أي بيانات.
const V = 'bf-v1';
self.addEventListener('install', e => { self.skipWaiting(); });
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== V).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
// مفتاح الكاش بدون ?query عشان الرقم السري في الرابط (لو موجود) ما يتخزنش
const keyOf = u => new Request(u.origin + u.pathname);
self.addEventListener('fetch', e => {
  const r = e.request;
  if (r.method !== 'GET') return;
  const u = new URL(r.url);
  if (u.origin !== location.origin) return;
  const key = keyOf(u);
  e.respondWith(
    fetch(r).then(res => {
      if (res.ok && res.type === 'basic') { const cp = res.clone(); caches.open(V).then(c => c.put(key, cp)); }
      return res;
    }).catch(() => caches.match(key).then(m => m || (u.pathname.endsWith('/') ? caches.match(new Request(u.origin + u.pathname + 'index.html')) : undefined)))
  );
});
