const CACHE = 'liftlog-v10';
const SHELL = ['./', 'index.html', 'manifest.webmanifest', 'icon.svg'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
// Network first with cache fallback, so updates show up on the next open and the app still works offline.
self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);
  if (e.request.method !== 'GET' || url.origin !== location.origin) return;
  e.respondWith(
    fetch(e.request).then(res => {
      if (res.ok) { const copy = res.clone(); caches.open(CACHE).then(c => c.put(e.request, copy)); }
      return res;
    }).catch(() => caches.match(e.request))
  );
});

// Workout notification actions: forward to the open app, or open it.
self.addEventListener('notificationclick', e => {
  const a = e.action;
  if (e.notification.tag === 'liftlog-buzz') { e.notification.close(); return; }
  e.waitUntil(self.clients.matchAll({ type: 'window', includeUncontrolled: true }).then(cs => {
    if (a === 'pause' || a === 'skip' || a === 'add30') { cs.forEach(c => c.postMessage({ type: 'wk', act: a })); return; }
    e.notification.close();
    if (cs.length) return cs[0].focus();
    return self.clients.openWindow('./');
  }));
});
