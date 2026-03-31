const CACHE_NAME = 'keio-navi-v1';
const STATIC_ASSETS = [
  'members.html',
  'writing.html',
  'interview-page.html',
  'resources.html',
  'glossary.html',
  'community.html',
  'account.html',
  'auth.js'
];

self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME).then(cache => cache.addAll(STATIC_ASSETS))
  );
  self.skipWaiting();
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', event => {
  if (event.request.method !== 'GET') return;
  event.respondWith(
    fetch(event.request).then(response => {
      const clone = response.clone();
      caches.open(CACHE_NAME).then(cache => cache.put(event.request, clone));
      return response;
    }).catch(() => caches.match(event.request))
  );
});

// ── Push通知ハンドラ ──────────────────────────────
self.addEventListener('push', event => {
  let data = { title: '就活ナビ', body: '新しいお知らせがあります。' };
  try {
    if (event.data) {
      const parsed = event.data.json();
      data = Object.assign(data, parsed);
    }
  } catch (e) {
    if (event.data) {
      data.body = event.data.text();
    }
  }

  const options = {
    body: data.body || '',
    icon: data.icon || 'data:image/svg+xml,<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100"><rect width="100" height="100" rx="20" fill="%230a1a3e"/><text x="50" y="65" font-size="50" text-anchor="middle" fill="%23c9a84c" font-family="serif">就</text></svg>',
    badge: data.badge || undefined,
    tag: data.tag || 'keio-navi-push',
    data: { url: data.url || 'members.html' },
    actions: data.actions || [
      { action: 'open', title: '開く' },
      { action: 'close', title: '閉じる' }
    ],
    requireInteraction: false,
  };

  event.waitUntil(
    self.registration.showNotification(data.title, options)
  );
});

self.addEventListener('notificationclick', event => {
  event.notification.close();

  if (event.action === 'close') return;

  const targetUrl = (event.notification.data && event.notification.data.url) || 'members.html';

  event.waitUntil(
    clients.matchAll({ type: 'window', includeUncontrolled: true }).then(windowClients => {
      for (const client of windowClients) {
        if (client.url.includes(targetUrl) && 'focus' in client) {
          return client.focus();
        }
      }
      if (clients.openWindow) {
        return clients.openWindow(targetUrl);
      }
    })
  );
});
