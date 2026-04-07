const CACHE_NAME = 'keio-navi-static-v2';
const CACHEABLE_DESTINATIONS = new Set(['style', 'script', 'image', 'font']);

function isKeioNaviCacheName(name) {
  return typeof name === 'string' && name.indexOf('keio-navi-') === 0;
}

function isCacheableRequest(request) {
  if (!request || request.method !== 'GET' || request.mode === 'navigate') return false;

  const url = new URL(request.url);
  if (url.origin !== self.location.origin) return false;
  if (!CACHEABLE_DESTINATIONS.has(request.destination || '')) return false;
  if (/\.html?$/i.test(url.pathname)) return false;

  return true;
}

function sanitizeTargetUrl(rawUrl) {
  const fallback = 'members.html';
  const raw = String(rawUrl || '').trim();
  if (!raw) return fallback;
  if (/^[a-zA-Z][a-zA-Z0-9+.-]*:/.test(raw) || raw.startsWith('//')) return fallback;

  try {
    const url = new URL(raw, self.location.origin + '/');
    if (url.origin !== self.location.origin) return fallback;
    const path = url.pathname || '/';
    if (path !== '/' && !path.endsWith('.html')) return fallback;
    return (path === '/' ? '/members.html' : path).replace(/^\/+/, '') + url.search + url.hash;
  } catch (error) {
    return fallback;
  }
}

async function clearKeioNaviCaches() {
  const keys = await caches.keys();
  await Promise.all(
    keys
      .filter(isKeioNaviCacheName)
      .map(key => caches.delete(key))
  );
}

self.addEventListener('install', () => {
  self.skipWaiting();
});

self.addEventListener('activate', event => {
  event.waitUntil((async () => {
    const keys = await caches.keys();
    await Promise.all(
      keys
        .filter(key => isKeioNaviCacheName(key) && key !== CACHE_NAME)
        .map(key => caches.delete(key))
    );
    await self.clients.claim();
  })());
});

self.addEventListener('message', event => {
  if (!event.data || event.data.type !== 'CLEAR_RUNTIME_CACHE') return;

  event.waitUntil((async () => {
    await clearKeioNaviCaches();
    if (self.registration.getNotifications) {
      const notifications = await self.registration.getNotifications();
      notifications.forEach(notification => notification.close());
    }
  })());
});

self.addEventListener('fetch', event => {
  if (!isCacheableRequest(event.request)) return;

  event.respondWith((async () => {
    const cache = await caches.open(CACHE_NAME);
    const cached = await cache.match(event.request);

    try {
      const response = await fetch(event.request);
      if (response && response.ok && response.type === 'basic') {
        cache.put(event.request, response.clone());
      }
      return response;
    } catch (error) {
      if (cached) return cached;
      throw error;
    }
  })());
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
    data: { url: sanitizeTargetUrl(data.url || 'members.html') },
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

  const targetUrl = new URL(
    sanitizeTargetUrl((event.notification.data && event.notification.data.url) || 'members.html'),
    self.location.origin + '/'
  ).href;

  event.waitUntil(
    clients.matchAll({ type: 'window', includeUncontrolled: true }).then(windowClients => {
      for (const client of windowClients) {
        try {
          const clientUrl = new URL(client.url);
          const nextUrl = new URL(targetUrl);
          if ((clientUrl.href === nextUrl.href || clientUrl.pathname === nextUrl.pathname) && 'focus' in client) {
            return client.focus();
          }
        } catch (error) {
          // noop
        }
      }
      if (clients.openWindow) {
        return clients.openWindow(targetUrl);
      }
    })
  );
});
