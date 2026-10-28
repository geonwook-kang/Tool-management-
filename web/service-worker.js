/* Only the immutable public application shell is cached. User files and
   authentication/API responses always remain outside this cache. */
'use strict';
const APP_VERSION = '__PWA_BUILD__';
const APP_ASSETS = /*__PWA_ASSETS__*/[];
const CACHE_PREFIX = 'tool-manager-shell:' + new URL(self.registration.scope).pathname + ':';
const CACHE_NAME = CACHE_PREFIX + APP_VERSION;
const assetURLs = APP_ASSETS.map(asset => new URL(asset.path, self.registration.scope).href);
const allowedURLs = new Set(assetURLs);
const indexURL = new URL('./index.html', self.registration.scope).href;
const scopeURL = new URL(self.registration.scope);
const hex = bytes => Array.from(new Uint8Array(bytes), value => value.toString(16).padStart(2, '0')).join('');

self.addEventListener('install', event => {
  event.waitUntil((async () => {
    const cache = await caches.open(CACHE_NAME);
    try {
      await Promise.all(APP_ASSETS.map(async asset => {
        const url = new URL(asset.path, self.registration.scope);
        const response = await fetch(new Request(url.href, { cache: 'reload', credentials: 'omit', redirect: 'error' }));
        if (!response.ok || response.type === 'opaque') throw new Error('Application asset unavailable: ' + asset.path);
        const bytes = await response.clone().arrayBuffer();
        if (bytes.byteLength !== asset.bytes || hex(await crypto.subtle.digest('SHA-256', bytes)) !== asset.sha256) {
          throw new Error('Application asset verification failed: ' + asset.path);
        }
        await cache.put(url.href, response);
      }));
      // Do not skipWaiting here. A new shell cannot interrupt OCR or a save.
    } catch (error) {
      await caches.delete(CACHE_NAME);
      throw error;
    }
  })());
});

self.addEventListener('activate', event => {
  event.waitUntil((async () => {
    for (const name of await caches.keys()) if (name.startsWith(CACHE_PREFIX) && name !== CACHE_NAME) await caches.delete(name);
    await self.clients.claim();
  })());
});

self.addEventListener('fetch', event => {
  const request = event.request;
  if (request.method !== 'GET') return;
  const url = new URL(request.url);
  if (url.origin !== scopeURL.origin) return;
  // Navigations to the app's exact entry points can use the shell. Query
  // strings (including sign-in callbacks) are never stored as cache keys.
  const entry = request.mode === 'navigate' && (url.pathname === scopeURL.pathname || url.pathname === new URL(indexURL).pathname);
  const key = entry ? indexURL : request.url;
  if (!entry && !allowedURLs.has(key)) return;
  event.respondWith((async () => {
    const cached = await (await caches.open(CACHE_NAME)).match(key);
    return cached || fetch(request);
  })());
});

function askClient(client) {
  return new Promise(resolve => {
    const channel = new MessageChannel();
    const timer = setTimeout(() => { channel.port1.close(); resolve(false); }, 2500);
    channel.port1.onmessage = event => {
      clearTimeout(timer); channel.port1.close(); resolve(event.data && event.data.idle === true);
    };
    try { client.postMessage({ type: 'TOOL_PWA_CAN_ACTIVATE' }, [channel.port2]); }
    catch (_) { clearTimeout(timer); channel.port1.close(); resolve(false); }
  });
}
let activationTask = null;
self.addEventListener('message', event => {
  const data = event.data || {}, port = event.ports && event.ports[0];
  if (data.type === 'TOOL_PWA_STATUS') {
    event.waitUntil((async () => {
      const cache = await caches.open(CACHE_NAME);
      const complete = (await Promise.all(assetURLs.map(url => cache.match(url)))).every(Boolean);
      if (port) port.postMessage({ version: APP_VERSION, complete });
    })());
  } else if (data.type === 'TOOL_PWA_ACTIVATE') {
    if (!activationTask) activationTask = (async () => {
      const windows = (await self.clients.matchAll({ type: 'window', includeUncontrolled: true }))
        .filter(client => client.url.startsWith(self.registration.scope));
      const idle = windows.length > 0 && (await Promise.all(windows.map(askClient))).every(Boolean);
      if (!idle) return { accepted: false, reason: 'busy' };
      await self.skipWaiting(); return { accepted: true };
    })().finally(() => { activationTask = null; });
    event.waitUntil(activationTask.then(result => { if (port) port.postMessage(result); }));
  }
});
