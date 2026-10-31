/* Only the immutable public application shell is cached. User files and
   authentication/API responses always remain outside this cache. */
'use strict';
const APP_VERSION = 'b54f25967775e9bb569b734c';
const APP_ASSETS = [{"path":"./icons/apple-touch-icon.png","sha256":"e6c7075c557e9fab37ec1cf0e3dd0216120673cc643eb5a202a1e506af22344d","bytes":655},{"path":"./icons/icon-192.png","sha256":"ecc591e93d232b07c23403556f2efdfe5ebeb1620b0273bc08bcc6726628ced0","bytes":690},{"path":"./icons/icon-512.png","sha256":"90fdf79867e54a48e188431ef9312e2fb3c0b8b422d6a539c535aa4888249b39","bytes":1923},{"path":"./icons/icon-maskable-512.png","sha256":"b9cac0d3c09c130a3fb8c18cee9a8bc6521a3a85e931566b651a91c705c20b66","bytes":1898},{"path":"./index.html","sha256":"51ab1a6cc3103c3b54b4a6315bb834d8c8938097b36852d89978a0b5fbb56245","bytes":10753069},{"path":"./manifest.webmanifest","sha256":"8be0450f7cf5cedccf142cb360b712a452ef7d02cd4aa601dd9182ba9ff7843b","bytes":711},{"path":"./pwa.js","sha256":"43d35eb8a7bc3269b2fbd6abab4cd9bc61f93c60d14daa65a313465000f51310","bytes":8513}];
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
