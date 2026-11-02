// Fridgey service worker: lets Android offer "Install app". It does not cache anything;
// every request goes straight to the network so the app is always the latest version.
self.addEventListener('install', () => self.skipWaiting());
self.addEventListener('activate', e => e.waitUntil(self.clients.claim()));
self.addEventListener('fetch', () => {});
