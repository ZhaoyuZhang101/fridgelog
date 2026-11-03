// Fridgey service worker. It shows phone notifications (through OneSignal) and lets Android offer "Install app".
// It does not cache anything, so the app is always the latest version.
try { importScripts('https://cdn.onesignal.com/sdks/web/v16/OneSignalSDK.sw.js'); } catch (e) { /* offline: notifications load next time */ }
self.addEventListener('install', () => self.skipWaiting());
self.addEventListener('activate', e => e.waitUntil(self.clients.claim()));
self.addEventListener('fetch', () => {});
