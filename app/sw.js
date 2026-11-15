// Offline: appka se uloží do cache. Nová verze čeká na potvrzení v appce.
const VERSION = '1.17.0';
// ostrá a testovací verze běží na stejném webu, mají proto oddělenou cache
const PREFIX = self.location.pathname.indexOf('/test/') >= 0 ? 'fotoapp-test-v' : 'fotoapp-v';
const CACHE = PREFIX + VERSION;
const SHELL = ['./', 'index.html', 'manifest.webmanifest', 'icon-96.png', 'icon-192.png', 'icon-512.png', 'apple-touch-icon.png'];
self.addEventListener('install', e => e.waitUntil(
  caches.open(CACHE).then(c => c.addAll(SHELL.map(u => new Request(u, { cache: 'reload' }))))));
self.addEventListener('message', e => { if (e.data === 'skip') self.skipWaiting(); });
self.addEventListener('activate', e => e.waitUntil(
  caches.keys().then(ks => Promise.all(ks.filter(k => (k.indexOf(PREFIX) === 0 || (PREFIX === 'fotoapp-v' && k.indexOf('meridla-v') === 0)) && k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim())));
self.addEventListener('fetch', e => {
  const u = new URL(e.request.url);
  if (e.request.method !== 'GET' || u.origin !== location.origin) return; // volání na Google nikdy z cache
  e.respondWith(caches.match(e.request).then(hit => hit || fetch(e.request)));
});
