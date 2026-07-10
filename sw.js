/* 中古品在庫管理 ― サービスワーカー（常に最新を優先／オフライン時のみキャッシュ） */
const CACHE = 'chuko-navi-v2';
const SHELL = ['./', './index.html', './manifest.json', './icon-192.png', './icon-512.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)));
  self.skipWaiting();
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))));
  self.clients.claim();
});
self.addEventListener('fetch', e => {
  if (e.request.method !== 'GET') return;                 // 書き込み(POST)はそのままネットへ
  const url = new URL(e.request.url);
  if (url.origin !== location.origin) return;             // GAS等の外部APIはそのままネットへ
  // ネット優先：常に最新を取得。取れた分はキャッシュ更新。オフライン時のみキャッシュを使う。
  e.respondWith(
    fetch(e.request).then(res => {
      const copy = res.clone();
      caches.open(CACHE).then(c => c.put(e.request, copy));
      return res;
    }).catch(() => caches.match(e.request))
  );
});
