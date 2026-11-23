/* 비서 앱 서비스워커 — 화면 파일을 폰에 저장해 두고 즉시 띄운 뒤, 뒤에서 새 버전을 받아 둡니다(다음 실행 때 반영).
 * 데이터 요청(script.google.com)은 건드리지 않습니다. */
const CACHE = 'biseo-shell-v9';
const SHELL = ['./', 'index.html', 'manifest.webmanifest', 'icon-192.png', 'icon-512.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  const same = url.origin === self.location.origin;
  const font = url.hostname === 'cdn.jsdelivr.net';
  if (!same && !font) return;
  e.respondWith(caches.open(CACHE).then(async c => {
    const key = same && req.mode === 'navigate' ? 'index.html' : req;
    const hit = await c.match(key);
    const net = fetch(req).then(res => { if (res && (res.ok || res.type === 'opaque')) c.put(key, res.clone()); return res; }).catch(() => null);
    if (hit) { e.waitUntil(net); return hit; }
    return (await net) || new Response('오프라인입니다', { status: 503, headers: { 'Content-Type': 'text/plain; charset=utf-8' } });
  }));
});
