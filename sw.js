const CACHE = "jt2-v2";
const FILES = ["./", "./index.html", "./manifest.json"];
// 設定系は常に最新を取りに行く（変更が全端末にすぐ届くように）
const ALWAYS_FRESH = ["config.js", "voices.json"];

self.addEventListener("install", e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(FILES)).then(() => self.skipWaiting()));
});
self.addEventListener("activate", e => {
  e.waitUntil(caches.keys().then(ks =>
    Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))
  ).then(() => self.clients.claim()));
});
self.addEventListener("fetch", e => {
  const url = new URL(e.request.url);
  if (e.request.method !== "GET" || url.hostname.indexOf("google") >= 0) return;

  const fresh = ALWAYS_FRESH.some(f => url.pathname.endsWith("/" + f));
  if (fresh) {
    e.respondWith(
      fetch(new Request(url.href, { cache: "no-store" }))
        .then(res => { const c = res.clone(); caches.open(CACHE).then(x => x.put(e.request, c)).catch(() => {}); return res; })
        .catch(() => caches.match(e.request))
    );
    return;
  }
  e.respondWith(
    fetch(e.request).then(res => {
      const copy = res.clone();
      caches.open(CACHE).then(c => c.put(e.request, copy)).catch(() => {});
      return res;
    }).catch(() => caches.match(e.request).then(r => r || caches.match("./index.html")))
  );
});
