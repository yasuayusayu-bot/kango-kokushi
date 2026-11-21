const C="kokushi-v7-2026-11-21b";
const FILES=["./","index.html","config.js","manifest.json","icon-180.png","icon-192.png","icon-512.png"];
self.addEventListener("install",e=>{e.waitUntil(caches.open(C).then(c=>c.addAll(FILES)));self.skipWaiting();});
self.addEventListener("activate",e=>{e.waitUntil(caches.keys().then(ks=>Promise.all(ks.filter(k=>k!==C).map(k=>caches.delete(k)))));self.clients.claim();});
self.addEventListener("fetch",e=>{
  const u=new URL(e.request.url); if(e.request.method!=="GET"||u.origin!==location.origin) return;
  // ページ本体は最新を優先（オフライン時はキャッシュ）
  if(e.request.mode==="navigate"||/(index|admin)\.html$|config\.js$/.test(u.pathname)){
    const key=u.pathname.endsWith("config.js")?"config.js":u.pathname.endsWith("admin.html")?"admin.html":"index.html";
    e.respondWith(fetch(e.request).then(r=>{const cp=r.clone();caches.open(C).then(c=>c.put(key,cp));return r;}).catch(()=>caches.match(key)));return;}
  e.respondWith(caches.match(e.request).then(r=>r||fetch(e.request)));
});
