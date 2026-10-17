// Offline cache so the app opens without internet after the first visit.
const CACHE='pjp-suraj-v3';
const FILES=['./','index.html','local-backend.js','vendor/xlsx.min.js','manifest.webmanifest','icon.svg','icon-192.png','icon-512.png'];

self.addEventListener('install',e=>{
  e.waitUntil(caches.open(CACHE).then(c=>c.addAll(FILES)).then(()=>self.skipWaiting()));
});
self.addEventListener('activate',e=>{
  e.waitUntil(caches.keys().then(keys=>Promise.all(keys.filter(k=>k!==CACHE).map(k=>caches.delete(k)))).then(()=>self.clients.claim()));
});
// Network first, cache fallback: updates show up when online, still works offline.
self.addEventListener('fetch',e=>{
  if(e.request.method!=='GET') return;
  e.respondWith(
    // no-cache: always revalidate with the server so a new deploy shows up on the next open.
    fetch(e.request,{cache:'no-cache'}).then(res=>{
      const copy=res.clone();
      caches.open(CACHE).then(c=>c.put(e.request,copy));
      return res;
    }).catch(()=>caches.match(e.request,{ignoreSearch:true}))
  );
});
