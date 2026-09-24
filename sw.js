const SHELL_CACHE='voice-finance-shell-v22';
const SHELL_FILES=['./','./index.html','./v22.html','./manifest.webmanifest'];

self.addEventListener('install',event=>{
  event.waitUntil(
    caches.open(SHELL_CACHE)
      .then(cache=>cache.addAll(SHELL_FILES))
      .then(()=>self.skipWaiting())
  );
});

self.addEventListener('activate',event=>{
  event.waitUntil(
    caches.keys()
      .then(keys=>Promise.all(keys.filter(k=>k!==SHELL_CACHE).map(k=>caches.delete(k))))
      .then(()=>self.clients.claim())
  );
});

self.addEventListener('fetch',event=>{
  const req=event.request;
  if(req.method!=='GET')return;

  const url=new URL(req.url);
  if(url.origin!==self.location.origin)return;

  event.respondWith(
    caches.match(req).then(cached=>{
      const fresh=fetch(req).then(resp=>{
        if(resp&&resp.ok){
          const copy=resp.clone();
          caches.open(SHELL_CACHE).then(cache=>cache.put(req,copy));
        }
        return resp;
      }).catch(()=>cached);

      return cached||fresh;
    })
  );
});
