// sw.js —— bank2excel-h5 Service Worker
// 缓存策略:
//   cache-first        : wheel(17MB) + Pyodide CDN(固定版本) —— 最怕重复下载
//   stale-while-reval. : index/src/python 源码 —— 更新时能拿到新版, 离线可用
// 版本号: 更新部署时 bump CACHE 名, 自动清旧缓存

const CACHE = "bank2excel-h5-v3";  // bump: v2→v3 (forbid src/ caching to fix Pages deploy lag)

const CACHE_FIRST = [
  /^https:\/\/cdn\.jsdelivr\.net\/pyodide\//,           // Pyodide 运行时(固定 0.27.2)
  /\/wheels\//,                                         // 17MB wheel(同源)
  /^https:\/\/cdn\.jsdelivr\.net\/gh\/jxuzhi-lab\/bank2excel-h5@main\/wheels\//,
];

const NETWORK_FIRST = [
  /\/src\/worker\.js/,   // worker.js 必须拿最新(否则 micropip kwargs 等代码 bug 不会修)
];

const SWR = [
  /\/src\//,          // 应用 JS 模块(index/app/bridge 等)
  /\/python\//,       // 引擎源码
  /\/index\.html$/,
  /\/privacy\.html$/,
  /\/manifest\.json$/,
];

self.addEventListener("install", () => self.skipWaiting());
self.addEventListener("activate", (e) => {
  e.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
      // v3 一次性自我注销: 强制浏览器清掉所有旧 SW 控制权, 下次硬刷
      // 会以全新 SW 接管(避免旧 SW 缓存的旧 worker.js 反复触发 keep_going 报错)
      .then(() => {
        if (CACHE === "bank2excel-h5-v3") {
          // 仅本版本自我注销; 之后版本不再触发
          return self.registration.unregister();
        }
      })
  );
});

async function cacheFirst(req) {
  const cached = await caches.match(req);
  if (cached) return cached;
  try {
    const resp = await fetch(req);
    if (resp && (resp.status === 200 || resp.type === "opaque")) {
      const clone = resp.clone();
      caches.open(CACHE).then((c) => c.put(req, clone));
    }
    return resp;
  } catch (e) {
    return cached; // 断网且无缓存时返回 undefined → 浏览器兜底
  }
}

async function staleWhileRevalidate(req) {
  const cached = await caches.match(req);
  const network = fetch(req)
    .then((resp) => {
      if (resp && resp.status === 200) {
        const clone = resp.clone();
        caches.open(CACHE).then((c) => c.put(req, clone));
      }
      return resp;
    })
    .catch(() => null);
  if (cached) return cached;
  return network || fetch(req);
}

async function networkFirst(req) {
  // 必须拿最新: 拿不到时回退到缓存(用于离线场景); 都不行就交给浏览器
  try {
    const resp = await fetch(req);
    if (resp && resp.status === 200) {
      const clone = resp.clone();
      caches.open(CACHE).then((c) => c.put(req, clone));
    }
    return resp;
  } catch (e) {
    const cached = await caches.match(req);
    if (cached) return cached;
    throw e;
  }
}

self.addEventListener("fetch", (e) => {
  if (e.request.method !== "GET") return;
  const url = new URL(e.request.url);
  if (url.origin === self.location.origin && url.pathname === "/") return; // 根路径直通
  if (CACHE_FIRST.some((p) => p.test(url.href))) {
    e.respondWith(cacheFirst(e.request));
    return;
  }
  if (NETWORK_FIRST.some((p) => p.test(url.pathname))) {
    e.respondWith(networkFirst(e.request));
    return;
  }
  if (SWR.some((p) => p.test(url.pathname))) {
    e.respondWith(staleWhileRevalidate(e.request));
    return;
  }
  // 其余(根路径页面导航等)走网络
});
