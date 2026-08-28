// bridge.js —— 主线程 ↔ Worker 消息协议封装
// 用法:
//   const engine = createEngine();
//   await engine.init(onProgress);                    // 首次加载引擎
//   engine.convert(fileBytes, {password, sheet}, cb)  // cb: {onProgress,onDone,onError}
//   engine.diagnose(fileBytes, filename, cb)

export function createEngine() {
  // 页面经 ES module 加载 app.js → 这里同样用 module Worker。
  // worker.js 用 importScripts 加载 Pyodide(经典范式), 因此 worker 本身保持 classic 类型。
  const worker = new Worker(new URL("./worker.js", import.meta.url));
  let nextId = 1;
  const pending = new Map();   // id → {onProgress,onDone,onError}
  let readyResolve = null;
  let readyReject = null;

  worker.onmessage = (e) => {
    const { type, id } = e.data || {};
    if (type === "ready") { if (readyResolve) readyResolve(); return; }
    if (type === "error" && id === undefined) { if (readyReject) readyReject(new Error(e.data.message)); return; }
    const p = pending.get(id);
    if (!p) return;
    if (type === "progress") { p.onProgress && p.onProgress(e.data.page, e.data.total); return; }
    if (type === "done") { pending.delete(id); p.onDone && p.onDone(e.data); return; }
    if (type === "error") {
      pending.delete(id);
      const err = new Error(e.data.message);
      err.detail = e.data.detail;  // 完整 Python traceback(worker 附加)
      p.onError && p.onError(err);
      return;
    }
    if (type === "diag") { pending.delete(id); p.onDone && p.onDone(e.data.json); return; }
  };
  worker.onerror = (e) => {
    // Worker 级错误(脚本加载失败等): 拒绝 ready 与全部 pending
    const msg = e.message || "worker 加载失败";
    if (readyReject) readyReject(new Error(msg));
    for (const [, p] of pending) p.onError && p.onError(new Error(msg));
    pending.clear();
  };

  function init() {
    return new Promise((resolve, reject) => {
      readyResolve = resolve; readyReject = reject;
      worker.postMessage({ type: "init" });
    });
  }

  // 输入归一化: File/Blob → Uint8Array(ArrayBuffer 视图), 其余原样
  async function toBytes(input) {
    if (input instanceof Uint8Array) return input;
    if (ArrayBuffer.isView(input)) return new Uint8Array(input.buffer, input.byteOffset, input.byteLength);
    if (input instanceof ArrayBuffer) return new Uint8Array(input);
    if (typeof Blob !== "undefined" && input instanceof Blob) {
      return new Uint8Array(await input.arrayBuffer());
    }
    throw new Error("不支持的输入类型: " + (input && input.constructor && input.constructor.name));
  }

  async function convert(pdfBytes, { password = "", sheet = "对账单" } = {}, cb = {}) {
    const id = nextId++;
    pending.set(id, cb);
    const u8 = await toBytes(pdfBytes);
    worker.postMessage({ type: "convert", id, pdfBytes: u8, password, sheet }, [u8.buffer]);
    return id;
  }

  async function diagnose(pdfBytes, filename, cb = {}) {
    const id = nextId++;
    pending.set(id, cb);
    const u8 = await toBytes(pdfBytes);
    worker.postMessage({ type: "diagnose", id, pdfBytes: u8, filename }, [u8.buffer]);
    return id;
  }

  return { init, convert, diagnose };
}
