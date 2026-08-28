// worker.js —— bank2excel-h5 转换引擎宿主(Web Worker)
// 职责: Pyodide 运行时加载 + 引擎/shim 注入(FS+import 范式) + convert_bytes 调用
// 协议(与主线程 bridge 对应):
//   收 {type:"init"}                     → 加载引擎, 回 {type:"ready"} / {type:"error"}
//   收 {type:"convert", id, pdfBytes, password, sheet}
//       → 回 {type:"progress", id, page, total}(每页)
//        → 回 {type:"done", id, xlsx(ArrayBuffer), rows, report, stats, pages, sheet, warnings}
//        或  {type:"error", id, message, stage}
//   收 {type:"diagnose", id, pdfBytes, filename}
//       → 回 {type:"diag", id, json}

let py = null;
let initPromise = null;

// ---- Pyodide 资源 URL(相对 worker 脚本位置解析) ----
const PYODIDE_VER = "0.27.2";
const PYODIDE_URL = `https://cdn.jsdelivr.net/pyodide/v${PYODIDE_VER}/full/pyodide.js`;
const WHEEL_URL = "../wheels/pymupdf-1.26.7-cp312-abi3-pyodide_2024_0_wasm32.whl";
const ENGINE_URL = "../python/extract_bank_statement.py";
const SHIM_URL = "../python/shim.py";

async function fetchPySource(url, label) {
  const r = await fetch(url);
  if (!r.ok) throw new Error(`${label} HTTP ${r.status}`);
  const t = await r.text();
  if (t.trimStart().startsWith("<")) {
    throw new Error(`${label} 返回 HTML(路径 404): ${url}`);
  }
  return t;
}

async function loadPyodideScript() {
  await new Promise((res, rej) => {
    importScripts(PYODIDE_URL);  // Worker 内同步 importScripts
    res();
  });
}

async function init() {
  if (py) return py;
  if (!initPromise) {
    initPromise = (async () => {
      await loadPyodideScript();
      self.pyodide = await loadPyodide();
      py = self.pyodide;

      // 1) openpyxl + xlsxwriter (纯 Python, micropip 可装)
      await py.loadPackage("micropip");
      await py.runPythonAsync(
        `import micropip; await micropip.install("openpyxl"); await micropip.install("xlsxwriter")`);

      // 2) PyMuPDF WASM wheel(共享库, 必须 loadPackage)
      await py.loadPackage(WHEEL_URL);

      // 3) 引擎 + shim: 写虚拟 FS 后 import(不触发 __main__ 守卫)
      const [engSrc, shimSrc] = await Promise.all([
        fetchPySource(ENGINE_URL, "引擎源码"),
        fetchPySource(SHIM_URL, "shim"),
      ]);
      py.runPython(`import os; os.environ["PYODIDE"]="1"; os.makedirs("/lib/engine", exist_ok=True)`);
      py.FS.writeFile("/lib/engine/extract_bank_statement.py", engSrc);
      py.FS.writeFile("/lib/engine/shim.py", shimSrc);
      py.runPython(`import sys; sys.path.insert(0, "/lib/engine"); import shim`);
      py.runPython(`assert hasattr(shim, "convert_bytes")`);

      // 4) 页级进度桥: Python 侧留 install hook, JS 函数经 set_js_progress 注入
      //    (直接在 JsProxy 模块对象上 setattr JS 函数在部分版本会触发
      //     "Unsupported data type", 走 Python 函数中转更稳)
      py.runPython(`
def _install_js_progress(fn):
    import shim
    shim.js_progress_cb = fn
`);
      const installFn = py.globals.get("_install_js_progress");
      installFn((page, total) => {
        if (currentConvertId !== null) {
          self.postMessage({ type: "progress", id: currentConvertId, page, total });
        }
      });
      return py;
    })();
    initPromise.catch(() => { initPromise = null; py = null; });
  }
  return initPromise;
}

let currentConvertId = null;

function errPayload(e) {
  // Pyodide 的 PythonError.message 含完整 traceback; 取最后非空行(通常是
  // "XXXError: 中文消息")作为 UI 展示文本, 全文放 detail 供调试。
  const full = String(e.message || e);
  const lines = full.trim().split("\n").filter((l) => l.trim());
  let msg = lines[lines.length - 1] || full;
  // 去掉 "File ... line N, in ..." 样式的前缀噪音
  msg = msg.replace(/^.*?(?=[A-Za-z_]+Error:)/, "");
  return { message: msg.slice(0, 500), detail: full.slice(-1200) };
}

self.onmessage = async (e) => {
  const { type, id } = e.data || {};
  try {
    if (type === "init") {
      await init();
      self.postMessage({ type: "ready" });
      return;
    }
    if (type === "convert") {
      await init();
      currentConvertId = id;
      try {
        // FS 写入必须 Uint8Array(纯 ArrayBuffer 在部分 Pyodide 版本报 Unsupported data type)
        const u8 = e.data.pdfBytes instanceof Uint8Array
          ? e.data.pdfBytes : new Uint8Array(e.data.pdfBytes);
        py.FS.writeFile("input.pdf", u8);
        // 参数经 globals 传入(不拼源码, 避免引号注入与编码问题)
        py.globals.set("_pdf_password", e.data.password || "");
        py.globals.set("_pdf_sheet", e.data.sheet || "对账单");
        const js = py.runPython(`
import json, shim
r = shim.convert_bytes(
    open("input.pdf","rb").read(),
    password=_pdf_password,
    sheet=_pdf_sheet,
)
open("out.xlsx","wb").write(r["xlsx"])
json.dumps({
  "rows": r["rows"], "report": r["report"], "stats": r["stats"],
  "pages": r.get("pages"), "sheet": r["sheet"], "warnings": r["warnings"],
})
`);
        const meta = JSON.parse(js);
        const xlsx = py.FS.readFile("out.xlsx");
        // Transferable: 零拷贝转出
        self.postMessage({
          type: "done", id,
          xlsx: xlsx.buffer.slice(xlsx.byteOffset, xlsx.byteOffset + xlsx.byteLength),
          rows: meta.rows, report: meta.report, stats: meta.stats,
          pages: meta.pages, sheet: meta.sheet, warnings: meta.warnings,
        }, [xlsx.buffer]);
      } finally {
        currentConvertId = null;
      }
      return;
    }
    if (type === "diagnose") {
      await init();
      const u8 = e.data.pdfBytes instanceof Uint8Array
        ? e.data.pdfBytes : new Uint8Array(e.data.pdfBytes);
      py.FS.writeFile("diag_input.pdf", u8);
      const json = py.runPython(`
import json, shim
shim.export_diagnosis(open("diag_input.pdf","rb").read(), ${JSON.stringify(e.data.filename || "")})
`);
      self.postMessage({ type: "diag", id, json });
      return;
    }
  } catch (err) {
    self.postMessage({ type: "error", id, ...errPayload(err) });
  }
};
