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
// 资源列表(主+备): 优先 jsdelivr 中国节点, 失败回退同源 GitHub Pages
// 资源顺序 = [jsdelivr_gh_mirror, github_pages_same_origin]
// jsdelivr 的 gh 镜像对中国大陆移动网络通常比 GitHub Pages 快
function makeUrlPairs(relPath) {
  return [
    `https://cdn.jsdelivr.net/gh/jxuzhi-lab/bank2excel-h5@main/${relPath}`,
    new URL(relPath, self.location.href).toString(),
  ];
}
const WHEEL_PAIRS = makeUrlPairs("wheels/pymupdf-1.26.7-cp312-abi3-pyodide_2024_0_wasm32.whl");
const ENGINE_PAIRS = makeUrlPairs("python/extract_bank_statement.py");
const SHIM_PAIRS = makeUrlPairs("python/shim.py");

// 顺序尝试直到一个成功; 全失败抛 last error
async function fetchWithFallback(pairs, label) {
  let lastErr;
  for (const url of pairs) {
    try {
      const r = await fetch(url);
      if (!r.ok) throw new Error(`${label} HTTP ${r.status} (${url})`);
      return { response: r, url };
    } catch (e) {
      reportLog(`${label} 失败: ${url} → ${e.message}`);
      lastErr = e;
    }
  }
  throw lastErr;
}

async function fetchPySource(pairs, label) {
  const { response, url } = await fetchWithFallback(pairs, label);
  const t = await response.text();
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
      reportStage("pyodide-cdn", 5);
      await loadPyodideScript();
      reportStage("pyodide-init", 15);
      self.pyodide = await loadPyodide();
      py = self.pyodide;

      // 1) openpyxl + xlsxwriter (纯 Python, micropip 可装)
      reportStage("micropip", 25);
      await py.loadPackage("micropip");
      reportStage("openpyxl", 40);
      await py.runPythonAsync(
        `import micropip; await micropip.install("openpyxl")`);
      reportStage("xlsxwriter", 55);
      await py.runPythonAsync(
        `import micropip; await micropip.install("xlsxwriter")`);

      // 2) PyMuPDF WASM wheel(共享库)——用 micropip.install 单 URL 重试,
      //    比 loadPackage(URL列表)更稳。带 30s/URL 硬超时, 移动网卡住就跳下一个。
      reportStage("wheel-17mb", 70);
      let wheelInstalled = false;
      for (let i = 0; i < WHEEL_PAIRS.length; i++) {
        const url = WHEEL_PAIRS[i];
        const tag = i === 0 ? "主" : "备";
        reportLog(`wheel: 尝试${tag}源 ${i + 1}/${WHEEL_PAIRS.length} (${url})`);
        try {
          await Promise.race([
            py.runPythonAsync(`
import asyncio
import sys
import micropip
micropip.add_wheel_log_handler(lambda *a, **k: None)  # 抑制默认 stdout 日志
try:
    await asyncio.wait_for(
        micropip.install(${JSON.stringify(url)}, keep_going=False, deps=False, pre=False),
        timeout=60,
    )
except asyncio.TimeoutError:
    raise RuntimeError("wheel install timeout 60s")
`),
            new Promise((_, rej) => setTimeout(() => rej(new Error("wheel install timeout 60s")), 70_000)),
          ]);
          reportLog(`wheel: 来自 ${url} 安装成功`);
          wheelInstalled = true; break;
        } catch (e) {
          reportLog(`wheel: ${url} 失败 → ${e.message || e}`);
        }
      }
      if (!wheelInstalled) throw new Error("wheel 安装失败(主+备 URL 均失败, 请检查网络或刷新重试)");
      reportStage("engine-fetch", 85);

      // 3) 引擎 + shim: 写虚拟 FS 后 import(不触发 __main__ 守卫)
      const [engSrc, shimSrc] = await Promise.all([
        fetchPySource(ENGINE_PAIRS, "引擎源码"),
        fetchPySource(SHIM_PAIRS, "shim"),
      ]);
      reportStage("engine-import", 95);
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
      reportStage("ready", 100);
      return py;
    })();
    initPromise.catch((e) => {
      reportLog("init 失败: " + (e.message || e));
      initPromise = null; py = null;
    });
  }
  return initPromise;
}

let currentConvertId = null;

// ---- 阶段上报(替换假进度, 微信/慢网场景下能精确定位卡点) ----
function reportStage(stage, pct) {
  try { self.postMessage({ type: "stage", stage, pct }); } catch (e) {}
}
function reportLog(text) {
  try { self.postMessage({ type: "log", text }); } catch (e) {}
}

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
      // ready 在 init() 内 reportStage("ready", 100) 之后由 onmessage 这里再发一次
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
    reportLog("错误: " + (err.message || err));
    self.postMessage({ type: "error", id, ...errPayload(err) });
  }
};

self.onerror = (e) => {
  reportLog("worker onerror: " + (e.message || "unknown"));
};
