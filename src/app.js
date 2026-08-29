// app.js —— UI 状态机: idle → loading-engine → ready → converting → done|error
import { createEngine } from "./bridge.js";
import { downloadXlsx, isWeChatInnerBrowser, isIOS } from "./download.js";
import { showDiagnoseOffer } from "./diagnose.js";

const $ = (s) => document.querySelector(s);
const CONTACT_HTML = "邮箱: <b><a href='mailto:436089745@qq.com'>436089745@qq.com</a></b>"; // 诊断包发往此处

const state = { engine: null, engineReady: false, queue: [], current: null, results: [] };

// ---------- UI 骨架 ----------
const els = {
  drop: $("#drop"), file: $("#file"), pick: $("#pick"),
  engineBar: $("#engine-bar"), engineBarFill: $("#engine-bar-fill"),
  engineText: $("#engine-text"),
  pwdRow: $("#pwd-row"), pwd: $("#pwd"),
  queueList: $("#queue-list"), results: $("#results"),
  wechatTip: $("#wechat-tip"),
};

function setState(name) { document.body.dataset.state = name; }

// ---------- 引擎预加载 ----------
async function ensureEngine() {
  if (state.engineReady) return;
  setState("loading-engine");
  els.engineBar.style.display = "block";
  els.engineBarFill.style.width = "0%";
  state.engine = state.engine || createEngine();
  // 真实阶段上报: 每步 worker postMessage 上来, 立即更新 UI
  const STAGE_LABEL = {
    "pyodide-cdn":   "① 加载 Pyodide 运行时(CDN)...",
    "pyodide-init":  "② 初始化 Pyodide...",
    "micropip":      "③ 加载 micropip...",
    "openpyxl":      "④ 安装 openpyxl(PyPI)...",
    "xlsxwriter":    "⑤ 安装 xlsxwriter(PyPI)...",
    "wheel-17mb":    "⑥ 加载 PyMuPDF 引擎(17MB, 较慢)...",
    "engine-fetch":  "⑦ 拉取转换引擎源码...",
    "engine-import": "⑧ 注入引擎到 Pyodide...",
    "ready":         "引擎就绪 ✓",
  };
  state.engine.onStage((stage, pct) => {
    els.engineBarFill.style.width = (pct || 0) + "%";
    els.engineText.textContent = STAGE_LABEL[stage] || (stage + " " + (pct || 0) + "%");
  });
  try {
    await state.engine.init();
    setTimeout(() => { els.engineBar.style.display = "none"; }, 800);
    state.engineReady = true;
    setState("ready");
    drainQueue();
  } catch (e) {
    els.engineText.innerHTML = `❌ 引擎加载失败: ${escapeHtml(e.message)}<br>` +
      `建议: ① 微信内置浏览器可能对 jsdelivr/PyPI 限速或阻断 → 复制本页链接到系统浏览器(Chrome)打开; ` +
      `② 首次加载需联网拉取约 30MB 资源, 请检查网络; ③ 断网/防火墙场景无法使用。`;
    setState("error");
  }
}

// ---------- 文件选择/拖拽 ----------
els.pick.addEventListener("click", () => els.file.click());
els.file.addEventListener("change", () => addFiles(els.file.files));
["dragover", "dragenter"].forEach((ev) =>
  els.drop.addEventListener(ev, (e) => { e.preventDefault(); els.drop.classList.add("over"); }));
["dragleave", "drop"].forEach((ev) =>
  els.drop.addEventListener(ev, (e) => { e.preventDefault(); els.drop.classList.remove("over"); }));
els.drop.addEventListener("drop", (e) => addFiles(e.dataTransfer.files));

function addFiles(fileList) {
  const files = [...fileList].filter((f) => /\.pdf$/i.test(f.name) || f.type === "application/pdf");
  if (!files.length) { alert("请选择 PDF 文件"); return; }
  // 加密 PDF 提示密码框: 华夏/中行等网银导出常加密; 无法预知, 先统一展示可折叠密码框
  els.pwdRow.style.display = "block";
  for (const f of files) {
    state.queue.push({ file: f, status: "queued" });
    renderQueueItem(f.name, f.size);
  }
  // >300 页护栏在转换时已知页数后提示; 此处按文件体积粗提示
  for (const f of files) {
    if (f.size > 20 * 1024 * 1024) {
      toast(`「${f.name}」较大(${(f.size / 1048576).toFixed(1)}MB), 手机上可能较慢, 建议电脑浏览器处理。`);
    }
  }
  drainQueue();
}

// ---------- 队列处理 ----------
function renderQueueItem(name, size) {
  const row = document.createElement("div");
  row.className = "queue-item";
  row.innerHTML = `<span>${escapeHtml(name)}</span><span class="muted">(${(size / 1024).toFixed(0)} KB) — 排队中</span>`;
  els.queueList.appendChild(row);
  return row;
}

async function drainQueue() {
  if (!state.engineReady) { await ensureEngine(); if (!state.engineReady) return; }
  while (state.queue.length) {
    const item = state.queue.shift();
    await convertOne(item);
  }
  setState("ready");
}

function convertOne(item) {
  return new Promise((resolve) => {
    item.status = "converting";
    const row = addResultRow(item.file.name, "转换中...");
    const password = els.pwd.value.trim();
    setState("converting");
    const t0 = performance.now();
    state.engine.convert(item.file, { password }, {
      onProgress: (page, total) => {
        row.querySelector(".status").innerHTML =
          `<span class="pct">${page}/${total} 页</span>`;
      },
      onDone: (d) => {
        item.status = "done";
        const ms = Math.round(performance.now() - t0);
        row.querySelector(".status").innerHTML = `✅ 成功`;
        renderResultCard(row, item, d, ms);
        state.results.push({ name: item.file.name, ...d });
        resolve();
      },
      onError: (e) => {
        item.status = "error";
        const msg = e.message || "转换失败";
        row.querySelector(".status").innerHTML = `❌ 失败`;
        const detail = document.createElement("div");
        detail.className = "err-detail";
        detail.textContent = msg;
        if (e.detail) {
          const pre = document.createElement("details");
          pre.innerHTML = `<summary style="cursor:pointer;color:#8b949e">完整堆栈</summary>`;
          const preBody = document.createElement("pre");
          preBody.style.cssText = "font-size:11px;color:#8b949e;white-space:pre-wrap;max-height:200px;overflow:auto";
          preBody.textContent = e.detail;
          pre.appendChild(preBody);
          detail.appendChild(pre);
        }
        row.appendChild(detail);
        // 未识别格式 → 诊断包引导(排除密码错/扫描件, 那些有自己的提示)
        if (/表头|识别|记录/.test(msg)) {
          item.file.arrayBuffer().then((buf) => {
            showDiagnoseOffer(detail, {
              pdfBytes: new Uint8Array(buf),
              filename: item.file.name,
              rawError: msg,
              engine: state.engine,
              contactHTML: CONTACT_HTML,
            });
          });
        } else if (/加密|密码/.test(msg)) {
          detail.textContent += " —— 请在“PDF 密码(可选)”框输入密码后重试";
          els.pwdRow.style.display = "block";
        } else if (/扫描|OCR/.test(msg)) {
          detail.textContent += " —— 该文件是扫描件/纯图片, 无文字层, 本工具暂不支持";
        }
        resolve();
      },
    });
  });
}

// ---------- 结果卡片 ----------
function addResultRow(name, statusText) {
  const row = document.createElement("div");
  row.className = "result-row";
  row.innerHTML = `<span class="name">${escapeHtml(name)}</span><span class="status">${statusText}</span>`;
  els.results.prepend(row);
  return row;
}

function renderResultCard(row, item, d, ms) {
  const card = document.createElement("div");
  card.className = "result-card";
  const rep = d.report || [];
  const statsHtml = (Array.isArray(rep) ? rep : [String(rep)])
    .map((l) => `<div>${escapeHtml(String(l))}</div>`).join("");
  const wechatWarn = isWeChatInnerBrowser()
    ? `<div class="wechat-warn">检测到微信内嵌浏览器, 可能无法直接下载 → 请点右上角「···」选择「在浏览器打开」后再下载。</div>`
    : "";
  card.innerHTML = `
    <div class="rc-head">
      <span>${d.pages ?? "?"} 页 · ${d.rows ?? "?"} 笔 · 耗时 ${ms}ms</span>
      <button class="btn dl">⬇ 下载 Excel</button>
    </div>
    <div class="rc-stats">${statsHtml}</div>
    ${wechatWarn}
  `;
  card.querySelector(".dl").addEventListener("click", () => {
    const out = item.file.name.replace(/\.pdf$/i, "") + ".xlsx";
    downloadXlsx(d.xlsx, out);
  });
  row.appendChild(card);
}

// ---------- 杂项 ----------
function toast(msg) {
  const t = document.createElement("div");
  t.className = "toast";
  t.textContent = msg;
  document.body.appendChild(t);
  setTimeout(() => t.remove(), 6000);
}

// 应急: 清空 Service Worker + CacheStorage, 用于"代码已修但浏览器还在跑旧 SW"的情形
async function resetBrowserCache() {
  if (navigator.serviceWorker) {
    const regs = await navigator.serviceWorker.getRegistrations();
    for (const r of regs) await r.unregister();
  }
  if (typeof caches !== "undefined") {
    const names = await caches.keys();
    for (const n of names) await caches.delete(n);
  }
  location.reload();
}
function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, (c) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;",
  }[c]));
}

// 自动检测旧 SW: 启动时若有未识别版(sw.js 不含 v3+ 标记), 主动清理并刷新
(async function autoResetOldSW() {
  if (!navigator.serviceWorker) return;
  try {
    const regs = await navigator.serviceWorker.getRegistrations();
    for (const r of regs) {
      const script = r.active && r.active.scriptURL;
      // 用 v3 以后的 SW 都会带自身版本号在 URL 上; 旧版没带 → 清掉
      // 简单粗暴但有效: 任何已激活的 sw.js 都强制 unregister 一次
      if (script && /sw\.js(\?|$)/.test(script)) {
        await r.unregister();
        // 清掉 SW 缓存
        if (typeof caches !== "undefined") {
          const names = await caches.keys();
          for (const n of names) await caches.delete(n);
        }
        location.reload();
        return;
      }
    }
  } catch (e) { /* ignore */ }
})();

// 微信内嵌环境: 页面加载即提示
if (isWeChatInnerBrowser()) {
  els.wechatTip.style.display = "block";
}
// 首屏即预加载引擎(用户选文件时通常已就绪)
ensureEngine();
