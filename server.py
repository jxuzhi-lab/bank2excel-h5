# -*- coding: utf-8 -*-
"""server.py —— bank2excel-h5 私有转换服务(完整引擎版, 2026-08-29)。

升级自 shim/PYODIDE 版: 直接调用完整引擎 convert_pdf(非 PYODIDE 路径),
激活引擎内置的两级失败升级链 ——
  1) 表头识别失败 → 视觉兜底(视觉只读一次表头列名, 锚点反查后仍由规则管道提取);
  2) 仍失败 → 自动 onboarding 生成格式描述符, 经真实提取验证通过后缓存并重试;
     描述符按首页文字指纹缓存, 同格式之后零成本秒转。

视觉 provider 由环境变量驱动(vision_utils.resolve):
  - 默认 auto: 无 node/vision.js 且未配置 API → 自动降级 none(纯规则), 零外部调用;
  - BANK_PDF_VISION_PROVIDER=api + BANK_PDF_VISION_API_BASE/API_KEY(/API_MODEL)
    → 未知格式自动走 OpenAI 兼容视觉模型兜底(标准库 urllib, 无新增 pip 依赖)。

安全/成本控制(环境变量, 均可省略):
  - VLM_BUDGET_PER_HOUR   全局每小时视觉调用预算, 超限自动降级纯规则(默认 40, 0=不限);
  - RATE_LIMIT_PER_MIN    每 IP 每分钟转换请求数(默认 12, 0=不限);
  - MAX_CONCURRENT        同时进行的转换数(默认 2);
  - DESCRIPTOR_CACHE_DIR  描述符缓存目录(建议挂卷持久化, 默认 ./_descriptors);
  - B2X_MAX_MB            上传上限 MB(默认 200)。

隐私: 文件在本机临时目录转换完即删; 仅当视觉启用且规则失败时, 第 1 页渲染图
才会发给所配置的外部视觉模型; 描述符只含列模板(列名+x 区间), 不含任何账单数据。
"""
import collections
import json
import os
import sys
import tempfile
import threading
import time

from fastapi import FastAPI, File, HTTPException, Request, UploadFile  # noqa: E402
from fastapi.responses import HTMLResponse, Response  # noqa: E402

BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(BASE, "python"))

import extract_bank_statement as eng  # noqa: E402  完整引擎(PYODIDE 未设 → CPython 全功能)
import vision_utils as vu  # noqa: E402
import ocr_layer  # noqa: E402  服务端扫描件 OCR 夹心层(OCR_ENABLED 可关)

app = FastAPI(title="bank2excel-h5 server", docs_url=None, redoc_url=None)

MAX_MB = int(os.environ.get("B2X_MAX_MB", "200") or "200")
DESC_CACHE_DIR = os.environ.get(
    "DESCRIPTOR_CACHE_DIR", os.path.join(BASE, "_descriptors"))
VLM_BUDGET = int(os.environ.get("VLM_BUDGET_PER_HOUR", "40") or "0")
RATE_PER_MIN = int(os.environ.get("RATE_LIMIT_PER_MIN", "12") or "0")
MAX_CONCURRENT = max(1, int(os.environ.get("MAX_CONCURRENT", "2") or "2"))
OCR_ENABLED = os.environ.get("OCR_ENABLED", "1").strip().lower() not in (
    "0", "false", "off", "no")

try:
    os.makedirs(DESC_CACHE_DIR, exist_ok=True)
except OSError:
    DESC_CACHE_DIR = os.path.join(tempfile.gettempdir(), "b2x_descriptors")
    os.makedirs(DESC_CACHE_DIR, exist_ok=True)

# ---- 视觉调用预算(包装 call_vision_raw 这唯一出口; 超限抛 VisionDisabledError,
# ---- 引擎各级升级链会将其按 RuntimeError 优雅降级, 不影响正常转换) ----
_orig_call_vision_raw = vu.call_vision_raw
_vlm_lock = threading.Lock()
_vlm_times = collections.deque()


def _budgeted_call_vision_raw(png_path, prompt, retries=1, timeout=150):
    if VLM_BUDGET:
        with _vlm_lock:
            now = time.time()
            while _vlm_times and now - _vlm_times[0] > 3600:
                _vlm_times.popleft()
            if len(_vlm_times) >= VLM_BUDGET:
                raise vu.VisionDisabledError(
                    f"VLM 小时预算已用尽({VLM_BUDGET}/小时), 本次降级纯规则路径")
            _vlm_times.append(now)
    return _orig_call_vision_raw(png_path, prompt, retries=retries, timeout=timeout)


vu.call_vision_raw = _budgeted_call_vision_raw

VLM_PROVIDER = vu.resolve_vision_provider(None)

# ---- 每 IP 限频(滑动窗口; uvicorn 多 worker 时各进程独立计数, 限额为近似值) ----
_rl_lock = threading.Lock()
_rl_hits = {}
_convert_sem = threading.Semaphore(MAX_CONCURRENT)


def _rate_limited(ip):
    if not RATE_PER_MIN:
        return False
    now = time.time()
    with _rl_lock:
        if len(_rl_hits) > 4096:  # 防表无限膨胀
            for k in [k for k, d in _rl_hits.items() if not d or now - d[-1] > 300]:
                _rl_hits.pop(k, None)
        dq = _rl_hits.setdefault(ip, collections.deque())
        while dq and now - dq[0] > 60:
            dq.popleft()
        if len(dq) >= RATE_PER_MIN:
            return True
        dq.append(now)
        return False


def _diag_payload(out_path, exc):
    """失败时聚合结构化错误: message/stage/suggestion + 诊断包 JSON(无 PNG)。"""
    detail = {"message": str(exc).split("\n")[-1][:300]}
    diag_path = os.path.splitext(out_path)[0] + ".diag.json"
    if os.path.exists(diag_path):
        try:
            with open(diag_path, encoding="utf-8") as f:
                diag = json.load(f)
            for k in ("stage", "suggestion", "expected", "escalation"):
                if diag.get(k):
                    detail[k] = diag[k]
            detail["diag"] = diag
            detail["diag"].pop("png", None)
        except Exception:  # noqa: BLE001
            pass
    return detail


@app.get("/", response_class=HTMLResponse)
def index():
    return INDEX_HTML.replace("__PRIVACY_NOTE__", PRIVACY_NOTE[VLM_PROVIDER != "none"])


@app.get("/api/health")
def health():
    try:
        n_desc = len([f for f in os.listdir(DESC_CACHE_DIR) if f.endswith(".json")])
    except OSError:
        n_desc = -1
    return {
        "ok": True,
        "engine": "full(escalation+onboard)",
        "vision_provider": VLM_PROVIDER,
        "vlm_budget_per_hour": VLM_BUDGET or None,
        "ocr_enabled": OCR_ENABLED,
        "descriptor_cache": {"dir": DESC_CACHE_DIR, "count": n_desc},
    }


@app.post("/api/convert")
def convert(request: Request, file: UploadFile = File(...), password: str = ""):
    client_ip = request.client.host if request.client else ""
    if _rate_limited(client_ip):
        raise HTTPException(429, detail="请求过于频繁, 请稍后再试")
    data = file.file.read()
    if len(data) > MAX_MB * 1024 * 1024:
        raise HTTPException(400, f"文件超过 {MAX_MB}MB 上限")
    if not data:
        raise HTTPException(400, "空文件")
    if not _convert_sem.acquire(timeout=120):
        raise HTTPException(503, detail="当前转换并发已满, 请稍后重试")
    try:
        with tempfile.TemporaryDirectory(prefix="b2x_") as td:
            pdf_path = os.path.join(td, "input.pdf")
            out_path = os.path.join(td, "out.xlsx")
            with open(pdf_path, "wb") as f:
                f.write(data)
            # 扫描件 OCR 夹心层(2026-08-29): 含无文字层的页时, 先 OCR 成
            # 不可见文字夹心层再走引擎 —— 对纯文字层 PDF 零开销。
            convert_src, ocr_pages = pdf_path, 0
            if OCR_ENABLED:
                try:
                    _doc = ocr_layer.pymupdf.open(pdf_path)
                    _need = ocr_layer.needs_ocr(_doc, max_pages=ocr_layer.PAGE_CAP)
                    _doc.close()
                    if _need:
                        scan_pdf = os.path.join(td, "scan_sandwich.pdf")
                        ocr_pages, _total, _secs = ocr_layer.build_text_layer(
                            pdf_path, scan_pdf)
                        if ocr_pages:
                            convert_src = scan_pdf
                except Exception as oe:  # noqa: BLE001 (OCR 失败按原样走引擎报扫描件错误)
                    ocr_pages = -1
                    print(f"[提示] OCR 预处理失败: {oe}")
            try:
                eng.convert_pdf(
                    convert_src, out_path=out_path, sheet="对账单",
                    password=password or None,
                    quick_classify=True, diag=True,
                    onboard_cache=DESC_CACHE_DIR,
                )
            except Exception as e:  # noqa: BLE001
                raise HTTPException(422, detail=_diag_payload(out_path, e))
            with open(out_path, "rb") as f:
                xlsx = f.read()
    finally:
        _convert_sem.release()
    out_name = os.path.splitext(file.filename or "对账单")[0] + ".xlsx"
    # 中文文件名: RFC 5987 编码(Header 必须 latin-1 可表示)
    from urllib.parse import quote
    cd = f"attachment; filename*=UTF-8''{quote(out_name)}"
    resp_headers = {"Content-Disposition": cd}
    if ocr_pages:
        resp_headers["X-OCR-Pages"] = str(ocr_pages)  # 负数=OCR 失败仍由规则路径拒绝
    return Response(
        content=xlsx,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=resp_headers,
    )


PRIVACY_NOTE = {
    True: ("隐私说明: 转换在您自己的服务器上完成, 文件转换后即删。未能自动识别的新格式"
           "会把第 1 页渲染图发送给您所配置的外部视觉模型辅助识别, 识别成功后仅保存"
           "不含任何数据的列模板(描述符)。"),
    False: ("隐私说明: 转换在您自己的服务器上完成, 文件转换后即删, 不发送任何数据给第三方。"),
}

INDEX_HTML = r"""<!DOCTYPE html>
<html lang="zh-CN"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>对账单 PDF→Excel · 私有服务</title>
<style>
  body{font-family:-apple-system,"Microsoft YaHei",sans-serif;background:#f6f8fa;margin:0;color:#1f2328}
  .wrap{max-width:640px;margin:0 auto;padding:24px 16px}
  h1{font-size:20px}#drop{border:2px dashed #b6c2cf;border-radius:12px;background:#fff;padding:36px 16px;text-align:center;margin-top:16px}
  #drop.over{border-color:#1a73e8;background:#f0f6ff}.btn{background:#1a73e8;color:#fff;border:0;padding:11px 26px;border-radius:8px;font-size:16px;cursor:pointer}
  .muted{color:#57606a;font-size:13px;margin-top:8px}
  #status{margin-top:14px;font-size:14px}#status .err{color:#d32f2f}
  .stat,.erbox{background:#fff;border:1px solid #e5e7eb;border-radius:10px;padding:12px;margin-top:12px;font-size:14px}
  .erbox{border-color:#f1b8b4;background:#fff5f5}
  pre{white-space:pre-wrap;font-size:12px;color:#57606a}
  .ok{color:#1a8a2e;font-weight:600}
  .linkbtn{background:none;border:none;color:#1a73e8;cursor:pointer;font-size:13px;padding:0;text-decoration:underline}
</style></head><body><div class="wrap">
<h1>银行对账单 PDF → Excel</h1>
<p class="muted">私有转换服务 · 文件转换完即删 · __PRIVACY_NOTE__</p>
<div id="drop">
  <div style="font-size:16px">点选或拖入对账单 PDF</div>
  <div style="margin-top:12px"><button class="btn" onclick="document.getElementById('f').click()">选择 PDF</button></div>
  <input type="file" id="f" accept=".pdf" hidden>
  <div class="muted">单个文件 ≤ 200MB · 加密 PDF 可附 ?password=</div>
</div>
<div id="status"></div>
<script>
const f=document.getElementById('f'),drop=document.getElementById('drop'),st=document.getElementById('status');
['dragover','dragenter'].forEach(e=>drop.addEventListener(e,x=>{x.preventDefault();drop.classList.add('over')}));
['dragleave','drop'].forEach(e=>drop.addEventListener(e,x=>{x.preventDefault();drop.classList.remove('over')}));
drop.addEventListener('drop',e=>{if(e.dataTransfer.files.length)go(e.dataTransfer.files[0])});
f.addEventListener('change',()=>{if(f.files.length)go(f.files[0])});
function showErr(d){
  const box=document.createElement('div');box.className='erbox';
  let h='<div class="err">❌ '+(d.message||'转换失败')+'</div>';
  if(d.stage)h+='<div class="muted">失败阶段: '+d.stage+'</div>';
  if(d.suggestion)h+='<div class="muted">建议: '+d.suggestion+'</div>';
  if(d.diag){
    h+='<button class="linkbtn" id="dl">下载诊断包(JSON, 发给维护者可加快适配)</button>';
    box.innerHTML=h;
    box.querySelector('#dl').addEventListener('click',()=>{
      const blob=new Blob([JSON.stringify(d.diag,null,1)],{type:'application/json'});
      const a=document.createElement('a');a.href=URL.createObjectURL(blob);
      a.download='bank2excel-diagnosis-'+Date.now()+'.json';a.click();
      setTimeout(()=>URL.revokeObjectURL(a.href),30000);
    });
  }else{box.innerHTML=h}
  st.innerHTML='';st.appendChild(box);
}
async function go(file){
  st.innerHTML='⏳ 上传并转换中...';
  const fd=new FormData();fd.append('file',file);
  const t0=performance.now();
  try{
    const r=await fetch('/api/convert',{method:'POST',body:fd});
    if(!r.ok){
      const j=await r.json().catch(()=>({}));
      const d=(j.detail&&typeof j.detail==='object')?j.detail:{message:(j.detail||('HTTP '+r.status))};
      showErr(d);return;
    }
    const blob=await r.blob();
    const name=file.name.replace(/\.pdf$/i,'')+'.xlsx';
    const a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download=name;a.click();
    setTimeout(()=>URL.revokeObjectURL(a.href),30000);
    const mb=(blob.size/1048576).toFixed(2);
    st.innerHTML=`<div class="stat"><span class="ok">✅ 转换成功</span><br>${name} (${mb}MB) · 耗时 ${Math.round(performance.now()-t0)/1000}s<br>文件已开始下载。</div>`;
  }catch(e){st.innerHTML=`<div class="err">❌ 网络错误: ${e.message}</div>`}
}
</script></div></body></html"""


def main():
    import argparse
    import uvicorn
    ap = argparse.ArgumentParser(description="bank2excel-h5 私有转换服务")
    ap.add_argument("--host", default="127.0.0.1")
    ap.add_argument("--port", type=int, default=8000)
    ap.add_argument("--workers", type=int, default=1)
    args = ap.parse_args()
    uvicorn.run(app, host=args.host, port=args.port, workers=args.workers)


if __name__ == "__main__":
    main()
