# -*- coding: utf-8 -*-
"""server.py —— 银行对账单 PDF→Excel 私有转换服务(手机端稳定方案)。

用法:
  python server.py [--host 0.0.0.0] [--port 8000]

依赖: fastapi uvicorn python-multipart(pip install fastapi uvicorn python-multipart)
引擎: 复用 bank2excel-h5/python/ 的 shim.convert_bytes(本地 CPython + PyMuPDF)
"""
import argparse
import io
import os
import sys
import tempfile
import traceback

# 引擎路径: 本文件位于 bank2excel-h5/server.py, 引擎在 ./python/
HERE = os.path.dirname(os.path.abspath(__file__))
ENGINE_DIR = os.path.join(HERE, "python")
sys.path.insert(0, ENGINE_DIR)
os.environ.setdefault("PYODIDE", "")  # 本地 CPython 模式(非 PYODIDE)

import shim  # noqa: E402

from fastapi import FastAPI, File, UploadFile, HTTPException  # noqa: E402
from fastapi.responses import HTMLResponse, Response  # noqa: E402
from fastapi.middleware.cors import CORSMiddleware  # noqa: E402

app = FastAPI(title="bank2excel-h5 server", docs_url=None, redoc_url=None)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"], allow_methods=["*"], allow_headers=["*"],
)

MAX_MB = 200  # 单文件上限


@app.get("/", response_class=HTMLResponse)
def index():
    return INDEX_HTML


@app.post("/api/convert")
async def convert(file: UploadFile = File(...), password: str = ""):
    data = await file.read()
    if len(data) > MAX_MB * 1024 * 1024:
        raise HTTPException(400, f"文件超过 {MAX_MB}MB 上限")
    if not data:
        raise HTTPException(400, "空文件")
    try:
        r = shim.convert_bytes(data, password=password, sheet="对账单")
    except Exception as e:
        detail = str(e).split("\n")[-1][:200]
        raise HTTPException(422, detail=f"转换失败: {detail}")
    xlsx = r["xlsx"]
    out_name = os.path.splitext(file.filename or "对账单")[0] + ".xlsx"
    # 中文文件名: RFC 5987 编码(Header 必须 latin-1 可表示)
    from urllib.parse import quote
    cd = f"attachment; filename*=UTF-8''{quote(out_name)}"
    return Response(
        content=bytes(xlsx),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": cd},
    )


INDEX_HTML = """<!DOCTYPE html>
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
  .stat{background:#fff;border:1px solid #e5e7eb;border-radius:10px;padding:12px;margin-top:12px;font-size:14px}
  pre{white-space:pre-wrap;font-size:12px;color:#57606a}
  .ok{color:#1a8a2e;font-weight:600}
</style></head><body><div class="wrap">
<h1>银行对账单 PDF → Excel</h1>
<p class="muted">私有转换服务 · 文件直接到你自己的电脑处理, 转换完即丢</p>
<div id="drop">
  <div style="font-size:16px">点选或拖入对账单 PDF</div>
  <div style="margin-top:12px"><button class="btn" onclick="document.getElementById('f').click()">选择 PDF</button></div>
  <input type="file" id="f" accept=".pdf" hidden>
  <div class="muted">单个文件 ≤ 200MB</div>
</div>
<div id="status"></div>
<script>
const f=document.getElementById('f'),drop=document.getElementById('drop'),st=document.getElementById('status');
['dragover','dragenter'].forEach(e=>drop.addEventListener(e,x=>{x.preventDefault();drop.classList.add('over')}));
['dragleave','drop'].forEach(e=>drop.addEventListener(e,x=>{x.preventDefault();drop.classList.remove('over')}));
drop.addEventListener('drop',e=>{if(e.dataTransfer.files.length)go(e.dataTransfer.files[0])});
f.addEventListener('change',()=>{if(f.files.length)go(f.files[0])});
async function go(file){
  st.innerHTML='⏳ 上传并转换中...';
  const fd=new FormData();fd.append('file',file);
  const t0=performance.now();
  try{
    const r=await fetch('/api/convert',{method:'POST',body:fd});
    if(!r.ok){const j=await r.json().catch(()=>({}));st.innerHTML=`<div class="err">❌ ${j.detail||'HTTP '+r.status}</div>`;return}
    const blob=await r.blob();
    const name=file.name.replace(/\.pdf$/i,'')+'.xlsx';
    const a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download=name;a.click();
    setTimeout(()=>URL.revokeObjectURL(a.href),30000);
    const mb=(blob.size/1048576).toFixed(2);
    st.innerHTML=`<div class="stat"><span class="ok">✅ 转换成功</span><br>${name} (${mb}MB) · 耗时 ${Math.round(performance.now()-t0)/1000}s<br>文件已开始下载。</div>`;
  }catch(e){st.innerHTML=`<div class="err">❌ 网络错误: ${e.message}</div>`}
}
</script></div></body></html>"""


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--host", default="0.0.0.0")
    ap.add_argument("--port", type=int, default=8000)
    args = ap.parse_args()
    import uvicorn
    print(f"* 服务启动: http://{args.host}:{args.port}  (Ctrl+C 退出)")
    uvicorn.run(app, host=args.host, port=args.port, log_level="warning")


if __name__ == "__main__":
    main()
