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
import hashlib
import json
import os
import sys
import tempfile
import threading
import uuid
import time

from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile  # noqa: E402
from fastapi.responses import HTMLResponse, Response  # noqa: E402

BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(BASE, "python"))

import log_store  # noqa: E402  转换日志(SQLite, 保留 30 天)
import extract_bank_statement as eng  # noqa: E402  完整引擎(PYODIDE 未设 → CPython 全功能)
import vision_utils as vu  # noqa: E402
import ocr_layer  # noqa: E402  服务端扫描件 RapidOCR 夹心层(OCR_ENABLED 可关)
import glm_ocr  # noqa: E402  GLM-OCR 主路径(GLM_OCR_KEY 配置后启用)

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


def _diag_actual_placeholder():
    pass


def _write_glm_xlsx(header_names, records, out_path):
    """GLM-OCR 直接产出: 复用引擎写出链(typed_records 类型化 + 流式 xlsx)。"""
    typed = eng.typed_records(header_names, records)
    aligns, widths = eng._output_layout(header_names)
    eng._write_xlsx(out_path, "对账单", header_names, typed, aligns, widths,
                    keep_text=False)


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
def convert(request: Request, file: UploadFile = File(...),
            password: str = Form("")):  # Form() 必须显式声明, 否则被解析为 query 参数(踩坑)
    if _rate_limited(request.client.host if request.client else ""):
        raise HTTPException(429, detail="请求过于频繁, 请稍后再试")
    data = file.file.read()
    if len(data) > MAX_MB * 1024 * 1024:
        raise HTTPException(400, f"文件超过 {MAX_MB}MB 上限")
    if not data:
        raise HTTPException(400, "空文件")
    t0, usage = time.time(), {}
    stage, err_msg, rows = None, None, None
    if not _convert_sem.acquire(timeout=120):
        raise HTTPException(503, detail="当前转换并发已满, 请稍后重试")
    try:
        xlsx_bytes, res = _do_convert(data, password, usage)
        ocr_mode, ocr_pages, rows = res["ocr_mode"], res["ocr_pages"], res["rows"]
        stage = "ok"
    except HTTPException as he:
        if isinstance(he.detail, dict):
            stage = he.detail.get("stage") or stage
            err_msg = he.detail.get("message") or str(he.detail)[:200]
        else:
            err_msg = str(he.detail)
        raise
    except Exception as e:  # noqa: BLE001
        stage, err_msg = stage or "error", str(e)[:200]
        raise
    finally:
        _convert_sem.release()
        try:
            vlm = getattr(vu._LAST_API_USAGE, "usage", None) or {}
            u = dict(usage)
            u["vlm_calls"] = vlm.get("calls", 0)
            u["vlm_prompt"] = vlm.get("prompt", 0)
            u["vlm_completion"] = vlm.get("completion", 0)
            log_store.add_entry(file.filename, len(data),
                                "fail" if err_msg else "ok", stage,
                                ocr_mode or "rule", time.time() - t0,
                                rows, u, err_msg)
        except Exception:  # noqa: BLE001
            pass
    out_name = os.path.splitext(file.filename or "对账单")[0] + ".xlsx"
    # 中文文件名: RFC 5987 编码(Header 必须 latin-1 可表示)
    from urllib.parse import quote
    cd = "attachment; filename*=UTF-8''" + quote(out_name)
    resp_headers = {"Content-Disposition": cd}
    if ocr_pages:
        resp_headers["X-OCR-Pages"] = str(ocr_pages)
        if ocr_mode:
            resp_headers["X-OCR-Mode"] = ocr_mode
        if res.get("warning"):
            resp_headers["X-OCR-Warning"] = quote(res["warning"])
    return Response(content=xlsx_bytes, media_type="application/vnd"
                    ".openxmlformats-officedocument.spreadsheetml.sheet",
                    headers=resp_headers)


def _do_convert(data, password, usage):
    """共享转换核心(同步接口与异步任务队列共用): 落盘→OCR 分支→引擎→读出 xlsx。
    返回 (xlsx_bytes, {"ocr_mode","ocr_pages","rows","warning"});
    失败抛 HTTPException(带 stage/suggestion 结构化 detail)。"""
    ocr_mode, ocr_pages, glm_direct = "", 0, False
    ocr_warning = ""
    rows = None
    stage = err_msg = None
    try:
        with tempfile.TemporaryDirectory(prefix="b2x_") as td:
            pdf_path = os.path.join(td, "input.pdf")
            out_path = os.path.join(td, "out.xlsx")
            with open(pdf_path, "wb") as f:
                f.write(data)
            # 扫描件 OCR: GLM-OCR 主路径(退化自愈) → 百度定向兜底 → 夹心层保底
            glm_direct = False
            if OCR_ENABLED:
                try:
                    _doc = ocr_layer.pymupdf.open(pdf_path)
                    try:
                        _need = ocr_layer.needs_ocr(_doc, max_pages=ocr_layer.PAGE_CAP)
                        _n_pages = len(_doc)
                    finally:
                        # 加密文件 needs_ocr 会抛异常, close 必须兜底
                        # (Windows 句柄不释放会让临时目录清理失败→500, 踩坑)
                        _doc.close()
                    if _need:
                        glm = None
                        if glm_ocr.get_key():
                            try:
                                glm = glm_ocr.parse_scanned_pdf(
                                    pdf_path, _n_pages, usage=usage)
                            except Exception as ge:  # noqa: BLE001
                                print(f"[提示] GLM-OCR 异常, 转夹心层: {ge}", flush=True)
                                glm = None
                        if glm is not None:
                            bad_idx = [p["index"] + 1 for p in glm["pages"]
                                       if p["hard_fail"]]
                            if glm["errors"]:
                                print(f"[提示] GLM-OCR: {glm['errors']}", flush=True)
                            if bad_idx:
                                raise HTTPException(
                                    422, detail={
                                        "message":
                                            f"第 {bad_idx} 页 OCR 识别失败(重试后仍无有效"
                                            "数据), 为保证数据完整未输出结果",
                                        "stage": "ocr",
                                        "suggestion": "请重试一次; 若持续失败请导出诊断反馈",
                                        "fail_pages": bad_idx})
                            glm["pages"], gaps = glm_ocr.refine_gaps(
                                pdf_path, glm["pages"], debug=True, usage=usage)
                            _hdr, _recs, _meta = glm_ocr.build_engine_records(glm["pages"])
                            rows = len(_recs)
                            _write_glm_xlsx(_hdr, _recs, out_path)
                            glm_direct = True
                            ocr_pages = _n_pages
                            ocr_mode = ("glm-ocr+baidu"
                                        if any(p.get("source") == "baidu"
                                               for p in glm["pages"]) else "glm-ocr")
                            salv = [p["index"] + 1 for p in glm["pages"]
                                    if p["salvaged"] and p.get("source") != "baidu"]
                            if salv or gaps:
                                parts = []
                                if salv:
                                    parts.append(f"第 {salv} 页识别曾截断已抢救")
                                parts.append(
                                    f"余额链断点行: {gaps[:8]} — 请核对断点处是否遗漏记录"
                                    if gaps else "余额链完整")
                                ocr_warning = "; ".join(parts)
                        if not glm_direct:
                            scan_pdf = os.path.join(td, "scan_sandwich.pdf")
                            ocr_pages, _total, _secs = ocr_layer.build_text_layer(
                                pdf_path, scan_pdf)
                            if ocr_pages:
                                pdf_path = scan_pdf
                                ocr_mode = "ocr-sandwich"
                except HTTPException:
                    raise
                except Exception as oe:  # noqa: BLE001 (OCR 失败按原样走引擎报扫描件错误)
                    ocr_pages = -1
                    ocr_mode = f"error: {oe}"
                    err_msg = f"OCR 预处理失败: {oe}"
                    print(f"[提示] OCR 预处理失败: {oe}", flush=True)
            try:
                r = eng.convert_pdf(
                    pdf_path, out_path=out_path, sheet="对账单",
                    password=password or None,
                    quick_classify=True, diag=True,
                    onboard_cache=DESC_CACHE_DIR,
                )
                try:
                    rows = r[1].get("rows")
                except Exception:  # noqa: BLE001
                    rows = None
            except HTTPException as he:
                if isinstance(he.detail, dict):
                    stage = he.detail.get("stage") or stage
                    err_msg = he.detail.get("message") or str(he.detail)[:200]
                raise
            except Exception as e:  # noqa: BLE001
                raise HTTPException(422, detail=_diag_payload(out_path, e))
            with open(out_path, "rb") as f:
                xlsx_bytes = f.read()
        return xlsx_bytes, {"ocr_mode": ocr_mode, "ocr_pages": ocr_pages,
                            "rows": rows, "warning": ocr_warning}
    except HTTPException as he:
        if isinstance(he.detail, dict):
            stage = he.detail.get("stage") or stage
            err_msg = he.detail.get("message") or str(he.detail)[:200]
        raise
    except Exception as e:  # noqa: BLE001
        stage, err_msg = stage or "error", str(e)[:200]
        raise HTTPException(422, detail={"message": err_msg, "stage": stage})


# ---- 异步任务队列(移动端友好): 提交即返回 task_id, 轮询取结果 ----
# 上传后连接断开/切后台不影响服务端继续转换, 规避手机浏览器杀连接问题。
TASKS = {}
_tasks_lock = threading.Lock()
_RESULTS_DIR = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "_task_results")
os.makedirs(_RESULTS_DIR, exist_ok=True)


def _task_worker(tid, data, filename, password, usage):
    t0 = time.time()
    with _tasks_lock:
        TASKS[tid]["status"] = "running"
    stage = err_msg = rows = None
    ocr_mode, ocr_pages = "", 0
    try:
        if not _convert_sem.acquire(timeout=300):
            raise HTTPException(503, detail="当前转换并发已满")
        try:
            with tempfile.TemporaryDirectory(prefix="b2x_") as td:
                pdf_path = os.path.join(td, "input.pdf")
                out_path = os.path.join(td, "out.xlsx")
                with open(pdf_path, "wb") as f:
                    f.write(data)
                try:
                    xlsx_bytes, res = _do_convert(data, password, usage)
                except HTTPException as he:
                    if isinstance(he.detail, dict):
                        stage = he.detail.get("stage")
                        err_msg = he.detail.get("message") or str(he.detail)[:200]
                    raise
                out_name = os.path.splitext(filename or "对账单")[0] + ".xlsx"
                # 结果存任务目录(临时目录会随请求结束删除——踩坑)
                keep = os.path.join(_RESULTS_DIR, tid + ".xlsx")
                with open(keep, "wb") as f:
                    f.write(xlsx_bytes)
                with _tasks_lock:
                    TASKS[tid].update({
                        "status": "ok", "rows": res["rows"],
                        "ocr_pages": res["ocr_pages"], "ocr_mode": res["ocr_mode"],
                        "warning": res["warning"], "result_path": keep,
                        "out_name": out_name,
                        "duration_s": round(time.time() - t0, 1)})
                rows = res["rows"]
        except HTTPException as he:
            if isinstance(he.detail, dict):
                stage = he.detail.get("stage")
                err_msg = he.detail.get("message") or str(he.detail)[:200]
            raise
    except Exception as e:  # noqa: BLE001
        with _tasks_lock:
            TASKS[tid].update({"status": "fail", "stage": stage or "error",
                               "error": err_msg or str(e)[:200],
                               "duration_s": round(time.time() - t0, 1)})
    finally:
        _convert_sem.release()
        try:
            vlm = getattr(vu._LAST_API_USAGE, "usage", None) or {}
            u = dict(usage)
            u["vlm_calls"] = vlm.get("calls", 0)
            u["vlm_prompt"] = vlm.get("prompt", 0)
            u["vlm_completion"] = vlm.get("completion", 0)
            log_store.add_entry(filename, len(data),
                                "fail" if err_msg else "ok", stage,
                                ocr_mode or "rule", time.time() - t0,
                                rows, u, err_msg)
        except Exception:  # noqa: BLE001
            pass


def _purge_tasks():
    cutoff = time.time() - 6 * 3600
    with _tasks_lock:
        for tid in [t for t, v in TASKS.items()
                    if v.get("created", 0) < cutoff]:
            TASKS.pop(tid, None)
            try:
                os.remove(os.path.join(_RESULTS_DIR, tid + ".xlsx"))
            except OSError:
                pass


@app.post("/api/tasks")
def create_task(request: Request, file: UploadFile = File(...), password: str = Form("")):
    if _rate_limited(request.client.host if request.client else ""):
        raise HTTPException(429, detail="请求过于频繁, 请稍后再试")
    data = file.file.read()
    if len(data) > MAX_MB * 1024 * 1024:
        raise HTTPException(400, f"文件超过 {MAX_MB}MB 上限")
    if not data:
        raise HTTPException(400, "空文件")
    _purge_tasks()
    tid = uuid.uuid4().hex
    with _tasks_lock:
        TASKS[tid] = {"status": "queued", "created": time.time(),
                      "filename": file.filename or "对账单.pdf"}
    threading.Thread(target=_task_worker,
                     args=(tid, data, file.filename, password, {}),
                     daemon=True).start()
    return {"task_id": tid}


@app.get("/api/tasks/{tid}")
def task_status(tid: str):
    with _tasks_lock:
        t = TASKS.get(tid)
        if not t:
            raise HTTPException(404, "任务不存在或已过期")
        return {k: t.get(k) for k in ("status", "stage", "error", "rows",
                                      "ocr_pages", "ocr_mode", "warning",
                                      "duration_s", "filename")}


@app.get("/api/tasks/{tid}/result")
def task_result(tid: str):
    with _tasks_lock:
        t = TASKS.get(tid)
        if not t:
            raise HTTPException(404, "任务不存在或已过期")
        if t.get("status") != "ok":
            raise HTTPException(409, "任务尚未完成")
        path, out_name = t["result_path"], t.get("out_name", "对账单.xlsx")
    from urllib.parse import quote
    with open(path, "rb") as f:
        content = f.read()
    return Response(
        content=content,
        media_type="application/vnd.openxmlformats-officedocument"
                   ".spreadsheetml.sheet",
        headers={"Content-Disposition":
                 "attachment; filename*=UTF-8''" + quote(out_name)})




# ---- /admin 转换日志后台(密码保护, 2026-08-30; v7 补丁曾误删, 已恢复) ----
ADMIN_PASSWORD = os.environ.get("ADMIN_PASSWORD", "")
ADMIN_COOKIE = "b2x_admin"


def _admin_token():
    return hashlib.sha256(("b2x-admin:" + ADMIN_PASSWORD).encode()).hexdigest()[:32]


def _admin_authed(request: Request) -> bool:
    return bool(ADMIN_PASSWORD) and request.cookies.get(ADMIN_COOKIE) == _admin_token()


@app.get("/admin")
def admin_page(request: Request):
    if not ADMIN_PASSWORD:
        return HTMLResponse("<h2>未配置 ADMIN_PASSWORD 环境变量, 后台已禁用</h2>", 403)
    if not _admin_authed(request):
        return HTMLResponse(LOGIN_HTML)
    return HTMLResponse(ADMIN_HTML)


@app.post("/admin/login")
def admin_login(password: str = Form("")):
    if ADMIN_PASSWORD and password == ADMIN_PASSWORD:
        resp = Response(status_code=303)
        resp.headers["Location"] = "/admin"
        resp.set_cookie(ADMIN_COOKIE, _admin_token(), httponly=True,
                        samesite="lax", max_age=86400 * 7)
        return resp
    return HTMLResponse(LOGIN_HTML)


@app.get("/admin/logout")
def admin_logout():
    resp = Response(status_code=303)
    resp.headers["Location"] = "/admin"
    resp.delete_cookie(ADMIN_COOKIE)
    return resp


@app.get("/admin/api/summary")
def admin_summary(request: Request, days: int = 7):
    if not _admin_authed(request):
        raise HTTPException(401, "未登录")
    log_store.maybe_purge()
    return log_store.stats(days=min(max(days, 1), 30))


@app.get("/admin/api/logs")
def admin_logs(request: Request, days: int = 7, status: str = "",
               limit: int = 300, offset: int = 0):
    if not _admin_authed(request):
        raise HTTPException(401, "未登录")
    return log_store.query(days=min(max(days, 1), 30),
                           status=status or None,
                           limit=min(limit, 100000), offset=offset)


@app.get("/admin/api/export")
def admin_export(request: Request, days: int = 30):
    if not _admin_authed(request):
        raise HTTPException(401, "未登录")
    return Response(content=log_store.export_json(days=min(max(days, 1), 30)),
                    media_type="application/json")

log_store.maybe_purge(force=True)  # 启动即清理过期日志

PRIVACY_NOTE = {
    True: ("隐私说明: 转换在您自己的服务器上完成, 文件转换后即删。未能自动识别的新格式"
           "会把第 1 页渲染图发送给您所配置的外部视觉模型辅助识别, 识别成功后仅保存"
           "不含任何数据的列模板(描述符)。"),
    False: ("隐私说明: 转换在您自己的服务器上完成, 文件转换后即删, 不发送任何数据给第三方。"),
}

ADMIN_HTML = r"""<!DOCTYPE html>
<html lang="zh-CN"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>转换日志后台 · bank2excel</title>
<style>
  *{box-sizing:border-box}
  body{font-family:-apple-system,"PingFang SC","Microsoft YaHei",sans-serif;background:#10141c;color:#e6e9ef;margin:0}
  .wrap{max-width:1100px;margin:0 auto;padding:22px 14px 40px}
  .head{display:flex;align-items:center;justify-content:space-between;margin-bottom:16px;gap:10px;flex-wrap:wrap}
  h1{font-size:19px;margin:0}
  .tools{display:flex;gap:8px;align-items:center;flex-wrap:wrap}
  select,button,input{background:#1a2130;color:#e6e9ef;border:1px solid #2b3346;border-radius:8px;padding:7px 11px;font-size:13px}
  button{cursor:pointer}button:hover{background:#232c40}
  .cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:10px;margin-bottom:14px}
  .c{background:#171d2b;border:1px solid #2b3346;border-radius:12px;padding:12px 14px}
  .c .v{font-size:21px;font-weight:700}
  .c .k{font-size:12px;color:#8b94a7;margin-top:2px}
  .c.ok .v{color:#4ade80}.c.fail .v{color:#f87171}
  .chart{background:#171d2b;border:1px solid #2b3346;border-radius:12px;padding:14px;margin-bottom:14px}
  .chart .t{font-size:12.5px;color:#8b94a7;margin-bottom:8px}
  .bars{display:flex;align-items:flex-end;gap:5px;height:90px}
  .bars>div{flex:1;background:#3b82f6;border-radius:4px 4px 0 0;min-height:2px;position:relative}
  .bars>div span{position:absolute;bottom:-18px;left:0;right:0;text-align:center;font-size:10px;color:#8b94a7}
  .bars>div b{position:absolute;top:-16px;left:0;right:0;text-align:center;font-size:10px;color:#e6e9ef}
  table{width:100%;border-collapse:collapse;font-size:12.5px;background:#171d2b;border:1px solid #2b3346;border-radius:12px;overflow:hidden}
  th{background:#1a2130;color:#8b94a7;text-align:left;padding:8px 9px;font-weight:600;white-space:nowrap}
  td{padding:7px 9px;border-top:1px solid #232b3d;white-space:nowrap;max-width:230px;overflow:hidden;text-overflow:ellipsis}
  tr:hover td{background:#1a2233}
  .ok{color:#4ade80}.fail{color:#f87171}.m{color:#8b94a7}
</style></head><body><div class="wrap">
<div class="head">
  <h1>转换日志后台</h1>
  <div class="tools">
    <select id="days"><option value="1">今天</option><option value="7" selected>近 7 天</option><option value="30">近 30 天</option></select>
    <select id="st"><option value="">全部状态</option><option value="ok">成功</option><option value="fail">失败</option></select>
    <input id="q" placeholder="搜文件名" style="width:150px">
    <button onclick="load()">刷新</button>
    <label style="font-size:12.5px;color:#8b94a7"><input type="checkbox" id="auto" checked style="margin-right:4px">30s 自动</label>
    <button onclick="exportCsv()">导出 CSV</button>
    <button onclick="location='/admin/logout'">退出</button>
  </div>
</div>
<div class="cards" id="cards"></div>
<div class="chart"><div class="t">每日转换量(近 14 天)</div><div class="bars" id="bars"></div><div style="height:18px"></div></div>
<table id="tbl"><thead><tr>
  <th>请求时间</th><th>文件</th><th>状态</th><th>失败阶段</th><th>通道</th>
  <th>耗时</th><th>行数</th><th>GLM-OCR</th><th>视觉兜底</th><th>百度点数</th><th>错误</th>
</tr></thead><tbody></tbody></table>
<script>
function fmtT(s){return s>=60?(s/60).toFixed(1)+' 分':s.toFixed(1)+' 秒'}
function esc(s){return String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]))}
async function load(){
  const days=document.getElementById('days').value, st=document.getElementById('st').value,
        q=document.getElementById('q').value.trim();
  const [sum, logs] = await Promise.all([
    fetch('/admin/api/summary?days='+days).then(r=>r.json()),
    fetch('/admin/api/logs?days='+days+'&status='+st+'&limit=300').then(r=>r.json())
  ]);
  const rate = sum.total ? Math.round(sum.ok*100/sum.total) : '-';
  document.getElementById('cards').innerHTML = [
    ['总数', sum.total, ''], ['成功', sum.ok, 'ok'], ['失败', sum.fail, 'fail'],
    ['成功率', rate==='-'?'-':rate+'%', ''],
    ['总耗时', fmtT(sum.duration_sum), ''],
    ['GLM-OCR', sum.glm.calls+' 次 / '+(sum.glm.prompt+sum.glm_completion).toLocaleString()+' tok', ''],
    ['视觉兜底', sum.vlm.calls+' 次 / '+(sum.vlm.prompt+sum.vlm_completion).toLocaleString()+' tok', ''],
    ['百度', sum.baidu_calls+' 次 / '+sum.baidu_points+' 点', ''],
  ].map(c=>'<div class="c '+c[2]+'"><div class="v">'+c[1]+'</div><div class="k">'+c[0]+'</div></div>').join('');
  const mx = Math.max(1, ...sum.daily.map(d=>d.total));
  document.getElementById('bars').innerHTML = sum.daily.map(d=>
    '<div style="height:'+Math.max(3,d.total*100/mx)+'%"><b>'+d.total+'</b><span>'+d.day.slice(5)+'</span></div>').join('');
  const q2=q.toLowerCase();
  const tb=document.querySelector('#tbl tbody');
  tb.innerHTML = logs.filter(l=>!q2||String(l.filename||'').toLowerCase().includes(q2))
    .map(l=>'<tr><td>'+esc(l.ts)+'</td><td title="'+esc(l.filename)+'">'+esc((l.filename||'').slice(0,32))+
      '</td><td class="'+l.status+'">'+(l.status==='ok'?'成功':'失败')+'</td><td class="m">'+esc(l.stage||'-')+
      '</td><td class="m">'+esc(l.mode||'-')+'</td><td>'+fmtT(l.duration_s)+'</td><td>'+(l.rows??'-')+
      '</td><td class="m">'+(l.glm_calls?(l.glm_calls+'次/'+(l.glm_prompt+l.glm_completion).toLocaleString()):'-')+
      '</td><td class="m">'+(l.vlm_calls?(l.vlm_calls+'次/'+(l.vlm_prompt+l.vlm_completion).toLocaleString()):'-')+
      '</td><td class="m">'+(l.baidu_calls?(l.baidu_calls*25+' 点'):'-')+
      '</td><td class="m" title="'+esc(l.error)+'">'+esc((l.error||'').slice(0,26))+'</td></tr>').join('')
    || '<tr><td colspan="11" style="text-align:center;color:#8b94a7;padding:20px">无记录</td></tr>';
}
async function exportCsv(){
  const days=document.getElementById('days').value;
  const logs = await fetch('/admin/api/logs?days='+days+'&limit=100000').then(r=>r.json());
  const cols = ['ts','filename','size','status','stage','mode','duration_s','rows','glm_calls','glm_prompt','glm_completion','vlm_calls','vlm_prompt','vlm_completion','baidu_calls','error'];
  const csv = '\ufeff' + cols.join(',') + '\n' + logs.map(l=>cols.map(c=>'"'+String(l[c]??'').replace(/"/g,'""')+'"').join(',')).join('\n');
  const a=document.createElement('a');a.href=URL.createObjectURL(new Blob([csv],{type:'text/csv'}));
  a.download='convert-logs-'+Date.now()+'.csv';a.click();
}
document.getElementById('days').onchange=load;
document.getElementById('st').onchange=load;
document.getElementById('q').oninput=load;
setInterval(()=>{if(document.getElementById('auto').checked)load()},30000);
load();
</script></div></body></html>
"""

LOGIN_HTML = r"""<!DOCTYPE html><html lang="zh-CN"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1"><title>后台登录</title>
<style>body{font-family:-apple-system,"PingFang SC","Microsoft YaHei",sans-serif;background:#10141c;color:#e6e9ef}
.login{max-width:340px;margin:14vh auto;background:#171d2b;border:1px solid #2b3346;border-radius:14px;padding:26px}
input{width:100%;margin:10px 0 14px;background:#1a2130;color:#e6e9ef;border:1px solid #2b3346;border-radius:8px;padding:10px;font-size:14px}
.btn{width:100%;background:#3b82f6;border:0;color:#fff;padding:10px;border-radius:8px;font-weight:600;cursor:pointer}
h2{font-size:17px;margin:0 0 6px}.m{color:#8b94a7;font-size:12.5px}</style></head><body>
<div class="login"><h2>转换日志后台</h2><div class="m">请输入管理密码</div>
<form method="post" action="/admin/login"><input type="password" name="password" autofocus>
<button class="btn" type="submit">登录</button></form></div></body></html>
"""

INDEX_HTML = r"""<!DOCTYPE html>
<html lang="zh-CN"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>对账单 PDF → Excel · 私有转换</title>
<style>
  *{box-sizing:border-box}
  body{font-family:-apple-system,"PingFang SC","Microsoft YaHei",sans-serif;background:#f3f5f9;margin:0;color:#1f2328;-webkit-font-smoothing:antialiased}
  .wrap{max-width:760px;margin:0 auto;padding:24px 14px 44px}
  .head{display:flex;align-items:center;justify-content:space-between;margin-bottom:16px}
  h1{font-size:21px;margin:0}
  .sub{color:#8a94a6;font-size:12.5px;margin-top:3px}
  .helpbtn{background:#fff;border:1px solid #d5dbe3;color:#344054;border-radius:999px;
    padding:6px 14px;font-size:13px;cursor:pointer;flex-shrink:0;font-weight:600}
  .helpbtn:hover{background:#f6f8fb}
  .card{background:#fff;border:1px solid #e7eaf0;border-radius:14px;box-shadow:0 1px 3px rgba(16,24,40,.05);padding:18px;margin-bottom:14px}
  #drop{border:2px dashed #c3cdda;border-radius:12px;padding:26px 14px;text-align:center;cursor:pointer;transition:all .15s}
  #drop:hover{border-color:#1a73e8;background:#f5f9ff}
  #drop.over{border-color:#1a73e8;background:#eef5ff;transform:scale(1.01)}
  #drop .big{font-size:15.5px;font-weight:600;margin-bottom:5px}
  #drop .hint{color:#8a94a6;font-size:12.5px}
  #drop .hint.warn{color:#d92d20;font-weight:600}
  .queue{margin-top:4px}
  .item{display:flex;align-items:center;gap:9px;padding:10px 2px;border-bottom:1px solid #f0f2f6;font-size:13px}
  .item:last-child{border-bottom:0}
  .ficon{width:32px;height:32px;border-radius:8px;background:#eef3fb;color:#1a73e8;display:flex;align-items:center;justify-content:center;font-size:10.5px;font-weight:700;flex-shrink:0}
  .fmeta{flex:1;min-width:0}
  .fname{font-weight:600;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
  .fmsg{color:#8a94a6;font-size:12px;margin-top:2px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
  .fmsg.err{color:#d92d20}
  .fmsg.ok{color:#12805c}
  .badge{flex-shrink:0;font-size:11.5px;padding:3px 9px;border-radius:999px;background:#f2f4f7;color:#667085;font-weight:600}
  .badge.run{background:#e5f0ff;color:#1a73e8}
  .badge.ok{background:#e6f6ef;color:#12805c}
  .badge.fail{background:#fdecea;color:#d92d20}
  .pin{flex-shrink:0;width:96px;border:1px solid #d5dbe3;border-radius:7px;padding:6px 8px;font-size:12.5px;outline:none;transition:border .15s}
  .pin:focus{border-color:#1a73e8;box-shadow:0 0 0 3px rgba(26,115,232,.12)}
  .pin:disabled{background:#f6f8fb;color:#98a2b3}
  .btn{background:#1a73e8;color:#fff;border:0;padding:10px 22px;border-radius:9px;font-size:14px;font-weight:600;cursor:pointer;transition:background .15s}
  .btn:hover{background:#1662c6}
  .btn:disabled{background:#c2d6f2;cursor:not-allowed}
  .btn.ghost{background:#fff;color:#344054;border:1px solid #d5dbe3}
  .btn.ghost:hover{background:#f6f8fb}
  .linkbtn{background:none;border:none;color:#1a73e8;cursor:pointer;font-size:12.5px;padding:0;text-decoration:underline;flex-shrink:0}
  .linkbtn.del{color:#d92d20;font-size:15px;text-decoration:none;font-weight:700;padding:2px 6px;border-radius:6px}
  .linkbtn.del:hover{background:#fdecea}
  .linkbtn.del:disabled{color:#d0d5dd;cursor:not-allowed;background:none}
  .bar{display:flex;align-items:center;justify-content:space-between;margin-top:14px;gap:10px}
  .progress{color:#667085;font-size:13px}
  .empty{color:#98a2b3;font-size:13px;text-align:center;padding:14px 0 4px}
  .err-detail{background:#fff8f7;border:1px solid #f4d3cf;border-radius:8px;padding:8px 10px;margin-top:6px;font-size:12.5px;color:#7a271a;line-height:1.55}
  /* ---- 帮助弹窗(二级界面) ---- */
  #mask{position:fixed;inset:0;background:rgba(15,23,42,.45);display:none;align-items:flex-end;justify-content:center;z-index:50}
  #mask.show{display:flex}
  #help{background:#fff;width:100%;max-width:560px;max-height:82vh;overflow-y:auto;
    border-radius:18px 18px 0 0;padding:20px 20px 28px}
  @media(min-width:600px){#mask{align-items:center}#help{border-radius:16px;padding-bottom:20px}}
  #help h2{font-size:17px;margin:0 0 12px;display:flex;justify-content:space-between;align-items:center}
  #help .x{background:none;border:none;font-size:20px;color:#667085;cursor:pointer;padding:2px 8px}
  #help h3{font-size:13.5px;margin:14px 0 6px;color:#1a73e8}
  #help p{font-size:13px;color:#475467;line-height:1.75;margin:0 0 4px}
  #help .hl{background:#eef5ff;border-radius:8px;padding:10px 12px;font-size:12.5px;color:#3b4453;line-height:1.7;margin-top:12px}
</style></head><body><div class="wrap">
<div class="head">
  <div><h1>对账单 PDF → Excel</h1><div class="sub">文件转换完即删 · 私有部署</div></div>
  <button class="helpbtn" id="helpbtn">ⓘ 帮助</button>
</div>
<div class="card">
  <div id="drop">
    <div class="big">点选或拖入对账单 PDF</div>
    <div class="hint" id="hint">可多选, 最多 10 个 · 单个 ≤ 200MB</div>
  </div>
  <input type="file" id="f" accept=".pdf" multiple hidden>
</div>
<div class="card">
  <div class="queue" id="queue"><div class="empty">队列为空 — 先添加文件</div></div>
  <div class="bar">
    <div class="progress" id="prog"></div>
    <div style="display:flex;gap:8px">
      <button class="btn ghost" id="clear" disabled>清空</button>
      <button class="btn" id="start" disabled>开始转换</button>
    </div>
  </div>
</div>
<div id="mask"><div id="help">
  <h2>使用帮助 <button class="x" id="helpx">✕</button></h2>
  <h3>基本流程</h3>
  <p>1. 点选或拖入 PDF（可多选, 最多 10 个）→ 2. 加密的文件在行内密码框填入打开密码 → 3. 点"开始转换", 逐个排队转换 → 4. 完成后点该行的"下载 xlsx"保存到手机。</p>
  <h3>密码说明</h3>
  <p>加密 PDF（如华夏银行）在该文件行内的密码框填写打开密码, 未填则按无密码转换。密码只保存在你自己的浏览器里、随转换使用, 服务器不做任何存储。</p>
  <h3>队列与刷新</h3>
  <p>队列自动保存在本机浏览器（IndexedDB）, 刷新页面不丢失: 待转换的文件、密码、失败原因、已完成的结果都会恢复, 可重新下载。点每行 ✕ 可单独移除。</p>
  <h3>失败怎么办</h3>
  <p>失败的文件会显示原因和建议, 可点"重试"重新转换; 网络中断类失败会自动重试一次。若是新格式识别失败, 点"下载诊断包"（脱敏 JSON, 不含真实数据）发给维护者, 下个版本即可支持。</p>
  <p style="margin-top:6px">提交后即使手机锁屏/切后台/网络闪断, 转换仍会在服务器继续完成——回到本页稍等即可收取结果（结果保留 6 小时）。若某浏览器反复报网络错误, 建议换一个浏览器打开。</p>
  <div class="hl">__PRIVACY_NOTE__</div>
</div></div>
<script>
const MAXQ=10;
const drop=document.getElementById('drop'),hint=document.getElementById('hint'),f=document.getElementById('f'),
      queueEl=document.getElementById('queue'),prog=document.getElementById('prog'),
      startBtn=document.getElementById('start'),clearBtn=document.getElementById('clear');
const HINT_DEFAULT='可多选, 最多 10 个 · 单个 ≤ 200MB';
let items=[];  // {file, pwd, status, msg, blob, detail}
let running=false;

// ---- IndexedDB 持久化 ----
let _dbPromise=null;
function idb(){
  if(!_dbPromise){
    _dbPromise=new Promise((res,rej)=>{
      const r=indexedDB.open('b2xq',1);
      r.onupgradeneeded=()=>r.result.createObjectStore('items',{keyPath:'id'});
      r.onsuccess=()=>res(r.result);
      r.onerror=()=>rej(r.error);
      setTimeout(()=>rej(new Error('idb open timeout')),2500);
    }).catch(e=>{_dbPromise=null;throw e});
  }
  return _dbPromise;
}
async function persist(){
  try{
    const db=await idb();
    const tx=db.transaction('items','readwrite');
    tx.objectStore('items').clear();
    items.forEach((it,i)=>tx.objectStore('items').put({
      id:i, file:it.file, pwd:it.pwd||'', status:it.status,
      msg:it.msg||'', blob:it.blob||null, detail:it.detail||null }));
  }catch(e){}
}
async function restore(){
  try{
    const db=await Promise.race([idb(),
      new Promise((_,rej)=>setTimeout(()=>rej(new Error('idb timeout')),2500))]);
    const rq=db.transaction('items','readonly').objectStore('items').getAll();
    rq.onsuccess=()=>{
      items=(rq.result||[]).sort((a,b)=>a.id-b.id).map(x=>({
        file:x.file, pwd:x.pwd||'',
        status:x.status==='run'?'pending':x.status,
        msg:x.msg, blob:x.blob||null, detail:x.detail||null,
      })).map(x=>({...x, blobUrl:x.blob?URL.createObjectURL(x.blob):null}));
      render();
    };
    rq.onerror=()=>render();
  }catch(e){render()}
}
// ---- 帮助弹窗 ----
const mask=document.getElementById('mask');
document.getElementById('helpbtn').onclick=()=>mask.classList.add('show');
document.getElementById('helpx').onclick=()=>mask.classList.remove('show');
mask.addEventListener('click',e=>{if(e.target===mask)mask.classList.remove('show')});
// ---- 渲染 ----
function fmtSize(n){return n>1048576?(n/1048576).toFixed(1)+' MB':Math.max(1,n/1024).toFixed(0)+' KB'}
function flashHint(msg){
  hint.textContent=msg;hint.classList.add('warn');
  setTimeout(()=>{hint.textContent=HINT_DEFAULT;hint.classList.remove('warn')},4000);
}
function render(){
  queueEl.innerHTML = items.length ? '' : '<div class="empty">队列为空 — 先添加文件</div>';
  items.forEach((it,i)=>{
    const d=document.createElement('div');d.className='item';
    const badge={pending:'等待',run:'转换中',ok:'完成',fail:'失败'}[it.status];
    let msg=it.msg||fmtSize(it.file.size);
    let actions='';
    if(it.status==='ok'&&it.blobUrl)
      actions='<button class="linkbtn" data-dl="'+i+'">下载 xlsx</button>';
    if(it.status==='fail')
      actions='<button class="linkbtn" data-retry="'+i+'">重试</button>'+
        (it.detail&&it.detail.diag?'<button class="linkbtn" data-diag="'+i+'">下载诊断包</button>':'');
    const lock=it.pwd?'🔒':'';
    d.innerHTML='<div class="ficon">PDF</div><div class="fmeta"><div class="fname">'+
      escapeHtml(it.file.name)+' '+lock+'</div><div class="fmsg '+(it.status==='fail'?'err':it.status==='ok'?'ok':'')+'">'+
      escapeHtml(msg)+'</div>'+(it.status==='fail'&&it.detail&&it.detail.suggestion?'<div class="err-detail">建议: '+escapeHtml(it.detail.suggestion)+'</div>':'')+
      '</div><input type="password" class="pin" data-pwd="'+i+'" placeholder="无密码" value="'+escapeHtml(it.pwd||'')+'" '+(running?'disabled':'')+
      '><span class="badge '+it.status+'">'+badge+'</span>'+(actions||'')+
      '<button class="linkbtn del" data-del="'+i+'" title="从队列移除" '+(it.status==='run'?'disabled':'')+'>✕</button>';
    queueEl.appendChild(d);
  });
  const done=items.filter(x=>x.status==='ok'||x.status==='fail').length;
  prog.textContent=items.length?('进度 '+done+' / '+items.length):'';
  startBtn.disabled=!items.some(x=>x.status==='pending')||running;
  clearBtn.disabled=!items.length||running;
  queueEl.querySelectorAll('[data-pwd]').forEach(inp=>inp.oninput=e=>{items[+e.target.dataset.pwd].pwd=e.target.value;persist()});
  queueEl.querySelectorAll('[data-dl]').forEach(b=>b.onclick=()=>doDownload(items[+b.dataset.dl]));
  queueEl.querySelectorAll('[data-diag]').forEach(b=>b.onclick=()=>doDiag(items[+b.dataset.diag]));
  queueEl.querySelectorAll('[data-retry]').forEach(b=>b.onclick=()=>{
    const it=items[+b.dataset.retry];
    it.status='pending';it.msg=null;it._retried=false;persist();render();
    if(!running)startAll();
  });
  queueEl.querySelectorAll('[data-del]').forEach(b=>b.onclick=()=>removeAt(+b.dataset.del));
}
function removeAt(i){
  if(items[i].status==='run')return;
  if(items[i].blobUrl)URL.revokeObjectURL(items[i].blobUrl);
  items.splice(i,1);persist();render();
}
function escapeHtml(s){return String(s||'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]))}
function addFiles(files){
  let added=0,skipped=0;
  for(const f of files){
    if(items.length>=MAXQ){skipped++;continue}
    if(!/\.pdf$/i.test(f.name))continue;
    if(items.some(x=>x.file.name===f.name&&x.file.size===f.size))continue;
    items.push({file:f,pwd:'',status:'pending',msg:null,blob:null,blobUrl:null,detail:null});
    added++;
  }
  if(skipped)flashHint('队列上限 '+MAXQ+' 个, 已跳过 '+skipped+' 个');
  if(added)persist();
  render();
}
function doDownload(it){
  if(!it.blobUrl)it.blobUrl=URL.createObjectURL(it.blob);
  const a=document.createElement('a');a.href=it.blobUrl;
  a.download=it.file.name.replace(/\.pdf$/i,'')+'.xlsx';a.click();
}
function doDiag(it){
  const blob=new Blob([JSON.stringify(it.detail.diag,null,1)],{type:'application/json'});
  const a=document.createElement('a');a.href=URL.createObjectURL(blob);
  a.download='bank2excel-diagnosis-'+Date.now()+'.json';a.click();
  setTimeout(()=>URL.revokeObjectURL(a.href),30000);
}
async function convertOne(it){
  it.status='run';it.msg='上传中…';it._retried=it._retried||false;render();
  const fd=new FormData();fd.append('file',it.file);fd.append('password',it.pwd||'');
  const t0=performance.now();
  let pollFails=0;
  try{
    // 提交任务: 上传完成后连接即可断开, 服务端后台继续转换(移动端友好)
    const up=await fetch('/api/tasks',{method:'POST',body:fd});
    if(!up.ok){
      const j=await up.json().catch(()=>({}));
      const d=(j.detail&&typeof j.detail==='object')?j.detail:{message:String(j.detail||('HTTP '+up.status))};
      it.status='fail';it.detail=d;it.msg=d.message||('HTTP '+up.status);
      persist();render();return;
    }
    const tid=(await up.json()).task_id;
    // 轮询结果(容忍最多 6 次瞬断; 服务端不受影响继续转换)
    while(true){
      await new Promise(r=>setTimeout(r,2500));
      let s=null;
      try{
        s=await fetch('/api/tasks/'+tid).then(r=>r.json());
        pollFails=0;
      }catch(pe){
        if(++pollFails>=6)throw pe;
        continue;
      }
      if(s.status==='queued'||s.status==='running'){
        const sec=Math.round((performance.now()-t0)/1000);
        it.msg='转换中… '+sec+'s'+(s.ocr_pages?' · OCR '+s.ocr_pages+' 页':'');
        render();continue;
      }
      if(s.status==='ok'){
        const rb=await fetch('/api/tasks/'+tid+'/result');
        it.blob=await rb.blob();
        it.blobUrl=URL.createObjectURL(it.blob);
        it.status='ok';
        const warn=s.warning;
        it.msg=fmtSize(it.blob.size)+' · '+s.duration_s+'s'+
          (s.ocr_pages?' · OCR '+s.ocr_pages+' 页':'')+(warn?' · '+warn:'');
        break;
      }
      // fail
      it.status='fail';
      it.detail={message:s.error||'转换失败',stage:s.stage};
      it.msg=s.error||'转换失败';
      break;
    }
  }catch(e){
    const netErr=(e instanceof TypeError)||/fetch|network|Failed|abort/i.test(e.message||'');
    if(!it._retried&&netErr){
      it._retried=true;it.status='pending';it.msg='网络中断, 5 秒后自动重试…';
      render();await new Promise(r=>setTimeout(r,5000));
      return convertOne(it);
    }
    it.status='fail';it.detail={message:e.message};it.msg='网络错误: '+e.message;
  }
  persist();render();
}
async function startAll(){
  running=true;items.forEach(x=>{if(x.status==='pending')x._retried=false});render();
  let it;
  while((it=items.find(x=>x.status==='pending'))){
    await convertOne(it);
  }
  running=false;persist();render();
}
drop.onclick=()=>f.click();
['dragover','dragenter'].forEach(e=>drop.addEventListener(e,x=>{x.preventDefault();drop.classList.add('over')}));
['dragleave','drop'].forEach(e=>drop.addEventListener(e,x=>{x.preventDefault();drop.classList.remove('over')}));
drop.addEventListener('drop',e=>{if(e.dataTransfer.files.length)addFiles(e.dataTransfer.files)});
f.addEventListener('change',()=>{addFiles(f.files);f.value=''});
startBtn.onclick=startAll;
clearBtn.onclick=()=>{
  items.forEach(x=>{if(x.blobUrl)URL.revokeObjectURL(x.blobUrl)});
  items=[];persist();render();
};
restore();
</script></div></body></html>"""


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
