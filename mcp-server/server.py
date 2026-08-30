#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""bank2excel MCP server —— 让大模型通过 MCP 调用对账单转换服务。

工具:
  - convert_pdf            本地 PDF → Excel(走私有转换服务, 支持 OCR/加密/密码)
  - conversion_logs        查询最近的转换日志(需 B2X_ADMIN_PASSWORD)
  - conversion_stats       转换量与 API 消耗统计
  - service_health         服务健康状态

配置(环境变量):
  B2X_API_URL          服务地址, 默认线上 VPS http://124.223.110.222
  B2X_ADMIN_PASSWORD   日志后台密码(仅日志/统计工具需要)
  B2X_TIMEOUT          转换请求超时秒数, 默认 900(扫描件 OCR 较慢)

注册(ZCode 用户级 ~/.zcode/cli/config.json):
  "mcp": {"servers": {"bank2excel": {
      "command": "<python>", "args": ["<本目录>/server.py"],
      "env": {"B2X_API_URL": "...", "B2X_ADMIN_PASSWORD": "..."}}}}
依赖: pip install mcp
"""
import json
import mimetypes
import os
import urllib.error
import urllib.request
import uuid

from mcp.server.fastmcp import FastMCP

API_URL = os.environ.get("B2X_API_URL", "http://124.223.110.222").rstrip("/")
ADMIN_PASSWORD = os.environ.get("B2X_ADMIN_PASSWORD", "")
TIMEOUT = int(os.environ.get("B2X_TIMEOUT", "900"))

mcp = FastMCP("bank2excel")


def _post_convert(pdf_path, password):
    boundary = "----b2x" + uuid.uuid4().hex
    with open(pdf_path, "rb") as f:
        payload = f.read()
    fname = os.path.basename(pdf_path)
    body = (
        f"--{boundary}\r\n"
        f'Content-Disposition: form-data; name="file"; filename="{fname}"\r\n'
        f"Content-Type: application/pdf\r\n\r\n"
    ).encode("utf-8") + payload + (
        f"\r\n--{boundary}\r\n"
        f'Content-Disposition: form-data; name="password"\r\n\r\n{password}\r\n'
        f"--{boundary}--\r\n"
    ).encode("utf-8")
    req = urllib.request.Request(
        API_URL + "/api/convert", data=body, method="POST",
        headers={"Content-Type": f"multipart/form-data; boundary={boundary}"})
    try:
        with urllib.request.urlopen(req, timeout=TIMEOUT) as r:
            return r.status, r.read(), dict(r.headers)
    except urllib.error.HTTPError as e:
        return e.code, e.read(), dict(e.headers)


def _get_json(path, timeout=30):
    req = urllib.request.Request(API_URL + path, method="GET")
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.status, json.loads(r.read().decode("utf-8", "replace"))


@mcp.tool()
def convert_pdf(file_path: str, password: str = "", output_dir: str = "") -> str:
    """把本地银行对账单/微信支付流水的 PDF 转换成 Excel(xlsx)。

    走私有转换服务: 自动识别 13+ 家银行格式; 扫描件自动 OCR;
    未见过的新格式自动视觉学习; 加密 PDF 通过 password 传打开密码。

    Args:
        file_path: PDF 文件的本地路径(必填)
        password: PDF 打开密码(加密文件才需要, 默认空)
        output_dir: 输出目录, 默认与源文件同目录
    """
    file_path = os.path.abspath(os.path.expanduser(file_path))
    if not os.path.isfile(file_path):
        return json.dumps({"ok": False, "error": f"文件不存在: {file_path}"},
                          ensure_ascii=False)
    if not file_path.lower().endswith(".pdf"):
        return json.dumps({"ok": False, "error": "仅支持 .pdf 文件"}, ensure_ascii=False)
    status, body, headers = _post_convert(file_path, password)
    if status != 200:
        try:
            detail = json.loads(body.decode("utf-8", "replace")).get("detail", {})
        except Exception:  # noqa: BLE001
            detail = {"message": body.decode("utf-8", "replace")[:300]}
        if isinstance(detail, str):
            detail = {"message": detail}
        return json.dumps({
            "ok": False, "http": status,
            "error": detail.get("message", "转换失败"),
            "stage": detail.get("stage"),
            "suggestion": detail.get("suggestion"),
        }, ensure_ascii=False)
    out_dir = os.path.abspath(os.path.expanduser(output_dir)) if output_dir \
        else os.path.dirname(file_path)
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(
        out_dir, os.path.splitext(os.path.basename(file_path))[0] + ".xlsx")
    with open(out_path, "wb") as f:
        f.write(body)
    return json.dumps({
        "ok": True,
        "output_path": out_path,
        "xlsx_bytes": len(body),
        "ocr_pages": headers.get("X-OCR-Pages"),
        "ocr_mode": headers.get("X-OCR-Mode"),
        "warning": headers.get("X-OCR-Warning"),
        "note": "扫描件转换请人工核对 warning 中的余额链断点提示",
    }, ensure_ascii=False)


@mcp.tool()
def conversion_logs(days: int = 7, limit: int = 20, status: str = "") -> str:
    """查询最近的转换日志(时间/文件/状态/耗时/通道/各 API 消耗)。
    需要 B2X_ADMIN_PASSWORD 环境变量。

    Args:
        days: 查询最近几天, 默认 7(最大 30, 服务端只保留 30 天)
        limit: 返回条数上限, 默认 20
        status: 过滤 "ok"/"fail", 空为全部
    """
    if not ADMIN_PASSWORD:
        return json.dumps({"ok": False,
                           "error": "未配置 B2X_ADMIN_PASSWORD, 无法查询日志"},
                          ensure_ascii=False)
    import http.cookiejar
    import urllib.parse
    login_data = urllib.parse.urlencode({"password": ADMIN_PASSWORD}).encode()
    cj = http.cookiejar.CookieJar()
    opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(cj))
    req = urllib.request.Request(API_URL + "/admin/login", data=login_data,
                                 method="POST")
    with opener.open(req, timeout=30) as r:
        pass
    qs = urllib.parse.urlencode({"days": max(1, min(days, 30)),
                                 "limit": limit, "status": status})
    req2 = urllib.request.Request(API_URL + "/admin/api/logs?" + qs)
    with opener.open(req2, timeout=30) as r:
        logs = json.loads(r.read().decode("utf-8", "replace"))
    slim = [{k: l.get(k) for k in ("ts", "filename", "status", "stage", "mode",
                                   "duration_s", "rows", "glm_calls",
                                   "vlm_calls", "baidu_calls", "error")}
            for l in logs]
    return json.dumps({"ok": True, "count": len(slim), "logs": slim},
                      ensure_ascii=False)


@mcp.tool()
def conversion_stats(days: int = 7) -> str:
    """转换服务的汇总统计: 总量/成功率/耗时/各 API 调用与 token 消耗。
    需要 B2X_ADMIN_PASSWORD 环境变量。

    Args:
        days: 统计最近几天, 默认 7
    """
    if not ADMIN_PASSWORD:
        return json.dumps({"ok": False,
                           "error": "未配置 B2X_ADMIN_PASSWORD, 无法查询统计"},
                          ensure_ascii=False)
    import http.cookiejar
    import urllib.parse
    login_data = urllib.parse.urlencode({"password": ADMIN_PASSWORD}).encode()
    cj = http.cookiejar.CookieJar()
    opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(cj))
    req = urllib.request.Request(API_URL + "/admin/login", data=login_data,
                                 method="POST")
    with opener.open(req, timeout=30):
        pass
    req2 = urllib.request.Request(
        API_URL + "/admin/api/summary?" +
        urllib.parse.urlencode({"days": max(1, min(days, 30))}))
    with opener.open(req2, timeout=30) as r:
        stats = json.loads(r.read().decode("utf-8", "replace"))
    return json.dumps({"ok": True, "days": days, "stats": stats},
                      ensure_ascii=False)


@mcp.tool()
def service_health() -> str:
    """检查转换服务的健康状态(引擎模式/视觉通道/OCR 开关/描述符缓存量)。"""
    try:
        status, health = _get_json("/api/health")
        return json.dumps({"ok": status == 200, "health": health},
                          ensure_ascii=False)
    except Exception as e:  # noqa: BLE001
        return json.dumps({"ok": False, "error": str(e)}, ensure_ascii=False)


if __name__ == "__main__":
    mcp.run()
