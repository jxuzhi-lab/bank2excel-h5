# -*- coding: utf-8 -*-
"""test_vps_fallback.py —— T1 服务端未知格式兜底链路端到端冒烟(2026-08-29)。

不依赖真实 VLM: 本脚本自建一个 OpenAI 兼容假视觉端点, 并生成一份
"表头关键词零命中"的未知格式合成 PDF, 对 server.py 起的服务验证:

  A. 已知格式样本 → 200 + xlsx;
  B. 未知格式 + 无视觉(provider=none) → 422 结构化错误(stage/suggestion/diag);
  C. 未知格式 + api provider → 升级链自动 onboarding → 200 + xlsx;
  D. 同文件再传一次 → 命中描述符指纹缓存, 不再发起视觉请求;
  E. api 请求形态校验(Authorization/model/image_url)。

用法(项目 venv):
    python tests/test_vps_fallback.py
仅用标准库 + pymupdf + fastapi/uvicorn(服务端既有依赖), 无新增依赖。
"""
import json
import os
import subprocess
import sys
import tempfile
import threading
import time
import urllib.request
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

import pymupdf

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
SERVER_PY = os.path.join(ROOT, "server.py")
SAMPLES = os.path.join(os.path.dirname(ROOT), "测试样本")

FAKE_PORT = 18765
COLUMNS_ANSWER = "记账日期,交易说明,资金流入,资金流出,剩余金额"


# ---------- 假 OpenAI 兼容视觉端点 ----------
class FakeVLM(BaseHTTPRequestHandler):
    requests_seen = []          # [(path, model, has_bearer, has_image, prompt_head)]
    header_ask_count = 0        # "读表头列名"类请求计数(用于缓存命中断言)

    def log_message(self, *a):  # 静默
        pass

    def do_POST(self):
        body = json.loads(self.rfile.read(int(self.headers["Content-Length"])))
        auth = self.headers.get("Authorization", "")
        msgs = body.get("messages", [])
        content = msgs[0].get("content", []) if msgs else []
        prompt = next((c.get("text", "") for c in content if c.get("type") == "text"), "")
        has_image = any(c.get("type") == "image_url" for c in content)
        FakeVLM.requests_seen.append(
            (self.path, body.get("model"), auth.startswith("Bearer "), has_image, prompt[:24]))
        if "表头列名" in prompt:  # 仅"读表头列名"类请求(classify 的选项文案含"列名", 勿误计)
            FakeVLM.header_ask_count += 1
            answer = COLUMNS_ANSWER
        elif "类型" in prompt:
            answer = "A, 表格页"
        elif "日期" in prompt:
            answer = "2025-01-03"
        else:
            answer = "A, 表格页"
        resp = json.dumps({
            "choices": [{"message": {"content": answer}}],
        }).encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(resp)))
        self.end_headers()
        self.wfile.write(resp)


def start_fake_vlm():
    srv = ThreadingHTTPServer(("127.0.0.1", FAKE_PORT), FakeVLM)
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    return srv


# ---------- 未知格式合成 PDF(表头关键词零命中) ----------
def make_unknown_pdf(path):
    doc = pymupdf.open()
    page = doc.new_page()  # A4 595x842
    xs = {"date": 50, "memo": 120, "in": 260, "out": 340, "bal": 440}
    for name, x in zip(("记账日期", "交易说明", "资金流入", "资金流出", "剩余金额"),
                       (xs["date"], xs["memo"], xs["in"], xs["out"], xs["bal"])):
        page.insert_text((x, 70), name, fontname="china-s", fontsize=10)
    bal = 10000.00
    for i in range(15):
        y = 95 + i * 16
        day = f"2025-01-{i + 1:02d}"
        page.insert_text((xs["date"], y), day, fontname="helv", fontsize=9)
        page.insert_text((xs["memo"], y), f"购物事项{i:02d}号", fontname="china-s", fontsize=9)
        if i % 2 == 0:
            amt = 100.00 + i
            bal += amt
            page.insert_text((xs["in"], y), f"{amt:,.2f}", fontname="helv", fontsize=9)
        else:
            amt = 50.00 + i
            bal -= amt
            page.insert_text((xs["out"], y), f"{amt:,.2f}", fontname="helv", fontsize=9)
        page.insert_text((xs["bal"], y), f"{bal:,.2f}", fontname="helv", fontsize=9)
    doc.save(path)
    doc.close()


# ---------- HTTP 小工具 ----------
def http_post(url, data, headers=None, timeout=180):
    req = urllib.request.Request(url, data=data,
                                 headers=headers or {"Content-Type": "application/json"},
                                 method="POST")
    try:
        with urllib.request.urlopen(req, timeout=timeout) as r:
            return r.status, r.read()
    except urllib.error.HTTPError as e:
        return e.code, e.read()


def http_get(url, timeout=15):
    with urllib.request.urlopen(url, timeout=timeout) as r:
        return r.status, json.loads(r.read().decode("utf-8"))


def start_server(port, extra_env):
    env = dict(os.environ)
    env.pop("PYODIDE", None)
    env["PYTHONIOENCODING"] = "utf-8"
    env.update(extra_env)
    proc = subprocess.Popen(
        [sys.executable, SERVER_PY, "--host", "127.0.0.1", "--port", str(port)],
        env=env, cwd=ROOT,
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    base = f"http://127.0.0.1:{port}"
    for _ in range(60):
        try:
            http_get(base + "/api/health")
            return proc, base
        except Exception:
            time.sleep(0.5)
    raise RuntimeError(f"server:{port} 启动超时")


def pick_sample():
    """挑一份最小的已知格式样本。"""
    cands = sorted(
        (f for f in os.listdir(SAMPLES)
         if f.endswith(".pdf") and os.path.getsize(os.path.join(SAMPLES, f)) < 300_000),
        key=lambda f: os.path.getsize(os.path.join(SAMPLES, f)))
    if not cands:
        raise RuntimeError("测试样本目录无可用 PDF")
    return os.path.join(SAMPLES, cands[0])


def main():
    tmp = tempfile.mkdtemp(prefix="b2x_t1_")
    unknown_pdf = os.path.join(tmp, "unknown.pdf")
    make_unknown_pdf(unknown_pdf)
    known_pdf = pick_sample()
    desc_dir = os.path.join(tmp, "desc")
    print(f"[setup] 未知格式合成 PDF: {unknown_pdf}")
    print(f"[setup] 已知格式样本: {os.path.basename(known_pdf)}")

    fake = start_fake_vlm()
    api_env = {
        "BANK_PDF_VISION_PROVIDER": "api",
        "BANK_PDF_VISION_API_BASE": f"http://127.0.0.1:{FAKE_PORT}",
        "BANK_PDF_VISION_API_KEY": "test-key-123",
        "BANK_PDF_VISION_API_MODEL": "test-vlm",
        "DESCRIPTOR_CACHE_DIR": desc_dir,
        "VLM_BUDGET_PER_HOUR": "40",
    }
    none_env = {"DESCRIPTOR_CACHE_DIR": os.path.join(tmp, "desc_none"),
                # 显式禁用: 开发机可能残留可用的 vision.js, 会干扰"无视觉"场景
                "BANK_PDF_VISION_PROVIDER": "none"}

    results = []
    try:
        # --- api provider 实例 ---
        proc_a, base_a = start_server(18766, api_env)
        # --- provider=none 实例 ---
        proc_n, base_n = start_server(18767, none_env)

        # 0) health: provider 解析
        _, h = http_get(base_a + "/api/health")
        ok0 = h["vision_provider"] == "api" and h["engine"].startswith("full")
        results.append(("health: api provider + 完整引擎", ok0, h))
        _, h2 = http_get(base_n + "/api/health")
        ok0b = h2["vision_provider"] == "none"
        results.append(("health: 无配置时自动降级 none", ok0b, h2["vision_provider"]))

        # A) 已知格式 → 200
        with open(known_pdf, "rb") as f:
            st, body = http_post(base_a + "/api/convert", f.read(),
                                 {"Content-Type": "multipart/form-data; boundary=X"})
        # urllib 不自动构造 multipart; 改用手工 multipart:
        st, body = post_multipart(base_a + "/api/convert", known_pdf)
        okA = st == 200 and body[:2] == b"PK"
        results.append(("A 已知格式 → 200 xlsx", okA, (st, len(body))))

        # B) 未知格式 + none → 422 结构化
        st, body = post_multipart(base_n + "/api/convert", unknown_pdf)
        detail = {}
        try:
            detail = json.loads(body.decode("utf-8")).get("detail", {})
        except Exception:
            pass
        okB = st == 422 and isinstance(detail, dict) and detail.get("stage") \
            and detail.get("suggestion")
        results.append(("B 未知格式+无视觉 → 422 结构化(stage/suggestion/diag)",
                        okB, (st, detail.get("stage"), detail.get("message", "")[:40])))

        # C) 未知格式 + api → 升级链兜底成功
        st, body = post_multipart(base_a + "/api/convert", unknown_pdf)
        okC = st == 200 and body[:2] == b"PK"
        desc_files = os.listdir(desc_dir) if os.path.isdir(desc_dir) else []
        results.append(("C 未知格式+api → VLM 兜底转换成功", okC, (st, len(body))))
        results.append(("C2 描述符提取验证后已入缓存", len(desc_files) == 1, desc_files))

        # D) 同文件再传 → 命中指纹缓存, 不再发起"读表头"视觉请求
        asks_before = FakeVLM.header_ask_count
        st, body = post_multipart(base_a + "/api/convert", unknown_pdf)
        asks_after = FakeVLM.header_ask_count
        okD = st == 200 and asks_after == asks_before
        results.append(("D 二次请求命中描述符缓存(零新增视觉请求)", okD,
                        (st, f"header_ask {asks_before}->{asks_after}")))

        # E) api 请求形态
        r = FakeVLM.requests_seen[0] if FakeVLM.requests_seen else None
        okE = bool(r) and r[0] == "/chat/completions" and r[1] == "test-vlm" \
            and r[2] and r[3]
        results.append(("E api 请求形态(path/model/Bearer/image)", okE, r[:4] if r else None))

        for proc in (proc_a, proc_n):
            proc.terminate()
    finally:
        fake.shutdown()
    # 不清理 tmp, 便于失败时人工检查

    print("\n================ 冒烟结果 ================")
    n_ok = 0
    for name, ok, info in results:
        print(f"{'PASS' if ok else 'FAIL'}  {name}   {info}")
        n_ok += bool(ok)
    print(f"---------- {n_ok}/{len(results)} ----------")
    return 0 if n_ok == len(results) else 1


def post_multipart(url, file_path, field="file", timeout=180):
    """手工 multipart/form-data(标准库, 不引 requests)。"""
    import mimetypes
    import uuid
    boundary = "----b2x" + uuid.uuid4().hex
    with open(file_path, "rb") as f:
        payload = f.read()
    fname = os.path.basename(file_path)
    ctype = mimetypes.guess_type(fname)[0] or "application/pdf"
    body = (
        f"--{boundary}\r\n"
        f'Content-Disposition: form-data; name="{field}"; filename="{fname}"\r\n'
        f"Content-Type: {ctype}\r\n\r\n"
    ).encode("utf-8") + payload + f"\r\n--{boundary}--\r\n".encode("utf-8")
    return http_post(url, body, {"Content-Type": f"multipart/form-data; boundary={boundary}"},
                     timeout=timeout)


if __name__ == "__main__":
    sys.exit(main())
