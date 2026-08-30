#!/usr/bin/env python3
"""视觉模型封装(2026-08-19 升级): 统一调用 claude-vision-skill/vision.js。

职责:
  - 渲染 PDF 页 → PNG(供视觉模型阅读);
  - 调用 node vision.js 提问, 失败自动重试 1 次;
  - 输出解析: 按约定取"最后一个含逗号且含中文的行"(列名等逗号分隔答案),
    其余问答类型提供专用解析函数;
  - 绝不打印/暴露 API key(.env 由 vision.js 自行加载)。

视觉 provider(2026-08-22): 默认 "visionjs"; 若接入的是原生多模态模型(可自行读图),
设 BANK_PDF_VISION_PROVIDER=model 即可让脚本把"读图"这一环交给当前模型——
渲染页保留为 PNG → 抛 ModelVisionPendingError 并写入待读图请求(含回答文件路径),
代理用模型看图后把答案写入回答文件, 重跑同一命令即命中答案继续; 全程不调外部
vision.js。见 SKILL.md「视觉角色: 原生多模态模型替代 vision.js」。

视觉原则(与 SKILL.md 一致): 视觉只做"每格式一次/每文件一次/异常按需"的高层
判断, 不做逐页逐格处理; 全量执行仍由规则管道负责。
"""
import base64
import hashlib
import json
import os
import re
import shutil
import subprocess
import tempfile
import urllib.request

import pymupdf

# vision.js 绝对路径。跨平台部署(2026-08-19): 优先读环境变量
# BANK_PDF_VISION_JS(在其他平台把 claude-vision-skill 的 vision.js 路径
# 指过来即可), 缺省回退本机固定路径(Windows 开发机)。
VISION_JS = os.environ.get(
    "BANK_PDF_VISION_JS",
    os.path.expanduser(r"~/.codex/skills/claude-vision-skill/vision.js"),
)
_VISION_ENV_BASE = os.path.dirname(VISION_JS)

VISION_PROVIDER_VAR = "BANK_PDF_VISION_PROVIDER"
MODEL_VISION_DIR_VAR = "BANK_PDF_MODEL_VISION_DIR"

# ---- api provider(2026-08-29): 服务端直连 OpenAI 兼容视觉模型 ----
# 供私有 VPS 服务端使用: 浏览器端(Pyodide)永远短路视觉, 不受影响。
# 配置(环境变量):
#   BANK_PDF_VISION_PROVIDER=api
#   BANK_PDF_VISION_API_BASE  如 https://open.bigmodel.cn/api/paas/v4 (不带 /chat/completions)
#   BANK_PDF_VISION_API_KEY   API Key
#   BANK_PDF_VISION_API_MODEL 模型名(默认 glm-4.6v)
VISION_API_BASE_VAR = "BANK_PDF_VISION_API_BASE"
VISION_API_KEY_VAR = "BANK_PDF_VISION_API_KEY"
VISION_API_MODEL_VAR = "BANK_PDF_VISION_API_MODEL"


def _api_config():
    """读取 api provider 配置, 返回 (base, key, model)。绝不打印 key。"""
    base = os.environ.get(VISION_API_BASE_VAR, "").strip().rstrip("/")
    key = os.environ.get(VISION_API_KEY_VAR, "").strip()
    model = os.environ.get(VISION_API_MODEL_VAR, "glm-4.6v").strip()
    return base, key, model


_LAST_API_USAGE = __import__("threading").local()  # 每请求线程累计 api 视觉 token(服务端日志用)


class VisionDisabledError(RuntimeError):
    """当前环境未配置任何可用的视觉能力(无 vision.js/node 或显式禁用)。"""


def _visionjs_available():
    """外部 vision.js 是否可运行(路径存在且能找到 node)。"""
    if not os.path.isfile(VISION_JS):
        return False
    if shutil.which("node"):
        return True
    for cand in (r"C:/Program Files/nodejs/node.exe",
                 r"C:/Program Files (x86)/nodejs/node.exe"):
        if cand and os.path.isfile(cand):
            return True
    return False


def resolve_vision_provider(provider=None):
    """把配置解析为实际 provider。

  - "none"/"off"/"disabled"/"0"/"false": 无视觉, 仅规则/文字层;
  - "model": 原生多模态模型回环(写待读图请求);
  - "visionjs": 外部 vision.js;
  - "api": 服务端直连 OpenAI 兼容视觉模型(需配置 API_BASE/API_KEY);
  - "auto"/未设置: 显式环境变量优先; 否则外部 vision.js 可用则用,
    api 配置齐全则用 api, 都不可用自动降级 none。
  无图像能力环境绝不会卡在外部视觉调用或模型读图待办上。
  """
    if provider is None:
        raw = os.environ.get(VISION_PROVIDER_VAR, "auto")
    else:
        raw = provider
    raw = str(raw or "auto").strip().lower()
    # 显式传 "auto" 时仍尊重已设置的环境变量(none/model/visionjs/api), 便于 CI/进程级降级。
    if raw == "auto":
        env_raw = str(os.environ.get(VISION_PROVIDER_VAR, "auto")).strip().lower()
        if env_raw in ("none", "model", "visionjs", "api"):
            raw = env_raw
    if raw in ("none", "off", "disabled", "0", "false", "no"):
        return "none"
    if raw in ("model", "visionjs", "api"):
        return raw
    if raw in ("auto", "", "default"):
        _ab, _ak, _ = _api_config()
        if _visionjs_available():
            return "visionjs"
        if _ab and _ak:
            return "api"
        return "none"
    return "none"


def set_vision_provider(provider):
    """解析并写入环境变量, 返回实际 provider。"""
    resolved = resolve_vision_provider(provider)
    os.environ[VISION_PROVIDER_VAR] = resolved
    return resolved


def get_vision_provider():
    """返回实际 provider: visionjs / model / none(未配置或不可用时不再默认外部视觉)。"""
    return resolve_vision_provider(os.environ.get(VISION_PROVIDER_VAR, "auto"))


class ModelVisionPendingError(Exception):
    """原生多模态模型回环待办: 脚本需要当前模型读图, 已写出请求并暂存 PNG。

    属性:
      png           待读图 PNG 绝对路径(已复制到请求目录, 不会被调用方清理);
      prompt        要问模型的问题;
      answer_file   代理应将模型答案写入的文本文件(原始文本, 可多行);
      request_dir   请求目录(含 request.json / page.png)。

    处理方式: 代理用模型读 png 后, 把答案写入 answer_file (遵循 _last_zh_line
    约定: 列名答案放"最后一个含逗号且含中文的行"), 然后重跑同一命令即命中答案。
    """

    def __init__(self, png, prompt, answer_file, request_dir):
        super().__init__(f"需要当前模型读图: {png}")
        self.png = png
        self.prompt = prompt
        self.answer_file = answer_file
        self.request_dir = request_dir


def _node_cmd():
    node = shutil.which("node")
    if not node:
        # 常见安装位置兜底
        for cand in (r"C:/Program Files/nodejs/node.exe",
                     r"C:/Program Files (x86)/nodejs/node.exe"):
            if cand and os.path.isfile(cand):
                node = cand
                break
    if not node:
        raise RuntimeError("未找到 node, 无法调用 vision.js 视觉模型")
    return node


def _model_vision_call(png_path, prompt, retries=1, timeout=150):
    """原生多模态模型回环: 不调外部 vision.js。

    1) 把渲染页复制到请求目录并保留(调用方 finally 清理的是原临时 PNG, 请求副本不删);
    2) 写 request.json(含 png/prompt/answer_file/约定);
    3) 抛 ModelVisionPendingError 供上层感知"需要当前模型读图"。

    回答文件命中(代理已写入)则直接返回原文, 相当于 vision.js 存在的效果。
    """
    if not os.path.exists(png_path):
        raise RuntimeError(f"渲染 PNG 不存在: {png_path}")
    # 请求目录按 (prompt + 页面内容) 哈希, 保证"同一文件同一页同一问题"落同一目录:
    # 代理写答案后重跑可命中; 不同文件/页互不串扰(内容不同 → 目录不同)。
    with open(png_path, "rb") as f:
        digest = hashlib.sha256((prompt.strip() or "_").encode("utf-8")
                                + f.read()).hexdigest()[:20]
    base = os.environ.get(
        MODEL_VISION_DIR_VAR,
        os.path.join(tempfile.gettempdir(), "bs_model_vision"),
    )
    request_dir = os.path.join(base, digest)
    os.makedirs(request_dir, exist_ok=True)
    png_copy = os.path.join(request_dir, "page.png")
    answer_file = os.path.join(request_dir, "answer.txt")
    # 幂等: 同请求已有回答则直接返回(重跑命中), 不再抛待办。
    if os.path.exists(answer_file):
        with open(answer_file, encoding="utf-8") as f:
            body = f.read().strip()
        if body:
            return body
    shutil.copy2(png_path, png_copy)
    manifest = {
        "png": png_copy,
        "prompt": prompt,
        "answer_file": answer_file,
        "instruction": (
            "用当前原生多模态模型读 png 后, 把答案以文本形式写入 answer.txt。"
            "列名答案要求: 列出该表格表头列名, 用英文逗号分隔, 且放在最后一"
            "个含逗号且含中文的行; 不要输出解释/序号。"
        ),
    }
    with open(os.path.join(request_dir, "request.json"), "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)
    raise ModelVisionPendingError(
        png_copy, prompt, answer_file, request_dir)


def _api_vision_call(png_path, prompt, retries=1, timeout=120):
    """provider=api: 直连 OpenAI 兼容 /chat/completions 视觉模型。
    只用标准库 urllib, 服务端镜像无需新增依赖。返回原始回答文本。"""
    base, key, model = _api_config()
    if not base or not key:
        raise VisionDisabledError(
            f"api provider 配置不完整(需 {VISION_API_BASE_VAR} 和 {VISION_API_KEY_VAR})")
    if not os.path.exists(png_path):
        raise RuntimeError(f"渲染 PNG 不存在: {png_path}")
    with open(png_path, "rb") as f:
        b64 = base64.b64encode(f.read()).decode("ascii")
    payload = {
        "model": model,
        "temperature": 0,
        "messages": [{
            "role": "user",
            "content": [
                {"type": "image_url",
                 "image_url": {"url": "data:image/png;base64," + b64}},
                {"type": "text", "text": prompt},
            ],
        }],
    }
    url = base + "/chat/completions"
    last_err = None
    for _attempt in range(retries + 1):
        try:
            req = urllib.request.Request(
                url,
                data=json.dumps(payload).encode("utf-8"),
                headers={"Content-Type": "application/json",
                         "Authorization": "Bearer " + key},
                method="POST",
            )
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                body = json.loads(resp.read().decode("utf-8", "replace"))
            _u = body.get("usage") or {}
            _prev = getattr(_LAST_API_USAGE, "usage", None) or {"calls": 0, "prompt": 0, "completion": 0}
            _LAST_API_USAGE.usage = {"calls": _prev["calls"] + 1,
                                     "prompt": _prev["prompt"] + int(_u.get("prompt_tokens", 0)),
                                     "completion": _prev["completion"] + int(_u.get("completion_tokens", 0))}
            msg = (body.get("choices") or [{}])[0].get("message", {})
            out = (msg.get("content") or "").strip()
            if not out:
                raise RuntimeError(f"视觉 API 无输出: {str(body)[:200]}")
            return out
        except VisionDisabledError:
            raise
        except Exception as e:  # noqa: BLE001 (网络/超时/限流等偶发失败, 重试)
            last_err = str(e)
    raise RuntimeError(f"视觉 API 调用失败(重试 {retries} 次后): {last_err}")


def call_vision_raw(png_path, prompt, retries=1, timeout=150):
    """调用 vision.js 并返回原始 stdout 文本。失败重试 retries 次后抛 RuntimeError。
    provider=none 时立即抛出可识别异常, 不渲染、不调外部模型、不进入待办循环。
    API key 由 vision.js 从同目录 .env 加载, 本函数不读取也不打印。"""
    provider = get_vision_provider()
    if provider == "none":
        raise VisionDisabledError(
            "当前视觉功能已禁用/不可用(BANK_PDF_VISION_PROVIDER=none); "
            "请改用文字层确定性路径, 或设置 BANK_PDF_VISION_PROVIDER=visionjs/model")
    if provider == "model":
        return _model_vision_call(png_path, prompt, retries=retries, timeout=timeout)
    if provider == "api":
        return _api_vision_call(png_path, prompt, retries=retries, timeout=timeout)
    if not os.path.exists(png_path):
        raise RuntimeError(f"渲染 PNG 不存在: {png_path}")
    node = _node_cmd()
    last_err = None
    for attempt in range(retries + 1):
        try:
            env = dict(os.environ)
            # 工作目录设为 vision.js 所在目录, 保证 dotenv 能加载 .env(兜底)
            r = subprocess.run(
                [node, VISION_JS, os.path.abspath(png_path), prompt],
                capture_output=True, text=True, encoding="utf-8", errors="replace",
                timeout=timeout, env=env, cwd=_VISION_ENV_BASE,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
            if r.returncode != 0:
                last_err = (r.stderr or r.stdout or "").strip()[-400:]
                raise RuntimeError(f"vision.js 退出码 {r.returncode}")
            out = (r.stdout or "").strip()
            if not out:
                last_err = (r.stderr or "").strip()[-400:]
                raise RuntimeError("vision.js 无输出")
            return out
        except RuntimeError:
            raise
        except Exception as e:  # noqa: BLE001 (超时/编码等偶发失败, 重试)
            last_err = str(e)
    raise RuntimeError(f"视觉调用失败(重试 {retries} 次后): {last_err}")


def _last_zh_line(text, require_comma=True):
    """取最后一个满足条件的行: 默认需含逗号且含中文(列名答案约定);
    无候选时放宽到"含中文的行", 再放宽到非空行。"""
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    for need_comma in (require_comma, False):
        cand = [ln for ln in lines
                if (not need_comma or "," in ln or "，" in ln)
                and re.search(r"[\u4e00-\u9fff]", ln)]
        if cand:
            return cand[-1]
    return lines[-1] if lines else ""


def call_vision(png_path, prompt, retries=1, timeout=150):
    """调用视觉并返回解析后的答案文本(最后一个含逗号且含中文的行)。"""
    return _last_zh_line(call_vision_raw(png_path, prompt, retries=retries, timeout=timeout))


def parse_columns_answer(text):
    """把视觉表头列名答案解析为锚点列表(按逗号/顿号/换行分隔, 去空)。"""
    anchors = []
    for chunk in re.split(r"[,\n，、]", text):
        a = chunk.strip().strip(";；:：").strip()
        if a:
            anchors.append(a)
    return anchors


def parse_date_answer(text):
    """从视觉回答中提取日期样例文本(首个日期形态 token)。"""
    m = re.search(r"(?:19|20)\d{2}[年/\-]?\d{1,2}[月/\-]?\d{1,2}日?|\d{8}", text)
    return m.group(0).replace("年", "/").replace("月", "/").replace("日", "") if m else ""


def parse_layout_answer(text):
    """把视觉布局回答映射为布局枚举。返回 None 表示无法判定。"""
    if not text:
        return None
    t = text
    if any(k in t for k in ("网格", "表格线", "格子")):
        return "grid"
    if any(k in t for k in ("竖排", "纵向排", "每条记录 2 行", "每条记录占 2 行")):
        return "vertical"
    if any(k in t for k in ("多行", "列式", "单元格拆", "堆叠")):
        return "columnar-multiline"
    if any(k in t for k in ("单行", "一行", "常规", "普通")):
        return "single-line"
    return None


def render_page_png(doc, pno, matrix_zoom=2.0, out_path=None):
    """渲染 PDF 第 pno 页为 PNG(默认 2x)。返回输出路径。"""
    out = out_path or os.path.join(
        tempfile.gettempdir(), f"bs2x_page{pno+1}_{os.getpid()}.png")
    pix = doc[pno].get_pixmap(matrix=pymupdf.Matrix(matrix_zoom, matrix_zoom))
    pix.save(out)
    return out


def ask_header_columns(doc, pno=0, password=None, retries=1, debug=False):
    """渲染第 pno 页并让视觉读出表头列名。返回锚点列表(失败返回 [])。"""
    png = render_page_png(doc, pno)
    try:
        text = call_vision(
            png,
            "请只输出这张银行对账单/交易明细表格的表头列名，从上到下、从左到右用英文逗号分隔，"
            "不要输出解释、序号或其它内容。例如：交易时间,摘要,借方发生额,贷方发生额,余额,对方户名",
            retries=retries,
        )
        anchors = parse_columns_answer(text)
        if debug:
            print(f"[DEBUG] 视觉读表头(第{pno+1}页): {anchors}")
        return anchors
    finally:
        try:
            os.remove(png)
        except OSError:
            pass
