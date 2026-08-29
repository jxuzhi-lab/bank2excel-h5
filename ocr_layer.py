# -*- coding: utf-8 -*-
"""ocr_layer.py —— 服务端扫描件 OCR 夹心层(仅 VPS/CPython, 2026-08-29)。

原理: 对无文字层的页, 渲染位图 → RapidOCR 逐行识别 → 把识别文本作为
"不可见文字"(render_mode=3)按 OCR 检测框坐标写回 PDF(夹心层/sandwich),
之后的转换完全复用引擎既有规则管道 —— 引擎零改动。

边界:
  - 服务端专属能力(浏览器/PYODIDE 不可能做 OCR);
  - 干净打印体效果好; 拍照歪斜/低分辨率/印章大面积遮挡时识别率明显下降,
    OCR 引入的识别错误会原样进入转换结果(内容由调用方与基准核对);
  - OCR_ENABLED=0 可整体关闭; 页数上限 _PAGE_CAP 防止极端大件长时间占用 CPU。
"""
import os
import re
import threading
import time

import pymupdf

ZOOM = float(os.environ.get("OCR_ZOOM", "3.0"))    # 渲染倍率(~216dpi)
PAGE_CAP = int(os.environ.get("OCR_PAGE_CAP", "50"))
MIN_WORDS = int(os.environ.get("OCR_MIN_WORDS", "5"))  # 少于该词数视为无文字层

_lock = threading.Lock()
_engine = None

# 文本行 → 词元: ASCII 连续串(日期/金额/单号)整体, 中文连续串整体, 其余单列
_TOKEN_RE = re.compile(
    r"[A-Za-z0-9/\-.:%—]+|[\u4e00-\u9fff（）()]+|\s+|[^A-Za-z0-9/\-.:%—\u4e00-\u9fff（）()]+")
_CJK_RE = re.compile(r"[\u4e00-\u9fff]")


def get_engine():
    """懒加载单例(每进程一份模型, 首次调用初始化约 2-5s)。"""
    global _engine
    if _engine is None:
        with _lock:
            if _engine is None:
                from rapidocr_onnxruntime import RapidOCR
                _engine = RapidOCR()
    return _engine


def needs_ocr(doc, max_pages=None):
    """是否含有(近乎)无文字层的页。"""
    n = len(doc) if max_pages is None else min(len(doc), max_pages)
    return any(len(doc[p].get_text("words")) < MIN_WORDS for p in range(n))


def _tokens(line_text):
    """把 OCR 行文本拆成词元(拆掉空格)。"""
    return [t for t in _TOKEN_RE.findall(line_text) if t.strip()]


def _insert_line(page, scale, box, text):
    """把一行 OCR 结果按检测框位置写成不可见文字(词元级横向按字宽比例铺开)。"""
    xs = [p[0] for p in box]
    ys = [p[1] for p in box]
    px0, px1 = min(xs), max(xs)
    py0, py1 = min(ys), max(ys)
    h_pt = (py1 - py0) * scale
    if h_pt <= 0 or not text.strip():
        return
    fontsize = max(4.0, h_pt * 0.82)
    y_base = (py0 + (py1 - py0) * 0.80) * scale  # OCR 框底上方 ~80% 处 ≈ 基线
    toks = _tokens(text)
    if not toks:
        return
    # 估算词元宽: CJK 字宽≈fontsize, ASCII≈0.55*fontsize
    widths = [len(t) * fontsize * (1.0 if _CJK_RE.search(t) else 0.55) for t in toks]
    total = sum(widths)
    avail = max((px1 - px0) * scale, total * 0.5)  # 框宽过窄时防重叠, 允许溢出
    x = px0 * scale
    step = avail / total
    for t, w in zip(toks, widths):
        page.insert_text(
            (x, y_base), t, fontsize=fontsize,
            fontname="china-s" if _CJK_RE.search(t) else "helv",
            render_mode=3)  # 3 = 不可见
        x += w * step


def build_text_layer(pdf_path, out_path, debug=False):
    """对 pdf_path 中无文字层的页做 OCR, 输出夹心层 PDF 到 out_path。
    返回 (ocr_pages, total_pages, seconds)。所有页都有文字层时 ocr_pages=0,
    此时 out_path 仍是原样拷贝(调用方也可直接用原文件)。"""
    t0 = time.time()
    eng = get_engine()
    doc = pymupdf.open(pdf_path)
    total = len(doc)
    ocr_pages = 0
    try:
        for pno in range(min(total, PAGE_CAP)):
            page = doc[pno]
            if len(page.get_text("words")) >= MIN_WORDS:
                continue
            pix = page.get_pixmap(matrix=pymupdf.Matrix(ZOOM, ZOOM))
            result, _elapse = eng(pix.tobytes("jpeg"))
            scale = page.rect.width / pix.width
            if result:
                for box, text, _score in result:
                    try:
                        _insert_line(page, scale, box, text)
                    except Exception:  # noqa: BLE001 (单行放置失败不影响整页)
                        pass
            ocr_pages += 1
        doc.save(out_path, garbage=3, deflate=True)
    finally:
        doc.close()
    return ocr_pages, total, round(time.time() - t0, 1)
