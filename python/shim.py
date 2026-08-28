# -*- coding: utf-8 -*-
"""shim.py —— bank2excel-h5 浏览器端入口(PYODIDE 模式)。

职责:
  1. convert_bytes(): 字节流进 / 字节流出的一站式转换
     PDF bytes → 引擎提取 → xlsx bytes + 统计报告 + 警告日志
  2. export_diagnosis(): 未识别格式时导出脱敏诊断包 JSON(M4 闭环用)
  3. 页级进度回调 on_progress: 每页提取后上报 (页号, 总页数)

设计约束(与引擎 H5 改造对齐):
  - 运行时无 LLM/无视觉/无文件系统: PYODIDE=1 分支已短路 vision/onboard/诊断落盘
  - xlsx 优先 xlsxwriter(Pyodide 可 micropip 安装), 缺失回退 openpyxl
  - 日志经引擎 log() 收集进 LOG_BUFFER, 转换完随结果返回
"""
import io
import json
import re

import extract_bank_statement as eng  # 引擎(同目录, H5 改造版)


def _set_env():
    """幂等设置 PYODIDE 环境开关(须在引擎解析 os.environ 前生效;
    Pyodide 下重复调用无副作用)。"""
    import os
    os.environ["PYODIDE"] = "1"
    eng.PYODIDE = True


# 进度回调: 由 JS 侧注入(py.globals.set("js_progress_cb", ...))后生效
js_progress_cb = None


def _progress(pno, total):
    """引擎页级进度桥(extract_statement 每页循环内被调用)。
    双通道: 1) JS 注入的 js_progress_cb(仅 JS 函数, Pyodide 自动桥接);
            2) 返回值轮询用 LAST_PROGRESS 记录。"""
    global LAST_PROGRESS
    LAST_PROGRESS = (pno, total)
    if js_progress_cb is not None:
        try:
            js_progress_cb(pno, total)
        except Exception:
            pass


LAST_PROGRESS = (0, 0)


def get_progress():
    """JS 侧轮询接口: 返回 (当前页, 总页数)。"""
    return list(LAST_PROGRESS)


def convert_bytes(pdf_bytes, password="", sheet="对账单", debug=False):
    """主入口: 转换单个对账单 PDF。

    Args:
        pdf_bytes: bytes/bytearray, PDF 文件内容
        password:   PDF 打开密码(加密对账单, 如华夏/中行), 无则空串
        sheet:      工作表名(≤31 字符, 自动截断)
        debug:      透传引擎 debug(日志更详细)

    Returns:
        dict: {
          "ok": True,
          "xlsx": bytes,                 # xlsx 文件内容
          "sheet": str,                  # 实际工作表名
          "rows": int,                   # 数据行数(不含表头)
          "report": {...},               # 提取统计(笔数/借贷方合计/余额链等)
          "stats": {...},
          "meta":  {...},                # 页面级元数据(页数/每页记录)
          "pages": int,                  # PDF 总页数
          "warnings": [...],             # 转换过程日志(非 DEBUG)
        }
    Raises:
        原 SDK 语义异常原样抛出(密码错/扫描件/表头未识别/记录为空...),
        由 JS 侧捕获并映射 UI 文案。
    """
    _set_env()
    eng.LOG_BUFFER.clear()

    if not isinstance(pdf_bytes, (bytes, bytearray)):
        pdf_bytes = bytes(pdf_bytes)

    # 挂接页级进度(引擎 PROGRESS_CB → shim._progress → js_progress_cb/轮询)
    eng.set_progress_cb(_progress)
    global LAST_PROGRESS
    LAST_PROGRESS = (0, 0)

    # 引擎 convert_pdf 的 PYODIDE 分支: open_pdf 收 bytes 走 stream,
    # out_path 传 BytesIO, xlsx 写入内存; 失败不落诊断包直接抛。
    buf = io.BytesIO()
    report, stats, meta, _out, sheet_name, n_cols = eng.convert_pdf(
        pdf_bytes,
        out_path=buf,
        sheet=(sheet or "对账单")[:31],
        debug=debug,
        password=password or None,
        vision_fallback=False,   # 无视觉环境
        auto_onboard=False,      # 无 onboarding(无视觉/无文件系统)
        diag=False,              # 不写诊断包(无文件系统)
        write_meta=False,
        quick_classify=True,     # 扫描件前置拦截(纯文字层判定, 不用视觉)
        vision_provider="none",
    )
    return {
        "ok": True,
        "xlsx": buf.getvalue(),
        "sheet": sheet_name,
        "rows": stats.get("rows"),
        "report": report,
        "stats": stats,
        "meta": meta,
        "pages": (meta or {}).get("page_count"),
        "warnings": list(eng.LOG_BUFFER),
    }


# ---------------- M4: 脱敏诊断包 ----------------

_MASK_PATTERNS = [
    # 金额: 数字+可选千分位+两位小数 → 0.01
    (re.compile(r"\d{1,3}(?:,\d{3})+\.\d{2}"), "0.01"),
    (re.compile(r"[-+]?\d+\.\d{2}(?!\d)"), "0.01"),
    # 长数字串(账号/流水号/凭证号 ≥8 位) → 掩码
    (re.compile(r"\d{8,}"), "*" * 8),
]


def _mask_text(s):
    """单字符串脱敏: 金额→0.01, 长数字→*"""
    s = str(s)
    for pat, repl in _MASK_PATTERNS:
        s = pat.sub(repl, s)
    return s


def export_diagnosis(pdf_bytes, filename="未命名.pdf", max_pages=3):
    """未识别格式时导出脱敏诊断包(几 KB JSON, 不含真实数据)。

    内容: 文件名/页数/前 N 页的 表头候选词+y坐标+x区间 + 前几行词骨架
    (所有金额→0.01, 账号→*, 户名等文本保留结构但数字脱敏),
    维护者拿到后可离线分析版式并生成格式描述符。"""
    _set_env()
    import pymupdf

    doc = pymupdf.Document(stream=bytes(pdf_bytes))
    pages = []
    try:
        if doc.needs_pass:
            pages.append({"error": "encrypted"})
        n = min(len(doc), max_pages)
        for pno in range(n):
            page = doc[pno]
            words = page.get_text("words")
            if page.rotation != 0:
                words = eng.normalize_page_words(page, words)
            # 只保留坐标骨架 + 脱敏词文本
            skeleton = [
                {
                    "x0": round(w[0], 1), "y0": round(w[1], 1),
                    "x1": round(w[2], 1), "y1": round(w[3], 1),
                    "t": _mask_text(w[4]),
                }
                for w in words[:200]  # 每页最多 200 词, 控制体积
            ]
            pages.append({
                "page": pno + 1,
                "rotation": page.rotation,
                "size": [round(page.rect.width, 1), round(page.rect.height, 1)],
                "n_words": len(words),
                "words": skeleton,
            })
    finally:
        doc.close()

    return json.dumps({
        "format": "bank2excel-h5-diagnosis/1",
        "filename": _mask_text(filename),
        "total_pages": len(doc),
        "engine": "extract_bank_statement(h5)",
        "pages": pages,
    }, ensure_ascii=False, separators=(",", ":"))
