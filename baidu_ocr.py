# -*- coding: utf-8 -*-
"""baidu_ocr.py —— 百度智能云表格识别兜底(服务端专属, 2026-08-30)。

定位: GLM-OCR 主路径的**定向精锐兜底** —— 仅当余额链断点重试后仍有缺笔时,
把断点所在页送百度表格文字识别V2 重做(25 点/次 ≈ 0.025 元/页, 单文档封顶
BAIDU_OCR_MAX_PAGES 页), 质量实测满分(民生 401/401)。

集成方式(直接网格重组): 百度返回单元格网格(row/col 索引 + words), 表头行
文本按包含关系映射到目标(GLM)表头列, 数据格多行文本按"数量匹配/形态规则"
(日期时间→时间列, 长数字→号码列, 其余→文本列)拆分到目标列 —— 确定性重组,
不经引擎, 无夹心脆弱性。

竖拼已实测否决: 单页三值 21/21, 2 页拼 37/41, 3 页拼 17/63 —— 服务端对超高
图内部降采样, 质量随图高单调劣化, 不省这个点数。

配置(环境变量): BAIDU_OCR_API_KEY / BAIDU_OCR_SECRET_KEY(未配置则兜底禁用) /
BAIDU_OCR_MAX_PAGES(默认 4, 单文档百度兜底页数上限)。
"""
import base64
import json
import os
import re
import threading
import time
import urllib.parse
import urllib.request
from collections import defaultdict

import pymupdf

TABLE_URL = "https://aip.baidubce.com/rest/2.0/ocr/v1/table"
TOKEN_URL = "https://aip.baidubce.com/oauth/2.0/token"
ZOOM = float(os.environ.get("BAIDU_OCR_ZOOM", "2.5"))

_tok_lock = threading.Lock()
_tok_cache = {"token": None, "expire": 0.0}
_qps_lock = threading.Lock()
_last_call = [0.0]

DATE_RE = re.compile(r"^\d{4}[/\-.]\d{1,2}[/\-.]\d{1,2}")
TIME_RE = re.compile(r"^\d{1,2}:\d{2}(:\d{2})?$")
DIGITS_RE = re.compile(r"^[\d,]{6,}(\.\d{2})?$")


def get_keys():
    return (os.environ.get("BAIDU_OCR_API_KEY", "").strip(),
            os.environ.get("BAIDU_OCR_SECRET_KEY", "").strip())


def available():
    ak, sk = get_keys()
    return bool(ak and sk)


def _get_token():
    ak, sk = get_keys()
    with _tok_lock:
        if _tok_cache["token"] and time.time() < _tok_cache["expire"] - 3600:
            return _tok_cache["token"]
        url = (TOKEN_URL + "?grant_type=client_credentials"
               + "&client_id=" + urllib.parse.quote(ak)
               + "&client_secret=" + urllib.parse.quote(sk))
        with urllib.request.urlopen(url, timeout=30) as r:
            tok = json.loads(r.read().decode())
        if not tok.get("access_token"):
            raise RuntimeError(f"百度 token 获取失败: {str(tok)[:120]}")
        _tok_cache["token"] = tok["access_token"]
        _tok_cache["expire"] = time.time() + float(tok.get("expires_in", 2592000))
        return _tok_cache["token"]


def _throttle(min_interval=0.55):
    """免费档 QPS≈2, 进程内节流。"""
    with _qps_lock:
        wait = min_interval - (time.time() - _last_call[0])
        if wait > 0:
            time.sleep(wait)
        _last_call[0] = time.time()


def table_ocr_cells(pdf_path, page_index, retries=3):
    """单页表格识别, 返回 cells 列表(无表格/失败抛 RuntimeError)。"""
    tok = _get_token()
    doc = pymupdf.open(pdf_path)
    try:
        pix = doc[page_index].get_pixmap(matrix=pymupdf.Matrix(ZOOM, ZOOM))
    finally:
        doc.close()
    b64 = base64.b64encode(pix.tobytes("jpeg")).decode()
    last = None
    for _attempt in range(retries):
        _throttle()
        req = urllib.request.Request(
            TABLE_URL + "?access_token=" + tok,
            data=urllib.parse.urlencode({"image": b64}).encode(),
            headers={"Content-Type": "application/x-www-form-urlencoded"})
        try:
            with urllib.request.urlopen(req, timeout=120) as r:
                resp = json.loads(r.read().decode())
            if "error_code" in resp:
                if resp.get("error_code") == 18:  # QPS 超限
                    time.sleep(2)
                    continue
                raise RuntimeError(f"百度表格识别错误: {str(resp)[:150]}")
            cells = []
            for t in resp.get("tables_result", []):
                cells += t.get("body", [])
            if not cells:
                raise RuntimeError("百度表格识别: 未返回表格")
            return cells
        except RuntimeError:
            raise
        except Exception as e:  # noqa: BLE001
            last = str(e)
            time.sleep(2 * (_attempt + 1))
    raise RuntimeError(f"百度表格识别失败(重试 {retries} 次): {last}")


def _norm_name(s):
    return re.sub(r"[\s|/｜]", "", str(s or ""))


def _match_targets(cell_text, target_header):
    """表头格文本 → 匹配的目标列索引列表(按子串出现顺序)。"""
    norm = _norm_name(cell_text)
    if not norm:
        return []
    exact = [i for i, t in enumerate(target_header) if _norm_name(t) == norm]
    if exact:
        return exact
    hits = [i for i, t in enumerate(target_header)
            if _norm_name(t) and _norm_name(t) in norm]
    hits.sort(key=lambda i: norm.find(_norm_name(target_header[i])))
    return hits


def _split_cell_lines(lines, targets, target_header):
    """单元格文本行 → 多目标列分配: 数量匹配按序; 否则按形态
    (日期/时间→时间类列, 长数字→号码类列, 其余→首个文本列)。"""
    if len(targets) == 1:
        return {targets[0]: " ".join(lines)}
    if len(lines) == len(targets):
        return dict(zip(targets, lines))
    out = {t: [] for t in targets}
    tmeta = [(t, str(target_header[t])) for t in targets]
    for raw in lines:
        ln = raw.strip()
        if not ln:
            continue
        placed = False
        if DATE_RE.match(ln) or TIME_RE.match(ln):
            for t, name in tmeta:
                if "时间" in name or "日期" in name:
                    out[t].append(ln)
                    placed = True
                    break
        if not placed and DIGITS_RE.match(ln.replace(",", "")):
            for t, name in tmeta:
                if any(k in name for k in ("号码", "流水号", "账号", "单号")):
                    out[t].append(ln)
                    placed = True
                    break
        if not placed:
            for t, name in tmeta:
                if ("时间" not in name and "日期" not in name
                        and not any(k in name for k in ("号码", "流水号"))):
                    out[t].append(ln)
                    placed = True
                    break
        if not placed and tmeta:
            out[tmeta[0][0]].append(ln)
    return {t: " ".join(v) for t, v in out.items()}


def reconstruct_records(cells, target_header):
    """百度单元格网格 → 目标表头列序的记录行(纯网格重组, 不经引擎)。"""
    header_cells = sorted([c for c in cells if c.get("row_start") == 0],
                          key=lambda c: c.get("col_start", 0))
    col_targets = {}
    for hc in header_cells:
        col_targets[hc.get("col_start")] = (
            _match_targets(hc.get("words", ""), target_header) or [None])
    rows = defaultdict(dict)
    for c in cells:
        r = c.get("row_start", 0)
        if r == 0:
            continue
        targets = col_targets.get(c.get("col_start"), [None])
        lines = [ln for ln in str(c.get("words", "")).split("\n") if ln.strip()]
        for t, v in _split_cell_lines(lines, targets, target_header).items():
            if t is not None:
                rows[r][t] = v
    recs = [[rows[r].get(t, "") for t in range(len(target_header))]
            for r in sorted(rows)]
    recs = [_postfix_row(r, target_header) for r in recs]
    return _repair_balances(recs, target_header)


def _fnum(s):
    s = str(s).replace(",", "").strip()
    try:
        return round(float(s), 2)
    except Exception:
        return None


def _repair_balances(recs, target_header):
    """余额链定向修复: 百度偶发把余额并进相邻列(如流水号格粘成
    '348,523.46 31392...'), 该行余额列为空。已知前一行余额与本行借贷
    → 期望余额可算; 在行内各格找该值, 找到即搬回余额列。"""
    bi = _find_col(target_header, "余额", "结余", "结存")
    if bi is None:
        return recs
    deb_i = _find_col(target_header, "借方", "支出", "汇出")
    cre_i = _find_col(target_header, "贷方", "收入", "汇入")
    prev = None
    repaired = 0
    for r in recs:
        bal = _fnum(r[bi]) if bi < len(r) else None
        if bal is None:
            deb = _fnum(r[deb_i]) if deb_i is not None and deb_i < len(r) else None
            cre = _fnum(r[cre_i]) if cre_i is not None and cre_i < len(r) else None
            if prev is not None:
                expect = round(prev + (cre or 0) - (deb or 0), 2)
                # 行内找期望余额(容忍千分位/空白粘连形态)
                for ci in range(len(r)):
                    if ci == bi:
                        continue
                    cell = str(r[ci])
                    for cand in (f"{expect:,.2f}", f"{expect:.2f}"):
                        if cand in cell:
                            r[ci] = cell.replace(cand, " ").strip()
                            r[bi] = f"{expect:.2f}"
                            bal = expect
                            repaired += 1
                            break
                    if bal is not None:
                        break
        if bal is not None:
            prev = bal
    return recs


def _find_col(target_header, *keys):
    return next((i for i, n in enumerate(target_header)
                 if any(k in str(n) for k in keys)), None)


def _postfix_row(row, target_header):
    """行级修复: 日期列值粘连摘要文本时拆回(百度 cells 偶发把
    "2025/03/13一般账户转账" 粘成一行, 摘要列留空)。"""
    ti = _find_col(target_header, "交易时间", "交易日期", "记账日期", "日期")
    mi = _find_col(target_header, "摘要", "用途", "备注")
    if ti is None or mi is None or ti >= len(row) or mi >= len(row):
        return row
    v = str(row[ti]).strip()
    m = re.match(r"^(\d{4}[/\-.]\d{1,2}[/\-.]\d{1,2})(.*)$", v)
    if not m:
        return row
    date_part, rest = m.group(1), m.group(2).strip()
    tm = re.search(r"(\d{1,2}:\d{2}(?::\d{2})?)\s*$", rest)
    time_part = ""
    if tm:
        time_part = tm.group(1)
        rest = rest[:tm.start()].strip()
    if rest and not str(row[mi]).strip():
        row[mi] = rest
    row[ti] = (date_part + " " + time_part).strip()
    return row


def redo_pages(pdf_path, page_indexes, target_header, debug=False, usage=None):
    """百度兜底入口: 逐页 表格识别→网格重组→对齐目标表头。
    返回 {page_index: {"header", "records"}}, 失败页不在结果中。"""
    results = {}
    cap = int(os.environ.get("BAIDU_OCR_MAX_PAGES", "4"))
    for pi in sorted(page_indexes)[:cap]:
        try:
            cells = table_ocr_cells(pdf_path, pi)
            if usage is not None:
                usage["baidu_calls"] = usage.get("baidu_calls", 0) + 1  # 25 点/次
            recs = reconstruct_records(cells, target_header)
            if recs:
                results[pi] = {"header": list(target_header), "records": recs}
        except Exception as e:  # noqa: BLE001 (单页失败不影响其余页)
            if debug:
                print(f"[百度兜底] 第{pi+1}页失败: {e}", flush=True)
    return results
