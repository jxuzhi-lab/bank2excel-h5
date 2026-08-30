# -*- coding: utf-8 -*-
"""glm_ocr.py —— GLM-OCR 主路径扫描件解析(服务端专属, 2026-08-30)。

流程: 逐页渲染 JPEG → 智谱 layout_parsing(glm-ocr) → md_results 中的 HTML
<table> 解析为记录行。相比 RapidOCR 夹心层(ocr_layer.py): 无水印污染、表头
结构准确, 但约 10% 页会生成退化(输出截断)——本模块逐页检测截断(表格未闭合
或 completion 顶限), 截断页交由调用方回退夹心层。

计费: 0.2 元/百万 token(输入输出同价), 实测 ~4700 token/页(A4 2.5x)。
配置(环境变量): GLM_OCR_KEY(智谱 API key, 未配置则模块整体禁用) /
GLM_OCR_CONCURRENCY(默认 2, 实测并发 4 会触发限流) /
GLM_OCR_ZOOM(渲染倍率, 默认 2.5) / GLM_OCR_PAGE_CAP(最大页数, 默认 60)。
"""
import base64
import json
import os
import re
import threading
import time
import urllib.request
from html.parser import HTMLParser
from concurrent.futures import ThreadPoolExecutor

import pymupdf

API_URL = "https://open.bigmodel.cn/api/paas/v4/layout_parsing"
TRUNC_COMPLETION = 8000  # completion 达到该值视为顶限(实测退化页精确 8289)

TOKEN_RE = re.compile(
    r"[A-Za-z0-9/\-.:%—]+|[\u4e00-\u9fff（）()]+|\s+|[^A-Za-z0-9/\-.:%—\u4e00-\u9fff（）()]+")


class _TableParser(HTMLParser):
    """宽松 HTML 表格解析: 未闭合的 <table>(截断页)也能取回已完成的行。"""
    def __init__(self):
        super().__init__()
        self.rows, self._row, self._cell = [], None, None

    def handle_starttag(self, tag, attrs):
        if tag == "tr":
            self._row = []
        elif tag in ("td", "th") and self._row is not None:
            self._cell = []

    def handle_endtag(self, tag):
        if tag == "tr" and self._row is not None:
            self.rows.append(self._row)
            self._row = None
        elif tag in ("td", "th") and self._cell is not None:
            self._row.append("".join(self._cell).strip())
            self._cell = None

    def handle_data(self, d):
        if self._cell is not None:
            self._cell.append(d)


def get_key():
    return os.environ.get("GLM_OCR_KEY", "").strip()


def _parse_page_md(md, header_len=None):
    """从单页 md 提取 (表头行|None, 数据行列表, 是否截断)。
    header_len 给定时, 丢弃列数不符的畸形尾行(截断抢救的关键一步)。"""
    header = None
    rows = []
    truncated = False
    for tbl in re.findall(r"<table[^>]*>.*?(?:</table>|$)", md, re.S):
        tp = _TableParser()
        tp.feed(tbl)
        if "</table>" not in tbl:
            truncated = True
        for row in tp.rows:
            if not row:
                continue
            if header is None and "交易时间" in row[0]:
                header = row
                continue
            if header is None and any(k in "".join(row) for k in
                                      ("日期", "时间", "摘要", "余额")) and not re.search(
                    r"\d{1,2}[:：]\d{2}", row[0]):
                header = row
                continue
            rows.append(row)
    if header_len is not None:
        rows = [r for r in rows if len(r) == header_len]
    return header, rows, truncated


def _row_sig(row):
    """行签名(用于拆分重叠区去重): 日期列+金额特征组合。"""
    return tuple(c.replace(",", "").replace(" ", "") for c in row[:2]) + \
        (row[-1].replace(",", "") if row else "",)


def _render_page(pdf_path, page_index, clip=None, zoom=None):
    doc = pymupdf.open(pdf_path)
    zoom = float(zoom or os.environ.get("GLM_OCR_ZOOM", "2.5"))
    try:
        page = doc[page_index]
        pix = page.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom),
                              clip=clip) if clip is not None else \
            page.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom))
        return base64.b64encode(pix.tobytes("jpeg")).decode()
    finally:
        doc.close()


def _split_clips(pdf_path, page_index, depth):
    """把页垂直切 2^depth 段(相邻段 12% 重叠防切行)。"""
    doc = pymupdf.open(pdf_path)
    try:
        h = doc[page_index].rect.height
        w = doc[page_index].rect.width
    finally:
        doc.close()
    n = 2 ** depth
    seg = h / n
    ov = seg * 0.12
    clips = []
    for i in range(n):
        y0 = max(0.0, i * seg - (ov if i else 0))
        y1 = min(h, (i + 1) * seg + (ov if i < n - 1 else 0))
        clips.append(pymupdf.Rect(0, y0, w, y1))
    return clips


def _ocr_page_resilient(pdf_path, page_index, depth=0, max_depth=2, zoom=None,
                        usage=None):
    """单页识别, 带退化自愈:
    截断页(生成退化)递归对半拆分重试(退化跟内容走, 拆分后每次生成都重新
    开始, 多数可跳出); 每级用完整行抢救(丢弃列数不符的畸形尾行), 重叠区
    按行签名去重。返回 (header, rows, salvaged: bool, hard_fail: bool)。"""
    b64 = _render_page(pdf_path, page_index, zoom=zoom)
    md, comp, err = _call_layout_parsing(b64, usage=usage)
    if err is not None:
        return None, [], False, True
    header, rows, trunc = _parse_page_md(md)
    if comp >= TRUNC_COMPLETION:
        trunc = True
    if not trunc:
        return header, rows, False, not rows
    if depth >= max_depth:
        # 抢救完整行(需要表头定列数; 无表头按最常见列数兜底)
        hlen = len(header) if header else _modal_len(rows)
        if hlen:
            _h2, rows2, _t = _parse_page_md(md, header_len=hlen)
            return header, rows2, True, not rows2
        return header, rows, True, not rows
    # 递归拆分
    seen, merged, salvaged = set(), [], False
    hdr = header
    for clip in _split_clips(pdf_path, page_index, depth + 1):
        b64i = _render_page(pdf_path, page_index, clip=clip, zoom=zoom)
        mdi, compi, erri = _call_layout_parsing(b64i, usage=usage)
        if erri is not None:
            salvaged = True
            continue
        hi, ri, ti = _parse_page_md(mdi)
        if compi >= TRUNC_COMPLETION:
            ti = True
        if ti:
            salvaged = True
            hlen = len(hi) if hi else (_modal_len(ri) or (len(hdr) if hdr else 0))
            if hlen:
                _h3, ri, _t3 = _parse_page_md(mdi, header_len=hlen)
        if hi and hdr is None:
            hdr = hi
        for r in ri:
            if hdr and len(r) != len(hdr):
                continue
            sig = _row_sig(r)
            if sig in seen:
                continue
            seen.add(sig)
            merged.append(r)
    return hdr, merged, salvaged, not merged


def _modal_len(rows):
    """行列表中最常见的列数(畸形尾行会稀释但正确列数占多数)。"""
    from collections import Counter
    if not rows:
        return 0
    return Counter(len(r) for r in rows).most_common(1)[0][0]


def _call_layout_parsing(b64, retries=3, usage=None):
    body = json.dumps({"model": "glm-ocr",
                       "file": "data:image/jpeg;base64," + b64}).encode()
    last = None
    for attempt in range(retries):
        req = urllib.request.Request(
            API_URL, data=body,
            headers={"Content-Type": "application/json",
                     "Authorization": "Bearer " + get_key()})
        try:
            with urllib.request.urlopen(req, timeout=300) as r:
                resp = json.loads(r.read().decode())
            md = resp.get("md_results")
            if isinstance(md, list):
                md = md[0] if md else ""
            u = resp.get("usage", {})
            if usage is not None:
                with _usage_lock:
                    usage["glm_calls"] = usage.get("glm_calls", 0) + 1
                    usage["glm_prompt"] = usage.get("glm_prompt", 0) + int(u.get("prompt_tokens", 0))
                    usage["glm_completion"] = usage.get("glm_completion", 0) + int(u.get("completion_tokens", 0))
            return md or "", int(u.get("completion_tokens", 0)), None
        except urllib.error.HTTPError as e:
            err = e.read().decode("utf-8", "replace")[:200]
            last = f"HTTP {e.code}: {err}"
            if e.code in (429, 500, 502, 503, 504):
                time.sleep(5 * (attempt + 1))
                continue
            break
        except Exception as e:  # noqa: BLE001
            last = f"{type(e).__name__}: {e}"
            time.sleep(5 * (attempt + 1))
    return "", 0, last


_usage_lock = threading.Lock()


def parse_scanned_pdf(pdf_path, page_count, debug=False, usage=None):
    """usage: 可选 dict, 就地累计 GLM-OCR 真实 token 消耗(线程安全)。"""
    """GLM-OCR 主路径入口(逐页, 截断页自动拆分重试)。

    返回 dict: {"pages": [{index, header, rows, salvaged, hard_fail}...],
                "tokens": int, "errors": [str]}
    salvaged=True 的页数据完整性与余额链由调用方校验(balance_gaps)。"""
    cap = int(os.environ.get("GLM_OCR_PAGE_CAP", "60"))
    pages_idx = list(range(min(page_count, cap)))
    if not get_key():
        return None

    conc = max(1, min(int(os.environ.get("GLM_OCR_CONCURRENCY", "2")), 4))
    results = {}
    errors = []

    def work(i):
        hdr, rows, salvaged, hard = _ocr_page_resilient(
            pdf_path, i, usage=usage)
        return i, {"index": i, "header": hdr, "rows": rows,
                   "salvaged": salvaged, "hard_fail": hard}

    with ThreadPoolExecutor(conc) as ex:
        for i, res in ex.map(work, pages_idx):
            results[i] = res
            if res["hard_fail"]:
                errors.append(f"第{i+1}页识别失败(重试后仍无有效数据)")
    return {"pages": [results[i] for i in pages_idx],
            "errors": errors}


BALANCE_KEYS = ("余额", "结余", "结存")


def balance_gaps(header, records):
    """余额链断点检测(引擎三重校验哲学的平移): 上一行余额 ± 借贷 ≠ 下一行
    余额 → 中间可能丢记录。返回断点所在记录序号(1-based)列表。
    无法定位余额列/金额列时返回 []。"""
    if not header or not records:
        return []
    bal_i = next((i for i, n in enumerate(header)
                  if any(k in n for k in BALANCE_KEYS)), None)
    if bal_i is None:
        return []
    deb_i = next((i for i, n in enumerate(header)
                  if any(k in n for k in ("借方", "支出", "汇出"))), None)
    cre_i = next((i for i, n in enumerate(header)
                  if any(k in n for k in ("贷方", "收入", "汇入"))), None)
    if deb_i is None and cre_i is None:
        # 单金额列(发生额/收入-支出单边)无法确定资金方向 → 期望余额不可算,
        # 跳过断点检测(北京银行等过滤型/单列格式), 防误报。
        amt_i = next((i for i, n in enumerate(header)
                      if "发生额" in n or "金额" in n), None)
        if amt_i is not None:
            return []

    def num(s):
        s = str(s).replace(",", "").replace("元", "").strip()
        try:
            return round(float(s), 2)
        except Exception:
            return None

    gaps = []
    prev_bal = None
    for idx, r in enumerate(records):
        if bal_i >= len(r):
            continue
        bal = num(r[bal_i])
        if bal is None:
            continue
        if prev_bal is not None:
            d = num(r[deb_i]) if deb_i is not None and deb_i < len(r) else None
            c = num(r[cre_i]) if cre_i is not None and cre_i < len(r) else None
            expect = prev_bal + (c or 0) - (d or 0)
            if abs(expect - bal) > 0.01:
                gaps.append(idx + 1)  # 第 idx+1 行之前可能缺记录
        prev_bal = bal
    return gaps


def _merge_pages(pages):
    """页列表 → (header, records, 各页行数前缀和)。"""
    header = None
    records = []
    counts = []
    for pg in pages:
        if header is None and pg["header"]:
            header = pg["header"]
        records.extend(pg["rows"])
        counts.append(len(pg["rows"]))
    if header is None:
        raise RuntimeError("GLM-OCR: 所有页均未解析出表头")
    return header, records, counts


def build_engine_records(pages):
    """把逐页解析结果拼成引擎可用形态(表头取首个非空页, 记录按页序顺接)。"""
    header, records, _counts = _merge_pages(pages)
    return header, records, {"page_count": len(pages),
                             "truncated_words": 0,
                             "pages": [{"records": c} for c in _counts],
                             "wechat": False}


def refine_gaps(pdf_path, pages, max_rounds=2, debug=False, usage=None):
    """余额链断点驱动的定向重试:
    断点行 → 映射回所在页(按页行数前缀和) → 该页整页重新识别(跨次调用存在
    非确定性, 实测同一退化页跨轮可恢复) → 重组 → 再检测; 直至无断点或轮次用尽。
    返回 (pages, gaps)。"""
    for _round in range(max_rounds):
        try:
            header, records, counts = _merge_pages(pages)
        except RuntimeError:
            return pages, [], []
        gaps = balance_gaps(header, records)
        if not gaps:
            return pages, [], []
        # 断点所在记录序号(1-based) → 页索引
        bad_page_idx = set()
        cum = 0
        for pi, cnt in enumerate(counts):
            for gpos in gaps:
                if cum < gpos <= cum + cnt:
                    bad_page_idx.add(pi)
            cum += cnt
        if not bad_page_idx:
            return pages, gaps, baidu_pages
        if debug:
            print(f"[GLM-OCR] 第{_round+1}轮定向重试: 断点{gaps} → 页 "
                  f"{sorted(i+1 for i in bad_page_idx)}", flush=True)
        for pi in sorted(bad_page_idx):
            # 重试轮换渲染倍率: 退化与 tokenization 相关, 改变切词可能跳出循环
            z = float(os.environ.get("GLM_OCR_ZOOM", "2.5")) * (1.25 if _round % 2 == 0 else 0.8)
            hdr2, rows2, salv2, hard2 = _ocr_page_resilient(
                pdf_path, pi, zoom=z, usage=usage)
            if rows2 and not hard2:
                pages[pi] = {"index": pi, "header": hdr2 or pages[pi]["header"],
                             "rows": rows2, "salvaged": salv2,
                             "hard_fail": False}
    # GLM 轮次用尽仍断点 → 百度表格识别精锐兜底(只送断点页, ~0.025 元/页)
    try:
        header, records, counts = _merge_pages(pages)
        gaps = balance_gaps(header, records)
    except RuntimeError:
        return pages, [], []
    if not gaps:
        return pages, [], []
    try:
        import baidu_ocr
        if not baidu_ocr.available():
            return pages, gaps, []
    except Exception:  # noqa: BLE001
        return pages, gaps, []
    # 内容锚定定位断点页: 断点前一行的行签名在哪个页就重做哪个页
    # (按累计行数映射会因各页行数与基准漂移而错页, 已实测踩坑)
    def _row_sig(r):
        return tuple(str(c).replace(",", "").replace(" ", "") for c in r[:2]) + \
            (str(r[-1]).replace(",", "") if r else "",)

    bad_page_idx = set()
    for gpos in gaps:
        anchor = None
        if gpos >= 2:
            anchor = _row_sig(records[gpos - 2])
        for pi, pg in enumerate(pages):
            if not pg["rows"]:
                continue
            if anchor is None or any(_row_sig(r) == anchor for r in pg["rows"]):
                # 锚行所在页; 首行断点(无锚)时取含断点后行之前的所有页中最后一个
                bad_page_idx.add(pi)
                break
        else:
            if pages:
                bad_page_idx.add(len(pages) - 1)
    if debug:
        print(f"[GLM-OCR] 升级百度兜底: 断点{gaps} → 页 "
              f"{sorted(i+1 for i in bad_page_idx)}", flush=True)
    target = next((p["header"] for p in pages if p["header"]), None)
    if target is None:
        return pages, gaps, []
    redone = baidu_ocr.redo_pages(pdf_path, sorted(bad_page_idx),
                                  target_header=target, debug=debug,
                                  usage=usage)
    baidu_pages = []
    for pi, res in redone.items():
        if pi < len(pages) and res["records"]:
            pages[pi] = {"index": pi, "header": target,
                         "rows": res["records"], "salvaged": True,
                         "hard_fail": False, "source": "baidu"}
            baidu_pages.append(pi + 1)
    if baidu_pages:
        _header0, _recs0, _c0 = _merge_pages(pages)
        _dedup_adjacent(pages, header=_header0)
        _fill_baidu_balances(pages, _header0)
        header, records, _c = _merge_pages(pages)
        gaps = balance_gaps(header, records)
    return pages, gaps, baidu_pages


def _fill_baidu_balances(pages, header):
    """百度来源页的空余额用余额链补全: prev ± 本行借贷可精确推导
    (百度页金额实测 100% 准确, 推导值可靠); 仅填 baidu 来源页的行。"""
    if not header:
        return
    bi = _find_col_idx(header, "余额", "结余", "结存")
    deb_i = _find_col_idx(header, "借方", "支出", "汇出")
    cre_i = _find_col_idx(header, "贷方", "收入", "汇入")
    if bi is None:
        return

    def fnum(s):
        s = str(s).replace(",", "").strip()
        try:
            return round(float(s), 2)
        except Exception:
            return None

    prev = None
    for pg in pages:
        for r in pg["rows"]:
            bal = fnum(r[bi]) if bi < len(r) else None
            if pg.get("source") == "baidu" and prev is not None:
                deb = fnum(r[deb_i]) if deb_i is not None and deb_i < len(r) else None
                cre = fnum(r[cre_i]) if cre_i is not None and cre_i < len(r) else None
                if deb is not None or cre is not None:
                    expect = round(prev + (cre or 0) - (deb or 0), 2)
                    # 百度页金额实测高度可靠: 空余额填推导值, 错误余额纠错
                    if bal is None or abs(bal - expect) > 0.01:
                        r[bi] = f"{expect:.2f}"
                        bal = expect
            if bal is not None:
                prev = bal


def _dedup_adjacent(pages, header=None):
    """页缝相邻去重: (日期数字串, 余额数字串) 精确键相同的相邻记录只留一条
    (GLM 页边界与 PDF 真实页边界可能差一行, 与百度页拼接时产生重复)。
    用数字字符串而非 float: 20 位流水号/大额余额会超出 float64 精度导致误判。"""
    import re as _re
    all_recs = []
    for pg in pages:
        all_recs += pg["rows"]
    di = _find_col_idx(header, "交易时间", "交易日期", "日期", "记账日期")
    bi = _find_col_idx(header, "余额", "结余", "结存")

    def key(r):
        if not r:
            return None
        dk = _re.sub(r"\D", "", str(r[di]))[:8] if di is not None and di < len(r) else ""
        bk = _re.sub(r"[^\d.]", "", str(r[bi])) if bi is not None and bi < len(r) else ""
        return (dk, bk)

    out = []
    for r in all_recs:
        k = key(r)
        # 余额为空的行不可安全判重(同日多笔空余额是百度输出常态), 不折叠
        if out and k is not None and k[1] and k == key(out[-1]):
            continue
        out.append(r)
    remaining = list(out)
    for pg in pages:
        k = min(len(pg["rows"]), len(remaining))
        pg["rows"], remaining = remaining[:k], remaining[k:]
    if remaining:
        pages[-1]["rows"] += remaining
    return out


def _find_col_idx(header, *keys):
    if not header:
        return None
    return next((i for i, n in enumerate(header)
                 if any(k in str(n) for k in keys)), None)


def mask_records(records):
    """诊断/日志用: 脱敏记录行(金额→0.01, 数字串打码), 避免泄漏。"""
    out = []
    for r in records:
        row = []
        for c in r:
            c = re.sub(r"\d", "0", str(c)) if re.search(r"\d{4,}", str(c)) else c
            row.append(str(c)[:20])
        out.append(row)
    return out
