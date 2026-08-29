#!/usr/bin/env python3
"""格式描述符 onboarding(2026-08-19 升级): 把"每格式一次"的格式知识沉淀为
format_descriptor.json, 供 extract_bank_statement.py --descriptor 驱动提取,
跳过逐文件的启发式试探(日期格式/布局/页脚/列边界)。

描述符字段:
  - date_pattern: 日期列多数日期形态的正则(如 ^\\d{4}/\\d{1,2}/\\d{1,2}$);
  - layout: single-line / columnar-multiline / vertical(微信竖排) / grid(表格网格);
  - reverse_chronological: 是否时间倒序(首页首条日期 vs 末页末条日期, 确定性判定);
  - header_every_page: 是否每页都有表头(逐页 detect_header_line, 确定性判定);
  - footer_keywords: 页脚样例关键词(逐页页脚行词频统计, 确定性判定);
  - columns: [{name, x0, x1, semantic}] 列模板 + 语义标签(P2-6);
  - anchors: 视觉锚点反查信息(命中/缺失)。

模式:
  --heuristic  纯规则: 列模板来自 detect_column_boundaries, 不调视觉(最快,
               适合批量同格式复用与无视觉环境);
  --anchors    手工/外部传入锚点(逗号分隔): 渲染 → 文字层反查 → 列模板;
  --vision     脚本自动调 vision.js 读表头列名/日期样例/布局(每格式一次)。

用法:
  python onboard_format.py --input 对账单.pdf --heuristic [--output desc.json]
  python onboard_format.py --input 对账单.pdf --anchors "账号,交易时间,..."
  python onboard_format.py --input 对账单.pdf --vision

输出 JSON 可直接用于: python extract_bank_statement.py --input <pdf> --descriptor <json>
"""
import argparse
import json
import os
import re
import sys
from collections import Counter, defaultdict
from datetime import datetime

import pymupdf

from extract_bank_statement import (
    DEFAULT_DATE_PATTERNS,
    build_columns_from_anchors,
    detect_column_boundaries,
    detect_header_line,
    detect_table_grid,
    find_date_rows,
    ground_anchors_on_page,
    ground_header_anchors,
    infer_semantic,
    is_footer_word,
    is_multiline_column_layout,
    is_wechat_doc,
    match_column,
    _centered_mis_slice,
    normalize_page_words,
    refine_cols_with_data_x,
    wechat_column_bounds,
)


# ---------- 确定性推断 ----------
def _page_words_list(doc):
    out = []
    for pno in range(len(doc)):
        pw = doc[pno].get_text("words")
        if doc[pno].rotation != 0:
            pw = normalize_page_words(doc[pno], pw)
        out.append(pw)
    return out


def _is_date_col(ci, cols):
    return ci is not None and any(k in cols[ci][2] for k in ("日期", "时间"))


def infer_date_pattern(page_words_list, cols, debug=False):
    """从日期列文字层统计各日期正则命中数, 取多数形态。
    返回 (pattern_str, sample_date)。限列失败(特殊列名)时放宽到全页。"""
    cnt = Counter()
    sample = None
    for pw in page_words_list:
        for w in pw:
            ci = match_column(w[0], w[4], cols)
            if not _is_date_col(ci, cols):
                continue
            for pat in DEFAULT_DATE_PATTERNS:
                if pat.match(w[4]):
                    cnt[pat.pattern] += 1
                    if sample is None:
                        sample = w[4]
                    break
    if not cnt:
        for pw in page_words_list:
            for w in pw:
                for pat in DEFAULT_DATE_PATTERNS:
                    if pat.match(w[4]):
                        cnt[pat.pattern] += 1
                        if sample is None:
                            sample = w[4]
                        break
    if not cnt:
        return None, None
    pattern, _n = cnt.most_common(1)[0]
    if debug:
        print(f"[DEBUG] 日期形态推断: {pattern} sample={sample} 统计={dict(cnt)}")
    return pattern, sample


def _date_key(text):
    for fmt in ("%Y/%m/%d", "%Y-%m-%d", "%Y%m%d", "%Y-%m/%d"):
        try:
            return datetime.strptime(text, fmt)
        except (ValueError, TypeError):
            pass
    return None


def infer_reverse_chronological(page_words_list, cols, date_pattern_str):
    """时间倒序判定(确定性): 首页最早日期 vs 末页最晚日期。"""
    if not date_pattern_str:
        return None
    pat = re.compile(date_pattern_str)

    def collect(pw):
        vals = []
        for w in pw:
            ci = match_column(w[0], w[4], cols)
            if _is_date_col(ci, cols) and pat.match(w[4]):
                vals.append(w[4])
        return vals

    firsts = collect(page_words_list[0])
    lasts = []
    for pw in reversed(page_words_list):
        lasts = collect(pw)
        if lasts:
            break
    if not firsts or not lasts:
        return None
    k_first = [k for k in (_date_key(v) for v in firsts) if k]
    k_last = [k for k in (_date_key(v) for v in lasts) if k]
    if not k_first or not k_last:
        return None
    if min(k_first) > max(k_last):
        return True
    if min(k_first) < max(k_last):
        return False
    return None


def infer_layout(doc, page_words_list, header_page_idx, header, cols, band_bottom,
                 is_wc, debug=False):
    """布局判定(确定性优先): 微信竖排 > 表格网格 > 17 列签名+动态形态 >
    居中多行单元格(浦发) > 单行。与 extract_statement 的执行路径保持一致。"""
    if is_wc:
        return "vertical"
    try:
        grid = detect_table_grid(doc[header_page_idx].get_drawings())
    except Exception:  # noqa: BLE001
        grid = None
    if grid is not None:
        return "grid"
    words = page_words_list[header_page_idx]
    region = [w for w in words if (band_bottom or 0) + 3 < w[1]]
    date_rows = find_date_rows(words, DEFAULT_DATE_PATTERNS, cols)
    names = [c[2] for c in cols]
    sig17 = (any("交易时间" in n for n in names)
             and any("记账日期" in n for n in names)
             and any(("账户明细编号" in n) or ("交易介质编号" in n) for n in names))
    if sig17 and is_multiline_column_layout(region, date_rows, cols):
        return "columnar-multiline"
    # 居中多行单元格(2026-08-21, 浦发电子对账单): 单元格以记录行带垂直居中,
    # 首行高出日期行 11-17pt, absorb=6 切片会把首行切给上一条记录; 切片错位词
    # ≥2 → columnar-multiline(精确边界)。与 extract_statement 的 _centered 门控一致。
    if len(date_rows) >= 2 and _centered_mis_slice(region, date_rows, cols) >= 2:
        return "columnar-multiline"
    return "single-line"


def infer_header_every_page(page_words_list):
    """每页表头统计(确定性): 命中表头行的页数占比。"""
    total = len(page_words_list)
    if total == 0:
        return False, 0.0
    hits = sum(1 for pw in page_words_list if detect_header_line(pw) is not None)
    return hits == total, hits / total


def infer_footer_keywords(page_words_list, header_y, band_bottom, max_kw=10):
    """页脚样例关键词(确定性): 逐页页脚行(表头下 100pt 且命中页脚形态)中
    出现 ≥2 页的 ≥2 字中文词。仅作描述符信息与弱辅助(extract 端只对
    表头下 100pt 的行生效)。"""
    rows_text = []
    for pw in page_words_list:
        page_rows = defaultdict(list)
        for w in pw:
            page_rows[round(w[1], 1)].append(w)
        for y, ws in page_rows.items():
            if y > (header_y or 0) + 100 and (band_bottom or 0) < y:
                text = "".join(w[4] for w in ws)
                if is_footer_word(text):
                    rows_text.append(text)
    cnt = Counter()
    for text in rows_text:
        for tok in re.findall(r"[\u4e00-\u9fff]{2,}", text):
            cnt[tok] += 1
    multi_page = [t for t, n in cnt.items() if n >= 2]
    return sorted(multi_page, key=lambda t: -cnt[t])[:max_kw]


def _vision_layout_answer(doc, pno=0):
    """视觉布局判定(仅 --vision 模式用, 失败返回 None)。"""
    try:
        from vision_utils import call_vision, render_page_png
        png = render_page_png(doc, pno)
        try:
            ans = call_vision(
                png,
                "这是银行对账单的一页。请判断表格布局类型，只回答一个字母和简短理由："
                "A 每条记录一行(单行)  B 每条记录多行纵向堆叠(列式单元格拆行)  "
                "C 竖排式(每条记录占2行, 如微信支付)  D 有完整表格网格线。"
                "格式如: D, 有网格线",
                retries=1,
            )
        finally:
            try:
                os.remove(png)
            except OSError:
                pass
        t = (ans or "")[0].upper()
        m = {"A": "single-line", "B": "columnar-multiline",
             "C": "vertical", "D": "grid"}
        return m.get(t)
    except Exception as e:  # noqa: BLE001
        print(f"[提示] 视觉布局判定失败: {e}", file=sys.stderr)
        return None


def build_format_descriptor(pdf_path, password=None, mode="heuristic",
                            anchors=None, vision=False, debug=False):
    """生成格式描述符 dict(供 CLI 与 batch_convert 复用)。
    mode: heuristic(纯规则) / anchors(给定锚点) / vision(脚本调视觉读锚点)。
    返回 dict; 失败抛 RuntimeError。"""
    pdf_path = os.path.abspath(pdf_path)
    if not os.path.exists(pdf_path):
        raise FileNotFoundError(f"文件不存在: {pdf_path}")
    doc = pymupdf.open(pdf_path)
    try:
        if doc.needs_pass:
            if not password:
                raise RuntimeError("PDF 已加密, 请用 --password 提供打开密码")
            if not doc.authenticate(password):
                raise RuntimeError("PDF 密码错误")
        page_words_list = _page_words_list(doc)
        is_wc = is_wechat_doc(doc)
        # ---- 找表头页 / 初始列模板 ----
        header_page_idx = None
        header = None
        cols = None
        band_bottom = None
        header_y = None
        for pno in range(min(len(doc), 20)):
            cand = detect_header_line(page_words_list[pno])
            if cand is not None:
                cand_cols, cand_bb = detect_column_boundaries(page_words_list[pno], cand)
                if cand_cols:
                    header_page_idx = pno
                    header = cand
                    cols = cand_cols
                    band_bottom = cand_bb
                    header_y = cand[0]
                    break
        if is_wc:
            # 微信竖排: 固定 8 列模板(动态锚点 P2-5 见下)
            cols = wechat_column_bounds()
            header_page_idx = 0
            header_y = 214.9
            band_bottom = header_y
        # ---- 视觉锚点(可选) ----
        anchors_used = None
        matched = None
        if mode in ("anchors", "vision"):
            if mode == "vision":
                from vision_utils import ask_header_columns
                anchors_used = ask_header_columns(doc, pno=header_page_idx or 0,
                                                  retries=1, debug=debug)
                if not anchors_used:
                    raise RuntimeError("视觉未读出行列名, 无法生成描述符")
            elif anchors:
                anchors_used = [a.strip() for a in str(anchors).split(",") if a.strip()]
            if is_wc and anchors_used:
                # 微信动态列锚点: 视觉读表头 + 数据列 x0 恢复(失败回退硬编码)
                from extract_bank_statement import _wechat_dynamic_bounds
                wc_bounds = _wechat_dynamic_bounds(
                    page_words_list[0], anchors_used, debug=debug)
                if wc_bounds is not None:
                    cols = wc_bounds
                    matched = [m for m in anchors_used]
            if anchors_used and not is_wc:
                if header is not None:
                    _ground = ground_header_anchors(
                        page_words_list[header_page_idx], anchors_used,
                        header[0], header[1], debug=debug)
                    # 2026-08-21 修复: ground_header_anchors 可能返回 None
                    # (锚点全部未命中时), 直接解包会崩 "cannot unpack non-iterable
                    # NoneType"; 未命中时回退启发式列识别(cols 保持 detect 结果)。
                    if _ground is not None:
                        matched, _bb = _ground
                        cols = build_columns_from_anchors(matched)
                        cols = refine_cols_with_data_x(
                            page_words_list[header_page_idx], cols, band_bottom)
                else:
                    # 启发式表头识别失败(印章遮挡/特殊列名): 用视觉锚点逐页
                    # 反推表头带(与 _vision_fallback_extract 同路径, 2026-08-20 修复)。
                    for pno in range(min(len(doc), 5)):
                        grounded = ground_anchors_on_page(
                            page_words_list[pno], anchors_used, debug=debug)
                        if grounded is None:
                            continue
                        header_page_idx = pno
                        header_y = grounded[0]
                        band_bottom = grounded[3]
                        cols = build_columns_from_anchors(grounded[2])
                        cols = refine_cols_with_data_x(
                            page_words_list[pno], cols, grounded[3])
                        matched = list(anchors_used)
                        break
        if cols is None:
            raise RuntimeError("未能自动识别表头, 无法生成描述符; 请用 --vision/--anchors 提供表头锚点")
        # ---- 确定性推断 ----
        date_pattern, date_sample = infer_date_pattern(page_words_list, cols, debug=debug)
        layout = infer_layout(doc, page_words_list, header_page_idx or 0, header,
                              cols, band_bottom, is_wc, debug=debug)
        reverse = infer_reverse_chronological(page_words_list, cols, date_pattern)
        header_every, header_ratio = infer_header_every_page(page_words_list)
        footer_kw = infer_footer_keywords(page_words_list, header_y, band_bottom)
        semantics = [infer_semantic(c[2]) for c in cols]
        # ---- 视觉日期样例/布局交叉(仅 vision 模式) ----
        vision_date = None
        if mode == "vision" and date_pattern is None:
            try:
                from vision_utils import call_vision, parse_date_answer, render_page_png
                png = render_page_png(doc, header_page_idx or 0)
                try:
                    ans = call_vision(
                        png,
                        "请找出该页表格的日期/交易时间列第一条数据记录的日期，"
                        "只输出日期原文，例如 2025/01/03 或 20250103 或 2025-01-03，不要输出其它文字。",
                        retries=1,
                    )
                finally:
                    try:
                        os.remove(png)
                    except OSError:
                        pass
                vision_date = parse_date_answer(ans)
                if vision_date:
                    for pat in DEFAULT_DATE_PATTERNS:
                        if pat.match(vision_date):
                            date_pattern = pat.pattern
                            date_sample = vision_date
                            break
            except Exception as e:  # noqa: BLE001
                print(f"[提示] 视觉日期样例读取失败: {e}", file=sys.stderr)
        # 视觉布局仅作参考字段: 不覆盖确定性布局。确定性判定(微信竖排/表格网格/
        # 17 列签名+动态形态)更可靠——视觉常把英文子表头/多行单元格误判为
        # "列式多行"(华夏英文表头行曾触发强制 17 列边界逻辑导致错位)。
        vision_layout = _vision_layout_answer(doc, header_page_idx or 0) if mode == "vision" else None
        descriptor = {
            "source": os.path.basename(pdf_path),
            "page": (header_page_idx or 0) + 1,
            "wechat": is_wc,
            "date_pattern": date_pattern,
            "date_sample": date_sample or vision_date,
            "layout": layout,
            "vision_layout": vision_layout,
            "reverse_chronological": reverse,
            "header_every_page": header_every,
            "header_pages_ratio": round(header_ratio, 3),
            "footer_keywords": footer_kw,
            "columns": [
                {"name": name, "x0": round(lo, 2), "x1": round(hi, 2), "semantic": sem}
                for (lo, hi, name), sem in zip(cols, semantics)
            ],
            "anchors": {
                "total": len(anchors_used or []),
                "matched": len(matched) if matched is not None else None,
                "list": list(anchors_used) if anchors_used else [],
            },
        }
        return descriptor
    finally:
        doc.close()


def main():
    ap = argparse.ArgumentParser(description="格式描述符 onboarding(生成 format_descriptor.json)")
    ap.add_argument("--input", required=True, help="输入对账单 PDF")
    ap.add_argument("--heuristic", action="store_true",
                    help="纯规则生成描述符(不调视觉, 列模板来自启发式识别)")
    ap.add_argument("--anchors", help="表头列名锚点, 逗号分隔(手动/外部视觉读出)")
    ap.add_argument("--vision", action="store_true",
                    help="脚本自动调 vision.js 读表头列名/日期/布局")
    ap.add_argument("--no-vision", dest="vision_provider",
                    action="store_const", const="none",
                    help="关闭视觉能力(仅 --heuristic/--anchors 可用)")
    ap.add_argument("--vision-provider", default="auto",
                    choices=["auto", "visionjs", "model", "none"],
                    help="auto(外部 vision.js 可用则用, 否则 none)")
    ap.add_argument("--render-page", type=int, default=1,
                    help="渲染第几页供视觉阅读(默认 1; 0=自动找含表头的页)")
    ap.add_argument("--render-png", help="渲染输出 PNG 路径(仅 --vision 调试用)")
    ap.add_argument("--output", help="描述符 JSON 路径(默认 <pdf>_format_descriptor.json)")
    ap.add_argument("--password", help="PDF 打开密码")
    ap.add_argument("--debug", action="store_true")
    args = ap.parse_args()

    if not args.input:
        print("错误: 缺少 --input", file=sys.stderr)
        ap.print_usage(file=sys.stderr)
        sys.exit(2)
    try:
        from vision_utils import set_vision_provider
        resolved_provider = set_vision_provider(args.vision_provider)
    except Exception:  # noqa: BLE001
        resolved_provider = "none"
    mode = "heuristic" if args.heuristic else ("vision" if args.vision else "anchors")
    if mode == "vision" and resolved_provider == "none":
        print("错误: 当前视觉功能不可用/已禁用(none); 请改用 --heuristic 或 --anchors",
              file=sys.stderr)
        sys.exit(2)
    if mode == "anchors" and not args.anchors:
        print("错误: 请用 --anchors 提供表头列名, 或 --heuristic / --vision 自动生成",
              file=sys.stderr)
        sys.exit(2)
    try:
        desc = build_format_descriptor(
            args.input, password=args.password, mode=mode,
            anchors=args.anchors, vision=args.vision, debug=args.debug)
    except Exception as e:
        print(f"错误: {e}", file=sys.stderr)
        sys.exit(2)

    out_path = args.output or os.path.splitext(os.path.abspath(args.input))[0] + "_format_descriptor.json"
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(desc, f, ensure_ascii=False, indent=2)

    print(f"格式描述符: {out_path}")
    print(f"  日期形态: {desc['date_pattern']} (样例 {desc['date_sample']})")
    print(f"  布局: {desc['layout']} | 倒序: {desc['reverse_chronological']} "
          f"| 每页表头: {desc['header_every_page']} ({desc['header_pages_ratio']:.0%})")
    print(f"  页脚关键词: {desc['footer_keywords'] or '无'}")
    print("  列模板(锚点盒 → 中点边界 + 语义):")
    for c in desc["columns"]:
        print(f"    {c['name']:22s} x=[{c['x0']:7.1f}, {c['x1']:7.1f}) sem={c['semantic']}")
    if desc["anchors"]["total"]:
        print(f"  锚点反查: 命中 {desc['anchors']['matched']}/{desc['anchors']['total']}")
    print("下一步: python extract_bank_statement.py --input <pdf> --descriptor <json>")
    sys.exit(0)


if __name__ == "__main__":
    main()
