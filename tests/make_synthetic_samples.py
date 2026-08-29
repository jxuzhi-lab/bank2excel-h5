# -*- coding: utf-8 -*-
"""make_synthetic_samples.py —— M6: 生成无隐私的合成测试样本。

背景:
  真实对账单 PDF 无法脱敏后公开(对文本做 redact 会破坏 Form XObject 坐标系,
  导致基于坐标的表头识别失效; 字符替换又受限于内嵌字体子集)。方案定为:
  公开仓库只放**合成 PDF**(程序生成、标准表格版式), 真实样本回归仅本地保留。

合成样本设计(覆盖引擎两条提取分支):
  sample_bank.xlsx 版式 → 标准"表格型"对账单:
    标题行 + 表头行(交易日期/摘要/收入/支出/余额) + N 行数据(日期为记录锚)
    全部数据为随机生成(姓名=张三/李四, 账号=6222***, 金额=递增随机)
  sample_wechat.xlsx 版式 → 微信证明风格(竖排拆词表头简版)

输出:
  tests/pdf/         合成 PDF(覆盖原真实样本副本)
  tests/fixtures/    对应基准 JSON(由合成数据直出, 天然自洽)
"""
import io
import json
import os
import random
import sys

import pymupdf

HERE = os.path.dirname(os.path.abspath(__file__))
PDF_DIR = os.path.join(HERE, "pdf")
FXT_DIR = os.path.join(HERE, "fixtures")

random.seed(20260829)  # 可复现


def gen_bank_records(n):
    """生成 n 条合成银行记录(日期递增, 金额随机, 余额链一致)。"""
    recs = []
    balance = round(random.uniform(1000, 50000), 2)
    day = 1
    for i in range(n):
        day += random.randint(0, 2)
        date = f"2026-0{random.randint(1,3)}-{day:02d}" if day <= 28 else f"2026-04-{day % 28 + 1:02d}"
        t = f"{random.randint(8,20):02d}:{random.randint(0,59):02d}:{random.randint(0,59):02d}"
        amount = round(random.uniform(1, 5000), 2)
        direction = random.choice(["收入", "支出"])
        if direction == "收入":
            balance = round(balance + amount, 2)
            credit, debit = amount, ""
        else:
            balance = round(balance - amount, 2)
            credit, debit = "", amount
        recs.append({
            "date": date, "time": t,
            "summary": random.choice(["转账", "消费", "工资", "退款", "利息", "缴费"]),
            "name": random.choice(["张三", "李四", "王五", "赵六"]),
            "account": f"6222****{random.randint(1000,9999)}",
            "income": credit, "expense": debit,
            "balance": balance,
        })
    return recs


def draw_bank_pdf(path, records):
    """画标准表格型对账单 PDF(与引擎表头关键词对齐)。"""
    doc = pymupdf.open()
    page = doc.new_page(width=595, height=842)  # A4
    y = 60
    page.insert_text((160, y), "XX银行个人账户交易流水明细清单",
                     fontsize=14, fontname="china-s")
    y += 24
    page.insert_text((60, y), f"账号：6222****8888    户名：张三    币种：人民币",
                     fontsize=9, fontname="china-s")
    y += 20
    cols = [("交易日期", 60, 55), ("摘要", 120, 55), ("收入", 180, 55),
            ("支出", 240, 55), ("余额", 300, 60), ("对方户名", 365, 60),
            ("对方账号", 430, 90)]
    # 表头
    for name, x, w in cols:
        page.insert_text((x, y), name, fontsize=9, fontname="china-s")
    y += 14
    page.draw_line(pymupdf.Point(55, y - 9), pymupdf.Point(520, y - 9), width=0.5)
    # 数据行
    for r in records:
        vals = [r["date"], r["summary"], str(r["income"]), str(r["expense"]),
                f"{r['balance']:.2f}", r["name"], r["account"]]
        for (name, x, w), v in zip(cols, vals):
            page.insert_text((x, y), str(v), fontsize=8, fontname="china-s")
        y += 13
        if y > 800:  # 分页
            page = doc.new_page(width=595, height=842)
            y = 60
            for name, x, w in cols:
                page.insert_text((x, y), name, fontsize=9, fontname="china-s")
            y += 14
    doc.save(path, deflate=True)
    doc.close()


def gen_wechat_records(n):
    recs = []
    day = 1
    for i in range(n):
        day += random.randint(0, 2)
        date = f"2026-0{random.randint(1,3)}-{day % 28 + 1:02d}"
        t = f"{random.randint(8,20):02d}:{random.randint(0,59):02d}"
        amount = round(random.uniform(0.5, 800), 2)
        direction = random.choice(["收入", "支出"])
        recs.append({
            "date": date, "time": t,
            "type": random.choice(["转账", "二维码付款", "红包", "商户消费"]),
            "direction": direction,
            "counterparty": random.choice(["张三", "李四", "某商户", "王五"]),
            "amount": amount,
            "paymethod": random.choice(["零钱", "工商银行(7458)"]),
            "status": "已到账",
        })
    return recs


def draw_wechat_pdf(path, records):
    """微信证明风格: 签名标题 + 表头 + 数据行。"""
    doc = pymupdf.open()
    page = doc.new_page(width=595, height=842)
    y = 60
    page.insert_text((180, y), "微信支付交易明细证明", fontsize=14, fontname="china-s")
    y += 22
    page.insert_text((60, y), "兹证明：张三（居民身份证：110101********001X），在其微信号：wxid_test中的交易明细信息如下",
                     fontsize=8, fontname="china-s")
    y += 20
    cols = [("交易单号", 60, 80), ("交易时间", 145, 60), ("交易类型", 210, 55),
            ("收/支", 270, 35), ("交易对方", 310, 60), ("金额(元)", 375, 50),
            ("支付方式", 430, 65), ("当前状态", 500, 50)]
    for name, x, w in cols:
        page.insert_text((x, y), name, fontsize=8, fontname="china-s")
    y += 13
    txid = 420000260000000000
    for r in records:
        txid += random.randint(1, 9)
        vals = [str(txid), f"{r['date']} {r['time']}", r["type"], r["direction"],
                r["counterparty"], f"{r['amount']:.2f}", r["paymethod"], r["status"]]
        for (name, x, w), v in zip(cols, vals):
            page.insert_text((x, y), str(v), fontsize=8, fontname="china-s")
        y += 13
        if y > 800:
            page = doc.new_page(width=595, height=842)
            y = 60
            for name, x, w in cols:
                page.insert_text((x, y), name, fontsize=8, fontname="china-s")
            y += 13
    doc.save(path, deflate=True)
    doc.close()


def cells_from_pdf(path):
    """用引擎实际转换合成 PDF → 规范化 cells(与 make_fixtures 同语义)。
    引擎提取成功 = 合成样本对管道有效; 基准来自引擎输出, 验证的是
    '浏览器内转换 == 本地转换' 的一致性(而非格式识别)。"""
    from datetime import date, datetime
    here = os.path.dirname(os.path.abspath(__file__))
    sys.path.insert(0, os.path.join(os.path.dirname(here), "python"))
    os.environ["PYODIDE"] = "1"
    import shim
    import extract_bank_statement as eng
    eng.PYODIDE = True
    import openpyxl

    def norm(v):
        if isinstance(v, datetime):
            return v.strftime("%Y%m%d %H:%M:%S")
        if isinstance(v, date):
            return v.strftime("%Y%m%d")
        if v is None:
            return ""
        try:
            return str(float(v))
        except Exception:
            return str(v).strip()

    with open(path, "rb") as f:
        r = shim.convert_bytes(f.read())
    wb = openpyxl.load_workbook(io.BytesIO(r["xlsx"]))
    ws = wb.active
    cells = [[norm(c.value) for c in row] for row in ws.iter_rows()]
    return cells, r["rows"]


def main():
    os.makedirs(PDF_DIR, exist_ok=True)
    os.makedirs(FXT_DIR, exist_ok=True)
    manifest = []
    specs = [
        ("synthetic_bank_50.pdf", gen_bank_records, draw_bank_pdf, 50),
        ("synthetic_bank_300.pdf", gen_bank_records, draw_bank_pdf, 300),
        ("synthetic_wechat_60.pdf", gen_wechat_records, draw_wechat_pdf, 60),
    ]
    for i, (fname, gen, draw, n) in enumerate(specs, 1):
        recs = gen(n)
        p = os.path.join(PDF_DIR, fname)
        draw(p, recs)
        try:
            cells, rows = cells_from_pdf(p)
            slug = f"synthetic_{i:02d}.json"
            with open(os.path.join(FXT_DIR, slug), "w", encoding="utf-8") as f:
                json.dump({"rows": len(cells), "cols": max(len(x) for x in cells),
                           "cells": cells}, f, ensure_ascii=False,
                          separators=(",", ":"))
            manifest.append({"name": fname, "pdf": "pdf/" + fname,
                             "fixture": "fixtures/" + slug, "rows": len(cells)})
            print(f"OK  {fname}  {rows} rows extracted")
        except Exception as e:  # noqa: BLE001
            print(f"FAIL {fname}: {str(e)[:90]}")
    with open(os.path.join(FXT_DIR, "manifest.json"), "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=1)
    print(f"\n{len(manifest)}/{len(specs)} synthetic samples -> {PDF_DIR}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
