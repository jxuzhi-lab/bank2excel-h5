# -*- coding: utf-8 -*-
"""mask_pdfs.py —— M6 脱敏 PDF 生成(方案 D, 保留提取管道兼容性)。

规则(逐词 redact):
  - 日期/时间词: 保留(记录切分锚点, 非隐私)
  - 纯数字金额词: 数字→0.01(保留正负号)
  - ≥8 位数字词(账号/流水号): 等长 '*'(数字字体字形可用)
  - 表头候选词(与已知列名词表匹配): 保留(表头检测依赖)
  - 其余中文/混合词: 整词删除(redact 不替换文本 → 留白, 无字形问题)

用法: python mask_pdfs.py <源目录> <输出目录>
"""
import glob
import os
import re
import sys

import pymupdf

MONEY = re.compile(r'^[-+]?(\d{1,3}(?:,\d{3})+|\d+)(\.\d{2})$')
LONGNUM = re.compile(r'^\d{8,}$')
DATEISH = re.compile(r'^[12]\d{3}([-/.年月日时: ]\S*)?$')
TIMEISH = re.compile(r'^\d{1,2}:\d{2}(:\d{2})?$')
CN_RE = re.compile(r'[\u4e00-\u9fff]')

# 表头词表(与引擎 DEFAULT_HEADER_KEYWORDS 对齐 + 常见列名)
HEADER_WORDS = {
    "交易时间", "时间", "日期", "交易日期", "记账日期", "摘要", "业务摘要", "用途", "备注",
    "附言", "凭证种类", "凭证类型", "凭证号码", "凭证号", "序号", "交易序号", "流水号",
    "借方发生额", "贷方发生额", "借方", "贷方", "支出", "收入", "交易金额", "金额",
    "发生额", "余额", "联机余额", "账户余额", "币种", "钞汇", "钞汇标志", "对方户名",
    "对方账号", "对方账户", "收付款人", "交易地点", "交易机构", "交易渠道", "交易类型",
    "货币", "单位", "户名", "账号", "账户", "对手信息", "对方信息", "现转标志",
    "交易渠道/交易机构", "记账金额", "姓名", "客户姓名",
    # 微信/招行等竖排拆词表头的列名段
    "交易单号", "商户单号", "交易对方", "交易方式", "当前状态", "交易类型",
    "收/支", "收/支/其他", "金额(元)", "支付方式", "交易日", "交易时间/交易类型",
    "交易时间交易类型收/支/其他交易方式金额(元)",
    # 各银行补充列名
    "账户明细编号", "交易介质编号", "记账日期", "记账时间", "对手户名", "对手账号",
    "用途/摘要", "交易摘要", "转账金额", "手续费", "纸币/硬币", "余额(元)", "发生额(元)",
    "现钞/现汇", "账/卡号", "网点号", "柜员号", "授权码", "附言/用途",
    "收入金额", "支出金额", "本期合计", "上期余额", "本期余额",
}

# 文档分流/结构签名词(引擎靠它们识别文档类型, 必须保留)
SIGNATURE_WORDS = {
    "微信支付交易明细证明", "微信支付", "交易明细对应时间段", "具体交易明细",
    "兹证明", "居民身份证", "微信号", "交易明细信息如下",
    "北京银行个人客户交易流水清单", "客户账户历史明细", "账户历史明细",
    "历史明细", "交易流水清单", "账户明细", "交易流水明细清单",
    "个人客户交易流水", "对账单", "明细清单", "交易明细表",
}
# 竖排拆词合并表头(整词出现时保留)
HEADER_WORDS |= SIGNATURE_WORDS


def classify(text):
    """返回 'keep' | 'mask_money' | 'mask_num' | 'drop'"""
    t = text.strip()
    if not t:
        return "keep"
    # 签名词/结构词优先: 含签名子串的词(标题行/兹证明长句)必须保留,
    # 否则引擎的文档分流(is_wechat_doc 等)失效
    for sig in SIGNATURE_WORDS:
        if sig in t:
            return "keep"
    if t in HEADER_WORDS:
        return "keep"
    if DATEISH.match(t) or TIMEISH.match(t):
        return "keep"
    if MONEY.match(t):
        return "mask_money"
    if LONGNUM.match(t):
        return "mask_num"
    # 竖排拆词表头词(多列名连写, 含至少 2 个表头段): 保留
    # 例: "交易时间交易类型收/支/其他交易方式金额(元)"
    hits = sum(1 for hw in HEADER_WORDS if hw in t)
    if hits >= 2 and CN_RE.search(t):
        return "keep"
    # 含金额/账号的混合词(如 "转账1,234.56元"): 整词处理
    if re.search(r"\d{6,}", t):
        return "drop"
    if re.search(r"\d+\.\d{2}", t):
        return "mask_money"
    return "drop"


def mask_pdfs(src_dir, dst_dir):
    os.makedirs(dst_dir, exist_ok=True)
    stat = []
    for p in sorted(glob.glob(os.path.join(src_dir, "*.pdf"))):
        doc = pymupdf.open(p)
        for page in doc:
            for w in page.get_text("words"):
                x0, y0, x1, y1, text = w[0], w[1], w[2], w[3], w[4]
                act = classify(text)
                if act == "keep":
                    continue
                if act == "mask_money":
                    repl = ("-" if text.strip().startswith("-") else "") + "0.01"
                elif act == "mask_num":
                    repl = "*" * len(text.strip())
                else:
                    repl = ""
                page.add_redact_annot((x0, y0, x1, y1), text=repl)
            page.apply_redactions()
        out = os.path.join(dst_dir, os.path.basename(p))
        doc.save(out, garbage=3, deflate=True)
        doc.close()
        stat.append(os.path.basename(p))
    return stat


if __name__ == "__main__":
    src = sys.argv[1] if len(sys.argv) > 1 else r"C:/Users/Administrator/Documents/银行对账单转化pdf/测试样本"
    dst = sys.argv[2] if len(sys.argv) > 2 else os.path.join(
        os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "tests", "pdf_masked")
    done = mask_pdfs(src, dst)
    print(f"masked {len(done)} pdfs -> {dst}")
