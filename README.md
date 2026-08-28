# bank2excel-h5

银行对账单 / 微信支付流水 PDF → Excel,纯前端本地转换。

**数据不出设备**:页面内置 Pyodide + PyMuPDF(WebAssembly)引擎,转换全程在你的浏览器内完成,无服务器、无上传、无埋点。

## 使用

托管后打开首页,拖入对账单 PDF 即可(支持多文件队列、加密 PDF 密码、失败时导出脱敏诊断包)。

- 首次访问需加载 ~30MB 引擎(Service Worker 缓存后二次访问秒开)
- 已适配:民生/建行/工行/招行/北京/华夏/北京农商/浦发/微信支付等十余种格式

## 开发

```bash
python -m http.server 8765
# 打开 http://127.0.0.1:8765/
```

- `tests/regression.html` 浏览器内全量回归(13 样本 vs 基准 fixtures)
- `python/m1_check.py` 本地 CPython 回归
- `.github/workflows/build-wheel.yml` PyMuPDF Pyodide wheel 云构建(备胎,本地已构建好 wheels/)

## 隐私

见 [privacy.html](privacy.html)。诊断包仅含脱敏版式骨架(金额→0.01, 账号→*), 是否导出由你决定。
