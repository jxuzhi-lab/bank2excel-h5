# bank2excel-h5 维护手册（2026-08-29 版）

> 本手册供新会话接手维护时阅读。读完本文即可了解：项目全貌、双通道部署、更新流程、踩坑记录、待办。
> 项目状态：**M0-M7 全部完成，双通道线上运行中**。

---


## ⚠️ 交接速览（2026-08-30 深夜, 长期无人接管前的最终状态）

**全部已上线并验证**：dual 通道 OCR（GLM-OCR 主路径 + 百度定向兜底，民生压测 401/401 满分）、
网页多文件队列 UI（上限 10/串行转换/每文件密码/删除/刷新持久化 IndexedDB）、
转换日志后台 `/admin`（ADMIN_PASSWORD 见 VPS compose；记录时间/文件/耗时/通道/三路 API token，
SQLite 存于挂卷 ./logs，保留 30 天自动清理；支持筛选/搜索/CSV 导出/30s 自动刷新）。

**长期无人接管注意事项**：
1. 三把 key 全在 VPS compose：智谱(GLM-OCR+VLM)、百度(体验额度约余 900 次≈900 页)。
   百度额度用完后自动只剩 GLM-OCR 主路径（功能不中断，质量上限降为"缺页标记"级）。
2. 智谱 600 万 token 赠送额度：按当前用量够数年。用尽后 glm-ocr 转按量 0.2 元/百万。
3. 磁盘占用：logs 挂卷（30 天自动清理）+ descriptors 挂卷（每格式几 KB）+ 容器日志
   `docker logs`（uvicorn 访问日志会持续增长，长期不接管可定期
   `sudo truncate -s 0 $(sudo docker inspect --format='{{.LogPath}}' bank2excel)`）。
4. 重启/重建不丢数据：descriptors 与 logs 都是挂卷。但重建后检查
   `sudo chown -R 1000:1000 /opt/bank2excel-h5/{logs,descriptors}`（root 挂载目录坑）。
5. 出问题先看：`sudo docker logs --tail 100 bank2excel` + `/admin` 失败行 + 诊断包。

---

## ⚠️ 新环境速览（2026-08-29 迁移至本机维护，先读这段）

- **Python**：系统 `python` 只有 3.7.8（过旧，缺 pymupdf≥1.24，跑不了引擎，勿直接用）。
  维护统一用项目 venv：`C:\Users\Administrator\Documents\银行对账单转化pdf\.venv`
  （Python 3.12.10，依赖与 requirements-server.txt 锁定版完全一致；m1_check 13/13 已验证）。
- **git 仓库镜像**：已从 `/tmp/repo_work` 迁至 **`repo_work/`（项目根目录下）**——/tmp 会被系统清理，仓库放那里随时可能丢。
  注意：推送凭据（GitHub PAT）明文存在 `repo_work/.git/config` 的 remote URL 里，勿截图/外泄该文件。
- **旧环境残留**：`C:\Users\Administrator\.workbuddy\binaries\python\envs\default`（Python 3.13.14，依赖非锁定版本），历史服务进程曾用它，维护时**勿依赖**。
- 本机 8766 服务已改由 `start_server.bat` → 项目 .venv 接管（2026-08-29 已重启验证）。
  server.py 内嵌 HTML 的 SyntaxWarning（JS 正则 `\.`）已用 r""" 前缀修复。

---

## 一、项目一句话

银行对账单/微信支付流水 **PDF → Excel** 转换工具。同一套 Python 引擎（`extract_bank_statement.py`），两条线上通道：

| 通道 | 访问地址 | 技术 | 适用 |
|------|---------|------|------|
| **A. GitHub Pages（浏览器内）** | https://jxuzhi-lab.github.io/bank2excel-h5/ | Pyodide WASM，纯前端 | PC/手机临时用，数据不出设备 |
| **B. 私有 VPS（服务端）** | http://124.223.110.222 | FastAPI + Docker，本地 CPython | 手机稳定使用（推荐主通道） |

- 仓库：`github.com/jxuzhi-lab/bank2excel-h5`（public），Pages 指向 main 分支
- 引擎识别 13 种格式（民生/建行/工行/招行/北京银行/华夏/北京农商/浦发/微信支付×2/邮储/8264/旅立方等），本地回归 **13/13 全绿**

---

## 二、代码结构（bank2excel-h5/）

```
python/
  extract_bank_statement.py   # 引擎(3638+ 行, 唯一真源)
  shim.py                     # PYODIDE 字节流桥接层(convert_bytes 入口)
  m1_check.py                 # 本地 13 样本回归(must 13/13)
  requirements-server.txt     # 服务端依赖(固定版本)
src/
  worker.js                   # Web Worker: Pyodide 加载 + 8 阶段上报 + wheel 主备源
  bridge.js                   # 主线程↔Worker 协议
  app.js                      # UI 逻辑(上传/进度/下载/诊断)
  diagnose.js                 # 诊断包导出
  download.js                 # xlsx 下载
wheels/
  pymupdf-1.26.7-cp312-abi3-pyodide_2024_0_wasm32.whl  # 17MB WASM wheel(勿删)
server.py                     # FastAPI 私有服务(通道 B)
Dockerfile / docker-compose.yml / requirements-server.txt
deploy_vps.md                 # VPS 购买+部署指南
deploy_vps.py                 # 一键远程部署脚本(paramiko)
sw.js / manifest.json / icon-*.png   # PWA
tests/                        # 合成样本回归集(公开) + 生成脚本
start_server.bat              # 本机服务一键启动(8766)
```

---

## 三、引擎同步 SOP（技能 ↔ H5 双副本，最易踩坑）

存在两份引擎代码：
- **真源**：`C:/Users/Administrator/Documents/银行对账单转化pdf/scripts/extract_bank_statement.py`（本地技能版，日常更新在这）
- **H5 副本**：`bank2excel-h5/python/extract_bank_statement.py`（带 PYODIDE 改造）

**技能更新 → H5 的固定流程**：
```bash
# 1. 以技能新版为基底覆盖
cp scripts/extract_bank_statement.py bank2excel-h5/python/extract_bank_statement.py
# 2. 重新施加 PYODIDE 补丁(在 H5 副本上, 按注释标记逐点重放):
#    - import io; PYODIDE 环境分支 + LOG_BUFFER/PROGRESS_CB/set_progress_cb/open_pdf
#    - log() PYODIDE 分支
#    - 3 处 pymupdf.open → open_pdf
#    - 诊断 PNG 落盘 PYODIDE 跳过
#    - sheet 名 PYODIDE 分支 / write_meta 跳过
#    - vision 4 处短路(_vision_fallback_extract/classify_doc/quick_classify/_auto_onboard_retry)
#    - 页级进度钩子 PROGRESS_CB(pno+1, page_count)
# 3. 回归验证(必须全绿):
cd bank2excel-h5/python && python m1_check.py   # 期望 13/13 PASS
# 4. 推送: git add + commit + push(仓库在 项目根/repo_work)
```

---

## 四、更新部署流程

### 通道 A（GitHub Pages）
```bash
cd "C:/Users/Administrator/Documents/银行对账单转化pdf/repo_work"   # 本地仓库镜像(勿放回 /tmp)
cp <修改文件> .     # 覆盖改动
git add -A && git commit -m "..." && git push origin main
# Pages 自动构建, 约 1-3 分钟生效; 验证: curl https://jxuzhi-lab.github.io/bank2excel-h5/
```

### 通道 B（VPS）
```bash
# 方式一(代码改动): 上传后 compose 重建
#   Dockerfile/python/ 改动 → 重新构建镜像
ssh ubuntu@124.223.110.222
cd /opt/bank2excel-h5 && sudo docker compose up -d --build
# 方式二(仅数据/配置): 容器内已 COPY 的代码改了要重建, 日志看 docker logs -f bank2excel
```
> VPS 上代码更新：**推荐 SFTP 少量文件**（`update_vps.py` 仅本地存在, 含 codeload 流程; 实测 codeload 从 VPS 下载被限速 ~32KB/s, 整包 zip 极慢, 变更文件少时直接 SFTP 覆盖 + `docker compose up -d --build` 最快）。用 `update_vps.py`（依赖 deploy_vps.py 的凭据, 两者都不入仓库）或手动 SFTP 后 sudo cp。

---

## 五、关键踩坑记录（务必先读）

1. **micropip API 兼容（血泪教训）**：Pyodide 0.27.2 内置 micropip **只有最简签名可用**。
   - `micropip.install(url)` ✅；`keep_going=` / `pre=` / `add_wheel_log_handler()` ❌（会抛 unexpected keyword argument / AttributeError，且报错极隐蔽）
   - 之前"移动端加载失败"排障 3 小时，真凶就是这行 API，与缓存/SW 无关
2. **移动端 WeChat X5 掐断 17MB 下载**：微信内置浏览器对跨域大文件 fetch 有限制（AbortError）。**无法代码根治**，只能引导"用 Chrome 打开"或走 VPS 通道。手机 Firefox/Chrome 正常。
3. **Service Worker 缓存坑**：SW 会缓存旧 worker.js 导致"改了不生效"。措施：
   - `src/worker.js` 走 **network-first**（永不缓存）
   - SW 缓存版本 bump 时改 `sw.js` 顶部 `CACHE` 常量 + `index.html` 注册 `?v=N`
   - app.js 有 autoResetOldSW 兜底（启动时 unregister 旧 SW）
4. **VPS 部署链路的网络坑**（国内服务器）：
   - GitHub 主站/`git clone` 超时 → 用 `codeload.github.com/.../zip/refs/heads/main`（该域名可达）
   - Docker Hub 超时 → 配镜像加速器 `/etc/docker/daemon.json`：`mirror.ccs.tencentyun.com`（腾讯云内网）
   - `fonts-noto-cjk` 拉取极慢 → 服务端不需要字体渲染，Dockerfile 已移除
   - **systemd 起 docker 需先起 containerd**：`systemctl start containerd` 再 `start docker`
   - ubuntu 用户已免密 sudo（sudo -n true 可用），部署脚本直接 sudo 即可
5. **本地服务端口**：8766（本机）、VPS 80。VPS 曾因旧进程占端口报 winerror 10048。
6. **隐私策略**：公开仓库 tests/ 只放**合成样本**（make_synthetic_samples.py 生成，seed=20260829）；真实样本仅本地（`测试样本/` 目录，13 个 PDF + 基准 xlsx）不入库。

---

## 六、当前运行状态

| 项 | 状态 |
|----|------|
| GitHub Pages 站点 | ✅ 在线（engine/worker/wheel 均 200） |
| VPS 服务 | ✅ 在线（124.223.110.222:80，**T1 完整引擎版 2026-08-29 部署**，vision 未配 key 时为 none） |
| 本机 8766 服务 | ✅ 运行中（start_server.bat / .venv，T1 完整引擎版） |
| 引擎回归 | ✅ 13/13（m1_check.py） |
| 合成样本回归 | ✅ 3/3 |

---

## 七、待办（未做）

1. **M4 诊断包闭环演练**：导出诊断包 → 开发者侧导入分析流程，未实操验证
2. **VPS 安全**：用户改 SSH 密码（已提醒）；可选加 token 鉴权（server.py 加装饰器）
3. **域名 + HTTPS**（可选）：IP 访问无法申请免费证书；需域名+ICP 备案，或用 Cloudflare Tunnel
4. **README 完善**：仓库 README.md 可补双通道说明
5. ~~VPS Dockerfile 已改未推 GitHub~~ → **已解决（2026-08-29 验证）**：commit `b0e3f8c`（Dockerfile 改造 + 本手册）已推送，本地 main 与 origin/main 一致。

---

## 八、T1 未知格式识别（2026-08-29 新增，服务端兜底架构）

**核心机制**（详见方案讨论记录）：未知格式识别放 VPS 端，识别本身仍由规则管道做，VLM 只做"每文件一次读表头列名"的高层判断。`server.py` 已从 shim/PYODIDE 版升级为**完整引擎**（`eng.convert_pdf`，非 PYODIDE 路径），失败自动走两级升级链：

1. 表头识别失败 → 视觉兜底（读列名 → 锚点反查 → 规则管道重提）
2. 兜底成功 → 锚点补全描述符 → **真实提取校验通过才落缓存**（防坏描述符污染）
3. 同格式后续请求 → 首页文字指纹（sha256）命中缓存 → **零视觉成本直取**
4. 缓存重试失败（版式漂移）→ 强制刷新缓存重走兜底

**视觉 provider**（`vision_utils.py`，四态）：`visionjs`（外部 node 脚本，开发机残留可用）/ `model`（代理读图回环）/ **`api`（新增：直连 OpenAI 兼容视觉接口，标准库 urllib，无新增依赖）** / `none`（纯规则）。auto 顺序：显式 env → visionjs 可用 → api 配置齐全 → none。

**VPS 启用 VLM 兜底**（docker-compose.yml 已留注释模板，三行配上即生效）：
```yaml
- BANK_PDF_VISION_PROVIDER=api
- BANK_PDF_VISION_API_BASE=https://open.bigmodel.cn/api/paas/v4   # 或其它 OpenAI 兼容端点
- BANK_PDF_VISION_API_KEY=sk-xxx
# 可选: BANK_PDF_VISION_API_MODEL(默认 glm-4.6v)
```

**成本/安全控制**（server.py，env 可调）：`VLM_BUDGET_PER_HOUR=40`（超限自动降级纯规则，包装在 call_vision_raw 唯一出口）/ `RATE_LIMIT_PER_MIN=12`（每 IP）/ `MAX_CONCURRENT=2`。**隐私**：文件转换完即删；仅视觉启用且规则失败时，第 1 页渲染图才发给所配 VLM；描述符只含列模板不含数据。

**VPS 端已启用**（2026-08-29）：智谱 key 已配置在 **VPS 本地** `/opt/bank2excel-h5/docker-compose.yml`（model=glm-4.6v, 预算 40 次/时），**仓库里的 compose 只留占位注释, key 不入 git**。启用后线上已实测：未知格式合成 PDF → glm-4.6v 读表头 → 转换成功（3.4s）→ 描述符入卷缓存 → 二次请求命中缓存零视觉成本（1.6s）。

**运维坑（已修, 记录防回退）**：`./descriptors` 卷目录若被 Docker 首次自动创建会归 root 所有, 容器内 `app` 用户(uid 1000)写不进（报 Permission denied, 转换仍成功但不缓存）。修复：`sudo chown -R 1000:1000 /opt/bank2excel-h5/descriptors`。

**扫描件 OCR（2026-08-29 新增, 服务端专属）**：`ocr_layer.py`（仓库根, server 专属模块）实现"夹心层"方案——无文字层的页渲染位图 → RapidOCR（rapidocr_onnxruntime 1.4.4, CPU 推理, 模型进程内懒加载）逐行识别 → 识别文本按检测框坐标以**不可见文字**(render_mode=3)写回 PDF → 引擎零改动直接转换。server.py 在 `OCR_ENABLED=1`（默认开）时自动走此路径, 成功响应带 `X-OCR-Pages` 头。**实测**（北京银行 3 页扫描件 180dpi JPEG）：58/58 行恢复, 交易日期/发生额/余额三列与登记基准 100% 一致, 金额合计分毫不差；已知边界：单字符列（钞汇）与脱敏 `*` 号跨列会漂移（行级 OCR 框固有限制）。性能：~8s/页（含模型首次加载）, 容器内存峰值 ~330MB。

**OCR 部署坑（逐个踩过）**：
- Dockerfile 新增模块（ocr_layer.py）**必须加 COPY 行**, 否则容器 import 失败无限重启（Restarting 循环）
- `deb.debian.org` 国内不稳定会卡死 apt → Dockerfile sed 换 `mirrors.tuna.tsinghua.edu.cn`（trixie 的源在 /etc/apt/sources.list.d/debian.sources）
- opencv 需要 `libgl1 libglib2.0-0`（apt 已加）
- Docker workers 2→1：OCR 模型每 worker 独立加载 ~500MB, 2G 小鸡扛不住双份；单 worker + MAX_CONCURRENT=2 线程并发够用
- SSH 断连后 `setsid docker compose up -d --build &` 会留孤儿进程, 多个 up 并发报"容器名已被占用"互相打架 → 先 `pkill -9 -f 'docker compose'` 清场再触发

**水印复杂样本 OCR 压测结论（2026-08-29, 工商22页/244水印词 + 民生20页/回归基准）**：管道全程不崩, 但质量显著退化, 四个坑：
1. **水印文字泄漏进表头带**（最致命, 民生实锤）：文字层版的水印靠引擎"跨页同位置固定元素"清除, 要求 text+x+y 精确匹配；OCR 坐标逐页抖动 → 判定失效, 页首行"中国民生银行…香山支行"被当成表头一部分 → 表头多出一列, 列错位级联全部行（借方合计虚增 714 万）。
2. **合并表头词被 OCR 拆散**（民生"交易时间|摘要"拆成两列, 11 列 vs 基准 10 列；凭证号码时而并入摘要时而独立 → 行间列跳变）。
3. **丢行**（民生 -12、工商 -16）：同日多条记录的日期词 OCR 失败或被并 → 整条记录被吸入上一行。
4. 反直觉：工商对角印章/防伪码**零泄漏**（RapidOCR 读不了斜排文字反而因祸得福）, 其损害主要来自表头列结构错位（10 列 vs 11 列, 余额列 0/525）。水印密度不决定成败（北京银行 111 词/3 页曾 100% 通过）, **水印/页眉是否落在表头带附近才是关键**。
→ 改进方向（未做）：ocr_layer 内做跨页水印行抑制（相似文本+近似位置的行不写入夹心层）；表头列数与启发式不一致时用 VLM 复核；把"笔数校验"结果暴露给用户提示丢行风险。

→ 改进方向（未做）：ocr_layer 内做跨页水印行抑制（相似文本+近似位置的行不写入夹心层）；表头列数与启发式不一致时用 VLM 复核；把"笔数校验"结果暴露给用户提示丢行风险。

**GLM-OCR 备选方案实测（2026-08-30, 未集成）**：智谱 `glm-ocr` 走专用端点 `POST /paas/v4/layout_parsing`，`file` 传 `data:image/jpeg;base64,...`（裸 base64/文件上传均不支持, files 接口 purpose=ocr 报错）。实测北京银行 3 页扫描件（2.5x JPEG）：**平均 4704 tokens/页**（prompt~2650 + completion~2050）, 8-14s/页, 返回结构化 markdown（含印章文字, **无坐标**——夹心层方案用不了, 集成需 md→记录转换层）。定价 0.2 元/百万 token（输入输出同价）；用户持有 8 元/5000 万 token 资源包（0.16 元/M）≈ **0.00075 元/页, 整包约 1 万页**。牌价下 0.00094 元/页。注意 token 消耗与渲染分辨率正相关, 降 zoom 可省钱。

**GLM-OCR 水印样本全量实验（2026-08-30, 民生 20 页扫描件 = RapidOCR 失败的同一样本）**：
- **水印完胜**：水印/开户机构文字只出现在表外文档文本, 表格内 0 污染, 表头 10 列与基准完全一致——夹心层的头号坑（水印进表头带→列错位级联）被根治。
- **表格质量高**：359 条记录零脏行零多行, 金额格式干净, HTML `<table>` 解析容易。
- **新坑——页级输出截断**：2/20 页（10%）生成退化, completion 跑满 ~8289 token 内部上限, `<table>` 无闭合、末行切在流水号中间; temperature=0 下跨 zoom/跨 max_tokens 参数都确定性复现, **不可重试修复**。整页静默丢失 42 笔（借方合计虚差 149 万）。
- **缓解路径明确**：截断可检测（闭合标签+completion 阈值）→ 失败页回退 RapidOCR 夹心层。推荐架构：**GLM-OCR 主路径（含逐页校验）+ 夹心层兜底 + md→记录转换层**（HTML 表格解析已验证可行）。
- 运维注意：并发 4 会触发限流（部分页失败需补跑）, 客户端要限速; 全实验含重试共耗 ~21 万 token ≈ 0.034 元。

**百度智能云表格识别对比实测（2026-08-30, 未集成, 与 GLM-OCR 同样本对垒）**：接口 `POST /rest/2.0/ocr/v1/table?access_token=`（OAuth client_credentials 取 token, 表单 form 传 base64 image, 免费档 QPS≈2 需客户端节流）。返回**单元格级网格**（row/col 索引 + cell_location 坐标 + words）。民生 20 页：**401/401 行零丢失、借/贷/余额三值 401/401 全对、零截断零失败, 59s（QPS 节流下）**；北京 3 页 58/58 全对。已知瑕疵：竖排表头词被网格合并（"交易时间 摘要"成一列）→ 行列数 8/9/10 参差, 集成需列重排层；个别行余额并入流水号格（值仍在, 可被余额链校正）；11 条长截断摘要边缘字数差。对比结论：**数据质量三方案最优**（RapidOCR 夹心层列错位污染 / GLM-OCR 确定性截断丢行 / 百度表格全对）, 且单元格级坐标支持"cell 级夹心"集成路线。限制：体验额度 1000 次（=1000 页）, 正式计费走 OCR 共享资源包（¥9.9/1万点首购, 12 个月有效）, **表格文字识别V2 为 25 点/次 = 0.02475 元/页, 是 GLM-OCR 资源包价的 33 倍** → 定位应为"精锐兜底"而非主力：接在 GLM-OCR 余额链断点检测后, 仅对断点页(2-3 页)送百度重做(≈0.05 元/份), 全单成本仍 ~0.02 元级, 质量上限拉到满分级。省点子方向(未验证)：tables_result 支持多表格数组 + 图片长边上限 8194px → 2.5x 下两页竖拼一次调用可减半点数, 需实测跨页粘连风险。


**Dual 通道已上线（2026-08-30 部署 VPS 并线上满分验收）**：`baidu_ocr.py`（新）+ `glm_ocr.py` 升级。架构：GLM-OCR 主路径 → 余额链断点检测 → GLM 定向重试（换 zoom 两轮）→ 仍断点则**百度表格识别精锐兜底**（仅断点页, 封顶 BAIDU_OCR_MAX_PAGES=4 页/文档）。断点页定位用**内容锚定**（断点前行签名所在页, 累计行数映射会漂移——踩坑）。百度结果经**直接网格重组**（表头包含关系映射 + 行形态拆分, 不经引擎——cell 级夹心+引擎提取路线实测脆弱已弃）合入, 配套四层修复：粘连行拆分（日期+摘要偶发粘连）、页缝相邻去重（键=(日期数字串, 余额数字串)——**禁用 float**, 20 位流水号超 float64 精度会误删合法行——踩坑；余额为空的行不参与折叠——踩坑）、百度页空余额链式补全、错误余额链式纠错（仅百度来源行, 金额实测高度可靠）。**民生压测终局：401/401 行、借贷合计分毫不差、逐行余额 401/401、余额链完整**（RapidOCR 389 行列污染 / GLM 单通道 381 行缺 20）。竖拼实测否决：单页三值 21/21 → 2 页拼 37/41 → 3 页拼 17/63, 服务端对超高图内部降采样, 质量随图高单调劣化。配置：BAIDU_OCR_API_KEY/SECRET_KEY（未配置自动跳过百度级, 降级为 GLM 单通道）。VPS compose 已配 BAIDU_OCR_API_KEY/SECRET_KEY（百度体验额度 1000 次约余 900）；线上实测民生扫描件 401/401 + 借贷合计分毫不差 + 逐行余额 401/401, 模式头 glm-ocr+baidu。**Dockerfile 已加 COPY glm_ocr.py / baidu_ocr.py**（漏 COPY 会容器崩溃循环——已知坑）。


**PaddleOCR-VL 文档解析实测（2026-08-30, 未集成）**：接口 `POST /rest/2.0/brain/online/v2/paddle-vl-parser/task`（异步: 提交返回 task_id → `/task/query` 轮询 5-10s 起; form 表单, 参数 `file_data`=base64 + `file_name` 必填）。**整份 PDF 一次调用**（≤500 页/100M）。民生 20 页水印扫描件实测：单调用 64s 出全部 20 页; 内容存在性 **396/401（约 99%）**, 缺一个 4 行连续块+1 行（恰是历来最难的 3 月 16 日页边界区）; 输出 markdown 表头 10 列全分离（无表格 V2 的合并表头问题）, 另有 cells+matrix 结构与 span_boxes 行坐标（需开 return_span_boxes）。坑：markdown 的 `|---|` 分隔行与页边界重复行需解析层清洗（437 原始行→清洗后对位评估才有意义）。成本 0.009 元/页（9 元/1000 页）, 为 GLM-OCR 的 12 倍、表格 V2 的 36%。**结论：不替代 GLM-OCR 主路径（贵 12 倍）; 作兜底通道与表格 V2 二选一——小额定向兜底（2-3 页）表格 V2 更便宜且质量满分, 整文档级重做（>8 页）PaddleOCR-VL 更划算且集成简单（单调用）+ 免费跨页合并 merge_tables/水印擦除 erase_watermark 等开关未实测。当前架构维持 GLM-OCR 主 + 表格 V2 定向兜底不变。**

**关键文件**：
**关键文件**：
- `python/vision_utils.py`、`python/onboard_format.py`：从 scripts/ 真源同步（vision_utils 含 api provider）。**改 scripts/ 后记得 cp 到 python/**（引擎双副本 SOP 的扩展）
- `tests/test_vps_fallback.py`：端到端冒烟（自建假 VLM 端点 + 合成未知格式 PDF），8 场景断言升级链/缓存/协议形态，`python tests/test_vps_fallback.py` 即可跑
- 描述符缓存：VPS 挂卷 `./descriptors:/data/descriptors`（重建不丢）；本机服务默认 `bank2excel-h5/_descriptors/`
- 描述符回流（未做）：VPS 缓存审核后入库 → H5/Pages 通道也能消费描述符（纯规则数据）

**已知行为**：无视觉配置时未知格式返回 422 结构化错误（stage/suggestion + 诊断包 JSON，前端可一键下载），替代原来的裸错误文本。

---

## 九、快速参考

```bash
# 本地回归(用项目 venv, 系统 python 3.7 跑不了)
cd bank2excel-h5/python && "C:/Users/Administrator/Documents/银行对账单转化pdf/.venv/Scripts/python.exe" m1_check.py

# T1 兜底链路端到端冒烟(自建假 VLM, 无需真实 key)
cd bank2excel-h5 && "C:/Users/Administrator/Documents/银行对账单转化pdf/.venv/Scripts/python.exe" tests/test_vps_fallback.py

# 本机服务
cd bank2excel-h5 && start_server.bat   # 或 python server.py --port 8766

# VPS 运维(ssh ubuntu@124.223.110.222, 免密 sudo)
docker ps; docker logs -f bank2excel; docker restart bank2excel
cd /opt/bank2excel-h5 && sudo docker compose up -d --build   # 重建

# 移动端入口
http://124.223.110.222        # 主用(VPS)
https://jxuzhi-lab.github.io/bank2excel-h5/  # 备用(Pages)

# 测试样本(本地真实, 不入库)
C:/Users/Administrator/Documents/银行对账单转化pdf/测试样本/*.pdf
```
