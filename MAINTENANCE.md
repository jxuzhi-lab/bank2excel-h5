# bank2excel-h5 维护手册（2026-09-01 版）

> 本手册供新会话接手维护时阅读。读完本文即可了解：项目全貌、双通道部署、更新流程、踩坑记录、待办。
> 项目状态：**全功能运行中——规则管道 13+ 格式 / dual OCR 通道 / 异步任务模式 / 日志后台 / MCP 服务**。

---

## ⚠️ 交接速览（2026-09-01 最终状态，覆盖此前速览）

**已上线并全量验证的能力栈**（自上而下依次兜底）：
1. **规则管道**：13+ 家银行/微信格式文字层 PDF 直接转换；未注册新格式经 VLM 学表头→描述符缓存→秒转
2. **扫描件 OCR dual 通道**：GLM-OCR 主路径（逐页识别+截断自愈+余额链校验）→ 断点页百度表格 V2 定向重做（≤4 页/文档）→ RapidOCR 夹心层保底（无百度 key 时）
3. **异步任务模式**（移动端治本方案）：`POST /api/tasks` 提交即回 → `GET /api/tasks/{id}` 轮询 → `/result` 下载（结果保留 6h）；原同步 `/api/convert` 保留（MCP/curl 用）
4. **网页 UI（v8）**：多文件队列（上限 10、串行转换、每行独立密码、单行删除、✕ 重试、IndexedDB 刷新持久化、帮助弹窗、无自动下载弹窗）
5. **日志后台 `/admin`**：密码 `-TfLutI-9XZS`（VPS compose `ADMIN_PASSWORD`）；SQLite 挂卷存 30 天自动清理；时间/文件/耗时/通道/三路 API token/错误；CSV 导出
6. **MCP 服务**：`mcp-server/`（独立 venv，mcp<2），已注册本机 `~/.zcode/cli/config.json`，4 工具（convert_pdf/conversion_logs/conversion_stats/service_health），指向 VPS

**凭据与额度**（全在 VPS `/opt/bank2excel-h5/docker-compose.yml`）：
- 智谱 key：GLM-OCR + VLM 共用；600 万赠送 token 按当前用量够数年
- 百度 key：表格 V2 体验额度 1000 次约余 900（每次=1 页）；用尽自动降级 GLM 单通道（不中断）
- `ADMIN_PASSWORD=-TfLutI-9XZS`（后台登录；MCP 端对应 `B2X_ADMIN_PASSWORD`，已配在本机 ~/.zcode/config.json）

**长期无人接管注意事项**：
1. 磁盘：logs 挂卷（30 天自动清理）+ descriptors 挂卷（每格式几 KB）+ `_task_results/`（6h 自动清理）
   + 容器日志 `docker logs` 会持续增长（可定期
   `sudo truncate -s 0 $(sudo docker inspect --format='{{.LogPath}}' bank2excel)`）
2. 重建容器后检查 `sudo chown -R 1000:1000 /opt/bank2excel-h5/{logs,descriptors}`（root 挂载目录坑）
3. 移动端兼容：转换走异步任务模式，锁屏/切后台/断网均不丢结果（FirefoxFM 等杀连接浏览器已治本）
4. 出问题三步：`sudo docker logs --tail 100 bank2excel` → `/admin` 失败行 → 用户侧诊断包

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


## 十、开发会话记录（2026-08-29 ~ 09-01，本环境从接手到交付全程）

### D1（08-29）环境迁移 + T1 未知格式识别
- 环境接手：装 Python 3.12.10 + 项目 .venv（系统 3.7.8 跑不了引擎）；仓库镜像从 /tmp 迁至 repo_work/；
  修正手册过时待办（Dockerfile 改造其实已推送）
- T1 架构：server.py 从 shim/PYODIDE 阉割版改为**完整引擎**；vision_utils 加 `api` provider
  （OpenAI 兼容视觉接口，标准库 urllib）；描述符缓存加**校验门禁**（提取成功才落缓存）+
  **缓存优先**（兜底前先查缓存，同格式零视觉成本）；docker-compose 挂 descriptors 卷
- 8 场景端到端冒烟（tests/test_vps_fallback.py，假 VLM+合成未知格式 PDF）全绿；部署 VPS

### D2（08-29）VLM 启用 + 描述符回流前置
- 智谱 key 配上（glm-4v-flash→后改 glm-4.6v）；线上实测 VLM 兜底全链路；
- 修坑：descriptors 卷 root 属主导致写缓存 Permission denied（chown 1000）

### D3（08-30 上午）扫描件 OCR（RapidOCR 夹心层 v1）
- ocr_layer.py：无文字层页渲染→RapidOCR→不可见文字按坐标写回（夹心层），**引擎零改动**
- 北京银行 3 页扫描件：58/58 行，核心三列 100%；server.py 接 OCR 分支 + X-OCR-Pages 头
- 部署坑：Dockerfile 漏 COPY ocr_layer.py→崩溃循环；deb.debian.org 卡死→换清华源；
  opencv 需 libgl1；workers 2→1（OCR 模型内存）

### D4（08-30 下午）水印压测暴露夹心层上限
- 工商(244 水印词)/民生(回归基准)扫描件压测：行数丢/列错位/借方合计虚增 714 万
- 根因：OCR 坐标抖动使"跨页固定元素"清洗失效；水印进表头带是致命项

### D5（08-30 晚）GLM-OCR 探测→集成→dual 通道
- glm-ocr 探测：layout_parsing 端点，4704 token/页，markdown 输出无坐标
- 民生全量：水印零污染（根治），但 2/20 页确定性截断（completion 顶 8289）→缺 20 笔
- 集成 dual：GLM 主路径+余额链断点检测+百度表格 V2 定向兜底（只送坏页 ≤4）
- 修复串坑：断点→页映射改内容锚定（累计行数漂移）；页缝去重键禁用 float（20 位流水号超精度）；
  空余额行不参与判重；百度页空余额链式补全+错误余额链式纠错
- **终局：401/401 行、借贷合计分毫不差、逐行余额 401/401、余额链完整**

### D6（08-30 深夜）日志后台 + MCP
- log_store.py（SQLite 30 天保留）+ 三路 API token 计量（glm/vlm/baidu）+ /admin 图形后台
  （统计卡片/14 天柱状图/筛选搜索/CSV 导出，密码保护）
- 容器内 sqlite cursor.description 全 None→列名硬编码；logs 卷 chown；重建丢日志库→挂卷
- mcp-server/（FastMCP 独立 venv，mcp<2）：convert_pdf/logs/stats/health 四工具，
  注册进 ~/.zcode/cli/config.json；修 mcp2.x 改名与 cookiejar 类名两坑

### D7（08-31 ~ 09-01）网页 UI 六轮迭代 + 密码 bug + 异步任务
- v3 队列上限 10/串行/每行密码；v4 行删除+IndexedDB 持久化（修连接堆积 bug）；
  v5 小字收进帮助弹窗；v6 失败重试按钮+网络错误自动重试
- **E2E 用户模拟抓到致命 bug**：password 参数没加 Form()，被解析为 query——
  **网页密码功能从未生效**；修复+三场景验证+CLI 术语文案中性化
- v7 异步任务模式：移动 Firefox 杀后台连接致"第三文件网络错误"（服务端日志证明全部成功）→
  POST /api/tasks 提交即回+轮询+结果 6h；同步接口保留（MCP 用）
- v8 移除自动下载弹窗（移动端体验）
- PaddleOCR-VL 探测（未集成）：单调用整 PDF 64s，内容 396/401，0.009 元/页；
  结论：不替代主路径，整文档重做场景备用
- 企微机器人知识文件 docs/知识文件-企业微信机器人-转换服务.md（LLM 调用手册，凭据隔离）

### D8（09-01）下载链接文件名修复（批量下载分不清谁是谁）
- 症状：转换成功后下载的 xlsx 文件名不区分文件——iOS Safari（预览/主屏 PWA 模式）和
  微信 X5 会忽略 `blob:` URL 上的 `<a download>` 文件名属性，批量文件全部存成同一个通用名
- 修复（server.py 内嵌页 + 任务接口）：
  1. `convertOne` 记录 `it.tid`，成功后记 `it.out_name`（取自服务端）；`persist/restore`
     增存 `tid/out_name`，刷新后仍可下载
  2. `doDownload` 优先走同源 `/api/tasks/{tid}/result`（响应头 `Content-Disposition`
     携带按原 PDF 名生成的文件名，iOS/X5 下载管理器都认）；无 tid（6h 过期回退）才用 blob+download
  3. `/api/tasks/{tid}` 状态接口补充返回 `out_name`
- 验证：本地两文件名 PDF（中文名 UTF-8）E2E——状态接口 out_name 正确、两任务
  Content-Disposition 各带其名、GUI 点击 spy 显示 href 指向 /result、刷新后 IndexedDB
  恢复 tid/out_name 依旧、真实点击服务端日志 200；VPS 部署后外网复测同绿
- 注意：本地 WASM 页（index.html + src/app.js，数据不出设备）下载仍走 blob+download，
  该模式无服务端通道，iOS 命名问题无解，属已知限制

### 本会话最终遗留（均不阻塞）
- 百度体验额度用尽后 dual 降级为 GLM 单通道；PaddleOCR-VL 作为整文档重做备选未接线
- 非 PDF 拖入仅静默跳过（可加提示）；垃圾 PDF 会先烧一次 VLM 调用（可加零成本预检）
- 手头缺真加密样本（本次用 pymupdf 自造验证）；m1_check 样本集未加新格式
- 手机端 Firefox 实测反馈仍欢迎（异步模式理论上已治本）

---
