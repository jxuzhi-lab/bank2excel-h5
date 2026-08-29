# bank2excel-h5 维护手册（2026-08-29 版）

> 本手册供新会话接手维护时阅读。读完本文即可了解：项目全貌、双通道部署、更新流程、踩坑记录、待办。
> 项目状态：**M0-M7 全部完成，双通道线上运行中**。

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
> VPS 上代码更新：因 GitHub 主站被墙，用 codeload 下 zip 覆盖 `/opt/bank2excel-h5`（见 deploy_vps.py 逻辑），或用 SFTP 传文件到家目录后 sudo cp。

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
| VPS 服务 | ✅ 在线（124.223.110.222:80，Docker 自动重启） |
| 本机 8766 服务 | ✅ 运行中（start_server.bat 可重启） |
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

## 八、快速参考

```bash
# 本地回归(用项目 venv, 系统 python 3.7 跑不了)
cd bank2excel-h5/python && "C:/Users/Administrator/Documents/银行对账单转化pdf/.venv/Scripts/python.exe" m1_check.py

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
