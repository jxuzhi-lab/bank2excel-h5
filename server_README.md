# 私有转换服务（server）

手机端稳定方案：把转换放到自己电脑上跑（本地 CPython + PyMuPDF 原生引擎），手机只负责上传/下载，**不再下载 17MB 引擎**，100% 稳定、秒开。

## 启动

```bash
# 双击 start_server.bat 或命令行:
python server.py --host 0.0.0.0 --port 8766
```

启动后本机服务在 `http://127.0.0.1:8766`。

## 手机访问

### 同一 WiFi（最简单）
1. 查电脑局域网 IP：`ipconfig` → 找 `IPv4 地址`（如 `192.168.1.5`）
2. 手机浏览器打开 `http://192.168.1.5:8766`
3. 首次 Windows 防火墙会弹窗 → 点"允许访问"

### 不在同一网络（外网/流量）
用内网穿透，任选其一：
- **cpolar**（推荐，免费隧道）：`cpolar authtoken <你的token>` → `cpolar http 8766` → 得到 `https://xxx.r3.cpolar.cn`
- **ngrok**：`ngrok http 8766`
- **Tailscale**（更稳，组网后手机访问电脑 Tailscale IP:8766）

## 接口

- `GET /` 手机友好上传页
- `POST /api/convert` multipart: `file=<pdf>`（可选 `password=<加密密码>`）→ 返回 xlsx 文件

## 说明

- 复用 `python/` 下同一套引擎（`shim.convert_bytes`），与 GitHub Pages 版转换结果完全一致
- 依赖：`pip install fastapi uvicorn python-multipart`（已装于工作区 venv）
- 单文件上限 200MB；转换在内存完成，文件不留盘
