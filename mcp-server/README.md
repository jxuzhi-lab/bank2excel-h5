# bank2excel MCP Server

让大模型（ZCode / Claude / Cursor 等支持 MCP 的客户端）直接调用对账单转换服务。

## 工具

| 工具 | 功能 | 备注 |
|------|------|------|
| `convert_pdf` | 本地 PDF → Excel（自动存到源文件旁） | 支持加密 PDF（password 参数）、扫描件 OCR、新格式视觉学习 |
| `conversion_logs` | 查询最近转换日志（时间/文件/状态/耗时/通道/API 消耗） | 需 `B2X_ADMIN_PASSWORD` |
| `conversion_stats` | 汇总统计（总量/成功率/各 API token/百度点数） | 需 `B2X_ADMIN_PASSWORD` |
| `service_health` | 服务健康状态 | 无需配置 |

## 环境变量

| 变量 | 默认 | 说明 |
|------|------|------|
| `B2X_API_URL` | `http://124.223.110.222` | 转换服务地址（本机调试用 `http://127.0.0.1:8766`） |
| `B2X_ADMIN_PASSWORD` | 空 | 日志后台密码（VPS compose 里的 `ADMIN_PASSWORD`） |
| `B2X_TIMEOUT` | `900` | 转换超时秒数（扫描件 OCR 较慢，勿设太小） |

## 安装与注册（ZCode）

```bash
# 1. 独立环境(不污染服务端 venv; 需 mcp 1.x, 2.x API 不兼容)
python -m venv venv
venv/Scripts/pip install "mcp<2"

# 2. 注册到用户级配置 ~/.zcode/cli/config.json
{
  "mcp": {
    "servers": {
      "bank2excel": {
        "command": "<绝对路径>/mcp-server/venv/Scripts/python.exe",
        "args": ["<绝对路径>/mcp-server/server.py"],
        "env": {
          "B2X_API_URL": "http://124.223.110.222",
          "B2X_ADMIN_PASSWORD": "<VPS compose 里的 ADMIN_PASSWORD>"
        }
      }
    }
  }
}
```

重启客户端即自动连接（用户级配置所有工作区可用）。其他 MCP 客户端（Claude Desktop 等）用标准 `mcpServers` 格式注册同一命令即可。

## 大模型提示词示例

- “把 D:\账单\华夏银行流水.pdf 转成 Excel，密码在文件名里”
- “查一下最近的转换记录，有没有失败的”
- “转换服务健康状态如何？这个月消耗了多少 API 额度”

## 注意

- `convert_pdf` 对扫描件（OCR）可能耗时 1-5 分钟，MCP 客户端超时建议 ≥10 分钟
- 服务端日志只保留 30 天；日志/统计工具需要管理密码
- venv 目录不入库（.gitignore）
