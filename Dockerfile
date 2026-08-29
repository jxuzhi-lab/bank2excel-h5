# bank2excel-h5 私有转换服务 — VPS 部署
# 基础镜像: 官方 Python 3.12 slim
# 构建: docker build -t bank2excel-server .
# 运行: docker run -d --name bank2excel -p 80:8000 --restart unless-stopped bank2excel-server
FROM python:3.12-slim

WORKDIR /app

# 系统依赖(字体/时区, pymupdf 纯 wheel 无需编译)
RUN apt-get update && apt-get install -y --no-install-recommends \
    fonts-noto-cjk \
    tzdata \
    && rm -rf /var/lib/apt/lists/* \
    && ln -sf /usr/share/zoneinfo/Asia/Shanghai /etc/localtime

# Python 依赖(固定版本保证可复现)
COPY requirements-server.txt .
RUN pip install --no-cache-dir -r requirements-server.txt

# 引擎(server.py 运行时 import shim → extract_bank_statement)
COPY server.py .
COPY python/ python/

# 非 root 运行(更安全)
RUN useradd -m app && chown -R app:app /app
USER app

EXPOSE 8000
CMD ["uvicorn", "server:app", "--host", "0.0.0.0", "--port", "8000", "--workers", "2"]
