# bank2excel-h5 私有转换服务 — VPS 部署
# 基础镜像: 官方 Python 3.12 slim
# 构建: docker build -t bank2excel-server .
# 运行: docker run -d --name bank2excel -p 80:8000 --restart unless-stopped bank2excel-server
FROM python:3.12-slim

WORKDIR /app

# 系统依赖: tzdata + OpenCV 运行库(扫描件 OCR 用; 不需要中文字体渲染)
# apt 走清华镜像(deb.debian.org 国内不稳定, 曾卡死构建)
RUN sed -i 's|deb.debian.org|mirrors.tuna.tsinghua.edu.cn|g' /etc/apt/sources.list.d/debian.sources \
    && apt-get update && apt-get install -y --no-install-recommends \
    tzdata libgl1 libglib2.0-0 \
    && rm -rf /var/lib/apt/lists/* \
    && ln -sf /usr/share/zoneinfo/Asia/Shanghai /etc/localtime

# Python 依赖(固定版本; 用清华 PyPI 镜像加速国内构建)
COPY requirements-server.txt .
RUN pip install --no-cache-dir -i https://pypi.tuna.tsinghua.edu.cn/simple -r requirements-server.txt

# 引擎(server.py 运行时 import shim → extract_bank_statement)
COPY server.py .
COPY ocr_layer.py .
COPY python/ python/

# 非 root 运行(更安全)
RUN useradd -m app && chown -R app:app /app
USER app

EXPOSE 8000
CMD ["uvicorn", "server:app", "--host", "0.0.0.0", "--port", "8000", "--workers", "1"]
