# Office2PDF 公开转换服务
#
# 构建上下文包含 server/ 目录和根目录的 office2pdf.py。
#
# 本地构建：
#   docker build -f server/Dockerfile -t office2pdf-api .
#   docker run -p 7860:7860 office2pdf-api
FROM python:3.11-slim

ENV PYTHONUNBUFFERED=1 \
    HOME=/tmp \
    PORT=7860 \
    MAX_FILE_MB=40 \
    MAX_CONCURRENCY=2 \
    CONVERT_TIMEOUT=600

# LibreOffice（转换引擎）+ poppler（图片型 PDF）+ 中文字体（避免方框）
RUN apt-get update && apt-get install -y --no-install-recommends \
        libreoffice-writer libreoffice-calc libreoffice-impress \
        poppler-utils \
        fonts-noto-cjk fonts-noto-cjk-extra fonts-wqy-zenhei \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY server/requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

COPY server/app.py office2pdf.py ./

# 预热一次 LibreOffice，让容器首请求不必等字体缓存
RUN soffice --headless --version || libreoffice --headless --version

EXPOSE 7860
HEALTHCHECK --interval=30s --timeout=5s --start-period=20s \
    CMD python -c "import urllib.request,sys; sys.exit(0 if urllib.request.urlopen('http://127.0.0.1:7860/health', timeout=4).status==200 else 1)"

CMD ["sh", "-c", "uvicorn app:app --host 0.0.0.0 --port ${PORT:-7860} --timeout-keep-alive 75"]
