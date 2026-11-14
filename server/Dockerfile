# Office2PDF 公开转换服务
#
# 自包含镜像：app.py + requirements.txt 来自构建上下文，office2pdf.py 由
# OFFICE2PDF_RAW 拉取（默认取 doc-converter 的 main 分支），因此这个目录
# 可以单独上传成 Hugging Face Space，无需携带整个仓库。
#
# 本地构建：
#   docker build -t office2pdf-api server/
#   docker run -p 7860:7860 office2pdf-api
FROM python:3.11-slim

ENV PYTHONUNBUFFERED=1 \
    HOME=/tmp \
    PORT=7860 \
    MAX_FILE_MB=40 \
    MAX_CONCURRENCY=2 \
    CONVERT_TIMEOUT=600

ARG OFFICE2PDF_RAW=https://raw.githubusercontent.com/cyx-weimijiheheda/doc-converter/main/office2pdf.py

# LibreOffice（转换引擎）+ poppler（图片型 PDF）+ 中文字体（避免方框）
RUN apt-get update && apt-get install -y --no-install-recommends \
        libreoffice-writer libreoffice-calc libreoffice-impress \
        poppler-utils \
        fonts-noto-cjk fonts-noto-cjk-extra fonts-wqy-zenhei \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

COPY app.py ./
RUN python -c "import urllib.request; urllib.request.urlretrieve('${OFFICE2PDF_RAW}', 'office2pdf.py')" \
    && python -c "import ast; ast.parse(open('office2pdf.py').read())" \
    && echo "office2pdf.py fetched"

# 预热一次 LibreOffice，让容器首请求不必等字体缓存
RUN soffice --headless --version || libreoffice --headless --version

EXPOSE 7860
HEALTHCHECK --interval=30s --timeout=5s --start-period=20s \
    CMD python -c "import urllib.request,sys; sys.exit(0 if urllib.request.urlopen('http://127.0.0.1:7860/health', timeout=4).status==200 else 1)"

CMD ["sh", "-c", "uvicorn app:app --host 0.0.0.0 --port ${PORT:-7860} --timeout-keep-alive 75"]
