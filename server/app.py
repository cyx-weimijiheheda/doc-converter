"""Office2PDF 公开转换服务（FastAPI）。

给 docs/index.html 的「公开服务模式」当后端：访问者打开网页拖文件即可，
不需要 GitHub 账号、不需要 Personal Access Token。

接口：
  GET  /health       健康检查（页面「检查连接」用）
  POST /api/convert  单个文件 → application/pdf
  GET  /             极简说明页

防护：
  - 单文件大小上限 MAX_FILE_MB（默认 40）
  - 进程内并发上限 MAX_CONCURRENCY（默认 2），排队而不是打爆 LibreOffice
  - 文件名白名单化，临时目录独立，响应结束即清理
  - 只接受 office2pdf 支持的扩展名

环境变量：MAX_FILE_MB / MAX_CONCURRENCY / CONVERT_TIMEOUT / ALLOWED_ORIGINS
"""

from __future__ import annotations

import asyncio
import os
import re
import sys
import tempfile
import time
from pathlib import Path
from urllib.parse import quote

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, HTMLResponse
from starlette.background import BackgroundTask

# 允许把本文件放在 server/ 下，同时复用仓库根的 office2pdf.py
ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import office2pdf  # noqa: E402

MAX_FILE_MB = float(os.getenv("MAX_FILE_MB", "40"))
MAX_OUTPUT_MB = float(os.getenv("MAX_OUTPUT_MB", "0"))
MAX_CONCURRENCY = max(1, int(os.getenv("MAX_CONCURRENCY", "2")))
CONVERT_TIMEOUT = int(os.getenv("CONVERT_TIMEOUT", "600"))
ALLOWED = office2pdf.ALL_EXTS

app = FastAPI(title="Office2PDF", version="1.0", docs_url="/docs")

app.add_middleware(
    CORSMiddleware,
    allow_origins=[o for o in os.getenv("ALLOWED_ORIGINS", "*").split(",") if o],
    allow_methods=["GET", "POST", "OPTIONS"],
    allow_headers=["*"],
    max_age=86400,
)

_sem = asyncio.Semaphore(MAX_CONCURRENCY)


def _safe_name(name: str) -> str:
    """只保留纯文件名，去掉路径与控制字符，避免路径穿越。"""
    base = Path(name or "input").name
    base = re.sub(r'[\\/:*?"<>|\x00-\x1f]', "_", base).strip(" .") or "input"
    return base[-120:]


def _convert_sync(src: Path, mode: str, dpi: int) -> Path:
    ext = src.suffix.lower()
    if ext not in ALLOWED:
        raise HTTPException(415, f"不支持的格式：{ext or '(无扩展名)'}")

    outdir = src.parent / "out"
    outdir.mkdir(parents=True, exist_ok=True)

    soffice = None
    if ext in office2pdf.SUPPORTED_EXTS or ext in office2pdf.MARKDOWN_EXTS:
        try:
            soffice = office2pdf.find_soffice()
        except office2pdf.ConvertError as e:
            raise HTTPException(503, f"服务器未就绪：{str(e).splitlines()[0]}") from e

    try:
        out, _skip = office2pdf.convert_one(
            str(src), str(outdir), soffice,
            as_image=(mode == "image"),
            dpi=dpi,
            img_format="png",
            overwrite=True,
            timeout=CONVERT_TIMEOUT,
            verbose=False,
        )
    except office2pdf.ConvertError as e:
        raise HTTPException(500, f"转换失败：{e}") from e
    except OSError as e:
        raise HTTPException(500, f"文件系统错误：{e}") from e

    out_path = Path(out) if out else None
    if not out_path or not out_path.is_file():
        raise HTTPException(500, "转换未产生输出文件")
    return out_path


@app.get("/health")
async def health() -> dict:
    try:
        soffice = office2pdf.find_soffice()
        lo_ok, lo_path = True, soffice
    except office2pdf.ConvertError:
        lo_ok, lo_path = False, None
    return {
        "ok": True,
        "service": "office2pdf",
        "version": "1.0",
        "libreoffice": lo_ok,
        "soffice": lo_path,
        "md2pdf": office2pdf.find_md2pdf(),
        "max_file_mb": MAX_FILE_MB,
        "max_concurrency": MAX_CONCURRENCY,
        "formats": sorted(ALLOWED),
    }


@app.post("/api/convert")
async def convert(
    file: UploadFile = File(...),
    mode: str = Form("text"),
    dpi: int = Form(150),
):
    name = _safe_name(file.filename or "input")
    ext = Path(name).suffix.lower()
    if ext not in ALLOWED:
        raise HTTPException(415, f"不支持的格式：{ext or '(无扩展名)'}")
    if mode not in ("text", "image"):
        mode = "text"
    dpi = 150 if dpi < 60 or dpi > 600 else int(dpi)

    limit = int(MAX_FILE_MB * 1024 * 1024)
    workdir = Path(tempfile.mkdtemp(prefix="o2p_api_"))

    def cleanup() -> None:
        import shutil
        shutil.rmtree(workdir, ignore_errors=True)

    src = workdir / name
    written = 0
    try:
        with src.open("wb") as fh:
            while chunk := await file.read(1 << 20):
                written += len(chunk)
                if written > limit:
                    raise HTTPException(413, f"文件超过 {MAX_FILE_MB:g} MB 上限")
                fh.write(chunk)
    except HTTPException:
        cleanup()
        raise
    except OSError as e:
        cleanup()
        raise HTTPException(500, f"写入临时文件失败：{e}") from e
    finally:
        await file.close()

    if written == 0:
        cleanup()
        raise HTTPException(400, "空文件")

    t0 = time.monotonic()
    try:
        async with _sem:  # 限并发，超出排队
            out_path = await asyncio.to_thread(_convert_sync, src, mode, dpi)
    except HTTPException:
        cleanup()
        raise
    except Exception as e:  # noqa: BLE001 —— 兜底，避免半成品临时目录泄漏
        cleanup()
        raise HTTPException(500, f"内部错误：{type(e).__name__}: {e}") from e

    elapsed = time.monotonic() - t0
    if MAX_OUTPUT_MB > 0 and out_path.stat().st_size > MAX_OUTPUT_MB * 1024 * 1024:
        cleanup()
        raise HTTPException(413, f"转换结果超过 {MAX_OUTPUT_MB:g} MB 上限，请缩小文件或降低 DPI")
    download_name = Path(_safe_name(out_path.name)).name
    headers = {
        "X-Convert-Seconds": f"{elapsed:.2f}",
        # HTTP 头只能是 latin-1，文件名按 URL 编码传给前端
        "X-Output-Name": quote(download_name),
        "Access-Control-Expose-Headers": "X-Convert-Seconds, X-Output-Name, Content-Disposition",
    }
    return FileResponse(
        path=out_path,
        media_type="application/pdf",
        filename=download_name,
        headers=headers,
        background=BackgroundTask(cleanup),
    )


@app.get("/", response_class=HTMLResponse)
async def index() -> str:
    return f"""<!DOCTYPE html>
<html lang="zh-CN"><head><meta charset="utf-8"><title>Office2PDF API</title>
<style>body{{font:15px/1.7 -apple-system,"Noto Sans CJK SC",sans-serif;max-width:680px;margin:60px auto;padding:0 20px;color:#1f2328}}
code{{background:#f2f4f7;padding:1px 6px;border-radius:5px}}pre{{background:#f6f8fa;padding:12px;border-radius:8px;overflow:auto}}</style>
</head><body>
<h1>Office2PDF 转换服务</h1>
<p>把网页里的「服务地址」填成本站地址即可使用，访问者无需任何 token。</p>
<p>健康检查：<a href="/health">/health</a>　接口文档：<a href="/docs">/docs</a></p>
<pre>curl -F "file=@report.docx" -F "mode=text" -F "dpi=150" \\
  {os.getenv("PUBLIC_BASE_URL", "http://127.0.0.1:8000")}/api/convert -o report.pdf</pre>
<p>单文件上限 {MAX_FILE_MB:g} MB，并发 {MAX_CONCURRENCY}。支持：
{", ".join(sorted(ALLOWED))}</p>
</body></html>"""
