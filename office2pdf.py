#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
office2pdf —— 将 Office 文档（Word / PPT / Excel）、Markdown、或 PDF 转换为 PDF 的命令行工具。

支持四种转换：
  1. Office -> 文本型 PDF（默认）：可选中、可搜索、可复制文字，体积小。
     由 LibreOffice 无头模式（soffice --headless）完成排版渲染。
  2. Office -> 图片型 PDF（--image）：把每一页栅格化成图片再打包成 PDF。
     不可选中文字（可防止文字被提取 / 保证跨平台版式 100% 一致），体积较大。
  3. Markdown -> PDF：使用 md2pdf 把 Markdown 渲染为 PDF。
  4. PDF -> 图片型 PDF：直接把一个已有的（文本型）PDF 渲染为图片组成的 PDF，
     同样不可选中文字；无需 LibreOffice，仅需 poppler + img2pdf/Pillow。

依赖：
  - Office 转换：LibreOffice（soffice）。
  - Markdown 转 PDF：md2pdf。
  - 图片型（含 PDF -> 图片型）：poppler（pdftoppm）+ img2pdf（缺失时回退 Pillow）。
  - 页数统计（可选）：poppler 的 pdfinfo。

用法示例：
  python office2pdf.py report.docx                      # 转成文本型 report.pdf
  python office2pdf.py notes.md                        # 转成 notes.pdf
  python office2pdf.py a.doc b.pptx c.xlsx -o out/      # 批量转，输出到 out/
  python office2pdf.py slides.pptx --image --dpi 200    # Office 转图片型 PDF（200 DPI）
  python office2pdf.py report.pdf --dpi 300             # PDF 渲染为图片型 report_image.pdf
  python office2pdf.py ./docs --recursive --image       # 递归处理目录
  python office2pdf.py --check                          # 检查本机依赖是否就绪
  python office2pdf.py ./docs -o out --json             # 输出机器可读的 JSON 结果
"""

from __future__ import annotations

import argparse
import glob
import json
import pathlib
import os
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import unicodedata
from dataclasses import dataclass

# ---------------------------------------------------------------------------
# Office 输入格式 -> 对应的 LibreOffice 过滤模块（仅作展示与提示用）
# ---------------------------------------------------------------------------
SUPPORTED_EXTS = {
    # Word
    ".doc": "writer",
    ".docx": "writer",
    ".odt": "writer",
    ".rtf": "writer",
    # PowerPoint
    ".ppt": "impress",
    ".pptx": "impress",
    ".odp": "impress",
    # Excel
    ".xls": "calc",
    ".xlsx": "calc",
    ".ods": "calc",
    ".csv": "calc",
}

# Markdown 输入格式
MARKDOWN_EXTS = {".md", ".markdown", ".mdown", ".mkd"}

# 可直接栅格化为图片型 PDF 的输入格式（无需 LibreOffice）
PDF_EXTS = {".pdf"}

# 全部受支持的输入格式
ALL_EXTS = set(SUPPORTED_EXTS) | PDF_EXTS | MARKDOWN_EXTS

# 图片型 PDF 支持的栅格化格式
IMG_FORMATS = {"png", "jpg", "jpeg", "tiff", "ppm"}

# 转换结果状态
OK, SKIP, FAIL = "ok", "skip", "fail"


class ConvertError(RuntimeError):
    """转换失败时抛出，携带可供定位的信息。"""


# ===========================================================================
# 输出渲染层：颜色 / 图标 / 宽度对齐 / 进度 / 汇总表
# ===========================================================================
class Style:
    """ANSI 着色器；不可用时所有方法退化为原样返回。"""

    CODES = {
        "bold": "1",
        "dim": "2",
        "red": "31",
        "green": "32",
        "yellow": "33",
        "blue": "34",
        "magenta": "35",
        "cyan": "36",
        "grey": "90",
    }

    def __init__(self, enabled: bool) -> None:
        self.enabled = enabled

    def __call__(self, text: str, *names: str) -> str:
        if not self.enabled or not names:
            return text
        codes = ";".join(self.CODES[n] for n in names if n in self.CODES)
        return f"\033[{codes}m{text}\033[0m" if codes else text

    def bold(self, text: str) -> str:
        return self(text, "bold")

    def dim(self, text: str) -> str:
        return self(text, "dim")

    def ok(self, text: str) -> str:
        return self(text, "green")

    def warn(self, text: str) -> str:
        return self(text, "yellow")

    def err(self, text: str) -> str:
        return self(text, "red")

    def info(self, text: str) -> str:
        return self(text, "cyan")

    def key(self, text: str) -> str:
        return self(text, "bold", "cyan")


# 图标：终端编码不支持 UTF-8 时退化为 ASCII 标记
GLYPHS_UNICODE = {OK: "✓", SKIP: "↷", FAIL: "✗", "run": "⟳"}
GLYPHS_ASCII = {OK: "OK", SKIP: "-=", FAIL: "!!", "run": ">>"}


def _unicode_ok(stream) -> bool:
    enc = (getattr(stream, "encoding", None) or "").lower()
    return "utf" in enc


def display_width(text: str) -> int:
    """按终端显示宽度计算（CJK 全角字符占 2 列）。"""
    width = 0
    for ch in text:
        if unicodedata.combining(ch):
            continue
        width += 2 if unicodedata.east_asian_width(ch) in ("W", "F") else 1
    return width


def fit(text: str, width: int) -> str:
    """把文本裁到指定显示宽度，中段省略，避免长文件名把表格顶歪。"""
    if display_width(text) <= width:
        return text
    if width <= 1:
        return "…" if width else ""
    keep = width - 1
    head = keep // 2
    tail = keep - head
    out, acc = [], 0
    for ch in text:
        w = 2 if unicodedata.east_asian_width(ch) in ("W", "F") else 1
        if acc + w > head:
            break
        out.append(ch)
        acc += w
    left = "".join(out)
    out, acc = [], 0
    for ch in reversed(text):
        w = 2 if unicodedata.east_asian_width(ch) in ("W", "F") else 1
        if acc + w > tail:
            break
        out.append(ch)
        acc += w
    return left + "…" + "".join(reversed(out))


def pad(text: str, width: int, align: str = "left") -> str:
    """按显示宽度补空格对齐。"""
    gap = max(0, width - display_width(text))
    if align == "right":
        return " " * gap + text
    return text + " " * gap


def human_size(num: float) -> str:
    for unit in ("B", "KB", "MB", "GB"):
        if num < 1024 or unit == "GB":
            if unit == "B":
                return f"{num:.0f} B"
            return f"{num:.1f} {unit}"
        num /= 1024
    return f"{num:.1f} GB"


def human_time(seconds: float) -> str:
    if seconds < 1:
        return f"{seconds * 1000:.0f}ms"
    if seconds < 60:
        return f"{seconds:.1f}s"
    minutes, sec = divmod(seconds, 60)
    return f"{int(minutes)}m{sec:04.1f}s"


def pdf_pages(path: str) -> int | None:
    """取 PDF 页数：优先 pdfinfo，失败则退化为解析 /Type /Page 计数。"""
    pdfinfo = shutil.which("pdfinfo")
    if pdfinfo:
        try:
            proc = subprocess.run(
                [pdfinfo, path],
                stdout=subprocess.PIPE, stderr=subprocess.DEVNULL, timeout=20,
            )
            for line in proc.stdout.decode("utf-8", "replace").splitlines():
                if line.startswith("Pages:"):
                    return int(line.split(":", 1)[1].strip())
        except (OSError, ValueError, subprocess.SubprocessError):
            pass
    try:
        with open(path, "rb") as f:
            blob = f.read()
    except OSError:
        return None
    count = len(re.findall(rb"/Type\s*/Page[^s]", blob))
    return count or None


class Spinner:
    """耗时任务进行中的单行动态提示；仅 TTY 且未禁用时启用。"""

    UNICODE_FRAMES = "⠋⠙⠹⠸⠼⠴⠦⠧⠇⠏"
    ASCII_FRAMES = "|/-\\"

    def __init__(self, enabled: bool, label: str, unicode_ok: bool = True) -> None:
        self.enabled = enabled
        self.label = label
        self.frames = self.UNICODE_FRAMES if unicode_ok else self.ASCII_FRAMES
        self._stop = threading.Event()
        self._thread: threading.Thread | None = None
        self._t0 = 0.0

    def start(self) -> None:
        if not self.enabled:
            return
        self._t0 = time.monotonic()
        self._thread = threading.Thread(target=self._loop, daemon=True)
        self._thread.start()

    def _loop(self) -> None:
        i = 0
        while not self._stop.wait(0.12):
            elapsed = time.monotonic() - self._t0
            frame = self.frames[i % len(self.frames)]
            i += 1
            sys.stdout.write(f"\r\033[K  {frame} {self.label} · {elapsed:.1f}s")
            sys.stdout.flush()

    def stop(self) -> float:
        if not self.enabled:
            return 0.0
        self._stop.set()
        if self._thread:
            self._thread.join()
        sys.stdout.write("\r\033[K")
        sys.stdout.flush()
        return time.monotonic() - self._t0


@dataclass
class JobResult:
    """单个文件的转换结果。"""

    index: int
    total: int
    src: str
    status: str
    out: str | None = None
    note: str = ""
    pages: int | None = None
    size: int | None = None
    duration: float = 0.0

    def as_dict(self) -> dict:
        return {
            "index": self.index,
            "total": self.total,
            "input": self.src,
            "output": self.out,
            "status": self.status,
            "note": self.note,
            "pages": self.pages,
            "size": self.size,
            "duration": round(self.duration, 3),
        }


class Reporter:
    """统一负责所有人类可读输出：表头、逐文件行、进度、汇总。"""

    def __init__(
        self,
        *,
        color: bool,
        unicode_ok: bool,
        verbose: bool,
        quiet: bool,
        progress: bool,
        silent: bool = False,
    ) -> None:
        self.st = Style(color)
        self.color = color
        self.glyphs = GLYPHS_UNICODE if unicode_ok else GLYPHS_ASCII
        self.unicode_ok = unicode_ok
        self.verbose = verbose
        self.quiet = quiet
        self.progress = progress
        self.silent = silent
        self.results: list[JobResult] = []
        self.t0 = time.monotonic()
        self._name_w = 24
        self._out_w = 28

    # -- 表头 ---------------------------------------------------------------
    def header(self, total: int, mode: str, detail: str = "") -> None:
        self._name_w = 28
        self._out_w = 32
        if self.quiet or self.silent:
            return
        bar = "─" * 68 if self.unicode_ok else "-" * 68
        title = self.st.bold("Office2PDF")
        subtitle = f"{total} 个文件 · {mode}"
        if detail:
            subtitle += f" · {detail}"
        print(f"{title}  {self.st.dim(subtitle)}")
        print(self.st.dim(bar))
        self._bar = bar

    # -- 单文件行 -----------------------------------------------------------
    def begin(self, index: int, total: int, src: str) -> Spinner:
        label = f"[{index}/{total}] {fit(os.path.basename(src), self._name_w)}"
        spinner = Spinner(
            enabled=self.progress and not self.quiet,
            label=label,
            unicode_ok=self.unicode_ok,
        )
        spinner.start()
        return spinner

    def end(self, result: JobResult) -> None:
        self.results.append(result)
        if self.silent:
            return
        if self.quiet:
            if result.status == FAIL:
                sys.stdout.flush()
                print(f"[失败] {result.src}: {result.note}", file=sys.stderr)
            return

        count = f"[{result.index}/{result.total}]"
        name = fit(os.path.basename(result.src), self._name_w)
        target = ""
        if result.out:
            target = self.st.dim(fit(os.path.basename(result.out), self._out_w))

        if result.status == OK:
            head = self.st.ok(f"{self.glyphs[OK]} {count} {pad(name, self._name_w)}")
            metrics = []
            if result.pages:
                metrics.append(f"{result.pages} 页")
            if result.size is not None:
                metrics.append(human_size(result.size))
            metrics.append(human_time(result.duration))
            tail = self.st.dim("  " + " · ".join(metrics))
            print(f"{head}  →  {target}{tail}")
        elif result.status == SKIP:
            head = self.st.warn(f"{self.glyphs[SKIP]} {count} {pad(name, self._name_w)}")
            print(f"{head}  →  {target}  {self.st.warn('跳过：' + result.note)}")
        else:
            head = self.st.err(f"{self.glyphs[FAIL]} {count} {pad(name, self._name_w)}")
            # 行内只放首行摘要，完整原因留到结尾的失败清单里
            brief = str(result.note).strip().splitlines()[0] if str(result.note).strip() else "未知错误"
            sys.stdout.flush()
            print(f"{head}  {self.st.err('失败：' + brief)}", file=sys.stderr)

        if self.verbose and result.status == OK:
            print(self.st.dim(f"     输入 {result.src}"))

    # -- 汇总 ---------------------------------------------------------------
    def summary(self) -> tuple[int, int, int]:
        ok = sum(1 for r in self.results if r.status == OK)
        skipped = sum(1 for r in self.results if r.status == SKIP)
        failed = sum(1 for r in self.results if r.status == FAIL)
        elapsed = time.monotonic() - self.t0
        total_size = sum(r.size or 0 for r in self.results if r.status == OK)

        if not self.quiet and not self.silent:
            print(self.st.dim(getattr(self, "_bar", "-" * 68)))
            parts = [
                self.st.ok(f"成功 {ok}"),
                self.st.warn(f"跳过 {skipped}") if skipped else self.st.dim("跳过 0"),
                self.st.err(f"失败 {failed}") if failed else self.st.dim("失败 0"),
            ]
            print(f"{self.st.bold('汇总')}  " + self.st.dim(" · ").join(parts))
            extra = [f"总耗时 {human_time(elapsed)}"]
            if total_size:
                extra.append(f"产出 {human_size(total_size)}")
            if ok:
                extra.append(f"平均 {human_time(elapsed / len(self.results))}/个")
            print(self.st.dim("      " + " · ".join(extra)))

            if failed:
                sys.stdout.flush()
                print(self.st.err("\n失败清单："), file=sys.stderr)
                for r in self.results:
                    if r.status == FAIL:
                        print(f"  {self.st.err('✗' if self.unicode_ok else '!!')} {r.src}", file=sys.stderr)
                        for line in str(r.note).splitlines():
                            print(self.st.dim(f"    {line}"), file=sys.stderr)

        return ok, skipped, failed


# ---------------------------------------------------------------------------
# 环境探测
# ---------------------------------------------------------------------------
def find_soffice(explicit: str | None = None) -> str:
    """定位 LibreOffice 的 soffice 可执行文件。

    依次尝试：用户显式指定的路径 -> PATH -> 各平台常见安装位置。
    """
    candidates = []

    if explicit:
        candidates.append(explicit)

    # 从 PATH 中找（Linux/macOS 常为 soffice，Windows 为 soffice.exe）
    for name in ("soffice", "libreoffice", "soffice.exe", "libreoffice.exe"):
        p = shutil.which(name)
        if p:
            candidates.append(p)

    # 常见固定路径
    candidates += [
        "/usr/bin/soffice",
        "/usr/bin/libreoffice",
        "/usr/lib/libreoffice/program/soffice",
        "/usr/lib64/libreoffice/program/soffice",
        # macOS
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        # Windows（常见安装目录）
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]

    for c in candidates:
        if c and os.path.isfile(c) and os.access(c, os.X_OK):
            return c
        # Windows 上 os.access 的 X_OK 不可靠，存在即认为可用
        if c and os.path.isfile(c):
            return c

    raise ConvertError(
        "未找到 LibreOffice (soffice)。请先安装：\n"
        "  - Ubuntu/Debian: sudo apt install libreoffice\n"
        "  - Fedora/RHEL:   sudo dnf install libreoffice\n"
        "  - macOS:         安装 LibreOffice 官方 dmg\n"
        "  - Windows:       安装 LibreOffice，或用 --soffice 指定 soffice.exe 路径\n"
        "仅做无界面转换可安装轻量的 libreoffice-headless / libreoffice-writer+calc+impress。"
    )


def check_pdftoppm() -> str:
    p = shutil.which("pdftoppm")
    if not p:
        raise ConvertError(
            "图片型 PDF 需要 poppler 的 pdftoppm。请安装：\n"
            "  - Ubuntu/Debian: sudo apt install poppler-utils\n"
            "  - Fedora/RHEL:   sudo dnf install poppler-utils\n"
            "  - macOS:         brew install poppler\n"
            "  - Windows:       下载 poppler 并加入 PATH"
        )
    return p


def find_img2pdf() -> bool:
    try:
        import img2pdf  # noqa: F401
        return True
    except ImportError:
        return False


def find_md2pdf() -> bool:
    try:
        import md2pdf  # noqa: F401
        return True
    except ImportError:
        return False


# ---------------------------------------------------------------------------
# 转换核心
# ---------------------------------------------------------------------------
def run(cmd: list[str], timeout: int | None = None) -> subprocess.CompletedProcess:
    """运行命令并返回 CompletedProcess；失败时抛出带 stderr 的异常。"""
    try:
        proc = subprocess.run(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            timeout=timeout,
        )
    except FileNotFoundError as e:
        raise ConvertError(f"命令不存在：{cmd[0]} ({e})") from e
    except subprocess.TimeoutExpired as e:
        raise ConvertError(f"命令超时：{' '.join(cmd)}") from e
    if proc.returncode != 0:
        err = proc.stderr.decode("utf-8", errors="replace").strip()
        out = proc.stdout.decode("utf-8", errors="replace").strip()
        detail = err or out or "无输出"
        hint = ""
        if proc.returncode == 134:
            hint = (
                "\n\n提示：exit=134 是 LibreOffice 内部崩溃（SIGABRT），"
                "常见于处理老式 Microsoft Equation 3.0 / OLE 公式对象。"
                "请将 LibreOffice 升级到 26.8+ 版本（可用 --soffice 指定新版本路径），"
                "该崩溃在新版本中已修复。"
            )
        raise ConvertError(f"命令执行失败 (exit={proc.returncode})：{' '.join(cmd)}\n{detail}{hint}")
    return proc


def convert_text(soffice: str, src: str, outdir: str, timeout: int | None) -> str:
    """用 LibreOffice 无头模式把单个文件转成文本型 PDF，返回生成的 PDF 路径。

    LibreOffice 的输出文件名固定为「源文件名.pdf」，因此每次使用独立临时目录，
    避免同名覆盖与并发冲突，再由调用方把结果移动到目标位置。
    """
    tmp = tempfile.mkdtemp(prefix="office2pdf_")
    try:
        # 每个进程用独立的 LibreOffice 用户配置目录：
        # 并发调用时若共用 ~/.config/libreoffice，第二个实例会因为 profile 被占用而静默失败。
        profile = os.path.join(tmp, "lo_profile")
        os.makedirs(profile, exist_ok=True)
        run(
            [
                soffice,
                f"-env:UserInstallation={pathlib.Path(profile).as_uri()}",
                "--headless",
                "--norestore",
                "--invisible",
                "--convert-to", "pdf",
                "--outdir", tmp,
                src,
            ],
            timeout=timeout,
        )
        base = os.path.splitext(os.path.basename(src))[0]
        out_pdf = os.path.join(tmp, base + ".pdf")
        if not os.path.isfile(out_pdf):
            # 某些版本对含空格/特殊字符的文件名输出可能不一致，兜底查找
            pdfs = glob.glob(os.path.join(tmp, "*.pdf"))
            if len(pdfs) == 1:
                out_pdf = pdfs[0]
            else:
                raise ConvertError(f"LibreOffice 未生成预期 PDF：{src}")
        return out_pdf
    except Exception:
        shutil.rmtree(tmp, ignore_errors=True)
        raise


def convert_markdown(src: str, outdir: str, *, overwrite: bool, verbose: bool) -> str:
    """把 Markdown 文档转换为 PDF，依赖 md2pdf。"""
    base = os.path.splitext(os.path.basename(src))[0]
    final_pdf = os.path.join(outdir, base + ".pdf")
    if os.path.exists(final_pdf) and not overwrite:
        return final_pdf

    os.makedirs(outdir, exist_ok=True)
    if verbose:
        print(f"  [转换] {src}", file=sys.stderr)

    try:
        import md2pdf
    except ImportError as e:
        raise ConvertError(
            "Markdown 转 PDF 需要 md2pdf。请安装：\n"
            "  pip install md2pdf"
        ) from e

    callable_target = getattr(md2pdf, "md2pdf", None)
    if not callable(callable_target):
        try:
            mod = __import__("md2pdf.core", fromlist=["md2pdf"])
            callable_target = getattr(mod, "md2pdf", None)
        except ImportError:
            callable_target = None

    if not callable(callable_target):
        raise ConvertError("已安装 md2pdf，但未找到兼容的调用入口。")

    try:
        callable_target(pdf=pathlib.Path(final_pdf), md=pathlib.Path(src))
        return final_pdf
    except TypeError as e:
        raise ConvertError("md2pdf 调用失败：API 与当前版本不兼容。") from e


def rasterize(pdftoppm: str, pdf: str, workdir: str, dpi: int, fmt: str) -> list[str]:
    """用 pdftoppm 把 PDF 每一页栅格化成图片，返回图片路径列表（按页码排序）。"""
    prefix = os.path.join(workdir, "page")
    ext = "jpg" if fmt.lower() in ("jpg", "jpeg") else fmt.lower()
    run([pdftoppm, f"-{ext}", "-r", str(dpi), pdf, prefix])
    imgs = sorted(glob.glob(f"{prefix}-*.{ext}"))
    if not imgs:
        raise ConvertError(f"栅格化未产生任何图片：{pdf}")
    return imgs


def images_to_pdf(images: list[str], out_pdf: str, dpi: int = 150) -> None:
    """把图片列表合并为单个 PDF，优先 img2pdf（无损、体积小），回退 Pillow。"""
    if find_img2pdf():
        import img2pdf
        with open(out_pdf, "wb") as f:
            f.write(img2pdf.convert(images))
        return

    try:
        from PIL import Image
    except ImportError as e:
        raise ConvertError(
            "图片型 PDF 需要 img2pdf 或 Pillow。请安装其一：\n"
            "  pip install img2pdf"
        ) from e

    first = Image.open(images[0]).convert("RGB")
    rest = [Image.open(i).convert("RGB") for i in images[1:]]
    first.save(out_pdf, "PDF", save_all=True, append_images=rest, resolution=float(dpi))
    for im in [first, *rest]:
        im.close()


# ---------------------------------------------------------------------------
# PDF -> 图片型 PDF
# ---------------------------------------------------------------------------
def pdf_to_image_pdf(
    src: str,
    outdir: str,
    *,
    dpi: int,
    img_format: str,
    overwrite: bool,
    verbose: bool,
) -> str:
    """把文本型 PDF 渲染为图片组成的 PDF（不可选中文字，无需 LibreOffice）。"""
    base = os.path.splitext(os.path.basename(src))[0]
    # 输入与输出同为 .pdf，加后缀避免覆盖源文件
    final_pdf = os.path.join(outdir, base + "_image.pdf")
    if os.path.exists(final_pdf) and not overwrite:
        return final_pdf

    os.makedirs(outdir, exist_ok=True)
    if verbose:
        print(f"  [栅格化] {src}", file=sys.stderr)

    pdftoppm = check_pdftoppm()
    workdir = tempfile.mkdtemp(prefix="office2pdf_img_")
    try:
        images = rasterize(pdftoppm, src, workdir, dpi, img_format)
        images_to_pdf(images, final_pdf, dpi=dpi)
    finally:
        shutil.rmtree(workdir, ignore_errors=True)
    return final_pdf


# ---------------------------------------------------------------------------
# 单文件转换入口
# ---------------------------------------------------------------------------
def planned_output(src: str, outdir: str, as_image: bool) -> str:
    """提前算出输出文件名，用于「已存在则跳过」判断与展示。"""
    base = os.path.splitext(os.path.basename(src))[0]
    ext = os.path.splitext(src)[1].lower()
    if ext in PDF_EXTS:
        return os.path.join(outdir, base + "_image.pdf")
    return os.path.join(outdir, base + ".pdf")


def convert_one(
    src: str,
    outdir: str,
    soffice: str | None,
    *,
    as_image: bool,
    dpi: int,
    img_format: str,
    overwrite: bool,
    timeout: int | None,
    verbose: bool,
) -> tuple[str | None, str]:
    """转换单个文件，返回 (输出路径, 跳过原因)；未跳过时原因为空串。"""
    ext = os.path.splitext(src)[1].lower()
    final_pdf = planned_output(src, outdir, as_image)

    if ext not in ALL_EXTS:
        raise ConvertError(f"不支持的格式：{src}")

    if ext in SUPPORTED_EXTS and soffice is None:
        raise ConvertError("缺少 LibreOffice（soffice），无法转换 Office 文档。")

    if os.path.exists(final_pdf) and not overwrite:
        return final_pdf, "输出已存在，用 --overwrite 覆盖"

    # PDF 输入：直接渲染为图片型 PDF（无需 LibreOffice）
    if ext in PDF_EXTS:
        return pdf_to_image_pdf(
            src, outdir, dpi=dpi, img_format=img_format,
            overwrite=overwrite, verbose=verbose,
        ), ""

    if ext in MARKDOWN_EXTS:
        if not as_image:
            return convert_markdown(src, outdir, overwrite=overwrite, verbose=verbose), ""
        with tempfile.TemporaryDirectory(prefix="office2pdf_md_") as tmp:
            text_pdf = convert_markdown(src, tmp, overwrite=True, verbose=verbose)
            return pdf_to_image_pdf(text_pdf, outdir, dpi=dpi, img_format=img_format,
                                    overwrite=overwrite, verbose=verbose), ""

    os.makedirs(outdir, exist_ok=True)
    if verbose:
        print(f"  [转换] {src}", file=sys.stderr)

    # 1) 文本型 PDF（临时目录内）
    text_pdf = convert_text(soffice, src, outdir, timeout)
    try:
        if not as_image:
            shutil.move(text_pdf, final_pdf)
            return final_pdf, ""

        # 2) 图片型 PDF：栅格化 -> 打包
        pdftoppm = check_pdftoppm()
        workdir = tempfile.mkdtemp(prefix="office2pdf_img_")
        try:
            images = rasterize(pdftoppm, text_pdf, workdir, dpi, img_format)
            images_to_pdf(images, final_pdf, dpi=dpi)
        finally:
            shutil.rmtree(workdir, ignore_errors=True)
        return final_pdf, ""
    finally:
        # 文本型临时 PDF 位于独立临时目录，一并清理
        shutil.rmtree(os.path.dirname(text_pdf), ignore_errors=True)


# ---------------------------------------------------------------------------
# 输入展开：文件 / 目录 / 通配符
# ---------------------------------------------------------------------------
def collect_inputs(inputs: list[str], recursive: bool) -> list[str]:
    files: list[str] = []
    for item in inputs:
        if os.path.isdir(item):
            if recursive:
                for root, _dirs, names in os.walk(item):
                    for n in names:
                        if os.path.splitext(n)[1].lower() in ALL_EXTS:
                            files.append(os.path.join(root, n))
            else:
                for n in os.listdir(item):
                    p = os.path.join(item, n)
                    if os.path.isfile(p) and os.path.splitext(n)[1].lower() in ALL_EXTS:
                        files.append(p)
        elif any(c in item for c in "*?["):
            for p in glob.glob(item):
                if os.path.isfile(p) and os.path.splitext(p)[1].lower() in ALL_EXTS:
                    files.append(p)
        else:
            if not os.path.isfile(item):
                raise ConvertError(f"文件不存在：{item}")
            files.append(item)
    # 去重并保持顺序
    seen, result = set(), []
    for f in files:
        ap = os.path.abspath(f)
        if ap not in seen:
            seen.add(ap)
            result.append(ap)
    return result


# ---------------------------------------------------------------------------
# 环境自检
# ---------------------------------------------------------------------------
def run_check(soffice_explicit: str | None, color: bool = True) -> int:
    st = Style(color)
    unicode_ok = _unicode_ok(sys.stdout)
    mark = {"ok": st.ok("✓"), "no": st.warn("○"), "bad": st.err("✗")}
    if not unicode_ok:
        mark = {"ok": st.ok("[OK]"), "no": st.warn("[--]"), "bad": st.err("[!!]")}

    print(st.bold("Office2PDF 环境自检"))
    print(st.dim("─" * 68 if unicode_ok else "-" * 68))
    ok = True
    rows: list[tuple[str, str, str]] = []

    try:
        p = find_soffice(soffice_explicit)
        rows.append((mark["ok"], "LibreOffice", p))
    except ConvertError as e:
        ok = False
        rows.append((mark["bad"], "LibreOffice", "未就绪 —— " + str(e).splitlines()[0]))

    if find_md2pdf():
        rows.append((mark["ok"], "md2pdf", "已安装（Markdown → PDF）"))
    else:
        rows.append((mark["no"], "md2pdf", "未安装（pip install md2pdf）"))

    pdftoppm = shutil.which("pdftoppm")
    rows.append(
        (mark["ok"], "poppler", pdftoppm) if pdftoppm
        else (mark["no"], "poppler", "未安装（仅图片模式需要）")
    )

    if find_img2pdf():
        import img2pdf
        rows.append((mark["ok"], "img2pdf", getattr(img2pdf, "__version__", "已安装")))
    else:
        try:
            import PIL
            rows.append((mark["ok"], "Pillow", f"回退引擎 {getattr(PIL, '__version__', '')}".strip()))
        except ImportError:
            rows.append((mark["no"], "img2pdf/Pillow", "未安装（pip install img2pdf）"))

    pdfinfo = shutil.which("pdfinfo")
    rows.append(
        (mark["ok"], "pdfinfo", pdfinfo) if pdfinfo
        else (mark["no"], "pdfinfo", "未安装（页数显示退化为估算）")
    )

    name_w = max(display_width(r[1]) for r in rows) + 2
    for icon, name, detail in rows:
        print(f"  {icon} {pad(st.bold(name), name_w)} {detail}")

    print(st.dim("─" * 68 if unicode_ok else "-" * 68))
    if ok:
        print(st.ok("核心依赖就绪，可以直接转换。"))
    else:
        print(st.err("核心依赖缺失：Office 转换暂不可用。"))
    return 0 if ok else 1


# ---------------------------------------------------------------------------
# 主入口
# ---------------------------------------------------------------------------
def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="office2pdf",
        description="将 Word/PPT/Excel 文档、Markdown、或 PDF 批量转换为 PDF（文本型或图片型）。",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__.split("用法示例：")[1] if "用法示例：" in __doc__ else None,
    )
    parser.add_argument("inputs", nargs="*", help="待转换的文件、目录或通配符（可多个）")
    parser.add_argument("-o", "--outdir", default=None, help="输出目录（默认与源文件同目录）")
    parser.add_argument("-r", "--recursive", action="store_true", help="递归处理目录下的所有支持文件")
    parser.add_argument("--image", action="store_true", help="Office 输出图片型 PDF（PDF 输入会自动图片化，无需此标志）")
    parser.add_argument("--dpi", type=int, default=150, help="图片型 PDF 的分辨率，默认 150")
    parser.add_argument("--img-format", default="png", choices=sorted(IMG_FORMATS),
                        help="图片型 PDF 的栅格化格式，默认 png")
    parser.add_argument("--soffice", default=None, help="显式指定 soffice 可执行文件路径")
    parser.add_argument("--overwrite", action="store_true", help="覆盖已存在的同名 PDF")
    parser.add_argument("--timeout", type=int, default=600, help="单文件转换超时秒数，默认 600")
    parser.add_argument("-c", "--check", action="store_true", help="检查环境依赖后退出")
    parser.add_argument("-v", "--verbose", action="store_true", help="显示详细过程")
    parser.add_argument("-q", "--quiet", action="store_true", help="只输出错误与最终统计")
    parser.add_argument("--json", action="store_true", help="以 JSON 输出结果（便于脚本/CI 消费）")
    parser.add_argument("--no-color", action="store_true", help="禁用 ANSI 颜色")
    parser.add_argument("--no-progress", action="store_true", help="禁用动态进度（适合日志/CI）")

    args = parser.parse_args(argv)

    # 颜色：显式关闭 > NO_COLOR > FORCE_COLOR > TTY 探测
    color = sys.stdout.isatty() and os.environ.get("TERM", "") != "dumb"
    if os.environ.get("NO_COLOR"):
        color = False
    if os.environ.get("FORCE_COLOR"):
        color = True
    if args.no_color or args.json:
        color = False
    unicode_ok = _unicode_ok(sys.stdout)
    progress = sys.stdout.isatty() and not args.no_progress and not args.json and not args.quiet

    if args.check:
        return run_check(args.soffice, color=color)

    if not args.inputs:
        parser.error("请提供至少一个输入文件/目录，或使用 --check 检查环境")

    st = Style(color)

    # 展开输入
    try:
        files = collect_inputs(args.inputs, args.recursive)
    except ConvertError as e:
        print(st.err(f"错误：{e}"), file=sys.stderr)
        return 2
    if not files:
        print(st.err("未找到任何支持的文档（.doc/.docx/.ppt/.pptx/.xls/.xlsx/.md/.pdf 等）。"), file=sys.stderr)
        return 2

    # 仅当存在 Office 文档时才需要 LibreOffice；纯 PDF / Markdown 输入无需 soffice
    has_office = any(os.path.splitext(f)[1].lower() in SUPPORTED_EXTS for f in files)
    has_pdf = any(os.path.splitext(f)[1].lower() in PDF_EXTS for f in files)
    has_markdown = any(os.path.splitext(f)[1].lower() in MARKDOWN_EXTS for f in files)
    soffice = None
    if has_office:
        try:
            soffice = find_soffice(args.soffice)
        except ConvertError as e:
            print(st.err(f"错误：{e}"), file=sys.stderr)
            return 2

    parts = []
    if has_office:
        parts.append("Office→" + ("图片型" if args.image else "文本型"))
    if has_markdown:
        parts.append("Markdown→PDF")
    if has_pdf:
        parts.append("PDF→图片型")
    mode = " + ".join(parts) or "文本型"
    detail = f"dpi={args.dpi} · {args.img_format}" if (args.image or has_pdf) else ""

    reporter = Reporter(
        color=color,
        unicode_ok=unicode_ok,
        verbose=args.verbose,
        quiet=args.quiet,
        progress=progress,
        silent=args.json,
    )
    # 先按文件名长度收窄列宽，让表格更紧凑
    names = [os.path.basename(f) for f in files]
    if names:
        reporter._name_w = min(34, max(18, max(display_width(n) for n in names) + 1))
    reporter.header(len(files), mode, detail)

    for i, src in enumerate(files, 1):
        outdir = args.outdir or os.path.dirname(src)
        spinner = reporter.begin(i, len(files), src)
        t0 = time.monotonic()
        result = JobResult(index=i, total=len(files), src=src, status=OK)
        try:
            out, skip_reason = convert_one(
                src, outdir, soffice,
                as_image=args.image,
                dpi=args.dpi,
                img_format=args.img_format,
                overwrite=args.overwrite,
                timeout=args.timeout,
                verbose=args.verbose,
            )
            result.out = out
            if skip_reason:
                result.status = SKIP
                result.note = skip_reason
            elif out and os.path.isfile(out):
                result.size = os.path.getsize(out)
                result.pages = pdf_pages(out)
        except ConvertError as e:
            result.status = FAIL
            result.note = str(e)
        except OSError as e:
            result.status = FAIL
            result.note = f"文件系统错误：{e}"
        except Exception as e:  # 未预期的异常也要计入汇总，方便批量排查
            result.status = FAIL
            result.note = f"{type(e).__name__}: {e}"
        finally:
            spinner.stop()
        result.duration = time.monotonic() - t0
        reporter.end(result)

    ok_count, _skipped, fail_count = reporter.summary()

    if args.json:
        payload = {
            "mode": mode,
            "total": len(files),
            "ok": ok_count,
            "skipped": _skipped,
            "failed": fail_count,
            "results": [r.as_dict() for r in reporter.results],
        }
        print(json.dumps(payload, ensure_ascii=False, indent=2))

    return 0 if fail_count == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
