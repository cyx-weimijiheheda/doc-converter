#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
office2pdf —— 将 Office 文档（Word / PPT / Excel）或 PDF 转换为 PDF 的命令行工具。

支持三种转换：
  1. Office -> 文本型 PDF（默认）：可选中、可搜索、可复制文字，体积小。
     由 LibreOffice 无头模式（soffice --headless）完成排版渲染。
  2. Office -> 图片型 PDF（--image）：把每一页栅格化成图片再打包成 PDF。
     不可选中文字（可防止文字被提取 / 保证跨平台版式 100% 一致），体积较大。
  3. PDF -> 图片型 PDF：直接把一个已有的（文本型）PDF 渲染为图片组成的 PDF，
     同样不可选中文字；无需 LibreOffice，仅需 poppler + img2pdf/Pillow。

依赖：
  - Office 转换：LibreOffice（soffice）。
  - 图片型（含 PDF -> 图片型）：poppler（pdftoppm）+ img2pdf（缺失时回退 Pillow）。

用法示例：
  python office2pdf.py report.docx                      # 转成文本型 report.pdf
  python office2pdf.py a.doc b.pptx c.xlsx -o out/      # 批量转，输出到 out/
  python office2pdf.py slides.pptx --image --dpi 200    # Office 转图片型 PDF（200 DPI）
  python office2pdf.py report.pdf --dpi 300             # PDF 渲染为图片型 report_image.pdf
  python office2pdf.py ./docs --recursive --image       # 递归处理目录
  python office2pdf.py --check                          # 检查本机依赖是否就绪
"""

from __future__ import annotations

import argparse
import glob
import os
import shutil
import subprocess
import sys
import tempfile

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

# 可直接栅格化为图片型 PDF 的输入格式（无需 LibreOffice）
PDF_EXTS = {".pdf"}

# 全部受支持的输入格式
ALL_EXTS = set(SUPPORTED_EXTS) | PDF_EXTS

# 图片型 PDF 支持的栅格化格式
IMG_FORMATS = {"png", "jpg", "jpeg", "tiff", "ppm"}


class ConvertError(RuntimeError):
    """转换失败时抛出，携带可供定位的信息。"""


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
        run(
            [
                soffice,
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
        print(f"  [跳过] 已存在：{final_pdf}")
        return final_pdf

    os.makedirs(outdir, exist_ok=True)
    if verbose:
        print(f"  [栅格化] {src}")

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
def convert_one(
    src: str,
    outdir: str,
    soffice: str,
    *,
    as_image: bool,
    dpi: int,
    img_format: str,
    overwrite: bool,
    timeout: int | None,
    verbose: bool,
) -> str:
    ext = os.path.splitext(src)[1].lower()

    # PDF 输入：直接渲染为图片型 PDF（无需 LibreOffice）
    if ext in PDF_EXTS:
        return pdf_to_image_pdf(
            src, outdir, dpi=dpi, img_format=img_format,
            overwrite=overwrite, verbose=verbose,
        )

    if ext not in SUPPORTED_EXTS:
        raise ConvertError(f"不支持的格式：{src}")

    base = os.path.splitext(os.path.basename(src))[0]
    final_pdf = os.path.join(outdir, base + ".pdf")
    if os.path.exists(final_pdf) and not overwrite:
        print(f"  [跳过] 已存在：{final_pdf}")
        return final_pdf

    os.makedirs(outdir, exist_ok=True)
    if verbose:
        print(f"  [转换] {src}")

    # 1) 文本型 PDF（临时）
    text_pdf = convert_text(soffice, src, outdir, timeout)
    try:
        if not as_image:
            # 直接移动到位
            shutil.move(text_pdf, final_pdf)
            return final_pdf

        # 2) 图片型 PDF：栅格化 -> 打包
        pdftoppm = check_pdftoppm()
        workdir = tempfile.mkdtemp(prefix="office2pdf_img_")
        try:
            images = rasterize(pdftoppm, text_pdf, workdir, dpi, img_format)
            images_to_pdf(images, final_pdf, dpi=dpi)
        finally:
            shutil.rmtree(workdir, ignore_errors=True)
        return final_pdf
    finally:
        # 文本型临时 PDF 位于独立临时目录，一并清理
        d = os.path.dirname(text_pdf)
        shutil.rmtree(d, ignore_errors=True)


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
def run_check(soffice_explicit: str | None) -> int:
    print("=== 环境自检 ===")
    ok = True

    # LibreOffice
    try:
        p = find_soffice(soffice_explicit)
        print(f"[OK]  LibreOffice: {p}")
    except ConvertError as e:
        print(f"[!!]  LibreOffice 未就绪\n      {e}")
        ok = False

    # poppler（图片模式）
    pdftoppm = shutil.which("pdftoppm")
    if pdftoppm:
        print(f"[OK]  poppler (pdftoppm): {pdftoppm}")
    else:
        print("[--]  poppler 未安装（仅图片模式需要）")

    # 图片打包库
    if find_img2pdf():
        import img2pdf
        print(f"[OK]  img2pdf: {getattr(img2pdf, '__version__', '已安装')}")
    else:
        try:
            import PIL  # noqa: F401
            print("[OK]  Pillow（回退引擎）")
        except ImportError:
            print("[--]  img2pdf / Pillow 均未安装（仅图片模式需要，pip install img2pdf）")

    print("=== 自检结束 ===" if ok else "=== 自检结束（存在缺失，见上） ===")
    return 0 if ok else 1


# ---------------------------------------------------------------------------
# 主入口
# ---------------------------------------------------------------------------
def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="office2pdf",
        description="将 Word/PPT/Excel 文档或 PDF 批量转换为 PDF（文本型或图片型）。",
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
    args = parser.parse_args(argv)

    if args.check:
        return run_check(args.soffice)

    if not args.inputs:
        parser.error("请提供至少一个输入文件/目录，或使用 --check 检查环境")

    # 展开输入
    try:
        files = collect_inputs(args.inputs, args.recursive)
    except ConvertError as e:
        print(f"错误：{e}", file=sys.stderr)
        return 2
    if not files:
        print("未找到任何支持的文档（.doc/.docx/.ppt/.pptx/.xls/.xlsx/.pdf 等）。", file=sys.stderr)
        return 2

    # 仅当存在 Office 文档时才需要 LibreOffice；纯 PDF 输入无需 soffice
    has_office = any(os.path.splitext(f)[1].lower() in SUPPORTED_EXTS for f in files)
    has_pdf = any(os.path.splitext(f)[1].lower() in PDF_EXTS for f in files)
    soffice = None
    if has_office:
        try:
            soffice = find_soffice(args.soffice)
        except ConvertError as e:
            print(f"错误：{e}", file=sys.stderr)
            return 2

    parts = []
    if has_office:
        parts.append("Office→" + ("图片型" if args.image else "文本型"))
    if has_pdf:
        parts.append("PDF→图片型")
    mode = " + ".join(parts) or "文本型"
    print(f"共发现 {len(files)} 个文件，模式：{mode}")
    if args.image or has_pdf:
        print(f"  图片参数：dpi={args.dpi}, format={args.img_format}")

    ok_count, fail_count = 0, 0
    failed: list[tuple[str, str]] = []

    for i, src in enumerate(files, 1):
        outdir = args.outdir or os.path.dirname(src)
        print(f"[{i}/{len(files)}] {os.path.basename(src)}")
        try:
            out = convert_one(
                src, outdir, soffice,
                as_image=args.image,
                dpi=args.dpi,
                img_format=args.img_format,
                overwrite=args.overwrite,
                timeout=args.timeout,
                verbose=args.verbose,
            )
            ok_count += 1
            print(f"      -> {out}")
        except ConvertError as e:
            fail_count += 1
            failed.append((src, str(e)))
            print(f"      [失败] {e}", file=sys.stderr)

    print(f"\n完成：成功 {ok_count}，失败 {fail_count}")
    if failed:
        print("失败清单：", file=sys.stderr)
        for src, err in failed:
            print(f"  - {src}\n    {err}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
