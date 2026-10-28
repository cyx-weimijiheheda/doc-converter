#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
doc2htmlpdf —— 将 Markdown / Office 文档批量转换为 HTML 和 / 或 PDF。

支持输入：
  Markdown : .md .markdown .mdown .mkd
  Office   : .doc .docx .docm .odt .rtf .ppt .pptx .odp .xls .xlsx .ods .csv

输出目标（--to）：
  html | pdf | both（默认 both）

转换引擎：
  Markdown -> HTML : Python `markdown` 库（extra + fenced_code + tables + codehilite）
  HTML     -> PDF  : 自动选择可用引擎，优先级
                     WeasyPrint（质量最好） > Chromium 无头打印 > LibreOffice Writer/Web
  Office   -> HTML/PDF : LibreOffice（soffice --headless --convert-to）

示例：
  python doc2htmlpdf.py README.md                     # 生成 README.html + README.pdf
  python doc2htmlpdf.py README.md --to pdf            # 只生成 PDF
  python doc2htmlpdf.py ./docs -r --to both -o out/   # 递归转换目录
  python doc2htmlpdf.py a.docx b.pptx --to html       # Office -> HTML
  python doc2htmlpdf.py --check                       # 检查依赖
"""
from __future__ import annotations

import argparse
import glob as globmod
import os
import shutil
import subprocess
import sys
import tempfile

VERSION = "1.0.0"

MD_EXTS = {".md", ".markdown", ".mdown", ".mkd"}
OFFICE_EXTS = {
    ".doc", ".docx", ".docm", ".odt", ".rtf",
    ".ppt", ".pptx", ".odp",
    ".xls", ".xlsx", ".ods", ".csv",
}
ALL_EXTS = MD_EXTS | OFFICE_EXTS


class ConvertError(Exception):
    pass


# ---------------------------------------------------------------------------
# 依赖探测
# ---------------------------------------------------------------------------
def find_soffice(explicit: str | None = None) -> str | None:
    if explicit:
        if os.path.isfile(explicit) and os.access(explicit, os.X_OK):
            return explicit
        found = shutil.which(explicit)
        if found:
            return found
        raise ConvertError(f"指定的 soffice 不可用：{explicit}")

    found = shutil.which("soffice") or shutil.which("libreoffice")
    if found:
        return found

    candidates = [
        "/usr/bin/soffice",
        "/usr/local/bin/soffice",
        "/usr/lib/libreoffice/program/soffice",
        "/opt/libreoffice/program/soffice",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    ]
    for c in candidates:
        if os.path.isfile(c) and os.access(c, os.X_OK):
            return c
    for c in globmod.glob("/opt/libreoffice*/program/soffice"):
        if os.access(c, os.X_OK):
            return c
    return None


def find_chromium() -> str | None:
    for name in ("chromium", "chromium-browser", "google-chrome",
                 "google-chrome-stable", "chrome", "msedge"):
        p = shutil.which(name)
        if p:
            return p
    return None


def has_weasyprint() -> bool:
    try:
        import weasyprint  # noqa: F401
        return True
    except Exception:
        return False


def has_markdown() -> bool:
    try:
        import markdown  # noqa: F401
        return True
    except Exception:
        return False


# ---------------------------------------------------------------------------
# Markdown -> HTML
# ---------------------------------------------------------------------------
def pygments_css(style: str = "default") -> str:
    try:
        from pygments.formatters import HtmlFormatter
        return HtmlFormatter(style=style).get_style_defs(".codehilite")
    except Exception:
        return ""


def md_to_fragment(text: str, with_toc: bool = False) -> tuple[str, str]:
    """把 Markdown 源码转成 HTML 片段，返回 (body_html, toc_html)。"""
    try:
        import markdown
    except ImportError:
        raise ConvertError(
            "缺少 Python 库 markdown，请先安装：pip install markdown"
        )

    extensions = [
        "extra", "sane_lists", "fenced_code", "tables",
        "attr_list", "def_list", "footnotes", "codehilite",
    ]
    configs = {
        "codehilite": {"guess_lang": False, "noclasses": False, "linenums": False},
    }
    if with_toc:
        extensions.append("toc")
        configs["toc"] = {"permalink": False, "toc_depth": "2-4"}

    md = markdown.Markdown(extensions=extensions, extension_configs=configs)
    body = md.convert(text)
    toc = getattr(md, "toc", "") if with_toc else ""
    return body, toc


BASE_CSS = r"""
@page { size: A4; margin: 18mm 16mm 20mm; }
* { box-sizing: border-box; }
html { -webkit-text-size-adjust: 100%; }
body {
  margin: 0; padding: 0; background: #ffffff; color: #24292f;
  font-family: "Noto Sans CJK SC","Source Han Sans SC","Source Han Sans CN",
               "PingFang SC","Hiragino Sans GB","Microsoft YaHei",
               "Helvetica Neue", Arial, sans-serif;
  font-size: 15px; line-height: 1.75;
  -webkit-font-smoothing: antialiased;
}
.markdown-body { max-width: 880px; margin: 0 auto; padding: 28px 22px 56px; }
h1, h2, h3, h4, h5, h6 { font-weight: 600; line-height: 1.3; margin: 1.6em 0 .6em; }
h1 { font-size: 2em; border-bottom: 1px solid #e5e7eb; padding-bottom: .3em; }
h2 { font-size: 1.5em; border-bottom: 1px solid #eef0f2; padding-bottom: .25em; }
h3 { font-size: 1.25em; }
h4 { font-size: 1.05em; }
h1:first-child, h2:first-child { margin-top: 0; }
p { margin: .8em 0; }
a { color: #0969da; text-decoration: none; }
a:hover { text-decoration: underline; }
strong { font-weight: 600; }
ul, ol { padding-left: 1.7em; margin: .8em 0; }
li { margin: .3em 0; }
li > ul, li > ol { margin: .2em 0; }
blockquote {
  margin: 1em 0; padding: .6em 1.1em; color: #57606a;
  border-left: 4px solid #d0d7de; background: #f6f8fa; border-radius: 0 6px 6px 0;
}
blockquote > :first-child { margin-top: 0; }
blockquote > :last-child { margin-bottom: 0; }
code {
  font-family: "JetBrains Mono","Cascadia Code","Fira Code",Consolas,
               "Courier New","Noto Sans Mono CJK SC",monospace;
  background: #f3f4f6; padding: .15em .38em; border-radius: 4px; font-size: .9em;
}
pre {
  background: #f6f8fa; padding: 14px 16px; border-radius: 8px;
  overflow: auto; line-height: 1.55; border: 1px solid #eaecef;
}
pre code { background: transparent; padding: 0; font-size: .87em; }
table { border-collapse: collapse; width: 100%; margin: 1em 0; font-size: .95em; }
th, td { border: 1px solid #dfe2e5; padding: 8px 12px; text-align: left; vertical-align: top; }
th { background: #f6f8fa; font-weight: 600; }
tbody tr:nth-child(even) td { background: #fafbfc; }
img { max-width: 100%; height: auto; }
hr { border: 0; border-top: 1px solid #e5e7eb; margin: 2em 0; }
.toc {
  background: #f6f8fa; border: 1px solid #e5e7eb; border-radius: 8px;
  padding: 12px 20px; margin: 0 0 1.6em;
}
.toc > ul { padding-left: 1em; }
.toc ul { list-style: none; }
.toc a { color: #24292f; }
.toc a:hover { color: #0969da; }
/* 打印优化 */
pre, table, blockquote, img { page-break-inside: avoid; }
h1, h2, h3, h4 { page-break-after: avoid; }
"""

PAGE_TEMPLATE = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{title}</title>
{base}
<style>
{css}
</style>
</head>
<body>
<article class="markdown-body">
{toc}
{body}
</article>
</body>
</html>
"""


def render_page(title: str, body: str, toc: str = "", extra_css: str = "",
                base_href: str | None = None) -> str:
    css = BASE_CSS + "\n" + pygments_css() + "\n" + (extra_css or "")
    base = f'<base href="{base_href}">' if base_href else ""
    toc_block = ""
    if toc:
        toc_block = f'<nav class="toc">\n{toc}\n</nav>'
    return PAGE_TEMPLATE.format(
        title=html_escape(title), base=base, css=css, toc=toc_block, body=body
    )


def html_escape(s: str) -> str:
    return (s.replace("&", "&amp;").replace("<", "&lt;")
             .replace(">", "&gt;").replace('"', "&quot;"))


# ---------------------------------------------------------------------------
# HTML -> PDF（多引擎降级）
# ---------------------------------------------------------------------------
def html_to_pdf(html_path: str, pdf_path: str, soffice: str | None,
                engine: str = "auto") -> str:
    """把 HTML 文件转成 PDF，返回实际使用的引擎名。"""
    os.makedirs(os.path.dirname(os.path.abspath(pdf_path)), exist_ok=True)

    if engine in ("auto", "weasyprint") and has_weasyprint():
        try:
            import weasyprint
            weasyprint.HTML(filename=html_path).write_pdf(pdf_path)
            if os.path.getsize(pdf_path) > 0:
                return "weasyprint"
        except Exception:
            if engine == "weasyprint":
                raise
            # 继续尝试下一个引擎

    if engine in ("auto", "chromium"):
        chrome = find_chromium()
        if chrome:
            uri = "file://" + os.path.abspath(html_path)
            cmd = [
                chrome, "--headless", "--no-sandbox", "--disable-gpu",
                f"--print-to-pdf={os.path.abspath(pdf_path)}",
                "--no-pdf-header-footer", "--print-to-pdf-no-header",
                uri,
            ]
            try:
                subprocess.run(cmd, capture_output=True, timeout=180)
            except Exception:
                pass
            if os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 0:
                return "chromium"
            if engine == "chromium":
                raise ConvertError("Chromium 打印 PDF 失败")

    if engine in ("auto", "soffice"):
        if not soffice:
            raise ConvertError(
                "没有可用的 HTML→PDF 引擎：请安装 WeasyPrint / Chromium，"
                "或安装 LibreOffice。"
            )
        outdir = os.path.dirname(os.path.abspath(pdf_path))
        cmd = [
            soffice, "--headless", "--norestore", "--invisible",
            "--convert-to", "pdf:writer_web_pdf_Export",
            "--outdir", outdir, os.path.abspath(html_path),
        ]
        proc = subprocess.run(cmd, capture_output=True)
        produced = os.path.join(
            outdir, os.path.splitext(os.path.basename(html_path))[0] + ".pdf")
        if os.path.exists(produced) and os.path.abspath(produced) != os.path.abspath(pdf_path):
            os.replace(produced, pdf_path)
        if not (os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 0):
            err = (proc.stderr or proc.stdout).decode("utf-8", "replace")[:400]
            raise ConvertError(f"LibreOffice HTML→PDF 失败：{err}")
        return "soffice"

    raise ConvertError(f"未知的 PDF 引擎：{engine}")


# ---------------------------------------------------------------------------
# 单文件转换
# ---------------------------------------------------------------------------
def run_soffice(soffice: str, target: str, outdir: str, src: str) -> str:
    """用 LibreOffice 把 src 转成 target（html/pdf），返回产物路径。"""
    cmd = [
        soffice, "--headless", "--norestore", "--invisible",
        "--convert-to", target, "--outdir", outdir, os.path.abspath(src),
    ]
    proc = subprocess.run(cmd, capture_output=True)
    base = os.path.splitext(os.path.basename(src))[0]
    ext = "pdf" if target.startswith("pdf") else "html"
    produced = os.path.join(outdir, f"{base}.{ext}")
    if not os.path.exists(produced):
        err = (proc.stderr or proc.stdout).decode("utf-8", "replace")[:400]
        raise ConvertError(f"LibreOffice 转换失败（{target}）：{err}")
    return produced


def convert_markdown(src: str, outdir: str, want_html: bool, want_pdf: bool,
                     opts, soffice: str | None) -> list[str]:
    with open(src, "r", encoding="utf-8", errors="replace") as f:
        text = f.read()

    base = os.path.splitext(os.path.basename(src))[0]
    body, toc = md_to_fragment(text, with_toc=opts.toc)

    extra_css = ""
    if opts.css:
        with open(opts.css, "r", encoding="utf-8") as f:
            extra_css = f.read()

    src_dir = os.path.dirname(os.path.abspath(src))
    page = render_page(
        title=opts.title or base, body=body, toc=toc,
        extra_css=extra_css, base_href="file://" + src_dir + "/",
    )

    outputs: list[str] = []
    html_path = os.path.join(outdir, base + ".html")
    pdf_path = os.path.join(outdir, base + ".pdf")

    if want_html:
        with open(html_path, "w", encoding="utf-8") as f:
            f.write(page)
        outputs.append(html_path)

    if want_pdf:
        if want_html:
            engine = html_to_pdf(html_path, pdf_path, soffice, opts.engine)
        else:
            fd, tmp = tempfile.mkstemp(suffix=".html")
            os.close(fd)
            with open(tmp, "w", encoding="utf-8") as f:
                f.write(page)
            try:
                engine = html_to_pdf(tmp, pdf_path, soffice, opts.engine)
            finally:
                os.remove(tmp)
        if opts.verbose:
            print(f"      PDF 引擎：{engine}")
        outputs.append(pdf_path)

    return outputs


def convert_office(src: str, outdir: str, want_html: bool, want_pdf: bool,
                   opts, soffice: str | None) -> list[str]:
    if not soffice:
        raise ConvertError(
            "Office 文档转换需要 LibreOffice，未找到 soffice。"
            "请安装 LibreOffice 或用 --soffice 指定路径。"
        )
    outputs: list[str] = []
    if want_html:
        outputs.append(run_soffice(soffice, "html", outdir, src))
    if want_pdf:
        outputs.append(run_soffice(soffice, "pdf", outdir, src))
    return outputs


def convert_one(src: str, outdir: str, want_html: bool, want_pdf: bool,
                opts, soffice: str | None) -> list[str]:
    ext = os.path.splitext(src)[1].lower()
    if ext not in ALL_EXTS:
        raise ConvertError(f"不支持的格式：{src}")
    base = os.path.splitext(os.path.basename(src))[0]

    # 覆盖检查
    planned = []
    if want_html:
        planned.append(os.path.join(outdir, base + ".html"))
    if want_pdf:
        planned.append(os.path.join(outdir, base + ".pdf"))
    if not opts.overwrite:
        exist = [p for p in planned if os.path.exists(p)]
        if exist and len(exist) == len(planned):
            return ["__SKIP__"]

    os.makedirs(outdir, exist_ok=True)
    if ext in MD_EXTS:
        return convert_markdown(src, outdir, want_html, want_pdf, opts, soffice)
    return convert_office(src, outdir, want_html, want_pdf, opts, soffice)


# ---------------------------------------------------------------------------
# 输入收集
# ---------------------------------------------------------------------------
def collect_inputs(paths: list[str], recursive: bool) -> list[str]:
    files: list[str] = []
    seen: set[str] = set()

    def add(p: str):
        ap = os.path.abspath(p)
        if ap not in seen:
            seen.add(ap)
            files.append(p)

    for p in paths:
        if os.path.isfile(p):
            add(p)
        elif os.path.isdir(p):
            it = os.walk(p) if recursive else [(p, [], os.listdir(p))]
            for root, _dirs, names in it:
                for n in names:
                    fp = os.path.join(root, n)
                    if os.path.isfile(fp) and os.path.splitext(n)[1].lower() in ALL_EXTS:
                        add(fp)
        else:
            matches = globmod.glob(p)
            for m in matches:
                if os.path.isfile(m) and os.path.splitext(m)[1].lower() in ALL_EXTS:
                    add(m)
    return files


# ---------------------------------------------------------------------------
# --check
# ---------------------------------------------------------------------------
def do_check(soffice_arg: str | None) -> int:
    ok = True
    print("依赖检查：")
    m = has_markdown()
    print(f"  [{'✔' if m else '✘'}] Python markdown "
          f"{'已安装' if m else '缺失 → pip install markdown'}")
    ok &= m

    if has_weasyprint():
        import weasyprint
        print(f"  [✔] WeasyPrint {weasyprint.__version__}（HTML→PDF 首选）")
    else:
        print("  [ ] WeasyPrint 未安装（可选，pip install weasyprint）")

    chrome = find_chromium()
    print(f"  [{'✔' if chrome else ' '}] Chromium {'：' + chrome if chrome else '未找到（可选）'}")

    soffice = find_soffice(soffice_arg) if True else None
    if soffice:
        try:
            out = subprocess.run([soffice, "--version"], capture_output=True,
                                 text=True, timeout=30).stdout.strip()
        except Exception:
            out = ""
        print(f"  [✔] LibreOffice：{soffice}")
        if out:
            print(f"        {out.splitlines()[0]}")
    else:
        print("  [✘] LibreOffice（soffice）未找到 —— Office 转换与兜底 PDF 需要它")
        ok = False
    print()
    print("结论：" + ("依赖齐全，可以使用。" if ok else "存在缺失，请先补齐上述 ✘ 项。"))
    return 0 if ok else 2


# ---------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------
def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="doc2htmlpdf",
        description="将 Markdown / Office 文档批量转换为 HTML 和/或 PDF。",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument("inputs", nargs="*", help="输入文件 / 目录 / 通配符")
    parser.add_argument("-o", "--outdir", default=".", help="输出目录（默认当前目录）")
    parser.add_argument("--to", choices=["html", "pdf", "both"], default="both",
                        help="输出目标（默认 both）")
    parser.add_argument("-r", "--recursive", action="store_true", help="递归处理目录")
    parser.add_argument("--toc", action="store_true", help="为 Markdown 生成目录")
    parser.add_argument("--title", help="HTML <title>（默认取文件名）")
    parser.add_argument("--css", help="附加自定义 CSS 文件")
    parser.add_argument("--engine", choices=["auto", "weasyprint", "chromium", "soffice"],
                        default="auto", help="HTML→PDF 引擎（默认 auto 自动降级）")
    parser.add_argument("--soffice", help="soffice 可执行文件路径")
    parser.add_argument("--overwrite", action="store_true", help="覆盖已存在的输出")
    parser.add_argument("--check", action="store_true", help="检查依赖后退出")
    parser.add_argument("--version", action="version", version=f"doc2htmlpdf {VERSION}")
    parser.add_argument("-v", "--verbose", action="store_true", help="输出详细信息")

    opts = parser.parse_args(argv)

    if opts.check:
        return do_check(opts.soffice)

    if not opts.inputs:
        parser.print_help()
        return 2

    want_html = opts.to in ("html", "both")
    want_pdf = opts.to in ("pdf", "both")

    files = collect_inputs(opts.inputs, opts.recursive)
    if not files:
        print("未找到任何支持的文件（.md/.docx/.pptx/.xlsx 等）。", file=sys.stderr)
        return 2

    # 只有需要 Office 或需要 soffice 兜底时才定位 soffice
    need_soffice = any(os.path.splitext(f)[1].lower() in OFFICE_EXTS for f in files)
    try:
        soffice = find_soffice(opts.soffice)
    except ConvertError as e:
        print(f"错误：{e}", file=sys.stderr)
        return 2
    if need_soffice and not soffice:
        print("错误：输入包含 Office 文档，但未找到 LibreOffice（soffice）。",
              file=sys.stderr)
        return 2
    if want_pdf and not soffice and not has_weasyprint() and not find_chromium():
        print("错误：需要 HTML→PDF 引擎，但未找到 WeasyPrint / Chromium / LibreOffice。",
              file=sys.stderr)
        return 2

    print(f"共发现 {len(files)} 个文件，输出：{opts.to} → {opts.outdir}")
    ok = fail = skip = 0
    failures: list[tuple[str, str]] = []

    for i, src in enumerate(files, 1):
        name = os.path.relpath(src)
        print(f"[{i}/{len(files)}] {name}")
        try:
            results = convert_one(src, opts.outdir, want_html, want_pdf, opts, soffice)
            if results == ["__SKIP__"]:
                print("      跳过（已存在，用 --overwrite 覆盖）")
                skip += 1
                continue
            for r in results:
                print(f"      -> {r}")
            ok += 1
        except ConvertError as e:
            print(f"      [失败] {e}", file=sys.stderr)
            failures.append((src, str(e)))
            fail += 1
        except Exception as e:  # noqa: BLE001
            print(f"      [失败] {type(e).__name__}: {e}", file=sys.stderr)
            failures.append((src, f"{type(e).__name__}: {e}"))
            fail += 1

    print(f"\n完成：成功 {ok}，失败 {fail}" + (f"，跳过 {skip}" if skip else ""))
    if failures:
        print("失败清单：")
        for src, err in failures:
            print(f"  - {src}\n    {err}")
    return 0 if fail == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
