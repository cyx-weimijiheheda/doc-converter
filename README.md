# doc2htmlpdf

把 **Markdown / Office 文档**批量转换成 **HTML 和 / 或 PDF** 的命令行工具。纯命令行、零 GUI、跨平台，Linux 上开箱即用。

> 顺带附赠一个 `office2pdf.py`（Office → 文本型/图片型 PDF）和一个 `upgrade_libreoffice.sh`（一键升级 LibreOffice，修复老 PPT 转换崩溃）。

---

## 功能特性

- **多种输入**：`.md` `.markdown` `.mdown` `.mkd`，以及 `.doc` `.docx` `.odt` `.rtf` `.ppt` `.pptx` `.odp` `.xls` `.xlsx` `.ods` `.csv`
- **两种输出**：HTML、PDF，或同时生成
- **批量 & 递归**：支持多文件、目录、通配符、`-r` 递归
- **Markdown 排版**：GitHub 风格主题，支持表格、围栏代码块、代码高亮、脚注、目录（`--toc`）
- **PDF 引擎自动降级**：WeasyPrint → Chromium 无头 → LibreOffice，哪个可用用哪个
- **中文字体友好**：优先思源黑体 / Noto Sans CJK，PDF 字体自动嵌入
- **可定制**：`--css` 追加自定义样式，`--title` 指定标题

## 支持矩阵

| 输入 | → HTML | → PDF |
|---|:---:|:---:|
| Markdown (`.md` 等) | ✅ | ✅ |
| Word (`.doc` `.docx` `.odt` `.rtf`) | ✅ | ✅ |
| PowerPoint (`.ppt` `.pptx` `.odp`) | ✅ | ✅ |
| Excel (`.xls` `.xlsx` `.ods` `.csv`) | ✅ | ✅ |

## 安装

### 1. Python 依赖

```bash
pip install -r requirements.txt
```

`requirements.txt` 只强制 `markdown`；`weasyprint` 为可选项，装了能显著提升 Markdown→PDF 的排版质量。

### 2. 系统依赖

- **LibreOffice**（Office 文档转换必需；也是 HTML→PDF 的兜底引擎）
  - Fedora/RHEL：`sudo dnf install libreoffice libreoffice-writer libreoffice-calc libreoffice-impress`
  - Debian/Ubuntu：`sudo apt install libreoffice`
  - macOS：`brew install --cask libreoffice`
- **中文字体**（可选，避免方框）：`sudo dnf install google-noto-sans-cjk-fonts` / `sudo apt install fonts-noto-cjk`

检查环境是否就绪：

```bash
python doc2htmlpdf.py --check
```

## 快速开始

```bash
# Markdown → HTML + PDF
python doc2htmlpdf.py README.md

# 只生成 PDF，并生成目录
python doc2htmlpdf.py README.md --to pdf --toc

# 批量 + 指定输出目录
python doc2htmlpdf.py a.md b.docx c.xlsx -o out/

# 递归转换整个目录
python doc2htmlpdf.py ./docs -r --to both -o out/

# Office 文档 → HTML
python doc2htmlpdf.py 报告.docx 幻灯片.pptx --to html

# 指定 PDF 引擎
python doc2htmlpdf.py note.md --to pdf --engine weasyprint
```

## 参数说明

| 参数 | 说明 |
|---|---|
| `inputs` | 输入文件 / 目录 / 通配符（可多个） |
| `-o, --outdir` | 输出目录（默认当前目录） |
| `--to {html,pdf,both}` | 输出目标，默认 `both` |
| `-r, --recursive` | 递归处理目录 |
| `--toc` | 为 Markdown 生成目录（放在正文前） |
| `--title` | HTML `<title>`（默认取文件名） |
| `--css FILE` | 追加自定义 CSS 文件 |
| `--engine {auto,weasyprint,chromium,soffice}` | HTML→PDF 引擎，默认 `auto` 自动降级 |
| `--soffice PATH` | 显式指定 soffice 路径 |
| `--overwrite` | 覆盖已存在的输出（默认跳过） |
| `--check` | 检查依赖后退出 |
| `-v, --verbose` | 显示所用引擎等细节 |
| `--version` | 显示版本 |

## PDF 引擎说明

Markdown 先转成带样式的 HTML，再渲染成 PDF。引擎按 `auto` 顺序自动选择：

1. **WeasyPrint** —— 纯 Python，CSS 支持好、排版精致，首选（`pip install weasyprint`）
2. **Chromium / Chrome** —— 无头 `--print-to-pdf`，渲染最接近浏览器
3. **LibreOffice** —— 兜底，任何装了 LO 的机器都能用（CSS 支持有限）

Office 文档则直接由 LibreOffice 转 HTML / PDF，保真度最高。

## 附赠工具

### office2pdf.py —— Office / PDF 专用转换

```bash
python office2pdf.py report.docx                       # Office → 文本型 PDF
python office2pdf.py slides.pptx --image --dpi 200     # Office → 图片型 PDF
python office2pdf.py doc.pdf --dpi 300                 # 文本型 PDF → 图片型 PDF
python office2pdf.py ./docs -r -o out/                 # 批量递归
python office2pdf.py --check                           # 环境检查
```

- **文本型 PDF**：可选中、可搜索、体积小（LibreOffice headless）
- **图片型 PDF**：每页栅格化后打包，不可选中文字、版式 100% 一致（poppler + img2pdf）

图片型 PDF 额外依赖：`poppler-utils`（`pdftoppm`）与 `img2pdf`（或 Pillow 兜底）。

### upgrade_libreoffice.sh —— 修复老 PPT 转换崩溃

老式二进制 `.ppt` 若包含 Microsoft Equation 3.0 公式对象，旧版 LibreOffice 会崩溃（`exit=134`）。本脚本在**不使用 flatpak** 的前提下把 LibreOffice 升级到 26.8+：

```bash
bash upgrade_libreoffice.sh              # 升级 + 清理冗余下载包
bash upgrade_libreoffice.sh --no-clean   # 只升级
bash upgrade_libreoffice.sh --clean-only # 只清理
```

升级路径：先试 `dnf` 仓库 → 不够则下载官方 RPM 包、本地提取、在 `/usr/local/bin` 生成 `soffice` 包装脚本。

## 目录结构

```
doc2htmlpdf/
├── doc2htmlpdf.py        # 主工具：md/office → html/pdf
├── office2pdf.py         # Office/PDF → PDF（含图片型）
├── upgrade_libreoffice.sh# LibreOffice 升级脚本
├── examples/
│   └── sample.md         # 示例 Markdown
├── requirements.txt
├── LICENSE
└── README.md
```

## 许可

[MIT](LICENSE)
