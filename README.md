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
python office2pdf.py notes.md                          # Markdown → PDF
python office2pdf.py ./docs -r -o out/                 # 批量递归
python office2pdf.py ./docs -o out --json              # 机器可读结果
python office2pdf.py --check                           # 环境检查
```

- **文本型 PDF**：可选中、可搜索、体积小（LibreOffice headless）
- **图片型 PDF**：每页栅格化后打包，不可选中文字、版式 100% 一致（poppler + img2pdf）
- **输出显示**：彩色状态行、逐文件耗时/体积/页数、结尾汇总与失败清单，CJK 按双宽对齐

图片型 PDF 额外依赖：`poppler-utils`（`pdftoppm`、`pdfinfo`）与 `img2pdf`（或 Pillow 兜底）。

| 参数 | 说明 |
|---|---|
| `-o, --outdir` | 输出目录，默认与源文件同目录 |
| `-r, --recursive` | 递归处理目录 |
| `--image` | Office → 图片型 PDF（PDF 输入自动图片化，无需此标志） |
| `--dpi` | 图片型分辨率，默认 150 |
| `--img-format` | 栅格化格式：`png`（默认）/`jpg`/`jpeg`/`tiff`/`ppm` |
| `--overwrite` | 覆盖已存在的同名 PDF（默认跳过） |
| `--timeout` | 单文件超时秒数，默认 600 |
| `--soffice` | 显式指定 soffice 路径 |
| `-v, --verbose` | 输出完整输入路径等细节 |
| `-q, --quiet` | 只输出错误与最终统计 |
| `--json` | 纯 JSON 结果（此时不打印人读表格） |
| `--no-color` / `--no-progress` | 关闭颜色 / 动态进度（日志、CI 友好） |
| `-c, --check` | 环境自检 |

颜色与进度会自动判定：`NO_COLOR`、`FORCE_COLOR`、`TERM=dumb`、是否 TTY 均生效；输出被重定向或管道时退化为纯文本行。

```
Office2PDF  4 个文件 · Office→文本型 + Markdown→PDF
────────────────────────────────────────────────────────────────────
✓ [1/4] 测试文档.docx                 →  测试文档.pdf  1 页 · 54.2 KB · 1.5s
✓ [2/4] 演示文稿.pptx                 →  演示文稿.pdf  2 页 · 20.9 KB · 527ms
✓ [3/4] 表格.xlsx                     →  表格.pdf  2 页 · 26.6 KB · 1.4s
↷ [4/4] 测试文档.md                   →  测试文档.pdf  跳过：输出已存在，用 --overwrite 覆盖
────────────────────────────────────────────────────────────────────
汇总  成功 3 · 跳过 1 · 失败 0
      总耗时 3.5s · 产出 101.7 KB · 平均 866ms/个
```

退出码：全部成功 `0`，有失败 `1`，参数/输入错误 `2`。

### upgrade_libreoffice.sh —— 修复老 PPT 转换崩溃

老式二进制 `.ppt` 若包含 Microsoft Equation 3.0 公式对象，旧版 LibreOffice 会崩溃（`exit=134`）。本脚本在**不使用 flatpak** 的前提下把 LibreOffice 升级到 26.8+：

```bash
bash upgrade_libreoffice.sh              # 升级 + 清理冗余下载包
bash upgrade_libreoffice.sh --no-clean   # 只升级
bash upgrade_libreoffice.sh --clean-only # 只清理
```

升级路径：先试 `dnf` 仓库 → 不够则下载官方 RPM 包、本地提取、在 `/usr/local/bin` 生成 `soffice` 包装脚本。

## 可视化批量转换（GitHub Pages）

仓库自带一个静态页面 `docs/index.html`：拖进文件、选模式、点一下，就能在 GitHub Actions 的 Ubuntu 机器上用真实 LibreOffice 批量转换并下载 PDF。

**Pages 本身能不能直接转换？不能。** GitHub Pages 只发 HTML/CSS/JS，没有能执行 `soffice` 的后端进程，浏览器里也没有可用的 LibreOffice WASM 全功能构建——纯前端转 docx 排版必然走样。所以这里采用「Pages 当门面 + Actions 当后端」：

```
浏览器 (Pages 静态页)
  │  ① 用你的 token 把文件 PUT 到仓库 inbox/<job>/
  │  ② POST /dispatches 触发 convert 工作流
  ▼
GitHub Actions (ubuntu-latest + 真实 LibreOffice)
  │  ③ office2pdf.py 转换 → 提交 outbox/<job>/
  ▼
浏览器 ④ 轮询 outbox/<job>/_status.json → 逐个下载 PDF
```

### 启用步骤

1. 推到 GitHub，默认分支为 `main`（`repository_dispatch` 只监听默认分支上的工作流）。
2. **Settings → Pages → Build and deployment → Source 选 `GitHub Actions`**。
3. 跑一次 `pages` 工作流（推送改动到 `docs/**` 时自动触发），得到 `https://<owner>.github.io/doc-converter/`。
4. 打开页面填 `owner` / `repo` / `branch`（项目站点会自动猜）与 Personal Access Token：
   - 经典 token：勾 `repo` + `workflow`
   - 细粒度 token：`Contents: Read and write`、`Actions: Read and write`
5. 选文件 → 选模式（文本型 / 图片型 + DPI）→ **开始转换**，通常 1–3 分钟。

### 限制（照实说）

| 限制 | 数值 / 原因 |
|---|---|
| 单文件大小 | 页面硬限制 20 MB（Contents API 上限 100 MB，浏览器 base64 再膨胀 1.33 倍） |
| 单次作业时长 | 60 分钟超时（`convert.yml` 的 `timeout-minutes`） |
| 并发 | 同一 `job` 串行、不同 `job` 并行 |
| 私有仓库 | 消耗 Actions 分钟数；Pages 可见性取决于套餐 |
| 结果留存 | PDF 会提交进 `outbox/`，长期使用需定期清理旧作业目录 |
| 安全 | token 只存在浏览器 localStorage，可一键清除；页面不经过第三方服务器 |

要给「不用填 token 的人」用，就需要真后端：把 `office2pdf.py` 套一层 FastAPI，Dockerfile 里 `apt install libreoffice poppler-utils`，部署到 Hugging Face Spaces / Fly.io / Render，或自建 VPS 反代并加限流。

## 目录结构

```
doc-converter/
├── doc2htmlpdf.py         # 主工具：md/office → html/pdf
├── office2pdf.py          # Office/MD/PDF → PDF（含图片型、CLI 输出渲染）
├── upgrade_libreoffice.sh # LibreOffice 升级脚本
├── docs/index.html        # GitHub Pages 可视化批量转换页面
├── .github/workflows/
│   ├── pages.yml          # 发布 docs/ 到 Pages
│   └── convert.yml        # Actions 后端：装 LibreOffice 并批量转换
├── examples/
│   └── sample.md          # 示例 Markdown
├── requirements.txt
├── LICENSE
└── README.md
```

## 许可

[MIT](LICENSE)
