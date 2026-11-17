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

仓库自带静态页面 `docs/index.html`，提供**两条**运行路径，默认走第一条：

| 模式 | 访客要做什么 | 适用 |
|---|---|---|
| **公开服务（默认）** | 什么都不用做——打开页面、拖文件、下载 | 给别人用、当公共小工具 |
| GitHub Actions | 自己填一个 PAT（写权限在你仓库） | 只有自己用、不想额外部署服务 |

**Pages 本身能不能直接转换？不能。** GitHub Pages 只发 HTML/CSS/JS，没有能执行 `soffice` 的进程；浏览器里也没有可用的 LibreOffice WASM 全功能构建，纯前端转 docx 排版必然走样（公式、图表、分页全崩）。所以只要要「真 LibreOffice 的保真度」，就必须有一台能跑 `soffice` 的机器。

于是所有「不用 token」的方案都收敛到同一件事：**你自己部署一次服务，替所有访客承担那一次认证**。

```
访客浏览器 (Pages 静态页)
  │  ① POST 文件（FormData，无任何凭据）
  ▼
你的转换服务（server/ · Docker + LibreOffice）
  │  ② 服务端转换
  ▼
访客浏览器 ③ 直接拿回 PDF，自动下载
```

### 1. 部署转换服务（一次性，约 1 分钟）

服务镜像使用 `server/` 中的配置以及仓库根目录的 `office2pdf.py`。

**A. Hugging Face Space（Docker Space 的账号和费用要求以平台为准）**

1. 新建 Space → SDK 选 **Docker** → 模板选 Blank
2. 将 `server/README.md`、`server/Dockerfile` 和 `office2pdf.py` 放到 Space 根目录；将 `server/app.py` 和 `server/requirements.txt` 放到 Space 的 `server/` 目录
3. 等构建完成，拿到地址 `https://<user>-<space>.hf.space`

**B. 自己 VPS / 任意 Docker 主机**

```bash
docker build -f server/Dockerfile -t office2pdf-api .
docker run -d --name office2pdf -p 7860:7860 --restart unless-stopped office2pdf-api
# 前面挂 Caddy/Nginx 拿 HTTPS，并在反代层加限流
```

**C. 本机先试**

```bash
pip install -r server/requirements.txt
uvicorn server.app:app --host 0.0.0.0 --port 8000
curl http://127.0.0.1:8000/health
```

### 2. 把地址接进页面

三种方式，任选其一：

- 页面顶部「转换服务地址」填一次（存 localStorage，之后一直生效），点「检查连接」
- 链接直接带参数：`https://<owner>.github.io/doc-converter/?api=https://your-space.hf.space`
- 部署者写死在 `docs/index.html` 的 `const DEFAULT_API = "";`，访客打开即用

之后任何人打开页面就能拖文件转换，**不用 GitHub 账号、不用 token**，文件只发往你自己这台服务。

### 3. Pages 发布

推到 GitHub → **Settings → Pages → Source 选 `GitHub Actions`** → 跑一次 `pages` 工作流，
得到 `https://<owner>.github.io/doc-converter/`。（推送改动到 `docs/**` 时会自动重新部署。）

### 限额（`server/` 的环境变量）

| 变量 | 默认 | 说明 |
|---|---|---|
| `MAX_FILE_MB` | 40 | 单文件上限，公开部署建议 20 |
| `MAX_CONCURRENCY` | 2 | 同时转换数，超出排队（LibreOffice 很吃内存） |
| `CONVERT_TIMEOUT` | 600 | 单文件超时（秒） |
| `ALLOWED_ORIGINS` | `*` | CORS 允许来源，可收紧到你的 Pages 域名 |

服务端对每个请求用独立临时目录、独立 LibreOffice profile（并发安全），响应结束即清理；
文件名白名单化，拒绝路径穿越；只接受 `office2pdf.py` 支持的扩展名。

### Actions 模式（备用）

不开服务也能用，代价是使用者得自备 token：页面把文件提交到仓库 `inbox/<job>/` → 触发
`convert` 工作流 → Actions 里转换并提交到 `outbox/<job>/` → 页面轮询 `_status.json` 后下载。

- 单文件 20 MB（Contents API + base64 膨胀）
- 单作业 60 分钟超时
- token 需要 `repo` + `workflow`（经典）或 `Contents: RW` + `Actions: R`（细粒度）

这条路**不适合当公共入口**：任何拿到 token 的人都能改写你的仓库。所以它只是备用，不是默认。

## 目录结构

```
doc-converter/
├── doc2htmlpdf.py         # 主工具：md/office → html/pdf
├── office2pdf.py          # Office/MD/PDF → PDF（含图片型、CLI 输出渲染）
├── upgrade_libreoffice.sh # LibreOffice 升级脚本
├── docs/index.html        # GitHub Pages 可视化批量转换页面（公开服务 / Actions 双模式）
├── server/                # 公开转换服务（FastAPI + Dockerfile，可直接传成 HF Space）
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
