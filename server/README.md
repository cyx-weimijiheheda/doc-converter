---
title: Office2PDF
emoji: 📄
colorFrom: blue
colorTo: green
sdk: docker
app_port: 7860
pinned: false
---

# Office2PDF 公开转换服务

把 Word / PPT / Excel / Markdown / PDF 批量转成 PDF，给 GitHub Pages 页面当后端。
访问者打开网页拖文件即可，**不需要 GitHub 账号、不需要任何 token**。

## 上传成 Hugging Face Space

将 `server/README.md` 和 `server/Dockerfile` 放到 Space 根目录，
将 `server/app.py`、`server/requirements.txt` 放到 Space 的 `server/` 目录，
并将仓库根目录的 `office2pdf.py` 放到 Space 根目录。构建使用这些文件的当前版本。

拿到 Space 地址（形如 `https://<user>-<space>.hf.space`）后，填进
`docs/index.html` 页面顶部的「服务地址」，或直接用
`https://<owner>.github.io/doc-converter/?api=https://<user>-<space>.hf.space`。

## Vercel 容器部署

将整个仓库导入 Vercel；根目录的 `Dockerfile.vercel` 用于构建服务。
部署后访问服务地址的 `/health`，确认 LibreOffice 就绪，再把服务地址填入 Pages。
Vercel Functions 的请求和响应大小有限制，镜像默认将输入和输出限制为 4 MiB；
服务可能冷启动，资源与运行时间受所选套餐限制。公开使用前应部署额外的限流保护。

## 本地 / 自建 VPS

```bash
pip install -r server/requirements.txt
uvicorn server.app:app --host 0.0.0.0 --port 8000
# 或
docker build -f server/Dockerfile -t office2pdf-api . && docker run -p 7860:7860 office2pdf-api
```

## 接口

| 方法 | 路径 | 说明 |
|---|---|---|
| GET | `/health` | 健康检查（LibreOffice / md2pdf / 限额） |
| POST | `/api/convert` | `file` + `mode=text\|image` + `dpi=60..300` → `application/pdf` |
| GET | `/docs` | OpenAPI 文档 |

## 限额（环境变量可调）

| 变量 | 默认 | 说明 |
|---|---|---|
| `MAX_FILE_MB` | 40 | 单文件大小上限 |
| `MAX_OUTPUT_MB` | 0 | 输出大小上限，0 表示不限制（Vercel 镜像默认 4） |
| `MAX_CONCURRENCY` | 2 | 同时转换数，超出排队 |
| `CONVERT_TIMEOUT` | 600 | 单文件转换超时（秒） |
| `ALLOWED_ORIGINS` | `*` | CORS 允许来源，逗号分隔 |

公开部署时建议把 `MAX_FILE_MB` 调到 20、`MAX_CONCURRENCY` 保持 2，并在反代层加限流。
