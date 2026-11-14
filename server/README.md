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

## 上传成 Hugging Face Space（4 个文件）

在本目录下，把 `README.md`、`Dockerfile`、`app.py`、`requirements.txt` 传到新建的
Docker Space 根目录即可。构建时镜像会自己拉取最新的 `office2pdf.py`。

拿到 Space 地址（形如 `https://<user>-<space>.hf.space`）后，填进
`docs/index.html` 页面顶部的「服务地址」，或直接用
`https://<owner>.github.io/doc-converter/?api=https://<user>-<space>.hf.space`。

## 本地 / 自建 VPS

```bash
pip install -r server/requirements.txt
uvicorn server.app:app --host 0.0.0.0 --port 8000
# 或
docker build -t office2pdf-api server/ && docker run -p 7860:7860 office2pdf-api
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
| `MAX_CONCURRENCY` | 2 | 同时转换数，超出排队 |
| `CONVERT_TIMEOUT` | 600 | 单文件转换超时（秒） |
| `ALLOWED_ORIGINS` | `*` | CORS 允许来源，逗号分隔 |

公开部署时建议把 `MAX_FILE_MB` 调到 20、`MAX_CONCURRENCY` 保持 2，并在反代层加限流。
