# doc2htmlpdf 示例文档

这是一个用于验证 **Markdown → HTML / PDF** 转换的示例文件，包含中文、代码、表格、列表与引用。

## 功能一览

| 能力 | 输入 | 输出 |
|---|---|---|
| Markdown | `.md` | HTML / PDF |
| Word | `.docx` / `.doc` | HTML / PDF |
| PowerPoint | `.pptx` / `.ppt` | HTML / PDF |
| Excel | `.xlsx` / `.xls` | HTML / PDF |

## 有序列表

1. 安装依赖
2. 运行转换
3. 查看输出

## 无序列表

- 支持批量处理
- 支持目录递归（`-r`）
- 支持 4 种 PDF 引擎自动降级

## 代码块

```python
def hello(name: str) -> str:
    """返回问候语"""
    return f"你好，{name}！"

print(hello("世界"))
```

行内代码：`python doc2htmlpdf.py README.md --to both`

## 引用

> 好的工具应该让人少想一步，而不是多配一遍环境。

## 链接与强调

访问 [GitHub](https://github.com) 了解更多；*斜体*、**粗体**、~~删除线~~。

### 三级标题

更深层级的内容用于测试目录（`--toc`）生成效果。
