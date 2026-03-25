# doc-conversion Skill

将 Word (.docx) 和 PDF 文件高质量转换为 Markdown 格式的 Claude Code skill。

## 功能特性

- **Word (.docx) → Markdown**：保留标题层级、表格、图片、列表等格式
- **PDF → Markdown**：支持文字型 PDF 和扫描版 PDF（OCR）
- **批量转换**：一键转换整个目录的文档
- **智能降级**：缺少工具时自动使用备选方案

## 快速安装

### 方式一：安装脚本（推荐）

```bash
# 克隆仓库
git clone <your-repo-url> ~/doc-conversion

# 运行安装脚本
cd ~/doc-conversion
./install.sh
```

### 方式二：手动安装

```bash
# 创建软链接
ln -s /path/to/doc-conversion ~/.claude/skills/doc-conversion
```

### 方式三：直接复制

```bash
cp -r /path/to/doc-conversion ~/.claude/skills/doc-conversion
```

## 使用方法

在 Claude Code 中：

1. 输入 `/skills`
2. 选择 `doc-conversion`
3. 提供要转换的文件路径

### 转换示例

```
把 /path/to/document.docx 转成 markdown 格式
这个 Word 文档里有表格，帮我转成 markdown，表格要保留
将这个 PDF 转换为 markdown：/path/to/manual.pdf
这是个扫描版 PDF，帮我转成 markdown：/path/to/scanned.pdf
把 /path/to/docs/ 目录下所有 PDF 文件都转成 markdown
```

## 系统依赖

### 基础工具

```bash
# macOS
brew install pandoc poppler tesseract

# Linux (Ubuntu/Debian)
sudo apt install pandoc poppler-utils tesseract-ocr

# Linux (CentOS/RHEL)
sudo yum install pandoc poppler-utils tesseract
```

### Python 库

```bash
pip install python-docx pymupdf pdf2image pytesseract markdownify
```

## 命令行用法

也可以直接使用转换脚本：

```bash
# 单文件转换
python scripts/convert.py document.docx output.md
python scripts/convert.py document.pdf output.md

# OCR 模式（扫描版 PDF）
python scripts/convert.py --ocr scanned.pdf output.md

# 批量转换
python scripts/convert.py --batch ./input/dir ./output/dir
```

## 输出目录结构

```
output/
├── document.md              # 主 Markdown 文件
└── document_assets/         # 资源目录（如有）
    ├── image-1.png
    ├── image-2.png
    └── table-1.html         # 复杂表格
```

## 支持的文件格式

| 格式 | 支持程度 | 说明 |
|------|----------|------|
| .docx | 完全支持 | 保留所有格式 |
| .doc | 部分支持 | 需要额外工具 |
| PDF (文字型) | 完全支持 | 直接提取文字 |
| PDF (扫描版) | 支持 | 需要 OCR |

## 故障排查

### pandoc 未安装

```
解决：brew install pandoc
备选：自动使用 python-docx 库
```

### pdftotext 未安装

```
解决：brew install poppler
备选：自动使用 pymupdf 库
```

### OCR 失败

```
解决：brew install tesseract
需要：pip install pdf2image pytesseract
```

## 卸载

```bash
rm -rf ~/.claude/skills/doc-conversion
```

## License

MIT
