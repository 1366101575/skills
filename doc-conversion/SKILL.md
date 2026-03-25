---
name: doc-conversion
description: 文档转换专家 - 将 Word (.docx) 和 PDF 文件转换为 Markdown 格式。支持保留复杂格式（表格、图片、标题层级）和处理扫描版 PDF（OCR）。当用户需要转换 .docx、.pdf 文件为 Markdown、需要提取文档内容、或提到"转成 markdown"/"转换为 md"时使用此 skill。支持单文件和批量转换。
---

# 文档转换 Skill

本 skill 专注于将 Word 文档和 PDF 文件高质量地转换为 Markdown 格式。

## 何时使用

- 用户需要转换 `.docx` 或 `.pdf` 文件为 Markdown
- 用户需要提取文档内容并保持格式
- 用户提到"转成 markdown"、"转换为 md"、"docx 转 markdown"、"pdf 转 markdown"等
- 用户需要批量转换多个文档

## 转换原则

### 1. 格式保留优先级

**Word 文档 (.docx):**
- 标题层级 → Markdown 标题 (`#`, `##`, `###`)
- 表格 → Markdown 表格（复杂表格使用 HTML）
- 图片 → Markdown 图片语法，保留原文说明
- 列表 → Markdown 列表 (`-`, `1.`)
- 代码块 → Markdown 代码块（带语言标识）
- 粗体/斜体 → `**text**` / `*text*`
- 链接 → Markdown 链接 `[text](url)`

**PDF 文档:**
- 文字型 PDF → 直接提取文字，保持原有结构
- 扫描版 PDF → 使用 OCR 识别文字
- 复杂排版 → 尽可能用 Markdown + HTML 还原

### 2. 输出质量要求

- 保持原文档的标题层级结构
- 表格必须完整保留（必要时使用 HTML 表格）
- 图片提取到独立目录，Markdown 中引用
- 清理页眉、页脚、页码等冗余内容
- 保持段落间距和可读性

## 工作流程

### Step 1: 识别输入

1. 确认用户要转换的文件路径
2. 检查文件类型（`.docx`、`.pdf`、`.md`）
3. 确认是否需要批量处理

### Step 2: 选择转换工具

根据文件类型选择合适的工具：

**Word → Markdown:**
- 首选 `pandoc`（如果已安装）：`pandoc input.docx -f docx -t markdown -o output.md`
- 备选：使用 Python `python-docx` + 自定义转换脚本

**PDF → Markdown:**
- 文字型 PDF：`pdftotext` 或 Python `pymupdf`
- 扫描版 PDF：Python `pdf2image` + `pytesseract` (OCR)
- 复杂 PDF：使用 `marker-pdf` 或 `nougat` 等深度学习模型

### Step 3: 执行转换

1. 创建输出目录（用于存放提取的图片等资源）
2. 执行转换命令
3. 后处理 Markdown 输出（清理、格式化）

### Step 4: 验证输出

1. 检查 Markdown 文件是否生成
2. 验证标题层级是否正确
3. 确认表格、图片已正确转换
4. 向用户展示转换结果

## 工具依赖

### 推荐安装

```bash
# 基础工具
brew install pandoc          # Word 转换
brew install poppler         # PDF 工具 (pdftotext, pdfimages)
brew install tesseract       # OCR 识别

# Python 库
pip install python-docx      # 读取 Word
pip install pymupdf          # 读取 PDF
pip install pdf2image        # PDF 转图片
pip install pytesseract      # OCR
pip install markdownify      # HTML 转 Markdown
```

### 备选方案

如果系统缺少某些工具，skill 应自动降级使用备选方案：
- 无 `pandoc` → 使用 Python `python-docx`
- 无 `tesseract` → 仅支持文字型 PDF，扫描版提示用户安装
- 无 `poppler` → 使用 `pymupdf` 替代

## 输出目录结构

转换后的文件组织：

```
output/
├── document.md              # 主 Markdown 文件
└── document_assets/         # 资源目录（如有）
    ├── image-1.png
    ├── image-2.png
    └── table-1.html         # 复杂表格
```

## 错误处理

1. **文件不存在** → 提示用户检查文件路径
2. **文件格式不支持** → 告知支持的文件类型
3. **OCR 失败** → 建议用户安装 tesseract 或提供文字型 PDF
4. **转换质量差** → 提供手动调整建议

## 批量转换

当用户需要批量转换时：
1. 确认目标目录
2. 遍历所有匹配文件
3. 逐个转换并报告进度
4. 生成转换报告（成功/失败数量）
