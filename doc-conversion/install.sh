#!/bin/bash
# doc-conversion skill 安装脚本
# 用法：./install.sh

set -e

SKILL_NAME="doc-conversion"
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SKILL_TARGET="$HOME/.claude/skills/$SKILL_NAME"

echo "========================================"
echo "  安装 $SKILL_NAME skill"
echo "========================================"
echo ""

# 创建 skills 目录（如果不存在）
mkdir -p "$HOME/.claude/skills"

# 检查是否已安装
if [ -L "$SKILL_TARGET" ] || [ -d "$SKILL_TARGET" ]; then
    echo "检测到已存在的安装，是否覆盖？(y/n)"
    read -r response
    if [[ "$response" =~ ^[Yy]$ ]]; then
        rm -rf "$SKILL_TARGET"
    else
        echo "安装已取消"
        exit 0
    fi
fi

# 创建软链接
ln -sf "$SCRIPT_DIR" "$SKILL_TARGET"

echo ""
echo "✓ 安装完成！"
echo ""
echo "使用方法:"
echo "  1. 在 Claude Code 中输入 /skills"
echo "  2. 选择 $SKILL_NAME"
echo ""
echo "可选：安装系统依赖以获得完整功能"
echo "  brew install pandoc poppler tesseract"
echo "  pip install python-docx pymupdf pdf2image pytesseract"
echo ""
