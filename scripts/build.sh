#!/usr/bin/env bash
# =============================================================================
# build.sh - 从 docx 直接生成 pdf（保留原 Word 排版）
#
# 用途：当你修改了 文档版/晨读背诵内容vX.Y.docx 后，重新生成对应的 pdf
#
# 用法：
#   bash scripts/build.sh                              # 自动找最新的 docx
#   bash scripts/build.sh 文档版/晨读背诵内容v0.5.docx # 指定输入
#   bash scripts/build.sh 文档版/晨读背诵内容v0.5.docx my-note  # 自定义输出名
# =============================================================================

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
ROOT_DIR="$(dirname "$SCRIPT_DIR")"
DOCX_DIR="$ROOT_DIR/文档版"
PDF_DIR="$ROOT_DIR/PDF版"

# ---------- 颜色（必须在前面定义，因为 case 会调用 warn）----------
RED='\033[0;31m'; GREEN='\033[0;32m'; YELLOW='\033[1;33m'; CYAN='\033[0;36m'; NC='\033[0m'
info()  { echo -e "${CYAN}→${NC} $*"; }
ok()    { echo -e "${GREEN}✓${NC} $*"; }
warn()  { echo -e "${YELLOW}⚠${NC} $*"; }
fail()  { echo -e "${RED}✗${NC} $*"; exit 1; }

# ---------- 参数解析 ----------
if [ $# -ge 1 ]; then
    INPUT_DOCX="$1"
else
    # 默认取 文档版/ 下最新的 docx
    INPUT_DOCX="$(ls -t "$DOCX_DIR"/*.docx 2>/dev/null | head -1 || true)"
    [ -z "$INPUT_DOCX" ] && fail "文档版/ 下没有 docx 文件"
fi

[ -f "$INPUT_DOCX" ] || fail "输入文件不存在: $INPUT_DOCX"

if [ $# -ge 2 ]; then
    OUTPUT_NAME="$2"
else
    # 默认沿用原文件名（去掉扩展名）
    OUTPUT_NAME="$(basename "$INPUT_DOCX" .docx)"
fi

# 如果输出名包含非 ASCII 字符（中文等），自动转英文以避免 GitHub raw 404
# 原因：raw.githubusercontent.com 对中文路径下的文件返回 404（GitHub 已知 bug）
case "$OUTPUT_NAME" in
    *[![:ascii:]]*)
        version=$(echo "$OUTPUT_NAME" | grep -oE 'v[0-9]+\.[0-9]+' || echo "")
        if [ -n "$version" ]; then
            OUTPUT_NAME="cheatsheet-${version}"
        else
            OUTPUT_NAME="cheatsheet"
        fi
        warn "检测到中文文件名，自动改为英文: $OUTPUT_NAME"
        ;;
esac

PDF_FILENAME="${OUTPUT_NAME}.pdf"

# ---------- 前置检查 ----------
command -v soffice >/dev/null || command -v libreoffice >/dev/null \
    || fail "LibreOffice 未安装。运行：brew install --cask libreoffice / apt install libreoffice"

LO_CMD="$(command -v soffice || command -v libreoffice)"

# ---------- 执行 ----------
info "输入: $INPUT_DOCX"
info "输出: $PDF_DIR/$PDF_FILENAME"

(cd "$PDF_DIR" && "$LO_CMD" --headless --convert-to pdf --outdir . "$INPUT_DOCX" >/dev/null)

# LibreOffice 输出文件名沿用输入的 basename，如果跟目标不一致则改名
GENERATED="$PDF_DIR/$(basename "$INPUT_DOCX" .docx).pdf"
if [ "$GENERATED" != "$PDF_DIR/$PDF_FILENAME" ] && [ -f "$GENERATED" ]; then
    mv "$GENERATED" "$PDF_DIR/$PDF_FILENAME"
fi

if [ -f "$PDF_DIR/$PDF_FILENAME" ]; then
    SIZE=$(du -h "$PDF_DIR/$PDF_FILENAME" | cut -f1)
    ok "pdf 已生成 ($SIZE): PDF版/$PDF_FILENAME"
else
    fail "pdf 生成失败"
fi
echo ""