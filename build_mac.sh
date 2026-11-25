#!/bin/bash
# ============================================================
# PPT 工具箱 · macOS 打包脚本
#
# 用法（在 macOS 上，任意目录执行）：
#     bash build_mac.sh
#
# 产物：dist/PPT工具箱.app 以及 dist/PPT工具箱-mac.zip
# 依赖：Python 3.9+（建议 3.11 / 3.12）、Xcode 命令行工具（xcode-select --install）
# ============================================================
set -e

cd "$(dirname "$0")"

# 允许外部指定解释器：CI 里可能用 actions/setup-python 的 python3，
# 也可能因为缺 tkinter 退回到 /opt/homebrew/bin/python3.12
PYTHON_BIN="${PYTHON:-python3}"

say() { printf "\n\033[1;32m▸ %s\033[0m\n" "$1"; }
warn() { printf "\033[1;33m！ %s\033[0m\n" "$1"; }

# ---------- 0. 环境检查 ----------
if [ "$(uname)" != "Darwin" ]; then
    warn "这不是 macOS，PyInstaller 无法交叉编译。请把整个目录拷到 Mac 上再执行本脚本。"
    exit 1
fi

if ! command -v "$PYTHON_BIN" >/dev/null 2>&1; then
    warn "没找到 $PYTHON_BIN。请先安装 Python 3.11（https://www.python.org/downloads/macos/）"
    exit 1
fi

PYVER=$("$PYTHON_BIN" -c 'import sys;print("%d.%d"%sys.version_info[:2])')
say "使用 Python $PYVER（$PYTHON_BIN）"

if ! "$PYTHON_BIN" -c "import tkinter" >/dev/null 2>&1; then
    warn "当前 python3 没有 tkinter（系统自带 Python 与部分 Homebrew 版本没有）。"
    echo "   办法一：装 python.org 的官方安装包（推荐）"
    echo "   办法二：brew install python-tk"
    exit 1
fi

# ---------- 1. 虚拟环境 ----------
if [ ! -d ".venv-mac" ]; then
    say "创建虚拟环境 .venv-mac"
    "$PYTHON_BIN" -m venv .venv-mac
fi
# shellcheck disable=SC1091
source .venv-mac/bin/activate

say "安装依赖"
python -m pip install --upgrade pip -q
python -m pip install -r requirements-mac.txt -q

# ---------- 2. 图标 ----------
say "生成 app.icns"
if [ -f "logo_512.png" ] || [ -f "assets/logo.png" ]; then
    SRC="logo_512.png"
    [ -f "$SRC" ] || SRC="assets/logo.png"
    rm -rf icon.iconset app.icns
    mkdir -p icon.iconset
    for s in 16 32 64 128 256 512; do
        sips -z $s $s "$SRC" --out "icon.iconset/icon_${s}x${s}.png" >/dev/null 2>&1
    done
    # @2x 版本
    for s in 16 32 128 256 512; do
        d=$((s * 2))
        sips -z $d $d "$SRC" --out "icon.iconset/icon_${s}x${s}@2x.png" >/dev/null 2>&1
    done
    iconutil -c icns icon.iconset -o app.icns >/dev/null 2>&1 || warn "iconutil 失败，将使用默认图标"
    rm -rf icon.iconset
else
    warn "没找到 logo 图片，跳过图标"
fi

# ---------- 3. 打包 ----------
say "开始打包（首次约 2–4 分钟）"
rm -rf build dist
python -m PyInstaller --clean --noconfirm PPT工具箱_mac.spec

APP="dist/PPT工具箱.app"
if [ ! -d "$APP" ]; then
    warn "打包失败：没有生成 $APP"
    exit 1
fi

# ---------- 4. 签名（ad-hoc，Apple Silicon 上必须）----------
say "临时签名（ad-hoc）"
codesign --force --deep --sign - "$APP" || warn "签名失败，可能需要在「系统设置 → 隐私与安全性」里手动允许"

# ---------- 5. 压缩，便于拷贝 ----------
say "压缩为 zip"
ditto -c -k --sequesterRsrc --keepParent "$APP" "dist/PPT工具箱-mac.zip"

SIZE=$(du -sh "$APP" | awk '{print $1}')
say "完成：$APP（$SIZE）"
echo "   压缩包：dist/PPT工具箱-mac.zip"
echo
echo "首次打开如果提示「无法验证开发者」："
echo "   右键 App → 打开 → 打开；或执行  xattr -dr com.apple.quarantine \"$APP\""
