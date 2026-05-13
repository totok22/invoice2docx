#!/bin/bash
# Build standalone app for macOS
# Usage: ./build_mac.sh

set -e

echo "=== 发票 Word 生成器 - macOS 打包 ==="

# Ensure we're in the right directory
cd "$(dirname "$0")"

PYTHON_BIN="${PYTHON_BIN:-python3}"
if [ -x ".venv/bin/python" ]; then
    PYTHON_BIN=".venv/bin/python"
fi
FLET_BIN="${FLET_BIN:-flet}"
if [ -x ".venv/bin/flet" ]; then
    FLET_BIN=".venv/bin/flet"
fi

# Install dependencies if needed
"$PYTHON_BIN" -m pip install -r requirements.txt

# Package with flet
echo "Packaging with flet..."
"$FLET_BIN" pack main.py \
    --yes \
    --name "发票Word生成器" \
    --add-data "engine.py:." \
    --add-data "默认报账说明模板.docx:." \
    --add-data "默认验收单模板.docx:." \
    --product-name "发票Word生成器" \
    --product-version "0.1.0" \
    --copyright "BITFSAE" \
    --bundle-id "org.bitfsae.invoice2docx"

APP_PLIST="dist/发票Word生成器.app/Contents/Info.plist"
if [ -f "$APP_PLIST" ]; then
    /usr/libexec/PlistBuddy -c "Set :CFBundleShortVersionString 0.1.0" "$APP_PLIST" 2>/dev/null || \
        /usr/libexec/PlistBuddy -c "Add :CFBundleShortVersionString string 0.1.0" "$APP_PLIST"
    /usr/libexec/PlistBuddy -c "Set :CFBundleVersion 0.1.0" "$APP_PLIST" 2>/dev/null || \
        /usr/libexec/PlistBuddy -c "Add :CFBundleVersion string 0.1.0" "$APP_PLIST"
    codesign --force --deep --sign - "dist/发票Word生成器.app"
fi

echo ""
echo "=== 打包完成 ==="
echo "输出位置: dist/"
