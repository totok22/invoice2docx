#!/bin/bash
# Build standalone app for macOS
# Usage: ./build_mac.sh

set -e

echo "=== 发票 Word 生成器 - macOS 打包 ==="

# Ensure we're in the right directory
cd "$(dirname "$0")"

# Install dependencies if needed
pip install -r requirements.txt 2>/dev/null || pip3 install -r requirements.txt

# Package with flet
echo "Packaging with flet..."
flet pack main.py \
    --name "发票Word生成器" \
    --icon assets/icon.png \
    --add-data "engine.py:." \
    --product-name "发票Word生成器" \
    --product-version "1.0.0" \
    --copyright "FSAE Team"

echo ""
echo "=== 打包完成 ==="
echo "输出位置: dist/"
echo ""
echo "如果没有 icon.png，可以去掉 --icon 参数重新运行。"
