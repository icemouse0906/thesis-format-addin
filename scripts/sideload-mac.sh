#!/usr/bin/env bash
set -euo pipefail

ROOT="$(cd "$(dirname "$0")/.." && pwd)"
WEF_DIR="${HOME}/Library/Containers/com.microsoft.Word/Data/Documents/wef"

mkdir -p "$WEF_DIR"
cp "$ROOT/manifest.xml" "$WEF_DIR/thesis-format-manifest.xml"
echo "已复制 manifest → $WEF_DIR/thesis-format-manifest.xml"
echo "请完全退出 Word 后重新打开，在「开始」选项卡查找「格式助手」。"
echo "同时确保已运行: npm run dev"
