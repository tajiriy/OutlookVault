#!/usr/bin/env bash
# OutlookArchiver テスト実行スクリプト
# 使い方:
#   ./test.sh          # 全テスト実行
#   ./test.sh Release  # Release ビルドのテスト実行

set -euo pipefail

VSTEST="C:\\Program Files\\Microsoft Visual Studio\\18\\Community\\Common7\\IDE\\CommonExtensions\\Microsoft\\TestWindow\\vstest.console.exe"
TEST_DLL="d:\\Development\\VisualStudioProjects\\OutlookArchiver\\OutlookArchiver.Tests\\bin\\${1:-Debug}\\OutlookArchiver.Tests.dll"

echo "=== テスト実行 ==="
powershell.exe -NoProfile -Command "& '$VSTEST' '$TEST_DLL'"
