#!/usr/bin/env bash
# OutlookArchiver ビルドスクリプト
# 使い方:
#   ./build.sh          # Debug ビルド（デフォルト）
#   ./build.sh Release  # Release ビルド

set -euo pipefail

MSBUILD="C:\\Program Files\\Microsoft Visual Studio\\18\\Community\\MSBuild\\Current\\Bin\\MSBuild.exe"
PROJECT="d:\\Development\\VisualStudioProjects\\OutlookArchiver\\OutlookArchiver\\OutlookArchiver.vbproj"
TEST_PROJECT="d:\\Development\\VisualStudioProjects\\OutlookArchiver\\OutlookArchiver.Tests\\OutlookArchiver.Tests.vbproj"
CONFIG="${1:-Debug}"

echo "=== NuGet パッケージ復元 ==="
powershell.exe -NoProfile -Command "& '$MSBUILD' '$PROJECT' /t:Restore /p:Configuration=$CONFIG"

echo ""
echo "=== ビルド (${CONFIG}) ==="
powershell.exe -NoProfile -Command "& '$MSBUILD' '$PROJECT' /p:Configuration=$CONFIG"

echo ""
echo "=== テストプロジェクト: NuGet パッケージ復元 ==="
powershell.exe -NoProfile -Command "& '$MSBUILD' '$TEST_PROJECT' /t:Restore /p:Configuration=$CONFIG"

echo ""
echo "=== テストプロジェクト: ビルド (${CONFIG}) ==="
powershell.exe -NoProfile -Command "& '$MSBUILD' '$TEST_PROJECT' /p:Configuration=$CONFIG"
