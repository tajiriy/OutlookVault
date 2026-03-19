#!/usr/bin/env bash
# OutlookVault ビルドスクリプト
# 使い方:
#   ./build.sh          # Debug ビルド（デフォルト）
#   ./build.sh Release  # Release ビルド

set -euo pipefail

# vswhere で最新の Visual Studio から MSBuild.exe を検出
VSWHERE="C:\\Program Files (x86)\\Microsoft Visual Studio\\Installer\\vswhere.exe"
VS_PATH=$(powershell.exe -NoProfile -Command "& '$VSWHERE' -latest -requires Microsoft.Component.MSBuild -property installationPath" | tr -d '\r')
if [ -z "$VS_PATH" ]; then
    echo "エラー: MSBuild がインストールされた Visual Studio が見つかりません" >&2
    exit 1
fi
MSBUILD="${VS_PATH}\\MSBuild\\Current\\Bin\\MSBuild.exe"
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd -W | sed 's|/|\\\\|g')"
PROJECT="${SCRIPT_DIR}\\OutlookVault\\OutlookVault.vbproj"
TEST_PROJECT="${SCRIPT_DIR}\\OutlookVault.Tests\\OutlookVault.Tests.vbproj"
CONFIG="${1:-Debug}"

# === バージョン PATCH +1 ===
ASSEMBLY_INFO="${SCRIPT_DIR}\\OutlookVault\\My Project\\AssemblyInfo.vb"
ASSEMBLY_INFO_UNIX="$(dirname "$0")/OutlookVault/My Project/AssemblyInfo.vb"

if [ -f "$ASSEMBLY_INFO_UNIX" ]; then
    # 現在のバージョンを読み取り PATCH を +1
    CURRENT=$(grep 'AssemblyVersion(' "$ASSEMBLY_INFO_UNIX" | head -1 | sed 's/.*"\([0-9]*\.[0-9]*\.[0-9]*\.[0-9]*\)".*/\1/')
    if [ -n "$CURRENT" ]; then
        MAJOR=$(echo "$CURRENT" | cut -d. -f1)
        MINOR=$(echo "$CURRENT" | cut -d. -f2)
        PATCH=$(echo "$CURRENT" | cut -d. -f3)
        NEW_PATCH=$((PATCH + 1))
        NEW_VER="${MAJOR}.${MINOR}.${NEW_PATCH}.0"
        sed -i "s/AssemblyVersion(\"${CURRENT}\")/AssemblyVersion(\"${NEW_VER}\")/" "$ASSEMBLY_INFO_UNIX"
        sed -i "s/AssemblyFileVersion(\"${CURRENT}\")/AssemblyFileVersion(\"${NEW_VER}\")/" "$ASSEMBLY_INFO_UNIX"
        echo "=== バージョン更新: ${CURRENT} → ${NEW_VER} ==="
    fi
fi

echo ""
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
