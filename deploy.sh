#!/usr/bin/env bash
# OutlookArchiver デプロイスクリプト
# 使い方:
#   ./deploy.sh                        # デフォルト: C:\Tools\OutlookArchiver にデプロイ
#   ./deploy.sh /c/MyApp/OutlookArchiver  # デプロイ先を指定

set -euo pipefail

DEPLOY_DIR="${1:-/c/Tools/OutlookArchiver}"
BUILD_DIR="d:/Development/VisualStudioProjects/OutlookArchiver/OutlookArchiver/bin/Release"

echo "=== OutlookArchiver デプロイ ==="
echo "デプロイ先: ${DEPLOY_DIR}"
echo ""

# Release ビルド
echo "=== Release ビルド ==="
./build.sh Release

# ビルド成果物の存在確認
if [ ! -f "${BUILD_DIR}/OutlookArchiver.exe" ]; then
    echo "エラー: Release ビルドの成果物が見つかりません: ${BUILD_DIR}/OutlookArchiver.exe"
    exit 1
fi

# デプロイ先ディレクトリの作成
mkdir -p "${DEPLOY_DIR}"

# 実行ファイルと依存DLLをコピー
echo ""
echo "=== ファイルコピー ==="
FILES=(
    "OutlookArchiver.exe"
    "OutlookArchiver.exe.config"
    "EntityFramework.dll"
    "EntityFramework.SqlServer.dll"
    "Microsoft.Office.Interop.Outlook.dll"
    "System.Data.SQLite.dll"
    "System.Data.SQLite.EF6.dll"
    "System.Data.SQLite.Linq.dll"
)

for file in "${FILES[@]}"; do
    if [ -f "${BUILD_DIR}/${file}" ]; then
        cp "${BUILD_DIR}/${file}" "${DEPLOY_DIR}/"
        echo "  コピー: ${file}"
    else
        echo "  警告: ${file} が見つかりません（スキップ）"
    fi
done

# SQLite ネイティブDLL (x64/x86)
for arch in x64 x86; do
    if [ -d "${BUILD_DIR}/${arch}" ]; then
        mkdir -p "${DEPLOY_DIR}/${arch}"
        cp "${BUILD_DIR}/${arch}/SQLite.Interop.dll" "${DEPLOY_DIR}/${arch}/"
        echo "  コピー: ${arch}/SQLite.Interop.dll"
    fi
done

# data フォルダの処理
if [ -d "${DEPLOY_DIR}/data" ]; then
    echo ""
    echo "※ デプロイ先に既存の data フォルダがあります。データは上書きしません。"
    echo "  データを移行する場合は手動でコピーしてください。"
else
    echo ""
    echo "※ デプロイ先に data フォルダがありません。"
    echo "  初回起動時に自動作成されます。"
    echo "  既存データを移行する場合は data フォルダをコピーしてください。"
fi

echo ""
echo "=== デプロイ完了 ==="
echo "実行: ${DEPLOY_DIR}/OutlookArchiver.exe"
