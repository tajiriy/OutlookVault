# OutlookArchiver デプロイスクリプト
# 使い方:
#   .\deploy.ps1                              # デフォルト: C:\Tools\OutlookArchiver にデプロイ
#   .\deploy.ps1 -DeployDir "D:\MyApp\OA"     # デプロイ先を指定

param(
    [string]$DeployDir = "C:\Tools\OutlookArchiver"
)

$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$BuildDir = Join-Path $ProjectRoot "OutlookArchiver\bin\Release"
$MSBuild = "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe"
$Project = Join-Path $ProjectRoot "OutlookArchiver\OutlookArchiver.vbproj"

Write-Host "=== OutlookArchiver デプロイ ==="
Write-Host "デプロイ先: $DeployDir"
Write-Host ""

# Release ビルド
Write-Host "=== NuGet パッケージ復元 ==="
& $MSBuild $Project /t:Restore /p:Configuration=Release
if ($LASTEXITCODE -ne 0) { throw "NuGet 復元に失敗しました" }

Write-Host ""
Write-Host "=== Release ビルド ==="
& $MSBuild $Project /p:Configuration=Release
if ($LASTEXITCODE -ne 0) { throw "ビルドに失敗しました" }

# ビルド成果物の存在確認
$exePath = Join-Path $BuildDir "OutlookArchiver.exe"
if (-not (Test-Path $exePath)) {
    throw "エラー: Release ビルドの成果物が見つかりません: $exePath"
}

# デプロイ先ディレクトリの作成
if (-not (Test-Path $DeployDir)) {
    New-Item -ItemType Directory -Path $DeployDir -Force | Out-Null
}

# 実行ファイルと依存DLLをコピー
Write-Host ""
Write-Host "=== ファイルコピー ==="
$files = @(
    "OutlookArchiver.exe"
    "OutlookArchiver.exe.config"
    "EntityFramework.dll"
    "EntityFramework.SqlServer.dll"
    "Microsoft.Office.Interop.Outlook.dll"
    "System.Data.SQLite.dll"
    "System.Data.SQLite.EF6.dll"
    "System.Data.SQLite.Linq.dll"
)

foreach ($file in $files) {
    $src = Join-Path $BuildDir $file
    if (Test-Path $src) {
        Copy-Item $src -Destination $DeployDir -Force
        Write-Host "  コピー: $file"
    } else {
        Write-Host "  警告: $file が見つかりません（スキップ）" -ForegroundColor Yellow
    }
}

# SQLite ネイティブDLL (x64/x86)
foreach ($arch in @("x64", "x86")) {
    $srcDir = Join-Path $BuildDir $arch
    if (Test-Path $srcDir) {
        $destDir = Join-Path $DeployDir $arch
        if (-not (Test-Path $destDir)) {
            New-Item -ItemType Directory -Path $destDir -Force | Out-Null
        }
        Copy-Item (Join-Path $srcDir "SQLite.Interop.dll") -Destination $destDir -Force
        Write-Host "  コピー: $arch\SQLite.Interop.dll"
    }
}

# data フォルダの処理
$dataDir = Join-Path $DeployDir "data"
Write-Host ""
if (Test-Path $dataDir) {
    Write-Host "※ デプロイ先に既存の data フォルダがあります。データは上書きしません。"
    Write-Host "  データを移行する場合は手動でコピーしてください。"
} else {
    Write-Host "※ デプロイ先に data フォルダがありません。"
    Write-Host "  初回起動時に自動作成されます。"
    Write-Host "  既存データを移行する場合は data フォルダをコピーしてください。"
}

Write-Host ""
Write-Host "=== デプロイ完了 ==="
Write-Host "実行: $DeployDir\OutlookArchiver.exe"
