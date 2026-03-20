# OutlookVault デプロイスクリプト
# 使い方:
#   .\deploy.ps1                              # デフォルト: C:\Tools\OutlookVault にデプロイ
#   .\deploy.ps1 -DeployDir "D:\MyApp\OA"     # デプロイ先を指定
#   .\deploy.ps1 -ExcludeConfig               # config を除外（既存設定を保持）

param(
    [string]$DeployDir = "C:\Tools\OutlookVault",
    [switch]$ExcludeConfig
)

$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$BuildDir = Join-Path $ProjectRoot "OutlookVault\bin\Release"
# vswhere で最新の Visual Studio から MSBuild.exe を検出
$vswhere = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
if (-not (Test-Path $vswhere)) {
    throw "vswhere.exe が見つかりません: $vswhere"
}
$vsPath = & $vswhere -latest -requires Microsoft.Component.MSBuild -property installationPath
if (-not $vsPath) {
    throw "MSBuild がインストールされた Visual Studio が見つかりません"
}
$MSBuild = Join-Path $vsPath "MSBuild\Current\Bin\MSBuild.exe"
if (-not (Test-Path $MSBuild)) {
    throw "MSBuild.exe が見つかりません: $MSBuild"
}
$Project = Join-Path $ProjectRoot "OutlookVault\OutlookVault.vbproj"

Write-Host "=== OutlookVault デプロイ ==="
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
$exePath = Join-Path $BuildDir "OutlookVault.exe"
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
    "OutlookVault.exe"
    "OutlookVault.exe.config"
    "EntityFramework.dll"
    "EntityFramework.SqlServer.dll"
    "Microsoft.Office.Interop.Outlook.dll"
    "System.Data.SQLite.dll"
    "System.Data.SQLite.EF6.dll"
    "System.Data.SQLite.Linq.dll"
)

foreach ($file in $files) {
    if ($ExcludeConfig -and $file -eq "OutlookVault.exe.config") {
        $destConfig = Join-Path $DeployDir $file
        if (Test-Path $destConfig) {
            Write-Host "  スキップ: $file（-ExcludeConfig: 既存設定を保持）" -ForegroundColor Cyan
        } else {
            Write-Host "  警告: $file をスキップしましたが、デプロイ先に既存ファイルがありません" -ForegroundColor Yellow
        }
        continue
    }
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

# ユーザーマニュアル (Help フォルダ丸ごとコピー)
$helpSrcDir = Join-Path $BuildDir "Help"
if (Test-Path $helpSrcDir) {
    $helpDestDir = Join-Path $DeployDir "Help"
    Copy-Item $helpSrcDir -Destination $helpDestDir -Recurse -Force
    Write-Host "  コピー: Help\ (user-manual.html, style.css, screenshots)"
} else {
    Write-Host "  警告: Help フォルダが見つかりません（スキップ）" -ForegroundColor Yellow
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
Write-Host "実行: $DeployDir\OutlookVault.exe"
