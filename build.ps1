# OutlookArchiver ビルドスクリプト
# 使い方:
#   .\build.ps1          # Debug ビルド（デフォルト）
#   .\build.ps1 Release  # Release ビルド

param(
    [string]$Config = "Debug"
)

$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$Project = Join-Path $ProjectRoot "OutlookArchiver\OutlookArchiver.vbproj"
$TestProject = Join-Path $ProjectRoot "OutlookArchiver.Tests\OutlookArchiver.Tests.vbproj"

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

Write-Host "=== NuGet パッケージ復元 ==="
& $MSBuild $Project /t:Restore /p:Configuration=$Config
if ($LASTEXITCODE -ne 0) { throw "NuGet 復元に失敗しました" }

Write-Host ""
Write-Host "=== ビルド (${Config}) ==="
& $MSBuild $Project /p:Configuration=$Config
if ($LASTEXITCODE -ne 0) { throw "ビルドに失敗しました" }

Write-Host ""
Write-Host "=== テストプロジェクト: NuGet パッケージ復元 ==="
& $MSBuild $TestProject /t:Restore /p:Configuration=$Config
if ($LASTEXITCODE -ne 0) { throw "テストプロジェクトの NuGet 復元に失敗しました" }

Write-Host ""
Write-Host "=== テストプロジェクト: ビルド (${Config}) ==="
& $MSBuild $TestProject /p:Configuration=$Config
if ($LASTEXITCODE -ne 0) { throw "テストプロジェクトのビルドに失敗しました" }
