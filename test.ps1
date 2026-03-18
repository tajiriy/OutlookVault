# OutlookArchiver テスト実行スクリプト
# 使い方:
#   .\test.ps1          # 全テスト実行（Debug）
#   .\test.ps1 Release  # Release ビルドのテスト実行

param(
    [string]$Config = "Debug"
)

$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$TestDll = Join-Path $ProjectRoot "OutlookArchiver.Tests\bin\$Config\OutlookArchiver.Tests.dll"

# vswhere で最新の Visual Studio から vstest.console.exe を検出
$vswhere = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
if (-not (Test-Path $vswhere)) {
    throw "vswhere.exe が見つかりません: $vswhere"
}
$vsPath = & $vswhere -latest -requires Microsoft.VisualStudio.PackageGroup.TestTools.Core -property installationPath
if (-not $vsPath) {
    throw "テストツールがインストールされた Visual Studio が見つかりません"
}
$vstest = Join-Path $vsPath "Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe"
if (-not (Test-Path $vstest)) {
    throw "vstest.console.exe が見つかりません: $vstest"
}

if (-not (Test-Path $TestDll)) {
    throw "テスト DLL が見つかりません: $TestDll`n先に .\build.ps1 を実行してください"
}

Write-Host "=== テスト実行 ==="
& $vstest $TestDll
if ($LASTEXITCODE -ne 0) { throw "テストに失敗しました" }
