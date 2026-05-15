# Publish-MOSapp-With-VSTO.ps1
# 表紙（MOSapp）を発行し、VSTO アドインをビルドして publish\WordAddIn\ に同梱する

$ErrorActionPreference = "Stop"

# リポジトリルート（このスクリプトの置き場所・絶対パス）
$scriptDir = if ($PSCommandPath) { (Get-Item (Split-Path -Parent $PSCommandPath)).FullName } elseif ($PSScriptRoot) { (Get-Item $PSScriptRoot).FullName } else { (Get-Location).Path }

# 表紙プロジェクトと発行先（表紙 = MOSapp\MOSapp\ 内の MOSapp.csproj、発行は同フォルダの publish）
$tocCsproj   = Join-Path $scriptDir "MOSapp\MOSapp\MOSapp.csproj"
$publishDir  = Join-Path $scriptDir "MOSapp\MOSapp\publish"
if (-not (Test-Path $tocCsproj)) { throw "Table of contents project not found: $tocCsproj (scriptDir=$scriptDir)" }
$wordAddInDir = Join-Path $publishDir "WordAddIn"

# VSTO アドイン
$vstoProjPath = Join-Path $scriptDir "MOSapp\MOS Word app\New_MOSWordVSTOAddIn\New_MOSWordVSTOAddIn\New_MOSWordVSTOAddIn.csproj"
$configuration = "Release"
$vstoBinDir    = Join-Path $scriptDir "MOSapp\MOS Word app\New_MOSWordVSTOAddIn\New_MOSWordVSTOAddIn\bin\$configuration"

Write-Host "=== Publish MOSapp (表紙) + VSTO 同梱 ===" -ForegroundColor Cyan

# MSBuild を探す（VS 2022 / 18 など）
$msbuild = $null
foreach ($base in @(
    "${env:ProgramFiles}\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe",
    "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe",
    "${env:ProgramFiles}\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe"
)) {
    if ($base -and (Test-Path $base)) { $msbuild = $base; break }
}
if (-not $msbuild) { $msbuild = "msbuild.exe" }
Write-Host "MSBuild: $msbuild" -ForegroundColor Gray

# [1/3] 表紙を発行
Write-Host "[1/3] Publishing MOSapp (表紙)..." -ForegroundColor Yellow
& $msbuild $tocCsproj /t:Publish /p:Configuration=$configuration /p:PublishDir=$publishDir /v:minimal
if ($LASTEXITCODE -ne 0) { throw "Failed to publish MOSapp." }

# [2/3] VSTO アドインをビルド
Write-Host "[2/3] Building VSTO add-in..." -ForegroundColor Yellow
& $msbuild $vstoProjPath /t:Build /p:Configuration=$configuration /v:minimal
if ($LASTEXITCODE -ne 0) { throw "Failed to build VSTO add-in." }

if (-not (Test-Path $vstoBinDir)) {
    throw "VSTO build output not found: $vstoBinDir"
}

# [3/3] VSTO 出力を publish\WordAddIn にコピー
Write-Host "[3/3] Copying VSTO to publish\WordAddIn..." -ForegroundColor Yellow
New-Item -ItemType Directory -Path $wordAddInDir -Force | Out-Null

Get-ChildItem -Path $vstoBinDir -Filter "New_MOSWordVSTOAddIn.*" | ForEach-Object {
    Copy-Item $_.FullName -Destination $wordAddInDir -Force
    Write-Host "  Copied $($_.Name)" -ForegroundColor Gray
}
$vstoDeps = @("Microsoft.Office.Tools.Common.v4.0.Utilities.dll")
foreach ($name in $vstoDeps) {
    $src = Join-Path $vstoBinDir $name
    if (Test-Path $src) {
        Copy-Item $src -Destination $wordAddInDir -Force
        Write-Host "  Copied $name" -ForegroundColor Gray
    }
}

Write-Host "=== Done ===" -ForegroundColor Green
Write-Host "Publish folder: $publishDir"
Write-Host "VSTO folder:   $wordAddInDir"
