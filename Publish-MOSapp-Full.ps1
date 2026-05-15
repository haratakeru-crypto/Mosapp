# Publish-MOSapp-Full.ps1
# 他PCでも表紙から各科目exe・JSON問題文・採点が動作する発行物を作る一括スクリプト。
# 手順: (1) 3科目を Release ビルド (2) 表紙 Publish + VSTO 同梱 (3) 3科目の bin\Release を発行フォルダにマージ

$ErrorActionPreference = "Stop"

$scriptDir = if ($PSCommandPath) { (Get-Item (Split-Path -Parent $PSCommandPath)).FullName } elseif ($PSScriptRoot) { (Get-Item $PSScriptRoot).FullName } else { (Get-Location).Path }

$configuration = "Release"
$publishDir   = Join-Path $scriptDir "MOSapp\MOSapp\publish"

# 3科目の csproj（Release ビルド用）
$appProjects = @(
    (Join-Path $scriptDir "MOSapp\mos_xaml_app\MOSExcelMogiApp.csproj"),
    (Join-Path $scriptDir "MOSapp\MOS Word app\MOS Word app.csproj"),
    (Join-Path $scriptDir "MOSapp\Mos PowerPoint Mogi App\MOS PowerPoint app.csproj")
)

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

# [1/4] 3科目を Release ビルド
Write-Host "[1/4] Building 3 subject apps (Release)..." -ForegroundColor Cyan
foreach ($proj in $appProjects) {
    $name = [System.IO.Path]::GetFileNameWithoutExtension($proj)
    if (-not (Test-Path $proj)) {
        Write-Host "  Skip (not found): $name" -ForegroundColor Gray
        continue
    }
    Write-Host "  Building $name..." -ForegroundColor Yellow
    & $msbuild $proj /t:Build /p:Configuration=$configuration /v:minimal
    if ($LASTEXITCODE -ne 0) { throw "Failed to build $name." }
}

# [2/4] 表紙 Publish + VSTO 同梱
Write-Host "[2/4] Publishing table of contents + VSTO..." -ForegroundColor Cyan
& (Join-Path $scriptDir "Publish-MOSapp-With-VSTO.ps1")
if ($LASTEXITCODE -ne 0) { throw "Failed to run Publish-MOSapp-With-VSTO.ps1." }

# [3/4] 発行先の存在確認（表紙が MOSapp\MOSapp\publish に出力される想定）
if (-not (Test-Path $publishDir)) {
    $fallback = Join-Path $scriptDir "MOSapp\MOSapp\MOSapp\publish"
    if (Test-Path $fallback) { $publishDir = $fallback } else { throw "Publish folder not found: $publishDir" }
}

# [4/4] 3科目の bin\Release を発行フォルダにマージ
Write-Host "[4/4] Merging 3 apps into publish..." -ForegroundColor Cyan
& (Join-Path $scriptDir "Merge-AppOutputs-ToPublish.ps1") -PublishDir $publishDir
if ($LASTEXITCODE -ne 0) { throw "Failed to run Merge-AppOutputs-ToPublish.ps1." }

Write-Host "=== Full publish done ===" -ForegroundColor Green
Write-Host "Publish folder: $publishDir"
Write-Host "Other PCs: copy this folder (or Application Files\MOSapp_1_0_0_*) so that MOSapp.exe and the 3 subject exes + JSON/DLLs are in the same folder."
