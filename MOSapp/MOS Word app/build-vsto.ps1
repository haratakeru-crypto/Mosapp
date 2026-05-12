# VSTOアドインのビルドスクリプト
param(
    [string]$Configuration = "Debug",
    [switch]$RestoreNuGet,
    [switch]$Clean
)

$ErrorActionPreference = "Stop"

Write-Host "=== VSTOアドイン ビルドスクリプト ===" -ForegroundColor Cyan

$solutionPath = "MOS Word app.sln"
$projectPath = "MOSWordVSTOAddIn\MOSWordVSTOAddIn.csproj"

# ソリューションの存在確認
if (-not (Test-Path $solutionPath)) {
    Write-Error "ソリューションが見つかりません: $solutionPath"
    exit 1
}

# NuGetパッケージの復元
if ($RestoreNuGet -or $Clean) {
    Write-Host "`nNuGetパッケージを復元しています..." -ForegroundColor Yellow
    nuget restore $solutionPath
    if ($LASTEXITCODE -ne 0) {
        Write-Warning "NuGetパッケージの復元に失敗しました。続行します..."
    }
}

# クリーンビルド
if ($Clean) {
    Write-Host "`nクリーンビルドを実行しています..." -ForegroundColor Yellow
    msbuild $solutionPath /t:Clean /p:Configuration=$Configuration
}

# MSBuildのパスを取得
$msbuildPaths = @(
    "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe",
    "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe",
    "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\MSBuild.exe",
    "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2019\Community\MSBuild\Current\Bin\MSBuild.exe",
    "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2019\Professional\MSBuild\Current\Bin\MSBuild.exe",
    "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2019\Enterprise\MSBuild\Current\Bin\MSBuild.exe"
)

$msbuild = $null
foreach ($path in $msbuildPaths) {
    if (Test-Path $path) {
        $msbuild = $path
        break
    }
}

if (-not $msbuild) {
    $msbuild = "msbuild.exe"
    Write-Host "MSBuildのパスを特定できませんでした。環境変数のMSBuildを使用します。" -ForegroundColor Yellow
}

# ビルド実行
Write-Host "`nビルドを実行しています..." -ForegroundColor Yellow
Write-Host "MSBuild: $msbuild" -ForegroundColor Gray
Write-Host "プロジェクト: $projectPath" -ForegroundColor Gray
Write-Host "構成: $Configuration" -ForegroundColor Gray

& $msbuild $projectPath /p:Configuration=$Configuration /p:Platform="Any CPU" /v:minimal /nologo

if ($LASTEXITCODE -ne 0) {
    Write-Error "ビルドに失敗しました。"
    exit 1
}

# ビルド出力の確認
Write-Host "`nビルド出力を確認しています..." -ForegroundColor Yellow
$outputDir = "MOSWordVSTOAddIn\bin\$Configuration"
$requiredFiles = @(
    "MOSWordVSTOAddIn.dll",
    "MOSWordVSTOAddIn.dll.manifest",
    "MOSWordVSTOAddIn.vsto"
)

$allFilesExist = $true
foreach ($file in $requiredFiles) {
    $filePath = Join-Path $outputDir $file
    if (Test-Path $filePath) {
        Write-Host "  ✓ $file" -ForegroundColor Green
    } else {
        Write-Host "  ✗ $file (見つかりません)" -ForegroundColor Red
        $allFilesExist = $false
    }
}

if ($allFilesExist) {
    Write-Host "`n✓ ビルドが正常に完了しました！" -ForegroundColor Green
    Write-Host "出力ディレクトリ: $outputDir" -ForegroundColor Cyan
} else {
    Write-Warning "`n一部のファイルが見つかりませんでした。VSTOが正しくインストールされているか確認してください。"
}

Write-Host "`n次のステップ:" -ForegroundColor Cyan
Write-Host "1. Visual Studioでプロジェクトを開いて、プロパティを確認してください" -ForegroundColor White
Write-Host "2. F5キーでデバッグを開始して、Wordでアドインが動作することを確認してください" -ForegroundColor White

