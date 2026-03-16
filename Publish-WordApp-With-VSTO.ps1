# Publish-WordApp-With-VSTO.ps1
# MOS Word app を公開し、そのあと VSTO をビルドして publish\WordAddIn\ にコピーする

$ErrorActionPreference = "Stop"

# パス類（環境に合わせて調整）
$solutionRoot  = "C:\Users\kouza\source\repos\MOSapp"
$wordProjPath  = Join-Path $solutionRoot "MOSapp\MOS Word app\MOS Word app.csproj"
$vstoProjPath  = Join-Path $solutionRoot "MOSapp\MOS Word app\New_MOSWordVSTOAddIn\New_MOSWordVSTOAddIn\New_MOSWordVSTOAddIn.csproj"
$configuration = "Release"
$publishDir    = Join-Path $solutionRoot "MOSapp\MOS Word app\publish\"
$vstoBinDir    = Join-Path $solutionRoot "MOSapp\MOS Word app\New_MOSWordVSTOAddIn\New_MOSWordVSTOAddIn\bin\$configuration"
$wordAddInDir  = Join-Path $publishDir "WordAddIn"
$checkerSrcDir = Join-Path $solutionRoot "MOSapp\MOS Word app\bin\$configuration\Dlls"
$checkerDstDir = Join-Path $publishDir "Dlls"

Write-Host "=== Publish MOS Word app + VSTO ===" -ForegroundColor Cyan

# MSBuild の場所を探す（VS2022 Community 前提、無ければ msbuild.exe 任せ）
$msbuild = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe"
if (-not (Test-Path $msbuild)) {
    $msbuild = "${env:ProgramFiles}\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe"
}
if (-not (Test-Path $msbuild)) {
    $msbuild = "msbuild.exe"
}

Write-Host "[1/3] Publishing MOS Word app..." -ForegroundColor Yellow
& $msbuild $wordProjPath /t:Publish /p:Configuration=$configuration /p:PublishDir=$publishDir
if ($LASTEXITCODE -ne 0) { throw "Failed to publish MOS Word app." }

Write-Host "[2/3] Building VSTO add-in..." -ForegroundColor Yellow
& $msbuild $vstoProjPath /t:Build /p:Configuration=$configuration
if ($LASTEXITCODE -ne 0) { throw "Failed to build VSTO project." }

if (-not (Test-Path $vstoBinDir)) {
    throw "VSTO build output folder not found: $vstoBinDir"
}

Write-Host "[3/3] Copying VSTO outputs to publish\WordAddIn..." -ForegroundColor Yellow
New-Item -ItemType Directory -Path $wordAddInDir -Force | Out-Null

Get-ChildItem -Path $vstoBinDir -Filter "New_MOSWordVSTOAddIn.*" | ForEach-Object {
    Copy-Item $_.FullName -Destination $wordAddInDir -Force
}
# VSTO 依存 DLL（.vsto が参照するもの）も WordAddIn にコピー
$vstoDeps = @("Microsoft.Office.Tools.Common.v4.0.Utilities.dll")
foreach ($name in $vstoDeps) {
    $src = Join-Path $vstoBinDir $name
    if (Test-Path $src) {
        Copy-Item $src -Destination $wordAddInDir -Force
    }
}

# チェッカー DLL (WordChecker1_n.dll) を publish\Dlls にコピー
if (Test-Path $checkerSrcDir) {
    Write-Host "Copying checker DLLs to publish\Dlls..." -ForegroundColor Yellow
    New-Item -ItemType Directory -Path $checkerDstDir -Force | Out-Null
    Get-ChildItem -Path $checkerSrcDir -Filter "*.dll" | ForEach-Object {
        Copy-Item $_.FullName -Destination $checkerDstDir -Force
    }
    # Application Files 内のバージョン付きフォルダにも Dlls をコピー（ClickOnce インストール後に exe の横で見つかるようにする）
    $appFilesDir = Join-Path $publishDir "Application Files"
    if (Test-Path $appFilesDir) {
        $versionedFolders = Get-ChildItem -Path $appFilesDir -Directory | Where-Object { $_.Name -like "MOS Word app_*" }
        foreach ($vf in $versionedFolders) {
            $dllsInVersioned = Join-Path $vf.FullName "Dlls"
            New-Item -ItemType Directory -Path $dllsInVersioned -Force | Out-Null
            Get-ChildItem -Path $checkerSrcDir -Filter "*.dll" | ForEach-Object {
                Copy-Item $_.FullName -Destination $dllsInVersioned -Force
            }
            Write-Host "Copying checker DLLs to Application Files\$($vf.Name)\Dlls..." -ForegroundColor Yellow
        }
    }
} else {
    Write-Host "Checker DLL source folder not found: $checkerSrcDir" -ForegroundColor Yellow
}

Write-Host "=== Done ===" -ForegroundColor Green
Write-Host "Publish folder: $publishDir"
Write-Host "VSTO folder: $wordAddInDir"

