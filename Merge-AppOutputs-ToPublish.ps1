# Merge-AppOutputs-ToPublish.ps1
# 発行フォルダ内のサブフォルダに Excel/Word/PowerPoint の bin\Release 一式をマージする。
# 表紙のみルートに表示し、3科目exeはサブフォルダに集約する。SubjectSelectionWindow.SubjectAppsSubfolderName と一致させること。
param(
    [string]$PublishDir
)

$ErrorActionPreference = "Stop"
$scriptDir = if ($PSCommandPath) { (Get-Item (Split-Path -Parent $PSCommandPath)).FullName } elseif ($PSScriptRoot) { (Get-Item $PSScriptRoot).FullName } else { (Get-Location).Path }

if (-not $PublishDir) {
    $PublishDir = Join-Path $scriptDir "MOSapp\MOSapp\publish"
    if (-not (Test-Path $PublishDir)) { $PublishDir = Join-Path $scriptDir "MOSapp\MOSapp\MOSapp\publish" }
}
if (-not (Test-Path $PublishDir)) { throw "Publish folder not found: $PublishDir" }

# 3科目exeを格納するサブフォルダ名。MOSapp Views\SubjectSelectionWindow.xaml.cs の SubjectAppsSubfolderName と一致させること。
$SubjectAppsSubfolderName = "App"

$configuration = "Release"
$appBins = @(
    (Join-Path $scriptDir "MOSapp\mos_xaml_app\bin\$configuration"),
    (Join-Path $scriptDir "MOSapp\MOS Word app\bin\$configuration"),
    (Join-Path $scriptDir "MOSapp\Mos PowerPoint Mogi App\bin\$configuration")
)

function Merge-Directory($src, $dest) {
    if (-not (Test-Path $src)) { return }
    Get-ChildItem -Path $src -Recurse -File | ForEach-Object {
        $rel = $_.FullName.Substring($src.Length).TrimStart('\')
        $target = Join-Path $dest $rel
        $targetDir = Split-Path -Parent $target
        if (-not (Test-Path $targetDir)) { New-Item -ItemType Directory -Path $targetDir -Force | Out-Null }
        Copy-Item $_.FullName -Destination $target -Force
    }
}

$destSub = Join-Path $PublishDir $SubjectAppsSubfolderName
if (-not (Test-Path $destSub)) { New-Item -ItemType Directory -Path $destSub -Force | Out-Null }

Write-Host "=== Merge 3 apps output into publish\$SubjectAppsSubfolderName ===" -ForegroundColor Cyan
Write-Host "PublishDir: $PublishDir" -ForegroundColor Gray
Write-Host "DestSub:    $destSub" -ForegroundColor Gray

foreach ($bin in $appBins) {
    $name = Split-Path (Split-Path $bin -Parent) -Leaf
    if (Test-Path $bin) {
        Write-Host "Merging $name\bin\Release -> publish\$SubjectAppsSubfolderName..." -ForegroundColor Yellow
        Merge-Directory -src $bin -dest $destSub
    } else {
        Write-Host "Skip (not found): $bin" -ForegroundColor Gray
    }
}

$appFiles = Join-Path $PublishDir "Application Files"
if (Test-Path $appFiles) {
    $versioned = Get-ChildItem -Path $appFiles -Directory | Where-Object { $_.Name -like "MOSapp_1_0_0_*" } | Sort-Object Name -Descending | Select-Object -First 1
    if ($versioned) {
        $versionedSub = Join-Path $versioned.FullName $SubjectAppsSubfolderName
        if (-not (Test-Path $versionedSub)) { New-Item -ItemType Directory -Path $versionedSub -Force | Out-Null }
        Write-Host "Merging into Application Files\$($versioned.Name)\$SubjectAppsSubfolderName..." -ForegroundColor Yellow
        foreach ($bin in $appBins) {
            $name = Split-Path (Split-Path $bin -Parent) -Leaf
            if (Test-Path $bin) { Merge-Directory -src $bin -dest $versionedSub }
        }
    }
}

Write-Host "=== Merge done ===" -ForegroundColor Green
