# Rebuild-And-Install-WordVSTO.ps1
# VSTOアドインをアンインストール → ビルド → インストール

$ErrorActionPreference = "Stop"
$AddInName = "New_MOSWordVSTOAddIn"
$ScriptDir = $PSScriptRoot
if (-not $ScriptDir) { $ScriptDir = Get-Location }

Write-Host "=== VSTO 再ビルド・インストール ===" -ForegroundColor Cyan

# 1) Word を終了
Write-Host "Word を終了しています..."
Get-Process -Name "WINWORD" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2

# 2) 現在の VSTO をアプリ一覧からアンインストール
$uninstallKey = Get-ChildItem -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*" -ErrorAction SilentlyContinue |
    Get-ItemProperty -ErrorAction SilentlyContinue |
    Where-Object { $_.DisplayName -like "*$AddInName*" -or $_.DisplayName -like "*MOS*Word*Add*In*" }
if ($uninstallKey -and $uninstallKey.UninstallString) {
    Write-Host "アンインストール実行: $($uninstallKey.DisplayName)"
    $cmd = $uninstallKey.UninstallString
    if ($cmd -match "rundll32") {
        Start-Process -FilePath "rundll32.exe" -ArgumentList "dfshim.dll,ShArpMaintain $AddInName.application" -Wait -NoNewWindow
    } else {
        Start-Process -FilePath "cmd.exe" -ArgumentList "/c `"$cmd`"" -Wait
    }
    Start-Sleep -Seconds 2
} else {
    Write-Host "既存のアドインのアンインストール情報が見つかりません（スキップ）"
}

# 3) VSTO をビルド
$vstoProj = Join-Path $ScriptDir "New_MOSWordVSTOAddIn\New_MOSWordVSTOAddIn\New_MOSWordVSTOAddIn.csproj"
$msbuild = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe"
if (-not (Test-Path $msbuild)) { $msbuild = "${env:ProgramFiles}\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" }
if (-not (Test-Path $msbuild)) { $msbuild = "msbuild.exe" }

if (Test-Path $vstoProj) {
    Write-Host "ビルド中: $vstoProj"
    & $msbuild $vstoProj /t:Rebuild /p:Configuration=Debug /v:minimal
    if ($LASTEXITCODE -ne 0) { throw "ビルドに失敗しました" }
} else {
    $vstoProj = Join-Path $ScriptDir "MOSWordVSTOAddIn\MOSWordVSTOAddIn.csproj"
    if (Test-Path $vstoProj) {
        Write-Host "ビルド中: $vstoProj"
        & $msbuild $vstoProj /t:Rebuild /p:Configuration=Debug /v:minimal
        if ($LASTEXITCODE -ne 0) { throw "ビルドに失敗しました" }
    } else { throw "VSTO プロジェクトが見つかりません" }
}

# 4) .vsto を探してインストール
$vstoPath = $null
foreach ($base in @($ScriptDir, (Join-Path $ScriptDir "bin\Debug"))) {
    $p = Join-Path $base "New_MOSWordVSTOAddIn\New_MOSWordVSTOAddIn\bin\Debug\$AddInName.vsto"
    if (Test-Path $p) { $vstoPath = $p; break }
}
if (-not $vstoPath) {
    $p = Join-Path $ScriptDir "MOSWordVSTOAddIn\bin\Debug\MOSWordVSTOAddIn.vsto"
    if (Test-Path $p) { $vstoPath = $p }
}
if ($vstoPath) {
    Write-Host "インストール実行: $vstoPath"
    Start-Process -FilePath $vstoPath -Wait
} else {
    Write-Host "警告: .vsto が見つかりません。手動でインストールしてください。"
}
Write-Host "=== 完了 ===" -ForegroundColor Green
