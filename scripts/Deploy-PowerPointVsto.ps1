# Deploy-PowerPointVsto.ps1
# ビルド → 既存VSTOアンインストール → 新規インストール（PowerPointに反映）
$ErrorActionPreference = "Stop"
$root = Join-Path $PSScriptRoot "..\MOSapp\Mos PowerPoint Mogi App"
$vstoDir = Join-Path $root "PowerPointAddIn1\bin\Debug"
$vstoPath = Join-Path $vstoDir "PowerPointAddIn1.vsto"
$vstoinstaller = "${env:CommonProgramFiles}\Microsoft Shared\VSTO\10.0\VSTOInstaller.exe"

# 1) PowerPoint 終了（任意）
Get-Process -Name "POWERPNT" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1

# 2) ビルド（MSBuild は Developer PowerShell または VS の MSBuild を使用）
$csproj = Join-Path $root "PowerPointAddIn1\PowerPointAddIn1.csproj"
& msbuild $csproj /t:Build /p:Configuration=Debug /v:minimal
if ($LASTEXITCODE -ne 0) { throw "Build failed." }

# 3) アンインストール（既にインストール済みの場合）
if (Test-Path $vstoinstaller) {
    & $vstoinstaller /Uninstall $vstoPath /Silent 2>$null
    Start-Sleep -Seconds 1
}

# 4) インストール（/Silent で証明書未信頼の場合はダイアログが出る場合あり）
& $vstoinstaller /Install $vstoPath /Silent
Write-Host "Done. Start PowerPoint to load the add-in."
