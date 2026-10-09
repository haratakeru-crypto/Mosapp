# Rebuild-And-Install-ExcelVSTO.ps1
# Uninstall -> rebuild -> install Excel VSTO add-in (ExcelAddIn1)

$ErrorActionPreference = "Stop"
$AddInName = "ExcelAddIn1"
$Configuration = if ($args[0]) { $args[0] } else { "Debug" }
$ScriptDir = $PSScriptRoot
if (-not $ScriptDir) { $ScriptDir = Get-Location }

function Write-Step([string]$Message) {
    Write-Host $Message -ForegroundColor Cyan
}

function Resolve-MsBuildPath {
    foreach ($candidate in @(
        "${env:ProgramFiles}\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe",
        "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe",
        "${env:ProgramFiles}\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe"
    )) {
        if ($candidate -and (Test-Path $candidate)) { return $candidate }
    }
    return "msbuild.exe"
}

function Resolve-VstoInstallerPath {
    $path = Join-Path ${env:CommonProgramFiles} "Microsoft Shared\VSTO\10.0\VSTOInstaller.exe"
    if (Test-Path $path) { return $path }
    return $null
}

function Resolve-VstoManifestPath {
    param([string]$RootDir, [string]$Config, [string]$Name)
    $candidates = @(
        (Join-Path $RootDir "ExcelAddIn1\bin\$Config\$Name.vsto"),
        (Join-Path $RootDir "ExcelAddIn1\bin\Debug\$Name.vsto"),
        (Join-Path $RootDir "ExcelAddIn1\bin\Release\$Name.vsto")
    )
    foreach ($p in $candidates) {
        if (Test-Path $p) { return (Resolve-Path $p).Path }
    }
    return $null
}

function Enable-ExcelAddInLoadBehavior {
    param([string]$Name)
    $keyPath = "HKCU:\Software\Microsoft\Office\Excel\Addins\$Name"
    if (-not (Test-Path $keyPath)) { return }
    Set-ItemProperty -Path $keyPath -Name "LoadBehavior" -Value 3 -Type DWord -ErrorAction SilentlyContinue
    foreach ($ver in @("16.0", "15.0")) {
        $resBase = "HKCU:\Software\Microsoft\Office\$ver\Excel\Resiliency"
        foreach ($sub in @("DisabledItems", "CrashingAddinList")) {
            $p = Join-Path $resBase $sub
            if (-not (Test-Path $p)) { continue }
            Get-ChildItem $p -ErrorAction SilentlyContinue | ForEach-Object {
                $props = Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue
                $nameVal = [string]$props.Name
                $desc = [string]$props.Description
                if ($nameVal -match [regex]::Escape($Name) -or $desc -match [regex]::Escape($Name)) {
                    Remove-Item $_.PSPath -Recurse -Force -ErrorAction SilentlyContinue
                    Write-Step "Cleared Excel resiliency entry: $sub\$($_.PSChildName)"
                }
            }
        }
    }
}

function Test-ExcelAddInRegistered {
    param([string]$Name)
    $keyPath = "HKCU:\Software\Microsoft\Office\Excel\Addins\$Name"
    if (-not (Test-Path $keyPath)) { return $false }
    $props = Get-ItemProperty $keyPath
    if ($null -eq $props.Manifest -or [string]::IsNullOrWhiteSpace($props.Manifest)) { return $false }
    $load = [int]$props.LoadBehavior
    return $load -ne 0
}

Write-Step "=== Excel VSTO rebuild and install ($Configuration) ==="

Write-Step "Stopping Excel..."
Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2

$vstoInstaller = Resolve-VstoInstallerPath
$installedManifest = $null
$regKey = "HKCU:\Software\Microsoft\Office\Excel\Addins\$AddInName"
if (Test-Path $regKey) {
    $installedManifest = (Get-ItemProperty $regKey -ErrorAction SilentlyContinue).Manifest
}

if ($vstoInstaller -and $installedManifest) {
    Write-Step "Uninstalling VSTO: $installedManifest"
    & $vstoInstaller /Uninstall $installedManifest /Silent
    Start-Sleep -Seconds 2
}

if ($vstoInstaller) {
    foreach ($cfg in @("Release", "Debug")) {
        $altVsto = Join-Path $ScriptDir "ExcelAddIn1\bin\$cfg\$AddInName.vsto"
        if (Test-Path $altVsto) {
            Write-Step "Uninstalling VSTO (if present): $altVsto"
            & $vstoInstaller /Uninstall $altVsto /Silent
        }
    }
    Start-Sleep -Seconds 2
}

$msbuild = Resolve-MsBuildPath
Write-Step "MSBuild: $msbuild"

$vstoProj = Join-Path $ScriptDir "ExcelAddIn1\ExcelAddIn1.csproj"
if (-not (Test-Path $vstoProj)) {
    throw "VSTO project not found: $vstoProj"
}

Write-Step "Building: $vstoProj ($Configuration)"
& $msbuild $vstoProj /t:Rebuild /p:Configuration=$Configuration /v:minimal
if ($LASTEXITCODE -ne 0) { throw "Build failed (exit $LASTEXITCODE)" }

$vstoPath = Resolve-VstoManifestPath -RootDir $ScriptDir -Config $Configuration -Name $AddInName
if (-not $vstoPath) {
    throw ".vsto not found. Check $Configuration build output."
}

Write-Step "Installing: $vstoPath"
if ($vstoInstaller) {
    & $vstoInstaller /Install $vstoPath /Silent
    if ($LASTEXITCODE -ne 0) {
        Write-Host "VSTOInstaller failed; launching .vsto directly..." -ForegroundColor Yellow
        Start-Process -FilePath $vstoPath -Wait
    }
}
else {
    Start-Process -FilePath $vstoPath -Wait
}

Start-Sleep -Seconds 2
Enable-ExcelAddInLoadBehavior -Name $AddInName
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\Excel\Addins\$AddInName" -Name "LoadBehavior" -Value 3 -Type DWord -ErrorAction SilentlyContinue
if (Test-ExcelAddInRegistered -Name $AddInName) {
    $props = Get-ItemProperty "HKCU:\Software\Microsoft\Office\Excel\Addins\$AddInName"
    Write-Step "=== Done: Excel add-in registered ==="
    Write-Host "Manifest: $($props.Manifest)"
    Write-Host "LoadBehavior: $($props.LoadBehavior)"
}
else {
    Write-Host "WARN: Excel add-in registry entry not found after install." -ForegroundColor Yellow
    Write-Host "Run manually: $vstoPath" -ForegroundColor Yellow
    exit 1
}
