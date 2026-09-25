$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$files = Get-ChildItem 'c:\Users\kouza\Desktop' -File -Filter '*.xlsm'
foreach ($f in $files) {
    Write-Host ("FOUND: [{0}] len={1}" -f $f.Name, $f.Length)
}
$src = $files[0].FullName
Write-Host "SRC=$src"

$tmp = Join-Path $env:TEMP 'mos_excel_vocab_unzip'
if (Test-Path $tmp) { Remove-Item $tmp -Recurse -Force }
$extracted = Join-Path $tmp 'extracted'
New-Item -ItemType Directory -Path $extracted -Force | Out-Null

Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.IO.Compression.ZipFile]::ExtractToDirectory($src, $extracted)

Write-Host '=== ZIP TREE ==='
Get-ChildItem -Path $extracted -Recurse | ForEach-Object {
    $rel = $_.FullName.Substring($extracted.Length + 1)
    if ($_.PSIsContainer) { Write-Host "[D] $rel" } else { Write-Host ("[F] {0} ({1})" -f $rel, $_.Length) }
}

Write-Host '=== workbook.xml ==='
Get-Content -LiteralPath (Join-Path $extracted 'xl\workbook.xml') -Raw -Encoding UTF8

Write-Host '=== workbook.xml.rels ==='
Get-Content -LiteralPath (Join-Path $extracted 'xl\_rels\workbook.xml.rels') -Raw -Encoding UTF8
