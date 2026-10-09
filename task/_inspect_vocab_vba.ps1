$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$bin = Join-Path $env:TEMP 'mos_excel_vocab_unzip\extracted\xl\vbaProject.bin'
$bytes = [System.IO.File]::ReadAllBytes($bin)
Write-Host ("VBA_SIZE={0}" -f $bytes.Length)

# Extract printable ASCII and UTF-16LE strings (length >= 4)
function Get-AsciiStrings([byte[]]$b, [int]$minLen=4) {
    $sb = New-Object System.Text.StringBuilder
    $list = New-Object System.Collections.Generic.List[string]
    for ($i=0; $i -lt $b.Length; $i++) {
        $c = $b[$i]
        if ($c -ge 32 -and $c -le 126) {
            [void]$sb.Append([char]$c)
        } else {
            if ($sb.Length -ge $minLen) { $list.Add($sb.ToString()) }
            [void]$sb.Clear()
        }
    }
    if ($sb.Length -ge $minLen) { $list.Add($sb.ToString()) }
    return $list
}

function Get-Utf16Strings([byte[]]$b, [int]$minLen=4) {
    $list = New-Object System.Collections.Generic.List[string]
    $sb = New-Object System.Text.StringBuilder
    for ($i=0; $i -lt $b.Length - 1; $i += 2) {
        $lo = $b[$i]; $hi = $b[$i+1]
        if ($hi -eq 0 -and $lo -ge 32 -and $lo -le 126) {
            [void]$sb.Append([char]$lo)
        } elseif ($hi -eq 0 -and ($lo -ge 0xA0 -or ($lo -ge 0x80))) {
            # skip high ascii in utf16 for simplicity
            if ($sb.Length -ge $minLen) { $list.Add($sb.ToString()) }
            [void]$sb.Clear()
        } else {
            if ($sb.Length -ge $minLen) { $list.Add($sb.ToString()) }
            [void]$sb.Clear()
        }
    }
    if ($sb.Length -ge $minLen) { $list.Add($sb.ToString()) }
    return $list
}

$ascii = Get-AsciiStrings $bytes 4
$utf16 = Get-Utf16Strings $bytes 4

Write-Host "`n=== ASCII strings (interesting) ==="
$ascii | Where-Object {
    $_ -match 'Sub |Function |Module|Sheet|ThisWorkbook|Attribute|Option|Dim |If |Then|Else|MsgBox|Range|Cells|Value|Button|Tab|正解|不正解|答え|判定|Random|Shuffle|Click|CommandButton|UserForm|Private|Public|End Sub|End Function|Worksheet|ActiveSheet|Target|Answer|Check|Score|点数|関数|タブ'
} | Select-Object -Unique | ForEach-Object { Write-Host $_ }

Write-Host "`n=== ALL ASCII strings len>=6 (dedup) ==="
$ascii | Where-Object { $_.Length -ge 6 } | Select-Object -Unique | ForEach-Object { Write-Host $_ }

Write-Host "`n=== UTF16 strings len>=4 ==="
$utf16 | Where-Object { $_.Length -ge 4 } | Select-Object -Unique | ForEach-Object { Write-Host $_ }

# Also dump raw as escaped for module names near dir stream markers
Write-Host "`n=== Near 'Module' / 'VBA' markers ==="
$text = [System.Text.Encoding]::GetEncoding(1252).GetString($bytes)
$patterns = @('Module','ThisWorkbook','Sheet','Attribute VB_','Sub ','Function ','UserForm','CommandButton')
foreach ($p in $patterns) {
    $idx = 0
    $count = 0
    while (($idx = $text.IndexOf($p, $idx)) -ge 0 -and $count -lt 20) {
        $start = [Math]::Max(0, $idx - 20)
        $len = [Math]::Min(120, $text.Length - $start)
        $snip = $text.Substring($start, $len) -replace '[\x00-\x1F]', '.'
        Write-Host ("[{0} @{1}] {2}" -f $p, $idx, $snip)
        $idx += $p.Length
        $count++
    }
}
