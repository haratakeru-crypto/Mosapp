$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$extracted = Join-Path $env:TEMP 'mos_excel_vocab_unzip\extracted'

# --- shared strings ---
[xml]$ssXml = Get-Content -LiteralPath (Join-Path $extracted 'xl\sharedStrings.xml') -Encoding UTF8
$ns = New-Object System.Xml.XmlNamespaceManager($ssXml.NameTable)
$ns.AddNamespace('m','http://schemas.openxmlformats.org/spreadsheetml/2006/main')
$strings = @()
foreach ($si in $ssXml.SelectNodes('//m:si',$ns)) {
    $tNodes = $si.SelectNodes('.//m:t',$ns)
    $text = ($tNodes | ForEach-Object { $_.'#text' }) -join ''
    $strings += $text
}
Write-Host ("SHARED_STRINGS_COUNT={0}" -f $strings.Count)
$strings | ForEach-Object -Begin { $i=0 } -Process { Write-Host ("SS[{0}]={1}" -f $i, $_); $i++ }

function Get-ColIndex([string]$ref) {
    if ($ref -match '^([A-Z]+)') {
        $col = 0
        foreach ($ch in $Matches[1].ToCharArray()) { $col = $col * 26 + ([int][char]$ch - [int][char]'A' + 1) }
        return $col
    }
    return 0
}

function Dump-Sheet([string]$sheetFile, [string]$sheetName) {
    Write-Host ("`n======== SHEET: {0} ({1}) ========" -f $sheetName, $sheetFile)
    [xml]$xml = Get-Content -LiteralPath (Join-Path $extracted "xl\worksheets\$sheetFile") -Encoding UTF8
    $ns2 = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
    $ns2.AddNamespace('m','http://schemas.openxmlformats.org/spreadsheetml/2006/main')

    $rows = @{}
    foreach ($c in $xml.SelectNodes('//m:sheetData/m:row/m:c',$ns2)) {
        $ref = $c.GetAttribute('r')
        if ($ref -match '^([A-Z]+)(\d+)$') {
            $col = Get-ColIndex $Matches[1]
            $row = [int]$Matches[2]
            $t = $c.GetAttribute('t')
            $vNode = $c.SelectSingleNode('m:v',$ns2)
            $isNode = $c.SelectSingleNode('m:is',$ns2)
            $val = ''
            if ($t -eq 's' -and $vNode) {
                $idx = [int]$vNode.'#text'
                $val = $strings[$idx]
            } elseif ($t -eq 'inlineStr' -and $isNode) {
                $val = ($isNode.SelectNodes('.//m:t',$ns2) | ForEach-Object { $_.'#text' }) -join ''
            } elseif ($vNode) {
                $val = $vNode.'#text'
            }
            if (-not $rows.ContainsKey($row)) { $rows[$row] = @{} }
            $rows[$row][$col] = $val
        }
    }

    $maxCol = 0
    foreach ($r in $rows.Keys) { foreach ($c in $rows[$r].Keys) { if ($c -gt $maxCol) { $maxCol = $c } } }
    $sortedRows = $rows.Keys | Sort-Object
    Write-Host ("ROWS={0} MAXCOL={1}" -f $sortedRows.Count, $maxCol)
    foreach ($r in $sortedRows) {
        $cells = @()
        for ($c=1; $c -le $maxCol; $c++) {
            if ($rows[$r].ContainsKey($c)) { $cells += $rows[$r][$c] } else { $cells += '' }
        }
        Write-Host ("R{0}|{1}" -f $r, ($cells -join ' || '))
    }
}

$map = @{
    'sheet1.xml'='タブ単語帳'
    'sheet2.xml'='タブテスト'
    'sheet3.xml'='関数単語帳'
    'sheet4.xml'='関数テスト'
    'sheet5.xml'='関数テスト 1-8'
    'sheet6.xml'='関数'
    'sheet7.xml'='タブボタン対応'
}
foreach ($k in ($map.Keys | Sort-Object)) { Dump-Sheet $k $map[$k] }

# table1
Write-Host "`n=== table1.xml ==="
Get-Content -LiteralPath (Join-Path $extracted 'xl\tables\table1.xml') -Raw -Encoding UTF8

# ctrlProps / vml
Write-Host "`n=== ctrlProps ==="
Get-ChildItem (Join-Path $extracted 'xl\ctrlProps') | ForEach-Object {
    Write-Host $_.Name
    Get-Content $_.FullName -Raw -Encoding UTF8
}
Write-Host "`n=== vml drawings ==="
Get-ChildItem (Join-Path $extracted 'xl\drawings') -Filter '*.vml' | ForEach-Object {
    Write-Host $_.Name
    Get-Content $_.FullName -Raw -Encoding UTF8
}

# sheet rels for controls
Write-Host "`n=== sheet rels ==="
Get-ChildItem (Join-Path $extracted 'xl\worksheets\_rels') | ForEach-Object {
    Write-Host $_.Name
    Get-Content $_.FullName -Raw -Encoding UTF8
}
