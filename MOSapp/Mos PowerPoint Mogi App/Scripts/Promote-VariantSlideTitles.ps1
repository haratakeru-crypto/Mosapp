#Requires -Version 5.1
<#
.SYNOPSIS
  類題 PPTX の「タイトル」テキストを（非表示の）タイトルプレースホルダーへ複製し、
  目的別スライドショー等でスライド名が表示されるようにする。

.DESCRIPTION
  類題テンプレではタイトルPHが Visible=false で、見た目は「タイトル」テキストボックス側にある。
  見た目を変えず、PH に同じ文言を入れるだけにする（TB は削除しない）。
#>
[CmdletBinding()]
param(
    [string]$Root = "C:\MOSTest\PowerPoint365\Tab1\PracticeVariant1",
    [switch]$IncludeWorkingCopies,
    [switch]$WhatIf
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Get-ShapeText([object]$sh) {
    try {
        if ([int]$sh.HasTextFrame -ne -1) { return "" }
        $t = $sh.TextFrame.TextRange.Text
        if ($null -eq $t) { return "" }
        return ([string]$t).Trim()
    }
    catch { return "" }
}

function Find-PlaceholderByTypes([object]$slide, [int[]]$types) {
    for ($j = 1; $j -le $slide.Shapes.Count; $j++) {
        $sh = $slide.Shapes.Item($j)
        if ([int]$sh.Type -ne 14) { continue }
        try {
            $pt = [int]$sh.PlaceholderFormat.Type
            if ($types -contains $pt) { return $sh }
        }
        catch { }
    }
    return $null
}

function Find-NamedShape([object]$slide, [string]$name) {
    for ($j = 1; $j -le $slide.Shapes.Count; $j++) {
        $sh = $slide.Shapes.Item($j)
        if ($sh.Name -eq $name) { return $sh }
    }
    return $null
}

function Set-PlaceholderTextKeepHidden([object]$ph, [string]$text) {
    if ($null -eq $ph) { return $false }
    if ([string]::IsNullOrWhiteSpace($text)) { return $false }
    # 表示状態は触らない（元が非表示なら非表示のまま）
    $ph.TextFrame.TextRange.Text = $text
    return $true
}

function Promote-Presentation([string]$path) {
    Write-Host "---- $path"
    $pres = $script:PptApp.Presentations.Open($path, $false, $false, $false)
    $promotedTitle = 0
    $promotedSub = 0
    $already = 0
    $noPh = 0
    $noSrc = 0
    try {
        for ($i = 1; $i -le $pres.Slides.Count; $i++) {
            $slide = $pres.Slides.Item($i)
            $titlePh = Find-PlaceholderByTypes $slide @(1, 3)
            $titleTb = Find-NamedShape $slide "タイトル"
            $tbText = ""
            if ($null -ne $titleTb) { $tbText = Get-ShapeText $titleTb }

            if ($null -eq $titlePh) {
                Write-Host ("  Slide {0}: no title PH (layout={1})" -f $i, $slide.CustomLayout.Name)
                $noPh++
            }
            elseif ([string]::IsNullOrWhiteSpace($tbText)) {
                $phNow = Get-ShapeText $titlePh
                if ([string]::IsNullOrWhiteSpace($phNow)) {
                    Write-Host ("  Slide {0}: no タイトル source text" -f $i)
                    $noSrc++
                }
                else {
                    $already++
                }
            }
            else {
                $phNow = Get-ShapeText $titlePh
                if ($phNow -eq $tbText) {
                    $already++
                }
                elseif ($WhatIf) {
                    Write-Host ("  Slide {0}: WOULD set title PH from '{1}'" -f $i, $tbText)
                    $promotedTitle++
                }
                else {
                    [void](Set-PlaceholderTextKeepHidden $titlePh $tbText)
                    Write-Host ("  Slide {0}: title PH from '{1}'" -f $i, $tbText)
                    $promotedTitle++
                }
            }

            $subPh = Find-PlaceholderByTypes $slide @(4)
            $subTb = Find-NamedShape $slide "副題"
            if ($null -ne $subPh -and $null -ne $subTb) {
                $subTbText = Get-ShapeText $subTb
                $subPhText = Get-ShapeText $subPh
                if (-not [string]::IsNullOrWhiteSpace($subTbText) -and $subPhText -ne $subTbText) {
                    if ($WhatIf) {
                        Write-Host ("  Slide {0}: WOULD set subtitle PH from '{1}'" -f $i, $subTbText)
                    }
                    else {
                        [void](Set-PlaceholderTextKeepHidden $subPh $subTbText)
                        Write-Host ("  Slide {0}: subtitle PH from '{1}'" -f $i, $subTbText)
                    }
                    $promotedSub++
                }
            }
        }

        if (-not $WhatIf) {
            $pres.Save()
        }
    }
    finally {
        $pres.Close()
    }

    return [pscustomobject]@{
        File          = [IO.Path]::GetFileName($path)
        PromotedTitle = $promotedTitle
        PromotedSub   = $promotedSub
        AlreadyOk     = $already
        NoTitlePh     = $noPh
        NoSource      = $noSrc
    }
}

if (-not (Test-Path -LiteralPath $Root)) {
    throw "Root not found: $Root"
}

$targets = New-Object System.Collections.Generic.List[string]
$templates = Join-Path $Root "Templates"
if (Test-Path -LiteralPath $templates) {
    Get-ChildItem -LiteralPath $templates -Filter "similar*.pptx" | Sort-Object Name | ForEach-Object { $targets.Add($_.FullName) }
}
if ($IncludeWorkingCopies) {
    Get-ChildItem -LiteralPath $Root -Filter "similar*.pptx" -File | Sort-Object Name | ForEach-Object { $targets.Add($_.FullName) }
}

if ($targets.Count -eq 0) {
    throw "No similar*.pptx found under $Root"
}

Write-Host ("Files: {0}" -f $targets.Count)
$script:PptApp = New-Object -ComObject PowerPoint.Application
try { $script:PptApp.Visible = 1 } catch { }

$results = @()
try {
    foreach ($path in $targets) {
        $results += Promote-Presentation $path
    }
}
finally {
    try { $script:PptApp.Quit() } catch { }
    if ($null -ne $script:PptApp) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($script:PptApp)
    }
    $script:PptApp = $null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

$results | Format-Table -AutoSize
Write-Host "Done."
