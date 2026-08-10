#Requires -Version 5.1
<#
.SYNOPSIS
  MOSTest データ配置用の Tab1 フォルダ構成を作成します（Office ファイルは含めません）。

.DESCRIPTION
  Excel / Word / PowerPoint の統一規約に合わせ、Templates → Tab1 直下 → Initial の
  3 か所へ実データを手動配置できるよう、空のディレクトリツリーを作成します。
  既存フォルダがあっても上書きせず、再実行しても安全です。

.PARAMETER RootPath
  データルート。既定は C:\MOSTest

.PARAMETER IncludePracticeVariants
  Excel Tab1 向けの PracticeVariant1〜5 フォルダ（各 Templates 含む）も作成します。

.EXAMPLE
  powershell -ExecutionPolicy Bypass -File .\scripts\Initialize-MOSTestFolders.ps1

.EXAMPLE
  powershell -ExecutionPolicy Bypass -File .\scripts\Initialize-MOSTestFolders.ps1 -IncludePracticeVariants
#>
[CmdletBinding()]
param(
    [string]$RootPath = 'C:\MOSTest',
    [switch]$IncludePracticeVariants
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Ensure-Directory {
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
        Write-Host "  [作成] $Path"
    }
    else {
        Write-Host "  [既存] $Path"
    }
}

function Get-WordExtension {
    param([int]$ProjectId)
    if ($ProjectId -eq 7) { return '.doc' }
    return '.docx'
}

Write-Host '=== MOSTest Tab1 フォルダ構成の初期化 ==='
Write-Host "RootPath: $RootPath"
Write-Host ''

$created = @()

$subjects = @(
    @{ Name = 'Excel365'; Extension = '.xlsx'; ProjectCount = 10 }
    @{ Name = 'Word365'; Extension = $null; ProjectCount = 10 }
    @{ Name = 'PowerPoint365'; Extension = '.pptx'; ProjectCount = 10 }
)

foreach ($subject in $subjects) {
    $subjectRoot = Join-Path $RootPath $subject.Name
    $templatesTab = Join-Path $subjectRoot 'Templates\Tab1'
    $tabWorking = Join-Path $subjectRoot 'Tab1'
    $tabInitial = Join-Path $subjectRoot 'Tab1\Initial'

    Write-Host "[$($subject.Name)]"
    Ensure-Directory -Path $templatesTab
    Ensure-Directory -Path $tabWorking
    Ensure-Directory -Path $tabInitial

    $created += $templatesTab
    $created += $tabWorking
    $created += $tabInitial
    Write-Host ''
}

if ($IncludePracticeVariants) {
    Write-Host '[Excel365 PracticeVariant]'
    for ($variant = 1; $variant -le 5; $variant++) {
        $variantRoot = Join-Path $RootPath "Excel365\Tab1\PracticeVariant$variant"
        $variantTemplates = Join-Path $variantRoot 'Templates'
        Ensure-Directory -Path $variantRoot
        Ensure-Directory -Path $variantTemplates
        $created += $variantRoot
        $created += $variantTemplates
    }
    Write-Host ''
}

Write-Host '=== 手動配置チェックリスト（Tab1 / Project1〜10） ==='
Write-Host ''
Write-Host '各科目で、同じ教材ファイルを次の 3 か所へ配置してください:'
Write-Host '  1. Templates\Tab1\Project{N}.{ext}  （リセット元・編集可）'
Write-Host '  2. Tab1\Project{N}.{ext}              （作業ファイル）'
Write-Host '  3. Tab1\Initial\Project{N}.{ext}      （採点用初期状態）'
Write-Host ''

for ($projectId = 1; $projectId -le 10; $projectId++) {
    $wordExt = Get-WordExtension -ProjectId $projectId
    Write-Host "Project$projectId :"
    Write-Host "  Excel      : $RootPath\Excel365\Templates\Tab1\Project$projectId.xlsx"
    Write-Host "               $RootPath\Excel365\Tab1\Project$projectId.xlsx"
    Write-Host "               $RootPath\Excel365\Tab1\Initial\Project$projectId.xlsx"
    Write-Host "  Word       : $RootPath\Word365\Templates\Tab1\Project$projectId$wordExt"
    Write-Host "               $RootPath\Word365\Tab1\Project$projectId$wordExt"
    Write-Host "               $RootPath\Word365\Tab1\Initial\Project$projectId$wordExt"
    Write-Host "  PowerPoint : $RootPath\PowerPoint365\Templates\Tab1\Project$projectId.pptx"
    Write-Host "               $RootPath\PowerPoint365\Tab1\Project$projectId.pptx"
    Write-Host "               $RootPath\PowerPoint365\Tab1\Initial\Project$projectId.pptx"
    Write-Host ''
}

Write-Host '=== 配置後の確認手順 ==='
Write-Host '1. 各アプリを起動し、Tab1 Project1 が開けることを確認'
Write-Host '2. リセット実行後、Templates / Tab1 直下 / Initial の内容が一致することを確認'
Write-Host ''
Write-Host '=== 保護ビュー対策（配置後に実行） ==='
Write-Host "Get-ChildItem '$RootPath\Word365' -Recurse -Include *.doc,*.docx | Unblock-File"
Write-Host "Get-ChildItem '$RootPath\PowerPoint365' -Recurse -Include *.pptx | Unblock-File"
Write-Host "Get-ChildItem '$RootPath\Excel365' -Recurse -Include *.xlsx | Unblock-File"
Write-Host ''
Write-Host '=== 完了 ==='
