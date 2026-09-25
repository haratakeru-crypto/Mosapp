$ErrorActionPreference = "Continue"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$outDir = "c:\Users\kouza\source\repos\MOSapp\task\_vocab_extract\vba_export"
New-Item -ItemType Directory -Force -Path $outDir | Out-Null
$src = (Get-ChildItem "c:\Users\kouza\Desktop\*.xlsm" | Select-Object -First 1).FullName
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Open($src, $null, $true)
$log = New-Object System.Collections.Generic.List[string]
$log.Add("Opened $src")
try {
  $vb = $wb.VBProject
  foreach ($comp in $vb.VBComponents) {
    $name = $comp.Name
    $n = $comp.CodeModule.CountOfLines
    $code = ""
    if ($n -gt 0) { $code = $comp.CodeModule.Lines(1, $n) }
    [System.IO.File]::WriteAllText((Join-Path $outDir ($name + ".bas.txt")), $code, [System.Text.UTF8Encoding]::new($false))
    $log.Add("exported $name lines=$n")
  }
} catch {
  $log.Add("VBProject fail: $($_.Exception.Message)")
}
# dump formulas for key sheets
foreach ($sh in @($wb.Worksheets | ForEach-Object { $_ })) {
  $name = $sh.Name
  try {
    $used = $sh.UsedRange
    $addr = $used.Address($false,$false)
    $fml = $used.Formula
    $val = $used.Value2
    # serialize as TSV of formulas
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("ADDR=$addr")
    if ($fml -is [object[,]]) {
      $r1 = $fml.GetLowerBound(0); $r2 = $fml.GetUpperBound(0)
      $c1 = $fml.GetLowerBound(1); $c2 = $fml.GetUpperBound(1)
      for ($r=$r1; $r -le $r2; $r++) {
        $row = @()
        for ($c=$c1; $c -le $c2; $c++) {
          $cell = $fml[$r,$c]
          if ($null -eq $cell) { $row += "" } else { $row += [string]$cell }
        }
        [void]$sb.AppendLine(($row -join "`t"))
      }
    } else {
      [void]$sb.AppendLine([string]$fml)
    }
    $safe = ($name -replace '[\\/:*?\"<>| ]','_')
    [System.IO.File]::WriteAllText((Join-Path $outDir ("fml_$safe.tsv")), $sb.ToString(), [System.Text.UTF8Encoding]::new($false))
    $log.Add("fml dump $name")
  } catch {
    $log.Add("fml fail $name : $($_.Exception.Message)")
  }
}
$wb.Close($false)
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
[System.IO.File]::WriteAllText((Join-Path $outDir "_ps_log.txt"), ($log -join "`n"), [System.Text.UTF8Encoding]::new($false))
Write-Host "PS_DONE"