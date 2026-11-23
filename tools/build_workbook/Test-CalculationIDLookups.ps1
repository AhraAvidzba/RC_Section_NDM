# ==========================================================================
# ДЛЯ ТЕСТОВ: точный поиск ID в формулах оформления без решения НДС.
# Книга открывается только для чтения; подмены выполняются в памяти и не сохраняются.
# ==========================================================================
param([Parameter(Mandatory=$true)][string]$WorkbookPath,[Parameter(Mandatory=$true)][string]$EvidenceDirectory,[string]$ExistingID='6')
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$path=(Resolve-Path -LiteralPath (Join-Path $root $WorkbookPath)).Path
if($path -eq (Join-Path $root 'workbook/output/RC_Section_NDM.xlsm')){throw 'Use an isolated workbook'}
$directory=Join-Path $root $EvidenceDirectory;[void][IO.Directory]::CreateDirectory($directory)
$app=$null;$book=$null;$owned=0;$ticks=0;$checks=@()
Add-Type 'using System;using System.Runtime.InteropServices;public static class CalculationIDTestWindow{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);}'
try{
 $memory=Get-CimInstance Win32_PerfFormattedData_PerfOS_Memory
 if($memory.CommitLimit-$memory.CommittedBytes -lt 7GB){throw 'Insufficient memory'}
 $app=New-Object -ComObject Excel.Application;$app.Visible=$false;$app.DisplayAlerts=$false;$app.EnableEvents=$false;$app.AutomationSecurity=3
 [uint32]$id=0;[void][CalculationIDTestWindow]::GetWindowThreadProcessId([IntPtr]$app.Hwnd,[ref]$id);$owned=[int]$id;$ticks=(Get-Process -Id $owned).StartTime.ToUniversalTime().Ticks
 Write-Output ('OWNED_EXCEL: '+$owned+'; ticks='+$ticks)
 $book=$app.Workbooks.Open($path,0,$true);if(-not $book.ReadOnly){throw 'Read-only verification required'}
 $sheet=$book.Worksheets.Item('Расчет');$results=$book.Worksheets.Item('Results')
 $range=$sheet.UsedRange;$data=$range.Formula;$lookups=@();$sources=@{}
 $pattern='MATCH\((\$?[A-Z]+\$?\d+)(?:&"")?,(Results!\$?[A-Z]+\$?\d+:\$?[A-Z]+\$?\d+),0\)'
 for($r=1;$r -le $data.GetLength(0);$r++){
  for($c=1;$c -le $data.GetLength(1);$c++){
   $formula=[string]$data[$r,$c];if(-not $formula.StartsWith('=')){continue}
   $matches=[regex]::Matches($formula,$pattern)
   if(-not $matches.Count){continue}
   foreach($match in @($matches | Group-Object Value | ForEach-Object {$_.Group[0]})){
    $key=$match.Groups[1].Value;$source=$match.Groups[2].Value
    if(-not $sources.ContainsKey($source)){
     $ids=$results.Range($source.Substring(8));$found=$null
     for($row=1;$row -le $ids.Rows.Count;$row++){
      if([string]$ids.Cells.Item($row,1).Value2 -ceq $ExistingID){$found=$ids.Cells.Item($row,1);break}
     }
     if($null -eq $found){throw ('Fixture lacks ID in '+$source)}
     $sources[$source]=[ordered]@{Cell=$found;Text=($found.Value2 -is [string]);Index=$row}
    }
    $lookups+=[ordered]@{Address=$range.Cells.Item($r,$c).Address($false,$false);Formula=('='+$match.Value.Replace('MATCH('+$key,"MATCH('Расчет'!"+$key));Expected=$sources[$source].Index}
   }
  }
 }
 if(-not $lookups.Count){throw 'No exact ID lookups found'}
 $scratch=$book.Worksheets.Add();$expressions=New-Object 'object[,]' $lookups.Count,1
 for($i=0;$i -lt $lookups.Count;$i++){$expressions[$i,0]=$lookups[$i].Formula}
 $scratch.Range('A1').Resize($lookups.Count,1).Formula=$expressions
 foreach($scenario in @('Numeric','Text')){
  if($scenario -eq 'Numeric'){
   $results.Range('G12:R12').Value2=[double]$ExistingID
   foreach($source in $sources.Values){if($source.Text){$source.Cell.Value2=[string]$ExistingID}else{$source.Cell.Value2=[double]$ExistingID}}
  }else{
   $results.Range('G12:R12').Value2='CASE-6'
   foreach($source in $sources.Values){$source.Cell.Value2='CASE-6'}
  }
  $app.CalculateFull();$actual=$scratch.Range('A1').Resize($lookups.Count,1).Value2
  for($i=0;$i -lt $lookups.Count;$i++){
   $passed=($actual[($i+1),1] -eq $lookups[$i].Expected)
   $checks+=[ordered]@{Scenario=$scenario;Address=$lookups[$i].Address;Passed=$passed;Actual=$actual[($i+1),1];Expected=$lookups[$i].Expected}
  }
 }
 $results.Range('G12:R12').Value2='-';$app.CalculateFull()
 foreach($item in $lookups){
  $value=$sheet.Range($item.Address).Value2
  $checks+=[ordered]@{Scenario='NoSelection';Address=$item.Address;Passed=([string]$value -ceq '-');Actual=$value;Expected='-'}
 }
 $checks | ConvertTo-Json -Depth 6 | Set-Content -Encoding UTF8 -LiteralPath (Join-Path $directory 'LookupChecks.json')
 $failed=@($checks | Where-Object {-not $_.Passed})
 [ordered]@{Status=$(if($failed.Count){'FAIL'}else{'PASS'});LookupCells=$lookups.Count;Assertions=$checks.Count;Failed=$failed.Count;Scenarios=@('Numeric','Text','NoSelection')} | ConvertTo-Json -Depth 5 | Set-Content -Encoding UTF8 -LiteralPath (Join-Path $directory 'LookupSummary.json')
 if($failed.Count){throw ($failed | ConvertTo-Json -Depth 6)}
 Write-Output 'LOOKUPS_PASS'
}finally{
 if($book){$book.Close($false)};if($app){$app.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
 if($owned){$p=Get-Process -Id $owned -ErrorAction SilentlyContinue;if($p -and $p.ProcessName -eq 'EXCEL' -and $p.StartTime.ToUniversalTime().Ticks -eq $ticks){Stop-Process -Id $owned -Force -ErrorAction SilentlyContinue}}
}
