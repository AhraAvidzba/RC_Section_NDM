# ДЛЯ ТЕСТОВ: проверяет назначение полей ввода и заливки в собственной копии книги.
$ErrorActionPreference='Stop';$root=(Get-Location).Path
. (Join-Path $root 'tools/build_workbook/SettingsCatalog.ps1')
$directory=Join-Path $root 'outputs/CircleOpeningStyle_2026-11-28';$source=Join-Path $root 'workbook/candidates/CircleOpening/RC_Section_NDM.xlsm'
$before=Join-Path $directory 'Before.xlsm';$native=Join-Path $directory 'NativeFields.xlsm';if(Test-Path -LiteralPath $before){throw 'Primary style source already recorded.'}
Copy-Item -LiteralPath $source -Destination $before;Copy-Item -LiteralPath $source -Destination $native
$m=Get-CimInstance Win32_PerfFormattedData_PerfOS_Memory;if($m.AvailableMBytes-lt4096-or($m.CommitLimit-$m.CommittedBytes)-lt7GB){throw 'Insufficient memory reserve'}
$m|Select-Object AvailableMBytes,CommitLimit,CommittedBytes|ConvertTo-Json|Set-Content (Join-Path $directory 'Memory.json') -Encoding UTF8
Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public static class InputStyleTestWindow{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);}'
$app=$null;$book=$null;$owned=0;$ticks=0;$mask=@{Config=@();'Сечения'=@()};$fails=@();$count=0
try{
 $previous=@(Get-Process EXCEL -ErrorAction SilentlyContinue|ForEach-Object Id)
 $app=New-Object -ComObject Excel.Application;$app.Visible=$false;$app.EnableEvents=$false;$app.DisplayAlerts=$false;$app.AutomationSecurity=3
 [uint32]$id=0;[void][InputStyleTestWindow]::GetWindowThreadProcessId([IntPtr]$app.Hwnd,[ref]$id);if($previous-contains[int]$id){$app=$null;throw 'Own Excel required'}
 $owned=[int]$id;$ticks=(Get-Process -Id $owned).StartTime.ToUniversalTime().Ticks
 @{PID=$owned;StartTicksUTC=[string]$ticks}|ConvertTo-Json|Set-Content (Join-Path $directory 'Ownership.json') -Encoding UTF8
 $book=$app.Workbooks.Open($native,0,$false);if($book.ReadOnly){throw 'Own copy must be writable'}
 Apply-ConfigUserInputAlignment $book
 foreach($item in @(Get-ConfigUserInputTableSpecs)){
  $range=$book.Names.Item($item.Name).RefersToRange;$data=$range.Value2
  for($r=2;$r-le$range.Rows.Count;$r++){
   $key=[string]$data[$r,1];if($key.StartsWith('[')-or($item.Name-eq'rngSystemSettings'-and-not$key.Contains('.'))){continue}
   foreach($c in $item.Columns){$cell=$range.Cells.Item($r,$c);if($cell.MergeCells){$cell=$cell.MergeArea.Cells.Item(1,1)};$address=$cell.Address($false,$false);$mask['Config']+=,$address;$count++;if($cell.Interior.Color-ne15135487){$fails+=,'Config!'+$address}}
  }
 }
 foreach($def in @(Get-GeometrySettingsCatalog)){
  $range=$book.Names.Item($def.RangeName).RefersToRange
  foreach($block in @(Get-GeometryTableBlocks $range)){
   for($r=$block.FirstRow;$r-le$block.LastRow;$r++){
    $key=[string]$range.Cells.Item($r,1).Value2;if(-not$key-or$key-eq'Units'){continue}
    foreach($column in $block.Columns.Keys){
     if($column-in@('key','name','unit','note')){continue}
     if($block.Key-eq'Main'-and$key-like'Opening*'-and$column-in@('n','endMode')){continue}
     if($block.Key-eq'Main'-and$key-eq'OpeningCircle'-and$column-notin@('as','d')){continue}
     $cell=$range.Cells.Item($r,[int]$block.Columns[$column]);$address=$cell.Address($false,$false);$mask['Сечения']+=,$address;$count++;if($cell.Interior.Color-ne15135487){$fails+=,'Сечения!'+$address}
    }
   }
  }
 }
 foreach($name in @($mask.Keys)){$mask[$name]=@($mask[$name]|Sort-Object -Unique)}
 $mask|ConvertTo-Json -Depth 4|Set-Content (Join-Path $directory 'InputFillMask.json') -Encoding UTF8
 @{Status=$(if($fails.Count){'FAIL'}else{'PASS'});Checks=$count;Failed=$fails;SourceSHA256=(Get-FileHash -LiteralPath $source).Hash}|ConvertTo-Json -Depth 4|Set-Content (Join-Path $directory 'NativeFieldCheck.json') -Encoding UTF8
 if($fails.Count){throw 'Native field fill failed'}
 $book.Save();$book.Close($false);$book=$null;Write-Output ('INPUT STYLE PASS: '+$count+' fields')
}finally{
 if($book){$book.Close($false)}
 if($app){$app.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
 if($owned){$p=Get-Process -Id $owned -ErrorAction SilentlyContinue;if($p-and$p.StartTime.ToUniversalTime().Ticks-eq$ticks-and$p.ProcessName-eq'EXCEL'){Stop-Process -Id $owned -Force}}
}
