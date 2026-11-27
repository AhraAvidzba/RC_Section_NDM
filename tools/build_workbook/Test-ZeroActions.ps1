# ==========================================================================
# ДЛЯ ТЕСТОВ: ранний пропуск нулевых NM и самостоятельное T в Excel/File.
# ==========================================================================
param([Parameter(Mandatory=$true)][string]$WorkbookPath,[Parameter(Mandatory=$true)][string]$EvidenceDirectory,[switch]$SkipSuites)
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$source=(Resolve-Path -LiteralPath (Join-Path $root $WorkbookPath)).Path
$directory=Join-Path $root $EvidenceDirectory;[void][IO.Directory]::CreateDirectory($directory)
$path=Join-Path $directory 'Runtime.xlsm';Copy-Item -LiteralPath $source -Destination $path
$app=$null;$book=$null;$owned=0;$ticks=0;$checks=@()

# -------------------- ДЛЯ ТЕСТОВ --------------------
# Меняет значение по структурному ключу исключительно в проверочной копии.
function Test-ZeroSetting([string]$Key,[object]$Value,[string]$Table='rngSystemSettings'){
 $range=$book.Names.Item($Table).RefersToRange
 for($row=2;$row -le $range.Rows.Count;$row++){if([string]$range.Cells.Item($row,1).Value2 -ceq $Key){$range.Cells.Item($row,2).Value2=[string]$Value;return}}
 throw ('Missing setting '+$Key)
}

# -------------------- ДЛЯ ТЕСТОВ --------------------
# Сохраняет каждое утверждение, чтобы причина отказа оставалась доступной при остановке.
function Test-ZeroAssert([string]$Name,[bool]$Condition,[object]$Actual){
 $script:checks+= [ordered]@{Name=$Name;Passed=$Condition;Actual=$Actual}
 $script:checks | ConvertTo-Json -Depth 7 | Set-Content -Encoding UTF8 -LiteralPath (Join-Path $directory 'RuntimeChecks.json')
 if(-not $Condition){throw ('Failed: '+$Name+'; actual='+[string]$Actual)}
}

Add-Type 'using System;using System.Runtime.InteropServices;public static class ZeroActionsTestWindow{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);}'
try{
 $memory=Get-CimInstance Win32_PerfFormattedData_PerfOS_Memory
 if([double]$memory.CommitLimit-[double]$memory.CommittedBytes -lt 7GB){throw 'Insufficient commit headroom'}
 $app=New-Object -ComObject Excel.Application;$app.Visible=$false;$app.DisplayAlerts=$false;$app.EnableEvents=$false;$app.AutomationSecurity=1
 [uint32]$id=0;[void][ZeroActionsTestWindow]::GetWindowThreadProcessId([IntPtr]$app.Hwnd,[ref]$id);$owned=[int]$id;$ticks=(Get-Process -Id $owned).StartTime.ToUniversalTime().Ticks
 Write-Output ('OWNED_EXCEL: '+$owned+'; ticks='+$ticks)
 $book=$app.Workbooks.Open($path,0,$false);if($book.ReadOnly){throw 'Read-only fixture'}
 $prefix="'"+$book.Name+"'!"
 $suites=@('modTestBatchCalculation.RunZeroLoadCombinationTests','modTestBatchCalculation.RunAxialLoadZeroFilterTests','modTestBatchCalculation.RunLoadReferenceMomentFilterTests',
 'modTestTorsionPost.RunTorsionOnlyTests','modTestTorsionPost.RunTorsionPostTests','modTestTorsionPost.RunTorsionPostAvailabilityTests','modTestTorsionPost.RunTorsionPostUnitTests',
 'modTestTorsionPost.RunTorsionPostOutputTests','modTestTorsionPost.RunTorsionBarLengthTests','modTestTorsionPost.RunTorsionPostGuardTests','modTestTorsionPost.RunTorsionPostDiagnosticTests','modTestTorsionPost.RunTorsionGeometryInvariantTests')
 if(-not $SkipSuites){foreach($suite in $suites){
  $result=[string]$app.Run($prefix+$suite)
  $result | Set-Content -Encoding UTF8 -LiteralPath (Join-Path $directory ($suite.Split('.')[-1]+'.txt'))
  Test-ZeroAssert $suite ($result -match '(?i)failed[=:]\s*0\b' -and $result -notmatch '(?i)failed[=:]\s*[1-9]|RUNTIME ERROR|RUNTIME:') $result
  Write-Output ('PASS '+$suite)
 }}
 [void]$app.Run($prefix+'modTestPerformance.TestConfigurePerformanceFixture','DIRECT',1,'Newton','Circle','Normal','No')
 foreach($item in @(@('General.NonCriticalMessagesEnabled','No'),@('General.ExecutionReportEnabled','No'),@('Plot.Enabled','No'),@('General.ReformatResultsEnabled','No'),@('Results.SnapshotName','ZeroActions'),@('Calculation.ZeroAxialForcePerDepth','0.001'),@('Calculation.ZeroMomentPerDepth','0.001'),@('Load.ReferenceOffsetX','0'),@('Load.ReferenceOffsetY','0'))){Test-ZeroSetting $item[0] $item[1]}
 foreach($item in @(@('Force','tf'),@('Moment','tf*m'),@('Length','mm'))){Test-ZeroSetting $item[0] $item[1] 'rngUnitSettings'}
 foreach($item in @(@('Torsion.HoopAxisOffset',20),@('Torsion.HoopDiameter',8),@('Torsion.HoopSpacing',100),@('Torsion.BarBand',100),@('Torsion.StrutAngle',45),@('Torsion.Rs',350),@('Torsion.Rsw',300),@('Torsion.StateType','StrengthState'),@('Torsion.ConcreteMethod','None'))){Test-ZeroSetting $item[0] $item[1] 'rngTorsionSettings'}
 $circle=$book.Names.Item('rngCircleGeometry').RefersToRange
 for($row=1;$row -le $circle.Rows.Count;$row++){if([string]$circle.Cells.Item($row,1).Value2 -ceq 'Circle' -and [string]$circle.Cells.Item($row,2).Value2 -ceq 'По окружности'){$circle.Cells.Item($row,5).Value2='12';break}}
 $loads=$book.Names.Item('rngLoadCombinations').RefersToRange;$torque=$book.Names.Item('rngTorsionCombinations').RefersToRange
 $summary=$book.Names.Item('rngBatchSummary').RefersToRange.Cells.Item(1,1)
 $torsion=$book.Names.Item('rngTorsionSummaryAnchor').RefersToRange.Cells.Item(1,1)
 foreach($storage in @('Excel','File')){
  Test-ZeroSetting 'Results.SnapshotStorage' $storage
  [void]$app.Run($prefix+'modWorkbookCalculation.FormSectionForWorkbook',$book)
  [void]$loads.Offset(1,0).Resize($loads.Rows.Count-1,$loads.Columns.Count).ClearContents()
  [void]$torque.Offset(1,0).Resize($torque.Rows.Count-1,2).ClearContents()
  $ids=@('ZERO_N','ZERO_M','ONLY_T','FILTERED_T','KEEP')
  for($i=0;$i -lt $ids.Count;$i++){$loads.Cells.Item($i+2,1).Value2=$ids[$i];$loads.Cells.Item($i+2,5).Value2='PR1'}
  $loads.Cells.Item(7,1).Value2='NO_LOAD';$loads.Cells.Item(8,1).Value2='EXPLICIT_ZERO'
  $loads.Cells.Item(8,2).Resize(1,3).Value2=0.0
  $loads.Cells.Item(2,2).Value2=0.0001;$loads.Cells.Item(3,3).Value2=0.0001
  $loads.Cells.Item(5,2).Value2=0.0001;$loads.Cells.Item(5,3).Value2=0.0001;$loads.Cells.Item(6,2).Value2=-10.0
  $torque.Cells.Item(2,1).Value2='ONLY_T';$torque.Cells.Item(2,2).Value2=1.0
  $torque.Cells.Item(3,1).Value2='FILTERED_T';$torque.Cells.Item(3,2).Value2=1.0
  $message=[string]$app.Run($prefix+'modWorkbookCalculation.RunSectionCalculationForWorkbook',$book)
  Test-ZeroAssert ($storage+'.mixed.message') ($message -match '(?m)^Пропущено сочетаний, в которых N, Mx или My малы и обнулены фильтром, а T не задан: 2\. ID: ZERO_N, ZERO_M\r?$') $message
  Test-ZeroAssert ($storage+'.mixed.filteredIDs') ($message -notmatch '(?m)^ID:' -and $message -match '2\. ID: ZERO_N, ZERO_M') $message
  Test-ZeroAssert ($storage+'.mixed.emptyLoads.notReported') ($message -notmatch 'NO_LOAD|EXPLICIT_ZERO') $message
  Test-ZeroAssert ($storage+'.mixed.emptyLoads.notCalculated') ([string]$summary.Offset(17,0).Value2 -eq '' -and [string]$summary.Offset(18,0).Value2 -eq '') @($summary.Offset(17,0).Value2,$summary.Offset(18,0).Value2)
  Test-ZeroAssert ($storage+'.skipped.N') ([string]$summary.Offset(12,0).Value2 -eq '') $summary.Offset(12,0).Value2
  Test-ZeroAssert ($storage+'.skipped.M') ([string]$summary.Offset(13,0).Value2 -eq '') $summary.Offset(13,0).Value2
  Test-ZeroAssert ($storage+'.T.ID') ([string]$summary.Offset(14,0).Value2 -ceq 'ONLY_T') $summary.Offset(14,0).Value2
  Test-ZeroAssert ($storage+'.filteredT.ID') ([string]$summary.Offset(15,0).Value2 -ceq 'FILTERED_T') $summary.Offset(15,0).Value2
  Test-ZeroAssert ($storage+'.remaining.N') ([string]$summary.Offset(16,0).Value2 -ceq 'KEEP') $summary.Offset(16,0).Value2
  # DisplayColumns соответствует короткой шапке рабочего writer-а; сравнение физических сумм выполняется по двум независимым колонкам.
  $map=@(1,39,38,0,0,0,3,0,2,37,4,5,6,7,8,9,10,11,43,12,13,14,15,16,17,18,19,20,21,22,23,40,41,42,24,25,26,27,28,29,44,30,31,32,33,34,35,36)
  foreach($row in @(0,1)){
   $sigmaNM=$torsion.Offset($row,[Array]::IndexOf($map,30)).Value2
   $sigmaT=$torsion.Offset($row,[Array]::IndexOf($map,31)).Value2
   $sigmaSum=$torsion.Offset($row,[Array]::IndexOf($map,33)).Value2
   Test-ZeroAssert ($storage+'.T.NMzero.'+$row) ($null -ne $sigmaNM -and [double]$sigmaNM -eq 0) $sigmaNM
   Test-ZeroAssert ($storage+'.T.sum.'+$row) ($sigmaT -gt 0 -and [double]$sigmaSum -eq [double]$sigmaT) @($sigmaT,$sigmaSum)
  }
  $loads.Cells.Item(6,2).Value2=0.0001 # После фильтрации остаётся только T: повторный запуск очищает состояния KEEP.
  $message=[string]$app.Run($prefix+'modWorkbookCalculation.RunSectionCalculationForWorkbook',$book)
  Test-ZeroAssert ($storage+'.Tonly.message') ($message -match 'НДС от N, Mx и My не рассчитывался') $message
  $index=$book.Names.Item('rngNDMElementResultIndex').RefersToRange.Cells.Item(1,1)
  Test-ZeroAssert ($storage+'.Tonly.noNMstates') ([string]$index.Offset(2,0).Value2 -eq '') $index.Offset(2,0).Value2
  if($storage -eq 'File'){[void]$app.Run($prefix+'modCalculationSnapshot.UpdateSectionResultsForWorkbook',$book)}
  [void]$torque.Offset(1,0).Resize($torque.Rows.Count-1,2).ClearContents()
  $message=[string]$app.Run($prefix+'modWorkbookCalculation.RunSectionCalculationForWorkbook',$book)
  Test-ZeroAssert ($storage+'.allZero.message') ($message -match 'Результаты обновлены без расчётных сочетаний') $message
  Test-ZeroAssert ($storage+'.allZero.noOldID') ([string]$summary.Offset(14,0).Value2 -eq '' -and [string]$torsion.Value2 -eq '') @($summary.Offset(14,0).Value2,$torsion.Value2)
  $torque.Cells.Item(2,1).Value2='ONLY_T';$torque.Cells.Item(2,2).Value2='bad T'
  $message=[string]$app.Run($prefix+'modWorkbookCalculation.RunSectionCalculationForWorkbook',$book)
  Test-ZeroAssert ($storage+'.invalidT.reasonInMessage') ($message -match 'Момент T должен быть числом') $message
  Test-ZeroAssert ($storage+'.invalidT.reasonInResult') ([string]$torsion.Offset(0,1).Value2 -match 'Момент T должен быть числом') $torsion.Offset(0,1).Value2
  $torque.Cells.Item(2,1).Value2='UNKNOWN';$torque.Cells.Item(2,2).Value2=1.0
  $message=[string]$app.Run($prefix+'modWorkbookCalculation.RunSectionCalculationForWorkbook',$book)
  Test-ZeroAssert ($storage+'.orphanT.reasonInMessage') ($message -match 'ID не найден в rngLoadCombinations') $message
  $torque.Cells.Item(2,1).Value2='ONLY_T';$loads.Cells.Item(4,2).Value2='bad N'
  $message=[string]$app.Run($prefix+'modWorkbookCalculation.RunSectionCalculationForWorkbook',$book)
  Test-ZeroAssert ($storage+'.invalidNM.notAssumedZero') ([string]$torsion.Offset(0,[Array]::IndexOf($map,36)).Value2 -eq 'N/A') $torsion.Offset(0,[Array]::IndexOf($map,36)).Value2
  $loads.Cells.Item(4,2).Value2=0.0;[void]$torque.Offset(1,0).Resize($torque.Rows.Count-1,2).ClearContents()
  [void]$app.Run($prefix+'modWorkbookCalculation.RunSectionCalculationForWorkbook',$book)
 }
 $book.Save();$book.Close($false);$book=$null
 $book=$app.Workbooks.Open($path,0,$false)
 [void]$app.Run($prefix+'modCalculationSnapshot.UpdateSectionResultsForWorkbook',$book)
 Test-ZeroAssert 'File.reopen.restoredWithoutOldResults' ([string]$book.Names.Item('rngTorsionSummaryAnchor').RefersToRange.Value2 -eq '') 'Restored'
 [ordered]@{Status='PASS';Assertions=$checks.Count;PID=$owned;StartTicksUTC=$ticks} | ConvertTo-Json | Set-Content -Encoding UTF8 -LiteralPath (Join-Path $directory 'Summary.json')
 Write-Output 'ZERO_ACTIONS_PASS'
}catch{
 [ordered]@{Message=$_.Exception.Message;Stack=$_.ScriptStackTrace;Position=$_.InvocationInfo.PositionMessage} | ConvertTo-Json -Depth 5 | Set-Content -Encoding UTF8 -LiteralPath (Join-Path $directory 'Failure.json')
 if($book){$book.Save()}
 throw
}finally{
 if($book){$book.Close($false)};if($app){$app.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
 if($owned){$p=Get-Process -Id $owned -ErrorAction SilentlyContinue;if($p -and $p.ProcessName -eq 'EXCEL' -and $p.StartTime.ToUniversalTime().Ticks -eq $ticks){Stop-Process -Id $owned -Force -ErrorAction SilentlyContinue}}
}
