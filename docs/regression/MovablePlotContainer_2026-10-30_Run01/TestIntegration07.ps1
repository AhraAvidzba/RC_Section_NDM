# -------------------- ДЛЯ ТЕСТОВ --------------------
# Запускает адресные проверки и сохранение/перенос реальной схемы только в собственной копии новой книги.
$ErrorActionPreference='Stop'
Import-Module 'C:/Windows/System32/WindowsPowerShell/v1.0/Modules/Microsoft.PowerShell.Utility/Microsoft.PowerShell.Utility.psd1' -Global
$p=$PSScriptRoot;$out=Join-Path $p 'Integration07';if(Test-Path -LiteralPath $out){throw 'Inspect existing directed attempt first'}
[void][IO.Directory]::CreateDirectory($out)
$source=Join-Path $p 'Build02/RC_Section_NDM.xlsm';$path=Join-Path $out 'TestWorkbook.xlsm';Copy-Item -LiteralPath $source -Destination $path
$m=Get-CimInstance Win32_PerfFormattedData_PerfOS_Memory
if($m.AvailableMBytes-lt8192-or($m.CommitLimit-$m.CommittedBytes)-lt14GB){throw 'Insufficient RAM/commit reserve'}
Add-Type 'using System;using System.Runtime.InteropServices;public static class TestContainerWindow{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);}'
$app=$null;$book=$null;$ownedNew=$false;$quit=$false;$results=@()
$r=[ordered]@{Passed=$false;SourceSHA256=(Get-FileHash -LiteralPath $source).Hash;FullAcceptanceClaim=$false}

# -------------------- ДЛЯ ТЕСТОВ --------------------
# Ищет единственную настройку по ключу, не завися от её координат.
function Set-TestContainerSetting([string]$Key,[object]$Value){
 $range=$book.Names.Item('rngSystemSettings').RefersToRange;$values=$range.Value2;$count=0
 for($row=1;$row-le$values.GetLength(0);$row++){if([string]$values[$row,1]-ceq$Key){$range.Cells.Item($row,2).Value2=$Value;$count++}}
 if($count-ne1){throw ('Setting not unique: '+$Key)}
}

# -------------------- ДЛЯ ТЕСТОВ --------------------
# Проверяет сохранённое положение и отсутствие второй схемы на исходном листе.
function Test-ContainerPlacement([string]$Stage){
 $chart=$book.Worksheets.Item('Лист схемы').ChartObjects().Item('chtNDMSectionPlot')
 if($book.Worksheets.Item('Расчет').ChartObjects().Count-ne0){throw ('Unexpected replacement chart: '+$Stage)}
 if([Math]::Abs($chart.Left-73)-gt0.01-or[Math]::Abs($chart.Top-91)-gt0.01-or[Math]::Abs($chart.Width-720)-gt0.01-or[Math]::Abs($chart.Height-460)-gt0.01){throw ('Placement changed: '+$Stage)}
 $count=0;foreach($shape in $chart.Chart.Shapes){if([string]$shape.Name-like'NDMPlot_Element*'){$count++}}
 if($count-eq0-and$chart.Chart.SeriesCollection.Count-eq0){throw ('Empty rendered section: '+$Stage)}
 return [ordered]@{Stage=$Stage;Sheet='Лист схемы';Name=[string]$chart.Name;Left=$chart.Left;Top=$chart.Top;Width=$chart.Width;Height=$chart.Height;ElementShapes=$count;Series=$chart.Chart.SeriesCollection.Count;NoReplacementOnCalculationSheet=$true}
}
try{
 $existing=@([Diagnostics.Process]::GetProcessesByName('EXCEL')|ForEach-Object {$_.Id})
 $app=New-Object -ComObject Excel.Application;[uint32]$id=0;[void][TestContainerWindow]::GetWindowThreadProcessId([IntPtr]$app.Hwnd,[ref]$id)
 if($id-le0-or$existing-contains[int]$id){$app=$null;throw 'New own Excel required'}
 $ownedNew=$true;$ticks=(Get-Process -Id $id).StartTime.ToUniversalTime().Ticks
 [ordered]@{PID=[int]$id;StartTicksUTC=[string]$ticks;FreeRAMGiB=$m.AvailableMBytes/1024;FreeCommitGiB=($m.CommitLimit-$m.CommittedBytes)/1GB}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $out 'Ownership.json') -Encoding UTF8
 $app.Visible=$false;$app.DisplayAlerts=$false;$app.EnableEvents=$false;$app.AutomationSecurity=1
 $book=$app.Workbooks.Open($path,0,$false)
 foreach($macro in @('modTestPlotConfig.RunPlotContainerTests')){
  Write-Output ('START '+$macro);$started=[DateTime]::UtcNow
  $text=[string]$app.Run("'"+$book.Name+"'!"+$macro)
  [IO.File]::WriteAllText((Join-Path $out ($macro+'.txt')),$text,[Text.UTF8Encoding]::new($false))
  $counts=[regex]::Matches($text,'(?im)\bfailed\s*=\s*(\d+)')
  $passed=$counts.Count-gt0-and@($counts|Where-Object {[long]$_.Groups[1].Value-ne0}).Count-eq0-and$text-notmatch'(?im)^FAIL\b'
  $results+=,[ordered]@{Macro=$macro;Passed=$passed;Seconds=([DateTime]::UtcNow-$started).TotalSeconds}
  $results|ConvertTo-Json -Depth 6|Set-Content -LiteralPath (Join-Path $out 'Macros.json') -Encoding UTF8
  Write-Output ($text-split'\r?\n'|Where-Object {$_-match'^TOTAL'}|Select-Object -Last 1)
  if(-not$passed){throw ('Directed suite failed: '+$macro)}
 }
 Set-TestContainerSetting 'Geometry.Type' 'Circle';Set-TestContainerSetting 'Results.SnapshotStorage' 'Excel'
 Set-TestContainerSetting 'General.NonCriticalMessagesEnabled' 'No';Set-TestContainerSetting 'Plot.AutoUpdateAfterCalculation' 'Yes'
 $book.Worksheets.Item('Расчет').ChartObjects().Item('chtNDMSectionPlot').Delete()
 $message=[string]$app.Run("'"+$book.Name+"'!modWorkbookCalculation.FormSectionForWorkbook",$book)
 [IO.File]::WriteAllText((Join-Path $out 'FormationMessage.txt'),$message,[Text.UTF8Encoding]::new($false))
 if($message-notmatch'автоматически создана') {throw 'Actual formation omitted recovery notice with ordinary messages disabled'}
 $book.Worksheets.Item('Расчет').ChartObjects().Item('chtNDMSectionPlot').Delete()
 $notice=[string]$app.Run("'"+$book.Name+"'!modTestPlotConfig.TestUpdatePlotContainerNotice",$book)
 [IO.File]::WriteAllText((Join-Path $out 'ManualRecoveryNotice.txt'),$notice,[Text.UTF8Encoding]::new($false))
 if($notice-notmatch'^True\|'-or$notice-notmatch'автоматически создана'){throw 'Manual recovery notice/attention flag absent'}
 $target=$book.Worksheets.Add();$target.Name='Лист схемы'
 $chart=$book.Worksheets.Item('Расчет').ChartObjects().Item('chtNDMSectionPlot');$relocated=$chart.Chart.Location(2,$target.Name);$relocated.Parent.Name='chtNDMSectionPlot'
 $chart=$target.ChartObjects().Item('chtNDMSectionPlot');$chart.Left=73;$chart.Top=91;$chart.Width=720;$chart.Height=460
 $book.Save();$book.Close($false);$book=$null
 $book=$app.Workbooks.Open($path,0,$false)
 $audit=@();$audit+=,(Test-ContainerPlacement 'SavedReopened')
 [void]$app.Run("'"+$book.Name+"'!modWorkbookCalculation.UpdateSectionGeometryPreviewForWorkbook",$book)
 $audit+=,(Test-ContainerPlacement 'FormationPreviewAfterReopen')
 [void]$app.Run("'"+$book.Name+"'!modCalculationSnapshot.UpdateSectionResultsForWorkbook",$book)
 $audit+=,(Test-ContainerPlacement 'ManualUpdateExcel')
 $loads=$book.Names.Item('rngLoadCombinations').RefersToRange
 $loads.Cells.Item(2,1).Value2='LC_CONTAINER';$loads.Cells.Item(2,2).Value2=1.0;$loads.Cells.Item(2,3).Value2=0.0;$loads.Cells.Item(2,4).Value2=0.0
 $formula=[string]$loads.Cells.Item(2,5).Validation.Formula1
 $profiles=$book.Worksheets.Item('Config').Range($formula.TrimStart('='));$loads.Cells.Item(2,5).Value2=[string]$profiles.Cells.Item(1,1).Value2
 $duration=$book.Names.Item('rngStabilityDurationLoads').RefersToRange
 $duration.Cells.Item(2,1).Value2='LC_CONTAINER';$duration.Cells.Item(2,2).Value2=0.0;$duration.Cells.Item(2,3).Value2=0.0;$duration.Cells.Item(2,4).Value2=0.0
 $message=[string]$app.Run("'"+$book.Name+"'!modWorkbookCalculation.RunSectionCalculationForWorkbook",$book)
 [IO.File]::WriteAllText((Join-Path $out 'CalculationMessage.txt'),$message,[Text.UTF8Encoding]::new($false))
 $audit+=,(Test-ContainerPlacement 'AutomaticAfterCalculation')
 Set-TestContainerSetting 'Results.SnapshotStorage' 'File';Set-TestContainerSetting 'Results.SnapshotName' 'TestMovablePlotContainer'
 [void]$app.Run("'"+$book.Name+"'!modWorkbookCalculation.FormSectionForWorkbook",$book)
 $audit+=,(Test-ContainerPlacement 'FormationFile')
 $duration=$book.Names.Item('rngStabilityDurationLoads').RefersToRange
 $duration.Cells.Item(2,1).Value2='LC_CONTAINER';$duration.Cells.Item(2,2).Value2=0.0;$duration.Cells.Item(2,3).Value2=0.0;$duration.Cells.Item(2,4).Value2=0.0
 $message=[string]$app.Run("'"+$book.Name+"'!modWorkbookCalculation.RunSectionCalculationForWorkbook",$book)
 [IO.File]::WriteAllText((Join-Path $out 'CalculationFileMessage.txt'),$message,[Text.UTF8Encoding]::new($false))
 [void]$app.Run("'"+$book.Name+"'!modCalculationSnapshot.UpdateSectionResultsForWorkbook",$book)
 $audit+=,(Test-ContainerPlacement 'ManualUpdateFile')
 $audit|ConvertTo-Json -Depth 6|Set-Content -LiteralPath (Join-Path $out 'Placement.json') -Encoding UTF8
 $book.Worksheets.Item('Лист схемы').ChartObjects().Item('chtNDMSectionPlot').Delete()
 $recovery=[string]$app.Run("'"+$book.Name+"'!modWorkbookCalculation.RunSectionCalculationForWorkbook",$book)
 [IO.File]::WriteAllText((Join-Path $out 'AutomaticRecoveryMessage.txt'),$recovery,[Text.UTF8Encoding]::new($false))
 if($recovery-notmatch'автоматически создана'-or$book.Worksheets.Item('Расчет').ChartObjects().Count-ne1){throw 'Automatic calculation omitted recovery notice'}
 $r.AutomaticCalculationRecoveryNotice=$true
 $r.Passed=$true;$r.PlacementStages=$audit.Count;$r.Macros=$results;$r.FormationRecoveryNotice=$true
}catch{$r.Error=$_.Exception.ToString();throw}
finally{
 if($book){$book.Close($false)};$book=$null
 if($app-and$ownedNew){$app.Quit();$quit=$true};$app=$null
 $chart=$null;$target=$null;$profiles=$null;$loads=$null;[GC]::Collect();[GC]::WaitForPendingFinalizers();[GC]::Collect()
 $r.QuitRequested=$quit;$r.ForceKilled=$false;$r.SourceAfterSHA256=(Get-FileHash -LiteralPath $source).Hash
 $r|ConvertTo-Json -Depth 8|Set-Content -LiteralPath (Join-Path $out 'Result.json') -Encoding UTF8
}
