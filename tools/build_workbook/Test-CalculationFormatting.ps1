# Проверяет обычный расчёт и обновление схемы в отдельной копии книги.
# Сохраняет снимки для сравнения оформления всех листов, кроме Results.
param([Parameter(Mandatory=$true)][string]$WorkbookPath,[Parameter(Mandatory=$true)][string]$EvidenceDirectory)
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$source=(Resolve-Path -LiteralPath (Join-Path $root $WorkbookPath)).Path
$directory=[IO.Path]::GetFullPath((Join-Path $root $EvidenceDirectory))
$allowed=[IO.Path]::GetFullPath((Join-Path $root 'docs/regression'))+[IO.Path]::DirectorySeparatorChar
if(-not $directory.StartsWith($allowed,[StringComparison]::OrdinalIgnoreCase)){throw 'Evidence must remain within docs/regression.'}
[void][IO.Directory]::CreateDirectory($directory)
$hash=(Get-FileHash -LiteralPath $source).Hash
$private=Join-Path $directory 'Runtime.xlsm'
[IO.File]::Copy($source,$private,$false)
. (Join-Path $PSScriptRoot 'SettingsCatalog.ps1')
$areas=@(Get-WorkbookPrintAreas $private)
Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public static class RuntimeLayoutWindow{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);}'
$app=$null;$book=$null;$owned=0;$ticks=0;$results=@()
try{
 $app=New-Object -ComObject Excel.Application
 $app.Visible=$false;$app.DisplayAlerts=$false;$app.EnableEvents=$false;$app.AutomationSecurity=1
 [uint32]$id=0;[void][RuntimeLayoutWindow]::GetWindowThreadProcessId([IntPtr]$app.Hwnd,[ref]$id)
 $owned=[int]$id;$ticks=(Get-Process -Id $owned).StartTime.ToUniversalTime().Ticks
 Write-Output ('OWNED_EXCEL: id='+$owned+'; utcStartTicks='+$ticks)
 $book=$app.Workbooks.Open($private,0,$false)
 if($book.ReadOnly){throw 'Private workbook is read-only.'}
 $prefix="'"+$book.Name+"'!"
 [void]$app.Run($prefix+'modTestShapeArchitecture.ConfigureShapeFixture','Circle')
 [void]$app.Run($prefix+'modTestPerformance.ConfigurePerformanceFixture','DIRECT',1,'Newton','Saved','Normal','No')
 $settings=$book.Names.Item('rngSystemSettings').RefersToRange
 foreach($key in @('Plot.Enabled','General.NonCriticalMessagesEnabled')){
  for($r=2;$r -le $settings.Rows.Count;$r++){if([string]$settings.Cells.Item($r,1).Value2 -eq $key){$settings.Cells.Item($r,2).Value2=if($key -eq 'Plot.Enabled'){'Yes'}else{'No'}}}
 }
 # Пользовательское оформление должно сохраниться, даже если оно отличается от шаблона.
 foreach($sheet in $book.Worksheets){
  if($sheet.Name -eq 'Results'){continue}
  $sheet.Rows.Item(1).RowHeight=26.25;$sheet.Columns.Item(2).ColumnWidth=23.5
  $sheet.Range('A1').Font.Color=8388736;$sheet.Range('A1').Interior.Color=16777164
 }
 $book.SaveCopyAs((Join-Path $directory 'BeforeCalculation.xlsm'))
 foreach($iteration in 1,2){
  $watch=[Diagnostics.Stopwatch]::StartNew()
  $message=[string]$app.Run($prefix+'modWorkbookCalculation.RunSectionCalculationForWorkbook',$book,$true)
  $watch.Stop()
  if($message -match 'NumFail|InputErr|CalcErr|ошиб'){throw ('Calculation failed: '+$message)}
  $results+=,[ordered]@{Stage=('Calculation'+$iteration);Seconds=$watch.Elapsed.TotalSeconds;Message=$message;PlotSeries=[int]$book.Worksheets.Item('Расчет').ChartObjects('chtNDMSectionPlot').Chart.SeriesCollection().Count}
  $book.SaveCopyAs((Join-Path $directory ('AfterCalculation'+$iteration+'.xlsm')))
  [void]$app.Run($prefix+'modWorkbookCalculation.UpdateSectionPlotForWorkbook',$book,$true)
  $book.SaveCopyAs((Join-Path $directory ('AfterPlot'+$iteration+'.xlsm')))
 }
 $tableTest=[string]$app.Run($prefix+'modTestWorkbookInterface.RunSavedLoadTableLayoutTests')
 if($tableTest -notmatch 'TOTAL_LOAD_TABLE_LAYOUT: passed=\d+; failed=0'){throw $tableTest}
 [ordered]@{Status='RUNTIME_COMPLETED';SourceWorkbookSHA256=$hash;Calculations=$results;LoadTableTest=$tableTest;OwnedExcelPID=$owned;OwnedExcelStartTicks=$ticks} | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $directory 'Run.json') -Encoding UTF8
}finally{
 if($book){$book.Close($false)}
 if($app){$app.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
 if($owned){$p=Get-Process -Id $owned -ErrorAction SilentlyContinue;if($p -and $p.ProcessName -eq 'EXCEL' -and $p.StartTime.ToUniversalTime().Ticks -eq $ticks){Stop-Process -Id $owned -Force}}
}
if((Get-FileHash -LiteralPath $source).Hash -cne $hash){throw 'Source changed during formatting check.'}
foreach($name in @('BeforeCalculation','AfterCalculation1','AfterCalculation2','AfterPlot1','AfterPlot2')){Restore-WorkbookPrintAreas (Join-Path $directory ($name+'.xlsm')) $areas}
Write-Output 'RUNTIME_COMPLETED: compare saved formatting snapshots before accepting.'
