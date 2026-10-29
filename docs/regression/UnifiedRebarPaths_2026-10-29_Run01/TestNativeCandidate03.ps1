# -------------------- ДЛЯ ТЕСТОВ --------------------
# Проверяет сохранённую чистую сборку без запуска макросов и сохранения.
$ErrorActionPreference='Stop'
Import-Module 'C:/Windows/System32/WindowsPowerShell/v1.0/Modules/Microsoft.PowerShell.Utility/Microsoft.PowerShell.Utility.psd1' -Global
$repoOuter=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../../..')).Path
$directory='C:/Users/avidzba/.codex/visualizations/2026/10/12/01a12ece-8192-7141-89b3-f83f1176184c/UnifiedRebarPaths_2026-10-29_Run01'
Add-Type 'using System;using System.Runtime.InteropServices;public static class TestNativeFinalOwner{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);}'
$m=Get-CimInstance Win32_PerfFormattedData_PerfOS_Memory
if($m.AvailableMBytes-lt8192 -or ($m.CommitLimit-$m.CommittedBytes)-lt14GB){throw 'Memory reserve insufficient'}
$app=$null;$book=$null;$project=$null;$component=$null;$code=$null
$ownedNew=$false
try{
 $preexistingExcelIDs=@([Diagnostics.Process]::GetProcessesByName('EXCEL')|ForEach-Object {$_.Id})
 $app=New-Object -ComObject Excel.Application
 [uint32]$id=0;[void][TestNativeFinalOwner]::GetWindowThreadProcessId([IntPtr]$app.Hwnd,[ref]$id);$owned=[int]$id;$ticks=(Get-Process -Id $owned).StartTime.ToUniversalTime().Ticks
 if($preexistingExcelIDs-contains$owned){$app=$null;throw 'Existing Excel returned; no user instance modified or closed'}
 $ownedNew=$true
 [ordered]@{PID=$owned;StartTicksUTC=[string]$ticks}|ConvertTo-Json|Set-Content (Join-Path $directory 'Native03Ownership.json') -Encoding UTF8
 $app.Visible=$false;$app.DisplayAlerts=$false;$app.EnableEvents=$false;$app.AutomationSecurity=3
 foreach($case in @('Main')){
  $path=Join-Path $PSScriptRoot 'Build03/RC_Section_NDM.xlsm';$hash=(Get-FileHash -LiteralPath $path).Hash
  $book=$app.Workbooks.Open($path,0,$true);if(-not$book.ReadOnly){throw 'Readonly required'}
  $project=$book.VBProject;if($project.VBComponents.Count-ne181){throw '181 real components required'}
  $out=Join-Path $directory ('Native03/'+$case);[void][IO.Directory]::CreateDirectory($out)
  $all=New-Object Text.StringBuilder;[void]$all.AppendLine('VBA PROJECT EXPORT');[void]$all.AppendLine('Workbook: '+$book.Name);[void]$all.AppendLine('Date: '+(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
  $names=@()
  for($i=1;$i-le$project.VBComponents.Count;$i++){
   $component=$project.VBComponents.Item($i);$code=$component.CodeModule;$n=[int]$code.CountOfLines;$body=if($n){[string]$code.Lines(1,$n)}else{''}
   [void]$all.AppendLine('='*100);[void]$all.AppendLine('COMPONENT: '+$component.Name);[void]$all.AppendLine('TYPE: '+$component.Type);[void]$all.AppendLine('LINES: '+$n);[void]$all.AppendLine('='*100);if($n){[void]$all.AppendLine($body)};[void]$all.AppendLine()
   [IO.File]::WriteAllText((Join-Path $out ($component.Name+'.txt')),$body,[Text.Encoding]::UTF8);$names+=[ordered]@{Name=[string]$component.Name;Type=[int]$component.Type;Lines=$n}
  }
  [IO.File]::WriteAllText((Join-Path $out 'VBA_All_Code.txt'),$all.ToString(),[Text.Encoding]::UTF8)
  $range=$book.Names.Item('rngSystemSettings').RefersToRange;$data=$range.Value2;$settings=@{}
  for($row=2;$row-le$data.GetLength(0);$row++){
   $key=[string]$data[$row,1]
   if($key-and-not$key.StartsWith('[')){$settings[$key]=[ordered]@{Value=$data[$row,2];Address=$range.Cells.Item($row,2).Address()}}
  }
  foreach($pair in @(@('Mesh.CellGeometry','Contour'),@('Mesh.StepMode','Auto 0.1'),@('Results.SnapshotStorage','Excel'),@('Output.CombinationMode','Worst'))){if($settings[$pair[0]].Value-cne$pair[1]){throw ('Incorrect clean default '+$pair[0])}}
  foreach($key in @('Mesh.CellGeometry','Geometry.Type','Results.SnapshotStorage','Output.CombinationMode')){
   $cell=$book.Worksheets.Item('Config').Range($settings[$key].Address)
   if($cell.Validation.Type-ne3-or-not$cell.Validation.InCellDropdown){throw ('Required saved dropdown missing: '+$key)}
   $settings[$key]['ValidationFormula']=[string]$cell.Validation.Formula1
  }
  $cell=$null
  $clean=@{}
  foreach($name in @('rngNDMSectionGeometry','rngNDMSectionProperties','rngNDMMaterialDiagrams','rngNDMSectionContours','rngNDMElementResults','rngNDMElementGeometry','rngNDMSectionAnnotations')){
   $range=$book.Names.Item($name).RefersToRange
   if([string]$range.Offset(1,0).Value2-ne''){throw ('Not a clean snapshot: '+$name)}
   $clean[$name]=$range.Address()
  }
  $range=$book.Names.Item('rngLoadCombinations').RefersToRange;$data=$range.Value2
  for($row=2;$row-le$data.GetLength(0);$row++){foreach($column in 1..4){if([string]$data[$row,$column]-ne''){throw 'User loads must not be copied'}}}
  $sheetOrder=(($book.Worksheets|ForEach-Object {$_.Name})-join'|')
  if($sheetOrder-cne'Config|Сечения|Справка|Расчет|Results'){throw 'Unexpected sheet order'}
  $plotSheet=$book.Worksheets.Item('Расчет');$plotCharts=$plotSheet.ChartObjects();$elementShapes=0;$charts=@()
  for($j=1;$j-le$plotCharts.Count;$j++){
   $plotChart=$plotCharts.Item($j).Chart
   for($k=1;$k-le$plotChart.Shapes.Count;$k++){if([string]$plotChart.Shapes.Item($k).Name-match'^NDMPlot_Element(Concrete|Rebar)'){$elementShapes++}}
   $charts+=,[ordered]@{Name=[string]$plotCharts.Item($j).Name;Shapes=$plotChart.Shapes.Count;Series=$plotChart.SeriesCollection().Count}
  }
  if($elementShapes-ne0){throw 'Clean template contains previous section element shapes'}
  [ordered]@{FromWorkbookLayout=$true;Charts=$charts;ElementShapes=$elementShapes}|ConvertTo-Json -Depth 5|Set-Content (Join-Path $out 'EmptyPlotContainers.json') -Encoding UTF8
  $plotChart=$null;$plotCharts=$null;$plotSheet=$null
  [ordered]@{Settings=$settings;CleanSnapshotAnchors=$clean;SheetOrder=$sheetOrder;UserLoadsCopied=$false;PreviousPlotCopied=$false;CalculationsOrBehaviorTestsRun=$false}|ConvertTo-Json -Depth 7|Set-Content (Join-Path $out 'CleanStructure.json') -Encoding UTF8
  $range=$null
  [ordered]@{Case=$case;SHA256=$hash;ReadOnly=[bool]$book.ReadOnly;MacroRunImportSave=$false;Components=$names;ConfigPresent=($book.Worksheets.Item('Config').Name-eq'Config');GeometryAnchor=$book.Names.Item('rngNDMSectionGeometry').RefersTo;BoundariesAnchor=$book.Names.Item('rngNDMElementGeometry').RefersTo}|ConvertTo-Json -Depth 6|Set-Content (Join-Path $out 'NativeVerification.json') -Encoding UTF8
  $code=$null;$component=$null;$project=$null;[GC]::Collect();[GC]::WaitForPendingFinalizers()
  $book.Close($false);$book=$null;[GC]::Collect();[GC]::WaitForPendingFinalizers()
  if((Get-FileHash -LiteralPath $path).Hash-cne$hash){throw 'Readonly package changed'}
  Write-Output ('NATIVE_READONLY: '+$case+'; 181 components')
 }
}catch{
 [ordered]@{PrimaryError=$_.Exception.ToString();Stack=$_.ScriptStackTrace}|ConvertTo-Json|Set-Content (Join-Path $directory 'Native03Failure.json') -Encoding UTF8
 throw
}finally{
 $cell=$null;$range=$null;$code=$null;$component=$null;$project=$null;[GC]::Collect();[GC]::WaitForPendingFinalizers()
 if($book){$book.Close($false)};$book=$null;[GC]::Collect();[GC]::WaitForPendingFinalizers()
 if($app-and$ownedNew){$app.Quit()};$app=$null;[GC]::Collect();[GC]::WaitForPendingFinalizers();[GC]::Collect();[GC]::WaitForPendingFinalizers()
 [ordered]@{PID=$owned;StartTicksUTC=[string]$ticks;QuitRequested=$true;ForceKilled=$false}|ConvertTo-Json|Set-Content (Join-Path $directory 'Native03Quit.json') -Encoding UTF8
}
