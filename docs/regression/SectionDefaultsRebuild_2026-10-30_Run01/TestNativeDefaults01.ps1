# -------------------- ДЛЯ ТЕСТОВ --------------------
# Проверяет чистую сохранённую сборку и каждый пункт списков; открывает только собственную копию без макросов и сохранения.
param()
$ErrorActionPreference='Stop'
Import-Module 'C:/Windows/System32/WindowsPowerShell/v1.0/Modules/Microsoft.PowerShell.Utility/Microsoft.PowerShell.Utility.psd1' -Global
$p=$PSScriptRoot;$root='C:/Users/avidzba/Desktop/программы/В_РАБОТЕ!/RC_Section_NDM'
. (Join-Path $root 'tools/build_workbook/GeometryTables.ps1')
$path=Join-Path $p 'Build/RC_Section_NDM.xlsm';$out=Join-Path $p 'Native01'
if(Test-Path -LiteralPath $out){throw 'Inspect existing native attempt first'}
[void][IO.Directory]::CreateDirectory($out)
$memory=Get-CimInstance Win32_PerfFormattedData_PerfOS_Memory
if($memory.AvailableMBytes-lt8192-or($memory.CommitLimit-$memory.CommittedBytes)-lt14GB){throw 'Insufficient RAM/commit reserve'}
Add-Type 'using System;using System.Runtime.InteropServices;public static class TestDefaultsWindow{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);}'
$app=$null;$book=$null;$owned=0;$ownedNew=$false;$quit=$false
$result=[ordered]@{Passed=$false;BookSHA256=(Get-FileHash -LiteralPath $path).Hash;SavedReopened=$true;ReadOnly=$true;MacrosRun=$false;SavedAgain=$false;FullAcceptanceClaim=$false}

# -------------------- ДЛЯ ТЕСТОВ --------------------
# Проверяет варианты отдельно и отклонение склеенных строк без нормализации разделителей; возвращает исходное значение.
function Test-DefaultsDropdown([string]$Key,[object]$Cell,[string[]]$Choices){
 $original=$Cell.Formula
 try{
  if($Cell.Validation.Type-ne3-or-not$Cell.Validation.InCellDropdown){throw ('Missing saved list: '+$Key)}
  foreach($choice in $Choices){$Cell.Value2=$choice;if(-not$Cell.Validation.Value){throw ('Rejected separate choice: '+$Key+'/'+$choice)}}
  foreach($bad in @(($Choices-join','),($Choices-join';'),'__UNKNOWN_CHOICE__')){$Cell.Value2=$bad;if($Cell.Validation.Value){throw ('Accepted invalid list choice: '+$Key+'/'+$bad)}}
  return [ordered]@{Key=$Key;Address=$Cell.Address();Choices=$Choices;Accepted=$Choices.Count;Rejected=3;Formula1=[string]$Cell.Validation.Formula1}
 }finally{if($null-eq$original){$Cell.ClearContents()|Out-Null}else{$Cell.Formula=$original}}
}
try{
 $existing=@([Diagnostics.Process]::GetProcessesByName('EXCEL')|ForEach-Object {$_.Id})
 $app=New-Object -ComObject Excel.Application
 [uint32]$id=0;[void][TestDefaultsWindow]::GetWindowThreadProcessId([IntPtr]$app.Hwnd,[ref]$id);$owned=[int]$id
 if($owned-le0-or$existing-contains$owned){$app=$null;throw 'New own Excel required'}
 $ownedNew=$true;$ticks=(Get-Process -Id $owned).StartTime.ToUniversalTime().Ticks
 [ordered]@{PID=$owned;StartTicksUTC=[string]$ticks;FreeRAMGiB=$memory.AvailableMBytes/1024;FreeCommitGiB=($memory.CommitLimit-$memory.CommittedBytes)/1GB}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $out 'Ownership.json') -Encoding UTF8
 $app.Visible=$false;$app.DisplayAlerts=$false;$app.EnableEvents=$false;$app.AutomationSecurity=3
 $book=$app.Workbooks.Open($path,0,$true);if(-not$book.ReadOnly){throw 'Readonly opening required'}
 $project=$book.VBProject;if($project.VBComponents.Count-ne181){throw '181 components required'}
 $dump=New-Object Text.StringBuilder;[void]$dump.AppendLine('VBA PROJECT EXPORT');[void]$dump.AppendLine('Date: '+[DateTime]::UtcNow.ToString('o'))
 foreach($component in $project.VBComponents){
  $code=$component.CodeModule;$count=[int]$code.CountOfLines;$body=if($count){[string]$code.Lines(1,$count)}else{''}
  [void]$dump.AppendLine('COMPONENT: '+$component.Name);[void]$dump.AppendLine('TYPE: '+$component.Type);[void]$dump.AppendLine('LINES: '+$count);[void]$dump.AppendLine('');[void]$dump.AppendLine($body)
 }
 [IO.File]::WriteAllText((Join-Path $out 'VBA_All_Code.txt'),$dump.ToString(),[Text.UTF8Encoding]::new($false))
 $result.Components=181
 $layout=Get-Content -LiteralPath (Join-Path $root 'tools/build_workbook/WorkbookLayout.json') -Raw -Encoding UTF8|ConvertFrom-Json
 $audit=@();$lists=@();$dashes=0
 foreach($definition in @(Get-GeometrySettingsCatalog)){
  $table=$book.Names.Item($definition.RangeName).RefersToRange
  $bounds=@($layout.Names|Where-Object Name -eq $definition.RangeName)[0]
  if($table.Address()-cne$bounds.Address-or$table.Rows.Count-ne$bounds.Rows-or$table.Columns.Count-ne$bounds.Columns){throw ('Table bounds differ: '+$definition.Type)}
  foreach($block in $definition.Blocks){foreach($row in $block.Rows){
   for($column=2;$column-lt$block.Keys.Count;$column++){
    $key=[string]$block.Keys[$column];$expected=$row[$column];$cell=Get-GeometryTableCell $table $block.Key ([string]$row[0]) $key
    if($expected -is [pscustomobject] -and $expected.PSObject.Properties['Formula']){if(-not$cell.HasFormula){throw 'Missing catalog formula'};continue}
    if($key-ceq'unit'-and[string]$expected-ceq'мм'){
     if(-not$cell.HasFormula-or[string]$cell.Value2-cne'mm'){throw 'Input unit must reference configured millimetres'}
    }elseif([string]$cell.Value2-cne[string]$expected){throw ('Catalog value differs: '+$definition.Type+'/'+$block.Key+'/'+$row[0]+'/'+$key+' actual='+$cell.Value2+' expected='+$expected)}
    if($cell.WrapText){throw 'Unexpected geometry wrapping'}
    if($key-eq'note'){if($cell.HorizontalAlignment-ne-4152){throw 'Comment must be right aligned'}}
    elseif($key-notin@('unit')){
     if([string]$expected-ceq[char]0x2014){if($cell.Interior.Color-eq15135487){throw 'Projection dash marked editable'};$dashes++}
     elseif(-not$cell.HasFormula-and$cell.Interior.Color-ne15135487){throw ('Editable fill missing: '+$definition.Type+'/'+$row[0]+'/'+$key)}
    }
    foreach($edge in 7..10){if($cell.Borders.Item($edge).LineStyle-eq-4142){throw 'Missing geometry cell border'}}
    $audit+=,[ordered]@{Type=$definition.Type;Block=$block.Key;Row=[string]$row[0];Field=$key;Value=$cell.Value2;Address=$cell.Address()}
    if($key-eq'endMode'-and[string]$row[0]-notlike'Opening*'){$lists+=,(Test-DefaultsDropdown ($definition.Type+'/'+$row[0]+'/'+$key) $cell @('IncludeBoth','ExcludeBoth','IncludeStart','IncludeEnd'))}
    if($key-like'bind*'){$lists+=,(Test-DefaultsDropdown ($definition.Type+'/'+$row[0]+'/'+$key) $cell @('EachBar','EverySecondBar'))}
    if($key-like'pos*'){$lists+=,(Test-DefaultsDropdown ($definition.Type+'/'+$row[0]+'/'+$key) $cell @('Stacked','SideBySide'))}
   }
  }}
 }
 $audit|ConvertTo-Json -Depth 7|Set-Content -LiteralPath (Join-Path $out 'CatalogReadback.json') -Encoding UTF8
 $range=$book.Names.Item('rngSystemSettings').RefersToRange;$data=$range.Value2;$settings=@{}
 for($r=2;$r-le$data.GetLength(0);$r++){if([string]$data[$r,1]){$settings[[string]$data[$r,1]]=[ordered]@{Value=$data[$r,2];Row=$r}}}
 $choices=@{'Geometry.Type'=@('Circle','Rectangles','PointedSection','HollowRectangle');'Mesh.CellGeometry'=@('Adaptive','Contour');'Mesh.StepMode'=@('Auto 0.05','Auto 0.1','Auto 0.2','User');'Results.SnapshotStorage'=@('Excel','File');'Output.CombinationMode'=@('Worst','User')}
 foreach($key in $choices.Keys){$lists+=,(Test-DefaultsDropdown $key $range.Cells.Item($settings[$key].Row,2) $choices[$key])}
 foreach($pair in @(@('Mesh.CellGeometry','Contour'),@('Mesh.StepMode','Auto 0.1'),@('Results.SnapshotStorage','Excel'),@('Output.CombinationMode','Worst'))){if($settings[$pair[0]].Value-cne$pair[1]){throw ('Wrong clean Config default: '+$pair[0])}}
 $lists|ConvertTo-Json -Depth 7|Set-Content -LiteralPath (Join-Path $out 'NativeDropdowns.json') -Encoding UTF8
 $result.CatalogCells=$audit.Count;$result.Dropdowns=$lists.Count;$result.ProjectionDashes=$dashes
 foreach($name in @('rngNDMSectionGeometry','rngNDMSectionProperties','rngNDMMaterialDiagrams','rngNDMSectionContours','rngNDMElementResults','rngNDMElementGeometry','rngNDMSectionAnnotations')){if([string]$book.Names.Item($name).RefersToRange.Offset(1,0).Value2-ne''){throw ('Unclean Results: '+$name)}}
 $loads=$book.Names.Item('rngLoadCombinations').RefersToRange.Value2
 for($r=2;$r-le$loads.GetLength(0);$r++){foreach($c in 1..4){if([string]$loads[$r,$c]){throw 'User loads copied into fresh book'}}}
 $plot=$book.Worksheets.Item('Расчет');$shapes=0
 foreach($chartObject in $plot.ChartObjects()){foreach($shape in $chartObject.Chart.Shapes){if([string]$shape.Name-match'^NDMPlot_Element(Concrete|Rebar)'){$shapes++}}}
 if($shapes){throw 'Previous plot copied'}
 $order=($book.Worksheets|ForEach-Object {$_.Name})-join'|'
 if($order-cne'Config|Сечения|Справка|Расчет|Results'){throw 'Wrong sheet order'}
 $result.CleanLoadsResultsAndPlot=$true;$result.SheetOrder=$order
 $scratch=$app.Workbooks.Add()
 try{
  foreach($image in @(@('Сечения','A60:K88','PointedSection'),@('Сечения','A92:K128','HollowRectangle'),@('Config','A1:K30','Config'),@('Справка','A711:F723','Help'),@('Расчет','A1:AJ38','Calculation'),@('Results','A150:AF159','Results'))){
   $sheet=$book.Worksheets.Item($image[0]);$capture=$sheet.Range($image[1]);$sheet.Activate();$capture.CopyPicture(1,2)
   $canvas=$scratch.Worksheets.Item(1);$canvas.Activate();$box=$canvas.ChartObjects().Add(0,0,[double]$capture.Width,[double]$capture.Height)
   try{$box.Activate();$box.Chart.Paste();if(-not$box.Chart.Export((Join-Path $out ($image[2]+'.png')),'PNG')){throw 'PNG export failed'}}finally{$box.Delete();$box=$null}
  }
 }finally{$scratch.Close($false);$scratch=$null}
 $result.Passed=$true
}catch{$result.Error=$_.Exception.ToString();throw}
finally{
 if($book){$book.Close($false)};$book=$null
 if($app-and$ownedNew){$app.Quit();$quit=$true};$app=$null
 $code=$null;$component=$null;$project=$null;$table=$null;$cell=$null;$range=$null;$plot=$null;$shape=$null;$chartObject=$null;$capture=$null;$canvas=$null;$sheet=$null
 [GC]::Collect();[GC]::WaitForPendingFinalizers();[GC]::Collect()
 $result.QuitRequested=$quit;$result.ForceKilled=$false;$result.BookAfterSHA256=(Get-FileHash -LiteralPath $path).Hash
 if($result.BookSHA256-cne$result.BookAfterSHA256){$result.Passed=$false;$result.Error='Readonly package changed'}
 $result|ConvertTo-Json -Depth 8|Set-Content -LiteralPath (Join-Path $out 'Result.json') -Encoding UTF8
}
