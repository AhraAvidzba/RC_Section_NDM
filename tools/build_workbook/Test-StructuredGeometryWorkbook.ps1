# Проверяет лист сечений, независимость таблиц, списки и формулы единиц в отдельном Excel.
param([Parameter(Mandatory=$true)][string]$WorkbookPath,[Parameter(Mandatory=$true)][string]$ReportPath,[switch]$CheckDefaults,[string]$PreservedInputs='')
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
. (Join-Path $PSScriptRoot 'SettingsCatalog.ps1')
$path=(Resolve-Path -LiteralPath (Join-Path $root $WorkbookPath)).Path
$report=Join-Path $root $ReportPath
$hash=(Get-FileHash -LiteralPath $path).Hash
Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public static class SectionsValidationWindow{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);}'
$app=$null;$book=$null;$owned=0;$ticks=0;$defaults=0;$options=0;$preserved=0;$units=New-Object 'System.Collections.Generic.List[object]';$tables=@();$snapshots=@{}
try{
 $app=New-Object -ComObject Excel.Application
 $app.Visible=$false;$app.DisplayAlerts=$false;$app.EnableEvents=$false;$app.AutomationSecurity=3
 [uint32]$id=0;[void][SectionsValidationWindow]::GetWindowThreadProcessId([IntPtr]$app.Hwnd,[ref]$id)
 $owned=[int]$id;$ticks=(Get-Process -Id $owned).StartTime.ToUniversalTime().Ticks
 $book=$app.Workbooks.Open($path,0,$true)
 if(-not $book.ReadOnly){throw 'Read-only opening required.'}
 $order=@($book.Worksheets | ForEach-Object {$_.Name})
 if(($order -join '|') -cne 'Config|Сечения|Справка|Расчет|Results'){throw ('Sheet order: '+($order -join '|'))}
 $names=@($book.Names | ForEach-Object {$_.Name})
 if(@($names | Where-Object {$_ -match '_Choices$|rngRebarInput|RoundedRectangle'}).Count){throw 'Obsolete named ranges remain.'}
 $config=$book.Worksheets.Item('Config');$sections=$book.Worksheets.Item('Сечения')
 for($c=9;$c -le 24;$c++){if($CheckDefaults -and $config.Columns.Item($c).ColumnWidth -ne 17){throw "Config width C${c}: $($config.Columns.Item($c).ColumnWidth)"}}
 for($c=1;$c -le 11;$c++){if($sections.Columns.Item($c).ColumnWidth -ne 17){throw 'Sections columns must be equal.'}}
 # В скрытой области разрешены только три компактных динамических списка BR:BT.
 $values=$config.UsedRange.Value2
 for($r=1;$r -le $values.GetLength(0);$r++){for($c=73;$c -le $values.GetLength(1);$c++){if($null -ne $values[$r,$c]){throw "Obsolete hidden data R${r}C$c"}}}
 $last=0
 foreach($def in @(Get-GeometrySettingsCatalog)){
  $range=$book.Names.Item($def.RangeName).RefersToRange
  if($range.Worksheet.Name -cne 'Сечения' -or $range.Column -ne 1 -or $range.Row -le $last -or $range.Columns.Count -ne 11){throw ('Table placement: '+$def.RangeName)}
  $last=$range.Row+$range.Rows.Count
  if($range.FormatConditions.Count -ne 0 -or $range.WrapText -ne $false -or $range.RowHeight -ne 15){throw 'Static non-wrapping table formatting required.'}
  foreach($edge in 7,8,9,10,11,12){if($range.Borders.Item($edge).LineStyle -ne -4115){throw 'Full dotted border required.'}}
  $blocks=@(Get-GeometryTableBlocks $range)
  if(($blocks.Key -join '|') -cne 'General|Main|Extra'){throw 'One complete General/Main/Extra table required.'}
  for($b=0;$b -lt $blocks.Count;$b++){
   $block=$blocks[$b];$expected=$def.Blocks[$b]
   for($c=0;$c -lt $expected.Keys.Count;$c++){
    if([string]$range.Cells.Item($block.KeyRow,$c+1).Value2 -cne $expected.Keys[$c]){throw 'Column key mismatch.'}
    if([string]$range.Cells.Item($block.HeaderRow,$c+1).Value2 -cne $expected.Headers[$c]){throw 'Header mismatch.'}
    if($expected.Units -and $expected.Units[$c] -ceq 'мм'){$units.Add($range.Cells.Item($block.FirstRow,$c+1))}
   }
   $first=$block.FirstRow;if($expected.Units){$first++}
   if($block.LastRow-$first+1 -ne $expected.Rows.Count){throw 'Wrong row count.'}
   $note=$range.Cells.Item($block.HeaderRow,[int]$block.Columns['note']).Resize($block.LastRow-$block.HeaderRow+1,1)
   if($note.HorizontalAlignment -ne -4152 -or $note.WrapText -ne $false){throw 'Comment alignment changed.'}
   for($r=0;$r -lt $expected.Rows.Count;$r++){
    $row=$first+$r;$record=$expected.Rows[$r]
    if([string]$range.Cells.Item($row,1).Value2 -cne $record[0]){throw 'Local row key mismatch.'}
    foreach($key in $block.Columns.Keys){
     $c=[int]$block.Columns[$key];$cell=$range.Cells.Item($row,$c)
     if($key -eq 'unit' -and $record[$c-1] -ceq 'мм'){$units.Add($cell)}
     if($key -in @('key','name','unit','note')){continue}
     $editable=([string]$cell.Value2 -notin @('-','—') -and -not $cell.HasFormula)
     if($editable -and ($cell.Interior.Color -ne 15135487 -or $cell.HorizontalAlignment -ne -4108)){throw ('Input format: '+$def.RangeName+'/'+$block.Key+'/'+$record[0]+'/'+$key)}
     if(-not $editable -and $cell.Interior.Color -eq 15135487){throw 'Derived field highlighted as input.'}
     if($CheckDefaults){if([string]$cell.Value2 -cne [string]$record[$c-1]){throw ('Default mismatch: '+$def.RangeName+'/'+$block.Key+'/'+$record[0]+'/'+$key)};$defaults++}
     if($key -in @('bind2','bind3','pos2','pos3') -or ($key -eq 'val' -and $record[0] -match 'Shape|\.Type$')){
      $formula=[string]$cell.Validation.Formula1
      $expectedList='EachBar,EverySecondBar'
      if($key -like 'pos*'){$expectedList='Stacked,SideBySide'}
      elseif($record[0] -eq 'Shape'){$expectedList='Rectangle,LSection,TwoRectangles'}
      elseif($key -eq 'val'){$expectedList='Simple,Tapered'}
      if($formula.Replace(';',',') -cne $expectedList){throw ('Inline choices mismatch: '+$formula)}
      $old=$cell.Formula
      foreach($option in $formula.Split([string]$app.International(5))){$cell.Value2=$option;if(-not $cell.Validation.Value){throw "Valid choice rejected: $($def.RangeName)/$($block.Key)/$($record[0])/$key; option=$option; Formula1=$($cell.Validation.Formula1); Value2=$($cell.Value2); separator=$($app.International(5))"};$options++}
      $cell.Value2='__INVALID__';if($cell.Validation.Value){throw 'Invalid choice accepted.'};$cell.Formula=$old
     }
    }
   }
  }
  $snapshots[$def.RangeName]=$range.Formula
  $tables+=,[ordered]@{Name=$def.RangeName;Address=$range.Address();Blocks=$blocks.Key;RebarRows=$def.Blocks[1].Rows.Count}
 }
 if($PreservedInputs){
  $receipt=Get-Content -LiteralPath (Join-Path $root $PreservedInputs) -Raw -Encoding UTF8 | ConvertFrom-Json
  foreach($field in $receipt.InputFields){
   $range=$book.Names.Item($field.Table).RefersToRange
   $cell=Get-GeometryTableCell $range $field.Block $field.Row $field.Column
   if([string]$cell.Value2 -cne [string]$field.Value){throw ('Preserved input differs: '+($field|ConvertTo-Json -Compress))};$preserved++
  }
  if($preserved -ne 294){throw "Expected 294 preserved fields; actual=$preserved"}
 }
 $settings=$book.Names.Item('rngSystemSettings').RefersToRange;$selector=$null
 for($r=2;$r -le $settings.Rows.Count;$r++){if([string]$settings.Cells.Item($r,1).Value2 -eq 'Geometry.Type'){$selector=$settings.Cells.Item($r,2)}}
 foreach($type in @('Circle','RectSet','PointedRect','HollowRectangle')){
  $selector.Value2=$type;$config.Calculate();$sections.Calculate()
  foreach($name in $snapshots.Keys){$now=$book.Names.Item($name).RefersToRange.Formula;$old=$snapshots[$name];for($r=1;$r -le $old.GetLength(0);$r++){for($c=1;$c -le $old.GetLength(1);$c++){if([string]$old[$r,$c] -cne [string]$now[$r,$c]){throw 'Type switch changed stored inputs.'}}}}
 }
 $unitRange=$book.Names.Item('rngUnitSettings').RefersToRange;$lengthCell=$null
 for($r=2;$r -le $unitRange.Rows.Count;$r++){if([string]$unitRange.Cells.Item($r,1).Value2 -eq 'Length'){$lengthCell=$unitRange.Cells.Item($r,2)}}
 if($null -eq $lengthCell -or $units.Count -eq 0){throw 'Length unit coverage missing.'}
 foreach($entry in @(@('mm','mm'),@('cm','cm'),@('m','m'))){
  $lengthCell.Value2=$entry[0];$unitRange.Calculate();$sections.Calculate()
  foreach($cell in $units){if(-not $cell.HasFormula -or [string]$cell.Value2 -cne $entry[1]){throw "Unit formula mismatch: $($cell.Address()); expected=$($entry[1]); actual=$($cell.Value2); formula=$($cell.Formula)"}}
 }
 $book.Close($false);$book=$null
 [ordered]@{Status='PASS';WorkbookSHA256=$hash;SheetOrder=$order;Names=$names.Count;ChoiceNames=0;Tables=$tables;DefaultFieldsChecked=$defaults;PreservedFields=$preserved;ValidDropdownChoicesChecked=$options;DynamicUnitCells=$units.Count;LengthUnits=@('mm','cm','m');IndependentInputsPreservedOnSwitch=$true;ConditionalFormattingRules=0;HiddenDynamicColumns='BR:BT';ReadOnly=$true} | ConvertTo-Json -Depth 7 | Set-Content -LiteralPath $report -Encoding UTF8
}finally{
 if($book){$book.Close($false)}
 if($app){$app.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
 if($owned){$p=Get-Process -Id $owned -ErrorAction SilentlyContinue;if($p -and $p.ProcessName -eq 'EXCEL' -and $p.StartTime.ToUniversalTime().Ticks -eq $ticks){Stop-Process -Id $owned -Force}}
}
if((Get-FileHash -LiteralPath $path).Hash -cne $hash){throw 'Source changed during read-only check.'}
Write-Output "SECTIONS_STRUCTURE_PASS: defaults=$defaults; inputs=$preserved; units=$($units.Count); choices=$options"
