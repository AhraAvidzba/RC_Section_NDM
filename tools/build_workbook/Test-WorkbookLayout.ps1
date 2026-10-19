# Сверяет фактическую геометрию всей книги с сохранённым расположением.
# Проверка читает отдельный файл, не меняя проверяемую книгу.
param([Parameter(Mandatory=$true)][string]$WorkbookPath,[Parameter(Mandatory=$true)][string]$ReportPath,[string]$LayoutPath='tools/build_workbook/WorkbookLayout.json')
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'WorkbookLayout.ps1')
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$path=(Resolve-Path -LiteralPath (Join-Path $root $WorkbookPath)).Path
$layout=Read-WorkbookLayout (Join-Path $root $LayoutPath);$xml=Get-WorkbookLayoutXmlSheets $path
$hash=(Get-FileHash -LiteralPath $path).Hash;$checks=New-Object 'System.Collections.Generic.List[object]'
function Check-Layout([string]$Name,[bool]$Passed,[object]$Actual,[object]$Expected){$checks.Add([ordered]@{Name=$Name;Passed=$Passed;Actual=$Actual;Expected=$Expected})}
Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public static class LayoutCheckWindow{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);}'
$app=$null;$book=$null;$owned=0;$ticks=0
try{
 $app=New-Object -ComObject Excel.Application;$app.Visible=$false;$app.DisplayAlerts=$false;$app.EnableEvents=$false;$app.AutomationSecurity=3
 [uint32]$id=0;[void][LayoutCheckWindow]::GetWindowThreadProcessId([IntPtr]$app.Hwnd,[ref]$id);$owned=[int]$id;$ticks=(Get-Process -Id $owned).StartTime.ToUniversalTime().Ticks
 $book=$app.Workbooks.Open($path,0,$true)
 if(-not $book.ReadOnly){throw 'Read-only opening required.'}
 $actual=Get-WorkbookLayoutSnapshot $book $xml
 Check-Layout 'SheetOrder' (($actual.Sheets.Name -join '|') -ceq ($layout.Sheets.Name -join '|')) ($actual.Sheets.Name -join '|') ($layout.Sheets.Name -join '|')
 Check-Layout 'AllNames' (($actual.Names.Name -join '|') -ceq ($layout.Names.Name -join '|')) ($actual.Names.Name -join '|') ($layout.Names.Name -join '|')
 foreach($item in $layout.Names){$found=@($actual.Names | Where-Object Name -eq $item.Name)[0];Check-Layout ('Name.'+$item.Name) ($found.Sheet -ceq $item.Sheet -and $found.Address -ceq $item.Address) ($found.Sheet+'!'+$found.Address) ($item.Sheet+'!'+$item.Address)}
 foreach($item in $layout.OtherTables){$found=@($actual.OtherTables | Where-Object Key -eq $item.Key)[0];Check-Layout ('Table.'+$item.Key) ($found.Sheet -ceq $item.Sheet -and $found.Address -ceq $item.Address) ($found.Sheet+'!'+$found.Address) ($item.Sheet+'!'+$item.Address)}
 foreach($item in $layout.Sheets){
  $found=@($actual.Sheets | Where-Object Name -eq $item.Name)[0]
  Check-Layout ($item.Name+'.Visibility') ($found.Visible -eq $item.Visible) $found.Visible $item.Visible
  Check-Layout ($item.Name+'.PrintArea') ($found.PrintArea -ceq $item.PrintArea) $found.PrintArea $item.PrintArea
  $bounds=@(1,16385)+@($item.Columns | ForEach-Object {$_.First;$_.Last+1})+@($found.Columns | ForEach-Object {$_.First;$_.Last+1})
  $bounds=@($bounds | Sort-Object -Unique)
  for($i=0;$i -lt $bounds.Count-1;$i++){
   $start=[int]$bounds[$i];$end=[int]$bounds[$i+1]-1
   $expected=@($item.Columns | Where-Object {$_.First -le $start -and $_.Last -ge $start})
   $observed=@($found.Columns | Where-Object {$_.First -le $start -and $_.Last -ge $start})
   $width=[double]$item.StandardWidth;$hidden=$false;if($expected.Count){$width=[double]$expected[0].Width;$hidden=[bool]$expected[0].Hidden}
   $got=[double]$found.StandardWidth;$gotHidden=$false;if($observed.Count){$got=[double]$observed[0].Width;$gotHidden=[bool]$observed[0].Hidden}
   Check-Layout ($item.Name+'.Columns.'+$start+':'+$end) ([Math]::Abs($width-$got) -lt 0.005 -and $gotHidden -eq $hidden) @($got,$gotHidden) @($width,$hidden)
  }
  $rows=@($item.Rows.Row)+@($found.Rows.Row) | Sort-Object -Unique
  $sectionHeaderRows=@();if($item.Name -eq 'Сечения'){foreach($name in $layout.Names | Where-Object Sheet -eq 'Сечения'){$sectionHeaderRows+=[int]$book.Worksheets.Item('Сечения').Range($name.Address).Row-1}}
  $expectedRows=@{};foreach($record in $item.Rows){$expectedRows[[int]$record.Row]=$record}
  $observedRows=@{};foreach($record in $found.Rows){$observedRows[[int]$record.Row]=$record}
  foreach($row in $rows){
   $expected=@();$observed=@();if($expectedRows.ContainsKey([int]$row)){$expected=@($expectedRows[[int]$row])};if($observedRows.ContainsKey([int]$row)){$observed=@($observedRows[[int]$row])}
   $height=[double]$item.StandardHeight;$hidden=$false;if($expected.Count){$height=[double]$expected[0].Height;$hidden=[bool]$expected[0].Hidden}
   if($row -in $sectionHeaderRows){$height=[double]$layout.SectionHeaders.Height}
   $got=[double]$found.StandardHeight;$gotHidden=$false;if($observed.Count){$got=[double]$observed[0].Height;$gotHidden=[bool]$observed[0].Hidden}
   Check-Layout ($item.Name+'.Row.'+$row) ([Math]::Abs($height-$got) -lt 0.01 -and $gotHidden -eq $hidden) @($got,$gotHidden) @($height,$hidden)
  }
  Check-Layout ($item.Name+'.ObjectNames') (($found.Objects.Name -join '|') -ceq ($item.Objects.Name -join '|')) ($found.Objects.Name -join '|') ($item.Objects.Name -join '|')
  foreach($object in $item.Objects){
   $got=@($found.Objects | Where-Object Name -eq $object.Name)[0]
   foreach($property in @('Left','Top','Width','Height')){Check-Layout ($item.Name+'.'+$object.Name+'.'+$property) ([Math]::Abs([double]$got.$property-[double]$object.$property) -lt 0.05) $got.$property $object.$property}
   Check-Layout ($item.Name+'.'+$object.Name+'.Placement') ($got.Placement -eq $object.Placement -and $got.Visible -eq $object.Visible -and $got.LockAspectRatio -eq $object.LockAspectRatio) @($got.Placement,$got.Visible,$got.LockAspectRatio) @($object.Placement,$object.Visible,$object.LockAspectRatio)
  }
  Check-Layout ($item.Name+'.ExcelTables') (($found.Tables | ConvertTo-Json -Compress) -ceq ($item.Tables | ConvertTo-Json -Compress)) $found.Tables $item.Tables
 }
 $headers=@()
 foreach($name in @('rngCircleGeometry','rngRectSetGeometry','rngPointedRectGeometry','rngHollowRectangleGeometry')){
  $table=$book.Names.Item($name).RefersToRange;$cell=$table.Worksheet.Cells.Item($table.Row-1,$table.Column)
  $text=[string]$cell.Value2;$start=$text.IndexOf('(Подробнее)')+1
  $ok=($cell.Interior.Color -eq $layout.SectionHeaders.Color -and $cell.RowHeight -eq $layout.SectionHeaders.Height -and $cell.Characters(1,1).Font.Size -eq $layout.SectionHeaders.FontSize -and $cell.Characters(1,1).Font.Bold)
  $ok=$ok -and $cell.Worksheet.Cells.Item($cell.Row,16384).Interior.Color -eq $layout.SectionHeaders.Color
  Check-Layout ('Header.'+$name) $ok @($cell.Interior.Color,$cell.RowHeight,$cell.Characters(1,1).Font.Size) @($layout.SectionHeaders.Color,$layout.SectionHeaders.Height,$layout.SectionHeaders.FontSize)
  Check-Layout ('HeaderLink.'+$name) ($start -gt 0 -and $cell.Hyperlinks.Count -eq 1 -and $cell.Characters($start,11).Font.Color -eq 16711680) $cell.Hyperlinks.Count 1
  $headers+=,[ordered]@{Name=$name;Row=[int]$cell.Row;Caption=$text}
 }
 $failed=@($checks | Where-Object {-not $_.Passed})
 [ordered]@{Status=if($failed.Count){'FAIL'}else{'PASS'};WorkbookSHA256=$hash;NamedRanges=$layout.Names.Count;OtherTables=$layout.OtherTables.Count;Objects=(($layout.Sheets | ForEach-Object {$_.Objects.Count} | Measure-Object -Sum).Sum);Checks=$checks.Count;Headers=$headers;Failures=$failed;Details=$checks} | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath (Join-Path $root $ReportPath) -Encoding UTF8
 $failed | Select-Object -First 12 | ConvertTo-Json -Depth 4
}finally{
 if($book){$book.Close($false)}
 if($app){$app.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
 if($owned){$p=Get-Process -Id $owned -ErrorAction SilentlyContinue;if($p -and $p.ProcessName -eq 'EXCEL' -and $p.StartTime.ToUniversalTime().Ticks -eq $ticks){Stop-Process -Id $owned -Force}}
}
if((Get-FileHash -LiteralPath $path).Hash -cne $hash){throw 'Source changed during read-only layout check.'}
Write-Output ('WORKBOOK_LAYOUT: checks='+$checks.Count+'; failures='+$failed.Count)
if($failed.Count){exit 1}
