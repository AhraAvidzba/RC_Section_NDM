# Переносит пользовательский ввод в заново собранную книгу в каталоге проверки.
# Одноразовое преобразование удалённых полей не входит в рабочий VBA.
param([Parameter(Mandatory=$true)][string]$BeforeWorkbook,
      [Parameter(Mandatory=$true)][string]$RebuiltWorkbook,
      [Parameter(Mandatory=$true)][string]$ReportPath)
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
. (Join-Path $PSScriptRoot 'GeometryTables.ps1')
$before=(Resolve-Path -LiteralPath $BeforeWorkbook).Path
$rebuilt=(Resolve-Path -LiteralPath $RebuiltWorkbook).Path
$allowed=[IO.Path]::GetFullPath((Join-Path $root 'docs/regression/SharedRebarInput_2026-10-18/SectionNames_2026-10-20'))+[IO.Path]::DirectorySeparatorChar
if(-not $before.StartsWith($allowed,[StringComparison]::OrdinalIgnoreCase) -or -not $rebuilt.StartsWith($allowed,[StringComparison]::OrdinalIgnoreCase)){throw 'Only isolated section-name evidence books are allowed.'}
$oldHash=(Get-FileHash -LiteralPath $before).Hash
Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public static class SectionReleaseWindow{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr h,out uint id);}'
$app=$null;$old=$null;$book=$null;$owned=0;$ticks=0;$fields=@();$configCount=0
try{
 $app=New-Object -ComObject Excel.Application
 $app.Visible=$false;$app.DisplayAlerts=$false;$app.EnableEvents=$false;$app.AutomationSecurity=3
 [uint32]$id=0;[void][SectionReleaseWindow]::GetWindowThreadProcessId([IntPtr]$app.Hwnd,[ref]$id)
 $owned=[int]$id;$ticks=(Get-Process -Id $owned).StartTime.ToUniversalTime().Ticks
 Write-Output "OWNED_EXCEL: id=$owned; startTicks=$ticks"
 $old=$app.Workbooks.Open($before,0,$true);$book=$app.Workbooks.Open($rebuilt,0,$false);$app.Calculation=-4135
 foreach($sheetName in @('Config','Расчет')){
  $source=$old.Worksheets.Item($sheetName);$target=$book.Worksheets.Item($sheetName)
  foreach($cell in $source.UsedRange.Cells){
   if([int]$cell.Interior.Color -ne 15135487){continue}
   $to=$target.Cells.Item([int]$cell.Row,[int]$cell.Column)
   $formula=$cell.Formula
   if($formula -is [string]){$formula=$formula.Replace('PointedRect','PointedSection').Replace('RectSet','Rectangles')}
   if($formula -is [string]){$to.Formula=$formula}else{$to.Value2=[double]$formula};$configCount++
  }
 }
 foreach($definition in @(Get-GeometrySettingsCatalog)){
  $oldName=$definition.RangeName.Replace('PointedSection','PointedRect').Replace('Rectangles','RectSet')
  $from=$old.Names.Item($oldName).RefersToRange;$to=$book.Names.Item($definition.RangeName).RefersToRange
  foreach($block in $definition.Blocks){foreach($row in $block.Rows){for($c=0;$c -lt $block.Keys.Count;$c++){
   $key=[string]$block.Keys[$c]
   if($key -in @('key','name','unit','note') -or [string]$row[$c] -in @('-','—')){continue}
   $source=Get-GeometryTableCell $from $block.Key $row[0] $key
   $target=Get-GeometryTableCell $to $block.Key $row[0] $key
   $value=$source.Formula
   if($definition.Type -eq 'Rectangles' -and $block.Key -eq 'General'){
    $oldShape=[string](Get-GeometryTableCell $from 'General' 'Shape' 'val').Value2
    if($row[0] -eq 'Shape'){
     if($oldShape -eq 'Rectangle'){$value='OneRect'}else{$value='TwoRect'}
    }elseif($row[0] -eq 'XOffset' -and $oldShape -in @('LSection','Rectangles')){$value=0}
   }
   if($definition.Type -eq 'PointedSection' -and $block.Key -eq 'General' -and $row[0] -like '*.W'){
    $type=Get-GeometryTableCell $from 'General' ($row[0].Replace('.W','.Type')) 'val'
    if([string]$type.Value2 -eq 'Simple'){$value=0}
   }
   if($value -is [string]){$target.Formula=$value}else{$target.Value2=[double]$value}
   $fields+=,[ordered]@{Table=$definition.RangeName;Block=$block.Key;Row=$row[0];Column=$key;Value=$target.Value2;Formula=$target.Formula;Address=($target.Worksheet.Name+'!'+$target.Address($false,$false))}
  }}}
 }
 $book.Worksheets.Item('Config').Calculate();$book.Worksheets.Item('Сечения').Calculate()
 $book.Worksheets.Item('Расчет').Activate();$book.Save();$book.Close($false);$book=$null
 $old.Close($false);$old=$null
 if((Get-FileHash -LiteralPath $before).Hash -cne $oldHash){throw 'Original input snapshot changed.'}
 [ordered]@{Passed=$true;BeforeSHA256=$oldHash;RebuiltSHA256=(Get-FileHash -LiteralPath $rebuilt).Hash;ConfigAndCalculationCells=$configCount;InputFields=$fields;RemovedInputs=@('Left.Type','Right.Type');InactiveWidthsNormalizedToZero=$true;RectanglesChoices=@('TwoRect','OneRect');FormerLSectionOffsetNormalizedToZero=$true} | ConvertTo-Json -Depth 7 | Set-Content -LiteralPath $ReportPath -Encoding UTF8
 Write-Output "INPUTS_PRESERVED: config/calculation=$configCount; geometry=$($fields.Count)"
}catch{
 Write-Output $_.ScriptStackTrace
 Write-Output $_.InvocationInfo.PositionMessage
 throw
}finally{
 if($book){$book.Close($false)};if($old){$old.Close($false)}
 if($app){$app.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
 if($owned){$p=Get-Process -Id $owned -ErrorAction SilentlyContinue;if($p -and $p.ProcessName -eq 'EXCEL' -and $p.StartTime.ToUniversalTime().Ticks -eq $ticks){Stop-Process -Id $owned -Force}}
}
