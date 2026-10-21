# Одноразово добавляет новую настройку и её справку в проверочную копию книги.
# Существующие значения, правые таблицы Config и пользовательские листы сохраняются.
param([Parameter(Mandatory=$true)][string]$WorkbookPath,[switch]$RepairValidation)
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../../..')).Path
. (Join-Path $root 'tools/build_workbook/SettingsCatalog.ps1')
$path=(Resolve-Path -LiteralPath $WorkbookPath).Path
if(-not $path.StartsWith((Join-Path $root 'docs/regression/ResultsFormatting_2026-10-21'),[StringComparison]::OrdinalIgnoreCase)){throw 'Only the private candidate may be updated.'}
$key='General.ReformatResultsEnabled'
$catalog=@(Get-SystemSettingsCatalog)
$record=@($catalog[0].Rows | Where-Object {$_[0] -eq $key})[0]
if(-not $record){throw 'Missing setting catalog entry.'}
$excel=$null;$book=$null
try {
 $excel=New-Object -ComObject Excel.Application
 $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false;$excel.AutomationSecurity=3
 $book=$excel.Workbooks.Open($path)
 $config=$book.Worksheets.Item('Config');$help=$book.Worksheets.Item('Справка')
 $range=$book.Names.Item('rngSystemSettings').RefersToRange;$values=$range.Value2;$row=0
 if($RepairValidation){
  for($r=1;$r -le $values.GetLength(0);$r++){
   if([string]$values[$r,1] -eq $key){
    $cell=$range.Cells.Item($r,2);$cell.Validation.Delete();$cell.Validation.Add(3,1,1,('Yes'+[string]$excel.International(5)+'No'))
    $book.Save();Write-Output 'VALIDATION_REPAIRED';return
   }
  }
  throw 'Setting to repair was not found.'
 }
 for($r=1;$r -le $values.GetLength(0);$r++){
  if([string]$values[$r,1] -eq $key){throw 'Setting already exists.'}
  if([string]$values[$r,1] -eq 'General.NonCriticalMessagesEnabled'){$row=[int]$range.Row+$r}
 }
 if(-not $row){throw 'Missing General insertion point.'}
 $links=New-Object Collections.Generic.List[object]
 foreach($sheet in $book.Worksheets){foreach($link in $sheet.Hyperlinks){$links.Add(@{Link=$link;Before=[string]$link.SubAddress})}}
 $config.Cells.Item($row,1).Resize(1,5).Insert(-4121) | Out-Null
 $config.Cells.Item(($row-1),1).Resize(1,5).Copy($config.Cells.Item($row,1))
 $data=New-Object 'object[,]' 1,5
 for($c=0;$c -lt 4;$c++){$data[0,$c]=[string]$record[$c]}
 $data[0,4]='Подробнее'
 $config.Cells.Item($row,1).Resize(1,5).Value2=$data
 $cell=$config.Cells.Item($row,2)
 $cell.Validation.Delete();$cell.Validation.Add(3,1,1,('Yes'+[string]$excel.International(5)+'No'))
 $cell.Interior.Color=15135487;$cell.HorizontalAlignment=-4108
 foreach($item in $links){
  $before=$item.Before
  if($before -match "^'?Config'?!(\`$?[A-E])(\`$?)(\d+)$" -and [int]$Matches[3] -ge $row){
   if([string]$item.Link.SubAddress -ceq $before){$item.Link.SubAddress="'Config'!"+$Matches[1]+$Matches[2]+([int]$Matches[3]+1)}
  }
 }
 $last=$help.Cells.Find('*',$help.Cells.Item(1,1),-4123,2,1,2,$false)
 $guideRow=[int]$last.Row+2;$next=Add-GuideTitle $help $guideRow ($key+' — повторное оформление Results') 11
 foreach($line in @(Get-SettingInstructionLines $key ([string]$record[3]))){$next=Add-GuideParagraph $help $next $line}
 $guide=$help.Cells.Item($guideRow,1).Resize(($next-$guideRow),6)
 $guide.Font.Name='Arial';$guide.WrapText=$false;$guide.RowHeight=15
 [void]$help.Hyperlinks.Add($help.Cells.Item($guideRow,7),'',("'Config'!B"+$row),'','К настройке')
 $config.Cells.Item($row,5).Hyperlinks.Delete()
 [void]$config.Hyperlinks.Add($config.Cells.Item($row,5),'',("'Справка'!A"+$guideRow),'','Подробнее')
 $config.Cells.Item($row,5).Font.Color=16711680;$config.Cells.Item($row,5).Font.Underline=2
 $book.Save()
 [ordered]@{Added=$key;Default='Yes';ConfigRow=$row;HelpRow=$guideRow;PrivateWorkbook=$path} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path (Split-Path -Parent $path) 'Setting.json') -Encoding UTF8
} finally {
 if($book){$book.Close($false)}
 if($excel){$excel.Quit()}
}
Write-Output 'SETTING_ADDED'
