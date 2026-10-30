# -------------------- ДЛЯ ТЕСТОВ --------------------
# Подтверждает, что после интеграции изменены только ожидаемая строка ошибки пустого снимка и тестовый модуль.
$ErrorActionPreference='Stop'
$p=$PSScriptRoot
& (Join-Path $p 'TestVerifySource.ps1') -ExportPath (Join-Path $p 'Native02/VBA_All_Code.txt') -ReportPath (Join-Path $p 'SourceEquality03.json')

# -------------------- ДЛЯ ТЕСТОВ --------------------
# Разбирает весь нативный экспорт по фактическим длинам CodeModule, а не по поисковым фрагментам.
function Read-TestExport([string]$Path){
 $lines=[IO.File]::ReadAllLines($Path,[Text.Encoding]::UTF8);$all=@{}
 for($i=0;$i-lt$lines.Length;$i++){
  if(-not$lines[$i].StartsWith('COMPONENT: ')){continue}
  $name=$lines[$i].Substring(11);$count=[int]$lines[$i+2].Substring(7)
  $body=if($count){$lines[($i+4)..($i+3+$count)]-join"`n"}else{''}
  $all[$name]=$body;$i=$i+3+$count
 };return $all
}
$old=Read-TestExport (Join-Path $p 'Native01/VBA_All_Code.txt')
$new=Read-TestExport (Join-Path $p 'Native02/VBA_All_Code.txt')
if($old.Count-ne181-or$new.Count-ne181){throw 'Exact181 inventories required'}
$changed=@()
foreach($name in $old.Keys){if([PostAuditVbaCanonicalText]::Normalize($old[$name])-cne[PostAuditVbaCanonicalText]::Normalize($new[$name])){$changed+=$name}}
if(($changed|Sort-Object)-join'|'-cne'modTestPlotConfig|modWorkbookCalculation'){throw ('Unexpected post-integration code changes: '+($changed-join','))}
$literal='"В Results нет выбранного состояния для схемы. Выполните расчет или проверьте ProfileId и Visualization.State."'
$replacement='WorkbookMessageParagraph('+ $literal + ', containerNotice)'
$expected=$old['modWorkbookCalculation'].Replace($literal,$replacement)
$expectedNormalized=[PostAuditVbaCanonicalText]::Normalize($expected).Trim()
$actualNormalized=[PostAuditVbaCanonicalText]::Normalize($new['modWorkbookCalculation']).Trim()
if($expected-ceq$old['modWorkbookCalculation']-or$expectedNormalized-cne$actualNormalized){
 [IO.File]::WriteAllText((Join-Path $p 'ExpectedDelta03.txt'),$expectedNormalized)
 [IO.File]::WriteAllText((Join-Path $p 'ActualDelta03.txt'),$actualNormalized)
 throw 'Production delta is not the one approved empty-snapshot message expression'
}
[ordered]@{Passed=$true;ComponentsCompared=181;ChangedComponents=$changed;ProductionChangeOnlyEmptySnapshotMessage=$true;Other179BodiesEqual=$true;IntegrationSHA=(Get-FileHash -LiteralPath (Join-Path $p 'Build02/RC_Section_NDM.xlsm')).Hash;FinalSHA=(Get-FileHash -LiteralPath (Join-Path $p 'Final03/RC_Section_NDM.xlsm')).Hash}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $p 'CodeDelta03.json') -Encoding UTF8
Get-Content -LiteralPath (Join-Path $p 'CodeDelta03.json') -Raw -Encoding UTF8
