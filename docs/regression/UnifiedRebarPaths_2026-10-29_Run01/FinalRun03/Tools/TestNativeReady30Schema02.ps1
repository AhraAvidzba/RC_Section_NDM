# -------------------- ДЛЯ ТЕСТОВ --------------------
# Только после полного PASS экспортирует точную готовую книгу read-only, без импорта, сохранения или запуска макросов.
param([Parameter(Mandatory=$true)][string]$LaunchDirectory)
$ErrorActionPreference='Stop'
Import-Module 'C:/Windows/System32/WindowsPowerShell/v1.0/Modules/Microsoft.PowerShell.Utility/Microsoft.PowerShell.Utility.psd1' -Global
$root='C:/Users/avidzba/Desktop/программы/В_РАБОТЕ!/RC_Section_NDM'
$launch=(Resolve-Path -LiteralPath (Join-Path $root $LaunchDirectory)).Path
$request=Get-Content -LiteralPath (Join-Path $launch 'LaunchRequest.json') -Raw -Encoding UTF8|ConvertFrom-Json
$complete=Get-Content -LiteralPath (Join-Path $root ($request.EvidenceDirectory+'/CompleteCurrentAcceptance.json')) -Raw -Encoding UTF8|ConvertFrom-Json
$closure=Get-Content -LiteralPath (Join-Path $launch 'AllOwnClosuresReviewed.json') -Raw -Encoding UTF8|ConvertFrom-Json
if($complete.Status-cne'PASS'-or$complete.SourceSHA256-cne$request.SourceSHA256-or$complete.CoreSuites-ne182-or$complete.FocusedSingleCases-ne122-or-not$closure.AllOwnIdentitiesTerminal-or-not$closure.PrimaryLogsReviewed-or$closure.SourceSHA256-cne$request.SourceSHA256-or$closure.CoreMacrosReviewed-ne182-or$closure.FocusedSingleCasesReviewed-ne122-or$closure.FocusedRecordCount-ne26){throw 'Complete current full PASS182/122/26 and own closures required; no native process started'}
$proof=Get-Content -LiteralPath (Join-Path $PSScriptRoot 'Ready30SchemaPreparation02.json') -Raw -Encoding UTF8|ConvertFrom-Json
if($proof.LaunchDirectory-cne$LaunchDirectory-or$proof.AcceptedMainSHA256-cne$request.SourceSHA256-or$proof.AcceptedCommit-cne$request.SourceCommit-or$proof.OnlyVBAChanged-or-not$proof.ResultsPartByteExact-or-not$proof.CalculationSheetPartByteExact-or-not$proof.OriginalSharedStringsByteExact){throw 'Preparation does not match this accepted full source'}
if((Get-FileHash -LiteralPath (Join-Path $root 'workbook/output/RC_Section_NDM.xlsm')).Hash-cne$request.SourceSHA256){throw 'Accepted main package changed'}
if((Get-FileHash -LiteralPath $proof.Source).Hash-cne$proof.SourceSHA256){throw 'Original Ready30 changed; no native process started'}
$path=[string]$proof.Ready
if((Get-FileHash -LiteralPath $path).Hash-cne$proof.ReadySHA256){throw 'Prepared ready package changed'}
. (Join-Path $root 'tools/build_workbook/GeometryTables.ps1')
$scratch=$PSScriptRoot
$out=Join-Path $scratch 'Ready30SchemaNative02';if(Test-Path -LiteralPath $out){throw 'Fresh readonly evidence required'};[void][IO.Directory]::CreateDirectory($out)
$memory=Get-CimInstance Win32_PerfFormattedData_PerfOS_Memory
if($memory.AvailableMBytes-lt8192-or($memory.CommitLimit-$memory.CommittedBytes)-lt14GB){throw 'RAM/commit reserve insufficient'}
Add-Type 'using System;using System.Runtime.InteropServices;public static class TestReadyThirtySchemaOwner02{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);}'
$app=$null;$book=$null;$project=$null;$component=$null;$code=$null;$range=$null;$ownedNew=$false;$quitRequested=$false
# -------------------- ДЛЯ ТЕСТОВ --------------------
# Проверяет каждый настоящий вариант отдельно; склеенные строки обоих видов и неизвестное значение обязаны отклоняться.
function Test-ReadyDropdown([string]$SettingKey,[object]$TargetCell,[string[]]$AllowedChoices){
 $savedFormula=$TargetCell.Formula
 try{
  if($TargetCell.Validation.Type-ne3-or-not$TargetCell.Validation.InCellDropdown){throw ('Saved dropdown absent: '+$SettingKey)}
  if($AllowedChoices.Count-lt2-or@($AllowedChoices|Sort-Object -Unique).Count-ne$AllowedChoices.Count){throw 'Distinct explicit choices required'}
  foreach($choice in $AllowedChoices){$TargetCell.Value2=$choice;if(-not$TargetCell.Validation.Value){throw ('Saved choice rejected: '+$SettingKey+' = '+$choice)}}
  foreach($invalid in @(($AllowedChoices-join','),($AllowedChoices-join';'),'__INVALID_READY_CHOICE__')){
   $TargetCell.Value2=$invalid
   if($TargetCell.Validation.Value){throw ('Joined or unknown choice accepted: '+$SettingKey+' = '+$invalid)}
  }
  return [ordered]@{Key=$SettingKey;Address=$TargetCell.Address();NativeFormula1=[string]$TargetCell.Validation.Formula1;Options=$AllowedChoices;EachOptionAccepted=$true;JoinedCommaRejected=$true;JoinedSemicolonRejected=$true;UnknownRejected=$true;SavedPackageAlreadyReopened=$true;SourceSeparatorNotNormalized=$true}
 }finally{
  if($null-eq$savedFormula){[void]$TargetCell.ClearContents()}else{$TargetCell.Formula=$savedFormula}
 }
}
try{
 $existing=@([Diagnostics.Process]::GetProcessesByName('EXCEL')|ForEach-Object {$_.Id})
 $app=New-Object -ComObject Excel.Application
 [uint32]$id=0;[void][TestReadyThirtySchemaOwner02]::GetWindowThreadProcessId([IntPtr]$app.Hwnd,[ref]$id);$owned=[int]$id
 if($owned-le0-or$existing-contains$owned){$app=$null;throw 'New own Excel identity required'}
 $ownedNew=$true;$ticks=(Get-Process -Id $owned).StartTime.ToUniversalTime().Ticks
 [ordered]@{PID=$owned;StartTicksUTC=[string]$ticks;RAMGiB=$memory.AvailableMBytes/1024;CommitFreeGiB=($memory.CommitLimit-$memory.CommittedBytes)/1GB}|ConvertTo-Json|Set-Content (Join-Path $out 'Ownership.json') -Encoding UTF8
 $app.Visible=$false;$app.DisplayAlerts=$false;$app.EnableEvents=$false;$app.AutomationSecurity=3
 $book=$app.Workbooks.Open($path,0,$true);if(-not$book.ReadOnly){throw 'Readonly required'}
 $project=$book.VBProject;if($project.VBComponents.Count-ne181){throw '181 native components required'}
 $all=New-Object Text.StringBuilder;[void]$all.AppendLine('VBA PROJECT EXPORT');[void]$all.AppendLine('Workbook: '+$book.Name);[void]$all.AppendLine('Date: '+(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
 $components=@()
 for($i=1;$i-le$project.VBComponents.Count;$i++){
  $component=$project.VBComponents.Item($i);$code=$component.CodeModule;$n=[int]$code.CountOfLines;$body=if($n){[string]$code.Lines(1,$n)}else{''}
  [IO.File]::WriteAllText((Join-Path $out ($component.Name+'.txt')),$body,[Text.Encoding]::UTF8)
  [void]$all.AppendLine('='*100);[void]$all.AppendLine('COMPONENT: '+$component.Name);[void]$all.AppendLine('TYPE: '+$component.Type);[void]$all.AppendLine('LINES: '+$n);[void]$all.AppendLine('='*100);if($n){[void]$all.AppendLine($body)};[void]$all.AppendLine()
  $components+=,[ordered]@{Name=[string]$component.Name;Type=[int]$component.Type;Lines=$n}
 }
 [IO.File]::WriteAllText((Join-Path $out 'VBA_All_Code.txt'),$all.ToString(),[Text.Encoding]::UTF8)
 $range=$book.Names.Item('rngNDMSectionGeometry').RefersToRange;$anchor=$range.Cells.Item(1,1);$sheet=$range.Worksheet
 $last=$sheet.Cells.Item($sheet.Rows.Count,$anchor.Column+1).End(-4162).Row
 $geometry=$anchor.Resize($last-$anchor.Row+1,16).Value2;$concrete=0;$rebar=0
 for($r=2;$r-le$geometry.GetLength(0);$r++){
  switch -CaseSensitive ([string]$geometry[$r,3]){'Concrete'{$concrete++};'Rebar'{$rebar++};default{throw 'Unknown saved element role'}}
 }
 if($concrete-ne1232-or$rebar-ne8){throw ('Expected real heavy 1232/8; actual='+$concrete+'/'+$rebar)}
 $range=$book.Names.Item('rngLoadCombinations').RefersToRange;$values=$range.Value2;$loads=@()
 for($r=2;$r-le$values.GetLength(0);$r++){if([string]$values[$r,1]){$loads+=[string]$values[$r,1]}}
 $expected=@(1..30|ForEach-Object {'PERF_'+('{0:00}'-f$_)})
 if($loads.Count-ne30-or($loads-join'|')-cne($expected-join'|')){throw 'Original thirty load IDs required'}
 $range=$book.Names.Item('rngBatchSummary').RefersToRange;$summary=$range.Cells.Item(1,1).Resize(50,33).Value2;$inputErrors=@()
 for($r=1;$r-le50;$r++){if([string]$summary[$r,1]-match'^PERF_\d+$'-and[string]$summary[$r,4]-ceq'InputErr'){$inputErrors+=[string]$summary[$r,1]}}
 if(($inputErrors-join'|')-cne'PERF_14|PERF_21|PERF_27|PERF_28'){throw 'Exactly the original four historical InputErr IDs must remain unchanged'}
 $range=$book.Names.Item('rngSystemSettings').RefersToRange;$settings=$range.Value2;$selected=@{}
 for($r=2;$r-le$settings.GetLength(0);$r++){if([string]$settings[$r,1]-in@('Results.SnapshotStorage','Results.SnapshotName','Output.CombinationMode','Output.CombinationID','Mesh.CellGeometry')){$selected[[string]$settings[$r,1]]=$settings[$r,2]}}
 if($selected['Results.SnapshotStorage']-cne'Excel'-or$selected['Output.CombinationMode']-cne'Worst'){throw 'Expected ready Excel snapshot and approved Worst selection'}
 $validationChoices=@{
  'Geometry.Type'=@('Circle','Rectangles','PointedSection','HollowRectangle')
  'Mesh.CellGeometry'=@('Adaptive','Contour')
  'Mesh.StepMode'=@('Auto 0.05','Auto 0.1','Auto 0.2','User')
  'Results.SnapshotStorage'=@('Excel','File')
  'Output.CombinationMode'=@('Worst','User')
 }
 $validationAudit=@()
 foreach($settingKey in @('Geometry.Type','Mesh.CellGeometry','Mesh.StepMode','Results.SnapshotStorage','Output.CombinationMode')){
  $validationCell=$null
  for($r=2;$r-le$settings.GetLength(0);$r++){if([string]$settings[$r,1]-ceq$settingKey){$validationCell=$range.Cells.Item($r,2);break}}
  if($null-eq$validationCell){throw ('Required setting absent: '+$settingKey)}
  $validationAudit+=,(Test-ReadyDropdown $settingKey $validationCell $validationChoices[$settingKey])
 }
 $validationCell=$book.Names.Item('rngLoadCombinations').RefersToRange.Cells.Item(2,5)
 $profileFormula=[string]$validationCell.Validation.Formula1
 $profileRange=$book.Worksheets.Item('Config').Range($profileFormula.TrimStart('='))
 if($profileRange.Rows.Count-ne1-or$profileRange.Columns.Count-ne4){throw 'Current direct visible four-profile row required'}
 $profileChoices=@($profileRange.Value2|ForEach-Object{[string]$_})
 $validationAudit+=,(Test-ReadyDropdown 'Load.ProfileID' $validationCell $profileChoices)

 # ДЛЯ ТЕСТОВ: проверка сохранённых Main-списков и неизменных внешних габаритов всех таблиц.
 $modeAudit=@();$modeAccepted=0;$modeRejected=0;$projected=0
 $layout=Get-Content -LiteralPath (Join-Path $root 'tools/build_workbook/WorkbookLayout.json') -Raw -Encoding UTF8|ConvertFrom-Json
 foreach($definition in @(Get-GeometrySettingsCatalog)){
  $table=$book.Names.Item($definition.RangeName).RefersToRange
  $layoutName=@($layout.Names|Where-Object Name -eq $definition.RangeName)[0]
  if($table.Address()-cne$layoutName.Address-or$table.Rows.Count-ne$layoutName.Rows-or$table.Columns.Count-ne$layoutName.Columns){throw ('Ready Main outer table bounds changed: '+$definition.RangeName)}
  $block=@(Get-GeometryTableBlocks $table|Where-Object Key -eq 'Main')[0]
  if([int]$block.Columns.endMode-ne8-or[int]$block.Columns.note-ne9){throw 'Ready Main field order differs'}
  foreach($row in @(($definition.Blocks|Where-Object Key -eq 'Main').Rows)){
   $key=[string]$row[0];$cell=Get-GeometryTableCell $table 'Main' $key 'endMode';$note=Get-GeometryTableCell $table 'Main' $key 'note'
   if([string]$cell.Value2-cne[string]$row[7]){throw ('Ready default must preserve placement: '+$definition.Type+'/'+$key)}
   if($cell.HorizontalAlignment-ne-4108-or$cell.WrapText-or$note.HorizontalAlignment-ne-4152-or$note.WrapText){throw 'Ready Main alignment/wrapping changed'}
   if($key-like'Opening*'){
    if([string]$cell.Value2-cne[char]0x2014-or[int]$cell.Interior.Color-eq15135487){throw 'Projection must remain a non-input dash'}
    $projected++
   }else{
    if([int]$cell.Interior.Color-ne15135487-or-not$cell.Validation.ShowError){throw 'Editable Main list has wrong fill or validation'}
    $audit=Test-ReadyDropdown ($definition.Type+'/'+$key+'/endMode') $cell @('IncludeBoth','ExcludeBoth','IncludeStart','IncludeEnd')
    $saved=$cell.Formula
    try{
     foreach($bad in @('Auto','ClosedUniform')){$cell.Value2=$bad;if($cell.Validation.Value){throw ('Removed mode accepted: '+$bad)};$modeRejected++}
    }finally{$cell.Formula=$saved}
    $modeAccepted+=4;$modeRejected+=3;$modeAudit+=,$audit
   }
   foreach($edge in 7..10){if($cell.Borders.Item($edge).LineStyle-eq-4142-or$note.Borders.Item($edge).LineStyle-eq-4142){throw 'Ready Main border missing'}}
  }
 }
 if($modeAccepted-ne68-or$modeRejected-ne85-or$projected-ne4){throw 'Ready Main native list inventory differs'}
 $modeAudit|ConvertTo-Json -Depth 8|Set-Content (Join-Path $out 'MainDropdownAcceptance.json') -Encoding UTF8
 # ДЛЯ ТЕСТОВ: общие объяснения не дублируются в разделах типов; все ссылки ведут в один блок.
 $help=$book.Worksheets.Item('Справка');$used=$help.UsedRange;$text=$used.Value2
 $common=@();for($r=1;$r-le$text.GetLength(0);$r++){if([string]$text[$r,2]-like'Поле «Расстановка» (endMode)*IncludeBoth*ExcludeBoth*IncludeStart*IncludeEnd*'){$common+=($used.Row+$r-1)}}
 if($common.Count-ne1){throw 'Exactly one common EndMode description required'}
 $targets=@();foreach($link in $help.Hyperlinks){if([string]$link.SubAddress-ceq"'Справка'!A711"){$targets+=[string]$link.Range.Address()}}
 if($targets.Count-ne4){throw 'Four type links to shared reinforcement help required'}
 [ordered]@{CommonModeDescriptionRow=$common[0];SharedTargets=$targets;MainAcceptedChoices=$modeAccepted;MainRejectedChoices=$modeRejected;ProjectionDashes=$projected;SchemaPreparationProof='Ready30SchemaPreparation02.json';HistoricalResultsPreserved=$true}|ConvertTo-Json -Depth 5|Set-Content (Join-Path $out 'SchemaNativeAcceptance.json') -Encoding UTF8
 $help=$null;$used=$null;$table=$null;$cell=$null;$note=$null;$link=$null
 $validationCell=$null;$profileRange=$null
 [ordered]@{Status='PASS_SAVED_READY_NATIVE_DROPDOWNS';Choices=$validationAudit;OriginalValuesRestoredInMemory=$true;WorkbookSaved=$false;SourceFileUnchanged=$true}|ConvertTo-Json -Depth 8|Set-Content (Join-Path $out 'DropdownAcceptance.json') -Encoding UTF8
 [ordered]@{Status='PASS_READONLY_READY_CODE_AND_INPUT_INVENTORY';SHA256=$proof.ReadySHA256;AcceptedMainSHA256=$request.SourceSHA256;ReadOnly=$true;MacrosImportedRunSaved=$false;Components=$components;Concrete=$concrete;Rebar=$rebar;LoadIDs=$loads;HistoricalInputErrIDs=$inputErrors;Settings=$selected;SchemaPreparationProof='Ready30SchemaPreparation02.json';HistoricalResultsPreserved=$true}|ConvertTo-Json -Depth 8|Set-Content (Join-Path $out 'NativeVerification.json') -Encoding UTF8
 Write-Output 'READY_NATIVE_READONLY: 181 components; 1232/8; 30 loads; four historical InputErr preserved'
}catch{
 [ordered]@{Error=$_.Exception.ToString();Stack=$_.ScriptStackTrace}|ConvertTo-Json|Set-Content (Join-Path $out 'Failure.json') -Encoding UTF8
 throw
}finally{
 $validationCell=$null;$profileRange=$null;$sheet=$null;$anchor=$null;$range=$null;$code=$null;$component=$null;$project=$null;[GC]::Collect();[GC]::WaitForPendingFinalizers()
 if($book){$book.Close($false)};$book=$null;[GC]::Collect();[GC]::WaitForPendingFinalizers()
 if($app-and$ownedNew){
  $ownProcess=Get-Process -Id $owned -ErrorAction Stop
  if([string]$ownProcess.StartTime.ToUniversalTime().Ticks-cne[string]$ticks-or$ownProcess.ProcessName-ine'EXCEL'){throw 'Own Excel identity changed; no Quit performed'}
  $app.Quit();$quitRequested=$true
 };$app=$null;[GC]::Collect();[GC]::WaitForPendingFinalizers();[GC]::Collect();[GC]::WaitForPendingFinalizers()
 [ordered]@{PID=$owned;StartTicksUTC=[string]$ticks;QuitRequested=$quitRequested;ForceKilled=$false;NewOwnInstanceVerified=$ownedNew}|ConvertTo-Json|Set-Content (Join-Path $out 'Quit.json') -Encoding UTF8

 $process=$null;try{$process=Get-Process -Id $owned -ErrorAction Stop}catch{if($_.FullyQualifiedErrorId-notlike'NoProcessFoundForGivenId*'){throw}}
 if($process-and[string]$process.StartTime.ToUniversalTime().Ticks-ceq[string]$ticks){[void]$process.WaitForExit(10000)}
 $process=$null;try{$process=Get-Process -Id $owned -ErrorAction Stop}catch{if($_.FullyQualifiedErrorId-notlike'NoProcessFoundForGivenId*'){throw}}
 if($process-and[string]$process.StartTime.ToUniversalTime().Ticks-ceq[string]$ticks){throw 'Exact own Ready Excel still live after normal Quit'}
 [ordered]@{PID=$owned;StartTicksUTC=[string]$ticks;ExactIdentityTerminal=$true;QuitRequested=$quitRequested;ForceKilled=$false}|ConvertTo-Json|Set-Content (Join-Path $out 'OwnClosure.json') -Encoding UTF8
 if((Get-FileHash -LiteralPath $path).Hash-cne$proof.ReadySHA256){throw 'Readonly ready package changed'}
 if((Get-FileHash -LiteralPath $proof.Source).Hash-cne$proof.SourceSHA256){throw 'Original Ready30 changed'}
}
