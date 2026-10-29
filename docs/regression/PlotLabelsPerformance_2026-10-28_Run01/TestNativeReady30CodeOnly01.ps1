# -------------------- ДЛЯ ТЕСТОВ --------------------
# Только после полного PASS экспортирует точную готовую книгу read-only, без импорта, сохранения или запуска макросов.
param([Parameter(Mandatory=$true)][string]$LaunchDirectory)
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../../..')).Path
$launch=(Resolve-Path -LiteralPath (Join-Path $root $LaunchDirectory)).Path
$request=Get-Content -LiteralPath (Join-Path $launch 'LaunchRequest.json') -Raw -Encoding UTF8|ConvertFrom-Json
$complete=Get-Content -LiteralPath (Join-Path $root ($request.EvidenceDirectory+'/CompleteCurrentAcceptance.json')) -Raw -Encoding UTF8|ConvertFrom-Json
$closure=Get-Content -LiteralPath (Join-Path $launch 'AllOwnClosuresReviewed.json') -Raw -Encoding UTF8|ConvertFrom-Json
if($complete.Status-cne'PASS'-or$complete.SourceSHA256-cne$request.SourceSHA256-or-not$closure.AllOwnIdentitiesTerminal){throw 'Complete current full PASS and own closures required; no native process started'}
$proof=Get-Content -LiteralPath (Join-Path $PSScriptRoot 'Ready30CodePreparation01.json') -Raw -Encoding UTF8|ConvertFrom-Json
$path=[string]$proof.Ready
if((Get-FileHash -LiteralPath $path).Hash-cne$proof.ReadySHA256){throw 'Prepared ready package changed'}
$scratch='C:/Users/avidzba/.codex/visualizations/2026/10/12/01a12ece-8192-7141-89b3-f83f1176184c/PlotLabelsPerformance_2026-10-28_Run01'
$out=Join-Path $scratch 'Ready30Native01';if(Test-Path -LiteralPath $out){throw 'Fresh readonly evidence required'};[void][IO.Directory]::CreateDirectory($out)
$memory=Get-CimInstance Win32_PerfFormattedData_PerfOS_Memory
if($memory.AvailableMBytes-lt8192-or($memory.CommitLimit-$memory.CommittedBytes)-lt14GB){throw 'RAM/commit reserve insufficient'}
Add-Type 'using System;using System.Runtime.InteropServices;public static class TestReadyThirtyOwner{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);}'
$app=$null;$book=$null;$project=$null;$component=$null;$code=$null;$range=$null;$ownedNew=$false
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
 [uint32]$id=0;[void][TestReadyThirtyOwner]::GetWindowThreadProcessId([IntPtr]$app.Hwnd,[ref]$id);$owned=[int]$id
 if($existing-contains$owned){$app=$null;throw 'New own Excel identity required'}
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
 if($inputErrors.Count-ne4){throw 'Exactly four historical InputErr results must remain unchanged'}
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
 $validationCell=$null;$profileRange=$null
 [ordered]@{Status='PASS_SAVED_READY_NATIVE_DROPDOWNS';Choices=$validationAudit;OriginalValuesRestoredInMemory=$true;WorkbookSaved=$false;SourceFileUnchanged=$true}|ConvertTo-Json -Depth 8|Set-Content (Join-Path $out 'DropdownAcceptance.json') -Encoding UTF8
 [ordered]@{Status='PASS_READONLY_READY_CODE_AND_INPUT_INVENTORY';SHA256=$proof.ReadySHA256;AcceptedMainSHA256=$request.SourceSHA256;ReadOnly=$true;MacrosImportedRunSaved=$false;Components=$components;Concrete=$concrete;Rebar=$rebar;LoadIDs=$loads;HistoricalInputErrIDs=$inputErrors;Settings=$selected;OnlyVBAChangedProof='Ready30CodePreparation01.json'}|ConvertTo-Json -Depth 8|Set-Content (Join-Path $out 'NativeVerification.json') -Encoding UTF8
 Write-Output 'READY_NATIVE_READONLY: 181 components; 1232/8; 30 loads; four historical InputErr preserved'
}catch{
 [ordered]@{Error=$_.Exception.ToString();Stack=$_.ScriptStackTrace}|ConvertTo-Json|Set-Content (Join-Path $out 'Failure.json') -Encoding UTF8
 throw
}finally{
 $validationCell=$null;$profileRange=$null;$sheet=$null;$anchor=$null;$range=$null;$code=$null;$component=$null;$project=$null;[GC]::Collect();[GC]::WaitForPendingFinalizers()
 if($book){$book.Close($false)};$book=$null;[GC]::Collect();[GC]::WaitForPendingFinalizers()
 if($app-and$ownedNew){$app.Quit()};$app=$null;[GC]::Collect();[GC]::WaitForPendingFinalizers();[GC]::Collect();[GC]::WaitForPendingFinalizers()
 [ordered]@{PID=$owned;StartTicksUTC=[string]$ticks;QuitRequested=$true;ForceKilled=$false}|ConvertTo-Json|Set-Content (Join-Path $out 'Quit.json') -Encoding UTF8
 if((Get-FileHash -LiteralPath $path).Hash-cne$proof.ReadySHA256){throw 'Readonly ready package changed'}
}
