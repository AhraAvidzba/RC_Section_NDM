# -------------------- ДЛЯ ТЕСТОВ --------------------
# Проверяет пустой и заполненный прямой список ID, а также сохранность отказа повреждённых источников.
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../../../..')).Path
$manifest=Get-Content -LiteralPath (Join-Path $PSScriptRoot 'DeferredSourceManifest.json') -Raw -Encoding UTF8|ConvertFrom-Json
$source=Join-Path $root $manifest.Source;$candidate=Join-Path $root $manifest.Candidate
if((Get-FileHash -LiteralPath $source).Hash-cne$manifest.OriginalSHA256-or(Get-FileHash -LiteralPath $candidate).Hash-cne$manifest.CandidateSHA256){throw 'Source/candidate mismatch'}
$closed=Get-Content -LiteralPath (Join-Path $root 'docs/regression/FullSuiteLaunch_2026-10-28_PlotLabels_Run01/AllOwnClosuresReviewed.json') -Raw -Encoding UTF8|ConvertFrom-Json
if(-not$closed.AllOwnIdentitiesTerminal){throw 'First full own closure proof required'}
$dir='C:/Users/avidzba/.codex/visualizations/2026/10/12/01a12ece-8192-7141-89b3-f83f1176184c/PlotLabelsPerformance_2026-10-28_Run01/ParallelFix04/NativeDynamic02'
if(Test-Path -LiteralPath $dir){throw 'Fresh evidence required'};[void][IO.Directory]::CreateDirectory($dir)
$path=Join-Path $dir 'RC_Section_NDM.xlsm'
Copy-Item -LiteralPath (Join-Path $root 'docs/regression/FullOnOff_2026-10-28_PlotLabels_Run01/Source/RC_Section_NDM.xlsm') -Destination $path
$memory=Get-CimInstance Win32_PerfFormattedData_PerfOS_Memory
if($memory.AvailableMBytes-lt8192-or($memory.CommitLimit-$memory.CommittedBytes)-lt14GB){throw 'RAM/commit reserve insufficient'}
Add-Type 'using System;using System.Runtime.InteropServices;public static class TestDynamicIdOwner{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);}'
$app=$null;$book=$null;$project=$null;$code=$null;$compile=$null;$range=$null;$idCell=$null;$profileCell=$null;$ownedNew=$false;$records=@()
# -------------------- ДЛЯ ТЕСТОВ --------------------
# Находит настройку по ключу в текущем именованном диапазоне без фиксированного адреса.
function Test-SettingCell([string]$Key){
 $area=$book.Names.Item('rngSystemSettings').RefersToRange;$data=$area.Value2
 for($r=2;$r-le$data.GetLength(0);$r++){if([string]$data[$r,1]-ceq$Key){return $area.Cells.Item($r,2)}}
 throw ('Setting absent: '+$Key)
}
# -------------------- ДЛЯ ТЕСТОВ --------------------
# После реального Save/Close/Open вызывает только проверяемую группу представления Config.
function Test-PresentationReport([string]$Name,[bool]$RequirePass){
 $book.Save();$book.Close($false);$script:book=$null;[GC]::Collect();[GC]::WaitForPendingFinalizers()
 $script:book=$app.Workbooks.Open($path,0,$false)
 $report=[string]$app.Run("'$($book.Name)'!modTestWorkbookInterface.RunConfigPresentationTests")
 [IO.File]::WriteAllText((Join-Path $dir ($Name+'.txt')),$report,[Text.Encoding]::UTF8)
 $total=[regex]::Match($report,'TOTAL_CONFIG_PRESENTATION: passed=(\d+); failed=(\d+)')
 if(-not$total.Success){throw ('No terminal config total: '+$Name)}
 $failures=@([regex]::Matches($report,'(?m)^FAIL[: ][^\r\n]*')|ForEach-Object Value)
 if($RequirePass-and([int]$total.Groups[2].Value-ne0-or$failures.Count)){throw $report}
 return [ordered]@{Name=$Name;Passed=[int]$total.Groups[1].Value;Failed=[int]$total.Groups[2].Value;Failures=$failures;SaveCloseOpen=$true;ReportSHA256=(Get-FileHash -LiteralPath (Join-Path $dir ($Name+'.txt'))).Hash}
}
try{
 $existing=@([Diagnostics.Process]::GetProcessesByName('EXCEL')|ForEach-Object {$_.Id})
 $app=New-Object -ComObject Excel.Application
 [uint32]$id=0;[void][TestDynamicIdOwner]::GetWindowThreadProcessId([IntPtr]$app.Hwnd,[ref]$id);$owned=[int]$id
 if($existing-contains$owned){$app=$null;throw 'New own Excel required'};$ownedNew=$true
 $ticks=(Get-Process -Id $owned).StartTime.ToUniversalTime().Ticks
 [ordered]@{PID=$owned;StartTicksUTC=[string]$ticks;RAMGiB=$memory.AvailableMBytes/1024;CommitFreeGiB=($memory.CommitLimit-$memory.CommittedBytes)/1GB}|ConvertTo-Json|Set-Content (Join-Path $dir 'Ownership.json') -Encoding UTF8
 $app.Visible=$false;$app.DisplayAlerts=$false;$app.EnableEvents=$false;$app.AutomationSecurity=1
 $book=$app.Workbooks.Open($path,0,$false)
 [void]$app.Run("'$($book.Name)'!modTestPerformance.TestConfigurePerformanceFixture",'NDS',1,'Newton','Rectangles','Normal','No')
 $range=$book.Names.Item('rngLoadCombinations').RefersToRange
 [void]$range.Cells.Item(2,1).Resize($range.Rows.Count-1,1).ClearContents();$range=$null
 $original=Test-PresentationReport 'Original_Empty' $false
 if($original.Failed-ne2-or@($original.Failures|Where-Object{$_-notmatch'config\.dropdown\.(L71\.vertical|B19\.notEmpty)$'}).Count){throw 'Expected exactly the two original dropdown prerequisites'}
 $project=$book.VBProject;$code=$project.VBComponents.Item('modTestWorkbookInterface').CodeModule
 $body=[regex]::Replace([IO.File]::ReadAllText($candidate,[Text.Encoding]::UTF8),'(?m)^Attribute [^\r\n]*\r?\n','')
 $code.DeleteLines(1,$code.CountOfLines);$code.AddFromString($body)
 $app.VBE.ActiveVBProject=$project;$compile=$app.VBE.CommandBars.FindControl(1,578);if($compile.Enabled){$compile.Execute()};if($compile.Enabled){throw 'Candidate compile failed'}
 $compile=$null;$code=$null;$project=$null;[GC]::Collect();[GC]::WaitForPendingFinalizers()
 foreach($mode in @('No','Yes')){
  $cell=Test-SettingCell 'General.DiagramExtension';$cell.Value2=$mode;$cell=$null
  foreach($kind in @('Normal','Empty')){
   $range=$book.Names.Item('rngLoadCombinations').RefersToRange
   [void]$range.Cells.Item(2,1).Resize($range.Rows.Count-1,1).ClearContents()
   if($kind-eq'Normal'){$range.Cells.Item(2,1).Value2='LC1'};$range=$null
   $record=Test-PresentationReport ($kind+'_'+$mode) $true
   if($record.Passed-lt1326){throw 'Original config assertions must remain'}
   $records+=,$record
  }
 }
 $idCell=Test-SettingCell 'Output.CombinationID';$savedIdFormula=[string]$idCell.Validation.Formula1
 $idCell.Validation.Modify(3,1,1,'=$H$72:$H$100');$idCell=$null
 $bad=Test-PresentationReport 'WrongIdSource_Empty' $false
 if(-not@($bad.Failures|Where-Object{$_-match'\.directLoadIDs$'}).Count){throw 'Wrong empty ID source must still fail'}
 $records+=,$bad
 $idCell=Test-SettingCell 'Output.CombinationID';$idCell.Validation.Modify(3,1,1,$savedIdFormula);$idCell=$null
 $profileCell=$book.Names.Item('rngLoadCombinations').RefersToRange.Cells.Item(2,5);$savedProfile=[string]$profileCell.Validation.Formula1
 [void]$book.Worksheets.Item('Config').Range('A1048575:A1048576').ClearContents();$profileCell.Validation.Modify(3,1,1,'=$A$1048575:$A$1048576');$profileCell=$null
 $bad=Test-PresentationReport 'EmptyRequiredProfileSource' $false
 if(-not@($bad.Failures|Where-Object{$_-match'\.notEmptyOrUnformedLoadIDs$'}).Count){throw 'Other empty sources must still fail'}
 $records+=,$bad
 $profileCell=$book.Names.Item('rngLoadCombinations').RefersToRange.Cells.Item(2,5);$profileCell.Validation.Modify(3,1,1,$savedProfile);$profileCell=$null
 $records+=,(Test-PresentationReport 'Restored_Empty_Yes' $true)
 [ordered]@{Status='PASS_ISOLATED_DYNAMIC_ID_AND_HORIZONTAL_PROFILE_TESTS';Original=$original;Records=$records;ProductionChanged=$false;RootSourceUnchanged=((Get-FileHash -LiteralPath $source).Hash-ceq$manifest.OriginalSHA256);FullPass=$false}|ConvertTo-Json -Depth 9|Set-Content (Join-Path $PSScriptRoot 'NativeAcceptance01.json') -Encoding UTF8
 Write-Output 'DYNAMIC_ID_GATE: normal/empty No+Yes passed; wrong-ID-source and other-empty-source still rejected; restored passed'
}catch{
 [ordered]@{Error=$_.Exception.ToString();Stack=$_.ScriptStackTrace}|ConvertTo-Json|Set-Content (Join-Path $dir 'Failure.json') -Encoding UTF8;throw
}finally{
 $cell=$null;$range=$null;$idCell=$null;$profileCell=$null;$compile=$null;$code=$null;$project=$null;[GC]::Collect();[GC]::WaitForPendingFinalizers()
 if($book){$book.Close($false)};$book=$null;[GC]::Collect();[GC]::WaitForPendingFinalizers()
 if($app-and$ownedNew){$app.Quit()};$app=$null;[GC]::Collect();[GC]::WaitForPendingFinalizers();[GC]::Collect();[GC]::WaitForPendingFinalizers()
 [ordered]@{PID=$owned;StartTicksUTC=[string]$ticks;QuitRequested=$true;ForceKilled=$false}|ConvertTo-Json|Set-Content (Join-Path $dir 'Quit.json') -Encoding UTF8
}
