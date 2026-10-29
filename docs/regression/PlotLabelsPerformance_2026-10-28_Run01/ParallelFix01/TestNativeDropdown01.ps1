# -------------------- ДЛЯ ТЕСТОВ --------------------
# В отдельной копии воспроизводит ошибочную предпосылку и проверяет исходные варианты списка после Save/Close/Open.
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../../../..')).Path
$manifest=Get-Content -LiteralPath (Join-Path $PSScriptRoot 'DeferredSourceManifest.json') -Raw -Encoding UTF8|ConvertFrom-Json
$source=Join-Path $root $manifest.Source;$candidate=Join-Path $root $manifest.Candidate
if((Get-FileHash -LiteralPath $source).Hash-cne$manifest.OriginalSHA256-or(Get-FileHash -LiteralPath $candidate).Hash-cne$manifest.CandidateSHA256){throw 'Original/candidate source guard failed'}
$scratch='C:/Users/avidzba/.codex/visualizations/2026/10/12/01a12ece-8192-7141-89b3-f83f1176184c/PlotLabelsPerformance_2026-10-28_Run01'
$dir=Join-Path $scratch 'ParallelFix01/NativeDropdown01';if(Test-Path -LiteralPath $dir){throw 'Fresh native evidence required'};[void][IO.Directory]::CreateDirectory($dir)
$path=Join-Path $dir 'RC_Section_NDM.xlsm';Copy-Item -LiteralPath (Join-Path $root 'docs/regression/FullOnOff_2026-10-28_PlotLabels_Run01/Source/RC_Section_NDM.xlsm') -Destination $path
$memory=Get-CimInstance Win32_PerfFormattedData_PerfOS_Memory
if($memory.AvailableMBytes-lt8192-or($memory.CommitLimit-$memory.CommittedBytes)-lt14GB){throw 'Insufficient total native RAM/commit reserve'}
Add-Type 'using System;using System.Runtime.InteropServices;public static class TestDropdownFixOwner{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);}'
$app=$null;$book=$null;$code=$null;$compile=$null;$project=$null;$cell=$null;$range=$null;$records=@();$ownedNew=$false
# Меняет только существующую настройку в частной копии.
function Set-TestSetting([string]$Key,[string]$Value){
 $range=$book.Names.Item('rngSystemSettings').RefersToRange;$values=$range.Value2
 for($r=2;$r-le$values.GetLength(0);$r++){if([string]$values[$r,1]-ceq$Key){$range.Cells.Item($r,2).Value2=$Value;return}}
 throw ('Missing setting '+$Key)
}
try{
 $existing=@([Diagnostics.Process]::GetProcessesByName('EXCEL')|ForEach-Object {$_.Id})
 $app=New-Object -ComObject Excel.Application
 [uint32]$id=0;[void][TestDropdownFixOwner]::GetWindowThreadProcessId([IntPtr]$app.Hwnd,[ref]$id);$owned=[int]$id
 if($existing-contains$owned){$app=$null;throw 'Own new Excel required'};$ownedNew=$true
 $ticks=(Get-Process -Id $owned).StartTime.ToUniversalTime().Ticks
 [ordered]@{PID=$owned;StartTicksUTC=[string]$ticks;RAMGiB=$memory.AvailableMBytes/1024;CommitFreeGiB=($memory.CommitLimit-$memory.CommittedBytes)/1GB}|ConvertTo-Json|Set-Content (Join-Path $dir 'Ownership.json') -Encoding UTF8
 $app.Visible=$false;$app.DisplayAlerts=$false;$app.EnableEvents=$false;$app.AutomationSecurity=1
 $book=$app.Workbooks.Open($path,0,$false)
 [void]$app.Run("'$($book.Name)'!modTestPerformance.TestConfigurePerformanceFixture",'NDS',1,'Newton','Rectangles','Normal','No')
 $original=[string]$app.Run("'$($book.Name)'!modTestWorkbookInterface.RunConfigPresentationTests")
 [IO.File]::WriteAllText((Join-Path $dir 'Original_No.txt'),$original,[Text.Encoding]::UTF8)
 $failures=@([regex]::Matches($original,'(?m)^FAIL[: ][^\r\n]*')|ForEach-Object Value)
 if($failures.Count-ne1-or$failures[0]-notmatch'config\.dropdown\.L71\.vertical'){throw ('Expected only the recorded vertical-source prerequisite failure: '+($failures-join';'))}
 $project=$book.VBProject;$code=$project.VBComponents.Item('modTestWorkbookInterface').CodeModule
 $body=[regex]::Replace([IO.File]::ReadAllText($candidate,[Text.Encoding]::UTF8),'(?m)^Attribute [^\r\n]*\r?\n','')
 $code.DeleteLines(1,$code.CountOfLines);$code.AddFromString($body)
 $app.VBE.ActiveVBProject=$project;$compile=$app.VBE.CommandBars.FindControl(1,578);if($compile.Enabled){$compile.Execute()};if($compile.Enabled){throw 'Private candidate compile failed'}
 $compile=$null;$code=$null;$project=$null;[GC]::Collect();[GC]::WaitForPendingFinalizers()
 foreach($mode in @('No','Yes')){
  Set-TestSetting 'General.DiagramExtension' $mode
  $book.Save();$book.Close($false);$book=$null;[GC]::Collect();[GC]::WaitForPendingFinalizers()
  $book=$app.Workbooks.Open($path,0,$false)
  $report=[string]$app.Run("'$($book.Name)'!modTestWorkbookInterface.RunConfigPresentationTests")
  [IO.File]::WriteAllText((Join-Path $dir ('Candidate_'+$mode+'.txt')),$report,[Text.Encoding]::UTF8)
  $total=[regex]::Match($report,'TOTAL_CONFIG_PRESENTATION: passed=(\d+); failed=(\d+)')
  if(-not$total.Success-or[int]$total.Groups[2].Value-ne0-or$report-match'(?m)^FAIL[: ]'){throw $report}
  $cell=$book.Names.Item('rngLoadCombinations').RefersToRange.Cells.Item(2,5)
  $formula=[string]$cell.Validation.Formula1;$range=$book.Worksheets.Item('Config').Range($formula.TrimStart('='))
  if($range.Rows.Count-ne1-or$range.Columns.Count-ne4){throw 'Expected visible four-profile row'}
  $options=@();foreach($value in $range.Value2){$options+=[string]$value}
  $saved=$cell.Formula
  foreach($value in $options){$cell.Value2=$value;if(-not$cell.Validation.Value){throw ('Profile option rejected '+$value)}}
  foreach($bad in @(($options-join','),($options-join';'),'__INVALID_DROPDOWN_OPTION__')){$cell.Value2=$bad;if($cell.Validation.Value){throw ('Joined or unknown option incorrectly accepted '+$bad)}}
  $cell.Formula=$saved
  $records+=,[ordered]@{Mode=$mode;Passed=[int]$total.Groups[1].Value;Failed=0;SaveCloseOpen=$true;Formula1=$formula;Source=$range.Address();SourceRows=$range.Rows.Count;SourceColumns=$range.Columns.Count;Options=$options;EveryOptionAccepted=$true;JoinedCommaRejected=$true;JoinedSemicolonRejected=$true;UnknownRejected=$true}
  $records|ConvertTo-Json -Depth 6|Set-Content (Join-Path $PSScriptRoot 'NativeDropdownAcceptance01.json') -Encoding UTF8
  Write-Output ('DIRECTED_DROPDOWN '+$mode+': '+$total.Groups[1].Value+'/0; saved/reopened; each four options accepted; joined/unknown rejected')
 }
}catch{
 [ordered]@{Error=$_.Exception.ToString();Stack=$_.ScriptStackTrace}|ConvertTo-Json|Set-Content (Join-Path $dir 'Failure.json') -Encoding UTF8;throw
}finally{
 $cell=$null;$range=$null;$compile=$null;$code=$null;$project=$null;[GC]::Collect();[GC]::WaitForPendingFinalizers()
 if($book){$book.Close($false)};$book=$null;[GC]::Collect();[GC]::WaitForPendingFinalizers()
 if($app-and$ownedNew){$app.Quit()};$app=$null;[GC]::Collect();[GC]::WaitForPendingFinalizers();[GC]::Collect();[GC]::WaitForPendingFinalizers()
 [ordered]@{PID=$owned;StartTicksUTC=[string]$ticks;QuitRequested=$true;ForceKilled=$false;RootSourceUnchanged=((Get-FileHash -LiteralPath $source).Hash-ceq$manifest.OriginalSHA256)}|ConvertTo-Json|Set-Content (Join-Path $dir 'Quit.json') -Encoding UTF8
}
