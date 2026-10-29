# -------------------- ДЛЯ ТЕСТОВ --------------------
# Tests missing-setting section navigation in an owned saved/reopened workbook; production is unchanged.
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../../../..')).Path
$manifest=Get-Content -LiteralPath (Join-Path $PSScriptRoot 'DeferredSourceManifest.json') -Raw -Encoding UTF8|ConvertFrom-Json
$source=Join-Path $root $manifest.Source;$candidate=Join-Path $root $manifest.Candidate
if((Get-FileHash -LiteralPath $source).Hash-cne$manifest.OriginalSHA256-or(Get-FileHash -LiteralPath $candidate).Hash-cne$manifest.CandidateSHA256){throw 'Original/candidate source guard failed'}
$scratch='C:/Users/avidzba/.codex/visualizations/2026/10/12/01a12ece-8192-7141-89b3-f83f1176184c/PlotLabelsPerformance_2026-10-28_Run01'
$dir=Join-Path $scratch 'ParallelFix03/NativeSettings01';if(Test-Path -LiteralPath $dir){throw 'Fresh native evidence required'};[void][IO.Directory]::CreateDirectory($dir)
$path=Join-Path $dir 'RC_Section_NDM.xlsm';Copy-Item -LiteralPath (Join-Path $root 'docs/regression/FullOnOff_2026-10-28_PlotLabels_Run01/Source/RC_Section_NDM.xlsm') -Destination $path
$memory=Get-CimInstance Win32_PerfFormattedData_PerfOS_Memory
if($memory.AvailableMBytes-lt8192-or($memory.CommitLimit-$memory.CommittedBytes)-lt14GB){throw 'Insufficient total native RAM/commit reserve'}
Add-Type 'using System;using System.Runtime.InteropServices;public static class TestMissingSectionFixOwner{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);}'
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
 [uint32]$id=0;[void][TestMissingSectionFixOwner]::GetWindowThreadProcessId([IntPtr]$app.Hwnd,[ref]$id);$owned=[int]$id
 if($existing-contains$owned){$app=$null;throw 'Own new Excel required'};$ownedNew=$true
 $ticks=(Get-Process -Id $owned).StartTime.ToUniversalTime().Ticks
 [ordered]@{PID=$owned;StartTicksUTC=[string]$ticks;RAMGiB=$memory.AvailableMBytes/1024;CommitFreeGiB=($memory.CommitLimit-$memory.CommittedBytes)/1GB}|ConvertTo-Json|Set-Content (Join-Path $dir 'Ownership.json') -Encoding UTF8
 $app.Visible=$false;$app.DisplayAlerts=$false;$app.EnableEvents=$false;$app.AutomationSecurity=1
 $book=$app.Workbooks.Open($path,0,$false)
 [void]$app.Run("'$($book.Name)'!modTestPerformance.TestConfigurePerformanceFixture",'NDS',1,'Newton','Rectangles','Normal','No')
 $original=[string]$app.Run("'$($book.Name)'!modTestConfiguration.RunAudit03SettingsTableGuardTests")
 [IO.File]::WriteAllText((Join-Path $dir 'Original_No.txt'),$original,[Text.Encoding]::UTF8)
 $failures=@([regex]::Matches($original,'(?m)^FAIL[: ][^\r\n]*')|ForEach-Object Value)
 $expected=@($manifest.Failures|ForEach-Object {'FAIL: '+$_}|Sort-Object)
 if($failures.Count-ne6-or(($failures|Sort-Object)-join'|')-cne($expected-join'|')){throw ('Expected only six recorded section caption mismatches: '+($failures-join';'))}
 $cfg=$book.Worksheets.Item('Config').UsedRange.Value2;$captionCount=0
 foreach($value in $cfg){if([string]$value-ceq('['+$manifest.ExpectedCaption+']')){$captionCount++}}
 if($captionCount-ne1){throw 'Current visible Config section caption must occur exactly once'}

 $project=$book.VBProject;$code=$project.VBComponents.Item('modTestConfiguration').CodeModule
 $body=[regex]::Replace([IO.File]::ReadAllText($candidate,[Text.Encoding]::UTF8),'(?m)^Attribute [^\r\n]*\r?\n','')
 $code.DeleteLines(1,$code.CountOfLines);$code.AddFromString($body)
 $app.VBE.ActiveVBProject=$project;$compile=$app.VBE.CommandBars.FindControl(1,578);if($compile.Enabled){$compile.Execute()};if($compile.Enabled){throw 'Private candidate compile failed'}
 $compile=$null;$code=$null;$project=$null;[GC]::Collect();[GC]::WaitForPendingFinalizers()
 foreach($mode in @('No','Yes')){
  Set-TestSetting 'General.DiagramExtension' $mode
  $book.Save();$book.Close($false);$book=$null;[GC]::Collect();[GC]::WaitForPendingFinalizers()
  $book=$app.Workbooks.Open($path,0,$false)
  $report=[string]$app.Run("'$($book.Name)'!modTestConfiguration.RunAudit03SettingsTableGuardTests")
  [IO.File]::WriteAllText((Join-Path $dir ('Candidate_'+$mode+'.txt')),$report,[Text.Encoding]::UTF8)
  $total=[regex]::Match($report,'TOTAL_AUDIT03_SETTINGS_TABLE_GUARDS: passed=(\d+); failed=(\d+)')
  if(-not$total.Success-or[int]$total.Groups[2].Value-ne0-or$report-match'(?m)^FAIL[: ]'){throw $report}
  if([int]$total.Groups[1].Value-ne620){throw 'All original 620 settings guard assertions must remain'}
  $records+=,[ordered]@{Mode=$mode;Passed=[int]$total.Groups[1].Value;Failed=0;SaveCloseOpen=$true;ExpectedCaption=$manifest.ExpectedCaption;VisibleCaptionCount=$captionCount;AllOriginalAssertionsPreserved=$true;ProductionChanged=$false}
  $records|ConvertTo-Json -Depth 6|Set-Content (Join-Path $PSScriptRoot 'NativeSettingsAcceptance01.json') -Encoding UTF8
  Write-Output ('DIRECTED_MISSING_SECTION '+$mode+': '+$total.Groups[1].Value+'/0; saved/reopened; exact current visible caption')
 }
}catch{
 [ordered]@{Error=$_.Exception.ToString();Stack=$_.ScriptStackTrace}|ConvertTo-Json|Set-Content (Join-Path $dir 'Failure.json') -Encoding UTF8;throw
}finally{
 $cell=$null;$range=$null;$compile=$null;$code=$null;$project=$null;[GC]::Collect();[GC]::WaitForPendingFinalizers()
 if($book){$book.Close($false)};$book=$null;[GC]::Collect();[GC]::WaitForPendingFinalizers()
 if($app-and$ownedNew){$app.Quit()};$app=$null;[GC]::Collect();[GC]::WaitForPendingFinalizers();[GC]::Collect();[GC]::WaitForPendingFinalizers()
 [ordered]@{PID=$owned;StartTicksUTC=[string]$ticks;QuitRequested=$true;ForceKilled=$false;RootSourceUnchanged=((Get-FileHash -LiteralPath $source).Hash-ceq$manifest.OriginalSHA256)}|ConvertTo-Json|Set-Content (Join-Path $dir 'Quit.json') -Encoding UTF8
}
