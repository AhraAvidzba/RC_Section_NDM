# ДЛЯ ТЕСТОВ: точная готовая книга, оба десятичных знака Excel, бинарные и Excel-снимки.
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSHOME 'Modules/Microsoft.PowerShell.Utility/Microsoft.PowerShell.Utility.psd1') -Global -ErrorAction Stop
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../../../..')).Path
$source=Join-Path $root 'workbook/output/RC_Section_NDM.xlsm'
$hash=(Get-FileHash -LiteralPath $source).Hash
$primary=Join-Path $PSScriptRoot 'SnapshotSeparators.txt'
if(Test-Path -LiteralPath $primary){throw 'Existing primary must be preserved.'}
$memory=Get-CimInstance Win32_OperatingSystem
if($memory.FreePhysicalMemory*1KB -lt 4GB -or $memory.FreeVirtualMemory*1KB -lt 8GB){throw 'Insufficient reserve; no Excel started.'}
$temp=Join-Path $env:TEMP ('NDM_LocaleSnapshots_'+[Guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($temp)
$copy=Join-Path $temp 'RC_Section_NDM.xlsm';Copy-Item -LiteralPath $source -Destination $copy
$before=@(Get-Process EXCEL -ErrorAction SilentlyContinue|ForEach-Object Id)
Add-Type 'using System;using System.Runtime.InteropServices;public static class NdmSnapshotLocaleOwner{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);}'
$app=$null;$book=$null;$owned=0;$ticks=0;$oldDecimal='';$records=@();$text=[Text.StringBuilder]::new()
$proof=[ordered]@{Status='STARTED';SourceSHA256=$hash;MacrosImported=$false;WindowsLocaleChanged=$false;Records=@()}
try{
 $app=New-Object -ComObject Excel.Application
 [uint32]$nativeId=0;[void][NdmSnapshotLocaleOwner]::GetWindowThreadProcessId([IntPtr]$app.Hwnd,[ref]$nativeId)
 $owned=[int]$nativeId;$ticks=(Get-Process -Id $owned).StartTime.ToUniversalTime().Ticks
 if($owned -in $before){throw 'Fresh own Excel required.'}
 $proof.PID=$owned;$proof.StartTicks=[string]$ticks
 $proof|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $PSScriptRoot 'SnapshotSeparators.json') -Encoding UTF8
 $app.Visible=$false;$app.DisplayAlerts=$false;$app.EnableEvents=$false;$app.AutomationSecurity=1
 $oldUse=$app.UseSystemSeparators;$oldDecimal=$app.DecimalSeparator;$oldThousands=$app.ThousandsSeparator
 $book=$app.Workbooks.Open($copy,0,$true)
 if(-not $book.ReadOnly){throw 'Read-only required.'}
 $settings=$book.Names.Item('rngSystemSettings').RefersToRange
 for($row=2;$row -le $settings.Rows.Count;$row++){if([string]$settings.Cells($row,1).Value2 -eq 'General.DiagramExtension'){$settings.Cells($row,2).Value2='No';break}}
 [void]$text.AppendLine('SOURCE_SHA256: '+$hash)
 [void]$text.AppendLine('TEST_EXCEL_PROCESS: id='+$owned+'; startTicks='+$ticks)
 foreach($separator in @('.',',')){
  $app.UseSystemSeparators=$false;$app.DecimalSeparator=$separator;$app.ThousandsSeparator=if($separator -eq '.') {','}else{'.'}
  [void]$text.AppendLine('LOCALE_STARTED: '+$app.DecimalSeparator)
  $body=[string]$app.Run(("'"+$book.Name+"'!modTestCalculationSnapshot.RunCalculationSnapshotTests"))
  [void]$text.AppendLine($body)
  $passed=$body -match 'TOTAL CALCULATION SNAPSHOT: passed=417; failed=0' -and $body -notmatch '(?m)^(FAIL:|RUNTIME ERROR|RESTORE ERROR|CLEANUP ERROR)'
  $records+=@{ExcelDecimalSeparator=$separator;Passed=$passed;ExpectedChecks=417}
  [void]$text.AppendLine('LOCALE_FINISHED: '+$separator+'; passed='+$passed)
  [IO.File]::WriteAllText($primary,$text.ToString(),[Text.UTF8Encoding]::new($true))
  if(-not $passed){throw ('Snapshot checks failed for separator '+$separator)}
 }
 $proof.Status='PASS_SNAPSHOT_DOT_COMMA_NATIVE_EXACT_BOOK'
} catch{$proof.Status='REVIEW_REQUIRED';$proof.Error=$_.Exception.ToString();throw}
finally{
 if($app -and $oldDecimal){$app.UseSystemSeparators=$false;$app.DecimalSeparator=$oldDecimal;$app.ThousandsSeparator=$oldThousands;$app.UseSystemSeparators=$oldUse}
 if($book){$book.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)}
 if($app){$app.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
 $proof.SourceUnchanged=(Get-FileHash -LiteralPath $source).Hash -ceq $hash
 $proof.Records=$records;$proof.Finished=(Get-Date).ToString('o')
 $process=Get-Process -Id $owned -ErrorAction SilentlyContinue
 $proof.OwnIdentityStillLive=($process -and [string]$process.StartTime.ToUniversalTime().Ticks -ceq [string]$ticks)
 $proof|ConvertTo-Json -Depth 5|Set-Content -LiteralPath (Join-Path $PSScriptRoot 'SnapshotSeparators.json') -Encoding UTF8
}
Write-Output 'Native snapshot suite passed for both Excel decimal separators.'
