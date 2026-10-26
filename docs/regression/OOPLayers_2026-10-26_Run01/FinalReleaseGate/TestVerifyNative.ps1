# ДЛЯ ТЕСТОВ: нативная проверка готовой книги ReadOnly и фактический экспорт AllCode без импорта/сохранения.
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSHOME 'Modules/Microsoft.PowerShell.Utility/Microsoft.PowerShell.Utility.psd1') -Global -ErrorAction Stop
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../../../..')).Path
$source=Join-Path $root 'workbook/output/RC_Section_NDM.xlsm'
$hash=(Get-FileHash -LiteralPath $source).Hash
# ДЛЯ ТЕСТОВ: окончательная проверка разрешена только после закрытия всего полного набора.
$completePath=Join-Path $root 'docs/regression/FullSuiteLaunch_2026-10-26_OOP_Run01/Completion.json'
if(-not(Test-Path -LiteralPath $completePath)){throw 'Full driver completion is absent; no native Excel started.'}
$complete=Get-Content -LiteralPath $completePath -Raw -Encoding UTF8|ConvertFrom-Json
$reconciliationPath=Join-Path $root 'docs/regression/OOPLayers_2026-10-26_Run01/RunnerFooterFix/ExitReconciliation.json'
$reconciliation=Get-Content -LiteralPath $reconciliationPath -Raw -Encoding UTF8|ConvertFrom-Json
if($reconciliation.Status -ne 'PASS_POST_ACCEPTANCE_FOOTER_RECONCILED' -or $reconciliation.SourceSHA256 -cne $hash -or -not $reconciliation.AllMandatoryTestStagesPassed -or -not $reconciliation.AllPrimaryReportsRead){throw 'Whole primary corpus has not passed on this exact workbook.'}
if((Get-FileHash -LiteralPath $completePath).Hash -cne $reconciliation.OriginalDriverCompletionSHA256 -or $complete.SourceSHA256 -cne $hash){throw 'Original driver evidence changed after exit reconciliation.'}
$closurePath=Join-Path $PSScriptRoot 'ProcessClosureBeforeNative.json'
if(-not(Test-Path -LiteralPath $closurePath)){throw 'All-owned-process closure proof is absent.'}
$closure=Get-Content -LiteralPath $closurePath -Raw -Encoding UTF8|ConvertFrom-Json
if(-not $closure.AllOwnedIdentitiesClosed -or $closure.SourceSHA256 -cne $hash){throw 'Closure proof does not cover this completed workbook.'}
foreach($identity in $closure.Identities){
    $running=Get-Process -Id $identity.PID -ErrorAction SilentlyContinue
    if($running -and $running.StartTime.ToUniversalTime().Ticks.ToString() -eq [string]$identity.StartTicks){throw 'A full-run owned identity is still alive.'}
}
$memory=Get-CimInstance Win32_OperatingSystem
if($memory.FreePhysicalMemory*1KB -lt 4GB -or $memory.FreeVirtualMemory*1KB -lt 8GB){throw 'Insufficient reserve; no Excel started.'}
Add-Type 'using System;using System.Runtime.InteropServices;public static class NdmNativeMessageOwner{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);}'
$beforePids=@(Get-Process EXCEL -ErrorAction SilentlyContinue|ForEach-Object Id)
$app=$null;$book=$null;$owned=0;$ticks=0;$ownConfirmed=$false;$records=@();$bodies=[ordered]@{};$text=New-Object Text.StringBuilder
$proof=[ordered]@{At=(Get-Date).ToString('o');Status='STARTED';SourceSHA256=$hash;FreePhysicalGiB=[math]::Round($memory.FreePhysicalMemory*1KB/1GB,3);FreeCommitGiB=[math]::Round($memory.FreeVirtualMemory*1KB/1GB,3);ImportPerformed=$false;MacrosRun=$false;SavePerformed=$false}
try{
    $app=New-Object -ComObject Excel.Application
    [uint32]$nativeId=0;[void][NdmNativeMessageOwner]::GetWindowThreadProcessId([IntPtr]$app.Hwnd,[ref]$nativeId)
    $owned=[int]$nativeId;$ticks=(Get-Process -Id $owned).StartTime.ToUniversalTime().Ticks
    if($owned -in $beforePids){throw 'New owned Excel required.'}
    $ownConfirmed=$true; $proof.PID=$owned;$proof.StartTicksUtc=[string]$ticks
    $proof|ConvertTo-Json -Depth 5|Set-Content -LiteralPath (Join-Path $PSScriptRoot 'NativeVerification.json') -Encoding UTF8
    $app.Visible=$false;$app.DisplayAlerts=$false;$app.EnableEvents=$false;$app.AutomationSecurity=3
    $book=$app.Workbooks.Open($source,0,$true)
    if(-not $book.ReadOnly){throw 'Read-only required.'}
    $proof.ReadOnly=$book.ReadOnly;$proof.ComponentCount=$book.VBProject.VBComponents.Count
    if($proof.ComponentCount -ne 178){throw 'Unexpected component count.'}
    [void]$text.AppendLine('VBA PROJECT EXPORT');[void]$text.AppendLine('Workbook: '+$book.Name);[void]$text.AppendLine('Date: '+(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
    foreach($component in $book.VBProject.VBComponents){
        $count=$component.CodeModule.CountOfLines
        $body=if($count){$component.CodeModule.Lines(1,$count)}else{''}
        $bodies[$component.Name]=$body
        [void]$text.AppendLine(('='*100));[void]$text.AppendLine('COMPONENT: '+$component.Name);[void]$text.AppendLine('TYPE: '+$component.Type);[void]$text.AppendLine('LINES: '+$count);[void]$text.AppendLine(('='*100));if($count){[void]$text.AppendLine($body)};[void]$text.AppendLine()
        $records+=@{Name=$component.Name;Type=$component.Type;Lines=$count}
    }
    $proof.NamesCount=$book.Names.Count;$proof.Sheets=@($book.Worksheets|ForEach-Object Name)
    $config=$book.Worksheets('Config');$proof.CaptionH53=[string]$config.Range('H53').Value2;$proof.CaptionH58=[string]$config.Range('H58').Value2
    $proof.SourceEventTemplate='tools/build_workbook/WorkbookHelpEvents.vba';$proof.DocumentModuleName=[string]$book.CodeName
    $book.Close($false);$book=$null
    [IO.File]::WriteAllText((Join-Path $PSScriptRoot 'NativeSnapshot.txt'),$text.ToString(),(New-Object Text.UTF8Encoding($true)))
    $bodies|ConvertTo-Json -Depth 3|Set-Content -LiteralPath (Join-Path $PSScriptRoot 'PublishedNativeCode.json') -Encoding UTF8
    $proof.Status='PASS_NATIVE_READONLY_EXPORT';$proof.Components=$records
}finally{
    if($book){$book.Close($false)}
    if($app){
        $ownNow=Get-Process -Id $owned -ErrorAction SilentlyContinue
        if($ownConfirmed -and $ownNow -and $ownNow.StartTime.ToUniversalTime().Ticks -eq $ticks){$app.Quit()}
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)
    }
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
$proof.SourceUnchanged=(Get-FileHash -LiteralPath $source).Hash -ceq $hash
if(-not $proof.SourceUnchanged){throw 'Output changed during read-only verification.'}
$proof.Scope='NATIVE_READONLY_EXPORT_ONLY'
$proof|ConvertTo-Json -Depth 6|Set-Content -LiteralPath (Join-Path $PSScriptRoot 'NativeVerification.json') -Encoding UTF8
Write-Output ('Native output read-only verified: components='+$proof.ComponentCount+'; SHA256='+$hash)
