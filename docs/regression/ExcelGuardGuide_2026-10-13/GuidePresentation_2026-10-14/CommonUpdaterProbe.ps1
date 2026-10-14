# Изображения фактических строк Excel; основная книга и проверяемые копии не сохраняются.
$ErrorActionPreference='Stop'
. 'tools/build_workbook/SectionRecipeGuide.ps1'
$root=(Get-Location).Path
$directory=Join-Path (Get-Location).Path 'docs/regression/ExcelGuardGuide_2026-10-13/GuidePresentation_2026-10-14'
Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public static class GuidePreviewWindow{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr hWnd,out uint id);}'
$excel=$null;$book=$null;$scratch=$null;$ownedPid=0;$ownedTicks=0
try {
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false;$excel.AutomationSecurity=3
    [uint32]$id=0;[void][GuidePreviewWindow]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$id)
    $ownedPid=[int]$id;$ownedTicks=(Get-Process -Id $ownedPid).StartTime.ToUniversalTime().Ticks
    Write-Output ('OWNED_EXCEL: id='+$ownedPid+'; utcStartTicks='+$ownedTicks)
    $path=Join-Path $directory 'BuilderProbe.xlsm'
    Copy-Item -LiteralPath (Join-Path $directory 'Candidate.xlsm') -Destination $path
    $book=$excel.Workbooks.Open($path,0,$false)
    if($book.ReadOnly){throw 'Builder probe is read-only.'}
    $first=Update-SectionRecipeGuide $book $root
    $second=Update-SectionRecipeGuide $book $root
    if($first.StartRow -ne $second.StartRow -or $first.EndRow -ne $second.EndRow -or $first.BodySHA256 -cne $second.BodySHA256){throw 'Common updater is not idempotent.'}
    $range=$book.Names.Item('rngSectionRecipeGuide').RefersToRange
    $cell=$range.Cells.Item(2,1);$cell.Value2='USER EDIT'
    $errorText=''
    try {Update-SectionRecipeGuide $book $root | Out-Null} catch {$errorText=$_.Exception.Message}
    if($errorText -notlike '*content was edited*' -or [string]$cell.Value2 -cne 'USER EDIT'){throw 'Native common updater destroyed user edit.'}
    [ordered]@{Passed=$true;Rows=$range.Rows.Count;BlankLastRow=([string]$range.Cells.Item($range.Rows.Count,1).Value2 -eq '');Idempotent=$true;ManualEditRejected=$true;MainWorkbookUnchanged=$true} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $directory 'CommonUpdaterChecks.json') -Encoding UTF8
    $book.Close($false);$book=$null
    Write-Output 'COMMON_NATIVE_GUIDE_UPDATER_PASS'
} finally {
    if($scratch){$scratch.Close($false)}
    if($book){$book.Close($false)}
    if($excel){$excel.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
    if($ownedPid){
        $p=Get-Process -Id $ownedPid -ErrorAction SilentlyContinue
        if($p -and $p.ProcessName -eq 'EXCEL' -and $p.StartTime.ToUniversalTime().Ticks -eq $ownedTicks){Stop-Process -Id $ownedPid -Force}
    }
}
