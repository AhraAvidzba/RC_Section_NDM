# Сравнивает схемы baseline/candidate на одинаковых синтетических данных.
# Временные диагностические методы существуют только в копиях в памяти Excel.
$ErrorActionPreference='Stop'
$root=(Get-Location).Path;$directory=$PSScriptRoot
Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public static class GoalNativeWindow{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr hWnd,out uint id);}'
$protected=@('workbook/output/RC_Section_NDM.xlsm','workbook/output/VBA_All_Code.txt','workbook/output/RC_Section_NDM_execution_report.txt')
$hashes=@{};foreach($path in $protected){$hashes[$path]=(Get-FileHash -LiteralPath (Join-Path $root $path)).Hash}
foreach($stage in @('Baseline','Candidate')) {
    $source=Join-Path $directory ($stage+'.xlsm')
    $sourceHash=(Get-FileHash -LiteralPath $source).Hash
    $target=Join-Path $directory ($stage+'Objects')
    if(Test-Path -LiteralPath $target){throw 'Existing native evidence must not be replaced.'}
    New-Item -ItemType Directory -Path $target | Out-Null
    $path=Join-Path $target 'Probe.xlsm';Copy-Item -LiteralPath $source -Destination $path
    $excel=$null;$book=$null;$ownedPid=0;$ownedTicks=0
    try {
        $excel=New-Object -ComObject Excel.Application
        $excel.Visible=$false;$excel.EnableEvents=$false;$excel.DisplayAlerts=$false;$excel.AutomationSecurity=1
        [uint32]$id=0;[void][GoalNativeWindow]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$id)
        $ownedPid=[int]$id;$ownedTicks=(Get-Process -Id $ownedPid).StartTime.ToUniversalTime().Ticks
        Write-Output ('OWNED_EXCEL: '+$ownedPid+'; stage='+$stage)
        $book=$excel.Workbooks.Open($path,0,$false)
        if($book.ReadOnly){throw 'Probe is read-only.'}
        $excel.Calculation=-4135
        $book.VBProject.VBComponents.Item('modTestPlotConfig').CodeModule.AddFromString([IO.File]::ReadAllText((Join-Path $directory 'PlotObjectProbe.bas'),[Text.Encoding]::UTF8))
        $book.VBProject.VBComponents.Item('CSectionPlotter').CodeModule.AddFromString(@'
' ДЛЯ ТЕСТОВ: захваченная рамка, только в отдельной fixture.
Public Function GoalCapturedFrame() As Variant
    GoalCapturedFrame = Array(mMinX, mMaxX, mMinY, mMaxY, mPlotLeft, mPlotTop, mPlotWidth, mPlotHeight)
End Function
'@)
        $control=$excel.VBE.CommandBars.FindControl(1,578)
        if($control -and $control.Enabled){$control.Execute()}
        $watch=[Diagnostics.Stopwatch]::StartNew()
        $result=[string]$excel.Run("'$($book.Name)'!modTestPlotConfig.GoalPlotObjectProbe",$target)
        $watch.Stop()
        $result | Set-Content -LiteralPath (Join-Path $target 'Result.txt') -Encoding UTF8
        if($result -notmatch '^TOTAL_GOAL_PLOT_OBJECTS: passed=162; failed=0$'){throw $result}
        [ordered]@{Passed=$true;Cases=162;Stage=$stage;Seconds=$watch.Elapsed.TotalSeconds;SourceSHA256=$sourceHash;TemporaryHelpersPublished=$false} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $target 'Run.json') -Encoding UTF8
        Write-Output ($stage+': '+$result)
        $book.Close($false);$book=$null
    } finally {
        if($book){$book.Close($false)}
        if($excel){$excel.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
        [GC]::Collect();[GC]::WaitForPendingFinalizers()
        if($ownedPid){$p=Get-Process -Id $ownedPid -ErrorAction SilentlyContinue;if($p -and $p.ProcessName -eq 'EXCEL' -and $p.StartTime.ToUniversalTime().Ticks -eq $ownedTicks){Stop-Process -Id $ownedPid -Force}}
        if((Get-FileHash -LiteralPath $source).Hash -cne $sourceHash){throw 'Evidence source changed.'}
        foreach($protectedPath in $protected){if((Get-FileHash -LiteralPath (Join-Path $root $protectedPath)).Hash -cne $hashes[$protectedPath]){throw 'Production changed during native probe.'}}
    }
}
