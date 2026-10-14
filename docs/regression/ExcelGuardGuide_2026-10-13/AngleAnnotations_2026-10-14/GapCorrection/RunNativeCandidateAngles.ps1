# Baseline/candidate native renderer matrix. Temporary code stays only in memory.
$ErrorActionPreference='Stop';$root=(Get-Location).Path;$directory=$PSScriptRoot
Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public static class NativeAngleWindow{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr hWnd,out uint id);}'
foreach($stage in @('Candidate')){
    $source=Join-Path $directory ($stage+'.xlsm');$hash=(Get-FileHash -LiteralPath $source).Hash
    $folder=Join-Path $directory ($stage+'NativeAngles');if(Test-Path -LiteralPath $folder){throw 'Existing native matrix must not be replaced.'}
    New-Item -ItemType Directory -Path $folder | Out-Null
    $path=Join-Path $folder 'NativeProbe.xlsm';Copy-Item -LiteralPath $source -Destination $path
    $excel=$null;$book=$null;$pidOwned=0;$ticks=0
    try{
        $excel=New-Object -ComObject Excel.Application
        $excel.Visible=$false;$excel.EnableEvents=$false;$excel.DisplayAlerts=$false;$excel.AutomationSecurity=1;$excel.ScreenUpdating=$true
        [uint32]$id=0;[void][NativeAngleWindow]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$id)
        $pidOwned=[int]$id;$ticks=(Get-Process -Id $pidOwned).StartTime.ToUniversalTime().Ticks
        Write-Output ('OWNED_EXCEL: '+$pidOwned+'; stage='+$stage)
        $book=$excel.Workbooks.Open($path,0,$false);$excel.Calculation=-4135
        $book.VBProject.VBComponents.Item('modTestPlotConfig').CodeModule.AddFromString([IO.File]::ReadAllText((Join-Path $directory 'NativeAngleProbe.bas'),[Text.Encoding]::UTF8))
        $book.VBProject.VBComponents.Item('CSectionPlotter').CodeModule.AddFromString(@'
' ДЛЯ ТЕСТОВ: renderer вызывается только в отдельной копии, метод не сохраняется.
Public Sub GoalAngleRenderForTests(ByVal chartObject As Object, ByVal layout As CPlotAnnotationLayout)
    Set mChartObject = chartObject: Set mChart = chartObject.Chart
    mDimensionArrowType = "Triangle": mDimensionArrowSize = "Medium"
    ClearGeneratedShapes
    RenderAnnotationLayout layout
End Sub
'@)
        $control=$excel.VBE.CommandBars.FindControl(1,578);if($control.Enabled){$control.Execute()}
        if($control.Enabled){throw 'Native helper failed compilation.'}
        $result=[string]$excel.Run("'$($book.Name)'!modTestPlotConfig.GoalNativeAngleProbe",$folder)
        $result | Set-Content -LiteralPath (Join-Path $folder 'Result.txt') -Encoding UTF8
        if($result -notmatch 'passed=56; failed=0'){throw $result}
        [ordered]@{Passed=$true;Cases=56;SourceSHA256=$hash;TemporaryCodeSaved=$false} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $folder 'Run.json') -Encoding UTF8
        Write-Output ($stage+': '+$result)
        $book.Close($false);$book=$null
    }finally{
        if($book){$book.Close($false)}
        if($excel){$excel.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
        [GC]::Collect();[GC]::WaitForPendingFinalizers()
        if($pidOwned){$p=Get-Process -Id $pidOwned -ErrorAction SilentlyContinue;if($p -and $p.ProcessName -eq 'EXCEL' -and $p.StartTime.ToUniversalTime().Ticks -eq $ticks){Stop-Process -Id $pidOwned -Force}}
        if((Get-FileHash -LiteralPath $source).Hash -cne $hash){throw 'Evidence source changed.'}
    }
}
