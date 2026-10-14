# Изображения фактических строк Excel; основная книга и проверяемые копии не сохраняются.
$ErrorActionPreference='Stop'
$directory=Join-Path (Get-Location).Path 'docs/regression/ExcelGuardGuide_2026-10-13/GuideContourSummary_2026-10-14'
Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public static class GuidePreviewWindow{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr hWnd,out uint id);}'
$excel=$null;$book=$null;$scratch=$null;$ownedPid=0;$ownedTicks=0
try {
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false;$excel.AutomationSecurity=3
    [uint32]$id=0;[void][GuidePreviewWindow]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$id)
    $ownedPid=[int]$id;$ownedTicks=(Get-Process -Id $ownedPid).StartTime.ToUniversalTime().Ticks
    Write-Output ('OWNED_EXCEL: id='+$ownedPid+'; utcStartTicks='+$ownedTicks)
    foreach($name in @('Baseline','Candidate')) {
        $path=Join-Path $directory ($name+'.xlsm')
        $before=(Get-FileHash -LiteralPath $path).Hash
        $book=$excel.Workbooks.Open($path,0,$true)
        if(-not $book.ReadOnly){throw 'Preview workbook is not read-only.'}
        $guide=$book.Names.Item('rngSectionRecipeGuide').RefersToRange
        $sheet=$guide.Worksheet;$sheet.Activate()
        $last=$guide.Row+$guide.Rows.Count-1
        $range=$sheet.Range($sheet.Cells.Item($guide.Row+5,1),$sheet.Cells.Item($guide.Row+19,6))
        $range.EntireRow.Hidden=$false
        $range.CopyPicture(1,-4147)
        $scratch=$excel.Workbooks.Add()
        $chart=$scratch.Worksheets.Item(1).ChartObjects().Add(0,0,[double]$range.Width,[double]$range.Height)
        $chart.Activate();$chart.Chart.Paste()
        if(-not $chart.Chart.Export((Join-Path $directory ($name+'.png')),'PNG')){throw 'Preview export failed.'}
        $scratch.Close($false);$scratch=$null
        $book.Close($false);$book=$null
        if((Get-FileHash -LiteralPath $path).Hash -cne $before){throw 'Preview source was changed.'}
    }
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
