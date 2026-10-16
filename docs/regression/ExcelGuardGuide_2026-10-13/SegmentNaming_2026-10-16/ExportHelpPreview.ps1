# Изображение фактической Справки из собственной проверочной книги; файл не сохраняется.
$ErrorActionPreference='Stop'
Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public static class NamingPreview{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr hWnd,out uint id);}'
$excel=$null;$book=$null;$ownedPid=0;$ticks=0
try {
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false;$excel.AutomationSecurity=3
    [uint32]$id=0;[void][NamingPreview]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$id)
    $ownedPid=[int]$id;$ticks=(Get-Process -Id $ownedPid).StartTime.ToUniversalTime().Ticks
    $book=$excel.Workbooks.Open((Join-Path $PSScriptRoot 'Accepted.xlsm'),0,$true)
    $sheet=$book.Names.Item('rngSectionRecipeGuide').RefersToRange.Worksheet;$sheet.Activate()
    $range=$sheet.Range('A580:F586');$range.EntireRow.Hidden=$false
    $range.CopyPicture(1,-4147)
    $scratch=$book.Worksheets.Add()
    $chart=$scratch.ChartObjects().Add(0,0,[double]$range.Width,[double]$range.Height)
    $chart.Activate();$chart.Chart.Paste()
    if(-not $chart.Chart.Export((Join-Path $PSScriptRoot 'HelpNamedSegments.png'),'PNG')){throw 'Export failed'}
    Write-Output 'HELP_PREVIEW_EXPORTED'
} finally {
    if($book){$book.Close($false)}
    if($excel){$excel.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
    $process=Get-Process -Id $ownedPid -ErrorAction SilentlyContinue
    if($process -and $process.ProcessName -eq 'EXCEL' -and $process.StartTime.ToUniversalTime().Ticks -eq $ticks){
        try {Stop-Process -Id $ownedPid -Force} catch {if(Get-Process -Id $ownedPid -ErrorAction SilentlyContinue){throw}}
    }
}
