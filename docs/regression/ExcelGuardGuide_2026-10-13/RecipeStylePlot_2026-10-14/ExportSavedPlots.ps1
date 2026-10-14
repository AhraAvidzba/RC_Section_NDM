# Получает растровые изображения из уже сохранённых графических fixtures.
# Chart активируется только в собственном Excel, исходные книги не сохраняются.
$ErrorActionPreference='Stop';$directory=$PSScriptRoot
Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public static class SavedPlotWindow{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr hWnd,out uint id);}'
$excel=$null;$book=$null;$ownedPid=0;$ownedTicks=0;$checks=@()
try {
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false;$excel.EnableEvents=$false;$excel.DisplayAlerts=$false;$excel.AutomationSecurity=3;$excel.ScreenUpdating=$true
    [uint32]$id=0;[void][SavedPlotWindow]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$id)
    $ownedPid=[int]$id;$ownedTicks=(Get-Process -Id $ownedPid).StartTime.ToUniversalTime().Ticks
    Write-Output ('OWNED_EXCEL: '+$ownedPid)
    foreach($stage in @('BaselineObjects','CandidateObjects')) {
        $folder=Join-Path $directory $stage
        foreach($file in Get-ChildItem -LiteralPath $folder -Filter 'Case*.xlsx') {
            $hash=(Get-FileHash -LiteralPath $file.FullName).Hash
            $book=$excel.Workbooks.Open($file.FullName,0,$true)
            $sheet=$book.Worksheets.Item('Расчет');$sheet.Activate()
            $chart=$sheet.ChartObjects().Item('chtNDMSectionPlot');$chart.Activate();$chart.Chart.Refresh()
            $png=Join-Path $folder ($file.BaseName+'.png')
            $ok=$chart.Chart.Export($png,'PNG')
            if(-not $ok -or (Get-Item -LiteralPath $png).Length -le 0){throw 'Native chart image export is empty.'}
            $book.Close($false);$book=$null
            if((Get-FileHash -LiteralPath $file.FullName).Hash -cne $hash){throw 'Read-only plot source changed.'}
            $checks+=,[ordered]@{Stage=$stage;File=$file.Name;Bytes=(Get-Item -LiteralPath $png).Length;SourceUnchanged=$true}
        }
    }
    [ordered]@{Passed=$true;Images=$checks} | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $directory 'NativeImages.json') -Encoding UTF8
} finally {
    if($book){$book.Close($false)}
    if($excel){$excel.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
    if($ownedPid){$p=Get-Process -Id $ownedPid -ErrorAction SilentlyContinue;if($p -and $p.ProcessName -eq 'EXCEL' -and $p.StartTime.ToUniversalTime().Ticks -eq $ownedTicks){Stop-Process -Id $ownedPid -Force}}
}
