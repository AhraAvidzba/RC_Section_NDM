# Обновляет только сохранённую схему кандидата по прежним Results, без расчёта.
$ErrorActionPreference='Stop';$root=(Get-Location).Path;$directory=$PSScriptRoot
. (Join-Path $root 'tools/build_workbook/SettingsCatalog.ps1')
$path=Join-Path $directory 'Candidate.xlsm';$before=(Get-FileHash -LiteralPath $path).Hash
$printAreas=@(Get-WorkbookPrintAreas $path)
$main=Join-Path $root 'workbook/output/RC_Section_NDM.xlsm';$mainHash=(Get-FileHash -LiteralPath $main).Hash
Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public static class RefreshWindow{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr hWnd,out uint id);}'
$excel=$null;$book=$null;$ownedPid=0;$ownedTicks=0
try{
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false;$excel.EnableEvents=$false;$excel.DisplayAlerts=$false;$excel.AutomationSecurity=1
    [uint32]$id=0;[void][RefreshWindow]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$id)
    $ownedPid=[int]$id;$ownedTicks=(Get-Process -Id $ownedPid).StartTime.ToUniversalTime().Ticks
    Write-Output ('OWNED_EXCEL: '+$ownedPid)
    $book=$excel.Workbooks.Open($path,0,$false);if($book.ReadOnly){throw 'Candidate is read-only.'}
    $excel.Calculation=-4135
    $sheet=$book.Worksheets.Item('Расчет');$chart=$sheet.ChartObjects().Item('chtNDMSectionPlot')
    $frame=@([double]$chart.Left,[double]$chart.Top,[double]$chart.Width,[double]$chart.Height);$oldShapes=[int]$chart.Chart.Shapes.Count
    $excel.Run("'$($book.Name)'!modWorkbookCalculation.UpdateSectionPlotForWorkbook",$book,$true)
    $newFrame=@([double]$chart.Left,[double]$chart.Top,[double]$chart.Width,[double]$chart.Height)
    if(($frame -join ',') -cne ($newFrame -join ',')){throw 'Saved chart frame changed.'}
    $newShapes=[int]$chart.Chart.Shapes.Count
    $control=$excel.VBE.CommandBars.FindControl(1,578);if($control -and $control.Enabled){$control.Execute()}
    $book.Save();$book.Close($false);$book=$null
    Restore-WorkbookPrintAreas $path $printAreas
    [ordered]@{Passed=$true;Compiled=$true;BeforeSHA256=$before;AfterSHA256=(Get-FileHash -LiteralPath $path).Hash;OldShapes=$oldShapes;NewShapes=$newShapes;ChartFrame=$frame;RecalculationRun=$false;InputAndResultsVerificationPending=$true} | ConvertTo-Json -Depth 3 | Set-Content -LiteralPath (Join-Path $directory 'ChartRefresh.json') -Encoding UTF8
}finally{
    if($book){$book.Close($false)}
    if($excel){$excel.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
    if($ownedPid){$p=Get-Process -Id $ownedPid -ErrorAction SilentlyContinue;if($p -and $p.ProcessName -eq 'EXCEL' -and $p.StartTime.ToUniversalTime().Ticks -eq $ownedTicks){Stop-Process -Id $ownedPid -Force}}
    if((Get-FileHash -LiteralPath $main).Hash -cne $mainHash){throw 'Production changed.'}
}
