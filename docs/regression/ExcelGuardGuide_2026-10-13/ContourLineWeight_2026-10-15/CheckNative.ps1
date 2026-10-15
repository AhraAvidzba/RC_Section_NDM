# Проверяет новую строку, ссылки и начальную сборку в собственном Excel без сохранения.
$ErrorActionPreference='Stop'
$root=(Get-Location).Path
$directory=Join-Path $root 'docs/regression/ExcelGuardGuide_2026-10-13/ContourLineWeight_2026-10-15'
. (Join-Path $root 'tools/build_workbook/SettingsCatalog.ps1')
$excel=$null;$book=$null;$fresh=$null;$ownedPid=0;$ticks=0
Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public static class ContourCheckWindow{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr hWnd,out uint id);}'
try {
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false;$excel.AutomationSecurity=3
    [uint32]$id=0;[void][ContourCheckWindow]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$id)
    $ownedPid=[int]$id;$ticks=(Get-Process -Id $ownedPid).StartTime.ToUniversalTime().Ticks
    Write-Output ('OWNED_EXCEL: '+$ownedPid)
    $book=$excel.Workbooks.Open((Join-Path $directory 'Candidate.xlsm'),0,$true)
    $table=$book.Names.Item('rngPlotAnnotationSettings').RefersToRange
    $sheet=$table.Worksheet
    $cell=$null
    for($r=2;$r -le $table.Rows.Count;$r++){if([string]$table.Cells($r,1).Value2 -eq 'ContourLineWeight'){$cell=$table.Cells($r,2)}}
    if(-not $cell -or [double]$cell.Value2 -ne 2 -or $cell.MergeArea.Columns.Count -ne 2 -or $cell.HorizontalAlignment -ne -4108){throw 'Actual common input/default/alignment mismatch.'}
    if($cell.Validation.Type -ne 2 -or $cell.Validation.Operator -ne 5 -or $cell.Validation.IgnoreBlank){throw 'Actual positive-number validation mismatch.'}
    foreach($edge in @(7,8,9,10)){if($table.Borders.Item($edge).LineStyle -eq -4142){throw 'Missing table edge.'}}
    $link=$table.Cells($table.Rows.Count,6).Hyperlinks.Item(1)
    $help=$book.Worksheets.Item('Справка')
    if($link.SubAddress -ne "'Справка'!A1763" -or [string]$help.Range('A1763').Value2 -ne 'Настройки оформления схемы: ContourLineWeight' -or $help.Range('A1763').Hyperlinks.Item(1).SubAddress -ne "'Config'!I84"){throw 'Help navigation mismatch.'}
    if([string]$sheet.Cells($table.Row-1,$table.Column).Value2 -ne 'Настройки оформления схемы (Подробнее)'){throw 'Actual Config title mismatch.'}
    $book.Close($false);$book=$null
    $fresh=$excel.Workbooks.Add(-4167);$sheet=$fresh.Worksheets.Item(1);$sheet.Name='Config'
    Add-PlotAnnotationSettingsTable $fresh $sheet 2 8
    $table=$fresh.Names.Item('rngPlotAnnotationSettings').RefersToRange
    $cell=$table.Cells($table.Rows.Count,2)
    if([string]$table.Cells($table.Rows.Count,1).Value2 -ne 'ContourLineWeight' -or [double]$cell.Value2 -ne 2 -or $cell.MergeArea.Columns.Count -ne 2 -or $cell.Validation.Type -ne 2 -or $cell.Validation.Operator -ne 5){throw 'Fresh builder differs from current workbook.'}
    $result=[ordered]@{Passed=$true;CurrentInput='Config!I84:J84';Default=2;Units='pt';CommonInputMergedAndCentered=$true;PositiveValidation=$true;TableAllEdgesPresent=$true;HelpNavigation=$true;FreshBuilderSameSetting=$true}
    $result | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $directory 'NativeCheck.json') -Encoding UTF8
    Write-Output 'NATIVE_SETTING_AND_BUILDER_PASS'
} finally {
    if($book){$book.Close($false)}
    if($fresh){$fresh.Close($false)}
    if($excel){$excel.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
    if($ownedPid){$process=Get-Process -Id $ownedPid -ErrorAction SilentlyContinue;if($process -and $process.ProcessName -eq 'EXCEL' -and $process.StartTime.ToUniversalTime().Ticks -eq $ticks){Stop-Process -Id $ownedPid -Force}}
}
