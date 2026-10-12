# Адресно заменяет согласованные пояснения; данные и структура листов сохраняются.
param([Parameter(Mandatory=$true)][string]$WorkbookPath,[Parameter(Mandatory=$true)][string]$ReportPath)
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
. (Join-Path $PSScriptRoot 'SettingsCatalog.ps1')
. (Join-Path $PSScriptRoot 'HelpOutline.ps1')
. (Join-Path $PSScriptRoot 'StandaloneHelpText.ps1')
$allowed=[IO.Path]::GetFullPath((Join-Path $root 'docs/regression/GeometryReview_2026-10-12'))+[IO.Path]::DirectorySeparatorChar
$path=(Resolve-Path -LiteralPath (Join-Path $root $WorkbookPath)).Path
$report=[IO.Path]::GetFullPath((Join-Path $root $ReportPath))
if (-not $path.StartsWith($allowed,[StringComparison]::OrdinalIgnoreCase) -or -not $report.StartsWith($allowed,[StringComparison]::OrdinalIgnoreCase)) {throw 'Only private review paths are allowed.'}
$areas=@(Get-WorkbookPrintAreas (Join-Path $root 'docs/regression/GeometryReview_2026-10-12/Baseline/RC_Section_NDM.xlsm'))
Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public static class StandaloneTextWindow{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr hWnd,out uint id);}'
$excel=$null;$book=$null;$ownedPid=0;$ownedTicks=0;$changes=@()
try {
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false;$excel.AutomationSecurity=3
    [uint32]$id=0;[void][StandaloneTextWindow]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$id)
    $ownedPid=[int]$id;$ownedTicks=(Get-Process -Id $ownedPid).StartTime.ToUniversalTime().Ticks
    Write-Output ('OWNED_EXCEL: id='+$ownedPid+'; utcStartTicks='+$ownedTicks)
    $book=$excel.Workbooks.Open($path,0,$false)
    if ($book.ReadOnly) {throw 'Private review book is read-only.'}
    $changes=@(Update-StandaloneHelpText $book)
    $sheet=$book.Worksheets.Item('Справка')
    $norm=$sheet.Cells.Find('СП 35.13330.2011',$sheet.Cells.Item(1,1),-4163,1)
    if ($null -eq $norm) {throw 'SP35 source row is missing.'}
    $sources=Get-Content -LiteralPath (Join-Path $PSScriptRoot 'NormativeHelpSources.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    $manual=@($sources | Where-Object {$_.kind -eq 'manual'})
    if ($manual.Count -ne 1) {throw 'This directed update expects one manual.'}
    $row=[int]$norm.Row+1
    if ([string]$sheet.Cells.Item($row,1).Value2 -cne [string]$manual[0].caption) {
        $target=$sheet.Range($sheet.Cells.Item($row,1),$sheet.Cells.Item($row,6))
        if ($excel.WorksheetFunction.CountA($target) -ne 0) {throw 'Manual source row is not empty; preserve existing Help.'}
        $sheet.Range($sheet.Cells.Item($norm.Row,1),$sheet.Cells.Item($norm.Row,6)).Copy($target) | Out-Null
        $sheet.Cells.Item($row,1).Value2=[string]$manual[0].caption
        $sheet.Cells.Item($row,2).Value2=[string]$manual[0].description
        $target.WrapText=$true
        $target.Borders.Item(10).LineStyle=1
        $changes+=,[ordered]@{Sheet=$sheet.Name;Address=$target.Address();Before='';After=[string]$manual[0].description}
    }
    if ([string]$sheet.Cells.Item($row,2).Value2 -cne [string]$manual[0].description -or [int]$sheet.Rows.Item($row).OutlineLevel -ne 2) {throw 'Manual source is outside Config outline or has stale text.'}
    $sheet.Rows.Item($row).RowHeight=30
    $important=[int]$row+1
    if ([string]$sheet.Cells.Item($important,1).Value2 -ceq 'Важно') {
        $expectedImportant='Справка описывает работу программы и смысл ее настроек. Если применение пункта СП еще не проверено, это прямо указано в соответствующем пояснении.'
        if ([string]$sheet.Cells.Item($important,2).Value2 -cne $expectedImportant) {throw 'Important explanation was edited; preserve user text.'}
        $sheet.Range($sheet.Cells.Item($important,1),$sheet.Cells.Item($important,6)).ClearContents() | Out-Null
        Resize-HelpEmptyRows $book $sheet $important 1 $false
        $changes+=,[ordered]@{Sheet=$sheet.Name;Address='A'+$important+':F'+$important;Before='Важно';After=''}
    }
    $outline=Set-HelpSectionOrderAndOutline $book
    # Картинка подтверждает чтение пояснений и рамку источников в настоящем Excel.
    $initialSheet=[string]$book.ActiveSheet.Name
    $sheet.Activate();$excel.ScreenUpdating=$true
    $config=$sheet.Cells.Find('Справка по настройкам листа Config',$sheet.Cells.Item(1,1),-4163,1)
    $pictureRange=$sheet.Range($sheet.Cells.Item(1,1),$sheet.Cells.Item([int]$config.Row+4,6))
    $pictureRange.CopyPicture(1,-4147)
    $scratch=$book.Worksheets.Add()
    try {
        $chart=$scratch.ChartObjects().Add(0,0,[double]$pictureRange.Width,[double]$pictureRange.Height)
        $chart.Activate();$chart.Chart.Paste()
        if (-not $chart.Chart.Export((Join-Path (Split-Path -Parent $report) 'NormativeSources.png'),'PNG')) {throw 'Normative source image export failed.'}
    } finally {$scratch.Delete();$sheet.Activate()}
    $sheet.Range('A1').Select();$excel.ActiveWindow.ScrollRow=1;$excel.ActiveWindow.ScrollColumn=1
    $book.Worksheets.Item($initialSheet).Activate()
    $book.Save()
} finally {
    if ($book) {$book.Close($false)}
    if ($excel) {$excel.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
    if ($ownedPid) {
        $owned=Get-Process -Id $ownedPid -ErrorAction SilentlyContinue
        if ($owned -and $owned.ProcessName -eq 'EXCEL' -and $owned.StartTime.ToUniversalTime().Ticks -eq $ownedTicks) {Stop-Process -Id $ownedPid -Force}
    }
}
Restore-WorkbookPrintAreas $path $areas
[ordered]@{Status='STANDALONE_TEXTS_UPDATED';Changes=$changes;Outline=$outline;WorkbookSHA256=(Get-FileHash -LiteralPath $path).Hash} | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $report -Encoding UTF8
Write-Output ('STANDALONE_TEXTS_UPDATED: '+$changes.Count)
