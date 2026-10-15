# Проверяет готовый файл штатным Excel без макросов, пересчёта и сохранения.
$ErrorActionPreference = 'Stop'
$root = (Get-Location).Path
$directory = Join-Path $root 'docs/regression/ExcelGuardGuide_2026-10-13/FormattingLayout_2026-10-15'
$path = Join-Path $directory 'Candidate.xlsm'
$before = (Get-FileHash -LiteralPath $path).Hash
$excel = $null; $book = $null; $ownedPid = 0; $ownedTicks = 0
Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public static class FormattingExcelWindow{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr hWnd,out uint id);}'
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 3; $excel.EnableEvents = $false
    [uint32]$windowPid = 0
    [void][FormattingExcelWindow]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$windowPid)
    $ownedPid = [int]$windowPid
    $ownedTicks = (Get-Process -Id $ownedPid).StartTime.ToUniversalTime().Ticks
    $book = $excel.Workbooks.Open($path, 0, $true)
    if (-not $book.ReadOnly) { throw 'Read-only verification expected.' }
    $excel.Calculation = -4135
    $config = $book.Worksheets.Item('Config')
    if ($excel.WorksheetFunction.CountA($config.Range('H85:P85')) -ne 0) { throw 'Separator is not empty.' }
    # Верхняя/нижняя грань может читаться из соседней таблицы; внутри промежутка рамок нет.
    foreach ($edge in @(7,10,11,12)) {
        if ($config.Range('H85:P85').Borders.Item($edge).LineStyle -ne -4142) { throw 'Separator has a cell border.' }
    }
    $caption = 'Настройки оформления схемы'
    $link = '(Подробнее)'
    $heading = $config.Range('H69')
    $start = ([string]$heading.Value2).IndexOf($link) + 1
    if ($heading.Characters($start, $link.Length).Font.Color -ne 16711680 -or $heading.Characters($start, $link.Length).Font.Underline -ne 2) { throw 'Link text lacks native blue/underline formatting.' }
    if ($heading.Characters(1, $caption.Length).Font.Color -ne 0 -or $heading.Characters(1, $caption.Length).Font.Underline -ne -4142) { throw 'Caption format differs from other headings.' }
    if ($heading.Hyperlinks.Count -ne 1 -or $heading.Hyperlinks.Item(1).SubAddress -cne "'Справка'!A724") { throw 'Heading hyperlink changed.' }
    $names = [ordered]@{
        rngPlotAnnotationSettings = 'H70:M84'; rngCircleGeometry = 'H87:K98'
        rngRectSetGeometry = 'H101:P128'; rngRoundedRectangleGeometry = 'H131:P152'
        rngHollowRectangleGeometry = 'H156:P183'; rngSystemSettings = 'A3:E138'
    }
    foreach ($entry in $names.GetEnumerator()) {
        if ($book.Names.Item($entry.Key).RefersToRange.Address($false,$false) -cne $entry.Value) { throw ('Wrong named range: ' + $entry.Key) }
    }
    foreach ($address in @('I97','I98')) {
        if ($config.Range($address).Validation.Type -ne 3) { throw ('Missing independent row list: ' + $address) }
    }
    $units = [ordered]@{}
    foreach ($address in @('J88','N111','O121','K133','N166','O176','O183')) {
        $cell = $config.Range($address)
        if (-not $cell.HasFormula -or [string]$cell.Value2 -cne 'mm') { throw ('Moved unit formula changed: ' + $address) }
        $units[$address] = [string]$cell.Formula
    }
    # Существующий сборщик уже оставляет один промежуток после полной таблицы.
    . (Join-Path $root 'tools/build_workbook/SettingsCatalog.ps1')
    $circleHeader = 70 + (Get-PlotAnnotationSettingsCatalog).Count + 1 + 2
    if ($circleHeader -ne 87) { throw 'Fresh builder spacing differs.' }
    [ordered]@{
        Passed = $true; OpenedWithoutRepair = $true; ReadOnly = $book.ReadOnly
        SeparatorEmptyAndWithoutBorders = $true; BlueUnderlinedLink = $true
        CaptionBlackWithoutUnderline = $true; LinkTarget = $heading.Hyperlinks.Item(1).SubAddress
        Names = $names; CircleIndependentRowChoices = @('I97','I98')
        MovedUnitFormulas = $units; NativeComponents = $book.VBProject.VBComponents.Count
        UserContourLineWeight = $config.Range('I84').Value2
        FreshBuilderCircleHeaderRow = $circleHeader
        OwnedExcelPID = $ownedPid
    } | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directory 'NativeCheck.json') -Encoding UTF8
} finally {
    if ($book) { $book.Close($false) }
    if ($excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    if ($ownedPid) {
        $owned = Get-Process -Id $ownedPid -ErrorAction SilentlyContinue
        if ($owned -and $owned.ProcessName -eq 'EXCEL' -and $owned.StartTime.ToUniversalTime().Ticks -eq $ownedTicks) { Stop-Process -Id $ownedPid -Force }
    }
}
if ((Get-FileHash -LiteralPath $path).Hash -cne $before) { throw 'Read-only check modified the file.' }
Write-Output 'NATIVE_CONFIG_LAYOUT_PASS'
