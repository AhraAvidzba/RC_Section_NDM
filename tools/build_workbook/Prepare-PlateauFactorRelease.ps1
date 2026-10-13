# Готовит числовой наклон плато на копии актуальной книги, сохраняя ввод и Results.
param([string]$Directory='docs/regression/GeometryFull_2026-10-13/PlateauFactor')
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
. (Join-Path $PSScriptRoot 'SettingsCatalog.ps1')
$source=Join-Path $root 'workbook/output/RC_Section_NDM.xlsm'
$expected='871B42914FCB2FCF0168F43896D66BA603B6B170F6E6A9F9DB9206CDC3CA9FEC'
if ((Get-FileHash -LiteralPath $source).Hash -cne $expected) {throw 'Актуальная книга изменилась: требуется повторная проверка ввода.'}
$directoryPath=[IO.Path]::GetFullPath((Join-Path $root $Directory))
$allowed=[IO.Path]::GetFullPath((Join-Path $root 'docs/regression/GeometryFull_2026-10-13'))+[IO.Path]::DirectorySeparatorChar
if (-not $directoryPath.StartsWith($allowed,[StringComparison]::OrdinalIgnoreCase)) {throw 'Копия должна находиться в каталоге доказательств.'}
if (Test-Path -LiteralPath $directoryPath) {throw 'Каталог уже существует: исходные доказательства не перезаписываются.'}
New-Item -ItemType Directory -Path $directoryPath | Out-Null
$target=Join-Path $directoryPath 'RC_Section_NDM.xlsm'
Copy-Item -LiteralPath $source -Destination (Join-Path $directoryPath 'Before.xlsm')
Copy-Item -LiteralPath $source -Destination $target
Copy-Item -LiteralPath (Join-Path $root 'workbook/output/VBA_All_Code.txt') -Destination (Join-Path $directoryPath 'BeforeVBA.txt')
$printAreas=@(Get-WorkbookPrintAreas $target)
Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public static class PlateauFactorWindow { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId); }'
$excel=$null;$book=$null;$ownedPid=0;$ownedTicks=0
try {
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false;$excel.AutomationSecurity=3
    [uint32]$id=0;[void][PlateauFactorWindow]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$id)
    $ownedPid=[int]$id;$ownedTicks=(Get-Process -Id $ownedPid).StartTime.ToUniversalTime().Ticks
    Write-Output ('OWNED_EXCEL: id='+$ownedPid+'; utcStartTicks='+$ownedTicks)
    $book=$excel.Workbooks.Open($target,0,$false)
    if ($book.ReadOnly) {throw 'Проверяемая копия открылась только для чтения.'}
    $setting=Update-PlateauSlopeFactorSetting $book
    $config=$book.Worksheets.Item('Config')
    $cell=$config.Range([string]$setting.Address)
    if ([double]$cell.Value2 -ne 0.001) {throw 'Миграция Yes не установила согласованный default 0,001.'}
    # Проверяем все виды переноса на собственной копии. Итог возвращаем к
    # согласованному default; остальные строки Config не пересобираются.
    $checks=@()
    $cell.Offset(0,-1).Value2='Solver.PlateauSlope.Enabled';$cell.Value2='No'
    Update-PlateauSlopeFactorSetting $book | Out-Null
    if ([double]$cell.Value2 -ne 0) {throw 'Миграция No должна сохранять горизонтальное плато.'}
    $checks+='No -> 0'
    $cell.Formula='=3/1000'
    Update-PlateauSlopeFactorSetting $book | Out-Null
    if ([string]$cell.Formula -cne '=3/1000') {throw 'Повторное обновление изменило пользовательскую формулу.'}
    $checks+='Numeric formula preserved'
    $cell.Offset(0,-1).Value2='Solver.PlateauSlope.Enabled';$cell.Value2='Yes'
    Update-PlateauSlopeFactorSetting $book | Out-Null
    if ([double]$cell.Value2 -ne 0.001) {throw 'Миграция Yes должна устанавливать согласованный default 0,001.'}
    $checks+='Yes -> 0.001'
    $title=$config.Cells.Find('Контрольные точки диаграмм*',$config.Cells.Item(1,1),-4163,1,1,1,$false,$false,$false)
    if ($null -eq $title) {throw 'Нет контрольной таблицы диаграмм.'}
    Add-MaterialDiagramControlTables $config ($title.Row+1) $title.Column -CoordinatesOnly
    $config.Calculate()
    $system=$book.Names.Item('rngSystemSettings').RefersToRange
    $extension=0
    for ($row=2;$row -le $system.Rows.Count;$row++) {
        if ([string]$system.Cells.Item($row,1).Value2 -ceq 'General.DiagramExtension') {$extension=$system.Row+$row-1}
    }
    if ($cell.Row+1 -ne $extension) {throw 'Коэффициент должен находиться сразу перед General.DiagramExtension.'}
    $guide=$book.Worksheets.Item('Справка')
    $link=$config.Cells.Item($cell.Row,5).Hyperlinks.Item(1)
    if ([string]$guide.Range(([string]$link.SubAddress -split '!')[-1]).Value2 -cne 'General.PlateauSlopeFactor') {throw 'Ссылка ведёт к другой настройке.'}
    if ([int]$cell.HorizontalAlignment -ne -4108 -or [int]$cell.Validation.Type -ne 2 -or [int]$cell.Validation.Operator -ne 7) {throw 'Оформление или числовая проверка ввода не установлены.'}
    if ([int]$config.Cells.Item($cell.Row,5).Borders.Item(10).LineStyle -eq -4142) {throw 'Нет правой рамки Справки.'}
    $book.Save();$book.Close($false);$book=$null
    Restore-WorkbookPrintAreas $target $printAreas
    if ((Get-FileHash -LiteralPath $source).Hash -cne $expected) {throw 'Исходная книга изменилась во время подготовки.'}
    [ordered]@{SourceSHA256=$expected;WorkbookSHA256=(Get-FileHash -LiteralPath $target).Hash;Setting=$setting;MigrationChecks=$checks;ProductionChanged=$false} |
        ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directoryPath 'Preparation.json') -Encoding UTF8
    Write-Output 'PRIVATE_FACTOR_READY: migration, placement, validation, hyperlink and border PASS.'
} finally {
    if ($book) {$book.Close($false)}
    if ($excel) {$excel.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
    if ($ownedPid) {
        $owned=Get-Process -Id $ownedPid -ErrorAction SilentlyContinue
        if ($owned -and $owned.ProcessName -eq 'EXCEL' -and $owned.StartTime.ToUniversalTime().Ticks -eq $ownedTicks) {Stop-Process -Id $ownedPid -Force}
    }
}
