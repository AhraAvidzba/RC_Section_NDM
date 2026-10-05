# Адресная приемка независимых селекторов RectSet в сохраненной книге.
# Проверяет фактические ячейки, справку и перенос объединенного ввода на две
# стороны. Все экспериментальные изменения остаются в read-only COM-копии;
# исходная книга не сохраняется, ее SHA256 должен остаться неизменным.
param(
    [Parameter(Mandatory=$true)][string]$WorkbookPath,
    [Parameter(Mandatory=$true)][string]$ReportPath,
    [Parameter(Mandatory=$true)][string]$PreviewPath
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'SettingsCatalog.ps1')
$root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$source = (Resolve-Path -LiteralPath $WorkbookPath).Path
$report = Join-Path $root $ReportPath
$preview = Join-Path $root $PreviewPath
$beforeHash = (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash
$lines = New-Object 'System.Collections.Generic.List[string]'
$script:failed = 0
$excel = $null
$book = $null

# ДЛЯ ТЕСТОВ: фиксирует проверяемый контракт, не подменяя отказ успехом.
function Assert-RectSet([string]$Name, [bool]$Passed) {
    if (-not $Passed) { $script:failed++ }
    $lines.Add("RECTSET|$Name|passed=$Passed")
}

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 3
    $book = $excel.Workbooks.Open($source, 0, $true)
    if (-not $book.ReadOnly) { throw 'Проверяемая книга должна быть read-only.' }
    $target = $book.Names.Item('rngRectSetGeometry').RefersToRange
    foreach ($row in 21..28) {
        foreach ($column in @(3, 4, 6, 7)) {
            $cell = $target.Cells.Item($row, $column)
            Assert-RectSet ("unmerged.$row.$column") (-not $cell.MergeCells)
            Assert-RectSet ("dropdown.$row.$column") ($cell.Validation.Type -eq 3 -and $cell.Validation.InCellDropdown -and ([string]$cell.Validation.Formula1).StartsWith('='))
            Assert-RectSet ("centered.$row.$column") ($cell.HorizontalAlignment -eq -4108 -and $cell.VerticalAlignment -eq -4108)
            Assert-RectSet ("bottomBorder.$row.$column") ($cell.Borders.Item(9).LineStyle -eq -4115 -and $cell.Borders.Item(9).Weight -eq 2)
        }
    }
    $guide = $book.Worksheets.Item('Справка').UsedRange.Value2
    $body = ($guide | ForEach-Object { [string]$_ }) -join ' '
    Assert-RectSet 'guide.independent' ($body.Contains('Выбор одной грани не изменяет противоположную'))
    Assert-RectSet 'guide.noSharedContract' (-not $body.Contains('общие для двух сторон каждой грани'))

    # Печатаем только измененный блок Config, не меняя сам output-файл.
    $sheet = $target.Worksheet
    $sheet.PageSetup.PrintArea = $target.Address()
    $sheet.PageSetup.Orientation = 2
    $sheet.PageSetup.PaperSize = 9
    $sheet.PageSetup.Zoom = $false
    $sheet.PageSetup.FitToPagesWide = 1
    $sheet.PageSetup.FitToPagesTall = 1
    $sheet.ExportAsFixedFormat(0, $preview)
    $lines.Add("PREVIEW: $preview")

    # Противоположные стороны уже независимой таблицы сохраняют разный выбор.
    foreach ($firstRow in @(21, 23, 25, 27)) {
        foreach ($column in @(3, 4, 6, 7)) {
            if ($column -in @(3, 6)) { $choices = @('Stacked', 'SideBySide') }
            else { $choices = @('EachBar', 'EverySecondBar') }
            $target.Cells.Item($firstRow, $column).Value2 = $choices[0]
            $target.Cells.Item(($firstRow + 1), $column).Value2 = $choices[1]
        }
    }
    $snapshot = ConvertTo-Json -InputObject $target.Formula -Depth 4 -Compress
    $records = @(Set-RectSetIndependentSelectorLayout $target)
    Assert-RectSet 'migration.independentValuesPreserved' ($snapshot -ceq (ConvertTo-Json -InputObject $target.Formula -Depth 4 -Compress))
    Assert-RectSet 'migration.noNewMerges' (@($records | Where-Object WasMerged).Count -eq 0)

    # Отдельная временная таблица моделирует прежние объединения и validation.
    $fixtureSheet = $book.Worksheets.Add()
    $fixture = $fixtureSheet.Range('A1:I28')
    $fixtureSheet.Range('L1').Value2 = 'Stacked'
    $fixtureSheet.Range('L2').Value2 = 'SideBySide'
    $fixtureSheet.Range('M1').Value2 = 'EachBar'
    $fixtureSheet.Range('M2').Value2 = 'EverySecondBar'
    foreach ($firstRow in @(21, 23, 25, 27)) {
        foreach ($column in @(3, 4, 6, 7)) {
            $first = $fixture.Cells.Item($firstRow, $column)
            if ($column -in @(3, 6)) { $value = 'SideBySide'; $list = '=$L$1:$L$2' }
            else { $value = 'EverySecondBar'; $list = '=$M$1:$M$2' }
            $first.Value2 = $value
            $first.Validation.Add(3, 1, 1, $list)
            $first.Validation.InCellDropdown = $true
            $first.Resize(2, 1).Merge()
        }
    }
    $records = @(Set-RectSetIndependentSelectorLayout $fixture)
    Assert-RectSet 'migration.sixteenMergedPairs' (@($records | Where-Object WasMerged).Count -eq 16)
    foreach ($record in $records) {
        $first = $fixtureSheet.Range($record.SharedAddress)
        $second = $fixtureSheet.Range($record.Address)
        Assert-RectSet ('migration.preserved.' + $record.Address) ([string]$first.Formula -ceq [string]$record.SharedFormula -and [string]$second.Formula -ceq [string]$record.SharedFormula)
        Assert-RectSet ('migration.split.' + $record.Address) (-not $first.MergeCells -and -not $second.MergeCells)
        Assert-RectSet ('migration.validation.' + $record.Address) ($second.Validation.Type -eq 3 -and $second.Validation.InCellDropdown -and $second.Validation.Formula1 -ceq $first.Validation.Formula1)
    }
    $snapshot = ConvertTo-Json -InputObject $fixture.Formula -Depth 4 -Compress
    Set-RectSetIndependentSelectorLayout $fixture | Out-Null
    Assert-RectSet 'migration.idempotent' ($snapshot -ceq (ConvertTo-Json -InputObject $fixture.Formula -Depth 4 -Compress))
} catch {
    $script:failed++
    $lines.Add("SCRIPT ERROR: $($_.Exception.Message); $($_.ScriptStackTrace)")
} finally {
    if ($book) { $book.Close($false) }
    if ($excel) { $excel.Quit(); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) | Out-Null }
    $afterHash = (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash
    Assert-RectSet 'sourceUnchanged' ($beforeHash -eq $afterHash)
    $lines.Add("TOTAL_RECTSET_PRESENTATION_MIGRATION: failed=$script:failed; sourceSHA256=$afterHash")
    $lines | Set-Content -LiteralPath $report -Encoding UTF8
    $lines
}
if ($script:failed -gt 0) { exit 1 }
