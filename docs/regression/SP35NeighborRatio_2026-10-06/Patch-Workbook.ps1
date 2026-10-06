# Готовит копию пользовательской книги с исправленной геометрией области Ar.
# Не импортирует AutoCAD, не решает НДС и сохраняет текущие Config/Results.
# Добавляет только порог t/L, уточняет нумерацию рядов и обновляет справку.
$ErrorActionPreference = 'Stop'
$root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../../..')).Path
. (Join-Path $root 'tools/build_workbook/SettingsCatalog.ps1')
$source = Join-Path $root 'workbook/output/RC_Section_NDM.xlsm'
$directory = Join-Path $PSScriptRoot 'Publication'
New-Item -ItemType Directory -Path $directory -Force | Out-Null
$target = Join-Path $directory 'RC_Section_NDM.xlsm'
$backup = Join-Path $directory 'BeforePatch.xlsm'
$sourceHash = (Get-FileHash -LiteralPath $source).Hash
Copy-Item -LiteralPath $source -Destination $backup -Force
Copy-Item -LiteralPath $source -Destination $target -Force
$printAreas = @(Get-WorkbookPrintAreas $target)
$snapshots = @{}; $validations = @{}; $widths = @(); $configWidths = @(); $resultsSnapshot = $null
$insertIndex = 0; $rowToleranceIndex = 0
$excel = $null; $book = $null

# Проверяет исходные ячейки и списки Config с учетом одной новой строки.
# Разрешено только уточнение комментария RowTolerance; Results не меняется.
function Assert-Preserved([object]$Book) {
    $cells = 0
    foreach ($name in $snapshots.Keys) {
        $range = $Book.Names.Item($name).RefersToRange
        $actual = $range.Formula; $expected = $snapshots[$name]
        for ($r = 1; $r -le $expected.GetLength(0); $r++) {
            $actualRow = $r
            if ($name -eq 'rngSystemSettings' -and $r -ge $insertIndex) { $actualRow++ }
            for ($c = 1; $c -le $range.Columns.Count; $c++) {
                if ($name -eq 'rngSystemSettings' -and $r -eq $rowToleranceIndex -and $c -eq 4) { continue }
                if ([string]$actual[$actualRow,$c] -cne [string]$expected[$r,$c]) { throw "Config changed: $name row=$r column=$c" }
                $validationKey = "$name/$r/$c"
                if ($validations.ContainsKey($validationKey)) {
                    if ([string]$range.Cells.Item($actualRow,$c).Validation.Formula1 -cne $validations[$validationKey]) {
                        throw "Validation changed: $validationKey"
                    }
                }
                $cells++
            }
        }
    }
    for ($c = 1; $c -le $configWidths.Count; $c++) {
        if ($Book.Worksheets.Item('Config').Columns.Item($c).ColumnWidth -ne $configWidths[$c - 1]) { throw "Config width changed: column=$c" }
    }
    $settings = $Book.Names.Item('rngSystemSettings').RefersToRange
    if ($settings.Rows.Count -ne $snapshots['rngSystemSettings'].GetLength(0) + 1) { throw 'Settings range did not expand by one row.' }
    if ([string]$settings.Cells.Item($insertIndex,1).Value2 -ne 'SLS.Crack.SP35.NeighborRatioLimit' -or
        [double]$settings.Cells.Item($insertIndex,2).Value2 -ne 0.2) { throw 'New setting not saved.' }
    if ($settings.Cells.Item($insertIndex,5).Hyperlinks.Count -ne 1) { throw 'New setting help link missing.' }
    $sheet = $Book.Worksheets.Item('Results')
    for ($c = 1; $c -le $widths.Count; $c++) {
        if ($sheet.Columns.Item($c).ColumnWidth -ne $widths[$c - 1]) { throw "Results width changed: column=$c" }
    }
    $actual = $sheet.Range($resultsAddress).Formula
    for ($r = 1; $r -le $resultsSnapshot.GetLength(0); $r++) {
        for ($c = 1; $c -le $resultsSnapshot.GetLength(1); $c++) {
            if ([string]$actual[$r,$c] -cne [string]$resultsSnapshot[$r,$c]) { throw "Results changed: row=$r column=$c" }
        }
    }
    return $cells
}

# Выгружает сохраненный VBE-код книги для проверки фактического соответствия src/tests.
function Export-Vba([object]$Book, [string]$Path) {
    $writer = New-Object IO.StreamWriter($Path, $false, (New-Object Text.UTF8Encoding($true)))
    try {
        $writer.WriteLine('VBA PROJECT EXPORT')
        $writer.WriteLine("Workbook: $($Book.Name)")
        $writer.WriteLine("Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
        foreach ($component in $Book.VBProject.VBComponents) {
            $module = $component.CodeModule
            $writer.WriteLine(('=' * 100)); $writer.WriteLine("COMPONENT: $($component.Name)")
            $writer.WriteLine("TYPE: $($component.Type)"); $writer.WriteLine("LINES: $($module.CountOfLines)")
            $writer.WriteLine(('=' * 100))
            if ($module.CountOfLines) { $writer.WriteLine($module.Lines(1, $module.CountOfLines)) }
        }
    } finally { $writer.Dispose() }
}

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false; $excel.AutomationSecurity = 1
    $book = $excel.Workbooks.Open($target)
    foreach ($name in (Get-ConfigNamedRangeNames)) {
        $range = $book.Names.Item($name).RefersToRange
        $snapshots[$name] = $range.Formula
        for ($r = 1; $r -le $range.Rows.Count; $r++) {
            for ($c = 1; $c -le $range.Columns.Count; $c++) {
                try {
                    if ($range.Cells.Item($r,$c).Validation.Type -eq 3) {
                        $validations["$name/$r/$c"] = [string]$range.Cells.Item($r,$c).Validation.Formula1
                    }
                } catch { }
            }
        }
    }
    $config = $book.Worksheets.Item('Config')
    for ($c = 1; $c -le $config.UsedRange.Columns.Count; $c++) { $configWidths += $config.Columns.Item($c).ColumnWidth }
    $results = $book.Worksheets.Item('Results')
    $resultsAddress = $results.UsedRange.Address(); $resultsSnapshot = $results.UsedRange.Formula
    for ($c = 1; $c -le $results.UsedRange.Columns.Count; $c++) { $widths += $results.Columns.Item($c).ColumnWidth }
    foreach ($relative in @('src/Batch/CBatchSectionCalculator.cls', 'src/Crack/CCrackWidthCalculator.cls',
            'src/Crack/CSP35CrackData.cls', 'tests/modTestSP35CrackWidth.bas')) {
        $file = Get-Item -LiteralPath (Join-Path $root $relative)
        $component = $book.VBProject.VBComponents.Item($file.BaseName)
        $book.VBProject.VBComponents.Remove($component)
        $text = [IO.File]::ReadAllText($file.FullName, [Text.Encoding]::UTF8)
        $body = [regex]::Match($text, '(?ms)^Option Explicit.*').Value
        if ([string]::IsNullOrWhiteSpace($body)) { throw "Missing Option Explicit: $relative" }
        $type = 1; if ($file.Extension -eq '.cls') { $type = 2 }
        $component = $book.VBProject.VBComponents.Add($type); $component.Name = $file.BaseName
        $component.CodeModule.AddFromString(($body -split "`r?`n") -join "`r`n")
    }
    $settings = $book.Names.Item('rngSystemSettings').RefersToRange
    for ($r = 1; $r -le $settings.Rows.Count; $r++) {
        if ([string]$settings.Cells.Item($r,1).Value2 -eq 'SLS.Crack.SP35.NeighborRatioLimit') { throw 'Setting already present; rebase patch.' }
        if ([string]$settings.Cells.Item($r,1).Value2 -eq 'SLS.Crack.SP35.RowTolerance') { $rowToleranceIndex = $r }
    }
    if ($rowToleranceIndex -eq 0 -or $settings.Column -ne 1 -or $settings.Columns.Count -ne 5) { throw 'Unexpected settings layout.' }
    $insertIndex = $rowToleranceIndex + 1
    $insertRow = $settings.Row + $insertIndex - 1
    $startRow = $settings.Row; $lastRow = $settings.Row + $settings.Rows.Count - 1
    # Вставляем только A:E. Геометрия/единицы/сочетания справа не сдвигаются.
    $config.Range("A${insertRow}:E${insertRow}").Insert(-4121) | Out-Null
    $book.Names.Item('rngSystemSettings').RefersTo = "='Config'!`$A`$$startRow`:`$E`$$($lastRow + 1)"
    $settings = $book.Names.Item('rngSystemSettings').RefersToRange
    $newRow = $settings.Rows.Item($insertIndex)
    $settings.Rows.Item($rowToleranceIndex).Copy() | Out-Null
    $newRow.PasteSpecial(-4122) | Out-Null
    $newRow.Validation.Delete()
    foreach ($group in (Get-SystemSettingsCatalog)) {
        foreach ($entry in $group.Rows) {
            if ($entry[0] -eq 'SLS.Crack.SP35.RowTolerance') { $settings.Cells.Item($rowToleranceIndex,4).Value2 = [string]$entry[3] }
            if ($entry[0] -eq 'SLS.Crack.SP35.NeighborRatioLimit') {
                $newRow.Cells.Item(1,1).Value2 = [string]$entry[0]
                $newRow.Cells.Item(1,2).Value2 = 0.2
                $newRow.Cells.Item(1,3).Value2 = [string]$entry[2]
                $newRow.Cells.Item(1,4).Value2 = [string]$entry[3]
            }
        }
    }
    $newRow.Cells.Item(1,2).NumberFormat = '0.0###'
    $newRow.Cells.Item(1,2).HorizontalAlignment = -4108
    $newRow.Cells.Item(1,2).VerticalAlignment = -4108
    Add-SettingsInstructions $book $config ($book.Worksheets.Item('Справка'))
    for ($c = 1; $c -le $configWidths.Count; $c++) { $config.Columns.Item($c).ColumnWidth = $configWidths[$c - 1] }
    $preserved = Assert-Preserved $book
    $book.Save(); $book.Close($false); $book = $null
    Restore-WorkbookPrintAreas $target $printAreas
    $book = $excel.Workbooks.Open($target, 0, $true)
    $preserved = Assert-Preserved $book
    Export-Vba $book (Join-Path $directory 'VBA_All_Code.txt')
    if ((Get-FileHash -LiteralPath $source).Hash -ne $sourceHash) { throw 'User workbook changed while preparing repair.' }
    [pscustomobject]@{
        SourceSHA256=$sourceHash; SourceUnchanged=$true; PreparedSHA256=(Get-FileHash -LiteralPath $target).Hash
        ConfigCellsPreserved=$preserved; ConfigCommentsUpdated=1; ConfigSettingsAdded=1
        ConfigValidationListsPreserved=$validations.Count; ConfigWidthsPreserved=$configWidths.Count; ResultsAddress=$resultsAddress
        ResultsCellsPreserved=$resultsSnapshot.Length; ResultsWidthsPreserved=$widths.Count; StateSolveExecuted=$false
    } | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $directory 'Manifest.json') -Encoding UTF8
    Write-Output "PATCH_OK: Config=$preserved; Results=$($resultsSnapshot.Length); widths=$($widths.Count); no state solve"
} finally {
    if ($null -ne $book) { $book.Close($false) }
    if ($null -ne $excel) { $excel.Quit() }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
