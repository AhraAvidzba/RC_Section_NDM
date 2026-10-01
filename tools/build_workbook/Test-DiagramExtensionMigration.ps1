# Проверяет точечную миграцию Config на независимых книгах, без изменения
# пользовательской output-книги. Сохраняет фикстуры и отчет для воспроизведения.
param([string]$ReportPath = "docs/regression/Audit02/diagram_extension_migration_2026-10-01.txt")
$ErrorActionPreference = "Stop"
. (Join-Path $PSScriptRoot "SettingsCatalog.ps1")
$root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot "../..")).Path
$fixtureRoot = Join-Path ([IO.Path]::GetTempPath()) ("RC_NDM_ExtensionMigration_" + [Guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $fixtureRoot | Out-Null
$report = New-Object System.Collections.Generic.List[string]
$script:passed = 0
$excel = $null
$workbook = $null

# Прерывает тест при первом нарушении; формирует проверяемый журнал assertions.
function Assert-Migration {
    param([string]$Label, [bool]$Condition)
    if (-not $Condition) { throw "FAIL: $Label" }
    $script:passed++
    $report.Add("OK: $Label")
}

# Создает минимум реального Config с общим блоком, Solver и соседней таблицей.
# Не использует defaults генератора, чтобы old No проверялось до их подстановки.
function New-MigrationFixture {
    param([object]$Application, [object]$Case)
    $book = $Application.Workbooks.Add(-4167)
    $config = $book.Worksheets.Item(1)
    $config.Name = "Config"
    $guide = $book.Worksheets.Add()
    $guide.Name = "Справка"
    $guide.Cells.Item(1, 1).Value2 = "Существующая справка"
    $guide.Cells.Item(2, 1).Value2 = "Соседний блок не должен измениться."
    $data = New-Object 'object[,]' 15, 5
    $data[0, 0] = "Параметр"; $data[0, 1] = "Значение"
    $data[1, 0] = "[Общие]"
    $data[2, 0] = "General.ExecutionReportEnabled"; $data[2, 1] = "No"
    $data[5, 0] = "[Solver]"
    $data[6, 0] = "Solver.MaxIterations"; $data[6, 1] = "37"
    $data[7, 0] = "Solver.Method"; $data[7, 1] = "Secant"
    if ($Case.FullGeneral) {
        $data[3, 0] = "General.UnchangedA"; $data[3, 1] = "A"
        $data[4, 0] = "General.UnchangedB"; $data[4, 1] = "B"
    }
    if ($Case.HasLegacy) {
        $data[10, 0] = "Solver.DirectState.DiagramExtension"
        $data[10, 1] = $Case.Legacy
    }
    if ($Case.HasCanonical) {
        $index = 3
        if ($Case.CanonicalOutside) { $index = 11 }
        $data[$index, 0] = "General.DiagramExtension"
        $data[$index, 1] = $Case.Canonical
    }
    $config.Range("A3:E17").Value2 = $data
    $config.Range("H14").Formula = "=2+3"
    $config.PageSetup.PrintArea = '$A$1:$E$17'
    $book.Names.Add("rngSystemSettings", "=Config!`$A`$3:`$E`$17") | Out-Null
    $book.Names.Add("rngMigrationNeighbor", "=Config!`$H`$14") | Out-Null
    return $book
}

# Читает видимое значение настройки из фактического именованного диапазона.
function Get-MigrationSetting {
    param([object]$Book, [string]$Key)
    $range = $Book.Names.Item("rngSystemSettings").RefersToRange
    $data = $range.Value2
    $matches = @()
    for ($r = 2; $r -le $range.Rows.Count; $r++) {
        if ([string]$data[$r, 1] -eq $Key) {
            $matches += @{ Value = [string]$data[$r, 2]; Row = $range.Row + $r - 1 }
        }
    }
    return ,$matches
}

# Снимок нужен для проверки полного отсутствия изменений при ошибке ввода
# и идемпотентности успешного повторного запуска, включая порядок строк.
function Get-MigrationSnapshot {
    param([object]$Book)
    $parts = New-Object System.Collections.Generic.List[string]
    $range = $Book.Names.Item("rngSystemSettings").RefersToRange
    $parts.Add([string]$Book.Names.Item("rngSystemSettings").RefersTo)
    $data = $range.Value2
    for ($r = 1; $r -le $range.Rows.Count; $r++) {
        for ($c = 1; $c -le $range.Columns.Count; $c++) { $parts.Add([string]($data[$r, $c])) }
    }
    $guide = $Book.Worksheets.Item("Справка").UsedRange
    $parts.Add([string]$guide.Address())
    $parts.Add([string]$Book.Names.Item("rngMigrationNeighbor").RefersTo)
    return $parts -join "|"
}

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $cases = @(
        @{ Name="oldNo"; HasLegacy=$true; Legacy="No"; Expected="No" },
        @{ Name="oldYes"; HasLegacy=$true; Legacy="Yes"; Expected="Yes" },
        @{ Name="newNo"; HasCanonical=$true; Canonical="No"; Expected="No" },
        @{ Name="newYes"; HasCanonical=$true; Canonical="Yes"; Expected="Yes" },
        @{ Name="conflict"; HasLegacy=$true; Legacy="No"; HasCanonical=$true; Canonical="Yes"; Expected="Yes"; Warning=$true },
        @{ Name="equivalent"; HasLegacy=$true; Legacy="1"; HasCanonical=$true; Canonical="Yes"; Expected="Yes" },
        @{ Name="missing"; Expected="Yes" },
        @{ Name="invalidNew"; HasLegacy=$true; Legacy="Yes"; HasCanonical=$true; Canonical="invalid"; Invalid=$true },
        @{ Name="emptyNew"; HasLegacy=$true; Legacy="Yes"; HasCanonical=$true; Canonical=""; Invalid=$true },
        @{ Name="invalidOld"; HasLegacy=$true; Legacy="invalid"; Invalid=$true },
        @{ Name="invalidOldWithNew"; HasLegacy=$true; Legacy="invalid"; HasCanonical=$true; Canonical="No"; Expected="No"; Warning=$true },
        @{ Name="insertGeneral"; HasLegacy=$true; Legacy="No"; FullGeneral=$true; Expected="No" },
        @{ Name="moveNew"; HasLegacy=$true; Legacy="Yes"; HasCanonical=$true; Canonical="No"; CanonicalOutside=$true; Expected="No"; Warning=$true }
    )
    foreach ($case in $cases) {
        $workbook = New-MigrationFixture $excel $case
        $path = Join-Path $fixtureRoot ($case.Name + ".xlsx")
        $workbook.SaveAs($path, 51)
        $workbook.Close($false)
        $printAreas = @(Get-WorkbookPrintAreas $path)
        Assert-Migration "migration.$($case.Name).initialPrintArea" ($printAreas.Count -eq 1)
        $workbook = $excel.Workbooks.Open($path)
        $before = Get-MigrationSnapshot $workbook
        $errorText = ""
        $result = $null
        try { $result = Invoke-DiagramExtensionMigration $workbook }
        catch { $errorText = $_.Exception.Message }
        $label = "migration.$($case.Name)"
        if ($case.Invalid) {
            Assert-Migration "$label.invalidRejected" ($errorText -like "*должна быть Yes или No*")
            Assert-Migration "$label.unchanged" ((Get-MigrationSnapshot $workbook) -eq $before)
            $workbook.Close($false)
            $workbook = $null
            continue
        }
        if ($errorText) { throw "$label failed: $errorText" }
        Assert-Migration "$label.value" ($result.Value -eq $case.Expected)
        Assert-Migration "$label.warning" ([bool]$result.Warning -eq [bool]$case.Warning)
        $canonical = Get-MigrationSetting $workbook "General.DiagramExtension"
        Assert-Migration "$label.oneCanonical" ($canonical.Count -eq 1)
        Assert-Migration "$label.noLegacy" ((Get-MigrationSetting $workbook "Solver.DirectState.DiagramExtension").Count -eq 0)
        Assert-Migration "$label.generalBlock" ($canonical[0].Row -gt 4 -and $canonical[0].Row -le 8)
        Assert-Migration "$label.neighborIterations" ((Get-MigrationSetting $workbook "Solver.MaxIterations")[0].Value -eq "37")
        Assert-Migration "$label.neighborMethod" ((Get-MigrationSetting $workbook "Solver.Method")[0].Value -eq "Secant")
        Assert-Migration "$label.neighborGeneral" ((Get-MigrationSetting $workbook "General.ExecutionReportEnabled")[0].Value -eq "No")
        $neighbor = $workbook.Names.Item("rngMigrationNeighbor").RefersToRange
        Assert-Migration "$label.rightTable" ($neighbor.Formula -eq "=2+3" -and $neighbor.Address() -eq '$H$14')
        $input = $workbook.Worksheets.Item("Config").Cells.Item($result.Row, 2)
        Assert-Migration "$label.validation" ($input.Validation.Type -eq 3 -and $input.Validation.InCellDropdown -and -not $input.Validation.IgnoreBlank)
        Assert-Migration "$label.alignment" ($input.HorizontalAlignment -eq -4108 -and $input.VerticalAlignment -eq -4108)
        Assert-Migration "$label.help" ($workbook.Worksheets.Item("Config").Cells.Item($result.Row, 5).Hyperlinks.Count -eq 1)
        Assert-Migration "$label.guideNeighbor" ($workbook.Worksheets.Item("Справка").Cells.Item(2, 1).Value2 -eq "Соседний блок не должен измениться.")
        $stable = Get-MigrationSnapshot $workbook
        $repeat = Invoke-DiagramExtensionMigration $workbook
        Assert-Migration "$label.idempotent" ((Get-MigrationSnapshot $workbook) -eq $stable -and $repeat.Row -eq $result.Row -and -not $repeat.Warning)
        $workbook.Save()
        $workbook.Close($false)
        Restore-WorkbookPrintAreas $path $printAreas
        $savedAreas = @(Get-WorkbookPrintAreas $path)
        Assert-Migration "$label.printAreaPreserved" ($savedAreas.Count -eq 1 -and $savedAreas[0].Xml.Contains('$A$1:$E$17'))
        $workbook = $excel.Workbooks.Open($path)
        Assert-Migration "$label.saveReopen" ((Get-MigrationSnapshot $workbook) -eq $stable)
        $workbook.Close($false)
        $workbook = $null
    }
    $report.Add("TOTAL_MIGRATION: passed=$script:passed; failed=0")
    $report.Add("Fixtures: $fixtureRoot")
} catch {
    $report.Add("$($_.Exception.Message)")
    $report.Add("TOTAL_MIGRATION: passed=$script:passed; failed=1")
    throw
} finally {
    if ($workbook) { $workbook.Close($false); [Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null }
    if ($excel) { $excel.Quit(); [Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    $fullReport = Join-Path $root $ReportPath
    New-Item -ItemType Directory -Force -Path (Split-Path -Parent $fullReport) | Out-Null
    $report | Set-Content -LiteralPath $fullReport -Encoding UTF8
    $report
}
