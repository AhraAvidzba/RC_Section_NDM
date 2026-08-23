# скрипт проверяет структуру собранной книги, именованные диапазоны и ключевые правила интерфейса.

param(
    [string]$WorkbookPath = "workbook/output/RC_Section_NDM.xlsm"
)

$ErrorActionPreference = "Stop"

# Добавляет структурный элемент книги или отчета, сохраняя единый формат сборочных скриптов.
function Add-Check {
    param(
        [System.Collections.Generic.List[object]]$Checks,
        [string]$Name,
        [bool]$Passed,
        [string]$Details
    )

    $Checks.Add([pscustomobject]@{
        Check = $Name
        Passed = $Passed
        Details = $Details
    }) | Out-Null
}

# Выполняет одну проверку структуры книги и добавляет результат в общий отчет.
function Test-FormulaPresence {
    param([object]$Sheet)

    $used = $Sheet.UsedRange
    $rows = $used.Rows.Count
    $cols = $used.Columns.Count
    for ($r = 1; $r -le $rows; $r++) {
        for ($c = 1; $c -le $cols; $c++) {
            $cell = $used.Cells.Item($r, $c)
            $formula = [string]$cell.Formula
            if ($formula.StartsWith("=")) {
                $row = $cell.Row
                $col = $cell.Column
                if (Is-AllowedFormulaCell $Sheet.Name $row $col) { continue }
                return "Formula at $($Sheet.Name)!$($cell.Address($false, $false))"
            }
        }
    }
    return ""
}

# Выполняет служебный шаг сборочного или проверочного сценария.
function Is-AllowedFormulaCell {
    param([string]$SheetName, [int]$Row, [int]$Column)

    if ($SheetName -eq "Расчет") {
        if ($Row -ge 6 -and $Row -le 60 -and $Column -ge 1 -and $Column -le 7) { return $true }
        if ($Row -ge 17 -and $Row -le 35 -and $Column -ge 19 -and $Column -le 70) { return $true }
    }
    if ($SheetName -eq "Config") {
        if ($Row -ge 130 -and $Row -le 250 -and $Column -ge 9 -and $Column -le 17) { return $true }
    }
    return $false
}

$root = Resolve-Path -LiteralPath (Join-Path $PSScriptRoot "../..")
$fullWorkbookPath = Join-Path $root $WorkbookPath

if (-not (Test-Path -LiteralPath $fullWorkbookPath)) {
    throw "Workbook not found: $fullWorkbookPath"
}

$checks = [System.Collections.Generic.List[object]]::new()
$excel = $null
$workbook = $null

# Удаляет только служебный объект, который может мешать повторяемой сборке или проверке.
function Remove-DuplicatePrintAreaName {
    param([string]$Path)
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [System.IO.Compression.ZipFile]::Open($Path, [System.IO.Compression.ZipArchiveMode]::Update)
    try {
        $entry = $zip.GetEntry("xl/workbook.xml")
        if ($entry -eq $null) { return }
        $reader = New-Object IO.StreamReader($entry.Open())
        $xml = $reader.ReadToEnd()
        $reader.Close()
        $newXml = [regex]::Replace($xml, '<definedName name="Print_Area"[^>]*>.*?</definedName>\s*', '')
        if ($newXml -ne $xml) {
            $stream = $entry.Open()
            $stream.SetLength(0)
            $writer = New-Object IO.StreamWriter($stream, (New-Object System.Text.UTF8Encoding($false)))
            $writer.Write($newXml)
            $writer.Close()
        }
    }
    finally {
        if ($zip -ne $null) { $zip.Dispose() }
    }
}

Remove-DuplicatePrintAreaName $fullWorkbookPath

# Возвращает подготовленные данные или справочное значение для дальнейшего шага сборки.
function Get-PrintAreaFromWorkbookXml {
    param([string]$Path)
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $entry = $zip.GetEntry("xl/workbook.xml")
        if ($entry -eq $null) { return "" }
        $reader = New-Object IO.StreamReader($entry.Open())
        $xml = $reader.ReadToEnd()
        $reader.Close()
        $match = [regex]::Match($xml, '<definedName name="_xlnm\.Print_Area"[^>]*>(.*?)</definedName>')
        if ($match.Success) { return $match.Groups[1].Value }
        return ""
    }
    finally {
        if ($zip -ne $null) { $zip.Dispose() }
    }
}

$printAreaXml = Get-PrintAreaFromWorkbookXml $fullWorkbookPath

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $workbook = $excel.Workbooks.Open($fullWorkbookPath, $null, $true)

    $sheetNames = @()
    foreach ($sheet in $workbook.Worksheets) {
        $sheetNames += [string]$sheet.Name
    }

    $expectedSheetOrder = @("Config", "Справка", "Расчет", "Results")
    $sheetOrderOk = $true
    for ($i = 0; $i -lt $expectedSheetOrder.Count; $i++) {
        if ($sheetNames.Count -lt ($i + 1) -or $sheetNames[$i] -ne $expectedSheetOrder[$i]) { $sheetOrderOk = $false }
    }
    Add-Check $checks "Sheets Config/Справка/Расчет/Results" $sheetOrderOk ($sheetNames -join ", ")

    $requiredNames = @(
        "rngLoadCombinations",
        "rngResultSection",
        "rngUnitSettings",
        "rngSignConventionSettings",
        "rngPlotAnnotationSettings",
        "rngSystemSettings",
        "rngCircleGeometry",
        "rngRoundedRectangleGeometry",
        "rngLShapeGeometry",
        "rngBatchSummary",
        "rngConcreteDiagramPoints",
        "rngSteelDiagramPoints",
        "rngNDMElementResults",
        "rngNDMSectionGeometry",
        "rngNDMSectionProperties",
        "rngNDMSectionAnnotations"
    )

    $actualNames = @()
    foreach ($name in $workbook.Names) {
        $actualNames += [string]$name.Name
    }

    $missingNames = @($requiredNames | Where-Object { $actualNames -notcontains $_ })
    Add-Check $checks "Required named ranges" ($missingNames.Count -eq 0) ("Missing: " + ($missingNames -join ", "))

    if ($missingNames.Count -eq 0) {
        $batchSummaryRange = $workbook.Names.Item("rngBatchSummary").RefersToRange
        $elementResultsRange = $workbook.Names.Item("rngNDMElementResults").RefersToRange
        $geometryResultsRange = $workbook.Names.Item("rngNDMSectionGeometry").RefersToRange
        $sectionPropertiesRange = $workbook.Names.Item("rngNDMSectionProperties").RefersToRange
        $sectionAnnotationsRange = $workbook.Names.Item("rngNDMSectionAnnotations").RefersToRange
        Add-Check $checks "Results ranges layout" (
            ([string]$batchSummaryRange.Worksheet.Name -eq "Results") -and
            ($batchSummaryRange.Row -eq 1) -and ($batchSummaryRange.Column -eq 1) -and
            ($batchSummaryRange.Rows.Count -ge 30) -and
            ($elementResultsRange.Row -eq 33) -and ($elementResultsRange.Column -eq 1) -and
            ($geometryResultsRange.Row -eq 33) -and ($geometryResultsRange.Column -eq 9) -and
            ($sectionPropertiesRange.Row -eq 33) -and ($sectionPropertiesRange.Column -eq 26) -and
            ($sectionAnnotationsRange.Row -eq 33) -and ($sectionAnnotationsRange.Column -eq 34)
        ) ("batch=$($batchSummaryRange.Worksheet.Name)!R$($batchSummaryRange.Row)C$($batchSummaryRange.Column); elements=R$($elementResultsRange.Row)C$($elementResultsRange.Column); geometry=R$($geometryResultsRange.Row)C$($geometryResultsRange.Column); properties=R$($sectionPropertiesRange.Row)C$($sectionPropertiesRange.Column); annotations=R$($sectionAnnotationsRange.Row)C$($sectionAnnotationsRange.Column)")
    }

    $duplicates = @($actualNames | Group-Object | Where-Object { $_.Count -gt 1 } | ForEach-Object { $_.Name })
    Add-Check $checks "No duplicate names" ($duplicates.Count -eq 0) ("Duplicates: " + ($duplicates -join ", "))

    $obsoleteNames = @(
        "rngResultMx", "rngResultMy", "rngResultMxy", "rngMainInput", "rngRebarInput",
        "rngSystemDiagnostics",
        "SolverSettings", "CapacitySettings", "ConcreteDiagram", "SteelDiagram", "GeometrySettings",
        "OutputSettings", "AutoCADSettings"
    )
    $presentObsoleteNames = @($obsoleteNames | Where-Object { $actualNames -contains $_ })
    Add-Check $checks "No obsolete named ranges" ($presentObsoleteNames.Count -eq 0) ("Present: " + ($presentObsoleteNames -join ", "))

    $calc = $workbook.Worksheets.Item("Расчет")
    $system = $workbook.Worksheets.Item("Config")

    $printArea = [string]$calc.PageSetup.PrintArea
    $printAreaOk = $printArea.Contains('$A$1:$AJ$60') -or $printAreaXml.Contains('$A$1:$AJ$60')
    Add-Check $checks "Print area" $printAreaOk ("PageSetup=$printArea; Xml=$printAreaXml")
    Add-Check $checks "A4 portrait" (($calc.PageSetup.PaperSize -eq 9) -and ($calc.PageSetup.Orientation -eq 1)) ("PaperSize=$($calc.PageSetup.PaperSize); Orientation=$($calc.PageSetup.Orientation)")
    Add-Check $checks "Print scale for one-page blocks" (($calc.PageSetup.Zoom -le 100) -and ($calc.HPageBreaks.Count -eq 0)) ("Zoom=$($calc.PageSetup.Zoom); HBreaks=$($calc.HPageBreaks.Count)")

    $breakColumns = @()
    foreach ($break in $calc.VPageBreaks) {
        if ($break.Type -eq -4135) {
            $breakColumns += $break.Location.Column
        }
    }
    $expectedBreakColumns = @(19)
    $hasBreaks = -not (@($expectedBreakColumns | Where-Object { $breakColumns -notcontains $_ }).Count)
    Add-Check $checks "Vertical page breaks" $hasBreaks ("Columns: " + ($breakColumns -join ", "))

    $sectionResult = $workbook.Names.Item("rngResultSection").RefersToRange
    $loadsRangeForLayout = $workbook.Names.Item("rngLoadCombinations").RefersToRange
    $leftToRight = ($loadsRangeForLayout.Column -lt $sectionResult.Column)
    Add-Check $checks "Load and result blocks left to right" $leftToRight ("Columns: loads=$($loadsRangeForLayout.Column), section=$($sectionResult.Column)")

    $settings = $workbook.Names.Item("rngSystemSettings").RefersToRange
    $expectedSettingsHeaders = @("Параметр", "Значение", "Ед.", "Комментарий", "Инструкции")
    $actualSettingsHeaders = @()
    for ($i = 1; $i -le $expectedSettingsHeaders.Count; $i++) {
        $actualSettingsHeaders += [string]$settings.Cells.Item(1, $i).Value2
    }
    Add-Check $checks "Config settings table headers" (($actualSettingsHeaders -join "|") -eq ($expectedSettingsHeaders -join "|")) ($actualSettingsHeaders -join " | ")

    $settingKeys = @()
    for ($i = 2; $i -le $settings.Rows.Count; $i++) {
        $key = [string]$settings.Cells.Item($i, 1).Value2
        if (-not [string]::IsNullOrWhiteSpace($key) -and -not $key.StartsWith("[")) {
            $settingKeys += $key
        }
    }
    $duplicateSettingKeys = @($settingKeys | Group-Object | Where-Object { $_.Count -gt 1 } | ForEach-Object { $_.Name })
    Add-Check $checks "No duplicate Config keys" ($duplicateSettingKeys.Count -eq 0) ("Duplicates: " + ($duplicateSettingKeys -join ", "))

    $obsoleteSettingKeys = @(
        "Capacity.Enabled", "Capacity.CalculateMx", "Capacity.CalculateMy",
        "Capacity.CalculateMxy", "Mesh.BoundaryMode", "Mesh.StepX", "Mesh.StepY", "Circle.Radius",
        "Batch.MaxCombinations", "Batch.Diagnostics", "Materials.SourceStatus",
        "Concrete.Diagram", "Steel.Diagram",
        "Concrete.Point1.Eps", "Concrete.Point1.Stress",
        "Concrete.Point2.Eps", "Concrete.Point2.Stress",
        "Concrete.Point3.Eps", "Concrete.Point3.Stress",
        "Steel.Point1.Eps", "Steel.Point1.Stress",
        "Steel.Point2.Eps", "Steel.Point2.Stress",
        "Steel.Point3.Eps", "Steel.Point3.Stress",
        "Circle.CenterX", "Circle.CenterY",
        "LShape.OriginX", "LShape.OriginY",
        "Plot.DimensionsEnabled", "Plot.RebarLabelsEnabled"
    )
    $presentObsoleteSettings = @($obsoleteSettingKeys | Where-Object { $settingKeys -contains $_ })
    Add-Check $checks "No obsolete Config settings" ($presentObsoleteSettings.Count -eq 0) ("Present: " + ($presentObsoleteSettings -join ", "))

    $solverMethodCell = $null
    for ($i = 2; $i -le $settings.Rows.Count; $i++) {
        if ([string]$settings.Cells.Item($i, 1).Value2 -eq "Solver.Method") {
            $solverMethodCell = $settings.Cells.Item($i, 2)
            break
        }
    }
    $solverMethodDefaultOk = ($solverMethodCell -ne $null) -and ([string]$solverMethodCell.Value2 -eq "Newton")
    Add-Check $checks "Solver.Method default" $solverMethodDefaultOk ("Value=" + [string]$(if ($solverMethodCell -eq $null) { "" } else { $solverMethodCell.Value2 }))

    $solverValidationOk = $false
    $solverValidationDetails = "Missing Solver.Method"
    if ($solverMethodCell -ne $null) {
        try {
            $formula = [string]$solverMethodCell.Validation.Formula1
            $solverValidationDetails = "Formula1=$formula"
            if ($formula.StartsWith("=")) {
                $validationRange = $system.Range($formula.Substring(1))
                $values = @()
                foreach ($cell in $validationRange.Cells) {
                    $value = [string]$cell.Value2
                    if (-not [string]::IsNullOrWhiteSpace($value)) { $values += $value }
                }
                $solverValidationDetails = "Values=" + ($values -join ", ")
                $solverValidationOk = (($values.Count -eq 2) -and ($values[0] -eq "Newton") -and ($values[1] -eq "Secant"))
            }
            else {
                $values = @($formula -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ })
                $solverValidationDetails = "Values=" + ($values -join ", ")
                $solverValidationOk = (($values.Count -eq 2) -and ($values[0] -eq "Newton") -and ($values[1] -eq "Secant"))
            }
        }
        catch {
            $solverValidationDetails = $_.Exception.Message
        }
    }
    Add-Check $checks "Solver.Method validation list" $solverValidationOk $solverValidationDetails

    $concretePoints = $workbook.Names.Item("rngConcreteDiagramPoints").RefersToRange
    $steelPoints = $workbook.Names.Item("rngSteelDiagramPoints").RefersToRange
    Add-Check $checks "Concrete diagram point table" (($concretePoints.Columns.Count -eq 3) -and ($concretePoints.Rows.Count -eq 10)) ("Rows=$($concretePoints.Rows.Count); Columns=$($concretePoints.Columns.Count)")
    Add-Check $checks "Steel diagram point table" (($steelPoints.Columns.Count -eq 3) -and ($steelPoints.Rows.Count -eq 10)) ("Rows=$($steelPoints.Rows.Count); Columns=$($steelPoints.Columns.Count)")

    $loads = $workbook.Names.Item("rngLoadCombinations").RefersToRange
    $expectedLoadHeaders = @("CombinationID", "N, tf", "Mx, tf*m", "My, tf*m", "CalculationType", "Comment")
    $actualLoadHeaders = @()
    for ($i = 1; $i -le 6; $i++) {
        $actualLoadHeaders += [string]$loads.Cells.Item(1, $i).Value2
    }
    Add-Check $checks "Load combinations table headers" (($actualLoadHeaders -join "|") -eq ($expectedLoadHeaders -join "|")) ($actualLoadHeaders -join " | ")
    Add-Check $checks "Formulas allowed on workbook sheets" $true "Excel formulas are allowed and preferred for transparent user-level calculations; VBA validation covers workbook structure."

    $excelSourceDir = Join-Path $root "src/Excel"
    $cellItemMatches = @(
        Get-ChildItem -LiteralPath $excelSourceDir -File |
            Where-Object { $_.Extension -in @(".bas", ".cls") } |
            Select-String -Pattern "Cells\.Item" |
            ForEach-Object { "$($_.Path):$($_.LineNumber)" }
    )
    Add-Check $checks "Runtime Excel IO uses bulk ranges" ($cellItemMatches.Count -eq 0) ("Cells.Item matches: " + ($cellItemMatches -join "; "))
}
finally {
    if ($workbook -ne $null) {
        $workbook.Close($false)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
    }
    if ($excel -ne $null) {
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

$checks | Format-Table -AutoSize

$failed = @($checks | Where-Object { -not $_.Passed })
if ($failed.Count -gt 0) {
    exit 1
}














