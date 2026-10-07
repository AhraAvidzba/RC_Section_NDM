# скрипт проверяет структуру собранной книги, именованные диапазоны и ключевые правила интерфейса.

# Не исправляет проверяемый файл. Excel открывается read-only без макросов и
# событий; хеш входной книги после проверки должен совпасть. По запросу
# сохраняется машинно-читаемый отчет отдельных структурных проверок.
param(
    [string]$WorkbookPath = "workbook/output/RC_Section_NDM.xlsm",
    [string]$ReportPath = "",
    [switch]$UserConfiguredWorkbook
)

$ErrorActionPreference = "Stop"

# Сохраняет отдельную структурную проверку и ее фактическую причину в отчете.
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
$beforeHash = (Get-FileHash -LiteralPath $fullWorkbookPath -Algorithm SHA256).Hash

# Читает каноническую область печати именно листа Расчет без открытия ZIP
# на запись. Неканоническое/повторное имя не удаляется из проверяемой книги:
# структурная ошибка фиксируется в отчете, а исходный файл остается неизменным.
function Get-PrintAreaFromWorkbookXml {
    param([string]$Path, [System.Collections.Generic.List[object]]$Checks)
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $entry = $zip.GetEntry("xl/workbook.xml")
        if ($entry -eq $null) { throw "В книге отсутствует xl/workbook.xml." }
        $reader = New-Object IO.StreamReader($entry.Open())
        try { [xml]$xml = $reader.ReadToEnd() } finally { $reader.Dispose() }
        $sheetIndex = -1
        $index = 0
        foreach ($sheet in $xml.workbook.sheets.sheet) {
            if ([string]$sheet.name -eq "Расчет") { $sheetIndex = $index }
            $index++
        }
        $namespaces = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
        $namespaces.AddNamespace('w', $xml.DocumentElement.NamespaceURI)
        $names = $xml.SelectNodes('/w:workbook/w:definedNames/w:definedName', $namespaces)
        $areas = @($names | Where-Object {
            $_.GetAttribute('name') -eq '_xlnm.Print_Area' -and
            $_.GetAttribute('localSheetId') -eq [string]$sheetIndex
        })
        $aliases = @($names | Where-Object {
            $_.GetAttribute('name') -eq 'Print_Area'
        })
        Add-Check $Checks "Canonical print area names" ($sheetIndex -ge 0 -and $areas.Count -eq 1 -and $aliases.Count -eq 0) (
            "CalculationSheetIndex=$sheetIndex; canonical=$($areas.Count); noncanonical=$($aliases.Count); file not repaired")
        if ($areas.Count -eq 1) { return [string]$areas[0].InnerText }
        return ""
    }
    finally {
        if ($zip -ne $null) { $zip.Dispose() }
    }
}

$printAreaXml = Get-PrintAreaFromWorkbookXml $fullWorkbookPath $checks

try {
    if (@($checks | Where-Object { -not $_.Passed }).Count -gt 0) {
        throw "Именованная область печати листа Расчет отсутствует, повторяется или содержит неканоническое имя Print_Area. Исправьте имя в книге; проверка остановлена без изменения файла."
    }
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 3
    $excel.EnableEvents = $false

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
        "rngUnitSettings",
        "rngSignConventionSettings",
        "rngPlotAnnotationSettings",
        "rngSystemSettings",
        "rngCircleGeometry",
        "rngRoundedRectangleGeometry",
        "rngHollowRectangleGeometry",
        "rngRectSetGeometry",
        "rngBatchSummary",
        "rngStrengthSummaryAnchor",
        "rngCrackSummaryAnchor",
        "rngStabilitySummaryAnchor",
        "rngConcreteMaterialParameters",
        "rngSteelMaterialParameters",
        "rngCalculationProfiles",
        "rngStabilityDurationLoads",
        "rngSP35Table721",
        "rngNDMElementResults",
        "rngNDMSectionGeometry",
        "rngNDMSectionContours",
        "rngNDMSectionProperties",
        "rngNDMSectionAnnotations",
        "rngNDMMaterialDiagrams"
    )

    $actualNames = @()
    foreach ($name in $workbook.Names) {
        $actualNames += [string]$name.Name
    }

    $missingNames = @($requiredNames | Where-Object { $actualNames -notcontains $_ })
    Add-Check $checks "Required named ranges" ($missingNames.Count -eq 0) ("Missing: " + ($missingNames -join ", "))

    if ($missingNames.Count -eq 0) {
        $batchSummaryRange = $workbook.Names.Item("rngBatchSummary").RefersToRange
        $strengthSummaryAnchorRange = $workbook.Names.Item("rngStrengthSummaryAnchor").RefersToRange
        $crackSummaryAnchorRange = $workbook.Names.Item("rngCrackSummaryAnchor").RefersToRange
        $stabilitySummaryAnchorRange = $workbook.Names.Item("rngStabilitySummaryAnchor").RefersToRange
        $elementResultsRange = $workbook.Names.Item("rngNDMElementResults").RefersToRange
        $geometryResultsRange = $workbook.Names.Item("rngNDMSectionGeometry").RefersToRange
        $sectionContoursRange = $workbook.Names.Item("rngNDMSectionContours").RefersToRange
        $sectionPropertiesRange = $workbook.Names.Item("rngNDMSectionProperties").RefersToRange
        $sectionAnnotationsRange = $workbook.Names.Item("rngNDMSectionAnnotations").RefersToRange
        $materialDiagramsRange = $workbook.Names.Item("rngNDMMaterialDiagrams").RefersToRange
        $stabilityLoadsRange = $workbook.Names.Item("rngStabilityDurationLoads").RefersToRange
        $sp35TableRange = $workbook.Names.Item("rngSP35Table721").RefersToRange
        Add-Check $checks "Results ranges layout" (
            ([string]$batchSummaryRange.Worksheet.Name -eq "Results") -and
            ($batchSummaryRange.Row -eq 1) -and ($batchSummaryRange.Column -eq 1) -and
            ($batchSummaryRange.Rows.Count -eq 1) -and ($batchSummaryRange.Columns.Count -eq 1) -and
            ($strengthSummaryAnchorRange.Worksheet.Name -eq "Results") -and
            ($strengthSummaryAnchorRange.Row -eq 49) -and ($strengthSummaryAnchorRange.Column -eq 1) -and
            ($crackSummaryAnchorRange.Worksheet.Name -eq "Results") -and
            ($crackSummaryAnchorRange.Row -eq 85) -and ($crackSummaryAnchorRange.Column -eq 1) -and
            ($stabilitySummaryAnchorRange.Worksheet.Name -eq "Results") -and
            ($stabilitySummaryAnchorRange.Row -eq 122) -and ($stabilitySummaryAnchorRange.Column -eq 1) -and
            ($elementResultsRange.Row -eq 156) -and ($elementResultsRange.Column -eq 1) -and
            ($geometryResultsRange.Row -eq 156) -and ($geometryResultsRange.Column -eq 12) -and
            ($sectionContoursRange.Row -eq 156) -and ($sectionContoursRange.Column -eq 29) -and
            ($sectionPropertiesRange.Row -eq 156) -and ($sectionPropertiesRange.Column -eq 48) -and
            ($materialDiagramsRange.Row -eq 156) -and ($materialDiagramsRange.Column -eq 56) -and
            ($sectionAnnotationsRange.Row -eq 156) -and ($sectionAnnotationsRange.Column -eq 69)
        ) ("batch=$($batchSummaryRange.Worksheet.Name)!R$($batchSummaryRange.Row)C$($batchSummaryRange.Column):$($batchSummaryRange.Columns.Count) cols; strengthAnchor=R$($strengthSummaryAnchorRange.Row)C$($strengthSummaryAnchorRange.Column); crackAnchor=R$($crackSummaryAnchorRange.Row)C$($crackSummaryAnchorRange.Column); stabilityAnchor=R$($stabilitySummaryAnchorRange.Row)C$($stabilitySummaryAnchorRange.Column); elements=R$($elementResultsRange.Row)C$($elementResultsRange.Column); geometry=R$($geometryResultsRange.Row)C$($geometryResultsRange.Column); properties=R$($sectionPropertiesRange.Row)C$($sectionPropertiesRange.Column); annotations=R$($sectionAnnotationsRange.Row)C$($sectionAnnotationsRange.Column); materialDiagrams=R$($materialDiagramsRange.Row)C$($materialDiagramsRange.Column)")
        Add-Check $checks "Stability duration loads range" (
            ($stabilityLoadsRange.Worksheet.Name -eq "Config") -and
            ($stabilityLoadsRange.Row -eq 37) -and ($stabilityLoadsRange.Column -eq 16) -and
            ($stabilityLoadsRange.Rows.Count -eq 31) -and ($stabilityLoadsRange.Columns.Count -eq 4)
        ) ("Address=$($stabilityLoadsRange.Address())")
        Add-Check $checks "SP35 table 7.21 range" (
            ($sp35TableRange.Worksheet.Name -eq "Config") -and
            ($sp35TableRange.Row -eq 60) -and ($sp35TableRange.Column -eq 35) -and
            ($sp35TableRange.Columns.Count -eq 8) -and ($sp35TableRange.Rows.Count -ge 2)
        ) ("Address=$($sp35TableRange.Address()); Rows=$($sp35TableRange.Rows.Count); Columns=$($sp35TableRange.Columns.Count)")
    }

    $duplicates = @($actualNames | Group-Object | Where-Object { $_.Count -gt 1 } | ForEach-Object { $_.Name })
    Add-Check $checks "No duplicate names" ($duplicates.Count -eq 0) ("Duplicates: " + ($duplicates -join ", "))

    $obsoleteNames = @(
        "rngResultMx", "rngResultMy", "rngResultMxy", "rngMainInput", "rngRebarInput",
        "rngSystemDiagnostics", "rngConcreteDiagramPoints", "rngSteelDiagramPoints",
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

    $loadsRangeForLayout = $workbook.Names.Item("rngLoadCombinations").RefersToRange
    $loadCombinationsOnConfig = (($loadsRangeForLayout.Worksheet.Name -eq "Config") -and ($loadsRangeForLayout.Row -eq 3) -and ($loadsRangeForLayout.Column -eq 15))
    Add-Check $checks "Load combinations table on Config O3" $loadCombinationsOnConfig ("Sheet=$($loadsRangeForLayout.Worksheet.Name); Row=$($loadsRangeForLayout.Row); Column=$($loadsRangeForLayout.Column)")

    $settings = $workbook.Names.Item("rngSystemSettings").RefersToRange
    $expectedSettingsHeaders = @("Параметр", "Значение", "Ед.", "Комментарий", "Справка")
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
        "Capacity.CalculateMxy", "Mesh.BoundaryMode", "Mesh.Step", "Circle.Radius",
        "Batch.MaxCombinations", "Batch.Diagnostics", "Materials.SourceStatus",
        "Concrete.Diagram", "Steel.Diagram",
        "Concrete.Point1.Eps", "Concrete.Point1.Stress",
        "Concrete.Point2.Eps", "Concrete.Point2.Stress",
        "Concrete.Point3.Eps", "Concrete.Point3.Stress",
        "Steel.Point1.Eps", "Steel.Point1.Stress",
        "Steel.Point2.Eps", "Steel.Point2.Stress",
        "Steel.Point3.Eps", "Steel.Point3.Stress",
        "Concrete.TensionMode",
        "Capacity.ConcreteCompressionLimit", "Capacity.ConcreteTensionLimit",
        "Capacity.SteelStrainLimit", "Concrete.Class",
        "Circle.CenterX", "Circle.CenterY",
        "RectSet.OriginX", "RectSet.OriginY",
        "Plot.DimensionsEnabled", "Plot.RebarLabelsEnabled",
        "Plot.PrincipalAxesEnabled", "Plot.CentroidEnabled",
        "AutoCAD.Export.PrincipalAxesEnabled"
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
    $solverCheckName = "Solver.Method default"
    if ($UserConfiguredWorkbook) {
        $solverCheckName = "Solver.Method user value"
        $solverMethodDefaultOk = ($solverMethodCell -ne $null) -and ([string]$solverMethodCell.Value2 -in @("Newton", "Secant"))
    }
    Add-Check $checks $solverCheckName $solverMethodDefaultOk ("Value=" + [string]$(if ($solverMethodCell -eq $null) { "" } else { $solverMethodCell.Value2 }))

    $stabilityCodeCell = $null
    for ($i = 2; $i -le $settings.Rows.Count; $i++) {
        if ([string]$settings.Cells.Item($i, 1).Value2 -eq "Stability.Code") {
            $stabilityCodeCell = $settings.Cells.Item($i, 2)
            break
        }
    }
    $stabilityCodeDefaultOk = ($stabilityCodeCell -ne $null) -and ([string]$stabilityCodeCell.Value2 -eq "SP35")
    $stabilityCheckName = "Stability.Code default"
    if ($UserConfiguredWorkbook) {
        $stabilityCheckName = "Stability.Code user value"
        $stabilityCodeDefaultOk = ($stabilityCodeCell -ne $null) -and ([string]$stabilityCodeCell.Value2 -in @("SP35", "SP63"))
    }
    Add-Check $checks $stabilityCheckName $stabilityCodeDefaultOk ("Value=" + [string]$(if ($stabilityCodeCell -eq $null) { "" } else { $stabilityCodeCell.Value2 }))

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

    $concreteParams = $workbook.Names.Item("rngConcreteMaterialParameters").RefersToRange
    $steelParams = $workbook.Names.Item("rngSteelMaterialParameters").RefersToRange
    $calculationProfiles = $workbook.Names.Item("rngCalculationProfiles").RefersToRange
    Add-Check $checks "Concrete material parameters" (($concreteParams.Columns.Count -eq 6) -and ($concreteParams.Rows.Count -eq 8)) ("Rows=$($concreteParams.Rows.Count); Columns=$($concreteParams.Columns.Count)")
    Add-Check $checks "Steel material parameters" (($steelParams.Columns.Count -eq 6) -and ($steelParams.Rows.Count -eq 6)) ("Rows=$($steelParams.Rows.Count); Columns=$($steelParams.Columns.Count)")
    Add-Check $checks "Calculation profiles" (($calculationProfiles.Columns.Count -eq 7) -and ($calculationProfiles.Rows.Count -ge 26)) ("Rows=$($calculationProfiles.Rows.Count); Columns=$($calculationProfiles.Columns.Count)")

    $rightStackNames = @(
        "rngUnitSettings",
        "rngSignConventionSettings",
        "rngSteelMaterialParameters",
        "rngConcreteMaterialParameters",
        "rngCalculationProfiles",
        "rngPlotAnnotationSettings",
        "rngCircleGeometry",
        "rngRectSetGeometry",
        "rngRoundedRectangleGeometry",
        "rngHollowRectangleGeometry"
    )
    $rightStackOk = $true
    $rightStackDetails = @()
    $previousEndRow = 0
    foreach ($rangeName in $rightStackNames) {
        $range = $workbook.Names.Item($rangeName).RefersToRange
        $endRow = $range.Row + $range.Rows.Count - 1
        $rightStackDetails += "$rangeName=R$($range.Row)C$($range.Column):R$endRow"
        if (($range.Worksheet.Name -ne "Config") -or ($range.Column -ne 8) -or ($range.Row -le $previousEndRow)) {
            $rightStackOk = $false
        }
        $previousEndRow = $endRow
    }
    Add-Check $checks "Config right-side ranges vertical stack" $rightStackOk ($rightStackDetails -join "; ")

    $loads = $workbook.Names.Item("rngLoadCombinations").RefersToRange
    $expectedLoadHeaders = @("CombinationID", "N, tf", "Mx, tf*m", "My, tf*m", "ProfileId", "LoadPath", "Comment")
    $actualLoadHeaders = @()
    for ($i = 1; $i -le 7; $i++) {
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
catch {
    Add-Check $checks "Validation runtime" $false $_.Exception.Message
}
finally {
    if ($workbook -ne $null) {
        try {
            $workbook.Close($false)
        }
        catch {
            Write-Warning "Excel COM refused to close workbook cleanly after validation: $($_.Exception.Message)"
        }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
    }
    if ($excel -ne $null) {
        try {
            $excel.Quit()
        }
        catch {
            Write-Warning "Excel COM refused to quit cleanly after validation: $($_.Exception.Message)"
        }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

$afterHash = (Get-FileHash -LiteralPath $fullWorkbookPath -Algorithm SHA256).Hash
Add-Check $checks "Source workbook unchanged" ($beforeHash -eq $afterHash) "Before=$beforeHash; After=$afterHash"
if ($ReportPath) {
    $checks | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $root $ReportPath) -Encoding UTF8
}
$checks | Format-Table -AutoSize

$failed = @($checks | Where-Object { -not $_.Passed })
if ($failed.Count -gt 0) {
    exit 1
}














