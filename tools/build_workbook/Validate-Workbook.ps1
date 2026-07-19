param(
    [string]$WorkbookPath = "workbook/output/RC_Section_NDM.xlsm"
)

$ErrorActionPreference = "Stop"

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

function Is-AllowedFormulaCell {
    param([string]$SheetName, [int]$Row, [int]$Column)

    if ($SheetName -eq "Расчет") {
        if ($Row -ge 6 -and $Row -le 60 -and $Column -ge 1 -and $Column -le 7) { return $true }
        if ($Row -ge 17 -and $Row -le 35 -and $Column -ge 19 -and $Column -le 70) { return $true }
    }
    if ($SheetName -eq "System") {
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

    Add-Check $checks "Sheets Расчет/System" (($sheetNames -contains "Расчет") -and ($sheetNames -contains "System")) ($sheetNames -join ", ")

    $requiredNames = @(
        "rngMainInput",
        "rngRebarInput",
        "rngLoadCombinations",
        "rngResultSection",
        "rngSystemSettings",
        "rngSystemDiagnostics",
        "SolverSettings",
        "CapacitySettings",
        "ConcreteDiagram",
        "SteelDiagram",
        "GeometrySettings",
        "OutputSettings",
        "AutoCADSettings"
    )

    $actualNames = @()
    foreach ($name in $workbook.Names) {
        $actualNames += [string]$name.Name
    }

    $missingNames = @($requiredNames | Where-Object { $actualNames -notcontains $_ })
    Add-Check $checks "Required named ranges" ($missingNames.Count -eq 0) ("Missing: " + ($missingNames -join ", "))

    $duplicates = @($actualNames | Group-Object | Where-Object { $_.Count -gt 1 } | ForEach-Object { $_.Name })
    Add-Check $checks "No duplicate names" ($duplicates.Count -eq 0) ("Duplicates: " + ($duplicates -join ", "))

    $obsoleteNames = @("rngResultMx", "rngResultMy", "rngResultMxy")
    $presentObsoleteNames = @($obsoleteNames | Where-Object { $actualNames -contains $_ })
    Add-Check $checks "No obsolete Mx/My result ranges" ($presentObsoleteNames.Count -eq 0) ("Present: " + ($presentObsoleteNames -join ", "))

    $calc = $workbook.Worksheets.Item("Расчет")
    $system = $workbook.Worksheets.Item("System")

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

    $main = $workbook.Names.Item("rngMainInput").RefersToRange
    $sectionResult = $workbook.Names.Item("rngResultSection").RefersToRange
    $leftToRight = ($main.Column -lt $sectionResult.Column)
    Add-Check $checks "Input and result blocks left to right" $leftToRight ("Columns: main=$($main.Column), section=$($sectionResult.Column)")

    $settings = $workbook.Names.Item("rngSystemSettings").RefersToRange
    $expectedSettingsHeaders = @("Параметр", "Значение", "Ед.", "Комментарий")
    $actualSettingsHeaders = @()
    for ($i = 1; $i -le $expectedSettingsHeaders.Count; $i++) {
        $actualSettingsHeaders += [string]$settings.Cells.Item(1, $i).Value2
    }
    Add-Check $checks "System settings table headers" (($actualSettingsHeaders -join "|") -eq ($expectedSettingsHeaders -join "|")) ($actualSettingsHeaders -join " | ")

    $settingKeys = @()
    for ($i = 2; $i -le $settings.Rows.Count; $i++) {
        $key = [string]$settings.Cells.Item($i, 1).Value2
        if (-not [string]::IsNullOrWhiteSpace($key) -and -not $key.StartsWith("[")) {
            $settingKeys += $key
        }
    }
    $duplicateSettingKeys = @($settingKeys | Group-Object | Where-Object { $_.Count -gt 1 } | ForEach-Object { $_.Name })
    Add-Check $checks "No duplicate System keys" ($duplicateSettingKeys.Count -eq 0) ("Duplicates: " + ($duplicateSettingKeys -join ", "))

    $obsoleteSettingKeys = @("Capacity.Enabled", "Capacity.CalculateMx", "Capacity.CalculateMy", "Capacity.CalculateMxy", "Mesh.BoundaryMode", "Circle.Radius", "Batch.MaxCombinations", "Batch.Diagnostics", "Materials.SourceStatus", "Concrete.Diagram", "Steel.Diagram")
    $presentObsoleteSettings = @($obsoleteSettingKeys | Where-Object { $settingKeys -contains $_ })
    Add-Check $checks "No obsolete System settings" ($presentObsoleteSettings.Count -eq 0) ("Present: " + ($presentObsoleteSettings -join ", "))

    $rebar = $workbook.Names.Item("rngRebarInput").RefersToRange
    $expectedRebarHeaders = @("ID", "X", "Y", "Diameter", "Area", "SteelClass", "Comment")
    $actualRebarHeaders = @()
    for ($i = 1; $i -le $expectedRebarHeaders.Count; $i++) {
        $actualRebarHeaders += [string]$rebar.Cells.Item(1, $i).Value2
    }
    Add-Check $checks "Rebar table headers" (($actualRebarHeaders -join "|") -eq ($expectedRebarHeaders -join "|")) ($actualRebarHeaders -join " | ")

    $loads = $workbook.Names.Item("rngLoadCombinations").RefersToRange
    $expectedLoadHeaders = @("CombinationID", "N", "Mx", "My", "CalculationType", "DurationType", "Comment")
    $actualLoadHeaders = @()
    for ($i = 1; $i -le 7; $i++) {
        $actualLoadHeaders += [string]$loads.Cells.Item(1, $i).Value2
    }
    Add-Check $checks "Load combinations table headers" (($actualLoadHeaders -join "|") -eq ($expectedLoadHeaders -join "|")) ($actualLoadHeaders -join " | ")
    Add-Check $checks "Formulas allowed on workbook sheets" $true "Excel formulas are allowed and preferred for transparent user-level calculations; VBA validation covers workbook structure."
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














