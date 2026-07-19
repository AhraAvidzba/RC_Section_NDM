param(
    [string]$WorkbookPath = "workbook/output/RC_Section_NDM.xlsm"
)

$ErrorActionPreference = "Stop"

function Get-VbaComponentName {
    param([string]$Path)
    $source = [System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::Default)
    foreach ($line in ($source -split "`r?`n")) {
        if ($line -match '^Attribute\s+VB_Name\s*=\s*"([^"]+)"') {
            return $Matches[1]
        }
    }
    throw "VBA source has no Attribute VB_Name: $Path"
}

function Import-VbaSourceFile {
    param(
        [object]$Workbook,
        [string]$Path
    )

    $componentName = Get-VbaComponentName $Path
    $existing = $null
    foreach ($component in $Workbook.VBProject.VBComponents) {
        if ($component.Name -eq $componentName) {
            $existing = $component
            break
        }
    }
    if ($existing -ne $null) {
        $Workbook.VBProject.VBComponents.Remove($existing)
    }

    $extension = [System.IO.Path]::GetExtension($Path).ToLowerInvariant()
    $componentType = 1
    if ($extension -eq ".cls") { $componentType = 2 }
    $component = $Workbook.VBProject.VBComponents.Add($componentType)
    $component.Name = $componentName

    $source = [System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::Default)
    $lines = $source -split "`r?`n"
    $bodyLines = New-Object System.Collections.Generic.List[string]
    $insideHeader = $false
    foreach ($line in $lines) {
        if ($line -match '^VERSION\s+') { continue }
        if ($line -match '^BEGIN\s*$') { $insideHeader = $true; continue }
        if ($insideHeader) {
            if ($line -match '^END\s*$') { $insideHeader = $false }
            continue
        }
        if ($line -match '^Attribute\s+VB_') { continue }
        $bodyLines.Add($line)
    }
    $body = ($bodyLines -join "").Trim()
    if ($body.Length -gt 0) {
        $component.CodeModule.AddFromString($body)
    }
}

function Set-SystemSetting {
    param(
        [object]$Workbook,
        [string]$Key,
        [string]$Value,
        [string]$DefaultValue,
        [string]$Unit,
        [string]$Purpose,
        [string]$Source
    )

    $settings = $Workbook.Names.Item("rngSystemSettings").RefersToRange
    $row = $null
    for ($i = 2; $i -le $settings.Rows.Count; $i++) {
        if ([string]::Equals([string]$settings.Cells.Item($i, 1).Value2, $Key, [System.StringComparison]::OrdinalIgnoreCase)) {
            $row = $i
            break
        }
    }
    if ($row -eq $null) {
        $row = $settings.Rows.Count
        while ($row -gt 2 -and [string]::IsNullOrWhiteSpace([string]$settings.Cells.Item($row, 1).Value2)) {
            $row--
        }
        $row++
        if ($row -gt $settings.Rows.Count) {
            $sheet = $settings.Worksheet
            $insertAt = $settings.Row + $settings.Rows.Count
            $sheet.Rows.Item($insertAt).Insert() | Out-Null
            $newRange = $sheet.Range($sheet.Cells.Item($settings.Row, $settings.Column), $sheet.Cells.Item($settings.Row + $settings.Rows.Count, $settings.Column + $settings.Columns.Count - 1))
            $Workbook.Names.Item("rngSystemSettings").RefersTo = "=" + $newRange.Address($true, $true, 1, $true)
            $settings = $Workbook.Names.Item("rngSystemSettings").RefersToRange
            $row = $settings.Rows.Count
        }
    }

    $settings.Cells.Item($row, 1).Value2 = $Key
    $settings.Cells.Item($row, 2).Value2 = $Value
    $settings.Cells.Item($row, 3).Value2 = $DefaultValue
    $settings.Cells.Item($row, 4).Value2 = $Unit
    $settings.Cells.Item($row, 5).Value2 = $Purpose
    $settings.Cells.Item($row, 6).Value2 = $Source
    $settings.Cells.Item($row, 7).Value2 = "No"
}

$root = Resolve-Path -LiteralPath (Join-Path $PSScriptRoot "../..")
$fullWorkbookPath = Join-Path $root $WorkbookPath
if (-not (Test-Path -LiteralPath $fullWorkbookPath)) { throw "Workbook not found: $fullWorkbookPath" }

$excel = $null
$workbook = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false
    $excel.AutomationSecurity = 1

    $workbook = $excel.Workbooks.Open($fullWorkbookPath)

    $files = @(
        "src/Materials/CConcreteDiagramMaterial.cls",        "src/Crack/CCrackWidthCalculator.cls",
        "src/Excel/CCapacityResultWriter.cls",
        "tests/modTestCrackWidth.bas"
    )
    foreach ($relative in $files) {
        Import-VbaSourceFile $workbook (Join-Path $root $relative)
    }

    Set-SystemSetting $workbook "Concrete.TensionMode" "Ignore" "Ignore" "Ignore/UseDiagram" "Concrete tension behavior for cracked and material-diagram calculations" "Stage 7 project setting; normative applicability pending"
    Set-SystemSetting $workbook "CrackWidth.Enabled" "Yes" "Yes" "Yes/No" "Calculate already formed normal crack width" "Stage 7 project setting"
    Set-SystemSetting $workbook "CrackWidth.Allowable" "0.3" "0.3" "mm" "Provisional allowable crack width" "PROVISIONAL_FOR_SOLVER_TESTING"
    Set-SystemSetting $workbook "CrackWidth.CrackSpacing" "200" "200" "mm" "Provisional normal crack spacing" "PROVISIONAL_FOR_SOLVER_TESTING"
    Set-SystemSetting $workbook "CrackWidth.StrainFactor" "1" "1" "" "Provisional strain factor for crack width" "PROVISIONAL_FOR_SOLVER_TESTING"
    Set-SystemSetting $workbook "CrackWidth.DurationFactor" "1" "1" "" "Provisional load-duration factor for crack width" "PROVISIONAL_FOR_SOLVER_TESTING"

    $workbook.Save()
    Write-Output "Workbook stage 7 refreshed: $fullWorkbookPath"
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

