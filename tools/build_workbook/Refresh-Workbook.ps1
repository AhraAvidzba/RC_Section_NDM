param(
    [string]$WorkbookPath = "workbook/output/RC_Section_NDM.xlsm"
)

$ErrorActionPreference = "Stop"

function Add-VbaSourceFile {
    param(
        [object]$Workbook,
        [string]$Path
    )

    $extension = [System.IO.Path]::GetExtension($Path).ToLowerInvariant()
    if ($extension -eq ".frm") {
        $Workbook.VBProject.VBComponents.Import($Path) | Out-Null
        return
    }

    $source = [System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::Default)
    $lines = $source -split "`r?`n"
    $componentName = $null
    foreach ($line in $lines) {
        if ($line -match '^Attribute\s+VB_Name\s*=\s*"([^"]+)"') {
            $componentName = $Matches[1]
            break
        }
    }
    if ([string]::IsNullOrWhiteSpace($componentName)) {
        throw "VBA source has no Attribute VB_Name: $Path"
    }

    $componentType = 1
    if ($extension -eq ".cls") {
        $componentType = 2
    }

    $component = $Workbook.VBProject.VBComponents.Add($componentType)
    $component.Name = $componentName

    $bodyLines = New-Object System.Collections.Generic.List[string]
    $insideVbaHeaderBlock = $false
    foreach ($line in $lines) {
        if ($line -match '^VERSION\s+') { continue }
        if ($line -match '^BEGIN\s*$') {
            $insideVbaHeaderBlock = $true
            continue
        }
        if ($insideVbaHeaderBlock) {
            if ($line -match '^END\s*$') {
                $insideVbaHeaderBlock = $false
            }
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

function Import-VbaSourceTree {
    param(
        [object]$Workbook,
        [string]$RootPath
    )

    $sourceDirs = @(
        (Join-Path $RootPath "src/Common"),
        (Join-Path $RootPath "src/Interfaces"),
        (Join-Path $RootPath "src/Geometry"),
        (Join-Path $RootPath "src/Section"),
        (Join-Path $RootPath "src/Materials"),
        (Join-Path $RootPath "src/Solver"),
        (Join-Path $RootPath "src/Crack"),
        (Join-Path $RootPath "src/Batch"),
        (Join-Path $RootPath "src/Excel"),
        (Join-Path $RootPath "tests")
    )

    foreach ($dir in $sourceDirs) {
        if (Test-Path -LiteralPath $dir) {
            Get-ChildItem -LiteralPath $dir -Recurse -File |
                Where-Object { $_.Extension -in @(".bas", ".cls", ".frm") } |
                Sort-Object FullName |
                ForEach-Object {
                    Add-VbaSourceFile $Workbook $_.FullName
                }
        }
    }
}

function Remove-ImportedVbaComponents {
    param([object]$Workbook)

    $components = @()
    foreach ($component in $Workbook.VBProject.VBComponents) {
        if ($component.Type -ne 100) {
            $components += $component
        }
    }

    foreach ($component in $components) {
        $Workbook.VBProject.VBComponents.Remove($component)
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
            throw "rngSystemSettings has no free row for setting $Key. Run full Build-Workbook.ps1 after closing Excel."
        }
    }

    $settings.Cells.Item($row, 1).Value2 = $Key
    $settings.Cells.Item($row, 2).Value2 = $Value
    $settings.Cells.Item($row, 3).Value2 = $DefaultValue
    $settings.Cells.Item($row, 4).Value2 = $Unit
    $settings.Cells.Item($row, 5).Value2 = $Purpose
    $settings.Cells.Item($row, 6).Value2 = $Source
    $settings.Cells.Item($row, 7).Value2 = "РќРµС‚"
}

$root = Resolve-Path -LiteralPath (Join-Path $PSScriptRoot "../..")
$fullWorkbookPath = Join-Path $root $WorkbookPath

if (-not (Test-Path -LiteralPath $fullWorkbookPath)) {
    throw "Workbook not found: $fullWorkbookPath"
}

$excel = $null
$workbook = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false
    $excel.AutomationSecurity = 1

    $workbook = $excel.Workbooks.Open($fullWorkbookPath)
    Remove-ImportedVbaComponents $workbook
    Import-VbaSourceTree $workbook $root

    Set-SystemSetting $workbook "Capacity.InitialLambda" "1" "1" "" "РќР°С‡Р°Р»СЊРЅС‹Р№ РјРЅРѕР¶РёС‚РµР»СЊ РґР»СЏ РїРѕРёСЃРєР° РІРµСЂС…РЅРµР№ РіСЂР°РЅРёС†С‹" "РџСЂРѕРµРєС‚РЅР°СЏ РЅР°СЃС‚СЂРѕР№РєР° СЌС‚Р°РїР° 5"
    Set-SystemSetting $workbook "Capacity.MaxLambda" "64" "64" "" "РџСЂРµРґРµР»СЊРЅРѕРµ Р·РЅР°С‡РµРЅРёРµ lambda РїСЂРё СЂР°СЃС€РёСЂРµРЅРёРё СЃРєРѕР±РєРё" "РџСЂРѕРµРєС‚РЅР°СЏ РЅР°СЃС‚СЂРѕР№РєР° СЌС‚Р°РїР° 5"
    Set-SystemSetting $workbook "Capacity.ToleranceLambda" "0.01" "0.01" "" "Р”РѕРїСѓСЃРє РѕРґРЅРѕРјРµСЂРЅРѕРіРѕ РїРѕРёСЃРєР° РїСЂРµРґРµР»СЊРЅРѕРіРѕ РјРЅРѕР¶РёС‚РµР»СЏ" "РџСЂРѕРµРєС‚РЅР°СЏ РЅР°СЃС‚СЂРѕР№РєР° СЌС‚Р°РїР° 5"
    Set-SystemSetting $workbook "Capacity.MaxRetries" "4" "4" "С€С‚" "Р§РёСЃР»Рѕ РїРѕРІС‚РѕСЂРѕРІ РїРѕСЃР»Рµ С‡РёСЃР»РµРЅРЅРѕР№ РЅРµСЃС…РѕРґРёРјРѕСЃС‚Рё РїСЂРѕР±С‹" "РџСЂРѕРµРєС‚РЅР°СЏ РЅР°СЃС‚СЂРѕР№РєР° СЌС‚Р°РїР° 5"
    Set-SystemSetting $workbook "Capacity.BaseLoadSteps" "8" "8" "С€С‚" "Р‘Р°Р·РѕРІРѕРµ С‡РёСЃР»Рѕ РІРЅСѓС‚СЂРµРЅРЅРёС… СЃС‚СѓРїРµРЅРµР№ РЅР°РіСЂСѓР·РєРё РІ CSectionSolver" "РџСЂРѕРµРєС‚РЅР°СЏ РЅР°СЃС‚СЂРѕР№РєР° СЌС‚Р°РїР° 5"
    Set-SystemSetting $workbook "Capacity.SolverMaxIterations" "60" "60" "С€С‚" "РњР°РєСЃРёРјСѓРј РёС‚РµСЂР°С†РёР№ Newton РЅР° СЃС‚СѓРїРµРЅСЊ РїСЂРё РїРѕРёСЃРєРµ РЅРµСЃСѓС‰РµР№ СЃРїРѕСЃРѕР±РЅРѕСЃС‚Рё" "РџСЂРѕРµРєС‚РЅР°СЏ РЅР°СЃС‚СЂРѕР№РєР° СЌС‚Р°РїР° 5"
    Set-SystemSetting $workbook "Capacity.ConcreteCompressionLimit" "-0.0035" "-0.0035" "" "РџСЂРµРґРµР» РґРµС„РѕСЂРјР°С†РёРё Р±РµС‚РѕРЅР° РґР»СЏ С„РёРєСЃР°С†РёРё ConcreteStrainLimit" "PROVISIONAL_FOR_SOLVER_TESTING"
    Set-SystemSetting $workbook "Capacity.SteelStrainLimit" "0.025" "0.025" "" "РџСЂРµРґРµР» РґРµС„РѕСЂРјР°С†РёРё РѕР±С‹С‡РЅРѕР№ РЅРµРЅР°РїСЂСЏРіР°РµРјРѕР№ Р°СЂРјР°С‚СѓСЂС‹ РґР»СЏ SteelStrainLimit" "PROVISIONAL_FOR_SOLVER_TESTING"
    Set-SystemSetting $workbook "Concrete.TensionMode" "Ignore" "Ignore" "Ignore/UseDiagram" "Concrete tension behavior for cracked and material-diagram calculations" "Stage 7 project setting; normative applicability pending"
    Set-SystemSetting $workbook "CrackWidth.Enabled" "Да" "Да" "Да/Нет" "Calculate already formed normal crack width" "Stage 7 project setting"
    Set-SystemSetting $workbook "CrackWidth.Allowable" "0.3" "0.3" "мм" "Provisional allowable crack width" "PROVISIONAL_FOR_SOLVER_TESTING"
    Set-SystemSetting $workbook "CrackWidth.CrackSpacing" "200" "200" "мм" "Provisional normal crack spacing" "PROVISIONAL_FOR_SOLVER_TESTING"
    Set-SystemSetting $workbook "CrackWidth.StrainFactor" "1" "1" "" "Provisional strain factor for crack width" "PROVISIONAL_FOR_SOLVER_TESTING"
    Set-SystemSetting $workbook "CrackWidth.DurationFactor" "1" "1" "" "Provisional load-duration factor for crack width" "PROVISIONAL_FOR_SOLVER_TESTING"

    $workbook.Save()
    Write-Output "Workbook refreshed: $fullWorkbookPath"
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


