# скрипт обновляет отдельные части существующей книги без ручного импорта модулей через редактор VBA.

param(
    [string]$WorkbookPath = "workbook/output/RC_Section_NDM.xlsm"
)

$ErrorActionPreference = "Stop"

# Добавляет структурный элемент книги или отчета, сохраняя единый формат сборочных скриптов.
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

    $source = [System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::UTF8)
    $lines = $source -split "`r?`n"
    $componentName = $null
    foreach ($line in $lines) {
        if ($line -match '^\uFEFF?Attribute\s+VB_Name\s*=\s*"([^"]+)"') {
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
        if ($line -match '^\uFEFF?VERSION\s+') { continue }
        if ($line -match '^\uFEFF?BEGIN\s*$') {
            $insideVbaHeaderBlock = $true
            continue
        }
        if ($insideVbaHeaderBlock) {
            if ($line -match '^\uFEFF?END\s*$') {
                $insideVbaHeaderBlock = $false
            }
            continue
        }
        if ($line -match '^\uFEFF?Attribute\s+VB_') { continue }
        $bodyLines.Add($line)
    }

    $body = ($bodyLines -join "`r`n").Trim()
    if ($body.Length -gt 0) {
        $component.CodeModule.AddFromString($body)
    }
}

# Импортирует исходные VBA-модули в книгу, сохраняя воспроизводимость сборки.
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
        (Join-Path $RootPath "src/Stability"),
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

# Удаляет только служебный объект, который может мешать повторяемой сборке или проверке.
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

# Устанавливает значение, оформление или именованный диапазон в книге через Excel COM.
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
    $settings.Cells.Item($row, 7).Value2 = "Нет"
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

    Set-SystemSetting $workbook "Capacity.InitialLambda" "1" "1" "" "Начальный множитель для поиска верхней границы" "Проектная настройка этапа 5"
    Set-SystemSetting $workbook "Capacity.MaxLambda" "64" "64" "" "Предельное значение lambda при расширении скобки" "Проектная настройка этапа 5"
    Set-SystemSetting $workbook "Capacity.ToleranceLambda" "0.01" "0.01" "" "Допуск одномерного поиска предельного множителя" "Проектная настройка этапа 5"
    Set-SystemSetting $workbook "Capacity.MaxRetries" "4" "4" "шт" "Число повторов после численной несходимости пробы" "Проектная настройка этапа 5"
    Set-SystemSetting $workbook "Capacity.BaseLoadSteps" "8" "8" "шт" "Базовое число внутренних ступеней нагрузки в CSectionSolver" "Проектная настройка этапа 5"
    Set-SystemSetting $workbook "Capacity.SolverMaxIterations" "60" "60" "шт" "Максимум итераций Newton на ступень при поиске несущей способности" "Проектная настройка этапа 5"
    Set-SystemSetting $workbook "SLS.Crack.Allowable" "0.3" "0.3" "мм" "User-defined allowable crack width a_crc,ult" "Stage 7 project setting"
    Set-SystemSetting $workbook "SLS.Crack.TensionZoneMode" "Effective" "Effective" "Effective/FullTension" "Concrete tension zone for Abt in crack width calculation" "Stage 7 project setting"
    Set-SystemSetting $workbook "SLS.Crack.CoverDistanceMode" "GlobalExtreme" "GlobalExtreme" "NearestContour/GlobalExtreme" "Cover distance a_s mode" "Stage 7 project setting"
    Set-SystemSetting $workbook "SLS.Crack.Phi1" "1.4" "1.4" "" "Crack coefficient phi1" "Stage 7 project setting"
    Set-SystemSetting $workbook "SLS.Crack.Phi2" "0.5" "0.5" "" "Crack coefficient phi2" "Stage 7 project setting"
    Set-SystemSetting $workbook "SLS.Crack.Phi3Mode" "Auto" "Auto" "Auto/User" "Phi3 mode" "Stage 7 project setting"
    Set-SystemSetting $workbook "SLS.Crack.Phi3" "1" "1" "" "User phi3 when Phi3Mode=User" "Stage 7 project setting"
    Set-SystemSetting $workbook "SLS.Crack.PsiMode" "User" "User" "User/Auto" "Psi_s mode: user value or auto after failed first crack-width check" "Stage 7 project setting"
    Set-SystemSetting $workbook "SLS.Crack.PsiS" "1" "1" "" "User psi_s when PsiMode=User" "Stage 7 project setting"

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


