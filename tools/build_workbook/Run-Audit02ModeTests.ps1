# Выполняет все штатные VBA suites на независимой копии заданной книги.
# Меняет только глобальную настройку копии: исходная baseline/output книга
# остается неизменной. Explicit-On setup внутри исторических тестов сохранен.
param(
    [Parameter(Mandatory=$true)][string]$SourceWorkbook,
    [string]$SettingKey = "General.DiagramExtension",
    [ValidateSet("Yes", "No")][string]$Mode = "No",
    [Parameter(Mandatory=$true)][string]$ReportPath,
    [string[]]$Macro = @()
)
$ErrorActionPreference = "Stop"
. (Join-Path $PSScriptRoot "SettingsCatalog.ps1")
$root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot "../..")).Path
$sourcePath = (Resolve-Path -LiteralPath $SourceWorkbook).Path
$sourceHash = (Get-FileHash -LiteralPath $sourcePath -Algorithm SHA256).Hash
$fixtureRoot = Join-Path ([IO.Path]::GetTempPath()) ("RC_NDM_ModeSuite_" + [Guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $fixtureRoot | Out-Null
$fixturePath = Join-Path $fixtureRoot "RC_Section_NDM.xlsm"
Copy-Item -LiteralPath $sourcePath -Destination $fixturePath
$printAreas = @(Get-WorkbookPrintAreas $fixturePath)
$lines = New-Object System.Collections.Generic.List[string]
$excel = $null
$workbook = $null
$failed = $false

# Находит единственную строку настройки: неизвестный формат не заменяет
# настройку другим default и не позволяет получить ложный Off-прогон.
function Get-ModeSettingCell {
    param([object]$Book, [string]$Key)
    $range = $Book.Names.Item("rngSystemSettings").RefersToRange
    $data = $range.Value2
    $row = 0
    for ($r = 2; $r -le $range.Rows.Count; $r++) {
        if ([string]$data[$r, 1] -eq $Key) {
            if ($row) { throw "В тестовой книге несколько строк $Key." }
            $row = $r
        }
    }
    if (-not $row) { throw "В тестовой книге отсутствует $Key." }
    return $range.Cells.Item($row, 2)
}

try {
    $lines.Add("SOURCE: $sourcePath")
    $lines.Add("SOURCE_SHA256: $sourceHash")
    $lines.Add("FIXTURE: $fixturePath")
    $lines.Add("GLOBAL_MODE: $SettingKey=$Mode; explicit-On tests retain their setup")
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1
    $workbook = $excel.Workbooks.Open($fixturePath)
    (Get-ModeSettingCell $workbook $SettingKey).Value2 = $Mode
    $workbook.Save()
    $workbook.Close($false)
    $workbook = $null
    Restore-WorkbookPrintAreas $fixturePath $printAreas
    $macros = @(
        "modTestGeometry.RunGeometryTests",
        "modTestMaterialDiagrams.RunMaterialDiagramTests",
        "modTestSectionSolver.RunSectionSolverTests",
        "modTestCapacitySolver.RunCapacitySolverTests",
        "modTestCrackWidth.RunCrackWidthTests",
        "modTestBatchCalculation.RunBatchCalculationTests",
        "modTestWorkbookInterface.RunWorkbookInterfaceTests",
        "modTestRegressionBaseline.RunRegressionBaselineTests"
    )
    if ($Macro.Count -gt 0) { $macros = $Macro }
    foreach ($macro in $macros) {
        $workbook = $excel.Workbooks.Open($fixturePath, $null, $true)
        $lines.Add("===== $macro =====")
        if ([string](Get-ModeSettingCell $workbook $SettingKey).Value2 -ne $Mode) {
            throw "Перед suite $macro не сохранено требуемое значение $Mode."
        }
        $result = [string]$excel.Run("'RC_Section_NDM.xlsm'!$macro")
        foreach ($line in ($result -split "`r?`n")) {
            $lines.Add($line)
            if ($line -match "failed=([1-9][0-9]*)|^FAIL:|RUNTIME ERROR") { $failed = $true }
        }
        $modeAfter = [string](Get-ModeSettingCell $workbook $SettingKey).Value2
        $lines.Add("MODE_AFTER_SUITE: $modeAfter")
        if ($modeAfter -ne $Mode) {
            $lines.Add("SUITE_MODE_OVERRIDE: explicit test setup left $modeAfter; changes discarded on close")
        }
        $workbook.Close($false)
        $workbook = $null
    }
    if ((Get-FileHash -LiteralPath $sourcePath -Algorithm SHA256).Hash -ne $sourceHash) {
        throw "Исходная книга изменилась во время проверки независимой копии."
    }
    $lines.Add("SOURCE_UNCHANGED: True")
} catch {
    $failed = $true
    $lines.Add("SCRIPT ERROR: $($_.Exception.Message)")
} finally {
    if ($workbook) { $workbook.Close($false); [Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null }
    if ($excel) { $excel.Quit(); [Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    $fullReport = Join-Path $root $ReportPath
    New-Item -ItemType Directory -Force -Path (Split-Path -Parent $fullReport) | Out-Null
    $lines | Set-Content -LiteralPath $fullReport -Encoding UTF8
    $lines | Where-Object { $_ -match "^SOURCE|^FIXTURE|^GLOBAL_MODE|^TOTAL|^FAIL:|^SCRIPT ERROR|^RUNTIME ERROR|^MODE_AFTER" }
}
if ($failed) { exit 1 }
