# служебный PowerShell-скрипт поддерживает сборку, проверку или обновление Excel-книги проекта.

param(
    [string]$WorkbookPath = "workbook/output/RC_Section_NDM.xlsm",
    [string]$ReportPath = "docs/regression/Stage01_RegressionBaseline_Raw.txt"
)

$ErrorActionPreference = "Stop"

$root = Resolve-Path -LiteralPath (Join-Path $PSScriptRoot "../..")
$fullWorkbookPath = Join-Path $root $WorkbookPath
$fullReportPath = Join-Path $root $ReportPath

if (-not (Test-Path -LiteralPath $fullWorkbookPath)) {
    throw "Workbook not found: $fullWorkbookPath"
}

$reportDirectory = Split-Path -Parent $fullReportPath
if (-not (Test-Path -LiteralPath $reportDirectory)) {
    New-Item -ItemType Directory -Path $reportDirectory | Out-Null
}

$excel = $null
$workbook = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1

    $workbook = $excel.Workbooks.Open($fullWorkbookPath)
    $result = $excel.Run("'RC_Section_NDM.xlsm'!modTestRegressionBaseline.RunRegressionBaselineTests")
    $result | Set-Content -LiteralPath $fullReportPath -Encoding UTF8
    Write-Output $result
    Write-Output "Regression baseline report: $fullReportPath"
}
finally {
    if ($workbook -ne $null) {
        try {
            $workbook.Close($false)
        }
        catch {
            Write-Warning "Excel workbook refused to close cleanly after regression baseline tests: $($_.Exception.Message)"
        }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
    }
    if ($excel -ne $null) {
        try {
            $excel.Quit()
        }
        catch {
            Write-Warning "Excel COM refused to quit cleanly after regression baseline tests: $($_.Exception.Message)"
        }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
