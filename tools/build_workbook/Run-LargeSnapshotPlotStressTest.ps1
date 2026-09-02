# служебный PowerShell-скрипт запускает тяжелую проверку крупного Results snapshot и Excel-схемы.

param(
    [string]$WorkbookPath = "workbook/output/RC_Section_NDM.xlsm"
)

$ErrorActionPreference = "Stop"

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
    $excel.AutomationSecurity = 1
    $excel.EnableEvents = $false

    $workbook = $excel.Workbooks.Open($fullWorkbookPath)
    Start-Sleep -Seconds 1

    $result = $excel.Run("'RC_Section_NDM.xlsm'!modTestWorkbookInterface.RunLargeSnapshotPlotStressTest")
    Write-Output $result
}
finally {
    if ($workbook -ne $null) {
        $workbook.Close($false)
        [Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
    }
    if ($excel -ne $null) {
        $excel.Quit()
        [Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}
