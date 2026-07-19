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

    $workbook = $excel.Workbooks.Open($fullWorkbookPath)
    $result = $excel.Run("'RC_Section_NDM.xlsm'!modTestLinearCore.RunLinearCoreTests")
    Write-Output $result
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





