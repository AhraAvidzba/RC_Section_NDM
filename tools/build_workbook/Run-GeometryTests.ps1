# служебный PowerShell-скрипт поддерживает сборку, проверку или обновление Excel-книги проекта.

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

    $compileControl = $excel.VBE.CommandBars.FindControl(1, 578)
    if ($compileControl -ne $null) {
        $compileControl.Execute()
    }

    $result = $excel.Run("'RC_Section_NDM.xlsm'!modTestGeometry.RunGeometryTests")
    Write-Output $result
}
finally {
    if ($workbook -ne $null) {
        try {
            $workbook.Close($false)
        }
        catch {
            Write-Warning "Excel workbook refused to close cleanly after geometry tests: $($_.Exception.Message)"
        }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
    }
    if ($excel -ne $null) {
        try {
            $excel.Quit()
        }
        catch {
            Write-Warning "Excel COM refused to quit cleanly after geometry tests: $($_.Exception.Message)"
        }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}








