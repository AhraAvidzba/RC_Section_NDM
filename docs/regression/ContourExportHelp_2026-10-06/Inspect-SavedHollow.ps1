# Читает сохраненные настройки и R22 без запуска расчетов и без правок книги.
$ErrorActionPreference = 'Stop'
$root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../../..')).Path
$excel = $null; $book = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.EnableEvents = $false; $excel.AutomationSecurity = 3
    $book = $excel.Workbooks.Open((Join-Path $root 'workbook/output/RC_Section_NDM.xlsm'), 0, $true)
    foreach ($name in 'rngHollowRectangleGeometry', 'rngLoadCombinations', 'rngUnitSettings') {
        $data = $book.Names.Item($name).RefersToRange.Value2
        Write-Output $name
        for ($r = 1; $r -le $data.GetLength(0); $r++) {
            $values = @(); for ($c = 1; $c -le $data.GetLength(1); $c++) { $values += [string]$data[$r,$c] }
            if (($values -join '|') -match '\S') { Write-Output ($values -join '|') }
        }
    }
    $data = $book.Names.Item('rngNDMSectionGeometry').RefersToRange.CurrentRegion.Value2
    for ($r = 1; $r -le $data.GetLength(0); $r++) {
        if ($r -eq 1 -or [string]$data[$r,2] -eq 'R22' -or [string]$data[$r,3] -eq 'Contour') {
            $values = @(); for ($c = 1; $c -le $data.GetLength(1); $c++) { $values += [string]$data[$r,$c] }
            Write-Output ($values -join '|')
        }
    }
} finally {
    if ($null -ne $book) { $book.Close($false) }
    if ($null -ne $excel) { $excel.Quit() }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
