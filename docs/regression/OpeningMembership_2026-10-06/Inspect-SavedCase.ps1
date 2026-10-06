# Читает сохраненный пользовательский пример без запуска расчета и без записи книги.
$ErrorActionPreference = 'Stop'
$root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../../..')).Path
$path = Join-Path $root 'workbook/output/RC_Section_NDM.xlsm'
$before = (Get-FileHash -LiteralPath $path).Hash
$excel = $null; $book = $null; $records = @{}
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false; $excel.AutomationSecurity = 3
    $book = $excel.Workbooks.Open($path, 0, $true)
    foreach ($name in @('rngSystemSettings','rngHollowRectangleGeometry','rngLoadCombinations','rngUnitSettings',
        'rngNDMSectionGeometry','rngNDMElementResults','rngNDMSectionAnnotations','rngCrackSummaryAnchor')) {
        $range = $book.Names.Item($name).RefersToRange
        if ($name -match '^rngNDM') { $range = $range.CurrentRegion }
        if ($name -eq 'rngCrackSummaryAnchor') { $range = $range.Offset(-4,0).Resize(10,73) }
        $data = $range.Value2; $rows = New-Object 'System.Collections.Generic.List[object]'
        for ($r = 1; $r -le $data.GetLength(0); $r++) {
            $row = @(); for ($c = 1; $c -le $data.GetLength(1); $c++) { $row += $data[$r,$c] }
            $rows.Add($row)
        }
        $records[$name] = @{Address=$range.Address(); Rows=$rows.ToArray()}
    }
} finally {
    if ($null -ne $book) { $book.Close($false) }
    if ($null -ne $excel) { $excel.Quit() }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
if ((Get-FileHash -LiteralPath $path).Hash -ne $before) { throw 'Read-only inspection changed the workbook.' }
$records['SHA256'] = $before
$records | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath (Join-Path $PSScriptRoot 'SavedCase.json') -Encoding UTF8
Write-Output "SAVED_CASE_READ_ONLY_OK $before"
