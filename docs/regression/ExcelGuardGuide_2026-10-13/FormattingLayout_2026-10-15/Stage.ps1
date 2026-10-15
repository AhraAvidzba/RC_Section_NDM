# Сохраняет актуальные пользовательские значения перед точечной правкой оформления.
$ErrorActionPreference = 'Stop'
$root = (Get-Location).Path
$directory = Join-Path $root 'docs/regression/ExcelGuardGuide_2026-10-13/FormattingLayout_2026-10-15'
$bookPath = Join-Path $root 'workbook/output/RC_Section_NDM.xlsm'
$baseline = Join-Path $directory 'Baseline.xlsm'
if (Test-Path -LiteralPath $baseline) { throw 'Baseline already exists.' }
$lock = $null
try { $lock = [IO.File]::Open($bookPath, 'Open', 'Read', 'None') } catch [IO.IOException] {}
$savedOpenBook = $false
if ($lock) { $lock.Dispose() } else {
    $book = [Runtime.InteropServices.Marshal]::BindToMoniker($bookPath)
    try {
        if ([string]$book.FullName -cne $bookPath -or $book.ReadOnly -or -not $book.Application.Ready) { throw 'Unexpected or unavailable workbook.' }
        if (-not $book.Saved) { $book.Save(); $savedOpenBook = $true }
        $book.Close($false)
    } finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($book) }
}
Copy-Item -LiteralPath $bookPath -Destination $baseline
[ordered]@{
    SavedOpenWorkbook = $savedOpenBook
    WorkbookSHA256 = (Get-FileHash -LiteralPath $bookPath).Hash
    AllCodeSHA256 = (Get-FileHash -LiteralPath (Join-Path $root 'workbook/output/VBA_All_Code.txt')).Hash
    ExecutionReportSHA256 = (Get-FileHash -LiteralPath (Join-Path $root 'workbook/output/RC_Section_NDM_execution_report.txt')).Hash
} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $directory 'Before.json') -Encoding UTF8
Write-Output 'USER_BASELINE_CAPTURED'
