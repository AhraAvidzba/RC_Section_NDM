# Публикует только проверенное оформление; текущая книга защищена от перезаписи новых правок.
$ErrorActionPreference = 'Stop'
$root = (Get-Location).Path
$directory = Join-Path $root 'docs/regression/ExcelGuardGuide_2026-10-13/FormattingLayout_2026-10-15'
function Receipt([string]$name) { Get-Content -LiteralPath (Join-Path $directory $name) -Raw -Encoding UTF8 | ConvertFrom-Json }
$before = Receipt 'Before.json'
$preserved = Receipt 'Preservation.json'
$native = Receipt 'NativeCheck.json'
if (-not $preserved.Passed -or -not $preserved.VbaProjectByteIdentical -or -not $native.Passed) { throw 'Layout acceptance failed.' }
$candidate = Join-Path $directory 'Candidate.xlsm'
$book = Join-Path $root 'workbook/output/RC_Section_NDM.xlsm'
$code = Join-Path $root 'workbook/output/VBA_All_Code.txt'
$report = Join-Path $root 'workbook/output/RC_Section_NDM_execution_report.txt'
if ((Get-FileHash -LiteralPath $candidate).Hash.ToLowerInvariant() -cne $preserved.WorkbookSHA256) { throw 'Candidate changed after acceptance.' }
if ((Get-FileHash -LiteralPath $code).Hash -cne $before.AllCodeSHA256 -or (Get-FileHash -LiteralPath $report).Hash -cne $before.ExecutionReportSHA256) { throw 'User code/report changed during verification.' }
[byte[]]$bytes = [IO.File]::ReadAllBytes($candidate)
[byte[]]$rollback = [IO.File]::ReadAllBytes((Join-Path $directory 'Baseline.xlsm'))
$stream = [IO.File]::Open($book, 'Open', 'ReadWrite', 'None')
$digest = [Security.Cryptography.SHA256]::Create()
try {
    $hash = [BitConverter]::ToString($digest.ComputeHash($stream)).Replace('-', '')
    if ($hash -cne $before.WorkbookSHA256) { throw 'Production workbook changed; preserve the new user data.' }
    try {
        $stream.Position = 0; $stream.Write($bytes, 0, $bytes.Length)
        $stream.SetLength($bytes.Length); $stream.Flush($true)
    } catch {
        $stream.Position = 0; $stream.Write($rollback, 0, $rollback.Length)
        $stream.SetLength($rollback.Length); $stream.Flush($true)
        throw
    }
} finally { $stream.Dispose(); $digest.Dispose() }
if ((Get-FileHash -LiteralPath $book).Hash.ToLowerInvariant() -cne $preserved.WorkbookSHA256) { throw 'Publication hash mismatch.' }
[ordered]@{
    Status = 'CONFIG_LAYOUT_PUBLISHED'; PublishedAt = (Get-Date).ToString('o')
    WorkbookSHA256 = (Get-FileHash -LiteralPath $book).Hash
    AllCodeSHA256 = (Get-FileHash -LiteralPath $code).Hash
    AllCodeUnchanged = $true; UserExecutionReportUnchanged = $true
} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $directory 'Publication.json') -Encoding UTF8
Write-Output 'CONFIG_LAYOUT_PUBLISHED'
