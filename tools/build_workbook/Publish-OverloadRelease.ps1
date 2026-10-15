# Обновляет книгу и ALL CODE только после численной и файловой приёмки.
$ErrorActionPreference='Stop'
$root=(Get-Location).Path
$directory=Join-Path $root 'docs/regression/Performance/Overload_2026-10-15'
$release=Join-Path $directory 'Release'
function Receipt([string]$path){Get-Content -LiteralPath $path -Raw -Encoding UTF8 | ConvertFrom-Json}
$before=Receipt (Join-Path $release 'Before.json')
$package=Receipt (Join-Path $release 'Package.json')
$measurements=Receipt (Join-Path $directory 'MeasurementsAcceptance.json')
$equality=Receipt (Join-Path $root 'docs/regression/ExcelGuardGuide_2026-10-13/Overload_2026-10-15/CandidateSourceEquality.json')
$guards=Get-Content -LiteralPath (Join-Path $release 'OutputGuards.txt') -Raw -Encoding UTF8
if($package.Status -ne 'OVERLOAD_RELEASE_PACKAGE_PASS' -or $measurements.Status -ne 'OVERLOAD_TIMING_ACCEPTANCE_PASS' -or -not $equality.Passed -or $guards -notmatch 'TOTAL_OVERLOAD_INITIAL_POINT: passed=51; failed=0' -or $guards -notmatch 'WATCHDOG_COMPLETED: exit=0'){throw 'Acceptance is incomplete.'}
$candidate=Join-Path $release 'Accepted.xlsm'
$book=Join-Path $root 'workbook/output/RC_Section_NDM.xlsm'
$code=Join-Path $root 'workbook/output/VBA_All_Code.txt'
$report=Join-Path $root 'workbook/output/RC_Section_NDM_execution_report.txt'
if((Get-FileHash -LiteralPath $candidate).Hash.ToLowerInvariant() -cne $package.WorkbookSHA256){throw 'Accepted workbook changed.'}
if((Get-FileHash -LiteralPath $report).Hash -cne $before.ExecutionReportSHA256){throw 'User execution report changed; preserve it.'}
[byte[]]$newBook=[IO.File]::ReadAllBytes($candidate)
[byte[]]$oldBook=[IO.File]::ReadAllBytes((Join-Path $release 'Baseline.xlsm'))
[byte[]]$newCode=[IO.File]::ReadAllBytes((Join-Path $release 'CandidateNativeVBA.txt'))
$bookStream=$null;$codeStream=$null;$digest=[Security.Cryptography.SHA256]::Create()
try{
    $bookStream=[IO.File]::Open($book,'Open','ReadWrite','None')
    $codeStream=[IO.File]::Open($code,'Open','ReadWrite','None')
    $bookHash=[BitConverter]::ToString($digest.ComputeHash($bookStream)).Replace('-','')
    $codeHash=[BitConverter]::ToString($digest.ComputeHash($codeStream)).Replace('-','')
    if($bookHash -cne $before.WorkbookSHA256 -or $codeHash -cne $before.AllCodeSHA256){throw 'New user changes exist; reconcile before publishing.'}
    $codeStream.Position=0
    [byte[]]$oldCode=New-Object byte[] ([int]$codeStream.Length)
    if($codeStream.Read($oldCode,0,$oldCode.Length) -ne $oldCode.Length){throw 'Cannot capture previous ALL CODE.'}
    try{
        $bookStream.Position=0;$bookStream.Write($newBook,0,$newBook.Length);$bookStream.SetLength($newBook.Length);$bookStream.Flush($true)
        $codeStream.Position=0;$codeStream.Write($newCode,0,$newCode.Length);$codeStream.SetLength($newCode.Length);$codeStream.Flush($true)
    }catch{
        $bookStream.Position=0;$bookStream.Write($oldBook,0,$oldBook.Length);$bookStream.SetLength($oldBook.Length);$bookStream.Flush($true)
        $codeStream.Position=0;$codeStream.Write($oldCode,0,$oldCode.Length);$codeStream.SetLength($oldCode.Length);$codeStream.Flush($true)
        throw
    }
}finally{if($codeStream){$codeStream.Dispose()};if($bookStream){$bookStream.Dispose()};$digest.Dispose()}
[ordered]@{
    Status='OVERLOAD_RELEASE_PUBLISHED';PublishedAt=(Get-Date).ToString('o')
    WorkbookSHA256=(Get-FileHash -LiteralPath $book).Hash
    AllCodeSHA256=(Get-FileHash -LiteralPath $code).Hash
    UserExecutionReportUnchanged=$true;ActualWorkbookReadbackPending=$true
} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $release 'Publication.json') -Encoding UTF8
Write-Output 'OVERLOAD_RELEASE_PUBLISHED; actual workbook readback remains.'
