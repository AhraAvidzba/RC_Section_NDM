# Публикует только принятый кандидат; двойной hash-guard сохраняет текущие пользовательские данные.
$ErrorActionPreference='Stop'
$root=(Get-Location).Path
$directory=Join-Path $root 'docs/regression/ExcelGuardGuide_2026-10-13/CircleEverySecond_2026-10-14'
function Receipt([string]$name) { Get-Content -LiteralPath (Join-Path $directory $name) -Raw -Encoding UTF8 | ConvertFrom-Json }
$before=Receipt 'Before.json'
$preservation=Receipt 'Preservation.json'
$equality=Receipt 'FinalSourceEquality.json'
$help=Receipt 'StandaloneHelp.json'
$directed=Receipt 'Directed01/Run.json'
$finalDirected=Receipt 'FinalWorkbookDirected/Run.json'
$factory=Receipt 'FactoryTest.json'
$migration=Receipt 'RecipeConfigUpdate.json'
if(-not $preservation.Passed -or -not $preservation.NativeChangesRestricted -or -not $preservation.NoEmbeddedMarkers -or -not $preservation.UserDataFormulasStylesNamesAndDrawingsPreserved -or -not $preservation.ModelsEqualExcept72NewInputCells -or -not $equality.Passed -or $help.status -ne 'STANDALONE_HELP_PASS' -or -not $directed.Passed -or -not $finalDirected.Passed -or -not $factory.Passed -or -not $migration.Idempotent){throw 'Circle binding acceptance is incomplete.'}
$candidate=Join-Path $directory 'Candidate.xlsm'
$dump=Join-Path $directory 'FinalNativeVBA.txt'
$bookPath=Join-Path $root 'workbook/output/RC_Section_NDM.xlsm'
$codePath=Join-Path $root 'workbook/output/VBA_All_Code.txt'
$reportPath=Join-Path $root 'workbook/output/RC_Section_NDM_execution_report.txt'
if((Get-FileHash -LiteralPath $candidate).Hash.ToLowerInvariant() -cne $preservation.CandidateSHA256 -or (Get-FileHash -LiteralPath $dump).Hash.ToLowerInvariant() -cne $preservation.NativeSHA256){throw 'Accepted candidate/dump changed.'}
if((Get-FileHash -LiteralPath $reportPath).Hash -cne $before.ExecutionReportSHA256){throw 'User execution report changed; do not modify it.'}
$lock=$null
try {$lock=[IO.File]::Open($bookPath,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::None)} catch [IO.IOException] {}
if($lock){$lock.Dispose()} else {
    # Пользователь ранее разрешил сохранить и закрыть именно эту книгу; прочие книги остаются открыты.
    $book=[Runtime.InteropServices.Marshal]::BindToMoniker($bookPath)
    try {
        if([string]$book.FullName -cne $bookPath -or -not $book.Application.Ready -or $book.ReadOnly){throw 'Unexpected or unavailable open workbook.'}
        if(-not $book.Saved){$book.Save()}
        $book.Close($false)
    } finally {[void][Runtime.InteropServices.Marshal]::ReleaseComObject($book)}
}
$native=[IO.File]::ReadAllText($dump,[Text.Encoding]::UTF8)
if($native -notmatch '(?m)^Workbook: Candidate.xlsm\r?$'){throw 'Unexpected native dump identity.'}
$native=$native.Replace('Workbook: Candidate.xlsm','Workbook: RC_Section_NDM.xlsm')
$utf8=New-Object Text.UTF8Encoding($true)
[byte[]]$bookBytes=[IO.File]::ReadAllBytes($candidate)
[byte[]]$oldBookBytes=[IO.File]::ReadAllBytes((Join-Path $directory 'Baseline.xlsm'))
[byte[]]$codeBytes=$utf8.GetPreamble()+$utf8.GetBytes($native)
$bookStream=$null;$codeStream=$null;$digest=[Security.Cryptography.SHA256]::Create()
try {
    $bookStream=[IO.File]::Open($bookPath,[IO.FileMode]::Open,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None)
    $codeStream=[IO.File]::Open($codePath,[IO.FileMode]::Open,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None)
    $bookHash=[BitConverter]::ToString($digest.ComputeHash($bookStream)).Replace('-','')
    $codeHash=[BitConverter]::ToString($digest.ComputeHash($codeStream)).Replace('-','')
    if($bookHash -cne $before.WorkbookSHA256 -or $codeHash -cne $before.AllCodeSHA256){throw 'Production data changed since baseline; preserve it and reconcile before publishing.'}
    $codeStream.Position=0
    [byte[]]$oldCodeBytes=New-Object byte[] ([int]$codeStream.Length)
    $read=$codeStream.Read($oldCodeBytes,0,$oldCodeBytes.Length)
    if($read -ne $oldCodeBytes.Length){throw 'Cannot capture previous ALL CODE for rollback.'}
    [IO.File]::WriteAllBytes((Join-Path $directory 'AllCodeBeforePublication.txt'),$oldCodeBytes)
    try {
        $bookStream.Position=0;$bookStream.Write($bookBytes,0,$bookBytes.Length);$bookStream.SetLength($bookBytes.Length);$bookStream.Flush($true)
        $codeStream.Position=0;$codeStream.Write($codeBytes,0,$codeBytes.Length);$codeStream.SetLength($codeBytes.Length);$codeStream.Flush($true)
    } catch {
        $bookStream.Position=0;$bookStream.Write($oldBookBytes,0,$oldBookBytes.Length);$bookStream.SetLength($oldBookBytes.Length);$bookStream.Flush($true)
        $codeStream.Position=0;$codeStream.Write($oldCodeBytes,0,$oldCodeBytes.Length);$codeStream.SetLength($oldCodeBytes.Length);$codeStream.Flush($true)
        throw
    }
} finally {
    if($codeStream){$codeStream.Dispose()}
    if($bookStream){$bookStream.Dispose()}
    $digest.Dispose()
}
if((Get-FileHash -LiteralPath $bookPath).Hash.ToLowerInvariant() -cne $preservation.CandidateSHA256){throw 'Published book differs from accepted candidate.'}
if((Get-FileHash -LiteralPath $reportPath).Hash -cne $before.ExecutionReportSHA256){throw 'Execution report changed during publication.'}
[ordered]@{PublishedAt=(Get-Date).ToString('o');WorkbookSHA256=(Get-FileHash -LiteralPath $bookPath).Hash;AllCodeSHA256=(Get-FileHash -LiteralPath $codePath).Hash;UserExecutionReportUnchanged=$true;ReadbackPending=$true} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $directory 'Publication.json') -Encoding UTF8
Write-Output 'AUDIT02_PAIR_PUBLISHED: actual workbook readback remains.'
