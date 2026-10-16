# Публикует только проверенную пару книга/ALL CODE с исключительной блокировкой файлов.
param([string]$EvidenceDirectory='docs/regression/ExcelGuardGuide_2026-10-13/RecipeSimplification_2026-10-16')
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$directory=(Resolve-Path -LiteralPath (Join-Path $root $EvidenceDirectory)).Path
if(-not $directory.StartsWith((Join-Path $root 'docs/regression')+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)){throw 'Unexpected publication evidence directory.'}
function Receipt([string]$name){Get-Content -LiteralPath (Join-Path $directory $name) -Raw -Encoding UTF8 | ConvertFrom-Json}
$before=Receipt 'Before.json'
$package=Receipt 'Package.json'
$acceptance=Receipt 'Acceptance.json'
if($acceptance.Status -cne 'RECIPE_SIMPLIFICATION_PASS' -or $package.Status -cne 'RECIPE_PACKAGE_PASS'){throw 'Acceptance is incomplete.'}
foreach($name in @('AcceptedSourceEquality.json','AcceptedGuidePresentation.json')){if(-not (Receipt $name).Passed){throw ('Acceptance failed: '+$name)}}
$bookPath=Join-Path $root 'workbook/output/RC_Section_NDM.xlsm'
$codePath=Join-Path $root 'workbook/output/VBA_All_Code.txt'
$reportPath=Join-Path $root 'workbook/output/RC_Section_NDM_execution_report.txt'
$accepted=Join-Path $directory 'Accepted.xlsm'
$dump=Join-Path $directory 'AcceptedNativeVBA.txt'
if((Get-FileHash -LiteralPath $accepted).Hash.ToLowerInvariant() -cne $package.AcceptedSHA256){throw 'Accepted workbook changed.'}
if((Get-FileHash -LiteralPath $dump).Hash.ToLowerInvariant() -cne (Receipt 'AcceptedGuidePresentation.json').NativeSHA256){throw 'Accepted native dump changed.'}
if((Get-FileHash -LiteralPath $reportPath).Hash -cne $before.UserReportSHA256){throw 'User execution report changed; preserve it.'}
$raw=[IO.File]::ReadAllText($dump,[Text.Encoding]::UTF8)
if($raw -notmatch '(?m)^Workbook: Accepted.xlsm\r?$'){throw 'Unexpected native dump identity.'}
$raw=$raw.Replace('Workbook: Accepted.xlsm','Workbook: RC_Section_NDM.xlsm')
$utf8=New-Object Text.UTF8Encoding($true)
[byte[]]$newCode=$utf8.GetPreamble()+$utf8.GetBytes($raw)
[byte[]]$newBook=[IO.File]::ReadAllBytes($accepted)
[byte[]]$oldBook=[IO.File]::ReadAllBytes((Join-Path $directory 'Baseline.xlsm'))
[byte[]]$oldCode=[IO.File]::ReadAllBytes((Join-Path $directory 'BaselineAllCode.txt'))
$bookStream=$null;$codeStream=$null;$digest=[Security.Cryptography.SHA256]::Create()
try {
    $bookStream=[IO.File]::Open($bookPath,[IO.FileMode]::Open,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None)
    $codeStream=[IO.File]::Open($codePath,[IO.FileMode]::Open,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None)
    $bookHash=[BitConverter]::ToString($digest.ComputeHash($bookStream)).Replace('-','')
    $codeHash=[BitConverter]::ToString($digest.ComputeHash($codeStream)).Replace('-','')
    if($bookHash -cne $before.WorkbookSHA256 -or $codeHash -cne $before.AllCodeSHA256){throw 'User workbook/code changed; preserve and reconcile it before publication.'}
    try {
        $bookStream.Position=0;$bookStream.Write($newBook,0,$newBook.Length);$bookStream.SetLength($newBook.Length);$bookStream.Flush($true)
        $codeStream.Position=0;$codeStream.Write($newCode,0,$newCode.Length);$codeStream.SetLength($newCode.Length);$codeStream.Flush($true)
    } catch {
        $bookStream.Position=0;$bookStream.Write($oldBook,0,$oldBook.Length);$bookStream.SetLength($oldBook.Length);$bookStream.Flush($true)
        $codeStream.Position=0;$codeStream.Write($oldCode,0,$oldCode.Length);$codeStream.SetLength($oldCode.Length);$codeStream.Flush($true)
        throw
    }
} finally {
    if($codeStream){$codeStream.Dispose()}
    if($bookStream){$bookStream.Dispose()}
    $digest.Dispose()
}
[ordered]@{Status='PUBLISHED_READBACK_PENDING';WorkbookSHA256=(Get-FileHash -LiteralPath $bookPath).Hash;AllCodeSHA256=(Get-FileHash -LiteralPath $codePath).Hash;UserExecutionReportUnchanged=((Get-FileHash -LiteralPath $reportPath).Hash -ceq $before.UserReportSHA256);PublishedAt=(Get-Date).ToString('o')} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $directory 'Publication.json') -Encoding UTF8
Write-Output 'RECIPE_PAIR_PUBLISHED: native readback remains.'
