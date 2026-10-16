# Переносит принятую пару книга/ALL CODE, проверяя исходные хеши и сохраняя пользовательский отчёт.
$ErrorActionPreference='Stop'
$root=(Get-Location).Path;$d=$PSScriptRoot
function Receipt([string]$name){Get-Content -LiteralPath (Join-Path $d $name) -Raw -Encoding UTF8 | ConvertFrom-Json}
$before=Receipt 'Before.json';$package=Receipt 'Package.json'
if((Receipt 'Acceptance.json').Status -cne 'RECIPE_BOUNDARIES_PASS' -or $package.Status -cne 'RECIPE_BOUNDARIES_PACKAGE_PASS'){throw 'Acceptance incomplete'}
foreach($name in @('AcceptedSourceEquality.json','AcceptedGuidePresentation.json')){if(-not (Receipt $name).Passed){throw ('Verification failed: '+$name)}}
$bookPath=Join-Path $root 'workbook/output/RC_Section_NDM.xlsm';$codePath=Join-Path $root 'workbook/output/VBA_All_Code.txt'
$reportPath=Join-Path $root 'workbook/output/RC_Section_NDM_execution_report.txt'
if((Get-FileHash -LiteralPath $reportPath).Hash -cne $before.Hashes.'workbook/output/RC_Section_NDM_execution_report.txt'){throw 'Preserve the changed user report'}
$accepted=Join-Path $d 'Accepted.xlsm'
if((Get-FileHash -LiteralPath $accepted).Hash.ToLowerInvariant() -cne $package.AcceptedSHA256){throw 'Accepted book changed'}
[byte[]]$newBook=[IO.File]::ReadAllBytes($accepted)
[byte[]]$oldBook=[IO.File]::ReadAllBytes((Join-Path $d 'Baseline.xlsm'))
[byte[]]$oldCode=[IO.File]::ReadAllBytes((Join-Path $d 'BaselineNativeVBA.txt'))
$raw=[IO.File]::ReadAllText((Join-Path $d 'AcceptedNativeVBA.txt'),[Text.Encoding]::UTF8)
if($raw -notmatch '(?m)^Workbook: Accepted.xlsm\r?$'){throw 'Unexpected export identity'}
$encoding=New-Object Text.UTF8Encoding($true)
[byte[]]$newCode=$encoding.GetPreamble()+$encoding.GetBytes($raw.Replace('Workbook: Accepted.xlsm','Workbook: RC_Section_NDM.xlsm'))
$bookStream=$null;$codeStream=$null;$digest=[Security.Cryptography.SHA256]::Create()
try {
    $bookStream=[IO.File]::Open($bookPath,'Open','ReadWrite','None');$codeStream=[IO.File]::Open($codePath,'Open','ReadWrite','None')
    $bookHash=[BitConverter]::ToString($digest.ComputeHash($bookStream)).Replace('-','')
    $codeHash=[BitConverter]::ToString($digest.ComputeHash($codeStream)).Replace('-','')
    if($bookHash -cne $before.Hashes.'workbook/output/RC_Section_NDM.xlsm' -or $codeHash -cne $before.Hashes.'workbook/output/VBA_All_Code.txt'){throw 'User files changed; preserve them'}
    try {
        $bookStream.Position=0;$bookStream.Write($newBook,0,$newBook.Length);$bookStream.SetLength($newBook.Length);$bookStream.Flush($true)
        $codeStream.Position=0;$codeStream.Write($newCode,0,$newCode.Length);$codeStream.SetLength($newCode.Length);$codeStream.Flush($true)
    } catch {
        $bookStream.Position=0;$bookStream.Write($oldBook,0,$oldBook.Length);$bookStream.SetLength($oldBook.Length);$bookStream.Flush($true)
        $codeStream.Position=0;$codeStream.Write($oldCode,0,$oldCode.Length);$codeStream.SetLength($oldCode.Length);$codeStream.Flush($true)
        throw
    }
} finally {if($bookStream){$bookStream.Dispose()};if($codeStream){$codeStream.Dispose()};$digest.Dispose()}
[ordered]@{Status='PUBLISHED_READBACK_PENDING';WorkbookSHA256=(Get-FileHash -LiteralPath $bookPath).Hash.ToLowerInvariant();
    AllCodeSHA256=(Get-FileHash -LiteralPath $codePath).Hash.ToLowerInvariant();UserReportUnchanged=$true;NativeComponents=158;PublishedAt=(Get-Date).ToString('o')} |
    ConvertTo-Json | Set-Content -LiteralPath (Join-Path $d 'Publication.json') -Encoding UTF8
Write-Output 'RECIPE_BOUNDARIES_PUBLISHED: native readback remains'
