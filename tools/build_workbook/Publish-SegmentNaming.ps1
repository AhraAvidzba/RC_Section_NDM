# Публикует проверенную книгу и экспорт, сохраняя актуальные пользовательские данные и чужой отчёт.
$ErrorActionPreference='Stop'
$root=(Get-Location).Path
$directory=Join-Path $root 'docs/regression/ExcelGuardGuide_2026-10-13/SegmentNaming_2026-10-16'
function Receipt([string]$name){Get-Content -LiteralPath (Join-Path $directory $name) -Raw -Encoding UTF8 | ConvertFrom-Json}
$before=Receipt 'Before.json';$package=Receipt 'Package.json';$acceptance=Receipt 'Acceptance.json'
if($package.Status -cne 'SEGMENT_NAMING_PACKAGE_PASS' -or $acceptance.Status -cne 'SEGMENT_NAMING_PASS'){throw 'Acceptance is incomplete'}
foreach($name in @('AcceptedSourceEquality.json','AcceptedGuidePresentation.json','SourceScope.json')){if(-not (Receipt $name).Passed){throw ('Verification failed: '+$name)}}
$bookPath=Join-Path $root 'workbook/output/RC_Section_NDM.xlsm'
$codePath=Join-Path $root 'workbook/output/VBA_All_Code.txt'
$reportPath=Join-Path $root 'workbook/output/RC_Section_NDM_execution_report.txt'
$accepted=Join-Path $directory 'Accepted.xlsm'
if((Get-FileHash -LiteralPath $accepted).Hash.ToLowerInvariant() -cne $package.AcceptedSHA256){throw 'Accepted workbook changed'}
if((Get-FileHash -LiteralPath $reportPath).Hash -cne $before.Hashes.'workbook/output/RC_Section_NDM_execution_report.txt'){throw 'User report changed; preserve it'}
$raw=[IO.File]::ReadAllText((Join-Path $directory 'AcceptedNativeVBA.txt'),[Text.Encoding]::UTF8)
if($raw -notmatch '(?m)^Workbook: Accepted.xlsm\r?$'){throw 'Unexpected export identity'}
$raw=$raw.Replace('Workbook: Accepted.xlsm','Workbook: RC_Section_NDM.xlsm')
$utf8=New-Object Text.UTF8Encoding($true)
[byte[]]$newCode=$utf8.GetPreamble()+$utf8.GetBytes($raw)
[byte[]]$newBook=[IO.File]::ReadAllBytes($accepted)
[byte[]]$oldBook=[IO.File]::ReadAllBytes((Join-Path $directory 'Baseline.xlsm'))
[byte[]]$oldCode=[IO.File]::ReadAllBytes((Join-Path $directory 'BaselineNativeVBA.txt'))
$bookStream=$null;$codeStream=$null;$digest=[Security.Cryptography.SHA256]::Create()
try {
    $bookStream=[IO.File]::Open($bookPath,'Open','ReadWrite','None')
    $codeStream=[IO.File]::Open($codePath,'Open','ReadWrite','None')
    $bookHash=[BitConverter]::ToString($digest.ComputeHash($bookStream)).Replace('-','')
    $codeHash=[BitConverter]::ToString($digest.ComputeHash($codeStream)).Replace('-','')
    if($bookHash -cne $before.Hashes.'workbook/output/RC_Section_NDM.xlsm' -or $codeHash -cne $before.Hashes.'workbook/output/VBA_All_Code.txt'){throw 'Published files changed; preserve them'}
    try {
        $bookStream.Position=0;$bookStream.Write($newBook,0,$newBook.Length);$bookStream.SetLength($newBook.Length);$bookStream.Flush($true)
        $codeStream.Position=0;$codeStream.Write($newCode,0,$newCode.Length);$codeStream.SetLength($newCode.Length);$codeStream.Flush($true)
    } catch {
        $bookStream.Position=0;$bookStream.Write($oldBook,0,$oldBook.Length);$bookStream.SetLength($oldBook.Length);$bookStream.Flush($true)
        $codeStream.Position=0;$codeStream.Write($oldCode,0,$oldCode.Length);$codeStream.SetLength($oldCode.Length);$codeStream.Flush($true)
        throw
    }
} finally {
    if($codeStream){$codeStream.Dispose()};if($bookStream){$bookStream.Dispose()};$digest.Dispose()
}
[ordered]@{Status='PUBLISHED_READBACK_PENDING';WorkbookSHA256=(Get-FileHash -LiteralPath $bookPath).Hash.ToLowerInvariant();
    AllCodeSHA256=(Get-FileHash -LiteralPath $codePath).Hash.ToLowerInvariant();NativeComponents=158;
    UserDataBaselineSHA256=$package.BaselineSHA256;UserExecutionReportUnchanged=$true;PublishedAt=(Get-Date).ToString('o')} |
    ConvertTo-Json | Set-Content -LiteralPath (Join-Path $directory 'Publication.json') -Encoding UTF8
Write-Output 'SEGMENT_NAMING_PUBLISHED: independent native readback remains.'
