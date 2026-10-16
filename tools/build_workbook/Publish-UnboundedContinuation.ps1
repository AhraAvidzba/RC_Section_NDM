# Публикует проверенные книгу и ALL CODE, затем принимает независимое чтение VBA.
param([switch]$Readback)
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$directory=Join-Path $root 'docs/regression/ExcelGuardGuide_2026-10-13/ExtensionLimit100_2026-10-16'
function Receipt([string]$name){Get-Content -LiteralPath (Join-Path $directory $name) -Raw -Encoding UTF8 | ConvertFrom-Json}
function Hash([string]$path){(Get-FileHash -LiteralPath $path).Hash}
function StreamHash($stream){$digest=[Security.Cryptography.SHA256]::Create();try{[BitConverter]::ToString($digest.ComputeHash($stream)).Replace('-','')}finally{$digest.Dispose()}}
$bookPath=Join-Path $root 'workbook/output/RC_Section_NDM.xlsm'
$codePath=Join-Path $root 'workbook/output/VBA_All_Code.txt'
$reportPath=Join-Path $root 'workbook/output/RC_Section_NDM_execution_report.txt'
$before=Receipt 'Before.json';$package=Receipt 'FinalPackage.json';$acceptance=Receipt 'Acceptance.json'
if(-not $package.Passed -or -not $acceptance.Passed){throw 'Continuation acceptance is incomplete.'}
if((Hash $reportPath) -cne $before.Hashes.'workbook/output/RC_Section_NDM_execution_report.txt'){throw 'Preserve the changed user report before publication.'}
$accepted=Join-Path $directory 'AcceptedFinal.xlsm'
if((Hash $accepted) -cne $package.AcceptedSHA256){throw 'Accepted workbook changed.'}
$bookStream=$null;$codeStream=$null
try {
    if($Readback){
        $publication=Receipt 'Publication.json'
        if(-not (Receipt 'PublishedSourceEquality.json').Passed){throw 'Published VBA does not match source.'}
        $newCode=[IO.File]::ReadAllBytes((Join-Path $directory 'PublishedNativeVBA.txt'))
        $bookStream=[IO.File]::Open($bookPath,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::Read)
        $codeStream=[IO.File]::Open($codePath,[IO.FileMode]::Open,[IO.FileAccess]::ReadWrite,[IO.FileShare]::Read)
        if((StreamHash $bookStream) -cne $package.AcceptedSHA256 -or (StreamHash $codeStream) -cne $publication.AllCodeSHA256){throw 'Published pair changed before readback.'}
        $codeStream.Position=0;$codeStream.Write($newCode,0,$newCode.Length);$codeStream.SetLength($newCode.Length);$codeStream.Flush($true)
    } else {
        if(-not (Receipt 'FinalCandidateSourceEquality.json').Passed){throw 'Compiled reference VBA does not match source.'}
        $native=[IO.File]::ReadAllText((Join-Path $directory 'FinalCandidateNativeVBA.txt'),[Text.Encoding]::UTF8)
        if(([regex]::Matches($native,'(?m)^Workbook: Continuation.xlsm\r?$')).Count -ne 1){throw 'Unexpected compiled reference dump.'}
        $native=$native.Replace('Workbook: Continuation.xlsm','Workbook: RC_Section_NDM.xlsm')
        $encoding=New-Object Text.UTF8Encoding($true)
        [byte[]]$newCode=$encoding.GetPreamble()+$encoding.GetBytes($native)
        [byte[]]$newBook=[IO.File]::ReadAllBytes($accepted)
        [byte[]]$oldBook=[IO.File]::ReadAllBytes((Join-Path $directory 'Baseline.xlsm'))
        [byte[]]$oldCode=[IO.File]::ReadAllBytes((Join-Path $directory 'BaselineAllCode.txt'))
        $bookStream=[IO.File]::Open($bookPath,[IO.FileMode]::Open,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None)
        $codeStream=[IO.File]::Open($codePath,[IO.FileMode]::Open,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None)
        if((StreamHash $bookStream) -cne $before.Hashes.'workbook/output/RC_Section_NDM.xlsm' -or
           (StreamHash $codeStream) -cne $before.Hashes.'workbook/output/VBA_All_Code.txt'){throw 'User workbook or ALL CODE changed; reconcile before publication.'}
        try {
            $bookStream.Position=0;$bookStream.Write($newBook,0,$newBook.Length);$bookStream.SetLength($newBook.Length);$bookStream.Flush($true)
            $codeStream.Position=0;$codeStream.Write($newCode,0,$newCode.Length);$codeStream.SetLength($newCode.Length);$codeStream.Flush($true)
        } catch {
            $bookStream.Position=0;$bookStream.Write($oldBook,0,$oldBook.Length);$bookStream.SetLength($oldBook.Length);$bookStream.Flush($true)
            $codeStream.Position=0;$codeStream.Write($oldCode,0,$oldCode.Length);$codeStream.SetLength($oldCode.Length);$codeStream.Flush($true)
            throw
        }
    }
} finally {
    if($codeStream){$codeStream.Dispose()}
    if($bookStream){$bookStream.Dispose()}
}
[ordered]@{Published=$true;Readback=[bool]$Readback;WorkbookSHA256=(Hash $bookPath);AllCodeSHA256=(Hash $codePath);
    UserExecutionReportUnchanged=((Hash $reportPath) -ceq $before.Hashes.'workbook/output/RC_Section_NDM_execution_report.txt');
    PublishedAt=(Get-Date).ToString('o')} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $directory 'Publication.json') -Encoding UTF8
Write-Output ('CONTINUATION_PUBLISHED: readback='+[bool]$Readback)
