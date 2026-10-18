# Заменяет рабочую книгу принятой копией и её фактическим VBA. Отчёт пользователя сохраняется.
param([switch]$Readback)
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$directory=Join-Path $root 'docs/regression/SharedRebarInput_2026-10-18'
function Receipt([string]$name){Get-Content -LiteralPath (Join-Path $directory $name) -Raw -Encoding UTF8 | ConvertFrom-Json}
function Hash([string]$path){(Get-FileHash -LiteralPath $path).Hash}
function StreamHash($stream){$digest=[Security.Cryptography.SHA256]::Create();try{$stream.Position=0;[BitConverter]::ToString($digest.ComputeHash($stream)).Replace('-','')}finally{$digest.Dispose()}}
$bookPath=Join-Path $root 'workbook/output/RC_Section_NDM.xlsm'
$codePath=Join-Path $root 'workbook/output/VBA_All_Code.txt'
$reportPath=Join-Path $root 'workbook/output/RC_Section_NDM_execution_report.txt'
$before=Receipt 'Before.json';$acceptance=Receipt 'Acceptance.json'
if(-not $acceptance.Passed -or -not (Receipt 'AcceptedSourceEquality.json').Passed){throw 'Shared reinforcement acceptance is incomplete.'}
if((Hash $reportPath) -cne $before.'workbook/output/RC_Section_NDM_execution_report.txt'){throw 'User report changed before publication.'}
$accepted=Join-Path $directory 'Accepted.xlsm'
if((Hash $accepted) -cne $acceptance.AcceptedSHA256){throw 'Accepted workbook changed.'}
$bookStream=$null;$codeStream=$null
try{
    if($Readback){
        $publication=Receipt 'Publication.json'
        if(-not (Receipt 'PublishedSourceEquality.json').Passed){throw 'Published VBA differs from source.'}
        $newCode=[IO.File]::ReadAllBytes((Join-Path $directory 'PublishedNativeVBA.txt'))
        $bookStream=[IO.File]::Open($bookPath,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::Read)
        $codeStream=[IO.File]::Open($codePath,[IO.FileMode]::Open,[IO.FileAccess]::ReadWrite,[IO.FileShare]::Read)
        if((StreamHash $bookStream) -cne $acceptance.AcceptedSHA256 -or (StreamHash $codeStream) -cne $publication.AllCodeSHA256){throw 'Published files changed before readback.'}
        $codeStream.Position=0;$codeStream.Write($newCode,0,$newCode.Length);$codeStream.SetLength($newCode.Length);$codeStream.Flush($true)
    }else{
        $native=[IO.File]::ReadAllText((Join-Path $directory 'AcceptedFinalNativeVBA.txt'),[Text.Encoding]::UTF8)
        if(([regex]::Matches($native,'(?m)^Workbook: Accepted.xlsm\r?$')).Count -ne 1){throw 'Unexpected native VBA export.'}
        $native=$native.Replace('Workbook: Accepted.xlsm','Workbook: RC_Section_NDM.xlsm')
        $encoding=New-Object Text.UTF8Encoding($true)
        [byte[]]$newCode=$encoding.GetPreamble()+$encoding.GetBytes($native)
        [byte[]]$newBook=[IO.File]::ReadAllBytes($accepted)
        [byte[]]$oldBook=[IO.File]::ReadAllBytes((Join-Path $directory 'Baseline.xlsm'))
        [byte[]]$oldCode=[IO.File]::ReadAllBytes($codePath)
        $bookStream=[IO.File]::Open($bookPath,[IO.FileMode]::Open,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None)
        $codeStream=[IO.File]::Open($codePath,[IO.FileMode]::Open,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None)
        if((StreamHash $bookStream) -cne $before.'workbook/output/RC_Section_NDM.xlsm' -or (StreamHash $codeStream) -cne $before.'workbook/output/VBA_All_Code.txt'){throw 'User workbook or ALL CODE changed.'}
        try{
            $bookStream.Position=0;$bookStream.Write($newBook,0,$newBook.Length);$bookStream.SetLength($newBook.Length);$bookStream.Flush($true)
            $codeStream.Position=0;$codeStream.Write($newCode,0,$newCode.Length);$codeStream.SetLength($newCode.Length);$codeStream.Flush($true)
        }catch{
            $bookStream.Position=0;$bookStream.Write($oldBook,0,$oldBook.Length);$bookStream.SetLength($oldBook.Length);$bookStream.Flush($true)
            $codeStream.Position=0;$codeStream.Write($oldCode,0,$oldCode.Length);$codeStream.SetLength($oldCode.Length);$codeStream.Flush($true)
            throw
        }
    }
}finally{
    if($codeStream){$codeStream.Dispose()}
    if($bookStream){$bookStream.Dispose()}
}
[ordered]@{Passed=$true;Readback=[bool]$Readback;WorkbookSHA256=(Hash $bookPath);AllCodeSHA256=(Hash $codePath);
    UserExecutionReportUnchanged=((Hash $reportPath) -ceq $before.'workbook/output/RC_Section_NDM_execution_report.txt');PublishedAt=(Get-Date).ToString('o')} |
    ConvertTo-Json | Set-Content -LiteralPath (Join-Path $directory 'Publication.json') -Encoding UTF8
Write-Output ('SHARED_REBAR_PUBLISHED: readback='+[bool]$Readback)
