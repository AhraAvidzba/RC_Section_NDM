# Публикует проверенный VBA и Справку, сохраняя исходные данные и результаты пользователя.
param([switch]$Readback)
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$directory=Join-Path $root 'docs/regression/SharedRebarInput_2026-10-18/OffsetExport_2026-10-19'
function Receipt([string]$name){Get-Content -LiteralPath (Join-Path $directory $name) -Raw -Encoding UTF8 | ConvertFrom-Json}
function Hash([string]$path){(Get-FileHash -LiteralPath $path).Hash}
function StreamHash($stream){$digest=[Security.Cryptography.SHA256]::Create();try{$stream.Position=0;[BitConverter]::ToString($digest.ComputeHash($stream)).Replace('-','')}finally{$digest.Dispose()}}
$bookPath=Join-Path $root 'workbook/output/RC_Section_NDM.xlsm'
$codePath=Join-Path $root 'workbook/output/VBA_All_Code.txt'
$reportPath=Join-Path $root 'workbook/output/RC_Section_NDM_execution_report.txt'
$package=Receipt 'Package.json'
if($package.Status -cne 'PRESERVATION_PASS' -or -not $package.AllInputAndResultCellsPreserved -or -not $package.AllStylesDimensionsAndObjectsPreserved){throw 'Preservation check is incomplete.'}
foreach($name in @('AcceptedV3Equality.json','KernelMath_Final/Comparison.json','SavedNative_Verified/Acceptance.json')){
    if(-not (Receipt $name).Passed){throw "Acceptance failed: $name"}
}
foreach($mode in @('Yes','No')){
    $log=[IO.File]::ReadAllText((Join-Path $directory "Release_$mode.txt"),[Text.Encoding]::UTF8)
    if(-not $log.Contains('WATCHDOG_COMPLETED: exit=0; source unchanged=True')){throw "Directed mode incomplete: $mode"}
}
$reportHash='9426FB7EC9A0A5CED9CB75021E71F7967DE0384733E908CC8D720BE5E9AE865E'
if((Hash $reportPath) -cne $reportHash){throw 'User execution report changed before publication.'}
$accepted=Join-Path $directory 'AcceptedV3.xlsm'
if((Hash $accepted) -cne $package.AcceptedSHA256){throw 'Accepted workbook changed.'}
$bookStream=$null;$codeStream=$null
try{
    if($Readback){
        $publication=Receipt 'Publication.json'
        if(-not (Receipt 'PublishedEquality.json').Passed){throw 'Published VBA differs from source.'}
        [byte[]]$newCode=[IO.File]::ReadAllBytes((Join-Path $directory 'PublishedNative.txt'))
        $bookStream=[IO.File]::Open($bookPath,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::Read)
        $codeStream=[IO.File]::Open($codePath,[IO.FileMode]::Open,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None)
        if((StreamHash $bookStream) -cne $package.AcceptedSHA256 -or (StreamHash $codeStream) -cne $publication.AllCodeSHA256){throw 'Published files changed before readback.'}
        $codeStream.Position=0;$codeStream.Write($newCode,0,$newCode.Length);$codeStream.SetLength($newCode.Length);$codeStream.Flush($true)
    }else{
        $native=[IO.File]::ReadAllText((Join-Path $directory 'AcceptedV3Native.txt'),[Text.Encoding]::UTF8)
        if(([regex]::Matches($native,'(?m)^Workbook: AcceptedV3.xlsm\r?$')).Count -ne 1){throw 'Unexpected native VBA export.'}
        $native=$native.Replace('Workbook: AcceptedV3.xlsm','Workbook: RC_Section_NDM.xlsm')
        $encoding=New-Object Text.UTF8Encoding($true)
        [byte[]]$newCode=$encoding.GetPreamble()+$encoding.GetBytes($native)
        [byte[]]$newBook=[IO.File]::ReadAllBytes($accepted)
        [byte[]]$oldBook=[IO.File]::ReadAllBytes((Join-Path $directory 'Saved.xlsm'))
        [byte[]]$oldCode=[IO.File]::ReadAllBytes($codePath)
        $bookStream=[IO.File]::Open($bookPath,[IO.FileMode]::Open,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None)
        $codeStream=[IO.File]::Open($codePath,[IO.FileMode]::Open,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None)
        if((StreamHash $bookStream) -cne $package.SavedSHA256 -or (StreamHash $codeStream) -cne $package.CodeSHA256){throw 'User workbook or ALL CODE changed.'}
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
    UserExecutionReportUnchanged=((Hash $reportPath) -ceq $reportHash);PublishedAt=(Get-Date).ToString('o')} |
    ConvertTo-Json | Set-Content -LiteralPath (Join-Path $directory 'Publication.json') -Encoding UTF8
Write-Output ('OFFSET_EXPORT_PUBLISHED: readback='+[bool]$Readback)
