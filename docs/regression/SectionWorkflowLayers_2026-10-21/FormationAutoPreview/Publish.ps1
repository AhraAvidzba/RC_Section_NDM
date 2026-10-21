# Выпускает проверенную книгу и её фактический VBA; посторонний отчёт не меняет.
param([switch]$Readback)
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../../../..')).Path
function Hash([string]$path){(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash}
function Receipt([string]$name){Get-Content -LiteralPath (Join-Path $PSScriptRoot $name) -Raw -Encoding UTF8 | ConvertFrom-Json}
function StreamHash($stream){
 $hash=[Security.Cryptography.SHA256]::Create()
 try{$stream.Position=0;[BitConverter]::ToString($hash.ComputeHash($stream)).Replace('-','')}finally{$hash.Dispose()}
}
$bookPath=Join-Path $root 'workbook/output/RC_Section_NDM.xlsm'
$codePath=Join-Path $root 'workbook/output/VBA_All_Code.txt'
$reportPath=Join-Path $root 'workbook/output/RC_Section_NDM_execution_report.txt'
$reportHash=(Receipt 'Package.json').UserExecutionReportSHA256
$oldCodeHash='2304F0103C3E77DB711EDC90567BB2236449EAF5236663099F8DB02C2E1C910A'
$package=Receipt 'Package.json'
$acceptance=Receipt 'Acceptance.json'
if(-not $package.Passed -or -not $package.AllInputAndResultCellsPreserved -or -not $acceptance.Passed){throw 'Incomplete acceptance.'}
if((Hash $reportPath) -cne $reportHash){throw 'User execution report changed; publication stopped.'}
$bookStream=$null;$codeStream=$null
try{
 if($Readback){
  $publication=Receipt 'Publication.json'
  if(-not (Receipt 'PublishedEquality.json').Passed){throw 'Published VBA differs from source.'}
  [byte[]]$newCode=[IO.File]::ReadAllBytes((Join-Path $PSScriptRoot 'PublishedNative.txt'))
  $bookStream=[IO.File]::Open($bookPath,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::Read)
  $codeStream=[IO.File]::Open($codePath,[IO.FileMode]::Open,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None)
  if((StreamHash $bookStream) -cne $package.AcceptedSHA256 -or (StreamHash $codeStream) -cne $publication.AllCodeSHA256){throw 'Published files changed.'}
  $codeStream.Position=0;$codeStream.Write($newCode,0,$newCode.Length);$codeStream.SetLength($newCode.Length);$codeStream.Flush($true)
 }else{
  if(-not (Receipt 'AcceptedEquality.json').Passed){throw 'Candidate VBA differs from source.'}
  $candidate=Join-Path $PSScriptRoot 'Accepted.xlsm'
  if((Hash $candidate) -cne $package.AcceptedSHA256){throw 'Candidate changed.'}
  $text=[IO.File]::ReadAllText((Join-Path $PSScriptRoot 'AcceptedNative.txt'),[Text.Encoding]::UTF8)
  if(([regex]::Matches($text,'(?m)^Workbook: Accepted.xlsm\r?$')).Count -ne 1){throw 'Unexpected native export.'}
  $text=$text.Replace('Workbook: Accepted.xlsm','Workbook: RC_Section_NDM.xlsm')
  $encoding=New-Object Text.UTF8Encoding($true)
  [byte[]]$newCode=$encoding.GetPreamble()+$encoding.GetBytes($text)
  [byte[]]$newBook=[IO.File]::ReadAllBytes($candidate)
  $bookStream=[IO.File]::Open($bookPath,[IO.FileMode]::Open,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None)
  $codeStream=[IO.File]::Open($codePath,[IO.FileMode]::Open,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None)
  if((StreamHash $bookStream) -cne $package.OriginalSHA256 -or (StreamHash $codeStream) -cne $oldCodeHash){throw 'User workbook or ALL CODE changed.'}
  $oldBook=New-Object byte[] $bookStream.Length;$bookStream.Position=0;[void]$bookStream.Read($oldBook,0,$oldBook.Length)
  $oldCode=New-Object byte[] $codeStream.Length;$codeStream.Position=0;[void]$codeStream.Read($oldCode,0,$oldCode.Length)
  try{
   $bookStream.Position=0;$bookStream.Write($newBook,0,$newBook.Length);$bookStream.SetLength($newBook.Length);$bookStream.Flush($true)
   $codeStream.Position=0;$codeStream.Write($newCode,0,$newCode.Length);$codeStream.SetLength($newCode.Length);$codeStream.Flush($true)
  }catch{
   $bookStream.Position=0;$bookStream.Write($oldBook,0,$oldBook.Length);$bookStream.SetLength($oldBook.Length);$bookStream.Flush($true)
   $codeStream.Position=0;$codeStream.Write($oldCode,0,$oldCode.Length);$codeStream.SetLength($oldCode.Length);$codeStream.Flush($true)
   throw
  }
 }
}finally{if($codeStream){$codeStream.Dispose()};if($bookStream){$bookStream.Dispose()}}
[ordered]@{Passed=$true;Readback=[bool]$Readback;WorkbookSHA256=(Hash $bookPath);AllCodeSHA256=(Hash $codePath);
 UserExecutionReportUnchanged=((Hash $reportPath) -ceq $reportHash);PublishedAt=(Get-Date).ToString('o')} |
 ConvertTo-Json | Set-Content -LiteralPath (Join-Path $PSScriptRoot 'Publication.json') -Encoding UTF8
Write-Output ('FORMATION_PREVIEW_PUBLISHED: readback='+[bool]$Readback)
