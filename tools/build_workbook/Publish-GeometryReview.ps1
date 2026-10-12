# Публикует проверенную пару книги и VBA; защищает данные и остальные output-файлы.
param([string]$Directory='docs/regression/GeometryReview_2026-10-12')
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$release=[IO.Path]::GetFullPath((Join-Path $root $Directory))
$expected=[IO.Path]::GetFullPath((Join-Path $root 'docs/regression/GeometryReview_2026-10-12'))
if ($release -cne $expected) {throw 'Only the reviewed release directory is allowed.'}
function Json([string]$Name) {Get-Content -LiteralPath (Join-Path $release $Name) -Raw -Encoding UTF8 | ConvertFrom-Json}
function Hash([string]$Path) {(Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant()}
function BytesHash([byte[]]$Bytes) {
    $sha=[Security.Cryptography.SHA256]::Create()
    try {return [BitConverter]::ToString($sha.ComputeHash($Bytes)).Replace('-','').ToLowerInvariant()} finally {$sha.Dispose()}
}
function Read-LockedBytes($Stream) {
    $buffer=New-Object IO.MemoryStream
    try {$Stream.Position=0;$Stream.CopyTo($buffer);return ,$buffer.ToArray()} finally {$buffer.Dispose()}
}
function Write-LockedBytes($Stream,[byte[]]$Bytes) {
    $Stream.Position=0;$Stream.SetLength(0);$Stream.Write($Bytes,0,$Bytes.Length);$Stream.Flush($true)
}
$data=Json 'UserDataAccepted.json';$help=Json 'HelpFinalRelease/HelpEvidence.json'
$standalone=Json 'StandaloneHelpAccepted.json';$source=Json 'SourceEqualityAccepted.json'
if ($data.status -ne 'REVIEW_XML_PRESERVATION_PASS' -or $help.Status -ne 'DIRECTED_REVIEW_HELP_PASS' -or
    $standalone.status -ne 'STANDALONE_HELP_PASS' -or -not $source.AllPassed -or
    $source.ActualComponents -ne 149 -or $source.RepositoryComponents -ne 144) {throw 'Review publication gates are incomplete.'}
$book=Join-Path $release 'Candidate.xlsm';$dump=Join-Path $release 'VBA_Ready.txt'
$bookHash=Hash $book;$dumpHash=Hash $dump
if ($bookHash -ne $data.afterSHA256 -or $bookHash -ne $standalone.workbookSHA256 -or $dumpHash -ne $standalone.dumpSHA256) {throw 'Reviewed artifacts changed.'}
$files=@(Get-ChildItem -LiteralPath (Join-Path $root 'src'),(Join-Path $root 'tests') -Recurse -File | Where-Object Extension -in '.bas','.cls')
foreach ($check in $source.Checks) {
    $file=@($files | Where-Object BaseName -ceq $check.Component)
    if ($file.Count -ne 1 -or (Hash $file[0].FullName) -ne $check.SourceSHA256.ToLowerInvariant()) {throw ('Source changed: '+$check.Component)}
}
if ((Hash (Join-Path $PSScriptRoot 'Build-Workbook.ps1')) -ne $source.BuildScriptSHA256.ToLowerInvariant()) {throw 'Fresh build source changed.'}
$output=Join-Path $root 'workbook/output';$targetBook=Join-Path $output 'RC_Section_NDM.xlsm';$targetDump=Join-Path $output 'VBA_All_Code.txt'
$foreign=@(Get-ChildItem -LiteralPath $output -Recurse -File | Where-Object FullName -notin @($targetBook,$targetDump) |
    ForEach-Object {[ordered]@{Path=$_.FullName.Substring($root.Length+1);SHA256=Hash $_.FullName}})
$backup=Join-Path $release 'OriginalOutputBackup'
if (Test-Path -LiteralPath $backup) {throw 'Review backup already exists; inspect prior publication.'}
$bookStream=$null;$dumpStream=$null;$written=$false
try {
    # Блокировка действует до окончания записи; пользовательский Excel не закрывается.
    $bookStream=[IO.File]::Open($targetBook,[IO.FileMode]::Open,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None)
    $dumpStream=[IO.File]::Open($targetDump,[IO.FileMode]::Open,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None)
    [byte[]]$originalBook=Read-LockedBytes $bookStream;[byte[]]$originalDump=Read-LockedBytes $dumpStream
    if ((BytesHash $originalBook) -ne $data.beforeSHA256 -or (BytesHash $originalDump) -ne (Hash (Join-Path $release 'Baseline/VBA_All_Code.txt'))) {throw 'Output changed since baseline; preserve new user changes.'}
    New-Item -ItemType Directory -Path $backup | Out-Null
    [IO.File]::WriteAllBytes((Join-Path $backup 'RC_Section_NDM.xlsm'),$originalBook)
    [IO.File]::WriteAllBytes((Join-Path $backup 'VBA_All_Code.txt'),$originalDump)
    $written=$true
    Write-LockedBytes $bookStream ([IO.File]::ReadAllBytes($book))
    Write-LockedBytes $dumpStream ([IO.File]::ReadAllBytes($dump))
    if ((BytesHash (Read-LockedBytes $bookStream)) -ne $bookHash -or (BytesHash (Read-LockedBytes $dumpStream)) -ne $dumpHash) {throw 'Published bytes differ.'}
    foreach ($item in $foreign) {if ((Hash (Join-Path $root $item.Path)) -ne $item.SHA256) {throw ('Unrelated output changed: '+$item.Path)}}
} catch {
    if ($written) {Write-LockedBytes $bookStream $originalBook;Write-LockedBytes $dumpStream $originalDump}
    throw
} finally {
    if ($dumpStream) {$dumpStream.Dispose()}
    if ($bookStream) {$bookStream.Dispose()}
}
[ordered]@{Published=$true;DateUTC=(Get-Date).ToUniversalTime().ToString('o');WorkbookSHA256=$bookHash;ExportSHA256=$dumpHash;
    OriginalWorkbookSHA256=$data.beforeSHA256;OriginalDumpSHA256=BytesHash $originalDump;
    UserDataPreserved=$true;ForeignOutputsPreserved=$true;ForeignOutputFiles=$foreign;Backup=$backup;
    SourceCodeCommit=(& git -C $root rev-parse HEAD);ActualVBAComponents=149;PushPerformed=$false} |
    ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $release 'Publication.json') -Encoding UTF8
Write-Output ('GEOMETRY_REVIEW_PUBLISHED: '+$bookHash+'; VBA='+$dumpHash+'; foreign files='+$foreign.Count)
