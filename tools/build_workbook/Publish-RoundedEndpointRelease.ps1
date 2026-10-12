# Публикует F02 только после всех согласованных проверок, сохраняя исходную пару и другие output-файлы.
param([string]$Directory='docs/regression/RoundedEndpoint_2026-10-12',[switch]$ForCodeReview)
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$release=[IO.Path]::GetFullPath((Join-Path $root $Directory))
$expected=[IO.Path]::GetFullPath((Join-Path $root 'docs/regression/RoundedEndpoint_2026-10-12'))
if ($release -cne $expected) {throw 'Only the rounded endpoint evidence directory is allowed.'}
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
if ($ForCodeReview) {
    # Пользователь остановил полные наборы и запросил проверяемую книгу/ALL CODE.
    # Это публикация для внешнего аудита, а не итоговая полная приёмка.
    $data=Json 'UserDataReview30.json';$help=Json 'HelpCodeReview30/HelpEvidence.json'
    $standalone=Json 'StandaloneHelpReview30.json';$source=Json 'SourceEqualityReview30.json';$testedSource=Json 'SourceEqualityReview30Ready.json'
    $modes=Json 'DirectedCodeReviewEvidence.json';$ordinary=Json 'OrdinaryReviewEquality.json'
    $structureFile='StructureReview30.json';$bookFile='Review30Release.xlsm';$dumpFile='VBA_Review30.txt';$testedFile='Review30Ready.xlsm'
} else {
    $data=Json 'UserDataRelease.json';$help=Json 'HelpAccepted/HelpEvidence.json'
    $standalone=Json 'StandaloneHelpRelease.json';$source=Json 'SourceEqualityRelease.json';$testedSource=Json 'SourceEqualityFinal.json'
    $modes=Json 'FullOnOffEvidence.json';$ordinary=Json 'OrdinaryEquality.json'
    $structureFile='StructureRelease.json';$bookFile='Release.xlsm';$dumpFile='VBA_Release.txt';$testedFile='Ready.xlsm'
}
if ($data.status -ne 'REVIEW_XML_PRESERVATION_PASS' -or $help.Status -ne 'DIRECTED_REVIEW_HELP_PASS' -or
    $standalone.status -ne 'STANDALONE_HELP_PASS' -or -not $source.AllPassed -or -not $testedSource.AllPassed -or
    $source.ActualComponents -ne 149 -or $source.RepositoryComponents -ne 144 -or
    $modes.status -ne 'PASS' -or $ordinary.status -ne 'PASS') {throw 'Rounded endpoint publication gates are incomplete.'}
if ($ForCodeReview) {
    if ($modes.fullAcceptance -ne 'DEFERRED_BY_USER' -or $modes.reports.Count -ne 5) {throw 'Directed code review gates are incomplete.'}
} elseif ($modes.reports.Count -ne 4) {throw 'Full acceptance gates are incomplete.'}
$structure=@(Get-Content -LiteralPath (Join-Path $release $structureFile) -Raw -Encoding UTF8 | ConvertFrom-Json)
if (-not $structure.Count -or @($structure | Where-Object {-not $_.Passed}).Count) {throw 'Workbook structure has failed checks.'}
$book=Join-Path $release $bookFile;$dump=Join-Path $release $dumpFile
$bookHash=Hash $book;$dumpHash=Hash $dump;$testedHash=Hash (Join-Path $release $testedFile)
if ($bookHash -ne $data.afterSHA256 -or $bookHash -ne $standalone.workbookSHA256 -or $dumpHash -ne $standalone.dumpSHA256) {throw 'Reviewed artifacts changed.'}
foreach ($mode in $modes.reports) {
    if ($mode.status -ne 'PASS' -or $mode.sourceSHA256 -ne $testedHash -or $mode.reportSHA256 -ne (Hash (Join-Path $release $mode.name))) {throw 'Full on/off evidence changed.'}
}
$files=@(Get-ChildItem -LiteralPath (Join-Path $root 'src'),(Join-Path $root 'tests') -Recurse -File | Where-Object Extension -in '.bas','.cls')
foreach ($catalog in @($source,$testedSource)) {
    foreach ($check in $catalog.Checks) {
        $file=@($files | Where-Object BaseName -ceq $check.Component)
        if ($file.Count -ne 1 -or (Hash $file[0].FullName) -ne $check.SourceSHA256.ToLowerInvariant()) {throw ('Source changed: '+$check.Component)}
    }
    if ((Hash (Join-Path $PSScriptRoot 'Build-Workbook.ps1')) -ne $catalog.BuildScriptSHA256.ToLowerInvariant()) {throw 'Fresh build source changed.'}
}
$output=Join-Path $root 'workbook/output';$targetBook=Join-Path $output 'RC_Section_NDM.xlsm';$targetDump=Join-Path $output 'VBA_All_Code.txt'
$foreign=@(Get-ChildItem -LiteralPath $output -Recurse -File | Where-Object FullName -notin @($targetBook,$targetDump) |
    ForEach-Object {[ordered]@{Path=$_.FullName.Substring($root.Length+1);SHA256=Hash $_.FullName}})
$backup=Join-Path $release 'OriginalOutputBackup'
if (Test-Path -LiteralPath $backup) {throw 'Backup already exists; inspect prior publication.'}
$bookStream=$null;$dumpStream=$null;$written=$false
try {
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
    SourceCodeCommit=(& git -C $root rev-parse HEAD);ActualVBAComponents=149;
    FullOnOff=$(if ($ForCodeReview) {'DEFERRED_BY_USER'} else {'PASS'});
    PublicationPurpose=$(if ($ForCodeReview) {'EXTERNAL_CODE_REVIEW'} else {'FULL_ACCEPTANCE'});PushPerformed=$false} |
    ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $release 'Publication.json') -Encoding UTF8
Write-Output ('ROUNDED_ENDPOINT_PUBLISHED: '+$bookHash+'; VBA='+$dumpHash+'; foreign files='+$foreign.Count)
