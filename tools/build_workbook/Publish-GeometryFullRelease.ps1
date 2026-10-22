# Публикует только полностью проверенную пару книга/VBA, сохраняя пользовательские данные.
param([string]$ReleaseDirectory='docs/regression/GeometryFull_2026-10-13')
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$release=[IO.Path]::GetFullPath((Join-Path $root $ReleaseDirectory))
$expected=[IO.Path]::GetFullPath((Join-Path $root 'docs/regression/GeometryFull_2026-10-13'))
if ($release -cne $expected) {throw 'Unexpected acceptance directory.'}
function Json([string]$Name) {Get-Content -LiteralPath (Join-Path $release $Name) -Raw -Encoding UTF8 | ConvertFrom-Json}
function Hash([string]$Path) {(Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant()}
function BytesHash([byte[]]$Bytes) {
 $digest=[Security.Cryptography.SHA256]::Create()
 try {[BitConverter]::ToString($digest.ComputeHash($Bytes)).Replace('-','').ToLowerInvariant()} finally {$digest.Dispose()}
}
function ReadBytes($Stream) {
 $buffer=New-Object IO.MemoryStream
 try {$Stream.Position=0;$Stream.CopyTo($buffer);return ,$buffer.ToArray()} finally {$buffer.Dispose()}
}
function WriteBytes($Stream,[byte[]]$Bytes) {
 $Stream.Position=0;$Stream.SetLength(0);$Stream.Write($Bytes,0,$Bytes.Length);$Stream.Flush($true)
}
$acceptance=Json 'FINAL_ACCEPTANCE_EVIDENCE.json'
$full=Json 'FinalFullEvidence.json'
$candidate=Json 'FinalReleaseCandidate.json'
$data=Json $candidate.data
$source=Json $candidate.source
$standalone=Json $candidate.standalone
$bridge=Json $candidate.nativeBridge
$lambdaBridge=Json $candidate.lambdaBridge
$profileBridge=Json $candidate.profileBridge
$help=Json 'FinalHelp/HelpEvidence.json'
# Windows PowerShell 5 возвращает JSON-массив одним объектом в pipeline.
# Прямое присваивание сохраняет его 27 записей без дополнительной оболочки.
$structure=Json $candidate.structure
foreach ($check in $structure) {
 if ($check.Passed -isnot [bool]) {throw 'Structure proof must contain explicit Boolean results.'}
}
if ($candidate.book -ne 'ReleaseSnapshotFixed.xlsm' -or $candidate.dump -ne 'VBA_SnapshotFixed.txt') {throw 'Unexpected release candidate.'}
$book=Join-Path $release $candidate.book;$dump=Join-Path $release $candidate.dump
$bookHash=Hash $book;$dumpHash=Hash $dump
if ($acceptance.status -ne 'PASS' -or $full.status -ne 'PASS' -or $full.reports.Count -ne 6 -or
    $data.status -ne 'REVIEW_XML_PRESERVATION_PASS' -or $data.userCellsChecked -ne 52003 -or
    -not $source.AllPassed -or $source.ActualComponents -ne 149 -or $source.RepositoryComponents -ne 144 -or
    $standalone.status -ne 'STANDALONE_HELP_PASS' -or $standalone.apiMembers -ne 275 -or
    $help.Status -ne 'DIRECTED_REVIEW_HELP_PASS' -or -not $help.Idempotent -or -not $help.ReopenText -or
    $bridge.status -ne 'PASS' -or -not $bridge.productionUnchanged -or -not $bridge.unchangedOtherMethods -or
    ($bridge.changed -join ',') -ne 'modTestContourMesh' -or $bridge.changedMethod -ne 'modTestContourMesh.TestCheckNativeFullExport' -or
    $lambdaBridge.status -ne 'PASS' -or -not $lambdaBridge.productionUnchanged -or -not $lambdaBridge.unchangedOtherCode -or
    ($lambdaBridge.changed -join ',') -ne 'modTestConfiguration' -or $lambdaBridge.changedMethod -ne 'modTestConfiguration.TestCheckCapacityNumericEffects' -or
    $profileBridge.status -ne 'PASS' -or -not $profileBridge.productionUnchanged -or -not $profileBridge.unchangedOtherCode -or -not $profileBridge.tolerancesUnchanged -or
    ($profileBridge.changed -join ',') -ne 'modTestProfileConfig' -or $profileBridge.changedMethod -ne 'modTestProfileConfig.TestCheckVisualizationValues' -or
    $structure.Count -ne 27 -or @($structure | Where-Object {-not $_.Passed}).Count) {throw 'Full publication gates are incomplete.'}
if ($acceptance.sourceSHA256 -ne $bookHash -or $full.sourceSHA256 -ne $bridge.testedWorkbookSHA256 -or
    $bridge.testedWorkbookSHA256 -ne (Hash (Join-Path $release 'FinalReady.xlsm')) -or
    $bridge.releaseWorkbookSHA256 -ne $lambdaBridge.beforeWorkbookSHA256 -or
    $bridge.releaseWorkbookSHA256 -ne (Hash (Join-Path $release $candidate.intermediateBook)) -or
    $bridge.afterDumpSHA256 -ne $lambdaBridge.beforeDumpSHA256 -or
    $bridge.afterDumpSHA256 -ne (Hash (Join-Path $release $candidate.intermediateDump)) -or
    $lambdaBridge.afterWorkbookSHA256 -ne $profileBridge.beforeWorkbookSHA256 -or
    $lambdaBridge.afterDumpSHA256 -ne $profileBridge.beforeDumpSHA256 -or
    $lambdaBridge.afterWorkbookSHA256 -ne (Hash (Join-Path $release 'ReleaseLambdaFixed.xlsm')) -or
    $lambdaBridge.afterDumpSHA256 -ne (Hash (Join-Path $release 'VBA_LambdaFixed.txt')) -or
    $profileBridge.afterWorkbookSHA256 -ne $bookHash -or $profileBridge.afterDumpSHA256 -ne $dumpHash -or
    $data.afterSHA256 -ne $bookHash -or $standalone.workbookSHA256 -ne $bookHash -or
    $standalone.dumpSHA256 -ne $dumpHash) {throw 'Verified publication artifacts changed.'}
$required=@('Full8','Geometry14','ValidationExtras9','ValidationPlan7','Domain21','Family9_SP35_SP63','NativeCAD5','OneClass','WorkflowPerformance','PlotModes')
foreach ($name in $required) {
 $group=@($acceptance.groups | Where-Object name -eq $name)
 if ($group.Count -ne 1 -or $group[0].status -ne 'PASS' -or ($group[0].modes -join ',') -ne 'Off,On') {throw ('Acceptance group incomplete: '+$name)}
}
foreach ($record in @($acceptance.artifacts)+@($full.reports)) {
 $relative=if ($record.path) {$record.path} else {$record.name}
 $path=[IO.Path]::GetFullPath((Join-Path $release $relative))
 if (-not $path.StartsWith($release+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)) {throw 'Artifact escapes the acceptance directory.'}
 $recordHash=if ($record.sha256) {$record.sha256} else {$record.reportSHA256}
 if ((Hash $path) -ne $recordHash) {throw ('Acceptance evidence changed: '+$relative)}
}
$files=@(Get-ChildItem -LiteralPath (Join-Path $root 'src'),(Join-Path $root 'tests') -Recurse -File | Where-Object Extension -in '.bas','.cls')
foreach ($check in $source.Checks) {
 $file=@($files | Where-Object BaseName -ceq $check.Component)
 if ($file.Count -ne 1 -or (Hash $file[0].FullName) -ne $check.SourceSHA256.ToLowerInvariant()) {throw ('Source changed: '+$check.Component)}
}
if ((Hash (Join-Path $PSScriptRoot 'Build-Workbook.ps1')) -ne $source.BuildScriptSHA256.ToLowerInvariant()) {throw 'Build source changed.'}
$output=Join-Path $root 'workbook/output';$targetBook=Join-Path $output 'RC_Section_NDM.xlsm';$targetDump=Join-Path $output 'VBA_All_Code.txt'
$foreign=@(Get-ChildItem -LiteralPath $output -Recurse -File | Where-Object FullName -notin @($targetBook,$targetDump) |
 ForEach-Object {[ordered]@{Path=$_.FullName.Substring($root.Length+1);SHA256=Hash $_.FullName}})
$backup=Join-Path $release 'OriginalOutputBackup'
if (Test-Path -LiteralPath $backup) {throw 'Publication backup already exists; inspect prior outcome.'}
$bookStream=$null;$dumpStream=$null;$written=$false
try {
 $bookStream=[IO.File]::Open($targetBook,[IO.FileMode]::Open,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None)
 $dumpStream=[IO.File]::Open($targetDump,[IO.FileMode]::Open,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None)
 [byte[]]$originalBook=ReadBytes $bookStream;[byte[]]$originalDump=ReadBytes $dumpStream
 if ((BytesHash $originalBook) -ne $data.beforeSHA256 -or (BytesHash $originalDump) -ne (Hash (Join-Path $release 'Baseline/VBA_All_Code.txt'))) {throw 'Output changed since baseline; preserve new user data first.'}
 New-Item -ItemType Directory -Path $backup | Out-Null
 [IO.File]::WriteAllBytes((Join-Path $backup 'RC_Section_NDM.xlsm'),$originalBook)
 [IO.File]::WriteAllBytes((Join-Path $backup 'VBA_All_Code.txt'),$originalDump)
 $written=$true
 WriteBytes $bookStream ([IO.File]::ReadAllBytes($book));WriteBytes $dumpStream ([IO.File]::ReadAllBytes($dump))
 if ((BytesHash (ReadBytes $bookStream)) -ne $bookHash -or (BytesHash (ReadBytes $dumpStream)) -ne $dumpHash) {throw 'Published bytes differ.'}
 foreach ($item in $foreign) {if ((Hash (Join-Path $root $item.Path)) -ne $item.SHA256) {throw ('Unrelated output changed: '+$item.Path)}}
} catch {
 if ($written) {WriteBytes $bookStream $originalBook;WriteBytes $dumpStream $originalDump}
 throw
} finally {
 if ($dumpStream) {$dumpStream.Dispose()};if ($bookStream) {$bookStream.Dispose()}
}
[ordered]@{Published=$true;DateUTC=(Get-Date).ToUniversalTime().ToString('o');WorkbookSHA256=$bookHash;ExportSHA256=$dumpHash;
 OriginalWorkbookSHA256=$data.beforeSHA256;OriginalDumpSHA256=BytesHash $originalDump;
 UserDataPreserved=$true;ForeignOutputsPreserved=$true;ForeignOutputFiles=$foreign;Backup=$backup;
 SourceCodeCommit=(& git -C $root rev-parse HEAD);ActualVBAComponents=149;FullOnOff='PASS';PushPerformed=$false} |
 ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $release 'Publication.json') -Encoding UTF8
Write-Output ('FULL_GEOMETRY_PUBLISHED: '+$bookHash+'; VBA='+$dumpHash+'; foreign files='+$foreign.Count)
