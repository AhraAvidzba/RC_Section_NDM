# Публикует только проверенную книгу/dump. Исходные пользовательские данные,
# execution report, formula_images и другие output-файлы защищены hash guard.
param([string]$Directory='docs/regression/GeometryDeclarative_2026-10-12/FinalRelease')
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$release=[IO.Path]::GetFullPath((Join-Path $root $Directory))
$allowed=[IO.Path]::GetFullPath((Join-Path $root 'docs/regression/GeometryDeclarative_2026-10-12'))+[IO.Path]::DirectorySeparatorChar
if (-not $release.StartsWith($allowed,[StringComparison]::OrdinalIgnoreCase)) {throw 'Release path is outside geometry evidence.'}
function Read-Json([string]$Path) {Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json}
function Hash([string]$Path) {(Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant()}
$evidence=Read-Json (Join-Path $root 'docs/geometry_refactor/FINAL_ACCEPTANCE_EVIDENCE.json')
$preserved=Read-Json (Join-Path $release 'UserDataPreservation.json')
$guide=Read-Json (Join-Path $release 'GuideFinalPreservation.json')
$source=Read-Json (Join-Path $release 'SourceEquality.json')
if ($evidence.status -ne 'READY_TO_PUBLISH' -or $preserved.status -ne 'USER_DATA_PRESERVED' -or
    $guide.status -ne 'XML_PRESERVATION_PASS' -or -not $source.AllPassed -or $source.ActualComponents -ne 147 -or
    $source.RepositoryComponents -ne 142 -or $evidence.nativeCAD.status -ne 'PASS') {throw 'Release gates are incomplete.'}
$sourceFiles=@(Get-ChildItem -LiteralPath (Join-Path $root 'src'),(Join-Path $root 'tests') -Recurse -File | Where-Object Extension -in '.bas','.cls')
foreach($check in $source.Checks) {
    $files=@($sourceFiles | Where-Object BaseName -eq $check.Component)
    if ($files.Count -ne 1 -or (Hash $files[0].FullName) -ne $check.SourceSHA256.ToLowerInvariant()) {throw ('Source changed: '+$check.Component)}
}
if ((Hash (Join-Path $root 'tools/build_workbook/Build-Workbook.ps1')) -ne $source.BuildScriptSHA256.ToLowerInvariant()) {throw 'Fresh build source changed.'}
$book=Join-Path $release 'RC_Section_NDM.xlsm';$dump=Join-Path $release 'VBA_All_Code.txt'
$output=Join-Path $root 'workbook/output'
$targetBook=Join-Path $output 'RC_Section_NDM.xlsm';$targetDump=Join-Path $output 'VBA_All_Code.txt'
if ((Hash $book) -ne $evidence.workbookSHA256 -or (Hash $dump) -ne $evidence.vbaSHA256 -or
    (Hash $targetBook) -ne $preserved.beforeSHA256) {throw 'Release or original workbook changed; do not overwrite user changes.'}
$foreign=@(Get-ChildItem -LiteralPath $output -Recurse -File | Where-Object FullName -notin @($targetBook,$targetDump) |
    ForEach-Object {[ordered]@{Path=$_.FullName.Substring($root.Length+1);SHA256=Hash $_.FullName}})
# Отказ при открытой исходной книге не закрывает пользовательский Excel.
foreach($path in @($targetBook,$targetDump)) {
    $probe=[IO.File]::Open($path,[IO.FileMode]::Open,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None)
    $probe.Dispose()
}
$backup=Join-Path $release 'OriginalOutputBackup'
if (Test-Path -LiteralPath $backup) {throw 'Backup already exists: review prior publication instead of overwriting it.'}
New-Item -ItemType Directory -Path $backup | Out-Null
Copy-Item -LiteralPath $targetBook -Destination (Join-Path $backup 'RC_Section_NDM.xlsm')
Copy-Item -LiteralPath $targetDump -Destination (Join-Path $backup 'VBA_All_Code.txt')
$originalDumpHash=Hash $targetDump
try {
    Copy-Item -LiteralPath $book -Destination $targetBook
    Copy-Item -LiteralPath $dump -Destination $targetDump
    if ((Hash $targetBook) -ne $evidence.workbookSHA256 -or (Hash $targetDump) -ne $evidence.vbaSHA256) {throw 'Published pair differs.'}
    foreach($item in $foreign) {if ((Hash (Join-Path $root $item.Path)) -ne $item.SHA256) {throw ('Unrelated output changed: '+$item.Path)}}
} catch {
    # Восстанавливаем только свои опубликованные байты, не чужую позднюю правку.
    if ((Hash $targetBook) -eq $evidence.workbookSHA256) {Copy-Item -LiteralPath (Join-Path $backup 'RC_Section_NDM.xlsm') -Destination $targetBook}
    if ((Hash $targetDump) -eq $evidence.vbaSHA256) {Copy-Item -LiteralPath (Join-Path $backup 'VBA_All_Code.txt') -Destination $targetDump}
    throw
}
[ordered]@{Published=$true;DateUTC=(Get-Date).ToUniversalTime().ToString('o');WorkbookSHA256=Hash $targetBook;
    ExportSHA256=Hash $targetDump;OriginalWorkbookSHA256=$preserved.beforeSHA256;OriginalDumpSHA256=$originalDumpHash;
    UserDataPreserved=$true;ForeignOutputsPreserved=$true;ForeignOutputFiles=$foreign;Backup=$backup;
    SourceCodeCommit=(& git -C $root rev-parse HEAD);ActualVBAComponents=147;PushPerformed=$false} |
    ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $release 'Publication.json') -Encoding UTF8
Write-Output ('GEOMETRY_RELEASE_PUBLISHED: '+(Hash $targetBook)+'; VBA='+(Hash $targetDump)+'; foreign files preserved='+$foreign.Count)
