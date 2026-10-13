# Публикует проверенный числовой наклон плато и native ALL CODE одной парой.
# Ввод и Results защищены независимой XML-проверкой; другие output-файлы не меняются.
param([string]$Directory='docs/regression/GeometryFull_2026-10-13/PlateauFactorV2')
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$directoryPath=[IO.Path]::GetFullPath((Join-Path $root $Directory))
$allowed=[IO.Path]::GetFullPath((Join-Path $root 'docs/regression/GeometryFull_2026-10-13'))+[IO.Path]::DirectorySeparatorChar
if (-not $directoryPath.StartsWith($allowed,[StringComparison]::OrdinalIgnoreCase)) {throw 'Unexpected publication directory.'}
function Json([string]$Name) {Get-Content -LiteralPath (Join-Path $directoryPath $Name) -Raw -Encoding UTF8 | ConvertFrom-Json}
function Hash([string]$Path) {(Get-FileHash -LiteralPath $Path).Hash.ToLowerInvariant()}
function BytesHash([byte[]]$Bytes) {
    $sha=[Security.Cryptography.SHA256]::Create()
    try {[BitConverter]::ToString($sha.ComputeHash($Bytes)).Replace('-','').ToLowerInvariant()} finally {$sha.Dispose()}
}
$candidate=Join-Path $directoryPath 'RC_Section_NDM.xlsm'
$dump=Join-Path $directoryPath 'VBA_All_Code.txt'
$candidateHash=Hash $candidate;$dumpHash=Hash $dump
$data=Json 'UserDataV2.json';$native=Json 'NativeSourceEquality.json';$help=Json 'StandaloneHelp.json'
$comparison=Json 'StatusComparison.json';$structure=Json 'Structure.json';$prepare=Json 'Preparation.json'
if (-not $data.AllPassed -or $data.UserCellsChecked -ne 107461 -or
    -not $native.AllPassed -or $native.ActualComponents -ne 149 -or $native.DocumentComponents -ne 5 -or
    $help.status -ne 'STANDALONE_HELP_PASS' -or $help.apiMembers -ne 275 -or
    -not $comparison.AllPassed -or $comparison.ComparedPairs -ne 24 -or $comparison.NewNumericalFailures -ne 0 -or
    $structure.Count -ne 27 -or @($structure | Where-Object {-not $_.Passed}).Count) {throw 'Directed publication gates are incomplete.'}
if ($data.WorkbookSHA256 -ne $candidateHash -or $native.WorkbookSHA256 -ne $candidateHash -or
    $help.workbookSHA256 -ne $candidateHash -or $native.DumpSHA256 -ne $dumpHash -or $help.dumpSHA256 -ne $dumpHash) {throw 'A verified artifact changed.'}
foreach ($check in $native.Checks) {
    $files=@(Get-ChildItem -LiteralPath (Join-Path $root 'src'),(Join-Path $root 'tests') -Recurse -File | Where-Object BaseName -ceq $check.Component)
    if ($files.Count -ne 1 -or (Hash $files[0].FullName) -ne $check.SourceSHA256.ToLowerInvariant()) {throw ('Source changed after tests: '+$check.Component)}
}
foreach ($gate in @(@('MaterialFactorV2.txt','TOTAL_MATERIAL: passed=25867; failed=0'),
        @('Results_Config.txt','TOTAL_PLATEAU_RESULTS: passed=186; failed=0'),
        @('Results_Config.txt','TOTAL_CONFIG_PRESENTATION: passed=1589; failed=0'))) {
    $text=Get-Content -LiteralPath (Join-Path $directoryPath $gate[0]) -Raw -Encoding UTF8
    if (-not $text.Contains($gate[1]) -or $text -match '(?m)^FAIL:|RUNTIME ERROR|SCRIPT ERROR|RESTORE ERROR' -or
        $text -notmatch ('SOURCE_SHA256: '+$candidateHash) -or
        -not $text.Contains('WATCHDOG_COMPLETED: exit=0; source unchanged=True')) {throw ('Directed test failed or is stale: '+$gate[0])}
}
foreach ($record in $comparison.ReportSHA256.PSObject.Properties) {
    if ((Hash (Join-Path $directoryPath $record.Name)) -ne $record.Value) {throw ('Comparison report changed: '+$record.Name)}
}
foreach ($probeName in @('BaselineProbe','ReleaseProbe')) {
    $probeDirectory=Join-Path $directoryPath $probeName
    $probe=Get-Content -LiteralPath (Join-Path $probeDirectory 'Preparation.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    $expectedSource=if ($probeName -eq 'BaselineProbe') {$data.BeforeSHA256} else {$candidateHash}
    if (-not $probe.Compiled -or $probe.SourceSHA256 -ne $expectedSource -or
        $probe.ProbeWorkbookSHA256 -ne (Hash (Join-Path $probeDirectory 'RC_Section_NDM.xlsm')) -or
        $probe.ExtraHelperSHA256 -ne (Hash (Join-Path $PSScriptRoot 'PlateauFactorControlProbe.bas')) -or
        $probe.HelperSHA256 -ne (Hash (Join-Path $PSScriptRoot 'FocusedAcceptanceProbe.bas'))) {throw 'Control probe is stale.'}
}
$production=Join-Path $root 'workbook/output/RC_Section_NDM.xlsm'
$publishedDump=Join-Path $root 'workbook/output/VBA_All_Code.txt'
$expectedBefore=$prepare.SourceSHA256.ToLowerInvariant()
if ((Hash $production) -ne $expectedBefore -or (Hash $publishedDump) -ne (Hash (Join-Path $directoryPath 'BeforeVBA.txt'))) {throw 'User workbook or dump changed before publication.'}
$other=@{}
foreach ($file in (Get-ChildItem -LiteralPath (Join-Path $root 'workbook/output') -File -Recurse)) {
    if ($file.FullName -in @($production,$publishedDump)) {continue}
    $other[$file.FullName]=(Hash $file.FullName)
}
$bookBytes=[IO.File]::ReadAllBytes($candidate);$dumpBytes=[IO.File]::ReadAllBytes($dump)
$bookStream=$null;$dumpStream=$null;$published=$false
try {
    # Исключительная блокировка не позволяет заменить книгу поверх сохранения Excel.
    $bookStream=[IO.File]::Open($production,[IO.FileMode]::Open,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None)
    $dumpStream=[IO.File]::Open($publishedDump,[IO.FileMode]::Open,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None)
    $oldBook=New-Object byte[] ([int]$bookStream.Length);$oldDump=New-Object byte[] ([int]$dumpStream.Length)
    $bookOffset=0;while ($bookOffset -lt $oldBook.Length) {$count=$bookStream.Read($oldBook,$bookOffset,$oldBook.Length-$bookOffset);if (-not $count) {throw 'Incomplete workbook read.'};$bookOffset+=$count}
    $dumpOffset=0;while ($dumpOffset -lt $oldDump.Length) {$count=$dumpStream.Read($oldDump,$dumpOffset,$oldDump.Length-$dumpOffset);if (-not $count) {throw 'Incomplete dump read.'};$dumpOffset+=$count}
    if ((BytesHash $oldBook) -ne $expectedBefore -or (BytesHash $oldDump) -ne (Hash (Join-Path $directoryPath 'BeforeVBA.txt'))) {throw 'Output changed before the exclusive lock.'}
    [IO.File]::WriteAllBytes((Join-Path $directoryPath 'PublishedBefore.xlsm'),$oldBook)
    [IO.File]::WriteAllBytes((Join-Path $directoryPath 'PublishedBeforeVBA.txt'),$oldDump)
    try {
        $bookStream.Position=0;$bookStream.Write($bookBytes,0,$bookBytes.Length);$bookStream.SetLength($bookBytes.Length);$bookStream.Flush($true)
        $dumpStream.Position=0;$dumpStream.Write($dumpBytes,0,$dumpBytes.Length);$dumpStream.SetLength($dumpBytes.Length);$dumpStream.Flush($true)
        $published=$true
    } catch {
        $bookStream.Position=0;$bookStream.Write($oldBook,0,$oldBook.Length);$bookStream.SetLength($oldBook.Length);$bookStream.Flush($true)
        $dumpStream.Position=0;$dumpStream.Write($oldDump,0,$oldDump.Length);$dumpStream.SetLength($oldDump.Length);$dumpStream.Flush($true)
        throw
    }
} finally {
    if ($dumpStream) {$dumpStream.Dispose()}
    if ($bookStream) {$bookStream.Dispose()}
}
if (-not $published -or (Hash $production) -ne $candidateHash -or (Hash $publishedDump) -ne $dumpHash) {throw 'Published pair does not match verified bytes.'}
$following=@(Get-ChildItem -LiteralPath (Join-Path $root 'workbook/output') -File -Recurse | Where-Object FullName -notin @($production,$publishedDump))
if ($following.Count -ne $other.Count) {throw 'Other output inventory changed.'}
foreach ($file in $following) {if ($other[$file.FullName] -ne (Hash $file.FullName)) {throw ('Other output changed: '+$file.FullName)}}
[ordered]@{AllPassed=$true;Published=$true;BeforeWorkbookSHA256=$expectedBefore;WorkbookSHA256=$candidateHash;DumpSHA256=$dumpHash;
    DefaultFactor=0.001;Slope='0.001*E';ComparedPairs=24;NewNumericalFailures=0;OtherOutputFilesPreserved=$other.Count;
    Compiled=$true;PushPerformed=$false;Backup='PublishedBefore.xlsm + PublishedBeforeVBA.txt';
    VerificationSHA256=[ordered]@{UserData=(Hash (Join-Path $directoryPath 'UserDataV2.json'));Code=(Hash (Join-Path $directoryPath 'NativeSourceEquality.json'));Statuses=(Hash (Join-Path $directoryPath 'StatusComparison.json'))}} |
    ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directoryPath 'Publication.json') -Encoding UTF8
Write-Output ('PUBLISHED_FACTOR_PASS: workbook and ALL CODE match tested pair; other output preserved='+$other.Count)
