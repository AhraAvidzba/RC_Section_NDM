# Publishes only a fully verified native candidate; existing reports/images stay untouched.
param([string]$Directory='docs/regression/PlateauSlope_2026-10-11/ReleaseActualDiagrams',
      [string]$ExpectedDumpSHA256='D8E0064A46140C865D9E6BBB53449BB5F1E01530AE9AA6290BF15647DF48DF4F')
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$evidence=[IO.Path]::GetFullPath((Join-Path $root $Directory))
$allowed=[IO.Path]::GetFullPath((Join-Path $root 'docs/regression/PlateauSlope_2026-10-11'))+[IO.Path]::DirectorySeparatorChar
$openIssues=[IO.Path]::GetFullPath((Join-Path $root 'docs/regression/OpenIssues_2026-10-11'))+[IO.Path]::DirectorySeparatorChar
if (-not $evidence.StartsWith($allowed,[StringComparison]::OrdinalIgnoreCase) -and
    -not $evidence.StartsWith($openIssues,[StringComparison]::OrdinalIgnoreCase)) {throw 'Unexpected evidence path.'}
if ($evidence.StartsWith($openIssues,[StringComparison]::OrdinalIgnoreCase)) {
    $stage=Get-Content -LiteralPath (Join-Path $openIssues 'FocusedV2/LoadStage.json') -Raw | ConvertFrom-Json
    $review=Get-Content -LiteralPath (Join-Path $openIssues 'ForceReview.json') -Raw | ConvertFrom-Json
    $native=Get-Content -LiteralPath (Join-Path $openIssues 'NativeCAD/Stage.json') -Raw | ConvertFrom-Json
    if (-not $stage.LoadStageComplete -or -not $stage.SourceUnchanged -or -not $stage.ProductionUnchanged -or
        @($stage.Groups | Where-Object {-not $_.Completed}).Count -or -not $native.Passed -or
        $review.SourceSHA256 -cne $stage.SourceSHA256 -or $review.Inconclusive -ne 0 -or
        $review.ReviewedStates -ne $review.ProvenImpossible -or
        -not $review.AllNumericalMessagesCovered -or
        @($stage.NumericalFailuresForReview).Count -ne $review.ReviewedMessages) {throw 'Directed numerical review or native CAD gates failed.'}
    foreach ($name in @('OpenIssuesFinal_No.txt','OpenIssuesFinal_Yes.txt')) {
        $body=Get-Content -LiteralPath (Join-Path $openIssues $name) -Raw -Encoding UTF8
        if ($body -notmatch 'TOTAL_OPEN_ISSUES: passed=6; failed=0' -or $body -match '(?m)^FAIL:|RUNTIME ERROR:|SCRIPT ERROR:') {throw 'Original issue regression failed.'}
    }
}
$acceptance=Get-Content -LiteralPath (Join-Path $evidence 'Acceptance.json') -Raw | ConvertFrom-Json
$equality=Get-Content -LiteralPath (Join-Path $evidence 'SourceEquality.json') -Raw | ConvertFrom-Json
$validation=Get-Content -LiteralPath (Join-Path $evidence 'Validation.json') -Raw | ConvertFrom-Json
if (-not $acceptance.InputsPreserved -or -not $acceptance.SaveReopenExact -or -not $acceptance.RepeatExact -or $acceptance.PassiveSolveCount -ne 0 -or -not $acceptance.PassiveExportExact) {throw 'Acceptance failed.'}
if (-not $equality.AllPassed -or $equality.RepositoryComponents -ne 132 -or @($equality.GeneratedDocumentChecks).Count -ne 5) {throw 'Source equality failed.'}
if ($validation.Count -ne 27 -or @($validation | Where-Object {-not $_.Passed}).Count) {throw 'Workbook structure failed.'}
foreach ($check in $equality.Checks) {
    $files=@(Get-ChildItem -LiteralPath (Join-Path $root 'src'),(Join-Path $root 'tests') -Recurse -File | Where-Object BaseName -eq $check.Component)
    if ($files.Count -ne 1 -or (Get-FileHash -LiteralPath $files[0].FullName).Hash -cne $check.SourceSHA256) {throw 'A source changed after native verification.'}
}
foreach ($gate in @(@('MaterialPlateau.txt',2047,'PLATEAU_SLOPE'),@('ConfigPresentation.txt',1592,'CONFIG_PRESENTATION'),@('MaterialControls.txt',2191,'MATERIAL_CONTROL_UNITS'),@('ResultsDiagrams.txt',124,'PLATEAU_RESULTS'))) {
    $text=Get-Content -LiteralPath (Join-Path $evidence $gate[0]) -Raw
    if (-not $text.Contains(('TOTAL_{0}: passed={1}; failed=0' -f $gate[2],$gate[1])) -or $text -match '(?m)^FAIL:|RUNTIME ERROR:|SCRIPT ERROR:') {throw 'A directed gate failed.'}
}
$book=Join-Path $root 'workbook/output/RC_Section_NDM.xlsm'
$dump=Join-Path $root 'workbook/output/VBA_All_Code.txt'
$candidate=Join-Path $evidence 'RC_Section_NDM.xlsm'
$candidateDump=Join-Path $evidence 'VBA_All_Code.txt'
if ((Get-FileHash -LiteralPath $book).Hash -cne $acceptance.SourceSHA256 -or (Get-FileHash -LiteralPath $candidate).Hash -cne $acceptance.ReleaseSHA256) {throw 'A workbook changed after verification.'}
if ((Get-FileHash -LiteralPath $dump).Hash -cne $ExpectedDumpSHA256) {throw 'The existing VBA dump changed.'}
if ((Get-FileHash -LiteralPath $candidateDump).Hash -cne (Get-FileHash -LiteralPath (Join-Path $evidence 'FinalVBA.txt')).Hash) {throw 'The dump is not the final native export.'}
$report=Join-Path $root 'workbook/output/RC_Section_NDM_execution_report.txt'
$reportHash=(Get-FileHash -LiteralPath $report).Hash
$backup=Join-Path $evidence 'PublicationBackup'
if (Test-Path -LiteralPath $backup) {throw 'Publication has already been attempted.'}
New-Item -ItemType Directory -Path $backup | Out-Null
$bookStage=$book+'.plateau-stage';$dumpStage=$dump+'.plateau-stage'
$bookLock=$null;$dumpLock=$null;$bookPublished=$false;$dumpPublished=$false
try {
    # Allow atomic replacement but block any simultaneous read/write open by Excel.
    $bookLock=[IO.File]::Open($book,'Open','ReadWrite','Delete')
    $dumpLock=[IO.File]::Open($dump,'Open','ReadWrite','Delete')
    Copy-Item -LiteralPath $candidate -Destination $bookStage
    Copy-Item -LiteralPath $candidateDump -Destination $dumpStage
    [IO.File]::Replace($bookStage,$book,(Join-Path $backup 'RC_Section_NDM.xlsm'))
    $bookPublished=$true
    [IO.File]::Replace($dumpStage,$dump,(Join-Path $backup 'VBA_All_Code.txt'))
    $dumpPublished=$true
} catch {
    if ($bookPublished) {[IO.File]::Replace((Join-Path $backup 'RC_Section_NDM.xlsm'),$book,$null)}
    if ($dumpPublished) {[IO.File]::Replace((Join-Path $backup 'VBA_All_Code.txt'),$dump,$null)}
    throw
} finally {
    if ($bookLock) {$bookLock.Dispose()};if ($dumpLock) {$dumpLock.Dispose()}
    if (Test-Path -LiteralPath $bookStage) {Remove-Item -LiteralPath $bookStage}
    if (Test-Path -LiteralPath $dumpStage) {Remove-Item -LiteralPath $dumpStage}
}
if ((Get-FileHash -LiteralPath $book).Hash -cne $acceptance.ReleaseSHA256 -or (Get-FileHash -LiteralPath $report).Hash -cne $reportHash) {throw 'Published hashes differ or the existing report changed.'}
[ordered]@{WorkbookSHA256=(Get-FileHash -LiteralPath $book).Hash;DumpSHA256=(Get-FileHash -LiteralPath $dump).Hash;ExistingReportPreserved=$true;InputsPreserved=$true;ActualConfigAndResultsDiagrams=$true;RepositoryComponents=132;GeneratedComponents=5;PublishedAt=(Get-Date -Format o)} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $evidence 'Publication.json') -Encoding UTF8
Write-Output 'PUBLISHED: verified workbook and final VBA dump; existing report preserved.'
