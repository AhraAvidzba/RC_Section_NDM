# Publishes only the verified candidate, with a guard against a new user save.
$ErrorActionPreference = 'Stop'
$root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../../..')).Path
$publication = Join-Path $PSScriptRoot 'Publication'
$manifest = Get-Content -LiteralPath (Join-Path $publication 'Manifest.json') -Raw | ConvertFrom-Json
$source = Join-Path $root 'workbook/output/RC_Section_NDM.xlsm'
$candidate = Join-Path $publication 'RC_Section_NDM.xlsm'
$checks = Get-Content -LiteralPath (Join-Path $publication 'StructureChecks.json') -Raw | ConvertFrom-Json
if ($checks.Count -ne 27 -or @($checks | Where-Object { -not $_.Passed }).Count -gt 0) { throw 'Structure checks are not successful.' }
if ((Get-FileHash -LiteralPath $source).Hash -ne $manifest.SourceSHA256) { throw 'User workbook changed; publication cancelled.' }
if ((Get-FileHash -LiteralPath $candidate).Hash -ne $manifest.PreparedSHA256) { throw 'Candidate changed; publication cancelled.' }
if ((Get-Content -LiteralPath (Join-Path $publication 'SourceContracts.txt') -Raw) -notmatch 'matchingModules=119; sourceModules=119; failed=0') { throw 'Source contracts are not successful.' }
foreach ($test in @(
        @('modTestGeometryQuery.RunGeometryQueryTests', 'TOTAL: passed=987; failed=0'),
        @('modTestSP35CrackWidth.RunSP35PreparationTests', 'TOTAL_SP35_PREPARATION: passed=201; failed=0'),
        @('modTestSP35CrackWidth.RunSP35NeighborSettingsTests', 'TOTAL_SP35_NEIGHBOR_SETTINGS: passed=18; failed=0'),
        @('modTestSP35CrackWidth.RunSP35LocalOpeningTests', 'TOTAL_SP35_LOCAL_OPENING: passed=36; failed=0'),
        @('modTestSP35CrackWidth.RunSP35FormulaTests', 'TOTAL_SP35_FORMULA: passed=27; failed=0'),
        @('modTestSP35CrackWidth.RunSP35PipelineTests', 'TOTAL_SP35_PIPELINE: passed=34; failed=0'),
        @('modTestSP35CrackWidth.RunSP35EndToEndTests', 'TOTAL_SP35_END_TO_END: passed=270; failed=0'),
        @('modTestSP35CrackWidth.RunSP35CurrentProjectionTests', 'TOTAL_SP35_CURRENT_PROJECTION: passed=57; failed=0'),
        @('modTestWorkbookInterface.RunConfigPresentationTests', 'TOTAL_CONFIG_PRESENTATION: passed=1583; failed=0'))) {
    $path = Join-Path $PSScriptRoot ("FinalSavedCode/$($test[0]).txt")
    $text = Get-Content -LiteralPath $path -Raw
    if (-not $text.Contains($test[1]) -or $text -match '(?m)^FAIL:') { throw "Saved-code test failed: $($test[0])" }
}
$report = Join-Path $root 'workbook/output/RC_Section_NDM_execution_report.txt'
$reportHash = (Get-FileHash -LiteralPath $report).Hash
Copy-Item -LiteralPath $candidate -Destination $source -Force
Copy-Item -LiteralPath (Join-Path $publication 'VBA_All_Code.txt') -Destination (Join-Path $root 'workbook/output/VBA_All_Code.txt') -Force
if ((Get-FileHash -LiteralPath $source).Hash -ne $manifest.PreparedSHA256) { throw 'Published workbook hash mismatch.' }
if ((Get-FileHash -LiteralPath $report).Hash -ne $reportHash) { throw 'Calculation report unexpectedly changed.' }
Write-Output "PUBLISHED: $($manifest.PreparedSHA256); execution report preserved"
