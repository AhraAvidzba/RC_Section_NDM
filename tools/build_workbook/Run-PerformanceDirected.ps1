# Runs only directed gates of the performance stage, never the old full On/Off suite.
# Each watchdog owns its copy and Excel PID; the source book is hash-guarded.
param(
    [string]$SourceWorkbook = 'docs/regression/Performance/BatchV3Candidate/RC_Section_NDM.xlsm',
    [string]$Directory = 'docs/regression/Performance/DirectedFinal'
)
$ErrorActionPreference = 'Stop'
$root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$allowed = [IO.Path]::GetFullPath((Join-Path $root 'docs/regression/Performance')) + [IO.Path]::DirectorySeparatorChar
$output = [IO.Path]::GetFullPath((Join-Path $root $Directory))
if (-not $output.StartsWith($allowed,[StringComparison]::OrdinalIgnoreCase)) { throw 'Evidence path is outside Performance.' }
New-Item -ItemType Directory -Path $output -Force | Out-Null
$source = (Resolve-Path -LiteralPath (Join-Path $root $SourceWorkbook)).Path
$hash = (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash
$tests = @(
    'modTestGeometryQuery.RunGeometryQueryTests',
    'modTestSP35CrackWidth.RunSP35PreparationTests',
    'modTestSP35CrackWidth.RunSP35LocalOpeningTests',
    'modTestSP35CrackWidth.RunSP35CurrentProjectionTests',
    'modTestSP35CrackWidth.RunSP35EndToEndTests',
    'modTestCapacitySolver.RunAudit03CapacityContracts',
    'modTestCapacitySolver.RunAudit03CapacityLifecycle',
    'modTestCapacitySolver.RunAudit03MultiplierTypedFaults',
    'modTestBatchCalculation.RunPostAudit03LifecycleTests',
    'modTestBatchCalculation.RunAudit03PublishedResultTests',
    'modTestBatchCalculation.RunPostAudit03SnapshotOwnerTests',
    'modTestBatchCalculation.RunPostAudit03SnapshotRepeatTests',
    'modTestBatchCalculation.RunAudit03StatusPaletteTests',
    'modTestWorkbookInterface.RunAudit03ReadLifecycleTests',
    'modTestWorkbookInterface.RunAudit03ImportedSnapshotUnitChangeTests',
    'modTestWorkbookInterface.RunConfigPresentationTests',
    'modTestWorkbookInterface.RunAudit03ExcelGuardTests'
)
$records = @()
foreach ($macro in $tests) {
    $name = ($macro -split '\.')[-1]
    $report = Join-Path $Directory ($name + '.txt')
    $timer = [Diagnostics.Stopwatch]::StartNew()
    & (Join-Path $PSScriptRoot 'Run-Audit03Watchdog.ps1') -SourceWorkbook $source -ReportPath $report -Macro $macro -TimeoutSeconds 600 -Mode No
    $exitCode = $LASTEXITCODE
    $timer.Stop()
    $text = Get-Content -LiteralPath (Join-Path $root $report) -Encoding UTF8 -Raw
    $passed = $exitCode -eq 0 -and $text -notmatch '(?m)^FAIL:' -and $text -match 'failed=0'
    $records += [ordered]@{Macro=$macro;Seconds=$timer.Elapsed.TotalSeconds;ExitCode=$exitCode;Passed=$passed;Report=$report}
    $records | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $output 'Summary.json') -Encoding UTF8
    Write-Output ($records[-1] | ConvertTo-Json -Compress)
    if ((Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash -ne $hash) { throw 'Source changed.' }
}
if (@($records | Where-Object {-not $_.Passed}).Count -gt 0) { throw 'Some directed gates failed; inspect the preserved reports.' }
