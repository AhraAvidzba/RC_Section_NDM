# Directed exact per-LC isolation, not a benchmark and not the old full On/Off suite.
param([string]$SourceWorkbook='docs/regression/Performance/FinalCandidate/RC_Section_NDM.xlsm',
    [string]$Directory='docs/regression/Performance/IsolationFinal')
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$source=(Resolve-Path -LiteralPath (Join-Path $root $SourceWorkbook)).Path
$output=Join-Path $root $Directory
New-Item -ItemType Directory -Path $output -Force | Out-Null
$records=@()
foreach($family in @('CRACK_SP35','CRACK_SP63','STABILITY_SP35','STABILITY_SP63','CAPACITY','MIXED','TINY','REPEAT','IMPORTED_CRACK_SP35')) {
    $report=Join-Path $Directory ($family+'.txt')
    $timer=[Diagnostics.Stopwatch]::StartNew()
    & (Join-Path $PSScriptRoot 'Run-Audit03Watchdog.ps1') -SourceWorkbook $source -ReportPath $report -Macro 'modTestPerformance.RunPerformanceBatchIsolationTests' -MacroArgument1 $family -Mode Yes -TimeoutSeconds 600
    $exitCode=$LASTEXITCODE;$timer.Stop()
    $text=Get-Content -LiteralPath (Join-Path $root $report) -Raw -Encoding UTF8
    $passed=$exitCode -eq 0 -and $text -notmatch '(?m)^FAIL:' -and $text -match 'failed=0'
    $records += [ordered]@{Family=$family;Seconds=$timer.Elapsed.TotalSeconds;ExitCode=$exitCode;Passed=$passed;Report=$report}
    $records | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $output 'Summary.json') -Encoding UTF8
    Write-Output ($records[-1] | ConvertTo-Json -Compress)
    if(-not $passed) {throw "Isolation gate failed: $family"}
}
