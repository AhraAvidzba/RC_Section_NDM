# Final stage jobs stay sequential. They use only private performance copies;
# a failing measurement or directed test stops publication preparation.
param([ValidateSet('Export','Imported','UserInputs','Readers','Release','Build','Validate')][string]$StartAt='Export')
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$baseline='docs/regression/Performance/FinalBaselineV2/RC_Section_NDM.xlsm'
$candidate='docs/regression/Performance/FinalCandidateV2/RC_Section_NDM.xlsm'
$order=@('Export','Imported','UserInputs','Readers','Build','Release','Validate')
$started=$false
foreach($stage in $order) {
    if($stage -eq $StartAt) {$started=$true}
    if(-not $started) {continue}
    $timer=[Diagnostics.Stopwatch]::StartNew()
    switch($stage) {
        'Export' {& (Join-Path $PSScriptRoot 'Measure-PerformanceStorage.ps1') -Group Export -Directory docs/regression/Performance/ExportMeasurements -BaselineWorkbook $baseline -CandidateWorkbook $candidate -Family DIRECT}
        'Imported' {& (Join-Path $PSScriptRoot 'Measure-PerformanceStorage.ps1') -Group Batch -Directory docs/regression/Performance/ImportedMeasurements -BaselineWorkbook $baseline -CandidateWorkbook $candidate -Family IMPORTED_CRACK_SP35}
        'UserInputs' {& (Join-Path $PSScriptRoot 'Measure-PerformanceStorage.ps1') -Group Batch -Directory docs/regression/Performance/UserInputsMeasurements -BaselineWorkbook $baseline -CandidateWorkbook $candidate -Family USER_SAVED -Counts 1 -BatchModes Full -RunsPerOpen 2 -KeepInputs}
        'Readers' {
            foreach($macro in @('modTestWorkbookInterface.RunAudit03ReadLifecycleTests','modTestWorkbookInterface.RunAudit03ImportedSnapshotUnitChangeTests')) {
                $name=($macro -split '\.')[-1]
                $report="docs/regression/Performance/FinalReaders/$name.txt"
                & (Join-Path $PSScriptRoot 'Run-Audit03Watchdog.ps1') -SourceWorkbook (Join-Path $root $candidate) -ReportPath $report -Macro $macro -Mode No -TimeoutSeconds 600
                if($LASTEXITCODE -ne 0) {throw "Reader gate failed: $macro"}
            }
        }
        'Release' {
            & (Join-Path $PSScriptRoot 'Prepare-PerformanceRelease.ps1') -SourceWorkbook $candidate -Directory docs/regression/Performance/ReleaseV2
            & (Join-Path $PSScriptRoot 'Export-Audit03VbaSnapshot.ps1') -WorkbookPath docs/regression/Performance/ReleaseV2/RC_Section_NDM.xlsm -OutputPath docs/regression/Performance/ReleaseV2/VBA_All_Code.txt
            & (Join-Path $PSScriptRoot 'Export-Audit03VbaSnapshot.ps1') -WorkbookPath docs/regression/Performance/CompiledFreshBuildV2/RC_Section_NDM.xlsm -OutputPath docs/regression/Performance/CompiledFreshBuildV2/VBA_All_Code.txt
            & (Join-Path $PSScriptRoot 'Compare-PerformanceVba.ps1') -ExportPath docs/regression/Performance/ReleaseV2/VBA_All_Code.txt -ReportPath docs/regression/Performance/ReleaseV2/SourceEquality.json -ReferenceExportPath docs/regression/Performance/CompiledFreshBuildV2/VBA_All_Code.txt
        }
        'Build' {
            & (Join-Path $PSScriptRoot 'Build-Workbook.ps1') -OutputPath docs/regression/Performance/FreshBuild/RC_Section_NDM.xlsm
            & (Join-Path $PSScriptRoot 'Prepare-PerformanceWorkbook.ps1') -Directory docs/regression/Performance/CompiledFreshBuildV2 -SourceWorkbook docs/regression/Performance/FreshBuild/RC_Section_NDM.xlsm -TestModuleOnly -SmokeMacro modTestWorkbookInterface.RunConfigPresentationTests
        }
        'Validate' {
            & powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'Validate-Workbook.ps1') -WorkbookPath docs/regression/Performance/CompiledFreshBuildV2/RC_Section_NDM.xlsm -ReportPath docs/regression/Performance/CompiledFreshBuildV2/Validation.json
            if($LASTEXITCODE -ne 0) {throw 'Fresh build validation process failed.'}
            $checks=Get-Content -LiteralPath (Join-Path $root 'docs/regression/Performance/CompiledFreshBuildV2/Validation.json') -Raw | ConvertFrom-Json
            if(@($checks | Where-Object {-not $_.Passed}).Count -gt 0) {throw 'Fresh build validation failed.'}
            & powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'Validate-Workbook.ps1') -WorkbookPath docs/regression/Performance/ReleaseV2/RC_Section_NDM.xlsm -ReportPath docs/regression/Performance/ReleaseV2/Validation.json -UserConfiguredWorkbook
            if($LASTEXITCODE -ne 0) {throw 'User release validation process failed.'}
            $checks=Get-Content -LiteralPath (Join-Path $root 'docs/regression/Performance/ReleaseV2/Validation.json') -Raw | ConvertFrom-Json
            if(@($checks | Where-Object {-not $_.Passed}).Count -gt 0) {throw 'User release validation failed.'}
        }
    }
    $timer.Stop()
    [ordered]@{Stage=$stage;Passed=$true;Seconds=$timer.Elapsed.TotalSeconds} | ConvertTo-Json -Compress | Add-Content -LiteralPath (Join-Path $root 'docs/regression/Performance/FinalStages.jsonl') -Encoding UTF8
    Write-Output "FINAL_STAGE_COMPLETED: $stage; seconds=$($timer.Elapsed.TotalSeconds)"
}
