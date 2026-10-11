# The load stage of FocusedValidationPlan; no old exhaustive suite is called.
param(
    [string]$ProbeDirectory='docs/regression/FocusedAcceptance_2026-10-10_V2',
    [ValidateRange(60,3600)][int]$MaxSeconds=1800,
    [switch]$RunCalculations,
    [switch]$ResumeIncomplete
)
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$directory=Join-Path $root $ProbeDirectory
$source=Join-Path $directory 'RC_Section_NDM.xlsm'
$preparation=Get-Content -LiteralPath (Join-Path $directory 'Preparation.json') -Raw | ConvertFrom-Json
$manifest=Get-Content -LiteralPath (Join-Path $directory 'InputManifest.json') -Raw | ConvertFrom-Json
if (-not $preparation.Compiled -or -not $manifest.Passed -or
    $preparation.ProbeWorkbookSHA256 -cne (Get-FileHash -LiteralPath $source).Hash -or
    $preparation.SourceSHA256 -cne (Get-FileHash -LiteralPath $preparation.SourceWorkbook).Hash -or
    $preparation.HelperSHA256 -cne (Get-FileHash -LiteralPath (Join-Path $PSScriptRoot 'FocusedAcceptanceProbe.bas')).Hash) {
    throw 'Source or selected inputs changed; prepare a new probe before running.'
}
$jobs=@()
foreach ($group in $manifest.Groups) {
    foreach ($mode in @('No','Yes')) {
        $timeout=150
        if ($group.Group -eq 'Saved') { $timeout=350 }
        if ($group.Group -eq 'Paths') { $timeout=100 }
        $jobs+=[pscustomobject]@{Shape=$group.Shape; Group=$group.Group; Mode=$mode; Cases=$group.CasesPerMode; Timeout=$timeout}
    }
}
if ($jobs.Count -ne 12 -or ($jobs | Measure-Object Cases -Sum).Sum -ne 122) { throw 'Selected coverage differs from plan.' }
if (-not $RunCalculations) {
    $jobs | Format-Table Shape,Group,Mode,Cases,Timeout -AutoSize
    Write-Output ('DRY_RUN_ONLY: 122 selected cases; max stage time='+$MaxSeconds+' sec; no calculation executed.')
    exit 0
}
$summaryPath=Join-Path $directory 'LoadStage.json'
if ((Test-Path -LiteralPath $summaryPath) -and -not $ResumeIncomplete) { throw 'Existing evidence cannot be overwritten; prepare a new probe directory.' }
$watch=[Diagnostics.Stopwatch]::StartNew()
$records=@(); $unresolved=@(); $reason=''; $errorText=''
if ($ResumeIncomplete) {
    $previous=Get-Content -LiteralPath $summaryPath -Raw | ConvertFrom-Json
    if ($previous.LoadStageComplete -or $previous.SourceSHA256 -cne $preparation.SourceSHA256 -or
        $previous.ProbeSHA256 -cne $preparation.ProbeWorkbookSHA256) { throw 'Cannot resume completed or changed evidence.' }
    $archive=Join-Path $directory 'LoadStage.initial.json'
    if (Test-Path -LiteralPath $archive) { throw 'A previous resume already exists.' }
    Copy-Item -LiteralPath $summaryPath -Destination $archive
    $records=@($previous.Groups); $unresolved=@($previous.NumericalFailuresForReview)
}
try {
    foreach ($job in $jobs) {
        if (@($records | Where-Object {$_.Shape -eq $job.Shape -and $_.Group -eq $job.Group -and $_.Mode -eq $job.Mode}).Count) { continue }
        $remaining=[math]::Floor($MaxSeconds-$watch.Elapsed.TotalSeconds)
        if ($remaining -lt 10) { $reason='Incomplete: stage deadline reached'; break }
        $reportRelative=$ProbeDirectory+'/'+$job.Shape+'_'+$job.Group+'_'+$job.Mode+'.txt'
        $report=Join-Path $root $reportRelative
        $timeout=[math]::Min($job.Timeout,$remaining)
        Write-Output ('FOCUSED_STARTED: '+$job.Shape+'/'+$job.Group+'/'+$job.Mode+'; remaining='+$remaining+' sec')
        & (Join-Path $PSScriptRoot 'Run-Audit03Watchdog.ps1') -SourceWorkbook $source -ReportPath $reportRelative `
            -Macro modTestBatchCalculation.RunFocusedAcceptanceCases -MacroArgument1 $job.Shape `
            -MacroArgument2 $job.Group -Mode $job.Mode -TimeoutSeconds $timeout -VerifyResultsReopen -VerifyStatusReopen
        $code=$LASTEXITCODE
        $body=Get-Content -LiteralPath $report -Raw -Encoding UTF8
        $total=[regex]::Match($body,'TOTAL_FOCUSED_ACCEPTANCE: shape=([^;]+); group=([^;]+); cases=(\d+); maxSeconds=([^;]+); passed=(\d+); failed=(\d+)')
        $completed=$code -eq 0 -and $total.Success -and $body -notmatch 'WATCHDOG_FAILURE|RUNTIME ERROR|SCRIPT ERROR|RESTORE ERROR|(?m)^FAIL:'
        $record=[ordered]@{Shape=$job.Shape; Group=$job.Group; Mode=$job.Mode; Completed=$completed;
            ExpectedCases=$job.Cases; Report=$reportRelative; ReportSHA256=(Get-FileHash -LiteralPath $report).Hash}
        if ($total.Success) {
            $record.Cases=[int]$total.Groups[3].Value
            $record.MaximumSeconds=[double]::Parse($total.Groups[4].Value,[Globalization.CultureInfo]::InvariantCulture)
            $record.Passed=[int]$total.Groups[5].Value; $record.Failed=[int]$total.Groups[6].Value
            if ($record.Cases -ne $job.Cases -or $record.Failed -gt 0 -or $record.MaximumSeconds -gt 10) {
                $record.Completed=$false; $completed=$false
            }
        }
        foreach ($line in $body -split '\r?\n' | Where-Object { $_ -like 'FOCUSED_UNRESOLVED|*' }) {
            $unresolved+=@{Mode=$job.Mode; Report=$reportRelative; Detail=$line}
        }
        $records+=$record
        $records | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $directory 'LoadProgress.json') -Encoding UTF8
        if (-not $completed -and -not $ResumeIncomplete) { $reason='Incomplete: a selected group failed or exceeded its time limit'; break }
    }
} catch {
    $reason='Incomplete: runner error'; $errorText=$_.Exception.Message
} finally {
    $watch.Stop()
    $sourceUnchanged=(Get-FileHash -LiteralPath $source).Hash -ceq $preparation.ProbeWorkbookSHA256
    $productionUnchanged=(Get-FileHash -LiteralPath $preparation.SourceWorkbook).Hash -ceq $preparation.SourceSHA256
    $complete=$records.Count -eq $jobs.Count -and @($records | Where-Object {$_.Cases -ne $_.ExpectedCases}).Count -eq 0
    $passed=$complete -and @($records | Where-Object {-not $_.Completed}).Count -eq 0 -and $unresolved.Count -eq 0 -and $sourceUnchanged -and $productionUnchanged
    [ordered]@{Passed=$passed; LoadStageComplete=$complete; AcceptanceComplete=$false;
        SourceSHA256=$preparation.SourceSHA256; ProbeSHA256=$preparation.ProbeWorkbookSHA256;
        ElapsedSeconds=$watch.Elapsed.TotalSeconds; LimitSeconds=$MaxSeconds; Groups=$records;
        NumericalFailuresForReview=$unresolved; StopReason=$reason; Error=$errorText;
        SourceUnchanged=$sourceUnchanged; ProductionUnchanged=$productionUnchanged;
        RemainingPlanStages=@('Batch15','NativeCAD','StabilityAndAsymmetricSteel','IndependentControl','FinalOutput');
        ExhaustiveSuiteExecuted=$false} |
        ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $summaryPath -Encoding UTF8
}
Write-Output ('FOCUSED_LOAD_STAGE: complete='+$complete+'; passed='+$passed+'; unresolved='+$unresolved.Count+'; '+$reason)
if (-not $passed) { exit 1 }
