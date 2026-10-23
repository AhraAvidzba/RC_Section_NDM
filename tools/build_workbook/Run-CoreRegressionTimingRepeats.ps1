# ДЛЯ ТЕСТОВ: три чередующихся A/B повтора групп с подозрительным временем.
# Полная группа сохраняет прежнее положение холодного и прогретых замеров.
param([string]$Directory='docs/regression/Performance/CoreRegression_2026-10-15',[int]$MaxSeconds=1200)
$ErrorActionPreference='Stop'
$root=(Get-Location).Path
$directoryPath=[IO.Path]::GetFullPath((Join-Path $root $Directory))
if(-not $directoryPath.StartsWith((Join-Path $root 'docs/regression/Performance/'),[StringComparison]::OrdinalIgnoreCase)){throw 'Performance evidence directory required.'}
$summary=Get-Content -LiteralPath (Join-Path $directoryPath 'Comparison.json') -Raw -Encoding UTF8 | ConvertFrom-Json
if(-not $summary.complete -or -not $summary.allPairedOutputsExact){throw 'Complete main comparison required first.'}
$candidates=@($summary.timingRepeatCandidates | Group-Object shape,group,mode | ForEach-Object {$_.Group[0]})
$original=Get-Content -LiteralPath (Join-Path $directoryPath 'Progress.json') -Raw -Encoding UTF8 | ConvertFrom-Json
foreach($version in @('Before','After')){
    $folder=Join-Path $directoryPath $version
    $preparation=Get-Content -LiteralPath (Join-Path $folder 'Preparation.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    if((Get-FileHash -LiteralPath (Join-Path $folder 'RC_Section_NDM.xlsm')).Hash -cne $preparation.ProbeWorkbookSHA256){throw 'Prepared source changed.'}
}
$records=New-Object 'System.Collections.Generic.List[object]'
$stopwatch=[Diagnostics.Stopwatch]::StartNew()
$expectedJobs=$candidates.Count*6
$stopReason=''
try{
    foreach($candidate in $candidates){
        for($repeat=1;$repeat -le 3;$repeat++){
            $order=@('Before','After')
            if($repeat % 2 -eq 0){$order=@('After','Before')}
            foreach($version in $order){
                $remaining=[math]::Floor($MaxSeconds-$stopwatch.Elapsed.TotalSeconds)
                if($remaining -lt 10){throw 'Repeat deadline reached; incomplete evidence.'}
                $id='Repeat'+$repeat+'_'+$version+'_'+$candidate.shape+'_'+$candidate.group+'_'+$candidate.mode
                $reportRelative=$Directory+'/'+$id+'.txt'
                $report=Join-Path $root $reportRelative
                if(Test-Path -LiteralPath $report){throw ('Existing repeat evidence: '+$id)}
                $base=@($original | Where-Object {$_.Version -eq $version -and $_.Shape -eq $candidate.shape -and $_.Group -eq $candidate.group -and $_.Mode -eq $candidate.mode})
                if($base.Count -ne 1){throw 'Original comparison group was not found uniquely.'}
                Write-Output ('REPEAT_STARTED: '+$id)
                $arguments=@{SourceWorkbook=(Join-Path $directoryPath ($version+'/RC_Section_NDM.xlsm'));ReportPath=$reportRelative;Mode=$candidate.mode;TimeoutSeconds=[math]::Min($remaining,360);VerifyResultsReopen=$true;VerifyStatusReopen=$true}
                if($candidate.group -eq 'Batch15'){$arguments.Macro='modTestBatchCalculation.TestRunCoreRegressionBatch15'}
                else{$arguments.Macro='modTestBatchCalculation.TestRunFocusedAcceptanceCases';$arguments.MacroArgument1=$candidate.shape;$arguments.MacroArgument2=$candidate.group}
                & (Join-Path $PSScriptRoot 'Run-Audit03Watchdog.ps1') @arguments
                $exit=$LASTEXITCODE
                $body=Get-Content -LiteralPath $report -Raw -Encoding UTF8
                $match=[regex]::Match($body,'(?m)^TOTAL_(?:FOCUSED_ACCEPTANCE|CORE_BATCH15):[^\r\n]*passed=(\d+); failed=(\d+)')
                $passed=$exit -eq 0 -and $match.Success -and $match.Groups[2].Value -eq '0' -and $body -notmatch '(?m)^(FAIL:|RUNTIME ERROR:|WATCHDOG_FAILURE:|RESTORE ERROR:)'
                $records.Add([ordered]@{ID=$id;Repeat=$repeat;Version=$version;Shape=$candidate.shape;Group=$candidate.group;Mode=$candidate.mode;ExpectedCases=$base[0].ExpectedCases;Passed=$passed;ExitCode=$exit;Report=$reportRelative;ReportSHA256=(Get-FileHash -LiteralPath $report).Hash})
                ConvertTo-Json -InputObject @($records.ToArray()) -Depth 5 | Set-Content -LiteralPath (Join-Path $directoryPath 'TimingRepeatsProgress.json') -Encoding UTF8
                Write-Output ('REPEAT_FINISHED: '+$id+'; passed='+$passed+'; jobs='+$records.Count+'/'+$expectedJobs)
            }
        }
    }
}catch{$stopReason=$_.Exception.Message;throw}
finally{
    $stopwatch.Stop()
    [ordered]@{AllJobsExecuted=($records.Count -eq $expectedJobs);AllAssertionsPassed=($records.Count -eq $expectedJobs -and @($records | Where-Object {-not $_.Passed}).Count -eq 0);ExpectedJobs=$expectedJobs;ElapsedSeconds=$stopwatch.Elapsed.TotalSeconds;StopReason=$stopReason;Records=$records.ToArray()} | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $directoryPath 'TimingRepeatsRun.json') -Encoding UTF8
}
