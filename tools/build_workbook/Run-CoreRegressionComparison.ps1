# Сравнивает согласованные 122 случая и batch-15 до/после, последовательно.
param([string]$Directory='docs/regression/Performance/CoreRegression_2026-10-15',[int]$MaxSeconds=3600)
$ErrorActionPreference='Stop'
$root=(Get-Location).Path
$directoryPath=[IO.Path]::GetFullPath((Join-Path $root $Directory))
if(-not $directoryPath.StartsWith((Join-Path $root 'docs/regression/Performance/'),[StringComparison]::OrdinalIgnoreCase)){throw 'Performance evidence directory required.'}
$records=New-Object 'System.Collections.Generic.List[object]'
$manifest=@{}
foreach($version in @('Before','After')){
    $folder=Join-Path $directoryPath $version
    $preparation=Get-Content -LiteralPath (Join-Path $folder 'Preparation.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    $manifest[$version]=Get-Content -LiteralPath (Join-Path $folder 'InputManifest.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    if(-not $preparation.Compiled -or -not $manifest[$version].Passed -or
       (Get-FileHash -LiteralPath (Join-Path $folder 'RC_Section_NDM.xlsm')).Hash -cne $preparation.ProbeWorkbookSHA256){throw 'Prepared source changed.'}
}
$jobs=@()
foreach($group in $manifest['After'].Groups){
    $old=$manifest['Before'].Groups | Where-Object {$_.Shape -eq $group.Shape -and $_.Group -eq $group.Group}
    if(-not $old -or $old.ReportSHA256 -cne $group.ReportSHA256){throw 'A/B manifests differ.'}
    foreach($mode in @('No','Yes')){
        $jobs+=[pscustomobject]@{Shape=$group.Shape;Group=$group.Group;Mode=$mode;Cases=$group.CasesPerMode;Macro='modTestBatchCalculation.TestRunFocusedAcceptanceCases';Timeout=360}
    }
}
if($jobs.Count -ne 12 -or ($jobs | Measure-Object Cases -Sum).Sum -ne 122){throw '122-case coverage mismatch.'}
foreach($mode in @('No','Yes')){$jobs+=[pscustomobject]@{Shape='SavedWorkbook';Group='Batch15';Mode=$mode;Cases=15;Macro='modTestBatchCalculation.TestRunCoreRegressionBatch15';Timeout=360}}
$stopwatch=[Diagnostics.Stopwatch]::StartNew()
$stoppedReason='';$jobIndex=0
try{
    foreach($job in $jobs){
        $order=@('Before','After')
        if($jobIndex % 2 -eq 1){$order=@('After','Before')}
        foreach($version in $order){
            $remaining=[math]::Floor($MaxSeconds-$stopwatch.Elapsed.TotalSeconds)
            if($remaining -lt 10){$stoppedReason='Deadline reached; incomplete jobs remain.';break}
            $id=$version+'_'+$job.Shape+'_'+$job.Group+'_'+$job.Mode
            $reportRelative=$Directory+'/'+$id+'.txt'
            $report=Join-Path $root $reportRelative
            if(Test-Path -LiteralPath $report){throw ('Existing run evidence: '+$id)}
            [ordered]@{ActiveJob=$id;ElapsedSeconds=$stopwatch.Elapsed.TotalSeconds;CompletedJobs=$records.Count;TotalJobs=28;RunnerPID=$PID} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $directoryPath 'Running.json') -Encoding UTF8
            Write-Output ('CORE_STARTED: '+$id+'; remaining='+$remaining+' s')
            $arguments=@{SourceWorkbook=(Join-Path $directoryPath ($version+'/RC_Section_NDM.xlsm'));ReportPath=$reportRelative;Macro=$job.Macro;Mode=$job.Mode;TimeoutSeconds=[math]::Min($remaining,$job.Timeout);VerifyResultsReopen=$true;VerifyStatusReopen=$true}
            if($job.Group -ne 'Batch15'){$arguments.MacroArgument1=$job.Shape;$arguments.MacroArgument2=$job.Group}
            & (Join-Path $PSScriptRoot 'Run-Audit03Watchdog.ps1') @arguments
            $exit=$LASTEXITCODE
            $body=Get-Content -LiteralPath $report -Raw -Encoding UTF8
            $match=[regex]::Match($body,'(?m)^TOTAL_(?:FOCUSED_ACCEPTANCE|CORE_BATCH15):[^\r\n]*passed=(\d+); failed=(\d+)')
            $passed=$exit -eq 0 -and $match.Success -and $match.Groups[2].Value -eq '0' -and $body -notmatch '(?m)^(FAIL:|RUNTIME ERROR:|WATCHDOG_FAILURE:|RESTORE ERROR:)'
            $records.Add([ordered]@{ID=$id;Version=$version;Shape=$job.Shape;Group=$job.Group;Mode=$job.Mode;ExpectedCases=$job.Cases;Passed=$passed;ExitCode=$exit;Report=$reportRelative;ReportSHA256=(Get-FileHash -LiteralPath $report).Hash})
            ConvertTo-Json -InputObject @($records.ToArray()) -Depth 5 | Set-Content -LiteralPath (Join-Path $directoryPath 'Progress.json') -Encoding UTF8
            Write-Output ('CORE_FINISHED: '+$id+'; passed='+$passed+'; jobs='+$records.Count+'/28')
        }
        if($stoppedReason){break}
        $jobIndex++
    }
}catch{$stoppedReason=$_.Exception.Message;throw}
finally{
    $stopwatch.Stop()
    [ordered]@{AllJobsExecuted=($records.Count -eq 28);AllAssertionsPassed=($records.Count -eq 28 -and @($records | Where-Object {-not $_.Passed}).Count -eq 0);ElapsedSeconds=$stopwatch.Elapsed.TotalSeconds;StopReason=$stoppedReason;Records=$records.ToArray();OldExhaustiveMatrixExecuted=$false} | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $directoryPath 'Run.json') -Encoding UTF8
}
