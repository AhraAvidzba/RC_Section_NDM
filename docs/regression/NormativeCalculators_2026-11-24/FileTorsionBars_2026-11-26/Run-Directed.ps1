# ДЛЯ ТЕСТОВ: связанный вывод кручения после расчёта и при обновлении File/Excel.
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../../../..')).Path
$prefix='docs/regression/NormativeCalculators_2026-11-24/FileTorsionBars_2026-11-26/'
$source=Join-Path $PSScriptRoot 'Candidate.xlsm'
$jobs=@(
    @{Name='Tables';Macros='modTestTorsionPost.RunTorsionPostOutputTests,modTestTorsionPost.RunTorsionBarLengthTests,modTestTorsionPost.RunTorsionPostGuardTests';Argument='';Count=3},
    @{Name='File';Macros='modTestTorsionConcrete.RunTorsionConcreteWorkbookTests';Argument='File';Count=1},
    @{Name='Excel';Macros='modTestTorsionConcrete.RunTorsionConcreteWorkbookTests';Argument='Excel';Count=1}
)
$runs=@()
foreach($mode in @('No','Yes')){
    foreach($job in $jobs){
        $os=Get-CimInstance Win32_OperatingSystem
        $memory=Get-CimInstance Win32_PerfFormattedData_PerfOS_Memory
        if($os.FreePhysicalMemory/1MB -lt 7 -or ($memory.CommitLimit-$memory.CommittedBytes)/1GB -lt 7){throw 'Insufficient memory reserve for Excel'}
        $report=$prefix+$job.Name+'_'+$mode+'.txt'
        if(Test-Path -LiteralPath (Join-Path $root $report)){
            Write-Output ('REVIEW_EXISTING '+$job.Name+' '+$mode)
            $closure=Get-Content -LiteralPath ((Join-Path $root $report)+'.ProcessClosure.json') -Raw | ConvertFrom-Json
            if(-not $closure.Passed -or $closure.ChildExitCode -ne 0){throw 'Existing run did not close successfully'}
        }else{
            Write-Output ('START '+$job.Name+' '+$mode)
            & (Join-Path $root 'tools/build_workbook/Run-Audit03Watchdog.ps1') -SourceWorkbook $source -ReportPath $report -Macro $job.Macros -MacroArgument1 $job.Argument -Mode $mode -TimeoutSeconds 600
            if($LASTEXITCODE -ne 0){throw ('Related tests failed: '+$report)}
        }
        $body=[IO.File]::ReadAllText((Join-Path $root $report))
        if($body -match '(?m)^FAIL\b|Failed=[1-9]|failed=[1-9]|^RUNTIME[: ]|RUNTIME ERROR|WATCHDOG_FAILURE'){throw ('Failed assertions: '+$report)}
        if(-not $body.Contains('SOURCE_SHA256: '+(Get-FileHash -LiteralPath $source).Hash)){throw 'Source provenance differs'}
        if([regex]::Matches($body,'(?m)^SUITE_FINISHED:').Count -ne $job.Count){throw 'Incomplete related scope'}
        $runs+=,[ordered]@{Job=$job.Name;Mode=$mode;Macros=$job.Macros.Split(',');Argument=$job.Argument;Report=$report;SHA256=(Get-FileHash -LiteralPath (Join-Path $root $report)).Hash;Passed=$true}
        $runs|ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $PSScriptRoot 'Runs.json') -Encoding UTF8
    }
}
Write-Output 'RELATED_COMPLETE: ten macro executions; no full suite started.'
