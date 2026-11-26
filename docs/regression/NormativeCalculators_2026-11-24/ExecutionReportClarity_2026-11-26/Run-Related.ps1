# ДЛЯ ТЕСТОВ: оформление диагностик, причины образования трещин и счетчик исключенных сочетаний.
param([string]$Suffix='')
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../../../..')).Path
$prefix='docs/regression/NormativeCalculators_2026-11-24/ExecutionReportClarity_2026-11-26/'
$macros='modTestWorkbookInterface.RunExecutionReportFormattingTests,modTestBatchCalculation.RunAudit03FormationReportTests,modTestBatchCalculation.RunZeroLoadCombinationTests,modTestBatchCalculation.RunNormativeStabilityTests,modTestTorsionPost.RunTorsionOnlyTests'
$runs=@()
foreach($mode in @('No','Yes')){
    $os=Get-CimInstance Win32_OperatingSystem
    $memory=Get-CimInstance Win32_PerfFormattedData_PerfOS_Memory
    if($os.FreePhysicalMemory/1MB -lt 7 -or ($memory.CommitLimit-$memory.CommittedBytes)/1GB -lt 7){throw 'Insufficient memory reserve'}
    $report=$prefix+'Related_'+$mode+$Suffix+'.txt'
    if(Test-Path -LiteralPath (Join-Path $root $report)){throw 'Fresh evidence required'}
    & (Join-Path $root 'tools/build_workbook/Run-Audit03Watchdog.ps1') -SourceWorkbook (Join-Path $PSScriptRoot 'Candidate.xlsm') -ReportPath $report -Macro $macros -Mode $mode -TimeoutSeconds 900
    if($LASTEXITCODE -ne 0){throw ('Related execution failed: '+$mode)}
    $body=[IO.File]::ReadAllText((Join-Path $root $report))
    if($body -match '(?m)^FAIL\b|[Ff]ailed=[1-9]|^RUNTIME[: ]|WATCHDOG_FAILURE'){throw ('Related assertions failed: '+$report)}
    if([regex]::Matches($body,'(?m)^SUITE_FINISHED:').Count -ne 5){throw 'Related scope incomplete'}
    $runs+=,[ordered]@{Mode=$mode;Macros=$macros.Split(',');Report=$report;SHA256=(Get-FileHash -LiteralPath (Join-Path $root $report)).Hash;Passed=$true}
    $runs | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $PSScriptRoot ('RelatedRuns'+$Suffix+'.json')) -Encoding UTF8
}
Write-Output 'RELATED_COMPLETE: 10 macro executions; no full run.'
