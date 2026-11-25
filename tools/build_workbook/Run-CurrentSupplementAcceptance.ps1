# ДЛЯ ТЕСТОВ: обязательное дополнительное покрытие на тех же байтах после полного набора.
param(
    [Parameter(Mandatory=$true)][string]$PrimaryDirectory,
    [Parameter(Mandatory=$true)][string]$Deadline
)
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$primary=(Resolve-Path -LiteralPath (Join-Path $root $PrimaryDirectory)).Path
$run=Get-Content -LiteralPath (Join-Path $primary 'Run.json') -Raw -Encoding UTF8 | ConvertFrom-Json
if($run.Status -eq 'RUNNING' -or $run.CompletedJobs.Count -ne 12 -or -not $run.FinishedUTC -or -not $run.FocusedAcceptance){throw 'All primary groups must reach a terminal state before supplements.'}
$focused=Get-Content -LiteralPath (Join-Path $primary 'Focused/Run.json') -Raw -Encoding UTF8 | ConvertFrom-Json
if(-not $focused.FinishedUTC){throw 'Focused plan has not finished.'}
$frozen=Join-Path $primary 'Source/RC_Section_NDM.xlsm'
$hash=(Get-FileHash -LiteralPath $frozen).Hash
if($hash -cne $run.SourceSHA256 -or $hash -cne $focused.SourceSHA256){throw 'Source provenance differs.'}
$limit=[DateTimeOffset]::Parse($Deadline,[Globalization.CultureInfo]::InvariantCulture)
if([DateTimeOffset]::UtcNow.AddMinutes(35) -ge $limit){throw 'Supplement forecast does not fit the deadline.'}
$directory=Join-Path $primary 'Supplement'
if(Test-Path -LiteralPath $directory){throw 'A fresh supplement directory is required.'}
New-Item -ItemType Directory -Path $directory | Out-Null
Copy-Item -LiteralPath $PSCommandPath -Destination (Join-Path $directory 'SupplementRunner.ps1')
$macros=@(
    'modTestBatchCalculation.RunZeroLoadCombinationTests',
    'modTestTorsionPost.RunTorsionOnlyTests',
    'modTestWorkbookInterface.RunAudit03ExcelGuardTests',
    'modTestWorkbookInterface.RunTestCalculationFreshnessWorkflow',
    'modTestBatchCalculation.RunNormativeStabilityTests',
    'modTestBatchCalculation.RunNormativePartialBranchFailureTest',
    'modTestCrackWidth.RunNormativeBoundaryTests'
)
$patterns=@(
    '(?m)^TOTAL_ZERO_LOAD_COMBINATIONS: passed=([1-9][0-9]*); failed=0\r?$',
    '(?m)^TorsionOnly: Passed=([1-9][0-9]*); Failed=0\r?$',
    '(?m)^TOTAL_AUDIT03_EXCEL_GUARD: passed=([1-9][0-9]*); failed=0\r?$',
    '(?m)^TOTAL_FRESHNESS_WORKFLOW: passed=([1-9][0-9]*); failed=0\r?$',
    '(?m)^TOTAL_NORMATIVE_STABILITY: passed=([1-9][0-9]*); failed=0\r?$',
    '(?m)^TOTAL_NORMATIVE_PARTIAL_BRANCH: passed=11; failed=0\r?$',
    '(?m)^TOTAL_NORMATIVE_BOUNDARY: passed=13; failed=0\r?$'
)
$results=@();$allPassed=$true
foreach($mode in @('No','Yes')){
    $relative=Join-Path $PrimaryDirectory ('Supplement/'+$mode+'.txt')
    # Watchdog проверяет общий резерв памяти перед каждым новым Excel.
    & (Join-Path $PSScriptRoot 'Run-Audit03Watchdog.ps1') -SourceWorkbook $frozen -ReportPath $relative -Mode $mode -Macro ($macros -join ',') -TimeoutSeconds 1800
    $exitCode=$LASTEXITCODE
    $path=Join-Path $root $relative
    $text=[IO.File]::ReadAllText($path)
    $errors=@([regex]::Matches($text,'(?im)^(FAIL(?::| )|RUNTIME ERROR|SCRIPT ERROR|WATCHDOG_FAILURE|RESTORE ERROR|CLEANUP ERROR)[^\r\n]*')|ForEach-Object Value)
    $starts=@([regex]::Matches($text,'(?m)^SUITE_STARTED: ([^;]+);')|ForEach-Object {$_.Groups[1].Value})
    $finishes=@([regex]::Matches($text,'(?m)^SUITE_FINISHED: ([^;]+);')|ForEach-Object {$_.Groups[1].Value})
    $passed=$exitCode -eq 0 -and $errors.Count -eq 0 -and ($starts -join ',') -ceq ($macros -join ',') -and ($finishes -join ',') -ceq ($macros -join ',')
    foreach($pattern in $patterns){$passed=$passed -and ([regex]::Matches($text,$pattern).Count -eq 1)}
    $closure=Get-Content -LiteralPath ($path+'.ProcessClosure.json') -Raw -Encoding UTF8|ConvertFrom-Json
    $passed=$passed -and $closure.Passed -and $text.Contains('WATCHDOG_COMPLETED: exit=0; source unchanged=True') -and ((Get-FileHash -LiteralPath $frozen).Hash -ceq $hash)
    if(-not $passed){$allPassed=$false}
    $results+=,[ordered]@{Mode=$mode;Passed=$passed;ExitCode=$exitCode;Macros=$macros;Report=$relative;ReportSHA256=(Get-FileHash -LiteralPath $path).Hash;Errors=$errors;ProcessClosed=$closure.Passed}
    [ordered]@{Status='RUNNING';SourceSHA256=$hash;ExpectedExecutions=14;Results=$results}|ConvertTo-Json -Depth 7|Set-Content -LiteralPath (Join-Path $directory 'Acceptance.json') -Encoding UTF8
}
[ordered]@{Status=if($allPassed){'PASS'}else{'REVIEW_REQUIRED'};SourceSHA256=$hash;ExpectedExecutions=14;Results=$results;FinishedUTC=[DateTimeOffset]::UtcNow.ToString('o')}|ConvertTo-Json -Depth 7|Set-Content -LiteralPath (Join-Path $directory 'Acceptance.json') -Encoding UTF8
if(-not $allPassed){exit 1}
