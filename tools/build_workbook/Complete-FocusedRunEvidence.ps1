# Records a negative/incomplete acceptance honestly; never publishes a workbook.
param([datetime]$StartedUtc='2026-10-10T19:23:38Z', [datetime]$DeadlineUtc='2026-10-10T20:23:38Z')
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$first=Join-Path $root 'docs/regression/FocusedAcceptance_2026-10-10_V5'
$latest=Join-Path $root 'docs/regression/FocusedAcceptance_2026-10-10_V9'
$target=Join-Path $first 'RunOutcome.json'
if(Test-Path -LiteralPath $target) { throw 'Final evidence already exists.' }
$tailPath=Join-Path $latest 'TailStage.json'
if(-not(Test-Path -LiteralPath $tailPath)) { throw 'Addressed test runner is still active or did not finalize.' }
function Read-Json([string]$Path) { Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json }
function Summarize-Report([string]$Path) {
    $body=Get-Content -LiteralPath $Path -Raw -Encoding UTF8
    $totals=[regex]::Matches($body,'(?m)^TOTAL[^\r\n]*passed=(\d+);\s*failed=(\d+)')
    $passed=0; $failed=0
    foreach($total in $totals) { $passed+=[int]$total.Groups[1].Value; $failed+=[int]$total.Groups[2].Value }
    return @{Path=$Path; SHA256=(Get-FileHash -LiteralPath $Path).Hash; Suites=$totals.Count;
        Passed=$passed; Failed=$failed; Completed=($body -match 'WATCHDOG_COMPLETED: exit=0;')}
}
$candidate=Join-Path $root 'docs/regression/PostAudit03/Axial1000_2026-10-09/FocusedMeritV3/RC_Section_NDM.xlsm'
$equality=Read-Json (Join-Path (Split-Path -Parent $candidate) 'VbaEquality.json')
$load=Read-Json (Join-Path $first 'LoadStage.json')
$tail=Read-Json $tailPath
$now=[datetime]::UtcNow
$mainHash=(Get-FileHash -LiteralPath (Join-Path $root 'workbook/output/RC_Section_NDM.xlsm')).Hash
$expectedMain='17474D47DDBEB05A8A40A74911698E0FA0316F769582552DD90ED327A9CB36EA'
$external=Join-Path $root 'docs/regression/PostAudit03/Axial1000_2026-10-09/FocusedExternalMerit/Results.json'
[ordered]@{
    Status='NotAccepted'; AcceptanceComplete=$false; ExhaustiveSuiteExecuted=$false;
    StartedUtc=$StartedUtc.ToUniversalTime().ToString('o'); DeadlineUtc=$DeadlineUtc.ToUniversalTime().ToString('o');
    FinishedUtc=$now.ToString('o'); ElapsedSeconds=($now-$StartedUtc.ToUniversalTime()).TotalSeconds;
    InitialLoadStage=$load; InitialNativeStage=(Read-Json (Join-Path $first 'NativeCAD/Stage.json'));
    InitialIndependentControl=(Read-Json $external); InitialNumericalReview=(Read-Json (Join-Path $first 'NumericalFailureReview.json'));
    CorrectedCandidateSHA256=(Get-FileHash -LiteralPath $candidate).Hash; CorrectedVbaEquality=$equality;
    CorrectedContractsNo=(Summarize-Report (Join-Path $first 'CorrectedContracts_No.txt'));
    OriginalPlateauControl=(Summarize-Report (Join-Path $first 'OriginalForcePathPlateau_No.txt'));
    CorrectedTail=$tail; PrimaryWorkbookPublished=$false; PrimaryWorkbookSHA256=$mainHash;
    PrimaryWorkbookUnchanged=($mainHash -ceq $expectedMain);
    Remaining=@('Repeat remaining On/Off loads on final code','Recheck strong Off overload timing','Final output/data/presentation preservation before publication');
    Meaning='All evidence belongs to its recorded candidate. Initial positive stages are not full acceptance of later corrections.'
} | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $target -Encoding UTF8
Write-Output ('FOCUSED_RUN_OUTCOME: NotAccepted; mainUnchanged='+($mainHash -ceq $expectedMain))
