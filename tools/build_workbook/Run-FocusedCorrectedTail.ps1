# Addressed repeats only; the original failed load evidence remains authoritative.
param([Parameter(Mandatory=$true)][datetime]$DeadlineUtc)
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$output='docs/regression/FocusedAcceptance_2026-10-10_V9'
$candidate='docs/regression/PostAudit03/Axial1000_2026-10-09/FocusedMeritV3/RC_Section_NDM.xlsm'
$batch='docs/regression/PostAudit03/Axial1000_2026-10-09/FocusedBatch15MeritV3/RC_Section_NDM.xlsm'
$probe=$output+'/RC_Section_NDM.xlsm'
$sourceHash=(Get-FileHash -LiteralPath (Join-Path $root $candidate)).Hash
$jobs=@(
    @{Name='Batch15_No'; Source=$batch; Macro='RunFocusedBatch15'; Mode='No'; A=''; B=''},
    @{Name='Batch15_Yes'; Source=$batch; Macro='RunFocusedBatch15'; Mode='Yes'; A=''; B=''},
    @{Name='Saved_Yes'; Source=$probe; Macro='RunFocusedAcceptanceCases'; Mode='Yes'; A='SavedWorkbook'; B='Saved'},
    @{Name='RoundedMx1_Yes'; Source=$probe; Macro='RunFocusedSingleLoad'; Mode='Yes'; A='RoundedRectangle'; B='Mx1'}
)
$records=@(); $reason=''
try {
    foreach($job in $jobs) {
        $remaining=[math]::Floor(($DeadlineUtc.ToUniversalTime()-[datetime]::UtcNow).TotalSeconds)
        if($remaining -lt 15) { $reason='Incomplete: approved overall deadline reached.'; break }
        $relative=$output+'/'+$job.Name+'.txt'
        $path=Join-Path $root $relative
        if(Test-Path -LiteralPath $path) { throw 'Existing tail evidence cannot be overwritten.' }
        Write-Output ('CORRECTED_TAIL_STARTED: '+$job.Name+'; remaining='+$remaining)
        & (Join-Path $PSScriptRoot 'Run-Audit03Watchdog.ps1') -SourceWorkbook $job.Source -ReportPath $relative `
            -Macro ('modTestBatchCalculation.'+$job.Macro) -MacroArgument1 $job.A -MacroArgument2 $job.B `
            -Mode $job.Mode -TimeoutSeconds ([math]::Min(150,$remaining)) -VerifyResultsReopen -VerifyStatusReopen
        $body=Get-Content -LiteralPath $path -Raw -Encoding UTF8
        $records+=@{Name=$job.Name; ExitCode=$LASTEXITCODE; Report=$relative; SHA256=(Get-FileHash -LiteralPath $path).Hash;
            Passed=($LASTEXITCODE -eq 0 -and $body -notmatch '(?m)^FAIL:|RUNTIME ERROR|WATCHDOG_FAILURE')}
        $records | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $root ($output+'/TailProgress.json')) -Encoding UTF8
    }
} catch { $reason=$_.Exception.Message }
finally {
    [ordered]@{SourceSHA256=$sourceHash; Passed=($records.Count -eq $jobs.Count -and @($records | Where-Object {-not $_.Passed}).Count -eq 0);
        AcceptanceComplete=$false; Records=$records; StopReason=$reason; DeadlineUtc=$DeadlineUtc.ToUniversalTime().ToString('o');
        CandidateUnchanged=((Get-FileHash -LiteralPath (Join-Path $root $candidate)).Hash -ceq $sourceHash)} |
        ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $root ($output+'/TailStage.json')) -Encoding UTF8
}
