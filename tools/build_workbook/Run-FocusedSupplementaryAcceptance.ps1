# Selected contracts, full special cases and one 15-load batch per mode.
param([string]$CandidateDirectory='docs/regression/PostAudit03/Axial1000_2026-10-09/FocusedMeritV1',
      [string]$ProbeDirectory='docs/regression/FocusedAcceptance_2026-10-10_V5',
      [string]$LoadEvidenceDirectory='',
      [Parameter(Mandatory=$true)][datetime]$DeadlineUtc)
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$candidate=Join-Path (Join-Path $root $CandidateDirectory) 'RC_Section_NDM.xlsm'
$probe=Join-Path (Join-Path $root $ProbeDirectory) 'RC_Section_NDM.xlsm'
$directory=Join-Path (Join-Path $root $ProbeDirectory) 'Supplementary'
if (Test-Path -LiteralPath $directory) { throw 'Supplementary evidence already exists.' }
New-Item -ItemType Directory -Path $directory | Out-Null
if (-not $LoadEvidenceDirectory) { $LoadEvidenceDirectory=$ProbeDirectory }
$load=Get-Content -LiteralPath (Join-Path (Join-Path $root $LoadEvidenceDirectory) 'LoadStage.json') -Raw | ConvertFrom-Json
if (-not $load.LoadStageComplete -or -not $load.SourceUnchanged -or -not $load.ProductionUnchanged) {
    throw 'Selected load stage is incomplete; supplementary acceptance refused.'
}
$sourceHash=(Get-FileHash -LiteralPath $candidate).Hash
$prep=Get-Content -LiteralPath (Join-Path (Join-Path $root $ProbeDirectory) 'Preparation.json') -Raw | ConvertFrom-Json
if ($prep.SourceSHA256 -cne $sourceHash -or $load.SourceSHA256 -cne $sourceHash -or $prep.ProbeWorkbookSHA256 -cne (Get-FileHash -LiteralPath $probe).Hash) {
    throw 'Candidate/probe provenance changed.'
}
$watch=[Diagnostics.Stopwatch]::StartNew()
$records=@(); $passed=$false; $reason=''
function Remaining-Seconds {
    $remaining=[math]::Floor(($DeadlineUtc.ToUniversalTime()-[datetime]::UtcNow).TotalSeconds)
    if ($remaining -lt 15) { throw 'Incomplete: overall one-hour deadline reached.' }
    return [int]$remaining
}
function Invoke-SelectedGate([string]$Source,[string]$Name,[string]$Mode,[string]$Argument='', [switch]$Reopen) {
    $timeout=[math]::Min(180,(Remaining-Seconds))
    $relative=$ProbeDirectory+'/Supplementary/'+(($Name -split '\.')[-1])+'_'+$Argument+'_'+$Mode+'.txt'
    $report=Join-Path $root $relative
    Write-Output ('SELECTED_STARTED: '+$Name+'/'+$Argument+'/'+$Mode)
    & (Join-Path $PSScriptRoot 'Run-Audit03Watchdog.ps1') -SourceWorkbook $Source -ReportPath $relative `
        -Macro $Name -MacroArgument1 $Argument -Mode $Mode -TimeoutSeconds $timeout `
        -VerifyResultsReopen:$Reopen -VerifyStatusReopen:$Reopen
    $code=$LASTEXITCODE
    $body=Get-Content -LiteralPath $report -Raw -Encoding UTF8
    $totals=[regex]::Matches($body,'(?m)^TOTAL[^\r\n]*passed=(\d+);\s*failed=(\d+)')
    if ($code -ne 0 -or $totals.Count -eq 0 -or
        $body -match 'RUNTIME ERROR|SCRIPT ERROR|WATCHDOG_FAILURE|RESTORE ERROR|(?m)^FAIL:' -or
        @($totals | Where-Object {$_.Groups[2].Value -ne '0'}).Count) {
        throw ('Selected gate failed: '+$Name+'/'+$Argument+'/'+$Mode)
    }
    if ($Name -like '*RunFocusedBatch15') {
        $timing=[regex]::Match($body,'CRACK_SEARCH_PROBE\|case=batch15\|mode=[^|]+\|count=15\|elapsedSec=([^|\r\n]+)')
        if (-not $timing.Success -or [double]::Parse($timing.Groups[1].Value,[Globalization.CultureInfo]::InvariantCulture) -gt 10) {
            throw 'Batch15 did not complete within 10 seconds.'
        }
    }
    $script:records+=@{Macro=$Name; Argument=$Argument; Mode=$Mode; Passed=[int]$totals[$totals.Count-1].Groups[1].Value;
        Failed=0; Report=$relative; ReportSHA256=(Get-FileHash -LiteralPath $report).Hash; SourceSHA256=(Get-FileHash -LiteralPath $Source).Hash}
    $script:records | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $directory 'Progress.json') -Encoding UTF8
}
try {
    foreach ($mode in @('No','Yes')) {
        foreach ($case in @('Stability','AsymmetricSteel')) {
            Invoke-SelectedGate $probe 'modTestBatchCalculation.RunFocusedSpecialCase' $mode $case -Reopen
        }
    }
    foreach ($mode in @('No','Yes')) {
        $macros=if ($mode -eq 'No') {
            @('RunLimitGuessScalingTests','RunConcreteMeshCenterTests','RunCenteredNumericalTests','RunCommonStrainLimitTests',
              'RunAsymmetricSteelStrengthTests','RunAxialCapacityRegressionTests','RunForcePathPlateauTests',
              'RunAudit03CapacityContracts','RunAudit03MultiplierTypedFaults','RunAudit03CapacityLifecycle',
              'RunAudit03FormationContracts','RunAudit03FormationLifecycle','RunFormationOutcomeTests',
              'RunAudit03ExtremeFormationTests','RunAudit03UnitSignConsumerTests','RunConfigPresentationTests','RunCapacityOutputContractTests')
        } else {
            @('RunCapacityOutputContractTests','RunAudit03FormationContracts','RunAudit03ExtremeFormationTests',
              'RunFormationOutcomeTests','RunAsymmetricSteelStrengthTests')
        }
        foreach ($macro in $macros) { Invoke-SelectedGate $candidate $macro $mode }
    }
    [void](Remaining-Seconds)
    $series=Join-Path $root 'docs/regression/PostAudit03/Axial1000_2026-10-09'
    $probeRelative=[IO.Path]::GetFullPath((Join-Path $root $ProbeDirectory))
    $batchDirectory='FocusedBatch15Merit'
    if (Test-Path -LiteralPath (Join-Path $series $batchDirectory)) { throw 'Batch evidence already exists.' }
    # The batch copy retains only the test capture from the private probe.
    $sourceRelative='../../'+(Split-Path -Leaf $probeRelative)
    & (Join-Path $series 'PrepareSlowCaseProbe.ps1') -SourceDirectory $sourceRelative -OutputDirectory $batchDirectory `
        -HelperFile SavedCrackSearchProbe.bas -SecondHelperFile '../../../../tools/build_workbook/FocusedBatch15Probe.bas'
    $batch=Join-Path $series ($batchDirectory+'/RC_Section_NDM.xlsm')
    foreach ($mode in @('No','Yes')) { Invoke-SelectedGate $batch 'modTestBatchCalculation.RunFocusedBatch15' $mode -Reopen }
    $passed=$true
} catch { $reason=$_.Exception.Message; Write-Output ('SELECTED_FAILURE: '+$reason) }
finally {
    $candidateUnchanged=(Get-FileHash -LiteralPath $candidate).Hash -ceq $sourceHash
    if (-not $candidateUnchanged) { $passed=$false; $reason='Candidate changed during selected checks.' }
    [ordered]@{Passed=$passed; AcceptanceComplete=$false; SourceSHA256=$sourceHash; ElapsedSeconds=$watch.Elapsed.TotalSeconds;
        DeadlineUtc=$DeadlineUtc.ToUniversalTime().ToString('o'); Records=$records; StopReason=$reason;
        CandidateUnchanged=$candidateUnchanged; ExhaustiveSuiteExecuted=$false} |
        ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $directory 'Stage.json') -Encoding UTF8
}
Write-Output ('SUPPLEMENTARY_STAGE: passed='+$passed+'; completedGates='+$records.Count)
if (-not $passed) { exit 1 }
