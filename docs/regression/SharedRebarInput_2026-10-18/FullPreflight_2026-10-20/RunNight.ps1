# Waits for the verified Config process, then starts the complete acceptance once.
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../../../..')).Path
Set-Location -LiteralPath $root
$evidence='docs/regression/FullOnOff_2026-10-20_Run01'
$deadline='2026-10-21T10:00:00+03:00'
$python=Join-Path $env:USERPROFILE '.cache/codex-runtimes/codex-primary-runtime/dependencies/python/python.exe'
$bookHash='02088CA93F43461192606083C5A160B0716822406AF91C1C1E6F87C03439AEEA'
$configHash='C0C90E0CE773ACBDA37CD95F7BBE2EA15A9F86D52F82E8801E280B0DD66B48C2'
$state=[ordered]@{Status='WAITING_CONFIG';DriverPID=$PID;DriverStartTicks=(Get-Process -Id $PID).StartTime.ToUniversalTime().Ticks;
    Started=[DateTimeOffset]::Now.ToString('o');Deadline=$deadline;EvidenceDirectory=$evidence;SourceSHA256=$bookHash;
    ConfigPID=209872;ConfigStartTicks=639281142568628275;DriverExit=1;SummaryExit=1;CombinedExit=1}
function Save-State {$state | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $PSScriptRoot 'NightDriverResult.json') -Encoding UTF8}
function Read-Log([string]$name){[IO.File]::ReadAllText((Join-Path $PSScriptRoot $name),[Text.Encoding]::UTF8)}
function Assert-PassLog([string]$body,[string]$hash){
    if($body -notmatch ('(?m)^SOURCE_SHA256: '+$hash+'\r?$') -or
        -not $body.Contains('WATCHDOG_COMPLETED: exit=0; source unchanged=True') -or
        $body -match '(?m)^(FAIL:|RUNTIME ERROR|SCRIPT ERROR|WATCHDOG_FAILURE|RESTORE ERROR|CLEANUP ERROR)'){
        throw 'A required directed log failed or used a different workbook.'
    }
}
Save-State
try {
    if(Test-Path -LiteralPath $evidence){throw 'The full-run directory already exists; automatic restart is forbidden.'}
    $config=Get-Process -Id $state.ConfigPID -ErrorAction SilentlyContinue
    if($config){
        if($config.ProcessName -ne 'powershell' -or $config.StartTime.ToUniversalTime().Ticks -ne $state.ConfigStartTicks){throw 'Config PID identity changed.'}
        $null=$config.Handle
        while(-not $config.WaitForExit(30000)){$config.Refresh()}
        $config.WaitForExit()
    }
    $body=Read-Log 'Revision02/SharedConfig.txt'
    Assert-PassLog $body $configHash
    if($body -notmatch '(?m)^TOTAL_SHARED_CONFIG_PREFLIGHT: passed=([1-9][0-9]*); failed=0\r?$'){throw 'Config totals are incomplete.'}
    $configPassed=[int]$Matches[1]
    if(([regex]::Matches($body,'(?m)^SUITE_STARTED: ')).Count -ne 1 -or
        $body -notmatch '(?m)^SUITE_FINISHED: modTestWorkbookInterface.RunSharedConfigPreflightTests;'){throw 'Config did not complete.'}
    $context=Get-Content -LiteralPath (Join-Path $PSScriptRoot 'SharedContextEquality.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    if(-not $context.Passed -or $context.SharedBookSHA256 -ne $configHash -or $context.PublishedBookSHA256 -ne $bookHash -or
        -not $context.ProductionCodeUnchanged -or -not $context.SharedEntryAndGeometryConfigModulesEquivalent){throw 'Config code provenance is incomplete.'}
    $body=Read-Log 'ReleaseFinal/CorrectedExpectations.txt'
    Assert-PassLog $body $bookHash
    $expected=@('modTestGeometryReview.RunGeometryReviewTests','modTestRecipeOperations.RunSavedPointedOffsetTests','modTestRecipeOperations.RunSavedOffsetCalculationTests')
    $actual=@([regex]::Matches($body,'(?m)^SUITE_FINISHED: ([^;]+);') | ForEach-Object {$_.Groups[1].Value})
    if(($actual -join '|') -cne ($expected -join '|')){throw 'Corrected checks are incomplete.'}
    $body=Read-Log 'Expectations.txt'
    $plan=Get-Content -LiteralPath (Join-Path $PSScriptRoot 'ExpectationPlan.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    $actual=@([regex]::Matches($body,'(?m)^SUITE_FINISHED: ([^;]+);') | ForEach-Object {$_.Groups[1].Value})
    $failures=@([regex]::Matches($body,'(?m)^FAIL:[^\r\n]*') | ForEach-Object {$_.Value})
    if($actual.Count -ne 58 -or ($actual -join '|') -cne ($plan.Macros -join '|') -or
        $failures.Count -ne 1 -or $failures[0] -cne 'FAIL: collapsed.extra.row' -or
        -not $body.Contains('WATCHDOG_COMPLETED: exit=1; source unchanged=True') -or
        $body -match '(?m)^(RUNTIME ERROR|SCRIPT ERROR|WATCHDOG_FAILURE|RESTORE ERROR|CLEANUP ERROR)'){throw 'The original 58-group preflight has unreviewed failures.'}
    if((Get-FileHash -LiteralPath 'workbook/output/RC_Section_NDM.xlsm').Hash -cne $bookHash -or
        (Get-FileHash -LiteralPath 'workbook/output/VBA_All_Code.txt').Hash -cne '7503AAE16CEAF73362FFD3D38256741CB87269CE32301B472CE31D212398A044'){
        throw 'Published workbook or ALL CODE changed.'
    }
    $equality=Get-Content -LiteralPath (Join-Path $PSScriptRoot 'PublishedSourceEquality.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    if(-not $equality.Passed){throw 'Published source equality failed.'}
    $state['ConfigPassedAssertions']=$configPassed
    $state['PreflightFinished']=[DateTimeOffset]::Now.ToString('o')
    $summary=Get-Content -LiteralPath (Join-Path $PSScriptRoot 'PreflightSummary.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    $summary.Status='PASSED_AFTER_TARGETED_CORRECTIONS'
    $summary | Add-Member -NotePropertyName ConfigPassedAssertions -NotePropertyValue $configPassed
    $summary | Add-Member -NotePropertyName ConfigCodeProvenance -NotePropertyValue 'SharedContextEquality.json'
    $summary | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath (Join-Path $PSScriptRoot 'PreflightSummary.json') -Encoding UTF8
    $state.Status='FULL_STARTING';Save-State
    Write-Output ('FULL_STARTING: '+$evidence+'; '+[DateTimeOffset]::Now.ToString('o'))
    & (Join-Path $PSHOME 'powershell.exe') -NoProfile -ExecutionPolicy Bypass -File tools/build_workbook/Run-FullOnOffAcceptance.ps1 -SourceWorkbook workbook/output/RC_Section_NDM.xlsm -EvidenceDirectory $evidence -Deadline $deadline -ForecastHours 6 -Python $python
    $state.DriverExit=$LASTEXITCODE
    if(Test-Path -LiteralPath ($evidence+'/Run.json')){
        & $python tools/build_workbook/Summarize-FullOnOffAcceptance.py --directory $evidence --check
        $state.SummaryExit=$LASTEXITCODE
        if($state.DriverExit -eq 0 -and $state.SummaryExit -eq 0){
            & $python tools/build_workbook/Finalize-CompleteAcceptance.py --core-directory $evidence --focused-directory ($evidence+'/Focused')
            $state.CombinedExit=$LASTEXITCODE
        }
    }
    $state.Status=if($state.DriverExit -eq 0 -and $state.SummaryExit -eq 0 -and $state.CombinedExit -eq 0){'PASS'}else{'REVIEW_REQUIRED'}
} catch {
    $state.Status='REVIEW_REQUIRED';$state['Error']=$_.Exception.Message
    Write-Output ('NIGHT_DRIVER_ERROR: '+$_.Exception.Message)
} finally {
    $state['Finished']=[DateTimeOffset]::Now.ToString('o');Save-State
    Write-Output ('NIGHT_DRIVER_STATUS: '+$state.Status)
}
if($state.Status -ne 'PASS'){exit 1}
