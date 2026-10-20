# Directed comparisons; not the historical exhaustive settings matrix.
param(
    [string]$ProbeWorkbook='docs/regression/PlateauSlope_2026-10-11/ProbeV3/RC_Section_NDM.xlsm',
    [string]$ReportDirectory='docs/regression/PlateauSlope_2026-10-11/ResultsV1',
    [ValidateSet('Materials','Saved','Generated','Paths','SlowRepeated','PureAxial','PureRepeated','Batch15','BatchRepeated','Cracked','ForcedAxial','Extended')][string]$Stage='Materials',
    [switch]$SimpleInitialGuess
)
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$probe=(Resolve-Path -LiteralPath (Join-Path $root $ProbeWorkbook)).Path
$production=Join-Path $root 'workbook/output/RC_Section_NDM.xlsm'
$productionHash=(Get-FileHash -LiteralPath $production).Hash
if ($Stage -eq 'Extended') {
    foreach ($nextStage in @('PureAxial','Batch15','Paths','Cracked','SlowRepeated','ForcedAxial')) {
        & $PSCommandPath -ProbeWorkbook $ProbeWorkbook -ReportDirectory $ReportDirectory -Stage $nextStage -SimpleInitialGuess:$SimpleInitialGuess
    }
    return
}
$jobs=[Collections.Generic.List[object]]::new()
if ($Stage -eq 'ForcedAxial') {
    foreach ($mode in @('No','Yes')) {
        $jobs.Add(@{Name="ForcedAxial_${mode}"; Mode=$mode; Macro='modPlateauSlopeProbe.RunPlateauForcedAxial'; Argument=''; Timeout=300})
    }
} elseif ($Stage -eq 'Materials') {
    $jobs.Add(@{Name='Materials'; Mode='No'; Macro='modPlateauSlopeProbe.RunPlateauMaterialChecks'; Argument=''; Timeout=120})
} else {
    foreach ($mode in @('No','Yes')) {
        foreach ($alpha in @('0','0.0001')) {
            $label=if ($alpha -eq '0') {'SlopeOff'} else {'SlopeOn'}
            $simple=if ($SimpleInitialGuess -and $alpha -ne '0') {'1'} else {'0'}
            if ($Stage -eq 'Saved') {
                $jobs.Add(@{Name="Saved_${mode}_${label}"; Mode=$mode; Argument="$alpha|$simple|SavedWorkbook|Saved||Auto|Auto|1"; Timeout=600})
            } elseif ($Stage -eq 'Generated') {
                foreach ($shape in @('Rectangles','Circle','PointedSection','HollowRectangle')) {
                    $jobs.Add(@{Name="${shape}_${mode}_${label}"; Mode=$mode; Argument="$alpha|$simple|$shape|Generated||Auto|Auto|1"; Timeout=300})
                }
            } elseif ($Stage -eq 'Paths') {
                foreach ($method in @('UltimateStrain','LoadMultiplier')) {
                    $jobs.Add(@{Name="Paths_${method}_${mode}_${label}"; Mode=$mode; Argument="$alpha|$simple|SavedWorkbook|Paths||$method|$method|1"; Timeout=600})
                }
            } elseif ($Stage -eq 'SlowRepeated') {
                foreach ($load in @('overloadMixed1','overloadMixed1000000','N2098Plateau')) {
                    $jobs.Add(@{Name="${load}_${mode}_${label}"; Mode=$mode; Argument="$alpha|$simple|SavedWorkbook|Saved|$load|Auto|Auto|3"; Timeout=300})
                }
            } elseif ($Stage -in @('PureAxial','PureRepeated','Batch15','BatchRepeated')) {
                $group=if ($Stage -eq 'PureRepeated') {'PureAxial'} elseif ($Stage -eq 'BatchRepeated') {'Batch15'} else {$Stage}
                $repetitions=if ($Stage.EndsWith('Repeated')) {3} else {1}
                $jobs.Add(@{Name="${Stage}_${mode}_${label}"; Mode=$mode; Argument="$alpha|$simple|SavedWorkbook|$group||Auto|Auto|$repetitions"; Timeout=300})
            } elseif ($Stage -eq 'Cracked') {
                foreach ($shape in @('SavedWorkbook','HollowRectangle')) {
                    $jobs.Add(@{Name="Cracked_${shape}_${mode}_${label}"; Mode=$mode; Argument="$alpha|$simple|$shape|Cracked||Auto|Auto|1"; Timeout=600})
                }
            }
        }
    }
}
$manifest=[Collections.Generic.List[object]]::new()
foreach ($job in $jobs) {
    if ((Get-PSDrive C).Free -lt 256MB) {throw 'Insufficient disk space for isolated test copies.'}
    $macro=if ($job.Macro) {$job.Macro} else {'modTestBatchCalculation.RunPlateauSlopeCases'}
    $report=Join-Path $ReportDirectory ($job.Name+'.txt')
    Write-Output ('PLATEAU_JOB_START: '+$job.Name)
    & (Join-Path $PSScriptRoot 'Run-Audit03Watchdog.ps1') -SourceWorkbook $probe -ReportPath $report -Macro $macro -MacroArgument1 $job.Argument -Mode $job.Mode -TimeoutSeconds $job.Timeout
    if (Select-String -LiteralPath (Join-Path $root $report) -Pattern '^SCRIPT ERROR:' -Quiet) {throw 'Test infrastructure failed; remaining measurements were not started.'}
    $manifest.Add([ordered]@{Name=$job.Name; ExitCode=$LASTEXITCODE; Mode=$job.Mode; Argument=$job.Argument; Report=$report})
    if ((Get-FileHash -LiteralPath $production).Hash -cne $productionHash) { throw 'Production book changed during experiments.' }
}
$manifest | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $root (Join-Path $ReportDirectory ($Stage+'_Manifest.json'))) -Encoding UTF8
Write-Output ('PLATEAU_STAGE_COMPLETED: '+$Stage+'; jobs='+$jobs.Count+'; production unchanged=True')
