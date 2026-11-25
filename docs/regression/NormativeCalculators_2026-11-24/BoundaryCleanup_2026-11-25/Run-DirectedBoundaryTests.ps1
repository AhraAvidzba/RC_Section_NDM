# ДЛЯ ТЕСТОВ: адресная приёмка границ обязанностей на двух версиях книги.
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../../../..')).Path
$relative='docs/regression/NormativeCalculators_2026-11-24/BoundaryCleanup_2026-11-25'
$macros=@(
 'modTestCrackWidth.RunCrackWidthTests',
 'modTestCrackWidth.RunAudit03SigmaSCrcBoundaryTests',
 'modTestCrackWidth.RunNoFormationTensionWidthTests',
 'modTestSP35CrackWidth.RunSP35FormulaTests',
 'modTestSP35CrackWidth.RunSP35PreparationTests',
 'modTestSP35CrackWidth.RunSP35PipelineTests',
 'modTestBatchCalculation.RunNormativeStabilityTests',
 'modTestProfileConfig.RunStabilityDurationMatchingTests'
)
$runs=@()
foreach($label in @('Baseline','Candidate')){
 foreach($mode in @('No','Yes')){
  $source=Join-Path $PSScriptRoot ($label+'.xlsm')
  $report=$relative+'/'+$label+'_'+$mode+'.txt'
  if(Test-Path -LiteralPath (Join-Path $root $report)){throw 'A fresh test report is required.'}
  Write-Output ('START '+$label+' '+$mode+'; '+[DateTimeOffset]::Now.ToString('o'))
  & (Join-Path $root 'tools/build_workbook/Run-Audit03Watchdog.ps1') -SourceWorkbook $source -ReportPath $report -Mode $mode -Macro ($macros -join ',') -TimeoutSeconds 1200 -NormativeCoreGuard
  if($LASTEXITCODE -ne 0){throw ('Native tests failed: '+$label+' '+$mode)}
  $body=[IO.File]::ReadAllText((Join-Path $root $report))
  if($body -match '(?m)^FAIL[: ]|failed=[1-9]|RUNTIME ERROR|WATCHDOG_FAILURE'){throw ('Failed report: '+$report)}
  if([regex]::Matches($body,'(?m)^SUITE_FINISHED:').Count -ne $macros.Count){throw 'Incomplete native test set.'}
  if([regex]::Matches($body,'(?m)^NORMATIVE_CORE_GUARD_RESTORED: True').Count -ne $macros.Count){throw 'Guard restoration differs.'}
  $runs+=,[ordered]@{Label=$label;Mode=$mode;SourceSHA256=(Get-FileHash -LiteralPath $source).Hash;Report=$report;Macros=$macros;Passed=$true}
  $runs|ConvertTo-Json -Depth 5|Set-Content -LiteralPath (Join-Path $PSScriptRoot 'DirectedRuns.json') -Encoding UTF8
 }
}
$extra=@('modTestCrackWidth.RunNormativeBoundaryTests','modTestBatchCalculation.RunNormativePartialBranchFailureTest')
foreach($mode in @('No','Yes')){
 $report=$relative+'/Boundary_'+$mode+'.txt'
 & (Join-Path $root 'tools/build_workbook/Run-Audit03Watchdog.ps1') -SourceWorkbook (Join-Path $PSScriptRoot 'Candidate.xlsm') -ReportPath $report -Mode $mode -Macro ($extra -join ',') -TimeoutSeconds 300 -NormativeCoreGuard
 if($LASTEXITCODE -ne 0){throw ('Boundary tests failed: '+$mode)}
 $body=[IO.File]::ReadAllText((Join-Path $root $report))
 if($body -notmatch '(?m)^TOTAL_NORMATIVE_BOUNDARY: passed=13; failed=0\r?$' -or $body -notmatch '(?m)^TOTAL_NORMATIVE_PARTIAL_BRANCH: passed=11; failed=0\r?$'){throw 'Boundary totals differ.'}
 $runs+=,[ordered]@{Label='Boundary';Mode=$mode;SourceSHA256=(Get-FileHash -LiteralPath (Join-Path $PSScriptRoot 'Candidate.xlsm')).Hash;Report=$report;Macros=$extra;Passed=$true}
 $runs|ConvertTo-Json -Depth 5|Set-Content -LiteralPath (Join-Path $PSScriptRoot 'DirectedRuns.json') -Encoding UTF8
}
Write-Output 'DIRECTED_BOUNDARY_COMPLETE: 32 comparison executions and 4 boundary executions; inspect exact report equality.'
