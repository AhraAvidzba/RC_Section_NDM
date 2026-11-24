# ДЛЯ ТЕСТОВ: сравнивает нормативные классы на существующих адресных наборах.
# Обе книги должны быть подготовлены заранее; исходники и рабочая книга не меняются.
param(
    [Parameter(Mandatory=$true)][string]$BaselineWorkbook,
    [Parameter(Mandatory=$true)][string]$CandidateWorkbook,
    [Parameter(Mandatory=$true)][string]$EvidenceDirectory,
    [ValidateSet('Baseline','Candidate')][string[]]$Labels=@('Baseline','Candidate'),
    [ValidateSet('No','Yes')][string[]]$Modes=@('No','Yes')
)
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$directory=Join-Path $root $EvidenceDirectory
New-Item -ItemType Directory -Path $directory -Force | Out-Null
$macros=@(
    'modTestCrackWidth.RunCrackWidthTests',
    'modTestCrackWidth.RunAudit03SigmaSCrcBoundaryTests',
    'modTestCrackWidth.RunNoFormationTensionWidthTests',
    'modTestCrackWidth.RunCrackGeometryTests',
    'modTestCrackWidth.RunCrackTranslationTests',
    'modTestSP35CrackWidth.RunSP35FormulaTests',
    'modTestSP35CrackWidth.RunSP35PreparationTests',
    'modTestSP35CrackWidth.RunSP35NeighborSettingsTests',
    'modTestSP35CrackWidth.RunSP35LocalOpeningTests',
    'modTestSP35CrackWidth.RunSP35PipelineTests',
    'modTestSP35CrackWidth.RunSP35EndToEndTests',
    'modTestBatchCalculation.RunNormativeStabilityTests',
    'modTestBatchCalculation.RunAudit03SP35TableBehaviorTests',
    'modTestBatchCalculation.RunAudit03SP35TableDiagnosticTests',
    'modTestBatchCalculation.RunLocaleSP35InterpolationTests',
    'modTestBatchCalculation.RunAudit03CrackSpacingWarningTests',
    'modTestConfiguration.RunAudit03StabilityConfigTests',
    'modTestConfiguration.RunAudit03StabilityLengthRangeTests',
    'modTestProfileConfig.RunStabilityDurationMatchingTests'
)
$runs=@()
foreach($label in $Labels) {
    $source=if($label -eq 'Baseline'){$BaselineWorkbook}else{$CandidateWorkbook}
    foreach($mode in $Modes) {
        $relative=Join-Path $EvidenceDirectory ($label+'_'+$mode+'.txt')
        Write-Output ('START '+$label+' '+$mode+'; '+[DateTimeOffset]::Now.ToString('o'))
        # Watchdog перед каждым Excel проверяет RAM/commit и записывает точную идентичность собственного PID.
        & (Join-Path $PSScriptRoot 'Run-Audit03Watchdog.ps1') -SourceWorkbook $source -ReportPath $relative -Mode $mode -Macro ($macros -join ',') -TimeoutSeconds 3600 -FailFast -NormativeCoreGuard
        if($LASTEXITCODE -ne 0){throw ('Native test process failed: '+$label+' '+$mode)}
        $report=Join-Path $root $relative
        $text=[IO.File]::ReadAllText($report)
        if($text -match '(?m)^FAIL[: ]|failed=[1-9]|RUNTIME ERROR|WATCHDOG_FAILURE'){throw ('Failed test report: '+$report)}
        $finished=[regex]::Matches($text,'(?m)^SUITE_FINISHED:').Count
        if($finished -ne $macros.Count){throw ('Incomplete directed set: '+$label+' '+$mode)}
        $runs+=,[ordered]@{Label=$label;Mode=$mode;WorkbookSHA256=(Get-FileHash -LiteralPath $source).Hash;Suites=$finished;Report=$relative;Passed=$true}
        $runs | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directory 'DirectedRuns.json') -Encoding UTF8
        Write-Output ('PASS '+$label+' '+$mode+'; suites='+$finished)
    }
}
