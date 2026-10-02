# Присоединяет вручную проверенную per-key трассировку к фактическому census.
# Существование строки или совпадение имени теста не является PASS: обязательны
# завершенный положительный лог и конкретные assertions активного поведения.
param(
    [string]$RegistryPath = 'docs/regression/Audit03/config_field_registry_2026-10-02.json',
    [string[]]$EvidencePath = @('docs/NDM_Audit03_Config_Behavior_Evidence.json'),
    [string]$OutputPrefix = 'docs/regression/Audit03/config_behavior_registry_v48_2026-10-02'
)
$ErrorActionPreference = 'Stop'
$root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$registry = Get-Content -LiteralPath (Join-Path $root $RegistryPath) -Raw -Encoding UTF8 | ConvertFrom-Json
$reviewed = @{}
foreach ($evidenceFile in $EvidencePath) {
$evidence = Get-Content -LiteralPath (Join-Path $root $evidenceFile) -Raw -Encoding UTF8 | ConvertFrom-Json
$positive = Get-Content -LiteralPath (Join-Path $root $evidence.PositiveReport) -Raw -Encoding UTF8
if (-not $positive.Contains($evidence.PositiveGate) -or
    -not $positive.Contains('WATCHDOG_COMPLETED: exit=0; source unchanged=True')) {
    throw 'Положительное directed evidence не завершено или не совпадает с gate.'
}
$fullReports = foreach ($path in $evidence.FullReports) {
    $report = Get-Content -LiteralPath (Join-Path $root $path) -Raw -Encoding UTF8
    if (-not $report.Contains('WATCHDOG_COMPLETED: exit=0; source unchanged=True')) {
        throw "Полный регрессионный прогон не завершен: $path"
    }
    $requiredSuites = @('modTestGeometry.RunGeometryTests', 'modTestMaterialDiagrams.RunMaterialDiagramTests',
        'modTestSectionSolver.RunSectionSolverTests', 'modTestCapacitySolver.RunCapacitySolverTests',
        'modTestCrackWidth.RunCrackWidthTests', 'modTestBatchCalculation.RunBatchCalculationTests',
        'modTestWorkbookInterface.RunWorkbookInterfaceTests', 'modTestRegressionBaseline.RunRegressionBaselineTests')
    foreach ($suite in $requiredSuites) {
        if (-not $report.Contains("SUITE_FINISHED: $suite;")) { throw "В full evidence отсутствует suite $suite : $path" }
    }
    if ($report -match '(?m)^TOTAL[^\r\n]*failed=[1-9]') { throw "В full evidence есть ошибки: $path" }
    $report
}
foreach ($entry in $evidence.Entries) {
    if ($reviewed.ContainsKey($entry.Id)) { throw "Повторный key evidence: $($entry.Id)" }
    $fields = @($registry.Fields | Where-Object Id -EQ $entry.Id)
    if ($fields.Count -ne 1) { throw "Нужна одна фактическая строка census для $($entry.Id)" }
    $pattern = '(?m)^OK: ' + [regex]::Escape($entry.AssertionPrefix)
    if ($positive -notmatch $pattern) { throw "Нет активного assertion для $($entry.Id)" }
    foreach ($report in $fullReports) {
        if ($report -notmatch $pattern) { throw "Full suite не содержит active evidence для $($entry.Id)" }
    }
    $field = $fields[0]
    $field.Type = $entry.Type
    $field.InternalUnits = $entry.InternalUnits
    $field.RuntimeDefault = $entry.RuntimeDefault
    $field.Activity = $entry.Activity
    $field.Consumers = $evidence.Consumer
    $field.ExpectedEffect = $entry.ExpectedEffect
    $field.TestId = $evidence.TestId
    $field.Evidence = @($evidence.PositiveReport) + @($evidence.FullReports)
    $field.CoverageStatus = 'ActiveBehaviorAccepted:FullRangeReviewNotComplete'
    $field | Add-Member -NotePropertyName ValidRange -NotePropertyValue $entry.ValidRange -Force
    $field | Add-Member -NotePropertyName RangeReview -NotePropertyValue $entry.RangeReview -Force
    if ($entry.CanonicalKey) { $field | Add-Member -NotePropertyName CanonicalKey -NotePropertyValue $entry.CanonicalKey -Force }
    if ($entry.BlankAndErrorContract) { $field.BlankAndErrorContract = $entry.BlankAndErrorContract }
    if ($entry.MutationReviewed) {
        $mutation = Get-Content -LiteralPath (Join-Path $root $evidence.MutationReport) -Raw -Encoding UTF8
        if ($mutation -notmatch '(?m)^FAIL: audit03.effect.Solver.Method.Secant' -or
            -not $mutation.Contains('WATCHDOG_COMPLETED: exit=1; source unchanged=True')) {
            throw 'Отрицательная Method-mutation не обнаружена соответствующим тестом.'
        }
        $field | Add-Member -NotePropertyName MutationEvidence -NotePropertyValue $evidence.MutationReport -Force
    }
    $reviewed[$entry.Id] = $true
}
}
$acceptedCount = @($registry.Fields | Where-Object CoverageStatus -Like 'ActiveBehaviorAccepted:*').Count
$registry | Add-Member -NotePropertyName ActiveBehaviorReviewedFields -NotePropertyValue $acceptedCount -Force
$prefix = [IO.Path]::GetFullPath((Join-Path $root $OutputPrefix))
$registry | ConvertTo-Json -Depth 50 | Set-Content -LiteralPath ($prefix + '.json') -Encoding UTF8
$registry.Fields | Select-Object Id, CanonicalKey, Block, Address, Role, Type, UserUnits, InternalUnits,
    NewWorkbookDefault, RuntimeDefault, Activity, Consumers, ExpectedEffect, TestId,
    @{N='Evidence'; E={ $_.Evidence -join '; ' }}, CoverageStatus, ValidRange, RangeReview,
    MutationEvidence | Export-Csv -LiteralPath ($prefix + '.csv') -NoTypeInformation -Encoding UTF8
Write-Output "CONFIG_EVIDENCE: fields=$($registry.Fields.Count); activeReviewed=$acceptedCount; fullAcceptance=False"
