# Присоединяет адресные доказательства для 104 служебных контрольных ячеек.
# Это производные формулы/нулевые точки, а не 104 пользовательские настройки.
# Не исключает остальные поля из знаменателя и не объявляет общий Config PASS.
param(
    [Parameter(Mandatory=$true)][string]$RegistryPath,
    [Parameter(Mandatory=$true)][string]$PositiveReport,
    [Parameter(Mandatory=$true)][string[]]$FullReports,
    [Parameter(Mandatory=$true)][string]$OutputPrefix
)
$ErrorActionPreference = 'Stop'
$root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$allowed = [IO.Path]::GetFullPath((Join-Path $root 'docs/regression/Audit03')) + [IO.Path]::DirectorySeparatorChar
$prefix = [IO.Path]::GetFullPath((Join-Path $root $OutputPrefix))
if (-not $prefix.StartsWith($allowed, [StringComparison]::OrdinalIgnoreCase)) {
    throw 'Evidence output must stay in docs/regression/Audit03.'
}
foreach ($suffix in @('.json', '.csv')) {
    if (Test-Path -LiteralPath ($prefix + $suffix)) { throw 'Evidence already exists; use a new output prefix.' }
}
$registry = Get-Content -LiteralPath (Join-Path $root $RegistryPath) -Raw -Encoding UTF8 | ConvertFrom-Json
$directed = Get-Content -LiteralPath (Join-Path $root $PositiveReport) -Raw -Encoding UTF8
if (-not $directed.Contains('TOTAL_MATERIAL_CONTROL_UNITS: passed=1096; failed=0') -or
        -not $directed.Contains('WATCHDOG_COMPLETED: exit=0; source unchanged=True')) {
    throw 'Directed control-table gate is incomplete.'
}
$reports = @($directed)
foreach ($path in $FullReports) {
    $report = Get-Content -LiteralPath (Join-Path $root $path) -Raw -Encoding UTF8
    if (-not $report.Contains('WATCHDOG_COMPLETED: exit=0; source unchanged=True') -or
            $report -match '(?m)^TOTAL[^\r\n]*failed=[1-9]') {
        throw "Full regression gate is incomplete: $path"
    }
    foreach ($suite in @('Geometry', 'MaterialDiagrams', 'SectionSolver', 'CapacitySolver',
            'CrackWidth', 'BatchCalculation', 'WorkbookInterface', 'RegressionBaseline')) {
        if ($report -notmatch ('(?m)^SUITE_FINISHED: modTest' + [regex]::Escape($suite) + '\.')) {
            throw "Missing full suite $suite in $path"
        }
    }
    $reports += $report
}
if ($FullReports.Count -lt 1) { throw 'At least one completed full regression gate is required.' }
foreach ($report in $reports) {
    foreach ($unit in @('MPa', 'kPa', 'Pa', 'kgf/cm2', 'tf/m2')) {
        for ($point = 0; $point -lt 52; $point++) {
            foreach ($quantity in @('strain', 'stress')) {
                $assertion = 'audit03.materialControls.' + $unit + '.' + $quantity + '.' + $point
                if ($report -notmatch ('(?m)^OK: ' + [regex]::Escape($assertion) + '; actual=')) {
                    throw "No physical-provider point comparison: $assertion"
                }
            }
        }
    }
}
$fields = @($registry.Fields | Where-Object Block -EQ 'UnnamedMaterialDiagramControl')
if ($fields.Count -ne 104) { throw 'The inventory must retain all 104 control cells.' }
foreach ($field in $fields) {
    if ($field.Role -notin @('DerivedFormula', 'ReadOnlyControlValue')) { throw 'Unexpected control-cell role.' }
    $cell = ([string]$field.Address -split '!')[-1]
    $kind = if ($field.Role -eq 'DerivedFormula') { 'directFormula' } else { 'zeroLiteral' }
    foreach ($report in $reports) {
        foreach ($unit in @('MPa', 'kPa', 'Pa', 'kgf/cm2', 'tf/m2')) {
            $assertion = 'audit03.materialControls.' + $unit + '.' + $kind + '.' + $cell
            if ($report -notmatch ('(?m)^OK: ' + [regex]::Escape($assertion) + '(?:;|\r?$)')) {
                throw "No exact control-cell assertion: $assertion"
            }
        }
        if (-not $report.Contains('OK: audit03.materialControls.noSolve')) {
            throw 'The control-table test must prove that it does not solve equilibrium.'
        }
    }
    $field.Type = if ($field.Role -eq 'DerivedFormula') { 'ReadOnly arithmetic formula' } else { 'ReadOnly zero point' }
    $field.RuntimeDefault = 'No independent runtime setting; control-only value'
    $field.ExpectedEffect = 'Five INPUT Stress units; 52 physical-provider points, direct arithmetic links or literal zero, unit-aware table/charts; no state solve'
    $field.TestId = 'modTestMaterialDiagrams.TestAudit03MaterialControlUnits'
    $field.Evidence = @($PositiveReport) + $FullReports
    $field.CoverageStatus = 'DerivedControlAccepted:NotIndependentUserInput'
    $field.InternalUnits = 'Dimensionless strain or MPa after CUnitSystem conversion'
    $field.UserUnits = 'Dimensionless strain or current INPUT Stress unit'
    $field.BlankAndErrorContract = 'Generated control cell; do not edit independently. Formula/literal contract checked at the actual cell address.'
}
$registry | Add-Member -NotePropertyName DerivedControlReviewedFields -NotePropertyValue $fields.Count -Force
$registry | ConvertTo-Json -Depth 50 | Set-Content -LiteralPath ($prefix + '.json') -Encoding UTF8
$registry.Fields | Select-Object Id, Block, Address, Role, Type, UserUnits, InternalUnits,
    RuntimeDefault, ExpectedEffect, TestId, @{N='Evidence'; E={ $_.Evidence -join '; ' }},
    CoverageStatus | Export-Csv -LiteralPath ($prefix + '.csv') -NoTypeInformation -Encoding UTF8
Write-Output "CONTROL_TABLE_EVIDENCE: fields=$($registry.Fields.Count); derivedReviewed=$($fields.Count); fullAcceptance=False"
