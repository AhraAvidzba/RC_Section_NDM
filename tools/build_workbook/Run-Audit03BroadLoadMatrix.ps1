# Выполняет широкую матрицу на изолированных копиях одной проверенной книги.
# Каждый VBA-runner проверяет все десять путевых вариантов, typed результаты,
# комментарии четырех writer-ов/txt-отчета и сохраненный Results. При отказе
# серия останавливается; существующее отрицательное доказательство не заменяется.
param(
    [Parameter(Mandatory=$true)][string]$SourceWorkbook,
    [Parameter(Mandatory=$true)][ValidatePattern('^[A-Za-z0-9_-]+$')][string]$Version,
    [string[]]$Shapes = @('CircleSym', 'CircleUneven', 'RectRectangle', 'RectL',
        'RectTwoLeft', 'RectTwoRight', 'RoundedSimple', 'RoundedTapered',
        'RoundedMixed', 'HollowCentered', 'HollowOffset', 'HollowThin', 'ImportedFixture'),
    [ValidateSet('Light', 'Stress')][string[]]$Families = @('Light', 'Stress'),
    [ValidateSet('No', 'Yes')][string[]]$Modes = @('No', 'Yes'),
    [ValidateRange(10, 3600)][int]$TimeoutSeconds = 3600,
    [switch]$Resume
)
$ErrorActionPreference = 'Stop'
$root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$source = (Resolve-Path -LiteralPath $SourceWorkbook).Path
$sourceHash = (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash
$date = Get-Date -Format 'yyyy-MM-dd'
$completed = 0
$skipped = 0
$independentCases = 0

# Повторно использует только полностью завершенный отчет этой же source-книги.
# Наличие файла/строки TOTAL без успешного save-reopen и watchdog недостаточно.
function Test-CompletedMatrixReport([string]$Path, [int]$ExpectedCases, [string]$Shape,
        [string]$Family, [string]$Mode) {
    $text = Get-Content -LiteralPath $Path -Raw -Encoding UTF8
    return $text.Contains("SOURCE_SHA256: $sourceHash") -and
        $text.Contains("TOTAL_AUDIT03_LOAD_MATRIX: shape=$Shape; family=$Family; independentCases=$ExpectedCases;") -and
        $text.Contains("MODE_AFTER_SUITE: $Mode") -and
        ($text -match 'TOTAL_AUDIT03_LOAD_MATRIX:[^\r\n]*failed=0(?:\r?\n|$)') -and
        ($text -match 'RESULTS_SAVE_REOPEN:[^\r\n]*equal=True;') -and
        $text.Contains('SOURCE_UNCHANGED: True') -and
        $text.Contains('WATCHDOG_COMPLETED: exit=0; source unchanged=True')
}

foreach ($shape in $Shapes) {
    if ($shape -notmatch '^[A-Za-z]+$') { throw 'Недопустимое имя тестовой формы.' }
    foreach ($family in $Families) {
        $expectedCases = if ($family -eq 'Light') { 490 } else { 240 }
        foreach ($mode in $Modes) {
            $modeName = if ($mode -eq 'Yes') { 'on' } else { 'off' }
            $relativeReport = "docs/regression/Audit03/broad_matrix_${shape}_${family}_${modeName}_${Version}_${date}.txt"
            $report = Join-Path $root $relativeReport
            if (Test-Path -LiteralPath $report) {
                if ($Resume -and (Test-CompletedMatrixReport $report $expectedCases $shape $family $mode)) {
                    $skipped++
                    $independentCases += $expectedCases
                    Write-Output "MATRIX_RESUMED: $shape/$family/$mode; cases=$expectedCases"
                    continue
                }
                throw "Отчет уже существует и не разрешен для reuse: $relativeReport. Выберите новый Version; отрицательный/неполный лог сохраняется."
            }
            Write-Output "MATRIX_STARTED: $shape/$family/$mode; sourceSHA=$sourceHash"
            & (Join-Path $PSScriptRoot 'Run-Audit03Watchdog.ps1') -SourceWorkbook $source `
                -ReportPath $relativeReport -Macro 'modTestBatchCalculation.RunAudit03BroadLoadMatrixTests' `
                -MacroArgument1 $shape -MacroArgument2 $family -Mode $mode `
                -VerifyResultsReopen -TimeoutSeconds $TimeoutSeconds
            if ($LASTEXITCODE -ne 0) { throw "Матрица остановлена на $shape/$family/$mode. Контрпример: $relativeReport" }
            if (-not (Test-CompletedMatrixReport $report $expectedCases $shape $family $mode)) {
                throw "Матрица не подтвердила полный контракт $shape/$family/${mode}: $relativeReport"
            }
            $completed++
            $independentCases += $expectedCases
            Write-Output "MATRIX_COMPLETED: $shape/$family/$mode; cases=$expectedCases"
        }
    }
}
if ((Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash -ne $sourceHash) {
    throw 'Source-книга изменилась во время матрицы.'
}
Write-Output "TOTAL_AUDIT03_MATRIX_RUNNER: newRuns=$completed; reusedRuns=$skipped; independentCases=$independentCases; sourceUnchanged=True"
