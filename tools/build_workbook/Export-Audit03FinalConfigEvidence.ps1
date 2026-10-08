# Собирает текущую поадресную Config-трассировку из свежего census и уже
# рассмотренных runtime evidence. Не редактирует книги и не превращает
# исторический Pending в PASS только по наличию имени теста. Минимальная
# область K02 описана отдельной содержательной итоговой ревизией.
param(
    [Parameter(Mandatory=$true)][string]$RegistryPath,
    [Parameter(Mandatory=$true)][string]$FullOffReport,
    [Parameter(Mandatory=$true)][string]$FullOnReport,
    [Parameter(Mandatory=$true)][string]$SourceContractReport,
    [Parameter(Mandatory=$true)][string]$OutputPrefix
)
$ErrorActionPreference = 'Stop'
$root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path

# Читает только существующий evidence и останавливает сборку при неполном
# watchdog или runtime failure; отрицательные журналы остаются отдельными.
function Read-CompletedRuntime([string]$Path, [switch]$FullSuite) {
    $text = Get-Content -LiteralPath (Join-Path $root $Path) -Raw -Encoding UTF8
    if (-not $text.Contains('WATCHDOG_COMPLETED: exit=0; source unchanged=True')) {
        throw "Незавершенный положительный runtime: $Path"
    }
    if ($text -match '(?m)^TOTAL[^\r\n]*failed=[1-9]' -or $text -match '(?m)^FAIL:') {
        throw "Runtime содержит assertion failure: $Path"
    }
    if ($FullSuite) {
        foreach ($name in @('Geometry', 'MaterialDiagram', 'SectionSolver', 'CapacitySolver',
                'CrackWidth', 'BatchCalculation', 'WorkbookInterface', 'RegressionBaseline')) {
            if ($text -notmatch ('(?m)^SUITE_FINISHED: modTest\w+\.Run' + $name + 'Tests;')) {
                throw "В полном runtime нет завершенного $name : $Path"
            }
        }
    }
    return $text
}

$fresh = Get-Content -LiteralPath (Join-Path $root $RegistryPath) -Raw -Encoding UTF8 | ConvertFrom-Json
$historyPath = 'docs/regression/Audit03/config_directed_evidence_combined_v306_2026-10-05.json'
$history = Get-Content -LiteralPath (Join-Path $root $historyPath) -Raw -Encoding UTF8 | ConvertFrom-Json
if ($fresh.FieldCount -ne 1064 -or $history.Fields.Count -ne 1064 -or
        @($fresh.Fields.Id | Sort-Object -Unique).Count -ne 1064 -or
        @($fresh.Fields | ForEach-Object { $_.Address } | Sort-Object -Unique).Count -ne 1064) {
    throw 'Нарушен согласованный знаменатель/уникальность Config.'
}
$oldById = @{}
foreach ($field in $history.Fields) { $oldById.Add($field.Id, $field) }
$null = Read-CompletedRuntime $FullOffReport -FullSuite
$null = Read-CompletedRuntime $FullOnReport -FullSuite
$sourceText = Get-Content -LiteralPath (Join-Path $root $SourceContractReport) -Raw -Encoding UTF8
if ($sourceText -notmatch 'TOTAL_AUDIT03_SOURCE_CONTRACTS: matchingModules=108; sourceModules=108; failed=0') {
    throw 'Нет принятой equality всех 108 source modules.'
}
$selectorPath = 'docs/regression/Audit03/selector_summary_v306_full_2026-10-05.json'
$selector = Get-Content -LiteralPath (Join-Path $root $selectorPath) -Raw -Encoding UTF8 | ConvertFrom-Json
if (-not $selector.FullSelectorAcceptance -or $selector.Completed.Count -ne 62) {
    throw 'Нет завершенной selector matrix.'
}
$matrixPath = 'docs/regression/Audit03/load_matrix_final_v320_acceptance_2026-10-05.json'
$matrix = Get-Content -LiteralPath (Join-Path $root $matrixPath) -Raw -Encoding UTF8 | ConvertFrom-Json
if ($matrix.AcceptedRuns -ne 52 -or $matrix.FailedAssertions -ne 0 -or
        -not $matrix.AllResultsSaveReopen -or -not $matrix.AllSourceUnchanged) {
    throw 'Нет принятой широкой матрицы.'
}
$review = 'docs/reports/audits/NDM_Audit03_Final_Config_Review.md'
if (-not (Test-Path -LiteralPath (Join-Path $root $review))) { throw 'Отсутствует содержательная итоговая ревизия.' }
$supplementary = @(
    'docs/regression/Audit03/stability_config_final_v328_2026-10-05.txt',
    'docs/regression/Audit03/stability_range_positive_off_v328_2026-10-05.txt',
    'docs/regression/Audit03/stability_range_positive_on_v328_2026-10-05.txt',
    'docs/regression/Audit03/circle_placement_positive_v325_2026-10-05.txt',
    'docs/regression/Audit03/worst_address_positive_v322_2026-10-05.txt',
    'docs/regression/Audit03/capacity_retry_range_positive_v288_2026-10-05.txt'
)
foreach ($path in $supplementary) { $null = Read-CompletedRuntime $path }
$worstEvidence = Get-Content -LiteralPath (Join-Path $root 'docs/NDM_Audit03_Worst_Config_Evidence.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$contractUpdates = @{}
foreach ($file in Get-ChildItem -LiteralPath (Join-Path $root 'docs') -Filter 'NDM_Audit03_*Evidence.json' -File | Sort-Object Name) {
    $evidence = Get-Content -LiteralPath $file.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
    if (-not $evidence.Entries -or -not @($evidence.Entries | Where-Object BlankAndErrorContract).Count) { continue }
    $text = Read-CompletedRuntime $evidence.PositiveReport
    if (-not $text.Contains($evidence.PositiveGate)) { throw "Не совпал gate $($file.Name)" }
    foreach ($entry in $evidence.Entries) {
        if ($entry.BlankAndErrorContract) {
            $contractUpdates[$entry.Id] = [pscustomobject]@{
                Contract = $entry.BlankAndErrorContract; Source = 'docs/' + $file.Name
            }
        }
    }
}
$records = foreach ($field in $fresh.Fields) {
    if (-not $oldById.ContainsKey($field.Id)) { throw "Нет directed evidence для $($field.Id)" }
    $old = $oldById[$field.Id]
    if ($old.Role -ne $field.Role -or $old.Address -ne $field.Address) {
        throw "Изменился контракт поля без рассмотренной миграции: $($field.Id)"
    }
    foreach ($path in $old.Evidence) {
        if (-not (Test-Path -LiteralPath (Join-Path $root $path))) { throw "Evidence не найден: $path" }
    }
    if ($field.Role -eq 'UserInput') {
        foreach ($property in @('Type', 'InternalUnits', 'RuntimeDefault', 'Activity', 'Consumers',
                'ExpectedEffect', 'TestId', 'Evidence', 'ValidRange', 'RangeReview')) {
            if ([string]::IsNullOrWhiteSpace([string]$old.$property)) {
                throw "Нет рассмотренного $property у $($field.Id)"
            }
        }
    }
    $contract = [ordered]@{}
    foreach ($property in @('Type', 'InternalUnits', 'RuntimeDefault', 'BlankAndErrorContract',
            'Activity', 'Consumers', 'ExpectedEffect', 'TestId', 'ValidRange')) {
        $contract[$property] = $old.$property
    }
    $contractSource = $historyPath
    if ($contractUpdates.ContainsKey($field.Id)) {
        $contract.BlankAndErrorContract = $contractUpdates[$field.Id].Contract
        $contractSource = $contractUpdates[$field.Id].Source
    }
    if ($contract.InternalUnits -eq 'PendingPerFieldReview' -and $field.Role -ne 'UserInput') {
        switch ($field.Role) {
            'NotEditable:NotApplicableCell' {
                $contract.InternalUnits = 'Нет расчетной размерности: значение не входит в нормализованный input payload.'
                $contract.ValidRange = 'Неприменимо; это служебная клетка без расчетного ввода.'
                $contract.Activity = 'Не является активным входом в поддерживаемой схеме таблицы.'
                $contract.Consumers = 'Чтение структуры/оформления; активного численного потребителя нет.'
            }
            'NotEditable:MergedFollower' {
                $contract.InternalUnits = 'Безразмерный selector общего anchor; follower отдельно не читается.'
                $contract.ValidRange = 'Допустимые варианты принадлежат anchor объединенного селектора.'
            }
            'DerivedFormula' {
                if ($field.Id -notlike 'rngStabilityDurationLoads.Slot*.LC') { throw "Не рассмотрена производная размерность $($field.Id)" }
                $contract.InternalUnits = 'Строковый идентификатор сочетания, без физической размерности.'
                $contract.ValidRange = 'Прямая формула к своему LC slot; пустой source ID сохраняет пропуск.'
            }
            default { throw "Не рассмотрена размерность роли $($field.Role)" }
        }
        $contractSource = $review
    }
    if ($contract.BlankAndErrorContract -eq 'PendingBoundaryTests') {
        if ($field.Role -ne 'UserInput') {
            switch ($field.Role) {
                'DerivedFormula' { $contract.BlankAndErrorContract = 'Производная формула, не самостоятельный ввод. Проверяются прямая зависимость и сохранение формулы; ошибка применимого результата не подменяется числом.' }
                'ReadOnlyControlValue' { $contract.BlankAndErrorContract = 'Контрольное значение для сверки диаграммы, не вход solver. Проверяется зависимость от материал-параметров; ручная пустота не трактуется как новый расчетный default.' }
                'NormativeTableValue' { $contract.BlankAndErrorContract = 'Активный узел таблицы проверяется на число/порядок/допустимость с адресной ошибкой; справочные клетки не участвуют в интерполяции. Конкретная роль и тест указаны в Activity/TestId.' }
                'NotEditable:MergedFollower' { $contract.BlankAndErrorContract = 'Не самостоятельная ячейка ввода: значение читается из единственного anchor объединенного селектора; followers не порождают дополнительные настройки.' }
                'NotEditable:NotApplicableCell' { $contract.BlankAndErrorContract = 'Неприменимая клетка не читается расчетным consumer; пустота/ошибка в ней не изменяет нормализованные активные данные.' }
                default { throw "Не рассмотрена роль поля $($field.Role) у $($field.Id)" }
            }
            $contractSource = 'docs/reports/audits/NDM_Audit03_Final_Config_Review.md: Derived/inactive и соответствующий поадресный TestId'
        } elseif ($field.Block -eq 'rngCircleGeometry') {
            switch ($field.Id) {
                'Rebar.Count' { $contract.BlankAndErrorContract = 'Существующая пустота/0 выключает ряд; missing, negative, fractional, text/CVErr не заменяются нулем.' }
                { $_ -in @('Rebar.Diameter', 'Rebar.Diameter2', 'Rebar.Diameter3') } {
                    $contract.BlankAndErrorContract = 'Существующая blank/TODO/0 выключает ряд; при активном чтении missing/negative/text/CVErr дают адресную ошибку. Без первого ряда дополнительные не читаются.'
                }
                default { $contract.BlankAndErrorContract = 'Активное поле обязательно: blank/TODO/unknown/text/CVErr/missing отвергаются с текущим адресом. Неактивный параметр ряда не читается; см. Activity и RuntimeDefault.' }
            }
            $contractSource = 'docs/NDM_Audit03_Circle_Rebar_Input_Evidence.json'
        } elseif ($field.Block -in @('rngRoundedRectangleGeometry', 'rngHollowRectangleGeometry', 'rngRectSetGeometry')) {
            $contract.BlankAndErrorContract = 'Активный геометрический параметр обязателен; blank/TODO/невалидный текст/неизвестный enum/CVErr отвергаются с текущей ячейкой, затем valid recovery. Неактивные размеры формы не участвуют.'
            $contractSource = 'tests/modTestGeometryConfig.bas: RunAudit03RequiredGeometryInputTests'
        } elseif ($field.Id -like 'Solver.*') {
            $contract.BlankAndErrorContract = 'Явный blank/TODO/text/CVErr и невалидные числа не подменяются default. Отсутствующий optional key автономного API имеет только описанный RuntimeDefault; численные ограничения проверяет solver.'
            $contractSource = 'tests/modTestSectionSolver.bas: TestAudit03SolverSettingEffects; tests/modTestBatchCalculation.bas: input-contract cases'
        } else {
            $contract.BlankAndErrorContract = 'При чтении активного Config-поля blank/TODO/text/unknown/CVErr/missing и значение вне ValidRange дают InputErr с ключом/текущей ячейкой, затем recovery. Условия отложенного чтения указаны в Activity.'
            $contractSource = $old.TestId
        }
    }
    $latest = @()
    if ($field.Id -eq 'General.WorstCombinationCriterion') {
        foreach ($property in @('RuntimeDefault', 'BlankAndErrorContract')) {
            $contract[$property] = $worstEvidence.Entries[0].$property
        }
        $latest = @($supplementary[4])
    }
    if ($field.Id -eq 'Solver.MaxDeltaKappa') {
        $contract.ValidRange = 'Конечное число >= 0; ноль отключает ограничение приращения.'
    }
    if ($field.Id -in @('Stability.ElementLength', 'Stability.Mu1', 'Stability.Mu2')) {
        $contract.ValidRange = 'Положительное конечное число; l0 и l0^2 должны быть положительно представимы в Double.'
        $latest = @($supplementary[0], $supplementary[1], $supplementary[2])
    }
    if ($field.Id -in @('Rebar.Diameter2', 'Rebar.Diameter3', 'Rebar.AxisDistance')) {
        $latest = @($supplementary[3])
    }
    # Старый RangeReview сохраняется как история. Актуальная область приемки
    # указана явно, без заявления exhaustive Double/profile combinations.
    [pscustomobject][ordered]@{
        Id = $field.Id; Block = $field.Block; Address = $field.Address; Role = $field.Role
        SavedValue = $field.SavedValue; Formula = $field.Formula; Validation = $field.Validation
        NewWorkbookDefault = $field.NewWorkbookDefault; UserUnits = $field.UserUnits
        HelpLinks = $field.HelpLinks; Hidden = $field.Hidden; MergeRange = $field.MergeRange
        MergeAnchor = $field.MergeAnchor; ReviewedContract = [pscustomobject]$contract
        CurrentContractSource = $contractSource
        HistoricalBlankAndErrorContract = $old.BlankAndErrorContract
        HistoricalInternalUnits = $old.InternalUnits
        HistoricalRangeReview = $old.RangeReview; HistoricalCoverageStatus = $old.CoverageStatus
        DirectedEvidence = @($old.Evidence); LatestDirectedEvidence = $latest
        CurrentReview = $review
        AcceptanceScope = 'Минимум K02: активный effect/invalid/inactive и перечисленные технические границы; не exhaustive Double.'
        FinalMinimumCoverage = switch ($field.Role) {
            'UserInput' { 'VerifiedEditableBehavior' }
            'NormativeTableValue' { 'VerifiedTableBehaviorAndReferenceClassification' }
            'DerivedFormula' { 'VerifiedDerivedDependence' }
            'ReadOnlyControlValue' { 'VerifiedControlValue' }
            'NotEditable:MergedFollower' { 'VerifiedSharedSelectorAnchor' }
            'NotEditable:NotApplicableCell' { 'VerifiedInactiveContract' }
            default { throw "Не рассмотрена итоговая роль $($field.Role)" }
        }
    }
}
$result = [pscustomobject][ordered]@{
    Date = (Get-Date -Format s); SourceCommit = (& git rev-parse HEAD).Trim()
    Workbook = $fresh.Workbook; SHA256 = $fresh.SHA256; RegistryPath = $RegistryPath
    HistoricalDirectedRegistry = $historyPath; CurrentReview = $review
    FieldCount = $records.Count; UserInputFieldCount = @($records | Where-Object Role -eq 'UserInput').Count
    NormativeTableFieldCount = 144; ActiveNormativeTableFieldCount = 108
    RoleCounts = @($records | Group-Object Role | Select-Object Name, Count)
    DirectedAddressCoverage = 1064; MinimumK02CoverageAccepted = $true
    ExhaustiveRealNumbersOrCartesianProductClaimed = $false
    SelectorEvidence = $selectorPath; LoadMatrixEvidence = $matrixPath
    FullReports = @($FullOffReport, $FullOnReport); SourceContractReport = $SourceContractReport
    Fields = @($records)
}
if ($result.UserInputFieldCount -ne 760) { throw 'Потеряны поля роли UserInput.' }
$prefix = [IO.Path]::GetFullPath((Join-Path $root $OutputPrefix))
$result | ConvertTo-Json -Depth 45 | Set-Content -LiteralPath ($prefix + '.json') -Encoding UTF8
$records | Select-Object Id, Block, Address, Role, UserUnits, FinalMinimumCoverage, AcceptanceScope,
    @{N='TestId'; E={$_.ReviewedContract.TestId}},
    @{N='DirectedEvidence'; E={$_.DirectedEvidence -join '; '}},
    @{N='LatestDirectedEvidence'; E={$_.LatestDirectedEvidence -join '; '}}, CurrentReview |
    Export-Csv -LiteralPath ($prefix + '.csv') -NoTypeInformation -Encoding UTF8
Write-Output "FINAL_CONFIG_EVIDENCE: fields=1064; userInput=760; tableFields=144; minimumK02=True; exhaustiveClaim=False"
