# Выполняет selector cases на отдельных копиях согласованной книги.
# Использует существующие load-matrix assertions для State/path/comment/output,
# добавляя только тестовое семейство из шести нагрузок во временный VBA-модуль.
# Production-компоненты проверяются на неизменность. Все пути и действующие
# numerical tolerances сохраняются; эта серия не заменяет полную Light/Stress.
param(
    [Parameter(Mandatory=$true)][string]$SourceWorkbook,
    [Parameter(Mandatory=$true)][string]$MatrixPath,
    [Parameter(Mandatory=$true)][ValidatePattern('^[A-Za-z0-9_-]+$')][string]$Version,
    [string[]]$CaseIds = @(),
    [ValidateRange(10,3600)][int]$TimeoutSeconds = 600,
    [switch]$Resume
)
$ErrorActionPreference = 'Stop'
$root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$source = (Resolve-Path -LiteralPath $SourceWorkbook).Path
$sourceHash = (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash
$matrix = Get-Content -LiteralPath (Join-Path $root $MatrixPath) -Raw -Encoding UTF8 | ConvertFrom-Json
$matrixHash = (Get-FileHash -LiteralPath (Join-Path $root $MatrixPath) -Algorithm SHA256).Hash
if ($matrix.RequiredPairs -ne $matrix.CoveredPairs -or $matrix.HighRiskQuads -ne 48) { throw 'Selector plan has an incomplete structural gate.' }
$cases = @($matrix.Cases)
if ($CaseIds.Count) {
    $CaseIds = @($CaseIds | ForEach-Object { $_ -split ',' })
    $cases = @($cases | Where-Object { $CaseIds -contains $_.Id })
    if ($cases.Count -ne $CaseIds.Count) { throw 'Requested selector scope contains unknown or duplicate case IDs.' }
}
if (-not $cases.Count) { throw 'Empty selector runtime scope is not a successful matrix.' }
$folder = Join-Path ([IO.Path]::GetTempPath()) ('RC_NDM_Selector_' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $folder | Out-Null
$completed = [Collections.Generic.List[object]]::new()
$date = Get-Date -Format 'yyyy-MM-dd'

# Находит единственное поле по настоящему ключу/профилю, без адресов Config.
# Неизвестная строка или дубликат является ошибкой setup, а не новым default.
function Set-SelectorValue([object]$Book, [string]$RangeName, [string]$Key, [object]$Value, [string]$Profile='') {
    $range = $Book.Names.Item($RangeName).RefersToRange
    $data = $range.Value2
    $keyColumn = if ($Profile) { 2 } else { 1 }
    $row = 0
    for ($r=1; $r -le $range.Rows.Count; $r++) {
        if ([string]$data[$r,$keyColumn] -ceq $Key) {
            if ($row) { throw "Duplicate selector key: $Key" }
            $row=$r
        }
    }
    if (-not $row) { throw "Missing selector key: $Key" }
    $column = 2
    if ($Profile) {
        $column=0
        for ($r=1; $r -le $range.Rows.Count; $r++) {
            for ($c=3; $c -le $range.Columns.Count; $c++) {
                if ([string]$data[$r,$c] -ceq $Profile) {
                    if ($column -and $column -ne $c) { throw "Ambiguous selector profile: $Profile" }
                    $column=$c
                }
            }
        }
        if (-not $column) { throw "Missing selector profile: $Profile" }
    }
    $range.Cells.Item($row,$column).Value2 = $Value
}

# Считает hash фактических production-компонентов до и после временного
# расширения тестового семейства. Код численных владельцев не редактируется.
function Get-ProductionCodeHash([object]$Book) {
    $parts = [Collections.Generic.List[string]]::new()
    foreach ($component in $Book.VBProject.VBComponents) {
        if ($component.Name -like 'modTest*' -or $component.Name -like 'CFake*' -or $component.Name -like 'CTest*') { continue }
        $parts.Add($component.Name)
        if ($component.CodeModule.CountOfLines) { $parts.Add($component.CodeModule.Lines(1,$component.CodeModule.CountOfLines)) }
    }
    $sha = [Security.Cryptography.SHA256]::Create()
    try { return [BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes(($parts -join "`n")))).Replace('-','') }
    finally { $sha.Dispose() }
}

# Создает один изолированный Config. Тестовая вставка использует существующие
# размерно раздельные nc/nt/Mref; ни критерий, ни assertions не изменяются.
function New-SelectorFixture([object]$Case, [string]$Destination) {
    Copy-Item -LiteralPath $source -Destination $Destination
    $excel=$null; $book=$null
    try {
        $excel=New-Object -ComObject Excel.Application
        $excel.Visible=$false; $excel.DisplayAlerts=$false; $excel.EnableEvents=$false
        $excel.AutomationSecurity=3
        $book=$excel.Workbooks.Open($Destination,0,$false)
        $productionBefore=Get-ProductionCodeHash $book
        $module=$book.VBProject.VBComponents.Item('modTestBatchCalculation').CodeModule
        $code=$module.Lines(1,$module.CountOfLines)
        $needle='    If family = "Smoke" Then'
        if ([regex]::Matches($code,[regex]::Escape($needle)).Count -ne 1) { throw 'Test load-family insertion point is ambiguous.' }
        $insert=@'
    ' ДЛЯ ТЕСТОВ: selector interaction использует отдельный малый набор.
    ' Те же State/path/comment/writer assertions выполняются для всех 12 путей.
    If family = "Selector" Then
        result.Add Array("regularC", -0.05 * nc, 0.03 * mxRef, -0.02 * myRef)
        result.Add Array("regularT", 0.05 * nt, -0.03 * mxRef, 0.02 * myRef)
        result.Add Array("axialC", -0.05 * nc, 0#, 0#)
        result.Add Array("axialT", 0.75 * nt, 0#, 0#)
        result.Add Array("overloadC", -2# * nc, 0#, 0#)
        result.Add Array("overloadMixed", -2# * nc, 2# * mxRef, -2# * myRef)
    ElseIf family = "Smoke" Then
'@
        $module.DeleteLines(1,$module.CountOfLines)
        $insert=($insert -split '\r?\n') -join "`r`n"
        $module.AddFromString($code.Replace($needle,$insert))
        if ((Get-ProductionCodeHash $book) -ne $productionBefore) { throw 'Production code changed while preparing a selector fixture.' }
        $v=$Case.Values
        foreach ($entry in @(
            @('General.DiagramExtension',$v.Extension), @('Capacity.SolutionStrategy',$v.CapacityStrategy),
            @('SLS.Crack.InitiationSolutionStrategy',$v.FormationStrategy), @('Capacity.SearchMethod',$v.SearchMethod),
            @('Solver.Method',$v.SolverMethod), @('SLS.Crack.PsiMode',$v.PsiMode), @('SLS.Crack.PsiS','0.8')
        )) { Set-SelectorValue $book 'rngSystemSettings' $entry[0] $entry[1] }
        foreach ($profile in @('PR1','PR2')) {
            foreach ($role in @('Strength','CrackInitiation','CrackedState')) {
                Set-SelectorValue $book 'rngCalculationProfiles' ("MaterialModel.$role.ConcreteDiagram") $v.ConcreteDiagram $profile
                Set-SelectorValue $book 'rngCalculationProfiles' ("MaterialModel.$role.SteelDiagram") $v.SteelDiagram $profile
            }
            Set-SelectorValue $book 'rngCalculationProfiles' 'MaterialModel.Strength.ConcreteTension' $v.StrengthTension $profile
            Set-SelectorValue $book 'rngCalculationProfiles' 'MaterialModel.CrackInitiation.ConcreteTension' 'UseDiagram' $profile
            Set-SelectorValue $book 'rngCalculationProfiles' 'MaterialModel.CrackedState.ConcreteTension' 'Ignore' $profile
        }
        $book.Save()
        $book.Close($false); $book=$null
        return $productionBefore
    } finally {
        if ($book) { $book.Close($false) }
        if ($excel) { $excel.Quit(); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) | Out-Null }
    }
}

# Принимает только завершенные actual runtime logs с 72 path cases,
# save/reopen, typed/comment assertions и фактической выбранной системой моделей.
function Test-SelectorReport([string]$Path, [object]$Case) {
    $text=Get-Content -LiteralPath $Path -Raw -Encoding UTF8
    $v=$Case.Values
    if (-not $text.Contains('WATCHDOG_COMPLETED: exit=0; source unchanged=True') -or
        $text -notmatch 'TOTAL_AUDIT03_LOAD_MATRIX:[^\r\n]*independentCases=72;[^\r\n]*failed=0' -or
        $text -notmatch 'RESULTS_SAVE_REOPEN:[^\r\n]*equal=True;' -or
        -not $text.Contains("MODE_AFTER_SUITE: $($v.Extension)")) { return $false }
    $directStates=@([regex]::Matches($text,'(?m)^MATRIX_STATE:[^\r\n]*\|type=StrengthState\|spec=([^\r\n]+)\r?$'))
    if (-not $directStates.Count) { return $false }
    foreach ($state in $directStates) {
        $expected='ULS(I)|' + $v.ConcreteDiagram + '|' + $v.StrengthTension + '|' + $v.SteelDiagram + '|'
        if (-not $state.Groups[1].Value.StartsWith($expected)) { return $false }
    }
    if ($v.PsiMode -eq 'User') {
        $widths=@([regex]::Matches($text,'(?m)^MATRIX_CRACK_DATA:[^\r\n]*\|calculated=True\|formationDataAvailable=(True|False)\|[^\r\n]*\|psi=([^|]+)\|'))
        if (-not $widths.Count) { return $false }
        foreach ($width in $widths) {
            # Доступность Formation/Post - typed runtime факт из matrix oracle.
            # Действующий контракт использует fallback 1 при недоступной точке
            # независимо от User; текст комментария не назначает этот результат.
            $expectedPsi=if ($width.Groups[1].Value -eq 'True') { 0.8 } else { 1.0 }
            if ([double]::Parse($width.Groups[2].Value,[Globalization.CultureInfo]::InvariantCulture) -ne $expectedPsi) { return $false }
        }
    }
    return $true
}

foreach ($case in $cases) {
    $relativeReport="docs/regression/Audit03/selector_$($case.Id)_${Version}_${date}.txt"
    $report=Join-Path $root $relativeReport
    if (Test-Path -LiteralPath $report) {
        $old=Get-Content -LiteralPath $report -Raw -Encoding UTF8
        if (-not $Resume -or -not (Test-SelectorReport $report $case) -or
            -not $old.Contains("SELECTOR_ORIGINAL_SOURCE_SHA256: $sourceHash") -or
            -not $old.Contains("SELECTOR_MATRIX_SHA256: $matrixHash")) { throw "Existing selector evidence is not reusable: $relativeReport" }
        $completed.Add(@{Id=$case.Id; Report=$relativeReport; Reused=$true})
        continue
    }
    $fixture=Join-Path $folder ($case.Id + '.xlsm')
    try { $productionHash=New-SelectorFixture $case $fixture }
    catch {
        @("SELECTOR_SETUP_FAILURE: $($_.Exception.Message)", "STACK: $($_.ScriptStackTrace)",
            "SELECTOR_ORIGINAL_SOURCE_SHA256: $sourceHash", "SELECTOR_MATRIX_SHA256: $matrixHash",
            'RUNTIME_STATUS: NotRun; no equilibrium outcome assigned') |
            Set-Content -LiteralPath $report -Encoding UTF8
        throw
    }
    Write-Output "SELECTOR_STARTED: $($case.Id); values=$($case.Values | ConvertTo-Json -Compress)"
    & (Join-Path $PSScriptRoot 'Run-Audit03Watchdog.ps1') -SourceWorkbook $fixture -ReportPath $relativeReport `
        -Macro 'modTestBatchCalculation.RunAudit03BroadLoadMatrixTests' -MacroArgument1 $case.Values.Shape `
        -MacroArgument2 'Selector' -Mode $case.Values.Extension -VerifyResultsReopen -TimeoutSeconds $TimeoutSeconds
    Add-Content -LiteralPath $report -Encoding UTF8 -Value @(
        "SELECTOR_ORIGINAL_SOURCE_SHA256: $sourceHash", "SELECTOR_MATRIX_SHA256: $matrixHash",
        "SELECTOR_PRODUCTION_CODE_SHA256: $productionHash", ('SELECTOR_VALUES: ' + ($case.Values | ConvertTo-Json -Compress))
    )
    if ($LASTEXITCODE -ne 0 -or -not (Test-SelectorReport $report $case)) { throw "Selector case failed; preserve counterexample: $relativeReport" }
    $completed.Add(@{Id=$case.Id; Report=$relativeReport; Reused=$false})
    Write-Output "SELECTOR_COMPLETED: $($case.Id); independentPathCases=72"
}
if ((Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash -ne $sourceHash) { throw 'Original source workbook changed.' }
$all=$completed.Count -eq $matrix.CaseCount
$summary=Join-Path $root "docs/regression/Audit03/selector_summary_${Version}_${date}.json"
if (Test-Path -LiteralPath $summary) { throw 'Selector summary already exists; choose a new Version.' }
@{ OriginalSourceSHA256=$sourceHash; MatrixSHA256=$matrixHash; Completed=$completed; CaseCount=$completed.Count;
    IndependentPathCases=72*$completed.Count; FullSelectorAcceptance=$all; SourceUnchanged=$true } |
    ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $summary -Encoding UTF8
Write-Output "TOTAL_AUDIT03_SELECTOR_RUNNER: completed=$($completed.Count); fullSelectorAcceptance=$all; independentPathCases=$(72*$completed.Count); sourceUnchanged=True"
