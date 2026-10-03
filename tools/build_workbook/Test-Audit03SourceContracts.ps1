# Проверяет конкретные архитектурные границы Audit03 и код реальной книги.
# Не заменяет VBA/runtime/UI приемку и не переписывает исторический Audit02 gate.
param(
    [string]$BaselineRoot = 'C:\Users\avidzba\AppData\Local\Temp\RC_NDM_Audit03_Baseline_df10412f',
    [string]$ExportPath = 'docs/regression/Audit03/VBA_All_Code.txt',
    [string]$ReportPath = 'docs/regression/Audit03/source_contracts_2026-10-02.txt'
)
$ErrorActionPreference = 'Stop'
$root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$lines = New-Object System.Collections.Generic.List[string]
$failed = 0

# Фиксирует проверяемый контракт отдельно от числа совпавших модулей.
function Assert-Contract([string]$name, [bool]$passed, [string]$detail) {
    if (-not $passed) { $script:failed++ }
    $lines.Add("CONTRACT|$name|passed=$passed|$detail")
}

# Удаляет только экспортную оболочку модуля, как штатный VBA-import.
function Get-SourceBody([string]$text) {
    $result = New-Object System.Collections.Generic.List[string]
    $header = $false
    foreach ($line in $text -split "`r?`n") {
        if ($line -match '^\uFEFF?VERSION\s+') { continue }
        if ($line -match '^\uFEFF?BEGIN\s*$') { $header = $true; continue }
        if ($header) {
            if ($line -match '^\uFEFF?END\s*$') { $header = $false }
            continue
        }
        if ($line -match '^\uFEFF?Attribute\s+VB_') { continue }
        $result.Add($line.TrimEnd())
    }
    return ($result -join "`n").Trim()
}

# Сравнивает именно VBE-представление: системная CP1251, регистр идентификаторов
# и 15 значащих цифр Double. Строки и комментарии не заменяются другим текстом.
function Get-VbeCanonicalBody([string]$text) {
    $body = Get-SourceBody $text
    $encoding = [Text.Encoding]::GetEncoding(1251)
    $body = $encoding.GetString($encoding.GetBytes($body))
    $canonical = foreach ($line in $body -split "`n") {
        [regex]::Replace($line, '"(?:[^"]|"")*"|''.*$|(?<![\w&])(?:\d+\.\d*|\d+)(?:[Ee][+-]?\d+)?#?(?!\w)', {
            param($token)
            if ($token.Value.StartsWith('"') -or $token.Value.StartsWith("'")) { return $token.Value }
            $number = [double]::Parse($token.Value.TrimEnd('#'), [Globalization.CultureInfo]::InvariantCulture)
            return $number.ToString('G15', [Globalization.CultureInfo]::InvariantCulture)
        })
    }
    return $canonical -join "`n"
}

# Убирает комментарии только для поиска запрещенных кодовых зависимостей.
function Get-CodeOnly([string]$path) {
    return (Get-Content -LiteralPath (Join-Path $root $path) -Encoding UTF8 |
        Where-Object { $_ -notmatch "^\s*'" }) -join "`n"
}

$production = @(Get-ChildItem -LiteralPath (Join-Path $root 'src') -Recurse -File -Filter '*.cls')
$testClasses = @(Get-ChildItem -LiteralPath (Join-Path $root 'tests') -Recurse -File -Filter '*.cls')
$beforeClasses = @(Get-ChildItem -LiteralPath (Join-Path $BaselineRoot 'src') -Recurse -File -Filter '*.cls' |
    ForEach-Object { $_.Name })
# Пользовательский scope LoadPath переименовал существующий описатель без
# добавления ответственности/класса. Сопоставляем его с baseline по этой явной
# паре; остальные новые классы по-прежнему запрещены.
$comparableClasses = @($production.Name | ForEach-Object { if ($_ -eq 'CLoadPathDescriptor.cls') { 'CCapacityLoadPath.cls' } else { $_ } })
$difference = @(Compare-Object ($beforeClasses | Sort-Object) ($comparableClasses | Sort-Object))
$added = @($difference | Where-Object { $_.SideIndicator -eq '=>' })
$removed = @($difference | Where-Object { $_.SideIndicator -eq '<=' } | ForEach-Object { $_.InputObject } | Sort-Object)
$expectedRemoved = @('CBatchStatusPolicy.cls', 'CCrackWidthFormulaCalculator.cls')
Assert-Contract 'noNewClasses' ($added.Count -eq 0 -and $testClasses.Count -eq 3) "newProduction=$($added.Count); test=$($testClasses.Count)"
Assert-Contract 'twoSpecifiedMerges' ($production.Count -eq 83 -and ($removed -join '|') -eq ($expectedRemoved -join '|')) "production=$($production.Count); removed=$($removed -join ',')"

$sources = @(Get-ChildItem -Path (Join-Path $root 'src'), (Join-Path $root 'tests') -Recurse -File |
    Where-Object { $_.Extension -in '.cls', '.bas' })
foreach ($file in $sources) {
    $code = Get-Content -LiteralPath $file.FullName -Raw -Encoding UTF8
    Assert-Contract "removedConsumers.$($file.BaseName)" ($code -notmatch '\bCBatchStatusPolicy\b|\bCCrackWidthFormulaCalculator\b') 'No deleted-class consumer or compatibility alias'
}
foreach ($path in @('src/Interfaces/ILimitSearchProblem.cls', 'src/Solver/CLoadMultiplierSearch.cls',
        'src/Solver/CUltimateStrainSearch.cls', 'src/Solver/CLimitSearchResult.cls')) {
    $code = Get-CodeOnly $path
    Assert-Contract "generic.$path" ($code -notmatch 'DomainContext|As CCrackWidthCalculator|As CCrackFormationCalculator|As CCapacitySolver|CStateRepository') 'No domain downcast or shared probe repository'
}
$width = Get-CodeOnly 'src/Crack/CCrackWidthCalculator.cls'
foreach ($name in @('CrackWidthFromData', 'UtilizationFromData')) {
    $match = [regex]::Match($width, "(?ms)^Public Function $name\b.*?^End Function")
    Assert-Contract "pureFormula.$name" ($match.Success -and $match.Value -notmatch 'CSectionState|CSectionSolver|\brs\w+\b|NumFail|\.Solve|SetWidthResult|\bm[A-Z]\w+') 'Prepared numbers only; no State/status/solve or result mutation'
}
foreach ($path in @('src/Crack/CCrackWidthCalculator.cls', 'src/Crack/CLongitudinalCrackCalculator.cls', 'src/Stability/CStabilityCalculator.cls')) {
    $code = Get-CodeOnly $path
    Assert-Contract "formulaStatuses.$path" ($code -notmatch '\brsNumericalFailure\b|"NumFail"|\.Solve(?:By|With|\b)') 'No equilibrium solve or locally assigned numerical failure'
}
$batch = Get-CodeOnly 'src/Batch/CBatchSectionCalculator.cls'
$packagingCalls = [regex]::Matches($batch, '(?m)^\s*StoreCrackAggregateSnapshot\s').Count
Assert-Contract 'oneFinalCrackPackaging' ($packagingCalls -eq 1) "calls=$packagingCalls"
$meta = Get-CodeOnly 'src/Common/CResultMeta.cls'
Assert-Contract 'noIndependentLifecycleSetters' ($meta -notmatch 'Property Let (InternalStatus|ResultCode|ResultKind|Applies|Calculated)\b') 'Atomic SetResult owns typed outcome and lifecycle'

# Проверяет реальные границы публикации, а не только отсутствие setter-а.
# Mutable fill API остаются у builders, но не меняют уже выданные snapshots.
foreach ($path in @('src/Batch/CCombinationResult.cls', 'src/Batch/CStrengthResult.cls',
        'src/Batch/CDirectStateResult.cls', 'src/Batch/CSectionStateResult.cls',
        'src/Batch/CStateRepository.cls', 'src/Crack/CCrackResult.cls',
        'src/Crack/CCrackFormationResult.cls', 'src/Crack/CCrackWidthResult.cls',
        'src/Crack/CLongitudinalCrackResult.cls', 'src/Solver/CCapacityResult.cls',
        'src/Stability/CStabilityResult.cls')) {
    $code = Get-CodeOnly $path
    $mutators = [regex]::Matches($code, '(?m)^Public Sub [^\r\n]*(?:\r?\n\s+[^\r\n]*)*?\r?\n\s+AssertWritable\b')
    $declared = [regex]::Matches($code, '(?m)^Public Sub ').Count
    Assert-Contract "publishedGuard.$path" ($code -match 'Friend Sub Freeze\(\)' -and
        $code -match 'If mFrozen Then Err.Raise vbObjectError \+ 4220' -and $mutators.Count -eq $declared) "guarded=$($mutators.Count); mutators=$declared"
    Assert-Contract "noPublicResultFields.$path" ($code -notmatch '(?m)^Public \w+ As ') 'Published values are read-only; fill operations have lifecycle guard'
}
Assert-Contract 'batchPublicationBoundary' ($batch -match 'mResults\(index\)\.Freeze' -and
    $batch -match '(?ms)Private Sub ClearResultRow.*?Set mResults\(index\) = New CCombinationResult') 'ResultAt freezes tree; repeated Execute replaces LC instead of resetting an old published object'

foreach ($path in @('src/Geometry/CGeometryRoundedRectangle.cls', 'src/Geometry/CGeometryHollowRectangle.cls')) {
    $code = Get-CodeOnly $path
    $contains = [regex]::Match($code, '(?ms)^Public Function ContainsPoint\b.*?^End Function')
    Assert-Contract "preparedGeometry.$path" ($contains.Success -and
        $contains.Value -match 'If Not IsValid\(message\) Then Exit Function' -and
        $contains.Value -notmatch 'ValidateCurrentGeometry|RebuildContour|OpeningFitsInsideOuter' -and
        $code -match 'mValidationReady = False' -and $code -match 'mGeometryValid = IsValid\(message\)') 'Initialize validates new parameters; hot point queries use the prepared contour'
}
$geometryTypes = Get-CodeOnly 'src/Common/modGeometryTypes.bas'
Assert-Contract 'stableGeometryPiLiteral' ($geometryTypes -match 'GEOM_PI As Double = 3\.14159265358979 \+ 3\.10862446895044E-15') 'Same full-precision Double before and after VBE literal canonicalization'

# Начальная ширина принадлежит сборке, не повторному расчету. Runtime-тест
# дополнительно проверяет пользовательские ширины после Clear/Write/reopen.
foreach ($path in @('src/Excel/CBatchResultWriter.cls', 'src/Excel/CStrengthSummaryWriter.cls',
        'src/Excel/CCrackSummaryWriter.cls', 'src/Excel/CStabilitySummaryWriter.cls',
        'src/Excel/CNDMResultsWriter.cls')) {
    $code = Get-CodeOnly $path
    Assert-Contract "preserveColumnWidths.$path" ($code -notmatch '\.(?:ColumnWidth|StandardWidth)\s*=') 'Calculation output does not resize worksheet columns'
    Assert-Contract "noTemporaryWriterTrace.$path" ($code -notmatch 'Audit03Trace|Win32_Process|PrivatePageCount') 'Temporary profiling API absent in production writer'
}
$build = Get-CodeOnly 'tools/build_workbook/Build-Workbook.ps1'
Assert-Contract 'columnWidthOwnedByBuild' ($build -match '\$results\.Range\("A:CF"\)\.ColumnWidth\s*=\s*10' -and
    $build -match '\$results\.Columns\.Item\(1\)\.ColumnWidth\s*=\s*15' -and
    $build -match '\$results\.Columns\.Item\(2\)\.ColumnWidth\s*=\s*18' -and
    $build -match '\$results\.Columns\.Item\(3\)\.ColumnWidth\s*=\s*21') 'Initial Results widths belong only to workbook construction'
Assert-Contract 'snapshotTitleBoundedInBuild' ($build -match '\$results\.Range\("A154:BJ154"\)\.Interior\.Color' -and
    $build -notmatch '\$results\.Rows\.Item\(154\)\.Interior') 'Snapshot title fill does not format a whole worksheet row'
$snapshot = Get-CodeOnly 'src/Excel/CNDMResultsWriter.cls'
Assert-Contract 'snapshotTitleBoundedAtRuntime' ($snapshot -match 'Set title = SnapshotTitleRange\(workbook\)' -and
    $snapshot -notmatch 'Worksheet\.Rows\.Item\([^\r\n]+\)\.(?:Clear|Interior)') 'Snapshot title write/clear preserves cells outside the snapshot span'

$exportFile = (Resolve-Path -LiteralPath (Join-Path $root $ExportPath)).Path
$export = Get-Content -LiteralPath $exportFile -Raw -Encoding UTF8
$components = @{}
foreach ($match in [regex]::Matches($export, '(?ms)^COMPONENT: ([^\r\n]+)\r?\nTYPE: \d+\r?\nLINES: \d+\r?\n=+\r?\n(.*?)(?=^=+\r?\nCOMPONENT:|\z)')) {
    $components[$match.Groups[1].Value] = Get-VbeCanonicalBody $match.Groups[2].Value
}
$matched = 0
foreach ($file in $sources) {
    $source = Get-Content -LiteralPath $file.FullName -Raw -Encoding UTF8
    if ($source -notmatch 'Attribute\s+VB_Name\s*=\s*"([^"]+)"') { throw "Missing VB_Name: $file" }
    $name = $Matches[1]
    $body = Get-VbeCanonicalBody $source
    $same = $components.ContainsKey($name) -and [string]::Equals($body, $components[$name], [StringComparison]::OrdinalIgnoreCase)
    Assert-Contract "export.$name" $same 'Same source/VBE body, only export representation normalized'
    if ($same) { $matched++ }
}
Assert-Contract 'noDeletedExportModules' (-not $components.ContainsKey('CBatchStatusPolicy') -and -not $components.ContainsKey('CCrackWidthFormulaCalculator')) 'Deleted classes absent in actual workbook export'
$lines.Add("ARTIFACT|$ExportPath|SHA256=$((Get-FileHash -LiteralPath $exportFile -Algorithm SHA256).Hash)")
$lines.Add("TOTAL_AUDIT03_SOURCE_CONTRACTS: matchingModules=$matched; sourceModules=$($sources.Count); failed=$failed")
$reportFile = Join-Path $root $ReportPath
New-Item -ItemType Directory -Force -Path (Split-Path -Parent $reportFile) | Out-Null
$lines | Set-Content -LiteralPath $reportFile -Encoding UTF8
$lines
if ($failed) { exit 1 }
