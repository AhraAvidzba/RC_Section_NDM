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
$difference = @(Compare-Object ($beforeClasses | Sort-Object) ($production.Name | Sort-Object))
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
