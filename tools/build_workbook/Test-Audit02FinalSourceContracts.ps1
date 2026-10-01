# Итоговая статическая проверка границ Audit02 и фактического VBA-export.
# Не запускает и не заменяет Excel-тесты: они подтверждаются отдельными логами.
param(
    [string]$BaselineRoot = 'C:\Users\avidzba\AppData\Local\Temp\RC_NDM_Audit02_Baseline_855626e6',
    [string]$ReportPath = 'docs/regression/Audit02/final_source_contracts_2026-10-01.txt'
)
$ErrorActionPreference = 'Stop'
$root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$lines = New-Object System.Collections.Generic.List[string]
$failed = 0

# Записывает утверждение с конкретным доказательством или причиной отказа.
function Assert-Contract([string]$name, [bool]$passed, [string]$detail) {
    if (-not $passed) { $script:failed++ }
    $lines.Add("CONTRACT|$name|passed=$passed|$detail")
}

# Удаляет только export-header и Attribute строки, как штатный VBA-import.
function Get-SourceBody([string]$text) {
    $result = New-Object System.Collections.Generic.List[string]
    $header = $false
    foreach ($line in $text -split "`r?`n") {
        if ($line -match '^\uFEFF?VERSION\s+') { continue }
        if ($line -match '^\uFEFF?BEGIN\s*$') { $header=$true; continue }
        if ($header) { if ($line -match '^\uFEFF?END\s*$') { $header=$false }; continue }
        if ($line -match '^\uFEFF?Attribute\s+VB_') { continue }
        $result.Add($line.TrimEnd())
    }
    return ($result -join "`n").Trim()
}

# VBA-editor применяет системную ANSI-кодировку и канонизирует Double literals:
# 1e-8 -> 0.00000001, удаляет лишний # и сохраняет 15 значащих цифр.
# Нормализуем именно это представление; строки и числовой текст комментариев
# не переписываем. Изменение произвольной константы проверку не проходит.
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

$production = @(Get-ChildItem -LiteralPath (Join-Path $root 'src') -Recurse -File -Filter '*.cls')
$beforeClasses = @(Get-ChildItem -LiteralPath (Join-Path $BaselineRoot 'src') -Recurse -File -Filter '*.cls' | ForEach-Object { $_.Name })
$classDifference = @(Compare-Object ($beforeClasses | Sort-Object) ($production.Name | Sort-Object))
Assert-Contract 'noNewOrDeletedProductionClasses' ($classDifference.Count -eq 0) "baseline=$($beforeClasses.Count); current=$($production.Count)"

$generic = @('src/Interfaces/ILimitSearchProblem.cls','src/Solver/CLoadMultiplierSearch.cls','src/Solver/CUltimateStrainSearch.cls','src/Solver/CLimitSearchResult.cls')
foreach ($path in $generic) {
    $code = Get-Content -LiteralPath (Join-Path $root $path) -Raw -Encoding UTF8
    Assert-Contract "generic.$path" ($code -notmatch 'DomainContext|As CCrackWidthCalculator|As CCrackFormationCalculator|As CCapacitySolver') 'No DomainContext/downcast/domain-object storage'
}
$formula = Get-Content -LiteralPath (Join-Path $root 'src/Crack/CCrackWidthFormulaCalculator.cls') -Raw -Encoding UTF8
$formulaCode = ($formula -split "`r?`n" | Where-Object { $_ -notmatch "^\s*'" }) -join "`n"
Assert-Contract 'pureCrackWidthFormula' ($formulaCode -notmatch 'CSectionState|CSectionSolver|\brs\w+\b|NumFail|\.Solve') 'Only prepared numeric arguments; no state/status/solve'
foreach ($path in @('src/Crack/CCrackWidthCalculator.cls','src/Crack/CLongitudinalCrackCalculator.cls','src/Stability/CStabilityCalculator.cls')) {
    $code = (Get-Content -LiteralPath (Join-Path $root $path) -Encoding UTF8 | Where-Object { $_ -notmatch "^\s*'" }) -join "`n"
    Assert-Contract "formulaStatuses.$path" ($code -notmatch '\brsNumericalFailure\b|"NumFail"|\.Solve(?:By|With|\b)') 'No equilibrium solve or locally assigned numerical failure'
}

$exportPath = Join-Path $root 'workbook/output/VBA_All_Code.txt'
$export = Get-Content -LiteralPath $exportPath -Raw -Encoding UTF8
$components = @{}
foreach ($match in [regex]::Matches($export, '(?ms)^COMPONENT: ([^\r\n]+)\r?\nTYPE: \d+\r?\nLINES: \d+\r?\n=+\r?\n(.*?)(?=^=+\r?\nCOMPONENT:|\z)')) {
    $components[$match.Groups[1].Value] = Get-VbeCanonicalBody $match.Groups[2].Value
}
$matched = 0
foreach ($file in Get-ChildItem -Path (Join-Path $root 'src'),(Join-Path $root 'tests') -Recurse -File | Where-Object { $_.Extension -in '.cls','.bas' }) {
    $source = Get-Content -LiteralPath $file.FullName -Raw -Encoding UTF8
    if ($source -notmatch 'Attribute\s+VB_Name\s*=\s*"([^"]+)"') { throw "Missing VB_Name: $file" }
    $name = $Matches[1]
    $sourceBody = Get-VbeCanonicalBody $source
    $same = $components.ContainsKey($name) -and [string]::Equals($sourceBody, $components[$name], [StringComparison]::OrdinalIgnoreCase)
    Assert-Contract "export.$name" $same 'Same VBE source body; newline/spacing/casing/CP1251/G15 literal representation normalized'
    if ($same) { $matched++ }
}
foreach ($path in @('workbook/output/RC_Section_NDM.xlsm','workbook/output/VBA_All_Code.txt','docs/regression/Audit02/RC_Section_NDM_clean_final.xlsm')) {
    $hash = (Get-FileHash -LiteralPath (Join-Path $root $path) -Algorithm SHA256).Hash
    $lines.Add("ARTIFACT|$path|SHA256=$hash")
}
$lines.Add("TOTAL_SOURCE_CONTRACTS: matchingModules=$matched; failed=$failed")
$lines | Set-Content -LiteralPath (Join-Path $root $ReportPath) -Encoding UTF8
$lines
if ($failed) { exit 1 }
