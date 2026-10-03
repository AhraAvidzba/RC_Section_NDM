# Сопоставляет численные assertions общих тестов двух выполненных VBA suites.
# Это проверка опубликованных чисел журнала, не замена инженерных oracle,
# компонентных допусков, полного regression или сравнения сохраненных Results.
param(
    [Parameter(Mandatory=$true)][string]$BaselineReport,
    [Parameter(Mandatory=$true)][string]$CurrentReport,
    [Parameter(Mandatory=$true)][string]$ReportPath
)
$ErrorActionPreference = 'Stop'
$culture = [Globalization.CultureInfo]::InvariantCulture

# Повторяющиеся ID различаются номером появления; новые ID не сдвигают
# существующие записи. Извлекается только actual численного assertion.
function Read-NumericAssertions([string]$Path) {
    $values = [ordered]@{}
    $occurrences = @{}
    foreach ($line in [IO.File]::ReadAllLines((Resolve-Path -LiteralPath $Path).Path, [Text.Encoding]::UTF8)) {
        $match = [regex]::Match($line, '^(?:OK|FAIL):\s*(?<id>[^;|]+)(?:;|\|).*?actual\s*[=:]\s*(?<value>[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?)(?=[;|\s]|$)')
        if (-not $match.Success) { continue }
        $id = $match.Groups['id'].Value.Trim()
        $number = 1
        if ($occurrences.ContainsKey($id)) { $number = $occurrences[$id] + 1 }
        $occurrences[$id] = $number
        $key = $id + '#' + $number
        $values[$key] = [pscustomobject]@{
            Text = $match.Groups['value'].Value
            Value = [double]::Parse($match.Groups['value'].Value, $culture)
        }
    }
    if ($values.Count -eq 0) { throw "В журнале нет численных assertions: $Path" }
    return $values
}

$before = Read-NumericAssertions $BaselineReport
$after = Read-NumericAssertions $CurrentReport
$differences = [Collections.Generic.List[object]]::new()
$missing = [Collections.Generic.List[string]]::new()
$compared = 0
foreach ($key in $before.Keys) {
    if (-not $after.Contains($key)) { $missing.Add($key); continue }
    $compared++
    if ($before[$key].Value -ne $after[$key].Value) {
        $differences.Add([pscustomobject]@{
            Assertion = $key
            Baseline = $before[$key].Text
            Current = $after[$key].Text
            AbsoluteDifference = [Math]::Abs($after[$key].Value - $before[$key].Value)
        })
    }
}
$result = [ordered]@{
    BaselineReport = $BaselineReport
    CurrentReport = $CurrentReport
    BaselineSHA256 = (Get-FileHash -LiteralPath $BaselineReport -Algorithm SHA256).Hash
    CurrentSHA256 = (Get-FileHash -LiteralPath $CurrentReport -Algorithm SHA256).Hash
    BaselineNumericAssertions = $before.Count
    CurrentNumericAssertions = $after.Count
    Compared = $compared
    Missing = @($missing)
    Differences = @($differences)
    AcceptanceScope = 'Exact numeric actual-values in shared assertion IDs; no new tolerance'
    Passed = ($missing.Count -eq 0 -and $differences.Count -eq 0)
}
$result | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $ReportPath -Encoding UTF8
Write-Output "SUITE_NUMBER_COMPARISON: compared=$compared; missing=$($missing.Count); differences=$($differences.Count)"
if (-not $result.Passed) { exit 1 }
