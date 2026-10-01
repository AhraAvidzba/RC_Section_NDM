# Сверяет исходные AssertClose по процедуре и выражению test-ID.
# Изменение пути actual разрешено, expected/tolerance должны сохраниться.
# Это дополнение к выполненным VBA suites, не замена численной проверки.
param(
    [string]$BaselineRoot = 'C:\Users\avidzba\AppData\Local\Temp\RC_NDM_Audit02_Baseline_855626e6',
    [string]$ReportPath = 'docs/regression/Audit02/expected_preservation_2026-10-01.txt'
)
$ErrorActionPreference = 'Stop'
$root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$lines = New-Object System.Collections.Generic.List[string]

# Разделяет аргументы VBA с учетом строк и скобок вложенных вызовов.
function Split-VbaArguments([string]$value) {
    $items = New-Object System.Collections.Generic.List[string]
    $quoted = $false
    $depth = 0
    $start = 0
    for ($i = 0; $i -lt $value.Length; $i++) {
        $ch = $value[$i]
        if ($ch -eq '"') {
            if ($quoted -and $i + 1 -lt $value.Length -and $value[$i+1] -eq '"') { $i++; continue }
            $quoted = -not $quoted
        } elseif (-not $quoted) {
            if ($ch -eq '(') { $depth++ }
            elseif ($ch -eq ')') { $depth-- }
            elseif ($ch -eq ',' -and $depth -eq 0) {
                $items.Add($value.Substring($start, $i-$start).Trim())
                $start = $i+1
            }
        }
    }
    $items.Add($value.Substring($start).Trim())
    return $items.ToArray()
}

# Читает полные AssertClose statements, включая стандартные continuation.
function Get-CloseAssertions([string]$path) {
    $result = @{}
    $occurrences = @{}
    $procedure = ''
    $statement = ''
    foreach ($line in Get-Content -LiteralPath $path -Encoding UTF8) {
        if ($line -match '^\s*(Public|Private)\s+(Sub|Function)\s+(\w+)') { $procedure = $Matches[3] }
        if (-not $statement -and $line -notmatch '^\s*AssertClose\s+') { continue }
        $statement += ' ' + $line.Trim()
        if ($statement -match '\s+_$') { $statement = $statement.Substring(0,$statement.Length-1); continue }
        $arguments = @(Split-VbaArguments ($statement.Trim() -replace '^AssertClose\s+', ''))
        if ($arguments.Count -ne 5) { throw "Не распознано AssertClose ($path): $statement" }
        $key = $procedure + '|' + ($arguments[1] -replace '\s+', ' ')
        if (-not $occurrences.ContainsKey($key)) { $occurrences[$key] = 0 }
        $occurrences[$key]++
        $key += '|' + $occurrences[$key]
        $result[$key] = [pscustomobject]@{ Expected=$arguments[3]; Tolerance=$arguments[4] }
        $statement = ''
    }
    return $result
}

$checked = 0
$failed = 0
$migrated = 0
# Только явно проверенные замены прежнего scalar API на тот же typed источник.
# Произвольное новое выражение либо изменение tolerance не разрешается.
$apiMigrations = @{
    'batch.MaxConcreteCompressionStress(1) / batch.LongitudinalCrackRbMc2(1)' = 'batch.ResultAt(1).CrackResult.Longitudinal.MaxCompressionStress / batch.ResultAt(1).CrackResult.Longitudinal.RbMc2'
    'batch.Epsilon0(1)' = 'batch.ResultAt(1).StrengthResult.DirectState.StateResult.Epsilon0'
    'batch.KappaX(1)' = 'batch.ResultAt(1).StrengthResult.DirectState.StateResult.KappaX'
    'batch.KappaY(1)' = 'batch.ResultAt(1).StrengthResult.DirectState.StateResult.KappaY'
    'batch.Mxint(1)' = 'batch.ResultAt(1).StrengthResult.DirectState.StateResult.Mxint'
    'batch.Myint(1)' = 'batch.ResultAt(1).StrengthResult.DirectState.StateResult.Myint'
    'batch.Nint(1)' = 'batch.ResultAt(1).StrengthResult.DirectState.StateResult.Nint'
    'batch.StabilityPhiLTable1(1)' = 'batch.ResultAt(1).StabilityResult.PhiLTable1'
    'Abs(crackHigh.Mcrc)' = 'Abs(crackHigh.FormationResult.Mcrc)'
    'Abs(crackLow.Mcrc)' = 'Abs(crackLow.FormationResult.Mcrc)'
    'crack.Ncrc' = 'crack.FormationResult.Ncrc'
    'crackAutoN.Ncrc' = 'crackAutoN.FormationResult.Ncrc'
    'crackN.Ncrc' = 'crackN.FormationResult.Ncrc'
    'Abs(crackNMxy.LambdaCrc * -15000000#)' = 'Abs(crackNMxy.FormationResult.LambdaCrc * -15000000#)'
    'crackNMxy.LambdaCrc * -20000#' = 'crackNMxy.FormationResult.LambdaCrc * -20000#'
}
$procedures = 0
foreach ($file in Get-ChildItem -LiteralPath (Join-Path $BaselineRoot 'tests') -Filter '*.bas') {
    $currentPath = Join-Path (Join-Path $root 'tests') $file.Name
    $before = Get-CloseAssertions $file.FullName
    $after = Get-CloseAssertions $currentPath
    foreach ($key in $before.Keys | Sort-Object) {
        if (-not $after.ContainsKey($key)) { $failed++; $lines.Add("MISSING|$($file.Name)|$key"); continue }
        $a = $before[$key]
        $b = $after[$key]
        $sameTolerance = ($a.Tolerance -replace '\s+', '') -ceq ($b.Tolerance -replace '\s+', '')
        if ($sameTolerance -and $apiMigrations.ContainsKey($a.Expected) -and
            $apiMigrations[$a.Expected] -ceq $b.Expected) {
            $migrated++
            $lines.Add("API_ONLY|$($file.Name)|$key|oldExpected=$($a.Expected)|newExpected=$($b.Expected)|unchangedTolerance=$($a.Tolerance)")
        } elseif (($a.Expected -replace '\s+', '') -cne ($b.Expected -replace '\s+', '') -or -not $sameTolerance) {
            $failed++
            $lines.Add("CHANGED|$($file.Name)|$key|oldExpected=$($a.Expected)|newExpected=$($b.Expected)|oldTolerance=$($a.Tolerance)|newTolerance=$($b.Tolerance)")
        } else {
            $checked++
        }
    }
    $currentText = Get-Content -LiteralPath $currentPath -Encoding UTF8 -Raw
    foreach ($match in [regex]::Matches((Get-Content -LiteralPath $file.FullName -Encoding UTF8 -Raw), '(?m)^Private Sub (Test\w+)\(')) {
        $name = $match.Groups[1].Value
        if ($currentText -notmatch "(?m)^Private Sub $name\(") { $failed++; $lines.Add("MISSING_TEST_PROCEDURE|$($file.Name)|$name") }
        else { $procedures++ }
    }
    $lines.Add("FILE|$($file.Name)|baseline=$($before.Count)|current=$($after.Count)")
}
$lines.Add("TOTAL_EXPECTED_PRESERVATION: unchanged=$checked; reviewedApiOnly=$migrated; changedOrMissing=$failed; originalTestProcedures=$procedures")
$lines | Set-Content -LiteralPath (Join-Path $root $ReportPath) -Encoding UTF8
$lines
if ($failed) { exit 1 }
