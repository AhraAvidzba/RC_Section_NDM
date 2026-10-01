# Однократная механическая миграция Audit02: потребители скалярного Batch API
# переходят к каноническому result-tree. Скрипт не меняет формулы и expected.
param([switch]$Apply)
$ErrorActionPreference = 'Stop'
$root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$batchPath = Join-Path $root 'src/Batch/CBatchSectionCalculator.cls'
$combinationPath = Join-Path $root 'src/Batch/CCombinationResult.cls'
$batch = [IO.File]::ReadAllText($batchPath)
$combination = [IO.File]::ReadAllText($combinationPath)
$fieldMap = @{}

# Карта строится из существующих присваиваний, а не из сходства имен.
foreach ($entry in @(
    @('capacity', 'StrengthResult.Capacity'),
    @('formation', 'CrackResult.Formation'),
    @('widthResult', 'CrackResult.Width'),
    @('longitudinal', 'CrackResult.Longitudinal')
)) {
    $pattern = '(?m)^\s*(\w+) = ' + [regex]::Escape($entry[0]) + '\.(\w+)\s*$'
    foreach ($match in [regex]::Matches($combination, $pattern)) {
        $fieldMap[$match.Groups[1].Value] = $entry[1] + '.' + $match.Groups[2].Value
    }
}
$stabilityFields = [regex]::Matches([IO.File]::ReadAllText((Join-Path $root 'src/Stability/CStabilityResult.cls')),
    '(?m)^Public (\w+) As ') | ForEach-Object { $_.Groups[1].Value }
foreach ($field in $stabilityFields) {
    $fieldMap['Stability' + $field] = 'StabilityResult.' + $field
}
$fieldMap['StabilityUtil'] = 'StabilityResult.Utilization'
$fieldMap['MaxConcreteCompressionStress'] = 'CrackResult.Longitudinal.MaxCompressionStress'
foreach ($field in @('Epsilon0','KappaX','KappaY','Nint','Mxint','Myint','ResidualN','ResidualMx','ResidualMy',
    'MinConcreteStrain','MaxConcreteStrain','MinSteelStrain','MaxSteelStrain','StateAvailable','StateConverged')) {
    $fieldMap[$field] = 'StrengthResult.DirectState.' + $field
}
$fieldMap['StateIterations'] = 'StrengthResult.DirectState.Iterations'
$fieldMap['DirectStateStatus'] = 'StrengthResult.DirectState.Status'
$fieldMap['CapacityStatus'] = 'StrengthResult.Capacity.Status'
$fieldMap['CapacityLimitState'] = 'StrengthResult.Capacity.LimitState'
$fieldMap['CapacitySolutionMethod'] = 'StrengthResult.Capacity.SolutionMethod'
$fieldMap['CapacityPathResolved'] = 'StrengthResult.Capacity.PathResolved'
$fieldMap['CrackStatus'] = 'NormalCrackStatus'
$fieldMap['LongitudinalCrackStatus'] = 'CrackResult.Longitudinal.Status'
$fieldMap['StabilityStatus'] = 'StabilityResult.Status'
$fieldMap['InternalStatus'] = 'OverallMeta.InternalStatus'
$fieldMap['ResultCode'] = 'OverallMeta.ResultCode'
$fieldMap['ResultComment'] = 'OverallMeta.ResultComment'

$getterMap = @{}
$members = [regex]::Matches($batch,
    '(?ms)^Public Property Get (\w+)\(ByVal index As Long\)[^\r\n]*\r?\n.*?^End Property\r?\n')
foreach ($member in $members) {
    $name = $member.Groups[1].Value
    $field = [regex]::Match($member.Value, 'mResults\(index\)\.(\w+)')
    if (-not $field.Success) { continue }
    $value = $field.Groups[1].Value
    if ($name -in @('Status','OverallStatus')) { $value = 'Status' }
    elseif ($name -eq 'DirectStateStatus') { $value = 'StrengthResult.DirectState.Status' }
    elseif ($name -eq 'CapacityStatus') { $value = 'StrengthResult.Capacity.Status' }
    elseif ($name -eq 'CrackStatus') { $value = 'NormalCrackStatus' }
    elseif ($name -eq 'LongitudinalCrackStatus') { $value = 'CrackResult.Longitudinal.Status' }
    elseif ($name -eq 'StabilityStatus') { $value = 'StabilityResult.Status' }
    elseif ($name -eq 'CrackSummaryStatus') { $value = 'CrackSummaryStatus' }
    elseif ($fieldMap.ContainsKey($value)) { $value = $fieldMap[$value] }
    $getterMap[$name] = $value
}

# Разбирает скобки VBA-вызова, включая вложенные функции и строковые литералы.
function Convert-IndexedCalls {
    param([string]$Text, [string]$Receiver, [hashtable]$Map, [string]$ResultReceiver)
    $pattern = '(?i)\b' + [regex]::Escape($Receiver) + '\.(\w+)\('
    $matches = [regex]::Matches($Text, $pattern)
    for ($i = $matches.Count - 1; $i -ge 0; $i--) {
        $match = $matches[$i]
        $name = $match.Groups[1].Value
        if (-not $Map.ContainsKey($name)) { continue }
        $start = $match.Index + $match.Length
        $depth = 1; $quoted = $false; $end = $start
        for (; $end -lt $Text.Length; $end++) {
            $char = $Text[$end]
            if ($char -eq '"') {
                if ($quoted -and $end + 1 -lt $Text.Length -and $Text[$end + 1] -eq '"') { $end++; continue }
                $quoted = -not $quoted
            } elseif (-not $quoted) {
                if ($char -eq '(') { $depth++ }
                elseif ($char -eq ')') { $depth--; if ($depth -eq 0) { break } }
            }
        }
        if ($depth -ne 0) { throw "Незакрытый вызов $Receiver.$name" }
        $argument = $Text.Substring($start, $end - $start)
        $replacement = $ResultReceiver + '(' + $argument + ').' + $Map[$name]
        $Text = $Text.Remove($match.Index, $end - $match.Index + 1).Insert($match.Index, $replacement)
    }
    return $Text
}

$files = @(Get-ChildItem -LiteralPath (Join-Path $root 'src'),(Join-Path $root 'tests') -Recurse -File |
    Where-Object { $_.Extension -in @('.cls','.bas') })
$changed = 0
foreach ($file in $files) {
    if ($file.FullName -eq $combinationPath) { continue }
    $text = [IO.File]::ReadAllText($file.FullName)
    $old = $text
    $receivers = [regex]::Matches($text, '(?i)\b(\w+) As CBatchSectionCalculator\b') |
        ForEach-Object { $_.Groups[1].Value } | Sort-Object -Unique
    foreach ($receiver in $receivers) {
        $text = Convert-IndexedCalls $text $receiver $getterMap ($receiver + '.ResultAt')
    }
    $resultReceivers = [regex]::Matches($text, '(?i)\b(\w+) As CCombinationResult\b') |
        ForEach-Object { $_.Groups[1].Value } | Sort-Object -Unique
    foreach ($receiver in $resultReceivers) {
        foreach ($field in $fieldMap.Keys) {
            $text = [regex]::Replace($text, '(?i)\b' + [regex]::Escape($receiver) + '\.' +
                [regex]::Escape($field) + '\b', $receiver + '.' + $fieldMap[$field])
        }
    }
    if ($file.FullName -eq $batchPath) {
        foreach ($member in $members) {
            if ($getterMap.ContainsKey($member.Groups[1].Value)) { $text = $text.Replace($member.Value, '') }
        }
        foreach ($field in $fieldMap.Keys) {
            $text = [regex]::Replace($text, '(?i)(mResults\([^\r\n]*?\))\.' + [regex]::Escape($field) + '\b',
                '$1.' + $fieldMap[$field])
        }
    }
    if ($text -ne $old) {
        $changed++
        Write-Output ($file.FullName.Substring($root.Length + 1))
        if ($Apply) { [IO.File]::WriteAllText($file.FullName, $text, (New-Object Text.UTF8Encoding($false))) }
    }
}
Write-Output "Changed files: $changed; forwarding getters: $($getterMap.Count); flat mappings: $($fieldMap.Count)"
