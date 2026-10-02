# Converts the controlled benchmark record format to typed metrics and aggregates.
# Raw logs remain the evidence; no assertion failures are discarded.
param(
    [Parameter(Mandatory=$true)][string]$ReportPath,
    [Parameter(Mandatory=$true)][string]$OutputPrefix
)
$ErrorActionPreference = 'Stop'
$records = New-Object System.Collections.Generic.List[object]
$raw = Get-Content -LiteralPath $ReportPath -Encoding UTF8
foreach ($line in $raw) {
    if ($line -notmatch '^PERF\|([^|]+)\|run=(\d+)\|(.+)$') { continue }
    $record = [ordered]@{ Version = $Matches[1]; Run = [int]$Matches[2] }
    foreach ($field in $Matches[3] -split '; ') {
        $parts = $field.Split('=', 2)
        if ($parts.Length -ne 2) { throw "Malformed benchmark field: $field" }
        $value = $parts[1]
        $number = 0.0
        if ([double]::TryParse($value, [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$number)) {
            $record[$parts[0]] = $number
        } else {
            $record[$parts[0]] = $value
        }
    }
    $records.Add([pscustomobject]$record)
}
if ($records.Count -eq 0) { throw 'No benchmark records.' }
if (@($raw | Where-Object { $_ -match '^SCRIPT ERROR|^RUNTIME ERROR|^WATCHDOG_FAILURE|^FAIL:' }).Count) { throw 'Benchmark has a failure.' }
if (@($records | Where-Object failed -ne 0).Count) { throw 'Benchmark assertions failed.' }
$summary = New-Object System.Collections.Generic.List[object]
foreach ($group in $records | Group-Object Version, case) {
    if ($group.Count -lt 5) { throw "Fewer than five repeats: $($group.Name)" }
    $item = [ordered]@{ Version = $group.Group[0].Version; Case = $group.Group[0].case; Repeats = $group.Count }
    foreach ($metric in @('elapsedSec', 'geometrySec', 'coreSec', 'packagingSec', 'writeSec', 'solves', 'probes', 'planeProbes', 'iterations', 'retries', 'retryFormatCalls', 'probeReuse', 'stateReuse', 'duplicateSolveAttempts', 'unconvergedProbes', 'unconfirmedPhysical')) {
        if ($group.Group[0].PSObject.Properties.Name -notcontains $metric) { continue }
        $values = @($group.Group | ForEach-Object { [double]$_.$metric } | Sort-Object)
        $middle = [int][math]::Floor($values.Count / 2)
        $median = $values[$middle]
        if ($values.Count % 2 -eq 0) { $median = ($values[$middle - 1] + $values[$middle]) / 2 }
        $item[$metric + 'Median'] = $median
        $item[$metric + 'Min'] = $values[0]
        $item[$metric + 'P95'] = $values[[int][math]::Ceiling($values.Count * 0.95) - 1]
    }
    $item['ModelSizes'] = @($group.Group.modelSizes | Sort-Object -Unique) -join '; '
    $summary.Add([pscustomobject]$item)
}
$records | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath ($OutputPrefix + '_records.json') -Encoding UTF8
$summary | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath ($OutputPrefix + '_summary.json') -Encoding UTF8
$summary | Export-Csv -LiteralPath ($OutputPrefix + '_summary.csv') -Encoding UTF8 -NoTypeInformation
Write-Output "PERFORMANCE_SUMMARY: records=$($records.Count); groups=$($summary.Count); repeats>=5; failed=0"
