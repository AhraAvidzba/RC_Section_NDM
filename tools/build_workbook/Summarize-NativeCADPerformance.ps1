# Reports only accepted native measurements, excluding warmups.
param([string]$Directory='docs/regression/Performance/NativeCADV4')
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$path=Join-Path $root $Directory
$accepted=Get-Content -LiteralPath (Join-Path $path 'Acceptance.json') -Raw | ConvertFrom-Json
if(-not $accepted.Passed -or -not $accepted.ExactABGeometry -or $accepted.Repetitions -lt 5) {throw 'No accepted native series.'}
$records=Get-Content -LiteralPath (Join-Path $path 'Measurements.json') -Raw | ConvertFrom-Json
$records=@($records | Where-Object {-not $_.Warmup})
function Stats($values) {
    $values=@($values | Sort-Object)
    if($values.Count -lt 5) {throw 'At least five repetitions are required.'}
    $middle=[int][Math]::Floor($values.Count/2)
    $median=$values[$middle];if($values.Count%2 -eq 0) {$median=($values[$middle-1]+$values[$middle])/2}
    return [ordered]@{N=$values.Count;Median=$median;Min=$values[0];Max=$values[-1];P95=$values[[Math]::Ceiling(0.95*$values.Count)-1]}
}
$summary=@()
foreach($scenario in @('Grid','Rotated','GridWithOtherEntities')) {
    $a=@($records | Where-Object {$_.Scenario -eq $scenario -and $_.Version -eq 'A'})
    $b=@($records | Where-Object {$_.Scenario -eq $scenario -and $_.Version -eq 'B'})
    foreach($run in $a.Run) {
        $left=@($a | Where-Object Run -eq $run);$right=@($b | Where-Object Run -eq $run)
        if($left.Count -ne 1 -or $right.Count -ne 1 -or $left[0].SignatureSHA256 -cne $right[0].SignatureSHA256) {throw 'Exact pair agreement missing.'}
    }
    foreach($operation in @('Export','Import')) {
        $field=$operation+'Seconds';$sa=Stats $a.$field;$sb=Stats $b.$field
        $summary += [ordered]@{Scenario=$scenario;Operation=$operation;A=$sa;B=$sb;TimeReductionPercent=100*(1-$sb.Median/$sa.Median);SpeedupRatio=$sa.Median/$sb.Median;ExactGeometry=$true}
    }
}
$summary | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $path 'Summary.json') -Encoding UTF8
$summary | ForEach-Object {'{0} {1}: {2:N6} -> {3:N6} sec; reduction={4:N2}%; ratio={5:N3}' -f $_.Scenario,$_.Operation,$_.A.Median,$_.B.Median,$_.TimeReductionPercent,$_.SpeedupRatio}
