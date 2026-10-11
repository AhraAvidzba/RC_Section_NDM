param([string]$ReportDirectory='docs/regression/PlateauSlope_2026-10-11/ResultsV2')
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$directory=Join-Path $root $ReportDirectory
$culture=[Globalization.CultureInfo]::InvariantCulture
function Read-Cases([string]$Path) {
    foreach ($line in Get-Content -LiteralPath $Path -Encoding UTF8) {
        if (-not $line.StartsWith('PLATEAU_CASE|')) { continue }
        $record=[ordered]@{}
        foreach ($field in ($line -split '\|' | Select-Object -Skip 1)) {
            $pair=$field -split '=',2
            $record[$pair[0]]=$pair[1]
        }
        [pscustomobject]$record
    }
}
function Numeric($Value) { return [double]::Parse([string]$Value,$culture) }
function Relative($Before,$After) {
    $a=Numeric $Before; $b=Numeric $After
    if ($a -eq 0) { if ($b -eq 0) { return 0.0 } else { return $null } }
    return 100.0*[Math]::Abs($b-$a)/[Math]::Abs($a)
}
$pairs=[Collections.Generic.List[object]]::new()
$summaries=[Collections.Generic.List[object]]::new()
foreach ($offFile in Get-ChildItem -LiteralPath $directory -Filter '*_SlopeOff.txt') {
    $onPath=Join-Path $directory ($offFile.Name.Replace('_SlopeOff','_SlopeOn'))
    if (-not (Test-Path -LiteralPath $onPath)) { continue }
    $off=@(Read-Cases $offFile.FullName); $on=@(Read-Cases $onPath)
    if ($off.Count -eq 0 -or $off.Count -ne $on.Count) { Write-Warning ('Incomplete pair: '+$offFile.Name); continue }
    $group=$offFile.BaseName.Replace('_SlopeOff','')
    $groupPairs=[Collections.Generic.List[object]]::new()
    for ($i=0; $i -lt $off.Count; $i++) {
        $a=$off[$i]; $b=$on[$i]
        if ($a.id -cne $b.id) { throw 'Case ordering mismatch.' }
        $record=[ordered]@{Group=$group; Case=$a.id; Sequence=$i+1;
            OffSeconds=(Numeric $a.seconds); OnSeconds=(Numeric $b.seconds);
            OffIterations=(Numeric $a.stateIterations)+(Numeric $a.ultimateIterations);
            OnIterations=(Numeric $b.stateIterations)+(Numeric $b.ultimateIterations);
            OffCalls=(Numeric $a.solverCalls); OnCalls=(Numeric $b.solverCalls);
            OffNumFail=(Numeric $a.numFailStates); OnNumFail=(Numeric $b.numFailStates);
            OffNumFailResults=(Numeric $a.numFailResults); OnNumFailResults=(Numeric $b.numFailResults);
            OffStatus=$a.overall; OnStatus=$b.overall; OffCapacityStatus=$a.capacityStatus; OnCapacityStatus=$b.capacityStatus;
            OffFormationStatus=$a.formationStatus; OnFormationStatus=$b.formationStatus;
            OffWidthStatus=$a.widthStatus; OnWidthStatus=$b.widthStatus;
            OffLambda=(Numeric $a.lambdaUlt); OnLambda=(Numeric $b.lambdaUlt);
            LambdaDeltaPercent=$(if ($a.capacityStatus -ne 'NumFail' -and $b.capacityStatus -ne 'NumFail') { Relative $a.lambdaUlt $b.lambdaUlt } else {$null});
            NultDeltaPercent=$(if ($a.capacityStatus -ne 'NumFail' -and $b.capacityStatus -ne 'NumFail') {Relative $a.Nult $b.Nult} else {$null});
            MxUltDeltaPercent=$(if ($a.capacityStatus -ne 'NumFail' -and $b.capacityStatus -ne 'NumFail') {Relative $a.MxUlt $b.MxUlt} else {$null});
            MyUltDeltaPercent=$(if ($a.capacityStatus -ne 'NumFail' -and $b.capacityStatus -ne 'NumFail') {Relative $a.MyUlt $b.MyUlt} else {$null});
            McrcDeltaPercent=$(if ($a.hasMcrc -eq 'True' -and $b.hasMcrc -eq 'True') {Relative $a.Mcrc $b.Mcrc} else {$null});
            WidthDeltaPercent=$(if ($a.widthStatus -eq $b.widthStatus) {Relative $a.width $b.width} else {$null});
            ArDeltaPercent=$(if ($a.widthStatus -eq $b.widthStatus) {Relative $a.Ar $b.Ar} else {$null});
            SigmaSDeltaPercent=$(if ($a.widthStatus -eq $b.widthStatus) {Relative $a.sigmaS $b.sigmaS} else {$null});
            AsDeltaPercent=$(if ($a.widthStatus -eq $b.widthStatus) {Relative $a.As $b.As} else {$null});
            OffWidth=(Numeric $a.width); OnWidth=(Numeric $b.width); OffAr=(Numeric $a.Ar); OnAr=(Numeric $b.Ar);
            OffAnchor=$a.anchor; OnAnchor=$b.anchor;
            OffActiveSet=(Numeric $a.activeSetIterations); OnActiveSet=(Numeric $b.activeSetIterations);
            OffRetries=(Numeric $a.retryCalls); OnRetries=(Numeric $b.retryCalls);
            OffFallbacks=(Numeric $a.multiplierFallbacks); OnFallbacks=(Numeric $b.multiplierFallbacks);
            OffExtensionGuesses=(Numeric $a.extensionGuesses); OnExtensionGuesses=(Numeric $b.extensionGuesses);
            OffCapacityRetries=(Numeric $a.capacityRetries); OnCapacityRetries=(Numeric $b.capacityRetries);
            OffFormationFullStarts=(Numeric $a.formationFullStarts); OnFormationFullStarts=(Numeric $b.formationFullStarts);
            OffWarmRestarts=(Numeric $a.warmRestarts); OnWarmRestarts=(Numeric $b.warmRestarts);
            OffProjectedRecoveries=(Numeric $a.projectedRecoveries); OnProjectedRecoveries=(Numeric $b.projectedRecoveries)}
        $row=[pscustomobject]$record; $pairs.Add($row); $groupPairs.Add($row)
    }
    $sumOff=($groupPairs | Measure-Object OffSeconds -Sum).Sum
    $sumOn=($groupPairs | Measure-Object OnSeconds -Sum).Sum
    $summaries.Add([pscustomobject][ordered]@{Group=$group; Cases=$off.Count; OffSeconds=$sumOff; OnSeconds=$sumOn;
        ChangePercent=100.0*($sumOn/$sumOff-1.0);
        OffIterations=($groupPairs | Measure-Object OffIterations -Sum).Sum; OnIterations=($groupPairs | Measure-Object OnIterations -Sum).Sum;
        OffCalls=($groupPairs | Measure-Object OffCalls -Sum).Sum; OnCalls=($groupPairs | Measure-Object OnCalls -Sum).Sum;
        OffNumFail=($groupPairs | Measure-Object OffNumFail -Sum).Sum; OnNumFail=($groupPairs | Measure-Object OnNumFail -Sum).Sum;
        OffNumFailResults=($groupPairs | Measure-Object OffNumFailResults -Sum).Sum; OnNumFailResults=($groupPairs | Measure-Object OnNumFailResults -Sum).Sum;
        OffActiveSet=($groupPairs | Measure-Object OffActiveSet -Sum).Sum; OnActiveSet=($groupPairs | Measure-Object OnActiveSet -Sum).Sum;
        OffRetries=($groupPairs | Measure-Object OffRetries -Sum).Sum; OnRetries=($groupPairs | Measure-Object OnRetries -Sum).Sum;
        OffExtensionGuesses=($groupPairs | Measure-Object OffExtensionGuesses -Sum).Sum; OnExtensionGuesses=($groupPairs | Measure-Object OnExtensionGuesses -Sum).Sum;
        OffFallbacks=($groupPairs | Measure-Object OffFallbacks -Sum).Sum; OnFallbacks=($groupPairs | Measure-Object OnFallbacks -Sum).Sum;
        MaxLambdaDeltaPercent=($groupPairs | Measure-Object LambdaDeltaPercent -Maximum).Maximum;
        MaxMcrcDeltaPercent=($groupPairs | Measure-Object McrcDeltaPercent -Maximum).Maximum;
        MaxWidthDeltaPercent=($groupPairs | Measure-Object WidthDeltaPercent -Maximum).Maximum;
        MaxArDeltaPercent=($groupPairs | Measure-Object ArDeltaPercent -Maximum).Maximum;
        MaxSigmaSDeltaPercent=($groupPairs | Measure-Object SigmaSDeltaPercent -Maximum).Maximum})
}
if ($pairs.Count) { $pairs | Export-Csv -LiteralPath (Join-Path $directory 'Comparisons.csv') -NoTypeInformation -Encoding UTF8 }
if ($summaries.Count) {
    $summaries | Export-Csv -LiteralPath (Join-Path $directory 'Summary.csv') -NoTypeInformation -Encoding UTF8
    $summaries | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $directory 'Summary.json') -Encoding UTF8
    $summaries | Format-Table Group,Cases,OffSeconds,OnSeconds,ChangePercent,OffIterations,OnIterations,OffNumFail,OnNumFail -AutoSize
}
Write-Output 'PHYSICAL_OR_STATUS_CHANGES:'
$pairs | Where-Object {$_.LambdaDeltaPercent -gt 2 -or $_.McrcDeltaPercent -gt 2 -or $_.WidthDeltaPercent -gt 2 -or $_.ArDeltaPercent -gt 2 -or
    $_.OffStatus -cne $_.OnStatus -or $_.OffCapacityStatus -cne $_.OnCapacityStatus -or $_.OffFormationStatus -cne $_.OnFormationStatus -or $_.OffWidthStatus -cne $_.OnWidthStatus} |
    Format-Table Group,Case,OffStatus,OnStatus,LambdaDeltaPercent,McrcDeltaPercent,WidthDeltaPercent,OffAnchor,OnAnchor -AutoSize
