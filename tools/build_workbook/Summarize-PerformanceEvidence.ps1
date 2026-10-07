# Read-only aggregation of accepted alternating A/B series; pilots and wrong-argument
# folders are intentionally excluded. Full macro phases are not inferred as zero cost.
param([string]$Directory='docs/regression/Performance')
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$evidence=Join-Path $root $Directory
$names=@('SolverMeasurements','DirectMeasurements','SingleFullRepeat','Crack35Measurements',
    'Crack63MeasurementsV2','Stability35MeasurementsV2','Stability63MeasurementsV2',
    'CapacityAutoMeasurementsV2','CapacityMultiplierMeasurementsV2','MixedMeasurementsV2',
    'ReverseMeasurementsV2','ShuffleMeasurementsV2','CircleMeasurementsV2',
    'RoundedMeasurementsV2','OffMeasurementsV2','StorageQPCMeasurementsV2',
    'ExportMeasurements','ImportedMeasurements','UserInputsMeasurements')
function Statistics($values) {
    $sorted=@($values | Sort-Object)
    if($sorted.Count -lt 5) {throw 'Fewer than five observations.'}
    $middle=[int][Math]::Floor($sorted.Count/2)
    $median=[double]$sorted[$middle]
    if($sorted.Count%2 -eq 0) {$median=([double]$sorted[$middle-1]+[double]$sorted[$middle])/2}
    return [ordered]@{N=$sorted.Count;Median=$median;Min=[double]$sorted[0];Max=[double]$sorted[-1]}
}
$result=@()
foreach($name in $names) {
    $path=Join-Path $evidence $name
    if(-not(Test-Path -LiteralPath (Join-Path $path 'Summary.json'))) {continue}
    $raw=Get-Content -LiteralPath (Join-Path $path 'Raw.json') -Raw | ConvertFrom-Json
    $raw=@($raw)
    foreach($row in $raw) {
        if($row.PSObject.Properties.Name -notcontains 'RunWithinOpen') {$row | Add-Member -NotePropertyName RunWithinOpen -NotePropertyValue 1}
    }
    foreach($case in @($raw.Case | Select-Object -Unique)) {
        $rows=@($raw | Where-Object {$_.Case -eq $case})
        $reference=$rows[0].Values
        foreach($row in $rows) {
            if($row.Pilot) {throw "Pilot used as acceptance: $name"}
            if(@($row.Values.PSObject.Properties).Count -ne @($reference.PSObject.Properties).Count) {throw 'A/B fields differ.'}
            foreach($property in $reference.PSObject.Properties) {
                if($row.Values.($property.Name) -cne $property.Value) {throw "Exact mismatch: $name/$case/$($property.Name)"}
            }
        }
        foreach($run in @($rows.RunWithinOpen | Select-Object -Unique)) {
            $pair=[ordered]@{Evidence=$name;Case=$case;RunWithinOpen=$run;ExactValues=$true}
            foreach($version in @('A','B')) {
                $selected=@($rows | Where-Object {$_.Version -eq $version -and $_.RunWithinOpen -eq $run})
                $metrics=[ordered]@{Total=(Statistics $selected.Seconds)}
                if($case -notmatch '-Full-') {
                    foreach($key in @('preparation','core','clear','packing','summary','snapshot','plot')) {
                        if($selected[0].Technical.PSObject.Properties.Name -contains $key) {
                            $values=@($selected | ForEach-Object {[double]::Parse($_.Technical.$key,[Globalization.CultureInfo]::InvariantCulture)})
                            $metrics[$key]=Statistics $values
                        }
                    }
                }
                $metrics['PrivateBytesBefore']=Statistics @($selected | ForEach-Object {$_.MemoryBefore.PrivateBytes})
                $metrics['PrivateBytesAfter']=Statistics @($selected | ForEach-Object {$_.MemoryAfter.PrivateBytes})
                $metrics['PeakWorkingSet']=Statistics @($selected | ForEach-Object {$_.MemoryAfter.PeakWorkingSet})
                $metrics['TechnicalSamples']=@($selected | ForEach-Object {$_.Technical})
                $pair[$version]=$metrics
            }
            $pair['TimeReductionPercent']=100*(1-$pair.B.Total.Median/$pair.A.Total.Median)
            $result += $pair
        }
    }
}
$result | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath (Join-Path $evidence 'AcceptedMeasurements.json') -Encoding UTF8
$lines=@('# Accepted Performance Measurements','',
    'A = frozen production baseline; B = candidate. Seconds are median [min; max], five alternating A/B observations per row.',
    'Full phases are NotMeasured; phase metrics belong only to separate Staged runs. Memory includes Excel allocator and benchmark fingerprint marshaling.',
    '', '| Evidence / Case | Run | A, s | B, s | Time Reduction | Exact |', '|---|---:|---:|---:|---:|---|')
foreach($row in $result) {
    $a=$row.A.Total;$b=$row.B.Total
    $lines += '| {0} / {1} | {2} | {3:F6} [{4:F6}; {5:F6}] | {6:F6} [{7:F6}; {8:F6}] | {9:F2}% | Yes |' -f $row.Evidence,$row.Case,$row.RunWithinOpen,$a.Median,$a.Min,$a.Max,$b.Median,$b.Min,$b.Max,$row.TimeReductionPercent
}
$lines | Set-Content -LiteralPath (Join-Path $evidence 'AcceptedMeasurements.md') -Encoding UTF8
$index=@()
$indexedFolders=@($names)+@('DirectedFinal','IsolationFinal','FinalReaders','GuardFixedTests',
    'NumericTests','StorageTests','FinalCandidateV2','FinalBaselineV2','CompiledCandidate',
    'PreparedCandidate','NumericCandidate','StorageCandidate','ReleaseV2','CompiledFreshBuildV2')
foreach($folder in $indexedFolders) {
    $path=Join-Path $evidence $folder
    if(-not(Test-Path -LiteralPath $path)) {continue}
    foreach($file in @(Get-ChildItem -LiteralPath $path -File | Sort-Object Name)) {
        if($file.Extension -eq '.xlsm') {continue}
        $index += [ordered]@{Path=$folder+'/'+$file.Name;Bytes=$file.Length;
            SHA256=(Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash}
    }
}
[ordered]@{MeasurementFolders=$names;Cases=$result.Count;Files=$index;
    Note='Full fingerprints and private fixtures remain in the workspace. The index records their evidence paths/hashes; private XLSM copies are excluded. Pilots/wrong-argument series are not accepted measurements.'} |
    ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $evidence 'EvidenceManifest.json') -Encoding UTF8
Write-Output "Accepted cases: $($result.Count); exact A/B fields verified."
