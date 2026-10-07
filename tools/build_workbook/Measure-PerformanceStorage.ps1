# Alternating A/B in one owned Excel process; source books remain read-only copies.
# Timers belong to VBA operations, not setup/assertions. Five observations use max, not P95.
param(
    [string]$Directory = 'docs/regression/Performance/StorageMeasurements', [int]$Repeats = 5,
    [ValidateSet('Storage','Solver','Batch','Export')][string]$Group = 'Storage',
    [string]$BaselineWorkbook = 'docs/regression/Performance/StorageBaseline/RC_Section_NDM.xlsm',
    [string]$CandidateWorkbook = 'docs/regression/Performance/StorageCandidate/RC_Section_NDM.xlsm',
    [int]$ModelCount = 50000, [int]$ExtentCount = 100000, [int]$SolverDivisions = 40,
    [string]$Family = 'DIRECT', [string]$Method = 'Newton', [string]$Shape = 'Saved',
    [string]$Order = 'Normal', [string]$Extension = 'Yes', [int[]]$Counts = @(1,10,30),
    [ValidateSet('Staged','Full')][string[]]$BatchModes = @('Staged','Full'), [switch]$Pilot, [int]$RunsPerOpen = 1,
    [switch]$MeasureOverhead, [switch]$KeepInputs
)
$ErrorActionPreference = 'Stop'
if ($Repeats -lt 5 -and -not $Pilot) { throw 'At least five A/B observations are required.' }
$root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$directoryPath = [IO.Path]::GetFullPath((Join-Path $root $Directory))
$allowed = [IO.Path]::GetFullPath((Join-Path $root 'docs/regression/Performance')) + [IO.Path]::DirectorySeparatorChar
if (-not $directoryPath.StartsWith($allowed, [StringComparison]::OrdinalIgnoreCase)) { throw 'Evidence path is outside Performance.' }
New-Item -ItemType Directory -Path $directoryPath -Force | Out-Null
$paths = @{
    A = Join-Path $root $BaselineWorkbook
    B = Join-Path $root $CandidateWorkbook
}
$hashes = @{}; foreach ($key in @('A','B')) { $hashes[$key] = (Get-FileHash -LiteralPath $paths[$key] -Algorithm SHA256).Hash }
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class PerformanceExcelIdentity {
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);
}
'@
function Get-ComProperty([object]$Target, [string]$Property) {
    if ($null -eq $Target) { throw "Missing COM target: $Property" }
    $value = $Target.GetType().InvokeMember($Property, [Reflection.BindingFlags]::GetProperty, $null, $Target, $null)
    if ($null -eq $value) { throw "Missing COM property: $Property" }
    return ,$value
}
function Get-Memory([int]$ProcessId) {
    $process = Get-Process -Id $ProcessId
    return [ordered]@{PrivateBytes=$process.PrivateMemorySize64; WorkingSet=$process.WorkingSet64; PeakWorkingSet=$process.PeakWorkingSet64}
}
$excel = $null; $book = $null; $records = New-Object 'System.Collections.Generic.List[object]'
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.EnableEvents = $false; $excel.AutomationSecurity = 1
    $window = [IntPtr][long](Get-ComProperty $excel 'Hwnd'); [uint32]$excelProcessId = 0
    [void][PerformanceExcelIdentity]::GetWindowThreadProcessId($window, [ref]$excelProcessId)
    $process = Get-Process -Id $excelProcessId
    $imagePath = $process.MainModule.FileName
    $pe = [IO.File]::ReadAllBytes($imagePath); $offset = [BitConverter]::ToInt32($pe,60)
    $machine = [BitConverter]::ToUInt16($pe,$offset+4); $bitness = 64; if ($machine -eq 332) {$bitness=32}; $pe=$null
    [ordered]@{ExcelPID=$excelProcessId; StartTicks=$process.StartTime.ToUniversalTime().Ticks; ExcelVersion=[string]$excel.Version; ExcelBuild=[string]$excel.Build; ExcelBitness=$bitness; ExcelExecutable=$imagePath; Sources=$hashes} |
        ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directoryPath 'Environment.json') -Encoding UTF8
    $cases = @(@{Name='model'; Count=$ModelCount}, @{Name='extent'; Count=$ExtentCount})
    if ($Group -eq 'Solver') {
        $cases = @(); foreach ($method in @('Newton','Secant')) { foreach ($count in @(1,10,30)) {
            $cases += @{Name="solver-$method-$count"; Method=$method; Count=$count}
        }}
    }
    if ($Group -eq 'Batch') {
        $cases = @(); foreach ($mode in $BatchModes) { foreach ($count in $Counts) {
            $cases += @{Name="batch-$Family-$Method-$Shape-$Order-$Extension-$mode-$count"; Mode=$mode; Count=$count}
        }}
    }
    if ($Group -eq 'Export') {
        $cases=@(); foreach($count in @(1,10,30)) {$cases += @{Name="export-$count";Count=$count}}
    }
    foreach ($case in $cases) {
        for ($repeat=1; $repeat -le $Repeats; $repeat++) {
            foreach ($version in @('A','B')) {
                $books = Get-ComProperty $excel 'Workbooks'
                $book = $books.Open($paths[$version])
                $excel.Calculation = -4135
                if ($Group -eq 'Batch' -or $Group -eq 'Export') {
                    if(-not $KeepInputs) {[void]$excel.Run("'$($book.Name)'!modTestPerformance.ConfigurePerformanceFixture",[string]$Family,[int]$case.Count,[string]$Method,[string]$Shape,[string]$Order,[string]$Extension)}
                    if ($repeat -eq 1) {
                        $inputs=@{}; foreach($name in @('rngSystemSettings','rngCalculationProfiles','rngLoadCombinations','rngStabilityDurationLoads')) {
                            $inputs[$name]=$book.Names.Item($name).RefersToRange.Value2
                        }
                        $inputs | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directoryPath "$($case.Name)-$version-input.json") -Encoding UTF8
                    }
                    if($Group -eq 'Export') {[void]$excel.Run("'$($book.Name)'!modTestPerformance.MeasurePerformanceBatch",'Full')}
                }
                for ($run=1; $run -le $RunsPerOpen; $run++) {
                $before = Get-Memory $excelProcessId
                if ($case.Name -eq 'model') { $result = [string]$excel.Run("'$($book.Name)'!modTestPerformance.MeasurePerformanceModel", [int]$case.Count) }
                elseif ($case.Name -eq 'extent') { $result = [string]$excel.Run("'$($book.Name)'!modTestPerformance.MeasurePerformanceExtent", [bool]($version -eq 'B'), [int]$case.Count) }
                elseif ($Group -eq 'Solver') { $result = [string]$excel.Run("'$($book.Name)'!modTestPerformance.MeasurePerformanceSolver", [string]$case.Method, [int]$SolverDivisions, [int]$case.Count) }
                elseif ($Group -eq 'Export') { $result = [string]$excel.Run("'$($book.Name)'!modTestPerformance.MeasurePerformanceExport",10) }
                else { $result = [string]$excel.Run("'$($book.Name)'!modTestPerformance.MeasurePerformanceBatch", [string]$case.Mode) }
                $after = Get-Memory $excelProcessId
                $delimiter = ';'; if ($Group -ne 'Storage') { $delimiter = [string][char]30 }
                $fields = @{}; foreach ($part in $result -split $delimiter) { $pair = $part.Trim() -split '=',2; if($pair.Count -eq 2){$fields[$pair[0]]=$pair[1]} }
                $seconds = [double]::Parse($fields['seconds'],[Globalization.CultureInfo]::InvariantCulture)
                $fields.Remove('seconds')
                $technical = @{}
                foreach ($key in @('preparation','core','clear','packing','summary','snapshot','plot','evaluations','K','reuse','solves','domainBuilds','mechanicalBuilds','released')) {
                    if ($fields.ContainsKey($key)) { $technical[$key]=$fields[$key]; $fields.Remove($key) }
                }
                if ($fields.ContainsKey('fingerprint')) {
                    $fingerprintPath = Join-Path $directoryPath "$($case.Name)-$version-$repeat-$run-fingerprint.txt"
                    [IO.File]::WriteAllText($fingerprintPath,[string]$fields['fingerprint'],[Text.UTF8Encoding]::new($false))
                    $fields['fingerprint']=(Get-FileHash -LiteralPath $fingerprintPath -Algorithm SHA256).Hash
                }
                $record = [ordered]@{Case=$case.Name; Version=$version; Repeat=$repeat; RunWithinOpen=$run; Pilot=[bool]$Pilot; Count=$case.Count; Seconds=$seconds; Values=$fields; Technical=$technical; MemoryBefore=$before; MemoryAfter=$after}
                $records.Add($record)
                $records | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $directoryPath 'Raw.json') -Encoding UTF8
                Write-Output ($record | ConvertTo-Json -Depth 5 -Compress)
                if($Group -eq 'Batch' -and $MeasureOverhead) {
                    $overhead=[string]$excel.Run("'$($book.Name)'!modTestPerformance.MeasurePerformanceOverhead")
                    $overhead | Set-Content -LiteralPath (Join-Path $directoryPath "$($case.Name)-$version-$repeat-$run-overhead.txt") -Encoding UTF8
                }
                if($Group -eq 'Batch' -and $case.Mode -eq 'Full') {
                    $reportPath=Join-Path (Split-Path -Parent $paths[$version]) 'RC_Section_NDM_execution_report.txt'
                    if(Test-Path -LiteralPath $reportPath) {
                        Copy-Item -LiteralPath $reportPath -Destination (Join-Path $directoryPath "$($case.Name)-$version-$repeat-$run-execution-report.txt")
                    }
                }
                }
                $book.Close($false); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book); $book=$null
                [GC]::Collect(); [GC]::WaitForPendingFinalizers()
            }
        }
    }
    $summary = @()
    foreach ($case in $cases) {
        $reference = @($records | Where-Object {$_.Case -eq $case.Name})[0].Values
        foreach ($record in @($records | Where-Object {$_.Case -eq $case.Name})) {
            if ($record.Values.Count -ne $reference.Count) { throw 'A/B field set differs.' }
            foreach ($key in $reference.Keys) { if ($record.Values[$key] -cne $reference[$key]) { throw "A/B difference: $($case.Name).$key" } }
        }
        foreach ($version in @('A','B')) { foreach ($run in 1..$RunsPerOpen) {
            $seconds = @($records | Where-Object {$_.Case -eq $case.Name -and $_.Version -eq $version -and $_.RunWithinOpen -eq $run} | ForEach-Object {$_.Seconds} | Sort-Object)
            $median=$seconds[[int][Math]::Floor($seconds.Count/2)]
            if ($seconds.Count % 2 -eq 0) {$median=($seconds[$seconds.Count/2-1]+$seconds[$seconds.Count/2])/2}
            $summary += [ordered]@{Case=$case.Name; Version=$version; RunWithinOpen=$run; Observations=$seconds.Count; Median=$median; Min=$seconds[0]; Max=$seconds[-1]; ExactValues=$true; FullPhasesMeasured=($Group -ne 'Batch' -or $case.Mode -ne 'Full')}
        }}
    }
    foreach ($key in @('A','B')) { if ((Get-FileHash -LiteralPath $paths[$key] -Algorithm SHA256).Hash -ne $hashes[$key]) { throw "Source changed: $key" } }
    $summary | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directoryPath 'Summary.json') -Encoding UTF8
}
finally {
    if ($null -ne $book) {try {$book.Close($false)} catch {Write-Warning $_}; [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)}
    if ($null -ne $excel) {try {$excel.Quit()} catch {Write-Warning $_}; [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
