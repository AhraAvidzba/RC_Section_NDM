# Sequential performance-stage gates. No concurrent Excel jobs or old full On/Off suite.
# Every child owns its hidden Excel; timeout is preserved as a failed gate.
param(
    [string]$BaselineWorkbook = 'docs/regression/Performance/BatchV3Baseline/RC_Section_NDM.xlsm',
    [string]$CandidateWorkbook = 'docs/regression/Performance/BatchV3Candidate/RC_Section_NDM.xlsm',
    [switch]$SkipDirected,
    [string]$EvidenceSuffix = ''
)
$ErrorActionPreference = 'Stop'
$root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$evidence = Join-Path $root 'docs/regression/Performance'
function Quote-Argument([string]$value) {
    if ($value.Contains('"')) { throw 'Quote in child argument.' }
    return '"' + $value + '"'
}
$jobs = @(
    @{Name='DirectedFinal'; Tool='Run-PerformanceDirected.ps1'; Args=@('-SourceWorkbook',$CandidateWorkbook)},
    @{Name='Crack63Measurements'; Family='CRACK_SP63'},
    @{Name='Stability35Measurements'; Family='STABILITY_SP35'},
    @{Name='Stability63Measurements'; Family='STABILITY_SP63'},
    @{Name='CapacityAutoMeasurements'; Family='CAPACITY'; Method='Newton'},
    @{Name='CapacityMultiplierMeasurements'; Family='CAPACITY_MULTIPLIER'; Method='Secant'; Shape='RectSet'},
    @{Name='MixedMeasurements'; Family='MIXED'},
    @{Name='ReverseMeasurements'; Family='CRACK_SP35'; Order='Reverse'; Count=30},
    @{Name='ShuffleMeasurements'; Family='CRACK_SP35'; Order='Shuffle'; Count=30},
    @{Name='CircleMeasurements'; Family='DIRECT'; Shape='Circle'; Count=30},
    @{Name='RoundedMeasurements'; Family='DIRECT'; Shape='RoundedRectangle'; Count=30},
    @{Name='OffMeasurements'; Family='DIRECT'; Extension='No'; Count=30},
    @{Name='StorageQPCMeasurements'; Group='Storage'}
)
$records = @()
foreach ($job in $jobs) {
    if($SkipDirected -and $job.Tool) {continue}
    $directory = 'docs/regression/Performance/' + $job.Name + $EvidenceSuffix
    $output = Join-Path $evidence ($job.Name + $EvidenceSuffix)
    New-Item -ItemType Directory -Path $output -Force | Out-Null
    $tool = 'Measure-PerformanceStorage.ps1'
    if ($job.Tool) {$tool=$job.Tool}
    $arguments = @('-NoProfile','-ExecutionPolicy','Bypass','-File',(Join-Path $PSScriptRoot $tool),'-Directory',$directory)
    if ($job.Tool) {
        $arguments += $job.Args
    } else {
        $group='Batch'; if($job.Group){$group=$job.Group}
        $arguments += @('-Group',$group,'-BaselineWorkbook',$BaselineWorkbook,'-CandidateWorkbook',$CandidateWorkbook)
        foreach($key in @('Family','Method','Shape','Order','Extension')) {
            if($job[$key]) {$arguments += @(('-'+$key),([string]$job[$key]))}
        }
        if($job.ContainsKey('Count')) {$arguments += @('-Counts',[string]$job['Count'],'-BatchModes','Staged')}
    }
    $arguments | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $output 'Arguments.json') -Encoding UTF8
    $command = ($arguments | ForEach-Object {Quote-Argument $_}) -join ' '
    $timer=[Diagnostics.Stopwatch]::StartNew()
    $child=Start-Process -FilePath (Join-Path $PSHOME 'powershell.exe') -ArgumentList $command -WorkingDirectory $root -WindowStyle Hidden -PassThru -RedirectStandardOutput (Join-Path $output 'runner.stdout.log') -RedirectStandardError (Join-Path $output 'runner.stderr.log')
    $null=$child.Handle
    $completed=$child.WaitForExit(3600000)
    if(-not $completed) {
        Stop-Process -Id $child.Id -Force -ErrorAction SilentlyContinue
        $identityPath=Join-Path $output 'Environment.json'
        if(Test-Path -LiteralPath $identityPath) {
            $identity=Get-Content -LiteralPath $identityPath -Raw | ConvertFrom-Json
            $owned=Get-Process -Id $identity.ExcelPID -ErrorAction SilentlyContinue
            if($owned -and $owned.ProcessName -eq 'EXCEL' -and $owned.StartTime.ToUniversalTime().Ticks -eq $identity.StartTicks) {Stop-Process -Id $owned.Id -Force}
        }
        throw "Performance gate timed out: $($job.Name)"
    }
    $child.WaitForExit(); $child.Refresh(); $timer.Stop()
    if(-not $job.Tool -and $child.ExitCode -eq 0) {
        $raw=Get-Content -LiteralPath (Join-Path $output 'Raw.json') -Raw | ConvertFrom-Json
        if($job.Family -and @($raw | Where-Object {$_.Case -notlike "batch-$($job.Family)-*"}).Count -gt 0) {throw 'Runner executed a different family.'}
        $expected=60; if($job.ContainsKey('Count')) {$expected=10}; if($job.Group -eq 'Storage') {$expected=20}
        if(@($raw).Count -ne $expected) {throw "Runner observation count differs: $(@($raw).Count), expected $expected"}
        if($job.ContainsKey('Count') -and @($raw | Where-Object {$_.Count -ne $job['Count']}).Count -gt 0) {throw 'Runner load count differs.'}
    }
    $records += [ordered]@{Name=$job.Name+$EvidenceSuffix;ExitCode=$child.ExitCode;Seconds=$timer.Elapsed.TotalSeconds;ArgumentsValidated=$true}
    $records | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $evidence ('GateSeries'+$EvidenceSuffix+'.json')) -Encoding UTF8
    Write-Output ($records[-1] | ConvertTo-Json -Compress)
    if($child.ExitCode -ne 0) {throw "Performance gate failed: $($job.Name). Preserved logs must be reviewed."}
}
