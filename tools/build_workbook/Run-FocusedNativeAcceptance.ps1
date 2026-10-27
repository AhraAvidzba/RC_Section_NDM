# Bounded native CAD tests; only processes with recorded PID and start time are closed.
param([string]$SourceWorkbook='workbook/output/RC_Section_NDM.xlsm',
      [string]$ReportDirectory='docs/regression/FocusedAcceptance_2026-10-10_V3/NativeCAD',
      [ValidateRange(30,900)][int]$MaxSeconds=420,
      [string]$DeadlineUTC='')
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSHOME 'Modules/Microsoft.PowerShell.Utility/Microsoft.PowerShell.Utility.psd1') -Global -ErrorAction Stop
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$watch=[Diagnostics.Stopwatch]::StartNew()
$nativeDeadline=[DateTimeOffset]::UtcNow.AddSeconds($MaxSeconds)
if ($DeadlineUTC) { $outer=[DateTimeOffset]::Parse($DeadlineUTC,[Globalization.CultureInfo]::InvariantCulture); if ($outer -lt $nativeDeadline) { $nativeDeadline=$outer } }
if ($nativeDeadline -le [DateTimeOffset]::UtcNow) { throw 'Native CAD deadline already reached; no child started.' }
$results=@()
$stageDirectory=Join-Path $root $ReportDirectory
[void][IO.Directory]::CreateDirectory($stageDirectory)
$nativeRunner=Join-Path $PSScriptRoot 'Run-PostAudit03NativeCAD.ps1'
$nativeRunnerHash=(Get-FileHash -LiteralPath $nativeRunner).Hash
Copy-Item -LiteralPath $nativeRunner -Destination (Join-Path $stageDirectory 'NativeRunner.ps1')
function Stop-Owned([string]$Directory) {
    foreach ($item in @(@{File='NativeExcelOwnership.json'; Name='EXCEL'}, @{File='NativeCADOwnership.json'; Name='acad'})) {
        $identityPath=Join-Path $Directory $item.File
        if (-not (Test-Path -LiteralPath $identityPath)) { continue }
        $identity=Get-Content -LiteralPath $identityPath -Raw | ConvertFrom-Json
        $owned=Get-Process -Id $identity.OwnedPID -ErrorAction SilentlyContinue
        if ($owned -and $identity.StartTicks -and $owned.ProcessName -eq $item.Name -and
            $owned.StartTime.ToUniversalTime().Ticks -eq [long]$identity.StartTicks) {
            Stop-Process -Id $owned.Id -Force -ErrorAction SilentlyContinue
        }
    }
}
foreach ($macro in @('RunRealAutoCADFarOriginTests','RunRealAutoCADFarCapacityTests')) {
    $remaining=[math]::Floor(($nativeDeadline-[DateTimeOffset]::UtcNow).TotalSeconds)
    if ($remaining -lt 10) { throw 'Incomplete: native CAD stage deadline reached.' }
    $relative=$ReportDirectory+'/'+$macro
    $directory=Join-Path $root $relative
    if (Test-Path -LiteralPath $directory) { throw 'Native evidence already exists; choose a new output directory.' }
    New-Item -ItemType Directory -Path $directory | Out-Null
    $arguments='-NoProfile -ExecutionPolicy Bypass -File "'+(Join-Path $PSScriptRoot 'Run-PostAudit03NativeCAD.ps1')+'" -SourceWorkbook "'+$SourceWorkbook+'" -ReportDirectory "'+$relative+'" -Macro '+$macro+' -DeadlineUTC "'+$nativeDeadline.ToString('o')+'"'
    $child=Start-Process -FilePath (Join-Path $PSHOME 'powershell.exe') -ArgumentList $arguments -WorkingDirectory $root `
        -WindowStyle Hidden -PassThru -RedirectStandardOutput (Join-Path $directory 'stdout.log') -RedirectStandardError (Join-Path $directory 'stderr.log')
    $null=$child.Handle
    Write-Output ('NATIVE_STARTED: '+$macro+'; budget='+$remaining+' sec')
    try {
        $completed=$child.WaitForExit($remaining*1000)
        if (-not $completed) { Stop-Process -Id $child.Id -Force; throw 'Incomplete: native CAD watchdog timeout.' }
        $child.WaitForExit(); $child.Refresh()
        if ($child.ExitCode -ne 0) { throw ('Native CAD child failed: '+$macro) }
        $report=Join-Path $directory 'NativeCAD.txt'
        $body=Get-Content -LiteralPath $report -Raw -Encoding UTF8
        if ($body -notmatch 'TOTAL[^\r\n]*failed=0\b' -or $body -match '(?m)^FAIL:|RUNTIME ERROR') { throw 'Native CAD assertions failed.' }
        $results+=@{Macro=$macro; Passed=$true; ReportSHA256=(Get-FileHash -LiteralPath $report).Hash}
        Write-Output (($body -split '\r?\n' | Where-Object {$_ -match '^TOTAL'}) -join '')
    } finally { Stop-Owned $directory }
}
if ((Get-FileHash -LiteralPath $nativeRunner).Hash -cne $nativeRunnerHash) {throw 'Native runner changed during the stage.'}
[ordered]@{Passed=$true; ElapsedSeconds=$watch.Elapsed.TotalSeconds; LimitSeconds=$MaxSeconds; Results=$results; UserDocumentsUsed=$false;RunnerSHA256=$nativeRunnerHash} |
    ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path (Join-Path $root $ReportDirectory) 'Stage.json') -Encoding UTF8
Write-Output 'NATIVE_FOCUSED_STAGE: passed=True'
exit 0
