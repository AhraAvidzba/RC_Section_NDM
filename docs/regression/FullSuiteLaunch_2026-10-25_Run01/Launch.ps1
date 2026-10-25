# ДЛЯ ТЕСТОВ: фиксирует входы и запускает один полный набор через Task Scheduler.
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../../..')).Path
Set-Location -LiteralPath $root
$evidence='docs/regression/FullOnOff_2026-10-25_Current_Run01'
if(Test-Path -LiteralPath (Join-Path $root $evidence)){throw 'Fresh full evidence directory required.'}
if(Test-Path -LiteralPath (Join-Path $PSScriptRoot 'FullLaunch.json')){throw 'Launch already recorded; never duplicate.'}
& git diff --quiet HEAD -- src tests tools/build_workbook workbook/output/RC_Section_NDM.xlsm workbook/output/VBA_All_Code.txt
if($LASTEXITCODE -ne 0){throw 'Inputs must be committed before full launch.'}
$memory=Get-CimInstance Win32_OperatingSystem
if($memory.FreePhysicalMemory*1KB -lt 4GB -or $memory.FreeVirtualMemory*1KB -lt 8GB){throw 'Insufficient memory reserve; no full driver started.'}
$source='workbook/output/RC_Section_NDM.xlsm';$started=[DateTimeOffset]::Now;$deadline=$started.AddHours(8)
$components=@(Get-ChildItem -LiteralPath (Join-Path $root 'src'),(Join-Path $root 'tests') -Recurse -File|Where-Object {$_.Extension -in '.cls','.bas'})
$scripts=@(Get-ChildItem -LiteralPath (Join-Path $root 'tools/build_workbook') -Recurse -File|Where-Object {$_.Extension -in '.ps1','.py','.bas','.vba','.json'})
function Test-ImmutableEntry([string]$Path){
    $absolute=(Resolve-Path -LiteralPath $Path).Path
    if(-not $absolute.StartsWith($root+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)){throw 'Input outside workspace.'}
    [ordered]@{Path=$absolute.Substring($root.Length+1).Replace('\','/');SHA256=(Get-FileHash -LiteralPath $absolute).Hash}
}
$taskName='RC_NDM_Full_20261025_Run01'
if(Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue){throw 'Own launch task already exists; inspect it instead of relaunching.'}
$protected=@($source,'workbook/output/VBA_All_Code.txt','AGENTS.md','docs/FocusedValidationPlan.md')+@('FullDriver.ps1','FullDriverHost.ps1','Launch.ps1','Read-CurrentStatus.ps1'|ForEach-Object {Join-Path $PSScriptRoot $_})
$request=[ordered]@{Status='AUTHORIZED_READY_TO_START';Started=$started.ToString('o');Deadline=$deadline.ToString('o');ForecastHours=4.5;SourceWorkbook=$source;SourceSHA256=(Get-FileHash -LiteralPath $source).Hash;SourceCommit=(& git rev-parse HEAD);Branch=(& git branch --show-current);EvidenceDirectory=$evidence;ExpectedSuites=168;IncludesFocusedAcceptance=$true;CoverageVersion=5;Python='C:/Users/avidzba/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/python.exe';ProtectedInputs=@(foreach($f in $protected){Test-ImmutableEntry $f});SourceComponents=@(foreach($f in $components){Test-ImmutableEntry $f.FullName});ParticipatingScripts=@(foreach($f in $scripts){Test-ImmutableEntry $f.FullName});UserAuthorization='Explicit user goal: run complete tests and fix errors; one new full ON/OFF plus all mandatory Focused stages';NewLongRepeatRequiresSeparateAuthorization=$true;ScheduledTaskName=$taskName;LaunchMethod='Windows Task Scheduler, interactive current user, on-demand single run outside Codex process tree';FreePhysicalBeforeGiB=[math]::Round($memory.FreePhysicalMemory*1KB/1GB,3);FreeCommitBeforeGiB=[math]::Round($memory.FreeVirtualMemory*1KB/1GB,3)}
$request|ConvertTo-Json -Depth 7|Set-Content -LiteralPath (Join-Path $PSScriptRoot 'LaunchRequest.json') -Encoding UTF8
$identity=[Security.Principal.WindowsIdentity]::GetCurrent().Name
$hostPath=Join-Path $PSScriptRoot 'FullDriverHost.ps1'
$action=New-ScheduledTaskAction -Execute (Join-Path $PSHOME 'powershell.exe') -Argument ('-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "'+$hostPath+'"') -WorkingDirectory $root
$principal=New-ScheduledTaskPrincipal -UserId $identity -LogonType Interactive -RunLevel Limited
$settings=New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit ([TimeSpan]::Zero) -MultipleInstances IgnoreNew
Register-ScheduledTask -TaskName $taskName -Action $action -Principal $principal -Settings $settings -Description 'One explicitly authorized RC_Section_NDM full test run; no recurring trigger.'|Out-Null
Export-ScheduledTask -TaskName $taskName|Set-Content -LiteralPath (Join-Path $PSScriptRoot 'ScheduledTask.xml') -Encoding UTF8
Start-ScheduledTask -TaskName $taskName
$driverRun=Join-Path $PSScriptRoot 'DriverRun.json'
for($attempt=0;$attempt -lt 40;$attempt++){if(Test-Path -LiteralPath $driverRun){break};Start-Sleep -Seconds 1}
if(-not(Test-Path -LiteralPath $driverRun)){throw 'Task requested; driver identity not yet available. Inspect task; do not retry automatically.'}
$driver=Get-Content -LiteralPath $driverRun -Raw -Encoding UTF8|ConvertFrom-Json
$process=Get-Process -Id $driver.PID -ErrorAction Stop
if($process.StartTime.ToUniversalTime().Ticks -ne [long]$driver.StartTicks){throw 'Driver identity mismatch.'}
[ordered]@{Status='STARTED';PID=$driver.PID;StartTicks=[string]$driver.StartTicks;Started=$started.ToString('o');Deadline=$deadline.ToString('o');SourceSHA256=$request.SourceSHA256;SourceCommit=$request.SourceCommit;Branch=$request.Branch;EvidenceDirectory=$evidence;ExpectedSuites=168;IncludesFocusedAcceptance=$true;RequestSHA256=(Get-FileHash -LiteralPath (Join-Path $PSScriptRoot 'LaunchRequest.json')).Hash;ScheduledTaskName=$taskName}|ConvertTo-Json -Depth 5|Set-Content -LiteralPath (Join-Path $PSScriptRoot 'FullLaunch.json') -Encoding UTF8
Get-Content -LiteralPath (Join-Path $PSScriptRoot 'FullLaunch.json') -Raw -Encoding UTF8
