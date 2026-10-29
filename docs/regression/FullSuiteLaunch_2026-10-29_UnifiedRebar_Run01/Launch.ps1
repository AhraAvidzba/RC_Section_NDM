# ДЛЯ ТЕСТОВ: фиксирует входы и запускает один полный набор через Task Scheduler.
param([switch]$UserConfirmedFullRun,[string]$UserAuthorizationText='')
$ErrorActionPreference='Stop'
# Отдельное подтверждение нового длинного набора требуется до любых запусков и записей.
if(-not $UserConfirmedFullRun -or [string]::IsNullOrWhiteSpace($UserAuthorizationText)){throw 'WAITING_EXPLICIT_NEW_FULL_AUTHORIZATION; no launch, task or workbook process created.'}
Import-Module (Join-Path $PSHOME 'Modules/Microsoft.PowerShell.Utility/Microsoft.PowerShell.Utility.psd1') -Global -ErrorAction Stop
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../../..')).Path
if($PSVersionTable.PSVersion.Major -ne 5){throw 'Use Windows PowerShell 5.1.'}
Import-Module ScheduledTasks -ErrorAction Stop
Set-Location -LiteralPath $root

# Перед стартом все адресные процессы должны быть завершены и проверены по PID и UTC ticks.
$eligibility=Get-Content -LiteralPath (Join-Path $PSScriptRoot 'ReleaseEligibility.json') -Raw -Encoding UTF8|ConvertFrom-Json
if(-not $eligibility.Ready -or -not $eligibility.AllOwnClosed){throw 'Release eligibility or own closure proof missing.'}
foreach($own in $eligibility.OwnIdentities){
 $p=$null;try{$p=Get-Process -Id ([int]$own.PID) -ErrorAction Stop}catch{if($_.FullyQualifiedErrorId-notlike'NoProcessFoundForGivenId*'){throw}}
 if($p-and$p.ProcessName-ieq'EXCEL'){
  if($null-eq$p.StartTime){throw 'Previous own identity unreadable'}
  if([string]$p.StartTime.ToUniversalTime().Ticks-ceq[string]$own.StartTicksUTC){throw 'Previous own Excel still live'}
 }
}
$epoch=Get-Content -LiteralPath (Join-Path $PSScriptRoot 'SourceEpoch.json') -Raw -Encoding UTF8|ConvertFrom-Json
$evidence='docs/regression/FullOnOff_2026-10-29_UnifiedRebar_Run01'
if(Test-Path -LiteralPath (Join-Path $root $evidence)){throw 'Fresh full evidence directory required.'}
if((Test-Path -LiteralPath (Join-Path $PSScriptRoot 'FullLaunch.json')) -or (Test-Path -LiteralPath (Join-Path $PSScriptRoot 'LaunchRequest.json'))){throw 'Launch already recorded; never duplicate.'}
# Подтверждение относится только к проверенной опубликованной версии.
if((& git rev-parse HEAD) -cne $epoch.HEAD -or (& git branch --show-current) -cne $epoch.Branch -or (Get-FileHash -LiteralPath 'workbook/output/RC_Section_NDM.xlsm').Hash -cne $epoch.BookSHA256 -or (Get-FileHash -LiteralPath 'workbook/output/VBA_All_Code.txt').Hash -cne $epoch.AllCodeSHA256){throw 'Reviewed published source epoch differs.'}
& git diff --quiet HEAD -- src tests tools/build_workbook workbook/output/RC_Section_NDM.xlsm workbook/output/VBA_All_Code.txt AGENTS.md docs/FocusedValidationPlan.md
if($LASTEXITCODE -ne 0){throw 'Inputs must be committed before full launch.'}
$untracked=@(& git ls-files --others --exclude-standard -- src tests tools/build_workbook)
if($untracked.Count){throw ('Uncommitted new participating inputs: '+($untracked -join ', '))}
$memory=Get-CimInstance Win32_PerfFormattedData_PerfOS_Memory
if($null -eq $memory -or $null -eq $memory.AvailableMBytes -or $null -eq $memory.CommitLimit -or $null -eq $memory.CommittedBytes -or [double]$memory.CommitLimit -le 0){throw 'Actual memory counters unavailable; no full process started.'}
$freeRAM=[double]$memory.AvailableMBytes/1024
$freeCommit=([double]$memory.CommitLimit-[double]$memory.CommittedBytes)/1GB
if($freeRAM -lt 8 -or $freeCommit -lt 14){throw 'Insufficient memory reserve; no full driver started.'}
$source='workbook/output/RC_Section_NDM.xlsm';$started=[DateTimeOffset]::Now;$deadline=$started.AddHours(8)
$overallDeadline=[DateTimeOffset]::Parse('2026-10-30T11:00:00+03:00',[Globalization.CultureInfo]::InvariantCulture)
if($deadline -gt $overallDeadline){$deadline=$overallDeadline}
if($started.AddHours(4.5).AddMinutes(40) -gt $overallDeadline){throw 'Full forecast plus final review does not fit the user deadline; no launch'}
$components=@(Get-ChildItem -LiteralPath (Join-Path $root 'src'),(Join-Path $root 'tests') -Recurse -File|Where-Object {$_.Extension -in '.cls','.bas'})
$scripts=@(Get-ChildItem -LiteralPath (Join-Path $root 'tools/build_workbook') -Recurse -File|Where-Object {$_.Extension -in '.ps1','.py','.bas','.vba','.json'})
function Test-ImmutableEntry([string]$Path){
    $absolute=(Resolve-Path -LiteralPath $Path).Path
    if(-not $absolute.StartsWith($root+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)){throw 'Input outside workspace.'}
    [ordered]@{Path=$absolute.Substring($root.Length+1).Replace('\','/');SHA256=(Get-FileHash -LiteralPath $absolute).Hash}
}
$taskName='RC_NDM_Full_20261029_UnifiedRebar_Run01'
if(Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue){throw 'Own launch task already exists; inspect it instead of relaunching.'}
$protected=@($source,'workbook/output/VBA_All_Code.txt','AGENTS.md','docs/FocusedValidationPlan.md')+@('FullDriver.ps1','FullDriverHost.ps1','Launch.ps1','Read-CurrentStatus.ps1','TestObserveCompleted.py','Test-ObserverSelftest.py','Read-UiProgressSupplement.ps1','SourceEpoch.json','ReleaseEligibility.json'|ForEach-Object {Join-Path $PSScriptRoot $_})
$request=[ordered]@{Status='AUTHORIZED_READY_TO_START';Started=$started.ToString('o');Deadline=$deadline.ToString('o');ForecastHours=4.5;SourceWorkbook=$source;SourceSHA256=(Get-FileHash -LiteralPath $source).Hash;SourceCommit=(& git rev-parse HEAD);AllCodeSHA256=(Get-FileHash -LiteralPath 'workbook/output/VBA_All_Code.txt').Hash;Branch=(& git branch --show-current);EvidenceDirectory=$evidence;ExpectedSuites=182;IncludesFocusedAcceptance=$true;CoverageVersion=7;Python='C:/Users/avidzba/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/python.exe';ProtectedInputs=@(foreach($f in $protected){Test-ImmutableEntry $f});SourceComponents=@(foreach($f in $components){Test-ImmutableEntry $f.FullName});ParticipatingScripts=@(foreach($f in $scripts){Test-ImmutableEntry $f.FullName});SourcePublicationContract='Clean full rebuilt workbook; user-data transfer is excluded from this acceptance source. Root publication is required before this actual launch.';UserAuthorization=$UserAuthorizationText;ExplicitUserConfirmation=[bool]$UserConfirmedFullRun;NewLongRepeatRequiresSeparateAuthorization=$false;RepeatsAuthorizedWithinOverallDeadline=$true;OverallUserDeadline=$overallDeadline.ToString('o');ReviewReserveMinutes=40;RepeatStartCondition='Start only if forecast of the entire remaining full composition and final review fits before the overall deadline';ScheduledTaskName=$taskName;LaunchMethod='Windows Task Scheduler, interactive current user, on-demand single run outside Codex process tree';FreePhysicalBeforeGiB=[math]::Round($freeRAM,3);FreeCommitBeforeGiB=[math]::Round($freeCommit,3);CommitLimitBytes=[double]$memory.CommitLimit;CommittedBytes=[double]$memory.CommittedBytes;RequiredRAMGiB=8;RequiredCommitGiB=14}
$request|ConvertTo-Json -Depth 7|Set-Content -LiteralPath (Join-Path $PSScriptRoot 'LaunchRequest.json') -Encoding UTF8
$identity=[Security.Principal.WindowsIdentity]::GetCurrent().Name
$hostPath=Join-Path $PSScriptRoot 'FullDriverHost.ps1'
$action=New-ScheduledTaskAction -Execute (Join-Path $PSHOME 'powershell.exe') -Argument ('-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "'+$hostPath+'"') -WorkingDirectory $root
$principal=New-ScheduledTaskPrincipal -UserId $identity -LogonType Interactive -RunLevel Limited
$settings=New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit ([TimeSpan]::Zero) -MultipleInstances IgnoreNew
Register-ScheduledTask -TaskName $taskName -Action $action -Principal $principal -Settings $settings -Description 'One explicitly authorized RC_Section_NDM full test run; no recurring trigger.'|Out-Null
Export-ScheduledTask -TaskName $taskName|Set-Content -LiteralPath (Join-Path $PSScriptRoot 'ScheduledTask.xml') -Encoding UTF8
$registered=Get-ScheduledTask -TaskName $taskName
if(@($registered.Triggers|Where-Object {$null -ne $_}).Count -gt 0){throw 'Unexpected trigger: full task must be on-demand only.'}
Start-ScheduledTask -TaskName $taskName
$driverRun=Join-Path $PSScriptRoot 'DriverRun.json'
for($attempt=0;$attempt -lt 40;$attempt++){if(Test-Path -LiteralPath $driverRun){break};Start-Sleep -Seconds 1}
if(-not(Test-Path -LiteralPath $driverRun)){throw 'Task requested; driver identity not yet available. Inspect task; do not retry automatically.'}
$driver=Get-Content -LiteralPath $driverRun -Raw -Encoding UTF8|ConvertFrom-Json
# -------------------- ДЛЯ ТЕСТОВ --------------------
# Только явный NotFound означает завершение; ошибка доступа не скрывается.
function Test-ReadLaunchDriver([int]$DriverId,[string]$ExpectedTicks){
    try {$process=Get-Process -Id $DriverId -ErrorAction Stop}
    catch {
        if($_.FullyQualifiedErrorId -like 'NoProcessFoundForGivenId*'){return $null}
        throw ('Driver identity is unknown: '+$_.Exception.Message)
    }
    $start=$process.StartTime
    if($null -eq $start){throw 'Driver identity is unknown: StartTime is unavailable.'}
    if($process.ProcessName -ne 'powershell' -or [string]$start.ToUniversalTime().Ticks -cne $ExpectedTicks){throw 'Driver identity mismatch.'}
    return $process
}
$process=Test-ReadLaunchDriver ([int]$driver.PID) ([string]$driver.StartTicks)
$hostRun=Get-Content -LiteralPath (Join-Path $PSScriptRoot 'HostRun.json') -Raw -Encoding UTF8|ConvertFrom-Json
[ordered]@{Status=$(if($process){'STARTED'}else{'DRIVER_ALREADY_TERMINAL_REVIEW_COMPLETION'});HostPID=$hostRun.PID;HostStartTicks=[string]$hostRun.StartTicks;PID=$driver.PID;StartTicks=[string]$driver.StartTicks;Started=$started.ToString('o');Deadline=$deadline.ToString('o');SourceSHA256=$request.SourceSHA256;SourceCommit=$request.SourceCommit;Branch=$request.Branch;EvidenceDirectory=$evidence;ExpectedSuites=182;IncludesFocusedAcceptance=$true;CoverageVersion=7;AllCodeSHA256=$request.AllCodeSHA256;RequestSHA256=(Get-FileHash -LiteralPath (Join-Path $PSScriptRoot 'LaunchRequest.json')).Hash;ScheduledTaskName=$taskName}|ConvertTo-Json -Depth 5|Set-Content -LiteralPath (Join-Path $PSScriptRoot 'FullLaunch.json') -Encoding UTF8
Get-Content -LiteralPath (Join-Path $PSScriptRoot 'FullLaunch.json') -Raw -Encoding UTF8
