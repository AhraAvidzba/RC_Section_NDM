# ДЛЯ ТЕСТОВ: однократный родитель драйвера вне процесса Codex.
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSHOME 'Modules/Microsoft.PowerShell.Utility/Microsoft.PowerShell.Utility.psd1') -Global -ErrorAction Stop
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../../..')).Path
$hostProcess=Get-Process -Id $PID
[ordered]@{PID=$PID;StartTicks=[string]$hostProcess.StartTime.ToUniversalTime().Ticks;StartedUTC=[DateTimeOffset]::UtcNow.ToString('o')}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $PSScriptRoot 'HostRun.json') -Encoding UTF8
$driver=Join-Path $PSScriptRoot 'FullDriver.ps1'
$process=Start-Process -FilePath (Join-Path $PSHOME 'powershell.exe') -ArgumentList @('-NoProfile','-ExecutionPolicy','Bypass','-File',('"'+$driver+'"')) -WorkingDirectory $root -WindowStyle Hidden -RedirectStandardOutput (Join-Path $PSScriptRoot 'FullLaunch.stdout.log') -RedirectStandardError (Join-Path $PSScriptRoot 'FullLaunch.stderr.log') -PassThru
$null=$process.Handle
$driverTicks=$process.StartTime.ToUniversalTime().Ticks
[ordered]@{PID=$PID;StartTicks=[string]$hostProcess.StartTime.ToUniversalTime().Ticks;DriverPID=$process.Id;DriverStartTicks=[string]$driverTicks;StartedUTC=[DateTimeOffset]::UtcNow.ToString('o')}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $PSScriptRoot 'HostRun.json') -Encoding UTF8
$process.WaitForExit();$process.Refresh()
$exitCode=$process.ExitCode
[ordered]@{FinishedUTC=[DateTimeOffset]::UtcNow.ToString('o');DriverExitCode=$exitCode}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $PSScriptRoot 'HostCompletion.json') -Encoding UTF8
if($null -eq $exitCode){throw 'Driver exit code unavailable.'}
exit $exitCode
