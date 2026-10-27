# Выполняет VBA-тесты на независимой книге через существующий mode-runner.
# Ограничивает время COM-вызова и закрывает созданный тестом Excel при зависании;
# timeout является ошибкой проверки, а не успешным численным результатом.
param(
    [Parameter(Mandatory=$true)][string]$SourceWorkbook,
    [Parameter(Mandatory=$true)][string]$ReportPath,
    [string]$Macro = "",
    [string]$MacroArgument1 = "",
    [string]$MacroArgument2 = "",
    [ValidateSet("Rectangles", "Circle", "PointedSection", "HollowRectangle")][string]$FixtureGeometryType,
    [switch]$VerifyResultsReopen,
    [switch]$VerifyStatusReopen,
    [switch]$FailFast,
    [ValidateSet("Yes", "No")][string]$Mode = "No",
    [ValidateRange(10, 172800)][int]$TimeoutSeconds = 120
)
$ErrorActionPreference = "Stop"
Import-Module (Join-Path $PSHOME 'Modules/Microsoft.PowerShell.Utility/Microsoft.PowerShell.Utility.psd1') -Global -ErrorAction Stop
$root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot "../..")).Path
$source = (Resolve-Path -LiteralPath $SourceWorkbook).Path
$hash = (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash
$report = Join-Path $root $ReportPath
New-Item -ItemType Directory -Force -Path (Split-Path -Parent $report) | Out-Null

# Кавычки сохраняют пробелы в известных путях аргументов дочернего PowerShell.
function Quote-ProcessArgument([string]$Value) {
    if ($Value.Contains('"')) { throw "Аргумент watchdog содержит недопустимую кавычку." }
    return '"' + $Value + '"'
}

# Сохраняет журнал последних assertions рядом с основным отчетом, включая
# остановленный macro. Это диагностический артефакт, а не доказательство PASS.
function Save-ProbeProgress {
    if (-not (Test-Path -LiteralPath $report)) { return }
    $fixtureLine = Get-Content -LiteralPath $report -Encoding UTF8 | Where-Object { $_ -like 'FIXTURE: *' } | Select-Object -First 1
    if (-not $fixtureLine) { return }
    $fixture = $fixtureLine.Substring('FIXTURE: '.Length)
    $progress = Join-Path (Split-Path -Parent $fixture) 'Audit03_Search_Progress.txt'
    if (Test-Path -LiteralPath $progress) {
        Copy-Item -LiteralPath $progress -Destination ($report + '.progress.txt') -Force
    }
    foreach ($name in @('Config_Preflight_Progress','Config_Preflight_Timeline')) {
        $progress = Join-Path (Split-Path -Parent $fixture) ($name + '.txt')
        if (Test-Path -LiteralPath $progress) {
            Copy-Item -LiteralPath $progress -Destination ($report + '.' + $name + '.txt') -Force
        }
    }
}

# -------------------- ДЛЯ ТЕСТОВ --------------------
# Только явный NoProcessFoundForGivenId доказывает отсутствие. Остальные ошибки не маскируются под exit.
function Get-TestOwnedExcelExitState([int]$ID,[long]$Ticks){
    $state=[ordered]@{State='Unknown';PID=$ID;StartTicksUTC=[string]$Ticks;Reason='';ObservedUTC=[DateTimeOffset]::UtcNow.ToString('o')}
    if($ID -le0 -or $Ticks -le0){$state.Reason='Missing or invalid identity';return [pscustomobject]$state}
    try{$item=Get-Process -Id $ID -ErrorAction Stop}catch{
        if($_.FullyQualifiedErrorId -match '^NoProcessFoundForGivenId(,|$)'){$state.State='Absent';$state.Reason='Explicit NoProcessFoundForGivenId'}else{$state.Reason='Process query failed: '+$_.Exception.Message}
        return [pscustomobject]$state
    }
    if($null -eq $item){$state.Reason='Process query returned null without explicit not-found';return [pscustomobject]$state}
    try{
        $name=[string]$item.ProcessName;$started=$item.StartTime
        if($null -eq $started){throw 'StartTime is null'}
        $observed=$started.ToUniversalTime().Ticks
        if($name -ine 'EXCEL' -or $observed -ne $Ticks){$state.Reason='PID identity mismatch; absence not asserted'}else{$state.State='Live';$state.Reason='Exact own Excel is still alive'}
    }catch{$state.Reason='Process identity unavailable: '+$_.Exception.Message}
    return [pscustomobject]$state
}

# -------------------- ДЛЯ ТЕСТОВ --------------------
# Одна первичная identity обязана быть записана дочерним mode-runner до открытия копии.
function Read-TestOwnedExcelIdentity([string]$Primary){
    try{$identityRows=@(Get-Content -LiteralPath $Primary -Encoding UTF8 -ErrorAction Stop|Where-Object {$_ -match '^TEST_EXCEL_PROCESS: id=(\d+); startTicks=(\d+)$'})}catch{return [pscustomobject]@{Valid=$false;PID=0;Ticks=0L;Reason='Primary identity could not be read'}}
    if($identityRows.Count -ne1){return [pscustomobject]@{Valid=$false;PID=0;Ticks=0L;Reason='Exactly one primary Excel identity required'}}
    $null=$identityRows[0] -match '^TEST_EXCEL_PROCESS: id=(\d+); startTicks=(\d+)$'
    try{$id=[int]$Matches[1];$ticks=[long]$Matches[2]}catch{return [pscustomobject]@{Valid=$false;PID=0;Ticks=0L;Reason='Invalid identity range'}}
    return [pscustomobject]@{Valid=($id -gt0 -and $ticks -gt0);PID=$id;Ticks=$ticks;Reason='Primary TEST_EXCEL_PROCESS'}
}

# -------------------- ДЛЯ ТЕСТОВ --------------------
# Проверяется только после actual child exit. Обычное завершение никого не останавливает.
function Get-TestPostChildExcelClosure([string]$Primary,[int]$GraceMilliseconds=3000){
    $identity=Read-TestOwnedExcelIdentity $Primary;$observations=@();$timer=[Diagnostics.Stopwatch]::StartNew()
    do{
        $state=Get-TestOwnedExcelExitState ([int]$identity.PID) ([long]$identity.Ticks);$observations+=,$state
        if($state.State -eq 'Absent' -or -not $identity.Valid -or $timer.ElapsedMilliseconds -ge $GraceMilliseconds){break}
        Start-Sleep -Milliseconds 200
    }while($true)
    return [pscustomobject]@{Passed=($identity.Valid -and $state.State -eq 'Absent');Identity=$identity;FinalState=$state;Observations=$observations;GraceMilliseconds=$GraceMilliseconds;OrdinaryPathKilledAnything=$false}
}

$arguments = @(
    "-NoProfile", "-ExecutionPolicy", "Bypass", "-File",
    (Quote-ProcessArgument (Join-Path $PSScriptRoot "Run-Audit02ModeTests.ps1")),
    "-SourceWorkbook", (Quote-ProcessArgument $source),
    "-ReportPath", (Quote-ProcessArgument $ReportPath),
    "-Mode", $Mode
)
if (-not [string]::IsNullOrWhiteSpace($Macro)) {
    $arguments += @("-Macro", (Quote-ProcessArgument $Macro))
}
if ($MacroArgument1) { $arguments += @("-MacroArgument1", (Quote-ProcessArgument $MacroArgument1)) }
if ($MacroArgument2) { $arguments += @("-MacroArgument2", (Quote-ProcessArgument $MacroArgument2)) }
if ($FixtureGeometryType) { $arguments += @("-FixtureGeometryType", (Quote-ProcessArgument $FixtureGeometryType)) }
if ($VerifyResultsReopen) { $arguments += "-VerifyResultsReopen" }
if ($VerifyStatusReopen) { $arguments += "-VerifyStatusReopen" }
if ($FailFast) { $arguments += "-FailFast" }
$arguments = $arguments -join " "
# ДЛЯ ТЕСТОВ: перед новым Excel оставляет запас для Codex и пика проверки.
# Во время ожидания не запускает дочерний процесс и не закрывает чужие приложения.
$memoryGateStarted=[DateTimeOffset]::UtcNow
while($true){
    $memory=Get-CimInstance Win32_PerfFormattedData_PerfOS_Memory
    $freePhysicalBytes=[double]$memory.AvailableMBytes*1MB
    $freeCommitBytes=[double]$memory.CommitLimit-[double]$memory.CommittedBytes
    $ready=$freePhysicalBytes -ge 4GB -and $freeCommitBytes -ge 8GB
    [ordered]@{CheckedUTC=[DateTimeOffset]::UtcNow.ToString('o');Ready=$ready;FreePhysicalGiB=[math]::Round($freePhysicalBytes/1GB,3);FreeCommitGiB=[math]::Round($freeCommitBytes/1GB,3);CommitLimitBytes=[double]$memory.CommitLimit;CommittedBytes=[double]$memory.CommittedBytes;RequiredPhysicalGiB=4;RequiredCommitGiB=8;NewExcelStarted=$false}|ConvertTo-Json|Set-Content -LiteralPath ($report+'.memory.json') -Encoding UTF8
    if($ready){break}
    if(([DateTimeOffset]::UtcNow-$memoryGateStarted).TotalSeconds -ge 300){Add-Content -LiteralPath $report -Encoding UTF8 -Value 'WATCHDOG_FAILURE: memory reserve below 4 GiB physical or 8 GiB commit for five minutes; no test Excel started';throw 'Memory reserve insufficient; no test Excel started.'}
    Start-Sleep -Seconds 10
}
$process = Start-Process -FilePath (Join-Path $PSHOME "powershell.exe") `
    -ArgumentList $arguments -WorkingDirectory $root -WindowStyle Hidden -PassThru `
    -RedirectStandardOutput ($report + ".stdout.log") -RedirectStandardError ($report + ".stderr.log")
$null = $process.Handle
if (-not $process.WaitForExit($TimeoutSeconds * 1000)) {
    # Удерживаемый handle относится к запущенному child; PID не переоткрывается вслепую.
    $childStopError=$null
    try{if(-not $process.HasExited){$process.Kill()};[void]$process.WaitForExit(3000)}catch{$childStopError=$_.Exception.Message}
    $identity=Read-TestOwnedExcelIdentity $report;$state=Get-TestOwnedExcelExitState ([int]$identity.PID) ([long]$identity.Ticks)
    $excelStopRequested=$false;$excelStopError=$null
    if($identity.Valid -and $state.State -eq 'Live'){
        try{
            $own=Get-Process -Id ([int]$identity.PID) -ErrorAction Stop;$null=$own.Handle
            if($own.ProcessName -ieq 'EXCEL' -and $own.StartTime.ToUniversalTime().Ticks -eq [long]$identity.Ticks){$excelStopRequested=$true;$own.Kill();[void]$own.WaitForExit(3000)}else{throw 'Timeout identity changed before cleanup'}
        }catch{$excelStopError=$_.Exception.Message}
    }
    $final=Get-TestOwnedExcelExitState ([int]$identity.PID) ([long]$identity.Ticks)
    [ordered]@{Stage='TIMEOUT';Passed=$false;Identity=$identity;Before=$state;After=$final;ChildStopError=$childStopError;ExcelStopRequested=$excelStopRequested;ExcelStopError=$excelStopError;ExactOwnExcelStopRequestedOnlyByExplicitTimeout=$excelStopRequested}|ConvertTo-Json -Depth 7|Set-Content -LiteralPath ($report+'.ProcessClosure.json') -Encoding UTF8
    Add-Content -LiteralPath $report -Encoding UTF8 -Value "WATCHDOG_FAILURE: timeout=$TimeoutSeconds sec; macro=$Macro; ownExcelState=$($final.State); exactCleanupRequested=$excelStopRequested"
    Save-ProbeProgress
    Write-Output "WATCHDOG_FAILURE: $Macro exceeded $TimeoutSeconds seconds."
    exit 1
}
$process.WaitForExit()
$process.Refresh()
Save-ProbeProgress
$exitCode = $process.ExitCode
if ($null -eq $exitCode) { throw "Watchdog не получил код завершения дочернего runner-а." }
$closure=Get-TestPostChildExcelClosure $report
$cleanupPath=$report+'.RunnerCleanup.json';$cleanup=$null;$cleanupValid=$false
try{
    $cleanup=Get-Content -LiteralPath $cleanupPath -Raw -Encoding UTF8 -ErrorAction Stop|ConvertFrom-Json
    $cleanupValid=$cleanup.OwnerVerified -and [int]$cleanup.PID -eq [int]$closure.Identity.PID -and [long]$cleanup.StartTicksUTC -eq [long]$closure.Identity.Ticks -and $cleanup.QuitRequested -and $cleanup.QuitSucceeded -and @($cleanup.Errors).Count -eq0
}catch{$cleanupReadError=$_.Exception.Message}
[ordered]@{Stage='AFTER_CHILD_EXIT';Passed=($closure.Passed -and $cleanupValid);ChildHasExited=$process.HasExited;ChildExitCode=$exitCode;Closure=$closure;RunnerCleanupValid=$cleanupValid;RunnerCleanup=$cleanup;RunnerCleanupReadError=$cleanupReadError;OrdinaryPathKilledAnything=$false}|ConvertTo-Json -Depth 9|Set-Content -LiteralPath ($report+'.ProcessClosure.json') -Encoding UTF8
if(-not $closure.Passed -or -not $cleanupValid){
    Add-Content -LiteralPath $report -Encoding UTF8 -Value "CLEANUP ERROR: post-child Excel closure=$($closure.FinalState.State); cleanupRequestsValid=$cleanupValid; no process terminated on ordinary path"
    Add-Content -LiteralPath $report -Encoding UTF8 -Value 'WATCHDOG_FAILURE: own Excel exit not proven after child exit'
    exit 1
}
if ((Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash -ne $hash) {
    Add-Content -LiteralPath $report -Encoding UTF8 -Value "WATCHDOG_FAILURE: source workbook changed"
    exit 1
}
Add-Content -LiteralPath $report -Encoding UTF8 -Value "WATCHDOG_COMPLETED: exit=$exitCode; source unchanged=True"
Write-Output "WATCHDOG_COMPLETED: exit=$exitCode; report=$report"
exit $exitCode
