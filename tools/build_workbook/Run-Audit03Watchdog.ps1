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
    $memory=Get-CimInstance Win32_OperatingSystem
    $ready=$memory.FreePhysicalMemory*1KB -ge 4GB -and $memory.FreeVirtualMemory*1KB -ge 8GB
    [ordered]@{CheckedUTC=[DateTimeOffset]::UtcNow.ToString('o');Ready=$ready;FreePhysicalGiB=[math]::Round($memory.FreePhysicalMemory*1KB/1GB,3);FreeCommitGiB=[math]::Round($memory.FreeVirtualMemory*1KB/1GB,3);RequiredPhysicalGiB=4;RequiredCommitGiB=8;NewExcelStarted=$false}|ConvertTo-Json|Set-Content -LiteralPath ($report+'.memory.json') -Encoding UTF8
    if($ready){break}
    if(([DateTimeOffset]::UtcNow-$memoryGateStarted).TotalSeconds -ge 300){Add-Content -LiteralPath $report -Encoding UTF8 -Value 'WATCHDOG_FAILURE: memory reserve below 4 GiB physical or 8 GiB commit for five minutes; no test Excel started';throw 'Memory reserve insufficient; no test Excel started.'}
    Start-Sleep -Seconds 10
}
$process = Start-Process -FilePath (Join-Path $PSHOME "powershell.exe") `
    -ArgumentList $arguments -WorkingDirectory $root -WindowStyle Hidden -PassThru `
    -RedirectStandardOutput ($report + ".stdout.log") -RedirectStandardError ($report + ".stderr.log")
$null = $process.Handle
if (-not $process.WaitForExit($TimeoutSeconds * 1000)) {
    Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue
    # Закрываем только PID, явно записанный дочерним runner-ом. Время старта
    # защищает от повторного использования PID; новые пользовательские окна
    # Excel никогда не входят в область очистки этого watchdog.
    $identity = Get-Content -LiteralPath $report -Encoding UTF8 -ErrorAction SilentlyContinue |
        Where-Object { $_ -match '^TEST_EXCEL_PROCESS: id=(\d+); startTicks=(\d+)$' } | Select-Object -First 1
    if ($identity -match '^TEST_EXCEL_PROCESS: id=(\d+); startTicks=(\d+)$') {
        $testExcelProcessId = [int]$Matches[1]
        $startTicks = [long]$Matches[2]
        $excel = Get-Process -Id $testExcelProcessId -ErrorAction SilentlyContinue
        if ($excel -and $excel.ProcessName -eq 'EXCEL' -and $excel.StartTime.ToUniversalTime().Ticks -eq $startTicks) {
            Stop-Process -Id $testExcelProcessId -Force -ErrorAction SilentlyContinue
        }
    }
    Add-Content -LiteralPath $report -Encoding UTF8 -Value "WATCHDOG_FAILURE: timeout=$TimeoutSeconds sec; macro=$Macro; test Excel terminated"
    Save-ProbeProgress
    Write-Output "WATCHDOG_FAILURE: $Macro exceeded $TimeoutSeconds seconds."
    exit 1
}
$process.WaitForExit()
$process.Refresh()
Save-ProbeProgress
$exitCode = $process.ExitCode
if ($null -eq $exitCode) { throw "Watchdog не получил код завершения дочернего runner-а." }
if ((Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash -ne $hash) {
    Add-Content -LiteralPath $report -Encoding UTF8 -Value "WATCHDOG_FAILURE: source workbook changed"
    exit 1
}
Add-Content -LiteralPath $report -Encoding UTF8 -Value "WATCHDOG_COMPLETED: exit=$exitCode; source unchanged=True"
Write-Output "WATCHDOG_COMPLETED: exit=$exitCode; report=$report"
exit $exitCode
