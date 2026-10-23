# ДЛЯ ТЕСТОВ: текущая проверка размеров смещения -50/0/50 на собственной копии.
# Исторический before/after здесь не запускается; пользовательская книга не открывается.
param(
    [Parameter(Mandatory=$true)][string]$SourceWorkbook,
    [Parameter(Mandatory=$true)][ValidatePattern('^[A-Fa-f0-9]{64}$')][string]$ExpectedSourceSHA256,
    [Parameter(Mandatory=$true)][string]$ReportPath,
    [ValidateSet('No','Yes')][string]$Mode='No',
    [ValidateRange(10,600)][int]$TimeoutSeconds=120,
    [switch]$InternalPrepare
)
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$workspace=$root.TrimEnd('\')+'\'
$sourcePath=$SourceWorkbook
if(-not [IO.Path]::IsPathRooted($sourcePath)){$sourcePath=Join-Path $root $sourcePath}
$source=(Resolve-Path -LiteralPath $sourcePath).Path
if(-not $source.StartsWith($workspace,[StringComparison]::OrdinalIgnoreCase)){throw 'Источник должен находиться в рабочей папке проекта.'}
$sourceHash=(Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash
if($sourceHash -cne $ExpectedSourceSHA256.ToUpperInvariant()){throw 'SHA256 исходной книги не совпадает с явно указанным значением.'}
$report=$ReportPath
if(-not [IO.Path]::IsPathRooted($report)){$report=Join-Path $root $report}
$report=[IO.Path]::GetFullPath($report)
$evidenceRoot=[IO.Path]::GetFullPath((Join-Path $root 'docs/regression')).TrimEnd('\')+'\'
if(-not $report.StartsWith($evidenceRoot,[StringComparison]::OrdinalIgnoreCase)){throw 'Новый ReportPath должен находиться в docs/regression.'}
$watchdogReport=$report.Substring($workspace.Length) # Центральный runner ожидает путь отчёта относительно корня проекта.
$fixtureDirectory=$report+'.fixture'
foreach($path in @($report,$report+'.json',$report+'.stdout.log',$report+'.stderr.log')){
    if(Test-Path -LiteralPath $path){throw ('Доказательства уже существуют; укажите новый ReportPath: '+$path)}
}
if(-not $InternalPrepare -and (Test-Path -LiteralPath $fixtureDirectory)){throw 'Частная проверка уже существует; нужен новый ReportPath.'}
$protected=@{}
foreach($relative in @('workbook/output/RC_Section_NDM.xlsm','workbook/output/VBA_All_Code.txt','workbook/output/RC_Section_NDM_execution_report.txt')){
    $path=Join-Path $root $relative
    if(Test-Path -LiteralPath $path){$protected[$path]=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash}
}
$private=Join-Path $fixtureDirectory 'RC_Section_NDM.xlsm'
if($InternalPrepare){
    if(-not (Test-Path -LiteralPath $private) -or (Test-Path -LiteralPath (Join-Path $fixtureDirectory 'Preparation.json'))){throw 'Не найден новый private input родительского wrapper.'}
}else{
    New-Item -ItemType Directory -Path $fixtureDirectory | Out-Null
    Copy-Item -LiteralPath $source -Destination $private
}
if((Get-FileHash -LiteralPath $private -Algorithm SHA256).Hash -cne $sourceHash){throw 'Частная копия не совпала с исходной книгой.'}
$helper=Join-Path $PSScriptRoot 'RectanglesOffsetProbe.bas'
$helperHash=(Get-FileHash -LiteralPath $helper -Algorithm SHA256).Hash

# Право Close/Quit/Stop относится только к новой паре PID/время старта.
function Test-OwnedProcessIdentity([int]$ProcessId,[long]$StartTicks,[string]$ProcessName='EXCEL'){
    if($ProcessId -le 0 -or $StartTicks -le 0){return $false}
    $current=Get-Process -Id $ProcessId -ErrorAction SilentlyContinue
    return ($null -ne $current -and $current.ProcessName -eq $ProcessName -and $current.StartTime.ToUniversalTime().Ticks -eq $StartTicks)
}
function Quote-ProcessArgument([string]$Value){if($Value.Contains('"')){throw 'Кавычка в аргументе не поддерживается.'};return '"'+$Value+'"'}

$excel=$null;$book=$null;$ownedPid=0;$ownedTicks=0;$compiled=$false;$testExit=$null
if($InternalPrepare){
try{
    Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public static class RectanglesCurrentWindow{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);}'
    $existingExcelIds=@([Diagnostics.Process]::GetProcessesByName('EXCEL') | ForEach-Object {$_.Id})
    $excel=New-Object -ComObject Excel.Application
    [uint32]$id=0;[void][RectanglesCurrentWindow]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$id)
    if($id -eq 0 -or $existingExcelIds -contains [int]$id){throw 'COM вернул существующий или неопределённый Excel: настройки не изменены, закрытие запрещено.'}
    $newProcess=Get-Process -Id $id -ErrorAction Stop
    $newTicks=$newProcess.StartTime.ToUniversalTime().Ticks
    if(-not (Test-OwnedProcessIdentity ([int]$id) $newTicks)){throw 'Не подтверждена идентичность нового Excel.'}
    $ownedPid=[int]$id;$ownedTicks=$newTicks
    [ordered]@{PID=$ownedPid;StartTicks=$ownedTicks} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $fixtureDirectory 'OwnExcel.json') -Encoding UTF8
    Write-Output ('OWNED_EXCEL: id='+$ownedPid+'; utcStartTicks='+$ownedTicks)
    $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false;$excel.AutomationSecurity=3
    $book=$excel.Workbooks.Open($private,0,$false)
    if($book.ReadOnly -or [string]$book.FullName -cne $private){throw 'Не удалось открыть свою частную книгу для добавления тестового фрагмента.'}
    $module=$book.VBProject.VBComponents.Item('modTestSectionRecipes').CodeModule
    $body=$module.Lines(1,$module.CountOfLines)
    if($body -match '(?im)^Public Function RunRectanglesOffsetDimensionTests\b'){throw 'Источник уже содержит частный probe; нужна чистая текущая книга.'}
    $module.AddFromString([IO.File]::ReadAllText($helper,[Text.Encoding]::UTF8))
    $excel.VBE.ActiveVBProject=$book.VBProject
    $compile=$excel.VBE.CommandBars.FindControl(1,578)
    if($null -eq $compile){throw 'Команда VBE Compile недоступна.'}
    if($compile.Enabled){$compile.Execute()}
    if($compile.Enabled){throw 'Частный Rectangles probe не скомпилирован.'}
    $compiled=$true
    $book.Save() # Сохраняется только собственная копия: watchdog откроет её независимую копию.
}finally{
    try{if($book -and (Test-OwnedProcessIdentity $ownedPid $ownedTicks)){$book.Close($false)}}finally{
        try{if($excel){try{if(Test-OwnedProcessIdentity $ownedPid $ownedTicks){$excel.Quit()}}finally{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}}}finally{
            [GC]::Collect();[GC]::WaitForPendingFinalizers()
            if(Test-OwnedProcessIdentity $ownedPid $ownedTicks){Stop-Process -Id $ownedPid -Force -ErrorAction SilentlyContinue}
            if((Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash -cne $sourceHash){throw 'Исходная книга изменилась во время подготовки.'}
            foreach($path in $protected.Keys){if((Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash -cne $protected[$path]){throw ('Защищённый файл изменился: '+$path)}}
        }
    }
}
[ordered]@{OwnPID=$ownedPid;OwnStartTicks=$ownedTicks;Compiled=$compiled;SourceSHA256=$sourceHash;HelperSHA256=$helperHash;PrivateSHA256=(Get-FileHash -LiteralPath $private -Algorithm SHA256).Hash} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $fixtureDirectory 'Preparation.json') -Encoding UTF8
exit 0
}
# Подготовка тоже ограничена временем: модальная ошибка VBE не оставляет вечный Excel.
$arguments=@('-NoProfile','-ExecutionPolicy','Bypass','-File',(Quote-ProcessArgument $PSCommandPath),'-SourceWorkbook',(Quote-ProcessArgument $source),'-ExpectedSourceSHA256',$sourceHash,'-ReportPath',(Quote-ProcessArgument $report),'-Mode',$Mode,'-TimeoutSeconds',[string]$TimeoutSeconds,'-InternalPrepare') -join ' '
$process=Start-Process -FilePath (Join-Path $PSHOME 'powershell.exe') -ArgumentList $arguments -WorkingDirectory $root -WindowStyle Hidden -PassThru -RedirectStandardOutput (Join-Path $fixtureDirectory 'Prepare.stdout.log') -RedirectStandardError (Join-Path $fixtureDirectory 'Prepare.stderr.log')
$null=$process.Handle;$driverTicks=$process.StartTime.ToUniversalTime().Ticks
if(-not $process.WaitForExit($TimeoutSeconds*1000)){
    if(Test-OwnedProcessIdentity $process.Id $driverTicks 'powershell'){Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue}
    $identityPath=Join-Path $fixtureDirectory 'OwnExcel.json'
    if(Test-Path -LiteralPath $identityPath){$identity=Get-Content -LiteralPath $identityPath -Raw -Encoding UTF8 | ConvertFrom-Json;if(Test-OwnedProcessIdentity ([int]$identity.PID) ([long]$identity.StartTicks)){Stop-Process -Id ([int]$identity.PID) -Force -ErrorAction SilentlyContinue}}
    throw 'Превышено время частной compile-подготовки Rectangles; проверка не пройдена.'
}
$process.WaitForExit();$process.Refresh()
if($process.ExitCode -ne 0){throw 'Частная подготовка Rectangles не завершена; см. Prepare.stderr.log.'}
$preparation=Get-Content -LiteralPath (Join-Path $fixtureDirectory 'Preparation.json') -Raw -Encoding UTF8 | ConvertFrom-Json
if(-not $preparation.Compiled -or $preparation.SourceSHA256 -cne $sourceHash -or $preparation.HelperSHA256 -cne $helperHash){throw 'Не совпало доказательство частной подготовки.'}
$compiled=$true;$ownedPid=[int]$preparation.OwnPID;$ownedTicks=[long]$preparation.OwnStartTicks
$privateHash=(Get-FileHash -LiteralPath $private -Algorithm SHA256).Hash
if($privateHash -cne $preparation.PrivateSHA256){throw 'Частная книга изменилась после compile-подготовки.'}
try{
    & (Join-Path $PSScriptRoot 'Run-Audit03Watchdog.ps1') -SourceWorkbook $private -ReportPath $watchdogReport -Macro 'modTestSectionRecipes.RunRectanglesOffsetDimensionTests' -Mode $Mode -TimeoutSeconds $TimeoutSeconds
    $testExit=$LASTEXITCODE
    $actual=Get-Content -LiteralPath $report -Raw -Encoding UTF8
    if($testExit -ne 0 -or $actual -notmatch 'TOTAL_RECTANGLES_OFFSET: passed=\d+; failed=0\b' -or $actual -match '(?m)^FAIL:|RUNTIME ERROR|WATCHDOG_FAILURE'){throw 'Проверка Rectangles завершилась с ошибкой; подробности в новом отчёте.'}
}finally{
    $sourceUnchanged=((Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash -ceq $sourceHash)
    $privateUnchanged=((Get-FileHash -LiteralPath $private -Algorithm SHA256).Hash -ceq $privateHash)
    $protectedUnchanged=$true
    foreach($path in $protected.Keys){if((Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash -cne $protected[$path]){$protectedUnchanged=$false}}
    [ordered]@{Source=$source;SourceSHA256=$sourceHash;HelperSHA256=$helperHash;PrivateSHA256=$privateHash;OwnPID=$ownedPid;OwnStartTicks=$ownedTicks;Compiled=$compiled;Macro='modTestSectionRecipes.RunRectanglesOffsetDimensionTests';Mode=$Mode;ExitCode=$testExit;SourceUnchanged=$sourceUnchanged;PrivateUnchanged=$privateUnchanged;ProtectedUnchanged=$protectedUnchanged;HistoricalBeforeExecuted=$false} | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath ($report+'.json') -Encoding UTF8
    if(-not $sourceUnchanged -or -not $privateUnchanged -or -not $protectedUnchanged){throw 'Контроль неизменности исходных и опубликованных файлов не пройден.'}
}
Write-Output ('RECTANGLES_OFFSET_CURRENT_CHECK_FINISHED: '+$report)
