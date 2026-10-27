# Runs native contour acceptance in an owned Autodesk AutoCAD instance.
# Existing user/SOFiPLUS instances are never edited, closed, or reused.
param([string]$SourceWorkbook='workbook/output/RC_Section_NDM.xlsm',
      [string]$ReportDirectory='docs/regression/PostAudit03/NativeCAD',
      [ValidateSet('RunRealAutoCADContourTests','RunRealAutoCADContourFormatTests','RunRealAutoCADContourMeshTests','RunRealAutoCADFarOriginTests','RunRealAutoCADFarCapacityTests','RunSP35WorkbookTests')]
      [string]$Macro='RunRealAutoCADContourTests',
      [ValidateSet('No','Yes')][string]$Mode,
      [switch]$LaunchInstalledAutoCAD,
      [switch]$AcquireInsideExcel,
      [switch]$ShowOwnedWindow,
      [string]$DeadlineUTC='')
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSHOME 'Modules/Microsoft.PowerShell.Utility/Microsoft.PowerShell.Utility.psd1') -Global -ErrorAction Stop
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$directory=Join-Path $root $ReportDirectory
New-Item -ItemType Directory -Path $directory -Force | Out-Null
$path=Join-Path $directory 'RC_Section_NDM.xlsm'
$sourcePath=$SourceWorkbook
if (-not [IO.Path]::IsPathRooted($sourcePath)) { $sourcePath=Join-Path $root $sourcePath }
$sourcePath=(Resolve-Path -LiteralPath $sourcePath).Path
$sourceHash=(Get-FileHash -LiteralPath $sourcePath -Algorithm SHA256).Hash
$nativePassed=$false
Copy-Item -LiteralPath $sourcePath -Destination $path -Force
$class=(Get-ItemProperty -LiteralPath 'Registry::HKEY_CLASSES_ROOT\AutoCAD.Application.24.2\CLSID').'(default)'
$server=(Get-ItemProperty -LiteralPath "Registry::HKEY_CLASSES_ROOT\CLSID\$class\LocalServer32").'(default)'
$installedAcad='C:\Program Files\Autodesk\AutoCAD 2023\acad.exe'
if (-not $LaunchInstalledAutoCAD -and $server -notlike "$installedAcad*") {throw "Unsafe CAD COM registration: $server"}
if ($LaunchInstalledAutoCAD -and -not (Test-Path -LiteralPath $installedAcad)) {throw 'Installed Autodesk acad.exe was not found.'}
Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class PostAuditCADIdentity {
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);

}
'@
$existing=@(Get-Process acad -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
if($LaunchInstalledAutoCAD -and $existing.Count -gt 0){throw 'Explicit ROT launch requires no existing AutoCAD; use verified direct COM activation with an open user instance.'}
# Проверяет сохранённую пару PID/время старта перед Close/Quit/Stop-Process.
# Отсутствующий или повторно использованный PID не даёт права закрывать приложение.
function Test-OwnedProcessIdentity {
    param([int]$ProcessId, [long]$StartTicks, [string]$ProcessName)
    if ($ProcessId -le 0 -or $StartTicks -le 0) { return $false }
    $current=Get-Process -Id $ProcessId -ErrorAction SilentlyContinue
    return ($null -ne $current -and $current.ProcessName -eq $ProcessName -and
        $current.StartTime.ToUniversalTime().Ticks -eq $StartTicks)
}

# -------------------- ДЛЯ ТЕСТОВ --------------------
# Проверяет фактическую RAM и запас commit перед каждым новым CAD/Excel, включая уже работающий CAD.
# Ожидание не превышает 300 с и общий deadline; отказ счётчиков запрещает создание приложения.
function Wait-TestNativeCreationMemory {
    param([ValidateSet('InstalledCAD','CADCOM','Excel')][string]$CreationKind, [string]$DeadlineUTC='')
    $started=[DateTimeOffset]::UtcNow
    $limit=$started.AddSeconds(300)
    if ($DeadlineUTC) { $outer=[DateTimeOffset]::Parse($DeadlineUTC,[Globalization.CultureInfo]::InvariantCulture); if ($outer -lt $limit) { $limit=$outer } }
    $proofPath=Join-Path $directory ('MemoryBefore'+$CreationKind+'.json')
    while ($true) {
        $proof=[ordered]@{CheckedUTC=[DateTimeOffset]::UtcNow.ToString('o');Status='MEMORY_QUERY_FAILED';Ready=$false;CreationKind=$CreationKind;RequiredPhysicalGiB=4;RequiredCommitGiB=8;WaitLimitUTC=$limit.ToString('o');NewApplicationStarted=$false}
        try {
            $memory=Get-CimInstance Win32_PerfFormattedData_PerfOS_Memory -ErrorAction Stop
            foreach ($name in @('AvailableMBytes','CommitLimit','CommittedBytes')) { if ($null -eq $memory.PSObject.Properties[$name] -or $null -eq $memory.$name) { throw ('Missing memory counter: '+$name) } }
            $physical=[double]$memory.AvailableMBytes*1MB
            $commit=[double]$memory.CommitLimit-[double]$memory.CommittedBytes
            if ([double]::IsNaN($physical) -or [double]::IsInfinity($physical) -or [double]::IsNaN($commit) -or [double]::IsInfinity($commit) -or $physical -lt 0 -or [double]$memory.CommitLimit -le 0 -or [double]$memory.CommittedBytes -lt 0) { throw 'Invalid memory counters.' }
            $proof['FreePhysicalGiB']=$physical/1GB; $proof['FreeCommitGiB']=$commit/1GB
            $proof['CommitLimitBytes']=[double]$memory.CommitLimit; $proof['CommittedBytes']=[double]$memory.CommittedBytes
        } catch {
            $proof['Error']=$_.Exception.Message
            $proof | ConvertTo-Json | Set-Content -LiteralPath $proofPath -Encoding UTF8
            throw ('Memory counters unavailable; no '+$CreationKind+' started.')
        }
        $remaining=($limit-[DateTimeOffset]::UtcNow).TotalSeconds
        $ready=$physical -ge 4GB -and $commit -ge 8GB -and $remaining -gt 0
        $proof.Ready=$ready; $proof.Status=if ($remaining -le 0) {'MEMORY_WAIT_DEADLINE'} elseif ($ready) {'READY'} else {'WAITING_FOR_MEMORY'}
        $proof['WaitedSeconds']=([DateTimeOffset]::UtcNow-$started).TotalSeconds
        $proof | ConvertTo-Json | Set-Content -LiteralPath $proofPath -Encoding UTF8
        if ($ready) { return }
        if ($remaining -le 0) { throw ('Memory reserve/deadline insufficient (4 GiB physical, 8 GiB commit); no '+$CreationKind+' started.') }
        Start-Sleep -Milliseconds ([int][math]::Max(1,[math]::Min(10000,[math]::Ceiling($remaining*1000))))
    }
}

$excel=$null; $book=$null; $acad=$null; $owned=$false; $pidValue=0; $cadStartTicks=0
$ownedExcelPid=0; $ownedExcelTicks=0
$launchedCad=$null; $launchedTicks=0
function Invoke-CadReady([scriptblock]$Action) {
    $deadline=[DateTime]::UtcNow.AddSeconds(90)
    while ($true) {
        try { $value = & $Action; return ,$value }
        catch {
            # IDispatch через Reflection может иметь несколько оболочек.
            # Повторяется только известный отказ COM занятого приложения.
            $nativeException=$_.Exception
            while ($nativeException.InnerException) {$nativeException=$nativeException.InnerException}
            $errorCode=$nativeException.HResult
            if ($errorCode -notin @(-2147418111,-2147417846) -or [DateTime]::UtcNow -ge $deadline) { throw }
            Write-Warning ('CAD_COM_READY_RETRY: HRESULT='+$errorCode)
            Start-Sleep -Milliseconds 1000
        }
    }
}
try {
    if ($AcquireInsideExcel -and ($existing.Count -gt 0 -or $Macro -ne 'RunRealAutoCADContourTests')) {throw 'Native ROT acquisition requires no existing acad process and the contour entrypoint.'}
    if ($LaunchInstalledAutoCAD) {
        # ProgID может быть занят SOFiPLUS. Запускаем проверенный acad.exe и
        # подключаемся только к его ROT-объекту, не меняя регистрацию вручную.
        Write-Output ('CAD_EXPLICIT_LAUNCH: ' + $installedAcad + '; registered server=' + $server)
        Wait-TestNativeCreationMemory -CreationKind 'InstalledCAD' -DeadlineUTC $DeadlineUTC
        $existing=@([Diagnostics.Process]::GetProcessesByName('acad') | ForEach-Object { $_.Id })
        if ($existing.Count -gt 0) { throw 'CAD appeared before explicit own launch.' }
        $launchedCad=Start-Process -FilePath $installedAcad -ArgumentList '/nologo' -WindowStyle Hidden -PassThru
        $launchedTicks=$launchedCad.StartTime.ToUniversalTime().Ticks
        [ordered]@{Executable=$installedAcad;OwnedPID=$launchedCad.Id;StartTicks=$launchedTicks;ExistingPIDs=$existing;UserDocumentsUsed=$false} |
            ConvertTo-Json | Set-Content -LiteralPath (Join-Path $directory 'NativeCADOwnership.json') -Encoding UTF8
        Write-Output ('CAD_LAUNCHED: PID=' + $launchedCad.Id + '; startTicks=' + $launchedTicks)
        $deadline=[DateTime]::UtcNow.AddSeconds(180)
        while ($null -eq $acad) {
            if ($launchedCad.HasExited) {throw 'Owned Autodesk AutoCAD exited before COM readiness.'}
            try { $acad=[Runtime.InteropServices.Marshal]::GetActiveObject('AutoCAD.Application.24.2') }
            catch {
                if ([DateTime]::UtcNow -ge $deadline) {throw 'Owned Autodesk AutoCAD did not publish its ROT object within 180 seconds.'}
                Start-Sleep -Milliseconds 1000
            }
        }
    } else {
        Write-Output ('CAD_ACTIVATE: ' + $server)
        Wait-TestNativeCreationMemory -CreationKind 'CADCOM' -DeadlineUTC $DeadlineUTC
        $existing=@([Diagnostics.Process]::GetProcessesByName('acad') | ForEach-Object { $_.Id })
        if ($AcquireInsideExcel -and $existing.Count -gt 0) { throw 'CAD appeared before own ROT acquisition.' }
        $acad=New-Object -ComObject AutoCAD.Application.24.2
    }
    $fullName=[string](Invoke-CadReady { $acad.GetType().InvokeMember('FullName',[Reflection.BindingFlags]::GetProperty,$null,$acad,$null) })
    if ([string]::IsNullOrWhiteSpace($fullName)) {throw 'AutoCAD FullName is unavailable through IDispatch.'}
    if ($fullName -notlike 'C:\Program Files\Autodesk\AutoCAD 2023\acad.exe') {throw "Not Autodesk AutoCAD: $fullName"}
    [uint32]$processId=0
    $window=Invoke-CadReady { $acad.GetType().InvokeMember('HWND',[Reflection.BindingFlags]::GetProperty,$null,$acad,$null) }
    [void][PostAuditCADIdentity]::GetWindowThreadProcessId([IntPtr][long]$window,[ref]$processId)
    $pidValue=[int]$processId
    $owned=($pidValue -gt 0 -and $existing -notcontains $pidValue)
    if ($LaunchInstalledAutoCAD) {
        $owned=$owned -and $pidValue -eq $launchedCad.Id -and
            (Get-Process -Id $pidValue).StartTime.ToUniversalTime().Ticks -eq $launchedTicks
    }
    if (-not $owned) {throw "CAD instance is not owned by this test: PID=$pidValue"}
    Write-Output "CAD_OWNED: PID=$pidValue; executable=$fullName"
    $cadStartTicks=(Get-Process -Id $pidValue).StartTime.ToUniversalTime().Ticks
    [ordered]@{Executable=$fullName; OwnedPID=$pidValue; StartTicks=$cadStartTicks; ExistingPIDs=$existing; UserDocumentsUsed=$false} |
        ConvertTo-Json | Set-Content -LiteralPath (Join-Path $directory 'NativeCADOwnership.json') -Encoding UTF8
    Invoke-CadReady { $acad.Visible=[bool]$ShowOwnedWindow }
    # Готовим первый документ в управляющем COM-процессе до передачи
    # application в Excel: скрытый CAD без документа может зависнуть
    # на первом межпроцессном чтении Documents из VBA.
    $template=Join-Path $env:LOCALAPPDATA 'Autodesk/AutoCAD 2023/R24.2/rus/Template/acadiso.dwt'
    if (-not (Test-Path -LiteralPath $template)) {throw 'Standard AutoCAD test template is missing.'}
    $documents=Invoke-CadReady { ,$acad.GetType().InvokeMember('Documents',[Reflection.BindingFlags]::GetProperty,$null,$acad,$null) }
    if ($Macro -in @('RunRealAutoCADFarOriginTests','RunRealAutoCADFarCapacityTests')) {
        $bootstrap=Invoke-CadReady { $acad.Documents.Add() }
        $template=''
    } else {
        $bootstrap=Invoke-CadReady { $acad.Documents.Add($template) }
    }
    Write-Output ('CAD_BOOTSTRAP_DOCUMENT: ' + [string]$bootstrap.Name)
    Wait-TestNativeCreationMemory -CreationKind 'Excel' -DeadlineUTC $DeadlineUTC
    $existingExcelIds=@([Diagnostics.Process]::GetProcessesByName('EXCEL') | ForEach-Object { $_.Id })
    $excel=New-Object -ComObject Excel.Application
    [uint32]$excelId=0
    [void][PostAuditCADIdentity]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$excelId)
    if ($excelId -eq 0 -or $existingExcelIds -contains [int]$excelId) {
        throw 'COM returned an existing or unidentified Excel; no properties changed and no ownership file written.'
    }
    $newExcelProcess=Get-Process -Id $excelId -ErrorAction Stop
    $newExcelTicks=$newExcelProcess.StartTime.ToUniversalTime().Ticks
    if (-not (Test-OwnedProcessIdentity ([int]$excelId) $newExcelTicks 'EXCEL')) { throw 'New Excel identity could not be verified.' }
    $ownedExcelPid=[int]$excelId; $ownedExcelTicks=$newExcelTicks
    $excel.Visible=$false; $excel.DisplayAlerts=$false; $excel.EnableEvents=$false; $excel.AutomationSecurity=1
    [ordered]@{OwnedPID=$ownedExcelPid; StartTicks=$ownedExcelTicks; ExistingPIDs=$existingExcelIds; ProcessName='EXCEL'} |
        ConvertTo-Json | Set-Content -LiteralPath (Join-Path $directory 'NativeExcelOwnership.json') -Encoding UTF8
    $book=$excel.Workbooks.Open($path)
    $excel.DisplayAlerts=$false
    # Отмена фонового теста ограничивается внешним watchdog, а не Escape в другом окне.
    $excel.EnableCancelKey=0
    if ($Mode) {
        # Режим меняется только в private native-копии; исходная книга не сохраняется.
        $settingsTable=$book.Names.Item('rngSystemSettings').RefersToRange
        $modeCell=$null
        for ($r=2;$r -le $settingsTable.Rows.Count;$r++) {
            if ([string]$settingsTable.Cells.Item($r,1).Value2 -ceq 'General.DiagramExtension') {$modeCell=$settingsTable.Cells.Item($r,2);break}
        }
        if ($null -eq $modeCell) {throw 'DiagramExtension is missing in native fixture.'}
        $modeCell.Value2=$Mode
        Write-Output ('NATIVE_DIAGRAM_MODE: '+$Mode+'; fixture only')
    }
    # The saved candidate already contains the guarded optional CAD entrypoint.
    # Test it without replacing its VBA or saving fixture mutations.
    if ($AcquireInsideExcel) {
        $result=[string]$excel.Run("'$($book.Name)'!modTestAutoCADContours.$Macro",$null,$template)
    } elseif ($Macro -eq 'RunSP35WorkbookTests') {
        $result=[string]$excel.Run("'$($book.Name)'!modTestSP35CrackWidth.$Macro",$acad)
    } else {
        $result=[string]$excel.Run("'$($book.Name)'!modTestAutoCADContours.$Macro",$acad,$template)
    }
    $result | Set-Content -LiteralPath (Join-Path $directory 'NativeCAD.txt') -Encoding UTF8
    [ordered]@{Executable=$fullName; OwnedPID=$pidValue; COMServer=$server; UserDocumentsUsed=$false} |
        ConvertTo-Json | Set-Content -LiteralPath (Join-Path $directory 'NativeCADIdentity.json') -Encoding UTF8
    Write-Output (($result -split "`r?`n" | Where-Object {$_ -match '^TOTAL|^FAIL:|^NATIVE_DWG:'}) -join "`n")
    if ($result -notmatch 'TOTAL[^\r\n]*failed=0\b' -or $result -match '(?m)^FAIL:') {throw 'Native CAD test failed.'}
    $nativePassed=$true
} finally {
    if ($book -and (Test-OwnedProcessIdentity $ownedExcelPid $ownedExcelTicks 'EXCEL')) {try {$book.Close($false)} catch {}}
    if ($excel) {
        if (Test-OwnedProcessIdentity $ownedExcelPid $ownedExcelTicks 'EXCEL') {try {$excel.Quit()} catch {}}
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
    }
    if ($acad) {
        if ($owned -and (Test-OwnedProcessIdentity $pidValue $cadStartTicks 'acad')) {try {Invoke-CadReady { $acad.Quit() }} catch {Write-Warning $_.Exception.Message}}
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($acad)
    }
    if ($launchedCad) {
        $remaining=Get-Process -Id $launchedCad.Id -ErrorAction SilentlyContinue
        if ($remaining -and $remaining.ProcessName -eq 'acad' -and
            $remaining.StartTime.ToUniversalTime().Ticks -eq $launchedTicks) {
            Stop-Process -Id $remaining.Id -Force
        }
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    $sourceUnchanged=(Get-FileHash -LiteralPath $sourcePath -Algorithm SHA256).Hash -eq $sourceHash
    $invocation=[ordered]@{Macro=$Macro;Mode=$Mode;SourceWorkbook=$sourcePath;SourceSHA256=$sourceHash;SourceUnchanged=$sourceUnchanged;Passed=$nativePassed;OwnedCadPID=$pidValue;LaunchedCadPID=$(if ($launchedCad) {$launchedCad.Id} else {0});LaunchedCadStartTicks=$launchedTicks;RunnerSHA256=(Get-FileHash -LiteralPath $PSCommandPath).Hash}
    [IO.File]::WriteAllText((Join-Path $directory 'NativeInvocation.json'),($invocation | ConvertTo-Json),[Text.Encoding]::UTF8)
    if (-not $sourceUnchanged) {throw 'Native source workbook changed.'}
}
