# ДЛЯ ТЕСТОВ: открывает частную копию для нативной проверки списка интегрирования.
# Ввод выполняется через Computer Use; скрипт только подготавливает книгу и сохраняет её по файлу Close_N.request.
param([int]$Phase=1,[string]$EvidenceDirectory='docs/regression/MTLI_2026-11-28/DropdownNative01')
$ErrorActionPreference='Stop'
$root=(Get-Location).Path
$dir=Join-Path $root $EvidenceDirectory
New-Item -ItemType Directory -Path $dir -Force|Out-Null
$path=Join-Path $dir 'MTLI_Dropdown.xlsm'
$ready=Join-Path $dir ('Ready_'+$Phase+'.json')
$close=Join-Path $dir ('Close_'+$Phase+'.request')
if(Test-Path -LiteralPath $ready){throw 'Native UI phase already has primary evidence'}
if($Phase -eq 1){Copy-Item -LiteralPath (Join-Path $root 'workbook/experimental/RC_Section_NDM_MTLI.xlsm') -Destination $path}
$memory=Get-CimInstance Win32_PerfFormattedData_PerfOS_Memory
if($memory.AvailableMBytes -lt 4096 -or ($memory.CommitLimit-$memory.CommittedBytes) -lt 7GB){throw 'Insufficient memory before native UI Excel'}
$memory|Select-Object AvailableMBytes,CommitLimit,CommittedBytes|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $dir ('Memory_'+$Phase+'.json')) -Encoding UTF8
Add-Type 'using System;using System.Runtime.InteropServices;public static class MTLITestDropdownWindow{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr hwnd,out uint id);}'
$previous=@(Get-Process EXCEL -ErrorAction SilentlyContinue|ForEach-Object Id)
$excel=$null;$book=$null;$owned=0;$ticks=0
try {
    $excel=New-Object -ComObject Excel.Application
    [uint32]$id=0;[void][MTLITestDropdownWindow]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$id)
    $owned=[int]$id;$ticks=(Get-Process -Id $owned).StartTime.ToUniversalTime().Ticks
    if($owned -in $previous){throw 'Existing user Excel reused'}
    $excel.Visible=$true;$excel.DisplayAlerts=$true;$excel.EnableEvents=$false;$excel.AutomationSecurity=3
    $book=$excel.Workbooks.Open($path,0,$false)
    $range=$book.Names.Item('rngExperimentalNumericsSettings').RefersToRange
    $record=[ordered]@{Phase=$Phase;PID=$owned;StartTicksUTC=[string]$ticks;Path=$path;MethodAddress=($range.Worksheet.Name+'!'+$range.Cells.Item(2,2).Address($false,$false));InitialMethod=[string]$range.Cells.Item(2,2).Value2;InitialTolerance=[string]$range.Cells.Item(3,2).Value2;ListFormula=[string]$range.Cells.Item(2,2).Validation.Formula1}
    $record|ConvertTo-Json|Set-Content -LiteralPath $ready -Encoding UTF8
    $deadline=[DateTime]::UtcNow.AddMinutes(20)
    while(-not(Test-Path -LiteralPath $close)){
        if([DateTime]::UtcNow -ge $deadline){throw 'Native UI observation deadline reached'}
        Start-Sleep -Seconds 1
    }
    $record.FinalMethod=[string]$range.Cells.Item(2,2).Value2
    $record.FinalTolerance=[string]$range.Cells.Item(3,2).Value2
    $record.FinalList=[string]$range.Cells.Item(2,2).Validation.Formula1
    $book.Save();$book.Close($false);$book=$null
    $record.SHA256=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
    $record|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $dir ('Closed_'+$Phase+'.json')) -Encoding UTF8
} finally {
    if($book){$book.Close($false)}
    if($excel){
        if($excel.Workbooks.Count -eq 0){$excel.Quit()}
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
    }
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
