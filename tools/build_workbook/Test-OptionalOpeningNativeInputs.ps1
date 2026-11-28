# ДЛЯ ТЕСТОВ: собственная видимая копия для нативной проверки новых списков, с сохранением и повторным открытием.
param([Parameter(Mandatory=$true)][string]$SourceWorkbook,[Parameter(Mandatory=$true)][string]$EvidenceDirectory,[ValidateSet(1,2)][int]$Phase=1)
$ErrorActionPreference='Stop';$root=(Get-Location).Path
$source=(Resolve-Path -LiteralPath $SourceWorkbook).Path;$directory=[IO.Path]::GetFullPath((Join-Path $root $EvidenceDirectory))
if(-not $directory.StartsWith((Join-Path $root 'docs/regression/CircleOpening_2026-11-28/'),[StringComparison]::OrdinalIgnoreCase)){throw 'Evidence path must stay inside CircleOpening.'}
[void][IO.Directory]::CreateDirectory($directory)
$memory=Get-CimInstance Win32_PerfFormattedData_PerfOS_Memory;if($memory.AvailableMBytes-lt4096-or($memory.CommitLimit-$memory.CommittedBytes)-lt7GB){throw 'Insufficient memory'}
$memory|Select-Object AvailableMBytes,CommitLimit,CommittedBytes|ConvertTo-Json|Set-Content (Join-Path $directory "Memory_$Phase.json") -Encoding UTF8
if(Test-Path (Join-Path $directory "Ready_$Phase.json")){throw 'Primary UI phase cannot be overwritten.'}
$path=Join-Path $directory 'OptionalOpening_UI.xlsm';if($Phase-eq1){if(Test-Path -LiteralPath $path){throw 'UI copy already exists.'};Copy-Item -LiteralPath $source -Destination $path}
Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public static class OpeningNativeTestWindow{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);}'
$app=$null;$book=$null;$owned=0;$ticks=0;$fields=@('E23','F23','I23','J23')
try{
 $previous=@(Get-Process EXCEL -ErrorAction SilentlyContinue|ForEach-Object Id)
 $app=New-Object -ComObject Excel.Application;$app.AutomationSecurity=3;$app.EnableEvents=$false;$app.DisplayAlerts=$true
 [uint32]$id=0;[void][OpeningNativeTestWindow]::GetWindowThreadProcessId([IntPtr]$app.Hwnd,[ref]$id);if($previous-contains[int]$id){$app=$null;throw 'Independent own Excel required'}
 $owned=[int]$id;$ticks=(Get-Process -Id $owned).StartTime.ToUniversalTime().Ticks
 $book=$app.Workbooks.Open($path,0,$false);if($book.ReadOnly){throw 'Native UI copy must be writable.'};$app.Visible=$true
 $sheet=$book.Worksheets.Item('Сечения');$initial=@{};$validation=@{}
 foreach($field in $fields){$initial[$field]=$sheet.Range($field).Value2;$validation[$field]=[ordered]@{Type=$sheet.Range($field).Validation.Type;Formula1=$sheet.Range($field).Validation.Formula1;ShowError=$sheet.Range($field).Validation.ShowError}}
 @{Phase=$Phase;PID=$owned;StartTicksUTC=[string]$ticks;Workbook=$path;SourceSHA256=(Get-FileHash -LiteralPath $source).Hash;Initial=$initial;Validation=$validation}|ConvertTo-Json -Depth 5|Set-Content (Join-Path $directory "Ready_$Phase.json") -Encoding UTF8
 Write-Output "NATIVE UI READY phase $Phase"
 $clock=[Diagnostics.Stopwatch]::StartNew();$finish=Join-Path $directory "Finish_$Phase.json"
 while(-not(Test-Path -LiteralPath $finish)){if($clock.Elapsed.TotalMinutes-gt15){throw 'Native UI phase expired.'};Start-Sleep -Milliseconds 300}
 $saved=@{};foreach($field in $fields){$saved[$field]=$sheet.Range($field).Value2}
 $book.Save();$book.Close($false);$book=$null
 @{Phase=$Phase;Saved=$saved;SavedFileSHA256=(Get-FileHash -LiteralPath $path).Hash}|ConvertTo-Json -Depth 4|Set-Content (Join-Path $directory "Closed_$Phase.json") -Encoding UTF8
}finally{
 if($book){$book.Close($false)}
 if($app){if($app.Workbooks.Count-eq0){$app.Quit()};[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
