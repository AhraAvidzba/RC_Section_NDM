# Адресные проверки и снимки в собственной копии; рабочая книга и пользовательский отчёт защищены hash.
param(
    [Parameter(Mandatory=$true)][string]$SourceWorkbook,
    [Parameter(Mandatory=$true)][string]$Stage,
    [Parameter(Mandatory=$true)][string[]]$Macros,
    [string]$Argument1='',
    [switch]$Snapshots
)
$ErrorActionPreference='Stop'
$root=(Get-Location).Path
$evidence=Join-Path $root 'docs/regression/ExcelGuardGuide_2026-10-13/RecipeStylePlot_2026-10-14'
if($Stage -notmatch '^[A-Za-z][A-Za-z0-9_-]*$'){throw 'Stage must be a local directory name.'}
$directory=Join-Path $evidence $Stage
New-Item -ItemType Directory -Path $directory -Force | Out-Null
$copy=Join-Path $directory 'GeometrySolidProbe.xlsm'
if(Test-Path -LiteralPath $copy){throw 'Run copy already exists; choose a new stage.'}
$source=(Resolve-Path -LiteralPath (Join-Path $root $SourceWorkbook)).Path
if(-not $source.StartsWith($evidence+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)){throw 'Only the recorded evidence source is permitted.'}
$protected=@('workbook/output/RC_Section_NDM.xlsm','workbook/output/VBA_All_Code.txt','workbook/output/RC_Section_NDM_execution_report.txt')
$before=@{}
foreach($path in $protected){$before[$path]=(Get-FileHash -LiteralPath (Join-Path $root $path)).Hash}
Copy-Item -LiteralPath $source -Destination $copy
$excel=$null;$book=$null;$ownedPid=0;$ownedTicks=0;$runs=@();$success=$false
Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public static class SolidExcelIdentity{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr hWnd,out uint id);}'
try {
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false;$excel.AutomationSecurity=1
    [uint32]$id=0;[void][SolidExcelIdentity]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$id)
    $ownedPid=[int]$id;$ownedTicks=(Get-Process -Id $ownedPid).StartTime.ToUniversalTime().Ticks
    [ordered]@{PID=$ownedPid;StartTimeTicks=$ownedTicks} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $directory 'OwnedExcel.json') -Encoding UTF8
    Write-Output ('OWNED_EXCEL: '+$ownedPid)
    $book=$excel.Workbooks.Open($copy,0,$false)
    $excel.Calculation=-4135
    foreach($macro in @($Macros | ForEach-Object {$_ -split ','})) {
        Write-Output ('RUN '+$macro)
        $watch=[Diagnostics.Stopwatch]::StartNew()
        $qualified="'$($book.Name)'!$macro"
        if($Argument1){$result=[string]$excel.Run($qualified,$Argument1)}else{$result=[string]$excel.Run($qualified)}
        $watch.Stop()
        $result.TrimEnd([char[]]"`r`n") | Set-Content -LiteralPath (Join-Path $directory ($macro+'.txt')) -Encoding UTF8
        $passed=$result -notmatch '(?m)^FAIL[: ]|failed=[1-9]|RUNTIME ERROR:|^ERROR:'
        if(-not $Snapshots){$passed=$passed -and $result -match '(?m)^TOTAL[^\r\n]*failed=0\b'}
        if($Snapshots){$passed=$passed -and $result.Length -gt 0}
        $runs+=@([ordered]@{Macro=$macro;Seconds=$watch.Elapsed.TotalSeconds;Characters=$result.Length;Passed=$passed;Snapshot=[bool]$Snapshots})
        if($Snapshots){Write-Output ('SNAPSHOT_CHARS: '+$result.Length)}else{Write-Output (($result -split "`r?`n" | Where-Object {$_ -match '^TOTAL|^FAIL|^RUNTIME|^OBSERVED'}) -join "`n")}
        if(-not $passed){throw ('Macro did not pass: '+$macro)}
    }
    $success=$true
} finally {
    if($book){try{$book.Close($false)}finally{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)}}
    if($excel){try{$excel.Quit()}finally{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
    if($ownedPid){$owned=Get-Process -Id $ownedPid -ErrorAction SilentlyContinue;if($owned -and $owned.ProcessName -eq 'EXCEL' -and $owned.StartTime.ToUniversalTime().Ticks -eq $ownedTicks){Stop-Process -Id $ownedPid -Force}}
    foreach($path in $before.Keys){if((Get-FileHash -LiteralPath (Join-Path $root $path)).Hash -cne $before[$path]){throw ('Protected production file changed: '+$path)}}
    [ordered]@{Stage=$Stage;Passed=$success;Runs=$runs;SourceSHA256=(Get-FileHash -LiteralPath $source).Hash;ProductionUnchanged=$true;FinishedAt=(Get-Date).ToString('o')} | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directory 'Run.json') -Encoding UTF8
}
