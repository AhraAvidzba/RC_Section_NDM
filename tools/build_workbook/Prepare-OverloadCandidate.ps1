# Обновляет только перечисленные компоненты отдельной измерительной копии.
param(
    [string]$SourceWorkbook='docs/regression/Performance/Overload_2026-10-15/ProbeBaseline02/RC_Section_NDM.xlsm',
    [string]$OutputDirectory='docs/regression/Performance/Overload_2026-10-15/ProbeCandidate',
    [switch]$Refresh
)
$ErrorActionPreference='Stop'
$root=(Get-Location).Path
$source=(Resolve-Path -LiteralPath $SourceWorkbook).Path
$directory=[IO.Path]::GetFullPath((Join-Path $root $OutputDirectory))
if(-not $directory.StartsWith((Join-Path $root 'docs/regression/Performance/'),[StringComparison]::OrdinalIgnoreCase)){throw 'Performance directory required.'}
New-Item -ItemType Directory -Path $directory -Force | Out-Null
$path=Join-Path $directory 'RC_Section_NDM.xlsm'
if(Test-Path -LiteralPath $path){if(-not $Refresh){throw 'Candidate already exists.'}} else {Copy-Item -LiteralPath $source -Destination $path}
$excel=$null;$book=$null;$ownedPid=0;$ownedTicks=0
Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public static class OverloadCandidateWindow{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr w,out uint id);}'
try {
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false;$excel.AutomationSecurity=3
    [uint32]$windowPid=0;[void][OverloadCandidateWindow]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$windowPid)
    $ownedPid=[int]$windowPid;$ownedTicks=(Get-Process -Id $ownedPid).StartTime.ToUniversalTime().Ticks
    Write-Output ('OWNED_EXCEL: id='+$ownedPid+'; utcStartTicks='+$ownedTicks)
    $book=$excel.Workbooks.Open($path,0,$false)
    $files=@('src/Solver/CCapacitySolver.cls','src/Solver/CCapacityLimitSearchProblem.cls','src/Solver/CLimitSearchCoordinator.cls','src/Solver/CUltimateStrainSearch.cls','tests/modTestCapacitySolver.bas')
    foreach($file in $files){
        $name=[IO.Path]::GetFileNameWithoutExtension($file)
        $body=[IO.File]::ReadAllText((Join-Path $root $file),[Text.Encoding]::UTF8)
        $body=$body.Substring($body.IndexOf('Option Explicit'))
        $module=$book.VBProject.VBComponents.Item($name).CodeModule
        $module.DeleteLines(1,$module.CountOfLines);$module.AddFromString($body)
    }
    $excel.VBE.ActiveVBProject=$book.VBProject
    $compile=$excel.VBE.CommandBars.FindControl(1,578)
    if($compile.Enabled){$compile.Execute()}
    if($compile.Enabled){throw 'Candidate compilation failed.'}
    $book.Save();$book.Close($false);$book=$null
    [ordered]@{Compiled=$true;UpdatedSources=$files;SourceSHA256=(Get-FileHash -LiteralPath $source).Hash;CandidateSHA256=(Get-FileHash -LiteralPath $path).Hash;ProductionWorkbookChanged=$false} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $directory 'Preparation.json') -Encoding UTF8
} finally {
    if($book){$book.Close($false)}
    if($excel){$excel.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
    if($ownedPid){$owned=Get-Process -Id $ownedPid -ErrorAction SilentlyContinue; if($owned -and $owned.ProcessName -eq 'EXCEL' -and $owned.StartTime.ToUniversalTime().Ticks -eq $ownedTicks){Stop-Process -Id $ownedPid -Force}}
}
