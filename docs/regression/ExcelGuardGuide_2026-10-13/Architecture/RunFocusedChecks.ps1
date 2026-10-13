# Focused architecture probes on an owned copy; production files are never written.
param([string]$CopyName='ArchitectureProbe_02.xlsm', [string]$Receipt='RuntimeChecks.json', [switch]$AnnotationsOnly)
$ErrorActionPreference='Stop'
$root=(Get-Location).Path
$relative='docs/regression/ExcelGuardGuide_2026-10-13/Architecture'
$directory=Join-Path $root $relative
$source=Join-Path $root 'workbook/output/RC_Section_NDM.xlsm'
$report=Join-Path $root 'workbook/output/RC_Section_NDM_execution_report.txt'
$allCode=Join-Path $root 'workbook/output/VBA_All_Code.txt'
if($CopyName -notmatch '^ArchitectureProbe_\d+\.xlsm$' -or $Receipt -notmatch '^\w+\.json$'){throw 'Evidence file names must be local leaves.'}
$copy=Join-Path $directory $CopyName
if(Test-Path -LiteralPath $copy){throw 'Private copy exists; do not replace.'}
$before=@{}
foreach($path in @($source,$report,$allCode)){$before[$path]=(Get-FileHash -LiteralPath $path).Hash}
Copy-Item -LiteralPath $source -Destination $copy
$excel=$null;$book=$null;$ownedPid=0;$ownedTicks=0
$runs=@()
Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public static class ArchitectureExcelIdentity{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr hWnd,out uint id);}'
try {
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false
    $excel.AutomationSecurity=1
    [uint32]$id=0;[void][ArchitectureExcelIdentity]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$id)
    $ownedPid=[int]$id;$ownedTicks=(Get-Process -Id $ownedPid).StartTime.ToUniversalTime().Ticks
    [ordered]@{PID=$ownedPid;StartTimeTicks=$ownedTicks} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $directory 'OwnedExcel.json') -Encoding UTF8
    Write-Output ('OWNED_EXCEL: '+$ownedPid)
    $book=$excel.Workbooks.Open($copy,0,$false)
    $excel.Calculation=-4135
    $text=[IO.File]::ReadAllText((Join-Path $directory 'modGeometryArchitectureProbe.bas'),[Text.Encoding]::UTF8)
    $body=[regex]::Match($text,'(?ms)^Option Explicit.*').Value
    $component=$book.VBProject.VBComponents.Add(1);$component.Name='modGeometryArchitectureProbe'
    $component.CodeModule.AddFromString(($body -split "`r?`n") -join "`r`n")
    $book.Save()
    $macros=@('modGeometryArchitectureProbe.RunGeometryArchitectureProbe','modTestDeclarativeGeometry.RunCoreGeometryTests','modTestSectionRecipes.RunRecipeDefinitionTests','modTestSectionRecipes.RunRecipeRuleTests','modTestSectionRecipes.RunRecipeMeshBoundaryTests','modTestSectionRecipes.RunRecipeAnnotationCountTests')
    if($AnnotationsOnly){$macros=@('modGeometryArchitectureProbe.RunAnnotationOwnershipProbe')}
    foreach($macro in $macros) {
        Write-Output ('RUN '+$macro)
        $watch=[Diagnostics.Stopwatch]::StartNew()
        $result=[string]$excel.Run("'$($book.Name)'!$macro")
        $watch.Stop()
        $result.TrimEnd([char[]]"`r`n") | Set-Content -LiteralPath (Join-Path $directory ($macro+'.txt')) -Encoding UTF8
        $success=$result -match '(?m)^TOTAL[^\r\n]*failed=0\b' -and $result -notmatch '(?m)^FAIL[: ]|failed=[1-9]|RUNTIME ERROR:'
        $runs+=@([ordered]@{Macro=$macro;Seconds=$watch.Elapsed.TotalSeconds;Passed=$success})
        Write-Output (($result -split "`r?`n" | Where-Object {$_ -match '^TOTAL|^OBSERVED|^FAIL|^RUNTIME'}) -join "`n")
        if(-not $success){throw ('Probe failed: '+$macro)}
    }
} finally {
    if($book){$book.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)}
    if($excel){$excel.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
    if($ownedPid){$owned=Get-Process -Id $ownedPid -ErrorAction SilentlyContinue;if($owned -and $owned.ProcessName -eq 'EXCEL' -and $owned.StartTime.ToUniversalTime().Ticks -eq $ownedTicks){Stop-Process -Id $ownedPid -Force}}
    foreach($path in $before.Keys){if((Get-FileHash -LiteralPath $path).Hash -cne $before[$path]){throw ('Production file changed during audit: '+$path)}}
    [ordered]@{Runs=$runs;ProductionWorkbookUnchanged=$true;AllCodeUnchanged=$true;ExecutionReportUnchanged=$true;WorkbookSHA256=$before[$source];FinishedAt=(Get-Date).ToString('o')} | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directory $Receipt) -Encoding UTF8
}
