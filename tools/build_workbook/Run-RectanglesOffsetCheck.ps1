# ДЛЯ ТЕСТОВ: три знака смещения на отдельных книгах, без расчёта НДС.
param([string]$Directory='docs/regression/ExcelGuardGuide_2026-10-13/RectanglesOffset_2026-10-15')
$ErrorActionPreference='Stop'
$root=(Get-Location).Path
$evidence=[IO.Path]::GetFullPath((Join-Path $root $Directory))
if(-not $evidence.StartsWith((Join-Path $root 'docs/regression/ExcelGuardGuide_2026-10-13/'),[StringComparison]::OrdinalIgnoreCase)){throw 'Unexpected evidence directory.'}
if(Test-Path -LiteralPath $evidence){throw 'Use a fresh evidence directory.'}
New-Item -ItemType Directory -Path $evidence | Out-Null
$production=Join-Path $root 'workbook/output/RC_Section_NDM.xlsm'
$baseline=Join-Path $evidence 'Baseline.xlsm'
$lock=$null
try{$lock=[IO.File]::Open($production,'Open','Read','None')}catch [IO.IOException]{}
if($lock){$lock.Dispose()}else{
    $userBook=[Runtime.InteropServices.Marshal]::BindToMoniker($production)
    try{
        if([string]$userBook.FullName -cne $production -or $userBook.ReadOnly -or -not $userBook.Application.Ready){throw 'User workbook is unavailable.'}
        if(-not $userBook.Saved){$userBook.Save()}
        $userBook.Close($false)
    }finally{[void][Runtime.InteropServices.Marshal]::ReleaseComObject($userBook)}
}
Copy-Item -LiteralPath $production -Destination $baseline
[ordered]@{WorkbookSHA256=(Get-FileHash -LiteralPath $production).Hash;AllCodeSHA256=(Get-FileHash -LiteralPath (Join-Path $root 'workbook/output/VBA_All_Code.txt')).Hash;ExecutionReportSHA256=(Get-FileHash -LiteralPath (Join-Path $root 'workbook/output/RC_Section_NDM_execution_report.txt')).Hash} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $evidence 'Before.json') -Encoding UTF8
function Invoke-Driver([string]$Name,[string[]]$Arguments){
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot $Name) @Arguments
    if($LASTEXITCODE -ne 0){throw ('Driver failed: '+$Name)}
}
$candidateDirectory=Join-Path $evidence 'Candidate'
New-Item -ItemType Directory -Path $candidateDirectory | Out-Null
$candidate=Join-Path $candidateDirectory 'RC_Section_NDM.xlsm'
Copy-Item -LiteralPath $baseline -Destination $candidate
Invoke-Driver 'Update-PrivateWorkbookCode.ps1' @('-WorkbookPath',($Directory+'/Candidate/RC_Section_NDM.xlsm'),'-ModuleNames','CSectionRecipeRectangles')
Invoke-Driver 'Export-Audit03VbaSnapshot.ps1' @('-WorkbookPath',($Directory+'/Baseline.xlsm'),'-OutputPath',($Directory+'/BaselineNativeVBA.txt'))
Invoke-Driver 'Export-Audit03VbaSnapshot.ps1' @('-WorkbookPath',($Directory+'/Candidate/RC_Section_NDM.xlsm'),'-OutputPath',($Directory+'/CandidateNativeVBA.txt'))
Invoke-Driver 'Verify-PostAudit03VbaEquality.ps1' @('-ExportPath',($Directory+'/CandidateNativeVBA.txt'),'-ReportPath',($Directory+'/CandidateSourceEquality.json'))
Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public static class OffsetProbeWindow{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);}'
foreach($version in @('Before','After')){
    $folder=Join-Path $evidence $version
    New-Item -ItemType Directory -Path $folder | Out-Null
    $path=Join-Path $folder 'RC_Section_NDM.xlsm'
    Copy-Item -LiteralPath $(if($version -eq 'Before'){$baseline}else{$candidate}) -Destination $path
    $excel=$null;$book=$null;$ownedId=0;$ownedTicks=0
    try{
        $excel=New-Object -ComObject Excel.Application
        $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false;$excel.AutomationSecurity=3
        [uint32]$windowId=0;[void][OffsetProbeWindow]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$windowId)
        $ownedId=[int]$windowId;$ownedTicks=(Get-Process -Id $ownedId).StartTime.ToUniversalTime().Ticks
        $book=$excel.Workbooks.Open($path,0,$false)
        if($book.ReadOnly){throw 'Probe is read-only.'}
        $book.VBProject.VBComponents.Item('modTestSectionRecipes').CodeModule.AddFromString([IO.File]::ReadAllText((Join-Path $PSScriptRoot 'RectanglesOffsetProbe.bas'),[Text.Encoding]::UTF8))
        $excel.VBE.ActiveVBProject=$book.VBProject
        $compile=$excel.VBE.CommandBars.FindControl(1,578)
        if($compile.Enabled){$compile.Execute()}
        if($compile.Enabled){throw 'Offset probe compilation failed.'}
        $book.Save();$book.Close($false);$book=$null
    }finally{
        if($book){$book.Close($false)}
        if($excel){$excel.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
        [GC]::Collect();[GC]::WaitForPendingFinalizers()
        if($ownedId){$owned=Get-Process -Id $ownedId -ErrorAction SilentlyContinue;if($owned -and $owned.ProcessName -eq 'EXCEL' -and $owned.StartTime.ToUniversalTime().Ticks -eq $ownedTicks){Stop-Process -Id $ownedId -Force}}
    }
    $macro=if($version -eq 'Before'){'DumpRectanglesOffsetBefore'}else{'RunRectanglesOffsetDimensionTests'}
    Invoke-Driver 'Run-Audit03Watchdog.ps1' @('-SourceWorkbook',($Directory+'/'+$version+'/RC_Section_NDM.xlsm'),'-ReportPath',($Directory+'/'+$version+'.txt'),'-Macro',('modTestSectionRecipes.'+$macro),'-Mode','No','-TimeoutSeconds','120')
}
Write-Output 'RECTANGLES_OFFSET_NATIVE_CHECKS_FINISHED'
