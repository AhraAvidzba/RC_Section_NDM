# Read-only native output verification and actual AllCode export. No imports or Save.
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../../..')).Path
$source=Join-Path $root 'workbook/output/RC_Section_NDM.xlsm'
$hash=(Get-FileHash -LiteralPath $source).Hash
$memory=Get-CimInstance Win32_OperatingSystem
if($memory.FreePhysicalMemory*1KB -lt 2GB -or $memory.FreeVirtualMemory*1KB -lt 2GB){throw 'Insufficient reserve; no Excel started.'}
Add-Type 'using System;using System.Runtime.InteropServices;public static class NdmNativeMessageOwner{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);}'
$beforePids=@(Get-Process EXCEL -ErrorAction SilentlyContinue|ForEach-Object Id)
$app=$null;$book=$null;$owned=0;$ticks=0;$records=@();$bodies=[ordered]@{};$text=New-Object Text.StringBuilder
$proof=[ordered]@{At=(Get-Date).ToString('o');Status='STARTED';SourceSHA256=$hash;FreePhysicalGiB=[math]::Round($memory.FreePhysicalMemory*1KB/1GB,3);FreeCommitGiB=[math]::Round($memory.FreeVirtualMemory*1KB/1GB,3);ImportPerformed=$false;MacrosRun=$false;SavePerformed=$false}
try{
    $app=New-Object -ComObject Excel.Application
    [uint32]$nativeId=0;[void][NdmNativeMessageOwner]::GetWindowThreadProcessId([IntPtr]$app.Hwnd,[ref]$nativeId)
    $owned=[int]$nativeId;$ticks=(Get-Process -Id $owned).StartTime.ToUniversalTime().Ticks
    if($owned -in $beforePids){throw 'New owned Excel required.'}
    $proof.PID=$owned;$proof.StartTicksUtc=[string]$ticks
    $proof|ConvertTo-Json -Depth 5|Set-Content -LiteralPath (Join-Path $PSScriptRoot 'NativeVerification.json') -Encoding UTF8
    $app.Visible=$false;$app.DisplayAlerts=$false;$app.EnableEvents=$false;$app.AutomationSecurity=3
    $book=$app.Workbooks.Open($source,0,$true)
    if(-not $book.ReadOnly){throw 'Read-only required.'}
    $proof.ReadOnly=$book.ReadOnly;$proof.ComponentCount=$book.VBProject.VBComponents.Count
    if($proof.ComponentCount -ne 178){throw 'Unexpected component count.'}
    [void]$text.AppendLine('VBA PROJECT EXPORT');[void]$text.AppendLine('Workbook: '+$book.Name);[void]$text.AppendLine('Date: '+(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
    foreach($component in $book.VBProject.VBComponents){
        $count=$component.CodeModule.CountOfLines
        $body=if($count){$component.CodeModule.Lines(1,$count)}else{''}
        $bodies[$component.Name]=$body
        [void]$text.AppendLine(('='*100));[void]$text.AppendLine('COMPONENT: '+$component.Name);[void]$text.AppendLine('TYPE: '+$component.Type);[void]$text.AppendLine('LINES: '+$count);[void]$text.AppendLine(('='*100));if($count){[void]$text.AppendLine($body)};[void]$text.AppendLine()
        $records+=@{Name=$component.Name;Type=$component.Type;Lines=$count}
    }
    $proof.NamesCount=$book.Names.Count;$proof.Sheets=@($book.Worksheets|ForEach-Object Name)
    $config=$book.Worksheets('Config');$proof.CaptionH53=[string]$config.Range('H53').Value2;$proof.CaptionH58=[string]$config.Range('H58').Value2
    $proof.SourceEventTemplate='tools/build_workbook/WorkbookHelpEvents.vba';$proof.DocumentModuleName=[string]$book.CodeName
    $book.Close($false);$book=$null
    [IO.File]::WriteAllText((Join-Path $PSScriptRoot 'NativeSnapshot.txt'),$text.ToString(),(New-Object Text.UTF8Encoding($true)))
    $bodies|ConvertTo-Json -Depth 3|Set-Content -LiteralPath (Join-Path $PSScriptRoot 'PublishedNativeCode.json') -Encoding UTF8
    $proof.Status='PASS_NATIVE_READONLY_EXPORT';$proof.Components=$records
}finally{
    if($book){$book.Close($false)}
    if($app){$app.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
$proof.SourceUnchanged=(Get-FileHash -LiteralPath $source).Hash -ceq $hash
if(-not $proof.SourceUnchanged){throw 'Output changed during read-only verification.'}
$proof.FullRegressionPass=$false
$proof|ConvertTo-Json -Depth 6|Set-Content -LiteralPath (Join-Path $PSScriptRoot 'NativeVerification.json') -Encoding UTF8
Write-Output ('Native output read-only verified: components='+$proof.ComponentCount+'; SHA256='+$hash)
