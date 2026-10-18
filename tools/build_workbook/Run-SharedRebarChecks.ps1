# Направленные проверки переноса армирования; исходная книга не сохраняется.
param([Parameter(Mandatory=$true)][string]$WorkbookPath,
      [Parameter(Mandatory=$true)][string]$EvidenceDirectory,
      [switch]$InputTests,[switch]$CalculationsOnly,[string[]]$OnlyMacros=@())
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$path=(Resolve-Path -LiteralPath (Join-Path $root $WorkbookPath)).Path
$directory=[IO.Path]::GetFullPath((Join-Path $root $EvidenceDirectory))
if(-not $directory.StartsWith($root+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)){throw 'Evidence must stay in the workspace.'}
[void][IO.Directory]::CreateDirectory($directory)
$before=(Get-FileHash -LiteralPath $path).Hash
$existing=@(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)
Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public static class SharedRebarChecksWindow{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr hWnd,out uint id);}'
$excel=$null;$book=$null;$ownedPid=0;$ticks=0;$receipts=@()
try{
    $excel=New-Object -ComObject Excel.Application
    [uint32]$id=0;[void][SharedRebarChecksWindow]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$id)
    $ownedPid=[int]$id;if($ownedPid -in $existing){throw 'Separate Excel process required.'}
    $ticks=(Get-Process -Id $ownedPid).StartTime.ToUniversalTime().Ticks
    @{PID=$ownedPid;StartTicks=$ticks} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $directory 'Ownership.json') -Encoding UTF8
    $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false;$excel.AutomationSecurity=1
    $book=$excel.Workbooks.Open($path,0,$true)
    $compile=$excel.VBE.CommandBars.FindControl(1,578);if($compile -and $compile.Enabled){$compile.Execute()}
    $macros=@('modTestSectionRecipes.DumpCircleRecipeMatrix','modTestSectionRecipes.DumpRectSetRecipeMatrix',
        'modTestSectionRecipes.DumpRoundedRecipeMatrix','modTestSectionRecipes.DumpHollowRecipeMatrix',
        'modTestSectionRecipes.DumpHollowRecipeEdgeMatrix','modTestAudit02Snapshots.DumpRecipePathContracts')
    if($InputTests){$macros+=@('modTestSectionInputReader.RunSectionInputReaderTests','modTestSectionRecipes.RunRecipeDefinitionTests',
        'modTestSectionRecipes.RunRecipeAnnotationCountTests','modTestSectionRecipes.RunCircleBindingTests')}
    if($CalculationsOnly){$macros=@('modTestAudit02Snapshots.DumpAudit02Calculations:SP35','modTestAudit02Snapshots.DumpAudit02Calculations:SP63')}
    if($OnlyMacros.Count){$macros=$OnlyMacros}
    foreach($entry in $macros){
        $parts=$entry -split ':';$macro=$parts[0]
        $watch=[Diagnostics.Stopwatch]::StartNew()
        Write-Output ('RUN: '+$macro)
        $name="'"+$book.Name+"'!"+$macro
        $suffix='';if($parts.Count -gt 1){$body=[string]$excel.Run($name,[string]$parts[1]);$suffix='_'+$parts[1]}else{$body=[string]$excel.Run($name)}
        $file=Join-Path $directory (($macro -split '\.')[-1]+$suffix+'.txt')
        [IO.File]::WriteAllText($file,$body,[Text.UTF8Encoding]::new($false))
        if($body -match '(?m)^FAIL:|RUNTIME ERROR|failed=[1-9]'){throw ('Failed checks: '+$macro)}
        if(-not $body.Trim()){throw ('Empty result: '+$macro)}
        $receipts+=@{Macro=$entry;Seconds=$watch.Elapsed.TotalSeconds;SHA256=(Get-FileHash -LiteralPath $file).Hash}
        Write-Output ('PASS: '+$macro+'; seconds='+[Math]::Round($watch.Elapsed.TotalSeconds,2))
    }
}finally{
    if($book){$book.Close($false)}
    if($excel){$excel.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
    if($ownedPid -and $ticks){$process=Get-Process -Id $ownedPid -ErrorAction SilentlyContinue;if($process -and $process.ProcessName -eq 'EXCEL' -and $process.StartTime.ToUniversalTime().Ticks -eq $ticks){Stop-Process -Id $ownedPid -Force}}
}
if((Get-FileHash -LiteralPath $path).Hash -cne $before){throw 'Source workbook changed.'}
@{Passed=$true;WorkbookSHA256=$before;Checks=$receipts} | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directory 'Checks.json') -Encoding UTF8
