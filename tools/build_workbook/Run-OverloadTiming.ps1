# Направленный замер перегрузок: штатный запуск с Results, отчётом и схемой.
param(
    [string]$WorkbookPath='docs/regression/Performance/Overload_2026-10-15/ProbeBaseline02/RC_Section_NDM.xlsm',
    [string]$Directory='docs/regression/Performance/Overload_2026-10-15/Baseline',
    [string[]]$Shapes=@('SavedWorkbook'),
    [string[]]$Families=@('Axial','Mx','My','Mixed'),
    [double[]]$Overloads=@(2,3,4),
    [string[]]$Strategies=@('Auto'),
    [int]$Repeats=1,
    [switch]$UserCircleCase,
    [switch]$CircleFixture,
    [double]$MeshStep=0
)
$ErrorActionPreference='Stop'
$root=(Get-Location).Path
$source=(Resolve-Path -LiteralPath $WorkbookPath).Path
$output=[IO.Path]::GetFullPath((Join-Path $root $Directory))
if(-not $output.StartsWith((Join-Path $root 'docs/regression/Performance/'),[StringComparison]::OrdinalIgnoreCase)){throw 'Performance evidence directory required.'}
New-Item -ItemType Directory -Path $output -Force | Out-Null
$sourceHash=(Get-FileHash -LiteralPath $source).Hash
$records=New-Object 'System.Collections.Generic.List[object]'
$excel=$null; $book=$null; $ownedPid=0; $ownedTicks=0
$culture=[Globalization.CultureInfo]::InvariantCulture
Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public static class OverloadTimingWindow{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr w,out uint id);}'
function Measure-Case([string]$shape,[string]$id,[double[]]$load,[string]$path,[string]$strategy){
    $numbers=@($load | ForEach-Object {$_.ToString('R',$culture)})
    $specification=(@($shape,$id)+$numbers+@($path,$strategy)) -join '|'
    Write-Output ('START '+$id+'; '+$specification)
    $raw=[string]$excel.Run("'$($book.Name)'!modTestBatchCalculation.TestRunOverloadTimingCase",$specification)
    [IO.File]::WriteAllText((Join-Path $output ($id+'.txt')),$raw,(New-Object Text.UTF8Encoding($true)))
    $report=Join-Path (Split-Path -Parent $source) 'RC_Section_NDM_execution_report.txt'
    if(Test-Path -LiteralPath $report){Copy-Item -LiteralPath $report -Destination (Join-Path $output ($id+'.execution.txt')) -Force}
    if($raw -notmatch '^OVERLOAD\|'){throw ('Case failed: '+$raw)}
    $row=[ordered]@{Shape=$shape;Specification=$specification}
    foreach($part in ($raw -split '\r?\n')[0].Split('|')){
        $pair=$part.Split('=',2)
        if($pair.Count -eq 2){$row[$pair[0]]=$pair[1]}
    }
    $records.Add($row)
    ConvertTo-Json -InputObject @($records.ToArray()) -Depth 5 | Set-Content -LiteralPath (Join-Path $output 'Measurements.json') -Encoding UTF8
    Write-Output (($raw -split '\r?\n')[0])
    return ,$row
}
try {
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false; $excel.DisplayAlerts=$false; $excel.EnableEvents=$false; $excel.AutomationSecurity=1
    [uint32]$windowPid=0; [void][OverloadTimingWindow]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$windowPid)
    $ownedPid=[int]$windowPid; $ownedTicks=(Get-Process -Id $ownedPid).StartTime.ToUniversalTime().Ticks
    [ordered]@{ExcelPID=$ownedPid;StartTicks=$ownedTicks;SourceSHA256=$sourceHash;ExcelVersion=[string]$excel.Version;ExcelBuild=[string]$excel.Build} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $output 'Environment.json') -Encoding UTF8
    $book=$excel.Workbooks.Open($source,0,$false)
    $excel.Calculation=-4135
    if($MeshStep -gt 0){
        $system=$book.Names.Item('rngSystemSettings').RefersToRange
        for($row=2;$row -le $system.Rows.Count;$row++){
            if([string]$system.Cells.Item($row,1).Value2 -in @('Mesh.StepX','Mesh.StepY')){$system.Cells.Item($row,2).Value2=$MeshStep}
        }
    }
    if($UserCircleCase -or $CircleFixture){
        $circle=$book.Names.Item('rngCircleGeometry').RefersToRange
        $circleValues=[ordered]@{}
        $requested=@{'Circle.Diameter'=300;'Rebar.Count'=8;'Rebar.Diameter'=20;'Rebar.Diameter2'=0;'Rebar.Diameter3'=0}
        for($row=2;$row -le $circle.Rows.Count;$row++){
            $key=[string]$circle.Cells.Item($row,1).Value2
            if($requested.ContainsKey($key)){$circle.Cells.Item($row,2).Value2=[double]$requested[$key]}
            $circleValues[$key]=$circle.Cells.Item($row,2).Value2
        }
        $circleValues | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $output 'CircleInputs.json') -Encoding UTF8
    }
    if($UserCircleCase){
        for($repeat=1;$repeat -le $Repeats;$repeat++){
        foreach($axis in @('Mx','My')){
            [double[]]$load=@(-4903325,980665000,0)
            if($axis -eq 'My'){$load=@(-4903325,0,980665000)}
            $id='UserCircle500N100'+$axis
            if($Repeats -gt 1){$id+='_r'+$repeat}
            Measure-Case 'Circle' $id $load 'Auto' 'Auto' | ForEach-Object {if($_ -is [string]){Write-Output $_}}
        }
        }
    } else {
    foreach($shape in $Shapes){foreach($family in $Families){
        [double[]]$seed=switch($family){
            'Axial' {@(-4903325,0,0)}
            'Mx' {@(0,980665000,0)}
            'My' {@(0,0,980665000)}
            'Mixed' {@(-1961330,490332500,735498750)}
            'UserRatio' {@(-4903325,980665000,0)}
            default {throw 'Unknown load family.'}
        }
        $prefix=$shape+'_'+$family
        $calibration=Measure-Case $shape ($prefix+'_calibration') $seed ("$([char]955)*NMxy") 'Auto'
        $calibration=$calibration | Select-Object -Last 1
        $lambda=[double]::Parse($calibration['lambda'],$culture)
        if($calibration['capacity'] -notin @('OK','FAIL') -or $lambda -le 0){throw ('Calibration did not confirm a capacity: '+$prefix)}
        foreach($factor in $Overloads){foreach($strategy in $Strategies){for($repeat=1;$repeat -le $Repeats;$repeat++){
            [double[]]$load=@($seed | ForEach-Object {$_*$lambda*$factor})
            $id=$prefix+'_'+$factor.ToString('0.###',$culture)+'x_'+$strategy+'_'+$repeat
            Measure-Case $shape $id $load 'Auto' $strategy | ForEach-Object {if($_ -is [string]){Write-Output $_}}
        }}}
    }}
    }
} finally {
    if($book){$book.Close($false)}
    if($excel){$excel.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
    if($ownedPid){$owned=Get-Process -Id $ownedPid -ErrorAction SilentlyContinue; if($owned -and $owned.ProcessName -eq 'EXCEL' -and $owned.StartTime.ToUniversalTime().Ticks -eq $ownedTicks){Stop-Process -Id $ownedPid -Force}}
}
if((Get-FileHash -LiteralPath $source).Hash -cne $sourceHash){throw 'Probe workbook changed on disk.'}
[ordered]@{Status='OVERLOAD_MEASUREMENTS_COMPLETE';Count=$records.Count;SourceUnchanged=$true} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $output 'Completion.json') -Encoding UTF8
