# ДЛЯ ТЕСТОВ: готовит частную книгу из согласованной main, не меняя пользовательский Excel.
param([string]$Label='Compile01')
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSHOME 'Modules/Microsoft.PowerShell.Utility/Microsoft.PowerShell.Utility.psd1') -Global
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../../..')).Path
$directory=Join-Path $PSScriptRoot $Label
if(Test-Path -LiteralPath $directory){throw 'Use a fresh preparation label.'}
[void][IO.Directory]::CreateDirectory($directory)
$baseline=Join-Path $directory 'AcceptedMain.xlsm'
$candidate=Join-Path $directory 'RC_Section_NDM.xlsm'
$python='C:/Users/avidzba/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/python.exe'
$extractor=Join-Path $directory 'ExtractAccepted.py'
[IO.File]::WriteAllText($extractor,'import subprocess,sys; f=open(sys.argv[1],"xb"); subprocess.run(["git","show","1a0c03a82ecb2fdbbabdbdb1e2bff8f554136516:workbook/output/RC_Section_NDM.xlsm"],stdout=f,check=True); f.close()',[Text.UTF8Encoding]::new($false))
& $python $extractor $baseline
if($LASTEXITCODE -ne 0){throw 'Cannot read accepted workbook from Git.'}
$baselineHash=(Get-FileHash -LiteralPath $baseline).Hash
if($baselineHash -ne '53919E6CF11018B769191059E17977D60F3201D5857991FFE49E9326E1CC613A'){throw 'Accepted source hash mismatch.'}
Copy-Item -LiteralPath $baseline -Destination $candidate
$paths=@(@(git -C $root diff --name-only main -- src tests)+@(git -C $root ls-files --others --exclude-standard -- src tests) | Sort-Object -Unique)
$inputs=@()
foreach($relative in $paths){
    if([IO.Path]::GetExtension($relative) -notin '.cls','.bas'){continue}
    $source=Join-Path $root $relative
    $body=[regex]::Match([IO.File]::ReadAllText($source,[Text.Encoding]::UTF8),'(?ms)^Option Explicit.*').Value
    if(-not $body){throw ('Missing Option Explicit: '+$relative)}
    $copy=Join-Path (Join-Path $directory 'Inputs') $relative
    [void][IO.Directory]::CreateDirectory((Split-Path -Parent $copy))
    Copy-Item -LiteralPath $source -Destination $copy
    $inputs+=,[pscustomobject]@{Path=$relative;Name=[IO.Path]::GetFileNameWithoutExtension($relative);SHA256=(Get-FileHash -LiteralPath $copy).Hash;Body=$body;Kind=$(if([IO.Path]::GetExtension($relative) -eq '.cls'){2}else{1})}
}
$memory=Get-CimInstance Win32_PerfFormattedData_PerfOS_Memory
$freeRAM=[double]$memory.AvailableMBytes*1MB
$freeCommit=[double]$memory.CommitLimit-[double]$memory.CommittedBytes
[ordered]@{ObservedUTC=[DateTimeOffset]::UtcNow.ToString('o');FreePhysicalGiB=$freeRAM/1GB;FreeCommitGiB=$freeCommit/1GB;Ready=($freeRAM -ge 4GB -and $freeCommit -ge 8GB)} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $directory 'MemoryGate.json') -Encoding UTF8
if($freeRAM -lt 4GB -or $freeCommit -lt 8GB){throw 'Insufficient combined reserve; no Excel started.'}
Add-Type 'using System;using System.Runtime.InteropServices;public static class NdmContourProbeOwner{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);}'
$prior=@(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)
$app=$null;$book=$null;$owned=0;$ticks=0
try{
    $app=New-Object -ComObject Excel.Application
    [uint32]$nativeID=0;[void][NdmContourProbeOwner]::GetWindowThreadProcessId([IntPtr]$app.Hwnd,[ref]$nativeID)
    $owned=[int]$nativeID;$ticks=(Get-Process -Id $owned).StartTime.ToUniversalTime().Ticks
    if($owned -in $prior){throw 'Preparation requires a new owned Excel.'}
    [ordered]@{PID=$owned;StartTicksUTC=[string]$ticks;SourceSHA256=$baselineHash;Status='PREPARING';Label=$Label} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $directory 'Ownership.json') -Encoding UTF8
    $app.Visible=$false;$app.DisplayAlerts=$false;$app.EnableEvents=$false;$app.AutomationSecurity=3
    $book=$app.Workbooks.Open($candidate,0,$false)
    foreach($sourceInput in $inputs){
        $component=$null
        try{$component=$book.VBProject.VBComponents.Item($sourceInput.Name)}catch{}
        if(-not $component){$component=$book.VBProject.VBComponents.Add([int]$sourceInput.Kind);$component.Name=$sourceInput.Name}
    }
    foreach($sourceInput in $inputs){
        $component=$book.VBProject.VBComponents.Item($sourceInput.Name)
        $module=$component.CodeModule
        if($module.CountOfLines -gt 0){$module.DeleteLines(1,$module.CountOfLines)}
        $module.AddFromString(($sourceInput.Body -split '\r?\n') -join "`r`n")
    }
    $config=$book.Worksheets.Item('Config');$settings=$book.Names.Item('rngSystemSettings').RefersToRange
    $keyRow=0;$stepRow=0
    for($i=1;$i -le $settings.Rows.Count;$i++){
        $key=[string]$settings.Cells.Item($i,1).Value2
        if($key -ceq 'Mesh.CellGeometry'){$keyRow=$settings.Row+$i-1}
        if($key -ceq 'Mesh.StepMode'){$stepRow=$settings.Row+$i-1}
    }
    if(-not $keyRow){
        if(-not $stepRow){throw 'Base mesh setting not found.'}
        $config.Range('A'+$stepRow+':E'+$stepRow).Insert(-4121)
        $config.Range('A'+($stepRow+1)+':E'+($stepRow+1)).Copy($config.Range('A'+$stepRow+':E'+$stepRow))
        $keyRow=$stepRow
        $config.Cells.Item($keyRow,1).Value2='Mesh.CellGeometry';$config.Cells.Item($keyRow,2).Value2='Adaptive';$config.Cells.Item($keyRow,3).Value2='-'
        $config.Cells.Item($keyRow,4).Value2='Adaptive: прежняя сетка. Contour: фактическая геометрия ячеек.'
        $config.Cells.Item($keyRow,2).Validation.Delete()
        $config.Cells.Item($keyRow,2).Validation.Add(3,1,1,('Adaptive'+[string]$app.International(5)+'Contour'))
    }
    $results=$book.Worksheets.Item('Results')
    $anchor=$null;try{$anchor=$book.Names.Item('rngNDMElementGeometry').RefersToRange}catch{}
    if(-not $anchor){[void]$book.Names.Add('rngNDMElementGeometry',"=Results!`$CR`$156")}
    $headers=@('RunID','LoopID','SegmentID','Sequence','LoopRole','SegmentType','StartX','StartY','EndX','EndY','CenterX','CenterY','Radius','SweepAngle, rad','SourceID','Comment','LengthUnit (SectionXY)','ElementID')
    for($i=0;$i -lt $headers.Count;$i++){$results.Cells.Item(156,96+$i).Value2=$headers[$i]}
    $results.Range('CR155:DI155').Merge();$results.Cells.Item(155,96).Value2='Фактические границы бетонных элементов'
    $results.Range('CR155:DI156').Font.Bold=$true
    $app.VBE.ActiveVBProject=$book.VBProject
    $compile=$app.VBE.CommandBars.FindControl(1,578)
    if(-not $compile){throw 'VBE compile command unavailable.'}
    if($compile.Enabled){$compile.Execute()}
    if($compile.Enabled){
        $pane=$app.VBE.ActiveCodePane;$sl=0;$sc=0;$el=0;$ec=0
        $pane.GetSelection([ref]$sl,[ref]$sc,[ref]$el,[ref]$ec)
        [ordered]@{Component=$pane.CodeModule.Parent.Name;Line=$sl;Text=$pane.CodeModule.Lines([math]::Max(1,$sl-2),5)} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $directory 'CompileFailure.json') -Encoding UTF8
        throw 'Native compile failed; diagnostic preserved.'
    }
    $components=$book.VBProject.VBComponents.Count
    if($components -ne 181){throw ('Unexpected component inventory: '+$components)}
    $book.Save();$book.Close($false);$book=$null
    $manifest=@($inputs|Select-Object Path,Name,SHA256)
    [ordered]@{Status='COMPILED_PRIVATE';Label=$Label;Workbook=$candidate;WorkbookSHA256=(Get-FileHash -LiteralPath $candidate).Hash;SourceSHA256=$baselineHash;Components=$components;Inputs=$manifest;FullRegressionRun=$false} | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $directory 'Preparation.json') -Encoding UTF8
}finally{
    $process=Get-Process -Id $owned -ErrorAction SilentlyContinue
    if($process -and $process.ProcessName -eq 'EXCEL' -and $process.StartTime.ToUniversalTime().Ticks -eq $ticks -and $owned -notin $prior){if($book){$book.Close($false)};if($app){$app.Quit()}}
    $component=$null;$module=$null;$compile=$null;$pane=$null;$anchor=$null;$results=$null;$settings=$null;$config=$null
    if($app){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
Get-Content -LiteralPath (Join-Path $directory 'Preparation.json') -Raw
