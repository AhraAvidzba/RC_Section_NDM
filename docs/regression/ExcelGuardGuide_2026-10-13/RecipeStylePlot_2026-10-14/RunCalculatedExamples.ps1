# Настоящие расчёты и PNG двух сечений. Изменяются только отдельные копии.
$ErrorActionPreference='Stop';$root=(Get-Location).Path;$directory=$PSScriptRoot
Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public static class ExampleWindow{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr hWnd,out uint id);}'
function Set-Value($cell,$value) {[void]$cell.GetType().InvokeMember('Value2',[Reflection.BindingFlags]::SetProperty,$null,$cell,[object[]](,$value))}
function Set-Key($book,[string]$key,$value) {
    $table=$book.Names.Item('rngSystemSettings').RefersToRange
    for($i=1;$i -le $table.Rows.Count;$i++){if([string]$table.Cells.Item($i,1).Value2 -ceq $key){Set-Value $table.Cells.Item($i,2) $value;return}}
    throw ('Missing setting: '+$key)
}
function Capture-Shapes($chartObject) {
    $list=@()
    foreach($shape in $chartObject.Chart.Shapes){
        $item=[ordered]@{Name=[string]$shape.Name;Type=[int]$shape.Type;Left=[double]$shape.Left;Top=[double]$shape.Top;Width=[double]$shape.Width;Height=[double]$shape.Height;Rotation=[double]$shape.Rotation}
        if($shape.Type -eq 5){$nodes=@();for($i=1;$i -le $shape.Nodes.Count;$i++){$p=$shape.Nodes.Item($i).Points;$nodes+=,@([double]$p.GetValue(1,1),[double]$p.GetValue(1,2))};$item.Nodes=$nodes}
        $list+=,$item
    }
    return ,$list
}
$protected=@('workbook/output/RC_Section_NDM.xlsm','workbook/output/VBA_All_Code.txt','workbook/output/RC_Section_NDM_execution_report.txt');$hashes=@{}
foreach($path in $protected){$hashes[$path]=(Get-FileHash -LiteralPath (Join-Path $root $path)).Hash}
$records=@()
foreach($example in @('BaselineG','GSection','HollowRounded')){
    $folder=Join-Path $directory ('Calculated'+$example)
    if(Test-Path -LiteralPath $folder){throw 'Existing example evidence must not be replaced.'}
    New-Item -ItemType Directory -Path $folder | Out-Null
    $source=Join-Path $directory $(if($example -eq 'BaselineG'){'Baseline.xlsm'}else{'Candidate.xlsm'})
    $path=Join-Path $folder ($example+'.xlsm');Copy-Item -LiteralPath $source -Destination $path
    $excel=$null;$book=$null;$ownedPid=0;$ownedTicks=0
    try{
        $excel=New-Object -ComObject Excel.Application
        $excel.Visible=$false;$excel.EnableEvents=$false;$excel.DisplayAlerts=$false;$excel.AutomationSecurity=1;$excel.ScreenUpdating=$true
        [uint32]$id=0;[void][ExampleWindow]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$id)
        $ownedPid=[int]$id;$ownedTicks=(Get-Process -Id $ownedPid).StartTime.ToUniversalTime().Ticks
        Write-Output ('OWNED_EXCEL: '+$ownedPid+'; example='+$example)
        $book=$excel.Workbooks.Open($path,0,$false);if($book.ReadOnly){throw 'Example is read-only.'}
        $excel.Calculation=-4135
        Set-Key $book 'Geometry.Source' 'Generated'
        Set-Key $book 'Geometry.Type' $(if($example -eq 'HollowRounded'){'HollowRectangle'}else{'RectSet'})
        Set-Key $book 'Load.ReferenceOffsetX' 0;Set-Key $book 'Load.ReferenceOffsetY' 0
        Set-Key $book 'Mesh.StepX' 40;Set-Key $book 'Mesh.StepY' 30;Set-Key $book 'Mesh.BoundarySubdivisions' 2
        Set-Key $book 'Plot.Enabled' 'Yes';Set-Key $book 'Plot.AutoUpdateAfterCalculation' 'Yes';Set-Key $book 'Plot.LoadCase' 'EXAMPLE'
        Set-Key $book 'Plot.PrincipalAxesMode' 'Transformed';Set-Key $book 'Plot.LoadApplicationPointEnabled' 'Yes'
        if($example -eq 'HollowRounded'){
            $table=$book.Names.Item('rngHollowRectangleGeometry').RefersToRange
            Set-Value $table.Cells.Item(3,2) 0;Set-Value $table.Cells.Item(4,2) 0
            $sizes=@(1200,1000,160,700,400,80);for($i=1;$i -le 6;$i++){Set-Value $table.Cells.Item(7,$i) $sizes[$i-1]}
        }else{
            $table=$book.Names.Item('rngRectSetGeometry').RefersToRange
            Set-Value $table.Cells.Item(3,2) 'LSection';Set-Value $table.Cells.Item(4,2) 0
        }
        $lc=$book.Names.Item('rngLoadCombinations').RefersToRange
        $lc.Offset(1,0).Resize($lc.Rows.Count-1,$lc.Columns.Count).ClearContents()
        Set-Value $lc.Cells.Item(2,1) 'EXAMPLE';Set-Value $lc.Cells.Item(2,2) 50;Set-Value $lc.Cells.Item(2,3) 5;Set-Value $lc.Cells.Item(2,4) 2
        Set-Value $lc.Cells.Item(2,5) 'PR1';Set-Value $lc.Cells.Item(2,6) 'Auto';Set-Value $lc.Cells.Item(2,7) 'Контроль схемы; точка (0; 0) от центра бетона'
        $book.Worksheets.Item('Config').Calculate()
        $watch=[Diagnostics.Stopwatch]::StartNew()
        $message=[string]$excel.Run("'$($book.Name)'!modWorkbookCalculation.RunSectionCalculationForWorkbook",$book,$false)
        $watch.Stop();$message | Set-Content -LiteralPath (Join-Path $folder 'CalculationMessage.txt') -Encoding UTF8
        if($message -notmatch 'Расчет завершен'){throw $message}
        Write-Output ('CALCULATED: '+$example+'; seconds='+$watch.Elapsed.TotalSeconds)
        $module=$book.VBProject.VBComponents.Item('modTestPlotConfig').CodeModule;$originalLines=$module.CountOfLines
        $module.AddFromString([IO.File]::ReadAllText((Join-Path $directory 'CalculatedExampleProbe.bas'),[Text.Encoding]::UTF8))
        $snapshot=[string]$excel.Run("'$($book.Name)'!modTestPlotConfig.GoalCalculatedExampleSnapshot")
        $snapshot | Set-Content -LiteralPath (Join-Path $folder 'CalculatedData.tsv') -Encoding UTF8
        $module.DeleteLines($originalLines+1,$module.CountOfLines-$originalLines)
        if($snapshot -match 'STATE\t(NumFail|FAIL|InputErr)' -or $snapshot -match 'STATE[^\r\n]*\tTrue\t'){throw 'No physically admissible calculated state for illustration.'}
        $sheet=$book.Worksheets.Item('Расчет');$chart=$sheet.ChartObjects().Item('chtNDMSectionPlot')
        $originLeft=[double]$chart.Left;$originTop=[double]$chart.Top
        foreach($mode in @('Transformed','Concrete')){
            if($example -ne 'GSection' -and $mode -eq 'Concrete'){continue}
            $chart.Left=$originLeft;$chart.Top=$originTop;$chart.Width=1000;$chart.Height=740
            Set-Key $book 'Plot.PrincipalAxesMode' $mode
            foreach($frame in @('Original','Moved','Resized')){
                if($example -ne 'GSection' -and $frame -ne 'Original'){continue}
                if($frame -eq 'Moved'){$chart.Left=[double]$chart.Left+215;$chart.Top=[double]$chart.Top+117}
                if($frame -eq 'Resized'){$chart.Width=1200;$chart.Height=850}
                $excel.Run("'$($book.Name)'!modWorkbookCalculation.UpdateSectionPlotForWorkbook",$book,$true)
                $before=Capture-Shapes $chart
                $sheet.Activate();$chart.Activate();$chart.Chart.Refresh()
                $png=Join-Path $folder ($mode+$frame+'.png')
                $ok=$chart.Chart.Export($png,'PNG')
                if(-not $ok -or (Get-Item -LiteralPath $png).Length -le 0){throw 'Empty native chart image.'}
                $after=Capture-Shapes $chart
                [ordered]@{Example=$example;Mode=$mode;Frame=$frame;ChartLeft=[double]$chart.Left;ChartTop=[double]$chart.Top;ChartWidth=[double]$chart.Width;ChartHeight=[double]$chart.Height;BeforeExport=$before;Shapes=$after;PNGBytes=(Get-Item -LiteralPath $png).Length} | ConvertTo-Json -Depth 7 | Set-Content -LiteralPath (Join-Path $folder ($mode+$frame+'.json')) -Encoding UTF8
            }
        }
        Set-Key $book 'Plot.PrincipalAxesMode' 'Transformed'
        $book.Save();$book.Close($false);$book=$null
        $records+=,[ordered]@{Example=$example;Passed=$true;Seconds=$watch.Elapsed.TotalSeconds;Loads=@(50,5,2);InputUnits=@('tf','tf*m','tf*m');LoadOffsets=@(0,0);CalculationProfile='PR1';SourceSHA256=(Get-FileHash -LiteralPath $source).Hash;WorkbookSHA256=(Get-FileHash -LiteralPath $path).Hash;TemporaryHelperRemoved=$true}
    }finally{
        if($book){$book.Close($false)}
        if($excel){$excel.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
        [GC]::Collect();[GC]::WaitForPendingFinalizers()
        if($ownedPid){$p=Get-Process -Id $ownedPid -ErrorAction SilentlyContinue;if($p -and $p.ProcessName -eq 'EXCEL' -and $p.StartTime.ToUniversalTime().Ticks -eq $ownedTicks){Stop-Process -Id $ownedPid -Force}}
        foreach($protectedPath in $protected){if((Get-FileHash -LiteralPath (Join-Path $root $protectedPath)).Hash -cne $hashes[$protectedPath]){throw 'Production changed during example preparation.'}}
    }
}
[ordered]@{Passed=$true;Examples=$records;ProductionUnchanged=$true} | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directory 'CalculatedExamples.json') -Encoding UTF8
