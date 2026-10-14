# Пример проходит через обычные builder / Results / reader / plotter; основная книга не открывается.
$ErrorActionPreference='Stop';$directory=$PSScriptRoot
$source=Join-Path $directory 'Candidate.xlsm';$hash=(Get-FileHash -LiteralPath $source).Hash
$folder=Join-Path $directory 'TaperedFacet02'
$path=Join-Path $folder 'FacetProbe.xlsm'
if(Test-Path -LiteralPath $folder){
    if(@(Get-ChildItem -LiteralPath $folder).Count -ne 1 -or (Get-FileHash -LiteralPath $path).Hash -cne $hash){throw 'Existing evidence must not be replaced.'}
}else{New-Item -ItemType Directory -Path $folder | Out-Null;Copy-Item -LiteralPath $source -Destination $path}
Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public static class FacetWindow{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr hWnd,out uint id);}'
$excel=$null;$book=$null;$pidOwned=0;$ticks=0
try{
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false;$excel.EnableEvents=$false;$excel.DisplayAlerts=$false;$excel.AutomationSecurity=1;$excel.ScreenUpdating=$true
    [uint32]$id=0;[void][FacetWindow]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$id)
    $pidOwned=[int]$id;$ticks=(Get-Process -Id $pidOwned).StartTime.ToUniversalTime().Ticks
    Write-Output ('OWNED_EXCEL: '+$pidOwned)
    $book=$excel.Workbooks.Open($path,0,$false);$excel.Calculation=-4135
    $module=$book.VBProject.VBComponents.Item('modTestShapeArchitecture').CodeModule
    $body=[string]$module.Lines(1,$module.CountOfLines)
    $old='    chart.Refresh'+"`r`n"+'    If Not chart.Export(outputDirectory'
    $new='    ThisWorkbook.Worksheets("Расчет").Activate'+"`r`n"+'    chart.Parent.Activate'+"`r`n"+'    chart.Shapes("NDMPlot_Title").TextFrame2.TextRange.Text = "Наклонная грань: подпись арматуры и размер Aligned"'+"`r`n"+$old
    if(-not $body.Contains($old)){throw 'Expected exporter body not found.'}
    $module.DeleteLines(1,$module.CountOfLines);$module.AddFromString($body.Replace($old,$new))
    $module.AddFromString([IO.File]::ReadAllText((Join-Path $directory 'TaperedFacetPreview.bas'),[Text.Encoding]::UTF8))
    $recipe=$book.VBProject.VBComponents.Item('CSectionRecipeRoundedRectangle').CodeModule
    $recipeBody=[string]$recipe.Lines(1,$recipe.CountOfLines)
    $finish='    Set ISectionRecipe_CreateDefinition = facade.Finish'
    if(-not $recipeBody.Contains($finish)){throw 'Expected recipe finish not found.'}
    $recipe.DeleteLines(1,$recipe.CountOfLines)
    $recipe.AddFromString($recipeBody.Replace($finish,'    GoalAppendSlopeAnnotations facade, sizes'+"`r`n"+$finish+"`r`n"+'    Set modTestShapeArchitecture.GoalFacetDefinition = ISectionRecipe_CreateDefinition'))
    $recipe.AddFromString([IO.File]::ReadAllText((Join-Path $directory 'TaperedFacetProbe.bas'),[Text.Encoding]::UTF8))
    $control=$excel.VBE.CommandBars.FindControl(1,578);if($control.Enabled){$control.Execute()};if($control.Enabled){throw 'Probe did not compile.'}
    $result=[string]$excel.Run("'$($book.Name)'!modTestShapeArchitecture.GoalTaperedFacetPreview",$folder)
    $result | Set-Content -LiteralPath (Join-Path $folder 'GeometryAndShapes.tsv') -Encoding UTF8
    if($result -match '(?m)^FAIL:'){throw $result}
    $chart=$book.Worksheets.Item('Расчет').ChartObjects('chtNDMSectionPlot').Chart
    $metrics=@();foreach($shape in $chart.Shapes){
        if([int]$shape.Type -eq 17 -and $shape.TextFrame2.HasText){
            $metrics += [ordered]@{Name=[string]$shape.Name;Text=[string]$shape.TextFrame2.TextRange.Text;Left=[double]$shape.Left;Top=[double]$shape.Top;Width=[double]$shape.Width;Height=[double]$shape.Height;Rotation=[double]$shape.Rotation}
        }
    }
    $metrics | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $folder 'NativeTextMetrics.json') -Encoding UTF8
    if((Get-Item -LiteralPath (Join-Path $folder 'RoundedTapered.png')).Length -eq 0){throw 'Native image is empty.'}
    $result2=[string]$excel.Run("'$($book.Name)'!modTestShapeArchitecture.GoalTaperedFacetPreview",$folder)
    if($result2 -cne $result){throw 'Repeated native drawing changed.'}
    [ordered]@{Passed=$true;NativeExcelExport=$true;RepeatIdentical=$true;SourceSHA256=$hash;TemporaryCodeSaved=$false} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $folder 'Run.json') -Encoding UTF8
    $book.Close($false);$book=$null
    Write-Output 'TAPERED_FACET_PASS'
}finally{
    if($book){$book.Close($false)}
    if($excel){$excel.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
    if($pidOwned){$p=Get-Process -Id $pidOwned -ErrorAction SilentlyContinue;if($p -and $p.ProcessName -eq 'EXCEL' -and $p.StartTime.ToUniversalTime().Ticks -eq $ticks){Stop-Process -Id $pidOwned -Force}}
    if((Get-FileHash -LiteralPath $source).Hash -cne $hash){throw 'Source changed.'}
}
