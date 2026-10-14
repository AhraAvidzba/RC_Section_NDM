# Real standard recipes, same input/settings, native PNG before/after.
# Chart activation is injected only into the private test module in memory.
$ErrorActionPreference='Stop';$directory=$PSScriptRoot
Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public static class LegacyPreviewWindow{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr hWnd,out uint id);}'
foreach($stage in @('Baseline','Candidate')){
    $source=Join-Path $directory ($stage+'.xlsm');$hash=(Get-FileHash -LiteralPath $source).Hash
    $folder=Join-Path $directory ($stage+'LegacyPreviews');if(Test-Path -LiteralPath $folder){throw 'Existing previews must not be replaced.'}
    New-Item -ItemType Directory -Path $folder | Out-Null
    $path=Join-Path $folder 'LegacyPreviewProbe.xlsm';Copy-Item -LiteralPath $source -Destination $path
    $excel=$null;$book=$null;$pidOwned=0;$ticks=0
    try{
        $excel=New-Object -ComObject Excel.Application
        $excel.Visible=$false;$excel.EnableEvents=$false;$excel.DisplayAlerts=$false;$excel.AutomationSecurity=1;$excel.ScreenUpdating=$true
        [uint32]$id=0;[void][LegacyPreviewWindow]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$id)
        $pidOwned=[int]$id;$ticks=(Get-Process -Id $pidOwned).StartTime.ToUniversalTime().Ticks
        Write-Output ('OWNED_EXCEL: '+$pidOwned+'; stage='+$stage)
        $book=$excel.Workbooks.Open($path,0,$false);$excel.Calculation=-4135
        $module=$book.VBProject.VBComponents.Item('modTestShapeArchitecture').CodeModule
        $body=[string]$module.Lines(1,$module.CountOfLines)
        $old='    chart.Refresh'+"`r`n"+'    If Not chart.Export(outputDirectory'
        $new='    ThisWorkbook.Worksheets("Расчет").Activate'+"`r`n"+'    chart.Parent.Activate'+"`r`n"+$old
        if(-not $body.Contains($old)){throw 'Expected native exporter body not found.'}
        $module.DeleteLines(1,$module.CountOfLines);$module.AddFromString($body.Replace($old,$new))
        $control=$excel.VBE.CommandBars.FindControl(1,578);if($control.Enabled){$control.Execute()};if($control.Enabled){throw 'Private exporter did not compile.'}
        $result=[string]$excel.Run("'$($book.Name)'!modTestShapeArchitecture.DumpShapePreviewSnapshot",$folder)
        $result | Set-Content -LiteralPath (Join-Path $folder 'Shapes.tsv') -Encoding UTF8
        if($result -match '(?m)^FAIL:'){throw $result}
        $images=@(Get-ChildItem -LiteralPath $folder -Filter '*.png')
        if($images.Count -ne 9 -or @($images | Where-Object Length -eq 0).Count){throw 'Expected nine nonempty native images.'}
        [ordered]@{Passed=$true;Images=9;SourceSHA256=$hash;PrivateActivationSaved=$false} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $folder 'Run.json') -Encoding UTF8
        $book.Close($false);$book=$null
        Write-Output ('LEGACY_PREVIEWS: '+$stage+'; images=9')
    }finally{
        if($book){$book.Close($false)}
        if($excel){$excel.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
        [GC]::Collect();[GC]::WaitForPendingFinalizers()
        if($pidOwned){$p=Get-Process -Id $pidOwned -ErrorAction SilentlyContinue;if($p -and $p.ProcessName -eq 'EXCEL' -and $p.StartTime.ToUniversalTime().Ticks -eq $ticks){Stop-Process -Id $pidOwned -Force}}
        if((Get-FileHash -LiteralPath $source).Hash -cne $hash){throw 'Evidence source changed.'}
    }
}
