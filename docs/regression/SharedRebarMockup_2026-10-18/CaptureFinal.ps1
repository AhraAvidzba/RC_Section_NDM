# Повторный снимок широкого диапазона целиком, без сохранения книги.
$ErrorActionPreference='Stop'
$manifest=Get-Content -LiteralPath (Join-Path $PSScriptRoot 'Manifest.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$path=[string]$manifest.file;$before=(Get-FileHash -LiteralPath $path).Hash
$excel=$null;$book=$null
try{
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false;$excel.AutomationSecurity=3
    $book=$excel.Workbooks.Open($path,0,$true);$sheet=$book.Worksheets.Item('Config')
    foreach($item in $manifest.captureRanges){
        if($item.type){$sheet.Range('B7').Value2=[string]$item.type;$sheet.Calculate()}
        $sheet.Activate();$excel.ActiveWindow.Zoom=50;$excel.ActiveWindow.ScrollColumn=8;$excel.ActiveWindow.ScrollRow=1
        $source=$sheet.Range([string]$item.range);$source.CopyPicture(1,-4147) | Out-Null
        $scratch=$book.Worksheets.Add()
        try{
            $chart=$scratch.ChartObjects().Add(0,0,[double]$source.Width,[double]$source.Height)
            $chart.Activate();$chart.Chart.Paste()
            if(-not $chart.Chart.Export((Join-Path $PSScriptRoot ($item.name+'_Native.png')),'PNG')){throw 'Native capture failed.'}
        }finally{$scratch.Delete();$sheet.Activate()}
        if($item.type){$sheet.Range('B7').Value2='LSection';$sheet.Calculate()}
    }
}finally{if($book){$book.Close($false)};if($excel){$excel.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}}
if((Get-FileHash -LiteralPath $path).Hash -cne $before){throw 'Read-only capture changed the workbook.'}
Write-Output 'FULL_RANGE_CAPTURE: both ranges; source unchanged.'
