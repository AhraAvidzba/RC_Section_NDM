# Проверяет открытый Freeform в отдельной книге Excel; рабочая книга не открывается.
$ErrorActionPreference='Stop'
$directory=$PSScriptRoot
Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public static class OpenArcWindow{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr hWnd,out uint id);}'
$excel=$null;$book=$null;$ownedPid=0;$ownedTicks=0
try {
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false;$excel.EnableEvents=$false;$excel.DisplayAlerts=$false;$excel.AutomationSecurity=3
    [uint32]$id=0;[void][OpenArcWindow]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$id)
    $ownedPid=[int]$id;$ownedTicks=(Get-Process -Id $ownedPid).StartTime.ToUniversalTime().Ticks
    Write-Output ('OWNED_EXCEL: '+$ownedPid)
    $book=$excel.Workbooks.Add()
    $sheet=$book.Worksheets.Item(1)
    $chart=$sheet.ChartObjects().Add(0,0,400,400).Chart
    $chart.ChartType=-4169
    $builder=$chart.Shapes.BuildFreeform(1,330,200)
    for($i=1;$i -le 18;$i++) {
        $angle=[Math]::PI/2*$i/18
        $builder.AddNodes(0,1,200+130*[Math]::Cos($angle),200-130*[Math]::Sin($angle))
    }
    $shape=$builder.ConvertToShape()
    $shape.Name='OpenArc';$shape.Fill.Visible=0;$shape.Line.Visible=-1;$shape.Line.ForeColor.RGB=0;$shape.Line.Weight=1.7
    if(-not $chart.Export((Join-Path $directory 'OpenArc.png'),'PNG')){throw 'Export failed.'}
    $book.SaveAs((Join-Path $directory 'OpenArc.xlsx'),51)
    [ordered]@{Nodes=$shape.Nodes.Count;Shapes=$chart.Shapes.Count;FillVisible=$shape.Fill.Visible;ShapeType=$shape.Type} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $directory 'OpenArc.json') -Encoding UTF8
    $book.Close($false);$book=$null
} finally {
    if($book){$book.Close($false)}
    if($excel){$excel.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
    if($ownedPid){$p=Get-Process -Id $ownedPid -ErrorAction SilentlyContinue;if($p -and $p.ProcessName -eq 'EXCEL' -and $p.StartTime.ToUniversalTime().Ticks -eq $ownedTicks){Stop-Process -Id $ownedPid -Force}}
}
