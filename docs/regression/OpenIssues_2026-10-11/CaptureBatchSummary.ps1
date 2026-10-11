# Native Excel cell rendering; never saves the source workbook.
param([string]$OutputFile='docs/regression/OpenIssues_2026-10-11/BatchSummary.png')
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../../..')).Path
$source=Join-Path $root 'workbook/output/RC_Section_NDM.xlsm'
$before=(Get-FileHash -LiteralPath $source).Hash
$output=[IO.Path]::GetFullPath((Join-Path $root $OutputFile))
if (-not $output.StartsWith($root+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)) {throw 'Output must stay in the workspace.'}
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public static class SummaryCaptureWindow { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId); }'
$excel=$null;$book=$null;$picture=$null;$ownedPid=0;$ownedTicks=0
try {
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false;$excel.AutomationSecurity=3
    [uint32]$id=0;[void][SummaryCaptureWindow]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$id)
    $ownedPid=[int]$id;$ownedTicks=(Get-Process -Id $ownedPid).StartTime.ToUniversalTime().Ticks
    Write-Output ('OWNED_EXCEL: id='+$ownedPid+'; utcStartTicks='+$ownedTicks)
    $book=$excel.Workbooks.Open($source,0,$true)
    $sheet=$book.Worksheets.Item('Results');$sheet.Activate()
    $anchor=$book.Names.Item('rngBatchSummary').RefersToRange
    $area=$anchor.Resize(4,13)
    for ($attempt=0;$attempt -lt 3 -and $null -eq $picture;$attempt++) {
        $area.CopyPicture(1,2)
        Start-Sleep -Milliseconds 300
        $picture=[Windows.Forms.Clipboard]::GetImage()
    }
    if ($null -eq $picture) {throw 'Excel did not provide a rendered image.'}
    $picture.Save($output,[Drawing.Imaging.ImageFormat]::Png)
    Write-Output ('IMAGE: '+$output+'; size='+$picture.Width+'x'+$picture.Height)
    foreach ($cell in @('A1','A2','A3','A4','G1','G2','G3')) {Write-Output ($cell+'='+$sheet.Range($cell).Text)}
} finally {
    if ($picture) {$picture.Dispose()}
    if ($book) {$book.Close($false)}
    if ($excel) {$excel.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
    if ($ownedPid) {
        $owned=Get-Process -Id $ownedPid -ErrorAction SilentlyContinue
        if ($owned -and $owned.ProcessName -eq 'EXCEL' -and $owned.StartTime.ToUniversalTime().Ticks -eq $ownedTicks) {Stop-Process -Id $ownedPid -Force}
    }
    if ((Get-FileHash -LiteralPath $source).Hash -cne $before) {throw 'Source workbook changed during capture.'}
}
