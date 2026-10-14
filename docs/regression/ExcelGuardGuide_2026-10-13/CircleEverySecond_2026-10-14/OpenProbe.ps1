# ДЛЯ ТЕСТОВ: открывает только диагностические копии и закрывает собственный Excel.
$ErrorActionPreference='Stop'
$taskDirectory=Join-Path (Get-Location).Path 'docs/regression/ExcelGuardGuide_2026-10-13/CircleEverySecond_2026-10-14'
Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public static class BindingOpenIdentity{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr hWnd,out uint id);}'
$excel=$null;$book=$null;$ownedId=0;$ownedTicks=0;$rows=@()
try {
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false;$excel.AutomationSecurity=3
    [uint32]$id=0;[void][BindingOpenIdentity]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$id)
    $ownedId=[int]$id;$ownedTicks=(Get-Process -Id $ownedId).StartTime.ToUniversalTime().Ticks
    foreach($name in @('NativeCodeOnly','Styles','Workbook','Config','Help','Chain')) {
        $path=Join-Path $taskDirectory ($name+'.xlsm')
        try {
            $book=$excel.Workbooks.Open($path,0,$true)
            $rows+=@([ordered]@{Variant=$name;Opened=$true})
            $book.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book);$book=$null
        } catch {$rows+=@([ordered]@{Variant=$name;Opened=$false;Error=$_.Exception.Message})}
    }
    $rows | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $taskDirectory 'OpenProbe.json') -Encoding UTF8
    $rows | ConvertTo-Json -Compress
} finally {
    if($book){$book.Close($false)}
    if($excel){$excel.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
    if($ownedId){$owned=Get-Process -Id $ownedId -ErrorAction SilentlyContinue;if($owned -and $owned.ProcessName -eq 'EXCEL' -and $owned.StartTime.ToUniversalTime().Ticks -eq $ownedTicks){Stop-Process -Id $ownedId -Force}}
}
