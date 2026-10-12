# ДЛЯ ТЕСТОВ: минимальная native-проба SAFEARRAY после скалярных записей.
$ErrorActionPreference='Stop'
Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public static class GuideBulkWindow{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr hWnd,out uint id);}'
$excel=$null;$book=$null;$ownedPid=0;$ownedTicks=0;$records=@()
try {
    $excel=New-Object -ComObject Excel.Application;$excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false
    [uint32]$id=0;[void][GuideBulkWindow]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$id)
    $ownedPid=[int]$id;$ownedTicks=(Get-Process -Id $ownedPid).StartTime.ToUniversalTime().Ticks
    $book=$excel.Workbooks.Add();$sheet=$book.Worksheets.Item(1)
    $sheet.Cells.Item(1,1).Value2='Scalar';$sheet.Cells.Item(2,1).Value2='Scalar2'
    $range=$sheet.Range('A4:F9');$range.Merge($true);$range.NumberFormat='@'
    $sheet.Range('A6:F6').UnMerge();$sheet.Range('C6:F6').Merge()
    $values=New-Object 'object[,]' 6,6
    for ($row=0;$row -lt 6;$row++) {$values[$row,0]='Row'+$row}
    $values[2,1]='Type';$values[2,2]='Meaning'
    foreach ($mode in @('Direct','Reflection','SingleColumn')) {
        try {
            if ($mode -eq 'Direct') {$range.Value2=$values}
            elseif ($mode -eq 'Reflection') {[void]$range.GetType().InvokeMember('Value2',[Reflection.BindingFlags]::SetProperty,$null,$range,[object[]](,$values))}
            else {
                $first=New-Object 'object[,]' 6,1
                for ($row=0;$row -lt 6;$row++) {$first[$row,0]='Row'+$row}
                $sheet.Range('A4:A9').Value2=$first
            }
            $records+=,[ordered]@{Mode=$mode;Passed=$true;A4=[string]$sheet.Cells.Item(4,1).Value2}
        } catch {$records+=,[ordered]@{Mode=$mode;Passed=$false;Reason=$_.Exception.Message;Stack=$_.ScriptStackTrace}}
    }
    $records | ConvertTo-Json | Set-Content -LiteralPath docs/regression/GeometryDeclarative_2026-10-12/FreshBuild/BulkWriteProbe.json -Encoding UTF8
    $records | ConvertTo-Json -Compress | Write-Output
} finally {
    if ($book) {$book.Close($false)}
    if ($excel) {$excel.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
    if ($ownedPid) {
        $owned=Get-Process -Id $ownedPid -ErrorAction SilentlyContinue
        if ($owned -and $owned.ProcessName -eq 'EXCEL' -and $owned.StartTime.ToUniversalTime().Ticks -eq $ownedTicks) {Stop-Process -Id $ownedPid -Force}
    }
}
