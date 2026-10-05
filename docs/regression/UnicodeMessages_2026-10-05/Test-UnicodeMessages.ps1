param([string]$WorkbookPath, [string]$ReportPath)

$ErrorActionPreference = 'Stop'
$excel = $null
$book = $null
$lines = New-Object System.Collections.Generic.List[string]
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = 1
    $book = $excel.Workbooks.Open($WorkbookPath, 0, $true)
    $lambda = [char]0x3BB
    $message = "Для пути ${lambda}*Mx в колонке LoadPath соответствующая нагрузка должна быть ненулевой."
    $response = $excel.Run("'$($book.Name)'!modWorkbookMessages.ShowWorkbookMessage", $message, 48, 'NDM Unicode warning test')
    if ([int]$response -ne 1) { throw "Unexpected warning response: $response" }
    $lines.Add('PASS: warning dialog, OK button response = 1')
    $symbols = -join ([char[]](0x3BB, 0x3C8, 0x3C3, 0x3B5, 0x3BA, 0x2264, 0x2265, 0x2205, 0x2248, 0xB2, 0xB3))
    $message = "Проверка символов: $symbols`r`n${lambda}*Mx; ${lambda}*My; ${lambda}*Mxy; ${lambda}*N; ${lambda}*NMxy`r`nКириллица и путь: C:\Расчет\RC_Section_NDM.xlsm"
    $response = $excel.Run("'$($book.Name)'!modWorkbookMessages.ShowWorkbookMessage", $message, 64, 'NDM Unicode information test')
    if ([int]$response -ne 1) { throw "Unexpected information response: $response" }
    $lines.Add('PASS: information dialog, OK button response = 1')
    $lines.Add('PASS: no calculations or worksheets were changed; workbook opened read-only')
    $lines.Add('Components: ' + $book.VBProject.VBComponents.Count)
    $lines.Add('Office: ' + $excel.Version)
}
catch {
    $lines.Add('FAIL: ' + $_.Exception.Message)
    throw
}
finally {
    [IO.File]::WriteAllLines($ReportPath, $lines, (New-Object Text.UTF8Encoding($true)))
    if ($book -ne $null) {
        try { $book.Close($false) } catch { Write-Warning $_.Exception.Message }
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($book)
    }
    if ($excel -ne $null) {
        try { $excel.Quit() } catch { Write-Warning $_.Exception.Message }
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
$lines
