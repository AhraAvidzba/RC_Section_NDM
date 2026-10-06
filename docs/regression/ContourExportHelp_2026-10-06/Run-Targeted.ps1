# Запускает только выбранные VBA-проверки в собственной копии Excel.
# При повторной локальной правке импортирует только указанные модули, без
# пересборки справки и без тяжелых общих матриц нагрузок.
param(
    [string]$WorkbookPath = 'workbook/output/RC_Section_NDM.xlsm',
    [string]$ReportName = 'Targeted',
    [string]$ReportDirectory = '',
    [string[]]$ImportModules = @(),
    [string[]]$Macros = @('modTestGeometryQuery.RunGeometryQueryTests')
)
$ErrorActionPreference = 'Stop'
$root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../../..')).Path
. (Join-Path $root 'tools/build_workbook/SettingsCatalog.ps1')
$directory = Join-Path $PSScriptRoot $ReportName
if ($ReportDirectory) { $directory = Join-Path $root $ReportDirectory }
New-Item -ItemType Directory -Path $directory -Force | Out-Null
$copy = Join-Path $directory 'RC_Section_NDM.xlsm'
Copy-Item -LiteralPath (Join-Path $root $WorkbookPath) -Destination $copy -Force
$printAreas = @(Get-WorkbookPrintAreas $copy)
$excel = $null; $book = $null
$reportPath = Join-Path $root 'workbook/output/RC_Section_NDM_execution_report.txt'
$reportBytes = [IO.File]::ReadAllBytes($reportPath)
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.EnableEvents = $false; $excel.AutomationSecurity = 1
    $book = $excel.Workbooks.Open($copy)
    foreach ($relative in ($ImportModules | ForEach-Object { $_ -split ',' })) {
        $file = Get-Item -LiteralPath (Join-Path $root $relative)
        $component = $null; try { $component = $book.VBProject.VBComponents.Item($file.BaseName) } catch { }
        if ($null -ne $component) { $book.VBProject.VBComponents.Remove($component) }
        $body = [regex]::Match([IO.File]::ReadAllText($file.FullName, [Text.Encoding]::UTF8), '(?ms)^Option Explicit.*').Value
        if ([string]::IsNullOrWhiteSpace($body)) { throw "No module body: $file" }
        $type = 1; if ($file.Extension -eq '.cls') { $type = 2 }
        $component = $book.VBProject.VBComponents.Add($type); $component.Name = $file.BaseName
        $component.CodeModule.AddFromString(($body -split "`r?`n") -join "`r`n")
    }
    $book.Save()
    foreach ($macro in ($Macros | ForEach-Object { $_ -split ',' })) {
        Write-Output "RUN $macro"
        $text = [string]$excel.Run("'$($book.Name)'!$macro")
        $report = (($text -split "`r?`n" | ForEach-Object { $_.TrimEnd() }) -join "`r`n").TrimEnd()
        $report | Set-Content -LiteralPath (Join-Path $directory "$macro.txt") -Encoding UTF8
        Write-Output (($text -split "`r?`n" | Where-Object { $_ -match '^TOTAL|^FAIL:' }) -join "`n")
        if ($text -match '(?m)^FAIL[: ]|Failed: [1-9]|failed=[1-9]|RUNTIME ERROR:' -or $text -notmatch '(?m)^TOTAL[^\r\n]*failed=0\b') { throw "VBA test failed: $macro" }
    }
} finally {
    if ($null -ne $book) { $book.Close($false) }
    if ($null -ne $excel) { $excel.Quit() }
    Restore-WorkbookPrintAreas $copy $printAreas
    [IO.File]::WriteAllBytes($reportPath, $reportBytes)
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
