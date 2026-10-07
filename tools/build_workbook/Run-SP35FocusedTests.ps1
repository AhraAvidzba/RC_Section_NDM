# Адресные тесты разработки СП 35 в отдельной копии книги; исходные данные пользователя сохраняются.
param(
    [string]$WorkbookPath = 'workbook/output/RC_Section_NDM.xlsm',
    [string]$ReportDirectory = 'docs/regression/SP35/Baseline',
    [string[]]$Macros = @('modTestCrackWidth.RunCrackWidthTests', 'modTestGeometry.RunGeometryTests'),
    [switch]$ImportSources
)

$ErrorActionPreference = 'Stop'
$Macros = @($Macros | ForEach-Object { $_ -split ',' })
$root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
. (Join-Path $PSScriptRoot 'SettingsCatalog.ps1')
$source = Join-Path $root $WorkbookPath
$directory = Join-Path $root $ReportDirectory
New-Item -ItemType Directory -Path $directory -Force | Out-Null
$copy = Join-Path $directory 'RC_Section_NDM.xlsm'
if ([IO.Path]::GetFullPath($source) -eq [IO.Path]::GetFullPath($copy)) { throw 'Test copy must differ from source workbook.' }
Copy-Item -LiteralPath $source -Destination $copy -Force
$printAreas = @(Get-WorkbookPrintAreas $copy)
$excel = $null
$book = $null
$report = Join-Path $root 'workbook/output/RC_Section_NDM_execution_report.txt'
$savedReport = $null
if (Test-Path -LiteralPath $report) { $savedReport = [IO.File]::ReadAllBytes($report) }
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1
    $book = $excel.Workbooks.Open($copy)
    if ($ImportSources) {
        foreach ($file in Get-ChildItem -LiteralPath (Join-Path $root 'src'), (Join-Path $root 'tests') -Recurse -File | Where-Object Extension -in '.bas', '.cls') {
            $name = $file.BaseName
            $component = $null
            try { $component = $book.VBProject.VBComponents.Item($name) } catch { }
            if ($null -ne $component) { $book.VBProject.VBComponents.Remove($component) }
            $sourceText = [IO.File]::ReadAllText($file.FullName, [Text.Encoding]::UTF8)
            $body = [regex]::Match($sourceText, '(?ms)^Option Explicit.*').Value
            if ([string]::IsNullOrWhiteSpace($body)) { throw "Missing Option Explicit: $($file.FullName)" }
            $type = 1
            if ($file.Extension -eq '.cls') { $type = 2 }
            $component = $book.VBProject.VBComponents.Add($type)
            $component.Name = $name
            $component.CodeModule.AddFromString(($body -split "`r?`n") -join "`r`n")
        }
        # Сохраняем только собственную копию до тестовых мутаций, чтобы ее
        # VBA-код на диске соответствовал действительно проверяемому исходнику.
        $book.Save()
    }
    foreach ($macro in $Macros) {
        $watch = [Diagnostics.Stopwatch]::StartNew()
        Write-Output "RUN $macro"
        try { $result = [string]$excel.Run("'$($book.Name)'!$macro") }
        catch {
            Write-Output "Excel failed: $($_.Exception.Message)"
            $excel.Visible = $true
            Start-Sleep -Seconds 40
            throw
        }
        $watch.Stop()
        $result | Set-Content -LiteralPath (Join-Path $directory "$macro.txt") -Encoding UTF8
        Write-Output (($result -split "`r?`n" | Where-Object { $_ -match '^TOTAL|^FAIL:' }) -join "`n")
        Write-Output "ELAPSED $($watch.Elapsed.TotalSeconds) seconds"
        if ($result -match '(?m)^FAIL[: ]|Failed: [1-9]|failed=[1-9]|RUNTIME ERROR:') { throw "Failed assertions or VBA runtime error in $macro" }
        if ($result -notmatch '(?m)^TOTAL[^\r\n]*failed=0\b') { throw "Missing successful test summary in $macro" }
    }
}
finally {
    if ($null -ne $book) {
        try { $book.Close($false) } catch { Write-Warning $_.Exception.Message }
        [Runtime.InteropServices.Marshal]::ReleaseComObject($book) | Out-Null
    }
    if ($null -ne $excel) {
        try { $excel.Quit() } catch { Write-Warning $_.Exception.Message }
        [Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    # Восстанавливаем каноническое имя области печати только собственной копии:
    # локализованный Excel переименовывает его при сохранении импортированного VBA.
    if ($ImportSources) { Restore-WorkbookPrintAreas $copy $printAreas }
    if ($null -ne $savedReport) { [IO.File]::WriteAllBytes($report, $savedReport) }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
