# Адресные тесты разработки СП 35 в отдельной копии книги; исходные данные пользователя сохраняются.
param(
    [string]$WorkbookPath = 'workbook/output/RC_Section_NDM.xlsm',
    [string]$ReportDirectory = 'docs/regression/SP35/Baseline',
    [string[]]$Macros = @('modTestCrackWidth.RunCrackWidthTests', 'modTestGeometry.RunGeometryTests'),
    [switch]$ImportSources,
    [string]$PreparationExportPath = ''
)

$ErrorActionPreference = 'Stop'
$Macros = @($Macros | ForEach-Object { $_ -split ',' })
$root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../../..')).Path
. (Join-Path $root 'tools/build_workbook/SettingsCatalog.ps1')
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
        # Только тестовая копия старого шаблона: освобождаем место под 25 колонок
        # геометрии. Основная книга и пользовательские именованные якоря не меняются.
        $geometry = $book.Names.Item('rngNDMSectionGeometry').RefersToRange
        $properties = $book.Names.Item('rngNDMSectionProperties').RefersToRange
        if ($properties.Column - $geometry.Column -lt 27) {
            foreach ($name in 'rngNDMSectionProperties', 'rngNDMMaterialDiagrams', 'rngNDMSectionAnnotations') {
                $anchor = $book.Names.Item($name).RefersToRange
                $book.Names.Item($name).RefersTo = "='$($anchor.Worksheet.Name)'!$($anchor.Offset(0, 10).Address())"
            }
        }
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
    if ($PreparationExportPath) {
        # Только воспроизведение: в собственной копии заменяем подготовку
        # прежним кодом, выгруженным из книги до исправления. Тесты остаются новыми.
        $export = Get-Content -LiteralPath (Join-Path $root $PreparationExportPath) -Raw -Encoding UTF8
        $match = [regex]::Match($export, '(?ms)^COMPONENT: CSP35CrackData\r?\nTYPE: \d+\r?\nLINES: \d+\r?\n=+\r?\n(.*?)(?=^=+\r?\nCOMPONENT:|\z)')
        if (-not $match.Success) { throw 'Baseline CSP35CrackData export not found.' }
        $component = $book.VBProject.VBComponents.Item('CSP35CrackData')
        $component.CodeModule.DeleteLines(1, $component.CodeModule.CountOfLines)
        $component.CodeModule.AddFromString($match.Groups[1].Value.Trim())
        $book.Save()
    }
    foreach ($macro in $Macros) {
        $watch = [Diagnostics.Stopwatch]::StartNew()
        Write-Output "RUN $macro"
        try { $result = [string]$excel.Run("'$($book.Name)'!$macro") }
        catch {
            Write-Output "Excel failed: $($_.Exception.Message)"
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
