# скрипт обновляет отдельные части существующей книги без ручного импорта модулей через редактор VBA.

param(
    [string]$WorkbookPath = "workbook/output/RC_Section_NDM.xlsm"
)

$ErrorActionPreference = "Stop"

# Возвращает подготовленные данные или справочное значение для дальнейшего шага сборки.
function Get-VbaComponentName {
    param([string]$Path)
    $source = [System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::Default)
    foreach ($line in ($source -split "`r?`n")) {
        if ($line -match '^Attribute\s+VB_Name\s*=\s*"([^"]+)"') {
            return $Matches[1]
        }
    }
    throw "VBA source has no Attribute VB_Name: $Path"
}

# Импортирует исходные VBA-модули в книгу, сохраняя воспроизводимость сборки.
function Import-VbaSourceFile {
    param(
        [object]$Workbook,
        [string]$Path
    )

    $componentName = Get-VbaComponentName $Path
    $existing = $null
    foreach ($component in $Workbook.VBProject.VBComponents) {
        if ($component.Name -eq $componentName) {
            $existing = $component
            break
        }
    }
    if ($existing -ne $null) {
        $Workbook.VBProject.VBComponents.Remove($existing)
    }

    $extension = [System.IO.Path]::GetExtension($Path).ToLowerInvariant()
    $componentType = 1
    if ($extension -eq ".cls") { $componentType = 2 }
    $component = $Workbook.VBProject.VBComponents.Add($componentType)
    $component.Name = $componentName

    $source = [System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::Default)
    $lines = $source -split "`r?`n"
    $bodyLines = New-Object System.Collections.Generic.List[string]
    $insideHeader = $false
    foreach ($line in $lines) {
        if ($line -match '^VERSION\s+') { continue }
        if ($line -match '^BEGIN\s*$') { $insideHeader = $true; continue }
        if ($insideHeader) {
            if ($line -match '^END\s*$') { $insideHeader = $false }
            continue
        }
        if ($line -match '^Attribute\s+VB_') { continue }
        $bodyLines.Add($line)
    }
    $body = ($bodyLines -join "").Trim()
    if ($body.Length -gt 0) {
        $component.CodeModule.AddFromString($body)
    }
}

# Устанавливает значение, оформление или именованный диапазон в книге через Excel COM.
function Set-SystemSetting {
    param(
        [object]$Workbook,
        [string]$Key,
        [string]$Value,
        [string]$DefaultValue,
        [string]$Unit,
        [string]$Purpose,
        [string]$Source
    )

    $settings = $Workbook.Names.Item("rngSystemSettings").RefersToRange
    $row = $null
    for ($i = 2; $i -le $settings.Rows.Count; $i++) {
        if ([string]::Equals([string]$settings.Cells.Item($i, 1).Value2, $Key, [System.StringComparison]::OrdinalIgnoreCase)) {
            $row = $i
            break
        }
    }
    if ($row -eq $null) {
        $row = $settings.Rows.Count
        while ($row -gt 2 -and [string]::IsNullOrWhiteSpace([string]$settings.Cells.Item($row, 1).Value2)) {
            $row--
        }
        $row++
        if ($row -gt $settings.Rows.Count) {
            $sheet = $settings.Worksheet
            $insertAt = $settings.Row + $settings.Rows.Count
            $sheet.Rows.Item($insertAt).Insert() | Out-Null
            $newRange = $sheet.Range($sheet.Cells.Item($settings.Row, $settings.Column), $sheet.Cells.Item($settings.Row + $settings.Rows.Count, $settings.Column + $settings.Columns.Count - 1))
            $Workbook.Names.Item("rngSystemSettings").RefersTo = "=" + $newRange.Address($true, $true, 1, $true)
            $settings = $Workbook.Names.Item("rngSystemSettings").RefersToRange
            $row = $settings.Rows.Count
        }
    }

    $settings.Cells.Item($row, 1).Value2 = $Key
    $settings.Cells.Item($row, 2).Value2 = $Value
    $settings.Cells.Item($row, 3).Value2 = $DefaultValue
    $settings.Cells.Item($row, 4).Value2 = $Unit
    $settings.Cells.Item($row, 5).Value2 = $Purpose
    $settings.Cells.Item($row, 6).Value2 = $Source
    $settings.Cells.Item($row, 7).Value2 = "No"
}

# Удаляет только служебный объект, который может мешать повторяемой сборке или проверке.
function Remove-DuplicatePrintAreaName {
    param([string]$Path)

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [System.IO.Compression.ZipFile]::Open($Path, [System.IO.Compression.ZipArchiveMode]::Update)
    try {
        $entry = $zip.GetEntry("xl/workbook.xml")
        if ($entry -eq $null) { return }
        $reader = New-Object IO.StreamReader($entry.Open())
        $xml = $reader.ReadToEnd()
        $reader.Close()
        $newXml = [regex]::Replace($xml, '<definedName name="Print_Area"[^>]*>.*?</definedName>\s*', '')
        if ($newXml -ne $xml) {
            $stream = $entry.Open()
            $stream.SetLength(0)
            $writer = New-Object IO.StreamWriter($stream, (New-Object System.Text.UTF8Encoding($false)))
            $writer.Write($newXml)
            $writer.Close()
        }
    }
    finally {
        if ($zip -ne $null) { $zip.Dispose() }
    }
}

$root = Resolve-Path -LiteralPath (Join-Path $PSScriptRoot "../..")
$fullWorkbookPath = Join-Path $root $WorkbookPath
if (-not (Test-Path -LiteralPath $fullWorkbookPath)) { throw "Workbook not found: $fullWorkbookPath" }
Remove-DuplicatePrintAreaName $fullWorkbookPath

$excel = $null
$workbook = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false
    $excel.AutomationSecurity = 1

    $workbook = $excel.Workbooks.Open($fullWorkbookPath)

    $files = @(
        "src/Batch/CBatchSectionCalculator.cls",
        "src/Excel/CLoadCombinationReader.cls",
        "src/Excel/CBatchResultWriter.cls",
        "tests/modTestCapacitySolver.bas",
        "tests/modTestBatchCalculation.bas"
    )
    foreach ($relative in $files) {
        Import-VbaSourceFile $workbook (Join-Path $root $relative)
    }

    $calc = $workbook.Worksheets.Item(1)
    $workbook.Names.Item("rngLoadCombinations").RefersTo = "=" + $calc.Range("A37:G57").Address($true, $true, 1, $true)
    $calc.PageSetup.PrintArea = $calc.Range("A1:BT57").Address($true, $true)


    $workbook.Save()
    Write-Output "Workbook stage 8 refreshed: $fullWorkbookPath"
}
finally {
    if ($workbook -ne $null) {
        $workbook.Close($false)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
    }
    if ($excel -ne $null) {
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Remove-DuplicatePrintAreaName $fullWorkbookPath
