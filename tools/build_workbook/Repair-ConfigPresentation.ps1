# Готовит исправленную копию существующей книги без пересборки пользовательского
# ввода. Источники validation берутся из контрольной сборки, VBA - из src/tests;
# текущие Config, именованные якоря, Results и ширины столбцов сохраняются.
param(
    [Parameter(Mandatory=$true)][string]$ReferenceWorkbook,
    [string]$CurrentWorkbook = 'workbook/output/RC_Section_NDM.xlsm',
    [string]$ReportDirectory = 'docs/regression/ConfigRepair_2026-10-06/Publication',
    [switch]$MigrateAutoCADCommonSettings
)
$ErrorActionPreference = 'Stop'
$root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
. (Join-Path $PSScriptRoot 'SettingsCatalog.ps1')
$current = (Resolve-Path -LiteralPath (Join-Path $root $CurrentWorkbook)).Path
$reference = (Resolve-Path -LiteralPath (Join-Path $root $ReferenceWorkbook)).Path
$directory = [IO.Path]::GetFullPath((Join-Path $root $ReportDirectory))
New-Item -ItemType Directory -Path $directory -Force | Out-Null
$target = Join-Path $directory 'RC_Section_NDM.xlsm'
$backup = Join-Path $directory 'BeforeRepair.xlsm'
$referenceCopy = Join-Path $directory 'ReferenceWorkbook.xlsm'
if ($target -eq $current -or $target -eq $reference) { throw 'Repair copy must differ from both inputs.' }
$sourceHash = (Get-FileHash -LiteralPath $current).Hash
Copy-Item -LiteralPath $current -Destination $backup -Force
Copy-Item -LiteralPath $current -Destination $target -Force
if ($referenceCopy -ne $reference) { Copy-Item -LiteralPath $reference -Destination $referenceCopy -Force }
$printAreas = @(Get-WorkbookPrintAreas $target)
$excel = $null; $book = $null; $referenceBook = $null
$snapshots = @{}
$records = New-Object 'System.Collections.Generic.List[object]'
$migrations = New-Object 'System.Collections.Generic.List[object]'

# Переносит только измененный реестр общих настроек по машинным ключам.
# Старые отдельные слои объединяются явно; при различии выбирается значение
# активного источника геометрии, оба прежних значения остаются в отчете.
function Move-AutoCADCommonSettings([object]$Book, [object]$Reference) {
    $oldRange = $Book.Names.Item('rngSystemSettings').RefersToRange
    $oldValues = $oldRange.Formula
    $values = @{}
    for ($r=2; $r -le $oldRange.Rows.Count; $r++) {
        $key = [string]$oldValues[$r,1]
        if ($key -and -not $key.StartsWith('[')) {
            if ($values.ContainsKey($key)) { throw "Duplicate original setting: $key" }
            $values[$key] = $oldValues[$r,2]
        }
    }
    $aliases = @{
        'AutoCAD.Common.ConcreteLayer' = @('AutoCAD.Import.ConcreteLayer','AutoCAD.Layer.Concrete')
        'AutoCAD.Common.RebarLayer' = @('AutoCAD.Import.RebarLayer','AutoCAD.Layer.Rebar')
        'AutoCAD.Export.CrackInteractionLayer' = @('AutoCAD.Common.CrackInteractionLayer')
    }
    foreach ($key in $aliases.Keys) {
        if ($values.ContainsKey($key)) { continue }
        $oldKeys = $aliases[$key]
        $available = @($oldKeys | Where-Object { $values.ContainsKey($_) })
        if ($available.Count -eq 0) { throw "Missing migration source: $key" }
        $selected = $available[0]
        if ($oldKeys.Count -eq 2 -and [string]$values['Geometry.Source'] -ne 'AutoCAD' -and $values.ContainsKey($oldKeys[1])) {
            $selected = $oldKeys[1]
        }
        $values[$key] = $values[$selected]
        $previous = @{}
        foreach ($oldKey in $available) { $previous[$oldKey] = $values[$oldKey] }
        $migrations.Add([pscustomobject]@{Key=$key; SelectedSource=$selected; Value=$values[$key]; PreviousValues=$previous})
    }
    $radiusKey = 'SLS.Crack.SP35.InteractionRadiusMode'
    $radius = [string]$values[$radiusKey]
    if ($radius -match '^D([356])$') {
        $values[$radiusKey] = $Matches[1] + 'd'
        $migrations.Add([pscustomobject]@{Key=$radiusKey; SelectedSource=$radius; Value=$values[$radiusKey]})
    }
    $from = $Reference.Names.Item('rngSystemSettings').RefersToRange
    $to = $oldRange.Cells.Item(1,1).Resize($from.Rows.Count,$from.Columns.Count)
    $removed = @($aliases.Values | ForEach-Object { $_ })
    $newKeys = @{}
    $referenceValues = $from.Value2
    for ($r=2; $r -le $from.Rows.Count; $r++) {
        $key = [string]$referenceValues[$r,1]
        if ($key -and -not $key.StartsWith('[')) { $newKeys[$key]=$r }
    }
    foreach ($key in $values.Keys) {
        if (-not $newKeys.ContainsKey($key) -and $key -notin $removed) { throw "Unmapped original setting: $key" }
    }
    foreach ($key in $newKeys.Keys) {
        if (-not $values.ContainsKey($key)) { throw "New setting lacks explicit source: $key" }
        if ([string]$values[$key] -match '^=' -and $aliases.ContainsKey($key)) {
            throw "Layer formula needs an explicit reference migration: $key"
        }
    }
    $oldRange.Clear()
    $from.Copy($to) | Out-Null
    $to.Formula = $from.Formula
    $Book.Names.Item('rngSystemSettings').RefersTo = "='$($to.Worksheet.Name)'!$($to.Address())"
    foreach ($key in $newKeys.Keys) {
        $to.Cells.Item($newKeys[$key],2).Formula = $values[$key]
        if ([string]$to.Cells.Item($newKeys[$key],2).Formula -cne [string]$values[$key]) { throw "Migration changed input: $key" }
    }
    Write-Output "SYSTEM_SETTINGS_MIGRATED $($newKeys.Count) keys"
}

# Проверяет сохранность всех ячеек таблиц Config, включая формулы, пустоты
# и пользовательские значения новых параметров СП 35, отсутствующих в старом census.
function Assert-ConfigPreserved([object]$Book) {
    $count = 0
    foreach ($name in $snapshots.Keys) {
        $range = $Book.Names.Item($name).RefersToRange
        $actual = $range.Formula
        $expected = $snapshots[$name]
        for ($r = 1; $r -le $range.Rows.Count; $r++) {
            for ($c = 1; $c -le $range.Columns.Count; $c++) {
                if ([string]$actual[$r, $c] -cne [string]$expected[$r, $c]) {
                    throw "Config changed: $name at $($range.Cells.Item($r,$c).Address())."
                }
                $count++
            }
        }
    }
    return $count
}

# Выгружает фактически сохраненные VBE-компоненты единым UTF-8 TXT для трассировки.
function Export-RepairedVba([object]$Book, [string]$Path) {
    $writer = New-Object IO.StreamWriter($Path, $false, (New-Object Text.UTF8Encoding($true)))
    try {
        $writer.WriteLine('VBA PROJECT EXPORT')
        $writer.WriteLine("Workbook: $($Book.Name)")
        $writer.WriteLine("Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
        foreach ($component in $Book.VBProject.VBComponents) {
            $module = $component.CodeModule
            $writer.WriteLine(('=' * 100))
            $writer.WriteLine("COMPONENT: $($component.Name)")
            $writer.WriteLine("TYPE: $($component.Type)")
            $writer.WriteLine("LINES: $($module.CountOfLines)")
            $writer.WriteLine(('=' * 100))
            if ($module.CountOfLines) { $writer.WriteLine($module.Lines(1, $module.CountOfLines)) }
        }
    } finally { $writer.Dispose() }
}

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false; $excel.AutomationSecurity = 1
    # Excel не открывает в одном Application две книги с одинаковым basename.
    $referenceBook = $excel.Workbooks.Open($referenceCopy, 0, $true)
    $book = $excel.Workbooks.Open($target)
    $config = $book.Worksheets.Item('Config')
    $referenceConfig = $referenceBook.Worksheets.Item('Config')
    $sources = @{}
    if ($MigrateAutoCADCommonSettings) { Move-AutoCADCommonSettings $book $referenceBook }
    # Проверяем размеры до любых мутаций и сохраняем весь Config-ввод.
    foreach ($name in (Get-ConfigNamedRangeNames)) {
        $from = $referenceBook.Names.Item($name).RefersToRange
        $to = $book.Names.Item($name).RefersToRange
        if ($from.Rows.Count -ne $to.Rows.Count -or $from.Columns.Count -ne $to.Columns.Count) {
            throw "Config dimensions differ: $name; explicit migration required."
        }
        $snapshots[$name] = $to.Formula
    }
    foreach ($name in (Get-ConfigNamedRangeNames)) {
        $from = $referenceBook.Names.Item($name).RefersToRange
        $to = $book.Names.Item($name).RefersToRange
        $to.Validation.Delete()
        for ($r = 1; $r -le $from.Rows.Count; $r++) {
            for ($c = 1; $c -le $from.Columns.Count; $c++) {
                $cell = $from.Cells.Item($r, $c)
                $type = 0
                try { $type = $cell.Validation.Type } catch { }
                if ($type -ne 3) { continue }
                $formula = [string]$cell.Validation.Formula1
                if (-not $formula.StartsWith('=')) { throw "Non-range validation source: $name." }
                if (-not $sources.ContainsKey($formula)) {
                    $source = $referenceConfig.Range($formula.Substring(1))
                    $destination = $config.Range($source.Address())
                    $source.Copy($destination) | Out-Null
                    $config.Columns.Item($source.Column).Hidden = $true
                    $sources[$formula] = $true
                }
                $destinationCell = $to.Cells.Item($r, $c)
                $destinationCell.Validation.Add(3, 1, 1, $formula)
                $destinationCell.Validation.IgnoreBlank = $cell.Validation.IgnoreBlank
                $destinationCell.Validation.InCellDropdown = $true
                $records.Add([pscustomobject]@{Table=$name; Cell=$destinationCell.Address(); Source=$formula})
            }
        }
    }
    Apply-ConfigNamedRangeBorders $book
    foreach ($file in Get-ChildItem -LiteralPath (Join-Path $root 'src'), (Join-Path $root 'tests') -Recurse -File | Where-Object Extension -in '.bas', '.cls') {
        $component = $null
        try { $component = $book.VBProject.VBComponents.Item($file.BaseName) } catch { }
        if ($null -ne $component) { $book.VBProject.VBComponents.Remove($component) }
        $text = [IO.File]::ReadAllText($file.FullName, [Text.Encoding]::UTF8)
        $body = [regex]::Match($text, '(?ms)^Option Explicit.*').Value
        if ([string]::IsNullOrWhiteSpace($body)) { throw "Missing Option Explicit: $($file.FullName)" }
        $type = 1
        if ($file.Extension -eq '.cls') { $type = 2 }
        $component = $book.VBProject.VBComponents.Add($type)
        $component.Name = $file.BaseName
        $component.CodeModule.AddFromString(($body -split "`r?`n") -join "`r`n")
    }
    if ($MigrateAutoCADCommonSettings) { Add-SettingsInstructions $book $config $book.Worksheets.Item('Справка') | Out-Null }
    $preserved = Assert-ConfigPreserved $book
    Write-Output "CONFIG_PRESERVED $preserved"
    $book.Save()
    $book.Close($false); $book = $null
    Restore-WorkbookPrintAreas $target $printAreas
    $book = $excel.Workbooks.Open($target)
    $preserved = Assert-ConfigPreserved $book
    $test = [string]$excel.Run("'$($book.Name)'!modTestWorkbookInterface.RunConfigPresentationTests")
    $test | Set-Content -LiteralPath (Join-Path $directory 'ConfigPresentation.txt') -Encoding UTF8
    Write-Output (($test -split "`r?`n" | Where-Object { $_ -match '^TOTAL|^FAIL:|^CONFIG_' }) -join "`n")
    if ($test -match '(?m)^FAIL[: ]|failed=[1-9]|RUNTIME ERROR:') { throw 'Config presentation test failed.' }
    $message = [string]$excel.Run("'$($book.Name)'!modWorkbookCalculation.RunSectionCalculationForWorkbook", $book, $false)
    Write-Output "CALCULATION_FINISHED $message"
    if ($message -match 'Расчет не выполнен') { throw $message }
    $preserved = Assert-ConfigPreserved $book
    $book.Save()
    $book.Close($false); $book = $null
    Restore-WorkbookPrintAreas $target $printAreas
    $book = $excel.Workbooks.Open($target, 0, $true)
    $preserved = Assert-ConfigPreserved $book
    Export-RepairedVba $book (Join-Path $directory 'VBA_All_Code.txt')
    if ((Get-FileHash -LiteralPath $current).Hash -ne $sourceHash) { throw 'Original workbook changed while preparing repair.' }
    $records | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $directory 'ValidationSources.json') -Encoding UTF8
    $migrations | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directory 'SettingMigrations.json') -Encoding UTF8
    [pscustomobject]@{
        SourceSHA256=$sourceHash; SourceUnchanged=$true; PreparedWorkbook=$target
        ConfigCellsPreserved=$preserved; ValidationCells=$records.Count; DistinctSources=$sources.Count
        CalculationMessage=$message; PreparedSHA256=(Get-FileHash -LiteralPath $target).Hash
    } | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $directory 'RepairManifest.json') -Encoding UTF8
}
finally {
    if ($null -ne $book) { $book.Close($false) }
    if ($null -ne $referenceBook) { $referenceBook.Close($false) }
    if ($null -ne $excel) { $excel.Quit() }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
