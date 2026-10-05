# Готовит отдельную книгу для публикации СП 35 с сохранением пользовательского
# ввода. Исходная рабочая книга открывается read-only и не заменяется этим скриптом.
param(
    [Parameter(Mandatory=$true)][string]$CandidateWorkbook,
    [string]$CurrentWorkbook = 'workbook/output/RC_Section_NDM.xlsm',
    [string]$ReportDirectory = 'docs/regression/SP35/Publication'
)
$ErrorActionPreference = 'Stop'
$root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
. (Join-Path $PSScriptRoot 'SettingsCatalog.ps1')
$current = (Resolve-Path -LiteralPath (Join-Path $root $CurrentWorkbook)).Path
$candidate = (Resolve-Path -LiteralPath (Join-Path $root $CandidateWorkbook)).Path
$directory = [IO.Path]::GetFullPath((Join-Path $root $ReportDirectory))
if ($directory -eq (Split-Path -Parent $current)) { throw 'Publication must be prepared separately.' }
New-Item -ItemType Directory -Path $directory -Force | Out-Null
$targetPath = Join-Path $directory 'RC_Section_NDM.xlsm'
if ($targetPath -eq $current -or $targetPath -eq $candidate) { throw 'Publication copy must differ from both inputs.' }
$currentHash = (Get-FileHash -LiteralPath $current -Algorithm SHA256).Hash
$sourceCopy = Join-Path $directory 'BeforePublication.xlsm'
Copy-Item -LiteralPath $current -Destination $sourceCopy -Force
Copy-Item -LiteralPath $candidate -Destination $targetPath -Force
$printAreas = @(Get-WorkbookPrintAreas $targetPath)
$registry = Import-Csv -LiteralPath (Join-Path $root 'docs/regression/Audit03/config_field_registry_clean_v322_2026-10-05.csv')
$census = Get-Content -LiteralPath (Join-Path $root 'docs/regression/Audit03/config_census_clean_v322_2026-10-05.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$rectangles = @{}
foreach ($range in $census.NamedRanges) { $rectangles[$range.Name] = $range.Rectangle }
$ledger = New-Object 'System.Collections.Generic.List[object]'
$excel = $null
$sourceBook = $null
$targetBook = $null

# Возвращает адрес ячейки реестра относительно диапазона, а не старого листа.
function Get-FieldOffset([object]$Field) {
    if ($Field.Address -notmatch '!([A-Z]+)([0-9]+)$') { throw "Unsupported field address: $($Field.Address)" }
    $column = 0
    foreach ($char in $Matches[1].ToCharArray()) { $column = 26 * $column + [int]$char - 64 }
    $row = [int]$Matches[2]
    $rectangle = $rectangles[$Field.Block]
    if ($null -eq $rectangle) { throw "Missing named range metadata: $($Field.Block)" }
    return @(($row - [int]$rectangle.Top), ($column - [int]$rectangle.Left))
}

# Находит строку настройки по ключу в фактическом именованном диапазоне.
# Дубли и отсутствующие настройки запрещены, чтобы не переносить неверный ввод.
function Find-SettingCell([object]$Book, [string]$Key) {
    $range = $Book.Names.Item('rngSystemSettings').RefersToRange
    $data = $range.Value2
    $found = 0
    for ($i = 1; $i -le $range.Rows.Count; $i++) {
        if ([string]$data[$i, 1] -eq $Key) {
            if ($found) { throw "Duplicate setting: $Key" }
            $found = $i
        }
    }
    if (-not $found) { throw "Missing setting: $Key" }
    return $range.Cells.Item($found, 2)
}

# Сохраняет число, текст или пустоту без изменения типа. Формулы потребовали бы
# отдельного переноса ссылок; при их наличии останавливаемся без изменения source.
function Copy-InputCell([string]$Id, [string]$Block, [object]$From, [object]$To) {
    if ([bool]$From.HasFormula) {
        # ID длительной части уже связан с тем же rngLoadCombinations формулой
        # шаблона. Сохраняем именно формулу, а не ее текущий кешированный ID.
        if ($Block -ne 'rngStabilityDurationLoads' -or $Id -notmatch '\.1$' -or
            [string]$From.Formula -cne [string]$To.Formula) {
            throw "Formula needs explicit migration: $Id at $($From.Address())"
        }
        $ledger.Add([pscustomobject]@{
            Id = $Id; Block = $Block; SourceAddress = $From.Address(); TargetAddress = $To.Address()
            Value = $From.Value2; Formula = [string]$From.Formula; TargetSheet = $To.Worksheet.Name
        })
        return
    }
    $value = $From.Value2
    # Отдельные типизированные COM-вызовы избегают ошибки кэша late binding
    # Windows PowerShell при чередовании текста и Double на одном call-site.
    if ($null -eq $value) { $To.ClearContents() | Out-Null }
    elseif ($value -is [string]) { $To.Value2 = [string]$value }
    elseif ($value -is [bool]) { $To.Value2 = [bool]$value }
    else { $To.Value2 = [double]$value }
    $ledger.Add([pscustomobject]@{
        Id = $Id; Block = $Block; SourceAddress = $From.Address(); TargetAddress = $To.Address()
        Value = $value; TargetSheet = $To.Worksheet.Name
    })
}

# Проверяет весь перенесенный ввод после расчета и после повторного открытия.
function Assert-PreservedInputs([object]$Book) {
    foreach ($entry in $ledger) {
        $cell = $Book.Worksheets.Item($entry.TargetSheet).Range($entry.TargetAddress)
        $actual = $cell.Value2
        $expected = $entry.Value
        if ($entry.PSObject.Properties.Name -contains 'Formula') {
            if ([string]$cell.Formula -cne $entry.Formula) { throw "Formula changed: $($entry.Id)" }
            continue
        }
        if ($null -eq $expected -or [string]$expected -eq '') {
            if ($null -ne $actual -and [string]$actual -ne '') { throw "Blank input changed: $($entry.Id)" }
        } elseif ($expected -is [string]) {
            if ($actual -isnot [string] -or $actual -cne $expected) { throw "Text input changed: $($entry.Id)" }
        } elseif ($actual -ne $expected) { throw "Numeric input changed: $($entry.Id)" }
    }
}

# Генерирует единый UTF-8 TXT именно из реально сохраненных компонентов книги.
# Этот файл не собирается склейкой src и содержит также document modules.
function Export-AllVba([object]$Book, [string]$Path) {
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
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = 1
    $sourceBook = $excel.Workbooks.Open($sourceCopy, 0, $true)
    $targetBook = $excel.Workbooks.Open($targetPath)
    # Кандидат может предшествовать последней оформительской правке. Публикуем
    # актуальные src/tests, сохраняя созданные сборкой document modules.
    foreach ($file in Get-ChildItem -LiteralPath (Join-Path $root 'src'), (Join-Path $root 'tests') -Recurse -File | Where-Object Extension -in '.bas', '.cls') {
        $component = $null
        try { $component = $targetBook.VBProject.VBComponents.Item($file.BaseName) } catch { }
        if ($null -ne $component) { $targetBook.VBProject.VBComponents.Remove($component) }
        $text = [IO.File]::ReadAllText($file.FullName, [Text.Encoding]::UTF8)
        $body = [regex]::Match($text, '(?ms)^Option Explicit.*').Value
        if ([string]::IsNullOrWhiteSpace($body)) { throw "Missing Option Explicit: $($file.FullName)" }
        $type = 1
        if ($file.Extension -eq '.cls') { $type = 2 }
        $component = $targetBook.VBProject.VBComponents.Add($type)
        $component.Name = $file.BaseName
        $component.CodeModule.AddFromString(($body -split "`r?`n") -join "`r`n")
    }
    if ([string](Find-SettingCell $sourceBook 'Geometry.Source').Value2 -ne 'Generated') {
        throw 'Imported geometry requires a separate snapshot-preserving migration; source workbook was not modified.'
    }
    foreach ($field in $registry | Where-Object Role -eq 'UserInput') {
        if ($field.Block -in @('rngLoadCombinations', 'rngStabilityDurationLoads')) { continue }
        if ($field.Block -eq 'rngSystemSettings') {
            $from = Find-SettingCell $sourceBook $field.Id
            # Однократный перенос удаленного поля книги, не runtime-алиас.
            $targetKey = $field.Id
            if ($targetKey -eq 'AutoCAD.Layer.Contour') { $targetKey = 'AutoCAD.Common.SectionContourLayer' }
            $to = Find-SettingCell $targetBook $targetKey
        } else {
            $offset = Get-FieldOffset $field
            $fromRange = $sourceBook.Names.Item($field.Block).RefersToRange
            $toRange = $targetBook.Names.Item($field.Block).RefersToRange
            $from = $fromRange.Cells.Item(1 + $offset[0], 1 + $offset[1])
            $to = $toRange.Cells.Item(1 + $offset[0], 1 + $offset[1])
        }
        Copy-InputCell $field.Id $field.Block $from $to
    }
    # Переносим все фактические строки сочетаний и длительной части, включая пустые.
    # При несовпадении емкости не вставляем строки наугад в другие Config-таблицы.
    foreach ($name in 'rngLoadCombinations', 'rngStabilityDurationLoads', 'rngSP35Table721') {
        $from = $sourceBook.Names.Item($name).RefersToRange
        $to = $targetBook.Names.Item($name).RefersToRange
        if ($from.Rows.Count -ne $to.Rows.Count -or $from.Columns.Count -ne $to.Columns.Count) {
            throw "Named range dimensions changed: $name; explicit resizing required."
        }
        for ($row = 2; $row -le $from.Rows.Count; $row++) {
            for ($col = 1; $col -le $from.Columns.Count; $col++) {
                Copy-InputCell "$name.$row.$col" $name $from.Cells.Item($row, $col) $to.Cells.Item($row, $col)
            }
        }
    }
    # Единовременная публикация сохраняет выбранные пользователем ширины Results.
    # Runtime writer-ы этим не занимаются.
    for ($col = 1; $col -le 84; $col++) {
        $targetBook.Worksheets.Item('Results').Columns.Item($col).ColumnWidth = $sourceBook.Worksheets.Item('Results').Columns.Item($col).ColumnWidth
    }
    Assert-PreservedInputs $targetBook
    Write-Output "INPUTS_PRESERVED $($ledger.Count)"
    # Справка и ссылки обновляются штатным каталогом без переноса старых текстов.
    # Сохранность каждого пользовательского поля подтверждаем до и после.
    Add-SettingsInstructions $targetBook $targetBook.Worksheets.Item('Config') $targetBook.Worksheets.Item('Справка')
    Write-Output 'HELP_UPDATED'
    Assert-PreservedInputs $targetBook
    $message = [string]$excel.Run("'$($targetBook.Name)'!modWorkbookCalculation.RunSectionCalculationForWorkbook", $targetBook, $false)
    Write-Output "CALCULATION_FINISHED $message"
    Assert-PreservedInputs $targetBook
    $targetBook.Save()
    $targetBook.Close($false)
    $targetBook = $null
    $sourceBook.Close($false)
    $sourceBook = $null
    Restore-WorkbookPrintAreas $targetPath $printAreas
    $targetBook = $excel.Workbooks.Open($targetPath, 0, $true)
    Assert-PreservedInputs $targetBook
    Export-AllVba $targetBook (Join-Path $directory 'VBA_All_Code.txt')
    $currentAfter = (Get-FileHash -LiteralPath $current -Algorithm SHA256).Hash
    if ($currentAfter -ne $currentHash) { throw 'Source workbook changed during preparation.' }
    $manifest = [pscustomobject]@{
        CurrentWorkbook = $current; SourceSHA256 = $currentHash; SourceUnchanged = $true
        CandidateWorkbook = $candidate; PreparedWorkbook = $targetPath
        InputCellsPreserved = $ledger.Count; InputsVerifiedAfterReopen = $true
        NewCrackCode = [string](Find-SettingCell $targetBook 'SLS.Crack.Code').Value2
        CalculationMessage = $message
        PreparedSHA256 = (Get-FileHash -LiteralPath $targetPath -Algorithm SHA256).Hash
    }
    $ledger | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directory 'PreservedInputs.json') -Encoding UTF8
    $manifest | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directory 'PublicationManifest.json') -Encoding UTF8
    $manifest | Format-List
}
catch {
    Write-Output $_.ScriptStackTrace
    throw
}
finally {
    if ($null -ne $targetBook) { $targetBook.Close($false) }
    if ($null -ne $sourceBook) { $sourceBook.Close($false) }
    if ($null -ne $excel) { $excel.Quit() }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
