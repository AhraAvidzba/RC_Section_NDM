# Читает фактический Config непосредственно из сохраненного XLSM без Excel.
# Фиксирует также неназванные ячейки, скрытые строки/колонки, формулы и validation.
# Это структурный census, не декларация поведенческого покрытия всех настроек.
param(
    [string]$WorkbookPath = 'workbook/output/RC_Section_NDM.xlsm',
    [string]$OutputPath = 'docs/regression/Audit03/config_census_2026-10-02.json'
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
$root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$path = (Resolve-Path -LiteralPath (Join-Path $root $WorkbookPath)).Path
$beforeHash = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
$archive = [IO.Compression.ZipFile]::OpenRead($path)

# Возвращает XML известной части OpenXML; отсутствие обязательной части явно
# останавливает census вместо пустого списка, который мог бы выглядеть полным.
function Read-XmlPart([string]$part, [bool]$required = $true) {
    $entry = $archive.GetEntry($part)
    if (-not $entry) {
        if ($required) { throw "В книге отсутствует $part." }
        return $null
    }
    $reader = [IO.StreamReader]::new($entry.Open())
    try { return [xml]$reader.ReadToEnd() } finally { $reader.Dispose() }
}

# Раскрывает shared/inline rich text без потери отдельных текстовых фрагментов.
function Get-CellText($cell) {
    $valueNode = $cell.SelectSingleNode('./*[local-name()="v"]')
    if ([string]$cell.t -eq 's') { return $sharedStrings[[int]$valueNode.InnerText] }
    if ([string]$cell.t -eq 'inlineStr') {
        return (@($cell.SelectNodes('.//*[local-name()="t"]') | ForEach-Object { $_.InnerText }) -join '')
    }
    if ($valueNode) { return [string]$valueNode.InnerText }
    return ''
}

# Переводит буквенную колонку в индекс для проверки принадлежности диапазону.
function Get-ColumnNumber([string]$letters) {
    $number = 0
    foreach ($char in $letters.ToUpperInvariant().ToCharArray()) {
        $number = 26 * $number + [int]$char - [int][char]'A' + 1
    }
    return $number
}

# Разбирает только адреса сохраненных ячеек/прямоугольных диапазонов. Динамические
# defined names сохраняются как текст и отдельно отмечаются, не теряются молча.
function Get-Rectangle([string]$address) {
    if ($address -notmatch '^\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$') { return $null }
    $left = Get-ColumnNumber $Matches[1]
    $top = [int]$Matches[2]
    $right = $left
    $bottom = $top
    if ($Matches[3]) {
        $right = Get-ColumnNumber $Matches[3]
        $bottom = [int]$Matches[4]
    }
    return [pscustomobject]@{ Left = $left; Top = $top; Right = $right; Bottom = $bottom }
}

# Индексирует ограниченный диапазон один раз. Массовый обход ячеек затем
# получает принадлежность за один lookup, без повторного перебора validation.
function Add-RectangleToIndex($rectangle, [string]$label, [hashtable]$index) {
    if (-not $rectangle) { return }
    for ($row = $rectangle.Top; $row -le $rectangle.Bottom; $row++) {
        for ($column = $rectangle.Left; $column -le $rectangle.Right; $column++) {
            $key = "$row,$column"
            if (-not $index.ContainsKey($key)) {
                $index[$key] = [System.Collections.Generic.List[string]]::new()
            }
            if (-not $index[$key].Contains($label)) { $index[$key].Add($label) }
        }
    }
}

try {
    $book = Read-XmlPart 'xl/workbook.xml'
    $relationships = Read-XmlPart 'xl/_rels/workbook.xml.rels'
    $config = @($book.workbook.sheets.sheet | Where-Object { $_.name -eq 'Config' })
    if ($config.Count -ne 1) { throw "Ожидался один лист Config, найдено $($config.Count)." }
    $relationshipId = $config[0].GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
    $relationship = @($relationships.Relationships.Relationship | Where-Object { $_.Id -eq $relationshipId })
    if ($relationship.Count -ne 1) { throw 'Не найдена однозначная часть листа Config.' }
    $target = [string]$relationship[0].Target
    if ($target.StartsWith('/')) { $sheetPart = $target.TrimStart('/') } else { $sheetPart = 'xl/' + $target }
    $sheet = Read-XmlPart $sheetPart
    $styles = Read-XmlPart 'xl/styles.xml'
    $baseFormats = @($styles.styleSheet.cellStyleXfs.xf)
    $effectiveStyles = @($styles.styleSheet.cellXfs.xf | ForEach-Object {
        $alignment = $_.SelectSingleNode('./*[local-name()="alignment"]')
        if (-not $alignment -and $_.HasAttribute('xfId')) {
            $alignment = $baseFormats[[int]$_.xfId].SelectSingleNode('./*[local-name()="alignment"]')
        }
        $horizontal = 'general'; $vertical = 'bottom'
        if ($alignment) {
            if ($alignment.HasAttribute('horizontal')) { $horizontal = [string]$alignment.horizontal }
            if ($alignment.HasAttribute('vertical')) { $vertical = [string]$alignment.vertical }
        }
        [pscustomobject]@{ Horizontal = $horizontal; Vertical = $vertical }
    })
    $columnStyleIndex = @{}
    foreach ($columnRange in $sheet.worksheet.cols.col) {
        if ($columnRange.HasAttribute('style')) {
            for ($column = [int]$columnRange.min; $column -le [int]$columnRange.max; $column++) {
                $columnStyleIndex[$column] = [int]$columnRange.style
            }
        }
    }
    $stringsPart = Read-XmlPart 'xl/sharedStrings.xml' $false
    $sharedStrings = @()
    if ($stringsPart) {
        $sharedStrings = @($stringsPart.sst.si | ForEach-Object {
            @($_.SelectNodes('.//*[local-name()="t"]') | ForEach-Object { $_.InnerText }) -join ''
        })
    }
    $namedRanges = @($book.workbook.definedNames.definedName | Where-Object { $_.InnerText -match "^'?Config'?!(.+)$" } |
        ForEach-Object {
            $formula = $_.InnerText
            $address = $formula.Substring($formula.IndexOf('!') + 1)
            [pscustomobject]@{ Name = [string]$_.name; RefersTo = $formula; Rectangle = (Get-Rectangle $address) }
        })
    $validations = @($sheet.worksheet.dataValidations.dataValidation | Where-Object { $_ } | ForEach-Object {
        [pscustomobject]@{
            Range = [string]$_.sqref; Type = [string]$_.type; Operator = [string]$_.operator
            Formula1 = [string]$_.formula1; Formula2 = [string]$_.formula2
            AllowBlank = [string]$_.allowBlank; ErrorStyle = [string]$_.errorStyle
            ShowErrorMessage = [string]$_.showErrorMessage; ErrorTitle = [string]$_.errorTitle
            Error = [string]$_.error; Rectangles = @(([string]$_.sqref -split ' ') | ForEach-Object { Get-Rectangle $_ })
        }
    })
    $hiddenColumns = @($sheet.worksheet.cols.col | Where-Object { $_.hidden -eq '1' })
    $hyperlinks = @($sheet.worksheet.hyperlinks.hyperlink | Where-Object { $_ } | ForEach-Object {
        [pscustomobject]@{ Range = [string]$_.ref; Location = [string]$_.location; Display = [string]$_.display }
    })
    $mergedRanges = @($sheet.worksheet.mergeCells.mergeCell | Where-Object { $_ } | ForEach-Object {
        $address = [string]$_.ref
        [pscustomobject]@{ Range = $address; Anchor = ($address -split ':')[0]; Rectangle = (Get-Rectangle $address) }
    })
    $namesIndex = @{}
    foreach ($range in $namedRanges) { Add-RectangleToIndex $range.Rectangle $range.Name $namesIndex }
    $validationIndex = @{}
    foreach ($validation in $validations) {
        foreach ($rectangle in $validation.Rectangles) {
            Add-RectangleToIndex $rectangle $validation.Range $validationIndex
        }
    }
    $mergeIndex = @{}
    foreach ($merged in $mergedRanges) {
        Add-RectangleToIndex $merged.Rectangle $merged.Range $mergeIndex
    }
    $hiddenColumnIndex = @{}
    foreach ($columnRange in $hiddenColumns) {
        for ($column = [int]$columnRange.min; $column -le [int]$columnRange.max; $column++) {
            $hiddenColumnIndex[$column] = $true
        }
    }
    $cells = New-Object System.Collections.Generic.List[object]
    foreach ($row in $sheet.worksheet.sheetData.row) {
        foreach ($cell in $row.c) {
            $address = [string]$cell.r
            $position = Get-Rectangle $address
            if (-not $position) { throw "Не удалось разобрать адрес Config!$address." }
            $key = "$($position.Top),$($position.Left)"
            $names = @()
            if ($namesIndex.ContainsKey($key)) { $names = @($namesIndex[$key].ToArray()) }
            $validation = @()
            if ($validationIndex.ContainsKey($key)) { $validation = @($validationIndex[$key].ToArray()) }
            $formulaNode = $cell.SelectSingleNode('./*[local-name()="f"]')
            $formula = ''
            if ($formulaNode) { $formula = [string]$formulaNode.InnerText }
            $mergeRange = ''
            $mergeAnchor = ''
            if ($mergeIndex.ContainsKey($key)) {
                if ($mergeIndex[$key].Count -ne 1) { throw "Overlapping merged areas at Config!$address." }
                $mergeRange = $mergeIndex[$key][0]
                $mergeAnchor = ($mergeRange -split ':')[0]
            }
            $styleIndex = 0
            if ($cell.HasAttribute('s')) { $styleIndex = [int]$cell.s }
            elseif ($row.HasAttribute('s') -and $row.customFormat -eq '1') { $styleIndex = [int]$row.s }
            elseif ($columnStyleIndex.ContainsKey($position.Left)) { $styleIndex = $columnStyleIndex[$position.Left] }
            if ($styleIndex -ge $effectiveStyles.Count) { throw "Unknown style index at Config!$address : $styleIndex" }
            $cells.Add([pscustomobject]@{
                Address = $address; Value = (Get-CellText $cell); Formula = $formula
                OpenXmlType = [string]$cell.t; Style = [string]$cell.s
                RowHidden = ($row.hidden -eq '1'); ColumnHidden = $hiddenColumnIndex.ContainsKey($position.Left)
                NamedRanges = $names; Validations = $validation
                MergeRange = $mergeRange; MergeAnchor = $mergeAnchor
                EffectiveStyle = $styleIndex
                HorizontalAlignment = $effectiveStyles[$styleIndex].Horizontal
                VerticalAlignment = $effectiveStyles[$styleIndex].Vertical
                CoverageStatus = 'NotYetMappedToBehaviorTest'
            })
        }
    }
    $result = [pscustomobject]@{
        Workbook = $WorkbookPath; SHA256 = $beforeHash; Sheet = 'Config'; SheetPart = $sheetPart
        CapturedAt = [DateTimeOffset]::Now.ToString('o'); CellCount = $cells.Count
        NamedRangeCount = $namedRanges.Count; ValidationCount = $validations.Count
        SheetState = [string]$config[0].state; SheetProtection = [string]$sheet.worksheet.sheetProtection.sheet
        NamedRanges = $namedRanges; Validations = $validations; Hyperlinks = $hyperlinks; MergedRanges = $mergedRanges
        HiddenRows = @($sheet.worksheet.sheetData.row | Where-Object { $_.hidden -eq '1' } | ForEach-Object { [int]$_.r })
        HiddenColumns = @($hiddenColumns | ForEach-Object { [pscustomobject]@{ From = [int]$_.min; To = [int]$_.max } })
        Cells = $cells.ToArray()
    }
    $output = Join-Path $root $OutputPath
    New-Item -ItemType Directory -Force -Path (Split-Path -Parent $output) | Out-Null
    $result | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $output -Encoding UTF8
    if ((Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash -ne $beforeHash) { throw 'Книга изменилась при read-only census.' }
    Write-Output "CONFIG_CENSUS: cells=$($cells.Count); ranges=$($namedRanges.Count); validations=$($validations.Count); source unchanged=True; output=$OutputPath"
} finally {
    $archive.Dispose()
}
