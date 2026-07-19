param(
    [string]$OutputPath = "workbook/output/RC_Section_NDM.xlsm"
)

$ErrorActionPreference = "Stop"

function Set-Cell {
    param(
        [object]$Sheet,
        [int]$Row,
        [int]$Column,
        [object]$Value,
        [switch]$Bold,
        [int]$Color = -1,
        [int]$InteriorColor = -1
    )

    $cell = $Sheet.Cells.Item($Row, $Column)
    $cell.Value2 = $Value
    if ($Bold) { $cell.Font.Bold = $true }
    if ($Color -ge 0) { $cell.Font.Color = $Color }
    if ($InteriorColor -ge 0) { $cell.Interior.Color = $InteriorColor }
    return $cell
}

function Set-Border {
    param([object]$Range)

    foreach ($idx in 7, 8, 9, 10, 11, 12) {
        $Range.Borders.Item($idx).LineStyle = 1
        $Range.Borders.Item($idx).Weight = 2
        $Range.Borders.Item($idx).Color = 12632256
    }
}

function ConvertTo-ExcelColumn {
    param([int]$ColumnNumber)
    $name = ""
    while ($ColumnNumber -gt 0) {
        $mod = ($ColumnNumber - 1) % 26
        $name = [char](65 + $mod) + $name
        $ColumnNumber = [math]::Floor(($ColumnNumber - $mod) / 26)
    }
    $name
}

function Add-WorkbookName {
    param(
        [object]$Workbook,
        [string]$Name,
        [object]$Sheet,
        [string]$Address
    )

    $sheetName = [string]$Sheet.Name
    $sheetName = $sheetName.Replace("'", "''")
    $address = "='" + $sheetName + "'!" + $Address
    try { $Workbook.Names.Item($Name).RefersTo = $address }
    catch { $Workbook.Names.Add($Name, $address) | Out-Null }
}

function Import-VbaSourceTree {
    param(
        [object]$Workbook,
        [string]$RootPath
    )

    $sourceDirs = @(
        (Join-Path $RootPath "src/Common"),
        (Join-Path $RootPath "src/Interfaces"),
        (Join-Path $RootPath "src/Geometry"),
        (Join-Path $RootPath "src/Section"),
        (Join-Path $RootPath "src/Materials"),
        (Join-Path $RootPath "src/Solver"),
        (Join-Path $RootPath "src/Crack"),
        (Join-Path $RootPath "src/Batch"),
        (Join-Path $RootPath "src/Excel"),
        (Join-Path $RootPath "tests")
    )

    foreach ($dir in $sourceDirs) {
        if (Test-Path -LiteralPath $dir) {
            Get-ChildItem -LiteralPath $dir -Recurse -File |
                Where-Object { $_.Extension -in @(".bas", ".cls", ".frm") } |
                Sort-Object FullName |
                ForEach-Object {
                    Add-VbaSourceFile $Workbook $_.FullName
                }
        }
    }
}

function Add-VbaSourceFile {
    param(
        [object]$Workbook,
        [string]$Path
    )

    $extension = [System.IO.Path]::GetExtension($Path).ToLowerInvariant()
    if ($extension -eq ".frm") {
        $Workbook.VBProject.VBComponents.Import($Path) | Out-Null
        return
    }

    $source = [System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::UTF8)
    $lines = $source -split "`r?`n"
    $componentName = $null
    foreach ($line in $lines) {
        if ($line -match '^\uFEFF?Attribute\s+VB_Name\s*=\s*"([^"]+)"') {
            $componentName = $Matches[1]
            break
        }
    }
    if ([string]::IsNullOrWhiteSpace($componentName)) {
        throw "VBA source has no Attribute VB_Name: $Path"
    }

    $componentType = 1
    if ($extension -eq ".cls") {
        $componentType = 2
    }

    $component = $Workbook.VBProject.VBComponents.Add($componentType)
    $component.Name = $componentName

    $bodyLines = New-Object System.Collections.Generic.List[string]
    $insideVbaHeaderBlock = $false
    foreach ($line in $lines) {
        if ($line -match '^\uFEFF?VERSION\s+') { continue }
        if ($line -match '^\uFEFF?BEGIN\s*$') {
            $insideVbaHeaderBlock = $true
            continue
        }
        if ($insideVbaHeaderBlock) {
            if ($line -match '^\uFEFF?END\s*$') {
                $insideVbaHeaderBlock = $false
            }
            continue
        }
        if ($line -match '^\uFEFF?Attribute\s+VB_') { continue }
        $bodyLines.Add($line)
    }

    $body = ($bodyLines -join "`r`n").Trim()
    if ($body.Length -gt 0) {
        $component.CodeModule.AddFromString($body)
    }
}

function Add-BlockHeader {
    param(
        [object]$Sheet,
        [int]$StartColumn,
        [string]$Title
    )

    $endColumn = $StartColumn + 17
    $range = $Sheet.Range($Sheet.Cells.Item(1, $StartColumn), $Sheet.Cells.Item(2, $endColumn))
    $range.Merge() | Out-Null
    $range.Value2 = $Title
    $range.Font.Bold = $true
    $range.Font.Size = 16
    $range.Font.Color = 16777215
    $range.Interior.Color = 5383702
    $range.HorizontalAlignment = -4108
    $range.VerticalAlignment = -4108
}

function Add-SectionTitle {
    param(
        [object]$Sheet,
        [int]$Row,
        [int]$StartColumn,
        [int]$EndColumn,
        [string]$Title
    )

    $range = $Sheet.Range($Sheet.Cells.Item($Row, $StartColumn), $Sheet.Cells.Item($Row, $EndColumn))
    $range.Merge() | Out-Null
    $range.Value2 = $Title
    $range.Font.Bold = $true
    $range.Font.Color = 16777215
    $range.Interior.Color = 8355711
    $range.HorizontalAlignment = -4131
}

function Add-KeyValueRows {
    param(
        [object]$Sheet,
        [int]$StartRow,
        [int]$StartColumn,
        [array]$Rows
    )

    $row = $StartRow
    foreach ($item in $Rows) {
        $Sheet.Cells.Item($row, $StartColumn).Value2 = $item[0]
        $Sheet.Cells.Item($row, $StartColumn + 4).Value2 = $item[1]
        $Sheet.Cells.Item($row, $StartColumn + 8).Value2 = $item[2]
        $Sheet.Cells.Item($row, $StartColumn + 12).Value2 = $item[3]
        $row++
    }

    $range = $Sheet.Range($Sheet.Cells.Item($StartRow, $StartColumn), $Sheet.Cells.Item($row - 1, $StartColumn + 16))
    Set-Border $range
    $Sheet.Range($Sheet.Cells.Item($StartRow, $StartColumn), $Sheet.Cells.Item($row - 1, $StartColumn + 3)).Font.Bold = $true
}

function Add-InputValidationList {
    param(
        [object]$Cell,
        [string]$List
    )

    $Cell.Validation.Delete()
    $Cell.Validation.Add(3, 1, 1, $List)
}

function Add-ResultBlock {
    param(
        [object]$Sheet,
        [int]$StartColumn,
        [string]$Title
    )

    Add-BlockHeader $Sheet $StartColumn $Title
    Add-SectionTitle $Sheet 4 $StartColumn ($StartColumn + 17) "Расчет по прочности и трещиностойкости"

    $rows = @(
        @("Сочетание", "", "", ""),
        @("N", "", "Н", ""),
        @("Mx", "", "Н*мм", ""),
        @("My", "", "Н*мм", ""),
        @("Статус", "Заполняется после расчета", "", ""),
        @("Примечание", "Единая постановка N + Mx + My", "", "")
    )
    Add-KeyValueRows $Sheet 7 $StartColumn $rows

    Add-SectionTitle $Sheet 16 $StartColumn ($StartColumn + 17) "Результаты"
    $resultHeaders = @("Показатель", "Значение", "Ед.", "Комментарий")
    for ($i = 0; $i -lt $resultHeaders.Count; $i++) {
        $cell = Set-Cell $Sheet 18 ($StartColumn + $i * 4) $resultHeaders[$i] -Bold -InteriorColor 14277081
        $Sheet.Range($cell, $Sheet.Cells.Item(18, $StartColumn + $i * 4 + 3)).Merge() | Out-Null
    }
    $resultRange = $Sheet.Range($Sheet.Cells.Item(18, $StartColumn), $Sheet.Cells.Item(36, $StartColumn + 15))
    Set-Border $resultRange

    Add-SectionTitle $Sheet 36 $StartColumn ($StartColumn + 17) "Диагностика"
    $diagnosticRange = $Sheet.Range($Sheet.Cells.Item(38, $StartColumn), $Sheet.Cells.Item(50, $StartColumn + 17))
    Set-Border $diagnosticRange
}
function Add-MainInputBlock {
    param([object]$Sheet)

    Add-BlockHeader $Sheet 1 "Исходные данные"

    Add-SectionTitle $Sheet 4 1 18 "Геометрия и материалы"
    $mainRows = @(
        @("Объект / элемент", "", "", ""),
        @("Тип сечения", "RoundedRectangle", "", "Задается на листе System"),
        @("Ширина", "", "мм", ""),
        @("Высота", "", "мм", ""),
        @("Радиус верхний левый", "", "мм", ""),
        @("Радиус верхний правый", "", "мм", ""),
        @("Радиус нижний правый", "", "мм", ""),
        @("Радиус нижний левый", "", "мм", ""),
        @("Диаметр круга", "", "мм", "Используется при Geometry.Type = Circle"),
        @("Центр круга X", "0", "мм", "Используется при Geometry.Type = Circle"),
        @("Центр круга Y", "0", "мм", "Используется при Geometry.Type = Circle"),
        @("Класс бетона", "", "", ""),
        @("Класс арматуры", "", "", ""),
        @("Защитный слой до оси арматуры", "", "мм", ""),
        @("Диаметр стержней", "", "мм", "Для автоматической арматуры круглого сечения")
    )
    Add-KeyValueRows $Sheet 6 1 $mainRows

    Add-SectionTitle $Sheet 23 1 18 "Контрольная таблица арматуры"
    $rebarHeaders = @("ID", "X", "Y", "Diameter", "Area", "SteelClass", "Comment")
    for ($i = 0; $i -lt $rebarHeaders.Count; $i++) {
        Set-Cell $Sheet 25 ($i + 1) $rebarHeaders[$i] -Bold -InteriorColor 14277081 | Out-Null
    }
    $rebarRange = $Sheet.Range($Sheet.Cells.Item(25, 1), $Sheet.Cells.Item(35, 7))
    Set-Border $rebarRange

    Add-SectionTitle $Sheet 38 1 18 "Сочетания нагрузок"
    $loadHeaders = @("CombinationID", "N", "Mx", "My", "CalculationType", "DurationType", "Comment")
    for ($i = 0; $i -lt $loadHeaders.Count; $i++) {
        Set-Cell $Sheet 40 ($i + 1) $loadHeaders[$i] -Bold -InteriorColor 14277081 | Out-Null
    }
    $loadRange = $Sheet.Range($Sheet.Cells.Item(40, 1), $Sheet.Cells.Item(60, 7))
    Set-Border $loadRange
}
function Add-CalculationButtons {
    param([object]$Sheet)

    $left = $Sheet.Cells.Item(4, 38).Left
    $top = $Sheet.Cells.Item(4, 38).Top

    $runButton = $Sheet.Shapes.AddShape(1, $left, $top, 150, 28)
    $runButton.Name = "btnRunSectionCalculation"
    $runButton.TextFrame.Characters().Text = "Выполнить расчет"
    $runButton.OnAction = "RunSectionCalculation"
    $runButton.Fill.ForeColor.RGB = 5383702
    $runButton.Line.ForeColor.RGB = 5383702
    $runButton.TextFrame.Characters().Font.Color = 16777215
    $runButton.TextFrame.Characters().Font.Bold = $true

    $clearButton = $Sheet.Shapes.AddShape(1, $left, $top + 36, 150, 28)
    $clearButton.Name = "btnClearSectionResults"
    $clearButton.TextFrame.Characters().Text = "Очистить результаты"
    $clearButton.OnAction = "ClearSectionResults"
    $clearButton.Fill.ForeColor.RGB = 8355711
    $clearButton.Line.ForeColor.RGB = 8355711
    $clearButton.TextFrame.Characters().Font.Color = 16777215
    $clearButton.TextFrame.Characters().Font.Bold = $true

    $acadButton = $Sheet.Shapes.AddShape(1, $left, $top + 72, 150, 28)
    $acadButton.Name = "btnExportStressToAutoCAD"
    $acadButton.TextFrame.Characters().Text = "Экспорт в AutoCAD"
    $acadButton.OnAction = "ExportSectionStressToAutoCAD"
    $acadButton.Fill.ForeColor.RGB = 10053171
    $acadButton.Line.ForeColor.RGB = 10053171
    $acadButton.TextFrame.Characters().Font.Color = 16777215
    $acadButton.TextFrame.Characters().Font.Bold = $true
}

$root = Resolve-Path -LiteralPath (Join-Path $PSScriptRoot "../..")
. (Join-Path $PSScriptRoot "SettingsCatalog.ps1")
$fullOutputPath = Join-Path $root $OutputPath
$outputDir = Split-Path -Parent $fullOutputPath
New-Item -ItemType Directory -Force -Path $outputDir | Out-Null

$excel = $null
$workbook = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false

    $workbook = $excel.Workbooks.Add()

    while ($workbook.Worksheets.Count -lt 2) {
        $workbook.Worksheets.Add() | Out-Null
    }
    while ($workbook.Worksheets.Count -gt 2) {
        $workbook.Worksheets.Item($workbook.Worksheets.Count).Delete()
    }

    $calc = $workbook.Worksheets.Item(1)
    $system = $workbook.Worksheets.Item(2)
    $calc.Name = "Расчет"
    $system.Name = "System"

    Add-MainInputBlock $calc
    Add-ResultBlock $calc 19 "Расчет N + Mx + My"

    for ($col = 1; $col -le 36; $col++) {
        $calc.Columns.Item($col).ColumnWidth = 3.7
    }
    for ($row = 1; $row -le 50; $row++) {
        $calc.Rows.Item($row).RowHeight = 14
    }
    $calc.Rows.Item(1).RowHeight = 24
    $calc.Rows.Item(2).RowHeight = 24
    $calc.Range("A1:AJ60").Font.Name = "Arial"
    $calc.Range("A1:AJ60").Font.Size = 9
    $calc.Range("A1:AJ60").VerticalAlignment = -4108

    $systemRanges = Add-SystemSettings $system
    $system.Range("A1:O80").Font.Name = "Arial"
    $system.Range("A1:O80").Font.Size = 9
    $system.Columns.Item(1).ColumnWidth = 32
    $system.Columns.Item(2).ColumnWidth = 18
    $system.Columns.Item(3).ColumnWidth = 22
    $system.Columns.Item(4).ColumnWidth = 14
    $system.Columns.Item(5).ColumnWidth = 48
    $system.Columns.Item(6).ColumnWidth = 34
    $system.Columns.Item(7).ColumnWidth = 22

    Add-WorkbookName $workbook "rngMainInput" $calc '$A$6:$Q$20'
    Add-WorkbookName $workbook "rngRebarInput" $calc '$A$25:$G$34'
    Add-WorkbookName $workbook "rngLoadCombinations" $calc '$A$40:$G$60'
    Add-WorkbookName $workbook "rngResultSection" $calc '$S$17:$AH$35'

    $calc.PageSetup.PaperSize = 9
    $calc.PageSetup.Orientation = 1
    $calc.PageSetup.Zoom = 80
    $calc.PageSetup.PrintArea = '$A$1:$AJ$60'
    $calc.PageSetup.LeftMargin = $excel.CentimetersToPoints(0.7)
    $calc.PageSetup.RightMargin = $excel.CentimetersToPoints(0.7)
    $calc.PageSetup.TopMargin = $excel.CentimetersToPoints(0.8)
    $calc.PageSetup.BottomMargin = $excel.CentimetersToPoints(0.8)
    $calc.PageSetup.CenterHorizontally = $true

    $calc.ResetAllPageBreaks()
    $calc.VPageBreaks.Add($calc.Cells.Item(1, 19)) | Out-Null
    $calc.PageSetup.Zoom = 80
    Add-CalculationButtons $calc

    $system.PageSetup.PrintArea = ""

    Import-VbaSourceTree $workbook $root

    $workbook.SaveAs($fullOutputPath, 52)
    Write-Output "Workbook built: $fullOutputPath"
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
























