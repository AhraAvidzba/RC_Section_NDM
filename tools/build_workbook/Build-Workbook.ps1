# скрипт собирает рабочую книгу Excel из исходных VBA-модулей, настроек и шаблонной структуры листов.

param(
    [string]$OutputPath = "workbook/output/RC_Section_NDM.xlsm",
    [string]$FormulaImageDirectory = ""
)

$ErrorActionPreference = "Stop"

# Устанавливает значение, оформление или именованный диапазон в книге через Excel COM.
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

# Устанавливает значение, оформление или именованный диапазон в книге через Excel COM.
function Set-Border {
    param([object]$Range)

    foreach ($idx in 7, 8, 9, 10, 11, 12) {
        $Range.Borders.Item($idx).LineStyle = 1
        $Range.Borders.Item($idx).Weight = 2
        $Range.Borders.Item($idx).Color = 12632256
    }
}

# Преобразует техническое представление в формат, удобный для Excel или отчета.
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

# Добавляет структурный элемент книги или отчета, сохраняя единый формат сборочных скриптов.
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

# Импортирует исходные VBA-модули в книгу, сохраняя воспроизводимость сборки.
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
        (Join-Path $RootPath "src/Stability"),
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
                    $sourcePath = $_.FullName
                    try {
                        Add-VbaSourceFile $Workbook $sourcePath
                    }
                    catch {
                        throw "Failed to import VBA source '$sourcePath': $($_.Exception.Message)"
                    }
                }
        }
    }
}

# Добавляет обработчики событий книги, которые нельзя импортировать как обычный стандартный модуль.
function Add-WorkbookEventHandlers {
    param([object]$Workbook)

    $code = [IO.File]::ReadAllText((Join-Path $PSScriptRoot 'WorkbookHelpEvents.vba'),[Text.Encoding]::UTF8)

    $componentName = [string]$Workbook.CodeName
    if ([string]::IsNullOrWhiteSpace($componentName)) {
        $componentName = "ThisWorkbook"
    }
    $component = $Workbook.VBProject.VBComponents.Item($componentName)
    if ($component.CodeModule.CountOfLines -gt 0) {
        $component.CodeModule.DeleteLines(1, $component.CodeModule.CountOfLines)
    }
    $component.CodeModule.AddFromString($code)
    foreach ($document in $Workbook.VBProject.VBComponents) {
        if ($document.Type -eq 100 -and $document.Name -ne $componentName -and $document.CodeModule.CountOfLines -eq 0) {
            $document.CodeModule.AddFromString("Option Explicit`r`n`r`n' Лист книги: отображение пользовательских данных и результатов.")
        }
    }
}

# Импортирует один VBA-файл в книгу. Формы .frm Excel принимает через Import,
# а классы и стандартные модули сначала создаются с правильным типом компонента
# и затем заполняются текстом из исходников.
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

# Добавляет структурный элемент книги или отчета, сохраняя единый формат сборочных скриптов.
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

# Добавляет структурный элемент книги или отчета, сохраняя единый формат сборочных скриптов.
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

# Добавляет структурный элемент книги или отчета, сохраняя единый формат сборочных скриптов.
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

# Добавляет структурный элемент книги или отчета, сохраняя единый формат сборочных скриптов.
function Add-InputValidationList {
    param(
        [object]$Cell,
        [string]$List
    )

    $Cell.Validation.Delete()
    $Cell.Validation.Add(3, 1, 1, $List)
}

# Добавляет структурный элемент книги или отчета, сохраняя единый формат сборочных скриптов.
function Add-MainInputBlock {
    param([object]$Sheet)

    Add-BlockHeader $Sheet 1 "Сочетания нагрузок"

    Add-SectionTitle $Sheet 38 1 18 "Сочетания нагрузок"
    $loadHeaders = @("CombinationID", "N", "Mx", "My", "ProfileId", "LoadPath", "Comment")
    for ($i = 0; $i -lt $loadHeaders.Count; $i++) {
        Set-Cell $Sheet 40 ($i + 1) $loadHeaders[$i] -Bold -InteriorColor 14277081 | Out-Null
    }
    $Sheet.Cells.Item(40, 2).Formula = "=`"N, `"&INDEX(rngUnitSettings,MATCH(`"Force`",INDEX(rngUnitSettings,,1),0),2)"
    $Sheet.Cells.Item(40, 3).Formula = "=`"Mx, `"&INDEX(rngUnitSettings,MATCH(`"Moment`",INDEX(rngUnitSettings,,1),0),2)"
    $Sheet.Cells.Item(40, 4).Formula = "=`"My, `"&INDEX(rngUnitSettings,MATCH(`"Moment`",INDEX(rngUnitSettings,,1),0),2)"
    $loadRange = $Sheet.Range($Sheet.Cells.Item(40, 1), $Sheet.Cells.Item(70, 7))
    Set-Border $loadRange

    $profileListColumn = 52
    $capacityLoadPathListColumn = 53
    $profileOptions = @("PR1", "PR2", "PR3", "PR4")
    for ($i = 0; $i -lt $profileOptions.Count; $i++) {
        $Sheet.Cells.Item($i + 1, $profileListColumn).Value2 = $profileOptions[$i]
    }
    $profileListAddress = '=$AZ$1:$AZ$' + $profileOptions.Count
    $profileRange = $Sheet.Range($Sheet.Cells.Item(41, 5), $Sheet.Cells.Item(70, 5))
    $profileRange.Validation.Delete()
    $profileRange.Validation.Add(3, 1, 1, $profileListAddress)
    $profileRange.Validation.IgnoreBlank = $false
    $profileRange.Validation.InCellDropdown = $true

    $lambda = [char]0x03BB
    $capacityLoadPathOptions = @("Auto", "$lambda*Mx", "$lambda*My", "$lambda*Mxy", "$lambda*N", "$lambda*NMxy")
    for ($i = 0; $i -lt $capacityLoadPathOptions.Count; $i++) {
        $Sheet.Cells.Item($i + 1, $capacityLoadPathListColumn).Value2 = $capacityLoadPathOptions[$i]
    }
    $capacityLoadPathListAddress = '=$BA$1:$BA$' + $capacityLoadPathOptions.Count
    $capacityLoadPathRange = $Sheet.Range($Sheet.Cells.Item(41, 6), $Sheet.Cells.Item(70, 6))
    $capacityLoadPathRange.Validation.Delete()
    $capacityLoadPathRange.Validation.Add(3, 1, 1, $capacityLoadPathListAddress)
    $capacityLoadPathRange.Validation.IgnoreBlank = $true
    $capacityLoadPathRange.Validation.InCellDropdown = $true
    $capacityLoadPathRange.Value2 = "Auto"
    $Sheet.Columns.Item($profileListColumn).Hidden = $true
    $Sheet.Columns.Item($capacityLoadPathListColumn).Hidden = $true
}

# Центрирует подпись внутри Shape-кнопки. Используем TextFrame для
# совместимости с Excel VBA/COM и, если доступно, TextFrame2 для более
# надежного вертикального якоря в новых версиях Office.
function Center-ShapeButtonText {
    param([object]$Button)

    $Button.TextFrame.HorizontalAlignment = -4108
    $Button.TextFrame.VerticalAlignment = -4108
    $Button.TextFrame.MarginLeft = 0
    $Button.TextFrame.MarginRight = 0
    $Button.TextFrame.MarginTop = 0
    $Button.TextFrame.MarginBottom = 0
    try {
        $Button.TextFrame2.TextRange.ParagraphFormat.Alignment = 2
        $Button.TextFrame2.VerticalAnchor = 3
        $Button.TextFrame2.MarginLeft = 0
        $Button.TextFrame2.MarginRight = 0
        $Button.TextFrame2.MarginTop = 0
        $Button.TextFrame2.MarginBottom = 0
    } catch {
        # В старых Excel TextFrame2 может быть недоступен; TextFrame выше уже достаточно.
    }
}

# Добавляет структурный элемент книги или отчета, сохраняя единый формат сборочных скриптов.
function Add-CalculationButtons {
    param([object]$Sheet)

    $left = $Sheet.Cells.Item(4, 38).Left
    $top = $Sheet.Cells.Item(4, 38).Top
    $buttonWidth = 235
    $buttonHeight = 32
    $buttonStep = 40

    $runButton = $Sheet.Shapes.AddShape(1, $left, $top + $buttonStep, $buttonWidth, $buttonHeight)
    $runButton.Name = "btnRunSectionCalculation"
    $runButton.TextFrame.Characters().Text = "Выполнить расчет"
    $runButton.OnAction = "RunSectionCalculation"
    $runButton.Fill.ForeColor.RGB = 5383702
    $runButton.Line.ForeColor.RGB = 5383702
    $runButton.TextFrame.Characters().Font.Color = 16777215
    $runButton.TextFrame.Characters().Font.Bold = $true
    $runButton.TextFrame.Characters().Font.Size = 9
    Center-ShapeButtonText $runButton

    $importButton = $Sheet.Shapes.AddShape(1, $left, $top, $buttonWidth, $buttonHeight)
    $importButton.Name = "btnImportGeometryFromAutoCAD"
    $importButton.TextFrame.Characters().Text = "Сформировать сечение"
    $importButton.OnAction = "FormSection"
    $importButton.Fill.ForeColor.RGB = 10053171
    $importButton.Line.ForeColor.RGB = 10053171
    $importButton.TextFrame.Characters().Font.Color = 16777215
    $importButton.TextFrame.Characters().Font.Bold = $true
    $importButton.TextFrame.Characters().Font.Size = 9
    Center-ShapeButtonText $importButton

    $clearButton = $Sheet.Shapes.AddShape(1, $left, $top + 4 * $buttonStep, $buttonWidth, $buttonHeight)
    $clearButton.Name = "btnClearAutoCADDrawing"
    $clearButton.TextFrame.Characters().Text = "Очистить подписи в AutoCAD"
    $clearButton.OnAction = "ClearAutoCADDrawing"
    $clearButton.Fill.ForeColor.RGB = 8355711
    $clearButton.Line.ForeColor.RGB = 8355711
    $clearButton.TextFrame.Characters().Font.Color = 16777215
    $clearButton.TextFrame.Characters().Font.Bold = $true
    $clearButton.TextFrame.Characters().Font.Size = 9
    Center-ShapeButtonText $clearButton

    $acadButton = $Sheet.Shapes.AddShape(1, $left, $top + 3 * $buttonStep, $buttonWidth, $buttonHeight)
    $acadButton.Name = "btnExportStressToAutoCAD"
    $acadButton.TextFrame.Characters().Text = "Экспорт в AutoCAD"
    $acadButton.OnAction = "ExportSectionStressToAutoCAD"
    $acadButton.Fill.ForeColor.RGB = 10053171
    $acadButton.Line.ForeColor.RGB = 10053171
    $acadButton.TextFrame.Characters().Font.Color = 16777215
    $acadButton.TextFrame.Characters().Font.Bold = $true
    $acadButton.TextFrame.Characters().Font.Size = 9
    Center-ShapeButtonText $acadButton

    $plotButton = $Sheet.Shapes.AddShape(1, $left, $top + 2 * $buttonStep, $buttonWidth, $buttonHeight)
    $plotButton.Name = "btnUpdateSectionPlot"
    $plotButton.TextFrame.Characters().Text = "Обновить схему"
    $plotButton.OnAction = "UpdateSectionPlot"
    $plotButton.Fill.ForeColor.RGB = 5287936
    $plotButton.Line.ForeColor.RGB = 5287936
    $plotButton.TextFrame.Characters().Font.Color = 16777215
    $plotButton.TextFrame.Characters().Font.Bold = $true
    $plotButton.TextFrame.Characters().Font.Size = 9
    Center-ShapeButtonText $plotButton
}

$root = Resolve-Path -LiteralPath (Join-Path $PSScriptRoot "../..")
. (Join-Path $PSScriptRoot "SettingsCatalog.ps1")
$script:WorkbookLayout = Read-WorkbookLayout
. (Join-Path $PSScriptRoot "SectionRecipeGuide.ps1")
if (-not [string]::IsNullOrWhiteSpace($FormulaImageDirectory)) {
    $script:FormulaImageDir = [IO.Path]::GetFullPath((Join-Path $root $FormulaImageDirectory))
    $workspace = [IO.Path]::GetFullPath([string]$root) + [IO.Path]::DirectorySeparatorChar
    if (-not ($script:FormulaImageDir + [IO.Path]::DirectorySeparatorChar).StartsWith($workspace, [StringComparison]::OrdinalIgnoreCase)) {
        throw 'Formula image output must stay inside the workspace.'
    }
}
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

    while ($workbook.Worksheets.Count -lt 4) {
        $workbook.Worksheets.Add() | Out-Null
    }
    while ($workbook.Worksheets.Count -gt 4) {
        $workbook.Worksheets.Item($workbook.Worksheets.Count).Delete()
    }

    $system = $workbook.Worksheets.Item(1)
    $instructions = $workbook.Worksheets.Item(2)
    $calc = $workbook.Worksheets.Item(3)
    $results = $workbook.Worksheets.Item(4)
    $system.Name = "Config"
    $instructions.Name = "Справка"
    $calc.Name = "Расчет"
    $results.Name = "Results"

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
    $results.Rows.RowHeight = 15

    $systemRanges = Add-SystemSettings $system
    Add-SettingsInstructions $workbook $system $instructions
    # Последний содержательный раздел добавляется после старых формул/рисунков.
    Update-SectionRecipeGuide $workbook ([string]$root) | Out-Null
    # Hyperlink-стиль Excel может сбросить рамку ячейки Подробнее.
    # Финальную обводку делаем после добавления всех ссылок Config.
    Apply-ConfigNamedRangeBorders $workbook
    $system.Range("A1:U80").Font.Name = "Arial"
    $system.Range("A1:U80").Font.Size = 9
    # Размеры столбцов Config уже назначены Add-SystemSettings; повторно их не меняем.
    $results.Range("A1:AE1").Font.Bold = $true
    # Заливка относится только к построенному снимку, не ко всей строке листа.
    $results.Range("A154:CL154").Interior.Color = 15652797
    $results.Cells.Item(154, 1).Value2 = "Расчетный снимок Results: элементы, геометрия, свойства сечения, диаграммы материалов и аннотации"
    $results.Cells.Item(154, 1).Font.Name = "Arial"
    $results.Cells.Item(154, 1).Font.Size = 12
    $results.Cells.Item(154, 1).Font.Bold = $true
    $results.Cells.Item(154, 1).Font.Color = 255
    $results.Cells.Item(154, 1).HorizontalAlignment = -4131
    $results.Cells.Item(154, 1).VerticalAlignment = -4108
    $results.Range("A156:I156").Font.Bold = $true
    $results.Range("L156:Z156").Font.Bold = $true
    $results.Range("AC156:AS156").Font.Bold = $true
    $results.Range("AV156:BA156").Font.Bold = $true
    $results.Range("BD156:BN156").Font.Bold = $true
    $results.Range("BQ156:CL156").Font.Bold = $true
    # Явная ширина используемых колонок не пересчитывается Excel при смене
    # шрифта ячеек, в отличие от общей default-ширины листа. Это только сборка.
    $results.Range("A:CL").ColumnWidth = 10
    $results.Columns.Item(1).ColumnWidth = 15
    $results.Columns.Item(2).ColumnWidth = 18
    $results.Columns.Item(3).ColumnWidth = 21

    Add-WorkbookName $workbook "rngBatchSummary" $results '$A$1'
    Add-WorkbookName $workbook "rngStrengthSummaryAnchor" $results '$A$49'
    Add-WorkbookName $workbook "rngCrackSummaryAnchor" $results '$A$85'
    Add-WorkbookName $workbook "rngStabilitySummaryAnchor" $results '$A$122'
    Add-WorkbookName $workbook "rngNDMElementResults" $results '$A$156'
    Add-WorkbookName $workbook "rngNDMSectionGeometry" $results '$L$156'
    Add-WorkbookName $workbook "rngNDMSectionContours" $results '$AD$156'
    Add-WorkbookName $workbook "rngNDMSectionProperties" $results '$AW$156'
    Add-WorkbookName $workbook "rngNDMMaterialDiagrams" $results '$BE$156'
    Add-WorkbookName $workbook "rngNDMSectionAnnotations" $results '$BR$156'
    $contourHeaders = @('RunID', 'LoopID', 'SegmentID', 'Sequence', 'LoopRole', 'SegmentType', 'StartX', 'StartY', 'EndX', 'EndY', 'CenterX', 'CenterY', 'Radius', 'SweepAngle, rad', 'SourceID', 'Comment', 'LengthUnit (SectionXY)')
    $contourHeaderValues = New-Object 'object[,]' 1,$contourHeaders.Count
    for ($i = 0; $i -lt $contourHeaders.Count; $i++) { $contourHeaderValues[0,$i] = $contourHeaders[$i] }
    $results.Range('AD156:AT156').Value2 = $contourHeaderValues
    $results.Range('AD155:AT155').Merge()
    $results.Cells.Item(155, 30).Value2 = 'Достоверные контуры и отверстия сечения'
    $results.Range('AD155:AT156').Font.Bold = $true
    Set-Border $results.Range('AD156:AT156')

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
    Add-WorkbookEventHandlers $workbook
    # Начальная структура нового шаблона создаётся сборщиком, а не рабочим VBA.
    $recipeTypes = @($excel.Run("'" + $workbook.Name + "'!modSectionRecipeSetup.RegisteredSectionRecipeTypes"))
    Get-ConfigValidationListAddress $system $recipeTypes 'rngSectionRecipeTypes' | Out-Null
    $excel.Run("'" + $workbook.Name + "'!modSectionRecipeSetup.UpdateSectionRecipeTypes")

    Apply-WorkbookLayout $workbook $script:WorkbookLayout
    $workbook.SaveAs($fullOutputPath, 52)
    Write-Output "Workbook built: $fullOutputPath"
}
catch {
    [ordered]@{Status='FAIL';Reason=$_.Exception.Message;ScriptStackTrace=$_.ScriptStackTrace} |
        ConvertTo-Json | Set-Content -LiteralPath (Join-Path $outputDir ('BuildFailed_'+(Get-Date -Format 'yyyyMMdd_HHmmss')+'.json')) -Encoding UTF8
    throw
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
























