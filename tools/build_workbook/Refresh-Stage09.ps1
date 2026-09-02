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
    $settings.Cells.Item($row, 7).Value2 = "���"
}

# Удаляет только служебный объект, который может мешать повторяемой сборке или проверке.
function Remove-SystemSetting {
    param(
        [object]$Workbook,
        [string]$Key
    )
    $settings = $Workbook.Names.Item("rngSystemSettings").RefersToRange
    for ($i = 2; $i -le $settings.Rows.Count; $i++) {
        if ([string]::Equals([string]$settings.Cells.Item($i, 1).Value2, $Key, [System.StringComparison]::OrdinalIgnoreCase)) {
            $settings.Worksheet.Rows.Item($settings.Row + $i - 1).Delete() | Out-Null
            return
        }
    }
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

# Устанавливает значение, оформление или именованный диапазон в книге через Excel COM.
function Set-CellText {
    param([object]$Sheet, [int]$Row, [int]$Column, [string]$Text)
    $Sheet.Cells.Item($Row, $Column).Value2 = $Text
}

# Устанавливает значение, оформление или именованный диапазон в книге через Excel COM.
function Set-ValidationList {
    param([object]$Cell, [string]$List)
    $Cell.Validation.Delete()
    $Cell.Validation.Add(3, 1, 1, $List)
}

# Добавляет структурный элемент книги или отчета, сохраняя единый формат сборочных скриптов.
function Add-Or-Replace-Button {
    param(
        [object]$Sheet,
        [string]$Name,
        [string]$Text,
        [string]$Macro,
        [double]$Left,
        [double]$Top,
        [int]$Color
    )

    foreach ($shape in @($Sheet.Shapes)) {
        if ($shape.Name -eq $Name) {
            $shape.Delete()
            break
        }
    }
    $button = $Sheet.Shapes.AddShape(1, $Left, $Top, 150, 28)
    $button.Name = $Name
    $button.TextFrame.Characters().Text = $Text
    $button.OnAction = $Macro
    $button.Fill.ForeColor.RGB = $Color
    $button.Line.ForeColor.RGB = $Color
    $button.TextFrame.Characters().Font.Color = 16777215
    $button.TextFrame.Characters().Font.Bold = $true
}

# Выполняет служебный шаг сборочного или проверочного сценария.
function Update-CalculationSheetLayout {
    param([object]$Workbook)

    $calc = $Workbook.Worksheets.Item(1)
    $mainLabels = @(
        @("������ / �������", "", "", ""),
        @("��� �������", "Circle", "", "RoundedRectangle ��� Circle"),
        @("������", "", "��", "�� ������������ ��� Circle"),
        @("������", "", "��", "�� ������������ ��� Circle"),
        @("������ ������� �����", "", "��", "�� ������������ ��� Circle"),
        @("������ ������� ������", "", "��", "�� ������������ ��� Circle"),
        @("������ ������ ������", "", "��", "�� ������������ ��� Circle"),
        @("������ ������ �����", "", "��", "�� ������������ ��� Circle"),
        @("������� ����� D", "300", "��", "������������ ��� Geometry.Type = Circle"),
        @("����� ����� X", "0", "��", "������������ ��� Geometry.Type = Circle"),
        @("����� ����� Y", "0", "��", "������������ ��� Geometry.Type = Circle"),
        @("����� ������", "", "", ""),
        @("����� ��������", "A400", "", ""),
        @("���������� �� ��� �������� as", "40", "��", "�� ����� ������� �� ��� �������"),
        @("���������� �������� n", "8", "��", "������������� ���������� �� ����������"),
        @("������� �������� ds", "20", "��", "���������� ��� ���� ���������� ��������"),
        @("���������� ������ ���������", "0.3", "��", "���������; ��������� ��������� ������ �� System")
    )
    for ($i = 0; $i -lt $mainLabels.Count; $i++) {
        $row = 6 + $i
        Set-CellText $calc $row 1 $mainLabels[$i][0]
        if ([string]::IsNullOrWhiteSpace([string]$calc.Cells.Item($row, 5).Value2) -and -not [string]::IsNullOrWhiteSpace($mainLabels[$i][1])) {
            Set-CellText $calc $row 5 $mainLabels[$i][1]
        }
        Set-CellText $calc $row 9 $mainLabels[$i][2]
        Set-CellText $calc $row 13 $mainLabels[$i][3]
    }
    Set-ValidationList $calc.Cells.Item(7, 5) "RoundedRectangle,Circle"
    $axisDistanceValue = 0.0
    if (-not [double]::TryParse([string]$calc.Cells.Item(19, 5).Value2, [ref]$axisDistanceValue) -or $axisDistanceValue -lt 1) { Set-CellText $calc 19 5 "40" }
    if ([string]::IsNullOrWhiteSpace([string]$calc.Cells.Item(20, 5).Value2)) { Set-CellText $calc 20 5 "8" }
    if ([string]::IsNullOrWhiteSpace([string]$calc.Cells.Item(21, 5).Value2)) { Set-CellText $calc 21 5 "20" }
    $crackWidthValue = 0.0
    if (-not [double]::TryParse([string]$calc.Cells.Item(22, 5).Value2, [ref]$crackWidthValue)) { Set-CellText $calc 22 5 "0.3" }
    Set-CellText $calc 23 1 "��������: ������������� ������������ ����������"
    $rebarHeaders = @("ID", "X", "Y", "Diameter", "Area", "SteelClass", "Comment")
    for ($i = 0; $i -lt $rebarHeaders.Count; $i++) {
        Set-CellText $calc 25 ($i + 1) $rebarHeaders[$i]
    }

    Set-CellText $calc 38 1 "��������� ��������"
    $loadHeaders = @("CombinationID", "N", "Mx", "My", "ProfileId", "Comment")
    for ($i = 0; $i -lt $loadHeaders.Count; $i++) {
        Set-CellText $calc 40 ($i + 1) $loadHeaders[$i]
    }

    $Workbook.Names.Item("rngMainInput").RefersTo = "=" + $calc.Range("A6:Q22").Address($true, $true, 1, $true)
    $Workbook.Names.Item("rngRebarInput").RefersTo = "=" + $calc.Range("A25:G34").Address($true, $true, 1, $true)
    $Workbook.Names.Item("rngLoadCombinations").RefersTo = "=" + $calc.Range("A40:F60").Address($true, $true, 1, $true)
    foreach ($obsoleteName in @("rngResultMx", "rngResultMy", "rngResultMxy")) {
        try { $Workbook.Names.Item($obsoleteName).Delete() } catch { }
    }
    $calc.PageSetup.PrintArea = $calc.Range("A1:BT60").Address($true, $true)

    $left = $calc.Cells.Item(4, 74).Left
    $top = $calc.Cells.Item(4, 74).Top
    Add-Or-Replace-Button $calc "btnRunSectionCalculation" "��������� ������" "RunSectionCalculation" $left $top 5383702
    Add-Or-Replace-Button $calc "btnClearAutoCADDrawing" "Очистить чертеж AutoCAD" "ClearAutoCADDrawing" $left ($top + 36) 8355711
    Add-Or-Replace-Button $calc "btnExportStressToAutoCAD" "��������� � AutoCAD" "ExportSectionStressToAutoCAD" $left ($top + 72) 10053171
}

$root = Resolve-Path -LiteralPath (Join-Path $PSScriptRoot "../..")
. (Join-Path $PSScriptRoot "SettingsCatalog.ps1")
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
        "src/Common/modGeometryTypes.bas",
        "src/Geometry/CFiberMeshBuilder.cls",
        "src/Geometry/CGeometryCircle.cls",
        "src/Geometry/CGeometryRoundedRectangle.cls",
        "src/Section/CCircleRebarLayoutBuilder.cls",
        "src/Section/CSectionPropertiesCalculator.cls",
        "src/Section/CRebarLayout.cls",
        "src/Materials/CMaterialDiagram.cls",        "src/Materials/CLinearConcreteMaterial.cls",
        "src/Materials/CLinearSteelMaterial.cls",
        "src/Solver/CLinearSystem3x3.cls",
        "src/Solver/CCapacitySolver.cls",
        "src/Solver/CSectionSolver.cls",
        "src/Crack/CCrackWidthCalculator.cls",
        "src/Batch/CBatchSectionCalculator.cls",
        "src/Excel/CSystemSettingsReader.cls",
        "src/Excel/CLoadCombinationReader.cls",
        "src/Excel/CBatchResultWriter.cls",
        "src/Excel/modWorkbookCalculation.bas",
        "src/Excel/modAutoCADStressExport.bas",
        "tests/modTestGeometry.bas",
        "tests/modTestMaterialDiagrams.bas",
        "tests/modTestSectionSolver.bas",
        "tests/modTestCapacitySolver.bas",
        "tests/modTestCrackWidth.bas",
        "tests/modTestBatchCalculation.bas",
        "tests/modTestWorkbookInterface.bas",
        "tests/modTestRegressionBaseline.bas"
    )
    foreach ($relative in $files) {
        Import-VbaSourceFile $workbook (Join-Path $root $relative)
    }

    Update-CalculationSheetLayout $workbook

    $systemSheet = $workbook.Worksheets.Item("System")
    Apply-SystemSettingsLayout $workbook $systemSheet | Out-Null
    Apply-CalculationSettingsLinks $workbook

    $workbook.Save()
    Write-Output "Workbook stage 9 refreshed: $fullWorkbookPath"
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












