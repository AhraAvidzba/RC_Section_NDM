function Get-SystemSettingsCatalog {
    @(
        @{ Name = "GeometrySettings"; Title = "[Геометрия и сетка]"; Rows = @(
            @("Geometry.Type", "Circle", "", "Тип сечения. Поддерживается: RoundedRectangle или Circle. Это единственный источник выбора геометрии."),
            @("Circle.Diameter", "300", "мм", "Диаметр круглого сечения. Используется только при Geometry.Type = Circle."),
            @("Circle.CenterX", "0", "мм", "Координата X центра круглого сечения."),
            @("Circle.CenterY", "0", "мм", "Координата Y центра круглого сечения."),
            @("Mesh.StepX", "20", "мм", "Базовый шаг бетонной волоконной сетки по X."),
            @("Mesh.StepY", "20", "мм", "Базовый шаг бетонной волоконной сетки по Y."),
            @("Mesh.BoundarySubdivisions", "1", "шт", "Количество подъячеек по каждой оси для граничной базовой ячейки; 1 - быстрый режим по центру ячейки.")
        )},
        @{ Name = "ConcreteDiagram"; Title = "[Бетон и диаграмма]"; Rows = @(
            @("Concrete.Class", "B30", "", "Класс бетона для текущего набора параметров."),
            @("Concrete.Rb.ULS", "15.5", "МПа", "Расчетное сопротивление бетона осевому сжатию."),
            @("Concrete.Rbt.ULS", "1.10", "МПа", "Расчетное сопротивление бетона осевому растяжению."),
            @("Concrete.Rb.SLS", "22.0", "МПа", "Сопротивление бетона сжатию для эксплуатационных расчетов."),
            @("Concrete.Rbt.SLS", "1.80", "МПа", "Сопротивление бетона растяжению для режима UseDiagram."),
            @("Concrete.Eb", "32500", "МПа", "Начальный модуль деформации бетона."),
            @("Concrete.TensionMode", "Ignore", "", "Работа растянутого бетона: Ignore - нулевое напряжение и касательная жесткость; UseDiagram - растянутая ветвь диаграммы."),
            @("Concrete.Point1.Eps", "-0.0015", "", "Первая пользовательская точка сжатой ветви диаграммы бетона: деформация."),
            @("Concrete.Point1.Stress", "-15.5", "МПа", "Первая пользовательская точка сжатой ветви диаграммы бетона: напряжение."),
            @("Concrete.Point2.Eps", "-0.002", "", "Вторая пользовательская точка сжатой ветви диаграммы бетона: деформация. По умолчанию лежит на площадке текущей двухлинейной формы."),
            @("Concrete.Point2.Stress", "-15.5", "МПа", "Вторая пользовательская точка сжатой ветви диаграммы бетона: напряжение."),
            @("Concrete.Point3.Eps", "-0.0035", "", "Третья пользовательская точка сжатой ветви диаграммы бетона: деформация. По умолчанию задает конец площадки текущей двухлинейной формы."),
            @("Concrete.Point3.Stress", "-15.5", "МПа", "Напряжение в третьей пользовательской точке бетона.")
        )},
        @{ Name = "SteelDiagram"; Title = "[Арматура и диаграмма]"; Rows = @(
            @("Steel.Class", "A400", "", "Класс обычной ненапрягаемой арматуры."),
            @("Steel.Rs.ULS", "350", "МПа", "Расчетное сопротивление арматуры растяжению."),
            @("Steel.Rsc.ULS", "350", "МПа", "Расчетное сопротивление арматуры сжатию."),
            @("Steel.Es", "200000", "МПа", "Модуль упругости обычной ненапрягаемой арматуры."),
            @("Steel.Point1.Eps", "0.00175", "", "Первая пользовательская точка растянутой ветви диаграммы арматуры: деформация."),
            @("Steel.Point1.Stress", "350", "МПа", "Первая пользовательская точка растянутой ветви диаграммы арматуры: напряжение."),
            @("Steel.Point2.Eps", "0.01", "", "Вторая пользовательская точка растянутой ветви диаграммы арматуры: деформация. По умолчанию лежит на площадке текущей двухлинейной формы."),
            @("Steel.Point2.Stress", "350", "МПа", "Вторая пользовательская точка растянутой ветви диаграммы арматуры: напряжение."),
            @("Steel.Point3.Eps", "0.025", "", "Третья пользовательская точка растянутой ветви диаграммы арматуры: деформация. По умолчанию задает конец площадки текущей двухлинейной формы."),
            @("Steel.Point3.Stress", "350", "МПа", "Напряжение в третьей пользовательской точке арматуры.")
        )},
        @{ Name = "SolverSettings"; Title = "[Решатель равновесия]"; Rows = @(
            @("Calculation.Mode", "DirectState", "", "Режим расчета: DirectState - НДС по заданным усилиям; FullCapacity - поиск lambdaUltimate; LinearMatrix - один линейно-упругий матричный расчет."),
            @("Solver.MaxIterations", "40", "шт", "Максимальное число итераций Ньютона на ступень нагрузки."),
            @("Solver.LoadSteps", "1", "шт", "Число ступеней приложения нагрузки в прямом НДМ."),
            @("Solver.ToleranceN", "1", "Н", "Абсолютный допуск равновесия по продольной силе."),
            @("Solver.ToleranceMx", "1000", "Н*мм", "Абсолютный допуск равновесия по моменту Mx."),
            @("Solver.ToleranceMy", "1000", "Н*мм", "Абсолютный допуск равновесия по моменту My."),
            @("Solver.LineSearchEnabled", "Yes", "", "Включить line search в CSectionSolver."),
            @("Solver.DampingInitial", "1", "", "Начальный коэффициент демпфирования шага Ньютона."),
            @("Solver.MinLineSearchAlpha", "0.03125", "", "Минимальный коэффициент alpha при line search."),
            @("Solver.MaxDeltaEpsilon0", "0.0005", "", "Ограничение приращения epsilon0 за одну итерацию."),
            @("Solver.MaxDeltaKappa", "0.00001", "1/мм", "Ограничение приращения кривизны за одну итерацию."),
            @("Solver.DiagnosticsEnabled", "No", "", "Записывать подробный журнал итераций решателя.")
        )},
        @{ Name = "CapacitySettings"; Title = "[Поиск несущей способности]"; Rows = @(
            @("Capacity.InitialLambda", "1", "", "Начальный множитель lambda для поиска верхней границы несущей способности."),
            @("Capacity.MaxLambda", "64", "", "Предельное значение lambda при расширении расчетной скобки."),
            @("Capacity.ToleranceLambda", "0.01", "", "Допуск одномерного поиска предельного множителя."),
            @("Capacity.MaxRetries", "0", "шт", "Число повторов после численной несходимости при поиске несущей способности."),
            @("Capacity.BaseLoadSteps", "1", "шт", "Базовое число ступеней нагрузки внутри CSectionSolver при режиме FullCapacity."),
            @("Capacity.SolverMaxIterations", "60", "шт", "Максимум итераций Ньютона на ступень при поиске несущей способности."),
            @("Capacity.ConcreteCompressionLimit", "-0.0035", "", "Предел деформации бетона для фиксации ConcreteStrainLimit."),
            @("Capacity.SteelStrainLimit", "0.025", "", "Предел деформации обычной ненапрягаемой арматуры для SteelStrainLimit.")
        )},
        @{ Name = "OutputSettings"; Title = "[Вывод и трещины]"; Rows = @(
            @("CrackWidth.Enabled", "Yes", "", "Включить расчет ширины раскрытия уже образовавшихся нормальных трещин."),
            @("CrackWidth.Allowable", "0.3", "мм", "Допустимая ширина раскрытия трещин для коэффициента использования."),
            @("CrackWidth.CrackSpacing", "200", "мм", "Временный параметр расстояния между трещинами до окончательной нормативной трассировки."),
            @("CrackWidth.StrainFactor", "1", "", "Временный коэффициент для прозрачной формулы ширины раскрытия на листе."),
            @("CrackWidth.DurationFactor", "1", "", "Временный коэффициент длительности действия нагрузки.")
        )},
        @{ Name = "AutoCADSettings"; Title = "[AutoCAD]"; Rows = @()}
    )
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

function Set-WorkbookNameByBounds {
    param([object]$Workbook, [string]$Name, [object]$Sheet, [int]$FirstRow, [int]$FirstColumn, [int]$LastRow, [int]$LastColumn)
    $sheetName = ([string]$Sheet.Name).Replace("'", "''")
    $firstCol = ConvertTo-ExcelColumn $FirstColumn
    $lastCol = ConvertTo-ExcelColumn $LastColumn
    $rangeAddress = "$" + $firstCol + "$" + $FirstRow + ":$" + $lastCol + "$" + $LastRow
    $address = "='" + $sheetName + "'!" + $rangeAddress
    try { $Workbook.Names.Item($Name).RefersTo = $address }
    catch { $Workbook.Names.Add($Name, $address) | Out-Null }
}

function Apply-SystemSettingsLayout {
    param([object]$Workbook, [object]$Sheet)

    $catalog = Get-SystemSettingsCatalog
    $Sheet.Range("A1:O260").ClearContents()
    $Sheet.Range("A1:O260").Validation.Delete()

    $Sheet.Cells.Item(1, 1).Value2 = "System"
    $Sheet.Cells.Item(1, 1).Font.Bold = $true
    $Sheet.Cells.Item(1, 1).Font.Size = 16

    $headers = @("Параметр", "Значение", "Ед.", "Комментарий")
    for ($i = 0; $i -lt $headers.Count; $i++) {
        $cell = $Sheet.Cells.Item(3, $i + 1)
        $cell.Value2 = $headers[$i]
        $cell.Font.Bold = $true
        $cell.Interior.Color = 14277081
    }

    $row = 4
    foreach ($section in $catalog) {
        $sectionStart = $row
        $Sheet.Cells.Item($row, 1).Value2 = $section.Title
        $sectionRange = $Sheet.Range($Sheet.Cells.Item($row, 1), $Sheet.Cells.Item($row, 4))
        $sectionRange.Font.Bold = $true
        $sectionRange.Interior.Color = 15921906
        $row++
        foreach ($setting in $section.Rows) {
            for ($c = 0; $c -lt 4; $c++) { $Sheet.Cells.Item($row, $c + 1).Value2 = $setting[$c] }
            $row++
        }
        Set-WorkbookNameByBounds $Workbook $section.Name $Sheet $sectionStart 1 ($row - 1) 4
        $row++
    }

    $settingsRange = $Sheet.Range($Sheet.Cells.Item(3, 1), $Sheet.Cells.Item($row - 2, 4))
    Set-WorkbookNameByBounds $Workbook "rngSystemSettings" $Sheet 3 1 ($row - 2) 4

    $validationLists = [ordered]@{
        "Geometry.Type" = @("RoundedRectangle", "Circle")
        "Concrete.TensionMode" = @("Ignore", "UseDiagram")
        "Calculation.Mode" = @("DirectState", "FullCapacity", "LinearMatrix")
        "Solver.LineSearchEnabled" = @("Yes", "No")
        "Solver.DiagnosticsEnabled" = @("Yes", "No")
        "CrackWidth.Enabled" = @("Yes", "No")
    }

    $listColumn = 8
    foreach ($key in $validationLists.Keys) {
        $options = $validationLists[$key]
        for ($i = 0; $i -lt $options.Count; $i++) {
            $Sheet.Cells.Item($i + 1, $listColumn).Value2 = $options[$i]
        }
        $colName = ConvertTo-ExcelColumn $listColumn
        for ($r = 4; $r -le ($row - 2); $r++) {
            if ([string]$Sheet.Cells.Item($r, 1).Value2 -eq $key) {
                $cell = $Sheet.Cells.Item($r, 2)
                $cell.Validation.Delete()
                $listAddress = "=$" + $colName + '$1:$' + $colName + '$' + $options.Count
                $cell.Validation.Add(3, 1, 1, $listAddress)
            }
        }
        $listColumn++
    }

    $diagStart = $row + 2
    $Sheet.Cells.Item($diagStart, 1).Value2 = "Диагностика"
    $Sheet.Cells.Item($diagStart, 1).Font.Bold = $true
    $Sheet.Cells.Item($diagStart, 1).Interior.Color = 8355711
    $Sheet.Cells.Item($diagStart, 1).Font.Color = 16777215

    $diagHeaders = @("Iteration", "LoadFactor", "Epsilon0", "KappaX", "KappaY", "Nint", "Mxint", "Myint", "ResidualN", "ResidualMx", "ResidualMy", "StepAlpha", "MatrixDeterminant", "MatrixConditionEstimate", "StopReason")
    for ($i = 0; $i -lt $diagHeaders.Count; $i++) {
        $cell = $Sheet.Cells.Item($diagStart + 2, $i + 1)
        $cell.Value2 = $diagHeaders[$i]
        $cell.Font.Bold = $true
        $cell.Interior.Color = 14277081
    }
    $diagnosticRange = $Sheet.Range($Sheet.Cells.Item($diagStart + 2, 1), $Sheet.Cells.Item($diagStart + 12, $diagHeaders.Count))
    Set-WorkbookNameByBounds $Workbook "rngSystemDiagnostics" $Sheet ($diagStart + 2) 1 ($diagStart + 12) $diagHeaders.Count

    $Sheet.Columns.Item(1).ColumnWidth = 34
    $Sheet.Columns.Item(2).ColumnWidth = 18
    $Sheet.Columns.Item(3).ColumnWidth = 12
    $Sheet.Columns.Item(4).ColumnWidth = 86
    $Sheet.Columns.Item("H:M").Hidden = $true
    $Sheet.Range("A1:O260").Font.Name = "Arial"
    $Sheet.Range("A1:O260").Font.Size = 9

    return @{ Settings = $settingsRange; Diagnostics = $diagnosticRange }
}

function Apply-CalculationSettingsLinks {
    param([object]$Workbook)
    $calc = $Workbook.Worksheets.Item(1)
    $calc.Cells.Item(7, 5).Formula = '=VLOOKUP("Geometry.Type",System!$A:$B,2,FALSE)'
    $calc.Cells.Item(14, 5).Formula = '=VLOOKUP("Circle.Diameter",System!$A:$B,2,FALSE)'
    $calc.Cells.Item(15, 5).Formula = '=VLOOKUP("Circle.CenterX",System!$A:$B,2,FALSE)'
    $calc.Cells.Item(16, 5).Formula = '=VLOOKUP("Circle.CenterY",System!$A:$B,2,FALSE)'
}

function Add-SystemSettings {
    param([object]$Sheet)
    $result = Apply-SystemSettingsLayout $Sheet.Parent $Sheet
    Apply-CalculationSettingsLinks $Sheet.Parent
    $result
}
