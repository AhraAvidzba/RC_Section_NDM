function Get-SystemSettingsCatalog {
    @(
        @{ Name = "GeometrySettings"; Title = "[Геометрия и сетка]"; Rows = @(
            @("Geometry.Source", "Generated", "", "Источник расчетной геометрии: Generated - построить сетку встроенными генераторами; AutoCAD - импортировать только Region из активного чертежа AutoCAD. Единицы AutoCAD всегда считаются мм."),
            @("Geometry.Type", "LShape", "", "Тип сечения для Geometry.Source = Generated. Поддерживается: RoundedRectangle, Circle или LShape."),
            @("Mesh.Step", "50", "мм", "Базовый шаг квадратной бетонной волоконной сетки для Geometry.Source = Generated. Один и тот же шаг используется по X и Y."),
            @("Mesh.BoundarySubdivisions", "1", "шт", "Количество подъячеек по каждой оси для граничной базовой ячейки при Geometry.Source = Generated; 1 - быстрый режим по центру ячейки."),
            @("Load.ReferenceOffsetX", "0", "мм", "Смещение точки приложения нагрузок по X относительно центра тяжести приведенного сечения. 0 означает, что Mx/My заданы относительно этого центра."),
            @("Load.ReferenceOffsetY", "0", "мм", "Смещение точки приложения нагрузок по Y относительно центра тяжести приведенного сечения. Рабочие моменты для решателя автоматически переносятся к текущим координатам волокон.")
        )},
        @{ Name = "ConcreteDiagram"; Title = "[Бетон и диаграмма]"; Rows = @(
            @("Concrete.Class", "B30", "", "Класс бетона для текущего набора параметров."),
            @("Concrete.Rb.ULS", "15.5", "МПа", "Расчетное сопротивление бетона осевому сжатию."),
            @("Concrete.Rbt.ULS", "1.10", "МПа", "Расчетное сопротивление бетона осевому растяжению."),
            @("Concrete.Rb.SLS", "22.0", "МПа", "Сопротивление бетона сжатию для эксплуатационных расчетов."),
            @("Concrete.Rbt.SLS", "1.80", "МПа", "Сопротивление бетона растяжению для режима UseDiagram."),
            @("Concrete.Eb", "32500", "МПа", "Начальный модуль деформации бетона."),
            @("Concrete.TensionMode", "Ignore", "", "Работа растянутого бетона: Ignore - нулевое напряжение и касательная жесткость; UseDiagram - использовать точки диаграммы бетона в растянутой зоне."),
            @("Capacity.ConcreteCompressionLimit", "-0.0035", "", "Предельная сжимающая деформация бетона для фиксации ConcreteStrainLimit при поиске несущей способности."),
            @("Capacity.ConcreteTensionLimit", "0.00015", "", "Предельная растягивающая деформация бетона для ConcreteTensionStrainLimit; используется только при Concrete.TensionMode = UseDiagram.")
        )},
        @{ Name = "SteelDiagram"; Title = "[Арматура и диаграмма]"; Rows = @(
            @("Steel.Class", "A400", "", "Класс обычной ненапрягаемой арматуры."),
            @("Steel.Rs.ULS", "350", "МПа", "Расчетное сопротивление арматуры растяжению."),
            @("Steel.Rsc.ULS", "350", "МПа", "Расчетное сопротивление арматуры сжатию."),
            @("Steel.Es", "200000", "МПа", "Модуль упругости обычной ненапрягаемой арматуры."),
            @("Capacity.SteelStrainLimit", "0.025", "", "Предельная деформация обычной ненапрягаемой арматуры для фиксации SteelStrainLimit при поиске несущей способности.")
        )},
        @{ Name = "SolverSettings"; Title = "[Решатель равновесия]"; Rows = @(
            @("Calculation.Mode", "FullCapacity", "", "Режим расчета: DirectState - НДС по заданным усилиям; FullCapacity - поиск lambdaUltimate."),
            @("Solver.Method", "Newton", "", "Метод решения системы равновесия CSectionSolver. Допустимо только Newton или Secant; по умолчанию Newton; пустое или иное значение дает InputError без скрытого fallback."),
            @("Solver.MaxIterations", "40", "шт", "Применяется к Newton и Secant. Максимальное число итераций на ступень нагрузки; увеличение повышает шанс сходимости, но увеличивает время."),
            @("Solver.LoadSteps", "1", "шт", "Применяется к Newton и Secant. Число ступеней приложения нагрузки в прямом НДМ; больше ступеней обычно устойчивее, но медленнее."),
            @("Solver.ToleranceN", "1", "Н", "Применяется к Newton и Secant. Абсолютный допуск равновесия по продольной силе; меньше значение строже и может увеличить число итераций."),
            @("Solver.ToleranceMx", "1000", "Н*мм", "Применяется к Newton и Secant. Абсолютный допуск равновесия по моменту Mx; меньше значение строже и может увеличить число итераций."),
            @("Solver.ToleranceMy", "1000", "Н*мм", "Применяется к Newton и Secant. Абсолютный допуск равновесия по моменту My; меньше значение строже и может увеличить число итераций."),
            @("Solver.LineSearchEnabled", "Yes", "", "Применяется к Newton и Secant. Включает line search; повышает устойчивость шага, но может добавить вычисления внутренних усилий."),
            @("Solver.DampingInitial", "1", "", "Применяется к Newton и Secant. Начальный коэффициент демпфирования шага; меньше значение делает шаг осторожнее."),
            @("Solver.MinLineSearchAlpha", "0.03125", "", "Применяется к Newton и Secant. Минимальный коэффициент alpha при line search; меньше значение разрешает более сильное дробление шага."),
            @("Solver.MaxDeltaEpsilon0", "0.0005", "", "Применяется к Newton и Secant. Ограничение приращения epsilon0 за одну итерацию; меньше значение устойчивее, но медленнее."),
            @("Solver.MaxDeltaKappa", "0.00001", "1/мм", "Применяется к Newton и Secant. Ограничение приращения кривизны за одну итерацию; меньше значение устойчивее, но медленнее."),
            @("Solver.SecantMaxRestarts", "2", "шт", "Применяется только к Secant. Максимальное число контролируемых перезапусков приближенной матрицы чувствительности."),
            @("Solver.SecantMinStepNorm", "0.000000000001", "", "Применяется только к Secant. Минимально допустимая норма secant-шага; слишком малый шаг считается застоем.")
        )},
        @{ Name = "CapacitySettings"; Title = "[Поиск несущей способности]"; Rows = @(
            @("Capacity.Method", "UltimateStrain", "", "Актуально для всех расчетов несущей способности. Метод определения: LoadMultiplier - масштабирование заданного вектора моментов; UltimateStrain - прямой поиск состояния по достижению предельной деформации при сохранении направления Mx/My."),
            @("Capacity.SearchMethod", "Bisection", "", "Метод поиска несущей способности для Capacity.Method = LoadMultiplier. Поддерживается только Bisection, Brent или Secant. Пустое или другое значение дает InputError."),
            @("Capacity.InitialLambda", "1", "", "Актуально только для Capacity.Method = LoadMultiplier. Начальный множитель lambda при поиске верхней границы несущей способности."),
            @("Capacity.ToleranceLambda", "0.01", "", "Актуально только для Capacity.Method = LoadMultiplier. Общий допуск одномерного поиска lambdaUltimate для Bisection, Brent и Secant."),
            @("Capacity.MaxRetries", "0", "шт", "Актуально только для Capacity.Method = LoadMultiplier. Число повторов после численной несходимости пробного расчета."),
            @("Capacity.BaseLoadSteps", "1", "шт", "Актуально только для Capacity.Method = LoadMultiplier. Базовое число ступеней нагрузки внутри CSectionSolver для каждого пробного lambda."),
            @("Capacity.ToleranceStrain", "0.00001", "", "Актуально только для Capacity.Method = UltimateStrain. Абсолютный допуск достижения предельной деформации критического бетонного волокна или стержня."),
            @("Capacity.MaxLambda", "64", "", "Актуально для обоих методов. В LoadMultiplier задает предел расширения расчетной скобки; в UltimateStrain используется как защитный верхний предел найденного lambda."),
            @("Capacity.SolverMaxIterations", "60", "шт", "Актуально для обоих методов. В LoadMultiplier задает максимум итераций Ньютона на ступень внутреннего решателя; в UltimateStrain задает максимум итераций прямого поиска предельного состояния.")
        )},
        @{ Name = "OutputSettings"; Title = "[Вывод и трещины]"; Rows = @(
            @("CrackWidth.Enabled", "Yes", "", "Включить расчет ширины раскрытия уже образовавшихся нормальных трещин."),
            @("CrackWidth.Allowable", "0.3", "мм", "Допустимая ширина раскрытия трещин для коэффициента использования."),
            @("CrackWidth.CrackSpacing", "200", "мм", "Временный параметр расстояния между трещинами до окончательной нормативной трассировки."),
            @("CrackWidth.StrainFactor", "1", "", "Временный коэффициент для прозрачной формулы ширины раскрытия на листе."),
            @("CrackWidth.DurationFactor", "1", "", "Временный коэффициент длительности действия нагрузки.")
        )},
        @{ Name = "AutoCADExportSettings"; Title = "[AutoCAD export]"; Rows = @(
            @("AutoCAD.Export.CombinationID", "Worst", "", "Какое сочетание экспортировать в AutoCAD: Worst - определяющее сочетание из последнего расчета; либо конкретный CombinationID из rngLoadCombinations."),
            @("AutoCAD.Export.NeutralLineEnabled", "Yes", "", "Выгружать нейтральную линию в AutoCAD: Yes - выводить; No - не выводить."),
            @("AutoCAD.Export.PrincipalAxesEnabled", "Yes", "", "Выгружать главные центральные оси приведенного сечения в AutoCAD: Yes - выводить; No - не выводить."),
            @("AutoCAD.Export.LoadPointEnabled", "Yes", "", "Выгружать точку приложения нагрузки в AutoCAD: Yes - выводить; No - не выводить."),
            @("AutoCAD.Export.ResultType", "Stress", "", "Что выводить цветом и подписями в AutoCAD: Stress - напряжения; Strain - деформации. Используются данные последнего расчетного snapshot."),
            @("AutoCAD.Export.LabelMode", "NamesAndValues", "", "Формат текстовых подписей при выгрузке в AutoCAD: ValuesOnly - только значение выбранного ResultType; NamesAndValues - имя элемента и значение. При NamesAndValues имена выводятся для бетона и арматуры."),
            @("AutoCAD.Layer.Concrete", "Concrete", "", "Слой для областей бетонных волокон. Цвет каждой области задается по знаку напряжения: растяжение, сжатие или нейтральное состояние."),
            @("AutoCAD.Layer.Rebar", "Reinf", "", "Слой для всех областей продольной арматуры. Цвет каждого стержня задается по знаку напряжения: растяжение, сжатие или нейтральное состояние."),
            @("AutoCAD.Layer.ConcreteTension", "Anno_Concrete_Positive", "", "Слой для подписей положительных напряжений бетона, то есть растянутых бетонных волокон."),
            @("AutoCAD.Layer.ConcreteCompression", "Anno_Concrete_Negative", "", "Слой для подписей отрицательных, нулевых и почти нулевых напряжений бетона."),
            @("AutoCAD.Layer.RebarTension", "Anno_Rebar_Positive", "", "Слой для подписей положительных напряжений арматуры, то есть растянутых стержней."),
            @("AutoCAD.Layer.RebarCompression", "Anno_Rebar_Negative", "", "Слой для подписей отрицательных, нулевых и почти нулевых напряжений арматуры."),
            @("AutoCAD.Color.ConcreteTension", "9", "", "AutoCAD ColorIndex для растянутых бетонных областей и их подписей."),
            @("AutoCAD.Color.ConcreteCompression", "5", "", "AutoCAD ColorIndex для сжатых бетонных областей и их подписей."),
            @("AutoCAD.Color.RebarTension", "1", "", "AutoCAD ColorIndex для растянутых стержней и их подписей."),
            @("AutoCAD.Color.RebarCompression", "6", "", "AutoCAD ColorIndex для сжатых стержней и их подписей."),
            @("AutoCAD.Color.Neutral", "8", "", "AutoCAD ColorIndex для областей и подписей с нулевыми или почти нулевыми напряжениями бетона и арматуры. Слой подписи при этом остается compression-слоем соответствующего материала.")
        )},
        @{ Name = "AutoCADImportSettings"; Title = "[AutoCAD import]"; Rows = @(
            @("AutoCAD.Import.ConcreteLayer", "Concrete", "", "Слой бетонных областей для импорта. Импортируются только AutoCAD Region на этом слое; единицы чертежа считаются мм."),
            @("AutoCAD.Import.RebarLayer", "Reinf", "", "Слой областей арматуры для импорта. Импортируются только AutoCAD Region на этом слое; Steel.Class берется из System, разные классы арматуры в одном сечении не поддерживаются."),
            @("AutoCAD.Import.MinArea", "0.000001", "мм2", "Минимальная площадь Region для импорта; области с меньшей площадью игнорируются.")
        )},
        @{ Name = "PlotSettings"; Title = "[Схема сечения]"; Rows = @(
            @("Plot.Enabled", "Yes", "", "Включить построение схемы сечения на листе Расчет по последнему расчетному snapshot на листе Results."),
            @("Plot.AutoUpdateAfterCalculation", "Yes", "", "Автоматически обновлять схему после выполнения расчета. Кнопка Обновить схему всегда читает только Results и не запускает расчет."),
            @("Plot.LoadCase", "Worst", "", "Какое сочетание показывать на схеме: Worst - определяющее сочетание из последнего расчета; либо конкретный CombinationID из rngLoadCombinations."),
            @("Plot.ResultType", "Stress", "", "Что показывать цветом и численными подписями: Stress - напряжения; Strain - деформации. Используется сохраненный snapshot, а не текущие единицы System."),
            @("Plot.ResultGradient", "Yes", "", "Включить цветовое различение результата на схеме. Цвет физического состояния берется из PhysicalState snapshot."),
            @("Plot.ResultLabelsEnabled", "No", "", "Показывать пространственно распределенные численные подписи выбранного ResultType для бетонных элементов."),
            @("Plot.ResultLabelSpacing", "100", "мм", "Минимальный пространственный шаг между численными подписями результата на схеме."),
            @("Plot.ResultPrecision", "1", "шт", "Количество знаков после запятой для численных подписей и легенды схемы."),
            @("Plot.NeutralLineEnabled", "Yes", "", "Показывать нейтральную линию выбранного сочетания по Epsilon0, KappaX, KappaY из Results."),
            @("Plot.PrincipalAxesEnabled", "Yes", "", "Показывать главные центральные оси приведенного сечения."),
            @("Plot.LoadApplicationPointEnabled", "Yes", "", "Показывать точку приложения нагрузки из последнего расчетного snapshot."),
            @("Plot.CentroidEnabled", "Yes", "", "Показывать центр тяжести приведенного сечения."),
            @("Plot.LegendEnabled", "Yes", "", "Показывать легенду физического состояния и выбранного ResultType справа от схемы."),
            @("Plot.LegendMode", "Separate", "", "Separate - отдельные легенды для арматуры и бетона; Common - одна общая легенда. В режиме Common используются цвета Plot.Color.RebarCompression и Plot.Color.RebarTension."),
            @("Plot.Color.RebarCompression", "30,80,220", "RGB", "Цвет максимального сжатия арматуры. В режиме Plot.LegendMode=Common используется как цвет сжатия для всех элементов."),
            @("Plot.Color.RebarTension", "210,30,20", "RGB", "Цвет максимального растяжения арматуры. В режиме Plot.LegendMode=Common используется как цвет растяжения для всех элементов."),
            @("Plot.Color.ConcreteCompression", "125,35,210", "RGB", "Цвет максимального сжатия бетона при Plot.LegendMode=Separate."),
            @("Plot.Color.ConcreteTension", "0,155,85", "RGB", "Цвет максимального растяжения бетона при Plot.LegendMode=Separate; неработающий растянутый бетон остается серым.")
        )}
    )
}

function Get-UnitSettingsCatalog {
    @(
        @("Length", "mm", "mm", "mm"),
        @("Area", "mm2", "mm2", "mm2"),
        @("Force", "tf", "N", "tf"),
        @("Moment", "tf*m", "N*mm", "tf*m"),
        @("Stress", "MPa", "MPa", "MPa"),
        @("Curvature", "1/mm", "1/mm", "1/mm")
    )
}

function Get-SignConventionSettingsCatalog {
    @(
        @("+N", "Compression", "Tension"),
        @("+Mx", "+Y tension", "+Y tension"),
        @("+My", "+X tension", "+X tension")
    )
}

function Get-GeometrySettingsCatalog {
    @(
        @{ RangeName = "rngCircleGeometry"; Title = "Круглое сечение"; StartRow = 16; StartColumn = 6; Rows = @(
            @("Circle.Diameter", "300", "мм", "Диаметр круглого сечения."),
            @("Rebar.AxisDistance", "40", "мм", "Расстояние от грани круга до оси стержней as."),
            @("Rebar.Count", "8", "шт", "Количество продольных стержней по окружности."),
            @("Rebar.Diameter", "20", "мм", "Диаметр продольных стержней первого ряда."),
            @("дополнительные ряды арматуры", "", "", "Параметры дополнительных рядов. Пустой или нулевой диаметр означает, что ряд не создается."),
            @("Rebar.Diameter2", "", "мм", "Диаметр стержней второго ряда. Второй ряд строится от стержней первого ряда."),
            @("Rebar.Diameter3", "", "мм", "Диаметр стержней третьего ряда. Третий ряд строится от первого ряда; при совпадении Rebar.Loc3row с Rebar.Loc2row и наличии второго ряда перескакивает второй ряд."),
            @("Rebar.Loc2row", "Stacked", "", "Расположение второго ряда: Stacked - внутрь сечения по радиусу к центру; SideBySide - справа по часовой касательной."),
            @("Rebar.Loc3row", "Stacked", "", "Расположение третьего ряда: Stacked - внутрь сечения по радиусу к центру; SideBySide - справа по часовой касательной.")
        )},
        @{ RangeName = "rngRoundedRectangleGeometry"; Title = "Скругленный прямоугольник"; StartRow = 16; StartColumn = 10; Rows = @(
            @("RoundedRectangle.Width", "300", "мм", "Ширина прямоугольного сечения со скруглениями."),
            @("RoundedRectangle.Height", "200", "мм", "Высота прямоугольного сечения со скруглениями."),
            @("RoundedRectangle.RadiusTopLeft", "0", "мм", "Радиус верхнего левого угла."),
            @("RoundedRectangle.RadiusTopRight", "0", "мм", "Радиус верхнего правого угла."),
            @("RoundedRectangle.RadiusBottomRight", "0", "мм", "Радиус нижнего правого угла."),
            @("RoundedRectangle.RadiusBottomLeft", "0", "мм", "Радиус нижнего левого угла.")
        )},
        @{ RangeName = "rngLShapeGeometry"; Title = "Г-образное сечение"; StartRow = 30; StartColumn = 6; FaceTable = $true; Rows = @(
            @("величина размера", "550", "250", "250", "600", "мм", "H1 и B1 - высота и ширина верхнего прямоугольника; H2 и B2 - высота и ширина нижнего прямоугольника; полная высота сечения равна H1 + H2."),
            @("as_1", "40", "40", "40", "40", "мм", "Отступ от грани _1. Нумерация граней идет слева направо для H и сверху вниз для B: _1 - левая/верхняя грань."),
            @("as_2", "40", "40", "40", "40", "мм", "Отступ от грани _2. Нумерация граней идет слева направо для H и сверху вниз для B: _2 - правая/нижняя грань."),
            @("d_1", "32", "32", "32", "32", "мм", "Диаметр стержней у грани _1; если n_1 = 0, значение не используется."),
            @("d_2", "32", "32", "32", "32", "мм", "Диаметр стержней у грани _2; если n_2 = 0, значение не используется."),
            @("n_1", "5", "2", "2", "5", "шт", "Количество стержней у грани _1; 0 означает, что арматура у этой грани не создается."),
            @("n_2", "5", "2", "2", "5", "шт", "Количество стержней у грани _2; 0 означает, что арматура у этой грани не создается."),
            @("t1_1", "80", "80", "80", "80", "мм", "Отступ первого стержня от начала грани _1 при обходе по часовой стрелке."),
            @("t2_1", "80", "80", "80", "80", "мм", "Отступ последнего стержня от конца грани _1 при обходе по часовой стрелке."),
            @("t1_2", "80", "80", "80", "80", "мм", "Отступ первого стержня от начала грани _2 при обходе по часовой стрелке."),
            @("t2_2", "80", "80", "80", "80", "мм", "Отступ последнего стержня от конца грани _2 при обходе по часовой стрелке."),
            @("дополнительные ряды арматуры", "", "", "", "", "", "Параметры дополнительных рядов. Пустой диаметр означает, что ряд не создается. Суффикс _1 - левая грань для H или верхняя грань для B; _2 - правая грань для H или нижняя грань для B."),
            @("d_2row_1", "", "", "", "", "мм", "Диаметр стержней второго ряда у грани _1. Если первый ряд этой грани не задан, второй ряд не создается."),
            @("d_2row_2", "", "", "", "", "мм", "Диаметр стержней второго ряда у грани _2. Если первый ряд этой грани не задан, второй ряд не создается."),
            @("d_3row_1", "", "", "", "", "мм", "Диаметр стержней третьего ряда у грани _1. Ставится относительно первого ряда; при совпадении loc_3row с loc_2row перескакивает второй ряд."),
            @("d_3row_2", "", "", "", "", "мм", "Диаметр стержней третьего ряда у грани _2. Ставится относительно первого ряда; при совпадении loc_3row с loc_2row перескакивает второй ряд."),
            @("loc_2row", "Stacked", "Stacked", "Stacked", "Stacked", "", "Расположение второго ряда, общее для граней _1/_2: Stacked - внутрь сечения; SideBySide - для H вниз, для B вправо."),
            @("loc_3row", "Stacked", "Stacked", "Stacked", "Stacked", "", "Расположение третьего ряда, общее для граней _1/_2: Stacked - внутрь сечения; SideBySide - для H вниз, для B вправо."),
            @("bind_2row", "EachBar", "EachBar", "EachBar", "EachBar", "", "Привязка второго ряда: EachBar - к каждому стержню первого ряда; EverySecondBar - через один стержень первого ряда."),
            @("bind_3row", "EachBar", "EachBar", "EachBar", "EachBar", "", "Привязка третьего ряда: EachBar - к каждому стержню первого ряда; EverySecondBar - через один стержень первого ряда.")
        )}
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

function Add-UnitSettingsTable {
    param([object]$Workbook, [object]$Sheet, [int]$HeaderRow, [int]$StartColumn)

    $Sheet.Cells.Item($HeaderRow - 1, $StartColumn).Value2 = "Units"
    $Sheet.Range($Sheet.Cells.Item($HeaderRow - 1, $StartColumn), $Sheet.Cells.Item($HeaderRow - 1, $StartColumn + 3)).Merge() | Out-Null
    $Sheet.Cells.Item($HeaderRow - 1, $StartColumn).Font.Bold = $true
    $Sheet.Cells.Item($HeaderRow - 1, $StartColumn).Interior.Color = 15921906

    $headers = @("Quantity", "INPUT", "INTERNAL", "OUTPUT")
    for ($i = 0; $i -lt $headers.Count; $i++) {
        $cell = $Sheet.Cells.Item($HeaderRow, $StartColumn + $i)
        $cell.Value2 = $headers[$i]
        $cell.Font.Bold = $true
        $cell.Interior.Color = 14277081
    }

    $rows = Get-UnitSettingsCatalog
    for ($r = 0; $r -lt $rows.Count; $r++) {
        for ($c = 0; $c -lt 4; $c++) {
            $Sheet.Cells.Item($HeaderRow + 1 + $r, $StartColumn + $c).Value2 = $rows[$r][$c]
        }
        $Sheet.Cells.Item($HeaderRow + 1 + $r, $StartColumn + 2).Interior.Color = 15921906
    }

    $validationLists = @{
        "Length" = @("mm", "cm", "m")
        "Area" = @("mm2", "cm2", "m2")
        "Force" = @("N", "kN", "tf")
        "Moment" = @("N*mm", "kN*m", "tf*m")
        "Stress" = @("Pa", "kPa", "MPa", "kgf/cm2", "tf/m2")
        "Curvature" = @("1/mm", "1/m")
    }
    $listColumn = 70
    for ($r = 0; $r -lt $rows.Count; $r++) {
        $quantity = [string]$rows[$r][0]
        $options = $validationLists[$quantity]
        for ($i = 0; $i -lt $options.Count; $i++) {
            $Sheet.Cells.Item($i + 1, $listColumn).Value2 = $options[$i]
        }
        $colName = ConvertTo-ExcelColumn $listColumn
        $listAddress = "=$" + $colName + '$1:$' + $colName + '$' + $options.Count
        foreach ($valueColumn in @($($StartColumn + 1), $($StartColumn + 3))) {
            $cell = $Sheet.Cells.Item($HeaderRow + 1 + $r, $valueColumn)
            $cell.Validation.Delete()
            $cell.Validation.Add(3, 1, 1, $listAddress)
            $cell.Validation.IgnoreBlank = $false
            $cell.Validation.InCellDropdown = $true
        }
        $listColumn++
    }

    Set-WorkbookNameByBounds $Workbook "rngUnitSettings" $Sheet $HeaderRow $StartColumn ($HeaderRow + $rows.Count) ($StartColumn + 3)
}

function Add-SignConventionSettingsTable {
    param([object]$Workbook, [object]$Sheet, [int]$HeaderRow, [int]$StartColumn)

    $Sheet.Cells.Item($HeaderRow - 1, $StartColumn).Value2 = "Sign convention"
    $Sheet.Range($Sheet.Cells.Item($HeaderRow - 1, $StartColumn), $Sheet.Cells.Item($HeaderRow - 1, $StartColumn + 2)).Merge() | Out-Null
    $Sheet.Cells.Item($HeaderRow - 1, $StartColumn).Font.Bold = $true
    $Sheet.Cells.Item($HeaderRow - 1, $StartColumn).Interior.Color = 15921906

    $headers = @("Quantity", "USER", "INTERNAL")
    for ($i = 0; $i -lt $headers.Count; $i++) {
        $cell = $Sheet.Cells.Item($HeaderRow, $StartColumn + $i)
        $cell.Value2 = $headers[$i]
        $cell.Font.Bold = $true
        $cell.Interior.Color = 14277081
    }

    $rows = Get-SignConventionSettingsCatalog
    for ($r = 0; $r -lt $rows.Count; $r++) {
        for ($c = 0; $c -lt 3; $c++) {
            $Sheet.Cells.Item($HeaderRow + 1 + $r, $StartColumn + $c).Value2 = $rows[$r][$c]
        }
        $Sheet.Cells.Item($HeaderRow + 1 + $r, $StartColumn + 2).Interior.Color = 15921906
    }

    $validationLists = @{
        "+N" = @("Tension", "Compression")
        "+Mx" = @("+Y tension", "-Y tension")
        "+My" = @("+X tension", "-X tension")
    }
    $listColumn = 76
    for ($r = 0; $r -lt $rows.Count; $r++) {
        $quantity = [string]$rows[$r][0]
        $options = $validationLists[$quantity]
        for ($i = 0; $i -lt $options.Count; $i++) {
            $Sheet.Cells.Item($i + 1, $listColumn).Value2 = $options[$i]
        }
        $colName = ConvertTo-ExcelColumn $listColumn
        $listAddress = "=$" + $colName + '$1:$' + $colName + '$' + $options.Count
        $cell = $Sheet.Cells.Item($HeaderRow + 1 + $r, $StartColumn + 1)
        $cell.Validation.Delete()
        $cell.Validation.Add(3, 1, 1, $listAddress)
        $cell.Validation.IgnoreBlank = $false
        $cell.Validation.InCellDropdown = $true
        $listColumn++
    }

    Set-WorkbookNameByBounds $Workbook "rngSignConventionSettings" $Sheet $HeaderRow $StartColumn ($HeaderRow + $rows.Count) ($StartColumn + 2)
}

function Get-PlotAnnotationSettingsCatalog {
    @(
        @("Enabled", "Yes", "Yes", "Yes/No", "Включает вывод соответствующего типа аннотаций на схеме."),
        @("Placement", "Outside", "Outside", "", "Для арматуры задает сторону подписи от линии осей стержней; для размеров задает только сторону текста относительно размерной линии."),
        @("Offset", "60", "100", "мм", "Визуальный зазор в реальных миллиметрах сечения: для арматуры откладывается только от линии осей стержней, Placement задает сторону; для размеров - от грани до размерной линии."),
        @("LineEnabled", "Yes", "-", "Yes/No", "Показывать короткую линию обозначения арматуры; если No, остается только текст подписи."),
        @("LineWeight", "1.35", "1.75", "pt", "Толщина основной линии аннотации."),
        @("ExtensionLineWeight", "-", "0.85", "pt", "Толщина выносных линий; применяется только для размерных линий."),
        @("TextHeight", "25", "25", "мм", "Высота текста в реальных миллиметрах сечения; при выводе пересчитывается в размер шрифта Chart."),
        @("TextGap", "15", "15", "мм", "Зазор между линией аннотации и текстом в реальных миллиметрах сечения."),
        @("ArrowType", "-", "Triangle", "", "Тип стрелки размерной линии."),
        @("ArrowSize", "-", "Medium", "", "Размер стрелок размерной линии."),
        @("Color", "20,30,90", "20,30,90", "RGB", "Цвет текста и основной линии в формате R,G,B."),
        @("ExtensionLineColor", "-", "140,140,140", "RGB", "Цвет выносных линий размеров в формате R,G,B.")
    )
}

function Add-PlotAnnotationSettingsTable {
    param([object]$Workbook, [object]$Sheet, [int]$HeaderRow, [int]$StartColumn)

    $title = "[Аннотации схемы]"
    $Sheet.Cells.Item($HeaderRow - 1, $StartColumn).Value2 = $title
    $Sheet.Range($Sheet.Cells.Item($HeaderRow - 1, $StartColumn), $Sheet.Cells.Item($HeaderRow - 1, $StartColumn + 4)).Merge() | Out-Null
    $Sheet.Cells.Item($HeaderRow - 1, $StartColumn).Font.Bold = $true
    $Sheet.Cells.Item($HeaderRow - 1, $StartColumn).Interior.Color = 15921906

    $headers = @("Параметр", "Обозначения арматуры", "Размерные линии", "Ед.", "Комментарий")
    for ($i = 0; $i -lt $headers.Count; $i++) {
        $cell = $Sheet.Cells.Item($HeaderRow, $StartColumn + $i)
        $cell.Value2 = $headers[$i]
        $cell.Font.Bold = $true
        $cell.Interior.Color = 14277081
    }

    $rows = Get-PlotAnnotationSettingsCatalog
    for ($r = 0; $r -lt $rows.Count; $r++) {
        for ($c = 0; $c -lt 5; $c++) {
            $Sheet.Cells.Item($HeaderRow + 1 + $r, $StartColumn + $c).Value2 = $rows[$r][$c]
        }
    }

    Set-WorkbookNameByBounds $Workbook "rngPlotAnnotationSettings" $Sheet $HeaderRow $StartColumn ($HeaderRow + $rows.Count) ($StartColumn + 4)
}

function Get-InputUnitFormulaForUnit {
    param([string]$UnitText)
    switch ($UnitText) {
        "мм" { return '=INDEX(rngUnitSettings,MATCH("Length",INDEX(rngUnitSettings,,1),0),2)' }
        "мм2" { return '=INDEX(rngUnitSettings,MATCH("Area",INDEX(rngUnitSettings,,1),0),2)' }
        "Н" { return '=INDEX(rngUnitSettings,MATCH("Force",INDEX(rngUnitSettings,,1),0),2)' }
        "Н*мм" { return '=INDEX(rngUnitSettings,MATCH("Moment",INDEX(rngUnitSettings,,1),0),2)' }
        "МПа" { return '=INDEX(rngUnitSettings,MATCH("Stress",INDEX(rngUnitSettings,,1),0),2)' }
        "1/мм" { return '=INDEX(rngUnitSettings,MATCH("Curvature",INDEX(rngUnitSettings,,1),0),2)' }
        default { return $null }
    }
}

function Set-InputUnitCell {
    param([object]$Cell, [string]$UnitText)
    $formula = Get-InputUnitFormulaForUnit $UnitText
    if ($null -ne $formula) {
        $Cell.Formula = $formula
    } else {
        $Cell.Value2 = $UnitText
    }
}

function Apply-SystemSettingsLayout {
    param([object]$Workbook, [object]$Sheet)

    $catalog = Get-SystemSettingsCatalog
    $Sheet.Range("A1:DK260").ClearContents()
    $Sheet.Range("A1:DK260").Validation.Delete()

    $Sheet.Cells.Item(1, 1).Value2 = "Settings"
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
            Set-InputUnitCell $Sheet.Cells.Item($row, 3) ([string]$setting[2])
            $row++
        }
        $row++
    }

    $settingsRange = $Sheet.Range($Sheet.Cells.Item(3, 1), $Sheet.Cells.Item($row - 2, 4))
    Set-WorkbookNameByBounds $Workbook "rngSystemSettings" $Sheet 3 1 ($row - 2) 4

    Add-MaterialDiagramTable $Workbook $Sheet "rngConcreteDiagramPoints" 3 6 "Точки диаграммы бетона" @(
        @{ Point = 1; Strain = 0.0001; Stress = 1.8 },
        @{ Point = 2; Strain = 0; Stress = 0 },
        @{ Point = 3; Strain = -0.0015; Stress = -15.5 },
        @{ Point = 4; Strain = -0.002; Stress = -15.5 },
        @{ Point = 5; Strain = -0.0035; Stress = -15.5 }
    )
    Add-MaterialDiagramTable $Workbook $Sheet "rngSteelDiagramPoints" 3 10 "Точки диаграммы арматуры" @(
        @{ Point = 1; Strain = -0.025; Stress = -350 },
        @{ Point = 2; Strain = -0.01; Stress = -350 },
        @{ Point = 3; Strain = -0.00175; Stress = -350 },
        @{ Point = 4; Strain = 0; Stress = 0 },
        @{ Point = 5; Strain = 0.00175; Stress = 350 },
        @{ Point = 6; Strain = 0.01; Stress = 350 },
        @{ Point = 7; Strain = 0.025; Stress = 350 }
    )

    Add-UnitSettingsTable $Workbook $Sheet 3 14
    Add-SignConventionSettingsTable $Workbook $Sheet 13 14
    Add-PlotAnnotationSettingsTable $Workbook $Sheet 20 16

    foreach ($geometryTable in (Get-GeometrySettingsCatalog)) {
        if ($geometryTable.ContainsKey("FaceTable") -and $geometryTable.FaceTable) {
            Add-LShapeFaceSettingsTable $Workbook $Sheet $geometryTable.RangeName $geometryTable.StartRow $geometryTable.StartColumn $geometryTable.Title $geometryTable.Rows
        } else {
            Add-GeometrySettingsTable $Workbook $Sheet $geometryTable.RangeName $geometryTable.StartRow $geometryTable.StartColumn $geometryTable.Title $geometryTable.Rows
        }
    }

    $validationLists = [ordered]@{
        "Geometry.Source" = @("Generated", "AutoCAD")
        "Geometry.Type" = @("RoundedRectangle", "Circle", "LShape")
        "Concrete.TensionMode" = @("Ignore", "UseDiagram")
        "Calculation.Mode" = @("DirectState", "FullCapacity")
        "Solver.Method" = @("Newton", "Secant")
        "Capacity.Method" = @("LoadMultiplier", "UltimateStrain")
        "Capacity.SearchMethod" = @("Bisection", "Brent", "Secant")
        "Solver.LineSearchEnabled" = @("Yes", "No")
        "CrackWidth.Enabled" = @("Yes", "No")
        "AutoCAD.Export.ResultType" = @("Stress", "Strain")
        "AutoCAD.Export.LabelMode" = @("ValuesOnly", "NamesAndValues")
        "AutoCAD.Export.NeutralLineEnabled" = @("Yes", "No")
        "AutoCAD.Export.PrincipalAxesEnabled" = @("Yes", "No")
        "AutoCAD.Export.LoadPointEnabled" = @("Yes", "No")
        "Plot.Enabled" = @("Yes", "No")
        "Plot.AutoUpdateAfterCalculation" = @("Yes", "No")
        "Plot.ResultType" = @("Stress", "Strain")
        "Plot.ResultGradient" = @("Yes", "No")
        "Plot.ResultLabelsEnabled" = @("Yes", "No")
        "Plot.NeutralLineEnabled" = @("Yes", "No")
        "Plot.PrincipalAxesEnabled" = @("Yes", "No")
        "Plot.LoadApplicationPointEnabled" = @("Yes", "No")
        "Plot.CentroidEnabled" = @("Yes", "No")
        "Plot.LegendEnabled" = @("Yes", "No")
        "Plot.LegendMode" = @("Separate", "Common")
    }

    $listColumn = 80
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

    $exportCombinationListColumn = $listColumn
    $Sheet.Cells.Item(1, $exportCombinationListColumn).Value2 = "Worst"
    for ($i = 1; $i -le 20; $i++) {
        $Sheet.Cells.Item($i + 1, $exportCombinationListColumn).Formula = '=IF(INDEX(rngLoadCombinations,' + ($i + 1) + ',1)="","",INDEX(rngLoadCombinations,' + ($i + 1) + ',1))'
    }
    $exportCombinationColName = ConvertTo-ExcelColumn $exportCombinationListColumn
    $exportCombinationListAddress = "=$" + $exportCombinationColName + '$1:$' + $exportCombinationColName + '$21'
    for ($r = 4; $r -le ($row - 2); $r++) {
        if ([string]$Sheet.Cells.Item($r, 1).Value2 -eq "AutoCAD.Export.CombinationID" -or [string]$Sheet.Cells.Item($r, 1).Value2 -eq "Plot.LoadCase") {
            $cell = $Sheet.Cells.Item($r, 2)
            $cell.Validation.Delete()
            $cell.Validation.Add(3, 1, 1, $exportCombinationListAddress)
            $cell.Validation.IgnoreBlank = $false
            $cell.Validation.InCellDropdown = $true
        }
    }

    Add-PlotAnnotationValidation $Sheet

    $Sheet.Columns.Item(1).ColumnWidth = 34
    $Sheet.Columns.Item(2).ColumnWidth = 18
    $Sheet.Columns.Item(3).ColumnWidth = 12
    $Sheet.Columns.Item(4).ColumnWidth = 86
    $Sheet.Columns.Item(14).ColumnWidth = 14
    $Sheet.Columns.Item(15).ColumnWidth = 16
    $Sheet.Columns.Item(16).ColumnWidth = 16
    $Sheet.Columns.Item(17).ColumnWidth = 16
    $Sheet.Columns.Item("BR:DK").Hidden = $true
    $Sheet.Range("A1:DK260").Font.Name = "Arial"
    $Sheet.Range("A1:DK260").Font.Size = 9

    return @{ Settings = $settingsRange }
}

function Add-PlotAnnotationValidation {
    param([object]$Sheet)

    $validationSources = @{
        "Enabled" = @{ Column = 110; Values = @("Yes", "No") }
        "Placement" = @{ Column = 111; Values = @("Outside", "Inside") }
        "LineEnabled" = @{ Column = 110; Values = @("Yes", "No") }
        "ArrowType" = @{ Column = 112; Values = @("Triangle", "Stealth", "Diamond", "Oval", "Open") }
        "ArrowSize" = @{ Column = 113; Values = @("Small", "Medium", "Wide") }
    }
    foreach ($key in $validationSources.Keys) {
        $source = $validationSources[$key]
        for ($i = 0; $i -lt $source.Values.Count; $i++) {
            $Sheet.Cells.Item($i + 1, $source.Column).Value2 = $source.Values[$i]
        }
    }

    $range = $Sheet.Parent.Names.Item("rngPlotAnnotationSettings").RefersToRange
    $rows = $range.Value2
    for ($r = 2; $r -le $range.Rows.Count; $r++) {
        $paramName = [string]$rows[$r, 1]
        $source = $validationSources[$paramName]

        if ($null -ne $source) {
            for ($c = 2; $c -le 3; $c++) {
                if ([string]$rows[$r, $c] -ne "-") {
                    $cell = $range.Cells.Item($r, $c)
                    $cell.Validation.Delete()
                    $colName = ConvertTo-ExcelColumn $source.Column
                    $listAddress = "=$" + $colName + '$1:$' + $colName + '$' + $source.Values.Count
                    $cell.Validation.Add(3, 1, 1, $listAddress)
                    $cell.Validation.IgnoreBlank = $false
                    $cell.Validation.InCellDropdown = $true
                }
            }
        }
    }
}

function Add-GeometrySettingsTable {
    param(
        [object]$Workbook,
        [object]$Sheet,
        [string]$RangeName,
        [int]$HeaderRow,
        [int]$StartColumn,
        [string]$Title,
        [array]$Rows
    )

    $Sheet.Cells.Item($HeaderRow - 1, $StartColumn).Value2 = $Title
    $Sheet.Range($Sheet.Cells.Item($HeaderRow - 1, $StartColumn), $Sheet.Cells.Item($HeaderRow - 1, $StartColumn + 3)).Merge() | Out-Null
    $Sheet.Cells.Item($HeaderRow - 1, $StartColumn).Font.Bold = $true
    $Sheet.Cells.Item($HeaderRow - 1, $StartColumn).Interior.Color = 15921906

    $headers = @("Параметр", "Значение", "Ед.", "Комментарий")
    for ($i = 0; $i -lt $headers.Count; $i++) {
        $cell = $Sheet.Cells.Item($HeaderRow, $StartColumn + $i)
        $cell.Value2 = $headers[$i]
        $cell.Font.Bold = $true
        $cell.Interior.Color = 14277081
    }

    for ($r = 0; $r -lt $Rows.Count; $r++) {
        for ($c = 0; $c -lt 4; $c++) {
            $Sheet.Cells.Item($HeaderRow + 1 + $r, $StartColumn + $c).Value2 = $Rows[$r][$c]
        }
        Set-InputUnitCell $Sheet.Cells.Item($HeaderRow + 1 + $r, $StartColumn + 2) ([string]$Rows[$r][2])
    }

    $locOptions = @("Stacked", "SideBySide")
    $listColumn = 61
    for ($i = 0; $i -lt $locOptions.Count; $i++) {
        $Sheet.Cells.Item($i + 1, $listColumn).Value2 = $locOptions[$i]
    }
    $listColName = ConvertTo-ExcelColumn $listColumn
    $listAddress = "=$" + $listColName + '$1:$' + $listColName + '$' + $locOptions.Count
    for ($r = 0; $r -lt $Rows.Count; $r++) {
        $parameterName = [string]$Rows[$r][0]
        if ($parameterName -eq "Rebar.Loc2row" -or $parameterName -eq "Rebar.Loc3row") {
            $cell = $Sheet.Cells.Item($HeaderRow + 1 + $r, $StartColumn + 1)
            $cell.Validation.Delete()
            $cell.Validation.Add(3, 1, 1, $listAddress)
            $cell.Validation.IgnoreBlank = $true
            $cell.Validation.InCellDropdown = $true
        }
    }

    Set-WorkbookNameByBounds $Workbook $RangeName $Sheet $HeaderRow $StartColumn ($HeaderRow + $Rows.Count) ($StartColumn + 3)
}

function Add-LShapeFaceSettingsTable {
    param(
        [object]$Workbook,
        [object]$Sheet,
        [string]$RangeName,
        [int]$HeaderRow,
        [int]$StartColumn,
        [string]$Title,
        [array]$Rows
    )

    $Sheet.Cells.Item($HeaderRow - 1, $StartColumn).Value2 = $Title
    $Sheet.Range($Sheet.Cells.Item($HeaderRow - 1, $StartColumn), $Sheet.Cells.Item($HeaderRow - 1, $StartColumn + 8)).Merge() | Out-Null
    $Sheet.Cells.Item($HeaderRow - 1, $StartColumn).Font.Bold = $true
    $Sheet.Cells.Item($HeaderRow - 1, $StartColumn).Interior.Color = 15921906

    $geometryValues = @{
        "H1" = $Rows[0][1]
        "B1" = $Rows[0][2]
        "H2" = $Rows[0][3]
        "B2" = $Rows[0][4]
    }
    $mainRows = @(
        @("H1 - левая",   $Rows[1][1], $Rows[3][1], $Rows[5][1], $Rows[7][1],  $Rows[8][1],  "мм", "Грань _1: левая грань H1."),
        @("H1 - правая",  $Rows[2][1], $Rows[4][1], $Rows[6][1], $Rows[9][1],  $Rows[10][1], "мм", "Грань _2: правая грань H1."),
        @("B1 - верхняя", $Rows[1][2], $Rows[3][2], $Rows[5][2], $Rows[7][2],  $Rows[8][2],  "мм", "Грань _1: верхняя грань B1."),
        @("B1 - нижняя",  $Rows[2][2], $Rows[4][2], $Rows[6][2], $Rows[9][2],  $Rows[10][2], "мм", "Грань _2: нижняя грань B1."),
        @("H2 - левая",   $Rows[1][3], $Rows[3][3], $Rows[5][3], $Rows[7][3],  $Rows[8][3],  "мм", "Грань _1: левая грань H2."),
        @("H2 - правая",  $Rows[2][3], $Rows[4][3], $Rows[6][3], $Rows[9][3],  $Rows[10][3], "мм", "Грань _2: правая грань H2."),
        @("B2 - верхняя", $Rows[1][4], $Rows[3][4], $Rows[5][4], $Rows[7][4],  $Rows[8][4],  "мм", "Грань _1: верхняя грань B2."),
        @("B2 - нижняя",  $Rows[2][4], $Rows[4][4], $Rows[6][4], $Rows[9][4],  $Rows[10][4], "мм", "Грань _2: нижняя грань B2.")
    )
    $extraRows = @(
        @("H1 - левая",   $Rows[12][1], $Rows[16][1], $Rows[18][1], $Rows[14][1], $Rows[17][1], $Rows[19][1], "мм", "Дополнительные ряды у грани _1 H1."),
        @("H1 - правая",  $Rows[13][1], $Rows[16][1], $Rows[18][1], $Rows[15][1], $Rows[17][1], $Rows[19][1], "мм", "Дополнительные ряды у грани _2 H1."),
        @("B1 - верхняя", $Rows[12][2], $Rows[16][2], $Rows[18][2], $Rows[14][2], $Rows[17][2], $Rows[19][2], "мм", "Дополнительные ряды у грани _1 B1."),
        @("B1 - нижняя",  $Rows[13][2], $Rows[16][2], $Rows[18][2], $Rows[15][2], $Rows[17][2], $Rows[19][2], "мм", "Дополнительные ряды у грани _2 B1."),
        @("H2 - левая",   $Rows[12][3], $Rows[16][3], $Rows[18][3], $Rows[14][3], $Rows[17][3], $Rows[19][3], "мм", "Дополнительные ряды у грани _1 H2."),
        @("H2 - правая",  $Rows[13][3], $Rows[16][3], $Rows[18][3], $Rows[15][3], $Rows[17][3], $Rows[19][3], "мм", "Дополнительные ряды у грани _2 H2."),
        @("B2 - верхняя", $Rows[12][4], $Rows[16][4], $Rows[18][4], $Rows[14][4], $Rows[17][4], $Rows[19][4], "мм", "Дополнительные ряды у грани _1 B2."),
        @("B2 - нижняя",  $Rows[13][4], $Rows[16][4], $Rows[18][4], $Rows[15][4], $Rows[17][4], $Rows[19][4], "мм", "Дополнительные ряды у грани _2 B2.")
    )

    $Sheet.Cells.Item($HeaderRow, $StartColumn).Value2 = "ГЕОМЕТРИЯ"
    $Sheet.Cells.Item($HeaderRow, $StartColumn).Font.Bold = $true
    $Sheet.Cells.Item($HeaderRow, $StartColumn).Interior.Color = 15921906
    $geomHeaders = @("H1", "B1", "H2", "B2")
    for ($i = 0; $i -lt $geomHeaders.Count; $i++) {
        $cell = $Sheet.Cells.Item($HeaderRow + 1, $StartColumn + $i)
        $cell.Value2 = $geomHeaders[$i]
        $cell.Font.Bold = $true
        $cell.Interior.Color = 14277081
        $Sheet.Cells.Item($HeaderRow + 2, $StartColumn + $i).Value2 = $geometryValues[$geomHeaders[$i]]
    }
    Set-InputUnitCell $Sheet.Cells.Item($HeaderRow + 1, $StartColumn + 4) "мм"

    $mainHeaderRow = $HeaderRow + 5
    $Sheet.Cells.Item($mainHeaderRow - 1, $StartColumn).Value2 = "ОСНОВНОЕ АРМИРОВАНИЕ"
    $Sheet.Cells.Item($mainHeaderRow - 1, $StartColumn).Font.Bold = $true
    $Sheet.Cells.Item($mainHeaderRow - 1, $StartColumn).Interior.Color = 15921906
    $mainHeaders = @("Грань", "as", "d", "n", "t нач.", "t кон.", "Ед.", "Комментарий")
    for ($i = 0; $i -lt $mainHeaders.Count; $i++) {
        $cell = $Sheet.Cells.Item($mainHeaderRow, $StartColumn + $i)
        $cell.Value2 = $mainHeaders[$i]
        $cell.Font.Bold = $true
        $cell.Interior.Color = 14277081
    }
    for ($r = 0; $r -lt $mainRows.Count; $r++) {
        for ($c = 0; $c -lt $mainRows[$r].Count; $c++) {
            $Sheet.Cells.Item($mainHeaderRow + 1 + $r, $StartColumn + $c).Value2 = $mainRows[$r][$c]
        }
        Set-InputUnitCell $Sheet.Cells.Item($mainHeaderRow + 1 + $r, $StartColumn + 6) "мм"
    }

    $extraHeaderRow = $mainHeaderRow + 11
    $Sheet.Cells.Item($extraHeaderRow - 1, $StartColumn).Value2 = "ДОПОЛНИТЕЛЬНЫЕ РЯДЫ"
    $Sheet.Cells.Item($extraHeaderRow - 1, $StartColumn).Font.Bold = $true
    $Sheet.Cells.Item($extraHeaderRow - 1, $StartColumn).Interior.Color = 15921906
    $extraHeaders = @("Грань", "d2", "положение", "привязка", "d3", "положение", "привязка", "Ед.", "Комментарий")
    for ($i = 0; $i -lt $extraHeaders.Count; $i++) {
        $cell = $Sheet.Cells.Item($extraHeaderRow, $StartColumn + $i)
        $cell.Value2 = $extraHeaders[$i]
        $cell.Font.Bold = $true
        $cell.Interior.Color = 14277081
    }
    for ($r = 0; $r -lt $extraRows.Count; $r++) {
        for ($c = 0; $c -lt $extraRows[$r].Count; $c++) {
            $Sheet.Cells.Item($extraHeaderRow + 1 + $r, $StartColumn + $c).Value2 = $extraRows[$r][$c]
        }
        Set-InputUnitCell $Sheet.Cells.Item($extraHeaderRow + 1 + $r, $StartColumn + 7) "мм"
    }

    $locOptions = @("Stacked", "SideBySide")
    $listColumn = 114
    for ($i = 0; $i -lt $locOptions.Count; $i++) {
        $Sheet.Cells.Item($i + 1, $listColumn).Value2 = $locOptions[$i]
    }
    $listColName = ConvertTo-ExcelColumn $listColumn
    $listAddress = "=$" + $listColName + '$1:$' + $listColName + '$' + $locOptions.Count
    $bindOptions = @("EachBar", "EverySecondBar")
    $bindListColumn = 115
    for ($i = 0; $i -lt $bindOptions.Count; $i++) {
        $Sheet.Cells.Item($i + 1, $bindListColumn).Value2 = $bindOptions[$i]
    }
    $bindListColName = ConvertTo-ExcelColumn $bindListColumn
    $bindListAddress = "=$" + $bindListColName + '$1:$' + $bindListColName + '$' + $bindOptions.Count

    for ($r = 0; $r -lt $extraRows.Count; $r++) {
        foreach ($valueColumn in @(2, 5)) {
            $cell = $Sheet.Cells.Item($extraHeaderRow + 1 + $r, $StartColumn + $valueColumn)
            $cell.Validation.Delete()
            $cell.Validation.Add(3, 1, 1, $listAddress)
            $cell.Validation.IgnoreBlank = $true
            $cell.Validation.InCellDropdown = $true
        }
        foreach ($valueColumn in @(3, 6)) {
            $cell = $Sheet.Cells.Item($extraHeaderRow + 1 + $r, $StartColumn + $valueColumn)
            $cell.Validation.Delete()
            $cell.Validation.Add(3, 1, 1, $bindListAddress)
            $cell.Validation.IgnoreBlank = $true
            $cell.Validation.InCellDropdown = $true
        }
    }

    Set-WorkbookNameByBounds $Workbook $RangeName $Sheet $HeaderRow $StartColumn ($extraHeaderRow + $extraRows.Count) ($StartColumn + 8)
}

function Add-MaterialDiagramTable {
    param(
        [object]$Workbook,
        [object]$Sheet,
        [string]$RangeName,
        [int]$HeaderRow,
        [int]$StartColumn,
        [string]$Title,
        [array]$Points
    )

    $Sheet.Cells.Item($HeaderRow - 1, $StartColumn).Value2 = $Title
    $Sheet.Range($Sheet.Cells.Item($HeaderRow - 1, $StartColumn), $Sheet.Cells.Item($HeaderRow - 1, $StartColumn + 2)).Merge() | Out-Null
    $Sheet.Cells.Item($HeaderRow - 1, $StartColumn).Font.Bold = $true
    $Sheet.Cells.Item($HeaderRow - 1, $StartColumn).Interior.Color = 15921906

    $headers = @("Point", "Strain", "Stress")
    for ($i = 0; $i -lt $headers.Count; $i++) {
        $cell = $Sheet.Cells.Item($HeaderRow, $StartColumn + $i)
        $cell.Value2 = $headers[$i]
        $cell.Font.Bold = $true
        $cell.Interior.Color = 14277081
    }
    $Sheet.Cells.Item($HeaderRow, $StartColumn + 2).Formula = '="Stress, "&INDEX(rngUnitSettings,MATCH("Stress",INDEX(rngUnitSettings,,1),0),2)'

    for ($r = 1; $r -le 10; $r++) {
        for ($c = 0; $c -lt 3; $c++) {
            $Sheet.Cells.Item($HeaderRow + $r, $StartColumn + $c).ClearContents()
        }
    }

    for ($i = 0; $i -lt $Points.Count; $i++) {
        $targetRow = $HeaderRow + 1 + $i
        $Sheet.Cells.Item($targetRow, $StartColumn).Value2 = [double]$Points[$i].Point
        $Sheet.Cells.Item($targetRow, $StartColumn + 1).Value2 = [double]$Points[$i].Strain
        $Sheet.Cells.Item($targetRow, $StartColumn + 2).Value2 = [double]$Points[$i].Stress
    }

    Set-WorkbookNameByBounds $Workbook $RangeName $Sheet ($HeaderRow + 1) $StartColumn ($HeaderRow + 10) ($StartColumn + 2)
}

function Add-SystemSettings {
    param([object]$Sheet)
    $result = Apply-SystemSettingsLayout $Sheet.Parent $Sheet
    $result
}
