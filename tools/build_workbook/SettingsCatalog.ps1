# скрипт описывает каталог настроек Config и строит таблицы настроек, единиц, геометрии и валидации.

. (Join-Path $PSScriptRoot 'WorkbookLayout.ps1')

$script:FormulaImageCounter = 0
$script:FormulaImageRows = New-Object System.Collections.Generic.List[object]
$script:FormulaImageDir = Join-Path (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot "../..")) "workbook/output/formula_images"
$script:ConfigGroupHeaderColor = 14277081
$script:ConfigSubgroupHeaderColor = 15921906
$script:ConfigTableHeaderColor = 14277081

function Get-SystemSettingsCatalog {
    $lambda = [char]0x03BB
    $mu = [char]0x03BC
    @(
        @{ Name = "GeneralSettings"; Title = "[Общие]"; Rows = @(
            @("General.ExecutionReportEnabled", "Yes", "-", "Включить подробную диагностику выполнения расчета. Диагностика обновляется при каждом новом запуске."),
            @("General.NonCriticalMessagesEnabled", "Yes", "-", "Показывать информационные окна об успешном расчете, обновлении схемы, импорте и экспорте. Ошибки и предупреждения выводятся всегда."),
            @("General.ReformatResultsEnabled", "Yes", "-", "Повторно оформлять таблицы Results и их шапки. Yes — полное оформление; No — сохранять готовое, обновляя данные, единицы, цвета статусов и необходимое оформление новых строк."),
            @("Results.SnapshotStorage", "Excel", "-", "Excel — подробные результаты хранятся на листе Results. File — полный снимок хранится в файле .ndm рядом с книгой, на листе остаются сводки и общая геометрия. Кнопка «Обновить схему и результаты» загружает снимок без повторного расчёта."),
            @("Results.SnapshotName", "NDM_Snapshot", "-", "Имя снимка без пути и расширения. Файл сохраняется в формате .ndm рядом с книгой. Новый расчёт заменяет файл с этим именем; обновление схемы и результатов загружает его без повторного решения.")
        )},
        @{ Name = "LoadDiagramSettings"; Title = "[Нагрузки и диаграммы материалов]"; Rows = @(
            @("Load.ReferenceOffsetX", "0", "мм", "Смещение точки приложения нагрузок по X относительно центра тяжести бетонного сечения. 0 означает, что N, Mx и My заданы относительно центра бетона без учета арматуры."),
            @("Load.ReferenceOffsetY", "0", "мм", "Смещение точки приложения нагрузок по Y относительно центра тяжести бетонного сечения. Рабочие моменты для решателя автоматически переносятся к координатам расчетных элементов."),
            @("Calculation.ZeroMomentPerDepth", "0.0005", "Н*мм/мм", "Порог обнуления малых моментов: M_tol = значение·h. Убирает численный остаток после переносов и поворотов; слишком большое значение может обнулить значимый малый момент."),
            @("General.WorstCombinationCriterion", "StrengthCapacity", "-", "Критерий выбора сочетания Worst в верхней сводке Results: прочность по деформациям, прочность по несущей, трещины или устойчивость."),
            @("General.PlateauSlopeFactor", "0.001", "-", "Коэффициент α определяет наклон ненулевых плато бетона и арматуры: Et = α·E. По умолчанию α = 0,001, наклон плато — 0,001·E (0,1% от E своей ветви). 0 — горизонтальное плато. Наклон помогает сходимости и меняет напряжения; предельные деформации сохраняются."),
            @("General.DiagramExtension", "Yes", "-", "Численное продолжение диаграмм: продолжает активные ветви за физическими пределами. Рекомендуется для сходимости и диагностики. Не увеличивает физические пределы и не делает состояние в продолжении допустимым. При No работают обычные диаграммы.")
        )},
        @{ Name = "GeometrySettings"; Title = "[Геометрия]"; Rows = @(
            @("Geometry.Source", "Generated", "-", "Generated - встроенная геометрия; AutoCAD - импорт Region или генерация бетона по точным контурам при пустом слое бетона. Арматура импортируется из Region. Геометрия AutoCAD всегда в мм."),
            @("Geometry.Type", "Rectangles", "-", "Тип сечения для Geometry.Source = Generated. Поддерживается: PointedSection, HollowRectangle, Circle или Rectangles.")
        )},
        @{ Name = "MeshSettings"; Title = "[Сетка геометрии]"; Rows = @(
            @("Mesh.MinElementArea", "0.000001", "мм2", "Минимальная площадь бетонного элемента или стержня при формировании сечения. Части меньшей площади пропускаются, равной — сохраняются. Применяется к встроенным формам и импорту, включая части после обрезания. Ввод в INPUT Area; 0 отключает фильтр."),
            @("Mesh.CellGeometry", "Contour", "-", "Adaptive — более быстрое формирование прямоугольной сетки с приближённой границей. Contour — точное обрезание ячеек по наружным контурам и отверстиям с сохранением прямых и круговых дуг."),
            @("Mesh.StepMode", "Auto 0.1", "-", "Auto — шаг как доля меньшего габарита сечения; User — независимые ручные шаги Mesh.StepX и Mesh.StepY. Adaptive уточняет пограничные ячейки; Contour обрезает их по фактической границе."),
            @("Mesh.StepX", "50", "мм", "Ручной шаг бетонной сетки по X при Mesh.StepMode = User. В Auto сохраняется, но не используется."),
            @("Mesh.StepY", "50", "мм", "Ручной шаг бетонной сетки по Y при Mesh.StepMode = User. В Auto сохраняется, но не используется; автоматические шаги X/Y одинаковы."),
            @("Mesh.BoundarySubdivisions", "2", "шт", "Только Mesh.CellGeometry = Adaptive: уровни деления пересечённых ячеек пополам по X/Y; N = 2 даёт h/4. В Contour сохраняется, но не используется.")
        )},
        @{ Name = "SolverSettings"; Title = "[Настройки решателя НДС]"; Rows = @(
            @("[Общие настройки]", "", "-", ""),
            @("Solver.Method", "Newton", "-", "Метод поиска равновесия НДС. Допустимо только Newton или Secant; пустое или неизвестное значение считается ошибкой исходных данных."),
            @("Solver.MaxIterations", "80", "шт", "Применяется к Newton и Secant. Максимальное число итераций на ступень нагрузки; увеличение повышает шанс сходимости, но увеличивает время."),
            @("Solver.LoadSteps", "1", "шт", "Применяется к Newton и Secant. Число ступеней приложения нагрузки в прямом НДМ; больше ступеней обычно устойчивее, но медленнее."),
            @("Solver.ToleranceN", "0.0001", "Н", "Применяется к Newton и Secant. Абсолютный допуск равновесия по продольной силе в выбранной INPUT-единице силы; значение переводится во внутренние Н."),
            @("Solver.ToleranceMx", "0.0001", "Н*мм", "Применяется к Newton и Secant. Абсолютный допуск равновесия по моменту Mx в выбранной INPUT-единице момента; значение переводится во внутренние Н*мм."),
            @("Solver.ToleranceMy", "0.0001", "Н*мм", "Применяется к Newton и Secant. Абсолютный допуск равновесия по моменту My в выбранной INPUT-единице момента; значение переводится во внутренние Н*мм."),
            @("[Secant]", "", "-", ""),
            @("Solver.SecantMaxRestarts", "2", "шт", "Применяется только к Secant. Максимальное число контролируемых перезапусков приближенной матрицы чувствительности.")
        )},
        @{ Name = "CapacitySettings"; Title = "[Поиск предельной несущей способности]"; Rows = @(
            @("[Общие настройки]", "", "-", ""),
            @("Capacity.SolutionStrategy", "Auto", "-", "Предпочтительная стратегия поиска предельной нагрузки: Auto, UltimateStrain или LoadMultiplier. Для отдельных траекторий программа может выбрать более устойчивую ветку и покажет ее в Results."),
            @("Capacity.SolverMaxIterations", "60", "шт", "Актуально для всех методов. Ограничивает число попыток найти равновесие или предельное состояние."),
            @("[UltimateStrain]", "", "-", ""),
            @("Capacity.ToleranceStrain", "0.00001", "-", "Актуально для UltimateStrain и первой ветки Auto. Абсолютный допуск достижения предельной деформации критического бетонного волокна или стержня."),
            @("[LoadMultiplier]", "", "-", ""),
            @("Capacity.MaxLambda", "1024", "-", "Только для LoadMultiplier. Верхняя граница расширения интервала поиска λ."),
            @("Capacity.SearchMethod", "Bisection", "-", "Метод одномерного поиска для LoadMultiplier и для повторной попытки в Auto. Поддерживается только Bisection, Brent или Secant."),
            @("Capacity.InitialLambda", "1", "-", "Начальный множитель λ при поиске верхней границы несущей способности."),
            @("Capacity.ToleranceLambda", "0.01", "-", "Общий допуск одномерного поиска λ для Bisection, Brent и Secant."),
            @("Capacity.MaxRetries", "0", "шт", "Число повторов той же λ-точки после NumFail/SingularTangent. 0 - без повторов; при повторе LoadSteps = BaseLoadSteps·2^attempt."),
            @("Capacity.BaseLoadSteps", "1", "шт", "Базовое число ступеней нагрузки первой попытки λ-точки LoadMultiplier. При MaxRetries > 0 повторы автоматически используют 2, 4, ... базовых шага.")
        )},
        @{ Name = "StabilitySettings"; Title = "[Продольный изгиб и устойчивость]"; Rows = @(
            @("[Общие настройки]", "", "-", ""),
            @("Stability.Code", "SP35", "-", "Выбор методики учета продольного изгиба: SP63 или SP35. Расчет включается отдельно в профиле."),
            @("Stability.ElementLength", "8000", "мм", "Геометрическая длина элемента L. Расчетная длина по главным плоскостям получается как l_0,1 = μ_1·L и l_0,2 = μ_2·L."),
            @("Stability.Mu1", "2", "-", "Коэффициент расчетной длины для плоскости 1: изгиб вокруг главной оси 1 с моментом инерции I_1."),
            @("Stability.Mu2", "2", "-", "Коэффициент расчетной длины для плоскости 2: изгиб вокруг главной оси 2 с моментом инерции I_2."),
            @("Stability.SystemType", "Determinate", "-", "Расчетная схема элемента: Determinate - статически определимая; Indeterminate - статически неопределимая."),
            @("Stability.ZeroMomentEccentricitySign1", "1", "-", "Направление случайного эксцентриситета при M1=0: положительная нормаль к главной оси 1 или обратная сторона."),
            @("Stability.ZeroMomentEccentricitySign2", "1", "-", "Направление случайного эксцентриситета при M2=0: положительная нормаль к главной оси 2 или обратная сторона."),
            @("Stability.AccidentalEccentricityMode", "AutoWith$($mu)L", "-", "Как задавать случайный эксцентриситет: AutoWithL - по геометрической длине L; AutoWith$($mu)L - по расчетной длине $($mu)L; User - взять заданный модуль."),
            @("Stability.AccidentalEccentricityUser1", "10", "мм", "Модуль случайного эксцентриситета плоскости 1 для режима User: плечо по нормали к главной оси 1."),
            @("Stability.AccidentalEccentricityUser2", "10", "мм", "Модуль случайного эксцентриситета плоскости 2 для режима User: плечо по нормали к главной оси 2."),
            @("Stability.AccidentalEccentricityPlanes", "OnlyMomentPlane", "-", "OnlyMomentPlane - учитывать случайный эксцентриситет только в плоскостях с моментом; BothPlanes - учитывать в обеих главных плоскостях."),
            @("Stability.PhiLMode", "Auto", "-", "Auto - коэффициент длительности считается по длительной части момента; PhiL2 - принять φ_l = 2 в расчетных ветвях N_cr/η."),
            @("[СП 63]", "", "-", ""),
            @("Stability.SP63.Ks", "0.7", "-", "Коэффициент ks при вкладе арматуры в условную жесткость D по СП 63."),
            @("Stability.SP63.DeltaEMin", "0.15", "-", "Нижнее ограничение δ_e в формуле условной жесткости СП 63."),
            @("Stability.SP63.DeltaEMax", "1.5", "-", "Верхнее ограничение δ_e в формуле условной жесткости СП 63."),
            @("[СП 35]", "", "-", ""),
            @("Stability.SP35.PhiP", "1", "-", "Коэффициент φ_p. Для ненапрягаемой арматуры в текущей постановке обычно принимается 1."),
            @("Stability.SP35.NOverNcrLimit", "0.7", "-", "Контрольное ограничение N/N_cr для ветви СП 35 с коэффициентом η; в Results выводится запас Stability.SP35.NOverNcrLimit·N_cr/N.")
        )},
        @{ Name = "OutputSettings"; Title = "[Трещины]"; Rows = @(
            @("[Общие настройки]", "", "-", ""),
            @("SLS.Crack.Code", "SP35", "-", "Методика ширины нормальных трещин: SP63 или SP35. Обе используют готовое текущее НДС; выбор не связан с Stability.Code."),
            @("SLS.Crack.Allowable", "0.3", "мм", "Допустимая ширина раскрытия a_crc,ult, задается пользователем по нормам и условиям эксплуатации."),
            @("SLS.Crack.InitiationSolutionStrategy", "Auto", "-", "Как искать уровень образования трещины: Auto, UltimateStrain или LoadMultiplier."),
            @("SLS.Crack.ToleranceLambda", "0.01", "-", "Допуск интервала λ при поиске образования трещины методом LoadMultiplier, в том числе при переходе из Auto. Не меняет физические пределы материалов и допуски равновесия."),
            @("[SP63]", "", "-", ""),
            @("SLS.Crack.TensionZoneMode", "Effective", "-", "Зона бетона A_bt: Effective - эффективная зона по 2a_s и 0.5h; FullTension - вся фактически растянутая зона."),
            @("SLS.Crack.CoverDistanceMode", "GlobalExtreme", "-", "Как считать a_s: GlobalExtreme - до дальней растянутой линии всего сечения; NearestContour - до ближайшего бетонного контура по нормали."),
            @("SLS.Crack.Phi1", "1.4", "-", "Коэффициент длительности действия нагрузки для формулы a_crc; для продолжительного раскрытия обычно 1.4."),
            @("SLS.Crack.Phi2", "0.5", "-", "Коэффициент поверхности арматуры в формуле a_crc. Пользователь задает значение напрямую: обычно 0.5 для периодического профиля и 0.8 для гладкого."),
            @("SLS.Crack.Phi3Mode", "Auto", "-", "Как принимать φ_3: Auto - по характеру продольной силы; User - брать введенное значение SLS.Crack.Phi3."),
            @("SLS.Crack.Phi3", "1", "-", "Пользовательское значение φ_3, если SLS.Crack.Phi3Mode = User."),
            @("SLS.Crack.PsiMode", "AlwaysCalc", "-", "При доступной точке образования: User - SLS.Crack.PsiS; Auto - уточнение при непрохождении; AlwaysCalc - всегда уточнение по σ_s,crc. Без пригодной точки: ψ_s = 1 с предупреждением."),
            @("SLS.Crack.SigmaSCrcAveragingMode", "AllSelected", "-", "Как усреднять только σ_s,crc: AllSelected - учитывать все выбранные стержни со знаком; TensionOnly - сжатые принимать как 0."),
            @("SLS.Crack.PsiS", "1", "-", "Пользовательское значение ψ_s, если SLS.Crack.PsiMode = User."),
            @("[SP35]", "", "-", ""),
            @("SLS.Crack.SP35.GroupGapTolerance", "10", "мм", "Принято в программе: Максимальное расстояние между поверхностями растянутых арматурных стержней для их объединения в группу; ввод в INPUT Length."),
            @("SLS.Crack.SP35.RowTolerance", "50", "мм", "Принято в программе: Разнесение центров групп одного ряда по нормали к НЛ. Не выбирает боковых соседей. Ряд 1 - наружный, номера растут к НЛ. Начальное значение 50 мм; ввод в INPUT Length."),
            @("SLS.Crack.SP35.NeighborRatioLimit", "0.2", "-", "Принято в программе: Порог t/L между рассматриваемой группой растянутых арматурных стержней и ближайшей группой стержней слева/справа вдоль НЛ. Ниже порога - полный r; при равенстве/выше - min(r; t/2). L - расстояние между центрами групп, t - расстояние между их проекциями на НЛ."),
            @("SLS.Crack.SP35.RadiusDiameterMode", "Max", "-", "Принято в программе: Max/Min/Average выбирает диаметр только для радиусов зоны. Сумма βnd всегда использует реальные диаметры всех стержней."),
            @("SLS.Crack.SP35.InteractionRadiusMode", "6d", "-", "Радиус взаимодействия по СП 35: 3d, 5d или 6d. Один вариант применяется в боковом направлении и к нейтральной линии."),
            @("SLS.Crack.SP35.RebarProfile", "Periodic", "-", "Профиль арматуры для размерного коэффициента раскрытия ψ по СП 35: Periodic - периодический; Smooth - гладкий.")
        )},
        @{ Name = "AutoCADSettings"; Title = "[AutoCAD]"; Rows = @(
            @("[Общие настройки]", "", "", ""),
            @("AutoCAD.Common.ConcreteLayer", "Concrete", "-", "Единый слой бетонных элементов для импорта и экспорта. Импортируются Region; экспортируется расчетная геометрия с цветом по физическому состоянию из Results."),
            @("AutoCAD.Common.RebarLayer", "Reinf", "-", "Единый слой продольной арматуры для импорта и экспорта. Импортируются Region; цвет экспортируемых стержней определяется их физическим состоянием из Results."),
            @("AutoCAD.Common.SectionContourLayer", "RC_NDM_Contour", "-", "Общий слой точной наружной границы для импорта и экспорта. Отсутствующий или пустой слой допустим: без контура используются границы бетонных элементов."),
            @("AutoCAD.Common.OpeningContourLayer", "RC_NDM_Opening", "-", "Общий слой отдельных отверстий для импорта и экспорта. Задается независимо от наружного контура; без него отверстия вычитаются из бетонной сетки. Допускает несколько контуров; отсутствующий или пустой слой пропускается. Координаты всегда в мм."),
            @("[Импорт]", "", "", ""),
            @("[Экспорт]", "", "", ""),
            @("AutoCAD.Export.CombinationID", "Worst", "-", "Какое сочетание экспортировать в AutoCAD: Worst - строка, выбранная в верхней сводке Results по General.WorstCombinationCriterion; либо конкретный CombinationID."),
            @("AutoCAD.Export.NeutralLineEnabled", "Yes", "-", "Выгружать нейтральную линию в AutoCAD: Yes - выводить; No - не выводить."),
            @("AutoCAD.Export.PrincipalAxesMode", "Transformed", "-", "Какие главные центральные оси выгружать в AutoCAD: Transformed - приведенного сечения; Concrete - бетонного; None - не выводить."),
            @("AutoCAD.Export.LoadPointEnabled", "Yes", "-", "Выгружать точку приложения нагрузки в AutoCAD: Yes - выводить; No - не выводить."),
            @("AutoCAD.Export.ContourEnabled", "Yes", "-", "Выгружать параметрический контур сечения в AutoCAD: Yes - вывести контур, если он есть в Results; No - не выводить."),
            @("AutoCAD.Export.ContourFormat", "Region", "-", "Формат контуров сечения: Region - бетонная область со всеми отверстиями; Polyline - отдельные замкнутые полилинии наружной границы и отверстий. Не меняет расчетные контуры трещин."),
            @("AutoCAD.Export.ExportCrackInteractionContour", "Yes", "-", "Вывести фактически полученную площадь взаимодействия при выбранном профилем PreCrackState, PostCrackState или CrackedState. Нового расчета при экспорте нет."),
            @("AutoCAD.Export.LabelMode", "NamesAndValues", "-", "Формат текстовых подписей при выгрузке в AutoCAD: ValuesOnly - только выбранная величина профиля; NamesAndValues - имя элемента и значение. При NamesAndValues имена выводятся для бетона и арматуры."),
            @("AutoCAD.Export.CrackInteractionLayer", "RC_NDM_CrackArea", "-", "Экспортный слой рассчитанной области нормальных трещин по СП 35 или СП 63: наружные контуры и отверстия. При импорте не используется."),
            @("AutoCAD.Layer.ConcreteTension", "Anno_Concrete_Tension", "-", "Слой для подписей бетонных волокон, которые физически находятся в растяжении."),
            @("AutoCAD.Layer.ConcreteCompression", "Anno_Concrete_Compression", "-", "Слой для подписей бетонных волокон, которые физически находятся в сжатии; NearZero и выключенное растяжение идут сюда нейтральным цветом."),
            @("AutoCAD.Layer.RebarTension", "Anno_Rebar_Tension", "-", "Слой для подписей стержней, которые физически находятся в растяжении."),
            @("AutoCAD.Layer.RebarCompression", "Anno_Rebar_Compression", "-", "Слой для подписей стержней, которые физически находятся в сжатии; NearZero идет сюда нейтральным цветом."),
            @("AutoCAD.Color.ConcreteTension", "9", "ColorIndex", "AutoCAD ColorIndex для растянутых бетонных областей и их подписей."),
            @("AutoCAD.Color.ConcreteCompression", "5", "ColorIndex", "AutoCAD ColorIndex для сжатых бетонных областей и их подписей."),
            @("AutoCAD.Color.RebarTension", "1", "ColorIndex", "AutoCAD ColorIndex для растянутых стержней и их подписей."),
            @("AutoCAD.Color.RebarCompression", "6", "ColorIndex", "AutoCAD ColorIndex для сжатых стержней и их подписей."),
            @("AutoCAD.Color.Neutral", "8", "ColorIndex", "AutoCAD ColorIndex для областей и подписей с нулевыми или почти нулевыми напряжениями бетона и арматуры. Слой подписи при этом остается compression-слоем соответствующего материала.")
        )},
        @{ Name = "PlotSettings"; Title = "[Схема сечения]"; Rows = @(
            @("Plot.Enabled", "Yes", "-", "Включить построение схемы сечения на листе Расчет по данным последнего расчета на листе Results."),
            @("Plot.AutoUpdateAfterCalculation", "Yes", "-", "При Plot.Enabled=Yes обязателен выбор Yes/No: обновлять схему после расчета или сохранять готовую схему. При Plot.Enabled=No настройка не используется. Кнопка обновления читает только Results."),
            @("Plot.LoadCase", "Worst", "-", "Какое сочетание показывать на схеме: Worst - строка, выбранная в верхней сводке Results по General.WorstCombinationCriterion; либо конкретный CombinationID."),
            @("Plot.ResultGradient", "Yes", "-", "Включить цветовое различение результата на схеме. Цвет физического состояния берется из PhysicalState, записанного в Results."),
            @("Plot.ResultLabelsEnabled", "No", "-", "Показывать пространственно распределенные численные подписи величины Visualization.Quantity выбранного профиля для бетонных элементов."),
            @("Plot.ResultLabelSpacing", "100", "мм", "Минимальный пространственный шаг между численными подписями результата на схеме."),
            @("Plot.NeutralLineEnabled", "Yes", "-", "Показывать нейтральную линию выбранного сочетания по ε_0, κ_x, κ_y из Results."),
            @("Plot.PrincipalAxesMode", "Transformed", "-", "Какие главные центральные оси показывать на схеме: Transformed - приведенного сечения; Concrete - бетонного; None - не показывать."),
            @("Plot.LoadApplicationPointEnabled", "Yes", "-", "Показывать точку приложения нагрузки из последнего расчета."),
            @("Plot.AxisLabelsEnabled", "Yes", "-", "Показывать маленькие курсивные подписи глобальных осей X и Y внутри рабочей области схемы."),
            @("Plot.AxisLabelsFontSize", "7.5", "pt", "Размер шрифта подписей осей X и Y на схеме в пунктах Excel."),
            @("Plot.ContourEnabled", "Yes", "-", "Показывать параметрический контур сечения на схеме, если он сохранен в Results. Для импортированной сетки контур не восстанавливается."),
            @("Plot.LegendEnabled", "Yes", "-", "Показывать легенду физического состояния и выбранной величины профиля справа от схемы."),
            @("Plot.LegendMode", "Separate", "-", "Separate - отдельные легенды для арматуры и бетона; Common - одна общая легенда. В режиме Common используются цвета Plot.Color.RebarCompression и Plot.Color.RebarTension."),
            @("Plot.Color.RebarCompression", "30,80,220", "RGB", "Цвет максимального сжатия арматуры. В режиме Plot.LegendMode=Common используется как цвет сжатия для всех элементов."),
            @("Plot.Color.RebarTension", "210,30,20", "RGB", "Цвет максимального растяжения арматуры. В режиме Plot.LegendMode=Common используется как цвет растяжения для всех элементов."),
            @("Plot.Color.ConcreteCompression", "125,35,210", "RGB", "Цвет максимального сжатия бетона при Plot.LegendMode=Separate."),
            @("Plot.Color.ConcreteTension", "0,155,85", "RGB", "Цвет максимального растяжения бетона при Plot.LegendMode=Separate; неработающий растянутый бетон остается серым.")
        )}
    )
}

# Задаёт начальные INPUT/INTERNAL/OUTPUT единицы длины, площади, силы, момента, напряжения и кривизны для таблицы Config.
function Get-UnitSettingsCatalog {
    @(
        @("Length", "mm", "mm", "m"),
        @("Area", "mm2", "mm2", "m2"),
        @("Force", "tf", "N", "tf"),
        @("Moment", "tf*m", "N*mm", "tf*m"),
        @("Stress", "MPa", "MPa", "MPa"),
        @("Curvature", "1/mm", "1/mm", "1/mm")
    )
}

# Задаёт пользовательские и внутренние знаки N, Mx и My для таблицы системы знаков Config.
function Get-SignConventionSettingsCatalog {
    @(
        @("+N", "Compression", "Tension"),
        @("+Mx", "+Y tension", "+Y tension"),
        @("+My", "+X tension", "+X tension")
    )
}

# Возвращает параметры бетона, из которых код строит диаграммы I/II ГПС.
# Серые вычисляемые точки диаграмм выводятся отдельно и не являются вводом.
function Get-ConcreteMaterialParametersCatalog {
    @(
        @("Concrete.R.ULS(I)", "15.5", "1.10", "МПа", "I ГПС по СП: Сжатие = R_b, Растяжение = R_bt. Пользователь вводит итоговые расчетные сопротивления с нужными коэффициентами условий работы. Оба сопротивления вводите положительными, без знака минус для сжатия."),
        @("Concrete.R.SLS(II)", "22.0", "1.80", "МПа", "II ГПС по СП: Сжатие = R_b,ser, Растяжение = R_bt,ser. Коэффициенты условий работы учитываются пользователем до ввода. Оба сопротивления вводите положительными, без знака минус для сжатия."),
        @("Concrete.Rb.mc2", "14.6", "-", "МПа", "Допустимое осевое сжатие бетона для проверки продольных трещин по СП 35; значение вводится уже с учетом m_b13 и других нужных коэффициентов. Вводите положительное значение, без знака минус."),
        @("Concrete.E", "32500", "32500", "МПа", "Начальный модуль бетона: E_b при сжатии и E_bt при растяжении; по СП 63 E_bt = E_b. Оба модуля вводите положительными."),
        @("Concrete.TwoLine.Eb1Red", "0.0015", "0.00008", "-", "Для TwoLine: ε_b1,red при сжатии и ε_bt1,red при растяжении. Обе деформации вводите положительными по модулю, без знака минус для сжатия."),
        @("Concrete.ThreeLine.Eb0", "0.002", "0.0001", "-", "Для ThreeLine: ε_b0 и ε_bt0 - деформации начала горизонтального участка. Обе деформации вводите положительными по модулю, без знака минус для сжатия."),
        @("Concrete.TwoThreeLine.Eb2", "0.0035", "0.00015", "-", "Предельные деформации ε_b2 и ε_bt2 для TwoLine и ThreeLine. Обе деформации вводите положительными по модулю, без знака минус для сжатия.")
    )
}

# Возвращает параметры арматуры, из которых строятся диаграммы I/II ГПС.
# Rsc,ser и Esc оставлены как программное расширение для симметрии.
function Get-SteelMaterialParametersCatalog {
    @(
        @("Steel.R.ULS(I)", "350", "350", "МПа", "I ГПС по СП: Сжатие = R_sc, Растяжение = R_s. Оба сопротивления вводите положительными, без знака минус для сжатия."),
        @("Steel.R.SLS(II)", "390", "390", "МПа", "II ГПС: Сжатие = R_sc,ser, Растяжение = R_s,ser. R_sc,ser — отдельный параметр сжатой ветви программы. Оба сопротивления вводите положительными, без знака минус для сжатия."),
        @("Steel.E", "200000", "200000", "МПа", "Модуль арматуры: E_sc при сжатии и E_s при растяжении; по СП 63 модуль принимается одинаковым. Оба модуля вводите положительными."),
        @("Steel.TwoLine.Es2", "0.025", "0.025", "-", "Для TwoLine: предельные деформации арматуры ε_sc2 при сжатии и ε_s2 при растяжении. Обе деформации вводите положительными по модулю, без знака минус для сжатия."),
        @("Steel.ThreeLine.Es2", "0.015", "0.015", "-", "Для ThreeLine: предельные деформации арматуры ε_sc2 при сжатии и ε_s2 при растяжении. Обе деформации вводите положительными по модулю, без знака минус для сжатия.")
    )
}

# Возвращает вертикальную таблицу расчетных профилей.
# Строки описывают frontend-контракт профиля, а столбцы PR1/PR2/... являются
# стабильными ProfileId, которые назначаются сочетаниям нагрузок.
function Get-CalculationProfilesCatalog {
    @(
        @{ Caption = "[Общее]"; Key = ""; PR1 = ""; PR2 = ""; PR3 = ""; PR4 = ""; Comment = "" },
        @{ Caption = "Имя профиля"; Key = "Profile.DisplayName"; PR1 = "PR1"; PR2 = "PR2"; PR3 = "PR3"; PR4 = "PR4"; Comment = "Короткое имя профиля для пользователя; можно переименовать без изменения технического ID профиля." },
        @{ Caption = "Описание"; Key = "Profile.Description"; PR1 = "Прочность, несущая способность и устойчивость"; PR2 = "Расчет раскрытия нормальных и продольных трещин"; PR3 = "Прочность, capacity и трещины"; PR4 = "Только прямое НДС по прочности"; Comment = "Пояснение, что делает профиль." },

        @{ Caption = "[Запрашиваемые расчеты]"; Key = ""; PR1 = ""; PR2 = ""; PR3 = ""; PR4 = ""; Comment = "" },
        @{ Caption = "НДС по прочности"; Key = "Calculation.Strength.DirectState"; PR1 = "Yes"; PR2 = "No"; PR3 = "Yes"; PR4 = "Yes"; Comment = "Состояние НДС по модели прочности." },
        @{ Caption = "Несущая способность"; Key = "Calculation.Strength.Capacity"; PR1 = "Yes"; PR2 = "No"; PR3 = "Yes"; PR4 = "No"; Comment = "Поиск предельной точки по LoadPath сочетания." },
        @{ Caption = "Раскрытие трещин"; Key = "Calculation.Crack.Width"; PR1 = "No"; PR2 = "Yes"; PR3 = "Yes"; PR4 = "No"; Comment = "Нормальные трещины и проверка продольных трещин." },
        @{ Caption = "Продольный изгиб"; Key = "Calculation.Stability.Enabled"; PR1 = "Yes"; PR2 = "No"; PR3 = "No"; PR4 = "No"; Comment = "Включает расчет продольного изгиба и устойчивости перед остальными проверками профиля." },

        @{ Caption = "[Модель устойчивости]"; Key = ""; PR1 = ""; PR2 = ""; PR3 = ""; PR4 = ""; Comment = "" },
        @{ Caption = "Характеристики материалов"; Key = "MaterialModel.Stability.ValueSet"; PR1 = "ULS(I)"; PR2 = "ULS(I)"; PR3 = "ULS(I)"; PR4 = "ULS(I)"; Comment = "Для расчета устойчивости по СП применяются расчетные характеристики I ГПС: бетон Rb/Rbt, арматура Rsc/Rs." },

        @{ Caption = "[Модель прочности]"; Key = ""; PR1 = ""; PR2 = ""; PR3 = ""; PR4 = ""; Comment = "" },
        @{ Caption = "Характеристики материалов"; Key = "MaterialModel.Strength.ValueSet"; PR1 = "ULS(I)"; PR2 = "ULS(I)"; PR3 = "ULS(I)"; PR4 = "ULS(I)"; Comment = "Для прочности по СП используются характеристики I ГПС: бетон Rb/Rbt, арматура Rsc/Rs." },
        @{ Caption = "Диаграмма бетона"; Key = "MaterialModel.Strength.ConcreteDiagram"; PR1 = "ThreeLine"; PR2 = "ThreeLine"; PR3 = "ThreeLine"; PR4 = "ThreeLine"; Comment = "СП 63, п. 6.1.23: для прочности применяется двух- или трехлинейная диаграмма бетона." },
        @{ Caption = "Растянутый бетон"; Key = "MaterialModel.Strength.ConcreteTension"; PR1 = "Ignore"; PR2 = "Ignore"; PR3 = "Ignore"; PR4 = "Ignore"; Comment = "СП 63, п. 8.1.20: при расчете прочности растянутый бетон допускается не учитывать." },
        @{ Caption = "Диаграмма арматуры"; Key = "MaterialModel.Strength.SteelDiagram"; PR1 = "TwoLine"; PR2 = "TwoLine"; PR3 = "TwoLine"; PR4 = "TwoLine"; Comment = "СП 63, п. 6.2.13: TwoLine для физического предела текучести, ThreeLine для условного." },

        @{ Caption = "[Модель начала трещинообразования]"; Key = ""; PR1 = ""; PR2 = ""; PR3 = ""; PR4 = ""; Comment = "" },
        @{ Caption = "Характеристики материалов"; Key = "MaterialModel.CrackInitiation.ValueSet"; PR1 = "SLS(II)"; PR2 = "SLS(II)"; PR3 = "SLS(II)"; PR4 = "SLS(II)"; Comment = "Для проверки образования нормальной трещины используются характеристики II ГПС: Rb,ser/Rbt,ser и Rs,ser." },
        @{ Caption = "Диаграмма бетона"; Key = "MaterialModel.CrackInitiation.ConcreteDiagram"; PR1 = "ThreeLine"; PR2 = "ThreeLine"; PR3 = "ThreeLine"; PR4 = "ThreeLine"; Comment = "СП 63, п. 6.1.24: для образования трещин основная модель бетона - ThreeLine с растяжением." },
        @{ Caption = "Растянутый бетон"; Key = "MaterialModel.CrackInitiation.ConcreteTension"; PR1 = "UseDiagram"; PR2 = "UseDiagram"; PR3 = "UseDiagram"; PR4 = "UseDiagram"; Comment = "Для стадии образования трещины растянутая ветвь бетона должна учитываться; иначе критерий εbt,ult не определяется по этой модели." },
        @{ Caption = "Диаграмма арматуры"; Key = "MaterialModel.CrackInitiation.SteelDiagram"; PR1 = "TwoLine"; PR2 = "TwoLine"; PR3 = "TwoLine"; PR4 = "TwoLine"; Comment = "СП 63, п. 6.2.13: TwoLine для физического предела текучести, ThreeLine для условного." },

        @{ Caption = "[Модель НДС с раскрытыми трещинами]"; Key = ""; PR1 = ""; PR2 = ""; PR3 = ""; PR4 = ""; Comment = "" },
        @{ Caption = "Характеристики материалов"; Key = "MaterialModel.CrackedState.ValueSet"; PR1 = "SLS(II)"; PR2 = "SLS(II)"; PR3 = "SLS(II)"; PR4 = "SLS(II)"; Comment = "Для раскрытия трещин по СП используются характеристики II ГПС." },
        @{ Caption = "Диаграмма бетона"; Key = "MaterialModel.CrackedState.ConcreteDiagram"; PR1 = "TwoLine"; PR2 = "TwoLine"; PR3 = "TwoLine"; PR4 = "TwoLine"; Comment = "СП 63, п. 6.1.26: после образования трещин НДС допускается считать по TwoLine или ThreeLine." },
        @{ Caption = "Растянутый бетон"; Key = "MaterialModel.CrackedState.ConcreteTension"; PR1 = "Ignore"; PR2 = "Ignore"; PR3 = "Ignore"; PR4 = "Ignore"; Comment = "Для уже образовавшейся трещины растянутый бетон в НДС не учитывается." },
        @{ Caption = "Диаграмма арматуры"; Key = "MaterialModel.CrackedState.SteelDiagram"; PR1 = "TwoLine"; PR2 = "TwoLine"; PR3 = "TwoLine"; PR4 = "TwoLine"; Comment = "СП 63, п. 6.2.13: TwoLine для физического предела текучести, ThreeLine для условного." },

        @{ Caption = "[Настройки визуализации]"; Key = ""; PR1 = ""; PR2 = ""; PR3 = ""; PR4 = ""; Comment = "" },
        @{ Caption = "Выводимое состояние"; Key = "Visualization.State"; PR1 = "StrengthState"; PR2 = "CrackedState"; PR3 = "StrengthState"; PR4 = "StrengthState"; Comment = "Какое расчетное состояние показывать на схеме и в AutoCAD." },
        @{ Caption = "Выводимая величина"; Key = "Visualization.Quantity"; PR1 = "Stress"; PR2 = "Stress"; PR3 = "Stress"; PR4 = "Strain"; Comment = "Что показывать: напряжения Stress или деформации Strain." },
        @{ Caption = "Точность Stress"; Key = "Visualization.StressPrecision"; PR1 = "1"; PR2 = "1"; PR3 = "1"; PR4 = "1"; Comment = "Знаки после запятой для подписей напряжений на схеме и в AutoCAD; влияет только на текст, не на Results." },
        @{ Caption = "Точность Strain"; Key = "Visualization.StrainPrecision"; PR1 = "6"; PR2 = "6"; PR3 = "6"; PR4 = "6"; Comment = "Знаки после запятой для подписей деформаций на схеме и в AutoCAD; удобно оставлять больше знаков, чем для Stress." }
    )
}

# Подключает построение таблиц сечений и общий сохранённый макет книги.
. (Join-Path $PSScriptRoot 'GeometryTables.ps1')

# Преобразует положительный номер столбца в буквенный адрес Excel, включая многобуквенные имена.
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

# Создаёт или обновляет имя книги на прямоугольник по четырём границам строк/столбцов; экранирует имя листа.
function Set-WorkbookNameByBounds {
    param([object]$Workbook, [string]$Name, [object]$Sheet, [int]$FirstRow, [int]$FirstColumn, [int]$LastRow, [int]$LastColumn)
    $sheetName = ([string]$Sheet.Name).Replace("'", "''")
    $firstCol = ConvertTo-ExcelColumn $FirstColumn
    $lastCol = ConvertTo-ExcelColumn $LastColumn
    $rangeAddress = "$" + $firstCol + "$" + $FirstRow + ":$" + $lastCol + "$" + $LastRow
    $address = "='" + $sheetName + "'!" + $rangeAddress
    try { $Workbook.Names.Item($Name).RefersTo = $address }
    catch {
        try { $Workbook.Names.Add($Name, $address) | Out-Null }
        catch { throw "Не удалось назначить имя '$Name' диапазону '$address': $($_.Exception.Message)" }
    }
}

# Возвращает именованные диапазоны, которые физически живут на Config и
# должны выглядеть как отдельные редактируемые таблицы.
function Get-ConfigNamedRangeNames {
    @(
        "rngSystemSettings",
        "rngUnitSettings",
        "rngSignConventionSettings",
        "rngCircleGeometry",
        "rngRectanglesGeometry",
        "rngPointedSectionGeometry",
        "rngHollowRectangleGeometry",
        "rngConcreteMaterialParameters",
        "rngSteelMaterialParameters",
        "rngCalculationProfiles",
        "rngPlotAnnotationSettings",
        "rngLoadCombinations",
        "rngStabilityDurationLoads",
        "rngSP35Table721"
    )
}

# Обводит каждую ячейку таблиц Config тонкой пунктирной рамкой.
# Это визуально показывает пользователю фактические границы именованных
# диапазонов без отдельной толстой линии в конце таблицы.
function Apply-ConfigNamedRangeBorders {
    param([object]$Workbook)

    foreach ($rangeName in (Get-ConfigNamedRangeNames)) {
        try {
            $range = $Workbook.Names.Item($rangeName).RefersToRange
        } catch {
            # Диапазон может отсутствовать в промежуточной сборке; валидатор
            # отдельно проверит обязательные имена книги.
            continue
        }
        if($rangeName -in @('rngCircleGeometry','rngRectanglesGeometry','rngPointedSectionGeometry','rngHollowRectangleGeometry')){
            Format-GeometrySettingsTable $Workbook $rangeName $range
            continue
        }

        $range.Borders.LineStyle = -4115
        $range.Borders.Weight = 2
        $range.Borders.Color = 0
        # После Hyperlink-стиля Excel общая коллекция Borders не всегда
        # восстанавливает внешнюю грань отдельных ячеек Справка.
        $rightEdge = $range.Columns.Item($range.Columns.Count).Borders.Item(10)
        $rightEdge.Weight = 2
        $rightEdge.Color = 0
        $rightEdge.LineStyle = -4115
    }
}

# Начинает распределение служебных списков Config для одной сборки листа.
# Все таблицы используют этот реестр: добавление новых настроек не может
# перезаписать списки геометрии, единиц или параметров схемы.
function Reset-ConfigValidationSources {
    param([object]$Sheet)
    $script:ConfigValidationSheet = $Sheet
    $script:ConfigValidationNextColumn = 70
    if($null -ne $script:WorkbookLayout){$last=($script:WorkbookLayout.HelperColumns.PSObject.Properties.Value | Measure-Object -Maximum).Maximum;$script:ConfigValidationNextColumn=[int]$last+1}
    $script:ConfigValidationLists = @{}
}

# Выделяет отдельную скрытую колонку для статического или динамического
# списка. Очищается только выделяемый служебный диапазон, не пользовательские таблицы.
function Get-NextConfigValidationColumn {
    param([object]$Sheet,[string]$SourceKey='')
    if ($null -eq $script:ConfigValidationSheet -or
        $script:ConfigValidationSheet.Parent.Name -ne $Sheet.Parent.Name -or
        $script:ConfigValidationSheet.Name -ne $Sheet.Name) {
        throw 'Служебные списки Config не инициализированы для этой книги.'
    }
    $column = Get-WorkbookLayoutHelperColumn $SourceKey
    if($column -eq 0){$column=$script:ConfigValidationNextColumn;$script:ConfigValidationNextColumn++}
    if ($column -gt $Sheet.Columns.Count) { throw 'Недостаточно колонок для служебных списков Config.' }
    $Sheet.Range($Sheet.Cells.Item(1, $column), $Sheet.Cells.Item(501, $column)).ClearContents() | Out-Null
    $Sheet.Columns.Item($column).Hidden = $true
    return $column
}

# Возвращает прямую ссылку на вертикальный список допустимых значений.
# Короткие постоянные наборы хранятся в проверке данных; большие и динамические используют общий источник.
function Get-ConfigValidationListAddress {
    param([object]$Sheet, [array]$Options, [string]$IndependentName = '')
    if ($Options.Count -eq 0) { throw 'Пустой список допустимых значений Config.' }
    $inline=@($Options | ForEach-Object {[string]$_}) -join ','
    if (-not $IndependentName -and $inline.Length -le 255 -and -not @($Options | Where-Object {[string]$_ -match '[,;\r\n]'}).Count) { return ($Options -join [string]$Sheet.Application.International(5)) }
    $key = ConvertTo-Json -InputObject @($Options) -Compress
    if ($IndependentName) { $key = $IndependentName + ':' + $key }
    if ($script:ConfigValidationLists.ContainsKey($key)) { return $script:ConfigValidationLists[$key] }
    $column = Get-NextConfigValidationColumn $Sheet $IndependentName
    for ($i = 0; $i -lt $Options.Count; $i++) {
        $Sheet.Cells.Item($i + 1, $column).Value2 = [string]$Options[$i]
    }
    $colName = ConvertTo-ExcelColumn $column
    $address = '=$' + $colName + '$1:$' + $colName + '$' + $Options.Count
    if ($IndependentName) {
        Set-WorkbookNameByBounds $Sheet.Parent $IndependentName $Sheet 1 $column $Options.Count $column
        $address = '=' + $IndependentName
    }
    $script:ConfigValidationLists[$key] = $address
    return $address
}

# Возвращает фактические таблицы форматирования внутри именованных диапазонов.
# Составная геометрия содержит несколько таблиц с разными колонками: правило
# общего комментария не должно затрагивать расположенные ниже числа и selectors.
function Get-ConfigFormattingTables {
    param([object]$Workbook)
    foreach($name in @(Get-ConfigNamedRangeNames)){
        try{$range=$Workbook.Names.Item($name).RefersToRange}catch{continue}
        if([string]$range.Worksheet.Name -notin @('Config','Сечения')){continue}
        if($name -in @((Get-GeometrySettingsCatalog).RangeName)){
            foreach($block in @(Get-GeometryTableBlocks $range)){
                [pscustomobject]@{Name=$name;Range=$range.Cells.Item($block.HeaderRow,1).Resize($block.LastRow-$block.HeaderRow+1,[int]$block.Columns['note']);Geometry=$true}
            }
        }else{[pscustomobject]@{Name=$name;Range=$range;Geometry=$false}}
    }
}

# Выравнивает комментарии в пределах их собственной таблицы, а не всей
# составной геометрии. В геометрии, армировании и профилях комментарии справа;
# в остальных таблицах — слева, без переноса, строка 15 пунктов.
function Apply-ConfigCommentColumnAlignment {
    param([object]$Workbook)

    foreach ($table in (Get-ConfigFormattingTables $Workbook)) {
        try {
            $range = $table.Range
            if ($range.Rows.Count -lt 2) { continue }

            for ($c = 1; $c -le $range.Columns.Count; $c++) {
                $header = [string]$range.Cells.Item(1, $c).Value2
                $headerRow = 1
                if (($header -ne "Комментарий") -and ($header -ne "Комментарии") -and ($header -ne "Comment") -and ($range.Rows.Count -ge 2)) {
                    $header = [string]$range.Cells.Item(2, $c).Value2
                    $headerRow = 2
                }
                if ($header.ToLowerInvariant() -eq "комментарий" -or $header -eq "Комментарии" -or $header -eq "Comment") {
                    for($r=$headerRow;$r -le $range.Rows.Count;$r++){
                        $cell=$range.Worksheet.Cells.Item($range.Row+$r-1,$range.Column+$c-1)
                        $textRange=$cell.MergeArea
                        if($textRange.Column -ne $cell.Column){continue} # Это заголовок раздела, а не ячейка комментария.
                        $textRange.HorizontalAlignment=$(if($table.Geometry -or $table.Name -eq 'rngCalculationProfiles'){-4152}else{-4131})
                        $textRange.WrapText=$false
                        if(-not $table.Geometry){$cell.EntireRow.RowHeight=[double]15}
                    }
                }
            }
        } catch {
            # Форматирование не должно скрывать реальную ошибку сборки книги:
            # обязательные диапазоны проверяются отдельным Validate-Workbook.
        }
    }
}

# Центрирует только те ячейки Config, куда пользователь вводит значение или
# выбирает вариант из выпадающего списка. Заголовки, единицы, комментарии и
# ссылки на справку сохраняют собственное выравнивание.
function Set-ConfigInputCellAlignment {
    param([object]$Range)

    if ($null -eq $Range) { return }
    $Range.HorizontalAlignment = -4108
    $Range.VerticalAlignment = -4108
}

# Центрирует короткие служебные колонки Config вместе с их заголовками.
# Такие поля читаются как компактные метки или ссылки, поэтому центральное
# выравнивание выглядит аккуратнее, чем текстовое прижатие к левому краю.
function Apply-ConfigShortColumnAlignment {
    param([object]$Workbook)

    foreach ($table in (Get-ConfigFormattingTables $Workbook)) {
        try {
            $range = $table.Range

            for ($c = 1; $c -le $range.Columns.Count; $c++) {
                $header = [string]$range.Cells.Item(1, $c).Value2
                if ($header -eq "Ед." -or $header -eq "ед" -or $header -eq "Справка" -or $header -eq "INTERNAL") {
                    Set-ConfigInputCellAlignment $range.Offset(0, $c - 1).Resize($range.Rows.Count, 1)
                } elseif ($header -eq "Значение") {
                    Set-ConfigInputCellAlignment $range.Cells.Item(1, $c)
                }
            }
        } catch {
            # Форматирование коротких колонок не должно прерывать сборку книги.
        }
    }
}

# Применяет единое правило выравнивания пользовательских ячеек ввода на Config.
# Список колонок задан явно по именованным диапазонам, чтобы не центрировать
# описательные поля и не ломать читаемость комментариев.
function Apply-ConfigUserInputAlignment {
    param([object]$Workbook)

    $items = @(
        @{ Name = "rngSystemSettings"; Columns = @(2) },
        @{ Name = "rngUnitSettings"; Columns = @(2, 4) },
        @{ Name = "rngSignConventionSettings"; Columns = @(2) },
        @{ Name = "rngSteelMaterialParameters"; Columns = @(2, 3) },
        @{ Name = "rngConcreteMaterialParameters"; Columns = @(2, 3) },
        @{ Name = "rngCalculationProfiles"; Columns = @(3, 4, 5, 6) },
        @{ Name = "rngPlotAnnotationSettings"; Columns = @(2, 3) },
        @{ Name = "rngLoadCombinations"; Columns = @(1, 2, 3, 4, 5, 6, 7); CommentColumns = @(7) },
        @{ Name = "rngStabilityDurationLoads"; Columns = @(1, 2, 3, 4) },
        @{ Name = "rngSP35Table721"; Columns = @(1, 2, 3, 4, 5, 6, 7, 8) }
    )

    foreach ($item in $items) {
        try {
            $range = $Workbook.Names.Item($item.Name).RefersToRange
            if ([string]$range.Worksheet.Name -ne "Config") { continue }
            if ($range.Rows.Count -lt 2) { continue }

            foreach ($columnIndex in $item.Columns) {
                if($columnIndex -notin $item.CommentColumns){
                    Set-ConfigInputCellAlignment $range.Offset(1, $columnIndex - 1).Resize($range.Rows.Count - 1, 1)
                }
                for($r=2;$r -le $range.Rows.Count;$r++){
                    $rowKey=[string]$range.Cells.Item($r,1).Value2
                    if($rowKey.StartsWith('[') -or ($item.Name -eq 'rngSystemSettings' -and -not $rowKey.Contains('.'))){continue}
                    $cell=$range.Cells.Item($r,$columnIndex)
                    $inputArea=$cell
                    if($cell.MergeCells){
                        $inputArea=$cell.MergeArea
                        if($inputArea.Row -ne $cell.Row -or $inputArea.Column -ne $cell.Column){continue}
                    }
                    if($columnIndex -in $item.CommentColumns -or (-not $cell.HasFormula -and [string]$cell.Value2 -notin @('-','—'))){$inputArea.Interior.Color=15135487}
                }
            }
        } catch {
            # Диапазон может отсутствовать в промежуточной сборке; валидатор
            # отдельно проверит обязательные имена книги.
        }
    }

    foreach($name in @((Get-GeometrySettingsCatalog).RangeName)){
        $range=$Workbook.Names.Item($name).RefersToRange
        Set-GeometryEditableCellFormat $range
        foreach($block in @(Get-GeometryTableBlocks $range)){
            foreach($key in @($block.Columns.Keys)){
                if($key -in @('key','name','unit','note')){continue}
                Set-ConfigInputCellAlignment $range.Cells.Item($block.FirstRow,[int]$block.Columns[$key]).Resize($block.LastRow-$block.FirstRow+1,1)
            }
        }
    }
}

# Возвращает расширенное описание настройки для листа "Справка".
function Get-SettingInstructionLines {
    param([string]$Key, [string]$Comment)

    $lead = Get-InstructionLead $Key $Comment

    switch -Wildcard ($Key) {
        "Mesh.StepY" { return @(
            "При Mesh.StepMode = User задайте положительный шаг по Y в INPUT-единице длины; он не копируется из Mesh.StepX. В Auto это поле сохраняется, но не используется. Общие правила выбора шага и точности сетки приведены в разделе Mesh.StepX.",
            "",
            "",
            "",
            "",
            "",
            ""
        ) }
        "Load.ReferenceOffsetY" { return @(
            "Смещение по Y отсчитывается от центра тяжести бетона и задаётся независимо от смещения по X. Допустимые значения, единицы и перенос моментов описаны в разделе Load.ReferenceOffsetX.",
            "",
            "",
            "",
            "",
            ""
        ) }
        "Stability.Mu2" { return @(
            "Укажите положительный безразмерный коэффициент для главной плоскости 2; у круга она соответствует изгибу My. Общие ограничения ввода и связь с расчётной длиной описаны в разделе Stability.Mu1; коэффициенты двух плоскостей задаются независимо.",
            "",
            "",
            "",
            "",
            "",
            ""
        ) }
        "Stability.ZeroMomentEccentricitySign2" { return @(
            "Укажите 1 или −1 для направления случайного эксцентриситета при нулевом моменте во второй главной плоскости; для круга это направление по X при изгибе My. Общий порядок применения описан в разделе Stability.ZeroMomentEccentricitySign1; знак второй плоскости задаётся отдельно.",
            "",
            "",
            "",
            "",
            ""
        ) }
        "Stability.AccidentalEccentricityUser2" { return @(
            "В режиме User задайте положительный модуль эксцентриситета второй главной плоскости в INPUT-единице длины; для круга это плечо по X при изгибе My. Общие ограничения ввода, в том числе для автоматических режимов, описаны в разделе Stability.AccidentalEccentricityUser1; направление задаётся отдельно от модуля.",
            "",
            "",
            "",
            "",
            "",
            ""
        ) }
        "AutoCAD.Layer.ConcreteCompression" { return @(
            "Здесь задаётся отдельное имя слоя подписей сжатого бетона. Ограничения имени, создание и очистка слоёв описаны в разделе AutoCAD.Layer.ConcreteTension.",
            "",
            "",
            "",
            "",
            ""
        ) }
        "AutoCAD.Layer.RebarTension" { return @(
            "Здесь задаётся отдельное имя слоя подписей растянутых стержней арматуры. Ограничения имени, создание и очистка слоёв описаны в разделе AutoCAD.Layer.ConcreteTension.",
            "",
            "",
            "",
            "",
            ""
        ) }
        "AutoCAD.Layer.RebarCompression" { return @(
            "Здесь задаётся отдельное имя слоя подписей сжатых стержней арматуры. Ограничения имени, создание и очистка слоёв описаны в разделе AutoCAD.Layer.ConcreteTension.",
            "",
            "",
            "",
            "",
            ""
        ) }
        "AutoCAD.Color.ConcreteCompression" { return @(
            "Укажите собственный цвет сжатого бетона и его подписей. Формат ACI, допустимые индексы и выбор физического состояния описаны в разделе AutoCAD.Color.ConcreteTension.",
            "",
            ""
        ) }
        "AutoCAD.Color.RebarTension" { return @(
            "Укажите собственный цвет растянутых стержней арматуры и их подписей. Формат ACI, допустимые индексы и выбор физического состояния описаны в разделе AutoCAD.Color.ConcreteTension.",
            "",
            ""
        ) }
        "AutoCAD.Color.RebarCompression" { return @(
            "Укажите собственный цвет сжатых стержней арматуры и их подписей. Формат ACI, допустимые индексы и выбор физического состояния описаны в разделе AutoCAD.Color.ConcreteTension.",
            "",
            ""
        ) }
        "AutoCAD.Color.Neutral" { return @(
            "Этот цвет используется для нулевых и почти нулевых напряжений, независимо от цвета сжатия или растяжения. Формат ACI и допустимые индексы описаны в разделе AutoCAD.Color.ConcreteTension.",
            "",
            ""
        ) }
        "Plot.Color.RebarTension" { return @(
            "Укажите собственный цвет крайнего растяжения арматуры; в режиме Common он также задаёт растянутый конец общей шкалы. Формат RGB и ограничения ввода описаны в разделе Plot.Color.RebarCompression.",
            "",
            "",
            ""
        ) }
        "Plot.Color.ConcreteCompression" { return @(
            "Укажите собственный цвет крайнего сжатия бетона для режима Separate. Формат RGB и ограничения ввода описаны в разделе Plot.Color.RebarCompression.",
            "",
            "",
            ""
        ) }
        "Plot.Color.ConcreteTension" { return @(
            "Укажите собственный цвет крайнего растяжения работающего бетона для режима Separate. Формат RGB и ограничения ввода описаны в разделе Plot.Color.RebarCompression.",
            "",
            "",
            ""
        ) }
        "Concrete.ThreeLine.Eb0" { return @($lead) + @(
            "Здесь задаются деформации начала плато трёхлинейной диаграммы, отдельно для сжатия и растяжения. Общие правила ввода деформаций и нормативные ссылки приведены в разделе Concrete.TwoLine.Eb1Red; численные значения разных точек диаграммы задаются независимо.",
            "",
            "",
            ""
        ) }
        "Concrete.TwoThreeLine.Eb2" { return @($lead) + @(
            "Здесь задаются предельные деформации бетонной диаграммы, отдельно для сжатия и растяжения. Общие правила ввода деформаций и нормативные ссылки приведены в разделе Concrete.TwoLine.Eb1Red; эти пределы используются и для TwoLine, и для ThreeLine.",
            "",
            "",
            ""
        ) }
        "Steel.ThreeLine.Es2" { return @($lead) + @(
            "Здесь задаются предельные деформации трёхлинейной диаграммы арматуры, отдельно для сжатия и растяжения. Общие правила ввода и ограничение одной моделью арматуры на сечение описаны в разделе Steel.TwoLine.Es2; пределы двух диаграмм задаются независимо.",
            "",
            ""
        ) }
        "General.ExecutionReportEnabled" { return @(
            "Yes включает подробную диагностику выполнения расчета.",
            "Подробная диагностика обновляется при каждом запуске.",
            "В отчет попадают основные шаги запуска: чтение настроек, построение геометрии, расчет сочетаний, запись Results и обновление схемы.",
            "Для каждого сочетания дополнительно выводятся подробности численного поиска. Объем диагностической информации увеличивается при большом числе сочетаний или долгом поиске решения.",
            "No отключает подробную диагностику. Расчетная логика и результаты при этом не меняются.",
            "Выберите Yes или No. Если строка настройки отсутствует, значение пустое, неизвестное или формула содержит ошибку, программа укажет место и способ исправления на Config. Без корректного выбора действие не выполняется."
        ) }
        "General.NonCriticalMessagesEnabled" { return @(
            "Yes показывает обычные информационные окна после успешных действий: расчет выполнен, схема обновлена, геометрия импортирована, экспорт в AutoCAD завершен.",
            "No отключает информационные окна успешного завершения действий.",
            "Ошибки, предупреждения и сообщения о невыполненном действии выводятся независимо от этой настройки.",
            "Выберите Yes или No. Кнопки проверяют эту настройку до расчета, очистки, перерисовки, импорта или изменения чертежа. Потерянная строка, пустое или ошибочное значение не включает окна по умолчанию: показывается причина и текущее место ошибки в Config."
        ) }
        "General.ReformatResultsEnabled" { return @(
            "Yes заново оформляет таблицы Results и их шапки после формирования сечения и расчета. Это значение по умолчанию.",
            "No сохраняет готовые шрифты, рамки, выравнивание и объединения шапок. Данные, актуальные единицы, цвета расчетных статусов и необходимое оформление новых строк обновляются.",
            "Если диапазон перенесён в пустое место, подписи шапки не совпадают с ожидаемыми или новые строки ещё не оформлены, программа заново оформляет соответствующий блок даже при No. Она проверяет подписи и отдельные признаки готового оформления, а не сохранность каждого шрифта и рамки.",
            "Настройка не отключает расчетные проверки, подробные численные результаты, текстовый отчет или схему. После ручного изменения оформления выберите Yes, чтобы восстановить штатный вид таблиц.",
            "Выберите Yes или No. Пустое, ошибочное значение и отсутствие настройки останавливают действие до изменения Results с указанием места ошибки на Config."
        ) }
        "Results.SnapshotStorage" { return @(
            "Если выбрано Excel, результаты текущего расчёта сохраняются на листе Results. В нижних таблицах остаются состояния элементов, данные сочетаний и аннотации, необходимые для схемы. Это режим по умолчанию.",
            "Если выбрано File, полный снимок расчёта сохраняется в одном файле .ndm рядом с книгой. Он содержит геометрию, использованные диаграммы, состояния элементов и все данные для восстановления результатов НДС, несущей способности, трещин и устойчивости.",
            "В File источник — файл Results.SnapshotName. Для справки на Results обновляются верхние результаты и четыре нижние таблицы: элементы бетона и арматуры, характеристики сечения, достоверные контуры и диаграммы материалов. Удаление этих данных не меняет снимок и не мешает расчёту. Выбор сочетания, состояния и оформление остаются на Config. Состояния, аннотации и границы элементов хранятся в файле.",
            "Кнопка «Обновить схему и результаты» при выбранном File загружает снимок с именем Results.SnapshotName и восстанавливает схему и подробные таблицы без нового расчёта. Так можно переключаться между ранее сохранёнными расчётами разных сечений. Текущие настройки Config и входные нагрузки при загрузке не меняются.",
            "При выбранном Excel расчёт, схема и экспорт читают только Results, без загрузки файла. При выбранном File они сразу читают файл Results.SnapshotName; предварительно нажимать кнопку обновления не требуется. Переключение режима само по себе не переносит снимок между листом и файлом.",
            "Если предварительная оценка превышает 60000 строк состояний или фактических границ элементов, книга предлагает выбрать File, чтобы уменьшить размер книги и объём записи на лист. Можно согласиться или продолжить в Excel. Если таблица физически не помещается на листе, потребуется выбрать File.",
            "«Сформировать сечение» в File также сохраняет точную геометрию в .ndm. Такой файл явно помечен GeometryPreview, не содержит результатов НДС или сочетаний и позволяет после повторного открытия построить схему, экспортировать геометрию и выполнить расчёт по сохранённому сечению.",
            "Файл читается самой книгой без установки дополнительной программы. Для работы нужны настольный Excel для Windows и разрешённые макросы. При переносе книги на другой компьютер сохраните рядом нужные файлы снимков."
        ) }
        "Results.SnapshotName" { return @(
            "Укажите имя снимка, например «Опора_1», без пути и расширения. В режиме File книга сохраняет и читает «Опора_1.ndm» в своей папке. В режиме Excel эта настройка не используется.",
            "В File кнопка «Расчёт» берёт геометрию из указанного файла и заменяет его полным новым расчётом. «Сформировать сечение» заменяет этот файл новой геометрией без НДС. Для нового имени сначала сформируйте сечение или скопируйте существующий снимок под этим именем; это позволяет сохранить прежний расчёт.",
            "Чтобы посмотреть сохранённый расчёт, выберите File, укажите его имя и нажмите «Обновить схему и результаты». Геометрия, диаграммы, схема и подробные результаты восстановятся из файла. Текущие настройки и входные нагрузки останутся прежними.",
            "Перед использованием File сохраните книгу в локальную папку. Не включайте в имя символы пути и запрещённые Windows символы. Если файл отсутствует или повреждён, книга сообщит причину и сохранит уже записанные результаты.",
            "Поддерживаются также снимки .ndms, если рядом нет одноимённого .ndm. При наличии обоих файлов используется .ndm; новый расчёт всегда сохраняется с расширением .ndm."
        ) }
        "General.WorstCombinationCriterion" { return @(
            "Настройка выбирает, какое сочетание считать Worst в верхней сводке Results.",
            "StrengthStrain выбирает сочетание с минимальным запасом по предельным деформациям модели прочности.",
            "StrengthCapacity выбирает сочетание с минимальным запасом по найденному множителю несущей способности. Это значение по умолчанию новой книги.",
            "Cracks выбирает сочетание с минимальным запасом по нормальным или продольным трещинам среди включенных проверок.",
            "Stability выбирает сочетание с минимальным запасом по расчету продольного изгиба и устойчивости. Для ветви СП 35 с условием N/N_cr ≤ Stability.SP35.NOverNcrLimit запас считается именно относительно этого лимита.",
            "Plot.LoadCase = Worst и AutoCAD.Export.CombinationID = Worst не пересчитывают критерий самостоятельно: они читают уже выбранное сочетание из верхней строки rngBatchSummary.",
            "Явно пустое, ошибочное или неизвестное значение этого критерия является ошибкой Config: сочетания не передаются в решатель и получают InputErr с указанием настройки. Другая проверка вместо выбранной молча не подставляется."
        ) }
        "Calculation.ZeroMomentPerDepth" { return @(
            "Введите неотрицательное число: 0 отключает инженерное зануление, положительное число задает порог. Отрицательное число, пустая ячейка, текстовая заглушка, ошибка формулы и потерянная строка дают InputErr с текущим адресом или разделом для восстановления. Значение по умолчанию вместо ошибочного ввода не подставляется.",
            "Инженерный порог обнуления малых моментов после переноса усилий, поворота осей и учета устойчивости. Исключает влияние численных остатков на выбор расчетной ветви.",
            "Порог для каждой плоскости считается так: M_tol = Calculation.ZeroMomentPerDepth · h.",
            "h - бетонный габарит сечения в рассматриваемой плоскости. Для Mx берется размер бетона по Y, для My - размер по X. В расчете устойчивости после перехода в главные оси используются h1 и h2.",
            "Значение вводится в текущих INPUT-единицах момент/длина. При стандартных единицах tf*m и mm значение 0.0005 означает 0.5 tf*m/m.",
            "Если |M| ≤ M_tol, программа зануляет момент до выбора расчетной ветки DirectState, Capacity, Crack или Stability.",
            "Порог не является допуском равновесия. Завышенное значение исключает из расчета реальные малые моменты и эксцентриситеты; при их проверке порог следует уменьшить или отключить."
        ) }
        "Geometry.Source" { return @(
            "Источник обязателен: выберите Generated или AutoCAD. Отсутствующая строка, пустая ячейка, текстовая заглушка и неизвестное значение являются ошибкой ввода; программа не выбирает Generated вместо пользователя.",
            "Generated - параметрическое построение бетонной сетки, арматуры и аннотаций для Circle, PointedSection, HollowRectangle и Rectangles.",
            "При выборе AutoCAD удобнее задать наружный контур и отверстия, оставив слой бетонных элементов пустым: программа построит сетку с настройками Mesh. Дополнительно можно импортировать готовые бетонные Region на AutoCAD.Common.ConcreteLayer. В обоих случаях арматура берётся из Region на AutoCAD.Common.RebarLayer. Нажмите «Сформировать сечение»: геометрия сохранится в Results, а при Plot.Enabled = Yes обновится схема.",
            "«Выполнить расчет» читает геометрию из Results в режиме Excel или из файла Results.SnapshotName в File. Нагрузки и расчётные настройки берутся из Config. После изменения размеров, армирования, сетки или чертежа заново сформируйте сечение. Выбор Geometry.Source и Geometry.Type сам по себе не меняет сохранённую геометрию.",
            "Геометрия AutoCAD считается заданной в мм: координаты и размеры - мм, площади - мм2, инерции - мм4. Чертеж для импорта нужно готовить в миллиметрах; INPUT не задает коэффициент масштаба геометрии. После импорта программа работает с этим сечением так же, как с геометрией, построенной внутри книги.",
            "У импортированной из AutoCAD сетки обычно нет сведений о смысловых гранях и группах арматуры. Поэтому программа сохраняет фактические элементы и результаты, но не пытается угадывать специальные подписи граней по координатам."
        ) }
        "Geometry.Type" { return @(
            "Доступны Circle, PointedSection, HollowRectangle и Rectangles. На листе «Сечения» каждому типу соответствует своя таблица с геометрией и армированием. При смене выбранного типа данные остальных сечений сохраняются. Оранжевым выделены доступные поля ввода, включая необязательные.",
            "При Geometry.Source = Generated выберите тип сечения и заполните все используемые размеры его формы. Пустая ячейка, текст вместо размера, ошибка формулы, нечисловое значение или потерянная строка не заменяются размерами шаблона. Сообщение указывает настройку и ее фактическую ячейку; для отсутствующей строки указывается раздел, где ее нужно восстановить.",
            "Не используется при Geometry.Source = AutoCAD: форму задают импортированные элементы либо точные наружные контуры и отверстия.",
            "Если для выбранной формы нет специальных аннотаций, расчет все равно выполняется. На схеме будут показаны геометрия и расчетные значения без этих дополнительных подписей."
        ) }
        "Mesh.CellGeometry" { return @(
            "Выберите способ построения бетонных элементов. Adaptive создаёт прямоугольную сетку с уточнением у границы. Contour обрезает базовые ячейки по контурам бетона и отверстий. По умолчанию выбран Contour.",
            "Mesh.StepMode определяет основной шаг: Auto выбирает его по габаритам сечения, User использует независимые Mesh.StepX и Mesh.StepY. Ячейки целиком внутри бетона остаются прямоугольными, целиком вне бетона — отбрасываются.",
            "Adaptive обычно формирует сетку в 3–4 раза быстрее Contour; выигрыш зависит от формы и шага. Это вариант для быстрого построения, когда допустимо приближение границы. Даже при Mesh.BoundarySubdivisions > 1 край сечения остаётся ступенчатым, из прямоугольников.",
            "В Contour программа обрезает базовую ячейку по прямым и круговым дугам с учётом всех отверстий, без дополнительного дробления. Если после обрезания получились отдельные бетонные части, каждая становится элементом. Mesh.BoundarySubdivisions не используется и не перезаписывается.",
            "В режиме Contour форма элементов автоматически создаётся по заданным контурам сечения и отверстий или читается из AutoCAD. Описывать каждый элемент вручную не нужно. Если форму определить невозможно, программа сообщает об ошибке без замены элемента прямоугольником.",
            "Имена C1, C2 и далее программа присваивает автоматически. Общие свойства элементов находятся в rngNDMSectionGeometry, их границы — в rngNDMElementGeometry, с привязкой по ElementID. В Excel таблица границ находится перед аннотациями Results. В File на листе остаётся шапка, а полные границы хранятся в связанном файле .ndm.",
            "Точность НДС проверяйте сравнением расчётов с разным основным шагом сетки. Площадь и инерции берутся по фактической форме элементов; напряжения рассчитываются по принятой волоконной модели. На схеме уменьшается только изображение элементов, их расчётная площадь сохраняется.",
            "После изменения способа или шага заново нажмите «Сформировать сечение». Сохранённая геометрия Results и ранее записанные файлы сами не перестраиваются."
        ) }
        "Mesh.StepMode" { return @(
            "Выберите Auto 0.05, Auto 0.1, Auto 0.2 или User. По умолчанию установлен Auto 0.1. Настройка применяется при формировании сечения по параметрам и при создании сетки по контурам AutoCAD; готовые бетонные элементы AutoCAD не разбиваются заново.",
            "В Auto программа берёт меньший из двух полных габаритов бетонного сечения по осям X и Y и умножает его на выбранный коэффициент. Шаг одинаков по обеим осям: h = min(B, H) · k. Например, для габаритов 300 × 500 мм режим Auto 0.1 даёт шаг 30 × 30 мм.",
            "Габариты определяются по точным наружным контурам с учётом положения и поворота сечения, а не по отверстию или отдельной стенке. Ручные Mesh.StepX и Mesh.StepY в Auto не используются и не перезаписываются.",
            "В режиме User программа использует именно два ручных шага Mesh.StepX и Mesh.StepY в единицах INPUT Length. Задайте оба положительными числами; они могут различаться.",
            "При Mesh.CellGeometry = Adaptive параметр Mesh.BoundarySubdivisions работает в Auto и User одинаково. По умолчанию 2: минимальный размер равен четверти основного шага. При Contour граничные ячейки обрезаются по фактическому контуру без дробления.",
            "Меньший коэффициент создаёт более мелкую сетку и обычно увеличивает время расчёта. Это способ выбора дискретизации, а не нормативный коэффициент. После изменения режима или размеров заново нажмите «Сформировать сечение»."
        ) }
        "Mesh.StepX" { return @(
            "При Mesh.StepMode = User задайте положительный шаг бетонной сетки по X в единицах INPUT Length. Например, шагу 50 мм соответствуют 50 при выборе mm, 5 при cm или 0.05 при m.",
            "В режиме Auto ручное поле сохраняется, но не используется. При выборе User оно читается вместе с независимо заданным Mesh.StepY. Если шаги равны, полные ячейки квадратные; если различаются — прямоугольные.",
            "Меньший шаг точнее описывает форму сечения и распределение напряжений, но увеличивает число элементов и время расчёта. Это параметр дискретизации, а не коэффициент СП.",
            "Пустое или неверное значение в режиме User нужно исправить: программа не копирует шаг из другой оси и не подставляет значение шаблона."
        ) }
        "Mesh.BoundarySubdivisions" { return @(
            "Параметр используется только при Mesh.CellGeometry = Adaptive. В Contour он сохраняется без чтения и не влияет на точное обрезание базовых ячеек.",
            "Укажите целое число уровней ≥ 1; по умолчанию — 2. Действует в Auto и User при построении по параметрам или контурам AutoCAD. Пустая ячейка, текст, ошибка формулы и дробь недопустимы.",
            "Значение N — число последовательных уровней. Каждый уровень делит только пересечённую границей ячейку пополам по X и Y, на четыре дочерние. Затем проверяются уже дочерние ячейки.",
            "Минимальные размеры равны h_x/2^N и h_y/2^N: при N = 1 это h/2, при 2 — h/4, при 4 — h/16. Например, при основном шаге 50 × 100 мм и N = 2 самые мелкие элементы имеют размеры 12,5 × 25 мм.",
            "Проверяются наружные границы и все отверстия, включая маленькое отверстие целиком внутри ячейки и отдельный остров бетона. Касание или совпадение со стороной без захода во внутренность не требует дробления.",
            "На последнем уровне ячейка включается по принадлежности центра бетону. Площадь сетки — сумма Width × Height элементов разных размеров. Точного обрезания граничных ячеек нет.",
            "Внутренние ячейки остаются крупными, наружные отбрасываются. Последняя базовая ячейка проверяется и при центре за габаритом: её дочерние части могут попадать в бетон.",
            "Большее N может увеличить время построения и расчёта. Сравнивайте площадь, центр тяжести, инерции и результаты при разных шагах. Если программа сообщает о переполнении из-за числа ячеек, увеличьте основной шаг или уменьшите N; это техническое ограничение, не требование СП.",
            "Mesh.BoundarySubdivisions — численная аппроксимация геометрии. Прямого требования СП к этому параметру нет."
        ) }
        "Circle.Diameter" { return @(
            "Наружный диаметр бетонного сечения. Сетка строится с шагами Mesh.StepX и Mesh.StepY.",
            "Для активного круглого сечения диаметр обязателен и должен быть положительным. Пустая ячейка, текстовая заглушка или потерянная строка не означают диаметр по умолчанию; введите число в текущих INPUT-единицах длины.",
            "Центр круга является внутренней технической точкой генератора. Пользовательская точка приложения нагрузок задается отдельно через Load.ReferenceOffsetX/Y относительно центра тяжести бетонного сечения.",
            "Для схемы создается размер диаметра D. Повторять размер круга с нескольких сторон не требуется."
        ) }
        "Rebar.AxisDistance" { return @(
            "Это расстояние от наружной грани круга до оси стержней первого ряда.",
            "Радиус линии арматуры: r_s = 0.5·D - a_s, где D - наружный диаметр бетона, a_s - расстояние от наружной грани до оси стержня.",
            "Для активного первого ряда значение обязательно: Rebar.Diameter/2 < a_s < Circle.Diameter/2. Пустота или ошибка формулы не заменяются отступом шаблона и требуют исправления указанной ячейки.",
            "Если первого ряда нет из-за пустого/нулевого количества или диаметра, отступ не читается и не меняет раскладку."
        ) }
        "Rebar.Count" { return @(
            "Это количество стержней первого ряда, равномерно расставленных по окружности.",
            "Пустой или нулевой Rebar.Count либо пустой/нулевой Rebar.Diameter выключает первый и дополнительные ряды. Для существующего первого ряда требуется целое количество не меньше 2. Отрицательное, дробное или непредставимое количество не округляется и не заменяется нулем.",
            "Угловой шаг между стержнями: Δφ = 2π·n⁻¹, где n - заданное количество стержней первого ряда."
        ) }
        "Rebar.Diameter" { return @(
            "Это диаметр стержней первого ряда круглого сечения.",
            "Площадь одного стержня: A_s = 0.25·π·d².",
            "Если диаметр равен 0 или пустой, первый ряд считается отсутствующим и дополнительные ряды не строятся."
        ) }
        "Rebar.Diameter2" { return @(
            "Это опциональный диаметр стержней второго ряда круглого сечения.",
            "Пустое или нулевое значение означает, что второй ряд не создается.",
            "Второй ряд всегда привязывается к уже созданным стержням первого ряда. Если первого ряда нет, второй ряд игнорируется."
        ) }
        "Rebar.Diameter3" { return @(
            "Это опциональный диаметр стержней третьего ряда круглого сечения.",
            "Пустое или нулевое значение означает, что третий ряд не создается.",
            "Третий ряд строится относительно первого ряда. Если второй ряд есть и расположен в том же направлении, третий ряд перескакивает второй, чтобы стержни не совпали."
        ) }
        "Rebar.Loc2row" { return @(
            "Определяет, куда ставить стержень второго ряда относительно каждого стержня первого ряда.",
            "Stacked: стержень ставится внутрь сечения по радиусу к центру.",
            "SideBySide: стержень ставится справа по часовой касательной.",
            "Выбор обязателен, когда второй ряд существует. Пустое или неизвестное положение не подставляет Stacked; сообщение указывает ячейку для исправления.",
            "Если второй ряд не задан по диаметру или нет первого ряда, эта настройка не используется."
        ) }
        "Rebar.Loc3row" { return @(
            "Определяет, куда ставить стержень третьего ряда относительно каждого стержня первого ряда.",
            "Stacked: внутрь по радиусу к центру. SideBySide: справа по часовой касательной.",
            "Выбор обязателен только для существующего третьего ряда. Пустое или неизвестное положение активного ряда дает InputErr; у отключенного ряда оно не читается.",
            "Если второй ряд уже занимает выбранное направление, третий ряд ставится дальше, за вторым рядом."
        ) }
        "Rebar.Bind2row" { return @(
            "Выбирает стержни первого ряда, рядом с которыми создаётся второй ряд: EachBar — все стержни; EverySecondBar — первый, третий и следующие нечётные.",
            "При нечётном числе основных стержней выбирается (n+1)/2 позиций. При отключённом втором или первом ряде настройка не используется."
        ) }
        "Rebar.Bind3row" { return @(
            "Выбирает стержни первого ряда для третьего ряда независимо от второго: EachBar — все; EverySecondBar — первый, третий и следующие нечётные.",
            "При отключённом третьем или первом ряде настройка не используется. Положение Stacked/SideBySide выбирается отдельно."
        ) }
        "PointedSection.*" { return @(
            "Эта настройка задается на листе «Сечения» в таблице «Клиновидное сечение». Геометрия выбирается через Geometry.Source = Generated и Geometry.Type = PointedSection.",
            "Размеры и названия сторон пояснены в общем разделе этой таблицы. W = 0 задаёт сторону без выступа, W > 0 — боковой выступ. Радиус R2 при W = 0 не используется.",
            "Активный размер обязателен. Число 0 допустимо только там, где оно означает отсутствие скругления или выноса; ошибочный ввод программа попросит исправить."
        ) }
        "Load.ReferenceOffset*" { return @(
            "Оба смещения обязательны: допускаются конечные положительные, отрицательные и нулевые числа в INPUT-единице длины. Пустота, текстовая заглушка, ошибка формулы и утраченная строка не означают нулевой эксцентриситет; сообщение указывает действующее поле Config и что нужно исправить.",
            "Если Load.ReferenceOffsetX = 0 и Load.ReferenceOffsetY = 0, заданные N, Mx и My относятся к центру тяжести бетонного сечения без учета продольной арматуры. Техническое начало координат встроенной формы на физический результат не влияет.",
            "Смещения учитываются при переносе усилий к расчетной точке. Изменение Load.ReferenceOffsetX/Y изменяет моменты от эксцентриситета N.",
            "Число вводится в текущей INPUT-единице длины. Например эксцентриситет 100 mm задается как 100 при mm, 10 при cm или 0.1 при m.",
            "Равновесие внутренних усилий НДМ соответствует общей постановке СП 63.13330.2018, п. 8.1.23. Принятая в программе система координат и знаков описана в начале этой справки и на листе Config в блоках Единицы измерения и Система знаков.",
            "Практическая проверка: если задана только сила N, а оба смещения равны нулю, нагрузка проходит через центр тяжести бетона. При несимметричной арматуре небольшая кривизна может возникнуть физически: центр приведенного сечения и центр бетона не обязаны совпадать."
        ) }
        "ContourLineWeight" { return @(
            "ContourLineWeight — общая толщина наружных границ и границ отверстий на Excel-схеме, в пунктах (pt). Значение по умолчанию — 2 pt.",
            "Одинаково применяется к прямым, дугам и окружностям у созданных программой и импортированных сечений. Не зависит от единиц длины и масштаба диаграммы. Введите положительное число; при Plot.ContourEnabled = No значение не используется."
        ) }
        "Offset" { return @($lead) + @(
            "Offset остается геометрическим отступом в миллиметрах сечения и не зависит от TextUnits.",
            "Для включенного типа аннотаций значение обязательно и не меньше нуля. Пустая ячейка, текст или ошибка формулы не заменяются запасным отступом.",
            "У размеров полученный отступ задаёт расстояние от измеряемой границы до размерной линии. Для сечений до 1000 мм включительно базовый Offset не увеличивается.",
            "У арматуры полученный отступ отсчитывается от бетонной границы. При Outside на прямой стороне расстояние от оси ряда до линии подписи равно as + Offset × max(1; D/1000). TextGap задаёт отдельный зазор от линии до текста и этим коэффициентом не умножается.",
            "После перевода в экранные пункты Excel отступ ограничивается для читаемости: 4–180 pt у размеров и 2–120 pt у арматуры. Это ограничение отображения, а не изменение геометрии сечения.",
            "Для постоянного экранного размера текста и зазора применяется TextUnits = pt; Offset остается геометрическим отступом."
        ) }
        "TextUnits" { return @($lead) + @(
            "Для включенного типа аннотаций обязателен один из вариантов mm или pt. Недопустимое или пустое значение дает сообщение с текущей ячейкой Config; программа не подставляет mm вместо ошибочного ввода.",
            "mm - TextHeight и TextGap задаются в реальных миллиметрах сечения. Это удобно, когда оформление должно оставаться пропорциональным самой модели.",
            "pt - TextHeight и TextGap задаются в обычных пунктах Excel. Это удобно для сравнения сечений разных габаритов в одной и той же области построения: текст остается примерно одинакового визуального размера.",
            "Настройка задается отдельно для двух типов аннотаций: в колонке Обозначения арматуры для Plot.RebarLabels.*, в колонке Размерные линии для Plot.Dimensions.*.",
            "Столбец Ед. у строк TextHeight и TextGap меняется формулой по выбранному TextUnits. Если для арматуры и размеров выбраны разные единицы, ячейка показывает обе.",
            "Для стабильной читаемости при разных габаритах сечения удобно выбирать pt. Режим mm нужен для пропорциональной привязки текста к масштабу модели."
        ) }
        "TextHeight" { return @($lead) + @(
            "TextHeight задает размер шрифта текста аннотации.",
            "Для включенного типа аннотаций задайте положительное число. Итоговый экранный шрифт ограничивается 5–28 pt для читаемости; при Enabled=No параметры отключенной группы не используются.",
            "Если TextUnits = mm, значение переводится из миллиметров сечения в пункты Excel с учетом масштаба рабочей области схемы. Поэтому при одном и том же окне схемы крупное сечение даст визуально меньший текст, а маленькое — больший.",
            "Если TextUnits = pt, значение сразу задает размер шрифта Excel в пунктах. В этом режиме текст почти не зависит от реальных габаритов сечения, поэтому его удобнее читать при сравнении разных сечений.",
            "Настройка не меняет расчет и сохраненные результаты. Она применяется при каждом обновлении схемы по текущему Config."
        ) }
        "TextGap" { return @($lead) + @(
            "TextGap задает зазор от линии аннотации до текстовой подписи.",
            "Для размеров это расстояние от размерной линии до текста. Для арматуры - от короткой линии обозначения или точки расположения подписи до текста.",
            "Если TextUnits=mm, зазор масштабируется как реальная величина сечения. Если TextUnits=pt, зазор задается как экранное расстояние Excel и выглядит стабильнее при смене размера сечения.",
            "TextGap не заменяет Offset. Offset двигает саму линию аннотации относительно сечения, а TextGap двигает только текст относительно уже поставленной линии."
        ) }
        "Concrete.R.SLS(II)" { return @($lead) + @(
            "Сопротивления бетона II группы: R_b,ser при сжатии и R_bt,ser при растяжении.",
            "Эти значения нужны для образования трещины и текущего НДС с трещинами. Модели выбираются на Config в профиле настройками с началом MaterialModel.CrackInitiation и MaterialModel.CrackedState.",
            "Введите уже итоговые значения с применимыми коэффициентами условий работы. Единица задается на Config в Units.Stress.Input; программа сама выполнит пересчет."
        ) }
        "Concrete.R.ULS(I)" { return @($lead) + @(
            "Сопротивления бетона I группы: R_b при сжатии и R_bt при растяжении.",
            "Эти сопротивления используются для прочности и устойчивости и не заменяются сопротивлениями II группы. В расчетном профиле на листе Config настройка MaterialModel.Strength.ConcreteTension = Ignore выключает растянутую ветвь бетона, а UseDiagram учитывает ее.",
            "Введите итоговые сопротивления с применимыми коэффициентами условий работы. Единица - Config Units.Stress.Input; повторно коэффициенты к сопротивлениям программа не применяет."
        ) }
        "Concrete.Rb.mc2" { return @($lead) + @(
            "Допустимое сжимающее напряжение R_b,mc2 для проверки продольных трещин по II группе.",
            "При Calculation.Crack.Width = Yes проверяется максимальное по модулю сжимающее напряжение бетона текущего НДС. При отсутствии сжатого бетона проверка имеет N/A.",
            "Запас SF_long = R_b,mc2 / σ_c,max; при SF_long < 1 статус FAIL.",
            "СП 35: область проверки - п. 7.2 и таблица 7.1; условие - п. 7.100; сопротивления - таблица 7.6; условия работы - таблица 7.7. Введите значение с уже учтенным m_b13 - коэффициентом условий работы для продольных трещин - и другими нужными поправками."
        ) }
        "Concrete.E" { return @($lead) + @(
            "Начальные модули бетона E_b и E_bt задаются отдельно для сжатия и растяжения в Units.Stress.Input.",
            "СП 63, п. 6.1.15 принимает одинаковые модули. Отдельная колонка растяжения позволяет задать другую растянутую ветвь, если это обосновано расчетной постановкой.",
            "Модули используются при построении диаграмм и определении приведенных характеристик сечения."
        ) }
        "Concrete.TwoLine.Eb1Red" { return @(
            "Задайте деформации перехода к плато двухлинейной диаграммы бетона: ε_b1,red для сжатия и ε_bt1,red для растяжения.",
            "Оба числа вводятся положительными по модулю. Физические знаки назначает программа. Деформации безразмерны и не пересчитываются при смене единиц.",
            "Значения используются, когда для соответствующей проверки в профиле выбрана бетонная диаграмма TwoLine.",
            "Диаграммы бетона описаны в СП 63, пп. 6.1.14, 6.1.20–6.1.24 и 6.1.26. Применимые характеристики и условия работы выбирает пользователь."
        ) }
        "Steel.R.ULS(I)" { return @($lead) + @(
            "Сопротивления арматуры I группы R_sc и R_s задаются в колонках сжатия и растяжения в Units.Stress.Input.",
            "Эти значения используются в прочности и устойчивости. TwoLine/ThreeLine - двух-/трехлинейная диаграмма; ее тип выбирается в расчетном профиле на Config.",
            "Производные точки деформаций строятся по СП 63, пп. 6.2.11, 6.2.14 и 6.2.15."
        ) }
        "Steel.R.SLS(II)" { return @($lead) + @(
            "Сопротивления арматуры II группы R_s,ser и R_sc,ser задаются в колонках растяжения и сжатия в Units.Stress.Input.",
            "СП 63 использует R_s,ser. Отдельное R_sc,ser предусмотрено программой для задания сжатой ветви арматуры. В шаблоне принято R_sc,ser = R_s,ser.",
            "Эти характеристики используются при образовании трещины и для текущего НДС с трещинами, а не вместо сопротивлений I группы для прочности."
        ) }
        "Steel.E" { return @($lead) + @(
            "Модули арматуры E_sc и E_s задаются отдельно для сжатия и растяжения в Units.Stress.Input.",
            "СП 63, п. 6.2.12 принимает одинаковые модули при растяжении и сжатии. Раздельные колонки позволяют задать ветви независимо при обоснованной постановке.",
            "E_s используется в диаграмме растяжения и формуле ширины нормальных трещин."
        ) }
        "Steel.TwoLine.Es2" { return @(
            "Задайте предельные деформации двухлинейной диаграммы арматуры: ε_sc2 для сжатия и ε_s2 для растяжения. Значения используются при выбранной диаграмме TwoLine.",
            "Оба числа вводятся положительными по модулю и не пересчитываются при смене единиц. Физический знак назначает программа.",
            "В одном сечении используется одна модель обычной ненапрягаемой арматуры. Для стержней разных классов или профилей требуется отдельное инженерное решение: программа не проверяет их нормативную совместимость автоматически."
        ) }
        "Solver.Method" { return @(
            "Newton уточняет равновесие по текущей касательной жесткости сечения. Обычно это основной рабочий вариант.",
            "Secant использует приближенную зависимость изменения усилий от деформаций и при застое начинает приближение заново. Он полезен для контрольной проверки, но может работать дольше.",
            "Искомые параметры плоскости деформаций: ε_0, κ_x и κ_y.",
            "Выбор влияет на прямое НДС и пробные состояния поиска LoadMultiplier. UltimateStrain решает отдельную предельную задачу. Пустой или неизвестный метод дает InputErr, а не скрытое переключение."
        ) }
        "General.PlateauSlopeFactor" { return @(
            "В разделе [Общие] на Config задайте неотрицательный числовой коэффициент α. Наклон ненулевых плато физических диаграмм бетона и арматуры равен Et = α·E. При α = 0 плато горизонтальное; при α > 0 напряжение на плато изменяется с деформацией.",
            "По умолчанию α = 0,001: наклон плато равен 0,001·E, то есть 0,1% от модуля соответствующей ветви. E — Eb/Ebt для сжатого/растянутого бетона и Esc/Es для сжатой/растянутой арматуры. Например, при E = 200000 МПа наклон равен 200 МПа. Это численная настройка программы для улучшения сходимости, а не нормативный коэффициент.",
            "Напряжение в начале плато сохраняется; далее его изменение равно α·E·Δε, где Δε — приращение деформации от начала плато. Напряжения и касательная жесткость определяются по одной диаграмме. Увеличение коэффициента сильнее меняет напряжения, несущую способность, образование и раскрытие трещин; улучшение сходимости не гарантируется для любого сочетания.",
            "Предельные деформации не изменяются. Выключенный растянутый бетон остается выключенным. Настройка независима от General.DiagramExtension, которое продолжает диаграммы за физическими пределами.",
            "Контрольные точки и графики материалов на Config учитывают текущее значение коэффициента. Results сохраняет фактически использованные диаграммы последнего расчета, включая наклон и разрешенное продолжение. После изменения коэффициента выполните расчет заново. Пустая ячейка, отрицательное значение и текст вместо числа недопустимы."
        ) }
        "Solver.MaxIterations" { return @(
            "Настройка задает наибольшее число итераций решателя на одной ступени нагрузки.",
            "Если равновесие не найдено за это число итераций, проверка получает статус NumFail. Подробная причина указывается в диагностическом отчёте.",
            "Значение новой книги - 80. Длина шага подбирается автоматически по уменьшению невязки; увеличение ограничения числа итераций не меняет допуски равновесия.",
            "Увеличение может помочь тяжелой нелинейной задаче, но увеличивает время и не исправляет ошибочные исходные данные."
        ) }
        "Solver.LoadSteps" { return @(
            "Настройка задает число ступеней приложения нагрузки при прямом расчете НДС.",
            "Равновесие определяется последовательно от частичной до полной нагрузки N + Mx + My.",
            "Больше ступеней обычно устойчивее для нелинейных диаграмм, но каждая ступень требует своих итераций."
        ) }
        "General.DiagramExtension" { return @(
            "Yes линейно продолжает активные ветви за физическими пределами с наклоном 0,01*E без конечной границы деформаций. No использует обычные диаграммы. Выбор обязателен.",
            "Физические сопротивления и предельные деформации от этого не растут. Равновесие только в продолженной части не является допустимым состоянием; оно может помочь подтвердить BaseFail - отказ постоянной части выбранного пути - вместо численной неудачи NumFail.",
            "Выключенный растянутый бетон остается выключенным. Для найденного состояния ExtUsed = Yes показывает фактическое использование продолженной части. Промежуточная итерация вне предела с возвратом в физический диапазон сама по себе не делает результат отказавшим.",
            "Настройка действует при расчёте прочности, несущей способности и трещин; формулы устойчивости не меняются. Физическое равновесие при Yes/No должно совпадать в пределах точности. В Results выводится участок диаграммы, охватывающий деформации всех найденных НДС с этой моделью. Границы вывода не ограничивают расчёт."
        ) }
        "Solver.ToleranceN" { return @(
            "Задайте положительный допуск по силе N в INPUT Force; единица выбирается на Config в Units.Force.Input.",
            "Условие сходимости: |N_int - N| ≤ Solver.ToleranceN, где N_int - внутренняя сила сечения.",
            "Около 1 Н соответствует 0,001 кН или 0,0001 tf. Меньший допуск строже, но может увеличить число итераций или дать NumFail - численную неудачу."
        ) }
        "Solver.ToleranceMx" { return @(
            "Задайте положительный допуск по моменту Mx в INPUT Moment; единица выбирается на Config в Units.Moment.Input.",
            "Условие сходимости: |Mx_int - Mx| ≤ Solver.ToleranceMx, где Mx_int - внутренний момент сечения по X.",
            "Около 1000 Н·мм соответствует 0,001 кН·м или 0,0001 tf·m. Меньший допуск строже, но может увеличить число итераций или дать NumFail."
        ) }
        "Solver.ToleranceMy" { return @(
            "Задайте положительный допуск по моменту My в INPUT Moment; единица выбирается на Config в Units.Moment.Input.",
            "Условие сходимости: |My_int - My| ≤ Solver.ToleranceMy, где My_int - внутренний момент сечения по Y.",
            "Смысл и пересчет такие же, как у Config Solver.ToleranceMx. Меньший допуск строже, но может увеличить число итераций или дать NumFail."
        ) }





        "Solver.SecantMaxRestarts" { return @(
            "Настройка ограничивает число перестроений приближенной матрицы чувствительности при методе Secant.",
            "Перестроение выполняется при застое, когда текущая матрица не обеспечивает уменьшение невязки.",
            "Настройка используется только при Solver.Method = Secant."
        ) }

        "Capacity.SolutionStrategy" { return @(
            "Выберите способ поиска предельной точки для LoadPath из таблицы сочетаний: Auto, UltimateStrain или LoadMultiplier. λ - множитель выбранных нагрузок; исходное сочетание соответствует λ = 1.",
            "UltimateStrain одновременно ищет равновесие и достижение предельной деформации. LoadMultiplier подбирает λ пробными расчетами равновесия. Auto сначала пробует UltimateStrain, а при численной неудаче повторяет тот же путь через LoadMultiplier.",
            "Для λ*N, а также λ*NMxy без пользовательских моментов сразу применяется LoadMultiplier: чистая силовая задача может быть вырожденной на горизонтальном плато диаграммы. Это не замена LoadPath.",
            "CapacitySolutionMethod в Results показывает фактически использованный метод. Подробное построение путей описано в разделах сочетаний и несущей способности.",
            "Оба метода используют одинаковые материалы и физические пределы. Несходимость - NumFail, а не доказательство разрушения. Пустой или неизвестный выбор требует исправления Config."
        ) }
        "Capacity.SearchMethod" { return @(
            "Настройка выбирает численный поиск λ - множителя нагрузок - только для фактически используемого LoadMultiplier.",
            "Bisection делит надежный интервал поиска пополам; Brent сохраняет интервал и пробует более быстрые шаги по промежуточным точкам; Secant прогнозирует следующую точку по двум предыдущим и чувствительнее к неровной зависимости.",
            "В успешном прямом UltimateStrain эта настройка не участвует, но допустимый выбор Bisection, Brent или Secant обязателен. Ошибка не заменяется другим методом.",
            "Это способы численного поиска программы, а не отдельные нормативные коэффициенты. Условия перехода к LoadMultiplier описаны у Config Capacity.SolutionStrategy."
        ) }
        "Capacity.InitialLambda" { return @(
            "Это начальная величина λ для метода LoadMultiplier на выбранной траектории LoadPath.",
            "Введите положительное число. Пустая ячейка, текст или ошибка формулы остановят запуск; сообщение укажет, какое поле Config нужно исправить.",
            "Если при этом λ сечение еще не достигло предельного состояния, программа расширяет верхнюю границу поиска.",
            "Если начальное значение слишком маленькое, может потребоваться больше расчетных точек. Если слишком большое, первая попытка чаще будет несходящейся или сразу за пределом.",
            "Обычно удобно начать с 1: это соответствует исходному уровню выбранной части нагрузки. Для очень слабых сечений можно поставить 0.5, для заведомо малых нагрузок — 2. Множитель λ безразмерен, поэтому при переходе с tf*m на kN*m значение не пересчитывается."
        ) }
        "Capacity.ToleranceLambda" { return @(
            "Это точность одномерного поиска предельного множителя λ в LoadMultiplier.",
            "Допустимо число больше нуля. Явно пустая или неверная ячейка, ноль и отрицательное значение дают InputErr с актуальным адресом Config до запуска НДС. Допустимость значения проверяется и тогда, когда выбран UltimateStrain, хотя само значение не участвует в его успешном прямом поиске.",
            "Bisection, Brent и Secant останавливаются, когда интервал/шаг по λ становится достаточно малым по этому допуску.",
            "Это допуск именно по коэффициенту λ, а не по деформациям. Для UltimateStrain используется Capacity.ToleranceStrain.",
            "Практические ориентиры: 0.1 - грубая быстрая оценка; 0.01 - обычный инженерный расчет; 0.001 - более точный, но медленный поиск. Значение безразмерно и не зависит от выбранных единиц N/Mx/My."
        ) }
        "Capacity.MaxRetries" { return @(
            "Это число дополнительных попыток для одной и той же точки λ, если внутренний расчет НДС не сошелся.",
            "Допустимо целое число не меньше нуля. Ноль корректен и отключает дополнительные повторы; дробь, отрицательное значение, пустая ячейка, текстовая заглушка и ошибка формулы дают InputErr с актуальным местом исправления до запуска НДС.",
            "Первая попытка имеет номер attempt = 0. Если она дала NumFail или SingularTangent, следующая попытка решает ту же самую точку λ, но с более осторожными численными настройками.",
            "Связь с Capacity.BaseLoadSteps прямая: LoadSteps_eff = Capacity.BaseLoadSteps · 2^{attempt}. При BaseLoadSteps = 1 и MaxRetries = 0 выполняется только первая попытка с 1 ступенью; при MaxRetries = 1 добавляется единственный повтор с 2 ступенями.",
            "Число ступеней после всех разрешённых удвоений не должно превышать 2147483646. Например, при BaseLoadSteps = 1 допускается не более 30 повторов. Это техническая граница программы; если она превышена, уменьшите число ступеней или повторов в указанной ячейке.",
            "Повтор использует больше ступеней нагрузки. Длина каждого шага подбирается автоматически без изменения заданной физической нагрузки.",
            "Если расчет при каждом пробном λ сходится с первой попытки, MaxRetries не добавляет повторных решений НДС. Повторы запускаются только после численной несходимости конкретной пробы. На некоторых осевых путях такие пробы встречаются часто, поэтому разрешенные повторы заметно увеличивают время.",
            "Если повторы исчерпаны, точка λ получает NumFail. Такая несходимость не считается разрушением сечения.",
            "Значение по умолчанию - 0. Один повтор ограничивает дополнительные затраты одной попыткой с уменьшенным шагом; 2...3 повтора увеличивают затраты при несходимости. Параметр не зависит от единиц измерения."
        ) }
        "Capacity.BaseLoadSteps" { return @(
            "Исходное число ступеней приложения нагрузки при проверке каждого пробного множителя λ в LoadMultiplier.",
            "Допустимо целое число больше нуля. Ноль, отрицательное значение, дробь, явно пустая ячейка или ошибка формулы дают InputErr до запуска НДС; комментарий называет текущую ячейку Config и действие пользователя.",
            "BaseLoadSteps задаёт число ступеней первой попытки при каждом пробном λ. Если MaxRetries разрешает повтор, число ступеней увеличивается по формуле BaseLoadSteps·2^attempt.",
            "Для всех разрешённых повторов должно выполняться BaseLoadSteps·2^{MaxRetries} ≤ 2147483646. При превышении программа попросит уменьшить число ступеней или повторов. Не используйте очень большие значения без необходимости: они могут сильно замедлить расчёт.",
            "Например при BaseLoadSteps = 1 и MaxRetries = 0 каждая λ-точка решается один раз с одной ступенью. Если поставить MaxRetries = 1, несошедшаяся точка будет повторена с двумя ступенями и более осторожным шагом кривизны.",
            "Большее базовое значение делает каждую первую λ-точку устойчивее, но замедляет весь LoadMultiplier даже там, где сходимость была бы нормальной.",
            "Настройка не используется как физический шаг увеличения нагрузки. Внешний поиск λ выполняют Bisection, Brent или Secant.",
            "Значение по умолчанию - 1. Дополнительные ступени только после численной неудачи задаются через Capacity.MaxRetries.",
            "Значения 2...4 повышают дробность приложения нагрузки и затраты каждой λ-точки. Дальнейшее увеличение оправдано только при подтвержденном влиянии числа ступеней на сходимость."
        ) }
        "Capacity.ToleranceStrain" { return @(
            "Задайте положительный безразмерный допуск для попадания в предельную деформацию при UltimateStrain. Он не является самим физическим пределом.",
            "Физический предел берется из диаграммы выбранной модели материала. ε_ult означает его модуль; ult - предельный. Допуск отношения текущей деформации к пределу равен max(10⁻⁶; Capacity.ToleranceStrain / max(|ε_ult|; 10⁻⁹)).",
            "Финальная точка также должна удовлетворять равновесию и физическим пределам. Численный допуск не разрешает превышение диаграммы и не обязан менять уже более точно найденный результат.",
            "Для LoadMultiplier точность множителя задает Config Capacity.ToleranceLambda. Допустимое значение Capacity.ToleranceStrain все равно обязательно.",
            "Меньший допуск может повысить точность, но увеличить число итераций. Он безразмерен и не меняется вместе с единицами."
        ) }
        "Capacity.MaxLambda" { return @(
            "Это защитная верхняя граница расширения интервала в методе LoadMultiplier.",
            "Допустимо число больше нуля. Ноль, отрицательное значение, явно пустая ячейка или ошибка формулы дают InputErr с актуальным адресом Config до запуска НДС. Если InitialLambda больше MaxLambda, первая положительная проба ограничивается MaxLambda, а не выходит за него.",
            "В LoadMultiplier программа расширяет интервал поиска только до этого значения. Если до MaxLambda предел не найден, успешный результат не выдаётся; причина указывается в комментарии.",
            "В UltimateStrain эта настройка не ограничивает найденную λ: там программа напрямую ищет предельную плоскость деформаций и принимает результат, если λ положительная и численно конечная.",
            "Ориентиры для LoadMultiplier: 10...20 удобно для быстрой оценки; 1024 - значение по умолчанию с большим запасом для малых исходных нагрузок и контрольных расчетов. λ безразмерна и не зависит от выбранных единиц."
        ) }
        "Capacity.SolverMaxIterations" { return @(
            "Для LoadMultiplier эта настройка задает, сколько итераций можно сделать на каждой проверяемой точке λ, прежде чем считать эту точку численно несошедшейся.",
            "Допустимо целое число больше нуля. Ноль, отрицательное значение, дробь, явно пустая ячейка или ошибка формулы дают InputErr с ключом и текущим адресом Config до запуска НДС.",
            "Для UltimateStrain эта же настройка ограничивает число итераций прямого поиска предельной плоскости деформаций.",
            "На одной итерации Newton программа может несколько раз оценивать усилия: для проверки равновесия, уточнения чувствительности к деформациям и выбора шага. Поэтому число записей в диагностике может быть больше числа итераций.",
            "Недостаточный лимит может дать NumFail для физически допустимой нагрузки. Завышенный лимит увеличивает время неуспешного поиска.",
            "Лимит следует увеличивать при остановке по числу итераций после проверки исходных данных и диаграмм. Увеличение не исправляет ошибки единиц, материалов или геометрии.",
            "Capacity.SolverMaxIterations - это защитный численный параметр программы. Он не является расчетным коэффициентом СП."
        ) }
        "Stability.Code" { return @(
            "Выберите SP63 или SP35. Перед расчётом программа проверяет обязательные настройки устойчивости. Если значение отсутствует или неверно, сообщение укажет поле Config и способ исправления.",
            "SP63 включает учет продольного изгиба по формулам СП 63.13330.2018, п. 8.1.15: условная жесткость D, критическая сила N_cr и коэффициент η.",
            "SP35 включает ветвь СП 35.13330.2011: для малых эксцентриситетов используется отдельная проверка устойчивости, для больших - увеличение момента через η.",
            "В обоих вариантах эксцентриситет продольной силы считается относительно центра тяжести приведенного сечения бетон + арматура. Пользователь при этом вводит нагрузки относительно бетонного центра и Load.ReferenceOffset; программа переносит моменты к приведенному центру перед расчетом устойчивости.",
            "Расчет запускается только для тех сочетаний, где выбранный профиль имеет Calculation.Stability.Enabled = Yes."
        ) }
        "Stability.ElementLength" { return @(
            "Это геометрическая длина элемента L до умножения на μ1/μ2.",
            "Введите положительное число в текущей INPUT-единице длины. Ноль и отрицательная длина являются ошибкой ввода, а не основанием взять 8000 по умолчанию.",
            "Для критической силы и гибкости программа считает расчетные длины l_0,1 = μ1·L и l_0,2 = μ2·L.",
            "Если длина L или коэффициент μ настолько велики или малы, что расчётную длину нельзя вычислить надёжно, программа укажет ошибку ввода. Проверьте единицы и значения в указанных ячейках. Это ограничение числового расчёта, а не дополнительный нормативный предел.",
            "По СП 63 п. 8.1.7 и СП 35 п. 7.52 случайный эксцентриситет в режиме AutoWithL берется именно по L, а не по l_0.",
            "Программа не анализирует закрепления элемента автоматически: пользователь задает длину и коэффициенты расчетной длины уже по своей расчетной схеме."
        ) }
        "Stability.Mu1" { return @(
            "Задайте коэффициент расчётной длины μ1 для изгиба вокруг главной оси 1 приведённого сечения. Расчётная длина равна μ1·L, где L задаётся в Stability.ElementLength.",
            "Введите положительное число без единиц. В новой книге установлено 2, но нужное значение выбирайте по своей расчётной схеме: программа не определяет закрепления автоматически.",
            "Для круга и других почти изотропных сечений главная ось 1 совпадает с X: плоскость соответствует изгибу Mx, а эксцентриситет и габарит определяются по Y.",
            "Если μ1·L невозможно надёжно вычислить из-за слишком большого или малого числа, программа укажет участвующие ячейки. Исправьте коэффициент, длину или её единицы."
        ) }
        "Stability.SystemType" { return @(
            "Допустимы только Determinate и Indeterminate, без учета регистра. Неизвестный вариант не трактуется как определимая схема и дает InputErr с указанием места исправления.",
            "Determinate выбирают для статически определимой расчетной схемы: программа складывает эксцентриситет от заданного момента и случайный эксцентриситет.",
            "Indeterminate выбирают для статически неопределимой расчетной схемы: программа берет расчетный эксцентриситет не меньше случайного, то есть большее значение по модулю.",
            "В обозначениях СП 63 это формирование e_0 с учетом e_a; в обозначениях СП 35 - формирование e_c с учетом e_c,sl.",
            "Настройка влияет только на правило объединения основного и случайного эксцентриситета. Остальные формулы СП 63 или СП 35 выбираются настройкой Stability.Code."
        ) }
        "Stability.ZeroMomentEccentricitySign1" { return @(
            "Укажите направление случайного эксцентриситета в первой главной плоскости, когда момент в ней равен нулю. Значение 1 задаёт положительное направление, −1 — противоположное.",
            "Для круга и других почти изотропных сечений это направление по Y при изгибе Mx. В общем случае используются главные оси приведённого сечения.",
            "Здесь задаётся только знак. Величину эксцентриситета определяет Stability.AccidentalEccentricityMode. Ноль, пустая ячейка и другое число требуют исправления."
        ) }
        "Stability.AccidentalEccentricityMode" { return @(
            "Допустимы AutoWithL, AutoWithμL (также запись AutoWithMuL) и User. Неизвестный вариант не заменяется автоматическим режимом. Ввод проверяется и при растягивающем сочетании, для которого сама устойчивость неприменима.",
            "AutoWithL - нормативный режим по геометрической длине: СП 63 использует e_a = max(L/600; h/30; 10 мм), СП 35 использует e_c,sl = L/400.",
            "AutoWithμL - режим по умолчанию в книге: вместо геометрической длины L в случайном эксцентриситете используется расчетная длина l_0 = μL.",
            "User означает, что пользователь сам задает модуль случайного эксцентриситета отдельно для двух главных плоскостей: Stability.AccidentalEccentricityUser1/2.",
            "Знак эксцентриситета программа выбирает отдельно: по знаку момента в главной плоскости, а при чистом сжатии - по настройкам Stability.ZeroMomentEccentricitySign1/2."
        ) }
        "Stability.AccidentalEccentricityPlanes" { return @(
            "Допустимы OnlyMomentPlane и BothPlanes. Неизвестный вариант, пустая ячейка или ошибка формулы дают InputErr с актуальным адресом Config.",
            "OnlyMomentPlane сохраняет принятую упрощенную схему: если момент есть только в одной главной плоскости, случайный эксцентриситет добавляется только в этой плоскости.",
            "BothPlanes добавляет случайный эксцентриситет в обеих главных плоскостях. Если момент в плоскости ненулевой, знак берется по направлению увеличения этого момента; если момент равен нулю, знак берется из Stability.ZeroMomentEccentricitySign1/2.",
            "Настройка не запускает перебор знаков и не добавляет дополнительных НДМ-расчетов."
        ) }
        "Stability.AccidentalEccentricityUser1" { return @(
            "В режиме Stability.AccidentalEccentricityMode = User задайте положительную величину случайного эксцентриситета первой главной плоскости. Используются единицы INPUT Length.",
            "Для круга и других почти изотропных сечений это плечо по Y при изгибе Mx. Направление определяется отдельно; отрицательное число здесь не заменяется своим модулем.",
            "В автоматических режимах поле не участвует в формуле, но должно содержать корректное неотрицательное число; допускается 0."
        ) }
        "Stability.PhiLMode" { return @(
            "Допустимы Auto и PhiL2. Неизвестный режим не заменяется значением по умолчанию, даже если текущая табличная ветвь не использует эту настройку.",
            "Auto считает коэффициент длительности по нормативному отношению длительного момента к полному моменту.",
            "Для СП 63 моменты берутся относительно центра наиболее растянутого или наименее сжатого стержня арматуры.",
            "Для СП 35 в ветви e_c > r моменты берутся относительно соответствующей грани сечения; при противоположных знаках полного и длительного моментов применяется специальное правило СП 35.",
            "PhiL2 принимает φ_l = 2 как консервативный пользовательский режим только в расчетных ветвях N_cr/η.",
            "В режиме Auto программа ограничивает коэффициент диапазоном 1...2. Нижняя граница 1 принята как консервативное проектное правило: длительная часть не должна искусственно уменьшать расчетный эффект.",
            "В табличной ветви СП 35 при e_c ≤ r эта настройка не заменяет табличный φ_l: там всегда используются φ_m и φ_l из таблицы 7.21.",
            "Для СП 63 таблица нагрузок означает постоянные + длительные нагрузки. Для СП 35 таблица означает постоянные нагрузки."
        ) }
        "Stability.SP63.Ks" { return @(
            "Коэффициент k_s учитывает вклад арматуры в условную жесткость D = k_b·E_b·I_b + k_s·E_s·I_s по СП 63.",
            "Задайте положительное число. Некорректное значение дает InputErr с адресом Config, в том числе при неактивной ветви СП 63."
        ) }
        "Stability.SP63.DeltaEMin" { return @(
            "Нижняя граница δ_e = e_0/h в формуле k_b по СП 63. Принятая нормативная граница - 0.15.",
            "Задайте неотрицательное число; ноль отключает эту программную границу. Если верхняя граница DeltaEMax также включена, нижняя не должна превышать верхнюю. Некорректный диапазон дает InputErr с адресом Config, в том числе при неактивной ветви СП 63."
        ) }
        "Stability.SP63.DeltaEMax" { return @(
            "Верхняя граница δ_e = e_0/h в формуле k_b по СП 63. Принятая нормативная граница - 1.5.",
            "Задайте неотрицательное число; ноль отключает эту программную границу. Если нижняя граница DeltaEMin также включена, верхняя не должна быть меньше нижней. Некорректный диапазон дает InputErr с адресом Config, в том числе при неактивной ветви СП 63."
        ) }
        "Stability.SP35.PhiP" { return @(
            "Пользовательский коэффициент φ_p для устойчивости по СП 35. Для ненапрягаемой арматуры в текущей постановке обычно используется 1.",
            "Задайте положительное число. Некорректный ввод дает InputErr до расчетов, в том числе при неактивной ветви СП 35; допустимость числа не означает нормативного подтверждения произвольного коэффициента."
        ) }
        "Stability.SP35.NOverNcrLimit" { return @(
            "Предел отношения N/N_cr для ветви η СП 35. В Results условие выводится как коэффициент запаса NOverNcrLimit·N_cr/N: значение больше 1 означает, что оно выполняется.",
            "Задайте положительное число. Некорректный ввод дает InputErr с адресом Config, в том числе при неактивной ветви СП 35."
        ) }
        "SLS.Crack.Code" { return @(
            "Выберите SP35 или SP63 для ширины нормальных трещин. Общий расчет включается настройкой профиля Calculation.Crack.Width на Config.",
            "Методика устойчивости выбирается независимо настройкой Config Stability.Code. Продольные трещины проверяются по текущему сжатому бетону при любой методике ширины."
        ) }
        "SLS.Crack.SP35.GroupGapTolerance" { return @(
            "Задайте наибольшее расстояние между поверхностями растянутых стержней, при котором они объединяются в группу. Число неотрицательное, в INPUT Length; в новой книге 10 мм.",
            "0 объединяет касающиеся стержни. Положительный допуск разрешает небольшой зазор и является допущением программы, не требованием СП. Группа неделима: в площадь взаимодействия она входит целиком по своему центру."
        ) }
        "SLS.Crack.SP35.RowTolerance" { return @(
            "Задайте допустимый разнос центров групп, относящихся к одному ряду арматуры, по нормали к нейтральной линии (НЛ). Используются единицы INPUT Length; значение по умолчанию 50 мм.",
            "Первый ряд арматуры находится у растянутой поверхности бетона; следующие ряды нумеруются в направлении нейтральной линии. Этот допуск только распределяет группы по рядам: боковые границы площади взаимодействия и выбор соседей он не меняет."
        ) }
        "SLS.Crack.SP35.NeighborRatioLimit" { return @(
            "Порог сравнивается отдельно для ближайшей растянутой группы арматурных стержней слева и справа от рассматриваемой группы вдоль нейтральной линии (НЛ). L - полное расстояние между центрами групп, t - расстояние между их проекциями на НЛ.",
            "Допустимый диапазон 0...1; значение по умолчанию 0,2. При t/L ниже порога боковой отступ равен r; при равенстве или превышении - min(r; t/2)."
        ) }
        "SLS.Crack.SP35.RadiusDiameterMode" { return @(
            "Выберите диаметр d для расстояния r, которым ограничивается площадь взаимодействия: Max - наибольший, Min - наименьший, Average - среднее арифметическое диаметров отдельных стержней.",
            "Для боковых границ учитываются стержни рассматриваемой группы, для границы в сторону нейтральной линии - стержни опорного ряда. В формуле радиуса армирования R_r всегда складываются реальные диаметры, независимо от этого выбора. R_r - отношение площади бетона к вкладу арматуры."
        ) }
        "SLS.Crack.SP35.InteractionRadiusMode" { return @(
            "Выберите 3d, 5d или 6d: этот множитель задает отступ r для боковых границ площади взаимодействия и границы в сторону нейтральной линии. В новой книге установлен вариант 6d.",
            "СП 35, п. 7.110: 6d - общий случай; 3d - непрямоугольное сечение с равномерной арматурой по контуру; 5d - пучки и канаты. Последние не поддерживаются этой версией; выбор режима требует проверки нормативной применимости."
        ) }
        "SLS.Crack.SP35.RebarProfile" { return @(
            "Выберите Periodic для периодического профиля или Smooth для гладкой арматуры. Этот выбор задает зависимость коэффициента ψ от радиуса армирования R_r в формуле ширины.",
            "ψ - размерный коэффициент раскрытия, R_r - отношение площади взаимодействия к вкладу арматуры. По СП 35 п. 7.109 арматура в стенках также относится к первой зависимости: для этого случая выбирайте Periodic."
        ) }
        "SLS.Crack.Allowable" { return @(
            "Это предельно допустимая ширина раскрытия нормальной трещины a_crc,ult.",
            "Программа не назначает это значение автоматически. Пользователь выбирает его по нормам, категории трещиностойкости, условиям эксплуатации и проектным требованиям.",
            "После расчета программа выводит коэффициент запаса по трещинам: SF_crc = a_crc,ult·a_crc⁻¹.",
            "Если SF_crc < 1, рассчитанная ширина раскрытия превышает заданный пользователем предел. Это не численная ошибка, а результат проверки.",
            "Число вводится в текущей INPUT-единице длины. Например допустимые 0.3 mm задаются как 0.3 при Units.Length.Input = mm, 0.03 при cm или 0.0003 при m.",
            "Условие проверки раскрытия трещин приведено в СП 63.13330.2018, п. 8.2.6, формула (8.118). Конкретное допустимое значение зависит от расчетной ситуации и должно быть подтверждено пользователем."
        ) }
        "SLS.Crack.InitiationSolutionStrategy" { return @(
            "Выберите способ поиска точки образования нормальной трещины. Физический путь задается отдельно столбцом LoadPath в таблице сочетаний на Config.",
            "UltimateStrain напрямую ищет равновесие на первом деформационном пределе. LoadMultiplier подбирает множитель нагрузки λ пробными расчетами и уточняет найденную предельную плоскость. λ = 1 соответствует исходному сочетанию.",
            "Проверяются образование трещины, сжатие бетона, растяжение и сжатие арматуры. Если другой предел достигнут раньше трещины, он не выдается за точку трещинообразования. NotCracked требует подтвержденного допустимого НДС при исходной нагрузке.",
            "Auto сначала пробует UltimateStrain, при численной неудаче повторяет тот же путь через LoadMultiplier. На пути λ*N такой резерв используется и при явно выбранном UltimateStrain: постоянные моменты и эксцентриситет не меняются. Отрицательный λ сам по себе не означает отсутствие трещины. Неудача поиска точки образования не блокирует ширину при пригодном текущем НДС."
        ) }
        "SLS.Crack.ToleranceLambda" { return @(
            "Введите положительное число без единиц. Оно ограничивает разность верхней и нижней пробных λ при поиске точки образования трещины методом LoadMultiplier. λ = 1 соответствует исходному сочетанию.",
            "Больший допуск может уменьшить число проб, но дает менее точное начальное приближение предельной нагрузки. Найденная плоскость проверяется по равновесию и физическим пределам материалов. Если точка не подтверждена, программа дополнительно сужает тот же интервал, не обнуляя предельное число итераций.",
            "Настройка применяется также при переходе Auto к LoadMultiplier. Прямой поиск UltimateStrain и отдельный допуск несущей способности Capacity.ToleranceLambda она не изменяет."
        ) }
        "SLS.Crack.Phi1" { return @(
            "Коэффициент длительности нагрузки φ_1 в формуле ширины СП 63. Для продолжительного раскрытия - 1,4 по п. 8.2.6.",
            "Введите применимое положительное значение на Config. Программа рассчитывает продолжительное раскрытие; отдельная процедура непродолжительного раскрытия по нескольким состояниям не реализована. Для СП 35 эта настройка не используется."
        ) }
        "SLS.Crack.Phi2" { return @(
            "Коэффициент профиля арматуры φ_2 в формуле ширины СП 63.",
            "Периодическому профилю соответствует 0,5, гладкому - 0,8 по СП 63, п. 8.2.6. Введите применимое значение на Config; при SLS.Crack.Code = SP35 эта настройка не используется."
        ) }
        "SLS.Crack.Phi3Mode" { return @(
            "Режим определения коэффициента характера работы элемента φ_3 в формуле ширины СП 63.",
            "Auto принимает 1,2 при растягивающей продольной силе, 1,0 в остальных реализованных случаях. User использует число Config SLS.Crack.Phi3.",
            "Настройка не меняет НДС или бетонную площадь. Для SLS.Crack.Code = SP35 она не используется."
        ) }
        "SLS.Crack.Phi3" { return @(
            "Коэффициент φ_3 применяется при SLS.Crack.Phi3Mode = User и должен быть положительным.",
            "В режиме Auto число не подставляется в формулу, но для СП 63 ячейка все равно должна содержать корректный положительный коэффициент. Для SLS.Crack.Code = SP35 настройка не используется."
        ) }
        "SLS.Crack.PsiMode" { return @(
            "Режим определения ψ_s - коэффициента неравномерности деформаций арматуры между трещинами в формуле СП 63. При SLS.Crack.Code = SP35 не используется.",
            "User при пригодной точке образования берет число из Config SLS.Crack.PsiS. Auto сначала считает ширину с 1 и уточняет коэффициент только при превышении допуска Config SLS.Crack.Allowable. AlwaysCalc выполняет уточнение всегда, когда оно возможно.",
            "Уточнение сравнивает средние по площади напряжения одного и того же набора стержней при текущей нагрузке и сразу после образования трещины. Формула и обозначения приведены в разделе СП 63.",
            "При непригодной точке трещинообразования либо отсутствии растяжения выбранной арматуры в ней ширина по допустимому текущему НДС определяется с ψ_s = 1 во всех режимах. Предупреждение в ResultComment выводится только при фактическом использовании коэффициента.",
            "Подтвержденный NotCracked означает отсутствие нормальной трещины: ширина не считается, имеет N/A, пользовательское PsiS не меняется и предупреждения о назначении 1 нет. Продольные трещины при пригодном текущем НДС проверяются независимо."
        ) }
        "SLS.Crack.SigmaSCrcAveragingMode" { return @(
            "Способ учета сжатых стержней при определении среднего по площади напряжения σ_s,crc сразу после трещинообразования по СП 63.",
            "AllSelected учитывает напряжения со знаком. TensionOnly принимает сжимающие напряжения за 0, но сохраняет площади этих стержней в знаменателе среднего.",
            "Набор стержней выбирается по текущему НДС и не пересобирается для порогового состояния. Настройка влияет только на уточнение коэффициента Config SLS.Crack.PsiMode, а не на текущие напряжения, площадь или диаметр арматуры."
        ) }
        "SLS.Crack.PsiS" { return @(
            "Пользовательский коэффициент неравномерности деформаций ψ_s применяется при SLS.Crack.PsiMode = User и пригодных данных трещинообразования.",
            "Для СП 63 положительное число обязательно и в режимах Auto/AlwaysCalc, хотя тогда оно не задает итоговый коэффициент. Для SLS.Crack.Code = SP35 эта настройка не используется.",
            "Правила уточнения и принятия 1 при недоступных пороговых данных описаны у Config SLS.Crack.PsiMode. При подтвержденном NotCracked число не меняется: формула ширины не вызывается."
        ) }
        "SLS.Crack.TensionZoneMode" { return @(
            "Площадь бетона A_bt для СП 63: FullTension - вся растянутая часть; Effective - ограниченная полоса у растянутой поверхности.",
            "Для Effective: h_bt = min(max(x_t; 2a_s); 0,5h), где x_t - глубина растянутой зоны до НЛ; a_s - отступ крайнего выбранного стержня от бетонной границы; h - габарит по нормали к НЛ.",
            "При x_t < 2a_s полоса может зайти за нейтральную линию, но сжатые стержни не добавляются в расчет ширины. Смысл отступа задается отдельно на Config в SLS.Crack.CoverDistanceMode.",
            "Площадь определяется пересечением полосы с бетонной областью за вычетом учитываемых отверстий, в том числе с частичным пересечением элементов. Центр бетонного элемента не является правилом включения всей его площади. Выбор точного контура или сетки описан в конце раздела подготовки AutoCAD.",
            "Формула расстояния между трещинами и ограничения приведены в методике СП 63. Для SLS.Crack.Code = SP35 эта настройка не применяется."
        ) }
        "SLS.Crack.CoverDistanceMode" { return @(
            "Способ определения a_s от оси крайнего растянутого стержня до бетонной границы по нормали к НЛ. Применяется только в СП 63.",
            "NearestContour берет ближайший выход из бетона в сторону растянутой поверхности. GlobalExtreme берет дальнюю опорную линию всего сечения в том же направлении; для невыпуклой формы расстояние может быть больше.",
            "Это расстояние участвует в эффективной полосе Config SLS.Crack.TensionZoneMode = Effective. Оно не меняет НДС, прочность или устойчивость. При центральном растяжении без нейтральной линии настройка не участвует."
        ) }
        "SLS.Crack.*" { return @(
            "Общие формулы и порядок расчета приведены в разделе расчета ширины раскрытия нормальных трещин; здесь описывается только выбранная настройка."
        ) }
        "AutoCAD.Common.SectionContourLayer" { return @(
            "Единый слой точной наружной границы для импорта и экспорта. На нем допустимы замкнутая 2D-полилиния, окружность или Region - объект «Область». Region с вырезанными отверстиями передает также все эти отверстия; взрывать его не нужно.",
            "При импорте готовой бетонной сетки слой наружного контура может отсутствовать или быть пустым. Если наружный контур задан без отверстий, для области трещин внутри него принимается сплошной бетон. Чтобы учесть отверстия, задайте их тоже.",
            "Подготовка чертежа, пространство «Модель», система координат, замыкание полилиний и правила учёта отверстий описаны в первом разделе справки. Если готовой бетонной сетки нет, программа строит её по наружному контуру и отверстиям; готовую сетку эти контуры не изменяют."
        ) }
        "AutoCAD.Common.OpeningContourLayer" { return @(
            "Единый слой отдельно заданных отверстий для импорта и экспорта. Можно разместить несколько замкнутых 2D-полилиний, окружностей или Region - объектов «Область». Отверстия, уже вырезанные в наружном Region, здесь не дублируйте.",
            "Отверстия можно задать и без наружного контура: наружная граница берется по сетке, а вычитаются только эти точные отверстия. Пустой или отсутствующий слой не блокирует импорт.",
            "Отверстия не должны касаться или пересекать друг друга и известную наружную границу. Остальные правила и четыре случая учета границ приведены в первом разделе подготовки AutoCAD."
        ) }
        "AutoCAD.Common.ConcreteLayer" { return @(
            "Единый слой расчетных бетонных элементов для импорта и экспорта. Region — объект «Область» — на этом слое задаёт бетонные элементы; в Contour несвязные части становятся отдельными элементами. Полилинии, блоки и текст не становятся расчетными элементами.",
            "Если на слое нет ни одного объекта, бетонная сетка автоматически генерируется по точному наружному контуру и всем заданным отверстиям с настройками Mesh. Если нет и наружного контура, импорт отклоняется. Непригодные объекты или Region, отсеянные общим порогом площади, не считаются пустым слоем и не запускают генерацию.",
            "Имя обязательно и должно отличаться от Config AutoCAD.Common.RebarLayer. Точные границы задаются отдельно на Config через AutoCAD.Common.SectionContourLayer и AutoCAD.Common.OpeningContourLayer.",
            "Координаты и размеры Region читаются как мм, площадь как мм2, инерции как мм4 независимо от INPUT/OUTPUT и единиц чертежа. Config Mesh.MinElementArea - только пользовательский порог в INPUT Area; он не масштабирует чертеж.",
            "Для НДС сохраняются фактические площадь, центр и инерции Region. В Mesh.CellGeometry = Contour сохраняется его точная форма из прямых и круговых дуг, включая отверстия и отдельные части. В Adaptive схема использует прямоугольную оболочку: Rectangle — заданный прямоугольник, Equivalent rectangle — эквивалентный прямоугольник, Equivalent square — равноплощадный квадрат.",
            "Только в Adaptive направление эквивалентного квадрата берётся по пригодной прямой грани. У круга и других элементов без определимого направления - по среднему направлению остальных ориентированных бетонных элементов. Если такого направления нет, угол 0. Поэтому углы около +89° и -89° дают вертикаль, а не горизонталь.",
            "В Contour схема и экспорт используют фактическую геометрию элемента. В Adaptive для схемы и границ сетки используются сохранённые ширина, высота и поворот; они не заменяют реальные инерции. Цвет при экспорте определяется физическим сжатием или растяжением из сохранённого Results."
        ) }
        "AutoCAD.Common.RebarLayer" { return @(
            "Единый слой арматуры для импорта и экспорта. Каждый Region - объект «Область» - на этом слое становится одним стержнем; эквивалентный круглый диаметр определяется по его площади.",
            "Имя обязательно и должно отличаться от Config AutoCAD.Common.ConcreteLayer. Материал выбирается при расчете по профилю на Config, а не по имени слоя.",
            "Цвет экспортируемого стержня показывает его физическое сжатие или растяжение из Results. Разные материалы и профили арматуры в одном сечении эта версия не поддерживает."
        ) }
        "Mesh.MinElementArea" { return @(
            "Минимальная допустимая площадь бетонного элемента или арматурного стержня. Порог действует при любом способе формирования сечения.",
            "В Adaptive сравнивается площадь полученного прямоугольника, в Contour — фактическая площадь каждой связной части после обрезания. При импорте AutoCAD сравниваются конечные бетонные части и арматурные Region.",
            "Область с площадью строго меньше порога пропускается; равная порогу сохраняется. Ноль отключает фильтр. Слишком большой порог может удалить существенную часть бетона или все стержни: проверяйте состав и площадь сформированного сечения.",
            "Порог вводится в INPUT Area и один раз переводится в мм². Например 25 mm2, 0.25 cm2 и 0.000025 m2 задают один порог. Масштаб чертежа AutoCAD остаётся 1 единица = 1 мм.",
            "Изменение порога применяется при следующем нажатии «Сформировать сечение». Чтение готового Results и восстановление поддерживаемого снимка не удаляют сохранённые элементы.",
            "Задайте неотрицательное число. Пустая ячейка, ошибка формулы, отрицательное или слишком большое для перевода значение останавливают формирование с указанием фактической ячейки Mesh.MinElementArea."
        ) }
        "AutoCAD.Export.CombinationID" { return @(
            "Выбор обязателен: укажите Worst или ID сохраненного сочетания. Пустая ячейка, текстовая заглушка, ошибка Excel и отсутствующая строка не заменяются Worst; сообщение указывает текущую ячейку Config и способ исправления.",
            "Worst означает: взять определяющее сочетание из верхней сводки Results. Критерий выбора задается General.WorstCombinationCriterion.",
            "Геометрия экспортируется в AutoCAD всегда в миллиметрах: координаты, размеры, диаметры, контур, оси и точка приложения нагрузки. Выбор OUTPUT для длины в Config не меняет масштаб чертежа: одна единица AutoCAD соответствует 1 мм.",
            "Если указан конкретный номер, например LC2, в AutoCAD выгружаются напряжения/деформации именно для этого сочетания.",
            "Значение должно совпадать с CombinationID, который есть в результатах последнего расчета. Если указать несуществующее сочетание, экспорт остановится с понятной ошибкой.",
            "Если сечение сформировано, но расчёт ещё не выполнен, экспортируются только бетон, арматура и включённые геометрические контуры. Программа всегда предупреждает, что передано сечение без рассчитанного состояния. Напряжения, деформации и нейтральная линия не выводятся.",
            "Экспорт не запускает новый расчёт или импорт. Для сохранённого рассчитанного состояния используются его профиль, единицы и результаты; ошибка такого снимка не заменяется экспортом одной геометрии. После изменения нагрузок выполните расчёт, чтобы экспортировать актуальное состояние."
        ) }
        "AutoCAD.Export.LabelMode" { return @($lead) + @(
            "Выбор обязателен: пустота, текстовая заглушка, неизвестное значение или ошибка формулы не заменяются NamesAndValues, а дают сообщение с текущим адресом Config и допустимыми вариантами.",
            "ValuesOnly выводит только численное значение выбранной профильной величины Visualization.Quantity.",
            "NamesAndValues добавляет к значению имя расчетного элемента, например C1 или R1.",
            "При NamesAndValues имена выводятся и для бетона, и для арматуры, если они сохранены в Results."
        ) }
        "AutoCAD.Export.PrincipalAxesMode" { return @($lead) + @(
            "Выбор обязателен; пустое или неизвестное значение не включает оси автоматически. Сообщение ошибки указывает текущую ячейку Config и варианты исправления.",
            "Transformed выводит главные центральные оси приведенного сечения бетон + арматура.",
            "Concrete выводит главные центральные оси только бетонного сечения.",
            "None полностью отключает выгрузку главных осей. Точка приложения нагрузки и нейтральная линия управляются своими настройками.",
            "Экспорт берет центр и угол выбранных осей из сохраненных Results и не пересчитывает свойства сечения заново."
        ) }
        "AutoCAD.Export.ContourEnabled" { return @(
            "Yes выгружает сохраненную наружную границу сечения и отверстия. Формат задает AutoCAD.Export.ContourFormat. No отключает только исходные контуры сечения, не расчетную область трещин.",
            "Для Circle, Rectangles, PointedSection и HollowRectangle контур приходит из генератора вместе с прямыми участками и дугами. Дуги передаются как дуговые сегменты полилинии, а не как набор коротких линий.",
            "Исходный внешний контур и отверстия экспортируются из rngNDMSectionContours. Без отдельно импортированных точных контуров эта опция не подменяет их приблизительной границей сетки.",
            "Имена Common.SectionContourLayer и Common.OpeningContourLayer обязательны; отсутствующие слои создаются. Распределение контуров по слоям и цветам зависит от выбранного формата и описано в его справке. Это не слой рассчитанной области Ar.",
            "Каждый экспорт добавляет новые контуры без поиска совпадающих объектов и без автоматического удаления прежнего вывода. За чистоту места экспорта отвечает пользователь: заранее удалите ненужные объекты или выберите свободную область чертежа."
        ) }
        "AutoCAD.Export.ContourFormat" { return @(
            "Region (по умолчанию): известный наружный контур и все отверстия образуют Region бетона на Common.SectionContourLayer. Каждая отдельная бетонная часть получает свой Region. Все границы одного Region имеют единый цвет ACI 30 (оранжевый).",
            "Если известны только отверстия, Region выводит их отдельно на Common.OpeningContourLayer цветом ACI 4 (голубой), не создавая наружный контур по сетке.",
            "Polyline: наружная граница и каждое отверстие выводятся отдельными замкнутыми полилиниями. Наружные границы - на Common.SectionContourLayer, ACI 30; отверстия - на Common.OpeningContourLayer, ACI 4. Прямые и круговые дуги сохраняются точно.",
            "Оба формата можно импортировать обратно. Выбор влияет только на исходные контуры сечения и действует при ContourEnabled = Yes. Контур расчетной области трещин остается полилинией на собственном слое. Пустое или неизвестное значение активной настройки требует исправления в указанной ячейке Config."
        ) }
        "AutoCAD.Export.ExportCrackInteractionContour" { return @(
            "Yes выводит сохраненную критическую область СП 35 или фактическую область СП 63. No не меняет расчет и не выводит эту дополнительную геометрию.",
            "Область выводится только вместе с состоянием трещин, выбранным Visualization.State профиля: PreCrackState, PostCrackState или CrackedState. Для StrengthState/CapacityState она не выводится.",
            "Экспорт использует последний расчет на Results и сохраняет наружные контуры, отверстия и круговые дуги. Если область не рассчитана или устарела, она не выводится. Превышение допуска ширины не мешает посмотреть рассчитанную область.",
            "Все координаты экспорта всегда в мм. Слой задается AutoCAD.Export.CrackInteractionLayer, цвет объектов - ACI 31. Каждый запуск добавляет новый контур без автоматической очистки места экспорта."
        ) }
        "AutoCAD.Export.CrackInteractionLayer" { return @(
            "На этом отдельном слое сохраняется рассчитанная область взаимодействия для трещин. Он не предназначен для сетки, наружного контура или отверстий исходной модели. Не размещайте здесь исходные объекты: команда очистки удаляет оформление NDM.",
            "Настройка обязательна при включенном AutoCAD.Export.ExportCrackInteractionContour и при ручной очистке оформления. Объекты области взаимодействия получают цвет ACI 31, в том числе на существующем слое."
        ) }
        "AutoCAD.Export.NeutralLineEnabled" { return @(
            "Если выбрано Yes, вместе с выбранным НДС в AutoCAD выводится его нейтральная линия. Если выбрано No, линия не выводится.",
            "Экспорт использует сохранённые данные расчёта и не решает НДС заново. Укажите Yes или No; неверный выбор будет отмечен в сообщении с адресом Config."
        ) }
        "AutoCAD.Export.LoadPointEnabled" { return @(
            "Если выбрано Yes, в AutoCAD выводится точка приложения нагрузки выбранного сочетания. Если выбрано No, точка не выводится.",
            "Настройка не меняет положение нагрузки в расчёте. Экспорт читает сохранённые данные; укажите Yes или No в соответствующей ячейке Config."
        ) }
        "AutoCAD.Layer.ConcreteTension" { return @(
            "Укажите имя слоя подписей растянутых бетонных элементов. Если слоя нет, программа создаст его. Существующий слой используется с его текущими свойствами.",
            "Имя обязательно, не длиннее 255 символов. Запрещены управляющие символы, запятая, двойные кавычки и символы < > / \ : ; ? * | = '. При EXTNAMES=0 действуют дополнительные ограничения самого чертежа.",
            "Не совмещайте слой подписей с базовыми слоями бетона, арматуры и контуров. Команда очистки удаляет оформление, поэтому не храните на нём пользовательские линии и тексты. Объекты Region при очистке сохраняются."
        ) }
        "AutoCAD.Color.ConcreteTension" { return @(
            "Укажите цвет растянутых бетонных элементов и их подписей в AutoCAD. Используется физическое растяжение из сохранённого результата, независимо от выбранных знаков нагрузок.",
            "Введите целый индекс ACI от 1 до 255. Значения 0 (ByBlock) и 256 (ByLayer) здесь недопустимы; дробные значения не округляются.",
            "Выбор меняет оформление экспорта, а не напряжения или результат проверки. При неверном значении сообщение укажет ячейку Config."
        ) }
        "Plot.Enabled" { return @(
            "Yes обновляет схему на листе «Расчет». No сохраняет прежнюю схему без обновления.",
            "Схема строится только по последнему сохраненному Results и не обращается к текущей геометрии Config.",
            "Выберите Yes или No для обновления схемы, в том числе сразу после формирования сечения. При No существующая схема сохраняется, а настройки скрытых элементов не используются. Неверный выбор требует исправления указанной ячейки на листе Config."
        ) }
        "Plot.AutoUpdateAfterCalculation" { return @(
            "Yes автоматически перестраивает схему сразу после успешного расчета.",
            "No оставляет схему как есть. Для ручного обновления нажмите «Обновить схему и результаты».",
            "Выбор обязателен, если Plot.Enabled = Yes. Пустота, текстовая заглушка, ошибка Excel и отсутствие строки требуют исправления Config до начала расчета и очистки предыдущих Results.",
            "При Plot.Enabled = No настройка не используется, существующая схема сохраняется даже при отсутствии доступного НДС. Если оба переключателя Yes, но расчетных состояний нет, прежняя схема очищается с пояснением отсутствия данных.",
            "В обоих случаях обновление схемы только читает Results и не запускает расчет заново."
        ) }
        "Plot.LoadCase" { return @(
            "Worst показывает определяющее сочетание из верхней сводки Results. Критерий выбора задается General.WorstCombinationCriterion.",
            "Конкретный CombinationID, например LC2, показывает сохраненные напряжения, деформации и плоскость деформаций именно этого сочетания.",
            "Переключение LoadCase не создает геометрию заново: меняется только отображаемое расчетное состояние из Results.",
            "Укажите сочетание, чтобы показать его НДС либо схему геометрии с пояснением, почему это НДС отсутствует. Пустая ячейка и текстовая заглушка не означают Worst. Сразу после формирования сечения показывается предварительная схема без выбора сочетания; эта настройка тогда не используется."
        ) }
        "Plot.ResultGradient" { return @(
            "Yes включает цветовой градиент расчетных элементов по величине Visualization.Quantity выбранного профиля.",
            "No отключает градации интенсивности: работающие элементы получают один цвет своего материала и физического состояния, заданный в Plot.Color.*.",
            "Неработающий растянутый бетон остается отдельным серым состоянием и не смешивается с обычным градиентом."
        ) }
        "Plot.ResultLabelsEnabled" { return @(
            "Yes добавляет на схему численные подписи величины Visualization.Quantity выбранного профиля.",
            "Подписываются только бетонные элементы с пространственным шагом Plot.ResultLabelSpacing.",
            "Неработающий растянутый бетон не подписывается, чтобы схема не показывала нулевые значения как полезные напряжения.",
            "Для схемы найденного НДС выберите Yes или No. На предварительной схеме одной геометрии расчетных значений еще нет, поэтому этот переключатель и шаг подписей не используются."
        ) }
        "Plot.ResultLabelSpacing" { return @(
            "Это положительный пространственный шаг ячеек, по которым выбираются численные подписи результата. Он всегда задается в мм независимо от Units.Length.Output.",
            "Чем больше значение, тем меньше подписей и чище схема. Чем меньше значение, тем подробнее поле значений, но выше риск визуальной перегрузки.",
            "Шаг применяется пространственно, а не по номеру ElementID; в каждой ячейке выбирается ближайший к ее центру работающий бетонный элемент. Это не проверка точного минимального расстояния между любой парой подписей.",
            "При Plot.ResultLabelsEnabled=Yes значение обязательно и больше нуля. Пустота, текстовая заглушка, ноль и отрицательный шаг требуют исправления указанной ячейки; автоматический шаг вместо них не назначается. При No параметр не читается."
        ) }
        "Plot.LegendMode" { return @(
            "Separate строит две шкалы: отдельно для бетона и отдельно для арматуры.",
            "Common строит одну общую шкалу для всех элементов. В этом режиме цвета берутся из Plot.Color.RebarCompression и Plot.Color.RebarTension.",
            "Раздельные шкалы удобны, когда диапазоны напряжений бетона и арматуры сильно отличаются."
        ) }
        "Plot.PrincipalAxesMode" { return @(
            "Transformed показывает главные центральные оси приведенного сечения бетон + арматура.",
            "Concrete показывает главные центральные оси только бетонного сечения.",
            "None полностью скрывает главные оси на схеме. Точка приложения нагрузки и нейтральная линия управляются отдельными настройками.",
            "Схема берет центр и угол выбранных осей из сохраненных Results. Изменение этого режима после расчета только меняет отображение уже сохраненных осей.",
            "Выберите режим для расчетной схемы, в том числе если выбранное НДС отсутствует и показывается только геометрия. Пустая ячейка, текстовая заглушка или неизвестный вариант требуют исправления Config до очистки старой схемы. Предварительная схема сразу после формирования сечения не показывает эти вспомогательные оси и не использует их настройки."
        ) }
        "Plot.AxisLabelsEnabled" { return @(
            "Yes показывает маленькие курсивные подписи Ось X и Ось Y внутри рабочей области схемы.",
            "No скрывает только эти подписи. Сами оси, нейтральная линия, контур и точка приложения нагрузки управляются своими настройками.",
            "Подписи нужны как визуальная подсказка направления глобальных осей и не влияют на расчет или данные Results."
        ) }
        "Plot.AxisLabelsFontSize" { return @(
            "Размер задается в пунктах Excel (pt), как обычный размер шрифта в книге.",
            "По умолчанию используется 7.5 pt - это текущий небольшой размер подписей осей.",
            "При Plot.AxisLabelsEnabled=Yes значение обязательно и больше нуля; ошибочный ввод требует исправления и не заменяется размером шаблона. При No высота подписей не читается."
        ) }
        "Plot.ContourEnabled" { return @(
            "Yes показывает на Excel-схеме сохраненные точные наружные контуры и отверстия из Results. Они могут быть получены как из генератора книги, так и при импорте из AutoCAD.",
            "Этот контур нужен как читаемая внешняя граница сечения поверх расчетной сетки. Он не влияет на расчет и не заменяет сами бетонные элементы.",
            "Если точные контуры не сохранены, программа не придумывает их по сетке. No скрывает только этот дополнительный контур."
        ) }
        "Plot.Color.*" { return @($lead) + @(
            "Значение задается в формате R,G,B, например 30,80,220.",
            "Нужны ровно три целых числа от 0 до 255. Значение 0,0,0 является черным цветом. Пустой или ошибочный активный цвет требует исправления; дробные компоненты не округляются.",
            "Цвет используется для крайнего значения соответствующего материала и физического состояния в легенде и на элементах.",
            "При Plot.LegendMode = Common используются цвета арматуры как общая шкала для всех элементов."
        ) }
        "Plot.NeutralLineEnabled" { return @(
            "Если выбрано Yes, на схеме показывается нейтральная линия выбранного сочетания, определённая по сохранённому полю деформаций. Если выбрано No, линия скрыта.",
            "Линия доступна только для найденного НДС. При показе одной геометрии этот переключатель не используется.",
            "Укажите Yes или No. Настройка меняет только отображение; неверное значение потребует исправить указанную ячейку Config."
        ) }
        "Plot.LoadApplicationPointEnabled" { return @(
            "Если выбрано Yes, на схеме показывается точка приложения нагрузки сохранённого расчёта. Если выбрано No, точка скрыта.",
            "Переключатель не переносит нагрузку и не меняет эксцентриситет: он управляет только изображением уже заданной точки.",
            "Укажите Yes или No. При неверном значении программа укажет ячейку Config до изменения существующей схемы."
        ) }
        "Plot.LegendEnabled" { return @(
            "Если выбрано Yes, рядом со схемой выводится легенда выбранной величины и физических состояний материалов. Если выбрано No, легенда скрыта.",
            "Легенда доступна для найденного НДС. При показе одной геометрии без расчётных значений этот переключатель не используется.",
            "Укажите Yes или No. Настройка меняет только отображение; неверное значение потребует исправить указанную ячейку Config."
        ) }
        "Enabled" { return @($lead) + @(
            "Столбцы «Обозначения арматуры» и «Размерные линии» включаются независимо. Можно оставить размеры сечения и скрыть подписи арматуры либо сделать наоборот.",
            "Значение No скрывает только выбранный тип подписей на Excel-схеме. Геометрия, рассчитанные значения и экспорт в AutoCAD сохраняются; остальные параметры отключенного столбца не используются."
        ) }
        "Placement" { return @($lead) + @(
            "Выберите Outside или Inside в нужном столбце таблицы. Для подписи арматуры это направление от бетонной границы наружу или в бетон. Для размеров меняется сторона текста относительно размерной линии, а сама линия остается на месте.",
            "Расстояние до линии задается отдельно строкой Offset, а зазор между текстом и линией — строкой TextGap. Placement меняет только размещение подписи и не меняет координаты сечения."
        ) }
        "LineEnabled" { return @($lead) + @(
            "Эта строка относится только к столбцу «Обозначения арматуры». Размерные и выносные линии она не скрывает: их показ управляется столбцом «Размерные линии».",
            "При No остается текст с количеством и диаметрами стержней, а короткая линия рядом с ним не рисуется. Толщина этой линии LineWeight тогда не используется. Чтобы убрать и текст, выберите No в строке Enabled того же столбца."
        ) }
        "ArrowType" { return @($lead) + @(
            "Выберите форму стрелок: Triangle - треугольник, Stealth - треугольник с вырезом, Diamond - ромб, Oval - овал, Open - открытая стрелка.",
            "Это только оформление размерных линий Excel-схемы. На размеры сечения, расчет и экспорт в AutoCAD выбор не влияет."
        ) }
        "ArrowSize" { return @($lead) + @(
            "Выберите размер стрелок: Small - малый, Medium - средний, Wide - широкий. Это варианты Excel, а не длина в мм.",
            "Выбор меняет только стрелки размерных линий Excel-схемы."
        ) }
        "Color" { return @($lead) + @(
            "Задайте три целых числа от 0 до 255 через запятую: красный, зеленый, синий. Например 0,0,0 - черный, 255,0,0 - красный.",
            "Колонка арматуры задает цвет ее подписей и основных линий; колонка размеров - цвет размерного текста и основных линий. Это только Excel-схема."
        ) }
        "ExtensionLineColor" { return @($lead) + @(
            "Задайте цвет выносных линий размеров тремя целыми числами от 0 до 255: красный, зеленый, синий. Например 140,140,140 - серый.",
            "Настройка относится только к размерным линиям Excel-схемы. Цвет основных линий и текста задается отдельно строкой Color."
        ) }
        "LineWeight" { return @($lead) + @(
            "Введите положительную толщину основной линии в пунктах Excel (pt). Арматурные подписи и размеры имеют отдельные значения в своих колонках.",
            "Это только оформление Excel-схемы, без влияния на расчет и экспорт в AutoCAD."
        ) }
        "ExtensionLineWeight" { return @($lead) + @(
            "Введите положительную толщину выносных линий размеров в пунктах Excel (pt). Основная размерная линия настраивается отдельно строкой LineWeight.",
            "Настройка меняет только оформление размерных линий Excel-схемы."
        ) }
        default { return @(
            $lead,
            "Единица указана в столбце «Ед.» на Config и связана с текущими INPUT через формулу, если параметр размерный.",
            "Расчетные зависимости приведены в соответствующем методическом разделе справки."
        ) }
    }
}

# Делает первую строку инструкции одинаково понятной: пользователь сразу видит,
# какой именно ключ открыт и что конкретно он меняет. Остальные строки только
# раскрывают последствия и формулы для этой настройки.
function Get-InstructionLead {
    param([string]$Key, [string]$Comment)

    if ([string]::IsNullOrWhiteSpace($Comment)) {
        return "$Key — настройка на листе Config. Единица и допустимые варианты указаны рядом с ячейкой ввода."
    }
    if ($Key -in @('ArrowType','ArrowSize','ExtensionLineWeight','ExtensionLineColor')) {
        return "$Key — строка столбца «Размерные линии» на листе Config: Plot.Dimensions.$Key. $Comment"
    }
    if ($Key -eq 'LineEnabled') {
        return "$Key — строка столбца «Обозначения арматуры» на листе Config: Plot.RebarLabels.$Key. $Comment"
    }
    if ($Key -in @('Enabled','Placement','Offset','LineWeight','TextUnits','TextHeight','TextGap','Color')) {
        return "$Key — строка таблицы «Настройки оформления схемы» на листе Config. В столбце арматуры это Plot.RebarLabels.$Key, в столбце размеров — Plot.Dimensions.$Key. $Comment"
    }
    return $Comment
}

# Формирует список блоков справки для Config и специальных таблиц настроек.
function Get-SettingsInstructionCatalog {
    $items = New-Object System.Collections.Generic.List[object]
    $items.Add(@{ Key = "Units"; Title = "Единицы измерения"; Lines = @(
        "На листе Config в таблице «Единицы измерения» выберите единицы ввода в столбце INPUT и единицы результатов в столбце OUTPUT. Столбец INTERNAL показывает постоянные внутренние единицы программы; его не меняйте.",
        "Длина задается строкой Length, площадь - Area, сила - Force, момент - Moment, напряжение - Stress, кривизна - Curvature. Ключи этих настроек имеют вид Units.Length.Input и Units.Length.Output; остальные строятся так же по имени строки.",
        "Допустимые единицы: Length - mm, cm, m; Area - mm2, cm2, m2; Force - N, kN, tf; Moment - N*mm, kN*m, tf*m; Stress - Pa, kPa, MPa, kgf/cm2, tf/m2; Curvature - 1/mm, 1/m.",
        "Внутри используются мм, мм2, Н, Н*мм, МПа и 1/мм. Программа сама переводит ввод и вывод. INPUT и OUTPUT могут отличаться.",
        "При смене INPUT уже введенные числа не пересчитываются. Чтобы сохранить ту же задачу, переведите нагрузки, размеры, параметры материалов и размерные настройки Config, включая численные допуски, в новые единицы.",
        "Смена OUTPUT не меняет расчетное НДС. Новые результаты записываются в выбранных выходных единицах; координаты сформированного сечения сохраняются в единицах его таблицы и правильно читаются независимо от текущего OUTPUT.",
        "Геометрия AutoCAD имеет постоянный масштаб: 1 единица чертежа = 1 мм. Импорт принимает размеры в мм, площади в мм2, инерции в мм4; экспорт геометрии всегда в мм независимо от INPUT/OUTPUT. Настройка Config Mesh.MinElementArea - исключение: введенный порог площади читается в INPUT Area.",
        "После импорта можно сменить INPUT/OUTPUT и считать без повторного импорта: сохраненное сечение сохраняет физические размеры, новый Results записывается в текущих OUTPUT. Нагрузки, материалы и настройки Config читаются в текущих INPUT, поэтому их числа при смене единиц нужно пересчитать.",
        "Пустая или неизвестная единица - ошибка ввода. Исправьте указанную ячейку Config; значение шаблона вместо нее не подставляется."
    )}) | Out-Null
    $items.Add(@{ Key = "SignConvention"; Title = "Система знаков"; Lines = @(
        "На Config в таблице «Система знаков» выберите правила в колонке USER: Sign.N.User - для силы N, Sign.Mx.User и Sign.My.User - для моментов Mx и My. Один пользовательский выбор действует при вводе и выводе. INTERNAL показывает фиксированное внутреннее правило.",
        "Для N вариант Compression означает, что положительное пользовательское число задает сжатие; Tension - растяжение. Для Mx вариант +Y tension означает растяжение стороны +Y при положительном моменте, -Y tension - стороны -Y. Для My аналогично выбирается +X или -X.",
        "Напряжения Stress и деформации Strain всегда сохраняют физический знак: плюс - растяжение, минус - сжатие. Кривизны KappaX/KappaY описывают поле деформаций во внутренних осях; смена пользовательского знака момента их повторно не переворачивает.",
        "Цвета показывают физическое состояние бетона и арматуры, а не выбранный знак нагрузки. Положительные запасы, жесткости и пределы тоже не меняют знак.",
        "Пустое или неизвестное правило нужно исправить в указанной ячейке. Сохранённый Results относится к знакам последнего расчета; для результатов в новых знаках выполните расчет заново.",
        "Внутри сжатие отрицательное, растяжение положительное согласно СП 63 п. 8.1.22. Настройки пользовательских знаков эту математику не меняют."
    )}) | Out-Null
    $items.Add(@{ Key = "LoadCombinations"; Title = "Сочетания нагрузок"; Lines = @(
        "Сначала нажмите «Сформировать сечение»: программа построит его по параметрам листа «Сечения» или импортирует из AutoCAD согласно Geometry.Source. Затем задайте сочетания на Config и нажмите «Выполнить расчёт». Расчёт использует сохранённое сечение; после изменения геометрии его нужно сформировать заново.",
        "CombinationID - идентификатор сочетания; N, Mx и My вводятся в INPUT и пользовательской системе знаков. Пустые Mx/My принимаются равными 0.",
        "Пустой CombinationID отключает строку, даже если нагрузки заполнены: сохраните их и верните номер, когда сочетание снова понадобится. Номер с нулевыми N, Mx и My также пропускается без ошибки. При заполненном номере и хотя бы одной ненулевой нагрузке выполняется обычный расчёт. Если активных сочетаний нет, расчёт не запускается, сохранённые результаты остаются на листе.",
        "ProfileId выбирает профиль по имени Config Profile.DisplayName. Профиль задает запрошенные проверки и модели материалов, а не очередность действий программы. Comment - ваше пояснение; оно попадет в сводку и заголовок схемы.",
        "LoadPath - общий путь для несущей способности и образования трещины. λ - множитель, которым увеличиваются выбранные нагрузки. При λ = 1 это исходное сочетание; остальные составляющие выбранного пути сохраняются.",
        "Auto выбирает путь по нагрузкам: только N - λ*N; только Mx - λ*Mx; только My - λ*My; оба момента без N - λ*Mxy. Моменты по двум осям считаются одним изгибным действием.",
        "При N вместе с моментами Auto сначала пробует путь по моментам, затем λ*N, затем λ*NMxy, если предыдущая попытка не дала предельную точку. Найденный предел с λ < 1 завершает поиск: это результат для нагрузки выше предела, а не неудачная попытка.",
        "Пустой LoadPath равен Auto, в том числе пустой результат формулы. Неизвестный вариант и ошибка формулы требуют исправления. Явно выбранный путь не заменяется другим.",
        "Момент от смещенной продольной силы масштабируется вместе с N. Смещения задайте на Config в Load.ReferenceOffsetX и Load.ReferenceOffsetY. Например при λ*N пользовательские Mx/My постоянны, а моменты от смещенной силы меняются вместе с ней.",
        "Ненулевые малые нагрузки не исчезают из пути из-за допусков равновесия. Для специального обнуления малых моментов есть Config Calculation.ZeroMomentPerDepth; 0 выключает это правило.",
        "Численные способы поиска задаются отдельно на Config: Capacity.SolutionStrategy - для несущей способности, SLS.Crack.InitiationSolutionStrategy - для образования трещины. Их описание приведено в соответствующих методических разделах.",
        "В Results показываются принятый путь и причина отказа. BaseFail - физический отказ постоянной части пути; NumFail - численная неудача. NotCracked означает подтвержденное отсутствие нормальной трещины при текущей нагрузке. Неудачная проба большей нагрузки не делает текущее НДС отказавшим.",
        "В шаблоне 30 строк, но можно изменить их количество; оставьте хотя бы одну. Расширяйте именно именованный диапазон rngLoadCombinations через Формулы → Диспетчер имен или вставкой ячеек внутри таблицы. Не вставляйте целую строку листа, если рядом есть другие таблицы.",
        "На Results между блоками нужны две пустые строки. Если сочетания не помещаются, сообщение укажет, какие именованные блоки нужно сдвинуть. Сохраняйте имена диапазонов; их адрес можно менять через Диспетчер имен."
    ); Tables = @(
        @{
            Title = "Варианты LoadPath"
            Headers = @("Путь", "Масштабируемые составляющие", "Постоянные составляющие", "Условие применимости", "Недопустимый случай")
            Rows = @(
                @("Auto", "Зависит от ненулевых N, Mx/My.", "Зависит от выбранного пути.", "Хотя бы одна ненулевая нагрузка.", "Для полностью нулевого сочетания."),
                @("λ*Mx", "Пользовательский Mx.", "N, My и моменты от смещения N.", "Ненулевой Mx.", "При Mx = 0."),
                @("λ*My", "Пользовательский My.", "N, Mx и моменты от смещения N.", "Ненулевой My.", "При My = 0."),
                @("λ*Mxy", "Оба пользовательских момента.", "N и моменты от смещения N.", "Ненулевой Mx или My.", "При Mx = My = 0."),
                @("λ*N", "N и моменты от ее смещения.", "Пользовательские Mx/My.", "Ненулевая N.", "При N = 0."),
                @("λ*NMxy", "Все нагрузки, включая моменты от смещения N.", "Ничего.", "Хотя бы одна ненулевая нагрузка.", "Для полностью нулевого сочетания.")
            )
        }
    )}) | Out-Null
    $items.Add(@{ Key = "StabilityLoads"; Title = "Нагрузки для устойчивости"; Lines = @(
        "На Config в таблице «Нагрузки для устойчивости» задайте часть каждого полного сочетания: для СП 63 — постоянные и длительные нагрузки, для СП 35 — только постоянные. Методика выбирается настройкой Stability.Code.",
        "CombinationID связывает строку с полным сочетанием по имени, а не по номеру строки. N - сила, Mx/My - моменты; они вводятся в тех же INPUT и знаках и относятся к той же точке приложения. Смещения на Config задают Load.ReferenceOffsetX и Load.ReferenceOffsetY.",
        "При включённой устойчивости отсутствие нужного CombinationID даёт InputErr. В найденной строке пустые N, Mx и My означают 0; текст и ошибки формул недопустимы. При повторе ID берётся первая строка.",
        "Таблица используется при Calculation.Stability.Enabled = Yes. Выделение длительной части и проверка состава сочетания выполняются пользователем; автоматическое ограничение полной нагрузкой не предусмотрено.",
        "Ошибка этой таблицы дает InputErr по устойчивости с местом исправления. Прочность и трещины продолжаются с исходными усилиями; это не означает, что устойчивость проверена успешно."
    )}) | Out-Null
    $items.Add(@{ Key = "SP35Table721"; Title = "Таблица 7.21 СП 35"; Lines = @(
        "На Config таблица 7.21 СП 35 содержит коэффициенты для устойчивости при малом эксцентриситете. Для поддерживаемой ненапрягаемой арматуры используются верхние значения нормативной таблицы.",
        "l_0/i - гибкость: расчетная длина l_0, деленная на радиус инерции i. φ_m - коэффициент для кратковременной нагрузки, φ_l - для длительной. Индексы m/l обозначают эти две табличные величины.",
        "Расчетные колонки l_0/i, φ_m и φ_l должны содержать положительные числа. Гибкость должна строго возрастать сверху вниз. Пустое число или ошибка формулы требуют исправления указанной ячейки.",
        "l_0/b и l_0/d - справочные колонки; b и d обозначают размеры из нормативной таблицы. Они не участвуют в расчете. Промежуточные значения определяются линейной интерполяцией между соседними строками и колонками.",
        "Последняя строка служит для очень большой гибкости. При сборке она заполнена коэффициентами последней нормативной строки; измененные значения должны быть обоснованы пользователем.",
        "Таблица используется только в соответствующей ветви СП 35. В других ветвях она не меняет результат. Формула предельной силы и поправка при большой доле арматуры объяснены в разделе устойчивости."
    )}) | Out-Null
    $items.Add(@{ Key = "ConcreteMaterialParameters"; Title = "Параметры бетона"; Lines = @(
        "На Config задайте уже принятые характеристики бетона с учетом нужных коэффициентов условий работы. Программа не умножает их повторно.",
        "Concrete.R.ULS(I) задает R_b и R_bt для I группы; Concrete.R.SLS(II) - R_b,ser и R_bt,ser для II группы. Колонки соответствуют сжатию и растяжению.",
        "Настройка Concrete.E задает E_b - начальный модуль сжатия и E_bt - растяжения. По СП 63 п. 6.1.15 они равны. Раздельные колонки позволяют явно задать обе ветви. Модуль сдвига G = 0,4E_b не заменяет E_bt.",
        "Concrete.TwoThreeLine.Eb2 задает предельные деформации ε_b2 и ε_bt2 сжатой и растянутой ветвей.",
        "TwoLine означает двухлинейную диаграмму, ThreeLine - трехлинейную. Вид выбирается на Config настройками профиля MaterialModel.Strength.ConcreteDiagram, MaterialModel.CrackInitiation.ConcreteDiagram и MaterialModel.CrackedState.ConcreteDiagram для соответствующей проверки.",
        "Сопротивления и модули вводятся в INPUT Stress - единице напряжения. Деформации безразмерны. Серые контрольные точки и графики нужны для проверки человеком, а не для ручного задания расчетной кривой.",
        "Нормативные ссылки: СП 63 пп. 6.1.14, 6.1.20-6.1.24 и 6.1.26. Настройка Concrete.Rb.mc2 отдельно задает допустимое сжатие для проверки продольных трещин."
    )}) | Out-Null
    $items.Add(@{ Key = "SteelMaterialParameters"; Title = "Параметры арматуры"; Lines = @(
        "На Config задайте характеристики обычной ненапрягаемой арматуры. Разные материалы арматуры в одном сечении эта версия не поддерживает.",
        "Steel.R.ULS(I) задает R_sc и R_s для I группы; Steel.R.SLS(II) - R_sc,ser и R_s,ser для II группы. Раздельное R_sc,ser предусмотрено программой; в шаблоне принято R_sc,ser = R_s,ser.",
        "Настройка Steel.E задает E_sc и E_s - модули сжатой и растянутой арматуры. По СП 63 п. 6.2.12 они равны. Сопротивления и модули вводятся в INPUT Stress - единице напряжения.",
        "Steel.TwoLine.Es2 и Steel.ThreeLine.Es2 задают предельные деформации ε_sc2 и ε_s2 соответствующей диаграммы.",
        "TwoLine означает двухлинейную диаграмму, ThreeLine - трехлинейную. Вид выбирается на Config настройками профиля MaterialModel.Strength.SteelDiagram, MaterialModel.CrackInitiation.SteelDiagram и MaterialModel.CrackedState.SteelDiagram для соответствующего расчета.",
        "По СП 63 п. 6.2.13 TwoLine относится к физическому пределу текучести, ThreeLine - к условному. Производные точки строятся по пп. 6.2.11, 6.2.14 и 6.2.15; форму можно проверить по контрольным таблицам Config."
    )}) | Out-Null
    $items.Add(@{ Key = "CalculationProfiles"; Title = "Настройка расчетных профилей"; Lines = @(
        "Все перечисленные здесь настройки задаются на Config в таблице расчетных профилей. Каждая колонка PR1...PR4 - отдельный профиль. Profile.DisplayName задает его имя для выбора в сочетании, Profile.Description - короткое пояснение.",
        "Назначьте профиль строке сочетания в столбце ProfileId. Само заполнение колонки профиля расчет не запускает.",
        "Включите нужные проверки значением Yes, остальные отключите значением No. Calculation.Strength.DirectState задает текущее НДС по прочности, Calculation.Strength.Capacity — несущую способность, Calculation.Crack.Width — нормальные и продольные трещины, Calculation.Stability.Enabled — устойчивость и продольный изгиб.",
        "Если устойчивость не проходит, остальные запрошенные проверки продолжаются с исходными усилиями. Отказ устойчивости остается в общей сводке, даже если прочность или трещины прошли. При успешной проверке дальше используются усилия с учетом продольного изгиба.",
        "MaterialModel.Stability.ValueSet выбирает характеристики для устойчивости; нормативный вариант - ULS(I). Диаграммы этой проверке не назначаются.",
        "Для прочности настройка MaterialModel.Strength.ValueSet выбирает характеристики, MaterialModel.Strength.ConcreteDiagram и MaterialModel.Strength.SteelDiagram - диаграммы, MaterialModel.Strength.ConcreteTension - работу растянутого бетона.",
        "Для образования трещины эти четыре настройки начинаются с MaterialModel.CrackInitiation, для текущего НДС с трещинами - с MaterialModel.CrackedState. Полные ключи указаны в строках Config. При включенных трещинах обе модели обязательны.",
        "TwoLine означает двухлинейную диаграмму, ThreeLine — трехлинейную. Ignore исключает растянутую ветвь бетона, UseDiagram учитывает ее. Для трещин применяется SLS(II): образование проверяется с растянутым бетоном, текущее НДС — без него.",
        "Visualization.State выбирает состояние для схемы и AutoCAD: StrengthState - текущее по прочности; CapacityState - предельное; PreCrackState - перед образованием трещины; PostCrackState - сразу после; CrackedState - текущее с отключенным растянутым бетоном.",
        "Visualization.Quantity выбирает Stress - напряжения или Strain - деформации. Visualization.StressPrecision и Visualization.StrainPrecision задают число знаков после запятой в подписях, целое от 0 до 10. Они не округляют расчетные данные Results.",
        "Если выбранное состояние не получено, схема и экспорт сообщат причину, но не пересчитают его сами и не подменят другим. Без пригодной точки образования нет соответствующих PreCrackState/PostCrackState; текущее состояние рассчитывается независимо.",
        "Пустые и неизвестные активные настройки требуют исправления. Сообщение укажет профиль, точный ключ и текущую ячейку. Программа не выбирает модель вместо пользователя."
    )}) | Out-Null
    $items.Add(@{ Key = "MaterialDiagramControlTables"; Title = "Контрольные точки диаграмм"; Lines = @(
        "Таблицы на Config показывают точки двухлинейных TwoLine и трехлинейных ThreeLine диаграмм, полученные из введенных параметров бетона и арматуры. Соседние графики помогают проверить их форму.",
        "ε - относительная деформация, безразмерная; σ - напряжение в INPUT Stress. Единицу задает Config Units.Stress.Input. OUTPUT Stress не меняет эту контрольную таблицу.",
        "Расчет строит диаграммы по параметрам материалов, а не читает эти контрольные точки. Не редактируйте формулы таблицы: результаты от этого не изменятся, но контрольный график станет неверным.",
        "На графиках показаны диаграммы I группы предельных состояний. Материальные модели конкретного расчета выбирайте на Config в расчетных профилях."
    )}) | Out-Null
    $items.Add(@{ Key = "PlotAnnotationSettings"; Title = "Настройки оформления схемы"; Lines = @(
        "Таблица «Настройки оформления схемы» на Config управляет размерами, подписями арматуры и толщиной границ сечения на Excel-схеме. Колонка арматуры соответствует ключам Plot.RebarLabels.*, колонка размеров - Plot.Dimensions.*. Звездочка означает имя строки: например, Enabled дает Plot.RebarLabels.Enabled или Plot.Dimensions.Enabled.",
        "В строке Enabled выберите Yes, чтобы показать соответствующие подписи, или No, чтобы скрыть их. При No остальные параметры этого столбца не используются. Эти настройки не меняют расчет и экспорт в AutoCAD.",
        "Offset — базовый отступ в мм от бетона до размерной линии или линии подписи арматуры. Для обеих групп он умножается на max(1; D/1000), где D — наибольший габарит сечения в мм. При D ≤ 1000 мм коэффициент равен 1. Placement выбирает сторону подписи арматуры или текста размера.",
        "TextUnits выбирает pt - пункты Excel или mm - миллиметры сечения. TextHeight задает высоту текста, TextGap - зазор между текстом и линией в выбранных единицах. При pt размер шрифта почти не зависит от масштаба схемы; при mm он меняется вместе с геометрией.",
        "LineWeight - толщина основной линии в пунктах. Plot.RebarLabels.LineEnabled оставляет или убирает короткую линию рядом с подписью. Plot.Dimensions.ExtensionLineWeight задает толщину выносных линий, Plot.Dimensions.ArrowType и Plot.Dimensions.ArrowSize - вид и размер стрелок.",
        "Color задает цвет текста и основной линии; Plot.Dimensions.ExtensionLineColor - цвет выносных линий. Формат RGB: три целых компонента 0...255.",
        "Включенные подписи требуют корректных применимых настроек: TextHeight и LineWeight больше нуля, Offset и TextGap неотрицательны. Сообщение об ошибке указывает ячейку Config, которую нужно исправить."
    )}) | Out-Null
    $items.Add(@{ Key = "DependentRebarRows"; Title = "Как задать сечение и армирование"; Lines = @(
        "На листе Config выберите Geometry.Source = Generated и нужный тип Geometry.Type. Затем заполните его таблицу на листе «Сечения». В блоке General задаются размеры бетона, в блоке Main — основные ряды арматуры по сторонам, в блоке Extra — второй и третий ряды арматуры. Длины вводятся в единицах INPUT Length. Оранжевым выделены доступные поля ввода, в том числе необязательные.",
        "В блоке Main задайте диаметр стержней d и их количество n для каждой нужной стороны. Если любое из этих полей оставить пустым или указать 0, основной ряд арматуры и связанные с ним дополнительные ряды не создаются. Количество стержней n у отверстия HollowRectangle вводить не нужно: оно определяется проекцией наружного ряда. Противоположные стороны настраиваются независимо.",
        "Для включённого ряда арматуры заполните поле «До оси» (as): это расстояние от границы бетона до центров стержней, которое должно быть больше половины диаметра d. В полях «Отступ начала» (s0) и «Отступ конца» (s1) задайте отступы вдоль линии раскладки арматуры; если отступ не нужен, укажите 0. Количество стержней должно быть целым положительным числом, для круга — не меньше 2. Направления раскладки указаны в справке соответствующего типа сечения.",
        "Дополнительные ряды арматуры задаются в блоке Extra таблицы. Пустое поле «Диаметр 2» (d2) или «Диаметр 3» (d3), а также значение 0 отключают соответствующий ряд арматуры. Если диаметр задан, выберите привязку и положение этого ряда. EachBar создаёт стержни у каждого стержня основного ряда, EverySecondBar — у первого, третьего и следующих нечётных. Оба дополнительных ряда связаны с основным, поэтому третий ряд можно включить без второго. Настройки отключённого ряда не используются.",
        "Поля «Смещение 2» (a2) и «Смещение 3» (a3) задают расстояния от основного ряда арматуры до базовых линий дополнительных рядов. Смещение направлено поперёк линии раскладки в бетон: у наружной границы — внутрь сечения, у отверстия — от пустоты в бетон. Пустое поле означает 0. Оба расстояния отсчитываются от основного ряда независимо: например, 100 и 200 не складываются друг с другом.",
        "При одинаковом базовом смещении стержни дополнительных рядов располагаются рядом друг с другом. Положение Stacked размещает их вглубь бетона, а SideBySide — вдоль линии раскладки арматуры. При нулевом смещении к этой группе относится и основной стержень; при положительном смещении группа начинается с первого включённого дополнительного ряда. К сумме радиусов соседних стержней добавлены 2,5 мм на выступы арматуры периодического профиля. Расчётные диаметры и площади стержней при этом не меняются.",
        "При смещении проверяется весь контур сечения. Если скругление исчезло, соседние участки соединяются без этой дуги. Стержни, которые не попали на допустимую смещённую линию, не добавляются; программа сообщает об этом предупреждением. Если стержни смещённых рядов соседних сторон пересеклись, остаётся первый добавленный стержень, также с предупреждением. Остальные недопустимые пересечения требуют исправления. В подписях арматуры учитываются только фактически добавленные стержни.",
        "Нажмите «Сформировать сечение»: геометрия сохранится в Results, а при Plot.Enabled = Yes обновится схема. Проверьте схему и предупреждения, затем нажмите «Выполнить расчёт». После изменения размеров, армирования или сетки сформируйте сечение заново: сами правки входной таблицы не заменяют сохранённую геометрию. В обязательных числовых полях укажите допустимый ноль явно, если величина отсутствует."
    )}) | Out-Null
    $items.Add(@{ Key = "CircleGeometry"; Title = "Круглое сечение"; RelatedKey = "DependentRebarRows"; Lines = @(
        "Для круглого сечения выберите Geometry.Type = Circle. В блоке General таблицы укажите наружный диаметр бетонного сечения D, который должен быть больше 0. В блоках Main и Extra используется одна строка Circle для раскладки арматуры по всей окружности. Отдельные стороны для круга не задаются.",
        "Если нужен основной ряд арматуры, в блоке Main задайте диаметр стержней d больше 0 и их целое количество n не меньше 2. Расстояние до оси as должно быть больше d/2 и меньше D/2. В полях концевых отступов s0 и s1 укажите неотрицательные значения. Если арматура не нужна, оставьте поле d или n пустым либо укажите 0; расстояние as и концевые отступы тогда не используются.",
        "Раскладка арматуры начинается над центром бетонного сечения, на «12 часах», и идёт по часовой стрелке. При s0 = s1 = 0 стержни равномерно распределяются по полной окружности без повторения начальной точки.",
        "Если s0 или s1 больше 0, стержни располагаются по оставшейся дуге, включая оба её конца. Оба отступа отмеряются вдоль окружности центров стержней; их сумма должна быть меньше π·(D − 2as). Чтобы использовать полную окружность, оставьте оба отступа равными 0.",
        "Дополнительные ряды арматуры задаются в блоке Extra таблицы, также в строке Circle. Отдельное количество стержней для них не указывается. Положение Stacked направлено к центру бетонного сечения, а SideBySide — по касательной в направлении обхода окружности. Одна подпись арматуры показывает суммарное количество стержней каждого диаметра во всех рядах.",
        "Подробнее: общие правила заполнения Main/Extra и формирования сечения."
    )}) | Out-Null
    $items.Add(@{ Key = "PointedSectionGeometry"; Title = "Клиновидное сечение"; RelatedKey = "DependentRebarRows"; Lines = @(
        "Для клиновидного сечения выберите Geometry.Type = PointedSection. В блоке General таблицы задайте ширину прямой базы B и высоту H больше 0. Для левой стороны Left и правой стороны Right отдельно укажите вынос боковой вершины W и радиус R1 на переходе от верхней и нижней граней к боковой стороне. Значения W и радиусов не могут быть отрицательными; 0 означает отсутствие выноса или скругления.",
        "При W = 0 боковая сторона остаётся вертикальной. Радиус R2 этой стороны не используется, поэтому его поле можно оставить пустым. При W > 0 сторона образует выступ из двух наклонных граней, и для его вершины нужно задать радиус R2. Значение R2 = 0 оставляет вершину острой. Левая и правая стороны могут иметь разные параметры.",
        "В блоках Main и Extra таблицы строки Top и Bottom задают верхний и нижний ряды арматуры, а Left и Right — ряды по боковым сторонам. Верхний и нижний ряды идут слева направо, оба боковых — сверху вниз. Сверху и снизу арматура располагается на прямом участке между скруглениями, сбоку — по всей боковой линии с её дугами. Отступы s0 и s1 отсчитываются вдоль соответствующей линии раскладки арматуры.",
        "Если скругления не помещаются при заданных ширине B, высоте H и выносе W, уменьшите радиусы или увеличьте размеры сечения. Если исчез прямой участок для включённого верхнего или нижнего ряда арматуры, измените радиусы либо отключите этот ряд, указав диаметр d или количество стержней n равным 0. При W = 0 и исчезнувшей вертикальной прямой боковой ряд арматуры может идти по оставшимся дугам.",
        "Подробнее: общие правила заполнения Main/Extra и формирования сечения."
    )}) | Out-Null
    $items.Add(@{ Key = "HollowRectangleGeometry"; Title = "Пустотелый прямоугольник"; RelatedKey = "DependentRebarRows"; Lines = @(
        "Для пустотелого прямоугольника выберите Geometry.Type = HollowRectangle. В блоке General таблицы задайте наружные ширину B и высоту H, а также ширину Bo и высоту Ho отверстия. Все четыре размера должны быть больше 0. Затем укажите радиусы углов R и Ro и смещения центра отверстия Xo вправо и Yo вверх. Если скругление или смещение не нужно, укажите 0. Отверстие должно целиком помещаться внутри наружного контура.",
        "В блоке Main таблицы строки Top, Bottom, Left и Right задают наружные ряды арматуры. Для каждого ряда укажите расстояние as, диаметр d, количество стержней n и концевые отступы s0 и s1. Строки OpeningTop, OpeningBottom, OpeningLeft и OpeningRight задают арматуру по соответствующим сторонам отверстия. Горизонтальные линии раскладки идут слева направо, вертикальные — снизу вверх. По этим направлениям отсчитываются отступы s0 и s1.",
        "Количество стержней n у отверстия вводить не нужно: в таблице стоит прочерк, а положения стержней получаются проекцией основного наружного ряда на линию раскладки у отверстия. Если проекция не попала на рабочую длину этой линии, стержень не добавляется. Для включённого ряда арматуры у отверстия задайте собственные d, as, s0 и s1. Без соответствующего наружного ряда он не создаётся. Пустое поле диаметра d или значение 0 отключает этот ряд арматуры.",
        "Расстояние до оси as нужно указать для всех четырёх сторон отверстия, даже если их ряды арматуры выключены: эти значения ограничивают линии раскладки у отверстия. Расстояния не могут быть отрицательными, а для включённого ряда as должно быть больше d/2. Дополнительные ряды арматуры необязательны и задаются в блоке Extra отдельно для наружных сторон и сторон отверстия. Если боковая прямая исчезла при включённой арматуре, уменьшите радиус или отступ либо отключите этот боковой ряд арматуры.",
        "Подробнее: общие правила заполнения Main/Extra и формирования сечения."
    )}) | Out-Null
    $items.Add(@{ Key = "RectanglesGeometry"; Title = "Набор прямоугольников (Г-сечение)"; RelatedKey = "DependentRebarRows"; Lines = @(
        "Для сечения из прямоугольников выберите Geometry.Type = Rectangles. В блоке General таблицы настройка Shape определяет форму: OneRect создаёт один прямоугольник, TwoRect — два смежных прямоугольника. Задайте высоту H1 и ширину B1 первого прямоугольника больше 0. В варианте TwoRect первый прямоугольник находится над вторым.",
        "Для варианта TwoRect также задайте положительные высоту H2 и ширину B2 второго прямоугольника. Параметр XOffset смещает левый край верхнего прямоугольника вправо относительно левого края нижнего. При 0 левые края совпадают, отрицательное значение смещает верхний прямоугольник влево. По ширине прямоугольники должны иметь общий участок: −B1 < XOffset < B2.",
        "Для варианта OneRect используются только размеры H1 и B1 и строки армирования R1. Размеры H2 и B2, смещение XOffset и все поля строк R2 в блоках Main и Extra не используются, поэтому их можно оставить пустыми.",
        "В блоках Main и Extra таблицы строки R1.Top, R1.Bottom, R1.Left и R1.Right задают арматуру по сторонам первого прямоугольника; такие же строки с началом R2 относятся ко второму. Верхние ряды арматуры идут слева направо, нижние — справа налево, левые — снизу вверх, правые — сверху вниз. Отступы s0 и s1 отмеряются от концов полной линии соответствующей грани.",
        "Арматуру для каждой строки таблицы задайте отдельно. На общей границе двух прямоугольников ряд арматуры не отключается автоматически. Если он не нужен, укажите для него диаметр d или количество стержней n равным 0. Дополнительные ряды арматуры необязательны; отдельное количество стержней для них не задаётся.",
        "Подробнее: общие правила заполнения Main/Extra и формирования сечения."
    )}) | Out-Null
    foreach ($section in (Get-SystemSettingsCatalog)) {
        foreach ($row in $section.Rows) {
            $key = [string]$row[0]
            if ($key.Length -gt 0 -and -not $key.StartsWith("[")) {
                $items.Add(@{ Key = $key; Title = $key; Lines = (Get-SettingInstructionLines $key ([string]$row[3])) }) | Out-Null
            }
        }
    }
    foreach ($row in (Get-PlotAnnotationSettingsCatalog)) {
        $key = [string]$row[0]
        $items.Add(@{ Key = $key; Title = "Настройки оформления схемы: $key"; Lines = (Get-SettingInstructionLines $key ([string]$row[4])) }) | Out-Null
    }
    foreach ($row in (Get-ConcreteMaterialParametersCatalog)) {
        $key = [string]$row[0]
        $items.Add(@{ Key = $key; Title = $key; Lines = (Get-SettingInstructionLines $key ([string]$row[4])) }) | Out-Null
    }
    foreach ($row in (Get-SteelMaterialParametersCatalog)) {
        $key = [string]$row[0]
        $items.Add(@{ Key = $key; Title = $key; Lines = (Get-SettingInstructionLines $key ([string]$row[4])) }) | Out-Null
    }

    # Справка по повторному оформлению Results завершает каталог согласно утверждённому макету листа.
    $finalHelpBlock = @($items | Where-Object { $_.Key -eq 'General.ReformatResultsEnabled' })
    if ($finalHelpBlock.Count -ne 1) { throw 'В каталоге справки должен быть один блок General.ReformatResultsEnabled.' }
    $items.Remove($finalHelpBlock[0]) | Out-Null
    $items.Add($finalHelpBlock[0]) | Out-Null
    $items
}

# Добавляет строку-заголовок в user guide на листе "Справка".
function Add-GuideTitle {
    param([object]$Sheet, [int]$Row, [string]$Text, [int]$FontSize = 13)

    $range = $Sheet.Range($Sheet.Cells.Item($Row, 1), $Sheet.Cells.Item($Row, 6))
    $range.Merge() | Out-Null
    $range.Value2 = $Text
    $range.Font.Bold = $true
    $range.Font.Size = $FontSize
    $range.Interior.Color = 15921906
    $range.HorizontalAlignment = -4131
    $range.VerticalAlignment = -4108
    $range.WrapText = $false
    $range.EntireRow.RowHeight = [Math]::Max(21, [Math]::Ceiling($FontSize * 1.6))
    return $Row + 1
}

# Разбивает длинный абзац справки на несколько строк.
# Excel плохо подбирает высоту объединенных ячеек при AutoFit, поэтому длинные
# пояснения надежнее хранить несколькими короткими строками, а не одной высокой.
function Split-GuideParagraphText {
    param([string]$Text, [int]$MaxLength = 200)

    $parts = New-Object System.Collections.Generic.List[string]
    $remaining = ([string]$Text).Trim()
    while ($remaining.Length -gt $MaxLength) {
        $breakAt = -1
        foreach ($separator in @(". ", "; ", ": ", ", ", " ")) {
            $candidate = $remaining.LastIndexOf($separator, [Math]::Min($MaxLength, $remaining.Length - 1))
            if ($candidate -gt $breakAt) {
                $breakAt = $candidate + $separator.Length
            }
        }
        if ($breakAt -le 0) {
            $breakAt = $MaxLength
        }

        $parts.Add($remaining.Substring(0, $breakAt).Trim()) | Out-Null
        $remaining = $remaining.Substring($breakAt).Trim()
    }

    if ($remaining.Length -gt 0) {
        $parts.Add($remaining) | Out-Null
    }
    return $parts
}

# Делит абзац на заданное число строк без изменения его слов и адресов следующих разделов.
# При нуле сохраняется обычное разбиение по200символов; фиксированное число строк согласовано с макетом книги.
function Split-GuideParagraphRows {
    param([string]$Text, [int]$RowCount = 0)

    if ($RowCount -le 0) { return Split-GuideParagraphText $Text }
    $parts = New-Object System.Collections.Generic.List[string]
    $words = @(([string]$Text).Trim() -split '\s+')
    for ($index = 0; $index -lt $RowCount - 1; $index++) {
        $remainingLength = ($words -join ' ').Length
        $remainingRows = $RowCount - $index
        $target = [Math]::Max(1, [Math]::Floor(($remainingLength + $remainingRows - 1) / $remainingRows))
        $take = 1
        while ($take -lt $words.Count -and (($words | Select-Object -First ($take + 1)) -join ' ').Length -le $target) { $take++ }
        $parts.Add((($words | Select-Object -First $take) -join ' ')) | Out-Null
        $words = @($words | Select-Object -Skip $take)
    }
    $parts.Add(($words -join ' ')) | Out-Null
    return $parts
}

# Добавляет обычный абзац справки; длинный текст заранее режется на несколько
# строк, чтобы в готовой книге он не обрезался внутри объединенной ячейки.
function Add-GuideParagraph {
    param([object]$Sheet, [int]$Row, [string]$Text, [switch]$Emphasis, [int]$RowCount = 0)

    $currentRow = $Row
    foreach ($line in (Split-GuideParagraphRows $Text $RowCount)) {
        $range = $Sheet.Range($Sheet.Cells.Item($currentRow, 1), $Sheet.Cells.Item($currentRow, 6))
        $range.Merge() | Out-Null
        # Абзац всегда текст: перенесенная часть может начинаться со знака =.
        $range.NumberFormat = "@"
        $range.Value2 = [string]$line
        $range.Font.Bold = [bool]$Emphasis
        $range.WrapText = $true
        $range.VerticalAlignment = -4160
        $currentRow++
    }
    return $currentRow
}

# Удаляет только картинки с префиксом NDMHelpFormula_ перед заполнением справки.
# Cells.Clear не удаляет Shapes; явная очистка предотвращает наложение
# формульных изображений при повторном построении листа.
function Clear-GuideFormulaPictures {
    param([object]$Sheet)

    try {
        for ($i = $Sheet.Shapes.Count; $i -ge 1; $i--) {
            $shape = $Sheet.Shapes.Item($i)
            if ([string]$shape.Name -like "NDMHelpFormula_*") {
                $shape.Delete()
            }
        }
    }
    catch {
        # Если в конкретной версии Excel доступ к Shapes временно недоступен,
        # справка все равно должна собраться; новые формулы будут записаны текстом.
    }
}

# Загружает стандартный .NET-рендеринг формул и подготавливает каталог PNG.
# Картинки надежнее OLE-уравнений при автоматической сборке xlsm, а вручную
# нарисованные дроби выглядят ближе к инженерной записи, чем плоский текст.
function Initialize-FormulaImageRenderer {
    try {
        if (-not (Test-Path -LiteralPath $script:FormulaImageDir)) {
            New-Item -ItemType Directory -Force -Path $script:FormulaImageDir | Out-Null
        }
        Add-Type -AssemblyName System.Drawing -ErrorAction Stop
        return $true
    }
    catch {
        return $false
    }
}

# Создает уникальный путь для очередной формулы. Номер нужен, чтобы Excel
# не переиспользовал кэш картинки при повторной сборке.
function New-FormulaImagePath {
    param([string]$Name)

    $script:FormulaImageCounter++
    $safeName = ($Name -replace "[^\p{L}\p{Nd}_-]+", "_")
    if ([string]::IsNullOrWhiteSpace($safeName)) { $safeName = "formula" }
    return (Join-Path $script:FormulaImageDir ("formula_{0:000}_{1}.png" -f $script:FormulaImageCounter, $safeName))
}

# Создает шрифт для картинки формулы. Cambria Math обычно установлен вместе
# с Office и хорошо отображает греческие символы и индексы; если его нет,
# .NET автоматически подберет доступный fallback.
function New-FormulaFont {
    param([single]$Size, [bool]$Bold = $false)

    $style = [System.Drawing.FontStyle]::Regular
    if ($Bold) { $style = [System.Drawing.FontStyle]::Bold }
    return New-Object System.Drawing.Font("Cambria Math", $Size, $style, [System.Drawing.GraphicsUnit]::Pixel)
}

# Разбирает компактную инженерную запись формулы на фрагменты.
# Символ "_" превращает следующий фрагмент в настоящий нижний индекс,
# а "^" - в верхний индекс. В сам PNG эти служебные символы не попадают.
function ConvertTo-MathRuns {
    param([string]$Text)

    $runs = New-Object System.Collections.Generic.List[object]
    $i = 0
    while ($i -lt $Text.Length) {
        $ch = $Text[$i]
        if ($ch -eq "_" -or $ch -eq "^") {
            $kind = if ($ch -eq "_") { "Sub" } else { "Sup" }
            $i++
            if ($i -ge $Text.Length) { break }

            $script = ""
            if ($Text[$i] -eq "{") {
                $i++
                while ($i -lt $Text.Length -and $Text[$i] -ne "}") {
                    $script += [string]$Text[$i]
                    $i++
                }
                if ($i -lt $Text.Length -and $Text[$i] -eq "}") { $i++ }
            }
            else {
                while ($i -lt $Text.Length) {
                    $c = [string]$Text[$i]
                    if ($c -match "[\p{L}\p{Nd},\.]") {
                        $script += $c
                        $i++
                    }
                    else {
                        break
                    }
                }
            }

            if (-not [string]::IsNullOrWhiteSpace($script)) {
                $runs.Add(@{ Kind = $kind; Text = $script }) | Out-Null
            }
        }
        else {
            $base = ""
            while ($i -lt $Text.Length -and $Text[$i] -ne "_" -and $Text[$i] -ne "^") {
                $base += [string]$Text[$i]
                $i++
            }
            if ($base.Length -gt 0) {
                $runs.Add(@{ Kind = "Base"; Text = $base }) | Out-Null
            }
        }
    }
    return ,$runs
}

# Измеряет формулу с индексами как единую строку. Это нужно, чтобы дроби и
# обычные формулы центрировались по реальной картинке, а не по сырому тексту.
function Measure-MathText {
    param([System.Drawing.Graphics]$Graphics, [string]$Text, [System.Drawing.Font]$BaseFont, [System.Drawing.Font]$ScriptFont)

    $runs = ConvertTo-MathRuns $Text
    $width = 0.0
    $baseSize = $Graphics.MeasureString("Mg", $BaseFont)
    $height = $baseSize.Height * 1.25
    for ($i = 0; $i -lt $runs.Count; $i++) {
        $run = $runs[$i]
        $font = if ($run.Kind -eq "Base") { $BaseFont } else { $ScriptFont }
        $size = $Graphics.MeasureString([string]$run.Text, $font)
        $advance = $size.Width
        # GDI+ MeasureString дает небольшой боковой запас. Для индексов это
        # выглядит как лишний пробел, поэтому поджимаем базовый символ к индексу.
        if ($run.Kind -eq "Base" -and $i + 1 -lt $runs.Count -and $runs[$i + 1].Kind -ne "Base") {
            $advance = [Math]::Max(0, $advance - $ScriptFont.Size * 0.25)
        }
        $width += $advance
    }
    return @{ Width = $width; Height = $height; Runs = $runs; BaseHeight = $baseSize.Height }
}

# Отделяет нормативную приписку от самой формулы. Формула остается крупной,
# а ссылка на СП/EN рисуется меньшим шрифтом, чтобы не спорить с выражением.
function Split-FormulaReference {
    param([string]$Text)

    $match = [regex]::Match($Text, "^(.*?)(\s{2,}(?:СП|EN)\s+.*)$")
    if ($match.Success) {
        return @{
            Formula = $match.Groups[1].Value.TrimEnd()
            Reference = $match.Groups[2].Value.Trim()
        }
    }
    return @{ Formula = $Text; Reference = "" }
}

# Рисует строку формулы с нижними/верхними индексами. Визуально это ближе к
# Excel Equation, но не требует OLE-объектов и не ломается при сборке.
function Draw-MathText {
    param(
        [System.Drawing.Graphics]$Graphics,
        [System.Collections.IEnumerable]$Runs,
        [System.Drawing.Font]$BaseFont,
        [System.Drawing.Font]$ScriptFont,
        [System.Drawing.Brush]$Brush,
        [double]$X,
        [double]$CenterY
    )

    $baseSize = $Graphics.MeasureString("Mg", $BaseFont)
    $scriptSize = $Graphics.MeasureString("Mg", $ScriptFont)
    $baseTop = $CenterY - $baseSize.Height / 2
    $currentX = $X
    $runArray = @($Runs)
    for ($i = 0; $i -lt $runArray.Count; $i++) {
        $run = $runArray[$i]
        if ($run.Kind -eq "Base") {
            $Graphics.DrawString([string]$run.Text, $BaseFont, $Brush, [single]$currentX, [single]$baseTop)
            $advance = $Graphics.MeasureString([string]$run.Text, $BaseFont).Width
            if ($i + 1 -lt $runArray.Count -and $runArray[$i + 1].Kind -ne "Base") {
                $advance = [Math]::Max(0, $advance - $ScriptFont.Size * 0.25)
            }
            $currentX += $advance
        }
        elseif ($run.Kind -eq "Sub") {
            $top = $baseTop + $baseSize.Height * 0.52
            $Graphics.DrawString([string]$run.Text, $ScriptFont, $Brush, [single]$currentX, [single]$top)
            $currentX += $Graphics.MeasureString([string]$run.Text, $ScriptFont).Width
        }
        elseif ($run.Kind -eq "Sup") {
            $top = $baseTop - $scriptSize.Height * 0.2
            $Graphics.DrawString([string]$run.Text, $ScriptFont, $Brush, [single]$currentX, [single]$top)
            $currentX += $Graphics.MeasureString([string]$run.Text, $ScriptFont).Width
        }
    }
    return $currentX - $X
}

# Рисует одну строку формулы в PNG с высоким внутренним разрешением.
function New-TextFormulaImage {
    param([string]$Text)

    if (-not (Initialize-FormulaImageRenderer)) { return $null }

    $font = New-FormulaFont 32
    $scriptFont = New-FormulaFont 21
    $referenceFont = New-FormulaFont 16
    $referenceScriptFont = New-FormulaFont 11
    $paddingX = 24
    $paddingY = 12
    $scale = 2.0
    $referenceGap = 18
    $parts = Split-FormulaReference $Text
    $measureBitmap = New-Object System.Drawing.Bitmap(10, 10)
    $measureGraphics = [System.Drawing.Graphics]::FromImage($measureBitmap)
    $measureGraphics.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit
    $math = Measure-MathText $measureGraphics $parts.Formula $font $scriptFont
    $referenceMath = Measure-MathText $measureGraphics $parts.Reference $referenceFont $referenceScriptFont
    $measureGraphics.Dispose()
    $measureBitmap.Dispose()

    $extraGap = if ($referenceMath.Width -gt 0) { $referenceGap } else { 0 }
    $width = [int][Math]::Ceiling($math.Width + $extraGap + $referenceMath.Width + 2 * $paddingX)
    $height = [int][Math]::Ceiling([Math]::Max($math.Height, $referenceMath.Height) + 2 * $paddingY)
    $bitmap = New-Object System.Drawing.Bitmap($width, $height)
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $graphics.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit
    $graphics.Clear([System.Drawing.Color]::FromArgb(250, 250, 252))
    $brush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(30, 30, 30))
    Draw-MathText $graphics $math.Runs $font $scriptFont $brush $paddingX ($height / 2) | Out-Null
    if ($referenceMath.Width -gt 0) {
        Draw-MathText $graphics $referenceMath.Runs $referenceFont $referenceScriptFont $brush ($paddingX + $math.Width + $referenceGap) ($height / 2) | Out-Null
    }

    $path = New-FormulaImagePath "text"
    $bitmap.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)
    $graphics.Dispose()
    $brush.Dispose()
    $bitmap.Dispose()
    $font.Dispose()
    $scriptFont.Dispose()
    $referenceFont.Dispose()
    $referenceScriptFont.Dispose()

    return @{ Path = $path; WidthPoints = $width / $scale; HeightPoints = $height / $scale }
}

# Рисует формулу с настоящей горизонтальной дробной чертой.
# Prefix и Suffix остаются обычными частями формулы, а числитель/знаменатель
# центрируются относительно одной дроби.
function New-FractionFormulaImage {
    param(
        [string]$Prefix,
        [string]$Numerator,
        [string]$Denominator,
        [string]$Suffix = ""
    )

    if (-not (Initialize-FormulaImageRenderer)) { return $null }

    $font = New-FormulaFont 32
    $scriptFont = New-FormulaFont 21
    $smallFont = New-FormulaFont 29
    $smallScriptFont = New-FormulaFont 19
    $referenceFont = New-FormulaFont 16
    $referenceScriptFont = New-FormulaFont 11
    $paddingX = 24
    $paddingY = 12
    $scale = 2.0
    $gap = 10
    $lineGap = 5
    $referenceGap = 18
    $suffixParts = Split-FormulaReference $Suffix
    $measureBitmap = New-Object System.Drawing.Bitmap(10, 10)
    $measureGraphics = [System.Drawing.Graphics]::FromImage($measureBitmap)
    $measureGraphics.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit

    $prefixMath = Measure-MathText $measureGraphics $Prefix $font $scriptFont
    $suffixMath = Measure-MathText $measureGraphics $suffixParts.Formula $font $scriptFont
    $referenceMath = Measure-MathText $measureGraphics $suffixParts.Reference $referenceFont $referenceScriptFont
    $numMath = Measure-MathText $measureGraphics $Numerator $smallFont $smallScriptFont
    $denMath = Measure-MathText $measureGraphics $Denominator $smallFont $smallScriptFont
    $fracWidth = [Math]::Max($numMath.Width, $denMath.Width) + 18
    $fracHeight = $numMath.Height + $denMath.Height + 2 * $lineGap
    $bodyHeight = [Math]::Max([Math]::Max([Math]::Max($prefixMath.Height, $suffixMath.Height), $referenceMath.Height), $fracHeight)
    $extraReferenceGap = if ($referenceMath.Width -gt 0) { $referenceGap } else { 0 }
    $width = [int][Math]::Ceiling($paddingX * 2 + $prefixMath.Width + $gap + $fracWidth + $gap + $suffixMath.Width + $extraReferenceGap + $referenceMath.Width)
    $height = [int][Math]::Ceiling($paddingY * 2 + $bodyHeight)
    $measureGraphics.Dispose()
    $measureBitmap.Dispose()

    $bitmap = New-Object System.Drawing.Bitmap($width, $height)
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $graphics.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit
    $graphics.Clear([System.Drawing.Color]::FromArgb(250, 250, 252))
    $brush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(30, 30, 30))
    $pen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(30, 30, 30), 2)

    $centerY = $paddingY + $bodyHeight / 2
    $x = $paddingX
    Draw-MathText $graphics $prefixMath.Runs $font $scriptFont $brush $x $centerY | Out-Null
    $x += $prefixMath.Width + $gap

    $fracX = $x
    $lineY = $centerY
    $numX = $fracX + ($fracWidth - $numMath.Width) / 2
    $denX = $fracX + ($fracWidth - $denMath.Width) / 2
    Draw-MathText $graphics $numMath.Runs $smallFont $smallScriptFont $brush $numX ($lineY - $lineGap - $numMath.Height / 2) | Out-Null
    $graphics.DrawLine($pen, $fracX, $lineY, $fracX + $fracWidth, $lineY)
    Draw-MathText $graphics $denMath.Runs $smallFont $smallScriptFont $brush $denX ($lineY + $lineGap + $denMath.Height / 2) | Out-Null
    $x += $fracWidth + $gap

    if (-not [string]::IsNullOrWhiteSpace($suffixParts.Formula)) {
        Draw-MathText $graphics $suffixMath.Runs $font $scriptFont $brush $x $centerY | Out-Null
        $x += $suffixMath.Width
    }
    if ($referenceMath.Width -gt 0) {
        $referenceX = $x
        if ($suffixMath.Width -gt 0) { $referenceX += $referenceGap }
        Draw-MathText $graphics $referenceMath.Runs $referenceFont $referenceScriptFont $brush $referenceX $centerY | Out-Null
    }

    $path = New-FormulaImagePath "fraction"
    $bitmap.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)
    $graphics.Dispose()
    $brush.Dispose()
    $pen.Dispose()
    $bitmap.Dispose()
    $font.Dispose()
    $scriptFont.Dispose()
    $smallFont.Dispose()
    $smallScriptFont.Dispose()
    $referenceFont.Dispose()
    $referenceScriptFont.Dispose()

    return @{ Path = $path; WidthPoints = $width / $scale; HeightPoints = $height / $scale }
}

# Рисует составную формулу, где в одной строке может быть несколько настоящих
# дробей. Нужна для методических блоков СП 35/СП 63, где текстовая запись
# вида a/b внутри длинной формулы читается плохо.
# Radical и DenominatorRadical сохраняют настоящую черту корня над всем
# выражением, в том числе индексами и степенями, а не печатают один знак √.
function New-CompositeFormulaImage {
    param(
        [object[]]$Segments,
        [string]$Reference = ""
    )

    if (-not (Initialize-FormulaImageRenderer)) { return $null }

    $font = New-FormulaFont 32
    $scriptFont = New-FormulaFont 21
    $smallFont = New-FormulaFont 29
    $smallScriptFont = New-FormulaFont 19
    $referenceFont = New-FormulaFont 16
    $referenceScriptFont = New-FormulaFont 11
    $paddingX = 24
    $paddingY = 12
    $scale = 2.0
    $segmentGap = 2
    $fractionLineGap = 5
    $referenceGap = 18

    $measureBitmap = New-Object System.Drawing.Bitmap(10, 10)
    $measureGraphics = [System.Drawing.Graphics]::FromImage($measureBitmap)
    $measureGraphics.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit

    $layouts = @()
    $bodyWidth = 0.0
    $bodyHeight = 0.0
    foreach ($segment in $Segments) {
        $kind = [string]($segment["Kind"])
        if ($kind -eq "Fraction") {
            $numMath = Measure-MathText $measureGraphics ([string]$segment["Numerator"]) $smallFont $smallScriptFont
            $denMath = Measure-MathText $measureGraphics ([string]$segment["Denominator"]) $smallFont $smallScriptFont
            $denRadical = [bool]$segment["DenominatorRadical"]
            $denWidth = $denMath.Width
            $denHeight = $denMath.Height
            if ($denRadical) { $denWidth += 28; $denHeight += 10 }
            $fracWidth = [Math]::Max($numMath.Width, $denWidth) + 18
            $fracHeight = $numMath.Height + $denHeight + 2 * $fractionLineGap
            $layout = @{
                Kind = "Fraction"
                NumMath = $numMath
                DenMath = $denMath
                DenRadical = $denRadical
                DenWidth = $denWidth
                DenHeight = $denHeight
                Width = $fracWidth
                Height = $fracHeight
            }
        }
        elseif ($kind -eq "Radical") {
            $textMath = Measure-MathText $measureGraphics ([string]$segment["Radicand"]) $font $scriptFont
            $layout = @{ Kind = "Radical"; Math = $textMath; Width = $textMath.Width + 28; Height = $textMath.Height + 10 }
        }
        else {
            $textMath = Measure-MathText $measureGraphics ([string]$segment["Text"]) $font $scriptFont
            $layout = @{
                Kind = "Text"
                Math = $textMath
                Width = $textMath.Width
                Height = $textMath.Height
            }
        }
        $layouts += $layout
        $bodyWidth += [double]$layout.Width
        $bodyHeight = [Math]::Max($bodyHeight, [double]$layout.Height)
    }
    if ($layouts.Count -gt 1) { $bodyWidth += $segmentGap * ($layouts.Count - 1) }

    $referenceMath = Measure-MathText $measureGraphics $Reference $referenceFont $referenceScriptFont
    $extraReferenceGap = if ($referenceMath.Width -gt 0) { $referenceGap } else { 0 }
    $width = [int][Math]::Ceiling($paddingX * 2 + $bodyWidth + $extraReferenceGap + $referenceMath.Width)
    $height = [int][Math]::Ceiling($paddingY * 2 + [Math]::Max($bodyHeight, $referenceMath.Height))
    $measureGraphics.Dispose()
    $measureBitmap.Dispose()

    $bitmap = New-Object System.Drawing.Bitmap($width, $height)
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $graphics.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit
    $graphics.Clear([System.Drawing.Color]::FromArgb(250, 250, 252))
    $brush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(30, 30, 30))
    $pen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(30, 30, 30), 2)

    $centerY = $paddingY + [Math]::Max($bodyHeight, $referenceMath.Height) / 2
    $x = $paddingX
    foreach ($layout in $layouts) {
        if ($layout.Kind -eq "Fraction") {
            $fracX = $x
            $lineY = $centerY
            $numX = $fracX + ($layout.Width - $layout.NumMath.Width) / 2
            $denX = $fracX + ($layout.Width - $layout.DenWidth) / 2
            Draw-MathText $graphics $layout.NumMath.Runs $smallFont $smallScriptFont $brush $numX ($lineY - $fractionLineGap - $layout.NumMath.Height / 2) | Out-Null
            $graphics.DrawLine($pen, $fracX, $lineY, $fracX + $layout.Width, $lineY)
            $denCenterY = $lineY + $fractionLineGap + $layout.DenHeight / 2
            if ($layout.DenRadical) {
                Draw-RadicalMath $graphics $layout.DenMath $smallFont $smallScriptFont $brush $pen $denX $denCenterY
            }
            else { Draw-MathText $graphics $layout.DenMath.Runs $smallFont $smallScriptFont $brush $denX $denCenterY | Out-Null }
        }
        elseif ($layout.Kind -eq "Radical") {
            Draw-RadicalMath $graphics $layout.Math $font $scriptFont $brush $pen $x $centerY
        }
        else {
            Draw-MathText $graphics $layout.Math.Runs $font $scriptFont $brush $x $centerY | Out-Null
        }
        $x += [double]$layout.Width + $segmentGap
    }
    if ($referenceMath.Width -gt 0) {
        Draw-MathText $graphics $referenceMath.Runs $referenceFont $referenceScriptFont $brush ($paddingX + $bodyWidth + $referenceGap) $centerY | Out-Null
    }

    $path = New-FormulaImagePath "composite"
    $bitmap.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)
    $graphics.Dispose()
    $brush.Dispose()
    $pen.Dispose()
    $bitmap.Dispose()
    $font.Dispose()
    $scriptFont.Dispose()
    $smallFont.Dispose()
    $smallScriptFont.Dispose()
    $referenceFont.Dispose()
    $referenceScriptFont.Dispose()

    return @{ Path = $path; WidthPoints = $width / $scale; HeightPoints = $height / $scale }
}

# Рисует математический корень с общей верхней чертой и готовыми индексами.
# Layout заранее резервирует размер, поэтому знак и степень не обрезаются.
function Draw-RadicalMath {
    param([System.Drawing.Graphics]$Graphics, [object]$Math, [System.Drawing.Font]$Font,
        [System.Drawing.Font]$ScriptFont, [System.Drawing.Brush]$Brush, [System.Drawing.Pen]$Pen,
        [double]$X, [double]$CenterY)
    $top = $CenterY - $Math.Height / 2 - 2
    $Graphics.DrawLine($Pen, [single]($X + 1), [single]($CenterY + 1), [single]($X + 6), [single]($CenterY - 3))
    $Graphics.DrawLine($Pen, [single]($X + 6), [single]($CenterY - 3), [single]($X + 12), [single]($CenterY + 12))
    $Graphics.DrawLine($Pen, [single]($X + 12), [single]($CenterY + 12), [single]($X + 20), [single]$top)
    $Graphics.DrawLine($Pen, [single]($X + 20), [single]$top, [single]($X + 28 + $Math.Width), [single]$top)
    Draw-MathText $Graphics $Math.Runs $Font $ScriptFont $Brush ($X + 24) ($CenterY + 3) | Out-Null
}

# Делит строку матрицы на ячейки. В исходных строках генератора отдельные
# элементы матрицы разделяются несколькими пробелами, а одинарные пробелы
# остаются внутри выражения, например "N_int - N".
function Split-MatrixFormulaCells {
    param([string]$RowText)

    $cells = @()
    foreach ($cell in ([regex]::Split($RowText.Trim(), "\s{2,}"))) {
        if (-not [string]::IsNullOrWhiteSpace($cell)) {
            $cells += $cell.Trim()
        }
    }
    if ($cells.Count -eq 0 -and -not [string]::IsNullOrWhiteSpace($RowText)) {
        $cells += $RowText.Trim()
    }
    return ,$cells
}

# Измеряет одну ячейку матрицы. Если ячейка записана через "/", она считается
# дробью и дальше рисуется с настоящей горизонтальной чертой.
function Measure-MatrixFormulaCell {
    param(
        [System.Drawing.Graphics]$Graphics,
        [string]$Text,
        [System.Drawing.Font]$Font,
        [System.Drawing.Font]$ScriptFont
    )

    $slashIndex = $Text.IndexOf("/")
    if ($slashIndex -gt 0 -and $slashIndex -lt ($Text.Length - 1) -and $Text.IndexOf("/", $slashIndex + 1) -lt 0) {
        $numerator = $Text.Substring(0, $slashIndex).Trim()
        $denominator = $Text.Substring($slashIndex + 1).Trim()
        if (-not [string]::IsNullOrWhiteSpace($numerator) -and -not [string]::IsNullOrWhiteSpace($denominator)) {
            $numMath = Measure-MathText $Graphics $numerator $Font $ScriptFont
            $denMath = Measure-MathText $Graphics $denominator $Font $ScriptFont
            $lineGap = 5
            return @{
                Kind = "Fraction"
                NumMath = $numMath
                DenMath = $denMath
                LineGap = $lineGap
                Width = [Math]::Max($numMath.Width, $denMath.Width) + 16
                Height = $numMath.Height + $denMath.Height + 2 * $lineGap
            }
        }
    }

    $math = Measure-MathText $Graphics $Text $Font $ScriptFont
    return @{
        Kind = "Text"
        Math = $math
        Width = $math.Width
        Height = $math.Height
    }
}

# Рисует одну уже измеренную ячейку матрицы. Для дроби центрирует числитель
# и знаменатель относительно общей дробной черты.
function Draw-MatrixFormulaCell {
    param(
        [System.Drawing.Graphics]$Graphics,
        [hashtable]$Layout,
        [System.Drawing.Font]$Font,
        [System.Drawing.Font]$ScriptFont,
        [System.Drawing.Brush]$Brush,
        [System.Drawing.Pen]$Pen,
        [double]$X,
        [double]$CenterY
    )

    if ($Layout.Kind -eq "Fraction") {
        $lineY = $CenterY
        $numX = $X + ($Layout.Width - $Layout.NumMath.Width) / 2
        $denX = $X + ($Layout.Width - $Layout.DenMath.Width) / 2
        Draw-MathText $Graphics $Layout.NumMath.Runs $Font $ScriptFont $Brush $numX ($lineY - $Layout.LineGap - $Layout.NumMath.Height / 2) | Out-Null
        $Graphics.DrawLine($Pen, $X, $lineY, $X + $Layout.Width, $lineY)
        Draw-MathText $Graphics $Layout.DenMath.Runs $Font $ScriptFont $Brush $denX ($lineY + $Layout.LineGap + $Layout.DenMath.Height / 2) | Out-Null
        return
    }

    Draw-MathText $Graphics $Layout.Math.Runs $Font $ScriptFont $Brush $X $CenterY | Out-Null
}

# Рисует многострочную векторно-матричную формулу как PNG.
# Строки вида "R = [ ... ]" превращаются в одну матрицу с общей большой
# скобкой, чтобы в справке оставалась инженерная запись, а не псевдотекст.
function New-MatrixFormulaImage {
    param([string[]]$Lines)

    if (-not (Initialize-FormulaImageRenderer)) { return $null }

    $font = New-FormulaFont 32
    $scriptFont = New-FormulaFont 21
    $paddingX = 28
    $paddingY = 16
    $scale = 2.0
    $rowGap = 6
    $prefixGap = 16
    $bracketWidth = 14
    $matrixPad = 14
    $columnGap = 26

    $measureBitmap = New-Object System.Drawing.Bitmap(10, 10)
    $measureGraphics = [System.Drawing.Graphics]::FromImage($measureBitmap)
    $measureGraphics.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit

    $matrixRows = @()
    $prefixText = ""
    $prefixRow = 0
    $fallbackLines = @()
    for ($i = 0; $i -lt $Lines.Count; $i++) {
        $line = [string]$Lines[$i]
        $leftBracket = $line.IndexOf("[")
        $rightBracket = $line.LastIndexOf("]")
        if ($leftBracket -ge 0 -and $rightBracket -gt $leftBracket) {
            $prefix = $line.Substring(0, $leftBracket).Trim()
            if (-not [string]::IsNullOrWhiteSpace($prefix)) {
                $prefixText = $prefix
                $prefixRow = $i
            }
            $matrixRows += $line.Substring($leftBracket + 1, $rightBracket - $leftBracket - 1).Trim()
        }
        else {
            $fallbackLines += $line
        }
    }

    if ($matrixRows.Count -eq 0 -or $fallbackLines.Count -gt 0) {
        $lineLayouts = @()
        $maxWidth = 0.0
        $lineHeight = 0.0
        foreach ($line in $Lines) {
            $math = Measure-MathText $measureGraphics $line $font $scriptFont
            $lineLayouts += $math
            $maxWidth = [Math]::Max($maxWidth, $math.Width)
            $lineHeight = [Math]::Max($lineHeight, $math.Height)
        }
        $width = [int][Math]::Ceiling($paddingX * 2 + $maxWidth)
        $height = [int][Math]::Ceiling($paddingY * 2 + $lineLayouts.Count * $lineHeight + [Math]::Max(0, $lineLayouts.Count - 1) * $rowGap)
        $measureGraphics.Dispose()
        $measureBitmap.Dispose()

        $bitmap = New-Object System.Drawing.Bitmap($width, $height)
        $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
        $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
        $graphics.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit
        $graphics.Clear([System.Drawing.Color]::FromArgb(250, 250, 252))
        $brush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(30, 30, 30))

        $y = $paddingY + $lineHeight / 2
        foreach ($layout in $lineLayouts) {
            Draw-MathText $graphics $layout.Runs $font $scriptFont $brush (($width - $layout.Width) / 2) $y | Out-Null
            $y += $lineHeight + $rowGap
        }

        $path = New-FormulaImagePath "matrix"
        $bitmap.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)
        $graphics.Dispose()
        $brush.Dispose()
        $bitmap.Dispose()
        $font.Dispose()
        $scriptFont.Dispose()
        return @{ Path = $path; WidthPoints = $width / $scale; HeightPoints = $height / $scale }
    }

    $rowLayouts = @()
    $columnWidths = @()
    $rowHeights = @()
    $maxColumnCount = 0
    foreach ($matrixRow in $matrixRows) {
        $cells = Split-MatrixFormulaCells $matrixRow
        $cellLayouts = @()
        $rowHeight = 0.0
        for ($j = 0; $j -lt $cells.Count; $j++) {
            $cellLayout = Measure-MatrixFormulaCell $measureGraphics $cells[$j] $font $scriptFont
            $cellLayouts += $cellLayout
            $rowHeight = [Math]::Max($rowHeight, [double]$cellLayout.Height)
            if ($columnWidths.Count -le $j) {
                $columnWidths += [double]$cellLayout.Width
            }
            else {
                $columnWidths[$j] = [Math]::Max([double]$columnWidths[$j], [double]$cellLayout.Width)
            }
        }
        $rowLayouts += ,$cellLayouts
        $rowHeights += $rowHeight
        $maxColumnCount = [Math]::Max($maxColumnCount, $cells.Count)
    }
    $prefixMath = Measure-MathText $measureGraphics $prefixText $font $scriptFont
    $prefixWidth = if ([string]::IsNullOrWhiteSpace($prefixText)) { 0.0 } else { $prefixMath.Width + $prefixGap }
    $matrixContentWidth = 0.0
    foreach ($columnWidth in $columnWidths) {
        $matrixContentWidth += [double]$columnWidth
    }
    if ($columnWidths.Count -gt 1) {
        $matrixContentWidth += $columnGap * ($columnWidths.Count - 1)
    }
    $matrixHeight = 0.0
    foreach ($heightValue in $rowHeights) {
        $matrixHeight += [double]$heightValue
    }
    if ($rowHeights.Count -gt 1) {
        $matrixHeight += $rowGap * ($rowHeights.Count - 1)
    }
    $width = [int][Math]::Ceiling($paddingX * 2 + $prefixWidth + $bracketWidth * 2 + $matrixPad * 2 + $matrixContentWidth)
    $height = [int][Math]::Ceiling($paddingY * 2 + $matrixHeight)
    $measureGraphics.Dispose()
    $measureBitmap.Dispose()

    $bitmap = New-Object System.Drawing.Bitmap($width, $height)
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $graphics.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit
    $graphics.Clear([System.Drawing.Color]::FromArgb(250, 250, 252))
    $brush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(30, 30, 30))
    $pen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(30, 30, 30), 2)

    $matrixTop = $paddingY
    $matrixBottom = $height - $paddingY
    $matrixLeft = $paddingX + $prefixWidth
    $contentLeft = $matrixLeft + $bracketWidth + $matrixPad
    $rightBracketX = $contentLeft + $matrixContentWidth + $matrixPad

    if (-not [string]::IsNullOrWhiteSpace($prefixText)) {
        $prefixCenterY = $matrixTop
        for ($i = 0; $i -lt $prefixRow; $i++) {
            $prefixCenterY += [double]$rowHeights[$i] + $rowGap
        }
        $prefixCenterY += [double]$rowHeights[$prefixRow] / 2
        Draw-MathText $graphics $prefixMath.Runs $font $scriptFont $brush $paddingX $prefixCenterY | Out-Null
    }

    # Левая и правая квадратные скобки рисуются общей высоты матрицы.
    $graphics.DrawLine($pen, $matrixLeft + $bracketWidth, $matrixTop, $matrixLeft, $matrixTop)
    $graphics.DrawLine($pen, $matrixLeft, $matrixTop, $matrixLeft, $matrixBottom)
    $graphics.DrawLine($pen, $matrixLeft, $matrixBottom, $matrixLeft + $bracketWidth, $matrixBottom)
    $graphics.DrawLine($pen, $rightBracketX, $matrixTop, $rightBracketX + $bracketWidth, $matrixTop)
    $graphics.DrawLine($pen, $rightBracketX + $bracketWidth, $matrixTop, $rightBracketX + $bracketWidth, $matrixBottom)
    $graphics.DrawLine($pen, $rightBracketX + $bracketWidth, $matrixBottom, $rightBracketX, $matrixBottom)

    $yTop = $matrixTop
    for ($i = 0; $i -lt $rowLayouts.Count; $i++) {
        $cellLayouts = $rowLayouts[$i]
        $centerY = $yTop + [double]$rowHeights[$i] / 2
        $x = $contentLeft
        for ($j = 0; $j -lt $cellLayouts.Count; $j++) {
            $cellLayout = $cellLayouts[$j]
            $cellX = $x + ([double]$columnWidths[$j] - [double]$cellLayout.Width) / 2
            Draw-MatrixFormulaCell $graphics $cellLayout $font $scriptFont $brush $pen $cellX $centerY
            $x += [double]$columnWidths[$j] + $columnGap
        }
        $yTop += [double]$rowHeights[$i] + $rowGap
    }

    $path = New-FormulaImagePath "matrix"
    $bitmap.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)
    $graphics.Dispose()
    $brush.Dispose()
    $pen.Dispose()
    $bitmap.Dispose()
    $font.Dispose()
    $scriptFont.Dispose()

    return @{ Path = $path; WidthPoints = $width / $scale; HeightPoints = $height / $scale }
}

# Вставляет PNG формулы в объединенную строку справки. Если картинку создать
# не удалось, вызывающий метод заранее оставляет в ячейке текстовую формулу.
function Add-GuideFormulaPicture {
    param([object]$Sheet, [int]$Row, [hashtable]$ImageInfo, [string]$FallbackText)

    $range = $Sheet.Range($Sheet.Cells.Item($Row, 1), $Sheet.Cells.Item($Row, 6))
    $range.Merge() | Out-Null
    $range.Value2 = $FallbackText
    $range.Font.Name = "Arial"
    $range.Font.Size = 12
    $range.Font.Italic = $true
    $range.Interior.Color = 16448250
    $range.HorizontalAlignment = -4108
    $range.VerticalAlignment = -4108

    if ($null -eq $ImageInfo -or -not (Test-Path -LiteralPath $ImageInfo.Path)) {
        return $Row + 1
    }

    $maxWidth = [double]$range.Width - 16
    $width = [double]$ImageInfo.WidthPoints
    $height = [double]$ImageInfo.HeightPoints
    if ($width -gt $maxWidth -and $maxWidth -gt 0) {
        $factor = $maxWidth / $width
        $width *= $factor
        $height *= $factor
    }

    $rowHeight = [Math]::Max(36, $height + 8)
    $Sheet.Rows.Item($Row).RowHeight = $rowHeight
    $range.Value2 = ""

    $left = [double]$range.Left + ([double]$range.Width - $width) / 2
    $top = [double]$range.Top + ($rowHeight - $height) / 2
    $picture = $Sheet.Shapes.AddPicture($ImageInfo.Path, $false, $true, $left, $top, $width, $height)
    $picture.Name = "NDMHelpFormula_" + ("{0:000}" -f $script:FormulaImageCounter)
    $picture.AlternativeText = $FallbackText
    # xlMoveAndSize: формула ведет себя как содержимое строки и скрывается вместе с группой справки.
    $picture.Placement = 1
    $script:FormulaImageRows.Add(@{ Row = $Row; Height = $rowHeight; Name = $picture.Name; PictureWidth = $width; PictureHeight = $height }) | Out-Null
    return $Row + 1
}

# Добавляет компактную формулу в отдельном выделенном блоке.
function Add-GuideFormula {
    param([object]$Sheet, [int]$Row, [string]$Text)

    $image = New-TextFormulaImage $Text
    return (Add-GuideFormulaPicture $Sheet $Row $image $Text)
}

# Добавляет многострочный формульный блок для векторно-матричных записей.
# Такие формулы вставляются картинками с общей матричной скобкой, как и остальные
# важные формулы справки, а текст остается только fallback-ом при сбое рендера.
function Add-GuideMatrixFormula {
    param([object]$Sheet, [int]$Row, [string[]]$Lines)

    $fallback = ($Lines -join "`r`n")
    $image = New-MatrixFormulaImage $Lines
    return (Add-GuideFormulaPicture $Sheet $Row $image $fallback)
}

# Рисует ключевую формулу с дробью в виде картинки. Это сохраняет внешний вид
# полноценной инженерной записи и не зависит от OLE Equation Editor.
function Add-GuideFractionFormula {
    param(
        [object]$Sheet,
        [int]$Row,
        [string]$Prefix,
        [string]$Numerator,
        [string]$Denominator,
        [string]$Suffix = ""
    )

    $fallback = ($Prefix + " (" + $Numerator + ") / (" + $Denominator + ") " + $Suffix).Trim()
    $image = New-FractionFormulaImage $Prefix $Numerator $Denominator $Suffix
    return (Add-GuideFormulaPicture $Sheet $Row $image $fallback)
}

# Добавляет составную формулу с несколькими дробями. Fallback остается
# человекочитаемым текстом на случай, если PNG-рендер недоступен.
function Add-GuideCompositeFormula {
    param(
        [object]$Sheet,
        [int]$Row,
        [object[]]$Segments,
        [string]$Reference = "",
        [string]$FallbackText = ""
    )

    if ([string]::IsNullOrWhiteSpace($FallbackText)) {
        $pieces = @()
        foreach ($segment in $Segments) {
            if ([string]($segment["Kind"]) -eq "Fraction") {
                $pieces += "((" + [string]$segment["Numerator"] + ") / (" + [string]$segment["Denominator"] + "))"
            }
            else {
                $pieces += [string]$segment["Text"]
            }
        }
        $FallbackText = ($pieces -join "")
        if (-not [string]::IsNullOrWhiteSpace($Reference)) {
            $FallbackText = $FallbackText + "      " + $Reference
        }
    }

    $image = New-CompositeFormulaImage $Segments $Reference
    return (Add-GuideFormulaPicture $Sheet $Row $image $FallbackText)
}

# Добавляет крупный раздел справки с обычным заголовком и группируемым телом.
# Заголовок остается видимым всегда, а строки тела можно свернуть/развернуть
# стандартным Excel outline.
function Add-CollapsibleGuideSection {
    param(
        [object]$Sheet,
        [int]$StartRow,
        [string]$Title,
        [scriptblock]$BodyBuilder
    )

    $row = Add-GuideTitle $Sheet $StartRow $Title 16
    $bodyStartRow = $row
    $row = & $BodyBuilder $Sheet $row
    $bodyEndRow = $row - 1

    if ($bodyEndRow -ge $bodyStartRow) {
        try {
            $Sheet.Range($Sheet.Cells.Item($bodyStartRow, 1), $Sheet.Cells.Item($bodyEndRow, 1)).EntireRow.Group() | Out-Null
        }
        catch {
            # Если группировка недоступна в конкретном COM-сеансе Excel,
            # сама справка должна остаться читаемой без outline.
        }
    }

    return $row
}

# Формирует методический блок по поиску НДС от текущего сочетания.
# Это описание прямого CSectionSolver: как из заданных N/Mx/My получается
# равновесная плоскость деформаций и какие настройки Config влияют на процесс.
function Add-DirectStateMethodologyGuide {
    param([object]$Sheet, [int]$StartRow)

    $row = Add-GuideParagraph $Sheet $StartRow "Прямой расчет определяет НДС сечения при заданных N, Mx и My и проверяет равновесие и предельные деформации материалов."
    $row = Add-GuideTitle $Sheet $row "1. Расчетный профиль и материалы" 13
    $row = Add-GuideParagraph $Sheet $row "На Config в расчетном профиле включите Calculation.Strength.DirectState = Yes. Материалы этой проверки задаются там же настройками MaterialModel.Strength.ValueSet, MaterialModel.Strength.ConcreteDiagram, MaterialModel.Strength.ConcreteTension и MaterialModel.Strength.SteelDiagram: набор характеристик, диаграммы бетона и стали, работа растянутого бетона."
    $row = Add-GuideParagraph $Sheet $row "Ошибки обязательных настроек сопровождаются ключом, текущим адресом Config и требуемым исправлением. Удаленную строку следует восстановить в указанном разделе, утраченную ссылку на таблицу - через «Формулы → Диспетчер имен»."
    $row = Add-GuideParagraph $Sheet $row "Таблицы допускается перемещать с сохранением именованного диапазона на всю таблицу, включая заголовок. Диапазон должен оставаться непрерывным. Адреса диагностики определяются при чтении данных; комментарии ранее сохраненного Results относятся к предыдущему запуску."
    $row = Add-GuideTitle $Sheet $row "2. Нагрузки и поле деформаций" 13
    $row = Add-GuideParagraph $Sheet $row "Внесите N, Mx и My в таблицу сочетаний на Config в единицах INPUT. Пользовательская точка приложения отсчитывается от центра бетона без арматуры. Ее смещения задаются настройками Config Load.ReferenceOffsetX и Load.ReferenceOffsetY."
    $row = Add-GuideFormula $Sheet $row "M_x = M_{x,user} + N·(Y_ref - Y_calc)"
    $row = Add-GuideParagraph $Sheet $row "M_x,user - заданный момент; Y_ref и Y_calc - координаты пользовательской и расчетной точек приложения. M_x - момент после переноса к расчетной точке."
    $row = Add-GuideFormula $Sheet $row "M_y = M_{y,user} + N·(X_ref - X_calc)"
    $row = Add-GuideParagraph $Sheet $row "X_ref и X_calc - координаты тех же точек по X. Моменты от эксцентриситета N учитываются при переносе усилий."
    $row = Add-GuideFormula $Sheet $row "ε(x,y) = ε_0 + κ_x·y + κ_y·x"
    $row = Add-GuideParagraph $Sheet $row "ε_0 - деформация в начале координат; κ_x и κ_y - кривизны. Во внутренней системе знаков растяжение положительно, сжатие отрицательно."
    $row = Add-GuideTitle $Sheet $row "3. Условия равновесия" 13
    $row = Add-GuideParagraph $Sheet $row "Напряжения определяются по деформациям элементов и выбранным диаграммам материалов. Параметры плоскости деформаций уточняются до выполнения трех условий равновесия:"
    $row = Add-GuideFormula $Sheet $row "|N_int - N| ≤ T_N;    |Mx_int - Mx| ≤ T_Mx;    |My_int - My| ≤ T_My"
    $row = Add-GuideParagraph $Sheet $row "N_int, Mx_int и My_int - внутренние усилия. T_N, T_Mx и T_My задаются настройками Solver.ToleranceN, Solver.ToleranceMx и Solver.ToleranceMy на Config."
    $row = Add-GuideParagraph $Sheet $row "Допуски вводятся в INPUT Force и INPUT Moment - единицах ввода силы и момента. Они определяют точность равновесия, а не число знаков в таблице. Малые моменты зануляет другая настройка Config, Calculation.ZeroMomentPerDepth; 0 отключает такое зануление."
    $row = Add-GuideTitle $Sheet $row "4. Численный метод и параметры итераций" 13
    $row = Add-GuideParagraph $Sheet $row "В настройке Solver.Method выберите Newton или Secant. Newton использует касательную жесткость сечения, а Secant уточняет приближенную матрицу чувствительности по предыдущим состояниям." -RowCount 1
    $row = Add-GuideParagraph $Sheet $row "Solver.MaxIterations ограничивает число итераций на ступени нагрузки. Solver.LoadSteps задает число ступеней; их увеличение может повысить устойчивость сходимости ценой дополнительного времени."
    $row = Add-GuideParagraph $Sheet $row "General.PlateauSlopeFactor задаёт коэффициент α для наклона ненулевых плато бетона и арматуры: Et = α·E своей ветви. По умолчанию α = 0,001, наклон плато — 0,001·E (0,1% от E). 0 сохраняет горизонтальное плато. Наклон помогает сходимости и меняет напряжения, но не предельные деформации; выключенный растянутый бетон не включается. После изменения коэффициента выполните расчёт заново."
    $row = Add-GuideParagraph $Sheet $row "Решатель автоматически пробует полный шаг и уменьшает его, если равновесие ухудшается. Расчет ведется относительно центра бетонной сетки; исходные координаты геометрии и точка, к которой относится ε_0 в Results, не меняются."
    $row = Add-GuideParagraph $Sheet $row "На Config задаются метод, допуски равновесия, число ступеней нагрузки и предельное число итераций. Для Secant настройка Solver.SecantMaxRestarts ограничивает число перестроений приближенной матрицы. Малый шаг сам по себе не означает найденное равновесие."
    $row = Add-GuideTitle $Sheet $row "5. Продолжение диаграмм и статус" 13
    $row = Add-GuideParagraph $Sheet $row "General.DiagramExtension = Yes линейно продолжает активные ветви с наклоном 0,01*E без конечной границы деформаций. Физические сопротивления и предельные деформации материалов сохраняются." -Emphasis
    $row = Add-GuideParagraph $Sheet $row "Если найденное равновесие внутри физических пределов, статус - OK. Если оно найдено только за пределами диаграммы, статус - FAIL. NumFail означает, что численный поиск запускался, но равновесие не найдено; это не доказательство разрушения. Выключенный растянутый бетон остается выключенным и при продолжении диаграмм."
    $row = Add-GuideTitle $Sheet $row "6. Результаты и диагностика" 13
    $row = Add-GuideParagraph $Sheet $row "В подробном блоке прочности показаны напряжения, деформации бетона и стали, h - высота сечения по направлению изгиба, x - высота сжатой зоны, ExtUsed - фактическое использование продолженной диаграммы. Геометрические свойства сечения находятся в отдельной таблице."
    $row = Add-GuideParagraph $Sheet $row "εbc/εbt и εsc/εst - крайние деформации сжатия/растяжения бетона и арматуры. Результаты выводятся в OUTPUT; напряжения и деформации сохраняют внутреннюю систему знаков."
    $row = Add-GuideParagraph $Sheet $row "Для подробной диагностики включите на Config General.ExecutionReportEnabled = Yes. Диагностика обновляется при каждом запуске."
    return $row + 2
}

# Формирует методический блок по поиску несущей способности по выбранной
# λ-траектории. Описывает общую постановку, UltimateStrain,
# LoadMultiplier и автоматический выбор фактического метода.
function Add-CapacityMethodologyGuide {
    param([object]$Sheet, [int]$StartRow)

    $row = Add-GuideParagraph $Sheet $StartRow "Несущая способность определяется достижением предельной деформации бетона или арматуры на выбранной траектории нагружения. Множитель λ = 1 соответствует заданному сочетанию."
    $row = Add-GuideTitle $Sheet $row "1. Расчетный профиль и путь нагружения" 13
    $row = Add-GuideParagraph $Sheet $row "Расчет включается в профиле настройкой Calculation.Strength.Capacity = Yes. Используется модель прочности MaterialModel.Strength.*, общая с прямым расчетом НДС."
    $row = Add-GuideParagraph $Sheet $row "LoadPath в строке сочетания задает масштабируемые составляющие: Mx, My, оба момента, N или все усилия. Auto выбирает путь по ненулевым составляющим; пустое значение принимается как Auto."
    $row = Add-GuideParagraph $Sheet $row "При N + M режим Auto последовательно пробует моментный путь, λ*N и λ*NMxy до получения предельной точки. Найденный предел с λ < 1 завершает поиск. Явно заданный путь не заменяется другим."
    $row = Add-GuideTitle $Sheet $row "2. Постоянная и увеличиваемая части" 13
    $row = Add-GuideFormula $Sheet $row "N(λ) = N_0 + λ·N_base"
    $row = Add-GuideParagraph $Sheet $row "N_0 - постоянная часть силы; N_base - масштабируемая составляющая."
    $row = Add-GuideFormula $Sheet $row "M_x(λ) = M_{x,0} + λ·M_{x,base};    M_y(λ) = M_{y,0} + λ·M_{y,base}"
    $row = Add-GuideParagraph $Sheet $row "Индексы 0 и base у моментов также обозначают постоянную и масштабируемую части. При λ = 1 восстанавливается исходное сочетание."
    $row = Add-GuideParagraph $Sheet $row "Если сила приложена со смещением, заданным на Config настройками Load.ReferenceOffsetX и Load.ReferenceOffsetY, ее добавочный момент увеличивается вместе с N. Пользовательский момент масштабируется только тогда, когда выбранный путь включает этот момент."
    $row = Add-GuideParagraph $Sheet $row "Например, задано N = 100 тс и Mx = 50 тс·м; смещение силы добавляет еще 10 тс·м. При пути λ*N получаем N = λ·100 и Mx = 50 + λ·10. При λ*Mxy сила и ее добавка остаются постоянными: N = 100, Mx = 10 + λ·50."
    $row = Add-GuideTitle $Sheet $row "3. Поиск предельного состояния" 13
    $row = Add-GuideParagraph $Sheet $row "В настройке Capacity.SolutionStrategy выберите способ поиска. UltimateStrain совместно решает условия равновесия и предельной деформации, а LoadMultiplier уточняет λ последовательными расчетами равновесия. Auto начинает с UltimateStrain и при численной неудаче повторяет тот же путь через LoadMultiplier." -RowCount 2
    $row = Add-GuideParagraph $Sheet $row "Длина численного шага выбирается автоматически. Если обычный шаг UltimateStrain не улучшает решение, программа пробует резервный шаг. Успех подтверждается только равновесием и достижением физического предела с заданными допусками; сам по себе малый шаг не означает успешный расчет."
    $row = Add-GuideParagraph $Sheet $row "Для силового осевого пути λ*N и для λ*NMxy без пользовательских моментов применяется LoadMultiplier. На горизонтальном участке диаграммы прямой поиск предельной деформации может не дать единственного устойчивого решения."
    $row = Add-GuideParagraph $Sheet $row "Контролируются сжатие и растяжение бетона и стали. Растяжение бетона проверяется только при включенной растянутой ветви. Пределы берутся из диаграмм выбранной модели, а не из численных допусков." -Emphasis
    $row = Add-GuideCompositeFormula $Sheet $row @(
        @{ Kind = "Text"; Text = "max(" },
        @{ Kind = "Fraction"; Numerator = "|ε_b|"; Denominator = "ε_{b,ult}" },
        @{ Kind = "Text"; Text = "; " },
        @{ Kind = "Fraction"; Numerator = "|ε_s|"; Denominator = "ε_{s,ult}" },
        @{ Kind = "Text"; Text = ") = 1" }
    ) "" "max(|ε_b|/ε_{b,ult}; |ε_s|/ε_{s,ult}) = 1"
    $row = Add-GuideParagraph $Sheet $row "ε_b,ult и ε_s,ult - модули соответствующих предельных деформаций сжатия или растяжения. Максимум определяется по всем активным пределам бетона и арматуры."
    $row = Add-GuideParagraph $Sheet $row "Точность UltimateStrain задает настройка Config Capacity.ToleranceStrain. Число итераций обоих способов ограничивает Capacity.SolverMaxIterations. Допуски равновесия задаются на Config настройками Solver.ToleranceN, Solver.ToleranceMx и Solver.ToleranceMy."
    $row = Add-GuideTitle $Sheet $row "4. Настройки пробного поиска" 13
    $row = Add-GuideParagraph $Sheet $row "Для LoadMultiplier на Config настройка Capacity.InitialLambda задает первую положительную пробу, Capacity.MaxLambda - техническую границу расширения поиска, Capacity.ToleranceLambda - точность λ. MaxLambda не увеличивает физическую несущую способность и не ограничивает успешный прямой поиск UltimateStrain."
    $row = Add-GuideParagraph $Sheet $row "Настройка Config Capacity.SearchMethod выбирает уточнение интервала: Bisection делит его пополам; Brent сочетает деление с ускоренными пробами; Secant использует изменение результата между предыдущими пробами. Методы ищут один и тот же физический предел."
    $row = Add-GuideParagraph $Sheet $row "На Config Capacity.BaseLoadSteps задает число ступеней приложения нагрузки одной пробы, Capacity.MaxRetries - дополнительные попытки при численной несходимости. На каждом повторе число ступеней удваивается; численный шаг выбирается автоматически. Значения λ и усилий этой пробы не меняются."
    $row = Add-GuideParagraph $Sheet $row "Настройка Config General.DiagramExtension может помогать найти равновесие пробной нагрузки за пределами диаграммы. Физический критерий по деформациям остается исходным: численное продолжение не прибавляет сечению несущую способность."
    $row = Add-GuideTitle $Sheet $row "5. Результаты поиска" 13
    $row = Add-GuideFormula $Sheet $row "SF_capacity = λ_ult"
    $row = Add-GuideParagraph $Sheet $row "SF_capacity - коэффициент запаса по несущей способности; λ_ult - найденный предельный множитель. При λ_ult ≥ 1 заданная нагрузка проходит проверку, при λ_ult < 1 - не проходит."
    $row = Add-GuideParagraph $Sheet $row "BaseFail означает подтвержденный физический отказ постоянной части при λ = 0. NumFail означает, что предельную точку не удалось надежно найти численно; сам по себе он не доказывает разрушение. Комментарий объясняет причину и использованный путь."
    $row = Add-GuideParagraph $Sheet $row "В Results также показаны N_ult, Mx_ult и My_ult - усилия в найденной предельной точке, фактический метод и LoadPath. Усилия возвращаются к пользовательской точке приложения и выводятся в OUTPUT и пользовательских знаках; напряжения и деформации сохраняют физические знаки."
    return $row + 2
}

# Формирует верхний методический раздел листа "Справка" по фактическому
# алгоритму расчета нормальных трещин. Подробные настройки ниже остаются
# отдельными блоками, а этот раздел читается как общий user guide.
function Add-CrackWidthMethodologyGuide {
    param([object]$Sheet, [int]$StartRow)

    $row = Add-GuideParagraph $Sheet $StartRow "SLS.Crack.Code на Config выбирает методику расчета ширины нормальных трещин: SP63 или SP35. Расчетные зависимости и особенности применения приведены в отдельных подразделах."
    $row = Add-GuideParagraph $Sheet $row "Обе методики используют текущее НДС и границы бетона для площади взаимодействия. Точные контуры уточняют расчет трещин, но не меняют характеристики бетонной сетки для устойчивости. Методика устойчивости выбирается независимо настройкой Config Stability.Code."
    $row = Add-GuideTitle $Sheet $row "СП 63: расчет ширины нормальных трещин" 15
    $row = Add-SP63CrackWidthMethodologyGuide $Sheet $row
    return (Add-SP35CrackWidthMethodologyGuide $Sheet $row)
}

# Сохраняет последовательное объяснение существующей методики СП 63;
# меняется только общее вычисление геометрии, но не формула и ее коэффициенты.
function Add-SP63CrackWidthMethodologyGuide {
    param([object]$Sheet, [int]$StartRow)

    $row = Add-GuideParagraph $Sheet $StartRow "Расчет применяется при SLS.Crack.Code = SP63 и Calculation.Crack.Width = Yes в назначенном профиле. Реализовано продолжительное раскрытие для сочетаний II группы; расчет непродолжительного раскрытия по нескольким состояниям не предусмотрен."
    $row = Add-GuideTitle $Sheet $row "1. Условие проверки и общая формула" 13
    $row = Add-GuideFormula $Sheet $row "a_crc ≤ a_crc,ult      СП 63, п. 8.2.6, ф. (8.118)"
    $row = Add-GuideParagraph $Sheet $row "a_crc,ult задается настройкой SLS.Crack.Allowable в INPUT Length. Расчетная ширина a_crc выводится в OUTPUT Length."
    $row = Add-GuideFractionFormula $Sheet $row "a_crc = φ_1·φ_2·φ_3·ψ_s ·" "σ_s" "E_s" "· l_s      СП 63, п. 8.2.15, ф. (8.128)"
    $row = Add-GuideParagraph $Sheet $row "φ_1, φ_2 и φ_3 учитывают длительность нагрузки, поверхность арматуры и характер работы элемента; ψ_s - неравномерность деформаций арматуры между трещинами. σ_s - напряжение выбранной растянутой арматуры; E_s - ее модуль упругости; l_s - расчетное расстояние между трещинами."

    $row = Add-GuideTitle $Sheet $row "2. Трещинообразование и применимость расчета" 13
    $row = Add-GuideParagraph $Sheet $row "Проверка трещинообразования предшествует расчету ширины. Текущее НДС для раскрытия определяется при заданном сочетании без учета растянутого бетона."
    $row = Add-GuideParagraph $Sheet $row "Путь поиска задается столбцом LoadPath в таблице сочетаний на Config; пустое значение означает Auto. Численный способ выбирается настройкой Config SLS.Crack.InitiationSolutionStrategy. Допуск уточнения λ методом LoadMultiplier задается отдельно в SLS.Crack.ToleranceLambda. Отдельного пути только для трещин нет."
    $row = Add-GuideFormula $Sheet $row "λ_crc ≤ 1: трещина образуется;    λ_crc > 1: трещина еще не образуется"
    $row = Add-GuideParagraph $Sheet $row "λ_crc - найденный множитель выбранной части нагрузки в момент образования трещины. λ = 1 соответствует текущему сочетанию. Нормальная трещина может быть также достоверно исключена без найденной пороговой точки; тогда выводится NotCracked."
    $row = Add-GuideParagraph $Sheet $row "Подтвержденный NotCracked исключает расчет ширины: статус проверки N/A, численная ячейка a_crc остается пустой. Это не расчетная ширина, равная нулю. Проверка продольных трещин при этом сохраняется." -Emphasis
    $row = Add-GuideParagraph $Sheet $row "Отказ более высокой пробной нагрузки не означает отказа текущего сочетания. Если допустимое равновесие при λ = 1 не достигает критерия трещины, а более высокая проба достигает другого физического предела, текущее сочетание остается NotCracked."

    $row = Add-GuideTitle $Sheet $row "3. Центральное растяжение" 13
    $row = Add-GuideParagraph $Sheet $row "Центральное растяжение определяется относительно приведенного центра сечения бетон + арматура. При смещении точки приложения, заданном на Config настройками Load.ReferenceOffsetX и Load.ReferenceOffsetY, даже одна сила N может создавать изгиб и требовать общей проверки."
    $row = Add-GuideFormula $Sheet $row "N > N_crc      СП 63, п. 8.2.4, ф. (8.117)"
    $row = Add-GuideParagraph $Sheet $row "N - растягивающая сила; N_crc - сила образования нормальных трещин при центральном растяжении."
    $row = Add-GuideFormula $Sheet $row "N_crc = A_red · R_bt,ser      СП 63, п. 8.2.13, ф. (8.127)"
    $row = Add-GuideParagraph $Sheet $row "A_red - приведенная площадь для проверки центрального растяжения; R_bt,ser задается в колонке растяжения строки Concrete.R.SLS(II)."
    $row = Add-GuideCompositeFormula $Sheet $row @(
        @{ Kind = "Text"; Text = "A_red = A_b + A_s,total · " },
        @{ Kind = "Fraction"; Numerator = "E_s"; Denominator = "E_b" }
    ) "" "A_red = A_b + A_s,total · E_s/E_b"
    $row = Add-GuideParagraph $Sheet $row "A_b и A_s,total - полные площади бетона и продольной арматуры. E_b и E_s берутся из выбранной модели по строкам Concrete.E и Steel.E."
    $row = Add-GuideFractionFormula $Sheet $row "λ_crc =" "N_crc" "N" ""
    $row = Add-GuideParagraph $Sheet $row "Для центрального растяжения множитель равен отношению пороговой силы к текущей. При расчете ширины берется вся площадь бетона за вычетом отверстий и вся фактически растянутая арматура."

    $row = Add-GuideTitle $Sheet $row "4. Образование трещины при изгибе и внецентренной силе" 13
    $row = Add-GuideParagraph $Sheet $row "В общем случае точка трещинообразования определяется достижением предельной растягивающей деформации бетона. При одновременном сжатии и растяжении сечения:"
    $row = Add-GuideFormula $Sheet $row "ε_{bt,ult} = ε_{bt2}      СП 63, п. 8.2.14 и 8.1.30"
    $row = Add-GuideParagraph $Sheet $row "ε_bt,ult - критерий трещинообразования; ε_bt2 - предельная деформация растянутой ветви диаграммы бетона."
    $row = Add-GuideParagraph $Sheet $row "Если все сечение растянуто, но кривизна ненулевая, применяется формула:"
    $row = Add-GuideCompositeFormula $Sheet $row @(
        @{ Kind = "Text"; Text = "ε_{bt,ult} = ε_{bt2} - (ε_{bt2} - ε_{bt0}) · " },
        @{ Kind = "Fraction"; Numerator = "ε_1"; Denominator = "ε_2" }
    ) "СП 63, п. 8.1.30, ф. (8.54)" "ε_{bt,ult} = ε_{bt2} - (ε_{bt2} - ε_{bt0}) · ε_1/ε_2"
    $row = Add-GuideParagraph $Sheet $row "ε_bt0 - деформация начала предельного участка растянутой диаграммы; ε_1 и ε_2 - растягивающие деформации противоположных граней. Используется модель MaterialModel.CrackInitiation.* назначенного профиля."
    $row = Add-GuideParagraph $Sheet $row "Настройка Config General.DiagramExtension может помочь численному поиску равновесия, но не увеличивает физический порог образования. На Results сохраняются сила и момент найденной точки, использованный путь и причина неудачи, если точка не найдена."

    $row = Add-GuideTitle $Sheet $row "5. Текущее НДС и площадь взаимодействия" 13
    $row = Add-GuideParagraph $Sheet $row "Для расчета ширины используются напряжения текущего сочетания. В профиле задаются MaterialModel.CrackedState.ValueSet = SLS(II) и MaterialModel.CrackedState.ConcreteTension = Ignore."
    $row = Add-GuideParagraph $Sheet $row "Настройка Config SLS.Crack.TensionZoneMode выбирает бетонную площадь A_bt: FullTension - всю растянутую часть; Effective - эффективную полосу с ограничениями. A_bt означает площадь бетона расчетной растянутой зоны."
    $row = Add-GuideFormula $Sheet $row "h_bt = min(max(x_t; 2a_s); 0.5·h)"
    $row = Add-GuideParagraph $Sheet $row "h_bt - глубина эффективной полосы; x_t - глубина растянутой части до НЛ; a_s - расстояние от оси крайнего выбранного стержня до растянутой границы бетона; h - габарит сечения по нормали к НЛ."
    $row = Add-GuideParagraph $Sheet $row "SLS.Crack.CoverDistanceMode задает способ определения a_s: NearestContour - до ближайшей границы бетона по нормали к НЛ; GlobalExtreme - до дальней растянутой опорной линии всего сечения."
    $row = Add-GuideParagraph $Sheet $row "Если x_t меньше 2a_s, Effective расширяет бетонную полосу до 2a_s, но не глубже 0,5h. Она может зайти за НЛ. Сжатые стержни из этой полосы все равно не входят в расчет раскрытия: выбирается только растянутая арматура."
    $row = Add-GuideParagraph $Sheet $row "Площадь считается по пересечению полосы с бетоном и выбранными отверстиями. Правила четырех случаев задания контуров приведены в конце раздела подготовки AutoCAD. Без точного наружного контура граница по сетке приближенная; реальная площадь и инерции элементов для НДС не заменяются."

    $row = Add-GuideTitle $Sheet $row "6. Арматура, напряжение и эквивалентный диаметр" 13
    $row = Add-GuideParagraph $Sheet $row "A_s - сумма площадей растянутых стержней, оси которых попали в расчетную зону. Напряжение этих стержней усредняется по площади:"
    $row = Add-GuideFractionFormula $Sheet $row "σ_s =" "Σ(A_s,i · σ_s,i)" "ΣA_s,i" ""
    $row = Add-GuideParagraph $Sheet $row "Суммирование выполняется по выбранным растянутым стержням; A_s,i и σ_s,i - площадь и текущее напряжение i-го стержня. Усреднение по площади принято в программе как обобщение НДМ для произвольного сечения; отдельной формулой СП 63 оно не является."
    $row = Add-GuideFractionFormula $Sheet $row "d_s,eq =" "Σ(n_i·d_i^2)" "Σ(n_i·d_i)" ""
    $row = Add-GuideParagraph $Sheet $row "d_s,eq - эквивалентный диаметр; n_i - число выбранных стержней диаметра d_i. Для смешанных диаметров применено инженерное дополнение по EN 1992-1-1; общий случай отдельно не раскрыт в СП 63."

    $row = Add-GuideTitle $Sheet $row "7. Расстояние между трещинами" 13
    $row = Add-GuideFractionFormula $Sheet $row "l_{s,0} = 0.5 ·" "A_bt" "A_s" "· d_s,eq      СП 63, п. 8.2.17, ф. (8.136)"
    $row = Add-GuideParagraph $Sheet $row "l_s,0 - исходное расстояние между трещинами, до применения нормативных ограничений. A_bt, A_s и d_s,eq определяются по выбранной расчетной зоне."
    $row = Add-GuideFormula $Sheet $row "l_s = max(l_{s,0}; 10·d_s,eq; 100 мм)"
    $row = Add-GuideParagraph $Sheet $row "Сначала применяется нижняя граница: принятое расстояние l_s не меньше 10 диаметров и 100 мм."
    $row = Add-GuideFormula $Sheet $row "l_s = min(l_s; 40·d_s,eq; 400 мм)      СП 63, п. 8.2.17"
    $row = Add-GuideParagraph $Sheet $row "Затем применяется верхняя граница: не больше 40 диаметров и 400 мм. Если d_s,eq больше 40 мм, нижнее и верхнее ограничения противоречат друг другу. Программа использует верхнюю границу 400 мм и явно предупреждает об этом в комментарии выполненного расчета."

    $row = Add-GuideTitle $Sheet $row "8. Коэффициенты, задаваемые на Config" 13
    $row = Add-GuideParagraph $Sheet $row "SLS.Crack.Phi1 задает φ_1: для продолжительного раскрытия 1,4. SLS.Crack.Phi2 задает φ_2: 0,5 для периодического профиля и 0,8 для гладкой арматуры. Применимость коэффициентов определяется расчетной ситуацией."
    $row = Add-GuideParagraph $Sheet $row "Настройка Config SLS.Crack.Phi3Mode = Auto принимает φ_3 = 1,2 при растягивающей продольной силе и 1,0 в остальных реализованных случаях. При User используется число из настройки Config SLS.Crack.Phi3."
    $row = Add-GuideParagraph $Sheet $row "Настройка SLS.Crack.PsiMode на листе Config выбирает правило определения ψ_s. User использует число из SLS.Crack.PsiS, если получены пригодные данные трещинообразования. Auto сначала принимает 1 и уточняет коэффициент только при превышении допуска ширины. AlwaysCalc выполняет уточнение всегда, когда оно возможно." -RowCount 2
    $row = Add-GuideFractionFormula $Sheet $row "ψ_s = 1 - 0.8 ·" "σ_s,crc" "σ_s" "      СП 63, п. 8.2.18, ф. (8.137)"
    $row = Add-GuideParagraph $Sheet $row "σ_s,crc - среднее по площади напряжение тех же выбранных стержней сразу после образования трещины; σ_s - их среднее при текущем сочетании. Индекс crc обозначает пороговое состояние. Набор стержней не меняется между этими двумя состояниями. Значение ψ_s ограничивается диапазоном 0...1."
    $row = Add-GuideParagraph $Sheet $row "Настройка Config SLS.Crack.SigmaSCrcAveragingMode задает учет стержней, ставших сжатыми в пороговом состоянии. AllSelected учитывает их напряжения со знаком; TensionOnly принимает их напряжения за 0, но сохраняет площади в знаменателе среднего."
    $row = Add-GuideParagraph $Sheet $row "Если пригодной точки образования нет либо выбранная арматура в ней не растянута, но растянута при текущем сочетании, ширина продолжает считаться с ψ_s = 1 во всех режимах. Предупреждение выводится только если это значение действительно использовано в формуле." -Emphasis
    $row = Add-GuideParagraph $Sheet $row "При подтвержденном NotCracked формула ширины не вызывается, пользовательское SLS.Crack.PsiS не меняется и предупреждения о ψ_s = 1 нет. Отсутствие или физический отказ текущего НДС блокирует обе проверки трещин; ошибочные обязательные настройки требуют исправления."

    $row = Add-GuideTitle $Sheet $row "9. Ширина, запас и результат" 13
    $row = Add-GuideFractionFormula $Sheet $row "SF_crc =" "a_crc,ult" "a_crc" ""
    $row = Add-GuideParagraph $Sheet $row "SF_crc - коэффициент запаса по ширине. Если он меньше 1, допуск превышен и статус - FAIL. Для неприменимой или пропущенной проверки ширина и запас не выводятся. Рассчитанное a_crc = 0 - другой случай, не подтверждение отсутствия трещины."
    $row = Add-GuideParagraph $Sheet $row "Подробный блок Results содержит площадь, состав арматуры, напряжения, коэффициенты и ширину. В сводке выводятся общий статус и минимальный запас; в ResultComment - причины пропуска и предупреждения."

    $row = Add-GuideTitle $Sheet $row "10. Независимая проверка продольных трещин" 13
    $row = Add-GuideParagraph $Sheet $row "Она контролирует максимальное сжимающее напряжение в бетоне текущего НДС и выполняется как часть Calculation.Crack.Width. Отдельный поиск равновесия для этой проверки не нужен."
    $row = Add-GuideFormula $Sheet $row "σ_c,max ≤ R_b,mc2"
    $row = Add-GuideParagraph $Sheet $row "σ_c,max - максимальное по модулю сжимающее напряжение бетона текущего НДС; R_b,mc2 задается настройкой Concrete.Rb.mc2."
    $row = Add-GuideFractionFormula $Sheet $row "SF_long =" "R_b,mc2" "σ_c,max" ""
    $row = Add-GuideParagraph $Sheet $row "SF_long - запас по продольным трещинам. При отсутствии сжатого бетона проверка неприменима и получает N/A независимо от результата расчета нормальных трещин."
    $row = Add-GuideParagraph $Sheet $row "По СП 35 п. 7.100 и таблицам 7.6, 7.7 пользователь задает итоговое допустимое сопротивление с учетом применимых коэффициентов условий работы, включая m_b13. m_b13 - коэффициент условий работы бетона для этой проверки. Программа эти коэффициенты повторно не умножает."
    return $row + 2
}

# Рисует поясняющие схемы общей методики, не привязанные к типу сечения.
# Это учебная иллюстрация правил, а не альтернативный расчет области A_r.
# Картинка вставляется тем же механизмом, что формулы, и скрывается с группой.
function New-SP35GuideScheme {
    param([ValidateSet('Groups', 'Rows', 'Contours', 'MeshCentral')][string]$Kind)

    if (-not (Initialize-FormulaImageRenderer)) { return $null }
    $bitmap = New-Object System.Drawing.Bitmap(1100, 340)
    $g = [System.Drawing.Graphics]::FromImage($bitmap)
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit
    $g.Clear([System.Drawing.Color]::White)
    $pen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(45, 55, 60), 2)
    $guide = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(90, 105, 110), 2)
    $guide.DashStyle = [System.Drawing.Drawing2D.DashStyle]::Dash
    $boundary = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(15, 125, 140), 3)
    $opening = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(175, 45, 65), 2)
    $fill = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(218, 240, 229))
    $bar = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(45, 55, 60))
    $text = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(25, 35, 40))
    $font = New-Object System.Drawing.Font('Arial', 19)
    try {
        switch ($Kind) {
            'Groups' {
                for ($row = 0; $row -lt 4; $row++) {
                    $y = 55 + 70 * $row
                    for ($i = 0; $i -le $row; $i++) { $g.FillEllipse($bar, 60 + 47 * $i, $y, 36, 36) }
                    $g.DrawString(@('1 стержень: β=1.00', '2 стержня: β=0.85', '3 стержня: β=0.75', '>3 стержней: β=0.75 [NDM]')[$row], $font, $text, 300, $y + 3)
                }
                $g.FillEllipse($bar, 655, 80, 24, 24)
                $g.FillEllipse($bar, 755, 68, 48, 48)
                $g.DrawLine($guide, 680, 92, 752, 92)
                $g.DrawString('зазор между стержнями', $font, $text, 600, 35)
                $g.DrawLine($opening, 759, 110, 759, 145)
                $g.DrawString('центр группы', $font, $text, 815, 120)
                $g.DrawString('центр ближе', $font, $text, 650, 175)
                $g.DrawString('к большому стержню', $font, $text, 650, 205)
            }
            'Rows' {
                $g.DrawRectangle($pen, 70, 45, 890, 210)
                $g.FillRectangle($fill, 380, 45, 240, 180)
                $g.DrawLine($guide, 70, 265, 960, 265)
                $g.DrawString('НЛ', $font, $text, 975, 250)
                foreach ($x in @(260, 500, 740)) { $g.FillEllipse($bar, $x - 12, 70, 24, 24) }
                $g.FillEllipse($bar, 488, 130, 24, 24)
                $g.FillEllipse($bar, 495, 195, 10, 10)
                $g.DrawLine($boundary, 380, 45, 380, 225)
                $g.DrawLine($boundary, 620, 45, 620, 225)
                $g.DrawLine($boundary, 380, 225, 620, 225)
                $g.DrawString('t/L < q: r;  t/L ≥ q: min(r; t/2)', $font, $text, 280, 8)
                $g.DrawString('1 ряд: наружный', $font, $text, 780, 70)
                $g.DrawString('2 ряд: полный', $font, $text, 665, 128)
                $g.DrawString('3 ряд: малый', $font, $text, 665, 193)
                $g.DrawLine($guide, 365, 142, 365, 225)
                $g.DrawString('r от полного ряда', $font, $text, 110, 185)
                $g.DrawString('Ряды 1 → 2 → 3: от поверхности бетона к нейтральной линии (НЛ)', $font, $text, 80, 295)
            }
            'Contours' {
                $g.DrawRectangle($pen, 75, 50, 950, 220)
                $g.FillRectangle($fill, 260, 50, 560, 180)
                $g.FillRectangle([System.Drawing.Brushes]::White, 480, 115, 100, 70)
                $g.DrawRectangle($opening, 480, 115, 100, 70)
                $g.FillRectangle([System.Drawing.Brushes]::White, 710, 100, 25, 20)
                $g.DrawRectangle($opening, 710, 100, 25, 20)
                for ($i = 0; $i -lt 5; $i++) {
                    $x = 280 + 130 * $i; $end = 230
                    if ($i -eq 2) { $end = 115 }
                    $g.DrawLine($guide, $x, 50, $x, $end)
                    $g.DrawString('P' + ($i + 1), $font, $text, $x - 15, 20)
                    $g.DrawString('Q' + ($i + 1), $font, $text, $x - 15, $end + 3)
                }
                $g.DrawString('отверстие', $font, $text, 480, 140)
                $g.DrawString('местный вырез', $font, $text, 845, 140)
                $g.DrawString('без сужения ниже', $font, $text, 825, 170)
                $g.DrawString('Пустота между пятью пробами тоже отсекается точно', $font, $text, 210, 295)
            }
            'MeshCentral' {
                for ($i = 0; $i -lt 8; $i++) { for ($j = 0; $j -lt 4; $j++) {
                    if ($i -in 3, 4 -and $j -in 1, 2) { continue }
                    $g.DrawRectangle($guide, 55 + 45 * $i, 55 + 40 * $j, 45, 40)
                } }
                $g.DrawRectangle($boundary, 55, 55, 360, 160)
                $g.DrawRectangle($opening, 190, 95, 90, 80)
                $g.DrawEllipse($opening, 320, 145, 25, 25)
                $g.DrawString('без наружного контура:', $font, $text, 60, 245)
                $g.DrawString('наружная граница по сетке', $font, $text, 30, 275)
                $g.FillRectangle($fill, 590, 55, 420, 180)
                $g.DrawRectangle($boundary, 590, 55, 420, 180)
                $g.FillEllipse([System.Drawing.Brushes]::White, 700, 110, 70, 70)
                $g.DrawEllipse($opening, 700, 110, 70, 70)
                $g.FillRectangle([System.Drawing.Brushes]::White, 850, 105, 80, 65)
                $g.DrawRectangle($opening, 850, 105, 80, 65)
                foreach ($x in @(615, 975)) { foreach ($y in @(80, 200)) { $g.FillEllipse($bar, $x, $y, 12, 12) } }
                $g.DrawString('центральное растяжение:', $font, $text, 595, 245)
                $g.DrawString('вся площадь бетона без отверстий', $font, $text, 575, 275)
            }
        }
        $path = New-FormulaImagePath ('sp35_' + $Kind)
        $bitmap.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)
        return @{ Path = $path; WidthPoints = 660; HeightPoints = 204 }
    }
    finally {
        $g.Dispose(); $bitmap.Dispose(); $pen.Dispose(); $guide.Dispose(); $boundary.Dispose()
        $opening.Dispose(); $fill.Dispose(); $bar.Dispose(); $text.Dispose(); $font.Dispose()
    }
}

# Объясняет СП 35 от общей формулы к подготовленным группам и области.
# Нормативные правила отделены от программных допусков и геометрического
# обобщения NDM; все формулы используют уже найденное текущее НДС.
function Add-SP35CrackWidthMethodologyGuide {
    param([object]$Sheet, [int]$StartRow)

    $row = Add-GuideTitle $Sheet $StartRow "СП 35: расчет ширины нормальных трещин" 15
    $row = Add-GuideParagraph $Sheet $row "Ширина нормальных трещин определяется по текущему НДС и площади взаимодействия бетона с растянутой арматурой. Для каждой растянутой группы выполняется отдельная проверка; определяющим является наибольшее раскрытие."
    $row = Add-GuideParagraph $Sheet $row "Источники положений: [SP35] - требования СП 35; [METHOD] - Методическое пособие 2017 года к СП 35; [NDM] - инженерные допущения реализации. Пояснения пособия и допущения программы не заменяют нормативные требования."

    $row = Add-GuideTitle $Sheet $row "1. Исходные данные и применимость" 13
    $row = Add-GuideParagraph $Sheet $row "На Config задаются SLS.Crack.Code = SP35 и Calculation.Crack.Width = Yes в профиле, назначенном сочетанию."
    $row = Add-GuideParagraph $Sheet $row "Допустимая ширина задается настройкой SLS.Crack.Allowable в INPUT Length по требованиям проекта. Автоматическое назначение допуска не предусмотрено."
    $row = Add-GuideParagraph $Sheet $row "Проверка трещинообразования использует общий LoadPath сочетания; пустое значение принимается как Auto. Способ поиска задается SLS.Crack.InitiationSolutionStrategy, допуск λ для LoadMultiplier - SLS.Crack.ToleranceLambda. Подробные правила приведены в справке соответствующих настроек."
    $row = Add-GuideParagraph $Sheet $row "При увеличении нагрузки проверяется первый достигнутый предел: образование трещины, сжатие бетона, растяжение или сжатие арматуры. Если другой предел наступает раньше трещины, он не становится точкой трещинообразования. При подтвержденном допустимом НДС исходного сочетания без трещины выводится NotCracked."
    $row = Add-GuideParagraph $Sheet $row "Подтвержденный NotCracked исключает расчет ширины: статус N/A, численная ячейка пустая. При Cracked раскрытие определяется по напряжениям текущего сочетания."
    $row = Add-GuideParagraph $Sheet $row "Неудача поиска трещинообразования не блокирует расчет ширины при найденном физически допустимом текущем НДС; причина сохраняется в комментарии. Отсутствие или физическая недопустимость текущего НДС блокирует и ширину, и продольные трещины." -Emphasis
    $row = Add-GuideParagraph $Sheet $row "Проверка продольных трещин выполняется отдельно по сжатому бетону и сохраняется даже при NotCracked. Ее допустимое напряжение задается на Config настройкой Concrete.Rb.mc2."

    $row = Add-GuideTitle $Sheet $row "2. Группирование растянутой арматуры" 13
    $row = Add-GuideParagraph $Sheet $row "Отбираются стержни с положительными напряжением и деформацией в текущем НДС. Пользовательская система знаков нагрузок на отбор не влияет."
    $row = Add-GuideParagraph $Sheet $row "Группа содержит один или несколько растянутых стержней, объединенных по расстоянию между поверхностями. Предельный зазор задается SLS.Crack.SP35.GroupGapTolerance в INPUT Length; значение новой книги - 10 мм."
    $row = Add-GuideParagraph $Sheet $row "Если стержни 1 и 2 связаны по зазору, а стержни 2 и 3 также связаны, все три образуют одну группу 1-2-3. Сжатые стержни в группировании не участвуют. Ненулевой допуск зазора — допущение [NDM], а не требование СП." -RowCount 1
    $row = Add-GuideFractionFormula $Sheet $row "s_{12} = L_{12} -" "d_1 + d_2" "2" ""
    $row = Add-GuideParagraph $Sheet $row "s_12 - зазор между поверхностями; L_12 - расстояние между центрами; d_1 и d_2 - диаметры стержней. Индексы 1 и 2 относятся к стержням, не к рядам."
    $row = Add-GuideCompositeFormula $Sheet $row @(
        @{ Kind = "Text"; Text = "L_{12} = " },
        @{ Kind = "Radical"; Radicand = "(x_2-x_1)^2 + (y_2-y_1)^2" }
    ) "" "L_12 = √((x_2-x_1)^2 + (y_2-y_1)^2)"
    $row = Add-GuideParagraph $Sheet $row "x_1, y_1 и x_2, y_2 - координаты центров рассматриваемой пары стержней."
    $row = Add-GuideFractionFormula $Sheet $row "x_ц =" "A_1·x_1 + A_2·x_2 + ..." "A_1 + A_2 + ..." "; аналогично y_ц"
    $row = Add-GuideParagraph $Sheet $row "Центр группы определяется по фактическим площадям A_1, A_2, ... ее стержней. x_ц и y_ц - его координаты; суммирование ведется по полному составу группы."
    $row = Add-GuideParagraph $Sheet $row "По СП 35: Коэффициент β учитывает число стержней в группе: 1 стержень - β = 1,00; 2 - 0,85; 3 - 0,75 (таблица 7.26). Для двух и трех стержней таблица описывает группы без просветов между ними."
    $row = Add-GuideParagraph $Sheet $row "Принято в программе: Для группы более трех обычных стержней программа использует β = 0,75. Это принятое обобщение, не отдельное правило СП для такой группы. Напрягаемая арматура и канатные пучки в этой версии не поддерживаются."
    $row = Add-GuideParagraph $Sheet $row "Геометрическая середина между крайними стержнями не используется вместо центра по площадям."
    $row = Add-GuideFormulaPicture $Sheet $row (New-SP35GuideScheme 'Groups') "Группы стержней: число стержней, коэффициент β и положение центра."

    $row = Add-GuideTitle $Sheet $row "3. Площадь взаимодействия и бетонная граница" 13
    $row = Add-GuideParagraph $Sheet $row "Для каждой растянутой группы строится отдельная площадь взаимодействия с окружающим бетоном. Все группы рассматриваются последовательно."
    $row = Add-GuideParagraph $Sheet $row "Границы задаются в направлениях вдоль нейтральной линии (НЛ) и по нормали к ней, независимо от поворота сечения."
    $row = Add-GuideParagraph $Sheet $row "Со стороны растянутой поверхности используется фактическая граница бетонной части, содержащей рассматриваемую группу. Вышележащие стержни не заменяют эту границу. Отверстия вычитаются из полученной площади."
    $row = Add-GuideParagraph $Sheet $row "После построения боковых границ и ограничения к НЛ определяется пересечение участка с бетонной областью. Отверстия образуют локальные вырезы без изменения боковых границ за пределами отверстий."

    $row = Add-GuideTitle $Sheet $row "4. Боковые границы" 13
    $row = Add-GuideFormula $Sheet $row "r_бок = k·d_бок"
    $row = Add-GuideParagraph $Sheet $row "r_бок - исходное расстояние для определения боковых границ площади взаимодействия от центра рассматриваемой группы; k - множитель 3, 5 или 6; d_бок - принятый диаметр стержней этой группы."
    $row = Add-GuideParagraph $Sheet $row "Множитель выбирается на Config настройкой SLS.Crack.SP35.InteractionRadiusMode: 3d, 5d или 6d. Диаметр выбирается там же настройкой SLS.Crack.SP35.RadiusDiameterMode: Max - наибольший, Min - наименьший, Average - среднее арифметическое диаметров. Эти два выбора используются и сбоку, и в сторону НЛ."
    $row = Add-GuideParagraph $Sheet $row "По СП 35: Пункт 7.110 устанавливает 6d в общем случае и 3d для непрямоугольного сечения с равномерной арматурой по контуру. 5d относится к пучкам и канатам, не поддерживаемым этой версией. Выбор режима требует проверки его нормативной применимости."
    $row = Add-GuideParagraph $Sheet $row "Слева и справа вдоль НЛ отдельно выбирается ближайшая растянутая группа арматурных стержней по полному расстоянию L между центрами, независимо от принадлежности к ряду."
    $row = Add-GuideCompositeFormula $Sheet $row @(
        @{ Kind = "Text"; Text = "L = " },
        @{ Kind = "Radical"; Radicand = "(x_2-x_1)^2 + (y_2-y_1)^2" }
    ) "" "L = √((x_2-x_1)^2 + (y_2-y_1)^2)"
    $row = Add-GuideParagraph $Sheet $row "x_1, y_1 и x_2, y_2 - центры рассматриваемой и соседней групп. L - полное расстояние между центрами; t - расстояние между их проекциями на НЛ."
    $row = Add-GuideParagraph $Sheet $row "Порог q = SLS.Crack.SP35.NeighborRatioLimit сравнивается с t/L отдельно для каждой стороны. Допустимый диапазон 0...1; значение по умолчанию 0,2."
    $row = Add-GuideCompositeFormula $Sheet $row @(
        @{ Kind = "Fraction"; Numerator = "t"; Denominator = "L" },
        @{ Kind = "Text"; Text = " < q: b = r_бок;    " },
        @{ Kind = "Fraction"; Numerator = "t"; Denominator = "L" },
        @{ Kind = "Text"; Text = " ≥ q: b = min(r_бок; " },
        @{ Kind = "Fraction"; Numerator = "t"; Denominator = "2" },
        @{ Kind = "Text"; Text = ")" }
    ) "" "t/L < q: b = r_бок; t/L ≥ q: b = min(r_бок; t/2)"
    $row = Add-GuideParagraph $Sheet $row "b - принятое расстояние от центра группы до боковой границы. При t/L < q сохраняется r_бок; при t/L ≥ q расстояние ограничивается половиной t."
    $row = Add-GuideParagraph $Sheet $row "Если соседа с этой стороны нет, остается полный r_бок. Группа с той же проекцией на НЛ не считается левым или правым соседом. Если два соседа одинаково близки, используется устойчивый порядок их номеров."
    $row = Add-GuideParagraph $Sheet $row "Боковые границы - две нормали к НЛ. Близость бетонной грани или отверстия не меняет их положения: расчетный участок обрезается по фактическому контуру бетона."

    $row = Add-GuideTitle $Sheet $row "5. Ряды и граница в сторону НЛ" 13
    $row = Add-GuideParagraph $Sheet $row "В боковом участке группы объединяются в ряды по координате вдоль нормали к НЛ. Максимальный разнос центров одного ряда задается SLS.Crack.SP35.RowTolerance в INPUT Length; значение по умолчанию 50 мм."
    $row = Add-GuideParagraph $Sheet $row "Ряд 1 - ближайший к растянутой бетонной границе; нумерация возрастает к НЛ. RowTolerance определяет только состав рядов и не изменяет боковые границы." -Emphasis
    $row = Add-GuideParagraph $Sheet $row "По СП 35: Расстояние r к НЛ откладывается от оси последнего растянутого ряда. Если площадь его арматуры строго меньше половины площади каждого другого ряда участка, отсчет переносится на предпоследний полный ряд (п. 7.110, абзац о крайнем ряде)."
    $row = Add-GuideParagraph $Sheet $row "При выборе опорного ряда сравниваются суммы фактических площадей стержней целых принятых групп, без коэффициента β. Условие строгое: равенство половине не требует переноса."
    $row = Add-GuideFormula $Sheet $row "r_НЛ = k·d_ряд"
    $row = Add-GuideParagraph $Sheet $row "r_НЛ - расстояние от опорного ряда к НЛ; d_ряд выбирается по всем его стержням согласно SLS.Crack.SP35.RadiusDiameterMode. Множитель k общий с боковыми границами."
    $row = Add-GuideParagraph $Sheet $row "Граница ограничивается НЛ. Опорный ряд определяется в боковом участке бетона; после построения конечной площади состав входящих групп повторно проверяется по их центрам."
    $row = Add-GuideFormulaPicture $Sheet $row (New-SP35GuideScheme 'Rows') "Ряды нумеруются от бетона к НЛ; r откладывается от выбранного опорного ряда."
    $row = Add-GuideParagraph $Sheet $row "Методическое пояснение: Примеры областей показаны на рисунке 4.42 Методического пособия 2017 года, раздел 4.2.3. Правило переноса отсчета на предпоследний ряд приведено в самом СП 35."

    $row = Add-GuideTitle $Sheet $row "6. Состав арматуры и расчетное напряжение" 13
    $row = Add-GuideParagraph $Sheet $row "Группа неделима: центр внутри конечной площади или на ее границе включает всю группу; центр снаружи исключает ее целиком, в том числе отдельные стержни внутри площади. Частичного включения нет." -Emphasis
    $row = Add-GuideParagraph $Sheet $row "Число стержней, площадь арматуры, сумма диаметров и β относятся к полному составу принятой группы. При этом каждый исходный стержень должен находиться в бетоне; расположение в отверстии или за наружной границей является ошибкой геометрии."
    $row = Add-GuideParagraph $Sheet $row "Вышележащие растянутые группы, вошедшие в площадь взаимодействия, участвуют в сумме диаметров. Расчетное напряжение принимается только по рассматриваемой группе: максимальное растягивающее напряжение ее стержней в текущем НДС."
    $row = Add-GuideParagraph $Sheet $row "Остальные группы проверяются в отдельных итерациях со своими площадями взаимодействия и напряжениями. Если центр рассматриваемой группы не входит в ее конечную площадь, пропускается только эта проверка."

    $row = Add-GuideTitle $Sheet $row "7. Площадь бетона и радиус армирования" 13
    $row = Add-GuideFormula $Sheet $row "A_r = A_{бет,1} + A_{бет,2} + ... - (A_{отв,1} + A_{отв,2} + ...)"
    $row = Add-GuideParagraph $Sheet $row "A_r - площадь бетона конечного участка взаимодействия; A_бет,1, A_бет,2, ... - площади его бетонных частей до вычитания отверстий; A_отв,1, A_отв,2, ... - площади отверстий в пределах участка."
    $row = Add-GuideParagraph $Sheet $row "Можно учитывать несколько отверстий и несколько бетонных частей. В полом сечении рассматривается стенка, где находится выбранная группа; противоположная стенка через пустоту в эту область не добавляется."
    $row = Add-GuideParagraph $Sheet $row "Если точные границы заданы, учитываются их прямые участки и круговые дуги. Иначе используются границы бетонной сетки. Четыре случая выбора границ перечислены в конце раздела «Подготовка геометрии в AutoCAD и импорт». Для формулы, схемы и экспорта используется одна и та же конечная область."
    $row = Add-GuideParagraph $Sheet $row "P1...P5 и Q1...Q5 в Results и на схеме - контрольные пересечения с противоположными границами по пяти направлениям. Площадь определяется по всему контуру, включая отверстия между контрольными направлениями."
    $row = Add-GuideFormulaPicture $Sheet $row (New-SP35GuideScheme 'Contours') "Отверстия образуют местные вырезы, включая отверстия между контрольными направлениями P/Q."
    $row = Add-GuideFractionFormula $Sheet $row "R_r =" "A_r" "Σβnd" "      СП 35, п. 7.110, ф. (7.100)"
    $row = Add-GuideParagraph $Sheet $row "R_r - радиус армирования, расчетная характеристика, не геометрический отступ. Знаменатель учитывает все принятые целые группы: β - коэффициент группы, n - число стержней, d - диаметр."
    $row = Add-GuideParagraph $Sheet $row "Для одинаковых диаметров вклад группы равен β·n·d; для разных - β, умноженному на сумму фактических диаметров группы:"
    $row = Add-GuideFormula $Sheet $row "Σβnd = β_1·D_1 + β_2·D_2 + ..."
    $row = Add-GuideParagraph $Sheet $row "D_1, D_2, ... - суммы диаметров стержней принятых групп; β_1, β_2, ... - коэффициенты этих групп. Индексы обозначают группы, не ряды."
    $row = Add-GuideFormula $Sheet $row "Σβnd = 0.85·(20 + 25) + 1.00·16 = 54.25 мм"
    $row = Add-GuideParagraph $Sheet $row "Пример для группы из двух стержней 20 и 25 мм и одиночного стержня 16 мм. При A_r в мм2 радиус R_r получается в мм. RadiusDiameterMode влияет на геометрические отступы r, но не на сумму фактических диаметров в знаменателе."

    $row = Add-GuideTitle $Sheet $row "8. Формула ширины и напряжение арматуры" 13
    $row = Add-GuideCompositeFormula $Sheet $row @(
        @{ Kind = "Text"; Text = "ψ = 1.5 · " }, @{ Kind = "Radical"; Radicand = "R_r" },
        @{ Kind = "Text"; Text = "  (Periodic);    ψ = 0.35·R_r  (Smooth)" }
    ) "СП 35, п. 7.109" "ψ = 1.5·√R_r (Periodic); ψ = 0.35·R_r (Smooth)"
    $row = Add-GuideParagraph $Sheet $row "ψ - размерный коэффициент раскрытия; R_r и ψ в этих зависимостях выражаются в сантиметрах. SLS.Crack.SP35.RebarProfile: Periodic - периодический профиль, Smooth - гладкая арматура."
    $row = Add-GuideParagraph $Sheet $row "По СП 35: Для арматуры в стенках п. 7.109 также предписывает первую зависимость; в этом случае применяется Periodic."
    $row = Add-GuideFractionFormula $Sheet $row "a_cr =" "σ_s" "E_s" "· ψ ≤ Δ_cr      СП 35, п. 7.105, ф. (7.93)"
    $row = Add-GuideParagraph $Sheet $row "a_cr - ширина нормальной трещины; σ_s - наибольшее растягивающее напряжение стержня рассматриваемой группы при текущем сочетании; E_s - модуль упругости арматуры, заданный на Config строкой Steel.E в колонке растяжения; Δ_cr - допустимая ширина из настройки Config SLS.Crack.Allowable."
    $row = Add-GuideParagraph $Sheet $row "При ψ в сантиметрах формула дает a_cr в сантиметрах. В Results ширина переводится в OUTPUT Length."
    $row = Add-GuideFractionFormula $Sheet $row "a_{cr}[мм] =" "σ_s" "E_s" "· ψ[см] · 10"
    $row = Add-GuideParagraph $Sheet $row "Напряжения соседних групп и напряжения в точке трещинообразования в формулу раскрытия не подставляются."
    $row = Add-GuideFormula $Sheet $row "a_{cr} = max(a_{cr,1}; a_{cr,2}; ...)"
    $row = Add-GuideParagraph $Sheet $row "a_cr,1, a_cr,2, ... - результаты отдельных проверок групп. Определяющая группа выбирается по ширине, а не по максимальному напряжению: большая площадь взаимодействия может дать большее раскрытие при меньшем σ_s."
    $row = Add-GuideParagraph $Sheet $row "Методическое пояснение: Формулы и пояснения приведены в разделе 4.2.3 Методического пособия 2017 года, формулы (4.306)-(4.312)."

    $row = Add-GuideTitle $Sheet $row "9. Центральное растяжение" 13
    $row = Add-GuideParagraph $Sheet $row "По СП 35: При центральном растяжении берется вся площадь бетона за вычетом отверстий. НЛ отсутствует, поэтому локальные боковые границы и отсчет r от ряда не строятся."
    $row = Add-GuideParagraph $Sheet $row "В сумму Σβnd входят все фактически растянутые группы. Напряжение σ_s берется по наиболее растянутому стержню сечения. Правила заданных и неизвестных контуров остаются теми же."
    $row = Add-GuideFormulaPicture $Sheet $row (New-SP35GuideScheme 'MeshCentral') "Без точного наружного контура границы приближенные; при центральном растяжении используется вся площадь бетона."

    $row = Add-GuideTitle $Sheet $row "10. Результаты и экспорт площади взаимодействия" 13
    $row = Add-GuideParagraph $Sheet $row "Подробный блок Results содержит определяющую группу и стержень, A_r, R_r, ψ, σ_s и a_cr. Статусы: OK - допуск соблюден; FAIL - превышен; N/A - расчет ширины не выполнялся. Причина указывается в ResultComment."
    $row = Add-GuideParagraph $Sheet $row "Поля β, n, d и rebars раскрывают состав принятых рядов в формате «1 ряд: ... | 2 ряд: ...». G1, G2, ... - идентификаторы групп; «G2=R5, R6» указывает входящие стержни. По тем же идентификаторам приводятся β, число стержней и диаметры."
    $row = Add-GuideParagraph $Sheet $row "Фактически полученную площадь взаимодействия можно вывести в AutoCAD. На Config включите AutoCAD.Export.ExportCrackInteractionContour = Yes и задайте AutoCAD.Export.CrackInteractionLayer. В профиле настройкой Visualization.State выберите одно состояние трещин: PreCrackState - перед образованием, PostCrackState - сразу после образования, CrackedState - текущее НДС с отключенным растянутым бетоном." -Emphasis
    $row = Add-GuideParagraph $Sheet $row "Экспорт использует сохраненную область Results и всегда выводит геометрию в мм. Повторный расчет для одного экспорта не нужен. Нерассчитанная или устаревшая область не выводится; превышение допуска ширины не мешает посмотреть ее контур."
    return $row + 2
}

# Возвращает один текст предупреждения для сборки и узкого обновления справки.
function Get-AutoCADContourWarningText {
    return "ВНИМАНИЕ: для расчета раскрытия трещин настоятельно рекомендуется задать точный наружный контур и все отверстия. Иначе площадь взаимодействия для трещин и ширина раскрытия могут быть завышены или занижены из-за приближенных границ сетки, особенно у криволинейных и пустотелых сечений."
}

# Самостоятельная сворачиваемая инструкция подготовки AutoCAD: разделяет
# обязательные элементы, необязательные контуры и действия пользователя
# при импорте, изменении единиц и экспорте готовых результатов.
function Add-AutoCADGeometryPreparationGuide {
    param([object]$Sheet, [int]$StartRow)

    $row = Add-GuideParagraph $Sheet $StartRow "Удобнее всего задать наружный контур, отверстия и арматуру без бетонной сетки: программа построит её сама. Готовую сетку из Region тоже можно импортировать. Объекты должны находиться в пространстве «Модель» AutoCAD. Геометрия, нарисованная на листах чертежа, не импортируется." -RowCount 2
    $row = Add-GuideTitle $Sheet $row "1. Система координат и масштаб" 13
    $row = Add-GuideParagraph $Sheet $row "Импорт и экспорт выполняются в мировой системе координат МСК/WCS: _UCS → _World. При локальной или неизвестной системе координат импорт отклоняется. Программа не переключает ПСК." -Emphasis
    $row = Add-GuideParagraph $Sheet $row "Одна единица AutoCAD равна 1 мм. Сечение должно лежать в плоскости XY при Z = 0; поворот в этой плоскости допускается. Координаты, площади и инерции импортируются в мм, мм2 и мм4 независимо от INPUT/OUTPUT и единиц чертежа. Наклонные контуры, эллипсы и сплайны не поддерживаются."
    $row = Add-GuideTitle $Sheet $row "2. Рекомендуемый способ: только контуры и арматура" 13
    $row = Add-GuideParagraph $Sheet $row "Оставьте слой из AutoCAD.Common.ConcreteLayer пустым: на нём не должно быть бетонных элементов или других объектов. Размещать предварительно разбитую бетонную сетку не требуется." -Emphasis
    $row = Add-GuideParagraph $Sheet $row "Наружную замкнутую 2D-полилинию или окружность разместите на слое AutoCAD.Common.SectionContourLayer. Каждое отверстие задайте отдельной замкнутой полилинией, окружностью или Region на AutoCAD.Common.OpeningContourLayer. Поддерживаются прямые и круговые дуги."
    $row = Add-GuideParagraph $Sheet $row "Вместо отдельных границ можно задать один Region с вырезанными отверстиями на слое наружного контура. Вырежьте отверстия командой SUBTRACT. Программа извлечёт все границы этого Region; взрывать его и дублировать отверстия на отдельном слое не нужно. Окружность, нарисованная внутри сплошного Region без вычитания, не становится его отверстием."
    $row = Add-GuideParagraph $Sheet $row "Если внутри отверстия есть отдельная бетонная часть, задайте её отдельным Region на слое наружного контура. Для автоматического построения сетки всех бетонных частей слой бетонных элементов также оставьте пустым."
    $row = Add-GuideParagraph $Sheet $row "Арматуру задайте отдельными Region на слое AutoCAD.Common.RebarLayer. Каждый принятый Region становится одним стержнем; круглый эквивалентный диаметр определяется по его площади. Стержни должны находиться в бетоне."
    $row = Add-GuideParagraph $Sheet $row "На Config выберите Geometry.Source = AutoCAD. В блоке [Сетка геометрии] задайте Mesh.CellGeometry и шаг Mesh.StepMode: Auto выбирает шаг по габаритам, User использует Mesh.StepX и Mesh.StepY. Adaptive делит пограничные прямоугольники на заданное число уровней; Contour точно обрезает базовые ячейки по прямым и дугам без граничного дробления."
    $row = Add-GuideParagraph $Sheet $row "Нажмите «Сформировать сечение». Наружный контур и все отверстия будут учтены при построении бетонной сетки. Проверьте полученную площадь, количество элементов, стержни и схему. После изменения чертежа или настроек сетки сформируйте сечение заново."
    $row = Add-GuideParagraph $Sheet $row "Общий Mesh.MinElementArea в блоке [Сетка геометрии] исключает конечные бетонные части и стержни строго меньшей площади. Ввод задаётся в INPUT Area; равная порогу площадь сохраняется, ноль отключает фильтр. Слишком большой порог может удалить важные части сечения."
    $row = Add-GuideParagraph $Sheet $row "Границы должны быть замкнутыми, без разрывов, самопересечений и дублей. Отдельные отрезки и дуги объедините командой PEDIT в замкнутую полилинию либо создайте Region. Совпадения концов полилинии достаточно; повтор начальной вершины допускается. Отдельные отрезки нулевой длины пропускаются с предупреждением." -RowCount 2
    $row = Add-GuideParagraph $Sheet $row "Отверстия не должны касаться или пересекать друг друга и наружную границу. Контурные слои разблокируйте перед импортом. Исходные объекты не меняются; программа удаляет только созданные ею временные копии."
    $row = Add-GuideParagraph $Sheet $row "Для построения без готовой сетки наружный контур обязателен: одних отверстий недостаточно. Непригодные объекты или полностью отфильтрованные Region на слое бетонных элементов не считаются пустым слоем и не запускают построение сетки." -Emphasis
    $row = Add-GuideTitle $Sheet $row "3. Дополнительный способ: готовая бетонная сетка" 13
    $row = Add-GuideParagraph $Sheet $row "Разместите каждый бетонный элемент отдельным Region на слое AutoCAD.Common.ConcreteLayer. Сетка должна полностью описывать бетон сечения, включая фактические вырезы. Полилинии, блоки и текст на этом слое не становятся расчётными элементами. Арматура задаётся Region на отдельном слое AutoCAD.Common.RebarLayer."
    $row = Add-GuideParagraph $Sheet $row "При готовой сетке отдельные наружные контуры и отверстия уточняют область взаимодействия для трещин и схему. Они не вырезают бетон из готовых элементов и не меняют их площадь или инерции. Отверстие нужно учесть в самих бетонных Region либо воспользоваться рекомендуемым построением по контурам." -Emphasis
    $row = Add-GuideParagraph $Sheet $row "В Contour сохраняется точная форма каждого бетонного Region: прямые, круговые дуги и вырезанные отверстия; несвязные бетонные части становятся отдельными элементами. В Adaptive используются прямоугольные оболочки, построенные по фактическим площади и инерциям. Готовые бетонные элементы заново по шагу Mesh не разбиваются."
    $warningRow = $row
    $row = Add-GuideParagraph $Sheet $row (Get-AutoCADContourWarningText) -Emphasis
    $Sheet.Range($Sheet.Cells.Item($warningRow, 1), $Sheet.Cells.Item($row - 1, 6)).Font.Color = 255
    $row = Add-GuideParagraph $Sheet $row "Отсутствующий или пустой контурный слой допустим при готовой бетонной сетке. Для определения области трещин используются заданные точные границы либо границы элементов; четыре варианта перечислены в конце этого раздела. Ни один из них не перестраивает готовую сетку НДС."

    $row = Add-GuideParagraph $Sheet $row "Сформированное сечение можно экспортировать в AutoCAD до расчёта. Выгружаются только бетон, арматура и включённые геометрические контуры, без напряжений, деформаций и нейтральной линии. Предупреждение об отсутствии рассчитанного состояния показывается всегда."

    $row = Add-GuideTitle $Sheet $row "4. Импорт и контроль модели" 13
    $row = Add-GuideParagraph $Sheet $row "1. На Config выберите настройку Geometry.Source = AutoCAD."
    $row = Add-GuideParagraph $Sheet $row "2. Нажмите «Сформировать сечение». Проверьте число бетонных элементов и стержней, площадь, размеры и схему. Сообщение показывает время формирования и предупреждения о полученной геометрии."
    $row = Add-GuideParagraph $Sheet $row "3. Задайте сочетания и их расчетные профили на Config, затем нажмите «Выполнить расчет». После изменения чертежа заново нажмите «Сформировать сечение» перед расчётом."
    $row = Add-GuideParagraph $Sheet $row "Геометрия Results сохраняется с собственными единицами OUTPUT на момент импорта. Последующая смена единиц не меняет физические размеры сохраненного сечения; новый расчет выводится в текущих OUTPUT без повторного импорта."
    $row = Add-GuideParagraph $Sheet $row "При смене INPUT пересчитайте вручную уже введенные нагрузки, параметры материалов и размерные настройки Config. Это не относится к сохраненной импортированной геометрии. Если новый импорт завершился ошибкой, прежняя геометрия Results сохраняется."

    $row = Add-GuideTitle $Sheet $row "5. Экспорт сохраненных результатов" 13
    $row = Add-GuideParagraph $Sheet $row "Экспорт идёт в пространство «Модель» активного чертежа, даже если открыт лист. Используется последний расчёт на Results; после изменения нагрузок рассчитайте заново. Геометрия передаётся в мм независимо от OUTPUT. «Очистить подписи в AutoCAD» удаляет оформление, сохраняя бетон, арматуру и контуры; без слоёв оформления очистка сразу завершается." -Emphasis -RowCount 2
    $row = Add-GuideParagraph $Sheet $row "На Config включите AutoCAD.Export.ContourEnabled = Yes и выберите настройку AutoCAD.Export.ContourFormat:"
    $row = Add-GuideParagraph $Sheet $row "Region (по умолчанию) - бетонная область со всеми отверстиями на AutoCAD.Common.SectionContourLayer. Отдельные бетонные части экспортируются отдельными Region."
    $row = Add-GuideParagraph $Sheet $row "Polyline - отдельные замкнутые полилинии: наружная граница на AutoCAD.Common.SectionContourLayer, отверстия на AutoCAD.Common.OpeningContourLayer. Прямые участки и дуги сохраняются."
    $row = Add-GuideParagraph $Sheet $row "Если сохранены только отверстия, экспортируются только они на слой отверстий; наружная граница по сетке не создается. Оба формата применяются к сгенерированным и импортированным сечениям и поддерживают обратный импорт."
    $row = Add-GuideParagraph $Sheet $row "Каждый экспорт добавляет объекты. Совпадения не ищутся, старые контуры не удаляются: заранее подготовьте свободное место или уберите ненужные копии. Ответственность за чистоту чертежа при повторном импорте поясняется ниже."
    $row = Add-GuideParagraph $Sheet $row "Для экспорта рассчитанной площади взаимодействия для трещин, принятой в расчете ширины раскрытия, включите на Config AutoCAD.Export.ExportCrackInteractionContour = Yes и задайте AutoCAD.Export.CrackInteractionLayer. На том же листе в таблице «Настройка расчетных профилей», подразделе «Настройки визуализации», строке «Выводимое состояние» (Visualization.State) выберите PreCrackState, PostCrackState или CrackedState в колонке профиля нужного сочетания."
    $row = Add-GuideParagraph $Sheet $row "PreCrackState - состояние перед образованием трещины; PostCrackState - сразу после образования; CrackedState - текущее НДС с отключенным растянутым бетоном. Если в той же строке Visualization.State выбрано StrengthState (прочность) или CapacityState (несущая способность), площадь взаимодействия для трещин не экспортируется."
    $row = Add-GuideParagraph $Sheet $row "Экспортируется принятая в расчете ширины раскрытия площадь взаимодействия для трещин, сохраненная на Results и согласованная с его геометрией. Превышение допустимой ширины не препятствует экспорту ее контура. Актуальность после изменения исходных данных поясняется ниже."
    $row = Add-GuideParagraph $Sheet $row "Кнопка очистки удаляет оформление, но сохраняет бетон, арматуру, наружные контуры и отверстия, включая прежние экспортированные копии. Повторный экспорт сам чертеж не очищает."
    $row = Add-GuideParagraph $Sheet $row "Если бетонная сетка автоматически создана по контурам AutoCAD, сообщение об этом появляется только при формировании сечения. После расчёта и экспорта это сообщение не повторяется. Сведения о происхождении сетки сохраняются на листе Results." -Emphasis

    $row = Add-GuideTitle $Sheet $row "6. Поясняющие примечания" 13
    $row = Add-GuideParagraph $Sheet $row "Выбор границ площади взаимодействия для трещин зависит от заданных контуров. Если в Region на слое наружного контура отверстия уже вырезаны командой SUBTRACT, программа автоматически импортирует их как заданные отверстия. Отдельно размещать их на слое отверстий не нужно. Применяются четыре случая:"
    $row = Add-GuideParagraph $Sheet $row "1. Заданы только отверстия. Наружная граница берется по бетонной сетке, а вычитаются только заданные отверстия. Другие пустоты сетки не учитываются."
    $row = Add-GuideParagraph $Sheet $row "2. Задан только наружный контур. Внутри него принимается сплошной бетон, без отверстий. Если отверстия есть в действительности, задайте их вместе с наружным контуром: пустоты сетки их не заменят."
    $row = Add-GuideParagraph $Sheet $row "3. Заданы наружный контур и отверстия. Используется точная область внутри наружного контура за вычетом всех заданных отверстий. Границы сетки для этой площади не используются."
    $row = Add-GuideParagraph $Sheet $row "4. Контуры не заданы. Площадь взаимодействия для трещин определяется по бетонной сетке, включая все ее пустоты. Границы получаются приближенными."
    $labelRow = $row
    $row = Add-GuideParagraph $Sheet $row "Чистота чертежа и повторный импорт. Импортируются все объекты на заданных слоях независимо от того, созданы они пользователем или экспортированы программой. Ответственность за чистоту чертежа и отсутствие наложенных друг на друга дублей бетона, арматуры, наружных контуров и отверстий несет пользователь."
    $Sheet.Cells.Item($labelRow, 1).Characters(1, "Чистота чертежа и повторный импорт.".Length).Font.Bold = $true
    $row = Add-GuideParagraph $Sheet $row "Экспорт добавляет новые объекты и не удаляет прежние. Перед повторным импортом оставьте на импортируемых слоях только нужный набор геометрии без дублей; лишние копии удалите или перенесите на другие слои. Дубли контуров могут привести к ошибке импорта, дубли бетона и арматуры - к искажению расчета."
    $labelRow = $row
    $row = Add-GuideParagraph $Sheet $row "Актуальность площади взаимодействия для трещин. Она относится к геометрии, материалам и сочетанию сохраненного расчета. После изменения нагрузок, материалов или расчетных настроек выполните расчет заново: изменение Config само по себе не обновляет Results и не блокирует экспорт прежних результатов. Новый импорт заменяет геометрию и удаляет ранее рассчитанную площадь; после него требуется расчет."
    $Sheet.Cells.Item($labelRow, 1).Characters(1, "Актуальность площади взаимодействия для трещин.".Length).Font.Bold = $true
    $row = Add-GuideParagraph $Sheet $row "Если вручную скопировать в Results часть таблиц из другого расчета, площадь взаимодействия для трещин может оказаться несогласованной с сохраненной геометрией. Такой контур не экспортируется. Для восстановления согласованного Results повторите расчет."
    return $row + 2
}

# Формирует методический блок по фильтру продольного изгиба/устойчивости.
# Формулы выводятся тем же PNG-рендером, что и в разделе трещин, чтобы лист
# "Справка" оставался читаемым без OLE Equation и ручной верстки.
function Add-StabilityMethodologyGuide {
    param([object]$Sheet, [int]$StartRow)

    $row = Add-GuideParagraph $Sheet $StartRow "Проверка устойчивости выполняется перед расчетом НДС и учитывает случайный эксцентриситет и продольный изгиб по выбранной методике. Поддерживается ненапрягаемая арматура."
    $row = Add-GuideTitle $Sheet $row "1. Исходные данные" 13
    $row = Add-GuideParagraph $Sheet $row "Расчет включается настройкой профиля Calculation.Stability.Enabled = Yes. Stability.Code выбирает SP63 или SP35, MaterialModel.Stability.ValueSet - набор характеристик материалов. Для нормативной проверки применяется ULS(I)."
    $row = Add-GuideParagraph $Sheet $row "В таблице «Нагрузки для устойчивости» задайте постоянные + длительные нагрузки для SP63 либо постоянные для SP35. CombinationID связывает их с полным сочетанием. Числа вводятся в тех же единицах INPUT и знаках."
    $row = Add-GuideParagraph $Sheet $row "Для каждого сочетания с включенной устойчивостью добавьте строку с тем же CombinationID. В этой строке пустые N, Mx и My означают 0, а отсутствие самой строки является ошибкой ввода. Ошибочная формула и текст требуют исправления; длительную часть программа сама не выделяет и полной нагрузкой автоматически не ограничивает." -RowCount 1
    $row = Add-GuideParagraph $Sheet $row "Если устойчивость не проходит, прочность и трещины продолжаются с исходными усилиями для полноты диагностики. Общий статус сохраняет отказ устойчивости. При успешной проверке последующие расчеты получают увеличенные моменты." -Emphasis

    $row = Add-GuideTitle $Sheet $row "2. Главные плоскости и характеристики" 13
    $row = Add-GuideParagraph $Sheet $row "Программа переносит усилия от пользовательской точки к центру приведенного сечения бетон + арматура и поворачивает их в главные оси. Пользовательская точка задается на Config настройками Load.ReferenceOffsetX и Load.ReferenceOffsetY относительно центра бетона без арматуры."
    $row = Add-GuideFormula $Sheet $row "M_1 = M_x,red·cos α + M_y,red·sin α;    M_2 = -M_x,red·sin α + M_y,red·cos α"
    $row = Add-GuideParagraph $Sheet $row "M_1/M_2 - моменты в главных плоскостях; M_x,red/M_y,red - моменты по X/Y относительно приведенного центра. α - угол главной оси 1 от X. Для почти изотропного сечения фиксируется α = 0."
    $row = Add-GuideParagraph $Sheet $row "При α = 0 плоскость 1 соответствует Mx и размеру по Y, плоскость 2 - My и размеру по X. Площадь и инерции для устойчивости берутся по расчетной сетке; точные контуры трещин эту проверку не меняют."
    $row = Add-GuideFractionFormula $Sheet $row "n_s =" "E_s" "E_b" ""
    $row = Add-GuideParagraph $Sheet $row "Коэффициент приведения n_s определяется по модулям выбранной модели: Steel.E и Concrete.E на Config."
    $row = Add-GuideFormula $Sheet $row "A_red = A_b + (n_s - 1)·A_s"
    $row = Add-GuideParagraph $Sheet $row "A_red - приведенная площадь; A_b - площадь бетона, A_s - площадь арматуры. Вычитается единица, потому что бетон в месте стержней уже входит в сетку."
    $row = Add-GuideCompositeFormula $Sheet $row @(
        @{ Kind = "Text"; Text = "i = " },
        @{ Kind = "Text"; Text = "√(" },
        @{ Kind = "Fraction"; Numerator = "I_red"; Denominator = "A_red" },
        @{ Kind = "Text"; Text = ")" }
    ) "" "i = √(I_red/A_red)"
    $row = Add-GuideParagraph $Sheet $row "i и I_red - радиус и приведенный момент инерции проверяемой главной плоскости."
    $row = Add-GuideFractionFormula $Sheet $row "r =" "I_red" "A_red·c" ""
    $row = Add-GuideParagraph $Sheet $row "r - ядровое расстояние для выбора ветви СП 35; c - расстояние от приведенного центра до соответствующей крайней бетонной границы. r относится к линейному приведенному сечению, а не к точному началу растяжения при нелинейном НДС."

    $row = Add-GuideTitle $Sheet $row "3. Длина и случайный эксцентриситет" 13
    $row = Add-GuideFormula $Sheet $row "l_0,1 = μ_1·L;    l_0,2 = μ_2·L"
    $row = Add-GuideParagraph $Sheet $row "L = Stability.ElementLength; μ_1/μ_2 задаются в Stability.Mu1/Mu2 по принятой расчетной схеме. l_0,1/l_0,2 - расчетные длины главных плоскостей."
    $row = Add-GuideParagraph $Sheet $row "Настройка Config Stability.AccidentalEccentricityMode выбирает правило: AutoWithL - по L; AutoWithμL - по расчетной длине; User - значения из Stability.AccidentalEccentricityUser1 и Stability.AccidentalEccentricityUser2. AutoWithμL - проверочная опция программы, не замена нормативного правила по геометрической длине."
    $row = Add-GuideCompositeFormula $Sheet $row @(
        @{ Kind = "Text"; Text = "e_a = max(" },
        @{ Kind = "Fraction"; Numerator = "L"; Denominator = "600" },
        @{ Kind = "Text"; Text = "; " },
        @{ Kind = "Fraction"; Numerator = "h"; Denominator = "30" },
        @{ Kind = "Text"; Text = "; 10 мм)" }
    ) "СП 63, п. 8.1.7" "e_a = max(L/600; h/30; 10 мм)"
    $row = Add-GuideParagraph $Sheet $row "e_a - случайный эксцентриситет по СП 63; h - бетонный габарит проверяемой плоскости."
    $row = Add-GuideFractionFormula $Sheet $row "e_{c,sl} =" "L" "400" "      СП 35, п. 7.52"
    $row = Add-GuideParagraph $Sheet $row "e_c,sl - случайный эксцентриситет по СП 35. Он учитывается в усилиях без смещения геометрии сечения."
    $row = Add-GuideParagraph $Sheet $row "Настройка Config Stability.AccidentalEccentricityPlanes: BothPlanes учитывает его в обеих главных плоскостях; OnlyMomentPlane - только в плоскости с моментом, если момент задан в одной. Последний вариант - принятое упрощение проекта."
    $row = Add-GuideParagraph $Sheet $row "При ненулевом моменте случайное плечо направляется в сторону увеличения его модуля. При нулевом моменте направление задают настройки Config Stability.ZeroMomentEccentricitySign1 и Stability.ZeroMomentEccentricitySign2: 1 или -1."

    $row = Add-GuideTitle $Sheet $row "4. Расчетный эксцентриситет" 13
    $row = Add-GuideFractionFormula $Sheet $row "e_stat =" "M" "N" ""
    $row = Add-GuideParagraph $Sheet $row "e_stat - эксцентриситет исходной нагрузки в проверяемой главной плоскости. В формулах устойчивости N принимается положительным модулем сжимающей силы."
    $row = Add-GuideParagraph $Sheet $row "На Config настройка Stability.SystemType = Determinate задает статически определимую схему: фактическое и случайное плечи складываются. Indeterminate - статически неопределимую: берется большее по модулю с направлением расчетного момента."
    $row = Add-GuideFormula $Sheet $row "Determinate: e_0 = e_stat + e_a;    Indeterminate: e_0 = sign·max(|e_stat|; |e_a|)"
    $row = Add-GuideParagraph $Sheet $row "e_0 - расчетный эксцентриситет СП 63; sign - направление расчетного момента. Для СП 35 используются e_c и e_c,sl. Порог обнуления малых моментов задается Calculation.ZeroMomentPerDepth."

    $row = Add-GuideTitle $Sheet $row "5. Продольный изгиб по СП 63" 13
    $row = Add-GuideFractionFormula $Sheet $row "φ_l = 1 +" "M_{l1}" "M_1" "; φ_l ≤ 2      СП 63, п. 8.1.15"
    $row = Add-GuideParagraph $Sheet $row "φ_l - коэффициент длительности. В этой нормативной формуле M_1 - полный момент относительно наиболее растянутого или наименее сжатого стержня, M_l1 - соответствующая длительная часть. Здесь индекс 1 обозначает эти нормативные моменты, не номер главной плоскости. При отсутствии арматуры используется край бетона."
    $row = Add-GuideFractionFormula $Sheet $row "δ_e =" "|e_0|" "h" ""
    $row = Add-GuideParagraph $Sheet $row "δ_e - относительный эксцентриситет. Его границы на Config задают Stability.SP63.DeltaEMin и Stability.SP63.DeltaEMax."
    $row = Add-GuideFractionFormula $Sheet $row "k_b =" "0.15" "φ_l·(0.3 + δ_e)" ""
    $row = Add-GuideParagraph $Sheet $row "k_b - коэффициент бетонного вклада в жесткость."
    $row = Add-GuideFormula $Sheet $row "D = k_b·E_b·I_b + k_s·E_s·I_s      СП 63, п. 8.1.15"
    $row = Add-GuideParagraph $Sheet $row "D - условная жесткость; I_b/I_s - моменты инерции бетона/арматуры в проверяемой плоскости. k_s задается настройкой Config Stability.SP63.Ks."
    $row = Add-GuideFractionFormula $Sheet $row "N_cr =" "π^2·D" "l_0^2" ""
    $row = Add-GuideParagraph $Sheet $row "N_cr - критическая сила; l_0 - расчетная длина проверяемой плоскости."
    $row = Add-GuideFractionFormula $Sheet $row "η =" "N_cr" "N_cr - N" ""
    $row = Add-GuideParagraph $Sheet $row "η - коэффициент увеличения момента. При гибкости l_0/i ≤ 14 он принимается равным 1, хотя N_cr выводится справочно. Если N достигла N_cr, положительный коэффициент получить нельзя и устойчивость получает FAIL."
    $row = Add-GuideFormula $Sheet $row "M_design = η·N·e_0"
    $row = Add-GuideParagraph $Sheet $row "M_design - момент с учетом продольного изгиба. Stability.PhiLMode = PhiL2 задает φ_l = 2; Auto определяет его по нагрузкам с ограничением 1...2."

    $row = Add-GuideTitle $Sheet $row "6. СП 35 при большом эксцентриситете" 13
    $row = Add-GuideParagraph $Sheet $row "При |e_c| > r применяются критическая сила и увеличение момента. Коэффициент длительности считается по моментам относительно наиболее растянутой или наименее сжатой грани бетона:"
    $row = Add-GuideFractionFormula $Sheet $row "φ_l = 1 +" "M_l" "M" "; 1 ≤ φ_l ≤ 2      СП 35, п. 7.54, ф. (7.11)"
    $row = Add-GuideParagraph $Sheet $row "M - полный момент относительно этой грани, M_l - его постоянная часть. При противоположных знаках: если |e_c| ≤ 0,1h, берется φ_l = 1, иначе 1,05. Настройка Config Stability.PhiLMode = PhiL2 принимает 2."
    $row = Add-GuideCompositeFormula $Sheet $row @(
        @{ Kind = "Text"; Text = "δ = max(" },
        @{ Kind = "Fraction"; Numerator = "|e_c|"; Denominator = "h" },
        @{ Kind = "Text"; Text = "; 0.5 - 0.01·" },
        @{ Kind = "Fraction"; Numerator = "l_0"; Denominator = "h" },
        @{ Kind = "Text"; Text = " - 0.01·R_b)" }
    ) "" "δ = max(|e_c|/h; 0.5 - 0.01·l_0/h - 0.01·R_b)"
    $row = Add-GuideParagraph $Sheet $row "δ - относительная величина для жесткости; R_b - сопротивление бетона сжатию в МПа из настройки Config Concrete.R.ULS(I), колонки сжатия."
    $row = Add-GuideFractionFormula $Sheet $row "δ_p =" "δ" "φ_p" ""
    $row = Add-GuideParagraph $Sheet $row "φ_p задается настройкой Stability.SP35.PhiP."
    $row = Add-GuideFractionFormula $Sheet $row "k_1 =" "0.11" "0.1 + δ_p" "+ 0.1"
    $row = Add-GuideParagraph $Sheet $row "k_1 учитывает бетонный вклад в условную жесткость."
    $row = Add-GuideCompositeFormula $Sheet $row @(
        @{ Kind = "Text"; Text = "D = 6.4·E_b·[" },
        @{ Kind = "Fraction"; Numerator = "I_b"; Denominator = "φ_l" },
        @{ Kind = "Text"; Text = "·k_1 + " },
        @{ Kind = "Fraction"; Numerator = "E_s"; Denominator = "E_b" },
        @{ Kind = "Text"; Text = "·I_s]" }
    ) "СП 35, п. 7.54, ф. (7.10)" "D = 6.4·E_b·[I_b/φ_l·k_1 + (E_s/E_b)·I_s]"
    $row = Add-GuideParagraph $Sheet $row "D - условная жесткость этой ветви; модули E_b/E_s и инерции I_b/I_s имеют тот же смысл, что выше."
    $row = Add-GuideFractionFormula $Sheet $row "N_cr =" "D" "l_0^2" "      СП 35, п. 7.54, ф. (7.10)"
    $row = Add-GuideParagraph $Sheet $row "В этой формуле СП 35 нет дополнительного множителя π². Коэффициент η определяется как N_cr/(N_cr - N), расчетный момент - η·N·e_c."
    $row = Add-GuideParagraph $Sheet $row "Проверяется отношение N/N_cr. Его предел задается на Config настройкой Stability.SP35.NOverNcrLimit; запас равен этому пределу, умноженному на N_cr/N. При запасе меньше 1 устойчивость не проходит."

    $row = Add-GuideTitle $Sheet $row "7. СП 35 при малом эксцентриситете" 13
    $row = Add-GuideParagraph $Sheet $row "При |e_c| ≤ r коэффициенты берутся из таблицы 7.21 на Config. φ_m - коэффициент кратковременной нагрузки, φ_l - длительной; они интерполируются по гибкости l_0/i и, для φ_m, по отношению |e_c|/r."
    $row = Add-GuideFractionFormula $Sheet $row "q_l =" "N_l" "N" ""
    $row = Add-GuideParagraph $Sheet $row "q_l - доля постоянной сжимающей силы из таблицы нагрузок для устойчивости. Она не является табличным коэффициентом φ_l."
    $row = Add-GuideFractionFormula $Sheet $row "φ_{ml} =" "φ_m" "φ_l" ""
    $row = Add-GuideParagraph $Sheet $row "φ_m и φ_l принимаются по таблице 7.21; φ_ml - их отношение."
    $row = Add-GuideFractionFormula $Sheet $row "φ =" "φ_m" "q_l·φ_{ml} + (1 - q_l)" "      СП 35, п. 7.55, ф. (7.15)"
    $row = Add-GuideParagraph $Sheet $row "φ - итоговый коэффициент снижения несущей способности с учетом постоянной части. При ее отсутствии он равен φ_m, при полностью постоянной нагрузке - φ_l."
    $row = Add-GuideFormula $Sheet $row "N_ult,stab = φ·(R_b·A_b,stab + R_sc·A_s)"
    $row = Add-GuideParagraph $Sheet $row "N_ult,stab - предельная сила по устойчивости; R_sc задается в Steel.R.ULS(I). A_b,stab - бетонная площадь с поправкой данной ветви; A_s - площадь арматуры."
    $row = Add-GuideParagraph $Sheet $row "Если A_s больше 3% площади бетона, бетонный вклад уменьшается на A_s. Это поправка только данной формулы, не повторное вычитание арматуры из приведенного сечения. Настройка Config Stability.PhiLMode здесь не заменяет табличный φ_l."
    $row = Add-GuideFractionFormula $Sheet $row "SF_stab =" "N_ult,stab" "N" ""
    $row = Add-GuideParagraph $Sheet $row "SF_stab - коэффициент запаса устойчивости. Если он меньше 1, полная сила превышает найденный предел."

    $row = Add-GuideTitle $Sheet $row "8. Результаты проверки" 13
    $row = Add-GuideParagraph $Sheet $row "Подробная таблица показывает итоги, данные обеих главных плоскостей и только выбранный нормативный блок. При растяжении без сжатой силы проверка устойчивости неприменима и получает N/A."
    $row = Add-GuideParagraph $Sheet $row "Моменты приведены по главным осям, по центральным осям X/Y и в пользовательской точке. OUTPUT и пользовательские знаки выбираются на Config в таблицах единиц и знаков. Комментарий показывает ветвь, причину отказа или ошибку исходных данных."
    return $row + 2
}

# Формирует финальный справочный раздел по листу Results.
# Здесь описываются не настройки Config, а смысл основных выходных диапазонов
# расчетного снимка, особенно геометрические характеристики бетонного и
# приведенного сечения.
function Get-BatchSummaryInstruction {
    return "Над сводкой показаны время полного расчета и объем поиска: вызовы решателя НДС, его итерации, старты и итерации UltimateStrain. Старт - реально предпринятая попытка с начальной плоскостью; учитываются и неудачные попытки. Счетчики суммируются по всем сочетаниям одного запуска и обнуляются перед следующим."
}

function Add-ResultsSheetGuide {
    param([object]$Sheet, [int]$StartRow)

    $row = $StartRow
    $row = Add-GuideParagraph $Sheet $row "Лист Results хранит результаты последнего завершенного запуска, включая не прошедшие проверки. Схема и экспорт в AutoCAD читают эти сохраненные данные. Изменение настроек на листе Config само по себе не обновляет Results." -RowCount 1 -Emphasis
    $row = Add-GuideParagraph $Sheet $row "Числа выводятся в выбранных OUTPUT и пользовательских знаках. Нагрузки соответствуют правилам Config Sign.N.User, Sign.Mx.User, Sign.My.User; напряжения и деформации сохраняют физический знак: плюс - растяжение, минус - сжатие. Единицы заданы на Config ключами Units.*.Output, где звездочка означает Length, Area, Force, Moment, Stress или Curvature."
    $row = Add-GuideParagraph $Sheet $row "Ширину столбцов можно менять для удобства чтения: расчет и очистка сохраняют ваш выбор. В строках данных перенос текста отключен; при длинном комментарии расширьте столбец."
    $row++
    $row = Add-GuideTitle $Sheet $row "1. Сводка и подробные проверки" 13
    $row = Add-GuideParagraph $Sheet $row "В верхней сводке одна строка на сочетание: общий статус, прочность, трещины, устойчивость и запасы. Worst означает определяющее сочетание; критерий выбирается на Config в General.WorstCombinationCriterion. N/A не означает успех: проверка может быть выключена, неприменима или не иметь обязательного результата. Причина указана рядом."
    $row = Add-GuideParagraph $Sheet $row (Get-BatchSummaryInstruction)
    $row = Add-GuideParagraph $Sheet $row "В конце верхней сводки показаны N, Mx, My в точке приложения: исходные до фильтров и рабочие после учёта продольного изгиба, если он применим. Коэффициент η изменяет моменты, а не N. Выделенная строка нагрузок относится к определяющему сочетанию в самом верху сводки. Новые файловые снимки сохраняют этот блок и восстанавливают его без чтения текущих нагрузок Config. В снимках без этих необязательных полей значения остаются пустыми."
    $row = Add-GuideParagraph $Sheet $row "Блок прочности содержит текущее НДС и найденное предельное состояние; блок трещин - трещинообразование, ширину нормальных и проверку продольных трещин; блок устойчивости - продольный изгиб и итоговые усилия. Методики приведены выше."
    $row = Add-GuideParagraph $Sheet $row "ResultComment содержит расчетную ветвь, причину отказа или пропуска, предупреждения и адреса ошибок ввода. Общий комментарий объединяет сообщения отдельных проверок."
    $row++
    $row = Add-GuideTitle $Sheet $row "2. Элементы и известные контуры" 13
    $row = Add-GuideParagraph $Sheet $row "Таблица rngNDMSectionGeometry хранит бетонные элементы и арматуру. Area — площадь; LocalIx/LocalIy/LocalIxy — собственные моменты инерции. Width/Height/Rotation описывают прямоугольные элементы Adaptive. В Contour схема и экспорт используют фактические границы из rngNDMElementGeometry. ExtraData содержит необязательные сведения об источнике; расчёт их не использует."
    $row = Add-GuideParagraph $Sheet $row "В Contour при импорте площадь, центр и инерции вычисляются по фактической границе элемента. Схема и экспорт сохраняют прямые, дуги и отверстия. В Adaptive для схемы строится эквивалентный прямоугольник, если инерции позволяют определить его стороны и поворот; иначе — равноплощадный квадрат."
    $row = Add-GuideParagraph $Sheet $row "Отдельная таблица rngNDMSectionContours хранит достоверные наружные границы и отверстия - прямые, дуги и окружности. ID - имя контура, Outer - наружная граница, Opening - отверстие. Эти границы сохраняются один раз на сечение, а не на каждое сочетание."
    $row = Add-GuideParagraph $Sheet $row "Отверстия могут быть известны без наружного контура. Пустая таблица означает только отсутствие заданных точных границ, а не отсутствие пустот. Правила сочетания контуров и сетки перечислены в конце раздела подготовки AutoCAD."
    $row = Add-GuideParagraph $Sheet $row "Если повреждены координаты или обязательные поля, сообщение укажет ячейку Results. Нажмите «Сформировать сечение». Расчёт не перестраивает геометрию; размерные линии и подписи на схеме ему не нужны." -RowCount 1
    $row++
    $row = Add-GuideTitle $Sheet $row "3. Деформации и напряжения элементов" 13
    $row = Add-GuideParagraph $Sheet $row "В режиме Excel в таблице rngNDMElementResults сохраняются деформации и напряжения каждого элемента для полученных расчетных состояний. StrengthState — текущее НДС по прочности, CapacityState — предельное состояние, PreCrackState — состояние перед образованием трещины, PostCrackState — сразу после ее образования, CrackedState — текущее НДС с отключенным растянутым бетоном." -RowCount 2
    $row = Add-GuideParagraph $Sheet $row "В режиме Excel данные сочетаний и области трещин сохраняются в таблицах Results; кнопка «Обновить схему и результаты» обновляет только схему. В режиме File снимок с именем Results.SnapshotName хранится рядом с книгой. Та же кнопка загружает из него схему и подробные результаты без нового расчёта; полные состояния и границы элементов остаются в файле, на листе — их шапки." -RowCount 3
    $row = Add-GuideParagraph $Sheet $row "Сохраняются только фактически полученные состояния. Для схемы и AutoCAD выберите на Config в профиле Visualization.State и Visualization.Quantity: Stress - напряжение, Strain - деформация. Если выбранного состояния нет, программа сообщит причину и не подменит его другим."
    $row++
    $row = Add-GuideTitle $Sheet $row "4. Свойства всего сечения" 13
    $row = Add-GuideParagraph $Sheet $row "Таблица rngNDMSectionProperties показывает два представления: Concrete - только бетон; Transformed - приведенное сечение, где арматура заменена эквивалентной по жесткости площадью бетона."
    $row = Add-GuideFractionFormula $Sheet $row "n_s =" "E_s" "E_b" ""
    $row = Add-GuideParagraph $Sheet $row "Для приведения используются модули выбранной модели. Поскольку бетонная сетка включает площадь под стержнями, дополнительный вклад арматуры равен (n_s - 1)·A_s."
    $row = Add-GuideParagraph $Sheet $row "Area - площадь; CentroidX/Y - координаты центра тяжести. Ix/Iy/Ixy - моменты инерции относительно осей модели; Ixc/Iyc/Ixyc - относительно параллельных осей через центр тяжести. Индекс c здесь означает центральный момент."
    $row = Add-GuideParagraph $Sheet $row "PrincipalAngle - угол главной центральной оси 1 от +X, положительный против часовой стрелки. PrincipalI1 - больший главный момент инерции; PrincipalI2 - меньший. RadiusX/Y и PrincipalRadius1/2 - радиусы инерции по центральным и главным осям."
    $row++
    $row = Add-GuideTitle $Sheet $row "5. Ядровое расстояние" 13
    $row = Add-GuideParagraph $Sheet $row "CoreDistance - ядровое расстояние r для линейного сечения. Оно используется при выборе ветви устойчивости СП 35."
    $row = Add-GuideFractionFormula $Sheet $row "r =" "I" "A·c" ""
    $row = Add-GuideParagraph $Sheet $row "A - площадь, I - центральный момент инерции выбранного представления Concrete или Transformed; c - расстояние до крайней бетонной границы на стороне, противоположной смещению силы. Арматура меняет приведенные A/I, но не расширяет бетонный контур."
    $row = Add-GuideParagraph $Sheet $row "Plus/Minus относятся к направлению смещения силы. Например CoreDistanceYMinus - предел для смещения в сторону -Y; крайнее волокно для него находится с другой стороны. X/Y - центральные оси, 1/2 - главные оси. Для выбора ветви устойчивости используются значения приведенного сечения в соответствующей плоскости."
    $row = Add-GuideParagraph $Sheet $row "Это линейная характеристика. При нелинейной работе материалов, плато диаграммы или отключении растянутого бетона появление растяжения в НДС не обязано точно совпадать с ядровой границей."
    $row++
    $row = Add-GuideTitle $Sheet $row "6. Диаграммы, подписи и площадь взаимодействия" 13
    $row = Add-GuideParagraph $Sheet $row "Таблица rngNDMMaterialDiagrams хранит фактические точки диаграмм материалов последнего расчета. По ней можно проверить, какая модель использовалась для нужного состояния."
    $row = Add-GuideParagraph $Sheet $row "Таблица rngNDMSectionAnnotations хранит размерные линии, подписи арматуры и рассчитанную для сочетания площадь взаимодействия A_r. Исходные наружные контуры и отверстия сохраняются отдельно в таблице rngNDMSectionContours." -RowCount 1
    $row = Add-GuideParagraph $Sheet $row "Фактически полученную площадь взаимодействия можно вывести в AutoCAD: задайте на Config AutoCAD.Export.ExportCrackInteractionContour = Yes и в профиле Visualization.State выберите одно из состояний трещин. Экспорт использует сохраненную площадь, без нового расчета, и всегда выводит геометрию в мм." -Emphasis
    return $row + 2
}

# Добавляет лист "Справка" и прямые гиперссылки "Подробнее" из Config на конкретные ячейки справки.
function Add-SettingsInstructions {
    param([object]$Workbook, [object]$ConfigSheet, [object]$InstructionSheet)

    Clear-GuideFormulaPictures $InstructionSheet
    $script:FormulaImageCounter = 0
    $script:FormulaImageRows.Clear()

    $InstructionSheet.Cells.UnMerge() | Out-Null
    $InstructionSheet.Cells.Clear()
    try { $InstructionSheet.Cells.ClearOutline() } catch {}
    try { $InstructionSheet.Rows.Hidden = $false } catch {}
    try { $InstructionSheet.Outline.SummaryRow = 0 } catch {}
    $InstructionSheet.Cells.Validation.Delete()
    $InstructionSheet.Cells.Font.Name = "Arial"
    $InstructionSheet.Cells.Font.Size = 10
    $InstructionSheet.Columns.Item(1).ColumnWidth = 34
    $InstructionSheet.Columns.Item(2).ColumnWidth = 118
    $InstructionSheet.Columns.Item(3).ColumnWidth = 18
    $InstructionSheet.Columns.Item(4).ColumnWidth = 18
    $InstructionSheet.Columns.Item(5).ColumnWidth = 26
    $InstructionSheet.Columns.Item(6).ColumnWidth = 26
    $InstructionSheet.Columns.Item(1).WrapText = $true
    $InstructionSheet.Columns.Item(2).WrapText = $true
    $InstructionSheet.Columns.Item(3).WrapText = $true
    $InstructionSheet.Columns.Item(4).WrapText = $true
    $InstructionSheet.Columns.Item(5).WrapText = $true
    $InstructionSheet.Columns.Item(6).WrapText = $true

    $row = 1
    $row = Add-CollapsibleGuideSection $InstructionSheet $row "Подготовка геометрии в AutoCAD и импорт" {
        param($sheet, $startRow)
        Add-AutoCADGeometryPreparationGuide $sheet $startRow
    }
    $row = Add-CollapsibleGuideSection $InstructionSheet $row "Поиск НДС текущего сочетания" {
        param($sheet, $startRow)
        Add-DirectStateMethodologyGuide $sheet $startRow
    }
    $row = Add-CollapsibleGuideSection $InstructionSheet $row "Поиск предельной несущей способности" {
        param($sheet, $startRow)
        Add-CapacityMethodologyGuide $sheet $startRow
    }
    $row = Add-CollapsibleGuideSection $InstructionSheet $row "Расчет ширины раскрытия нормальных трещин" {
        param($sheet, $startRow)
        Add-CrackWidthMethodologyGuide $sheet $startRow
    }
    $row = Add-CollapsibleGuideSection $InstructionSheet $row "Расчет продольного изгиба и устойчивости" {
        param($sheet, $startRow)
        Add-StabilityMethodologyGuide $sheet $startRow
    }
    $row = Add-GuideTitle $InstructionSheet $row "Справка по настройкам Config и листа «Сечения»" 16
    $configSettingsBodyStartRow = $row

    $InstructionSheet.Cells.Item($row, 1).Value2 = "Нормативные источники"
    $InstructionSheet.Cells.Item($row, 1).Font.Bold = $true
    $normativeSources=Get-Content -LiteralPath (Join-Path $PSScriptRoot 'NormativeHelpSources.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    foreach ($manualSource in $normativeSources) {
        $row++
        $InstructionSheet.Cells.Item($row, 1).Value2 = [string]$manualSource.caption
        $InstructionSheet.Cells.Item($row, 2).Value2 = [string]$manualSource.description
    }
    $row += 2
    $InstructionSheet.Cells.Item($row, 1).Value2 = "Словарь пользовательских статусов"
    $InstructionSheet.Cells.Item($row, 1).Font.Bold = $true
    $InstructionSheet.Cells.Item($row, 1).Interior.Color = 15921906
    $InstructionSheet.Cells.Item($row, 2).Interior.Color = 15921906
    $statusTableStart = $row
    $row++
    $InstructionSheet.Cells.Item($row, 1).Value2 = "OK"
    $InstructionSheet.Cells.Item($row, 2).Value2 = "Расчет этой проверки выполнен успешно. Для прямого НДС равновесие найдено внутри физической диаграммы материала; для прочности или трещин проверка прошла."
    $row++
    $InstructionSheet.Cells.Item($row, 1).Value2 = "FAIL"
    $InstructionSheet.Cells.Item($row, 2).Value2 = "Инженерная проверка не проходит: превышен физический предел, допустимая ширина трещины или напряжение по продольным трещинам. Для прямого НДС также означает найденное равновесие за физическими пределами диаграмм; причину поясняет комментарий."
    $row++
    $InstructionSheet.Cells.Item($row, 1).Value2 = "BaseFail"
    $InstructionSheet.Cells.Item($row, 2).Value2 = "Постоянная часть выбранного пути уже не проходит физический критерий при λ = 0; λ - множитель выбранных нагрузок. Это подтвержденный физический отказ, не просто несходимость. При допустимом текущем НДС ширина продолжает рассчитываться; в СП 63 недоступное уточнение дает ψ_s = 1 с предупреждением, если формула выполнена. Подтвержденное отсутствие нормальной трещины имеет статус NotCracked, а не BaseFail."
    $row++
    $InstructionSheet.Cells.Item($row, 1).Value2 = "NumFail"
    $InstructionSheet.Cells.Item($row, 2).Value2 = "Численный поиск выполнялся, но программа не смогла надежно найти равновесие или предельную точку. Это не доказательство разрушения; проверьте диагностику, нагрузку, сетку и параметры расчета."
    $row++
    $InstructionSheet.Cells.Item($row, 1).Value2 = "InputErr"
    $InstructionSheet.Cells.Item($row, 2).Value2 = "Результат не получен из-за ошибки исходных данных или настройки: обязательная пустая ячейка, неизвестный вариант, недопустимая геометрия, материал, единицы или строка сочетания. Ошибка могла быть обнаружена до запуска либо при подготовке конкретной проверки; место и причина указаны в комментарии."
    $row++
    $InstructionSheet.Cells.Item($row, 1).Value2 = "CalcErr"
    $InstructionSheet.Cells.Item($row, 2).Value2 = "Нарушен внутренний контракт расчета или произошла программная ошибка. Это не физический отказ и не обычная несходимость. Сохраните комментарий и подробную диагностику."
    $row++
    $InstructionSheet.Cells.Item($row, 1).Value2 = "N/A"
    $InstructionSheet.Cells.Item($row, 2).Value2 = "Проверка не запрошена, неприменима либо не выполнена из-за отсутствия обязательного предыдущего результата. В последнем случае причина указана в комментарии, а фактический отказ остается в статусе соответствующего НДС или поиска и в общей сводке. N/A само по себе не означает успешную проверку."
    $InstructionSheet.Range($InstructionSheet.Cells.Item($statusTableStart, 1), $InstructionSheet.Cells.Item($row, 2)).Borders.LineStyle = 1
    $InstructionSheet.Range($InstructionSheet.Cells.Item($statusTableStart, 1), $InstructionSheet.Cells.Item($row, 2)).Borders.Color = 14277081
    $row += 3

    $anchors = @{}
    foreach ($item in (Get-SettingsInstructionCatalog)) {
        $anchors[$item.Key] = $row
        $heading = $InstructionSheet.Range($InstructionSheet.Cells.Item($row, 1), $InstructionSheet.Cells.Item($row, 6))
        $heading.Merge() | Out-Null
        $heading.Value2 = $item.Title
        $heading.Font.Bold = $true
        $heading.Interior.Color = 15921906
        $heading.WrapText = $false
        $row++
        $lineNumber = 1
        foreach ($line in $item.Lines) {
            if ([string]::IsNullOrWhiteSpace([string]$line)) {
                continue # Пустые слоты каталога не создают строки между пояснениями настройки.
            }
            Set-InstructionParagraphNumber $InstructionSheet.Cells.Item($row, 1) $lineNumber
            $InstructionSheet.Cells.Item($row, 2).Value2 = $line
            $InstructionSheet.Cells.Item($row, 2).IndentLevel = 1
            if ($line.StartsWith("Формула")) {
                $formulaCell = $InstructionSheet.Cells.Item($row, 2)
                $formulaCell.Font.Name = "Arial"
                $formulaCell.Font.Size = 11
                $formulaCell.Font.Italic = $true
                $formulaCell.Interior.Color = 16448250
                $formulaCell.Font.Color = 3355443
            }
            $row++
            $lineNumber++
        }
        if ($item.ContainsKey("Tables")) {
            foreach ($table in $item.Tables) {
                $row++
                $InstructionSheet.Cells.Item($row, 2).Value2 = $table.Title
                $InstructionSheet.Cells.Item($row, 2).Font.Bold = $true
                $InstructionSheet.Cells.Item($row, 2).Interior.Color = 15921906
                $row++

                for ($c = 0; $c -lt $table.Headers.Count; $c++) {
                    $cell = $InstructionSheet.Cells.Item($row, 2 + $c)
                    $cell.Value2 = $table.Headers[$c]
                    $cell.Font.Bold = $true
                    $cell.Interior.Color = 14277081
                    $cell.WrapText = $true
                    $cell.HorizontalAlignment = -4131
                }
                $row++

                $tableStartRow = $row - 1
                foreach ($dataRow in $table.Rows) {
                    for ($c = 0; $c -lt $dataRow.Count; $c++) {
                        $cell = $InstructionSheet.Cells.Item($row, 2 + $c)
                        $cell.Value2 = $dataRow[$c]
                        $cell.WrapText = $true
                        $cell.VerticalAlignment = -4160
                    }
                    $row++
                }

                $tableEndColumn = 1 + $table.Headers.Count
                $tableRange = $InstructionSheet.Range($InstructionSheet.Cells.Item($tableStartRow, 2), $InstructionSheet.Cells.Item($row - 1, $tableEndColumn))
                $tableRange.Borders.LineStyle = 1
                $tableRange.Borders.Color = 14277081
                $row++
            }
        }
        $row++
    }
    Add-RelatedInstructionLinks $InstructionSheet $anchors
    $configSettingsBodyEndRow = $row - 1
    if ($configSettingsBodyEndRow -ge $configSettingsBodyStartRow) {
        try {
            $InstructionSheet.Range($InstructionSheet.Cells.Item($configSettingsBodyStartRow, 1), $InstructionSheet.Cells.Item($configSettingsBodyEndRow, 1)).EntireRow.Group() | Out-Null
        }
        catch {
            # Если Excel не дал создать outline-группу, текстовый раздел все равно
            # остается на листе и ссылки "Подробнее" продолжают вести в нужные строки.
        }
    }

    $row = Add-CollapsibleGuideSection $InstructionSheet $row "Справка по результатам расчета листа Results" {
        param($sheet, $startRow)
        Add-ResultsSheetGuide $sheet $startRow
    }

    $InstructionSheet.Range($InstructionSheet.Cells.Item(1, 1), $InstructionSheet.Cells.Item($row, 6)).VerticalAlignment = -4160
    $InstructionSheet.Range($InstructionSheet.Cells.Item(1, 1), $InstructionSheet.Cells.Item($row, 6)).Borders.LineStyle = 1
    $InstructionSheet.Range($InstructionSheet.Cells.Item(1, 1), $InstructionSheet.Cells.Item($row, 6)).Borders.Color = 14277081
    $InstructionSheet.Rows.AutoFit()
    foreach ($formulaRow in $script:FormulaImageRows) {
        $InstructionSheet.Rows.Item($formulaRow.Row).RowHeight = $formulaRow.Height
        # AutoFit может сместить/масштабировать xlMoveAndSize-картинки.
        # После установки окончательной высоты возвращаем их внутрь своей строки.
        $range = $InstructionSheet.Range($InstructionSheet.Cells.Item($formulaRow.Row, 1), $InstructionSheet.Cells.Item($formulaRow.Row, 6))
        $picture = $InstructionSheet.Shapes.Item($formulaRow.Name)
        $picture.Width = $formulaRow.PictureWidth
        $picture.Height = $formulaRow.PictureHeight
        $picture.Left = [double]$range.Left + ([double]$range.Width - $formulaRow.PictureWidth) / 2
        $picture.Top = [double]$range.Top + ($formulaRow.Height - $formulaRow.PictureHeight) / 2
    }
    try {
        # Методические разделы и справка по Results свернуты по умолчанию.
        # Раздел с описанием настроек Config оставляем раскрытым: именно туда
        # ведут ссылки "Подробнее" с листа Config.
        $InstructionSheet.Outline.ShowLevels(1, [Type]::Missing)
    }
    catch {
        try { $InstructionSheet.Outline.ShowLevels(1) } catch {}
    }
    if ($configSettingsBodyEndRow -ge $configSettingsBodyStartRow) {
        try {
            $InstructionSheet.Range($InstructionSheet.Cells.Item($configSettingsBodyStartRow, 1), $InstructionSheet.Cells.Item($configSettingsBodyEndRow, 1)).EntireRow.Hidden = $false
            if ($configSettingsBodyStartRow -gt 1) {
                try { $InstructionSheet.Rows.Item($configSettingsBodyStartRow - 1).ShowDetail = $true } catch {}
            }
        }
        catch {
            # Если Excel не дал раскрыть последнюю группу, ссылки "Подробнее"
            # все равно смогут раскрыть нужную строку через обработчик клика.
        }
    }
    try { $InstructionSheet.Columns.Hidden = $false } catch {}

    $used = $ConfigSheet.UsedRange
    $ConfigSheet.Cells.Item(3, 5).Value2 = "Справка"
    $ConfigSheet.Cells.Item(3, 5).Font.Bold = $true
    $ConfigSheet.Cells.Item(3, 5).Interior.Color = 14277081
    $ConfigSheet.Columns.Item(5).ColumnWidth = 11
    for ($r = 4; $r -le ($used.Row + $used.Rows.Count + 5); $r++) {
        $key = [string]$ConfigSheet.Cells.Item($r, 1).Value2
        if ($anchors.ContainsKey($key)) {
            Add-InstructionHyperlink $ConfigSheet.Cells.Item($r, 5) $InstructionSheet.Name $anchors[$key]
        }
    }
    Add-NamedRangeInstructionLinks $ConfigSheet $InstructionSheet.Name $anchors "rngSteelMaterialParameters" 1 6
    Add-NamedRangeInstructionLinks $ConfigSheet $InstructionSheet.Name $anchors "rngConcreteMaterialParameters" 1 6
    Add-NamedRangeInstructionLinks $ConfigSheet $InstructionSheet.Name $anchors "rngPlotAnnotationSettings" 1 6

    $headerLinks = @{
        "Единицы измерения" = "Units"
        "Система знаков" = "SignConvention"
        "Сочетания нагрузок" = "LoadCombinations"
        "Нагрузки для устойчивости" = "StabilityLoads"
        "СП 35.13330.2011, таблица 7.21" = "SP35Table721"
        "Параметры бетона" = "ConcreteMaterialParameters"
        "Параметры арматуры" = "SteelMaterialParameters"
        "Материал бетона" = "ConcreteMaterialParameters"
        "Материал арматуры" = "SteelMaterialParameters"
        "Настройка расчетных профилей" = "CalculationProfiles"
        "Контрольные точки диаграмм" = "MaterialDiagramControlTables"
        "Настройки оформления схемы" = "PlotAnnotationSettings"
        "Круглое сечение" = "CircleGeometry"
        "Клиновидное сечение" = "PointedSectionGeometry"
        "Пустотелый прямоугольник" = "HollowRectangleGeometry"
        "Набор прямоугольников (Г-сечение)" = "RectanglesGeometry"
    }
    for ($r = 1; $r -le ($used.Row + $used.Rows.Count + 40); $r++) {
        for ($c = 1; $c -le 40; $c++) {
            $caption = [string]$ConfigSheet.Cells.Item($r, $c).Value2
            # При обновлении готовой книги заголовок уже содержит пометку ссылки.
            # Сравниваем исходное название, чтобы назначить новый адрес справки
            # и не накапливать повторные "(Подробнее)" при следующем обновлении.
            if ($caption.EndsWith(" (Подробнее)")) {
                $caption = $caption.Substring(0, $caption.Length - " (Подробнее)".Length)
            }
            if ($headerLinks.ContainsKey($caption)) {
                $targetKey = $headerLinks[$caption]
                if ($anchors.ContainsKey($targetKey)) {
                    Add-HeaderInstructionHyperlink $ConfigSheet.Cells.Item($r, $c) $caption $InstructionSheet.Name $anchors[$targetKey]
                }
            }
        }
    }
    $sections=$Workbook.Worksheets.Item('Сечения')
    foreach($definition in @(Get-GeometrySettingsCatalog)){
        $range=$Workbook.Names.Item($definition.RangeName).RefersToRange
        $caption=[string]$definition.Title
        if($headerLinks.ContainsKey($caption)){
            Add-HeaderInstructionHyperlink $sections.Cells.Item($range.Row-1,$range.Column) $caption $InstructionSheet.Name $anchors[$headerLinks[$caption]]
        }
    }
    # Добавление Hyperlink применяет стиль Excel и сбрасывает рамки Справка.
    # Правило действует и при обновлении справки уже существующей книги.
    Apply-ConfigNamedRangeBorders $Workbook
}

# Связывает краткую справку типа с общими правилами без повторения одного текста.
# Адрес последнего пояснения определяется каталогом, а не номером строки листа.
function Add-RelatedInstructionLinks {
    param([object]$InstructionSheet, [hashtable]$Anchors)
    foreach ($item in (Get-SettingsInstructionCatalog)) {
        if (-not $item.ContainsKey('RelatedKey')) { continue }
        if (-not $Anchors.ContainsKey($item.Key) -or -not $Anchors.ContainsKey($item.RelatedKey)) {
            throw "Не найден раздел справки для ссылки $($item.Key)."
        }
        $cell = $InstructionSheet.Cells.Item(($Anchors[$item.Key] + $item.Lines.Count), 2)
        $cell.Hyperlinks.Delete()
        $InstructionSheet.Hyperlinks.Add($cell, '', "'" + $InstructionSheet.Name + "'!A" + $Anchors[$item.RelatedKey], '', [string]$item.Lines[-1]) | Out-Null
        $cell.Font.Color = 16711680
        $cell.Font.Underline = 2
    }
}

# Рисует таблицу сочетаний нагрузок на Config и сразу назначает ей имя
# rngLoadCombinations. Расчетный код ниже по цепочке не знает, на каком
# листе находится таблица: reader всегда работает через это имя.
function Add-LoadCombinationsTable {
    param(
        [object]$Workbook,
        [object]$Sheet,
        [int]$HeaderRow,
        [int]$StartColumn
    )

    $titleRow = $HeaderRow - 1
    $Sheet.Cells.Item($titleRow, $StartColumn).Value2 = "Сочетания нагрузок"
    $Sheet.Range($Sheet.Cells.Item($titleRow, $StartColumn), $Sheet.Cells.Item($titleRow, $StartColumn + 6)).Merge() | Out-Null
    $Sheet.Cells.Item($titleRow, $StartColumn).Font.Bold = $true
    $Sheet.Cells.Item($titleRow, $StartColumn).Interior.Color = 15921906

    $loadHeaders = @("CombinationID", "N", "Mx", "My", "ProfileId", "LoadPath", "Comment")
    for ($i = 0; $i -lt $loadHeaders.Count; $i++) {
        $cell = $Sheet.Cells.Item($HeaderRow, $StartColumn + $i)
        $cell.Value2 = $loadHeaders[$i]
        $cell.Font.Bold = $true
        $cell.Interior.Color = $script:ConfigTableHeaderColor
    }
    $Sheet.Cells.Item($HeaderRow, $StartColumn + 1).Formula = "=`"N, `"&INDEX(rngUnitSettings,MATCH(`"Force`",INDEX(rngUnitSettings,,1),0),2)"
    $Sheet.Cells.Item($HeaderRow, $StartColumn + 2).Formula = "=`"Mx, `"&INDEX(rngUnitSettings,MATCH(`"Moment`",INDEX(rngUnitSettings,,1),0),2)"
    $Sheet.Cells.Item($HeaderRow, $StartColumn + 3).Formula = "=`"My, `"&INDEX(rngUnitSettings,MATCH(`"Moment`",INDEX(rngUnitSettings,,1),0),2)"

    $loadRange = $Sheet.Range($Sheet.Cells.Item($HeaderRow, $StartColumn), $Sheet.Cells.Item($HeaderRow + 30, $StartColumn + 6))
    $loadRange.Borders.LineStyle = 1
    $loadRange.Borders.Weight = 2
    $loadRange.Borders.Color = 12632256

    $defaultProfileRow = Get-CalculationProfilesCatalog | Where-Object { $_.Key -eq "Profile.DisplayName" } | Select-Object -First 1
    $defaultProfileName = $defaultProfileRow.PR1
    $Sheet.Cells.Item($HeaderRow + 1, $StartColumn + 4).Value2 = $defaultProfileName

    $lambda = [char]0x03BB
    $capacityLoadPathOptions = @("Auto", "$lambda*Mx", "$lambda*My", "$lambda*Mxy", "$lambda*N", "$lambda*NMxy")
    $capacityLoadPathListAddress = Get-ConfigValidationListAddress $Sheet $capacityLoadPathOptions
    $capacityLoadPathRange = $Sheet.Range($Sheet.Cells.Item($HeaderRow + 1, $StartColumn + 5), $Sheet.Cells.Item($HeaderRow + 30, $StartColumn + 5))
    $capacityLoadPathRange.Validation.Delete()
    $capacityLoadPathRange.Validation.Add(3, 1, 1, $capacityLoadPathListAddress)
    $capacityLoadPathRange.Validation.IgnoreBlank = $true
    $capacityLoadPathRange.Validation.InCellDropdown = $true
    $capacityLoadPathRange.Value2 = "Auto"

    Set-WorkbookNameByBounds $Workbook "rngLoadCombinations" $Sheet $HeaderRow $StartColumn ($HeaderRow + 30) ($StartColumn + 6)
}

# Подключает выпадающий список профилей после создания rngCalculationProfiles.
# Ячейки сочетаний остаются обычным пользовательским вводом без формул; динамика
# сидит только в служебном списке, который ссылается на строку Profile.DisplayName.
function Add-LoadProfileValidation {
    param([object]$Workbook, [object]$Sheet)

    $loadRange = $Workbook.Names.Item("rngLoadCombinations").RefersToRange
    $profilesRange = $Workbook.Names.Item("rngCalculationProfiles").RefersToRange

    $displayRow = 0
    for ($r = 1; $r -le $profilesRange.Rows.Count; $r++) {
        if ([string]$profilesRange.Cells.Item($r, 2).Value2 -eq "Profile.DisplayName") {
            $displayRow = $r
            break
        }
    }
    if ($displayRow -eq 0) { return }

    $profileCount = [Math]::Max(0, $profilesRange.Columns.Count - 3)
    if ($profileCount -le 0) { return }

    $ListColumn = Get-NextConfigValidationColumn $Sheet 'LoadProfiles'
    for ($i = 0; $i -lt $profileCount; $i++) {
        $sourceCell = $profilesRange.Cells.Item($displayRow, 3 + $i)
        $Sheet.Cells.Item($i + 1, $ListColumn).Formula = "=" + $sourceCell.Address($true, $true)
    }
    $Sheet.Columns.Item($ListColumn).Hidden = $true

    $colName = ConvertTo-ExcelColumn $ListColumn
    $listAddress = "=$" + $colName + '$1:$' + $colName + '$' + $profileCount
    $profileInputRange = $Sheet.Range($loadRange.Cells.Item(2, 5), $loadRange.Cells.Item($loadRange.Rows.Count, 5))
    $profileInputRange.Validation.Delete()
    $profileInputRange.Validation.Add(3, 1, 1, $listAddress)
    $profileInputRange.Validation.IgnoreBlank = $false
    $profileInputRange.Validation.InCellDropdown = $true
}

# Рисует таблицу постоянных/длительных нагрузок для расчета устойчивости.
# Она связана с основной таблицей только через CombinationID: Excel-слой
# прочитает значения, переведет единицы и передаст массив в batch.
function Add-StabilityDurationLoadsTable {
    param(
        [object]$Workbook,
        [object]$Sheet,
        [int]$HeaderRow,
        [int]$StartColumn
    )

    $titleRow = $HeaderRow - 1
    $Sheet.Cells.Item($titleRow, $StartColumn).Value2 = "Нагрузки для устойчивости"
    $Sheet.Range($Sheet.Cells.Item($titleRow, $StartColumn), $Sheet.Cells.Item($titleRow, $StartColumn + 3)).Merge() | Out-Null
    $Sheet.Cells.Item($titleRow, $StartColumn).Font.Bold = $true
    $Sheet.Cells.Item($titleRow, $StartColumn).Interior.Color = 15921906

    $headers = @("LC", "N", "Mx", "My")
    for ($i = 0; $i -lt $headers.Count; $i++) {
        $cell = $Sheet.Cells.Item($HeaderRow, $StartColumn + $i)
        $cell.Value2 = $headers[$i]
        $cell.Font.Bold = $true
        $cell.Interior.Color = $script:ConfigTableHeaderColor
    }
    $Sheet.Cells.Item($HeaderRow, $StartColumn + 1).Formula = "=`"N, `"&INDEX(rngUnitSettings,MATCH(`"Force`",INDEX(rngUnitSettings,,1),0),2)"
    $Sheet.Cells.Item($HeaderRow, $StartColumn + 2).Formula = "=`"Mx, `"&INDEX(rngUnitSettings,MATCH(`"Moment`",INDEX(rngUnitSettings,,1),0),2)"
    $Sheet.Cells.Item($HeaderRow, $StartColumn + 3).Formula = "=`"My, `"&INDEX(rngUnitSettings,MATCH(`"Moment`",INDEX(rngUnitSettings,,1),0),2)"

    for ($r = 1; $r -le 30; $r++) {
        $Sheet.Cells.Item($HeaderRow + $r, $StartColumn).Formula = '=IF(INDEX(rngLoadCombinations,' + ($r + 1) + ',1)="","",INDEX(rngLoadCombinations,' + ($r + 1) + ',1))'
        $Sheet.Cells.Item($HeaderRow + $r, $StartColumn + 1).Value2 = 0
        $Sheet.Cells.Item($HeaderRow + $r, $StartColumn + 2).Value2 = 0
        $Sheet.Cells.Item($HeaderRow + $r, $StartColumn + 3).Value2 = 0
    }

    Set-WorkbookNameByBounds $Workbook "rngStabilityDurationLoads" $Sheet $HeaderRow $StartColumn ($HeaderRow + 30) ($StartColumn + 3)
}

# Размещает на Config таблицу 7.21 СП 35. Расчетный код читает ее как
# исходный массив; встроенных hardcoded-значений таблицы в VBA быть не должно.
function Add-SP35Table721 {
    param(
        [object]$Workbook,
        [object]$Sheet,
        [int]$HeaderRow,
        [int]$StartColumn
    )

    $Sheet.Cells.Item($HeaderRow - 2, $StartColumn).Value2 = "СП 35.13330.2011, таблица 7.21"
    $Sheet.Range($Sheet.Cells.Item($HeaderRow - 2, $StartColumn), $Sheet.Cells.Item($HeaderRow - 2, $StartColumn + 7)).Merge() | Out-Null
    $Sheet.Cells.Item($HeaderRow - 2, $StartColumn).Font.Bold = $true
    $Sheet.Cells.Item($HeaderRow - 2, $StartColumn).Interior.Color = 15921906
    $Sheet.Cells.Item($HeaderRow - 1, $StartColumn).Value2 = "Верхние значения для ненапрягаемой арматуры; последняя строка пользовательская для очень большой гибкости."
    $Sheet.Range($Sheet.Cells.Item($HeaderRow - 1, $StartColumn), $Sheet.Cells.Item($HeaderRow - 1, $StartColumn + 7)).Merge() | Out-Null

    $headers = @("l_0/b", "l_0/d", "l_0/i", "φ_m q=0", "φ_m q=0.25", "φ_m q=0.50", "φ_m q=1.00", "φ_l")
    for ($i = 0; $i -lt $headers.Count; $i++) {
        $cell = $Sheet.Cells.Item($HeaderRow, $StartColumn + $i)
        $cell.Value2 = $headers[$i]
        $cell.Font.Bold = $true
        $cell.Interior.Color = $script:ConfigTableHeaderColor
    }

    $rows = @(
        @(4, 3.5, 14, 1.00, 0.90, 0.81, 0.69, 1.00),
        @(10, 8.7, 35, 1.00, 0.86, 0.77, 0.65, 0.84),
        @(12, 10.4, 40, 0.95, 0.83, 0.74, 0.62, 0.79),
        @(14, 12.1, 48.5, 0.90, 0.79, 0.70, 0.58, 0.70),
        @(16, 13.8, 55, 0.86, 0.75, 0.66, 0.55, 0.65),
        @(18, 15.6, 62.5, 0.82, 0.71, 0.62, 0.51, 0.56),
        @(20, 17.3, 70, 0.78, 0.67, 0.57, 0.48, 0.47),
        @(22, 19.1, 75, 0.72, 0.60, 0.52, 0.43, 0.41),
        @(24, 20.8, 83, 0.67, 0.55, 0.47, 0.38, 0.32),
        @(26, 22.5, 90, 0.62, 0.51, 0.44, 0.35, 0.25),
        @(28, 24.3, 97, 0.58, 0.49, 0.43, 0.34, 0.20),
        @(30, 26.0, 105, 0.53, 0.45, 0.39, 0.32, 0.16),
        @(32, 27.7, 110, 0.48, 0.41, 0.36, 0.31, 0.14),
        @(34, 29.0, 120, 0.43, 0.36, 0.31, 0.25, 0.10),
        @(38, 33.0, 130, 0.38, 0.32, 0.28, 0.24, 0.08),
        @(40, 34.6, 140, 0.35, 0.29, 0.25, 0.21, 0.07),
        @(43, 37.5, 150, 0.33, 0.28, 0.24, 0.21, 0.06),
        @(1000000, 1000000, 1000000, 0.33, 0.28, 0.24, 0.21, 0.06)
    )

    for ($r = 0; $r -lt $rows.Count; $r++) {
        for ($c = 0; $c -lt $headers.Count; $c++) {
            $rowIndex = [int]($HeaderRow + 1 + $r)
            $columnIndex = [int]($StartColumn + $c)
            $Sheet.Cells.Item($rowIndex, $columnIndex).Value2 = [double]$rows[$r][$c]
        }
    }

    Set-WorkbookNameByBounds $Workbook "rngSP35Table721" $Sheet $HeaderRow $StartColumn ($HeaderRow + $rows.Count) ($StartColumn + 7)
}

# Расставляет ссылки "Подробнее" внутри табличных блоков, у которых есть
# собственная колонка справки. Явная привязка к именованному диапазону не
# зависит от того, куда блок перенесли на листе Config.
function Add-NamedRangeInstructionLinks {
    param(
        [object]$ConfigSheet,
        [string]$InstructionSheetName,
        [hashtable]$Anchors,
        [string]$RangeName,
        [int]$KeyColumn,
        [int]$InstructionColumn
    )

    try {
        $range = $ConfigSheet.Parent.Names.Item($RangeName).RefersToRange
    } catch {
        return
    }

    for ($r = 2; $r -le $range.Rows.Count; $r++) {
        $key = [string]$range.Cells.Item($r, $KeyColumn).Value2
        if ($Anchors.ContainsKey($key)) {
            Add-InstructionHyperlink $range.Cells.Item($r, $InstructionColumn) $InstructionSheetName $Anchors[$key]
        }
    }
}

# Читает только значения свойств COM-формата, не сохраняя живые ссылки на Font/Interior.
# Theme-свойства для обычного RGB могут отсутствовать; Null обозначает смешанный шрифт.
function Get-InstructionFormatProperties {
    param([object]$Target, [string[]]$Names)
    $values = [ordered]@{}
    foreach ($name in $Names) {
        try { $values[$name] = $Target.$name } catch {}
    }
    return $values
}

# Восстанавливает сохранённые атрибуты в явном порядке: тема после RGB,
# LineStyle/Pattern последними, чтобы заливка «нет» и невидимая рамка не включились.
function Set-InstructionFormatProperties {
    param([object]$Target, [System.Collections.IDictionary]$Values, [string[]]$Names)
    foreach ($name in $Names) {
        if (-not $Values.Contains($name) -or $null -eq $Values[$name] -or $Values[$name] -is [DBNull]) { continue }
        $indexName = if ($name -eq 'Color') { 'ColorIndex' } elseif ($name -eq 'PatternColor') { 'PatternColorIndex' } else { '' }
        $themeName = if ($name -eq 'Color') { 'ThemeColor' } else { 'PatternThemeColor' }
        $hasTheme = $Values.Contains($themeName) -and $null -ne $Values[$themeName] -and $Values[$themeName] -isnot [DBNull] -and [int]$Values[$themeName] -gt 0
        if ($indexName -and $Values.Contains($indexName) -and [int]$Values[$indexName] -lt 0 -and -not $hasTheme) { continue }
        if ($name -in @('ThemeColor','PatternThemeColor') -and [int]$Values[$name] -lt 1) { continue }
        $propertyTarget = if ($name -in @('Locked','FormulaHidden')) { $Target.MergeArea } else { $Target } # Excel запрещает менять защиту отдельной части объединённой ячейки.
        try {
            $actual = $propertyTarget.$name
            if ($null -ne $actual -and $actual -isnot [DBNull] -and $actual -eq $Values[$name]) { continue } # Неизменённые грани не переписываем: Excel делит их с соседями; лишние setter-ы также включают apply* для унаследованных значений атрибутов.
        } catch {}
        $propertyTarget.$name = $Values[$name]
    }
}

# Сохраняет фактическое оформление всех ячеек MergeArea до изменения Hyperlink.
# Rich text копируется по сериям только при смешанном шрифте; листы, стили и clipboard не создаются.
function Save-InstructionCellFormat {
    param([object]$Cell)
    $fontNames = @('Name','Size','Bold','Italic','Strikethrough','Superscript','Subscript','OutlineFont','Shadow','ThemeFont','ColorIndex','Color','ThemeColor','TintAndShade','Underline')
    $cellNames = @('NumberFormat','HorizontalAlignment','VerticalAlignment','WrapText','ShrinkToFit','Orientation','AddIndent','IndentLevel','ReadingOrder','Locked','FormulaHidden')
    $fillNames = @('ColorIndex','Color','ThemeColor','TintAndShade','PatternColorIndex','PatternColor','PatternThemeColor','PatternTintAndShade','Pattern')
    $borderNames = @('ColorIndex','Color','ThemeColor','TintAndShade','Weight','LineStyle')
    foreach ($part in $Cell.MergeArea.Cells) {
        $style = $part.Style
        $styleName = if ($style -is [string]) { [string]$style } else { [string]$style.NameLocal }
        $font = Get-InstructionFormatProperties $part.Font $fontNames
        $fill = Get-InstructionFormatProperties $part.Interior $fillNames
        if ([int]$fill['Pattern'] -in @(4000,4001)) { throw 'Градиентная заливка ссылки требует отдельного сохранения; гиперссылка не изменена.' }
        $record = [pscustomobject]@{
            Row=[int]$part.Row; Column=[int]$part.Column; Style=$styleName
            Text=[string]$part.Value2; Font=$font; Interior=$fill
            Cell=(Get-InstructionFormatProperties $part $cellNames); Borders=@(); Runs=@()
        }
        foreach ($index in @(5,6,7,8,9,10,11,12)) {
            $record.Borders += [pscustomobject]@{Index=$index; Values=(Get-InstructionFormatProperties $part.Borders.Item($index) $borderNames)}
        }
        $mixed = @($font.Values | Where-Object {$null -eq $_ -or $_ -is [DBNull]}).Count -gt 0
        if ($mixed -and $record.Text.Length -gt 0) {
            $record.Font = Get-InstructionFormatProperties $part.Characters(1,1).Font $fontNames
            $lastKey = ''; $lastRun = $null
            for ($index = 1; $index -le $record.Text.Length; $index++) {
                $characterFont = Get-InstructionFormatProperties $part.Characters($index,1).Font $fontNames
                $key = ConvertTo-Json -InputObject $characterFont -Compress
                if ($lastRun -and $key -eq $lastKey) {
                    $lastRun.Length++
                } else {
                    $lastRun = [pscustomobject]@{Start=$index; Length=1; Font=$characterFont}
                    $record.Runs += $lastRun; $lastKey = $key
                }
            }
        }
        $record
    }
}

# Возвращает фактический формат перед намеренной ссылочной раскраской.
# Серии rich text восстанавливаются только для неизменившегося текстового префикса.
# Имя Style служит диагностикой: его назначение части MergeArea запрещено Excel;
# полное оформление восстанавливается сохранёнными фактическими атрибутами ниже.
function Restore-InstructionCellFormat {
    param([object]$Cell, [object[]]$Records)
    $fontNames = @('Name','Size','Bold','Italic','Strikethrough','Superscript','Subscript','OutlineFont','Shadow','ThemeFont','ColorIndex','Color','ThemeColor','TintAndShade','Underline')
    $cellNames = @('NumberFormat','HorizontalAlignment','VerticalAlignment','WrapText','ShrinkToFit','Orientation','AddIndent','IndentLevel','ReadingOrder','Locked','FormulaHidden')
    $fillNames = @('ColorIndex','Color','ThemeColor','TintAndShade','PatternColorIndex','PatternColor','PatternThemeColor','PatternTintAndShade','Pattern')
    $borderNames = @('ColorIndex','Color','ThemeColor','TintAndShade','Weight','LineStyle')
    foreach ($record in $Records) {
        $part = $Cell.Parent.Cells.Item($record.Row,$record.Column)
        Set-InstructionFormatProperties $part $record.Cell $cellNames
        Set-InstructionFormatProperties $part.Interior $record.Interior $fillNames
        Set-InstructionFormatProperties $part.Font $record.Font $fontNames
        foreach ($border in $record.Borders) {
            $borderTarget = $part.Borders.Item($border.Index)
            $actualBorder = Get-InstructionFormatProperties $borderTarget $borderNames
            if (-not $border.Values.Contains('LineStyle') -or $null -eq $border.Values['LineStyle'] -or $border.Values['LineStyle'] -is [DBNull]) { continue }
            if ([int]$border.Values['LineStyle'] -eq -4142) {
                if (-not $actualBorder.Contains('LineStyle') -or $null -eq $actualBorder['LineStyle'] -or $actualBorder['LineStyle'] -is [DBNull] -or [int]$actualBorder['LineStyle'] -ne -4142) { $borderTarget.LineStyle = -4142 }
                continue # У невидимой грани Color/ColorIndex/Weight не влияют на вид; их setter-ы способны создать линию и перераспределить рамку соседям.
            }
            $sameBorder = $true
            foreach ($name in @('LineStyle','Weight','Color','TintAndShade')) {
                if ($border.Values.Contains($name) -and $null -ne $border.Values[$name] -and $border.Values[$name] -isnot [DBNull]) {
                    if (-not $actualBorder.Contains($name) -or $null -eq $actualBorder[$name] -or $actualBorder[$name] -is [DBNull] -or $actualBorder[$name] -ne $border.Values[$name]) { $sameBorder = $false }
                }
            }
            $savedTheme = if ($border.Values.Contains('ThemeColor') -and $null -ne $border.Values['ThemeColor'] -and $border.Values['ThemeColor'] -isnot [DBNull]) { [int]$border.Values['ThemeColor'] } else { 0 }
            $actualTheme = if ($actualBorder.Contains('ThemeColor') -and $null -ne $actualBorder['ThemeColor'] -and $actualBorder['ThemeColor'] -isnot [DBNull]) { [int]$actualBorder['ThemeColor'] } else { 0 }
            if (($savedTheme -gt 0 -or $actualTheme -gt 0) -and $savedTheme -ne $actualTheme) { $sameBorder = $false }
            if ($sameBorder) { continue } # ColorIndex не различает исходный RGB и эквивалентный palette black; одинаковую видимую грань не переписываем.
            Set-InstructionFormatProperties $borderTarget $border.Values $borderNames
        }
        $text = [string]$part.Value2; $common = 0
        while ($common -lt $record.Text.Length -and $common -lt $text.Length -and $record.Text[$common] -ceq $text[$common]) { $common++ }
        foreach ($run in $record.Runs) {
            $length = [Math]::Min([int]$run.Length, $common - [int]$run.Start + 1)
            if ($length -gt 0) { Set-InstructionFormatProperties $part.Characters($run.Start,$length).Font $run.Font $fontNames }
        }
    }
}

# Создает прямую ссылку на ячейку листа "Справка" без использования именованных диапазонов.
function Add-InstructionHyperlink {
    param([object]$Cell, [string]$SheetName, [int]$TargetRow)
    $format = @(Save-InstructionCellFormat $Cell)
    try {
        $Cell.Hyperlinks.Delete()
        $Cell.Value2 = "Подробнее"
        $Cell.Parent.Hyperlinks.Add($Cell, "", "'" + $SheetName + "'!A" + $TargetRow, "", "Подробнее") | Out-Null
    } finally {
        Restore-InstructionCellFormat $Cell $format
    }
    Set-InstructionHyperlinkFont $Cell
}

# Задаёт одинаковый вид обычных ссылок «Подробнее», включая уже посещённые.
# Шрифт применяется после создания ссылки: Excel иначе подставляет стиль Hyperlink.
function Set-InstructionHyperlinkFont {
    param([object]$Cell)
    $Cell.Font.Name = 'Arial'; $Cell.Font.Size = 9
    $Cell.Font.Bold = $false; $Cell.Font.Italic = $false
    $Cell.Font.Color = 16711680; $Cell.Font.Underline = 2
}

# Нумерует пояснение внутри блока справки; тип числа и явное выравнивание
# устраняют разницу между текстовыми и числовыми номерами.
function Set-InstructionParagraphNumber {
    param([object]$Cell, [int]$Number)
    $Cell.Value2 = [double]$Number
    $Cell.NumberFormat = '0'; $Cell.HorizontalAlignment = -4152
    $Cell.Font.Name = 'Arial'; $Cell.Font.Size = 10
    $Cell.Font.Bold = $false; $Cell.Font.Italic = $false
    $Cell.Font.Color = 0; $Cell.Font.Underline = -4142
    $Cell.IndentLevel = 0
}

# Подбирает высоту обычных пояснений после восстановления сохранённого макета.
# Тексты каталога могут стать длиннее; заголовки остаются однострочными.
# Только сборка: расчёт и обновление схемы эту функцию не вызывают.
function Fit-SettingsInstructionRows {
    param([object]$Sheet)
    foreach ($item in (Get-SettingsInstructionCatalog)) {
        $heading = $Sheet.Cells.Find($item.Title, $Sheet.Cells.Item(1, 1), -4163, 1, 1, 1, $false, $false, $false)
        if ($null -eq $heading) { throw ('Не найден заголовок справки: ' + $item.Title) }
        $count = @($item.Lines | Where-Object {-not [string]::IsNullOrWhiteSpace([string]$_)}).Count
        $heading.EntireRow.RowHeight = [double]$Sheet.StandardHeight
        if ($count -gt 0) {
            $Sheet.Range($Sheet.Cells.Item($heading.Row + 1, 1), $Sheet.Cells.Item($heading.Row + $count, 6)).EntireRow.AutoFit() | Out-Null
        }
    }
}

# Делает заголовок блока самодостаточным: название остается на месте, а пометка "(Подробнее)"
# показывает пользователю, что по заголовку можно перейти к справке.
function Add-HeaderInstructionHyperlink {
    param([object]$Cell, [string]$Caption, [string]$SheetName, [int]$TargetRow)
    $displayText = $Caption + " (Подробнее)"
    $format = @(Save-InstructionCellFormat $Cell)
    try {
        $Cell.Hyperlinks.Delete()
        $Cell.Value2 = $displayText
        $Cell.Parent.Hyperlinks.Add($Cell, "", "'" + $SheetName + "'!A" + $TargetRow, "", $displayText) | Out-Null
    } finally {
        Restore-InstructionCellFormat $Cell $format
    }
    try {
        $linkText = "(Подробнее)"
        $linkStart = $displayText.IndexOf($linkText) + 1
        $Cell.Characters(1, $Caption.Length).Font.Color = 0
        $Cell.Characters(1, $Caption.Length).Font.Underline = -4142
        $Cell.Characters($linkStart, $linkText.Length).Font.Color = 16711680
        $Cell.Characters($linkStart, $linkText.Length).Font.Underline = 2
    } catch {
        # Если Excel не даст частично отформатировать merged-cell, сама ссылка все равно останется рабочей.
    }
}

# Строит таблицу INPUT/INTERNAL/OUTPUT, задаёт допустимые единицы отдельно для ввода и вывода и регистрирует rngUnitSettings.
function Add-UnitSettingsTable {
    param([object]$Workbook, [object]$Sheet, [int]$HeaderRow, [int]$StartColumn)

    $Sheet.Cells.Item($HeaderRow - 1, $StartColumn).Value2 = "Единицы измерения"
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
    for ($r = 0; $r -lt $rows.Count; $r++) {
        $quantity = [string]$rows[$r][0]
        $options = $validationLists[$quantity]
        $listAddress = Get-ConfigValidationListAddress $Sheet $options
        foreach ($valueColumn in @($($StartColumn + 1), $($StartColumn + 3))) {
            $cell = $Sheet.Cells.Item($HeaderRow + 1 + $r, $valueColumn)
            $cell.Validation.Delete()
            $cell.Validation.Add(3, 1, 1, $listAddress)
            $cell.Validation.IgnoreBlank = $false
            $cell.Validation.InCellDropdown = $true
        }
    }

    Set-WorkbookNameByBounds $Workbook "rngUnitSettings" $Sheet $HeaderRow $StartColumn ($HeaderRow + $rows.Count) ($StartColumn + 3)
}

# Строит таблицу USER/INTERNAL, разрешает выбор пользовательских знаков и регистрирует rngSignConventionSettings.
function Add-SignConventionSettingsTable {
    param([object]$Workbook, [object]$Sheet, [int]$HeaderRow, [int]$StartColumn)

    $Sheet.Cells.Item($HeaderRow - 1, $StartColumn).Value2 = "Система знаков"
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
    for ($r = 0; $r -lt $rows.Count; $r++) {
        $quantity = [string]$rows[$r][0]
        $options = $validationLists[$quantity]
        $listAddress = Get-ConfigValidationListAddress $Sheet $options
        $cell = $Sheet.Cells.Item($HeaderRow + 1 + $r, $StartColumn + 1)
        $cell.Validation.Delete()
        $cell.Validation.Add(3, 1, 1, $listAddress)
        $cell.Validation.IgnoreBlank = $false
        $cell.Validation.InCellDropdown = $true
    }

    Set-WorkbookNameByBounds $Workbook "rngSignConventionSettings" $Sheet $HeaderRow $StartColumn ($HeaderRow + $rows.Count) ($StartColumn + 2)
}

# Задаёт независимые значения оформления арматурных подписей и размерных линий, их единицы и пояснения.
function Get-PlotAnnotationSettingsCatalog {
    @(
        @("Enabled", "Yes", "Yes", "Yes/No", "Обязательный выбор Yes/No для соответствующего типа аннотаций. При No параметры отключенной группы не читаются."),
        @("Placement", "Outside", "Outside", "-", "Для арматуры задает сторону подписи от бетонной границы: Outside - наружу, Inside - внутрь бетона; для размеров задает только сторону текста относительно размерной линии."),
        @("Offset", "30", "100", "мм", "Базовый отступ, мм. Для размеров и арматуры умножается на max(1; D/1000), D — наибольший габарит сечения в мм. У арматуры отсчёт от бетона; Placement — сторона, TextGap — зазор до текста."),
        @("LineEnabled", "Yes", "-", "Yes/No", "Показывать короткую линию обозначения арматуры; если No, остается только текст подписи."),
        @("LineWeight", "1.35", "1.75", "pt", "Толщина основной линии аннотации."),
        @("ExtensionLineWeight", "-", "0.85", "pt", "Толщина выносных линий; применяется только для размерных линий."),
        @("TextUnits", "pt", "pt", "-", "Единицы для TextHeight и TextGap: pt - обычные пункты Excel, mm - реальные миллиметры сечения."),
        @("TextHeight", "13", "13", "мм", "Высота текста. При TextUnits=mm задается в миллиметрах сечения; при TextUnits=pt - в пунктах Excel и не зависит от габарита сечения."),
        @("TextGap", "9", "9", "мм", "Зазор между линией аннотации и текстом. Единицы выбираются строкой TextUnits для соответствующего типа аннотаций."),
        @("ArrowType", "-", "Triangle", "-", "Тип стрелки размерной линии."),
        @("ArrowSize", "-", "Medium", "-", "Размер стрелок размерной линии."),
        @("Color", "20,30,90", "20,30,90", "RGB", "Цвет текста и основной линии: три целых RGB-компонента 0–255. Значение 0,0,0 означает черный цвет."),
        @("ExtensionLineColor", "-", "140,140,140", "RGB", "Цвет выносных линий размеров: три целых RGB-компонента 0–255. Значение 0,0,0 означает черный цвет."),
        @("ContourLineWeight", "2", "", "pt", "Общая толщина наружного контура и границ отверстий на Excel-схеме. Прямые, дуги и окружности используют одно значение > 0; по умолчанию 2 pt.")
    )
}

# Выводит шестиколоночную таблицу аннотаций, хранит RGB как текст и связывает подписи единиц с TextUnits обеих групп.
function Add-PlotAnnotationSettingsTable {
    param([object]$Workbook, [object]$Sheet, [int]$HeaderRow, [int]$StartColumn)

    $title = "Настройки оформления схемы"
    $Sheet.Cells.Item($HeaderRow - 1, $StartColumn).Value2 = $title
    $Sheet.Range($Sheet.Cells.Item($HeaderRow - 1, $StartColumn), $Sheet.Cells.Item($HeaderRow - 1, $StartColumn + 5)).Merge() | Out-Null
    $Sheet.Cells.Item($HeaderRow - 1, $StartColumn).Font.Bold = $true
    $Sheet.Cells.Item($HeaderRow - 1, $StartColumn).Interior.Color = 15921906

    $headers = @("Параметр", "Обозначения арматуры", "Размерные линии", "Ед.", "Комментарий", "Справка")
    for ($i = 0; $i -lt $headers.Count; $i++) {
        $cell = $Sheet.Cells.Item($HeaderRow, $StartColumn + $i)
        $cell.Value2 = $headers[$i]
        $cell.Font.Bold = $true
        $cell.Interior.Color = 14277081
    }

    $rows = Get-PlotAnnotationSettingsCatalog
    for ($r = 0; $r -lt $rows.Count; $r++) {
        for ($c = 0; $c -lt 5; $c++) {
            $cell = $Sheet.Cells.Item($HeaderRow + 1 + $r, $StartColumn + $c)
            if ([string]$rows[$r][3] -eq "RGB" -and ($c -eq 1 -or $c -eq 2)) {
                $cell.NumberFormat = "@"
            }
            $cell.Value2 = $rows[$r][$c]
        }
        if ([string]$rows[$r][0] -eq "ContourLineWeight") {
            $contourInput = $Sheet.Cells.Item($HeaderRow + 1 + $r, $StartColumn + 1).Resize(1, 2)
            $contourInput.Merge() | Out-Null
            $contourInput.HorizontalAlignment = -4108
            $contourInput.Validation.Add(2, 1, 5, 0)
            $contourInput.Validation.IgnoreBlank = $false
            $contourInput.Validation.ShowError = $true
        }
    }

    $textUnitsExcelRow = 0
    for ($r = 0; $r -lt $rows.Count; $r++) {
        if ([string]$rows[$r][0] -eq "TextUnits") {
            $textUnitsExcelRow = $HeaderRow + 1 + $r
        }
    }
    if ($textUnitsExcelRow -gt 0) {
        $rebarUnitsCell = $Sheet.Cells.Item($textUnitsExcelRow, $StartColumn + 1).Address($false, $false)
        $dimensionUnitsCell = $Sheet.Cells.Item($textUnitsExcelRow, $StartColumn + 2).Address($false, $false)
        $unitFormula = '=IF(' + $rebarUnitsCell + '=' + $dimensionUnitsCell + ',IF(' + $rebarUnitsCell + '="pt","pt","мм"),"арм.="&IF(' + $rebarUnitsCell + '="pt","pt","мм")&"; разм.="&IF(' + $dimensionUnitsCell + '="pt","pt","мм"))'
        for ($r = 0; $r -lt $rows.Count; $r++) {
            $settingName = [string]$rows[$r][0]
            if ($settingName -eq "TextHeight" -or $settingName -eq "TextGap") {
                $Sheet.Cells.Item($HeaderRow + 1 + $r, $StartColumn + 3).Formula = $unitFormula
            }
        }
    }

    Set-WorkbookNameByBounds $Workbook "rngPlotAnnotationSettings" $Sheet $HeaderRow $StartColumn ($HeaderRow + $rows.Count) ($StartColumn + 5)
}

# Возвращает формулу подписи текущей входной единицы из rngUnitSettings; неизвестное обозначение оставляет без формулы.
function Get-InputUnitFormulaForUnit {
    param([string]$UnitText)
    switch ($UnitText) {
        "мм" { return '=INDEX(rngUnitSettings,MATCH("Length",INDEX(rngUnitSettings,,1),0),2)' }
        "мм2" { return '=INDEX(rngUnitSettings,MATCH("Area",INDEX(rngUnitSettings,,1),0),2)' }
        "Н" { return '=INDEX(rngUnitSettings,MATCH("Force",INDEX(rngUnitSettings,,1),0),2)' }
        "Н*мм" { return '=INDEX(rngUnitSettings,MATCH("Moment",INDEX(rngUnitSettings,,1),0),2)' }
        "Н*мм/мм" { return '=INDEX(rngUnitSettings,MATCH("Moment",INDEX(rngUnitSettings,,1),0),2)&"/"&INDEX(rngUnitSettings,MATCH("Length",INDEX(rngUnitSettings,,1),0),2)' }
        "МПа" { return '=INDEX(rngUnitSettings,MATCH("Stress",INDEX(rngUnitSettings,,1),0),2)' }
        "1/мм" { return '=INDEX(rngUnitSettings,MATCH("Curvature",INDEX(rngUnitSettings,,1),0),2)' }
        default { return $null }
    }
}

# Записывает формулу текущей входной единицы для поддержанного обозначения; остальные подписи сохраняет постоянным текстом.
function Set-InputUnitCell {
    param([object]$Cell, [string]$UnitText)
    $formula = Get-InputUnitFormulaForUnit $UnitText
    if ($null -ne $formula) {
        $Cell.Formula = $formula
    } else {
        $Cell.Value2 = $UnitText
    }
}

# Записывает число отдельным типизированным вызовом COM: PowerShell 5.1
# иначе может повторно использовать строковый setter после заполнения подписей.
function Set-ConfigNumericValue {
    param([object]$Cell, [double]$Value)
    $Cell.Value2 = $Value
}

# Восстанавливает только служебную колонку единиц существующей таблицы по
# действующему каталогу. Значения, комментарии и validation пользователя
# не назначает заново; размерные единицы сохраняют связь с INPUT-настройками.
# Возвращает измененные подписи для отчета явного обновления книги.
function Update-SystemSettingUnitCaptions {
    param([object]$Workbook)
    $unitsByKey = @{}
    foreach ($section in @(Get-SystemSettingsCatalog)) {
        foreach ($row in $section.Rows) {
            $key = [string]$row[0]
            if (-not $key.StartsWith('[')) { $unitsByKey[$key] = [string]$row[2] }
        }
    }
    $settings = $Workbook.Names.Item('rngSystemSettings').RefersToRange
    if ($settings.Columns.Count -lt 3 -or [string]$settings.Cells.Item(1, 3).Value2 -ne 'Ед.') {
        throw 'rngSystemSettings должен содержать текущую колонку «Ед.» после значения. Исправьте структуру таблицы Config перед обновлением справки.'
    }
    $values = $settings.Value2
    for ($row = 2; $row -le $settings.Rows.Count; $row++) {
        $key = [string]$values[$row, 1]
        if ($unitsByKey.ContainsKey($key)) {
            $cell = $settings.Cells.Item($row, 3)
            $before = [string]$cell.Formula
            Set-InputUnitCell $cell $unitsByKey[$key]
            $after = [string]$cell.Formula
            if ($before -ne $after) {
                @{ key=$key; address=([string]$settings.Worksheet.Name+'!'+$cell.Address($false, $false)); before=$before; after=$after }
            }
        }
    }
}

# Рисует симметричную таблицу параметров материала: отдельные значения для
# сжатия и растяжения, но один именованный диапазон и один human-readable блок.
function Add-MaterialParameterTable {
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
    $Sheet.Range($Sheet.Cells.Item($HeaderRow - 1, $StartColumn), $Sheet.Cells.Item($HeaderRow - 1, $StartColumn + 5)).Merge() | Out-Null
    $Sheet.Cells.Item($HeaderRow - 1, $StartColumn).Font.Bold = $true
    $Sheet.Cells.Item($HeaderRow - 1, $StartColumn).Interior.Color = 15921906

    $headers = @("Параметр", "Сжатие", "Растяжение", "Ед.", "Комментарий", "Справка")
    for ($i = 0; $i -lt $headers.Count; $i++) {
        $cell = $Sheet.Cells.Item($HeaderRow, $StartColumn + $i)
        $cell.Value2 = $headers[$i]
        $cell.Font.Bold = $true
        $cell.Interior.Color = 14277081
    }

    for ($r = 0; $r -lt $Rows.Count; $r++) {
        for ($c = 0; $c -lt 5; $c++) {
            $Sheet.Cells.Item($HeaderRow + 1 + $r, $StartColumn + $c).Value2 = $Rows[$r][$c]
        }
        Set-InputUnitCell $Sheet.Cells.Item($HeaderRow + 1 + $r, $StartColumn + 3) ([string]$Rows[$r][3])
    }

    Set-WorkbookNameByBounds $Workbook $RangeName $Sheet $HeaderRow $StartColumn ($HeaderRow + $Rows.Count) ($StartColumn + 5)
}

# Рисует вертикальную таблицу расчетных профилей.
# В named range входит две строки шапки: строка с общей подписью ProfileId и
# строка со стабильными ID PR1/PR2/... . Reader ищет именно PR-строку и не
# зависит от заранее заданного количества профилей.
function Add-CalculationProfilesTable {
    param([object]$Workbook, [object]$Sheet, [int]$HeaderRow, [int]$StartColumn)

    $title = "Настройка расчетных профилей"
    $Sheet.Cells.Item($HeaderRow - 1, $StartColumn).Value2 = $title
    $Sheet.Range($Sheet.Cells.Item($HeaderRow - 1, $StartColumn), $Sheet.Cells.Item($HeaderRow - 1, $StartColumn + 6)).Merge() | Out-Null
    $Sheet.Cells.Item($HeaderRow - 1, $StartColumn).Font.Bold = $true
    $Sheet.Cells.Item($HeaderRow - 1, $StartColumn).Interior.Color = 15921906

    $headers = @("Параметр", "Key", "PR1", "PR2", "PR3", "PR4", "Комментарий")
    $profileHeader = $Sheet.Range($Sheet.Cells.Item($HeaderRow, $StartColumn + 2), $Sheet.Cells.Item($HeaderRow, $StartColumn + 5))
    $profileHeader.Merge() | Out-Null
    $profileHeader.Value2 = "ProfileId"
    $Sheet.Range($Sheet.Cells.Item($HeaderRow, $StartColumn), $Sheet.Cells.Item($HeaderRow, $StartColumn + 6)).Font.Bold = $true
    $Sheet.Range($Sheet.Cells.Item($HeaderRow, $StartColumn), $Sheet.Cells.Item($HeaderRow, $StartColumn + 6)).Interior.Color = 14277081
    $Sheet.Range($Sheet.Cells.Item($HeaderRow, $StartColumn), $Sheet.Cells.Item($HeaderRow, $StartColumn + 6)).HorizontalAlignment = -4108

    for ($i = 0; $i -lt $headers.Count; $i++) {
        $cell = $Sheet.Cells.Item($HeaderRow + 1, $StartColumn + $i)
        $cell.Value2 = $headers[$i]
        $cell.Font.Bold = $true
        $cell.Interior.Color = 14277081
        $cell.HorizontalAlignment = -4108
    }

    $rows = Get-CalculationProfilesCatalog
    for ($r = 0; $r -lt $rows.Count; $r++) {
        $targetRow = $HeaderRow + 2 + $r
        $Sheet.Cells.Item($targetRow, $StartColumn).Value2 = $rows[$r].Caption
        $Sheet.Cells.Item($targetRow, $StartColumn + 1).Value2 = $rows[$r].Key
        $Sheet.Cells.Item($targetRow, $StartColumn + 2).Value2 = $rows[$r].PR1
        $Sheet.Cells.Item($targetRow, $StartColumn + 3).Value2 = $rows[$r].PR2
        $Sheet.Cells.Item($targetRow, $StartColumn + 4).Value2 = $rows[$r].PR3
        $Sheet.Cells.Item($targetRow, $StartColumn + 5).Value2 = $rows[$r].PR4
        $Sheet.Cells.Item($targetRow, $StartColumn + 6).Value2 = $rows[$r].Comment

        if (([string]$rows[$r].Caption).StartsWith("[")) {
            $sectionRow = $Sheet.Range($Sheet.Cells.Item($targetRow, $StartColumn), $Sheet.Cells.Item($targetRow, $StartColumn + 6))
            $sectionRow.Font.Bold = $true
            $sectionRow.Interior.Color = $script:ConfigSubgroupHeaderColor
            $sectionRow.Font.Color = 4210752
        }
    }

    Set-WorkbookNameByBounds $Workbook "rngCalculationProfiles" $Sheet $HeaderRow $StartColumn ($HeaderRow + $rows.Count + 1) ($StartColumn + 6)
}

# Служебные точки диаграмм нужны только как независимый контроль Config.
# Расчетное ядро их не читает: те же формулы реализованы в CMaterialModelProvider.
function Add-MaterialDiagramControlTables {
    param([object]$Sheet, [int]$HeaderRow, [int]$StartColumn, [switch]$CoordinatesOnly)

    function New-MaterialParameterRef {
        param([string]$RangeName, [string]$ParameterName, [int]$ValueColumn)
        $range = $Sheet.Range($RangeName)
        for ($rowIndex = 1; $rowIndex -le $range.Rows.Count; $rowIndex++) {
            if ([string]$range.Cells.Item($rowIndex, 1).Value2 -eq $ParameterName) {
                return $range.Cells.Item($rowIndex, $ValueColumn).Address($true, $true)
            }
        }
        throw "Не найден параметр $ParameterName в диапазоне $RangeName для контрольных точек диаграмм."
    }

    if (-not $CoordinatesOnly) {
        $Sheet.Cells.Item($HeaderRow - 1, $StartColumn).Value2 = "Контрольные точки диаграмм"
        $Sheet.Cells.Item($HeaderRow - 1, $StartColumn).Font.Bold = $true
        $Sheet.Cells.Item($HeaderRow - 1, $StartColumn).Interior.Color = 15921906
        $headers = @("Материал", "ГПС", "Тип", "Ветвь", "ε", "")
        for ($i = 0; $i -lt $headers.Count; $i++) {
            $cell = $Sheet.Cells.Item($HeaderRow, $StartColumn + $i)
            $cell.Value2 = $headers[$i]
            $cell.Font.Bold = $true
            $cell.Interior.Color = 14277081
        }
        # Формулы используют INPUT-значения; подпись следует их единице.
        $stressUnitRef = New-MaterialParameterRef "rngUnitSettings" "Stress" 2
        $Sheet.Cells.Item($HeaderRow, $StartColumn + 5).Formula = '="σ, "&' + $stressUnitRef
    }
    $slopeFactor = New-MaterialParameterRef "rngSystemSettings" "General.PlateauSlopeFactor" 2

    $concreteGroups = @(
        @{ Group = "I"; Rb = (New-MaterialParameterRef "rngConcreteMaterialParameters" "Concrete.R.ULS(I)" 2); Rbt = (New-MaterialParameterRef "rngConcreteMaterialParameters" "Concrete.R.ULS(I)" 3); Eb = (New-MaterialParameterRef "rngConcreteMaterialParameters" "Concrete.E" 2); Ebt = (New-MaterialParameterRef "rngConcreteMaterialParameters" "Concrete.E" 3) },
        @{ Group = "II"; Rb = (New-MaterialParameterRef "rngConcreteMaterialParameters" "Concrete.R.SLS(II)" 2); Rbt = (New-MaterialParameterRef "rngConcreteMaterialParameters" "Concrete.R.SLS(II)" 3); Eb = (New-MaterialParameterRef "rngConcreteMaterialParameters" "Concrete.E" 2); Ebt = (New-MaterialParameterRef "rngConcreteMaterialParameters" "Concrete.E" 3) }
    )
    $steelGroups = @(
        @{ Group = "I"; Rsc = (New-MaterialParameterRef "rngSteelMaterialParameters" "Steel.R.ULS(I)" 2); Rs = (New-MaterialParameterRef "rngSteelMaterialParameters" "Steel.R.ULS(I)" 3); Esc = (New-MaterialParameterRef "rngSteelMaterialParameters" "Steel.E" 2); Es = (New-MaterialParameterRef "rngSteelMaterialParameters" "Steel.E" 3) },
        @{ Group = "II"; Rsc = (New-MaterialParameterRef "rngSteelMaterialParameters" "Steel.R.SLS(II)" 2); Rs = (New-MaterialParameterRef "rngSteelMaterialParameters" "Steel.R.SLS(II)" 3); Esc = (New-MaterialParameterRef "rngSteelMaterialParameters" "Steel.E" 2); Es = (New-MaterialParameterRef "rngSteelMaterialParameters" "Steel.E" 3) }
    )
    $concreteUser = @{
        Eb1Red = (New-MaterialParameterRef "rngConcreteMaterialParameters" "Concrete.TwoLine.Eb1Red" 2); Ebt1Red = (New-MaterialParameterRef "rngConcreteMaterialParameters" "Concrete.TwoLine.Eb1Red" 3)
        Eb0 = (New-MaterialParameterRef "rngConcreteMaterialParameters" "Concrete.ThreeLine.Eb0" 2); Ebt0 = (New-MaterialParameterRef "rngConcreteMaterialParameters" "Concrete.ThreeLine.Eb0" 3)
        Eb2 = (New-MaterialParameterRef "rngConcreteMaterialParameters" "Concrete.TwoThreeLine.Eb2" 2); Ebt2 = (New-MaterialParameterRef "rngConcreteMaterialParameters" "Concrete.TwoThreeLine.Eb2" 3)
    }
    $steelUser = @{
        TwoLineEsc2 = (New-MaterialParameterRef "rngSteelMaterialParameters" "Steel.TwoLine.Es2" 2); TwoLineEs2 = (New-MaterialParameterRef "rngSteelMaterialParameters" "Steel.TwoLine.Es2" 3)
        ThreeLineEsc2 = (New-MaterialParameterRef "rngSteelMaterialParameters" "Steel.ThreeLine.Es2" 2); ThreeLineEs2 = (New-MaterialParameterRef "rngSteelMaterialParameters" "Steel.ThreeLine.Es2" 3)
    }

    $concreteEb1Red = $concreteUser["Eb1Red"]
    $concreteEbt1Red = $concreteUser["Ebt1Red"]
    $concreteEb0 = $concreteUser["Eb0"]
    $concreteEbt0 = $concreteUser["Ebt0"]
    $concreteEb2 = $concreteUser["Eb2"]
    $concreteEbt2 = $concreteUser["Ebt2"]
    $steelTwoLineEsc2 = $steelUser["TwoLineEsc2"]
    $steelTwoLineEs2 = $steelUser["TwoLineEs2"]
    $steelThreeLineEsc2 = $steelUser["ThreeLineEsc2"]
    $steelThreeLineEs2 = $steelUser["ThreeLineEs2"]

    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($g in $concreteGroups) {
        $group = $g["Group"]
        $rb = $g["Rb"]
        $rbt = $g["Rbt"]
        $eb = $g["Eb"]
        $ebt = $g["Ebt"]
        $rows.Add(@("Concrete", $group, "TwoLine", "Compression", ("=-{0}" -f $concreteEb2), ("=-({0}+{1}*{2}*({3}-{4}))" -f $rb, $slopeFactor, $eb, $concreteEb2, $concreteEb1Red))) | Out-Null
        $rows.Add(@("Concrete", $group, "TwoLine", "Compression", ("=-{0}" -f $concreteEb1Red), ("=-{0}" -f $rb))) | Out-Null
        $rows.Add(@("Concrete", $group, "TwoLine", "Zero", "0", "0")) | Out-Null
        $rows.Add(@("Concrete", $group, "TwoLine", "Tension", ("={0}" -f $concreteEbt1Red), ("={0}" -f $rbt))) | Out-Null
        $rows.Add(@("Concrete", $group, "TwoLine", "Tension", ("={0}" -f $concreteEbt2), ("={0}+{1}*{2}*({3}-{4})" -f $rbt, $slopeFactor, $ebt, $concreteEbt2, $concreteEbt1Red))) | Out-Null
        $rows.Add(@("Concrete", $group, "ThreeLine", "Compression", ("=-{0}" -f $concreteEb2), ("=-({0}+{1}*{2}*({3}-{4}))" -f $rb, $slopeFactor, $eb, $concreteEb2, $concreteEb0))) | Out-Null
        $rows.Add(@("Concrete", $group, "ThreeLine", "Compression", ("=-{0}" -f $concreteEb0), ("=-{0}" -f $rb))) | Out-Null
        $rows.Add(@("Concrete", $group, "ThreeLine", "Compression", ("=-0.6*{0}/{1}" -f $rb, $eb), ("=-0.6*{0}" -f $rb))) | Out-Null
        $rows.Add(@("Concrete", $group, "ThreeLine", "Zero", "0", "0")) | Out-Null
        $rows.Add(@("Concrete", $group, "ThreeLine", "Tension", ("=0.6*{0}/{1}" -f $rbt, $ebt), ("=0.6*{0}" -f $rbt))) | Out-Null
        $rows.Add(@("Concrete", $group, "ThreeLine", "Tension", ("={0}" -f $concreteEbt0), ("={0}" -f $rbt))) | Out-Null
        $rows.Add(@("Concrete", $group, "ThreeLine", "Tension", ("={0}" -f $concreteEbt2), ("={0}+{1}*{2}*({3}-{4})" -f $rbt, $slopeFactor, $ebt, $concreteEbt2, $concreteEbt0))) | Out-Null
    }
    foreach ($g in $steelGroups) {
        $group = $g["Group"]
        $rsc = $g["Rsc"]
        $rs = $g["Rs"]
        $esc = $g["Esc"]
        $es = $g["Es"]
        $epsSc0 = "({0}/{1}+0.002)" -f $rsc, $esc
        $epsSc1 = "(0.9*{0}/{1})" -f $rsc, $esc
        $epsScPl = "(2*{0}-{1})" -f $epsSc0, $epsSc1
        $epsS0 = "({0}/{1}+0.002)" -f $rs, $es
        $epsS1 = "(0.9*{0}/{1})" -f $rs, $es
        $epsSPl = "(2*{0}-{1})" -f $epsS0, $epsS1
        $rows.Add(@("Steel", $group, "TwoLine", "Compression", ("=-{0}" -f $steelTwoLineEsc2), ("=-({0}+{1}*{2}*({3}-{0}/{2}))" -f $rsc, $slopeFactor, $esc, $steelTwoLineEsc2))) | Out-Null
        $rows.Add(@("Steel", $group, "TwoLine", "Compression", ("=-{0}/{1}" -f $rsc, $esc), ("=-{0}" -f $rsc))) | Out-Null
        $rows.Add(@("Steel", $group, "TwoLine", "Zero", "0", "0")) | Out-Null
        $rows.Add(@("Steel", $group, "TwoLine", "Tension", ("={0}/{1}" -f $rs, $es), ("={0}" -f $rs))) | Out-Null
        $rows.Add(@("Steel", $group, "TwoLine", "Tension", ("={0}" -f $steelTwoLineEs2), ("={0}+{1}*{2}*({3}-{0}/{2})" -f $rs, $slopeFactor, $es, $steelTwoLineEs2))) | Out-Null
        $rows.Add(@("Steel", $group, "ThreeLine", "Compression", ("=-{0}" -f $steelThreeLineEsc2), ("=-(1.1*{0}+{1}*{2}*({3}-{4}))" -f $rsc, $slopeFactor, $esc, $steelThreeLineEsc2, $epsScPl))) | Out-Null
        $rows.Add(@("Steel", $group, "ThreeLine", "Compression", ("=-{0}" -f $epsScPl), ("=-1.1*{0}" -f $rsc))) | Out-Null
        $rows.Add(@("Steel", $group, "ThreeLine", "Compression", ("=-{0}" -f $epsSc0), ("=-{0}" -f $rsc))) | Out-Null
        $rows.Add(@("Steel", $group, "ThreeLine", "Compression", ("=-{0}" -f $epsSc1), ("=-0.9*{0}" -f $rsc))) | Out-Null
        $rows.Add(@("Steel", $group, "ThreeLine", "Zero", "0", "0")) | Out-Null
        $rows.Add(@("Steel", $group, "ThreeLine", "Tension", ("={0}" -f $epsS1), ("=0.9*{0}" -f $rs))) | Out-Null
        $rows.Add(@("Steel", $group, "ThreeLine", "Tension", ("={0}" -f $epsS0), ("={0}" -f $rs))) | Out-Null
        $rows.Add(@("Steel", $group, "ThreeLine", "Tension", ("={0}" -f $epsSPl), ("=1.1*{0}" -f $rs))) | Out-Null
        $rows.Add(@("Steel", $group, "ThreeLine", "Tension", ("={0}" -f $steelThreeLineEs2), ("=1.1*{0}+{1}*{2}*({3}-{4})" -f $rs, $slopeFactor, $es, $steelThreeLineEs2, $epsSPl))) | Out-Null
    }

    for ($r = 0; $r -lt $rows.Count; $r++) {
        for ($c = 0; $c -lt $rows[$r].Count; $c++) {
            if ($CoordinatesOnly -and $c -lt 4) {continue}
            $value = [string]$rows[$r][$c]
            $cell = $Sheet.Cells.Item($HeaderRow + 1 + $r, $StartColumn + $c)
            if ($value.StartsWith("=")) {
                try {
                    $cell.Formula = $value
                } catch {
                    $address = $cell.Address($false, $false)
                    throw "Не удалось записать формулу контрольной таблицы диаграмм в $address`: $value. $($_.Exception.Message)"
                }
            } else {
                $cell.Value2 = $value
            }
        }
    }

    if (-not $CoordinatesOnly) {
        $dataRange = $Sheet.Range($Sheet.Cells.Item($HeaderRow, $StartColumn), $Sheet.Cells.Item($HeaderRow + $rows.Count, $StartColumn + 5))
        $dataRange.Borders.LineStyle = 1
        $dataRange.Borders.Weight = 2
        $dataRange.Columns.AutoFit() | Out-Null
    }
}

function Remove-MaterialDiagramCharts {
    param([object]$Sheet)

    foreach ($name in @("chMaterialConcreteDiagram", "chMaterialSteelDiagram", "grpMaterialDiagrams")) {
        try { $Sheet.ChartObjects($name).Delete() } catch {}
        try { $Sheet.Shapes.Item($name).Delete() } catch {}
    }
}

function Add-MaterialDiagramCharts {
    param(
        [object]$Sheet,
        [int]$TopRow,
        [int]$StartColumn,
        [int]$ControlHeaderRow,
        [int]$ControlStartColumn
    )

    $left = $Sheet.Cells.Item($TopRow, $StartColumn).Left + ($Sheet.Columns.Item($StartColumn).Width / 2)
    $top = $Sheet.Cells.Item($TopRow, $StartColumn).Top
    $width = 453.54
    $height = 255
    $gap = 18
    $epsColumn = $ControlStartColumn + 4
    $sigmaColumn = $ControlStartColumn + 5

    Add-MaterialDiagramChart $Sheet "chMaterialConcreteDiagram" "Диаграмма растяжения/сжатия бетона (I ГПС)" `
        $left $top $width $height ($ControlHeaderRow + 1) ($ControlHeaderRow + 5) ($ControlHeaderRow + 6) ($ControlHeaderRow + 12) $epsColumn $sigmaColumn $ControlHeaderRow
    Add-MaterialDiagramChart $Sheet "chMaterialSteelDiagram" "Диаграмма растяжения/сжатия арматуры (I ГПС)" `
        $left ($top + $height + $gap) $width $height ($ControlHeaderRow + 25) ($ControlHeaderRow + 29) ($ControlHeaderRow + 30) ($ControlHeaderRow + 38) $epsColumn $sigmaColumn $ControlHeaderRow
}

function Add-MaterialDiagramChart {
    param(
        [object]$Sheet,
        [string]$Name,
        [string]$Title,
        [double]$Left,
        [double]$Top,
        [double]$Width,
        [double]$Height,
        [int]$TwoLineStartRow,
        [int]$TwoLineEndRow,
        [int]$ThreeLineStartRow,
        [int]$ThreeLineEndRow,
        [int]$EpsColumn,
        [int]$SigmaColumn,
        [int]$ControlHeaderRow
    )

    $chartObject = $Sheet.ChartObjects().Add($Left, $Top, $Width, $Height)
    $chartObject.Name = $Name
    $chart = $chartObject.Chart
    $chart.ChartType = 74
    $chart.HasTitle = $true
    $chart.ChartTitle.Text = $Title
    try { $chart.ChartTitle.Font.Size = 12 } catch {}
    $chart.HasLegend = $true

    while ($chart.SeriesCollection().Count -gt 0) {
        $chart.SeriesCollection(1).Delete()
    }

    Add-MaterialDiagramChartSeries $Sheet $chart "TwoLine" $TwoLineStartRow $TwoLineEndRow $EpsColumn $SigmaColumn 15773696
    Add-MaterialDiagramChartSeries $Sheet $chart "ThreeLine" $ThreeLineStartRow $ThreeLineEndRow $EpsColumn $SigmaColumn 49407

    $chart.Axes(2).HasTitle = $true
    $sigmaCaption = $Sheet.Cells.Item($ControlHeaderRow, $SigmaColumn).Address($true, $true)
    $chart.Axes(2).AxisTitle.Formula = "='" + $Sheet.Name.Replace("'", "''") + "'!" + $sigmaCaption

    try {
        $chart.Axes(1).HasTitle = $true
        $chart.Axes(1).AxisTitle.Text = "ε"
        $chart.Axes(1).CrossesAt = 0
        $chart.Axes(2).CrossesAt = 0
        $chart.Axes(1).ReversePlotOrder = $true
        $chart.Axes(2).ReversePlotOrder = $true
        $chart.Axes(1).HasMajorGridlines = $false
        $chart.Axes(2).HasMajorGridlines = $false
        $chart.Axes(1).HasMinorGridlines = $false
        $chart.Axes(2).HasMinorGridlines = $false
    } catch {}
}

function Add-MaterialDiagramChartSeries {
    param(
        [object]$Sheet,
        [object]$Chart,
        [string]$SeriesName,
        [int]$StartRow,
        [int]$EndRow,
        [int]$EpsColumn,
        [int]$SigmaColumn,
        [int]$Color
    )

    $series = $Chart.SeriesCollection().NewSeries()
    $series.Name = $SeriesName
    $series.XValues = $Sheet.Range($Sheet.Cells.Item($StartRow, $EpsColumn), $Sheet.Cells.Item($EndRow, $EpsColumn))
    $series.Values = $Sheet.Range($Sheet.Cells.Item($StartRow, $SigmaColumn), $Sheet.Cells.Item($EndRow, $SigmaColumn))
    try {
        $series.Format.Line.ForeColor.RGB = $Color
        $series.Format.Line.Weight = 1.5
        $series.MarkerStyle = 8
        $series.MarkerSize = 5
    } catch {}
}

# Сохраняет служебные области печати до Excel COM save. На локализованном
# Excel имя Print_Area может быть потеряно при повторном сохранении книги.
# Остальные имена/ячейки не входят в этот точечный снимок.
function Get-WorkbookPrintAreas {
    param([string]$Path)
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $reader = New-Object IO.StreamReader($zip.GetEntry("xl/workbook.xml").Open())
        try { [xml]$document = $reader.ReadToEnd() } finally { $reader.Dispose() }
        $areas = New-Object System.Collections.Generic.List[object]
        foreach ($node in $document.SelectNodes("//*[local-name()='definedName']")) {
            if ($node.GetAttribute("name") -in @("_xlnm.Print_Area", "Print_Area")) {
                $areas.Add(@{ Sheet = $node.GetAttribute("localSheetId"); Xml = $node.OuterXml })
            }
        }
        return $areas.ToArray()
    } finally { $zip.Dispose() }
}

# Восстанавливает только ранее существовавшие print names после закрытия Excel.
# Использует XML DOM, не строковые замены; не меняет пользовательские настройки.
function Restore-WorkbookPrintAreas {
    param([string]$Path, [object[]]$Areas)
    if (-not $Areas.Count) { return }
    Add-Type -AssemblyName System.IO.Compression
    $zip = [IO.Compression.ZipFile]::Open($Path, [IO.Compression.ZipArchiveMode]::Update)
    try {
        $entry = $zip.GetEntry("xl/workbook.xml")
        $reader = New-Object IO.StreamReader($entry.Open())
        try { [xml]$document = $reader.ReadToEnd() } finally { $reader.Dispose() }
        $definedNames = $document.SelectSingleNode("//*[local-name()='definedNames']")
        if ($null -eq $definedNames) {
            $definedNames = $document.CreateElement("definedNames", $document.DocumentElement.NamespaceURI)
            $next = $document.SelectSingleNode("//*[local-name()='calcPr']")
            if ($next) { $document.DocumentElement.InsertBefore($definedNames, $next) | Out-Null }
            else { $document.DocumentElement.AppendChild($definedNames) | Out-Null }
        }
        $seen = @{}
        foreach ($area in $Areas) {
            if ($seen.ContainsKey($area.Sheet)) { continue }
            $seen[$area.Sheet] = $true
            foreach ($existing in @($definedNames.ChildNodes)) {
                if ($existing.GetAttribute("name") -in @("_xlnm.Print_Area", "Print_Area") -and
                    $existing.GetAttribute("localSheetId") -eq $area.Sheet) {
                    $definedNames.RemoveChild($existing) | Out-Null
                }
            }
            [xml]$source = $area.Xml
            $copy = $document.ImportNode($source.DocumentElement, $true)
            $copy.SetAttribute("name", "_xlnm.Print_Area")
            $definedNames.AppendChild($copy) | Out-Null
        }
        $stream = $entry.Open()
        $stream.SetLength(0)
        $writer = New-Object IO.StreamWriter($stream, (New-Object Text.UTF8Encoding($false)))
        try { $document.Save($writer) } finally { $writer.Dispose() }
    } finally { $zip.Dispose() }
}

# Создаёт структуру нового шаблона; готовая книга не вызывает этот сборщик.
function Apply-SystemSettingsLayout {
    param([object]$Workbook, [object]$Sheet)

    $catalog = Get-SystemSettingsCatalog
    Reset-ConfigValidationSources $Sheet
    $Sheet.Range("A1:EF260").ClearContents()
    $Sheet.Range("A1:EF260").Validation.Delete()
    Remove-MaterialDiagramCharts $Sheet

    $Sheet.Cells.Item(1, 1).Value2 = "Config"
    $Sheet.Cells.Item(1, 1).Font.Bold = $true
    $Sheet.Cells.Item(1, 1).Font.Size = 16

    $headers = @("Параметр", "Значение", "Ед.", "Комментарий", "Справка")
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
        $sectionRange = $Sheet.Range($Sheet.Cells.Item($row, 1), $Sheet.Cells.Item($row, 5))
        $sectionRange.Font.Bold = $true
        $sectionRange.Interior.Color = $script:ConfigGroupHeaderColor
        $row++
        foreach ($setting in $section.Rows) {
            for ($c = 0; $c -lt 4; $c++) { $Sheet.Cells.Item($row, $c + 1).Value2 = $setting[$c] }
            Set-InputUnitCell $Sheet.Cells.Item($row, 3) ([string]$setting[2])
            $settingKey = [string]$setting[0]
            if ($settingKey.StartsWith("[")) {
                $subHeader = $Sheet.Range($Sheet.Cells.Item($row, 1), $Sheet.Cells.Item($row, 5))
                $subHeader.Font.Bold = $true
                $subHeader.Interior.Color = $script:ConfigSubgroupHeaderColor
                $subHeader.Font.Color = 4210752
            }
            $row++
        }
        $row++
    }

    $settingsRange = $Sheet.Range($Sheet.Cells.Item(3, 1), $Sheet.Cells.Item($row - 2, 5))
    Set-WorkbookNameByBounds $Workbook "rngSystemSettings" $Sheet 3 1 ($row - 2) 5

    # Правая часть листа собирается одним вертикальным стеком.
    # A:E занимает общий реестр rngSystemSettings, F:G оставлены пустым
    # визуальным зазором, а все специализированные диапазоны начинаются с H.
    $rightColumn = 8
    $rightRow = 3
    $rightBlockGap = 2
    $loadCombinationsHeaderRow = 3
    $loadCombinationsColumn = 15
    $materialControlHeaderRow = 2
    $materialControlColumn = 35
    $materialChartTopRow = 1
    $materialChartColumn = 22
    Add-SP35Table721 $Workbook $Sheet 60 35

    Add-UnitSettingsTable $Workbook $Sheet $rightRow $rightColumn
    $rightRow += (Get-UnitSettingsCatalog).Count + 1 + $rightBlockGap

    Add-LoadCombinationsTable $Workbook $Sheet $loadCombinationsHeaderRow $loadCombinationsColumn
    Add-StabilityDurationLoadsTable $Workbook $Sheet 37 16

    Add-SignConventionSettingsTable $Workbook $Sheet $rightRow $rightColumn
    $rightRow += (Get-SignConventionSettingsCatalog).Count + 1 + $rightBlockGap

    $steelRows = Get-SteelMaterialParametersCatalog
    Add-MaterialParameterTable $Workbook $Sheet "rngSteelMaterialParameters" $rightRow $rightColumn "Материал арматуры" $steelRows
    $rightRow += $steelRows.Count + 1 + $rightBlockGap

    $concreteRows = Get-ConcreteMaterialParametersCatalog
    Add-MaterialParameterTable $Workbook $Sheet "rngConcreteMaterialParameters" $rightRow $rightColumn "Материал бетона" $concreteRows
    $rightRow += $concreteRows.Count + 1 + $rightBlockGap

    Add-MaterialDiagramControlTables $Sheet $materialControlHeaderRow $materialControlColumn

    Add-CalculationProfilesTable $Workbook $Sheet $rightRow $rightColumn
    Add-LoadProfileValidation $Workbook $Sheet
    $rightRow += (Get-CalculationProfilesCatalog).Count + 2 + $rightBlockGap

    Add-PlotAnnotationSettingsTable $Workbook $Sheet $rightRow $rightColumn
    $rightRow += (Get-PlotAnnotationSettingsCatalog).Count + 1 + $rightBlockGap

    Add-SectionSettingsTables $Workbook $Sheet | Out-Null

    Add-MaterialDiagramCharts $Sheet $materialChartTopRow $materialChartColumn $materialControlHeaderRow $materialControlColumn

    $lambda = [char]0x03BB
    $mu = [char]0x03BC
    $validationLists = [ordered]@{
        "Geometry.Source" = @("Generated", "AutoCAD")
        "Mesh.CellGeometry" = @("Adaptive", "Contour")
        "Mesh.StepMode" = @("Auto 0.05", "Auto 0.1", "Auto 0.2", "User")
        "General.ExecutionReportEnabled" = @("Yes", "No")
        "General.NonCriticalMessagesEnabled" = @("Yes", "No")
        "General.ReformatResultsEnabled" = @("Yes", "No")
        "Results.SnapshotStorage" = @("Excel", "File")
        "General.WorstCombinationCriterion" = @("StrengthStrain", "StrengthCapacity", "Cracks", "Stability")
        "Stability.Code" = @("SP63", "SP35")
        "Stability.SystemType" = @("Determinate", "Indeterminate")
        "Stability.ZeroMomentEccentricitySign1" = @("1", "-1")
        "Stability.ZeroMomentEccentricitySign2" = @("1", "-1")
        "Stability.AccidentalEccentricityMode" = @("AutoWithL", "AutoWith$($mu)L", "User")
        "Stability.AccidentalEccentricityPlanes" = @("OnlyMomentPlane", "BothPlanes")
        "Stability.PhiLMode" = @("Auto", "PhiL2")
        "Solver.Method" = @("Newton", "Secant")
        "General.DiagramExtension" = @("Yes", "No")
        "Capacity.SolutionStrategy" = @("Auto", "UltimateStrain", "LoadMultiplier")
        "Capacity.SearchMethod" = @("Bisection", "Brent", "Secant")
        "SLS.Crack.Code" = @("SP63", "SP35")
        "SLS.Crack.SP35.RadiusDiameterMode" = @("Max", "Min", "Average")
        "SLS.Crack.SP35.InteractionRadiusMode" = @("3d", "5d", "6d")
        "SLS.Crack.SP35.RebarProfile" = @("Periodic", "Smooth")
        "SLS.Crack.Phi3Mode" = @("Auto", "User")
        "SLS.Crack.PsiMode" = @("User", "Auto", "AlwaysCalc")
        "SLS.Crack.SigmaSCrcAveragingMode" = @("AllSelected", "TensionOnly")
        "SLS.Crack.InitiationSolutionStrategy" = @("Auto", "UltimateStrain", "LoadMultiplier")
        "SLS.Crack.TensionZoneMode" = @("Effective", "FullTension")
        "SLS.Crack.CoverDistanceMode" = @("NearestContour", "GlobalExtreme")
        "AutoCAD.Export.LabelMode" = @("ValuesOnly", "NamesAndValues")
        "AutoCAD.Export.NeutralLineEnabled" = @("Yes", "No")
        "AutoCAD.Export.PrincipalAxesMode" = @("Transformed", "Concrete", "None")
        "AutoCAD.Export.LoadPointEnabled" = @("Yes", "No")
        "AutoCAD.Export.ContourEnabled" = @("Yes", "No")
        "AutoCAD.Export.ContourFormat" = @("Region", "Polyline")
        "AutoCAD.Export.ExportCrackInteractionContour" = @("Yes", "No")
        "Plot.Enabled" = @("Yes", "No")
        "Plot.AutoUpdateAfterCalculation" = @("Yes", "No")
        "Plot.ResultGradient" = @("Yes", "No")
        "Plot.ResultLabelsEnabled" = @("Yes", "No")
        "Plot.NeutralLineEnabled" = @("Yes", "No")
        "Plot.PrincipalAxesMode" = @("Transformed", "Concrete", "None")
        "Plot.LoadApplicationPointEnabled" = @("Yes", "No")
        "Plot.AxisLabelsEnabled" = @("Yes", "No")
        "Plot.ContourEnabled" = @("Yes", "No")
        "Plot.LegendEnabled" = @("Yes", "No")
        "Plot.LegendMode" = @("Separate", "Common")
    }

    foreach ($key in $validationLists.Keys) {
        $options = $validationLists[$key]
        $listAddress = Get-ConfigValidationListAddress $Sheet $options
        for ($r = 4; $r -le ($row - 2); $r++) {
            if ([string]$Sheet.Cells.Item($r, 1).Value2 -eq $key) {
                $cell = $Sheet.Cells.Item($r, 2)
                $cell.Validation.Delete()
                $cell.Validation.Add(3, 1, 1, $listAddress)
            }
        }
    }
    for ($r = 2; $r -le $settingsRange.Rows.Count; $r++) {
        if ([string]$settingsRange.Cells.Item($r, 1).Value2 -ceq 'General.PlateauSlopeFactor') {
            $cell = $settingsRange.Cells.Item($r, 2)
            Set-ConfigNumericValue $cell ([double]::Parse(([string]$cell.Value2), [Globalization.CultureInfo]::InvariantCulture))
            $cell.NumberFormat = '0.########'
            $cell.Validation.Delete(); $cell.Validation.Add(2, 1, 7, '0')
            $cell.Validation.IgnoreBlank = $false; $cell.Validation.ShowError = $true
        }
    }

    $exportCombinationListColumn = Get-NextConfigValidationColumn $Sheet 'LoadCases'
    $loadCasePickerHelperRows = $Workbook.Names.Item("rngLoadCombinations").RefersToRange.Rows.Count - 1
    $Sheet.Cells.Item(1, $exportCombinationListColumn).Value2 = "Worst"
    for ($i = 1; $i -le $loadCasePickerHelperRows; $i++) {
        $Sheet.Cells.Item($i + 1, $exportCombinationListColumn).Formula = '=IF(ROW()-1>ROWS(rngLoadCombinations)-1,"",INDEX(rngLoadCombinations,ROW(),1))' # Счётчик зависит от своей строки, а не от перемещаемых ячеек системных настроек A:E.
    }
    $exportCombinationColName = ConvertTo-ExcelColumn $exportCombinationListColumn
    $exportCombinationListAddress = "=$" + $exportCombinationColName + '$1:$' + $exportCombinationColName + '$' + ($loadCasePickerHelperRows + 1)
    for ($r = 2; $r -le $settingsRange.Rows.Count; $r++) {
        if ([string]$settingsRange.Cells.Item($r, 1).Value2 -eq "AutoCAD.Export.CombinationID" -or [string]$settingsRange.Cells.Item($r, 1).Value2 -eq "Plot.LoadCase") {
            $cell = $settingsRange.Cells.Item($r, 2)
            $cell.Validation.Delete()
            $cell.Validation.Add(3, 1, 1, $exportCombinationListAddress)
            $cell.Validation.IgnoreBlank = $false
            $cell.Validation.InCellDropdown = $true
        }
    }
    Add-CalculationProfilesValidation $Sheet
    Add-PlotAnnotationValidation $Sheet

    $Sheet.Columns.Item(1).ColumnWidth = 34
    $Sheet.Columns.Item(2).ColumnWidth = 18
    $Sheet.Columns.Item(3).ColumnWidth = 12
    $Sheet.Columns.Item(4).ColumnWidth = 20
    $Sheet.Columns.Item(5).ColumnWidth = 11
    $Sheet.Columns.Item(6).ColumnWidth = 11
    $Sheet.Columns.Item(7).ColumnWidth = 11
    $Sheet.Columns.Item(8).ColumnWidth = 32
    # Ширину Config назначаем один раз при сборке; I–U одинаковы для всех таблиц.
    for ($colIndex = 9; $colIndex -le 24; $colIndex++) {
        $Sheet.Columns.Item($colIndex).ColumnWidth = 17
    }
    for ($colIndex = 25; $colIndex -le 69; $colIndex++) {
        $Sheet.Columns.Item($colIndex).ColumnWidth = 8.43
    }
    $Sheet.Columns.Item("BR:EF").Hidden = $true
    $Sheet.Range("A1:EF260").Font.Name = "Arial"
    $Sheet.Range("A1:EF260").Font.Size = 9
    Apply-ConfigUserInputAlignment $Workbook
    Apply-ConfigCommentColumnAlignment $Workbook
    Apply-ConfigShortColumnAlignment $Workbook
    Apply-ConfigNamedRangeBorders $Workbook

    return @{ Settings = $settingsRange }
}

# Назначает списки выбора полям оформления арматуры и размеров; ячейки с прочерком не делает редактируемыми списками.
function Add-PlotAnnotationValidation {
    param([object]$Sheet)

    $validationSources = @{
        "Enabled" = @("Yes", "No")
        "Placement" = @("Outside", "Inside")
        "TextUnits" = @("mm", "pt")
        "LineEnabled" = @("Yes", "No")
        "ArrowType" = @("Triangle", "Stealth", "Diamond", "Oval", "Open")
        "ArrowSize" = @("Small", "Medium", "Wide")
    }

    $range = $Sheet.Parent.Names.Item("rngPlotAnnotationSettings").RefersToRange
    $rows = $range.Value2
    for ($r = 2; $r -le $range.Rows.Count; $r++) {
        $paramName = [string]$rows[$r, 1]
        $source = $validationSources[$paramName]

        if ($null -ne $source) {
            $listAddress = Get-ConfigValidationListAddress $Sheet $source
            for ($c = 2; $c -le 3; $c++) {
                if ([string]$rows[$r, $c] -ne "-") {
                    $cell = $range.Cells.Item($r, $c)
                    $cell.Validation.Delete()
                    $cell.Validation.Add(3, 1, 1, $listAddress)
                    $cell.Validation.IgnoreBlank = $false
                    $cell.Validation.InCellDropdown = $true
                }
            }
        }
    }
}

# Добавляет валидацию для профильной таблицы. Типы расчетов, материальные
# модели и визуализация выбираются в PR-колонках профилей.
function Add-CalculationProfilesValidation {
    param([object]$Sheet)

    $sources = @{
        "Calculation.Strength.DirectState" = @("Yes", "No")
        "Calculation.Strength.Capacity" = @("Yes", "No")
        "Calculation.Crack.Width" = @("Yes", "No")
        "Calculation.Stability.Enabled" = @("Yes", "No")
        "MaterialModel.Stability.ValueSet" = @("ULS(I)", "SLS(II)")
        "MaterialModel.Strength.ValueSet" = @("ULS(I)", "SLS(II)")
        "MaterialModel.CrackInitiation.ValueSet" = @("ULS(I)", "SLS(II)")
        "MaterialModel.CrackedState.ValueSet" = @("ULS(I)", "SLS(II)")
        "MaterialModel.Strength.ConcreteDiagram" = @("TwoLine", "ThreeLine")
        "MaterialModel.CrackInitiation.ConcreteDiagram" = @("TwoLine", "ThreeLine")
        "MaterialModel.CrackedState.ConcreteDiagram" = @("TwoLine", "ThreeLine")
        "MaterialModel.Strength.ConcreteTension" = @("Ignore", "UseDiagram")
        "MaterialModel.Strength.SteelDiagram" = @("TwoLine", "ThreeLine")
        "MaterialModel.CrackInitiation.SteelDiagram" = @("TwoLine", "ThreeLine")
        "MaterialModel.CrackedState.SteelDiagram" = @("TwoLine", "ThreeLine")
        "Visualization.State" = @("StrengthState", "CapacityState", "PreCrackState", "PostCrackState", "CrackedState")
        "Visualization.Quantity" = @("Stress", "Strain")
    }

    try {
        $range = $Sheet.Parent.Names.Item("rngCalculationProfiles").RefersToRange
        for ($r = 3; $r -le $range.Rows.Count; $r++) {
            $key = [string]$range.Cells.Item($r, 2).Value2
            if ($sources.ContainsKey($key)) {
                $source = $sources[$key]
                $listAddress = Get-ConfigValidationListAddress $Sheet $source
                for ($c = 3; $c -le 6; $c++) {
                    $cell = $range.Cells.Item($r, $c)
                    $cell.Validation.Delete()
                    $cell.Validation.Add(3, 1, 1, $listAddress)
                    $cell.Validation.IgnoreBlank = $true
                    $cell.Validation.InCellDropdown = $true
                }
            }
        }
    }
    catch {
        # Отсутствие таблицы будет поймано отдельной проверкой именованных диапазонов.
    }
}

# Записывает именованную таблицу точек диаграммы Strain/Stress, связывает заголовок напряжений с единицами вывода и резервирует десять строк.
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

# Обычный блок справки Config: заголовок и нумерованные пояснения в A/B.
# Обновление использует существующие строки, чтобы не сдвигать ссылки и группировку.
function Set-SettingInstructionBlock {
    param([object]$Sheet, [int]$Row, [string]$Title, [string[]]$Lines, [int]$LastRow = 0)
    if ($LastRow -eq 0) { $LastRow = $Row + $Lines.Count }
    if ($LastRow -lt $Row + $Lines.Count) { throw "Недостаточно строк справки для $Title." }
    $block = $Sheet.Range($Sheet.Cells.Item($Row, 1), $Sheet.Cells.Item($LastRow, 6))
    $block.UnMerge() | Out-Null
    $block.ClearContents() | Out-Null
    $block.Font.Name = 'Arial'; $block.Font.Size = 10
    $block.Font.Bold = $false; $block.Font.Italic = $false
    $block.Font.Color = 0; $block.Font.Underline = -4142
    $block.Interior.Pattern = -4142
    $block.NumberFormat = $Sheet.Cells.Item(1, 1).NumberFormat
    $block.IndentLevel = 0
    $block.HorizontalAlignment = 1; $block.VerticalAlignment = -4160
    $block.WrapText = $true
    $block.Borders.LineStyle = 1; $block.Borders.Weight = 2
    $block.Borders.Color = 14277081
    $block.Borders.Item(10).Color = 0
    $heading = $Sheet.Range($Sheet.Cells.Item($Row, 1), $Sheet.Cells.Item($Row, 6))
    $heading.Merge() | Out-Null
    $heading.Value2 = $Title; $heading.Font.Bold = $true
    $heading.Interior.Color = 15921906; $heading.WrapText = $false
    # Merge может вернуть цвет общей границы от соседней строки.
    $heading.Borders.Item(8).Color = 14277081
    $heading.Borders.Item(9).Color = 14277081
    for ($index = 0; $index -lt $Lines.Count; $index++) {
        Set-InstructionParagraphNumber $Sheet.Cells.Item(($Row + $index + 1), 1) ($index + 1)
        $text = $Sheet.Cells.Item(($Row + $index + 1), 2)
        $text.Value2 = [string]$Lines[$index]; $text.HorizontalAlignment = -4131
    }
    # AutoFit работает для обычных ячеек; объединённый заголовок однострочный.
    $block.EntireRow.AutoFit() | Out-Null
    $heading.EntireRow.RowHeight = [double]$Sheet.StandardHeight
    return $LastRow + 1
}

# Применяет общий макет настроек к родительской книге указанного листа
# и возвращает именованный диапазон, подготовленный Apply-SystemSettingsLayout.
function Add-SystemSettings {
    param([object]$Sheet)
    $result = Apply-SystemSettingsLayout $Sheet.Parent $Sheet
    $result
}
