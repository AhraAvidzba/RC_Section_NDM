# скрипт описывает каталог настроек Config и строит таблицы настроек, единиц, геометрии и валидации.

. (Join-Path $PSScriptRoot 'WorkbookLayout.ps1')
. (Join-Path $PSScriptRoot 'NormativeFormulaImages.ps1')

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
        @{ Name = "LoadDiagramSettings"; Title = "[Нагрузки, сочетания и диаграммы материалов]"; Rows = @(
            @("[Нагрузки]", "", "-", ""),
            @("Load.ReferenceOffsetX", "0", "мм", "Смещение точки приложения нагрузок по X относительно центра тяжести бетонного сечения без учёта арматуры."),
            @("Load.ReferenceOffsetY", "0", "мм", "Смещение точки приложения нагрузок по Y относительно центра тяжести бетонного сечения. Рабочие моменты для решателя автоматически переносятся к координатам расчётных элементов."),
            @("Calculation.ZeroMomentPerDepth", "0.0005", "Н*мм/мм", "Порог обнуления малых моментов: M_tol = значение·h. Убирает численный остаток после переносов и поворотов; слишком большое значение может обнулить значимый малый момент."),
            @("[Выводимое сочетание]", "", "-", ""),
            @("Output.CombinationMode", "Worst", "-", "Общий выбор сочетания для схемы Excel и экспорта AutoCAD: Worst — определяющее сочетание сохранённого расчёта; User — сочетание из Output.CombinationID."),
            @("General.WorstCombinationCriterion", "StrengthCapacity", "-", "Критерий выбора сочетания Worst при расчёте: прочность по деформациям, прочность по несущей, трещины или устойчивость. Схема и экспорт используют сохранённый выбор."),
            @("Output.CombinationID", "", "-", "ID сочетания для схемы и экспорта при Output.CombinationMode = User. Список берётся непосредственно из таблицы сочетаний; в режиме Worst поле не используется."),
            @("[Диаграммы материалов]", "", "-", ""),
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
            @("Stability.AccidentalEccentricityUser1", "10", "мм", "Модуль случайного эксцентриситета плоскости 1 для режима User: значение вдоль нормали к главной оси 1."),
            @("Stability.AccidentalEccentricityUser2", "10", "мм", "Модуль случайного эксцентриситета плоскости 2 для режима User: значение вдоль нормали к главной оси 2."),
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
            @("[Экспорт]", "", "", ""),
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
            @("Plot.Enabled", "Yes", "-", "Включить построение схемы в области chtNDMSectionPlot по сохранённому источнику Excel/File."),
            @("Plot.AutoUpdateAfterCalculation", "Yes", "-", "При Plot.Enabled=Yes обязателен выбор Yes/No: обновлять схему после расчета или сохранять готовую схему. При Plot.Enabled=No настройка не используется. Кнопка обновления читает выбранный источник Excel/File."),
            @("Plot.ResultGradient", "Yes", "-", "Включить цветовое различение результата на схеме. Цвет физического состояния берется из PhysicalState, записанного в Results."),
            @("Plot.ResultLabelsEnabled", "No", "-", "Показывать пространственно распределенные численные подписи величины Visualization.Quantity выбранного профиля для бетонных элементов."),
            @("Plot.ResultLabelSpacing", "100", "мм", "Минимальный пространственный шаг между численными подписями результата на схеме."),
            @("Plot.NeutralLineEnabled", "Yes", "-", "Показывать нейтральную линию выбранного сочетания по ε_0, κ_x, κ_y из Results."),
            @("Plot.NeutralLineWeight", "0.7", "pt", "Толщина нейтральной линии на Excel-схеме. Используется только при включённой нейтральной линии."),
            @("Plot.PrincipalAxesMode", "Transformed", "-", "Какие главные центральные оси показывать на схеме: Transformed - приведенного сечения; Concrete - бетонного; None - не показывать."),
            @("Plot.AxesLineWeight", "0.56", "pt", "Толщина обеих главных осей на Excel-схеме. Используется при Plot.PrincipalAxesMode = Transformed или Concrete."),
            @("Plot.LoadApplicationPointEnabled", "Yes", "-", "Показывать точку приложения нагрузки из последнего расчета."),
            @("Plot.AxisLabelsEnabled", "Yes", "-", "Показывать маленькие курсивные подписи глобальных осей X и Y внутри рабочей области схемы."),
            @("Plot.AxisLabelsFontSize", "7.5", "pt", "Размер шрифта подписей осей X и Y на схеме в пунктах Excel."),
            @("Plot.ContourEnabled", "Yes", "-", "Показывать параметрический контур сечения на схеме, если он сохранен в Results. Для импортированной сетки контур не восстанавливается."),
            @("Plot.LegendEnabled", "Yes", "-", "Показывать легенду физического состояния и выбранной величины профиля справа от схемы."),
            @("Plot.LegendMode", "Separate", "-", "Separate - отдельные легенды для арматуры и бетона; Common - одна общая легенда. В режиме Common используются цвета Plot.Color.RebarCompression и Plot.Color.RebarTension."),
            @("Plot.LegendScale", "1", "-", "Общий масштаб легенды вместе с надписями и отступами: 1 — исходный размер, 0.5 — вдвое меньше. Привязка к правому верхнему краю области схемы; освободившееся место используется для сечения."),
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
        @("Length", "mm", "mm", "mm"),
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

# Возвращает короткий список непосредственно для проверки данных, без служебных ячеек.
# Имена сочетаний и профилей передаются прямыми ссылками на их видимые таблицы отдельно.
function Get-ConfigValidationListAddress {
    param([object]$Sheet, [array]$Options)
    if ($Options.Count -eq 0) { throw 'Пустой список допустимых значений Config.' }
    $inline = @($Options | ForEach-Object {[string]$_}) -join [string]$Sheet.Application.International(5)
    if ($inline.Length -gt 255 -or @($Options | Where-Object {[string]$_ -match '[,;\r\n]'}).Count) {
        throw 'Перечисление не помещается в проверку данных: используйте прямую ссылку на видимую исходную таблицу.'
    }
    return $inline
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
        @{ Name = "rngSP35Table721"; Columns = @(1, 2, 3, 4, 5, 6) }
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

# Возвращает описание конкретной настройки из единственного каталога справки.
# Параметр Comment сохраняет интерфейс сборочных инструментов; текст не дублируется в комментариях ячеек.
function Get-SettingInstructionLines {
    param([string]$Key, [string]$Comment)
    $entry = @(Get-SettingsInstructionCatalog | Where-Object { $_.Key -ceq $Key })
    if ($entry.Count -ne 1) { throw "Нет единственного описания настройки справки: $Key" }
    return @($entry[0].Lines)
}

# Читает полные описания, проверяя покрытие всех действующих настроек сборки.
# Нормативные ссылки разрешаются из общего каталога, без второй редактируемой копии текста.
function Get-SettingsInstructionCatalog {
    $definitions = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'HelpSettingsDescriptions.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    $references = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'NormativeHelpReferences.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    $expected = @('Units','SignConvention','LoadCombinations','StabilityLoads','SP35Table721',
        'ConcreteMaterialParameters','SteelMaterialParameters','CalculationProfiles','MaterialDiagramControlTables',
        'PlotAnnotationSettings','DependentRebarRows','CircleGeometry','PointedSectionGeometry','HollowRectangleGeometry','RectanglesGeometry')
    foreach ($section in (Get-SystemSettingsCatalog)) {
        foreach ($row in $section.Rows) { $key=[string]$row[0]; if ($key -and -not $key.StartsWith('[')) { $expected += $key } }
    }
    foreach ($row in (Get-PlotAnnotationSettingsCatalog)) { $expected += [string]$row[0] }
    foreach ($row in (Get-ConcreteMaterialParametersCatalog)) { $expected += [string]$row[0] }
    foreach ($row in (Get-SteelMaterialParametersCatalog)) { $expected += [string]$row[0] }
    $items = New-Object System.Collections.Generic.List[object]
    $seen = @{}
    foreach ($definition in $definitions) {
        $key=[string]$definition.Key
        if ($seen.ContainsKey($key) -or $expected -cnotcontains $key) { throw "Повторный или неизвестный блок справки: $key" }
        $seen[$key]=$true
        $lines = @($definition.Lines | ForEach-Object {
            $line=[string]$_
            if ($line -ceq '{{NormativeReference}}') {
                $reference=$references.PSObject.Properties[$key]
                if ($null -eq $reference) { throw "Нет нормативной ссылки для $key" }
                [string]$reference.Value
            } else { $line }
        })
        if (-not $definition.Title -or $lines.Count -eq 0 -or @($lines | Where-Object { [string]::IsNullOrWhiteSpace($_) }).Count) { throw "Неполное описание справки: $key" }
        $item=@{Key=$key;Title=[string]$definition.Title;Lines=$lines}
        if ($definition.PSObject.Properties['RelatedKey']) { $item['RelatedKey']=[string]$definition.RelatedKey }
        if ($definition.PSObject.Properties['Tables']) { $item['Tables']=@($definition.Tables) }
        $items.Add($item) | Out-Null
    }
    foreach ($key in $expected) { if (-not $seen.ContainsKey($key)) { throw "Новая настройка требует описания справки: $key" } }
    foreach ($reference in $references.PSObject.Properties) {
        $entry=@($items | Where-Object { $_.Key -ceq $reference.Name })
        if ($entry.Count -ne 1 -or $entry[0].Lines -cnotcontains [string]$reference.Value) { throw "Потеряна нормативная ссылка: $($reference.Name)" }
    }
    foreach ($item in $items) {
        if ($item.ContainsKey('RelatedKey') -and -not $seen.ContainsKey($item.RelatedKey)) { throw "Потеряна связанная справка: $($item.Key)" }
    }
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
    $bitmap = New-Object System.Drawing.Bitmap(($width * 2), ($height * 2))
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    $graphics.ScaleTransform(2, 2) # Четыре пикселя на пункт при неизменном размере формулы на листе.
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

    $bitmap = New-Object System.Drawing.Bitmap(($width * 2), ($height * 2))
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    $graphics.ScaleTransform(2, 2) # Четыре пикселя на пункт при неизменном размере формулы на листе.
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

    $bitmap = New-Object System.Drawing.Bitmap(($width * 2), ($height * 2))
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    $graphics.ScaleTransform(2, 2) # Четыре пикселя на пункт при неизменном размере формулы на листе.
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

        $bitmap = New-Object System.Drawing.Bitmap(($width * 2), ($height * 2))
        $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    $graphics.ScaleTransform(2, 2) # Четыре пикселя на пункт при неизменном размере формулы на листе.
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

    $bitmap = New-Object System.Drawing.Bitmap(($width * 2), ($height * 2))
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    $graphics.ScaleTransform(2, 2) # Четыре пикселя на пункт при неизменном размере формулы на листе.
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

# Общий каталог местных расшифровок: каждая показанная формула объясняется непосредственно под собой.
$script:HelpFormulaExplanations = @{}
foreach ($entry in (Get-Content -LiteralPath (Join-Path $PSScriptRoot 'HelpFormulaExplanations.json') -Raw -Encoding UTF8 | ConvertFrom-Json)) {
    $caption = ([string]$entry.Caption).Trim()
    if ($script:HelpFormulaExplanations.ContainsKey($caption)) { throw "Повторная расшифровка формулы: $caption" }
    $script:HelpFormulaExplanations[$caption] = [string]$entry.Text
}

# Дописывает обозначения и индексы сразу после формулы; неизвестная формула требует своей расшифровки.
function Add-GuideFormulaExplanation {
    param([object]$Sheet, [int]$Row, [string]$Caption)
    $key = $Caption.Trim()
    if (-not $script:HelpFormulaExplanations.ContainsKey($key)) { throw "Нет расшифровки формулы справки: $key" }
    return (Add-GuideParagraph $Sheet $Row ([string]$script:HelpFormulaExplanations[$key]))
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
        return (Add-GuideFormulaExplanation $Sheet ($Row + 1) $FallbackText)
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
    return (Add-GuideFormulaExplanation $Sheet ($Row + 1) $FallbackText)
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
    $fallback = $Prefix + " (" + $Numerator + ") / (" + $Denominator + ") " + $Suffix
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

    if (-not [string]::IsNullOrWhiteSpace($Reference) -and -not $FallbackText.Contains($Reference)) { $FallbackText += "      " + $Reference }
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

    $row = Add-GuideParagraph $Sheet $StartRow "Прямой расчет определяет НДС сечения при заданных N, Mx и My и проверяет равновесие и предельные деформации материалов. СП 63, пп. 8.1.20–8.1.24 и 8.1.29–8.1.30."
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
    $row = Add-GuideParagraph $Sheet $row "ε_0 - деформация в начале координат; κ_x и κ_y - кривизны. Во внутренней системе знаков растяжение положительно, сжатие отрицательно. Гипотеза плоских сечений — СП 63, п. 8.1.20; уравнения НДМ — пп. 8.1.23–8.1.25."
    $row = Add-GuideTitle $Sheet $row "3. Условия равновесия" 13
    $row = Add-GuideParagraph $Sheet $row "Напряжения определяются по деформациям элементов и выбранным диаграммам материалов. Параметры плоскости деформаций уточняются до выполнения трех условий равновесия: СП 63, пп. 8.1.20–8.1.21 и 8.1.24–8.1.25."
    $row = Add-GuideFormula $Sheet $row "N_int = Σσ_i·A_i;    Mx_int = Σσ_i·A_i·y_i;    My_int = Σσ_i·A_i·x_i"
    $row = Add-GuideParagraph $Sheet $row "Сечение разделено на бетонные элементы и стержни. Напряжение σ_i каждого элемента умножается на его площадь A_i; сумма даёт внутреннюю силу N_int. Для моментов тот же вклад умножается на координату центра элемента относительно расчётной точки. Полученные усилия должны уравновесить заданные N, Mx и My. Это дискретная запись равновесия НДМ по СП 63, пп. 8.1.20–8.1.25."
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

    $row = Add-GuideParagraph $Sheet $StartRow "Несущая способность определяется достижением предельной деформации бетона или арматуры на выбранной траектории нагружения. Множитель λ = 1 соответствует заданному сочетанию. Критерии деформаций — СП 63, пп. 8.1.20, 8.1.29–8.1.30; траектория поиска задаётся программой."
    $row = Add-GuideTitle $Sheet $row "1. Расчетный профиль и путь нагружения" 13
    $row = Add-GuideParagraph $Sheet $row "Расчет включается в профиле настройкой Calculation.Strength.Capacity = Yes. Используется модель прочности MaterialModel.Strength.*, общая с прямым расчетом НДС."
    $row = Add-GuideParagraph $Sheet $row "LoadPath в строке сочетания задает масштабируемые составляющие: Mx, My, оба момента, N или все усилия. Auto выбирает путь по ненулевым составляющим; пустое значение принимается как Auto."
    $row = Add-GuideParagraph $Sheet $row "При N + M режим Auto последовательно пробует моментный путь, λ*N и λ*NMxy до получения предельной точки. Найденный предел с λ < 1 завершает поиск. Явно заданный путь не заменяется другим."
    $row = Add-GuideFormula $Sheet $row "SF_capacity = λ_ult"
    $row = Add-GuideParagraph $Sheet $row "SF_capacity - коэффициент запаса по несущей способности; λ_ult - найденный предельный множитель. При λ_ult ≥ 1 заданная нагрузка проходит проверку, при λ_ult < 1 - не проходит."
    $row = Add-GuideCompositeFormula $Sheet $row @(
        @{ Kind = "Text"; Text = "max(" },
        @{ Kind = "Fraction"; Numerator = "|ε_b|"; Denominator = "ε_{b,ult}" },
        @{ Kind = "Text"; Text = "; " },
        @{ Kind = "Fraction"; Numerator = "|ε_s|"; Denominator = "ε_{s,ult}" },
        @{ Kind = "Text"; Text = ") = 1" }
    ) "" "max(|ε_b|/ε_{b,ult}; |ε_s|/ε_{s,ult}) = 1"
    $row = Add-GuideParagraph $Sheet $row "ε_b,ult и ε_s,ult - модули соответствующих предельных деформаций сжатия или растяжения. Максимум определяется по всем активным пределам бетона и арматуры."
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
    $row = Add-GuideParagraph $Sheet $row "Точность UltimateStrain задает настройка Config Capacity.ToleranceStrain. Число итераций обоих способов ограничивает Capacity.SolverMaxIterations. Допуски равновесия задаются на Config настройками Solver.ToleranceN, Solver.ToleranceMx и Solver.ToleranceMy."
    $row = Add-GuideTitle $Sheet $row "4. Настройки пробного поиска" 13
    $row = Add-GuideParagraph $Sheet $row "Для LoadMultiplier на Config настройка Capacity.InitialLambda задает первую положительную пробу, Capacity.MaxLambda - техническую границу расширения поиска, Capacity.ToleranceLambda - точность λ. MaxLambda не увеличивает физическую несущую способность и не ограничивает успешный прямой поиск UltimateStrain."
    $row = Add-GuideParagraph $Sheet $row "Настройка Config Capacity.SearchMethod выбирает уточнение интервала: Bisection делит его пополам; Brent сочетает деление с ускоренными пробами; Secant использует изменение результата между предыдущими пробами. Методы ищут один и тот же физический предел."
    $row = Add-GuideParagraph $Sheet $row "На Config Capacity.BaseLoadSteps задает число ступеней приложения нагрузки одной пробы, Capacity.MaxRetries - дополнительные попытки при численной несходимости. На каждом повторе число ступеней удваивается; численный шаг выбирается автоматически. Значения λ и усилий этой пробы не меняются."
    $row = Add-GuideParagraph $Sheet $row "Настройка Config General.DiagramExtension может помогать найти равновесие пробной нагрузки за пределами диаграммы. Физический критерий по деформациям остается исходным: численное продолжение не прибавляет сечению несущую способность."
    $row = Add-GuideTitle $Sheet $row "5. Результаты поиска" 13
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
    $row = Add-GuideParagraph $Sheet $row "Для сочетания сначала проверяется образование нормальной трещины, затем находится текущее НДС от полной нагрузки по модели с отключённым растянутым бетоном. По этому НДС независимо проверяются продольные трещины. Если образование нормальной трещины достоверно исключено, её ширина не вычисляется; иначе применяются составляющие выбранной ниже методики. Непригодное текущее НДС не позволяет проверить ни ширину, ни продольные трещины. Точные контуры уточняют область взаимодействия, но не изменяют сетку и устойчивость."
    $row = Add-GuideTitle $Sheet $row "СП 63: расчет ширины нормальных трещин" 15
    $row = Add-SP63CrackWidthMethodologyGuide $Sheet $row
    $row = Add-SP35CrackWidthMethodologyGuide $Sheet $row
    return (Add-LongitudinalCrackGuide $Sheet $row)
}

# Объясняет ширину трещины по СП 63: общая зависимость, её составляющие, применимость и результат.
function Add-SP63CrackWidthMethodologyGuide {
    param([object]$Sheet, [int]$StartRow)

    $row = Add-GuideParagraph $Sheet $StartRow "Расчет применяется при SLS.Crack.Code = SP63 и Calculation.Crack.Width = Yes в назначенном профиле. Реализовано продолжительное раскрытие для сочетаний II группы; расчет непродолжительного раскрытия по нескольким состояниям не предусмотрен."
    $row = Add-GuideTitle $Sheet $row "1. Условие проверки и общая формула" 13
    $row = Add-GuideNormativeFormula $Sheet $row "SP63_CrackCondition"
    $row = Add-GuideParagraph $Sheet $row "Найденная ширина a_crc сравнивается с допустимой a_crc,ult. Допуск задайте в SLS.Crack.Allowable в единицах INPUT Length по условиям проекта. В Results программа переведёт найденную ширину в OUTPUT Length; изменение единиц не меняет саму проверку."
    $row = Add-GuideNormativeFormula $Sheet $row "SP63_CrackWidth"
    $row = Add-GuideParagraph $Sheet $row "φ_1, φ_2 и φ_3 учитывают длительность нагрузки, поверхность арматуры и характер работы элемента; ψ_s - неравномерность деформаций арматуры между трещинами. σ_s - напряжение выбранной растянутой арматуры; E_s - ее модуль упругости; l_s - расчетное расстояние между трещинами. СП 63, п. 8.2.15."

    $row = Add-GuideTitle $Sheet $row "2. Трещинообразование и применимость расчета" 13
    $row = Add-GuideParagraph $Sheet $row "Проверка трещинообразования предшествует расчету ширины. Текущее НДС для раскрытия определяется при заданном сочетании без учета растянутого бетона. СП 63, пп. 8.2.14–8.2.16."
    $row = Add-GuideParagraph $Sheet $row "Путь поиска задается столбцом LoadPath в таблице сочетаний на Config; пустое значение означает Auto. Численный способ выбирается настройкой Config SLS.Crack.InitiationSolutionStrategy. Допуск уточнения λ методом LoadMultiplier задается отдельно в SLS.Crack.ToleranceLambda. Отдельного пути только для трещин нет."
    $row = Add-GuideFormula $Sheet $row "λ_crc ≤ 1: трещина образуется;    λ_crc > 1: трещина еще не образуется"
    $row = Add-GuideParagraph $Sheet $row "λ_crc - найденный множитель выбранной части нагрузки в момент образования трещины. λ = 1 соответствует текущему сочетанию. Нормальная трещина может быть также достоверно исключена без найденной пороговой точки; тогда выводится NotCracked."
    $row = Add-GuideParagraph $Sheet $row "Подтвержденный NotCracked исключает расчет ширины: статус проверки N/A, численная ячейка a_crc остается пустой. Это не расчетная ширина, равная нулю. Проверка продольных трещин при этом сохраняется." -Emphasis
    $row = Add-GuideParagraph $Sheet $row "Отказ более высокой пробной нагрузки не означает отказа текущего сочетания. Если допустимое равновесие при λ = 1 не достигает критерия трещины, а более высокая проба достигает другого физического предела, текущее сочетание остается NotCracked."

    $row = Add-GuideTitle $Sheet $row "3. Центральное растяжение" 13
    $row = Add-GuideParagraph $Sheet $row "Центральное растяжение определяется относительно приведенного центра сечения бетон + арматура. При смещении точки приложения, заданном на Config настройками Load.ReferenceOffsetX и Load.ReferenceOffsetY, даже одна сила N может создавать изгиб и требовать общей проверки."
    $row = Add-GuideNormativeFormula $Sheet $row "SP63_TensionCondition"
    $row = Add-GuideParagraph $Sheet $row "N - растягивающая сила; N_crc - сила образования нормальных трещин при центральном растяжении."
    $row = Add-GuideNormativeFormula $Sheet $row "SP63_TensionForce"
    $row = Add-GuideParagraph $Sheet $row "A_red - приведенная площадь для проверки центрального растяжения; R_bt,ser задается в колонке растяжения строки Concrete.R.SLS(II)."
    $row = Add-GuideNormativeFormula $Sheet $row "SP63_Area"
    $row = Add-GuideParagraph $Sheet $row "Обозначения СП: A — площадь бетона; A_s и A_s′ — площади растянутой и сжатой арматуры; α = E_s/E_b. При центральном растяжении в программе суммируется вся продольная арматура. Модули заданы в Concrete.E и Steel.E."
    $row = Add-GuideFractionFormula $Sheet $row "λ_crc =" "N_crc" "N" ""
    $row = Add-GuideParagraph $Sheet $row "Для центрального растяжения множитель равен отношению пороговой силы к текущей. При расчете ширины берется вся площадь бетона за вычетом отверстий и вся фактически растянутая арматура."

    $row = Add-GuideTitle $Sheet $row "4. Образование трещины при изгибе и внецентренной силе" 13
    $row = Add-GuideParagraph $Sheet $row "В общем случае точка трещинообразования определяется достижением предельной растягивающей деформации бетона. При одновременном сжатии и растяжении сечения: СП 63, пп. 8.2.14 и 8.1.30."
    $row = Add-GuideFormula $Sheet $row "ε_{bt,ult} = ε_{bt2}      СП 63, п. 8.2.14 и 8.1.30"
    $row = Add-GuideParagraph $Sheet $row "ε_bt,ult - критерий трещинообразования; ε_bt2 - предельная деформация растянутой ветви диаграммы бетона. СП 63, пп. 8.2.14 и 8.1.30."
    $row = Add-GuideParagraph $Sheet $row "Если все сечение растянуто, но кривизна ненулевая, применяется формула:"
    $row = Add-GuideNormativeFormula $Sheet $row "SP63_TensionStrain"
    $row = Add-GuideParagraph $Sheet $row "ε_bt0 - деформация начала предельного участка растянутой диаграммы; ε_1 и ε_2 - растягивающие деформации противоположных граней. Используется модель MaterialModel.CrackInitiation.* назначенного профиля."
    $row = Add-GuideParagraph $Sheet $row "Настройка Config General.DiagramExtension может помочь численному поиску равновесия, но не увеличивает физический порог образования. На Results сохраняются сила и момент найденной точки, использованный путь и причина неудачи, если точка не найдена."

    $row = Add-GuideTitle $Sheet $row "5. Текущее НДС и площадь взаимодействия" 13
    $row = Add-GuideParagraph $Sheet $row "Для расчета ширины используются напряжения текущего сочетания. В профиле задаются MaterialModel.CrackedState.ValueSet = SLS(II) и MaterialModel.CrackedState.ConcreteTension = Ignore."
    $row = Add-GuideParagraph $Sheet $row "Настройка Config SLS.Crack.TensionZoneMode выбирает бетонную площадь A_bt: FullTension - всю растянутую часть; Effective - эффективную полосу с ограничениями. A_bt означает площадь бетона расчетной растянутой зоны. Ограничения высоты растянутой зоны — СП 63, п. 8.2.17."
    $row = Add-GuideFormula $Sheet $row "h_bt = min(max(x_t; 2a_s); 0.5·h)"
    $row = Add-GuideParagraph $Sheet $row "h_bt - глубина эффективной полосы; x_t - глубина растянутой части до НЛ; a_s - расстояние от оси крайнего выбранного стержня до растянутой границы бетона; h - габарит сечения по нормали к НЛ."
    $row = Add-GuideParagraph $Sheet $row "SLS.Crack.CoverDistanceMode задает способ определения a_s: NearestContour - до ближайшей границы бетона по нормали к НЛ; GlobalExtreme - до дальней растянутой опорной линии всего сечения."
    $row = Add-GuideParagraph $Sheet $row "Если x_t меньше 2a_s, Effective расширяет бетонную полосу до 2a_s, но не глубже 0,5h. Она может зайти за НЛ. Сжатые стержни из этой полосы все равно не входят в расчет раскрытия: выбирается только растянутая арматура. Границы 2a_s и 0,5h — СП 63, п. 8.2.17."
    $row = Add-GuideParagraph $Sheet $row "Площадь считается по пересечению полосы с бетоном и выбранными отверстиями. Правила четырех случаев задания контуров приведены в конце раздела подготовки AutoCAD. Без точного наружного контура граница по сетке приближенная; реальная площадь и инерции элементов для НДС не заменяются."

    $row = Add-GuideTitle $Sheet $row "6. Арматура, напряжение и эквивалентный диаметр" 13
    $row = Add-GuideParagraph $Sheet $row "A_s - сумма площадей растянутых стержней, оси которых попали в расчетную зону. Напряжение этих стержней усредняется по площади:"
    $row = Add-GuideFractionFormula $Sheet $row "σ_s =" "Σ(A_s,i · σ_s,i)" "ΣA_s,i" ""
    $row = Add-GuideParagraph $Sheet $row "Суммирование выполняется по выбранным растянутым стержням; A_s,i и σ_s,i - площадь и текущее напряжение i-го стержня. Усреднение по площади принято в программе как обобщение НДМ для произвольного сечения; отдельной формулой СП 63 оно не является."
    $row = Add-GuideFractionFormula $Sheet $row "d_s,eq =" "Σ(n_i·d_i^2)" "Σ(n_i·d_i)" ""
    $row = Add-GuideParagraph $Sheet $row "d_s,eq - эквивалентный диаметр; n_i - число выбранных стержней диаметра d_i. Для смешанных диаметров применено инженерное дополнение по EN 1992-1-1; общий случай отдельно не раскрыт в СП 63."

    $row = Add-GuideTitle $Sheet $row "7. Расстояние между трещинами" 13
    $row = Add-GuideNormativeFormula $Sheet $row "SP63_CrackSpacing"
    $row = Add-GuideParagraph $Sheet $row "В (8.136) d_s — номинальный диаметр арматуры. Для смешанных диаметров программа подставляет d_s,eq; исходное l_s до ограничений обозначается l_s,0. A_bt и A_s относятся к выбранной расчетной зоне."
    $row = Add-GuideFormula $Sheet $row "l_s = max(l_{s,0}; 10·d_s,eq; 100 мм)"
    $row = Add-GuideParagraph $Sheet $row "Сначала применяется нижняя граница: принятое расстояние l_s не меньше 10 диаметров и 100 мм. СП 63, п. 8.2.17."
    $row = Add-GuideFormula $Sheet $row "l_s = min(l_s; 40·d_s,eq; 400 мм)      СП 63, п. 8.2.17"
    $row = Add-GuideParagraph $Sheet $row "Затем применяется верхняя граница: не больше 40 диаметров и 400 мм. Если d_s,eq больше 40 мм, нижнее и верхнее ограничения противоречат друг другу. Программа использует верхнюю границу 400 мм и явно предупреждает об этом в комментарии выполненного расчета. Нормативные границы — СП 63, п. 8.2.17; разрешение противоречия при d_s,eq > 40 мм принято в программе."

    $row = Add-GuideTitle $Sheet $row "8. Коэффициенты, задаваемые на Config" 13
    $row = Add-GuideParagraph $Sheet $row "SLS.Crack.Phi1 задает φ_1: для продолжительного раскрытия 1,4. SLS.Crack.Phi2 задает φ_2: 0,5 для периодического профиля и 0,8 для гладкой арматуры. Применимость коэффициентов определяется расчетной ситуацией. СП 63, п. 8.2.15."
    $row = Add-GuideParagraph $Sheet $row "Настройка Config SLS.Crack.Phi3Mode = Auto принимает φ_3 = 1,2 при растягивающей продольной силе и 1,0 в остальных реализованных случаях. При User используется число из настройки Config SLS.Crack.Phi3. СП 63, п. 8.2.15."
    $row = Add-GuideParagraph $Sheet $row "Настройка SLS.Crack.PsiMode на листе Config выбирает правило определения ψ_s. User использует число из SLS.Crack.PsiS, если получены пригодные данные трещинообразования. Auto сначала принимает 1 и уточняет коэффициент только при превышении допуска ширины. AlwaysCalc выполняет уточнение всегда, когда оно возможно." -RowCount 2
    $row = Add-GuideNormativeFormula $Sheet $row "SP63_Psi"
    $row = Add-GuideParagraph $Sheet $row "σ_s,crc - среднее по площади напряжение тех же выбранных стержней сразу после образования трещины; σ_s - их среднее при текущем сочетании. Индекс crc обозначает пороговое состояние. Набор стержней не меняется между этими двумя состояниями. Значение ψ_s ограничивается диапазоном 0...1. Нормативная зависимость ψ_s — СП 63, п. 8.2.18; усреднение по площади принято в программе."
    $row = Add-GuideParagraph $Sheet $row "Настройка Config SLS.Crack.SigmaSCrcAveragingMode задает учет стержней, ставших сжатыми в пороговом состоянии. AllSelected учитывает их напряжения со знаком; TensionOnly принимает их напряжения за 0, но сохраняет площади в знаменателе среднего."
    $row = Add-GuideParagraph $Sheet $row "Если пригодной точки образования нет либо выбранная арматура в ней не растянута, но растянута при текущем сочетании, ширина продолжает считаться с ψ_s = 1 во всех режимах. Предупреждение выводится только если это значение действительно использовано в формуле. Допущение ψ_s = 1 предусмотрено СП 63, п. 8.2.15; порядок его применения при непригодной пороговой точке задан программой." -Emphasis
    $row = Add-GuideParagraph $Sheet $row "При подтвержденном NotCracked формула ширины не вызывается, пользовательское SLS.Crack.PsiS не меняется и предупреждения о ψ_s = 1 нет. Отсутствие или физический отказ текущего НДС блокирует обе проверки трещин; ошибочные обязательные настройки требуют исправления."

    $row = Add-GuideTitle $Sheet $row "9. Ширина, запас и результат" 13
    $row = Add-GuideFractionFormula $Sheet $row "SF_crc =" "a_crc,ult" "a_crc" ""
    $row = Add-GuideParagraph $Sheet $row "SF_crc - коэффициент запаса по ширине. Если он меньше 1, допуск превышен и статус - FAIL. Для неприменимой или пропущенной проверки ширина и запас не выводятся. Рассчитанное a_crc = 0 - другой случай, не подтверждение отсутствия трещины."
    $row = Add-GuideParagraph $Sheet $row "Подробный блок Results содержит площадь, состав арматуры, напряжения, коэффициенты и ширину. В сводке выводятся общий статус и минимальный запас; в ResultComment - причины пропуска и предупреждения."

    return $row + 2
}

# Объясняет независимую проверку продольных трещин, общую для обеих методик нормальных трещин.
function Add-LongitudinalCrackGuide {
    param([object]$Sheet, [int]$StartRow)

    $row = Add-GuideTitle $Sheet $StartRow "Продольные трещины: общая проверка по СП 35" 15
    $row = Add-GuideParagraph $Sheet $row "Эта отдельная проверка сжатого бетона выполняется при обеих методиках ширины нормальных трещин: SLS.Crack.Code = SP63 и SP35. Ниже приведены её общий критерий, запас и допустимое напряжение."
    $row = Add-GuideParagraph $Sheet $row "Она контролирует максимальное сжимающее напряжение в бетоне текущего НДС и выполняется как часть Calculation.Crack.Width. Отдельный поиск равновесия для этой проверки не нужен."
    $row = Add-GuideFormula $Sheet $row "σ_c,max ≤ R_b,mc2"
    $row = Add-GuideParagraph $Sheet $row "σ_c,max - максимальное по модулю сжимающее напряжение бетона текущего НДС; R_b,mc2 задается настройкой Concrete.Rb.mc2. Запрет продольных трещин и ограничение напряжений — СП 35, п. 7.100; сопротивления — п. 7.97 и таблицы 7.6–7.7."
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
    $row = Add-GuideParagraph $Sheet $row "Основная зависимость раскрытия и радиус армирования принимаются по СП 35, пп. 7.105, 7.109–7.110. Построение зоны разъясняет Методическое пособие 2017 года, раздел 4.2.3. Геометрические правила программы описаны отдельно там, где норматив не задаёт единственный способ построения; они не выдаются за дополнительные требования СП."

    $row = Add-GuideTitle $Sheet $row "1. Формула ширины и напряжение арматуры" 13
    $row = Add-GuideParagraph $Sheet $row "Для каждой растянутой группы находится ширина нормальной трещины. Общая зависимость связывает напряжение арматуры, её модуль и коэффициент, зависящий от окружающего бетона. Ниже последовательно подготавливаются группы и область взаимодействия, затем определяются составляющие этой зависимости и итоговая ширина."
    $row = Add-GuideNormativeFormula $Sheet $row "SP35_CrackWidth"
    $row = Add-GuideParagraph $Sheet $row "В (7.93) для ненапрягаемой арматуры σ = σ_s, E = E_s. σ_s — наибольшее растягивающее напряжение стержня рассматриваемой группы; E_s задан в Steel.E, колонке растяжения. Δ_cr — допустимая ширина из SLS.Crack.Allowable."
    $row = Add-GuideParagraph $Sheet $row "При ψ в сантиметрах формула дает a_cr в сантиметрах. В Results ширина переводится в OUTPUT Length."
    $row = Add-GuideTitle $Sheet $row "2. Исходные данные и применимость" 13
    $row = Add-GuideParagraph $Sheet $row "На Config задаются SLS.Crack.Code = SP35 и Calculation.Crack.Width = Yes в профиле, назначенном сочетанию."
    $row = Add-GuideParagraph $Sheet $row "Допустимая ширина задается настройкой SLS.Crack.Allowable в INPUT Length по требованиям проекта. Автоматическое назначение допуска не предусмотрено."
    $row = Add-GuideParagraph $Sheet $row "Проверка трещинообразования использует общий LoadPath сочетания; пустое значение принимается как Auto. Способ поиска задается SLS.Crack.InitiationSolutionStrategy, допуск λ для LoadMultiplier - SLS.Crack.ToleranceLambda. Подробные правила приведены в справке соответствующих настроек."
    $row = Add-GuideParagraph $Sheet $row "При увеличении нагрузки проверяется первый достигнутый предел: образование трещины, сжатие бетона, растяжение или сжатие арматуры. Если другой предел наступает раньше трещины, он не становится точкой трещинообразования. При подтвержденном допустимом НДС исходного сочетания без трещины выводится NotCracked."
    $row = Add-GuideParagraph $Sheet $row "Подтвержденный NotCracked исключает расчет ширины: статус N/A, численная ячейка пустая. При Cracked раскрытие определяется по напряжениям текущего сочетания."
    $row = Add-GuideParagraph $Sheet $row "Неудача поиска трещинообразования не блокирует расчет ширины при найденном физически допустимом текущем НДС; причина сохраняется в комментарии. Отсутствие или физическая недопустимость текущего НДС блокирует и ширину, и продольные трещины." -Emphasis
    $row = Add-GuideParagraph $Sheet $row "Проверка продольных трещин выполняется отдельно по сжатому бетону и сохраняется даже при NotCracked. Ее допустимое напряжение задается на Config настройкой Concrete.Rb.mc2."

    $row = Add-GuideTitle $Sheet $row "3. Группирование растянутой арматуры" 13
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
    $row = Add-GuideParagraph $Sheet $row "По СП 35: Коэффициент β учитывает число стержней в группе: 1 стержень - β = 1,00; 2 - 0,85; 3 - 0,75 (таблица 7.26). Для двух и трех стержней таблица описывает группы без просветов между ними. СП 35, п. 7.110 и таблица 7.26."
    $row = Add-GuideParagraph $Sheet $row "Принято в программе: Для группы более трех обычных стержней программа использует β = 0,75. Это принятое обобщение, не отдельное правило СП для такой группы. Напрягаемая арматура и канатные пучки в этой версии не поддерживаются."
    $row = Add-GuideParagraph $Sheet $row "Геометрическая середина между крайними стержнями не используется вместо центра по площадям."
    $row = Add-GuideFormulaPicture $Sheet $row (New-SP35GuideScheme 'Groups') "Группы стержней: число стержней, коэффициент β и положение центра."

    $row = Add-GuideTitle $Sheet $row "4. Состав арматуры и расчетное напряжение" 13
    $row = Add-GuideParagraph $Sheet $row "Группа неделима: центр внутри конечной площади или на ее границе включает всю группу; центр снаружи исключает ее целиком, в том числе отдельные стержни внутри площади. Частичного включения нет." -Emphasis
    $row = Add-GuideParagraph $Sheet $row "Число стержней, площадь арматуры, сумма диаметров и β относятся к полному составу принятой группы. При этом каждый исходный стержень должен находиться в бетоне; расположение в отверстии или за наружной границей является ошибкой геометрии."
    $row = Add-GuideParagraph $Sheet $row "Вышележащие растянутые группы, вошедшие в площадь взаимодействия, участвуют в сумме диаметров. Расчетное напряжение принимается только по рассматриваемой группе: максимальное растягивающее напряжение ее стержней в текущем НДС. Радиус взаимодействия — СП 35, п. 7.110; расчётное растягивающее напряжение — п. 7.105."
    $row = Add-GuideParagraph $Sheet $row "Остальные группы проверяются в отдельных итерациях со своими площадями взаимодействия и напряжениями. Если центр рассматриваемой группы не входит в ее конечную площадь, пропускается только эта проверка."

    $row = Add-GuideTitle $Sheet $row "5. Площадь взаимодействия и бетонная граница" 13
    $row = Add-GuideParagraph $Sheet $row "Для каждой растянутой группы строится отдельная площадь взаимодействия с окружающим бетоном. Все группы рассматриваются последовательно."
    $row = Add-GuideParagraph $Sheet $row "Границы задаются в направлениях вдоль нейтральной линии (НЛ) и по нормали к ней, независимо от поворота сечения."
    $row = Add-GuideParagraph $Sheet $row "Со стороны растянутой поверхности используется фактическая граница бетонной части, содержащей рассматриваемую группу. Вышележащие стержни не заменяют эту границу. Отверстия вычитаются из полученной площади."
    $row = Add-GuideParagraph $Sheet $row "После построения боковых границ и ограничения к НЛ определяется пересечение участка с бетонной областью. Отверстия образуют локальные вырезы без изменения боковых границ за пределами отверстий."

    $row = Add-GuideTitle $Sheet $row "6. Боковые границы" 13
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

    $row = Add-GuideTitle $Sheet $row "7. Ряды и граница в сторону НЛ" 13
    $row = Add-GuideParagraph $Sheet $row "В боковом участке группы объединяются в ряды по координате вдоль нормали к НЛ. Максимальный разнос центров одного ряда задается SLS.Crack.SP35.RowTolerance в INPUT Length; значение по умолчанию 50 мм."
    $row = Add-GuideParagraph $Sheet $row "Ряд 1 - ближайший к растянутой бетонной границе; нумерация возрастает к НЛ. RowTolerance определяет только состав рядов и не изменяет боковые границы." -Emphasis
    $row = Add-GuideParagraph $Sheet $row "По СП 35: Расстояние r к НЛ откладывается от оси последнего растянутого ряда. Если площадь его арматуры строго меньше половины площади каждого другого ряда участка, отсчет переносится на предпоследний полный ряд (п. 7.110, абзац о крайнем ряде)."
    $row = Add-GuideParagraph $Sheet $row "При выборе опорного ряда сравниваются суммы фактических площадей стержней целых принятых групп, без коэффициента β. Условие строгое: равенство половине не требует переноса. Правило половины площади — СП 35, п. 7.110."
    $row = Add-GuideFormula $Sheet $row "r_НЛ = k·d_ряд"
    $row = Add-GuideParagraph $Sheet $row "r_НЛ - расстояние от опорного ряда к НЛ; d_ряд выбирается по всем его стержням согласно SLS.Crack.SP35.RadiusDiameterMode. Множитель k общий с боковыми границами."
    $row = Add-GuideParagraph $Sheet $row "Граница ограничивается НЛ. Опорный ряд определяется в боковом участке бетона; после построения конечной площади состав входящих групп повторно проверяется по их центрам. Ограничение нейтральной осью — СП 35, п. 7.110."
    $row = Add-GuideFormulaPicture $Sheet $row (New-SP35GuideScheme 'Rows') "Ряды нумеруются от бетона к НЛ; r откладывается от выбранного опорного ряда."
    $row = Add-GuideParagraph $Sheet $row "Методическое пояснение: Примеры областей показаны на рисунке 4.42 Методического пособия 2017 года, раздел 4.2.3. Правило переноса отсчета на предпоследний ряд приведено в самом СП 35."

    $row = Add-GuideTitle $Sheet $row "8. Площадь бетона и радиус армирования" 13
    $row = Add-GuideFormula $Sheet $row "A_r = A_{бет,1} + A_{бет,2} + ... - (A_{отв,1} + A_{отв,2} + ...)"
    $row = Add-GuideParagraph $Sheet $row "A_r - площадь бетона конечного участка взаимодействия; A_бет,1, A_бет,2, ... - площади его бетонных частей до вычитания отверстий; A_отв,1, A_отв,2, ... - площади отверстий в пределах участка."
    $row = Add-GuideParagraph $Sheet $row "Можно учитывать несколько отверстий и несколько бетонных частей. В полом сечении рассматривается стенка, где находится выбранная группа; противоположная стенка через пустоту в эту область не добавляется."
    $row = Add-GuideParagraph $Sheet $row "Если точные границы заданы, учитываются их прямые участки и круговые дуги. Иначе используются границы бетонной сетки. Четыре случая выбора границ перечислены в конце раздела «Подготовка геометрии в AutoCAD и импорт». Для формулы, схемы и экспорта используется одна и та же конечная область."
    $row = Add-GuideParagraph $Sheet $row "P1...P5 и Q1...Q5 в Results и на схеме - контрольные пересечения с противоположными границами по пяти направлениям. Площадь определяется по всему контуру, включая отверстия между контрольными направлениями."
    $row = Add-GuideFormulaPicture $Sheet $row (New-SP35GuideScheme 'Contours') "Отверстия образуют местные вырезы, включая отверстия между контрольными направлениями P/Q."
    $row = Add-GuideNormativeFormula $Sheet $row "SP35_Radius"
    $row = Add-GuideParagraph $Sheet $row "R_r - радиус армирования, расчетная характеристика, не геометрический отступ. Знаменатель учитывает все принятые целые группы: β - коэффициент группы, n - число стержней, d - диаметр."
    $row = Add-GuideParagraph $Sheet $row "Для одинаковых диаметров вклад группы равен β·n·d; для разных - β, умноженному на сумму фактических диаметров группы:"
    $row = Add-GuideFormula $Sheet $row "Σβnd = β_1·D_1 + β_2·D_2 + ..."
    $row = Add-GuideParagraph $Sheet $row "D_1, D_2, ... - суммы диаметров стержней принятых групп; β_1, β_2, ... - коэффициенты этих групп. Индексы обозначают группы, не ряды."
    $row = Add-GuideFormula $Sheet $row "Σβnd = 0.85·(20 + 25) + 1.00·16 = 54.25 мм"
    $row = Add-GuideParagraph $Sheet $row "Пример для группы из двух стержней 20 и 25 мм и одиночного стержня 16 мм. При A_r в мм2 радиус R_r получается в мм. RadiusDiameterMode влияет на геометрические отступы r, но не на сумму фактических диаметров в знаменателе."

    $row = Add-GuideTitle $Sheet $row "9. Центральное растяжение" 13
    $row = Add-GuideParagraph $Sheet $row "По СП 35: При центральном растяжении берется вся площадь бетона за вычетом отверстий. НЛ отсутствует, поэтому локальные боковые границы и отсчет r от ряда не строятся. СП 35, п. 7.110."
    $row = Add-GuideParagraph $Sheet $row "В сумму Σβnd входят все фактически растянутые группы. Напряжение σ_s берется по наиболее растянутому стержню сечения. Правила заданных и неизвестных контуров остаются теми же."
    $row = Add-GuideFormulaPicture $Sheet $row (New-SP35GuideScheme 'MeshCentral') "Без точного наружного контура границы приближенные; при центральном растяжении используется вся площадь бетона."

    $row = Add-GuideTitle $Sheet $row "10. Коэффициент раскрытия и выбор определяющей группы" 13
    $row = Add-GuideCompositeFormula $Sheet $row @(
        @{ Kind = "Text"; Text = "ψ = 1,5 · " }, @{ Kind = "Radical"; Radicand = "R_r" },
        @{ Kind = "Text"; Text = "  (Periodic);    ψ = 0,35·R_r  (Smooth)" }
    ) "СП 35, п. 7.109" "ψ = 1,5·√R_r (Periodic); ψ = 0,35·R_r (Smooth)"
    $row = Add-GuideParagraph $Sheet $row "ψ - размерный коэффициент раскрытия; R_r и ψ в этих зависимостях выражаются в сантиметрах. SLS.Crack.SP35.RebarProfile: Periodic - периодический профиль, Smooth - гладкая арматура."
    $row = Add-GuideParagraph $Sheet $row "Для арматуры в стенках СП 35, п. 7.109 также предписывает зависимость с квадратным корнем; для неё выберите Periodic."
    $row = Add-GuideFractionFormula $Sheet $row "a_{cr}[мм] =" "σ_s" "E_s" "· ψ[см] · 10"
    $row = Add-GuideParagraph $Sheet $row "Напряжения соседних групп и напряжения в точке трещинообразования в формулу раскрытия не подставляются."
    $row = Add-GuideFormula $Sheet $row "a_{cr} = max(a_{cr,1}; a_{cr,2}; ...)"
    $row = Add-GuideParagraph $Sheet $row "a_cr,1, a_cr,2, ... - результаты отдельных проверок групп. Определяющая группа выбирается по ширине, а не по максимальному напряжению: большая площадь взаимодействия может дать большее раскрытие при меньшем σ_s."
    $row = Add-GuideParagraph $Sheet $row "Применение этих зависимостей разъясняется в Методическом пособии 2017 года, разделе 4.2.3, формулах (4.306)–(4.312). Нормативной основой проверки остаются указанные пункты СП 35."

    $row = Add-GuideTitle $Sheet $row "11. Результаты и экспорт площади взаимодействия" 13
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
    $row = Add-GuideParagraph $Sheet $row "Экспорт идёт в пространство «Модель» активного чертежа, даже если открыт лист. Источник задаёт Results.SnapshotStorage: Excel — лист Results, File — файл Results.SnapshotName рядом с книгой. После изменения нагрузок рассчитайте заново. Геометрия передаётся в мм независимо от OUTPUT. «Очистить подписи в AutoCAD» удаляет оформление, сохраняя бетон, арматуру и контуры; без слоёв оформления очистка сразу завершается." -Emphasis -RowCount 2
    $row = Add-GuideParagraph $Sheet $row "На Config включите AutoCAD.Export.ContourEnabled = Yes и выберите настройку AutoCAD.Export.ContourFormat:"
    $row = Add-GuideParagraph $Sheet $row "Region (по умолчанию) - бетонная область со всеми отверстиями на AutoCAD.Common.SectionContourLayer. Отдельные бетонные части экспортируются отдельными Region."
    $row = Add-GuideParagraph $Sheet $row "Polyline - отдельные замкнутые полилинии: наружная граница на AutoCAD.Common.SectionContourLayer, отверстия на AutoCAD.Common.OpeningContourLayer. Прямые участки и дуги сохраняются."
    $row = Add-GuideParagraph $Sheet $row "Если сохранены только отверстия, экспортируются только они на слой отверстий; наружная граница по сетке не создается. Оба формата применяются к сгенерированным и импортированным сечениям и поддерживают обратный импорт."
    $row = Add-GuideParagraph $Sheet $row "Каждый экспорт добавляет объекты. Совпадения не ищутся, старые контуры не удаляются: заранее подготовьте свободное место или уберите ненужные копии. Ответственность за чистоту чертежа при повторном импорте поясняется ниже."
    $row = Add-GuideParagraph $Sheet $row "Для экспорта рассчитанной площади взаимодействия для трещин, принятой в расчете ширины раскрытия, включите на Config AutoCAD.Export.ExportCrackInteractionContour = Yes и задайте AutoCAD.Export.CrackInteractionLayer. На том же листе в таблице «Настройка расчетных профилей», подразделе «Настройки визуализации», строке «Выводимое состояние» (Visualization.State) выберите PreCrackState, PostCrackState или CrackedState в колонке профиля нужного сочетания."
    $row = Add-GuideParagraph $Sheet $row "PreCrackState - состояние перед образованием трещины; PostCrackState - сразу после образования; CrackedState - текущее НДС с отключенным растянутым бетоном. Если в той же строке Visualization.State выбрано StrengthState (прочность) или CapacityState (несущая способность), площадь взаимодействия для трещин не экспортируется."
    $row = Add-GuideParagraph $Sheet $row "Экспортируется принятая в расчете ширины раскрытия площадь взаимодействия для трещин, сохраненная на Results и согласованная с его геометрией. Превышение допустимой ширины не препятствует экспорту ее контура. Актуальность после изменения исходных данных поясняется ниже."
    $row = Add-GuideParagraph $Sheet $row "Кнопка очистки удаляет оформление, но сохраняет бетон, арматуру, наружные контуры и отверстия, включая прежние экспортированные копии. Повторный экспорт сам чертеж не очищает."
    $row = Add-GuideParagraph $Sheet $row "Если бетонная сетка автоматически создана по контурам AutoCAD, сообщение об этом появляется только при формировании сечения. После расчёта и экспорта это сообщение не повторяется. Сведения о происхождении сетки сохраняются в снимке результатов." -Emphasis

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

    $row = Add-GuideParagraph $Sheet $StartRow "Проверка устойчивости выполняется перед расчетом НДС и учитывает случайный эксцентриситет и продольный изгиб по выбранной методике. Поддерживается ненапрягаемая арматура. СП 63, пп. 8.1.2, 8.1.7 и 8.1.15; СП 35, пп. 7.52–7.55."
    $row = Add-GuideTitle $Sheet $row "1. Исходные данные" 13
    $row = Add-GuideParagraph $Sheet $row "Расчет включается настройкой профиля Calculation.Stability.Enabled = Yes. Stability.Code выбирает SP63 или SP35, MaterialModel.Stability.ValueSet - набор характеристик материалов. Для нормативной проверки применяется ULS(I)."
    $row = Add-GuideParagraph $Sheet $row "В таблице «Нагрузки для устойчивости» задайте постоянные + длительные нагрузки для SP63 либо постоянные для SP35. CombinationID связывает их с полным сочетанием. Числа вводятся в тех же единицах INPUT и знаках. СП 63, п. 8.1.15; СП 35, пп. 7.54–7.55."
    $row = Add-GuideParagraph $Sheet $row "Для каждого сочетания с включенной устойчивостью добавьте строку с тем же CombinationID. В этой строке пустые N, Mx и My означают 0, а отсутствие самой строки является ошибкой ввода. Ошибочная формула и текст требуют исправления; длительную часть программа сама не выделяет и полной нагрузкой автоматически не ограничивает." -RowCount 1
    $row = Add-GuideParagraph $Sheet $row "Если устойчивость не проходит, прочность и трещины продолжаются с исходными усилиями для полноты диагностики. Общий статус сохраняет отказ устойчивости. При успешной проверке последующие расчеты получают увеличенные моменты." -Emphasis

    $row = Add-GuideParagraph $Sheet $row "Во всех формулах этого раздела N — положительный модуль сжимающей силы, а не знак N в таблице ввода. Сначала задаются длины, материалы, нагрузки и правила случайного эксцентриситета. После подготовки общих характеристик приведена общая формула каждой методики, а затем — вычисление её составляющих. Числовые индексы главных плоскостей не следует смешивать с индексами нормативных моментов."
    $row = Add-GuideParagraph $Sheet $row "СП 35, п. 7.52, исключает случайный эксцентриситет из расчёта по трещиностойкости. В программе успешный фильтр устойчивости изменяет усилия для всех проверок данного профиля. Поэтому для нормативной проверки трещин назначьте отдельный профиль с Calculation.Stability.Enabled = No; не объединяйте её с устойчивостью в одном профиле. Отключение устойчивости сохраняет заданные пользователем N, Mx и My, включая их перенос к расчётной точке." -Emphasis
    $row = Add-GuideTitle $Sheet $row "2. Длины и общие характеристики до выбора СП" 15
    $row = Add-GuideTitle $Sheet $row "Геометрическая длина, расчётные длины и случайный эксцентриситет" 13
    $row = Add-GuideFormula $Sheet $row "l_0,1 = μ_1·L;    l_0,2 = μ_2·L"
    $row = Add-GuideParagraph $Sheet $row "Геометрическая длина элемента L задаётся в Stability.ElementLength в единице INPUT Length; это длина вдоль элемента, не размер поперечного сечения. Коэффициенты μ_1 и μ_2 задаются независимо в Stability.Mu1 и Stability.Mu2 по условиям закрепления. Их произведения с L — расчётные длины первой и второй главных плоскостей. Например, при L = 8 м, μ_1 = 2 и μ_2 = 1 получаются 16 м и 8 м. Обе расчётные длины используются для гибкости и критических сил."
    $row = Add-GuideParagraph $Sheet $row "Stability.AccidentalEccentricityMode отдельно определяет длину для случайного эксцентриситета. В AutoWithL используется именно заданная геометрическая длина L. В AutoWithμL используется произведение L на коэффициент μ соответствующей плоскости: μ_1·L для первой, μ_2·L для второй. Этот выбор относится только к формуле случайного эксцентриситета: гибкость и критические силы всегда используют расчётные длины. В User автоматическая формула эксцентриситета не применяется — берутся положительные значения Stability.AccidentalEccentricityUser1 и Stability.AccidentalEccentricityUser2. Нормативная геометрическая длина и дополнительный режим по расчётной длине не отождествляются."
    $row = Add-GuideParagraph $Sheet $row "Stability.AccidentalEccentricityPlanes выбирает, где учитывать случайный эксцентриситет. BothPlanes включает его в обеих главных плоскостях. OnlyMomentPlane при моменте только в одной плоскости не добавляет случайный эксцентриситет во вторую; если оба момента ненулевые, обе плоскости получают случайный эксцентриситет. При чистом сжатии без моментов оба режима учитывают случайные эксцентриситеты в обеих плоскостях. Это правило программы не выполняет перебор направлений."
    $row = Add-GuideParagraph $Sheet $row "При ненулевом моменте случайный эксцентриситет направляется в сторону увеличения модуля этого момента. Если момент равен нулю, направление задаётся Stability.ZeroMomentEccentricitySign1 или Stability.ZeroMomentEccentricitySign2: 1 — положительное, −1 — противоположное направление нормали соответствующей плоскости. Stability.SystemType затем определяет объединение эксцентриситетов: Determinate складывает эксцентриситет нагрузки и случайный эксцентриситет; Indeterminate берёт большее по модулю. Эта настройка не выбирает длину вместо AccidentalEccentricityMode."

    $row = Add-GuideTitle $Sheet $row "Главные плоскости и характеристики сечения" 13
    $row = Add-GuideParagraph $Sheet $row "Программа переносит усилия от пользовательской точки к центру приведенного сечения бетон + арматура и поворачивает их в главные оси. Пользовательская точка задается на Config настройками Load.ReferenceOffsetX и Load.ReferenceOffsetY относительно центра бетона без арматуры."
    $row = Add-GuideFormula $Sheet $row "M_1 = M_x,red·cos α + M_y,red·sin α;    M_2 = -M_x,red·sin α + M_y,red·cos α"
    $row = Add-GuideParagraph $Sheet $row "M_1/M_2 - моменты в главных плоскостях; M_x,red/M_y,red - моменты по X/Y относительно приведенного центра. α - угол главной оси 1 от X. Для почти изотропного сечения фиксируется α = 0."
    $row = Add-GuideParagraph $Sheet $row "При α = 0 плоскость 1 соответствует Mx и размеру по Y, плоскость 2 - My и размеру по X. Площадь и инерции для устойчивости берутся по расчетной сетке; точные контуры трещин эту проверку не меняют."
    $row = Add-GuideFormula $Sheet $row "A_red = A_b + (n_s - 1)·A_s"
    $row = Add-GuideParagraph $Sheet $row "A_red - приведенная площадь; A_b - площадь бетона, A_s - площадь арматуры. Вычитается единица, потому что бетон в месте стержней уже входит в сетку."
    $row = Add-GuideFractionFormula $Sheet $row "n_s =" "E_s" "E_b" ""
    $row = Add-GuideParagraph $Sheet $row "Коэффициент приведения n_s определяется по модулям выбранной модели: Steel.E и Concrete.E на Config. СП 35, п. 7.48; СП 63, п. 8.2.12."
    $row = Add-GuideCompositeFormula $Sheet $row @(
        @{ Kind = "Text"; Text = "i = " },
        @{ Kind = "Text"; Text = "√(" },
        @{ Kind = "Fraction"; Numerator = "I_red"; Denominator = "A_red" },
        @{ Kind = "Text"; Text = ")" }
    ) "" "i = √(I_red/A_red)"
    $row = Add-GuideParagraph $Sheet $row "Радиус инерции i связывает приведённый момент инерции I_red с приведённой площадью A_red. По отношению расчётной длины l_0 к этому радиусу получают гибкость. Для каждой главной плоскости используются собственные I_red и i, поэтому результаты двух плоскостей могут различаться."
    $row = Add-GuideFractionFormula $Sheet $row "r =" "I_red" "A_red·c" ""
    $row = Add-GuideParagraph $Sheet $row "r - ядровое расстояние для выбора ветви СП 35; c - расстояние от приведенного центра до соответствующей крайней бетонной границы. r относится к линейному приведенному сечению, а не к точному началу растяжения при нелинейном НДС."

    $row = Add-GuideTitle $Sheet $row "СП 63: расчёт продольного изгиба и устойчивости" 15
    $row = Add-GuideTitle $Sheet $row "1. Расчётный момент и его увеличение" 13
    $row = Add-GuideFormula $Sheet $row "M_design = η·N·e_0"
    $row = Add-GuideParagraph $Sheet $row "M_design — момент с учётом прогиба сжатого элемента. Сжимающая сила N действует с расчётным эксцентриситетом e_0, а коэффициент η увеличивает этот момент из-за продольного изгиба. Далее сначала определяются начальный эксцентриситет и влияние длительной нагрузки, затем относительный эксцентриситет, бетонный коэффициент, жёсткость, критическая сила и коэффициент увеличения момента. Проверка ведётся отдельно в каждой главной плоскости. СП 63, пп. 8.1.2, 8.1.7 и 8.1.15."
    $row = Add-GuideTitle $Sheet $row "2. Расчётный и случайный эксцентриситеты" 13
    $row = Add-GuideCompositeFormula $Sheet $row @(
        @{ Kind = "Text"; Text = "e_a = max(" },
        @{ Kind = "Fraction"; Numerator = "L"; Denominator = "600" },
        @{ Kind = "Text"; Text = "; " },
        @{ Kind = "Fraction"; Numerator = "h"; Denominator = "30" },
        @{ Kind = "Text"; Text = "; 10 мм)" }
    ) "СП 63, п. 8.1.7" "e_a = max(L/600; h/30; 10 мм)"
    $row = Add-GuideParagraph $Sheet $row "e_a - случайный эксцентриситет по СП 63; h - бетонный габарит проверяемой плоскости. СП 63, п. 8.1.7."
    $row = Add-GuideFractionFormula $Sheet $row "e_stat =" "M" "N" ""
    $row = Add-GuideParagraph $Sheet $row "e_stat - эксцентриситет исходной нагрузки в проверяемой главной плоскости. В формулах устойчивости N принимается положительным модулем сжимающей силы."
    $row = Add-GuideParagraph $Sheet $row "На Config настройка Stability.SystemType = Determinate задает статически определимую схему: фактический и случайный эксцентриситеты складываются. Indeterminate - статически неопределимую: берется большее по модулю с направлением расчетного момента. СП 63, п. 8.1.7."
    $row = Add-GuideFormula $Sheet $row "Determinate: e_0 = e_stat + e_a;    Indeterminate: e_0 = sign·max(|e_stat|; |e_a|)"
    $row = Add-GuideParagraph $Sheet $row "e_0 - расчетный эксцентриситет СП 63; sign - направление расчетного момента. Порог обнуления малых моментов задается Calculation.ZeroMomentPerDepth."

    $row = Add-GuideTitle $Sheet $row "3. Влияние длительной нагрузки" 13
    $row = Add-GuideNormativeFormula $Sheet $row "SP63_PhiL"
    $row = Add-GuideParagraph $Sheet $row "φ_l - коэффициент длительности. В этой нормативной формуле M_1 - полный момент относительно наиболее растянутого или наименее сжатого стержня, M_l1 - соответствующая длительная часть. Здесь индекс 1 обозначает эти нормативные моменты, не номер главной плоскости. Для сечения без арматуры программа использует край бетона как проектное обобщение. Нормативное определение моментов относительно арматуры приведено в СП 63, п. 8.1.15."
    $row = Add-GuideParagraph $Sheet $row "Auto определяет φ_l по указанным постоянным и длительным нагрузкам, в пределах 1…2; PhiL2 принимает φ_l = 2. Выбор находится в Stability.PhiLMode. Коэффициент длительности меняет бетонный вклад в D, через него — N_cr и коэффициент η."
    $row = Add-GuideTitle $Sheet $row "4. Относительный эксцентриситет, жёсткость и критическая сила" 13
    $row = Add-GuideFractionFormula $Sheet $row "δ_e =" "e_0" "h" "      СП 63, п. 8.1.15"
    $row = Add-GuideParagraph $Sheet $row "δ_e — относительный эксцентриситет; e_0 в этой нормативной зависимости принимается по модулю. Его границы на Config задают Stability.SP63.DeltaEMin и Stability.SP63.DeltaEMax."
    $row = Add-GuideNormativeFormula $Sheet $row "SP63_Kb"
    $row = Add-GuideParagraph $Sheet $row "Коэффициент k_b уменьшает бетонный вклад E_b·I в условную жёсткость D. В знаменателе учитываются длительность нагрузки через φ_l и относительный эксцентриситет δ_e. Поэтому сначала определяется эксцентриситет силы относительно размера сечения, а длительная часть нагрузки вводится отдельно. СП 63, п. 8.1.15."
    $row = Add-GuideNormativeFormula $Sheet $row "SP63_D"
    $row = Add-GuideParagraph $Sheet $row "Условная жёсткость D складывается из бетонного и арматурного вкладов. E_b и E_s — модули упругости материалов, I и I_s — их моменты инерции в рассматриваемой плоскости. В Results бетонная инерция обозначается также I_b. Коэффициент k_b уже определён на предыдущем шаге; нормативное k_s = 0,7 задаётся в Stability.SP63.Ks. СП 63, п. 8.1.15."
    $row = Add-GuideNormativeFormula $Sheet $row "SP63_Ncr"
    $row = Add-GuideParagraph $Sheet $row "Критическая сила N_cr характеризует влияние продольного изгиба в выбранной главной плоскости. Она зависит от условной жёсткости D и расчётной длины l_0: большая длина уменьшает N_cr. Именно эту силу используют в знаменателе формулы η, а не найденную несущую способность сечения. СП 63, п. 8.1.15."
    $row = Add-GuideTitle $Sheet $row "5. Увеличение момента и применимость проверки" 13
    $row = Add-GuideNormativeFormula $Sheet $row "SP63_Eta"
    $row = Add-GuideParagraph $Sheet $row "η - коэффициент увеличения момента. При гибкости l_0/i ≤ 14 он принимается равным 1, хотя N_cr выводится справочно. При гибкости больше 14 и достижении N_cr действующей силой положительный коэффициент получить нельзя: проверка получает FAIL. При гибкости не больше 14 критическая сила выводится справочно и этот отказ по N_cr не применяется. Порог гибкости — СП 63, п. 8.1.2; коэффициент η — п. 8.1.15."
    $row = Add-GuideTitle $Sheet $row "СП 35: расчёт продольного изгиба и устойчивости" 15
    $row = Add-GuideParagraph $Sheet $row "Сначала сравниваются расчётный эксцентриситет e_c и ядровое расстояние r приведённого сечения. При |e_c| > r определяется увеличенный момент; при |e_c| ≤ r — предельная сжимающая сила. Это две ветви одной проверки, а не одновременное применение двух коэффициентов. СП 35, пп. 7.54–7.55 и 7.69."
    $row = Add-GuideTitle $Sheet $row "1. Случайный эксцентриситет СП 35" 13
    $row = Add-GuideFractionFormula $Sheet $row "e_{c,sl} =" "l" "400" "      СП 35, п. 7.52"
    $row = Add-GuideParagraph $Sheet $row "Полученное e_c,sl — дополнительный эксцентриситет продольной сжимающей силы. Сначала оно объединяется с эксцентриситетом заданного момента по Stability.SystemType, затем произведение силы на принятый эксцентриситет определяет момент для проверки. Поэтому случайный эксцентриситет влияет на расчётные моменты. СП 35, п. 7.52."
    $row = Add-GuideParagraph $Sheet $row "Фактический эксцентриситет равен M/N в рассматриваемой главной плоскости. Для статически определимой схемы случайный эксцентриситет прибавляется в сторону увеличения момента; для неопределимой берётся больший по модулю фактический или случайный эксцентриситет. Выполняемый расчёт получает длину l из выбранного AccidentalEccentricityMode: геометрическую L для AutoWithL либо расчётную μ_1·L/μ_2·L для AutoWithμL. В User случайный эксцентриситет берётся непосредственно из пользовательской ячейки и по длине не вычисляется. Stability.SystemType управляет сложением или сравнением эксцентриситетов, а не этим выбором длины."
    $row = Add-GuideTitle $Sheet $row "2. Общая формула момента при большом эксцентриситете" 13
    $row = Add-GuideFormula $Sheet $row "M_design = η·N·e_c"
    $row = Add-GuideParagraph $Sheet $row "M_design — момент с учётом продольного изгиба; e_c — расчётный эксцентриситет по СП 35. Для коэффициента η сначала находится критическая сила N_cr. Её сравнение с полной силой также ограничивает применимость этой ветви. СП 35, п. 7.54, формулы (7.8), (7.10) и (7.14)."
    $row = Add-GuideTitle $Sheet $row "3. Составляющие критической силы" 13
    $row = Add-GuideNormativeFormula $Sheet $row "SP35_PhiL"
    $row = Add-GuideParagraph $Sheet $row "M - полный момент относительно этой грани, M_l - его постоянная часть. При противоположных знаках: если |e_c| ≥ 0,1h, берется φ_l = 1, иначе 1,05. Настройка Config Stability.PhiLMode = PhiL2 принимает 2. СП 35, п. 7.54."
    $row = Add-GuideNormativeFormula $Sheet $row "SP35_DeltaMin"
    $row = Add-GuideParagraph $Sheet $row "По СП 35 δ = e_c/h, но не меньше δ_min по (7.12); для выбора ветви e_c — модуль эксцентриситета. R_b — сопротивление бетона сжатию в МПа из Config Concrete.R.ULS(I). СП 35, п. 7.54."
    $row = Add-GuideFractionFormula $Sheet $row "δ_p =" "δ" "φ_p" ""
    $row = Add-GuideParagraph $Sheet $row "Коэффициент φ_p учитывает предварительное напряжение в нормативной записи. Проект поддерживает ненапрягаемую арматуру: в этом случае φ_p = 1. Его значение задано в Stability.SP35.PhiP. Величина δ_p — только удобное обозначение отношения δ/φ_p при раскрытии формулы N_cr. СП 35, п. 7.54, формула (7.13)."
    $row = Add-GuideFractionFormula $Sheet $row "k_1 =" "0.11" "0.1 + δ_p" "+ 0.1"
    $row = Add-GuideParagraph $Sheet $row "Величина k_1 объединяет скобку 0,11/(0,1 + δ/φ_p) + 0,1 из формулы N_cr. Это вспомогательное обозначение позволяет показать бетонный вклад в D отдельно от арматурного; нового нормативного коэффициента оно не вводит. СП 35, п. 7.54, формула (7.10)."
    $row = Add-GuideCompositeFormula $Sheet $row @(
        @{ Kind = "Text"; Text = "D = 6.4·E_b·[" },
        @{ Kind = "Fraction"; Numerator = "I_b"; Denominator = "φ_l" },
        @{ Kind = "Text"; Text = "·k_1 + " },
        @{ Kind = "Fraction"; Numerator = "E_s"; Denominator = "E_b" },
        @{ Kind = "Text"; Text = "·I_s]" }
    ) "" "D = 6.4·E_b·[I_b/φ_l·k_1 + (E_s/E_b)·I_s]"
    $row = Add-GuideParagraph $Sheet $row "D — условная жёсткость этой ветви. E_b и E_s — модули упругости бетона и стали; I_b и I_s — моменты инерции бетонной сетки и арматуры в рассматриваемой главной плоскости. φ_l учитывает постоянную часть нагрузки; k_1 — скобка, зависящая от относительного эксцентриситета. СП 35, п. 7.54, формула (7.10)."

    $row = Add-GuideNormativeFormula $Sheet $row "SP35_Ncr"
    $row = Add-GuideParagraph $Sheet $row "В (7.10) n_1 = E_s/E_b; дополнительного множителя π² нет. Вспомогательные δ_p, k_1 и D выше раскрывают эту зависимость для программы: N_cr = D/l_0². Расчётный момент — η·N·e_c, η берётся по (7.8)."
    $row = Add-GuideNormativeFormula $Sheet $row "SP35_Eta"
    $row = Add-GuideParagraph $Sheet $row "При |e_c| > r применяются критическая сила и увеличение момента. Коэффициент длительности уже учтён в критической силе через моменты относительно выбранной грани бетона."
    $row = Add-GuideNormativeFormula $Sheet $row "SP35_NLimit"
    $row = Add-GuideParagraph $Sheet $row "Проверяется отношение N/N_cr. Его предел задается на Config настройкой Stability.SP35.NOverNcrLimit; запас равен этому пределу, умноженному на N_cr/N. При запасе меньше 1 устойчивость не проходит. Нормативный предел 0,7 — СП 35, п. 7.54, формула (7.14)."
    $row = Add-GuideTitle $Sheet $row "4. Малый эксцентриситет: предельная сжимающая сила" 13
    $row = Add-GuideFormula $Sheet $row "N_ult,stab = φ·(R_b·A_b,stab + R_sc·A_s)"
    $row = Add-GuideParagraph $Sheet $row "N_ult,stab - предельная сила по устойчивости; R_sc задается в Steel.R.ULS(I). A_b,stab - бетонная площадь с поправкой данной ветви; A_s - площадь арматуры. Условие для ненапрягаемой арматуры следует из СП 35, п. 7.69, перечисление а), формула (7.31)."
    $row = Add-GuideFractionFormula $Sheet $row "q_l =" "N_l" "N" ""
    $row = Add-GuideParagraph $Sheet $row "q_l - доля постоянной сжимающей силы из таблицы нагрузок для устойчивости. Она не является табличным коэффициентом φ_l."
    $row = Add-GuideFractionFormula $Sheet $row "φ_{ml} =" "φ_m" "φ_l" ""
    $row = Add-GuideParagraph $Sheet $row "φ_m и φ_l принимаются по таблице 7.21; φ_ml - их отношение."
    $row = Add-GuideNormativeFormula $Sheet $row "SP35_Phi"
    $row = Add-GuideParagraph $Sheet $row "φ - итоговый коэффициент снижения несущей способности с учетом постоянной части. При ее отсутствии он равен φ_m, при полностью постоянной нагрузке - φ_l."
    $row = Add-GuideParagraph $Sheet $row "При |e_c| ≤ r коэффициенты берутся из таблицы 7.21 на Config. φ_m - коэффициент кратковременной нагрузки, φ_l - длительной; они интерполируются по гибкости l_0/i и, для φ_m, по отношению |e_c|/r. СП 35, п. 7.55 и таблица 7.21."
    $row = Add-GuideParagraph $Sheet $row "Если A_s больше 3% площади бетона, бетонный вклад уменьшается на A_s. Это поправка только данной формулы, не повторное вычитание арматуры из приведенного сечения. Настройка Config Stability.PhiLMode здесь не заменяет табличный φ_l. СП 35, п. 7.69."

    $row = Add-GuideFractionFormula $Sheet $row "SF_stab =" "N_ult,stab" "N" ""
    $row = Add-GuideParagraph $Sheet $row "SF_stab - коэффициент запаса устойчивости. Если он меньше 1, полная сила превышает найденный предел."
    $row = Add-GuideTitle $Sheet $row "Результаты проверки устойчивости" 13
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
    $row = Add-GuideParagraph $Sheet $row "Лист Results показывает сохранённые результаты, включая не прошедшие проверки. Схема и экспорт читают лист или файл по настройке Results.SnapshotStorage. Изменение настроек Config не обновляет сохранённые результаты." -RowCount 1 -Emphasis
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
    $row = Add-GuideParagraph $Sheet $row "В Excel подробная таблица Results остаётся читаемой. Справа от аннотаций находится rngNDMElementResultIndex: небольшой служебный каталог относительных строк. Схема читает только выбранные участки. Технические ячейки помечены как блокируемые, но защита листа включается только пользователем по желанию. Не сортируйте и не удаляйте части таблицы вручную: для изменения результатов выполните расчёт заново. При чтении книги без индекса он строится один раз по ключевым столбцам; сохраните книгу. Кнопка «Обновить схему и результаты» обновляет только схему. В режиме File снимок с именем Results.SnapshotName хранится рядом с книгой. Та же кнопка читает его для схемы и обновляет справочные таблицы без нового расчёта; полные состояния и границы элементов остаются в файле, на листе — их шапки." -RowCount 6
    $row = Add-GuideParagraph $Sheet $row "Новые файлы NDMSNP06 содержат каталог независимых блоков. Просмотр проверяет каталог и только прочитанные блоки; повреждение другого состояния обнаружится при его чтении, полном восстановлении либо VerifyFile. Форматы NDMSNP04/05 и расширение .ndms читаются без конвертации; для них требуется полное чтение. Файловая привязка и обновление времени не выполняют скрытый полный SHA-проход по NDMSNP06." -RowCount 3
    $row = Add-GuideParagraph $Sheet $row "Сохраняются только фактически полученные состояния. Для схемы и AutoCAD выберите на Config в профиле Visualization.State и Visualization.Quantity: Stress - напряжение, Strain - деформация. Если выбранного состояния нет, программа сообщит причину и не подменит его другим."
    $row++
    $row = Add-GuideTitle $Sheet $row "Индекс сохранённых состояний элементов" 13
    $row = Add-GuideParagraph $Sheet $row "В Excel большая таблица результатов разделена на блоки одного сочетания, профиля и состояния. Таблица rngNDMElementResultIndex показывает, где начинается каждый блок и сколько строк он занимает. Поэтому при смене сочетания или состояния читается нужный блок, а не весь расчёт. В File адресный каталог хранится в самом снимке."
    $row = Add-GuideParagraph $Sheet $row "LoadCase, ProfileId и StateType обозначают сочетание, профиль и состояние. FirstRow — начало относительно именованного якоря таблицы результатов, RowCount — число строк блока. RunID связывает индекс с расчётом; TableRows и TableColumns задают размер основной таблицы. Строка Block описывает один участок. Напряжения, деформации и ElementID хранятся только в самой таблице результатов элементов."
    $row = Add-GuideParagraph $Sheet $row "Защита листа необязательна: программа её не включает и не меняет пароль. Если хотите запретить ручное редактирование результатов и индекса, включите защиту Results со своим паролем; технические ячейки уже помечены как блокируемые. Перед расчётом или очисткой снимите защиту, если она запрещает запись макросам. Чтение и просмотр готовых результатов не требуют снятия или включения защиты."
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
            $InstructionSheet.Cells.Item($row, 1).Value2 = "" # Пояснение читается как связные абзацы, без искусственного списка пунктов.
            $InstructionSheet.Cells.Item($row, 2).Value2 = $line
            $InstructionSheet.Cells.Item($row, 2).IndentLevel = 0
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
        "Круглое сечение Circle" = "CircleGeometry"
        "Клиновидное сечение PointedSection" = "PointedSectionGeometry"
        "Пустотелый прямоугольник HollowRectangle" = "HollowRectangleGeometry"
        "Набор прямоугольников Rectangles (Г-сечение)" = "RectanglesGeometry"
    }
    # Заголовки читаются одним блоком: дальнейшие действия меняют только ссылки
    # уже прочитанных ячеек и не требуют повторных обращений к COM за текстом.
    $lastHeaderRow = $used.Row + $used.Rows.Count + 40
    $headerText = $ConfigSheet.Range($ConfigSheet.Cells.Item(1, 1), $ConfigSheet.Cells.Item($lastHeaderRow, 40)).Value2
    for ($r = 1; $r -le $lastHeaderRow; $r++) {
        for ($c = 1; $c -le 40; $c++) {
            $caption = [string]$headerText[$r, $c]
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
# Ячейки сочетаний остаются обычным вводом; проверка данных прямо ссылается
# на видимую строку Profile.DisplayName, без скрытой копии имён профилей.
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

    $source = $profilesRange.Cells.Item($displayRow, 3).Resize(1, $profileCount)
    $listAddress = '=' + $source.Address($true, $true) # Имена профилей непосредственно из строки Profile.DisplayName (J39:M39).
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

    $Sheet.Cells.Item($HeaderRow - 3, $StartColumn).Value2 = "СП 35.13330.2011, таблица 7.21"
    $Sheet.Range($Sheet.Cells.Item($HeaderRow - 3, $StartColumn), $Sheet.Cells.Item($HeaderRow - 3, $StartColumn + 5)).Merge() | Out-Null
    $Sheet.Cells.Item($HeaderRow - 3, $StartColumn).Font.Bold = $true
    $Sheet.Cells.Item($HeaderRow - 3, $StartColumn).Interior.Color = 15921906
    $Sheet.Cells.Item($HeaderRow - 2, $StartColumn).Value2 = "Ненапрягаемая арматура; крайние строки — продление таблицы."
    $Sheet.Range($Sheet.Cells.Item($HeaderRow - 2, $StartColumn), $Sheet.Cells.Item($HeaderRow - 2, $StartColumn + 5)).Merge() | Out-Null

    $ratioHeader=$Sheet.Range($Sheet.Cells.Item($HeaderRow - 1, $StartColumn + 1), $Sheet.Cells.Item($HeaderRow - 1, $StartColumn + 4))
    $ratioHeader.Merge() | Out-Null
    $ratioHeader.Value2="ec/r"
    $ratioHeader.HorizontalAlignment=-4108
    $ratioHeader.Font.Bold=$true
    $ratioHeader.Interior.Color=$script:ConfigTableHeaderColor

    $headers = @("l_0/i", "φ_m (0)", "φ_m (0.25)", "φ_m (0.50)", "φ_m (1.00)", "φ_l")
    for ($i = 0; $i -lt $headers.Count; $i++) {
        $cell = $Sheet.Cells.Item($HeaderRow, $StartColumn + $i)
        $cell.Value2 = $headers[$i]
        $cell.Font.Bold = $true
        $cell.Interior.Color = $script:ConfigTableHeaderColor
    }

    # Крайние узлы продлевают таблицу постоянными коэффициентами, не добавляя нормативных строк.
    $rows = @(
        @(0, 1.00, 0.90, 0.81, 0.69, 1.00),
        @(14, 1.00, 0.90, 0.81, 0.69, 1.00),
        @(35, 1.00, 0.86, 0.77, 0.65, 0.84),
        @(40, 0.95, 0.83, 0.74, 0.62, 0.79),
        @(48.5, 0.90, 0.79, 0.70, 0.58, 0.70),
        @(55, 0.86, 0.75, 0.66, 0.55, 0.65),
        @(62.5, 0.82, 0.71, 0.62, 0.51, 0.56),
        @(70, 0.78, 0.67, 0.57, 0.48, 0.47),
        @(75, 0.72, 0.60, 0.52, 0.43, 0.41),
        @(83, 0.67, 0.55, 0.47, 0.38, 0.32),
        @(90, 0.62, 0.51, 0.44, 0.35, 0.25),
        @(97, 0.58, 0.49, 0.43, 0.34, 0.20),
        @(105, 0.53, 0.45, 0.39, 0.32, 0.16),
        @(110, 0.48, 0.41, 0.36, 0.31, 0.14),
        @(120, 0.43, 0.36, 0.31, 0.25, 0.10),
        @(130, 0.38, 0.32, 0.28, 0.24, 0.08),
        @(140, 0.35, 0.29, 0.25, 0.21, 0.07),
        @(150, 0.33, 0.28, 0.24, 0.21, 0.06),
        @(1000000, 0.33, 0.28, 0.24, 0.21, 0.06)
    )

    for ($r = 0; $r -lt $rows.Count; $r++) {
        for ($c = 0; $c -lt $headers.Count; $c++) {
            $rowIndex = [int]($HeaderRow + 1 + $r)
            $columnIndex = [int]($StartColumn + $c)
            $Sheet.Cells.Item($rowIndex, $columnIndex).Value2 = [double]$rows[$r][$c]
        }
    }

    Set-WorkbookNameByBounds $Workbook "rngSP35Table721" $Sheet $HeaderRow $StartColumn ($HeaderRow + $rows.Count) ($StartColumn + 5)
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
        @("ContourLineWeight", "2", "", "pt", "Общая толщина наружного контура и границ отверстий на Excel-схеме. Прямые, дуги и окружности используют одно значение > 0; по умолчанию 2 pt."),
        @("ShowSymbol", "-", "Yes", "Yes/No", "Для размеров: Yes — H=800 мм, B=500 мм; No — только 800 мм, 500 мм. Значение, единица и знак приближения сохраняются; подписи арматуры не меняются.")
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
        "Output.CombinationMode" = @("Worst", "User")
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

    $loads = $Workbook.Names.Item("rngLoadCombinations").RefersToRange
    $combinationSource = $loads.Cells.Item(2, 1).Resize($loads.Rows.Count - 1, 1)
    $combinationListAddress = '=' + $combinationSource.Address($true, $true) # Прямой столбец ID без заголовка и отдельного служебного Worst.
    for ($r = 2; $r -le $settingsRange.Rows.Count; $r++) {
        if ([string]$settingsRange.Cells.Item($r, 1).Value2 -eq "Output.CombinationID") {
            $cell = $settingsRange.Cells.Item($r, 2)
            $cell.Validation.Delete()
            $cell.Validation.Add(3, 1, 1, $combinationListAddress)
            $cell.Validation.IgnoreBlank = $true # Поле не обязательно в режиме Worst; User проверяет выбранный ID при действии.
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
    $Sheet.Range("A1:EF260").Font.Name = "Arial"
    $Sheet.Range("A1:EF260").Font.Size = 9
    Apply-ConfigUserInputAlignment $Workbook
    Apply-ConfigCommentColumnAlignment $Workbook
    Apply-ConfigShortColumnAlignment $Workbook
    Apply-ConfigNamedRangeBorders $Workbook

    return @{ Settings = $settingsRange }
}

# Назначает списки только применимым полям аннотаций; содержимое служебных ячеек не определяет применимость.
function Add-PlotAnnotationValidation {
    param([object]$Sheet)

    $validationSources = @{
        "Enabled" = @("Yes", "No")
        "Placement" = @("Outside", "Inside")
        "TextUnits" = @("mm", "pt")
        "LineEnabled" = @("Yes", "No")
        "ShowSymbol" = @("Yes", "No")
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
                $cell = $range.Cells.Item($r, $c)
                $cell.Validation.Delete()
                $applicable = -not (($c -eq 2 -and $paramName -in @('ExtensionLineWeight','ExtensionLineColor','ArrowType','ArrowSize','ShowSymbol')) -or ($c -eq 3 -and $paramName -eq 'LineEnabled'))
                if ($applicable) {
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
