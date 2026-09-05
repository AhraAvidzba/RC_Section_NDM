# скрипт описывает каталог настроек Config и строит таблицы настроек, единиц, геометрии и валидации.

$script:FormulaImageCounter = 0
$script:FormulaImageRows = New-Object System.Collections.Generic.List[object]
$script:FormulaImageDir = Join-Path (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot "../..")) "workbook/output/formula_images"
$script:ConfigGroupHeaderColor = 14277081
$script:ConfigSubgroupHeaderColor = 15921906
$script:ConfigTableHeaderColor = 14277081

function Get-SystemSettingsCatalog {
    @(
        @{ Name = "GeneralSettings"; Title = "[Общие]"; Rows = @(
            @("General.ExecutionReportEnabled", "No", "-", "Записывать пошаговый txt-отчет выполнения расчета в папку с книгой. Файл перезаписывается при каждом новом запуске."),
            @("General.NonCriticalMessagesEnabled", "Yes", "-", "Показывать информационные окна об успешном расчете, обновлении схемы, импорте и экспорте. Ошибки и предупреждения выводятся всегда.")
        )},
        @{ Name = "GeometrySettings"; Title = "[Геометрия и сетка]"; Rows = @(
            @("Geometry.Source", "Generated", "-", "Источник расчетной геометрии: Generated - построить сетку встроенными генераторами; AutoCAD - импортировать только Region из активного чертежа AutoCAD. Единицы AutoCAD всегда считаются мм."),
            @("Geometry.Type", "LShape", "-", "Тип сечения для Geometry.Source = Generated. Поддерживается: RoundedRectangle, Circle или LShape."),
            @("Mesh.Step", "50", "мм", "Базовый шаг квадратной бетонной волоконной сетки для Geometry.Source = Generated. Один и тот же шаг используется по X и Y."),
            @("Mesh.BoundarySubdivisions", "1", "шт", "Количество подъячеек по каждой оси для граничной базовой ячейки при Geometry.Source = Generated; 1 - быстрый режим по центру ячейки."),
            @("Load.ReferenceOffsetX", "0", "мм", "Смещение точки приложения нагрузок по X относительно центра тяжести бетонного сечения. 0 означает, что N, Mx и My заданы относительно центра бетона без учета арматуры."),
            @("Load.ReferenceOffsetY", "0", "мм", "Смещение точки приложения нагрузок по Y относительно центра тяжести бетонного сечения. Рабочие моменты для решателя автоматически переносятся к координатам расчетных элементов.")
        )},
        @{ Name = "SolverSettings"; Title = "[Настройки решателя НДС]"; Rows = @(
            @("Solver.Method", "Newton", "-", "Метод поиска равновесия НДС. Допустимо только Newton или Secant; пустое или неизвестное значение считается ошибкой исходных данных."),
            @("Solver.MaxIterations", "80", "шт", "Применяется к Newton и Secant. Максимальное число итераций на ступень нагрузки; увеличение повышает шанс сходимости, но увеличивает время."),
            @("Solver.LoadSteps", "1", "шт", "Применяется к Newton и Secant. Число ступеней приложения нагрузки в прямом НДМ; больше ступеней обычно устойчивее, но медленнее."),
            @("Solver.DirectState.DiagramExtension", "Yes", "-", "Помогает найти равновесие для перегруженного сочетания и показать FAIL вместо численной несходимости. Несущую способность и трещины не увеличивает."),
            @("Solver.ToleranceN", "0.0001", "Н", "Применяется к Newton и Secant. Абсолютный допуск равновесия по продольной силе в выбранной INPUT-единице силы; значение переводится во внутренние Н."),
            @("Solver.ToleranceMx", "0.0001", "Н*мм", "Применяется к Newton и Secant. Абсолютный допуск равновесия по моменту Mx в выбранной INPUT-единице момента; значение переводится во внутренние Н*мм."),
            @("Solver.ToleranceMy", "0.0001", "Н*мм", "Применяется к Newton и Secant. Абсолютный допуск равновесия по моменту My в выбранной INPUT-единице момента; значение переводится во внутренние Н*мм."),
            @("Solver.LineSearchEnabled", "Yes", "-", "Применяется к Newton и Secant. Разрешает дробить слишком резкий шаг поиска; расчет становится устойчивее, но иногда медленнее."),
            @("Solver.DampingInitial", "1", "-", "Применяется к Newton и Secant. Начальный коэффициент демпфирования шага; меньше значение делает шаг осторожнее."),
            @("Solver.MinLineSearchAlpha", "0.03125", "-", "Применяется к Newton и Secant. Минимальный коэффициент alpha при line search; меньше значение разрешает более сильное дробление шага."),
            @("Solver.MaxDeltaEpsilon0", "0.0005", "-", "Применяется к Newton и Secant. Ограничение приращения epsilon0 за одну итерацию; меньше значение устойчивее, но медленнее."),
            @("Solver.MaxDeltaKappa", "0.00001", "1/мм", "Применяется к Newton и Secant. Ограничение приращения кривизны за одну итерацию; меньше значение устойчивее, но медленнее."),
            @("Solver.SecantMaxRestarts", "2", "шт", "Применяется только к Secant. Максимальное число контролируемых перезапусков приближенной матрицы чувствительности."),
            @("Solver.SecantMinStepNorm", "0.000000000001", "-", "Применяется только к Secant. Минимально допустимая норма secant-шага; слишком малый шаг считается застоем.")
        )},
        @{ Name = "CapacitySettings"; Title = "[Поиск предельной несущей способности]"; Rows = @(
            @("[Общие настройки]", "", "-", ""),
            @("Capacity.SolutionStrategy", "Auto", "-", "Предпочтительная стратегия поиска предельной нагрузки: Auto, UltimateStrain или LoadMultiplier. Для отдельных траекторий программа может выбрать более устойчивую ветку и покажет ее в Results."),
            @("Capacity.MaxLambda", "64", "-", "Актуально для всех методов. В LoadMultiplier задает предел расширения расчетной скобки; в UltimateStrain используется как защитный верхний предел найденного lambda."),
            @("Capacity.SolverMaxIterations", "60", "шт", "Актуально для всех методов. Ограничивает число попыток найти равновесие или предельное состояние."),
            @("[UltimateStrain]", "", "-", ""),
            @("Capacity.ToleranceStrain", "0.00001", "-", "Актуально для UltimateStrain и первой ветки Auto. Абсолютный допуск достижения предельной деформации критического бетонного волокна или стержня."),
            @("[LoadMultiplier]", "", "-", ""),
            @("Capacity.SearchMethod", "Bisection", "-", "Метод одномерного поиска для LoadMultiplier и для повторной попытки в Auto. Поддерживается только Bisection, Brent или Secant."),
            @("Capacity.InitialLambda", "1", "-", "Начальный множитель lambda при поиске верхней границы несущей способности."),
            @("Capacity.ToleranceLambda", "0.01", "-", "Общий допуск одномерного поиска lambdaUltimate для Bisection, Brent и Secant."),
            @("Capacity.MaxRetries", "0", "шт", "Число повторов после численной несходимости пробного расчета в LoadMultiplier."),
            @("Capacity.BaseLoadSteps", "1", "шт", "Базовое число ступеней нагрузки для каждой пробной точки LoadMultiplier.")
        )},
        @{ Name = "StabilitySettings"; Title = "[Продольный изгиб и устойчивость]"; Rows = @(
            @("[Общие настройки]", "", "-", ""),
            @("Stability.Code", "SP63", "-", "Выбор методики учета продольного изгиба: SP63 или SP35. Расчет включается отдельно в профиле."),
            @("Stability.ElementLength", "3000", "мм", "Геометрическая длина элемента l. Расчетная длина по главным плоскостям получается как mu1*l и mu2*l."),
            @("Stability.Mu1", "1", "-", "Коэффициент расчетной длины для первой главной плоскости приведенного сечения."),
            @("Stability.Mu2", "1", "-", "Коэффициент расчетной длины для второй главной плоскости приведенного сечения."),
            @("Stability.SystemType", "Determinate", "-", "Determinate - статический момент плюс случайный эксцентриситет; Indeterminate - момент принимается не меньше случайного."),
            @("Stability.ZeroMomentEccentricitySign1", "1", "-", "Направление случайного эксцентриситета при нулевом моменте в первой главной плоскости: 1 или -1."),
            @("Stability.ZeroMomentEccentricitySign2", "1", "-", "Направление случайного эксцентриситета при нулевом моменте во второй главной плоскости: 1 или -1."),
            @("Stability.PhiLMode", "Auto", "-", "Auto - коэффициент длительности считается по длительной части момента; PhiL2 - принять phi_l = 2."),
            @("[СП 63]", "", "-", ""),
            @("Stability.SP63.Ks", "0.7", "-", "Коэффициент ks при вкладе арматуры в условную жесткость D по СП 63."),
            @("Stability.SP63.DeltaEMin", "0.15", "-", "Нижнее ограничение delta_e в формуле условной жесткости СП 63."),
            @("Stability.SP63.DeltaEMax", "1.5", "-", "Верхнее ограничение delta_e в формуле условной жесткости СП 63."),
            @("[СП 35]", "", "-", ""),
            @("Stability.SP35.PhiP", "1", "-", "Коэффициент phi_p. Для ненапрягаемой арматуры в текущей постановке обычно принимается 1."),
            @("Stability.SP35.NOverNcrLimit", "0.7", "-", "Контрольное ограничение N/Ncr для ветви СП 35 с коэффициентом eta.")
        )},
        @{ Name = "OutputSettings"; Title = "[Вывод и трещины]"; Rows = @(
            @("SLS.Crack.Allowable", "0.3", "мм", "Допустимая ширина раскрытия a_crc,ult, задается пользователем по нормам и условиям эксплуатации."),
            @("SLS.Crack.TensionZoneMode", "Effective", "-", "Зона бетона Abt: Effective - эффективная зона по 2a и 0.5h; FullTension - вся фактически растянутая зона."),
            @("SLS.Crack.Phi1", "1.4", "-", "Коэффициент длительности действия нагрузки для формулы a_crc; в текущей схеме продолжительного раскрытия используется 1.4."),
            @("SLS.Crack.Phi2", "0.5", "-", "Коэффициент поверхности арматуры в формуле a_crc. Пользователь задает значение напрямую: обычно 0.5 для периодического профиля и 0.8 для гладкого."),
            @("SLS.Crack.Phi3Mode", "Auto", "-", "Как принимать phi3: Auto - по характеру продольной силы; User - брать введенное значение SLS.Crack.Phi3."),
            @("SLS.Crack.Phi3", "1", "-", "Пользовательское значение phi3, если SLS.Crack.Phi3Mode = User."),
            @("SLS.Crack.PsiMode", "User", "-", "Как принимать psi_s: User - брать введенное значение SLS.Crack.PsiS; Auto - уточнять только если первая проверка с psi_s=1 не прошла."),
            @("SLS.Crack.PsiS", "1", "-", "Пользовательское значение psi_s, если SLS.Crack.PsiMode = User.")
        )},
        @{ Name = "AutoCADExportSettings"; Title = "[AutoCAD export]"; Rows = @(
            @("AutoCAD.Export.CombinationID", "Worst", "-", "Какое сочетание экспортировать в AutoCAD: Worst - определяющее сочетание из последнего расчета; либо конкретный CombinationID из rngLoadCombinations."),
            @("AutoCAD.Export.NeutralLineEnabled", "Yes", "-", "Выгружать нейтральную линию в AutoCAD: Yes - выводить; No - не выводить."),
            @("AutoCAD.Export.PrincipalAxesEnabled", "Yes", "-", "Выгружать главные центральные оси приведенного сечения в AutoCAD: Yes - выводить; No - не выводить."),
            @("AutoCAD.Export.LoadPointEnabled", "Yes", "-", "Выгружать точку приложения нагрузки в AutoCAD: Yes - выводить; No - не выводить."),
            @("AutoCAD.Export.LabelMode", "NamesAndValues", "-", "Формат текстовых подписей при выгрузке в AutoCAD: ValuesOnly - только выбранная величина профиля; NamesAndValues - имя элемента и значение. При NamesAndValues имена выводятся для бетона и арматуры."),
            @("AutoCAD.Layer.Concrete", "Concrete", "-", "Слой для областей бетонных волокон. Цвет каждой области задается по PhysicalState из Results: сжатие, растяжение или нейтральное состояние."),
            @("AutoCAD.Layer.Rebar", "Reinf", "-", "Слой для всех областей продольной арматуры. Цвет каждого стержня задается по PhysicalState из Results, а не по пользовательскому знаку Stress/Strain."),
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
        @{ Name = "AutoCADImportSettings"; Title = "[AutoCAD import]"; Rows = @(
            @("AutoCAD.Import.ConcreteLayer", "Concrete", "-", "Слой бетонных областей для импорта. Импортируются только AutoCAD Region на этом слое; единицы чертежа считаются мм."),
            @("AutoCAD.Import.RebarLayer", "Reinf", "-", "Слой областей арматуры для импорта. Импортируются только AutoCAD Region на этом слое; каждый импортированный стержень получает внутреннюю метку Rebar."),
            @("AutoCAD.Import.MinArea", "0.000001", "мм2", "Минимальная площадь Region для импорта; области с меньшей площадью игнорируются.")
        )},
        @{ Name = "PlotSettings"; Title = "[Схема сечения]"; Rows = @(
            @("Plot.Enabled", "Yes", "-", "Включить построение схемы сечения на листе Расчет по данным последнего расчета на листе Results."),
            @("Plot.AutoUpdateAfterCalculation", "Yes", "-", "Автоматически обновлять схему после выполнения расчета. Кнопка Обновить схему всегда читает только Results и не запускает расчет."),
            @("Plot.LoadCase", "Worst", "-", "Какое сочетание показывать на схеме: Worst - определяющее сочетание из последнего расчета; либо конкретный CombinationID из rngLoadCombinations."),
            @("Plot.ResultGradient", "Yes", "-", "Включить цветовое различение результата на схеме. Цвет физического состояния берется из PhysicalState, записанного в Results."),
            @("Plot.ResultLabelsEnabled", "No", "-", "Показывать пространственно распределенные численные подписи величины Visualization.Quantity выбранного профиля для бетонных элементов."),
            @("Plot.ResultLabelSpacing", "100", "мм", "Минимальный пространственный шаг между численными подписями результата на схеме."),
            @("Plot.ResultPrecision", "1", "шт", "Количество знаков после запятой для численных подписей и легенды схемы."),
            @("Plot.NeutralLineEnabled", "Yes", "-", "Показывать нейтральную линию выбранного сочетания по Epsilon0, KappaX, KappaY из Results."),
            @("Plot.PrincipalAxesEnabled", "Yes", "-", "Показывать главные центральные оси приведенного сечения."),
            @("Plot.LoadApplicationPointEnabled", "Yes", "-", "Показывать точку приложения нагрузки из последнего расчета."),
            @("Plot.CentroidEnabled", "Yes", "-", "Показывать центр тяжести приведенного сечения."),
            @("Plot.LegendEnabled", "Yes", "-", "Показывать легенду физического состояния и выбранной величины профиля справа от схемы."),
            @("Plot.LegendMode", "Separate", "-", "Separate - отдельные легенды для арматуры и бетона; Common - одна общая легенда. В режиме Common используются цвета Plot.Color.RebarCompression и Plot.Color.RebarTension."),
            @("Plot.Color.RebarCompression", "30,80,220", "RGB", "Цвет максимального сжатия арматуры. В режиме Plot.LegendMode=Common используется как цвет сжатия для всех элементов."),
            @("Plot.Color.RebarTension", "210,30,20", "RGB", "Цвет максимального растяжения арматуры. В режиме Plot.LegendMode=Common используется как цвет растяжения для всех элементов."),
            @("Plot.Color.ConcreteCompression", "125,35,210", "RGB", "Цвет максимального сжатия бетона при Plot.LegendMode=Separate."),
            @("Plot.Color.ConcreteTension", "0,155,85", "RGB", "Цвет максимального растяжения бетона при Plot.LegendMode=Separate; неработающий растянутый бетон остается серым.")
        )}
    )
}

# Возвращает подготовленные данные или справочное значение для дальнейшего шага сборки.
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

# Возвращает подготовленные данные или справочное значение для дальнейшего шага сборки.
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
        @("Concrete.R.ULS(I)", "15.5", "1.10", "МПа", "I ГПС по СП: Сжатие = Rb, Растяжение = Rbt. Пользователь вводит итоговые расчетные сопротивления с нужными коэффициентами условий работы."),
        @("Concrete.R.SLS(II)", "22.0", "1.80", "МПа", "II ГПС по СП: Сжатие = Rb,ser, Растяжение = Rbt,ser. Коэффициенты условий работы учитываются пользователем до ввода."),
        @("Concrete.Rb.mc2", "14.6", "-", "МПа", "Допустимое осевое сжатие бетона для проверки продольных трещин по СП 35; значение вводится уже с учетом mb13 и других нужных коэффициентов."),
        @("Concrete.E", "32500", "32500", "МПа", "Начальный модуль бетона: Eb при сжатии и Ebt при растяжении; по СП 63 Ebt = Eb."),
        @("Concrete.TwoLine.Eb1Red", "0.0015", "0.00008", "-", "Для TwoLine: eps_b1,red при сжатии и eps_bt1,red при растяжении."),
        @("Concrete.ThreeLine.Eb0", "0.002", "0.0001", "-", "Для ThreeLine: eps_b0 и eps_bt0 - деформации начала горизонтального участка."),
        @("Concrete.TwoThreeLine.Eb2", "0.0035", "0.00015", "-", "Предельные деформации eps_b2 и eps_bt2 для TwoLine и ThreeLine.")
    )
}

# Возвращает параметры арматуры, из которых строятся диаграммы I/II ГПС.
# Rsc,ser и Esc оставлены как программное расширение для симметрии.
function Get-SteelMaterialParametersCatalog {
    @(
        @("Steel.R.ULS(I)", "350", "350", "МПа", "I ГПС по СП: Сжатие = Rsc, Растяжение = Rs."),
        @("Steel.R.SLS(II)", "390", "390", "МПа", "II ГПС: Сжатие = Rsc,ser, Растяжение = Rs,ser; Rsc,ser - программный параметр для симметрии."),
        @("Steel.E", "200000", "200000", "МПа", "Модуль арматуры: Esc при сжатии и Es при растяжении; по СП 63 модуль принимается одинаковым."),
        @("Steel.TwoLine.Es2", "0.025", "0.025", "-", "Для TwoLine: eps_sc2 при сжатии и eps_s2 при растяжении."),
        @("Steel.ThreeLine.Es2", "0.015", "0.015", "-", "Для ThreeLine: eps_sc2 при сжатии и eps_s2 при растяжении.")
    )
}

# Возвращает вертикальную таблицу расчетных профилей.
# Строки описывают frontend-контракт профиля, а столбцы PR1/PR2/... являются
# стабильными ProfileId, которые назначаются сочетаниям нагрузок.
function Get-CalculationProfilesCatalog {
    @(
        @{ Caption = "[Общее]"; Key = ""; PR1 = ""; PR2 = ""; PR3 = ""; PR4 = ""; Comment = "" },
        @{ Caption = "Имя профиля"; Key = "Profile.DisplayName"; PR1 = "Прочность"; PR2 = "Трещины"; PR3 = "Полный расчет"; PR4 = "НДС"; Comment = "Короткое имя профиля для пользователя." },
        @{ Caption = "Описание"; Key = "Profile.Description"; PR1 = "НДС по прочности и несущая способность"; PR2 = "Расчет раскрытия нормальных и продольных трещин"; PR3 = "Прочность, capacity и трещины"; PR4 = "Только прямое НДС по прочности"; Comment = "Пояснение, что делает профиль." },

        @{ Caption = "[Запрашиваемые расчеты]"; Key = ""; PR1 = ""; PR2 = ""; PR3 = ""; PR4 = ""; Comment = "" },
        @{ Caption = "НДС по прочности"; Key = "Calculation.Strength.DirectState"; PR1 = "Yes"; PR2 = "No"; PR3 = "Yes"; PR4 = "Yes"; Comment = "Состояние НДС по модели прочности." },
        @{ Caption = "Несущая способность"; Key = "Calculation.Strength.Capacity"; PR1 = "Yes"; PR2 = "No"; PR3 = "Yes"; PR4 = "No"; Comment = "Поиск предельной точки по CapacityLoadPath сочетания." },
        @{ Caption = "Раскрытие трещин"; Key = "Calculation.Crack.Width"; PR1 = "No"; PR2 = "Yes"; PR3 = "Yes"; PR4 = "No"; Comment = "Нормальные трещины и проверка продольных трещин." },
        @{ Caption = "Продольный изгиб"; Key = "Calculation.Stability.Enabled"; PR1 = "No"; PR2 = "No"; PR3 = "No"; PR4 = "No"; Comment = "Включает расчет продольного изгиба и устойчивости перед остальными проверками профиля." },

        @{ Caption = "[Модель устойчивости]"; Key = ""; PR1 = ""; PR2 = ""; PR3 = ""; PR4 = ""; Comment = "" },
        @{ Caption = "Характеристики материалов"; Key = "MaterialModel.Stability.ValueSet"; PR1 = "ULS(I)"; PR2 = "ULS(I)"; PR3 = "ULS(I)"; PR4 = "ULS(I)"; Comment = "Для расчета устойчивости по СП применяются расчетные характеристики I ГПС: бетон Rb/Rbt, арматура Rsc/Rs." },

        @{ Caption = "[Модель прочности]"; Key = ""; PR1 = ""; PR2 = ""; PR3 = ""; PR4 = ""; Comment = "" },
        @{ Caption = "Характеристики материалов"; Key = "MaterialModel.Strength.ValueSet"; PR1 = "ULS(I)"; PR2 = "ULS(I)"; PR3 = "ULS(I)"; PR4 = "ULS(I)"; Comment = "Для прочности по СП используются характеристики I ГПС: бетон Rb/Rbt, арматура Rsc/Rs." },
        @{ Caption = "Диаграмма бетона"; Key = "MaterialModel.Strength.ConcreteDiagram"; PR1 = "TwoLine"; PR2 = "TwoLine"; PR3 = "TwoLine"; PR4 = "TwoLine"; Comment = "СП 63, п. 6.1.23: для прочности применяется двух- или трехлинейная диаграмма бетона." },
        @{ Caption = "Растянутый бетон"; Key = "MaterialModel.Strength.ConcreteTension"; PR1 = "Ignore"; PR2 = "Ignore"; PR3 = "Ignore"; PR4 = "Ignore"; Comment = "СП 63, п. 8.1.20: при расчете прочности растянутый бетон допускается не учитывать." },
        @{ Caption = "Диаграмма арматуры"; Key = "MaterialModel.Strength.SteelDiagram"; PR1 = "TwoLine"; PR2 = "TwoLine"; PR3 = "TwoLine"; PR4 = "TwoLine"; Comment = "СП 63, п. 6.2.13: TwoLine для физического предела текучести, ThreeLine для условного." },

        @{ Caption = "[Модель Mcrc]"; Key = ""; PR1 = ""; PR2 = ""; PR3 = ""; PR4 = ""; Comment = "" },
        @{ Caption = "Характеристики материалов"; Key = "MaterialModel.CrackInitiation.ValueSet"; PR1 = "SLS(II)"; PR2 = "SLS(II)"; PR3 = "SLS(II)"; PR4 = "SLS(II)"; Comment = "Для Mcrc по СП используются характеристики II ГПС: Rb,ser/Rbt,ser и Rs,ser." },
        @{ Caption = "Диаграмма бетона"; Key = "MaterialModel.CrackInitiation.ConcreteDiagram"; PR1 = "ThreeLine"; PR2 = "ThreeLine"; PR3 = "ThreeLine"; PR4 = "ThreeLine"; Comment = "СП 63, п. 6.1.24: для образования трещин основная модель бетона - ThreeLine с растяжением." },
        @{ Caption = "Растянутый бетон"; Key = "MaterialModel.CrackInitiation.ConcreteTension"; PR1 = "UseDiagram"; PR2 = "UseDiagram"; PR3 = "UseDiagram"; PR4 = "UseDiagram"; Comment = "Для Mcrc растянутая ветвь бетона должна учитываться; иначе момент образования трещин не определяется по этой модели." },
        @{ Caption = "Диаграмма арматуры"; Key = "MaterialModel.CrackInitiation.SteelDiagram"; PR1 = "TwoLine"; PR2 = "TwoLine"; PR3 = "TwoLine"; PR4 = "TwoLine"; Comment = "СП 63, п. 6.2.13: TwoLine для физического предела текучести, ThreeLine для условного." },

        @{ Caption = "[Модель НДС с трещинами]"; Key = ""; PR1 = ""; PR2 = ""; PR3 = ""; PR4 = ""; Comment = "" },
        @{ Caption = "Характеристики материалов"; Key = "MaterialModel.CrackedState.ValueSet"; PR1 = "SLS(II)"; PR2 = "SLS(II)"; PR3 = "SLS(II)"; PR4 = "SLS(II)"; Comment = "Для раскрытия трещин по СП используются характеристики II ГПС." },
        @{ Caption = "Диаграмма бетона"; Key = "MaterialModel.CrackedState.ConcreteDiagram"; PR1 = "TwoLine"; PR2 = "TwoLine"; PR3 = "TwoLine"; PR4 = "TwoLine"; Comment = "СП 63, п. 6.1.26: после образования трещин НДС допускается считать по TwoLine или ThreeLine." },
        @{ Caption = "Растянутый бетон"; Key = "MaterialModel.CrackedState.ConcreteTension"; PR1 = "Ignore"; PR2 = "Ignore"; PR3 = "Ignore"; PR4 = "Ignore"; Comment = "Для уже образовавшейся трещины растянутый бетон в НДС не учитывается." },
        @{ Caption = "Диаграмма арматуры"; Key = "MaterialModel.CrackedState.SteelDiagram"; PR1 = "TwoLine"; PR2 = "TwoLine"; PR3 = "TwoLine"; PR4 = "TwoLine"; Comment = "СП 63, п. 6.2.13: TwoLine для физического предела текучести, ThreeLine для условного." },

        @{ Caption = "[Настройки визуализации]"; Key = ""; PR1 = ""; PR2 = ""; PR3 = ""; PR4 = ""; Comment = "" },
        @{ Caption = "Выводимое состояние"; Key = "Visualization.State"; PR1 = "StrengthState"; PR2 = "CrackedState"; PR3 = "StrengthState"; PR4 = "StrengthState"; Comment = "Какое расчетное состояние показывать на схеме и в AutoCAD." },
        @{ Caption = "Выводимая величина"; Key = "Visualization.Quantity"; PR1 = "Stress"; PR2 = "Stress"; PR3 = "Stress"; PR4 = "Strain"; Comment = "Что показывать: напряжения Stress или деформации Strain." }
    )
}

# Возвращает подготовленные данные или справочное значение для дальнейшего шага сборки.
function Get-GeometrySettingsCatalog {
    @(
        @{ RangeName = "rngCircleGeometry"; Title = "Круглое сечение"; Rows = @(
            @("Circle.Diameter", "300", "мм", "Наружный диаметр бетонного круга; по нему строится сетка бетона и габаритная аннотация диаметра."),
            @("Rebar.AxisDistance", "40", "мм", "Расстояние от наружной грани круга до оси первого ряда стержней; это именно as до центра стержня, а не защитный слой до поверхности арматуры."),
            @("Rebar.Count", "8", "шт", "Количество стержней первого ряда, равномерно разложенных по окружности. Если указать 0, первый и дополнительные ряды не создаются."),
            @("Rebar.Diameter", "20", "мм", "Диаметр стержней первого ряда. Пустая ячейка или 0 отключает арматуру по окружности без ошибки."),
            @("Дополнительные ряды арматуры", "", "", ""),
            @("Rebar.Diameter2", "", "мм", "Диаметр второго ряда. Ряд появляется только если есть первый ряд и этот диаметр больше 0."),
            @("Rebar.Diameter3", "", "мм", "Диаметр третьего ряда. Строится от первого ряда; при совпадении направления со вторым рядом перескакивает его, чтобы стержни не легли друг на друга."),
            @("Rebar.Loc2row", "Stacked", "-", "Куда ставить второй ряд относительно каждого стержня первого ряда: Stacked - внутрь по радиусу, SideBySide - по касательной справа по часовой стрелке."),
            @("Rebar.Loc3row", "Stacked", "-", "Куда ставить третий ряд относительно первого ряда. Если выбран тот же путь, что у второго ряда, программа отодвигает третий ряд дальше.")
        )},
        @{ RangeName = "rngRoundedRectangleGeometry"; Title = "Скругленный прямоугольник"; Rows = @(
            @("RoundedRectangle.Width", "300", "мм", "Полная наружная ширина бетонного сечения по X до скругления углов."),
            @("RoundedRectangle.Height", "200", "мм", "Полная наружная высота бетонного сечения по Y до скругления углов."),
            @("RoundedRectangle.RadiusTopLeft", "0", "мм", "Радиус верхнего левого угла; 0 означает обычный прямой угол без скругления."),
            @("RoundedRectangle.RadiusTopRight", "0", "мм", "Радиус верхнего правого угла; применяется только к форме бетона, арматура для этого типа отдельно не генерируется."),
            @("RoundedRectangle.RadiusBottomRight", "0", "мм", "Радиус нижнего правого угла; слишком большой радиус будет ограничен проверками геометрии."),
            @("RoundedRectangle.RadiusBottomLeft", "0", "мм", "Радиус нижнего левого угла; все радиусы задаются независимо друг от друга.")
        )},
        @{ RangeName = "rngLShapeGeometry"; Title = "Г-образное сечение"; FaceTable = $true; Rows = @(
            @("величина размера", "550", "250", "250", "600", "мм", "Габариты двух прямоугольников Г-сечения: H1/B1 - верхняя полка/стенка, H2/B2 - нижняя полка; общая высота равна H1 + H2."),
            @("as_1", "40", "40", "40", "40", "мм", "Расстояние от грани _1 до оси первого ряда стержней. Для H-граней _1 слева, для B-граней _1 сверху."),
            @("as_2", "40", "40", "40", "40", "мм", "Расстояние от грани _2 до оси первого ряда стержней. Для H-граней _2 справа, для B-граней _2 снизу."),
            @("d_1", "32", "32", "32", "32", "мм", "Диаметр первого ряда у грани _1. Пустая ячейка или 0 отключает этот ряд даже при n_1 > 0."),
            @("d_2", "32", "32", "32", "32", "мм", "Диаметр первого ряда у грани _2. Пустая ячейка или 0 отключает этот ряд даже при n_2 > 0."),
            @("n_1", "5", "2", "2", "5", "шт", "Количество стержней первого ряда у грани _1. Если 0, ряд и его дополнительные стержни не создаются."),
            @("n_2", "5", "2", "2", "5", "шт", "Количество стержней первого ряда у грани _2. Если 0, ряд и его дополнительные стержни не создаются."),
            @("t1_1", "80", "80", "80", "80", "мм", "Отступ первого стержня от начала грани _1 по направлению обхода контура по часовой стрелке."),
            @("t2_1", "80", "80", "80", "80", "мм", "Отступ последнего стержня от конца грани _1; вместе с t1_1 задает длину, на которой раскладывается n_1."),
            @("t1_2", "80", "80", "80", "80", "мм", "Отступ первого стержня от начала грани _2 по направлению обхода контура по часовой стрелке."),
            @("t2_2", "80", "80", "80", "80", "мм", "Отступ последнего стержня от конца грани _2; вместе с t1_2 задает длину, на которой раскладывается n_2."),
            @("дополнительные ряды арматуры", "", "", "", "", "-", "Параметры дополнительных рядов. Пустой диаметр означает, что ряд не создается. Суффикс _1 - левая грань для H или верхняя грань для B; _2 - правая грань для H или нижняя грань для B."),
            @("d_2row_1", "", "", "", "", "мм", "Диаметр второго ряда у грани _1. Ряд строится только от существующих стержней первого ряда этой же грани."),
            @("d_2row_2", "", "", "", "", "мм", "Диаметр второго ряда у грани _2. Ряд строится только от существующих стержней первого ряда этой же грани."),
            @("d_3row_1", "", "", "", "", "мм", "Диаметр третьего ряда у грани _1. Если направление совпадает со вторым рядом, третий ряд ставится дальше от первого."),
            @("d_3row_2", "", "", "", "", "мм", "Диаметр третьего ряда у грани _2. Если направление совпадает со вторым рядом, третий ряд ставится дальше от первого."),
            @("loc_2row", "Stacked", "Stacked", "Stacked", "Stacked", "-", "Положение второго ряда: Stacked - внутрь сечения от первого ряда; SideBySide - вдоль грани рядом с первым рядом."),
            @("loc_3row", "Stacked", "Stacked", "Stacked", "Stacked", "-", "Положение третьего ряда относительно первого ряда; при совпадении с loc_2row программа учитывает второй ряд как занятое место."),
            @("bind_2row", "EachBar", "EachBar", "EachBar", "EachBar", "-", "Привязка второго ряда к первому: EachBar - к каждому стержню, EverySecondBar - через один."),
            @("bind_3row", "EachBar", "EachBar", "EachBar", "EachBar", "-", "Привязка третьего ряда к первому: EachBar - к каждому стержню, EverySecondBar - через один.")
        )}
    )
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

# Устанавливает значение, оформление или именованный диапазон в книге через Excel COM.
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

# Возвращает именованные диапазоны, которые физически живут на Config и
# должны выглядеть как отдельные редактируемые таблицы.
function Get-ConfigNamedRangeNames {
    @(
        "rngSystemSettings",
        "rngUnitSettings",
        "rngSignConventionSettings",
        "rngCircleGeometry",
        "rngRoundedRectangleGeometry",
        "rngLShapeGeometry",
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
            if ([string]$range.Worksheet.Name -ne "Config") { continue }

            $range.Borders.LineStyle = -4115
            $range.Borders.Weight = 2
            $range.Borders.Color = 0
        } catch {
            # Диапазон может отсутствовать в промежуточной сборке; валидатор
            # отдельно проверит обязательные имена книги.
        }
    }
}

# Единое правило для колонок "Комментарий/Комментарии" на Config:
# если справа нет других пользовательских колонок, текст прижимается вправо
# и визуально не растекается по пустым ячейкам; если справа есть "Справка",
# комментарий остается слева и читается как обычное описание.
function Apply-ConfigCommentColumnAlignment {
    param([object]$Workbook)

    foreach ($rangeName in (Get-ConfigNamedRangeNames)) {
        try {
            $range = $Workbook.Names.Item($rangeName).RefersToRange
            if ([string]$range.Worksheet.Name -ne "Config") { continue }
            if ($range.Rows.Count -lt 2) { continue }

            for ($c = 1; $c -le $range.Columns.Count; $c++) {
                $header = [string]$range.Cells.Item(1, $c).Value2
                $headerRow = 1
                if (($header -ne "Комментарий") -and ($header -ne "Комментарии") -and ($header -ne "Comment") -and ($range.Rows.Count -ge 2)) {
                    $header = [string]$range.Cells.Item(2, $c).Value2
                    $headerRow = 2
                }
                if ($header -eq "Комментарий" -or $header -eq "Комментарии" -or $header -eq "Comment") {
                    $alignment = -4131
                    if ($c -eq $range.Columns.Count) { $alignment = -4152 }
                    $range.Cells.Item($headerRow, $c).HorizontalAlignment = -4131
                    if ($range.Rows.Count -gt $headerRow) {
                        $range.Offset($headerRow, $c - 1).Resize($range.Rows.Count - $headerRow, 1).HorizontalAlignment = $alignment
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

    foreach ($rangeName in (Get-ConfigNamedRangeNames)) {
        try {
            $range = $Workbook.Names.Item($rangeName).RefersToRange
            if ([string]$range.Worksheet.Name -ne "Config") { continue }

            for ($c = 1; $c -le $range.Columns.Count; $c++) {
                $header = [string]$range.Cells.Item(1, $c).Value2
                if ($header -eq "Ед." -or $header -eq "Справка" -or $header -eq "INTERNAL") {
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
        @{ Name = "rngCircleGeometry"; Columns = @(2) },
        @{ Name = "rngRoundedRectangleGeometry"; Columns = @(2) },
        @{ Name = "rngLoadCombinations"; Columns = @(1, 2, 3, 4, 5, 6) },
        @{ Name = "rngStabilityDurationLoads"; Columns = @(1, 2, 3, 4) },
        @{ Name = "rngSP35Table721"; Columns = @(1, 2, 3, 4, 5, 6, 7, 8) }
    )

    foreach ($item in $items) {
        try {
            $range = $Workbook.Names.Item($item.Name).RefersToRange
            if ([string]$range.Worksheet.Name -ne "Config") { continue }
            if ($range.Rows.Count -lt 2) { continue }

            foreach ($columnIndex in $item.Columns) {
                Set-ConfigInputCellAlignment $range.Offset(1, $columnIndex - 1).Resize($range.Rows.Count - 1, 1)
            }
        } catch {
            # Диапазон может отсутствовать в промежуточной сборке; валидатор
            # отдельно проверит обязательные имена книги.
        }
    }

    try {
        $range = $Workbook.Names.Item("rngLShapeGeometry").RefersToRange
        if ([string]$range.Worksheet.Name -eq "Config" -and $range.Rows.Count -ge 23) {
            $sheet = $range.Worksheet
            $top = $range.Row
            $left = $range.Column
            Set-ConfigInputCellAlignment $sheet.Range($sheet.Cells.Item($top + 1, $left), $sheet.Cells.Item($top + 1, $left + 4))
            Set-ConfigInputCellAlignment $sheet.Range($sheet.Cells.Item($top + 2, $left), $sheet.Cells.Item($top + 2, $left + 3))
            Set-ConfigInputCellAlignment $sheet.Range($sheet.Cells.Item($top + 4, $left + 1), $sheet.Cells.Item($top + 4, $left + 6))
            Set-ConfigInputCellAlignment $sheet.Range($sheet.Cells.Item($top + 5, $left + 1), $sheet.Cells.Item($top + 12, $left + 5))
            Set-ConfigInputCellAlignment $sheet.Range($sheet.Cells.Item($top + 5, $left + 6), $sheet.Cells.Item($top + 12, $left + 6))
            Set-ConfigInputCellAlignment $sheet.Range($sheet.Cells.Item($top + 14, $left + 1), $sheet.Cells.Item($top + 14, $left + 7))
            Set-ConfigInputCellAlignment $sheet.Range($sheet.Cells.Item($top + 15, $left + 1), $sheet.Cells.Item($top + 22, $left + 6))
            Set-ConfigInputCellAlignment $sheet.Range($sheet.Cells.Item($top + 15, $left + 7), $sheet.Cells.Item($top + 22, $left + 7))
        }
    } catch {
        # LShape имеет составную таблицу, поэтому его ячейки ввода форматируются отдельно.
    }
}

# Возвращает расширенное описание настройки для листа "Справка".
function Get-SettingInstructionLines {
    param([string]$Key, [string]$Comment)

    $lead = Get-InstructionLead $Key $Comment

    switch -Wildcard ($Key) {
        "General.ExecutionReportEnabled" { return @($lead) + @(
            "Yes включает запись пошагового отчета выполнения расчета в txt-файл рядом с книгой.",
            "Файл называется RC_Section_NDM_execution_report.txt и перезаписывается при каждом новом запуске расчета, чтобы в папке не копились старые протоколы.",
            "В отчет попадают основные шаги запуска: чтение настроек, построение геометрии, расчет сочетаний, запись Results и обновление схемы.",
            "Для каждого сочетания дополнительно выводятся подробности численного поиска. Поэтому файл может быть большим, если сочетаний много или расчет сходится долго.",
            "No полностью отключает запись файла. Расчетная логика и результаты при этом не меняются."
        ) }
        "General.NonCriticalMessagesEnabled" { return @($lead) + @(
            "Yes показывает обычные информационные окна после успешных действий: расчет выполнен, схема обновлена, геометрия импортирована, экспорт в AutoCAD завершен.",
            "No убирает эти окна, чтобы пакетная работа не требовала каждый раз нажимать OK.",
            "Ошибки, предупреждения и сообщения о невыполненном действии выводятся независимо от этой настройки."
        ) }
        "Geometry.Source" { return @($lead) + @(
            "Generated: программа сама строит бетонную сетку, автоматически расставляет арматуру и готовит подписи для схемы. Этот режим удобен для параметрических сечений Circle, RoundedRectangle и LShape.",
            "AutoCAD: геометрия должна быть заранее загружена отдельной кнопкой 'Импортировать геометрию из AutoCAD'. Кнопка читает только объекты Region на слоях AutoCAD.Import.ConcreteLayer и AutoCAD.Import.RebarLayer, записывает их в Results и сразу показывает предварительную схему без расчетных значений.",
            "При последующем нажатии 'Выполнить расчет' программа не обращается к AutoCAD повторно. Она берет уже сохраненную импортированную геометрию из Results. Если Geometry.Source = AutoCAD, но предварительного импорта нет, расчет останавливается понятным сообщением.",
            "Все координаты и площади AutoCAD считаются заданными в миллиметрах. После импорта программа работает с такой геометрией так же, как с геометрией, построенной внутри книги.",
            "У импортированной из AutoCAD сетки обычно нет сведений о смысловых гранях и группах арматуры. Поэтому программа сохраняет фактические элементы и результаты, но не пытается угадывать специальные подписи граней по координатам."
        ) }
        "Geometry.Type" { return @($lead) + @(
            "Поддерживаемые варианты: Circle, RoundedRectangle, LShape. Для выбранной формы программа строит геометрию, расставляет арматуру и, если умеет, готовит подписи для схемы.",
            "Не используется при Geometry.Source = AutoCAD, потому что в этом режиме фактическая сетка приходит из чертежа.",
            "Если для выбранной формы нет специальных аннотаций, расчет все равно выполняется. На схеме будут показаны геометрия и расчетные значения без этих дополнительных подписей."
        ) }
        "Mesh.Step" { return @($lead) + @(
            "Используется один общий шаг по X и Y. Отдельные Mesh.StepX и Mesh.StepY намеренно не применяются, чтобы бетонные элементы оставались квадратными и визуализация не искажала размер маркеров.",
            "Меньший шаг повышает точность описания контура, площади, центра тяжести, главных осей и распределения напряжений, но увеличивает число бетонных элементов и время расчета.",
            "Формула: A_cell = s², где s = Mesh.Step.",
            "Формула для граничной ячейки: A_boundary = (n_inside / n_total) · s².",
            "Число вводится в текущей INPUT-единице длины. Эквивалентные записи для внутреннего шага 50 mm: 50 при Units.Length.Input = mm, 5 при cm или 0.05 при m.",
            "Mesh.Step - это численный параметр дискретизации НДМ, а не коэффициент СП. Его выбирают по требуемой точности и приемлемому времени расчета."
        ) }
        "Mesh.BoundarySubdivisions" { return @($lead) + @(
            "Значение 1 означает быстрый режим: граничная ячейка проверяется по центру. Значение больше 1 делит граничную ячейку на более мелкие части и считает, какая доля попала внутрь сечения.",
            "Приближенная площадь граничной ячейки: A_boundary = n_inside / n_total * Mesh.Step^2, где n_inside - число подъячеек внутри контура.",
            "Чем больше значение, тем лучше описываются круги, скругления и уступы, но тем дольше строится модель. Для грубого теста обычно достаточно 1, для проверки геометрии лучше увеличивать.",
            "Mesh.BoundarySubdivisions - это численная аппроксимация геометрии. Прямого требования СП к этому параметру нет."
        ) }
        "Circle.Diameter" { return @($lead) + @(
            "Это наружный диаметр бетонной окружности. По нему генератор определяет границу бетона и строит квадратную сетку Mesh.Step.",
            "Центр круга является внутренней технической точкой генератора. Пользовательская точка приложения нагрузок задается отдельно через Load.ReferenceOffsetX/Y относительно центра тяжести бетонного сечения.",
            "Для схемы создается размер диаметра D. Повторять размер круга с нескольких сторон не требуется."
        ) }
        "Rebar.AxisDistance" { return @($lead) + @(
            "Это расстояние от наружной грани круга до оси стержней первого ряда.",
            "Формула радиуса линии арматуры: r₁ = D/2 − as.",
            "Если расстояние слишком мало или уводит стержни за пределы сечения, генератор должен остановиться с ошибкой входных данных."
        ) }
        "Rebar.Count" { return @($lead) + @(
            "Это количество стержней первого ряда, равномерно расставленных по окружности.",
            "Если Rebar.Count = 0 или Rebar.Diameter пустой/равен 0, первый ряд не создается. Второй и третий ряды в этом случае также не создаются, даже если их диаметры заполнены.",
            "Угловой шаг между стержнями: Δφ = 2π / n."
        ) }
        "Rebar.Diameter" { return @($lead) + @(
            "Это диаметр стержней первого ряда круглого сечения.",
            "Площадь одного стержня: As = π·d²/4.",
            "Если диаметр равен 0 или пустой, первый ряд считается отсутствующим и дополнительные ряды не строятся."
        ) }
        "Rebar.Diameter2" { return @($lead) + @(
            "Это опциональный диаметр стержней второго ряда круглого сечения.",
            "Пустое или нулевое значение означает, что второй ряд не создается.",
            "Второй ряд всегда привязывается к уже созданным стержням первого ряда. Если первого ряда нет, второй ряд игнорируется."
        ) }
        "Rebar.Diameter3" { return @($lead) + @(
            "Это опциональный диаметр стержней третьего ряда круглого сечения.",
            "Пустое или нулевое значение означает, что третий ряд не создается.",
            "Третий ряд строится относительно первого ряда. Если второй ряд есть и расположен в том же направлении, третий ряд перескакивает второй, чтобы стержни не совпали."
        ) }
        "Rebar.Loc2row" { return @($lead) + @(
            "Определяет, куда ставить стержень второго ряда относительно каждого стержня первого ряда.",
            "Stacked: стержень ставится внутрь сечения по радиусу к центру.",
            "SideBySide: стержень ставится справа по часовой касательной.",
            "Если второй ряд не задан по диаметру или нет первого ряда, эта настройка не используется."
        ) }
        "Rebar.Loc3row" { return @($lead) + @(
            "Определяет, куда ставить стержень третьего ряда относительно каждого стержня первого ряда.",
            "Stacked: внутрь по радиусу к центру. SideBySide: справа по часовой касательной.",
            "Если второй ряд уже занимает выбранное направление, третий ряд ставится дальше, за вторым рядом."
        ) }
        "RoundedRectangle.*" { return @($lead) + @(
            "Настройка относится к геометрии прямоугольного сечения со скруглениями.",
            "Width и Height задают габариты бетона. RadiusTopLeft, RadiusTopRight, RadiusBottomRight и RadiusBottomLeft задают радиусы соответствующих углов.",
            "Если радиус равен 0, угол остается прямым. Если радиусы несовместимы с габаритами, генератор должен остановиться с ошибкой входных данных.",
            "Бетонная сетка строится квадратными элементами Mesh.Step с уточнением границы через Mesh.BoundarySubdivisions.",
            "Форма и радиусы являются геометрическими исходными данными пользователя. Программа не выполняет отдельную проверку конструктивных требований СП для скруглений."
        ) }
        "Load.ReferenceOffset*" { return @($lead) + @(
            "Если Load.ReferenceOffsetX = 0 и Load.ReferenceOffsetY = 0, заданные N, Mx и My относятся к центру тяжести бетонного сечения без учета продольной арматуры. Техническое начало координат формы, например нижний левый угол LShape, на физический результат не влияет.",
            "Если задан offset, программа учитывает дополнительный момент от смещения точки приложения нагрузки относительно бетонного центра.",
            "Инженерный смысл: осевая сила N при ненулевом offset создает дополнительный момент от эксцентриситета. Поэтому изменение offset при тех же N, Mx, My физически меняет расчетную задачу.",
            "Число вводится в текущей INPUT-единице длины. Например эксцентриситет 100 mm задается как 100 при mm, 10 при cm или 0.1 при m.",
            "Равновесие внутренних усилий НДМ соответствует общей постановке СП 63.13330.2018, п. 8.1.23. Принятая в программе система координат и знаков описана в начале этой справки и на листе Config в блоках Единицы измерения и Система знаков.",
            "Практическая проверка: при чистом N и нулевом offset нагрузка проходит через бетонный центр тяжести. Для сечения с несимметричной арматурой небольшая кривизна может появиться уже физически, потому что приведенный центр и бетонный центр не обязаны совпадать."
        ) }
        "Offset" { return @($lead) + @(
            "Offset остается геометрическим отступом в миллиметрах сечения и не зависит от TextUnits.",
            "Для размерных линий значение задает расстояние от грани или линии размера до самой размерной линии.",
            "Для подписей арматуры значение задает расстояние от линии осей стержней до текста подписи.",
            "Если нужно, чтобы при одинаковой области схемы текст не менял визуальный размер для разных габаритов сечения, меняй не Offset, а TextUnits на pt для TextHeight и TextGap."
        ) }
        "TextUnits" { return @($lead) + @(
            "mm - TextHeight и TextGap задаются в реальных миллиметрах сечения. Это удобно, когда оформление должно оставаться пропорциональным самой модели.",
            "pt - TextHeight и TextGap задаются в обычных пунктах Excel. Это удобно для сравнения сечений разных габаритов в одной и той же области построения: текст остается примерно одинакового визуального размера.",
            "Настройка задается отдельно для двух типов аннотаций: в колонке Обозначения арматуры для Plot.RebarLabels.*, в колонке Размерные линии для Plot.Dimensions.*.",
            "Столбец Ед. у строк TextHeight и TextGap меняется формулой по выбранному TextUnits. Если для арматуры и размеров выбраны разные единицы, ячейка показывает обе.",
            "Для стабильной читаемости при разных габаритах сечения удобно выбирать pt. Режим mm нужен для пропорциональной привязки текста к масштабу модели."
        ) }
        "TextHeight" { return @($lead) + @(
            "TextHeight задает размер шрифта текста аннотации.",
            "Если TextUnits=mm, значение сначала переводится из миллиметров сечения в points текущей PlotArea. Поэтому при одном и том же окне схемы крупное сечение даст визуально меньший текст, а маленькое - больший.",
            "Если TextUnits=pt, значение сразу считается размером шрифта Excel. В этом режиме текст почти не зависит от реальных габаритов сечения, что обычно удобнее для единообразной читаемости.",
            "Настройка не меняет расчет и сохраненные результаты. Она применяется при каждом обновлении схемы по текущему Config."
        ) }
        "TextGap" { return @($lead) + @(
            "TextGap задает зазор от линии аннотации до текстовой подписи.",
            "Для размерных линий это расстояние от размерной линии до текста размера. Для арматуры - расстояние от короткой линии обозначения или anchor-позиции подписи до текста.",
            "Если TextUnits=mm, зазор масштабируется как реальная величина сечения. Если TextUnits=pt, зазор задается как экранное расстояние Excel и выглядит стабильнее при смене размера сечения.",
            "TextGap не заменяет Offset. Offset двигает саму линию аннотации относительно сечения, а TextGap двигает только текст относительно уже поставленной линии."
        ) }
        "Concrete.R.SLS(II)" { return @($lead) + @(
            "Строка задает сопротивления бетона для расчетов II группы: Rb,ser при сжатии и Rbt,ser при растяжении.",
            "Эти значения используются при построении диаграмм Mcrc и CrackedNDS. Для Mcrc растянутый бетон включен фиксированно, для CrackedNDS растянутый бетон выключен фиксированно.",
            "Rbt,ser участвует в построении растянутой ветви диаграммы Mcrc и в центральной ветке образования трещин: N_crc = A_red · Rbt,ser. Для общего Mcrc предельный критерий задается через ε_bt,ult по СП 63, п. 8.2.14 и 8.1.30.",
            "Важно: пользователь вводит уже принятые расчетные значения. Если по СП 35 или условиям проекта к сопротивлениям бетона нужно применить коэффициенты условий работы m..., они должны быть учтены до ввода числа в Config.",
            "Фактические точки диаграммы строятся автоматически по этим параметрам и по таблице 'Настройки диаграмм для расчетов'. Контрольная таблица точек на листе нужна для проверки человеком и в расчет не подставляется.",
            "Значения вводятся в текущей INPUT-единице напряжения и перед расчетом переводятся в МПа."
        ) }
        "Concrete.R.ULS(I)" { return @($lead) + @(
            "Строка задает сопротивления бетона для расчетов I группы: Rb при сжатии и Rbt при растяжении.",
            "Эти значения используются для модели прочности. Они не должны подменяться значениями II группы.",
            "Если для Strength выбран режим растянутого бетона Ignore, положительная ветвь бетонной диаграммы в прочностном расчете отключается. Если выбран UseDiagram, растяжение идет по автоматически построенной ветви.",
            "Фактические предельные деформации Strength берутся из построенной диаграммы, а не из отдельных старых настроек пределов.",
            "Важно: программа не умножает Rb/Rbt на коэффициенты условий работы. В Config нужно вводить уже итоговые значения, которые пользователь считает применимыми для своей задачи."
        ) }
        "Concrete.Rb.mc2" { return @($lead) + @(
            "Concrete.Rb.mc2 задает допустимое сжимающее напряжение бетона для SLS(II)-проверки образования продольных трещин.",
            "Проверка выполняется как часть Calculation.Crack.Width, потому что СП 35 относит образование продольных трещин к расчетам трещиностойкости и требует сравнивать напряжения от нормативных нагрузок.",
            "Если профиль включает расчет трещин, программа берет максимальное по модулю сжимающее напряжение бетонных элементов из уже найденного состояния с трещинами и сравнивает его с Rb,mc2.",
            "Если в найденном НДС нет сжатого бетона, проверка получает статус N/A и не влияет на общий статус сочетания.",
            "Коэффициент запаса в Results записывается формулой Excel: SF_long = Rb,mc2 / sigma_c,max. Если SF_long < 1, статус проверки становится FAIL.",
            "СП 35.13330.2011, п. 7.2 и таблица 7.1 задают область проверки, п. 7.100 формулирует условие по нормативным нагрузкам, таблица 7.6 задает сопротивление, а таблица 7.7 - коэффициенты условий работы.",
            "По принятому правилу проекта пользователь вводит Rb,mc2 уже с учетом mb13 и всех других нужных коэффициентов m.... Программа не определяет эти коэффициенты и не применяет их повторно."
        ) }
        "Concrete.E" { return @($lead) + @(
            "Строка задает начальный модуль бетона при сжатии Eb и растяжении Ebt.",
            "По СП 63.13330.2018, п. 6.1.15 начальный модуль бетона при сжатии и растяжении принимается одинаковым: Ebt = Eb.",
            "Отдельное поле Ebt является программной возможностью задать растянутую ветвь бетона независимо от сжатой.",
            "Eb используется в построении диаграмм, начальной жесткости и приведенной площади A_red для центрального растяжения. Предельный критерий общего Mcrc задается через ε_bt,ult, а не через Rbt,ser / E_b."
        ) }
        "Concrete.*" { return @($lead) + @(
            "Это исходное значение материала бетона в Config. Из таких параметров программа автоматически строит расчетные диаграммы для I и II группы.",
            "Фактическое напряжение бетонного элемента в НДМ определяется уже построенной диаграммой выбранного расчета: Strength, Mcrc или CrackedNDS.",
            "Все пользовательские характеристики бетона считаются уже окончательно принятыми для конкретного расчета. Если по СП 35 нужны коэффициенты условий работы m..., температурные, стадийные или иные поправки, пользователь учитывает их в самом вводимом числе.",
            "Единицы напряжений вводятся в выбранной INPUT единице Stress и перед расчетом переводятся в МПа. Внутри ядра 1 МПа = 1 Н/мм2.",
            "Если Units.Stress.Input = MPa, значения вводятся как привычные МПа. Если выбраны kPa, то 15.5 MPa нужно ввести как 15500. Если Pa - как 15500000. Для kgf/cm2 ориентир: 15.5 MPa примерно 158 kgf/cm2. Для tf/m2 - примерно 1580 tf/m2.",
            "Расчетные диаграммы бетона строятся по СП 63.13330.2018, пп. 6.1.14, 6.1.20-6.1.24 и 6.1.26. Табличные значения сопротивлений пользователь задает сам."
        ) }
        "Steel.R.ULS(I)" { return @($lead) + @(
            "Строка задает сопротивления арматуры для расчетов I группы: Rsc при сжатии и Rs при растяжении.",
            "Эти значения используются при построении диаграммы прочности. Расчет получает уже готовую диаграмму и не выбирает Rsc/Rs самостоятельно.",
            "Для TwoLine и ThreeLine производные деформационные точки строятся автоматически по СП 63.13330.2018, пп. 6.2.11, 6.2.14 и 6.2.15."
        ) }
        "Steel.R.SLS(II)" { return @($lead) + @(
            "Строка задает сопротивления арматуры для расчетов II группы: Rs,ser при растяжении и программный параметр Rsc,ser при сжатии.",
            "СП 63 явно использует Rs,ser, а отдельное обозначение Rsc,ser для этой таблицы не вводит. В программе поле Rsc,ser оставлено для симметричного задания сжатой ветви; принято Rsc,ser = Rs,ser.",
            "Эти значения используются диаграммами Mcrc и CrackedNDS и не должны смешиваться с ULS(I)-сопротивлениями Strength."
        ) }
        "Steel.E" { return @($lead) + @(
            "Строка задает модуль арматуры при сжатии Esc и растяжении Es.",
            "По СП 63.13330.2018, п. 6.2.12 модуль Es принимается одинаковым при растяжении и сжатии: Esc = Es.",
            "Отдельное поле Esc является программной возможностью задать сжатую ветвь арматуры независимо от растянутой.",
            "Es используется в диаграммах арматуры и в формуле раскрытия трещины a_crc."
        ) }
        "Steel.*" { return @($lead) + @(
            "Это исходное значение материала обычной ненапрягаемой арматуры в Config. Программа не выполняет автоматическую нормативную проверку соответствия Rs/Rsc/Es выбранному профилю.",
            "Фактическое напряжение в каждом стержне определяется по диаграмме арматуры выбранного расчета. Тип диаграммы TwoLine/ThreeLine выбирается в отдельной таблице настроек диаграмм.",
            "Единицы Rs/Rsc/Es такие же, как Units.Stress.Input. Например 350 MPa вводится как 350 при MPa, 350000 при kPa, 350000000 при Pa, примерно 3569 при kgf/cm2 или примерно 35691 при tf/m2.",
            "Разные материалы арматуры в одном сечении сейчас не поддерживаются. При AutoCAD import достаточно указать, что Region относится к арматуре; профиль и диаграмма берутся из Config.",
            "Инженерное последствие: если в одном реальном сечении есть разные классы или профили арматуры, текущая версия программы требует отдельного инженерного решения или доработки модели материалов."
        ) }
        "Solver.Method" { return @($lead) + @(
            "Newton решает равновесие через касательную матрицу жесткости: K · Δu = −R, где Δu = [Δε₀; Δκx; Δκy].",
            "Secant использует самостоятельную секущую схему с приближенной матрицей чувствительности и контролируемыми рестартами. Это не скрытый вызов Newton.",
            "Для обычных расчетов и для всех CapacityLoadPath обычно предпочтителен Newton: он чаще быстрее и устойчивее, потому что использует локальную жесткость сечения на текущей плоскости деформаций.",
            "Secant полезен как диагностическая альтернатива, если хочется проверить чувствительность результата к методу равновесия. На нелинейных диаграммах и особенно внутри LoadMultiplier он часто медленнее, потому что может требовать больше пробных пересчетов и рестартов.",
            "Выбор Solver.Method влияет на прямой расчет НДС и на проверки равновесия внутри LoadMultiplier. Для UltimateStrain используется отдельная предельная постановка, но с теми же допусками и ограничениями шага.",
            "Пустое или недопустимое значение дает статус InputErr. Программа не переключается на другой метод молча."
        ) }
        "Solver.MaxIterations" { return @($lead) + @(
            "Это максимум итераций на одной ступени приложения нагрузки.",
            "Если равновесие не найдено за это число итераций, пользовательский статус получает NumFail, а подробная причина остается во внутренней диагностике.",
            "Базовое значение 80 выбрано с запасом для случаев, где приращение деформации ограничено Solver.MaxDeltaEpsilon0. Например, при сильном осевом растяжении программа может идти к равновесию маленькими шагами и 40 итераций бывает недостаточно.",
            "Увеличение может помочь тяжелой нелинейной задаче, но увеличивает время и не исправляет ошибочные исходные данные."
        ) }
        "Solver.LoadSteps" { return @($lead) + @(
            "Это число частей, на которые дробится заданная нагрузка в прямом расчете НДС.",
            "При нескольких ступенях программа сначала ищет равновесие для части нагрузки, затем постепенно доходит до полной N + Mx + My.",
            "Больше ступеней обычно устойчивее для нелинейных диаграмм, но каждая ступень требует своих итераций."
        ) }
        "Solver.DirectState.DiagramExtension" { return @($lead) + @(
            "Эта настройка нужна для прямого расчета состояния сечения от заданного сочетания N + Mx + My.",
            "Если нагрузка немного больше того, что описывает физическая диаграмма материала, обычная диаграмма выходит на последнюю точку и расчет может не найти равновесие. При Yes программа добавляет слабое техническое продолжение за последней физической деформацией, чтобы равновесие можно было найти и явно показать пользователю перегрузку.",
            "Такой результат не считается нормальным расчетным состоянием. Если техническое продолжение действительно понадобилось, в Summary будет FAIL, а на схеме и в AutoCAD появится предупреждение: ВНЕ ФИЗИЧЕСКОЙ ДИАГРАММЫ МАТЕРИАЛА.",
            "Extension не увеличивает несущую способность сечения. Расчет Capacity, образование трещины и расчет раскрытия трещин выполняются только по обычным физическим диаграммам без этого продолжения.",
            "Раскрытие трещин считается только тогда, когда исходное состояние с трещинами найдено внутри физической диаграммы: DirectStateStatus = OK и ExtensionUsed = False. Само включение этой настройки расчет трещин не запрещает."
        ) }
        "Solver.ToleranceN" { return @($lead) + @(
            "Это абсолютный допуск невязки по продольной силе N. Число в Config вводится в текущей INPUT-единице силы, указанной в блоке Units.",
            "Перед расчетом программа переводит этот допуск во внутренние Н и уже там проверяет равновесие.",
            "Критерий по силе считается выполненным, когда |Nint − N| не больше переведенного допуска.",
            "Практический ориентир для строгого расчета: при INPUT Force = N можно ставить около 1; при kN - около 0.001; при tf - около 0.0001. Все три варианта дают примерно один и тот же внутренний допуск порядка 1 Н.",
            "В столбце Ед. эта настройка должна показывать текущую INPUT-единицу силы формулой из rngUnitSettings. Если пользователь меняет Units.Force.Input, подпись единицы у Solver.ToleranceN меняется автоматически.",
            "Меньший допуск строже, но может увеличить число итераций или вызвать NumFail из-за слишком жесткого критерия."
        ) }
        "Solver.ToleranceMx" { return @($lead) + @(
            "Это абсолютный допуск невязки по моменту Mx. Число в Config вводится в текущей INPUT-единице момента, указанной в блоке Units.",
            "Перед расчетом программа переводит этот допуск во внутренние Н*мм и уже там проверяет равновесие.",
            "Критерий считается выполненным, когда |Mxint − Mx| не больше переведенного допуска.",
            "Практический ориентир для строгого расчета: при INPUT Moment = Н*мм можно ставить около 1000; при kN*m - около 0.001; при tf*m - около 0.0001. Все три варианта дают примерно один и тот же внутренний допуск порядка 1000 Н*мм.",
            "В столбце Ед. эта настройка показывает текущую INPUT-единицу момента из блока Units. Это важно для Capacity.SolutionStrategy = LoadMultiplier, потому что для каждой проверяемой точки нужно заново проверить равновесие.",
            "Допуск относится к равновесию решателя, а не к точности вывода момента в таблицу Results."
        ) }
        "Solver.ToleranceMy" { return @($lead) + @(
            "Это абсолютный допуск невязки по моменту My. Число в Config вводится в текущей INPUT-единице момента, указанной в блоке Units.",
            "Перед расчетом программа переводит этот допуск во внутренние Н*мм и уже там проверяет равновесие.",
            "Критерий считается выполненным, когда |Myint − My| не больше переведенного допуска.",
            "Практический ориентир такой же, как для Solver.ToleranceMx: 1000 при Н*мм; 0.001 при kN*m; 0.0001 при tf*m. Это не обязательные нормативные значения, а рабочие численные допуски.",
            "В столбце Ед. эта настройка должна показывать текущую INPUT-единицу момента формулой из rngUnitSettings.",
            "Слишком малое значение может сделать расчет чувствительным к дискретизации сетки и округлениям."
        ) }
        "Solver.LineSearchEnabled" { return @($lead) + @(
            "Line search проверяет, улучшает ли предложенный итерационный шаг равновесие.",
            "Если полный шаг слишком резкий, программа пробует уменьшенный шаг с коэффициентом alpha.",
            "Включение повышает устойчивость на нелинейных диаграммах, но может добавить вычисления внутренних усилий."
        ) }
        "Solver.DampingInitial" { return @($lead) + @(
            "Это начальный множитель итерационной поправки.",
            "Формула шага: u_new = u_old + alpha · Δu, где на первом пробном шаге alpha берется из Solver.DampingInitial.",
            "Значение меньше 1 делает первый шаг осторожнее. Это может улучшить сходимость, но обычно увеличивает число итераций.",
            "Обычный стартовый вариант - 1. Если расчет часто срывается на первом большом шаге, можно пробовать 0.5 или 0.25. Это безразмерный коэффициент, поэтому при смене N/kN/tf, mm/m или MPa/kPa его числовое значение менять не нужно."
        ) }
        "Solver.MinLineSearchAlpha" { return @($lead) + @(
            "Это минимально допустимый коэффициент alpha при дроблении итерационного шага line search.",
            "Если шаг уже уменьшен до этого alpha и равновесие все равно не улучшается, line search считается неуспешным.",
            "Меньшее значение разрешает сильнее дробить шаг и иногда спасает сложный расчет. Но если alpha приходится уменьшать слишком сильно, это часто признак плохой обусловленности, грубой сетки или проблемных диаграмм.",
            "Практический диапазон: 0.1 - грубый и быстрый line search; 0.03125 - рабочее значение; 0.01 - более терпеливый, но медленный режим. Это безразмерная доля шага, от выбранных единиц нагрузки и длины не зависит."
        ) }
        "Solver.MaxDeltaEpsilon0" { return @($lead) + @(
            "Это ограничение приращения средней деформации epsilon0 за одну итерацию.",
            "После вычисления поправки Δepsilon0 программа не разрешает ей быть больше заданного предела.",
            "Ограничение защищает расчет от нереалистичного скачка плоскости деформаций. Это численная защита, а не физический предел деформации материала.",
            "Ориентиры: 0.001 - более свободный шаг; 0.0005 - обычное значение; 0.0001 - осторожный режим для жестких диаграмм. Единица - относительная деформация, то есть безразмерная величина; при смене единиц нагрузки/длины значение не пересчитывается."
        ) }
        "Solver.MaxDeltaKappa" { return @($lead) + @(
            "Это ограничение приращения кривизны kappaX и kappaY за одну итерацию.",
            "После вычисления поправок ΔkappaX и ΔkappaY программа масштабирует шаг, если хотя бы одна кривизна меняется больше разрешенного предела.",
            "Настройка не задает предельную кривизну сечения. Она ограничивает только один итерационный скачок, чтобы плоскость деформаций не перескакивала через устойчивое решение.",
            "Меньшее значение обычно делает расчет спокойнее, но медленнее. Слишком большое значение может привести к срыву line search и пользовательскому статусу NumFail.",
            "Число вводится в текущей INPUT-единице кривизны. Эквивалентные записи одного внутреннего предела шага: 0.00001 при Units.Curvature.Input = 1/mm или примерно 0.01 при Units.Curvature.Input = 1/m."
        ) }
        "Solver.SecantMaxRestarts" { return @($lead) + @(
            "Это максимум перезапусков приближенной матрицы чувствительности в методе Secant.",
            "Перезапуск нужен, когда секущая матрица стала неудачной и дальнейшие шаги не улучшают невязку.",
            "Настройка используется только при Solver.Method = Secant."
        ) }
        "Solver.SecantMinStepNorm" { return @($lead) + @(
            "Это минимальная норма шага метода Secant.",
            "Если рассчитанный шаг становится меньше этого порога, программа считает, что метод застрял и не делает полезного продвижения.",
            "Настройка защищает от бесконечных микрошагов без улучшения равновесия.",
            "Обычно это очень малое безразмерное число: 1E-12 как рабочий уровень, 1E-10 для более грубой остановки, 1E-14 для более терпеливой проверки. От пользовательских единиц нагрузки и длины оно не зависит."
        ) }
        "Capacity.SolutionStrategy" { return @($lead) + @(
            "Capacity.SolutionStrategy задает предпочтительную стратегию поиска предельной точки для траектории, заданной в строке сочетания столбцом CapacityLoadPath. Это не жесткий запрет на другую ветку, потому что в некоторых траекториях программа выбирает более устойчивый путь и отдельно показывает фактический метод в Results.",
            "CapacityLoadPath показывает, какие компоненты нагрузки умножаются на λ: λ*Mx, λ*My, λ*Mxy, λ*N или λ*NMxy. Остальные компоненты остаются постоянными.",
            "Если масштабируется N, то момент от смещения точки приложения N масштабируется вместе с N. Это сохраняет одну и ту же линию действия продольной силы.",
            "Auto: рекомендуемый режим. Для изгибных траекторий программа сначала пробует быстрый UltimateStrain. Если эта задача не сошлась численно, она автоматически повторяет поиск через LoadMultiplier по той же самой λ-траектории.",
            "LoadMultiplier: программа подбирает множитель λ пробными расчетами НДС по общей формуле N = N0 + λ·Nbase, Mx = Mx0 + λ·Mxbase, My = My0 + λ·Mybase.",
            "UltimateStrain: программа напрямую ищет плоскость деформаций, где одновременно выполнены условия выбранной λ-траектории и достигнута одна из предельных деформаций бетона или арматуры.",
            "Важное исключение для силовой осевой траектории: если выбран λ*N или выбран λ*NMxy при нулевых пользовательских Mx/My, программа сразу использует LoadMultiplier. Такой путь масштабирует только продольную силу и моменты ее переноса, поэтому прямая предельная постановка UltimateStrain может быть вырожденной на горизонтальном плато диаграммы.",
            "В Summary это видно в отдельном столбце CapacitySolutionMethod. Он показывает не пожелание из Config, а фактическую ветку, которой программа реально нашла предельную точку: UltimateStrain или LoadMultiplier.",
            "Оба метода используют одну и ту же расчетную модель, одинаковые диаграммы материалов и одинаковый формат результата. Отдельных пользовательских режимов SolveMx/SolveMy/SolveMxy нет.",
            "Практический выбор: Auto удобен для обычной работы. UltimateStrain обычно быстрее для изгибных и смешанных траекторий, а LoadMultiplier устойчивее в вырожденных случаях, особенно при чистой продольной силе.",
            "Самые быстрые траектории обычно λ*Mx, λ*My и λ*Mxy при Capacity.SolutionStrategy = UltimateStrain или Auto. В этих случаях N остается постоянной, программа напрямую ищет предельную плоскость деформаций и затем получает λ из найденного предельного момента. Такой путь не требует многократно запускать полный прямой расчет НДС для разных λ.",
            "Более тяжелые и долгие режимы - λ*N и особенно λ*NMxy, а также любой путь при Capacity.SolutionStrategy = LoadMultiplier. Причина не в типе сечения, а в математической постановке: если масштабируется N, нужно проверять целую силовую траекторию N = N0 + λ·Nbase, Mx = Mx0 + λ·Mxbase, My = My0 + λ·Mybase. Для каждого пробного λ программа заново ищет равновесие, поэтому время растет примерно пропорционально числу пробных точек и количеству внутренних итераций.",
            "λ*NMxy является самым общим вариантом: одновременно меняются N, Mx и My, поэтому прямой UltimateStrain должен удерживать результат на трехмерной линии в пространстве усилий N-Mx-My. Это устойчиво и универсально, но численно тяжелее, чем типовой инженерный случай N + λ*Mxy с постоянной N.",
            "Если нужен быстрый расчет предельного момента при заданной продольной силе, обычно выбирают λ*Mxy или λ*Mx/λ*My. Если нужно именно предельное осевое усилие или масштабирование всего сочетания целиком, выбирают λ*N или λ*NMxy и принимают, что расчет может идти заметно дольше.",
            "Численная несходимость не является физическим разрушением. Если равновесие не найдено, пользовательский статус должен быть NumFail, а не FAIL по физическому пределу."
        ) }
        "Capacity.SearchMethod" { return @($lead) + @(
            "Capacity.SearchMethod работает только там, где фактически используется LoadMultiplier: при явном LoadMultiplier, при повторной попытке из Auto и при осевых траекториях, которые программа сразу считает через LoadMultiplier как более устойчивую постановку.",
            "Для λ*Mx, λ*My и λ*Mxy в обычном режиме Auto/UltimateStrain эта настройка не влияет, пока расчет не перешел в повторный поиск через LoadMultiplier.",
            "Bisection: самый предсказуемый и обычно предпочтительный вариант. Он сохраняет интервал, внутри которого меняется знак целевой функции, и на каждой итерации делит его пополам. Чаще всего он не самый быстрый по числу шагов, зато хорошо переносит неровную нелинейную функцию.",
            "Brent: обычно быстрее Bisection, если функция по λ достаточно гладкая. Метод тоже держит безопасный интервал, но при возможности делает более умный интерполяционный шаг. Если шаг получается рискованным, возвращается к безопасному шагу внутри интервала.",
            "Secant: потенциально быстрый на гладких задачах, но в текущих CapacityLoadPath чаще менее предпочтителен. Он берет две последние проверенные точки lambda и прогнозирует следующую точку; при плохом прогнозе, слишком малом шаге или застое может остановиться с NumFail.",
            "Практически: Bisection - надежная базовая стратегия; Brent - хороший кандидат, если хочется ускорить LoadMultiplier и проверка стабильна; Secant лучше использовать для диагностики, а не как основной режим.",
            "Пустое или недопустимое значение приводит к статусу InputErr. Программа не должна молча переключаться на Bisection или Brent.",
            "Bisection, Brent и Secant - это численные техники поиска предельного множителя. СП не предписывает конкретный алгоритм одномерного поиска для этой реализации."
        ) }
        "Capacity.InitialLambda" { return @($lead) + @(
            "Это первая пробная величина lambda для метода LoadMultiplier на выбранной траектории CapacityLoadPath.",
            "Если при этом lambda сечение еще не достигло предельного состояния, программа расширяет верхнюю границу поиска.",
            "Если начальное значение слишком маленькое, может потребоваться больше пробных расчетов. Если слишком большое, первая проба чаще будет несходящейся или сразу за пределом.",
            "Рекомендуемый старт - 1, потому что он соответствует проверке исходного уровня моментов. Для очень слабых сечений можно поставить 0.5; для заведомо малых нагрузок - 2. Lambda безразмерна, поэтому при переходе с tf*m на kN*m значение не пересчитывается."
        ) }
        "Capacity.ToleranceLambda" { return @($lead) + @(
            "Это точность одномерного поиска предельного множителя lambda в LoadMultiplier.",
            "Bisection, Brent и Secant останавливаются, когда интервал/шаг по lambda становится достаточно малым по этому допуску.",
            "Это допуск именно по коэффициенту lambda, а не по деформациям. Для UltimateStrain используется Capacity.ToleranceStrain.",
            "Практические ориентиры: 0.1 - грубая быстрая оценка; 0.01 - обычный инженерный расчет; 0.001 - более точный, но медленный поиск. Значение безразмерно и не зависит от выбранных единиц N/Mx/My."
        ) }
        "Capacity.MaxRetries" { return @($lead) + @(
            "Это число дополнительных попыток для пробной точки lambda, если внутренний расчет НДС не сошелся.",
            "Повтор обычно выполняется с более осторожным числом ступеней нагрузки, чтобы отличить физический предел от численной проблемы.",
            "Если retries исчерпаны, проба получает NumFail. Такая несходимость не считается разрушением сечения.",
            "Ориентиры: 0 - максимально быстро, без повторов; 1...2 - умеренная страховка; 3...4 - медленнее, но полезно для жестких диаграмм и крупной сетки. Это счетчик попыток, от единиц измерения не зависит."
        ) }
        "Capacity.BaseLoadSteps" { return @($lead) + @(
            "Это базовое число ступеней приложения нагрузки для каждой пробной lambda в LoadMultiplier.",
            "Больше ступеней делает пробный расчет устойчивее, но увеличивает время каждой проверки lambda.",
            "Настройка не используется как физический шаг увеличения нагрузки. Внешний поиск lambda выполняют Bisection, Brent или Secant.",
            "Для проверенного Г-сечения с траекторией λ*N расчет сходится уже при BaseLoadSteps = 1. Поэтому значение больше 1 не является обязательным условием Nult, а только численной страховкой.",
            "Ориентиры: 1 - быстро и обычно достаточно; 2...4 - мягче для нелинейных диаграмм; 8 и больше - только если отдельные пробные точки часто дают NumFail. Это количество ступеней, не размерная величина."
        ) }
        "Capacity.ToleranceStrain" { return @($lead) + @(
            "В UltimateStrain программа не перебирает lambda. Она подбирает плоскость деформаций так, чтобы одновременно были выполнены равновесие по N, сохранено направление Mx/My и один критический элемент дошел до своего предела деформации.",
            "Для силовой осевой траектории λ*N, а также для λ*NMxy без пользовательских моментов, этот параметр обычно не участвует: программа сразу применяет LoadMultiplier, где итоговую точность задает Capacity.ToleranceLambda.",
            "Capacity.ToleranceStrain говорит, насколько близко к этому пределу нужно попасть. Например при пределе бетона -0.0035 и допуске 0.00001 расчет может остановиться, когда критическая деформация отличается от предела примерно на эту величину.",
            "Это не физический предел. Физические пределы берутся из диаграммы выбранного расчета: бетонные пределы задает построенная диаграмма бетона, пределы арматуры - построенная диаграмма арматуры.",
            "Меньший допуск дает более строгую численную остановку, но может увеличить число итераций или привести к NumFail на сложной геометрии/диаграмме. Больший допуск ускоряет расчет, но делает найденный lambda менее точным.",
            "Ориентиры: 0.0001 - грубее и быстрее; 0.00001 - обычное значение; 0.000001 - строгая проверка, которая может потребовать больше итераций. Это относительная деформация, безразмерная величина; при смене единиц длины или нагрузки ее не пересчитывают.",
            "Capacity.ToleranceStrain - это численный допуск реализации. Он не является расчетным коэффициентом или отдельным требованием СП."
        ) }
        "Capacity.MaxLambda" { return @($lead) + @(
            "Это защитная верхняя граница для найденного коэффициента lambda.",
            "В LoadMultiplier программа расширяет интервал поиска только до этого значения. Если предельное состояние не найдено до MaxLambda, расчет должен завершиться диагностируемым статусом.",
            "В UltimateStrain настройка также ограничивает допустимый масштаб результата, чтобы ошибочные исходные данные не приводили к бессмысленно большой lambda.",
            "Ориентиры: 10...20 удобно для быстрой оценки; 64 - рабочее значение; больше 100 обычно имеет смысл только для очень малых исходных моментов. Lambda безразмерна и не зависит от выбранных единиц."
        ) }
        "Capacity.SolverMaxIterations" { return @($lead) + @(
            "Для LoadMultiplier эта настройка задает, сколько итераций можно сделать на каждой проверяемой точке lambda, прежде чем считать эту точку численно несошедшейся.",
            "Для UltimateStrain эта же настройка ограничивает число итераций прямого поиска предельной плоскости деформаций.",
            "Если значение слишком мало, сложное сочетание может получить NumFail даже тогда, когда физически сечение несет нагрузку. Если значение слишком велико, расчет может дольше пытаться сойтись на плохих исходных данных.",
            "Практически: увеличивать имеет смысл, когда диаграммы корректны, геометрия нормальная, но расчет останавливается по превышению числа итераций. Если ошибка вызвана неправильными единицами, пустыми диаграммами или нереалистичной нагрузкой, увеличение итераций проблему не исправит.",
            "Capacity.SolverMaxIterations - это защитный численный параметр программы. Он не является расчетным коэффициентом СП."
        ) }
        "Stability.Code" { return @($lead) + @(
            "SP63 включает учет продольного изгиба по формулам СП 63.13330.2018, п. 8.1.15: условная жесткость D, критическая сила Ncr и коэффициент eta.",
            "SP35 включает ветвь СП 35.13330.2011: для малых эксцентриситетов используется отдельная проверка устойчивости, для больших - увеличение момента через eta.",
            "В обоих вариантах эксцентриситет продольной силы считается относительно центра тяжести приведенного сечения бетон + арматура. Пользователь при этом вводит нагрузки относительно бетонного центра и Load.ReferenceOffset; программа переносит моменты к приведенному центру перед расчетом устойчивости.",
            "Расчет запускается только для тех сочетаний, где выбранный профиль имеет Calculation.Stability.Enabled = Yes."
        ) }
        "Stability.ElementLength" { return @($lead) + @(
            "Это расчетная геометрическая длина элемента до умножения на mu1/mu2.",
            "Для первой главной плоскости l0,1 = mu1*l, для второй l0,2 = mu2*l.",
            "Программа не анализирует закрепления элемента автоматически: пользователь задает длину и коэффициенты расчетной длины уже по своей расчетной схеме."
        ) }
        "Stability.Mu*" { return @($lead) + @(
            "Коэффициент задает расчетную длину в соответствующей главной плоскости приведенного сечения.",
            "mu1 относится к первой главной центральной плоскости приведенного сечения, mu2 - ко второй.",
            "Если оси сечения повернуты относительно X/Y, программа сама переводит моменты в главные оси, но выбор mu остается ответственностью пользователя."
        ) }
        "Stability.SystemType" { return @($lead) + @(
            "Determinate: случайный эксцентриситет добавляется к статическому моменту.",
            "Indeterminate: момент от расчета принимается, но его эксцентриситет должен быть не меньше случайного.",
            "В обозначениях СП 63 это формирование e0 с учетом ea; в обозначениях СП 35 - формирование ec с учетом ec,sl.",
            "Этот выбор влияет только на учет случайного эксцентриситета, а не на построение НДМ-диаграмм."
        ) }
        "Stability.ZeroMomentEccentricitySign*" { return @($lead) + @(
            "Если момент в главной плоскости равен нулю, знак случайного эксцентриситета нельзя определить из направления момента.",
            "Пользователь задает 1 или -1, чтобы явно указать, в какую сторону прикладывать случайный эксцентриситет: ea по СП 63 или ec,sl по СП 35.",
            "Настройка относится к главным плоскостям, а не к глобальным осям X/Y."
        ) }
        "Stability.PhiLMode" { return @($lead) + @(
            "Auto считает коэффициент длительности по отношению длительного момента к полному моменту.",
            "PhiL2 принимает phi_l = 2 как консервативный пользовательский режим.",
            "Для СП 63 таблица нагрузок означает постоянные + длительные нагрузки. Для СП 35 таблица означает постоянные нагрузки."
        ) }
        "Stability.SP63.*" { return @($lead) + @(
            "Эта настройка применяется только при Stability.Code = SP63.",
            "ks участвует в условной жесткости D = kb*Eb*Ib + ks*Es*Is.",
            "DeltaEMin и DeltaEMax ограничивают delta_e = e0/h в формуле kb. По СП 63 для расчета принимается диапазон 0.15...1.5.",
            "Материалы берутся строго из профиля: расчет устойчивости не подменяет ULS(I)/SLS(II) сам."
        ) }
        "Stability.SP35.*" { return @($lead) + @(
            "Эта настройка применяется только при Stability.Code = SP35.",
            "PhiP оставлен пользовательским коэффициентом. Для ненапрягаемой арматуры в текущей постановке обычно используется значение 1.",
            "NOverNcrLimit контролирует условие N/Ncr для ветви eta СП 35.",
            "Табличные коэффициенты СП 35 берутся из отдельной таблицы rngSP35Table721 на Config, а не из кода."
        ) }
        "SLS.Crack.Allowable" { return @($lead) + @(
            "Это предельно допустимая ширина раскрытия нормальной трещины a_crc,ult.",
            "Программа не назначает это значение автоматически. Пользователь выбирает его по нормам, категории трещиностойкости, условиям эксплуатации и проектным требованиям.",
            "После расчета программа выводит коэффициент запаса по трещинам: SFcrc = a_crc,ult / a_crc.",
            "Если SFcrc < 1, рассчитанная ширина раскрытия превышает заданный пользователем предел. Это не численная ошибка, а результат проверки.",
            "Число вводится в текущей INPUT-единице длины. Например допустимые 0.3 mm задаются как 0.3 при Units.Length.Input = mm, 0.03 при cm или 0.0003 при m.",
            "Условие проверки раскрытия трещин приведено в СП 63.13330.2018, п. 8.2.6, формула (8.118). Конкретное допустимое значение зависит от расчетной ситуации и должно быть подтверждено пользователем."
        ) }
        "SLS.Crack.Phi1" { return @($lead) + @(
            "Это коэффициент длительности действия нагрузки φ1 в формуле раскрытия нормальной трещины.",
            "В текущей реализованной схеме программа считает продолжительное раскрытие a_crc = a_crc1, поэтому используется φ1 = 1.4.",
            "Значение задается пользователем напрямую. Если нужна иная расчетная ситуация, пользователь меняет коэффициент осознанно; программа не подбирает его автоматически.",
            "Коэффициент входит в формулу раскрытия трещины по СП 63.13330.2018, п. 8.2.15, формула (8.128)."
        ) }
        "SLS.Crack.Phi2" { return @($lead) + @(
            "Это коэффициент φ2, учитывающий поверхность продольной арматуры в формуле раскрытия трещины.",
            "Теперь φ2 задается пользователем напрямую. Отдельного поля профиля арматуры нет, чтобы не держать два источника истины.",
            "Обычный ориентир по СП 63: 0.5 для арматуры периодического профиля и 0.8 для гладкой арматуры.",
            "Коэффициент входит в формулу раскрытия трещины по СП 63.13330.2018, п. 8.2.15, формула (8.128)."
        ) }
        "SLS.Crack.Phi3Mode" { return @($lead) + @(
            "Эта настройка задает, как принимать коэффициент φ3.",
            "Auto: программа выбирает φ3 по характеру продольной силы текущего сочетания. При растягивающей продольной силе принимается 1.2, иначе 1.0.",
            "User: программа берет численное значение из SLS.Crack.Phi3 и не пытается определять характер нагружения автоматически.",
            "Настройка влияет только на формулу a_crc и не меняет НДС, диаграммы материалов или состав расчетной зоны."
        ) }
        "SLS.Crack.Phi3" { return @($lead) + @(
            "Это пользовательское значение коэффициента φ3 для режима SLS.Crack.Phi3Mode = User.",
            "При SLS.Crack.Phi3Mode = Auto эта ячейка остается справочной: фактическое φ3 определяется программой по текущему сочетанию и выводится в Results.",
            "Коэффициент входит в формулу раскрытия трещины по СП 63.13330.2018, п. 8.2.15, формула (8.128)."
        ) }
        "SLS.Crack.PsiMode" { return @($lead) + @(
            "Эта настройка задает, как определить коэффициент psi_s в формуле раскрытия трещины.",
            "User: psi_s берется из SLS.Crack.PsiS. Это самый простой и быстрый режим: раскрытие считается по уже найденному НДС текущего сочетания, без отдельного поиска состояния образования трещины.",
            "Auto: программа сначала делает проверку с psi_s = 1. Если полученная ширина a_crc не превышает SLS.Crack.Allowable, расчет на этом заканчивается, а состояние образования трещины и sigma_s,crc не считаются.",
            "Если первая проверка в Auto не прошла, программа ищет долю заданной нагрузки, при которой в бетоне появляется нормальная трещина. В Results эта доля подписана как lambda_crc: например 0.70 означает, что трещина появляется примерно при 70% текущего сочетания.",
            "Для обычного изгиба программа масштабирует весь набор N, Mx и My в одной пропорции. Трещина считается образованной, когда растянутый бетон достигает предельной деформации eps_bt,ult по СП 63, п. 8.2.14 и 8.1.30.",
            "Для двузначной эпюры принимается eps_bt,ult = eps_bt2. Для однозначно растянутой эпюры с ненулевой кривизной применяется формула (8.54): eps_bt,ult = eps_bt2 - (eps_bt2 - eps_bt0) * eps1 / eps2. Для центрального растяжения используется отдельная сила образования трещин Ncrc.",
            "После нахождения уровня образования трещины программа отдельно считает состояние сразу после ее появления: растянутый бетон уже выключен, а нагрузка соответствует найденному уровню. Из этого состояния берется sigma_s,crc по тому же набору стержней, который выбран для текущего сочетания.",
            "sigma_s,crc считается как средневзвешенное по площади напряжение этих же стержней: сумма As_i * sigma_i,crc делится на сумму As_i. СП 63 задает саму величину sigma_s,crc в п. 8.2.18, но не описывает, как усреднять много стержней произвольного сечения N + Mx + My; это принятая НДМ-интерпретация.",
            "Формула: psi_s = 1 - 0.8 * sigma_s,crc / sigma_s. При численных выбросах значение ограничивается диапазоном 0...1.",
            "Если найденная доля больше 1, трещина при заданной нагрузке еще не образована, поэтому a_crc принимается равным 0.",
            "Важно: режим растянутого бетона из расчета прочности здесь не используется. Для поиска образования трещины берется модель MaterialModel.CrackInitiation из расчетного профиля. Нормативная настройка для Mcrc: SLS(II), ThreeLine, UseDiagram и кратковременные деформационные параметры.",
            "Базовая формула раскрытия приведена в СП 63.13330.2018, п. 8.2.15, формула (8.128). Для произвольного сечения N + Mx + My программа применяет НДМ-обобщение, потому что СП не дает отдельной пошаговой процедуры для такой дискретной модели."
        ) }
        "SLS.Crack.PsiS" { return @($lead) + @(
            "Это пользовательское значение коэффициента ψs для режима SLS.Crack.PsiMode = User.",
            "При SLS.Crack.PsiMode = Auto эта ячейка не задает итоговый коэффициент: программа сначала проверяет ψs = 1, а при непрохождении уточняет ψs по состоянию сразу после образования трещины.",
            "Настройка влияет только на формулу ширины раскрытия и не меняет НДС сечения.",
            "Автоматическое уточнение ψs выполняется по СП 63.13330.2018, п. 8.2.18, формула (8.137)."
        ) }
        "SLS.Crack.TensionZoneMode" { return @($lead) + @(
            "Эта настройка определяет, какая часть растянутого бетона входит в Abt для формулы расстояния между трещинами.",
            "FullTension: берется вся фактически растянутая зона бетона. Бетонный конечный элемент попадает в Abt, если его центр находится в растяжении по найденной плоскости деформаций.",
            "Effective: берется эффективная зона у наиболее растянутой поверхности. Ее глубина принимается h_bt = min(max(x_t, 2a), 0.5h), где x_t - фактическая растянутая зона. Если x_t < 2a, для Abt принимается полоса 2a, даже если она заходит за нейтральную линию; бетонные КЭ выбираются по попаданию центра в эту полосу. Стержни для As, sigma_s и d_s,eq при этом все равно берутся только растянутые.",
            "Геометрия произвольного сечения считается по нормали к нейтральной линии. h - габарит бетона по этой нормали; a - расстояние от внешней растянутой поверхности до оси наиболее крайнего растянутого стержня.",
            "Внешняя поверхность восстанавливается приближенно: каждый бетонный КЭ считается эквивалентным кругом площади A, r = sqrt(A/pi). Это используется только для s_t, s_c, h и a; усилия НДМ от этого не меняются.",
            "Abt считается дискретно: если центр бетонного КЭ попал в расчетную зону, учитывается вся его площадь; частичные площади не считаются. As включает только растянутые стержни, оси которых находятся внутри той же зоны.",
            "Расстояние между трещинами: ls0 = 0.5 * (Abt / As) * ds. Затем ls = max(ls0, 10ds, 100 mm), после этого ls = min(ls, 40ds, 400 mm).",
            "Если в зоне есть разные диаметры, программа использует ds,eq = sum(n_i*d_i^2) / sum(n_i*d_i). Это инженерное дополнение из EN 1992-1-1 для случая, который СП 63 явно не раскрывает.",
            "СП 63.13330.2018, п. 8.2.17 задает формулу (8.136) и ограничения 10ds...40ds, 100...400 mm. Для произвольного сечения, косого изгиба и дискретной сетки программа применяет описанное выше НДМ-обобщение."
        ) }
        "SLS.Crack.*" { return @($lead) + @(
            "Расчет трещин работает только для сочетаний II группы и использует тот же поиск равновесия, что и остальные состояния. Отдельного специального решателя для трещин нет.",
            "При PsiMode = User программа берет уже найденное НДС текущего сочетания и пользовательское SLS.Crack.PsiS. При PsiMode = Auto дополнительные расчеты состояния используются только если первая проверка с psi_s=1 не прошла.",
            "Напряжение sigma_s определяется напрямую из НДМ по выбранным растянутым стержням: sigma_s = sum(As_i * sigma_i) / sum(As_i), где sigma_i берется из диаграммы арматуры.",
            "Локальные зоны вокруг отдельных стержней и локальный максимум a_crc не реализуются. Используется один глобальный расчет для выбранной растянутой зоны.",
            "Для истинного центрального растяжения применяется отдельная ветка СП 63: если N > Ncrc по формуле (8.117), трещина считается образованной; Ncrc = Ared * Rbt,ser по формуле (8.127). После образования трещины берется вся площадь бетона и вся продольная растянутая арматура.",
            "Центральное растяжение распознается только по внешним нагрузкам относительно центра тяжести бетонного сечения: N должно быть растягивающим, а Mx_cg и My_cg должны быть равны нулю в пределах расчетного допуска.",
            "Mx_cg и My_cg учитывают Load.ReferenceOffset: если задано чистое N в таблице сочетаний, но точка приложения нагрузки смещена от бетонного центра, появляется внецентренное растяжение, и центральная ветка СП не включается.",
            "В сводке Results Ncrc выводится числом только для центрального растяжения, где оно действительно было рассчитано. В изгибной ветке или когда Ncrc не требовалось, выводится NotCalcul.",
            "Принятые ограничения: расчет наклонных трещин, напрягаемой арматуры и непродолжительного раскрытия не реализован."
        ) }
        "AutoCAD.Import.ConcreteLayer" { return @($lead) + @(
            "Импортер просматривает активный чертеж AutoCAD и берет бетонные Region только с этого слоя.",
            "Полилинии, блоки, тексты и объекты на других слоях не участвуют в расчете.",
            "Единицы чертежа считаются миллиметрами. Отдельного коэффициента масштаба импорта сейчас нет, поэтому геометрию в AutoCAD нужно готовить в мм.",
            "Каждый найденный бетонный Region превращается в расчетный бетонный элемент."
        ) }
        "AutoCAD.Import.RebarLayer" { return @($lead) + @(
            "Импортер берет арматурные Region только с этого слоя.",
            "Каждый найденный Region превращается в стержневой расчетный элемент. Геометрия импортируется без выбора диаграммы: материал будет выбран позже по расчетному профилю, а сам элемент получает техническую метку Rebar.",
            "Разные материалы или профили арматуры в одном импортированном сечении сейчас не поддерживаются."
        ) }
        "AutoCAD.Import.MinArea" { return @($lead) + @(
            "Это нижний порог площади Region для импорта.",
            "Region с площадью меньше этого значения игнорируется, чтобы случайный мусор чертежа или мелкие артефакты не попадали в расчетную модель.",
            "Если поставить значение слишком большим, можно случайно отбросить реальные маленькие арматурные области.",
            "Число вводится в текущей INPUT-единице площади. Порог 0.000001 mm2 соответствует 0.000001 при mm2, 0.00000001 при cm2 или 0.000000000001 при m2.",
            "AutoCAD-координаты при импорте считаются миллиметрами, а значение этой настройки вводится в текущей единице площади из блока Units."
        ) }
        "AutoCAD.Export.CombinationID" { return @($lead) + @(
            "Worst означает: взять определяющее сочетание из верхней сводки Results. Программа выбирает сочетание с минимальным прочностным запасом. Расчет трещин в этот выбор не входит.",
            "Если указан конкретный номер, например LC2, в AutoCAD выгружаются напряжения/деформации именно для этого сочетания.",
            "Значение должно совпадать с CombinationID, который есть в результатах последнего расчета. Если указать несуществующее сочетание, экспорт остановится с понятной ошибкой.",
            "Экспорт не запускает расчет заново. Если после расчета пользователь поменял нагрузки, сначала нужно выполнить новый расчет, и только потом экспортировать."
        ) }
        "AutoCAD.Export.LabelMode" { return @($lead) + @(
            "ValuesOnly выводит только численное значение выбранной профильной величины Visualization.Quantity.",
            "NamesAndValues добавляет к значению имя расчетного элемента, например C1 или R1.",
            "При NamesAndValues имена выводятся и для бетона, и для арматуры, если они сохранены в Results."
        ) }
        "AutoCAD.Export.*Enabled" { return @($lead) + @(
            "Yes включает вывод соответствующего объекта в AutoCAD. No пропускает его.",
            "Настройка управляет только оформлением экспорта и не меняет расчетные результаты.",
            "Экспорт читает последний Results и не запускает расчет заново."
        ) }
        "AutoCAD.Layer.*" { return @($lead) + @(
            "Слои создаются автоматически, если их нет в чертеже. Области бетона и арматуры лежат на своих базовых слоях, а текстовые подписи разделяются по физическому состоянию.",
            "Настройка влияет только на экспорт и не меняет расчет."
        ) }
        "AutoCAD.Color.*" { return @($lead) + @(
            "Нейтральный цвет применяется для нулевых или почти нулевых напряжений. Нулевые состояния не получают отдельного расчетного смысла, это только оформление."
        ) }
        "Plot.Enabled" { return @($lead) + @(
            "Yes разрешает построение Excel-схемы на листе Расчет. No отключает саму схему.",
            "Схема строится только по последнему сохраненному Results и не обращается к текущей геометрии Config."
        ) }
        "Plot.AutoUpdateAfterCalculation" { return @($lead) + @(
            "Yes автоматически перестраивает схему сразу после успешного расчета.",
            "No оставляет схему как есть; ее можно обновить отдельной кнопкой.",
            "В обоих случаях обновление схемы только читает Results и не запускает расчет заново."
        ) }
        "Plot.LoadCase" { return @($lead) + @(
            "Worst показывает определяющее сочетание из верхней сводки Results.",
            "Конкретный CombinationID, например LC2, показывает сохраненные напряжения, деформации и плоскость деформаций именно этого сочетания.",
            "Переключение LoadCase не создает геометрию заново: меняется только отображаемое расчетное состояние из Results."
        ) }
        "Plot.ResultGradient" { return @($lead) + @(
            "Yes включает цветовой градиент расчетных элементов по величине Visualization.Quantity выбранного профиля.",
            "No оставляет элементы в нейтральном оформлении без цветовой шкалы результата.",
            "Неработающий растянутый бетон остается отдельным серым состоянием и не смешивается с обычным градиентом."
        ) }
        "Plot.ResultLabelsEnabled" { return @($lead) + @(
            "Yes добавляет на схему численные подписи величины Visualization.Quantity выбранного профиля.",
            "Подписываются только бетонные элементы с пространственным шагом Plot.ResultLabelSpacing.",
            "Неработающий растянутый бетон не подписывается, чтобы схема не показывала нулевые значения как полезные напряжения."
        ) }
        "Plot.ResultLabelSpacing" { return @($lead) + @(
            "Это минимальное расстояние между численными подписями результата на схеме в модельных единицах длины.",
            "Чем больше значение, тем меньше подписей и чище схема. Чем меньше значение, тем подробнее поле значений, но выше риск визуальной перегрузки.",
            "Шаг применяется пространственно, а не по номеру ElementID."
        ) }
        "Plot.ResultPrecision" { return @($lead) + @(
            "Это количество знаков после запятой в подписях результата и легенде.",
            "Настройка влияет только на формат текста. Расчетные значения в Results не округляются заново."
        ) }
        "Plot.LegendMode" { return @($lead) + @(
            "Separate строит две шкалы: отдельно для бетона и отдельно для арматуры.",
            "Common строит одну общую шкалу для всех элементов. В этом режиме цвета берутся из Plot.Color.RebarCompression и Plot.Color.RebarTension.",
            "Раздельные шкалы удобны, когда диапазоны напряжений бетона и арматуры сильно отличаются."
        ) }
        "Plot.Color.*" { return @($lead) + @(
            "Значение задается в формате R,G,B, например 30,80,220.",
            "Цвет используется для крайнего значения соответствующего материала и физического состояния в легенде и на элементах.",
            "При Plot.LegendMode = Common используются цвета арматуры как общая шкала для всех элементов."
        ) }
        "Plot.*Enabled" { return @($lead) + @(
            "Yes показывает соответствующий объект на Excel-схеме. No скрывает его.",
            "Настройка влияет только на отображение уже сохраненных Results и не меняет расчет."
        ) }
        default { return @(
            $lead,
            "Единицу смотри в столбце Ед. на листе Config. Если там стоит формула, она подтягивает текущую INPUT-единицу из блока Units.",
            "Если эта настройка станет влиять на нормативно значимый расчетный критерий, инструкцию нужно расширить: добавить формулу, влияние на расчет и подтвержденную ссылку на СП."
        ) }
    }
}

# Делает первую строку инструкции одинаково понятной: пользователь сразу видит,
# какой именно ключ открыт и что конкретно он меняет. Остальные строки только
# раскрывают последствия и формулы для этой настройки.
function Get-InstructionLead {
    param([string]$Key, [string]$Comment)

    if ([string]::IsNullOrWhiteSpace($Comment)) {
        return "$Key - настройка Config. Отдельное подробное описание для нее пока не задано."
    }
    return "$Key - $Comment"
}

# Формирует список блоков справки для Config и специальных таблиц настроек.
function Get-SettingsInstructionCatalog {
    $items = New-Object System.Collections.Generic.List[object]
    $items.Add(@{ Key = "Units"; Title = "Единицы измерения"; Lines = @(
        "Этот блок задает пользовательские единицы ввода INPUT и вывода OUTPUT. Колонка INTERNAL является справочной и показывает фиксированные единицы расчетного ядра.",
        "Внутренние единицы ядра: длина - мм, площадь - мм2, сила - Н, момент - Н*мм, напряжение - МПа, кривизна - 1/мм.",
        "Формула: q_internal = q_user · k_input.",
        "Формула: q_output = q_internal / k_output.",
        "Все пересчеты единиц выполняются централизованно, поэтому расчетные формулы внутри программы работают в одной системе единиц.",
        "INPUT и OUTPUT могут быть разными. Например нагрузки удобно вводить в tf и tf*m, а результаты смотреть в kN и kN*m.",
        "Results является неизменяемым снимком последнего успешного расчета. Если после расчета пользователь поменял OUTPUT units, старые Results и схема не конвертируются заново; новые единицы применятся только после нового расчета.",
        "Выбор единиц не является расчетным допущением СП. При корректном обратимом пересчете физический результат должен сохраняться."
    )}) | Out-Null
    $items.Add(@{ Key = "SignConvention"; Title = "Система знаков"; Lines = @(
        "Этот блок задает пользовательскую систему знаков для ввода и вывода N, Mx и My.",
        "Внутри расчетного ядра используется фиксированная INTERNAL convention: растяжение положительное, сжатие отрицательное.",
        "Формула: ε(x, y) = ε₀ + κx·y + κy·x.",
        "Пользователь может выбрать, считать ли +N сжатием или растяжением. Программа преобразует внешний знак во внутренний перед расчетом и обратно перед записью Results.",
        "Сохраненный Output.SignConvention на Results является описанием сохраненных Results. Он не применяется повторно к Stress, N, Mx и My, потому что эти значения уже записаны с конечным пользовательским знаком.",
        "Цвета схемы определяются не по пользовательскому знаку Stress, а по сохраненному PhysicalState: Compression, Tension, InactiveTensionConcrete или NearZero.",
        "СП 63.13330.2018, п. 8.1.22 задает правило знаков для деформаций и напряжений НДМ: сжатие со знаком минус, растяжение со знаком плюс. Пользовательская настройка знаков меняет только ввод и вывод, а внутренняя математика остается той же."
    )}) | Out-Null
    $items.Add(@{ Key = "LoadCombinations"; Title = "Сочетания нагрузок"; Lines = @(
        "Таблица задает список сочетаний, которые будут рассчитаны при нажатии кнопки Выполнить расчет.",
        "CombinationID - короткое имя сочетания. Оно используется в Results, в заголовке схемы, в выборе Plot.LoadCase и AutoCAD.Export.CombinationID.",
        "N, Mx и My вводятся в текущих INPUT-единицах из блока единиц. Пустой Mx или My считается нулем, поэтому одноосный изгиб можно задавать как N + Mx или N + My без заполнения второго момента.",
        "ProfileId задает расчетный профиль из таблицы расчетных профилей. В строке сочетания нужно указать заголовок того профильного столбца, по которому это сочетание должно рассчитываться.",
        "Профиль определяет только, какие расчеты запрошены и какие материальные модели использовать. Порядок расчета и зависимости между состояниями задает программа, а не строка профиля.",
        "CapacityLoadPath задает, какие компоненты нагрузки масштабируются при поиске несущей способности: λ*Mx, λ*My, λ*Mxy, λ*N или λ*NMxy. Эта настройка находится в строке сочетания, потому что разные сочетания могут требовать разной траектории поиска.",
        "Если выбран λ*N, продольная сила N умножается на λ, а момент от ее смещенной линии действия масштабируется вместе с N. Если выбран λ*Mxy, N остается постоянной, а масштабируется только пользовательский вектор моментов.",
        "Если CapacityLoadPath оставлен пустым, программа сама выбирает понятный вариант: при наличии пользовательского Mx/My используется λ*Mxy, при чистой N - λ*N. Для прозрачности расчетов лучше задавать путь явно.",
        "Comment - пользовательское пояснение к сочетанию. Оно выводится в Summary и добавляется в заголовок схемы в скобках."
    ); Tables = @(
        @{
            Title = "Варианты CapacityLoadPath"
            Headers = @("Вариант", "Что масштабируется", "Что остается постоянным", "Минимально нужно задать", "Недопустимо")
            Rows = @(
                @("λ*Mx", "Пользовательский Mx.", "N, пользовательский My и моменты от эксцентриситета N.", "Mx <> 0.", "Пустой или нулевой Mx."),
                @("λ*My", "Пользовательский My.", "N, пользовательский Mx и моменты от эксцентриситета N.", "My <> 0.", "Пустой или нулевой My."),
                @("λ*Mxy", "Пользовательские Mx и My как один вектор изгиба.", "N и моменты от эксцентриситета N.", "Mx <> 0 или My <> 0.", "Одновременно Mx = 0 и My = 0."),
                @("λ*N", "N и моменты от эксцентриситета этой силы.", "Пользовательские Mx и My.", "N <> 0.", "Пустой или нулевой N."),
                @("λ*NMxy", "Полный внутренний вектор N, Mx, My.", "Ничего: все ненулевые компоненты входят в Base.", "N <> 0 или Mx <> 0 или My <> 0.", "Все масштабируемые компоненты равны нулю.")
            )
        },
        @{
            Title = "Как программа выбирает численный путь"
            Headers = @("Условие", "Что делает программа", "Что видно в Results", "Зачем это сделано", "Комментарий")
            Rows = @(
                @("Capacity.SolutionStrategy = Auto и изгибная траектория λ*Mx/λ*My/λ*Mxy.", "Сначала пробует UltimateStrain; при численной неудаче повторяет тот же CapacityLoadPath через LoadMultiplier.", "CapacitySolutionMethod показывает фактический метод.", "Обычно это самый быстрый путь для предельного момента при заданной N.", "Пользовательский CapacityLoadPath не меняется."),
                @("λ*N или λ*NMxy без пользовательских моментов.", "Использует LoadMultiplier как более устойчивую силовую ветку.", "CapacitySolutionMethod = LoadMultiplier.", "Чистая силовая траектория может быть вырожденной для UltimateStrain на плато диаграммы.", "CapacityLoadPath остается тем, который выбрал пользователь."),
                @("Capacity.SolutionStrategy = LoadMultiplier.", "Сразу выполняет одномерный поиск λ через Capacity.SearchMethod.", "CapacitySolutionMethod = LoadMultiplier.", "Нужен для сложных траекторий и контрольных расчетов.", "Bisection обычно самый предсказуемый, Brent часто быстрее, Secant чувствительнее."),
                @("CapacityLoadPath пустой.", "Автоматически выбирает λ*Mxy при наличии Mx/My, иначе λ*N при чистой N.", "В Results выводится фактически принятый CapacityLoadPath.", "Это защита для неполных строк сочетаний.", "Для прозрачности лучше заполнять путь явно."),
                @("Выбрана траектория, но ее масштабируемые компоненты нулевые.", "Возвращает InputErr.", "CapacityStatus = InputErr.", "Нельзя искать множитель для нулевого вектора Base.", "Например λ*Mx при Mx = 0.")
            )
        }
    )}) | Out-Null
    $items.Add(@{ Key = "StabilityLoads"; Title = "Нагрузки для устойчивости"; Lines = @(
        "Эта таблица задает длительную часть нагрузок, которая нужна только для расчета коэффициентов продольного изгиба.",
        "LC должен совпадать с CombinationID из основной таблицы сочетаний. N, Mx и My вводятся в тех же INPUT-единицах и той же системе знаков, что и основные сочетания.",
        "Для Stability.Code = SP63 пользователь указывает постоянные + длительные нагрузки.",
        "Для Stability.Code = SP35 пользователь указывает постоянные нагрузки.",
        "Если строка не заполнена или все усилия равны нулю, это не ошибка само по себе. Расчет продолжится, если формулы выбранного СП не требуют ненулевой длительной части.",
        "Load.ReferenceOffsetX/Y применяется к этой таблице так же, как к основной таблице сочетаний: пользовательские моменты сначала относятся к бетонному центру и заданному offset, затем для устойчивости переносятся к центру приведенного сечения.",
        "Таблица не запускает расчет сама. Устойчивость включается флагом Calculation.Stability.Enabled в расчетном профиле."
    )}) | Out-Null
    $items.Add(@{ Key = "SP35Table721"; Title = "Таблица 7.21 СП 35"; Lines = @(
        "Таблица содержит коэффициенты продольного изгиба из СП 35.13330.2011, таблица 7.21.",
        "В текущей программе используется ненапрягаемая арматура, поэтому в таблицу вынесены верхние значения.",
        "Расчетный класс устойчивости получает эту таблицу как массив из Config. В VBA нет скрытой резервной копии значений.",
        "Если выбран Stability.Code = SP35, таблица должна быть заполнена корректно. Пустая или поврежденная таблица приводит к InputErr для проверки устойчивости.",
        "Для промежуточных значений гибкости и эксцентриситета программа использует интерполяцию по загруженной таблице."
    )}) | Out-Null
    $items.Add(@{ Key = "ConcreteMaterialParameters"; Title = "Параметры бетона"; Lines = @(
        "Пользователь задает исходные сопротивления и модули бетона, а расчетные точки TwoLine/ThreeLine строятся программой автоматически.",
        "Для Strength используется I группа: Rb и Rbt. Для Mcrc и CrackedNDS используется II группа: Rb,ser и Rbt,ser.",
        "Eb и Ebt оставлены отдельными полями. Нормативно Ebt = Eb, потому что СП 63.13330.2018, п. 6.1.15 принимает начальный модуль бетона одинаковым при сжатии и растяжении.",
        "0.4·Eb в СП относится к модулю сдвига G, а не к Ebt.",
        "Деформационные параметры и контрольные точки бетонных диаграмм строятся по СП 63.13330.2018, пп. 6.1.14, 6.1.20-6.1.24 и 6.1.26.",
        "Серые контрольные таблицы точек на Config предназначены только для проверки человеком и формулами Excel. Расчетное ядро эти таблицы не читает."
    )}) | Out-Null
    $items.Add(@{ Key = "SteelMaterialParameters"; Title = "Параметры арматуры"; Lines = @(
        "Этот блок задает исходные сопротивления и модули обычной ненапрягаемой арматуры, из которых строятся диаграммы TwoLine/ThreeLine.",
        "Для Strength используется I группа: Rsc и Rs. Для Mcrc и CrackedNDS используется II группа: Rs,ser и программный параметр Rsc,ser.",
        "СП 63 отдельно не вводит Rsc,ser в том виде, как оно выведено в таблице. В программе поле оставлено для симметричного задания сжатой ветви; принято Rsc,ser = Rs,ser.",
        "Esc также является программным расширением для симметрии. Нормативно Esc = Es, потому что СП 63.13330.2018, п. 6.2.12 принимает модуль арматуры одинаковым при растяжении и сжатии.",
        "Производные точки арматурных диаграмм строятся по СП 63.13330.2018, пп. 6.2.11, 6.2.14 и 6.2.15. П. 6.2.13 задает смысл: TwoLine для физического предела текучести, ThreeLine для условного.",
        "Коэффициент phi2 для раскрытия трещин больше не берется из материала арматуры: он задается отдельной настройкой SLS.Crack.Phi2."
    )}) | Out-Null
    $items.Add(@{ Key = "CalculationProfiles"; Title = "Настройка расчетных профилей"; Lines = @(
        "Профиль используется только тогда, когда его ProfileId назначен строке сочетания. Заполненный PR3 или PR4 сам по себе не запускает расчет.",
        "Calculation.Strength.DirectState включает прямое НДС по модели прочности. Для этой модели по СП используются характеристики ULS(I): бетон Rb/Rbt и арматура Rsc/Rs.",
        "Calculation.Strength.Capacity включает поиск предельной несущей способности по выбранной λ-траектории сочетания. Используется та же модель материала, что и для НДС по прочности.",
        "Calculation.Crack.Width включает расчет нормальных трещин и существующую проверку продольных трещин. Для этих проверок по СП используются характеристики SLS(II): Rb,ser/Rbt,ser и Rs,ser.",
        "Calculation.Stability.Enabled включает учет продольного изгиба и проверку устойчивости перед последующими расчетами профиля.",
        "MaterialModel.Stability.ValueSet задает только набор характеристик материала для устойчивости. Диаграммы TwoLine/ThreeLine здесь не выбираются, потому что в формулах D, Ncr и eta используются R, Eb и Es.",
        "Для нормативного расчета устойчивости следует выбирать ULS(I): это расчет прочности сжатого элемента по первой группе предельных состояний.",
        "MaterialModel.Strength задает модель материала для НДС по прочности и для несущей способности. По СП 63, п. 6.1.23 бетон для прочности описывается TwoLine или ThreeLine; по п. 8.1.20 растянутый бетон при прочности допускается не учитывать.",
        "MaterialModel.CrackInitiation задает модель материала для состояния, когда нормальная трещина только появляется. По СП 63, п. 6.1.24 применяется II группа с работающей растянутой ветвью бетона; основной нормативный вариант для бетона - ThreeLine.",
        "MaterialModel.CrackedState задает модель материала для сечения после появления трещин. По СП 63, п. 6.1.26 допускается TwoLine или ThreeLine; растянутый бетон для уже образовавшейся трещины в НДС не учитывается.",
        "Calculation.Crack.Width включает расчет нормальных трещин и существующую проверку продольных трещин. Отдельного профиля, флага или ветки LongitudinalCrack нет.",
        "Модель CrackedState обязательна при включенном Calculation.Crack.Width. Модель CrackInitiation нужна только при SLS.Crack.PsiMode = Auto; если пользователь задает psi_s вручную, эта модель может быть не заполнена и не блокирует расчет.",
        "Visualization.State и Visualization.Quantity управляют только схемой и AutoCAD после расчета. Если выбранного состояния нет в Results, сам расчет не отменяется; сообщение появится только при попытке построить схему или экспорт.",
        "Visualization.State = StrengthState выводит НДС фактического сочетания по модели прочности.",
        "Visualization.State = CapacityState выводит предельное состояние, найденное расчетом несущей способности. Оно есть только если расчет несущей способности был включен и успешно нашел предельную точку.",
        "Visualization.State = CrackedState выводит НДС фактического сочетания как сечения с уже образовавшимися трещинами, без работы растянутого бетона.",
        "Visualization.State = BeforeMcrcState выводит состояние непосредственно перед образованием нормальной трещины, когда растянутый бетон еще работает. Оно создается только в режиме SLS.Crack.PsiMode = Auto, если первая проверка с psi_s = 1 не прошла и потребовался расчет момента образования трещины.",
        "Visualization.State = AfterMcrcState выводит состояние сразу после образования нормальной трещины: растянутый бетон уже выключен, а уровень нагрузки соответствует моменту появления трещины. Именно из этого состояния берется sigma_s,crc для уточнения psi_s.",
        "При центральном растяжении BeforeMcrcState не создается, потому что в этом случае нет нейтральной линии и программа определяет порог трещинообразования по расчетной силе Ncrc. Для проверки sigma_s,crc достаточно состояния AfterMcrcState сразу после появления трещины.",
        "Для арматуры во всех материальных моделях выбор TwoLine/ThreeLine читается по СП 63, п. 6.2.13: TwoLine соответствует арматуре с физическим пределом текучести, ThreeLine - арматуре с условным пределом текучести.",
        "Программа не подменяет параметры профиля автоматически. Если пользователь выбрал другой допустимый вариант диаграммы или учета растянутого бетона, расчет выполняется по явно заданной настройке профиля."
    )}) | Out-Null
    $items.Add(@{ Key = "MaterialDiagramControlTables"; Title = "Контрольные точки диаграмм"; Lines = @(
        "Служебные таблицы точек на Config нужны как независимый контроль того, какие координаты TwoLine/ThreeLine получаются из исходных параметров.",
        "Формулы Excel в этом блоке должны совпадать с теми формулами, по которым программа строит рабочие диаграммы материалов.",
        "Расчетное ядро не читает эти таблицы. Если пользователь случайно изменит формулы контрольного блока, расчетные результаты не изменятся, но проверка книги должна показать расхождение.",
        "Графики рядом с контрольными точками служат только для визуальной проверки формы диаграмм. Каждый график строится по строкам контрольной таблицы для I ГПС: одна линия TwoLine и одна линия ThreeLine."
    )}) | Out-Null
    $items.Add(@{ Key = "PlotAnnotationSettings"; Title = "Аннотации схемы"; Lines = @(
        "Таблица управляет только оформлением Shape-аннотаций схемы: размерными линиями, выносными линиями, текстом размеров и групповыми подписями арматуры вида 6Ø32 или 6Ø32 + 3Ø20.",
        "Расчетные объекты - бетонные элементы, стержни, нейтральная линия, главные оси и точка приложения нагрузки - не являются этими Shape-аннотациями и строятся отдельными графическими сериями.",
        "Offset всегда задается в реальных миллиметрах сечения. Для размерных линий это отступ от грани до размерной линии, для арматуры - отступ от линии осей стержней.",
        "TextUnits выбирает смысл TextHeight и TextGap отдельно для подписей арматуры и размерных линий. При pt текст задается в обычных пунктах Excel, поэтому визуальный размер почти не зависит от габаритов сечения при той же области построения. mm используется, когда текстовые размеры должны масштабироваться вместе с моделью.",
        "Если TextUnits=pt, числовые значения TextHeight и TextGap нужно задавать как размеры шрифта и зазоры в Excel points. При TextUnits=mm эти же строки трактуются как миллиметры модели.",
        "Для арматуры Placement задает сторону подписи относительно линии осей стержней: Inside - внутрь сечения, Outside - наружу.",
        "Для размерных линий Placement задает сторону текста относительно размерной линии. Сама размерная линия остается на расчетном месте.",
        "LineWeight, ArrowSize и ArrowType влияют только на внешний вид. Они не меняют расчет, Results и AutoCAD export."
    )}) | Out-Null
    $items.Add(@{ Key = "CircleGeometry"; Title = "Круглое сечение"; Lines = @(
        "Circle.Diameter задает внешний диаметр круглого бетонного сечения.",
        "Rebar.AxisDistance - расстояние от наружной грани круга до оси первого ряда стержней. Если Rebar.Count = 0 или Rebar.Diameter пустой/нулевой, первый ряд не создается.",
        "Rebar.Diameter2 и Rebar.Diameter3 задают дополнительные ряды. Пустой или нулевой диаметр означает, что соответствующий ряд пропускается без ошибки.",
        "Stacked ставит дополнительный стержень внутрь сечения по радиусу к центру. SideBySide ставит его по касательной справа по ходу часовой стрелки относительно стержня первого ряда.",
        "Semantic-аннотации круга строятся отдельным shape-builder: габаритный размер диаметра и групповые подписи арматуры затем выводятся тем же универсальным механизмом схемы, что и для остальных сечений."
    )}) | Out-Null
    $items.Add(@{ Key = "RoundedRectangleGeometry"; Title = "Скругленный прямоугольник"; Lines = @(
        "RoundedRectangle.Width и RoundedRectangle.Height задают основные габариты бетонного сечения.",
        "Радиусы RoundedRectangle.RadiusTopLeft/TopRight/BottomRight/BottomLeft задаются независимо для каждого угла. Нулевой радиус означает обычный прямой угол.",
        "Скругленный прямоугольник пока используется как геометрическая бетонная форма. Если для него не задан отдельный builder арматуры или аннотаций, расчет не падает: отсутствующие specialized-данные просто не выводятся.",
        "После генерации программа работает с сечением одинаково независимо от того, было оно построено в книге или импортировано из AutoCAD."
    )}) | Out-Null
    $items.Add(@{ Key = "LShapeGeometry"; Title = "Г-образное сечение"; Lines = @(
        "H1/B1 описывают верхний прямоугольник Г-сечения, H2/B2 - нижний прямоугольник. Полная высота расчетной формы равна H = H1 + H2.",
        "Основное армирование задается по смысловым граням H1, B1, H2, B2. Для каждой грани есть сторона _1 и _2: для вертикальных H это левая и правая грань, для горизонтальных B это верхняя и нижняя грань.",
        "Нумерация граней идет слева направо для вертикальных граней и сверху вниз для горизонтальных. Это правило используется в именах as_1/as_2, d_1/d_2, n_1/n_2, t1_1/t2_1 и t1_2/t2_2.",
        "Формула: step = (L_edge − t1 − t2) / (n − 1), если n > 1.",
        "Если n = 1, создается один стержень в допустимой позиции по принятой реализации builder.",
        "Если n_1 = 0 или d_1 = 0, ряд _1 не создается. Если первого ряда на стороне нет, дополнительные 2-й и 3-й ряды этой стороны также не создаются, даже если их диаметры указаны.",
        "Дополнительные ряды строятся от координат первого ряда. loc_2row/loc_3row задает направление Stacked или SideBySide, а bind_2row/bind_3row задает EachBar или EverySecondBar.",
        "Размеры и групповые подписи арматуры создаются по смысловым граням Г-сечения и затем сохраняются в Results. Схема не угадывает грани по координатам уже после расчета."
    )}) | Out-Null
    $items.Add(@{ Key = "BatchSummary"; Title = "Сводка пакетного расчета"; Lines = @(
        "Этот раздел описывает верхнюю таблицу Results: rngBatchSummary. В ней собраны основные итоги по всем сочетаниям, чтобы быстро увидеть, какое сочетание определяет прочность, какой запас дает выбранная λ-траектория и что получилось по раскрытию трещин.",
        "Первые строки - служебная часть. Определяющее сочетание по прочности берется по минимальному рассчитанному CapacitySafetyFactor. После перехода на CapacityLoadPath этот запас равен найденному lambda для выбранной траектории: λ*Mx, λ*My, λ*Mxy, λ*N или λ*NMxy. Проверка трещин в выбор worst для прочности не входит.",
        "Время расчета, с - полное время выполнения макроса от старта пользовательского расчета до завершения записи Results, обновления схемы и, если включен General.ExecutionReportEnabled, сохранения txt-отчета.",
        "Точка приложения нагрузки - смещение точки, относительно которой пользователь задает N, Mx и My, от центра тяжести бетонного сечения. При нулевых Load.ReferenceOffsetX/Y в summary выводится X=0 и Y=0.",
        "CombinationID (сочетание) - ID строки из таблицы сочетаний на листе Config.",
        "Comment (комментарий) - пользовательский комментарий к сочетанию. Этот же текст добавляется в заголовок схемы в скобках.",
        "ProfileId (профиль) - расчетный профиль, назначенный этому сочетанию. Профиль задает запрошенные расчеты, материальные модели и состояние для визуализации.",
        "OverallStatus (итог) - общий короткий статус сочетания. Он собирается из применимых статусов разделов с приоритетом InputErr -> NumFail -> FAIL -> OK. N/A не ухудшает итог, если раздел просто неприменим.",
        "DirectStateStatus (НДС) показывает результат прямого расчета состояния от заданного сочетания. OK - равновесие найдено внутри физической диаграммы материалов.",
        "DirectStateStatus = FAIL означает, что равновесие найдено, но состояние вышло за физические пределы диаграмм материалов. DirectStateStatus = NumFail означает, что численное равновесие не найдено.",
        "eps_c,min (бетон сжатие) - минимальная деформация бетонных элементов в найденном состоянии. В INTERNAL convention сжатие отрицательное, поэтому для сжатого бетона обычно контролируется отрицательный минимум.",
        "eps_c,max (бетон растяжение) - максимальная растягивающая деформация бетонных элементов.",
        "eps_s,min (арматура сжатие) - минимальная деформация стержней арматуры.",
        "eps_s,max (арматура растяжение) - максимальная растягивающая деформация стержней арматуры.",
        "eps_cu,comp (предел бетона сжатие) - предельная сжатая деформация бетонной диаграммы активного расчета.",
        "eps_cu,tens (предел бетона растяжение) - предельная растягивающая деформация бетонной диаграммы. Поле заполняется только если активная диаграмма действительно учитывает растянутый бетон.",
        "eps_su,comp (предел арматуры сжатие) и eps_su,tens (предел арматуры растяжение) - предельные деформации построенной диаграммы арматуры активного расчета.",
        "h в блоке деформаций - высота всего бетонного сечения по нормали к нейтральной линии именно для фактического состояния этого сочетания. Это не параметр предельного момента.",
        "x в блоке деформаций - высота сжатой зоны по той же нормали для фактического состояния сочетания. Если все сечение сжато, x близко к h; если все растянуто, x близко к 0.",
        "Пользовательский StrainSafetyFactor больше не выводится. Текущие деформации eps_* нужны для диагностики НДС, а не как самостоятельный коэффициент несущей способности.",
        "CapacityStatus (прочность) показывает результат поиска предельной точки по выбранной λ-траектории. OK - предельный множитель найден и он не меньше 1.",
        "CapacityStatus = FAIL означает, что предельная точка найдена, но исходное сочетание уже превышает несущую способность. CapacityStatus = NumFail означает, что численный поиск не дал надежного результата.",
        "CapacityLimitState (предельное состояние) - причина остановки поиска: предельная деформация бетона, предельная деформация арматуры, численная неудача и т.п. В коротком пользовательском Status численная несходимость показывается как NumFail.",
        "CapacityLoadPath (что масштабируется) - выбранная пользователем траектория: λ*Mx, λ*My, λ*Mxy, λ*N или λ*NMxy. Компоненты, которые не входят в путь, остаются постоянными.",
        "CapacitySolutionMethod (фактический метод) - фактический метод, которым программа реально нашла предельную точку: UltimateStrain или LoadMultiplier. Это поле может отличаться от предпочтительной Capacity.SolutionStrategy, если Auto перешел в повторный поиск или выбранная λ-траектория устойчивее считается через LoadMultiplier.",
        "lambdaUltimate (множитель) - найденный предельный множитель λ. Именно он выводится как CapacitySafetyFactor.",
        "Nult, Mxult, Myult - предельные компоненты только для тех нагрузок, которые входили в выбранную λ-траекторию. Если компонент не масштабировался, его ячейка остается пустой.",
        "Если выбрано λ*N, момент от смещения точки приложения продольной силы масштабируется вместе с N. Если выбрано λ*Mx/λ*My/λ*Mxy, N остается постоянной, а момент от ее эксцентриситета остается в постоянной части.",
        "h в блоке Capacity - высота бетонного сечения по нормали к нейтральной линии именно в найденном предельном состоянии λult. В DirectState или при неудачном capacity-поиске поле остается пустым.",
        "x в блоке Capacity - сжатая зона именно при найденном λult. Она не подменяется текущим состоянием сочетания, потому что нейтральная линия при увеличении нагрузки может повернуться.",
        "CapacitySafetyFactor (запас по λ) - формульная ссылка на lambdaUltimate. Этот коэффициент используется для выбора Worst по прочности.",
        "CrackStatus (статус трещин) - статус расчета ширины раскрытия нормальных трещин по II группе. Расчет выполняется только для профилей с Calculation.Crack.Width = Yes и только когда исходное НДС найдено внутри физической диаграммы материалов.",
        "CrackFormed (трещина есть) - Да, если трещина считается образованной; Нет, если при заданной нагрузке трещина еще не образовалась.",
        "h в блоке трещин - полный размер бетонного сечения по нормали к нейтральной линии, используемый для определения эффективной растянутой зоны.",
        "a (от края до стержня) - расстояние от внешнего края наиболее растянутого бетонного волокна до оси наиболее удаленного растянутого стержня по нормали к нейтральной линии.",
        "TensionZoneDepth (растянутая зона) - фактическая высота растянутой зоны бетона по нормали к нейтральной линии.",
        "EffectiveZoneDepth (эффективная зона) - принятая глубина зоны взаимодействия для Abt. В режиме Effective используется h_bt = min(max(x_t, 2a), 0.5h); зона может геометрически зайти за нейтральную линию ради условия 2a, и бетонные КЭ внутри этой полосы учитываются по центрам независимо от знака текущей деформации.",
        "As (растянутая арматура) - суммарная площадь растянутых стержней, оси которых попали в принятую расчетную зону.",
        "Abt (растянутый бетон) - суммарная площадь бетонных конечных элементов, центры которых попали в принятую растянутую/эффективную зону.",
        "ds_eq (экв. диаметр) - эквивалентный диаметр для группы стержней разных диаметров: ds,eq = sum(n_i*d_i²) / sum(n_i*d_i). Это инженерное дополнение по EN 1992-1-1 для случая, который СП 63 явно не раскрывает.",
        "phi1 (длительность) - коэффициент длительности действия нагрузки. В текущей реализации продолжительное раскрытие трещин по II группе использует phi1 = 1.4.",
        "phi2 (поверхность арматуры) - пользовательский коэффициент из SLS.Crack.Phi2. Ориентир: 0.5 для периодического профиля и 0.8 для гладкой арматуры.",
        "phi3 (характер нагрузки) - коэффициент характера работы элемента. При SLS.Crack.Phi3Mode = Auto программа принимает 1.2 для растягивающей продольной силы N и 1.0 для остальных случаев; при User берет SLS.Crack.Phi3.",
        "psi_s (работа растянутого бетона) - коэффициент учета работы бетона между трещинами. При пользовательском режиме берется SLS.Crack.PsiS. При Auto программа сначала проверяет с psi_s = 1; если проверка не прошла, уточняет коэффициент по состоянию сразу после образования трещины.",
        "sigma_s - средневзвешенное напряжение растянутой арматуры из выбранного набора стержней.",
        "sigma_s,crc - средневзвешенное по площади напряжение той же расчетной группы растянутых стержней сразу после образования трещины. Набор стержней берется из текущего сочетания и не переопределяется заново для состояния образования трещины.",
        "В сводке sigma_s,crc заполняется только если режим Auto действительно запустил уточнение psi_s после непрохождения первой проверки с psi_s = 1; иначе выводится NotCalcul.",
        "Es - модуль упругости арматуры из Steel.Es в выбранной OUTPUT-единице напряжений.",
        "ls_raw - исходный шаг трещин до ограничений: ls,raw = 0.5 · (Abt / As) · ds_eq.",
        "ls - принятый шаг трещин после ограничений. Сначала применяется нижняя граница: ls = max(ls_raw; 10·ds; 100 мм). Затем применяется верхняя граница: ls = min(ls; 40·ds; 400 мм).",
        "lambda_crc - доля заданной нагрузки, при которой появляется нормальная трещина. Например 0.70 означает, что трещина появляется примерно при 70% текущего сочетания. При центральном растяжении эта доля равна отношению силы образования трещин Ncrc к заданной продольной силе N. Если уточнение psi_s не выполнялось, выводится NotCalcul.",
        "Ncrc - продольная сила образования трещин при истинном центральном растяжении. Для нее используется отдельная ветка СП: Ncrc = Ared · Rbt,ser. Если центральная ветка не применялась или Ncrc не требовалось считать, в сводке выводится NotCalcul.",
        "TensionRebarCount - количество растянутых стержней, вошедших в расчет As и sigma_s.",
        "a_crc (ширина трещины) - расчетная ширина раскрытия трещины.",
        "Формула: a_crc = phi1 · phi2 · phi3 · psi_s · (sigma_s / Es) · ls.",
        "a_crc,ult (допустимая ширина) - пользовательское допустимое значение SLS.Crack.Allowable.",
        "CrackSafetyFactor (запас по трещинам) - отношение допустимой ширины к расчетной: SFcrc = a_crc,ult / a_crc. Если трещина не образована или a_crc = 0, запас не выводится.",
        "LongitudinalCrackStatus (продольные трещины) - статус SLS(II)-проверки сжатого бетона на продольные трещины. Проверка выполняется внутри Calculation.Crack.Width, когда состояние с трещинами найдено в физическом диапазоне диаграммы и в бетоне есть сжимающие напряжения.",
        "sigma_c,max (сжатие бетона) - максимальное по модулю сжимающее напряжение среди бетонных элементов в уже найденном НДС этого сочетания.",
        "Rb,mc2 (допустимое сжатие) - пользовательское допустимое значение для проверки продольных трещин по СП 35. Число вводится на Config уже с учетом применимых коэффициентов условий работы.",
        "LongitudinalCrackSafetyFactor (запас) - формула Results: SFlong = Rb,mc2 / sigma_c,max. Если Crack.Width не включен или сжатого бетона нет, проверка получает N/A и запас не выводится.",
        "Статус устойчивости показывает результат учета продольного изгиба перед НДМ. OK означает, что проверка выбранной методики прошла; FAIL - запас по устойчивости меньше 1 или нарушено ограничение выбранной ветви; N/A - продольная сила не сжимает сечение или расчет устойчивости не включен профилем.",
        "Контрольная сжимающая N - минимальная сила, по которой считается запас устойчивости. Для ветвей через eta это критическая сила Ncr. Для табличной ветви СП 35 это предельная сила Nult,stab, потому что Ncr в этой ветви не определяется.",
        "Усилия после устойчивости выводятся в двух привязках. Mx/My в точке нагрузки нужны для сопоставления с тем, как пользователь задал сочетание и Load.ReferenceOffsetX/Y. Mx/My в ц.т. приведенного сечения - это моменты, которые участвовали в расчете устойчивости; перед дальнейшим НДМ они переносятся к внутренним координатам элементов без повторного учета пользовательского эксцентриситета.",
        "Ncr в детальных блоках СП 35/СП 63 заполняется только там, где методика реально считает критическую силу. Если выбранная ветвь вместо Ncr считает Nult,stab, в колонке Ncr выводится N/A.",
        "Nult,stab в детальных блоках заполняется для табличной ветви СП 35. Для ветвей, где проверка идет через критическую силу Ncr и коэффициент eta, в колонке Nult,stab выводится N/A.",
        "Эксцентриситеты подписаны в обозначениях выбранного СП: для СП 63 выводятся e0 и ea, для СП 35 - ec и ec,sl.",
        "Запас устойчивости - отношение контрольной сжимающей силы к действующей сжимающей N. В зависимости от ветви это Ncr/N или Nult,stab/N.",
        "MinSafetyFactor (общий запас) - минимальный положительный коэффициент из CapacitySafetyFactor, CrackSafetyFactor, LongitudinalCrackSafetyFactor и StabilitySafetyFactor. Столбец информативный и специально выделен более темным серым цветом."
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
        $items.Add(@{ Key = $key; Title = "Аннотации схемы: $key"; Lines = (Get-SettingInstructionLines $key ([string]$row[4])) }) | Out-Null
    }
    foreach ($row in (Get-ConcreteMaterialParametersCatalog)) {
        $key = [string]$row[0]
        $items.Add(@{ Key = $key; Title = $key; Lines = (Get-SettingInstructionLines $key ([string]$row[4])) }) | Out-Null
    }
    foreach ($row in (Get-SteelMaterialParametersCatalog)) {
        $key = [string]$row[0]
        $items.Add(@{ Key = $key; Title = $key; Lines = (Get-SettingInstructionLines $key ([string]$row[4])) }) | Out-Null
    }
    foreach ($geometry in (Get-GeometrySettingsCatalog)) {
        if (-not ($geometry.ContainsKey("FaceTable") -and $geometry.FaceTable)) {
            foreach ($row in $geometry.Rows) {
                $key = [string]$row[0]
                if ($key.Length -gt 0 -and $key -notlike "дополнительные*") {
                    $items.Add(@{ Key = $key; Title = $key; Lines = (Get-SettingInstructionLines $key ([string]$row[3])) }) | Out-Null
                }
            }
        }
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

# Добавляет обычный абзац справки; длинный текст заранее режется на несколько
# строк, чтобы в готовой книге он не обрезался внутри объединенной ячейки.
function Add-GuideParagraph {
    param([object]$Sheet, [int]$Row, [string]$Text)

    $currentRow = $Row
    foreach ($line in (Split-GuideParagraphText $Text)) {
        $range = $Sheet.Range($Sheet.Cells.Item($currentRow, 1), $Sheet.Cells.Item($currentRow, 6))
        $range.Merge() | Out-Null
        $range.Value2 = $line
        $range.WrapText = $true
        $range.VerticalAlignment = -4160
        $currentRow++
    }
    return $currentRow
}

# Удаляет старые картинки формул, созданные прошлой сборкой справки.
# Cells.Clear не трогает Shapes, поэтому без отдельной очистки формулы
# могли бы накапливаться поверх новых строк после повторной пересборки книги.
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
    foreach ($run in $runs) {
        $font = if ($run.Kind -eq "Base") { $BaseFont } else { $ScriptFont }
        $size = $Graphics.MeasureString([string]$run.Text, $font)
        $width += $size.Width
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
    foreach ($run in $Runs) {
        if ($run.Kind -eq "Base") {
            $Graphics.DrawString([string]$run.Text, $BaseFont, $Brush, [single]$currentX, [single]$baseTop)
            $currentX += $Graphics.MeasureString([string]$run.Text, $BaseFont).Width
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
    $picture.Placement = 2
    $script:FormulaImageRows.Add(@{ Row = $Row; Height = $rowHeight }) | Out-Null
    return $Row + 1
}

# Добавляет компактную формулу в отдельном выделенном блоке.
function Add-GuideFormula {
    param([object]$Sheet, [int]$Row, [string]$Text)

    $image = New-TextFormulaImage $Text
    return (Add-GuideFormulaPicture $Sheet $Row $image $Text)
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

# Формирует верхний методический раздел листа "Справка" по фактическому
# алгоритму расчета нормальных трещин. Подробные настройки ниже остаются
# отдельными блоками, а этот раздел читается как общий user guide.
function Add-CrackWidthMethodologyGuide {
    param([object]$Sheet)

    $row = 1
    $row = Add-GuideTitle $Sheet $row "Расчет ширины раскрытия нормальных трещин в программе" 16
    $row = Add-GuideParagraph $Sheet $row "Этот раздел описывает реализованный расчет нормальных трещин для общего НДС N + Mx + My. Он нужен как пошаговая карта: какие данные берутся из НДМ, какие формулы СП применяются напрямую, а где используется инженерное обобщение для произвольного сечения."
    $row++

    $row = Add-GuideTitle $Sheet $row "Нормативные источники и статус реализации" 13
    $row = Add-GuideParagraph $Sheet $row "СП 63.13330.2018, разделы 8.1 и 8.2: НДМ нормального сечения, проверка раскрытия трещин, формулы (8.118), (8.119), (8.128), (8.136), (8.137). Конкретные ссылки в этом разделе привязаны именно к версии 2018 года."
    $row = Add-GuideParagraph $Sheet $row "СП 35.13330.2011 используется для отдельной SLS(II)-проверки продольных трещин по сжатому бетону. Нормальные трещины a_crc считаются по СП 63, а продольная проверка внутри Calculation.Crack.Width сравнивает σ_c,max с пользовательским R_b,mc2."
    $row = Add-GuideParagraph $Sheet $row "EN 1992-1-1 используется только как инженерное дополнение для эквивалентного диаметра при разных диаметрах стержней. Это не требование СП 63 и явно помечается как принятое обобщение."
    $row = Add-GuideParagraph $Sheet $row "В текущей версии расчет работает только для продолжительных сочетаний II группы: a_crc = a_crc1, φ1 = 1.4. Непродолжительное раскрытие a_crc1 + a_crc2 − a_crc3 не реализовано."
    $row = Add-GuideFormula $Sheet $row "a_crc = a_crc1      СП 63, п. 8.2.7, ф. (8.119)"
    $row++

    $row = Add-GuideTitle $Sheet $row "1. Исходное НДС для выбранного сочетания" 13
    $row = Add-GuideParagraph $Sheet $row "Для каждого сочетания программа сначала находит плоскость деформаций текущей нагрузки. Тип сечения при этом не важен: расчет использует бетонные элементы и стержни арматуры как единую модель сечения."
    $row = Add-GuideFormula $Sheet $row "ε(x, y) = ε₀ + κₓ·y + κᵧ·x"
    $row = Add-GuideParagraph $Sheet $row "В этой формуле ε(x,y) - относительная деформация точки с координатами x и y. ε₀ - деформация в начале принятой внутренней системы координат."
    $row = Add-GuideParagraph $Sheet $row "κₓ задает изменение деформации по координате y, а κᵧ - по координате x. Координаты, ε₀ и кривизны берутся из сохраненного расчетного состояния Results."
    $row = Add-GuideParagraph $Sheet $row "Из этого состояния берутся ε₀, κₓ, κᵧ, деформации и напряжения всех бетонных элементов и стержней. СП 63.13330.2018, п. 8.1.23 и 8.1.25 задают общий НДМ-подход; применение тех же данных для произвольного N + Mx + My в расчете трещин является принятой реализацией общей НДМ-постановки."
    $row = Add-GuideParagraph $Sheet $row "Режим SLS.Crack.PsiMode = User не ищет отдельно момент образования трещины: ψs берется из SLS.Crack.PsiS, а ширина считается по уже найденному НДС текущего сочетания."
    $row++

    $row = Add-GuideTitle $Sheet $row "2. Проверка факта образования трещины" 13
    $row = Add-GuideParagraph $Sheet $row "В режиме User программа считает уже образовавшуюся трещину и берет ψs из Config. В режиме Auto программа сначала считает ширину с ψs = 1 и сравнивает ее с a_crc,ult. Только если первая проверка не прошла, программа ищет уровень нагрузки, при котором появляется нормальная трещина."
    $row = Add-GuideFormula $Sheet $row "ε_bt,ult = ε_bt2      СП 63, п. 8.2.14 и 8.1.30, двузначная эпюра"
    $row = Add-GuideFormula $Sheet $row "ε_bt,ult = ε_bt2 - (ε_bt2 - ε_bt0) · ε1 / ε2      СП 63, п. 8.1.30, ф. (8.54)"
    $row = Add-GuideParagraph $Sheet $row "ε_bt,ult - предельная растягивающая деформация бетона, по которой определяется момент образования трещины. ε_bt2 берется как предельная деформация растяжения из выбранной диаграммы бетона."
    $row = Add-GuideParagraph $Sheet $row "ε_bt0 - деформация начала горизонтального участка трехлинейной растянутой ветви. В формуле (8.54) ε1 и ε2 - меньшая и большая по модулю деформации однозначно растянутой эпюры; принимается |ε2| >= |ε1|."
    $row = Add-GuideParagraph $Sheet $row "В общей изгибной ветке трещина считается появившейся, когда наиболее растянутый бетон достигает εbt,ult. Для двузначной эпюры используется εbt2; для однозначно растянутой эпюры с ненулевой кривизной используется формула (8.54). Для центрального растяжения применяется отдельная сила образования трещин Ncrc."
    $row = Add-GuideParagraph $Sheet $row "Для этого поиска растянутый бетон включается по MaterialModel.CrackInitiation расчетного профиля. Нормативная настройка для Mcrc: SLS(II), ThreeLine для бетона, UseDiagram для растянутого бетона и кратковременные деформационные характеристики."
    $row = Add-GuideParagraph $Sheet $row "В Results этот уровень выводится как λcrc: это доля заданного сочетания, при которой появляется трещина. Если λcrc > 1, то при текущей нагрузке трещина еще не образована и a_crc = 0. Если λcrc ≤ 1, расчет продолжается."
    $row++

    $row = Add-GuideTitle $Sheet $row "3. Центральное растяжение" 13
    $row = Add-GuideParagraph $Sheet $row "Если задано растягивающее N и моментный вектор относительно центра тяжести бетонного сечения равен нулю, включается отдельная ветка центрального растяжения. В этой ветке нейтральная линия отсутствует."
    $row = Add-GuideFormula $Sheet $row "N > N_crc      СП 63, п. 8.2.4, ф. (8.117)"
    $row = Add-GuideParagraph $Sheet $row "СП 63 для центрально-растянутых элементов задает отдельное условие образования трещин через продольную растягивающую силу: если внешнее растягивающее N больше N_crc, нужно выполнять расчет раскрытия трещин. Поэтому программа не пытается построить нейтральную линию там, где ее физически нет."
    $row = Add-GuideFormula $Sheet $row "N_crc = A_red · R_bt,ser      СП 63, п. 8.2.13, ф. (8.127)"
    $row = Add-GuideFormula $Sheet $row "λ_crc = N_crc / N"
    $row = Add-GuideParagraph $Sheet $row "N_crc - продольная сила образования трещин при центральном растяжении. Это не заданная нагрузка из таблицы сочетаний, а расчетная граница, с которой сравнивается заданное растягивающее N."
    $row = Add-GuideParagraph $Sheet $row "A_red - приведенная площадь сечения. R_bt,ser - сопротивление бетона растяжению для II группы предельных состояний."
    $row = Add-GuideParagraph $Sheet $row "Как программа распознает центральное растяжение: продольная сила должна быть растягивающей, а Mx_cg и My_cg относительно центра тяжести бетонного сечения должны быть равны нулю в пределах расчетного допуска."
    $row = Add-GuideParagraph $Sheet $row "При этом Mx_cg и My_cg учитывают Load.ReferenceOffset: Mx_cg = Mx_user + N·OffsetY, My_cg = My_user + N·OffsetX. Поэтому N при Mx_user = My_user = 0, но с ненулевым offset от бетонного центра, считается внецентренным растяжением и идет по общей ветке N + Mx + My."
    $row = Add-GuideParagraph $Sheet $row "В центральной ветке нет нейтральной линии, поэтому настройки SLS.Crack.TensionZoneMode = Effective/FullTension не должны менять результат. Программа принимает Abt равной всей площади бетона, а As - суммарной площади всех продольных стержней, которые находятся в растяжении в текущем НДС."
    $row = Add-GuideFormula $Sheet $row "A_red = A_b + A_s,total · E_s / E_b"
    $row = Add-GuideFractionFormula $Sheet $row "λ_crc =" "N_crc" "N" ""
    $row = Add-GuideParagraph $Sheet $row "Последовательность расчета такая: программа считает A_red, затем N_crc, затем λ_crc = N_crc / N. Если λ_crc > 1, заданная растягивающая сила меньше силы образования трещин, поэтому CrackFormed = Нет и a_crc = 0."
    $row = Add-GuideParagraph $Sheet $row "Если λ_crc ≤ 1, трещина считается образованной. Дальше программа использует всю площадь бетона Abt = A_b и всю растянутую продольную арматуру As, определяет σ_s по текущему НДС и считает ширину по общей формуле a_crc."
    $row = Add-GuideParagraph $Sheet $row "В формуле A_red величина A_b означает всю площадь бетона, а A_s,total - всю площадь продольной арматуры. Отношение E_s/E_b учитывает приведение арматуры к бетону."
    $row = Add-GuideParagraph $Sheet $row "Эта формула нужна только для N_crc и не подменяет расчет σ_s в раскрытой трещине."
    $row = Add-GuideParagraph $Sheet $row "После образования трещины для центральной ветки используются вся площадь бетона Abt и вся растянутая продольная арматура As. Если режим Auto требует уточнения ψs, программа не запускает общий изгибный поиск λcrc по εbt,ult, а использует λcrc = Ncrc / N и отдельно решает состояние после раскрытия трещины без растянутого бетона. Коэффициент φ3 принимается 1.2 как для растянутых элементов по СП 63.13330.2018, п. 8.2.15."
    $row++

    $row = Add-GuideTitle $Sheet $row "4. Нейтральная линия и направление расчета зоны" 13
    $row = Add-GuideParagraph $Sheet $row "В общем случае программа строит нормаль к нейтральной линии по найденным κₓ и κᵧ. Все расстояния для трещин считаются вдоль этой нормали, направленной в сторону растяжения."
    $row = Add-GuideFormula $Sheet $row "n = (κᵧ, κₓ) / √(κₓ² + κᵧ²)"
    $row = Add-GuideFormula $Sheet $row "s = nₓ·x + nᵧ·y"
    $row = Add-GuideParagraph $Sheet $row "n - единичный вектор нормали к нейтральной линии, направленный в сторону растяжения. nₓ и nᵧ - его проекции на оси X и Y."
    $row = Add-GuideParagraph $Sheet $row "s - координата точки вдоль этой нормали. Чем больше s, тем ближе точка к растянутому краю."
    $row = Add-GuideParagraph $Sheet $row "Даже если нейтральная линия находится за пределами сечения, используется тот же алгоритм: по координате s определяется растянутый край, сжатый край, h, a и зона Abt."
    $row++

    $row = Add-GuideTitle $Sheet $row "5. Внешняя поверхность бетонных элементов" 13
    $row = Add-GuideParagraph $Sheet $row "Для восстановления внешнего края по дискретной сетке каждый бетонный конечный элемент считается эквивалентным кругом только для геометрических расстояний s_t, s_c, h и a. Расчетная площадь элемента от этого не меняется."
    $row = Add-GuideFormula $Sheet $row "r_eq = √(A / π)"
    $row = Add-GuideFormula $Sheet $row "h = s_t − s_c"
    $row = Add-GuideParagraph $Sheet $row "r_eq - радиус круга, площадь которого равна площади бетонного элемента A. Это приближение используется только для восстановления внешних расстояний по дискретной сетке."
    $row = Add-GuideParagraph $Sheet $row "s_t - координата внешнего наиболее растянутого бетонного края по нормали n. s_c - координата внешнего наиболее сжатого края. Разность h = s_t - s_c дает высоту сечения по нормали к нейтральной линии."
    $row = Add-GuideParagraph $Sheet $row "Это инженерное обобщение для произвольного сечения и AutoCAD-import, где локальный контур элемента может быть неизвестен. Оно не является отдельной формулой СП 63."
    $row++

    $row = Add-GuideTitle $Sheet $row "6. Расстояние a и расчетная растянутая зона" 13
    $row = Add-GuideParagraph $Sheet $row "a измеряется не от центра крайнего бетонного элемента, а от восстановленного внешнего растянутого края до оси наиболее крайнего растянутого стержня по нормали к нейтральной линии."
    $row = Add-GuideFormula $Sheet $row "a = s_t − s_bar,max"
    $row = Add-GuideParagraph $Sheet $row "a - расстояние от внешнего растянутого края бетона до оси наиболее крайнего растянутого стержня. s_bar,max - максимальная координата оси растянутого стержня по той же нормали n."
    $row = Add-GuideParagraph $Sheet $row "Если арматура имеет несколько диаметров или рядов, выбирается именно крайний растянутый стержень из принятой расчетной группы."
    $row = Add-GuideParagraph $Sheet $row "SLS.Crack.TensionZoneMode = FullTension берет всю фактически растянутую зону. SLS.Crack.TensionZoneMode = Effective берет эффективную зону по логике п. 8.2.17 СП 63: ограничения 2a и 0.5h применяются вместе. Если фактическая растянутая зона меньше 2a, расчетная полоса может геометрически зайти за нейтральную линию; бетонные КЭ внутри этой полосы учитываются в Abt по положению центра."
    $row = Add-GuideFormula $Sheet $row "h_bt = min(max(x_t; 2a); 0.5h)"
    $row = Add-GuideParagraph $Sheet $row "h_bt - глубина эффективной зоны, а x_t - фактическая высота растянутой зоны до нейтральной линии. h - полный габарит сечения по нормали."
    $row = Add-GuideParagraph $Sheet $row "В режиме Effective бетонный КЭ входит в Abt, если его центр попал в принятую полосу h_bt. Дополнительная проверка растяжения для бетона здесь не применяется. Для As при этом выбираются только растянутые стержни, потому что sigma_s относится к растянутой арматуре."
    $row = Add-GuideParagraph $Sheet $row "Практические случаи для Effective: если x_t < 2a, глубина зоны принимается равной 2a, если только верхнее ограничение 0.5h не уменьшит ее. В этом случае по 2a расширяется именно бетонная зона Abt; сжатые стержни, попавшие в эту полосу геометрически, не включаются в As. Если 2a <= x_t <= 0.5h, принимается фактическая растянутая зона x_t. Если x_t > 0.5h, зона ограничивается глубиной 0.5h."
    $row = Add-GuideParagraph $Sheet $row "Abt считается дискретно: если центр бетонного элемента попал в принятую зону, учитывается вся его площадь; если не попал - площадь не учитывается. Частичные площади элементов текущая версия не считает."
    $row++

    $row = Add-GuideTitle $Sheet $row "7. Выбор арматуры и напряжения σs" 13
    $row = Add-GuideParagraph $Sheet $row "В As входят только растянутые стержни, оси которых находятся внутри принятой зоны по координате s. Для Effective при срабатывании нижнего ограничения 2a бетонная площадь Abt может увеличиться за счет полосы 2a, но сжатые стержни из этой полосы не добавляются в As, sigma_s и d_s,eq. Набор стержней выбирается один раз по текущему сочетанию и затем используется также для σs,crc в режиме Auto, если уточнение вообще потребовалось."
    $row = Add-GuideFractionFormula $Sheet $row "σ_s =" "Σ(A_s,i · σ_s,i)" "ΣA_s,i" ""
    $row = Add-GuideParagraph $Sheet $row "A_s,i - площадь i-го выбранного растянутого стержня. σ_s,i - напряжение в этом стержне, полученное из НДМ."
    $row = Add-GuideParagraph $Sheet $row "ΣA_s,i = A_s - суммарная площадь выбранной растянутой арматуры. σ_s - средневзвешенное напряжение арматуры, которое входит в формулу раскрытия трещины."
    $row = Add-GuideParagraph $Sheet $row "Такой расчет σs является инженерной НДМ-интерпретацией для произвольного сечения и косого изгиба. СП 63.13330.2018, п. 8.2.16 дает формулы для типовых схем, а программа берет напряжения непосредственно из НДМ по каждому стержню."
    $row++

    $row = Add-GuideTitle $Sheet $row "8. Эквивалентный диаметр при разных диаметрах" 13
    $row = Add-GuideParagraph $Sheet $row "СП 63 в формуле ls использует один диаметр ds и не раскрывает общий случай нескольких диаметров в выбранной группе стержней. Поэтому программа применяет инженерное дополнение по EN 1992-1-1."
    $row = Add-GuideFractionFormula $Sheet $row "d_s,eq =" "Σ(n_i·d_i²)" "Σ(n_i·d_i)" ""
    $row = Add-GuideParagraph $Sheet $row "d_s,eq - эквивалентный диаметр расчетной группы. d_i - диаметр стержней i-й группы, n_i - количество таких стержней."
    $row = Add-GuideParagraph $Sheet $row "Такой способ учета разных диаметров является инженерным дополнением. СП 63 не дает отдельной прямой формулы для этого случая."
    $row = Add-GuideParagraph $Sheet $row "Если все стержни одного диаметра, d_s,eq равен этому диаметру. Если диаметры разные, более крупные стержни получают больший вес, что инженерно согласуется с их большей площадью."
    $row++

    $row = Add-GuideTitle $Sheet $row "9. Расстояние между трещинами ls" 13
    $row = Add-GuideParagraph $Sheet $row "Базовое расстояние между нормальными трещинами считается по СП 63.13330.2018, п. 8.2.17, формула (8.136)."
    $row = Add-GuideFractionFormula $Sheet $row "l_s,0 = 0.5 ·" "A_bt" "A_s" "· d_s,eq      СП 63, п. 8.2.17, ф. (8.136)"
    $row = Add-GuideFormula $Sheet $row "l_s = max(l_s,0; 10·d_s; 100 мм)      СП 63, п. 8.2.17"
    $row = Add-GuideFormula $Sheet $row "l_s = min(l_s; 40·d_s; 400 мм)      СП 63, п. 8.2.17"
    $row = Add-GuideParagraph $Sheet $row "l_s,0 - расстояние между трещинами до применения нормативных ограничений. l_s - принятое расстояние после нижних и верхних ограничений."
    $row = Add-GuideParagraph $Sheet $row "A_bt - площадь бетона принятой растянутой зоны. A_s - площадь выбранной растянутой арматуры. d_s в ограничениях принимается как d_s,eq."
    $row = Add-GuideParagraph $Sheet $row "Если несколько ограничений действуют одновременно, сначала применяется нижняя граница, затем верхняя. В Results отдельно выводятся ls_raw до ограничений и окончательно принятое ls."
    $row++

    $row = Add-GuideTitle $Sheet $row "10. Коэффициенты φ1, φ2, φ3 и ψs" 13
    $row = Add-GuideParagraph $Sheet $row "Основная формула ширины раскрытия трещины соответствует СП 63.13330.2018, п. 8.2.15, формула (8.128). В текущей версии реализовано только продолжительное раскрытие: φ1 = 1.4."
    $row = Add-GuideFractionFormula $Sheet $row "a_crc = φ₁·φ₂·φ₃·ψ_s ·" "σ_s" "E_s" "· l_s      СП 63, п. 8.2.15, ф. (8.128)"
    $row = Add-GuideParagraph $Sheet $row "a_crc - расчетная ширина раскрытия нормальной трещины. E_s - модуль упругости арматуры. l_s - принятое расстояние между трещинами."
    $row = Add-GuideParagraph $Sheet $row "φ1 - коэффициент длительности действия нагрузки. В текущей версии считается только продолжительное раскрытие трещин, поэтому φ1 всегда принимается равным 1.4."
    $row = Add-GuideParagraph $Sheet $row "φ2 - коэффициент поверхности арматуры. Он задается пользователем напрямую в SLS.Crack.Phi2: обычный ориентир 0.5 для периодического профиля и 0.8 для гладкой арматуры."
    $row = Add-GuideParagraph $Sheet $row "φ3 - коэффициент характера работы элемента. При SLS.Crack.Phi3Mode = Auto программа принимает φ3 = 1.2 для растягивающей продольной силы N и φ3 = 1.0 для остальных реализованных случаев. При Phi3Mode = User берется SLS.Crack.Phi3."
    $row = Add-GuideParagraph $Sheet $row "ψ_s - коэффициент неравномерности деформаций растянутой арматуры между трещинами. При SLS.Crack.PsiMode = User берется SLS.Crack.PsiS."
    $row = Add-GuideParagraph $Sheet $row "При SLS.Crack.PsiMode = Auto программа сначала проверяет раскрытие с ψ_s = 1. Если a_crc <= a_crc,ult, расчет заканчивается, а σ_s,crc и λ_crc не считаются."
    $row = Add-GuideParagraph $Sheet $row "Если первая проверка Auto не проходит, программа ищет λ_crc и решает состояние λ_crc · (N, Mx, My) уже после образования трещины, то есть без работы растянутого бетона."
    $row = Add-GuideParagraph $Sheet $row "σ_s,crc определяется для той же расчетной группы стержней, которая была выбрана по текущему сочетанию для σ_s. Это важно: программа сравнивает напряжения одной и той же группы стержней до и сразу после образования трещины."
    $row = Add-GuideParagraph $Sheet $row "В общем изгибе в эту группу входят только стержни, которые в текущем сочетании растянуты и оси которых попали в принятую расчетную зону. В центральном растяжении берется вся продольная арматура, которая находится в растяжении."
    $row = Add-GuideParagraph $Sheet $row "Если режим Effective расширил бетонную зону Abt до 2a и эта полоса геометрически зашла за нейтральную линию, это влияет на Abt. Сжатые стержни из такой полосы в As, σ_s, d_s,eq и σ_s,crc не добавляются."
    $row = Add-GuideFractionFormula $Sheet $row "σ_s,crc =" "Σ(A_s,i · σ_s,i,crc)" "ΣA_s,i" ""
    $row = Add-GuideParagraph $Sheet $row "Здесь A_s,i - площадь i-го выбранного стержня, σ_s,i,crc - напряжение этого стержня из НДМ в состоянии λ_crc после образования трещины. Если из-за численного шума напряжение выбранного стержня получилось сжимающим, в среднем для раскрытия оно принимается равным нулю."
    $row = Add-GuideParagraph $Sheet $row "Что прямо сказано в СП: п. 8.2.18 вводит σ_s,crc как напряжение в продольной растянутой арматуре в сечении с трещиной сразу после образования нормальных трещин и использует его в формуле (8.137)."
    $row = Add-GuideParagraph $Sheet $row "Что является нашей реализацией: СП 63 не описывает отдельный алгоритм для произвольного дискретного сечения N + Mx + My с несколькими стержнями и разными диаметрами. Поэтому одно значение σ_s,crc программа получает как средневзвешенное по площади напряжение выбранной группы стержней."
    $row = Add-GuideFractionFormula $Sheet $row "ψ_s = 1 − 0.8 ·" "σ_s,crc" "σ_s" "      СП 63, п. 8.2.18, ф. (8.137)"
    $row = Add-GuideParagraph $Sheet $row "σ_s - средневзвешенное напряжение той же группы стержней в текущем сочетании. λ_crc - множитель нагрузки, при котором растянутый бетон достигает ε_bt,ult."
    $row = Add-GuideParagraph $Sheet $row "Критерий ε_bt,ult для λ_crc принят по СП 63.13330.2018, п. 8.2.14 и 8.1.30."
    $row++

    $row = Add-GuideTitle $Sheet $row "11. Итоговая проверка" 13
    $row = Add-GuideParagraph $Sheet $row "Условие проверки принято по СП 63.13330.2018, п. 8.2.6, формула (8.118). Допустимая ширина a_crc,ult всегда задается пользователем настройкой SLS.Crack.Allowable; программа не выбирает ее автоматически."
    $row = Add-GuideFormula $Sheet $row "a_crc ≤ a_crc,ult      СП 63, п. 8.2.6, ф. (8.118)"
    $row = Add-GuideFractionFormula $Sheet $row "SF_crc =" "a_crc,ult" "a_crc" ""
    $row = Add-GuideParagraph $Sheet $row "Обозначения: a_crc,ult - допустимая ширина раскрытия трещины, заданная пользователем; SF_crc - коэффициент запаса по раскрытию трещин. Если a_crc = 0, трещина не образована или расчет не применим, поэтому запас по раскрытию не ограничивает сечение."
    $row = Add-GuideParagraph $Sheet $row "В rngBatchSummary выводятся промежуточные величины h, a, TensionZoneDepth, EffectiveZoneDepth, As, Abt, d_s,eq, ls_raw, ls, φ1, φ2, φ3, Ncrc, λcrc, σs, σs,crc, ψs, Es и итоговые a_crc, a_crc,ult, CrackSafetyFactor."
    $row++

    $row = Add-GuideTitle $Sheet $row "12. Продольные трещины по сжатому бетону" 13
    $row = Add-GuideParagraph $Sheet $row "Эта проверка отделена от расчета ширины нормальных трещин a_crc. Она контролирует максимальное сжимающее напряжение в бетоне по уже найденному НДС сочетания и не запускает отдельный расчет равновесия."
    $row = Add-GuideParagraph $Sheet $row "По СП 35.13330.2011, п. 7.2 и таблице 7.1, образование продольных трещин относится к расчетам по второй группе предельных состояний. Поэтому программа выполняет эту проверку только как часть Calculation.Crack.Width. Если профиль не запрашивает раскрытие трещин, в Results выводится LongitudinalCrackStatus = N/A."
    $row = Add-GuideParagraph $Sheet $row "Программа просматривает бетонные элементы состояния с трещинами, берет максимальное по модулю сжимающее напряжение σ_c,max и сравнивает его с пользовательским Concrete.Rb.mc2. Если сжатого бетона нет, проверка неприменима: LongitudinalCrackStatus = N/A."
    $row = Add-GuideFormula $Sheet $row "σ_c,max ≤ R_b,mc2"
    $row = Add-GuideFractionFormula $Sheet $row "SF_long =" "R_b,mc2" "σ_c,max" ""
    $row = Add-GuideParagraph $Sheet $row "Concrete.Rb.mc2 вводится пользователем как итоговое допустимое значение для осевого сжатия при проверке продольных трещин. Программа не вычисляет и не умножает коэффициенты условий работы автоматически."
    $row = Add-GuideParagraph $Sheet $row "Нормативная привязка: СП 35.13330.2011, п. 7.100 требует проверять нормальные сжимающие напряжения от действующих нормативных нагрузок и воздействий. Таблица 7.6 выделяет расчетное сопротивление для осевого сжатия при проверках по предотвращению продольных трещин. Коэффициенты условий работы принимаются по применимым строкам таблицы 7.7; для этой проверки пользователь учитывает m_b13 и все другие нужные коэффициенты сам до ввода R_b,mc2."
    $row = Add-GuideParagraph $Sheet $row "То же правило принято для остальных характеристик бетона в Config: Rb, Rbt, Rb,ser и Rbt,ser считаются уже окончательно принятыми пользователем значениями с учетом нужных коэффициентов m.... Это исключает скрытое повторное применение коэффициентов внутри программы."
    $row++

    $row = Add-GuideTitle $Sheet $row "13. Принятые инженерные допущения" 13
    $row = Add-GuideParagraph $Sheet $row "1) Геометрические расстояния h и a для дискретных бетонных элементов восстанавливаются через эквивалентный круг r_eq = √(A/π). 2) Abt считается целыми элементами по попаданию центра в зону. 3) Для произвольного сечения используется один глобальный расчет σs и As, без локальных зон вокруг каждого стержня. 4) d_s,eq для разных диаметров взят по EN 1992-1-1 как инженерное дополнение. 5) Непродолжительное раскрытие и отдельный самостоятельный Mcrc не реализованы."
    $row = Add-GuideParagraph $Sheet $row "Если эти допущения будут пересмотрены, нужно одновременно изменить расчет трещин, вывод Results и этот раздел справки."
    return $row + 2
}

# Добавляет лист "Справка" и прямые гиперссылки "Подробнее" из Config на конкретные ячейки справки.
function Add-SettingsInstructions {
    param([object]$Workbook, [object]$ConfigSheet, [object]$InstructionSheet)

    Clear-GuideFormulaPictures $InstructionSheet
    $script:FormulaImageCounter = 0
    $script:FormulaImageRows.Clear()

    $InstructionSheet.Cells.Clear()
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

    $row = Add-CrackWidthMethodologyGuide $InstructionSheet
    $InstructionSheet.Cells.Item($row, 1).Value2 = "Справка по настройкам Config"
    $InstructionSheet.Cells.Item($row, 1).Font.Bold = $true
    $InstructionSheet.Cells.Item($row, 1).Font.Size = 16
    $row += 2
    $InstructionSheet.Cells.Item($row, 1).Value2 = "Нормативные источники"
    $InstructionSheet.Cells.Item($row, 1).Font.Bold = $true
    $row++
    $InstructionSheet.Cells.Item($row, 1).Value2 = "СП 63.13330.2018"
    $InstructionSheet.Cells.Item($row, 2).Value2 = "Бетонные и железобетонные конструкции. Основные положения. Используется как основной источник по НДМ и трещиностойкости; точные пункты указаны только там, где они уже подтверждены в документации проекта."
    $row++
    $InstructionSheet.Cells.Item($row, 1).Value2 = "СП 35.13330.2011"
    $InstructionSheet.Cells.Item($row, 2).Value2 = "Мосты и трубы. Используется как потенциальный специальный источник для мостовых параметров; непроверенные ссылки не подставляются автоматически."
    $row += 2
    $InstructionSheet.Cells.Item($row, 1).Value2 = "Важно"
    $InstructionSheet.Cells.Item($row, 2).Value2 = "Справка описывает фактически реализованное поведение программы. Если нормативная ссылка не подтверждена, это отмечается как требующее дальнейшей трассировки."
    $row += 2

    $InstructionSheet.Cells.Item($row, 1).Value2 = "Словарь пользовательских статусов"
    $InstructionSheet.Cells.Item($row, 1).Font.Bold = $true
    $InstructionSheet.Cells.Item($row, 1).Interior.Color = 15921906
    $InstructionSheet.Cells.Item($row, 2).Interior.Color = 15921906
    $row++
    $InstructionSheet.Cells.Item($row, 1).Value2 = "OK"
    $InstructionSheet.Cells.Item($row, 2).Value2 = "Расчет этой проверки выполнен успешно. Для прямого НДС равновесие найдено внутри физической диаграммы материала; для прочности или трещин проверка прошла."
    $row++
    $InstructionSheet.Cells.Item($row, 1).Value2 = "FAIL"
    $InstructionSheet.Cells.Item($row, 2).Value2 = "Численное решение есть, но инженерная проверка не проходит: превышена несущая способность, ширина нормальной трещины больше допуска, sigma_c,max больше Rb,mc2 по продольным трещинам или прямое НДС найдено только за пределами физической диаграммы."
    $row++
    $InstructionSheet.Cells.Item($row, 1).Value2 = "NumFail"
    $InstructionSheet.Cells.Item($row, 2).Value2 = "Программа не смогла найти равновесие или предельную точку с текущими численными настройками. Это не доказательство разрушения; нужно проверить диагностику, нагрузку, сетку и параметры расчета."
    $row++
    $InstructionSheet.Cells.Item($row, 1).Value2 = "InputErr"
    $InstructionSheet.Cells.Item($row, 2).Value2 = "Расчет не запускался из-за ошибки во входных данных: пустая или неизвестная настройка, недопустимая геометрия, материал, единицы или строка сочетания."
    $row++
    $InstructionSheet.Cells.Item($row, 1).Value2 = "N/A"
    $InstructionSheet.Cells.Item($row, 2).Value2 = "Проверка для этого сочетания не выполнялась и не считается ошибкой: например, она не включена выбранным профилем или неприменима к найденному состоянию."
    $InstructionSheet.Range($InstructionSheet.Cells.Item($row - 5, 1), $InstructionSheet.Cells.Item($row, 2)).Borders.LineStyle = 1
    $InstructionSheet.Range($InstructionSheet.Cells.Item($row - 5, 1), $InstructionSheet.Cells.Item($row, 2)).Borders.Color = 14277081
    $row += 3

    $anchors = @{}
    foreach ($item in (Get-SettingsInstructionCatalog)) {
        $anchors[$item.Key] = $row
        $InstructionSheet.Cells.Item($row, 1).Value2 = $item.Title
        $InstructionSheet.Cells.Item($row, 1).Font.Bold = $true
        $InstructionSheet.Cells.Item($row, 1).Interior.Color = 15921906
        $InstructionSheet.Cells.Item($row, 2).Interior.Color = 15921906
        $row++
        $lineNumber = 1
        foreach ($line in $item.Lines) {
            $InstructionSheet.Cells.Item($row, 1).Value2 = [string]$lineNumber
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
    $InstructionSheet.Range($InstructionSheet.Cells.Item(1, 1), $InstructionSheet.Cells.Item($row, 6)).VerticalAlignment = -4160
    $InstructionSheet.Range($InstructionSheet.Cells.Item(1, 1), $InstructionSheet.Cells.Item($row, 6)).Borders.LineStyle = 1
    $InstructionSheet.Range($InstructionSheet.Cells.Item(1, 1), $InstructionSheet.Cells.Item($row, 6)).Borders.Color = 14277081
    $InstructionSheet.Rows.AutoFit()
    foreach ($formulaRow in $script:FormulaImageRows) {
        $InstructionSheet.Rows.Item($formulaRow.Row).RowHeight = $formulaRow.Height
    }

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
        "Аннотации схемы" = "PlotAnnotationSettings"
        "Круглое сечение" = "CircleGeometry"
        "Скругленный прямоугольник" = "RoundedRectangleGeometry"
        "Г-образное сечение" = "LShapeGeometry"
    }
    for ($r = 1; $r -le ($used.Row + $used.Rows.Count + 40); $r++) {
        for ($c = 1; $c -le 40; $c++) {
            $caption = [string]$ConfigSheet.Cells.Item($r, $c).Value2
            if ($headerLinks.ContainsKey($caption)) {
                $targetKey = $headerLinks[$caption]
                if ($anchors.ContainsKey($targetKey)) {
                    Add-HeaderInstructionHyperlink $ConfigSheet.Cells.Item($r, $c) $caption $InstructionSheet.Name $anchors[$targetKey]
                }
            }
        }
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

    $loadHeaders = @("CombinationID", "N", "Mx", "My", "ProfileId", "CapacityLoadPath", "Comment")
    for ($i = 0; $i -lt $loadHeaders.Count; $i++) {
        $cell = $Sheet.Cells.Item($HeaderRow, $StartColumn + $i)
        $cell.Value2 = $loadHeaders[$i]
        $cell.Font.Bold = $true
        $cell.Interior.Color = $script:ConfigTableHeaderColor
    }
    $Sheet.Cells.Item($HeaderRow, $StartColumn + 1).Formula = "=`"N, `"&INDEX(rngUnitSettings,MATCH(`"Force`",INDEX(rngUnitSettings,,1),0),2)"
    $Sheet.Cells.Item($HeaderRow, $StartColumn + 2).Formula = "=`"Mx, `"&INDEX(rngUnitSettings,MATCH(`"Moment`",INDEX(rngUnitSettings,,1),0),2)"
    $Sheet.Cells.Item($HeaderRow, $StartColumn + 3).Formula = "=`"My, `"&INDEX(rngUnitSettings,MATCH(`"Moment`",INDEX(rngUnitSettings,,1),0),2)"

    $loadRange = $Sheet.Range($Sheet.Cells.Item($HeaderRow, $StartColumn), $Sheet.Cells.Item($HeaderRow + 20, $StartColumn + 6))
    $loadRange.Borders.LineStyle = 1
    $loadRange.Borders.Weight = 2
    $loadRange.Borders.Color = 12632256

    $profileListColumn = 132
    $capacityLoadPathListColumn = 133
    $profileOptions = @("PR1", "PR2", "PR3", "PR4")
    for ($i = 0; $i -lt $profileOptions.Count; $i++) {
        $Sheet.Cells.Item($i + 1, $profileListColumn).Value2 = $profileOptions[$i]
    }
    $profileColName = ConvertTo-ExcelColumn $profileListColumn
    $profileListAddress = "=$" + $profileColName + '$1:$' + $profileColName + '$' + $profileOptions.Count
    $profileRange = $Sheet.Range($Sheet.Cells.Item($HeaderRow + 1, $StartColumn + 4), $Sheet.Cells.Item($HeaderRow + 20, $StartColumn + 4))
    $profileRange.Validation.Delete()
    $profileRange.Validation.Add(3, 1, 1, $profileListAddress)
    $profileRange.Validation.IgnoreBlank = $false
    $profileRange.Validation.InCellDropdown = $true
    $Sheet.Cells.Item($HeaderRow + 1, $StartColumn + 4).Value2 = "PR1"

    $lambda = [char]0x03BB
    $capacityLoadPathOptions = @("$lambda*Mx", "$lambda*My", "$lambda*Mxy", "$lambda*N", "$lambda*NMxy")
    for ($i = 0; $i -lt $capacityLoadPathOptions.Count; $i++) {
        $Sheet.Cells.Item($i + 1, $capacityLoadPathListColumn).Value2 = $capacityLoadPathOptions[$i]
    }
    $capacityLoadPathColName = ConvertTo-ExcelColumn $capacityLoadPathListColumn
    $capacityLoadPathListAddress = "=$" + $capacityLoadPathColName + '$1:$' + $capacityLoadPathColName + '$' + $capacityLoadPathOptions.Count
    $capacityLoadPathRange = $Sheet.Range($Sheet.Cells.Item($HeaderRow + 1, $StartColumn + 5), $Sheet.Cells.Item($HeaderRow + 20, $StartColumn + 5))
    $capacityLoadPathRange.Validation.Delete()
    $capacityLoadPathRange.Validation.Add(3, 1, 1, $capacityLoadPathListAddress)
    $capacityLoadPathRange.Validation.IgnoreBlank = $true
    $capacityLoadPathRange.Validation.InCellDropdown = $true

    Set-WorkbookNameByBounds $Workbook "rngLoadCombinations" $Sheet $HeaderRow $StartColumn ($HeaderRow + 20) ($StartColumn + 6)
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

    for ($r = 1; $r -le 20; $r++) {
        $Sheet.Cells.Item($HeaderRow + $r, $StartColumn).Formula = '=IF(INDEX(rngLoadCombinations,' + ($r + 1) + ',1)="","",INDEX(rngLoadCombinations,' + ($r + 1) + ',1))'
        $Sheet.Cells.Item($HeaderRow + $r, $StartColumn + 1).Value2 = 0
        $Sheet.Cells.Item($HeaderRow + $r, $StartColumn + 2).Value2 = 0
        $Sheet.Cells.Item($HeaderRow + $r, $StartColumn + 3).Value2 = 0
    }

    Set-WorkbookNameByBounds $Workbook "rngStabilityDurationLoads" $Sheet $HeaderRow $StartColumn ($HeaderRow + 20) ($StartColumn + 3)
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
    $Sheet.Cells.Item($HeaderRow - 1, $StartColumn).Value2 = "Верхние значения для ненапрягаемой арматуры; используются только при Stability.Code = SP35."
    $Sheet.Range($Sheet.Cells.Item($HeaderRow - 1, $StartColumn), $Sheet.Cells.Item($HeaderRow - 1, $StartColumn + 7)).Merge() | Out-Null

    $headers = @("l0/b", "l0/d", "l0/i", "phi_m q=0", "phi_m q=0.25", "phi_m q=0.50", "phi_m q=1.00", "phi_l")
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
        @(43, 37.5, 150, 0.33, 0.28, 0.24, 0.21, 0.06)
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

# Создает прямую ссылку на ячейку листа "Справка" без использования именованных диапазонов.
function Add-InstructionHyperlink {
    param([object]$Cell, [string]$SheetName, [int]$TargetRow)
    $Cell.Hyperlinks.Delete()
    $Cell.Value2 = "Подробнее"
    $Cell.Parent.Hyperlinks.Add($Cell, "", "'" + $SheetName + "'!A" + $TargetRow, "", "Подробнее") | Out-Null
}

# Делает заголовок блока самодостаточным: название остается на месте, а пометка "(Подробнее)"
# показывает пользователю, что по заголовку можно перейти к справке.
function Add-HeaderInstructionHyperlink {
    param([object]$Cell, [string]$Caption, [string]$SheetName, [int]$TargetRow)
    $displayText = $Caption + " (Подробнее)"
    $Cell.Hyperlinks.Delete()
    $Cell.Value2 = $displayText
    $Cell.Parent.Hyperlinks.Add($Cell, "", "'" + $SheetName + "'!A" + $TargetRow, "", $displayText) | Out-Null
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

# Добавляет структурный элемент книги или отчета, сохраняя единый формат сборочных скриптов.
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

# Добавляет структурный элемент книги или отчета, сохраняя единый формат сборочных скриптов.
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

# Возвращает подготовленные данные или справочное значение для дальнейшего шага сборки.
function Get-PlotAnnotationSettingsCatalog {
    @(
        @("Enabled", "Yes", "Yes", "Yes/No", "Включает вывод соответствующего типа аннотаций на схеме."),
        @("Placement", "Outside", "Outside", "-", "Для арматуры задает сторону подписи от линии осей стержней; для размеров задает только сторону текста относительно размерной линии."),
        @("Offset", "60", "100", "мм", "Визуальный зазор в реальных миллиметрах сечения: для арматуры откладывается только от линии осей стержней, Placement задает сторону; для размеров - от грани до размерной линии."),
        @("LineEnabled", "Yes", "-", "Yes/No", "Показывать короткую линию обозначения арматуры; если No, остается только текст подписи."),
        @("LineWeight", "1.35", "1.75", "pt", "Толщина основной линии аннотации."),
        @("ExtensionLineWeight", "-", "0.85", "pt", "Толщина выносных линий; применяется только для размерных линий."),
        @("TextUnits", "pt", "pt", "-", "Единицы для TextHeight и TextGap: pt - обычные пункты Excel, mm - реальные миллиметры сечения."),
        @("TextHeight", "13", "13", "мм", "Высота текста. При TextUnits=mm задается в миллиметрах сечения; при TextUnits=pt - в пунктах Excel и не зависит от габарита сечения."),
        @("TextGap", "9", "9", "мм", "Зазор между линией аннотации и текстом. Единицы выбираются строкой TextUnits для соответствующего типа аннотаций."),
        @("ArrowType", "-", "Triangle", "-", "Тип стрелки размерной линии."),
        @("ArrowSize", "-", "Medium", "-", "Размер стрелок размерной линии."),
        @("Color", "20,30,90", "20,30,90", "RGB", "Цвет текста и основной линии в формате R,G,B."),
        @("ExtensionLineColor", "-", "140,140,140", "RGB", "Цвет выносных линий размеров в формате R,G,B.")
    )
}

# Добавляет структурный элемент книги или отчета, сохраняя единый формат сборочных скриптов.
function Add-PlotAnnotationSettingsTable {
    param([object]$Workbook, [object]$Sheet, [int]$HeaderRow, [int]$StartColumn)

    $title = "Аннотации схемы"
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

# Возвращает подготовленные данные или справочное значение для дальнейшего шага сборки.
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

# Устанавливает значение, оформление или именованный диапазон в книге через Excel COM.
function Set-InputUnitCell {
    param([object]$Cell, [string]$UnitText)
    $formula = Get-InputUnitFormulaForUnit $UnitText
    if ($null -ne $formula) {
        $Cell.Formula = $formula
    } else {
        $Cell.Value2 = $UnitText
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
    param([object]$Sheet, [int]$HeaderRow, [int]$StartColumn)

    function New-MaterialParameterRef {
        param([string]$RangeName, [string]$ParameterName, [int]$ValueColumn)
        return ('INDEX({0},MATCH("{1}",INDEX({0},,1),0),{2})' -f $RangeName, $ParameterName, $ValueColumn)
    }

    $Sheet.Cells.Item($HeaderRow - 1, $StartColumn).Value2 = "Контрольные точки диаграмм"
    $Sheet.Cells.Item($HeaderRow - 1, $StartColumn).Font.Bold = $true
    $Sheet.Cells.Item($HeaderRow - 1, $StartColumn).Interior.Color = 15921906
    $headers = @("Материал", "ГПС", "Тип", "Ветвь", "eps", "sigma")
    for ($i = 0; $i -lt $headers.Count; $i++) {
        $cell = $Sheet.Cells.Item($HeaderRow, $StartColumn + $i)
        $cell.Value2 = $headers[$i]
        $cell.Font.Bold = $true
        $cell.Interior.Color = 14277081
    }

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
        $rows.Add(@("Concrete", $group, "TwoLine", "Compression", ("=-{0}" -f $concreteEb2), ("=-{0}" -f $rb))) | Out-Null
        $rows.Add(@("Concrete", $group, "TwoLine", "Compression", ("=-{0}" -f $concreteEb1Red), ("=-{0}" -f $rb))) | Out-Null
        $rows.Add(@("Concrete", $group, "TwoLine", "Zero", "0", "0")) | Out-Null
        $rows.Add(@("Concrete", $group, "TwoLine", "Tension", ("={0}" -f $concreteEbt1Red), ("={0}" -f $rbt))) | Out-Null
        $rows.Add(@("Concrete", $group, "TwoLine", "Tension", ("={0}" -f $concreteEbt2), ("={0}" -f $rbt))) | Out-Null
        $rows.Add(@("Concrete", $group, "ThreeLine", "Compression", ("=-{0}" -f $concreteEb2), ("=-{0}" -f $rb))) | Out-Null
        $rows.Add(@("Concrete", $group, "ThreeLine", "Compression", ("=-{0}" -f $concreteEb0), ("=-{0}" -f $rb))) | Out-Null
        $rows.Add(@("Concrete", $group, "ThreeLine", "Compression", ("=-0.6*{0}/{1}" -f $rb, $eb), ("=-0.6*{0}" -f $rb))) | Out-Null
        $rows.Add(@("Concrete", $group, "ThreeLine", "Zero", "0", "0")) | Out-Null
        $rows.Add(@("Concrete", $group, "ThreeLine", "Tension", ("=0.6*{0}/{1}" -f $rbt, $ebt), ("=0.6*{0}" -f $rbt))) | Out-Null
        $rows.Add(@("Concrete", $group, "ThreeLine", "Tension", ("={0}" -f $concreteEbt0), ("={0}" -f $rbt))) | Out-Null
        $rows.Add(@("Concrete", $group, "ThreeLine", "Tension", ("={0}" -f $concreteEbt2), ("={0}" -f $rbt))) | Out-Null
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
        $rows.Add(@("Steel", $group, "TwoLine", "Compression", ("=-{0}" -f $steelTwoLineEsc2), ("=-{0}" -f $rsc))) | Out-Null
        $rows.Add(@("Steel", $group, "TwoLine", "Compression", ("=-{0}/{1}" -f $rsc, $esc), ("=-{0}" -f $rsc))) | Out-Null
        $rows.Add(@("Steel", $group, "TwoLine", "Zero", "0", "0")) | Out-Null
        $rows.Add(@("Steel", $group, "TwoLine", "Tension", ("={0}/{1}" -f $rs, $es), ("={0}" -f $rs))) | Out-Null
        $rows.Add(@("Steel", $group, "TwoLine", "Tension", ("={0}" -f $steelTwoLineEs2), ("={0}" -f $rs))) | Out-Null
        $rows.Add(@("Steel", $group, "ThreeLine", "Compression", ("=-{0}" -f $steelThreeLineEsc2), ("=-1.1*{0}" -f $rsc))) | Out-Null
        $rows.Add(@("Steel", $group, "ThreeLine", "Compression", ("=-{0}" -f $epsScPl), ("=-1.1*{0}" -f $rsc))) | Out-Null
        $rows.Add(@("Steel", $group, "ThreeLine", "Compression", ("=-{0}" -f $epsSc0), ("=-{0}" -f $rsc))) | Out-Null
        $rows.Add(@("Steel", $group, "ThreeLine", "Compression", ("=-{0}" -f $epsSc1), ("=-0.9*{0}" -f $rsc))) | Out-Null
        $rows.Add(@("Steel", $group, "ThreeLine", "Zero", "0", "0")) | Out-Null
        $rows.Add(@("Steel", $group, "ThreeLine", "Tension", ("={0}" -f $epsS1), ("=0.9*{0}" -f $rs))) | Out-Null
        $rows.Add(@("Steel", $group, "ThreeLine", "Tension", ("={0}" -f $epsS0), ("={0}" -f $rs))) | Out-Null
        $rows.Add(@("Steel", $group, "ThreeLine", "Tension", ("={0}" -f $epsSPl), ("=1.1*{0}" -f $rs))) | Out-Null
        $rows.Add(@("Steel", $group, "ThreeLine", "Tension", ("={0}" -f $steelThreeLineEs2), ("=1.1*{0}" -f $rs))) | Out-Null
    }

    for ($r = 0; $r -lt $rows.Count; $r++) {
        for ($c = 0; $c -lt $rows[$r].Count; $c++) {
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

    $dataRange = $Sheet.Range($Sheet.Cells.Item($HeaderRow, $StartColumn), $Sheet.Cells.Item($HeaderRow + $rows.Count, $StartColumn + 5))
    $dataRange.Borders.LineStyle = 1
    $dataRange.Borders.Weight = 2
    $dataRange.Columns.AutoFit() | Out-Null
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
        $left $top $width $height ($ControlHeaderRow + 1) ($ControlHeaderRow + 5) ($ControlHeaderRow + 6) ($ControlHeaderRow + 12) $epsColumn $sigmaColumn
    Add-MaterialDiagramChart $Sheet "chMaterialSteelDiagram" "Диаграмма растяжения/сжатия арматуры (I ГПС)" `
        $left ($top + $height + $gap) $width $height ($ControlHeaderRow + 25) ($ControlHeaderRow + 29) ($ControlHeaderRow + 30) ($ControlHeaderRow + 38) $epsColumn $sigmaColumn
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
        [int]$SigmaColumn
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

    try {
        $chart.Axes(1).HasTitle = $true
        $chart.Axes(1).AxisTitle.Text = "eps"
        $chart.Axes(2).HasTitle = $true
        $chart.Axes(2).AxisTitle.Text = "sigma, MPa"
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

# Выполняет служебный шаг сборочного или проверочного сценария.
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
    Add-MaterialDiagramControlTables $Sheet $materialControlHeaderRow $materialControlColumn
    Add-SP35Table721 $Workbook $Sheet 60 35

    Add-UnitSettingsTable $Workbook $Sheet $rightRow $rightColumn
    $rightRow += (Get-UnitSettingsCatalog).Count + 1 + $rightBlockGap

    Add-LoadCombinationsTable $Workbook $Sheet $loadCombinationsHeaderRow $loadCombinationsColumn
    Add-StabilityDurationLoadsTable $Workbook $Sheet 27 16

    Add-SignConventionSettingsTable $Workbook $Sheet $rightRow $rightColumn
    $rightRow += (Get-SignConventionSettingsCatalog).Count + 1 + $rightBlockGap

    $steelRows = Get-SteelMaterialParametersCatalog
    Add-MaterialParameterTable $Workbook $Sheet "rngSteelMaterialParameters" $rightRow $rightColumn "Материал арматуры" $steelRows
    $rightRow += $steelRows.Count + 1 + $rightBlockGap

    $concreteRows = Get-ConcreteMaterialParametersCatalog
    Add-MaterialParameterTable $Workbook $Sheet "rngConcreteMaterialParameters" $rightRow $rightColumn "Материал бетона" $concreteRows
    $rightRow += $concreteRows.Count + 1 + $rightBlockGap

    Add-CalculationProfilesTable $Workbook $Sheet $rightRow $rightColumn
    $rightRow += (Get-CalculationProfilesCatalog).Count + 2 + $rightBlockGap

    Add-PlotAnnotationSettingsTable $Workbook $Sheet $rightRow $rightColumn
    $rightRow += (Get-PlotAnnotationSettingsCatalog).Count + 1 + $rightBlockGap

    foreach ($geometryTable in (Get-GeometrySettingsCatalog)) {
        if ($geometryTable.ContainsKey("FaceTable") -and $geometryTable.FaceTable) {
            Add-LShapeFaceSettingsTable $Workbook $Sheet $geometryTable.RangeName $rightRow $rightColumn $geometryTable.Title $geometryTable.Rows
            $rightRow += 23 + $rightBlockGap
        } else {
            Add-GeometrySettingsTable $Workbook $Sheet $geometryTable.RangeName $rightRow $rightColumn $geometryTable.Title $geometryTable.Rows
            $rightRow += $geometryTable.Rows.Count + 1 + $rightBlockGap
        }
    }

    Add-MaterialDiagramCharts $Sheet $materialChartTopRow $materialChartColumn $materialControlHeaderRow $materialControlColumn

    $validationLists = [ordered]@{
        "Geometry.Source" = @("Generated", "AutoCAD")
        "General.ExecutionReportEnabled" = @("Yes", "No")
        "General.NonCriticalMessagesEnabled" = @("Yes", "No")
        "Stability.Code" = @("SP63", "SP35")
        "Stability.SystemType" = @("Determinate", "Indeterminate")
        "Stability.ZeroMomentEccentricitySign1" = @("1", "-1")
        "Stability.ZeroMomentEccentricitySign2" = @("1", "-1")
        "Stability.PhiLMode" = @("Auto", "PhiL2")
        "Geometry.Type" = @("RoundedRectangle", "Circle", "LShape")
        "Solver.Method" = @("Newton", "Secant")
        "Solver.DirectState.DiagramExtension" = @("Yes", "No")
        "Capacity.SolutionStrategy" = @("Auto", "UltimateStrain", "LoadMultiplier")
        "Capacity.SearchMethod" = @("Bisection", "Brent", "Secant")
        "Solver.LineSearchEnabled" = @("Yes", "No")
        "SLS.Crack.Phi3Mode" = @("Auto", "User")
        "SLS.Crack.PsiMode" = @("User", "Auto")
        "SLS.Crack.TensionZoneMode" = @("Effective", "FullTension")
        "AutoCAD.Export.LabelMode" = @("ValuesOnly", "NamesAndValues")
        "AutoCAD.Export.NeutralLineEnabled" = @("Yes", "No")
        "AutoCAD.Export.PrincipalAxesEnabled" = @("Yes", "No")
        "AutoCAD.Export.LoadPointEnabled" = @("Yes", "No")
        "Plot.Enabled" = @("Yes", "No")
        "Plot.AutoUpdateAfterCalculation" = @("Yes", "No")
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
    for ($r = 2; $r -le $settingsRange.Rows.Count; $r++) {
        if ([string]$settingsRange.Cells.Item($r, 1).Value2 -eq "AutoCAD.Export.CombinationID" -or [string]$settingsRange.Cells.Item($r, 1).Value2 -eq "Plot.LoadCase") {
            $cell = $settingsRange.Cells.Item($r, 2)
            $cell.Validation.Delete()
            $cell.Validation.Add(3, 1, 1, $exportCombinationListAddress)
            $cell.Validation.IgnoreBlank = $false
            $cell.Validation.InCellDropdown = $true
        }
    }
    $listColumn++

    Add-CalculationProfilesValidation $Sheet $listColumn
    $listColumn += 7

    Add-PlotAnnotationValidation $Sheet $listColumn

    $Sheet.Columns.Item(1).ColumnWidth = 34
    $Sheet.Columns.Item(2).ColumnWidth = 18
    $Sheet.Columns.Item(3).ColumnWidth = 12
    $Sheet.Columns.Item(4).ColumnWidth = 20
    $Sheet.Columns.Item(5).ColumnWidth = 11
    $Sheet.Columns.Item(6).ColumnWidth = 11
    $Sheet.Columns.Item(7).ColumnWidth = 11
    $Sheet.Columns.Item(8).ColumnWidth = 32
    foreach ($colIndex in @(9, 10, 11, 12, 13)) {
        $Sheet.Columns.Item($colIndex).ColumnWidth = 17
    }
    for ($colIndex = 14; $colIndex -le 69; $colIndex++) {
        $Sheet.Columns.Item($colIndex).ColumnWidth = 8.43
    }
    $Sheet.Columns.Item(15).ColumnWidth = 15
    $Sheet.Columns.Item(16).ColumnWidth = 11
    $Sheet.Columns.Item(17).ColumnWidth = 11
    $Sheet.Columns.Item(18).ColumnWidth = 11
    $Sheet.Columns.Item(19).ColumnWidth = 15
    $Sheet.Columns.Item(20).ColumnWidth = 18
    $Sheet.Columns.Item(21).ColumnWidth = 22
    $Sheet.Columns.Item("BR:EF").Hidden = $true
    $Sheet.Range("A1:EF260").Font.Name = "Arial"
    $Sheet.Range("A1:EF260").Font.Size = 9
    Apply-ConfigUserInputAlignment $Workbook
    Apply-ConfigCommentColumnAlignment $Workbook
    Apply-ConfigShortColumnAlignment $Workbook
    Apply-ConfigNamedRangeBorders $Workbook

    return @{ Settings = $settingsRange }
}

# Добавляет структурный элемент книги или отчета, сохраняя единый формат сборочных скриптов.
function Add-PlotAnnotationValidation {
    param([object]$Sheet, [int]$StartColumn)

    $validationSources = @{
        "Enabled" = @{ Column = $StartColumn; Values = @("Yes", "No") }
        "Placement" = @{ Column = ($StartColumn + 1); Values = @("Outside", "Inside") }
        "TextUnits" = @{ Column = ($StartColumn + 4); Values = @("mm", "pt") }
        "LineEnabled" = @{ Column = $StartColumn; Values = @("Yes", "No") }
        "ArrowType" = @{ Column = ($StartColumn + 2); Values = @("Triangle", "Stealth", "Diamond", "Oval", "Open") }
        "ArrowSize" = @{ Column = ($StartColumn + 3); Values = @("Small", "Medium", "Wide") }
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

# Добавляет валидацию для профильной таблицы. Типы расчетов, материальные
# модели и визуализация выбираются в PR-колонках профилей.
function Add-CalculationProfilesValidation {
    param([object]$Sheet, [int]$ListColumn)

    $sources = @{
        "Calculation.Strength.DirectState" = @{ Column = $ListColumn; Values = @("Yes", "No") }
        "Calculation.Strength.Capacity" = @{ Column = $ListColumn; Values = @("Yes", "No") }
        "Calculation.Crack.Width" = @{ Column = $ListColumn; Values = @("Yes", "No") }
        "Calculation.Stability.Enabled" = @{ Column = $ListColumn; Values = @("Yes", "No") }
        "MaterialModel.Stability.ValueSet" = @{ Column = ($ListColumn + 1); Values = @("ULS(I)", "SLS(II)") }
        "MaterialModel.Strength.ValueSet" = @{ Column = ($ListColumn + 1); Values = @("ULS(I)", "SLS(II)") }
        "MaterialModel.CrackInitiation.ValueSet" = @{ Column = ($ListColumn + 1); Values = @("ULS(I)", "SLS(II)") }
        "MaterialModel.CrackedState.ValueSet" = @{ Column = ($ListColumn + 1); Values = @("ULS(I)", "SLS(II)") }
        "MaterialModel.Strength.ConcreteDiagram" = @{ Column = ($ListColumn + 2); Values = @("TwoLine", "ThreeLine") }
        "MaterialModel.CrackInitiation.ConcreteDiagram" = @{ Column = ($ListColumn + 2); Values = @("TwoLine", "ThreeLine") }
        "MaterialModel.CrackedState.ConcreteDiagram" = @{ Column = ($ListColumn + 2); Values = @("TwoLine", "ThreeLine") }
        "MaterialModel.Strength.ConcreteTension" = @{ Column = ($ListColumn + 3); Values = @("Ignore", "UseDiagram") }
        "MaterialModel.Strength.SteelDiagram" = @{ Column = ($ListColumn + 4); Values = @("TwoLine", "ThreeLine") }
        "MaterialModel.CrackInitiation.SteelDiagram" = @{ Column = ($ListColumn + 4); Values = @("TwoLine", "ThreeLine") }
        "MaterialModel.CrackedState.SteelDiagram" = @{ Column = ($ListColumn + 4); Values = @("TwoLine", "ThreeLine") }
        "Visualization.State" = @{ Column = ($ListColumn + 5); Values = @("StrengthState", "CapacityState", "BeforeMcrcState", "AfterMcrcState", "CrackedState") }
        "Visualization.Quantity" = @{ Column = ($ListColumn + 6); Values = @("Stress", "Strain") }
    }

    foreach ($source in $sources.Values) {
        for ($i = 0; $i -lt $source.Values.Count; $i++) {
            $Sheet.Cells.Item($i + 1, $source.Column).Value2 = $source.Values[$i]
        }
    }

    try {
        $range = $Sheet.Parent.Names.Item("rngCalculationProfiles").RefersToRange
        for ($r = 3; $r -le $range.Rows.Count; $r++) {
            $key = [string]$range.Cells.Item($r, 2).Value2
            if ($sources.ContainsKey($key)) {
                $source = $sources[$key]
                $colName = ConvertTo-ExcelColumn $source.Column
                $listAddress = "=$" + $colName + '$1:$' + $colName + '$' + $source.Values.Count
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

# Добавляет структурный элемент книги или отчета, сохраняя единый формат сборочных скриптов.
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
        if ([string]$Rows[$r][0] -eq "Дополнительные ряды арматуры") {
            $separator = $Sheet.Range($Sheet.Cells.Item($HeaderRow + 1 + $r, $StartColumn), $Sheet.Cells.Item($HeaderRow + 1 + $r, $StartColumn + 3))
            $separator.Font.Bold = $true
            $separator.Interior.Color = $script:ConfigSubgroupHeaderColor
        }
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

# Добавляет структурный элемент книги или отчета, сохраняя единый формат сборочных скриптов.
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
        @("H1 - левая",   $Rows[12][1], $Rows[16][1], $Rows[18][1], $Rows[14][1], $Rows[17][1], $Rows[19][1], "мм", "Дополнительные ряды у левой грани H1 (грань _1)."),
        @("H1 - правая",  $Rows[13][1], $Rows[16][1], $Rows[18][1], $Rows[15][1], $Rows[17][1], $Rows[19][1], "мм", "Дополнительные ряды у правой грани H1 (грань _2)."),
        @("B1 - верхняя", $Rows[12][2], $Rows[16][2], $Rows[18][2], $Rows[14][2], $Rows[17][2], $Rows[19][2], "мм", "Дополнительные ряды у верхней грани B1 (грань _1)."),
        @("B1 - нижняя",  $Rows[13][2], $Rows[16][2], $Rows[18][2], $Rows[15][2], $Rows[17][2], $Rows[19][2], "мм", "Дополнительные ряды у нижней грани B1 (грань _2)."),
        @("H2 - левая",   $Rows[12][3], $Rows[16][3], $Rows[18][3], $Rows[14][3], $Rows[17][3], $Rows[19][3], "мм", "Дополнительные ряды у левой грани H2 (грань _1)."),
        @("H2 - правая",  $Rows[13][3], $Rows[16][3], $Rows[18][3], $Rows[15][3], $Rows[17][3], $Rows[19][3], "мм", "Дополнительные ряды у правой грани H2 (грань _2)."),
        @("B2 - верхняя", $Rows[12][4], $Rows[16][4], $Rows[18][4], $Rows[14][4], $Rows[17][4], $Rows[19][4], "мм", "Дополнительные ряды у верхней грани B2 (грань _1)."),
        @("B2 - нижняя",  $Rows[13][4], $Rows[16][4], $Rows[18][4], $Rows[15][4], $Rows[17][4], $Rows[19][4], "мм", "Дополнительные ряды у нижней грани B2 (грань _2).")
    )

    $Sheet.Cells.Item($HeaderRow, $StartColumn).Value2 = "Геометрия (H1/B1 - верхний прямоугольник; H2/B2 - нижний прямоугольник)"
    $Sheet.Cells.Item($HeaderRow, $StartColumn).Font.Bold = $true
    $Sheet.Cells.Item($HeaderRow, $StartColumn).Interior.Color = 15921906
    $Sheet.Cells.Item($HeaderRow, $StartColumn).WrapText = $false
    $geomHeaders = @("H1", "B1", "H2", "B2", "Ед.")
    for ($i = 0; $i -lt $geomHeaders.Count; $i++) {
        $cell = $Sheet.Cells.Item($HeaderRow + 1, $StartColumn + $i)
        $cell.Value2 = $geomHeaders[$i]
        $cell.Font.Bold = $true
        $cell.Interior.Color = 14277081
        if ($i -lt 4) {
            $Sheet.Cells.Item($HeaderRow + 2, $StartColumn + $i).Value2 = $geometryValues[$geomHeaders[$i]]
        }
    }
    Set-InputUnitCell $Sheet.Cells.Item($HeaderRow + 2, $StartColumn + 4) "мм"

    $mainHeaderRow = $HeaderRow + 4
    $Sheet.Cells.Item($mainHeaderRow - 1, $StartColumn).Value2 = "Основное армирование"
    $mainSectionRange = $Sheet.Range($Sheet.Cells.Item($mainHeaderRow - 1, $StartColumn), $Sheet.Cells.Item($mainHeaderRow - 1, $StartColumn + 7))
    $mainSectionRange.Font.Bold = $true
    $mainSectionRange.Interior.Color = 15921906
    $mainSectionRange.HorizontalAlignment = -4131
    $mainSectionRange.VerticalAlignment = -4108
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

    $extraHeaderRow = $mainHeaderRow + 10
    $Sheet.Cells.Item($extraHeaderRow - 1, $StartColumn).Value2 = "Дополнительные ряды"
    $extraSectionRange = $Sheet.Range($Sheet.Cells.Item($extraHeaderRow - 1, $StartColumn), $Sheet.Cells.Item($extraHeaderRow - 1, $StartColumn + 8))
    $extraSectionRange.Font.Bold = $true
    $extraSectionRange.Interior.Color = 15921906
    $extraSectionRange.HorizontalAlignment = -4131
    $extraSectionRange.VerticalAlignment = -4108
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
    $listColumn = 130
    for ($i = 0; $i -lt $locOptions.Count; $i++) {
        $Sheet.Cells.Item($i + 1, $listColumn).Value2 = $locOptions[$i]
    }
    $listColName = ConvertTo-ExcelColumn $listColumn
    $listAddress = "=$" + $listColName + '$1:$' + $listColName + '$' + $locOptions.Count
    $bindOptions = @("EachBar", "EverySecondBar")
    $bindListColumn = 131
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

    $Sheet.Range($Sheet.Cells.Item($mainHeaderRow, $StartColumn + 7), $Sheet.Cells.Item($mainHeaderRow + $mainRows.Count, $StartColumn + 7)).HorizontalAlignment = -4152
    $Sheet.Range($Sheet.Cells.Item($extraHeaderRow, $StartColumn + 8), $Sheet.Cells.Item($extraHeaderRow + $extraRows.Count, $StartColumn + 8)).HorizontalAlignment = -4152

    Set-WorkbookNameByBounds $Workbook $RangeName $Sheet $HeaderRow $StartColumn ($extraHeaderRow + $extraRows.Count) ($StartColumn + 8)
}

# Добавляет структурный элемент книги или отчета, сохраняя единый формат сборочных скриптов.
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

# Добавляет структурный элемент книги или отчета, сохраняя единый формат сборочных скриптов.
function Add-SystemSettings {
    param([object]$Sheet)
    $result = Apply-SystemSettingsLayout $Sheet.Parent $Sheet
    $result
}


