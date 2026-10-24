VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CBatchSectionCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Пакетный расчёт сочетаний нагрузок
' ==========================================================================
' Выполняет для каждого сочетания проверки, включённые его расчётным профилем:
' НДС, несущую способность, устойчивость и трещины. Сохраняет исходные нагрузки
' и результаты для записи на лист Results без повторного решения равновесия.
' Получает готовую модель CSectionModel, материалы и усилия N + Mx + My
' во внутренних единицах; чтение ячеек Excel и построение геометрии выполняются отдельно.

Private Const INITIAL_COMBINATION_CAPACITY As Long = 20 ' Начальный размер внутренних массивов; при расширении rngLoadCombinations хранилище растет.
Private Const COMBINATION_CAPACITY_STEP As Long = 20 ' Число дополнительных мест при расширении массивов, чтобы не копировать их для каждой новой строки сочетаний.
Private Const CAPACITY_PATH_NONE As String = "None" ' Обозначение отсутствия запущенного расчёта несущей способности для сочетания.
Private Const DEFAULT_ZERO_MOMENT_PER_DEPTH As Double = 4903.325 ' 0.5 tf*m/m во внутренних Н*мм/мм.

' Исходные данные одной строки сочетаний хранятся отдельно от результатов.
' Это позволяет сохранить нагрузки пользователя независимо от расчётных усилий.
Private Type TCombinationInput
    CombinationID As String ' Номер сочетания из входной таблицы.
    SourceDataOffset As Long ' Смещение строки сочетания от начала входной таблицы.
    SourceReference As String ' Фактический адрес исходной строки, переданный ридером без зависимости расчёта от Excel.
    N As Double ' Продольная сила во внутренних единицах, Н.
    Mx As Double ' Момент Mx во внутренних единицах, Н*мм.
    My As Double ' Момент My во внутренних единицах, Н*мм.
    ProfileId As String ' Имя профиля, задающего нужные проверки и материалы.
    LoadPath As String ' Способ изменения составляющих нагрузки при поиске предела.
    Comment As String ' Пользовательское пояснение сочетания.
    InputValid As Boolean ' Признак корректных исходных данных сочетания.
    InputError As String ' Причина ошибки исходных данных сочетания.
End Type

Private mSection As CSectionModel ' Одна расчётная модель сечения для всех сочетаний нагрузок.
Private mMaterialProvider As CMaterialModelProvider ' Подготовка диаграмм материалов для прочности, образования трещины и НДС с трещинами.
Private mProfiles As CCalculationProfileCatalog ' Каталог расчетных профилей PR1/PR2/... из Config.
Private mStatusPolicy As CResultStatusPolicy ' Единый словарь пользовательских статусов и правил их агрегирования.
Private mCount As Long ' Сколько сочетаний реально добавлено в пакет.

Private mInputs() As TCombinationInput ' Исходные данные LC, прочитанные из фактических строк rngLoadCombinations.
Private mWorkingLoads() As CSectionLoadState ' Рабочие усилия текущего запуска; исходная строка не изменяется.
Private mWorkingPaths() As String ' Проверенный выбор изменения нагрузки для запуска расчёта; исходный текст пользователя хранится отдельно.
Private mResults() As CCombinationResult ' Результаты каждого сочетания для записи сводки и Results; массив растёт с числом сочетаний.
Private mLoadReferenceX As Double ' X точки приложения нагрузки относительно расчетной модели, мм.
Private mLoadReferenceY As Double ' Y точки приложения нагрузки относительно расчетной модели, мм.
Private mLoadReferenceOffsetX As Double ' Смещение точки нагрузки от центра тяжести бетонного сечения по X, мм.
Private mLoadReferenceOffsetY As Double ' Смещение точки нагрузки от центра тяжести бетонного сечения по Y, мм.
Private mMomentZeroFilter As CMomentZeroFilter ' Общий инженерный фильтр практически нулевых моментов.
Private mZeroMomentPerDepth As Double ' Порог момента на единицу бетонного габарита, Н*мм/мм.
Private mConcreteDepthForMx As Double ' Бетонный габарит по Y для проверки Mx, мм.
Private mConcreteDepthForMy As Double ' Бетонный габарит по X для проверки My, мм.
Private mGoverningCombinationIndex As Long ' Номер сочетания, определяющего результат проверки прочности.
Private mCrackGoverningCombinationIndex As Long ' Номер сочетания с минимальным запасом по раскрытию трещин.
Private mWorstCombinationIndex As Long ' Номер сочетания Worst по критерию, выбранному пользователем для общей сводки.
Private mWorstCombinationCriterion As String ' Критерий выбора определяющего сочетания: деформации, несущая способность, трещины или устойчивость.
Private mSolverCallCount As Long ' Число вызовов CSectionSolver.Solve во всём пакетном расчёте.
Private mSolverIterationCount As Long ' Общее число фактически выполненных итераций решения НДС.
Private mUltimateStrainStartCount As Long ' Число попыток поиска предела с разными начальными полями деформаций.
Private mUltimateStrainIterationCount As Long ' Общее число итераций UltimateStrain для прочности и образования трещины.
Private mElapsedSeconds As Double ' Полное время пакетного расчета, сек.
Private mDiagnosticLog As String ' Краткий журнал выполнения отдельных сочетаний нагрузок.
Private mTransformedReferenceValid(1 To 2) As Boolean ' Сохранённые центры приведённого сечения для I и II ГПС в пределах одного запуска.
Private mTransformedReferenceX(1 To 2) As Double ' X центра приведенного сечения, мм.
Private mTransformedReferenceY(1 To 2) As Double ' Y центра приведенного сечения, мм.
Private mCrackGeometryQuery As CSectionGeometryQuery ' Общая бетонная область текущего Execute; готовится при первом обращении отдельно от расчётных НДС.
Private mCrackDomainBuildCount As Long ' Количество построений общей области; сам объект геометрических запросов освобождается до выхода из Execute.
Private mRunTransformedProperties(1 To 2) As CSectionPropertiesCalculator ' Приведённые характеристики для двух наборов модулей материалов текущего запуска.
Private mRunConcreteProperties As CSectionPropertiesCalculator ' Общие массивы геометрии для расчёта усилий; готовятся при первом обращении отдельно от области трещин.
Private mMechanicalPreparationBuildCount As Long ' Число фактических подготовок геометрических характеристик за запуск.
Private mExecuting As Boolean ' Временные общие данные подготовки, используемые только внутри Execute и освобождаемые перед выдачей результата.

Private mCrackAllowable As Double ' Допустимая ширина раскрытия трещины a_crc,ult, мм.
Private mCrackStandardCode As String ' Общий выбор SP63/SP35, независимый от Stability.Code.
Private mCrackSP35GroupGap As Double, mCrackSP35RowTolerance As Double ' Допуск объединения стержней в группу и допуск разделения рядов, мм.
Private mCrackSP35NeighborRatioLimit As Double ' Порог t/L для бокового ограничения ближайшим соседом.
Private mCrackSP35DiameterMode As String, mCrackSP35RadiusMode As String, mCrackSP35RebarProfile As String ' Выбор диаметра группы, радиуса взаимодействия и профиля арматуры для трещин СП 35.
Private mCrackPsiMode As String ' User, Auto или AlwaysCalc: как определять psi_s для нормальных трещин.
Private mCrackSigmaSCrcAveragingMode As String ' Как усреднять sigma_s,crc при уточнении psi_s по PostCrackState.
Private mCrackPhi1 As Double ' Пользовательский коэффициент phi1 для формулы раскрытия.
Private mCrackPhi2 As Double ' Пользовательский коэффициент phi2 для формулы раскрытия.
Private mCrackPhi3Mode As String ' User или Auto: как принимать phi3.
Private mCrackPhi3 As Double ' Пользовательский коэффициент phi3, если Phi3Mode = User.
Private mCrackPsiS As Double ' Пользовательский коэффициент psi_s, если PsiMode = User.
Private mCrackTensionZoneMode As String ' Effective или FullTension: какая бетонная зона входит в Abt.
Private mCrackCoverDistanceMode As String ' Как измерять a_s: до ближайшего контура или до дальней растянутой линии.
Private mCrackFormationSolutionStrategy As String ' Auto, UltimateStrain или LoadMultiplier для поиска уровня трещинообразования.
Private mSolverMethod As String ' Newton или Secant для прямого равновесия.
Private mSolverLoadSteps As Long ' Число ступеней приложения нагрузки в CSectionSolver.
Private mSolverMaxIterations As Long ' Максимальное число итераций решения НДС на одной ступени нагрузки.
Private mSolverToleranceN As Double ' Допуск равновесия по N, Н.
Private mSolverToleranceMx As Double ' Допуск равновесия по Mx, Н*мм.
Private mSolverToleranceMy As Double ' Допуск равновесия по My, Н*мм.
Private mSolverSecantMaxRestarts As Long ' Сколько раз Secant может перестроить приближенную матрицу.
Private mCapacityInitialLambda As Double ' Начальная величина множителя нагрузки lambda для поиска LoadMultiplier.
Private mCapacitySolutionStrategy As String ' Способ поиска предела: автоматически, изменением множителя нагрузки или по предельным деформациям.
Private mCapacitySearchMethod As String ' Численный метод уточнения предельной нагрузки в LoadMultiplier: Bisection, Brent или Secant.
Private mCapacityMaxLambda As Double ' Техническая верхняя граница lambda при поиске допустимой и предельной нагрузок.
Private mCapacityToleranceLambda As Double ' Точность lambda в одномерном поиске LoadMultiplier.
Private mCrackToleranceLambda As Double ' Точность lambda для поиска образования трещины.
Private mCapacityToleranceStrain As Double ' Точность попадания в предельную деформацию UltimateStrain.
Private mCapacityMaxRetries As Long ' Число повторных попыток решить НДС при том же lambda в случае несходимости.
Private mCapacityBaseLoadSteps As Long ' Число ступеней приложения нагрузки в первой попытке НДС при поиске несущей способности.
Private mCapacitySolverMaxIterations As Long ' Максимальное число итераций внутреннего решения НДС при поиске несущей способности.
Private mReport As CExecutionReport ' Необязательный человекочитаемый отчет выполнения; расчетная логика от него не зависит.
Private mReportUnits As CUnitSystem ' Единицы и знаки для вывода физических величин в отчёт выполнения.
Private mSettingsValid As Boolean ' False, если ApplySettings нашел ошибку Config и все LC должны получить InputErr.
Private mSettingsError As String ' Человекочитаемая причина ошибки чтения настроек Config.
Private mSettingLocations As Object ' Строковый снимок адресов настроек для пояснений результатов; живые ячейки в расчёте не используются.

Private mStabilityCode As String ' SP63 или SP35: выбранная нормативная ветвь фильтра устойчивости.
Private mStabilityElementLength As Double ' Геометрическая длина элемента l, мм.
Private mStabilityMu1 As Double ' Коэффициент расчетной длины для первой главной плоскости.
Private mStabilityMu2 As Double ' Коэффициент расчетной длины для второй главной плоскости.
Private mStabilitySystemType As String ' Determinate или Indeterminate: правило сложения основного и случайного эксцентриситета.
Private mStabilityZeroSign1 As Double ' Направление случайного эксцентриситета при M1=0: +1 или -1.
Private mStabilityZeroSign2 As Double ' Направление случайного эксцентриситета при M2=0: +1 или -1.
Private mStabilityPhiLMode As String ' Режим определения коэффициента длительности phi_l: Auto или PhiL2.
Private mStabilityAccidentalEccentricityMode As String ' Способ определения случайного эксцентриситета: по l, по mu*l или по заданному пользователем значению.
Private mStabilityAccidentalEccentricityPlanes As String ' OnlyMomentPlane или BothPlanes: в каких плоскостях учитывать случайный эксцентриситет.
Private mStabilityAccidentalEccentricityUser1 As Double ' Модуль пользовательского случайного эксцентриситета первой главной плоскости, мм.
Private mStabilityAccidentalEccentricityUser2 As Double ' Модуль пользовательского случайного эксцентриситета второй главной плоскости, мм.
Private mStabilitySP63Ks As Double ' Коэффициент ks для арматуры в D по СП 63.
Private mStabilitySP63DeltaEMin As Double ' Нижняя защита delta_e для численной устойчивости формулы.
Private mStabilitySP63DeltaEMax As Double ' Верхняя защита delta_e для численной устойчивости формулы.
Private mStabilitySP35PhiP As Double ' Коэффициент phi_p; для ненапрягаемой арматуры по умолчанию 1.
Private mStabilitySP35NOverNcrLimit As Double ' Ограничение N/Ncr для ветви СП 35.
Private mStabilitySP35Table721 As Variant ' Таблица 7.21 СП 35, прочитанная с Config.
Private mStabilitySP35TableInputError As String ' Причина неверной таблицы с актуальным адресом; применяется только в табличной ветви СП 35.
Private mStabilityLoadCount As Long ' Число строк постоянных/длительных нагрузок для устойчивости.
Private mStabilityLoadCombinationID() As String ' LC, к которому относится строка длительных нагрузок.
Private mStabilityLoadN() As Double ' N от постоянных или постоянных+длительных нагрузок, Н.
Private mStabilityLoadUserMx() As Double ' Пользовательский Mx длительной части в точке приложения нагрузки.
Private mStabilityLoadUserMy() As Double ' Пользовательский My длительной части в точке приложения нагрузки.
Private mStabilityLoadInputError() As String ' Ошибка прочтения длительной части только этого LC; пусто для корректной строки.

' Создаёт общие правила статусов, отчёт выполнения, фильтр малых моментов
' и начальные массивы сочетаний и длительных нагрузок. Начальные настройки
' нужны при программном использовании класса. При запуске из книги ApplySettings
' принимает Config; ошибочный явный ввод этими начальными настройками не заменяется.
Private Sub Class_Initialize()
    Set mStatusPolicy = New CResultStatusPolicy
    Set mReport = New CExecutionReport
    Set mMomentZeroFilter = New CMomentZeroFilter
    mSettingsValid = True
    mSettingsError = vbNullString
    mZeroMomentPerDepth = DEFAULT_ZERO_MOMENT_PER_DEPTH
    mMomentZeroFilter.Initialize mZeroMomentPerDepth
    EnsureCombinationCapacity INITIAL_COMBINATION_CAPACITY
    EnsureStabilityLoadCapacity INITIAL_COMBINATION_CAPACITY

    mCrackAllowable = 0.3
    mCrackStandardCode = "SP63"
    mCrackSP35GroupGap = 10#: mCrackSP35RowTolerance = 50#
    mCrackSP35NeighborRatioLimit = 0.2
    mCrackSP35DiameterMode = "Max": mCrackSP35RadiusMode = "6d": mCrackSP35RebarProfile = "Periodic"
    mCrackPsiMode = "AlwaysCalc"
    mCrackSigmaSCrcAveragingMode = "AllSelected"
    mCrackPhi1 = 1.4
    mCrackPhi2 = 0.5
    mCrackPhi3Mode = "Auto"
    mCrackPhi3 = 1#
    mCrackPsiS = 1#
    mCrackTensionZoneMode = "Effective"
    mCrackCoverDistanceMode = "GlobalExtreme"
    mCrackFormationSolutionStrategy = "Auto"
    mSolverMethod = "Newton"
    mSolverLoadSteps = 8
    mSolverMaxIterations = 80
    mSolverToleranceN = 5#
    mSolverToleranceMx = 5000#
    mSolverToleranceMy = 5000#
    mSolverSecantMaxRestarts = 2
    mCapacityInitialLambda = 1#
    mCapacitySolutionStrategy = "Auto"
    mCapacitySearchMethod = "Bisection"
    mCapacityMaxLambda = 1024#
    mCapacityToleranceLambda = 0.01
    mCrackToleranceLambda = 0.01
    mCapacityToleranceStrain = 0.00001
    mCapacityMaxRetries = 0
    mCapacityBaseLoadSteps = 1
    mCapacitySolverMaxIterations = 60
    mWorstCombinationCriterion = "StrengthCapacity"
    mStabilityCode = "SP35"
    mStabilityElementLength = 8000#
    mStabilityMu1 = 2#
    mStabilityMu2 = 2#
    mStabilitySystemType = "Determinate"
    mStabilityZeroSign1 = 1#
    mStabilityZeroSign2 = 1#
    mStabilityPhiLMode = "Auto"
    mStabilityAccidentalEccentricityMode = "AutoWithMuL"
    mStabilityAccidentalEccentricityPlanes = "OnlyMomentPlane"
    mStabilityAccidentalEccentricityUser1 = 10#
    mStabilityAccidentalEccentricityUser2 = 10#
    mStabilitySP63Ks = 0.7
    mStabilitySP63DeltaEMin = 0.15
    mStabilitySP63DeltaEMax = 1.5
    mStabilitySP35PhiP = 1#
    mStabilitySP35NOverNcrLimit = 0.7
End Sub

' Принимает готовую модель и материалы до добавления и расчёта сочетаний.
' Все сочетания используют одну CSectionModel без повторного построения геометрии.
Public Sub Initialize(ByVal section As CSectionModel, ByVal materialProvider As CMaterialModelProvider)
    If section Is Nothing Then Err.Raise vbObjectError + 3900, "CBatchSectionCalculator", "Модель сечения не передана."
    If section.ConcreteCount <= 0 Then Err.Raise vbObjectError + 3901, "CBatchSectionCalculator", "В модели сечения нет бетонных элементов."
    If materialProvider Is Nothing Then Err.Raise vbObjectError + 3902, "CBatchSectionCalculator", "Поставщик материалов не передан."
    Set mSection = section
    Set mMaterialProvider = materialProvider
    ClearTransformedReferenceCache
    mExecuting = False
    Set mCrackGeometryQuery = Nothing
    mCrackDomainBuildCount = 0
    mMechanicalPreparationBuildCount = 0
    ConfigureConcreteMomentDepths
End Sub

' Подключает профили, прочитанные из rngCalculationProfiles.
' Для каждой строки сочетаний выбирает профиль по её ProfileId.
Public Property Set ProfileCatalog(ByVal profiles As CCalculationProfileCatalog)
    Set mProfiles = profiles
End Property

' Подключает внешний отчёт выполнения и добавляет в него события расчёта.
' Открытие и сохранение файла выполняет вызывающий Excel-макрос.
Public Property Set ExecutionReport(ByVal report As CExecutionReport)
    Set mReport = report
    If mReport Is Nothing Then Set mReport = New CExecutionReport
    Set mReport.UnitSystem = mReportUnits
End Property

' Принимает численные настройки решения НДС и трещин из уже прочитанных данных Config.
' Через CUnitSystem переводит допуски сил, моментов и длин во внутренние единицы.
Public Sub ApplySettings(ByVal settings As CSystemSettingsReader, Optional ByVal units As CUnitSystem = Nothing)
    On Error GoTo InvalidSettings
    If settings Is Nothing Then Exit Sub
    mSettingsValid = True
    mSettingsError = vbNullString
    Dim settingIndex As Long, settingKey As String
    Set mSettingLocations = CreateObject("Scripting.Dictionary")
    mSettingLocations.CompareMode = vbTextCompare
    For settingIndex = 1 To settings.KeyCount
        settingKey = settings.KeyAt(settingIndex)
        mSettingLocations(settingKey) = settings.InputLocation(settingKey)
    Next settingIndex
    Set mReportUnits = units
    Set mReport.UnitSystem = units
    Dim conversionKey As String ' Настройка текущего пересчета единиц для точной диагностики переполнения.
    mCrackStandardCode = settings.GetRequiredChoice("SLS.Crack.Code", Array("SP63", "SP35")) ' Параметры трещин обязательны: пустая ячейка или отсутствующий ключ должны дать ошибку, а не изменить методику подстановкой начального значения.
    mCrackAllowable = settings.GetRequiredDouble("SLS.Crack.Allowable")
    mCrackFormationSolutionStrategy = settings.GetRequiredString("SLS.Crack.InitiationSolutionStrategy")
    Dim formationOptions As CCrackFormationCalculator
    Set formationOptions = New CCrackFormationCalculator
    mCrackFormationSolutionStrategy = formationOptions.NormalizeCrackFormationSolutionStrategy(mCrackFormationSolutionStrategy, settings) ' Владелец поиска проверяет стратегию до НДС и добавляет фактический адрес ошибочной настройки.
    mCrackToleranceLambda = settings.GetRequiredDouble("SLS.Crack.ToleranceLambda")
    formationOptions.LambdaTolerance = mCrackToleranceLambda
    formationOptions.ValidateSearchSettings settings
    mWorstCombinationCriterion = settings.GetRequiredChoice("General.WorstCombinationCriterion", _
        Array("StrengthStrain", "StrengthCapacity", "Cracks", "Stability"))
    If mCrackStandardCode = "SP35" Then
        mCrackSP35GroupGap = settings.GetRequiredDouble("SLS.Crack.SP35.GroupGapTolerance")
        mCrackSP35RowTolerance = settings.GetRequiredDouble("SLS.Crack.SP35.RowTolerance")
        mCrackSP35NeighborRatioLimit = settings.GetRequiredDouble("SLS.Crack.SP35.NeighborRatioLimit")
        mCrackSP35DiameterMode = settings.GetRequiredChoice("SLS.Crack.SP35.RadiusDiameterMode", Array("Max", "Min", "Average"))
        mCrackSP35RadiusMode = settings.GetRequiredChoice("SLS.Crack.SP35.InteractionRadiusMode", Array("3d", "5d", "6d"))
        mCrackSP35RebarProfile = settings.GetRequiredChoice("SLS.Crack.SP35.RebarProfile", Array("Periodic", "Smooth"))
    Else
        mCrackPsiMode = settings.GetRequiredString("SLS.Crack.PsiMode")
        mCrackSigmaSCrcAveragingMode = settings.GetRequiredString("SLS.Crack.SigmaSCrcAveragingMode")
        mCrackTensionZoneMode = settings.GetRequiredString("SLS.Crack.TensionZoneMode")
        mCrackCoverDistanceMode = settings.GetRequiredString("SLS.Crack.CoverDistanceMode")
        mCrackPhi1 = settings.GetRequiredDouble("SLS.Crack.Phi1")
        mCrackPhi2 = settings.GetRequiredDouble("SLS.Crack.Phi2")
        mCrackPhi3Mode = settings.GetRequiredString("SLS.Crack.Phi3Mode")
        mCrackPhi3 = settings.GetRequiredDouble("SLS.Crack.Phi3")
        mCrackPsiS = settings.GetRequiredDouble("SLS.Crack.PsiS")
    End If
    mZeroMomentPerDepth = settings.GetRequiredDouble("Calculation.ZeroMomentPerDepth")
    If settings.HasKey("Solver.Method") Then
        mSolverMethod = settings.GetRawString("Solver.Method", vbNullString)
    Else
        mSolverMethod = "Newton"
    End If
    mSolverLoadSteps = settings.GetLong("Solver.LoadSteps", mSolverLoadSteps, False)
    mSolverMaxIterations = settings.GetLong("Solver.MaxIterations", mSolverMaxIterations, False)
    mSolverToleranceN = settings.GetDouble("Solver.ToleranceN", mSolverToleranceN, False)
    mSolverToleranceMx = settings.GetDouble("Solver.ToleranceMx", mSolverToleranceMx, False)
    mSolverToleranceMy = settings.GetDouble("Solver.ToleranceMy", mSolverToleranceMy, False)
    mSolverSecantMaxRestarts = settings.GetLong("Solver.SecantMaxRestarts", mSolverSecantMaxRestarts, False)
    mCapacityInitialLambda = settings.GetDouble("Capacity.InitialLambda", mCapacityInitialLambda, False)
    mCapacitySolutionStrategy = CapacityTextSetting(settings, "Capacity.SolutionStrategy")
    mCapacitySearchMethod = CapacityTextSetting(settings, "Capacity.SearchMethod")
    mCapacityMaxLambda = settings.GetDouble("Capacity.MaxLambda", mCapacityMaxLambda, False)
    mCapacityToleranceLambda = settings.GetDouble("Capacity.ToleranceLambda", mCapacityToleranceLambda, False)
    mCapacityToleranceStrain = settings.GetDouble("Capacity.ToleranceStrain", mCapacityToleranceStrain, False)
    mCapacityMaxRetries = settings.GetLong("Capacity.MaxRetries", mCapacityMaxRetries, False)
    mCapacityBaseLoadSteps = settings.GetLong("Capacity.BaseLoadSteps", mCapacityBaseLoadSteps, False)
    mCapacitySolverMaxIterations = settings.GetLong("Capacity.SolverMaxIterations", mCapacitySolverMaxIterations, False)
    mStabilityCode = StabilityTextSetting(settings, "Stability.Code")
    mStabilityElementLength = StabilityNumericSetting(settings, "Stability.ElementLength")
    mStabilityMu1 = StabilityNumericSetting(settings, "Stability.Mu1")
    mStabilityMu2 = StabilityNumericSetting(settings, "Stability.Mu2")
    mStabilitySystemType = settings.GetRequiredString("Stability.SystemType")
    mStabilityZeroSign1 = StabilityNumericSetting(settings, "Stability.ZeroMomentEccentricitySign1")
    mStabilityZeroSign2 = StabilityNumericSetting(settings, "Stability.ZeroMomentEccentricitySign2")
    mStabilityPhiLMode = StabilityTextSetting(settings, "Stability.PhiLMode")
    mStabilityAccidentalEccentricityMode = StabilityTextSetting(settings, "Stability.AccidentalEccentricityMode")
    mStabilityAccidentalEccentricityPlanes = StabilityTextSetting(settings, "Stability.AccidentalEccentricityPlanes")
    mStabilityAccidentalEccentricityUser1 = StabilityNumericSetting(settings, "Stability.AccidentalEccentricityUser1")
    mStabilityAccidentalEccentricityUser2 = StabilityNumericSetting(settings, "Stability.AccidentalEccentricityUser2")
    mStabilitySP63Ks = StabilityNumericSetting(settings, "Stability.SP63.Ks")
    mStabilitySP63DeltaEMin = StabilityNumericSetting(settings, "Stability.SP63.DeltaEMin")
    mStabilitySP63DeltaEMax = StabilityNumericSetting(settings, "Stability.SP63.DeltaEMax")
    mStabilitySP35PhiP = StabilityNumericSetting(settings, "Stability.SP35.PhiP")
    mStabilitySP35NOverNcrLimit = StabilityNumericSetting(settings, "Stability.SP35.NOverNcrLimit")
    If Not units Is Nothing Then
        conversionKey = "SLS.Crack.Allowable"
        mCrackAllowable = units.InputLengthToInternal(mCrackAllowable)
        If mCrackStandardCode = "SP35" Then
            conversionKey = "SLS.Crack.SP35.GroupGapTolerance"
            mCrackSP35GroupGap = units.InputLengthToInternal(mCrackSP35GroupGap)
            conversionKey = "SLS.Crack.SP35.RowTolerance"
            mCrackSP35RowTolerance = units.InputLengthToInternal(mCrackSP35RowTolerance)
        End If
        conversionKey = "Stability.ElementLength"
        mStabilityElementLength = units.InputLengthToInternal(mStabilityElementLength)
        conversionKey = "Stability.AccidentalEccentricityUser1"
        mStabilityAccidentalEccentricityUser1 = units.InputLengthToInternal(mStabilityAccidentalEccentricityUser1)
        conversionKey = "Stability.AccidentalEccentricityUser2"
        mStabilityAccidentalEccentricityUser2 = units.InputLengthToInternal(mStabilityAccidentalEccentricityUser2)
        conversionKey = "Solver.ToleranceN" ' Отрицательный допуск остается ошибкой настройки и после перевода единиц.
        mSolverToleranceN = Sgn(mSolverToleranceN) * Abs(units.InputForceToInternal(mSolverToleranceN))
        conversionKey = "Solver.ToleranceMx"
        mSolverToleranceMx = Sgn(mSolverToleranceMx) * Abs(units.InputMomentMxToInternal(mSolverToleranceMx))
        conversionKey = "Solver.ToleranceMy"
        mSolverToleranceMy = Sgn(mSolverToleranceMy) * Abs(units.InputMomentMyToInternal(mSolverToleranceMy))
        conversionKey = "Calculation.ZeroMomentPerDepth"
        mZeroMomentPerDepth = Sgn(mZeroMomentPerDepth) * units.InputMomentPerLengthToInternal(mZeroMomentPerDepth) ' Переводим момент на единицу длины в Н*мм/мм, сохраняя отрицательный знак для проверки ошибочного ввода.
        conversionKey = vbNullString
    End If
    Dim widthOptions As CCrackWidthCalculator
    Set widthOptions = New CCrackWidthCalculator
    ConfigureCrackCalculator widthOptions
    widthOptions.ValidateSettings settings ' Проверка параметров выполняется до НДС и до вывода об отсутствии трещины; диапазоны не дублируются здесь.
    Dim capacityOptions As CCapacityCalculator
    Set capacityOptions = New CCapacityCalculator
    ConfigureCapacityCalculator capacityOptions
    capacityOptions.ValidateSettings settings ' Владелец расчёта проверяет диапазоны; прочитанные настройки дают фактический адрес ошибки Config.
    ' Расчёт устойчивости проверяет собственные значения и варианты настроек до запуска.
    ' Адрес ошибки берём из прочитанного Config после перевода единиц.
    Dim stabilityOptions As CStabilityCalculator
    Set stabilityOptions = New CStabilityCalculator
    stabilityOptions.ValidateSettings mStabilityCode, mStabilityElementLength, mStabilityMu1, mStabilityMu2, _
        mStabilitySystemType, mStabilityZeroSign1, mStabilityZeroSign2, mStabilityPhiLMode, _
        mStabilityAccidentalEccentricityMode, mStabilityAccidentalEccentricityPlanes, _
        mStabilityAccidentalEccentricityUser1, mStabilityAccidentalEccentricityUser2, _
        mStabilitySP63Ks, mStabilitySP63DeltaEMin, mStabilitySP63DeltaEMax, _
        mStabilitySP35PhiP, mStabilitySP35NOverNcrLimit, settings
    ConfigureMomentZeroFilter
    NormalizeAllCombinationMoments
    Exit Sub

InvalidSettings:
    mSettingsValid = False
    If Err.Number = 6 And Len(conversionKey) > 0 Then
        mSettingsError = settings.InputErrorMessage(conversionKey, _
            "Значение слишком велико для пересчета в выбранных INPUT-единицах.", _
            "Введите меньшее по модулю число и проверьте INPUT-единицы на листе Config.")
    ElseIf Err.Number = vbObjectError + 6010 Then
        mSettingsError = settings.InputErrorMessage("Calculation.ZeroMomentPerDepth", _
            Err.Description, "Введите ноль для отключения фильтра или положительное число в INPUT-единицах момент/длина.")
    Else
        mSettingsError = Err.Description
    End If
    ReportStep "Настройки расчета содержат ошибку: " & mSettingsError
End Sub

' Проверяет обязательный выбор способа расчёта несущей способности до решения НДС.
' Ошибка содержит адрес Config и допустимые варианты. Пустое значение
' не подменяется Auto или начальными настройками численного метода.
Private Function CapacityTextSetting(ByVal settings As CSystemSettingsReader, ByVal key As String) As String
    Dim value As String, allowed As String
    value = settings.GetRequiredString(key)
    If key = "Capacity.SolutionStrategy" Then
        allowed = "Auto, UltimateStrain или LoadMultiplier"
        Select Case LCase$(Trim$(value))
            Case "auto": CapacityTextSetting = "Auto"
            Case "ultimatestrain": CapacityTextSetting = "UltimateStrain"
            Case "loadmultiplier": CapacityTextSetting = "LoadMultiplier"
        End Select
    Else
        allowed = "Bisection, Brent или Secant"
        Select Case LCase$(Trim$(value))
            Case "bisection": CapacityTextSetting = "Bisection"
            Case "brent": CapacityTextSetting = "Brent"
            Case "secant": CapacityTextSetting = "Secant"
        End Select
    End If
    If Len(CapacityTextSetting) = 0 Then Err.Raise vbObjectError + 4313, "CBatchSectionCalculator", _
        settings.InputErrorMessage(key, "Неизвестный вариант: """ & value & """.", "Выберите " & allowed & ".")
End Function

' Читает обязательное число для устойчивости. Если настройка отсутствует,
' пуста или содержит текст, сообщает об ошибке ввода без подстановки значения по умолчанию.
Private Function StabilityNumericSetting(ByVal settings As CSystemSettingsReader, _
        ByVal key As String) As Double
    StabilityNumericSetting = settings.GetRequiredDouble(key)
End Function

' Читает обязательный текстовый переключатель устойчивости без подмены пустого
' значения. Допустимые варианты проверяются внутри CStabilityCalculator.
Private Function StabilityTextSetting(ByVal settings As CSystemSettingsReader, _
        ByVal key As String) As String
    StabilityTextSetting = settings.GetRequiredString(key)
End Function

' Очищает сочетания и их результаты перед чтением новой таблицы нагрузок.
' Статусы предыдущего запуска не должны попасть в новую сводку.
Public Sub ClearCombinations()
    mCount = 0
    ClearStabilityDurationLoads
    mGoverningCombinationIndex = 0
    mCrackGoverningCombinationIndex = 0
    mWorstCombinationIndex = 0
    mElapsedSeconds = 0#
    mSolverCallCount = 0
    mSolverIterationCount = 0
    mUltimateStrainStartCount = 0
    mUltimateStrainIterationCount = 0
    mDiagnosticLog = vbNullString
    mLoadReferenceX = 0#
    mLoadReferenceY = 0#
    mLoadReferenceOffsetX = 0#
    mLoadReferenceOffsetY = 0#
End Sub

' Очищает таблицу постоянных/длительных нагрузок для расчета устойчивости.
Public Sub ClearStabilityDurationLoads()
    mStabilityLoadCount = 0
    Dim i As Long
    For i = 1 To StabilityLoadCapacity()
        mStabilityLoadCombinationID(i) = vbNullString
        mStabilityLoadN(i) = 0#
        mStabilityLoadUserMx(i) = 0#
        mStabilityLoadUserMy(i) = 0#
        mStabilityLoadInputError(i) = vbNullString
    Next i
End Sub

' Добавляет строку постоянных или постоянных+длительных нагрузок для LC.
' Какие именно длительные нагрузки вводит пользователь, зависит от выбранного
' нормативного блока и описано на листе "Справка"; ядро получает уже готовые
' значения во внутренних единицах.
Public Sub AddStabilityDurationLoad(ByVal combinationID As String, ByVal nValue As Double, _
        ByVal mxValue As Double, ByVal myValue As Double)
    EnsureStabilityLoadCapacity mStabilityLoadCount + 1
    mStabilityLoadCount = mStabilityLoadCount + 1
    mStabilityLoadCombinationID(mStabilityLoadCount) = Trim$(combinationID)
    mStabilityLoadN(mStabilityLoadCount) = nValue
    mStabilityLoadUserMx(mStabilityLoadCount) = mxValue
    mStabilityLoadUserMy(mStabilityLoadCount) = myValue
    mStabilityLoadInputError(mStabilityLoadCount) = vbNullString
End Sub

' Сохраняет ошибку чтения длительной нагрузки, не заменяя её нулём.
' Ошибка блокирует только устойчивость данного сочетания; остальные
' запрошенные проверки продолжаются с исходными усилиями.
Public Sub AddInvalidStabilityDurationLoad(ByVal combinationID As String, ByVal reason As String)
    If Len(Trim$(reason)) = 0 Then Err.Raise vbObjectError + 4113, _
        "CBatchSectionCalculator", "Не передана причина ошибки дополнительной нагрузки устойчивости."
    AddStabilityDurationLoad combinationID, 0#, 0#, 0#
    mStabilityLoadInputError(mStabilityLoadCount) = reason
End Sub

' Принимает таблицу 7.21 СП 35 из Config и возможную ошибку с адресом ячейки.
' Расчёт устойчивости определяет, нужна ли таблица в выбранном режиме
' и можно ли использовать её данные.
Public Sub SetSP35Table721(ByVal tableData As Variant, Optional ByVal inputError As String = vbNullString)
    mStabilitySP35Table721 = tableData
    mStabilitySP35TableInputError = inputError
End Sub

' Добавляет корректно прочитанное сочетание нагрузок.
' Усилия уже должны быть во внутренних единицах и в INTERNAL-системе знаков.
Public Sub AddCombination(ByVal combinationID As String, ByVal nValue As Double, ByVal mxValue As Double, _
        ByVal myValue As Double, ByVal profileId As String, ByVal comment As String, _
        Optional ByVal loadPathText As String = vbNullString, Optional ByVal sourceDataOffset As Long = 0, _
        Optional ByVal sourceReference As String = vbNullString)
    AddCombinationInternal combinationID, nValue, mxValue, myValue, profileId, comment, True, vbNullString, loadPathText, sourceDataOffset, sourceReference
End Sub

' Сохраняет сочетание с ошибкой ввода для вывода причины в отчёте.
' НДС по нему не решается; в таблице ставится InputErr.
Public Sub AddInvalidCombination(ByVal combinationID As String, ByVal profileId As String, _
        ByVal comment As String, ByVal reason As String, Optional ByVal loadPathText As String = vbNullString, _
        Optional ByVal sourceDataOffset As Long = 0, Optional ByVal sourceReference As String = vbNullString)
    AddCombinationInternal combinationID, 0#, 0#, 0#, profileId, comment, False, reason, loadPathText, sourceDataOffset, sourceReference
End Sub

' Единая точка добавления строки сочетания. Через нее проходят и корректные
' LC, и строки с ошибкой ввода; это гарантирует одинаковую нумерацию,
' сохранение комментариев и нормализацию малых моментов только для валидных
' нагрузок.
Private Sub AddCombinationInternal(ByVal combinationID As String, ByVal nValue As Double, ByVal mxValue As Double, _
        ByVal myValue As Double, ByVal profileId As String, ByVal comment As String, _
        ByVal inputValid As Boolean, ByVal inputError As String, ByVal loadPathText As String, _
        ByVal sourceDataOffset As Long, ByVal sourceReference As String)
    EnsureCombinationCapacity mCount + 1
    mCount = mCount + 1
    If Len(Trim$(combinationID)) = 0 Then combinationID = "ROW" & CStr(mCount)
    mInputs(mCount).CombinationID = combinationID
    mInputs(mCount).SourceDataOffset = sourceDataOffset
    mInputs(mCount).SourceReference = sourceReference
    mInputs(mCount).N = nValue
    mInputs(mCount).Mx = mxValue
    mInputs(mCount).My = myValue
    mInputs(mCount).ProfileId = Trim$(profileId)
    mInputs(mCount).LoadPath = Trim$(loadPathText)
    mInputs(mCount).Comment = comment
    mInputs(mCount).InputValid = inputValid
    mInputs(mCount).InputError = inputError
    NormalizeCombinationMoments mCount
End Sub

' Расширяет хранилище LC без изменения уже прочитанных нагрузок и результатов.
Private Sub EnsureCombinationCapacity(ByVal requiredCount As Long)
    If requiredCount <= CombinationCapacity() Then Exit Sub

    Dim currentCapacity As Long
    Dim newCapacity As Long
    currentCapacity = CombinationCapacity()
    newCapacity = currentCapacity
    If newCapacity < INITIAL_COMBINATION_CAPACITY Then newCapacity = INITIAL_COMBINATION_CAPACITY
    Do While newCapacity < requiredCount
        newCapacity = newCapacity + COMBINATION_CAPACITY_STEP
    Loop

    If currentCapacity = 0 Then
        ReDim mInputs(1 To newCapacity)
        ReDim mWorkingLoads(1 To newCapacity)
        ReDim mWorkingPaths(1 To newCapacity)
        ReDim mResults(1 To newCapacity)
    Else
        ReDim Preserve mInputs(1 To newCapacity)
        ReDim Preserve mWorkingLoads(1 To newCapacity)
        ReDim Preserve mWorkingPaths(1 To newCapacity)
        ReDim Preserve mResults(1 To newCapacity)
    End If

    Dim i As Long
    For i = currentCapacity + 1 To newCapacity
        Set mResults(i) = New CCombinationResult
        mResults(i).Clear CAPACITY_PATH_NONE
    Next i
End Sub

' Возвращает текущую емкость массива LC; неинициализированный массив дает 0.
Private Function CombinationCapacity() As Long
    On Error GoTo EmptyStorage
    CombinationCapacity = UBound(mInputs)
    Exit Function
EmptyStorage:
End Function

' Расширяет таблицу длительных нагрузок устойчивости вместе с фактическим числом строк Config.
Private Sub EnsureStabilityLoadCapacity(ByVal requiredCount As Long)
    If requiredCount <= StabilityLoadCapacity() Then Exit Sub

    Dim currentCapacity As Long
    Dim newCapacity As Long
    currentCapacity = StabilityLoadCapacity()
    newCapacity = currentCapacity
    If newCapacity < INITIAL_COMBINATION_CAPACITY Then newCapacity = INITIAL_COMBINATION_CAPACITY
    Do While newCapacity < requiredCount
        newCapacity = newCapacity + COMBINATION_CAPACITY_STEP
    Loop

    If currentCapacity = 0 Then
        ReDim mStabilityLoadCombinationID(1 To newCapacity)
        ReDim mStabilityLoadN(1 To newCapacity)
        ReDim mStabilityLoadUserMx(1 To newCapacity)
        ReDim mStabilityLoadUserMy(1 To newCapacity)
        ReDim mStabilityLoadInputError(1 To newCapacity)
    Else
        ReDim Preserve mStabilityLoadCombinationID(1 To newCapacity)
        ReDim Preserve mStabilityLoadN(1 To newCapacity)
        ReDim Preserve mStabilityLoadUserMx(1 To newCapacity)
        ReDim Preserve mStabilityLoadUserMy(1 To newCapacity)
        ReDim Preserve mStabilityLoadInputError(1 To newCapacity)
    End If
End Sub

' Возвращает текущую емкость массива длительных нагрузок; неинициализированный массив дает 0.
Private Function StabilityLoadCapacity() As Long
    On Error GoTo EmptyStorage
    StabilityLoadCapacity = UBound(mStabilityLoadCombinationID)
    Exit Function
EmptyStorage:
End Function

' Переносит моменты из пользовательской точки приложения нагрузки
' к внутреннему началу координат CSectionModel для решения равновесия.
Public Sub ApplyLoadReference(ByVal referenceX As Double, ByVal referenceY As Double, _
        Optional ByVal baseX As Double = 0#, Optional ByVal baseY As Double = 0#)
    mLoadReferenceX = referenceX
    mLoadReferenceY = referenceY
    mLoadReferenceOffsetX = referenceX - baseX
    mLoadReferenceOffsetY = referenceY - baseY

    Dim i As Long
    For i = 1 To mCount
        If mInputs(i).InputValid Then NormalizeCombinationMoments i
    Next i

End Sub

' Обновляет бетонные габариты, по которым общий фильтр решает, считать ли
' Mx/My практически нулевыми. Для Mx нужен размер по Y, для My - по X.
Private Sub ConfigureConcreteMomentDepths()
    If mSection Is Nothing Then Exit Sub

    Dim props As CSectionPropertiesCalculator
    Set props = New CSectionPropertiesCalculator
    mConcreteDepthForMx = props.ProjectionDepth(mSection, 0#, 1#, False)
    mConcreteDepthForMy = props.ProjectionDepth(mSection, 1#, 0#, False)
End Sub

' Применяет пользовательский порог к общему фильтру. Отрицательное значение
' считается ошибкой настройки, потому что порог является модулем момента.
Private Sub ConfigureMomentZeroFilter()
    If mMomentZeroFilter Is Nothing Then Set mMomentZeroFilter = New CMomentZeroFilter
    mMomentZeroFilter.Initialize mZeroMomentPerDepth
End Sub

' Зануляет Mx, если он меньше инженерного порога для бетонного габарита по Y.
Private Function NormalizeMomentMx(ByVal momentValue As Double) As Double
    If mMomentZeroFilter Is Nothing Then
        NormalizeMomentMx = momentValue
    Else
        NormalizeMomentMx = mMomentZeroFilter.NormalizeMoment(momentValue, mConcreteDepthForMx)
    End If
End Function

' Зануляет My, если он меньше инженерного порога для бетонного габарита по X.
Private Function NormalizeMomentMy(ByVal momentValue As Double) As Double
    If mMomentZeroFilter Is Nothing Then
        NormalizeMomentMy = momentValue
    Else
        NormalizeMomentMy = mMomentZeroFilter.NormalizeMoment(momentValue, mConcreteDepthForMy)
    End If
End Function

' Заново строит рабочую нагрузку из неизменяемой строки LC. Фильтр и перенос
' применяются к исходным числам, а не к результатам предыдущего запуска.
Private Sub NormalizeCombinationMoments(ByVal index As Long)
    If index < 1 Or index > mCount Then Exit Sub
    Dim loadState As CSectionLoadState
    Set loadState = New CSectionLoadState
    If mSettingsValid And mInputs(index).InputValid Then
        loadState.Initialize mInputs(index).N, mInputs(index).Mx, mInputs(index).My, _
            mLoadReferenceX, mLoadReferenceY, mMomentZeroFilter, mConcreteDepthForMx, mConcreteDepthForMy
    Else
        loadState.Initialize mInputs(index).N, mInputs(index).Mx, mInputs(index).My, mLoadReferenceX, mLoadReferenceY
    End If
    Set mWorkingLoads(index) = loadState
    mWorkingPaths(index) = mInputs(index).LoadPath
End Sub

' Применяет текущую настройку фильтра к уже загруженным сочетаниям. Метод
' нужен, если ApplySettings вызывается после чтения строк LC.
Private Sub NormalizeAllCombinationMoments()
    Dim i As Long
    For i = 1 To mCount
        NormalizeCombinationMoments i
    Next i
End Sub

' Последовательно рассчитывает все сочетания и выбирает определяющее.
' Для каждого сочетания сначала применяется запрошенная устойчивость, затем
' несущая способность, НДС заданной нагрузки и проверки трещин.
' Состав проверок и материальные модели задаются расчётным профилем.
Public Sub Execute()
    On Error GoTo Failed

    ValidateReady
    mSection.BeginForceIntegrationRun mCount
    ClearTransformedReferenceCache ' Модель и материалы могли измениться между запусками. Ранее сохранённый расчётный центр в новом запуске использовать нельзя.
    ConfigureConcreteMomentDepths
    NormalizeAllCombinationMoments
    mExecuting = True
    mCrackDomainBuildCount = 0: mMechanicalPreparationBuildCount = 0
    Dim t0 As Double
    Dim stage As String
    Dim currentIndex As Long

    stage = "Initialize"
    t0 = Timer
    mGoverningCombinationIndex = 0
    mCrackGoverningCombinationIndex = 0
    mWorstCombinationIndex = 0
    mSolverCallCount = 0
    mSolverIterationCount = 0
    mUltimateStrainStartCount = 0
    mUltimateStrainIterationCount = 0
    mDiagnosticLog = vbNullString
    ResetSectionEquilibriumSolveCount
    ReportStep "Начат пакетный расчет. Количество сочетаний: " & CStr(mCount) & "."

    Dim i As Long
    For i = 1 To mCount ' Ошибка одного сочетания попадает в его статус и общую сводку, но не стирает результаты остальных строк.
        currentIndex = i
        stage = "ClearResultRow"
        ClearResultRow i
        stage = "RunCombination"
        RunCombination i
        If mInputs(i).InputValid Then
            stage = "SelectStrengthGoverning"
            If StrengthSafetyFactorForGoverning(i) > 0# Then
                If mGoverningCombinationIndex = 0 Then
                    mGoverningCombinationIndex = i
                ElseIf IsBetterGoverningCandidate(i, mGoverningCombinationIndex) Then
                    mGoverningCombinationIndex = i
                End If
            End If
            stage = "SelectCrackGoverning"
            If IsBetterCrackGoverningCandidate(i, mCrackGoverningCombinationIndex) Then
                mCrackGoverningCombinationIndex = i
            End If
        End If
    Next i
    stage = "SelectWorstCombination"
    SelectWorstCombination
    stage = "Finalize"
    mSolverCallCount = SectionEquilibriumSolveCount()
    mSolverIterationCount = SectionEquilibriumIterationCount()
    mUltimateStrainStartCount = modSolverWorkStats.UltimateStrainStartCount()
    mUltimateStrainIterationCount = modSolverWorkStats.UltimateStrainIterationCount()
    mElapsedSeconds = Timer - t0
    If mElapsedSeconds < 0# Then mElapsedSeconds = mElapsedSeconds + 86400#
    ReportStep "Пакетный расчет завершен. Время: " & FormatNumberInvariant(mElapsedSeconds) & " с."
    ReleaseRunPreparation
    Exit Sub

Failed:
    Dim failureNumber As Long, failureText As String
    failureNumber = Err.Number: failureText = Err.Description
    ReleaseRunPreparation
    Err.Raise failureNumber, "CBatchSectionCalculator.Execute", _
        stage & " LC=" & CStr(currentIndex) & ": " & failureText
End Sub

' Освобождает временную общую геометрию и при успехе, и при ошибке.
' Готовые результаты и хранилища НДС сочетаний сохраняются;
' для диагностики остаётся число выполненных подготовок.
Private Sub ReleaseRunPreparation()
    mExecuting = False
    If Not mCrackGeometryQuery Is Nothing Then mCrackDomainBuildCount = mCrackGeometryQuery.TestDomainBuildCount
    Set mCrackGeometryQuery = Nothing
    Dim i As Long
    For i = 1 To 2
        If Not mRunTransformedProperties(i) Is Nothing Then _
            mMechanicalPreparationBuildCount = mMechanicalPreparationBuildCount + mRunTransformedProperties(i).TestPreparationBuildCount
        Set mRunTransformedProperties(i) = Nothing
    Next i
    If Not mRunConcreteProperties Is Nothing Then _
        mMechanicalPreparationBuildCount = mMechanicalPreparationBuildCount + mRunConcreteProperties.TestPreparationBuildCount
    Set mRunConcreteProperties = Nothing
End Sub

' Рассчитывает одно сочетание нагрузок и собирает его итоговый статус.
' Несущая способность определяет запас, решение при заданной нагрузке — текущее НДС;
' проверки трещин используют собственные материальные модели выбранного профиля.
Private Sub RunCombination(ByVal index As Long)
    On Error GoTo Failed
    ReportSection "Сочетание " & CStr(index) & ": " & mInputs(index).CombinationID
    ReportValue "Комментарий", mInputs(index).Comment
    ReportValue "ID профиля", mInputs(index).ProfileId
    If Not mInputs(index).InputValid Then
        ReportStep "Решение НДС не запущено из-за ошибки исходных данных: " & mInputs(index).InputError
        MarkInvalid index, mInputs(index).InputError
        Exit Sub
    End If

    If Not mSettingsValid Then
        ReportStep "Решение НДС не запущено из-за ошибки настроек: " & mSettingsError
        MarkInvalid index, mSettingsError
        Exit Sub
    End If

    Dim pathDescriptor As CLoadPathDescriptor, normalizedPath As String
    Set pathDescriptor = New CLoadPathDescriptor
    normalizedPath = pathDescriptor.NormalizeKey(mWorkingPaths(index))
    If Len(normalizedPath) = 0 Then
        MarkInvalid index, "В сочетании " & mInputs(index).CombinationID & _
            " исправьте LoadPath: выберите Auto, lambda*Mx, lambda*My, lambda*Mxy, lambda*N или lambda*NMxy."
        Exit Sub
    End If
    mWorkingPaths(index) = normalizedPath

    If Len(Trim$(mInputs(index).ProfileId)) = 0 Then
        ReportStep "Сочетание не рассчитано: ProfileId не задан."
        MarkInvalid index, "ProfileId должен быть назначен сочетанию."
        Exit Sub
    End If

    Dim profile As CCalculationProfile
    If Not mProfiles.HasProfile(mInputs(index).ProfileId) Then
        MarkInvalid index, "Расчетный профиль " & mInputs(index).ProfileId & _
            " не найден. Выберите ProfileId из таблицы расчетных профилей на листе Config."
        Exit Sub
    End If
    Set profile = mProfiles.ProfileById(mInputs(index).ProfileId)
    mResults(index).SetProfileId profile.ProfileId
    ReportValue "Профиль", profile.ProfileId & " / " & profile.DisplayName
    ReportValue "N", mReport.FormatForce(mWorkingLoads(index).N)
    ReportValue "Mx расчетный", mReport.FormatMx(mWorkingLoads(index).InternalMx)
    ReportValue "My расчетный", mReport.FormatMy(mWorkingLoads(index).InternalMy)

    Dim profileError As String
    If Not profile.ValidateForCalculation(profileError) Then
        ReportStep "Сочетание не рассчитано: " & profileError
        MarkInvalid index, profileError
        Exit Sub
    End If

    ReportSection "Продольный изгиб и устойчивость"
    If profile.StabilityEnabled Then
        ReportStep "Запускается фильтр продольного изгиба/устойчивости по профилю."
        RunStabilityFilter index, profile.StabilityValueSet
    Else
        ReportStep "Устойчивость не запрошена профилем."
        mResults(index).ClearStabilityResult
        mResults(index).SetStabilityMeta NotRequestedMeta(rkStability)
    End If

    ReportSection "Расчет по модели прочности"
    If profile.StrengthCapacityEnabled Then
        ReportStep "Запускается проверка несущей способности по профилю."
        RunCapacity index, profile.StrengthSpec
    Else
        ReportStep "Проверка несущей способности не запрошена профилем."
        StoreNotCalculatedCapacity index
    End If

    If profile.StrengthDirectStateEnabled Then
        ReportStep "Запускается прямое НДС по модели прочности профиля."
        RunProfileState index, sstStrengthState, cpStrength, profile.StrengthSpec
    End If

    ReportSection "Расчет нормальных и продольных трещин"
    If profile.CrackWidthEnabled Then
        RunCrackWidth index, profile
    Else
        ReportStep "Расчет трещин не запрошен профилем."
    End If

FinalizeCombination:
    AppendDiagnostic index
    ReportCombinationSummary index
    Exit Sub

Failed:
    Dim failureNumber As Long, failureText As String, failureSource As String, failureMeta As CResultMeta
    failureNumber = Err.Number: failureText = Err.Description: failureSource = Err.Source
    ReportStep "Ошибка при расчете сочетания: " & failureText
    Set failureMeta = New CResultMeta
    failureMeta.SetResult rsInternalError, rcInternalError, rkGeneric, _
        "Расчет сочетания остановлен из-за ошибки выполнения: " & failureText, _
        failureSource & "; error=" & CStr(failureNumber)
    mResults(index).SetWorkflowMeta failureMeta
    AppendDiagnostic index
End Sub

' Применяет продольный изгиб к нагрузкам сочетания. CStabilityCalculator
' возвращает скорректированные N/Mx/My и поясняющие величины;
' дальнейшие проверки и общие статусы определяет пакетный расчёт.
Private Sub RunStabilityFilter(ByVal index As Long, ByVal valueSet As EMaterialValueSet)
    On Error GoTo Failed

    Dim durationInputError As String
    durationInputError = StabilityDurationInputError(mInputs(index).CombinationID)
    If Len(durationInputError) > 0 Then
        mResults(index).SetStabilityMeta InvalidInputMeta(rkStability, durationInputError), durationInputError
        ReportStep "Устойчивость не рассчитана: " & durationInputError
        Exit Sub
    End If

    Dim referenceX As Double
    Dim referenceY As Double
    TransformedReferencePoint valueSet, referenceX, referenceY

    Dim sustainedN As Double
    Dim sustainedMx As Double
    Dim sustainedMy As Double
    StabilityDurationLoadForCombination mInputs(index).CombinationID, referenceX, referenceY, _
        sustainedN, sustainedMx, sustainedMy

    Dim stability As CStabilityCalculator
    Set stability = New CStabilityCalculator
    Dim slot As Long
    slot = TransformedReferenceSlot(valueSet)
    If mRunTransformedProperties(slot) Is Nothing Then Set mRunTransformedProperties(slot) = New CSectionPropertiesCalculator
    If mRunConcreteProperties Is Nothing Then Set mRunConcreteProperties = New CSectionPropertiesCalculator
    stability.UseMechanicalPreparation mRunTransformedProperties(slot), mRunConcreteProperties
    Dim loadState As CSectionLoadState
    Set loadState = LoadStateForCombination(index)
    Dim referenceMx As Double
    Dim referenceMy As Double
    referenceMx = loadState.MxAboutPoint(referenceY)
    referenceMy = loadState.MyAboutPoint(referenceX)

    stability.Calculate mSection, mMaterialProvider, valueSet, mStabilityCode, _
        loadState.N, referenceMx, referenceMy, _
        sustainedN, sustainedMx, sustainedMy, _
        mStabilityElementLength, mStabilityMu1, mStabilityMu2, _
        mStabilitySystemType, mStabilityZeroSign1, mStabilityZeroSign2, _
        mStabilityPhiLMode, mStabilityAccidentalEccentricityMode, _
        mStabilityAccidentalEccentricityPlanes, _
        mStabilityAccidentalEccentricityUser1, mStabilityAccidentalEccentricityUser2, _
        mStabilitySP63Ks, mStabilitySP63DeltaEMin, _
        mStabilitySP63DeltaEMax, mStabilitySP35PhiP, _
        mStabilitySP35NOverNcrLimit, mStabilitySP35Table721, mMomentZeroFilter, mStabilitySP35TableInputError

    Dim stabilityMeta As CResultMeta
    Set stabilityMeta = stability.ResultMeta
    mResults(index).StoreStabilityResult stabilityMeta, stability, referenceX, referenceY
    If stability.InputValid And stability.Applicable And stability.CheckPassed Then _
        ApplyStabilityDesignLoads index, stability, referenceX, referenceY

    ReportValue "Статус устойчивости", mResults(index).StabilityResult.Status
    ReportMetaComment "Комментарий результата устойчивости", mResults(index).StabilityMeta
    ReportValue "Нормативный блок", mReport.StabilityCodeDisplayName(stability.Code)
    ReportValue "Материалы для устойчивости", MaterialValueSetToText(valueSet)
    ReportValue "Расчетная ветвь", mReport.StabilityBranchDisplayName(stability.Branch)
    ReportValue "Контрольная сила устойчивости", mReport.FormatCompressionForce(stability.CriticalForce)
    ReportValue "Минимальный запас устойчивости", FormatNumberInvariant(stability.ReserveFactor)
    ReportValue "Итоговые N/Mx/My после фильтра", _
        "N=" & mReport.FormatForce(mWorkingLoads(index).N) & "; " & _
        "Mx=" & mReport.FormatMx(mWorkingLoads(index).InternalMx) & "; " & _
        "My=" & mReport.FormatMy(mWorkingLoads(index).InternalMy)
    If Not mReport Is Nothing Then mReport.ReportStabilityDetails mResults(index).StabilityResult
    Exit Sub

Failed:
    mResults(index).SetStabilityMeta InvalidInputMeta(rkStability, Err.Description), Err.Description
    ReportStep "Устойчивость не рассчитана: " & Err.Description
End Sub

' Возвращает ошибку первой сопоставленной строки длительных нагрузок.
' Порядок совпадает с выбором самих усилий: последующая строка с тем же ID
' не должна менять уже выбранную строку, а прошлый запуск сбрасывается Clear.
Private Function StabilityDurationInputError(ByVal combinationID As String) As String
    Dim i As Long
    For i = 1 To mStabilityLoadCount
        If StrComp(mStabilityLoadCombinationID(i), Trim$(combinationID), vbTextCompare) = 0 Then
            StabilityDurationInputError = mStabilityLoadInputError(i)
            Exit Function
        End If
    Next i
End Function

' Берёт постоянную/длительную часть первой строки с нужным номером сочетания.
' Читатель таблицы заранее регистрирует ошибку, если соответствующей строки нет.
Private Sub StabilityDurationLoadForCombination(ByVal combinationID As String, _
        ByVal referenceX As Double, ByVal referenceY As Double, _
        ByRef nValue As Double, ByRef mxValue As Double, ByRef myValue As Double)
    Dim i As Long
    For i = 1 To mStabilityLoadCount
        If StrComp(mStabilityLoadCombinationID(i), Trim$(combinationID), vbTextCompare) = 0 Then
            Dim loadState As CSectionLoadState
            Set loadState = New CSectionLoadState
            loadState.Initialize mStabilityLoadN(i), mStabilityLoadUserMx(i), mStabilityLoadUserMy(i), _
                mLoadReferenceX, mLoadReferenceY, mMomentZeroFilter, mConcreteDepthForMx, mConcreteDepthForMy
            nValue = loadState.N
            mxValue = loadState.MxAboutPoint(referenceY)
            myValue = loadState.MyAboutPoint(referenceX)
            Exit Sub
        End If
    Next i
End Sub

' Заменяет только рабочую нагрузку расчетными усилиями устойчивости. Сам
' CStabilityCalculator работает с моментами относительно центра тяжести
' приведенного сечения, как требуют СП 35/63. Перед НДМ моменты снова
' переносятся к внутренней системе координат элементов.
Private Sub ApplyStabilityDesignLoads(ByVal index As Long, ByVal stability As CStabilityCalculator, _
        ByVal referenceX As Double, ByVal referenceY As Double)
    Dim loadState As CSectionLoadState
    Set loadState = New CSectionLoadState
    loadState.InitializeFromPointMoments stability.DesignN, stability.DesignMx, stability.DesignMy, _
        referenceX, referenceY, mLoadReferenceX, mLoadReferenceY, _
        mMomentZeroFilter, mConcreteDepthForMx, mConcreteDepthForMy

    Set mWorkingLoads(index) = loadState
    Dim finalDesignMx As Double
    Dim finalDesignMy As Double
    finalDesignMx = loadState.MxAboutPoint(referenceY) ' В сводку идут отфильтрованные усилия, используемые последующими проверками; момент переносится к выбранному центру.
    finalDesignMy = loadState.MyAboutPoint(referenceX)
    mResults(index).UpdateStabilityDesignLoads loadState.N, finalDesignMx, finalDesignMy, _
        stability.PrincipalAngleRad
End Sub

' Возвращает объект нагрузки LC: он единолично переносит моменты между
' пользовательской точкой, произвольным центром и внутренним началом модели.
Private Function LoadStateForCombination(ByVal index As Long) As CSectionLoadState
    Set LoadStateForCombination = mWorkingLoads(index)
End Function

' Возвращает центр приведенного сечения для проверки относительно выбранного
' набора характеристик. Stability и центральная ветка трещин используют один
' и тот же расчетный смысл: моменты проверяются относительно приведенного
' сечения бетон + арматура, а не только бетонного центра.
Private Sub TransformedReferencePoint(ByVal valueSet As EMaterialValueSet, _
        ByRef referenceX As Double, ByRef referenceY As Double)
    Dim slot As Long
    slot = TransformedReferenceSlot(valueSet)
    If mTransformedReferenceValid(slot) Then
        referenceX = mTransformedReferenceX(slot)
        referenceY = mTransformedReferenceY(slot)
        Exit Sub
    End If

    Dim props As CSectionPropertiesCalculator
    If mExecuting Then
        If mRunTransformedProperties(slot) Is Nothing Then Set mRunTransformedProperties(slot) = New CSectionPropertiesCalculator
        Set props = mRunTransformedProperties(slot)
    Else
        Set props = New CSectionPropertiesCalculator ' При отдельном справочном чтении сохраняем вычисленные координаты центра. Временный объект подготовки в завершённый запуск не возвращаем.
    End If
    props.EnsureTransformed mSection, _
        mMaterialProvider.ConcreteElasticModulusFromValueSet(valueSet), _
        mMaterialProvider.SteelElasticModulusFromValueSet(valueSet)
    referenceX = props.CentroidX
    referenceY = props.CentroidY
    mTransformedReferenceX(slot) = referenceX
    mTransformedReferenceY(slot) = referenceY
    mTransformedReferenceValid(slot) = True
End Sub

' Удаляет сохранённые центры приведённого сечения при смене модели или материалов.
' Эти координаты зависят от геометрии CSectionModel и выбранных модулей Es/Eb.
Private Sub ClearTransformedReferenceCache()
    Dim i As Long
    For i = 1 To 2
        mTransformedReferenceValid(i) = False
        mTransformedReferenceX(i) = 0#
        mTransformedReferenceY(i) = 0#
        Set mRunTransformedProperties(i) = Nothing
    Next i
    Set mRunConcreteProperties = Nothing
End Sub

' Переводит набор характеристик материалов в маленький индекс кэша.
Private Function TransformedReferenceSlot(ByVal valueSet As EMaterialValueSet) As Long
    If valueSet = mvsSLS Then
        TransformedReferenceSlot = 2
    Else
        TransformedReferenceSlot = 1
    End If
End Function

' Возвращает центр приведенного сечения для распознавания истинного
' центрального растяжения в расчете трещин. Если задана модель
' CrackInitiation, берем ее ValueSet, потому что Ncrc считается через Ared
' и Rbt этой модели; иначе используем модель CrackedState, с которой
' работает текущий расчет трещин.
Private Sub CrackCentralReferencePoint(ByVal profile As CCalculationProfile, _
        ByRef referenceX As Double, ByRef referenceY As Double)
    Dim valueSet As EMaterialValueSet
    valueSet = profile.CrackedStateSpec.ValueSet
    If Not profile.CrackInitiationSpec Is Nothing Then
        If profile.CrackInitiationSpec.IsComplete Then valueSet = profile.CrackInitiationSpec.ValueSet
    End If
    TransformedReferencePoint valueSet, referenceX, referenceY
End Sub

' Возвращает расчетный центр для табличного вывода трещин по конкретному LC.
' Это тот же центр приведенного сечения, относительно которого crack-калькулятор
' распознает центральное растяжение и ищет точку образования нормальной
' трещины.
Private Sub CrackReferencePointForCombination(ByVal index As Long, _
        ByRef referenceX As Double, ByRef referenceY As Double)
    Dim profile As CCalculationProfile
    Set profile = mProfiles.ProfileById(mInputs(index).ProfileId)
    CrackCentralReferencePoint profile, referenceX, referenceY
End Sub

' Возвращает центр приведенного сечения для расчетов по модели прочности.
' Относительно этой точки таблица прочности показывает N, Mx, My
' в группе "относительно центральных осей приведенного сечения".
Private Sub StrengthReferencePointForCombination(ByVal index As Long, _
        ByRef referenceX As Double, ByRef referenceY As Double)
    Dim profile As CCalculationProfile
    If Not mProfiles.HasProfile(mInputs(index).ProfileId) Then Exit Sub
    Set profile = mProfiles.ProfileById(mInputs(index).ProfileId)
    TransformedReferencePoint profile.StrengthSpec.ValueSet, referenceX, referenceY
End Sub

' Запускает поиск несущей способности для LC.
' Пользовательский столбец LoadPath задает, какие компоненты нагрузки
' умножаются на lambda. Все варианты дальше сводятся к одной математике:
' Target = Offset + lambda*Base для N, Mx и My.
Private Sub RunCapacity(ByVal index As Long, ByVal strengthSpec As CMaterialModelSpec)
    Dim loadPath As CLoadPathDescriptor
    Dim capacityCalculator As CCapacityCalculator
    Dim loadState As CSectionLoadState
    Dim reportText As String

    Set capacityCalculator = New CCapacityCalculator
    ConfigureCapacityCalculator capacityCalculator
    Set loadState = LoadStateForCombination(index)
    Set loadPath = capacityCalculator.BuildLoadPath(mWorkingPaths(index), loadState)
    ReportValue "Путь нагрузки для несущей", IIf(loadPath.HasPath, loadPath.DisplayName, mWorkingPaths(index))

    ReportValue "Стратегия поиска несущей", capacityCalculator.SolutionStrategy

    Dim strengthReferenceX As Double
    Dim strengthReferenceY As Double
    StrengthReferencePointForCombination index, strengthReferenceX, strengthReferenceY

    Dim capacityResult As CCapacityResult
    Set capacityResult = capacityCalculator.Calculate(mSection, _
        mMaterialProvider.ConcreteMaterialForEquilibriumFromSpec(strengthSpec), _
        mMaterialProvider.SteelMaterialForEquilibriumFromSpec(strengthSpec), strengthSpec, _
        loadPath, loadState, _
        strengthReferenceX, strengthReferenceY, mLoadReferenceX, mLoadReferenceY, mMomentZeroFilter, _
        mConcreteDepthForMx, mConcreteDepthForMy)

    mResults(index).StoreCapacityResult capacityResult
    If Not capacityResult.SearchExecuted Then
        ReportStep capacityResult.EarlyStopReport(loadPath)
        Exit Sub
    End If

    ReportStep capacityCalculator.SearchPlanReport(loadPath, capacityResult.PureAxialForMethod)
    reportText = capacityResult.MethodChoiceReport(capacityCalculator.SolutionStrategy, loadPath)
    If Len(reportText) > 0 Then ReportStep reportText

    ReportValue "Статус несущей способности", mResults(index).StrengthResult.Capacity.Status
    ReportMetaComment "Комментарий результата несущей", mResults(index).CapacityMeta
    ReportValue "Фактический метод несущей", mResults(index).StrengthResult.Capacity.SolutionMethod
    ReportValue "lambda несущей", FormatNumberInvariant(mResults(index).StrengthResult.Capacity.LambdaCapacity)
    ReportValue "Фактический путь несущей", mResults(index).CapacityPathDisplayName
    ReportBlock "Диагностика поиска несущей способности", capacityResult.DiagnosticLog
    reportText = capacityResult.LastSolverDiagnosticLog
    If Len(reportText) > 0 Then _
        ReportBlock "Итерации внутреннего CSectionSolver для последней capacity-точки", reportText
End Sub

' Решает НДС указанного типа по модели материалов выбранного профиля.
' Начальные приближения, повторные попытки и численное продолжение диаграмм
' обрабатывает общий CStateSolutionRunner.
Private Function RunProfileState(ByVal index As Long, ByVal stateType As ESectionStateType, _
        ByVal materialRole As ECalculationPurpose, ByVal spec As CMaterialModelSpec, _
        Optional ByVal useDiagramExtension As Boolean = True, _
        Optional ByVal mergeIntoDirectStateStatus As Boolean = True) As CStateSolutionRunner
    mResults(index).SetDiagramExtensionEnabled mMaterialProvider.DiagramExtensionEnabled
    Dim effectiveDiagramExtension As Boolean
    effectiveDiagramExtension = (useDiagramExtension And mResults(index).DiagramExtensionEnabled)

    Dim request As CStateRequest
    Set request = New CStateRequest
    request.Initialize stateType, materialRole, spec, mWorkingLoads(index).N, mWorkingLoads(index).InternalMx, mWorkingLoads(index).InternalMy, _
        effectiveDiagramExtension, effectiveDiagramExtension

    Dim provider As CStateProvider
    Set provider = CreateStateProvider(index)

    Dim stateResult As CSectionStateResult
    Set stateResult = provider.GetOrSolve(request)

    Dim runner As CStateSolutionRunner
    Set runner = provider.LastRunner

    If stateResult Is Nothing Then
        If mergeIntoDirectStateStatus Then MergeDirectStateMeta index, _
            MissingStateMeta(runner, "Прямое НДС не получено: решатель не вернул расчетное состояние.")
        ReportStep "Прямое НДС " & SectionStateDisplayName(stateType) & " не вернуло расчетное состояние."
        Set RunProfileState = runner
        Exit Function
    End If

    Dim service As CSectionSolver
    If Not runner Is Nothing Then Set service = runner.ResultSolver
    If stateType = sstStrengthState Then mResults(index).StoreSectionStateResult stateResult, _
        mMaterialProvider.ConcreteMaterialFromSpec(spec), mMaterialProvider.SteelMaterialFromSpec(spec)

    ReportValue "НДС " & SectionStateDisplayName(stateType) & " сошлось", IIf(stateResult.Converged, "да", "нет")
    If provider.LastStateWasReused Then
        ReportStep "Прямое НДС " & SectionStateDisplayName(stateType) & " взято из сохранённого результата без повторного решения равновесия."
    ElseIf Not runner Is Nothing Then
        ReportValue "Итераций НДС", CStr(runner.Iterations)
        ReportBlock "Сценарий прямого StateSolution", runner.DiagnosticLog
        If Not service Is Nothing Then ReportBlock "Итерации прямого CSectionSolver", service.DiagnosticLog
    End If

    If Not stateResult.Converged Then
        If mergeIntoDirectStateStatus Then MergeDirectStateMeta index, stateResult.ResultMeta
        ReportStep "Прямое НДС " & SectionStateDisplayName(stateType) & " не сошлось: " & stateResult.StopReason
        Set RunProfileState = runner
        Exit Function
    End If

    Dim stateStatus As String
    stateStatus = mStatusPolicy.ExternalStatus(stateResult.ResultMeta)
    If stateResult.ExtensionUsed Or Not stateResult.WithinPhysicalRange Then
        If mergeIntoDirectStateStatus Then MergeDirectStateMeta index, stateResult.ResultMeta
        ReportStep "Прямое НДС " & SectionStateDisplayName(stateType) & " найдено, но вышло за физический диапазон диаграммы материала."
    Else
        If mergeIntoDirectStateStatus Then MergeDirectStateMeta index, stateResult.ResultMeta
    End If
    ReportValue "Статус прямого НДС", stateStatus
    ReportMetaComment "Комментарий результата НДС", stateResult.ResultMeta
    ReportValue "Использовано продолжение диаграммы", IIf(stateResult.ExtensionUsed, "да", "нет")
    Set RunProfileState = runner
End Function

' Выполняет запрошенный профилем расчёт трещин. Сначала проверяет образование
' нормальной трещины и сохраняет пороговую нагрузку, если она найдена.
' Текущее НДС с трещинами по полной нагрузке сочетания рассчитывает всегда:
' оно нужно для продольных трещин и сохранения Results. Ширину определяет
' при наличии трещины либо пригодного текущего НДС с psi_s = 1, если порог
' не найден. Отсутствие текущего НДС делает обе проверки недоступными.
Private Sub RunCrackWidth(ByVal index As Long, ByVal profile As CCalculationProfile)
    Dim crack As CCrackWidthCalculator
    Set crack = New CCrackWidthCalculator
    ConfigureCrackCalculator crack
    ReportStep "Запускается проверка образования нормальной трещины."
    ReportValue "Путь образования трещины", mWorkingPaths(index)
    ReportValue "Стратегия поиска образования трещины", mCrackFormationSolutionStrategy
    If Not mReport Is Nothing Then mReport.ReportCrackFormationSearchStart mWorkingPaths(index)

    Dim crackReferenceX As Double
    Dim crackReferenceY As Double
    CrackCentralReferencePoint profile, crackReferenceX, crackReferenceY

    Dim loadState As CSectionLoadState
    Set loadState = LoadStateForCombination(index)

    ' Центральная ветка трещин проверяется относительно центра тяжести
    ' приведенного сечения той материальной модели, по которой определяется
    ' Ncrc. Load.ReferenceOffset и сдвиг от бетонного центра учтены внутри
    ' CSectionLoadState, поэтому здесь нет ручной формулы N*dx/dy.
    Dim formationCalculator As CCrackFormationCalculator
    Set formationCalculator = New CCrackFormationCalculator
    ConfigureCrackFormationCalculator formationCalculator, index
    Set formationCalculator.StateProvider = CreateStateProvider(index)
    Dim formationResult As CCrackFormationResult
    Set formationResult = formationCalculator.CheckFormation(mSection, _
        mMaterialProvider, profile.CrackedStateSpec, profile.CrackInitiationSpec, _
        loadState.N, loadState.InternalMx, loadState.InternalMy, loadState, _
        crackReferenceX, crackReferenceY)
    crack.InitializeForFormation formationResult
    If Not crack.Converged Then AppendCrackDiagnostic index, crack
    If Not formationResult.PreCrackState Is Nothing Then mResults(index).AddState formationResult.PreCrackState
    If Not formationResult.PostCrackState Is Nothing Then mResults(index).AddState formationResult.PostCrackState
    ReportValue "Фактический путь образования трещины", formationResult.FormationMethod
    ReportValue "lambda образования трещины", FormatNumberInvariant(formationResult.LambdaCrc)
    ReportValue "Ncrc", mReport.FormatForce(formationResult.FormationNcrc)
    ReportValue "Mxy,crc", mReport.FormatMomentMagnitude(formationResult.Mcrc)
    ReportValue "Причина остановки поиска трещины", crack.StopReason
    If Not mReport Is Nothing Then mReport.ReportCrackFormationChoice formationResult
    ReportBlock "Диагностика образования трещины", formationResult.DiagnosticLog

    ReportStep "Запускается расчёт НДС с трещинами от полной нагрузки сочетания по выбранному профилю."
    Dim runner As CStateSolutionRunner
    Set runner = RunProfileState(index, sstCrackedState, cpCrackedNDS, profile.CrackedStateSpec, True, False)
    Dim currentCrackedState As CSectionStateResult
    Set currentCrackedState = mResults(index).StateRepository.FindState(sstCrackedState)

    Dim currentStateMeta As CResultMeta
    Dim crackWidthMetaOverride As CResultMeta
    Dim longitudinalResult As CLongitudinalCrackResult
    Dim blockReason As String
    Dim longitudinalCalculator As CLongitudinalCrackCalculator
    Set longitudinalCalculator = New CLongitudinalCrackCalculator

    If Not CurrentCrackedStateReady(currentCrackedState, runner, currentStateMeta, blockReason) Then
        Set crackWidthMetaOverride = New CResultMeta
        crackWidthMetaOverride.SetBlockedByDependency rkCrackWidth, "CurrentCrackedState", blockReason
        Set longitudinalResult = longitudinalCalculator.BlockedByDependency( _
            "CurrentCrackedState", blockReason, mMaterialProvider.ConcreteParameters.RbMc2)
        ReportStep "Расчет ширины и продольных трещин не продолжается: " & blockReason
        GoTo ReportCrackResult
    End If

    Set longitudinalResult = longitudinalCalculator.Calculate(currentCrackedState, _
        mMaterialProvider.ConcreteParameters.RbMc2)
    If Not mReport Is Nothing Then mReport.ReportLongitudinalCrackResult longitudinalResult

    If Not crack.CrackFormed Then
        ReportStep "Ширина раскрытия не считается: нормальная трещина не образуется при текущей нагрузке."
        GoTo ReportCrackResult
    End If

    ReportStep "Запускается расчет ширины раскрытия нормальных трещин."
    If mCrackGeometryQuery Is Nothing Then Set mCrackGeometryQuery = New CSectionGeometryQuery
    crack.Calculate currentCrackedState, mSection, mMaterialProvider, profile.CrackedStateSpec, _
        formationResult, loadState, mCrackGeometryQuery
    If Not crack.Converged Then AppendCrackDiagnostic index, crack

ReportCrackResult:
    StoreCrackAggregateSnapshot index, crack, formationResult, currentStateMeta, _
        crackWidthMetaOverride, longitudinalResult
    ReportValue "Статус нормальных трещин", mResults(index).NormalCrackStatus
    ReportMetaComment "Комментарий образования трещины", mResults(index).CrackFormationMeta
    ReportMetaComment "Комментарий равновесия CrackedState", mResults(index).CrackCurrentStateMeta
    ReportMetaComment "Комментарий ширины нормальных трещин", mResults(index).CrackWidthMeta
    ReportMetaComment "Комментарий продольных трещин", mResults(index).LongitudinalCrackMeta
    ReportValue "Фактический путь образования трещины", mResults(index).CrackResult.Formation.FormationMethod
    ReportValue "lambda образования трещины", FormatNumberInvariant(mResults(index).CrackResult.Formation.LambdaCrc)
    If mResults(index).CrackResult.Formation.HasLimitPoint Then
        ReportValue "Трещина образована", IIf(mResults(index).CrackResult.Formation.CrackFormed, "да", "нет")
    Else
        ReportValue "Трещина образована", "пороговая точка не определена; см. результат поиска"
    End If
    ReportValue "a_crc", mReport.FormatLength(mResults(index).CrackResult.Width.CrackWidth)
    ReportValue "psi_s", FormatNumberInvariant(mResults(index).CrackResult.Width.PsiS)
    ReportValue "As", mReport.FormatArea(mResults(index).CrackResult.Width.AsTension)
    ReportValue "Abt", mReport.FormatArea(mResults(index).CrackResult.Width.Abt)
    ReportValue "ds_eq", mReport.FormatLength(mResults(index).CrackResult.Width.DsEquivalent)
    ReportValue "ls", mReport.FormatLength(mResults(index).CrackResult.Width.CrackSpacing)
    ReportBlock "Диагностика расчета трещин", crack.DiagnosticLog
End Sub

' После доступных этапов собирает один окончательный результат трещин.
' Использует фактические результаты этапов; числа и статусы прежнего запуска
' как запасные данные не подставляются.
Private Sub StoreCrackAggregateSnapshot(ByVal index As Long, ByVal crack As CCrackWidthCalculator, _
        ByVal formationResult As CCrackFormationResult, _
        ByVal currentStateMeta As CResultMeta, ByVal widthMeta As CResultMeta, _
        ByVal longitudinalResult As CLongitudinalCrackResult)
    Dim widthResult As CCrackWidthResult
    Set widthResult = New CCrackWidthResult
    If widthMeta Is Nothing Then
        widthResult.InitializeFromCalculator crack, crack.ResultMeta, mInputs(index).CombinationID
    Else
        widthResult.InitializeFromCalculator crack, widthMeta, mInputs(index).CombinationID
    End If

    Dim crackResult As CCrackResult
    Set crackResult = New CCrackResult
    crackResult.Initialize formationResult, currentStateMeta, widthResult, longitudinalResult
    mResults(index).StoreCrackAggregateResult crackResult
End Sub

' Проверяет пригодность текущего НДС с трещинами для нормальных и продольных трещин.
' Если НДС непригодно, возвращает фактическую причину этого этапа,
' а зависимые формулы отдельно отмечаются как недоступные.
Private Function CurrentCrackedStateReady(ByVal stateResult As CSectionStateResult, _
        ByVal runner As CStateSolutionRunner, ByRef currentStateMeta As CResultMeta, _
        ByRef blockReason As String) As Boolean
    Set currentStateMeta = CurrentCrackedStateMeta(stateResult, runner, blockReason)
    If stateResult Is Nothing Then Exit Function
    If Not stateResult.Converged Then Exit Function
    If stateResult.ExtensionUsed Or Not stateResult.WithinPhysicalRange Then Exit Function
    CurrentCrackedStateReady = True
End Function

' Возвращает статус текущего НДС с трещинами, даже если CStateProvider
' не создал объект состояния. Так сохраняется исходная причина неудачи
' без подмены её ошибкой формулы ширины трещины.
Private Function CurrentCrackedStateMeta(ByVal stateResult As CSectionStateResult, _
        ByVal runner As CStateSolutionRunner, ByRef blockReason As String) As CResultMeta
    If Not stateResult Is Nothing Then
        Set CurrentCrackedStateMeta = stateResult.ResultMeta
        blockReason = CurrentCrackedStateMeta.ResultComment
        If Len(Trim$(blockReason)) = 0 Then
            If Not stateResult.Converged Then blockReason = "Для заданного сочетания не найдено равновесие сечения с трещинами."
            If stateResult.ExtensionUsed Or Not stateResult.WithinPhysicalRange Then _
                blockReason = "НДС с трещинами найдено с выходом за физические пределы диаграмм материалов."
        End If
        Exit Function
    End If

    blockReason = "Текущее НДС для расчета трещин не получено: отсутствует расчетное состояние."
    If Not runner Is Nothing Then
        If Not runner.ResultSolver Is Nothing Then
            If Not runner.ResultSolver.Converged Then
                If Len(Trim$(runner.ResultSolver.StopReason)) > 0 Then blockReason = runner.ResultSolver.StopReason
            End If
        End If
    End If

    Set CurrentCrackedStateMeta = MissingStateMeta(runner, blockReason)
End Function

' Создаёт статус для случая, когда обязательное НДС не получено.
' Причина не угадывается по тексту blockReason: фиксируется отсутствие
' состояния, а подробности сохраняются как пояснение и диагностика.
Private Function MissingStateMeta(ByVal runner As CStateSolutionRunner, _
        ByVal blockReason As String) As CResultMeta
    Dim meta As CResultMeta
    Set meta = New CResultMeta

    If Not runner Is Nothing Then
        If Not runner.ResultSolver Is Nothing Then
            If Not runner.ResultSolver.Converged Then
                meta.SetSolverFailure runner.ResultSolver.FailureCode, rkDirectState, blockReason, _
                    (runner.ResultSolver.Iterations > 0)
                Set MissingStateMeta = meta
                Exit Function
            End If
        End If
    End If

    meta.SetResult rsInternalError, rcSolverDidNotReturnState, rkDirectState, blockReason, _
        vbNullString, True, False
    Set MissingStateMeta = meta
End Function

' Записывает в краткий журнал причину неудачи подготовки или расчёта трещин.
' Подробный отчёт содержит отдельный DiagnosticLog; исходная причина этапа
' сохраняется для диагностики и повторяемой проверки результатов.
Private Sub AppendCrackDiagnostic(ByVal index As Long, ByVal crack As CCrackWidthCalculator)
    If crack Is Nothing Then Exit Sub

    mDiagnosticLog = mDiagnosticLog & "combination=" & mInputs(index).CombinationID & _
        "; crackStopReason=" & crack.StopReason & vbCrLf
    If Len(crack.DiagnosticLog) > 0 Then
        mDiagnosticLog = mDiagnosticLog & crack.DiagnosticLog
    End If
End Sub

' Объединяет статусы нескольких решений НДС одного сочетания.
' N/A не ухудшает итог, а физический отказ, несходимость и ошибка ввода
' сохраняют соответствующую причину.
Private Sub MergeDirectStateMeta(ByVal index As Long, ByVal meta As CResultMeta)
    mResults(index).MergeDirectStateMeta meta
End Sub

' Отмечает несущую способность как незапрошенную, если её проверка
' не предусмотрена выбранной группой или режимом расчёта.
Private Sub StoreNotCalculatedCapacity(ByVal index As Long)
    Dim capacityResult As CCapacityResult
    Set capacityResult = New CCapacityResult
    capacityResult.InitializeNotRequested "Проверка несущей способности не запрошена профилем."
    mResults(index).StoreCapacityResult capacityResult
End Sub

' Сохраняет ошибку ввода через общий объект результата несущей способности.
' Даже если сочетание отклонено до решения НДС, этот раздел Results
' получает статус и пояснение тем же способом, что обычный расчёт предела.
Private Sub StoreInvalidCapacity(ByVal index As Long, ByVal reason As String)
    Dim capacityResult As CCapacityResult
    Set capacityResult = New CCapacityResult
    capacityResult.InitializeInvalidInput reason, CAPACITY_PATH_NONE
    mResults(index).StoreCapacityResult capacityResult
End Sub

' Создаёт статус проверки, отключённой профилем. Это не физическая
' неприменимость и не неудачная попытка расчёта: проверка не запускалась.
Private Function NotRequestedMeta(ByVal resultKind As EResultKind) As CResultMeta
    Dim meta As CResultMeta
    Set meta = New CResultMeta
    meta.SetNotRequested resultKind
    Set NotRequestedMeta = meta
End Function

' Создаёт статус и пояснение ошибки пользовательских данных или настроек.
Private Function InvalidInputMeta(ByVal resultKind As EResultKind, _
        ByVal reason As String) As CResultMeta
    Dim meta As CResultMeta
    Set meta = New CResultMeta
    meta.SetResult rsInvalidInput, rcInvalidInput, resultKind, reason
    Set InvalidInputMeta = meta
End Function

' Сохраняет InvalidInput для выбранного сочетания и не передаёт его в решатель.
' Причина выводится в таблице, а остальные сочетания продолжают рассчитываться.
Private Sub MarkInvalid(ByVal index As Long, ByVal reason As String)
    If Len(Trim$(reason)) = 0 Then reason = "Некорректное сочетание нагрузок."
    mResults(index).SetProfileId mInputs(index).ProfileId
    mResults(index).SetWorkflowMeta InvalidInputMeta(rkGeneric, reason)
    If mProfiles.HasProfile(mInputs(index).ProfileId) Then ' Ошибка одной строки относится только к запрошенным проверкам. Неизвестный профиль не дает права выдумывать набор активных расчетов.
        Dim profile As CCalculationProfile
        Set profile = mProfiles.ProfileById(mInputs(index).ProfileId)
        If profile.StrengthDirectStateEnabled Then mResults(index).SetDirectStateMeta InvalidInputMeta(rkDirectState, reason)
        If profile.StrengthCapacityEnabled Then StoreInvalidCapacity index, reason
        If profile.StabilityEnabled Then mResults(index).SetStabilityMeta InvalidInputMeta(rkStability, reason), reason
        If profile.CrackWidthEnabled Then mResults(index).CrackResult.SetInputFailure reason
    End If
    If Not mMaterialProvider Is Nothing Then
        mResults(index).SetDiagramExtensionEnabled mMaterialProvider.DiagramExtensionEnabled
    End If
    AppendDiagnostic index
End Sub

' Передаёт CCapacityCalculator общие настройки поиска предела и решения равновесия.
' Физические пределы материалов тот получает из диаграмм выбранного профиля.
Private Sub ConfigureCapacityCalculator(ByVal calculator As CCapacityCalculator)
    calculator.ConfigureSearch mCapacityInitialLambda, mCapacitySolutionStrategy, _
        mCapacitySearchMethod, mCapacityMaxLambda, mCapacityToleranceLambda, _
        mCapacityToleranceStrain, mCapacityMaxRetries, mCapacityBaseLoadSteps, _
        mCapacitySolverMaxIterations
    calculator.ConfigureEquilibriumSolver mSolverToleranceN, mSolverToleranceMx, _
        mSolverToleranceMy, mSolverMethod, mSolverSecantMaxRestarts, _
        DetailedDiagnosticsEnabled()
End Sub

' Создаёт CStateProvider для одного сочетания и подключает хранилище его НДС.
' По запросу он возвращает пригодное прежнее состояние или решает равновесие заново.
Private Function CreateStateProvider(ByVal index As Long) As CStateProvider
    Dim provider As CStateProvider
    Set provider = New CStateProvider
    provider.Initialize mSection, mMaterialProvider, mResults(index).StateRepository
    ConfigureStateProvider provider
    Set CreateStateProvider = provider
End Function

' Передаёт CStateProvider прочитанные численные настройки НДС,
' которые применяются общим CStateSolutionRunner ко всем запрашиваемым состояниям.
Private Sub ConfigureStateProvider(ByVal provider As CStateProvider)
    provider.LoadSteps = mSolverLoadSteps
    provider.SolverMethod = mSolverMethod
    provider.MaxIterations = mSolverMaxIterations
    provider.ToleranceN = mSolverToleranceN
    provider.ToleranceMx = mSolverToleranceMx
    provider.ToleranceMy = mSolverToleranceMy
    provider.SecantMaxRestarts = mSolverSecantMaxRestarts
    provider.DiagnosticsEnabled = DetailedDiagnosticsEnabled()
End Sub

' Передаёт параметры подготовки области и формулы ширины трещины.
' Их допустимость проверяет сам расчёт ширины; текущее НДС и его численные
' настройки подготавливаются отдельным объектом решения.
Private Sub ConfigureCrackCalculator(ByVal crack As CCrackWidthCalculator)
    crack.StandardCode = mCrackStandardCode
    crack.SP35GroupGapTolerance = mCrackSP35GroupGap
    crack.SP35RowTolerance = mCrackSP35RowTolerance
    crack.SP35NeighborRatioLimit = mCrackSP35NeighborRatioLimit
    crack.SP35RadiusDiameterMode = mCrackSP35DiameterMode
    crack.SP35InteractionRadiusMode = mCrackSP35RadiusMode
    crack.SP35RebarProfile = mCrackSP35RebarProfile
    crack.AllowableCrackWidth = mCrackAllowable
    crack.PsiMode = mCrackPsiMode
    crack.SigmaSCrcAveragingMode = mCrackSigmaSCrcAveragingMode
    crack.Phi1 = mCrackPhi1
    crack.Phi2 = mCrackPhi2
    crack.Phi3Mode = mCrackPhi3Mode
    crack.Phi3 = mCrackPhi3
    crack.PsiS = mCrackPsiS
    crack.TensionZoneMode = mCrackTensionZoneMode
    crack.CoverMode = mCrackCoverDistanceMode
    crack.DiagnosticsEnabled = DetailedDiagnosticsEnabled()
End Sub

' Передаёт численные настройки расчёту образования трещины.
' Расчёт ширины не выбирает траекторию нагрузки и не ищет состояния
' до и после образования трещины.
Private Sub ConfigureCrackFormationCalculator(ByVal formation As CCrackFormationCalculator, ByVal index As Long)
    formation.CrackFormationPath = mWorkingPaths(index)
    formation.CrackFormationSolutionStrategy = mCrackFormationSolutionStrategy
    formation.LambdaTolerance = mCrackToleranceLambda
    formation.SolverLoadSteps = mSolverLoadSteps
    formation.SolverMethod = mSolverMethod
    formation.SolverMaxIterations = mSolverMaxIterations
    formation.SolverToleranceN = mSolverToleranceN
    formation.SolverToleranceMx = mSolverToleranceMx
    formation.SolverToleranceMy = mSolverToleranceMy
    formation.SolverSecantMaxRestarts = mSolverSecantMaxRestarts
    formation.DiagnosticsEnabled = DetailedDiagnosticsEnabled()
End Sub

' Проверяет, нужен ли подробный журнал внутренних итераций.
' Краткая сводка по сочетаниям всегда остаётся в памяти для тестов и диагностики, но
' тяжелые строки CSectionSolver/CCapacitySolver/CCrackWidthCalculator имеют
' смысл только тогда, когда пользователь включил txt-отчет выполнения.
Private Function DetailedDiagnosticsEnabled() As Boolean
    If mReport Is Nothing Then Exit Function
    DetailedDiagnosticsEnabled = mReport.Enabled
End Function

' Проверяет готовность к запуску расчётов: модель сечения
' и хотя бы одно сочетание нагрузок уже должны быть заданы.
Private Sub ValidateReady()
    If mSection Is Nothing Then Err.Raise vbObjectError + 3920, "CBatchSectionCalculator", "Пакетный расчетчик не инициализирован моделью сечения."
    If mMaterialProvider Is Nothing Then Err.Raise vbObjectError + 3922, "CBatchSectionCalculator", "Пакетный расчетчик не инициализирован поставщиком материалов."
    If mProfiles Is Nothing Then Err.Raise vbObjectError + 3923, "CBatchSectionCalculator", "Каталог расчетных профилей не загружен."
    If mCount <= 0 Then Err.Raise vbObjectError + 3921, "CBatchSectionCalculator", "Не задано ни одного сочетания нагрузок."
End Sub

' Сбрасывает все результаты одной строки перед повторным расчетом LC.
' Входные усилия и комментарий не трогаются.
Private Sub ClearResultRow(ByVal index As Long)
    Set mResults(index) = New CCombinationResult ' Прежний результат мог быть передан классу записи на лист или другому потребителю. Новый запуск создаёт отдельный объект, не меняя уже выданный результат.
    mResults(index).Clear CAPACITY_PATH_NONE
    mResults(index).SetDirectStateMeta NotRequestedMeta(rkDirectState)
    mResults(index).SetStabilityMeta NotRequestedMeta(rkStability)
    mResults(index).CrackResult.SetNotRequested
    mResults(index).SetProfileId mInputs(index).ProfileId
End Sub

' Добавляет строку в краткий журнал пакетного расчёта.
' Подробные итерации численного решения сохраняются отдельно.
Private Sub AppendDiagnostic(ByVal index As Long)
    mDiagnosticLog = mDiagnosticLog & "combination=" & mInputs(index).CombinationID & _
        "; OverallStatus=" & mResults(index).Status & _
        "; DirectStateStatus=" & mResults(index).StrengthResult.DirectState.Status & _
        "; ExtensionUsed=" & CStr(mResults(index).ExtensionUsed) & _
        "; CapacityStatus=" & mResults(index).StrengthResult.Capacity.Status & _
        "; CrackStatus=" & mResults(index).NormalCrackStatus & _
        "; LongitudinalCrackStatus=" & mResults(index).CrackResult.Longitudinal.Status & _
        "; LongitudinalCrackUtil=" & FormatNumberInvariant(mResults(index).CrackResult.Longitudinal.Utilization) & _
        "; StabilityStatus=" & mResults(index).StabilityResult.Status & _
        "; StabilityUtil=" & FormatNumberInvariant(mResults(index).StabilityResult.Utilization) & _
        "; governingUtil=" & FormatNumberInvariant(mResults(index).GoverningUtil) & vbCrLf
End Sub

' Добавляет краткий итог сочетания в подробный отчёт выполнения.
' Основная сводка остаётся на Results; отчёт показывает ход расчёта.
Private Sub ReportCombinationSummary(ByVal index As Long)
    ReportValue "Итоговый статус", mResults(index).Status
    ReportValue "Статус прямого НДС", mResults(index).StrengthResult.DirectState.Status
    ReportValue "Использовано продолжение диаграммы", mReport.YesNoRu(mResults(index).ExtensionUsed)
    ReportValue "Статус несущей способности", mResults(index).StrengthResult.Capacity.Status
    ReportValue "Статус нормальных трещин", mResults(index).NormalCrackStatus
    ReportValue "Статус продольных трещин", mResults(index).CrackResult.Longitudinal.Status
    ReportValue "Использование по продольным трещинам", FormatNumberInvariant(mResults(index).CrackResult.Longitudinal.Utilization)
    ReportValue "Статус устойчивости", mResults(index).StabilityResult.Status
    ReportValue "Использование устойчивости", FormatNumberInvariant(mResults(index).StabilityResult.Utilization)
    ReportValue "Определяющее использование", FormatNumberInvariant(mResults(index).GoverningUtil)
    ReportMetaComment "Итоговый комментарий", mResults(index).OverallMeta
End Sub

' Добавляет заголовок раздела в текстовый отчёт выполнения, если он включён.
Private Sub ReportSection(ByVal title As String)
    If mReport Is Nothing Then Exit Sub
    If Not mReport.Enabled Then Exit Sub
    mReport.AddSection title
End Sub

' Добавляет название выполняемого этапа в отчёт, чтобы пользователь видел,
' какая проверка запускается сейчас.
Private Sub ReportStep(ByVal text As String)
    If mReport Is Nothing Then Exit Sub
    If Not mReport.Enabled Then Exit Sub
    mReport.AddStep text
End Sub

' Добавляет параметр и его значение в отчёт, если отчёт включён.
' Вызывающему расчёту отдельно проверять эту настройку не требуется.
Private Sub ReportValue(ByVal label As String, ByVal valueText As String)
    If mReport Is Nothing Then Exit Sub
    If Not mReport.Enabled Then Exit Sub
    mReport.AddValue label, valueText
End Sub

' Добавляет готовое пояснение ResultMeta нужной проверки в отчёт выполнения.
' Причины уже сформированы расчётом и при записи заново не собираются.
Private Sub ReportMetaComment(ByVal label As String, ByVal meta As CResultMeta)
    If meta Is Nothing Then Exit Sub
    If Len(Trim$(meta.ResultComment)) = 0 Then Exit Sub
    ReportValue label, meta.ResultComment
End Sub

' Добавляет многострочную диагностику в отчёт, предварительно
' преобразуя служебные записи решателя в понятный текст.
Private Sub ReportBlock(ByVal title As String, ByVal body As String)
    If mReport Is Nothing Then Exit Sub
    If Not mReport.Enabled Then Exit Sub
    mReport.AddDiagnosticBlock title, body
End Sub

' Выбирает определяющее сочетание прочности по минимальному положительному запасу.
' Общий выбор Worst по General.WorstCombinationCriterion выполняется отдельно.
Private Function IsBetterGoverningCandidate(ByVal candidateIndex As Long, ByVal currentIndex As Long) As Boolean
    Dim candidateSafety As Double
    Dim currentSafety As Double
    candidateSafety = StrengthSafetyFactorForGoverning(candidateIndex)
    currentSafety = StrengthSafetyFactorForGoverning(currentIndex)

    If candidateSafety <= 0# Then
        IsBetterGoverningCandidate = False
    ElseIf currentSafety <= 0# Then
        IsBetterGoverningCandidate = True
    Else
        IsBetterGoverningCandidate = (candidateSafety < currentSafety)
    End If
End Function

' Возвращает прочностной запас LC как найденный множитель lambda.
' Для любого LoadPath множитель одинаково характеризует запас выбранной
' траектории: lambda*Mx, lambda*My, lambda*Mxy, lambda*N или lambda*NMxy.
Private Function StrengthSafetyFactorForGoverning(ByVal index As Long) As Double
    StrengthSafetyFactorForGoverning = mResults(index).StrengthResult.Capacity.ReserveFactor
End Function

' Сравнивает только запасы по раскрытию трещин. Нулевой запас означает, что
' проверка трещин для LC не считалась или ширина раскрытия равна нулю.
Private Function IsBetterCrackGoverningCandidate(ByVal candidateIndex As Long, ByVal currentIndex As Long) As Boolean
    Dim candidateSafety As Double
    Dim currentSafety As Double
    candidateSafety = CrackSafetyFactorForGoverning(candidateIndex)
    If candidateSafety <= 0# Then Exit Function

    If currentIndex <= 0 Then
        IsBetterCrackGoverningCandidate = True
    Else
        currentSafety = CrackSafetyFactorForGoverning(currentIndex)
        IsBetterCrackGoverningCandidate = (currentSafety <= 0# Or candidateSafety < currentSafety)
    End If
End Function

' Возвращает запас по трещинам как величину, обратную использованию.
Private Function CrackSafetyFactorForGoverning(ByVal index As Long) As Double
    CrackSafetyFactorForGoverning = mResults(index).CrackResult.Width.ReserveFactor
End Function

' Выбирает сочетание Worst для верхней сводки по пользовательскому критерию.
' Его же используют схема Plot.LoadCase = Worst и экспорт AutoCAD
' с AutoCAD.Export.CombinationID = Worst.
Private Sub SelectWorstCombination()
    Dim i As Long
    mWorstCombinationIndex = 0
    For i = 1 To mCount
        If mInputs(i).InputValid Then
            If IsBetterWorstCandidate(i, mWorstCombinationIndex) Then mWorstCombinationIndex = i
        End If
    Next i
    If mWorstCombinationIndex <= 0 Then
        If mGoverningCombinationIndex > 0 Then
            mWorstCombinationIndex = mGoverningCombinationIndex
        Else
            For i = 1 To mCount
                If mInputs(i).InputValid Then
                    mWorstCombinationIndex = i
                    Exit For
                End If
            Next i
        End If
    End If
End Sub

' Сравнивает LC по минимальному положительному запасу выбранного типа.
' Нулевой запас означает, что соответствующая проверка для LC не применялась
' или не дала физически интерпретируемого результата.
Private Function IsBetterWorstCandidate(ByVal candidateIndex As Long, ByVal currentIndex As Long) As Boolean
    Dim candidateSafety As Double
    Dim currentSafety As Double
    candidateSafety = mResults(candidateIndex).ReserveForCriterion(mWorstCombinationCriterion)
    If candidateSafety <= 0# Then Exit Function

    If currentIndex <= 0 Then
        IsBetterWorstCandidate = True
    Else
        currentSafety = mResults(currentIndex).ReserveForCriterion(mWorstCombinationCriterion)
        IsBetterWorstCandidate = (currentSafety <= 0# Or candidateSafety < currentSafety)
    End If
End Function

' Форматирует число с точкой как десятичным разделителем для диагностической строки.
Private Function FormatNumberInvariant(ByVal value As Double) As String
    FormatNumberInvariant = Replace$(Format$(value, "0.############"), ",", ".")
End Function

' Сколько сочетаний реально добавлено в пакет.
Public Property Get Count() As Long
    Count = mCount
End Property

' Исходные данные LC, прочитанные из фактических строк rngLoadCombinations.
Public Property Get CombinationID(ByVal index As Long) As String
    CombinationID = mInputs(index).CombinationID
End Property

' Возвращает номер строки сочетания в данных rngLoadCombinations, начиная с 1.
' Это номер для записи Results; внутренний номер рассчитанного сочетания
' может отличаться из-за пропущенных пустых строк.
Public Property Get SourceDataOffset(ByVal index As Long) As Long
    SourceDataOffset = mInputs(index).SourceDataOffset
End Property

' Возвращает сохранённое место задания нагрузок; адрес получает ридер исходной таблицы.
' Для программно заданного сочетания возвращается пустая строка, адрес не выдумывается.
Public Property Get SourceReference(ByVal index As Long) As String
    SourceReference = mInputs(index).SourceReference
End Property

' Сопоставляет переданный расчётом ключ параметра с сохранённым местом его задания.
' Неизвестный ключ остаётся без адреса, чтобы не указывать пользователю постороннюю ячейку.
Public Function SettingLocation(ByVal key As String) As String
    If mSettingLocations Is Nothing Then Exit Function
    If mSettingLocations.Exists(key) Then SettingLocation = mSettingLocations(key)
End Function

' Разрешает ключи параметров по сохранённому снимку. Сводки и итоговое сообщение
' используют один способ сопоставления, не ищут ключи в тексте причины и не читают Excel.
Public Function SettingReferences(ByVal sourceKeys As String) As String
    Dim key As Variant, location As String
    If Len(sourceKeys) = 0 Then Exit Function
    For Each key In Split(sourceKeys, ";")
        location = SettingLocation(Trim$(CStr(key)))
        If Len(location) > 0 Then
            If Len(SettingReferences) > 0 Then SettingReferences = SettingReferences & " "
            SettingReferences = SettingReferences & Trim$(CStr(key)) & ": " & location & "."
        End If
    Next key
End Function

' Исходные данные LC, прочитанные из фактических строк rngLoadCombinations.
Public Property Get CombinationName(ByVal index As Long) As String
    CombinationName = mInputs(index).Comment
End Property

' Возвращает имя профиля из исходной строки до выполнения расчёта.
' Подготовка таблиц Results использует этот ввод, а не ещё пустые результаты.
Public Property Get InputProfileId(ByVal index As Long) As String
    InputProfileId = mInputs(index).ProfileId
End Property

' Возвращает рабочую продольную силу текущего расчёта, Н; исходная строка сочетания сохраняется.
Public Property Get N(ByVal index As Long) As Double
    N = mWorkingLoads(index).N
End Property

' Показывает, что исходная строка усилий прочитана без ошибки ввода.
' Нулевые заглушки ошибочной строки не выдаются за заданные нагрузки.
Public Property Get OriginalLoadsAvailable(ByVal index As Long) As Boolean
    OriginalLoadsAvailable = mInputs(index).InputValid
End Property

' Возвращает заданную N до фильтров и продольного изгиба, во внутренних Н.
Public Property Get OriginalN(ByVal index As Long) As Double
    OriginalN = mInputs(index).N
End Property

' Возвращает заданный Mx в пользовательской точке до фильтров, Н*мм.
Public Property Get OriginalMx(ByVal index As Long) As Double
    OriginalMx = mInputs(index).Mx
End Property

' Возвращает заданный My в пользовательской точке до фильтров, Н*мм.
Public Property Get OriginalMy(ByVal index As Long) As Double
    OriginalMy = mInputs(index).My
End Property

' Возвращает рабочий момент Mx относительно начала координат модели, Н*мм.
Public Property Get Mx(ByVal index As Long) As Double
    Mx = mWorkingLoads(index).InternalMx
End Property

' Возвращает рабочий момент My относительно начала координат модели, Н*мм.
Public Property Get My(ByVal index As Long) As Double
    My = mWorkingLoads(index).InternalMy
End Property

' Возвращает рабочий момент Mx в пользовательской точке приложения нагрузки, Н*мм.
Public Property Get UserMx(ByVal index As Long) As Double
    UserMx = mWorkingLoads(index).LoadPointMx
End Property

' Возвращает рабочий момент My в пользовательской точке приложения нагрузки, Н*мм.
Public Property Get UserMy(ByVal index As Long) As Double
    UserMy = mWorkingLoads(index).LoadPointMy
End Property

' X точки приложения нагрузки относительно расчетной модели, мм.
Public Property Get LoadReferenceX() As Double
    LoadReferenceX = mLoadReferenceX
End Property

' Y точки приложения нагрузки относительно расчетной модели, мм.
Public Property Get LoadReferenceY() As Double
    LoadReferenceY = mLoadReferenceY
End Property

' Смещение точки нагрузки от центра тяжести бетонного сечения по X, мм.
Public Property Get LoadReferenceOffsetX() As Double
    LoadReferenceOffsetX = mLoadReferenceOffsetX
End Property

' Смещение точки нагрузки от центра тяжести бетонного сечения по Y, мм.
Public Property Get LoadReferenceOffsetY() As Double
    LoadReferenceOffsetY = mLoadReferenceOffsetY
End Property

' Геометрическая длина элемента l, мм.
Public Property Get StabilityElementLength() As Double
    StabilityElementLength = mStabilityElementLength
End Property

' Коэффициент расчетной длины для первой главной плоскости.
Public Property Get StabilityMu1() As Double
    StabilityMu1 = mStabilityMu1
End Property

' Коэффициент расчетной длины для второй главной плоскости.
Public Property Get StabilityMu2() As Double
    StabilityMu2 = mStabilityMu2
End Property

' Возвращает общий объект результатов сочетания для чтения.
' Отдельные копии всех численных полей в пакетном расчёте не хранятся.
Public Property Get ResultAt(ByVal index As Long) As CCombinationResult
    If index < 1 Or index > mCount Then Err.Raise vbObjectError + 3943, _
        "CBatchSectionCalculator", "Номер результата сочетания вне диапазона пакета."
    mResults(index).Freeze
    Set ResultAt = mResults(index)
End Property

' Возвращает Mx после учёта устойчивости в пользовательской точке приложения нагрузки, Н*мм.
Public Property Get StabilityLoadPointMx(ByVal index As Long) As Double
    StabilityLoadPointMx = mWorkingLoads(index).LoadPointMx
End Property

' Возвращает My после устойчивости в пользовательской точке задания нагрузки.
Public Property Get StabilityLoadPointMy(ByVal index As Long) As Double
    StabilityLoadPointMy = mWorkingLoads(index).LoadPointMy
End Property

' Возвращает пользовательский предел N/Ncr для СП 35. Он же указывается
' в заголовках Results, чтобы формула запаса соответствовала настройке Config.
Public Property Get StabilitySP35NOverNcrLimit() As Double
    StabilitySP35NOverNcrLimit = mStabilitySP35NOverNcrLimit
End Property

' Усилия трещин относятся к сохраненной нагрузке LC после устойчивости.
' Продольная сила не зависит от точки приведения.
Public Property Get CrackReferenceN(ByVal index As Long) As Double
    CrackReferenceN = LoadStateForCombination(index).N
End Property

' Переносит Mx в принятую для трещин точку; добавка определяется N и ее Y.
Public Property Get CrackReferenceMx(ByVal index As Long) As Double
    Dim referenceX As Double
    Dim referenceY As Double
    CrackReferencePointForCombination index, referenceX, referenceY
    CrackReferenceMx = LoadStateForCombination(index).MxAboutPoint(referenceY)
End Property

' Переносит My в принятую для трещин точку; добавка определяется N и ее X.
Public Property Get CrackReferenceMy(ByVal index As Long) As Double
    Dim referenceX As Double
    Dim referenceY As Double
    CrackReferencePointForCombination index, referenceX, referenceY
    CrackReferenceMy = LoadStateForCombination(index).MyAboutPoint(referenceX)
End Property

' Усилие N в принятой точке приложения для проверки трещинообразования, Н.
Public Property Get CrackLoadPointN(ByVal index As Long) As Double
    CrackLoadPointN = LoadStateForCombination(index).N
End Property

' Усилие Mx в принятой точке приложения для проверки трещинообразования, Н·мм.
Public Property Get CrackLoadPointMx(ByVal index As Long) As Double
    CrackLoadPointMx = LoadStateForCombination(index).LoadPointMx
End Property

' Усилие My в принятой точке приложения для проверки трещинообразования, Н·мм.
Public Property Get CrackLoadPointMy(ByVal index As Long) As Double
    CrackLoadPointMy = LoadStateForCombination(index).LoadPointMy
End Property

' Возвращает подготовку диаграмм материалов для прочности, образования трещины и НДС с
' трещинами.
Public Property Get CrackSteelElasticModulus() As Double
    CrackSteelElasticModulus = mMaterialProvider.SteelElasticModulus(cpCrackedNDS)
End Property

' Предельная сжатая деформация бетона именно для модели прочности.
Public Function StrengthConcreteCompressionLimitForCombination(ByVal index As Long) As Double
    Dim spec As CMaterialModelSpec
    Set spec = StrengthMaterialSpecForCombination(index)
    If Not spec Is Nothing Then StrengthConcreteCompressionLimitForCombination = mMaterialProvider.ConcreteCompressionLimitFromSpec(spec)
End Function

' Предельная растянутая деформация бетона именно для модели прочности.
Public Function StrengthConcreteTensionLimitForCombination(ByVal index As Long) As Double
    Dim spec As CMaterialModelSpec
    Set spec = StrengthMaterialSpecForCombination(index)
    If Not spec Is Nothing Then StrengthConcreteTensionLimitForCombination = mMaterialProvider.ConcreteTensionLimitFromSpec(spec)
End Function

' Показывает, учитывает ли модель прочности растяжение бетона и нужно ли выводить его предел.
Public Function StrengthConcreteTensionLimitEnabledForCombination(ByVal index As Long) As Boolean
    Dim spec As CMaterialModelSpec
    Set spec = StrengthMaterialSpecForCombination(index)
    If Not spec Is Nothing Then StrengthConcreteTensionLimitEnabledForCombination = mMaterialProvider.ConcreteTensionLimitEnabledFromSpec(spec)
End Function

' Предельная сжатая деформация арматуры именно для модели прочности.
Public Function StrengthSteelCompressionLimitForCombination(ByVal index As Long) As Double
    Dim spec As CMaterialModelSpec
    Set spec = StrengthMaterialSpecForCombination(index)
    If Not spec Is Nothing Then StrengthSteelCompressionLimitForCombination = mMaterialProvider.SteelCompressionLimitFromSpec(spec)
End Function

' Предельная растянутая деформация арматуры именно для модели прочности.
Public Function StrengthSteelTensionLimitForCombination(ByVal index As Long) As Double
    Dim spec As CMaterialModelSpec
    Set spec = StrengthMaterialSpecForCombination(index)
    If Not spec Is Nothing Then StrengthSteelTensionLimitForCombination = mMaterialProvider.SteelTensionLimitFromSpec(spec)
End Function

' Возвращает фактические параметры материалов для подробной таблицы прочности.
' Сначала берёт их из сохранённого НДС. Если состояния нет, использует профиль
' для вывода пределов в строках, где расчёт не запускался или завершился ошибкой.
Private Function StrengthMaterialSpecForCombination(ByVal index As Long) As CMaterialModelSpec
    Dim stateResult As CSectionStateResult
    Set stateResult = mResults(index).StateRepository.FindState(sstStrengthState)
    If stateResult Is Nothing Then Set stateResult = mResults(index).StateRepository.FindState(sstCapacityState)
    If Not stateResult Is Nothing Then
        Set StrengthMaterialSpecForCombination = stateResult.MaterialSpec
        Exit Function
    End If

    Dim profile As CCalculationProfile
    If mProfiles.HasProfile(mInputs(index).ProfileId) Then
        Set profile = mProfiles.ProfileById(mInputs(index).ProfileId)
        If Not profile Is Nothing Then Set StrengthMaterialSpecForCombination = profile.StrengthSpec
    End If
End Function

' Возвращает индекс, ID и подпись сочетания, определяющего итог прочностного
' расчета. До выбора такого сочетания индекс равен нулю, строки остаются пустыми.
Public Property Get GoverningCombinationIndex() As Long
    GoverningCombinationIndex = mGoverningCombinationIndex
End Property

' Возвращает номер сочетания, определяющего результат проверки прочности.
Public Property Get GoverningCombinationID() As String
    If mGoverningCombinationIndex > 0 Then GoverningCombinationID = mInputs(mGoverningCombinationIndex).CombinationID
End Property

' Возвращает имя определяющего сочетания для общей сводки пакетного расчёта.
Public Property Get GoverningCombinationName() As String
    If mGoverningCombinationIndex > 0 Then GoverningCombinationName = mInputs(mGoverningCombinationIndex).Comment
End Property

' Возвращает данные сочетания с определяющим результатом проверки трещин.
' Отсутствие выбранного сочетания обозначается нулевым индексом и пустыми строками.
Public Property Get CrackGoverningCombinationIndex() As Long
    CrackGoverningCombinationIndex = mCrackGoverningCombinationIndex
End Property

' Индекс LC с минимальным запасом по раскрытию трещин.
Public Property Get CrackGoverningCombinationID() As String
    If mCrackGoverningCombinationIndex > 0 Then CrackGoverningCombinationID = mInputs(mCrackGoverningCombinationIndex).CombinationID
End Property

' Возвращает имя сочетания с минимальным запасом по раскрытию трещин.
Public Property Get CrackGoverningCombinationName() As String
    If mCrackGoverningCombinationIndex > 0 Then CrackGoverningCombinationName = mInputs(mCrackGoverningCombinationIndex).Comment
End Property

' Возвращает данные худшего сочетания по настроенному итоговому критерию.
' Этот выбор может отличаться от определяющего сочетания отдельного раздела.
Public Property Get WorstCombinationIndex() As Long
    WorstCombinationIndex = mWorstCombinationIndex
End Property

' Возвращает номер сочетания Worst по выбранному критерию общей сводки.
Public Property Get WorstCombinationID() As String
    If mWorstCombinationIndex > 0 Then WorstCombinationID = mInputs(mWorstCombinationIndex).CombinationID
End Property

' Возвращает имя сочетания Worst по выбранному критерию общей сводки.
Public Property Get WorstCombinationName() As String
    If mWorstCombinationIndex > 0 Then WorstCombinationName = mInputs(mWorstCombinationIndex).Comment
End Property

' Возвращает критерий выбора определяющего сочетания: деформации, несущая способность, трещины
' или устойчивость.
Public Property Get WorstCombinationCriterion() As String
    WorstCombinationCriterion = mWorstCombinationCriterion
End Property

' Возвращает число вызовов CSectionSolver.Solve во всём пакетном расчёте.
Public Property Get SolverCallCount() As Long
    SolverCallCount = mSolverCallCount
End Property

' Возвращает общее число фактически выполненных итераций НДС.
Public Property Get SolverIterationCount() As Long
    SolverIterationCount = mSolverIterationCount
End Property

' Возвращает число попыток поиска предела с разными начальными полями деформаций.
Public Property Get UltimateStrainStartCount() As Long
    UltimateStrainStartCount = mUltimateStrainStartCount
End Property

' Возвращает общее число итераций UltimateStrain для прочности и образования трещины.
Public Property Get UltimateStrainIterationCount() As Long
    UltimateStrainIterationCount = mUltimateStrainIterationCount
End Property

' Полное время пакетного расчета, сек.
Public Property Get ElapsedSeconds() As Double
    ElapsedSeconds = mElapsedSeconds
End Property

' Возвращает краткий журнал выполнения отдельных сочетаний нагрузок.
Public Property Get DiagnosticLog() As String
    DiagnosticLog = mDiagnosticLog
End Property

' Считает сочетания, завершившиеся ошибкой ввода по общим правилам статусов.
' Физический отказ и численная несходимость в это число не входят.
Public Property Get InvalidInputCount() As Long
    Dim i As Long
    For i = 1 To mCount
        If mStatusPolicy.IsInvalid(mResults(i).Status) Then InvalidInputCount = InvalidInputCount + 1
    Next i
End Property

' Находит первую ошибку входа в порядке строк сочетаний и возвращает сообщение
' с идентификатором строки. При отсутствии InputErr возвращает пустую строку.
Public Property Get FirstInvalidInputMessage() As String
    Dim i As Long
    For i = 1 To mCount
        If mStatusPolicy.IsInvalid(mResults(i).Status) Then
            FirstInvalidInputMessage = DetailedInvalidInputMessage(i)
            Exit Property
        End If
    Next i
End Property

' Возвращает первую человекочитаемую причину InputErr для итогового сообщения.
' Короткий статус остается в Results, а popup показывает, какая настройка,
' профиль или геометрическая величина фактически остановила расчет.
Private Function DetailedInvalidInputMessage(ByVal index As Long) As String
    Dim detail As String, references As String
    detail = Trim$(mInputs(index).InputError)
    If Len(detail) = 0 Then detail = Trim$(mResults(index).OverallMeta.ResultComment) ' Берём готовое пояснение общего результата, включая трещины; машинный код предела и краткий статус его не заменяют.
    If Len(detail) = 0 Then detail = Trim$(mResults(index).Status)
    If Len(detail) = 0 Then detail = mStatusPolicy.InputErr

    DetailedInvalidInputMessage = "CombinationID=" & DisplayCombinationID(index) & _
        ": " & detail & " (статус " & mResults(index).Status & ")"
    references = SettingReferences(mResults(index).OverallMeta.SourceKeys)
    If Len(references) > 0 Then DetailedInvalidInputMessage = DetailedInvalidInputMessage & " " & references
    If Len(mInputs(index).SourceReference) > 0 Then DetailedInvalidInputMessage = DetailedInvalidInputMessage & _
        " Исходная строка сочетания: " & mInputs(index).SourceReference & "."
End Function

' Формирует устойчивое имя LC для сообщения об ошибке даже при неполной строке
' сочетания, когда пользователь мог оставить CombinationID пустым.
Private Function DisplayCombinationID(ByVal index As Long) As String
    DisplayCombinationID = Trim$(mInputs(index).CombinationID)
    If Len(DisplayCombinationID) = 0 Then DisplayCombinationID = "#" & CStr(index)
End Function

' Читает сохраненную N прочностного LC; перенос точки не меняет силу.
Public Property Get StrengthReferenceN(ByVal index As Long) As Double
    StrengthReferenceN = LoadStateForCombination(index).N
End Property

' Переносит Mx в принятую для прочности точку, сохраняя нагрузку после устойчивости.
Public Property Get StrengthReferenceMx(ByVal index As Long) As Double
    Dim referenceX As Double
    Dim referenceY As Double
    StrengthReferencePointForCombination index, referenceX, referenceY
    StrengthReferenceMx = LoadStateForCombination(index).MxAboutPoint(referenceY)
End Property

' Переносит My в принятую для прочности точку, сохраняя нагрузку после устойчивости.
Public Property Get StrengthReferenceMy(ByVal index As Long) As Double
    Dim referenceX As Double
    Dim referenceY As Double
    StrengthReferencePointForCombination index, referenceX, referenceY
    StrengthReferenceMy = LoadStateForCombination(index).MyAboutPoint(referenceX)
End Property

' Усилие N в принятой точке приложения для проверки прочности, Н.
Public Property Get StrengthLoadPointN(ByVal index As Long) As Double
    StrengthLoadPointN = LoadStateForCombination(index).N
End Property

' Усилие Mx в принятой точке приложения для проверки прочности, Н·мм.
Public Property Get StrengthLoadPointMx(ByVal index As Long) As Double
    StrengthLoadPointMx = LoadStateForCombination(index).LoadPointMx
End Property

' Усилие My в принятой точке приложения для проверки прочности, Н·мм.
Public Property Get StrengthLoadPointMy(ByVal index As Long) As Double
    StrengthLoadPointMy = LoadStateForCombination(index).LoadPointMy
End Property

' Показывает, запрошено ли НДС прочности профилем сочетания.
' Если профиля нет, возвращает False; ошибку профиля обрабатывает общий запуск.
Public Property Get StrengthDirectStateRequested(ByVal index As Long) As Boolean
    If mProfiles.HasProfile(mInputs(index).ProfileId) Then _
        StrengthDirectStateRequested = mProfiles.ProfileById(mInputs(index).ProfileId).StrengthDirectStateEnabled
End Property

' Показывает, запрошена ли проверка несущей способности профилем данного сочетания.
Public Property Get StrengthCapacityRequested(ByVal index As Long) As Boolean
    If mProfiles.HasProfile(mInputs(index).ProfileId) Then _
        StrengthCapacityRequested = mProfiles.ProfileById(mInputs(index).ProfileId).StrengthCapacityEnabled
End Property

' Показывает, запрошен ли расчёт ширины раскрытия трещин профилем данного сочетания.
Public Property Get CrackWidthRequested(ByVal index As Long) As Boolean
    If mProfiles.HasProfile(mInputs(index).ProfileId) Then _
        CrackWidthRequested = mProfiles.ProfileById(mInputs(index).ProfileId).CrackWidthEnabled
End Property

' Показывает, запрошена ли проверка устойчивости профилем данного сочетания.
Public Property Get StabilityRequested(ByVal index As Long) As Boolean
    If mProfiles.HasProfile(mInputs(index).ProfileId) Then _
        StabilityRequested = mProfiles.ProfileById(mInputs(index).ProfileId).StabilityEnabled
End Property

' Считает сочетания, для которых сохранено хотя бы одно НДС.
' Наличие сохранённого состояния само по себе не означает успешной инженерной проверки.
Public Property Get StateAvailableCount() As Long
    Dim i As Long
    For i = 1 To mCount
        If mResults(i).StateRepository.StateCount > 0 Then StateAvailableCount = StateAvailableCount + 1
    Next i
End Property

' ==========================================================================
' ДЛЯ ТЕСТОВ
' ==========================================================================

' -------------------- ДЛЯ ТЕСТОВ --------------------
' возвращает число успешных построений общей области в последнем Execute.
Friend Property Get TestCrackDomainBuildCount() As Long
    TestCrackDomainBuildCount = mCrackDomainBuildCount
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' возвращает число фактических подготовок общих массивов геометрии
' за текущий запуск, без повторного решения НДС.
Friend Property Get TestMechanicalPreparationBuildCount() As Long
    TestMechanicalPreparationBuildCount = mMechanicalPreparationBuildCount
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' проверяет, освобождены ли после Execute ссылки на временные расчётные данные.
Friend Property Get TestRunPreparationReleased() As Boolean
    TestRunPreparationReleased = mCrackGeometryQuery Is Nothing And mRunConcreteProperties Is Nothing And _
        mRunTransformedProperties(1) Is Nothing And mRunTransformedProperties(2) Is Nothing
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' возвращает первый готовый результат трещин и повторяет обычный
' расчёт до завершения сочетания. Новые расчётчики используют уже найденные
' НДС из общего хранилища без нового решения равновесия.
Public Function TestRepeatCrackCalculation(ByVal index As Long) As CCrackResult
    If index < 1 Or index > mCount Then Err.Raise vbObjectError + 4270, _
        "CBatchSectionCalculator", "Индекс тестового сочетания вне диапазона."
    Dim profile As CCalculationProfile
    Set profile = mProfiles.ProfileById(mInputs(index).ProfileId)
    Set TestRepeatCrackCalculation = mResults(index).CrackResult
    TestRepeatCrackCalculation.Freeze
    RunCrackWidth index, profile
End Function
