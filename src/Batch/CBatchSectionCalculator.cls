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
' Пакетный расчет сочетаний нагрузок
' ==========================================================================
' CBatchSectionCalculator запускает один и тот же расчетный набор для всех LC:
' прямое НДС, несущую способность для 1 группы и трещины для 2 группы. Класс
' хранит результаты каждого сочетания в массивах, чтобы writer мог потом
' выгрузить сводку и Results без повторного обращения к solver-ам.
'
' Здесь нет чтения Excel-ячеек и нет генерации геометрии. На вход уже должны
' прийти CSectionModel, материалы и нормализованные усилия N + Mx + My.

Private Const MAX_COMBINATIONS As Long = 20 ' Максимум LC, поддерживаемый таблицей книги.
Private Const CAPACITY_PATH_NONE As String = "None" ' Capacity не запускался для LC.
Private Const DEFAULT_ZERO_MOMENT_PER_DEPTH As Double = 4903.325 ' 0.5 tf*m/m во внутренних Н*мм/мм.

' Данные одной строки таблицы сочетаний. Batch хранит их отдельно от результатов,
' чтобы было видно, где исходные нагрузки пользователя, а где то, что вернул расчет.
Private Type TCombinationInput
    CombinationID As String
    N As Double
    Mx As Double
    My As Double
    UserMx As Double
    UserMy As Double
    ProfileId As String
    CapacityLoadPath As String
    Comment As String
    InputValid As Boolean
    InputError As String
End Type

Private mSection As CSectionModel ' Расчетная модель сечения, общая для всех LC.
Private mMaterialProvider As CMaterialModelProvider ' Единая точка выдачи диаграмм для Strength/Mcrc/CrackedNDS.
Private mProfiles As CCalculationProfileCatalog ' Каталог расчетных профилей PR1/PR2/... из Config.
Private mStatusPolicy As CBatchStatusPolicy ' Единый словарь пользовательских статусов и правил их агрегирования.
Private mCount As Long ' Сколько сочетаний реально добавлено в пакет.

Private mInputs(1 To MAX_COMBINATIONS) As TCombinationInput ' Исходные данные LC, прочитанные из таблицы сочетаний.
Private mResults(1 To MAX_COMBINATIONS) As CCombinationResult ' Расчетные результаты LC для Summary/Results.
Private mLoadReferenceX As Double ' X точки приложения нагрузки относительно расчетной модели, мм.
Private mLoadReferenceY As Double ' Y точки приложения нагрузки относительно расчетной модели, мм.
Private mLoadReferenceOffsetX As Double ' Смещение точки нагрузки от центра тяжести бетонного сечения по X, мм.
Private mLoadReferenceOffsetY As Double ' Смещение точки нагрузки от центра тяжести бетонного сечения по Y, мм.
Private mMomentZeroFilter As CMomentZeroFilter ' Общий инженерный фильтр практически нулевых моментов.
Private mZeroMomentPerDepth As Double ' Порог момента на единицу бетонного габарита, Н*мм/мм.
Private mConcreteDepthForMx As Double ' Бетонный габарит по Y для проверки Mx, мм.
Private mConcreteDepthForMy As Double ' Бетонный габарит по X для проверки My, мм.
Private mGoverningCombinationIndex As Long ' Индекс LC, который стал определяющим в batch summary.
Private mCrackGoverningCombinationIndex As Long ' Индекс LC с минимальным запасом по раскрытию трещин.
Private mElapsedSeconds As Double ' Полное время пакетного расчета, сек.
Private mDiagnosticLog As String ' Короткий текстовый журнал по LC для диагностики batch-потока.

Private mCrackAllowable As Double ' Допустимая ширина раскрытия трещины a_crc,ult, мм.
Private mCrackPsiMode As String ' User или Auto: как определять psi_s для нормальных трещин.
Private mCrackPhi1 As Double ' Пользовательский коэффициент phi1 для формулы раскрытия.
Private mCrackPhi2 As Double ' Пользовательский коэффициент phi2 для формулы раскрытия.
Private mCrackPhi3Mode As String ' User или Auto: как принимать phi3.
Private mCrackPhi3 As Double ' Пользовательский коэффициент phi3, если Phi3Mode = User.
Private mCrackPsiS As Double ' Пользовательский коэффициент psi_s, если PsiMode = User.
Private mCrackTensionZoneMode As String ' Effective или FullTension: какая бетонная зона входит в Abt.
Private mCrackCoverDistanceMode As String ' Как измерять a_s: до ближайшего контура или до дальней растянутой линии.
Private mSolverMethod As String ' Newton или Secant для прямого равновесия.
Private mSolverLoadSteps As Long ' Число ступеней приложения нагрузки в CSectionSolver.
Private mSolverMaxIterations As Long ' Лимит итераций прямого solver-а на одной ступени.
Private mSolverToleranceN As Double ' Допуск равновесия по N, Н.
Private mSolverToleranceMx As Double ' Допуск равновесия по Mx, Н*мм.
Private mSolverToleranceMy As Double ' Допуск равновесия по My, Н*мм.
Private mSolverLineSearchEnabled As Boolean ' Разрешает дробить итерационный шаг, если полный шаг ухудшает невязку.
Private mSolverDampingInitial As Double ' Начальный alpha для line search общего решателя НДС.
Private mSolverMinLineSearchAlpha As Double ' Минимальный alpha, ниже которого line search считается неудачным.
Private mSolverMaxDeltaEpsilon0 As Double ' Защита от слишком большого шага по epsilon0.
Private mSolverMaxDeltaKappa As Double ' Защита от слишком большого шага по кривизнам.
Private mSolverSecantMaxRestarts As Long ' Сколько раз Secant может перестроить приближенную матрицу.
Private mSolverSecantMinStepNorm As Double ' Минимальная норма шага Secant до признания застоя.
Private mCapacityInitialLambda As Double ' Начальный шаг/точка поиска lambda для LoadMultiplier.
Private mCapacitySolutionStrategy As String ' Auto, LoadMultiplier или UltimateStrain.
Private mCapacitySearchMethod As String ' Bisection, Brent или Secant для LoadMultiplier.
Private mCapacityMaxLambda As Double ' Верхняя защитная граница lambda.
Private mCapacityToleranceLambda As Double ' Точность lambda в одномерном поиске LoadMultiplier.
Private mCapacityToleranceStrain As Double ' Точность попадания в предельную деформацию UltimateStrain.
Private mCapacityMaxRetries As Long ' Повторные попытки LoadMultiplier при несходимости пробной точки.
Private mCapacityBaseLoadSteps As Long ' Базовое число load steps для пробной точки capacity.
Private mCapacitySolverMaxIterations As Long ' Лимит итераций внутреннего solver-а при capacity-поиске.
Private mReport As CExecutionReport ' Необязательный человекочитаемый отчет выполнения; расчетная логика от него не зависит.

Private mStabilityCode As String ' SP63 или SP35: выбранная нормативная ветвь фильтра устойчивости.
Private mStabilityElementLength As Double ' Геометрическая длина элемента l, мм.
Private mStabilityMu1 As Double ' Коэффициент расчетной длины для первой главной плоскости.
Private mStabilityMu2 As Double ' Коэффициент расчетной длины для второй главной плоскости.
Private mStabilitySystemType As String ' Determinate или Indeterminate: правило сложения основного и случайного эксцентриситета.
Private mStabilityZeroSign1 As Double ' Направление случайного эксцентриситета при M1=0: +1 или -1.
Private mStabilityZeroSign2 As Double ' Направление случайного эксцентриситета при M2=0: +1 или -1.
Private mStabilityPhiLMode As String ' Auto или PhiL2.
Private mStabilityAccidentalEccentricityMode As String ' AutoWithL, AutoWithMuL или User для ea/ec,sl.
Private mStabilityAccidentalEccentricityPlanes As String ' OnlyMomentPlane или BothPlanes: в каких плоскостях учитывать случайный эксцентриситет.
Private mStabilityAccidentalEccentricityUser1 As Double ' Модуль пользовательского случайного эксцентриситета первой главной плоскости, мм.
Private mStabilityAccidentalEccentricityUser2 As Double ' Модуль пользовательского случайного эксцентриситета второй главной плоскости, мм.
Private mStabilitySP63Ks As Double ' Коэффициент ks для арматуры в D по СП 63.
Private mStabilitySP63DeltaEMin As Double ' Нижняя защита delta_e для численной устойчивости формулы.
Private mStabilitySP63DeltaEMax As Double ' Верхняя защита delta_e для численной устойчивости формулы.
Private mStabilitySP35PhiP As Double ' Коэффициент phi_p; для ненапрягаемой арматуры по умолчанию 1.
Private mStabilitySP35NOverNcrLimit As Double ' Ограничение N/Ncr для ветви СП 35.
Private mStabilitySP35Table721 As Variant ' Таблица 7.21 СП 35, прочитанная с Config.
Private mStabilityLoadCount As Long ' Число строк постоянных/длительных нагрузок для устойчивости.
Private mStabilityLoadCombinationID(1 To MAX_COMBINATIONS) As String ' LC, к которому относится строка длительных нагрузок.
Private mStabilityLoadN(1 To MAX_COMBINATIONS) As Double ' N от постоянных или постоянных+длительных нагрузок, Н.
Private mStabilityLoadUserMx(1 To MAX_COMBINATIONS) As Double ' Пользовательский Mx длительной части в точке приложения нагрузки.
Private mStabilityLoadUserMy(1 To MAX_COMBINATIONS) As Double ' Пользовательский My длительной части в точке приложения нагрузки.

' Инициализирует внутреннее состояние объекта значениями по умолчанию, безопасными для последующей настройки.
Private Sub Class_Initialize()
    Dim resultIndex As Long
    Set mStatusPolicy = New CBatchStatusPolicy
    Set mMomentZeroFilter = New CMomentZeroFilter
    mZeroMomentPerDepth = DEFAULT_ZERO_MOMENT_PER_DEPTH
    mMomentZeroFilter.Initialize mZeroMomentPerDepth
    For resultIndex = 1 To MAX_COMBINATIONS
        Set mResults(resultIndex) = New CCombinationResult
        mResults(resultIndex).Clear CAPACITY_PATH_NONE
    Next resultIndex

    mCrackAllowable = 0.3
    mCrackPsiMode = "User"
    mCrackPhi1 = 1.4
    mCrackPhi2 = 0.5
    mCrackPhi3Mode = "Auto"
    mCrackPhi3 = 1#
    mCrackPsiS = 1#
    mCrackTensionZoneMode = "Effective"
    mCrackCoverDistanceMode = "GlobalExtreme"
    mSolverMethod = "Newton"
    mSolverLoadSteps = 8
    mSolverMaxIterations = 80
    mSolverToleranceN = 5#
    mSolverToleranceMx = 5000#
    mSolverToleranceMy = 5000#
    mSolverLineSearchEnabled = True
    mSolverDampingInitial = 1#
    mSolverMinLineSearchAlpha = 0.03125
    mSolverMaxDeltaEpsilon0 = 0.0005
    mSolverMaxDeltaKappa = 0.00001
    mSolverSecantMaxRestarts = 2
    mSolverSecantMinStepNorm = 0.000000000001
    mCapacityInitialLambda = 1#
    mCapacitySolutionStrategy = "Auto"
    mCapacitySearchMethod = "Bisection"
    mCapacityMaxLambda = 64#
    mCapacityToleranceLambda = 0.01
    mCapacityToleranceStrain = 0.00001
    mCapacityMaxRetries = 0
    mCapacityBaseLoadSteps = 1
    mCapacitySolverMaxIterations = 60
    mStabilityCode = "SP63"
    mStabilityElementLength = 3000#
    mStabilityMu1 = 1#
    mStabilityMu2 = 1#
    mStabilitySystemType = "Determinate"
    mStabilityZeroSign1 = 1#
    mStabilityZeroSign2 = 1#
    mStabilityPhiLMode = "Auto"
    mStabilityAccidentalEccentricityMode = "AutoWithL"
    mStabilityAccidentalEccentricityPlanes = "OnlyMomentPlane"
    mStabilityAccidentalEccentricityUser1 = 10#
    mStabilityAccidentalEccentricityUser2 = 10#
    mStabilitySP63Ks = 0.7
    mStabilitySP63DeltaEMin = 0.15
    mStabilitySP63DeltaEMax = 1.5
    mStabilitySP35PhiP = 1#
    mStabilitySP35NOverNcrLimit = 0.7
End Sub

' Принимает готовую модель и материалы перед добавлением/расчетом LC.
' После этого класс не обращается к геометрическим генераторам: все сочетания
' считаются на одной и той же CSectionModel.
Public Sub Initialize(ByVal section As CSectionModel, ByVal materialProvider As CMaterialModelProvider)
    If section Is Nothing Then Err.Raise vbObjectError + 3900, "CBatchSectionCalculator", "Модель сечения не передана."
    If section.ConcreteCount <= 0 Then Err.Raise vbObjectError + 3901, "CBatchSectionCalculator", "В модели сечения нет бетонных элементов."
    If materialProvider Is Nothing Then Err.Raise vbObjectError + 3902, "CBatchSectionCalculator", "Material provider не передан."
    Set mSection = section
    Set mMaterialProvider = materialProvider
    ConfigureConcreteMomentDepths
End Sub

' Подключает каталог профилей, прочитанный Excel-слоем из rngCalculationProfiles.
' Batch дальше выбирает профиль только по ProfileId строки LC.
Public Property Set ProfileCatalog(ByVal profiles As CCalculationProfileCatalog)
    Set mProfiles = profiles
End Property

' Подключает внешний отчет выполнения. Batch только добавляет события, но не
' открывает и не сохраняет файл: жизненным циклом отчета управляет Excel-макрос.
Public Property Set ExecutionReport(ByVal report As CExecutionReport)
    Set mReport = report
End Property

' Считывает численные настройки solver-ов и трещин из уже загруженного reader-а.
' Если передан CUnitSystem, пользовательские допуски в силе/моменте/длине
' сразу переводятся во внутренние единицы ядра.
Public Sub ApplySettings(ByVal settings As CSystemSettingsReader, Optional ByVal units As CUnitSystem = Nothing)
    If settings Is Nothing Then Exit Sub
    Dim zeroMomentSettingRead As Boolean
    zeroMomentSettingRead = settings.HasKey("Calculation.ZeroMomentPerDepth")
    mCrackAllowable = settings.GetDouble("SLS.Crack.Allowable", mCrackAllowable)
    mCrackPsiMode = settings.GetString("SLS.Crack.PsiMode", mCrackPsiMode)
    mCrackTensionZoneMode = settings.GetString("SLS.Crack.TensionZoneMode", mCrackTensionZoneMode)
    If settings.HasKey("SLS.Crack.CoverDistanceMode") Then
        mCrackCoverDistanceMode = settings.GetString("SLS.Crack.CoverDistanceMode", mCrackCoverDistanceMode)
    Else
        mCrackCoverDistanceMode = settings.GetString("SLS.Crack.CoverMode", mCrackCoverDistanceMode)
    End If
    mCrackPhi1 = settings.GetDouble("SLS.Crack.Phi1", mCrackPhi1)
    mCrackPhi2 = settings.GetDouble("SLS.Crack.Phi2", mCrackPhi2)
    mCrackPhi3Mode = settings.GetString("SLS.Crack.Phi3Mode", mCrackPhi3Mode)
    mCrackPhi3 = settings.GetDouble("SLS.Crack.Phi3", mCrackPhi3)
    mCrackPsiS = settings.GetDouble("SLS.Crack.PsiS", mCrackPsiS)
    mZeroMomentPerDepth = settings.GetDouble("Calculation.ZeroMomentPerDepth", mZeroMomentPerDepth)
    If settings.HasKey("Solver.Method") Then
        mSolverMethod = settings.GetRawString("Solver.Method", vbNullString)
    Else
        mSolverMethod = "Newton"
    End If
    mSolverLoadSteps = settings.GetLong("Solver.LoadSteps", mSolverLoadSteps)
    mSolverMaxIterations = settings.GetLong("Solver.MaxIterations", mSolverMaxIterations)
    mSolverToleranceN = settings.GetDouble("Solver.ToleranceN", mSolverToleranceN)
    mSolverToleranceMx = settings.GetDouble("Solver.ToleranceMx", mSolverToleranceMx)
    mSolverToleranceMy = settings.GetDouble("Solver.ToleranceMy", mSolverToleranceMy)
    mSolverLineSearchEnabled = settings.GetBoolean("Solver.LineSearchEnabled", mSolverLineSearchEnabled)
    mSolverDampingInitial = settings.GetDouble("Solver.DampingInitial", mSolverDampingInitial)
    mSolverMinLineSearchAlpha = settings.GetDouble("Solver.MinLineSearchAlpha", mSolverMinLineSearchAlpha)
    mSolverMaxDeltaEpsilon0 = settings.GetDouble("Solver.MaxDeltaEpsilon0", mSolverMaxDeltaEpsilon0)
    mSolverMaxDeltaKappa = settings.GetDouble("Solver.MaxDeltaKappa", mSolverMaxDeltaKappa)
    mSolverSecantMaxRestarts = settings.GetLong("Solver.SecantMaxRestarts", mSolverSecantMaxRestarts)
    mSolverSecantMinStepNorm = settings.GetDouble("Solver.SecantMinStepNorm", mSolverSecantMinStepNorm)
    mCapacityInitialLambda = settings.GetDouble("Capacity.InitialLambda", mCapacityInitialLambda)
    mCapacitySolutionStrategy = settings.GetString("Capacity.SolutionStrategy", mCapacitySolutionStrategy)
    If settings.HasKey("Capacity.SearchMethod") Then
        mCapacitySearchMethod = settings.GetRawString("Capacity.SearchMethod", vbNullString)
    Else
        mCapacitySearchMethod = "Bisection"
    End If
    mCapacityMaxLambda = settings.GetDouble("Capacity.MaxLambda", mCapacityMaxLambda)
    mCapacityToleranceLambda = settings.GetDouble("Capacity.ToleranceLambda", mCapacityToleranceLambda)
    mCapacityToleranceStrain = settings.GetDouble("Capacity.ToleranceStrain", mCapacityToleranceStrain)
    mCapacityMaxRetries = settings.GetLong("Capacity.MaxRetries", mCapacityMaxRetries)
    mCapacityBaseLoadSteps = settings.GetLong("Capacity.BaseLoadSteps", mCapacityBaseLoadSteps)
    mCapacitySolverMaxIterations = settings.GetLong("Capacity.SolverMaxIterations", mCapacitySolverMaxIterations)
    mStabilityCode = settings.GetString("Stability.Code", mStabilityCode)
    mStabilityElementLength = StabilityNumericSetting(settings, "Stability.ElementLength", mStabilityElementLength)
    mStabilityMu1 = StabilityNumericSetting(settings, "Stability.Mu1", mStabilityMu1)
    mStabilityMu2 = StabilityNumericSetting(settings, "Stability.Mu2", mStabilityMu2)
    mStabilitySystemType = settings.GetString("Stability.SystemType", mStabilitySystemType)
    mStabilityZeroSign1 = settings.GetDouble("Stability.ZeroMomentEccentricitySign1", mStabilityZeroSign1)
    mStabilityZeroSign2 = settings.GetDouble("Stability.ZeroMomentEccentricitySign2", mStabilityZeroSign2)
    mStabilityPhiLMode = StabilityTextSetting(settings, "Stability.PhiLMode", mStabilityPhiLMode)
    mStabilityAccidentalEccentricityMode = StabilityTextSetting(settings, "Stability.AccidentalEccentricityMode", mStabilityAccidentalEccentricityMode)
    mStabilityAccidentalEccentricityPlanes = StabilityTextSetting(settings, "Stability.AccidentalEccentricityPlanes", mStabilityAccidentalEccentricityPlanes)
    mStabilityAccidentalEccentricityUser1 = StabilityNumericSetting(settings, "Stability.AccidentalEccentricityUser1", mStabilityAccidentalEccentricityUser1)
    mStabilityAccidentalEccentricityUser2 = StabilityNumericSetting(settings, "Stability.AccidentalEccentricityUser2", mStabilityAccidentalEccentricityUser2)
    mStabilitySP63Ks = StabilityNumericSetting(settings, "Stability.SP63.Ks", mStabilitySP63Ks)
    mStabilitySP63DeltaEMin = StabilityNumericSetting(settings, "Stability.SP63.DeltaEMin", mStabilitySP63DeltaEMin)
    mStabilitySP63DeltaEMax = StabilityNumericSetting(settings, "Stability.SP63.DeltaEMax", mStabilitySP63DeltaEMax)
    mStabilitySP35PhiP = StabilityNumericSetting(settings, "Stability.SP35.PhiP", mStabilitySP35PhiP)
    mStabilitySP35NOverNcrLimit = StabilityNumericSetting(settings, "Stability.SP35.NOverNcrLimit", mStabilitySP35NOverNcrLimit)
    If Not units Is Nothing Then
        mCrackAllowable = units.InputLengthToInternal(mCrackAllowable)
        mStabilityElementLength = units.InputLengthToInternal(mStabilityElementLength)
        mStabilityAccidentalEccentricityUser1 = units.InputLengthToInternal(mStabilityAccidentalEccentricityUser1)
        mStabilityAccidentalEccentricityUser2 = units.InputLengthToInternal(mStabilityAccidentalEccentricityUser2)
        mSolverToleranceN = Abs(units.InputForceToInternal(mSolverToleranceN))
        mSolverToleranceMx = Abs(units.InputMomentMxToInternal(mSolverToleranceMx))
        mSolverToleranceMy = Abs(units.InputMomentMyToInternal(mSolverToleranceMy))
        mSolverMaxDeltaKappa = Abs(units.InputCurvatureToInternal(mSolverMaxDeltaKappa))
        If zeroMomentSettingRead Then mZeroMomentPerDepth = units.InputMomentPerLengthToInternal(mZeroMomentPerDepth)
    End If
    ConfigureMomentZeroFilter
    NormalizeAllCombinationMoments
End Sub

' Читает числовую настройку устойчивости без подстановки default при пустой
' ячейке. Пустота сохраняется как 0, чтобы профиль с включенной устойчивостью
' получил InputErr в CStabilityCalculator, а не незаметный расчет по default.
Private Function StabilityNumericSetting(ByVal settings As CSystemSettingsReader, _
        ByVal key As String, ByVal defaultValue As Double) As Double
    If Not settings.HasKey(key) Then
        StabilityNumericSetting = defaultValue
        Exit Function
    End If

    Dim textValue As String
    textValue = Trim$(settings.GetRawString(key, vbNullString))
    If Len(textValue) = 0 Or UCase$(textValue) = "TODO" Then
        StabilityNumericSetting = 0#
    Else
        StabilityNumericSetting = CDbl(Replace$(textValue, ".", LocalDecimalSeparator()))
    End If
End Function

' Читает текстовый переключатель устойчивости без подмены пустого значения.
' Допустимые варианты проверяются внутри CStabilityCalculator.
Private Function StabilityTextSetting(ByVal settings As CSystemSettingsReader, _
        ByVal key As String, ByVal defaultValue As String) As String
    If settings.HasKey(key) Then
        StabilityTextSetting = Trim$(settings.GetRawString(key, vbNullString))
    Else
        StabilityTextSetting = defaultValue
    End If
End Function

' Локальный десятичный разделитель нужен только для raw-чтения настроек,
' когда обычный GetDouble нельзя использовать из-за его default-подстановки.
Private Function LocalDecimalSeparator() As String
    LocalDecimalSeparator = Mid$(CStr(1.1), 2, 1)
End Function

' Очищает список LC и накопленные результаты.
' Используется перед чтением новой таблицы сочетаний, чтобы старые статусы не
' попали в новый batch summary.
Public Sub ClearCombinations()
    mCount = 0
    ClearStabilityDurationLoads
    mGoverningCombinationIndex = 0
    mCrackGoverningCombinationIndex = 0
    mElapsedSeconds = 0#
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
    For i = 1 To MAX_COMBINATIONS
        mStabilityLoadCombinationID(i) = vbNullString
        mStabilityLoadN(i) = 0#
        mStabilityLoadUserMx(i) = 0#
        mStabilityLoadUserMy(i) = 0#
    Next i
End Sub

' Добавляет строку постоянных или постоянных+длительных нагрузок для LC.
' Какие именно длительные нагрузки вводит пользователь, зависит от выбранного
' нормативного блока и описано на листе "Справка"; ядро получает уже готовые
' значения во внутренних единицах.
Public Sub AddStabilityDurationLoad(ByVal combinationID As String, ByVal nValue As Double, _
        ByVal mxValue As Double, ByVal myValue As Double)
    If mStabilityLoadCount >= MAX_COMBINATIONS Then Err.Raise vbObjectError + 3925, _
        "CBatchSectionCalculator", "Поддерживается не более 20 строк нагрузок для устойчивости."
    mStabilityLoadCount = mStabilityLoadCount + 1
    mStabilityLoadCombinationID(mStabilityLoadCount) = Trim$(combinationID)
    mStabilityLoadN(mStabilityLoadCount) = nValue
    mStabilityLoadUserMx(mStabilityLoadCount) = mxValue
    mStabilityLoadUserMy(mStabilityLoadCount) = myValue
End Sub

' Принимает таблицу 7.21 СП 35 из Config. Таблица нужна только при выборе
' Stability.Code = SP35 и не должна быть зашита константами в код класса.
Public Sub SetSP35Table721(ByVal tableData As Variant)
    mStabilitySP35Table721 = tableData
End Sub

' Добавляет корректно прочитанное сочетание нагрузок.
' Усилия уже должны быть во внутренних единицах и в INTERNAL-системе знаков.
Public Sub AddCombination(ByVal combinationID As String, ByVal nValue As Double, ByVal mxValue As Double, _
        ByVal myValue As Double, ByVal profileId As String, ByVal comment As String, _
        Optional ByVal capacityLoadPath As String = vbNullString)
    AddCombinationInternal combinationID, nValue, mxValue, myValue, profileId, comment, True, vbNullString, capacityLoadPath
End Sub

' Сохраняет строку LC с ошибкой ввода, чтобы пользователь увидел ее в отчете.
' Такие сочетания не передаются в solver, но отображаются как InvalidInput.
Public Sub AddInvalidCombination(ByVal combinationID As String, ByVal profileId As String, _
        ByVal comment As String, ByVal reason As String, Optional ByVal capacityLoadPath As String = vbNullString)
    AddCombinationInternal combinationID, 0#, 0#, 0#, profileId, comment, False, reason, capacityLoadPath
End Sub

Private Sub AddCombinationInternal(ByVal combinationID As String, ByVal nValue As Double, ByVal mxValue As Double, _
        ByVal myValue As Double, ByVal profileId As String, ByVal comment As String, _
        ByVal inputValid As Boolean, ByVal inputError As String, ByVal capacityLoadPath As String)
    If mCount >= MAX_COMBINATIONS Then Err.Raise vbObjectError + 3911, "CBatchSectionCalculator", "Поддерживается не более 20 сочетаний нагрузок."
    mCount = mCount + 1
    If Len(Trim$(combinationID)) = 0 Then combinationID = "ROW" & CStr(mCount)
    mInputs(mCount).CombinationID = combinationID
    mInputs(mCount).N = nValue
    mInputs(mCount).Mx = mxValue
    mInputs(mCount).My = myValue
    mInputs(mCount).ProfileId = Trim$(profileId)
    mInputs(mCount).CapacityLoadPath = Trim$(capacityLoadPath)
    mInputs(mCount).Comment = comment
    mInputs(mCount).UserMx = mxValue
    mInputs(mCount).UserMy = myValue
    mInputs(mCount).InputValid = inputValid
    mInputs(mCount).InputError = inputError
    If inputValid Then NormalizeCombinationMoments mCount
End Sub

' Переносит пользовательские моменты из точки приложения нагрузки к координатам элементов.
' Пользователь задает Mx/My относительно выбранной точки нагрузки, а solver
' должен получить моменты относительно внутренней системы координат CSectionModel.
Public Sub ApplyLoadReference(ByVal referenceX As Double, ByVal referenceY As Double, _
        Optional ByVal baseX As Double = 0#, Optional ByVal baseY As Double = 0#)
    mLoadReferenceX = referenceX
    mLoadReferenceY = referenceY
    mLoadReferenceOffsetX = referenceX - baseX
    mLoadReferenceOffsetY = referenceY - baseY

    Dim i As Long
    ' Эквивалентный перенос момента от продольной силы:
    ' Mx_model = Mx_user + N * y_ref, My_model = My_user + N * x_ref.
    For i = 1 To mCount
        If mInputs(i).InputValid Then
            mInputs(i).Mx = mInputs(i).UserMx + mInputs(i).N * referenceY
            mInputs(i).My = mInputs(i).UserMy + mInputs(i).N * referenceX
            NormalizeCombinationMoments i
        End If
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

' Нормализует пару пользовательских моментов и заново переносит их к системе
' координат модели. Это единая точка для исходных LC и усилий после Stability.
Private Sub NormalizeCombinationMoments(ByVal index As Long)
    If index < 1 Or index > mCount Then Exit Sub
    If Not mInputs(index).InputValid Then Exit Sub

    mInputs(index).UserMx = NormalizeMomentMx(mInputs(index).UserMx)
    mInputs(index).UserMy = NormalizeMomentMy(mInputs(index).UserMy)
    mInputs(index).Mx = NormalizeMomentMx(mInputs(index).UserMx + mInputs(index).N * mLoadReferenceY)
    mInputs(index).My = NormalizeMomentMy(mInputs(index).UserMy + mInputs(index).N * mLoadReferenceX)
End Sub

' Применяет текущую настройку фильтра к уже загруженным сочетаниям. Метод
' нужен, если ApplySettings вызывается после чтения строк LC.
Private Sub NormalizeAllCombinationMoments()
    Dim i As Long
    For i = 1 To mCount
        NormalizeCombinationMoments i
    Next i
End Sub

' Последовательно считает все добавленные LC и выбирает определяющее сочетание.
' Порядок один для всех геометрий: capacity при необходимости, прямое НДС,
' затем трещины для сочетаний второй группы.
Public Sub Execute()
    On Error GoTo Failed

    ValidateReady
    ConfigureConcreteMomentDepths
    NormalizeAllCombinationMoments
    Dim t0 As Double
    Dim stage As String
    Dim currentIndex As Long

    stage = "Initialize"
    t0 = Timer
    mGoverningCombinationIndex = 0
    mCrackGoverningCombinationIndex = 0
    mDiagnosticLog = vbNullString
    ReportStep "Начат пакетный расчет. Количество сочетаний: " & CStr(mCount) & "."

    Dim i As Long
    ' Каждый LC изолирован: ошибка одного сочетания не должна стирать результаты
    ' соседних строк, но попадает в итоговый статус и batch summary.
    For i = 1 To mCount
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
    stage = "Finalize"
    mElapsedSeconds = Timer - t0
    ReportStep "Пакетный расчет завершен. Время: " & FormatNumberInvariant(mElapsedSeconds) & " с."
    Exit Sub

Failed:
    Err.Raise Err.Number, "CBatchSectionCalculator.Execute", _
        stage & " LC=" & CStr(currentIndex) & ": " & Err.Description
End Sub

' Считает одно сочетание нагрузок и собирает итоговый статус.
' Внутри сохраняется разделение: capacity отвечает за запас по несущей,
' прямой solve дает фактическое НДС, crack запускается только для 2 группы.
Private Sub RunCombination(ByVal index As Long)
    On Error GoTo Failed
    ReportSection "Сочетание " & CStr(index) & ": " & mInputs(index).CombinationID
    ReportValue "Комментарий", mInputs(index).Comment
    ReportValue "ProfileId", mInputs(index).ProfileId
    If Not mInputs(index).InputValid Then
        ReportStep "Сочетание не передано в solver из-за ошибки входных данных: " & mInputs(index).InputError
        MarkInvalid index, mInputs(index).InputError
        Exit Sub
    End If

    If Len(Trim$(mInputs(index).ProfileId)) = 0 Then
        ReportStep "Сочетание не рассчитано: ProfileId не задан."
        MarkInvalid index, "ProfileId должен быть назначен сочетанию."
        Exit Sub
    End If

    Dim profile As CCalculationProfile
    Set profile = mProfiles.ProfileById(mInputs(index).ProfileId)
    mResults(index).ProfileId = profile.ProfileId
    ReportValue "Профиль", profile.ProfileId & " / " & profile.DisplayName
    ReportValue "N", FormatNumberInvariant(mInputs(index).N) & " Н"
    ReportValue "Mx расчетный", FormatNumberInvariant(mInputs(index).Mx) & " Н*мм"
    ReportValue "My расчетный", FormatNumberInvariant(mInputs(index).My) & " Н*мм"

    Dim profileError As String
    If Not profile.ValidateForCalculation(NormalizedCrackPsiMode(), profileError) Then
        ReportStep "Сочетание не рассчитано: " & profileError
        MarkInvalid index, profileError
        Exit Sub
    End If

    mResults(index).Status = mStatusPolicy.NotApplicable
    mResults(index).DirectStateStatus = mStatusPolicy.NotApplicable
    mResults(index).CapacityStatus = mStatusPolicy.NotApplicable
    mResults(index).CrackStatus = mStatusPolicy.NotApplicable
    mResults(index).LongitudinalCrackStatus = mStatusPolicy.NotApplicable
    mResults(index).StabilityStatus = mStatusPolicy.NotApplicable

    If profile.StabilityEnabled Then
        ReportStep "Запускается фильтр продольного изгиба/устойчивости по профилю."
        RunStabilityFilter index, profile.StabilityValueSet
        If StabilityStopsDownstream(index) Then
            MarkDownstreamSkippedAfterStability index
            GoTo FinalizeCombination
        End If
    Else
        ReportStep "Устойчивость не запрошена профилем."
        mResults(index).ClearStabilityResult
        mResults(index).StabilityStatus = mStatusPolicy.NotApplicable
    End If

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

    If profile.CrackWidthEnabled Then
        RunCrackWidth index, profile
    End If

FinalizeCombination:
    mResults(index).GoverningUtil = MaxDouble(MaxDouble(MaxDouble(mResults(index).UtilCapacity, mResults(index).CrackUtil), _
        mResults(index).LongitudinalCrackUtil), mResults(index).StabilityUtil)
    mResults(index).Status = mStatusPolicy.Aggregate(mResults(index).DirectStateStatus, _
        mResults(index).CapacityStatus, mResults(index).CrackStatus, _
        mResults(index).LongitudinalCrackStatus, mResults(index).StabilityStatus)
    AppendDiagnostic index
    ReportCombinationSummary index
    Exit Sub

Failed:
    ReportStep "Ошибка при расчете сочетания: " & Err.Description
    MarkInvalid index, Err.Description
End Sub

' Применяет расчет продольного изгиба как изолированный фильтр усилий.
' CStabilityCalculator не знает о batch и профилях: он возвращает только
' расчетные N/Mx/My и диагностические величины, а batch решает, что делать
' дальше с пользовательскими статусами и последующими проверками НДМ.
Private Sub RunStabilityFilter(ByVal index As Long, ByVal valueSet As EMaterialValueSet)
    On Error GoTo Failed

    Dim referenceX As Double
    Dim referenceY As Double
    StabilityReferencePoint valueSet, referenceX, referenceY

    Dim sustainedN As Double
    Dim sustainedMx As Double
    Dim sustainedMy As Double
    StabilityDurationLoadForCombination mInputs(index).CombinationID, referenceX, referenceY, _
        sustainedN, sustainedMx, sustainedMy

    Dim stability As CStabilityCalculator
    Set stability = New CStabilityCalculator
    Dim referenceMx As Double
    Dim referenceMy As Double
    referenceMx = NormalizeMomentMx(MomentMxAboutPoint(index, referenceY))
    referenceMy = NormalizeMomentMy(MomentMyAboutPoint(index, referenceX))

    stability.Calculate mSection, mMaterialProvider, valueSet, mStabilityCode, _
        mInputs(index).N, referenceMx, referenceMy, _
        sustainedN, sustainedMx, sustainedMy, _
        mStabilityElementLength, mStabilityMu1, mStabilityMu2, _
        mStabilitySystemType, mStabilityZeroSign1, mStabilityZeroSign2, _
        mStabilityPhiLMode, mStabilityAccidentalEccentricityMode, _
        mStabilityAccidentalEccentricityPlanes, _
        mStabilityAccidentalEccentricityUser1, mStabilityAccidentalEccentricityUser2, _
        mStabilitySP63Ks, mStabilitySP63DeltaEMin, _
        mStabilitySP63DeltaEMax, mStabilitySP35PhiP, _
        mStabilitySP35NOverNcrLimit, mStabilitySP35Table721, mMomentZeroFilter

    mResults(index).StoreStabilityResult mStatusPolicy.StabilityStatusFromCalculator(stability), _
        stability, referenceX, referenceY
    If stability.InputValid And stability.Applicable And stability.CheckPassed Then _
        ApplyStabilityDesignLoads index, stability, referenceX, referenceY

    ReportValue "StabilityStatus", mResults(index).StabilityStatus
    ReportValue "Stability.Code", stability.Code
    ReportValue "Stability.ValueSet", MaterialValueSetToText(valueSet)
    ReportValue "Stability.Branch", stability.Branch
    ReportValue "Stability control force", FormatNumberInvariant(stability.CriticalForce) & " Н"
    ReportValue "Stability reserve", FormatNumberInvariant(stability.ReserveFactor)
    ReportValue "N/Mx/My после устойчивости", _
        FormatNumberInvariant(mInputs(index).N) & "; " & _
        FormatNumberInvariant(mInputs(index).Mx) & "; " & _
        FormatNumberInvariant(mInputs(index).My)
    Exit Sub

Failed:
    mResults(index).StabilityStatus = mStatusPolicy.InputErr
    mResults(index).StabilityBranch = Err.Description
    ReportStep "Устойчивость не рассчитана: " & Err.Description
End Sub

' Возвращает True, если фильтр устойчивости уже дал итог, после которого
' нельзя передавать усилия в НДМ. Так FAIL устойчивости не маскируется
' последующей численной несходимостью DirectState или Capacity.
Private Function StabilityStopsDownstream(ByVal index As Long) As Boolean
    StabilityStopsDownstream = mStatusPolicy.IsFailedOverall(mResults(index).StabilityStatus)
End Function

' Помечает последующие расчетные разделы как неприменимые к LC, который уже
' не прошел устойчивость. Пользовательский статус берется из общей policy:
' снаружи остается единый словарь OK/FAIL/NumFail/InputErr/N/A.
Private Sub MarkDownstreamSkippedAfterStability(ByVal index As Long)
    mResults(index).DirectStateStatus = mStatusPolicy.NotApplicable
    mResults(index).CapacityStatus = mStatusPolicy.NotApplicable
    mResults(index).CrackStatus = mStatusPolicy.NotApplicable
    mResults(index).LongitudinalCrackStatus = mStatusPolicy.NotApplicable
    ReportStep "НДС, несущая способность и трещины не запускались: сочетание не прошло проверку устойчивости."
End Sub

' Берет длительную часть нагрузки для LC. Если пользователь не заполнил строку,
' принимаются нули: это допустимый ввод и не блокирует основной расчет.
Private Sub StabilityDurationLoadForCombination(ByVal combinationID As String, _
        ByVal referenceX As Double, ByVal referenceY As Double, _
        ByRef nValue As Double, ByRef mxValue As Double, ByRef myValue As Double)
    Dim i As Long
    For i = 1 To mStabilityLoadCount
        If StrComp(mStabilityLoadCombinationID(i), Trim$(combinationID), vbTextCompare) = 0 Then
            nValue = mStabilityLoadN(i)
            mxValue = NormalizeMomentMx(mStabilityLoadUserMx(i) + mStabilityLoadN(i) * (mLoadReferenceY - referenceY))
            myValue = NormalizeMomentMy(mStabilityLoadUserMy(i) + mStabilityLoadN(i) * (mLoadReferenceX - referenceX))
            Exit Sub
        End If
    Next i
End Sub

' Подменяет усилия LC расчетными усилиями после фильтра устойчивости. Сам
' CStabilityCalculator работает с моментами относительно центра тяжести
' приведенного сечения, как требуют СП 35/63. Перед НДМ моменты снова
' переносятся к внутренней системе координат элементов.
Private Sub ApplyStabilityDesignLoads(ByVal index As Long, ByVal stability As CStabilityCalculator, _
        ByVal referenceX As Double, ByVal referenceY As Double)
    mInputs(index).N = stability.DesignN
    mInputs(index).Mx = stability.DesignMx + mInputs(index).N * referenceY
    mInputs(index).My = stability.DesignMy + mInputs(index).N * referenceX
    mInputs(index).UserMx = mInputs(index).Mx - mInputs(index).N * mLoadReferenceY
    mInputs(index).UserMy = mInputs(index).My - mInputs(index).N * mLoadReferenceX
    NormalizeCombinationMoments index

    ' Summary должен показывать те же моменты, которые после zero-filter ушли
    ' дальше в DirectState/Capacity/Crack, а не промежуточный DesignLoad.
    Dim finalDesignMx As Double
    Dim finalDesignMy As Double
    finalDesignMx = NormalizeMomentMx(MomentMxAboutPoint(index, referenceY))
    finalDesignMy = NormalizeMomentMy(MomentMyAboutPoint(index, referenceX))
    mResults(index).UpdateStabilityDesignLoads mInputs(index).N, finalDesignMx, finalDesignMy, _
        stability.PrincipalAngleRad
End Sub

' Возвращает Mx LC относительно произвольной расчетной точки.
' Для устойчивости сюда передается центр приведенного сечения, а для обычного
' НДМ пользовательская точка остается привязанной к бетонному центру.
Private Function MomentMxAboutPoint(ByVal index As Long, ByVal referenceY As Double) As Double
    MomentMxAboutPoint = mInputs(index).UserMx + mInputs(index).N * (mLoadReferenceY - referenceY)
End Function

' Возвращает My LC относительно произвольной расчетной точки.
Private Function MomentMyAboutPoint(ByVal index As Long, ByVal referenceX As Double) As Double
    MomentMyAboutPoint = mInputs(index).UserMy + mInputs(index).N * (mLoadReferenceX - referenceX)
End Function

' Возвращает центр приведенного сечения для нормативного расчета устойчивости.
' Приведение выполняется по модулю Eb/Es из того набора характеристик, который
' выбран в профиле строкой MaterialModel.Stability.ValueSet.
Private Sub StabilityReferencePoint(ByVal valueSet As EMaterialValueSet, _
        ByRef referenceX As Double, ByRef referenceY As Double)
    Dim props As CSectionPropertiesCalculator
    Set props = New CSectionPropertiesCalculator
    props.CalculateTransformedByModuli mSection, _
        mMaterialProvider.ConcreteElasticModulusFromValueSet(valueSet), _
        mMaterialProvider.SteelElasticModulusFromValueSet(valueSet)
    referenceX = props.CentroidX
    referenceY = props.CentroidY
End Sub

' Запускает поиск несущей способности для LC.
' Пользовательский столбец CapacityLoadPath задает, какие компоненты нагрузки
' умножаются на lambda. Все варианты дальше сводятся к одной математике:
' Target = Offset + lambda*Base для N, Mx и My.
Private Sub RunCapacity(ByVal index As Long, ByVal strengthSpec As CMaterialModelSpec)
    Dim loadPath As CCapacityLoadPath
    Dim cap As CCapacitySolver

    Set loadPath = BuildLoadPathForCombination(index)
    mResults(index).CapacityPathResolved = loadPath.Key
    ReportValue "CapacityLoadPath", IIf(loadPath.HasPath, loadPath.DisplayName, mInputs(index).CapacityLoadPath)

    If Not loadPath.Valid Then
        StoreInvalidCapacity index, loadPath.ErrorMessage
        ReportStep "Capacity остановлен: неизвестное значение CapacityLoadPath."
        Exit Sub
    End If

    If Not loadPath.HasPath Then
        ReportStep "Capacity не запускается: в сочетании нет ни момента, ни продольной силы."
        StoreSkippedCapacity index
        Exit Sub
    End If

    If Not loadPath.HasScaledLoad Then
        StoreInvalidCapacity index, "Для выбранного CapacityLoadPath соответствующая нагрузка должна быть ненулевой."
        ReportStep "Capacity остановлен: выбранная lambda-траектория не содержит ненулевой масштабируемой нагрузки."
        Exit Sub
    End If

    Set cap = New CCapacitySolver
    ConfigureCapacity cap, strengthSpec
    Select Case NormalizedCapacitySolutionStrategy()
        Case "Auto"
            If loadPath.ForceOnly Then
                cap.SolveByLoadPathMultiplier mSection, mMaterialProvider.ConcreteMaterialFromSpec(strengthSpec), _
                    mMaterialProvider.SteelMaterialFromSpec(strengthSpec), loadPath.NOffset, loadPath.NBase, _
                    loadPath.MxOffset, loadPath.MxBase, loadPath.MyOffset, loadPath.MyBase, True
            Else
                cap.SolveByAutoLoadPath mSection, mMaterialProvider.ConcreteMaterialFromSpec(strengthSpec), _
                    mMaterialProvider.SteelMaterialFromSpec(strengthSpec), loadPath.NOffset, loadPath.NBase, _
                    loadPath.MxOffset, loadPath.MxBase, loadPath.MyOffset, loadPath.MyBase
            End If
        Case "UltimateStrain"
            If loadPath.ForceOnly Then
                cap.SolveByLoadPathMultiplier mSection, mMaterialProvider.ConcreteMaterialFromSpec(strengthSpec), _
                    mMaterialProvider.SteelMaterialFromSpec(strengthSpec), loadPath.NOffset, loadPath.NBase, _
                    loadPath.MxOffset, loadPath.MxBase, loadPath.MyOffset, loadPath.MyBase, True
            Else
                cap.SolveByUltimateLoadPath mSection, mMaterialProvider.ConcreteMaterialFromSpec(strengthSpec), _
                    mMaterialProvider.SteelMaterialFromSpec(strengthSpec), loadPath.NOffset, loadPath.NBase, _
                    loadPath.MxOffset, loadPath.MxBase, loadPath.MyOffset, loadPath.MyBase
            End If
        Case "LoadMultiplier"
            cap.SolveByLoadPathMultiplier mSection, mMaterialProvider.ConcreteMaterialFromSpec(strengthSpec), _
                mMaterialProvider.SteelMaterialFromSpec(strengthSpec), loadPath.NOffset, loadPath.NBase, _
                loadPath.MxOffset, loadPath.MxBase, loadPath.MyOffset, loadPath.MyBase, loadPath.ForceOnly
        Case Else
            mResults(index).CapacityStatus = mStatusPolicy.InputErr
            mResults(index).CapacityLimitState = "Capacity.SolutionStrategy должен быть Auto, LoadMultiplier или UltimateStrain."
            mResults(index).CapacitySolutionMethod = vbNullString
            ReportStep "Capacity остановлен: неизвестный Capacity.SolutionStrategy."
            Exit Sub
    End Select

    StoreCapacity index, cap, loadPath, strengthSpec

    ReportValue "CapacityStatus", mResults(index).CapacityStatus
    ReportValue "LambdaCapacity", FormatNumberInvariant(mResults(index).LambdaCapacity)
    ReportValue "CapacityLoadPath", CapacityLoadPathDisplay(mResults(index).CapacityPathResolved)
    ReportBlock "Диагностика поиска несущей способности", cap.DiagnosticLog
    If Not cap.LastSolver Is Nothing Then
        ReportBlock "Итерации внутреннего CSectionSolver для последней capacity-точки", cap.LastSolver.DiagnosticLog
    End If
End Sub

' Выполняет один прямой StateSolution для заданного named-state профиля.
' Материальная модель строится по спецификации профиля, а управление retry и
' extension остается внутри CStateSolutionRunner.
Private Function RunProfileState(ByVal index As Long, ByVal stateType As ESectionStateType, _
        ByVal materialRole As ECalculationPurpose, ByVal spec As CMaterialModelSpec) As CStateSolutionRunner
    Dim concreteState As CMaterialDiagram
    Dim steelState As CMaterialDiagram
    Set concreteState = mMaterialProvider.ConcreteStateMaterialFromSpec(spec)
    Set steelState = mMaterialProvider.SteelStateMaterialFromSpec(spec)
    mResults(index).DirectStateDiagramExtensionEnabled = mMaterialProvider.DirectStateDiagramExtensionEnabled

    Dim runner As CStateSolutionRunner
    Set runner = New CStateSolutionRunner
    ConfigureStateRunner runner
    runner.Solve mSection, concreteState, steelState, mInputs(index).N, mInputs(index).Mx, mInputs(index).My, _
        mResults(index).DirectStateDiagramExtensionEnabled

    Dim service As CSectionSolver
    Set service = runner.ResultSolver
    If service Is Nothing Then
        MergeDirectStateStatus index, mStatusPolicy.NumFail
        ReportStep "Прямое НДС " & SectionStateTypeToText(stateType) & " не вернуло состояние solver-а."
        Set RunProfileState = runner
        Exit Function
    End If

    StoreSectionState index, service
    ReportValue "НДС " & SectionStateTypeToText(stateType) & " сошлось", IIf(service.Converged, "да", "нет")
    ReportValue "Итераций НДС", CStr(runner.Iterations)
    ReportBlock "Сценарий прямого StateSolution", runner.DiagnosticLog
    ReportBlock "Итерации прямого CSectionSolver", service.DiagnosticLog

    If Not service.Converged Then
        MergeDirectStateStatus index, mStatusPolicy.NumFail
        ReportStep "Прямое НДС " & SectionStateTypeToText(stateType) & " не сошлось: " & service.StopReason
        Set RunProfileState = runner
        Exit Function
    End If

    mResults(index).ExtensionUsed = (mResults(index).ExtensionUsed Or runner.ExtensionUsed)
    If runner.ExtensionUsed Or Not runner.WithinPhysicalRange Then
        MergeDirectStateStatus index, mStatusPolicy.Fail
        ReportStep "Прямое НДС " & SectionStateTypeToText(stateType) & " найдено, но вышло за физический диапазон диаграммы материала."
    Else
        MergeDirectStateStatus index, mStatusPolicy.OK
    End If
    ReportValue "DirectStateStatus", mResults(index).DirectStateStatus
    ReportValue "ExtensionUsed", IIf(runner.ExtensionUsed, "да", "нет")

    Dim stateResult As CSectionStateResult
    Set stateResult = New CSectionStateResult
    stateResult.InitializeFromSolver stateType, materialRole, spec, service, runner.ExtensionUsed, _
        mResults(index).DirectStateStatus
    mResults(index).AddState stateResult
    Set RunProfileState = runner
End Function

' Выполняет весь расчет раскрытия трещин как один запрошенный расчет профиля.
' Исходное CrackedState получает batch через общий StateSolutionRunner; сам
' CCrackWidthCalculator не повторяет этот solve и работает с готовым НДС.
Private Sub RunCrackWidth(ByVal index As Long, ByVal profile As CCalculationProfile)
    ReportStep "Запускается исходное НДС с раскрытой трещиной по профилю."

    Dim runner As CStateSolutionRunner
    Set runner = RunProfileState(index, sstCrackedState, cpCrackedNDS, profile.CrackedStateSpec)
    If runner Is Nothing Or runner.ResultSolver Is Nothing Then
        mResults(index).CrackStatus = mStatusPolicy.NotApplicable
        mResults(index).LongitudinalCrackStatus = mStatusPolicy.NotApplicable
        ReportStep "Расчет трещин не запускается: исходное CrackedState не получено."
        Exit Sub
    End If
    If Not runner.Converged Or runner.ExtensionUsed Or Not runner.WithinPhysicalRange Then
        mResults(index).CrackStatus = mStatusPolicy.NotApplicable
        mResults(index).LongitudinalCrackStatus = mStatusPolicy.NotApplicable
        ReportStep "Расчет трещин не запускается: исходное CrackedState не находится полностью в физическом диапазоне диаграммы."
        Exit Sub
    End If

    StoreLongitudinalCrackWithinCrackWidth index

    Dim crack As CCrackWidthCalculator
    Set crack = New CCrackWidthCalculator
    ConfigureCrackCalculator crack
    ReportStep "Запускается расчет ширины раскрытия нормальных трещин."

    Dim centroidMxForCrack As Double
    Dim centroidMyForCrack As Double
    ' Для отдельной ветки центрального растяжения важны моменты относительно
    ' центра тяжести бетонного сечения. Если точка приложения N смещена от
    ' бетонного центра, то даже при Mx_user = My_user = 0 появляется
    ' внецентренное растяжение.
    centroidMxForCrack = mInputs(index).UserMx + mInputs(index).N * mLoadReferenceOffsetY
    centroidMyForCrack = mInputs(index).UserMy + mInputs(index).N * mLoadReferenceOffsetX
    centroidMxForCrack = NormalizeMomentMx(centroidMxForCrack)
    centroidMyForCrack = NormalizeMomentMy(centroidMyForCrack)

    crack.Calculate runner.ResultSolver, mSection, mMaterialProvider, profile.CrackedStateSpec, _
        profile.CrackInitiationSpec, mInputs(index).N, mInputs(index).Mx, mInputs(index).My, _
        centroidMxForCrack, centroidMyForCrack
    mResults(index).CrackStatus = mStatusPolicy.CrackStatusFromCalculator(crack)
    mResults(index).StoreCrackResult crack
    If Not crack.BeforeMcrcState Is Nothing Then mResults(index).AddState crack.BeforeMcrcState
    If Not crack.AfterMcrcState Is Nothing Then mResults(index).AddState crack.AfterMcrcState

    ReportValue "CrackStatus", mResults(index).CrackStatus
    ReportValue "Трещина образована", IIf(mResults(index).CrackFormed, "да", "нет")
    ReportValue "a_crc", FormatNumberInvariant(mResults(index).CrackWidth) & " мм"
    ReportValue "psi_s", FormatNumberInvariant(mResults(index).CrackPsiS)
    ReportValue "As", FormatNumberInvariant(mResults(index).CrackAs) & " мм2"
    ReportValue "Abt", FormatNumberInvariant(mResults(index).CrackAbt) & " мм2"
    ReportValue "ds_eq", FormatNumberInvariant(mResults(index).CrackDsEquivalent) & " мм"
    ReportValue "ls", FormatNumberInvariant(mResults(index).CrackLs) & " мм"
    ReportBlock "Диагностика расчета трещин", crack.DiagnosticLog
End Sub

' Создает объект lambda-траектории для одной строки LC.
' Batch передает сюда только уже нормализованные усилия и точку приложения
' нагрузки, а все подробности интерпретации lambda*Mx/lambda*N остаются в CCapacityLoadPath.
Private Function BuildLoadPathForCombination(ByVal index As Long) As CCapacityLoadPath
    Dim loadPath As CCapacityLoadPath
    Set loadPath = New CCapacityLoadPath
    loadPath.Initialize mInputs(index).CapacityLoadPath, mInputs(index).N, mInputs(index).UserMx, mInputs(index).UserMy, _
        mLoadReferenceX, mLoadReferenceY, mMomentZeroFilter, mConcreteDepthForMx, mConcreteDepthForMy
    Set BuildLoadPathForCombination = loadPath
End Function

' Формирует человекочитаемую подпись load-path для Results и отчета.
' Само соответствие ключа и подписи хранится в CCapacityLoadPath, чтобы batch
' не становился вторым местом интерпретации lambda-траекторий.
Private Function CapacityLoadPathDisplay(ByVal loadPathKey As String) As String
    Dim loadPath As CCapacityLoadPath
    Set loadPath = New CCapacityLoadPath
    CapacityLoadPathDisplay = loadPath.DisplayNameForKey(loadPathKey)
End Function

' Приводит пользовательскую строку Capacity.SolutionStrategy к одному из трех
' внутренних имен. Пустой результат дальше превращается в InvalidInput.
Private Function NormalizedCapacitySolutionStrategy() As String
    Select Case LCase$(Trim$(mCapacitySolutionStrategy))
        Case "auto"
            NormalizedCapacitySolutionStrategy = "Auto"
        Case "loadmultiplier"
            NormalizedCapacitySolutionStrategy = "LoadMultiplier"
        Case "ultimatestrain"
            NormalizedCapacitySolutionStrategy = "UltimateStrain"
    End Select
End Function

' Нормализует режим psi_s для проверки обязательности MaterialModel.CrackInitiation.
Private Function NormalizedCrackPsiMode() As String
    Select Case LCase$(Trim$(mCrackPsiMode))
        Case "auto", "refined"
            NormalizedCrackPsiMode = "Auto"
        Case "user", "manual", "fixed1", "unity"
            NormalizedCrackPsiMode = "User"
        Case Else
            NormalizedCrackPsiMode = "User"
    End Select
End Function

' Объединяет статусы нескольких прямых StateSolution одного LC.
' N/A не ухудшает итог, а FAIL/NumFail/InputErr сохраняют общий смысл строки.
Private Sub MergeDirectStateStatus(ByVal index As Long, ByVal statusText As String)
    mResults(index).DirectStateStatus = mStatusPolicy.Aggregate(mResults(index).DirectStateStatus, _
        statusText, mStatusPolicy.NotApplicable)
End Sub

' Заполняет capacity-поля нейтральным статусом, когда эта проверка не должна
' выполняться для выбранной группы/режима.
Private Sub StoreNotCalculatedCapacity(ByVal index As Long)
    mResults(index).CapacityStatus = mStatusPolicy.NotApplicable
    mResults(index).CapacityLimitState = mStatusPolicy.NotApplicable
    mResults(index).CapacitySolutionMethod = vbNullString
    mResults(index).CapacityPathResolved = CAPACITY_PATH_NONE
End Sub

' Помечает capacity как пропущенную из-за нулевой нагрузки.
' Это не ошибка исходных данных: для нулевых N, Mx и My нельзя построить
' направление внешнего поиска несущей способности.
Private Sub StoreSkippedCapacity(ByVal index As Long)
    mResults(index).CapacityStatus = mStatusPolicy.NotApplicable
    mResults(index).CapacityLimitState = "ZeroBaseLoad"
    mResults(index).CapacitySolutionMethod = vbNullString
    mResults(index).CapacityPathResolved = CAPACITY_PATH_NONE
End Sub

' Помечает только ветку capacity как ошибку исходных данных.
' Прямое НДС при этом все равно может быть посчитано: ошибка относится к
' выбору траектории несущей способности, а не к самим нагрузкам LC.
Private Sub StoreInvalidCapacity(ByVal index As Long, ByVal reason As String)
    mResults(index).CapacityStatus = mStatusPolicy.InputErr
    mResults(index).CapacityLimitState = reason
    mResults(index).CapacitySolutionMethod = vbNullString
    mResults(index).CapacityPathResolved = CAPACITY_PATH_NONE
    mResults(index).LambdaCapacity = 0#
    mResults(index).UtilCapacity = 0#
    mResults(index).NUltimate = 0#
    mResults(index).MxUltimate = 0#
    mResults(index).MyUltimate = 0#
    mResults(index).MomentUltimate = 0#
End Sub

' Переносит результат CCapacitySolver в массивы batch-отчета.
' Здесь разделяется математическая сходимость capacity solver-а и простая
' инженерная проверка: lambda меньше или больше единицы.
Private Sub StoreCapacity(ByVal index As Long, ByVal cap As CCapacitySolver, _
        ByVal loadPath As CCapacityLoadPath, ByVal strengthSpec As CMaterialModelSpec)
    mResults(index).CapacityStatus = mStatusPolicy.CapacityStatusFromSolver(cap)
    mResults(index).CapacityLimitState = cap.LimitState
    mResults(index).CapacitySolutionMethod = cap.SolutionMethod
    mResults(index).CapacityPathResolved = loadPath.Key
    mResults(index).LambdaCapacity = cap.LambdaUltimate
    If cap.LambdaUltimate > 0# Then
        If loadPath.ScalesN Then mResults(index).NUltimate = cap.LambdaUltimate * mInputs(index).N
        If loadPath.ScalesMx Then mResults(index).MxUltimate = cap.LambdaUltimate * mInputs(index).UserMx
        If loadPath.ScalesMy Then mResults(index).MyUltimate = cap.LambdaUltimate * mInputs(index).UserMy
    End If
    mResults(index).MomentUltimate = Sqr(mResults(index).MxUltimate * mResults(index).MxUltimate + mResults(index).MyUltimate * mResults(index).MyUltimate)
    If cap.LambdaUltimate > 0# Then mResults(index).UtilCapacity = 1# / cap.LambdaUltimate
    If cap.Converged Then mResults(index).UtilCapacity = cap.Utilization

    ' h/x в блоке предельного момента должны относиться к найденному
    ' предельному состоянию, а не к фактическому НДС LC. Поэтому сохраняем
    ' отдельную плоскость деформаций только когда capacity-поиск действительно
    ' дал физически принятое предельное состояние.
    If cap.Converged And cap.LambdaUltimate > 0# Then
        If Not cap.LastSolver Is Nothing Then
            mResults(index).CapacityStateAvailable = True
            mResults(index).CapacityEpsilon0 = cap.LastSolver.Epsilon0
            mResults(index).CapacityKappaX = cap.LastSolver.KappaX
            mResults(index).CapacityKappaY = cap.LastSolver.KappaY
            Dim stateResult As CSectionStateResult
            Set stateResult = New CSectionStateResult
            stateResult.InitializeFromSolver sstCapacityState, cpStrength, strengthSpec, cap.LastSolver, False, _
                mResults(index).CapacityStatus
            mResults(index).AddState stateResult
        End If
    End If
End Sub

' Фиксирует InvalidInput по конкретному LC и не дает этому сочетанию пройти
' в solver. Так пользователь видит причину в таблице, а остальные LC считаются.
Private Sub MarkInvalid(ByVal index As Long, ByVal reason As String)
    If Len(Trim$(reason)) = 0 Then reason = "Некорректное сочетание нагрузок."
    mInputs(index).InputError = reason
    mResults(index).ProfileId = mInputs(index).ProfileId
    mResults(index).Status = mStatusPolicy.InputErr
    mResults(index).DirectStateStatus = mStatusPolicy.InputErr
    mResults(index).CapacityStatus = mStatusPolicy.InputErr
    mResults(index).CapacityLimitState = reason
    mResults(index).CapacitySolutionMethod = vbNullString
    mResults(index).CapacityPathResolved = CAPACITY_PATH_NONE
    mResults(index).CrackStatus = mStatusPolicy.NotApplicable
    mResults(index).LongitudinalCrackStatus = mStatusPolicy.NotApplicable
    If Not mMaterialProvider Is Nothing Then
        mResults(index).DirectStateDiagramExtensionEnabled = mMaterialProvider.DirectStateDiagramExtensionEnabled
    End If
    mResults(index).GoverningUtil = 0#
    AppendDiagnostic index
End Sub

' Выполняет SLS(II)-проверку продольных трещин по максимальному сжимающему
' напряжению бетона из уже найденного прямого НДС с трещинами. По СП 35 эта
' проверка относится к трещиностойкости, поэтому в новой профильной схеме она
' является частью Calculation.Crack.Width. Повторный расчет состояния здесь
' намеренно не нужен: сравниваем именно Stress/Strain рассчитанного LC.
Private Sub StoreLongitudinalCrackWithinCrackWidth(ByVal index As Long)
    If Not mResults(index).StateAvailable Or Not mResults(index).StateConverged Then
        mResults(index).LongitudinalCrackStatus = mStatusPolicy.NotApplicable
        ReportStep "Проверка продольных трещин не запускается: нет сходящегося прямого НДС."
        Exit Sub
    End If

    Dim maxCompressionStress As Double
    Dim rbMc2 As Double
    Dim utilization As Double
    maxCompressionStress = mResults(index).MaxConcreteCompressionStress
    rbMc2 = mMaterialProvider.ConcreteParameters.RbMc2

    If maxCompressionStress > 0.000000001 And rbMc2 > 0# Then
        utilization = maxCompressionStress / rbMc2
    End If

    mResults(index).StoreLongitudinalCrackResult _
        mStatusPolicy.LongitudinalCrackStatus(maxCompressionStress, rbMc2), _
        maxCompressionStress, rbMc2, utilization

    ReportValue "LongitudinalCrackStatus", mResults(index).LongitudinalCrackStatus
    If mResults(index).LongitudinalCrackStatus <> mStatusPolicy.NotApplicable Then
        ReportValue "sigma_c,max", FormatNumberInvariant(maxCompressionStress) & " МПа"
        ReportValue "Rb,mc2", FormatNumberInvariant(rbMc2) & " МПа"
        ReportValue "LongitudinalCrackUtilization", FormatNumberInvariant(utilization)
    Else
        ReportStep "Проверка продольных трещин неприменима: в прямом НДС нет сжатого бетона."
    End If
End Sub

' Передает в CCapacitySolver ровно те настройки, которые относятся к поиску
' несущей способности и внутренним пробным расчетам равновесия.
Private Sub ConfigureCapacity(ByVal cap As CCapacitySolver, ByVal strengthSpec As CMaterialModelSpec)
    cap.ConcreteCompressionLimit = mMaterialProvider.ConcreteCompressionLimitFromSpec(strengthSpec)
    cap.ConcreteTensionLimit = mMaterialProvider.ConcreteTensionLimitFromSpec(strengthSpec)
    cap.ConcreteTensionLimitEnabled = mMaterialProvider.ConcreteTensionLimitEnabledFromSpec(strengthSpec)
    cap.SteelStrainLimit = MaxDouble(Abs(mMaterialProvider.SteelCompressionLimitFromSpec(strengthSpec)), _
        Abs(mMaterialProvider.SteelTensionLimitFromSpec(strengthSpec)))
    cap.LambdaTolerance = mCapacityToleranceLambda
    cap.StrainTolerance = mCapacityToleranceStrain
    cap.SearchMethod = mCapacitySearchMethod
    cap.MaxLambda = mCapacityMaxLambda
    cap.InitialLambdaStep = mCapacityInitialLambda
    cap.MaxRetries = mCapacityMaxRetries
    cap.SolverBaseLoadSteps = mCapacityBaseLoadSteps
    cap.SolverMaxIterations = mCapacitySolverMaxIterations
    cap.SolverToleranceN = mSolverToleranceN
    cap.SolverToleranceMx = mSolverToleranceMx
    cap.SolverToleranceMy = mSolverToleranceMy
    cap.SolverMethod = mSolverMethod
    cap.SolverLineSearchEnabled = mSolverLineSearchEnabled
    cap.SolverDampingInitial = mSolverDampingInitial
    cap.SolverMinLineSearchAlpha = mSolverMinLineSearchAlpha
    cap.SolverMaxDeltaEpsilon0 = mSolverMaxDeltaEpsilon0
    cap.SolverMaxDeltaKappa = mSolverMaxDeltaKappa
    cap.SolverSecantMaxRestarts = mSolverSecantMaxRestarts
    cap.SolverSecantMinStepNorm = mSolverSecantMinStepNorm
    cap.DiagnosticsEnabled = DetailedDiagnosticsEnabled()
End Sub

' Настраивает runner прямого НДС общими solver-параметрами batch-расчета.
' Здесь остается только передача настроек; сама логика initial guess, retry
' и extension живет в CStateSolutionRunner.
Private Sub ConfigureStateRunner(ByVal runner As CStateSolutionRunner)
    runner.LoadSteps = mSolverLoadSteps
    runner.SolverMethod = mSolverMethod
    runner.MaxIterations = mSolverMaxIterations
    runner.ToleranceN = mSolverToleranceN
    runner.ToleranceMx = mSolverToleranceMx
    runner.ToleranceMy = mSolverToleranceMy
    runner.LineSearchEnabled = mSolverLineSearchEnabled
    runner.DampingInitial = mSolverDampingInitial
    runner.MinLineSearchAlpha = mSolverMinLineSearchAlpha
    runner.MaxDeltaEpsilon0 = mSolverMaxDeltaEpsilon0
    runner.MaxDeltaKappa = mSolverMaxDeltaKappa
    runner.SecantMaxRestarts = mSolverSecantMaxRestarts
    runner.SecantMinStepNorm = mSolverSecantMinStepNorm
    runner.DiagnosticsEnabled = DetailedDiagnosticsEnabled()
End Sub

' Передает в расчет трещин те же настройки прямого равновесия, что и в
' сервисный solve LC. Это важно для сопоставимости: состояние с раскрытой
' трещиной ищется тем же CSectionSolver, только с принудительно выключенным
' растянутым бетоном.
Private Sub ConfigureCrackCalculator(ByVal crack As CCrackWidthCalculator)
    crack.AllowableCrackWidth = mCrackAllowable
    crack.PsiMode = mCrackPsiMode
    crack.Phi1 = mCrackPhi1
    crack.Phi2 = mCrackPhi2
    crack.Phi3Mode = mCrackPhi3Mode
    crack.Phi3 = mCrackPhi3
    crack.PsiS = mCrackPsiS
    crack.TensionZoneMode = mCrackTensionZoneMode
    crack.CoverMode = mCrackCoverDistanceMode
    crack.SolverLoadSteps = mSolverLoadSteps
    crack.SolverMethod = mSolverMethod
    crack.SolverMaxIterations = mSolverMaxIterations
    crack.SolverToleranceN = mSolverToleranceN
    crack.SolverToleranceMx = mSolverToleranceMx
    crack.SolverToleranceMy = mSolverToleranceMy
    crack.SolverLineSearchEnabled = mSolverLineSearchEnabled
    crack.SolverDampingInitial = mSolverDampingInitial
    crack.SolverMinLineSearchAlpha = mSolverMinLineSearchAlpha
    crack.SolverMaxDeltaEpsilon0 = mSolverMaxDeltaEpsilon0
    crack.SolverMaxDeltaKappa = mSolverMaxDeltaKappa
    crack.SolverSecantMaxRestarts = mSolverSecantMaxRestarts
    crack.SolverSecantMinStepNorm = mSolverSecantMinStepNorm
    crack.DiagnosticsEnabled = DetailedDiagnosticsEnabled()
End Sub

' Проверяет, нужен ли подробный журнал внутренних итераций.
' Короткая batch-сводка всегда остается в памяти для тестов и диагностики, но
' тяжелые строки CSectionSolver/CCapacitySolver/CCrackWidthCalculator имеют
' смысл только тогда, когда пользователь включил txt-отчет выполнения.
Private Function DetailedDiagnosticsEnabled() As Boolean
    If mReport Is Nothing Then Exit Function
    DetailedDiagnosticsEnabled = mReport.Enabled
End Function

' Проверяет, что batch готов к расчету.
' Это последний защитный барьер перед запуском solver-ов: модель и хотя бы
' одно сочетание должны быть уже загружены.
Private Sub ValidateReady()
    If mSection Is Nothing Then Err.Raise vbObjectError + 3920, "CBatchSectionCalculator", "Пакетный расчетчик не инициализирован моделью сечения."
    If mMaterialProvider Is Nothing Then Err.Raise vbObjectError + 3922, "CBatchSectionCalculator", "Пакетный расчетчик не инициализирован material provider-ом."
    If mProfiles Is Nothing Then Err.Raise vbObjectError + 3923, "CBatchSectionCalculator", "Каталог расчетных профилей не загружен."
    If mCount <= 0 Then Err.Raise vbObjectError + 3921, "CBatchSectionCalculator", "Не задано ни одного сочетания нагрузок."
End Sub

' Сбрасывает все результаты одной строки перед повторным расчетом LC.
' Входные усилия и комментарий не трогаются.
Private Sub ClearResultRow(ByVal index As Long)
    If mResults(index) Is Nothing Then Set mResults(index) = New CCombinationResult
    mResults(index).Clear CAPACITY_PATH_NONE
    mResults(index).ProfileId = mInputs(index).ProfileId
End Sub

' Сохраняет найденную плоскость деформаций и внутренние усилия прямого solve.
' Эти данные затем уходят в Results и используются для схемы/AutoCAD export.
Private Sub StoreSectionState(ByVal index As Long, ByVal solver As CSectionSolver)
    mResults(index).StoreSectionState solver
End Sub

' Добавляет одну строку в текстовую диагностику batch-расчета.
' Это компактная диагностическая сводка, а не подробный журнал итераций solver-а.
Private Sub AppendDiagnostic(ByVal index As Long)
    mDiagnosticLog = mDiagnosticLog & "combination=" & mInputs(index).CombinationID & _
        "; OverallStatus=" & mResults(index).Status & _
        "; DirectStateStatus=" & mResults(index).DirectStateStatus & _
        "; ExtensionUsed=" & CStr(mResults(index).ExtensionUsed) & _
        "; CapacityStatus=" & mResults(index).CapacityStatus & _
        "; CrackStatus=" & mResults(index).CrackStatus & _
        "; LongitudinalCrackStatus=" & mResults(index).LongitudinalCrackStatus & _
        "; LongitudinalCrackUtil=" & FormatNumberInvariant(mResults(index).LongitudinalCrackUtil) & _
        "; StabilityStatus=" & mResults(index).StabilityStatus & _
        "; StabilityUtil=" & FormatNumberInvariant(mResults(index).StabilityUtil) & _
        "; governingUtil=" & FormatNumberInvariant(mResults(index).GoverningUtil) & vbCrLf
End Sub

' Добавляет в подробный txt-отчет краткий итог по текущему LC.
' Batch summary остается основным табличным выводом, а здесь пользователь
' видит ход выполнения без открытия Results и без чтения внутренних массивов.
Private Sub ReportCombinationSummary(ByVal index As Long)
    ReportValue "Итоговый Status", mResults(index).Status
    ReportValue "DirectStateStatus", mResults(index).DirectStateStatus
    ReportValue "ExtensionUsed", CStr(mResults(index).ExtensionUsed)
    ReportValue "CapacityStatus", mResults(index).CapacityStatus
    ReportValue "CrackStatus", mResults(index).CrackStatus
    ReportValue "LongitudinalCrackStatus", mResults(index).LongitudinalCrackStatus
    ReportValue "LongitudinalCrackUtilization", FormatNumberInvariant(mResults(index).LongitudinalCrackUtil)
    ReportValue "StabilityStatus", mResults(index).StabilityStatus
    ReportValue "StabilityUtilization", FormatNumberInvariant(mResults(index).StabilityUtil)
    ReportValue "GoverningUtilization", FormatNumberInvariant(mResults(index).GoverningUtil)
End Sub

Private Sub ReportSection(ByVal title As String)
    If mReport Is Nothing Then Exit Sub
    If Not mReport.Enabled Then Exit Sub
    mReport.AddSection title
End Sub

Private Sub ReportStep(ByVal text As String)
    If mReport Is Nothing Then Exit Sub
    If Not mReport.Enabled Then Exit Sub
    mReport.AddStep text
End Sub

Private Sub ReportValue(ByVal label As String, ByVal valueText As String)
    If mReport Is Nothing Then Exit Sub
    If Not mReport.Enabled Then Exit Sub
    mReport.AddValue label, valueText
End Sub

Private Sub ReportBlock(ByVal title As String, ByVal body As String)
    If mReport Is Nothing Then Exit Sub
    If Not mReport.Enabled Then Exit Sub
    mReport.AddBlock title, body
End Sub

' Утилитарный максимум без WorksheetFunction, чтобы не тащить Excel в hot path.
Private Function MaxDouble(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then MaxDouble = a Else MaxDouble = b
End Function

Private Function MinDouble(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then MinDouble = a Else MinDouble = b
End Function

' Выбирает худшее сочетание для batch summary и режима Worst.
' По пользовательскому требованию Worst относится именно к прочности: после
' отказа от пользовательского StrainSafetyFactor выбор идет по запасу
' найденного предельного момента, а трещины остаются отдельной проверкой.
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
' После перехода на CapacityLoadPath это единая величина для lambda*Mx,
' lambda*My, lambda*Mxy, lambda*N и lambda*NMxy; отдельные отношения Nult/N или Mult/Mxy больше
' не нужны для выбора Worst.
Private Function StrengthSafetyFactorForGoverning(ByVal index As Long) As Double
    If mStatusPolicy.IsFinished(mResults(index).CapacityStatus) And mResults(index).LambdaCapacity > 0# Then
        StrengthSafetyFactorForGoverning = mResults(index).LambdaCapacity
    End If
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
    If mResults(index).CrackUtil > 0# Then CrackSafetyFactorForGoverning = 1# / mResults(index).CrackUtil
End Function

' Берет минимальное положительное значение из двух запасов.
' Нулевые значения означают, что соответствующая проверка не рассчитана.
Private Function PositiveMin(ByVal firstValue As Double, ByVal secondValue As Double) As Double
    If firstValue > 0# And secondValue > 0# Then
        PositiveMin = MinDouble(firstValue, secondValue)
    ElseIf firstValue > 0# Then
        PositiveMin = firstValue
    ElseIf secondValue > 0# Then
        PositiveMin = secondValue
    End If
End Function

' Форматирует число с точкой как десятичным разделителем для диагностической строки.
Private Function FormatNumberInvariant(ByVal value As Double) As String
    FormatNumberInvariant = Replace$(Format$(value, "0.############"), ",", ".")
End Function

Public Property Get Count() As Long
    Count = mCount
End Property

Public Property Get CombinationID(ByVal index As Long) As String
    CombinationID = mInputs(index).CombinationID
End Property

Public Property Get CombinationName(ByVal index As Long) As String
    CombinationName = mInputs(index).Comment
End Property

Public Property Get ProfileId(ByVal index As Long) As String
    ProfileId = mInputs(index).ProfileId
End Property

Public Property Get CapacityLoadPath(ByVal index As Long) As String
    CapacityLoadPath = CapacityLoadPathDisplay(CapacityLoadPathKey(index))
End Property

Public Property Get CapacityLoadPathKey(ByVal index As Long) As String
    If Len(mResults(index).CapacityPathResolved) > 0 And mResults(index).CapacityPathResolved <> CAPACITY_PATH_NONE Then
        CapacityLoadPathKey = mResults(index).CapacityPathResolved
    Else
        CapacityLoadPathKey = BuildLoadPathForCombination(index).Key
    End If
End Property

Public Property Get N(ByVal index As Long) As Double
    N = mInputs(index).N
End Property

Public Property Get Mx(ByVal index As Long) As Double
    Mx = mInputs(index).Mx
End Property

Public Property Get My(ByVal index As Long) As Double
    My = mInputs(index).My
End Property

Public Property Get UserMx(ByVal index As Long) As Double
    UserMx = mInputs(index).UserMx
End Property

Public Property Get UserMy(ByVal index As Long) As Double
    UserMy = mInputs(index).UserMy
End Property

Public Property Get LoadReferenceX() As Double
    LoadReferenceX = mLoadReferenceX
End Property

Public Property Get LoadReferenceY() As Double
    LoadReferenceY = mLoadReferenceY
End Property

Public Property Get LoadReferenceOffsetX() As Double
    LoadReferenceOffsetX = mLoadReferenceOffsetX
End Property

Public Property Get LoadReferenceOffsetY() As Double
    LoadReferenceOffsetY = mLoadReferenceOffsetY
End Property

Public Property Get StabilityElementLength() As Double
    StabilityElementLength = mStabilityElementLength
End Property

Public Property Get StabilityMu1() As Double
    StabilityMu1 = mStabilityMu1
End Property

Public Property Get StabilityMu2() As Double
    StabilityMu2 = mStabilityMu2
End Property

Public Property Get Status(ByVal index As Long) As String
    Status = mResults(index).Status
End Property

Public Property Get OverallStatus(ByVal index As Long) As String
    OverallStatus = mResults(index).Status
End Property

Public Property Get DirectStateStatus(ByVal index As Long) As String
    DirectStateStatus = mResults(index).DirectStateStatus
End Property

Public Property Get CapacityStatus(ByVal index As Long) As String
    CapacityStatus = mResults(index).CapacityStatus
End Property

Public Property Get CapacityLimitState(ByVal index As Long) As String
    CapacityLimitState = mResults(index).CapacityLimitState
End Property

Public Property Get CapacitySolutionMethod(ByVal index As Long) As String
    CapacitySolutionMethod = mResults(index).CapacitySolutionMethod
End Property

Public Property Get CrackStatus(ByVal index As Long) As String
    CrackStatus = mResults(index).CrackStatus
End Property

Public Property Get LongitudinalCrackStatus(ByVal index As Long) As String
    LongitudinalCrackStatus = mResults(index).LongitudinalCrackStatus
End Property

Public Property Get ExtensionUsed(ByVal index As Long) As Boolean
    ExtensionUsed = mResults(index).ExtensionUsed
End Property

Public Property Get DirectStateDiagramExtensionEnabled(ByVal index As Long) As Boolean
    DirectStateDiagramExtensionEnabled = mResults(index).DirectStateDiagramExtensionEnabled
End Property

Public Property Get UtilizationCapacity(ByVal index As Long) As Double
    UtilizationCapacity = mResults(index).UtilCapacity
End Property

Public Property Get NUltimate(ByVal index As Long) As Double
    NUltimate = mResults(index).NUltimate
End Property

Public Property Get MxUltimate(ByVal index As Long) As Double
    MxUltimate = mResults(index).MxUltimate
End Property

Public Property Get MyUltimate(ByVal index As Long) As Double
    MyUltimate = mResults(index).MyUltimate
End Property

Public Property Get MomentUltimate(ByVal index As Long) As Double
    MomentUltimate = mResults(index).MomentUltimate
End Property

Public Property Get CapacityStateAvailable(ByVal index As Long) As Boolean
    CapacityStateAvailable = mResults(index).CapacityStateAvailable
End Property

Public Property Get CapacityEpsilon0(ByVal index As Long) As Double
    CapacityEpsilon0 = mResults(index).CapacityEpsilon0
End Property

Public Property Get CapacityKappaX(ByVal index As Long) As Double
    CapacityKappaX = mResults(index).CapacityKappaX
End Property

Public Property Get CapacityKappaY(ByVal index As Long) As Double
    CapacityKappaY = mResults(index).CapacityKappaY
End Property

Public Property Get MinConcreteStrain(ByVal index As Long) As Double
    MinConcreteStrain = mResults(index).MinConcreteStrain
End Property

Public Property Get MaxConcreteStrain(ByVal index As Long) As Double
    MaxConcreteStrain = mResults(index).MaxConcreteStrain
End Property

Public Property Get MinSteelStrain(ByVal index As Long) As Double
    MinSteelStrain = mResults(index).MinSteelStrain
End Property

Public Property Get MaxSteelStrain(ByVal index As Long) As Double
    MaxSteelStrain = mResults(index).MaxSteelStrain
End Property

Public Property Get CrackWidth(ByVal index As Long) As Double
    CrackWidth = mResults(index).CrackWidth
End Property

Public Property Get CrackAllowable(ByVal index As Long) As Double
    CrackAllowable = mResults(index).CrackAllowableValue
End Property

Public Property Get CrackUtilization(ByVal index As Long) As Double
    CrackUtilization = mResults(index).CrackUtil
End Property

Public Property Get MaxConcreteCompressionStress(ByVal index As Long) As Double
    MaxConcreteCompressionStress = mResults(index).MaxConcreteCompressionStress
End Property

Public Property Get LongitudinalCrackRbMc2(ByVal index As Long) As Double
    LongitudinalCrackRbMc2 = mResults(index).LongitudinalCrackRbMc2
End Property

Public Property Get LongitudinalCrackUtilization(ByVal index As Long) As Double
    LongitudinalCrackUtilization = mResults(index).LongitudinalCrackUtil
End Property

Public Property Get StabilityStatus(ByVal index As Long) As String
    StabilityStatus = mResults(index).StabilityStatus
End Property

Public Property Get StabilityCode(ByVal index As Long) As String
    StabilityCode = mResults(index).StabilityCode
End Property

Public Property Get StabilityBranch(ByVal index As Long) As String
    StabilityBranch = mResults(index).StabilityBranch
End Property

Public Property Get StabilityDesignN(ByVal index As Long) As Double
    StabilityDesignN = mResults(index).StabilityDesignN
End Property

Public Property Get StabilityDesignMx(ByVal index As Long) As Double
    StabilityDesignMx = mResults(index).StabilityDesignMx
End Property

Public Property Get StabilityDesignMy(ByVal index As Long) As Double
    StabilityDesignMy = mResults(index).StabilityDesignMy
End Property

Public Property Get StabilitySustainedN(ByVal index As Long) As Double
    StabilitySustainedN = mResults(index).StabilitySustainedN
End Property

Public Property Get StabilityMoment1(ByVal index As Long) As Double
    StabilityMoment1 = mResults(index).StabilityMoment1
End Property

Public Property Get StabilityMoment2(ByVal index As Long) As Double
    StabilityMoment2 = mResults(index).StabilityMoment2
End Property

Public Property Get StabilitySustainedMoment1(ByVal index As Long) As Double
    StabilitySustainedMoment1 = mResults(index).StabilitySustainedMoment1
End Property

Public Property Get StabilitySustainedMoment2(ByVal index As Long) As Double
    StabilitySustainedMoment2 = mResults(index).StabilitySustainedMoment2
End Property

Public Property Get StabilityDesignMoment1(ByVal index As Long) As Double
    StabilityDesignMoment1 = mResults(index).StabilityDesignMoment1
End Property

Public Property Get StabilityDesignMoment2(ByVal index As Long) As Double
    StabilityDesignMoment2 = mResults(index).StabilityDesignMoment2
End Property

' Возвращает Mx после устойчивости в той же пользовательской точке задания
' нагрузки. Берем нормализованное UserMx из текущего LC: это ровно тот момент,
' который затем участвует в DirectState/Capacity/Crack через общий load-path.
Public Property Get StabilityLoadPointMx(ByVal index As Long) As Double
    StabilityLoadPointMx = NormalizeMomentMx(mInputs(index).UserMx)
End Property

' Возвращает My после устойчивости в пользовательской точке задания нагрузки.
Public Property Get StabilityLoadPointMy(ByVal index As Long) As Double
    StabilityLoadPointMy = NormalizeMomentMy(mInputs(index).UserMy)
End Property

Public Property Get StabilityCriticalForce(ByVal index As Long) As Double
    StabilityCriticalForce = mResults(index).StabilityCriticalForce
End Property

Public Property Get StabilityReserveFactor(ByVal index As Long) As Double
    StabilityReserveFactor = mResults(index).StabilityReserveFactor
End Property

Public Property Get StabilityNcr1(ByVal index As Long) As Double
    StabilityNcr1 = mResults(index).StabilityNcr1
End Property

Public Property Get StabilityNcr2(ByVal index As Long) As Double
    StabilityNcr2 = mResults(index).StabilityNcr2
End Property

Public Property Get StabilityNultimate1(ByVal index As Long) As Double
    StabilityNultimate1 = mResults(index).StabilityNultimate1
End Property

Public Property Get StabilityNultimate2(ByVal index As Long) As Double
    StabilityNultimate2 = mResults(index).StabilityNultimate2
End Property

Public Property Get StabilityReserve1(ByVal index As Long) As Double
    StabilityReserve1 = mResults(index).StabilityReserve1
End Property

Public Property Get StabilityReserve2(ByVal index As Long) As Double
    StabilityReserve2 = mResults(index).StabilityReserve2
End Property

Public Property Get StabilityEta1(ByVal index As Long) As Double
    StabilityEta1 = mResults(index).StabilityEta1
End Property

Public Property Get StabilityEta2(ByVal index As Long) As Double
    StabilityEta2 = mResults(index).StabilityEta2
End Property

Public Property Get StabilityPhiL1(ByVal index As Long) As Double
    StabilityPhiL1 = mResults(index).StabilityPhiL1
End Property

Public Property Get StabilityPhiL2(ByVal index As Long) As Double
    StabilityPhiL2 = mResults(index).StabilityPhiL2
End Property

Public Property Get StabilityPhiLTable1(ByVal index As Long) As Double
    StabilityPhiLTable1 = mResults(index).StabilityPhiLTable1
End Property

Public Property Get StabilityPhiLTable2(ByVal index As Long) As Double
    StabilityPhiLTable2 = mResults(index).StabilityPhiLTable2
End Property

Public Property Get StabilityPhiM1(ByVal index As Long) As Double
    StabilityPhiM1 = mResults(index).StabilityPhiM1
End Property

Public Property Get StabilityPhiM2(ByVal index As Long) As Double
    StabilityPhiM2 = mResults(index).StabilityPhiM2
End Property

Public Property Get StabilityPhiValue1(ByVal index As Long) As Double
    StabilityPhiValue1 = mResults(index).StabilityPhiValue1
End Property

Public Property Get StabilityPhiValue2(ByVal index As Long) As Double
    StabilityPhiValue2 = mResults(index).StabilityPhiValue2
End Property

Public Property Get StabilityPhiP1(ByVal index As Long) As Double
    StabilityPhiP1 = mResults(index).StabilityPhiP1
End Property

Public Property Get StabilityPhiP2(ByVal index As Long) As Double
    StabilityPhiP2 = mResults(index).StabilityPhiP2
End Property

Public Property Get StabilityEccentricity1(ByVal index As Long) As Double
    StabilityEccentricity1 = mResults(index).StabilityEccentricity1
End Property

Public Property Get StabilityEccentricity2(ByVal index As Long) As Double
    StabilityEccentricity2 = mResults(index).StabilityEccentricity2
End Property

Public Property Get StabilityAccidentalEcc1(ByVal index As Long) As Double
    StabilityAccidentalEcc1 = mResults(index).StabilityAccidentalEcc1
End Property

Public Property Get StabilityAccidentalEcc2(ByVal index As Long) As Double
    StabilityAccidentalEcc2 = mResults(index).StabilityAccidentalEcc2
End Property

Public Property Get StabilityEffectiveLength1(ByVal index As Long) As Double
    StabilityEffectiveLength1 = mResults(index).StabilityEffectiveLength1
End Property

Public Property Get StabilityEffectiveLength2(ByVal index As Long) As Double
    StabilityEffectiveLength2 = mResults(index).StabilityEffectiveLength2
End Property

Public Property Get StabilitySlenderness1(ByVal index As Long) As Double
    StabilitySlenderness1 = mResults(index).StabilitySlenderness1
End Property

Public Property Get StabilitySlenderness2(ByVal index As Long) As Double
    StabilitySlenderness2 = mResults(index).StabilitySlenderness2
End Property

Public Property Get StabilityDepth1(ByVal index As Long) As Double
    StabilityDepth1 = mResults(index).StabilityDepth1
End Property

Public Property Get StabilityDepth2(ByVal index As Long) As Double
    StabilityDepth2 = mResults(index).StabilityDepth2
End Property

Public Property Get StabilityRadius1(ByVal index As Long) As Double
    StabilityRadius1 = mResults(index).StabilityRadius1
End Property

Public Property Get StabilityRadius2(ByVal index As Long) As Double
    StabilityRadius2 = mResults(index).StabilityRadius2
End Property

Public Property Get StabilityCoreDistance1(ByVal index As Long) As Double
    StabilityCoreDistance1 = mResults(index).StabilityCoreDistance1
End Property

Public Property Get StabilityCoreDistance2(ByVal index As Long) As Double
    StabilityCoreDistance2 = mResults(index).StabilityCoreDistance2
End Property

Public Property Get StabilityDelta1(ByVal index As Long) As Double
    StabilityDelta1 = mResults(index).StabilityDelta1
End Property

Public Property Get StabilityDelta2(ByVal index As Long) As Double
    StabilityDelta2 = mResults(index).StabilityDelta2
End Property

Public Property Get StabilityKb1(ByVal index As Long) As Double
    StabilityKb1 = mResults(index).StabilityKb1
End Property

Public Property Get StabilityKb2(ByVal index As Long) As Double
    StabilityKb2 = mResults(index).StabilityKb2
End Property

Public Property Get StabilityKs1(ByVal index As Long) As Double
    StabilityKs1 = mResults(index).StabilityKs1
End Property

Public Property Get StabilityKs2(ByVal index As Long) As Double
    StabilityKs2 = mResults(index).StabilityKs2
End Property

Public Property Get StabilityStiffnessD1(ByVal index As Long) As Double
    StabilityStiffnessD1 = mResults(index).StabilityStiffnessD1
End Property

Public Property Get StabilityStiffnessD2(ByVal index As Long) As Double
    StabilityStiffnessD2 = mResults(index).StabilityStiffnessD2
End Property

Public Property Get StabilityPlaneBranch1(ByVal index As Long) As String
    StabilityPlaneBranch1 = mResults(index).StabilityPlaneBranch1
End Property

Public Property Get StabilityPlaneBranch2(ByVal index As Long) As String
    StabilityPlaneBranch2 = mResults(index).StabilityPlaneBranch2
End Property

Public Property Get StabilityPlaneApplicable1(ByVal index As Long) As Boolean
    StabilityPlaneApplicable1 = mResults(index).StabilityPlaneApplicable1
End Property

Public Property Get StabilityPlaneApplicable2(ByVal index As Long) As Boolean
    StabilityPlaneApplicable2 = mResults(index).StabilityPlaneApplicable2
End Property

Public Property Get StabilityPlanePassed1(ByVal index As Long) As Boolean
    StabilityPlanePassed1 = mResults(index).StabilityPlanePassed1
End Property

Public Property Get StabilityPlanePassed2(ByVal index As Long) As Boolean
    StabilityPlanePassed2 = mResults(index).StabilityPlanePassed2
End Property

Public Property Get CrackFormed(ByVal index As Long) As Boolean
    CrackFormed = mResults(index).CrackFormed
End Property

Public Property Get CrackPhi1(ByVal index As Long) As Double
    CrackPhi1 = mResults(index).CrackPhi1
End Property

Public Property Get CrackPhi2(ByVal index As Long) As Double
    CrackPhi2 = mResults(index).CrackPhi2
End Property

Public Property Get CrackPhi3(ByVal index As Long) As Double
    CrackPhi3 = mResults(index).CrackPhi3
End Property

Public Property Get CrackPsiS(ByVal index As Long) As Double
    CrackPsiS = mResults(index).CrackPsiS
End Property

Public Property Get CrackSigmaS(ByVal index As Long) As Double
    CrackSigmaS = mResults(index).CrackSigmaS
End Property

Public Property Get CrackSigmaSCrc(ByVal index As Long) As Double
    CrackSigmaSCrc = mResults(index).CrackSigmaSCrc
End Property

Public Property Get CrackAs(ByVal index As Long) As Double
    CrackAs = mResults(index).CrackAs
End Property

Public Property Get CrackAbt(ByVal index As Long) As Double
    CrackAbt = mResults(index).CrackAbt
End Property

Public Property Get CrackDsEquivalent(ByVal index As Long) As Double
    CrackDsEquivalent = mResults(index).CrackDsEquivalent
End Property

Public Property Get CrackSpacingRaw(ByVal index As Long) As Double
    CrackSpacingRaw = mResults(index).CrackLsRaw
End Property

Public Property Get CrackSpacing(ByVal index As Long) As Double
    CrackSpacing = mResults(index).CrackLs
End Property

Public Property Get CrackLambdaCrc(ByVal index As Long) As Double
    CrackLambdaCrc = mResults(index).CrackLambdaCrc
End Property

Public Property Get CrackNcrc(ByVal index As Long) As Double
    CrackNcrc = mResults(index).CrackNcrc
End Property

Public Property Get CrackTensionRebarCount(ByVal index As Long) As Long
    CrackTensionRebarCount = mResults(index).CrackTensionRebarCount
End Property

Public Property Get CrackSectionDepthH(ByVal index As Long) As Double
    CrackSectionDepthH = mResults(index).CrackSectionDepthH
End Property

Public Property Get CrackCoverA(ByVal index As Long) As Double
    CrackCoverA = mResults(index).CrackCoverA
End Property

Public Property Get CrackTensionDepth(ByVal index As Long) As Double
    CrackTensionDepth = mResults(index).CrackTensionDepth
End Property

Public Property Get CrackEffectiveZoneDepth(ByVal index As Long) As Double
    CrackEffectiveZoneDepth = mResults(index).CrackEffectiveZoneDepth
End Property

Public Property Get CrackCentralBranch(ByVal index As Long) As Boolean
    CrackCentralBranch = mResults(index).CrackCentralBranch
End Property

Public Property Get GoverningUtilization(ByVal index As Long) As Double
    GoverningUtilization = mResults(index).GoverningUtil
End Property

Public Property Get CrackSteelElasticModulus() As Double
    CrackSteelElasticModulus = mMaterialProvider.SteelElasticModulus(cpCrackedNDS)
End Property

' Предельная сжатая деформация бетона для фактического НДС выбранного LC.
Public Function ConcreteCompressionLimitForCombination(ByVal index As Long) As Double
    Dim spec As CMaterialModelSpec
    Set spec = SummaryMaterialSpecForCombination(index)
    If Not spec Is Nothing Then ConcreteCompressionLimitForCombination = mMaterialProvider.ConcreteCompressionLimitFromSpec(spec)
End Function

' Предельная растянутая деформация бетона для фактического НДС выбранного LC.
Public Function ConcreteTensionLimitForCombination(ByVal index As Long) As Double
    Dim spec As CMaterialModelSpec
    Set spec = SummaryMaterialSpecForCombination(index)
    If Not spec Is Nothing Then ConcreteTensionLimitForCombination = mMaterialProvider.ConcreteTensionLimitFromSpec(spec)
End Function

' True, если текущая диаграмма LC учитывает растянутый бетон и его предел нужно выводить.
Public Function ConcreteTensionLimitEnabledForCombination(ByVal index As Long) As Boolean
    Dim spec As CMaterialModelSpec
    Set spec = SummaryMaterialSpecForCombination(index)
    If Not spec Is Nothing Then ConcreteTensionLimitEnabledForCombination = mMaterialProvider.ConcreteTensionLimitEnabledFromSpec(spec)
End Function

' Предельная сжатая деформация арматуры для фактического НДС выбранного LC.
Public Function SteelCompressionLimitForCombination(ByVal index As Long) As Double
    Dim spec As CMaterialModelSpec
    Set spec = SummaryMaterialSpecForCombination(index)
    If Not spec Is Nothing Then SteelCompressionLimitForCombination = mMaterialProvider.SteelCompressionLimitFromSpec(spec)
End Function

' Предельная растянутая деформация арматуры для фактического НДС выбранного LC.
Public Function SteelTensionLimitForCombination(ByVal index As Long) As Double
    Dim spec As CMaterialModelSpec
    Set spec = SummaryMaterialSpecForCombination(index)
    If Not spec Is Nothing Then SteelTensionLimitForCombination = mMaterialProvider.SteelTensionLimitFromSpec(spec)
End Function

' Выбирает спецификацию того StateSolution, который представлен в старых
' summary-полях деформаций. Предпочтение отдано прямому StrengthState, затем
' CrackedState; CapacityState используется только если других states нет.
Private Function SummaryMaterialSpecForCombination(ByVal index As Long) As CMaterialModelSpec
    Dim stateResult As CSectionStateResult
    Set stateResult = mResults(index).FindState(sstStrengthState)
    If stateResult Is Nothing Then Set stateResult = mResults(index).FindState(sstCrackedState)
    If stateResult Is Nothing Then Set stateResult = mResults(index).FindState(sstCapacityState)
    If Not stateResult Is Nothing Then Set SummaryMaterialSpecForCombination = stateResult.MaterialSpec
End Function

Public Property Get GoverningCombinationIndex() As Long
    GoverningCombinationIndex = mGoverningCombinationIndex
End Property

Public Property Get GoverningCombinationID() As String
    If mGoverningCombinationIndex > 0 Then GoverningCombinationID = mInputs(mGoverningCombinationIndex).CombinationID
End Property

Public Property Get GoverningCombinationName() As String
    If mGoverningCombinationIndex > 0 Then GoverningCombinationName = mInputs(mGoverningCombinationIndex).Comment
End Property

Public Property Get CrackGoverningCombinationIndex() As Long
    CrackGoverningCombinationIndex = mCrackGoverningCombinationIndex
End Property

Public Property Get CrackGoverningCombinationID() As String
    If mCrackGoverningCombinationIndex > 0 Then CrackGoverningCombinationID = mInputs(mCrackGoverningCombinationIndex).CombinationID
End Property

Public Property Get CrackGoverningCombinationName() As String
    If mCrackGoverningCombinationIndex > 0 Then CrackGoverningCombinationName = mInputs(mCrackGoverningCombinationIndex).Comment
End Property

Public Property Get ElapsedSeconds() As Double
    ElapsedSeconds = mElapsedSeconds
End Property

Public Property Get DiagnosticLog() As String
    DiagnosticLog = mDiagnosticLog
End Property

Public Property Get InvalidInputCount() As Long
    Dim i As Long
    For i = 1 To mCount
        If mStatusPolicy.IsInvalid(mResults(i).Status) Then InvalidInputCount = InvalidInputCount + 1
    Next i
End Property

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
    Dim detail As String
    detail = Trim$(mInputs(index).InputError)
    If Len(detail) = 0 And mStatusPolicy.IsInvalid(mResults(index).StabilityStatus) Then _
        detail = Trim$(mResults(index).StabilityBranch)
    If Len(detail) = 0 And mStatusPolicy.IsInvalid(mResults(index).CapacityStatus) Then _
        detail = Trim$(mResults(index).CapacityLimitState)
    If Len(detail) = 0 And mStatusPolicy.IsInvalid(mResults(index).DirectStateStatus) Then _
        detail = Trim$(mResults(index).Status)
    If Len(detail) = 0 Then detail = Trim$(mResults(index).Status)
    If Len(detail) = 0 Then detail = mStatusPolicy.InputErr

    DetailedInvalidInputMessage = "CombinationID=" & DisplayCombinationID(index) & _
        ": " & detail & " (статус " & mResults(index).Status & ")"
End Function

' Формирует устойчивое имя LC для сообщения об ошибке даже при неполной строке
' сочетания, когда пользователь мог оставить CombinationID пустым.
Private Function DisplayCombinationID(ByVal index As Long) As String
    DisplayCombinationID = Trim$(mInputs(index).CombinationID)
    If Len(DisplayCombinationID) = 0 Then DisplayCombinationID = "#" & CStr(index)
End Function

Public Property Get StateAvailable(ByVal index As Long) As Boolean
    StateAvailable = mResults(index).StateAvailable
End Property

Public Property Get NamedStateCount(ByVal index As Long) As Long
    NamedStateCount = mResults(index).StateCount
End Property

Public Property Get NamedState(ByVal index As Long, ByVal stateIndex As Long) As CSectionStateResult
    Set NamedState = mResults(index).StateAt(stateIndex)
End Property

Public Function FindNamedState(ByVal index As Long, ByVal stateType As ESectionStateType) As CSectionStateResult
    Set FindNamedState = mResults(index).FindState(stateType)
End Function

Public Property Get StateAvailableCount() As Long
    Dim i As Long
    For i = 1 To mCount
        If mResults(i).StateCount > 0 Then StateAvailableCount = StateAvailableCount + 1
    Next i
End Property

Public Property Get Epsilon0(ByVal index As Long) As Double
    Epsilon0 = mResults(index).Epsilon0
End Property

Public Property Get KappaX(ByVal index As Long) As Double
    KappaX = mResults(index).KappaX
End Property

Public Property Get KappaY(ByVal index As Long) As Double
    KappaY = mResults(index).KappaY
End Property

Public Property Get Nint(ByVal index As Long) As Double
    Nint = mResults(index).Nint
End Property

Public Property Get Mxint(ByVal index As Long) As Double
    Mxint = mResults(index).Mxint
End Property

Public Property Get Myint(ByVal index As Long) As Double
    Myint = mResults(index).Myint
End Property

Public Property Get ResidualN(ByVal index As Long) As Double
    ResidualN = mResults(index).ResidualN
End Property

Public Property Get ResidualMx(ByVal index As Long) As Double
    ResidualMx = mResults(index).ResidualMx
End Property

Public Property Get ResidualMy(ByVal index As Long) As Double
    ResidualMy = mResults(index).ResidualMy
End Property

Public Property Get StateIterations(ByVal index As Long) As Long
    StateIterations = mResults(index).StateIterations
End Property

' ==========================================================================
' ДЛЯ ТЕСТОВ
' ==========================================================================
' Эти свойства раскрывают промежуточные численные данные batch-расчета,
' которые удобны regression-тестам: отдельно проверяются lambda capacity
' и факт сходимости прямого НДС.

Public Property Get LambdaCapacity(ByVal index As Long) As Double
    LambdaCapacity = mResults(index).LambdaCapacity
End Property

Public Property Get StateConverged(ByVal index As Long) As Boolean
    StateConverged = mResults(index).StateConverged
End Property



