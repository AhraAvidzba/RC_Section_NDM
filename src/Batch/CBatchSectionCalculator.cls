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

Private Const INITIAL_COMBINATION_CAPACITY As Long = 20 ' Начальный размер внутренних массивов; при расширении rngLoadCombinations хранилище растет.
Private Const COMBINATION_CAPACITY_STEP As Long = 20 ' Шаг расширения, чтобы не делать ReDim Preserve на каждой новой строке LC.
Private Const CAPACITY_PATH_NONE As String = "None" ' Capacity не запускался для LC.
Private Const DEFAULT_ZERO_MOMENT_PER_DEPTH As Double = 4903.325 ' 0.5 tf*m/m во внутренних Н*мм/мм.

' Данные одной строки таблицы сочетаний. Batch хранит их отдельно от результатов,
' чтобы было видно, где исходные нагрузки пользователя, а где то, что вернул расчет.
Private Type TCombinationInput
    CombinationID As String
    SourceDataOffset As Long
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
Private mMaterialProvider As CMaterialModelProvider ' Единая точка выдачи диаграмм для Strength/CrackInitiation/CrackedNDS.
Private mProfiles As CCalculationProfileCatalog ' Каталог расчетных профилей PR1/PR2/... из Config.
Private mStatusPolicy As CBatchStatusPolicy ' Единый словарь пользовательских статусов и правил их агрегирования.
Private mCount As Long ' Сколько сочетаний реально добавлено в пакет.

Private mInputs() As TCombinationInput ' Исходные данные LC, прочитанные из фактических строк rngLoadCombinations.
Private mResults() As CCombinationResult ' Расчетные результаты LC для Summary/Results; размер массива растет вместе с количеством LC.
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
Private mWorstCombinationIndex As Long ' Индекс LC, выбранный как Worst по пользовательскому критерию Summary.
Private mWorstCombinationCriterion As String ' StrengthStrain, StrengthCapacity, Cracks или Stability.
Private mSolverCallCount As Long ' Сколько раз за batch запускался CSectionSolver.Solve для поиска равновесия НДС.
Private mElapsedSeconds As Double ' Полное время пакетного расчета, сек.
Private mDiagnosticLog As String ' Короткий текстовый журнал по LC для диагностики batch-потока.
Private mTransformedReferenceValid(1 To 2) As Boolean ' Кэш центров приведенного сечения для ULS/SLS внутри одного batch.
Private mTransformedReferenceX(1 To 2) As Double ' X центра приведенного сечения, мм.
Private mTransformedReferenceY(1 To 2) As Double ' Y центра приведенного сечения, мм.

Private mCrackAllowable As Double ' Допустимая ширина раскрытия трещины a_crc,ult, мм.
Private mCrackPsiMode As String ' User, Auto или AlwaysCalc: как определять psi_s для нормальных трещин.
Private mCrackSigmaSCrcAveragingMode As String ' Как усреднять sigma_s,crc при уточнении psi_s по PostCrackState.
Private mCrackPhi1 As Double ' Пользовательский коэффициент phi1 для формулы раскрытия.
Private mCrackPhi2 As Double ' Пользовательский коэффициент phi2 для формулы раскрытия.
Private mCrackPhi3Mode As String ' User или Auto: как принимать phi3.
Private mCrackPhi3 As Double ' Пользовательский коэффициент phi3, если Phi3Mode = User.
Private mCrackPsiS As Double ' Пользовательский коэффициент psi_s, если PsiMode = User.
Private mCrackTensionZoneMode As String ' Effective или FullTension: какая бетонная зона входит в Abt.
Private mCrackCoverDistanceMode As String ' Как измерять a_s: до ближайшего контура или до дальней растянутой линии.
Private mCrackFormationPath As String ' Путь поиска образования трещины: Auto, lambda*Mxy, lambda*N или lambda*NMxy.
Private mCrackFormationSolutionStrategy As String ' Auto, UltimateStrain или LoadMultiplier для поиска уровня трещинообразования.
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
Private mCapacityMaxLambda As Double ' Верхняя защитная граница lambda для расширения скобки LoadMultiplier.
Private mCapacityToleranceLambda As Double ' Точность lambda в одномерном поиске LoadMultiplier.
Private mCapacityToleranceStrain As Double ' Точность попадания в предельную деформацию UltimateStrain.
Private mCapacityMaxRetries As Long ' Повторные попытки LoadMultiplier при несходимости той же lambda-точки.
Private mCapacityBaseLoadSteps As Long ' Базовое число load steps для первой попытки lambda-точки capacity.
Private mCapacitySolverMaxIterations As Long ' Лимит итераций внутреннего solver-а при capacity-поиске.
Private mReport As CExecutionReport ' Необязательный человекочитаемый отчет выполнения; расчетная логика от него не зависит.
Private mReportUnits As CUnitSystem ' Единицы и знаки для человекочитаемых физических величин execution_report.
Private mSettingsValid As Boolean ' False, если ApplySettings нашел ошибку Config и все LC должны получить InputErr.
Private mSettingsError As String ' Человекочитаемая причина ошибки чтения настроек Config.

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
Private mStabilityLoadCombinationID() As String ' LC, к которому относится строка длительных нагрузок.
Private mStabilityLoadN() As Double ' N от постоянных или постоянных+длительных нагрузок, Н.
Private mStabilityLoadUserMx() As Double ' Пользовательский Mx длительной части в точке приложения нагрузки.
Private mStabilityLoadUserMy() As Double ' Пользовательский My длительной части в точке приложения нагрузки.

' Инициализирует внутреннее состояние объекта значениями по умолчанию, безопасными для последующей настройки.
Private Sub Class_Initialize()
    Set mStatusPolicy = New CBatchStatusPolicy
    Set mMomentZeroFilter = New CMomentZeroFilter
    mSettingsValid = True
    mSettingsError = vbNullString
    mZeroMomentPerDepth = DEFAULT_ZERO_MOMENT_PER_DEPTH
    mMomentZeroFilter.Initialize mZeroMomentPerDepth
    EnsureCombinationCapacity INITIAL_COMBINATION_CAPACITY
    EnsureStabilityLoadCapacity INITIAL_COMBINATION_CAPACITY

    mCrackAllowable = 0.3
    mCrackPsiMode = "AlwaysCalc"
    mCrackSigmaSCrcAveragingMode = "AllSelected"
    mCrackPhi1 = 1.4
    mCrackPhi2 = 0.5
    mCrackPhi3Mode = "Auto"
    mCrackPhi3 = 1#
    mCrackPsiS = 1#
    mCrackTensionZoneMode = "Effective"
    mCrackCoverDistanceMode = "GlobalExtreme"
    mCrackFormationPath = "Auto"
    mCrackFormationSolutionStrategy = "Auto"
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
    mCapacityMaxLambda = 1024#
    mCapacityToleranceLambda = 0.01
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

' Принимает готовую модель и материалы перед добавлением/расчетом LC.
' После этого класс не обращается к геометрическим генераторам: все сочетания
' считаются на одной и той же CSectionModel.
Public Sub Initialize(ByVal section As CSectionModel, ByVal materialProvider As CMaterialModelProvider)
    If section Is Nothing Then Err.Raise vbObjectError + 3900, "CBatchSectionCalculator", "Модель сечения не передана."
    If section.ConcreteCount <= 0 Then Err.Raise vbObjectError + 3901, "CBatchSectionCalculator", "В модели сечения нет бетонных элементов."
    If materialProvider Is Nothing Then Err.Raise vbObjectError + 3902, "CBatchSectionCalculator", "Поставщик материалов не передан."
    Set mSection = section
    Set mMaterialProvider = materialProvider
    ClearTransformedReferenceCache
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
    On Error GoTo InvalidSettings
    If settings Is Nothing Then Exit Sub
    mSettingsValid = True
    mSettingsError = vbNullString
    Set mReportUnits = units
    Dim zeroMomentSettingRead As Boolean
    zeroMomentSettingRead = settings.HasKey("Calculation.ZeroMomentPerDepth")
    mCrackAllowable = settings.GetDouble("SLS.Crack.Allowable", mCrackAllowable)
    mCrackPsiMode = settings.GetString("SLS.Crack.PsiMode", mCrackPsiMode)
    mCrackSigmaSCrcAveragingMode = settings.GetString("SLS.Crack.SigmaSCrcAveragingMode", mCrackSigmaSCrcAveragingMode)
    mCrackFormationPath = settings.GetString("SLS.Crack.InitiationLoadPath", mCrackFormationPath)
    mCrackFormationSolutionStrategy = settings.GetString("SLS.Crack.InitiationSolutionStrategy", mCrackFormationSolutionStrategy)
    mWorstCombinationCriterion = settings.GetString("General.WorstCombinationCriterion", mWorstCombinationCriterion)
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
    Exit Sub

InvalidSettings:
    mSettingsValid = False
    mSettingsError = Err.Description
    ReportStep "Настройки расчета содержат ошибку: " & mSettingsError
End Sub

' Читает обязательную числовую настройку устойчивости без скрытой подстановки
' default. Если строка Config не прочиталась, пуста или содержит не число,
' расчет останавливается как ошибка исходных данных.
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

' Очищает список LC и накопленные результаты.
' Используется перед чтением новой таблицы сочетаний, чтобы старые статусы не
' попали в новый batch summary.
Public Sub ClearCombinations()
    mCount = 0
    ClearStabilityDurationLoads
    mGoverningCombinationIndex = 0
    mCrackGoverningCombinationIndex = 0
    mWorstCombinationIndex = 0
    mElapsedSeconds = 0#
    mSolverCallCount = 0
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
        Optional ByVal capacityLoadPath As String = vbNullString, Optional ByVal sourceDataOffset As Long = 0)
    AddCombinationInternal combinationID, nValue, mxValue, myValue, profileId, comment, True, vbNullString, capacityLoadPath, sourceDataOffset
End Sub

' Сохраняет строку LC с ошибкой ввода, чтобы пользователь увидел ее в отчете.
' Такие сочетания не передаются в solver, но отображаются как InvalidInput.
Public Sub AddInvalidCombination(ByVal combinationID As String, ByVal profileId As String, _
        ByVal comment As String, ByVal reason As String, Optional ByVal capacityLoadPath As String = vbNullString, _
        Optional ByVal sourceDataOffset As Long = 0)
    AddCombinationInternal combinationID, 0#, 0#, 0#, profileId, comment, False, reason, capacityLoadPath, sourceDataOffset
End Sub

' Единая точка добавления строки сочетания. Через нее проходят и корректные
' LC, и строки с ошибкой ввода; это гарантирует одинаковую нумерацию,
' сохранение комментариев и нормализацию малых моментов только для валидных
' нагрузок.
Private Sub AddCombinationInternal(ByVal combinationID As String, ByVal nValue As Double, ByVal mxValue As Double, _
        ByVal myValue As Double, ByVal profileId As String, ByVal comment As String, _
        ByVal inputValid As Boolean, ByVal inputError As String, ByVal capacityLoadPath As String, _
        ByVal sourceDataOffset As Long)
    EnsureCombinationCapacity mCount + 1
    mCount = mCount + 1
    If Len(Trim$(combinationID)) = 0 Then combinationID = "ROW" & CStr(mCount)
    mInputs(mCount).CombinationID = combinationID
    mInputs(mCount).SourceDataOffset = sourceDataOffset
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
        ReDim mResults(1 To newCapacity)
    Else
        ReDim Preserve mInputs(1 To newCapacity)
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
    Else
        ReDim Preserve mStabilityLoadCombinationID(1 To newCapacity)
        ReDim Preserve mStabilityLoadN(1 To newCapacity)
        ReDim Preserve mStabilityLoadUserMx(1 To newCapacity)
        ReDim Preserve mStabilityLoadUserMy(1 To newCapacity)
    End If
End Sub

' Возвращает текущую емкость массива длительных нагрузок; неинициализированный массив дает 0.
Private Function StabilityLoadCapacity() As Long
    On Error GoTo EmptyStorage
    StabilityLoadCapacity = UBound(mStabilityLoadCombinationID)
    Exit Function
EmptyStorage:
End Function

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

' Нормализует пару пользовательских моментов и заново переносит их к системе
' координат модели. Это единая точка для исходных LC и усилий после Stability.
Private Sub NormalizeCombinationMoments(ByVal index As Long)
    If index < 1 Or index > mCount Then Exit Sub
    If Not mInputs(index).InputValid Then Exit Sub

    Dim loadState As CSectionLoadState
    Set loadState = LoadStateForCombination(index)
    mInputs(index).UserMx = loadState.LoadPointMx
    mInputs(index).UserMy = loadState.LoadPointMy
    mInputs(index).Mx = loadState.InternalMx
    mInputs(index).My = loadState.InternalMy
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
    mWorstCombinationIndex = 0
    mSolverCallCount = 0
    mDiagnosticLog = vbNullString
    ResetSectionEquilibriumSolveCount
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
    stage = "SelectWorstCombination"
    SelectWorstCombination
    stage = "Finalize"
    mSolverCallCount = SectionEquilibriumSolveCount()
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
    ReportValue "ID профиля", mInputs(index).ProfileId
    If Not mInputs(index).InputValid Then
        ReportStep "Сочетание не передано в solver из-за ошибки входных данных: " & mInputs(index).InputError
        MarkInvalid index, mInputs(index).InputError
        Exit Sub
    End If

    If Not mSettingsValid Then
        ReportStep "Сочетание не передано в solver из-за ошибки настроек: " & mSettingsError
        MarkInvalid index, mSettingsError
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
    ReportValue "N", ReportForceText(mInputs(index).N)
    ReportValue "Mx расчетный", ReportMxText(mInputs(index).Mx)
    ReportValue "My расчетный", ReportMyText(mInputs(index).My)

    Dim profileError As String
    If Not profile.ValidateForCalculation(profileError) Then
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

    ReportSection "Продольный изгиб и устойчивость"
    If profile.StabilityEnabled Then
        ReportStep "Запускается фильтр продольного изгиба/устойчивости по профилю."
        RunStabilityFilter index, profile.StabilityValueSet
    Else
        ReportStep "Устойчивость не запрошена профилем."
        mResults(index).ClearStabilityResult
        mResults(index).StabilityStatus = mStatusPolicy.NotApplicable
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
    mResults(index).GoverningUtil = MaxDouble(MaxDouble(MaxDouble(mResults(index).UtilCapacity, mResults(index).CrackUtil), _
        mResults(index).LongitudinalCrackUtil), mResults(index).StabilityUtil)
    mResults(index).Status = mStatusPolicy.AggregateMeta(mResults(index).DirectStateMeta, _
        mResults(index).CapacityMeta, mResults(index).CrackMeta, _
        mResults(index).LongitudinalCrackMeta, mResults(index).StabilityMeta)
    mResults(index).StoreOverallMeta mResults(index).OverallMeta
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
    TransformedReferencePoint valueSet, referenceX, referenceY

    Dim sustainedN As Double
    Dim sustainedMx As Double
    Dim sustainedMy As Double
    StabilityDurationLoadForCombination mInputs(index).CombinationID, referenceX, referenceY, _
        sustainedN, sustainedMx, sustainedMy

    Dim stability As CStabilityCalculator
    Set stability = New CStabilityCalculator
    Dim loadState As CSectionLoadState
    Set loadState = LoadStateForCombination(index)
    Dim referenceMx As Double
    Dim referenceMy As Double
    referenceMx = loadState.MxAboutPoint(referenceY)
    referenceMy = loadState.MyAboutPoint(referenceX)

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

    ReportValue "Статус устойчивости", mResults(index).StabilityStatus
    ReportValue "Нормативный блок", StabilityCodeDisplayName(stability.Code)
    ReportValue "Материалы для устойчивости", MaterialValueSetToText(valueSet)
    ReportValue "Расчетная ветвь", StabilityBranchDisplayName(stability.Branch)
    ReportValue "Контрольная сила устойчивости", ReportCompressionForceText(stability.CriticalForce)
    ReportValue "Минимальный запас устойчивости", FormatNumberInvariant(stability.ReserveFactor)
    ReportValue "Итоговые N/Mx/My после фильтра", _
        "N=" & ReportForceText(mInputs(index).N) & "; " & _
        "Mx=" & ReportMxText(mInputs(index).Mx) & "; " & _
        "My=" & ReportMyText(mInputs(index).My)
    ReportStabilityDetails stability
    Exit Sub

Failed:
    mResults(index).StabilityStatus = mStatusPolicy.InputErr
    mResults(index).StabilityBranch = Err.Description
    ReportStep "Устойчивость не рассчитана: " & Err.Description
End Sub

' Пишет в txt-отчет подробные данные фильтра устойчивости по двум главным
' плоскостям. Таблица Results остается основным машинным выводом, а этот блок
' нужен для быстрой трассировки того, какая ветвь СП сработала в сценарии LC.
Private Sub ReportStabilityDetails(ByVal stability As CStabilityCalculator)
    If stability Is Nothing Then Exit Sub

    If Len(stability.StopReason) > 0 Then _
        ReportValue "Причина остановки устойчивости", stability.StopReason
    ReportValue "Продольная сила для проверки", ReportForceText(stability.DesignN)
    ReportValue "Моменты после учета eta по главным плоскостям 1/2", _
        "M1=" & ReportMomentMagnitudeText(stability.DesignMoment1) & "; " & _
        "M2=" & ReportMomentMagnitudeText(stability.DesignMoment2)
    ReportValue "Длительная часть продольной силы", ReportSignedCompressionForceText(stability.SustainedN)
    ReportValue "Длительные моменты по главным плоскостям 1/2", _
        "M1l=" & ReportMomentMagnitudeText(stability.SustainedMoment1) & "; " & _
        "M2l=" & ReportMomentMagnitudeText(stability.SustainedMoment2)

    ReportValue "Плоскость 1: ветвь и итог", _
        StabilityBranchDisplayName(stability.PlaneBranch1) & "; " & StabilityPassDisplayName(stability.PlanePassed1)
    ReportValue "Плоскость 1: eta и запас", _
        FormatNumberInvariant(stability.Eta1) & "; " & FormatNumberInvariant(stability.Reserve1)
    ReportValue "Плоскость 1: Ncr/Nult", _
        "Ncr=" & ReportCompressionForceText(stability.Ncr1) & "; Nult=" & ReportCompressionForceText(stability.Nultimate1)
    ReportValue "Плоскость 1: h, i, r и гибкость l0/i", _
        "h=" & ReportLengthText(stability.Depth1) & "; " & _
        "i=" & ReportLengthText(stability.Radius1) & "; " & _
        "r=" & ReportLengthText(stability.CoreDistance1) & "; " & _
        "l0/i=" & FormatNumberInvariant(stability.Slenderness1)
    ReportValue "Плоскость 1: e, ea, phi_l и D", _
        "e=" & ReportLengthText(stability.Eccentricity1) & "; " & _
        "ea=" & ReportLengthText(stability.AccidentalEccentricity1) & "; " & _
        "phi_l=" & FormatNumberInvariant(stability.PhiL1) & "; " & _
        "D=" & ReportStiffnessText(stability.StiffnessD1)

    ReportValue "Плоскость 2: ветвь и итог", _
        StabilityBranchDisplayName(stability.PlaneBranch2) & "; " & StabilityPassDisplayName(stability.PlanePassed2)
    ReportValue "Плоскость 2: eta и запас", _
        FormatNumberInvariant(stability.Eta2) & "; " & FormatNumberInvariant(stability.Reserve2)
    ReportValue "Плоскость 2: Ncr/Nult", _
        "Ncr=" & ReportCompressionForceText(stability.Ncr2) & "; Nult=" & ReportCompressionForceText(stability.Nultimate2)
    ReportValue "Плоскость 2: h, i, r и гибкость l0/i", _
        "h=" & ReportLengthText(stability.Depth2) & "; " & _
        "i=" & ReportLengthText(stability.Radius2) & "; " & _
        "r=" & ReportLengthText(stability.CoreDistance2) & "; " & _
        "l0/i=" & FormatNumberInvariant(stability.Slenderness2)
    ReportValue "Плоскость 2: e, ea, phi_l и D", _
        "e=" & ReportLengthText(stability.Eccentricity2) & "; " & _
        "ea=" & ReportLengthText(stability.AccidentalEccentricity2) & "; " & _
        "phi_l=" & FormatNumberInvariant(stability.PhiL2) & "; " & _
        "D=" & ReportStiffnessText(stability.StiffnessD2)
End Sub

' Берет длительную часть нагрузки для LC. Если пользователь не заполнил строку,
' принимаются нули: это допустимый ввод и не блокирует основной расчет.
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

' Подменяет усилия LC расчетными усилиями после фильтра устойчивости. Сам
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

    mInputs(index).N = loadState.N
    mInputs(index).UserMx = loadState.LoadPointMx
    mInputs(index).UserMy = loadState.LoadPointMy
    mInputs(index).Mx = loadState.InternalMx
    mInputs(index).My = loadState.InternalMy

    ' Summary должен показывать те же моменты, которые после zero-filter ушли
    ' дальше в DirectState/Capacity/Crack, а не промежуточный DesignLoad.
    Dim finalDesignMx As Double
    Dim finalDesignMy As Double
    finalDesignMx = loadState.MxAboutPoint(referenceY)
    finalDesignMy = loadState.MyAboutPoint(referenceX)
    mResults(index).UpdateStabilityDesignLoads mInputs(index).N, finalDesignMx, finalDesignMy, _
        stability.PrincipalAngleRad
End Sub

' Возвращает объект нагрузки LC: он единолично переносит моменты между
' пользовательской точкой, произвольным центром и внутренним началом модели.
Private Function LoadStateForCombination(ByVal index As Long) As CSectionLoadState
    Dim loadState As CSectionLoadState
    Set loadState = New CSectionLoadState
    loadState.Initialize mInputs(index).N, mInputs(index).UserMx, mInputs(index).UserMy, _
        mLoadReferenceX, mLoadReferenceY, mMomentZeroFilter, mConcreteDepthForMx, mConcreteDepthForMy
    Set LoadStateForCombination = loadState
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
    Set props = New CSectionPropertiesCalculator
    props.CalculateTransformedByModuli mSection, _
        mMaterialProvider.ConcreteElasticModulusFromValueSet(valueSet), _
        mMaterialProvider.SteelElasticModulusFromValueSet(valueSet)
    referenceX = props.CentroidX
    referenceY = props.CentroidY
    mTransformedReferenceX(slot) = referenceX
    mTransformedReferenceY(slot) = referenceY
    mTransformedReferenceValid(slot) = True
End Sub

' Очищает кэш центров приведенного сечения, когда batch получает новую
' геометрию или новый provider материалов. Сами значения зависят только от
' CSectionModel и модулей Es/Eb выбранного набора характеристик.
Private Sub ClearTransformedReferenceCache()
    Dim i As Long
    For i = 1 To 2
        mTransformedReferenceValid(i) = False
        mTransformedReferenceX(i) = 0#
        mTransformedReferenceY(i) = 0#
    Next i
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
' Именно относительно этой точки новый strength summary показывает N/Mx/My
' в группе "относительно центральных осей приведенного сечения".
Private Sub StrengthReferencePointForCombination(ByVal index As Long, _
        ByRef referenceX As Double, ByRef referenceY As Double)
    Dim profile As CCalculationProfile
    If Not mProfiles.HasProfile(mInputs(index).ProfileId) Then Exit Sub
    Set profile = mProfiles.ProfileById(mInputs(index).ProfileId)
    TransformedReferencePoint profile.StrengthSpec.ValueSet, referenceX, referenceY
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
    ReportValue "Путь нагрузки для несущей", IIf(loadPath.HasPath, loadPath.DisplayName, mInputs(index).CapacityLoadPath)

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

    Dim pureAxialForMethod As Boolean
    pureAxialForMethod = CapacityPathIsPureAxialForMethod(index, loadPath)

    ReportValue "Стратегия поиска несущей", NormalizedCapacitySolutionStrategy()
    Select Case NormalizedCapacitySolutionStrategy()
        Case "Auto"
            If pureAxialForMethod Then
                ReportStep "Несущая способность: стратегия Auto. После переноса к центру приведенного сечения траектория является чисто осевой, поэтому сразу используется LoadMultiplier: для такой постановки UltimateStrain вырождается на горизонтальном участке диаграммы."
                cap.SolveByLoadPathMultiplier mSection, mMaterialProvider.ConcreteMaterialFromSpec(strengthSpec), _
                    mMaterialProvider.SteelMaterialFromSpec(strengthSpec), loadPath.NOffset, loadPath.NBase, _
                    loadPath.MxOffset, loadPath.MxBase, loadPath.MyOffset, loadPath.MyBase, True
            Else
                If loadPath.ForceOnly Then
                    ReportStep "Несущая способность: стратегия Auto. Пользователь масштабирует только N, но относительно центра приведенного сечения есть изгибающий момент; сначала пробуем UltimateStrain, затем при неудаче LoadMultiplier."
                Else
                    ReportStep "Несущая способность: стратегия Auto сначала пробует UltimateStrain; если предельная плоскость не находится устойчиво, тот же путь нагрузки повторяется через LoadMultiplier."
                End If
                cap.SolveByAutoLoadPath mSection, mMaterialProvider.ConcreteMaterialFromSpec(strengthSpec), _
                    mMaterialProvider.SteelMaterialFromSpec(strengthSpec), loadPath.NOffset, loadPath.NBase, _
                    loadPath.MxOffset, loadPath.MxBase, loadPath.MyOffset, loadPath.MyBase, _
                    False, loadPath.ForceOnly
            End If
        Case "UltimateStrain"
            If pureAxialForMethod Then
                ReportStep "Несущая способность: выбран UltimateStrain, но после переноса к центру приведенного сечения траектория чисто осевая и считается через LoadMultiplier как более устойчивая одномерная постановка."
                cap.SolveByLoadPathMultiplier mSection, mMaterialProvider.ConcreteMaterialFromSpec(strengthSpec), _
                    mMaterialProvider.SteelMaterialFromSpec(strengthSpec), loadPath.NOffset, loadPath.NBase, _
                    loadPath.MxOffset, loadPath.MxBase, loadPath.MyOffset, loadPath.MyBase, True
            ElseIf loadPath.ForceOnly Then
                ReportStep "Несущая способность: выбран UltimateStrain. Пользователь масштабирует только N, но относительно центра приведенного сечения есть изгибающий момент; сначала пробуем UltimateStrain, затем при неудаче LoadMultiplier."
                cap.SolveByAutoLoadPath mSection, mMaterialProvider.ConcreteMaterialFromSpec(strengthSpec), _
                    mMaterialProvider.SteelMaterialFromSpec(strengthSpec), loadPath.NOffset, loadPath.NBase, _
                    loadPath.MxOffset, loadPath.MxBase, loadPath.MyOffset, loadPath.MyBase, _
                    False, True
            Else
                ReportStep "Несущая способность: расчет идет напрямую через UltimateStrain по выбранной λ-траектории."
                cap.SolveByUltimateLoadPath mSection, mMaterialProvider.ConcreteMaterialFromSpec(strengthSpec), _
                    mMaterialProvider.SteelMaterialFromSpec(strengthSpec), loadPath.NOffset, loadPath.NBase, _
                    loadPath.MxOffset, loadPath.MxBase, loadPath.MyOffset, loadPath.MyBase
            End If
        Case "LoadMultiplier"
            ReportStep "Несущая способность: расчет идет через LoadMultiplier, то есть программа подбирает λ и для каждой lambda-точки заново решает НДС."
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
    ReportCapacityMethodChoice index, cap, loadPath, pureAxialForMethod

    ReportValue "Статус несущей способности", mResults(index).CapacityStatus
    ReportValue "Фактический метод несущей", mResults(index).CapacitySolutionMethod
    ReportValue "lambda несущей", FormatNumberInvariant(mResults(index).LambdaCapacity)
    ReportValue "Фактический путь несущей", CapacityLoadPathDisplay(mResults(index).CapacityPathResolved)
    ReportBlock "Диагностика поиска несущей способности", cap.DiagnosticLog
    If Not cap.LastSolver Is Nothing Then
        ReportBlock "Итерации внутреннего CSectionSolver для последней capacity-точки", cap.LastSolver.DiagnosticLog
    End If
End Sub

' Определяет, является ли силовая lambda-траектория действительно чисто
' осевой для выбора численного метода. Проверка выполняется не по сырым
' Mx/My из таблицы, а через CSectionLoadState: усилия переносятся к центру
' приведенного сечения модели прочности и проходят общий ZeroMoment-фильтр.
Private Function CapacityPathIsPureAxialForMethod(ByVal index As Long, _
        ByVal loadPath As CCapacityLoadPath) As Boolean
    If loadPath Is Nothing Then Exit Function
    If Not loadPath.ForceOnly Then Exit Function

    Dim referenceX As Double
    Dim referenceY As Double
    StrengthReferencePointForCombination index, referenceX, referenceY

    Dim loadState As CSectionLoadState
    Set loadState = LoadStateForCombination(index)

    CapacityPathIsPureAxialForMethod = MomentIsZeroForMethod(loadState.MxAboutPoint(referenceY), mConcreteDepthForMx) And _
        MomentIsZeroForMethod(loadState.MyAboutPoint(referenceX), mConcreteDepthForMy)
End Function

' Единая локальная проверка инженерного нуля для выбора метода capacity.
' Основное зануление уже делает CSectionLoadState, но повторная проверка
' через CMomentZeroFilter оставляет поведение устойчивым при прямых тестах
' или при отключенном фильтре.
Private Function MomentIsZeroForMethod(ByVal momentValue As Double, ByVal concreteDepth As Double) As Boolean
    If mMomentZeroFilter Is Nothing Then
        MomentIsZeroForMethod = (Abs(momentValue) <= 0.000000001)
    Else
        MomentIsZeroForMethod = mMomentZeroFilter.IsMomentEffectivelyZero(momentValue, concreteDepth)
    End If
End Function

' Поясняет в отчете, какая ветка поиска несущей фактически была принята.
' Низкоуровневый DiagnosticLog остается рядом, но здесь пользователь сразу
' видит смысл переключения Auto/UltimateStrain/LoadMultiplier по-русски.
Private Sub ReportCapacityMethodChoice(ByVal index As Long, ByVal cap As CCapacitySolver, _
        ByVal loadPath As CCapacityLoadPath, ByVal pureAxialForMethod As Boolean)
    If cap Is Nothing Then Exit Sub

    Dim strategy As String
    strategy = NormalizedCapacitySolutionStrategy()
    Dim methodName As String
    methodName = mResults(index).CapacitySolutionMethod

    If StrComp(strategy, "Auto", vbTextCompare) = 0 Then
        If StrComp(methodName, "UltimateStrain", vbTextCompare) = 0 Then
            ReportStep "Итог поиска несущей: Auto принял результат UltimateStrain, резервный переход на LoadMultiplier не потребовался."
        ElseIf StrComp(methodName, "LoadMultiplier", vbTextCompare) = 0 Then
            If pureAxialForMethod Then
                ReportStep "Итог поиска несущей: Auto использовал LoadMultiplier сразу, потому что траектория чисто осевая относительно центра приведенного сечения."
            Else
                ReportStep "Итог поиска несущей: Auto переключился на LoadMultiplier, потому что UltimateStrain не дал устойчивую предельную точку; подробная причина ниже в диагностике solver-а."
            End If
        End If
    ElseIf StrComp(strategy, "UltimateStrain", vbTextCompare) = 0 And _
            StrComp(methodName, "LoadMultiplier", vbTextCompare) = 0 Then
        If pureAxialForMethod Then
            ReportStep "Итог поиска несущей: фактически использован LoadMultiplier, так как траектория чисто осевая относительно центра приведенного сечения и не решается устойчиво прямым UltimateStrain."
        ElseIf loadPath.ForceOnly Then
            ReportStep "Итог поиска несущей: UltimateStrain не дал устойчивую предельную точку для силовой траектории с эксцентриситетом; фактически использован LoadMultiplier."
        Else
            ReportStep "Итог поиска несущей: фактически использован LoadMultiplier."
        End If
    End If
End Sub

' Выполняет один прямой StateSolution для заданного named-state профиля.
' Материальная модель строится по спецификации профиля, а управление retry и
' extension остается внутри CStateSolutionRunner.
Private Function RunProfileState(ByVal index As Long, ByVal stateType As ESectionStateType, _
        ByVal materialRole As ECalculationPurpose, ByVal spec As CMaterialModelSpec, _
        Optional ByVal useDirectStateExtension As Boolean = True, _
        Optional ByVal mergeIntoDirectStateStatus As Boolean = True) As CStateSolutionRunner
    Dim concreteState As CMaterialDiagram
    Dim steelState As CMaterialDiagram
    If useDirectStateExtension Then
        Set concreteState = mMaterialProvider.ConcreteStateMaterialFromSpec(spec)
        Set steelState = mMaterialProvider.SteelStateMaterialFromSpec(spec)
    Else
        Set concreteState = mMaterialProvider.ConcreteMaterialFromSpec(spec)
        Set steelState = mMaterialProvider.SteelMaterialFromSpec(spec)
    End If
    mResults(index).DirectStateDiagramExtensionEnabled = mMaterialProvider.DirectStateDiagramExtensionEnabled

    Dim runner As CStateSolutionRunner
    Set runner = New CStateSolutionRunner
    ConfigureStateRunner runner
    runner.Solve mSection, concreteState, steelState, mInputs(index).N, mInputs(index).Mx, mInputs(index).My, _
        (useDirectStateExtension And mResults(index).DirectStateDiagramExtensionEnabled)

    Dim service As CSectionSolver
    Set service = runner.ResultSolver
    If service Is Nothing Then
        If mergeIntoDirectStateStatus Then MergeDirectStateStatus index, mStatusPolicy.NumFail
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
        If mergeIntoDirectStateStatus Then MergeDirectStateStatus index, mStatusPolicy.NumFail
        ReportStep "Прямое НДС " & SectionStateTypeToText(stateType) & " не сошлось: " & service.StopReason
        Set RunProfileState = runner
        Exit Function
    End If

    mResults(index).ExtensionUsed = (mResults(index).ExtensionUsed Or runner.ExtensionUsed)
    Dim stateStatus As String
    If runner.ExtensionUsed Or Not runner.WithinPhysicalRange Then
        stateStatus = mStatusPolicy.Fail
        If mergeIntoDirectStateStatus Then MergeDirectStateStatus index, stateStatus
        ReportStep "Прямое НДС " & SectionStateTypeToText(stateType) & " найдено, но вышло за физический диапазон диаграммы материала."
    Else
        stateStatus = mStatusPolicy.OK
        If mergeIntoDirectStateStatus Then MergeDirectStateStatus index, stateStatus
    End If
    ReportValue "Статус прямого НДС", stateStatus
    ReportValue "Использовано продолжение диаграммы", IIf(runner.ExtensionUsed, "да", "нет")

    Dim stateResult As CSectionStateResult
    Set stateResult = New CSectionStateResult
    stateResult.InitializeFromSolver stateType, materialRole, spec, service, runner.ExtensionUsed, _
        stateStatus, runner.WithinPhysicalRange, runner.SolverCallCount, runner.DiagnosticLog
    mResults(index).AddState stateResult
    Set RunProfileState = runner
End Function

' Выполняет весь расчет раскрытия трещин как один запрошенный расчет профиля.
' Сначала обязательно проверяется факт образования нормальной трещины и
' сохраняется найденная пороговая точка, если она определима. Текущее
' CrackedState полного LC считается всегда: по нему выполняется проверка
' продольных трещин и записывается snapshot. Формула ширины раскрытия
' запускается только если нормальная трещина образована.
Private Sub RunCrackWidth(ByVal index As Long, ByVal profile As CCalculationProfile)
    Dim crack As CCrackWidthCalculator
    Set crack = New CCrackWidthCalculator
    ConfigureCrackCalculator crack
    ReportStep "Запускается проверка образования нормальной трещины."
    ReportValue "Путь образования трещины", mCrackFormationPath
    ReportValue "Стратегия поиска образования трещины", mCrackFormationSolutionStrategy
    ReportCrackFormationSearchStart

    Dim crackReferenceX As Double
    Dim crackReferenceY As Double
    CrackCentralReferencePoint profile, crackReferenceX, crackReferenceY

    Dim loadState As CSectionLoadState
    Set loadState = LoadStateForCombination(index)

    ' Центральная ветка трещин проверяется относительно центра тяжести
    ' приведенного сечения той материальной модели, по которой определяется
    ' Ncrc. Load.ReferenceOffset и сдвиг от бетонного центра учтены внутри
    ' CSectionLoadState, поэтому здесь нет ручной формулы N*dx/dy.
    crack.CheckFormationBeforeWidth mSection, mMaterialProvider, profile.CrackedStateSpec, _
        profile.CrackInitiationSpec, mInputs(index).N, mInputs(index).Mx, mInputs(index).My, _
        loadState, crackReferenceX, crackReferenceY
    mResults(index).CrackStatus = mStatusPolicy.CrackStatusFromCalculator(crack)
    mResults(index).StoreCrackResult crack
    If Not crack.Converged Then AppendCrackDiagnostic index, crack
    If Not crack.PreCrackState Is Nothing Then mResults(index).AddState crack.PreCrackState
    If Not crack.PostCrackState Is Nothing Then mResults(index).AddState crack.PostCrackState
    ReportValue "Фактический путь образования трещины", mResults(index).CrackFormationMethod
    ReportValue "lambda образования трещины", FormatNumberInvariant(mResults(index).CrackLambdaCrc)
    ReportValue "Ncrc", ReportForceText(mResults(index).CrackFormationNcrc)
    ReportValue "Mxy,crc", ReportMomentMagnitudeText(mResults(index).CrackMcrc)
    ReportValue "Причина остановки поиска трещины", crack.StopReason
    ReportCrackFormationChoice crack

    ReportStep "Запускается текущее CrackedState полного сочетания по профилю трещин."
    Dim runner As CStateSolutionRunner
    Set runner = RunProfileState(index, sstCrackedState, cpCrackedNDS, profile.CrackedStateSpec, True, False)
    If runner Is Nothing Or runner.ResultSolver Is Nothing Then
        mResults(index).CrackStatus = mStatusPolicy.NumFail
        mResults(index).LongitudinalCrackStatus = mStatusPolicy.NotApplicable
        ReportStep "Расчет трещин не продолжается: CrackedState не получен."
        GoTo ReportCrackResult
    End If
    If Not runner.Converged Then
        mResults(index).CrackStatus = mStatusPolicy.NumFail
        mResults(index).LongitudinalCrackStatus = mStatusPolicy.NotApplicable
        ReportStep "Расчет трещин не продолжается: CrackedState не сошелся."
        GoTo ReportCrackResult
    End If
    If runner.ExtensionUsed Or Not runner.WithinPhysicalRange Then
        mResults(index).CrackStatus = mStatusPolicy.Fail
        mResults(index).LongitudinalCrackStatus = mStatusPolicy.NotApplicable
        ReportStep "Расчет трещин не продолжается: равновесие CrackedState найдено только через продолжение диаграммы или вышло за физический диапазон."
        GoTo ReportCrackResult
    End If

    StoreLongitudinalCrackWithinCrackWidth index

    If Not crack.Converged Then
        ReportStep "Проверка образования трещины не завершилась корректно: " & crack.StopReason
        GoTo ReportCrackResult
    End If
    If Not crack.CrackFormed Then
        ReportStep "Ширина раскрытия не считается: нормальная трещина не образуется при текущей нагрузке."
        GoTo ReportCrackResult
    End If

    ReportStep "Запускается расчет ширины раскрытия нормальных трещин."
    crack.CalculateAfterFormationCheck runner.ResultSolver, mSection, mMaterialProvider, profile.CrackedStateSpec, _
        profile.CrackInitiationSpec, mInputs(index).N, mInputs(index).Mx, mInputs(index).My, _
        loadState, crackReferenceX, crackReferenceY
    mResults(index).CrackStatus = CrackWidthStatusAfterFormula(index, crack)
    mResults(index).StoreCrackResult crack
    If Not crack.Converged Then AppendCrackDiagnostic index, crack

ReportCrackResult:
    ReportValue "Статус нормальных трещин", mResults(index).CrackStatus
    ReportValue "Фактический путь образования трещины", mResults(index).CrackFormationMethod
    ReportValue "lambda образования трещины", FormatNumberInvariant(mResults(index).CrackLambdaCrc)
    ReportValue "Трещина образована", IIf(mResults(index).CrackFormed, "да", "нет")
    ReportValue "a_crc", ReportLengthText(mResults(index).CrackWidth)
    ReportValue "psi_s", FormatNumberInvariant(mResults(index).CrackPsiS)
    ReportValue "As", ReportAreaText(mResults(index).CrackAs)
    ReportValue "Abt", ReportAreaText(mResults(index).CrackAbt)
    ReportValue "ds_eq", ReportLengthText(mResults(index).CrackDsEquivalent)
    ReportValue "ls", ReportLengthText(mResults(index).CrackLs)
    ReportBlock "Диагностика расчета трещин", crack.DiagnosticLog
End Sub

' Возвращает пользовательский статус именно формульного этапа раскрытия
' нормальной трещины. NumFail допустим только если он уже виден в одном из
' НДС-состояний трещин: PreCrackState, PostCrackState или CrackedState.
' Если эти состояния сошлись, а калькулятор ширины оставил техническую
' NumericalFailure-причину, это не отдельная несходимость равновесия, поэтому
' для пользователя показывается FAIL, а подробность остается в txt-диагностике.
Private Function CrackWidthStatusAfterFormula(ByVal index As Long, _
        ByVal crack As CCrackWidthCalculator) As String
    Dim statusText As String
    statusText = mStatusPolicy.CrackStatusFromCalculator(crack)
    If statusText = mStatusPolicy.NumFail And Not HasNumFailCrackState(index) Then
        CrackWidthStatusAfterFormula = mStatusPolicy.Fail
    Else
        CrackWidthStatusAfterFormula = statusText
    End If
End Function

' Проверяет только те состояния трещин, где действительно решается равновесие
' сечения. Именно они имеют право поднимать пользовательский NumFail.
Private Function HasNumFailCrackState(ByVal index As Long) As Boolean
    HasNumFailCrackState = CrackStateStatusIsNumFail(index, sstPreCrackState) Or _
        CrackStateStatusIsNumFail(index, sstPostCrackState) Or _
        CrackStateStatusIsNumFail(index, sstCrackedState)
End Function

' Возвращает True, если конкретный named-state трещин сохранен как NumFail
' или содержит несошедшийся solver без более мягкого статуса.
Private Function CrackStateStatusIsNumFail(ByVal index As Long, _
        ByVal stateType As ESectionStateType) As Boolean
    Dim stateResult As CSectionStateResult
    Set stateResult = mResults(index).FindState(stateType)
    If stateResult Is Nothing Then Exit Function
    If mStatusPolicy.DisplayStatus(stateResult.Status) = mStatusPolicy.NumFail Then
        CrackStateStatusIsNumFail = True
    ElseIf Len(stateResult.Status) = 0 And Not stateResult.Converged Then
        CrackStateStatusIsNumFail = True
    End If
End Function

' Поясняет в txt-отчете стартовую логику поиска образования нормальной трещины.
' Низкоуровневый DiagnosticLog ниже сохраняет точные попытки, а эта строка
' нужна пользователю как быстрый ответ, что именно программа собирается делать.
Private Sub ReportCrackFormationSearchStart()
    If StrComp(mCrackFormationPath, "Auto", vbTextCompare) = 0 Then
        ReportStep "Трещинообразование: режим Auto. Программа пробует пути по порядку: " & _
            ChrW$(&H3BB) & "*Mxy, затем " & ChrW$(&H3BB) & "*N, затем " & _
            ChrW$(&H3BB) & "*NMxy. Если путь дает только резервный результат с " & _
            ChrW$(&H3C8) & "s = 1 без найденной точки образования трещины, программа переходит к следующему пути."
    Else
        ReportStep "Трещинообразование: используется фиксированный путь " & mCrackFormationPath & _
            "; Auto-перебор других путей не выполняется."
    End If
End Sub

' Поясняет итог выбора пути образования трещины для текущего LC. Это не заменяет
' DiagnosticLog solver-а, а переводит ключевой результат Auto/резервного
' нормальный инженерный язык прямо в отчете.
Private Sub ReportCrackFormationChoice(ByVal crack As CCrackWidthCalculator)
    If crack Is Nothing Then Exit Sub

    If Len(crack.CrackFormationMethod) > 0 Then
        ReportStep "Трещинообразование: принята точка по пути " & _
            crack.CrackFormationMethod & ". По этой точке записаны Ncrc/Mxy,crc и состояния PreCrackState/PostCrackState, если они применимы."
    ElseIf crack.CrackFormed And crack.Converged Then
        ReportStep "Трещинообразование: нормальная трещина считается образованной, но выбранный путь не дал отдельную расчетную точку образования трещины. Для этой проверки принят резервный вариант " & _
            ChrW$(&H3C8) & "s = 1; причина указана в CrackFormationStopReason и диагностике ниже."
    ElseIf crack.Converged Then
        ReportStep "Трещинообразование: проверка завершена, нормальная трещина при текущем сочетании не образуется."
    Else
        ReportStep "Трещинообразование: точку образования трещины найти не удалось; дальнейшая причина приведена в диагностике ниже."
    End If
End Sub

' Добавляет в компактный batch-журнал причину несходимости трещин.
' Подробный txt-отчет пишет тот же DiagnosticLog отдельно, а здесь строка
' нужна unit/regression-тестам и пользователю, когда общий статус только NumFail.
Private Sub AppendCrackDiagnostic(ByVal index As Long, ByVal crack As CCrackWidthCalculator)
    If crack Is Nothing Then Exit Sub

    mDiagnosticLog = mDiagnosticLog & "combination=" & mInputs(index).CombinationID & _
        "; crackStopReason=" & crack.StopReason & vbCrLf
    If Len(crack.DiagnosticLog) > 0 Then
        mDiagnosticLog = mDiagnosticLog & crack.DiagnosticLog
    End If
End Sub

' Создает объект lambda-траектории для одной строки LC.
' Batch передает сюда только уже нормализованные усилия и точку приложения
' нагрузки, а все подробности интерпретации lambda*Mx/lambda*N остаются в CCapacityLoadPath.
Private Function BuildLoadPathForCombination(ByVal index As Long) As CCapacityLoadPath
    Dim loadPath As CCapacityLoadPath
    Set loadPath = New CCapacityLoadPath
    loadPath.InitializeFromLoadState mInputs(index).CapacityLoadPath, LoadStateForCombination(index)
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
    mResults(index).CapacityStateN = 0#
    mResults(index).CapacityStateMx = 0#
    mResults(index).CapacityStateMy = 0#
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
        If Not cap.LastSolver Is Nothing Then
            mResults(index).CapacityStateN = cap.LastSolver.Nint
            mResults(index).CapacityStateMx = cap.LastSolver.Mxint
            mResults(index).CapacityStateMy = cap.LastSolver.Myint
            StoreCapacityUltimateAtLoadPoint index, cap.LastSolver.Nint, cap.LastSolver.Mxint, cap.LastSolver.Myint
        Else
            Dim targetN As Double
            Dim targetMx As Double
            Dim targetMy As Double
            loadPath.Target cap.LambdaUltimate, targetN, targetMx, targetMy
            StoreCapacityUltimateAtLoadPoint index, targetN, targetMx, targetMy
        End If
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
                mResults(index).CapacityStatus, True, 0, cap.DiagnosticLog
            mResults(index).AddState stateResult
        End If
    End If
End Sub

' Сохраняет полный предельный вектор CapacityState в пользовательской точке
' приложения нагрузки. Внутри solver хранит моменты относительно начала
' CSectionModel, поэтому для вывода Nult/Mxult/Myult используем общий класс
' переносов нагрузки, а не локальную формулу writer-а.
Private Sub StoreCapacityUltimateAtLoadPoint(ByVal index As Long, ByVal nValue As Double, _
        ByVal mxAboutOrigin As Double, ByVal myAboutOrigin As Double)
    Dim loadState As CSectionLoadState
    Set loadState = New CSectionLoadState
    loadState.InitializeFromPointMoments nValue, mxAboutOrigin, myAboutOrigin, _
        0#, 0#, mLoadReferenceX, mLoadReferenceY, mMomentZeroFilter, _
        mConcreteDepthForMx, mConcreteDepthForMy

    mResults(index).NUltimate = loadState.N
    mResults(index).MxUltimate = loadState.LoadPointMx
    mResults(index).MyUltimate = loadState.LoadPointMy
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
    mResults(index).StabilityStatus = mStatusPolicy.InputErr
    mResults(index).CapacityLimitState = reason
    mResults(index).CapacitySolutionMethod = vbNullString
    mResults(index).CapacityPathResolved = CAPACITY_PATH_NONE
    mResults(index).CrackStatus = mStatusPolicy.NotApplicable
    mResults(index).LongitudinalCrackStatus = mStatusPolicy.NotApplicable
    If Not mMaterialProvider Is Nothing Then
        mResults(index).DirectStateDiagramExtensionEnabled = mMaterialProvider.DirectStateDiagramExtensionEnabled
    End If
    mResults(index).GoverningUtil = 0#
    mResults(index).ResultComment = reason
    mResults(index).StoreOverallMeta mResults(index).OverallMeta
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

    ReportValue "Статус продольных трещин", mResults(index).LongitudinalCrackStatus
    If mResults(index).LongitudinalCrackStatus <> mStatusPolicy.NotApplicable Then
        ReportValue "sigma_c,max", ReportStressText(maxCompressionStress)
        ReportValue "Rb,mc2", ReportStressText(rbMc2)
        ReportValue "Использование по продольным трещинам", FormatNumberInvariant(utilization)
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
    crack.SigmaSCrcAveragingMode = mCrackSigmaSCrcAveragingMode
    crack.Phi1 = mCrackPhi1
    crack.Phi2 = mCrackPhi2
    crack.Phi3Mode = mCrackPhi3Mode
    crack.Phi3 = mCrackPhi3
    crack.PsiS = mCrackPsiS
    crack.TensionZoneMode = mCrackTensionZoneMode
    crack.CoverMode = mCrackCoverDistanceMode
    crack.CrackFormationPath = mCrackFormationPath
    crack.CrackFormationSolutionStrategy = mCrackFormationSolutionStrategy
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
    If mMaterialProvider Is Nothing Then Err.Raise vbObjectError + 3922, "CBatchSectionCalculator", "Пакетный расчетчик не инициализирован поставщиком материалов."
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
    ReportValue "Итоговый статус", mResults(index).Status
    ReportValue "Статус прямого НДС", mResults(index).DirectStateStatus
    ReportValue "Использовано продолжение диаграммы", YesNoRu(mResults(index).ExtensionUsed)
    ReportValue "Статус несущей способности", mResults(index).CapacityStatus
    ReportValue "Статус нормальных трещин", mResults(index).CrackStatus
    ReportValue "Статус продольных трещин", mResults(index).LongitudinalCrackStatus
    ReportValue "Использование по продольным трещинам", FormatNumberInvariant(mResults(index).LongitudinalCrackUtil)
    ReportValue "Статус устойчивости", mResults(index).StabilityStatus
    ReportValue "Использование устойчивости", FormatNumberInvariant(mResults(index).StabilityUtil)
    ReportValue "Определяющее использование", FormatNumberInvariant(mResults(index).GoverningUtil)
End Sub

' Добавляет смысловой заголовок в txt-отчет выполнения, если отчет включен
' пользователем. Вызов безопасен при выключенном отчете.
Private Sub ReportSection(ByVal title As String)
    If mReport Is Nothing Then Exit Sub
    If Not mReport.Enabled Then Exit Sub
    mReport.AddSection title
End Sub

' Добавляет строку-этап в txt-отчет выполнения. Используется orchestration-ом,
' чтобы показать пользователю, какая расчетная ветка запускается сейчас.
Private Sub ReportStep(ByVal text As String)
    If mReport Is Nothing Then Exit Sub
    If Not mReport.Enabled Then Exit Sub
    mReport.AddStep text
End Sub

' Добавляет пару "параметр - значение" в txt-отчет. Метод держит всю проверку
' включенности отчета в одном месте, чтобы расчетный код не дублировал guard.
Private Sub ReportValue(ByVal label As String, ByVal valueText As String)
    If mReport Is Nothing Then Exit Sub
    If Not mReport.Enabled Then Exit Sub
    mReport.AddValue label, valueText
End Sub

' Форматирует длину для execution_report в той же выходной единице, что и
' табличный Results. Внутри batch все еще хранит миллиметры.
Private Function ReportLengthText(ByVal value As Double) As String
    If mReportUnits Is Nothing Then
        ReportLengthText = FormatNumberInvariant(value) & " мм"
    Else
        ReportLengthText = FormatNumberInvariant(mReportUnits.InternalLengthToOutput(value)) & " " & mReportUnits.OutputLengthUnit
    End If
End Function

' Форматирует площадь в пользовательской выходной единице площади.
Private Function ReportAreaText(ByVal value As Double) As String
    If mReportUnits Is Nothing Then
        ReportAreaText = FormatNumberInvariant(value) & " мм2"
    Else
        ReportAreaText = FormatNumberInvariant(mReportUnits.InternalAreaToOutput(value)) & " " & mReportUnits.OutputAreaUnit
    End If
End Function

' Форматирует обычную знаковую продольную силу N с учетом Output.SignConvention.N.
Private Function ReportForceText(ByVal value As Double) As String
    If mReportUnits Is Nothing Then
        ReportForceText = FormatNumberInvariant(value) & " Н"
    Else
        ReportForceText = FormatNumberInvariant(mReportUnits.InternalForceToOutput(value)) & " " & mReportUnits.OutputForceUnit
    End If
End Function

' Форматирует положительную величину сжимающей силы, например Ncr или Nult.
' Внутренний знак восстанавливается как сжатие, чтобы пользователь видел
' тот же знак, что в таблице устойчивости.
Private Function ReportCompressionForceText(ByVal compressionMagnitude As Double) As String
    If mReportUnits Is Nothing Then
        ReportCompressionForceText = FormatNumberInvariant(-Abs(compressionMagnitude)) & " Н"
    Else
        ReportCompressionForceText = FormatNumberInvariant(mReportUnits.InternalForceToOutput(-Abs(compressionMagnitude))) & _
            " " & mReportUnits.OutputForceUnit
    End If
End Function

' Форматирует уже подписанную сжимающую силу, где положительное значение
' означает сжатие в локальной нормативной формуле устойчивости.
Private Function ReportSignedCompressionForceText(ByVal signedCompression As Double) As String
    If mReportUnits Is Nothing Then
        ReportSignedCompressionForceText = FormatNumberInvariant(-signedCompression) & " Н"
    Else
        ReportSignedCompressionForceText = FormatNumberInvariant(mReportUnits.InternalForceToOutput(-signedCompression)) & _
            " " & mReportUnits.OutputForceUnit
    End If
End Function

' Форматирует Mx с пользовательским знаком Output.SignConvention.Mx.
Private Function ReportMxText(ByVal value As Double) As String
    If mReportUnits Is Nothing Then
        ReportMxText = FormatNumberInvariant(value) & " Н*мм"
    Else
        ReportMxText = FormatNumberInvariant(mReportUnits.InternalMomentMxToOutput(value)) & " " & mReportUnits.OutputMomentUnit
    End If
End Function

' Форматирует My с пользовательским знаком Output.SignConvention.My.
Private Function ReportMyText(ByVal value As Double) As String
    If mReportUnits Is Nothing Then
        ReportMyText = FormatNumberInvariant(value) & " Н*мм"
    Else
        ReportMyText = FormatNumberInvariant(mReportUnits.InternalMomentMyToOutput(value)) & " " & mReportUnits.OutputMomentUnit
    End If
End Function

' Форматирует скалярный момент без смены знака: используется для M1/M2,
' Mxy,crc и других величин, где в отчете нужен модуль, а не глобальный Mx/My.
Private Function ReportMomentMagnitudeText(ByVal value As Double) As String
    If mReportUnits Is Nothing Then
        ReportMomentMagnitudeText = FormatNumberInvariant(value) & " Н*мм"
    Else
        ReportMomentMagnitudeText = FormatNumberInvariant(mReportUnits.InternalMomentMagnitudeToOutput(value)) & _
            " " & mReportUnits.OutputMomentUnit
    End If
End Function

' Форматирует напряжение/модуль в пользовательской выходной единице напряжений.
Private Function ReportStressText(ByVal value As Double) As String
    If mReportUnits Is Nothing Then
        ReportStressText = FormatNumberInvariant(value) & " МПа"
    Else
        ReportStressText = FormatNumberInvariant(mReportUnits.InternalStressToOutput(value)) & " " & mReportUnits.OutputStressUnit
    End If
End Function

' Форматирует жесткость D устойчивости. Размерность задается как сила*длина2
' и использует ту же пару output-единиц силы и длины, что Results.
Private Function ReportStiffnessText(ByVal value As Double) As String
    If mReportUnits Is Nothing Then
        ReportStiffnessText = FormatNumberInvariant(value) & " Н*мм2"
    Else
        ReportStiffnessText = FormatNumberInvariant(mReportUnits.InternalForceLengthSquaredToOutput(value)) & _
            " " & mReportUnits.OutputForceLengthSquaredUnit
    End If
End Function

' Добавляет многострочный диагностический блок в txt-отчет и перед записью
' переводит низкоуровневую диагностику solver-а в более читаемый вид.
Private Sub ReportBlock(ByVal title As String, ByVal body As String)
    If mReport Is Nothing Then Exit Sub
    If Not mReport.Enabled Then Exit Sub
    mReport.AddBlock title, HumanReadableReportBlock(title, body)
End Sub

' Преобразует низкоуровневую диагностику solver-ов в форму, пригодную для
' пользовательского txt-отчета. Внутренние DiagnosticLog остаются подробными
' и машинно-проверяемыми, но отчет не должен забиваться сотнями строк
' initialGuess/lambda_crc_iter без пояснений.
Private Function HumanReadableReportBlock(ByVal title As String, ByVal body As String) As String
    Dim text As String
    text = body
    If Len(text) = 0 Then Exit Function

    If InStr(1, title, "Диагностика расчета трещин", vbTextCompare) > 0 Then
        text = CompactCrackDiagnosticText(text)
    End If
    If InStr(1, title, "Сценарий прямого StateSolution", vbTextCompare) > 0 Or _
            InStr(1, title, "Итерации прямого CSectionSolver", vbTextCompare) > 0 Or _
            InStr(1, title, "Итерации внутреннего CSectionSolver", vbTextCompare) > 0 Then
        text = CompactStateSolutionDiagnosticText(text)
    End If
    If InStr(1, title, "Диагностика поиска несущей способности", vbTextCompare) > 0 Then
        text = CompactCapacityDiagnosticText(text)
    End If

    HumanReadableReportBlock = text
End Function

' Приводит подробный журнал трещин к читаемому виду: смысловые строки и
' итерации lambda_crc оставляет по порядку, а повторяющиеся warm-start
' сообщения заменяет короткой сводкой.
Private Function CompactCrackDiagnosticText(ByVal body As String) As String
    Dim lines() As String
    lines = Split(body, vbCrLf)

    Dim result As String
    Dim warmStartCount As Long
    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        Dim lineText As String
        lineText = Trim$(lines(i))
        If Len(lineText) = 0 Then
            ' Пустые строки внутри diagnostic-блока не несут смысла.
        ElseIf Left$(lineText, Len("lambda_crc_iter=")) = "lambda_crc_iter=" Then
            AppendReportLine result, TranslateDiagnosticLine(lineText)
        ElseIf Left$(lineText, Len("lambda_crc_path_iter=")) = "lambda_crc_path_iter=" Then
            AppendReportLine result, TranslateDiagnosticLine(lineText)
        ElseIf lineText = "initialGuess=warmStartSolver" Then
            warmStartCount = warmStartCount + 1
        Else
            AppendReportLine result, TranslateDiagnosticLine(lineText)
        End If
    Next i

    If warmStartCount > 0 Then _
        AppendReportLine result, "Для повторных НДС использовалась стартовая плоскость из предыдущего близкого решения; повторов: " & CStr(warmStartCount) & "."

    CompactCrackDiagnosticText = result
End Function

' Сжимает diagnostic CStateSolutionRunner. Особенно важно убрать десятки
' одинаковых warm-start строк: они полезны разработчику, но не пользователю.
Private Function CompactStateSolutionDiagnosticText(ByVal body As String) As String
    Dim lines() As String
    lines = Split(body, vbCrLf)

    Dim result As String
    Dim warmStartCount As Long
    Dim activeGuessCount As Long
    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        Dim lineText As String
        lineText = Trim$(lines(i))
        If Len(lineText) = 0 Then
            ' Пропускаем.
        ElseIf Left$(lineText, Len("step=")) = "step=" Then
            AppendReportLine result, TranslateDiagnosticLine(lineText)
        ElseIf lineText = "initialGuess=warmStartSolver" Then
            warmStartCount = warmStartCount + 1
        ElseIf lineText = "initialGuess=directStateActiveSet" Then
            activeGuessCount = activeGuessCount + 1
        Else
            AppendReportLine result, TranslateDiagnosticLine(lineText)
        End If
    Next i

    If warmStartCount > 0 Then _
        AppendReportLine result, "Стартовая плоскость: использовано предыдущее близкое решение; повторов: " & CStr(warmStartCount) & "."
    If activeGuessCount > 0 Then _
        AppendReportLine result, "Стартовая плоскость: подобрана по активной зоне сечения; повторов: " & CStr(activeGuessCount) & "."

    CompactStateSolutionDiagnosticText = result
End Function

' Сжимает диагностический блок поиска несущей способности. Полный список
' lambda-точек может быть большим, но для инженерной диагностики полезно
' видеть каждую проверенную точку и возможный retry НДС.
Private Function CompactCapacityDiagnosticText(ByVal body As String) As String
    Dim lines() As String
    lines = Split(body, vbCrLf)

    Dim result As String
    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        Dim lineText As String
        lineText = Trim$(lines(i))
        If Len(lineText) = 0 Then
            ' Пропускаем.
        ElseIf Left$(lineText, Len("probe=")) = "probe=" Then
            AppendReportLine result, TranslateDiagnosticLine(lineText)
        Else
            AppendReportLine result, TranslateDiagnosticLine(lineText)
        End If
    Next i

    CompactCapacityDiagnosticText = result
End Function

' Переводит короткие машинные diagnostic-строки в более понятные фразы.
Private Function TranslateDiagnosticLine(ByVal lineText As String) As String
    Dim translated As String
    translated = lineText
    If translated = "ultimatePath pureAxial redirected=LoadMultiplier" Or _
            translated = "Силовая траектория UltimateStrain перенаправлена на LoadMultiplier." Then
        TranslateDiagnosticLine = "Силовая траектория lambda*N перенаправлена на LoadMultiplier, потому что UltimateStrain для нее вырождается."
        Exit Function
    End If
    translated = Replace$(translated, "ultimatePath iter=", "итерация UltimateStrain по общей lambda-траектории=")
    translated = Replace$(translated, "ultimate iter=", "итерация UltimateStrain по моментной траектории=")
    translated = Replace$(translated, "lambda_crc=", "Найден lambda_crc=")
    translated = Replace$(translated, "lambda_crc_iter=", "итерация прямого поиска lambda_crc=")
    translated = Replace$(translated, "lambda_crc_path_iter=", "итерация UltimateStrain по lambda-траектории lambda_crc=")
    translated = Replace$(translated, "lambda_crc_load_probe=", "точка LoadMultiplier lambda_crc=")
    translated = Replace$(translated, "iterations=", "итераций=")
    translated = Replace$(translated, "method=", "метод=")
    translated = Replace$(translated, "formation=Ncrc", "Образование трещины найдено по центральному растяжению Ncrc")
    translated = Replace$(translated, "search=", "Поиск: ")
    translated = Replace$(translated, "searchMethod=", "метод поиска=")
    translated = Replace$(translated, "secantGuard=insideBracket", "защита Secant=точка возвращена внутрь скобки")
    translated = Replace$(translated, "rawLambda=", "исходная lambda=")
    translated = Replace$(translated, "guardedLambda=", "защищенная lambda=")
    translated = Replace$(translated, "bracketLow=", "нижняя граница=")
    translated = Replace$(translated, "bracketHigh=", "верхняя граница=")
    translated = Replace$(translated, "probe=", "точка lambda=")
    translated = Replace$(translated, "step=", "шаг=")
    translated = Replace$(translated, "iter=", "итерация=")
    translated = Replace$(translated, "directStateRetry=", "повтор прямого НДС=")
    translated = Replace$(translated, "extensionRetry=", "повтор после выхода за диаграмму=")
    translated = Replace$(translated, "directStateActiveSet", "активная зона прямого НДС")
    translated = Replace$(translated, "warmStartSolver", "предыдущее близкое решение")
    translated = Replace$(translated, "direct target without line search", "целевая нагрузка без line search")
    translated = Replace$(translated, "direct target", "целевая нагрузка")
    translated = Replace$(translated, "stepped", "ступенчатая нагрузка")
    translated = Replace$(translated, "attempt=", "повтор НДС=")
    translated = Replace$(translated, "converged=True", "сошлось=да")
    translated = Replace$(translated, "converged=False", "сошлось=нет")
    translated = Replace$(translated, "state=", "состояние=")
    translated = Replace$(translated, "stop=", "причина=")
    translated = Replace$(translated, "lineSearchUsed=", "line search использован=")
    translated = Replace$(translated, "lineSearchIterations=", "итераций line search=")
    translated = Replace$(translated, "internalForceEvaluationCount=", "оценок внутренних усилий=")
    translated = Replace$(translated, "internalNewtonCallCount=", "вызовов Newton=")
    translated = Replace$(translated, "result=converged", "результат=сошлось")
    translated = Replace$(translated, "result=failed", "результат=не сошлось")
    translated = Replace$(translated, "rPath1=", "невязка пути 1=")
    translated = Replace$(translated, "rPath2=", "невязка пути 2=")
    translated = Replace$(translated, "rLimit=", "невязка предельной деформации=")
    translated = Replace$(translated, "rDir=", "невязка направления=")
    translated = Replace$(translated, "residual=", "невязка образования трещины=")
    translated = Replace$(translated, "rN=", "невязка N=")
    translated = Replace$(translated, "rMomentDir=", "невязка направления момента=")
    translated = Replace$(translated, "eps_bt_ult=", "eps_bt,ult=")
    translated = Replace$(translated, "residualN=", "невязка N=")
    translated = Replace$(translated, "residualMx=", "невязка Mx=")
    translated = Replace$(translated, "residualMy=", "невязка My=")
    translated = Replace$(translated, "critical=", "критический элемент=")
    translated = Replace$(translated, "minConcreteStrain=", "мин. деформация бетона=")
    translated = Replace$(translated, "maxConcreteStrain=", "макс. деформация бетона=")
    translated = Replace$(translated, "maxAbsSteelStrain=", "макс. модуль деформации арматуры=")
    TranslateDiagnosticLine = translated
End Function

' Единая короткая русская запись булевых значений для отчета.
Private Function YesNoRu(ByVal value As Boolean) As String
    If value Then
        YesNoRu = "да"
    Else
        YesNoRu = "нет"
    End If
End Function

' Аккуратно добавляет строку в собираемый многострочный блок отчета.
Private Sub AppendReportLine(ByRef target As String, ByVal lineText As String)
    If Len(lineText) = 0 Then Exit Sub
    If Len(target) > 0 Then target = target & vbCrLf
    target = target & lineText
End Sub

' Русское название нормативного блока устойчивости для txt-отчета.
Private Function StabilityCodeDisplayName(ByVal codeText As String) As String
    Select Case UCase$(Trim$(codeText))
        Case "SP63"
            StabilityCodeDisplayName = "СП 63"
        Case "SP35"
            StabilityCodeDisplayName = "СП 35"
        Case Else
            StabilityCodeDisplayName = codeText
    End Select
End Function

' Русское описание внутренней ветви устойчивости. Расчетные идентификаторы
' остаются в коде, но отчет должен читаться без знания этих имен.
Private Function StabilityBranchDisplayName(ByVal branchText As String) As String
    Select Case LCase$(Trim$(branchText))
        Case "nocompression"
            StabilityBranchDisplayName = "устойчивость не требуется: нет сжатия"
        Case "sp63"
            StabilityBranchDisplayName = "СП 63: расчет eta и Ncr"
        Case "sp35-eta"
            StabilityBranchDisplayName = "СП 35: ветвь ec > r через eta и Ncr"
        Case "sp35-table"
            StabilityBranchDisplayName = "СП 35: табличная ветвь ec <= r через Nult,stab"
        Case "sp35-mixed"
            StabilityBranchDisplayName = "СП 35: плоскости попали в разные ветви"
        Case Else
            StabilityBranchDisplayName = branchText
    End Select
End Function

' Делает PASS/FAIL в отчете понятным русским текстом.
Private Function StabilityPassDisplayName(ByVal passed As Boolean) As String
    If passed Then
        StabilityPassDisplayName = "проверка пройдена"
    Else
        StabilityPassDisplayName = "проверка не пройдена"
    End If
End Function

' Утилитарный максимум без WorksheetFunction, чтобы не тащить Excel в hot path.
Private Function MaxDouble(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then MaxDouble = a Else MaxDouble = b
End Function

' Утилитарный минимум без WorksheetFunction; используется в расчетном слое,
' где нельзя вводить лишнюю зависимость от Excel-объектов.
Private Function MinDouble(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then MinDouble = a Else MinDouble = b
End Function

' Утилитарный максимум для Long без WorksheetFunction и без неявного Variant.
Private Function MaxLong(ByVal a As Long, ByVal b As Long) As Long
    If a > b Then MaxLong = a Else MaxLong = b
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

' Выбирает Worst-сочетание по пользовательскому критерию верхней сводки.
' Этот же индекс затем читают Plot.LoadCase = Worst и AutoCAD.Export.CombinationID = Worst.
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
    candidateSafety = WorstSafetyFactor(candidateIndex)
    If candidateSafety <= 0# Then Exit Function

    If currentIndex <= 0 Then
        IsBetterWorstCandidate = True
    Else
        currentSafety = WorstSafetyFactor(currentIndex)
        IsBetterWorstCandidate = (currentSafety <= 0# Or candidateSafety < currentSafety)
    End If
End Function

' Возвращает запас LC по критерию General.WorstCombinationCriterion.
' StrengthCapacity оставлен значением по умолчанию, чтобы сохранить старое
' поведение Worst до явного выбора пользователем другого критерия.
Public Function WorstSafetyFactor(ByVal index As Long) As Double
    Select Case NormalizedWorstCombinationCriterion()
        Case "StrengthStrain"
            WorstSafetyFactor = StrengthStrainReserve(index)
        Case "StrengthCapacity"
            WorstSafetyFactor = StrengthCapacityReserve(index)
        Case "Cracks"
            WorstSafetyFactor = PositiveMin(CrackReserveFactor(index), LongitudinalCrackReserveFactor(index))
        Case "Stability"
            WorstSafetyFactor = StabilityReserveForSummary(index)
    End Select
End Function

' Нормализует пользовательский критерий выбора Worst. Некорректная строка не
' роняет расчет, а возвращает прежний консервативный смысл по capacity.
Private Function NormalizedWorstCombinationCriterion() As String
    Select Case LCase$(Trim$(mWorstCombinationCriterion))
        Case "strengthstrain"
            NormalizedWorstCombinationCriterion = "StrengthStrain"
        Case "cracks"
            NormalizedWorstCombinationCriterion = "Cracks"
        Case "stability"
            NormalizedWorstCombinationCriterion = "Stability"
        Case Else
            NormalizedWorstCombinationCriterion = "StrengthCapacity"
    End Select
End Function

' Минимальный запас по предельным деформациям модели прочности.
' Берутся только реально существующие зоны: если растянутого бетона или
' сжатой арматуры в состоянии нет, эта составляющая не участвует в минимуме.
Public Function StrengthStrainReserve(ByVal index As Long) As Double
    Dim stateResult As CSectionStateResult
    Set stateResult = mResults(index).FindState(sstStrengthState)
    If stateResult Is Nothing Then Exit Function
    If Not mStatusPolicy.IsFinished(stateResult.Status) Then Exit Function

    Dim reserve As Double
    reserve = 1E+100
    AddStrainReserve reserve, stateResult.HasConcreteCompression, _
        StrengthConcreteCompressionLimitForCombination(index), stateResult.MinConcreteStrain
    If StrengthConcreteTensionLimitEnabledForCombination(index) Then _
        AddStrainReserve reserve, stateResult.HasConcreteTension, _
            StrengthConcreteTensionLimitForCombination(index), stateResult.MaxConcreteStrain
    AddStrainReserve reserve, stateResult.HasSteelCompression, _
        StrengthSteelCompressionLimitForCombination(index), stateResult.MinSteelStrain
    AddStrainReserve reserve, stateResult.HasSteelTension, _
        StrengthSteelTensionLimitForCombination(index), stateResult.MaxSteelStrain
    If reserve < 1E+90 Then StrengthStrainReserve = reserve
End Function

' Добавляет одну деформационную составляющую в общий минимум запаса.
Private Sub AddStrainReserve(ByRef reserve As Double, ByVal hasValue As Boolean, _
        ByVal limitValue As Double, ByVal actualValue As Double)
    If Not hasValue Then Exit Sub
    If Abs(limitValue) <= 0.000000000000001 Then Exit Sub
    If Abs(actualValue) <= 0.000000000000001 Then Exit Sub
    reserve = MinDouble(reserve, Abs(limitValue) / Abs(actualValue))
End Sub

' Запас по несущей способности: фактически найденный lambda выбранной
' траектории. Это ровно прежний критерий strength-Worst.
Public Function StrengthCapacityReserve(ByVal index As Long) As Double
    If mStatusPolicy.IsFinished(mResults(index).CapacityStatus) And mResults(index).LambdaCapacity > 0# Then _
        StrengthCapacityReserve = mResults(index).LambdaCapacity
End Function

' Запас по нормальным трещинам как обратная величина к использованию a_crc.
Public Function CrackReserveFactor(ByVal index As Long) As Double
    If mStatusPolicy.IsFinished(mResults(index).CrackStatus) And mResults(index).CrackUtil > 0# Then _
        CrackReserveFactor = 1# / mResults(index).CrackUtil
End Function

' Запас по продольным трещинам как обратная величина к sigma_b/Rb,mc2.
Public Function LongitudinalCrackReserveFactor(ByVal index As Long) As Double
    If mStatusPolicy.IsFinished(mResults(index).LongitudinalCrackStatus) And _
            mResults(index).LongitudinalCrackUtil > 0# Then _
        LongitudinalCrackReserveFactor = 1# / mResults(index).LongitudinalCrackUtil
End Function

' Минимальный запас по устойчивости с учетом пользовательского лимита СП 35
' для ветви N/Ncr <= Stability.SP35.NOverNcrLimit.
Public Function StabilityReserveForSummary(ByVal index As Long) As Double
    If Not mStatusPolicy.IsFinished(mResults(index).StabilityStatus) Then Exit Function
    StabilityReserveForSummary = PositiveMin( _
        StabilityPlaneReserveForSummary(index, 1), _
        StabilityPlaneReserveForSummary(index, 2))
    If StabilityReserveForSummary <= 0# And mResults(index).StabilityReserveFactor > 0# Then _
        StabilityReserveForSummary = mResults(index).StabilityReserveFactor
End Function

' Возвращает запас одной главной плоскости устойчивости. Для СП 35 eta-ветви
' в расчетный запас входит не просто Ncr/N, а проверяемое условие N/Ncr <= limit.
Private Function StabilityPlaneReserveForSummary(ByVal index As Long, ByVal planeIndex As Long) As Double
    Dim applicable As Boolean
    Dim rawReserve As Double
    Dim branchText As String
    If planeIndex = 1 Then
        applicable = mResults(index).StabilityPlaneApplicable1
        rawReserve = mResults(index).StabilityReserve1
        branchText = mResults(index).StabilityPlaneBranch1
    Else
        applicable = mResults(index).StabilityPlaneApplicable2
        rawReserve = mResults(index).StabilityReserve2
        branchText = mResults(index).StabilityPlaneBranch2
    End If
    If Not applicable Then Exit Function
    If rawReserve <= 0# Then Exit Function

    StabilityPlaneReserveForSummary = rawReserve
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

' Возвращает 1-based позицию LC внутри data-строк rngLoadCombinations.
' Это служебная величина только для writer-ов Results; расчетный индекс batch остается компактным.
Public Property Get SourceDataOffset(ByVal index As Long) As Long
    SourceDataOffset = mInputs(index).SourceDataOffset
End Property

Public Property Get CombinationName(ByVal index As Long) As String
    CombinationName = mInputs(index).Comment
End Property

Public Property Get ProfileId(ByVal index As Long) As String
    If Len(mResults(index).ProfileId) > 0 Then
        ProfileId = mResults(index).ProfileId
    Else
        ProfileId = mInputs(index).ProfileId
    End If
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

Public Property Get CapacityStateN(ByVal index As Long) As Double
    CapacityStateN = mResults(index).CapacityStateN
End Property

Public Property Get CapacityStateMx(ByVal index As Long) As Double
    CapacityStateMx = mResults(index).CapacityStateMx
End Property

Public Property Get CapacityStateMy(ByVal index As Long) As Double
    CapacityStateMy = mResults(index).CapacityStateMy
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

' Возвращает пользовательский лимит N/Ncr для eta-ветви СП 35.
' Writer-ы используют это значение в заголовках, чтобы формула запаса
' в Results совпадала с фактически выбранной настройкой Config.
Public Property Get StabilitySP35NOverNcrLimit() As Double
    StabilitySP35NOverNcrLimit = mStabilitySP35NOverNcrLimit
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

Public Property Get CrackFormationMethod(ByVal index As Long) As String
    CrackFormationMethod = mResults(index).CrackFormationMethod
End Property

Public Property Get CrackNcrc(ByVal index As Long) As Double
    CrackNcrc = mResults(index).CrackNcrc
End Property

Public Property Get CrackFormationNcrc(ByVal index As Long) As Double
    CrackFormationNcrc = mResults(index).CrackFormationNcrc
End Property

Public Property Get CrackMcrc(ByVal index As Long) As Double
    CrackMcrc = mResults(index).CrackMcrc
End Property

Public Property Get CrackAred(ByVal index As Long) As Double
    CrackAred = mResults(index).CrackAred
End Property

Public Property Get CrackRbtSer(ByVal index As Long) As Double
    CrackRbtSer = mResults(index).CrackRbtSer
End Property

Public Property Get CrackTensionRebarCount(ByVal index As Long) As Long
    CrackTensionRebarCount = mResults(index).CrackTensionRebarCount
End Property

Public Property Get CrackTensionRebarIds(ByVal index As Long) As String
    CrackTensionRebarIds = mResults(index).CrackTensionRebarIds
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

Public Property Get CrackReferenceN(ByVal index As Long) As Double
    CrackReferenceN = LoadStateForCombination(index).N
End Property

Public Property Get CrackReferenceMx(ByVal index As Long) As Double
    Dim referenceX As Double
    Dim referenceY As Double
    CrackReferencePointForCombination index, referenceX, referenceY
    CrackReferenceMx = LoadStateForCombination(index).MxAboutPoint(referenceY)
End Property

Public Property Get CrackReferenceMy(ByVal index As Long) As Double
    Dim referenceX As Double
    Dim referenceY As Double
    CrackReferencePointForCombination index, referenceX, referenceY
    CrackReferenceMy = LoadStateForCombination(index).MyAboutPoint(referenceX)
End Property

Public Property Get CrackLoadPointN(ByVal index As Long) As Double
    CrackLoadPointN = LoadStateForCombination(index).N
End Property

Public Property Get CrackLoadPointMx(ByVal index As Long) As Double
    CrackLoadPointMx = LoadStateForCombination(index).LoadPointMx
End Property

Public Property Get CrackLoadPointMy(ByVal index As Long) As Double
    CrackLoadPointMy = LoadStateForCombination(index).LoadPointMy
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

' True, если strength-диаграмма учитывает растянутый бетон и предел нужно показывать.
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

' Возвращает спецификацию материала strength-модели для подробного блока прочности.
' Сначала используем сохраненный named state, потому что он фиксирует реально
' примененную спецификацию, а при отсутствии состояния берем профиль как fallback
' для лимитов в строках, где расчет не запускался или завершился ошибкой.
Private Function StrengthMaterialSpecForCombination(ByVal index As Long) As CMaterialModelSpec
    Dim stateResult As CSectionStateResult
    Set stateResult = mResults(index).FindState(sstStrengthState)
    If stateResult Is Nothing Then Set stateResult = mResults(index).FindState(sstCapacityState)
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

Public Property Get WorstCombinationIndex() As Long
    WorstCombinationIndex = mWorstCombinationIndex
End Property

Public Property Get WorstCombinationID() As String
    If mWorstCombinationIndex > 0 Then WorstCombinationID = mInputs(mWorstCombinationIndex).CombinationID
End Property

Public Property Get WorstCombinationName() As String
    If mWorstCombinationIndex > 0 Then WorstCombinationName = mInputs(mWorstCombinationIndex).Comment
End Property

Public Property Get WorstCombinationCriterion() As String
    WorstCombinationCriterion = NormalizedWorstCombinationCriterion()
End Property

Public Property Get SolverCallCount() As Long
    SolverCallCount = mSolverCallCount
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

' Ищет сохраненное named-state по типу состояния. Этим пользуются схема,
' AutoCAD export и подробные writer-ы, чтобы не пересчитывать НДС повторно.
Public Function FindNamedState(ByVal index As Long, ByVal stateType As ESectionStateType) As CSectionStateResult
    Set FindNamedState = mResults(index).FindState(stateType)
End Function

Public Property Get StrengthReferenceN(ByVal index As Long) As Double
    StrengthReferenceN = LoadStateForCombination(index).N
End Property

Public Property Get StrengthReferenceMx(ByVal index As Long) As Double
    Dim referenceX As Double
    Dim referenceY As Double
    StrengthReferencePointForCombination index, referenceX, referenceY
    StrengthReferenceMx = LoadStateForCombination(index).MxAboutPoint(referenceY)
End Property

Public Property Get StrengthReferenceMy(ByVal index As Long) As Double
    Dim referenceX As Double
    Dim referenceY As Double
    StrengthReferencePointForCombination index, referenceX, referenceY
    StrengthReferenceMy = LoadStateForCombination(index).MyAboutPoint(referenceX)
End Property

Public Property Get StrengthLoadPointN(ByVal index As Long) As Double
    StrengthLoadPointN = LoadStateForCombination(index).N
End Property

Public Property Get StrengthLoadPointMx(ByVal index As Long) As Double
    StrengthLoadPointMx = LoadStateForCombination(index).LoadPointMx
End Property

Public Property Get StrengthLoadPointMy(ByVal index As Long) As Double
    StrengthLoadPointMy = LoadStateForCombination(index).LoadPointMy
End Property

Public Property Get StrengthDirectStateRequested(ByVal index As Long) As Boolean
    If mProfiles.HasProfile(mInputs(index).ProfileId) Then _
        StrengthDirectStateRequested = mProfiles.ProfileById(mInputs(index).ProfileId).StrengthDirectStateEnabled
End Property

Public Property Get StrengthCapacityRequested(ByVal index As Long) As Boolean
    If mProfiles.HasProfile(mInputs(index).ProfileId) Then _
        StrengthCapacityRequested = mProfiles.ProfileById(mInputs(index).ProfileId).StrengthCapacityEnabled
End Property

Public Property Get CrackWidthRequested(ByVal index As Long) As Boolean
    If mProfiles.HasProfile(mInputs(index).ProfileId) Then _
        CrackWidthRequested = mProfiles.ProfileById(mInputs(index).ProfileId).CrackWidthEnabled
End Property

Public Property Get StabilityRequested(ByVal index As Long) As Boolean
    If mProfiles.HasProfile(mInputs(index).ProfileId) Then _
        StabilityRequested = mProfiles.ProfileById(mInputs(index).ProfileId).StabilityEnabled
End Property

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
