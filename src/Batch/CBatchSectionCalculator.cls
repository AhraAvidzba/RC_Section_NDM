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
Private mStatusPolicy As CResultStatusPolicy ' Единый словарь пользовательских статусов и правил их агрегирования.
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
    If mReport Is Nothing Then Set mReport = New CExecutionReport
    Set mReport.UnitSystem = mReportUnits
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
    Set mReport.UnitSystem = units
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
' Используется перед чтением новой таблицы сочетаний, чтобы предыдущие статусы не
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
    mResults(index).SetProfileId profile.ProfileId
    ReportValue "Профиль", profile.ProfileId & " / " & profile.DisplayName
    ReportValue "N", mReport.FormatForce(mInputs(index).N)
    ReportValue "Mx расчетный", mReport.FormatMx(mInputs(index).Mx)
    ReportValue "My расчетный", mReport.FormatMy(mInputs(index).My)

    Dim profileError As String
    If Not profile.ValidateForCalculation(profileError) Then
        ReportStep "Сочетание не рассчитано: " & profileError
        MarkInvalid index, profileError
        Exit Sub
    End If

    mResults(index).SetDirectStateMeta NotApplicableMeta(rkDirectState)
    mResults(index).SetStabilityMeta NotApplicableMeta(rkStability)

    ReportSection "Продольный изгиб и устойчивость"
    If profile.StabilityEnabled Then
        ReportStep "Запускается фильтр продольного изгиба/устойчивости по профилю."
        RunStabilityFilter index, profile.StabilityValueSet
    Else
        ReportStep "Устойчивость не запрошена профилем."
        mResults(index).ClearStabilityResult
        mResults(index).SetStabilityMeta NotApplicableMeta(rkStability)
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
        "N=" & mReport.FormatForce(mInputs(index).N) & "; " & _
        "Mx=" & mReport.FormatMx(mInputs(index).Mx) & "; " & _
        "My=" & mReport.FormatMy(mInputs(index).My)
    If Not mReport Is Nothing Then mReport.ReportStabilityDetails mResults(index).StabilityResult
    Exit Sub

Failed:
    mResults(index).SetStabilityMeta InvalidInputMeta(rkStability, Err.Description), Err.Description
    ReportStep "Устойчивость не рассчитана: " & Err.Description
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
    Dim capacityCalculator As CCapacityCalculator
    Dim loadState As CSectionLoadState
    Dim reportText As String

    Set capacityCalculator = New CCapacityCalculator
    ConfigureCapacityCalculator capacityCalculator
    Set loadState = LoadStateForCombination(index)
    Set loadPath = capacityCalculator.BuildLoadPath(mInputs(index).CapacityLoadPath, loadState)
    ReportValue "Путь нагрузки для несущей", IIf(loadPath.HasPath, loadPath.DisplayName, mInputs(index).CapacityLoadPath)

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

' Выполняет один прямой StateSolution для заданного named-state профиля.
' Материальная модель строится по спецификации профиля, а управление retry и
' extension остается внутри CStateSolutionRunner.
Private Function RunProfileState(ByVal index As Long, ByVal stateType As ESectionStateType, _
        ByVal materialRole As ECalculationPurpose, ByVal spec As CMaterialModelSpec, _
        Optional ByVal useDiagramExtension As Boolean = True, _
        Optional ByVal mergeIntoDirectStateStatus As Boolean = True) As CStateSolutionRunner
    mResults(index).SetDiagramExtensionEnabled mMaterialProvider.DiagramExtensionEnabled
    Dim effectiveDiagramExtension As Boolean
    effectiveDiagramExtension = (useDiagramExtension And mResults(index).DiagramExtensionEnabled)

    Dim request As CStateRequest
    Set request = New CStateRequest
    request.Initialize stateType, materialRole, spec, mInputs(index).N, mInputs(index).Mx, mInputs(index).My, _
        effectiveDiagramExtension, effectiveDiagramExtension

    Dim provider As CStateProvider
    Set provider = CreateStateProvider(index)

    Dim stateResult As CSectionStateResult
    Set stateResult = provider.GetOrSolve(request)

    Dim runner As CStateSolutionRunner
    Set runner = provider.LastRunner

    If stateResult Is Nothing Then
        If mergeIntoDirectStateStatus Then MergeDirectStateMeta index, _
            NumericalFailureMeta(rkDirectState, "Прямое НДС не вернуло расчетное состояние.")
        ReportStep "Прямое НДС " & SectionStateTypeToText(stateType) & " не вернуло расчетное состояние."
        Set RunProfileState = runner
        Exit Function
    End If

    Dim service As CSectionSolver
    If Not runner Is Nothing Then Set service = runner.ResultSolver
    If stateType = sstStrengthState Then mResults(index).StoreSectionStateResult stateResult, _
        mMaterialProvider.ConcreteMaterialFromSpec(spec), mMaterialProvider.SteelMaterialFromSpec(spec)

    ReportValue "НДС " & SectionStateTypeToText(stateType) & " сошлось", IIf(stateResult.Converged, "да", "нет")
    If provider.LastStateWasReused Then
        ReportStep "Прямое НДС " & SectionStateTypeToText(stateType) & " переиспользовано из StateRepository."
    ElseIf Not runner Is Nothing Then
        ReportValue "Итераций НДС", CStr(runner.Iterations)
        ReportBlock "Сценарий прямого StateSolution", runner.DiagnosticLog
        If Not service Is Nothing Then ReportBlock "Итерации прямого CSectionSolver", service.DiagnosticLog
    End If

    If Not stateResult.Converged Then
        If mergeIntoDirectStateStatus Then MergeDirectStateMeta index, stateResult.ResultMeta
        ReportStep "Прямое НДС " & SectionStateTypeToText(stateType) & " не сошлось: " & stateResult.StopReason
        Set RunProfileState = runner
        Exit Function
    End If

    Dim stateStatus As String
    stateStatus = mStatusPolicy.ExternalStatus(stateResult.ResultMeta)
    If stateResult.ExtensionUsed Or Not stateResult.WithinPhysicalRange Then
        If mergeIntoDirectStateStatus Then MergeDirectStateMeta index, stateResult.ResultMeta
        ReportStep "Прямое НДС " & SectionStateTypeToText(stateType) & " найдено, но вышло за физический диапазон диаграммы материала."
    Else
        If mergeIntoDirectStateStatus Then MergeDirectStateMeta index, stateResult.ResultMeta
    End If
    ReportValue "Статус прямого НДС", stateStatus
    ReportMetaComment "Комментарий результата НДС", stateResult.ResultMeta
    ReportValue "Использовано продолжение диаграммы", IIf(stateResult.ExtensionUsed, "да", "нет")
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
    If Not mReport Is Nothing Then mReport.ReportCrackFormationSearchStart mCrackFormationPath

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
    ConfigureCrackFormationCalculator formationCalculator
    Set formationCalculator.StateProvider = CreateStateProvider(index)
    Dim formationResult As CCrackFormationResult
    Set formationResult = formationCalculator.CheckFormation(mSection, _
        mMaterialProvider, profile.CrackedStateSpec, profile.CrackInitiationSpec, _
        mInputs(index).N, mInputs(index).Mx, mInputs(index).My, loadState, _
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

    ReportStep "Запускается текущее CrackedState полного сочетания по профилю трещин."
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

    If Not crack.Converged Then
        ReportStep "Проверка образования трещины не завершилась корректно: " & crack.StopReason
        GoTo ReportCrackResult
    End If
    If Not crack.CrackFormed Then
        ReportStep "Ширина раскрытия не считается: нормальная трещина не образуется при текущей нагрузке."
        GoTo ReportCrackResult
    End If

    ReportStep "Запускается расчет ширины раскрытия нормальных трещин."
    crack.Calculate currentCrackedState, mSection, mMaterialProvider, profile.CrackedStateSpec, _
        formationResult, loadState
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
    ReportValue "Трещина образована", IIf(mResults(index).CrackResult.Formation.CrackFormed, "да", "нет")
    ReportValue "a_crc", mReport.FormatLength(mResults(index).CrackResult.Width.CrackWidth)
    ReportValue "psi_s", FormatNumberInvariant(mResults(index).CrackResult.Width.PsiS)
    ReportValue "As", mReport.FormatArea(mResults(index).CrackResult.Width.AsTension)
    ReportValue "Abt", mReport.FormatArea(mResults(index).CrackResult.Width.Abt)
    ReportValue "ds_eq", mReport.FormatLength(mResults(index).CrackResult.Width.DsEquivalent)
    ReportValue "ls", mReport.FormatLength(mResults(index).CrackResult.Width.CrackSpacing)
    ReportBlock "Диагностика расчета трещин", crack.DiagnosticLog
End Sub

' Упаковывает окончательный result-tree один раз после всех доступных стадий.
' Промежуточная диагностика записывается непосредственно из результатов этапа;
' ни числа, ни meta прошлой упаковки не используются как резервные данные.
Private Sub StoreCrackAggregateSnapshot(ByVal index As Long, ByVal crack As CCrackWidthCalculator, _
        ByVal formationResult As CCrackFormationResult, _
        ByVal currentStateMeta As CResultMeta, ByVal widthMeta As CResultMeta, _
        ByVal longitudinalResult As CLongitudinalCrackResult)
    Dim widthResult As CCrackWidthResult
    Set widthResult = New CCrackWidthResult
    If widthMeta Is Nothing Then
        widthResult.InitializeFromCalculator crack, crack.ResultMeta
    Else
        widthResult.InitializeFromCalculator crack, widthMeta
    End If

    Dim crackResult As CCrackResult
    Set crackResult = New CCrackResult
    crackResult.Initialize formationResult, currentStateMeta, widthResult, longitudinalResult
    mResults(index).StoreCrackAggregateResult crackResult
End Sub

' Проверяет, можно ли использовать CurrentCrackedState как исходные данные для
' ширины нормальных трещин и продольных трещин. Если нельзя, возвращает meta
' фактической причины отказа именно state-этапа, а зависимые формулы блокируются отдельно.
Private Function CurrentCrackedStateReady(ByVal stateResult As CSectionStateResult, _
        ByVal runner As CStateSolutionRunner, ByRef currentStateMeta As CResultMeta, _
        ByRef blockReason As String) As Boolean
    Set currentStateMeta = CurrentCrackedStateMeta(stateResult, runner, blockReason)
    If stateResult Is Nothing Then Exit Function
    If Not stateResult.Converged Then Exit Function
    If stateResult.ExtensionUsed Or Not stateResult.WithinPhysicalRange Then Exit Function
    CurrentCrackedStateReady = True
End Function

' Возвращает meta состояния CrackedState даже тогда, когда CStateProvider не смог
' создать готовый CSectionStateResult. Это позволяет сохранить причину отказа
' в result-tree и не подменять ее статусом формульного расчета ширины трещины.
Private Function CurrentCrackedStateMeta(ByVal stateResult As CSectionStateResult, _
        ByVal runner As CStateSolutionRunner, ByRef blockReason As String) As CResultMeta
    If Not stateResult Is Nothing Then
        Set CurrentCrackedStateMeta = stateResult.ResultMeta
        blockReason = CurrentCrackedStateMeta.ResultComment
        If Len(Trim$(blockReason)) = 0 Then
            If Not stateResult.Converged Then blockReason = "CurrentCrackedState не сошелся."
            If stateResult.ExtensionUsed Or Not stateResult.WithinPhysicalRange Then _
                blockReason = "CurrentCrackedState найден вне физического диапазона диаграмм."
        End If
        Exit Function
    End If

    blockReason = "CurrentCrackedState не получен."
    If Not runner Is Nothing Then
        If Not runner.ResultSolver Is Nothing Then
            If Len(Trim$(runner.ResultSolver.StopReason)) > 0 Then blockReason = runner.ResultSolver.StopReason
        End If
    End If

    Set CurrentCrackedStateMeta = MissingCurrentCrackedStateMeta(runner, blockReason)
End Function

' Формирует meta для случая, когда CurrentCrackedState не был создан вообще.
' Статус не выводится из текста blockReason: если typed-state не создан,
' runner может дать только факт отсутствия физического НДС, а подробная строка
' остается комментарием/diagnostic.
Private Function MissingCurrentCrackedStateMeta(ByVal runner As CStateSolutionRunner, _
        ByVal blockReason As String) As CResultMeta
    Dim meta As CResultMeta
    Set meta = New CResultMeta

    If Not runner Is Nothing Then
        If Not runner.ResultSolver Is Nothing Then
            If Not runner.ResultSolver.Converged Then
                meta.SetSolverFailure runner.ResultSolver.FailureCode, rkDirectState, blockReason, _
                    (runner.ResultSolver.Iterations > 0)
                Set MissingCurrentCrackedStateMeta = meta
                Exit Function
            End If
        End If
    End If

    meta.SetResult rsInternalError, rcSolverDidNotReturnState, rkDirectState, blockReason, _
        vbNullString, True, False
    Set MissingCurrentCrackedStateMeta = meta
End Function



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

' Объединяет meta нескольких прямых StateSolution одного LC.
' N/A не ухудшает итог, а FAIL/NumFail/InputErr сохраняют общий смысл строки.
Private Sub MergeDirectStateMeta(ByVal index As Long, ByVal meta As CResultMeta)
    mResults(index).MergeDirectStateMeta meta
End Sub

' Заполняет capacity-поля нейтральным статусом, когда эта проверка не должна
' выполняться для выбранной группы/режима.
Private Sub StoreNotCalculatedCapacity(ByVal index As Long)
    Dim capacityResult As CCapacityResult
    Set capacityResult = New CCapacityResult
    capacityResult.InitializeSkipped "Проверка несущей способности не запрошена профилем."
    mResults(index).StoreCapacityResult capacityResult
End Sub

' Сохраняет ошибку входных данных в capacity-блок через общий CCapacityResult.
' Даже когда общий LC отбрасывается до solver-а, capacity-часть Results должна
' проходить тем же result/meta-контрактом, что и штатный поиск несущей.
Private Sub StoreInvalidCapacity(ByVal index As Long, ByVal reason As String)
    Dim capacityResult As CCapacityResult
    Set capacityResult = New CCapacityResult
    capacityResult.InitializeInvalidInput reason, CAPACITY_PATH_NONE
    mResults(index).StoreCapacityResult capacityResult
End Sub

' Создает meta для отключенной или неприменимой ветки расчета.
Private Function NotApplicableMeta(ByVal resultKind As EResultKind, _
        Optional ByVal reason As String = vbNullString) As CResultMeta
    Dim meta As CResultMeta
    Set meta = New CResultMeta
    meta.SetNotApplicable resultKind, reason
    Set NotApplicableMeta = meta
End Function

' Создает meta для ошибки пользовательских данных или настроек.
Private Function InvalidInputMeta(ByVal resultKind As EResultKind, _
        ByVal reason As String) As CResultMeta
    Dim meta As CResultMeta
    Set meta = New CResultMeta
    meta.SetResult rsInvalidInput, rcInvalidInput, resultKind, reason
    Set InvalidInputMeta = meta
End Function

' Создает meta для настоящей численной несходимости state-solve/search.
Private Function NumericalFailureMeta(ByVal resultKind As EResultKind, _
        ByVal reason As String) As CResultMeta
    Dim meta As CResultMeta
    Set meta = New CResultMeta
    meta.SetResult rsNumericalFailure, rcNumericalFailure, resultKind, reason
    Set NumericalFailureMeta = meta
End Function

' Фиксирует InvalidInput по конкретному LC и не дает этому сочетанию пройти
' в solver. Так пользователь видит причину в таблице, а остальные LC считаются.
Private Sub MarkInvalid(ByVal index As Long, ByVal reason As String)
    If Len(Trim$(reason)) = 0 Then reason = "Некорректное сочетание нагрузок."
    mInputs(index).InputError = reason
    mResults(index).SetProfileId mInputs(index).ProfileId
    mResults(index).SetDirectStateMeta InvalidInputMeta(rkDirectState, reason)
    mResults(index).SetStabilityMeta InvalidInputMeta(rkStability, reason), reason
    StoreInvalidCapacity index, reason
    If Not mMaterialProvider Is Nothing Then
        mResults(index).SetDiagramExtensionEnabled mMaterialProvider.DiagramExtensionEnabled
    End If
    AppendDiagnostic index
End Sub


' Передает в CCapacityCalculator настройки поиска несущей и внутреннего
' равновесия. Материальные пределы calculator берет из построенных диаграмм
' для конкретного профиля, поэтому batch больше не настраивает CCapacitySolver.
Private Sub ConfigureCapacityCalculator(ByVal calculator As CCapacityCalculator)
    calculator.ConfigureSearch mCapacityInitialLambda, mCapacitySolutionStrategy, _
        mCapacitySearchMethod, mCapacityMaxLambda, mCapacityToleranceLambda, _
        mCapacityToleranceStrain, mCapacityMaxRetries, mCapacityBaseLoadSteps, _
        mCapacitySolverMaxIterations
    calculator.ConfigureEquilibriumSolver mSolverToleranceN, mSolverToleranceMx, _
        mSolverToleranceMy, mSolverMethod, mSolverLineSearchEnabled, _
        mSolverDampingInitial, mSolverMinLineSearchAlpha, mSolverMaxDeltaEpsilon0, _
        mSolverMaxDeltaKappa, mSolverSecantMaxRestarts, mSolverSecantMinStepNorm, _
        DetailedDiagnosticsEnabled()
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

' Создает provider именованных НДС для одного LC и подключает к нему repository
' этого сочетания. Сам provider не знает о профилях и статусной политике batch:
' он только переиспользует либо решает физическое состояние.
Private Function CreateStateProvider(ByVal index As Long) As CStateProvider
    Dim provider As CStateProvider
    Set provider = New CStateProvider
    provider.Initialize mSection, mMaterialProvider, mResults(index).StateRepository
    ConfigureStateProvider provider
    Set CreateStateProvider = provider
End Function

' Передает в StateProvider ровно те же настройки прямого НДС, которые раньше
' получал CStateSolutionRunner напрямую из batch-калькулятора.
Private Sub ConfigureStateProvider(ByVal provider As CStateProvider)
    provider.LoadSteps = mSolverLoadSteps
    provider.SolverMethod = mSolverMethod
    provider.MaxIterations = mSolverMaxIterations
    provider.ToleranceN = mSolverToleranceN
    provider.ToleranceMx = mSolverToleranceMx
    provider.ToleranceMy = mSolverToleranceMy
    provider.LineSearchEnabled = mSolverLineSearchEnabled
    provider.DampingInitial = mSolverDampingInitial
    provider.MinLineSearchAlpha = mSolverMinLineSearchAlpha
    provider.MaxDeltaEpsilon0 = mSolverMaxDeltaEpsilon0
    provider.MaxDeltaKappa = mSolverMaxDeltaKappa
    provider.SecantMaxRestarts = mSolverSecantMaxRestarts
    provider.SecantMinStepNorm = mSolverSecantMinStepNorm
    provider.DiagnosticsEnabled = DetailedDiagnosticsEnabled()
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
    crack.DiagnosticsEnabled = DetailedDiagnosticsEnabled()
End Sub

' Передает formation-настройки его владельцу. Width не хранит solver-options
' и не участвует в поиске Pre/PostCrackState или выборе физического пути.
Private Sub ConfigureCrackFormationCalculator(ByVal formation As CCrackFormationCalculator)
    formation.CrackFormationPath = mCrackFormationPath
    formation.CrackFormationSolutionStrategy = mCrackFormationSolutionStrategy
    formation.SolverLoadSteps = mSolverLoadSteps
    formation.SolverMethod = mSolverMethod
    formation.SolverMaxIterations = mSolverMaxIterations
    formation.SolverToleranceN = mSolverToleranceN
    formation.SolverToleranceMx = mSolverToleranceMx
    formation.SolverToleranceMy = mSolverToleranceMy
    formation.SolverLineSearchEnabled = mSolverLineSearchEnabled
    formation.SolverDampingInitial = mSolverDampingInitial
    formation.SolverMinLineSearchAlpha = mSolverMinLineSearchAlpha
    formation.SolverMaxDeltaEpsilon0 = mSolverMaxDeltaEpsilon0
    formation.SolverMaxDeltaKappa = mSolverMaxDeltaKappa
    formation.SolverSecantMaxRestarts = mSolverSecantMaxRestarts
    formation.SolverSecantMinStepNorm = mSolverSecantMinStepNorm
    formation.DiagnosticsEnabled = DetailedDiagnosticsEnabled()
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
    ' Прежний результат мог быть выдан writer-у или другому потребителю.
    ' Новый запуск не должен менять объект, на который у него осталась ссылка.
    Set mResults(index) = New CCombinationResult
    mResults(index).Clear CAPACITY_PATH_NONE
    mResults(index).SetProfileId mInputs(index).ProfileId
End Sub

' Добавляет одну строку в текстовую диагностику batch-расчета.
' Это компактная диагностическая сводка, а не подробный журнал итераций solver-а.
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

' Добавляет в подробный txt-отчет краткий итог по текущему LC.
' Batch summary остается основным табличным выводом, а здесь пользователь
' видит ход выполнения без открытия Results и без чтения внутренних массивов.
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

' Добавляет в txt-отчет комментарий из ResultMeta конкретного result-subtree.
' Writer/report не склеивает причины сам, а только
' показывает уже сформированный ResultComment ответственного расчетного слоя.
Private Sub ReportMetaComment(ByVal label As String, ByVal meta As CResultMeta)
    If meta Is Nothing Then Exit Sub
    If Len(Trim$(meta.ResultComment)) = 0 Then Exit Sub
    ReportValue label, meta.ResultComment
End Sub

' Добавляет многострочный диагностический блок в txt-отчет и перед записью
' переводит низкоуровневую диагностику solver-а в более читаемый вид.
Private Sub ReportBlock(ByVal title As String, ByVal body As String)
    If mReport Is Nothing Then Exit Sub
    If Not mReport.Enabled Then Exit Sub
    mReport.AddDiagnosticBlock title, body
End Sub

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
' Для CapacityLoadPath это единая величина для lambda*Mx,
' lambda*My, lambda*Mxy, lambda*N и lambda*NMxy; отдельные отношения Nult/N или Mult/Mxy больше
' не нужны для выбора Worst.
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
    candidateSafety = mResults(candidateIndex).ReserveForCriterion(NormalizedWorstCombinationCriterion())
    If candidateSafety <= 0# Then Exit Function

    If currentIndex <= 0 Then
        IsBetterWorstCandidate = True
    Else
        currentSafety = mResults(currentIndex).ReserveForCriterion(NormalizedWorstCombinationCriterion())
        IsBetterWorstCandidate = (currentSafety <= 0# Or candidateSafety < currentSafety)
    End If
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

' Возвращает профиль исходного сочетания до выполнения расчета. Preflight
' компоновки Results использует input, а не еще пустой расчетный result-tree.
Public Property Get InputProfileId(ByVal index As Long) As String
    InputProfileId = mInputs(index).ProfileId
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

' Возвращает каноническое дерево результата LC. Потребители читают его
' непосредственно, без скалярных forwarding-копий в batch-калькуляторе.
Public Property Get ResultAt(ByVal index As Long) As CCombinationResult
    If index < 1 Or index > mCount Then Err.Raise vbObjectError + 3943, _
        "CBatchSectionCalculator", "Номер результата сочетания вне диапазона пакета."
    mResults(index).Freeze
    Set ResultAt = mResults(index)
End Property

Public Property Get StabilityLoadPointMx(ByVal index As Long) As Double
    StabilityLoadPointMx = NormalizeMomentMx(mInputs(index).UserMx)
End Property

' Возвращает My после устойчивости в пользовательской точке задания нагрузки.
Public Property Get StabilityLoadPointMy(ByVal index As Long) As Double
    StabilityLoadPointMy = NormalizeMomentMy(mInputs(index).UserMy)
End Property

' Возвращает пользовательский лимит N/Ncr для eta-ветви СП 35.
' Writer-ы используют это значение в заголовках, чтобы формула запаса
' в Results совпадала с фактически выбранной настройкой Config.
Public Property Get StabilitySP35NOverNcrLimit() As Double
    StabilitySP35NOverNcrLimit = mStabilitySP35NOverNcrLimit
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
    If Len(detail) = 0 And mStatusPolicy.IsInvalid(mResults(index).StabilityResult.Status) Then _
        detail = Trim$(mResults(index).StabilityMeta.ResultComment)
    If Len(detail) = 0 And mStatusPolicy.IsInvalid(mResults(index).StrengthResult.Capacity.Status) Then _
        detail = Trim$(mResults(index).StrengthResult.Capacity.LimitState)
    If Len(detail) = 0 And mStatusPolicy.IsInvalid(mResults(index).StrengthResult.DirectState.Status) Then _
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
        If mResults(i).StateRepository.StateCount > 0 Then StateAvailableCount = StateAvailableCount + 1
    Next i
End Property

' ==========================================================================
' ДЛЯ ТЕСТОВ
' ==========================================================================
' Эти свойства раскрывают промежуточные численные данные batch-расчета,
' которые удобны regression-тестам: отдельно проверяются lambda capacity
' и факт сходимости прямого НДС.

' Повторяет production-маршрут до публикации всего LC и возвращает первый
' неизменяемый crack-result для сравнения. Repository остается рабочим:
' новые formation/width-калькуляторы получают найденные состояния без solve.
Public Function TestRepeatCrackCalculation(ByVal index As Long) As CCrackResult
    If index < 1 Or index > mCount Then Err.Raise vbObjectError + 4270, _
        "CBatchSectionCalculator", "Индекс тестового сочетания вне диапазона."
    Dim profile As CCalculationProfile
    Set profile = mProfiles.ProfileById(mInputs(index).ProfileId)
    Set TestRepeatCrackCalculation = mResults(index).CrackResult
    TestRepeatCrackCalculation.Freeze
    RunCrackWidth index, profile
End Function
