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
Private Const STATUS_OK As String = "OK" ' Расчетная ветка завершилась без ошибки.
Private Const STATUS_FAIL As String = "FAIL" ' Равновесие найдено, но физический предел или проверка не проходят.
Private Const STATUS_NUMFAIL As String = "NumFail" ' Solver не нашел равновесие в заданных численных ограничениях.
Private Const STATUS_INPUTERR As String = "InputErr" ' Ошибка исходных данных, а не численная несходимость.
Private Const STATUS_NA As String = "N/A" ' Расчетная проверка не относится к этому LC/режиму.
Private Const CALC_GROUP1 As String = "Group1" ' Первая группа: проверка несущей способности.
Private Const CALC_GROUP2 As String = "Group2" ' Вторая группа: эксплуатационная проверка трещин.
Private Const CAPACITY_PATH_NONE As String = "None" ' Capacity не запускался для LC.
Private Const CAPACITY_PATH_MX As String = "LambdaMx" ' Масштабируется только пользовательский Mx.
Private Const CAPACITY_PATH_MY As String = "LambdaMy" ' Масштабируется только пользовательский My.
Private Const CAPACITY_PATH_MXY As String = "LambdaMxy" ' Масштабируется пользовательский вектор Mx/My.
Private Const CAPACITY_PATH_N As String = "LambdaN" ' Масштабируется N вместе с моментом от эксцентриситета N.
Private Const CAPACITY_PATH_NMXY As String = "LambdaNMxy" ' Масштабируется весь расчетный вектор N/Mx/My.

Private mSection As CSectionModel ' Расчетная модель сечения, общая для всех LC.
Private mMaterialProvider As CMaterialModelProvider ' Единая точка выдачи диаграмм для Strength/Mcrc/CrackedNDS.
Private mCount As Long ' Сколько сочетаний реально добавлено в пакет.

Private mCombinationID(1 To MAX_COMBINATIONS) As String ' ID сочетания из таблицы нагрузок.
Private mN(1 To MAX_COMBINATIONS) As Double ' Продольная сила после переноса к координатам модели, Н.
Private mMx(1 To MAX_COMBINATIONS) As Double ' Mx после учета точки приложения нагрузки, Н*мм.
Private mMy(1 To MAX_COMBINATIONS) As Double ' My после учета точки приложения нагрузки, Н*мм.
Private mUserMx(1 To MAX_COMBINATIONS) As Double ' Mx как ввел пользователь относительно выбранной точки нагрузки.
Private mUserMy(1 To MAX_COMBINATIONS) As Double ' My как ввел пользователь относительно выбранной точки нагрузки.
Private mCalculationType(1 To MAX_COMBINATIONS) As String ' Group1/Group2: какая инженерная проверка нужна для LC.
Private mCapacityLoadPath(1 To MAX_COMBINATIONS) As String ' Выбор пользователя: какие компоненты масштабировать множителем lambda.
Private mComment(1 To MAX_COMBINATIONS) As String ' Комментарий пользователя из таблицы сочетаний.
Private mInputValid(1 To MAX_COMBINATIONS) As Boolean ' False, если строка LC прочитана, но входные данные некорректны.
Private mInputError(1 To MAX_COMBINATIONS) As String ' Текст причины InvalidInput для конкретного LC.
Private mLoadReferenceX As Double ' X точки приложения нагрузки относительно расчетной модели, мм.
Private mLoadReferenceY As Double ' Y точки приложения нагрузки относительно расчетной модели, мм.
Private mLoadReferenceOffsetX As Double ' Смещение точки нагрузки от центра тяжести бетонного сечения по X, мм.
Private mLoadReferenceOffsetY As Double ' Смещение точки нагрузки от центра тяжести бетонного сечения по Y, мм.

Private mStatus(1 To MAX_COMBINATIONS) As String ' OverallStatus: итоговый пользовательский статус LC.
Private mDirectStateStatus(1 To MAX_COMBINATIONS) As String ' Статус прямого НДС фактического сочетания.
Private mCapacityStatus(1 To MAX_COMBINATIONS) As String ' Статус поиска несущей способности.
Private mCapacityLimitState(1 To MAX_COMBINATIONS) As String ' Какой предел остановил capacity solver.
Private mCapacitySolutionMethod(1 To MAX_COMBINATIONS) As String ' Фактическая численная ветка, которая дала capacity-результат.
Private mCrackStatus(1 To MAX_COMBINATIONS) As String ' Статус расчета трещин для второй группы.
Private mExtensionUsed(1 To MAX_COMBINATIONS) As Boolean ' True, если прямой StateSolution вышел в numerical extension.
Private mDirectStateDiagramExtensionEnabled(1 To MAX_COMBINATIONS) As Boolean ' Снимок настройки extension для LC.
Private mLambdaCapacity(1 To MAX_COMBINATIONS) As Double ' Найденный множитель несущей способности.
Private mUtilCapacity(1 To MAX_COMBINATIONS) As Double ' Использование по несущей способности, обычно 1/lambda.
Private mCapacityPathResolved(1 To MAX_COMBINATIONS) As String ' Нормализованная capacity-траектория после fallback/валидации.
Private mNUltimate(1 To MAX_COMBINATIONS) As Double ' Предельная N для траекторий, где N входит в масштабируемый вектор.
Private mMxUltimate(1 To MAX_COMBINATIONS) As Double ' Предельный пользовательский Mx для траекторий, где масштабируется Mx.
Private mMyUltimate(1 To MAX_COMBINATIONS) As Double ' Предельный пользовательский My для траекторий, где масштабируется My.
Private mMomentUltimate(1 To MAX_COMBINATIONS) As Double ' Результирующий предельный момент из заполненных MxUltimate/MyUltimate.
Private mCapacityStateAvailable(1 To MAX_COMBINATIONS) As Boolean ' True, если сохранена плоскость деформаций именно предельного capacity-состояния.
Private mCapacityEpsilon0(1 To MAX_COMBINATIONS) As Double ' epsilon0 в состоянии найденного предельного момента, а не в фактическом LC.
Private mCapacityKappaX(1 To MAX_COMBINATIONS) As Double ' kappaX в состоянии найденного предельного момента.
Private mCapacityKappaY(1 To MAX_COMBINATIONS) As Double ' kappaY в состоянии найденного предельного момента.
Private mMinConcreteStrain(1 To MAX_COMBINATIONS) As Double ' Наиболее сжатая деформация бетона в фактическом НДС LC.
Private mMaxConcreteStrain(1 To MAX_COMBINATIONS) As Double ' Наиболее растянутая деформация бетона в фактическом НДС LC.
Private mMinSteelStrain(1 To MAX_COMBINATIONS) As Double ' Наиболее сжатая деформация арматуры в фактическом НДС LC.
Private mMaxSteelStrain(1 To MAX_COMBINATIONS) As Double ' Наиболее растянутая деформация арматуры в фактическом НДС LC.
Private mCrackWidth(1 To MAX_COMBINATIONS) As Double ' Расчетная ширина раскрытия трещины, мм.
Private mCrackAllowableValue(1 To MAX_COMBINATIONS) As Double ' Нормативный/пользовательский предел a_crc,ult, мм.
Private mCrackUtil(1 To MAX_COMBINATIONS) As Double ' Использование по ширине трещины.
Private mCrackFormed(1 To MAX_COMBINATIONS) As Boolean ' True, если при текущем LC трещина считается образованной.
Private mCrackPhi1(1 To MAX_COMBINATIONS) As Double ' Коэффициент длительности нагрузки в формуле a_crc.
Private mCrackPhi2(1 To MAX_COMBINATIONS) As Double ' Коэффициент профиля арматуры в формуле a_crc.
Private mCrackPhi3(1 To MAX_COMBINATIONS) As Double ' Коэффициент характера нагружения в формуле a_crc.
Private mCrackPsiS(1 To MAX_COMBINATIONS) As Double ' Коэффициент psi_s: 1 или Auto-уточнение после непрохождения первой проверки.
Private mCrackSigmaS(1 To MAX_COMBINATIONS) As Double ' Среднее напряжение выбранной растянутой арматуры, МПа.
Private mCrackSigmaSCrc(1 To MAX_COMBINATIONS) As Double ' Напряжение той же арматуры в состоянии lambda_crc, МПа.
Private mCrackAs(1 To MAX_COMBINATIONS) As Double ' Площадь выбранной растянутой арматуры As, мм2.
Private mCrackAbt(1 To MAX_COMBINATIONS) As Double ' Площадь растянутого бетона Abt, мм2.
Private mCrackDsEquivalent(1 To MAX_COMBINATIONS) As Double ' Эквивалентный диаметр стержней ds, мм.
Private mCrackLsRaw(1 To MAX_COMBINATIONS) As Double ' Неограниченное расстояние между трещинами ls, мм.
Private mCrackLs(1 To MAX_COMBINATIONS) As Double ' Расчетное ls после ограничений СП, мм.
Private mCrackLambdaCrc(1 To MAX_COMBINATIONS) As Double ' Множитель нагрузки в момент образования трещины.
Private mCrackNcrc(1 To MAX_COMBINATIONS) As Double ' Усилие образования трещины для центрального растяжения, Н.
Private mCrackTensionRebarCount(1 To MAX_COMBINATIONS) As Long ' Количество стержней, вошедших в As.
Private mCrackSectionDepthH(1 To MAX_COMBINATIONS) As Double ' Полная высота h по нормали к нейтральной линии для расчета трещин, мм.
Private mCrackCoverA(1 To MAX_COMBINATIONS) As Double ' Расстояние a от растянутого края до оси крайнего стержня, мм.
Private mCrackTensionDepth(1 To MAX_COMBINATIONS) As Double ' Высота фактической растянутой зоны бетона, мм.
Private mCrackEffectiveZoneDepth(1 To MAX_COMBINATIONS) As Double ' Фактическая глубина зоны Abt, мм.
Private mCrackCentralBranch(1 To MAX_COMBINATIONS) As Boolean ' True, если использована ветка центрального растяжения.
Private mGoverningUtil(1 To MAX_COMBINATIONS) As Double ' Максимальное использование из доступных проверок LC.
Private mGoverningCombinationIndex As Long ' Индекс LC, который стал определяющим в batch summary.
Private mCrackGoverningCombinationIndex As Long ' Индекс LC с минимальным запасом по раскрытию трещин.
Private mElapsedSeconds As Double ' Полное время пакетного расчета, сек.
Private mDiagnosticLog As String ' Короткий текстовый журнал по LC для отладки batch-потока.
Private mStateAvailable(1 To MAX_COMBINATIONS) As Boolean ' True, если для LC сохранена плоскость деформаций.
Private mStateConverged(1 To MAX_COMBINATIONS) As Boolean ' True, если прямой CSectionSolver сошелся.
Private mEpsilon0(1 To MAX_COMBINATIONS) As Double ' Свободный член плоскости деформаций LC.
Private mKappaX(1 To MAX_COMBINATIONS) As Double ' Кривизна при y в формуле epsilon = e0 + kx*y + ky*x.
Private mKappaY(1 To MAX_COMBINATIONS) As Double ' Кривизна при x в формуле epsilon = e0 + kx*y + ky*x.
Private mNint(1 To MAX_COMBINATIONS) As Double ' Внутренняя продольная сила из напряжений элементов, Н.
Private mMxint(1 To MAX_COMBINATIONS) As Double ' Внутренний Mx из напряжений элементов, Н*мм.
Private mMyint(1 To MAX_COMBINATIONS) As Double ' Внутренний My из напряжений элементов, Н*мм.
Private mResidualN(1 To MAX_COMBINATIONS) As Double ' Невязка Nint - N после прямого solve, Н.
Private mResidualMx(1 To MAX_COMBINATIONS) As Double ' Невязка Mxint - Mx после прямого solve, Н*мм.
Private mResidualMy(1 To MAX_COMBINATIONS) As Double ' Невязка Myint - My после прямого solve, Н*мм.
Private mStateIterations(1 To MAX_COMBINATIONS) As Long ' Число итераций прямого solver-а для LC.

Private mCrackAllowable As Double ' Допустимая ширина раскрытия трещины a_crc,ult, мм.
Private mCrackPsiMode As String ' Fixed1 или Auto: как определять psi_s для нормальных трещин.
Private mCrackTensionZoneMode As String ' Effective или FullTension: какая бетонная зона входит в Abt.
Private mCrackEnabled As Boolean ' Общий выключатель расчета трещин для сочетаний второй группы.
Private mCalculationMode As String ' DirectState или FullCapacity.
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
Private mCapacityCalculationScope As String ' Какие LC отправлять в поиск несущего момента: только Group1 или Group1+Group2.
Private mCapacitySearchMethod As String ' Bisection, Brent или Secant для LoadMultiplier.
Private mCapacityMaxLambda As Double ' Верхняя защитная граница lambda.
Private mCapacityToleranceLambda As Double ' Точность lambda в одномерном поиске LoadMultiplier.
Private mCapacityToleranceStrain As Double ' Точность попадания в предельную деформацию UltimateStrain.
Private mCapacityMaxRetries As Long ' Повторные попытки LoadMultiplier при несходимости пробной точки.
Private mCapacityBaseLoadSteps As Long ' Базовое число load steps для пробной точки capacity.
Private mCapacitySolverMaxIterations As Long ' Лимит итераций внутреннего solver-а при capacity-поиске.
Private mReport As CExecutionReport ' Необязательный человекочитаемый отчет выполнения; расчетная логика от него не зависит.

' Инициализирует внутреннее состояние объекта значениями по умолчанию, безопасными для последующей настройки.
Private Sub Class_Initialize()
    mCrackAllowable = 0.3
    mCrackPsiMode = "Fixed1"
    mCrackTensionZoneMode = "Effective"
    mCalculationMode = "FullCapacity"
    mSolverMethod = "Newton"
    mCrackEnabled = True
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
    mCapacityCalculationScope = "Group1+2"
    mCapacitySearchMethod = "Bisection"
    mCapacityMaxLambda = 64#
    mCapacityToleranceLambda = 0.01
    mCapacityToleranceStrain = 0.00001
    mCapacityMaxRetries = 0
    mCapacityBaseLoadSteps = 1
    mCapacitySolverMaxIterations = 60
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
End Sub

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
    mCrackAllowable = settings.GetDouble("SLS.Crack.Allowable", mCrackAllowable)
    mCrackPsiMode = settings.GetString("SLS.Crack.PsiMode", mCrackPsiMode)
    mCrackTensionZoneMode = settings.GetString("SLS.Crack.TensionZoneMode", mCrackTensionZoneMode)
    mCalculationMode = settings.GetString("Calculation.Mode", mCalculationMode)
    If settings.HasKey("Solver.Method") Then
        mSolverMethod = settings.GetRawString("Solver.Method", vbNullString)
    Else
        mSolverMethod = "Newton"
    End If
    mCrackEnabled = settings.GetBoolean("SLS.Crack.Enabled", mCrackEnabled)
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
    mCapacityCalculationScope = settings.GetString("Capacity.CalculationScope", mCapacityCalculationScope)
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
    If Not units Is Nothing Then
        mCrackAllowable = units.InputLengthToInternal(mCrackAllowable)
        mSolverToleranceN = Abs(units.InputForceToInternal(mSolverToleranceN))
        mSolverToleranceMx = Abs(units.InputMomentMxToInternal(mSolverToleranceMx))
        mSolverToleranceMy = Abs(units.InputMomentMyToInternal(mSolverToleranceMy))
        mSolverMaxDeltaKappa = Abs(units.InputCurvatureToInternal(mSolverMaxDeltaKappa))
    End If
End Sub

' Очищает список LC и накопленные результаты.
' Используется перед чтением новой таблицы сочетаний, чтобы старые статусы не
' попали в новый batch summary.
Public Sub ClearCombinations()
    mCount = 0
    mGoverningCombinationIndex = 0
    mCrackGoverningCombinationIndex = 0
    mElapsedSeconds = 0#
    mDiagnosticLog = vbNullString
    mLoadReferenceX = 0#
    mLoadReferenceY = 0#
    mLoadReferenceOffsetX = 0#
    mLoadReferenceOffsetY = 0#
End Sub

' Добавляет корректно прочитанное сочетание нагрузок.
' Усилия уже должны быть во внутренних единицах и в INTERNAL-системе знаков.
Public Sub AddCombination(ByVal combinationID As String, ByVal nValue As Double, ByVal mxValue As Double, _
        ByVal myValue As Double, ByVal calculationType As String, ByVal comment As String, _
        Optional ByVal capacityLoadPath As String = vbNullString)
    AddCombinationInternal combinationID, nValue, mxValue, myValue, calculationType, comment, True, vbNullString, capacityLoadPath
End Sub

' Сохраняет строку LC с ошибкой ввода, чтобы пользователь увидел ее в отчете.
' Такие сочетания не передаются в solver, но отображаются как InvalidInput.
Public Sub AddInvalidCombination(ByVal combinationID As String, ByVal calculationType As String, _
        ByVal comment As String, ByVal reason As String, Optional ByVal capacityLoadPath As String = vbNullString)
    AddCombinationInternal combinationID, 0#, 0#, 0#, calculationType, comment, False, reason, capacityLoadPath
End Sub

Private Sub AddCombinationInternal(ByVal combinationID As String, ByVal nValue As Double, ByVal mxValue As Double, _
        ByVal myValue As Double, ByVal calculationType As String, ByVal comment As String, _
        ByVal inputValid As Boolean, ByVal inputError As String, ByVal capacityLoadPath As String)
    If mCount >= MAX_COMBINATIONS Then Err.Raise vbObjectError + 3911, "CBatchSectionCalculator", "Поддерживается не более 20 сочетаний нагрузок."
    mCount = mCount + 1
    If Len(Trim$(combinationID)) = 0 Then combinationID = "ROW" & CStr(mCount)
    mCombinationID(mCount) = combinationID
    mN(mCount) = nValue
    mMx(mCount) = mxValue
    mMy(mCount) = myValue
    mCalculationType(mCount) = calculationType
    mCapacityLoadPath(mCount) = Trim$(capacityLoadPath)
    mComment(mCount) = comment
    mUserMx(mCount) = mxValue
    mUserMy(mCount) = myValue
    mInputValid(mCount) = inputValid
    mInputError(mCount) = inputError
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
        If mInputValid(i) Then
            mMx(i) = mUserMx(i) + mN(i) * referenceY
            mMy(i) = mUserMy(i) + mN(i) * referenceX
        End If
    Next i
End Sub

' Последовательно считает все добавленные LC и выбирает определяющее сочетание.
' Порядок один для всех геометрий: capacity при необходимости, прямое НДС,
' затем трещины для сочетаний второй группы.
Public Sub Execute()
    On Error GoTo Failed

    ValidateReady
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
        If mInputValid(i) Then
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
    ReportSection "Сочетание " & CStr(index) & ": " & mCombinationID(index)
    ReportValue "Комментарий", mComment(index)
    If Not mInputValid(index) Then
        ReportStep "Сочетание не передано в solver из-за ошибки входных данных: " & mInputError(index)
        MarkInvalid index, mInputError(index)
        Exit Sub
    End If
    If Not IsSupportedCalculationMode() Then
        ReportStep "Сочетание не рассчитано: неизвестный Calculation.Mode."
        MarkInvalid index, "Calculation.Mode должен быть DirectState или FullCapacity."
        Exit Sub
    End If

    Dim calculationType As String
    calculationType = NormalizedCombinationCalculationType(mCalculationType(index))
    ReportValue "Тип проверки", IIf(Len(calculationType) > 0, calculationType, mCalculationType(index))
    ReportValue "N", FormatNumberInvariant(mN(index)) & " Н"
    ReportValue "Mx расчетный", FormatNumberInvariant(mMx(index)) & " Н*мм"
    ReportValue "My расчетный", FormatNumberInvariant(mMy(index)) & " Н*мм"
    If Len(calculationType) = 0 Then
        ReportStep "Сочетание не рассчитано: CalculationType должен быть Group1 или Group2."
        MarkInvalid index, "CalculationType должен быть Group1 или Group2."
        Exit Sub
    End If
    If IsFullCapacityMode And Len(NormalizedCapacityCalculationScope()) = 0 Then
        ReportStep "Сочетание не рассчитано: неизвестная область расчета Capacity."
        MarkInvalid index, "Capacity.CalculationScope должен быть Group1Only или Group1+2."
        Exit Sub
    End If

    mStatus(index) = STATUS_NA
    If ShouldRunCapacity(calculationType) And IsFullCapacityMode Then
        ReportStep "Запускается проверка несущей способности для сочетания."
        RunCapacity index
    Else
        ReportStep "Проверка несущей способности не требуется для этого сочетания или режима Calculation.Mode."
        StoreNotCalculatedCapacity index
    End If
    RunSectionStateAndCrack index, ShouldRunCrack(calculationType)

    mGoverningUtil(index) = MaxDouble(mUtilCapacity(index), mCrackUtil(index))
    mStatus(index) = AggregateOverallStatus(index)
    AppendDiagnostic index
    ReportCombinationSummary index
    Exit Sub

Failed:
    ReportStep "Ошибка при расчете сочетания: " & Err.Description
    MarkInvalid index, Err.Description
End Sub

' Запускает поиск несущей способности для LC.
' Пользовательский столбец CapacityLoadPath задает, какие компоненты нагрузки
' умножаются на lambda. Все варианты дальше сводятся к одной математике:
' Target = Offset + lambda*Base для N, Mx и My.
Private Sub RunCapacity(ByVal index As Long)
    Dim loadPath As String
    Dim nOffset As Double
    Dim nBase As Double
    Dim mxOffset As Double
    Dim mxBase As Double
    Dim myOffset As Double
    Dim myBase As Double
    Dim cap As CCapacitySolver

    loadPath = CapacityLoadPathForCombination(index)
    mCapacityPathResolved(index) = loadPath
    ReportValue "CapacityLoadPath", IIf(Len(loadPath) > 0, CapacityLoadPathDisplay(loadPath), mCapacityLoadPath(index))

    If Len(loadPath) = 0 Then
        If Len(Trim$(mCapacityLoadPath(index))) > 0 Then
            StoreInvalidCapacity index, "CapacityLoadPath должен быть λ*Mx, λ*My, λ*Mxy, λ*N или λ*NMxy."
            ReportStep "Capacity остановлен: неизвестное значение CapacityLoadPath."
        Else
            ReportStep "Capacity не запускается: в сочетании нет ни момента, ни продольной силы."
            StoreSkippedCapacity index
        End If
        Exit Sub
    End If

    If Not CapacityLoadPathHasScaledLoad(index, loadPath) Then
        StoreInvalidCapacity index, "Для выбранного CapacityLoadPath соответствующая нагрузка должна быть ненулевой."
        ReportStep "Capacity остановлен: выбранная λ-траектория не содержит ненулевой масштабируемой нагрузки."
        Exit Sub
    End If

    BuildCapacityLoadPath index, loadPath, nOffset, nBase, mxOffset, mxBase, myOffset, myBase
    Set cap = New CCapacitySolver
    ConfigureCapacity cap
    Dim forceOnlyPath As Boolean
    ' Если пользователь не задавал Mx/My, то λ*N и λ*NMxy являются силовой
    ' траекторией продольной силы. Во внутренних моментах могут быть члены
    ' переноса от N к расчетной системе координат, но это не делает исходную
    ' пользовательскую задачу изгибной.
    forceOnlyPath = IsForceOnlyCapacityLoadPath(index, loadPath)
    Select Case NormalizedCapacitySolutionStrategy()
        Case "Auto"
            If forceOnlyPath Then
                cap.SolveByLoadPathMultiplier mSection, mMaterialProvider.ConcreteMaterial(cpStrength), _
                    mMaterialProvider.SteelMaterial(cpStrength), nOffset, nBase, mxOffset, mxBase, myOffset, myBase, True
            Else
                cap.SolveByAutoLoadPath mSection, mMaterialProvider.ConcreteMaterial(cpStrength), _
                    mMaterialProvider.SteelMaterial(cpStrength), nOffset, nBase, mxOffset, mxBase, myOffset, myBase
            End If
        Case "UltimateStrain"
            If forceOnlyPath Then
                cap.SolveByLoadPathMultiplier mSection, mMaterialProvider.ConcreteMaterial(cpStrength), _
                    mMaterialProvider.SteelMaterial(cpStrength), nOffset, nBase, mxOffset, mxBase, myOffset, myBase, True
            Else
                cap.SolveByUltimateLoadPath mSection, mMaterialProvider.ConcreteMaterial(cpStrength), _
                    mMaterialProvider.SteelMaterial(cpStrength), nOffset, nBase, mxOffset, mxBase, myOffset, myBase
            End If
        Case "LoadMultiplier"
            cap.SolveByLoadPathMultiplier mSection, mMaterialProvider.ConcreteMaterial(cpStrength), _
                mMaterialProvider.SteelMaterial(cpStrength), nOffset, nBase, mxOffset, mxBase, myOffset, myBase, _
                forceOnlyPath
        Case Else
            mCapacityStatus(index) = STATUS_INPUTERR
            mCapacityLimitState(index) = "Capacity.SolutionStrategy должен быть Auto, LoadMultiplier или UltimateStrain."
            mCapacitySolutionMethod(index) = vbNullString
            ReportStep "Capacity остановлен: неизвестный Capacity.SolutionStrategy."
            Exit Sub
    End Select

    StoreCapacity index, cap, loadPath

    ReportValue "CapacityStatus", mCapacityStatus(index)
    ReportValue "LambdaCapacity", FormatNumberInvariant(mLambdaCapacity(index))
    ReportValue "CapacityLoadPath", CapacityLoadPathDisplay(mCapacityPathResolved(index))
    ReportBlock "Диагностика поиска несущей способности", cap.DiagnosticLog
    If Not cap.LastSolver Is Nothing Then
        ReportBlock "Итерации внутреннего CSectionSolver для последней capacity-точки", cap.LastSolver.DiagnosticLog
    End If
End Sub

' Выполняет прямой расчет НДС и, если нужно, расчет ширины трещин.
' Даже в FullCapacity этот solve нужен отдельно: он показывает состояние
' именно от заданного LC, а не только предельное состояние capacity solver-а.
Private Sub RunSectionStateAndCrack(ByVal index As Long, ByVal crackRequested As Boolean)
    Dim basePurpose As ECalculationPurpose
    basePurpose = MaterialPurposeForCombination(index)

    Dim concreteState As CMaterialDiagram
    Dim steelState As CMaterialDiagram
    Set concreteState = mMaterialProvider.ConcreteStateMaterial(basePurpose)
    Set steelState = mMaterialProvider.SteelStateMaterial(basePurpose)
    mDirectStateDiagramExtensionEnabled(index) = mMaterialProvider.DirectStateDiagramExtensionEnabled

    Dim service As CSectionSolver
    Set service = New CSectionSolver
    ConfigureSectionSolver service
    ReportStep "Запускается прямой расчет НДС CSectionSolver."
    service.Solve mSection, concreteState, steelState, mN(index), mMx(index), mMy(index)
    StoreSectionState index, service
    ReportValue "НДС сошелся", IIf(service.Converged, "да", "нет")
    ReportValue "Итераций НДС", CStr(service.Iterations)
    ReportBlock "Итерации прямого CSectionSolver", service.DiagnosticLog
    If Not service.Converged And mDirectStateDiagramExtensionEnabled(index) Then
        If TrySolveStateWithExtensionWarmStart(index, concreteState, steelState, service) Then
            StoreSectionState index, service
            ReportStep "Прямое НДС сошлось после повторного старта в technical extension."
            ReportValue "Итераций повторного НДС", CStr(service.Iterations)
            ReportBlock "Итерации повторного CSectionSolver", service.DiagnosticLog
        End If
    End If
    If Not service.Converged Then
        mDirectStateStatus(index) = STATUS_NUMFAIL
        mCrackStatus(index) = STATUS_NA
        ReportStep "Расчет трещин не запускается, потому что прямое НДС не сошлось: " & service.StopReason
        Exit Sub
    End If

    Dim stateWithinPhysical As Boolean
    stateWithinPhysical = StateWithinPhysicalRange(service, concreteState, steelState)
    mExtensionUsed(index) = StateUsesExtension(service, concreteState, steelState)

    If mExtensionUsed(index) Or Not stateWithinPhysical Then
        mDirectStateStatus(index) = STATUS_FAIL
        ReportStep "Прямое НДС найдено, но вышло за физический диапазон диаграммы материала."
    Else
        mDirectStateStatus(index) = STATUS_OK
    End If
    ReportValue "DirectStateStatus", mDirectStateStatus(index)
    ReportValue "ExtensionUsed", IIf(mExtensionUsed(index), "да", "нет")

    If Not crackRequested Or Not mCrackEnabled Then
        mCrackStatus(index) = STATUS_NA
        If crackRequested Then
            ReportStep "Расчет трещин отключен настройкой SLS.Crack.Enabled."
        Else
            ReportStep "Расчет трещин не требуется для этого типа сочетания."
        End If
        Exit Sub
    End If

    If mDirectStateStatus(index) <> STATUS_OK Or mExtensionUsed(index) Then
        mCrackStatus(index) = STATUS_NA
        ReportStep "Расчет трещин не запускается: исходное НДС Group2 не находится полностью в физическом диапазоне диаграммы."
        Exit Sub
    End If

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
    centroidMxForCrack = mUserMx(index) + mN(index) * mLoadReferenceOffsetY
    centroidMyForCrack = mUserMy(index) + mN(index) * mLoadReferenceOffsetX

    crack.Calculate service, mSection, mMaterialProvider, basePurpose, _
        mN(index), mMx(index), mMy(index), _
        centroidMxForCrack, centroidMyForCrack
    mCrackStatus(index) = CrackStatusFromCalculator(crack)
    mCrackWidth(index) = crack.CrackWidth
    mCrackAllowableValue(index) = crack.AllowableCrackWidth
    mCrackUtil(index) = crack.Utilization
    mCrackFormed(index) = crack.CrackFormed
    mCrackPhi1(index) = crack.Phi1
    mCrackPhi2(index) = crack.Phi2
    mCrackPhi3(index) = crack.Phi3
    mCrackPsiS(index) = crack.PsiS
    mCrackSigmaS(index) = crack.SigmaS
    mCrackSigmaSCrc(index) = crack.SigmaSCrc
    mCrackAs(index) = crack.AsTension
    mCrackAbt(index) = crack.Abt
    mCrackDsEquivalent(index) = crack.DsEquivalent
    mCrackLsRaw(index) = crack.CrackSpacingRaw
    mCrackLs(index) = crack.CrackSpacing
    mCrackLambdaCrc(index) = crack.LambdaCrc
    mCrackNcrc(index) = crack.Ncrc
    mCrackTensionRebarCount(index) = crack.TensionRebarCount
    mCrackSectionDepthH(index) = crack.SectionDepthH
    mCrackCoverA(index) = crack.CoverA
    mCrackTensionDepth(index) = crack.TensionDepth
    mCrackEffectiveZoneDepth(index) = crack.EffectiveZoneDepth
    mCrackCentralBranch(index) = crack.CentralTensionBranch
    ReportValue "CrackStatus", mCrackStatus(index)
    ReportValue "Трещина образована", IIf(mCrackFormed(index), "да", "нет")
    ReportValue "a_crc", FormatNumberInvariant(mCrackWidth(index)) & " мм"
    ReportValue "psi_s", FormatNumberInvariant(mCrackPsiS(index))
    ReportValue "As", FormatNumberInvariant(mCrackAs(index)) & " мм2"
    ReportValue "Abt", FormatNumberInvariant(mCrackAbt(index)) & " мм2"
    ReportValue "ds_eq", FormatNumberInvariant(mCrackDsEquivalent(index)) & " мм"
    ReportValue "ls", FormatNumberInvariant(mCrackLs(index)) & " мм"
    ReportBlock "Диагностика расчета трещин", crack.DiagnosticLog
End Sub

Private Function NormalizedCombinationCalculationType(ByVal calculationType As String) As String
    Select Case LCase$(Trim$(calculationType))
        Case "group1", "1", "uls", "strength", "strengthonly"
            NormalizedCombinationCalculationType = CALC_GROUP1
        Case "group2", "2", "sls", "crack", "crackonly"
            NormalizedCombinationCalculationType = CALC_GROUP2
    End Select
End Function

' Показывает, должна ли для LC запускаться несущая способность.
' Пользовательская настройка Capacity.CalculationScope позволяет оставить
' поиск предельной λ-траектории только для первой группы или временно считать
' его также для Group2 как диагностический прочностной запас.
Private Function ShouldRunCapacity(ByVal calculationType As String) As Boolean
    Select Case NormalizedCapacityCalculationScope()
        Case "Group1Only"
            ShouldRunCapacity = (calculationType = CALC_GROUP1)
        Case "Group1AndGroup2"
            ShouldRunCapacity = (calculationType = CALC_GROUP1 Or calculationType = CALC_GROUP2)
    End Select
End Function

' Показывает, должна ли для LC запускаться проверка трещин.
' Текущая реализация считает только продолжительное раскрытие нормальных
' трещин для второй группы; сочетания первой группы сюда не попадают.
Private Function ShouldRunCrack(ByVal calculationType As String) As Boolean
    ShouldRunCrack = (calculationType = CALC_GROUP2)
End Function

' True, если в LC есть пользовательский момент Mx или My.
' Моменты от смещенной точки приложения N сюда не входят: пользователь не
' задавал их как отдельную нагрузку, они формируются переносом силы.
Private Function HasUserMomentVector(ByVal index As Long) As Boolean
    HasUserMomentVector = (Sqr(mUserMx(index) * mUserMx(index) + mUserMy(index) * mUserMy(index)) > 0.000000001)
End Function

' True для траекторий, где пользователь масштабирует только продольную силу.
' CapacityLoadPath при этом не переписывается: λ*NMxy остается λ*NMxy в
' Results, но если пользовательские Mx/My равны нулю, фактический вектор
' нагрузки содержит только N и моменты переноса от выбранной точки приложения.
Private Function IsForceOnlyCapacityLoadPath(ByVal index As Long, ByVal loadPath As String) As Boolean
    IsForceOnlyCapacityLoadPath = (loadPath = CAPACITY_PATH_N Or loadPath = CAPACITY_PATH_NMXY) And _
        Not HasUserMomentVector(index)
End Function

' Нормализует текст из Excel в короткий внутренний ключ load-path.
' Явный текст должен совпадать с пользовательской моделью λ*..., иначе
' строка сочетания получает InputErr вместо молчаливого авто-исправления.
Private Function NormalizedCapacityLoadPath(ByVal loadPath As String) As String
    Dim value As String
    value = LCase$(Trim$(loadPath))
    value = Replace$(value, " ", vbNullString)
    value = Replace$(value, ChrW$(&H3BB), "lambda")
    value = Replace$(value, "λ", "lambda")

    Select Case value
        Case "lambda*mx", "lambdamx", "mx"
            NormalizedCapacityLoadPath = CAPACITY_PATH_MX
        Case "lambda*my", "lambdamy", "my"
            NormalizedCapacityLoadPath = CAPACITY_PATH_MY
        Case "lambda*mxy", "lambdamxy", "mxy"
            NormalizedCapacityLoadPath = CAPACITY_PATH_MXY
        Case "lambda*n", "lambdan", "nload"
            NormalizedCapacityLoadPath = CAPACITY_PATH_N
        Case "lambda*nmxy", "lambdanmxy", "nmxy", "all"
            NormalizedCapacityLoadPath = CAPACITY_PATH_NMXY
    End Select
End Function

' Возвращает load-path, который реально будет использован для LC.
' Пустой столбец трактуется осторожно: моментное сочетание идет по λ*Mxy,
' чистая продольная сила - по λ*N. Явный неизвестный текст остается ошибкой.
Private Function CapacityLoadPathForCombination(ByVal index As Long) As String
    Dim explicitValue As String
    explicitValue = NormalizedCapacityLoadPath(mCapacityLoadPath(index))
    If Len(explicitValue) > 0 Then
        CapacityLoadPathForCombination = explicitValue
    ElseIf Len(Trim$(mCapacityLoadPath(index))) > 0 Then
        CapacityLoadPathForCombination = vbNullString
    ElseIf HasUserMomentVector(index) Then
        CapacityLoadPathForCombination = CAPACITY_PATH_MXY
    ElseIf Abs(mN(index)) > 0.000000001 Then
        CapacityLoadPathForCombination = CAPACITY_PATH_N
    End If
End Function

' Формирует человекочитаемую подпись load-path для Results и отчета.
Private Function CapacityLoadPathDisplay(ByVal loadPath As String) As String
    Select Case loadPath
        Case CAPACITY_PATH_MX
            CapacityLoadPathDisplay = ChrW$(&H3BB) & "*Mx"
        Case CAPACITY_PATH_MY
            CapacityLoadPathDisplay = ChrW$(&H3BB) & "*My"
        Case CAPACITY_PATH_MXY
            CapacityLoadPathDisplay = ChrW$(&H3BB) & "*Mxy"
        Case CAPACITY_PATH_N
            CapacityLoadPathDisplay = ChrW$(&H3BB) & "*N"
        Case CAPACITY_PATH_NMXY
            CapacityLoadPathDisplay = ChrW$(&H3BB) & "*NMxy"
        Case Else
            CapacityLoadPathDisplay = loadPath
    End Select
End Function

' Проверяет, есть ли в выбранном пути хотя бы одна ненулевая масштабируемая
' пользовательская компонента нагрузки.
Private Function CapacityLoadPathHasScaledLoad(ByVal index As Long, ByVal loadPath As String) As Boolean
    Select Case loadPath
        Case CAPACITY_PATH_MX
            CapacityLoadPathHasScaledLoad = (Abs(mUserMx(index)) > 0.000000001)
        Case CAPACITY_PATH_MY
            CapacityLoadPathHasScaledLoad = (Abs(mUserMy(index)) > 0.000000001)
        Case CAPACITY_PATH_MXY
            CapacityLoadPathHasScaledLoad = HasUserMomentVector(index)
        Case CAPACITY_PATH_N
            CapacityLoadPathHasScaledLoad = (Abs(mN(index)) > 0.000000001)
        Case CAPACITY_PATH_NMXY
            CapacityLoadPathHasScaledLoad = (Abs(mN(index)) + Abs(mUserMx(index)) + Abs(mUserMy(index)) > 0.000000001)
    End Select
End Function

' Переводит выбранный пользователем путь в математическую форму solver-а.
' Offset - неизменяемая часть внешней нагрузки, Base - часть, умножаемая
' на lambda. Моменты от эксцентриситета N попадают туда, куда физически
' попадает сама N: при λ*N они масштабируются, при λ*M остаются постоянными.
Private Sub BuildCapacityLoadPath(ByVal index As Long, ByVal loadPath As String, _
        ByRef nOffset As Double, ByRef nBase As Double, _
        ByRef mxOffset As Double, ByRef mxBase As Double, _
        ByRef myOffset As Double, ByRef myBase As Double)
    Dim mxFromN As Double
    Dim myFromN As Double
    mxFromN = mN(index) * mLoadReferenceY
    myFromN = mN(index) * mLoadReferenceX

    Select Case loadPath
        Case CAPACITY_PATH_MX
            nOffset = mN(index)
            mxOffset = mxFromN
            mxBase = mUserMx(index)
            myOffset = myFromN + mUserMy(index)
        Case CAPACITY_PATH_MY
            nOffset = mN(index)
            mxOffset = mxFromN + mUserMx(index)
            myOffset = myFromN
            myBase = mUserMy(index)
        Case CAPACITY_PATH_MXY
            nOffset = mN(index)
            mxOffset = mxFromN
            mxBase = mUserMx(index)
            myOffset = myFromN
            myBase = mUserMy(index)
        Case CAPACITY_PATH_N
            nBase = mN(index)
            mxOffset = mUserMx(index)
            mxBase = mxFromN
            myOffset = mUserMy(index)
            myBase = myFromN
        Case CAPACITY_PATH_NMXY
            nBase = mN(index)
            mxBase = mMx(index)
            myBase = mMy(index)
    End Select
End Sub

Private Function CapacityPathScalesN(ByVal loadPath As String) As Boolean
    CapacityPathScalesN = (loadPath = CAPACITY_PATH_N Or loadPath = CAPACITY_PATH_NMXY)
End Function

Private Function CapacityPathScalesMx(ByVal loadPath As String) As Boolean
    CapacityPathScalesMx = (loadPath = CAPACITY_PATH_MX Or loadPath = CAPACITY_PATH_MXY Or loadPath = CAPACITY_PATH_NMXY)
End Function

Private Function CapacityPathScalesMy(ByVal loadPath As String) As Boolean
    CapacityPathScalesMy = (loadPath = CAPACITY_PATH_MY Or loadPath = CAPACITY_PATH_MXY Or loadPath = CAPACITY_PATH_NMXY)
End Function

' Выбирает material purpose для фактического НДС конкретного LC.
' Group1 решается по Strength-модели I ГПС, Group2 - по CrackedNDS-модели
' II ГПС с неработающим растянутым бетоном, как принято для a_crc.
Private Function MaterialPurposeForCombination(ByVal index As Long) As ECalculationPurpose
    If NormalizedCombinationCalculationType(mCalculationType(index)) = CALC_GROUP2 Then
        MaterialPurposeForCombination = cpCrackedNDS
    Else
        MaterialPurposeForCombination = cpStrength
    End If
End Function

' Возвращает бетонную StateSolution-диаграмму для writer-а Results.
' Если прямой НДС решался с extension, именно эта диаграмма воспроизводит
' сохраненные Stress, которые пользователь увидит на Results/схеме/AutoCAD.
Public Function ConcreteMaterialForCombination(ByVal index As Long) As Object
    Set ConcreteMaterialForCombination = mMaterialProvider.ConcreteStateMaterial(MaterialPurposeForCombination(index))
End Function

' Возвращает StateSolution-диаграмму арматуры для writer-а Results.
Public Function SteelMaterialForCombination(ByVal index As Long) As Object
    Set SteelMaterialForCombination = mMaterialProvider.SteelStateMaterial(MaterialPurposeForCombination(index))
End Function

' True, когда выбран полный расчет с поиском несущей способности.
Private Function IsFullCapacityMode() As Boolean
    IsFullCapacityMode = (StrComp(mCalculationMode, "FullCapacity", vbTextCompare) = 0)
End Function

' True, когда пользователь запросил только фактическое НДС без capacity-поиска.
Private Function IsDirectStateMode() As Boolean
    IsDirectStateMode = (StrComp(mCalculationMode, "DirectState", vbTextCompare) = 0)
End Function

' Защищает RunCombination от неизвестных режимов Calculation.Mode.
Private Function IsSupportedCalculationMode() As Boolean
    IsSupportedCalculationMode = IsDirectStateMode() Or IsFullCapacityMode()
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

' Заполняет capacity-поля нейтральным статусом, когда эта проверка не должна
' выполняться для выбранной группы/режима.
Private Sub StoreNotCalculatedCapacity(ByVal index As Long)
    mCapacityStatus(index) = STATUS_NA
    mCapacityLimitState(index) = STATUS_NA
    mCapacitySolutionMethod(index) = vbNullString
    mCapacityPathResolved(index) = CAPACITY_PATH_NONE
End Sub

' Помечает capacity как пропущенную из-за нулевой нагрузки.
' Это не ошибка исходных данных: для нулевых N, Mx и My нельзя построить
' направление внешнего поиска несущей способности.
Private Sub StoreSkippedCapacity(ByVal index As Long)
    mCapacityStatus(index) = STATUS_NA
    mCapacityLimitState(index) = "ZeroBaseLoad"
    mCapacitySolutionMethod(index) = vbNullString
    mCapacityPathResolved(index) = CAPACITY_PATH_NONE
End Sub

' Помечает только ветку capacity как ошибку исходных данных.
' Прямое НДС при этом все равно может быть посчитано: ошибка относится к
' выбору траектории несущей способности, а не к самим нагрузкам LC.
Private Sub StoreInvalidCapacity(ByVal index As Long, ByVal reason As String)
    mCapacityStatus(index) = STATUS_INPUTERR
    mCapacityLimitState(index) = reason
    mCapacitySolutionMethod(index) = vbNullString
    mCapacityPathResolved(index) = CAPACITY_PATH_NONE
    mLambdaCapacity(index) = 0#
    mUtilCapacity(index) = 0#
    mNUltimate(index) = 0#
    mMxUltimate(index) = 0#
    mMyUltimate(index) = 0#
    mMomentUltimate(index) = 0#
End Sub

' Переносит результат CCapacitySolver в массивы batch-отчета.
' Здесь разделяется математическая сходимость capacity solver-а и простая
' инженерная проверка: lambda меньше или больше единицы.
Private Sub StoreCapacity(ByVal index As Long, ByVal cap As CCapacitySolver, ByVal loadPath As String)
    mCapacityStatus(index) = CapacityStatusFromSolver(cap)
    mCapacityLimitState(index) = cap.LimitState
    mCapacitySolutionMethod(index) = cap.SolutionMethod
    mCapacityPathResolved(index) = loadPath
    mLambdaCapacity(index) = cap.LambdaUltimate
    If cap.LambdaUltimate > 0# Then
        If CapacityPathScalesN(loadPath) Then mNUltimate(index) = cap.LambdaUltimate * mN(index)
        If CapacityPathScalesMx(loadPath) Then mMxUltimate(index) = cap.LambdaUltimate * mUserMx(index)
        If CapacityPathScalesMy(loadPath) Then mMyUltimate(index) = cap.LambdaUltimate * mUserMy(index)
    End If
    mMomentUltimate(index) = Sqr(mMxUltimate(index) * mMxUltimate(index) + mMyUltimate(index) * mMyUltimate(index))
    If cap.LambdaUltimate > 0# Then mUtilCapacity(index) = 1# / cap.LambdaUltimate
    If cap.Converged Then mUtilCapacity(index) = cap.Utilization

    ' h/x в блоке предельного момента должны относиться к найденному
    ' предельному состоянию, а не к фактическому НДС LC. Поэтому сохраняем
    ' отдельную плоскость деформаций только когда capacity-поиск действительно
    ' дал физически принятое предельное состояние.
    If cap.Converged And cap.LambdaUltimate > 0# Then
        If Not cap.LastSolver Is Nothing Then
            mCapacityStateAvailable(index) = True
            mCapacityEpsilon0(index) = cap.LastSolver.Epsilon0
            mCapacityKappaX(index) = cap.LastSolver.KappaX
            mCapacityKappaY(index) = cap.LastSolver.KappaY
        End If
    End If
End Sub

' Фиксирует InvalidInput по конкретному LC и не дает этому сочетанию пройти
' в solver. Так пользователь видит причину в таблице, а остальные LC считаются.
Private Sub MarkInvalid(ByVal index As Long, ByVal reason As String)
    If Len(Trim$(reason)) = 0 Then reason = "Некорректное сочетание нагрузок."
    mInputError(index) = reason
    mStatus(index) = STATUS_INPUTERR
    mDirectStateStatus(index) = STATUS_INPUTERR
    mCapacityStatus(index) = STATUS_INPUTERR
    mCapacityLimitState(index) = reason
    mCapacitySolutionMethod(index) = vbNullString
    mCapacityPathResolved(index) = CAPACITY_PATH_NONE
    mCrackStatus(index) = STATUS_NA
    If Not mMaterialProvider Is Nothing Then
        mDirectStateDiagramExtensionEnabled(index) = mMaterialProvider.DirectStateDiagramExtensionEnabled
    End If
    mGoverningUtil(index) = 0#
    AppendDiagnostic index
End Sub

' Передает в CCapacitySolver ровно те настройки, которые относятся к поиску
' несущей способности и внутренним пробным расчетам равновесия.
Private Sub ConfigureCapacity(ByVal cap As CCapacitySolver)
    cap.ConcreteCompressionLimit = mMaterialProvider.ConcreteCompressionLimit(cpStrength)
    cap.ConcreteTensionLimit = mMaterialProvider.ConcreteTensionLimit(cpStrength)
    cap.ConcreteTensionLimitEnabled = mMaterialProvider.ConcreteTensionLimitEnabled(cpStrength)
    cap.SteelStrainLimit = MaxDouble(Abs(mMaterialProvider.SteelCompressionLimit(cpStrength)), _
        Abs(mMaterialProvider.SteelTensionLimit(cpStrength)))
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

' Настраивает прямой CSectionSolver для расчета фактического состояния LC.
' Capacity-настройки сюда не попадают, чтобы прямое НДС не зависело от
' выбранного метода поиска предельного состояния.
Private Sub ConfigureSectionSolver(ByVal solver As CSectionSolver)
    solver.LoadSteps = mSolverLoadSteps
    solver.SolverMethod = mSolverMethod
    solver.MaxIterations = mSolverMaxIterations
    solver.ToleranceN = mSolverToleranceN
    solver.ToleranceMx = mSolverToleranceMx
    solver.ToleranceMy = mSolverToleranceMy
    solver.LineSearchEnabled = mSolverLineSearchEnabled
    solver.DampingInitial = mSolverDampingInitial
    solver.MinLineSearchAlpha = mSolverMinLineSearchAlpha
    solver.MaxDeltaEpsilon0 = mSolverMaxDeltaEpsilon0
    solver.MaxDeltaKappa = mSolverMaxDeltaKappa
    solver.SecantMaxRestarts = mSolverSecantMaxRestarts
    solver.SecantMinStepNorm = mSolverSecantMinStepNorm
    solver.DiagnosticsEnabled = DetailedDiagnosticsEnabled()
End Sub

' Передает в расчет трещин те же настройки прямого равновесия, что и в
' сервисный solve LC. Это важно для сопоставимости: состояние с раскрытой
' трещиной ищется тем же CSectionSolver, только с принудительно выключенным
' растянутым бетоном.
Private Sub ConfigureCrackCalculator(ByVal crack As CCrackWidthCalculator)
    crack.AllowableCrackWidth = mCrackAllowable
    crack.PsiMode = mCrackPsiMode
    crack.TensionZoneMode = mCrackTensionZoneMode
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
    If mCount <= 0 Then Err.Raise vbObjectError + 3921, "CBatchSectionCalculator", "Не задано ни одного сочетания нагрузок."
End Sub

' Сбрасывает все результаты одной строки перед повторным расчетом LC.
' Входные усилия и комментарий не трогаются.
Private Sub ClearResultRow(ByVal index As Long)
    mStatus(index) = vbNullString
    mDirectStateStatus(index) = vbNullString
    mCapacityStatus(index) = vbNullString
    mCapacityLimitState(index) = vbNullString
    mCapacitySolutionMethod(index) = vbNullString
    mCrackStatus(index) = vbNullString
    mLambdaCapacity(index) = 0#
    mUtilCapacity(index) = 0#
    mCapacityPathResolved(index) = CAPACITY_PATH_NONE
    mNUltimate(index) = 0#
    mMxUltimate(index) = 0#
    mMyUltimate(index) = 0#
    mCrackWidth(index) = 0#: mCrackAllowableValue(index) = 0#: mCrackUtil(index) = 0#
    mCrackFormed(index) = False
    mCrackPhi1(index) = 0#: mCrackPhi2(index) = 0#: mCrackPhi3(index) = 0#: mCrackPsiS(index) = 0#
    mCrackSigmaS(index) = 0#: mCrackSigmaSCrc(index) = 0#
    mCrackAs(index) = 0#: mCrackAbt(index) = 0#: mCrackDsEquivalent(index) = 0#
    mMomentUltimate(index) = 0#
    mCapacityStateAvailable(index) = False
    mCapacityEpsilon0(index) = 0#: mCapacityKappaX(index) = 0#: mCapacityKappaY(index) = 0#
    mMinConcreteStrain(index) = 0#: mMaxConcreteStrain(index) = 0#
    mMinSteelStrain(index) = 0#: mMaxSteelStrain(index) = 0#
    mCrackLsRaw(index) = 0#: mCrackLs(index) = 0#: mCrackLambdaCrc(index) = 0#: mCrackNcrc(index) = 0#
    mCrackTensionRebarCount(index) = 0#
    mCrackSectionDepthH(index) = 0#: mCrackCoverA(index) = 0#: mCrackTensionDepth(index) = 0#
    mCrackEffectiveZoneDepth(index) = 0#: mCrackCentralBranch(index) = False
    mGoverningUtil(index) = 0#
    mStateAvailable(index) = False
    mStateConverged(index) = False
    mEpsilon0(index) = 0#: mKappaX(index) = 0#: mKappaY(index) = 0#
    mNint(index) = 0#: mMxint(index) = 0#: mMyint(index) = 0#
    mResidualN(index) = 0#: mResidualMx(index) = 0#: mResidualMy(index) = 0#
    mStateIterations(index) = 0
    mExtensionUsed(index) = False
    mDirectStateDiagramExtensionEnabled(index) = False
End Sub

' Сохраняет найденную плоскость деформаций и внутренние усилия прямого solve.
' Эти данные затем уходят в Results и используются для схемы/AutoCAD export.
Private Sub StoreSectionState(ByVal index As Long, ByVal solver As CSectionSolver)
    mStateAvailable(index) = True
    mStateConverged(index) = solver.Converged
    mEpsilon0(index) = solver.Epsilon0
    mKappaX(index) = solver.KappaX
    mKappaY(index) = solver.KappaY
    mNint(index) = solver.Nint
    mMxint(index) = solver.Mxint
    mMyint(index) = solver.Myint
    mMinConcreteStrain(index) = solver.MinConcreteStrain
    mMaxConcreteStrain(index) = solver.MaxConcreteStrain
    mMinSteelStrain(index) = solver.MinSteelStrain
    mMaxSteelStrain(index) = solver.MaxSteelStrain
    mResidualN(index) = solver.ResidualN
    mResidualMx(index) = solver.ResidualMx
    mResidualMy(index) = solver.ResidualMy
    mStateIterations(index) = solver.Iterations
End Sub

' Добавляет одну строку в текстовую диагностику batch-расчета.
' Это компактная сводка для отладки, а не подробный журнал итераций solver-а.
Private Sub AppendDiagnostic(ByVal index As Long)
    mDiagnosticLog = mDiagnosticLog & "combination=" & mCombinationID(index) & _
        "; OverallStatus=" & mStatus(index) & _
        "; DirectStateStatus=" & mDirectStateStatus(index) & _
        "; ExtensionUsed=" & CStr(mExtensionUsed(index)) & _
        "; CapacityStatus=" & mCapacityStatus(index) & _
        "; CrackStatus=" & mCrackStatus(index) & _
        "; governingUtil=" & FormatNumberInvariant(mGoverningUtil(index)) & vbCrLf
End Sub

' Добавляет в подробный txt-отчет краткий итог по текущему LC.
' Batch summary остается основным табличным выводом, а здесь пользователь
' видит ход выполнения без открытия Results и без чтения внутренних массивов.
Private Sub ReportCombinationSummary(ByVal index As Long)
    ReportValue "Итоговый Status", mStatus(index)
    ReportValue "DirectStateStatus", mDirectStateStatus(index)
    ReportValue "ExtensionUsed", CStr(mExtensionUsed(index))
    ReportValue "CapacityStatus", mCapacityStatus(index)
    ReportValue "CrackStatus", mCrackStatus(index)
    ReportValue "GoverningUtilization", FormatNumberInvariant(mGoverningUtil(index))
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

' Собирает OverallStatus из статусов применимых расчетных разделов.
' N/A не ухудшает итог: это означает, что ветка не относится к LC или режиму.
Private Function AggregateOverallStatus(ByVal index As Long) As String
    Dim resultStatus As String
    resultStatus = STATUS_NA
    resultStatus = MergeOverallStatus(resultStatus, mDirectStateStatus(index))
    resultStatus = MergeOverallStatus(resultStatus, mCapacityStatus(index))
    resultStatus = MergeOverallStatus(resultStatus, mCrackStatus(index))
    AggregateOverallStatus = resultStatus
End Function

' Нормализует область запуска capacity-поиска. Алиасы оставлены на границе
' Excel-ввода, чтобы в расчетной логике дальше ходили только два значения.
Private Function NormalizedCapacityCalculationScope() As String
    Select Case LCase$(Trim$(mCapacityCalculationScope))
        Case "group1only", "group1", "uls"
            NormalizedCapacityCalculationScope = "Group1Only"
        Case "group1+2", "group1andgroup2", "group1+group2", "all", "both"
            NormalizedCapacityCalculationScope = "Group1AndGroup2"
    End Select
End Function

' Выбирает худший статус по утвержденному приоритету:
' InputErr -> NumFail -> FAIL -> OK, при этом N/A пропускается.
Private Function MergeOverallStatus(ByVal currentStatus As String, ByVal candidateStatus As String) As String
    candidateStatus = NormalizeUserStatus(candidateStatus)
    currentStatus = NormalizeUserStatus(currentStatus)
    If candidateStatus = STATUS_NA Then
        MergeOverallStatus = currentStatus
    ElseIf currentStatus = STATUS_NA Then
        MergeOverallStatus = candidateStatus
    ElseIf StatusPriority(candidateStatus) > StatusPriority(currentStatus) Then
        MergeOverallStatus = candidateStatus
    Else
        MergeOverallStatus = currentStatus
    End If
End Function

Private Function StatusPriority(ByVal statusText As String) As Long
    Select Case NormalizeUserStatus(statusText)
        Case STATUS_INPUTERR
            StatusPriority = 4
        Case STATUS_NUMFAIL
            StatusPriority = 3
        Case STATUS_FAIL
            StatusPriority = 2
        Case STATUS_OK
            StatusPriority = 1
        Case Else
            StatusPriority = 0
    End Select
End Function

' Нормализует внутренний текст в короткий пользовательский словарь статусов.
Private Function NormalizeUserStatus(ByVal statusText As String) As String
    Dim value As String
    value = Trim$(statusText)
    Select Case UCase$(value)
        Case "OK"
            NormalizeUserStatus = STATUS_OK
        Case "FAIL", "STRENGTHFAILED", "STRAINLIMITEXCEEDED"
            NormalizeUserStatus = STATUS_FAIL
        Case "NUMFAIL", "NUMERICALFAILURE", "SINGULARTANGENT"
            NormalizeUserStatus = STATUS_NUMFAIL
        Case "INPUTERR", "INPUTERROR", "INVALIDINPUT"
            NormalizeUserStatus = STATUS_INPUTERR
        Case "N/A", "NA", "NOTCALCULATED", "NOTCALCUL.", "SKIPPEDZEROMOMENT", "SKIPPEDZEROBASEMOMENT", ""
            NormalizeUserStatus = STATUS_NA
        Case Else
            If IsInvalidStatus(value) Then
                NormalizeUserStatus = STATUS_INPUTERR
            ElseIf IsNumericalStatus(value) Then
                NormalizeUserStatus = STATUS_NUMFAIL
            Else
                NormalizeUserStatus = STATUS_FAIL
            End If
    End Select
End Function

' True для расчетов, которые дали физически интерпретируемый результат:
' проверка либо прошла, либо показала недостаточную несущую способность.
Private Function IsFinishedStatus(ByVal statusText As String) As Boolean
    statusText = NormalizeUserStatus(statusText)
    IsFinishedStatus = (statusText = STATUS_OK Or statusText = STATUS_FAIL)
End Function

' Распознает статусы, означающие именно численную несходимость.
Private Function IsNumericalStatus(ByVal statusText As String) As Boolean
    IsNumericalStatus = (InStr(1, statusText, STATUS_NUMFAIL, vbTextCompare) > 0 Or _
        InStr(1, statusText, "NumericalFailure", vbTextCompare) > 0 Or _
        InStr(1, statusText, "SingularTangent", vbTextCompare) > 0)
End Function

' Делает вторую попытку прямого StateSolution, если первый Newton/Secant
' застрял на горизонтальном физическом плато. Диаграмма материала не меняется:
' стартовая плоскость лишь переносится ближе к technical extension, где уже
' есть малая касательная жесткость и solver может добрать формальное равновесие.
Private Function TrySolveStateWithExtensionWarmStart(ByVal index As Long, _
        ByVal concreteMaterial As CMaterialDiagram, ByVal steelMaterial As CMaterialDiagram, _
        ByRef service As CSectionSolver) As Boolean
    Dim guessBuilder As CStateGuessBuilder
    Set guessBuilder = New CStateGuessBuilder

    Dim eps0 As Double
    Dim kx As Double
    Dim ky As Double
    Dim attemptIndex As Long

    For attemptIndex = 1 To guessBuilder.ExtensionAttemptCount
        If guessBuilder.BuildAttemptForExtensionState(attemptIndex, mSection, service, _
                concreteMaterial, steelMaterial, mN(index), mMx(index), mMy(index), eps0, kx, ky) Then
            If TrySolveStateFromExtensionGuess(index, concreteMaterial, steelMaterial, _
                    eps0, kx, ky, guessBuilder.ExtensionAttemptName(attemptIndex), service) Then
                TrySolveStateWithExtensionWarmStart = True
                Exit Function
            End If
        End If
    Next attemptIndex
End Function

' Проверяет одну стартовую плоскость повторного StateSolution.
' Стартовые варианты строятся в CStateGuessBuilder, а здесь batch-слой
' запускает настоящий CSectionSolver и решает, подошла ли эта подсказка.
Private Function TrySolveStateFromExtensionGuess(ByVal index As Long, _
        ByVal concreteMaterial As CMaterialDiagram, ByVal steelMaterial As CMaterialDiagram, _
        ByVal eps0 As Double, ByVal kx As Double, ByVal ky As Double, _
        ByVal attemptName As String, ByRef service As CSectionSolver) As Boolean
    Dim startState As CSectionSolver
    Set startState = New CSectionSolver
    startState.EvaluateStrainPlane mSection, concreteMaterial, steelMaterial, eps0, kx, ky

    Dim retryLoadSteps As Long
    retryLoadSteps = mSolverLoadSteps
    If retryLoadSteps < 4 Then retryLoadSteps = 4

    If TryRunStateRetry(index, concreteMaterial, steelMaterial, eps0, kx, ky, _
            mN(index), mMx(index), mMy(index), 1, attemptName & " / direct target", service) Then
        TrySolveStateFromExtensionGuess = True
        Exit Function
    End If
    If TryRunStateRetry(index, concreteMaterial, steelMaterial, eps0, kx, ky, _
            mN(index), mMx(index), mMy(index), 1, attemptName & " / direct target without line search", service, True) Then
        TrySolveStateFromExtensionGuess = True
        Exit Function
    End If

    ' Повторный старт передает solver-у не только плоскость деформаций, но и
    ' усилия, которые эта плоскость уже создает. Тогда ступени идут от близкого
    ' extension-состояния к полному LC, а не от нулевой нагрузки из чужой
    ' деформированной формы; это особенно важно для перегруженного изгиба.
    TrySolveStateFromExtensionGuess = TryRunStateRetry(index, concreteMaterial, steelMaterial, eps0, kx, ky, _
        startState.Nint, startState.Mxint, startState.Myint, retryLoadSteps, attemptName & " / stepped", service)
End Function

' Запускает один повторный solve из подготовленной extension-плоскости.
' initialN/Mx/My определяют, от какого уже известного уровня нагрузки solver
' будет идти к полному LC: либо сразу от полного target, либо от усилий,
' созданных стартовой плоскостью.
Private Function TryRunStateRetry(ByVal index As Long, _
        ByVal concreteMaterial As CMaterialDiagram, ByVal steelMaterial As CMaterialDiagram, _
        ByVal eps0 As Double, ByVal kx As Double, ByVal ky As Double, _
        ByVal initialN As Double, ByVal initialMx As Double, ByVal initialMy As Double, _
        ByVal loadSteps As Long, ByVal attemptName As String, ByRef service As CSectionSolver, _
        Optional ByVal disableLineSearch As Boolean = False) As Boolean
    Dim retry As CSectionSolver
    Set retry = New CSectionSolver
    ConfigureSectionSolver retry
    If disableLineSearch Then retry.LineSearchEnabled = False
    retry.LoadSteps = loadSteps
    retry.SetInitialLoadState eps0, kx, ky, initialN, initialMx, initialMy
    ReportStep "Повторный старт StateSolution в extension: " & attemptName & "."
    retry.Solve mSection, concreteMaterial, steelMaterial, mN(index), mMx(index), mMy(index)

    If retry.Converged Then
        Set service = retry
        TryRunStateRetry = True
    Else
        mDiagnosticLog = mDiagnosticLog & "extensionRetryAttempt=" & attemptName & _
            "; extensionRetryStop=" & retry.StopReason & _
            "; retryResidualN=" & FormatNumberInvariant(retry.ResidualN) & _
            "; retryResidualMx=" & FormatNumberInvariant(retry.ResidualMx) & _
            "; retryResidualMy=" & FormatNumberInvariant(retry.ResidualMy) & vbCrLf
        ReportStep "Повторный старт в technical extension не сошелся (" & attemptName & "): " & retry.StopReason
        ReportBlock "Итерации несошедшегося повторного CSectionSolver", retry.DiagnosticLog
    End If
End Function

' Распознает статусы ошибок исходных данных.
Private Function IsInvalidStatus(ByVal statusText As String) As Boolean
    IsInvalidStatus = (InStr(1, statusText, STATUS_INPUTERR, vbTextCompare) > 0 Or _
        InStr(1, statusText, "InputError", vbTextCompare) > 0 Or _
        InStr(1, statusText, "InvalidInput", vbTextCompare) > 0)
End Function

' Переводит результат capacity solver-а в короткий пользовательский статус.
Private Function CapacityStatusFromSolver(ByVal cap As CCapacitySolver) As String
    If cap Is Nothing Then
        CapacityStatusFromSolver = STATUS_NUMFAIL
    ElseIf Not cap.Converged Then
        If IsInvalidStatus(cap.StopReason) Or IsInvalidStatus(cap.LimitState) Then
            CapacityStatusFromSolver = STATUS_INPUTERR
        Else
            CapacityStatusFromSolver = STATUS_NUMFAIL
        End If
    ElseIf cap.LambdaUltimate > 0# And cap.LambdaUltimate < 1# Then
        CapacityStatusFromSolver = STATUS_FAIL
    Else
        CapacityStatusFromSolver = STATUS_OK
    End If
End Function

' Переводит результат расчета трещин в тот же пользовательский словарь.
Private Function CrackStatusFromCalculator(ByVal crack As CCrackWidthCalculator) As String
    If crack Is Nothing Then
        CrackStatusFromCalculator = STATUS_NUMFAIL
    ElseIf crack.Converged Then
        If crack.CrackFormed And crack.AllowableCrackWidth > 0# And _
                crack.CrackWidth > crack.AllowableCrackWidth Then
            CrackStatusFromCalculator = STATUS_FAIL
        Else
            CrackStatusFromCalculator = STATUS_OK
        End If
    ElseIf IsInvalidStatus(crack.StopReason) Then
        CrackStatusFromCalculator = STATUS_INPUTERR
    Else
        CrackStatusFromCalculator = STATUS_NUMFAIL
    End If
End Function

' Проверяет, попала ли хотя бы одна деформация в техническую область extension.
Private Function StateUsesExtension(ByVal solver As CSectionSolver, _
        ByVal concreteMaterial As CMaterialDiagram, ByVal steelMaterial As CMaterialDiagram) As Boolean
    Dim i As Long
    Dim strain As Double

    For i = 1 To mSection.ConcreteCount
        strain = ElementStrain(solver, mSection.ConcreteX(i), mSection.ConcreteY(i))
        If concreteMaterial.IsInExtensionRange(strain) Then
            StateUsesExtension = True
            Exit Function
        End If
    Next i

    For i = 1 To mSection.RebarCount
        strain = ElementStrain(solver, mSection.RebarX(i), mSection.RebarY(i))
        If steelMaterial.IsInExtensionRange(strain) Then
            StateUsesExtension = True
            Exit Function
        End If
    Next i
End Function

' Проверяет физические пределы StateSolution. Бетонная растянутая зона с
' отключенным сопротивлением допускается: CMaterialDiagram сам знает, есть ли
' у конкретной диаграммы физический растянутый eps-предел.
Private Function StateWithinPhysicalRange(ByVal solver As CSectionSolver, _
        ByVal concreteMaterial As CMaterialDiagram, ByVal steelMaterial As CMaterialDiagram) As Boolean
    Dim i As Long
    Dim strain As Double

    For i = 1 To mSection.ConcreteCount
        strain = ElementStrain(solver, mSection.ConcreteX(i), mSection.ConcreteY(i))
        If Not concreteMaterial.IsInPhysicalRange(strain) Then Exit Function
    Next i

    For i = 1 To mSection.RebarCount
        strain = ElementStrain(solver, mSection.RebarX(i), mSection.RebarY(i))
        If Not steelMaterial.IsInPhysicalRange(strain) Then Exit Function
    Next i

    StateWithinPhysicalRange = True
End Function

Private Function ElementStrain(ByVal solver As CSectionSolver, ByVal x As Double, ByVal y As Double) As Double
    ElementStrain = solver.Epsilon0 + solver.KappaX * y + solver.KappaY * x
End Function

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
' После перехода на CapacityLoadPath это единая величина для λ*Mx, λ*My,
' λ*Mxy, λ*N и λ*NMxy; отдельные отношения Nult/N или Mult/Mxy больше
' не нужны для выбора Worst.
Private Function StrengthSafetyFactorForGoverning(ByVal index As Long) As Double
    If IsFinishedStatus(mCapacityStatus(index)) And mLambdaCapacity(index) > 0# Then
        StrengthSafetyFactorForGoverning = mLambdaCapacity(index)
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
    If mCrackUtil(index) > 0# Then CrackSafetyFactorForGoverning = 1# / mCrackUtil(index)
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
    CombinationID = mCombinationID(index)
End Property

Public Property Get CombinationName(ByVal index As Long) As String
    CombinationName = mComment(index)
End Property

Public Property Get CalculationType(ByVal index As Long) As String
    CalculationType = mCalculationType(index)
End Property

Public Property Get CapacityLoadPath(ByVal index As Long) As String
    CapacityLoadPath = CapacityLoadPathDisplay(CapacityLoadPathForCombination(index))
End Property

Public Property Get CapacityLoadPathKey(ByVal index As Long) As String
    CapacityLoadPathKey = CapacityLoadPathForCombination(index)
End Property

Public Property Get N(ByVal index As Long) As Double
    N = mN(index)
End Property

Public Property Get Mx(ByVal index As Long) As Double
    Mx = mMx(index)
End Property

Public Property Get My(ByVal index As Long) As Double
    My = mMy(index)
End Property

Public Property Get UserMx(ByVal index As Long) As Double
    UserMx = mUserMx(index)
End Property

Public Property Get UserMy(ByVal index As Long) As Double
    UserMy = mUserMy(index)
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

Public Property Get Status(ByVal index As Long) As String
    Status = mStatus(index)
End Property

Public Property Get OverallStatus(ByVal index As Long) As String
    OverallStatus = mStatus(index)
End Property

Public Property Get DirectStateStatus(ByVal index As Long) As String
    DirectStateStatus = mDirectStateStatus(index)
End Property

Public Property Get CapacityStatus(ByVal index As Long) As String
    CapacityStatus = mCapacityStatus(index)
End Property

Public Property Get CapacityLimitState(ByVal index As Long) As String
    CapacityLimitState = mCapacityLimitState(index)
End Property

Public Property Get CapacitySolutionMethod(ByVal index As Long) As String
    CapacitySolutionMethod = mCapacitySolutionMethod(index)
End Property

Public Property Get CrackStatus(ByVal index As Long) As String
    CrackStatus = mCrackStatus(index)
End Property

Public Property Get ExtensionUsed(ByVal index As Long) As Boolean
    ExtensionUsed = mExtensionUsed(index)
End Property

Public Property Get DirectStateDiagramExtensionEnabled(ByVal index As Long) As Boolean
    DirectStateDiagramExtensionEnabled = mDirectStateDiagramExtensionEnabled(index)
End Property

Public Property Get UtilizationCapacity(ByVal index As Long) As Double
    UtilizationCapacity = mUtilCapacity(index)
End Property

Public Property Get NUltimate(ByVal index As Long) As Double
    NUltimate = mNUltimate(index)
End Property

Public Property Get MxUltimate(ByVal index As Long) As Double
    MxUltimate = mMxUltimate(index)
End Property

Public Property Get MyUltimate(ByVal index As Long) As Double
    MyUltimate = mMyUltimate(index)
End Property

Public Property Get MomentUltimate(ByVal index As Long) As Double
    MomentUltimate = mMomentUltimate(index)
End Property

Public Property Get CapacityStateAvailable(ByVal index As Long) As Boolean
    CapacityStateAvailable = mCapacityStateAvailable(index)
End Property

Public Property Get CapacityEpsilon0(ByVal index As Long) As Double
    CapacityEpsilon0 = mCapacityEpsilon0(index)
End Property

Public Property Get CapacityKappaX(ByVal index As Long) As Double
    CapacityKappaX = mCapacityKappaX(index)
End Property

Public Property Get CapacityKappaY(ByVal index As Long) As Double
    CapacityKappaY = mCapacityKappaY(index)
End Property

Public Property Get MinConcreteStrain(ByVal index As Long) As Double
    MinConcreteStrain = mMinConcreteStrain(index)
End Property

Public Property Get MaxConcreteStrain(ByVal index As Long) As Double
    MaxConcreteStrain = mMaxConcreteStrain(index)
End Property

Public Property Get MinSteelStrain(ByVal index As Long) As Double
    MinSteelStrain = mMinSteelStrain(index)
End Property

Public Property Get MaxSteelStrain(ByVal index As Long) As Double
    MaxSteelStrain = mMaxSteelStrain(index)
End Property

Public Property Get CrackWidth(ByVal index As Long) As Double
    CrackWidth = mCrackWidth(index)
End Property

Public Property Get CrackAllowable(ByVal index As Long) As Double
    CrackAllowable = mCrackAllowableValue(index)
End Property

Public Property Get CrackUtilization(ByVal index As Long) As Double
    CrackUtilization = mCrackUtil(index)
End Property

Public Property Get CrackFormed(ByVal index As Long) As Boolean
    CrackFormed = mCrackFormed(index)
End Property

Public Property Get CrackPhi1(ByVal index As Long) As Double
    CrackPhi1 = mCrackPhi1(index)
End Property

Public Property Get CrackPhi2(ByVal index As Long) As Double
    CrackPhi2 = mCrackPhi2(index)
End Property

Public Property Get CrackPhi3(ByVal index As Long) As Double
    CrackPhi3 = mCrackPhi3(index)
End Property

Public Property Get CrackPsiS(ByVal index As Long) As Double
    CrackPsiS = mCrackPsiS(index)
End Property

Public Property Get CrackSigmaS(ByVal index As Long) As Double
    CrackSigmaS = mCrackSigmaS(index)
End Property

Public Property Get CrackSigmaSCrc(ByVal index As Long) As Double
    CrackSigmaSCrc = mCrackSigmaSCrc(index)
End Property

Public Property Get CrackAs(ByVal index As Long) As Double
    CrackAs = mCrackAs(index)
End Property

Public Property Get CrackAbt(ByVal index As Long) As Double
    CrackAbt = mCrackAbt(index)
End Property

Public Property Get CrackDsEquivalent(ByVal index As Long) As Double
    CrackDsEquivalent = mCrackDsEquivalent(index)
End Property

Public Property Get CrackSpacingRaw(ByVal index As Long) As Double
    CrackSpacingRaw = mCrackLsRaw(index)
End Property

Public Property Get CrackSpacing(ByVal index As Long) As Double
    CrackSpacing = mCrackLs(index)
End Property

Public Property Get CrackLambdaCrc(ByVal index As Long) As Double
    CrackLambdaCrc = mCrackLambdaCrc(index)
End Property

Public Property Get CrackNcrc(ByVal index As Long) As Double
    CrackNcrc = mCrackNcrc(index)
End Property

Public Property Get CrackTensionRebarCount(ByVal index As Long) As Long
    CrackTensionRebarCount = mCrackTensionRebarCount(index)
End Property

Public Property Get CrackSectionDepthH(ByVal index As Long) As Double
    CrackSectionDepthH = mCrackSectionDepthH(index)
End Property

Public Property Get CrackCoverA(ByVal index As Long) As Double
    CrackCoverA = mCrackCoverA(index)
End Property

Public Property Get CrackTensionDepth(ByVal index As Long) As Double
    CrackTensionDepth = mCrackTensionDepth(index)
End Property

Public Property Get CrackEffectiveZoneDepth(ByVal index As Long) As Double
    CrackEffectiveZoneDepth = mCrackEffectiveZoneDepth(index)
End Property

Public Property Get CrackCentralBranch(ByVal index As Long) As Boolean
    CrackCentralBranch = mCrackCentralBranch(index)
End Property

Public Property Get GoverningUtilization(ByVal index As Long) As Double
    GoverningUtilization = mGoverningUtil(index)
End Property

Public Property Get CrackSteelElasticModulus() As Double
    CrackSteelElasticModulus = mMaterialProvider.SteelElasticModulus(cpCrackedNDS)
End Property

' Предельная сжатая деформация бетона для фактического НДС выбранного LC.
Public Function ConcreteCompressionLimitForCombination(ByVal index As Long) As Double
    ConcreteCompressionLimitForCombination = mMaterialProvider.ConcreteCompressionLimit(MaterialPurposeForCombination(index))
End Function

' Предельная растянутая деформация бетона для фактического НДС выбранного LC.
Public Function ConcreteTensionLimitForCombination(ByVal index As Long) As Double
    ConcreteTensionLimitForCombination = mMaterialProvider.ConcreteTensionLimit(MaterialPurposeForCombination(index))
End Function

' True, если текущая диаграмма LC учитывает растянутый бетон и его предел нужно выводить.
Public Function ConcreteTensionLimitEnabledForCombination(ByVal index As Long) As Boolean
    ConcreteTensionLimitEnabledForCombination = mMaterialProvider.ConcreteTensionLimitEnabled(MaterialPurposeForCombination(index))
End Function

' Предельная сжатая деформация арматуры для фактического НДС выбранного LC.
Public Function SteelCompressionLimitForCombination(ByVal index As Long) As Double
    SteelCompressionLimitForCombination = mMaterialProvider.SteelCompressionLimit(MaterialPurposeForCombination(index))
End Function

' Предельная растянутая деформация арматуры для фактического НДС выбранного LC.
Public Function SteelTensionLimitForCombination(ByVal index As Long) As Double
    SteelTensionLimitForCombination = mMaterialProvider.SteelTensionLimit(MaterialPurposeForCombination(index))
End Function

Public Property Get GoverningCombinationIndex() As Long
    GoverningCombinationIndex = mGoverningCombinationIndex
End Property

Public Property Get GoverningCombinationID() As String
    If mGoverningCombinationIndex > 0 Then GoverningCombinationID = mCombinationID(mGoverningCombinationIndex)
End Property

Public Property Get GoverningCombinationName() As String
    If mGoverningCombinationIndex > 0 Then GoverningCombinationName = mComment(mGoverningCombinationIndex)
End Property

Public Property Get CrackGoverningCombinationIndex() As Long
    CrackGoverningCombinationIndex = mCrackGoverningCombinationIndex
End Property

Public Property Get CrackGoverningCombinationID() As String
    If mCrackGoverningCombinationIndex > 0 Then CrackGoverningCombinationID = mCombinationID(mCrackGoverningCombinationIndex)
End Property

Public Property Get CrackGoverningCombinationName() As String
    If mCrackGoverningCombinationIndex > 0 Then CrackGoverningCombinationName = mComment(mCrackGoverningCombinationIndex)
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
        If IsInvalidStatus(mStatus(i)) Then InvalidInputCount = InvalidInputCount + 1
    Next i
End Property

Public Property Get FirstInvalidInputMessage() As String
    Dim i As Long
    For i = 1 To mCount
        If IsInvalidStatus(mStatus(i)) Then
            FirstInvalidInputMessage = mInputError(i)
            If Len(Trim$(FirstInvalidInputMessage)) = 0 Then FirstInvalidInputMessage = mCapacityLimitState(i)
            If Len(Trim$(FirstInvalidInputMessage)) = 0 Then FirstInvalidInputMessage = mStatus(i)
            Exit Property
        End If
    Next i
End Property

Public Property Get StateAvailable(ByVal index As Long) As Boolean
    StateAvailable = mStateAvailable(index)
End Property

Public Property Get StateAvailableCount() As Long
    Dim i As Long
    For i = 1 To mCount
        If mStateAvailable(i) Then StateAvailableCount = StateAvailableCount + 1
    Next i
End Property

Public Property Get Epsilon0(ByVal index As Long) As Double
    Epsilon0 = mEpsilon0(index)
End Property

Public Property Get KappaX(ByVal index As Long) As Double
    KappaX = mKappaX(index)
End Property

Public Property Get KappaY(ByVal index As Long) As Double
    KappaY = mKappaY(index)
End Property

Public Property Get Nint(ByVal index As Long) As Double
    Nint = mNint(index)
End Property

Public Property Get Mxint(ByVal index As Long) As Double
    Mxint = mMxint(index)
End Property

Public Property Get Myint(ByVal index As Long) As Double
    Myint = mMyint(index)
End Property

Public Property Get ResidualN(ByVal index As Long) As Double
    ResidualN = mResidualN(index)
End Property

Public Property Get ResidualMx(ByVal index As Long) As Double
    ResidualMx = mResidualMx(index)
End Property

Public Property Get ResidualMy(ByVal index As Long) As Double
    ResidualMy = mResidualMy(index)
End Property

Public Property Get StateIterations(ByVal index As Long) As Long
    StateIterations = mStateIterations(index)
End Property

' ==========================================================================
' ДЛЯ ТЕСТОВ
' ==========================================================================
' Эти свойства раскрывают промежуточные численные данные batch-расчета,
' которые удобны regression-тестам: отдельно проверяются lambda capacity
' и факт сходимости прямого НДС.

Public Property Get LambdaCapacity(ByVal index As Long) As Double
    LambdaCapacity = mLambdaCapacity(index)
End Property

Public Property Get StateConverged(ByVal index As Long) As Boolean
    StateConverged = mStateConverged(index)
End Property



