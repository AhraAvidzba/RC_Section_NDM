VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCapacitySolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Поиск несущей способности сечения
' ==========================================================================
' Класс определяет предельное состояние поверх готового решателя равновесия.
' Само равновесие epsilon0/kappaX/kappaY всегда ищет CSectionSolver, а здесь
' решается внешняя задача: какой предельный множитель или какая предельная
' плоскость деформаций соответствует заданному N + Mx + My.
'
' На вход класса теперь передается универсальная траектория нагрузки:
' N = N0 + lambda*Nbase, Mx = Mx0 + lambda*Mxbase, My = My0 + lambda*Mybase.
' Это позволяет одним алгоритмом считать lambda*Mx, lambda*My, lambda*Mxy,
' lambda*N и lambda*NMxy,
' не плодя отдельные пользовательские ветви внутри расчетного ядра.

Private Const STATE_ACCEPTABLE As String = "Acceptable"
Private Const STATE_CONCRETE_LIMIT As String = "ConcreteStrainLimit"
Private Const STATE_CONCRETE_TENSION_LIMIT As String = "ConcreteTensionStrainLimit"
Private Const STATE_STEEL_LIMIT As String = "SteelStrainLimit"
Private Const STATE_NUMERICAL_FAILURE As String = "NumericalFailure"
Private Const STATE_SINGULAR_TANGENT As String = "SingularTangent"
Private Const STATE_INVALID_INPUT As String = "InvalidInput"
Private Const MOMENT_DIRECTION_TOLERANCE As Double = 0.001 ' Допустимое относительное отклонение от выбранной моментной траектории.
Private Const PROBE_CACHE_SIZE As Long = 12 ' Небольшое кольцо последних probe-точек LoadMultiplier.

Private mConverged As Boolean ' True, если выбранный метод capacity нашел предельное состояние.
Private mStopReason As String ' Причина остановки capacity solver для вывода в результаты и тесты.
Private mLimitState As String ' Физический или численный статус найденного предела.
Private mUpperLimitState As String ' Статус верхней границы при LoadMultiplier.
Private mLambdaUltimate As Double ' Предельный множитель к выбранной lambda-траектории нагрузки.
Private mNUltimate As Double ' Предельная продольная сила, если N входит в масштабируемый путь.
Private mMomentUltimate As Double ' Результирующий предельный момент sqrt(MxUltimate^2 + MyUltimate^2).
Private mMxUltimate As Double ' Предельный Mx, если Mx входит в масштабируемый путь.
Private mMyUltimate As Double ' Предельный My, если My входит в масштабируемый путь.
Private mSafetyFactor As Double ' Коэффициент запаса по capacity; сейчас совпадает с lambdaUltimate.
Private mUtilization As Double ' Использование несущей способности: 1 / lambdaUltimate.
Private mIterations As Long ' Число внешних итераций поиска несущей способности.
Private mRetryCount As Long ' Число повторов после численных отказов внутреннего решателя.
Private mLastSolver As CSectionSolver ' Последний внутренний расчет равновесия, успешный или диагностический.
Private mLastAcceptableSolver As CSectionSolver ' Последнее состояние, которое еще не вышло за физический предел.
Private mLastAcceptableN As Double ' N последнего acceptable-состояния.
Private mLastAcceptableMx As Double ' Mx последнего acceptable-состояния.
Private mLastAcceptableMy As Double ' My последнего acceptable-состояния.
Private mConcreteCompressionLimit As Double ' Предельная сжимающая деформация бетона.
Private mConcreteTensionLimit As Double ' Предельная растягивающая деформация бетона, если растянутая ветвь включена материалной моделью.
Private mConcreteTensionLimitEnabled As Boolean ' True, когда растянутый бетон реально ограничивает предельное состояние.
Private mSteelStrainLimit As Double ' Предельная деформация арматуры по модулю.
Private mLambdaTolerance As Double ' Точность одномерного поиска lambda для LoadMultiplier.
Private mStrainTolerance As Double ' Допуск достижения предельной деформации для UltimateStrain.
Private mSearchMethod As String ' Алгоритм одномерного поиска LoadMultiplier: Bisection, Brent или Secant.
Private mSolutionMethod As String ' Фактическая ветка, которая дала принятый результат: UltimateStrain или LoadMultiplier.
Private mMaxLambda As Double ' Верхняя граница расширения скобки LoadMultiplier, чтобы поиск не уходил бесконечно.
Private mInitialLambdaStep As Double ' Первый шаг расширения lambda при поиске расчетной скобки.
Private mMaxRetries As Long ' Допустимые повторы пробной точки после NumericalFailure.
Private mSolverBaseLoadSteps As Long ' Базовое число ступеней CSectionSolver для каждой probe-точки.
Private mSolverMaxIterations As Long ' Максимум внешних/внутренних итераций capacity-поиска.
Private mSolverToleranceN As Double ' Допуск внутреннего равновесия CSectionSolver по N.
Private mSolverToleranceMx As Double ' Допуск внутреннего равновесия CSectionSolver по Mx.
Private mSolverToleranceMy As Double ' Допуск внутреннего равновесия CSectionSolver по My.
Private mSolverMethod As String ' Метод внутреннего CSectionSolver: Newton или Secant.
Private mSolverLineSearchEnabled As Boolean ' Общая настройка line search для внутренних расчетов равновесия.
Private mSolverDampingInitial As Double ' Начальный alpha line search во внутренних расчетах равновесия.
Private mSolverMinLineSearchAlpha As Double ' Минимальный alpha line search во внутренних расчетах равновесия.
Private mSolverMaxDeltaEpsilon0 As Double ' Ограничение шага epsilon0 для внутреннего решателя.
Private mSolverMaxDeltaKappa As Double ' Ограничение шага кривизн для внутреннего решателя.
Private mSolverSecantMaxRestarts As Long ' Настройка рестартов внутреннего secant-решателя.
Private mSolverSecantMinStepNorm As Double ' Минимальный шаг внутреннего secant-решателя.
Private mMomentEquilibriumResidual As Double ' Невязка направления/момента в найденном предельном состоянии.
Private mDiagnosticLog As String ' Текстовая трассировка probe-точек и причин остановки.
Private mDiagnosticsEnabled As Boolean ' Отключает подробный журнал в массовом пользовательском расчете без txt-отчета.
Private mCriticalStrain As Double ' Деформация критического волокна или стержня.
Private mCriticalElement As String ' ID элемента, который первым определил предельное состояние.
Private mCriticalStrainUtilization As Double ' Отношение фактической деформации к своему пределу.
Private mProbeCacheCount As Long ' Сколько ячеек cache реально заполнено.
Private mProbeCacheNextSlot As Long ' Следующая ячейка кольцевого cache для записи.
Private mProbeCacheLambda(1 To PROBE_CACHE_SIZE) As Double ' Lambda ранее посчитанной probe-точки.
Private mProbeCacheN(1 To PROBE_CACHE_SIZE) As Double ' N этой probe-точки; нужен, чтобы не спутать разные траектории.
Private mProbeCacheMx(1 To PROBE_CACHE_SIZE) As Double ' Mx этой probe-точки.
Private mProbeCacheMy(1 To PROBE_CACHE_SIZE) As Double ' My этой probe-точки.
Private mProbeCacheState(1 To PROBE_CACHE_SIZE) As String ' Классификация probe-точки: acceptable/limit/numerical.
Private mProbeCacheCriticalStrain(1 To PROBE_CACHE_SIZE) As Double ' Критическая деформация, восстановленная без повторного solve.
Private mProbeCacheCriticalElement(1 To PROBE_CACHE_SIZE) As String ' Критический элемент, соответствующий cache-записи.
Private mProbeCacheCriticalUtil(1 To PROBE_CACHE_SIZE) As Double ' Utilization критической деформации для целевой функции.
Private mProbeCacheSolver(1 To PROBE_CACHE_SIZE) As CSectionSolver ' Solver-состояние probe-точки для дальнейшего warm-start.

' ------------------------------
' Настройки по умолчанию
' ------------------------------
' Эти значения используются в модульных тестах и как запасной набор, если
' книга еще не передала настройки с листа Config.
Private Sub Class_Initialize()
    mConcreteCompressionLimit = -0.0035
    mConcreteTensionLimit = 0.00015
    mConcreteTensionLimitEnabled = False
    mSteelStrainLimit = 0.025
    mLambdaTolerance = 0.001
    mStrainTolerance = 0.00001
    mSearchMethod = "Bisection"
    mMaxLambda = 1024#
    mInitialLambdaStep = 1#
    mMaxRetries = 4
    mSolverBaseLoadSteps = 8
    mSolverMaxIterations = 60
    mSolverToleranceN = 5#
    mSolverToleranceMx = 5000#
    mSolverToleranceMy = 5000#
    mSolverMethod = "Newton"
    mSolverLineSearchEnabled = True
    mSolverDampingInitial = 1#
    mSolverMinLineSearchAlpha = 0.03125
    mSolverMaxDeltaEpsilon0 = 0.0003
    mSolverMaxDeltaKappa = 0.000005
    mSolverSecantMaxRestarts = 2
    mSolverSecantMinStepNorm = 0.000000000001
    mDiagnosticsEnabled = True
    mProbeCacheNextSlot = 1
End Sub

' ------------------------------
' Чтение настроек capacity solver
' ------------------------------
' Настройки делятся на два уровня: внешний поиск несущей способности и
' внутренний CSectionSolver, который каждый раз ищет равновесие для probe-точки.
Public Sub ApplySettings(ByVal settings As CSystemSettingsReader, Optional ByVal units As CUnitSystem = Nothing)
    If settings Is Nothing Then Exit Sub
    mInitialLambdaStep = settings.GetDouble("Capacity.InitialLambda", mInitialLambdaStep)
    mMaxLambda = settings.GetDouble("Capacity.MaxLambda", mMaxLambda)
    mLambdaTolerance = settings.GetDouble("Capacity.ToleranceLambda", mLambdaTolerance)
    mStrainTolerance = settings.GetDouble("Capacity.ToleranceStrain", mStrainTolerance)
    If settings.HasKey("Capacity.SearchMethod") Then
        mSearchMethod = settings.GetRawString("Capacity.SearchMethod", vbNullString)
    Else
        mSearchMethod = "Bisection"
    End If
    mMaxRetries = settings.GetLong("Capacity.MaxRetries", mMaxRetries)
    mSolverBaseLoadSteps = settings.GetLong("Capacity.BaseLoadSteps", mSolverBaseLoadSteps)
    mSolverMaxIterations = settings.GetLong("Capacity.SolverMaxIterations", mSolverMaxIterations)
    mSolverToleranceN = settings.GetDouble("Solver.ToleranceN", mSolverToleranceN)
    mSolverToleranceMx = settings.GetDouble("Solver.ToleranceMx", mSolverToleranceMx)
    mSolverToleranceMy = settings.GetDouble("Solver.ToleranceMy", mSolverToleranceMy)
    If settings.HasKey("Solver.Method") Then
        mSolverMethod = settings.GetRawString("Solver.Method", vbNullString)
    Else
        mSolverMethod = "Newton"
    End If
    mSolverLineSearchEnabled = settings.GetBoolean("Solver.LineSearchEnabled", mSolverLineSearchEnabled)
    mSolverDampingInitial = settings.GetDouble("Solver.DampingInitial", mSolverDampingInitial)
    mSolverMinLineSearchAlpha = settings.GetDouble("Solver.MinLineSearchAlpha", mSolverMinLineSearchAlpha)
    mSolverMaxDeltaEpsilon0 = settings.GetDouble("Solver.MaxDeltaEpsilon0", mSolverMaxDeltaEpsilon0)
    mSolverMaxDeltaKappa = settings.GetDouble("Solver.MaxDeltaKappa", mSolverMaxDeltaKappa)
    mSolverSecantMaxRestarts = settings.GetLong("Solver.SecantMaxRestarts", mSolverSecantMaxRestarts)
    mSolverSecantMinStepNorm = settings.GetDouble("Solver.SecantMinStepNorm", mSolverSecantMinStepNorm)
    If Not units Is Nothing Then
        mSolverToleranceN = Abs(units.InputForceToInternal(mSolverToleranceN))
        mSolverToleranceMx = Abs(units.InputMomentMxToInternal(mSolverToleranceMx))
        mSolverToleranceMy = Abs(units.InputMomentMyToInternal(mSolverToleranceMy))
        mSolverMaxDeltaKappa = Abs(units.InputCurvatureToInternal(mSolverMaxDeltaKappa))
    End If
End Sub

' ------------------------------
' Метод UltimateStrain
' ------------------------------
' Ищет плоскость деформаций, при которой одновременно выполнены три условия:
' равновесие по N, сохранение направления Mx/My и достижение одного из
' пределов деформаций бетона или арматуры.
Public Sub SolveByUltimateStrain(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal nSpecified As Double, ByVal mxBase As Double, ByVal myBase As Double, _
        Optional ByVal mxOffset As Double = 0#, Optional ByVal myOffset As Double = 0#)
    SolveByUltimateLoadPath section, concreteMaterial, steelMaterial, _
        nSpecified, 0#, mxOffset, mxBase, myOffset, myBase
End Sub

' Ищет предельное состояние для произвольной lambda-траектории.
' Этот метод ничего не знает о названии пользовательского режима: он работает
' только с прямой формулой
' Target = Offset + lambda*Base для N, Mx и My.
Public Sub SolveByUltimateLoadPath(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal nOffset As Double, ByVal nBase As Double, _
        ByVal mxOffset As Double, ByVal mxBase As Double, _
        ByVal myOffset As Double, ByVal myBase As Double, _
        Optional ByVal allowPureAxialRedirect As Boolean = True)
    ClearResult
    mSolutionMethod = "UltimateStrain"

    If Not ValidateLoadPathCapacityInputs(section, concreteMaterial, steelMaterial, nBase, mxBase, myBase) Then Exit Sub
    If Not ValidateUltimateStrainSettings() Then Exit Sub
    If allowPureAxialRedirect And IsPureAxialLoadPath(nOffset, nBase, mxOffset, mxBase, myOffset, myBase) Then
        SolvePureAxialUltimateByLoadMultiplier section, concreteMaterial, steelMaterial, _
            nOffset, nBase, mxOffset, mxBase, myOffset, myBase
        Exit Sub
    End If
    If IsMomentOnlyLoadPath(nBase, mxBase, myBase) Then
        SolveMomentUltimateStrainPath section, concreteMaterial, steelMaterial, _
            nOffset, mxBase, myBase, mxOffset, myOffset
        Exit Sub
    End If

    Dim eps0 As Double
    Dim kx As Double
    Dim ky As Double
    Dim iter As Long
    Dim rPath1 As Double
    Dim rPath2 As Double
    Dim rLimit As Double
    Dim lambdaValue As Double
    Dim state As String
    Dim solver As CSectionSolver

    InitializeUltimateLoadPathGuess section, nOffset, nBase, mxOffset, mxBase, myOffset, myBase, eps0, kx, ky

    For iter = 1 To mSolverMaxIterations
        If Not EvaluateUltimateLoadPathResidual(section, concreteMaterial, steelMaterial, _
                nOffset, nBase, mxOffset, mxBase, myOffset, myBase, _
                eps0, kx, ky, solver, rPath1, rPath2, rLimit, lambdaValue, state) Then Exit For

        If Abs(rPath1) <= 1# And Abs(rPath2) <= 1# And Abs(rLimit) <= UltimateStrainRatioTolerance() Then
            Set mLastSolver = solver
            If Not IsValidUltimateLambda(lambdaValue, nOffset, nBase, mxBase, myBase, mxOffset, myOffset) Then
                mLimitState = STATE_NUMERICAL_FAILURE
                mStopReason = "UltimateStrain: найден недопустимый lambda: " & FormatNumberInvariant(lambdaValue)
                Exit Sub
            End If
            StoreLoadPathUltimateComponents lambdaValue, nOffset, nBase, mxBase, myBase, mxOffset, myOffset
            mLimitState = state
            mUpperLimitState = state
            mConverged = True
            Exit For
        End If

        If Not ApplyUltimateLoadPathNewtonStep(section, concreteMaterial, steelMaterial, _
                nOffset, nBase, mxOffset, mxBase, myOffset, myBase, _
                eps0, kx, ky, rPath1, rPath2, rLimit) Then Exit For
    Next iter

    If Not mConverged Then
        If Len(mLimitState) = 0 Then mLimitState = STATE_NUMERICAL_FAILURE
        If Len(mStopReason) = 0 Then mStopReason = "UltimateStrain: итерации Newton не сошлись."
        Exit Sub
    End If

    If Not mLastSolver Is Nothing Then
        mMomentEquilibriumResidual = Sqr((mLastSolver.Mxint - TargetMx(mLambdaUltimate, mxBase, mxOffset)) ^ 2 + _
            (mLastSolver.Myint - TargetMy(mLambdaUltimate, myBase, myOffset)) ^ 2)
    End If
    mSafetyFactor = mLambdaUltimate
    If mLambdaUltimate > 0# Then mUtilization = 1# / mLambdaUltimate
    mStopReason = "Поиск несущей способности через UltimateStrain сошелся."
End Sub

' Быстрая ветка UltimateStrain для моментных траекторий.
' Внешний API остается универсальным, но математически этот частный случай
' должен решаться в прямой моментной постановке: продольная сила фиксирована, а Mx/My
' масштабируются с сохранением направления моментного вектора.
Private Sub SolveMomentUltimateStrainPath(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal nSpecified As Double, ByVal mxBase As Double, ByVal myBase As Double, _
        ByVal mxOffset As Double, ByVal myOffset As Double)
    Dim eps0 As Double
    Dim kx As Double
    Dim ky As Double
    Dim iter As Long
    Dim rN As Double
    Dim rDir As Double
    Dim rLimit As Double
    Dim lambdaValue As Double
    Dim state As String
    Dim solver As CSectionSolver

    InitializeUltimateStrainGuess section, nSpecified, mxBase, myBase, eps0, kx, ky

    For iter = 1 To mSolverMaxIterations
        If Not EvaluateUltimateResidual(section, concreteMaterial, steelMaterial, _
                nSpecified, mxBase, myBase, mxOffset, myOffset, _
                eps0, kx, ky, solver, rN, rDir, rLimit, lambdaValue, state) Then Exit Sub

        If Abs(rN) <= 1# And Abs(rDir) <= MOMENT_DIRECTION_TOLERANCE And _
                Abs(rLimit) <= UltimateStrainRatioTolerance() Then
            Set mLastSolver = solver
            If Not IsValidUltimateLambda(lambdaValue, nSpecified, 0#, mxBase, myBase, mxOffset, myOffset) Then
                mLimitState = STATE_NUMERICAL_FAILURE
                mStopReason = "UltimateStrain: найден недопустимый lambda: " & FormatNumberInvariant(lambdaValue)
                Exit Sub
            End If
            StoreLoadPathUltimateComponents lambdaValue, nSpecified, 0#, mxBase, myBase, mxOffset, myOffset
            mLimitState = state
            mUpperLimitState = state
            mConverged = True
            Exit For
        End If

        If Not ApplyUltimateNewtonStep(section, concreteMaterial, steelMaterial, _
                nSpecified, mxBase, myBase, mxOffset, myOffset, _
                eps0, kx, ky, rN, rDir, rLimit) Then Exit Sub
    Next iter

    If Not mConverged Then
        If Len(mLimitState) = 0 Then mLimitState = STATE_NUMERICAL_FAILURE
        If Len(mStopReason) = 0 Then mStopReason = "UltimateStrain: итерации Newton не сошлись."
        Exit Sub
    End If

    If Not mLastSolver Is Nothing Then
        mMomentEquilibriumResidual = Sqr((mLastSolver.Mxint - TargetMx(mLambdaUltimate, mxBase, mxOffset)) ^ 2 + _
            (mLastSolver.Myint - TargetMy(mLambdaUltimate, myBase, myOffset)) ^ 2)
    End If
    mSafetyFactor = mLambdaUltimate
    If mLambdaUltimate > 0# Then mUtilization = 1# / mLambdaUltimate
    mStopReason = "Поиск несущей способности через UltimateStrain сошелся."
End Sub

' Математически чистая lambda*N без моментов не имеет устойчивой уникальной
' предельной плоскости для UltimateStrain: на горизонтальном плато одну и ту же
' N могут давать разные почти равномерные деформации. В этом вырожденном
' численном случае используем одномерный LoadMultiplier.
Private Sub SolvePureAxialUltimateByLoadMultiplier(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal nOffset As Double, ByVal nBase As Double, _
        ByVal mxOffset As Double, ByVal mxBase As Double, _
        ByVal myOffset As Double, ByVal myBase As Double)
    SolveByLoadPathMultiplier section, concreteMaterial, steelMaterial, _
        nOffset, nBase, mxOffset, mxBase, myOffset, myBase, True
    If Not mConverged Then AcceptPureAxialRetainedCapacity section, concreteMaterial

    If mDiagnosticsEnabled Then
        mDiagnosticLog = "Силовая траектория UltimateStrain перенаправлена на LoadMultiplier." & vbCrLf & mDiagnosticLog
    End If
    If mConverged Then
        mStopReason = "Чистая осевая траектория рассчитана через LoadMultiplier как устойчивую одномерную постановку. " & mStopReason
    End If
End Sub

' Автоматический режим сначала пробует UltimateStrain как более прямой способ
' найти предельную плоскость деформаций. Если эта постановка не сошлась
' численно, но входные данные корректны, расчет повторяется через
' LoadMultiplier по той же самой lambda-траектории. Это не меняет физический
' критерий несущей способности, а только выбирает более устойчивый численный
' путь для случаев с горизонтальными участками диаграмм.
Public Sub SolveByAutoLoadPath(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal nOffset As Double, ByVal nBase As Double, _
        ByVal mxOffset As Double, ByVal mxBase As Double, _
        ByVal myOffset As Double, ByVal myBase As Double, _
        Optional ByVal allowPureAxialRedirect As Boolean = True, _
        Optional ByVal allowForcePathRetainedLimit As Boolean = False)
    SolveByUltimateLoadPath section, concreteMaterial, steelMaterial, _
        nOffset, nBase, mxOffset, mxBase, myOffset, myBase, allowPureAxialRedirect
    If mConverged Or mLimitState = STATE_INVALID_INPUT Then Exit Sub

    Dim firstStopReason As String
    Dim firstLimitState As String
    Dim firstDiagnostics As String
    firstStopReason = mStopReason
    firstLimitState = mLimitState
    firstDiagnostics = mDiagnosticLog

    SolveByLoadPathMultiplier section, concreteMaterial, steelMaterial, _
        nOffset, nBase, mxOffset, mxBase, myOffset, myBase, allowForcePathRetainedLimit

    If mDiagnosticsEnabled Then
        mDiagnosticLog = "Auto: первая попытка UltimateStrain не дала принятую предельную точку; статус=" & _
            firstLimitState & "; причина=" & firstStopReason & vbCrLf & _
            firstDiagnostics & _
            "Auto: переходим к LoadMultiplier по той же lambda-траектории." & vbCrLf & _
            mDiagnosticLog
    End If
    If mConverged Then
        mStopReason = "Auto: UltimateStrain не сошелся, предельная точка найдена через LoadMultiplier. " & mStopReason
    End If
End Sub

' Чистое осевое масштабирование без постоянных моментов является
' вырожденным для UltimateStrain при горизонтальном плато диаграммы:
' разные равномерные деформации дают одну и ту же предельную N. Если прямой
' UltimateStrain не сошелся устойчиво, эта проверка разрешает перейти на ту
' же универсальную lambda-траекторию и найти границу одномерно.
Private Function IsPureAxialLoadPath(ByVal nOffset As Double, ByVal nBase As Double, _
        ByVal mxOffset As Double, ByVal mxBase As Double, _
        ByVal myOffset As Double, ByVal myBase As Double) As Boolean
    IsPureAxialLoadPath = Abs(nBase) > 0.000000001 And _
        Abs(nOffset) <= 0.000000001 And _
        Abs(mxOffset) <= 0.000000001 And Abs(mxBase) <= 0.000000001 And _
        Abs(myOffset) <= 0.000000001 And Abs(myBase) <= 0.000000001
End Function

' True для классической траектории "N = const, M = lambda*Mbase".
' Это lambda*Mx, lambda*My и lambda*Mxy. Для них внутри универсального API используется
' быстрая система UltimateStrain, потому что она лучше обусловлена
' именно для моментного поиска несущей способности.
Private Function IsMomentOnlyLoadPath(ByVal nBase As Double, ByVal mxBase As Double, ByVal myBase As Double) As Boolean
    IsMomentOnlyLoadPath = Abs(nBase) <= 0.000000001 And _
        Sqr(mxBase * mxBase + myBase * myBase) > 0.000000001
End Function

' При чистой lambda*N траектории горизонтальное плато диаграммы может дать
' устойчивую нижнюю границу capacity, а следующая probe-точка уже теряет
' касательную жесткость до формального пересечения eps_ult. Для такого
' осевого случая это означает достижение силового плато, поэтому принимаем
' последнюю сходящуюся lambda как найденную несущую способность.
Private Sub AcceptPureAxialRetainedCapacity(ByVal section As CSectionModel, ByVal concreteMaterial As Object)
    If mLambdaUltimate <= 0# Then Exit Sub
    If mLastAcceptableSolver Is Nothing Then Exit Sub

    Set mLastSolver = mLastAcceptableSolver
    UpdateCriticalState section, concreteMaterial, mLastAcceptableSolver
    mLimitState = CriticalLimitState()
    If mLimitState = STATE_NUMERICAL_FAILURE Then
        If mLastAcceptableN < 0# Then
            mLimitState = STATE_CONCRETE_LIMIT
        Else
            mLimitState = STATE_STEEL_LIMIT
        End If
    End If
    mUpperLimitState = mLimitState
    mSafetyFactor = mLambdaUltimate
    mUtilization = 1# / mLambdaUltimate
    mConverged = True
    AppendSearchDiagnostic "для чисто осевой траектории принята последняя сходящаяся lambda как силовое плато; lambda=" & _
        FormatNumberInvariant(mLambdaUltimate) & "; состояние=" & mLimitState
End Sub

' Стартовая плоскость для универсального UltimateStrain подбирается по
' направлению полной нагрузки при lambda = 1. Это только начальное приближение:
' дальше Newton одновременно уточняет равновесие вдоль выбранной траектории
' и доводит критическое волокно/стержень до физической предельной деформации.
Private Sub InitializeUltimateLoadPathGuess(ByVal section As CSectionModel, _
        ByVal nOffset As Double, ByVal nBase As Double, _
        ByVal mxOffset As Double, ByVal mxBase As Double, _
        ByVal myOffset As Double, ByVal myBase As Double, _
        ByRef eps0 As Double, ByRef kx As Double, ByRef ky As Double)
    Dim area As Double
    Dim maxAbsX As Double
    Dim maxAbsY As Double
    Dim i As Long
    Dim targetNValue As Double
    Dim targetMxValue As Double
    Dim targetMyValue As Double
    Dim momentNorm As Double
    Dim strainLimit As Double

    For i = 1 To section.ConcreteCount
        area = area + section.ConcreteArea(i)
        If Abs(section.ConcreteX(i)) > maxAbsX Then maxAbsX = Abs(section.ConcreteX(i))
        If Abs(section.ConcreteY(i)) > maxAbsY Then maxAbsY = Abs(section.ConcreteY(i))
    Next i
    If section.RebarCount > 0 Then
        For i = 1 To section.RebarCount
            area = area + section.RebarArea(i)
            If Abs(section.RebarX(i)) > maxAbsX Then maxAbsX = Abs(section.RebarX(i))
            If Abs(section.RebarY(i)) > maxAbsY Then maxAbsY = Abs(section.RebarY(i))
        Next i
    End If

    targetNValue = TargetN(1#, nOffset, nBase)
    targetMxValue = TargetMx(1#, mxBase, mxOffset)
    targetMyValue = TargetMy(1#, myBase, myOffset)
    If area > 0# Then eps0 = targetNValue / (area * 30000#)

    strainLimit = Abs(mConcreteCompressionLimit)
    If mSteelStrainLimit > 0# Then strainLimit = MinDouble(strainLimit, mSteelStrainLimit)
    If strainLimit <= 0# Then strainLimit = 0.001

    If targetNValue < 0# Then
        eps0 = MinDouble(eps0, 0.25 * mConcreteCompressionLimit)
    ElseIf targetNValue > 0# Then
        eps0 = MaxDouble(eps0, 0.25 * mSteelStrainLimit)
    End If

    momentNorm = Sqr(targetMxValue * targetMxValue + targetMyValue * targetMyValue)
    If momentNorm <= 0.000000001 Then Exit Sub

    If maxAbsY > 0# Then kx = Sgn(targetMxValue) * strainLimit / maxAbsY * Abs(targetMxValue) / momentNorm
    If maxAbsX > 0# Then ky = Sgn(targetMyValue) * strainLimit / maxAbsX * Abs(targetMyValue) / momentNorm
End Sub

' Оценивает плоскость деформаций для универсального UltimateStrain.
' Первые две невязки удерживают внутренние усилия на выбранной прямой
' load-path в пространстве N/Mx/My, третья отвечает за достижение предельной
' деформации материала.
Private Function EvaluateUltimateLoadPathResidual(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal nOffset As Double, ByVal nBase As Double, _
        ByVal mxOffset As Double, ByVal mxBase As Double, _
        ByVal myOffset As Double, ByVal myBase As Double, _
        ByVal eps0 As Double, ByVal kx As Double, ByVal ky As Double, _
        ByRef solver As CSectionSolver, ByRef rPath1 As Double, ByRef rPath2 As Double, _
        ByRef rLimit As Double, ByRef lambdaValue As Double, ByRef state As String) As Boolean
    On Error GoTo Failed

    Set solver = New CSectionSolver
    solver.EvaluateStrainPlane section, concreteMaterial, steelMaterial, eps0, kx, ky
    mIterations = mIterations + 1
    Set mLastSolver = solver
    UpdateCriticalState section, concreteMaterial, solver

    lambdaValue = LoadPathLambda(solver.Nint, solver.Mxint, solver.Myint, _
        nOffset, nBase, mxOffset, mxBase, myOffset, myBase)
    BuildLoadPathResiduals solver.Nint, solver.Mxint, solver.Myint, _
        nOffset, nBase, mxOffset, mxBase, myOffset, myBase, lambdaValue, rPath1, rPath2
    rLimit = mCriticalStrainUtilization - 1#
    state = CriticalLimitState()

    AppendSearchDiagnostic "ultimatePath iter=" & CStr(mIterations) & _
        "; eps0=" & FormatNumberInvariant(eps0) & _
        "; kappaX=" & FormatNumberInvariant(kx) & _
        "; kappaY=" & FormatNumberInvariant(ky) & _
        "; rPath1=" & FormatNumberInvariant(rPath1) & _
        "; rPath2=" & FormatNumberInvariant(rPath2) & _
        "; rLimit=" & FormatNumberInvariant(rLimit) & _
        "; lambda=" & FormatNumberInvariant(lambdaValue) & _
        "; critical=" & mCriticalElement
    EvaluateUltimateLoadPathResidual = True
    Exit Function

Failed:
    mLimitState = STATE_NUMERICAL_FAILURE
    mStopReason = "UltimateStrain: не удалось оценить состояние: " & Err.Description
    EvaluateUltimateLoadPathResidual = False
End Function

' Делает Newton-шаг для универсальной постановки UltimateStrain.
' Якобиан строится численно, потому что material diagram уже инкапсулирует
' кусочно-линейную диаграмму, а ручное дифференцирование дало бы дублирование.
Private Function ApplyUltimateLoadPathNewtonStep(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal nOffset As Double, ByVal nBase As Double, _
        ByVal mxOffset As Double, ByVal mxBase As Double, _
        ByVal myOffset As Double, ByVal myBase As Double, _
        ByRef eps0 As Double, ByRef kx As Double, ByRef ky As Double, _
        ByVal rPath1 As Double, ByVal rPath2 As Double, ByVal rLimit As Double) As Boolean
    On Error GoTo Failed

    Dim hEps As Double
    Dim hKx As Double
    Dim hKy As Double
    Dim j11 As Double, j21 As Double, j31 As Double
    Dim j12 As Double, j22 As Double, j32 As Double
    Dim j13 As Double, j23 As Double, j33 As Double
    Dim system As CLinearSystem3x3
    Dim dEps As Double
    Dim dKx As Double
    Dim dKy As Double
    Dim baseNorm As Double
    Dim alpha As Double
    Dim trialSolver As CSectionSolver
    Dim trialR1 As Double, trialR2 As Double, trialRLimit As Double
    Dim trialLambda As Double, trialState As String

    hEps = MaxDouble(0.0000001, Abs(eps0) * 0.0001)
    hKx = MaxDouble(0.000000001, Abs(kx) * 0.0001)
    hKy = MaxDouble(0.000000001, Abs(ky) * 0.0001)

    BuildUltimateLoadPathJacobianColumn section, concreteMaterial, steelMaterial, _
        nOffset, nBase, mxOffset, mxBase, myOffset, myBase, _
        eps0 + hEps, kx, ky, hEps, rPath1, rPath2, rLimit, j11, j21, j31
    BuildUltimateLoadPathJacobianColumn section, concreteMaterial, steelMaterial, _
        nOffset, nBase, mxOffset, mxBase, myOffset, myBase, _
        eps0, kx + hKx, ky, hKx, rPath1, rPath2, rLimit, j12, j22, j32
    BuildUltimateLoadPathJacobianColumn section, concreteMaterial, steelMaterial, _
        nOffset, nBase, mxOffset, mxBase, myOffset, myBase, _
        eps0, kx, ky + hKy, hKy, rPath1, rPath2, rLimit, j13, j23, j33

    Set system = New CLinearSystem3x3
    If Not system.Solve(j11, j12, j13, j21, j22, j23, j31, j32, j33, -rPath1, -rPath2, -rLimit) Then
        mLimitState = STATE_SINGULAR_TANGENT
        mStopReason = "UltimateStrain: не удалось решить касательную систему: " & system.StopReason
        Exit Function
    End If

    dEps = system.X1
    dKx = system.X2
    dKy = system.X3
    LimitUltimateIncrement dEps, dKx, dKy

    baseNorm = UltimateLoadPathResidualNorm(rPath1, rPath2, rLimit)
    alpha = mSolverDampingInitial
    If alpha <= 0# Or alpha > 1# Then alpha = 1#

    Do While alpha >= mSolverMinLineSearchAlpha
        If EvaluateUltimateLoadPathResidual(section, concreteMaterial, steelMaterial, _
                nOffset, nBase, mxOffset, mxBase, myOffset, myBase, _
                eps0 + alpha * dEps, kx + alpha * dKx, ky + alpha * dKy, _
                trialSolver, trialR1, trialR2, trialRLimit, trialLambda, trialState) Then
            If (Not mSolverLineSearchEnabled) Or UltimateLoadPathResidualNorm(trialR1, trialR2, trialRLimit) < baseNorm Then
                eps0 = eps0 + alpha * dEps
                kx = kx + alpha * dKx
                ky = ky + alpha * dKy
                ApplyUltimateLoadPathNewtonStep = True
                Exit Function
            End If
        End If
        alpha = alpha / 2#
    Loop

    mLimitState = STATE_NUMERICAL_FAILURE
    mStopReason = "UltimateStrain: line search не дал улучшения."
    Exit Function

Failed:
    mLimitState = STATE_NUMERICAL_FAILURE
    mStopReason = "UltimateStrain: шаг Newton не выполнен: " & Err.Description
End Function

' Строит один столбец численного Якобиана универсальной предельной задачи.
Private Sub BuildUltimateLoadPathJacobianColumn(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal nOffset As Double, ByVal nBase As Double, _
        ByVal mxOffset As Double, ByVal mxBase As Double, _
        ByVal myOffset As Double, ByVal myBase As Double, _
        ByVal eps0 As Double, ByVal kx As Double, ByVal ky As Double, ByVal stepValue As Double, _
        ByVal rPath1 As Double, ByVal rPath2 As Double, ByVal rLimit As Double, _
        ByRef j1 As Double, ByRef j2 As Double, ByRef j3 As Double)
    Dim solver As CSectionSolver
    Dim pR1 As Double, pR2 As Double, pLimit As Double
    Dim pLambda As Double, pState As String

    If Not EvaluateUltimateLoadPathResidual(section, concreteMaterial, steelMaterial, _
            nOffset, nBase, mxOffset, mxBase, myOffset, myBase, _
            eps0, kx, ky, solver, pR1, pR2, pLimit, pLambda, pState) Then Exit Sub
    j1 = (pR1 - rPath1) / stepValue
    j2 = (pR2 - rPath2) / stepValue
    j3 = (pLimit - rLimit) / stepValue
End Sub

' Сводит три невязки UltimateStrain в один скаляр для line-search и контроля
' улучшения решения: две невязки удерживают load-path, третья отвечает за
' достижение предельной деформации.
Private Function UltimateLoadPathResidualNorm(ByVal rPath1 As Double, ByVal rPath2 As Double, ByVal rLimit As Double) As Double
    Dim math As CLoadPathMath
    Set math = NewLoadPathMath()
    UltimateLoadPathResidualNorm = math.ResidualNorm(rPath1, rPath2, rLimit)
End Function

' Стартовая плоскость нужна только как начальное приближение UltimateStrain.
' Она не задает физический результат: дальнейшая итерация заново ищет N,
' направление моментов и достижение предельной деформации.
Private Sub InitializeUltimateStrainGuess(ByVal section As CSectionModel, _
        ByVal nSpecified As Double, ByVal mxBase As Double, ByVal myBase As Double, _
        ByRef eps0 As Double, ByRef kx As Double, ByRef ky As Double)
    Dim area As Double
    Dim maxAbsX As Double
    Dim maxAbsY As Double
    Dim i As Long

    For i = 1 To section.ConcreteCount
        area = area + section.ConcreteArea(i)
        If Abs(section.ConcreteX(i)) > maxAbsX Then maxAbsX = Abs(section.ConcreteX(i))
        If Abs(section.ConcreteY(i)) > maxAbsY Then maxAbsY = Abs(section.ConcreteY(i))
    Next i
    If section.RebarCount > 0 Then
        For i = 1 To section.RebarCount
            area = area + section.RebarArea(i)
            If Abs(section.RebarX(i)) > maxAbsX Then maxAbsX = Abs(section.RebarX(i))
            If Abs(section.RebarY(i)) > maxAbsY Then maxAbsY = Abs(section.RebarY(i))
        Next i
    End If

    If area > 0# Then eps0 = nSpecified / (area * 30000#)
    Dim baseNorm As Double
    baseNorm = Sqr(mxBase * mxBase + myBase * myBase)
    Dim strainLimit As Double
    strainLimit = Abs(mConcreteCompressionLimit)
    If mSteelStrainLimit > 0# Then strainLimit = MinDouble(strainLimit, mSteelStrainLimit)
    If strainLimit <= 0# Then strainLimit = 0.001

    If maxAbsY > 0# Then kx = Sgn(mxBase) * strainLimit / maxAbsY * Abs(mxBase) / baseNorm
    If maxAbsX > 0# Then ky = Sgn(myBase) * strainLimit / maxAbsX * Abs(myBase) / baseNorm
    If Abs(kx) <= 0.000000000001 And Abs(ky) <= 0.000000000001 Then kx = strainLimit / MaxDouble(maxAbsY, 1#)
End Sub

' Считает три безразмерные невязки UltimateStrain:
' rN - равновесие по продольной силе,
' rDir - сохранение направления вектора Mx/My,
' rLimit - достижение критической деформации.
Private Function EvaluateUltimateResidual(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal nSpecified As Double, ByVal mxBase As Double, ByVal myBase As Double, _
        ByVal mxOffset As Double, ByVal myOffset As Double, _
        ByVal eps0 As Double, ByVal kx As Double, ByVal ky As Double, _
        ByRef solver As CSectionSolver, ByRef rN As Double, ByRef rDir As Double, _
        ByRef rLimit As Double, ByRef lambdaValue As Double, ByRef state As String) As Boolean
    On Error GoTo Failed

    ' Здесь нет итерационного решения равновесия: мы оцениваем уже заданную
    ' плоскость деформаций и смотрим, какие внутренние усилия она дает.
    Set solver = New CSectionSolver
    solver.EvaluateStrainPlane section, concreteMaterial, steelMaterial, eps0, kx, ky
    mIterations = mIterations + 1
    Set mLastSolver = solver
    UpdateCriticalState section, concreteMaterial, solver

    Dim baseNorm As Double
    Dim baseNorm2 As Double
    baseNorm2 = mxBase * mxBase + myBase * myBase
    baseNorm = Sqr(baseNorm2)
    Dim userMxInt As Double
    Dim userMyInt As Double
    ' Моменты сравниваются в пользовательской точке приложения нагрузки:
    ' внутренние моменты сначала возвращаются из системы координат элементов.
    userMxInt = solver.Mxint - mxOffset
    userMyInt = solver.Myint - myOffset
    lambdaValue = (userMxInt * mxBase + userMyInt * myBase) / baseNorm2

    rN = (solver.Nint - nSpecified) / MaxDouble(mSolverToleranceN, Abs(nSpecified) * 0.000000001)
    Dim perpendicularMoment As Double
    perpendicularMoment = (userMxInt * myBase - userMyInt * mxBase) / baseNorm
    rDir = perpendicularMoment / MaxDouble(MaxDouble(Abs(lambdaValue) * baseNorm, baseNorm), 1#)
    rLimit = mCriticalStrainUtilization - 1#
    state = CriticalLimitState()

    AppendSearchDiagnostic "ultimate iter=" & CStr(mIterations) & _
        "; eps0=" & FormatNumberInvariant(eps0) & _
        "; kappaX=" & FormatNumberInvariant(kx) & _
        "; kappaY=" & FormatNumberInvariant(ky) & _
        "; rN=" & FormatNumberInvariant(rN) & _
        "; rDir=" & FormatNumberInvariant(rDir) & _
        "; rLimit=" & FormatNumberInvariant(rLimit) & _
        "; lambda=" & FormatNumberInvariant(lambdaValue) & _
        "; critical=" & mCriticalElement
    EvaluateUltimateResidual = True
    Exit Function

Failed:
    mLimitState = STATE_NUMERICAL_FAILURE
    mStopReason = "UltimateStrain: не удалось оценить состояние: " & Err.Description
    EvaluateUltimateResidual = False
End Function

' Делает один Newton-шаг для системы UltimateStrain. Якобиан строится
' численным дифференцированием трех невязок по epsilon0, kappaX и kappaY.
Private Function ApplyUltimateNewtonStep(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal nSpecified As Double, ByVal mxBase As Double, ByVal myBase As Double, _
        ByVal mxOffset As Double, ByVal myOffset As Double, _
        ByRef eps0 As Double, ByRef kx As Double, ByRef ky As Double, _
        ByVal rN As Double, ByVal rDir As Double, ByVal rLimit As Double) As Boolean
    On Error GoTo Failed

    ' Малые приращения выбираются относительно текущего масштаба неизвестных,
    ' чтобы получить численно устойчивые производные.
    Dim hEps As Double
    Dim hKx As Double
    Dim hKy As Double
    hEps = MaxDouble(0.0000001, Abs(eps0) * 0.0001)
    hKx = MaxDouble(0.000000001, Abs(kx) * 0.0001)
    hKy = MaxDouble(0.000000001, Abs(ky) * 0.0001)

    Dim j11 As Double, j21 As Double, j31 As Double
    Dim j12 As Double, j22 As Double, j32 As Double
    Dim j13 As Double, j23 As Double, j33 As Double
    BuildUltimateJacobianColumn section, concreteMaterial, steelMaterial, nSpecified, mxBase, myBase, mxOffset, myOffset, _
        eps0 + hEps, kx, ky, hEps, rN, rDir, rLimit, j11, j21, j31
    BuildUltimateJacobianColumn section, concreteMaterial, steelMaterial, nSpecified, mxBase, myBase, mxOffset, myOffset, _
        eps0, kx + hKx, ky, hKx, rN, rDir, rLimit, j12, j22, j32
    BuildUltimateJacobianColumn section, concreteMaterial, steelMaterial, nSpecified, mxBase, myBase, mxOffset, myOffset, _
        eps0, kx, ky + hKy, hKy, rN, rDir, rLimit, j13, j23, j33

    ' Решаем поправку сразу для всех трех условий UltimateStrain.
    Dim system As CLinearSystem3x3
    Set system = New CLinearSystem3x3
    If Not system.Solve(j11, j12, j13, j21, j22, j23, j31, j32, j33, -rN, -rDir, -rLimit) Then
        mLimitState = STATE_SINGULAR_TANGENT
        mStopReason = "UltimateStrain: не удалось решить касательную систему: " & system.StopReason
        Exit Function
    End If

    Dim dEps As Double
    Dim dKx As Double
    Dim dKy As Double
    dEps = system.X1
    dKx = system.X2
    dKy = system.X3
    LimitUltimateIncrement dEps, dKx, dKy

    ' Полный шаг может ухудшить невязку, поэтому принимаем его через простой
    ' line search с последовательным делением alpha пополам.
    Dim baseNorm As Double
    baseNorm = UltimateResidualNorm(rN, rDir, rLimit)
    Dim alpha As Double
    alpha = mSolverDampingInitial
    If alpha <= 0# Or alpha > 1# Then alpha = 1#

    Do While alpha >= mSolverMinLineSearchAlpha
        Dim trialSolver As CSectionSolver
        Dim trialRN As Double, trialRDir As Double, trialRLimit As Double
        Dim trialLambda As Double, trialState As String
        If EvaluateUltimateResidual(section, concreteMaterial, steelMaterial, nSpecified, mxBase, myBase, mxOffset, myOffset, _
                eps0 + alpha * dEps, kx + alpha * dKx, ky + alpha * dKy, trialSolver, _
                trialRN, trialRDir, trialRLimit, trialLambda, trialState) Then
            If (Not mSolverLineSearchEnabled) Or UltimateResidualNorm(trialRN, trialRDir, trialRLimit) < baseNorm Then
                eps0 = eps0 + alpha * dEps
                kx = kx + alpha * dKx
                ky = ky + alpha * dKy
                ApplyUltimateNewtonStep = True
                Exit Function
            End If
        End If
        alpha = alpha / 2#
    Loop

    mLimitState = STATE_NUMERICAL_FAILURE
    mStopReason = "UltimateStrain: line search не дал улучшения."
    Exit Function

Failed:
    mLimitState = STATE_NUMERICAL_FAILURE
    mStopReason = "UltimateStrain: шаг Newton не выполнен: " & Err.Description
End Function

' Строит один столбец численного Якобиана UltimateStrain: меняем одну
' неизвестную на stepValue и смотрим изменение всех трех невязок.
Private Sub BuildUltimateJacobianColumn(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal nSpecified As Double, ByVal mxBase As Double, ByVal myBase As Double, _
        ByVal mxOffset As Double, ByVal myOffset As Double, _
        ByVal eps0 As Double, ByVal kx As Double, ByVal ky As Double, ByVal stepValue As Double, _
        ByVal rN As Double, ByVal rDir As Double, ByVal rLimit As Double, _
        ByRef j1 As Double, ByRef j2 As Double, ByRef j3 As Double)
    Dim solver As CSectionSolver
    Dim pN As Double, pDir As Double, pLimit As Double
    Dim pLambda As Double, pState As String
    If Not EvaluateUltimateResidual(section, concreteMaterial, steelMaterial, nSpecified, mxBase, myBase, mxOffset, myOffset, _
            eps0, kx, ky, solver, pN, pDir, pLimit, pLambda, pState) Then Exit Sub
    j1 = (pN - rN) / stepValue
    j2 = (pDir - rDir) / stepValue
    j3 = (pLimit - rLimit) / stepValue
End Sub

' Норма смешивает три разных по смыслу остатка. Весовые множители не имеют
' нормативного смысла; они только помогают line search сравнивать пробные шаги.
Private Function UltimateResidualNorm(ByVal rN As Double, ByVal rDir As Double, ByVal rLimit As Double) As Double
    UltimateResidualNorm = Abs(rN) + Abs(rDir) * 100# + Abs(rLimit) * 10#
End Function

' Переводит абсолютный допуск деформации в относительный допуск по условию
' rLimit = utilization - 1, потому что критический предел может быть разным.
Private Function UltimateStrainRatioTolerance() As Double
    Dim limitAbs As Double
    limitAbs = Abs(mCriticalStrain)
    If InStr(1, mCriticalElement, ":Compression", vbTextCompare) > 0 Then
        limitAbs = Abs(mConcreteCompressionLimit)
    ElseIf InStr(1, mCriticalElement, ":Tension", vbTextCompare) > 0 Then
        limitAbs = Abs(mConcreteTensionLimit)
    ElseIf InStr(1, mCriticalElement, ":Steel", vbTextCompare) > 0 Then
        limitAbs = Abs(mSteelStrainLimit)
    End If
    UltimateStrainRatioTolerance = MaxDouble(0.000001, mStrainTolerance / MaxDouble(limitAbs, 0.000000001))
End Function

' Ограничивает приращения UltimateStrain теми же защитными пределами, что и
' внутренний решатель равновесия.
Private Sub LimitUltimateIncrement(ByRef dEps As Double, ByRef dKx As Double, ByRef dKy As Double)
    Dim scaleFactor As Double
    scaleFactor = 1#
    If mSolverMaxDeltaEpsilon0 > 0# And Abs(dEps) > mSolverMaxDeltaEpsilon0 Then
        scaleFactor = MinDouble(scaleFactor, mSolverMaxDeltaEpsilon0 / Abs(dEps))
    End If
    If mSolverMaxDeltaKappa > 0# And Abs(dKx) > mSolverMaxDeltaKappa Then
        scaleFactor = MinDouble(scaleFactor, mSolverMaxDeltaKappa / Abs(dKx))
    End If
    If mSolverMaxDeltaKappa > 0# And Abs(dKy) > mSolverMaxDeltaKappa Then
        scaleFactor = MinDouble(scaleFactor, mSolverMaxDeltaKappa / Abs(dKy))
    End If
    dEps = dEps * scaleFactor
    dKx = dKx * scaleFactor
    dKy = dKy * scaleFactor
End Sub

' ------------------------------
' Метод LoadMultiplier
' ------------------------------
' Сохраняет N постоянным, а Mx/My масштабирует одним lambda. Каждая probe-точка
' вызывает CSectionSolver; внешний алгоритм только ищет границу между
' acceptable-состоянием и физическим пределом.
Public Sub SolveByLoadMultiplier(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal nSpecified As Double, ByVal mxBase As Double, ByVal myBase As Double, _
        Optional ByVal mxOffset As Double = 0#, Optional ByVal myOffset As Double = 0#)
    SolveByLoadPathMultiplier section, concreteMaterial, steelMaterial, _
        nSpecified, 0#, mxOffset, mxBase, myOffset, myBase
End Sub

' Ищет предельный lambda для любой траектории Target = Offset + lambda*Base.
' Метод используется batch-слоем для новых вариантов lambda*Mx, lambda*My,
' lambda*Mxy, lambda*N и lambda*NMxy. Внутренний алгоритм скобки и Bisection/Brent/Secant остается
' тем же для всех вариантов CapacityLoadPath.
Public Sub SolveByLoadPathMultiplier(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal nOffset As Double, ByVal nBase As Double, _
        ByVal mxOffset As Double, ByVal mxBase As Double, _
        ByVal myOffset As Double, ByVal myBase As Double, _
        Optional ByVal allowForcePathRetainedLimit As Boolean = False)
    ClearResult
    mSolutionMethod = "LoadMultiplier"

    If Not ValidateLoadPathCapacityInputs(section, concreteMaterial, steelMaterial, nBase, mxBase, myBase) Then Exit Sub
    If Not ValidateLoadMultiplierSettings() Then Exit Sub

    SolveLoadMultiplierCore section, concreteMaterial, steelMaterial, nOffset, nBase, mxOffset, mxBase, myOffset, myBase, _
        allowForcePathRetainedLimit
End Sub

' Общий одномерный поиск предельного множителя. Входные Offset/Base задают
' любую допустимую lambda-траекторию нагрузки, а алгоритм скобки, probe-точек и
' Bisection/Brent/Secant остается одним для всех вариантов.
Private Sub SolveLoadMultiplierCore(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal nOffset As Double, ByVal nBase As Double, _
        ByVal mxOffset As Double, ByVal mxBase As Double, _
        ByVal myOffset As Double, ByVal myBase As Double, _
        ByVal allowForcePathRetainedLimit As Boolean)
    Dim lowLambda As Double
    Dim highLambda As Double
    Dim probeState As String
    lowLambda = 0#
    highLambda = mInitialLambdaStep
    If highLambda <= 0# Then highLambda = 1#

    ' Сначала проверяем точку lambda = 0. Если даже состояние с одним N уже
    ' неприемлемо, искать моментную несущую способность бессмысленно.
    probeState = ProbeWithRetries(section, concreteMaterial, steelMaterial, TargetN(lowLambda, nOffset, nBase), _
        TargetMx(lowLambda, mxBase, mxOffset), TargetMy(lowLambda, myBase, myOffset), lowLambda)
    If probeState <> STATE_ACCEPTABLE Then
        mLimitState = probeState
        mStopReason = "Начальное состояние при lambda = 0 не проходит физический критерий: " & probeState & "."
        Exit Sub
    End If
    RecordAcceptableState mLastSolver, TargetN(lowLambda, nOffset, nBase), TargetMx(lowLambda, mxBase, mxOffset), TargetMy(lowLambda, myBase, myOffset)

    ' Расширяем верхнюю границу, пока не получим физический предел. Между
    ' lowLambda и highLambda далее должен лежать корень целевой функции.
    Dim bracketFound As Boolean
    Do While highLambda <= mMaxLambda
        probeState = ProbeWithRetries(section, concreteMaterial, steelMaterial, TargetN(highLambda, nOffset, nBase), _
            TargetMx(highLambda, mxBase, mxOffset), TargetMy(highLambda, myBase, myOffset), highLambda)
        If probeState = STATE_ACCEPTABLE Then
            lowLambda = highLambda
            RecordAcceptableState mLastSolver, TargetN(highLambda, nOffset, nBase), TargetMx(highLambda, mxBase, mxOffset), TargetMy(highLambda, myBase, myOffset)
            highLambda = highLambda * 2#
        ElseIf IsPhysicalLimit(probeState) Then
            mUpperLimitState = probeState
            bracketFound = True
            Exit Do
        ElseIf IsNumericalFailure(probeState) Then
            highLambda = ReduceLambdaAfterNumericalFailure(lowLambda, highLambda, probeState)
            If highLambda <= lowLambda Then
                If allowForcePathRetainedLimit Then AcceptPureAxialRetainedCapacity section, concreteMaterial
                Exit Sub
            End If
        Else
            mLimitState = probeState
            mStopReason = "Поиск несущей способности остановлен: " & probeState & "."
            Exit Sub
        End If
    Loop

    If Not bracketFound Then
        If Len(mLimitState) = 0 Then mLimitState = STATE_NUMERICAL_FAILURE
        If Len(mStopReason) = 0 Then mStopReason = "До Capacity.MaxLambda не найдена верхняя физическая граница lambda."
        mLambdaUltimate = lowLambda
        If allowForcePathRetainedLimit Then AcceptPureAxialRetainedCapacity section, concreteMaterial
        Exit Sub
    End If

    ' Когда физическая скобка найдена, выбираем один из самостоятельных
    ' одномерных методов поиска lambda.
    Select Case NormalizedSearchMethod()
        Case "Bisection"
            SearchByBisection section, concreteMaterial, steelMaterial, nOffset, nBase, mxBase, myBase, mxOffset, myOffset, lowLambda, highLambda
        Case "Brent"
            SearchByBrent section, concreteMaterial, steelMaterial, nOffset, nBase, mxBase, myBase, mxOffset, myOffset, lowLambda, highLambda
        Case "Secant"
            SearchBySecant section, concreteMaterial, steelMaterial, nOffset, nBase, mxBase, myBase, mxOffset, myOffset, lowLambda, highLambda
        Case Else
            SetInvalidInput "Capacity.SearchMethod должен быть Bisection, Brent или Secant."
    End Select
    If Not mConverged And allowForcePathRetainedLimit Then AcceptPureAxialRetainedCapacity section, concreteMaterial
End Sub

' Устойчивый базовый метод: каждый раз делит физическую скобку пополам.
' Быстрее он не пытается быть, зато граница acceptable/limit не теряется.
Private Sub SearchByBisection(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal nOffset As Double, ByVal nBase As Double, _
        ByVal mxBase As Double, ByVal myBase As Double, _
        ByVal mxOffset As Double, ByVal myOffset As Double, _
        ByVal lowLambda As Double, ByVal highLambda As Double)
    AppendSearchDiagnostic "searchMethod=Bisection"
    Dim midLambda As Double
    Dim probeState As String
    Do While (highLambda - lowLambda) > mLambdaTolerance
        midLambda = (lowLambda + highLambda) / 2#
        probeState = ProbeWithRetries(section, concreteMaterial, steelMaterial, TargetN(midLambda, nOffset, nBase), _
            TargetMx(midLambda, mxBase, mxOffset), TargetMy(midLambda, myBase, myOffset), midLambda)
        If probeState = STATE_ACCEPTABLE Then
            lowLambda = midLambda
            RecordAcceptableState mLastSolver, TargetN(midLambda, nOffset, nBase), TargetMx(midLambda, mxBase, mxOffset), TargetMy(midLambda, myBase, myOffset)
        ElseIf IsPhysicalLimit(probeState) Then
            highLambda = midLambda
            mUpperLimitState = probeState
        ElseIf IsNumericalFailure(probeState) Then
            mLimitState = probeState
            mStopReason = "Внутри физической скобки возникла численная несходимость; сохранена последняя сходящаяся lambda."
            mLambdaUltimate = lowLambda
            If Not mLastAcceptableSolver Is Nothing Then Set mLastSolver = mLastAcceptableSolver
            Exit Sub
        Else
            mLimitState = probeState
            mStopReason = "Поиск Bisection остановлен: " & probeState & "."
            Exit Sub
        End If
    Loop

    StoreLoadPathUltimateComponents lowLambda, nOffset, nBase, mxBase, myBase, mxOffset, myOffset
    If Not mLastAcceptableSolver Is Nothing Then
        mMomentEquilibriumResidual = Sqr((mLastAcceptableSolver.Mxint - TargetMx(mLambdaUltimate, mxBase, mxOffset)) ^ 2 + _
            (mLastAcceptableSolver.Myint - TargetMy(mLambdaUltimate, myBase, myOffset)) ^ 2)
    End If
    mSafetyFactor = mLambdaUltimate
    If mLambdaUltimate > 0# Then mUtilization = 1# / mLambdaUltimate
    If Not mLastAcceptableSolver Is Nothing Then Set mLastSolver = mLastAcceptableSolver
    mLimitState = mUpperLimitState
    mConverged = True
    mStopReason = "Поиск несущей способности методом Bisection сошелся; верхнее физическое состояние: " & mUpperLimitState & "."
End Sub

' Защищенный Brent: хранит скобку как Bisection, но при хорошей геометрии
' функции пробует интерполяционный шаг. Если шаг рискованный, возвращается к
' безопасному движению внутри скобки.
Private Sub SearchByBrent(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal nOffset As Double, ByVal nBase As Double, _
        ByVal mxBase As Double, ByVal myBase As Double, _
        ByVal mxOffset As Double, ByVal myOffset As Double, _
        ByVal lowLambda As Double, ByVal highLambda As Double)
    AppendSearchDiagnostic "searchMethod=Brent"
    Dim a As Double, b As Double, c As Double
    Dim fa As Double, fb As Double, fc As Double
    Dim stateA As String, stateB As String, stateC As String
    Dim d As Double, e As Double
    Dim iter As Long

    a = lowLambda
    b = highLambda
    If Not EvaluateLoadMultiplierFunction(section, concreteMaterial, steelMaterial, nOffset, nBase, mxBase, myBase, mxOffset, myOffset, a, fa, stateA) Then Exit Sub
    If Not EvaluateLoadMultiplierFunction(section, concreteMaterial, steelMaterial, nOffset, nBase, mxBase, myBase, mxOffset, myOffset, b, fb, stateB) Then Exit Sub
    If fa >= 0# Or fb <= 0# Then
        SetInvalidInput "Методу Brent нужна скобка: нижняя lambda с запасом и верхняя lambda за физическим пределом."
        Exit Sub
    End If

    c = a
    fc = fa
    stateC = stateA
    d = b - a
    e = d

    For iter = 1 To mSolverMaxIterations
        If fb * fc > 0# Then
            c = a
            fc = fa
            stateC = stateA
            d = b - a
            e = d
        End If

        If Abs(fc) < Abs(fb) Then
            a = b: b = c: c = a
            fa = fb: fb = fc: fc = fa
            stateA = stateB: stateB = stateC: stateC = stateA
        End If

        Dim tol As Double
        Dim xm As Double
        tol = MaxDouble(mLambdaTolerance / 2#, 0.000000000001 * Abs(b))
        xm = (c - b) / 2#
        If Abs(xm) <= tol Then
            FinalizeLoadMultiplierSearchAfterProbe section, concreteMaterial, steelMaterial, nOffset, nBase, mxBase, myBase, mxOffset, myOffset, b, "Brent"
            Exit Sub
        End If

        ' Интерполяция разрешается только если она не разрушает скобку. Иначе
        ' Brent в этой итерации ведет себя как надежное деление интервала.
        If Abs(e) >= tol And Abs(fa) > Abs(fb) Then
            Dim s As Double, p As Double, q As Double
            Dim r As Double
            s = fb / fa
            If a = c Then
                p = 2# * xm * s
                q = 1# - s
            Else
                q = fa / fc
                r = fb / fc
                p = s * (2# * xm * q * (q - r) - (b - a) * (r - 1#))
                q = (q - 1#) * (r - 1#) * (s - 1#)
            End If
            If p > 0# Then q = -q
            p = Abs(p)
            If 2# * p < MinDouble(3# * xm * q - Abs(tol * q), Abs(e * q)) Then
                e = d
                d = p / q
            Else
                d = xm
                e = d
            End If
        Else
            d = xm
            e = d
        End If

        a = b
        fa = fb
        stateA = stateB
        If Abs(d) > tol Then
            b = b + d
        Else
            b = b + Sgn(xm) * tol
        End If

        If b <= lowLambda Or b >= highLambda Or Not IsFiniteNumber(b) Then
            mLimitState = STATE_NUMERICAL_FAILURE
            mStopReason = "Пробная точка Brent вышла за физическую скобку."
            Exit Sub
        End If
        If Not EvaluateLoadMultiplierFunction(section, concreteMaterial, steelMaterial, nOffset, nBase, mxBase, myBase, mxOffset, myOffset, b, fb, stateB) Then Exit Sub
    Next iter

    mLimitState = STATE_NUMERICAL_FAILURE
    mStopReason = "Поиск Brent достиг лимита итераций."
End Sub

' Честный метод секущих: новая lambda получается по двум предыдущим значениям
' целевой функции. Метод дополнительно хранит физическую скобку, чтобы не
' принять точку вне допустимого диапазона, но не подменяет шаг секущей
' бицекцией или Brent.
Private Sub SearchBySecant(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal nOffset As Double, ByVal nBase As Double, _
        ByVal mxBase As Double, ByVal myBase As Double, _
        ByVal mxOffset As Double, ByVal myOffset As Double, _
        ByVal lowLambda As Double, ByVal highLambda As Double)
    AppendSearchDiagnostic "searchMethod=Secant"
    Dim x0 As Double, x1 As Double, x2 As Double
    Dim f0 As Double, f1 As Double, f2 As Double
    Dim state0 As String, state1 As String, state2 As String
    Dim bracketLow As Double ' Последняя известная lambda с запасом по деформациям.
    Dim bracketHigh As Double ' Первая известная lambda за физическим пределом.
    Dim iter As Long

    x0 = lowLambda
    x1 = highLambda
    If Not EvaluateLoadMultiplierFunction(section, concreteMaterial, steelMaterial, nOffset, nBase, mxBase, myBase, mxOffset, myOffset, x0, f0, state0) Then Exit Sub
    If Not EvaluateLoadMultiplierFunction(section, concreteMaterial, steelMaterial, nOffset, nBase, mxBase, myBase, mxOffset, myOffset, x1, f1, state1) Then Exit Sub
    If f0 >= 0# Or f1 <= 0# Then
        SetInvalidInput "Методу Secant нужна скобка: нижняя lambda с запасом и верхняя lambda за физическим пределом."
        Exit Sub
    End If
    bracketLow = x0
    bracketHigh = x1

    For iter = 1 To mSolverMaxIterations
        If Abs(bracketHigh - bracketLow) <= mLambdaTolerance Or Abs(f1) <= mLambdaTolerance Then
            FinalizeLoadMultiplierSearchAfterProbe section, concreteMaterial, steelMaterial, nOffset, nBase, mxBase, myBase, mxOffset, myOffset, _
                bracketLow, "Secant"
            Exit Sub
        End If

        If Abs(f1 - f0) <= 0.000000000001 Then
            mLimitState = STATE_NUMERICAL_FAILURE
            mStopReason = "Поиск Secant остановлен: значения целевой функции слишком близки."
            Exit Sub
        End If

        ' Сначала строим настоящий шаг секущей по двум предыдущим значениям.
        ' Если он выходит за физическую скобку, не запускаем probe-точку
        ' заведомо вне диапазона, а демпфируем этот же шаг внутрь скобки.
        x2 = x1 - f1 * (x1 - x0) / (f1 - f0)
        If Not IsFiniteNumber(x2) Then
            mLimitState = STATE_NUMERICAL_FAILURE
            mStopReason = "Поиск Secant получил NaN или переполнение."
            Exit Sub
        End If
        x2 = GuardSecantCandidate(x2, x1, bracketLow, bracketHigh)
        If x2 <= bracketLow Or x2 >= bracketHigh Or Not IsFiniteNumber(x2) Then
            mLimitState = STATE_NUMERICAL_FAILURE
            mStopReason = "Пробная точка Secant вышла за физическую скобку."
            Exit Sub
        End If
        If Abs(x2 - x1) <= 0.000000000001 Then
            mLimitState = STATE_NUMERICAL_FAILURE
            mStopReason = "Поиск Secant застрял: новая lambda почти не отличается от предыдущей."
            Exit Sub
        End If

        If Not EvaluateLoadMultiplierFunction(section, concreteMaterial, steelMaterial, nOffset, nBase, mxBase, myBase, mxOffset, myOffset, x2, f2, state2) Then Exit Sub
        UpdateSecantBracket x2, f2, state2, bracketLow, bracketHigh
        If Abs(f2) <= mLambdaTolerance Or Abs(bracketHigh - bracketLow) <= mLambdaTolerance Then
            FinalizeLoadMultiplierSearchAfterProbe section, concreteMaterial, steelMaterial, nOffset, nBase, mxBase, myBase, mxOffset, myOffset, bracketLow, "Secant"
            Exit Sub
        End If

        x0 = x1
        f0 = f1
        state0 = state1
        x1 = x2
        f1 = f2
        state1 = state2
    Next iter

    mLimitState = STATE_NUMERICAL_FAILURE
    mStopReason = "Поиск Secant достиг лимита итераций."
End Sub

' Удерживает пробную lambda внутри найденной физической скобки.
' Это защита именно для Secant: базовый кандидат уже посчитан по секущей,
' а здесь слишком длинный шаг только уменьшается по тому же направлению.
Private Function GuardSecantCandidate(ByVal rawCandidate As Double, ByVal currentLambda As Double, _
        ByVal bracketLow As Double, ByVal bracketHigh As Double) As Double
    GuardSecantCandidate = rawCandidate
    If rawCandidate > bracketLow And rawCandidate < bracketHigh Then Exit Function

    If rawCandidate <= bracketLow Then
        If currentLambda > bracketLow And currentLambda < bracketHigh Then
            GuardSecantCandidate = (currentLambda + bracketLow) / 2#
        Else
            GuardSecantCandidate = bracketLow + 0.25 * (bracketHigh - bracketLow)
        End If
    ElseIf rawCandidate >= bracketHigh Then
        If currentLambda > bracketLow And currentLambda < bracketHigh Then
            GuardSecantCandidate = (currentLambda + bracketHigh) / 2#
        Else
            GuardSecantCandidate = bracketLow + 0.75 * (bracketHigh - bracketLow)
        End If
    End If

    AppendSearchDiagnostic "secantGuard=insideBracket; rawLambda=" & FormatNumberInvariant(rawCandidate) & _
        "; guardedLambda=" & FormatNumberInvariant(GuardSecantCandidate) & _
        "; bracketLow=" & FormatNumberInvariant(bracketLow) & _
        "; bracketHigh=" & FormatNumberInvariant(bracketHigh)
End Function

' Обновляет только защитную физическую скобку Secant.
' Отрицательная функция означает, что деформационный предел еще не достигнут;
' неотрицательная - что probe-точка уже является верхним физическим состоянием.
Private Sub UpdateSecantBracket(ByVal lambdaValue As Double, ByVal functionValue As Double, _
        ByVal probeState As String, ByRef bracketLow As Double, ByRef bracketHigh As Double)
    If functionValue < 0# Then
        If lambdaValue > bracketLow Then bracketLow = lambdaValue
    Else
        If lambdaValue < bracketHigh Then bracketHigh = lambdaValue
        If IsPhysicalLimit(probeState) Then mUpperLimitState = probeState
    End If
End Sub

' Единая целевая функция LoadMultiplier для Bisection, Brent и Secant.
' Отрицательное значение означает запас до предельной деформации, ноль -
' попадание в предел, положительное - выход за физический критерий.
Private Function EvaluateLoadMultiplierFunction(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal nOffset As Double, ByVal nBase As Double, _
        ByVal mxBase As Double, ByVal myBase As Double, _
        ByVal mxOffset As Double, ByVal myOffset As Double, _
        ByVal lambdaValue As Double, ByRef functionValue As Double, ByRef probeState As String) As Boolean
    probeState = ProbeWithRetries(section, concreteMaterial, steelMaterial, TargetN(lambdaValue, nOffset, nBase), _
        TargetMx(lambdaValue, mxBase, mxOffset), TargetMy(lambdaValue, myBase, myOffset), lambdaValue)
    If probeState = STATE_ACCEPTABLE Then
        functionValue = mCriticalStrainUtilization - 1#
        RecordAcceptableState mLastSolver, TargetN(lambdaValue, nOffset, nBase), TargetMx(lambdaValue, mxBase, mxOffset), TargetMy(lambdaValue, myBase, myOffset)
        EvaluateLoadMultiplierFunction = True
    ElseIf IsPhysicalLimit(probeState) Then
        functionValue = mCriticalStrainUtilization - 1#
        If functionValue < 0# Then functionValue = 0#
        mUpperLimitState = probeState
        EvaluateLoadMultiplierFunction = True
    ElseIf IsNumericalFailure(probeState) Then
        mLimitState = probeState
        mStopReason = "Во время поиска " & NormalizedSearchMethod() & " возникла численная несходимость."
    Else
        mLimitState = probeState
        mStopReason = "Поиск несущей способности остановлен: " & probeState & "."
    End If
End Function

' Фиксирует успешный результат LoadMultiplier: сохраняет lambda, предельные
' компоненты, запас и человекочитаемую причину остановки.
Private Sub FinalizeLoadMultiplierSearch(ByVal lambdaValue As Double, _
        ByVal nOffset As Double, ByVal nBase As Double, _
        ByVal mxBase As Double, ByVal myBase As Double, _
        ByVal mxOffset As Double, ByVal myOffset As Double, _
        ByVal limitState As String, ByVal methodName As String)
    StoreLoadPathUltimateComponents lambdaValue, nOffset, nBase, mxBase, myBase, mxOffset, myOffset
    If Not mLastSolver Is Nothing Then
        mMomentEquilibriumResidual = Sqr((mLastSolver.Mxint - TargetMx(mLambdaUltimate, mxBase, mxOffset)) ^ 2 + _
            (mLastSolver.Myint - TargetMy(mLambdaUltimate, myBase, myOffset)) ^ 2)
    End If
    mSafetyFactor = mLambdaUltimate
    If mLambdaUltimate > 0# Then mUtilization = 1# / mLambdaUltimate
    mLimitState = limitState
    If Not IsPhysicalLimit(mLimitState) Then mLimitState = mUpperLimitState
    mConverged = True
    mStopReason = "Поиск несущей способности методом " & methodName & " сошелся; верхнее физическое состояние: " & mUpperLimitState & "."
End Sub

' Перед финализацией пересчитывает probe в найденной точке, чтобы mLastSolver
' и критическое состояние соответствовали именно итоговой lambda.
Private Sub FinalizeLoadMultiplierSearchAfterProbe(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal nOffset As Double, ByVal nBase As Double, _
        ByVal mxBase As Double, ByVal myBase As Double, _
        ByVal mxOffset As Double, ByVal myOffset As Double, _
        ByVal lambdaValue As Double, ByVal methodName As String)
    Dim functionValue As Double
    Dim probeState As String
    If Not EvaluateLoadMultiplierFunction(section, concreteMaterial, steelMaterial, nOffset, nBase, mxBase, myBase, mxOffset, myOffset, _
            lambdaValue, functionValue, probeState) Then Exit Sub
    FinalizeLoadMultiplierSearch lambdaValue, nOffset, nBase, mxBase, myBase, mxOffset, myOffset, probeState, methodName
End Sub

' При двух физических граничных точках выбирает нижнюю lambda со стороны
' безопасной ветви, не полагаясь на порядок, в котором точки были найдены.
Private Function ChooseLowerPhysicalLambda(ByVal x0 As Double, ByVal f0 As Double, ByVal state0 As String, _
        ByVal x1 As Double, ByVal f1 As Double, ByVal state1 As String) As Double
    If f0 < 0# And f1 >= 0# Then
        ChooseLowerPhysicalLambda = x0
    ElseIf f1 < 0# And f0 >= 0# Then
        ChooseLowerPhysicalLambda = x1
    ElseIf f0 < f1 Then
        ChooseLowerPhysicalLambda = x0
    Else
        ChooseLowerPhysicalLambda = x1
    End If
End Function

' Выбирает lambda по самой устойчивой масштабируемой компоненте нагрузки.
' Компоненты имеют разные размерности, поэтому сравниваем их не напрямую, а
' относительно допусков равновесия solver-а.
Private Function LoadPathLambda(ByVal nValue As Double, ByVal mxValue As Double, ByVal myValue As Double, _
        ByVal nOffset As Double, ByVal nBase As Double, _
        ByVal mxOffset As Double, ByVal mxBase As Double, _
        ByVal myOffset As Double, ByVal myBase As Double) As Double
    Dim math As CLoadPathMath
    Set math = NewLoadPathMath()
    LoadPathLambda = math.Lambda(nValue, mxValue, myValue, _
        LoadPathFromComponents(nOffset, nBase, mxOffset, mxBase, myOffset, myBase))
End Function

' Собирает две независимые невязки, которые удерживают найденное внутреннее
' усилие на прямой Target = Offset + lambda*Base. Если компонент не
' масштабируется, она контролируется как постоянная. Если масштабируется,
' она должна совпасть со своим значением при том же lambda.
Private Sub BuildLoadPathResiduals(ByVal nValue As Double, ByVal mxValue As Double, ByVal myValue As Double, _
        ByVal nOffset As Double, ByVal nBase As Double, _
        ByVal mxOffset As Double, ByVal mxBase As Double, _
        ByVal myOffset As Double, ByVal myBase As Double, _
        ByVal lambdaValue As Double, ByRef rPath1 As Double, ByRef rPath2 As Double)
    Dim math As CLoadPathMath
    Set math = NewLoadPathMath()
    math.BuildResiduals nValue, mxValue, myValue, _
        LoadPathFromComponents(nOffset, nBase, mxOffset, mxBase, myOffset, myBase), _
        lambdaValue, rPath1, rPath2
End Sub

' Создает общий helper load-path с текущими solver tolerance capacity.
Private Function NewLoadPathMath() As CLoadPathMath
    Dim math As CLoadPathMath
    Set math = New CLoadPathMath
    math.Configure mSolverToleranceN, mSolverToleranceMx, mSolverToleranceMy
    Set NewLoadPathMath = math
End Function

' Упаковывает старые аргументы CCapacitySolver в общий CLoadPathVector, чтобы
' вся математика lambda-траектории проходила через CLoadPathMath.
Private Function LoadPathFromComponents(ByVal nOffset As Double, ByVal nBase As Double, _
        ByVal mxOffset As Double, ByVal mxBase As Double, _
        ByVal myOffset As Double, ByVal myBase As Double) As CLoadPathVector
    Dim loadPath As CLoadPathVector
    Set loadPath = New CLoadPathVector
    loadPath.Initialize nOffset, nBase, mxOffset, mxBase, myOffset, myBase
    Set LoadPathFromComponents = loadPath
End Function

' Возвращает целевой Mx на lambda-траектории.
Private Function TargetMx(ByVal lambdaValue As Double, ByVal mxBase As Double, ByVal mxOffset As Double) As Double
    TargetMx = mxOffset + lambdaValue * mxBase
End Function

' Возвращает целевой My на lambda-траектории.
Private Function TargetMy(ByVal lambdaValue As Double, ByVal myBase As Double, ByVal myOffset As Double) As Double
    TargetMy = myOffset + lambdaValue * myBase
End Function

' Возвращает целевую продольную силу на lambda-траектории.
Private Function TargetN(ByVal lambdaValue As Double, ByVal nOffset As Double, ByVal nBase As Double) As Double
    TargetN = nOffset + lambdaValue * nBase
End Function

' Записывает итоговые компоненты нагрузки для общей lambda-траектории.
' Этот низкоуровневый результат хранит полное предельное состояние solver-а:
' постоянные компоненты Offset и масштабируемые Base уже сложены вместе.
' Batch-слой при выводе дополнительно переносит полный найденный вектор
' обратно в пользовательскую точку приложения нагрузки.
Private Sub StoreLoadPathUltimateComponents(ByVal lambdaValue As Double, _
        ByVal nOffset As Double, ByVal nBase As Double, _
        ByVal mxBase As Double, ByVal myBase As Double, _
        ByVal mxOffset As Double, ByVal myOffset As Double)
    mLambdaUltimate = lambdaValue
    If CapacityScalesAxial(nBase) Then
        mNUltimate = TargetN(mLambdaUltimate, nOffset, nBase)
        mMxUltimate = TargetMx(mLambdaUltimate, mxBase, mxOffset)
        mMyUltimate = TargetMy(mLambdaUltimate, myBase, myOffset)
    Else
        mMxUltimate = mLambdaUltimate * mxBase
        mMyUltimate = mLambdaUltimate * myBase
    End If
    mMomentUltimate = Sqr(mMxUltimate * mMxUltimate + mMyUltimate * mMyUltimate)
End Sub

' Определяет, входит ли продольная сила в масштабируемую часть capacity-path.
' Для чисто моментной траектории N хранится как постоянный offset.
Private Function CapacityScalesAxial(ByVal nBase As Double) As Boolean
    CapacityScalesAxial = (Abs(nBase) > 0.000000001)
End Function

' Проверяет общую lambda-траекторию. Нулевой Base означает, что ни одна
' компонента нагрузки не масштабируется, поэтому предельный множитель не имеет
' инженерного смысла.
Private Function ValidateLoadPathCapacityInputs(ByVal section As CSectionModel, ByVal concreteMaterial As Object, _
        ByVal steelMaterial As Object, ByVal nBase As Double, ByVal mxBase As Double, ByVal myBase As Double) As Boolean
    If section Is Nothing Then
        SetInvalidInput "Модель сечения не передана."
    ElseIf section.ConcreteCount <= 0 Then
        SetInvalidInput "В модели сечения нет бетонных элементов."
    ElseIf concreteMaterial Is Nothing Then
        SetInvalidInput "Материал бетона не передан."
    ElseIf steelMaterial Is Nothing Then
        SetInvalidInput "Материал арматуры не передан."
    ElseIf Abs(nBase) <= 0.000000001 And Abs(mxBase) <= 0.000000001 And Abs(myBase) <= 0.000000001 Then
        SetInvalidInput "В траектории несущей способности нет масштабируемой нагрузки."
    ElseIf mSolverMaxIterations <= 0 Then
        SetInvalidInput "Настройки поиска несущей способности некорректны."
    Else
        ValidateLoadPathCapacityInputs = True
    End If
End Function

' Проверяет входные данные и прерывает выполнение понятной ошибкой, если расчетный контракт нарушен.
Private Function ValidateCapacityInputs(ByVal section As CSectionModel, ByVal concreteMaterial As Object, _
        ByVal steelMaterial As Object, ByVal mxBase As Double, ByVal myBase As Double) As Boolean
    If section Is Nothing Then
        SetInvalidInput "Модель сечения не передана."
    ElseIf section.ConcreteCount <= 0 Then
        SetInvalidInput "В модели сечения нет бетонных элементов."
    ElseIf concreteMaterial Is Nothing Then
        SetInvalidInput "Материал бетона не передан."
    ElseIf steelMaterial Is Nothing Then
        SetInvalidInput "Материал арматуры не передан."
    ElseIf Sqr(mxBase * mxBase + myBase * myBase) <= 0.000000001 Then
        SetInvalidInput "Базовый вектор момента должен быть ненулевым."
    ElseIf mSolverMaxIterations <= 0 Then
        SetInvalidInput "Настройки поиска несущей способности некорректны."
    Else
        ValidateCapacityInputs = True
    End If
End Function

' Проверяет настройки ветки LoadMultiplier перед одномерным поиском lambda.
' Здесь контролируются только параметры самой скобки/поиска, а не геометрия
' или материалы, которые проверяются выше по цепочке.
Private Function ValidateLoadMultiplierSettings() As Boolean
    If Len(NormalizedSearchMethod()) = 0 Then
        SetInvalidInput "Capacity.SearchMethod должен быть Bisection, Brent или Secant."
    ElseIf mInitialLambdaStep <= 0# Or mLambdaTolerance <= 0# Or mMaxLambda <= 0# Or _
            mMaxRetries < 0 Or mSolverBaseLoadSteps <= 0 Then
        SetInvalidInput "Настройки LoadMultiplier для несущей способности некорректны."
    Else
        ValidateLoadMultiplierSettings = True
    End If
End Function

' Проверяет настройки ветки UltimateStrain. В отличие от LoadMultiplier,
' здесь нет искусственного Capacity.MaxLambda: после найденной предельной
' плоскости достаточно проверить численную корректность lambda.
Private Function ValidateUltimateStrainSettings() As Boolean
    If mStrainTolerance <= 0# Then
        SetInvalidInput "Настройки UltimateStrain для несущей способности некорректны."
    Else
        ValidateUltimateStrainSettings = True
    End If
End Function

' Переводит расчет в InputErr без выбрасывания VBA-ошибки. Capacity solver
' возвращает статус пользователю через batch summary и подробный writer.
Private Sub SetInvalidInput(ByVal reason As String)
    mLimitState = STATE_INVALID_INPUT
    mStopReason = reason
End Sub

' Проверяет только численную пригодность найденного UltimateStrain-множителя.
' MaxLambda здесь намеренно не применяется: это граница расширения скобки
' LoadMultiplier, а UltimateStrain уже нашел предельную плоскость деформаций.
Private Function IsValidUltimateLambda(ByVal lambdaValue As Double, _
        ByVal nOffset As Double, ByVal nBase As Double, _
        ByVal mxBase As Double, ByVal myBase As Double, _
        ByVal mxOffset As Double, ByVal myOffset As Double) As Boolean
    If lambdaValue <= 0# Then Exit Function
    If Not IsFiniteNumber(lambdaValue) Then Exit Function
    If Not IsFiniteNumber(TargetN(lambdaValue, nOffset, nBase)) Then Exit Function
    If Not IsFiniteNumber(TargetMx(lambdaValue, mxBase, mxOffset)) Then Exit Function
    If Not IsFiniteNumber(TargetMy(lambdaValue, myBase, myOffset)) Then Exit Function
    IsValidUltimateLambda = True
End Function

' Выполняет probe-точку LoadMultiplier с ограниченным числом повторов.
' Повтор меняет только численные настройки внутреннего CSectionSolver
' (ступени/шаги), а физическая точка нагрузки lambda остается той же.
Private Function ProbeWithRetries(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal nSpecified As Double, ByVal mx As Double, ByVal my As Double, ByVal lambdaValue As Double) As String

    Dim cachedState As String
    If TryRestoreProbeCache(nSpecified, mx, my, lambdaValue, cachedState) Then
        ProbeWithRetries = cachedState
        Exit Function
    End If

    Dim attempt As Long
    Dim probeState As String
    For attempt = 0 To mMaxRetries
        probeState = ProbeOnce(section, concreteMaterial, steelMaterial, nSpecified, mx, my, lambdaValue, attempt)
        If Not IsNumericalFailure(probeState) Then
            StoreProbeCache nSpecified, mx, my, lambdaValue, probeState
            ProbeWithRetries = probeState
            Exit Function
        End If
        If attempt < mMaxRetries Then mRetryCount = mRetryCount + 1
    Next attempt

    StoreProbeCache nSpecified, mx, my, lambdaValue, probeState
    ProbeWithRetries = probeState
End Function

' Решает одну probe-точку capacity-поиска и классифицирует ее как приемлемую,
' физический предел, численную несходимость или ошибку входных данных.
Private Function ProbeOnce(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal nSpecified As Double, ByVal mx As Double, ByVal my As Double, _
        ByVal lambdaValue As Double, ByVal attempt As Long) As String

    On Error GoTo Failed
    mIterations = mIterations + 1

    Dim solver As CSectionSolver
    Set solver = New CSectionSolver
    solver.LoadSteps = mSolverBaseLoadSteps * (2 ^ attempt)
    solver.SolverMethod = mSolverMethod
    solver.MaxIterations = mSolverMaxIterations
    solver.ToleranceN = mSolverToleranceN
    solver.ToleranceMx = mSolverToleranceMx
    solver.ToleranceMy = mSolverToleranceMy
    solver.LineSearchEnabled = mSolverLineSearchEnabled
    solver.DampingInitial = mSolverDampingInitial
    solver.MinLineSearchAlpha = mSolverMinLineSearchAlpha
    solver.MaxDeltaEpsilon0 = mSolverMaxDeltaEpsilon0
    solver.MaxDeltaKappa = mSolverMaxDeltaKappa / (2 ^ attempt)
    solver.SecantMaxRestarts = mSolverSecantMaxRestarts
    solver.SecantMinStepNorm = mSolverSecantMinStepNorm
    solver.DiagnosticsEnabled = mDiagnosticsEnabled
    If Not mLastAcceptableSolver Is Nothing Then
        solver.SetInitialLoadState mLastAcceptableSolver.Epsilon0, mLastAcceptableSolver.KappaX, mLastAcceptableSolver.KappaY, _
            mLastAcceptableN, mLastAcceptableMx, mLastAcceptableMy
    End If
    ApplyPureBendingProbeGuess section, concreteMaterial, steelMaterial, solver, nSpecified, mx, my

    solver.Solve section, concreteMaterial, steelMaterial, nSpecified, mx, my
    Set mLastSolver = solver

    If Not solver.Converged Then
        If FailedProbeReachedPhysicalLimit(section, concreteMaterial, solver) Then
            ProbeOnce = CriticalLimitState()
        Else
            ProbeOnce = ClassifyNumericalFailure(solver.StopReason)
        End If
    Else
        UpdateCriticalState section, concreteMaterial, solver
        If mCriticalStrainUtilization >= 1# Then
            ProbeOnce = CriticalLimitState()
        Else
            ProbeOnce = STATE_ACCEPTABLE
        End If
    End If

    AppendDiagnostic solver, mx, my, lambdaValue, attempt, ProbeOnce
    Exit Function

Failed:
    ProbeOnce = STATE_INVALID_INPUT
    mLimitState = STATE_INVALID_INPUT
    mStopReason = "Пробная точка несущей способности не рассчиталась: " & Err.Description
End Function

' Для чистого изгиба N=0 каждая probe-точка LoadMultiplier может оказаться
' в той же численной ловушке, что и прямой DirectState: нулевая плоскость
' деформаций лежит ровно на изломе "сжатый бетон / выключенное растяжение".
' Поэтому для такой probe-точки старт строится от ее собственных целевых
' усилий, а не от предыдущей lambda: это устойчивее около границы плато и не
' зависит от того, какая точка перед этим была последней acceptable.
' Траектория Target = Offset + lambda*Base при этом не меняется: мы только
' подсказываем внутреннему CSectionSolver, с какой стороны излома начинать.
Private Sub ApplyPureBendingProbeGuess(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal solver As CSectionSolver, ByVal nSpecified As Double, _
        ByVal mx As Double, ByVal my As Double)
    If solver Is Nothing Then Exit Sub
    If section Is Nothing Then Exit Sub
    If Abs(nSpecified) > 0.000000001 Then Exit Sub
    If Sqr(mx * mx + my * my) <= 0.000000001 Then Exit Sub
    If TypeName(concreteMaterial) <> "CMaterialDiagram" Then Exit Sub
    If TypeName(steelMaterial) <> "CMaterialDiagram" Then Exit Sub

    Dim concreteDiagram As CMaterialDiagram
    Dim steelDiagram As CMaterialDiagram
    Set concreteDiagram = concreteMaterial
    Set steelDiagram = steelMaterial

    Dim guessBuilder As CStateGuessBuilder
    Set guessBuilder = New CStateGuessBuilder

    Dim eps0 As Double
    Dim kx As Double
    Dim ky As Double
    If guessBuilder.BuildForDirectState(section, concreteDiagram, steelDiagram, _
            nSpecified, mx, my, eps0, kx, ky) Then
        solver.ClearInitialState
        solver.SetInitialLoadState eps0, kx, ky, nSpecified, mx, my
        AppendSearchDiagnostic "для пробной точки чистого изгиба применена стартовая плоскость; масштабируемые моменты ненулевые"
    End If
End Sub

' Возвращает уже посчитанную probe-точку LoadMultiplier.
' Brent/Secant после нахождения скобки могут повторно спросить те же lambda,
' особенно нижнюю и верхнюю границы. Cache убирает этот повторный solve и
' восстанавливает все поля, от которых зависит целевая функция и warm-start.
Private Function TryRestoreProbeCache(ByVal nSpecified As Double, ByVal mx As Double, ByVal my As Double, _
        ByVal lambdaValue As Double, ByRef probeState As String) As Boolean
    Dim slot As Long
    For slot = 1 To mProbeCacheCount
        If ProbeCacheMatches(slot, nSpecified, mx, my, lambdaValue) Then
            probeState = mProbeCacheState(slot)
            Set mLastSolver = mProbeCacheSolver(slot)
            mCriticalStrain = mProbeCacheCriticalStrain(slot)
            mCriticalElement = mProbeCacheCriticalElement(slot)
            mCriticalStrainUtilization = mProbeCacheCriticalUtil(slot)
            AppendSearchDiagnostic "использована ранее рассчитанная пробная точка; lambda=" & FormatNumberInvariant(lambdaValue) & "; состояние=" & probeState
            TryRestoreProbeCache = True
            Exit Function
        End If
    Next slot
End Function

' Сравнивает probe-точку по lambda и целевым усилиям.
' Усилия проверяются отдельно от lambda, чтобы cache оставался безопасным даже
' после будущих расширений траектории нагрузки внутри одного solver-а.
Private Function ProbeCacheMatches(ByVal slot As Long, ByVal nSpecified As Double, ByVal mx As Double, ByVal my As Double, _
        ByVal lambdaValue As Double) As Boolean
    If Abs(mProbeCacheLambda(slot) - lambdaValue) > 0.000000000001 * MaxDouble(1#, Abs(lambdaValue)) Then Exit Function
    If Abs(mProbeCacheN(slot) - nSpecified) > MaxDouble(mSolverToleranceN * 0.000000001, 0.000001) Then Exit Function
    If Abs(mProbeCacheMx(slot) - mx) > MaxDouble(mSolverToleranceMx * 0.000000001, 0.001) Then Exit Function
    If Abs(mProbeCacheMy(slot) - my) > MaxDouble(mSolverToleranceMy * 0.000000001, 0.001) Then Exit Function
    ProbeCacheMatches = True
End Function

' Запоминает итог probe-точки после всех retry.
' NumericalFailure тоже можно сохранить: если метод снова спросит ту же lambda,
' повторная попытка с теми же настройками даст тот же отказ, а cache не даст
' впустую запускать внутренний CSectionSolver.
Private Sub StoreProbeCache(ByVal nSpecified As Double, ByVal mx As Double, ByVal my As Double, _
        ByVal lambdaValue As Double, ByVal probeState As String)
    If mProbeCacheNextSlot < 1 Or mProbeCacheNextSlot > PROBE_CACHE_SIZE Then mProbeCacheNextSlot = 1

    Dim slot As Long
    slot = mProbeCacheNextSlot
    mProbeCacheLambda(slot) = lambdaValue
    mProbeCacheN(slot) = nSpecified
    mProbeCacheMx(slot) = mx
    mProbeCacheMy(slot) = my
    mProbeCacheState(slot) = probeState
    mProbeCacheCriticalStrain(slot) = mCriticalStrain
    mProbeCacheCriticalElement(slot) = mCriticalElement
    mProbeCacheCriticalUtil(slot) = mCriticalStrainUtilization
    Set mProbeCacheSolver(slot) = mLastSolver

    If mProbeCacheCount < PROBE_CACHE_SIZE Then mProbeCacheCount = mProbeCacheCount + 1
    mProbeCacheNextSlot = slot + 1
    If mProbeCacheNextSlot > PROBE_CACHE_SIZE Then mProbeCacheNextSlot = 1
End Sub

' Полностью очищает cache между независимыми расчетами.
Private Sub ClearProbeCache()
    Dim slot As Long
    For slot = 1 To PROBE_CACHE_SIZE
        Set mProbeCacheSolver(slot) = Nothing
        mProbeCacheState(slot) = vbNullString
        mProbeCacheCriticalElement(slot) = vbNullString
    Next slot
    mProbeCacheCount = 0
    mProbeCacheNextSlot = 1
End Sub

' Проверяет неудачную probe-точку LoadMultiplier на физический предел.
' Около разрушения внутренний CSectionSolver может не успеть закрыть невязку
' равновесия, хотя рассчитанная плоскость деформаций уже прошла допустимую
' сжимающую/растягивающую деформацию. Такое состояние годится как верхняя
' граница lambda, но только если solver сделал больше одной итерации: одиночный
' неудачный шаг слишком груб, чтобы считать его надежным физическим пределом.
Private Function FailedProbeReachedPhysicalLimit(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal solver As CSectionSolver) As Boolean
    If solver Is Nothing Then Exit Function
    If solver.Iterations <= 1 Then Exit Function

    UpdateCriticalState section, concreteMaterial, solver
    If mCriticalStrainUtilization >= 1# And IsFiniteNumber(mCriticalStrainUtilization) And _
            Len(mCriticalElement) > 0 Then
        AppendSearchDiagnostic "несошедшаяся пробная точка достигла физического предела деформации; состояние=" & _
            CriticalLimitState() & "; использование=" & FormatNumberInvariant(mCriticalStrainUtilization) & _
            "; элемент=" & mCriticalElement & "; причина solver-а=" & solver.StopReason
        FailedProbeReachedPhysicalLimit = True
    End If
End Function

' Обновляет управляющую деформацию для последней probe-точки. Проверяются
' бетон, растянутый бетон при включенном лимите и арматура; максимальная
' utilization определяет тип предельного состояния.
Private Sub UpdateCriticalState(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal solver As CSectionSolver)
    mCriticalStrainUtilization = 0#
    mCriticalStrain = 0#
    mCriticalElement = vbNullString

    Dim i As Long
    Dim strain As Double
    Dim utilization As Double
    Dim xCoord As Double
    Dim yCoord As Double
    For i = 1 To section.ConcreteCount
        xCoord = section.ConcreteX(i)
        yCoord = section.ConcreteY(i)
        strain = solver.Epsilon0 + solver.KappaX * yCoord + solver.KappaY * xCoord
        If strain < 0# And mConcreteCompressionLimit < 0# Then
            utilization = strain / mConcreteCompressionLimit
            StoreCriticalIfLarger utilization, strain, section.ConcreteID(i) & ":Compression"
        ElseIf strain > 0# And mConcreteTensionLimitEnabled And mConcreteTensionLimit > 0# Then
            utilization = strain / mConcreteTensionLimit
            StoreCriticalIfLarger utilization, strain, section.ConcreteID(i) & ":Tension"
        End If
    Next i

    If section.RebarCount > 0 Then
        For i = 1 To section.RebarCount
            xCoord = section.RebarX(i)
            yCoord = section.RebarY(i)
            strain = solver.Epsilon0 + solver.KappaX * yCoord + solver.KappaY * xCoord
            If mSteelStrainLimit > 0# Then
                utilization = Abs(strain) / mSteelStrainLimit
                StoreCriticalIfLarger utilization, strain, section.RebarID(i) & ":Steel"
            End If
        Next i
    End If
End Sub

' Сохраняет наиболее опасный элемент, если его utilization больше текущего.
Private Sub StoreCriticalIfLarger(ByVal utilization As Double, ByVal strain As Double, ByVal elementName As String)
    If utilization > mCriticalStrainUtilization Then
        mCriticalStrainUtilization = utilization
        mCriticalStrain = strain
        mCriticalElement = elementName
    End If
End Sub

' Переводит имя критического элемента в расчетный LimitState.
Private Function CriticalLimitState() As String
    If InStr(1, mCriticalElement, ":Compression", vbTextCompare) > 0 Then
        CriticalLimitState = STATE_CONCRETE_LIMIT
    ElseIf InStr(1, mCriticalElement, ":Tension", vbTextCompare) > 0 Then
        CriticalLimitState = STATE_CONCRETE_TENSION_LIMIT
    ElseIf InStr(1, mCriticalElement, ":Steel", vbTextCompare) > 0 Then
        CriticalLimitState = STATE_STEEL_LIMIT
    Else
        CriticalLimitState = STATE_NUMERICAL_FAILURE
    End If
End Function

' Разделяет обычную NumericalFailure и подозрение на вырожденную касательную
' матрицу по тексту причины CSectionSolver.
Private Function ClassifyNumericalFailure(ByVal reason As String) As String
    If InStr(1, reason, "Tangent matrix", vbTextCompare) > 0 Or _
            InStr(1, reason, "касательн", vbTextCompare) > 0 Or _
            InStr(1, reason, "singular", vbTextCompare) > 0 Or _
            InStr(1, reason, "вырожден", vbTextCompare) > 0 Or _
            InStr(1, reason, "solve failed", vbTextCompare) > 0 Or _
            InStr(1, reason, "не реш", vbTextCompare) > 0 Then
        ClassifyNumericalFailure = STATE_SINGULAR_TANGENT
    Else
        ClassifyNumericalFailure = STATE_NUMERICAL_FAILURE
    End If
End Function

' Запоминает последнюю сходящуюся probe-точку. Она используется как warm-start
' для следующих lambda и как fallback, если верхняя физическая граница не
' найдена из-за численной несходимости.
Private Sub RecordAcceptableState(ByVal solver As CSectionSolver, ByVal nValue As Double, ByVal mxValue As Double, ByVal myValue As Double)
    Set mLastAcceptableSolver = solver
    mLastAcceptableN = nValue
    mLastAcceptableMx = mxValue
    mLastAcceptableMy = myValue
End Sub

' Уменьшает верхнюю границу lambda после NumericalFailure. Если интервал уже
' слишком мал, поиск возвращает последнюю acceptable-точку вместо того, чтобы
' выдавать неустойчивый физический предел.
Private Function ReduceLambdaAfterNumericalFailure(ByVal lowLambda As Double, ByVal highLambda As Double, ByVal probeState As String) As Double
    Dim reduced As Double
    reduced = lowLambda + (highLambda - lowLambda) / 2#
    If (reduced - lowLambda) <= mLambdaTolerance Then
        mLimitState = probeState
        mStopReason = "Численная несходимость возникла до нахождения верхней физической границы; сохранена последняя сходящаяся lambda."
        mLambdaUltimate = lowLambda
        If Not mLastAcceptableSolver Is Nothing Then Set mLastSolver = mLastAcceptableSolver
        ReduceLambdaAfterNumericalFailure = lowLambda
    Else
        AppendSearchDiagnostic "численная несходимость; верхняя lambda уменьшена с " & FormatNumberInvariant(highLambda) & _
            " до " & FormatNumberInvariant(reduced) & "; состояние=" & probeState
        ReduceLambdaAfterNumericalFailure = reduced
    End If
End Function

' Проверяет, является ли состояние физическим пределом несущей способности.
Private Function IsPhysicalLimit(ByVal state As String) As Boolean
    IsPhysicalLimit = (state = STATE_CONCRETE_LIMIT Or state = STATE_CONCRETE_TENSION_LIMIT Or _
        state = STATE_STEEL_LIMIT)
End Function

' Проверяет, является ли состояние численной неудачей, которую можно пытаться
' обходить retry, уменьшением lambda или fallback-стратегией.
Private Function IsNumericalFailure(ByVal state As String) As Boolean
    IsNumericalFailure = (state = STATE_NUMERICAL_FAILURE Or state = STATE_SINGULAR_TANGENT)
End Function

' Нормализует метод одномерного поиска LoadMultiplier из пользовательской
' настройки Capacity.SearchMethod.
Private Function NormalizedSearchMethod() As String
    Select Case LCase$(Trim$(mSearchMethod))
        Case "bisection"
            NormalizedSearchMethod = "Bisection"
        Case "brent"
            NormalizedSearchMethod = "Brent"
        Case "secant"
            NormalizedSearchMethod = "Secant"
    End Select
End Function

' Проверяет, что число конечно и пригодно для дальнейшей арифметики.
Private Function IsFiniteNumber(ByVal value As Double) As Boolean
    If value <> value Then Exit Function
    If Abs(value) > 1E+100 Then Exit Function
    IsFiniteNumber = True
End Function

' Записывает диагностическую строку одной probe-точки LoadMultiplier.
' Поле attempt означает номер retry той же физической lambda-точки.
Private Sub AppendDiagnostic(ByVal solver As CSectionSolver, ByVal mx As Double, ByVal my As Double, _
        ByVal lambdaValue As Double, ByVal attempt As Long, ByVal probeState As String)
    If Not mDiagnosticsEnabled Then Exit Sub
    mDiagnosticLog = mDiagnosticLog & "probe=" & CStr(mIterations) & _
        "; attempt=" & CStr(attempt) & _
        "; lambda=" & FormatNumberInvariant(lambdaValue) & _
        "; Mx=" & FormatNumberInvariant(mx) & _
        "; My=" & FormatNumberInvariant(my) & _
        "; converged=" & CStr(solver.Converged) & _
        "; state=" & probeState & _
        "; stop=" & solver.StopReason & _
        "; minConcreteStrain=" & FormatNumberInvariant(solver.MinConcreteStrain) & _
        "; maxConcreteStrain=" & FormatNumberInvariant(solver.MaxConcreteStrain) & _
        "; maxAbsSteelStrain=" & FormatNumberInvariant(solver.MaxAbsSteelStrain) & vbCrLf
End Sub

' Добавляет человекочитаемое сообщение о ходе одномерного поиска lambda.
Private Sub AppendSearchDiagnostic(ByVal message As String)
    If Not mDiagnosticsEnabled Then Exit Sub
    mDiagnosticLog = mDiagnosticLog & "search=" & message & vbCrLf
End Sub

' Очищает результаты предыдущего capacity-поиска, не меняя настройки solver-а
' и предельные деформации материалов.
Private Sub ClearResult()
    mConverged = False
    mStopReason = vbNullString
    mLimitState = vbNullString
    mUpperLimitState = vbNullString
    mLambdaUltimate = 0#
    mNUltimate = 0#
    mMomentUltimate = 0#
    mMxUltimate = 0#: mMyUltimate = 0#
    mSafetyFactor = 0#
    mUtilization = 0#
    mIterations = 0
    mRetryCount = 0
    mMomentEquilibriumResidual = 0#
    mCriticalStrain = 0#
    mCriticalElement = vbNullString
    mCriticalStrainUtilization = 0#
    mLastAcceptableN = 0#: mLastAcceptableMx = 0#: mLastAcceptableMy = 0#
    Set mLastSolver = Nothing
    Set mLastAcceptableSolver = Nothing
    mDiagnosticLog = vbNullString
    mSolutionMethod = vbNullString
    ClearProbeCache
End Sub

' Форматирует число для диагностического журнала с точкой как разделителем.
Private Function FormatNumberInvariant(ByVal value As Double) As String
    FormatNumberInvariant = Replace$(Format$(value, "0.############"), ",", ".")
End Function

' Утилитарный минимум без WorksheetFunction в расчетном ядре.
Private Function MinDouble(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then MinDouble = a Else MinDouble = b
End Function

' Утилитарный максимум без WorksheetFunction в расчетном ядре.
Private Function MaxDouble(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then MaxDouble = a Else MaxDouble = b
End Function

Public Property Get Converged() As Boolean
    Converged = mConverged
End Property

Public Property Get StopReason() As String
    StopReason = mStopReason
End Property

Public Property Get LimitState() As String
    LimitState = mLimitState
End Property

Public Property Get UpperLimitState() As String
    UpperLimitState = mUpperLimitState
End Property

Public Property Get LambdaUltimate() As Double
    LambdaUltimate = mLambdaUltimate
End Property

Public Property Get NUltimate() As Double
    NUltimate = mNUltimate
End Property

Public Property Get MomentUltimate() As Double
    MomentUltimate = mMomentUltimate
End Property

Public Property Get MxUltimate() As Double
    MxUltimate = mMxUltimate
End Property

Public Property Get MyUltimate() As Double
    MyUltimate = mMyUltimate
End Property

Public Property Get SafetyFactor() As Double
    SafetyFactor = mSafetyFactor
End Property

Public Property Get Utilization() As Double
    Utilization = mUtilization
End Property

Public Property Get Iterations() As Long
    Iterations = mIterations
End Property

Public Property Get MomentEquilibriumResidual() As Double
    MomentEquilibriumResidual = mMomentEquilibriumResidual
End Property

Public Property Get DiagnosticLog() As String
    DiagnosticLog = mDiagnosticLog
End Property

Public Property Let DiagnosticsEnabled(ByVal value As Boolean)
    mDiagnosticsEnabled = value
End Property

Public Property Get LastSolver() As CSectionSolver
    Set LastSolver = mLastSolver
End Property

Public Property Let ConcreteCompressionLimit(ByVal value As Double)
    mConcreteCompressionLimit = value
End Property

Public Property Let ConcreteTensionLimit(ByVal value As Double)
    mConcreteTensionLimit = value
End Property

Public Property Let ConcreteTensionLimitEnabled(ByVal value As Boolean)
    mConcreteTensionLimitEnabled = value
End Property

Public Property Let SteelStrainLimit(ByVal value As Double)
    mSteelStrainLimit = value
End Property

Public Property Let LambdaTolerance(ByVal value As Double)
    mLambdaTolerance = value
End Property

Public Property Let StrainTolerance(ByVal value As Double)
    mStrainTolerance = value
End Property

Public Property Let SearchMethod(ByVal value As String)
    mSearchMethod = value
End Property

Public Property Get SearchMethod() As String
    SearchMethod = mSearchMethod
End Property

Public Property Get SolutionMethod() As String
    SolutionMethod = mSolutionMethod
End Property

Public Property Let MaxLambda(ByVal value As Double)
    mMaxLambda = value
End Property

Public Property Let InitialLambdaStep(ByVal value As Double)
    mInitialLambdaStep = value
End Property

Public Property Let MaxRetries(ByVal value As Long)
    mMaxRetries = value
End Property

Public Property Let SolverBaseLoadSteps(ByVal value As Long)
    mSolverBaseLoadSteps = value
End Property

Public Property Let SolverMaxIterations(ByVal value As Long)
    mSolverMaxIterations = value
End Property

Public Property Let SolverToleranceN(ByVal value As Double)
    mSolverToleranceN = value
End Property

Public Property Let SolverToleranceMx(ByVal value As Double)
    mSolverToleranceMx = value
End Property

Public Property Let SolverToleranceMy(ByVal value As Double)
    mSolverToleranceMy = value
End Property

Public Property Let SolverMethod(ByVal value As String)
    mSolverMethod = value
End Property

Public Property Let SolverLineSearchEnabled(ByVal value As Boolean)
    mSolverLineSearchEnabled = value
End Property

Public Property Let SolverDampingInitial(ByVal value As Double)
    mSolverDampingInitial = value
End Property

Public Property Let SolverMinLineSearchAlpha(ByVal value As Double)
    mSolverMinLineSearchAlpha = value
End Property

Public Property Let SolverMaxDeltaEpsilon0(ByVal value As Double)
    mSolverMaxDeltaEpsilon0 = value
End Property

Public Property Let SolverMaxDeltaKappa(ByVal value As Double)
    mSolverMaxDeltaKappa = value
End Property

Public Property Let SolverSecantMaxRestarts(ByVal value As Long)
    mSolverSecantMaxRestarts = value
End Property

Public Property Let SolverSecantMinStepNorm(ByVal value As Double)
    mSolverSecantMinStepNorm = value
End Property

' ==========================================================================
' ДЛЯ ТЕСТОВ
' ==========================================================================
' Эти свойства раскрывают внутреннюю диагностику поиска несущей способности.
' Они нужны regression-тестам, чтобы убедиться, что solver действительно
' дошел до нужного предельного элемента, использовал retry-ветку и корректно
' оценил utilization. Пользовательский вывод берет уже агрегированные поля
' SafetyFactor, LimitState, MomentUltimate и статусы batch-расчета.

Public Property Get RetryCount() As Long
    RetryCount = mRetryCount
End Property

Public Property Get CriticalStrain() As Double
    CriticalStrain = mCriticalStrain
End Property

Public Property Get CriticalElement() As String
    CriticalElement = mCriticalElement
End Property

Public Property Get CriticalStrainUtilization() As Double
    CriticalStrainUtilization = mCriticalStrainUtilization
End Property








