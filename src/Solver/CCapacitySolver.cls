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
' На вход класса передается универсальная траектория нагрузки:
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
Private Const STATE_INTERNAL_ERROR As String = "InternalError" ' Нарушен программный контракт; повтор численного solve не помогает.
Private Const MOMENT_DIRECTION_TOLERANCE As Double = 0.001 ' Допустимое относительное отклонение от выбранной моментной траектории.
Private Const PROBE_CACHE_SIZE As Long = 12 ' Небольшое кольцо последних lambda-точек LoadMultiplier.

Private mConverged As Boolean ' True, если выбранный метод capacity нашел предельное состояние.
Private mStopReason As String ' Причина остановки capacity solver для вывода в результаты и тесты.
Private mLimitState As String ' Физический или численный статус найденного предела.
Private mFailureCode As ESolverFailureCode ' Точная техническая причина, независимая от текста комментария и LimitState.
Private mUpperLimitState As String ' Статус верхней границы при LoadMultiplier.
Private mInitialLambdaRejected As Boolean ' True, если lambda=0 уже не является acceptable-точкой capacity path.
Private mSearchBoundReached As Boolean ' MaxLambda проверена, но физический предел не установлен.
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
Private mSteelCompressionLimit As Double ' Модуль предельной сжимающей деформации арматуры.
Private mSteelTensionLimit As Double ' Предельная растягивающая деформация арматуры.
Private mLambdaTolerance As Double ' Точность одномерного поиска lambda для LoadMultiplier.
Private mStrainTolerance As Double ' Допуск достижения предельной деформации для UltimateStrain.
Private mSearchMethod As String ' Алгоритм одномерного поиска LoadMultiplier: Bisection, Brent или Secant.
Private mSolutionMethod As String ' Фактическая ветка, которая дала принятый результат: UltimateStrain или LoadMultiplier.
Private mMaxLambda As Double ' Верхняя граница расширения скобки LoadMultiplier, чтобы поиск не уходил бесконечно.
Private mInitialLambdaStep As Double ' Первый шаг расширения lambda при поиске расчетной скобки.
Private mMaxRetries As Long ' Допустимые повторы той же lambda-точки после NumericalFailure.
Private mSolverBaseLoadSteps As Long ' Базовое число ступеней CSectionSolver для первой попытки lambda-точки.
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
Private mDiagnosticLog As String ' Текстовая трассировка lambda-точек и причин остановки.
Private mDiagnosticsEnabled As Boolean ' Отключает подробный журнал в массовом пользовательском расчете без txt-отчета.
Private mProbeRunner As CStateSolutionRunner ' Общий pipeline и локальная память точных retry-attempts только текущего поиска.
Private mCriticalStrain As Double ' Деформация критического волокна или стержня.
Private mCriticalElement As String ' ID элемента, который первым определил предельное состояние.
Private mCriticalStrainUtilization As Double ' Отношение фактической деформации к своему пределу.
Private mUltimateMomentOnlyPath As Boolean ' True для частной UltimateStrain-постановки N=const, M=lambda*Mbase.
Private mProbeCacheCount As Long ' Сколько ячеек cache реально заполнено.
Private mProbeCacheNextSlot As Long ' Следующая ячейка кольцевого cache для записи.
Private mProbeCacheLambda(1 To PROBE_CACHE_SIZE) As Double ' Lambda ранее посчитанной точки поиска.
Private mProbeCacheN(1 To PROBE_CACHE_SIZE) As Double ' N этой lambda-точки; нужен, чтобы не спутать разные траектории.
Private mProbeCacheMx(1 To PROBE_CACHE_SIZE) As Double ' Mx этой lambda-точки.
Private mProbeCacheMy(1 To PROBE_CACHE_SIZE) As Double ' My этой lambda-точки.
Private mProbeCacheState(1 To PROBE_CACHE_SIZE) As String ' Классификация lambda-точки: acceptable/limit/numerical.
Private mProbeCacheCriticalStrain(1 To PROBE_CACHE_SIZE) As Double ' Критическая деформация, восстановленная без повторного solve.
Private mProbeCacheCriticalElement(1 To PROBE_CACHE_SIZE) As String ' Критический элемент, соответствующий cache-записи.
Private mProbeCacheCriticalUtil(1 To PROBE_CACHE_SIZE) As Double ' Utilization критической деформации для целевой функции.
Private mProbeCacheSolver(1 To PROBE_CACHE_SIZE) As CSectionSolver ' Solver-состояние lambda-точки для дальнейшего warm-start.

' ------------------------------
' Настройки по умолчанию
' ------------------------------
' Эти значения используются в модульных тестах и как запасной набор, если
' книга еще не передала настройки с листа Config.
Private Sub Class_Initialize()
    mConcreteCompressionLimit = -0.0035
    mConcreteTensionLimit = 0.00015
    mConcreteTensionLimitEnabled = False
    mSteelCompressionLimit = 0.025
    mSteelTensionLimit = 0.025
    mLambdaTolerance = 0.001
    mStrainTolerance = 0.00001
    mSearchMethod = "Bisection"
    mMaxLambda = 1024#
    mInitialLambdaStep = 1#
    mMaxRetries = 0
    mSolverBaseLoadSteps = 1
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
' внутренний CSectionSolver, который каждый раз ищет равновесие для lambda-точки.
' Явно пустые численные настройки не заменяются сохраненными defaults.
Public Sub ApplySettings(ByVal settings As CSystemSettingsReader, Optional ByVal units As CUnitSystem = Nothing)
    If settings Is Nothing Then Exit Sub
    mInitialLambdaStep = settings.GetDouble("Capacity.InitialLambda", mInitialLambdaStep, False)
    mMaxLambda = settings.GetDouble("Capacity.MaxLambda", mMaxLambda, False)
    mLambdaTolerance = settings.GetDouble("Capacity.ToleranceLambda", mLambdaTolerance, False)
    mStrainTolerance = settings.GetDouble("Capacity.ToleranceStrain", mStrainTolerance, False)
    If settings.HasKey("Capacity.SearchMethod") Then
        mSearchMethod = settings.GetRawString("Capacity.SearchMethod", vbNullString)
    Else
        mSearchMethod = "Bisection"
    End If
    mMaxRetries = settings.GetLong("Capacity.MaxRetries", mMaxRetries, False)
    mSolverBaseLoadSteps = settings.GetLong("Capacity.BaseLoadSteps", mSolverBaseLoadSteps, False)
    mSolverMaxIterations = settings.GetLong("Capacity.SolverMaxIterations", mSolverMaxIterations, False)
    mSolverToleranceN = settings.GetDouble("Solver.ToleranceN", mSolverToleranceN, False)
    mSolverToleranceMx = settings.GetDouble("Solver.ToleranceMx", mSolverToleranceMx, False)
    mSolverToleranceMy = settings.GetDouble("Solver.ToleranceMy", mSolverToleranceMy, False)
    If settings.HasKey("Solver.Method") Then
        mSolverMethod = settings.GetRawString("Solver.Method", vbNullString)
    Else
        mSolverMethod = "Newton"
    End If
    mSolverLineSearchEnabled = settings.GetBoolean("Solver.LineSearchEnabled", mSolverLineSearchEnabled)
    mSolverDampingInitial = settings.GetDouble("Solver.DampingInitial", mSolverDampingInitial, False)
    mSolverMinLineSearchAlpha = settings.GetDouble("Solver.MinLineSearchAlpha", mSolverMinLineSearchAlpha, False)
    mSolverMaxDeltaEpsilon0 = settings.GetDouble("Solver.MaxDeltaEpsilon0", mSolverMaxDeltaEpsilon0, False)
    mSolverMaxDeltaKappa = settings.GetDouble("Solver.MaxDeltaKappa", mSolverMaxDeltaKappa, False)
    mSolverSecantMaxRestarts = settings.GetLong("Solver.SecantMaxRestarts", mSolverSecantMaxRestarts, False)
    mSolverSecantMinStepNorm = settings.GetDouble("Solver.SecantMinStepNorm", mSolverSecantMinStepNorm, False)
    If Not units Is Nothing Then
        ' Отрицательный допуск остается ошибкой настройки и после перевода единиц.
        mSolverToleranceN = Sgn(mSolverToleranceN) * Abs(units.InputForceToInternal(mSolverToleranceN))
        mSolverToleranceMx = Sgn(mSolverToleranceMx) * Abs(units.InputMomentMxToInternal(mSolverToleranceMx))
        mSolverToleranceMy = Sgn(mSolverToleranceMy) * Abs(units.InputMomentMyToInternal(mSolverToleranceMy))
        mSolverMaxDeltaKappa = Sgn(mSolverMaxDeltaKappa) * Abs(units.InputCurvatureToInternal(mSolverMaxDeltaKappa))
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
    If Not EnsureSearchDependencies(section, concreteMaterial, steelMaterial) Then Exit Sub
    If allowPureAxialRedirect And IsPureAxialLoadPath(nOffset, nBase, mxOffset, mxBase, myOffset, myBase) Then
        SolvePureAxialUltimateByLoadMultiplier section, concreteMaterial, steelMaterial, _
            nOffset, nBase, mxOffset, mxBase, myOffset, myBase
        Exit Sub
    End If

    Dim request As CLimitSearchRequest
    Set request = New CLimitSearchRequest
    request.InitializeCapacity Me, section, concreteMaterial, steelMaterial, _
        "UltimateStrain", nOffset, nBase, mxOffset, mxBase, myOffset, myBase, _
        False, False

    Dim loadPath As CLoadPathVector
    Set loadPath = New CLoadPathVector
    loadPath.Initialize nOffset, nBase, mxOffset, mxBase, myOffset, myBase

    Dim search As CUltimateStrainSearch
    Set search = New CUltimateStrainSearch
    search.RunCapacityUltimateNewton request, loadPath
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
    If mConverged Or LimitSearchHasTerminalFailure Then Exit Sub

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

' Проверяет физическую предельную плоскость на силовой траектории.
' Удержанная проба дает лишь форму стартовой плоскости. После доведения
' до первого предела lambda определяется заново по внутренним усилиям,
' затем подтверждаются все три компоненты пути и физический критерий.
Private Sub AcceptPureAxialRetainedCapacity(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal loadPath As CLoadPathVector)
    If mSearchBoundReached Then Exit Sub
    If loadPath Is Nothing Then Exit Sub
    If mLastAcceptableSolver Is Nothing Then Exit Sub
    If Not mLastAcceptableSolver.Converged Then Exit Sub
    UpdateCriticalState section, concreteMaterial, mLastAcceptableSolver
    If mCriticalStrainUtilization <= 0# Then Exit Sub

    Dim strainScale As Double
    strainScale = 1# / mCriticalStrainUtilization
    Dim boundarySolver As CSectionSolver
    Set boundarySolver = New CSectionSolver
    boundarySolver.ToleranceN = mSolverToleranceN
    boundarySolver.ToleranceMx = mSolverToleranceMx
    boundarySolver.ToleranceMy = mSolverToleranceMy
    boundarySolver.EvaluateStrainPlane section, concreteMaterial, steelMaterial, _
        mLastAcceptableSolver.Epsilon0 * strainScale, _
        mLastAcceptableSolver.KappaX * strainScale, _
        mLastAcceptableSolver.KappaY * strainScale
    Dim math As CLoadPathMath
    Set math = NewLoadPathMath()
    Dim boundaryLambda As Double
    boundaryLambda = math.Lambda(boundarySolver.Nint, boundarySolver.Mxint, boundarySolver.Myint, loadPath)
    If boundaryLambda <= 0# Or boundaryLambda > mMaxLambda Then Exit Sub
    If Not boundarySolver.ConfirmEquilibrium(loadPath.TargetN(boundaryLambda), _
            loadPath.TargetMx(boundaryLambda), loadPath.TargetMy(boundaryLambda)) Then
        ' При несимметричном армировании масштабирование от нуля может
        ' вывести все стержни на плато и потерять моментную жесткость.
        ' Ставим нулевую деформацию на противоположном крае арматуры,
        ' сохраняя направление плоскости. Это только старт Newton,
        ' а не доказательство равновесия или несущей.
        Dim startSolver As CSectionSolver
        Set startSolver = mLastAcceptableSolver
        Dim anchorStrain As Double
        Dim edgeStrain As Double
        Dim edgeLimit As Double
        If mCriticalStrain > 0# Then
            anchorStrain = mLastAcceptableSolver.MinSteelStrain
            edgeStrain = mLastAcceptableSolver.MaxSteelStrain
            edgeLimit = mSteelTensionLimit
        Else
            anchorStrain = mLastAcceptableSolver.MaxSteelStrain
            edgeStrain = mLastAcceptableSolver.MinSteelStrain
            edgeLimit = -mSteelCompressionLimit
        End If
        If section.RebarCount > 0 And Abs(edgeStrain - anchorStrain) > 0.000000000001 Then
            strainScale = edgeLimit / (edgeStrain - anchorStrain)
            If strainScale > 1# Then
                Set startSolver = New CSectionSolver
                startSolver.EvaluateStrainPlane section, concreteMaterial, steelMaterial, _
                    (mLastAcceptableSolver.Epsilon0 - anchorStrain) * strainScale, _
                    mLastAcceptableSolver.KappaX * strainScale, _
                    mLastAcceptableSolver.KappaY * strainScale
            End If
        End If
        Dim problem As CCapacityLimitSearchProblem
        Set problem = New CCapacityLimitSearchProblem
        problem.Initialize Me
        Dim search As CUltimateStrainSearch
        Set search = New CUltimateStrainSearch
        mUltimateMomentOnlyPath = False
        If Not search.RunNewton(problem, section, concreteMaterial, steelMaterial, _
                loadPath, startSolver, _
                "Не удалось подтвердить равновесие осевого предела по выбранному пути.") Then Exit Sub
        Set boundarySolver = mLastSolver
        boundaryLambda = mLambdaUltimate
    End If
    UpdateCriticalState section, concreteMaterial, boundarySolver
    If Abs(mCriticalStrainUtilization - 1#) > UltimateStrainRatioTolerance() Then Exit Sub

    Set mLastSolver = boundarySolver
    StoreLoadPathUltimateComponents boundaryLambda, loadPath.NOffset, loadPath.NBase, _
        loadPath.MxBase, loadPath.MyBase, loadPath.MxOffset, loadPath.MyOffset
    mLimitState = CriticalLimitState()
    mUpperLimitState = mLimitState
    mNUltimate = boundarySolver.Nint
    mMxUltimate = boundarySolver.Mxint
    mMyUltimate = boundarySolver.Myint
    mMomentUltimate = Sqr(mMxUltimate * mMxUltimate + mMyUltimate * mMyUltimate)
    mSafetyFactor = mLambdaUltimate
    mUtilization = 1# / mLambdaUltimate
    mConverged = True
    mStopReason = "Осевое силовое плато подтверждено равновесием на физической предельной плоскости."
    AppendSearchDiagnostic "осевое плато подтверждено отдельной проверкой усилий на физическом пределе; lambda=" & _
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
    If mSteelCompressionLimit > 0# Then strainLimit = MinDouble(strainLimit, mSteelCompressionLimit)
    If mSteelTensionLimit > 0# Then strainLimit = MinDouble(strainLimit, mSteelTensionLimit)
    If strainLimit <= 0# Then strainLimit = 0.001

    If targetNValue < 0# Then
        ' Старт остается внутри меньшего физического предела материалов.
        ' Иначе при слабом пределе стали ограничение шага не дает вернуться
        ' из далекой бетонной деформации за доступное число итераций.
        eps0 = MinDouble(eps0, -0.25 * strainLimit)
    ElseIf targetNValue > 0# Then
        eps0 = MaxDouble(eps0, 0.25 * strainLimit)
    End If

    momentNorm = Sqr(targetMxValue * targetMxValue + targetMyValue * targetMyValue)
    If momentNorm <= 0.000000001 Then Exit Sub

    If maxAbsY > 0# Then kx = Sgn(targetMxValue) * strainLimit / maxAbsY * Abs(targetMxValue) / momentNorm
    If maxAbsX > 0# Then ky = Sgn(targetMyValue) * strainLimit / maxAbsX * Abs(targetMyValue) / momentNorm
End Sub

' Стартовая плоскость для частной моментной траектории N=const, M=lambda*Mbase.
' Она лучше обусловлена для lambda*Mx/lambda*My/lambda*Mxy, потому что
' начальная кривизна сразу направлена по масштабируемому моментному вектору.
Private Sub InitializeUltimateMomentGuess(ByVal section As CSectionModel, _
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
    If mSteelCompressionLimit > 0# Then strainLimit = MinDouble(strainLimit, mSteelCompressionLimit)
    If mSteelTensionLimit > 0# Then strainLimit = MinDouble(strainLimit, mSteelTensionLimit)
    If strainLimit <= 0# Then strainLimit = 0.001

    If maxAbsY > 0# And baseNorm > 0# Then kx = Sgn(mxBase) * strainLimit / maxAbsY * Abs(mxBase) / baseNorm
    If maxAbsX > 0# And baseNorm > 0# Then ky = Sgn(myBase) * strainLimit / maxAbsX * Abs(myBase) / baseNorm
    If Abs(kx) <= 0.000000000001 And Abs(ky) <= 0.000000000001 Then kx = strainLimit / MaxDouble(maxAbsY, 1#)
End Sub

' Оценивает частную моментную UltimateStrain-систему:
' rPath1 удерживает заданную продольную силу, rPath2 сохраняет направление
' моментного вектора, rLimit доводит критическое волокно/стержень до предела.
Private Function EvaluateUltimateMomentResidual(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal nSpecified As Double, ByVal mxBase As Double, ByVal myBase As Double, _
        ByVal mxOffset As Double, ByVal myOffset As Double, _
        ByVal eps0 As Double, ByVal kx As Double, ByVal ky As Double, _
        ByRef solver As CSectionSolver, ByRef rN As Double, ByRef rDir As Double, _
        ByRef rLimit As Double, ByRef lambdaValue As Double, ByRef state As String) As Boolean
    On Error GoTo Failed

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
    ' внутренние моменты сначала возвращаются из системы координат модели.
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
    EvaluateUltimateMomentResidual = True
    Exit Function

Failed:
    mLimitState = STATE_NUMERICAL_FAILURE
    mStopReason = "UltimateStrain: не удалось оценить моментное состояние: " & Err.Description
End Function

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
        limitAbs = SteelLimitForStrain(mCriticalStrain)
    End If
    UltimateStrainRatioTolerance = MaxDouble(0.000001, mStrainTolerance / MaxDouble(limitAbs, 0.000000001))
End Function

' ------------------------------
' Метод LoadMultiplier
' ------------------------------
' Сохраняет N постоянным, а Mx/My масштабирует одним lambda. Каждая lambda-точка
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
    If Not EnsureSearchDependencies(section, concreteMaterial, steelMaterial) Then Exit Sub
    Dim request As CLimitSearchRequest
    Set request = New CLimitSearchRequest
    request.InitializeCapacity Me, section, concreteMaterial, steelMaterial, _
        "LoadMultiplier", nOffset, nBase, mxOffset, mxBase, myOffset, myBase, _
        allowForcePathRetainedLimit, _
        IsPureAxialLoadPath(nOffset, nBase, mxOffset, mxBase, myOffset, myBase)

    Dim search As CLoadMultiplierSearch
    Set search = New CLoadMultiplierSearch
    search.Execute request
End Sub

' Готовит solver к общей ветке LoadMultiplier. Метод оставляет в CCapacitySolver
' только доменные проверки и хранение результата, а сам одномерный поиск lambda
' выполняет CLoadMultiplierSearch.
Public Function LimitSearchPrepareCapacityLoadMultiplier(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal nBase As Double, ByVal mxBase As Double, ByVal myBase As Double) As Boolean
    ClearResult
    mSolutionMethod = "LoadMultiplier"

    If Not ValidateLoadPathCapacityInputs(section, concreteMaterial, steelMaterial, nBase, mxBase, myBase) Then Exit Function
    If Not ValidateLoadMultiplierSettings() Then Exit Function
    LimitSearchPrepareCapacityLoadMultiplier = True
End Function

' Возвращает значение целевой функции для общей LoadMultiplier-точки capacity.
' Отрицательная функция означает запас до предельной деформации, ноль или плюс -
' достижение физического предела. Классификация самой lambda-точки остается здесь,
' потому что она зависит от предельных деформаций материалов strength-модели.
Public Function LimitSearchEvaluateCapacityLoadMultiplier(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal loadPath As CLoadPathVector, ByVal lambdaValue As Double, _
        ByRef functionValue As Double, ByRef probeState As String) As Boolean
    Dim nValue As Double
    Dim mxValue As Double
    Dim myValue As Double
    loadPath.Target lambdaValue, nValue, mxValue, myValue

    probeState = ProbeWithRetries(section, concreteMaterial, steelMaterial, _
        nValue, mxValue, myValue, lambdaValue)
    If probeState = STATE_ACCEPTABLE Then
        functionValue = mCriticalStrainUtilization - 1#
        RecordAcceptableState mLastSolver, nValue, mxValue, myValue
        LimitSearchEvaluateCapacityLoadMultiplier = True
    ElseIf IsPhysicalLimit(probeState) Then
        functionValue = mCriticalStrainUtilization - 1#
        If functionValue < 0# Then functionValue = 0#
        mUpperLimitState = probeState
        LimitSearchEvaluateCapacityLoadMultiplier = True
    ElseIf IsNumericalFailure(probeState) Then
        functionValue = 0#
        mLimitState = probeState
        If Len(mStopReason) = 0 Then mStopReason = "Во время поиска " & NormalizedSearchMethod() & " возникла численная несходимость."
        LimitSearchEvaluateCapacityLoadMultiplier = True
    Else
        functionValue = 0#
        mLimitState = probeState
        If Len(mStopReason) = 0 Then mStopReason = "Поиск несущей способности остановлен: " & CapacityStateDisplayName(probeState) & "."
        LimitSearchEvaluateCapacityLoadMultiplier = Not LimitSearchHasTerminalFailure
    End If
End Function

' Обрабатывает ситуацию, когда нижняя точка lambda = 0 уже не является
' приемлемой для capacity path. Здесь фиксируется только capacity-факт;
' внешний BaseFail позже формируется через ResultCode, а не через текст.
Public Sub LimitSearchHandleCapacityInitialFailure(ByVal probeState As String)
    mLimitState = probeState
    If LimitSearchHasTerminalFailure Then
        If Len(mStopReason) = 0 Then mStopReason = "Исходную точку пути несущей проверить нельзя: " & CapacityStateDisplayName(probeState) & "."
    ElseIf LimitSearchCapacityStateIsNumericalFailure(probeState) Then
        mStopReason = "Равновесие для исходной точки пути несущей не найдено численно: " & _
            CapacityStateDisplayName(probeState) & "."
    ElseIf LimitSearchCapacityStateIsPhysicalLimit(probeState) Then
        mInitialLambdaRejected = True
        mStopReason = "Несущая способность не проходит уже для исходной части выбранного пути: " & _
            CapacityStateDisplayName(probeState) & "."
    Else
        mLimitState = STATE_NUMERICAL_FAILURE
        mStopReason = "Исходная точка пути несущей вернула неизвестный статус: " & _
            CapacityStateDisplayName(probeState) & "."
    End If
End Sub

' Переводит внутренний код capacity-пробы в короткую инженерную причину для
' Results и batch summary. Машинный код остается в LimitState/ResultCode.
Private Function CapacityStateDisplayName(ByVal probeState As String) As String
    Select Case probeState
        Case STATE_CONCRETE_LIMIT
            CapacityStateDisplayName = "предельная деформация бетона уже превышена"
        Case STATE_CONCRETE_TENSION_LIMIT
            CapacityStateDisplayName = "предельная растягивающая деформация бетона уже превышена"
        Case STATE_STEEL_LIMIT
            CapacityStateDisplayName = "предельная деформация арматуры уже превышена"
        Case STATE_NUMERICAL_FAILURE
            CapacityStateDisplayName = "равновесие не сошлось численно"
        Case STATE_SINGULAR_TANGENT
            CapacityStateDisplayName = "касательная матрица вырождена"
        Case STATE_INVALID_INPUT
            CapacityStateDisplayName = "ошибка исходных данных или настроек"
        Case STATE_INTERNAL_ERROR
            CapacityStateDisplayName = "нарушен внутренний контракт расчета"
        Case Else
            CapacityStateDisplayName = probeState
    End Select
End Function

' Фиксирует, что до верхней границы MaxLambda физический предел не найден.
Public Sub LimitSearchHandleCapacityBracketNotFound(ByVal lowLambda As Double)
    mSearchBoundReached = True
    mConverged = False
    mLimitState = STATE_ACCEPTABLE
    mStopReason = "Достигнута заданная граница Capacity.MaxLambda; физический предел сечения не найден."
    mLambdaUltimate = 0#
End Sub

' Обрабатывает численную неудачу на этапе расширения верхней границы lambda.
Public Function LimitSearchReduceCapacityLambdaAfterNumericalFailure(ByVal lowLambda As Double, _
        ByVal highLambda As Double, ByVal probeState As String) As Double
    LimitSearchReduceCapacityLambdaAfterNumericalFailure = ReduceLambdaAfterNumericalFailure(lowLambda, highLambda, probeState)
End Function

' Сохраняет последний acceptable-result для чисто осевой силовой траектории,
' если физический предел не удалось отделить численно.
Public Sub LimitSearchAcceptPureAxialRetainedCapacity(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal loadPath As CLoadPathVector)
    AcceptPureAxialRetainedCapacity section, concreteMaterial, steelMaterial, loadPath
End Sub

' Сохраняет последнюю допустимую точку общего LoadMultiplier-search.
' Она нужна для финализации с нижней стороны физической скобки без повторного
' solve, а при численном отказе остается диагностикой, не готовой capacity.
Public Sub LimitSearchRecordCapacityAcceptable(ByVal nValue As Double, _
        ByVal mxValue As Double, ByVal myValue As Double)
    RecordAcceptableState mLastSolver, nValue, mxValue, myValue
End Sub

' Сохраняет численный отказ внутри физической скобки и последнюю допустимую
' lambda для диагностики. Такая точка сама по себе не становится capacity.
Public Sub LimitSearchHandleCapacityBracketNumericalFailure(ByVal lowLambda As Double, _
        ByVal probeState As String)
    mLimitState = probeState
    mStopReason = "Внутри физической скобки возникла численная несходимость; сохранена последняя сходящаяся lambda."
    mLambdaUltimate = lowLambda
    If Not mLastAcceptableSolver Is Nothing Then Set mLastSolver = mLastAcceptableSolver
End Sub

' Финализирует найденный LoadMultiplier-результат после того, как общий
' одномерный поиск выбрал итоговую lambda со стороны acceptable-ветки.
Public Sub LimitSearchFinalizeCapacityLoadMultiplier(ByVal lambdaValue As Double, _
        ByVal loadPath As CLoadPathVector, ByVal limitState As String, ByVal methodName As String, _
        ByVal solver As CSectionSolver)
    If Not solver Is Nothing Then Set mLastSolver = solver
    FinalizeLoadMultiplierSearch lambdaValue, _
        loadPath.NOffset, loadPath.NBase, loadPath.MxBase, loadPath.MyBase, _
        loadPath.MxOffset, loadPath.MyOffset, limitState, methodName
End Sub

' Финализирует Bisection по последней безопасной lambda-точке, уже
' посчитанной внутри общей скобки. Повторный solve в той же точке не нужен
' и не должен заменять подтвержденное равновесие неудачной повторной попыткой
' на жестких сценариях с малым числом итераций внутреннего CSectionSolver.
Public Sub LimitSearchFinalizeCapacityFromRetainedAcceptable(ByVal lambdaValue As Double, _
        ByVal loadPath As CLoadPathVector, ByVal methodName As String)
    If Not mLastAcceptableSolver Is Nothing Then Set mLastSolver = mLastAcceptableSolver
    FinalizeLoadMultiplierSearch lambdaValue, _
        loadPath.NOffset, loadPath.NBase, loadPath.MxBase, loadPath.MyBase, _
        loadPath.MxOffset, loadPath.MyOffset, mUpperLimitState, methodName
End Sub

' Переводит capacity-result в численную ошибку общего поиска.
Public Sub LimitSearchSetCapacityNumericalFailure(ByVal reason As String)
    mLimitState = STATE_NUMERICAL_FAILURE
    mFailureCode = sfcNumericalFailure
    mStopReason = reason
End Sub

' Переводит capacity-result в ошибку исходных данных общего поиска.
Public Sub LimitSearchSetCapacityInvalidInput(ByVal reason As String)
    SetInvalidInput reason
End Sub

' Добавляет строку в общий журнал capacity-search.
Public Sub LimitSearchAppendCapacityDiagnostic(ByVal message As String)
    AppendSearchDiagnostic message
End Sub

' Готовит capacity-контекст к общей ветке UltimateStrain. Сам Newton-цикл
' выполняет CUltimateStrainSearch, а здесь остаются проверки и хранение результата.
Public Function LimitSearchPrepareCapacityUltimate(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal loadPath As CLoadPathVector) As Boolean
    ClearResult
    mSolutionMethod = "UltimateStrain"
    If loadPath Is Nothing Then
        SetInvalidInput "Для UltimateStrain не задана lambda-траектория."
        Exit Function
    End If
    If Not ValidateLoadPathCapacityInputs(section, concreteMaterial, steelMaterial, _
            loadPath.NBase, loadPath.MxBase, loadPath.MyBase) Then Exit Function
    If Not ValidateUltimateStrainSettings() Then Exit Function
    mUltimateMomentOnlyPath = IsMomentOnlyLoadPath(loadPath.NBase, loadPath.MxBase, loadPath.MyBase)
    LimitSearchPrepareCapacityUltimate = True
End Function

' Строит стартовую плоскость для общего UltimateStrain capacity-поиска.
Public Sub LimitSearchInitializeUltimateGuess(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal loadPath As CLoadPathVector, ByVal startSolver As CSectionSolver, _
        ByRef eps0 As Double, ByRef kx As Double, ByRef ky As Double)
    If Not startSolver Is Nothing Then
        eps0 = startSolver.Epsilon0
        kx = startSolver.KappaX
        ky = startSolver.KappaY
        Exit Sub
    End If
    If mUltimateMomentOnlyPath Then
        InitializeUltimateMomentGuess section, loadPath.NOffset, loadPath.MxBase, _
            loadPath.MyBase, eps0, kx, ky
    Else
        InitializeUltimateLoadPathGuess section, loadPath.NOffset, loadPath.NBase, _
            loadPath.MxOffset, loadPath.MxBase, loadPath.MyOffset, loadPath.MyBase, _
            eps0, kx, ky
    End If
End Sub

' Возвращает три невязки общей UltimateStrain-задачи capacity.
Public Function LimitSearchEvaluateUltimateResidual(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal loadPath As CLoadPathVector, ByVal eps0 As Double, _
        ByVal kx As Double, ByVal ky As Double, ByRef solver As CSectionSolver, _
        ByRef rPath1 As Double, ByRef rPath2 As Double, ByRef rLimit As Double, _
        ByRef lambdaValue As Double, ByRef state As String) As Boolean
    If mUltimateMomentOnlyPath Then
        LimitSearchEvaluateUltimateResidual = EvaluateUltimateMomentResidual(section, _
            concreteMaterial, steelMaterial, loadPath.NOffset, loadPath.MxBase, _
            loadPath.MyBase, loadPath.MxOffset, loadPath.MyOffset, eps0, kx, ky, _
            solver, rPath1, rPath2, rLimit, lambdaValue, state)
    Else
        LimitSearchEvaluateUltimateResidual = EvaluateUltimateLoadPathResidual(section, _
            concreteMaterial, steelMaterial, loadPath.NOffset, loadPath.NBase, _
            loadPath.MxOffset, loadPath.MxBase, loadPath.MyOffset, loadPath.MyBase, _
            eps0, kx, ky, solver, rPath1, rPath2, rLimit, lambdaValue, state)
    End If
End Function

' Проверяет, что общая система UltimateStrain уже достаточно сошлась.
Public Function LimitSearchUltimateResidualReached(ByVal rPath1 As Double, _
        ByVal rPath2 As Double, ByVal rLimit As Double) As Boolean
    If mUltimateMomentOnlyPath Then
        LimitSearchUltimateResidualReached = (Abs(rPath1) <= 1# And _
            Abs(rPath2) <= MOMENT_DIRECTION_TOLERANCE And _
            Abs(rLimit) <= UltimateStrainRatioTolerance())
    Else
        LimitSearchUltimateResidualReached = (Abs(rPath1) <= 1# And _
            Abs(rPath2) <= 1# And Abs(rLimit) <= UltimateStrainRatioTolerance())
    End If
End Function

' Возвращает норму невязок для line-search общего UltimateStrain-цикла.
' Моментная постановка использует веса невязок 1, 100 и 10; остальные
' траектории нормирует CLoadPathMath по допускам capacity solver-а.
Public Function LimitSearchUltimateResidualNorm(ByVal rPath1 As Double, _
        ByVal rPath2 As Double, ByVal rLimit As Double) As Double
    If mUltimateMomentOnlyPath Then
        LimitSearchUltimateResidualNorm = Abs(rPath1) + Abs(rPath2) * 100# + Abs(rLimit) * 10#
    Else
        Dim math As CLoadPathMath
        Set math = NewLoadPathMath()
        LimitSearchUltimateResidualNorm = math.ResidualNorm(rPath1, rPath2, rLimit)
    End If
End Function

' Сохраняет точку UltimateStrain только после проверки усилий и физических
' деформаций. Допуск поиска критерия не разрешает выход за диаграмму:
' непринятая близкая плоскость уточняется следующим общим Newton-шагом.
Public Function LimitSearchFinalizeUltimateResult(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal loadPath As CLoadPathVector, _
        ByVal lambdaValue As Double, ByVal solver As CSectionSolver, _
        ByVal state As String, ByVal iterations As Long) As Boolean
    If solver Is Nothing Then Exit Function
    solver.ToleranceN = mSolverToleranceN
    solver.ToleranceMx = mSolverToleranceMx
    solver.ToleranceMy = mSolverToleranceMy
    If Not solver.ConfirmEquilibrium(loadPath.TargetN(lambdaValue), _
            loadPath.TargetMx(lambdaValue), loadPath.TargetMy(lambdaValue)) Then Exit Function
    Dim stateRunner As CStateSolutionRunner
    Set stateRunner = New CStateSolutionRunner
    If Not stateRunner.StateWithinPhysicalRange(section, solver, concreteMaterial, steelMaterial) Then
        AppendSearchDiagnostic "UltimateStrain: равновесие подтверждено, но деформации еще выходят за физические пределы; продолжаем уточнение предельной точки."
        Exit Function
    End If
    Set mLastSolver = solver
    If Not IsValidUltimateLambda(lambdaValue, loadPath.NOffset, loadPath.NBase, _
            loadPath.MxBase, loadPath.MyBase, loadPath.MxOffset, loadPath.MyOffset) Then
        mLimitState = STATE_NUMERICAL_FAILURE
        mStopReason = "UltimateStrain: найден недопустимый lambda: " & FormatNumberInvariant(lambdaValue)
        Exit Function
    End If
    StoreLoadPathUltimateComponents lambdaValue, loadPath.NOffset, loadPath.NBase, _
        loadPath.MxBase, loadPath.MyBase, loadPath.MxOffset, loadPath.MyOffset
    mLimitState = state
    mUpperLimitState = state
    mConverged = True
    LimitSearchFinalizeUltimateResult = True
End Function

' Завершает capacity-ветку UltimateStrain после общего цикла.
Public Sub LimitSearchCompleteCapacityUltimate(ByVal loadPath As CLoadPathVector)
    If Not mConverged Then
        If Len(mLimitState) = 0 Then mLimitState = STATE_NUMERICAL_FAILURE
        If Len(mStopReason) = 0 Then mStopReason = "UltimateStrain: итерации Newton не сошлись."
        Exit Sub
    End If

    If Not mLastSolver Is Nothing Then
        mMomentEquilibriumResidual = Sqr((mLastSolver.Mxint - loadPath.TargetMx(mLambdaUltimate)) ^ 2 + _
            (mLastSolver.Myint - loadPath.TargetMy(mLambdaUltimate)) ^ 2)
    End If
    mSafetyFactor = mLambdaUltimate
    If mLambdaUltimate > 0# Then mUtilization = 1# / mLambdaUltimate
    mStopReason = "Поиск несущей способности через UltimateStrain сошелся."
End Sub

' Переводит результат UltimateStrain в численную остановку.
Public Sub LimitSearchSetUltimateNumericalFailure(ByVal reason As String)
    If LimitSearchHasTerminalFailure Then Exit Sub
    If mFailureCode = sfcSingularTangent Then
        AppendSearchDiagnostic reason
        Exit Sub
    End If
    LimitSearchSetFailure sfcNumericalFailure, reason
End Sub

' Сохраняет typed-причину из общего Search. Терминальная конфигурационная
' ошибка не повторяется и не переинтерпретируется как физический отказ пути.
Public Sub LimitSearchSetFailure(ByVal failureCode As ESolverFailureCode, ByVal reason As String)
    mConverged = False
    Select Case failureCode
        Case sfcInvalidInput
            SetInvalidInput reason
        Case sfcInvalidConfiguration
            SetInvalidConfiguration reason
        Case sfcNumericalFailure, sfcSingularTangent
            mFailureCode = failureCode
            mLimitState = ClassifyNumericalFailure(failureCode)
            mStopReason = reason
        Case Else
            SetInternalError reason
    End Select
End Sub

' Настройки общего UltimateStrain-цикла.
Public Property Get LimitSearchUltimateMaxIterations() As Long
    LimitSearchUltimateMaxIterations = mSolverMaxIterations
End Property

Public Property Get LimitSearchUltimateDampingInitial() As Double
    LimitSearchUltimateDampingInitial = mSolverDampingInitial
End Property

Public Property Get LimitSearchUltimateMinLineSearchAlpha() As Double
    LimitSearchUltimateMinLineSearchAlpha = mSolverMinLineSearchAlpha
End Property

Public Property Get LimitSearchUltimateLineSearchEnabled() As Boolean
    LimitSearchUltimateLineSearchEnabled = mSolverLineSearchEnabled
End Property

Public Property Get LimitSearchUltimateMaxDeltaEpsilon0() As Double
    LimitSearchUltimateMaxDeltaEpsilon0 = mSolverMaxDeltaEpsilon0
End Property

Public Property Get LimitSearchUltimateMaxDeltaKappa() As Double
    LimitSearchUltimateMaxDeltaKappa = mSolverMaxDeltaKappa
End Property

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
    mStopReason = "Поиск несущей способности методом " & methodName & _
        " сошелся; верхняя граница подтверждена: " & CapacityStateDisplayName(mUpperLimitState) & "."
End Sub

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

' Упаковывает компонентную форму CCapacitySolver в общий CLoadPathVector, чтобы
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

' Проверяет программные зависимости до создания обязательного Search-request.
' При нарушении контракта очищает старый результат и сохраняет typed internal
' failure, чтобы публичный расчет не оставлял необработанную ошибку VBA.
Private Function EnsureSearchDependencies(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object) As Boolean
    If Not section Is Nothing Then
        If Not concreteMaterial Is Nothing Then
            If Not steelMaterial Is Nothing Then
                EnsureSearchDependencies = True
                Exit Function
            End If
        End If
    End If
    ClearResult
    If section Is Nothing Then
        SetInternalError "Модель сечения не передана расчетчику несущей способности."
    ElseIf concreteMaterial Is Nothing Then
        SetInternalError "Материал бетона не передан расчетчику несущей способности."
    Else
        SetInternalError "Материал арматуры не передан расчетчику несущей способности."
    End If
End Function

' Проверяет общую lambda-траекторию. Нулевой Base означает, что ни одна
' компонента нагрузки не масштабируется, поэтому предельный множитель не имеет
' инженерного смысла.
Private Function ValidateLoadPathCapacityInputs(ByVal section As CSectionModel, ByVal concreteMaterial As Object, _
        ByVal steelMaterial As Object, ByVal nBase As Double, ByVal mxBase As Double, ByVal myBase As Double) As Boolean
    If section Is Nothing Then
        SetInternalError "Модель сечения не передана расчетчику несущей способности."
    ElseIf section.ConcreteCount <= 0 Then
        SetInvalidInput "В модели сечения нет бетонных элементов."
    ElseIf concreteMaterial Is Nothing Then
        SetInternalError "Материал бетона не передан расчетчику несущей способности."
    ElseIf steelMaterial Is Nothing Then
        SetInternalError "Материал арматуры не передан расчетчику несущей способности."
    ElseIf Abs(nBase) <= 0.000000001 And Abs(mxBase) <= 0.000000001 And Abs(myBase) <= 0.000000001 Then
        SetInvalidInput "В траектории несущей способности нет масштабируемой нагрузки."
    ElseIf mSolverMaxIterations <= 0 Then
        SetInvalidConfiguration "Число итераций решателя несущей способности должно быть положительным."
    Else
        ValidateLoadPathCapacityInputs = True
    End If
End Function

' Проверяет модель, материалы и ненулевой масштабируемый вектор момента
' для поиска при фиксированной продольной силе. Отказ сохраняется typed:
' ошибка модели/ввода/настройки не подменяется численной несходимостью.
Private Function ValidateCapacityInputs(ByVal section As CSectionModel, ByVal concreteMaterial As Object, _
        ByVal steelMaterial As Object, ByVal mxBase As Double, ByVal myBase As Double) As Boolean
    If section Is Nothing Then
        SetInternalError "Модель сечения не передана расчетчику несущей способности."
    ElseIf section.ConcreteCount <= 0 Then
        SetInvalidInput "В модели сечения нет бетонных элементов."
    ElseIf concreteMaterial Is Nothing Then
        SetInternalError "Материал бетона не передан расчетчику несущей способности."
    ElseIf steelMaterial Is Nothing Then
        SetInternalError "Материал арматуры не передан расчетчику несущей способности."
    ElseIf Sqr(mxBase * mxBase + myBase * myBase) <= 0.000000001 Then
        SetInvalidInput "Базовый вектор момента должен быть ненулевым."
    ElseIf mSolverMaxIterations <= 0 Then
        SetInvalidConfiguration "Число итераций решателя несущей способности должно быть положительным."
    Else
        ValidateCapacityInputs = True
    End If
End Function

' Проверяет настройки ветки LoadMultiplier перед одномерным поиском lambda.
' Здесь контролируются только параметры самой скобки/поиска, а не геометрия
' или материалы, которые проверяются выше по цепочке.
Private Function ValidateLoadMultiplierSettings() As Boolean
    If Len(NormalizedSearchMethod()) = 0 Then
        SetInvalidConfiguration "Capacity.SearchMethod должен быть Bisection, Brent или Secant."
    ElseIf StrComp(Trim$(mSolverMethod), "Newton", vbTextCompare) <> 0 And _
            StrComp(Trim$(mSolverMethod), "Secant", vbTextCompare) <> 0 Then
        SetInvalidConfiguration "Solver.Method должен быть Newton или Secant."
    ElseIf mInitialLambdaStep <= 0# Or mLambdaTolerance <= 0# Or mMaxLambda <= 0# Or _
            mMaxRetries < 0 Or mSolverBaseLoadSteps <= 0 Then
        SetInvalidConfiguration "Настройки LoadMultiplier для несущей способности некорректны."
    Else
        ValidateLoadMultiplierSettings = True
    End If
End Function

' Проверяет настройки ветки UltimateStrain. В отличие от LoadMultiplier,
' здесь нет искусственного Capacity.MaxLambda: после найденной предельной
' плоскости достаточно проверить численную корректность lambda.
Private Function ValidateUltimateStrainSettings() As Boolean
    If mStrainTolerance <= 0# Then
        SetInvalidConfiguration "Настройки UltimateStrain для несущей способности некорректны."
    Else
        ValidateUltimateStrainSettings = True
    End If
End Function

' Переводит расчет в InputErr без выбрасывания VBA-ошибки. Capacity solver
' возвращает статус пользователю через batch summary и подробный writer.
Private Sub SetInvalidInput(ByVal reason As String)
    mLimitState = STATE_INVALID_INPUT
    mFailureCode = sfcInvalidInput
    mStopReason = reason
End Sub

' Ошибка численных настроек сохраняется отдельным typed-кодом и не запускает
' retries/fallback. Короткий код LimitState не является внешним display-статусом.
Private Sub SetInvalidConfiguration(ByVal reason As String)
    mLimitState = STATE_INVALID_INPUT
    mFailureCode = sfcInvalidConfiguration
    mStopReason = reason
End Sub

' Отсутствующая программная зависимость или неожиданная ошибка относятся
' к внутреннему контракту, а не к исходным данным пользователя.
Private Sub SetInternalError(ByVal reason As String)
    mLimitState = STATE_INTERNAL_ERROR
    mFailureCode = sfcInternalError
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

' Выполняет lambda-точку LoadMultiplier с ограниченным числом повторов.
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

' Решает одну lambda-точку capacity-поиска и классифицирует ее как приемлемую,
' физический предел, численную несходимость или ошибку входных данных.
Private Function ProbeOnce(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal nSpecified As Double, ByVal mx As Double, ByVal my As Double, _
        ByVal lambdaValue As Double, ByVal attempt As Long) As String

    On Error GoTo Failed
    mIterations = mIterations + 1

    If IsZeroLoadProbe(nSpecified, mx, my) Then
        ProbeOnce = EvaluateZeroLoadProbe(section, concreteMaterial, steelMaterial, lambdaValue, attempt)
        Exit Function
    End If

    Dim runner As CStateSolutionRunner
    If mProbeRunner Is Nothing Then
        Set mProbeRunner = New CStateSolutionRunner
        mProbeRunner.BeginRetryAttemptSession section, concreteMaterial, steelMaterial
    End If
    Set runner = mProbeRunner
    ConfigureCapacityProbeRunner runner, attempt
    Dim allowExtensionWarmStart As Boolean
    allowExtensionWarmStart = runner.MaterialsHaveExtension(concreteMaterial, steelMaterial)

    ' Для обычной probe-точки используем общий state-solve pipeline. Если есть
    ' предыдущая acceptable lambda, передаем ее как warm-start с явной стартовой
    ' нагрузкой. Для чистого изгиба без N используем отдельный вход runner-а:
    ' active-set плоскость сразу привязана к целевой lambda-точке и не проходит
    ' лишние ступени от нулевой нагрузки через излом диаграммы.
    If IsPureBendingProbe(nSpecified, mx, my) Then
        runner.SolveWithDirectTargetGuess section, concreteMaterial, steelMaterial, _
            nSpecified, mx, my, _
            "для lambda-точки чистого изгиба применена стартовая плоскость; масштабируемые моменты ненулевые", _
            allowExtensionWarmStart
    ElseIf Not mLastAcceptableSolver Is Nothing Then
        runner.SolveWithInitialLoadState section, concreteMaterial, steelMaterial, _
            nSpecified, mx, my, mLastAcceptableSolver.Epsilon0, _
            mLastAcceptableSolver.KappaX, mLastAcceptableSolver.KappaY, _
            mLastAcceptableN, mLastAcceptableMx, mLastAcceptableMy, allowExtensionWarmStart
    Else
        runner.Solve section, concreteMaterial, steelMaterial, nSpecified, mx, my, allowExtensionWarmStart
    End If

    Dim solver As CSectionSolver
    Set solver = runner.ResultSolver
    Set mLastSolver = solver
    If solver Is Nothing Then
        ProbeOnce = STATE_INTERNAL_ERROR
        SetInternalError "Пробная точка несущей способности не вернула результат решателя."
        Exit Function
    End If
    If Len(runner.DiagnosticLog) > 0 Then AppendSearchDiagnostic Trim$(runner.DiagnosticLog)

    If Not solver.Converged Then
        ProbeOnce = ClassifyNumericalFailure(solver.FailureCode)
        mFailureCode = solver.FailureCode
        mStopReason = solver.StopReason
    Else
        mFailureCode = sfcNone
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
    ProbeOnce = STATE_INTERNAL_ERROR
    SetInternalError "Пробная точка несущей способности не рассчиталась: " & Err.Description
End Function

' Передает настройки внутреннего CSectionSolver в общий runner для одной
' capacity-probe. Retry попытки меняют только численные параметры solve-а,
' физическая lambda-точка остается прежней.
Private Sub ConfigureCapacityProbeRunner(ByVal runner As CStateSolutionRunner, ByVal attempt As Long)
    runner.LoadSteps = mSolverBaseLoadSteps * (2 ^ attempt)
    runner.SolverMethod = mSolverMethod
    runner.MaxIterations = mSolverMaxIterations
    runner.ToleranceN = mSolverToleranceN
    runner.ToleranceMx = mSolverToleranceMx
    runner.ToleranceMy = mSolverToleranceMy
    runner.LineSearchEnabled = mSolverLineSearchEnabled
    runner.DampingInitial = mSolverDampingInitial
    runner.MinLineSearchAlpha = mSolverMinLineSearchAlpha
    runner.MaxDeltaEpsilon0 = mSolverMaxDeltaEpsilon0
    runner.MaxDeltaKappa = mSolverMaxDeltaKappa / (2 ^ attempt)
    runner.SecantMaxRestarts = mSolverSecantMaxRestarts
    runner.SecantMinStepNorm = mSolverSecantMinStepNorm
    runner.DiagnosticsEnabled = mDiagnosticsEnabled
End Sub

' True для probe чистого изгиба. В этом случае предыдущая acceptable lambda
' не используется как warm-start: общий runner сам строит active-set плоскость
' по целевым моментам текущей lambda-точки.
Private Function IsPureBendingProbe(ByVal nSpecified As Double, _
        ByVal mx As Double, ByVal my As Double) As Boolean
    IsPureBendingProbe = (Abs(nSpecified) <= 0.000000001 And _
        Sqr(mx * mx + my * my) > 0.000000001)
End Function

' Проверяет численно нулевой вектор пробной нагрузки. Только такой случай
' допускает проверку нулевой плоскости вместо итерационного подбора НДС.
Private Function IsZeroLoadProbe(ByVal nSpecified As Double, _
        ByVal mx As Double, ByVal my As Double) As Boolean
    IsZeroLoadProbe = (Abs(nSpecified) <= 0.000000001 And _
        Abs(mx) <= 0.000000001 And Abs(my) <= 0.000000001)
End Function

' Проверяет точное равновесие нулевой плоскости и сохраняет физические поля
' lambda=0. Ошибки модели/настроек сохраняют typed-причину, а не становятся
' автоматически допустимым начальным состоянием.
Private Function EvaluateZeroLoadProbe(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal lambdaValue As Double, ByVal attempt As Long) As String
    On Error GoTo Failed

    Dim solver As CSectionSolver
    Set solver = New CSectionSolver
    solver.SolverMethod = mSolverMethod
    solver.MaxIterations = mSolverMaxIterations
    solver.LoadSteps = mSolverBaseLoadSteps
    solver.ToleranceN = mSolverToleranceN
    solver.ToleranceMx = mSolverToleranceMx
    solver.ToleranceMy = mSolverToleranceMy
    solver.DiagnosticsEnabled = mDiagnosticsEnabled
    solver.EvaluateStrainPlane section, concreteMaterial, steelMaterial, 0#, 0#, 0#
    Set mLastSolver = solver
    If solver.FailureCode <> sfcNone Then
        mFailureCode = solver.FailureCode
        mStopReason = solver.StopReason
        EvaluateZeroLoadProbe = ClassifyNumericalFailure(mFailureCode)
        Exit Function
    End If
    If Not solver.ConfirmEquilibrium(0#, 0#, 0#) Then
        mFailureCode = solver.FailureCode
        mStopReason = solver.StopReason
        EvaluateZeroLoadProbe = STATE_NUMERICAL_FAILURE
        Exit Function
    End If

    Set mLastSolver = solver
    UpdateCriticalState section, concreteMaterial, solver
    EvaluateZeroLoadProbe = STATE_ACCEPTABLE
    If mDiagnosticsEnabled Then _
        AppendSearchDiagnostic "lambda=0: нулевая нагрузка принята как точное равновесие нулевой плоскости деформаций."
    AppendDiagnostic solver, 0#, 0#, lambdaValue, attempt, EvaluateZeroLoadProbe
    Exit Function

Failed:
    EvaluateZeroLoadProbe = STATE_INTERNAL_ERROR
    SetInternalError "Нулевую lambda-точку несущей способности проверить нельзя: " & Err.Description
End Function

' Возвращает рассчитанную probe-точку этой search-session по точному ключу
' lambda и усилий. Восстанавливает критерий и warm-start без повторного solve;
' локальный probe-cache не публикуется в shared StateRepository.
Private Function TryRestoreProbeCache(ByVal nSpecified As Double, ByVal mx As Double, ByVal my As Double, _
        ByVal lambdaValue As Double, ByRef probeState As String) As Boolean
    Dim slot As Long
    For slot = 1 To mProbeCacheCount
        If ProbeCacheMatches(slot, nSpecified, mx, my, lambdaValue) Then
            probeState = mProbeCacheState(slot)
            Set mLastSolver = mProbeCacheSolver(slot)
            mFailureCode = sfcNone
            mCriticalStrain = mProbeCacheCriticalStrain(slot)
            mCriticalElement = mProbeCacheCriticalElement(slot)
            mCriticalStrainUtilization = mProbeCacheCriticalUtil(slot)
            AppendSearchDiagnostic "использована ранее рассчитанная lambda-точка; lambda=" & FormatNumberInvariant(lambdaValue) & "; состояние=" & probeState
            TryRestoreProbeCache = True
            Exit Function
        End If
    Next slot
End Function

' Сравнивает lambda-точку по lambda и целевым усилиям.
' Разные представимые lambda нельзя объединять произвольным допуском: поиск
' может уточнять физический критерий точнее этого допуска. Повтор той же точки
' сохраняет точное совпадение lambda и всех трех целевых усилий.
Private Function ProbeCacheMatches(ByVal slot As Long, ByVal nSpecified As Double, ByVal mx As Double, ByVal my As Double, _
        ByVal lambdaValue As Double) As Boolean
    If mProbeCacheLambda(slot) <> lambdaValue Then Exit Function
    If mProbeCacheN(slot) <> nSpecified Then Exit Function
    If mProbeCacheMx(slot) <> mx Then Exit Function
    If mProbeCacheMy(slot) <> my Then Exit Function
    ProbeCacheMatches = True
End Function

' Запоминает только найденное равновесие lambda-точки после всех повторов.
' Неудачная попытка не закрывает возможность повторного solve с лучшим стартом.
Private Sub StoreProbeCache(ByVal nSpecified As Double, ByVal mx As Double, ByVal my As Double, _
        ByVal lambdaValue As Double, ByVal probeState As String)
    If probeState <> STATE_ACCEPTABLE Then
        If Not IsPhysicalLimit(probeState) Then Exit Sub
    End If
    If mLastSolver Is Nothing Then Exit Sub
    If Not mLastSolver.Converged Then Exit Sub
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

' Обновляет управляющую деформацию для последней lambda-точки. Проверяются
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
            If SteelLimitForStrain(strain) > 0# Then
                utilization = Abs(strain) / SteelLimitForStrain(strain)
                StoreCriticalIfLarger utilization, strain, section.RebarID(i) & ":Steel"
            End If
        Next i
    End If
End Sub

' Возвращает предел арматуры для знака проверяемой деформации.
' При симметричном тестовом контракте оба направления имеют одинаковый предел.
Private Function SteelLimitForStrain(ByVal strain As Double) As Double
    If strain < 0# Then
        SteelLimitForStrain = mSteelCompressionLimit
    Else
        SteelLimitForStrain = mSteelTensionLimit
    End If
End Function

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

' Разделяет виды неуспешного CSectionSolver по машинному коду, а не по тексту
' StopReason. Текст остается только пользовательской диагностикой и не влияет
' на статусную цепочку.
Private Function ClassifyNumericalFailure(ByVal failureCode As ESolverFailureCode) As String
    Select Case failureCode
        Case sfcSingularTangent
            ClassifyNumericalFailure = STATE_SINGULAR_TANGENT
        Case sfcInvalidInput
            ClassifyNumericalFailure = STATE_INVALID_INPUT
        Case sfcInvalidConfiguration
            ClassifyNumericalFailure = STATE_INVALID_INPUT
        Case sfcNumericalFailure
            ClassifyNumericalFailure = STATE_NUMERICAL_FAILURE
        Case Else
            ClassifyNumericalFailure = STATE_INTERNAL_ERROR
    End Select
End Function

' Запоминает последнюю сходящуюся lambda-точку. Она используется как warm-start
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

' Записывает диагностическую строку одной lambda-точки LoadMultiplier.
' Поле attempt - это номер попытки НДС для той же физической lambda-точки:
' 0 означает первую попытку, 1 - первый retry с более мягкими настройками.
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
    mInitialLambdaRejected = False
    mFailureCode = sfcNone
    mSearchBoundReached = False
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
    Set mProbeRunner = Nothing
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

' Проверяет, что состояние lambda-точки еще находится внутри физически
' допустимой области capacity.
Public Function LimitSearchCapacityStateIsAcceptable(ByVal state As String) As Boolean
    LimitSearchCapacityStateIsAcceptable = (state = STATE_ACCEPTABLE)
End Function

' Проверяет, что состояние lambda-точки является физическим пределом, а не
' численной неудачей или ошибкой исходных данных.
Public Function LimitSearchCapacityStateIsPhysicalLimit(ByVal state As String) As Boolean
    LimitSearchCapacityStateIsPhysicalLimit = IsPhysicalLimit(state)
End Function

' Проверяет, что состояние lambda-точки является численной неудачей, которую
' общий search может обработать уменьшением lambda или fallback-веткой.
Public Function LimitSearchCapacityStateIsNumericalFailure(ByVal state As String) As Boolean
    LimitSearchCapacityStateIsNumericalFailure = IsNumericalFailure(state)
End Function

' Проверяет, что доменный callback остановился из-за ошибки исходных данных.
' Общий coordinator использует это, чтобы не запускать fallback LoadMultiplier
' после заведомо некорректной постановки задачи.
Public Function LimitSearchCapacityStateIsInvalidInput(ByVal state As String) As Boolean
    LimitSearchCapacityStateIsInvalidInput = (state = STATE_INVALID_INPUT)
End Function

' Терминальная ошибка данных, настройки или программной зависимости запрещает
' повтор solve и переход к другому численному методу той же постановки.
Public Property Get LimitSearchHasTerminalFailure() As Boolean
    LimitSearchHasTerminalFailure = (mFailureCode = sfcInvalidInput Or _
        mFailureCode = sfcInvalidConfiguration Or mFailureCode = sfcInternalError)
End Property

' Передает адаптеру точную typed-причину для InternalStatus/ResultCode.
' Значение не определяется по словам StopReason или пользовательскому display.
Public Property Get FailureCode() As ESolverFailureCode
    FailureCode = mFailureCode
End Property

' Возвращает True, если capacity-path был остановлен уже на lambda=0.
' Низкоуровневый search только сообщает о rejected lower point; смысл
' "базовая часть пути уже не проходит" назначается capacity-result слоем.
Public Property Get LimitSearchInitialLambdaRejected() As Boolean
    LimitSearchInitialLambdaRejected = mInitialLambdaRejected
End Property

' Отличает техническую остановку MaxLambda от несходимости пробного равновесия.
Public Property Get SearchBoundReached() As Boolean
    SearchBoundReached = mSearchBoundReached
End Property

' Объединяет диагностику неудачной первой ветки с последующим fallback-расчетом
' LoadMultiplier. Search-слой сам выбирает порядок попыток, но только
' CCapacitySolver имеет право менять свой внутренний журнал и StopReason.
Public Sub LimitSearchMergeCapacityFallback(ByVal firstLimitState As String, _
        ByVal firstStopReason As String, ByVal firstDiagnostics As String, _
        ByVal firstAttemptMessage As String, ByVal transitionMessage As String, _
        ByVal successMessage As String)
    If mDiagnosticsEnabled Then
        Dim mergedDiagnostics As String
        mergedDiagnostics = firstAttemptMessage & "; статус=" & firstLimitState & _
            "; причина=" & firstStopReason & vbCrLf
        If Len(firstDiagnostics) > 0 Then mergedDiagnostics = mergedDiagnostics & firstDiagnostics
        mergedDiagnostics = mergedDiagnostics & transitionMessage & vbCrLf & mDiagnosticLog
        mDiagnosticLog = mergedDiagnostics
    End If

    If mConverged And Len(Trim$(successMessage)) > 0 Then _
        mStopReason = successMessage & " " & mStopReason
End Sub

' Возвращает нормализованное имя одномерного метода LoadMultiplier.
Public Function LimitSearchCapacitySearchMethod() As String
    LimitSearchCapacitySearchMethod = NormalizedSearchMethod()
End Function

' Возвращает первый шаг расширения lambda-скобки.
Public Property Get LimitSearchInitialLambdaStep() As Double
    LimitSearchInitialLambdaStep = mInitialLambdaStep
End Property

' Возвращает верхнюю границу расширения lambda-скобки.
Public Property Get LimitSearchMaxLambda() As Double
    LimitSearchMaxLambda = mMaxLambda
End Property

' Возвращает допуск одномерного поиска lambda.
Public Property Get LimitSearchLambdaTolerance() As Double
    LimitSearchLambdaTolerance = mLambdaTolerance
End Property

' Возвращает максимум итераций одномерного поиска.
Public Property Get LimitSearchMaxIterations() As Long
    LimitSearchMaxIterations = mSolverMaxIterations
End Property

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
    mSteelCompressionLimit = Abs(value)
    mSteelTensionLimit = Abs(value)
End Property

' Задает независимые физические пределы арматуры по двум знакам деформации.
Public Property Let SteelCompressionLimit(ByVal value As Double)
    mSteelCompressionLimit = Abs(value)
End Property

Public Property Let SteelTensionLimit(ByVal value As Double)
    mSteelTensionLimit = Abs(value)
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
