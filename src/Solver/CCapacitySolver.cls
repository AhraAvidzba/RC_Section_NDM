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
' Готовит критерий прочности, начальные поля деформаций и итоговый результат.
' Численные циклы выполняют CLoadMultiplierSearch и CUltimateStrainSearch;
' обычное равновесие пробных нагрузок решает CStateSolutionRunner.
' Нагрузка задаётся общей зависимостью: N = N0 + lambda*Nbase,
' Mx = Mx0 + lambda*Mxbase, My = My0 + lambda*Mybase.
' Промежуточные пробы не выдаются как конечные расчётные состояния.

Private Const STATE_ACCEPTABLE As String = "Acceptable" ' Код состояния, в котором предельные деформации не превышены.
Private Const STATE_CONCRETE_LIMIT As String = "ConcreteStrainLimit" ' Код достижения предельной деформации сжатого бетона.
Private Const STATE_CONCRETE_TENSION_LIMIT As String = "ConcreteTensionStrainLimit" ' Код достижения предельной деформации растянутого бетона.
Private Const STATE_STEEL_LIMIT As String = "SteelStrainLimit" ' Код достижения предельной деформации арматуры.
Private Const STATE_NUMERICAL_FAILURE As String = "NumericalFailure" ' Код численной неудачи поиска, не означающий разрушения сечения.
Private Const STATE_SINGULAR_TANGENT As String = "SingularTangent" ' Код вырожденной касательной матрицы.
Private Const STATE_INVALID_INPUT As String = "InvalidInput" ' Код ошибки исходных данных.
Private Const STATE_INTERNAL_ERROR As String = "InternalError" ' Обязательные программные данные не переданы или несовместимы; повторное решение НДС причину не исправит.
Private Const MOMENT_DIRECTION_TOLERANCE As Double = 0.001 ' Допустимое относительное отклонение от выбранной моментной траектории.
Private Const PROBE_CACHE_SIZE As Long = 12 ' Небольшое кольцо последних lambda-точек LoadMultiplier.

Private mConverged As Boolean ' Признак того, что выбранным методом найдено подтверждённое предельное состояние.
Private mStopReason As String ' Причина остановки поиска несущей способности для вывода и тестов.
Private mLimitState As String ' Физический или численный статус найденного предела.
Private mFailureCode As ESolverFailureCode ' Точная техническая причина, независимая от текста комментария и LimitState.
Private mFailureSettingKeys As String ' Ключи ошибочных исходных настроек, передаваемые с результатом для вывода фактических адресов.
Private mUpperLimitState As String ' Статус верхней границы при LoadMultiplier.
Private mInitialLambdaRejected As Boolean ' Признак того, что постоянная нагрузка при lambda = 0 уже не удовлетворяет критерию несущей способности.
Private mUltimateOriginX As Double, mUltimateOriginY As Double ' Координаты начала отсчёта моментов при поиске предела по деформациям, мм.
Private mSearchBoundReached As Boolean ' MaxLambda проверена, но физический предел не установлен.
Private mLastProbeNumericalFailure As Boolean ' Отказ равновесия, не исчерпание допуска/бюджета одномерного поиска.
Private mLambdaUltimate As Double ' Предельный множитель к выбранной lambda-траектории нагрузки.
Private mNUltimate As Double ' Предельная продольная сила, если N входит в масштабируемый путь.
Private mMomentUltimate As Double ' Результирующий предельный момент sqrt(MxUltimate^2 + MyUltimate^2).
Private mMxUltimate As Double ' Предельный Mx, если Mx входит в масштабируемый путь.
Private mMyUltimate As Double ' Предельный My, если My входит в масштабируемый путь.
Private mSafetyFactor As Double ' Коэффициент запаса несущей способности; сейчас равен найденному lambdaUltimate.
Private mUtilization As Double ' Использование несущей способности: 1 / lambdaUltimate.
Private mIterations As Long ' Число внешних итераций поиска несущей способности.
Private mSearchExecuted As Boolean ' Признак фактического запуска численного поиска; предварительная проверка ввода запуском не считается.
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
Private mSolverMaxIterations As Long ' Максимальное число внешних итераций поиска предела и внутренних итераций равновесия.
Private mSolverToleranceN As Double ' Допуск внутреннего равновесия CSectionSolver по N.
Private mSolverToleranceMx As Double ' Допуск внутреннего равновесия CSectionSolver по Mx.
Private mSolverToleranceMy As Double ' Допуск внутреннего равновесия CSectionSolver по My.
Private mSolverMethod As String ' Метод внутреннего CSectionSolver: Newton или Secant.
Private mSolverSecantMaxRestarts As Long ' Настройка рестартов внутреннего secant-решателя.
Private mMomentEquilibriumResidual As Double ' Невязка направления/момента в найденном предельном состоянии.
Private mDiagnosticLog As String ' Журнал проверенных значений lambda и причин остановки поиска.
Private mDiagnosticLines() As String, mDiagnosticCount As Long, mDiagnosticCapacity As Long ' Строки журнала, число записанных строк и вместимость массива.
Private mDiagnosticJoined As String, mDiagnosticDirty As Boolean ' Собранный текст журнала и признак появления новых строк после его сборки.
Private mDiagnosticsEnabled As Boolean ' Отключает подробный журнал в массовом пользовательском расчете без txt-отчета.
Private mProbeRunner As CStateSolutionRunner ' Общее решение НДС и сохранённые итоги точно совпадающих повторных попыток в пределах текущего поиска.
Private mCriticalStrain As Double ' Деформация критического волокна или стержня.
Private mCriticalElement As String ' ID элемента, который первым определил предельное состояние.
Private mCriticalLimitType As String ' Тип ограничения из общего деформационного критерия.
Private mCriticalStrainUtilization As Double ' Отношение фактической деформации к своему пределу.
Private mUltimateMomentOnlyPath As Boolean ' True для частной UltimateStrain-постановки N=const, M=lambda*Mbase.
Private mMeritJacobianForAffinePath As Boolean ' Масштабы шага для моментной системы и траекторий с постоянной частью.
Private mUltimateMomentNMeritWeight As Double ' Перевод осевой невязки из долей допуска в относительную силовую ошибку.
Private mUltimatePathMath As CLoadPathMath ' Общие уравнения траектории нагрузки для поиска по предельным деформациям.
Private mUltimatePathTolerance As Double ' Допустимое отклонение от траектории в нормированных уравнениях поиска.
Private mProbeCacheCount As Long ' Число сохранённых оценок пробных нагрузок.
Private mProbeCacheNextSlot As Long ' Место следующей записи; при заполнении списка новая оценка заменяет самую старую.
Private mProbeCacheLambda(1 To PROBE_CACHE_SIZE) As Double ' Lambda ранее посчитанной точки поиска.
Private mProbeCacheN(1 To PROBE_CACHE_SIZE) As Double ' N этой lambda-точки; нужен, чтобы не спутать разные траектории.
Private mProbeCacheMx(1 To PROBE_CACHE_SIZE) As Double ' Mx этой lambda-точки.
Private mProbeCacheMy(1 To PROBE_CACHE_SIZE) As Double ' My этой lambda-точки.
Private mProbeCacheState(1 To PROBE_CACHE_SIZE) As String ' Классификация lambda-точки: acceptable/limit/numerical.
Private mProbeCacheCriticalStrain(1 To PROBE_CACHE_SIZE) As Double ' Критическая деформация, восстановленная без повторного solve.
Private mProbeCacheCriticalElement(1 To PROBE_CACHE_SIZE) As String ' Критический элемент для сохранённой оценки нагрузки.
Private mProbeCacheCriticalType(1 To PROBE_CACHE_SIZE) As String ' Вид определяющего предела для каждой сохранённой пробной точки.
Private mProbeCacheCriticalUtil(1 To PROBE_CACHE_SIZE) As Double ' Коэффициент использования предельной деформации для поиска границы допустимых нагрузок.
Private mProbeCacheSolver(1 To PROBE_CACHE_SIZE) As CSectionSolver ' НДС пробной нагрузки для начального приближения следующего решения.

' ------------------------------
' Настройки по умолчанию
' ------------------------------
' Начальные значения предназначены для автономного API и модульных тестов.
' При расчёте сочетаний используются обязательные настройки из Config;
' отсутствующий или неверный ввод книги не заменяется этим набором.
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
    mSolverSecantMaxRestarts = 2
    mDiagnosticsEnabled = True
    mProbeCacheNextSlot = 1
End Sub

' ------------------------------
' Настройки поиска несущей способности
' ------------------------------
' Принимает настройки внешнего поиска предела и внутреннего CSectionSolver,
' который решает НДС каждой пробной нагрузки. Явно пустые обязательные
' численные настройки не заменяются начальными значениями.
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
    mSolverSecantMaxRestarts = settings.GetLong("Solver.SecantMaxRestarts", mSolverSecantMaxRestarts, False)
    If Not units Is Nothing Then
        mSolverToleranceN = Sgn(mSolverToleranceN) * Abs(units.InputForceToInternal(mSolverToleranceN)) ' Отрицательный допуск остается ошибкой настройки и после перевода единиц.
        mSolverToleranceMx = Sgn(mSolverToleranceMx) * Abs(units.InputMomentMxToInternal(mSolverToleranceMx))
        mSolverToleranceMy = Sgn(mSolverToleranceMy) * Abs(units.InputMomentMyToInternal(mSolverToleranceMy))
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

' Ищет предельное состояние для заданной зависимости нагрузки от lambda:
' Target = Offset + lambda*Base для N, Mx и My.
' Название пользовательского режима на сам численный алгоритм не влияет.
Public Sub SolveByUltimateLoadPath(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal nOffset As Double, ByVal nBase As Double, _
        ByVal mxOffset As Double, ByVal mxBase As Double, _
        ByVal myOffset As Double, ByVal myBase As Double, _
        Optional ByVal allowPureAxialRedirect As Boolean = True)
    If Not EnsureSearchDependencies(section, concreteMaterial, steelMaterial) Then Exit Sub
    If allowPureAxialRedirect And IsPureAxialLoadPath(section, nOffset, nBase, mxOffset, mxBase, myOffset, myBase) Then
        SolvePureAxialUltimateByLoadMultiplier section, concreteMaterial, steelMaterial, _
            nOffset, nBase, mxOffset, mxBase, myOffset, myBase
        Exit Sub
    End If

    Dim request As CLimitSearchRequest
    Set request = New CLimitSearchRequest
    InitializeCapacitySearchRequest request, Me, section, concreteMaterial, steelMaterial, _
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
        ReplaceDiagnosticLog "Силовая траектория UltimateStrain перенаправлена на LoadMultiplier." & vbCrLf & DiagnosticLog
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
    Dim firstSearchExecuted As Boolean ' Первая попытка остается выполненной даже при раннем отказе резервного метода.
    firstStopReason = mStopReason
    firstLimitState = mLimitState
    firstDiagnostics = DiagnosticLog
    firstSearchExecuted = mSearchExecuted

    SolveByLoadPathMultiplier section, concreteMaterial, steelMaterial, _
        nOffset, nBase, mxOffset, mxBase, myOffset, myBase, allowForcePathRetainedLimit
    mSearchExecuted = mSearchExecuted Or firstSearchExecuted

    If mDiagnosticsEnabled Then
        ReplaceDiagnosticLog "Auto: первая попытка UltimateStrain не дала принятую предельную точку; статус=" & _
            firstLimitState & "; причина=" & firstStopReason & vbCrLf & _
            firstDiagnostics & _
            "Auto: переходим к LoadMultiplier по той же lambda-траектории." & vbCrLf & _
            DiagnosticLog
    End If
    If mConverged Then
        mStopReason = "Auto: UltimateStrain не сошелся, предельная точка найдена через LoadMultiplier. " & mStopReason
    End If
End Sub

' Чистое осевое масштабирование без постоянных моментов является
' вырожденным для UltimateStrain при горизонтальном плато диаграммы:
' разные равномерные деформации дают одну и ту же предельную N. Эта проверка
' распознаёт осевую траекторию относительно центра бетонной сетки, чтобы
' разрешённое перенаправление в LoadMultiplier произошло до попытки Newton.
Private Function IsPureAxialLoadPath(ByVal section As CSectionModel, ByVal nOffset As Double, ByVal nBase As Double, _
        ByVal mxOffset As Double, ByVal mxBase As Double, _
        ByVal myOffset As Double, ByVal myBase As Double) As Boolean
    On Error GoTo NotAxial ' Ошибочную геометрию далее отклоняет общая проверка входа поиска с отдельным кодом причины.
    If section.ConcreteCount <= 0 Then Exit Function
    Dim math As CLoadPathMath, centerX As Double, centerY As Double, area As Double
    section.ConcreteMeshCentroid centerX, centerY, area
    Set math = New CLoadPathMath
    IsPureAxialLoadPath = math.IsPureAxialLoadPath(nOffset, nBase, mxOffset, mxBase, myOffset, myBase, centerX, centerY)
    Exit Function
NotAxial:
    Err.Clear
End Function

' True для классической траектории "N = const, M = lambda*Mbase".
' Это lambda*Mx, lambda*My и lambda*Mxy. Для них внутри универсального API используется
' быстрая система UltimateStrain, потому что она лучше обусловлена
' именно для моментного поиска несущей способности.
Private Function IsMomentOnlyLoadPath(ByVal nBase As Double, ByVal mxBase As Double, ByVal myBase As Double) As Boolean
    IsMomentOnlyLoadPath = Abs(nBase) <= 0.000000001 And _
        Sqr(mxBase * mxBase + myBase * myBase) > 0.000000001
End Function

' Проверяет физическую предельную плоскость на исходной траектории.
' Удержанная проба дает лишь форму стартовой плоскости. После доведения
' до первого предела lambda определяется заново по внутренним усилиям,
' затем подтверждаются все три компоненты пути и физический критерий.
Private Sub AcceptRetainedPhysicalCapacity(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal loadPath As CLoadPathVector)
    If LimitSearchHasTerminalFailure Then Exit Sub
    If mSearchBoundReached Then Exit Sub
    If loadPath Is Nothing Then Exit Sub
    If mLastAcceptableSolver Is Nothing Then Exit Sub
    If Not mLastAcceptableSolver.Converged Then Exit Sub
    mConverged = False ' Силовая скобка еще не подтверждает предельную плоскость на плато. Ее предварительный успех не должен останавливать Newton или делать непринятую пробу окончательным результатом поиска.
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
    PrepareUltimateNumericalPath section, loadPath
    Dim boundaryLambda As Double
    boundaryLambda = mUltimatePathMath.NumericalLambda(boundarySolver)
    If boundaryLambda <= 0# Then Exit Sub
    If boundaryLambda > mMaxLambda Then
        ' Проекция усилий может округлиться за техническую границу скобки.
        ' Принимаем саму границу только по исходным допускам N/Mx/My;
        ' произвольный допуск по lambda здесь не добавляется.
        If Not boundarySolver.ConfirmEquilibrium(loadPath.TargetN(mMaxLambda), _
                loadPath.TargetMx(mMaxLambda), loadPath.TargetMy(mMaxLambda)) Then Exit Sub
        boundaryLambda = mMaxLambda
    End If
    If Not boundarySolver.ConfirmEquilibrium(loadPath.TargetN(boundaryLambda), _
            loadPath.TargetMx(boundaryLambda), loadPath.TargetMy(boundaryLambda)) Then
        ' На плато масштабирование может потерять моментную жесткость.
        ' Старт привязываем к краям именно критического материала: предел
        ' бетона нельзя заменять существенно большей деформацией стали.
        ' Это только старт Newton, не доказательство равновесия или несущей.
        Dim startSolver As CSectionSolver
        Set startSolver = mLastAcceptableSolver
        Dim anchorStrain As Double
        Dim edgeStrain As Double
        Dim edgeLimit As Double
        If CriticalLimitState() = STATE_CONCRETE_LIMIT Then
            anchorStrain = mLastAcceptableSolver.MaxConcreteStrain
            edgeStrain = mLastAcceptableSolver.MinConcreteStrain
            edgeLimit = mConcreteCompressionLimit
        ElseIf CriticalLimitState() = STATE_CONCRETE_TENSION_LIMIT Then
            anchorStrain = mLastAcceptableSolver.MinConcreteStrain
            edgeStrain = mLastAcceptableSolver.MaxConcreteStrain
            edgeLimit = mConcreteTensionLimit
        ElseIf mCriticalStrain > 0# Then
            anchorStrain = mLastAcceptableSolver.MinSteelStrain
            edgeStrain = mLastAcceptableSolver.MaxSteelStrain
            edgeLimit = mSteelTensionLimit
        Else
            anchorStrain = mLastAcceptableSolver.MaxSteelStrain
            edgeStrain = mLastAcceptableSolver.MinSteelStrain
            edgeLimit = -mSteelCompressionLimit
        End If
        If Abs(edgeStrain - anchorStrain) > 0.000000000001 Then
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
                "Не удалось подтвердить равновесие предела по выбранному пути.") Then
            If LimitSearchHasTerminalFailure Then Exit Sub
            If Abs(edgeStrain - anchorStrain) <= 0.000000000001 Then Exit Sub
            ' Нулевой anchor может изменить активные ветви материалов.
            ' Повторяем тот же Newton, сохранив деформацию противоположного
            ' края критического материала из допустимой точки.
            strainScale = (edgeLimit - anchorStrain) / (edgeStrain - anchorStrain)
            If strainScale <= 1# Then Exit Sub
            Set startSolver = New CSectionSolver
            startSolver.EvaluateStrainPlane section, concreteMaterial, steelMaterial, _
                anchorStrain + (mLastAcceptableSolver.Epsilon0 - anchorStrain) * strainScale, _
                mLastAcceptableSolver.KappaX * strainScale, _
                mLastAcceptableSolver.KappaY * strainScale
            AppendSearchDiagnostic "повторный старт предельной плоскости сохраняет деформацию противоположного края критического материала"
            If Not search.RunNewton(problem, section, concreteMaterial, steelMaterial, _
                    loadPath, startSolver, _
                    "Не удалось подтвердить равновесие предела из сохраненного НДС.") Then Exit Sub
        End If
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
    mSafetyFactor = mLambdaUltimate
    mUtilization = 1# / mLambdaUltimate
    mConverged = True
    mStopReason = "Физический предел подтвержден равновесием предельной плоскости на исходной траектории нагрузки."
    AppendSearchDiagnostic "предел подтвержден отдельной проверкой усилий и деформаций из сохраненного НДС; lambda=" & _
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
    Dim originX As Double, originY As Double

    Dim properties As CSectionPropertiesCalculator
    Set properties = New CSectionPropertiesCalculator
    properties.CalculateConcrete section
    originX = properties.CentroidX: originY = properties.CentroidY

    For i = 1 To section.ConcreteCount
        area = area + section.ConcreteArea(i)
        maxAbsX = MaxDouble(maxAbsX, Abs(section.ConcreteX(i) - originX))
        maxAbsY = MaxDouble(maxAbsY, Abs(section.ConcreteY(i) - originY))
    Next i
    If section.RebarCount > 0 Then
        For i = 1 To section.RebarCount
            area = area + section.RebarArea(i)
            maxAbsX = MaxDouble(maxAbsX, Abs(section.RebarX(i) - originX))
            maxAbsY = MaxDouble(maxAbsY, Abs(section.RebarY(i) - originY))
        Next i
    End If

    targetNValue = TargetN(1#, nOffset, nBase)
    targetMxValue = TargetMx(1#, mxBase, mxOffset)
    targetMyValue = TargetMy(1#, myBase, myOffset)
    targetMxValue = targetMxValue - targetNValue * originY
    targetMyValue = targetMyValue - targetNValue * originX
    If area > 0# Then eps0 = targetNValue / (area * 30000#)

    strainLimit = Abs(mConcreteCompressionLimit)
    If mSteelCompressionLimit > 0# Then strainLimit = MinDouble(strainLimit, mSteelCompressionLimit)
    If mSteelTensionLimit > 0# Then strainLimit = MinDouble(strainLimit, mSteelTensionLimit)
    If strainLimit <= 0# Then strainLimit = 0.001

    If targetNValue < 0# Then
        eps0 = -0.25 * strainLimit ' Доля меньшего физического предела не исчезает при малой N и не растет до запредельных деформаций при перегрузке.
    ElseIf targetNValue > 0# Then
        eps0 = 0.25 * strainLimit
    End If

    momentNorm = Sqr(targetMxValue * targetMxValue + targetMyValue * targetMyValue)
    If momentNorm <= 0.000000001 Then Exit Sub

    If maxAbsY > 0# Then kx = Sgn(targetMxValue) * strainLimit / maxAbsY * Abs(targetMxValue) / momentNorm
    If maxAbsX > 0# Then ky = Sgn(targetMyValue) * strainLimit / maxAbsX * Abs(targetMyValue) / momentNorm
    eps0 = eps0 - kx * originY - ky * originX
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

    Dim originX As Double, originY As Double
    Dim properties As CSectionPropertiesCalculator
    Set properties = New CSectionPropertiesCalculator
    properties.CalculateConcrete section
    originX = properties.CentroidX: originY = properties.CentroidY

    For i = 1 To section.ConcreteCount
        area = area + section.ConcreteArea(i)
        maxAbsX = MaxDouble(maxAbsX, Abs(section.ConcreteX(i) - originX))
        maxAbsY = MaxDouble(maxAbsY, Abs(section.ConcreteY(i) - originY))
    Next i
    If section.RebarCount > 0 Then
        For i = 1 To section.RebarCount
            area = area + section.RebarArea(i)
            maxAbsX = MaxDouble(maxAbsX, Abs(section.RebarX(i) - originX))
            maxAbsY = MaxDouble(maxAbsY, Abs(section.RebarY(i) - originY))
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
    eps0 = eps0 - kx * originY - ky * originX
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
        ByRef rLimit As Double, ByRef lambdaValue As Double, ByRef state As String, _
        Optional ByVal centeredPlane As Boolean = False) As Boolean
    On Error GoTo Failed

    Set solver = New CSectionSolver
    If centeredPlane Then
        solver.EvaluatePhysicalCenterPlane section, concreteMaterial, steelMaterial, eps0, kx, ky
    Else
        solver.EvaluatePhysicalStrainPlane section, concreteMaterial, steelMaterial, eps0, kx, ky
    End If
    mIterations = mIterations + 1
    Set mLastSolver = solver
    UpdateCriticalState section, concreteMaterial, solver

    Dim baseNorm As Double
    Dim baseNorm2 As Double
    baseNorm2 = mxBase * mxBase + myBase * myBase
    baseNorm = Sqr(baseNorm2)

    Dim userMxInt As Double
    Dim userMyInt As Double
    ' И внутренние усилия, и исходный Offset переносим в центр бетона.
    ' Иначе погрешность N умножается на удаленные координаты чертежа.
    ' При rN=0 это та же траектория; финальная приемка проверяет исходные N/Mx/My.
    Dim numericalN As Double
    solver.NumericalEfforts numericalN, userMxInt, userMyInt
    userMxInt = userMxInt - (mxOffset - nSpecified * mUltimateOriginY)
    userMyInt = userMyInt - (myOffset - nSpecified * mUltimateOriginX)
    lambdaValue = (userMxInt * mxBase + userMyInt * myBase) / baseNorm2

    rN = (solver.Nint - nSpecified) / MaxDouble(mSolverToleranceN, Abs(nSpecified) * 0.000000001)
    Dim perpendicularMoment As Double
    perpendicularMoment = (userMxInt * myBase - userMyInt * mxBase) / baseNorm
    rDir = perpendicularMoment / MaxDouble(MaxDouble(Abs(lambdaValue) * baseNorm, baseNorm), 1#)
    rLimit = mCriticalStrainUtilization - 1#
    state = CriticalLimitState()

    AppendSearchDiagnostic "ultimate iter=" & CStr(mIterations) & _
        "; eps0=" & FormatNumberInvariant(solver.Epsilon0) & _
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
        ByRef rLimit As Double, ByRef lambdaValue As Double, ByRef state As String, _
        Optional ByVal centeredPlane As Boolean = False) As Boolean
    On Error GoTo Failed

    Set solver = New CSectionSolver
    If centeredPlane Then
        solver.EvaluatePhysicalCenterPlane section, concreteMaterial, steelMaterial, eps0, kx, ky
    Else
        solver.EvaluatePhysicalStrainPlane section, concreteMaterial, steelMaterial, eps0, kx, ky
    End If
    mIterations = mIterations + 1
    Set mLastSolver = solver
    UpdateCriticalState section, concreteMaterial, solver

    mUltimatePathMath.BuildNumericalResiduals solver, lambdaValue, rPath1, rPath2
    rLimit = mCriticalStrainUtilization - 1#
    state = CriticalLimitState()

    AppendSearchDiagnostic "ultimatePath iter=" & CStr(mIterations) & _
        "; eps0=" & FormatNumberInvariant(solver.Epsilon0) & _
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
    If mCriticalLimitType = STATE_CONCRETE_LIMIT Then
        limitAbs = Abs(mConcreteCompressionLimit)
    ElseIf mCriticalLimitType = STATE_CONCRETE_TENSION_LIMIT Then
        limitAbs = Abs(mConcreteTensionLimit)
    ElseIf mCriticalLimitType = STATE_STEEL_LIMIT Then
        limitAbs = SteelLimitForStrain(mCriticalStrain)
    End If
    UltimateStrainRatioTolerance = mStrainTolerance / MaxDouble(limitAbs, 0.000000001)
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

' Ищет предельный множитель lambda для Target = Offset + lambda*Base.
' Все способы LoadPath используют тот же поиск границ и уточнение
' Bisection, Brent или Secant независимо от того, какие усилия масштабируются.
Public Sub SolveByLoadPathMultiplier(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal nOffset As Double, ByVal nBase As Double, _
        ByVal mxOffset As Double, ByVal mxBase As Double, _
        ByVal myOffset As Double, ByVal myBase As Double, _
        Optional ByVal allowForcePathRetainedLimit As Boolean = False)
    If Not EnsureSearchDependencies(section, concreteMaterial, steelMaterial) Then Exit Sub
    Dim request As CLimitSearchRequest
    Set request = New CLimitSearchRequest
    InitializeCapacitySearchRequest request, Me, section, concreteMaterial, steelMaterial, _
        "LoadMultiplier", nOffset, nBase, mxOffset, mxBase, myOffset, myBase, _
        allowForcePathRetainedLimit, _
        IsPureAxialLoadPath(section, nOffset, nBase, mxOffset, mxBase, myOffset, myBase)

    Dim search As CLoadMultiplierSearch
    Set search = New CLoadMultiplierSearch
    search.Execute request
End Sub

' Готовит физические проверки и рабочие данные для LoadMultiplier.
' Одномерный численный поиск lambda выполняет общий CLoadMultiplierSearch.
Public Function LimitSearchPrepareCapacityLoadMultiplier(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal nBase As Double, ByVal mxBase As Double, ByVal myBase As Double) As Boolean
    ClearResult
    mSolutionMethod = "LoadMultiplier"

    If Not ValidateLoadPathCapacityInputs(section, concreteMaterial, steelMaterial, nBase, mxBase, myBase) Then Exit Function
    If Not ValidateLoadMultiplierSettings() Then Exit Function
    LimitSearchPrepareCapacityLoadMultiplier = True
End Function

' Оценивает критерий несущей способности одной пробной нагрузки lambda.
' Отрицательное значение означает запас до предельной деформации, ноль
' или положительное — достижение предела. Учитывает физические пределы
' фактических диаграмм модели прочности.
Public Function LimitSearchEvaluateCapacityLoadMultiplier(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal loadPath As CLoadPathVector, ByVal lambdaValue As Double, _
        ByRef functionValue As Double, ByRef probeState As String, _
        Optional ByVal startSolver As CSectionSolver = Nothing, _
        Optional ByVal allowRetries As Boolean = True) As Boolean
    mSearchExecuted = True
    Dim nValue As Double
    Dim mxValue As Double
    Dim myValue As Double
    loadPath.Target lambdaValue, nValue, mxValue, myValue

    If allowRetries Then
        probeState = ProbeWithRetries(section, concreteMaterial, steelMaterial, _
            nValue, mxValue, myValue, lambdaValue, startSolver)
    Else
        ' Предварительная проверка UltimateStrain требует только доказанного
        ' физического отказа. Неудачная проба не запускает удвоение ступеней:
        ' основной поиск сохранит все свои обычные старты и бюджеты.
        probeState = ProbeOnce(section, concreteMaterial, steelMaterial, _
            nValue, mxValue, myValue, lambdaValue, 0, startSolver)
    End If
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

' Сохраняет превышение критерия уже в постоянной нагрузке при lambda = 0.
' Внешний BaseFail затем определяется кодом причины, а не текстом комментария.
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

' Преобразует код исхода пробной нагрузки в короткое инженерное пояснение
' для Results и сводки. Сам код сохраняется в LimitState и ResultCode.
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

' Проверяет предельную плоскость исходной траектории из удержанного старта.
' Одна последняя acceptable-проба без подтверждения усилий и физического
' предела не становится успешным результатом несущей способности.
Public Sub LimitSearchAcceptRetainedPhysicalCapacity(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal loadPath As CLoadPathVector)
    AcceptRetainedPhysicalCapacity section, concreteMaterial, steelMaterial, loadPath
End Sub

' Сохраняет последнюю допустимую нагрузку и её НДС. При подтверждённой
' верхней границе они нужны для приёма предела с безопасной стороны без нового решения.
' При численной неудаче это только диагностика, а не найденная несущая способность.
Public Sub LimitSearchRecordCapacityAcceptable(ByVal nValue As Double, _
        ByVal mxValue As Double, ByVal myValue As Double)
    RecordAcceptableState mLastSolver, nValue, mxValue, myValue
End Sub

' Сохраняет численную неудачу между допустимой и предельной нагрузками
' и последнее допустимое lambda для диагностики. Отдельно эта точка
' несущую способность не подтверждает.
Public Sub LimitSearchHandleCapacityBracketNumericalFailure(ByVal lowLambda As Double, _
        ByVal probeState As String)
    mLimitState = probeState
    mStopReason = "При уточнении интервала поиска не удалось найти равновесие. Сохранён последний множитель нагрузки lambda с подтверждённым равновесием; предельная точка не подтверждена."
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

' Сохраняет численную неудачу общего поиска в результате несущей способности.
Public Sub LimitSearchSetCapacityNumericalFailure(ByVal reason As String)
    mLimitState = STATE_NUMERICAL_FAILURE
    mFailureCode = sfcNumericalFailure
    mStopReason = reason
End Sub

' Сохраняет ошибку исходных данных общего поиска в результате несущей способности.
Public Sub LimitSearchSetCapacityInvalidInput(ByVal reason As String)
    SetInvalidInput reason
End Sub

' Добавляет строку в общий журнал поиска несущей способности.
Public Sub LimitSearchAppendCapacityDiagnostic(ByVal message As String)
    AppendSearchDiagnostic message
End Sub

' Готовит физические проверки и рабочие данные для UltimateStrain.
' Общий цикл Ньютона выполняет CUltimateStrainSearch.
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
    mMeritJacobianForAffinePath = mUltimateMomentOnlyPath Or loadPath.NOffset <> 0# Or _
        loadPath.MxOffset <> 0# Or loadPath.MyOffset <> 0#
    LimitSearchPrepareCapacityUltimate = True
End Function

' Готовит начальное поле деформаций для общего поиска несущей способности по UltimateStrain.
Public Sub LimitSearchInitializeUltimateGuess(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal loadPath As CLoadPathVector, ByVal startSolver As CSectionSolver, _
        ByRef eps0 As Double, ByRef kx As Double, ByRef ky As Double)
    mSearchExecuted = True ' Общий Newton проверил настройки; отказ при подготовке плоскости тоже относится к реально начатой численной попытке, а не к раннему пропуску.
    PrepareUltimateNumericalPath section, loadPath
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
    ' Упругая оценка при перегрузке может оказаться далеко за диаграммой.
    ' Ограничиваем только старт общим критерием всех материалов, не усилия
    ' пути и не допуски. Переданную готовую плоскость выше не меняем.
    Dim initialState As CSectionSolver, builder As CStateGuessBuilder
    Set initialState = New CSectionSolver
    initialState.EvaluatePhysicalStrainPlane section, concreteMaterial, steelMaterial, eps0, kx, ky
    UpdateCriticalState section, concreteMaterial, initialState
    If mCriticalStrainUtilization > 1# Then
        Set builder = New CStateGuessBuilder
        builder.ScaleToPhysicalLimit section, mCriticalStrainUtilization, eps0, kx, ky
    End If
End Sub

' Численный контекст нужен и отдельной оценке плоскости, не только поиску.
' Повторное использование допускается лишь при точном совпадении геометрии,
' шести компонент пути и трех допусков; входной изменяемый path не сохраняется.
Private Sub PrepareUltimateNumericalPath(ByVal section As CSectionModel, ByVal loadPath As CLoadPathVector)
    If mUltimatePathMath Is Nothing Then Set mUltimatePathMath = New CLoadPathMath
    mUltimatePathMath.PrepareNumericalPath section, loadPath, mSolverToleranceN, mSolverToleranceMx, mSolverToleranceMy, True
    Dim frame As CSectionNumericalFrame
    Set frame = mUltimatePathMath.NumericalFrame
    mUltimateOriginX = frame.OriginX: mUltimateOriginY = frame.OriginY
    mUltimatePathTolerance = mUltimatePathMath.NumericalPathTolerance
    Dim forceScale As Double
    forceScale = MaxDouble(1#, MaxDouble(Abs(loadPath.NOffset), _
        Sqr(loadPath.MxBase * loadPath.MxBase + loadPath.MyBase * loadPath.MyBase) / frame.LengthScale))
    mUltimateMomentNMeritWeight = MaxDouble(mSolverToleranceN, Abs(loadPath.NOffset) * 0.000000001) / forceScale
End Sub

' Строит независимый старт по фактической жесткости и активным ветвям материалов.
' Масштабирование ограничивает только старт; предельные усилия подтвердит Newton.
Public Function LimitSearchAlternativeUltimateStart(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal loadPath As CLoadPathVector) As CSectionSolver
    If Not TypeOf concreteMaterial Is CMaterialDiagram Then Exit Function
    If Not TypeOf steelMaterial Is CMaterialDiagram Then Exit Function
    Dim builder As CStateGuessBuilder, concrete As CMaterialDiagram, steel As CMaterialDiagram
    Set builder = New CStateGuessBuilder
    Set concrete = concreteMaterial: Set steel = steelMaterial
    Dim eps0 As Double, kx As Double, ky As Double, solver As CSectionSolver
    If Not builder.BuildForDirectState(section, concrete, steel, loadPath.TargetN(1#), _
            loadPath.TargetMx(1#), loadPath.TargetMy(1#), eps0, kx, ky) Then Exit Function
    Set solver = New CSectionSolver
    solver.EvaluateStrainPlane section, concreteMaterial, steelMaterial, eps0, kx, ky
    UpdateCriticalState section, concreteMaterial, solver
    If mCriticalStrainUtilization <= 0# Then Exit Function
    builder.ScaleToPhysicalLimit section, mCriticalStrainUtilization, eps0, kx, ky
    solver.EvaluateStrainPlane section, concreteMaterial, steelMaterial, eps0, kx, ky
    AppendSearchDiagnostic "UltimateStrain: повторный старт по фактическим диаграммам и форме сечения."
    Set LimitSearchAlternativeUltimateStart = solver
End Function

' Возвращает две невязки траектории нагрузки и невязку предельной деформации для UltimateStrain.
Public Function LimitSearchEvaluateUltimateResidual(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal loadPath As CLoadPathVector, ByVal eps0 As Double, _
        ByVal kx As Double, ByVal ky As Double, ByRef solver As CSectionSolver, _
        ByRef rPath1 As Double, ByRef rPath2 As Double, ByRef rLimit As Double, _
        ByRef lambdaValue As Double, ByRef state As String, Optional ByVal centeredPlane As Boolean = False) As Boolean
    PrepareUltimateNumericalPath section, loadPath
    If mUltimateMomentOnlyPath Then
        LimitSearchEvaluateUltimateResidual = EvaluateUltimateMomentResidual(section, _
            concreteMaterial, steelMaterial, loadPath.NOffset, loadPath.MxBase, _
            loadPath.MyBase, loadPath.MxOffset, loadPath.MyOffset, eps0, kx, ky, _
            solver, rPath1, rPath2, rLimit, lambdaValue, state, centeredPlane)
    Else
        LimitSearchEvaluateUltimateResidual = EvaluateUltimateLoadPathResidual(section, _
            concreteMaterial, steelMaterial, loadPath.NOffset, loadPath.NBase, _
            loadPath.MxOffset, loadPath.MxBase, loadPath.MyOffset, loadPath.MyBase, _
            eps0, kx, ky, solver, rPath1, rPath2, rLimit, lambdaValue, state, centeredPlane)
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
        LimitSearchUltimateResidualReached = (Abs(rPath1) <= mUltimatePathTolerance And _
            Abs(rPath2) <= mUltimatePathTolerance And Abs(rLimit) <= UltimateStrainRatioTolerance())
    End If
End Function

' Для моментной и аффинной систем масштабы строк определяет текущий Якобиан.
' Однородная силовая траектория сохраняет норму уже масштабированного CLoadPathMath.
Friend Property Get LimitSearchJacobianMeritForPath() As Boolean
    LimitSearchJacobianMeritForPath = mMeritJacobianForAffinePath
End Property

' Вычисляет общую меру невязок для подбора шага UltimateStrain.
' В моментной постановке учитывает относительные ошибки усилий и деформаций,
' чтобы малый допуск N не мешал исправлению направления момента.
' При окончательном подтверждении используются исходные допуски.
Public Function LimitSearchUltimateResidualNorm(ByVal rPath1 As Double, _
        ByVal rPath2 As Double, ByVal rLimit As Double) As Double
    If mUltimateMomentOnlyPath Then
        LimitSearchUltimateResidualNorm = Abs(rPath1) * mUltimateMomentNMeritWeight + Abs(rPath2) + Abs(rLimit)
    Else
        Dim magnitude As Double ' Общий масштаб защищает квадраты; евклидова норма непрерывна при смене знака моментной невязки.
        magnitude = MaxDouble(Abs(rPath1), MaxDouble(Abs(rPath2), Abs(rLimit)))
        If magnitude > 0# Then LimitSearchUltimateResidualNorm = magnitude * _
            Sqr((rPath1 / magnitude) ^ 2 + (rPath2 / magnitude) ^ 2 + (rLimit / magnitude) ^ 2)
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
    If Not IsValidUltimateLambda(lambdaValue, loadPath.NOffset, loadPath.NBase, _
            loadPath.MxBase, loadPath.MyBase, loadPath.MxOffset, loadPath.MyOffset) Then
        Dim rejectionReason As String
        rejectionReason = "UltimateStrain: найден недопустимый lambda: " & FormatNumberInvariant(lambdaValue)
        AppendSearchDiagnostic rejectionReason
        LimitSearchSetFailure sfcNumericalFailure, rejectionReason
        Exit Function
    End If
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
    StoreLoadPathUltimateComponents lambdaValue, loadPath.NOffset, loadPath.NBase, _
        loadPath.MxBase, loadPath.MyBase, loadPath.MxOffset, loadPath.MyOffset
    mLimitState = state
    mUpperLimitState = state
    mConverged = True
    mFailureCode = sfcNone
    mFailureSettingKeys = vbNullString
    LimitSearchFinalizeUltimateResult = True
End Function

' Сохраняет инженерный итог несущей способности после общего поиска UltimateStrain.
Public Sub LimitSearchCompleteCapacityUltimate(ByVal loadPath As CLoadPathVector)
    If Not mConverged Then
        If Len(mLimitState) = 0 Then mLimitState = STATE_NUMERICAL_FAILURE
        If Len(mStopReason) = 0 Then mStopReason = "Поиск по предельной деформации (UltimateStrain) не завершён: итерации Newton не нашли равновесие."
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

' Сохраняет фактический код причины общего поиска. Ошибка настройки,
' запрещающая продолжение, не повторяется и не считается физическим отказом траектории.
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
' При масштабировании N возвращается полный целевой вектор; для моментного
' пути Nult пуст, а моменты относятся к масштабируемой части. Фактические
' полные усилия подтвержденной предельной плоскости хранятся в LastSolver.
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
        mNUltimate = 0#
        mMxUltimate = mLambdaUltimate * mxBase
        mMyUltimate = mLambdaUltimate * myBase
    End If
    mMomentUltimate = Sqr(mMxUltimate * mMxUltimate + mMyUltimate * mMyUltimate)
End Sub

' Проверяет, масштабируется ли N на выбранном пути нагрузки.
' При изменении только моментов продольная сила остаётся постоянной частью Offset.
Private Function CapacityScalesAxial(ByVal nBase As Double) As Boolean
    CapacityScalesAxial = (Abs(nBase) > 0.000000001)
End Function

' До создания запроса поиска проверяет обязательные программные объекты.
' При их отсутствии очищает старый результат и сохраняет внутреннюю ошибку
' с кодом причины вместо необработанного исключения VBA.
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
        SetInvalidConfiguration "Число итераций решателя несущей способности должно быть положительным.", "Capacity.SolverMaxIterations"
    Else
        Dim validation As CSectionSolver
        Set validation = New CSectionSolver ' Оба метода поиска используют общую проверку численных настроек без пробного решения НДС.
        validation.SolverMethod = mSolverMethod: validation.MaxIterations = mSolverMaxIterations
        validation.ToleranceN = mSolverToleranceN: validation.ToleranceMx = mSolverToleranceMx
        validation.ToleranceMy = mSolverToleranceMy: validation.SecantMaxRestarts = mSolverSecantMaxRestarts
        If validation.ValidateProblem(section, concreteMaterial, steelMaterial) Then
            ValidateLoadPathCapacityInputs = True
        Else
            Set mLastSolver = validation
            LimitSearchSetFailure validation.FailureCode, validation.StopReason
            mFailureSettingKeys = validation.FailureSettingKeys
        End If
    End If
End Function

' Определяет наибольшее число удвоений ступеней, представимое счетчиком Long.
' Возвращает -1 для неверного базового числа. Это техническая граница
' BaseLoadSteps*2^attempt, а не ограничение физических нагрузок или точности.
Friend Function MaximumRetryCountForLoadSteps(ByVal baseLoadSteps As Long) As Long
    MaximumRetryCountForLoadSteps = -1
    If baseLoadSteps < 1 Or baseLoadSteps > 2147483646 Then Exit Function
    Dim steps As Double, count As Long
    steps = CDbl(baseLoadSteps)
    Do While steps <= 1073741823#
        steps = steps * 2#
        count = count + 1
    Loop
    MaximumRetryCountForLoadSteps = count
End Function

' Проверяет настройки ветки LoadMultiplier перед одномерным поиском lambda.
' Здесь контролируются только параметры самой скобки/поиска, а не геометрия
' или материалы, которые проверяются выше по цепочке.
Private Function ValidateLoadMultiplierSettings() As Boolean
    If Len(NormalizedSearchMethod()) = 0 Then
        SetInvalidConfiguration "Capacity.SearchMethod должен быть Bisection, Brent или Secant.", "Capacity.SearchMethod"
    ElseIf mInitialLambdaStep <= 0# Then
        SetInvalidConfiguration "На листе Config в строке Capacity.InitialLambda задайте начальный шаг больше нуля.", "Capacity.InitialLambda"
    ElseIf mLambdaTolerance <= 0# Then
        SetInvalidConfiguration "На листе Config в строке Capacity.ToleranceLambda задайте допуск поиска больше нуля.", "Capacity.ToleranceLambda"
    ElseIf mMaxLambda <= 0# Then
        SetInvalidConfiguration "На листе Config в строке Capacity.MaxLambda задайте верхнюю границу поиска больше нуля.", "Capacity.MaxLambda"
    ElseIf mMaxRetries < 0 Then
        SetInvalidConfiguration "На листе Config в строке Capacity.MaxRetries задайте неотрицательное число повторных попыток.", "Capacity.MaxRetries"
    ElseIf mSolverBaseLoadSteps <= 0 Then
        SetInvalidConfiguration "На листе Config в строке Capacity.BaseLoadSteps задайте число ступеней нагрузки больше нуля.", "Capacity.BaseLoadSteps"
    ElseIf MaximumRetryCountForLoadSteps(mSolverBaseLoadSteps) < 0 Then
        SetInvalidConfiguration "Число ступеней Capacity.BaseLoadSteps не помещается в расчетный счетчик. Уменьшите его до 2147483646 или менее.", "Capacity.BaseLoadSteps"
    ElseIf mMaxRetries > MaximumRetryCountForLoadSteps(mSolverBaseLoadSteps) Then
        SetInvalidConfiguration "Число ступеней с учетом повторов не помещается в расчетный счетчик. Уменьшите Capacity.MaxRetries или Capacity.BaseLoadSteps.", "Capacity.MaxRetries;Capacity.BaseLoadSteps"
    Else
        ValidateLoadMultiplierSettings = True
    End If
End Function

' Проверяет параметры UltimateStrain. В отличие от LoadMultiplier,
' этот метод не ограничивается техническим Capacity.MaxLambda:
' после нахождения предельного поля проверяется корректность самого множителя.
Private Function ValidateUltimateStrainSettings() As Boolean
    If mStrainTolerance <= 0# Then
        SetInvalidConfiguration "На листе Config в строке Capacity.ToleranceStrain задайте допуск предельной деформации больше нуля.", "Capacity.ToleranceStrain"
    Else
        ValidateUltimateStrainSettings = True
    End If
End Function

' Сохраняет InputErr без необработанного исключения VBA.
' Причина затем выводится в общей сводке и подробной таблице.
Private Sub SetInvalidInput(ByVal reason As String)
    mLimitState = STATE_INVALID_INPUT
    mFailureCode = sfcInvalidInput
    mFailureSettingKeys = vbNullString
    mStopReason = reason
End Sub

' Сохраняет ошибку численных настроек отдельным кодом без повторных попыток
' и смены метода. Внутренний LimitState и внешнее обозначение статуса различаются.
Private Sub SetInvalidConfiguration(ByVal reason As String, Optional ByVal sourceKeys As String = vbNullString)
    mLimitState = STATE_INVALID_INPUT
    mFailureCode = sfcInvalidConfiguration
    mStopReason = reason
    mFailureSettingKeys = sourceKeys
End Sub

' Отсутствие обязательного программного объекта или неожиданная ошибка
' считаются внутренней ошибкой программы, а не ошибкой пользовательских данных.
Private Sub SetInternalError(ByVal reason As String)
    mLimitState = STATE_INTERNAL_ERROR
    mFailureCode = sfcInternalError
    mFailureSettingKeys = vbNullString
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
        ByVal nSpecified As Double, ByVal mx As Double, ByVal my As Double, ByVal lambdaValue As Double, _
        Optional ByVal startSolver As CSectionSolver = Nothing) As String

    Dim cachedState As String
    If TryRestoreProbeCache(nSpecified, mx, my, lambdaValue, cachedState) Then
        ProbeWithRetries = cachedState
        Exit Function
    End If

    Dim attempt As Long
    Dim probeState As String
    For attempt = 0 To mMaxRetries
        probeState = ProbeOnce(section, concreteMaterial, steelMaterial, nSpecified, mx, my, lambdaValue, attempt, startSolver)
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

' Решает НДС пробной нагрузки lambda и различает допустимую точку,
' физический предел, численную несходимость и ошибку ввода.
Private Function ProbeOnce(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal nSpecified As Double, ByVal mx As Double, ByVal my As Double, _
        ByVal lambdaValue As Double, ByVal attempt As Long, ByVal startSolver As CSectionSolver) As String

    On Error GoTo Failed
    mIterations = mIterations + 1
    mLastProbeNumericalFailure = False

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

    ' Обычную пробную нагрузку решаем общим способом. Если есть ненулевое
    ' равновесное НДС предыдущей нагрузки, используем его как старт.
    ' Для первого ненулевого шага строим начальное поле по целевой нагрузке:
    ' нулевое поле ещё не определяет работающие участки диаграмм чистого изгиба.
    Dim suppliedStart As Boolean
    If Not startSolver Is Nothing Then
        suppliedStart = startSolver.Converged And _
            Not IsZeroLoadProbe(startSolver.Nint, startSolver.Mxint, startSolver.Myint)
    End If
    If suppliedStart Then
        runner.SolveWithInitialLoadState section, concreteMaterial, steelMaterial, _
            nSpecified, mx, my, startSolver.Epsilon0, startSolver.KappaX, startSolver.KappaY, _
            startSolver.Nint, startSolver.Mxint, startSolver.Myint, allowExtensionWarmStart
    ElseIf Not mLastAcceptableSolver Is Nothing And _
            Not IsZeroLoadProbe(mLastAcceptableN, mLastAcceptableMx, mLastAcceptableMy) Then
        runner.SolveWithInitialLoadState section, concreteMaterial, steelMaterial, _
            nSpecified, mx, my, mLastAcceptableSolver.Epsilon0, _
            mLastAcceptableSolver.KappaX, mLastAcceptableSolver.KappaY, _
            mLastAcceptableN, mLastAcceptableMx, mLastAcceptableMy, allowExtensionWarmStart
    ElseIf IsPureBendingProbe(nSpecified, mx, my) Then
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
        mLastProbeNumericalFailure = (solver.FailureCode = sfcNumericalFailure Or solver.FailureCode = sfcSingularTangent)
        ProbeOnce = ClassifyNumericalFailure(solver.FailureCode)
        mFailureCode = solver.FailureCode
        mFailureSettingKeys = solver.FailureSettingKeys
        mStopReason = solver.StopReason
    Else
        mFailureCode = sfcNone
        mFailureSettingKeys = vbNullString
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

' Передаёт настройки внутреннего равновесия общему запуску НДС одной пробы.
' Повторные попытки меняют только численные параметры;
' проверяемая физическая нагрузка lambda остаётся той же.
Private Sub ConfigureCapacityProbeRunner(ByVal runner As CStateSolutionRunner, ByVal attempt As Long)
    runner.LoadSteps = mSolverBaseLoadSteps * (2 ^ attempt)
    runner.SolverMethod = mSolverMethod
    runner.MaxIterations = mSolverMaxIterations
    runner.ToleranceN = mSolverToleranceN
    runner.ToleranceMx = mSolverToleranceMx
    runner.ToleranceMy = mSolverToleranceMy
    runner.SecantMaxRestarts = mSolverSecantMaxRestarts
    runner.DiagnosticsEnabled = mDiagnosticsEnabled
End Sub

' Показывает, нужна ли для пробы чистого изгиба начальная плоскость
' по целевым моментам, если ненулевое прежнее равновесие ещё не найдено.
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

' Проверяет равновесие нулевого поля и сохраняет нагрузку при lambda = 0.
' Ошибки модели и настроек сохраняются как ошибки; начальная точка
' из-за них не считается автоматически допустимой.
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
        mFailureSettingKeys = solver.FailureSettingKeys
        mStopReason = solver.StopReason
        EvaluateZeroLoadProbe = ClassifyNumericalFailure(mFailureCode)
        Exit Function
    End If
    If Not solver.ConfirmEquilibrium(0#, 0#, 0#) Then
        mFailureCode = solver.FailureCode
        mFailureSettingKeys = solver.FailureSettingKeys
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

' Ищет уже оценённую пробу текущего поиска по точному совпадению lambda и усилий.
' Возвращает её критерий и начальное НДС без нового решения равновесия.
' Эти промежуточные пробы в общее хранилище конечных состояний не записываются.
Private Function TryRestoreProbeCache(ByVal nSpecified As Double, ByVal mx As Double, ByVal my As Double, _
        ByVal lambdaValue As Double, ByRef probeState As String) As Boolean
    Dim slot As Long
    For slot = 1 To mProbeCacheCount
        If ProbeCacheMatches(slot, nSpecified, mx, my, lambdaValue) Then
            probeState = mProbeCacheState(slot)
            Set mLastSolver = mProbeCacheSolver(slot)
            mFailureCode = sfcNone
            mFailureSettingKeys = vbNullString
            mLastProbeNumericalFailure = False
            mCriticalStrain = mProbeCacheCriticalStrain(slot)
            mCriticalElement = mProbeCacheCriticalElement(slot)
            mCriticalLimitType = mProbeCacheCriticalType(slot)
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
    mProbeCacheCriticalType(slot) = mCriticalLimitType
    mProbeCacheCriticalUtil(slot) = mCriticalStrainUtilization
    Set mProbeCacheSolver(slot) = mLastSolver

    If mProbeCacheCount < PROBE_CACHE_SIZE Then mProbeCacheCount = mProbeCacheCount + 1
    mProbeCacheNextSlot = slot + 1 ' Кольцевой буфер ограниченного размера заменяет самую старую пробу текущего поиска.
    If mProbeCacheNextSlot > PROBE_CACHE_SIZE Then mProbeCacheNextSlot = 1
End Sub

' Очищает все сохранённые оценки пробных нагрузок перед независимым расчётом.
Private Sub ClearProbeCache()
    Dim slot As Long
    For slot = 1 To PROBE_CACHE_SIZE
        Set mProbeCacheSolver(slot) = Nothing
        mProbeCacheState(slot) = vbNullString
        mProbeCacheCriticalElement(slot) = vbNullString
        mProbeCacheCriticalType(slot) = vbNullString
    Next slot
    mProbeCacheCount = 0
    mProbeCacheNextSlot = 1
End Sub

' Обновляет управляющую деформацию для последней lambda-точки. Проверяются
' бетон, растянутый бетон при включенном лимите и арматура; максимальная
' utilization определяет тип предельного состояния.
Private Sub UpdateCriticalState(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal solver As CSectionSolver)
    Dim limits As CSectionStrainLimits
    Set limits = New CSectionStrainLimits
    limits.Configure mConcreteCompressionLimit, mConcreteTensionLimit, _
        mSteelCompressionLimit, mSteelTensionLimit, mConcreteTensionLimitEnabled
    limits.Evaluate section, solver
    mCriticalStrainUtilization = limits.Utilization
    mCriticalStrain = limits.CriticalStrain
    mCriticalElement = limits.CriticalElement
    mCriticalLimitType = limits.LimitState
End Sub

' Возвращает предел арматуры для знака проверяемой деформации.
' В симметричной контрольной модели пределы сжатия и растяжения одинаковы.
Private Function SteelLimitForStrain(ByVal strain As Double) As Double
    If strain < 0# Then
        SteelLimitForStrain = mSteelCompressionLimit
    Else
        SteelLimitForStrain = mSteelTensionLimit
    End If
End Function

' Передает тип общего критерия, не распознавая материал по тексту ID.
Private Function CriticalLimitState() As String
    If Len(mCriticalLimitType) = 0 Then
        CriticalLimitState = STATE_NUMERICAL_FAILURE
    Else
        CriticalLimitState = mCriticalLimitType
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

' Сохраняет допустимую нагрузку и её НДС для следующих стартов и приёма
' уточнённого физического предела. Если верхний предел ещё не подтверждён,
' эта точка остаётся начальным приближением и диагностикой.
Private Sub RecordAcceptableState(ByVal solver As CSectionSolver, ByVal nValue As Double, ByVal mxValue As Double, ByVal myValue As Double)
    Set mLastAcceptableSolver = solver
    mLastAcceptableN = nValue
    mLastAcceptableMx = mxValue
    mLastAcceptableMy = myValue
End Sub

' Уменьшает верхнее lambda после численной неудачи. Если интервал слишком мал,
' сохраняет последнее найденное равновесие для диагностики.
' Физический предел не объявляется установленным; поиск остаётся неуспешным.
Private Function ReduceLambdaAfterNumericalFailure(ByVal lowLambda As Double, ByVal highLambda As Double, ByVal probeState As String) As Double
    Dim reduced As Double
    reduced = lowLambda + (highLambda - lowLambda) / 2#
    If lowLambda > 0# And (reduced - lowLambda) <= mLambdaTolerance Then
        mLimitState = probeState
        mStopReason = "Не удалось найти равновесие до определения верхней границы поиска. Сохранён последний множитель нагрузки lambda с подтверждённым равновесием; считать его несущей способностью нельзя."
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

' Проверяет, допускает ли численная неудача повторное решение,
' уменьшение lambda или переход к другому способу поиска.
Private Function IsNumericalFailure(ByVal state As String) As Boolean
    IsNumericalFailure = (state = STATE_NUMERICAL_FAILURE Or state = STATE_SINGULAR_TANGENT)
End Function

' Проверяет выбранный пользователем метод одномерного поиска Capacity.SearchMethod.
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

' Проверяет пригодность числа для дальнейшей арифметики поиска.
' Граница 1e100 защищает промежуточные произведения от переполнения
' и не является физическим пределом материала.
Private Function IsFiniteNumber(ByVal value As Double) As Boolean
    If value <> value Then Exit Function
    If Abs(value) > 1E+100 Then Exit Function
    IsFiniteNumber = True
End Function

' Записывает диагностику пробы LoadMultiplier. attempt показывает попытку
' НДС той же нагрузки: 0 — первая, 1 — первый повтор
' с изменёнными численными настройками.
Private Sub AppendDiagnostic(ByVal solver As CSectionSolver, ByVal mx As Double, ByVal my As Double, _
        ByVal lambdaValue As Double, ByVal attempt As Long, ByVal probeState As String)
    If Not mDiagnosticsEnabled Then Exit Sub
    AppendDiagnosticLine "probe=" & CStr(mIterations) & _
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
    AppendDiagnosticLine "search=" & message & vbCrLf
End Sub

' Строки сохраняются без повторного копирования всего журнала на каждой пробе.
Private Sub AppendDiagnosticLine(ByVal message As String)
    If mDiagnosticCount = mDiagnosticCapacity Then
        If mDiagnosticCapacity = 0 Then
            mDiagnosticCapacity = 256
        Else
            mDiagnosticCapacity = mDiagnosticCapacity * 2
        End If
        ReDim Preserve mDiagnosticLines(0 To mDiagnosticCapacity - 1)
    End If
    mDiagnosticLines(mDiagnosticCount) = message
    mDiagnosticCount = mDiagnosticCount + 1
    mDiagnosticDirty = True
End Sub

' Заменяет накопленный диагностический журнал подготовленным текстом без повторного решения.
Private Sub ReplaceDiagnosticLog(ByVal text As String)
    mDiagnosticLog = text: mDiagnosticJoined = text: mDiagnosticDirty = False
    Erase mDiagnosticLines
    mDiagnosticCount = 0: mDiagnosticCapacity = 0
End Sub

' Возвращает накопленный журнал поиска для сохранения в результате.
Private Function DiagnosticText() As String
    If mDiagnosticDirty Then
        Dim lines() As String, i As Long
        ReDim lines(0 To mDiagnosticCount - 1)
        For i = 0 To mDiagnosticCount - 1
            lines(i) = mDiagnosticLines(i)
        Next i
        mDiagnosticJoined = mDiagnosticLog & Join(lines, vbNullString)
        mDiagnosticDirty = False
    End If
    DiagnosticText = mDiagnosticJoined
End Function

' Очищает результаты предыдущего поиска, сохраняя численные настройки
' и предельные деформации материалов.
Private Sub ClearResult()
    mConverged = False
    mMeritJacobianForAffinePath = False
    mStopReason = vbNullString
    mLimitState = vbNullString
    mUpperLimitState = vbNullString
    mInitialLambdaRejected = False
    mFailureCode = sfcNone
    mFailureSettingKeys = vbNullString
    mSearchBoundReached = False
    mLastProbeNumericalFailure = False
    mLambdaUltimate = 0#
    mNUltimate = 0#
    mMomentUltimate = 0#
    mMxUltimate = 0#: mMyUltimate = 0#
    mSafetyFactor = 0#
    mUtilization = 0#
    mIterations = 0
    mSearchExecuted = False
    mRetryCount = 0
    mMomentEquilibriumResidual = 0#
    mCriticalStrain = 0#
    mCriticalElement = vbNullString
    mCriticalLimitType = vbNullString
    mCriticalStrainUtilization = 0#
    mLastAcceptableN = 0#: mLastAcceptableMx = 0#: mLastAcceptableMy = 0#
    Set mLastSolver = Nothing
    Set mLastAcceptableSolver = Nothing
    Set mProbeRunner = Nothing
    ReplaceDiagnosticLog vbNullString
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

' Проверяет, находится ли НДС пробной нагрузки внутри физических пределов материалов.
Public Function LimitSearchCapacityStateIsAcceptable(ByVal state As String) As Boolean
    LimitSearchCapacityStateIsAcceptable = (state = STATE_ACCEPTABLE)
End Function

' Проверяет, что состояние lambda-точки является физическим пределом, а не
' численной неудачей или ошибкой исходных данных.
Public Function LimitSearchCapacityStateIsPhysicalLimit(ByVal state As String) As Boolean
    LimitSearchCapacityStateIsPhysicalLimit = IsPhysicalLimit(state)
End Function

' Проверяет, можно ли после численной неудачи пробы уменьшить lambda
' или перейти к дополнительному методу поиска.
Public Function LimitSearchCapacityStateIsNumericalFailure(ByVal state As String) As Boolean
    LimitSearchCapacityStateIsNumericalFailure = IsNumericalFailure(state)
End Function

' Проверяет, вызвана ли остановка ошибкой исходных данных.
' При такой причине координатор не запускает дополнительный LoadMultiplier.
Public Function LimitSearchCapacityStateIsInvalidInput(ByVal state As String) As Boolean
    LimitSearchCapacityStateIsInvalidInput = (state = STATE_INVALID_INPUT)
End Function

' Терминальная ошибка данных, настройки или программной зависимости запрещает
' повтор solve и переход к другому численному методу той же постановки.
Public Property Get LimitSearchHasTerminalFailure() As Boolean
    LimitSearchHasTerminalFailure = (mFailureCode = sfcInvalidInput Or _
        mFailureCode = sfcInvalidConfiguration Or mFailureCode = sfcInternalError)
End Property

' Передаёт точные InternalStatus и ResultCode по признакам результата.
' Слова причины остановки и внешнего статуса на их выбор не влияют.
Public Property Get FailureCode() As ESolverFailureCode
    FailureCode = mFailureCode
End Property

' Сохраняет связь причины с настройками отдельно от её формулировки.
Public Property Get FailureSettingKeys() As String
    FailureSettingKeys = mFailureSettingKeys
End Property

' Показывает, отклонена ли уже постоянная нагрузка при lambda = 0.
' Общий численный поиск сообщает исход начальной точки;
' инженерный смысл BaseFail определяет результат несущей способности.
Public Property Get LimitSearchInitialLambdaRejected() As Boolean
    LimitSearchInitialLambdaRejected = mInitialLambdaRejected
End Property

' Отличает техническую остановку MaxLambda от несходимости пробного равновесия.
Public Property Get SearchBoundReached() As Boolean
    SearchBoundReached = mSearchBoundReached
End Property

' Отдает факт запуска независимо от успеха, числа завершенных итераций
' и наличия точки. BuildSnapshot сохраняет его без анализа диагностического текста.
Public Property Get SearchExecuted() As Boolean
    SearchExecuted = mSearchExecuted
End Property

' Объединяет диагностику первой неудачной попытки и последующего LoadMultiplier.
' Порядок методов выбирает общий поиск; внутренний журнал и причину
' остановки изменяет сам CCapacitySolver.
Public Sub LimitSearchMergeCapacityFallback(ByVal firstLimitState As String, _
        ByVal firstStopReason As String, ByVal firstDiagnostics As String, _
        ByVal firstAttemptMessage As String, ByVal transitionMessage As String, _
        ByVal successMessage As String, Optional ByVal firstSearchExecuted As Boolean = False)
    mSearchExecuted = mSearchExecuted Or firstSearchExecuted
    If mDiagnosticsEnabled Then
        Dim mergedDiagnostics As String
        mergedDiagnostics = firstAttemptMessage & "; статус=" & firstLimitState & _
            "; причина=" & firstStopReason & vbCrLf
        If Len(firstDiagnostics) > 0 Then mergedDiagnostics = mergedDiagnostics & firstDiagnostics
        mergedDiagnostics = mergedDiagnostics & transitionMessage & vbCrLf & DiagnosticLog
        ReplaceDiagnosticLog mergedDiagnostics
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

' Отказ равновесия, не исчерпание допуска/бюджета одномерного поиска.
Public Property Get LimitSearchHasFailedEquilibriumProbe() As Boolean
    LimitSearchHasFailedEquilibriumProbe = mLastProbeNumericalFailure
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

' Показывает, найдено ли подтверждённое предельное состояние выбранным методом.
Public Property Get Converged() As Boolean
    Converged = mConverged
End Property

' Возвращает причину остановки поиска несущей способности для вывода и тестов.
Public Property Get StopReason() As String
    StopReason = mStopReason
End Property

' Физический или численный статус найденного предела.
Public Property Get LimitState() As String
    LimitState = mLimitState
End Property

' Статус верхней границы при LoadMultiplier.
Public Property Get UpperLimitState() As String
    UpperLimitState = mUpperLimitState
End Property

' Предельный множитель к выбранной lambda-траектории нагрузки.
Public Property Get LambdaUltimate() As Double
    LambdaUltimate = mLambdaUltimate
End Property

' Предельная продольная сила, если N входит в масштабируемый путь.
Public Property Get NUltimate() As Double
    NUltimate = mNUltimate
End Property

' Результирующий предельный момент sqrt(MxUltimate^2 + MyUltimate^2).
Public Property Get MomentUltimate() As Double
    MomentUltimate = mMomentUltimate
End Property

' Предельный Mx, если Mx входит в масштабируемый путь.
Public Property Get MxUltimate() As Double
    MxUltimate = mMxUltimate
End Property

' Предельный My, если My входит в масштабируемый путь.
Public Property Get MyUltimate() As Double
    MyUltimate = mMyUltimate
End Property

' Возвращает коэффициент запаса несущей способности, равный найденному lambdaUltimate.
Public Property Get SafetyFactor() As Double
    SafetyFactor = mSafetyFactor
End Property

' Использование несущей способности: 1 / lambdaUltimate.
Public Property Get Utilization() As Double
    Utilization = mUtilization
End Property

' Число внешних итераций поиска несущей способности.
Public Property Get Iterations() As Long
    Iterations = mIterations
End Property

' Невязка направления/момента в найденном предельном состоянии.
Public Property Get MomentEquilibriumResidual() As Double
    MomentEquilibriumResidual = mMomentEquilibriumResidual
End Property

' Накопленный диагностический журнал текущего расчёта или поиска.
Public Property Get DiagnosticLog() As String
    DiagnosticLog = DiagnosticText()
End Property

' Отключает подробный журнал в массовом пользовательском расчете без txt-отчета.
Public Property Let DiagnosticsEnabled(ByVal value As Boolean)
    mDiagnosticsEnabled = value
End Property

' Последний внутренний расчет равновесия, успешный или диагностический.
Public Property Get LastSolver() As CSectionSolver
    Set LastSolver = mLastSolver
End Property

' Предельная сжимающая деформация бетона.
Public Property Let ConcreteCompressionLimit(ByVal value As Double)
    mConcreteCompressionLimit = value
End Property

' Предельная растягивающая деформация бетона, если растянутая ветвь включена материалной
' моделью.
Public Property Let ConcreteTensionLimit(ByVal value As Double)
    mConcreteTensionLimit = value
End Property

' True, когда растянутый бетон реально ограничивает предельное состояние.
Public Property Let ConcreteTensionLimitEnabled(ByVal value As Boolean)
    mConcreteTensionLimitEnabled = value
End Property

' Задаёт один и тот же модуль предельной деформации для сжатия и растяжения арматуры.
' Для независимых пределов используются SteelCompressionLimit и SteelTensionLimit.
Public Property Let SteelStrainLimit(ByVal value As Double)
    mSteelCompressionLimit = Abs(value)
    mSteelTensionLimit = Abs(value)
End Property

' Задает независимые физические пределы арматуры по двум знакам деформации.
Public Property Let SteelCompressionLimit(ByVal value As Double)
    mSteelCompressionLimit = Abs(value)
End Property

' Предельная растягивающая деформация арматуры.
Public Property Let SteelTensionLimit(ByVal value As Double)
    mSteelTensionLimit = Abs(value)
End Property

' Точность одномерного поиска lambda для LoadMultiplier.
Public Property Let LambdaTolerance(ByVal value As Double)
    mLambdaTolerance = value
End Property

' Допуск достижения предельной деформации для UltimateStrain.
Public Property Let StrainTolerance(ByVal value As Double)
    mStrainTolerance = value
End Property

' Алгоритм одномерного поиска LoadMultiplier: Bisection, Brent или Secant.
Public Property Let SearchMethod(ByVal value As String)
    mSearchMethod = value
End Property

' Алгоритм одномерного поиска LoadMultiplier: Bisection, Brent или Secant.
Public Property Get SearchMethod() As String
    SearchMethod = mSearchMethod
End Property

' Фактическая ветка, которая дала принятый результат: UltimateStrain или LoadMultiplier.
Public Property Get SolutionMethod() As String
    SolutionMethod = mSolutionMethod
End Property

' Верхняя граница расширения скобки LoadMultiplier, чтобы поиск не уходил бесконечно.
Public Property Let MaxLambda(ByVal value As Double)
    mMaxLambda = value
End Property

' Первый шаг расширения lambda при поиске расчетной скобки.
Public Property Let InitialLambdaStep(ByVal value As Double)
    mInitialLambdaStep = value
End Property

' Допустимые повторы той же lambda-точки после NumericalFailure.
Public Property Let MaxRetries(ByVal value As Long)
    mMaxRetries = value
End Property

' Базовое число ступеней CSectionSolver для первой попытки lambda-точки.
Public Property Let SolverBaseLoadSteps(ByVal value As Long)
    mSolverBaseLoadSteps = value
End Property

' Задаёт максимум итераций внутреннего CSectionSolver на одной ступени нагрузки.
' Бюджет внешнего поиска предела определяется отдельными настройками.
Public Property Let SolverMaxIterations(ByVal value As Long)
    mSolverMaxIterations = value
End Property

' Допуск внутреннего равновесия CSectionSolver по N.
Public Property Let SolverToleranceN(ByVal value As Double)
    mSolverToleranceN = value
End Property

' Допуск внутреннего равновесия CSectionSolver по Mx.
Public Property Let SolverToleranceMx(ByVal value As Double)
    mSolverToleranceMx = value
End Property

' Допуск внутреннего равновесия CSectionSolver по My.
Public Property Let SolverToleranceMy(ByVal value As Double)
    mSolverToleranceMy = value
End Property

' Метод внутреннего CSectionSolver: Newton или Secant.
Public Property Let SolverMethod(ByVal value As String)
    mSolverMethod = value
End Property

' Настройка рестартов внутреннего secant-решателя.
Public Property Let SolverSecantMaxRestarts(ByVal value As Long)
    mSolverSecantMaxRestarts = value
End Property

' ==========================================================================
' ДЛЯ ТЕСТОВ
' ==========================================================================

' -------------------- ДЛЯ ТЕСТОВ --------------------
' возвращает подробности поиска для проверки критического элемента,
' повторных решений и коэффициента использования. Пользовательский вывод
' использует готовые SafetyFactor, LimitState, MomentUltimate и статусы сочетания.
Public Property Get TestRetryCount() As Long
    TestRetryCount = mRetryCount
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' деформация критического волокна или стержня.
Public Property Get TestCriticalStrain() As Double
    TestCriticalStrain = mCriticalStrain
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' идентификатор элемента, который первым определил предельное состояние.
Public Property Get TestCriticalElement() As String
    TestCriticalElement = mCriticalElement
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' отношение фактической деформации к своему пределу.
Public Property Get TestCriticalStrainUtilization() As Double
    TestCriticalStrainUtilization = mCriticalStrainUtilization
End Property
