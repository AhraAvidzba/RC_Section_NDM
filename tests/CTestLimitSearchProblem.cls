VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTestLimitSearchProblem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' ДЛЯ ТЕСТОВ: простая доменная задача общего LimitSearch
' ==========================================================================
' Класс намеренно не использует CSectionSolver и реальные материалы. Он задает
' линейную функцию предела f(lambda) = lambda - lambdaRoot и позволяет проверять
' поведение общего search-ядра без влияния железобетонной физики.
'
' Такой тест защищает архитектурное требование: Bisection/Secant/Brent должны
' работать через ILimitSearchProblem, а не через конкретный Capacity или Crack
' calculator.

Implements ILimitSearchProblem

Private Const PROBE_ACCEPTABLE As String = "Acceptable" ' Точка находится до предела.
Private Const PROBE_LIMIT As String = "Limit" ' Точка достигла или превысила предел.
Private Const PROBE_NUMFAIL As String = "NumericalFailure" ' Искусственная численная ошибка.
Private Const PROBE_INPUTERR As String = "InvalidInput" ' Искусственная ошибка входных данных.

Private mSearchMethod As String ' Проверяемый одномерный метод search.
Private mRootLambda As Double ' Точный корень линейной тестовой функции.
Private mTolerance As Double ' Допуск lambda-скобки, передаваемый общему search.
Private mMaxIterations As Long ' Лимит итераций для проверки T02.
Private mInitialStep As Double ' Начальный шаг расширения скобки.
Private mMaxLambda As Double ' Техническая верхняя граница lambda.
Private mKind As EResultKind ' Вид результата, через который запускается общий search.
Private mConverged As Boolean ' True, если search финализировал точку.
Private mFinalLambda As Double ' Lambda, фактически принятая общим search.
Private mFinalMethod As String ' Метод, которым была принята итоговая точка.
Private mFinalIterations As Long ' Число итераций, переданное при финализации.
Private mLimitState As String ' Последний машинный статус fake-задачи.
Private mDiagnosticLog As String ' Накопленный диагностический текст.
Private mIterationLimitHit As Boolean ' True, если search остановился по лимиту итераций.
Private mProbeMaxLambda As Double ' Наибольшая реально проверенная точка.
Private mProbeCalls As Long ' Счетчик проб общего алгоритма.
Private mFailureAbove As Double ' Искусственная численная неудача выше этой lambda; 0 отключает ее.
Private mStepCriterion As Boolean ' Разрывный критерий не имеет нулевой невязки на соседних Double.
Private mRecoveryWithoutProgress As Boolean ' Намеренно неисправный callback recovery возвращает прежнюю границу.
Private mUltimateDamping As Double ' ДЛЯ ТЕСТОВ: начальное демпфирование Newton.
Private mUltimateMinAlpha As Double ' ДЛЯ ТЕСТОВ: нижняя граница line search.
Private mUltimateMaxDeltaEps As Double ' ДЛЯ ТЕСТОВ: проверка непредставимого приращения.
Private mUltimateInitialEps As Double ' ДЛЯ ТЕСТОВ: ненулевой старт для машинной стагнации.
Private mUltimateConstantNorm As Boolean ' ДЛЯ ТЕСТОВ: line search не может уменьшить норму.
Private mUltimateScenario As Long ' ДЛЯ ТЕСТОВ: направленный отказ callback-а.
Private mUltimateCalls As Long ' ДЛЯ ТЕСТОВ: число оценок невязки Newton.
Private mFailureCode As ESolverFailureCode ' ДЛЯ ТЕСТОВ: точная причина остановки.

' Настраивает линейную тестовую задачу перед запуском общего search.
Public Sub Configure(ByVal searchMethod As String, ByVal rootLambda As Double, _
        ByVal tolerance As Double, ByVal maxIterations As Long, _
        Optional ByVal initialStep As Double = 1#, Optional ByVal maxLambda As Double = 64#, _
        Optional ByVal resultKind As EResultKind = rkCapacity)
    mSearchMethod = searchMethod
    mRootLambda = rootLambda
    mTolerance = tolerance
    mMaxIterations = maxIterations
    mInitialStep = initialStep
    mMaxLambda = maxLambda
    mKind = resultKind
    mConverged = False
    mFinalLambda = 0#
    mFinalMethod = vbNullString
    mFinalIterations = 0
    mLimitState = vbNullString
    mDiagnosticLog = vbNullString
    mIterationLimitHit = False
    mProbeMaxLambda = 0#
    mProbeCalls = 0
    mFailureAbove = 0#
    mStepCriterion = False
    mRecoveryWithoutProgress = False
    mUltimateDamping = 1#
    mUltimateMinAlpha = 0.03125
    mUltimateMaxDeltaEps = 0.01
    mUltimateInitialEps = 0#
    mUltimateConstantNorm = False
    mUltimateScenario = 0
    mUltimateCalls = 0
    mFailureCode = sfcNone
End Sub

' ДЛЯ ТЕСТОВ: задает численную неудачу верхних проб, не физический предел.
Public Property Let FailureAbove(ByVal value As Double)
    mFailureAbove = value
End Property

' ДЛЯ ТЕСТОВ: не позволяет признать соседнюю граничную lambda точным корнем.
Public Property Let StepCriterion(ByVal value As Boolean)
    mStepCriterion = value
End Property

' ДЛЯ ТЕСТОВ: общий Search обязан остановить recovery, не сужающий скобку.
Public Property Let RecoveryWithoutProgress(ByVal value As Boolean)
    mRecoveryWithoutProgress = value
End Property

' ДЛЯ ТЕСТОВ: диагностические счетчики расширения и восстановления скобки.
Public Property Get ProbeMaxLambda() As Double
    ProbeMaxLambda = mProbeMaxLambda
End Property

Public Property Get ProbeCalls() As Long
    ProbeCalls = mProbeCalls
End Property

Public Property Get FinalLambda() As Double
    FinalLambda = mFinalLambda
End Property

Public Property Get FinalMethod() As String
    FinalMethod = mFinalMethod
End Property

Public Property Get FinalIterations() As Long
    FinalIterations = mFinalIterations
End Property

Public Property Get IterationLimitHit() As Boolean
    IterationLimitHit = mIterationLimitHit
End Property

Public Property Get Converged() As Boolean
    Converged = mConverged
End Property

Public Property Get LimitState() As String
    LimitState = mLimitState
End Property

Public Property Get DiagnosticLog() As String
    DiagnosticLog = mDiagnosticLog
End Property

Private Property Get ILimitSearchProblem_ResultKind() As EResultKind
    ILimitSearchProblem_ResultKind = mKind
End Property

Private Property Get ILimitSearchProblem_InitialLambdaStep() As Double
    ILimitSearchProblem_InitialLambdaStep = mInitialStep
End Property

Private Property Get ILimitSearchProblem_MaxLambda() As Double
    ILimitSearchProblem_MaxLambda = mMaxLambda
End Property

Private Property Get ILimitSearchProblem_LambdaTolerance() As Double
    ILimitSearchProblem_LambdaTolerance = mTolerance
End Property

Private Property Get ILimitSearchProblem_MaxIterations() As Long
    ILimitSearchProblem_MaxIterations = mMaxIterations
End Property

' ДЛЯ ТЕСТОВ: позволяет явно проверить исчерпание бюджета Bisection.
Private Property Get ILimitSearchProblem_BisectionMaxIterations() As Long
    ILimitSearchProblem_BisectionMaxIterations = mMaxIterations
End Property

Private Property Get ILimitSearchProblem_SearchMethod() As String
    ILimitSearchProblem_SearchMethod = mSearchMethod
End Property

Private Property Get ILimitSearchProblem_Converged() As Boolean
    ILimitSearchProblem_Converged = mConverged
End Property

Private Property Get ILimitSearchProblem_LimitState() As String
    ILimitSearchProblem_LimitState = mLimitState
End Property

Private Property Get ILimitSearchProblem_DiagnosticLog() As String
    ILimitSearchProblem_DiagnosticLog = mDiagnosticLog
End Property

' ДЛЯ ТЕСТОВ: общий сценарий выбирает путь, не добавляя доменную семантику
' к управляемому probe callback-у численных контрактных тестов.
Private Sub ILimitSearchProblem_BeginLoadPath(ByVal request As CLimitSearchRequest)
End Sub

' ДЛЯ ТЕСТОВ: история попыток хранится в настоящем CLimitSearchResult.
Private Sub ILimitSearchProblem_FinishLoadPath(ByVal request As CLimitSearchRequest, _
        ByVal previousPathComments As String)
End Sub

' ДЛЯ ТЕСТОВ: начало линейной функции должно лежать до заданного корня.
Private Function ILimitSearchProblem_AcceptInitialPoint(ByVal request As CLimitSearchRequest, _
        ByVal loadPath As CLoadPathVector, ByVal functionValue As Double, _
        ByVal probeState As String, ByVal solver As CSectionSolver) As Boolean
    If probeState <> PROBE_ACCEPTABLE Then
        ILimitSearchProblem_HandleInitialFailure probeState
        Exit Function
    End If
    ILimitSearchProblem_RecordAcceptable loadPath, 0#
    ILimitSearchProblem_AcceptInitialPoint = True
End Function

' Fake-задача не требует доменной подготовки перед поиском.
Private Function ILimitSearchProblem_Prepare(ByVal request As CLimitSearchRequest) As Boolean
    ILimitSearchProblem_Prepare = True
End Function

' ДЛЯ ТЕСТОВ: аналитическая система из трех независимых невязок позволяет
' проверить общий Newton без Capacity/Crack-калькуляторов и material API.
Private Function ILimitSearchProblem_PrepareUltimate(ByVal request As CLimitSearchRequest, _
        ByVal loadPath As CLoadPathVector) As Boolean
    ILimitSearchProblem_PrepareUltimate = Not loadPath Is Nothing
End Function

' ДЛЯ ТЕСТОВ: запускает настоящий общий Newton на аналитических callbacks.
' Section и материалы не нужны простой тестовой системе трех невязок.
Private Function ILimitSearchProblem_RunUltimateAttempts(ByVal request As CLimitSearchRequest) As Boolean
    Dim search As CUltimateStrainSearch
    Set search = New CUltimateStrainSearch
    ILimitSearchProblem_RunUltimateAttempts = search.RunNewton(Me, Nothing, Nothing, Nothing, _
        ILimitSearchProblem_BuildLoadPath(request), Nothing, "Тестовый поиск не сошелся.")
End Function

' ДЛЯ ТЕСТОВ: финальная lambda уже сохранена callback-ом; дополнительной
' инженерной обработки результата аналитической fixture нет.
Private Sub ILimitSearchProblem_CompleteUltimate(ByVal loadPath As CLoadPathVector)
End Sub

' ДЛЯ ТЕСТОВ: задает управляемый старт по epsilon0 и нулевые кривизны,
' чтобы воспроизводимо проверять нормальную сходимость и стагнацию общего Newton.
Private Sub ILimitSearchProblem_LimitSearchInitializeUltimateGuess(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal loadPath As CLoadPathVector, ByVal startSolver As CSectionSolver, _
        ByRef eps0 As Double, ByRef kx As Double, ByRef ky As Double)
    eps0 = mUltimateInitialEps: kx = 0#: ky = 0#
End Sub

' ДЛЯ ТЕСТОВ: вычисляет линейные невязки и считает реальные вызовы callback.
' Некоторые сценарии намеренно дают typed failure либо VBA error на второй
' пробе, чтобы проверить причину остановки на шаге построения Якобиана.
Private Function ILimitSearchProblem_LimitSearchEvaluateUltimateResidual(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal loadPath As CLoadPathVector, ByVal eps0 As Double, _
        ByVal kx As Double, ByVal ky As Double, ByRef solver As CSectionSolver, _
        ByRef rPath1 As Double, ByRef rPath2 As Double, ByRef rLimit As Double, _
        ByRef lambdaValue As Double, ByRef state As String) As Boolean
    Set solver = Nothing
    mUltimateCalls = mUltimateCalls + 1
    If mUltimateCalls = 2 Then
        Select Case mUltimateScenario
            Case 7
                ILimitSearchProblem_SetFailure sfcInvalidConfiguration, "Ошибочная настройка тестового callback-а Якобиана."
                Exit Function
            Case 8
                ILimitSearchProblem_SetFailure sfcInternalError, "Тестовый callback не получил обязательную зависимость."
                Exit Function
            Case 9
                ILimitSearchProblem_SetFailure sfcSingularTangent, "Вырождение тестовой касательной."
                Exit Function
            Case 10
                Err.Raise 9, "CTestLimitSearchProblem", "Тестовое нарушение программного контракта."
            Case 11
                Err.Raise 6, "CTestLimitSearchProblem", "Тестовое переполнение численной операции."
        End Select
    End If
    rPath1 = kx
    rPath2 = ky
    lambdaValue = eps0 / 0.001
    rLimit = lambdaValue - mRootLambda
    state = PROBE_LIMIT
    ILimitSearchProblem_LimitSearchEvaluateUltimateResidual = True
End Function

' ДЛЯ ТЕСТОВ: проверяет три допуска аналитической системы. Режим constant norm
' намеренно запрещает успех, создавая воспроизводимую стагнацию line search.
Private Function ILimitSearchProblem_LimitSearchUltimateResidualReached(ByVal rPath1 As Double, _
        ByVal rPath2 As Double, ByVal rLimit As Double) As Boolean
    ILimitSearchProblem_LimitSearchUltimateResidualReached = _
        Abs(rPath1) <= mTolerance And Abs(rPath2) <= mTolerance And Abs(rLimit) <= mTolerance
    If mUltimateConstantNorm Then ILimitSearchProblem_LimitSearchUltimateResidualReached = False
End Function

' ДЛЯ ТЕСТОВ: возвращает сумму модулей невязок либо постоянную норму,
' когда нужно проверить отсутствие улучшающего представимого шага.
Private Function ILimitSearchProblem_LimitSearchUltimateResidualNorm(ByVal rPath1 As Double, _
        ByVal rPath2 As Double, ByVal rLimit As Double) As Double
    ILimitSearchProblem_LimitSearchUltimateResidualNorm = Abs(rPath1) + Abs(rPath2) + Abs(rLimit)
    If mUltimateConstantNorm Then ILimitSearchProblem_LimitSearchUltimateResidualNorm = 1#
End Function

' ДЛЯ ТЕСТОВ: сохраняет найденную аналитическую lambda, метод и число итераций.
' Физического state нет: приемка production-предела проверяется отдельными tests.
Private Function ILimitSearchProblem_LimitSearchFinalizeUltimateResult(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal loadPath As CLoadPathVector, ByVal lambdaValue As Double, _
        ByVal solver As CSectionSolver, ByVal state As String, ByVal iterations As Long) As Boolean
    mConverged = True
    mFinalLambda = lambdaValue
    mFinalMethod = "UltimateStrain"
    mFinalIterations = iterations
    mLimitState = state
    ILimitSearchProblem_LimitSearchFinalizeUltimateResult = True
End Function

' ДЛЯ ТЕСТОВ: фиксирует численную остановку и ее причину без найденной точки.
Private Sub ILimitSearchProblem_LimitSearchSetUltimateNumericalFailure(ByVal reason As String)
    mConverged = False
    mLimitState = PROBE_NUMFAIL
    mFailureCode = sfcNumericalFailure
    AppendDiagnosticLine reason
End Sub

' ДЛЯ ТЕСТОВ: одинаковые численные ограничения простой линейной системы.
Private Property Get ILimitSearchProblem_LimitSearchUltimateMaxIterations() As Long
    ILimitSearchProblem_LimitSearchUltimateMaxIterations = mMaxIterations
End Property

Private Property Get ILimitSearchProblem_LimitSearchUltimateDampingInitial() As Double
    ILimitSearchProblem_LimitSearchUltimateDampingInitial = mUltimateDamping
End Property

Private Property Get ILimitSearchProblem_LimitSearchUltimateMinLineSearchAlpha() As Double
    ILimitSearchProblem_LimitSearchUltimateMinLineSearchAlpha = mUltimateMinAlpha
End Property

Private Property Get ILimitSearchProblem_LimitSearchUltimateLineSearchEnabled() As Boolean
    ILimitSearchProblem_LimitSearchUltimateLineSearchEnabled = True
End Property

Private Property Get ILimitSearchProblem_LimitSearchUltimateMaxDeltaEpsilon0() As Double
    ILimitSearchProblem_LimitSearchUltimateMaxDeltaEpsilon0 = mUltimateMaxDeltaEps
End Property

Private Property Get ILimitSearchProblem_LimitSearchUltimateMaxDeltaKappa() As Double
    ILimitSearchProblem_LimitSearchUltimateMaxDeltaKappa = 0.01
End Property

' Создает универсальную траекторию Offset + lambda*Base из request.
Private Function ILimitSearchProblem_BuildLoadPath(ByVal request As CLimitSearchRequest) As CLoadPathVector
    Dim loadPath As CLoadPathVector
    Set loadPath = New CLoadPathVector
    loadPath.Initialize request.NOffset, request.NBase, _
        request.MxOffset, request.MxBase, request.MyOffset, request.MyBase
    Set ILimitSearchProblem_BuildLoadPath = loadPath
End Function

' Оценивает одну lambda-точку линейной функции.
' Точка с нулевой невязкой достигла критерия; доменная финализация принимает
' аналитический корень без требования создать вымышленное материальное State.
Private Function ILimitSearchProblem_EvaluatePoint(ByVal request As CLimitSearchRequest, _
        ByVal loadPath As CLoadPathVector, ByVal startSolver As CSectionSolver, _
        ByVal lambdaValue As Double, ByRef functionValue As Double, _
        ByRef probeState As String, ByRef solver As CSectionSolver) As Boolean
    functionValue = lambdaValue - mRootLambda
    If mStepCriterion Then
        If lambdaValue < mRootLambda Then functionValue = -1# Else functionValue = 1#
    End If
    mProbeCalls = mProbeCalls + 1
    If lambdaValue > mProbeMaxLambda Then mProbeMaxLambda = lambdaValue
    Set solver = Nothing
    If mFailureAbove > 0# And lambdaValue > mFailureAbove Then
        probeState = PROBE_NUMFAIL
        Exit Function
    End If
    If functionValue < 0# Then
        probeState = PROBE_ACCEPTABLE
    Else
        probeState = PROBE_LIMIT
    End If
    Set solver = Nothing
    ILimitSearchProblem_EvaluatePoint = True
End Function

' ДЛЯ ТЕСТОВ: однотипные предикаты аналитического probe-протокола различают
' допустимую точку, предел, намеренную несходимость и ошибку входа.
Private Function ILimitSearchProblem_IsAcceptableState(ByVal probeState As String) As Boolean
    ILimitSearchProblem_IsAcceptableState = (probeState = PROBE_ACCEPTABLE)
End Function

Private Function ILimitSearchProblem_IsLimitState(ByVal probeState As String) As Boolean
    ILimitSearchProblem_IsLimitState = (probeState = PROBE_LIMIT)
End Function

Private Function ILimitSearchProblem_IsNumericalFailureState(ByVal probeState As String) As Boolean
    ILimitSearchProblem_IsNumericalFailureState = (probeState = PROBE_NUMFAIL)
End Function

Private Function ILimitSearchProblem_IsInvalidInputState(ByVal probeState As String) As Boolean
    ILimitSearchProblem_IsInvalidInputState = (probeState = PROBE_INPUTERR)
End Function

' ДЛЯ ТЕСТОВ: recovery использует середину оставшегося интервала.
Private Function ILimitSearchProblem_ReduceLambdaAfterNumericalFailure(ByVal lowLambda As Double, _
        ByVal highLambda As Double, ByVal probeState As String) As Double
    If mRecoveryWithoutProgress Then
        ILimitSearchProblem_ReduceLambdaAfterNumericalFailure = highLambda
    Else
        ILimitSearchProblem_ReduceLambdaAfterNumericalFailure = lowLambda + 0.5 * (highLambda - lowLambda)
    End If
End Function

' Запоминает последнюю acceptable-точку только в диагностике.
Private Sub ILimitSearchProblem_RecordAcceptable(ByVal loadPath As CLoadPathVector, _
        ByVal lambdaValue As Double)
    AppendDiagnosticLine "acceptable lambda=" & FormatNumberInvariant(lambdaValue)
End Sub

' ДЛЯ ТЕСТОВ: фиксирует остановку начальной пробы, не создавая предельную точку.
Private Sub ILimitSearchProblem_HandleInitialFailure(ByVal probeState As String)
    mLimitState = PROBE_NUMFAIL
    AppendDiagnosticLine "initialFailure=" & probeState
End Sub

' ДЛЯ ТЕСТОВ: отличает техническую границу поиска от найденного корня fixture.
Private Sub ILimitSearchProblem_HandleBracketNotFound(ByVal lowLambda As Double)
    mLimitState = "SearchBoundReached"
    AppendDiagnosticLine "bracketNotFound lowLambda=" & FormatNumberInvariant(lowLambda)
End Sub

' ДЛЯ ТЕСТОВ: сохраняет lambda неуспешной верхней пробы для проверки recovery.
Private Sub ILimitSearchProblem_HandleBracketNumericalFailure(ByVal lowLambda As Double, _
        ByVal failedLambda As Double, ByVal probeState As String)
    mLimitState = PROBE_NUMFAIL
    AppendDiagnosticLine "bracketNumericalFailure failedLambda=" & FormatNumberInvariant(failedLambda)
End Sub

' ДЛЯ ТЕСТОВ: лимит итераций bisection не должен превращаться в найденный предел.
Private Sub ILimitSearchProblem_HandleIterationLimit(ByVal maxIterations As Long, _
        ByVal remainingWidth As Double, ByVal tolerance As Double)
    mConverged = False
    mIterationLimitHit = True
    mLimitState = PROBE_NUMFAIL
    AppendDiagnosticLine "iterationLimit maxIterations=" & CStr(maxIterations) & _
        "; remainingWidth=" & FormatNumberInvariant(remainingWidth) & _
        "; tolerance=" & FormatNumberInvariant(tolerance)
End Sub

' ДЛЯ ТЕСТОВ: неизвестный probe-исход оставляет поиск неуспешным и диагностируемым.
Private Sub ILimitSearchProblem_HandleUnexpectedState(ByVal probeState As String)
    mLimitState = PROBE_NUMFAIL
    AppendDiagnosticLine "unexpectedState=" & probeState
End Sub

' ДЛЯ ТЕСТОВ: принимает ошибку параметров общего поиска без фиктивной сходимости.
Private Sub ILimitSearchProblem_SetInvalidInput(ByVal message As String)
    mLimitState = PROBE_INPUTERR
    AppendDiagnosticLine message
End Sub

' ДЛЯ ТЕСТОВ: принимает численную причину общего поиска и сохраняет ее в журнале.
Private Sub ILimitSearchProblem_SetNumericalFailure(ByVal message As String)
    mLimitState = PROBE_NUMFAIL
    AppendDiagnosticLine message
End Sub

' ДЛЯ ТЕСТОВ: callback журнала позволяет проверить реальный путь общего поиска.
Private Sub ILimitSearchProblem_AppendDiagnostic(ByVal message As String)
    AppendDiagnosticLine message
End Sub

' Принимает итоговую lambda-точку, выбранную общим search-ядром.
Private Function ILimitSearchProblem_FinalizeAt(ByVal request As CLimitSearchRequest, _
        ByVal loadPath As CLoadPathVector, ByVal lambdaValue As Double, _
        ByVal probeState As String, ByVal methodName As String, _
        ByVal solver As CSectionSolver, ByVal iterations As Long) As Boolean
    mConverged = True
    mFinalLambda = lambdaValue
    mFinalMethod = methodName
    mFinalIterations = iterations
    mLimitState = probeState
    AppendDiagnosticLine "final lambda=" & FormatNumberInvariant(lambdaValue) & _
        "; method=" & methodName
    ILimitSearchProblem_FinalizeAt = True
End Function

' ДЛЯ ТЕСТОВ: сохраняет выбранную сторону скобки по kind задачи, чтобы проверить
' generic финализацию Capacity/Formation без отдельного численного алгоритма.
Private Sub ILimitSearchProblem_FinalizeBracket(ByVal request As CLimitSearchRequest, _
        ByVal loadPath As CLoadPathVector, ByVal lowLambda As Double, ByVal lowSolver As CSectionSolver, _
        ByVal highLambda As Double, ByVal highSolver As CSectionSolver, _
        ByVal methodName As String, ByVal iterations As Long)
    Dim lambdaValue As Double
    If mKind = rkCrackFormation Then lambdaValue = highLambda Else lambdaValue = lowLambda
    mConverged = True
    mFinalLambda = lambdaValue
    mFinalMethod = methodName
    mLimitState = PROBE_ACCEPTABLE
    AppendDiagnosticLine "retained lambda=" & FormatNumberInvariant(lambdaValue) & _
        "; method=" & methodName
End Sub

' ДЛЯ ТЕСТОВ: отмечает вызов последней попытки финализации, не объявляя успех.
Private Sub ILimitSearchProblem_TryFinalizeAfterSearchFailure(ByVal request As CLimitSearchRequest)
    AppendDiagnosticLine "retainedCapacity"
End Sub

' ДЛЯ ТЕСТОВ: фиксирует полноценный численный результат без solver/calculator.
Private Function ILimitSearchProblem_BuildResult(ByVal request As CLimitSearchRequest, _
        ByVal searchSucceeded As Boolean) As CLimitSearchResult
    Dim meta As CResultMeta
    Set meta = New CResultMeta
    If mConverged Then
        meta.SetResult rsSuccess, rcCheckPassed, mKind, "Тестовый критерий найден."
    ElseIf mFailureCode <> sfcNone Then
        meta.SetSolverFailure mFailureCode, mKind, mDiagnosticLog, (mUltimateCalls > 0)
    ElseIf mLimitState = "SearchBoundReached" Then
        meta.SetResult rsNumericalFailure, rcSearchBoundReached, mKind, mLimitState
    ElseIf mLimitState = PROBE_INPUTERR Then
        meta.SetResult rsInvalidInput, rcInvalidInput, mKind, mLimitState
    Else
        meta.SetResult rsNumericalFailure, rcNumericalFailure, mKind, mLimitState
    End If
    Dim result As CLimitSearchResult
    Set result = New CLimitSearchResult
    result.Initialize meta, request.SolutionStrategy, mFinalMethod, mConverged, mFinalLambda, _
        mFinalLambda, 0#, 0#, Nothing, mLimitState, mDiagnosticLog
    Set ILimitSearchProblem_BuildResult = result
End Function

' ДЛЯ ТЕСТОВ: конфигурационная, входная и внутренняя ошибки запрещают повтор
' попытки; реальная численная неудача допускает предусмотренный fallback.
Private Property Get ILimitSearchProblem_StoppedTerminally() As Boolean
    ILimitSearchProblem_StoppedTerminally = (mFailureCode = sfcInvalidConfiguration Or _
        mFailureCode = sfcInvalidInput Or mFailureCode = sfcInternalError)
End Property

Private Property Get ILimitSearchProblem_StopReason() As String
    ILimitSearchProblem_StopReason = mLimitState
End Property

' ДЛЯ ТЕСТОВ: сбрасывает рабочую причину между попытками той же fixture.
Private Sub ILimitSearchProblem_ClearStopReason()
    mLimitState = vbNullString
End Sub

' ДЛЯ ТЕСТОВ: общий текст заполняет только отсутствующую конкретную причину.
Private Sub ILimitSearchProblem_SetStopReasonIfEmpty(ByVal reason As String)
    If Len(mLimitState) = 0 Then mLimitState = reason
End Sub

' ДЛЯ ТЕСТОВ: сохраняет сообщения перехода fallback в наблюдаемом порядке.
Private Sub ILimitSearchProblem_MergeFallback(ByVal firstLimitState As String, _
        ByVal firstStopReason As String, ByVal firstDiagnostics As String, _
        ByVal firstAttemptMessage As String, ByVal transitionMessage As String, _
        ByVal successMessage As String)
    AppendDiagnosticLine firstAttemptMessage
    AppendDiagnosticLine transitionMessage
    AppendDiagnosticLine successMessage
End Sub

' Добавляет строку в диагностический журнал fake-задачи.
Private Sub AppendDiagnosticLine(ByVal text As String)
    If Len(mDiagnosticLog) > 0 Then mDiagnosticLog = mDiagnosticLog & vbCrLf
    mDiagnosticLog = mDiagnosticLog & text
End Sub

' Форматирует числа одинаково для локальных test diagnostics.
Private Function FormatNumberInvariant(ByVal value As Double) As String
    FormatNumberInvariant = Replace$(Format$(value, "0.############"), ",", ".")
End Function

' ============================== ДЛЯ ТЕСТОВ ==============================
' Сохраняет typed failure общего Newton независимо от display-строк и текста.
Private Sub ILimitSearchProblem_SetFailure(ByVal failureCode As ESolverFailureCode, ByVal reason As String)
    mFailureCode = failureCode
    mConverged = False
    mLimitState = PROBE_NUMFAIL
    If failureCode = sfcInvalidConfiguration Or failureCode = sfcInvalidInput Then mLimitState = PROBE_INPUTERR
    AppendDiagnosticLine reason
End Sub

' ДЛЯ ТЕСТОВ: меняет ровно один guard/численный сценарий линейной системы.
' Первый сценарий воспроизводит бесконечное деление alpha при min alpha=0.
Public Sub ConfigureUltimateGuardCase(ByVal scenario As Long, Optional ByVal kind As EResultKind = rkCapacity)
    Configure "Bisection", 1.25, 0.000000001, 20, resultKind:=kind
    mUltimateScenario = scenario
    Select Case scenario
        Case 1
            mUltimateMinAlpha = 0#: mUltimateConstantNorm = True
        Case 2
            mUltimateDamping = 0#
        Case 3
            mUltimateDamping = 0.01: mUltimateMinAlpha = 0.1
        Case 4
            mMaxIterations = 0
        Case 5
            mUltimateMaxDeltaEps = -1#
        Case 6
            mUltimateInitialEps = 1#: mRootLambda = 1001#: mUltimateMaxDeltaEps = 1E-30
    End Select
End Sub

' ДЛЯ ТЕСТОВ: счетчик и typed-причина для независимых assertions.
Public Property Get UltimateCalls() As Long
    UltimateCalls = mUltimateCalls
End Property

Public Property Get FailureCode() As ESolverFailureCode
    FailureCode = mFailureCode
End Property
