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
' ДЛЯ ТЕСТОВ: управляемая задача общего поиска предела
' ==========================================================================
' Задаёт функцию f(lambda) = lambda - lambdaRoot и дополнительные
' аналитические невязки для проверки численных алгоритмов без материалов.
' Проверяет, что общий поиск работает через ILimitSearchProblem независимо
' от расчётов несущей способности и образования трещины.

Implements ILimitSearchProblem

Private Const PROBE_ACCEPTABLE As String = "Acceptable" ' Точка находится до предела.
Private Const PROBE_LIMIT As String = "Limit" ' Точка достигла или превысила предел.
Private Const PROBE_NUMFAIL As String = "NumericalFailure" ' Искусственная численная ошибка.
Private Const PROBE_INPUTERR As String = "InvalidInput" ' Искусственная ошибка входных данных.

Private mSearchMethod As String ' Численный метод одномерного поиска, который проверяется данным тестом.
Private mRootLambda As Double ' Точный корень линейной тестовой функции.
Private mTolerance As Double ' Допустимая ширина интервала lambda, передаваемая общему поиску.
Private mMaxIterations As Long ' Лимит итераций для проверки T02.
Private mInitialStep As Double ' Начальный шаг расширения скобки.
Private mMaxLambda As Double ' Техническая верхняя граница lambda.
Private mKind As EResultKind ' Вид инженерной задачи, с которым запускается общий поиск.
Private mConverged As Boolean ' Признак того, что общий поиск принял конечную точку.
Private mFinalLambda As Double ' Множитель lambda, фактически принятый общим поиском.
Private mFinalMethod As String ' Метод, которым была принята итоговая точка.
Private mFinalIterations As Long ' Число итераций, переданное при финализации.
Private mLimitState As String ' Последний сохранённый статус тестовой задачи.
Private mDiagnosticLog As String ' Накопленный диагностический текст.
Private mIterationLimitHit As Boolean ' Признак остановки общего поиска по максимальному числу итераций.
Private mProbeMaxLambda As Double ' Наибольшая реально проверенная точка.
Private mProbeCalls As Long ' Счетчик проб общего алгоритма.
Private mSearchExecuted As Boolean ' ДЛЯ ТЕСТОВ: факт начала пробы или подготовки старта, независимо от ее успеха.
Private mFailureAbove As Double ' Искусственная численная неудача выше этой lambda; 0 отключает ее.
Private mFailureAt As Double ' ДЛЯ ТЕСТОВ: неудачная проба возвращает True и нулевую невязку-заглушку.
Private mStepCriterion As Boolean ' Разрывный критерий не имеет нулевой невязки на соседних Double.
Private mRecoveryWithoutProgress As Boolean ' Намеренно неверная операция уменьшения lambda возвращает прежнюю границу для проверки защиты от застоя.
Private mUltimateInitialEps As Double ' ДЛЯ ТЕСТОВ: ненулевой старт для машинной стагнации.
Private mUltimateConstantNorm As Boolean ' ДЛЯ ТЕСТОВ: запрещает уменьшение общей меры невязок при подборе шага.
Private mUltimateScenario As Long ' ДЛЯ ТЕСТОВ: управляемая неудача оценки пробной точки.
Private mUltimateCalls As Long ' ДЛЯ ТЕСТОВ: число оценок невязки Newton.
Private mUltimateGuessCalls As Long ' ДЛЯ ТЕСТОВ: ограниченный возврат к общему старту.
Private mUltimateUsesWarmGuess As Boolean ' Использует ли тестовая задача предыдущее решение при поиске по деформациям.
Private mFailureCode As ESolverFailureCode ' ДЛЯ ТЕСТОВ: точная причина остановки.
Private mFailureReason As String ' ДЛЯ ТЕСТОВ: последняя причина операции задачи, сохранённая отдельно от общего журнала.
Private mMultiplierScenario As Long ' ДЛЯ ТЕСТОВ: управляемая неудача выбранной операции задачи при LoadMultiplier.
Private mFallbackCalls As Long ' ДЛЯ ТЕСТОВ: терминальная ошибка запрещает резервную финализацию.
Private mAcceptableProbeCalls As Long ' ДЛЯ ТЕСТОВ: успешные пробы между численными отказами.
Private mFinalizationMaxWidth As Double ' ДЛЯ ТЕСТОВ: приемке нужна более узкая скобка, чем пользовательский допуск.
Private mBracketFinalizationCalls As Long ' Число завершений поиска внутри найденного интервала.
Private mDeferOverallCompletion As Boolean ' ДЛЯ ТЕСТОВ: последующий этап еще не выполнен, но точка поиска принята.
Private mPointAccepted As Boolean ' Признак принятой тестовой точки поиска.

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Задаёт линейную контрольную задачу перед запуском общего поиска.
Public Sub TestConfigure(ByVal searchMethod As String, ByVal rootLambda As Double, _
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
    mSearchExecuted = False
    mFailureAbove = 0#
    mFailureAt = -1#
    mStepCriterion = False
    mRecoveryWithoutProgress = False
    mUltimateInitialEps = 0#
    mUltimateConstantNorm = False
    mUltimateScenario = 0
    mUltimateCalls = 0
    mUltimateGuessCalls = 0
    mUltimateUsesWarmGuess = False
    mFailureCode = sfcNone
    mFailureReason = vbNullString
    mMultiplierScenario = 0
    mAcceptableProbeCalls = 0
    mFallbackCalls = 0
    mFinalizationMaxWidth = 0#
    mBracketFinalizationCalls = 0
    mDeferOverallCompletion = False: mPointAccepted = False
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' отрицательная ширина моделирует постоянно неудачное уточнение.
Public Property Let TestFinalizationMaxWidth(ByVal value As Double)
    mFinalizationMaxWidth = value
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' число завершений поиска внутри найденного интервала lambda.
Public Property Get TestBracketFinalizationCalls() As Long
    TestBracketFinalizationCalls = mBracketFinalizationCalls
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' последующий этап еще не выполнен, но точка поиска принята.
Public Property Let TestDeferOverallCompletion(ByVal value As Boolean)
    mDeferOverallCompletion = value
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' задает численную неудачу верхних проб, не физический предел.
Public Property Let TestFailureAbove(ByVal value As Double)
    mFailureAbove = value
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' неудачная проба возвращает True и нулевую невязку-заглушку.
Public Property Let TestFailureAt(ByVal value As Double)
    mFailureAt = value
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' успешные пробы между численными отказами.
Public Property Get TestAcceptableProbeCalls() As Long
    TestAcceptableProbeCalls = mAcceptableProbeCalls
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' не позволяет признать соседнюю граничную lambda точным корнем.
Public Property Let TestStepCriterion(ByVal value As Boolean)
    mStepCriterion = value
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' общий поиск должен остановить уточнение,
' если предложенное lambda не уменьшает рабочий интервал.
Public Property Let TestRecoveryWithoutProgress(ByVal value As Boolean)
    mRecoveryWithoutProgress = value
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' диагностические счетчики расширения и восстановления скобки.
Public Property Get TestProbeMaxLambda() As Double
    TestProbeMaxLambda = mProbeMaxLambda
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Счетчик проб общего алгоритма.
Public Property Get TestProbeCalls() As Long
    TestProbeCalls = mProbeCalls
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Множитель lambda, фактически принятый общим поиском.
Public Property Get TestFinalLambda() As Double
    TestFinalLambda = mFinalLambda
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Метод, которым была принята итоговая точка.
Public Property Get TestFinalMethod() As String
    TestFinalMethod = mFinalMethod
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Число итераций, переданное при финализации.
Public Property Get TestFinalIterations() As Long
    TestFinalIterations = mFinalIterations
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Признак остановки общего поиска по максимальному числу итераций.
Public Property Get TestIterationLimitHit() As Boolean
    TestIterationLimitHit = mIterationLimitHit
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Признак того, что общий поиск принял конечную точку.
Public Property Get TestConverged() As Boolean
    TestConverged = mConverged
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Последний сохранённый статус тестовой задачи.
Public Property Get TestLimitState() As String
    TestLimitState = mLimitState
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Накопленный диагностический текст.
Public Property Get TestDiagnosticLog() As String
    TestDiagnosticLog = mDiagnosticLog
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Вид инженерной задачи, с которым запускается общий поиск.
Private Property Get ILimitSearchProblem_ResultKind() As EResultKind
    ILimitSearchProblem_ResultKind = mKind
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Начальный шаг расширения скобки.
Private Property Get ILimitSearchProblem_InitialLambdaStep() As Double
    ILimitSearchProblem_InitialLambdaStep = mInitialStep
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Техническая верхняя граница lambda.
Private Property Get ILimitSearchProblem_MaxLambda() As Double
    ILimitSearchProblem_MaxLambda = mMaxLambda
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Допустимая ширина интервала lambda, передаваемая общему поиску.
Private Property Get ILimitSearchProblem_LambdaTolerance() As Double
    ILimitSearchProblem_LambdaTolerance = mTolerance
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Лимит итераций для проверки T02.
Private Property Get ILimitSearchProblem_MaxIterations() As Long
    ILimitSearchProblem_MaxIterations = mMaxIterations
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' позволяет явно проверить исчерпание бюджета Bisection.
Private Property Get ILimitSearchProblem_BisectionMaxIterations() As Long
    ILimitSearchProblem_BisectionMaxIterations = mMaxIterations
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Численный метод одномерного поиска, который проверяется данным тестом.
Private Property Get ILimitSearchProblem_SearchMethod() As String
    ILimitSearchProblem_SearchMethod = mSearchMethod
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Признак того, что общий поиск принял конечную точку.
Private Property Get ILimitSearchProblem_Converged() As Boolean
    ILimitSearchProblem_Converged = mConverged
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Последний сохранённый статус тестовой задачи.
Private Property Get ILimitSearchProblem_LimitState() As String
    ILimitSearchProblem_LimitState = mLimitState
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Накопленный диагностический текст.
Private Property Get ILimitSearchProblem_DiagnosticLog() As String
    ILimitSearchProblem_DiagnosticLog = mDiagnosticLog
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' принимает выбранный путь общего сценария,
' сохраняя управляемую аналитическую оценку нагрузки без реальных материалов.
Private Sub ILimitSearchProblem_BeginLoadPath(ByVal request As CLimitSearchRequest)
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' история попыток хранится в настоящем CLimitSearchResult.
Private Sub ILimitSearchProblem_FinishLoadPath(ByVal request As CLimitSearchRequest, _
        ByVal previousPathComments As String)
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' начало линейной функции должно лежать до заданного корня.
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

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Контрольная задача не требует отдельной подготовки материалов или геометрии перед поиском.
Private Function ILimitSearchProblem_Prepare(ByVal request As CLimitSearchRequest) As Boolean
    If mMultiplierScenario = 13 Then Err.Raise 9, "CTestLimitSearchProblem", TestMultiplierFaultReason()
    ILimitSearchProblem_Prepare = True
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' задаёт три независимых аналитических невязки
' для проверки общего метода Ньютона без расчётов железобетона.
Private Function ILimitSearchProblem_PrepareUltimate(ByVal request As CLimitSearchRequest, _
        ByVal loadPath As CLoadPathVector) As Boolean
    ILimitSearchProblem_PrepareUltimate = Not loadPath Is Nothing
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' запускает настоящий общий Newton на аналитических callbacks.
' Section и материалы не нужны простой тестовой системе трех невязок.
Private Function ILimitSearchProblem_RunUltimateAttempts(ByVal request As CLimitSearchRequest) As Boolean
    Dim search As CUltimateStrainSearch
    Set search = New CUltimateStrainSearch
    ILimitSearchProblem_RunUltimateAttempts = search.RunNewton(Me, Nothing, Nothing, Nothing, _
        ILimitSearchProblem_BuildLoadPath(request), Nothing, "Тестовый поиск не сошелся.")
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' найденное lambda уже сохранено при приёме точки.
' Дополнительная инженерная обработка аналитическому примеру не требуется.
Private Sub ILimitSearchProblem_CompleteUltimate(ByVal loadPath As CLoadPathVector)
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' задает управляемый старт по epsilon0 и нулевые кривизны,
' чтобы воспроизводимо проверять нормальную сходимость и стагнацию общего Newton.
Private Sub ILimitSearchProblem_LimitSearchInitializeUltimateGuess(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal loadPath As CLoadPathVector, ByVal startSolver As CSectionSolver, _
        ByRef eps0 As Double, ByRef kx As Double, ByRef ky As Double)
    mSearchExecuted = True
    mUltimateGuessCalls = mUltimateGuessCalls + 1
    Select Case mUltimateScenario
        Case 18: Err.Raise 6
        Case 19: Err.Raise 9
        Case 20: Err.Raise 6, "CTestLimitSearchProblem", "Тестовое переполнение стартовой плоскости."
        Case 21: Err.Raise 9, "CTestLimitSearchProblem", "Тестовое нарушение подготовки стартовой плоскости."
    End Select
    eps0 = mUltimateInitialEps: kx = 0#: ky = 0#
    If mUltimateUsesWarmGuess And Not startSolver Is Nothing Then
        eps0 = startSolver.Epsilon0: kx = startSolver.KappaX: ky = startSolver.KappaY
    End If
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' вычисляет линейные невязки и считает обращения к задаче.
' Отдельные случаи намеренно возвращают код ошибки или исключение VBA
' на второй пробе для проверки остановки при вычислении Якобиана.
Private Function ILimitSearchProblem_LimitSearchEvaluateUltimateResidual(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal loadPath As CLoadPathVector, ByVal eps0 As Double, _
        ByVal kx As Double, ByVal ky As Double, ByRef solver As CSectionSolver, _
        ByRef rPath1 As Double, ByRef rPath2 As Double, ByRef rLimit As Double, _
        ByRef lambdaValue As Double, ByRef state As String, Optional ByVal centeredPlane As Boolean = False) As Boolean
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
            Case 14
                Err.Raise 6
            Case 15
                Err.Raise 9
            Case 16
                Err.Raise 11
            Case 17
                Err.Raise 11, "CTestLimitSearchProblem", "Тестовое деление на ноль."
        End Select
    End If
    rPath1 = kx
    rPath2 = ky
    lambdaValue = eps0 / 0.001
    rLimit = lambdaValue - mRootLambda
    If mUltimateScenario = 27 Then rLimit = Abs(lambdaValue) - mRootLambda
    state = PROBE_LIMIT
    ILimitSearchProblem_LimitSearchEvaluateUltimateResidual = True
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' проверяет три допуска аналитической системы.
' Режим постоянной меры невязок намеренно запрещает успех
' для воспроизводимой проверки застоя подбора шага.
Private Function ILimitSearchProblem_LimitSearchUltimateResidualReached(ByVal rPath1 As Double, _
        ByVal rPath2 As Double, ByVal rLimit As Double) As Boolean
    ILimitSearchProblem_LimitSearchUltimateResidualReached = _
        Abs(rPath1) <= mTolerance And Abs(rPath2) <= mTolerance And Abs(rLimit) <= mTolerance
    If mUltimateConstantNorm Then ILimitSearchProblem_LimitSearchUltimateResidualReached = False
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' возвращает сумму модулей невязок либо постоянную норму,
' когда нужно проверить отсутствие улучшающего представимого шага.
Private Function ILimitSearchProblem_LimitSearchUltimateResidualNorm(ByVal rPath1 As Double, _
        ByVal rPath2 As Double, ByVal rLimit As Double) As Double
    ILimitSearchProblem_LimitSearchUltimateResidualNorm = Abs(rPath1) + Abs(rPath2) + Abs(rLimit)
    If mUltimateConstantNorm Then ILimitSearchProblem_LimitSearchUltimateResidualNorm = 1#
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' признак масштабирования функции качества по матрице Якоби.
Private Property Get ILimitSearchProblem_LimitSearchUseJacobianMeritScaling() As Boolean
    ILimitSearchProblem_LimitSearchUseJacobianMeritScaling = False
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' сохраняет найденные lambda, метод и число итераций.
' Физическое НДС не создаётся; предел настоящего сечения проверяется отдельно.
Private Function ILimitSearchProblem_LimitSearchFinalizeUltimateResult(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal loadPath As CLoadPathVector, ByVal lambdaValue As Double, _
        ByVal solver As CSectionSolver, ByVal state As String, ByVal iterations As Long) As EUltimateFinalization
    If mUltimateScenario = 23 Then Exit Function
    If mUltimateScenario = 27 And lambdaValue < 0# Then
        ILimitSearchProblem_SetFailure sfcNumericalFailure, "Отрицательная ветвь тестового пути."
        ILimitSearchProblem_LimitSearchFinalizeUltimateResult = ufRejectedPoint
        Exit Function
    End If
    If mUltimateScenario = 24 Or mUltimateScenario = 25 Then
        ILimitSearchProblem_SetFailure sfcNumericalFailure, "Тестовая предельная точка не принадлежит допустимому направлению пути."
        ILimitSearchProblem_LimitSearchFinalizeUltimateResult = ufRejectedPoint
        Exit Function
    End If
    If mUltimateScenario = 26 Then
        ILimitSearchProblem_LimitSearchFinalizeUltimateResult = 99
        Exit Function
    End If
    mConverged = True
    mFinalLambda = lambdaValue
    mFinalMethod = "UltimateStrain"
    mFinalIterations = iterations
    mLimitState = state
    ILimitSearchProblem_LimitSearchFinalizeUltimateResult = ufAccepted
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' фиксирует численную остановку и ее причину без найденной точки.
Private Sub ILimitSearchProblem_LimitSearchSetUltimateNumericalFailure(ByVal reason As String)
    mConverged = False
    mLimitState = PROBE_NUMFAIL
    mFailureCode = sfcNumericalFailure
    TestAppendDiagnosticLine reason
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' одинаковые численные ограничения простой линейной системы.
Private Property Get ILimitSearchProblem_LimitSearchUltimateMaxIterations() As Long
    ILimitSearchProblem_LimitSearchUltimateMaxIterations = mMaxIterations
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Создает универсальную траекторию Offset + lambda*Base из request.
Private Function ILimitSearchProblem_BuildLoadPath(ByVal request As CLimitSearchRequest) As CLoadPathVector
    If mMultiplierScenario = 14 Then Err.Raise 9, "CTestLimitSearchProblem", TestMultiplierFaultReason()
    Dim loadPath As CLoadPathVector
    Set loadPath = New CLoadPathVector
    loadPath.Initialize request.NOffset, request.NBase, _
        request.MxOffset, request.MxBase, request.MyOffset, request.MyBase
    Set ILimitSearchProblem_BuildLoadPath = loadPath
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Оценивает линейную функцию при заданном lambda. Нулевая невязка
' означает достижение аналитического критерия; материальное НДС
' для подтверждения корня этой тестовой функции не требуется.
Private Function ILimitSearchProblem_EvaluatePoint(ByVal request As CLimitSearchRequest, _
        ByVal loadPath As CLoadPathVector, ByVal startSolver As CSectionSolver, _
        ByVal lambdaValue As Double, ByRef functionValue As Double, _
        ByRef probeState As String, ByRef solver As CSectionSolver) As Boolean
    mSearchExecuted = True
    functionValue = lambdaValue - mRootLambda
    If mStepCriterion Then
        If lambdaValue < mRootLambda Then functionValue = -1# Else functionValue = 1#
    End If
    mProbeCalls = mProbeCalls + 1
    If lambdaValue > mProbeMaxLambda Then mProbeMaxLambda = lambdaValue
    Set solver = Nothing
    If lambdaValue = mFailureAt Then
        functionValue = 0#
        probeState = PROBE_NUMFAIL
        ILimitSearchProblem_EvaluatePoint = True
        Exit Function
    End If
    Select Case mMultiplierScenario
        Case 1
            If lambdaValue = 0# Then Err.Raise 9, "CTestLimitSearchProblem", TestMultiplierFaultReason()
        Case 2
            If lambdaValue = 1# Then Err.Raise 9, "CTestLimitSearchProblem", TestMultiplierFaultReason()
        Case 3
            If lambdaValue = 1# Then Err.Raise 6, "CTestLimitSearchProblem", TestMultiplierFaultReason()
        Case 4, 5, 6, 7
            If lambdaValue = 1# Then
                Dim failureCode As ESolverFailureCode
                Select Case mMultiplierScenario
                    Case 4, 7: failureCode = sfcInvalidConfiguration
                    Case 5: failureCode = sfcInvalidInput
                    Case 6: failureCode = sfcInternalError
                End Select
                ILimitSearchProblem_SetFailure failureCode, TestMultiplierFaultReason()
                probeState = mLimitState
                If mMultiplierScenario = 7 Then Err.Raise 6, "CTestLimitSearchProblem", TestMultiplierFaultReason()
                Exit Function
            End If
        Case 8, 9
            If lambdaValue > 1# And lambdaValue < 2# Then
                If mMultiplierScenario = 8 Then Err.Raise 9, "CTestLimitSearchProblem", TestMultiplierFaultReason()
                ILimitSearchProblem_SetFailure sfcInvalidConfiguration, TestMultiplierFaultReason()
                probeState = mLimitState
                Exit Function
            End If
        Case 16
            If lambdaValue = 0# Then Err.Raise 6, "CTestLimitSearchProblem", TestMultiplierFaultReason()
    End Select
    If mFailureAbove > 0# And lambdaValue > mFailureAbove Then
        probeState = PROBE_NUMFAIL
        Exit Function
    End If
    If functionValue < 0# Then
        probeState = PROBE_ACCEPTABLE
        mAcceptableProbeCalls = mAcceptableProbeCalls + 1
    Else
        probeState = PROBE_LIMIT
    End If
    Set solver = Nothing
    ILimitSearchProblem_EvaluatePoint = True
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' различает допустимую точку, предел,
' намеренную численную неудачу и ошибку входных данных.
Private Function ILimitSearchProblem_IsAcceptableState(ByVal probeState As String) As Boolean
    ILimitSearchProblem_IsAcceptableState = (probeState = PROBE_ACCEPTABLE)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет физический предельный статус задачи; численная несходимость сама по себе не означает
' разрушение.
Private Function ILimitSearchProblem_IsLimitState(ByVal probeState As String) As Boolean
    ILimitSearchProblem_IsLimitState = (probeState = PROBE_LIMIT)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Распознаёт численную неудачу отдельно от физического предела.
Private Function ILimitSearchProblem_IsNumericalFailureState(ByVal probeState As String) As Boolean
    ILimitSearchProblem_IsNumericalFailureState = (probeState = PROBE_NUMFAIL)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' штатные аналитические callbacks не исключаются силовой границей.
Private Function ILimitSearchProblem_IsInvalidInputState(ByVal probeState As String) As Boolean
    ILimitSearchProblem_IsInvalidInputState = (probeState = PROBE_INPUTERR)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' recovery использует середину оставшегося интервала.
Private Function ILimitSearchProblem_ReduceLambdaAfterNumericalFailure(ByVal lowLambda As Double, _
        ByVal highLambda As Double, ByVal probeState As String) As Double
    If mRecoveryWithoutProgress Then
        ILimitSearchProblem_ReduceLambdaAfterNumericalFailure = highLambda
    Else
        ILimitSearchProblem_ReduceLambdaAfterNumericalFailure = lowLambda + 0.5 * (highLambda - lowLambda)
    End If
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Запоминает последнюю acceptable-точку только в диагностике.
Private Sub ILimitSearchProblem_RecordAcceptable(ByVal loadPath As CLoadPathVector, _
        ByVal lambdaValue As Double)
    TestAppendDiagnosticLine "acceptable lambda=" & TestFormatNumberInvariant(lambdaValue)
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' фиксирует остановку начальной пробы, не создавая предельную точку.
Private Sub ILimitSearchProblem_HandleInitialFailure(ByVal probeState As String)
    mLimitState = PROBE_NUMFAIL
    TestAppendDiagnosticLine "initialFailure=" & probeState
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' различает достижение технической границы поиска
' и подтверждённый корень аналитической функции.
Private Sub ILimitSearchProblem_HandleBracketNotFound(ByVal lowLambda As Double)
    If mMultiplierScenario = 12 Then
        ILimitSearchProblem_SetFailure sfcInvalidConfiguration, TestMultiplierFaultReason()
        Exit Sub
    End If
    mLimitState = "SearchBoundReached"
    TestAppendDiagnosticLine "bracketNotFound lowLambda=" & TestFormatNumberInvariant(lowLambda)
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' сохраняет lambda неуспешной верхней пробы для проверки recovery.
Private Sub ILimitSearchProblem_HandleBracketNumericalFailure(ByVal lowLambda As Double, _
        ByVal failedLambda As Double, ByVal probeState As String)
    mLimitState = PROBE_NUMFAIL
    TestAppendDiagnosticLine "bracketNumericalFailure failedLambda=" & TestFormatNumberInvariant(failedLambda)
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' лимит итераций bisection не должен превращаться в найденный предел.
Private Sub ILimitSearchProblem_HandleIterationLimit(ByVal maxIterations As Long, _
        ByVal remainingWidth As Double, ByVal tolerance As Double)
    mConverged = False
    mIterationLimitHit = True
    mLimitState = PROBE_NUMFAIL
    TestAppendDiagnosticLine "iterationLimit maxIterations=" & CStr(maxIterations) & _
        "; remainingWidth=" & TestFormatNumberInvariant(remainingWidth) & _
        "; tolerance=" & TestFormatNumberInvariant(tolerance)
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' неизвестный исход пробы оставляет поиск неуспешным
' и сохраняется для диагностики.
Private Sub ILimitSearchProblem_HandleUnexpectedState(ByVal probeState As String)
    mLimitState = PROBE_NUMFAIL
    TestAppendDiagnosticLine "unexpectedState=" & probeState
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' принимает ошибку параметров общего поиска без фиктивной сходимости.
Private Sub ILimitSearchProblem_SetInvalidInput(ByVal message As String)
    mLimitState = PROBE_INPUTERR
    TestAppendDiagnosticLine message
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' принимает численную причину общего поиска и сохраняет ее в журнале.
Private Sub ILimitSearchProblem_SetNumericalFailure(ByVal message As String)
    mLimitState = PROBE_NUMFAIL
    TestAppendDiagnosticLine message
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' принимает сообщения, позволяющие проверить фактическую последовательность поиска.
Private Sub ILimitSearchProblem_AppendDiagnostic(ByVal message As String)
    TestAppendDiagnosticLine message
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сохраняет конечную нагрузку lambda, выбранную общим численным поиском.
Private Function ILimitSearchProblem_FinalizeAt(ByVal request As CLimitSearchRequest, _
        ByVal loadPath As CLoadPathVector, ByVal lambdaValue As Double, _
        ByVal probeState As String, ByVal methodName As String, _
        ByVal solver As CSectionSolver, ByVal iterations As Long) As Boolean
    If mMultiplierScenario = 10 Then Err.Raise 9, "CTestLimitSearchProblem", TestMultiplierFaultReason()
    If mMultiplierScenario = 11 Then
        ILimitSearchProblem_SetFailure sfcInvalidConfiguration, TestMultiplierFaultReason()
        Exit Function
    End If
    mConverged = True
    mFinalLambda = lambdaValue
    mFinalMethod = methodName
    mFinalIterations = iterations
    mLimitState = probeState
    TestAppendDiagnosticLine "final lambda=" & TestFormatNumberInvariant(lambdaValue) & _
        "; method=" & methodName
    ILimitSearchProblem_FinalizeAt = True
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' принимает сторону интервала по виду задачи, проверяя
' общий порядок завершения несущей способности и образования трещины
' без отдельных численных алгоритмов.
Private Function ILimitSearchProblem_FinalizeBracket(ByVal request As CLimitSearchRequest, _
        ByVal loadPath As CLoadPathVector, ByVal lowLambda As Double, ByVal lowSolver As CSectionSolver, _
        ByVal highLambda As Double, ByVal highSolver As CSectionSolver, _
        ByVal methodName As String, ByVal iterations As Long) As Boolean
    mBracketFinalizationCalls = mBracketFinalizationCalls + 1
    mFinalIterations = iterations
    If mMultiplierScenario = 10 Then Err.Raise 9, "CTestLimitSearchProblem", TestMultiplierFaultReason()
    If mMultiplierScenario = 11 Then
        ILimitSearchProblem_SetFailure sfcInvalidConfiguration, TestMultiplierFaultReason()
        Exit Function
    End If
    If mFinalizationMaxWidth < 0# Or (mFinalizationMaxWidth > 0# And _
            highLambda - lowLambda > mFinalizationMaxWidth) Then
        ILimitSearchProblem_SetNumericalFailure "Тестовая физическая приемка требует более узкой скобки."
        Exit Function
    End If
    Dim lambdaValue As Double
    If mKind = rkCrackFormation Then lambdaValue = highLambda Else lambdaValue = lowLambda
    mConverged = Not mDeferOverallCompletion
    mPointAccepted = True
    mFinalLambda = lambdaValue
    mFinalMethod = methodName
    mLimitState = PROBE_ACCEPTABLE
    TestAppendDiagnosticLine "retained lambda=" & TestFormatNumberInvariant(lambdaValue) & _
        "; method=" & methodName
    ILimitSearchProblem_FinalizeBracket = True
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' отмечает вызов последней попытки финализации, не объявляя успех.
Private Sub ILimitSearchProblem_TryFinalizeAfterSearchFailure(ByVal request As CLimitSearchRequest)
    mFallbackCalls = mFallbackCalls + 1
    If mMultiplierScenario = 15 Then Err.Raise 9, "CTestLimitSearchProblem", TestMultiplierFaultReason()
    TestAppendDiagnosticLine "retainedCapacity"
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' сохраняет полный численный итог без реального решателя НДС.
Private Function ILimitSearchProblem_BuildResult(ByVal request As CLimitSearchRequest, _
        ByVal searchSucceeded As Boolean) As CLimitSearchResult
    Dim meta As CResultMeta
    Set meta = New CResultMeta
    If mConverged Or mPointAccepted Then
        meta.SetResult rsSuccess, rcCheckPassed, mKind, "Тестовый критерий найден."
    ElseIf mFailureCode <> sfcNone Then
        meta.SetSolverFailure mFailureCode, mKind, mDiagnosticLog, mSearchExecuted
    ElseIf mLimitState = "SearchBoundReached" Then
        meta.SetResult rsNumericalFailure, rcSearchBoundReached, mKind, mLimitState
    ElseIf mLimitState = PROBE_INPUTERR Then
        meta.SetResult rsInvalidInput, rcInvalidInput, mKind, mLimitState
    Else
        meta.SetResult rsNumericalFailure, rcNumericalFailure, mKind, mLimitState
    End If
    Dim result As CLimitSearchResult
    Set result = New CLimitSearchResult
    result.Initialize meta, request.SolutionStrategy, mFinalMethod, mConverged Or mPointAccepted, mFinalLambda, _
        mFinalLambda, 0#, 0#, Nothing, mLimitState, mDiagnosticLog, vbNullString, mSearchExecuted
    Set ILimitSearchProblem_BuildResult = result
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' ошибки настройки, ввода и программы запрещают повторную попытку.
' Только численная неудача позволяет предусмотренную смену метода.
Private Property Get ILimitSearchProblem_StoppedTerminally() As Boolean
    ILimitSearchProblem_StoppedTerminally = (mFailureCode = sfcInvalidConfiguration Or _
        mFailureCode = sfcInvalidInput Or mFailureCode = sfcInternalError)
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Последний сохранённый статус тестовой задачи.
Private Property Get ILimitSearchProblem_StopReason() As String
    ILimitSearchProblem_StopReason = mLimitState
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' очищает рабочую причину между попытками одной контрольной задачи.
Private Sub ILimitSearchProblem_ClearStopReason()
    mLimitState = vbNullString
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' общий текст заполняет только отсутствующую конкретную причину.
Private Sub ILimitSearchProblem_SetStopReasonIfEmpty(ByVal reason As String)
    If Len(mLimitState) = 0 Then mLimitState = reason
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' сохраняет порядок сообщений перехода к дополнительному методу.
Private Sub ILimitSearchProblem_MergeFallback(ByVal firstLimitState As String, _
        ByVal firstStopReason As String, ByVal firstDiagnostics As String, _
        ByVal firstAttemptMessage As String, ByVal transitionMessage As String, _
        ByVal successMessage As String, Optional ByVal firstSearchExecuted As Boolean = False)
    mSearchExecuted = mSearchExecuted Or firstSearchExecuted
    TestAppendDiagnosticLine firstAttemptMessage
    TestAppendDiagnosticLine transitionMessage
    TestAppendDiagnosticLine successMessage
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Добавляет строку в диагностический журнал тестовой задачи.
Private Sub TestAppendDiagnosticLine(ByVal text As String)
    If Len(mDiagnosticLog) > 0 Then mDiagnosticLog = mDiagnosticLog & vbCrLf
    mDiagnosticLog = mDiagnosticLog & text
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Форматирует числа одинаково для локальных test diagnostics.
Private Function TestFormatNumberInvariant(ByVal value As Double) As String
    TestFormatNumberInvariant = Replace$(Format$(value, "0.############"), ",", ".")
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' ============================== ДЛЯ ТЕСТОВ ==============================
' Сохраняет код ошибки общего метода Ньютона независимо от текста и внешнего статуса.
Private Sub ILimitSearchProblem_SetFailure(ByVal failureCode As ESolverFailureCode, ByVal reason As String)
    mFailureCode = failureCode
    mFailureReason = reason
    mConverged = False
    mLimitState = PROBE_NUMFAIL
    If failureCode = sfcInvalidConfiguration Or failureCode = sfcInvalidInput Then mLimitState = PROBE_INPUTERR
    TestAppendDiagnosticLine reason
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' меняет ровно один guard/численный сценарий линейной системы.
' Первый сценарий проверяет конечный автоматический поиск при неизменной норме.
Public Sub TestConfigureUltimateGuardCase(ByVal scenario As Long, Optional ByVal kind As EResultKind = rkCapacity)
    TestConfigure "Bisection", 1.25, 0.000000001, 20, resultKind:=kind
    mUltimateScenario = scenario
    Select Case scenario
        Case 1
            mUltimateConstantNorm = True
        Case 4
            mMaxIterations = 0
        Case 22, 23
            mRootLambda = 1#: mUltimateInitialEps = 0.001 - 0.000000000000001
            mTolerance = 0.000000000000001
        Case 24, 26
            mRootLambda = 1#: mUltimateInitialEps = 0.001
        Case 25
            mRootLambda = 1#: mMaxIterations = 1
    End Select
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' возвращает счётчик и код причины для независимой проверки результата.
Public Property Get TestUltimateCalls() As Long
    TestUltimateCalls = mUltimateCalls
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' теплый старт имеет два симметричных корня, принимается положительный.
Public Sub TestConfigureUltimateWarmCase(ByVal maxIterations As Long, ByVal kind As EResultKind, _
        Optional ByVal scenario As Long = 27, Optional ByVal coldStrain As Double = 0.001)
    TestConfigureUltimateGuardCase scenario, kind
    mMaxIterations = maxIterations: mRootLambda = 1#
    mUltimateInitialEps = coldStrain: mUltimateUsesWarmGuess = True
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' ограниченный возврат к общему старту.
Public Property Get TestUltimateGuessCalls() As Long
    TestUltimateGuessCalls = mUltimateGuessCalls
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' точная причина остановки.
Public Property Get TestFailureCode() As ESolverFailureCode
    TestFailureCode = mFailureCode
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' возвращает именно аргумент последнего SetFailure. Это позволяет
' обнаружить потерю причины, даже если прежний текст остался в DiagnosticLog.
Public Property Get TestFailureReason() As String
    TestFailureReason = mFailureReason
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' отказы до пробы, при расширении, внутри скобки и в финализации
' проверяют один общий error-boundary без зависимости от реальных материалов.
' Настройка сбрасывается Configure, поэтому recovery не наследует старую причину.
Public Sub TestConfigureMultiplierFaultCase(ByVal scenario As Long, ByVal searchMethod As String, _
        ByVal kind As EResultKind)
    TestConfigure searchMethod, 1.25, 0.000001, 80, resultKind:=kind
    mMultiplierScenario = scenario
    If scenario = 12 Then mRootLambda = 100#: mMaxLambda = 4#
    ' Дополнительный способ завершения проверяем после положительной допустимой пробы.
    ' Точная двоичная граница позволяет исключить влияние ранней остановки по допуску lambda.
    If scenario = 15 Then mFailureAbove = 0.5
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' учитывает только вызов дополнительного способа решения.
' После ошибки настройки, ввода или программы основной поиск вызывать его не должен.
Public Property Get TestFallbackCalls() As Long
    TestFallbackCalls = mFallbackCalls
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' уникальная русская причина позволяет проверить сохранность
' первичного сообщения, не назначая машинный статус по тексту сообщения.
Private Function TestMultiplierFaultReason() As String
    TestMultiplierFaultReason = "Контрольный отказ LoadMultiplier, сценарий " & CStr(mMultiplierScenario) & "."
End Function
