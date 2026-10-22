VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CStateProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Получение и повторное использование расчётного состояния
' ==========================================================================
' По запросу ищет подходящее состояние в CStateRepository. Если его нет,
' запускает CStateSolutionRunner и сохраняет полученный результат.
' Хранилище содержит конечные состояния, имеющие самостоятельный смысл
' в Results; промежуточные пробы поиска предела остаются внутри поиска.

Private mSection As CSectionModel ' Модель сечения для решения равновесия.
Private mMaterialProvider As CMaterialModelProvider ' Подготовка физических диаграмм и разрешённых вспомогательных продолжений.
Private mRepository As CStateRepository ' Хранилище расчётных состояний этого сочетания нагрузок.

Private mSolverMethod As String ' Newton или Secant.
Private mLoadSteps As Long ' Число ступеней приложения нагрузки.
Private mMaxIterations As Long ' Максимальное число итераций решения равновесия.
Private mToleranceN As Double ' Абсолютный допуск равновесия по N, Н.
Private mToleranceMx As Double ' Абсолютный допуск равновесия по Mx, Н*мм.
Private mToleranceMy As Double ' Абсолютный допуск равновесия по My, Н*мм.
Private mSecantMaxRestarts As Long ' Допустимое число обновлений матрицы секущих модулей.
Private mDiagnosticsEnabled As Boolean ' Нужно ли копить диагностический журнал.

Private mLastRunner As CStateSolutionRunner ' Объект последнего выполненного решения; Nothing, если использовано ранее сохранённое НДС.
Private mLastStateWasReused As Boolean ' Показывает, взято ли НДС из ранее сохранённых результатов.

' Задаёт начальные настройки решения равновесия для самостоятельного использования класса.
' При запуске из Excel они заменяются прочитанными настройками Config.
Private Sub Class_Initialize()
    mSolverMethod = "Newton"
    mLoadSteps = 8
    mMaxIterations = 80
    mToleranceN = 5#
    mToleranceMx = 5000#
    mToleranceMy = 5000#
    mSecantMaxRestarts = 2
    mDiagnosticsEnabled = True
End Sub

' Подключает модель, материалы и хранилище состояний. Если хранилище не передано,
' каждый запрос решается заново без использования прежних результатов.
Public Sub Initialize(ByVal section As CSectionModel, ByVal materialProvider As CMaterialModelProvider, _
        Optional ByVal repository As CStateRepository = Nothing)
    If section Is Nothing Then Err.Raise vbObjectError + 3990, "CStateProvider", "Модель сечения не передана."
    If materialProvider Is Nothing Then Err.Raise vbObjectError + 3991, "CStateProvider", "Поставщик материалов не передан."
    Set mSection = section
    Set mMaterialProvider = materialProvider
    Set mRepository = repository
End Sub

' Возвращает ранее найденное НДС для точно такого же запроса или решает равновесие заново.
' initialSolver задаёт только начальное приближение и не определяет совпадение запросов.
Public Function GetOrSolve(ByVal request As CStateRequest, _
        Optional ByVal initialSolver As CSectionSolver = Nothing) As CSectionStateResult
    If request Is Nothing Then Err.Raise vbObjectError + 3992, "CStateProvider", "Не передан запрос на расчёт НДС выбранного назначения."
    Set mLastRunner = Nothing
    mLastStateWasReused = False

    If mSection Is Nothing Or mMaterialProvider Is Nothing Then
        Dim missingMeta As CResultMeta
        Set missingMeta = New CResultMeta
        missingMeta.SetResult rsInternalError, rcInternalError, ResultKindForRequest(request), _
            "Расчетное состояние не получено: не подключена модель сечения или поставщик материалов.", _
            vbNullString, True, False
        Dim missingState As CSectionStateResult
        Set missingState = New CSectionStateResult
        missingState.InitializeFailure request, missingMeta
        If Not mRepository Is Nothing Then mRepository.StoreState missingState
        Set GetOrSolve = missingState
        Exit Function
    End If

    If Not mRepository Is Nothing Then
        BindRepositoryContext
        Dim cached As CSectionStateResult
        Set cached = mRepository.FindEquivalent(request)
        If Not cached Is Nothing Then
            mLastStateWasReused = True
            mRepository.StoreState cached
            Set GetOrSolve = cached
            Exit Function
        End If
    End If

    Dim runner As CStateSolutionRunner
    Set runner = New CStateSolutionRunner
    ConfigureRunner runner

    Dim concreteMaterial As CMaterialDiagram
    Dim steelMaterial As CMaterialDiagram
    Set concreteMaterial = ConcreteMaterialForRequest(request)
    Set steelMaterial = SteelMaterialForRequest(request)

    If initialSolver Is Nothing Then
        runner.Solve mSection, concreteMaterial, steelMaterial, _
            request.TargetN, request.TargetMx, request.TargetMy, request.AllowRetryStarts
    Else
        runner.SolveWithInitialSolver mSection, concreteMaterial, steelMaterial, _
            request.TargetN, request.TargetMx, request.TargetMy, initialSolver, request.AllowRetryStarts
    End If
    Set mLastRunner = runner

    Dim service As CSectionSolver
    Set service = runner.ResultSolver
    If service Is Nothing Then
        Set missingMeta = New CResultMeta
        missingMeta.SetResult rsInternalError, rcSolverDidNotReturnState, ResultKindForRequest(request), _
            "Расчетное НДС не получено: решатель не вернул результат для запрошенного состояния.", _
            runner.DiagnosticLog, True, False
        Set missingState = New CSectionStateResult
        missingState.InitializeFailure request, missingMeta
        If Not mRepository Is Nothing Then mRepository.StoreForRequest request, missingState
        Set GetOrSolve = missingState
        Exit Function
    End If

    Dim stateResult As CSectionStateResult
    Set stateResult = New CSectionStateResult
    stateResult.InitializeFromSolver request.StateType, request.MaterialRole, request.MaterialSpec, _
        service, runner.ExtensionUsed, runner.WithinPhysicalRange, _
        runner.SolverCallCount, runner.DiagnosticLog

    If Not mRepository Is Nothing Then mRepository.StoreForRequest request, stateResult
    Set GetOrSolve = stateResult
End Function

' Определяет, к какой проверке относится статус состояния, сохраняя причину решателя.
' Отсутствие обязательного программного объекта не выдаётся за ошибку пользовательских нагрузок.
Private Function ResultKindForRequest(ByVal request As CStateRequest) As EResultKind
    Select Case request.StateType
        Case sstCapacityState: ResultKindForRequest = rkCapacity
        Case sstPreCrackState, sstPostCrackState: ResultKindForRequest = rkCrackFormation
        Case Else: ResultKindForRequest = rkDirectState
    End Select
End Function

' Разрешает повторное использование НДС только для той же геометрии, материалов,
' допусков и режима продолжения диаграмм. Начальные приближения, число ступеней,
' повторные попытки и диагностика не меняют физическую постановку запроса.
Private Sub BindRepositoryContext()
    Dim acceptanceKey As String
    acceptanceKey = CStr(mMaterialProvider.DiagramExtensionEnabled) & "|" & _
        CStr(mToleranceN) & "|" & CStr(mToleranceMx) & "|" & CStr(mToleranceMy)
    mRepository.BindContext mSection, mMaterialProvider, acceptanceKey
End Sub

' Восстанавливает данные расчётного состояния для потребителя по сохранённому НДС.
' Повторное решение равновесия не выполняется, новая запись в хранилище не создаётся;
' объект последнего запуска LastRunner для этого не требуется.
Public Function SolverSnapshot(ByVal stateResult As CSectionStateResult) As CSectionSolver
    If stateResult Is Nothing Then Exit Function
    If Not stateResult.Converged Then Exit Function

    Dim request As CStateRequest
    Set request = New CStateRequest
    request.Initialize stateResult.StateType, stateResult.MaterialModelRole, stateResult.MaterialSpec, _
        stateResult.TargetN, stateResult.TargetMx, stateResult.TargetMy, stateResult.ExtensionUsed, False
    Dim runner As CStateSolutionRunner
    Set runner = New CStateSolutionRunner
    ConfigureRunner runner
    Set SolverSnapshot = runner.RestoreSnapshot(mSection, ConcreteMaterialForRequest(request), _
        SteelMaterialForRequest(request), stateResult)
End Function

' Передаёт сохранённые численные настройки новому объекту решения НДС.
Private Sub ConfigureRunner(ByVal runner As CStateSolutionRunner)
    runner.LoadSteps = mLoadSteps
    runner.MaxIterations = mMaxIterations
    runner.ToleranceN = mToleranceN
    runner.ToleranceMx = mToleranceMx
    runner.ToleranceMy = mToleranceMy
    runner.SolverMethod = mSolverMethod
    runner.SecantMaxRestarts = mSecantMaxRestarts
    runner.DiagnosticsEnabled = mDiagnosticsEnabled
End Sub

' Выбирает диаграмму бетона по назначению и режиму материалов в запросе.
Private Function ConcreteMaterialForRequest(ByVal request As CStateRequest) As CMaterialDiagram
    If request.UseDiagramExtension Then
        Set ConcreteMaterialForRequest = mMaterialProvider.ConcreteMaterialForEquilibriumFromSpec(request.MaterialSpec)
    Else
        Set ConcreteMaterialForRequest = mMaterialProvider.ConcreteMaterialFromSpec(request.MaterialSpec)
    End If
End Function

' Выбирает диаграмму арматуры по назначению и режиму материалов в запросе.
Private Function SteelMaterialForRequest(ByVal request As CStateRequest) As CMaterialDiagram
    If request.UseDiagramExtension Then
        Set SteelMaterialForRequest = mMaterialProvider.SteelMaterialForEquilibriumFromSpec(request.MaterialSpec)
    Else
        Set SteelMaterialForRequest = mMaterialProvider.SteelMaterialFromSpec(request.MaterialSpec)
    End If
End Function

' Возвращает объект последнего выполненного решения; Nothing, если взято ранее найденное НДС.
Public Property Get LastRunner() As CStateSolutionRunner
    Set LastRunner = mLastRunner
End Property

' Показывает, получено ли НДС из ранее сохранённых результатов.
Public Property Get LastStateWasReused() As Boolean
    LastStateWasReused = mLastStateWasReused
End Property

' Newton или Secant.
Public Property Let SolverMethod(ByVal value As String)
    mSolverMethod = value
End Property

' Число ступеней приложения нагрузки.
Public Property Let LoadSteps(ByVal value As Long)
    mLoadSteps = value
End Property

' Задаёт максимум итераций решения равновесия на одной ступени приложения нагрузки.
Public Property Let MaxIterations(ByVal value As Long)
    mMaxIterations = value
End Property

' Абсолютный допуск равновесия по N, Н.
Public Property Let ToleranceN(ByVal value As Double)
    mToleranceN = value
End Property

' Абсолютный допуск равновесия по Mx, Н*мм.
Public Property Let ToleranceMx(ByVal value As Double)
    mToleranceMx = value
End Property

' Абсолютный допуск равновесия по My, Н*мм.
Public Property Let ToleranceMy(ByVal value As Double)
    mToleranceMy = value
End Property

' Задаёт допустимое число перезапусков приближённой матрицы чувствительности в методе Secant.
Public Property Let SecantMaxRestarts(ByVal value As Long)
    mSecantMaxRestarts = value
End Property

' Нужно ли копить диагностический журнал.
Public Property Let DiagnosticsEnabled(ByVal value As Boolean)
    mDiagnosticsEnabled = value
End Property
