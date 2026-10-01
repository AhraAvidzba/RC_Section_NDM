VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCapacityCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Инженерная обвязка расчета несущей способности
' ==========================================================================
' CCapacityCalculator собирает один capacity-запрос поверх общей
' CLimitSearch-архитектуры. Он не хранит batch-строку, не пишет Results и не
' выбирает худшее сочетание. Его задача - получить уже подготовленные сечение,
' материалы, load path и настройки solver-а, выполнить общий LimitSearch и
' вернуть CCapacityResult.
'
' Такой слой скрывает детали CCapacitySolver от batch-калькулятора: batch
' получает готовый CCapacityResult и не разбирает внутренние шаги search.

Private mInitialLambdaStep As Double ' Первый шаг расширения lambda для LoadMultiplier.
Private mSolutionStrategy As String ' Пользовательская стратегия Capacity.SolutionStrategy после чтения Config.
Private mSearchMethod As String ' Алгоритм одномерного LoadMultiplier-поиска.
Private mMaxLambda As Double ' Защитная верхняя граница lambda при расширении скобки.
Private mLambdaTolerance As Double ' Допуск по lambda для одномерного поиска.
Private mStrainTolerance As Double ' Допуск попадания в предельную деформацию UltimateStrain.
Private mMaxRetries As Long ' Число повторов lambda-точки при численной неудаче.
Private mBaseLoadSteps As Long ' Базовое число load steps внутреннего CSectionSolver.
Private mCapacitySolverMaxIterations As Long ' Лимит итераций search/внутренних solves для capacity.
Private mSolverToleranceN As Double ' Допуск равновесия внутреннего CSectionSolver по N.
Private mSolverToleranceMx As Double ' Допуск равновесия внутреннего CSectionSolver по Mx.
Private mSolverToleranceMy As Double ' Допуск равновесия внутреннего CSectionSolver по My.
Private mSolverMethod As String ' Newton или Secant для внутренних расчетов равновесия.
Private mSolverLineSearchEnabled As Boolean ' Включен ли line search во внутренних расчетах.
Private mSolverDampingInitial As Double ' Начальный коэффициент демпфирования Newton.
Private mSolverMinLineSearchAlpha As Double ' Минимальный alpha line search.
Private mSolverMaxDeltaEpsilon0 As Double ' Ограничение шага epsilon0.
Private mSolverMaxDeltaKappa As Double ' Ограничение шага кривизн.
Private mSolverSecantMaxRestarts As Long ' Максимум рестартов Secant-решателя.
Private mSolverSecantMinStepNorm As Double ' Минимальная норма шага Secant-решателя.
Private mDiagnosticsEnabled As Boolean ' True, если нужно писать подробный журнал итераций.

' Задает безопасные значения по умолчанию на случай прямых unit-тестов.
' Основной batch-сценарий всегда переопределяет их настройками Config.
Private Sub Class_Initialize()
    mInitialLambdaStep = 1#
    mSolutionStrategy = "Auto"
    mSearchMethod = "Bisection"
    mMaxLambda = 1024#
    mLambdaTolerance = 0.01
    mStrainTolerance = 0.00001
    mMaxRetries = 0
    mBaseLoadSteps = 1
    mCapacitySolverMaxIterations = 60
    mSolverToleranceN = 0.01
    mSolverToleranceMx = 0.01
    mSolverToleranceMy = 0.01
    mSolverMethod = "Newton"
    mSolverLineSearchEnabled = True
    mSolverDampingInitial = 1#
    mSolverMinLineSearchAlpha = 0.000001
    mSolverMaxDeltaEpsilon0 = 0#
    mSolverMaxDeltaKappa = 0#
    mSolverSecantMaxRestarts = 0
    mSolverSecantMinStepNorm = 0#
    mDiagnosticsEnabled = False
End Sub

' Принимает пользовательские настройки capacity-search, уже прочитанные из
' Config. Класс не обращается к Excel, но именно здесь они превращаются в
' настройки CCapacitySolver.
Public Sub ConfigureSearch(ByVal initialLambdaStep As Double, ByVal solutionStrategy As String, _
        ByVal searchMethod As String, ByVal maxLambda As Double, ByVal lambdaTolerance As Double, _
        ByVal strainTolerance As Double, ByVal maxRetries As Long, ByVal baseLoadSteps As Long, _
        ByVal capacitySolverMaxIterations As Long)
    mInitialLambdaStep = initialLambdaStep
    mSolutionStrategy = solutionStrategy
    mSearchMethod = searchMethod
    mMaxLambda = maxLambda
    mLambdaTolerance = lambdaTolerance
    mStrainTolerance = strainTolerance
    mMaxRetries = maxRetries
    mBaseLoadSteps = baseLoadSteps
    mCapacitySolverMaxIterations = capacitySolverMaxIterations
End Sub

' Принимает настройки внутреннего CSectionSolver, который вызывается для
' пробных lambda-точек Capacity. Это те же параметры равновесия, что применяет
' batch для прямого НДС, но владение ими в capacity-сценарии остается здесь.
Public Sub ConfigureEquilibriumSolver(ByVal toleranceN As Double, ByVal toleranceMx As Double, _
        ByVal toleranceMy As Double, ByVal solverMethod As String, ByVal lineSearchEnabled As Boolean, _
        ByVal dampingInitial As Double, ByVal minLineSearchAlpha As Double, _
        ByVal maxDeltaEpsilon0 As Double, ByVal maxDeltaKappa As Double, _
        ByVal secantMaxRestarts As Long, ByVal secantMinStepNorm As Double, _
        ByVal diagnosticsEnabled As Boolean)
    mSolverToleranceN = toleranceN
    mSolverToleranceMx = toleranceMx
    mSolverToleranceMy = toleranceMy
    mSolverMethod = solverMethod
    mSolverLineSearchEnabled = lineSearchEnabled
    mSolverDampingInitial = dampingInitial
    mSolverMinLineSearchAlpha = minLineSearchAlpha
    mSolverMaxDeltaEpsilon0 = maxDeltaEpsilon0
    mSolverMaxDeltaKappa = maxDeltaKappa
    mSolverSecantMaxRestarts = secantMaxRestarts
    mSolverSecantMinStepNorm = secantMinStepNorm
    mDiagnosticsEnabled = diagnosticsEnabled
End Sub

' Выполняет расчет несущей способности для одного LC.
' Все аргументы уже нормализованы во внутренние единицы и знаки; Config сюда
' напрямую не читается.
Public Function Calculate(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal strengthSpec As CMaterialModelSpec, ByVal loadPath As CCapacityLoadPath, _
        ByVal loadState As CSectionLoadState, ByVal strengthReferenceX As Double, _
        ByVal strengthReferenceY As Double, ByVal statusPolicy As CBatchStatusPolicy, _
        ByVal loadReferenceX As Double, ByVal loadReferenceY As Double, _
        ByVal momentFilter As CMomentZeroFilter, ByVal concreteDepthForMx As Double, _
        ByVal concreteDepthForMy As Double) As CCapacityResult
    Dim result As CCapacityResult
    Set result = New CCapacityResult
    If statusPolicy Is Nothing Then Set statusPolicy = New CBatchStatusPolicy

    If loadPath Is Nothing Then
        result.InitializeInvalidInput "CapacityLoadPath не подготовлен."
        Set Calculate = result
        Exit Function
    End If

    If Not loadPath.Valid Then
        result.InitializeInvalidInput loadPath.ErrorMessage
        Set Calculate = result
        Exit Function
    End If

    If Not loadPath.HasPath Then
        result.InitializeSkipped "ZeroBaseLoad"
        Set Calculate = result
        Exit Function
    End If

    If Not loadPath.HasScaledLoad Then
        result.InitializeInvalidInput _
            "Для выбранного CapacityLoadPath соответствующая нагрузка должна быть ненулевой."
        Set Calculate = result
        Exit Function
    End If

    Dim pureAxialForMethod As Boolean
    pureAxialForMethod = IsPureAxialForMethod(loadPath, loadState, strengthReferenceX, _
        strengthReferenceY, momentFilter, concreteDepthForMx, concreteDepthForMy)

    Dim configuredSolver As CCapacitySolver
    Set configuredSolver = CreateConfiguredSolver(concreteMaterial, steelMaterial, strengthSpec)

    Dim request As CLimitSearchRequest
    Set request = New CLimitSearchRequest
    request.InitializeCapacity configuredSolver, section, concreteMaterial, steelMaterial, _
        SolutionStrategy, loadPath.NOffset, loadPath.NBase, loadPath.MxOffset, _
        loadPath.MxBase, loadPath.MyOffset, loadPath.MyBase, loadPath.ForceOnly, _
        pureAxialForMethod, strengthSpec

    Dim coordinator As CLimitSearchCoordinator
    Set coordinator = New CLimitSearchCoordinator

    Dim searchResult As CLimitSearchResult
    Set searchResult = coordinator.ExecuteCapacity(request)
    result.InitializeFromSearch searchResult, loadPath, statusPolicy, _
        pureAxialForMethod, loadReferenceX, loadReferenceY, momentFilter, _
        concreteDepthForMx, concreteDepthForMy

    Set Calculate = result
End Function

' Возвращает нормализованную стратегию Capacity.SolutionStrategy.
' Пустая строка сохраняет прежнее поведение для невалидной настройки: дальше
' search-слой сформирует controlled InputErr, а не молча подставит fallback.
Public Property Get SolutionStrategy() As String
    Select Case LCase$(Trim$(mSolutionStrategy))
        Case "auto"
            SolutionStrategy = "Auto"
        Case "loadmultiplier"
            SolutionStrategy = "LoadMultiplier"
        Case "ultimatestrain"
            SolutionStrategy = "UltimateStrain"
    End Select
End Property

' Создает capacity load-path из пользовательского текста и уже нормализованного
' состояния нагрузки. Это часть инженерной подготовки Capacity, поэтому batch
' не должен иметь собственную копию такой фабрики.
Public Function BuildLoadPath(ByVal rawPath As String, _
        ByVal loadState As CSectionLoadState) As CCapacityLoadPath
    Dim loadPath As CCapacityLoadPath
    Set loadPath = New CCapacityLoadPath
    loadPath.InitializeFromLoadState rawPath, loadState
    Set BuildLoadPath = loadPath
End Function

' Возвращает пользовательскую подпись нормализованного capacity-path для
' Results и execution report. Соответствие ключей и подписей остается рядом с
' объектом траектории, а calculator дает batch-слою короткую точку входа.
Public Function DisplayNameForKey(ByVal loadPathKey As String) As String
    Dim loadPath As CCapacityLoadPath
    Set loadPath = New CCapacityLoadPath
    DisplayNameForKey = loadPath.DisplayNameForKey(loadPathKey)
End Function

' Формирует русское пояснение для execution report перед запуском общего
' search. Это не меняет расчет: текст только объясняет, какую стратегию и
' почему выбрал capacity-calculator для текущей траектории.
Public Function SearchPlanReport(ByVal loadPath As CCapacityLoadPath, _
        ByVal pureAxialForMethod As Boolean) As String
    If loadPath Is Nothing Then
        SearchPlanReport = "Capacity остановлен: CapacityLoadPath не подготовлен."
        Exit Function
    End If

    Select Case SolutionStrategy
        Case "Auto"
            If pureAxialForMethod Then
                SearchPlanReport = "Несущая способность: стратегия Auto. После переноса к центру приведенного сечения траектория является чисто осевой, поэтому сразу используется LoadMultiplier: для такой постановки UltimateStrain вырождается на горизонтальном участке диаграммы."
            ElseIf loadPath.ForceOnly Then
                SearchPlanReport = "Несущая способность: стратегия Auto. Пользователь масштабирует только N, но относительно центра приведенного сечения есть изгибающий момент; сначала пробуем UltimateStrain, затем при неудаче LoadMultiplier."
            Else
                SearchPlanReport = "Несущая способность: стратегия Auto сначала пробует UltimateStrain; если предельная плоскость не находится устойчиво, тот же путь нагрузки повторяется через LoadMultiplier."
            End If
        Case "UltimateStrain"
            If pureAxialForMethod Then
                SearchPlanReport = "Несущая способность: выбран UltimateStrain, но после переноса к центру приведенного сечения траектория чисто осевая и считается через LoadMultiplier как более устойчивая одномерная постановка."
            ElseIf loadPath.ForceOnly Then
                SearchPlanReport = "Несущая способность: выбран UltimateStrain. Пользователь масштабирует только N, но относительно центра приведенного сечения есть изгибающий момент; сначала пробуем UltimateStrain, затем при неудаче LoadMultiplier."
            Else
                SearchPlanReport = "Несущая способность: расчет идет напрямую через UltimateStrain по выбранной lambda-траектории."
            End If
        Case "LoadMultiplier"
            SearchPlanReport = "Несущая способность: расчет идет через LoadMultiplier, то есть программа подбирает lambda и для каждой lambda-точки заново решает НДС."
        Case Else
            SearchPlanReport = "Capacity остановлен: неизвестный Capacity.SolutionStrategy."
    End Select
End Function

' Определяет, остается ли силовая lambda-траектория чисто осевой после переноса
' к расчетному центру приведенного сечения. Это влияет только на выбор численной
' ветки Capacity: для действительно осевой траектории устойчивее сразу работать
' через LoadMultiplier, а не через вырожденный UltimateStrain.
Private Function IsPureAxialForMethod(ByVal loadPath As CCapacityLoadPath, _
        ByVal loadState As CSectionLoadState, ByVal referenceX As Double, _
        ByVal referenceY As Double, ByVal momentFilter As CMomentZeroFilter, _
        ByVal concreteDepthForMx As Double, ByVal concreteDepthForMy As Double) As Boolean
    If loadPath Is Nothing Then Exit Function
    If loadState Is Nothing Then Exit Function
    If Not loadPath.ForceOnly Then Exit Function

    IsPureAxialForMethod = MomentIsZeroForMethod(loadState.MxAboutPoint(referenceY), _
        momentFilter, concreteDepthForMx) And _
        MomentIsZeroForMethod(loadState.MyAboutPoint(referenceX), momentFilter, concreteDepthForMy)
End Function

' Проверяет инженерно нулевой момент тем же фильтром, которым batch уже
' нормализует нагрузки перед solver-ом. Без фильтра остается только маленький
' численный guard, чтобы прямые unit-тесты не зависели от Nothing-объекта.
Private Function MomentIsZeroForMethod(ByVal momentValue As Double, _
        ByVal momentFilter As CMomentZeroFilter, ByVal concreteDepth As Double) As Boolean
    If momentFilter Is Nothing Then
        MomentIsZeroForMethod = (Abs(momentValue) <= 0.000000001)
    Else
        MomentIsZeroForMethod = momentFilter.IsMomentEffectivelyZero(momentValue, concreteDepth)
    End If
End Function

' Создает и настраивает CCapacitySolver для текущего material spec.
' Материальные пределы берутся из уже построенных диаграмм; пользовательские
' search-настройки и настройки внутреннего равновесия применяются из полей
' CCapacityCalculator.
Private Function CreateConfiguredSolver(ByVal concreteMaterial As Object, _
        ByVal steelMaterial As Object, ByVal strengthSpec As CMaterialModelSpec) As CCapacitySolver
    Dim cap As CCapacitySolver
    Set cap = New CCapacitySolver

    cap.ConcreteCompressionLimit = concreteMaterial.UltimateCompressionStrain
    cap.ConcreteTensionLimit = concreteMaterial.UltimateTensionStrain
    If Not strengthSpec Is Nothing Then _
        cap.ConcreteTensionLimitEnabled = (concreteMaterial.HasPhysicalTensionLimit And _
            strengthSpec.ConcreteTension = ctkUseDiagram)
    Dim steelCompressionLimit As Double
    Dim steelTensionLimit As Double
    steelCompressionLimit = Abs(steelMaterial.UltimateCompressionStrain)
    steelTensionLimit = Abs(steelMaterial.UltimateTensionStrain)
    cap.SteelCompressionLimit = steelCompressionLimit
    cap.SteelTensionLimit = steelTensionLimit

    cap.LambdaTolerance = mLambdaTolerance
    cap.StrainTolerance = mStrainTolerance
    cap.SearchMethod = mSearchMethod
    cap.MaxLambda = mMaxLambda
    cap.InitialLambdaStep = mInitialLambdaStep
    cap.MaxRetries = mMaxRetries
    cap.SolverBaseLoadSteps = mBaseLoadSteps
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
    cap.DiagnosticsEnabled = mDiagnosticsEnabled

    Set CreateConfiguredSolver = cap
End Function
