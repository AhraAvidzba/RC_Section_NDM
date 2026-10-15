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
' Подготовка расчёта несущей способности
' ==========================================================================
' Получает готовые сечение, материалы, траекторию нагрузки и численные настройки.
' Создаёт запрос общего поиска предела и возвращает CCapacityResult.
' Скрывает детали поиска от пакетного расчёта, который получает готовый итог
' и отдельно выбирает определяющее сочетание.

Private mInitialLambdaStep As Double ' Первый шаг расширения lambda для LoadMultiplier.
Private mSolutionStrategy As String ' Пользовательская стратегия Capacity.SolutionStrategy после чтения Config.
Private mSearchMethod As String ' Алгоритм одномерного LoadMultiplier-поиска.
Private mMaxLambda As Double ' Защитная верхняя граница lambda при расширении скобки.
Private mLambdaTolerance As Double ' Допуск по lambda для одномерного поиска.
Private mStrainTolerance As Double ' Допуск попадания в предельную деформацию UltimateStrain.
Private mMaxRetries As Long ' Число повторов lambda-точки при численной неудаче.
Private mBaseLoadSteps As Long ' Базовое число load steps внутреннего CSectionSolver.
Private mCapacitySolverMaxIterations As Long ' Максимальное число итераций поиска предела и внутренних решений НДС.
Private mSolverToleranceN As Double ' Допуск равновесия внутреннего CSectionSolver по N.
Private mSolverToleranceMx As Double ' Допуск равновесия внутреннего CSectionSolver по Mx.
Private mSolverToleranceMy As Double ' Допуск равновесия внутреннего CSectionSolver по My.
Private mSolverMethod As String ' Newton или Secant для внутренних расчетов равновесия.
Private mSolverSecantMaxRestarts As Long ' Максимум рестартов Secant-решателя.
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
    mSolverSecantMaxRestarts = 0
    mDiagnosticsEnabled = False
End Sub

' Принимает прочитанные настройки поиска несущей способности
' и сохраняет их для CCapacitySolver без обращения к Excel.
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

' Проверяет численные настройки несущей способности до подготовки поиска.
' При чтении Config добавляет фактический адрес ошибки. Проверяет и параметры
' других доступных стратегий, не запуская НДС и не меняя критерии материалов.
Public Sub ValidateSettings(Optional ByVal settings As CSystemSettingsReader = Nothing)
    If mCapacitySolverMaxIterations <= 0 Then RaiseInvalidSearchSetting "Capacity.SolverMaxIterations", _
        "Число итераций должно быть положительным.", "Введите целое число больше нуля.", settings
    If mStrainTolerance <= 0# Then RaiseInvalidSearchSetting "Capacity.ToleranceStrain", _
        "Допуск предельной деформации должен быть положительным.", "Введите число больше нуля.", settings
    If mMaxLambda <= 0# Then RaiseInvalidSearchSetting "Capacity.MaxLambda", _
        "Верхняя граница поиска должна быть положительной.", "Введите число больше нуля.", settings
    If mInitialLambdaStep <= 0# Then RaiseInvalidSearchSetting "Capacity.InitialLambda", _
        "Начальный множитель нагрузки должен быть положительным.", "Введите число больше нуля.", settings
    If mLambdaTolerance <= 0# Then RaiseInvalidSearchSetting "Capacity.ToleranceLambda", _
        "Допуск поиска множителя нагрузки должен быть положительным.", "Введите число больше нуля.", settings
    If mMaxRetries < 0 Then RaiseInvalidSearchSetting "Capacity.MaxRetries", _
        "Число дополнительных попыток не может быть отрицательным.", "Введите целое число не меньше нуля; 0 отключает повторы.", settings
    If mBaseLoadSteps <= 0 Then RaiseInvalidSearchSetting "Capacity.BaseLoadSteps", _
        "Число ступеней нагрузки должно быть положительным.", "Введите целое число больше нуля.", settings
    Dim solver As CCapacitySolver, maximumRetries As Long
    Set solver = New CCapacitySolver
    maximumRetries = solver.MaximumRetryCountForLoadSteps(mBaseLoadSteps)
    If maximumRetries < 0 Then RaiseInvalidSearchSetting "Capacity.BaseLoadSteps", _
        "Число ступеней не помещается в расчетный счетчик.", "Уменьшите значение до 2147483646 или менее.", settings
    If mMaxRetries > maximumRetries Then RaiseInvalidSearchSetting "Capacity.MaxRetries", _
        "Число ступеней с учетом повторов не помещается в расчетный счетчик.", _
        "Уменьшите число повторов до " & CStr(maximumRetries) & " или уменьшите Capacity.BaseLoadSteps.", settings
End Sub

' Сообщает об ошибке настройки с её именем и способом исправления.
' Адрес берётся из прочитанного Config; при отдельном программном вызове
' сохраняются причина и ключ параметра без вымышленного адреса ячейки.
Private Sub RaiseInvalidSearchSetting(ByVal key As String, ByVal reason As String, _
        ByVal correction As String, ByVal settings As CSystemSettingsReader)
    Dim message As String
    If settings Is Nothing Then
        message = "Настройка " & key & ": " & reason & " " & correction
    Else
        message = settings.InputErrorMessage(key, reason, correction)
    End If
    Err.Raise vbObjectError + 4190, "CCapacityCalculator", message
End Sub

' Принимает настройки CSectionSolver для пробных нагрузок при поиске предела.
' Их смысл совпадает с настройками обычного НДС, но здесь они сохраняются
' для самостоятельного расчёта несущей способности.
Public Sub ConfigureEquilibriumSolver(ByVal toleranceN As Double, ByVal toleranceMx As Double, _
        ByVal toleranceMy As Double, ByVal solverMethod As String, _
        ByVal secantMaxRestarts As Long, _
        ByVal diagnosticsEnabled As Boolean)
    mSolverToleranceN = toleranceN
    mSolverToleranceMx = toleranceMx
    mSolverToleranceMy = toleranceMy
    mSolverMethod = solverMethod
    mSolverSecantMaxRestarts = secantMaxRestarts
    mDiagnosticsEnabled = diagnosticsEnabled
End Sub

' Выполняет расчет несущей способности для одного LC.
' Все аргументы уже нормализованы во внутренние единицы и знаки; Config сюда
' напрямую не читается.
Public Function Calculate(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal strengthSpec As CMaterialModelSpec, ByVal loadPath As CLoadPathDescriptor, _
        ByVal loadState As CSectionLoadState, ByVal strengthReferenceX As Double, _
        ByVal strengthReferenceY As Double, _
        ByVal loadReferenceX As Double, ByVal loadReferenceY As Double, _
        ByVal momentFilter As CMomentZeroFilter, ByVal concreteDepthForMx As Double, _
        ByVal concreteDepthForMy As Double) As CCapacityResult
    Dim result As CCapacityResult
    Set result = New CCapacityResult
    If loadPath Is Nothing Then
        result.InitializeInvalidInput "LoadPath не подготовлен."
        Set Calculate = result
        Exit Function
    End If

    If Not loadPath.Valid Then
        result.InitializeInvalidInput loadPath.ErrorMessage
        Set Calculate = result
        Exit Function
    End If

    If Not loadPath.HasPath Then
        result.InitializeSkipped "Поиск несущей способности не выполнялся: масштабируемая нагрузка отсутствует."
        Set Calculate = result
        Exit Function
    End If

    If Not loadPath.HasScaledLoad Then
        result.InitializeInvalidInput _
            "Для выбранного LoadPath соответствующая нагрузка должна быть ненулевой."
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
    InitializeCapacitySearchRequest request, configuredSolver, section, concreteMaterial, steelMaterial, _
        SolutionStrategy, loadPath.NOffset, loadPath.NBase, loadPath.MxOffset, _
        loadPath.MxBase, loadPath.MyOffset, loadPath.MyBase, loadPath.ForceOnly, _
        pureAxialForMethod, strengthSpec
    Dim axialPath As CLoadPathDescriptor
    Set axialPath = BuildLoadPath("LambdaN", loadState)
    request.ConfigureLoadPathSelection loadPath.Key, loadState, _
        IsPureAxialForMethod(axialPath, loadState, strengthReferenceX, strengthReferenceY, _
            momentFilter, concreteDepthForMx, concreteDepthForMy)

    Dim coordinator As CLimitSearchCoordinator
    Set coordinator = New CLimitSearchCoordinator

    Dim searchResult As CLimitSearchResult
    Set searchResult = coordinator.ExecuteCapacity(request)
    result.InitializeFromSearch searchResult, request.SelectedLoadPath, _
        request.PureAxialForMethod, loadReferenceX, loadReferenceY, momentFilter, _
        concreteDepthForMx, concreteDepthForMy

    Set Calculate = result
End Function

' Возвращает проверенное написание Capacity.SolutionStrategy.
' Неизвестное значение даёт пустую строку и затем InputErr;
' другая стратегия вместо ошибочно заданной не подставляется.
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

' Готовит выбранное изменение нагрузки по тексту пользователя
' и общему объекту усилий. Пакетный расчёт использует этот метод
' без отдельной копии правил построения нагрузки.
Public Function BuildLoadPath(ByVal rawPath As String, _
        ByVal loadState As CSectionLoadState) As CLoadPathDescriptor
    Dim loadPath As CLoadPathDescriptor
    Set loadPath = New CLoadPathDescriptor
    loadPath.InitializeFromLoadState rawPath, loadState
    Set BuildLoadPath = loadPath
End Function

' Возвращает название принятого изменения нагрузки для Results и отчёта.
' Общие соответствия ключей и названий хранит CLoadPathDescriptor.
Public Function DisplayNameForKey(ByVal loadPathKey As String) As String
    Dim loadPath As CLoadPathDescriptor
    Set loadPath = New CLoadPathDescriptor
    DisplayNameForKey = loadPath.DisplayNameForKey(loadPathKey)
End Function

' Объясняет для отчёта выполнения, какой способ поиска выбран для текущей
' траектории нагрузки и почему. Численные результаты этот текст не меняет.
Public Function SearchPlanReport(ByVal loadPath As CLoadPathDescriptor, _
        ByVal pureAxialForMethod As Boolean) As String
    If loadPath Is Nothing Then
        SearchPlanReport = "Capacity остановлен: LoadPath не подготовлен."
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

' Проверяет, осталось ли изменение нагрузки чисто осевым после переноса
' к центру приведённого сечения. Для такой траектории выбирается устойчивый
' LoadMultiplier вместо вырожденной постановки UltimateStrain.
Private Function IsPureAxialForMethod(ByVal loadPath As CLoadPathDescriptor, _
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

' Проверяет малый момент тем же фильтром, что обычный пакетный расчёт.
' Если фильтр не передан, учитывает лишь небольшую численную погрешность
' для самостоятельного программного вызова.
Private Function MomentIsZeroForMethod(ByVal momentValue As Double, _
        ByVal momentFilter As CMomentZeroFilter, ByVal concreteDepth As Double) As Boolean
    If momentFilter Is Nothing Then
        MomentIsZeroForMethod = (Abs(momentValue) <= 0.000000001)
    Else
        MomentIsZeroForMethod = momentFilter.IsMomentEffectivelyZero(momentValue, concreteDepth)
    End If
End Function

' Создаёт CCapacitySolver с фактическими диаграммами и пределами материалов.
' Применяет сохранённые настройки поиска и внутреннего решения равновесия.
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
    cap.SolverSecantMaxRestarts = mSolverSecantMaxRestarts
    cap.DiagnosticsEnabled = mDiagnosticsEnabled

    Set CreateConfiguredSolver = cap
End Function
