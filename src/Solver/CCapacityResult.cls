VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCapacityResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Инженерный результат расчета несущей способности
' ==========================================================================
' CCapacityResult отделяет найденную limit-search точку от batch-таблицы.
' Класс принимает уже выполненный общий LimitSearch, вытаскивает из него
' физически значимые capacity-величины, переносит полный предельный вектор в
' пользовательскую точку приложения нагрузки и при необходимости готовит
' named-state CapacityState для Results snapshot.
'
' Здесь нет выбора search-алгоритма и нет чтения Config. Эти обязанности
' остаются соответственно в CLimitSearch* и вызывающем orchestration-слое.

Private Const CAPACITY_PATH_NONE As String = "None" ' Единый маркер: capacity для LC не выполнялся.

Private mSearchResult As CLimitSearchResult ' Типизированный результат общего LimitSearch.
Private mStateResult As CSectionStateResult ' Итоговый reusable named-state CapacityState, если он найден.
Private mStatus As String ' Пользовательский статус capacity для существующего вывода.
Private mLimitState As String ' Физическая причина предела или текст ошибки.
Private mSolutionMethod As String ' Фактически принятая ветка: UltimateStrain или LoadMultiplier.
Private mPathResolved As String ' Нормализованный ключ capacity load path.
Private mLambdaCapacity As Double ' Найденный множитель lambda.
Private mUtilCapacity As Double ' Использование несущей способности.
Private mNUltimate As Double ' Предельная N в точке приложения нагрузки.
Private mMxUltimate As Double ' Предельный Mx в точке приложения нагрузки.
Private mMyUltimate As Double ' Предельный My в точке приложения нагрузки.
Private mCapacityStateN As Double ' Полная N найденного CapacityState во внутреннем начале модели.
Private mCapacityStateMx As Double ' Полный Mx найденного CapacityState во внутреннем начале модели.
Private mCapacityStateMy As Double ' Полный My найденного CapacityState во внутреннем начале модели.
Private mMomentUltimate As Double ' Модуль полного предельного момента в точке приложения нагрузки.
Private mCapacityStateAvailable As Boolean ' True, если есть найденная плоскость деформаций CapacityState.
Private mCapacityEpsilon0 As Double ' epsilon0 найденного CapacityState.
Private mCapacityKappaX As Double ' kappaX найденного CapacityState.
Private mCapacityKappaY As Double ' kappaY найденного CapacityState.
Private mPureAxialForMethod As Boolean ' True, если capacity-calculator выбрал осевую численную ветку после переноса нагрузок к центру.
Private mResultMeta As CResultMeta ' Типизированный статус CapacityResult для общего result/status слоя.

' Инициализирует meta как неприменимую ветку, чтобы даже ранний выход имел
' безопасное машинное состояние.
Private Sub Class_Initialize()
    Set mResultMeta = New CResultMeta
    mResultMeta.SetNotApplicable rkCapacity
End Sub

' Формирует результат по успешному или неуспешному search-слою.
' Даже при numerical/input failure объект остается заполненным: batch получает
' единый контейнер и не разбирает поля CCapacitySolver вручную.
Public Sub InitializeFromSearch(ByVal searchResult As CLimitSearchResult, _
        ByVal loadPath As CCapacityLoadPath, ByVal strengthSpec As CMaterialModelSpec, _
        ByVal statusPolicy As CBatchStatusPolicy, ByVal pureAxialForMethod As Boolean, _
        ByVal loadReferenceX As Double, ByVal loadReferenceY As Double, _
        ByVal momentFilter As CMomentZeroFilter, ByVal concreteDepthForMx As Double, _
        ByVal concreteDepthForMy As Double)
    Set mSearchResult = searchResult
    mPureAxialForMethod = pureAxialForMethod
    If statusPolicy Is Nothing Then Set statusPolicy = New CBatchStatusPolicy

    If loadPath Is Nothing Then
        InitializeInvalid statusPolicy.InputErr, "CapacityLoadPath не подготовлен.", CAPACITY_PATH_NONE
        Exit Sub
    End If

    mPathResolved = loadPath.Key
    If searchResult Is Nothing Then
        Dim reason As String
        If Not searchResult Is Nothing Then reason = searchResult.Meta.ResultComment
        If Len(Trim$(reason)) = 0 Then reason = "Поиск несущей способности не вернул solver."
        InitializeInvalid statusPolicy.InputErr, reason, loadPath.Key
        Exit Sub
    End If

    If searchResult.ResultKind <> rkCapacity Then
        InitializeInvalid statusPolicy.CalcErr, "LimitSearch returned non-capacity result.", loadPath.Key
        Exit Sub
    End If

    Set mResultMeta = searchResult.Meta.Clone
    mStatus = statusPolicy.StatusFromMeta(mResultMeta)
    mLimitState = searchResult.LimitState
    If Len(Trim$(mLimitState)) = 0 Then mLimitState = mResultMeta.ResultComment
    mSolutionMethod = searchResult.ActualMethod
    mLambdaCapacity = searchResult.LambdaUltimate

    StoreUltimateVector loadPath, loadReferenceX, loadReferenceY, momentFilter, _
        concreteDepthForMx, concreteDepthForMy
    StoreUtilization
    StoreCapacityStateIfAvailable strengthSpec
End Sub

' Заполняет результат известной ошибкой до запуска search.
' Такой путь используется для некорректного или неприменимого capacity path.
Public Sub InitializeInvalid(ByVal statusText As String, ByVal reason As String, _
        Optional ByVal pathKey As String = CAPACITY_PATH_NONE)
    Set mSearchResult = Nothing
    mStatus = statusText
    mLimitState = reason
    mSolutionMethod = vbNullString
    mPathResolved = pathKey
    mLambdaCapacity = 0#
    mUtilCapacity = 0#
    mNUltimate = 0#
    mMxUltimate = 0#
    mMyUltimate = 0#
    mCapacityStateN = 0#
    mCapacityStateMx = 0#
    mCapacityStateMy = 0#
    mMomentUltimate = 0#
    mCapacityStateAvailable = False
    mPureAxialForMethod = False
    Set mStateResult = Nothing
    StoreMetaFromStatus statusText, reason
End Sub

' Заполняет результат штатным пропуском capacity без запуска search.
' Нулевая траектория не является ошибкой ввода: для нее просто нет направления,
' по которому можно искать предельный множитель lambda.
Public Sub InitializeSkipped(ByVal statusText As String, ByVal reason As String, _
        Optional ByVal pathKey As String = CAPACITY_PATH_NONE)
    InitializeInvalid statusText, reason, pathKey
End Sub

' Переносит найденный полный предельный вектор в пользовательскую точку
' приложения нагрузки через общий CSectionLoadState.
Private Sub StoreUltimateVector(ByVal loadPath As CCapacityLoadPath, _
        ByVal loadReferenceX As Double, ByVal loadReferenceY As Double, _
        ByVal momentFilter As CMomentZeroFilter, ByVal concreteDepthForMx As Double, _
        ByVal concreteDepthForMy As Double)
    If mSearchResult Is Nothing Then Exit Sub
    If mSearchResult.LambdaUltimate <= 0# Then Exit Sub

    Dim nValue As Double
    Dim mxAboutOrigin As Double
    Dim myAboutOrigin As Double

    If Not mSearchResult.LastSolver Is Nothing Then
        nValue = mSearchResult.LastSolver.Nint
        mxAboutOrigin = mSearchResult.LastSolver.Mxint
        myAboutOrigin = mSearchResult.LastSolver.Myint
        mCapacityStateN = nValue
        mCapacityStateMx = mxAboutOrigin
        mCapacityStateMy = myAboutOrigin
    Else
        loadPath.Target mSearchResult.LambdaUltimate, nValue, mxAboutOrigin, myAboutOrigin
        mCapacityStateN = nValue
        mCapacityStateMx = mxAboutOrigin
        mCapacityStateMy = myAboutOrigin
    End If

    Dim loadState As CSectionLoadState
    Set loadState = New CSectionLoadState
    loadState.InitializeFromPointMoments nValue, mxAboutOrigin, myAboutOrigin, _
        0#, 0#, loadReferenceX, loadReferenceY, momentFilter, _
        concreteDepthForMx, concreteDepthForMy

    mNUltimate = loadState.N
    mMxUltimate = loadState.LoadPointMx
    mMyUltimate = loadState.LoadPointMy
    mMomentUltimate = Sqr(mMxUltimate * mMxUltimate + mMyUltimate * mMyUltimate)
End Sub

' Сохраняет utilization по тем же правилам, что раньше применялись в batch.
' Если solver сам посчитал utilization, он является приоритетным источником.
Private Sub StoreUtilization()
    If mSearchResult Is Nothing Then Exit Sub
    If mSearchResult.LambdaUltimate > 0# Then mUtilCapacity = 1# / mSearchResult.LambdaUltimate
    If mSearchResult.Converged Then mUtilCapacity = mSearchResult.Utilization
End Sub

' Готовит CapacityState только для физически найденной предельной точки.
' Неуспешные probes и numerical failure не становятся reusable named-state.
Private Sub StoreCapacityStateIfAvailable(ByVal strengthSpec As CMaterialModelSpec)
    If mSearchResult Is Nothing Then Exit Sub
    If Not mSearchResult.Converged Then Exit Sub
    If mSearchResult.LambdaUltimate <= 0# Then Exit Sub
    If mSearchResult.LastSolver Is Nothing Then Exit Sub

    mCapacityStateAvailable = True
    mCapacityEpsilon0 = mSearchResult.LastSolver.Epsilon0
    mCapacityKappaX = mSearchResult.LastSolver.KappaX
    mCapacityKappaY = mSearchResult.LastSolver.KappaY

    Set mStateResult = New CSectionStateResult
    mStateResult.InitializeFromSolver sstCapacityState, cpStrength, strengthSpec, _
        mSearchResult.LastSolver, False, mStatus, True, 0, mSearchResult.DiagnosticLog
End Sub

Public Property Get SearchResult() As CLimitSearchResult
    Set SearchResult = mSearchResult
End Property

' Возвращает типизированный статус capacity-результата. CombinationResult
' хранит эту meta рядом со строковым статусом, чтобы итоговый OverallStatus
' собирался из InternalStatus/ResultCode, а не из повторного парсинга текста.
Public Property Get ResultMeta() As CResultMeta
    If mResultMeta Is Nothing Then StoreMetaFromStatus mStatus, mLimitState
    Set ResultMeta = mResultMeta.Clone
End Property

' Возвращает подробный журнал поиска через общий search-result.
' Batch-слой не должен обращаться к внутреннему DiagnosticLog solver-а напрямую.
Public Property Get DiagnosticLog() As String
    If Not mSearchResult Is Nothing Then DiagnosticLog = mSearchResult.DiagnosticLog
End Property

' Возвращает True, если для capacity действительно запускался LimitSearch.
' Ранние исходы вроде N/A или InputErr остаются полноценным результатом,
' но не содержат solver-фасад и не должны трактоваться batch-слоем как search.
Public Property Get SearchExecuted() As Boolean
    If Not mSearchResult Is Nothing Then SearchExecuted = mSearchResult.SearchExecuted
End Property

' Возвращает диагностический журнал последнего внутреннего CSectionSolver,
' если предельная lambda-точка действительно была решена равновесием.
Public Property Get LastSolverDiagnosticLog() As String
    If mSearchResult Is Nothing Then Exit Property
    LastSolverDiagnosticLog = mSearchResult.LastSolverDiagnosticLog
End Property

' Формирует русское пояснение ранней остановки до запуска LimitSearch.
' Batch-слой только выводит эту строку в отчет и не повторяет domain-логику.
Public Function EarlyStopReport(ByVal loadPath As CCapacityLoadPath) As String
    If loadPath Is Nothing Then
        EarlyStopReport = "Capacity остановлен: CapacityLoadPath не подготовлен."
    ElseIf Not loadPath.Valid Then
        EarlyStopReport = "Capacity остановлен: неизвестное значение CapacityLoadPath."
    ElseIf Not loadPath.HasPath Then
        EarlyStopReport = "Capacity не запускается: в сочетании нет ни момента, ни продольной силы."
    ElseIf Not loadPath.HasScaledLoad Then
        EarlyStopReport = "Capacity остановлен: выбранная lambda-траектория не содержит ненулевой масштабируемой нагрузки."
    Else
        EarlyStopReport = "Capacity остановлен: " & mLimitState
    End If
End Function

' Формирует пояснение, какая ветка поиска несущей была фактически принята.
' Сам выбор уже сделан search-слоем; здесь остается только инженерная трактовка
' для execution report.
Public Function MethodChoiceReport(ByVal solutionStrategy As String, _
        ByVal loadPath As CCapacityLoadPath) As String
    If Len(Trim$(mSolutionMethod)) = 0 Then Exit Function

    If StrComp(solutionStrategy, "Auto", vbTextCompare) = 0 Then
        If StrComp(mSolutionMethod, "UltimateStrain", vbTextCompare) = 0 Then
            MethodChoiceReport = "Итог поиска несущей: Auto принял результат UltimateStrain, резервный переход на LoadMultiplier не потребовался."
        ElseIf StrComp(mSolutionMethod, "LoadMultiplier", vbTextCompare) = 0 Then
            If mPureAxialForMethod Then
                MethodChoiceReport = "Итог поиска несущей: Auto использовал LoadMultiplier сразу, потому что траектория чисто осевая относительно центра приведенного сечения."
            Else
                MethodChoiceReport = "Итог поиска несущей: Auto переключился на LoadMultiplier, потому что UltimateStrain не дал устойчивую предельную точку; подробная причина ниже в диагностике solver-а."
            End If
        End If
    ElseIf StrComp(solutionStrategy, "UltimateStrain", vbTextCompare) = 0 And _
            StrComp(mSolutionMethod, "LoadMultiplier", vbTextCompare) = 0 Then
        If mPureAxialForMethod Then
            MethodChoiceReport = "Итог поиска несущей: фактически использован LoadMultiplier, так как траектория чисто осевая относительно центра приведенного сечения и не решается устойчиво прямым UltimateStrain."
        ElseIf Not loadPath Is Nothing And loadPath.ForceOnly Then
            MethodChoiceReport = "Итог поиска несущей: UltimateStrain не дал устойчивую предельную точку для силовой траектории с эксцентриситетом; фактически использован LoadMultiplier."
        Else
            MethodChoiceReport = "Итог поиска несущей: фактически использован LoadMultiplier."
        End If
    End If
End Function

Public Property Get StateResult() As CSectionStateResult
    Set StateResult = mStateResult
End Property

Public Property Get Status() As String
    Status = mStatus
End Property

Public Property Get LimitState() As String
    LimitState = mLimitState
End Property

Public Property Get SolutionMethod() As String
    SolutionMethod = mSolutionMethod
End Property

Public Property Get PathResolved() As String
    PathResolved = mPathResolved
End Property

Public Property Get LambdaCapacity() As Double
    LambdaCapacity = mLambdaCapacity
End Property

Public Property Get UtilCapacity() As Double
    UtilCapacity = mUtilCapacity
End Property

Public Property Get NUltimate() As Double
    NUltimate = mNUltimate
End Property

Public Property Get MxUltimate() As Double
    MxUltimate = mMxUltimate
End Property

Public Property Get MyUltimate() As Double
    MyUltimate = mMyUltimate
End Property

Public Property Get CapacityStateN() As Double
    CapacityStateN = mCapacityStateN
End Property

Public Property Get CapacityStateMx() As Double
    CapacityStateMx = mCapacityStateMx
End Property

Public Property Get CapacityStateMy() As Double
    CapacityStateMy = mCapacityStateMy
End Property

Public Property Get MomentUltimate() As Double
    MomentUltimate = mMomentUltimate
End Property

Public Property Get CapacityStateAvailable() As Boolean
    CapacityStateAvailable = mCapacityStateAvailable
End Property

Public Property Get CapacityEpsilon0() As Double
    CapacityEpsilon0 = mCapacityEpsilon0
End Property

Public Property Get CapacityKappaX() As Double
    CapacityKappaX = mCapacityKappaX
End Property

Public Property Get CapacityKappaY() As Double
    CapacityKappaY = mCapacityKappaY
End Property

Public Property Get PureAxialForMethod() As Boolean
    PureAxialForMethod = mPureAxialForMethod
End Property

' Совместимо строит meta для ранних исходов, где общего search-result еще нет.
' Это закрывает skipped/input-error capacity-ветки тем же типизированным путем.
Private Sub StoreMetaFromStatus(ByVal statusText As String, ByVal resultComment As String)
    Dim policy As CResultStatusPolicy
    Set policy = New CResultStatusPolicy
    Set mResultMeta = policy.MetaFromLegacyStatus(statusText, rkCapacity, resultComment)
End Sub
