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
Private mLimitState As String ' Физическая причина предела или текст ошибки.
Private mSolutionMethod As String ' Фактически принятая ветка: UltimateStrain или LoadMultiplier.
Private mPathResolved As String ' Нормализованный ключ capacity load path.
Private mLambdaCapacity As Double ' Найденный множитель lambda.
Private mUtilCapacity As Double ' Использование несущей способности.
Private mNUltimate As Double ' Предельная N в точке приложения нагрузки.
Private mMxUltimate As Double ' Предельный Mx в точке приложения нагрузки.
Private mMyUltimate As Double ' Предельный My в точке приложения нагрузки.
Private mMomentUltimate As Double ' Модуль полного предельного момента в точке приложения нагрузки.
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
        ByVal loadPath As CCapacityLoadPath, _
        ByVal pureAxialForMethod As Boolean, _
        ByVal loadReferenceX As Double, ByVal loadReferenceY As Double, _
        ByVal momentFilter As CMomentZeroFilter, ByVal concreteDepthForMx As Double, _
        ByVal concreteDepthForMy As Double)
    ResetNumericSnapshot
    Set mSearchResult = searchResult
    mPureAxialForMethod = pureAxialForMethod
    If loadPath Is Nothing Then
        InitializeInvalidInput "CapacityLoadPath не подготовлен.", CAPACITY_PATH_NONE
        Exit Sub
    End If

    mPathResolved = loadPath.Key
    If searchResult Is Nothing Then
        InitializeCalcError "Поиск несущей способности не вернул результат.", loadPath.Key
        Exit Sub
    End If

    If searchResult.ResultKind <> rkCapacity Then
        InitializeCalcError "Поиск вернул результат, не относящийся к несущей способности.", loadPath.Key
        Exit Sub
    End If

    Set mResultMeta = searchResult.Meta.Clone
    ' Численный поиск мог успешно найти предел ниже заданной нагрузки.
    ' Только инженерный слой трактует это как непрохождение текущего LC.
    If searchResult.Succeeded And searchResult.Converged Then
        If searchResult.LambdaUltimate > 0# And searchResult.LambdaUltimate < 1# Then
            mResultMeta.SetResult rsCheckFailed, rcPhysicalLimitExceeded, rkCapacity, _
                searchResult.Meta.ResultComment
        End If
    End If
    mLimitState = searchResult.LimitState
    If Len(Trim$(mLimitState)) = 0 Then mLimitState = mResultMeta.ResultComment
    mSolutionMethod = searchResult.ActualMethod
    mLambdaCapacity = searchResult.LambdaUltimate

    StoreUltimateVector loadPath, loadReferenceX, loadReferenceY, momentFilter, _
        concreteDepthForMx, concreteDepthForMy
    StoreUtilization
    StoreCapacityStateIfAvailable
End Sub

' Заполняет результат ошибкой входных данных до запуска search.
Public Sub InitializeInvalidInput(ByVal reason As String, _
        Optional ByVal pathKey As String = CAPACITY_PATH_NONE)
    ' При ошибке разбора пути он не выбран; служебный None не является
    ' фактически выполненной траекторией и не должен подменять пустое поле.
    If StrComp(pathKey, CAPACITY_PATH_NONE, vbTextCompare) = 0 Then pathKey = vbNullString
    InitializeEarlyResult rsInvalidInput, rcInvalidInput, reason, pathKey
End Sub

' Заполняет результат внутренней ошибкой до запуска search.
Public Sub InitializeCalcError(ByVal reason As String, _
        Optional ByVal pathKey As String = CAPACITY_PATH_NONE)
    InitializeEarlyResult rsInternalError, rcInternalError, reason, pathKey
End Sub

' Заполняет результат штатным пропуском capacity без запуска search.
' Нулевая траектория не является ошибкой ввода: для нее просто нет направления,
' по которому можно искать предельный множитель lambda.
Public Sub InitializeSkipped(ByVal reason As String, _
        Optional ByVal pathKey As String = CAPACITY_PATH_NONE)
    InitializeEarlyResult rsNotApplicable, rcNotApplicable, reason, pathKey
End Sub

' Заполняет результат, когда базовая точка lambda=0 уже нарушает критерий
' несущей способности. Для Capacity такая связка отображается как BaseFail.
Public Sub InitializeBaseFail(ByVal reason As String, _
        Optional ByVal pathKey As String = CAPACITY_PATH_NONE)
    InitializeEarlyResult rsCheckFailed, rcInitialStateBeyondLimit, NormalizeBaseFailReason(reason), pathKey, True
End Sub

' Убирает старые технические формулировки из пользовательского BaseFail.
' Машинный код причины остается в ResultCode = INITIAL_STATE_BEYOND_LIMIT.
Private Function NormalizeBaseFailReason(ByVal reason As String) As String
    NormalizeBaseFailReason = Trim$(reason)
    NormalizeBaseFailReason = Replace$(NormalizeBaseFailReason, _
        "Начальное состояние при lambda = 0 не проходит физический критерий:", _
        "Несущая способность не проходит уже для исходной части выбранного пути:", _
        1, 1, vbTextCompare)
    If Len(NormalizeBaseFailReason) = 0 Then _
        NormalizeBaseFailReason = "Несущая способность не проходит уже для исходной части выбранного пути."
End Function

' Общая запись раннего capacity-result без обратного маппинга из display-строки.
Private Sub InitializeEarlyResult(ByVal internalStatus As EResultInternalStatus, _
        ByVal resultCode As EResultCode, ByVal reason As String, _
        ByVal pathKey As String, Optional ByVal calculated As Boolean = False)
    ResetNumericSnapshot
    mLimitState = reason
    mSolutionMethod = vbNullString
    mPathResolved = pathKey
    Set mResultMeta = New CResultMeta
    mResultMeta.SetResult internalStatus, resultCode, rkCapacity, reason, vbNullString, True, calculated

End Sub

' Исключает утечку старой предельной точки при повторном заполнении объекта.
' Отсутствующие числа не заменяются прошлым успешным состоянием.
Private Sub ResetNumericSnapshot()
    Set mSearchResult = Nothing
    Set mStateResult = Nothing
    mLambdaCapacity = 0#
    mUtilCapacity = 0#
    mNUltimate = 0#
    mMxUltimate = 0#
    mMyUltimate = 0#
    mMomentUltimate = 0#
    mPureAxialForMethod = False
End Sub

' Переносит найденный полный предельный вектор в пользовательскую точку
' приложения нагрузки через общий CSectionLoadState.
Private Sub StoreUltimateVector(ByVal loadPath As CCapacityLoadPath, _
        ByVal loadReferenceX As Double, ByVal loadReferenceY As Double, _
        ByVal momentFilter As CMomentZeroFilter, ByVal concreteDepthForMx As Double, _
        ByVal concreteDepthForMy As Double)
    If mSearchResult Is Nothing Then Exit Sub
    If Not mSearchResult.HasLimitPoint Then Exit Sub

    Dim nValue As Double
    Dim mxAboutOrigin As Double
    Dim myAboutOrigin As Double

    nValue = mSearchResult.NUltimate
    mxAboutOrigin = mSearchResult.MxUltimate
    myAboutOrigin = mSearchResult.MyUltimate

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

' Определяет использование только по действительно найденному пределу.
' Последняя диагностическая проба и техническая граница не являются capacity.
Private Sub StoreUtilization()
    mUtilCapacity = 0#
    If mSearchResult Is Nothing Then Exit Sub
    If Not mSearchResult.HasLimitPoint Then Exit Sub
    If mSearchResult.LambdaUltimate > 0# Then mUtilCapacity = 1# / mSearchResult.LambdaUltimate
End Sub

' Готовит CapacityState только для физически найденной предельной точки.
' Неуспешные probes и numerical failure не становятся reusable named-state.
Private Sub StoreCapacityStateIfAvailable()
    Set mStateResult = Nothing
    If mSearchResult Is Nothing Then Exit Sub
    If Not mSearchResult.HasLimitPoint Then Exit Sub
    Set mStateResult = mSearchResult.PointState
End Sub

Public Property Get SearchResult() As CLimitSearchResult
    Set SearchResult = mSearchResult
End Property

' Возвращает независимый снимок meta канонического capacity-результата.
' Display-статус вычисляется из нее общей policy, отдельной строковой копии нет.
Public Property Get ResultMeta() As CResultMeta
    If mResultMeta Is Nothing Then InitializeCalcError "CapacityResult ResultMeta не сформирован.", mPathResolved
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
        Else
            MethodChoiceReport = "Итог поиска несущей: фактически использован LoadMultiplier."
            If Not loadPath Is Nothing Then
                If loadPath.ForceOnly Then _
                    MethodChoiceReport = "Итог поиска несущей: UltimateStrain не дал устойчивую предельную точку для силовой траектории с эксцентриситетом; фактически использован LoadMultiplier."
            End If
        End If
    End If
End Function

Public Property Get StateResult() As CSectionStateResult
    Set StateResult = mStateResult
End Property

Public Property Get Status() As String
    Dim policy As CResultStatusPolicy
    Set policy = New CResultStatusPolicy
    Status = policy.ExternalStatus(ResultMeta)
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
    If Not mStateResult Is Nothing Then CapacityStateN = mStateResult.Nint
End Property

Public Property Get CapacityStateMx() As Double
    If Not mStateResult Is Nothing Then CapacityStateMx = mStateResult.Mxint
End Property

Public Property Get CapacityStateMy() As Double
    If Not mStateResult Is Nothing Then CapacityStateMy = mStateResult.Myint
End Property

Public Property Get MomentUltimate() As Double
    MomentUltimate = mMomentUltimate
End Property

Public Property Get CapacityStateAvailable() As Boolean
    CapacityStateAvailable = Not mStateResult Is Nothing
End Property

Public Property Get CapacityEpsilon0() As Double
    If Not mStateResult Is Nothing Then CapacityEpsilon0 = mStateResult.Epsilon0
End Property

Public Property Get CapacityKappaX() As Double
    If Not mStateResult Is Nothing Then CapacityKappaX = mStateResult.KappaX
End Property

Public Property Get CapacityKappaY() As Double
    If Not mStateResult Is Nothing Then CapacityKappaY = mStateResult.KappaY
End Property

Public Property Get PureAxialForMethod() As Boolean
    PureAxialForMethod = mPureAxialForMethod
End Property

' Выдает найденный физический множитель как запас. Диагностическая lambda
' незавершенного поиска, BaseFail и отсутствующий предел запасом не являются.
Public Property Get ReserveFactor() As Double
    If mLambdaCapacity <= 0# Then Exit Property
    Dim policy As CResultStatusPolicy
    Set policy = New CResultStatusPolicy
    If Not policy.IsFinished(Status) Then Exit Property
    ReserveFactor = mLambdaCapacity
End Property
