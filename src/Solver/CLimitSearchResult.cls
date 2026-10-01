VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLimitSearchResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Результат общего поиска предельной точки
' ==========================================================================
' CLimitSearchResult хранит типизированный итог search-слоя и численный снимок
' найденной limit-точки. Downstream-слои читают именно этот result, а не
' доменный solver-фасад.
'
' Result является снимком результата search: он не держит ссылок на доменные
' калькуляторы, чтобы последующие попытки Auto-пути не меняли уже сохраненную
' диагностику.

Private mMeta As CResultMeta ' Машинная оценка результата для будущего ResultStatusPolicy-пути.
Private mRequestedStrategy As String ' Стратегия, заданная настройкой.
Private mActualMethod As String ' Метод, который фактически дал результат поиска.
Private mResultKind As EResultKind ' Инженерный контекст результата.
Private mSucceeded As Boolean ' True, если search-слой получил корректный результат для своей задачи.

Private mHasLimitPoint As Boolean ' Есть фактически найденная точка, не просто корректный исход без точки.
Private mPointState As CSectionStateResult ' Подтвержденный неизменяемый State найденной точки, если определен.
Private mSearchExecuted As Boolean ' True, если численный search действительно запускался.
Private mConverged As Boolean ' Итоговый флаг сходимости физической limit-точки.
Private mStopReason As String ' Техническая причина остановки search.
Private mLimitState As String ' Инженерная причина предельного состояния или ошибки.
Private mLambdaUltimate As Double ' Найденный множитель lambda.
Private mNUltimate As Double ' Предельная N в собственном представлении search.
Private mMxUltimate As Double ' Предельный Mx в собственном представлении search.
Private mMyUltimate As Double ' Предельный My в собственном представлении search.
Private mMomentUltimate As Double ' Модуль предельного момента, если он сформирован search.
Private mUtilization As Double ' Использование несущей способности, сформированное search.
Private mLastSolver As CSectionSolver ' Последний solver равновесия для физически найденной точки.
Private mLastSolverDiagnosticLog As String ' Снимок диагностики последнего solver-а.
Private mDiagnosticLog As String ' Снимок диагностического журнала search-слоя.

Private Sub Class_Initialize()
    Set mMeta = New CResultMeta
End Sub

' Сбрасывает численный снимок limit-search перед заполнением результата.
Private Sub ResetSearchSnapshot()
    mHasLimitPoint = False
    Set mPointState = Nothing
    mSearchExecuted = False
    mConverged = False
    mStopReason = vbNullString
    mLimitState = vbNullString
    mLambdaUltimate = 0#
    mNUltimate = 0#
    mMxUltimate = 0#
    mMyUltimate = 0#
    mMomentUltimate = 0#
    mUtilization = 0#
    Set mLastSolver = Nothing
    mLastSolverDiagnosticLog = vbNullString
    mDiagnosticLog = vbNullString
End Sub

' Сохраняет capacity-числа сразу в общий result, чтобы downstream-слой не разбирал CCapacitySolver.
Private Sub CaptureCapacitySnapshot(ByVal cap As CCapacitySolver)
    ResetSearchSnapshot
    If cap Is Nothing Then Exit Sub

    mSearchExecuted = True
    mConverged = cap.Converged
    mHasLimitPoint = cap.Converged And cap.LambdaUltimate > 0#
    mStopReason = cap.StopReason
    mLimitState = cap.LimitState
    mLambdaUltimate = cap.LambdaUltimate
    mNUltimate = cap.NUltimate
    mMxUltimate = cap.MxUltimate
    mMyUltimate = cap.MyUltimate
    mMomentUltimate = cap.MomentUltimate
    mUtilization = cap.Utilization
    mDiagnosticLog = cap.DiagnosticLog
    If Not cap.LastSolver Is Nothing Then
        Set mLastSolver = cap.LastSolver
        mLastSolverDiagnosticLog = cap.LastSolver.DiagnosticLog
    End If
End Sub

' Сохраняет найденную crack-точку и диагностику, независимо от того,
' достигнута ли эта точка текущим сочетанием. Состояние до трещины, когда
' оно доступно, является источником полного внутреннего вектора усилий.
Private Sub CaptureCrackSnapshot(ByVal crack As CCrackFormationCalculator)
    ResetSearchSnapshot
    If crack Is Nothing Then Exit Sub

    mSearchExecuted = True
    mHasLimitPoint = crack.HasLimitPoint
    mConverged = crack.Converged And mHasLimitPoint
    mStopReason = crack.StopReason
    mLimitState = crack.StopReason
    mLambdaUltimate = crack.LambdaCrc
    If mHasLimitPoint Then crack.PointLoads mNUltimate, mMxUltimate, mMyUltimate
    mMomentUltimate = crack.Mcrc
    If Not crack.PreCrackState Is Nothing Then
        Set mPointState = crack.PreCrackState
        mNUltimate = crack.PreCrackState.Nint
        mMxUltimate = crack.PreCrackState.Mxint
        mMyUltimate = crack.PreCrackState.Myint
    End If
    mDiagnosticLog = crack.DiagnosticLog
End Sub

' Фиксирует численный исход поиска точки. Найденная точка успешна независимо
' от того, достигается ли она до lambda=1: безопасность заданного сочетания
' оценивает инженерный CapacityResult, а не общий Search.
Public Sub InitializeFromCapacitySolver(ByVal cap As CCapacitySolver, ByVal requestedStrategy As String)
    mRequestedStrategy = requestedStrategy
    mResultKind = rkCapacity
    mSucceeded = False
    CaptureCapacitySnapshot cap
    If cap Is Nothing Then
        mActualMethod = vbNullString
        mMeta.SetResult rsInternalError, rcInternalError, rkCapacity, _
            "Не передан результат выполненного поиска несущей способности."
    Else
        mActualMethod = cap.SolutionMethod
        If cap.Converged Then
            mSucceeded = True
            mMeta.SetResult rsSuccess, rcCheckPassed, rkCapacity, cap.LimitState
        ElseIf cap.SearchBoundReached Then
            mMeta.SetResult rsNumericalFailure, rcSearchBoundReached, rkCapacity, cap.StopReason
        ElseIf cap.LimitSearchInitialLambdaRejected And _
                Not cap.LimitSearchCapacityStateIsInvalidInput(cap.LimitState) Then
            mMeta.SetResult rsCheckFailed, rcInitialStateBeyondLimit, rkCapacity, _
                CapacityInitialLambdaRejectedComment(cap.StopReason)
        ElseIf cap.LimitSearchCapacityStateIsInvalidInput(cap.LimitState) Then
            mMeta.SetResult rsInvalidInput, rcInvalidInput, rkCapacity, cap.StopReason
        ElseIf cap.LimitSearchCapacityStateIsNumericalFailure(cap.LimitState) Then
            mMeta.SetResult rsNumericalFailure, rcNumericalFailure, rkCapacity, cap.StopReason
        Else
            mMeta.SetResult rsInternalError, rcInternalError, rkCapacity, cap.StopReason
        End If
    End If
End Sub

' Возвращает пользовательский комментарий без служебных префиксов search.
' Сам ResultCode хранит машинный смысл BaseFail.
Private Function CapacityInitialLambdaRejectedComment(ByVal stopReason As String) As String
    Dim commentText As String
    commentText = Trim$(stopReason)
    commentText = Replace$(commentText, "INITIAL_STATE_BEYOND_LIMIT:", vbNullString, 1, 1, vbTextCompare)
    commentText = Replace$(commentText, _
        "Начальное состояние при lambda = 0 не проходит физический критерий:", _
        "Несущая способность не проходит уже для исходной части выбранного пути:", _
        1, 1, vbTextCompare)
    CapacityInitialLambdaRejectedComment = Trim$(commentText)
    If Len(CapacityInitialLambdaRejectedComment) = 0 Then _
        CapacityInitialLambdaRejectedComment = "Несущая способность не проходит уже для исходной части выбранного пути."
End Function

' Заполняет result по уже выполненному CCrackFormationCalculator в стадии
' образования трещины. Формульная ширина здесь не оценивается: речь только о
' search-результате Ncrc/Mcrc и связанных состояниях.
Public Sub InitializeFromCrackFormation(ByVal crack As CCrackFormationCalculator, _
        ByVal requestedStrategy As String, ByVal searchSucceeded As Boolean)
    mRequestedStrategy = requestedStrategy
    mResultKind = rkCrackFormation
    mActualMethod = requestedStrategy
    mSucceeded = searchSucceeded
    CaptureCrackSnapshot crack

    If crack Is Nothing Then
        mMeta.SetResult rsInternalError, rcInternalError, rkCrackFormation, _
            "Не передан результат выполненного поиска образования трещины."
    ElseIf crack.CrackFormationInternalStatus <> rsUnset Then
        mActualMethod = crack.FormationSolutionMethod
        mMeta.SetResult crack.CrackFormationInternalStatus, crack.CrackFormationResultCode, _
            rkCrackFormation, crack.StopReason
    ElseIf searchSucceeded Then
        mMeta.SetResult rsSuccess, rcCheckPassed, rkCrackFormation, crack.StopReason
    Else
        mMeta.SetResult rsInternalError, rcInternalError, rkCrackFormation, _
            "Поиск образования трещины завершен без машинного статуса результата."
    End If
    mSucceeded = LimitSearchMetaSucceeded(mMeta)
End Sub

' Возвращает True для корректного search-результата, включая физическое
' недостижение критерия. Численная ошибка и ошибка входных данных остаются
' неуспешными результатами.
Private Function LimitSearchMetaSucceeded(ByVal meta As CResultMeta) As Boolean
    If meta Is Nothing Then Exit Function
    LimitSearchMetaSucceeded = (meta.InternalStatus = rsSuccess Or _
        meta.InternalStatus = rsSuccessWithWarning)
End Function

' Заполняет result как ошибку настройки, когда стратегия поиска неизвестна.
Public Sub InitializeInvalidConfiguration(ByVal requestedStrategy As String, ByVal message As String, _
        Optional ByVal resultKind As EResultKind = rkCapacity)
    mRequestedStrategy = requestedStrategy
    mActualMethod = vbNullString
    mResultKind = resultKind
    mSucceeded = False
    ResetSearchSnapshot
    mStopReason = message
    mLimitState = message
    mMeta.SetResult rsInvalidConfiguration, rcInvalidConfiguration, resultKind, message
End Sub

Public Property Get Meta() As CResultMeta
    Set Meta = mMeta.Clone
End Property

Public Property Get RequestedStrategy() As String
    RequestedStrategy = mRequestedStrategy
End Property

Public Property Get ActualMethod() As String
    ActualMethod = mActualMethod
End Property

Public Property Get ResultKind() As EResultKind
    ResultKind = mResultKind
End Property

Public Property Get Succeeded() As Boolean
    Succeeded = mSucceeded
End Property

' Простые свойства ниже отдают снимок search-result без обращения к вызывающему batch-слою.
Public Property Get SearchExecuted() As Boolean
    SearchExecuted = mSearchExecuted
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

Public Property Get LambdaUltimate() As Double
    LambdaUltimate = mLambdaUltimate
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

Public Property Get MomentUltimate() As Double
    MomentUltimate = mMomentUltimate
End Property

Public Property Get Utilization() As Double
    Utilization = mUtilization
End Property

Public Property Get LastSolver() As CSectionSolver
    Set LastSolver = mLastSolver
End Property

Public Property Get LastSolverDiagnosticLog() As String
    LastSolverDiagnosticLog = mLastSolverDiagnosticLog
End Property

' Возвращает подробную диагностику search как сохраненный снимок.
Public Property Get DiagnosticLog() As String
    DiagnosticLog = mDiagnosticLog
End Property

' Различает найденную точку и успешный исход без предельной точки.
Public Property Get HasLimitPoint() As Boolean
    HasLimitPoint = mHasLimitPoint
End Property

' Возвращает готовое НДС точки без обращения к рабочему калькулятору.
Public Property Get PointState() As CSectionStateResult
    Set PointState = mPointState
End Property
