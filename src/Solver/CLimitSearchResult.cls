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
' CLimitSearchResult хранит типизированный итог search-слоя и ссылку на существующий
' CCapacitySolver. Ссылка нужна на переходном этапе: существующий вывод,
' snapshot и тесты уже читают много полей именно из CCapacitySolver.
'
' По мере дальнейшей миграции численные поля и ResultMeta можно будет читать
' напрямую из этого класса, не меняя публичную расчетную методику.

Private mCapacitySolver As CCapacitySolver ' Фактический solver, выполнивший расчет capacity.
Private mCrackCalculator As CCrackWidthCalculator ' Расчетчик, выполнивший поиск точки трещинообразования.
Private mMeta As CResultMeta ' Машинная оценка результата для будущего ResultStatusPolicy-пути.
Private mRequestedStrategy As String ' Стратегия, заданная настройкой.
Private mActualMethod As String ' Метод, который фактически записал CCapacitySolver.
Private mResultKind As EResultKind ' Инженерный контекст результата.
Private mSucceeded As Boolean ' True, если search-слой получил корректный результат для своей задачи.

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

Private Sub Class_Initialize()
    Set mMeta = New CResultMeta
End Sub

' Сбрасывает численный снимок limit-search перед заполнением результата.
Private Sub ResetSearchSnapshot()
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
End Sub

' Сохраняет capacity-числа сразу в общий result, чтобы downstream-слой не разбирал CCapacitySolver.
Private Sub CaptureCapacitySnapshot(ByVal cap As CCapacitySolver)
    ResetSearchSnapshot
    If cap Is Nothing Then Exit Sub

    mSearchExecuted = True
    mConverged = cap.Converged
    mStopReason = cap.StopReason
    mLimitState = cap.LimitState
    mLambdaUltimate = cap.LambdaUltimate
    mNUltimate = cap.NUltimate
    mMxUltimate = cap.MxUltimate
    mMyUltimate = cap.MyUltimate
    mMomentUltimate = cap.MomentUltimate
    mUtilization = cap.Utilization
    If Not cap.LastSolver Is Nothing Then
        Set mLastSolver = cap.LastSolver
        mLastSolverDiagnosticLog = cap.LastSolver.DiagnosticLog
    End If
End Sub

' Сохраняет краткую диагностику CrackFormation в той же result-структуре.
Private Sub CaptureCrackSnapshot(ByVal crack As CCrackWidthCalculator)
    ResetSearchSnapshot
    If crack Is Nothing Then Exit Sub

    mSearchExecuted = True
    mConverged = crack.Converged
    mStopReason = crack.StopReason
    mLimitState = crack.StopReason
End Sub

' Заполняет result по уже выполненному CCapacitySolver.
' Маппинг здесь намеренно простой: внешний статус по-прежнему формируется
' существующим CBatchStatusPolicy, а meta нужна как задел общей result-схемы.
Public Sub InitializeFromCapacitySolver(ByVal cap As CCapacitySolver, ByVal requestedStrategy As String)
    Set mCapacitySolver = cap
    Set mCrackCalculator = Nothing
    mRequestedStrategy = requestedStrategy
    mResultKind = rkCapacity
    mSucceeded = False
    CaptureCapacitySnapshot cap
    If cap Is Nothing Then
        mActualMethod = vbNullString
        mMeta.SetResult rsNumericalFailure, rcSolverDidNotReturnState, rkCapacity, _
            "Поиск несущей способности не вернул объект solver."
    Else
        mActualMethod = cap.SolutionMethod
        If cap.Converged Then
            mSucceeded = True
            If cap.LambdaUltimate > 0# And cap.LambdaUltimate < 1# Then
                mMeta.SetResult rsCheckFailed, rcPhysicalLimitExceeded, rkCapacity, cap.LimitState
            Else
                mMeta.SetResult rsSuccess, rcCheckPassed, rkCapacity, cap.LimitState
            End If
        ElseIf InStr(1, cap.StopReason, "INITIAL_STATE_BEYOND_LIMIT", vbTextCompare) > 0 Then
            mMeta.SetResult rsCheckFailed, rcInitialStateBeyondLimit, rkCapacity, cap.StopReason
        ElseIf InStr(1, cap.StopReason, "Input", vbTextCompare) > 0 Or _
                InStr(1, cap.StopReason, "ошиб", vbTextCompare) > 0 Then
            mMeta.SetResult rsInvalidInput, rcInvalidInput, rkCapacity, cap.StopReason
        Else
            mMeta.SetResult rsNumericalFailure, rcNumericalFailure, rkCapacity, cap.StopReason
        End If
    End If
End Sub

' Заполняет result по уже выполненному CCrackWidthCalculator в стадии
' образования трещины. Формульная ширина здесь не оценивается: речь только о
' search-результате Ncrc/Mcrc и связанных состояниях.
Public Sub InitializeFromCrackFormation(ByVal crack As CCrackWidthCalculator, _
        ByVal requestedStrategy As String, ByVal searchSucceeded As Boolean)
    Set mCapacitySolver = Nothing
    Set mCrackCalculator = crack
    mRequestedStrategy = requestedStrategy
    mResultKind = rkCrackFormation
    mActualMethod = requestedStrategy
    mSucceeded = searchSucceeded
    CaptureCrackSnapshot crack

    If crack Is Nothing Then
        mMeta.SetResult rsNumericalFailure, rcSolverDidNotReturnState, rkCrackFormation, _
            "Поиск образования трещины не вернул расчетчик."
    ElseIf searchSucceeded Then
        If crack.CrackFormed Then
            mMeta.SetResult rsSuccess, rcCheckPassed, rkCrackFormation, crack.StopReason
        Else
            mMeta.SetResult rsSuccess, rcCriterionNotReached, rkCrackFormation, crack.StopReason
        End If
    ElseIf crack.Converged And Not crack.CrackFormed Then
        mMeta.SetResult rsSuccess, rcCriterionNotReached, rkCrackFormation, crack.StopReason
    ElseIf InStr(1, crack.StopReason, "Input", vbTextCompare) > 0 Or _
            InStr(1, crack.StopReason, "ошиб", vbTextCompare) > 0 Then
        mMeta.SetResult rsInvalidInput, rcInvalidInput, rkCrackFormation, crack.StopReason
    ElseIf InStr(1, crack.StopReason, "Config", vbTextCompare) > 0 Or _
            InStr(1, crack.StopReason, "настрой", vbTextCompare) > 0 Then
        mMeta.SetResult rsInvalidConfiguration, rcInvalidConfiguration, rkCrackFormation, crack.StopReason
    Else
        mMeta.SetResult rsNumericalFailure, rcNumericalFailure, rkCrackFormation, crack.StopReason
    End If
End Sub

' Заполняет result как ошибку настройки, когда стратегия поиска неизвестна.
Public Sub InitializeInvalidConfiguration(ByVal requestedStrategy As String, ByVal message As String, _
        Optional ByVal resultKind As EResultKind = rkCapacity)
    Set mCapacitySolver = Nothing
    Set mCrackCalculator = Nothing
    mRequestedStrategy = requestedStrategy
    mActualMethod = vbNullString
    mResultKind = resultKind
    mSucceeded = False
    ResetSearchSnapshot
    mStopReason = message
    mLimitState = message
    mMeta.SetResult rsInvalidConfiguration, rcInvalidConfiguration, resultKind, message
End Sub

Public Property Get CapacitySolver() As CCapacitySolver
    Set CapacitySolver = mCapacitySolver
End Property

Public Property Get CrackCalculator() As CCrackWidthCalculator
    Set CrackCalculator = mCrackCalculator
End Property

Public Property Get Meta() As CResultMeta
    Set Meta = mMeta
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

' Возвращает подробную диагностику фактического search-объекта.
Public Property Get DiagnosticLog() As String
    If Not mCapacitySolver Is Nothing Then
        DiagnosticLog = mCapacitySolver.DiagnosticLog
    ElseIf Not mCrackCalculator Is Nothing Then
        DiagnosticLog = mCrackCalculator.DiagnosticLog
    Else
        DiagnosticLog = vbNullString
    End If
End Property
