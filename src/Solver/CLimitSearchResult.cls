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
Private mMeta As CResultMeta ' Машинная оценка результата для будущего ResultStatusPolicy-пути.
Private mRequestedStrategy As String ' Стратегия, заданная настройкой.
Private mActualMethod As String ' Метод, который фактически записал CCapacitySolver.
Private mResultKind As EResultKind ' Инженерный контекст результата.

Private Sub Class_Initialize()
    Set mMeta = New CResultMeta
End Sub

' Заполняет result по уже выполненному CCapacitySolver.
' Маппинг здесь намеренно простой: внешний статус по-прежнему формируется
' существующим CBatchStatusPolicy, а meta нужна как задел общей result-схемы.
Public Sub InitializeFromCapacitySolver(ByVal cap As CCapacitySolver, ByVal requestedStrategy As String)
    Set mCapacitySolver = cap
    mRequestedStrategy = requestedStrategy
    mResultKind = rkCapacity
    If cap Is Nothing Then
        mActualMethod = vbNullString
        mMeta.SetResult rsNumericalFailure, rcSolverDidNotReturnState, rkCapacity, _
            "Поиск несущей способности не вернул объект solver."
    Else
        mActualMethod = cap.SolutionMethod
        If cap.Converged Then
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

' Заполняет result как ошибку настройки, когда стратегия поиска неизвестна.
Public Sub InitializeInvalidConfiguration(ByVal requestedStrategy As String, ByVal message As String)
    Set mCapacitySolver = Nothing
    mRequestedStrategy = requestedStrategy
    mActualMethod = vbNullString
    mResultKind = rkCapacity
    mMeta.SetResult rsInvalidConfiguration, rcInvalidConfiguration, rkCapacity, message
End Sub

Public Property Get CapacitySolver() As CCapacitySolver
    Set CapacitySolver = mCapacitySolver
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
