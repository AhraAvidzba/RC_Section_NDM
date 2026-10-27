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
' Сохранённый результат общего поиска предельной нагрузки
' ==========================================================================
' Хранит исход поиска, фактический метод, диагностику и подтверждённое НДС
' найденной точки. Для материалов без CMaterialModelSpec сохраняет численные
' усилия; готовое именованное состояние передаётся отдельным снимком.
' Физический смысл критерия и проверку текущего сочетания определяет инженерная задача.

Private mMeta As CResultMeta ' Машинный исход, подготовленный владельцем задачи.
Private mRequestedStrategy As String ' Запрошенная численная стратегия.
Private mActualMethod As String ' Способ определения точки: численный метод либо аналитический Ncrc.
Private mHasLimitPoint As Boolean ' Найдена точка, а не только завершена процедура.
Private mPointState As CSectionStateResult ' Единственный снимок НДС найденной точки.
Private mSearchExecuted As Boolean ' Численный поиск действительно запускался.
Private mLambdaUltimate As Double ' Множитель нагрузки в найденной точке.
Private mNUltimate As Double ' Численные усилия для задачи, в которой не создаётся отдельное именованное НДС.
Private mMxUltimate As Double ' Mx низкоуровневой точки, Н*мм.
Private mMyUltimate As Double ' My низкоуровневой точки, Н*мм.
Private mLimitState As String ' Код управляющего критерия, без интерпретации.
Private mDiagnosticLog As String ' Независимый текстовый снимок всего поиска.
Private mLastSolverDiagnosticLog As String ' Журнал финального равновесия.

' Создает безопасный исход до получения результата задачи.
Private Sub Class_Initialize()
    Set mMeta = New CResultMeta
    mMeta.SetResult rsInternalError, rcInternalError, rkGeneric, "Результат поиска еще не заполнен.", _
        vbNullString, True, False
End Sub

' Сохраняет готовый исход поиска и найденные числа независимо от рабочих расчётных объектов.
' Если точки нет, очищает её числа и НДС. Факт численного поиска передаётся
' отдельно: аналитическая Ncrc или ранняя проверка ввода не означают запуск поиска.
Public Sub Initialize(ByVal meta As CResultMeta, ByVal requestedStrategy As String, _
        ByVal actualMethod As String, ByVal hasLimitPoint As Boolean, _
        ByVal lambdaValue As Double, ByVal nValue As Double, ByVal mxValue As Double, _
        ByVal myValue As Double, ByVal pointState As CSectionStateResult, _
        ByVal limitState As String, ByVal diagnosticLog As String, _
        Optional ByVal finalSolverDiagnosticLog As String = vbNullString, _
        Optional ByVal searchExecuted As Boolean = False)
    ResetPoint
    mRequestedStrategy = requestedStrategy
    mActualMethod = actualMethod
    mSearchExecuted = searchExecuted
    mLimitState = limitState
    mDiagnosticLog = diagnosticLog
    mLastSolverDiagnosticLog = finalSolverDiagnosticLog
    If meta Is Nothing Then
        mMeta.SetResult rsInternalError, rcInternalError, rkGeneric, "Поиск завершён без обязательного статуса и кода причины результата. Это внутренняя ошибка передачи результатов.", _
            vbNullString, True, searchExecuted
        Exit Sub
    End If
    Set mMeta = meta.Clone
    If Not hasLimitPoint Then Exit Sub
    If Not pointState Is Nothing Then
        If Not pointState.Converged Then
            mMeta.SetResult rsInternalError, rcInternalError, meta.ResultKind, _
                "Предельная точка передана без подтвержденного равновесия.", vbNullString, True, searchExecuted
            Exit Sub
        End If
        pointState.Freeze
        Set mPointState = pointState
    Else
        mNUltimate = nValue
        mMxUltimate = mxValue
        mMyUltimate = myValue
    End If
    mLambdaUltimate = lambdaValue
    mHasLimitPoint = True
End Sub

' Очищает только данные точки; метаданные и журнал заполняет Initialize.
Private Sub ResetPoint()
    mHasLimitPoint = False
    Set mPointState = Nothing
    mLambdaUltimate = 0#
    mNUltimate = 0#
    mMxUltimate = 0#
    mMyUltimate = 0#
End Sub

' Формирует ошибку стратегии до выполнения численного поиска.
Public Sub InitializeInvalidConfiguration(ByVal requestedStrategy As String, ByVal message As String, _
        Optional ByVal resultKind As EResultKind = rkCapacity)
    Dim meta As CResultMeta
    Set meta = New CResultMeta
    meta.SetResult rsInvalidConfiguration, rcInvalidConfiguration, resultKind, message
    Initialize meta, requestedStrategy, vbNullString, False, 0#, 0#, 0#, 0#, _
        Nothing, vbNullString, vbNullString, vbNullString, False
End Sub

' Создаёт независимую копию результата поиска для последующего хранения.
' Её повторное заполнение не меняет уже выданный инженерный итог.
' Конечное НДС остаётся тем же защищённым от изменений состоянием без копирования сетки.
Public Function Clone() As CLimitSearchResult
    Dim result As CLimitSearchResult
    Set result = New CLimitSearchResult
    result.Initialize mMeta, mRequestedStrategy, mActualMethod, mHasLimitPoint, _
        mLambdaUltimate, mNUltimate, mMxUltimate, mMyUltimate, mPointState, _
        mLimitState, mDiagnosticLog, mLastSolverDiagnosticLog, mSearchExecuted
    Set Clone = result
End Function

' Сохраняет причины выполненных неудачных путей и их журналы в общем результате.
' При успехе после смены пути добавляет предупреждение. Фактическая причина
' последней попытки сохраняется; текст не превращает отсутствие точки в успех.
' Ранее накопленные сообщения не повторяются. Статус определяет факт предыдущих попыток, а не текст.
Public Sub AddLoadPathHistory(ByVal history As String, ByVal pathDisplayName As String, _
        ByVal hadPreviousAttempts As Boolean, Optional ByVal diagnosticHistory As String = vbNullString)
    If Len(diagnosticHistory) > 0 Then
        If StrComp(Left$(mDiagnosticLog, Len(diagnosticHistory)), diagnosticHistory, vbBinaryCompare) <> 0 Then _
            mDiagnosticLog = diagnosticHistory & vbCrLf & mDiagnosticLog
    End If
    Dim status As EResultInternalStatus, sourceKeys As String, comment As String
    status = mMeta.InternalStatus
    sourceKeys = mMeta.SourceKeys
    If hadPreviousAttempts And status = rsSuccess Then status = rsSuccessWithWarning
    comment = mMeta.ResultComment
    If Len(Trim$(history)) > 0 Then comment = Trim$(history) & " Итоговый путь " & pathDisplayName & ": " & comment
    mMeta.SetResult status, mMeta.ResultCode, mMeta.ResultKind, comment, _
        mMeta.DiagnosticDetails, mMeta.Applies, mMeta.Calculated
    mMeta.SourceKeys = sourceKeys
End Sub

' Возвращает отдельную копию статуса, причины и пояснения вместо изменяемого внутреннего
' объекта.
Public Property Get Meta() As CResultMeta
    Set Meta = mMeta.Clone
End Property

' Возвращает стратегию поиска, указанную в исходном запросе.
Public Property Get RequestedStrategy() As String
    RequestedStrategy = mRequestedStrategy
End Property

' Способ определения точки: численный метод либо аналитический Ncrc.
Public Property Get ActualMethod() As String
    ActualMethod = mActualMethod
End Property

' Возвращает вид инженерной проверки, для которой выполнен поиск.
Public Property Get ResultKind() As EResultKind
    ResultKind = mMeta.ResultKind
End Property

' Показывает, признан ли сохранённый исход успешным по статусу результата.
Public Property Get Succeeded() As Boolean
    Succeeded = (mMeta.InternalStatus = rsSuccess Or mMeta.InternalStatus = rsSuccessWithWarning)
End Property

' Численный поиск действительно запускался.
Public Property Get SearchExecuted() As Boolean
    SearchExecuted = mSearchExecuted
End Property

' Найдена точка, а не только завершена процедура.
Public Property Get Converged() As Boolean
    Converged = mHasLimitPoint
End Property

' Возвращает сохранённое пояснение причины завершения поиска.
Public Property Get StopReason() As String
    StopReason = mMeta.ResultComment
End Property

' Код управляющего критерия, без интерпретации.
Public Property Get LimitState() As String
    LimitState = mLimitState
End Property

' Множитель нагрузки в найденной точке.
Public Property Get LambdaUltimate() As Double
    LambdaUltimate = mLambdaUltimate
End Property

' Возвращает продольную силу из конечного НДС, если оно есть, иначе
' из численных усилий результата поиска. Само чтение не подтверждает наличие физического
' предела.
Public Property Get NUltimate() As Double
    If mPointState Is Nothing Then
        NUltimate = mNUltimate
    Else
        NUltimate = mPointState.Nint
    End If
End Property

' Возвращает Mx из сохранённого НДС, если оно есть, иначе
' из численных усилий результата. Наличие предельной точки проверяется отдельно.
Public Property Get MxUltimate() As Double
    If mPointState Is Nothing Then
        MxUltimate = mMxUltimate
    Else
        MxUltimate = mPointState.Mxint
    End If
End Property

' Возвращает My по тому же источнику, что N/Mx, во внутренних Н*мм;
' сам accessor не назначает успех поиска и не создает новое состояние.
Public Property Get MyUltimate() As Double
    If mPointState Is Nothing Then
        MyUltimate = mMyUltimate
    Else
        MyUltimate = mPointState.Myint
    End If
End Property

' Журнал финального равновесия.
Public Property Get LastSolverDiagnosticLog() As String
    LastSolverDiagnosticLog = mLastSolverDiagnosticLog
End Property

' Независимый текстовый снимок всего поиска.
Public Property Get DiagnosticLog() As String
    DiagnosticLog = mDiagnosticLog
End Property

' Найдена точка, а не только завершена процедура.
Public Property Get HasLimitPoint() As Boolean
    HasLimitPoint = mHasLimitPoint
End Property

' Единственный снимок НДС найденной точки.
Public Property Get PointState() As CSectionStateResult
    Set PointState = mPointState
End Property
