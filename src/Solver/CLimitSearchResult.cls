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

' ===========================================================================
' Самодостаточный результат общего поиска предельной точки
' ===========================================================================
' Хранит численный исход, фактический метод, диагностику и подтвержденное НДС
' найденной точки. Не обращается к рабочему solver-у или инженерному calculator-у.
' Смысл физического критерия и проверка текущего LC остаются у доменной задачи.
' Для низкоуровневых материалов без material spec сохраняется численный вектор;
' production named-state передается готовым снимком.

Private mMeta As CResultMeta ' Машинный исход, подготовленный владельцем задачи.
Private mRequestedStrategy As String ' Запрошенная численная стратегия.
Private mActualMethod As String ' Фактически выполненный численный метод.
Private mHasLimitPoint As Boolean ' Найдена точка, а не только завершена процедура.
Private mPointState As CSectionStateResult ' Единственный снимок НДС найденной точки.
Private mSearchExecuted As Boolean ' Численный поиск действительно запускался.
Private mLambdaUltimate As Double ' Множитель нагрузки в найденной точке.
Private mNUltimate As Double ' Вектор для материала без named-state спецификации.
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

' Принимает готовый исход задачи и численный снимок без живых калькуляторов.
' Отсутствие найденной точки очищает ее числа и State при повторном заполнении.
Public Sub Initialize(ByVal meta As CResultMeta, ByVal requestedStrategy As String, _
        ByVal actualMethod As String, ByVal hasLimitPoint As Boolean, _
        ByVal lambdaValue As Double, ByVal nValue As Double, ByVal mxValue As Double, _
        ByVal myValue As Double, ByVal pointState As CSectionStateResult, _
        ByVal limitState As String, ByVal diagnosticLog As String, _
        Optional ByVal finalSolverDiagnosticLog As String = vbNullString, _
        Optional ByVal searchExecuted As Boolean = True)
    ResetPoint
    mRequestedStrategy = requestedStrategy
    mActualMethod = actualMethod
    mSearchExecuted = searchExecuted
    mLimitState = limitState
    mDiagnosticLog = diagnosticLog
    mLastSolverDiagnosticLog = finalSolverDiagnosticLog
    If meta Is Nothing Then
        mMeta.SetResult rsInternalError, rcInternalError, rkGeneric, "Поиск завершен без машинного результата.", _
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

' Создает независимый контейнер Search: его повторная Initialize не меняет
' опубликованный engineering result. Конечный named-state остается одним
' неизменяемым снимком, не копируется по волокнам и не превращается в probe.
Public Function Clone() As CLimitSearchResult
    Dim result As CLimitSearchResult
    Set result = New CLimitSearchResult
    result.Initialize mMeta, mRequestedStrategy, mActualMethod, mHasLimitPoint, _
        mLambdaUltimate, mNUltimate, mMxUltimate, mMyUltimate, mPointState, _
        mLimitState, mDiagnosticLog, mLastSolverDiagnosticLog, mSearchExecuted
    Set Clone = result
End Function

' Метаданные выдаются отдельной небольшой копией, а не mutable полем результата.
Public Property Get Meta() As CResultMeta
    Set Meta = mMeta.Clone
End Property

' Простые getters ниже читают только сохраненный снимок.
Public Property Get RequestedStrategy() As String
    RequestedStrategy = mRequestedStrategy
End Property

Public Property Get ActualMethod() As String
    ActualMethod = mActualMethod
End Property

Public Property Get ResultKind() As EResultKind
    ResultKind = mMeta.ResultKind
End Property

Public Property Get Succeeded() As Boolean
    Succeeded = (mMeta.InternalStatus = rsSuccess Or mMeta.InternalStatus = rsSuccessWithWarning)
End Property

Public Property Get SearchExecuted() As Boolean
    SearchExecuted = mSearchExecuted
End Property

Public Property Get Converged() As Boolean
    Converged = mHasLimitPoint
End Property

Public Property Get StopReason() As String
    StopReason = mMeta.ResultComment
End Property

Public Property Get LimitState() As String
    LimitState = mLimitState
End Property

Public Property Get LambdaUltimate() As Double
    LambdaUltimate = mLambdaUltimate
End Property

Public Property Get NUltimate() As Double
    If mPointState Is Nothing Then
        NUltimate = mNUltimate
    Else
        NUltimate = mPointState.Nint
    End If
End Property

Public Property Get MxUltimate() As Double
    If mPointState Is Nothing Then
        MxUltimate = mMxUltimate
    Else
        MxUltimate = mPointState.Mxint
    End If
End Property

Public Property Get MyUltimate() As Double
    If mPointState Is Nothing Then
        MyUltimate = mMyUltimate
    Else
        MyUltimate = mPointState.Myint
    End If
End Property

Public Property Get LastSolverDiagnosticLog() As String
    LastSolverDiagnosticLog = mLastSolverDiagnosticLog
End Property

Public Property Get DiagnosticLog() As String
    DiagnosticLog = mDiagnosticLog
End Property

Public Property Get HasLimitPoint() As Boolean
    HasLimitPoint = mHasLimitPoint
End Property

Public Property Get PointState() As CSectionStateResult
    Set PointState = mPointState
End Property
