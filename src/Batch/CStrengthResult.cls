VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CStrengthResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Сводный результат расчетов по модели прочности одного сочетания
' ==========================================================================
' CStrengthResult является промежуточным корневым узлом result-tree для direct
' state и capacity. Он не знает о профилях, Config и таблицах Results: сюда
' складываются уже полученные CDirectStateResult и CCapacityResult.
'
' На этапе 7 объект сосуществует со старым плоским фасадом CCombinationResult.
' Это позволяет постепенно переводить worst-combination, summary и writers на
' structured result graph без одномоментной большой переделки вывода.

Private mDirectState As CDirectStateResult ' Прямое НДС от заданного сочетания.
Private mCapacity As CCapacityResult ' Предельная несущая способность по выбранной lambda-траектории.

' Создает пустой direct-state, чтобы StrengthResult всегда имел нейтральную meta.
Private Sub Class_Initialize()
    Clear vbNullString
End Sub

' Очищает оба подрезультата; capacity остается Nothing, если расчет не запускался.
Public Sub Clear(ByVal defaultCapacityPath As String)
    Set mDirectState = New CDirectStateResult
    Set mCapacity = Nothing
End Sub

' Сохраняет статус прямого НДС без изменения его численных данных.
Public Sub SetDirectStatus(ByVal statusText As String, Optional ByVal resultComment As String = vbNullString)
    If mDirectState Is Nothing Then Set mDirectState = New CDirectStateResult
    mDirectState.SetStatus statusText, resultComment
End Sub

' Копирует прямое НДС из solver-а в структурированный direct-state result.
Public Sub StoreDirectFromSolver(ByVal solver As CSectionSolver, ByVal extensionEnabled As Boolean)
    If mDirectState Is Nothing Then Set mDirectState = New CDirectStateResult
    mDirectState.StoreSolver solver, extensionEnabled
End Sub

' Копирует прямое НДС из reusable named-state в структурированный direct-state result.
Public Sub StoreDirectFromStateResult(ByVal stateResult As CSectionStateResult, ByVal extensionEnabled As Boolean)
    If mDirectState Is Nothing Then Set mDirectState = New CDirectStateResult
    mDirectState.StoreStateResult stateResult, extensionEnabled
End Sub

' Сохраняет полный capacity-result как вторую ветвь strength-subtree.
Public Sub StoreCapacityResult(ByVal capacity As CCapacityResult)
    Set mCapacity = capacity
End Sub

' Возвращает direct-state branch. Writer-ы пока читают старый фасад, но тесты
' и последующая миграция могут проверять именно этот typed branch.
Public Property Get DirectState() As CDirectStateResult
    Set DirectState = mDirectState
End Property

' Возвращает capacity branch, если он был сформирован.
Public Property Get Capacity() As CCapacityResult
    Set Capacity = mCapacity
End Property

' Возвращает meta прямого НДС из structured result.
Public Property Get DirectMeta() As CResultMeta
    If mDirectState Is Nothing Then
        Dim meta As CResultMeta
        Set meta = New CResultMeta
        meta.SetNotApplicable rkDirectState
        Set DirectMeta = meta
    Else
        Set DirectMeta = mDirectState.Meta
    End If
End Property

' Возвращает meta capacity или N/A, если capacity не запускалась.
Public Property Get CapacityMeta() As CResultMeta
    If mCapacity Is Nothing Then
        Dim meta As CResultMeta
        Set meta = New CResultMeta
        meta.SetNotApplicable rkCapacity
        Set CapacityMeta = meta
    Else
        Set CapacityMeta = mCapacity.ResultMeta
    End If
End Property
