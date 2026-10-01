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
' Объект является источником typed meta для вывода strength-блока: writer-ы
' получают внешний статус через CResultStatusPolicy, а не выводят его из текста ошибок.

Private mDirectState As CDirectStateResult ' Прямое НДС от заданного сочетания.
Private mCapacity As CCapacityResult ' Предельная несущая способность по выбранной lambda-траектории.

' Создает пустой direct-state, чтобы StrengthResult всегда имел нейтральную meta.
Private Sub Class_Initialize()
    Clear vbNullString
End Sub

' Очищает оба подрезультата. Незапрошенная capacity имеет нейтральную typed meta,
' но не содержит фиктивной найденной точки или физического состояния.
Public Sub Clear(ByVal defaultCapacityPath As String)
    Set mDirectState = New CDirectStateResult
    Set mCapacity = New CCapacityResult
    mCapacity.InitializeSkipped "Несущая способность не запрошена.", defaultCapacityPath
End Sub

' Сохраняет typed meta прямого НДС без изменения его численных данных.
Public Sub SetDirectMeta(ByVal meta As CResultMeta)
    If mDirectState Is Nothing Then Set mDirectState = New CDirectStateResult
    mDirectState.SetMeta meta
End Sub

' Привязывает прямое НДС и исходные физические пределы к strength-результату.
Public Sub StoreDirectFromStateResult(ByVal stateResult As CSectionStateResult, ByVal extensionEnabled As Boolean, _
        Optional ByVal physicalConcrete As CMaterialDiagram = Nothing, _
        Optional ByVal physicalSteel As CMaterialDiagram = Nothing)
    If mDirectState Is Nothing Then Set mDirectState = New CDirectStateResult
    mDirectState.StoreStateResult stateResult, extensionEnabled, physicalConcrete, physicalSteel
End Sub

' Сохраняет полный capacity-result как вторую ветвь strength-subtree.
Public Sub StoreCapacityResult(ByVal capacity As CCapacityResult)
    Set mCapacity = capacity
End Sub

' Возвращает direct-state branch с численными данными и typed ResultMeta.
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
