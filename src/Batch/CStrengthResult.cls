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

Private mFrozen As Boolean ' После публикации заполнение/reset запрещены; численные данные не копируются.
Private mDirectState As CDirectStateResult ' Прямое НДС от заданного сочетания.
Private mCapacity As CCapacityResult ' Предельная несущая способность по выбранной lambda-траектории.
Private mPolicy As CResultStatusPolicy ' Общий typed/display приоритет, без собственной копии словаря.

' Создает пустой direct-state, чтобы StrengthResult всегда имел нейтральную meta.
Private Sub Class_Initialize()
    Clear vbNullString
End Sub

' Очищает оба подрезультата. Незапрошенная capacity имеет нейтральную typed meta,
' но не содержит фиктивной найденной точки или физического состояния.
Public Sub Clear(ByVal defaultCapacityPath As String)
    AssertWritable
    Set mDirectState = New CDirectStateResult
    Set mCapacity = New CCapacityResult
    mCapacity.InitializeSkipped "Несущая способность не запрошена.", defaultCapacityPath
End Sub

' Сохраняет typed meta прямого НДС без изменения его численных данных.
Public Sub SetDirectMeta(ByVal meta As CResultMeta)
    AssertWritable
    If mDirectState Is Nothing Then Set mDirectState = New CDirectStateResult
    mDirectState.SetMeta meta
End Sub

' Привязывает прямое НДС и исходные физические пределы к strength-результату.
Public Sub StoreDirectFromStateResult(ByVal stateResult As CSectionStateResult, ByVal extensionEnabled As Boolean, _
        Optional ByVal physicalConcrete As CMaterialDiagram = Nothing, _
        Optional ByVal physicalSteel As CMaterialDiagram = Nothing)
    AssertWritable
    If mDirectState Is Nothing Then Set mDirectState = New CDirectStateResult
    mDirectState.StoreStateResult stateResult, extensionEnabled, physicalConcrete, physicalSteel
End Sub

' Сохраняет полный capacity-result как вторую ветвь strength-subtree.
Public Sub StoreCapacityResult(ByVal capacity As CCapacityResult)
    AssertWritable
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

' Собирает итог только собственного поддерева DirectState/Capacity. Writer
' получает готовую причину, а Combination не повторяет strength-агрегацию.
Public Property Get ResultMeta() As CResultMeta
    If mPolicy Is Nothing Then Set mPolicy = New CResultStatusPolicy
    Dim direct As CResultMeta, capacity As CResultMeta, comment As String
    Set direct = DirectMeta
    Set capacity = CapacityMeta
    direct.AppendCommentTo comment
    capacity.AppendCommentTo comment
    Set ResultMeta = mPolicy.WorstResultMeta(rkGeneric, direct, capacity)
    ResultMeta.ResultComment = comment
End Property

' Объединяет причины последовательных прямых State одного workflow, сохраняя
' комментарий выбранной причины и не стирая его пустым аргументом агрегации.
Public Sub MergeDirectMeta(ByVal meta As CResultMeta)
    AssertWritable
    If meta Is Nothing Then Exit Sub
    If mPolicy Is Nothing Then Set mPolicy = New CResultStatusPolicy
    SetDirectMeta mPolicy.WorstResultMeta(rkDirectState, DirectMeta, meta)
End Sub

' Завершает жизненный цикл заполнения перед передачей snapshot потребителю.
' Повторный вызов безопасен; общие ссылки на named-state остаются теми же,
' но их fill/reset API больше не могут изменить опубликованный результат.
Friend Sub Freeze()
    If mFrozen Then Exit Sub
    If Not mDirectState Is Nothing Then mDirectState.Freeze
    If Not mCapacity Is Nothing Then mCapacity.Freeze
    mFrozen = True
End Sub

' Явно отклоняет изменение уже выданного результата вместо тихой порчи снимка.
' Для повторного расчета workflow создает новый объект, сохраняя прежние ссылки.
Private Sub AssertWritable()
    If mFrozen Then Err.Raise vbObjectError + 4220, "CStrengthResult", _
        "Опубликованный результат нельзя изменить; для нового расчета нужен новый объект."
End Sub
