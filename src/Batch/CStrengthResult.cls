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
' Общий результат проверок по модели прочности
' ==========================================================================
' Объединяет CDirectStateResult для заданной нагрузки и CCapacityResult
' для предельной несущей способности. Оба результата передаются уже готовыми.
' Статусы и пояснения этого блока служат источником данных для таблицы Results;
' видимый пользователю статус определяет CResultStatusPolicy.

Private mFrozen As Boolean ' После передачи готового результата его заполнение и очистка запрещены; сохранённые числа доступны для чтения.
Private mDirectState As CDirectStateResult ' Прямое НДС от заданного сочетания.
Private mCapacity As CCapacityResult ' Предельная несущая способность по выбранной lambda-траектории.
Private mPolicy As CResultStatusPolicy ' Общие правила выбора и вывода статуса; отдельный словарь обозначений здесь не хранится.

' Создаёт пустой результат НДС с признаком отсутствия выполненной проверки.
Private Sub Class_Initialize()
    Clear vbNullString
End Sub

' Очищает результаты НДС и несущей способности. Если поиск предела не запрошен,
' его результат отмечается как невыполненный без вымышленной предельной точки.
Public Sub Clear(ByVal defaultCapacityPath As String)
    AssertWritable
    Set mDirectState = New CDirectStateResult
    Set mCapacity = New CCapacityResult
    mCapacity.InitializeNotRequested "Несущая способность не запрошена.", defaultCapacityPath
End Sub

' Сохраняет статус и пояснение НДС, не меняя его численные данные.
Public Sub SetDirectMeta(ByVal meta As CResultMeta)
    AssertWritable
    If mDirectState Is Nothing Then Set mDirectState = New CDirectStateResult
    mDirectState.SetMeta meta
End Sub

' Связывает рассчитанное НДС и исходные физические пределы материалов с проверкой прочности.
Public Sub StoreDirectFromStateResult(ByVal stateResult As CSectionStateResult, ByVal extensionEnabled As Boolean, _
        Optional ByVal physicalConcrete As CMaterialDiagram = Nothing, _
        Optional ByVal physicalSteel As CMaterialDiagram = Nothing)
    AssertWritable
    If mDirectState Is Nothing Then Set mDirectState = New CDirectStateResult
    mDirectState.StoreStateResult stateResult, extensionEnabled, physicalConcrete, physicalSteel
End Sub

' Сохраняет полный результат несущей способности в составе проверок прочности.
Public Sub StoreCapacityResult(ByVal capacity As CCapacityResult)
    AssertWritable
    Set mCapacity = capacity
End Sub

' Возвращает результат НДС при заданной нагрузке с его числами, статусом и пояснением.
Public Property Get DirectState() As CDirectStateResult
    Set DirectState = mDirectState
End Property

' Возвращает результат несущей способности, если он был получен.
Public Property Get Capacity() As CCapacityResult
    Set Capacity = mCapacity
End Property

' Возвращает статус и пояснение расчёта НДС при заданной нагрузке.
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

' Возвращает статус несущей способности или N/A, если её расчёт не запускался.
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

' Собирает общий итог НДС и несущей способности. Класс записи на лист получает
' готовое пояснение; CCombinationResult не выполняет это объединение повторно.
Public Property Get ResultMeta() As CResultMeta
    If mPolicy Is Nothing Then Set mPolicy = New CResultStatusPolicy
    Dim direct As CResultMeta, capacity As CResultMeta, comment As String
    Set direct = DirectMeta
    Set capacity = CapacityMeta
    If direct.InternalStatus = rsInvalidInput And capacity.InternalStatus = rsInvalidInput And _
            StrComp(direct.ResultComment, capacity.ResultComment, vbBinaryCompare) = 0 Then
        ' При общей ошибке ввода ни НДС, ни несущая способность не рассчитывались.
        ' Общее пояснение относится ко всему блоку, а статусы отдельных проверок сохраняются.
        comment = "Прочность: " & direct.ResultComment
    Else
        direct.AppendCommentTo comment
        capacity.AppendCommentTo comment
    End If
    Set ResultMeta = mPolicy.WorstResultMeta(rkGeneric, direct, capacity)
    ResultMeta.ResultComment = comment
End Property

' Объединяет причины завершения последовательных решений НДС в одном запуске.
' Сохраняет пояснение выбранной причины, не заменяя его пустым текстом.
Public Sub MergeDirectMeta(ByVal meta As CResultMeta)
    AssertWritable
    If meta Is Nothing Then Exit Sub
    If mPolicy Is Nothing Then Set mPolicy = New CResultStatusPolicy
    SetDirectMeta mPolicy.WorstResultMeta(rkDirectState, DirectMeta, meta)
End Sub

' Завершает заполнение общего результата прочности, НДС и несущей способности.
' После этого они доступны только для чтения; повторный вызов безопасен,
' численные данные остаются в соответствующих объектах результатов.
Friend Sub Freeze()
    If mFrozen Then Exit Sub
    If Not mDirectState Is Nothing Then mDirectState.Freeze
    If Not mCapacity Is Nothing Then mCapacity.Freeze
    mFrozen = True
End Sub

' Запрещает изменение результата после передачи потребителям.
' Для повторного расчёта создаётся новый объект, а прежний результат сохраняется.
Private Sub AssertWritable()
    If mFrozen Then Err.Raise vbObjectError + 4220, "CStrengthResult", _
        "Опубликованный результат нельзя изменить; для нового расчета нужен новый объект."
End Sub
