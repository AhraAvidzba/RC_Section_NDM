VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CStateRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Хранилище расчётных состояний одного сочетания нагрузок
' ==========================================================================
' Сохраняет конечные состояния для Results и отдельно учитывает состояния,
' которые можно использовать при точном совпадении расчётного запроса.
' Неудачную попытку можно записать для диагностики, но FindEquivalent её
' не возвращает: повторное решение с другим начальным приближением остаётся возможным.

Private mFrozen As Boolean ' После передачи готового результата его заполнение и очистка запрещены; сохранённые числа доступны для чтения.
Private mStates As Collection ' Конечные расчётные состояния для вывода на Results.
Private mReusableByKey As Object ' Соответствие точного запроса ранее найденному равновесному НДС.
Private mContextSection As CSectionModel ' Объект геометрии, для которого сохранены расчётные состояния.
Private mContextMaterials As CMaterialModelProvider ' Фактически использованный набор материалов.
Private mSectionRevision As Long ' Версия геометрии на момент привязки контекста.
Private mMaterialRevision As Long ' Версия материалов на момент привязки контекста.
Private mAcceptanceKey As String ' Режим продолжения диаграмм и допуски равновесия.
Private mContextBound As Boolean ' Признак того, что общие условия повторного использования состояний уже заданы.

' Создаёт отдельные списки состояний для вывода и повторного использования.
' До BindContext нельзя подтвердить, что сохранённое вручную НДС относится к текущему запросу.
Private Sub Class_Initialize()
    Set mStates = New Collection
    Set mReusableByKey = CreateObject("Scripting.Dictionary")
End Sub

' Очищает состояния для вывода и повторного использования перед новым сочетанием нагрузок.
Public Sub Clear()
    AssertWritable
    Set mStates = New Collection
    Set mReusableByKey = CreateObject("Scripting.Dictionary")
    Set mContextSection = Nothing
    Set mContextMaterials = Nothing
    mContextBound = False
    mAcceptanceKey = vbNullString
End Sub

' Связывает сохранённые НДС с текущими моделью, материалами и настройками.
' При изменении этих условий прежние состояния больше не используются. Первый вызов
' не удаляет уже принятые результаты без ключа запроса, например найденное CapacityState.
Public Sub BindContext(ByVal section As CSectionModel, ByVal materials As CMaterialModelProvider, _
        ByVal acceptanceKey As String)
    AssertWritable
    If section Is Nothing Or materials Is Nothing Then _
        Err.Raise vbObjectError + 3986, "CStateRepository", "Не передан расчетный контекст repository."

    Dim changed As Boolean
    If mContextBound Then
        changed = (Not section Is mContextSection) Or (Not materials Is mContextMaterials)
        changed = changed Or (mSectionRevision <> section.Revision) Or _
            (mMaterialRevision <> materials.Revision) Or (mAcceptanceKey <> acceptanceKey)
        If changed Then Clear
    Else
        ' Без заданных условий расчёта нельзя подтвердить пригодность вручную сохранённого НДС
        ' для нового запроса.
        Set mReusableByKey = CreateObject("Scripting.Dictionary")
    End If
    Set mContextSection = section
    Set mContextMaterials = materials
    mSectionRevision = section.Revision
    mMaterialRevision = materials.Revision
    mAcceptanceKey = acceptanceKey
    mContextBound = True
End Sub

' Сохраняет готовое НДС без ключа запроса. Это нужно, например, для предельного
' состояния, найденного поиском несущей способности без отдельного CStateRequest.
Public Sub StoreState(ByVal stateResult As CSectionStateResult)
    AssertWritable
    If stateResult Is Nothing Then Exit Sub
    stateResult.Freeze
    Dim existingIndex As Long
    existingIndex = StateIndex(stateResult.StateType)
    If existingIndex > 0 Then
        If Not mStates.Item(existingIndex) Is stateResult Then ReplaceStateAt existingIndex, stateResult
        Exit Sub
    End If
    mStates.Add stateResult
End Sub

' Сохраняет состояние вместе с запросом. Для повторного использования допускает
' только найденное равновесие с полем деформаций. Ошибки NumFail/InputErr/CalcErr
' не запрещают последующую попытку решить тот же запрос.
Public Sub StoreForRequest(ByVal request As CStateRequest, ByVal stateResult As CSectionStateResult)
    AssertWritable
    If stateResult Is Nothing Then Exit Sub
    StoreState stateResult
    If request Is Nothing Then Exit Sub
    If Not IsReusablePhysicalState(stateResult) Then Exit Sub

    Dim key As String
    key = request.EquivalenceKey
    If mReusableByKey.Exists(key) Then
        Set mReusableByKey.Item(key) = stateResult
    Else
        mReusableByKey.Add key, stateResult
    End If
End Sub

' Ищет ранее найденное равновесное НДС для того же физического запроса.
' Неудачные попытки решения повторно используемым состоянием не считаются.
Public Function FindEquivalent(ByVal request As CStateRequest) As CSectionStateResult
    If request Is Nothing Then Exit Function
    Dim key As String
    key = request.EquivalenceKey
    If Not mReusableByKey.Exists(key) Then Exit Function
    Dim cached As CSectionStateResult
    Set cached = mReusableByKey.Item(key)
    If Not IsReusablePhysicalState(cached) Then Exit Function
    ' Дополнительно сравниваем исходные усилия: одинаковая текстовая запись ключа
    ' не должна объединять разные значения Double.
    If cached.TargetN <> request.TargetN Or cached.TargetMx <> request.TargetMx Or _
            cached.TargetMy <> request.TargetMy Then Exit Function
    Set FindEquivalent = cached
End Function

' Возвращает количество конечных расчётных состояний для записи в Results.
Public Property Get StateCount() As Long
    StateCount = mStates.Count
End Property

' Возвращает расчётное состояние по порядку его записи в Results.
Public Property Get StateAt(ByVal index As Long) As CSectionStateResult
    If index < 1 Or index > mStates.Count Then Err.Raise vbObjectError + 3985, "CStateRepository", "Индекс именованного состояния вне диапазона."
    Set StateAt = mStates.Item(index)
End Property

' Ищет конечное расчётное состояние заданного типа для этого сочетания.
Public Function FindState(ByVal stateType As ESectionStateType) As CSectionStateResult
    Dim i As Long
    For i = 1 To mStates.Count
        Dim stateResult As CSectionStateResult
        Set stateResult = mStates.Item(i)
        If stateResult.StateType = stateType Then
            Set FindState = stateResult
            Exit Function
        End If
    Next i
End Function

' Возвращает номер записи состояния данного типа в Results. Здесь хранится последний
' запрос этого типа, включая изменение усилий, повторное решение или получение прежнего НДС.
Private Function StateIndex(ByVal stateType As ESectionStateType) As Long
    Dim i As Long
    For i = 1 To mStates.Count
        Dim stateResult As CSectionStateResult
        Set stateResult = mStates.Item(i)
        If stateResult.StateType = stateType Then
            StateIndex = i
            Exit Function
        End If
    Next i
End Function

' Заменяет состояние в коллекции, сохраняя его порядковый номер для вывода в Results.
Private Sub ReplaceStateAt(ByVal index As Long, ByVal stateResult As CSectionStateResult)
    mStates.Add stateResult, , index
    mStates.Remove index + 1
End Sub

' Проверяет, можно ли использовать НДС повторно. Состояние с физическим FAIL
' можно использовать, если равновесие найдено; неудачная численная попытка
' без найденного поля деформаций для этого непригодна.
Private Function IsReusablePhysicalState(ByVal stateResult As CSectionStateResult) As Boolean
    If stateResult Is Nothing Then Exit Function
    If Not stateResult.Converged Then Exit Function
    Select Case stateResult.InternalStatus
        Case rsSuccess, rsSuccessWithWarning, rsCheckFailed
            IsReusablePhysicalState = True
    End Select
End Function

' Завершает заполнение хранилища и каждого сохранённого НДС, запрещая их изменение.
' Сетка не копируется; повторный вызов безопасен, добавление и очистка после него запрещены.
Friend Sub Freeze()
    If mFrozen Then Exit Sub
    Dim state As CSectionStateResult
    For Each state In mStates
        state.Freeze
    Next state
    mFrozen = True
End Sub

' Запрещает изменение результата после передачи потребителям.
' Для повторного расчёта создаётся новый объект, а прежний результат сохраняется.
Private Sub AssertWritable()
    If mFrozen Then Err.Raise vbObjectError + 4220, "CStateRepository", _
        "Опубликованный результат нельзя изменить; для нового расчета нужен новый объект."
End Sub
