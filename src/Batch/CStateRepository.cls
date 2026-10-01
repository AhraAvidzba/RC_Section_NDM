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
' Repository конечных named-state одного LC
' ==========================================================================
' CStateRepository хранит все конечные named-state, которые должны попасть в
' Results snapshot, и отдельно держит кэш переиспользуемых физических состояний.
' Неуспешные попытки solve можно сохранить для вывода статуса/диагностики, но
' FindEquivalent их не возвращает: новый solve с другим warm-start должен
' оставаться возможным.

Private mStates As Collection ' Все конечные named-state для вывода Results.
Private mReusableByKey As Object ' Dictionary: equivalence key -> успешно полученный физический state.
Private mContextSection As CSectionModel ' Identity геометрии, которой принадлежит reusable cache.
Private mContextMaterials As CMaterialModelProvider ' Identity эффективного набора материалов.
Private mSectionRevision As Long ' Версия геометрии на момент привязки контекста.
Private mMaterialRevision As Long ' Версия материалов на момент привязки контекста.
Private mAcceptanceKey As String ' Effective Extension и допуски приемки равновесия.
Private mContextBound As Boolean ' Контекст cache уже определен общим State pipeline.

Private Sub Class_Initialize()
    Set mStates = New Collection
    Set mReusableByKey = CreateObject("Scripting.Dictionary")
End Sub

' Очищает snapshot states и reusable-cache перед новым расчетом LC.
Public Sub Clear()
    Set mStates = New Collection
    Set mReusableByKey = CreateObject("Scripting.Dictionary")
    Set mContextSection = Nothing
    Set mContextMaterials = Nothing
    mContextBound = False
    mAcceptanceKey = vbNullString
End Sub

' Ограничивает reusable cache одним фактическим расчетным контекстом.
' При его смене старые состояния инвалидируются. Первый bind не удаляет
' unkeyed named states, уже принятые владельцем LC, например CapacityState.
Public Sub BindContext(ByVal section As CSectionModel, ByVal materials As CMaterialModelProvider, _
        ByVal acceptanceKey As String)
    If section Is Nothing Or materials Is Nothing Then _
        Err.Raise vbObjectError + 3986, "CStateRepository", "Не передан расчетный контекст repository."

    Dim changed As Boolean
    If mContextBound Then
        changed = (Not section Is mContextSection) Or (Not materials Is mContextMaterials)
        changed = changed Or (mSectionRevision <> section.Revision) Or _
            (mMaterialRevision <> materials.Revision) Or (mAcceptanceKey <> acceptanceKey)
        If changed Then Clear
    Else
        ' Без контекста вручную сохраненные states не являются доказанным reuse.
        Set mReusableByKey = CreateObject("Scripting.Dictionary")
    End If
    Set mContextSection = section
    Set mContextMaterials = materials
    mSectionRevision = section.Revision
    mMaterialRevision = materials.Revision
    mAcceptanceKey = acceptanceKey
    mContextBound = True
End Sub

' Сохраняет state без ключа переиспользования. Такой путь нужен для уже
' найденных capacity-точек и других физических states без CStateRequest.
Public Sub StoreState(ByVal stateResult As CSectionStateResult)
    If stateResult Is Nothing Then Exit Sub
    Dim existingIndex As Long
    existingIndex = StateIndex(stateResult.StateType)
    If existingIndex > 0 Then
        If Not mStates.Item(existingIndex) Is stateResult Then ReplaceStateAt existingIndex, stateResult
        Exit Sub
    End If
    mStates.Add stateResult
End Sub

' Сохраняет state, связанный с физическим request. В общий reusable cache
' попадают только сошедшиеся состояния, то есть состояния с реально полученной
' плоскостью деформаций. NumFail/InputErr/CalcErr не блокируют будущие попытки.
Public Sub StoreForRequest(ByVal request As CStateRequest, ByVal stateResult As CSectionStateResult)
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

' Ищет ранее полученное эквивалентное физическое состояние. Метод никогда не
' возвращает неуспешные попытки solve.
Public Function FindEquivalent(ByVal request As CStateRequest) As CSectionStateResult
    If request Is Nothing Then Exit Function
    Dim key As String
    key = request.EquivalenceKey
    If Not mReusableByKey.Exists(key) Then Exit Function
    Dim cached As CSectionStateResult
    Set cached = mReusableByKey.Item(key)
    If Not IsReusablePhysicalState(cached) Then Exit Function
    ' Строковое форматирование ключа не должно объединять разные Double-loads.
    If cached.TargetN <> request.TargetN Or cached.TargetMx <> request.TargetMx Or _
            cached.TargetMy <> request.TargetMy Then Exit Function
    Set FindEquivalent = cached
End Function

' Возвращает количество states, которые должны попасть в Results snapshot.
Public Property Get StateCount() As Long
    StateCount = mStates.Count
End Property

' Возвращает state по порядку записи в snapshot.
Public Property Get StateAt(ByVal index As Long) As CSectionStateResult
    If index < 1 Or index > mStates.Count Then Err.Raise vbObjectError + 3985, "CStateRepository", "Индекс именованного состояния вне диапазона."
    Set StateAt = mStates.Item(index)
End Property

' Ищет state по типу среди конечных states LC.
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

' Возвращает позицию named-state в snapshot. В этой позиции хранится последний
' запрос данного типа, включая новое усилие, повтор solve или cache-hit.
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

' Заменяет элемент Collection с сохранением его позиции в Results snapshot.
Private Sub ReplaceStateAt(ByVal index As Long, ByVal stateResult As CSectionStateResult)
    mStates.Add stateResult, , index
    mStates.Remove index + 1
End Sub

' Определяет, можно ли отдавать state через FindEquivalent. Физический FAIL из-за
' выхода за пределы диаграммы остается reusable, потому что равновесие найдено;
' численная неудача без найденного НДС reusable не является.
Private Function IsReusablePhysicalState(ByVal stateResult As CSectionStateResult) As Boolean
    If stateResult Is Nothing Then Exit Function
    If Not stateResult.Converged Then Exit Function
    Select Case stateResult.InternalStatus
        Case rsSuccess, rsSuccessWithWarning, rsCheckFailed
            IsReusablePhysicalState = True
    End Select
End Function
