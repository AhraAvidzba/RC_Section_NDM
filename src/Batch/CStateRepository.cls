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

Private Sub Class_Initialize()
    Set mStates = New Collection
    Set mReusableByKey = CreateObject("Scripting.Dictionary")
End Sub

' Очищает snapshot states и reusable-cache перед новым расчетом LC.
Public Sub Clear()
    Set mStates = New Collection
    Set mReusableByKey = CreateObject("Scripting.Dictionary")
End Sub

' Сохраняет state без ключа переиспользования. Такой путь нужен для уже
' найденных capacity-точек и переходного кода, где request пока не создается.
Public Sub StoreState(ByVal stateResult As CSectionStateResult)
    If stateResult Is Nothing Then Exit Sub
    If HasStateType(stateResult.StateType) Then Exit Sub
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
    If Not mReusableByKey.Exists(key) Then mReusableByKey.Add key, stateResult
End Sub

' Ищет ранее полученное эквивалентное физическое состояние. Метод никогда не
' возвращает неуспешные попытки solve.
Public Function FindEquivalent(ByVal request As CStateRequest) As CSectionStateResult
    If request Is Nothing Then Exit Function
    Dim key As String
    key = request.EquivalenceKey
    If mReusableByKey.Exists(key) Then Set FindEquivalent = mReusableByKey.Item(key)
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

' Проверяет, есть ли уже state такого типа в snapshot. В текущей архитектуре у
' одного LC должен быть не более одного StrengthState, CrackedState и т. п.
Private Function HasStateType(ByVal stateType As ESectionStateType) As Boolean
    HasStateType = Not FindState(stateType) Is Nothing
End Function

' Определяет, можно ли отдавать state через FindEquivalent. Физический FAIL из-за
' выхода за пределы диаграммы остается reusable, потому что равновесие найдено;
' численная неудача без найденного НДС reusable не является.
Private Function IsReusablePhysicalState(ByVal stateResult As CSectionStateResult) As Boolean
    IsReusablePhysicalState = stateResult.Converged
End Function
