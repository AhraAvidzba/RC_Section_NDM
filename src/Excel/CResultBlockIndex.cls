VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "CResultBlockIndex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Общий индекс фактических непрерывных участков результатов элементов
' ==========================================================================
' Знает ключи и относительные строки; физические адреса файла и Excel сюда не входят.
' Для indexOnly возвращает только RunID/LC/Profile/State; значения элементов читаются отдельным запросом.
Private mBlocks As Collection ' Каждый элемент: FirstRow, RowCount и девять исходных ячеек первой строки.
Private mHeader As Variant ' Исходная шапка сохраняет подписи и единицы.
Private mRunID As Double, mTotalRows As Long

' Создаёт независимый пустой индекс; порядок коллекции равен порядку результатов.
Private Sub Class_Initialize()
    Set mBlocks = New Collection: mTotalRows = 1
End Sub

' Одним проходом находит все участки, не объединяя раздельные повторения ключа.
' Двоичное разделение сохраняет исходный регистр; отбор ниже выполняется без учёта регистра.
Public Sub Build(ByRef values As Variant, ByVal runID As Double, Optional ByVal runIDsValidated As Boolean = False)
    Dim row As Long, first As Long, previousCase As String, previousProfile As String, previousState As String
    Dim currentCase As String, currentProfile As String, currentState As String
    Dim header(1 To 1, 1 To 9) As Variant, column As Long
    If LBound(values, 1) <> 1 Or LBound(values, 2) <> 1 Or UBound(values, 2) <> 9 Then Err.Raise 5, , "Неверная структура rngNDMElementResults."
    For column = 1 To 9: header(1, column) = values(1, column): Next column
    BuildCore values, header, runID, runIDsValidated
End Sub


' Однократно индексирует таблицу без каталога по RunID/LC/Profile/State, без всех девяти колонок.
Public Sub BuildKeyColumns(ByRef keys As Variant, ByRef header As Variant, ByVal runID As Double)
    If UBound(keys, 2) <> 4 Then Err.Raise 5, , "Для перестроения индекса нужны четыре ключевые колонки."
    BuildCore keys, header, runID, False
End Sub

' Общие правила группировки одинаковы для массива расчёта и однократного чтения ключей Excel.
Private Sub BuildCore(ByRef values As Variant, ByRef header As Variant, ByVal runID As Double, ByVal runIDsValidated As Boolean)
    Dim row As Long, first As Long, previousCase As String, previousProfile As String, previousState As String
    Dim caseKind As Long, profileKind As Long, stateKind As Long, changed As Boolean
    Initialize header, runID
    If UBound(values, 1) < 2 Then Exit Sub
    first = 2
    previousCase = CStr(values(2, 2)): previousProfile = CStr(values(2, 3)): previousState = CStr(values(2, 4))
    caseKind = VarType(values(2, 2)): profileKind = VarType(values(2, 3)): stateKind = VarType(values(2, 4))
    If Len(previousCase) = 0 Or Len(previousProfile) = 0 Or Len(previousState) = 0 Then Err.Raise 5, , "Пустой ключ сочетания, профиля или состояния."
    For row = 2 To UBound(values, 1)
        If Not runIDsValidated Then
            If Not IsNumeric(values(row, 1)) And VarType(values(row, 1)) <> vbDate Then Err.Raise 5, , "В результатах элементов отсутствует RunID."
            If CDbl(values(row, 1)) <> runID Then Err.Raise 5, , "Результаты элементов относятся к другому расчёту."
        End If
        changed = StrComp(values(row, 2), previousCase, vbBinaryCompare) <> 0 Or StrComp(values(row, 3), previousProfile, vbBinaryCompare) <> 0 Or _
            StrComp(values(row, 4), previousState, vbBinaryCompare) <> 0 Or VarType(values(row, 2)) <> caseKind Or VarType(values(row, 3)) <> profileKind Or VarType(values(row, 4)) <> stateKind
        If changed Then
            AppendArrayBlock values, first, row - first
            first = row
            previousCase = CStr(values(row, 2)): previousProfile = CStr(values(row, 3)): previousState = CStr(values(row, 4))
            caseKind = VarType(values(row, 2)): profileKind = VarType(values(row, 3)): stateKind = VarType(values(row, 4))
            If Len(previousCase) = 0 Or Len(previousProfile) = 0 Or Len(previousState) = 0 Then Err.Raise 5, , "Пустой ключ сочетания, профиля или состояния."
        End If
    Next row
    AppendArrayBlock values, first, UBound(values, 1) - first + 1
End Sub

' Устанавливает шапку до импорта записей каталога, проверяя фиксированный контракт.
Public Sub Initialize(ByRef header As Variant, ByVal runID As Double)
    If UBound(header, 1) <> 1 Or UBound(header, 2) <> 9 Then Err.Raise 5, , "Повреждена шапка индекса результатов."
    Dim column As Long, expected As Variant
    expected = Array("RunID", "LoadCase", "ProfileId", "StateType", "ElementID", "MaterialType", "Strain", "Stress", "PhysicalState")
    For column = 1 To 9
        If ResultHeaderBase(CStr(header(1, column))) <> CStr(expected(column - 1)) Then Err.Raise 5, , "Повреждена колонка индекса: " & CStr(column)
    Next column
    Set mBlocks = New Collection: mHeader = header: mRunID = runID: mTotalRows = 1
End Sub

' Копирует только одну представительскую строку; большая матрица не дублируется.
Private Sub AppendArrayBlock(ByRef values As Variant, ByVal first As Long, ByVal rows As Long)
    Dim sample(1 To 1, 1 To 9) As Variant, column As Long
    For column = 1 To UBound(values, 2): sample(1, column) = values(first, column): Next column
    Append first, rows, sample
End Sub

' Принимает проверенную запись каталога строго в исходном порядке без дыр и наложений.
Public Sub Append(ByVal first As Long, ByVal rows As Long, ByRef sample As Variant)
    If first <> mTotalRows + 1 Or rows < 1 Or CDbl(first) + rows > 2147483647# Then Err.Raise 5, , "Неверные относительные строки индекса результатов."
    If UBound(sample, 1) <> 1 Or UBound(sample, 2) <> 9 Then Err.Raise 5, , "Повреждена строка индекса результатов."
    If CDbl(sample(1, 1)) <> mRunID Then Err.Raise 5, , "Индекс принадлежит другому расчёту."
    If Len(CStr(sample(1, 2))) = 0 Or Len(CStr(sample(1, 3))) = 0 Or Len(CStr(sample(1, 4))) = 0 Then Err.Raise 5, , "Пустой ключ в индексе результатов."
    mBlocks.Add Array(first, rows, sample): mTotalRows = mTotalRows + rows
End Sub

' Возвращает все подходящие участки, включая несмежные повторения одного ключа.
' Пустой фильтр означает все значения соответствующего поля.
Public Function SelectBlocks(ByVal combinationID As String, ByVal stateType As String, Optional ByVal profileID As String = "") As Collection
    Dim selected As New Collection, block As Long, record As Variant, sample As Variant
    For block = 1 To mBlocks.Count
        record = mBlocks.Item(block): sample = record(2)
        If (Len(combinationID) = 0 Or StrComp(CStr(sample(1, 2)), combinationID, vbTextCompare) = 0) And _
                (Len(stateType) = 0 Or StrComp(CStr(sample(1, 4)), stateType, vbTextCompare) = 0) And _
                (Len(profileID) = 0 Or StrComp(CStr(sample(1, 3)), profileID, vbTextCompare) = 0) Then selected.Add block
    Next block
    Set SelectBlocks = selected
End Function

' Перечень состояний содержит только используемые ключи каждой пары LC/State, без копии результатов элемента.
' Ключ с длиной исключает коллизии встроенных разделителей в Unicode-идентификаторах.
Public Function StateRows(ByVal combinationID As String, ByVal stateType As String, Optional ByVal profileID As String = "") As Variant
    Dim selected As Collection: Set selected = SelectBlocks(combinationID, stateType, profileID)
    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary"): seen.CompareMode = vbTextCompare
    Dim block As Variant, sample As Variant, identity As String, count As Long, row As Long, column As Long
    For Each block In selected
        sample = Representative(CLng(block)): identity = CStr(Len(CStr(sample(1, 2)))) & ":" & CStr(sample(1, 2)) & CStr(sample(1, 4))
        If Not seen.Exists(identity) Then seen.Add identity, CLng(block): count = count + 1
    Next block
    Dim values() As Variant: ReDim values(1 To count + 1, 1 To 4)
    For column = 1 To 4: values(1, column) = mHeader(1, column): Next column
    row = 1
    For Each block In seen.Items
        row = row + 1: sample = Representative(CLng(block))
        For column = 1 To 4: values(row, column) = sample(1, column): Next column
    Next block
    StateRows = values
End Function

' Возвращает исходную шапку без доступа к массиву всех результатов.
Public Property Get Header() As Variant
    Header = mHeader
End Property

' Возвращает число непрерывных участков.
Public Property Get Count() As Long
    Count = mBlocks.Count
End Property

' Возвращает высоту полной таблицы вместе с шапкой.
Public Property Get TotalRows() As Long
    TotalRows = mTotalRows
End Property

' Возвращает исходный RunID для согласования с носителем.
Public Property Get RunID() As Double
    RunID = mRunID
End Property

' Возвращает первую строку относительно именованного якоря, включая шапку.
Public Function FirstRow(ByVal block As Long) As Long
    Dim record As Variant: record = mBlocks.Item(block): FirstRow = CLng(record(0))
End Function

' Возвращает число строк выбранного участка.
Public Function RowCount(ByVal block As Long) As Long
    Dim record As Variant: record = mBlocks.Item(block): RowCount = CLng(record(1))
End Function

' Возвращает независимую первую строку участка для индекса доступности.
Public Function Representative(ByVal block As Long) As Variant
    Dim record As Variant: record = mBlocks.Item(block): Representative = record(2)
End Function

' Дополняет индекс ключей только первой полной строкой каждого участка, до его публикации.
Public Sub SetRepresentative(ByVal block As Long, ByRef sample As Variant)
    Dim record As Variant: record = mBlocks.Item(block)
    Dim expected As Variant, column As Long: expected = record(2)
    If UBound(sample, 1) <> 1 Or UBound(sample, 2) <> 9 Then Err.Raise 5, , "Неверная представительская строка индекса."
    For column = 1 To 4
        If StrComp(CStr(sample(1, column)), CStr(expected(1, column)), vbBinaryCompare) <> 0 Then Err.Raise 5, , "Таблица изменилась во время построения индекса."
    Next column
    mBlocks.Remove block
    If block > mBlocks.Count Then
        mBlocks.Add Array(record(0), record(1), sample)
    Else
        mBlocks.Add Array(record(0), record(1), sample), , block
    End If
End Sub

' Принимает фактически созданный участок writer-а; обход всех результатов ради границ не нужен.
Friend Sub AppendPreparedBlock(ByRef values() As Variant, ByVal first As Long, ByVal rows As Long)
    If rows < 1 Then Exit Sub
    Dim sample(1 To 1, 1 To 9) As Variant, column As Long
    For column = 1 To 9: sample(1, column) = values(first, column): Next column
    Append first, rows, sample
End Sub
