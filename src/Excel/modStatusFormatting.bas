Attribute VB_Name = "modStatusFormatting"
Option Explicit

' ==========================================================================
' Цвета ячеек статусов на Results
' ==========================================================================
' Хранит общую палитру для записи результатов в Excel. Расчёт передаёт
' ResultMeta, CResultStatusPolicy определяет внешнее обозначение,
' а класс записи применяет к нему цвет из этого модуля.

' Возвращает утвержденный цвет фона для пользовательского статуса.
Public Function StatusFillColor(ByVal statusText As String) As Long
    Dim policy As CResultStatusPolicy
    Set policy = StatusPolicy()

    Select Case policy.ToUserStatus(statusText)
        Case policy.OK
            StatusFillColor = RGB(188, 222, 158)
        Case policy.Fail
            StatusFillColor = RGB(255, 190, 206)
        Case policy.BaseFail
            StatusFillColor = RGB(255, 190, 206)
        Case policy.NumFail
            StatusFillColor = RGB(255, 71, 71)
        Case policy.InputErr
            StatusFillColor = RGB(255, 230, 153)
        Case policy.CalcErr
            StatusFillColor = RGB(255, 71, 71)
        Case Else
            StatusFillColor = RGB(237, 237, 237)
    End Select
End Function

' Задаёт заливку одной ячейке или части строки Results по статусу.
' Шрифт, рамки и числовой формат сохраняет; вызывается
' после основной записи и оформления результатов.
Public Sub ApplyStatusFill(ByVal rangeObject As Object, ByVal statusText As String)
    If rangeObject Is Nothing Then Exit Sub

    With rangeObject
        .Interior.Color = StatusFillColor(statusText)
    End With
End Sub

' Читает таблицу одним массивом и красит ячейки одинакового цвета вместе.
' Адреса собираются в памяти; короткие порции обходят ограничение Excel
' на длину адреса. Прочие ячейки, шрифты и рамки не затрагиваются.
Public Sub ApplyTableStatusFills(ByVal target As Object, ByRef rowNumbers() As Long, _
        ByVal columnNumbers As Variant, Optional ByVal skipEmpty As Boolean = False, _
        Optional ByVal formationColumn As Long = 0, Optional ByRef preparedValues As Variant, _
        Optional ByVal preparedRowOffset As Long = 0)
    Dim values As Variant
    If IsMissing(preparedValues) Then
        values = target.Value2: preparedRowOffset = 0
    Else
        values = preparedValues
    End If
    Dim groups As Object: Set groups = CreateObject("Scripting.Dictionary")
    Dim column As Variant, row As Long, status As String, key As String
    Dim absoluteColumn As Long, columnText As String, remainder As Long
    Dim firstRow As Long, firstColumn As Long
    firstRow = target.Row: firstColumn = target.Column
    Dim sheet As Object: Set sheet = target.Worksheet
    Dim i As Long, spanFirst As Long, spanLast As Long, spanColor As String, skip As Boolean
    For Each column In columnNumbers
        absoluteColumn = firstColumn + CLng(column) - 1: columnText = vbNullString
        Do While absoluteColumn > 0
            remainder = (absoluteColumn - 1) Mod 26
            columnText = Chr$(65 + remainder) & columnText
            absoluteColumn = (absoluteColumn - 1) \ 26
        Loop
        spanFirst = 0
        For i = LBound(rowNumbers) To UBound(rowNumbers)
            row = rowNumbers(i)
            status = CStr(values(row - preparedRowOffset, CLng(column)))
            skip = skipEmpty And Len(Trim$(status)) = 0
            If CLng(column) = formationColumn Then
                If status = "Cracked" Or status = "NotCracked" Then skip = True
            End If
            If skip Then
                If spanFirst > 0 Then AppendStatusFillSpan sheet, groups, spanColor, columnText, firstRow + spanFirst - 1, firstRow + spanLast - 1
                spanFirst = 0
            Else
                key = CStr(StatusFillColor(status))
                If spanFirst > 0 And (key <> spanColor Or row <> spanLast + 1) Then
                    AppendStatusFillSpan sheet, groups, spanColor, columnText, firstRow + spanFirst - 1, firstRow + spanLast - 1
                    spanFirst = 0
                End If
                If spanFirst = 0 Then spanFirst = row: spanColor = key
                spanLast = row
            End If
        Next i
        If spanFirst > 0 Then AppendStatusFillSpan sheet, groups, spanColor, columnText, firstRow + spanFirst - 1, firstRow + spanLast - 1
    Next column
    Dim color As Variant
    For Each color In groups.Keys
        If Len(groups.Item(color)) > 0 Then sheet.Range(groups.Item(color)).Interior.Color = CLng(color)
    Next color
End Sub

' Объединяет соседние строки одного цвета, не затрагивая пропуски и незаполненные LC.
' Ограничение длины адреса сохраняет прежнюю совместимость Excel при больших пакетах.
Private Sub AppendStatusFillSpan(ByVal sheet As Object, ByVal groups As Object, ByVal color As String, _
        ByVal columnText As String, ByVal first As Long, ByVal last As Long)
    Dim address As String: address = columnText & CStr(first)
    If last > first Then address = address & ":" & columnText & CStr(last)
    If Not groups.Exists(color) Then groups.Add color, vbNullString
    If Len(groups.Item(color)) + Len(address) > 200 Then
        sheet.Range(groups.Item(color)).Interior.Color = CLng(color)
        groups.Item(color) = vbNullString
    End If
    If Len(groups.Item(color)) > 0 Then address = "," & address
    groups.Item(color) = groups.Item(color) & address
End Sub

' Возвращает общий словарь статусов для модуля цветового оформления.
Private Function StatusPolicy() As CResultStatusPolicy
    Static policy As CResultStatusPolicy
    If policy Is Nothing Then Set policy = New CResultStatusPolicy
    Set StatusPolicy = policy
End Function
