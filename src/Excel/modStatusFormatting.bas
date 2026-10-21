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
        Optional ByVal formationColumn As Long = 0)
    Dim values As Variant: values = target.Value2
    Dim groups As Object: Set groups = CreateObject("Scripting.Dictionary")
    Dim letters As Object: Set letters = CreateObject("Scripting.Dictionary")
    Dim column As Variant, row As Long, status As String, key As String, address As String
    Dim absoluteColumn As Long, columnText As String, remainder As Long
    Dim firstRow As Long, firstColumn As Long
    firstRow = target.Row: firstColumn = target.Column
    For Each column In columnNumbers
        absoluteColumn = firstColumn + CLng(column) - 1: columnText = vbNullString
        Do While absoluteColumn > 0
            remainder = (absoluteColumn - 1) Mod 26
            columnText = Chr$(65 + remainder) & columnText
            absoluteColumn = (absoluteColumn - 1) \ 26
        Loop
        letters.Add CLng(column), columnText
    Next column
    Dim i As Long
    For i = LBound(rowNumbers) To UBound(rowNumbers)
        row = rowNumbers(i)
        For Each column In columnNumbers
            status = CStr(values(row, CLng(column)))
            If skipEmpty And Len(Trim$(status)) = 0 Then GoTo NextCell
            If CLng(column) = formationColumn Then
                If status = "Cracked" Or status = "NotCracked" Then GoTo NextCell
            End If
            key = CStr(StatusFillColor(status))
            address = letters.Item(CLng(column)) & CStr(firstRow + row - 1)
            If Not groups.Exists(key) Then groups.Add key, vbNullString
            If Len(groups.Item(key)) + Len(address) > 200 Then
                target.Worksheet.Range(groups.Item(key)).Interior.Color = CLng(key)
                groups.Item(key) = vbNullString
            End If
            If Len(groups.Item(key)) > 0 Then address = "," & address
            groups.Item(key) = groups.Item(key) & address
NextCell:
        Next column
    Next i
    Dim color As Variant
    For Each color In groups.Keys
        target.Worksheet.Range(groups.Item(color)).Interior.Color = CLng(color)
    Next color
End Sub

' Возвращает общий словарь статусов для модуля цветового оформления.
Private Function StatusPolicy() As CResultStatusPolicy
    Static policy As CResultStatusPolicy
    If policy Is Nothing Then Set policy = New CResultStatusPolicy
    Set StatusPolicy = policy
End Function
