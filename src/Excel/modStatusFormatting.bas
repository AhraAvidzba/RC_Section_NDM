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

' Возвращает общий словарь статусов для модуля цветового оформления.
Private Function StatusPolicy() As CResultStatusPolicy
    Static policy As CResultStatusPolicy
    If policy Is Nothing Then Set policy = New CResultStatusPolicy
    Set StatusPolicy = policy
End Function
