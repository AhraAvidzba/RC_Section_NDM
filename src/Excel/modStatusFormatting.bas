Attribute VB_Name = "modStatusFormatting"
Option Explicit

' ==========================================================================
' Цветовое оформление пользовательских статусов Results
' ==========================================================================
' Модуль хранит единую палитру статусов для Excel-вывода. Расчетные классы
' продолжают работать только с текстовыми статусами, а writer-ы применяют эти
' цвета уже на последнем шаге оформления таблиц Results.

' Возвращает утвержденный цвет фона для пользовательского статуса.
Public Function StatusFillColor(ByVal statusText As String) As Long
    Dim policy As CBatchStatusPolicy
    Set policy = StatusPolicy()

    Select Case policy.ToUserStatus(statusText)
        Case policy.OK
            StatusFillColor = RGB(188, 222, 158)
        Case policy.Fail
            StatusFillColor = RGB(255, 190, 206)
        Case policy.NumFail
            StatusFillColor = RGB(255, 71, 71)
        Case policy.InputErr
            StatusFillColor = RGB(255, 230, 153)
        Case Else
            StatusFillColor = RGB(237, 237, 237)
    End Select
End Function

' Наносит статусную заливку на одну ячейку или блок строки Results.
' Шрифты, границы и числовые форматы не трогает: writer-ы вызывают эту
' процедуру только как последний слой цветовой индикации статуса.
Public Sub ApplyStatusFill(ByVal rangeObject As Object, ByVal statusText As String)
    If rangeObject Is Nothing Then Exit Sub

    With rangeObject
        .Interior.Color = StatusFillColor(statusText)
    End With
End Sub

' Возвращает общий словарь статусов для модуля цветового оформления.
Private Function StatusPolicy() As CBatchStatusPolicy
    Static policy As CBatchStatusPolicy
    If policy Is Nothing Then Set policy = New CBatchStatusPolicy
    Set StatusPolicy = policy
End Function
