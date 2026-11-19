VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CResultsLayoutValidator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Общие адресные рекомендации по вместимости таблиц Results
' ==========================================================================
' Проверяет интервалы строк по фактическим якорям до изменения листа.
' Одинаковые правила используются перед расчётом и при выводе снимка;
' класс не читает Excel, не вставляет строки и не изменяет данные.
Private mIssues As Collection ' Причины отказа, включая готовые рекомендации по вставке строк.

' Создаёт отдельный список проблем для одного действия пользователя.
Private Sub Class_Initialize()
    Set mIssues = New Collection
End Sub

' Возвращает общий список для проверок габаритов листа и отсутствующих якорей.
Public Property Get Issues() As Collection
    Set Issues = mIssues
End Property

' Проверяет свободный интервал; вставка выполняется над заголовком следующей таблицы.
Public Sub CheckRowsGap(ByVal rangeName As String, ByVal previousBottomRow As Long, _
        ByVal nextTopRow As Long, ByVal insertBeforeRow As Long, Optional ByVal gapRows As Long = 2)
    If previousBottomRow + gapRows < nextTopRow Then Exit Sub
    Dim count As Long: count = previousBottomRow + gapRows + 1 - nextTopRow
    If insertBeforeRow < 1 Then insertBeforeRow = nextTopRow
    mIssues.Add "На листе Results вставьте " & CStr(count) & " " & RowsWord(count) & _
        " над строкой " & CStr(insertBeforeRow) & ", чтобы опустить диапазон " & rangeName & "."
End Sub

' Собирает понятный итог с числом сочетаний, конкретными строками и доступным действием.
Public Function Message(ByVal combinationCount As Long, Optional ByVal stoppedAction As String = "Расчет не запущен.", _
        Optional ByVal canReduceInput As Boolean = True) As String
    Dim text As String, index As Long
    text = "Все сочетания (" & CStr(combinationCount) & _
        ") не помещаются в текущую раскладку или пределы листа Results. " & stoppedAction & vbCrLf & "Что сделать:"
    For index = 1 To mIssues.Count
        text = text & vbCrLf & CStr(index) & ") " & CStr(mIssues(index))
    Next index
    If canReduceInput Then text = text & vbCrLf & CStr(mIssues.Count + 1) & ") Либо на листе Config уменьшите число строк в rngLoadCombinations."
    Message = text
End Function

' Возвращает склонение количества вставляемых строк.
Private Function RowsWord(ByVal count As Long) As String
    Dim lastTwo As Long: lastTwo = Abs(count) Mod 100
    If lastTwo >= 11 And lastTwo <= 14 Then
        RowsWord = "строк"
    Else
        Select Case Abs(count) Mod 10
            Case 1: RowsWord = "строку"
            Case 2, 3, 4: RowsWord = "строки"
            Case Else: RowsWord = "строк"
        End Select
    End If
End Function
