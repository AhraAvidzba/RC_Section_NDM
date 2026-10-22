Attribute VB_Name = "modResultsFormatting"
Option Explicit

' ==========================================================================
' Повторное использование готовых таблиц Results
' ==========================================================================
' Сравнивает подписи шапок и записывает изменяемые части сводки массивами.
' Оформление не кэшируется: актуальные подписи читаются из самой книги.

' Проверяет текст шапки вместе с единицами без обращения к каждой ячейке.
' Для общей сводки можно сравнить только строки шапки, исключив метаданные.
Public Function ResultHeaderMatches(ByVal target As Object, ByRef expected As Variant, _
        Optional ByVal firstRow As Long = 1, Optional ByVal lastRow As Long = 0, _
        Optional ByVal firstColumn As Long = 1, Optional ByVal merges As Collection = Nothing) As Boolean
    If lastRow = 0 Then lastRow = UBound(expected, 1)
    Dim actual As Variant
    actual = target.Offset(firstRow - 1, 0).Resize(lastRow - firstRow + 1, UBound(expected, 2)).Value2
    Dim normalized As Variant: normalized = expected ' Копия массива нужна, чтобы объединённые заголовки не меняли входной шаблон.
    If Not merges Is Nothing Then NormalizeMergedHeader normalized, merges, firstRow, lastRow
    Dim row As Long, column As Long
    For row = firstRow To lastRow
        For column = firstColumn To UBound(expected, 2)
            If IsError(actual(row - firstRow + 1, column)) Then Exit Function
            If CStr(actual(row - firstRow + 1, column)) <> CStr(normalized(row, column)) Then Exit Function
        Next column
    Next row
    ResultHeaderMatches = True
End Function

' Excel переносит первую непустую подпись в начало объединяемой области.
' Повторяем только этот перенос в памяти, не меняя ячейки и исходный массив.
Private Sub NormalizeMergedHeader(ByRef data As Variant, ByVal blocks As Collection, _
        ByVal firstRow As Long, ByVal lastRow As Long)
    Dim block As Variant, row As Long, column As Long, value As Variant
    For Each block In blocks ' Четвёрка: первая/последняя строка, первый/последний столбец объединения; индексы массива 0..3.
        If CLng(block(0)) >= firstRow And CLng(block(1)) <= lastRow Then
            value = Empty
            For row = CLng(block(0)) To CLng(block(1))
                For column = CLng(block(2)) To CLng(block(3))
                    If IsEmpty(value) And Len(CStr(data(row, column))) > 0 Then value = data(row, column)
                    data(row, column) = Empty
                Next column
            Next row
            data(CLng(block(0)), CLng(block(2))) = value
        End If
    Next block
End Sub

' Обновляет один прямоугольник данных, сохраняя соседние объединённые шапки.
Public Sub WriteResultDataSlice(ByVal target As Object, ByRef data() As Variant, _
        ByVal firstRow As Long, ByVal firstColumn As Long, ByVal rows As Long, ByVal columns As Long)
    Dim output() As Variant, row As Long, column As Long
    ReDim output(1 To rows, 1 To columns)
    For row = 1 To rows
        For column = 1 To columns
            output(row, column) = data(firstRow + row - 1, firstColumn + column - 1)
        Next column
    Next row
    target.Offset(firstRow - 1, firstColumn - 1).Resize(rows, columns).Value2 = output
End Sub
