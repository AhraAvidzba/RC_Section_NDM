VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLoadCombinationReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Чтение таблицы сочетаний нагрузок
' ==========================================================================
' Reader переносит строки rngLoadCombinations в CBatchSectionCalculator. Пустой
' Mx или My считается нулем, полностью нулевая строка нагрузок пропускается,
' а неполные/некорректные строки сохраняются как InvalidInput, чтобы пользователь
' увидел их в отчете.
'
' Если передан CUnitSystem, усилия сразу переводятся из INPUT-единиц и
' пользовательской системы знаков во внутреннюю систему расчетного ядра.

' Загружает сочетания из именованного диапазона книги.
Public Sub LoadFromWorkbook(ByVal workbook As Object, ByVal batch As CBatchSectionCalculator, Optional ByVal units As CUnitSystem = Nothing)
    If workbook Is Nothing Then Err.Raise vbObjectError + 3950, "CLoadCombinationReader", "Книга Excel не передана."
    If batch Is Nothing Then Err.Raise vbObjectError + 3951, "CLoadCombinationReader", "Пакетный расчетчик не передан."
    LoadFromRange workbook.Names.Item("rngLoadCombinations").RefersToRange, batch, units
End Sub

' Загружает сочетания из уже переданного диапазона.
' Это используется и рабочей книгой, и тестами, чтобы проверять reader без
' ручного запуска Excel-интерфейса.
Public Sub LoadFromRange(ByVal source As Object, ByVal batch As CBatchSectionCalculator, Optional ByVal units As CUnitSystem = Nothing)
    If source Is Nothing Then Err.Raise vbObjectError + 3952, "CLoadCombinationReader", "Диапазон сочетаний нагрузок не найден."
    If batch Is Nothing Then Err.Raise vbObjectError + 3953, "CLoadCombinationReader", "Пакетный расчетчик не передан."
    If source.Rows.Count < 2 Then Err.Raise vbObjectError + 3954, "CLoadCombinationReader", _
        "rngLoadCombinations должен содержать строку заголовков и минимум одну строку сочетания."
    If source.Columns.Count < 5 Then Err.Raise vbObjectError + 3955, "CLoadCombinationReader", _
        "rngLoadCombinations должен содержать минимум столбцы CombinationID, N, Mx, My и ProfileId."

    Dim data As Variant
    data = source.Value2

    Dim rowIndex As Long
    ' Пустые строки оставляем пользователю для ввода, а не превращаем в ошибки.
    For rowIndex = 2 To UBound(data, 1)
        If IsEmptyCombinationRow(data, rowIndex) Then
            ' Строка свободна для будущего ввода LC.
        Else
            AddRow data, batch, rowIndex, units, source
        End If
    Next rowIndex
End Sub

' Читает одну строку LC и добавляет ее в batch.
' Ошибки ввода не прерывают весь расчет: строка попадает в отчет как InvalidInput,
' остальные сочетания продолжают считаться.
Private Sub AddRow(ByRef data As Variant, ByVal batch As CBatchSectionCalculator, ByVal rowIndex As Long, _
        ByVal units As CUnitSystem, ByVal source As Object)
    Dim combinationID As String
    Dim profileId As String
    Dim loadPathText As String
    Dim comment As String
    Dim errorText As String
    Dim pathLocation As String
    pathLocation = "LoadPath"
    If UBound(data, 2) >= 6 Then pathLocation = CellLocation(source, rowIndex, 6, "LoadPath")
    combinationID = ReadCellText(data(rowIndex, 1), CellLocation(source, rowIndex, 1, "CombinationID"), errorText)
    profileId = ReadCellText(data(rowIndex, 5), CellLocation(source, rowIndex, 5, "ProfileId"), errorText)
    If UBound(data, 2) >= 7 Then
        loadPathText = ReadCellText(data(rowIndex, 6), pathLocation, errorText)
        comment = ReadCellText(data(rowIndex, 7), CellLocation(source, rowIndex, 7, "Комментарий"), errorText)
    ElseIf UBound(data, 2) = 6 Then
        If StrComp(CStr(data(1, 6)), "LoadPath", vbTextCompare) = 0 Then
            loadPathText = ReadCellText(data(rowIndex, 6), pathLocation, errorText)
        Else
            comment = ReadCellText(data(rowIndex, 6), CellLocation(source, rowIndex, 6, "Комментарий"), errorText)
        End If
    End If
    Dim pathDescriptor As CLoadPathDescriptor
    Set pathDescriptor = New CLoadPathDescriptor
    If Len(pathDescriptor.NormalizeKey(loadPathText)) = 0 Then _
        AppendError errorText, pathLocation & ": выберите Auto, lambda*Mx, lambda*My, lambda*Mxy, lambda*N или lambda*NMxy"

    Dim nValue As Double
    Dim mxValue As Double
    Dim myValue As Double
    nValue = ReadOptionalLoadDouble(data(rowIndex, 2), CellLocation(source, rowIndex, 2, "N"), errorText)
    mxValue = ReadOptionalLoadDouble(data(rowIndex, 3), CellLocation(source, rowIndex, 3, "Mx"), errorText)
    myValue = ReadOptionalLoadDouble(data(rowIndex, 4), CellLocation(source, rowIndex, 4, "My"), errorText)

    ' Пропускаем только достоверно нулевые нагрузки. Сила и момент имеют
    ' разные размерности; reader не вводит инженерный порог их значимости.
    If Len(errorText) = 0 Then
        ConvertLoadsToInternal units, nValue, mxValue, myValue, errorText, source, rowIndex
        If Len(errorText) = 0 Then
            If nValue = 0# And mxValue = 0# And myValue = 0# Then Exit Sub
        End If
    End If

    If Len(Trim$(combinationID)) = 0 Then AppendError errorText, CellLocation(source, rowIndex, 1, "CombinationID") & " не заполнен"
    If Len(profileId) = 0 Then AppendError errorText, CellLocation(source, rowIndex, 5, "ProfileId") & " не заполнен"

    If Len(errorText) > 0 Then
        batch.AddInvalidCombination combinationID, profileId, comment, errorText, loadPathText, rowIndex - 1
    Else
        batch.AddCombination combinationID, nValue, mxValue, myValue, profileId, comment, loadPathText, rowIndex - 1
    End If
End Sub

' Проверяет, что пользователь действительно ничего не ввел в строке LC.
Private Function IsEmptyCombinationRow(ByRef data As Variant, ByVal rowIndex As Long) As Boolean
    Dim col As Long
    For col = 1 To UBound(data, 2)
        If Not IsBlankCell(data(rowIndex, col)) Then Exit Function
    Next col
    IsEmptyCombinationRow = True
End Function

' Читает нагрузку, где пустая ячейка означает 0.
' Это позволяет задавать одноосный изгиб как N + Mx при пустом My.
Private Function ReadOptionalLoadDouble(ByVal value As Variant, ByVal fieldName As String, ByRef errorText As String) As Double
    If IsError(value) Then
        AppendError errorText, fieldName & " содержит ошибку формулы Excel"
    ElseIf IsNull(value) Then
        AppendError errorText, fieldName & " содержит недопустимое значение Null"
    ElseIf IsBlankCell(value) Then
        Exit Function
    ElseIf Not IsNumeric(value) Then
        AppendError errorText, fieldName & " должен быть числом"
    Else
        On Error GoTo ConversionFailed
        ReadOptionalLoadDouble = CDbl(value)
    End If
    Exit Function
ConversionFailed:
    AppendError errorText, fieldName & " не помещается в допустимый числовой диапазон"
End Function

' Отличает пустую ячейку/пробелы от нуля, ошибки Excel и Null.
' Ошибочная ячейка не делает строку свободной для следующего сочетания.
Private Function IsBlankCell(ByVal value As Variant) As Boolean
    If IsEmpty(value) Then
        IsBlankCell = True
    ElseIf VarType(value) = vbString Then
        IsBlankCell = (Len(Trim$(value)) = 0)
    End If
End Function

' Читает текстовые поля строки, сохраняя ошибки формул как ошибки LC.
' Null недопустим, тогда как пустой optional-комментарий разрешен.
Private Function ReadCellText(ByVal value As Variant, ByVal fieldName As String, ByRef errorText As String) As String
    If IsError(value) Then
        AppendError errorText, fieldName & " содержит ошибку формулы Excel"
    ElseIf IsNull(value) Then
        AppendError errorText, fieldName & " содержит недопустимое значение Null"
    Else
        ReadCellText = Trim$(CStr(value))
    End If
End Function

' Переводит успешно прочитанные нагрузки через единый адаптер единиц.
' Переполнение конкретной компоненты является ошибкой строки, не всего пакета.
Private Sub ConvertLoadsToInternal(ByVal units As CUnitSystem, ByRef nValue As Double, _
        ByRef mxValue As Double, ByRef myValue As Double, ByRef errorText As String, _
        ByVal source As Object, ByVal rowIndex As Long)
    If units Is Nothing Then Exit Sub
    On Error GoTo ConversionFailed
    Dim fieldName As String
    fieldName = CellLocation(source, rowIndex, 2, "N"): nValue = units.InputForceToInternal(nValue)
    fieldName = CellLocation(source, rowIndex, 3, "Mx"): mxValue = units.InputMomentMxToInternal(mxValue)
    fieldName = CellLocation(source, rowIndex, 4, "My"): myValue = units.InputMomentMyToInternal(myValue)
    Exit Sub
ConversionFailed:
    AppendError errorText, fieldName & " не может быть переведено во внутренние единицы: числовой диапазон превышен"
End Sub

' Дополняет имя поля фактическим адресом исходной таблицы. Адрес формируется
' только Excel-reader-ом и передается как готовая причина; batch/writer не
' восстанавливают его по тексту и не читают исходные ячейки повторно.
Private Function CellLocation(ByVal source As Object, ByVal rowIndex As Long, _
        ByVal columnIndex As Long, ByVal fieldName As String) As String
    CellLocation = fieldName & ", " & source.Worksheet.Name & "!" & _
        source.Cells(rowIndex, columnIndex).Address(False, False)
End Function

' Добавляет текст ошибки в одну строку через точку с запятой.
Private Sub AppendError(ByRef errorText As String, ByVal item As String)
    If Len(errorText) > 0 Then errorText = errorText & "; "
    errorText = errorText & item
End Sub

