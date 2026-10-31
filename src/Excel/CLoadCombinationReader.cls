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
' Переносит строки rngLoadCombinations в CBatchSectionCalculator. Пустые
' компоненты N, Mx и My считаются нулём; полностью нулевая нагрузка пропускается,
' а неполные/некорректные строки сохраняются как InvalidInput, чтобы пользователь
' увидел их в отчете. Повторяющиеся заполненные ID прерывают запуск до добавления строк.
'
' Если передан CUnitSystem, усилия сразу переводятся из INPUT-единиц и
' пользовательской системы знаков во внутреннюю систему расчетного ядра.

' Загружает сочетания из именованного диапазона книги.
Public Sub LoadFromWorkbook(ByVal workbook As Object, ByVal batch As CBatchSectionCalculator, Optional ByVal units As CUnitSystem = Nothing)
    If workbook Is Nothing Then Err.Raise vbObjectError + 3950, "CLoadCombinationReader", "Книга Excel не передана."
    If batch Is Nothing Then Err.Raise vbObjectError + 3951, "CLoadCombinationReader", "Пакетный расчетчик не передан."
    LoadFromRange workbook.Names.Item("rngLoadCombinations").RefersToRange, batch, units
End Sub

' Читает сочетания из переданного диапазона. Общий способ используется
' обычной книгой и тестами без ручного запуска её интерфейса.
Public Sub LoadFromRange(ByVal source As Object, ByVal batch As CBatchSectionCalculator, Optional ByVal units As CUnitSystem = Nothing)
    If batch Is Nothing Then Err.Raise vbObjectError + 3953, "CLoadCombinationReader", "Пакетный расчетчик не передан."
    RequireSourceTable source

    Dim data As Variant
    data = source.Value2
    RequireUniqueIDs data, source

    Dim rowIndex As Long
    For rowIndex = 2 To UBound(data, 1) ' Пустой номер отключает строку, сохраняя введённые нагрузки и её исходное место.
        If Not IsBlankCell(data(rowIndex, 1)) Then AddRow data, batch, rowIndex, units, source
    Next rowIndex
End Sub

' Проверяет заполненные ID до чтения геометрии и запуска расчёта; пустой ID отключает строку.
' Регистр и крайние пробелы не различаются, как при выборе сохранённого сочетания.
Public Sub ValidateUniqueIDs(ByVal source As Object)
    RequireSourceTable source
    Dim data As Variant: data = source.Value2
    RequireUniqueIDs data, source
End Sub

' Проверяет форму именованной таблицы до блочного чтения её значений.
Private Sub RequireSourceTable(ByVal source As Object)
    If source Is Nothing Then Err.Raise vbObjectError + 3952, "CLoadCombinationReader", "Диапазон сочетаний нагрузок не найден."
    If source.Areas.Count <> 1 Then Err.Raise vbObjectError + 3956, "CLoadCombinationReader", _
        "Именованный диапазон rngLoadCombinations должен ссылаться на единую прямоугольную таблицу, а не на несколько отдельных областей." & vbCrLf & vbCrLf & _
        "Сейчас указан диапазон: " & source.Worksheet.Name & "!" & source.Address(False, False) & "." & vbCrLf & vbCrLf & _
        "Сделайте единую прямоугольную таблицу со строкой заголовков и всеми сочетаниями. В диспетчере имён Excel укажите её целиком для rngLoadCombinations."
    If source.Rows.Count < 2 Then Err.Raise vbObjectError + 3954, "CLoadCombinationReader", _
        "rngLoadCombinations должен содержать строку заголовков и минимум одну строку сочетания."
    If source.Columns.Count < 5 Then Err.Raise vbObjectError + 3955, "CLoadCombinationReader", _
        "rngLoadCombinations должен содержать минимум столбцы CombinationID, N, Mx, My и ProfileId."

End Sub

' Просматривает только ID; нагрузки и остальные поля отключённых строк здесь не читаются.
' Первый повтор прерывает чтение целиком, прежде чем в batch попадёт хотя бы одна строка.
Private Sub RequireUniqueIDs(ByRef data As Variant, ByVal source As Object)
    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare
    Dim rowIndex As Long, firstRow As Long, identifier As String
    For rowIndex = 2 To UBound(data, 1)
        If IsBlankCell(data(rowIndex, 1)) Or IsError(data(rowIndex, 1)) Or IsNull(data(rowIndex, 1)) Then GoTo NextIdentifier
        identifier = Trim$(CStr(data(rowIndex, 1))) ' Тот же ключ, который AddRow передаёт в пакетный расчёт.
        If seen.Exists(identifier) Then
            firstRow = CLng(seen(identifier))
            Err.Raise vbObjectError + 3957, "CLoadCombinationReader", _
                "Расчёт не запущен: заполненные ID сочетаний должны быть разными." & vbCrLf & vbCrLf & _
                "Повторяющийся CombinationID: «" & identifier & "»." & vbCrLf & vbCrLf & _
                "Адреса ячеек:" & vbCrLf & _
                CellLocation(source, firstRow, 1, "CombinationID (rngLoadCombinations)") & vbCrLf & _
                CellLocation(source, rowIndex, 1, "CombinationID (rngLoadCombinations)") & vbCrLf & vbCrLf & _
                "Задайте каждому сочетанию индивидуальный ID. Чтобы отключить строку, очистите её ID."
        End If
        seen.Add identifier, rowIndex
NextIdentifier:
    Next rowIndex
End Sub

' Читает одну строку таблицы нагрузок и добавляет сочетание в CBatchSectionCalculator.
' Ошибки ввода не прерывают весь расчет: строка попадает в отчет как InvalidInput,
' остальные сочетания продолжают считаться.
Private Sub AddRow(ByRef data As Variant, ByVal batch As CBatchSectionCalculator, ByVal rowIndex As Long, _
        ByVal units As CUnitSystem, ByVal source As Object)
    Dim combinationID As String
    Dim profileId As String
    Dim loadPathText As String
    Dim comment As String
    Dim errorText As String, hasError As Boolean
    Dim pathLocation As String
    pathLocation = "LoadPath"
    If UBound(data, 2) >= 6 Then pathLocation = CellLocation(source, rowIndex, 6, "LoadPath")
    combinationID = ReadCellText(data(rowIndex, 1), CellLocation(source, rowIndex, 1, "CombinationID"), errorText, hasError)
    Dim nValue As Double, mxValue As Double, myValue As Double
    nValue = ReadOptionalLoadDouble(data(rowIndex, 2), CellLocation(source, rowIndex, 2, "N"), errorText, hasError)
    mxValue = ReadOptionalLoadDouble(data(rowIndex, 3), CellLocation(source, rowIndex, 3, "Mx"), errorText, hasError)
    myValue = ReadOptionalLoadDouble(data(rowIndex, 4), CellLocation(source, rowIndex, 4, "My"), errorText, hasError)
    If Not hasError Then ' Точно нулевой LC не использует профиль, путь и комментарий; ненулевые малые нагрузки сохраняются.
        ConvertLoadsToInternal units, nValue, mxValue, myValue, errorText, hasError, source, rowIndex
        If Not hasError Then
            If nValue = 0# And mxValue = 0# And myValue = 0# Then Exit Sub
        End If
    End If
    profileId = ReadCellText(data(rowIndex, 5), CellLocation(source, rowIndex, 5, "ProfileId"), errorText, hasError)
    If UBound(data, 2) >= 7 Then
        loadPathText = ReadCellText(data(rowIndex, 6), pathLocation, errorText, hasError)
        comment = ReadCellText(data(rowIndex, 7), CellLocation(source, rowIndex, 7, "Комментарий"), errorText, hasError)
    ElseIf UBound(data, 2) = 6 Then
        If StrComp(CStr(data(1, 6)), "LoadPath", vbTextCompare) = 0 Then
            loadPathText = ReadCellText(data(rowIndex, 6), pathLocation, errorText, hasError)
        Else
            comment = ReadCellText(data(rowIndex, 6), CellLocation(source, rowIndex, 6, "Комментарий"), errorText, hasError)
        End If
    End If
    Dim pathDescriptor As CLoadPathDescriptor
    Set pathDescriptor = New CLoadPathDescriptor
    If Len(pathDescriptor.NormalizeKey(loadPathText)) = 0 Then _
        AppendError errorText, hasError, pathLocation & ": выберите Auto, lambda*Mx, lambda*My, lambda*Mxy, lambda*N или lambda*NMxy"

    If Len(Trim$(combinationID)) = 0 Then AppendError errorText, hasError, CellLocation(source, rowIndex, 1, "CombinationID") & " не заполнен"
    If Len(profileId) = 0 Then AppendError errorText, hasError, CellLocation(source, rowIndex, 5, "ProfileId") & " не заполнен"

    If hasError Then
        batch.AddInvalidCombination combinationID, profileId, comment, errorText, loadPathText, rowIndex - 1, RowLocation(source, rowIndex)
    Else
        batch.AddCombination combinationID, nValue, mxValue, myValue, profileId, comment, loadPathText, rowIndex - 1, RowLocation(source, rowIndex)
    End If
End Sub

' Читает нагрузку, где пустая ячейка означает 0.
' Это позволяет задавать одноосный изгиб как N + Mx при пустом My.
Private Function ReadOptionalLoadDouble(ByVal value As Variant, ByVal fieldName As String, ByRef errorText As String, ByRef hasError As Boolean) As Double
    Dim parsed As Double
    If IsError(value) Then
        AppendError errorText, hasError, fieldName & " содержит ошибку формулы Excel"
    ElseIf IsNull(value) Then
        AppendError errorText, hasError, fieldName & " содержит недопустимое значение Null"
    ElseIf IsBlankCell(value) Then
        Exit Function
    ElseIf Not TryReadStrictDouble(value, parsed) Then
        AppendError errorText, hasError, fieldName & " должен быть числом без разделителей разрядов"
    Else
        On Error GoTo ConversionFailed
        ReadOptionalLoadDouble = parsed
    End If
    Exit Function
ConversionFailed:
    AppendError errorText, hasError, fieldName & " не помещается в допустимый числовой диапазон"
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
Private Function ReadCellText(ByVal value As Variant, ByVal fieldName As String, ByRef errorText As String, ByRef hasError As Boolean) As String
    If IsError(value) Then
        AppendError errorText, hasError, fieldName & " содержит ошибку формулы Excel"
    ElseIf IsNull(value) Then
        AppendError errorText, hasError, fieldName & " содержит недопустимое значение Null"
    Else
        ReadCellText = Trim$(CStr(value))
    End If
End Function

' Переводит успешно прочитанные нагрузки через единый адаптер единиц.
' Переполнение конкретной компоненты является ошибкой строки, не всего пакета.
Private Sub ConvertLoadsToInternal(ByVal units As CUnitSystem, ByRef nValue As Double, _
        ByRef mxValue As Double, ByRef myValue As Double, ByRef errorText As String, ByRef hasError As Boolean, _
        ByVal source As Object, ByVal rowIndex As Long)
    If units Is Nothing Then Exit Sub
    On Error GoTo ConversionFailed
    Dim fieldName As String
    fieldName = CellLocation(source, rowIndex, 2, "N"): nValue = units.InputForceToInternal(nValue)
    fieldName = CellLocation(source, rowIndex, 3, "Mx"): mxValue = units.InputMomentMxToInternal(mxValue)
    fieldName = CellLocation(source, rowIndex, 4, "My"): myValue = units.InputMomentMyToInternal(myValue)
    Exit Sub
ConversionFailed:
    AppendError errorText, hasError, fieldName & " не может быть переведено во внутренние единицы: числовой диапазон превышен"
End Sub

' Передаёт адрес прочитанной строки вместе с сочетанием, чтобы сообщения
' расчётных результатов оставались адресными без повторного чтения Excel.
Private Function RowLocation(ByVal source As Object, ByVal rowIndex As Long) As String
    RowLocation = source.Worksheet.Name & "!" & source.Rows(rowIndex).Address(False, False)
End Function

' Добавляет к названию параметра фактический адрес его ячейки.
' Пакетный расчёт получает готовую причину без поиска адреса в тексте ошибки.
Private Function CellLocation(ByVal source As Object, ByVal rowIndex As Long, _
        ByVal columnIndex As Long, ByVal fieldName As String) As String
    CellLocation = fieldName & ", " & source.Worksheet.Name & "!" & _
        source.Cells(rowIndex, columnIndex).Address(False, False)
End Function

' Добавляет текст ошибки в одну строку через точку с запятой.
Private Sub AppendError(ByRef errorText As String, ByRef hasError As Boolean, ByVal item As String)
    hasError = True ' Факт ошибки сохраняется независимо от оформления или наличия пояснения.
    If Len(errorText) > 0 Then errorText = errorText & "; "
    errorText = errorText & item
End Sub
