Attribute VB_Name = "modIndexedExcelResults"
Option Explicit

' ==========================================================================
' Адресный индекс видимой таблицы результатов элементов Excel
' ==========================================================================
' Один якорь справа от аннотаций, относительные строки и фактически сохранённые ключи.
' Результаты элементов и индекс помечаются Locked; защиту листа включает только пользователь.
' Чтение не требует защиты: размеры, ключи и выбранные участки проверяются по фактическим данным.
Private Const INDEX_NAME As String = "rngNDMElementResultIndex"
Private Const INDEX_FORMAT As String = "NDMResultIndex03"
Private Const INDEX_COLUMNS As Long = 9
Private mTestRowsRead As Double, mTestBlocksRead As Long, mTestBootstrapRows As Double

' Находит единственный якорь, не создавая имён или скрытых копий для отдельных состояний.
Private Function ExistingIndexAnchor(ByVal workbook As Object) As Object
    Dim owner As Object
    On Error Resume Next
    Set owner = workbook.Names.Item(INDEX_NAME)
    On Error GoTo 0
    If Not owner Is Nothing Then Set ExistingIndexAnchor = owner.RefersToRange
End Function

' Создаёт свободную таблицу после аннотаций с двумя пустыми колонками между блоками.
' Повторный вызов использует тот же якорь; занятые соседние данные не очищает.
Public Function EnsureElementResultIndexAnchor(ByVal workbook As Object, ByVal rows As Long) As Object
    Dim anchor As Object: Set anchor = ExistingIndexAnchor(workbook)
    Dim elements As Object: Set elements = workbook.Names.Item("rngNDMElementResults").RefersToRange
    If anchor Is Nothing Then
        Dim annotations As Object: Set annotations = workbook.Names.Item("rngNDMSectionAnnotations").RefersToRange
        Set anchor = annotations.Cells(1, 1).Offset(0, 24 + 2)
        If anchor.Column + INDEX_COLUMNS - 1 > anchor.Worksheet.Columns.Count Or anchor.Row + rows - 1 > anchor.Worksheet.Rows.Count Then Err.Raise 5, , "Индекс результатов не помещается справа от аннотаций."
        If workbook.Application.WorksheetFunction.CountA(anchor.Offset(-1, 0).Resize(rows + 1, INDEX_COLUMNS)) <> 0 Then Err.Raise 5, , "Место для " & INDEX_NAME & " занято. Освободите область справа от аннотаций Results."
        workbook.Names.Add Name:=INDEX_NAME, RefersTo:="='" & Replace(anchor.Worksheet.Name, "'", "''") & "'!" & anchor.Address
    Else
        If anchor.Cells.CountLarge <> 1 Or Not anchor.Worksheet Is elements.Worksheet Then Err.Raise 5, , "Повреждён именованный якорь " & INDEX_NAME & "."
        If Len(CStr(anchor.Value2)) > 0 And CStr(anchor.Value2) <> "IndexKind" Then Err.Raise 5, , "В " & INDEX_NAME & " находятся посторонние данные."
        Dim oldRows As Long: oldRows = AnchoredRowCount(anchor)
        If oldRows < 1 Then oldRows = 1
        If rows > oldRows Then
            If workbook.Application.WorksheetFunction.CountA(anchor.Offset(oldRows, 0).Resize(rows - oldRows, INDEX_COLUMNS)) <> 0 Then Err.Raise 5, , "Под индексом находятся посторонние данные; расширение запрещено."
        End If
    End If
    Set EnsureElementResultIndexAnchor = anchor
End Function

' Проверяет пользовательскую защиту до записи, не включает и не снимает её.
' UserInterfaceOnly допускается только если пользователь сам разрешил запись макросам.
Public Sub PrepareIndexedResultsWrite(ByVal workbook As Object)
    Dim sheet As Object: Set sheet = workbook.Names.Item("rngNDMElementResults").RefersToRange.Worksheet
    If sheet.ProtectContents And Not sheet.ProtectionMode Then
        Err.Raise 5, "PrepareIndexedResultsWrite", "Лист Results защищён пользователем. Снимите защиту листа перед записью результатов; программа не меняет ваш пароль и режим защиты."
    End If
End Sub

' Сначала делает индекс недействительным; ошибка дальнейшей записи не оставляет прежний каталог.
' Пустая таблица после очистки не означает сохранённый расчёт из другого носителя.
Public Sub InvalidateElementResultIndex(ByVal workbook As Object)
    Dim anchor As Object: Set anchor = ExistingIndexAnchor(workbook)
    If anchor Is Nothing Then Set anchor = EnsureElementResultIndexAnchor(workbook, 2)
    PrepareIndexedResultsWrite workbook
    If Len(CStr(anchor.Value2)) > 0 And CStr(anchor.Value2) <> "IndexKind" Then Err.Raise 5, , "Нельзя очистить посторонние данные в " & INDEX_NAME & "."
    Dim rows As Long: rows = AnchoredRowCount(anchor)
    If rows > 1 Then anchor.Offset(1, 0).Resize(rows - 1, INDEX_COLUMNS).ClearContents
    anchor.Value2 = "IndexKind": anchor.Offset(1, 0).Value2 = "Invalid"
End Sub

' Готовит малую таблицу из общего индекса внутри измеряемого полного сохранения.
Public Function ElementResultIndexMatrix(ByVal index As CResultBlockIndex) As Variant
    Dim data() As Variant, names As Variant, column As Long, block As Long, sample As Variant
    names = Array("IndexKind", "RunID", "TableRows", "TableColumns", "FirstRow", "RowCount", "LoadCase", "ProfileId", "StateType")
    ReDim data(1 To index.Count + 2, 1 To INDEX_COLUMNS)
    For column = 1 To INDEX_COLUMNS: data(1, column) = names(column - 1): Next column
    data(2, 1) = INDEX_FORMAT: data(2, 2) = index.RunID: data(2, 3) = index.TotalRows: data(2, 4) = 9: data(2, 5) = 0: data(2, 6) = index.Count
    For block = 1 To index.Count
        sample = index.Representative(block)
        data(block + 2, 1) = "Block": data(block + 2, 2) = sample(1, 1)
        data(block + 2, 5) = index.FirstRow(block): data(block + 2, 6) = index.RowCount(block)
        For column = 2 To 4: data(block + 2, column + 5) = sample(1, column): Next column
    Next block
    ElementResultIndexMatrix = data
End Function

' Публикует индекс одним Value2 после успешной пакетной записи исходной таблицы.
' При ошибке индекс остаётся Invalid; ни старые, ни частичные результаты не выбираются по нему.
Public Sub PublishElementResultIndex(ByVal workbook As Object, ByRef data As Variant, ByVal formatter As CNDMResultsWriter)
    Dim anchor As Object: Set anchor = EnsureElementResultIndexAnchor(workbook, UBound(data, 1))
    Dim elements As Object: Set elements = workbook.Names.Item("rngNDMElementResults").RefersToRange
    Dim sheet As Object: Set sheet = elements.Worksheet
    PrepareIndexedResultsWrite workbook
    On Error GoTo Failed
    Dim column As Variant
    For Each column In Array(7, 8, 9): anchor.Resize(UBound(data, 1), INDEX_COLUMNS).Columns.Item(CLng(column)).NumberFormat = "@": Next column
    anchor.Resize(UBound(data, 1), INDEX_COLUMNS).Value2 = data
    formatter.FormatResultsBlock anchor, UBound(data, 1), INDEX_COLUMNS, "Индекс сохранённых состояний элементов", data
    anchor.Offset(1, 1).Resize(UBound(data, 1) - 1, 1).NumberFormat = "dd.mm.yy"
    elements.Resize(CLng(data(2, 3)), 9).Locked = True
    anchor.Offset(-1, 0).Resize(UBound(data, 1) + 1, INDEX_COLUMNS).Locked = True
    Exit Sub
Failed:
    Dim number As Long, reason As String: number = Err.Number: reason = Err.Description
    On Error Resume Next
    anchor.Offset(1, 0).Value2 = "Invalid"
    On Error GoTo 0
    Err.Raise number, "PublishElementResultIndex", reason
End Sub

' Перестраивает отсутствующий индекс книги без индекса один раз: четыре ключевые колонки и по одной строке на участок.
' После сохранения книги индекс читается непосредственно, повторного обхода ключей нет.
Private Function BootstrapElementResultIndex(ByVal workbook As Object, ByVal elements As Object, ByVal rows As Long) As CResultBlockIndex
    Dim header As Variant: header = elements.Resize(1, 9).Value2
    mTestRowsRead = mTestRowsRead + 1
    Dim index As New CResultBlockIndex, keys As Variant, runID As Double, block As Long, sample As Variant
    If rows > 1 Then
        keys = elements.Resize(rows, 4).Value2: runID = CDbl(keys(2, 1))
        mTestBootstrapRows = mTestBootstrapRows + rows: mTestRowsRead = mTestRowsRead + rows
        index.BuildKeyColumns keys, header, runID
        For block = 1 To index.Count
            sample = elements.Offset(index.FirstRow(block) - 1, 0).Resize(1, 9).Value2
            mTestRowsRead = mTestRowsRead + 1: index.SetRepresentative block, sample
        Next block
    Else
        index.Initialize header, 0#
    End If
    Dim values As Variant: values = ElementResultIndexMatrix(index)
    Dim anchor As Object: Set anchor = EnsureElementResultIndexAnchor(workbook, UBound(values, 1))
    Dim formatter As New CNDMResultsWriter
    PublishElementResultIndex workbook, values, formatter
    Set BootstrapElementResultIndex = index
End Function

' Читает компактный опубликованный индекс и проверяет размеры и ключи.
' Частичная публикация и противоречивые метаданные дают явную ошибку независимо от защиты листа.
Private Function ReadElementResultIndex(ByVal workbook As Object, ByVal elements As Object) As CResultBlockIndex
    Dim anchor As Object: Set anchor = ExistingIndexAnchor(workbook)
    Dim rows As Long
    If anchor Is Nothing Then
        rows = AnchoredRowCount(elements): If rows < 1 Then rows = 1
        Set ReadElementResultIndex = BootstrapElementResultIndex(workbook, elements, rows): Exit Function
    End If
    Dim metadata As Variant: metadata = anchor.Resize(2, INDEX_COLUMNS).Value2
    mTestRowsRead = mTestRowsRead + 2
    If Len(CStr(metadata(1, 1))) = 0 Then
        rows = AnchoredRowCount(elements): If rows < 1 Then rows = 1
        Set ReadElementResultIndex = BootstrapElementResultIndex(workbook, elements, rows): Exit Function
    End If
    If CStr(metadata(1, 1)) <> "IndexKind" Then Err.Raise 5, , "Повреждена шапка " & INDEX_NAME & "."
    If CStr(metadata(2, 1)) = "Invalid" Then
        If Len(CStr(elements.Offset(1, 0).Value2)) = 0 Then
            Set ReadElementResultIndex = BootstrapElementResultIndex(workbook, elements, 1): Exit Function
        End If
        Err.Raise 5, , "Индекс не опубликован; завершите расчёт либо очистите Results."
    End If
    If Len(CStr(metadata(2, 1))) = 0 Then
        rows = AnchoredRowCount(elements): If rows < 1 Then rows = 1
        Set ReadElementResultIndex = BootstrapElementResultIndex(workbook, elements, rows): Exit Function
    End If
    If CStr(metadata(2, 1)) <> INDEX_FORMAT Then Err.Raise 5, , "Неизвестный или повреждённый формат индекса результатов."
    rows = ExactIndexLong(metadata(2, 3))
    Dim count As Long: count = ExactIndexLong(metadata(2, 6))
    If rows < 1 Or ExactIndexLong(metadata(2, 4)) <> 9 Or count > rows - 1 Then Err.Raise 5, , "Повреждены размеры индекса."
    If CDbl(elements.Row) + rows - 1 > elements.Worksheet.Rows.Count Or CDbl(anchor.Row) + count + 1 > anchor.Worksheet.Rows.Count Then Err.Raise 5, , "Индекс выходит за границы листа."
    ' Граница и ключи проверяются без обхода всей таблицы и без обязательной защиты листа.
    If elements.Row + rows <= elements.Worksheet.Rows.Count Then
        If Len(CStr(elements.Offset(rows, 0).Value2)) > 0 Then Err.Raise 5, , "За сохранённой границей результатов появились строки без индекса."
        mTestRowsRead = mTestRowsRead + 1
    End If
    Dim header As Variant: header = elements.Resize(1, 9).Value2: mTestRowsRead = mTestRowsRead + 1
    Dim index As New CResultBlockIndex: index.Initialize header, CDbl(metadata(2, 2))
    If count > 0 Then
        Dim values As Variant: values = anchor.Offset(2, 0).Resize(count, INDEX_COLUMNS).Value2
        mTestRowsRead = mTestRowsRead + count
        Dim row As Long, column As Long, sample(1 To 1, 1 To 9) As Variant
        For row = 1 To count
            If CStr(values(row, 1)) <> "Block" Then Err.Raise 5, , "Повреждена запись индекса состояний."
            sample(1, 1) = values(row, 2)
            For column = 2 To 4: sample(1, column) = values(row, column + 5): Next column
            index.Append ExactIndexLong(values(row, 5)), ExactIndexLong(values(row, 6)), sample
        Next row
    End If
    If index.TotalRows <> rows Then Err.Raise 5, , "Индекс не покрывает всю таблицу результатов."
    Set ReadElementResultIndex = index
End Function

' Принимает только целые относительные строки без округления повреждённых значений.
Private Function ExactIndexLong(ByVal value As Variant) As Long
    If Not IsNumeric(value) Then Err.Raise 5, , "В индексе требуется целый номер строки."
    If CDbl(value) < 0# Or CDbl(value) > 2147483647# Then Err.Raise 5, , "Номер строки индекса вне диапазона Long."
    ExactIndexLong = CLng(value)
    If CDbl(ExactIndexLong) <> CDbl(value) Then Err.Raise 5, , "Дробный номер строки в индексе результатов."
End Function

' Читает только диапазоны выбранных участков; список состояний берётся из метаданных.
Public Function ReadIndexedExcelElementResults(ByVal workbook As Object, ByVal combinationID As String, ByVal stateType As String, ByVal indexOnly As Boolean, Optional ByVal profileID As String = "") As Variant
    Dim elements As Object: Set elements = workbook.Names.Item("rngNDMElementResults").RefersToRange
    mTestRowsRead = mTestRowsRead + 1
    If Len(CStr(elements.Cells(1, 1).Value2)) = 0 Then Exit Function
    Dim index As CResultBlockIndex: Set index = ReadElementResultIndex(workbook, elements)
    If indexOnly Then ReadIndexedExcelElementResults = index.StateRows(combinationID, stateType, profileID): Exit Function
    Dim blocks As Collection: Set blocks = index.SelectBlocks(combinationID, stateType, profileID)
    Dim block As Variant, rows As Long, used As Long, row As Long, column As Long, data As Variant, sample As Variant, result() As Variant, header As Variant
    For Each block In blocks: rows = rows + index.RowCount(CLng(block)): Next block
    ReDim result(1 To rows + 1, 1 To 9): header = index.Header
    For column = 1 To 9: result(1, column) = header(1, column): Next column
    used = 1
    For Each block In blocks
        data = elements.Offset(index.FirstRow(CLng(block)) - 1, 0).Resize(index.RowCount(CLng(block)), 9).Value2
        mTestRowsRead = mTestRowsRead + index.RowCount(CLng(block)): mTestBlocksRead = mTestBlocksRead + 1
        sample = index.Representative(CLng(block))
        For row = 1 To UBound(data, 1)
            For column = 1 To 4
                If StrComp(CStr(data(row, column)), CStr(sample(1, column)), vbBinaryCompare) <> 0 Then Err.Raise 5, , "Результаты изменены без обновления индекса: " & elements.Offset(index.FirstRow(CLng(block)) + row - 2, column - 1).Address(External:=True)
            Next column
            used = used + 1
            For column = 1 To 9: result(used, column) = data(row, column): Next column
        Next row
    Next block
    ReadIndexedExcelElementResults = result
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Реальные строки Value2 и блоки позволяют отличить адресный доступ от полного обхода.
Public Function TestIndexedExcelReadCounters(Optional ByVal reset As Boolean = False) As String
    TestIndexedExcelReadCounters = CStr(mTestRowsRead) & "|" & CStr(mTestBlocksRead) & "|" & CStr(mTestBootstrapRows)
    If reset Then mTestRowsRead = 0#: mTestBlocksRead = 0: mTestBootstrapRows = 0#
End Function
