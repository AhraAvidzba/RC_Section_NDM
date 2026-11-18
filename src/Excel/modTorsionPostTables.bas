Attribute VB_Name = "modTorsionPostTables"
Option Explicit

' ==========================================================================
' Собственные таблицы T, MNT и компактных усилий стержней
' ==========================================================================
' Матрицы отделены от обычных таблиц НДМ и хранятся как необязательные части
' снимка. Вывод ограничен собственными именованными диапазонами на Results.
' Имена, порядок, значения и оформление таблиц NM не меняются.

' Собирает строки T, компактные усилия и отдельное дополнение верхней сводки.
Public Sub SaveTorsionPostMatrices(ByVal snapshot As CCalculationSnapshot, ByVal results As Collection, _
        ByVal batch As CBatchSectionCalculator, ByVal units As CUnitSystem)
    Dim data() As Variant, bars() As Variant, summary As Variant, values As Variant
    Dim row As Long, column As Long, barRows As Long, result As CTorsionPostResult
    Dim count As Long: count = results.Count
    If count < 1 Then count = 1
    ReDim data(1 To count, 1 To TORSION_POST_COLUMNS)
    For Each result In results
        values = result.Bars
        If IsArray(values) Then barRows = barRows + UBound(values, 1)
    Next result
    ReDim bars(1 To barRows + 1, 1 To TORSION_BAR_COLUMNS)
    bars(1, 1) = "LoadCase": bars(1, 2) = "ElementID": bars(1, 3) = "F_T, " & units.OutputForceUnit
    Dim barRow As Long: barRow = 1
    For Each result In results
        row = row + 1: summary = result.Summary
        For column = 1 To TORSION_POST_COLUMNS
            data(row, column) = TorsionSummaryOutputValue(summary(1, column), column, units)
        Next column
        values = result.Bars
        If IsArray(values) Then
            Dim i As Long
            For i = 1 To UBound(values, 1)
                barRow = barRow + 1
                bars(barRow, 1) = values(i, 1): bars(barRow, 2) = values(i, 2)
                bars(barRow, 3) = Abs(units.InternalForceToOutput(CDbl(values(i, 3)))) ' Усилие растяжения от T показано положительной величиной.
            Next i
        End If
    Next result
    snapshot.AddTable TORSION_POST_HEADERS_KEY, TorsionPostHeaders(units)
    snapshot.AddTable TORSION_POST_DATA_KEY, data
    snapshot.AddTable TORSION_BAR_DATA_KEY, bars
    snapshot.AddTable TORSION_MNT_DATA_KEY, TorsionMNTSummary(results, batch)
End Sub

' Преобразует числа только через общий CUnitSystem; пустые поля и текст
' не превращаются в нули. Запасы и угловые коэффициенты безразмерны.
Private Function TorsionSummaryOutputValue(ByVal value As Variant, ByVal column As Long, ByVal units As CUnitSystem) As Variant
    TorsionSummaryOutputValue = value
    If IsEmpty(value) Then Exit Function
    If column = 39 Then
        TorsionSummaryOutputValue = Replace$(Replace$(CStr(value), vbCrLf, "; "), vbLf, "; ") ' Подробная причина остаётся компактным пояснением Results; цепочки адресов сохраняются.
        Exit Function
    End If
    If Not IsNumeric(value) Then Exit Function
    Select Case column
        Case 3, 26: TorsionSummaryOutputValue = units.InternalMomentMagnitudeToOutput(CDbl(value))
        Case 4, 9, 10, 16: TorsionSummaryOutputValue = units.InternalAreaToOutput(CDbl(value))
        Case 5, 17, 23, 37: TorsionSummaryOutputValue = units.InternalLengthToOutput(CDbl(value))
        Case 7: TorsionSummaryOutputValue = Abs(units.InternalForceToOutput(CDbl(value))) / units.InternalLengthToOutput(1#)
        Case 8: TorsionSummaryOutputValue = Abs(units.InternalForceToOutput(CDbl(value)))
        Case 32: TorsionSummaryOutputValue = Sgn(CDbl(value)) * Abs(units.InternalForceToOutput(CDbl(value))) ' Положительное усилие блока NMT обозначает растяжение.
        Case 12, 13, 18, 19, 24, 25, 30, 31, 33, 34: TorsionSummaryOutputValue = units.InternalStressToOutput(CDbl(value))
    End Select
End Function

' Публикует только дополнительные таблицы. Любой отказ вывода возвращается
' предупреждением после сохранения НДМ и не меняет его статусы.
Public Function RestoreTorsionPostTables(ByVal workbook As Object, ByVal source As Object) As String
    On Error GoTo Failed
    Dim target As Object: Set target = OptionalPostOutputRange(workbook, "rngTorsionPostSummary")
    If target Is Nothing Then
        If source.HasTable(TORSION_POST_DATA_KEY) Then Err.Raise 5, , "Выходной диапазон rngTorsionPostSummary не найден или его ссылка повреждена. Восстановите именованную таблицу блока NMT на Results."
        Exit Function
    End If
    If Not source.HasTable(TORSION_POST_DATA_KEY) Then
        RestoreTorsionPostTables = ClearTorsionPostBodies(workbook)
        Exit Function
    End If
    Dim headers As Variant, data As Variant, bars As Variant, mnt As Variant
    headers = source.Table(TORSION_POST_HEADERS_KEY): data = source.Table(TORSION_POST_DATA_KEY)
    bars = source.Table(TORSION_BAR_DATA_KEY): mnt = source.Table(TORSION_MNT_DATA_KEY)
    ValidateSnapshotTableShape TORSION_POST_HEADERS_KEY, headers
    ValidateSnapshotTableShape TORSION_POST_DATA_KEY, data
    ValidateSnapshotTableShape TORSION_BAR_DATA_KEY, bars
    ValidateSnapshotTableShape TORSION_MNT_DATA_KEY, mnt
    ValidatePostOutput workbook, "rngTorsionPostSummary", 4 + UBound(data, 1), TORSION_POST_DISPLAY_COLUMNS
    ValidatePostOutput workbook, "rngTorsionBarForces", UBound(bars, 1), TORSION_BAR_COLUMNS
    ValidatePostOutput workbook, "rngTorsionMNTSummary", UBound(mnt, 1), 2
    WritePostOutput workbook, "rngTorsionPostSummary", headers, data
    WritePostOutput workbook, "rngTorsionBarForces", bars
    WritePostOutput workbook, "rngTorsionMNTSummary", mnt
    FormatTorsionPostTables workbook, UBound(data, 1), UBound(mnt, 1)
    Exit Function
Failed:
    RestoreTorsionPostTables = "Не удалось вывести дополнительный блок T (кручение) / NMT. " & WorkbookDialogText(Err.Description, workbook) & _
        vbCrLf & "Результаты НДМ этим выводом не изменялись."
End Function

' Создаёт пустой вывод при полной сборке без формирования или решения НДМ.
' Выполняется сборщиком после импорта модулей и создания имён T / NMT.
Public Sub InitializeTorsionPostOutput()
    Dim settings As New CSystemSettingsReader, units As New CUnitSystem
    settings.LoadFromWorkbook ThisWorkbook: units.LoadFromSettings settings
    Dim snapshot As New CCalculationSnapshot, results As New Collection, batch As New CBatchSectionCalculator
    SaveTorsionPostMatrices snapshot, results, batch, units
    Dim warning As String: warning = RestoreTorsionPostTables(ThisWorkbook, snapshot)
    If Len(warning) > 0 Then Err.Raise 5, , warning
End Sub

' Очищает только тело дополнительных таблиц; основная сводка и НДС не затрагиваются.
Public Function ClearTorsionPostBodies(ByVal workbook As Object) As String
    On Error GoTo Failed
    Dim name As Variant, target As Object
    For Each name In Array("rngTorsionPostSummary", "rngTorsionBarForces", "rngTorsionMNTSummary")
        Set target = OptionalPostOutputRange(workbook, CStr(name))
        If Not target Is Nothing Then
            If Not target.Worksheet.Parent Is workbook Then Err.Raise 5, , CStr(name) & " ссылается на другую книгу. Исправьте имя в текущей книге."
            If target.Worksheet.Name <> "Results" Then Err.Raise 5, , CStr(name) & " должен находиться на Results."
            Dim expectedColumns As Long
            Select Case CStr(name)
                Case "rngTorsionPostSummary": expectedColumns = TORSION_POST_DISPLAY_COLUMNS
                Case "rngTorsionBarForces": expectedColumns = TORSION_BAR_COLUMNS
                Case Else: expectedColumns = 2
            End Select
            ValidatePostOutput workbook, CStr(name), target.Rows.Count, expectedColumns
            Dim header As Long
            Select Case CStr(name)
                Case "rngTorsionPostSummary": header = 4
                Case "rngTorsionBarForces": header = 1
                Case Else: header = 6
            End Select
            If target.Rows.Count > header Then target.Cells(header + 1, 1).Resize(target.Rows.Count - header, target.Columns.Count).ClearContents
        End If
    Next name
    Exit Function
Failed:
    ClearTorsionPostBodies = WorkbookDialogText("Дополнительный вывод T (кручение) / NMT не очищен: " & Err.Description, workbook)
End Function

' Проверяет размеры и отсутствие пересечения с любыми существующими
' таблицами NM до первого изменения выходных ячеек T / NMT.
Private Sub ValidatePostOutput(ByVal workbook As Object, ByVal name As String, ByVal rows As Long, ByVal columns As Long)
    Dim target As Object: Set target = OptionalPostOutputRange(workbook, name)
    If target Is Nothing Then Err.Raise 5, , "Не найден выходной диапазон " & name & ". Восстановите его из актуальной сборки."
    If Not target.Worksheet.Parent Is workbook Then Err.Raise 5, , name & " ссылается на другую книгу. Исправьте имя в текущей книге."
    If target.Worksheet.Name <> "Results" Or target.Areas.Count <> 1 Or target.Columns.Count <> columns Then _
        Err.Raise 5, , name & " должен ссылаться на один прямоугольный блок Results шириной " & CStr(columns) & " столбцов."
    If target.Row + rows - 1 > target.Worksheet.Rows.Count Then Err.Raise 5, , name & " выходит за пределы листа."
    If target.Rows.Count > rows Then rows = target.Rows.Count
    Dim bounds As Variant, item As Variant, protected As Object, left As Long, right As Long, bottom As Long
    left = target.Column: right = left + columns - 1: bottom = target.Row + rows - 1
    bounds = Array(Array("rngBatchSummary", 42, 36), Array("rngStrengthSummaryAnchor", 0, 49), _
        Array("rngCrackSummaryAnchor", 0, 72), Array("rngStabilitySummaryAnchor", 0, STABILITY_SNAPSHOT_COLUMNS), _
        Array("rngNDMSectionGeometry", 0, 16), Array("rngNDMSectionProperties", 0, 6), Array("rngNDMMaterialDiagrams", 0, 11), _
        Array("rngNDMSectionContours", 0, 17), Array("rngNDMElementResults", 0, 9), Array("rngNDMElementGeometry", 0, 18), _
        Array("rngNDMSectionAnnotations", 0, 24), Array("rngNDMElementResultIndex", 0, 9))
    For Each item In bounds
        Set protected = workbook.Names.Item(CStr(item(0))).RefersToRange
        If protected.Worksheet.Name = "Results" Then
            Dim protectedTop As Long: protectedTop = protected.Row
            Dim strengthWriter As New CStrengthSummaryWriter, crackWriter As New CCrackSummaryWriter, stabilityWriter As New CStabilitySummaryWriter
            Select Case CStr(item(0))
                Case "rngStrengthSummaryAnchor": protectedTop = protectedTop - strengthWriter.HeaderRowsAboveAnchor - 1
                Case "rngCrackSummaryAnchor": protectedTop = protectedTop - crackWriter.HeaderRowsAboveAnchor - 1
                Case "rngStabilitySummaryAnchor": protectedTop = protectedTop - stabilityWriter.HeaderRowsAboveAnchor - 1
                Case "rngBatchSummary"
                Case Else: protectedTop = protectedTop - 2 ' Заголовки технических таблиц и общая полоса над ними также принадлежат НДМ.
            End Select
            Dim protectedBottom As Long: protectedBottom = protected.Row + CLng(item(1)) - 1
            If CLng(item(1)) = 0 Then protectedBottom = protected.Worksheet.Rows.Count ' Нижняя таблица может содержать любое число элементов.
            If left <= protected.Column + CLng(item(2)) - 1 And right >= protected.Column And _
                    target.Row <= protectedBottom And bottom >= protectedTop Then Err.Raise 5, , _
                "Диапазон " & name & " пересекает таблицу НДМ " & CStr(item(0)) & ". Переместите дополнительную таблицу, сохранив свободную полосу."
        End If
    Next item
End Sub

' Пишет матрицу в собственный диапазон и сохраняет занимаемый размер.
' Удаляются только значения и объединения внутри этой дополнительной таблицы.
Private Sub WritePostOutput(ByVal workbook As Object, ByVal name As String, ByRef first As Variant, Optional ByRef second As Variant)
    Dim target As Object: Set target = workbook.Names.Item(name).RefersToRange
    Dim rows As Long: rows = UBound(first, 1)
    If IsArray(second) Then rows = rows + UBound(second, 1)
    If target.Rows.Count > rows Then rows = target.Rows.Count
    Dim columns As Long: columns = UBound(first, 2)
    If name = "rngTorsionPostSummary" Then columns = TORSION_POST_DISPLAY_COLUMNS
    Dim sizeChanged As Boolean: sizeChanged = (target.Rows.Count <> rows Or target.Columns.Count <> columns)
    Set target = target.Cells(1, 1).Resize(rows, columns)
    target.UnMerge: target.ClearContents
    target.Cells(1, 1).Resize(UBound(first, 1), UBound(first, 2)).Value2 = first
    If IsArray(second) Then target.Cells(UBound(first, 1) + 1, 1).Resize(UBound(second, 1), UBound(second, 2)).Value2 = second
    If sizeChanged Then workbook.Names.Item(name).RefersTo = "='" & target.Worksheet.Name & "'!" & target.Address
End Sub

' Применяет стиль существующих сводок только к ячейкам T / NMT. Ширины
' столбцов и высоты строк не меняются ни здесь, ни при обновлении схемы.
Private Sub FormatTorsionPostTables(ByVal workbook As Object, ByVal dataRows As Long, ByVal mntRows As Long)
    Dim name As Variant, target As Object, column As Variant, row As Long
    For Each name In Array("rngTorsionPostSummary", "rngTorsionBarForces", "rngTorsionMNTSummary")
        Set target = workbook.Names.Item(CStr(name)).RefersToRange
        target.Font.Name = "Arial": target.Font.Size = 9
        target.HorizontalAlignment = -4108: target.VerticalAlignment = -4108
        target.WrapText = False: target.NumberFormat = "0.000"
        target.Interior.Pattern = -4142
        target.Borders.LineStyle = 1: target.Borders.Weight = 2: target.Borders.Color = 0
        Dim headerRows As Long
        Select Case CStr(name)
            Case "rngTorsionPostSummary": headerRows = 4
            Case "rngTorsionBarForces": headerRows = 1
            Case Else: headerRows = 6
        End Select
        Dim header As Object: Set header = target.Cells(1, 1).Resize(headerRows, target.Columns.Count)
        header.NumberFormat = "General": header.Font.Bold = True: header.Interior.Color = 14277081: header.WrapText = True
        If CStr(name) = "rngTorsionPostSummary" Then
            target.Cells(5, 1).Resize(target.Rows.Count - 4, 1).NumberFormat = "General" ' Числовые ID не получают дробную часть от формата результатов.
            For Each column In Array(4, 9, 10, 16)
                target.Cells(5, column).Resize(dataRows, 1).NumberFormat = "0.000000" ' Площадь ветви хомута в м2 остаётся видимой ненулевой величиной.
            Next column
            target.Cells(5, 7).Resize(dataRows, 1).NumberFormat = "0.000000E+00" ' Малый поток в tf/mm показывается без округления до нуля.
            target.Cells(1, 1).Resize(1, TORSION_POST_DISPLAY_COLUMNS).Merge
            For Each column In Array(Array(1, 2), Array(3, 15), Array(16, 21), Array(22, 28), Array(29, 36), Array(37, 37), Array(38, TORSION_POST_DISPLAY_COLUMNS))
                target.Cells(2, column(0)).Resize(1, column(1) - column(0) + 1).Merge
            Next column
            For row = 3 To target.Rows.Count
                With target.Cells(row, 39).Resize(1, TORSION_POST_DISPLAY_COLUMNS - 38)
                    .Merge
                    .HorizontalAlignment = -4131 ' Выравнивается вся объединённая область; часть столбца пересекает соседние объединения Excel.
                End With
            Next row
            target.Rows(3).Font.Bold = False
            target.Rows(3).HorizontalAlignment = -4131
            target.Rows(3).WrapText = False
            target.Rows(4).WrapText = False
            For row = 5 To dataRows + 4
                For Each column In Array(15, 21, 28, 36, 38)
                    If Len(CStr(target.Cells(row, column).Value2)) > 0 Then ApplyStatusFill target.Cells(row, column), CStr(target.Cells(row, column).Value2)
                Next column
            Next row
        ElseIf CStr(name) = "rngTorsionBarForces" Then
            If target.Rows.Count > 1 Then target.Cells(2, 1).Resize(target.Rows.Count - 1, 2).NumberFormat = "General"
        ElseIf CStr(name) = "rngTorsionMNTSummary" Then
            target.Cells(1, 1).Resize(1, 2).Merge: target.Cells(2, 1).Resize(1, 2).Merge
            target.Cells(3, 1).Resize(3, 1).Merge: target.Cells(3, 2).Resize(3, 1).Merge
            For row = 7 To mntRows
                If Len(CStr(target.Cells(row, 2).Value2)) > 0 Then ApplyStatusFill target.Cells(row, 2), CStr(target.Cells(row, 2).Value2)
            Next row
        End If
    Next name
End Sub

' Возвращает только собственное имя заданной книги; другие открытые книги
' с такими же именами диапазонов не являются источниками или получателями.
Private Function OptionalPostOutputRange(ByVal workbook As Object, ByVal name As String) As Object
    On Error Resume Next
    Set OptionalPostOutputRange = workbook.Names.Item(name).RefersToRange
    On Error GoTo 0
End Function

' Даёт четыре яруса шапки: название, группы, русские пояснения и короткие
' обозначения с единицами. Размеры диапазона и порядок данных сохраняются.
Public Function TorsionPostHeaders(ByVal units As CUnitSystem) As Variant
    Dim result(1 To 4, 1 To TORSION_POST_COLUMNS) As Variant, labels As Variant, notes As Variant, column As Long
    Dim sigma As String, deltaSigma As String, unit As String
    sigma = ChrW$(&H3C3): deltaSigma = ChrW$(&H394) & sigma ' Греческие знаки не зависят от ANSI-кодировки сохранённого проекта VBA.
    labels = Array("ID", "State", "T", "A_o", "u_o", "cot " & ChrW$(&H3B8), "q", "F_T", "A_s", _
        "A_s,req", "ID", deltaSigma & "_T,max", "R_s", "k", "статус", "A_sw", "s", sigma & "_sw", "R_sw", "k", "статус", _
        "метод", "t_eff", sigma & "_diag", "R_diag", "T_lim", "k", "статус", _
        "ID", sigma & "_NM", deltaSigma & "_T", "F_NMT", sigma & "_NMT", "R_s", "k", "статус", "b_sel", "статус", "ResultComment")
    notes = Array("Номер сочетания", "Сохранённое состояние НДМ", "Заданный крутящий момент", _
        "Площадь внутри оси наружного хомута", "Периметр оси наружного хомута", "Котангенс угла бетонных раскосов", _
        "Касательный поток", "Суммарная продольная тяга", "Площадь отобранной продольной арматуры", _
        "Требуемая площадь продольной арматуры", "Стержень с наибольшей добавкой напряжения", "Максимальная растягивающая добавка от T", _
        "Сопротивление продольной арматуры", "Запас продольной арматуры от T", "Статус проверки продольной арматуры", _
        "Площадь одной ветви хомута", "Шаг замкнутых хомутов", "Напряжение ветви хомута", "Сопротивление ветви хомута", _
        "Запас поперечной арматуры", "Статус проверки хомута", "Метод проверки бетона", "Эффективная толщина бетонной полосы", _
        "Напряжение бетонного раскоса", "Допустимое напряжение бетонного раскоса", "Предельный крутящий момент по бетону", _
        "Запас бетона", "Статус проверки бетона", "Определяющий стержень по сумме NMT", "Сохранённое напряжение НДМ", _
        "Растягивающая добавка от T", "Суммарное усилие стержня NMT", "Алгебраическая сумма напряжений NMT", _
        "Сопротивление продольной арматуры", "Запас по сумме напряжений NMT", "Статус проверки суммы NMT", _
        "Ширина полосы отбора продольных стержней", "Итоговый статус постпроверки", "Причины отказов и предупреждения T и NMT")
    result(1, 1) = "Приближённая постпроверка T (кручение) и продольной арматуры NMT"
    For column = 1 To TORSION_POST_COLUMNS
        result(3, column) = notes(column - 1)
        Select Case column
            Case 1: result(2, column) = "Сочетание"
            Case 3: result(2, column) = "T — продольная арматура"
            Case 16: result(2, column) = "T — хомут"
            Case 22: result(2, column) = "T — бетон"
            Case 29: result(2, column) = "NMT — сумма напряжений продольной арматуры"
            Case 37: result(2, column) = "Отбор"
            Case 38: result(2, column) = "Постпроверка"
        End Select
        unit = vbNullString
        Select Case column
            Case 3, 26: unit = units.OutputMomentUnit
            Case 4, 9, 10, 16: unit = units.OutputAreaUnit
            Case 5, 17, 23, 37: unit = units.OutputLengthUnit
            Case 7: unit = units.OutputForceUnit & "/" & units.OutputLengthUnit
            Case 8, 32: unit = units.OutputForceUnit
            Case 12, 13, 18, 19, 24, 25, 30, 31, 33, 34: unit = units.OutputStressUnit
        End Select
        result(4, column) = labels(column - 1)
        If Len(unit) > 0 Then result(4, column) = result(4, column) & ", " & unit
    Next column
    TorsionPostHeaders = result
End Function

' Строки совпадают с верхней сводкой НДМ. Запас MNT — минимум проверок
' блока NMT; отказы и коэффициенты НДМ в него не подмешиваются.
Private Function TorsionMNTSummary(ByVal results As Collection, ByVal batch As CBatchSectionCalculator) As Variant
    Dim count As Long: count = batch.Count
    If count < 30 Then count = 30
    Dim output() As Variant: ReDim output(1 To count + 7, 1 To 2)
    output(1, 1) = "MNT": output(2, 1) = "T + NM"
    output(3, 1) = "Коэффициент запаса": output(3, 2) = "Статус"
    output(6, 1) = "min запаса": output(6, 2) = "T + NM"
    Dim byID As Object: Set byID = CreateObject("Scripting.Dictionary"): byID.CompareMode = vbTextCompare
    Dim result As CTorsionPostResult, row As Variant, policy As New CResultStatusPolicy
    Dim headerStatus As String, missingSum As Boolean
    For Each result In results
        row = result.Summary
        If policy.StatusPriority(CStr(row(1, 38))) > policy.StatusPriority(headerStatus) Then headerStatus = CStr(row(1, 38))
        If CStr(row(1, 38)) = policy.NotApplicable Then missingSum = True
        If Len(CStr(row(1, 1))) > 0 Then byID(CStr(row(1, 1))) = row
    Next result
    Dim i As Long, minimum As Double, overall As Double, column As Variant
    For i = 1 To batch.Count
        output(i + 7, 2) = policy.NotApplicable
        If byID.Exists(batch.CombinationID(i)) Then
            row = byID(batch.CombinationID(i)): output(i + 7, 2) = row(1, 38)
            If CStr(row(1, 38)) = policy.OK Or CStr(row(1, 38)) = policy.Fail Then
                minimum = 0#
                For Each column In Array(20, 27, 35)
                    If Not IsEmpty(row(1, column)) Then
                        If CDbl(row(1, column)) > 0# Then
                            If minimum = 0# Or CDbl(row(1, column)) < minimum Then minimum = CDbl(row(1, column))
                        End If
                    End If
                Next column
                If minimum > 0# Then
                    output(i + 7, 1) = minimum
                    If overall = 0# Or minimum < overall Then overall = minimum
                End If
            End If
        End If
    Next i
    If overall > 0# Then output(7, 1) = overall
    If missingSum And (Len(headerStatus) = 0 Or headerStatus = policy.OK) Then headerStatus = policy.NotApplicable
    output(7, 2) = headerStatus
    TorsionMNTSummary = output
End Function
