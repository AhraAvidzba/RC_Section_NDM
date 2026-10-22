Attribute VB_Name = "modSectionSnapshot"
Option Explicit

' ==========================================================================
' Чтение сохранённого сечения и таблиц Results
' ==========================================================================
' Восстанавливает геометрию в CSectionModel независимо от её источника.
' Расчёт, схема и экспорт используют одни данные; модуль не обращается
' к AutoCAD, не строит сетку и не решает систему равновесия.

Private mSnapshotUnits As CUnitSystem ' Преобразование единиц, записанных в Results, независимо от текущего Config.
Private mSnapshotReadTables As Object ' Таблицы одной операции чтения; освобождаются после восстановления модели или ошибки.
Private mSnapshotReadWorkbook As Object ' Книга, которой принадлежат временно прочитанные таблицы.
Private mSnapshotTableReadCount As Long ' ДЛЯ ТЕСТОВ: число загрузок целых таблиц; чтения при поиске границ сюда не входят.

' Восстанавливает расчётную модель из обязательных координат, площадей,
' инерций и исходных контуров Results в мм, мм2 и мм4. При includeOptionalData=True
' дополнительно читает JSON происхождения и RecipeTrace; для расчёта они не требуются.
' Восстановлением аннотаций управляет вызывающий ReadFormedSectionFromResults.
Public Function ReadSectionGeometryFromResults(ByVal workbook As Object, Optional ByVal sourceType As String = "Results", _
        Optional ByVal includeOptionalData As Boolean = False) As CSectionModel
    Dim anchor As Object: Set anchor = workbook.Names.Item("rngNDMSectionGeometry").RefersToRange
    Dim data As Variant: data = ReadAnchoredResultTable(workbook, "rngNDMSectionGeometry")
    If Not HasResultTableRows(data) Then Err.Raise vbObjectError + 4350, "ReadSectionGeometryFromResults", _
        "На листе Results нет сформированной геометрии. Нажмите «Сформировать сечение»."
    If UBound(data, 2) < 14 Then Err.Raise vbObjectError + 4355, "ReadSectionGeometryFromResults", _
        "Неверная структура rngNDMSectionGeometry. Нажмите «Сформировать сечение»."
    Dim colID As Long: colID = ResultColumn(data, "ElementID")
    Dim colType As Long: colType = ResultColumn(data, "MaterialType")
    Dim colX As Long: colX = ResultColumn(data, "X")
    Dim colY As Long: colY = ResultColumn(data, "Y")
    Dim colArea As Long: colArea = ResultColumn(data, "Area")
    Dim colShape As Long: colShape = ResultColumn(data, "GeometryInterpretationStatus")
    Dim colWidth As Long: colWidth = ResultColumn(data, "Width")
    Dim colHeight As Long: colHeight = ResultColumn(data, "Height")
    Dim colDiameter As Long: colDiameter = ResultColumn(data, "Diameter")
    Dim colRotation As Long: colRotation = ResultColumn(data, "Rotation")
    Dim colLocalIx As Long: colLocalIx = ResultColumn(data, "LocalIx")
    Dim colLocalIy As Long: colLocalIy = ResultColumn(data, "LocalIy")
    Dim colLocalIxy As Long: colLocalIxy = ResultColumn(data, "LocalIxy")
    Dim colComment As Long
    Dim lengthUnit As String: lengthUnit = ResultHeaderUnit(data, colX, "mm")
    Dim areaUnit As String: areaUnit = ResultHeaderUnit(data, colArea, "mm2")
    Dim fourthUnit As String: fourthUnit = ResultHeaderUnit(data, colLocalIx, lengthUnit & "4")
    Dim model As CSectionModel: Set model = New CSectionModel
    model.SourceType = sourceType
    model.ConcreteMeshSource = ReadLoadCasePropertyText(workbook, "ALL", "Geometry.ConcreteMeshSource")
    Dim row As Long, kind As String, id As String, info As Object, sourceName As String, handle As String
    Dim comment As String, steelClass As String, materialID As Long, extraColumn As Long
    If includeOptionalData Then
        extraColumn = ResultColumn(data, "ExtraData", False)
        colComment = ResultColumn(data, "Comment", False)
        Set info = ReadSectionOptionalJson(vbNullString)
    End If
    For row = 2 To UBound(data, 1)
        id = ReadSavedAnnotationText(data, anchor, row, colID)
        If Len(Trim$(id)) > 0 Then
            If IsError(data(row, 1)) Or IsNull(data(row, 1)) Then GoTo InvalidElementID
            If CStr(data(row, 1)) <> CStr(data(2, 1)) Then GoTo InvalidElementID
            kind = ReadSavedAnnotationText(data, anchor, row, colType)
            sourceName = vbNullString: handle = vbNullString: comment = vbNullString: steelClass = vbNullString: materialID = 1
            If includeOptionalData Then
                If extraColumn > 0 Then Set info = ReadSnapshotOptionalRecord(data, anchor, row, extraColumn)
                sourceName = SectionOptionalText(info, "SourceName"): handle = SectionOptionalText(info, "SourceHandle")
                steelClass = SectionOptionalText(info, "SteelClass")
                If info.Exists("MaterialID") Then materialID = CLng(SectionOptionalText(info, "MaterialID"))
                If colComment > 0 Then comment = ReadSavedAnnotationText(data, anchor, row, colComment)
            End If
            If StrComp(kind, "Concrete", vbTextCompare) = 0 Then
                Dim localIx As Double, localIy As Double, localIxy As Double
                localIx = ReadPositiveGeometryNumber(data, anchor, row, colLocalIx)
                localIy = ReadPositiveGeometryNumber(data, anchor, row, colLocalIy)
                localIxy = ReadGeometryNumber(data, anchor, row, colLocalIxy)
                If Abs(localIxy) / Sqr(localIx) / Sqr(localIy) >= 1# Then Err.Raise vbObjectError + 4355, _
                    "ReadSectionGeometryFromResults", "Несогласованные собственные моменты инерции бетонного элемента: " & _
                    anchor.Parent.Name & "!" & anchor.Offset(row - 1, colLocalIxy - 1).Address(False, False) & ". Нажмите «Сформировать сечение»."
                model.AddConcreteElement OutputLengthToInternalByUnit(ReadGeometryNumber(data, anchor, row, colX), lengthUnit), _
                    OutputLengthToInternalByUnit(ReadGeometryNumber(data, anchor, row, colY), lengthUnit), _
                    OutputAreaToInternalByUnit(ReadPositiveGeometryNumber(data, anchor, row, colArea), areaUnit), materialID, sourceName, handle, _
                    ReadSavedAnnotationText(data, anchor, row, colShape), _
                    OutputLengthToInternalByUnit(ReadPositiveGeometryNumber(data, anchor, row, colWidth), lengthUnit), _
                    OutputLengthToInternalByUnit(ReadPositiveGeometryNumber(data, anchor, row, colHeight), lengthUnit), _
                    ReadGeometryNumber(data, anchor, row, colRotation), comment, _
                    OutputFourthPowerLengthToInternalByUnit(localIx, fourthUnit), _
                    OutputFourthPowerLengthToInternalByUnit(localIy, fourthUnit), _
                    OutputFourthPowerLengthToInternalByUnit(localIxy, fourthUnit), True ' Это готовая расчётная граница ячейки. Угол уже определён при формировании; повторное усреднение направлений не требуется.
                If model.ConcreteID(model.ConcreteCount) <> id Then GoTo InvalidElementID
            ElseIf StrComp(kind, "Rebar", vbTextCompare) = 0 Then
                model.AddRebarElement OutputLengthToInternalByUnit(ReadGeometryNumber(data, anchor, row, colX), lengthUnit), _
                    OutputLengthToInternalByUnit(ReadGeometryNumber(data, anchor, row, colY), lengthUnit), _
                    OutputLengthToInternalByUnit(ReadPositiveGeometryNumber(data, anchor, row, colDiameter), lengthUnit), _
                    OutputAreaToInternalByUnit(ReadPositiveGeometryNumber(data, anchor, row, colArea), areaUnit), steelClass, materialID, sourceName, handle, comment
                If model.RebarID(model.RebarCount) <> id Then GoTo InvalidElementID
                If includeOptionalData Then model.RestoreRebarRecipeTrace model.RebarCount, SectionOptionalText(info, "RecipeTrace")
            Else
                GoTo InvalidElementID
            End If
        End If
    Next row
    If model.ConcreteCount <= 0 Then Err.Raise vbObjectError + 4351, "ReadSectionGeometryFromResults", "В Results нет бетонных элементов."
    ReadSavedSectionContours workbook, model.Contours
    Set ReadSectionGeometryFromResults = model
    Exit Function
InvalidElementID:
    Err.Raise vbObjectError + 4381, "ReadSectionGeometryFromResults", _
        "Нарушены тип, порядок или имя элемента: " & anchor.Parent.Name & "!" & _
        anchor.Offset(row - 1, colID - 1).Address(False, False) & ". Нажмите «Сформировать сечение»."
End Function

' Разбирает JSON из уже прочитанного массива. Адрес ячейки запрашивается
' только после ошибки, а не для каждого элемента полного восстановления.
Private Function ReadSnapshotOptionalRecord(ByRef data As Variant, ByVal anchor As Object, ByVal row As Long, ByVal column As Long) As Object
    Dim text As String: text = ReadSavedAnnotationText(data, anchor, row, column)
    On Error GoTo InvalidJson
    Set ReadSnapshotOptionalRecord = ReadSectionOptionalJson(text)
    Exit Function
InvalidJson:
    Set ReadSnapshotOptionalRecord = ReadSectionOptionalJson(text, _
        anchor.Parent.Name & "!" & anchor.Offset(row - 1, column - 1).Address(False, False))
End Function

' Число оформления читается без обращения к отдельной ячейке Excel.
' При неверном тексте добавляется адрес содержащего его JSON.
Private Function ReadSnapshotOptionalNumber(ByVal text As String, ByVal anchor As Object, ByVal row As Long, ByVal column As Long) As Double
    On Error GoTo InvalidNumber
    ReadSnapshotOptionalNumber = ReadSectionOptionalNumber(text, vbNullString)
    Exit Function
InvalidNumber:
    ReadSnapshotOptionalNumber = ReadSectionOptionalNumber(text, _
        anchor.Parent.Name & "!" & anchor.Offset(row - 1, column - 1).Address(False, False))
End Function

' Проверяет положительную физическую величину. Адрес запрашивается у Excel
' только при ошибке, поэтому чтение сетки остаётся блочным.
Private Function ReadPositiveGeometryNumber(ByRef data As Variant, ByVal anchor As Object, ByVal row As Long, ByVal column As Long) As Double
    ReadPositiveGeometryNumber = ReadGeometryNumber(data, anchor, row, column)
    If ReadPositiveGeometryNumber <= 0# Then Err.Raise vbObjectError + 4355, "ReadSectionGeometryFromResults", _
        "Значение " & CStr(data(1, column)) & " должно быть положительным: " & anchor.Parent.Name & "!" & _
        anchor.Offset(row - 1, column - 1).Address(False, False) & ". Нажмите «Сформировать сечение»."
End Function

' Читает единственный 17-колоночный формат достоверных контуров.
' Отсутствующая таблица или повреждение явно диагностируются.
' Единицы и SectionXY принадлежат снимку, а не текущему Config.
Public Function ReadSavedSectionContours(ByVal workbook As Object, _
        Optional ByVal contours As CSectionContours = Nothing) As CSectionContours
    Dim anchor As Object, data As Variant, headers As Variant, i As Long, row As Long
    On Error GoTo MissingAnchor
    Set anchor = workbook.Names.Item("rngNDMSectionContours").RefersToRange
    On Error GoTo 0
    data = ReadAnchoredResultTable(workbook, "rngNDMSectionContours")
    headers = Array("RunID", "LoopID", "SegmentID", "Sequence", "LoopRole", "SegmentType", _
        "StartX", "StartY", "EndX", "EndY", "CenterX", "CenterY", "Radius", "SweepAngle, rad", "SourceID", "Comment", "LengthUnit (SectionXY)")
    If Not IsArray(data) Then GoTo InvalidHeader
    If UBound(data, 2) <> 17 Then GoTo InvalidHeader
    For i = 0 To UBound(headers)
        If IsError(data(1, i + 1)) Then GoTo InvalidHeader
        If CStr(data(1, i + 1)) <> CStr(headers(i)) Then GoTo InvalidHeader
    Next i
    Dim restored As CSectionContours: Set restored = New CSectionContours
    Dim kind As String, id As String, loopID As String, role As String, source As String, comment As String, unitText As String
    Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double, cx As Double, cy As Double, radius As Double, sweep As Double, angle As Double
    Dim ids As Object: Set ids = CreateObject("Scripting.Dictionary")
    Dim snapshotRun As Variant, geometryAnchor As Object
    Set geometryAnchor = workbook.Names.Item("rngNDMSectionGeometry").RefersToRange
    If CStr(geometryAnchor.Value2) = "RunID" Then snapshotRun = geometryAnchor.Offset(1, 0).Value2
    For row = 2 To UBound(data, 1)
        If IsError(data(row, 1)) Or IsNull(data(row, 1)) Then GoTo InvalidRow
        If Len(CStr(data(row, 1))) = 0 Then GoTo InvalidRow
        If Not IsEmpty(snapshotRun) Then
            If CStr(data(row, 1)) <> CStr(snapshotRun) Then GoTo InvalidRow
        Else
            snapshotRun = data(row, 1)
        End If
        id = ReadSavedAnnotationText(data, anchor, row, 3)
        loopID = ReadSavedAnnotationText(data, anchor, row, 2): role = ReadSavedAnnotationText(data, anchor, row, 5)
        kind = ReadSavedAnnotationText(data, anchor, row, 6): source = ReadSavedAnnotationText(data, anchor, row, 15)
        comment = ReadSavedAnnotationText(data, anchor, row, 16): unitText = ReadSavedAnnotationText(data, anchor, row, 17)
        If Len(Trim$(id)) = 0 Or Len(Trim$(loopID)) = 0 Or Len(Trim$(unitText)) = 0 Then GoTo InvalidRow
        If ids.Exists(id) Then GoTo InvalidRow
        ids.Add id, True
        If ReadGeometryNumber(data, anchor, row, 4) <> row - 1 Then GoTo InvalidRow
        If StrComp(role, "Outer", vbTextCompare) <> 0 And StrComp(role, "Opening", vbTextCompare) <> 0 Then GoTo InvalidRow
        x1 = OutputLengthToInternalByUnit(ReadGeometryNumber(data, anchor, row, 7), unitText)
        y1 = OutputLengthToInternalByUnit(ReadGeometryNumber(data, anchor, row, 8), unitText)
        Select Case kind
            Case "CONTOUR_LINE"
                x2 = OutputLengthToInternalByUnit(ReadGeometryNumber(data, anchor, row, 9), unitText)
                y2 = OutputLengthToInternalByUnit(ReadGeometryNumber(data, anchor, row, 10), unitText)
                restored.AddContourLine id, x1, y1, x2, y2, comment, loopID, role, source
            Case "CONTOUR_CIRCLE"
                radius = OutputLengthToInternalByUnit(ReadGeometryNumber(data, anchor, row, 13), unitText)
                restored.AddContourCircle id, x1, y1, radius, comment, loopID, role, source
            Case "CONTOUR_ARC"
                cx = OutputLengthToInternalByUnit(ReadGeometryNumber(data, anchor, row, 11), unitText)
                cy = OutputLengthToInternalByUnit(ReadGeometryNumber(data, anchor, row, 12), unitText)
                sweep = ReadSavedContourArcSweep(data, anchor, row, 14)
                radius = Sqr((x1 - cx) ^ 2 + (y1 - cy) ^ 2)
                x2 = cx + (x1 - cx) * Cos(sweep) - (y1 - cy) * Sin(sweep) ' Поворот вектора начала однозначно восстанавливает конец, включая большие и отрицательные дуги.
                y2 = cy + (x1 - cx) * Sin(sweep) + (y1 - cy) * Cos(sweep)
                restored.AddContourArc id, x1, y1, x2, y2, cx, cy, radius, sweep, comment, loopID, role, source
            Case Else
                GoTo InvalidRow
        End Select
    Next row
    If contours Is Nothing Then
        Set contours = restored
    Else
        contours.Clear
        CopySectionContours restored, contours
    End If
    Set ReadSavedSectionContours = contours
    Exit Function
MissingAnchor:
    Err.Raise vbObjectError + 4376, "ReadSavedSectionContours", "В книге отсутствует таблица достоверных контуров rngNDMSectionContours. Нажмите «Сформировать сечение»."
InvalidHeader:
    Err.Raise vbObjectError + 4376, "ReadSavedSectionContours", "Неверная или пустая шапка таблицы контуров: " & anchor.Parent.Name & "!" & anchor.Address(False, False) & ". Нажмите «Сформировать сечение»."
InvalidRow:
    Err.Raise vbObjectError + 4376, "ReadSavedSectionContours", "Неверная запись контура: " & anchor.Parent.Name & "!" & anchor.Offset(row - 1, 0).Resize(1, 17).Address(False, False) & ". Проверьте ID, порядок, роль, тип и единицы."
End Function

' Переносит проверенный контейнер без геометрической классификации.
Private Sub CopySectionContours(ByVal source As CSectionContours, ByVal target As CSectionContours)
    Dim i As Long
    For i = 1 To source.Count
        Select Case source.SegmentType(i)
            Case "CONTOUR_LINE"
                target.AddContourLine source.ContourID(i), source.StartX(i), source.StartY(i), source.EndX(i), source.EndY(i), source.Comment(i), source.LoopID(i), source.LoopRole(i), source.SourceID(i)
            Case "CONTOUR_CIRCLE"
                target.AddContourCircle source.ContourID(i), source.StartX(i), source.StartY(i), source.Radius(i), source.Comment(i), source.LoopID(i), source.LoopRole(i), source.SourceID(i)
            Case "CONTOUR_ARC"
                target.AddContourArc source.ContourID(i), source.StartX(i), source.StartY(i), source.EndX(i), source.EndY(i), source.CenterX(i), source.CenterY(i), source.Radius(i), source.SweepAngle(i), source.Comment(i), source.LoopID(i), source.LoopRole(i), source.SourceID(i)
        End Select
    Next i
End Sub

' Читает текст сохранённой аннотации сечения. Ошибка формулы не заменяется
' пустой строкой; её адрес определяется по текущему началу именованной таблицы.
Public Function ReadSavedAnnotationText(ByRef data As Variant, ByVal anchor As Object, _
        ByVal row As Long, ByVal column As Long) As String
    If IsError(data(row, column)) Or IsNull(data(row, column)) Then Err.Raise vbObjectError + 4376, _
        "ReadSectionGeometryFromResults", "В сохраненном контуре Results поле " & CStr(data(1, column)) & _
        " содержит ошибку: " & anchor.Parent.Name & "!" & anchor.Offset(row - 1, column - 1).Address(False, False) & _
        ". Исправьте ячейку или нажмите «Сформировать сечение»."
    ReadSavedAnnotationText = CStr(data(row, column))
End Function

' Читает угол дуги из массива. Адрес ячейки запрашивается у Excel только
' при ошибке; число обращений не растёт с числом сохранённых дуг.
Public Function ReadSavedContourArcSweep(ByRef data As Variant, ByVal anchor As Object, _
        ByVal row As Long, ByVal column As Long) As Double
    On Error GoTo InvalidSweep
    ReadSavedContourArcSweep = ReadContourArcSweep(data(row, column), vbNullString)
    Exit Function
InvalidSweep:
    ReadSavedContourArcSweep = ReadContourArcSweep(data(row, column), _
        anchor.Parent.Name & "!" & anchor.Offset(row - 1, column - 1).Address(False, False))
End Function

' Возвращает источник уже сформированного сечения из свойств Results.
' Текущий выбор Geometry.Source не меняет происхождение сохранённого снимка.
Public Function ResultsGeometrySource(ByVal workbook As Object) As String
    On Error GoTo Failed
    Dim data As Variant
    data = ReadAnchoredResultTable(workbook, "rngNDMSectionProperties")
    If Not HasResultTableRows(data) Then Exit Function

    Dim colLoadCase As Long: colLoadCase = ResultColumn(data, "LoadCase")
    Dim colParameter As Long: colParameter = ResultColumn(data, "Parameter")
    Dim colValue As Long: colValue = ResultColumn(data, "Value")

    Dim rowIndex As Long
    For rowIndex = 2 To UBound(data, 1)
        If StrComp(CStr(data(rowIndex, colLoadCase)), "ALL", vbTextCompare) = 0 And _
                StrComp(CStr(data(rowIndex, colParameter)), "Geometry.Source", vbTextCompare) = 0 Then
            ResultsGeometrySource = Trim$(CStr(data(rowIndex, colValue)))
            Exit Function
        End If
    Next rowIndex
Failed:
End Function

' Читает текстовое свойство выбранного сочетания из таблицы параметров Results.
' Если свойства нет, текущее значение из Config не подставляется.
Public Function ReadLoadCasePropertyText(ByVal workbook As Object, ByVal combinationID As String, ByVal propertyName As String) As String
    On Error GoTo Failed
    Dim data As Variant
    data = ReadAnchoredResultTable(workbook, "rngNDMSectionProperties")
    If Not HasResultTableRows(data) Then Exit Function

    Dim colLoadCase As Long: colLoadCase = ResultColumn(data, "LoadCase")
    Dim colParameter As Long: colParameter = ResultColumn(data, "Parameter")
    Dim colValue As Long: colValue = ResultColumn(data, "Value")

    Dim rowIndex As Long
    For rowIndex = 2 To UBound(data, 1)
        If StrComp(Trim$(SafeText(data(rowIndex, colLoadCase))), combinationID, vbTextCompare) = 0 And _
                StrComp(Trim$(SafeText(data(rowIndex, colParameter))), propertyName, vbTextCompare) = 0 Then
            ReadLoadCasePropertyText = Trim$(SafeText(data(rowIndex, colValue)))
            Exit Function
        End If
    Next rowIndex
Failed:
End Function

' Читает численное поле готовой геометрии. Только разрешенная пустота означает
' отсутствие оболочки/локальной инерции; обязательные координаты/A/диаметр,
' текстовый префикс, Boolean и ошибка формулы не превращаются в ноль/число.
' Диагностика указывает фактическую ячейку даже после переноса таблицы.
Public Function ReadGeometryNumber(ByRef data As Variant, ByVal anchor As Object, _
        ByVal row As Long, ByVal column As Long, Optional ByVal allowBlank As Boolean = False) As Double
    Dim value As Variant
    value = data(row, column)
    If IsError(value) Or IsNull(value) Then GoTo InvalidNumber
    If VarType(value) = vbBoolean Or VarType(value) = vbDate Then GoTo InvalidNumber
    If Len(Trim$(CStr(value))) = 0 Then
        If allowBlank Then Exit Function
        GoTo InvalidNumber
    End If
    If Not IsNumeric(value) Then GoTo InvalidNumber
    On Error GoTo InvalidNumber
    ReadGeometryNumber = CDbl(value)
    Exit Function
InvalidNumber:
    Err.Raise vbObjectError + 4367, "ReadSectionGeometryFromResults", _
        "В сохраненной геометрии Results поле " & ResultHeaderBase(CStr(data(1, column))) & _
        " содержит нечисловое значение: " & anchor.Worksheet.Name & "!" & _
        anchor.Offset(row - 1, column - 1).Address(False, False) & _
        ". Нажмите «Сформировать сечение», чтобы восстановить корректную геометрию."
End Function

' Проверяет, есть ли в Variant-таблице заголовок и строка данных. Локальный
' перехват относится только к отсутствующему/непрямоугольному массиву.
Public Function HasResultTableRows(ByVal data As Variant) As Boolean
    On Error GoTo Failed
    HasResultTableRows = (UBound(data, 1) >= 2 And UBound(data, 2) >= 1)
    Exit Function
Failed:
    HasResultTableRows = False
End Function

' Находит обязательный столбец по названию без единиц измерения.
' Для обязательного поля отсутствие столбца является ошибкой структуры.
' Необязательное поле возвращает ноль, поэтому расчёт не требует JSON и комментариев.
Public Function ResultColumn(ByRef data As Variant, ByVal headerName As String, Optional ByVal isRequired As Boolean = True) As Long
    Dim colIndex As Long
    For colIndex = 1 To UBound(data, 2)
        If StrComp(ResultHeaderBase(CStr(data(1, colIndex))), headerName, vbTextCompare) = 0 Then
            ResultColumn = colIndex
            Exit Function
        End If
    Next colIndex
    If isRequired Then Err.Raise vbObjectError + 4355, "ResultColumn", "В Results не найден столбец: " & headerName
End Function

' Отделяет название столбца от единиц после запятой.
' Поиск поля в сохранённой таблице не зависит от выходных единиц.
Public Function ResultHeaderBase(ByVal headerText As String) As String
    Dim commaPos As Long
    commaPos = InStr(1, headerText, ",", vbTextCompare)
    If commaPos > 0 Then
        ResultHeaderBase = Trim$(Left$(headerText, commaPos - 1))
    Else
        ResultHeaderBase = Trim$(headerText)
    End If
End Function

' Читает таблицу Results от её именованной начальной ячейки.
' Не использует CurrentRegion: расположенные рядом пояснения и заголовки
' не должны попадать в массив расчётных данных.
Public Function ReadAnchoredResultTable(ByVal workbook As Object, ByVal rangeName As String) As Variant
    Dim useReadTables As Boolean
    useReadTables = Not mSnapshotReadTables Is Nothing And workbook Is mSnapshotReadWorkbook
    If useReadTables Then
        If mSnapshotReadTables.Exists(rangeName) Then
            ReadAnchoredResultTable = mSnapshotReadTables.Item(rangeName)
            Exit Function
        End If
    End If
    Dim anchor As Object
    Set anchor = workbook.Names.Item(rangeName).RefersToRange

    Dim columnCount As Long
    columnCount = AnchoredColumnCount(anchor)
    If columnCount <= 0 Then Exit Function

    Dim rowCount As Long
    rowCount = AnchoredRowCount(anchor)
    If rowCount <= 0 Then rowCount = 1

    ReadAnchoredResultTable = anchor.Resize(rowCount, columnCount).Value2
    If useReadTables Then
        mSnapshotReadTables.Add rangeName, ReadAnchoredResultTable
        mSnapshotTableReadCount = mSnapshotTableReadCount + 1
    End If
End Function

' Определяет ширину таблицы по непрерывной строке заголовков.
Public Function AnchoredColumnCount(ByVal anchor As Object) As Long
    AnchoredColumnCount = SnapshotAnchoredExtent(anchor, True, 256)
End Function

' Определяет число строк таблицы по её первому столбцу.
' В каждой строке данных там обязательно записан RunID, ElementID или AnnotationID.
Public Function AnchoredRowCount(ByVal anchor As Object) As Long
    AnchoredRowCount = SnapshotAnchoredExtent(anchor, False)
End Function

' Достает единицу измерения из заголовка вида "X, mm".
Public Function ResultHeaderUnit(ByRef data As Variant, ByVal colIndex As Long, ByVal defaultUnit As String) As String
    Dim headerText As String
    headerText = CStr(data(1, colIndex))

    Dim commaPos As Long
    commaPos = InStr(1, headerText, ",", vbTextCompare)
    If commaPos <= 0 Then
        ResultHeaderUnit = defaultUnit
        Exit Function
    End If

    Dim unitText As String
    unitText = Trim$(Mid$(headerText, commaPos + 1))
    Dim bracketPos As Long
    bracketPos = InStr(1, unitText, "(", vbTextCompare)
    If bracketPos > 0 Then unitText = Trim$(Left$(unitText, bracketPos - 1))

    If Len(unitText) = 0 Then unitText = defaultUnit
    ResultHeaderUnit = unitText
End Function

' Переводит длину из единиц сохранённых Results в мм.
Public Function OutputLengthToInternalByUnit(ByVal value As Double, ByVal unitText As String) As Double
    If mSnapshotUnits Is Nothing Then Set mSnapshotUnits = New CUnitSystem
    OutputLengthToInternalByUnit = mSnapshotUnits.OutputLengthToInternal(value, unitText)
End Function

' Переводит площадь из единиц сохранённых Results в мм2.
Public Function OutputAreaToInternalByUnit(ByVal value As Double, ByVal unitText As String) As Double
    If mSnapshotUnits Is Nothing Then Set mSnapshotUnits = New CUnitSystem
    OutputAreaToInternalByUnit = mSnapshotUnits.OutputAreaToInternal(value, unitText)
End Function

' Переводит собственные моменты инерции элемента из единиц Results в мм4.
Public Function OutputFourthPowerLengthToInternalByUnit(ByVal value As Double, ByVal unitText As String) As Double
    Dim baseUnit As String
    baseUnit = Trim$(unitText)
    If Right$(baseUnit, 1) = "4" Then baseUnit = Left$(baseUnit, Len(baseUnit) - 1)
    If mSnapshotUnits Is Nothing Then Set mSnapshotUnits = New CUnitSystem
    OutputFourthPowerLengthToInternalByUnit = mSnapshotUnits.OutputFourthPowerLengthToInternal(value, baseUnit)
End Function

' Переводит кривизну из единиц сохранённых Results в 1/мм.
Public Function OutputCurvatureToInternalByUnit(ByVal value As Double, ByVal unitText As String) As Double
    If Len(Trim$(unitText)) = 0 Then Err.Raise vbObjectError + 4364, "OutputCurvatureToInternalByUnit", _
        "В Results не указана единица кривизны. Повторите расчет, чтобы восстановить снимок."
    If mSnapshotUnits Is Nothing Then Set mSnapshotUnits = New CUnitSystem
    OutputCurvatureToInternalByUnit = value / mSnapshotUnits.InternalCurvatureToOutput(1#, unitText)
End Function

' Преобразует значение в текст через CStr; для ошибки Excel возвращает пустую строку.
Public Function SafeText(ByVal value As Variant) As String
    If IsError(value) Then Exit Function
    SafeText = CStr(value)
End Function

' Считает заполненные ячейки строки или столбца от заданного начала таблицы.
' Читает небольшими блоками Value2 до первой пустой ячейки, пустого текста
' или одних пробелов. Ошибка Excel вызывает ошибку чтения, а не завершает таблицу.
' Соседние ячейки за концом таблицы не включаются в результат.
Public Function SnapshotAnchoredExtent(ByVal anchor As Object, ByVal columns As Boolean, _
        Optional ByVal maxCount As Long = 0) As Long
    Const BLOCK_SIZE As Long = 1024
    Dim available As Long, offset As Long, count As Long, i As Long
    If columns Then
        available = anchor.Worksheet.Columns.Count - anchor.Column + 1
    Else
        available = anchor.Worksheet.Rows.Count - anchor.Row + 1
    End If
    If maxCount > 0 And maxCount < available Then available = maxCount
    Dim block As Variant, value As Variant
    Do While offset < available
        count = available - offset
        If count > BLOCK_SIZE Then count = BLOCK_SIZE
        If columns Then
            block = anchor.Offset(0, offset).Resize(1, count).Value2
        Else
            block = anchor.Offset(offset, 0).Resize(count, 1).Value2
        End If
        For i = 1 To count
            If count = 1 Then
                value = block
            ElseIf columns Then
                value = block(1, i)
            Else
                value = block(i, 1)
            End If
            If Len(Trim$(CStr(value))) = 0 Then Exit Function
            SnapshotAnchoredExtent = SnapshotAnchoredExtent + 1
        Next i
        offset = offset + count
    Loop
End Function

' Начинает одно чтение согласованного снимка: каждая таблица получается один раз.
Public Sub BeginSectionSnapshotRead(ByVal workbook As Object)
    Set mSnapshotReadTables = CreateObject("Scripting.Dictionary")
    Set mSnapshotReadWorkbook = workbook
    mSnapshotTableReadCount = 0
End Sub

' Освобождает массивы снимка после чтения или ошибки, сохраняя счётчик обращений.
Public Sub EndSectionSnapshotRead()
    Set mSnapshotReadTables = Nothing: Set mSnapshotReadWorkbook = Nothing
End Sub


' Читает обязательную геометрию для расчёта. Необязательные сведения и
' аннотации добавляются только по запросу; генератор и AutoCAD не вызываются.
Public Function ReadFormedSectionFromResults(ByVal workbook As Object, Optional ByVal includeOptionalData As Boolean = False) As CSectionModel
    On Error GoTo Failed
    BeginSectionSnapshotRead workbook ' Общие таблицы читаются один раз за операцию; кэш освобождается и при ошибке.
    Dim source As String: source = ResultsGeometrySource(workbook)
    If Len(source) = 0 Then Err.Raise vbObjectError + 4380, "ReadFormedSectionFromResults", _
        "В Results нет сформированного сечения. Сначала нажмите «Сформировать сечение»."
    Dim model As CSectionModel: Set model = ReadSectionGeometryFromResults(workbook, source, includeOptionalData)
    If includeOptionalData Then
        model.GeometryDiagnosticsText = ReadLoadCasePropertyText(workbook, "ALL", "Geometry.RecipeDiagnostics")
        RestoreModelAnnotations workbook, model
    End If
    EndSectionSnapshotRead
    Set ReadFormedSectionFromResults = model
    Exit Function
Failed:
    Dim number As Long, description As String, origin As String
    number = Err.Number: description = Err.Description: origin = Err.Source
    EndSectionSnapshotRead
    Err.Raise number, origin, description
End Function

' Восстанавливает исходные аннотации с числовыми значениями в мм и исходным
' обозначением размера; готовый округлённый текст не используется как ввод.
Private Sub RestoreModelAnnotations(ByVal workbook As Object, ByVal model As CSectionModel)
    Dim anchor As Object, data As Variant, row As Long, info As Object, kind As String, id As String, unit As String
    Set anchor = workbook.Names.Item("rngNDMSectionAnnotations").RefersToRange
    data = ReadAnchoredResultTable(workbook, "rngNDMSectionAnnotations")
    If Not HasResultTableRows(data) Then Exit Sub
    If UBound(data, 2) <> 24 Then Err.Raise vbObjectError + 4381, "RestoreModelAnnotations", "Неверная структура rngNDMSectionAnnotations."
    Dim text As String, value As Double, valueUnit As String, comment As String
    For row = 2 To UBound(data, 1)
        kind = ReadSavedAnnotationText(data, anchor, row, 2)
        If Left$(kind, 6) <> "CRACK_" Then
            id = ReadSavedAnnotationText(data, anchor, row, 3)
            unit = ReadSavedAnnotationText(data, anchor, row, 23)
            Set info = ReadSnapshotOptionalRecord(data, anchor, row, 24)
            text = ReadSavedAnnotationText(data, anchor, row, 10)
            value = ReadGeometryNumber(data, anchor, row, 11, True)
            valueUnit = ReadSavedAnnotationText(data, anchor, row, 12)
            comment = ReadSavedAnnotationText(data, anchor, row, 13)
            If info.Exists("OriginalText") Then text = SectionOptionalText(info, "OriginalText")
            If info.Exists("OriginalValue") Then value = ReadSnapshotOptionalNumber(SectionOptionalText(info, "OriginalValue"), anchor, row, 24)
            If info.Exists("OriginalUnit") Then valueUnit = SectionOptionalText(info, "OriginalUnit")
            If info.Exists("OriginalComment") Then comment = SectionOptionalText(info, "OriginalComment")
            model.Annotations.AddAnnotation kind, id, _
                OutputLengthToInternalByUnit(ReadGeometryNumber(data, anchor, row, 4), unit), _
                OutputLengthToInternalByUnit(ReadGeometryNumber(data, anchor, row, 5), unit), _
                OutputLengthToInternalByUnit(ReadGeometryNumber(data, anchor, row, 6), unit), _
                OutputLengthToInternalByUnit(ReadGeometryNumber(data, anchor, row, 7), unit), _
                ReadGeometryNumber(data, anchor, row, 8), ReadGeometryNumber(data, anchor, row, 9), text, value, valueUnit, comment
        End If
    Next row
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' ==========================================================================
' ДЛЯ ТЕСТОВ
' ==========================================================================
' Показывает число блочных чтений и освобождение таблиц после операции.
Public Function TestSectionSnapshotReadDiagnostics() As String
    TestSectionSnapshotReadDiagnostics = "tables=" & CStr(mSnapshotTableReadCount) & ";released=" & _
        CStr(mSnapshotReadTables Is Nothing And mSnapshotReadWorkbook Is Nothing)
End Function
