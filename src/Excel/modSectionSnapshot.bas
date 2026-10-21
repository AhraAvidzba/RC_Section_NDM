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

' Восстанавливает CSectionModel из Results во внутренних единицах.
' Площадь и моменты инерции импортированных элементов сохраняются отдельно
' от их изображения. Если сохранённой геометрии нет, выдаёт ошибку.
Public Function ReadSectionGeometryFromResults(ByVal workbook As Object, Optional ByVal sourceType As String = "Results", _
        Optional ByVal metadata As Object = Nothing) As CSectionModel
    Dim anchor As Object
    Set anchor = workbook.Names.Item("rngNDMSectionGeometry").RefersToRange

    Dim data As Variant
    data = ReadAnchoredResultTable(workbook, "rngNDMSectionGeometry")
    If Not HasResultTableRows(data) Then Err.Raise vbObjectError + 4350, "ReadSectionGeometryFromResults", _
        "На листе Results нет таблицы расчетной геометрии. Нажмите «Сформировать сечение»."
    If UBound(data, 2) <> 15 Then Err.Raise vbObjectError + 4355, "ReadSectionGeometryFromResults", _
        "Таблица расчетной геометрии Results должна содержать 15 столбцов. Нажмите «Сформировать сечение»."

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
    Dim colComment As Long: colComment = ResultColumn(data, "Comment")
    Dim lengthUnit As String: lengthUnit = ResultHeaderUnit(data, colX, "mm")
    Dim areaUnit As String: areaUnit = ResultHeaderUnit(data, colArea, "mm2")
    Dim fourthUnit As String: fourthUnit = ResultHeaderUnit(data, colLocalIx, lengthUnit & "4")

    Dim model As CSectionModel
    Set model = New CSectionModel
    model.SourceType = sourceType
    model.ConcreteMeshSource = ReadLoadCasePropertyText(workbook, "ALL", "Geometry.ConcreteMeshSource")

    Dim rowIndex As Long, info As Variant, elementType As String, elementID As String
    For rowIndex = 2 To UBound(data, 1)
        elementID = ReadSavedAnnotationText(data, anchor, rowIndex, colID)
        If Len(Trim$(elementID)) > 0 Then
            elementType = ReadSavedAnnotationText(data, anchor, rowIndex, colType)
            If Not metadata Is Nothing Then
                If IsError(data(rowIndex, 1)) Or IsNull(data(rowIndex, 1)) Then GoTo InvalidElementID
                If CStr(data(rowIndex, 1)) <> CStr(data(2, 1)) Then GoTo InvalidElementID
                info = RequiredSnapshotMetadata(metadata, elementType, elementID, anchor, rowIndex, colID)
                If elementType = "Concrete" Then
                    model.AddConcreteElement OutputLengthToInternalByUnit(ReadGeometryNumber(data, anchor, rowIndex, colX), lengthUnit), _
                        OutputLengthToInternalByUnit(ReadGeometryNumber(data, anchor, rowIndex, colY), lengthUnit), _
                        OutputAreaToInternalByUnit(ReadGeometryNumber(data, anchor, rowIndex, colArea), areaUnit), CLng(info(6)), CStr(info(4)), CStr(info(5)), _
                        CStr(info(9)), CDbl(info(10)), CDbl(info(11)), CDbl(info(12)), CStr(info(17)), _
                        OutputFourthPowerLengthToInternalByUnit(ReadGeometryNumber(data, anchor, rowIndex, colLocalIx), fourthUnit), _
                        OutputFourthPowerLengthToInternalByUnit(ReadGeometryNumber(data, anchor, rowIndex, colLocalIy), fourthUnit), _
                        OutputFourthPowerLengthToInternalByUnit(ReadGeometryNumber(data, anchor, rowIndex, colLocalIxy), fourthUnit), CBool(info(13))
                    If model.ConcreteID(model.ConcreteCount) <> elementID Then GoTo InvalidElementID
                    model.RestoreConcreteRotationKnowledge model.ConcreteCount, CBool(info(13))
                ElseIf elementType = "Rebar" Then
                    model.AddRebarElement OutputLengthToInternalByUnit(ReadGeometryNumber(data, anchor, rowIndex, colX), lengthUnit), _
                        OutputLengthToInternalByUnit(ReadGeometryNumber(data, anchor, rowIndex, colY), lengthUnit), _
                        OutputLengthToInternalByUnit(ReadGeometryNumber(data, anchor, rowIndex, colDiameter), lengthUnit), _
                        OutputAreaToInternalByUnit(ReadGeometryNumber(data, anchor, rowIndex, colArea), areaUnit), CStr(info(7)), CLng(info(6)), _
                        CStr(info(4)), CStr(info(5)), CStr(info(17))
                    If model.RebarID(model.RebarCount) <> elementID Then GoTo InvalidElementID
                    model.RestoreRebarRecipeTrace model.RebarCount, CStr(info(8))
                Else
                    GoTo InvalidElementID
                End If
            ElseIf StrComp(elementType, "Concrete", vbTextCompare) = 0 Then
                ' В Results уже записан окончательный угол изображения элемента.
                ' Повторный экспорт сохраняет его, в том числе нулевой угол горизонтальной
                ' грани.
                model.AddConcreteElement OutputLengthToInternalByUnit(ReadGeometryNumber(data, anchor, rowIndex, colX), lengthUnit), OutputLengthToInternalByUnit(ReadGeometryNumber(data, anchor, rowIndex, colY), lengthUnit), _
                    OutputAreaToInternalByUnit(ReadGeometryNumber(data, anchor, rowIndex, colArea), areaUnit), 1, vbNullString, vbNullString, _
                    CStr(data(rowIndex, colShape)), OutputLengthToInternalByUnit(ReadGeometryNumber(data, anchor, rowIndex, colWidth, True), lengthUnit), _
                    OutputLengthToInternalByUnit(ReadGeometryNumber(data, anchor, rowIndex, colHeight, True), lengthUnit), ReadGeometryNumber(data, anchor, rowIndex, colRotation, True), _
                    CStr(data(rowIndex, colComment)), OutputFourthPowerLengthToInternalByUnit(ReadGeometryNumber(data, anchor, rowIndex, colLocalIx, True), fourthUnit), _
                    OutputFourthPowerLengthToInternalByUnit(ReadGeometryNumber(data, anchor, rowIndex, colLocalIy, True), fourthUnit), OutputFourthPowerLengthToInternalByUnit(ReadGeometryNumber(data, anchor, rowIndex, colLocalIxy, True), fourthUnit), True
            ElseIf StrComp(CStr(data(rowIndex, colType)), "Rebar", vbTextCompare) = 0 Then
                model.AddRebarElement OutputLengthToInternalByUnit(ReadGeometryNumber(data, anchor, rowIndex, colX), lengthUnit), OutputLengthToInternalByUnit(ReadGeometryNumber(data, anchor, rowIndex, colY), lengthUnit), _
                    OutputLengthToInternalByUnit(ReadGeometryNumber(data, anchor, rowIndex, colDiameter), lengthUnit), OutputAreaToInternalByUnit(ReadGeometryNumber(data, anchor, rowIndex, colArea), areaUnit), _
                    vbNullString, 1, vbNullString, _
                    vbNullString, CStr(data(rowIndex, colComment))
            Else
                Err.Raise vbObjectError + 4355, "ReadSectionGeometryFromResults", _
                    "В таблице расчетных элементов Results допустимы только Concrete и Rebar: " & _
                    anchor.Parent.Name & "!" & anchor.Offset(rowIndex - 1, colType - 1).Address(False, False) & ". Нажмите «Сформировать сечение»."
            End If
        End If
    Next rowIndex

    If metadata Is Nothing Then model.ApplyAverageRotationToEquivalentAreaFallbacks
    If model.ConcreteCount <= 0 Then Err.Raise vbObjectError + 4351, "ReadSectionGeometryFromResults", _
        "В таблице Results нет бетонных элементов."
    ReadSavedSectionContours workbook, model.Contours
    Set ReadSectionGeometryFromResults = model
    Exit Function
InvalidElementID:
    Err.Raise vbObjectError + 4381, "ReadSectionGeometryFromResults", _
        "Нарушены тип, порядок или имя расчётного элемента: " & anchor.Parent.Name & "!" & _
        anchor.Offset(rowIndex - 1, colID - 1).Address(False, False) & ". Нажмите «Сформировать сечение»."
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
                ' Поворот вектора начала однозначно восстанавливает конец, включая большие и отрицательные дуги.
                x2 = cx + (x1 - cx) * Cos(sweep) - (y1 - cy) * Sin(sweep)
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

' Читает источник сохранённой геометрии из rngNDMSectionProperties.
' При выборе Geometry.Source = AutoCAD позволяет отличить импортированное сечение от
' параметрического.
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
        ". Повторите импорт геометрии или расчет, чтобы восстановить корректный снимок."
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
' Если столбца нет, сообщает об ошибке структуры Results до чтения данных.
Public Function ResultColumn(ByRef data As Variant, ByVal headerName As String) As Long
    Dim colIndex As Long
    For colIndex = 1 To UBound(data, 2)
        If StrComp(ResultHeaderBase(CStr(data(1, colIndex))), headerName, vbTextCompare) = 0 Then
            ResultColumn = colIndex
            Exit Function
        End If
    Next colIndex
    Err.Raise vbObjectError + 4355, "ResultColumn", "В Results не найден столбец: " & headerName
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

' Безопасно читает текст из ячейки Results.
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


' Читает полный снимок для расчёта, включая класс стали, исходные оболочки,
' связи стержней и аннотации. Config, генератор и AutoCAD не вызываются.
Public Function ReadFormedSectionFromResults(ByVal workbook As Object) As CSectionModel
    On Error GoTo Failed
    BeginSectionSnapshotRead workbook
    Dim source As String: source = ResultsGeometrySource(workbook)
    If Len(source) = 0 Then Err.Raise vbObjectError + 4380, "ReadFormedSectionFromResults", _
        "В Results нет сформированного сечения. Сначала нажмите «Сформировать сечение»."
    Dim metadata As Object: Set metadata = ReadModelMetadata(workbook)
    Dim model As CSectionModel: Set model = ReadSectionGeometryFromResults(workbook, source, metadata)
    model.GeometryDiagnosticsText = ReadLoadCasePropertyText(workbook, "ALL", "Geometry.RecipeDiagnostics")
    RestoreModelAnnotations workbook, model, metadata
    If metadata.Count <> model.ConcreteCount + model.RebarCount + model.AnnotationCount Then _
        Err.Raise vbObjectError + 4381, "ReadFormedSectionFromResults", _
            "Таблицы геометрии, аннотаций и rngNDMSectionModelMetadata не согласованы. Нажмите «Сформировать сечение»."
    EndSectionSnapshotRead
    Set ReadFormedSectionFromResults = model
    Exit Function
Failed:
    Dim number As Long, description As String, origin As String
    number = Err.Number: description = Err.Description: origin = Err.Source
    EndSectionSnapshotRead
    Err.Raise number, origin, description
End Function

' Проверяет структуру и идентификатор снимка один раз; точные ключи строк
' позволяют восстановить все элементы без повторного поиска по таблице.
Private Function ReadModelMetadata(ByVal workbook As Object) As Object
    Dim anchor As Object, data As Variant, geometry As Variant, headers As Variant
    Set anchor = workbook.Names.Item("rngNDMSectionModelMetadata").RefersToRange
    data = ReadAnchoredResultTable(workbook, "rngNDMSectionModelMetadata")
    headers = Array("RunID", "RecordType", "ItemID", "SourceName", "SourceHandle", "MaterialID", _
        "SteelClass", "RecipeTrace", "SourceShapeType", "SourceWidth_mm", "SourceHeight_mm", _
        "SourceRotation_rad", "SourceRotationKnown", "OriginalText", "OriginalValue", "OriginalUnit", "OriginalComment")
    If Not HasResultTableRows(data) Then GoTo InvalidMetadata
    If UBound(data, 2) <> 17 Then GoTo InvalidMetadata
    Dim row As Long, column As Long, kind As String, key As String, record(1 To 17) As Variant
    For column = 1 To 17
        If ReadSavedAnnotationText(data, anchor, 1, column) <> CStr(headers(column - 1)) Then GoTo InvalidMetadata
    Next column
    geometry = ReadAnchoredResultTable(workbook, "rngNDMSectionGeometry")
    If Not HasResultTableRows(geometry) Then GoTo InvalidMetadata
    Dim result As Object: Set result = CreateObject("Scripting.Dictionary")
    For row = 2 To UBound(data, 1)
        If IsError(data(row, 1)) Or IsNull(data(row, 1)) Then GoTo InvalidMetadata
        If CStr(data(row, 1)) <> CStr(geometry(2, 1)) Then GoTo InvalidMetadata
        kind = ReadSavedAnnotationText(data, anchor, row, 2)
        If kind <> "Concrete" And kind <> "Rebar" And kind <> "Annotation" Then GoTo InvalidMetadata
        key = kind & "|" & ReadSavedAnnotationText(data, anchor, row, 3)
        If result.Exists(key) Then GoTo InvalidMetadata
        For column = 1 To 17
            If IsError(data(row, column)) Or IsNull(data(row, column)) Then GoTo InvalidMetadata
            record(column) = data(row, column)
        Next column
        If kind <> "Annotation" Then
            record(6) = ReadGeometryNumber(data, anchor, row, 6)
            If CDbl(record(6)) < 1# Or CDbl(record(6)) > 2147483647# Or CDbl(record(6)) <> Fix(CDbl(record(6))) Then GoTo InvalidMetadata
        End If
        If kind = "Concrete" Then
            For column = 10 To 12: record(column) = ReadGeometryNumber(data, anchor, row, column): Next column
            If VarType(record(13)) <> vbBoolean Then GoTo InvalidMetadata
        ElseIf kind = "Annotation" Then
            record(15) = ReadGeometryNumber(data, anchor, row, 15)
        End If
        result.Add key, record
    Next row
    Set ReadModelMetadata = result
    Exit Function
InvalidMetadata:
    If row < 1 Then row = 1
    Err.Raise vbObjectError + 4381, "ReadModelMetadata", _
        "Повреждены исходные сведения снимка: " & anchor.Parent.Name & "!" & anchor.Offset(row - 1, 0).Address(False, False) & _
        ". Нажмите «Сформировать сечение»."
End Function

' Получает сведения только по точному типу и имени; отсутствующая строка
' не заменяется материалом или аннотацией по умолчанию.
Private Function RequiredSnapshotMetadata(ByVal metadata As Object, ByVal kind As String, _
        ByVal id As String, ByVal anchor As Object, ByVal row As Long, ByVal column As Long) As Variant
    Dim key As String: key = kind & "|" & id
    If Not metadata.Exists(key) Then Err.Raise vbObjectError + 4381, "RequiredSnapshotMetadata", _
        "Для " & id & " нет исходных сведений снимка: " & anchor.Parent.Name & "!" & _
        anchor.Offset(row - 1, column - 1).Address(False, False) & _
        ". Нажмите «Сформировать сечение»."
    RequiredSnapshotMetadata = metadata.Item(key)
End Function

' Восстанавливает исходные аннотации с числовыми значениями в мм и исходным
' обозначением размера; готовый округлённый текст не используется как ввод.
Private Sub RestoreModelAnnotations(ByVal workbook As Object, ByVal model As CSectionModel, ByVal metadata As Object)
    Dim anchor As Object, data As Variant, row As Long, record As Variant, kind As String, id As String
    Set anchor = workbook.Names.Item("rngNDMSectionAnnotations").RefersToRange
    data = ReadAnchoredResultTable(workbook, "rngNDMSectionAnnotations")
    If Not HasResultTableRows(data) Then Exit Sub
    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    If UBound(data, 2) <> 22 Then Err.Raise vbObjectError + 4381, "RestoreModelAnnotations", _
        "Неверная структура rngNDMSectionAnnotations. Нажмите «Сформировать сечение»."
    Dim headers As Variant, column As Long
    headers = Array("RunID", "AnnotationType", "AnnotationID", "StartX", "StartY", "EndX", "EndY", _
        "OutsideNormalX", "OutsideNormalY", "Text", "Value", "Unit", "Comment")
    For column = 1 To 13
        If ReadSavedAnnotationText(data, anchor, 1, column) <> CStr(headers(column - 1)) Then _
            Err.Raise vbObjectError + 4381, "RestoreModelAnnotations", _
                "Неверный заголовок аннотации: " & anchor.Parent.Name & "!" & anchor.Offset(0, column - 1).Address(False, False) & _
                ". Нажмите «Сформировать сечение»."
    Next column
    Dim unit As String: unit = ReadLoadCasePropertyText(workbook, "ALL", "Output.LengthUnit")
    For row = 2 To UBound(data, 1)
        kind = ReadSavedAnnotationText(data, anchor, row, 2)
        If Left$(kind, 6) <> "CRACK_" Then
            id = ReadSavedAnnotationText(data, anchor, row, 3)
            If seen.Exists(id) Then Err.Raise vbObjectError + 4381, "RestoreModelAnnotations", _
                "Повторяется аннотация " & id & ". Нажмите «Сформировать сечение»."
            seen.Add id, True
            record = RequiredSnapshotMetadata(metadata, "Annotation", id, anchor, row, 3)
            If IsError(data(row, 1)) Or IsNull(data(row, 1)) Then GoTo InvalidAnnotationRun
            If CStr(data(row, 1)) <> CStr(record(1)) Then GoTo InvalidAnnotationRun
            model.Annotations.AddAnnotation kind, id, _
                OutputLengthToInternalByUnit(ReadGeometryNumber(data, anchor, row, 4), unit), _
                OutputLengthToInternalByUnit(ReadGeometryNumber(data, anchor, row, 5), unit), _
                OutputLengthToInternalByUnit(ReadGeometryNumber(data, anchor, row, 6), unit), _
                OutputLengthToInternalByUnit(ReadGeometryNumber(data, anchor, row, 7), unit), _
                ReadGeometryNumber(data, anchor, row, 8), ReadGeometryNumber(data, anchor, row, 9), _
                CStr(record(14)), CDbl(record(15)), CStr(record(16)), CStr(record(17))
        End If
    Next row
    Exit Sub
InvalidAnnotationRun:
    Err.Raise vbObjectError + 4381, "RestoreModelAnnotations", _
        "Аннотация относится к другому снимку: " & anchor.Parent.Name & "!" & anchor.Offset(row - 1, 0).Address(False, False) & _
        ". Нажмите «Сформировать сечение»."
End Sub

' ДЛЯ ТЕСТОВ
' Показывает число блочных чтений и освобождение таблиц после операции.
Public Function SectionSnapshotReadDiagnostics() As String
    SectionSnapshotReadDiagnostics = "tables=" & CStr(mSnapshotTableReadCount) & ";released=" & _
        CStr(mSnapshotReadTables Is Nothing And mSnapshotReadWorkbook Is Nothing)
End Function
