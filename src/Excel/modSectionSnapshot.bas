Attribute VB_Name = "modSectionSnapshot"
Option Explicit
Private mTestBoundaryFullReads As Long, mTestBoundaryHashReads As Long ' ДЛЯ ТЕСТОВ: число полных и подтверждённых хешем элементов.

' ==========================================================================
' Чтение сохранённого сечения и таблиц Results
' ==========================================================================
' Восстанавливает геометрию в CSectionModel независимо от её источника.
' Расчёт, схема и экспорт используют одни данные; модуль не обращается
' к AutoCAD, не строит сетку и не решает систему равновесия.

Private mSnapshotUnits As CUnitSystem ' Преобразование единиц, записанных в Results, независимо от текущего Config.
Private mSnapshotReadTables As Object ' Таблицы одной операции чтения; освобождаются после восстановления модели или ошибки.
Private mSnapshotReadWorkbook As Object ' Книга, которой принадлежат временно прочитанные таблицы.
Private mSnapshotReadGeometryOnly As Boolean ' Формирование модели читает общую геометрию; экспорт также получает LC из полного снимка.
Private mSnapshotTableReadCount As Long ' ДЛЯ ТЕСТОВ: число загрузок целых таблиц; чтения при поиске границ сюда не входят.

' Положение и единицы общих колонок одной неизменяемой матрицы проверяются один раз.
' Числовые значения каждой строки проходят отдельную строгую проверку.
Private Type TSavedBoundaryLayout
    XColumn As Long
    YColumn As Long
    AreaColumn As Long
    IxColumn As Long
    IyColumn As Long
    IxyColumn As Long
    WidthColumn As Long
    HeightColumn As Long
    RotationColumn As Long
    LengthUnit As String
    AreaUnit As String
    InertiaUnit As String
End Type

' Восстанавливает расчётную модель из обязательных координат, площадей,
' инерций и контуров единого источника: выбранного файла либо Results, в мм, мм2 и мм4. При includeOptionalData=True
' дополнительно читает JSON происхождения и RecipeTrace; для расчёта они не требуются.
' Восстановлением аннотаций управляет вызывающий ReadFormedSectionFromResults.
Public Function ReadSectionGeometryFromResults(ByVal workbook As Object, Optional ByVal sourceType As String = "Results", _
        Optional ByVal includeOptionalData As Boolean = False) As CSectionModel
    Dim anchor As Object: Set anchor = workbook.Names.Item("rngNDMSectionGeometry").RefersToRange
    Dim data As Variant: data = ReadAnchoredResultTable(workbook, "rngNDMSectionGeometry")
    If Not HasResultTableRows(data) Then Err.Raise vbObjectError + 4350, "ReadSectionGeometryFromResults", _
        "На листе Results нет сформированной геометрии. Нажмите «Сформировать сечение»."
    If UBound(data, 2) < 14 Or UBound(data, 2) > 16 Then Err.Raise vbObjectError + 4355, "ReadSectionGeometryFromResults", _
        "Неверная структура rngNDMSectionGeometry: " & anchor.Parent.Name & "!" & _
        anchor.Resize(1, UBound(data, 2)).Address(False, False) & ". Допустимы 14 обязательных полей и только Comment/ExtraData после них. Нажмите «Сформировать сечение»."
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
    Dim boundaries As Object: Set boundaries = ReadSavedConcreteBoundaries(workbook, data)
    Dim exactRegion As CConcreteRegion
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
                If info.Exists("MaterialID") Then materialID = ReadSnapshotMaterialID(SectionOptionalText(info, "MaterialID"), anchor, row, extraColumn)
                If colComment > 0 Then comment = ReadSavedAnnotationText(data, anchor, row, colComment)
            End If
            If StrComp(kind, "Concrete", vbTextCompare) = 0 Then
                Dim localIx As Double, localIy As Double, localIxy As Double
                Dim shape As String: shape = ReadSavedAnnotationText(data, anchor, row, colShape)
                localIx = ReadPositiveGeometryNumber(data, anchor, row, colLocalIx)
                localIy = ReadPositiveGeometryNumber(data, anchor, row, colLocalIy)
                localIxy = ReadGeometryNumber(data, anchor, row, colLocalIxy)
                If Abs(localIxy) / Sqr(localIx) / Sqr(localIy) >= 1# Then Err.Raise vbObjectError + 4355, _
                    "ReadSectionGeometryFromResults", "Несогласованные собственные моменты инерции бетонного элемента: " & _
                    anchor.Parent.Name & "!" & anchor.Offset(row - 1, colLocalIxy - 1).Address(False, False) & ". Нажмите «Сформировать сечение»."
                Set exactRegion = Nothing
                If boundaries.Exists(id) Then Set exactRegion = boundaries.Item(id)
                model.AddConcreteElement OutputLengthToInternalByUnit(ReadGeometryNumber(data, anchor, row, colX), lengthUnit), _
                    OutputLengthToInternalByUnit(ReadGeometryNumber(data, anchor, row, colY), lengthUnit), _
                    OutputAreaToInternalByUnit(ReadPositiveGeometryNumber(data, anchor, row, colArea), areaUnit), materialID, sourceName, handle, _
                    shape, _
                    OutputLengthToInternalByUnit(ReadConcreteSnapshotDimension(data, anchor, row, colWidth, shape), lengthUnit), _
                    OutputLengthToInternalByUnit(ReadConcreteSnapshotDimension(data, anchor, row, colHeight, shape), lengthUnit), _
                    ReadGeometryNumber(data, anchor, row, colRotation), comment, _
                    OutputFourthPowerLengthToInternalByUnit(localIx, fourthUnit), _
                    OutputFourthPowerLengthToInternalByUnit(localIy, fourthUnit), _
                    OutputFourthPowerLengthToInternalByUnit(localIxy, fourthUnit), True, exactRegion ' Это готовая расчётная граница ячейки. Угол уже определён при формировании; повторное усреднение направлений не требуется.
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

' Читает определяющее сочетание сохранённого расчёта, не доверяя справочной сводке при File.
' Переключатели выбора сочетания остаются в Config; Worst принадлежит конкретному снимку.
Public Function ReadSavedWorstCombinationID(ByVal workbook As Object) As String
    Dim snapshot As CCalculationSnapshot: Set snapshot = CurrentCalculationSnapshot(workbook)
    If Not snapshot Is Nothing Then
        If snapshot.IsGeometryPreview Then Exit Function
        Dim summary As Variant: summary = snapshot.Table("Summary")
        ReadSavedWorstCombinationID = Trim$(SafeText(summary(1, 1)))
    Else
        ReadSavedWorstCombinationID = Trim$(SafeText(workbook.Names.Item("rngBatchSummary").RefersToRange.Cells(1, 1).Value2))
    End If
End Function

' Выбирает одно сочетание для схемы и AutoCAD по общим настройкам Config.
' User читает только ID; Worst игнорирует ID и использует сохранённый итог расчёта.
' При отсутствии определяющей строки сохранённые состояния проверяются в порядке таблицы.
Public Function ReadOutputCombinationID(ByVal workbook As Object, ByVal settings As CSystemSettingsReader) As String
    Dim mode As String: mode = settings.GetRequiredChoice("Output.CombinationMode", Array("Worst", "User"))
    If mode = "User" Then
        ReadOutputCombinationID = settings.GetRequiredString("Output.CombinationID")
        Exit Function
    End If
    ReadOutputCombinationID = ReadSavedWorstCombinationID(workbook)
    If Len(ReadOutputCombinationID) = 0 Then ReadOutputCombinationID = FirstOutputCombinationInTable(workbook, "rngNDMElementResults")
    If Len(ReadOutputCombinationID) = 0 Then ReadOutputCombinationID = FirstOutputCombinationInTable(workbook, "rngNDMSectionProperties")
    If Len(ReadOutputCombinationID) = 0 Then Err.Raise vbObjectError + 4702, "ReadOutputCombinationID", _
        settings.InputErrorMessage("Output.CombinationMode", _
            "В сохранённом снимке нет определяющего или другого рассчитанного сочетания для режима Worst.", _
            "Выполните расчёт либо выберите User и ID сочетания, уже сохранённого в текущем снимке.")
End Function

' Возвращает первое сохранённое сочетание, если верхняя сводка не содержит Worst.
' Общие свойства ALL не являются сочетанием. Отсутствующая необязательная таблица пропускается.
Private Function FirstOutputCombinationInTable(ByVal workbook As Object, ByVal rangeName As String) As String
    On Error GoTo Failed
    Dim data As Variant: data = ReadAnchoredResultTable(workbook, rangeName)
    If Not HasResultTableRows(data) Then Exit Function
    Dim column As Long: column = ResultColumn(data, "LoadCase")
    Dim row As Long, candidate As String
    For row = 2 To UBound(data, 1)
        candidate = Trim$(SafeText(data(row, column)))
        If Len(candidate) > 0 Then
            If rangeName <> "rngNDMSectionProperties" Or StrComp(candidate, "ALL", vbTextCompare) <> 0 Then
                FirstOutputCombinationInTable = candidate: Exit Function
            End If
        End If
    Next row
Failed:
End Function

' Читает собственную геометрию один раз блочными таблицами Results.
' Текущий Config, дополнительные JSON и внешний AutoCAD не участвуют;
' таблица границ обязательна для Contour; компактные Rectangle/Region хранятся без неё.
Public Function ReadSavedConcreteBoundaries(ByVal workbook As Object, Optional ByRef geometryData As Variant) As Object
    Dim data As Variant
    If IsArray(geometryData) Then data = geometryData Else data = ReadAnchoredResultTable(workbook, "rngNDMSectionGeometry")
    Dim boundaryData As Variant, anchor As Object, geometryAnchor As Object
    Set geometryAnchor = workbook.Names.Item("rngNDMSectionGeometry").RefersToRange
    On Error Resume Next
    Set anchor = workbook.Names.Item("rngNDMElementGeometry").RefersToRange
    On Error GoTo 0
    Dim saved As CCalculationSnapshot: Set saved = CurrentCalculationSnapshot(workbook)
    If Not saved Is Nothing Then
        Dim checked As Object
        If saved.TryCopyValidatedConcreteBoundaries(data, checked) Then
            Set ReadSavedConcreteBoundaries = checked
            Exit Function
        End If
    End If
    If Not anchor Is Nothing Or Not saved Is Nothing Then boundaryData = ReadAnchoredResultTable(workbook, "rngNDMElementGeometry")
    Set ReadSavedConcreteBoundaries = ReadConcreteBoundaryMap(boundaryData, data, "rngNDMElementGeometry", anchor, geometryAnchor)
End Function

' Проверяет явный маркер точной геометрии; компактная сетка не требует чтения большого файла НДС.
Public Function GeometryRequiresConcreteBoundaries(ByRef geometry As Variant) As Boolean
    If Not IsArray(geometry) Then Exit Function
    Dim column As Long: column = ResultColumn(geometry, "GeometryInterpretationStatus")
    Dim row As Long
    For row = 2 To UBound(geometry, 1)
        If IsError(geometry(row, column)) Or IsNull(geometry(row, column)) Then
            GeometryRequiresConcreteBoundaries = True: Exit Function
        End If
        If StrComp(CStr(geometry(row, column)), "Contour", vbTextCompare) = 0 Then GeometryRequiresConcreteBoundaries = True: Exit Function
    Next row
End Function

' Разбирает обязательные собственные границы из матриц Excel или файла.
' Contour требует точной связанной области; компактные Region/Rectangle не
' получают выдуманный контур. Возвращаемые координаты и инерции всегда в мм.
Public Function ReadConcreteBoundaryMap(ByRef boundaryData As Variant, ByRef geometryData As Variant, _
        Optional ByVal sourceReference As String = "rngNDMElementGeometry", _
        Optional ByVal boundaryAnchor As Object = Nothing, Optional ByVal geometryAnchor As Object = Nothing) As Object
    Dim result As Object: Set result = CreateObject("Scripting.Dictionary"): result.CompareMode = vbBinaryCompare
    Dim rows As Object: Set rows = CreateObject("Scripting.Dictionary"): rows.CompareMode = vbBinaryCompare
    Dim row As Long, id As String, kind As String, required As Object
    Set required = CreateObject("Scripting.Dictionary"): required.CompareMode = vbBinaryCompare
    If Not HasResultTableRows(geometryData) Then Err.Raise vbObjectError + 4383, "ReadConcreteBoundaryMap", "В снимке отсутствует основная геометрия элементов."
    Dim idColumn As Long: idColumn = ResultColumn(geometryData, "ElementID")
    Dim kindColumn As Long: kindColumn = ResultColumn(geometryData, "MaterialType")
    Dim shapeColumn As Long: shapeColumn = ResultColumn(geometryData, "GeometryInterpretationStatus")
    For row = 2 To UBound(geometryData, 1)
        id = SavedBoundaryText(geometryData, row, idColumn, "rngNDMSectionGeometry", geometryAnchor)
        If Len(id) > 0 Then
            If rows.Exists(id) Then SavedBoundaryError "Повторное имя элемента.", "rngNDMSectionGeometry", row, idColumn, geometryAnchor
            rows.Add id, row
            kind = SavedBoundaryText(geometryData, row, kindColumn, "rngNDMSectionGeometry", geometryAnchor)
            If StrComp(SavedBoundaryText(geometryData, row, shapeColumn, "rngNDMSectionGeometry", geometryAnchor), "Contour", vbTextCompare) = 0 Then
                If StrComp(kind, "Concrete", vbTextCompare) <> 0 Then SavedBoundaryError "Фактическая граница назначена не бетонному элементу.", "rngNDMSectionGeometry", row, kindColumn, geometryAnchor
                required.Add id, True
            End If
        End If
    Next row
    If Not IsArray(boundaryData) Then
        If required.Count > 0 Then SavedBoundaryError "Отсутствует обязательная таблица фактических границ бетонных элементов.", sourceReference, 1, 1, boundaryAnchor
        Set ReadConcreteBoundaryMap = result
        Exit Function
    End If
    Dim headers As Variant, column As Long
    headers = Array("RunID", "LoopID", "SegmentID", "Sequence", "LoopRole", "SegmentType", _
        "StartX", "StartY", "EndX", "EndY", "CenterX", "CenterY", "Radius", "SweepAngle, rad", "SourceID", "Comment", "LengthUnit (SectionXY)", "ElementID")
    If UBound(boundaryData, 2) <> 18 Then SavedBoundaryError "Неверная ширина таблицы фактических границ.", sourceReference, 1, 1, boundaryAnchor
    For column = 1 To 18
        If SavedBoundaryText(boundaryData, 1, column, sourceReference, boundaryAnchor) <> CStr(headers(column - 1)) Then _
            SavedBoundaryError "Неверный заголовок таблицы фактических границ.", sourceReference, 1, column, boundaryAnchor
    Next column
    Dim groups As Object: Set groups = CreateObject("Scripting.Dictionary"): groups.CompareMode = vbBinaryCompare
    Dim firstRows As Object: Set firstRows = CreateObject("Scripting.Dictionary"): firstRows.CompareMode = vbBinaryCompare
    Dim ids As Object: Set ids = CreateObject("Scripting.Dictionary"): ids.CompareMode = vbBinaryCompare
    Dim hashed As Object: Set hashed = ReadHashedConcreteBoundaries(boundaryData, geometryData, rows)
    Dim contours As CSectionContours, segmentID As String, loopID As String, role As String, owner As String, key As String, unitText As String
    Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double, cx As Double, cy As Double, radius As Double, sweep As Double, geometryRow As Long
    On Error GoTo InvalidContourRow
    For row = 2 To UBound(boundaryData, 1)
        owner = SavedBoundaryText(boundaryData, row, 18, sourceReference, boundaryAnchor)
        If Not required.Exists(owner) Then SavedBoundaryError "Граница не связана с существующим бетонным элементом типа Contour.", sourceReference, row, 18, boundaryAnchor
        geometryRow = CLng(rows.Item(owner))
        If SavedBoundaryRunID(boundaryData, row, sourceReference, boundaryAnchor) <> _
                SavedBoundaryRunID(geometryData, geometryRow, "rngNDMSectionGeometry", geometryAnchor) Then _
            SavedBoundaryError "Граница относится к другому расчётному снимку.", sourceReference, row, 1, boundaryAnchor
        If SavedBoundaryNumber(boundaryData, row, 4, sourceReference, boundaryAnchor) <> row - 1 Then _
            SavedBoundaryError "Нарушена последовательность сегментов фактической границы.", sourceReference, row, 4, boundaryAnchor
        If hashed.Exists(owner) Then GoTo NextSavedContourRow
        segmentID = SavedBoundaryText(boundaryData, row, 3, sourceReference, boundaryAnchor)
        loopID = SavedBoundaryText(boundaryData, row, 2, sourceReference, boundaryAnchor)
        role = SavedBoundaryText(boundaryData, row, 5, sourceReference, boundaryAnchor)
        If StrComp(role, "Outer", vbTextCompare) <> 0 And StrComp(role, "Opening", vbTextCompare) <> 0 Then _
            SavedBoundaryError "Роль фактической границы должна быть Outer или Opening.", sourceReference, row, 5, boundaryAnchor
        If Len(Trim$(segmentID)) = 0 Or Len(Trim$(loopID)) = 0 Then SavedBoundaryError "Не задано имя сегмента или замкнутого контура.", sourceReference, row, 3, boundaryAnchor
        key = CStr(Len(owner)) & ":" & owner & segmentID
        If ids.Exists(key) Then SavedBoundaryError "Повторный сегмент фактической границы.", sourceReference, row, 3, boundaryAnchor
        ids.Add key, True
        If Not groups.Exists(owner) Then
            Set contours = New CSectionContours: groups.Add owner, contours: firstRows.Add owner, row
        Else
            Set contours = groups.Item(owner)
        End If
        unitText = SavedBoundaryText(boundaryData, row, 17, sourceReference, boundaryAnchor)
        If Len(Trim$(unitText)) = 0 Then SavedBoundaryError "Не указана обязательная единица длины фактической границы.", sourceReference, row, 17, boundaryAnchor
        x1 = OutputLengthToInternalByUnit(SavedBoundaryNumber(boundaryData, row, 7, sourceReference, boundaryAnchor), unitText)
        y1 = OutputLengthToInternalByUnit(SavedBoundaryNumber(boundaryData, row, 8, sourceReference, boundaryAnchor), unitText)
        kind = SavedBoundaryText(boundaryData, row, 6, sourceReference, boundaryAnchor)
        Select Case kind
            Case "CONTOUR_LINE"
                x2 = OutputLengthToInternalByUnit(SavedBoundaryNumber(boundaryData, row, 9, sourceReference, boundaryAnchor), unitText)
                y2 = OutputLengthToInternalByUnit(SavedBoundaryNumber(boundaryData, row, 10, sourceReference, boundaryAnchor), unitText)
                contours.AddContourLine segmentID, x1, y1, x2, y2, vbNullString, loopID, role
            Case "CONTOUR_CIRCLE"
                radius = OutputLengthToInternalByUnit(SavedBoundaryNumber(boundaryData, row, 13, sourceReference, boundaryAnchor), unitText)
                contours.AddContourCircle segmentID, x1, y1, radius, vbNullString, loopID, role
            Case "CONTOUR_ARC"
                cx = OutputLengthToInternalByUnit(SavedBoundaryNumber(boundaryData, row, 11, sourceReference, boundaryAnchor), unitText)
                cy = OutputLengthToInternalByUnit(SavedBoundaryNumber(boundaryData, row, 12, sourceReference, boundaryAnchor), unitText)
                sweep = ReadContourArcSweep(boundaryData(row, 14), SavedBoundaryAddress(sourceReference, row, 14, boundaryAnchor))
                radius = Sqr((x1 - cx) ^ 2 + (y1 - cy) ^ 2)
                x2 = cx + (x1 - cx) * Cos(sweep) - (y1 - cy) * Sin(sweep)
                y2 = cy + (x1 - cx) * Sin(sweep) + (y1 - cy) * Cos(sweep)
                contours.AddContourArc segmentID, x1, y1, x2, y2, cx, cy, radius, sweep, vbNullString, loopID, role
            Case Else
                SavedBoundaryError "Неизвестный тип сегмента фактической границы.", sourceReference, row, 6, boundaryAnchor
        End Select
NextSavedContourRow:
    Next row
    On Error GoTo 0
    Dim item As Variant, region As CConcreteRegion, query As CSectionGeometryQuery: Set query = New CSectionGeometryQuery
    Dim loopNames As Object, actualLoopNames As Object
    Set loopNames = CreateObject("Scripting.Dictionary"): loopNames.CompareMode = vbBinaryCompare
    Set actualLoopNames = CreateObject("Scripting.Dictionary"): actualLoopNames.CompareMode = vbBinaryCompare
    Dim layout As TSavedBoundaryLayout, layoutReady As Boolean ' Один заголовок для всех элементов этого чтения; между действиями ничего не кэшируется.

    For Each item In required.Keys
        If hashed.Exists(CStr(item)) Then
            Set region = hashed.Item(CStr(item)): result.Add CStr(item), region
            mTestBoundaryHashReads = mTestBoundaryHashReads + 1
            GoTo NextSavedBoundaryElement
        End If
        If Not groups.Exists(CStr(item)) Then SavedBoundaryError "Для бетонного элемента " & CStr(item) & " отсутствует обязательная фактическая граница.", sourceReference, 1, 18, boundaryAnchor
        Set contours = groups.Item(CStr(item))
        On Error GoTo InvalidBoundary
        mTestBoundaryFullReads = mTestBoundaryFullReads + 1
        Set region = SavedConcreteRegion(contours, query, loopNames, actualLoopNames)
        If Not layoutReady Then
            ReadSavedBoundaryLayout geometryData, geometryAnchor, layout
            layoutReady = True
        End If
        ValidateSavedBoundaryProperties region, geometryData, CLng(rows.Item(CStr(item))), geometryAnchor, query, layout
        On Error GoTo 0
        result.Add CStr(item), region
NextSavedBoundaryElement:
    Next item
    Set ReadConcreteBoundaryMap = result
    Exit Function
InvalidBoundary:
    Dim reason As String: reason = Err.Description
    SavedBoundaryError "Некорректная фактическая граница элемента " & CStr(item) & ": " & reason, _
        sourceReference, CLng(firstRows.Item(CStr(item))), 18, boundaryAnchor
    Exit Function
InvalidContourRow:
    reason = Err.Description
    SavedBoundaryError reason, sourceReference, row, 1, boundaryAnchor
End Function

' Дата запуска до записи в Excel может иметь тип Date. Только эта колонка
' допускает Date и преобразует его числовое представление без текста локали;
' координаты, площади и остальные физические числа отвергают Date.
Private Function SavedBoundaryRunID(ByRef values As Variant, ByVal row As Long, ByVal reference As String, ByVal anchor As Object) As Double
    If VarType(values(row, 1)) = vbDate Then
        SavedBoundaryRunID = CDbl(values(row, 1))
    Else
        SavedBoundaryRunID = SavedBoundaryNumber(values, row, 1, reference, anchor)
    End If
End Function

' Читает строгое число из уже загруженной матрицы без CDbl текстовой локали.
' В файловом снимке адрес задаётся именем таблицы и относительной позицией.
Private Function SavedBoundaryNumber(ByRef values As Variant, ByVal row As Long, ByVal column As Long, _
        ByVal reference As String, ByVal anchor As Object) As Double
    Dim value As Variant: value = values(row, column)
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then GoTo InvalidNumber
    If VarType(value) = vbBoolean Or VarType(value) = vbDate Then GoTo InvalidNumber
    Dim parsed As Double
    If Not TryReadStrictDouble(value, parsed) Then GoTo InvalidNumber
    SavedBoundaryNumber = parsed
    Exit Function
InvalidNumber:
    SavedBoundaryError "Обязательное поле фактической геометрии содержит нечисловое значение.", reference, row, column, anchor
End Function

' Не скрывает ошибку формулы или Null под пустым ID либо ролью границы.
Private Function SavedBoundaryText(ByRef values As Variant, ByVal row As Long, ByVal column As Long, _
        ByVal reference As String, ByVal anchor As Object) As String
    If IsError(values(row, column)) Or IsNull(values(row, column)) Then _
        SavedBoundaryError "Поле фактической геометрии содержит ошибку.", reference, row, column, anchor
    SavedBoundaryText = CStr(values(row, column))
End Function

' Возвращает фактическую ячейку через именованный якорь, без жёсткой раскладки.
' Для матрицы файла сохраняется точная относительная строка и колонка.
Private Function SavedBoundaryAddress(ByVal reference As String, ByVal row As Long, ByVal column As Long, ByVal anchor As Object) As String
    If anchor Is Nothing Then
        SavedBoundaryAddress = reference & " [строка " & CStr(row) & ", столбец " & CStr(column) & "]"
    Else
        SavedBoundaryAddress = anchor.Parent.Name & "!" & anchor.Offset(row - 1, column - 1).Address(False, False)
    End If
End Function

' Выдаёт единую адресную диагностику обязательной геометрии до изменения Results.
Private Sub SavedBoundaryError(ByVal reason As String, ByVal reference As String, ByVal row As Long, ByVal column As Long, ByVal anchor As Object)
    Err.Raise vbObjectError + 4383, "ReadConcreteBoundaryMap", reason & " " & SavedBoundaryAddress(reference, row, column, anchor) & ". Нажмите «Сформировать сечение» или выберите исправный снимок."
End Sub

' Собирает область строгим общим валидатором, без нормализации повреждений.
' Сохранённые LoopID и роли сверяются с фактической вложенностью замкнутых
' границ; окружность разворачивается общим ядром в четыре точные дуги.
Private Function SavedConcreteRegion(ByVal contours As CSectionContours, ByVal query As CSectionGeometryQuery, _
        ByVal loops As Object, ByVal actualLoops As Object) As CConcreteRegion
    Dim edges() As TRegionEdge, sourceLoop() As String, sourceRole() As String, count As Long
    ReDim edges(1 To contours.Count * 4): ReDim sourceLoop(1 To contours.Count * 4): ReDim sourceRole(1 To contours.Count * 4)
    Dim i As Long, j As Long, circleEdges() As TRegionEdge, circleCount As Long, circleRegion As CConcreteRegion
    For i = 1 To contours.Count
        If contours.SegmentType(i) = "CONTOUR_CIRCLE" Then
            Set circleRegion = query.CircleRegion(contours.StartX(i), contours.StartY(i), contours.Radius(i))
            circleRegion.CopyEdges circleEdges, circleCount
            For j = 1 To circleCount
                count = count + 1: LSet edges(count) = circleEdges(j)
                sourceLoop(count) = contours.LoopID(i): sourceRole(count) = contours.LoopRole(i)
            Next j
        Else
            count = count + 1
            With edges(count)
                .X1 = contours.StartX(i): .Y1 = contours.StartY(i): .X2 = contours.EndX(i): .Y2 = contours.EndY(i)
                If contours.SegmentType(i) = "CONTOUR_ARC" Then
                    .CenterX = contours.CenterX(i): .CenterY = contours.CenterY(i): .Sweep = contours.SweepAngle(i)
                End If
            End With
            sourceLoop(count) = contours.LoopID(i): sourceRole(count) = contours.LoopRole(i)
        End If
    Next i
    Dim region As CConcreteRegion: Set region = query.BoundaryRegion(edges, count, "SavedConcreteBoundary")
    loops.RemoveAll: actualLoops.RemoveAll ' Эти рабочие словари принадлежат одному чтению и очищаются для каждого элемента; возвращаемая область от них не зависит.
    Dim actual() As TRegionEdge, actualCount As Long, loopID As Long, matched As Boolean, tolerance As Double
    region.CopyEdges actual, actualCount
    If actualCount <> count Then Err.Raise vbObjectError + 4383, "SavedConcreteRegion", "Число сегментов изменилось при проверке фактической границы."
    For i = 1 To count
        matched = False
        tolerance = GEOM_TOLERANCE + 0.000000000001 * GeomMax(GeomMax(Abs(edges(i).X1), Abs(edges(i).Y1)), GeomMax(Abs(edges(i).X2), Abs(edges(i).Y2)))
        For j = 1 To actualCount
            If SavedBoundarySameEdge(edges(i), actual(j), tolerance) Then
                loopID = actual(j).LoopID: matched = True
                Exit For
            End If
        Next j
        If Not matched Then Err.Raise vbObjectError + 4383, "SavedConcreteRegion", "Сегмент не принадлежит проверенной замкнутой границе."
        If (StrComp(sourceRole(i), "Opening", vbTextCompare) = 0) <> region.LoopIsOpening(loopID) Then _
            Err.Raise vbObjectError + 4383, "SavedConcreteRegion", "Роль Outer/Opening не соответствует вложенности фактической границы."
        If loops.Exists(sourceLoop(i)) Then
            If CLng(loops.Item(sourceLoop(i))) <> loopID Then Err.Raise vbObjectError + 4383, "SavedConcreteRegion", "Одним LoopID обозначены раздельные замкнутые границы."
        Else
            loops.Add sourceLoop(i), loopID
        End If
        If actualLoops.Exists(CStr(loopID)) Then
            If CStr(actualLoops.Item(CStr(loopID))) <> sourceLoop(i) Then Err.Raise vbObjectError + 4383, "SavedConcreteRegion", "Одна замкнутая граница разделена между разными LoopID."
        Else
            actualLoops.Add CStr(loopID), sourceLoop(i)
        End If
    Next i
    Set SavedConcreteRegion = region
End Function

' Узнаёт тот же аналитический сегмент после разрешённого изменения обхода.
' Одни совпавшие концы недостаточны: дополнительно совпадают центр и угол дуги.
Private Function SavedBoundarySameEdge(ByRef first As TRegionEdge, ByRef second As TRegionEdge, ByVal tolerance As Double) As Boolean
    Dim forward As Boolean, backward As Boolean
    forward = Abs(first.X1 - second.X1) <= tolerance And Abs(first.Y1 - second.Y1) <= tolerance And _
        Abs(first.X2 - second.X2) <= tolerance And Abs(first.Y2 - second.Y2) <= tolerance And Abs(first.Sweep - second.Sweep) <= 0.0000000001
    backward = Abs(first.X1 - second.X2) <= tolerance And Abs(first.Y1 - second.Y2) <= tolerance And _
        Abs(first.X2 - second.X1) <= tolerance And Abs(first.Y2 - second.Y1) <= tolerance And Abs(first.Sweep + second.Sweep) <= 0.0000000001
    If Not forward And Not backward Then Exit Function
    If Abs(first.Sweep) > 0.0000000001 Then
        If Abs(first.CenterX - second.CenterX) > tolerance Or Abs(first.CenterY - second.CenterY) > tolerance Then Exit Function
    End If
    SavedBoundarySameEdge = True
End Function

' Сверяет A, центр и собственные инерции общим геометрическим валидатором.
' Единицы и числовые поля проверяются до вызова; наружный обработчик
' добавляет фактический адрес границы без создания временной модели.
' Проверяет заголовок одной матрицы до чтения свойств её первого точного элемента.
' Адреса и требования совпадают с построчной проверкой; текущие единицы Config не используются.
Private Sub ReadSavedBoundaryLayout(ByRef data As Variant, ByVal anchor As Object, ByRef layout As TSavedBoundaryLayout)
    layout.XColumn = ResultColumn(data, "X")
    layout.YColumn = ResultColumn(data, "Y")
    layout.AreaColumn = ResultColumn(data, "Area")
    layout.IxColumn = ResultColumn(data, "LocalIx")
    layout.IyColumn = ResultColumn(data, "LocalIy")
    layout.IxyColumn = ResultColumn(data, "LocalIxy")
    layout.WidthColumn = ResultColumn(data, "Width")
    layout.HeightColumn = ResultColumn(data, "Height")
    layout.RotationColumn = ResultColumn(data, "Rotation")
    layout.LengthUnit = SavedBoundaryHeaderUnit(data, layout.XColumn, anchor)
    layout.AreaUnit = SavedBoundaryHeaderUnit(data, layout.AreaColumn, anchor)
    layout.InertiaUnit = SavedBoundaryHeaderUnit(data, layout.IxColumn, anchor)
    If Right$(layout.InertiaUnit, 1) <> "4" Then SavedBoundaryError "Собственная инерция требует единицы длины в четвёртой степени.", "rngNDMSectionGeometry", 1, layout.IxColumn, anchor
    Call SavedBoundaryHeaderUnit(data, layout.YColumn, anchor, layout.LengthUnit)
    Call SavedBoundaryHeaderUnit(data, layout.WidthColumn, anchor, layout.LengthUnit)
    Call SavedBoundaryHeaderUnit(data, layout.HeightColumn, anchor, layout.LengthUnit)
    Call SavedBoundaryHeaderUnit(data, layout.IyColumn, anchor, layout.InertiaUnit)
    Call SavedBoundaryHeaderUnit(data, layout.IxyColumn, anchor, layout.InertiaUnit)
    Call SavedBoundaryHeaderUnit(data, layout.RotationColumn, anchor, "rad")
End Sub

' Сверяет каждую строку с точной границей тем же общим валидатором.
' Общая шапка подготовлена для этого чтения; данные и числовые допуски не кэшируются.
Private Sub ValidateSavedBoundaryProperties(ByVal region As CConcreteRegion, ByRef data As Variant, ByVal row As Long, ByVal anchor As Object, _
        ByVal query As CSectionGeometryQuery, ByRef layout As TSavedBoundaryLayout)
    If SavedBoundaryNumber(data, row, layout.WidthColumn, "rngNDMSectionGeometry", anchor) < 0# Then _
        SavedBoundaryError "Ширина точного элемента не может быть отрицательной.", "rngNDMSectionGeometry", row, layout.WidthColumn, anchor
    If SavedBoundaryNumber(data, row, layout.HeightColumn, "rngNDMSectionGeometry", anchor) < 0# Then _
        SavedBoundaryError "Высота точного элемента не может быть отрицательной.", "rngNDMSectionGeometry", row, layout.HeightColumn, anchor
    Call SavedBoundaryNumber(data, row, layout.RotationColumn, "rngNDMSectionGeometry", anchor)
    query.ValidateBoundaryProperties region, _
        OutputLengthToInternalByUnit(SavedBoundaryNumber(data, row, layout.XColumn, "rngNDMSectionGeometry", anchor), layout.LengthUnit), _
        OutputLengthToInternalByUnit(SavedBoundaryNumber(data, row, layout.YColumn, "rngNDMSectionGeometry", anchor), layout.LengthUnit), _
        OutputAreaToInternalByUnit(SavedBoundaryNumber(data, row, layout.AreaColumn, "rngNDMSectionGeometry", anchor), layout.AreaUnit), _
        OutputFourthPowerLengthToInternalByUnit(SavedBoundaryNumber(data, row, layout.IxColumn, "rngNDMSectionGeometry", anchor), layout.InertiaUnit), _
        OutputFourthPowerLengthToInternalByUnit(SavedBoundaryNumber(data, row, layout.IyColumn, "rngNDMSectionGeometry", anchor), layout.InertiaUnit), _
        OutputFourthPowerLengthToInternalByUnit(SavedBoundaryNumber(data, row, layout.IxyColumn, "rngNDMSectionGeometry", anchor), layout.InertiaUnit)
    region.RememberPreparedProperties _
        OutputLengthToInternalByUnit(SavedBoundaryNumber(data, row, layout.XColumn, "rngNDMSectionGeometry", anchor), layout.LengthUnit), _
        OutputLengthToInternalByUnit(SavedBoundaryNumber(data, row, layout.YColumn, "rngNDMSectionGeometry", anchor), layout.LengthUnit), _
        OutputAreaToInternalByUnit(SavedBoundaryNumber(data, row, layout.AreaColumn, "rngNDMSectionGeometry", anchor), layout.AreaUnit), _
        OutputFourthPowerLengthToInternalByUnit(SavedBoundaryNumber(data, row, layout.IxColumn, "rngNDMSectionGeometry", anchor), layout.InertiaUnit), _
        OutputFourthPowerLengthToInternalByUnit(SavedBoundaryNumber(data, row, layout.IyColumn, "rngNDMSectionGeometry", anchor), layout.InertiaUnit), _
        OutputFourthPowerLengthToInternalByUnit(SavedBoundaryNumber(data, row, layout.IxyColumn, "rngNDMSectionGeometry", anchor), layout.InertiaUnit)
End Sub

' Точная геометрия запрещает подстановку текущих единиц при пустой шапке.
' Парные координаты и компоненты одного тензора имеют единый формат writer;
' отдельный выбор единицы площади остаётся допустимым.
Private Function SavedBoundaryHeaderUnit(ByRef data As Variant, ByVal column As Long, ByVal anchor As Object, Optional ByVal expected As String = vbNullString) As String
    Dim unit As String: unit = ResultHeaderUnit(data, column, vbNullString)
    If Len(unit) = 0 Then SavedBoundaryError "Не указана единица обязательного свойства точного элемента.", "rngNDMSectionGeometry", 1, column, anchor
    If Len(expected) > 0 Then
        If StrComp(unit, expected, vbTextCompare) <> 0 Then SavedBoundaryError "Несогласованные единицы свойств точного элемента.", "rngNDMSectionGeometry", 1, column, anchor
    End If
    SavedBoundaryHeaderUnit = unit
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

' Нулевые Width/Height допустимы только у точного непрямоугольного Contour:
' его настоящая граница обязательна в отдельной таблице. Размеры
' прямоугольника остаются строго положительными, отрицательные запрещены всем.
Private Function ReadConcreteSnapshotDimension(ByRef data As Variant, ByVal anchor As Object, ByVal row As Long, ByVal column As Long, ByVal shape As String) As Double
    Dim value As Double: value = ReadGeometryNumber(data, anchor, row, column)
    If value < 0# Or (value = 0# And StrComp(shape, "Contour", vbTextCompare) <> 0) Then _
        Err.Raise vbObjectError + 4355, "ReadSectionGeometryFromResults", "Неверный размер сохранённого бетонного элемента: " & _
            anchor.Parent.Name & "!" & anchor.Offset(row - 1, column - 1).Address(False, False) & ". Нажмите «Сформировать сечение»."
    ReadConcreteSnapshotDimension = value
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
    Dim snapshotRun As Variant, geometryData As Variant
    geometryData = ReadAnchoredResultTable(workbook, "rngNDMSectionGeometry") ' Номер запуска берётся из того же источника Excel/File, что и контуры; справочный Results в File не участвует.
    If HasResultTableRows(geometryData) Then
        If CStr(geometryData(1, 1)) = "RunID" Then snapshotRun = geometryData(2, 1)
    End If
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

' Возвращает источник уже сформированного сечения из свойств выбранного файла либо Results.
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


' Читает сохранённый целочисленный ID материала без округления дроби или группировки разрядов.
' MaterialID хранится целым Long; повреждённое поле указывает свою ячейку ExtraData.
Private Function ReadSnapshotMaterialID(ByVal text As String, ByVal anchor As Object, _
        ByVal row As Long, ByVal column As Long) As Long
    Dim parsed As Double
    If Not TryReadStrictDouble(text, parsed) Then GoTo InvalidNumber
    If parsed < -2147483648# Or parsed > 2147483647# Then GoTo InvalidNumber
    If Fix(parsed) <> parsed Then GoTo InvalidNumber
    ReadSnapshotMaterialID = CLng(parsed)
    Exit Function
InvalidNumber:
    Err.Raise vbObjectError + 4382, "ReadSectionGeometryFromResults", _
        "В сохранённых сведениях сечения MaterialID должен быть целым числом без разделителей разрядов: " & _
        anchor.Worksheet.Name & "!" & anchor.Offset(row - 1, column - 1).Address(False, False) & _
        ". Заново сформируйте сечение для восстановления этих сведений."
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
    Dim parsed As Double
    If Not TryReadStrictDouble(value, parsed) Then GoTo InvalidNumber
    On Error GoTo InvalidNumber
    ReadGeometryNumber = parsed
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

' Читает таблицу из единого источника текущего режима: выбранного файла либо Results.
' В File ни одна геометрическая таблица не дополняется ячейками листа.
' Для расчёта достаточно общих таблиц, для схемы/экспорта добавляются записи сочетаний.
Public Function ReadAnchoredResultTable(ByVal workbook As Object, ByVal rangeName As String) As Variant
    Dim snapshot As CCalculationSnapshot: Set snapshot = CurrentCalculationSnapshot(workbook)
    If Not snapshot Is Nothing Then
        If rangeName = "rngNDMElementGeometry" And Not snapshot.HasTable(rangeName) Then Exit Function ' Старый Adaptive-снимок может не иметь необязательной матрицы границ; Contour проверяется snapshot.Validate.
        If workbook Is mSnapshotReadWorkbook And mSnapshotReadGeometryOnly Then
            ReadAnchoredResultTable = snapshot.Table(rangeName)
        Else
            ReadAnchoredResultTable = snapshot.ResultTable(rangeName)
        End If
        Exit Function
    End If
    ReadAnchoredResultTable = ReadPhysicalResultTable(workbook, rangeName)
End Function

' Читает только физические ячейки. Используется при сборке полного снимка,
' чтобы загрузка LC-дополнения не обращалась рекурсивно к самой себе.
Public Function ReadPhysicalResultTable(ByVal workbook As Object, ByVal rangeName As String) As Variant
    Dim useReadTables As Boolean
    useReadTables = Not mSnapshotReadTables Is Nothing And workbook Is mSnapshotReadWorkbook
    If useReadTables Then
        If mSnapshotReadTables.Exists(rangeName) Then
            ReadPhysicalResultTable = mSnapshotReadTables.Item(rangeName)
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

    ReadPhysicalResultTable = anchor.Resize(rowCount, columnCount).Value2
    If useReadTables Then
        mSnapshotReadTables.Add rangeName, ReadPhysicalResultTable
        mSnapshotTableReadCount = mSnapshotTableReadCount + 1
    End If
End Function

' Выбирает сохранённое сочетание или только перечень состояний. Внешний файл
' проверяется так же, как при полном чтении; новые НДС и импорт не выполняются.
Public Function ReadElementResultTable(ByVal workbook As Object, Optional ByVal combinationID As String = "", _
        Optional ByVal stateType As String = "", Optional ByVal indexOnly As Boolean = False) As Variant
    Dim snapshot As CCalculationSnapshot: Set snapshot = CurrentCalculationSnapshot(workbook)
    If Not snapshot Is Nothing Then ReadElementResultTable = snapshot.SelectElementResults(combinationID, stateType, indexOnly): Exit Function
    Dim data As Variant: data = ReadAnchoredResultTable(workbook, "rngNDMElementResults")
    If Len(combinationID) = 0 And Len(stateType) = 0 And Not indexOnly Then ReadElementResultTable = data: Exit Function
    Dim rows As Long, i As Long, j As Long, used As Long, identity As String, seen As Object
    Dim caseColumn As Long: caseColumn = ResultColumn(data, "LoadCase")
    Dim stateColumn As Long: stateColumn = ResultColumn(data, "StateType")
    Set seen = CreateObject("Scripting.Dictionary"): seen.CompareMode = vbTextCompare
    For i = 2 To UBound(data, 1)
        If (Len(combinationID) = 0 Or StrComp(CStr(data(i, caseColumn)), combinationID, vbTextCompare) = 0) And _
                (Len(stateType) = 0 Or StrComp(CStr(data(i, stateColumn)), stateType, vbTextCompare) = 0) Then
            If indexOnly Then
                identity = CStr(data(i, caseColumn)) & Chr$(30) & CStr(data(i, stateColumn))
                If Not seen.Exists(identity) Then seen.Add identity, i: rows = rows + 1
            Else
                rows = rows + 1
            End If
        End If
    Next i
    Dim selected() As Variant: ReDim selected(1 To rows + 1, 1 To UBound(data, 2))
    For j = 1 To UBound(data, 2): selected(1, j) = data(1, j): Next j
    used = 1
    For i = 2 To UBound(data, 1)
        If (Len(combinationID) = 0 Or StrComp(CStr(data(i, caseColumn)), combinationID, vbTextCompare) = 0) And _
                (Len(stateType) = 0 Or StrComp(CStr(data(i, stateColumn)), stateType, vbTextCompare) = 0) Then
            If indexOnly Then
                identity = CStr(data(i, caseColumn)) & Chr$(30) & CStr(data(i, stateColumn))
                If seen.Exists(identity) Then
                    If CLng(seen.Item(identity)) = i Then
                        used = used + 1: For j = 1 To UBound(data, 2): selected(used, j) = data(i, j): Next j
                    End If
                End If
            Else
                used = used + 1: For j = 1 To UBound(data, 2): selected(used, j) = data(i, j): Next j
            End If
        End If
    Next i
    ReadElementResultTable = selected
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
    Const INITIAL_BLOCK_SIZE As Long = 1024
    Const MAX_BLOCK_SIZE As Long = 32768
    Dim available As Long, offset As Long, count As Long, i As Long, blockSize As Long
    blockSize = INITIAL_BLOCK_SIZE ' Начальный блок ограничен размером малой таблицы; для больших таблиц размер блока удваивается.
    If columns Then
        available = anchor.Worksheet.Columns.Count - anchor.Column + 1
    Else
        available = anchor.Worksheet.Rows.Count - anchor.Row + 1
    End If
    If maxCount > 0 And maxCount < available Then available = maxCount
    Dim block As Variant, value As Variant
    Do While offset < available
        count = available - offset
        If count > blockSize Then count = blockSize
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
        If blockSize < MAX_BLOCK_SIZE Then blockSize = blockSize * 2
    Loop
End Function

' Начинает одно чтение согласованного снимка: каждая таблица получается один раз.
Public Sub BeginSectionSnapshotRead(ByVal workbook As Object, Optional ByVal geometryOnly As Boolean = False)
    BeginCalculationSnapshotRead workbook
    Set mSnapshotReadTables = CreateObject("Scripting.Dictionary")
    Set mSnapshotReadWorkbook = workbook
    mSnapshotReadGeometryOnly = geometryOnly
    mSnapshotTableReadCount = 0
End Sub

' Освобождает массивы снимка после чтения или ошибки, сохраняя счётчик обращений.
Public Sub EndSectionSnapshotRead()
    Set mSnapshotReadTables = Nothing: Set mSnapshotReadWorkbook = Nothing
    mSnapshotReadGeometryOnly = False
    EndCalculationSnapshotRead
End Sub


' Читает обязательную геометрию для расчёта. Необязательные сведения и
' аннотации добавляются только по запросу; генератор и AutoCAD не вызываются.
Public Function ReadFormedSectionFromResults(ByVal workbook As Object, Optional ByVal includeOptionalData As Boolean = False, _
        Optional ByVal preservedWriter As CNDMResultsWriter = Nothing) As CSectionModel
    On Error GoTo Failed
    BeginSectionSnapshotRead workbook, True ' Модель использует только общую геометрию, без загрузки массивов расчётных состояний.
    Dim savedSnapshot As CCalculationSnapshot: Set savedSnapshot = CurrentCalculationSnapshot(workbook) ' Проверяем файл до опционального чтения подписи источника: его ошибку нельзя выдать за пустой Results.
    Dim source As String: source = ResultsGeometrySource(workbook)
    If Len(source) = 0 Then Err.Raise vbObjectError + 4380, "ReadFormedSectionFromResults", _
        "В Results нет сформированного сечения. Сначала нажмите «Сформировать сечение»."
    Dim model As CSectionModel: Set model = ReadSectionGeometryFromResults(workbook, source, includeOptionalData)
    If includeOptionalData Then
        model.GeometryDiagnosticsText = ReadLoadCasePropertyText(workbook, "ALL", "Geometry.RecipeDiagnostics")
        RestoreModelAnnotations workbook, model
    End If
    If Not preservedWriter Is Nothing Then preservedWriter.PreparePreservedSnapshotSource workbook ' Повторно использует уже прочитанные таблицы; после расчёта файл собирается только в памяти.
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

' Получает номера непустых строк готовой сводки для подсветки статусов.
' Источником служит сохранённый ID, а не повторная загрузка сочетаний Config.
Public Function ResultSnapshotStatusRows(ByVal target As Object, ByVal headerRows As Long, _
        Optional ByRef preparedValues As Variant, Optional ByVal headerRowsInValues As Long = 0) As Variant
    Dim count As Long, values As Variant
    If IsMissing(preparedValues) Then
        count = target.Rows.Count - headerRows
        If count < 1 Then Exit Function
        values = target.Offset(headerRows, 0).Resize(count, 3).Value2
    Else
        values = preparedValues: count = UBound(values, 1) - headerRowsInValues
    End If
    If count < 1 Then Exit Function
    Dim rows() As Long, i As Long, used As Long, sourceRow As Long: ReDim rows(1 To count)
    For i = 1 To count
        sourceRow = i: If Not IsMissing(preparedValues) Then sourceRow = i + headerRowsInValues
        If Len(CStr(values(sourceRow, 1))) > 0 Or Len(CStr(values(sourceRow, 2))) > 0 Or Len(CStr(values(sourceRow, 3))) > 0 Then
            used = used + 1: rows(used) = headerRows + i
        End If
    Next i
    If used = 0 Then Exit Function
    ReDim Preserve rows(1 To used): ResultSnapshotStatusRows = rows
End Function

' Строго распознаёт сохранённый Boolean без Excel/AutoCAD и без подстановки False при повреждении.
' Поддерживает пары True/False, Yes/No, Да/Нет, 1/0 и тип Boolean.
Public Function TryReadSnapshotBoolean(ByVal value As Variant, ByRef parsed As Boolean) As Boolean
    parsed = False
    On Error GoTo InvalidValue
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Or IsObject(value) Or IsArray(value) Then Exit Function
    If VarType(value) = vbBoolean Then
        parsed = CBool(value)
    Else
        Select Case LCase$(Trim$(CStr(value)))
            Case "true", "yes", "да", "1": parsed = True
            Case "false", "no", "нет", "0": parsed = False
            Case Else: Exit Function
        End Select
    End If
    TryReadSnapshotBoolean = True
InvalidValue:
End Function

' Проверяет семь внешних кодов сохранённого результата; неизвестная строка не становится CalcErr.
' Возвращает каноническое написание; адрес и ошибку чтения добавляет вызывающий адаптер.
Public Function TryReadSnapshotStatus(ByVal value As Variant, ByRef parsed As String) As Boolean
    parsed = vbNullString
    On Error GoTo InvalidValue
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Or IsObject(value) Or IsArray(value) Then Exit Function
    Select Case LCase$(Trim$(CStr(value)))
        Case "ok": parsed = "OK"
        Case "fail": parsed = "FAIL"
        Case "basefail": parsed = "BaseFail"
        Case "numfail": parsed = "NumFail"
        Case "inputerr": parsed = "InputErr"
        Case "calcerr": parsed = "CalcErr"
        Case "n/a": parsed = "N/A"
        Case Else: Exit Function
    End Select
    TryReadSnapshotStatus = True
InvalidValue:
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Показывает число блочных чтений и освобождение таблиц после операции.
Public Function TestSectionSnapshotReadDiagnostics() As String
    TestSectionSnapshotReadDiagnostics = "tables=" & CStr(mSnapshotTableReadCount) & ";released=" & _
        CStr(mSnapshotReadTables Is Nothing And mSnapshotReadWorkbook Is Nothing)
End Function


' Подтверждает только целиком неизменённые записи, выпущенные штатным writer-ом.
' При отсутствии или ошибке JSON/хеша передаёт запись полному разбору без подмены геометрии.
Private Function ReadHashedConcreteBoundaries(ByRef boundaries As Variant, ByRef geometry As Variant, ByVal geometryRows As Object) As Object
    Dim result As Object: Set result = CreateObject("Scripting.Dictionary"): result.CompareMode = vbBinaryCompare
    Set ReadHashedConcreteBoundaries = result
    If UBound(boundaries, 1) < 2 Then Exit Function
    Dim extraColumn As Long: extraColumn = ResultColumn(geometry, "ExtraData", False)
    If extraColumn = 0 Then Exit Function
    Dim first As Object, last As Object, counts As Object
    Set first = CreateObject("Scripting.Dictionary"): Set last = CreateObject("Scripting.Dictionary"): Set counts = CreateObject("Scripting.Dictionary")
    Dim row As Long, owner As String, key As Variant
    On Error GoTo Untrusted
    For row = 2 To UBound(boundaries, 1)
        owner = CStr(boundaries(row, 18))
        If Not first.Exists(owner) Then first.Add owner, row: counts.Add owner, 0
        last(owner) = row: counts(owner) = CLng(counts(owner)) + 1
    Next row
    Dim checksum As CFileChecksum: Set checksum = New CFileChecksum
    checksum.PrepareGeometryHashHeaders geometry, boundaries
    Dim region As CConcreteRegion, query As CSectionGeometryQuery: Set query = New CSectionGeometryQuery
    Dim layout As TSavedBoundaryLayout, layoutReady As Boolean
    For Each key In first.Keys
        owner = CStr(key)
        If geometryRows.Exists(owner) And CLng(counts(owner)) = CLng(last(owner)) - CLng(first(owner)) + 1 Then
            Set region = TryReadHashedBoundary(boundaries, geometry, CLng(geometryRows.Item(owner)), _
                CLng(first(owner)), CLng(last(owner)), extraColumn, checksum, query, layout, layoutReady)
            If Not region Is Nothing Then result.Add owner, region
        End If
    Next key
    Exit Function
Untrusted:
    Err.Clear ' Основной читатель выдаст адресную ошибку, если обязательные данные повреждены.
End Function

' Вычисляет контрольную сумму только по скалярам и при совпадении собирает отрезки/дуги без геометрического валидатора.
' Запись неподдерживаемого вида или без контрольной суммы проходит полный разбор.
Private Function TryReadHashedBoundary(ByRef boundaries As Variant, ByRef geometry As Variant, ByVal geometryRow As Long, _
        ByVal firstRow As Long, ByVal lastRow As Long, ByVal extraColumn As Long, ByVal checksum As CFileChecksum, _
        ByVal query As CSectionGeometryQuery, ByRef layout As TSavedBoundaryLayout, ByRef layoutReady As Boolean) As CConcreteRegion
    On Error GoTo Untrusted
    Dim digest As String: digest = SavedGeometryHashText(CStr(geometry(geometryRow, extraColumn)))
    If Len(digest) <> 16 Then Exit Function
    If StrComp(digest, checksum.PreparedGeometryHash(geometry, geometryRow, boundaries, firstRow, lastRow), vbBinaryCompare) <> 0 Then Exit Function
    Dim edges() As TRegionEdge, count As Long, row As Long, index As Long, kind As String, factor As Double, unit As String
    count = lastRow - firstRow + 1: ReDim edges(1 To count)
    For row = firstRow To lastRow
        index = row - firstRow + 1: unit = CStr(boundaries(row, 17))
        factor = OutputLengthToInternalByUnit(1#, unit)
        With edges(index)
            .X1 = CDbl(boundaries(row, 7)) * factor: .Y1 = CDbl(boundaries(row, 8)) * factor
            kind = CStr(boundaries(row, 6))
            Select Case kind
                Case "CONTOUR_LINE"
                    .X2 = CDbl(boundaries(row, 9)) * factor: .Y2 = CDbl(boundaries(row, 10)) * factor
                Case "CONTOUR_ARC"
                    .CenterX = CDbl(boundaries(row, 11)) * factor: .CenterY = CDbl(boundaries(row, 12)) * factor
                    .Sweep = ReadContourArcSweep(boundaries(row, 14), "rngNDMElementGeometry") ' Общий читатель проверяет угол дуги и возвращает его в радианах.
                    .X2 = .CenterX + (.X1 - .CenterX) * Cos(.Sweep) - (.Y1 - .CenterY) * Sin(.Sweep)
                    .Y2 = .CenterY + (.X1 - .CenterX) * Sin(.Sweep) + (.Y1 - .CenterY) * Cos(.Sweep)
                Case Else: Exit Function ' Writer выпускает только линии и дуги; внешние окружности проходят полный разбор.
            End Select
        End With
    Next row
    Dim region As CConcreteRegion: Set region = query.RestoreHashedBoundary(edges, count)
    If Not layoutReady Then ReadSavedBoundaryLayout geometry, Nothing, layout: layoutReady = True
    region.RememberPreparedProperties _
        OutputLengthToInternalByUnit(CDbl(geometry(geometryRow, layout.XColumn)), layout.LengthUnit), _
        OutputLengthToInternalByUnit(CDbl(geometry(geometryRow, layout.YColumn)), layout.LengthUnit), _
        OutputAreaToInternalByUnit(CDbl(geometry(geometryRow, layout.AreaColumn)), layout.AreaUnit), _
        OutputFourthPowerLengthToInternalByUnit(CDbl(geometry(geometryRow, layout.IxColumn)), layout.InertiaUnit), _
        OutputFourthPowerLengthToInternalByUnit(CDbl(geometry(geometryRow, layout.IyColumn)), layout.InertiaUnit), _
        OutputFourthPowerLengthToInternalByUnit(CDbl(geometry(geometryRow, layout.IxyColumn)), layout.InertiaUnit)
    region.RememberGeometryHash digest
    Set TryReadHashedBoundary = region
    Exit Function
Untrusted:
    Err.Clear
End Function

' Извлекает только каноническую шестнадцатеричную подпись writer-а, не создавая словарь для каждого элемента.
' Необычное оформление или отсутствующий маркер выбирает полный путь; остальные необязательные поля не интерпретируются.
Private Function SavedGeometryHashText(ByVal text As String) As String
    Dim marker As String: marker = Chr$(34) & "GeometryHash" & Chr$(34) & ":" & Chr$(34)
    Dim prefix As String: prefix = "{" & Chr$(34) & "schema" & Chr$(34) & ":" & Chr$(34) & "1" & Chr$(34)
    If Left$(text, Len(prefix)) <> prefix Or Right$(text, 1) <> "}" Then Exit Function
    Dim position As Long: position = InStr(1, text, marker, vbBinaryCompare)
    If position = 0 Then Exit Function
    If InStr(position + Len(marker), text, marker, vbBinaryCompare) > 0 Then Exit Function
    position = position + Len(marker)
    If Mid$(text, position + 16, 1) <> Chr$(34) Then Exit Function
    Dim value As String: value = Mid$(text, position, 16)
    Dim i As Long
    For i = 1 To 16
        If InStr(1, "0123456789ABCDEF", Mid$(value, i, 1), vbBinaryCompare) = 0 Then Exit Function
    Next i
    SavedGeometryHashText = value
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает независимые счётчики выбора полного/хеш-пути; reset не меняет рабочие данные.
Public Function TestSavedBoundaryReadCounts(Optional ByVal reset As Boolean = False) As String
    TestSavedBoundaryReadCounts = CStr(mTestBoundaryFullReads) & "/" & CStr(mTestBoundaryHashReads)
    If reset Then mTestBoundaryFullReads = 0: mTestBoundaryHashReads = 0
End Function
