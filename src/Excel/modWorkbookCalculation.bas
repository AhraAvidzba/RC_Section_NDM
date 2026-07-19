Attribute VB_Name = "modWorkbookCalculation"
Option Explicit

Public Sub RunSectionCalculation()
    On Error GoTo Failed
    Dim message As String
    message = RunSectionCalculationForWorkbook(ThisWorkbook, True)
    MsgBox message, vbInformation, "RC Section NDM"
    Exit Sub

Failed:
    MsgBox "Расчет не выполнен: " & Err.Description, vbExclamation, "RC Section NDM"
End Sub

Public Sub ClearSectionResults()
    On Error GoTo Failed
    ClearSectionResultsForWorkbook ThisWorkbook
    MsgBox "Результаты и диагностика очищены. Исходные данные не изменены.", vbInformation, "RC Section NDM"
    Exit Sub

Failed:
    MsgBox "Не удалось очистить результаты: " & Err.Description, vbExclamation, "RC Section NDM"
End Sub

Public Function RunSectionCalculationForWorkbook(ByVal workbook As Object, Optional ByVal showMessages As Boolean = False) As String
    If workbook Is Nothing Then Err.Raise vbObjectError + 4100, "RunSectionCalculationForWorkbook", "Workbook is missing."

    ClearSectionResultsForWorkbook workbook

    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook workbook

    Dim geometry As ISectionGeometry
    Set geometry = ReadGeometry(workbook, settings)

    Dim mesh As CFiberMeshBuilder
    Set mesh = New CFiberMeshBuilder
    mesh.BuildMesh geometry, settings.GetDouble("Mesh.StepX", 10#), settings.GetDouble("Mesh.StepY", 10#), 1, _
        MeshBoundarySubdivisions(settings)

    Dim rebars As CRebarLayout
    Set rebars = ReadRebars(workbook, geometry, settings)

    Dim concrete As CConcreteDiagramMaterial
    Set concrete = New CConcreteDiagramMaterial
    concrete.InitializeFromSettings settings

    Dim steel As CSteelDiagramMaterial
    Set steel = New CSteelDiagramMaterial
    steel.InitializeFromSettings settings


    Dim batch As CBatchSectionCalculator
    Set batch = New CBatchSectionCalculator
    batch.Initialize mesh, rebars, concrete, steel
    batch.ApplySettings settings

    Dim reader As CLoadCombinationReader
    Set reader = New CLoadCombinationReader
    reader.LoadFromWorkbook workbook, batch
    If batch.Count = 0 Then Err.Raise vbObjectError + 4101, "RunSectionCalculationForWorkbook", "Не задано ни одного сочетания нагрузок."

    batch.Execute

    Dim summaryWriter As CBatchResultWriter
    Set summaryWriter = New CBatchResultWriter
    summaryWriter.WriteSummary workbook, batch

    WriteFirstCombinationResults workbook, mesh, rebars, concrete, steel, settings

    RunSectionCalculationForWorkbook = "Расчет завершен. Обработано сочетаний: " & CStr(batch.Count) & _
        ". Определяющее сочетание: " & batch.GoverningCombinationID
End Function

Public Sub ClearSectionResultsForWorkbook(ByVal workbook As Object)
    If workbook Is Nothing Then Err.Raise vbObjectError + 4110, "ClearSectionResultsForWorkbook", "Workbook is missing."
    ClearResultRange workbook.Names.Item("rngResultSection").RefersToRange
    ClearPlainRange workbook.Names.Item("rngSystemDiagnostics").RefersToRange
    workbook.Worksheets.Item("System").Range("A92:M125").ClearContents
End Sub

Private Sub ClearResultRange(ByVal target As Object)
    Dim rowIndex As Long
    Dim colIndex As Variant
    For rowIndex = 1 To target.Rows.Count
        For Each colIndex In Array(1, 5, 9, 13)
            If Not target.Cells.Item(rowIndex, CLng(colIndex)).MergeCells Then
                target.Cells.Item(rowIndex, CLng(colIndex)).ClearContents
            End If
        Next colIndex
    Next rowIndex
End Sub

Private Sub ClearPlainRange(ByVal target As Object)
    Dim rowIndex As Long
    Dim colIndex As Long
    For rowIndex = 1 To target.Rows.Count
        For colIndex = 1 To target.Columns.Count
            target.Cells.Item(rowIndex, colIndex).ClearContents
        Next colIndex
    Next rowIndex
End Sub

Private Function ReadGeometry(ByVal workbook As Object, ByVal settings As CSystemSettingsReader) As ISectionGeometry
    Dim inputRange As Object
    Set inputRange = workbook.Names.Item("rngMainInput").RefersToRange

    Dim geometryType As String
    geometryType = settings.GetString("Geometry.Type", "RoundedRectangle")
    If Len(geometryType) = 0 Then geometryType = "RoundedRectangle"

    If StrComp(geometryType, "Circle", vbTextCompare) = 0 Then
        Dim circleGeom As CGeometryCircle
        Set circleGeom = New CGeometryCircle
        circleGeom.InitializeByDiameter settings.GetDouble("Circle.Diameter", 300#), _
            settings.GetDouble("Circle.CenterX", 0#), _
            settings.GetDouble("Circle.CenterY", 0#)
        Set ReadGeometry = circleGeom
    ElseIf StrComp(geometryType, "RoundedRectangle", vbTextCompare) = 0 Then
        Dim rect As CGeometryRoundedRectangle
        Set rect = New CGeometryRoundedRectangle
        rect.Initialize ReadRequiredDouble(inputRange.Cells.Item(3, 5).Value2, "Width"), _
            ReadRequiredDouble(inputRange.Cells.Item(4, 5).Value2, "Height"), _
            ReadOptionalDouble(inputRange.Cells.Item(5, 5).Value2, 0#), _
            ReadOptionalDouble(inputRange.Cells.Item(6, 5).Value2, 0#), _
            ReadOptionalDouble(inputRange.Cells.Item(7, 5).Value2, 0#), _
            ReadOptionalDouble(inputRange.Cells.Item(8, 5).Value2, 0#)
        Set ReadGeometry = rect
    Else
        Err.Raise vbObjectError + 4102, "ReadGeometry", "Unsupported section type: " & geometryType
    End If
End Function

Private Function MeshBoundarySubdivisions(ByVal settings As CSystemSettingsReader) As Long
    MeshBoundarySubdivisions = settings.GetLong("Mesh.BoundarySubdivisions", 1)
    If MeshBoundarySubdivisions < 1 Then MeshBoundarySubdivisions = 1
End Function

Private Function ReadRebars(ByVal workbook As Object, ByVal geometry As ISectionGeometry, ByVal settings As CSystemSettingsReader) As CRebarLayout
    Dim inputRange As Object
    Set inputRange = workbook.Names.Item("rngMainInput").RefersToRange

    Dim geometryType As String
    geometryType = settings.GetString("Geometry.Type", "RoundedRectangle")
    If StrComp(geometryType, "Circle", vbTextCompare) <> 0 Then
        Err.Raise vbObjectError + 4103, "ReadRebars", "Автоматическая расстановка арматуры сейчас поддерживает только Circle."
    End If

    Dim builder As CCircleRebarLayoutBuilder
    Set builder = New CCircleRebarLayoutBuilder

    Dim layout As CRebarLayout
    Set layout = builder.Build( _
        settings.GetDouble("Circle.Diameter", 300#), _
        settings.GetDouble("Circle.CenterX", 0#), _
        settings.GetDouble("Circle.CenterY", 0#), _
        ReadRequiredDouble(inputRange.Cells.Item(14, 5).Value2, "Rebar.AxisDistance"), _
        CLng(ReadRequiredDouble(inputRange.Cells.Item(15, 5).Value2, "Rebar.Count")), _
        ReadRequiredDouble(inputRange.Cells.Item(16, 5).Value2, "Rebar.Diameter"), _
        Trim$(CStr(inputRange.Cells.Item(13, 5).Value2)))

    WriteGeneratedRebarTable workbook, layout
    Set ReadRebars = layout
End Function

Private Sub WriteGeneratedRebarTable(ByVal workbook As Object, ByVal layout As CRebarLayout)
    Dim target As Object
    Set target = workbook.Names.Item("rngRebarInput").RefersToRange

    Dim rowIndex As Long
    Dim colIndex As Long
    For rowIndex = 2 To target.Rows.Count
        For colIndex = 1 To target.Columns.Count
            target.Cells.Item(rowIndex, colIndex).ClearContents
        Next colIndex
    Next rowIndex

    Dim countToWrite As Long
    countToWrite = layout.Count
    If countToWrite > target.Rows.Count - 1 Then countToWrite = target.Rows.Count - 1

    Dim i As Long
    For i = 1 To countToWrite
        WriteRebarRow target, i + 1, layout, i
    Next i

    If layout.Count > countToWrite Then
        target.Cells.Item(target.Rows.Count, 7).Value2 = "Полный список на System"
    End If

    WriteGeneratedRebarSystemTable workbook, layout
End Sub

Private Sub WriteGeneratedRebarSystemTable(ByVal workbook As Object, ByVal layout As CRebarLayout)
    Dim sheet As Object
    Set sheet = workbook.Worksheets.Item("System")
    sheet.Range("A130:G250").ClearContents
    sheet.Cells.Item(130, 1).Value2 = "Автоматически рассчитанные координаты арматуры"

    Dim headers As Variant
    headers = Array("ID", "X", "Y", "Diameter", "Area", "SteelClass", "Comment")
    Dim colIndex As Long
    For colIndex = 0 To 6
        sheet.Cells.Item(132, colIndex + 1).Value2 = headers(colIndex)
    Next colIndex

    Dim i As Long
    For i = 1 To layout.Count
        WriteRebarRow sheet.Range("A132:G250"), i + 1, layout, i
    Next i
End Sub

Private Sub WriteRebarRow(ByVal target As Object, ByVal rowIndex As Long, ByVal layout As CRebarLayout, ByVal barIndex As Long)
    target.Cells.Item(rowIndex, 1).Value2 = layout.BarID(barIndex)
    target.Cells.Item(rowIndex, 2).Value2 = layout.X(barIndex)
    target.Cells.Item(rowIndex, 3).Value2 = layout.Y(barIndex)
    target.Cells.Item(rowIndex, 4).Value2 = layout.Diameter(barIndex)
    target.Cells.Item(rowIndex, 5).Value2 = layout.Area(barIndex)
    target.Cells.Item(rowIndex, 6).Value2 = layout.SteelClass(barIndex)
    target.Cells.Item(rowIndex, 7).Value2 = "auto"
End Sub

Private Sub WriteFirstCombinationResults(ByVal workbook As Object, ByVal mesh As CFiberMeshBuilder, _
        ByVal rebars As CRebarLayout, ByVal concrete As Object, ByVal steel As Object, ByVal settings As CSystemSettingsReader)
    Dim source As Object
    Set source = workbook.Names.Item("rngLoadCombinations").RefersToRange

    Dim rowIndex As Long
    For rowIndex = 2 To source.Rows.Count
        If Not IsEmptyRow(source, rowIndex, 7) Then
            If HasPartialLoadRowError(source, rowIndex) Then Exit Sub
            Dim nValue As Double
            Dim mxValue As Double
            Dim myValue As Double
            nValue = CDbl(source.Cells.Item(rowIndex, 2).Value2)
            mxValue = CDbl(source.Cells.Item(rowIndex, 3).Value2)
            myValue = CDbl(source.Cells.Item(rowIndex, 4).Value2)
            WriteCapacityAndCrackResults workbook, mesh, rebars, concrete, steel, settings, nValue, mxValue, myValue
            Exit Sub
        End If
    Next rowIndex
End Sub

Private Sub WriteCapacityAndCrackResults(ByVal workbook As Object, ByVal mesh As CFiberMeshBuilder, _
        ByVal rebars As CRebarLayout, ByVal concrete As Object, ByVal steel As Object, ByVal settings As CSystemSettingsReader, _
        ByVal nValue As Double, ByVal mxValue As Double, ByVal myValue As Double)

    Dim writer As CCapacityResultWriter
    Set writer = New CCapacityResultWriter

    Dim calculationMode As String
    calculationMode = settings.GetString("Calculation.Mode", "DirectState")

    If StrComp(calculationMode, "LinearMatrix", vbTextCompare) = 0 Then
        WriteLinearMatrixResults workbook, mesh, rebars, settings, nValue, mxValue, myValue, writer
        Exit Sub
    End If

    If StrComp(calculationMode, "FullCapacity", vbTextCompare) <> 0 Then
        WriteDirectStateAndCrackResults workbook, mesh, rebars, concrete, steel, settings, nValue, mxValue, myValue, writer
        Exit Sub
    End If

    Dim capacity As CCapacitySolver
    Set capacity = New CCapacitySolver
    capacity.ApplySettings settings
    If Sqr(mxValue * mxValue + myValue * myValue) > 0.000000001 Then
        capacity.SolveByLoadMultiplier mesh, rebars, concrete, steel, nValue, mxValue, myValue
    End If
    writer.WriteCapacityResult workbook, capacity
    Dim service As CSectionSolver
    Set service = New CSectionSolver
    service.ApplySettings settings
    service.Solve mesh, rebars, concrete, steel, nValue, mxValue, myValue

    If service.Converged Then
        Dim crack As CCrackWidthCalculator
        Set crack = New CCrackWidthCalculator
        crack.ApplySettings settings
        If settings.GetBoolean("CrackWidth.Enabled", True) Then
            crack.Calculate service, rebars, steel
            writer.WriteCrackResult workbook, crack
            
            
            WriteCrackFormulaBlock workbook, service, rebars, settings
            LinkCrackResultRowsToFormulaBlock workbook
        End If
    End If
End Sub

Private Sub WriteLinearMatrixResults(ByVal workbook As Object, ByVal mesh As CFiberMeshBuilder, _
        ByVal rebars As CRebarLayout, ByVal settings As CSystemSettingsReader, _
        ByVal nValue As Double, ByVal mxValue As Double, ByVal myValue As Double, ByVal writer As CCapacityResultWriter)

    Dim linear As CLinearSectionSolver
    Set linear = New CLinearSectionSolver
    linear.Solve mesh, rebars, settings.GetDouble("Concrete.Eb", 32500#), _
        settings.GetDouble("Steel.Es", 200000#), nValue, mxValue, myValue

    writer.WriteLinearSectionResult workbook, linear
    
    
End Sub

Private Sub WriteDirectStateAndCrackResults(ByVal workbook As Object, ByVal mesh As CFiberMeshBuilder, _
        ByVal rebars As CRebarLayout, ByVal concrete As Object, ByVal steel As Object, ByVal settings As CSystemSettingsReader, _
        ByVal nValue As Double, ByVal mxValue As Double, ByVal myValue As Double, ByVal writer As CCapacityResultWriter)

    Dim service As CSectionSolver
    Set service = New CSectionSolver
    service.ApplySettings settings
    service.Solve mesh, rebars, concrete, steel, nValue, mxValue, myValue

    writer.WriteDirectSectionResult workbook, service
    
    

    If service.Converged And settings.GetBoolean("CrackWidth.Enabled", True) Then
        Dim crack As CCrackWidthCalculator
        Set crack = New CCrackWidthCalculator
        crack.ApplySettings settings
        crack.Calculate service, rebars, steel
        writer.WriteCrackResult workbook, crack
        
        
        WriteCrackFormulaBlock workbook, service, rebars, settings
        LinkCrackResultRowsToFormulaBlock workbook
    End If
End Sub

Private Sub WriteCrackFormulaBlock(ByVal workbook As Object, ByVal solver As CSectionSolver, _
        ByVal rebars As CRebarLayout, ByVal settings As CSystemSettingsReader)
    Dim sheet As Object
    Set sheet = workbook.Worksheets.Item("System")
    sheet.Range("I130:Q250").ClearContents

    sheet.Cells.Item(130, 9).Value2 = "Расчет трещин формулами Excel"
    sheet.Cells.Item(131, 9).Value2 = "НДС берется из CSectionSolver; геометрическая часть и ширина трещины считаются формулами листа."

    PutCrackFormulaHeader sheet, 133
    PutCrackValueRow sheet, 134, "Epsilon0", solver.Epsilon0, "", "Из найденного SLS-состояния CSectionSolver"
    PutCrackValueRow sheet, 135, "KappaX", solver.KappaX, "1/mm", "Кривизна по координате Y"
    PutCrackValueRow sheet, 136, "KappaY", solver.KappaY, "1/mm", "Кривизна по координате X"
    PutCrackFormulaRow sheet, 137, "Circle.Diameter", "=VLOOKUP(""Circle.Diameter"",$A:$B,2,FALSE)", "mm", "Диаметр круглого сечения"
    PutCrackFormulaRow sheet, 138, "Circle.CenterX", "=VLOOKUP(""Circle.CenterX"",$A:$B,2,FALSE)", "mm", "Координата центра X"
    PutCrackFormulaRow sheet, 139, "Circle.CenterY", "=VLOOKUP(""Circle.CenterY"",$A:$B,2,FALSE)", "mm", "Координата центра Y"
    PutCrackFormulaRow sheet, 140, "Rebar.AxisDistance", "='Расчет'!$E$19", "mm", "Расстояние от грани до оси стержней"
    PutCrackFormulaRow sheet, 141, "Rebar.Count", "='Расчет'!$E$20", "шт", "Количество стержней"
    PutCrackFormulaRow sheet, 142, "Rebar.Diameter", "='Расчет'!$E$21", "mm", "Диаметр стержней"
    PutCrackFormulaRow sheet, 143, "Rebar.Radius", "=$J$137/2-$J$140", "mm", "Радиус окружности осей стержней"
    PutCrackFormulaRow sheet, 144, "Rebar.StepAlong", "=2*PI()*$J$143/$J$141", "mm", "Длина дуги между соседними стержнями"
    PutCrackFormulaRow sheet, 145, "CrackWidth.Allowable", "='Расчет'!$E$22", "mm", "Допустимая ширина раскрытия из пользовательского ввода"

    Dim firstBarRow As Long
    Dim lastBarRow As Long
    firstBarRow = 160
    lastBarRow = firstBarRow + rebars.Count - 1

    PutCrackFormulaRow sheet, 146, "MaxSteelStrain", "=MAX(0,MAX($O$" & CStr(firstBarRow) & ":$O$" & CStr(lastBarRow) & "))", "", "Максимальная растягивающая деформация стержня"
    PutCrackFormulaRow sheet, 147, "MaxSteelStress", "=MAX(0,MAX($P$" & CStr(firstBarRow) & ":$P$" & CStr(lastBarRow) & "))", "MPa", "Напряжение наиболее растянутого стержня"
    PutCrackFormulaRow sheet, 148, "TensionRebarCount", "=COUNTIF($Q$" & CStr(firstBarRow) & ":$Q$" & CStr(lastBarRow) & ",TRUE)", "шт", "Количество стержней с eps_s > 0"
    PutCrackFormulaRow sheet, 149, "As.Tension", "=SUMIF($Q$" & CStr(firstBarRow) & ":$Q$" & CStr(lastBarRow) & ",TRUE,$N$" & CStr(firstBarRow) & ":$N$" & CStr(lastBarRow) & ")", "mm2", "Суммарная площадь растянутой арматуры"
    PutCrackFormulaRow sheet, 150, "Ar.Interaction", "=$J$149", "mm2", "Временная прозрачная формула: сейчас равна As.Tension; здесь заменить расчет площади взаимодействия по СП/пособию для выбранной геометрии"
    PutCrackFormulaRow sheet, 151, "Ar.Interaction.cm2", "=$J$150/100", "cm2", "То же значение в см2 для сравнения с контрольными программами"
    PutCrackValueRow sheet, 152, "CrackWidth.CrackSpacing", settings.GetDouble("CrackWidth.CrackSpacing", 200#), "mm", "Временный параметр; формулу расстояния между трещинами нужно раскрыть после трассировки СП"
    PutCrackValueRow sheet, 153, "CrackWidth.StrainFactor", settings.GetDouble("CrackWidth.StrainFactor", 1#), "", "Временный коэффициент"
    PutCrackValueRow sheet, 154, "CrackWidth.DurationFactor", settings.GetDouble("CrackWidth.DurationFactor", 1#), "", "Временный коэффициент длительности"
    PutCrackFormulaRow sheet, 155, "CrackWidth", "=$J$146*$J$152*$J$153*$J$154", "mm", "Предварительная формула; NDM дает eps_s, нормативная часть раскрыта на листе"
    PutCrackFormulaRow sheet, 156, "CrackUtilization", "=IF($J$145>0,$J$155/$J$145,""InvalidInput"")", "", "Коэффициент использования по ширине раскрытия"

    WriteCrackRebarFormulaTable sheet, firstBarRow - 2, rebars.Count, _
        settings.GetDouble("Steel.Es", 200000#), _
        settings.GetDouble("Steel.Rs.ULS", 350#), _
        settings.GetDouble("Steel.Rsc.ULS", 350#)
End Sub

Private Sub WriteCrackRebarFormulaTable(ByVal sheet As Object, ByVal headerRow As Long, ByVal barCount As Long, _
        ByVal steelEs As Double, ByVal steelRs As Double, ByVal steelRsc As Double)
    Dim headers As Variant
    headers = Array("ID", "Alpha", "X", "Y", "Diameter", "Area", "EpsS", "SigmaS", "Tension")

    Dim colIndex As Long
    For colIndex = 0 To 8
        sheet.Cells.Item(headerRow, 9 + colIndex).Value2 = headers(colIndex)
    Next colIndex

    Dim rowIndex As Long
    Dim i As Long
    For i = 1 To barCount
        rowIndex = headerRow + i
        sheet.Cells.Item(rowIndex, 9).Formula = "=""R""&ROW()-" & CStr(headerRow)
        sheet.Cells.Item(rowIndex, 10).Formula = "=2*PI()*(ROW()-" & CStr(headerRow + 1) & ")/$J$141"
        sheet.Cells.Item(rowIndex, 11).Formula = "=$J$138+$J$143*COS(J" & CStr(rowIndex) & ")"
        sheet.Cells.Item(rowIndex, 12).Formula = "=$J$139+$J$143*SIN(J" & CStr(rowIndex) & ")"
        sheet.Cells.Item(rowIndex, 13).Formula = "=$J$142"
        sheet.Cells.Item(rowIndex, 14).Formula = "=PI()*M" & CStr(rowIndex) & "^2/4"
        sheet.Cells.Item(rowIndex, 15).Formula = "=$J$134+$J$135*L" & CStr(rowIndex) & "+$J$136*K" & CStr(rowIndex)
        sheet.Cells.Item(rowIndex, 16).Formula = "=MAX(MIN(O" & CStr(rowIndex) & "*" & FormatFormulaNumber(steelEs) & "," & FormatFormulaNumber(steelRs) & "),-" & FormatFormulaNumber(steelRsc) & ")"
        sheet.Cells.Item(rowIndex, 17).Formula = "=O" & CStr(rowIndex) & ">0"
    Next i
End Sub

Private Sub LinkCrackResultRowsToFormulaBlock(ByVal workbook As Object)
    LinkCrackResultRange workbook.Names.Item("rngResultSection").RefersToRange
End Sub

Private Sub LinkCrackResultRange(ByVal target As Object)
    target.Cells.Item(17, 5).Formula = "='System'!$J$155"
    target.Cells.Item(18, 5).Formula = "='System'!$J$145"
    target.Cells.Item(18, 13).Formula = "=""Utilization=""&'System'!$J$156"
    target.Cells.Item(19, 5).Formula = "='System'!$J$146"
    target.Cells.Item(19, 13).Formula = "=""StressMPa=""&'System'!$J$147&""; tensionBars=""&'System'!$J$148&""; Ar_cm2=""&'System'!$J$151"
End Sub

Private Sub PutCrackFormulaHeader(ByVal sheet As Object, ByVal rowIndex As Long)
    sheet.Cells.Item(rowIndex, 9).Value2 = "Параметр"
    sheet.Cells.Item(rowIndex, 10).Value2 = "Значение / формула"
    sheet.Cells.Item(rowIndex, 11).Value2 = "Ед."
    sheet.Cells.Item(rowIndex, 12).Value2 = "Комментарий"
End Sub

Private Sub PutCrackValueRow(ByVal sheet As Object, ByVal rowIndex As Long, ByVal label As String, _
        ByVal value As Variant, ByVal unitName As String, ByVal comment As String)
    sheet.Cells.Item(rowIndex, 9).Value2 = label
    sheet.Cells.Item(rowIndex, 10).Value2 = value
    sheet.Cells.Item(rowIndex, 11).Value2 = unitName
    sheet.Cells.Item(rowIndex, 12).Value2 = comment
End Sub

Private Sub PutCrackFormulaRow(ByVal sheet As Object, ByVal rowIndex As Long, ByVal label As String, _
        ByVal formulaText As String, ByVal unitName As String, ByVal comment As String)
    sheet.Cells.Item(rowIndex, 9).Value2 = label
    sheet.Cells.Item(rowIndex, 10).Formula = formulaText
    sheet.Cells.Item(rowIndex, 11).Value2 = unitName
    sheet.Cells.Item(rowIndex, 12).Value2 = comment
End Sub

Private Function FormatFormulaNumber(ByVal value As Double) As String
    FormatFormulaNumber = Replace$(CStr(value), ",", ".")
End Function

Private Function IsEmptyRow(ByVal source As Object, ByVal rowIndex As Long, ByVal columnCount As Long) As Boolean
    Dim col As Long
    For col = 1 To columnCount
        If Len(Trim$(CStr(source.Cells.Item(rowIndex, col).Value2))) > 0 Then Exit Function
    Next col
    IsEmptyRow = True
End Function

Private Function HasPartialLoadRowError(ByVal source As Object, ByVal rowIndex As Long) As Boolean
    HasPartialLoadRowError = Len(Trim$(CStr(source.Cells.Item(rowIndex, 1).Value2))) = 0 Or _
        Not IsNumeric(source.Cells.Item(rowIndex, 2).Value2) Or _
        Not IsNumeric(source.Cells.Item(rowIndex, 3).Value2) Or _
        Not IsNumeric(source.Cells.Item(rowIndex, 4).Value2)
End Function

Private Function ReadRequiredDouble(ByVal value As Variant, ByVal fieldName As String) As Double
    If Len(Trim$(CStr(value))) = 0 Then Err.Raise vbObjectError + 4120, "ReadRequiredDouble", fieldName & " is empty."
    If Not IsNumeric(value) Then Err.Raise vbObjectError + 4121, "ReadRequiredDouble", fieldName & " is not numeric."
    ReadRequiredDouble = CDbl(value)
End Function

Private Function ReadOptionalDouble(ByVal value As Variant, ByVal defaultValue As Double) As Double
    If Len(Trim$(CStr(value))) = 0 Then
        ReadOptionalDouble = defaultValue
    ElseIf IsNumeric(value) Then
        ReadOptionalDouble = CDbl(value)
    Else
        Err.Raise vbObjectError + 4122, "ReadOptionalDouble", "Value is not numeric."
    End If
End Function











