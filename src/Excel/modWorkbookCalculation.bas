Attribute VB_Name = "modWorkbookCalculation"
Option Explicit

Public Sub RunSectionCalculation()
    On Error GoTo Failed
    Dim message As String
    message = RunSectionCalculationForWorkbook(ThisWorkbook, True)
    If InStr(1, message, "ошиб", vbTextCompare) > 0 Or InStr(1, message, "InvalidInput", vbTextCompare) > 0 Then
        MsgBox message, vbExclamation, "RC Section NDM"
    Else
        MsgBox message, vbInformation, "RC Section NDM"
    End If
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

Public Sub UpdateSectionPlot()
    On Error GoTo Failed
    UpdateSectionPlotForWorkbook ThisWorkbook
    MsgBox "Схема сечения обновлена по последнему расчетному снимку Results.", vbInformation, "RC Section NDM"
    Exit Sub

Failed:
    MsgBox "Схема не обновлена: " & Err.Description, vbExclamation, "RC Section NDM"
End Sub

Public Sub UpdateSectionPlotForWorkbook(ByVal workbook As Object)
    If workbook Is Nothing Then Err.Raise vbObjectError + 4140, "UpdateSectionPlotForWorkbook", "Книга Excel не передана."

    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook workbook

    If Not settings.GetBoolean("Plot.Enabled", True) Then Exit Sub

    Dim reader As CSectionPlotDataReader
    Set reader = New CSectionPlotDataReader
    reader.LoadFromWorkbook workbook, settings

    Dim plotter As CSectionPlotter
    Set plotter = New CSectionPlotter
    plotter.Draw workbook, reader, settings
End Sub

Public Function RunSectionCalculationForWorkbook(ByVal workbook As Object, Optional ByVal showMessages As Boolean = False) As String
    If workbook Is Nothing Then Err.Raise vbObjectError + 4100, "RunSectionCalculationForWorkbook", "Книга Excel не передана."

    ClearSectionResultsForWorkbook workbook

    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook workbook

    Dim units As CUnitSystem
    Set units = New CUnitSystem
    units.LoadFromSettings settings

    Dim section As CSectionModel
    Set section = BuildWorkbookSectionModel(workbook, settings, units)

    Dim concrete As CConcreteDiagramMaterial
    Set concrete = New CConcreteDiagramMaterial
    concrete.InitializeFromSettings settings, units

    Dim steel As CSteelDiagramMaterial
    Set steel = New CSteelDiagramMaterial
    steel.InitializeFromSettings settings, units


    Dim batch As CBatchSectionCalculator
    Set batch = New CBatchSectionCalculator
    batch.Initialize section, concrete, steel
    batch.ApplySettings settings, units

    Dim reader As CLoadCombinationReader
    Set reader = New CLoadCombinationReader
    reader.LoadFromWorkbook workbook, batch, units
    If batch.Count = 0 Then Err.Raise vbObjectError + 4101, "RunSectionCalculationForWorkbook", "Не задано ни одного сочетания нагрузок."

    ApplyLoadReferenceFromSettings section, concrete, steel, settings, units, batch

    batch.Execute

    Dim summaryWriter As CBatchResultWriter
    Set summaryWriter = New CBatchResultWriter
    summaryWriter.WriteSummary workbook, batch, units

    Dim ndmWriter As CNDMResultsWriter
    Set ndmWriter = New CNDMResultsWriter
    ndmWriter.WriteResults workbook, section, concrete, steel, batch, settings.GetString("Concrete.Class", "B30"), units

    WriteGoverningCombinationResults workbook, section, concrete, steel, settings, units, batch

    If settings.GetBoolean("Plot.AutoUpdateAfterCalculation", True) And batch.GoverningCombinationIndex > 0 Then
        UpdateSectionPlotForWorkbook workbook
    End If

    RunSectionCalculationForWorkbook = BuildCalculationMessage(section, settings, batch)
End Function

Private Function BuildCalculationMessage(ByVal section As CSectionModel, ByVal settings As CSystemSettingsReader, _
        ByVal batch As CBatchSectionCalculator) As String
    Dim prefix As String
    prefix = ImportedSectionMessage(section, settings)

    If batch.InvalidInputCount > 0 Then
        BuildCalculationMessage = prefix & "Расчет завершен с ошибками ввода. Обработано сочетаний: " & _
            CStr(batch.Count) & "; ошибок ввода: " & CStr(batch.InvalidInputCount) & "." & vbCrLf & _
            "Проверьте строки со статусом InvalidInput на листе Results." & vbCrLf & _
            "Первая ошибка: " & batch.FirstInvalidInputMessage
    Else
        BuildCalculationMessage = prefix & "Расчет завершен. Обработано сочетаний: " & CStr(batch.Count) & _
            ". Определяющее сочетание: " & batch.GoverningCombinationID
    End If
End Function

Private Function ImportedSectionMessage(ByVal section As CSectionModel, ByVal settings As CSystemSettingsReader) As String
    If section Is Nothing Then Exit Function
    If settings Is Nothing Then Exit Function

    If StrComp(settings.GetRawString("Geometry.Source", "Generated"), "AutoCAD", vbTextCompare) = 0 Then
        ImportedSectionMessage = "Импортировано из AutoCAD: бетонных Region: " & _
            CStr(section.ConcreteCount) & "; арматурных Region: " & CStr(section.RebarCount) & "." & vbCrLf
    End If
End Function

Private Sub ApplyLoadReferenceFromSettings(ByVal section As CSectionModel, _
        ByVal concrete As Object, ByVal steel As Object, ByVal settings As CSystemSettingsReader, _
        ByVal units As CUnitSystem, _
        ByVal batch As CBatchSectionCalculator)
    Dim referenceX As Double
    Dim referenceY As Double
    CalculateTransformedSectionCentroid section, concrete, steel, referenceX, referenceY

    batch.ApplyLoadReference referenceX + units.InputLengthToInternal(settings.GetDouble("Load.ReferenceOffsetX", 0#)), _
        referenceY + units.InputLengthToInternal(settings.GetDouble("Load.ReferenceOffsetY", 0#))
End Sub

Public Sub CalculateTransformedSectionCentroid(ByVal section As CSectionModel, _
        ByVal concrete As Object, ByVal steel As Object, ByRef referenceX As Double, ByRef referenceY As Double)
    Dim props As CSectionPropertiesCalculator
    Set props = New CSectionPropertiesCalculator
    props.CalculateTransformed section, concrete, steel
    referenceX = props.CentroidX
    referenceY = props.CentroidY
End Sub

Public Sub ClearSectionResultsForWorkbook(ByVal workbook As Object)
    If workbook Is Nothing Then Err.Raise vbObjectError + 4110, "ClearSectionResultsForWorkbook", "Книга Excel не передана."
    ClearResultRange workbook.Names.Item("rngResultSection").RefersToRange
    Dim summaryWriter As CBatchResultWriter
    Set summaryWriter = New CBatchResultWriter
    summaryWriter.ClearSummary workbook
    Dim ndmWriter As CNDMResultsWriter
    Set ndmWriter = New CNDMResultsWriter
    ndmWriter.ClearResults workbook
End Sub

Private Sub ClearResultRange(ByVal target As Object)
    target.Offset(0, 0).Resize(target.Rows.Count, 1).ClearContents
    target.Offset(0, 4).Resize(target.Rows.Count, 1).ClearContents
    target.Offset(0, 8).Resize(target.Rows.Count, 1).ClearContents
    target.Offset(0, 12).Resize(target.Rows.Count, 1).ClearContents
End Sub

Public Function ReadWorkbookGeometry(ByVal workbook As Object, ByVal settings As CSystemSettingsReader, Optional ByVal units As CUnitSystem = Nothing) As ISectionGeometry
    If units Is Nothing Then
        Set units = New CUnitSystem
        units.InitializeDefaults
    End If
    Dim geometryType As String
    geometryType = settings.GetString("Geometry.Type", "LShape")
    If Len(geometryType) = 0 Then geometryType = "LShape"

    If StrComp(geometryType, "Circle", vbTextCompare) = 0 Then
        Dim circleGeom As CGeometryCircle
        Set circleGeom = New CGeometryCircle
        circleGeom.InitializeByDiameter units.InputLengthToInternal(settings.GetDouble("Circle.Diameter", 300#)), 0#, 0#
        Set ReadWorkbookGeometry = circleGeom
    ElseIf StrComp(geometryType, "LShape", vbTextCompare) = 0 Then
        Dim lshape As CGeometryLShape
        Set lshape = New CGeometryLShape
        lshape.Initialize units.InputLengthToInternal(settings.GetDouble("LShape.B1", 250#)), _
            units.InputLengthToInternal(settings.GetDouble("LShape.H1", 550#)), _
            units.InputLengthToInternal(settings.GetDouble("LShape.B2", 600#)), _
            units.InputLengthToInternal(settings.GetDouble("LShape.H2", 250#)), 0#, 0#
        Set ReadWorkbookGeometry = lshape
    ElseIf StrComp(geometryType, "RoundedRectangle", vbTextCompare) = 0 Then
        Dim rect As CGeometryRoundedRectangle
        Set rect = New CGeometryRoundedRectangle
        rect.Initialize units.InputLengthToInternal(settings.GetDouble("RoundedRectangle.Width", 300#)), _
            units.InputLengthToInternal(settings.GetDouble("RoundedRectangle.Height", 200#)), _
            units.InputLengthToInternal(settings.GetDouble("RoundedRectangle.RadiusTopLeft", 0#)), _
            units.InputLengthToInternal(settings.GetDouble("RoundedRectangle.RadiusTopRight", 0#)), _
            units.InputLengthToInternal(settings.GetDouble("RoundedRectangle.RadiusBottomRight", 0#)), _
            units.InputLengthToInternal(settings.GetDouble("RoundedRectangle.RadiusBottomLeft", 0#))
        Set ReadWorkbookGeometry = rect
    Else
        Err.Raise vbObjectError + 4102, "ReadWorkbookGeometry", "Неподдерживаемый тип сечения: " & geometryType
    End If
End Function

Public Function BuildWorkbookSectionModel(ByVal workbook As Object, ByVal settings As CSystemSettingsReader, Optional ByVal units As CUnitSystem = Nothing) As CSectionModel
    If units Is Nothing Then
        Set units = New CUnitSystem
        units.InitializeDefaults
    End If
    If workbook Is Nothing Then Err.Raise vbObjectError + 4130, "BuildWorkbookSectionModel", "Книга Excel не передана."
    If settings Is Nothing Then Err.Raise vbObjectError + 4131, "BuildWorkbookSectionModel", "Настройки System не переданы."

    Dim geometrySource As String
    geometrySource = Trim$(settings.GetRawString("Geometry.Source", "Generated"))
    If Len(geometrySource) = 0 Then Err.Raise vbObjectError + 4132, "BuildWorkbookSectionModel", _
        "Geometry.Source должен быть Generated или AutoCAD."

    If StrComp(geometrySource, "Generated", vbTextCompare) = 0 Then
        Dim geometry As ISectionGeometry
        Set geometry = ReadWorkbookGeometry(workbook, settings, units)

        Dim mesh As CFiberMeshBuilder
        Set mesh = New CFiberMeshBuilder
        Dim meshStep As Double
        meshStep = units.InputLengthToInternal(settings.GetDouble("Mesh.Step", 10#))
        mesh.BuildMesh geometry, meshStep, meshStep, 1, WorkbookMeshBoundarySubdivisions(settings)

        Dim rebars As CRebarLayout
        Set rebars = ReadWorkbookRebars(workbook, geometry, settings, units)

        Set BuildWorkbookSectionModel = BuildGeneratedSectionModel(mesh, rebars, settings.GetString("Geometry.Type", "Generated"))
    ElseIf StrComp(geometrySource, "AutoCAD", vbTextCompare) = 0 Then
        Dim importer As CAutoCADSectionModelImporter
        Set importer = New CAutoCADSectionModelImporter
        Set BuildWorkbookSectionModel = importer.ImportFromActiveDocument(settings, units)
    Else
        Err.Raise vbObjectError + 4133, "BuildWorkbookSectionModel", _
            "Geometry.Source должен быть Generated или AutoCAD."
    End If
End Function

Public Function WorkbookMeshBoundarySubdivisions(ByVal settings As CSystemSettingsReader) As Long
    WorkbookMeshBoundarySubdivisions = settings.GetLong("Mesh.BoundarySubdivisions", 1)
    If WorkbookMeshBoundarySubdivisions < 1 Then WorkbookMeshBoundarySubdivisions = 1
End Function

Public Function ReadWorkbookRebars(ByVal workbook As Object, ByVal geometry As ISectionGeometry, ByVal settings As CSystemSettingsReader, Optional ByVal units As CUnitSystem = Nothing) As CRebarLayout
    If units Is Nothing Then
        Set units = New CUnitSystem
        units.InitializeDefaults
    End If
    Dim geometryType As String
    geometryType = settings.GetString("Geometry.Type", "LShape")

    Dim layout As CRebarLayout
    If StrComp(geometryType, "Circle", vbTextCompare) = 0 Then
        Dim circleBuilder As CCircleRebarLayoutBuilder
        Set circleBuilder = New CCircleRebarLayoutBuilder
        Set layout = circleBuilder.Build( _
            units.InputLengthToInternal(settings.GetDouble("Circle.Diameter", 300#)), _
            0#, _
            0#, _
            units.InputLengthToInternal(settings.GetDouble("Rebar.AxisDistance", 40#)), _
            settings.GetLong("Rebar.Count", 8), _
            units.InputLengthToInternal(settings.GetDouble("Rebar.Diameter", 20#)), _
            settings.GetString("Steel.Class", "A400"), _
            units.InputLengthToInternal(settings.GetDouble("Rebar.Diameter2", 0#)), _
            units.InputLengthToInternal(settings.GetDouble("Rebar.Diameter3", 0#)), _
            settings.GetString("Rebar.Loc2row", "Stacked"), _
            settings.GetString("Rebar.Loc3row", "Stacked"))
    ElseIf StrComp(geometryType, "LShape", vbTextCompare) = 0 Then
        Dim lshapeBuilder As CLShapeRebarLayoutBuilder
        Set lshapeBuilder = New CLShapeRebarLayoutBuilder
        Set layout = lshapeBuilder.Build( _
            units.InputLengthToInternal(settings.GetDouble("LShape.B1", 250#)), _
            units.InputLengthToInternal(settings.GetDouble("LShape.H1", 550#)), _
            units.InputLengthToInternal(settings.GetDouble("LShape.B2", 600#)), _
            units.InputLengthToInternal(settings.GetDouble("LShape.H2", 250#)), _
            0#, _
            0#, _
            LShapeFaceSettings(settings, units, "H1"), _
            LShapeFaceSettings(settings, units, "H2"), _
            LShapeFaceSettings(settings, units, "B1"), _
            LShapeFaceSettings(settings, units, "B2"), _
            settings.GetString("Steel.Class", "A400"))
    Else
        Err.Raise vbObjectError + 4103, "ReadWorkbookRebars", "Автоматическая расстановка арматуры поддерживается только для Circle и LShape."
    End If

    Set ReadWorkbookRebars = layout
End Function

Private Function LShapeFaceSettings(ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem, ByVal faceName As String) As Variant
    LShapeFaceSettings = Array( _
        units.InputLengthToInternal(settings.GetDouble("LShape." & faceName & ".as_1", 40#)), _
        units.InputLengthToInternal(settings.GetDouble("LShape." & faceName & ".as_2", 40#)), _
        units.InputLengthToInternal(settings.GetDouble("LShape." & faceName & ".d_1", 32#)), _
        units.InputLengthToInternal(settings.GetDouble("LShape." & faceName & ".d_2", 32#)), _
        settings.GetLong("LShape." & faceName & ".n_1", 0), _
        settings.GetLong("LShape." & faceName & ".n_2", 0), _
        units.InputLengthToInternal(settings.GetDouble("LShape." & faceName & ".StartOffset1", 80#)), _
        units.InputLengthToInternal(settings.GetDouble("LShape." & faceName & ".EndOffset1", 80#)), _
        units.InputLengthToInternal(settings.GetDouble("LShape." & faceName & ".StartOffset2", 80#)), _
        units.InputLengthToInternal(settings.GetDouble("LShape." & faceName & ".EndOffset2", 80#)), _
        units.InputLengthToInternal(settings.GetDouble("LShape." & faceName & ".d_2row_1", 0#)), _
        units.InputLengthToInternal(settings.GetDouble("LShape." & faceName & ".d_2row_2", 0#)), _
        units.InputLengthToInternal(settings.GetDouble("LShape." & faceName & ".d_3row_1", 0#)), _
        units.InputLengthToInternal(settings.GetDouble("LShape." & faceName & ".d_3row_2", 0#)), _
        settings.GetString("LShape." & faceName & ".loc_2row", "Stacked"), _
        settings.GetString("LShape." & faceName & ".loc_3row", "Stacked"), _
        settings.GetString("LShape." & faceName & ".bind_2row", "EachBar"), _
        settings.GetString("LShape." & faceName & ".bind_3row", "EachBar"))
End Function

Private Sub WriteGoverningCombinationResults(ByVal workbook As Object, ByVal section As CSectionModel, _
        ByVal concrete As Object, ByVal steel As Object, _
        ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem, ByVal batch As CBatchSectionCalculator)
    If batch Is Nothing Then Err.Raise vbObjectError + 4123, "WriteGoverningCombinationResults", "Результаты пакетного расчета отсутствуют."
    If batch.GoverningCombinationIndex <= 0 Then Exit Sub

    Dim index As Long
    index = batch.GoverningCombinationIndex
    WriteCapacityAndCrackResults workbook, section, concrete, steel, settings, _
        units, batch.N(index), batch.UserMx(index), batch.UserMy(index), batch.LoadReferenceX, batch.LoadReferenceY
End Sub

Private Sub WriteCapacityAndCrackResults(ByVal workbook As Object, ByVal section As CSectionModel, _
        ByVal concrete As Object, ByVal steel As Object, ByVal settings As CSystemSettingsReader, _
        ByVal units As CUnitSystem, _
        ByVal nValue As Double, ByVal userMxValue As Double, ByVal userMyValue As Double, _
        ByVal referenceX As Double, ByVal referenceY As Double)

    Dim writer As CCapacityResultWriter
    Set writer = New CCapacityResultWriter

    Dim internalMxValue As Double
    Dim internalMyValue As Double
    internalMxValue = userMxValue + nValue * referenceY
    internalMyValue = userMyValue + nValue * referenceX

    Dim calculationMode As String
    calculationMode = settings.GetString("Calculation.Mode", "FullCapacity")

    Select Case LCase$(Trim$(calculationMode))
        Case "directstate"
            WriteDirectStateAndCrackResults workbook, section, concrete, steel, settings, units, nValue, internalMxValue, internalMyValue, writer
            Exit Sub
        Case "fullcapacity"
        Case Else
            Err.Raise vbObjectError + 4124, "WriteCapacityAndCrackResults", _
                "Calculation.Mode должен быть DirectState или FullCapacity."
    End Select

    Dim capacity As CCapacitySolver
    Set capacity = New CCapacitySolver
    capacity.ApplySettings settings, units
    If Sqr(userMxValue * userMxValue + userMyValue * userMyValue) > 0.000000001 Then
        Select Case LCase$(Trim$(settings.GetRawString("Capacity.Method", vbNullString)))
            Case "ultimatestrain"
                capacity.SolveByUltimateStrain section, concrete, steel, nValue, userMxValue, userMyValue, _
                    nValue * referenceY, nValue * referenceX
            Case "loadmultiplier"
                capacity.SolveByLoadMultiplier section, concrete, steel, nValue, userMxValue, userMyValue, _
                    nValue * referenceY, nValue * referenceX
            Case Else
                Err.Raise vbObjectError + 4125, "WriteCapacityAndCrackResults", _
                    "Capacity.Method должен быть LoadMultiplier или UltimateStrain."
        End Select
    End If
    writer.WriteCapacityResult workbook, capacity, units
    Dim service As CSectionSolver
    Set service = New CSectionSolver
    service.ApplySettings settings, units
    service.Solve section, concrete, steel, nValue, internalMxValue, internalMyValue

    If service.Converged Then
        Dim crack As CCrackWidthCalculator
        Set crack = New CCrackWidthCalculator
        crack.ApplySettings settings, units
        If settings.GetBoolean("CrackWidth.Enabled", True) Then
            crack.Calculate service, section, steel
            writer.WriteCrackResult workbook, crack, units
        End If
    End If
End Sub

Private Sub WriteDirectStateAndCrackResults(ByVal workbook As Object, ByVal section As CSectionModel, _
        ByVal concrete As Object, ByVal steel As Object, ByVal settings As CSystemSettingsReader, _
        ByVal units As CUnitSystem, _
        ByVal nValue As Double, ByVal mxValue As Double, ByVal myValue As Double, ByVal writer As CCapacityResultWriter)

    Dim service As CSectionSolver
    Set service = New CSectionSolver
    service.ApplySettings settings, units
    service.Solve section, concrete, steel, nValue, mxValue, myValue

    writer.WriteDirectSectionResult workbook, service, units
    
    

    If service.Converged And settings.GetBoolean("CrackWidth.Enabled", True) Then
        Dim crack As CCrackWidthCalculator
        Set crack = New CCrackWidthCalculator
        crack.ApplySettings settings, units
        crack.Calculate service, section, steel
        writer.WriteCrackResult workbook, crack, units
    End If
End Sub
















