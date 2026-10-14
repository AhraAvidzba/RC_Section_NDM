' ДЛЯ ТЕСТОВ: добавляется только к modTestPlotConfig в отдельных копиях.
' Создаёт настоящие схемы Excel без расчёта НДС и сохраняет координаты объектов.
Public Function GoalPlotObjectProbe(ByVal directory As String) As String
    Dim fixture As Object, config As Object, sheet As Object, table As Object, source As Object, chart As Object
    Dim settings As CSystemSettingsReader, reader As CSectionPlotDataReader, plotter As CSectionPlotter
    Dim keys As Variant, values As Variant, geometry As Variant, contours As Variant, props(1 To 20, 1 To 4) As Variant
    Dim elements(1 To 3, 1 To 7) As Variant, annotation(1 To 3, 1 To 9) As Variant
    Dim angles As Variant, sweeps As Variant, modes As Variant, sizes As Variant, angle As Variant, sweep As Variant, axesMode As Variant
    Dim row As Long, column As Long, caseIndex As Long, sizeIndex As Long, repeat As Long, file As Integer, run As String
    Dim limit As Double, annotationIndex As Long
    Dim shape As Object, point As Variant, node As Long, frame As Variant, arcCount As Long, axisCount As Long
    Dim started As Double, elapsed As Double, solveCount As Long, failed As Boolean, errorText As String
    On Error GoTo Failed
    solveCount = SectionEquilibriumSolveCount()
    Set fixture = Application.Workbooks.Add(-4167)
    Set config = fixture.Worksheets(1): config.Name = "PlotConfig"
    Set sheet = fixture.Worksheets.Add: sheet.Name = "Results": fixture.Worksheets.Add.Name = "Расчет"
    Set table = Audit03NamedGeneralPlotSettingsTable(fixture, config)
    Audit03SetPlotSetting table, "Plot.LoadCase", "PLOT"
    Audit03SetPlotSetting table, "Plot.ContourEnabled", "Yes"
    Audit03SetPlotSetting table, "Plot.NeutralLineEnabled", "Yes"
    Audit03SetPlotSetting table, "Plot.LoadApplicationPointEnabled", "Yes"
    Audit03SetPlotSetting table, "Plot.Dimensions.Enabled", "Yes"
    Audit03SetPlotSetting table, "Plot.RebarLabels.Enabled", "Yes"
    Set source = ThisWorkbook.Names.Item("rngCalculationProfiles").RefersToRange
    config.Range("AF5").Resize(source.Rows.Count, source.Columns.Count).Value2 = source.Value2
    For row = 2 To source.Rows.Count
        If CStr(source.Cells(row, 2).Value2) = "Visualization.State" Then config.Range("AF5").Cells(row, 3).Value2 = "StrengthState"
        If CStr(source.Cells(row, 2).Value2) = "Visualization.Quantity" Then config.Range("AF5").Cells(row, 3).Value2 = "Stress"
    Next row
    fixture.Names.Add Name:="rngCalculationProfiles", RefersTo:="=PlotConfig!" & config.Range("AF5").Resize(source.Rows.Count, source.Columns.Count).Address
    geometry = Audit03ContourGeometrySnapshot("CONTOUR_ARC")
    sheet.Range("A5").Resize(3, 15).Value2 = geometry
    fixture.Names.Add Name:="rngNDMSectionGeometry", RefersTo:="=Results!$A$5"
    Audit03PrepareContourFixture fixture, "CONTOUR_ARC"
    keys = Array("Bounds.MinX", "Bounds.MaxX", "Bounds.MinY", "Bounds.MaxY", _
        "Concrete.CentroidX", "Concrete.CentroidY", "Concrete.PrincipalAngle", _
        "Transformed.CentroidX", "Transformed.CentroidY", "Transformed.PrincipalAngle", _
        "LoadReferenceX", "LoadReferenceY", "State.StrengthState.Epsilon0", _
        "State.StrengthState.KappaX", "State.StrengthState.KappaY", "ProfileId", _
        "State.StrengthState.ExtensionUsed", "State.StrengthState.Status", "Output.LengthUnit")
    values = Array(-150#, 150#, -150#, 150#, 20#, -10#, 0#, 28#, 15#, 0#, _
        25#, -15#, -0.0001, 0.000002, 0.000001, "PR1", "False", "OK", "mm")
    props(1, 1) = "LoadCase": props(1, 2) = "Parameter": props(1, 3) = "Value": props(1, 4) = "Unit"
    For row = 0 To UBound(keys)
        props(row + 2, 1) = "ALL": If row >= 12 And row <= 17 Then props(row + 2, 1) = "PLOT"
        props(row + 2, 2) = keys(row): props(row + 2, 3) = values(row): props(row + 2, 4) = "mm"
        If row = 6 Or row = 9 Or row = 12 Or row >= 15 Then props(row + 2, 4) = "-"
        If row = 13 Or row = 14 Then props(row + 2, 4) = "1/mm"
    Next row
    sheet.Range("AB5").Resize(20, 4).Value2 = props
    fixture.Names.Add Name:="rngNDMSectionProperties", RefersTo:="=Results!$AB$5"
    keys = Array("LoadCase", "ProfileId", "StateType", "ElementID", "Strain", "Stress, MPa", "PhysicalState")
    For column = 0 To UBound(keys): elements(1, column + 1) = keys(column): Next column
    elements(2, 1) = "PLOT": elements(2, 2) = "PR1": elements(2, 3) = "StrengthState": elements(2, 4) = "C1"
    elements(2, 5) = -0.0001: elements(2, 6) = -1.25: elements(2, 7) = "Compression"
    elements(3, 1) = "PLOT": elements(3, 2) = "PR1": elements(3, 3) = "StrengthState": elements(3, 4) = "R1"
    elements(3, 5) = 0.000015: elements(3, 6) = 2.75: elements(3, 7) = "Tension"
    sheet.Range("R35").Resize(3, 7).Value2 = elements
    fixture.Names.Add Name:="rngNDMElementResults", RefersTo:="=Results!$R$35"
    keys = Array("AnnotationType", "StartX", "StartY", "EndX", "EndY", "OutsideNormalX", "OutsideNormalY", "Text", "Unit")
    For column = 0 To UBound(keys): annotation(1, column + 1) = keys(column): Next column
    For row = 2 To 3
        For column = 2 To 7: annotation(row, column) = 0#: Next column
        annotation(row, 2) = -50#: annotation(row, 4) = 50#: annotation(row, 7) = 1#: annotation(row, 9) = "mm"
    Next row
    annotation(2, 1) = "DIMENSION": annotation(2, 8) = "DIM"
    annotation(3, 1) = "REBAR_ANNOTATION": annotation(3, 8) = "REBAR"
    sheet.Range("A30").Resize(3, 9).Value2 = annotation
    fixture.Names.Add Name:="rngNDMSectionAnnotations", RefersTo:="=Results!$A$30"
    Set settings = New CSystemSettingsReader: Set reader = New CSectionPlotDataReader: Set plotter = New CSectionPlotter
    Set chart = fixture.Worksheets.Item("Расчет").ChartObjects.Add(41#, 53#, 700#, 480#): chart.Name = "chtNDMSectionPlot"
    angles = Array(0#, 0.6, -0.7)
    sweeps = Array(GEOM_PI / 2#, -GEOM_PI / 2#, GEOM_PI, 1.75 * GEOM_PI, 0.001, "Circle")
    modes = Array("Concrete", "Transformed", "None")
    sizes = Array(Array(41#, 53#, 700#, 480#), Array(311#, 227#, 700#, 480#), Array(80#, 90#, 900#, 620#))
    file = FreeFile: Open directory & "\PlotObjects.txt" For Output As #file
    For Each axesMode In modes
        Audit03SetPlotSetting table, "Plot.PrincipalAxesMode", CStr(axesMode)
        For Each angle In angles
            sheet.Range("AB5").Cells(8, 3).Value2 = CDbl(angle)
            sheet.Range("AB5").Cells(11, 3).Value2 = CDbl(angle) + 0.15
            For Each sweep In sweeps
                If VarType(sweep) = vbString Then
                    contours = Audit03ContourSnapshot("CONTOUR_CIRCLE")
                    limit = 150#
                Else
                    contours = GoalClosedArcSnapshot(CDbl(sweep))
                    limit = 150#: If Abs(CDbl(sweep)) < 0.01 Then limit = 1500#
                End If
                sheet.Range("AB5").Cells(2, 3).Value2 = -limit: sheet.Range("AB5").Cells(3, 3).Value2 = limit
                sheet.Range("AB5").Cells(4, 3).Value2 = -limit: sheet.Range("AB5").Cells(5, 3).Value2 = limit
                sheet.Range("EZ5").Resize(3, 17).ClearContents
                sheet.Range("EZ5").Resize(UBound(contours, 1), 17).Value2 = contours
                For sizeIndex = 0 To 2
                    caseIndex = caseIndex + 1: run = CStr(caseIndex)
                    chart.Left = sizes(sizeIndex)(0): chart.Top = sizes(sizeIndex)(1)
                    chart.Width = sizes(sizeIndex)(2): chart.Height = sizes(sizeIndex)(3)
                    settings.LoadFromWorkbook fixture: reader.LoadFromWorkbook fixture, settings
                    plotter.Draw fixture, reader, settings
                    frame = plotter.GoalCapturedFrame()
                    Print #file, "CASE|" & run & "|" & CStr(axesMode) & "|" & CStr(angle) & "|" & CStr(sweep) & "|" & CStr(sizeIndex)
                    Print #file, "FRAME|" & run & "|" & Join(frame, "|")
                    Print #file, "DATA|" & run & "|" & CStr(reader.MinX) & "|" & CStr(reader.MaxX) & "|" & CStr(reader.MinY) & "|" & CStr(reader.MaxY) & "|" & CStr(reader.PrincipalAxesCentroidX(CStr(axesMode))) & "|" & CStr(reader.PrincipalAxesCentroidY(CStr(axesMode))) & "|" & CStr(reader.PrincipalAxesAngle(CStr(axesMode)))
                    For annotationIndex = 1 To reader.AnnotationCount
                        If reader.AnnotationType(annotationIndex) = "CONTOUR_ARC" Then Print #file, "ARC|" & run & "|" & CStr(reader.AnnotationStartX(annotationIndex)) & "|" & CStr(reader.AnnotationStartY(annotationIndex)) & "|" & CStr(reader.AnnotationEndX(annotationIndex)) & "|" & CStr(reader.AnnotationEndY(annotationIndex)) & "|" & CStr(reader.AnnotationNormalX(annotationIndex)) & "|" & CStr(reader.AnnotationNormalY(annotationIndex)) & "|" & CStr(reader.AnnotationSweepAngle(annotationIndex))
                    Next annotationIndex
                    axisCount = 0: arcCount = 0
                    For Each shape In chart.Chart.Shapes
                        Print #file, "SHAPE|" & run & "|" & shape.Name & "|" & CStr(shape.Type) & "|" & CStr(shape.Left) & "|" & CStr(shape.Top) & "|" & CStr(shape.Width) & "|" & CStr(shape.Height) & "|" & CStr(shape.Rotation)
                        If InStr(shape.Name, "Principal") > 0 Then axisCount = axisCount + 1
                        If InStr(shape.Name, "ContourArcLine") > 0 Then arcCount = arcCount + 1
                        If shape.Type = 5 Then
                            For node = 1 To shape.Nodes.Count
                                point = shape.Nodes.Item(node).Points
                                Print #file, "NODE|" & run & "|" & shape.Name & "|" & CStr(node) & "|" & CStr(point(1, 1)) & "|" & CStr(point(1, 2))
                            Next node
                        End If
                    Next shape
                    Print #file, "COUNTS|" & run & "|" & CStr(chart.Chart.Shapes.Count) & "|" & CStr(axisCount) & "|" & CStr(arcCount)
                    If caseIndex = 10 Or caseIndex = 16 Or caseIndex = 19 Or caseIndex = 20 Or caseIndex = 21 Or caseIndex = 25 Or caseIndex = 37 Then
                        chart.Chart.Export directory & "\Case" & run & ".png", "PNG"
                        fixture.SaveAs directory & "\Case" & run & ".xlsx", 51
                    End If
                    If sizeIndex = 0 And CDbl(angle) = 0# And VarType(sweep) <> vbString Then
                      If CDbl(sweep) = GEOM_PI / 2# Then
                        For repeat = 1 To 7
                            started = Timer: plotter.Draw fixture, reader, settings: elapsed = Timer - started
                            If elapsed < 0# Then elapsed = elapsed + 86400#
                            Print #file, "TIME|" & CStr(axesMode) & "|" & CStr(repeat) & "|" & CStr(elapsed)
                        Next repeat
                      End If
                    End If
                Next sizeIndex
            Next sweep
        Next angle
    Next axesMode
    Close #file: file = 0
    If SectionEquilibriumSolveCount() <> solveCount Then Err.Raise vbObjectError + 4701, "GoalPlotObjectProbe", "Плоттер запустил расчёт НДС."
    GoalPlotObjectProbe = "TOTAL_GOAL_PLOT_OBJECTS: passed=" & CStr(caseIndex) & "; failed=0"
    GoTo Cleanup
Failed:
    failed = True: errorText = "RUNTIME ERROR: " & CStr(Err.Number) & "|" & Err.Description
Cleanup:
    On Error Resume Next
    If file <> 0 Then Close #file
    If Not fixture Is Nothing Then fixture.Close False
    On Error GoTo 0
    If failed Then GoalPlotObjectProbe = errorText
End Function

' ДЛЯ ТЕСТОВ: дуга и её настоящая замыкающая граница образуют бетонную область.
' При полном обороте дополнительная нулевая линия не создаётся.
Private Function GoalClosedArcSnapshot(ByVal sweep As Double) As Variant
    Dim arc As Variant, data() As Variant, row As Long, column As Long, rows As Long, radius As Double
    arc = Audit03ContourSnapshot("CONTOUR_ARC")
    radius = 100#: If Abs(sweep) < 0.01 Then radius = 1000#
    arc(2, 7) = radius: arc(2, 13) = radius
    arc(2, 9) = radius * Cos(sweep): arc(2, 10) = radius * Sin(sweep): arc(2, 14) = sweep
    rows = 3: If Abs(Abs(sweep) - 2# * GEOM_PI) < 0.0000000001 Then rows = 2
    ReDim data(1 To rows, 1 To 17)
    For row = 1 To 2
        For column = 1 To 17: data(row, column) = arc(row, column): Next column
    Next row
    If rows = 3 Then
        For column = 1 To 17: data(3, column) = arc(2, column): Next column
        data(3, 3) = "R2": data(3, 4) = 2: data(3, 6) = "CONTOUR_LINE"
        data(3, 7) = arc(2, 9): data(3, 8) = arc(2, 10): data(3, 9) = arc(2, 7): data(3, 10) = arc(2, 8)
        For column = 11 To 14: data(3, column) = 0#: Next column
    End If
    GoalClosedArcSnapshot = data
End Function
