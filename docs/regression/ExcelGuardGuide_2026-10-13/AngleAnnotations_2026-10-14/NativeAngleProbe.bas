' ДЛЯ ТЕСТОВ: реальный renderer в отдельном Chart; методы не публикуются в книгу.
Public Function GoalNativeAngleProbe(ByVal directory As String) As String
    Dim book As Object, sheet As Object, config As Object, table As Object, settings As CSystemSettingsReader
    Dim layout As CPlotAnnotationLayout, plotter As CSectionPlotter, chart As Object, shape As Object
    Dim degrees As Variant, degree As Variant, side As Variant, aspect As Variant, sign As Variant, kind As Variant
    Dim angle As Double, dx As Double, dy As Double, nx As Double, ny As Double, i As Long, count As Long, file As Integer
    Dim text As String, fileBase As String, field As Variant, values As Variant, snapshot As String, previous As String
    On Error GoTo Failed
    Set book = Application.Workbooks.Add(-4167): Set sheet = book.Worksheets(1): sheet.Name = "Angles"
    Set config = book.Worksheets.Add: config.Name = "LayoutSettings"
    Set table = Audit03PlotSettingsTable(config.Range("A5"))
    For Each kind In Array("Dimensions", "RebarLabels")
        Audit03SetPlotSetting table, "Plot." & CStr(kind) & ".Enabled", "Yes"
        Audit03SetPlotSetting table, "Plot." & CStr(kind) & ".TextUnits", "pt"
        Audit03SetPlotSetting table, "Plot." & CStr(kind) & ".TextHeight", 13#
        Audit03SetPlotSetting table, "Plot." & CStr(kind) & ".TextGap", 9#
        Audit03SetPlotSetting table, "Plot." & CStr(kind) & ".Offset", 30#
    Next kind
    Audit03SetPlotSetting table, "Plot.RebarLabels.LineEnabled", "Yes"
    Set settings = New CSystemSettingsReader: Set plotter = New CSectionPlotter
    Set chart = sheet.ChartObjects.Add(20#, 20#, 900#, 780#)
    chart.Chart.ChartType = -4169: chart.Chart.HasLegend = False: chart.Chart.HasTitle = False
    On Error Resume Next
    chart.Chart.Axes(1).Delete: chart.Chart.Axes(2).Delete
    On Error GoTo Failed
    file = FreeFile: Open directory & Application.PathSeparator & "NativeAngles.tsv" For Output As #file
    degrees = Array(0#, 30#, 45#, 90#, 135#, 225#, 315#)
    For Each aspect In Array(430#, 640#)
        For Each side In Array("Outside", "Inside")
            Audit03SetPlotSetting table, "Plot.Dimensions.Placement", CStr(side)
            Audit03SetPlotSetting table, "Plot.RebarLabels.Placement", CStr(side)
            settings.LoadFromRange table
            For Each sign In Array(1#, -1#)
                For Each degree In degrees
                    count = count + 1: angle = CDbl(degree) * GEOM_PI / 180#
                    dx = Cos(angle): dy = Sin(angle): nx = -dy * CDbl(sign): ny = dx * CDbl(sign)
                    Set layout = New CPlotAnnotationLayout: layout.ConfigureFromSettings settings
                    layout.Initialize -200#, 200#, -200#, 200#, 80#, 80#, 640#, CDbl(aspect), "mm"
                    layout.AddDimension -80# * dx, -80# * dy, -20# * dx, -20# * dy, nx, ny, "L=60"
                    layout.AddRebarLabel 20# * dx, 20# * dy, 80# * dx, 80# * dy, nx, ny, "8" & ChrW$(216) & "20"
                    Print #file, "CASE" & vbTab & CStr(count) & vbTab & CStr(degree) & vbTab & CStr(aspect) & vbTab & CStr(side) & vbTab & CStr(sign)
                    For i = 1 To layout.Count
                        values = Array(i, layout.ItemKind(i), layout.X1(i), layout.Y1(i), layout.X2(i), layout.Y2(i), layout.Rotation(i), layout.FontSize(i), layout.Text(i))
                        Print #file, "LAYOUT" & vbTab & GoalNativeAngleValues(values)
                    Next i
                    previous = vbNullString
                    For i = 1 To 2
                        plotter.GoalAngleRenderForTests chart, layout
                        snapshot = GoalNativeShapeSnapshot(chart.Chart)
                        If i = 1 Then previous = snapshot Else If snapshot <> previous Then Err.Raise 5, , "Повторная отрисовка изменила координаты."
                    Next i
                    Print #file, snapshot
                    fileBase = directory & Application.PathSeparator & "Case" & CStr(count)
                    Application.DisplayAlerts = False: book.SaveAs fileBase & ".xlsx", 51
                    sheet.Activate: chart.Activate: chart.Chart.Refresh
                    If Not chart.Chart.Export(fileBase & ".png", "PNG") Then Err.Raise 5, , "Не получено изображение схемы."
                Next degree
            Next sign
        Next side
    Next aspect
    Close #file: file = 0: book.Close False
    GoalNativeAngleProbe = "TOTAL_NATIVE_ANGLE_PROBE: passed=" & CStr(count) & "; failed=0"
    Exit Function
Failed:
    GoalNativeAngleProbe = "FAIL: native angle probe; " & CStr(Err.Number) & "; " & Err.Description
    On Error Resume Next
    If file <> 0 Then Close #file
    If Not book Is Nothing Then book.Close False
End Function

Private Function GoalNativeAngleValues(ByVal values As Variant) As String
    Dim value As Variant, text As String
    For Each value In values
        If Len(text) > 0 Then text = text & vbTab
        If VarType(value) = vbDouble Or VarType(value) = vbSingle Then text = text & Replace(CStr(value), ",", ".") Else text = text & CStr(value)
    Next value
    GoalNativeAngleValues = text
End Function

Private Function GoalNativeShapeSnapshot(ByVal chart As Object) As String
    Dim shape As Object, values As Variant, text As String
    For Each shape In chart.Shapes
        values = Array(shape.Name, shape.Type, shape.Left, shape.Top, shape.Width, shape.Height, shape.Rotation)
        text = text & "SHAPE" & vbTab & GoalNativeAngleValues(values)
        If shape.Type = 17 Then
            values = Array(shape.TextFrame2.Orientation, shape.TextFrame.Characters.Text, shape.TextFrame2.TextRange.BoundWidth, shape.TextFrame2.TextRange.BoundHeight)
            text = text & vbTab & GoalNativeAngleValues(values)
        End If
        text = text & vbCrLf
    Next shape
    GoalNativeShapeSnapshot = text
End Function
