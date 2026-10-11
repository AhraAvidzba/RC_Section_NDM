Attribute VB_Name = "modTestShapeArchitecture"
Option Explicit

' Сохраняем все биты Double, а не 15 разрядов отображения CStr.
Private Type TSnapshotDouble
    value As Double
End Type
Private Type TSnapshotWords
    low As Long
    high As Long
End Type

' ДЛЯ ТЕСТОВ: независимый снимок сетки, всей арматуры и всех аннотаций.
' Этот entrypoint выполняется и на книге до рефакторинга, и на новой книге.
' Вход восстанавливается; расчетные Results не записываются.
Public Function DumpShapeArchitectureSnapshot(Optional ByVal fixtureFilter As String = "") As String
    Dim config As Object, saved As Variant, result As String, fixture As Variant
    Dim calculation As Long, events As Boolean, updating As Boolean, found As Boolean
    On Error GoTo Failed
    calculation = Application.Calculation: events = Application.EnableEvents: updating = Application.ScreenUpdating
    Application.EnableEvents = False: Application.ScreenUpdating = False
    Set config = ThisWorkbook.Worksheets("Config").UsedRange
    saved = config.Formula
    result = "SNAPSHOT_FORMAT: 2; doubles=IEEE754-hex; units=mm,mm2,mm4,rad,N,Nmm,MPa" & vbCrLf
    For Each fixture In Array("Circle", "Rectangle", "LSection", "TwoRectangles", _
            "RoundedSimple", "RoundedTapered", "RoundedMixed", "HollowCentered", "HollowOffset")
        If Len(fixtureFilter) = 0 Or CStr(fixture) = fixtureFilter Then
            found = True
            result = result & ShapeSnapshot(CStr(fixture))
        End If
    Next fixture
    If Not found Then Err.Raise vbObjectError + 5292, "modTestShapeArchitecture", "Неизвестный fixture: " & fixtureFilter
    GoTo Finished
Failed:
    result = result & "FAIL: snapshot; " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Finished:
    If Not config Is Nothing And Not IsEmpty(saved) Then config.Formula = saved
    Application.Calculation = calculation: Application.EnableEvents = events: Application.ScreenUpdating = updating
    DumpShapeArchitectureSnapshot = result
End Function

Private Function ShapeSnapshot(ByVal fixture As String, Optional ByVal suppliedModel As CSectionModel = Nothing, _
        Optional ByVal suppliedMesh As CFiberMeshBuilder = Nothing, Optional ByVal suppliedRebars As CRebarLayout = Nothing) As String
    Dim model As CSectionModel, properties As CSectionPropertiesCalculator, annotations As CSectionAnnotations
    Dim contours As CSectionContours, lines() As String, count As Long, i As Long, bounds As Variant
    Set model = suppliedModel
    If model Is Nothing Then Set model = FixtureModel(fixture)
    Set properties = New CSectionPropertiesCalculator: properties.PrepareSnapshot model
    Set contours = model.Contours: Set annotations = model.Annotations
    ReDim lines(1 To model.ConcreteCount + model.RebarCount + contours.Count + annotations.Count + 4)
    count = count + 1
    lines(count) = Row(fixture, "MODEL", "source", Array(model.SourceType, model.ConcreteMeshSource, _
        model.ConcreteCount, model.RebarCount, contours.Count, annotations.Count))
    count = count + 1
    lines(count) = Row(fixture, "PROPS", "concrete", Array(properties.Area, properties.StaticMomentX, _
        properties.StaticMomentY, properties.CentroidX, properties.CentroidY, properties.Ix, properties.Iy, properties.Ixy, _
        properties.Ixc, properties.Iyc, properties.Ixyc, properties.PrincipalI1, properties.PrincipalI2, _
        properties.PrincipalAngleRad, properties.RadiusX, properties.RadiusY, properties.PrincipalRadius1, properties.PrincipalRadius2))
    bounds = properties.SnapshotBounds
    count = count + 1: lines(count) = Row(fixture, "BOUNDS", "concrete", bounds)
    For i = 1 To model.ConcreteCount
        count = count + 1
        lines(count) = Row(fixture, "CONCRETE", model.ConcreteID(i), Array(model.ConcreteX(i), model.ConcreteY(i), _
            model.ConcreteArea(i), model.ConcreteMaterialID(i), model.ConcreteShapeType(i), model.ConcreteWidth(i), _
            model.ConcreteHeight(i), model.ConcreteRotation(i), model.ConcreteLocalIx(i), model.ConcreteLocalIy(i), _
            model.ConcreteLocalIxy(i), model.ConcreteSourceName(i), model.ConcreteSourceHandle(i), model.ConcreteComment(i)))
    Next i
    For i = 1 To contours.Count
        count = count + 1
        lines(count) = Row(fixture, "CONTOUR", contours.ContourID(i), Array(contours.SegmentType(i), contours.LoopID(i), _
            contours.LoopRole(i), contours.StartX(i), contours.StartY(i), contours.EndX(i), contours.EndY(i), _
            contours.CenterX(i), contours.CenterY(i), contours.Radius(i), contours.SweepAngle(i), contours.SourceID(i), contours.Comment(i)))
    Next i
    For i = 1 To model.RebarCount
        count = count + 1
        lines(count) = Row(fixture, "BAR", model.RebarID(i), Array(model.RebarX(i), model.RebarY(i), _
            model.RebarDiameter(i), model.RebarArea(i), model.RebarMaterialID(i), model.RebarSteelClass(i), _
            model.RebarSourceName(i), model.RebarSourceHandle(i), model.RebarComment(i)))
    Next i
    For i = 1 To annotations.Count
        count = count + 1
        lines(count) = Row(fixture, "ANNOTATION", annotations.AnnotationID(i), Array(annotations.AnnotationType(i), _
            annotations.StartX(i), annotations.StartY(i), annotations.EndX(i), annotations.EndY(i), _
            annotations.NormalX(i), annotations.NormalY(i), annotations.Text(i), annotations.Value(i), annotations.Unit(i), annotations.Comment(i)))
    Next i
    count = count + 1
    If suppliedMesh Is Nothing Then
        lines(count) = SourceSnapshot(fixture)
    Else
        lines(count) = SourceValuesSnapshot(fixture, suppliedMesh, suppliedRebars)
    End If
    ReDim Preserve lines(1 To count)
    ShapeSnapshot = Join(lines, vbNullString)
End Function

' ДЛЯ ТЕСТОВ: только фактический результат новой сборки, без прежних builders.
Public Function PreparedShapeSnapshot(ByVal fixture As String, ByVal model As CSectionModel, _
        ByVal mesh As CFiberMeshBuilder, ByVal rebars As CRebarLayout) As String
    If model Is Nothing Or mesh Is Nothing Or rebars Is Nothing Then Err.Raise vbObjectError + 5292, "modTestShapeArchitecture", "Неполный prepared snapshot."
    PreparedShapeSnapshot = ShapeSnapshot(fixture, model, mesh, rebars)
End Function

' ДЛЯ ТЕСТОВ: исходные fill factor и группы layout не хранятся в model.
' Читаем их отдельно тем же исходным pipeline; это не замер производительности.
Private Function SourceSnapshot(ByVal fixture As String) As String
    Dim settings As CSystemSettingsReader, units As CUnitSystem, registry As CSectionTypeRegistry
    Dim geometry As CGeometryRegion, mesh As CFiberMeshBuilder, rebars As CRebarLayout
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    Set registry = New CSectionTypeRegistry
    Set geometry = registry.CreateGeometry(settings, units)
    Set mesh = BuildConfiguredConcreteMesh(geometry, settings, units)
    Set rebars = registry.CreateRebars(geometry, settings, units)
    SourceSnapshot = SourceValuesSnapshot(fixture, mesh, rebars)
End Function

Private Function SourceValuesSnapshot(ByVal fixture As String, ByVal mesh As CFiberMeshBuilder, ByVal rebars As CRebarLayout) As String
    Dim lines() As String, count As Long, i As Long
    ReDim lines(1 To mesh.FiberCount + rebars.Count + rebars.AnnotationCount + 1)
    For i = 1 To mesh.FiberCount
        count = count + 1
        lines(count) = Row(fixture, "FIBER", CStr(i), Array(mesh.FiberX(i), mesh.FiberY(i), mesh.FiberArea(i), _
            mesh.FiberFillFactor(i), mesh.FiberWidth(i), mesh.FiberHeight(i), mesh.FiberMaterialID(i)))
    Next i
    For i = 1 To rebars.Count
        count = count + 1
        lines(count) = Row(fixture, "LAYOUT", rebars.BarID(i), Array(rebars.X(i), rebars.Y(i), rebars.Diameter(i), _
            rebars.Area(i), rebars.SteelClass(i), rebars.BarAnnotationGroupName(i), rebars.Comment(i)))
    Next i
    For i = 1 To rebars.AnnotationCount
        count = count + 1
        lines(count) = Row(fixture, "ANCHOR", rebars.AnnotationGroupName(i), Array(rebars.AnnotationStartX(i), _
            rebars.AnnotationStartY(i), rebars.AnnotationEndX(i), rebars.AnnotationEndY(i), _
            rebars.AnnotationNormalX(i), rebars.AnnotationNormalY(i), rebars.AnnotationAxisDistance(i)))
    Next i
    count = count + 1: lines(count) = InputSnapshot(fixture)
    SourceValuesSnapshot = Join(lines, vbNullString)
End Function

Private Function InputSnapshot(ByVal fixture As String) As String
    Dim name As Variant, table As Object, values As Variant, formulas As Variant, i As Long, j As Long
    Dim lines() As String, count As Long, capacity As Long
    For Each name In Array("rngSystemSettings", "rngUnitSettings", "rngSignConventionSettings", _
            "rngConcreteMaterialParameters", "rngSteelMaterialParameters", "rngCircleGeometry", _
            "rngRectSetGeometry", "rngRoundedRectangleGeometry", "rngHollowRectangleGeometry", "rngCalculationProfiles")
        Set table = ThisWorkbook.Names(CStr(name)).RefersToRange
        values = table.Value2: formulas = table.Formula
        capacity = capacity + table.Rows.Count * table.Columns.Count
        ReDim Preserve lines(1 To capacity)
        For i = 1 To table.Rows.Count
            For j = 1 To table.Columns.Count
                If Not IsEmpty(values(i, j)) Then
                    count = count + 1
                    lines(count) = Row(fixture, "INPUT", CStr(name) & "." & CStr(i) & "." & CStr(j), _
                        Array(values(i, j), formulas(i, j)))
                End If
            Next j
        Next i
    Next name
    ReDim Preserve lines(1 To count)
    InputSnapshot = Join(lines, vbNullString)
End Function

Private Function FixtureModel(ByVal fixture As String) As CSectionModel
    ConfigureShapeFixture fixture
    Dim settings As CSystemSettingsReader, units As CUnitSystem, registry As CSectionTypeRegistry
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    Set registry = New CSectionTypeRegistry
    Set FixtureModel = registry.BuildGeneratedModel(settings, units)
End Function

' ДЛЯ ТЕСТОВ: одинаковый ввод baseline/new, без сборки геометрии.
Public Sub ConfigureShapeFixture(ByVal fixture As String)
    Dim settings As CSystemSettingsReader, units As CUnitSystem, registry As CSectionTypeRegistry
    Dim shapeName As String, table As Object, factor As Double, i As Long, result As String
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    factor = units.InputLengthToInternal(1#)
    SetSetting "Geometry.Source", "Generated"
    SetSetting "Mesh.StepX", 40# / factor: SetSetting "Mesh.StepY", 30# / factor
    SetSetting "Mesh.BoundarySubdivisions", 2
    Select Case fixture
        Case "Circle"
            shapeName = "Circle"
            SetSetting "Circle.Diameter", 600# / factor
            SetSetting "Rebar.AxisDistance", 40# / factor
            SetSetting "Rebar.Count", 12: SetSetting "Rebar.Diameter", 20# / factor
            SetSetting "Rebar.Diameter2", 16# / factor: SetSetting "Rebar.Diameter3", 12# / factor
            SetSetting "Rebar.Loc2row", "Stacked": SetSetting "Rebar.Loc3row", "SideBySide"
        Case "Rectangle", "LSection", "TwoRectangles"
            shapeName = "RectSet": Set table = ThisWorkbook.Names("rngRectSetGeometry").RefersToRange
            table.Cells(3, 2).Value2 = fixture: table.Cells(4, 2).Value2 = -50# / factor
            table.Cells(8, 1).Value2 = 400# / factor: table.Cells(8, 2).Value2 = 300# / factor
            table.Cells(8, 3).Value2 = 200# / factor: table.Cells(8, 4).Value2 = 600# / factor
        Case "RoundedSimple", "RoundedTapered", "RoundedMixed", "RoundedMixedLeft"
            shapeName = "RoundedRectangle": Set table = ThisWorkbook.Names("rngRoundedRectangleGeometry").RefersToRange
            table.Cells(3, 2).Value2 = 600# / factor: table.Cells(3, 3).Value2 = 400# / factor
            table.Cells(7, 2).Value2 = IIf(fixture = "RoundedTapered" Or fixture = "RoundedMixedLeft", "Tapered", "Simple")
            table.Cells(7, 3).Value2 = IIf(fixture = "RoundedSimple" Or fixture = "RoundedMixedLeft", "Simple", "Tapered")
            table.Cells(8, 2).Value2 = 100# / factor: table.Cells(8, 3).Value2 = 130# / factor
            table.Cells(9, 2).Value2 = 30# / factor: table.Cells(9, 3).Value2 = 40# / factor
            table.Cells(10, 2).Value2 = 20# / factor: table.Cells(10, 3).Value2 = 25# / factor
        Case "HollowCentered", "HollowOffset"
            shapeName = "HollowRectangle": Set table = ThisWorkbook.Names("rngHollowRectangleGeometry").RefersToRange
            table.Cells(7, 1).Value2 = 1200# / factor: table.Cells(7, 2).Value2 = 1000# / factor
            table.Cells(7, 3).Value2 = 60# / factor: table.Cells(7, 4).Value2 = 500# / factor
            table.Cells(7, 5).Value2 = 400# / factor: table.Cells(7, 6).Value2 = 30# / factor
            table.Cells(3, 2).Value2 = IIf(fixture = "HollowOffset", 40#, 0#) / factor
            table.Cells(4, 2).Value2 = IIf(fixture = "HollowOffset", 50#, 0#) / factor
    End Select
    SetSetting "Geometry.Type", shapeName
    ThisWorkbook.Worksheets("Config").Calculate
End Sub

' ДЛЯ ТЕСТОВ: расширенный baseline Circle до переноса на рецепт.
' Каждая схема дополнительных рядов проверяется с четным/нечетным n.
Public Function DumpCircleGeometryMatrix() As String
    Dim config As Object, saved As Variant, result As String, seed As CSectionModel
    Dim counts As Variant, number As Variant, scenario As Long, fixture As String
    Dim settings As CSystemSettingsReader, units As CUnitSystem, registry As CSectionTypeRegistry, model As CSectionModel
    Dim calculation As Long, events As Boolean, updating As Boolean, factor As Double
    On Error GoTo Failed
    calculation = Application.Calculation: events = Application.EnableEvents: updating = Application.ScreenUpdating
    Application.EnableEvents = False: Application.ScreenUpdating = False
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = config.Formula
    Set seed = FixtureModel("Circle")
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings: factor = units.InputLengthToInternal(1#)
    counts = Array(2, 3, 12, 13)
    For Each number In counts
        For scenario = 0 To 6
            SetSetting "Rebar.Count", CLng(number)
            SetSetting "Rebar.Diameter2", IIf(scenario = 0 Or scenario = 2, 0#, 16# / factor)
            SetSetting "Rebar.Diameter3", IIf(scenario = 0 Or scenario = 1, 0#, 12# / factor)
            SetSetting "Rebar.Loc2row", IIf(scenario = 5 Or scenario = 6, "SideBySide", "Stacked")
            SetSetting "Rebar.Loc3row", IIf(scenario = 4 Or scenario = 6, "SideBySide", "Stacked")
            ThisWorkbook.Worksheets("Config").Calculate
            settings.LoadFromWorkbook ThisWorkbook: units.LoadFromSettings settings
            Set registry = New CSectionTypeRegistry: Set model = registry.BuildGeneratedModel(settings, units)
            fixture = "CircleMatrix.n" & CStr(number) & ".rows" & CStr(scenario)
            result = result & ShapeSnapshot(fixture, model)
        Next scenario
    Next number
    GoTo Finished
Failed:
    result = result & "FAIL: Circle matrix; " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Finished:
    If Not config Is Nothing And Not IsEmpty(saved) Then config.Formula = saved
    Application.Calculation = calculation: Application.EnableEvents = events: Application.ScreenUpdating = updating
    DumpCircleGeometryMatrix = result
End Function

' ДЛЯ ТЕСТОВ: ввод расширенной RectSet-матрицы без сборки формы.
Public Sub ConfigureRectSetScenario(ByVal fixture As String, ByVal scenario As Long)
    Dim table As Object, settings As CSystemSettingsReader, units As CUnitSystem, factor As Double
    Dim i As Long, row As Long, extra As Long, count As Long
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    factor = units.InputLengthToInternal(1#)
    Set table = ThisWorkbook.Names("rngRectSetGeometry").RefersToRange
    If scenario = 0 Then Exit Sub
    For i = 1 To 8
        row = 10 + i: extra = 20 + i
        If fixture = "Rectangle" And i > 4 Then GoTo InactiveSide
        count = 5
        If scenario = 2 Then count = 3
        If scenario = 4 Then count = 1
        table.Cells(row, 4).Value2 = count
        table.Cells(row, 5).Value2 = 70# / factor
        table.Cells(row, 6).Value2 = 90# / factor
        table.Cells(extra, 2).Value2 = IIf(scenario = 4, 0#, 16# / factor)
        table.Cells(extra, 5).Value2 = 12# / factor
        table.Cells(extra, 3).Value2 = IIf(scenario = 2 Or (scenario = 3 And i Mod 2 = 0), "SideBySide", "Stacked")
        table.Cells(extra, 6).Value2 = IIf(scenario = 2 Or (scenario = 3 And i Mod 2 = 1), "SideBySide", "Stacked")
        table.Cells(extra, 4).Value2 = IIf(scenario = 2 Or i Mod 2 = 0, "EverySecondBar", "EachBar")
        table.Cells(extra, 7).Value2 = IIf(scenario = 4 Or i Mod 2 = 1, "EverySecondBar", "EachBar")
        If scenario = 4 Then table.Cells(extra, 3).Value2 = "inactive value"
InactiveSide:
    Next i
    If scenario = 5 Then
        table.Cells(14, 4).Value2 = 0
        table.Cells(14, 2).Value2 = "inactive value"
        table.Cells(14, 5).Value2 = "inactive value"
        table.Cells(24, 2).Value2 = "inactive value"
        If fixture = "Rectangle" Then
            table.Cells(8, 3).Value2 = "inactive value": table.Cells(8, 4).Value2 = "inactive value"
            For i = 15 To 18: table.Cells(i, 3).Value2 = "inactive value": Next i
        End If
        If fixture <> "TwoRectangles" Then table.Cells(4, 2).Value2 = "inactive value"
    End If
    ThisWorkbook.Worksheets("Config").Calculate
End Sub

' ДЛЯ ТЕСТОВ: неизменный старый pipeline в отдельной baseline-книге.
Public Function DumpRectSetGeometryMatrix() As String
    Dim config As Object, saved As Variant, result As String, fixture As Variant, scenario As Long
    Dim settings As CSystemSettingsReader, units As CUnitSystem, registry As CSectionTypeRegistry, model As CSectionModel
    Dim calculation As Long, events As Boolean, updating As Boolean
    On Error GoTo Failed
    calculation = Application.Calculation: events = Application.EnableEvents: updating = Application.ScreenUpdating
    Application.EnableEvents = False: Application.ScreenUpdating = False
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = config.Formula
    For Each fixture In Array("Rectangle", "LSection", "TwoRectangles")
        For scenario = 0 To 5
            config.Formula = saved
            ConfigureShapeFixture "Circle": ConfigureShapeFixture CStr(fixture)
            ConfigureRectSetScenario CStr(fixture), scenario
            Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
            Set units = New CUnitSystem: units.LoadFromSettings settings
            Set registry = New CSectionTypeRegistry: Set model = registry.BuildGeneratedModel(settings, units)
            result = result & ShapeSnapshot("RectMatrix." & CStr(fixture) & ".case" & CStr(scenario), model)
        Next scenario
    Next fixture
    GoTo Finished
Failed:
    result = result & "FAIL: RectSet baseline matrix; " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Finished:
    If Not config Is Nothing And Not IsEmpty(saved) Then config.Formula = saved
    Application.Calculation = calculation: Application.EnableEvents = events: Application.ScreenUpdating = updating
    DumpRectSetGeometryMatrix = result
End Function

' ДЛЯ ТЕСТОВ: одинаковая матрица двух независимых сторон и активных дуг.
Public Sub ConfigureRoundedScenario(ByVal scenario As Long)
    Dim table As Object, data As Variant, settings As CSystemSettingsReader, units As CUnitSystem
    Dim factor As Double, mainHeader As Long, extraHeader As Long, i As Long, row As Long, extra As Long
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings: factor = units.InputLengthToInternal(1#)
    Set table = ThisWorkbook.Names("rngRoundedRectangleGeometry").RefersToRange: data = table.Value2
    mainHeader = FixtureHeaderRow(data, "Грань", 1): extraHeader = FixtureHeaderRow(data, "d2", 2)
    If scenario = 0 Then Exit Sub
    If scenario = 1 Then
        table.Cells(9, 2).Value2 = 80# / factor: table.Cells(9, 3).Value2 = 100# / factor
        table.Cells(10, 2).Value2 = 40# / factor: table.Cells(10, 3).Value2 = 50# / factor
    ElseIf scenario = 2 Then
        table.Cells(9, 2).Value2 = 0#: table.Cells(9, 3).Value2 = 0#
        table.Cells(10, 2).Value2 = 0#: table.Cells(10, 3).Value2 = 0#
    End If
    For i = 1 To 4
        row = mainHeader + i: extra = extraHeader + i
        table.Cells(row, 2).Value2 = (35# + 5# * CDbl(i)) / factor
        table.Cells(row, 3).Value2 = 20# / factor
        table.Cells(row, 4).Value2 = IIf(i <= 2, 7, 5)
        table.Cells(extra, 2).Value2 = 16# / factor: table.Cells(extra, 5).Value2 = 12# / factor
        table.Cells(extra, 3).Value2 = IIf(scenario = 3 And i Mod 2 = 0, "SideBySide", "Stacked")
        table.Cells(extra, 6).Value2 = IIf(scenario = 3 And i Mod 2 = 1, "SideBySide", "Stacked")
        table.Cells(extra, 4).Value2 = IIf(i Mod 2 = 0, "EverySecondBar", "EachBar")
        table.Cells(extra, 7).Value2 = IIf(i Mod 2 = 1, "EverySecondBar", "EachBar")
        If scenario = 4 Then
            table.Cells(row, 4).Value2 = 1: table.Cells(extra, 2).Value2 = 0#
            table.Cells(extra, 3).Value2 = "inactive value": table.Cells(extra, 4).Value2 = "inactive value"
        End If
    Next i
    If scenario = 5 Then
        table.Cells(mainHeader + 3, 4).Value2 = 0
        table.Cells(mainHeader + 3, 2).Value2 = "inactive value"
        table.Cells(extraHeader + 3, 2).Value2 = "inactive value"
        For i = 2 To 3
            If CStr(table.Cells(7, i).Value2) = "Simple" Then
                table.Cells(8, i).Value2 = "inactive value": table.Cells(10, i).Value2 = "inactive value"
            End If
        Next i
    End If
    ThisWorkbook.Worksheets("Config").Calculate
End Sub

Private Function FixtureHeaderRow(ByRef data As Variant, ByVal caption As String, ByVal column As Long) As Long
    Dim row As Long
    For row = LBound(data, 1) To UBound(data, 1)
        If StrComp(Trim$(CStr(data(row, column))), caption, vbTextCompare) = 0 Then FixtureHeaderRow = row: Exit Function
    Next row
    Err.Raise vbObjectError + 5590, "modTestShapeArchitecture", "Не найдена шапка fixture: " & caption
End Function

' ДЛЯ ТЕСТОВ: фиксирует старую расчетную ломаную до аналитического переноса.
Public Function DumpRoundedGeometryMatrix() As String
    Dim config As Object, saved As Variant, result As String, fixture As Variant, scenario As Long
    Dim settings As CSystemSettingsReader, units As CUnitSystem, registry As CSectionTypeRegistry, model As CSectionModel
    Dim calculation As Long, events As Boolean, updating As Boolean
    On Error GoTo Failed
    calculation = Application.Calculation: events = Application.EnableEvents: updating = Application.ScreenUpdating
    Application.EnableEvents = False: Application.ScreenUpdating = False
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = config.Formula
    For Each fixture In Array("RoundedSimple", "RoundedTapered", "RoundedMixed", "RoundedMixedLeft")
        For scenario = 0 To 5
            config.Formula = saved
            ConfigureShapeFixture "Circle": ConfigureShapeFixture CStr(fixture)
            ConfigureRoundedScenario scenario
            Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
            Set units = New CUnitSystem: units.LoadFromSettings settings
            Set registry = New CSectionTypeRegistry: Set model = registry.BuildGeneratedModel(settings, units)
            result = result & ShapeSnapshot("RoundedMatrix." & CStr(fixture) & ".case" & CStr(scenario), model)
        Next scenario
    Next fixture
    GoTo Finished
Failed:
    result = result & "FAIL: Rounded baseline matrix; " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Finished:
    If Not config Is Nothing And Not IsEmpty(saved) Then config.Formula = saved
    Application.Calculation = calculation: Application.EnableEvents = events: Application.ScreenUpdating = updating
    DumpRoundedGeometryMatrix = result
End Function

' ДЛЯ ТЕСТОВ: одинаковые параметры HollowRectangle для baseline и рецепта.
Public Sub ConfigureHollowScenario(ByVal scenario As Long)
    Dim table As Object, data As Variant, settings As CSystemSettingsReader, units As CUnitSystem
    Dim factor As Double, mainHeader As Long, extraHeader As Long, i As Long, row As Long, extra As Long
    If scenario = 0 Then Exit Sub
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings: factor = units.InputLengthToInternal(1#)
    Set table = ThisWorkbook.Names("rngHollowRectangleGeometry").RefersToRange: data = table.Value2
    mainHeader = FixtureHeaderRow(data, "as", 2): extraHeader = FixtureHeaderRow(data, "d2", 2)
    table.Cells(7, 3).Value2 = 180# / factor: table.Cells(7, 6).Value2 = 90# / factor
    If scenario = 2 Or scenario = 6 Then table.Cells(7, 6).Value2 = 0#
    If scenario = 8 Or scenario = 9 Then
        table.Cells(7, 1).Value2 = IIf(scenario = 8, 600#, 1000#) / factor
        table.Cells(7, 2).Value2 = 1000# / factor
        table.Cells(7, 3).Value2 = IIf(scenario = 8, 300#, 500#) / factor
        table.Cells(7, 4).Value2 = IIf(scenario = 8, 200#, 400#) / factor
        table.Cells(7, 5).Value2 = 400# / factor
        table.Cells(7, 6).Value2 = IIf(scenario = 8, 100#, 200#) / factor
    End If
    If scenario = 10 Then table.Cells(7, 3).Value2 = 0#: table.Cells(7, 6).Value2 = 0#
    If scenario = 3 Then
        table.Cells(3, 2).Value2 = 160# / factor: table.Cells(4, 2).Value2 = -180# / factor
    End If
    For i = 1 To 8
        row = mainHeader + i: extra = extraHeader + i
        table.Cells(row, 2).Value2 = (35# + 5# * CDbl(i)) / factor
        table.Cells(row, 3).Value2 = 20# / factor
        If i <= 4 Then table.Cells(row, 4).Value2 = IIf(i <= 2, 7, 9)
        table.Cells(extra, 2).Value2 = 16# / factor: table.Cells(extra, 5).Value2 = 12# / factor
        table.Cells(extra, 3).Value2 = IIf(scenario = 4 And i Mod 2 = 0, "SideBySide", "Stacked")
        table.Cells(extra, 6).Value2 = IIf(scenario = 4 And i Mod 2 = 1, "SideBySide", "Stacked")
        table.Cells(extra, 4).Value2 = IIf(i Mod 2 = 0, "EverySecondBar", "EachBar")
        table.Cells(extra, 7).Value2 = IIf(i Mod 2 = 1, "EverySecondBar", "EachBar")
        If scenario = 5 Then
            If i <= 4 Then table.Cells(row, 4).Value2 = 1
            table.Cells(extra, 2).Value2 = 0#
            table.Cells(extra, 3).Value2 = "inactive value": table.Cells(extra, 4).Value2 = "inactive value"
        End If
        If scenario = 6 And i = 5 Then
            table.Cells(row, 3).Value2 = 0#
            table.Cells(extra, 2).Value2 = "inactive value": table.Cells(extra, 5).Value2 = "inactive value"
        End If
        If scenario = 7 And i = 3 Then
            table.Cells(row, 4).Value2 = 0
            table.Cells(row, 2).Value2 = "inactive value"
            table.Cells(extra, 2).Value2 = "inactive value": table.Cells(extra, 5).Value2 = "inactive value"
        End If
        If (scenario = 8 Or scenario = 9) And (i = 1 Or i = 2 Or i = 5 Or i = 6) Then
            table.Cells(row, 3).Value2 = 0#: If i <= 4 Then table.Cells(row, 4).Value2 = 0
            table.Cells(extra, 2).Value2 = "inactive value": table.Cells(extra, 5).Value2 = "inactive value"
        End If
    Next i
    ThisWorkbook.Worksheets("Config").Calculate
End Sub

' ДЛЯ ТЕСТОВ: старая геометрия/проекции до переноса HollowRectangle.
Public Function DumpHollowGeometryMatrix() As String
    DumpHollowGeometryMatrix = HollowGeometryMatrix(0, 7)
End Function

' ДЛЯ ТЕСТОВ: исчезнувшие прямые при предельных радиусах и оба нулевых радиуса.
Public Function DumpHollowEdgeGeometryMatrix() As String
    DumpHollowEdgeGeometryMatrix = HollowGeometryMatrix(8, 10)
End Function

Private Function HollowGeometryMatrix(ByVal firstScenario As Long, ByVal lastScenario As Long) As String
    Dim config As Object, saved As Variant, result As String, fixture As Variant, scenario As Long
    Dim settings As CSystemSettingsReader, units As CUnitSystem, registry As CSectionTypeRegistry, model As CSectionModel
    Dim calculation As Long, events As Boolean, updating As Boolean
    On Error GoTo Failed
    calculation = Application.Calculation: events = Application.EnableEvents: updating = Application.ScreenUpdating
    Application.EnableEvents = False: Application.ScreenUpdating = False
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = config.Formula
    For Each fixture In Array("HollowCentered", "HollowOffset")
        For scenario = firstScenario To lastScenario
            config.Formula = saved
            ConfigureShapeFixture "Circle": ConfigureShapeFixture CStr(fixture)
            ConfigureHollowScenario scenario
            Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
            Set units = New CUnitSystem: units.LoadFromSettings settings
            Set registry = New CSectionTypeRegistry: Set model = registry.BuildGeneratedModel(settings, units)
            result = result & ShapeSnapshot("HollowMatrix." & CStr(fixture) & ".case" & CStr(scenario), model)
        Next scenario
    Next fixture
    GoTo Finished
Failed:
    result = result & "FAIL: Hollow baseline matrix; " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Finished:
    If Not config Is Nothing And Not IsEmpty(saved) Then config.Formula = saved
    Application.Calculation = calculation: Application.EnableEvents = events: Application.ScreenUpdating = updating
    HollowGeometryMatrix = result
End Function

' ДЛЯ ТЕСТОВ: одинаковый AddBar benchmark без Excel/solver внутри замера.
Public Function DumpRebarLayoutPerformance() As String
    Dim count As Variant, repeat As Long, i As Long, layout As CRebarLayout, started As Double, elapsed As Double
    Dim result As String
    On Error GoTo Failed
    For Each count In Array(1000, 2000, 4000, 8000)
        For repeat = 1 To 3
            Set layout = New CRebarLayout
            started = Timer
            For i = 1 To CLng(count)
                layout.AddBar "P" & CStr(i), 1000000# + CDbl(i) * 37#, -1000000# + CDbl(i Mod 17) * 43#, 20#, 0#, "Rebar"
            Next i
            elapsed = Timer - started
            If elapsed < 0# Then elapsed = elapsed + 86400#
            result = result & "LAYOUT_PERF|" & CStr(count) & "|" & CStr(repeat) & "|" & Replace$(CStr(elapsed), ",", ".") & "|" & CStr(layout.Count) & vbCrLf
        Next repeat
    Next count
    DumpRebarLayoutPerformance = result
    Exit Function
Failed:
    DumpRebarLayoutPerformance = result & "FAIL: layout performance; " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
End Function

' ДЛЯ ТЕСТОВ: расчетный эталон до/после рефакторинга без изменения Results.
' Три назначения материалов и четыре направления нагрузки на каждую форму;
' затем общий пакет несущей способности, трещинообразования и ширины СП 35.
Public Function DumpShapeCalculationSnapshot(Optional ByVal crackCode As String = "SP35") As String
    DumpShapeCalculationSnapshot = SourceCalculations(Array("Circle", "Rectangle", "LSection", "TwoRectangles", _
        "RoundedSimple", "RoundedTapered", "RoundedMixed", "HollowCentered", "HollowOffset"), crackCode)
End Function

' ДЛЯ ТЕСТОВ: исходный расчет четырех RoundedRectangle, включая точные
' контуры и прежние позиции арматуры. Вторая проба активирует радиусные пути.
Public Function DumpRoundedSourceCalculations(Optional ByVal crackCode As String = "SP35", Optional ByVal scenario As Long = 0) As String
    DumpRoundedSourceCalculations = SourceCalculations(Array("RoundedSimple", "RoundedTapered", "RoundedMixed", "RoundedMixedLeft"), crackCode, scenario)
End Function

' ДЛЯ ТЕСТОВ: отдельная baseline-книга с прежней production-сборкой Hollow.
Public Function DumpHollowSourceCalculations(Optional ByVal crackCode As String = "SP35", Optional ByVal scenario As Long = 0) As String
    DumpHollowSourceCalculations = SourceCalculations(Array("HollowCentered", "HollowOffset"), crackCode, 0, scenario)
End Function

Private Function SourceCalculations(ByVal fixtures As Variant, ByVal crackCode As String, Optional ByVal roundedScenario As Long = 0, _
        Optional ByVal hollowScenario As Long = 0) As String
    Dim config As Object, saved As Variant, result As String, fixture As Variant
    Dim calculation As Long, events As Boolean, updating As Boolean
    On Error GoTo Failed
    calculation = Application.Calculation: events = Application.EnableEvents: updating = Application.ScreenUpdating
    Application.EnableEvents = False: Application.ScreenUpdating = False
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = config.Formula
    Dim profiles As Object, i As Long, key As String
    Set profiles = ThisWorkbook.Names("rngCalculationProfiles").RefersToRange
    For i = 2 To profiles.Rows.Count
        key = CStr(profiles.Cells(i, 2).Value2)
        Select Case key
            Case "Calculation.Strength.DirectState", "Calculation.Strength.Capacity", "Calculation.Crack.Width"
                profiles.Cells(i, 3).Value2 = "Yes"
            Case "Calculation.Stability.Enabled", "Calculation.Crack.Longitudinal"
                profiles.Cells(i, 3).Value2 = "No"
        End Select
    Next i
    SetSetting "SLS.Crack.Code", crackCode
    ConfigureShapeFixture "Circle"
    For Each fixture In fixtures
        ConfigureShapeFixture CStr(fixture)
        If roundedScenario > 0 Then ConfigureRoundedScenario roundedScenario
        If hollowScenario > 0 Then ConfigureHollowScenario hollowScenario
        Dim settings As CSystemSettingsReader, units As CUnitSystem, registry As CSectionTypeRegistry, model As CSectionModel, label As String
        Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
        Set units = New CUnitSystem: units.LoadFromSettings settings
        Set registry = New CSectionTypeRegistry: Set model = registry.BuildGeneratedModel(settings, units)
        label = CStr(fixture): If roundedScenario > 0 Then label = label & ".Arc"
        If hollowScenario > 0 Then label = label & ".case" & CStr(hollowScenario)
        result = result & CalculationSnapshot(label, model)
    Next fixture
    GoTo Finished
Failed:
    result = result & "FAIL: calculation snapshot; " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Finished:
    If Not config Is Nothing And Not IsEmpty(saved) Then config.Formula = saved
    Application.Calculation = calculation: Application.EnableEvents = events: Application.ScreenUpdating = updating
    SourceCalculations = result
End Function

Private Function CalculationSnapshot(ByVal fixture As String, Optional ByVal suppliedModel As CSectionModel = Nothing) As String
    Dim model As CSectionModel, props As CSectionPropertiesCalculator
    Dim settings As CSystemSettingsReader, units As CUnitSystem, provider As CMaterialModelProvider
    Set model = suppliedModel
    If model Is Nothing Then Set model = FixtureModel(fixture)
    Set props = New CSectionPropertiesCalculator: props.CalculateConcrete model
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    Set provider = New CMaterialModelProvider: provider.Initialize settings, units
    Dim purpose As Variant, loadIndex As Long, solver As CSectionSolver, result As String
    Dim n As Double, mx As Double, my As Double, minX As Double, maxX As Double, minY As Double, maxY As Double
    props.CalculateProjection model, 1#, 0#, False, minX, maxX
    props.CalculateProjection model, 0#, 1#, False, minY, maxY
    For Each purpose In Array(cpStrength, cpMcrc, cpCrackedNDS)
        For loadIndex = 0 To 3
            n = -2# * props.Area
            mx = n * props.CentroidY: my = n * props.CentroidX
            If loadIndex = 1 Or loadIndex = 3 Then mx = mx + props.Area * (maxY - minY) * 0.2
            If loadIndex = 2 Or loadIndex = 3 Then my = my - props.Area * (maxX - minX) * 0.15
            Set solver = New CSectionSolver: solver.ApplySettings settings, units
            solver.Solve model, provider.ConcreteMaterial(CLng(purpose)), provider.SteelMaterial(CLng(purpose)), n, mx, my
            result = result & Row(fixture, "NDS", CStr(purpose) & "." & CStr(loadIndex), Array( _
                solver.Converged, solver.Epsilon0, solver.KappaX, solver.KappaY, solver.Nint, solver.Mxint, solver.Myint, _
                solver.MinConcreteStress, solver.MaxConcreteStress, solver.MinSteelStress, solver.MaxSteelStress, _
                solver.ResidualN, solver.ResidualMx, solver.ResidualMy, solver.MinConcreteStrain, solver.MaxConcreteStrain, _
                solver.MinSteelStrain, solver.MaxSteelStrain))
            result = result & Row(fixture, "LOAD", CStr(purpose) & "." & CStr(loadIndex), _
                Array(n, mx, my, props.CentroidX, props.CentroidY))
            If solver.Converged Then result = result & ElementStateSnapshot(fixture, CStr(purpose) & "." & CStr(loadIndex), _
                model, solver, provider.ConcreteMaterial(CLng(purpose)), provider.SteelMaterial(CLng(purpose)))
        Next loadIndex
    Next purpose
    Dim profiles As CCalculationProfileCatalog, batch As CBatchSectionCalculator, combination As CCombinationResult
    Set profiles = New CCalculationProfileCatalog: profiles.LoadFromWorkbook ThisWorkbook
    Set batch = New CBatchSectionCalculator: batch.Initialize model, provider
    Set batch.ProfileCatalog = profiles: batch.ApplySettings settings, units
    batch.ApplyLoadReference props.CentroidX, props.CentroidY, 0#, 0#
    batch.AddCombination "COMP", -2# * props.Area, 0#, 0#, "PR1", "Architecture control", "Auto"
    batch.AddCombination "BEND", -0.2 * props.Area, props.Area * (maxY - minY) * 1.6, _
        -props.Area * (maxX - minX) * 0.6, "PR1", "Architecture control", "Auto"
    batch.Execute
    For loadIndex = 1 To batch.Count
        Set combination = batch.ResultAt(loadIndex)
        result = result & Row(fixture, "BATCH", CStr(loadIndex), Array(combination.Status, combination.NormalCrackStatus, _
            combination.StrengthResult.Capacity.ResultMeta.InternalStatus, combination.CrackResult.Formation.ResultMeta.InternalStatus, _
            combination.CrackResult.Width.ResultMeta.Calculated, combination.StrengthResult.Capacity.LambdaCapacity, _
            combination.CrackResult.Formation.LambdaCrc, combination.CrackResult.Width.CrackWidth, combination.CrackResult.Width.Abt))
    Next loadIndex
    CalculationSnapshot = result
End Function

' ДЛЯ ТЕСТОВ: расчеты только переданной новой модели.
Public Function PreparedCalculationSnapshot(ByVal fixture As String, ByVal model As CSectionModel) As String
    If model Is Nothing Then Err.Raise vbObjectError + 5292, "modTestShapeArchitecture", "Расчетная модель не передана."
    PreparedCalculationSnapshot = CalculationSnapshot(fixture, model)
End Function

' ДЛЯ ТЕСТОВ: фактическая плоскость и выбранные материалы для каждого элемента.
' StrainAt сохраняет численный frame solver-а, включая удаленные координаты.
Private Function ElementStateSnapshot(ByVal fixture As String, ByVal stateID As String, ByVal model As CSectionModel, _
        ByVal solver As CSectionSolver, ByVal concrete As Object, ByVal steel As Object) As String
    Dim lines() As String, i As Long, strain As Double
    ReDim lines(1 To model.ConcreteCount + model.RebarCount)
    For i = 1 To model.ConcreteCount
        strain = solver.StrainAt(model.ConcreteX(i), model.ConcreteY(i))
        lines(i) = Row(fixture, "FIBER_STATE", stateID & "." & model.ConcreteID(i), Array(strain, concrete.StressAtStrain(strain)))
    Next i
    For i = 1 To model.RebarCount
        strain = solver.StrainAt(model.RebarX(i), model.RebarY(i))
        lines(model.ConcreteCount + i) = Row(fixture, "BAR_STATE", stateID & "." & model.RebarID(i), Array(strain, steel.StressAtStrain(strain)))
    Next i
    ElementStateSnapshot = Join(lines, vbNullString)
End Function

' ДЛЯ ТЕСТОВ: реальные Chart/Shapes и PNG геометрических эталонов.
' Results меняется только в закрываемой без сохранения fixture-книге runner-а.
' ДЛЯ ТЕСТОВ: прежний Hollow preview для сопоставления длинных подписей.
Public Function DumpHollowSourcePreview(ByVal outputDirectory As String, Optional ByVal scenario As Long = 0) As String
    Dim config As Object, saved As Variant, result As String, fixture As Variant
    Dim settings As CSystemSettingsReader, units As CUnitSystem, registry As CSectionTypeRegistry, model As CSectionModel
    Dim calculation As Long, events As Boolean, updating As Boolean
    On Error GoTo Failed
    calculation = Application.Calculation: events = Application.EnableEvents: updating = Application.ScreenUpdating
    Application.EnableEvents = False: Application.ScreenUpdating = True
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = config.Formula
    ConfigureShapeFixture "Circle"
    For Each fixture In Array("HollowCentered", "HollowOffset")
        ConfigureShapeFixture CStr(fixture): ConfigureHollowScenario scenario
        Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
        Set units = New CUnitSystem: units.LoadFromSettings settings
        Set registry = New CSectionTypeRegistry: Set model = registry.BuildGeneratedModel(settings, units)
        result = result & PreviewSnapshot(CStr(fixture) & ".case" & CStr(scenario), outputDirectory, model)
    Next fixture
    GoTo Finished
Failed:
    result = result & "FAIL: source Hollow preview; " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Finished:
    If Not config Is Nothing And Not IsEmpty(saved) Then config.Formula = saved
    Application.Calculation = calculation: Application.EnableEvents = events: Application.ScreenUpdating = updating
    DumpHollowSourcePreview = result
End Function

Public Function DumpShapePreviewSnapshot(ByVal outputDirectory As String, Optional ByVal fixtureFilter As String = "") As String
    Dim config As Object, saved As Variant, calculation As Long, events As Boolean, updating As Boolean
    Dim result As String, fixture As Variant
    On Error GoTo Failed
    calculation = Application.Calculation: events = Application.EnableEvents: updating = Application.ScreenUpdating
    Application.EnableEvents = False: Application.ScreenUpdating = True
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = config.Formula
    For Each fixture In Array("Circle", "Rectangle", "LSection", "TwoRectangles", _
            "RoundedSimple", "RoundedTapered", "RoundedMixed", "HollowCentered", "HollowOffset")
        If Len(fixtureFilter) = 0 Or CStr(fixture) = fixtureFilter Then result = result & PreviewSnapshot(CStr(fixture), outputDirectory)
    Next fixture
    GoTo Finished
Failed:
    result = result & "FAIL: preview snapshot; " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Finished:
    If Not config Is Nothing And Not IsEmpty(saved) Then config.Formula = saved
    Application.Calculation = calculation: Application.EnableEvents = events: Application.ScreenUpdating = updating
    DumpShapePreviewSnapshot = result
End Function

Private Function PreviewSnapshot(ByVal fixture As String, ByVal outputDirectory As String, Optional ByVal suppliedModel As CSectionModel = Nothing) As String
    Dim stage As String
    On Error GoTo Failed
    stage = "build"
    Dim model As CSectionModel, properties As CSectionPropertiesCalculator
    Dim settings As CSystemSettingsReader, units As CUnitSystem, writer As CNDMResultsWriter
    Dim reader As CSectionPlotDataReader, plotter As CSectionPlotter
    Set model = suppliedModel
    If model Is Nothing Then Set model = FixtureModel(fixture)
    Set properties = New CSectionPropertiesCalculator: properties.PrepareSnapshot model
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    stage = "write preview"
    Set writer = New CNDMResultsWriter: writer.WriteGeometryPreview ThisWorkbook, model, properties, units
    stage = "read preview"
    Set reader = New CSectionPlotDataReader: reader.LoadGeometryPreviewFromWorkbook ThisWorkbook, settings
    stage = "draw"
    Set plotter = New CSectionPlotter: plotter.Draw ThisWorkbook, reader, settings
    Dim container As Object, chart As Object, shape As Object, series As Object
    Dim result As String, i As Long, j As Long, xs As Variant, ys As Variant, shapeText As String
    stage = "chart metrics"
    Set container = ThisWorkbook.Worksheets("Расчет").ChartObjects("chtNDMSectionPlot")
    Set chart = container.Chart
    ' Geometry-only reader использует подпись импорта; эти fixtures сгенерированы.
    chart.Shapes("NDMPlot_Title").TextFrame2.TextRange.Text = "Сгенерированная геометрия: " & fixture
    result = Row(fixture, "PLOT", "container", Array(container.Left, container.Top, container.Width, container.Height, _
        chart.PlotArea.Left, chart.PlotArea.Top, chart.PlotArea.Width, chart.PlotArea.Height, _
        chart.SeriesCollection.Count, chart.Shapes.Count))
    For i = 1 To chart.SeriesCollection.Count
        stage = "series " & CStr(i)
        Set series = chart.SeriesCollection(i)
        xs = series.XValues: ys = series.Values
        For j = LBound(xs) To UBound(xs)
            result = result & Row(fixture, "PLOT_POINT", CStr(i) & "." & CStr(j), Array(CStr(series.Name), CDbl(xs(j)), CDbl(ys(j))))
        Next j
    Next i
    For Each shape In chart.Shapes
        stage = "shape " & CStr(shape.Name)
        shapeText = vbNullString
        If CLng(shape.Type) = 17 Or CLng(shape.Type) = 1 Then
            If shape.TextFrame2.HasText Then shapeText = CStr(shape.TextFrame2.TextRange.Text)
        End If
        result = result & Row(fixture, "PLOT_SHAPE", CStr(shape.Name), Array(CLng(shape.Type), _
            CDbl(shape.Left), CDbl(shape.Top), CDbl(shape.Width), CDbl(shape.Height), CDbl(shape.Rotation), shapeText))
    Next shape
    stage = "export PNG"
    chart.Refresh
    If Not chart.Export(outputDirectory & Application.PathSeparator & fixture & ".png", "PNG") Then _
        Err.Raise vbObjectError + 5294, "modTestShapeArchitecture", "Excel не сохранил PNG: " & fixture
    PreviewSnapshot = result
    Exit Function
Failed:
    Err.Raise Err.Number, "modTestShapeArchitecture.PreviewSnapshot", fixture & "; " & stage & "; " & Err.Description
End Function

' ДЛЯ ТЕСТОВ: Results -> reader -> plotter для переданной новой модели.
Public Function PreparedPreviewSnapshot(ByVal fixture As String, ByVal model As CSectionModel, ByVal outputDirectory As String) As String
    If model Is Nothing Then Err.Raise vbObjectError + 5292, "modTestShapeArchitecture", "Модель схемы не передана."
    PreparedPreviewSnapshot = PreviewSnapshot(fixture, outputDirectory, model)
End Function

Private Function Row(ByVal fixture As String, ByVal kind As String, ByVal id As String, ByVal values As Variant) As String
    Dim item As Variant
    Row = EscapeField(fixture) & vbTab & EscapeField(kind) & vbTab & EscapeField(id)
    For Each item In values: Row = Row & vbTab & SnapshotValue(item): Next item
    Row = Row & vbCrLf
End Function

Private Function SnapshotValue(ByVal value As Variant) As String
    Dim number As TSnapshotDouble, words As TSnapshotWords
    Select Case VarType(value)
        Case vbDouble, vbSingle, vbCurrency, vbDecimal
            number.value = CDbl(value): LSet words = number
            SnapshotValue = "d:" & Right$("00000000" & Hex$(words.high), 8) & Right$("00000000" & Hex$(words.low), 8)
        Case vbByte, vbInteger, vbLong
            SnapshotValue = "i:" & CStr(value)
        Case vbBoolean
            If value Then SnapshotValue = "b:1" Else SnapshotValue = "b:0"
        Case vbEmpty
            SnapshotValue = "e:"
        Case vbNull
            SnapshotValue = "n:"
        Case vbString
            SnapshotValue = "s:" & EscapeField(CStr(value))
        Case Else
            Err.Raise vbObjectError + 5293, "modTestShapeArchitecture", "Неподдерживаемое значение snapshot: " & CStr(VarType(value))
    End Select
End Function

Private Function EscapeField(ByVal value As String) As String
    EscapeField = Replace(Replace(Replace(Replace(value, "%", "%25"), vbTab, "%09"), vbCr, "%0D"), vbLf, "%0A")
End Function

Private Sub SetSetting(ByVal key As String, ByVal value As Variant)
    Dim name As Variant, table As Object, i As Long
    For Each name In Array("rngSystemSettings", "rngCircleGeometry")
        Set table = ThisWorkbook.Names(CStr(name)).RefersToRange
        For i = 2 To table.Rows.Count
            If CStr(table.Cells(i, 1).Value2) = key Then table.Cells(i, 2).Value2 = value: Exit Sub
        Next i
    Next name
    Err.Raise vbObjectError + 5291, "modTestShapeArchitecture", "Нет ключа fixture: " & key
End Sub
