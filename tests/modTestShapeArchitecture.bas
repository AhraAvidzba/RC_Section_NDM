Attribute VB_Name = "modTestShapeArchitecture"
Option Explicit

' ==========================================================================
' ДЛЯ ТЕСТОВ: точные снимки геометрии, входных таблиц, НДС и схем для сравнения общей архитектуры без округления Double.
' ==========================================================================

' Сохраняем все биты Double, а не 15 разрядов отображения CStr.
Private Type TSnapshotDouble
    value As Double
End Type
Private Type TSnapshotWords
    low As Long
    high As Long
End Type

' -------------------- ДЛЯ ТЕСТОВ --------------------
' сравнивает сетку, всю арматуру и все аннотации с неизменяемыми
' эталонами архива. Входные данные восстанавливает; расчётные Results не записывает.
Public Function TestDumpShapeArchitectureSnapshot(Optional ByVal fixtureFilter As String = "") As String
    Dim config As Object, saved As Variant, result As String, fixture As Variant
    Dim calculation As Long, events As Boolean, updating As Boolean, found As Boolean
    On Error GoTo Failed
    calculation = Application.Calculation: events = Application.EnableEvents: updating = Application.ScreenUpdating
    Application.EnableEvents = False: Application.ScreenUpdating = False
    Set config = ThisWorkbook.Worksheets("Config").UsedRange
    saved = TestFixtureSaveInputs(config)
    result = "SNAPSHOT_FORMAT: 2; doubles=IEEE754-hex; units=mm,mm2,mm4,rad,N,Nmm,MPa" & vbCrLf
    For Each fixture In Array("Circle", "Rectangle", "LSection", "TwoRectangles", _
            "RoundedSimple", "RoundedTapered", "RoundedMixed", "HollowCentered", "HollowOffset")
        If Len(fixtureFilter) = 0 Or CStr(fixture) = fixtureFilter Then
            found = True
            result = result & TestShapeSnapshot(CStr(fixture))
        End If
    Next fixture
    If Not found Then Err.Raise vbObjectError + 5292, "modTestShapeArchitecture", "Неизвестный fixture: " & fixtureFilter
    GoTo Finished
Failed:
    result = result & "FAIL: snapshot; " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Finished:
    If Not config Is Nothing And Not IsEmpty(saved) Then TestFixtureRestoreInputs config, saved
    Application.Calculation = calculation: Application.EnableEvents = events: Application.ScreenUpdating = updating
    TestDumpShapeArchitectureSnapshot = result
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сохраняет модель, свойства, контуры, сетку, арматуру и аннотации одной готовой сборки без потери битов Double.
Private Function TestShapeSnapshot(ByVal fixture As String, Optional ByVal suppliedModel As CSectionModel = Nothing, _
        Optional ByVal suppliedMesh As CFiberMeshBuilder = Nothing, Optional ByVal suppliedRebars As CRebarLayout = Nothing) As String
    Dim model As CSectionModel, properties As CSectionPropertiesCalculator, annotations As CSectionAnnotations
    Dim contours As CSectionContours, lines() As String, count As Long, i As Long, bounds As Variant, builder As CSectionBuilder
    Set model = suppliedModel
    If model Is Nothing Then
        Set builder = New CSectionBuilder: Set model = TestFixtureModel(fixture, builder)
        Set suppliedMesh = builder.Mesh: Set suppliedRebars = builder.Rebars
    End If
    If suppliedMesh Is Nothing Or suppliedRebars Is Nothing Then Err.Raise vbObjectError + 5292, "modTestShapeArchitecture", "Для снимка нужна одна полная prepared-сборка."
    Set properties = New CSectionPropertiesCalculator: properties.PrepareSnapshot model
    Set contours = model.Contours: Set annotations = model.Annotations
    ReDim lines(1 To model.ConcreteCount + model.RebarCount + contours.Count + annotations.Count + 4) ' Размер известен заранее; строки большого снимка не наращиваются повторной конкатенацией.
    count = count + 1
    lines(count) = TestRow(fixture, "MODEL", "source", Array(model.SourceType, model.ConcreteMeshSource, _
        model.ConcreteCount, model.RebarCount, contours.Count, annotations.Count))
    count = count + 1
    lines(count) = TestRow(fixture, "PROPS", "concrete", Array(properties.Area, properties.StaticMomentX, _
        properties.StaticMomentY, properties.CentroidX, properties.CentroidY, properties.Ix, properties.Iy, properties.Ixy, _
        properties.Ixc, properties.Iyc, properties.Ixyc, properties.PrincipalI1, properties.PrincipalI2, _
        properties.PrincipalAngleRad, properties.RadiusX, properties.RadiusY, properties.PrincipalRadius1, properties.PrincipalRadius2))
    bounds = properties.SnapshotBounds
    count = count + 1: lines(count) = TestRow(fixture, "BOUNDS", "concrete", bounds)
    For i = 1 To model.ConcreteCount
        count = count + 1
        lines(count) = TestRow(fixture, "CONCRETE", model.ConcreteID(i), Array(model.ConcreteX(i), model.ConcreteY(i), _
            model.ConcreteArea(i), model.ConcreteMaterialID(i), model.ConcreteShapeType(i), model.ConcreteWidth(i), _
            model.ConcreteHeight(i), model.ConcreteRotation(i), model.ConcreteLocalIx(i), model.ConcreteLocalIy(i), _
            model.ConcreteLocalIxy(i), model.ConcreteSourceName(i), model.ConcreteSourceHandle(i), model.ConcreteComment(i)))
    Next i
    For i = 1 To contours.Count
        count = count + 1
        lines(count) = TestRow(fixture, "CONTOUR", contours.ContourID(i), Array(contours.SegmentType(i), contours.LoopID(i), _
            contours.LoopRole(i), contours.StartX(i), contours.StartY(i), contours.EndX(i), contours.EndY(i), _
            contours.CenterX(i), contours.CenterY(i), contours.Radius(i), contours.SweepAngle(i), contours.SourceID(i), contours.Comment(i)))
    Next i
    For i = 1 To model.RebarCount
        count = count + 1
        lines(count) = TestRow(fixture, "BAR", model.RebarID(i), Array(model.RebarX(i), model.RebarY(i), _
            model.RebarDiameter(i), model.RebarArea(i), model.RebarMaterialID(i), model.RebarSteelClass(i), _
            model.RebarSourceName(i), model.RebarSourceHandle(i), model.RebarComment(i)))
    Next i
    For i = 1 To annotations.Count
        count = count + 1
        lines(count) = TestRow(fixture, "ANNOTATION", annotations.AnnotationID(i), Array(annotations.AnnotationType(i), _
            annotations.StartX(i), annotations.StartY(i), annotations.EndX(i), annotations.EndY(i), _
            annotations.NormalX(i), annotations.NormalY(i), annotations.Text(i), annotations.Value(i), annotations.Unit(i), annotations.Comment(i)))
    Next i
    count = count + 1
    lines(count) = TestSourceValuesSnapshot(fixture, suppliedMesh, suppliedRebars)
    ReDim Preserve lines(1 To count)
    TestShapeSnapshot = Join(lines, vbNullString)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' только фактический результат новой сборки, без отдельных построителей типов.
Public Function TestPreparedShapeSnapshot(ByVal fixture As String, ByVal model As CSectionModel, _
        ByVal mesh As CFiberMeshBuilder, ByVal rebars As CRebarLayout) As String
    If model Is Nothing Or mesh Is Nothing Or rebars Is Nothing Then Err.Raise vbObjectError + 5292, "modTestShapeArchitecture", "Неполный prepared snapshot."
    TestPreparedShapeSnapshot = TestShapeSnapshot(fixture, model, mesh, rebars)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Дополняет снимок исходными данными сетки и раскладки, включая группы и комментарии, чтобы обнаруживать скрытые изменения сборки.
Private Function TestSourceValuesSnapshot(ByVal fixture As String, ByVal mesh As CFiberMeshBuilder, ByVal rebars As CRebarLayout) As String
    Dim lines() As String, count As Long, i As Long
    ReDim lines(1 To mesh.FiberCount + rebars.Count + 1)
    For i = 1 To mesh.FiberCount
        count = count + 1
        lines(count) = TestRow(fixture, "FIBER", CStr(i), Array(mesh.FiberX(i), mesh.FiberY(i), mesh.FiberArea(i), _
            mesh.FiberFillFactor(i), mesh.FiberWidth(i), mesh.FiberHeight(i), mesh.FiberMaterialID(i)))
    Next i
    For i = 1 To rebars.Count
        count = count + 1
        lines(count) = TestRow(fixture, "LAYOUT", rebars.BarID(i), Array(rebars.X(i), rebars.Y(i), rebars.Diameter(i), _
            rebars.Area(i), rebars.SteelClass(i), rebars.BarAnnotationGroupName(i), rebars.Comment(i)))
    Next i

    count = count + 1: lines(count) = TestInputSnapshot(fixture)
    TestSourceValuesSnapshot = Join(lines, vbNullString)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сохраняет значения и формулы всех значимых входных таблиц в текущем порядке строк и столбцов.
Private Function TestInputSnapshot(ByVal fixture As String) As String
    Dim name As Variant, table As Object, values As Variant, formulas As Variant, i As Long, j As Long
    Dim lines() As String, count As Long, capacity As Long
    For Each name In Array("rngSystemSettings", "rngUnitSettings", "rngSignConventionSettings", _
            "rngConcreteMaterialParameters", "rngSteelMaterialParameters", "rngCircleGeometry", _
            "rngRectanglesGeometry", "rngPointedSectionGeometry", "rngHollowRectangleGeometry", "rngCalculationProfiles")
        Set table = ThisWorkbook.Names(CStr(name)).RefersToRange
        values = table.Value2: formulas = table.Formula
        capacity = capacity + table.Rows.Count * table.Columns.Count
        ReDim Preserve lines(1 To capacity)
        For i = 1 To table.Rows.Count
            For j = 1 To table.Columns.Count
                If Not IsEmpty(values(i, j)) Then
                    count = count + 1
                    lines(count) = TestRow(fixture, "INPUT", CStr(name) & "." & CStr(i) & "." & CStr(j), _
                        Array(values(i, j), formulas(i, j)))
                End If
            Next j
        Next i
    Next name
    ReDim Preserve lines(1 To count)
    TestInputSnapshot = Join(lines, vbNullString)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Настраивает контрольный вариант и строит его обычным конвейером чтения настроек, единиц и геометрии.
Private Function TestFixtureModel(ByVal fixture As String, Optional ByVal builder As CSectionBuilder = Nothing) As CSectionModel
    TestConfigureShapeFixture fixture
    Dim settings As CSystemSettingsReader, units As CUnitSystem
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    Set TestFixtureModel = BuildWorkbookSectionModel(ThisWorkbook, settings, units, builder)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' одинаковый ввод baseline/new, без сборки геометрии.
Public Sub TestConfigureShapeFixture(ByVal fixture As String)
    Dim settings As CSystemSettingsReader, units As CUnitSystem, builder As CSectionBuilder
    Dim shapeName As String, table As Object, factor As Double, i As Long, result As String
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    factor = units.InputLengthToInternal(1#)
    TestSetSetting "Geometry.Source", "Generated"
    TestSetSetting "Mesh.StepMode", "User"
    TestSetSetting "Mesh.StepX", 40# / factor: TestSetSetting "Mesh.StepY", 30# / factor
    TestSetSetting "Mesh.BoundarySubdivisions", 2
    Select Case fixture
        Case "Circle"
            shapeName = "Circle"
            TestFixtureResetReinforcement shapeName
            TestSetSetting "Circle.Diameter", 600# / factor
            TestSetSetting "Rebar.AxisDistance", 40# / factor
            TestSetSetting "Rebar.Count", 12: TestSetSetting "Rebar.Diameter", 20# / factor
            TestSetSetting "Rebar.Diameter2", 16# / factor: TestSetSetting "Rebar.Diameter3", 12# / factor
            TestSetSetting "Rebar.Loc2row", "Stacked": TestSetSetting "Rebar.Loc3row", "SideBySide"
        Case "Rectangle", "LSection", "TwoRectangles"
            shapeName = "Rectangles": Set table = ThisWorkbook.Names("rngRectanglesGeometry").RefersToRange
            TestFixtureInputCell(table, "General", "R").Value2 = 0#
            TestFixtureInputCell(table, "Detail", "XOffset", "val").Value2 = IIf(fixture = "LSection", 0#, -50#) / factor
            TestFixtureInputCell(table, "General", "H1", "val").Value2 = 400# / factor: TestFixtureInputCell(table, "General", "B1", "val").Value2 = 300# / factor
            ' При as=40 с двух сторон и s0+s1=160 положительному примеру нужна H2>240; H2=400 сохраняет допустимый участок 160 мм и n=5 без изменения проверок.
            TestFixtureInputCell(table, "General", "H2", "val").Value2 = 400# / factor: TestFixtureInputCell(table, "General", "B2", "val").Value2 = IIf(fixture = "Rectangle", 0#, 600#) / factor
        Case "RoundedSimple", "RoundedTapered", "RoundedMixed", "RoundedMixedLeft"
            shapeName = "PointedSection": Set table = ThisWorkbook.Names("rngPointedSectionGeometry").RefersToRange
            TestFixtureInputCell(table, "General", "B", "val").Value2 = 600# / factor: TestFixtureInputCell(table, "General", "H", "val").Value2 = 400# / factor
            TestFixtureInputCell(table, "Detail", "W", "left").Value2 = IIf(fixture = "RoundedTapered" Or fixture = "RoundedMixedLeft", 100#, 0#) / factor
            TestFixtureInputCell(table, "Detail", "W", "right").Value2 = IIf(fixture = "RoundedSimple" Or fixture = "RoundedMixedLeft", 0#, 130#) / factor
            TestFixtureInputCell(table, "Detail", "R1", "left").Value2 = 30# / factor: TestFixtureInputCell(table, "Detail", "R1", "right").Value2 = 40# / factor
            TestFixtureInputCell(table, "Detail", "R2", "left").Value2 = 20# / factor: TestFixtureInputCell(table, "Detail", "R2", "right").Value2 = 25# / factor
        Case "HollowCentered", "HollowOffset"
            shapeName = "HollowRectangle": Set table = ThisWorkbook.Names("rngHollowRectangleGeometry").RefersToRange
            TestFixtureInputCell(table, "General", "H", "val").Value2 = 1200# / factor: TestFixtureInputCell(table, "General", "B", "val").Value2 = 1000# / factor
            TestFixtureInputCell(table, "General", "R", "val").Value2 = 60# / factor: TestFixtureInputCell(table, "Detail", "Ho", "val").Value2 = 500# / factor
            TestFixtureInputCell(table, "Detail", "Bo", "val").Value2 = 400# / factor: TestFixtureInputCell(table, "Detail", "Ro", "val").Value2 = 30# / factor
            TestFixtureInputCell(table, "Detail", "Xo", "val").Value2 = IIf(fixture = "HollowOffset", 40#, 0#) / factor
            TestFixtureInputCell(table, "Detail", "Yo", "val").Value2 = IIf(fixture = "HollowOffset", 50#, 0#) / factor
    End Select
    TestSetSetting "Geometry.Type", shapeName
    If shapeName <> "Circle" Then TestFixtureResetReinforcement shapeName
    ThisWorkbook.Worksheets("Config").Calculate
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' расширенный baseline Circle до переноса на рецепт.
' Каждая схема дополнительных рядов проверяется с четным/нечетным n.
Public Function TestDumpCircleGeometryMatrix() As String
    Dim config As Object, saved As Variant, result As String, seed As CSectionModel
    Dim counts As Variant, number As Variant, scenario As Long, fixture As String
    Dim settings As CSystemSettingsReader, units As CUnitSystem, builder As CSectionBuilder, model As CSectionModel
    Dim calculation As Long, events As Boolean, updating As Boolean, factor As Double
    On Error GoTo Failed
    calculation = Application.Calculation: events = Application.EnableEvents: updating = Application.ScreenUpdating
    Application.EnableEvents = False: Application.ScreenUpdating = False
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = TestFixtureSaveInputs(config)
    Set seed = TestFixtureModel("Circle")
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings: factor = units.InputLengthToInternal(1#)
    counts = Array(2, 3, 12, 13)
    For Each number In counts
        For scenario = 0 To 6
            TestSetSetting "Rebar.Count", CLng(number)
            TestSetSetting "Rebar.Diameter2", IIf(scenario = 0 Or scenario = 2, 0#, 16# / factor)
            TestSetSetting "Rebar.Diameter3", IIf(scenario = 0 Or scenario = 1, 0#, 12# / factor)
            TestSetSetting "Rebar.Loc2row", IIf(scenario = 5 Or scenario = 6, "SideBySide", "Stacked")
            TestSetSetting "Rebar.Loc3row", IIf(scenario = 4 Or scenario = 6, "SideBySide", "Stacked")
            ThisWorkbook.Worksheets("Config").Calculate
            settings.LoadFromWorkbook ThisWorkbook: units.LoadFromSettings settings
            Set builder = New CSectionBuilder: Set model = BuildWorkbookSectionModel(ThisWorkbook, settings, units, builder)
            fixture = "CircleMatrix.n" & CStr(number) & ".rows" & CStr(scenario)
            result = result & TestShapeSnapshot(fixture, model, builder.Mesh, builder.Rebars)
        Next scenario
    Next number
    GoTo Finished
Failed:
    result = result & "FAIL: Circle matrix; " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Finished:
    If Not config Is Nothing And Not IsEmpty(saved) Then TestFixtureRestoreInputs config, saved
    Application.Calculation = calculation: Application.EnableEvents = events: Application.ScreenUpdating = updating
    TestDumpCircleGeometryMatrix = result
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' ввод расширенной Rectangles-матрицы без сборки формы.
Public Sub TestConfigureRectanglesScenario(ByVal fixture As String, ByVal scenario As Long)
    Dim keys As Variant
    Dim table As Object, settings As CSystemSettingsReader, units As CUnitSystem, factor As Double
    Dim i As Long, row As Long, extra As Long, count As Long
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    factor = units.InputLengthToInternal(1#)
    keys = Array("R1.Left", "R1.Right", "R1.Top", "R1.Bottom", "R2.Left", "R2.Right", "R2.Top", "R2.Bottom")
    Set table = ThisWorkbook.Names("rngRectanglesGeometry").RefersToRange
    If scenario = 0 Then Exit Sub
    For i = 1 To 8
        TestFixtureInputCell(table, "Extra", CStr(keys(i - 1)), "a2").Value2 = 0#: TestFixtureInputCell(table, "Extra", CStr(keys(i - 1)), "a3").Value2 = 0#
        If fixture = "Rectangle" And i > 4 Then GoTo InactiveSide
        count = 5
        If scenario = 2 Then count = 3
        If scenario = 4 Then count = 1
        TestFixtureInputCell(table, "Main", CStr(keys(i - 1)), "n").Value2 = count
        TestFixtureInputCell(table, "Main", CStr(keys(i - 1)), "s0").Value2 = 70# / factor
        TestFixtureInputCell(table, "Main", CStr(keys(i - 1)), "s1").Value2 = 90# / factor
        TestFixtureInputCell(table, "Extra", CStr(keys(i - 1)), "d2").Value2 = IIf(scenario = 4, 0#, 16# / factor)
        TestFixtureInputCell(table, "Extra", CStr(keys(i - 1)), "d3").Value2 = 12# / factor
        TestFixtureInputCell(table, "Extra", CStr(keys(i - 1)), "pos2").Value2 = IIf(scenario = 2 Or (scenario = 3 And i Mod 2 = 0), "SideBySide", "Stacked")
        TestFixtureInputCell(table, "Extra", CStr(keys(i - 1)), "pos3").Value2 = IIf(scenario = 2 Or (scenario = 3 And i Mod 2 = 1), "SideBySide", "Stacked")
        TestFixtureInputCell(table, "Extra", CStr(keys(i - 1)), "bind2").Value2 = IIf(scenario = 2 Or i Mod 2 = 0, "EverySecondBar", "EachBar")
        TestFixtureInputCell(table, "Extra", CStr(keys(i - 1)), "bind3").Value2 = IIf(scenario = 4 Or i Mod 2 = 1, "EverySecondBar", "EachBar")
        If scenario = 4 Then TestFixtureInputCell(table, "Extra", CStr(keys(i - 1)), "pos2").Value2 = "inactive value"
InactiveSide:
    Next i
    If scenario = 5 Then
        TestFixtureInputCell(table, "Main", "R1.Bottom", "n").Value2 = 0
        ' Заданный as отключённой стороны сохраняется: он ограничивает общий контур осей.
        TestFixtureInputCell(table, "Main", "R1.Bottom", "s0").Value2 = "inactive value"
        TestFixtureInputCell(table, "Extra", "R1.Bottom", "d2").Value2 = "inactive value"
        If fixture = "Rectangle" Then
            TestFixtureInputCell(table, "General", "H2", "val").Value2 = 0#: TestFixtureInputCell(table, "General", "B2", "val").Value2 = 0# ' Числовые размеры исключают часть; ошибочное армирование ниже не читается.
            For i = 15 To 18: TestFixtureInputCell(table, "Main", CStr(keys(i - 11)), "d").Value2 = "inactive value": Next i
        End If
        If fixture = "Rectangle" Then TestFixtureInputCell(table, "Detail", "XOffset", "val").Value2 = "inactive value"
    End If
    ThisWorkbook.Worksheets("Config").Calculate
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Выполняет сборку сечения в отдельной эталонной книге.
Public Function TestDumpRectanglesGeometryMatrix() As String
    Dim config As Object, saved As Variant, result As String, fixture As Variant, scenario As Long
    Dim settings As CSystemSettingsReader, units As CUnitSystem, builder As CSectionBuilder, model As CSectionModel
    Dim calculation As Long, events As Boolean, updating As Boolean
    On Error GoTo Failed
    calculation = Application.Calculation: events = Application.EnableEvents: updating = Application.ScreenUpdating
    Application.EnableEvents = False: Application.ScreenUpdating = False
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = TestFixtureSaveInputs(config)
    For Each fixture In Array("Rectangle", "LSection", "TwoRectangles")
        For scenario = 0 To 5
            TestFixtureRestoreInputs config, saved
            TestConfigureShapeFixture "Circle": TestConfigureShapeFixture CStr(fixture)
            TestConfigureRectanglesScenario CStr(fixture), scenario
            Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
            Set units = New CUnitSystem: units.LoadFromSettings settings
            Set builder = New CSectionBuilder: Set model = BuildWorkbookSectionModel(ThisWorkbook, settings, units, builder)
            result = result & TestShapeSnapshot("RectMatrix." & CStr(fixture) & ".case" & CStr(scenario), model, builder.Mesh, builder.Rebars)
        Next scenario
    Next fixture
    GoTo Finished
Failed:
    result = result & "FAIL: Rectangles baseline matrix; " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Finished:
    If Not config Is Nothing And Not IsEmpty(saved) Then TestFixtureRestoreInputs config, saved
    Application.Calculation = calculation: Application.EnableEvents = events: Application.ScreenUpdating = updating
    TestDumpRectanglesGeometryMatrix = result
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' одинаковая матрица двух независимых сторон и активных дуг.
Public Sub TestConfigureRoundedScenario(ByVal scenario As Long)
    Dim keys As Variant
    Dim table As Object, data As Variant, settings As CSystemSettingsReader, units As CUnitSystem
    Dim factor As Double, mainHeader As Long, extraHeader As Long, i As Long, row As Long, extra As Long
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings: factor = units.InputLengthToInternal(1#)
    keys = Array("Left", "Right", "Top", "Bottom")
    Set table = ThisWorkbook.Names("rngPointedSectionGeometry").RefersToRange: data = table.Value2
    If scenario = 0 Then Exit Sub
    If scenario = 1 Then
        TestFixtureInputCell(table, "Detail", "R1", "left").Value2 = 80# / factor: TestFixtureInputCell(table, "Detail", "R1", "right").Value2 = 100# / factor
        TestFixtureInputCell(table, "Detail", "R2", "left").Value2 = 40# / factor: TestFixtureInputCell(table, "Detail", "R2", "right").Value2 = 50# / factor
    ElseIf scenario = 2 Then
        TestFixtureInputCell(table, "Detail", "R1", "left").Value2 = 0#: TestFixtureInputCell(table, "Detail", "R1", "right").Value2 = 0#
        TestFixtureInputCell(table, "Detail", "R2", "left").Value2 = 0#: TestFixtureInputCell(table, "Detail", "R2", "right").Value2 = 0#
    End If
    For i = 1 To 4
        TestFixtureInputCell(table, "Extra", CStr(keys(i - 1)), "a2").Value2 = 0#: TestFixtureInputCell(table, "Extra", CStr(keys(i - 1)), "a3").Value2 = 0#
        TestFixtureInputCell(table, "Main", CStr(keys(i - 1)), "as").Value2 = (35# + 5# * CDbl(i)) / factor
        TestFixtureInputCell(table, "Main", CStr(keys(i - 1)), "d").Value2 = 20# / factor
        TestFixtureInputCell(table, "Main", CStr(keys(i - 1)), "n").Value2 = IIf(i <= 2, 7, 5)
        TestFixtureInputCell(table, "Extra", CStr(keys(i - 1)), "d2").Value2 = 16# / factor: TestFixtureInputCell(table, "Extra", CStr(keys(i - 1)), "d3").Value2 = 12# / factor
        TestFixtureInputCell(table, "Extra", CStr(keys(i - 1)), "pos2").Value2 = IIf(scenario = 3 And i Mod 2 = 0, "SideBySide", "Stacked")
        TestFixtureInputCell(table, "Extra", CStr(keys(i - 1)), "pos3").Value2 = IIf(scenario = 3 And i Mod 2 = 1, "SideBySide", "Stacked")
        TestFixtureInputCell(table, "Extra", CStr(keys(i - 1)), "bind2").Value2 = IIf(i Mod 2 = 0, "EverySecondBar", "EachBar")
        TestFixtureInputCell(table, "Extra", CStr(keys(i - 1)), "bind3").Value2 = IIf(i Mod 2 = 1, "EverySecondBar", "EachBar")
        If scenario = 4 Then
            TestFixtureInputCell(table, "Main", CStr(keys(i - 1)), "n").Value2 = 1: TestFixtureInputCell(table, "Extra", CStr(keys(i - 1)), "d2").Value2 = 0#
            TestFixtureInputCell(table, "Extra", CStr(keys(i - 1)), "pos2").Value2 = "inactive value": TestFixtureInputCell(table, "Extra", CStr(keys(i - 1)), "bind2").Value2 = "inactive value"
        End If
    Next i
    If scenario = 5 Then
        TestFixtureInputCell(table, "Main", "Top", "n").Value2 = 0
        ' Заданный as отключённой стороны сохраняется: он ограничивает общий контур осей.
        TestFixtureInputCell(table, "Extra", "Top", "d2").Value2 = "inactive value"
        For i = 2 To 3
            If CDbl(TestFixtureInputCell(table, "Detail", "W", IIf(i = 2, "left", "right")).Value2) = 0# Then
                TestFixtureInputCell(table, "Detail", "R2", IIf(i = 2, "left", "right")).Value2 = "inactive value"
            End If
        Next i
    End If
    ThisWorkbook.Worksheets("Config").Calculate
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Ищет строку контрольного заголовка в указанном столбце; отсутствие заголовка не маскируется нулевым индексом.
Private Function TestFixtureHeaderRow(ByRef data As Variant, ByVal caption As String, ByVal column As Long) As Long
    Dim row As Long
    For row = LBound(data, 1) To UBound(data, 1)
        If StrComp(Trim$(CStr(data(row, column))), caption, vbTextCompare) = 0 Then TestFixtureHeaderRow = row: Exit Function
    Next row
    Err.Raise vbObjectError + 5590, "modTestShapeArchitecture", "Не найдена шапка fixture: " & caption
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' фиксирует старую расчетную ломаную до аналитического переноса.
Public Function TestDumpRoundedGeometryMatrix() As String
    Dim config As Object, saved As Variant, result As String, fixture As Variant, scenario As Long
    Dim settings As CSystemSettingsReader, units As CUnitSystem, builder As CSectionBuilder, model As CSectionModel
    Dim calculation As Long, events As Boolean, updating As Boolean
    On Error GoTo Failed
    calculation = Application.Calculation: events = Application.EnableEvents: updating = Application.ScreenUpdating
    Application.EnableEvents = False: Application.ScreenUpdating = False
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = TestFixtureSaveInputs(config)
    For Each fixture In Array("RoundedSimple", "RoundedTapered", "RoundedMixed", "RoundedMixedLeft")
        For scenario = 0 To 5
            TestFixtureRestoreInputs config, saved
            TestConfigureShapeFixture "Circle": TestConfigureShapeFixture CStr(fixture)
            TestConfigureRoundedScenario scenario
            Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
            Set units = New CUnitSystem: units.LoadFromSettings settings
            Set builder = New CSectionBuilder: Set model = BuildWorkbookSectionModel(ThisWorkbook, settings, units, builder)
            result = result & TestShapeSnapshot("RoundedMatrix." & CStr(fixture) & ".case" & CStr(scenario), model, builder.Mesh, builder.Rebars)
        Next scenario
    Next fixture
    GoTo Finished
Failed:
    result = result & "FAIL: Rounded baseline matrix; " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Finished:
    If Not config Is Nothing And Not IsEmpty(saved) Then TestFixtureRestoreInputs config, saved
    Application.Calculation = calculation: Application.EnableEvents = events: Application.ScreenUpdating = updating
    TestDumpRoundedGeometryMatrix = result
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' одинаковые параметры HollowRectangle для baseline и рецепта.
Public Sub TestConfigureHollowScenario(ByVal scenario As Long)
    Dim keys As Variant
    Dim table As Object, data As Variant, settings As CSystemSettingsReader, units As CUnitSystem
    Dim factor As Double, mainHeader As Long, extraHeader As Long, i As Long, row As Long, extra As Long
    If scenario = 0 Then Exit Sub
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings: factor = units.InputLengthToInternal(1#)
    keys = Array("Left", "Right", "Top", "Bottom", "OpeningLeft", "OpeningRight", "OpeningTop", "OpeningBottom")
    Set table = ThisWorkbook.Names("rngHollowRectangleGeometry").RefersToRange: data = table.Value2
    TestFixtureInputCell(table, "General", "R", "val").Value2 = 180# / factor: TestFixtureInputCell(table, "Detail", "Ro", "val").Value2 = 90# / factor
    If scenario = 2 Or scenario = 6 Then TestFixtureInputCell(table, "Detail", "Ro", "val").Value2 = 0#
    If scenario = 8 Or scenario = 9 Then
        TestFixtureInputCell(table, "General", "H", "val").Value2 = IIf(scenario = 8, 600#, 1000#) / factor
        TestFixtureInputCell(table, "General", "B", "val").Value2 = 1000# / factor
        TestFixtureInputCell(table, "General", "R", "val").Value2 = IIf(scenario = 8, 300#, 500#) / factor
        TestFixtureInputCell(table, "Detail", "Ho", "val").Value2 = IIf(scenario = 8, 200#, 400#) / factor
        TestFixtureInputCell(table, "Detail", "Bo", "val").Value2 = 400# / factor
        TestFixtureInputCell(table, "Detail", "Ro", "val").Value2 = IIf(scenario = 8, 100#, 200#) / factor
    End If
    If scenario = 10 Then TestFixtureInputCell(table, "General", "R", "val").Value2 = 0#: TestFixtureInputCell(table, "Detail", "Ro", "val").Value2 = 0#
    If scenario = 3 Then
        TestFixtureInputCell(table, "Detail", "Xo", "val").Value2 = 160# / factor: TestFixtureInputCell(table, "Detail", "Yo", "val").Value2 = -180# / factor
    End If
    For i = 1 To 8
        TestFixtureInputCell(table, "Extra", CStr(keys(i - 1)), "a2").Value2 = 0#: TestFixtureInputCell(table, "Extra", CStr(keys(i - 1)), "a3").Value2 = 0#
        TestFixtureInputCell(table, "Main", CStr(keys(i - 1)), "as").Value2 = (35# + 5# * CDbl(i)) / factor
        TestFixtureInputCell(table, "Main", CStr(keys(i - 1)), "d").Value2 = 20# / factor
        If i <= 4 Then TestFixtureInputCell(table, "Main", CStr(keys(i - 1)), "n").Value2 = IIf(i <= 2, 7, 9)
        TestFixtureInputCell(table, "Extra", CStr(keys(i - 1)), "d2").Value2 = 16# / factor: TestFixtureInputCell(table, "Extra", CStr(keys(i - 1)), "d3").Value2 = 12# / factor
        TestFixtureInputCell(table, "Extra", CStr(keys(i - 1)), "pos2").Value2 = IIf(scenario = 4 And i Mod 2 = 0, "SideBySide", "Stacked")
        TestFixtureInputCell(table, "Extra", CStr(keys(i - 1)), "pos3").Value2 = IIf(scenario = 4 And i Mod 2 = 1, "SideBySide", "Stacked")
        TestFixtureInputCell(table, "Extra", CStr(keys(i - 1)), "bind2").Value2 = IIf(i Mod 2 = 0, "EverySecondBar", "EachBar")
        TestFixtureInputCell(table, "Extra", CStr(keys(i - 1)), "bind3").Value2 = IIf(i Mod 2 = 1, "EverySecondBar", "EachBar")
        If scenario = 5 Then
            If i <= 4 Then TestFixtureInputCell(table, "Main", CStr(keys(i - 1)), "n").Value2 = 1
            TestFixtureInputCell(table, "Extra", CStr(keys(i - 1)), "d2").Value2 = 0#
            TestFixtureInputCell(table, "Extra", CStr(keys(i - 1)), "pos2").Value2 = "inactive value": TestFixtureInputCell(table, "Extra", CStr(keys(i - 1)), "bind2").Value2 = "inactive value"
        End If
        If scenario = 6 And i = 5 Then
            TestFixtureInputCell(table, "Main", CStr(keys(i - 1)), "d").Value2 = 0#
            TestFixtureInputCell(table, "Extra", CStr(keys(i - 1)), "d2").Value2 = "inactive value": TestFixtureInputCell(table, "Extra", CStr(keys(i - 1)), "d3").Value2 = "inactive value"
        End If
        If scenario = 7 And i = 3 Then
            TestFixtureInputCell(table, "Main", CStr(keys(i - 1)), "n").Value2 = 0
            ' Заданный as отключённой стороны сохраняется: он ограничивает общий контур осей.
            TestFixtureInputCell(table, "Extra", CStr(keys(i - 1)), "d2").Value2 = "inactive value": TestFixtureInputCell(table, "Extra", CStr(keys(i - 1)), "d3").Value2 = "inactive value"
        End If
        If (scenario = 8 Or scenario = 9) And (i = 1 Or i = 2 Or i = 5 Or i = 6) Then
            TestFixtureInputCell(table, "Main", CStr(keys(i - 1)), "d").Value2 = 0#: If i <= 4 Then TestFixtureInputCell(table, "Main", CStr(keys(i - 1)), "n").Value2 = 0
            TestFixtureInputCell(table, "Extra", CStr(keys(i - 1)), "d2").Value2 = "inactive value": TestFixtureInputCell(table, "Extra", CStr(keys(i - 1)), "d3").Value2 = "inactive value"
        End If
    Next i
    ThisWorkbook.Worksheets("Config").Calculate
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Выводит контрольную геометрию и проекции HollowRectangle для сравнения реализаций.
Public Function TestDumpHollowGeometryMatrix() As String
    TestDumpHollowGeometryMatrix = TestHollowGeometryMatrix(0, 7)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' исчезнувшие прямые при предельных радиусах и оба нулевых радиуса.
Public Function TestDumpHollowEdgeGeometryMatrix() As String
    TestDumpHollowEdgeGeometryMatrix = TestHollowGeometryMatrix(8, 10)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сохраняет точные снимки заданного диапазона пустотелых сечений с центральным и смещённым отверстием.
Private Function TestHollowGeometryMatrix(ByVal firstScenario As Long, ByVal lastScenario As Long) As String
    Dim config As Object, saved As Variant, result As String, fixture As Variant, scenario As Long
    Dim settings As CSystemSettingsReader, units As CUnitSystem, builder As CSectionBuilder, model As CSectionModel
    Dim calculation As Long, events As Boolean, updating As Boolean
    On Error GoTo Failed
    calculation = Application.Calculation: events = Application.EnableEvents: updating = Application.ScreenUpdating
    Application.EnableEvents = False: Application.ScreenUpdating = False
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = TestFixtureSaveInputs(config)
    For Each fixture In Array("HollowCentered", "HollowOffset")
        For scenario = firstScenario To lastScenario
            TestFixtureRestoreInputs config, saved
            TestConfigureShapeFixture "Circle": TestConfigureShapeFixture CStr(fixture)
            TestConfigureHollowScenario scenario
            Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
            Set units = New CUnitSystem: units.LoadFromSettings settings
            Set builder = New CSectionBuilder: Set model = BuildWorkbookSectionModel(ThisWorkbook, settings, units, builder)
            result = result & TestShapeSnapshot("HollowMatrix." & CStr(fixture) & ".case" & CStr(scenario), model, builder.Mesh, builder.Rebars)
        Next scenario
    Next fixture
    GoTo Finished
Failed:
    result = result & "FAIL: Hollow baseline matrix; " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Finished:
    If Not config Is Nothing And Not IsEmpty(saved) Then TestFixtureRestoreInputs config, saved
    Application.Calculation = calculation: Application.EnableEvents = events: Application.ScreenUpdating = updating
    TestHollowGeometryMatrix = result
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' сравнивает время одинаковых вызовов AddBar; обращения к Excel и решение НДС в
' замер не входят.
Public Function TestDumpRebarLayoutPerformance() As String
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
    TestDumpRebarLayoutPerformance = result
    Exit Function
Failed:
    TestDumpRebarLayoutPerformance = result & "FAIL: layout performance; " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' сравнивает расчёт до и после рефакторинга без изменения Results.
' Для каждой формы использует три назначения материалов и четыре направления
' нагрузки, затем — несущую способность, образование и ширину трещин СП 35.
Public Function TestDumpShapeCalculationSnapshot(Optional ByVal crackCode As String = "SP35") As String
    TestDumpShapeCalculationSnapshot = TestSourceCalculations(Array("Circle", "Rectangle", "LSection", "TwoRectangles", _
        "RoundedSimple", "RoundedTapered", "RoundedMixed", "HollowCentered", "HollowOffset"), crackCode)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' исходный расчет четырех PointedSection, включая точные
' контуры и контрольные позиции арматуры. Вторая проба активирует радиусные пути.
Public Function TestDumpRoundedSourceCalculations(Optional ByVal crackCode As String = "SP35", Optional ByVal scenario As Long = 0) As String
    TestDumpRoundedSourceCalculations = TestSourceCalculations(Array("RoundedSimple", "RoundedTapered", "RoundedMixed", "RoundedMixedLeft"), crackCode, scenario)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' отдельная эталонная книга с контрольной рабочей сборкой пустотелого прямоугольника.
Public Function TestDumpHollowSourceCalculations(Optional ByVal crackCode As String = "SP35", Optional ByVal scenario As Long = 0) As String
    TestDumpHollowSourceCalculations = TestSourceCalculations(Array("HollowCentered", "HollowOffset"), crackCode, 0, scenario)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сохраняет результаты прямых расчётов и batch для выбранных вариантов и методики трещин, восстанавливая входные таблицы.
Private Function TestSourceCalculations(ByVal fixtures As Variant, ByVal crackCode As String, Optional ByVal roundedScenario As Long = 0, _
        Optional ByVal hollowScenario As Long = 0) As String
    Dim config As Object, saved As Variant, result As String, fixture As Variant
    Dim calculation As Long, events As Boolean, updating As Boolean
    On Error GoTo Failed
    calculation = Application.Calculation: events = Application.EnableEvents: updating = Application.ScreenUpdating
    Application.EnableEvents = False: Application.ScreenUpdating = False
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = TestFixtureSaveInputs(config)
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
    TestSetSetting "SLS.Crack.Code", crackCode
    TestConfigureShapeFixture "Circle"
    For Each fixture In fixtures
        TestConfigureShapeFixture CStr(fixture)
        If roundedScenario > 0 Then TestConfigureRoundedScenario roundedScenario
        If hollowScenario > 0 Then TestConfigureHollowScenario hollowScenario
        Dim settings As CSystemSettingsReader, units As CUnitSystem, builder As CSectionBuilder, model As CSectionModel, label As String
        Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
        Set units = New CUnitSystem: units.LoadFromSettings settings
        Set builder = New CSectionBuilder: Set model = BuildWorkbookSectionModel(ThisWorkbook, settings, units, builder)
        label = CStr(fixture): If roundedScenario > 0 Then label = label & ".Arc"
        If hollowScenario > 0 Then label = label & ".case" & CStr(hollowScenario)
        result = result & TestCalculationSnapshot(label, model)
    Next fixture
    GoTo Finished
Failed:
    result = result & "FAIL: calculation snapshot; " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Finished:
    If Not config Is Nothing And Not IsEmpty(saved) Then TestFixtureRestoreInputs config, saved
    Application.Calculation = calculation: Application.EnableEvents = events: Application.ScreenUpdating = updating
    TestSourceCalculations = result
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сохраняет НДС трёх режимов материалов и четырёх нагрузок, поэлементные состояния и два контрольных сочетания batch.
Private Function TestCalculationSnapshot(ByVal fixture As String, Optional ByVal suppliedModel As CSectionModel = Nothing) As String
    Dim model As CSectionModel, props As CSectionPropertiesCalculator
    Dim settings As CSystemSettingsReader, units As CUnitSystem, provider As CMaterialModelProvider
    Set model = suppliedModel
    If model Is Nothing Then Set model = TestFixtureModel(fixture)
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
            result = result & TestRow(fixture, "NDS", CStr(purpose) & "." & CStr(loadIndex), Array( _
                solver.Converged, solver.Epsilon0, solver.KappaX, solver.KappaY, solver.Nint, solver.Mxint, solver.Myint, _
                solver.MinConcreteStress, solver.MaxConcreteStress, solver.MinSteelStress, solver.MaxSteelStress, _
                solver.ResidualN, solver.ResidualMx, solver.ResidualMy, solver.MinConcreteStrain, solver.MaxConcreteStrain, _
                solver.MinSteelStrain, solver.MaxSteelStrain))
            result = result & TestRow(fixture, "LOAD", CStr(purpose) & "." & CStr(loadIndex), _
                Array(n, mx, my, props.CentroidX, props.CentroidY))
            If solver.Converged Then result = result & TestElementStateSnapshot(fixture, CStr(purpose) & "." & CStr(loadIndex), _
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
        result = result & TestRow(fixture, "BATCH", CStr(loadIndex), Array(combination.Status, combination.NormalCrackStatus, _
            combination.StrengthResult.Capacity.ResultMeta.InternalStatus, combination.CrackResult.Formation.ResultMeta.InternalStatus, _
            combination.CrackResult.Width.ResultMeta.Calculated, combination.StrengthResult.Capacity.LambdaCapacity, _
            combination.CrackResult.Formation.LambdaCrc, combination.CrackResult.Width.CrackWidth, combination.CrackResult.Width.Abt))
    Next loadIndex
    TestCalculationSnapshot = result
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' расчеты только переданной новой модели.
Public Function TestPreparedCalculationSnapshot(ByVal fixture As String, ByVal model As CSectionModel) As String
    If model Is Nothing Then Err.Raise vbObjectError + 5292, "modTestShapeArchitecture", "Расчетная модель не передана."
    TestPreparedCalculationSnapshot = TestCalculationSnapshot(fixture, model)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' вычисляет значения по фактическому полю деформаций и выбранным
' материалам каждого элемента. StrainAt использует численные координаты решателя,
' в том числе для моделей, удалённых от начала координат.
Private Function TestElementStateSnapshot(ByVal fixture As String, ByVal stateID As String, ByVal model As CSectionModel, _
        ByVal solver As CSectionSolver, ByVal concrete As Object, ByVal steel As Object) As String
    Dim lines() As String, i As Long, strain As Double
    ReDim lines(1 To model.ConcreteCount + model.RebarCount)
    For i = 1 To model.ConcreteCount
        strain = solver.StrainAt(model.ConcreteX(i), model.ConcreteY(i))
        lines(i) = TestRow(fixture, "FIBER_STATE", stateID & "." & model.ConcreteID(i), Array(strain, concrete.StressAtStrain(strain)))
    Next i
    For i = 1 To model.RebarCount
        strain = solver.StrainAt(model.RebarX(i), model.RebarY(i))
        lines(model.ConcreteCount + i) = TestRow(fixture, "BAR_STATE", stateID & "." & model.RebarID(i), Array(strain, steel.StressAtStrain(strain)))
    Next i
    TestElementStateSnapshot = Join(lines, vbNullString)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' строит настоящую диаграмму, фигуры и PNG геометрических эталонов.
' Results меняет только в отдельной тестовой копии, закрываемой без сохранения.
' Контрольный вид пустотелого прямоугольника нужен для сравнения длинных подписей.
Public Function TestDumpHollowSourcePreview(ByVal outputDirectory As String, Optional ByVal scenario As Long = 0) As String
    Dim config As Object, saved As Variant, result As String, fixture As Variant
    Dim settings As CSystemSettingsReader, units As CUnitSystem, builder As CSectionBuilder, model As CSectionModel
    Dim calculation As Long, events As Boolean, updating As Boolean
    On Error GoTo Failed
    calculation = Application.Calculation: events = Application.EnableEvents: updating = Application.ScreenUpdating
    Application.EnableEvents = False: Application.ScreenUpdating = True
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = TestFixtureSaveInputs(config)
    TestConfigureShapeFixture "Circle"
    For Each fixture In Array("HollowCentered", "HollowOffset")
        TestConfigureShapeFixture CStr(fixture): TestConfigureHollowScenario scenario
        Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
        Set units = New CUnitSystem: units.LoadFromSettings settings
        Set builder = New CSectionBuilder: Set model = BuildWorkbookSectionModel(ThisWorkbook, settings, units, builder)
        result = result & TestPreviewSnapshot(CStr(fixture) & ".case" & CStr(scenario), outputDirectory, model)
    Next fixture
    GoTo Finished
Failed:
    result = result & "FAIL: source Hollow preview; " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Finished:
    If Not config Is Nothing And Not IsEmpty(saved) Then TestFixtureRestoreInputs config, saved
    Application.Calculation = calculation: Application.EnableEvents = events: Application.ScreenUpdating = updating
    TestDumpHollowSourcePreview = result
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сохраняет схемы и их фактические координаты для выбранного контрольного сечения или всех вариантов.
Public Function TestDumpShapePreviewSnapshot(ByVal outputDirectory As String, Optional ByVal fixtureFilter As String = "") As String
    Dim config As Object, saved As Variant, calculation As Long, events As Boolean, updating As Boolean
    Dim result As String, fixture As Variant
    On Error GoTo Failed
    calculation = Application.Calculation: events = Application.EnableEvents: updating = Application.ScreenUpdating
    Application.EnableEvents = False: Application.ScreenUpdating = True
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = TestFixtureSaveInputs(config)
    For Each fixture In Array("Circle", "Rectangle", "LSection", "TwoRectangles", _
            "RoundedSimple", "RoundedTapered", "RoundedMixed", "HollowCentered", "HollowOffset")
        If Len(fixtureFilter) = 0 Or CStr(fixture) = fixtureFilter Then result = result & TestPreviewSnapshot(CStr(fixture), outputDirectory)
    Next fixture
    GoTo Finished
Failed:
    result = result & "FAIL: preview snapshot; " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Finished:
    If Not config Is Nothing And Not IsEmpty(saved) Then TestFixtureRestoreInputs config, saved
    Application.Calculation = calculation: Application.EnableEvents = events: Application.ScreenUpdating = updating
    TestDumpShapePreviewSnapshot = result
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Записывает геометрический снимок, повторно читает Results и сохраняет реальные точки и фигуры схемы вместе с PNG.
Private Function TestPreviewSnapshot(ByVal fixture As String, ByVal outputDirectory As String, Optional ByVal suppliedModel As CSectionModel = Nothing) As String
    Dim stage As String
    On Error GoTo Failed
    stage = "build"
    Dim model As CSectionModel, properties As CSectionPropertiesCalculator
    Dim settings As CSystemSettingsReader, units As CUnitSystem, writer As CNDMResultsWriter
    Dim reader As CSectionPlotDataReader, plotter As CSectionPlotter
    Set model = suppliedModel
    If model Is Nothing Then Set model = TestFixtureModel(fixture)
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
    ' Чтение только геометрии предполагает источник AutoCAD; эти тестовые модели получены
    ' генерацией.
    chart.Shapes("NDMPlot_Title").TextFrame2.TextRange.Text = "Сгенерированная геометрия: " & fixture
    result = TestRow(fixture, "PLOT", "container", Array(container.Left, container.Top, container.Width, container.Height, _
        chart.PlotArea.Left, chart.PlotArea.Top, chart.PlotArea.Width, chart.PlotArea.Height, _
        chart.SeriesCollection.Count, chart.Shapes.Count))
    For i = 1 To chart.SeriesCollection.Count
        stage = "series " & CStr(i)
        Set series = chart.SeriesCollection(i)
        xs = series.XValues: ys = series.Values
        For j = LBound(xs) To UBound(xs)
            result = result & TestRow(fixture, "PLOT_POINT", CStr(i) & "." & CStr(j), Array(CStr(series.Name), CDbl(xs(j)), CDbl(ys(j))))
        Next j
    Next i
    For Each shape In chart.Shapes
        stage = "shape " & CStr(shape.Name)
        shapeText = vbNullString
        If CLng(shape.Type) = 17 Or CLng(shape.Type) = 1 Then
            If shape.TextFrame2.HasText Then shapeText = CStr(shape.TextFrame2.TextRange.Text)
        End If
        result = result & TestRow(fixture, "PLOT_SHAPE", CStr(shape.Name), Array(CLng(shape.Type), _
            CDbl(shape.Left), CDbl(shape.Top), CDbl(shape.Width), CDbl(shape.Height), CDbl(shape.Rotation), shapeText))
    Next shape
    stage = "export PNG"
    chart.Refresh
    If Not chart.Export(outputDirectory & Application.PathSeparator & fixture & ".png", "PNG") Then _
        Err.Raise vbObjectError + 5294, "modTestShapeArchitecture", "Excel не сохранил PNG: " & fixture
    TestPreviewSnapshot = result
    Exit Function
Failed:
    Err.Raise Err.Number, "modTestShapeArchitecture.TestPreviewSnapshot", fixture & "; " & stage & "; " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' записывает переданную новую модель в Results, читает её и строит обычную схему.
Public Function TestPreparedPreviewSnapshot(ByVal fixture As String, ByVal model As CSectionModel, ByVal outputDirectory As String) As String
    If model Is Nothing Then Err.Raise vbObjectError + 5292, "modTestShapeArchitecture", "Модель схемы не передана."
    TestPreparedPreviewSnapshot = TestPreviewSnapshot(fixture, outputDirectory, model)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Создаёт строку снимка с экранированными полями и типизированным представлением каждого значения.
Private Function TestRow(ByVal fixture As String, ByVal kind As String, ByVal id As String, ByVal values As Variant) As String
    Dim item As Variant
    TestRow = TestEscapeField(fixture) & vbTab & TestEscapeField(kind) & vbTab & TestEscapeField(id)
    For Each item In values: TestRow = TestRow & vbTab & TestSnapshotValue(item): Next item
    TestRow = TestRow & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Кодирует тип значения явно; числа с плавающей точкой сохраняет побитово, чтобы формат Windows не скрывал расхождения.
Private Function TestSnapshotValue(ByVal value As Variant) As String
    Dim number As TSnapshotDouble, words As TSnapshotWords
    Select Case VarType(value)
        Case vbDouble, vbSingle, vbCurrency, vbDecimal
            number.value = CDbl(value): LSet words = number ' Копируются байты Double: снимок обнаруживает даже изменения младших разрядов.
            TestSnapshotValue = "d:" & Right$("00000000" & Hex$(words.high), 8) & Right$("00000000" & Hex$(words.low), 8)
        Case vbByte, vbInteger, vbLong
            TestSnapshotValue = "i:" & CStr(value)
        Case vbBoolean
            If value Then TestSnapshotValue = "b:1" Else TestSnapshotValue = "b:0"
        Case vbEmpty
            TestSnapshotValue = "e:"
        Case vbNull
            TestSnapshotValue = "n:"
        Case vbString
            TestSnapshotValue = "s:" & TestEscapeField(CStr(value))
        Case Else
            Err.Raise vbObjectError + 5293, "modTestShapeArchitecture", "Неподдерживаемое значение snapshot: " & CStr(VarType(value))
    End Select
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Экранирует проценты, табуляцию и переводы строк, сохраняя однозначные границы полей снимка.
Private Function TestEscapeField(ByVal value As String) As String
    TestEscapeField = Replace(Replace(Replace(Replace(value, "%", "%25"), vbTab, "%09"), vbCr, "%0D"), vbLf, "%0A")
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Меняет системное или геометрическое поле контрольного варианта по ключу; отсутствие поля прерывает подготовку проверки.
Private Sub TestSetSetting(ByVal key As String, ByVal value As Variant)
    Dim name As Variant, table As Object, i As Long
    If Left$(key, 6) = "Rebar." Or Left$(key, 7) = "Circle." Then
        Dim circleCell As Object
        Set circleCell = TestFixtureCircleFieldCell(ThisWorkbook, key): circleCell.Value2 = value: Exit Sub
    End If
    For Each name In Array("rngSystemSettings", "rngCircleGeometry")
        Set table = ThisWorkbook.Names(CStr(name)).RefersToRange
        For i = 2 To table.Rows.Count
            If CStr(table.Cells(i, 1).Value2) = key Then table.Cells(i, 2).Value2 = value: Exit Sub
        Next i
    Next name
    Err.Raise vbObjectError + 5291, "modTestShapeArchitecture", "Нет ключа fixture: " & key
End Sub
