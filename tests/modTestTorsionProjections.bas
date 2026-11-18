Attribute VB_Name = "modTestTorsionProjections"
Option Explicit

' ==================== ДЛЯ ТЕСТОВ ====================
' Независимый контроль равных точечных и непрерывных проекций, площади,
' равновесия, симметрии и независимости от направления и разбиения контура.
' ===================================================
Private mTestPassed As Long, mTestFailed As Long, mTestLog As String

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет подготовленные коэффициенты по аналитическим долям тяги.
Public Function RunTorsionProjectionTests() As String
    mTestPassed = 0: mTestFailed = 0: mTestLog = vbNullString
    On Error GoTo Failed
    TestProjectionComponents
    Dim section As CSectionModel, hoop As CTorsionHoopGeometry, demand As CTorsionLongitudinalDemand
    Dim coefficients() As Double, baseline() As Double, i As Long, caseIndex As Long, sign As Long
    For caseIndex = 0 To 7
        Set section = TestProjectionRectangle(caseIndex, False, 1)
        Set hoop = New CTorsionHoopGeometry: hoop.InitializeAutomatic section, 20#, 8#
        Set demand = New CTorsionLongitudinalDemand: demand.Initialize section, hoop, 1000#
        demand.CopyStressPerUnitForce coefficients
        For i = 1 To 4: TestProjectionNear "Rectangle corner " & caseIndex & ":" & i, coefficients(i), 29# / 16800#: Next i
        TestProjectionNear "Rectangle centre " & caseIndex, coefficients(5), 13# / 4200#
        TestProjectionEquilibrium section, coefficients, "Rectangle " & caseIndex
        If caseIndex = 0 Then baseline = coefficients
        For i = 1 To 5: TestProjectionNear "Contour order and segmentation " & caseIndex & ":" & i, coefficients(i), baseline(i): Next i
    Next caseIndex
    Set section = TestProjectionRectangle(0, False, 1)
    Set hoop = New CTorsionHoopGeometry: hoop.InitializeAutomatic section, 20#, 8#
    Set demand = New CTorsionLongitudinalDemand: demand.Initialize section, hoop, 25#
    demand.CopyStressPerUnitForce coefficients
    TestProjectionNear "Ambiguous bar outside selected band ignored", coefficients(5), 0#
    For i = 1 To 4: TestProjectionNear "Normal corner unchanged " & i, coefficients(i), 1# / 400#: Next i

    Set section = TestProjectionRectangle(0, True, 2)
    Set hoop = New CTorsionHoopGeometry: hoop.InitializeAutomatic section, 20#, 8#
    Set demand = New CTorsionLongitudinalDemand: demand.Initialize section, hoop, 1000#
    demand.CopyStressPerUnitForce coefficients
    For i = 1 To 4: TestProjectionNear "Square outer eighth " & i, coefficients(i), 1# / 800#: Next i
    TestProjectionNear "Shared centre station area ratio A", coefficients(5), 1# / 800#
    TestProjectionNear "Shared centre station area ratio B", coefficients(6), 1# / 800#
    TestProjectionEquilibrium section, coefficients, "Square shared projections"
    TestProjectionCalculatorOutput section

    Set section = New CSectionModel
    section.AddConcreteElement 0#, 0#, GEOM_PI * 10000#, 1, shapeType:="Circle", width:=200#, height:=200#
    section.Contours.AddContourCircle "Circle", 0#, 0#, 100#
    For i = 1 To 3: TestProjectionAddBar section, 0#, 0#, 100# * i: Next i
    Set hoop = New CTorsionHoopGeometry
    hoop.InitializeExplicit section, TestProjectionCirclePath(4, 1), 8#, 16#
    Set demand = New CTorsionLongitudinalDemand: demand.Initialize section, hoop, 1000#
    demand.CopyStressPerUnitForce coefficients
    For i = 1 To 3: TestProjectionNear "Only continuous projections " & i, coefficients(i), 1# / 600#: Next i
    TestProjectionEquilibrium section, coefficients, "Only continuous projections"
    TestProjectionCalculatorOutput section

    For caseIndex = 0 To 3
        sign = 1: If caseIndex Mod 2 = 1 Then sign = -1
        Set section = TestProjectionCircle()
        Set hoop = New CTorsionHoopGeometry
        hoop.InitializeExplicit section, TestProjectionCirclePath(4 + 4 * (caseIndex \ 2), sign), 8#, 16#
        Set demand = New CTorsionLongitudinalDemand: demand.Initialize section, hoop, 1000#
        demand.CopyStressPerUnitForce coefficients
        For i = 1 To 5: TestProjectionNear "Circle density and arc segmentation " & caseIndex & ":" & i, coefficients(i), 1# / 500#: Next i
        TestProjectionEquilibrium section, coefficients, "Circle " & caseIndex
    Next caseIndex
    For caseIndex = 0 To 3
        sign = 1: If caseIndex Mod 2 = 1 Then sign = -1
        Set section = TestProjectionStadiumSection()
        Set hoop = New CTorsionHoopGeometry
        hoop.InitializeExplicit section, TestProjectionStadiumPath(50#, sign, 1 + caseIndex \ 2), 8#, 16#
        Set demand = New CTorsionLongitudinalDemand: demand.Initialize section, hoop, 1000#
        demand.CopyStressPerUnitForce coefficients
        Dim expected As Double, perimeter As Double, arcLength As Double
        arcLength = 50# * GEOM_PI: perimeter = 200# + 2# * arcLength
        expected = 1# / (400# + 100# * perimeter / arcLength)
        TestProjectionNear "Stadium continuous centre " & caseIndex, coefficients(5), expected
        TestProjectionNear "Stadium left point " & caseIndex, coefficients(1), expected
        TestProjectionNear "Stadium right point " & caseIndex, coefficients(2), 1# / 400#
        TestProjectionNear "Stadium top " & caseIndex, coefficients(3), (0.75 - 200# * expected) / 200#
        TestProjectionNear "Stadium bottom " & caseIndex, coefficients(4), coefficients(3)
        TestProjectionEquilibrium section, coefficients, "Stadium " & caseIndex
    Next caseIndex
    GoTo Finished
Failed:
    mTestFailed = mTestFailed + 1: mTestLog = mTestLog & vbCrLf & "RUNTIME: " & Err.Description
Finished:
    RunTorsionProjectionTests = "TorsionProjections: Passed=" & mTestPassed & "; Failed=" & mTestFailed & mTestLog
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сверяет каждый участок Г-контура с ручными длинами: ось 2440 мм,
' неравномерные шаги, внешний и входящий углы, в том числе без стержня в углу.
Public Function RunTorsionCornerLengthTests() As String
    mTestPassed = 0: mTestFailed = 0: mTestLog = vbNullString
    On Error GoTo Failed
    Dim section As CSectionModel, hoop As CTorsionHoopGeometry, caseIndex As Long, expected As Variant
    expected = Array(230#, 280#, 270#, 220#, 200#, 170#, 250#, 220#, 270#, 330#)
    For caseIndex = 0 To 23
        Set section = TestCornerLSection(caseIndex, 0)
        Set hoop = New CTorsionHoopGeometry
        hoop.InitializeExplicit section, TestCornerLPath(True, caseIndex), 8#, 16#
        TestCornerAllocatedLengths section, hoop, expected, "L contour " & caseIndex
    Next caseIndex
    Set section = TestCornerLSection(0, 3)
    Set hoop = New CTorsionHoopGeometry
    hoop.InitializeExplicit section, TestCornerLPath(True, 0), 8#, 16#
    TestCornerAllocatedLengths section, hoop, Array(230#, 360#, 410#, 200#, 170#, 250#, 220#, 270#, 330#), "No convex corner bar"
    Set section = TestCornerLSection(17, 6)
    Set hoop = New CTorsionHoopGeometry
    hoop.InitializeExplicit section, TestCornerLPath(True, 17), 8#, 16#
    TestCornerAllocatedLengths section, hoop, Array(230#, 280#, 270#, 220#, 310#, 310#, 220#, 270#, 330#), "No reentrant corner bar"
    Set section = TestCornerLSection(0, 0)
    Set hoop = New CTorsionHoopGeometry: hoop.InitializeAutomatic section, 20#, 8#
    TestCornerAllocatedLengths section, hoop, expected, "Automatic L offset"
    TestProjectionCalculatorOutput section
    GoTo Finished
Failed:
    mTestFailed = mTestFailed + 1: mTestLog = mTestLog & vbCrLf & "RUNTIME: " & Err.Description
Finished:
    RunTorsionCornerLengthTests = "TorsionCornerLengths: Passed=" & mTestPassed & "; Failed=" & mTestFailed & mTestLog
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Получает фактически назначенную длину из доли усилия, не повторяя поиск
' проекций и соседей. Эталон задан вручную по десяти расстояниям вдоль оси.
Private Sub TestCornerAllocatedLengths(ByVal section As CSectionModel, ByVal hoop As CTorsionHoopGeometry, _
        ByVal expected As Variant, ByVal label As String)
    Dim demand As New CTorsionLongitudinalDemand, coefficients() As Double, i As Long, length As Double, total As Double
    TestProjectionNear label & " full perimeter", hoop.Length, 2440#
    demand.Initialize section, hoop, 30#: demand.CopyStressPerUnitForce coefficients
    For i = 1 To section.RebarCount
        length = coefficients(i) * section.RebarArea(i) * hoop.Length
        TestProjectionNear label & " tributary length " & i, length, CDbl(expected(i - 1))
        total = total + length
    Next i
    TestProjectionNear label & " no lost or duplicated perimeter", total, 2440#
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Задаёт Г-бетон 600 на 700 мм с толщиной ветвей 200 мм. Физические ID
' стержней следуют заданному списку, независимо от начала и обхода контура.
Private Function TestCornerLSection(ByVal caseIndex As Long, ByVal omitted As Long) As CSectionModel
    Dim section As New CSectionModel, path As CGeometryPath, edge As TRegionEdge, i As Long
    section.AddConcreteElement 300#, 100#, 120000#, 1, shapeType:="Rectangle", width:=600#, height:=200#
    section.AddConcreteElement 100#, 450#, 100000#, 1, shapeType:="Rectangle", width:=200#, height:=500#
    Set path = TestCornerLPath(False, caseIndex)
    For i = 1 To path.SegmentCount
        path.CopySegment i, edge
        section.Contours.AddContourLine "L" & i, edge.X1, edge.Y1, edge.X2, edge.Y2
    Next i
    Dim xs As Variant, ys As Variant
    xs = Array(35#, 200#, 565#, 565#, 300#, 165#, 165#, 165#, 35#, 35#)
    ys = Array(35#, 35#, 35#, 165#, 165#, 165#, 400#, 665#, 665#, 300#)
    For i = 0 To 9
        If i + 1 <> omitted Then TestProjectionAddBar section, CDbl(xs(i)), CDbl(ys(i)), 100#
    Next i
    Set TestCornerLSection = section
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Меняет начало, направление и разбиение рёбер, сохраняя одну геометрию.
' Вершины оси (20,20)–(580,20)–(580,180)–(180,180)–(180,680)–(20,680).
Private Function TestCornerLPath(ByVal axis As Boolean, ByVal caseIndex As Long) As CGeometryPath
    Dim xs As Variant, ys As Variant, path As New CGeometryPath
    xs = Array(0#, 600#, 600#, 200#, 200#, 0#): ys = Array(0#, 0#, 200#, 200#, 700#, 700#)
    If axis Then
        xs = Array(20#, 580#, 580#, 180#, 180#, 20#): ys = Array(20#, 20#, 180#, 180#, 680#, 680#)
    End If
    path.Initialize "TestL"
    Dim direction As Long, first As Long, nextPoint As Long, i As Long, start As Long
    direction = 1: If caseIndex Mod 2 = 1 Then direction = -1
    start = (caseIndex \ 2) Mod 6
    For i = 0 To 5
        first = ((start + direction * i) Mod 6 + 6) Mod 6
        nextPoint = ((first + direction) Mod 6 + 6) Mod 6
        If caseIndex >= 12 Then
            path.AddLine "A" & i, xs(first), ys(first), 0.5 * (xs(first) + xs(nextPoint)), 0.5 * (ys(first) + ys(nextPoint))
            path.AddLine "B" & i, 0.5 * (xs(first) + xs(nextPoint)), 0.5 * (ys(first) + ys(nextPoint)), xs(nextPoint), ys(nextPoint)
        Else
            path.AddLine "E" & i, xs(first), ys(first), xs(nextPoint), ys(nextPoint)
        End If
    Next i
    path.Seal True: Set TestCornerLPath = path
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сверяет распределение площади между точкой и связной дугой, включая шов.
Private Sub TestProjectionComponents()
    Dim points As New CTorsionProjectionStations
    points.Initialize 100#: points.AddArc 0#, 25#: points.AddArc 25#, 50#
    points.AddPoint 50#: points.AddPoint 75#: points.AddPoint 75#: points.Seal
    TestProjectionAssert "Arc segments merged", points.ArcCount = 1
    TestProjectionAssert "Duplicate endpoint area excluded", points.PointCount = 1
    TestProjectionNear "Isolated point physical area share", points.PointWeight, 0.5
    TestProjectionNear "Continuous area density", points.WeightDensityAt(10#), 0.01
    TestProjectionNear "Area accounted once", points.PointWeight + 50# * points.WeightDensityAt(10#), 1#
    TestProjectionNear "No density outside nearest arc", points.WeightDensityAt(90#), 0#
    Set points = New CTorsionProjectionStations
    points.Initialize 100#: points.AddArc 0#, 10#: points.AddArc 90#, 100#
    points.AddPoint 0#: points.AddPoint 50#: points.Seal
    TestProjectionAssert "Periodic arc components merged", points.ArcCount = 1 And points.PointCount = 1
    TestProjectionNear "Seam arc first side", points.WeightDensityAt(5#), 0.025
    TestProjectionNear "Seam arc second side", points.WeightDensityAt(95#), 0.025
    TestProjectionNear "Seam area accounted once", points.PointWeight + 20# * points.WeightDensityAt(95#), 1#
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Формирует прямоугольник с фиксированными ID; меняет только обход и деление рёбер.
Private Function TestProjectionRectangle(ByVal caseIndex As Long, ByVal square As Boolean, ByVal centres As Long) As CSectionModel
    Dim section As New CSectionModel, halfY As Double, barY As Double, x As Long, y As Long, i As Long
    halfY = 150#: barY = 115#: If square Then halfY = 100#: barY = 65#
    section.AddConcreteElement 0#, 0#, 400# * halfY, 1, shapeType:="Rectangle", width:=200#, height:=2# * halfY
    For x = -1 To 1 Step 2
        For y = -1 To 1 Step 2
            TestProjectionAddBar section, 65# * x, barY * y, 100#
        Next y
    Next x
    TestProjectionAddBar section, 0#, 0#, 100#
    If centres = 2 Then TestProjectionAddBar section, 0#, 0#, 300#
    Dim xs As Variant, ys As Variant, current As Long, following As Long, direction As Long
    xs = Array(-100#, 100#, 100#, -100#): ys = Array(-halfY, -halfY, halfY, halfY)
    direction = 1: If caseIndex Mod 2 = 1 Then direction = -1
    For i = 0 To 3
        current = ((caseIndex \ 2 + direction * i) Mod 4 + 4) Mod 4
        following = ((current + direction) Mod 4 + 4) Mod 4
        If caseIndex >= 4 Then
            section.Contours.AddContourLine "A" & i, xs(current), ys(current), 0.5 * (xs(current) + xs(following)), 0.5 * (ys(current) + ys(following))
            section.Contours.AddContourLine "B" & i, 0.5 * (xs(current) + xs(following)), 0.5 * (ys(current) + ys(following)), xs(following), ys(following)
        Else
            section.Contours.AddContourLine "E" & i, xs(current), ys(current), xs(following), ys(following)
        End If
    Next i
    Set TestProjectionRectangle = section
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Задаёт круг с четырьмя наружными стержнями и одним в центре.
Private Function TestProjectionCircle() As CSectionModel
    Dim section As New CSectionModel
    section.AddConcreteElement 0#, 0#, GEOM_PI * 10000#, 1, shapeType:="Circle", width:=200#, height:=200#
    section.Contours.AddContourCircle "Circle", 0#, 0#, 100#
    TestProjectionAddBar section, -65#, 0#, 100#: TestProjectionAddBar section, 65#, 0#, 100#
    TestProjectionAddBar section, 0#, 65#, 100#: TestProjectionAddBar section, 0#, -65#, 100#
    TestProjectionAddBar section, 0#, 0#, 100#
    Set TestProjectionCircle = section
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Строит одну окружность разным числом дуг и в обоих направлениях.
Private Function TestProjectionCirclePath(ByVal segments As Long, ByVal direction As Long) As CGeometryPath
    Dim path As New CGeometryPath, i As Long, sweep As Double
    path.Initialize "Circle": sweep = direction * 2# * GEOM_PI / segments
    For i = 0 To segments - 1: path.AddArc "Arc" & i, 0#, 0#, 80#, i * sweep, sweep: Next i
    path.Seal True: Set TestProjectionCirclePath = path
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Даёт стадион с непрерывной ближайшей полуокружностью только у одного стержня.
Private Function TestProjectionStadiumSection() As CSectionModel
    Dim section As New CSectionModel, path As CGeometryPath, edge As TRegionEdge, i As Long
    section.AddConcreteElement 0#, 0#, 33600#, 1, shapeType:="Rectangle", width:=240#, height:=140#
    Set path = TestProjectionStadiumPath(70#, 1, 1)
    For i = 1 To path.SegmentCount
        path.CopySegment i, edge
        If Abs(edge.Sweep) <= GEOM_PARAMETER_TOLERANCE Then
            section.Contours.AddContourLine "L" & i, edge.X1, edge.Y1, edge.X2, edge.Y2
        Else
            section.Contours.AddContourArc "A" & i, edge.X1, edge.Y1, edge.X2, edge.Y2, edge.CenterX, edge.CenterY, 70#, edge.Sweep
        End If
    Next i
    TestProjectionAddBar section, -85#, 0#, 100#: TestProjectionAddBar section, 85#, 0#, 100#
    TestProjectionAddBar section, 0#, 35#, 100#: TestProjectionAddBar section, 0#, -35#, 100#
    TestProjectionAddBar section, -50#, 0#, 100#
    Set TestProjectionStadiumSection = section
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Меняет деление полуокружностей, сохраняя фактическую ось и её размеры.
Private Function TestProjectionStadiumPath(ByVal radius As Double, ByVal direction As Long, ByVal pieces As Long) As CGeometryPath
    Dim path As New CGeometryPath, i As Long, sweep As Double, first As Double, centre As Double
    path.Initialize "Stadium": sweep = direction * GEOM_PI / pieces
    If direction = 1 Then path.AddLine "Bottom", -50#, -radius, 50#, -radius
    centre = 50#: first = -GEOM_PI / 2#: If direction = -1 Then centre = -50#
    For i = 0 To pieces - 1: path.AddArc "First" & i, centre, 0#, radius, first + i * sweep, sweep: Next i
    If direction = 1 Then path.AddLine "Top", 50#, radius, -50#, radius Else path.AddLine "Top", -50#, radius, 50#, radius
    centre = -centre: first = GEOM_PI / 2#
    For i = 0 To pieces - 1: path.AddArc "Second" & i, centre, 0#, radius, first + i * sweep, sweep: Next i
    If direction = -1 Then path.AddLine "Bottom", 50#, -radius, -50#, -radius
    path.Seal True: Set TestProjectionStadiumPath = path
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Задаёт физическую площадь и соответствующий диаметр без скрытой смены ID.
Private Sub TestProjectionAddBar(ByVal section As CSectionModel, ByVal x As Double, ByVal y As Double, ByVal area As Double)
    section.AddRebarElement x, y, Sqr(4# * area / GEOM_PI), area, "A400"
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет, что сумма сил относится к физическим стержням, без дублирующих ID.
Private Sub TestProjectionCalculatorOutput(ByVal section As CSectionModel)
    Dim p As TTorsionPostParameters, calculator As New CTorsionPostCalculator
    p.AxisOffset = 20#: p.HoopDiameter = 8#: p.HoopSpacing = 100#: p.BarBand = 1000#
    p.SteelResistance = 350#: p.HoopResistance = 300#: p.StrutAngle = 45#: p.ConcreteMethod = "None"
    calculator.Prepare section, p
    Dim result As CTorsionPostResult: Set result = calculator.Calculate("Equal", 1000000#, Nothing)
    Dim bars As Variant, row As Variant, i As Long, total As Double
    bars = result.Bars: row = result.Summary
    TestProjectionAssert "One output per physical bar", UBound(bars, 1) = section.RebarCount
    For i = 1 To UBound(bars, 1): total = total + bars(i, 3): Next i
    TestProjectionNear "Calculator force equilibrium", total, CDbl(row(1, 8))
    TestProjectionAssert "T available without projection error", CStr(row(1, 15)) = "OK" Or CStr(row(1, 15)) = "FAIL"
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Независимо суммирует коэффициент напряжения, умноженный на физическую площадь.
Private Sub TestProjectionEquilibrium(ByVal section As CSectionModel, ByRef coefficients() As Double, ByVal label As String)
    Dim i As Long, total As Double
    For i = 1 To section.RebarCount
        TestProjectionAssert label & " nonnegative " & i, coefficients(i) >= 0#
        total = total + coefficients(i) * section.RebarArea(i)
    Next i
    TestProjectionNear label & " physical force sum", total, 1#
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает аналитическое значение с относительным допуском.
Private Sub TestProjectionNear(ByVal label As String, ByVal actual As Double, ByVal expected As Double)
    TestProjectionAssert label, Abs(actual - expected) <= 0.000000001 * GeomMax(Abs(expected), 0.000001)
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сохраняет каждую проверку в первичном журнале, включая фактические отказы.
Private Sub TestProjectionAssert(ByVal label As String, ByVal passed As Boolean)
    If passed Then mTestPassed = mTestPassed + 1 Else mTestFailed = mTestFailed + 1
    mTestLog = mTestLog & vbCrLf & IIf(passed, "PASS ", "FAIL ") & label
End Sub
