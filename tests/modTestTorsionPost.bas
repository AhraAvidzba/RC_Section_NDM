Attribute VB_Name = "modTestTorsionPost"
Option Explicit

' ==================== ДЛЯ ТЕСТОВ ====================
' Независимые аналитические числа прямоугольной фермы; не калибровка норм.
' Все данные синтетические и заданы во внутренних мм, Н, Н*мм, МПа.
Private mTestPassed As Long, mTestFailed As Long, mTestLog As String

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сверяет геометрию оси, тягу, силы стержней, хомут и два предела бетона.
Public Function RunTorsionPostTests() As String
    mTestPassed = 0: mTestFailed = 0: mTestLog = vbNullString
    On Error GoTo Failed
    Dim section As CSectionModel: Set section = TestTorsionRectangle()
    Dim p As TTorsionPostParameters
    p.StateName = "StrengthState": p.Cover = 16#: p.HoopDiameter = 8#: p.HoopSpacing = 150#
    p.SteelResistance = 270#: p.HoopResistance = 170#: p.BarBand = 25#: p.StrutAngle = 45#
    p.ConcreteMethod = "Strut": p.ConcreteResistance = 10#: p.ConcreteEfficiency = 0.5
    Dim calculator As New CTorsionPostCalculator: calculator.Prepare section, p
    Dim result As CTorsionPostResult: Set result = calculator.Calculate("T01", 3900000#, Nothing)
    Dim row As Variant, bars As Variant: row = result.Summary: bars = result.Bars
    TestTorsionNear "Ao = 160*260", row(1, 4), 41600#
    TestTorsionNear "uo = 2*(160+260)", row(1, 5), 840#
    TestTorsionNear "q = T/(2Ao)", row(1, 7), 46.875
    TestTorsionNear "FT = q*uo", row(1, 8), 39375#
    TestTorsionNear "AsT = FT/Rs", row(1, 10), 145.833333333333
    TestTorsionNear "deltaSigma = FT/(4*As)", row(1, 12), 98.4375
    TestTorsionNear "HoopArea = pi*d*d/4", row(1, 16), 16# * 4# * Atn(1#)
    TestTorsionNear "HoopStress = q*s/Asw", row(1, 18), 46.875 * 150# / (16# * 4# * Atn(1#))
    TestTorsionNear "StrutThickness = 2*axisCover", row(1, 23), 40#
    TestTorsionNear "StrutStress = 2*q/t", row(1, 24), 2.34375
    TestTorsionNear "StrutLimit", row(1, 26), 8320000#
    TestTorsionAssert "NoNM -> N/A, T retained", CStr(row(1, 36)) = "N/A" And CStr(row(1, 15)) = "OK"
    TestTorsionAssert "Compact bars three columns", UBound(bars, 2) = 3
    Dim i As Long, total As Double
    For i = 1 To UBound(bars, 1)
        total = total + CDbl(bars(i, 3))
        TestTorsionNear "BarForce" & CStr(i), bars(i, 3), 9843.75
    Next i
    TestTorsionNear "Allocated force equilibrium", total, 39375#
    Set result = calculator.Calculate("Tminus", -3900000#, Nothing): row = result.Summary
    TestTorsionNear "T sign invariance", row(1, 8), total
    p.ConcreteMethod = "GorskyRectangle": p.GorskyCoefficient = 0.07
    Set calculator = New CTorsionPostCalculator: calculator.Prepare section, p
    Set result = calculator.Calculate("TG", 3900000#, Nothing): row = result.Summary
    TestTorsionNear "Gorsky coefficient*c*b2*h", row(1, 26), 8400000#
    TestTorsionNear "Gorsky changes no longitudinal force", row(1, 8), 39375#
    p.ConcreteMethod = "None"
    Set calculator = New CTorsionPostCalculator: calculator.Prepare section, p
    Set result = calculator.Calculate("TnoConcrete", 3900000#, Nothing): row = result.Summary
    TestTorsionAssert "Concrete disabled", IsEmpty(row(1, 27)) And CStr(row(1, 28)) = "N/A"
    Dim state As CSectionStateResult
    Set state = TestTorsionReadyPlane(section, -0.0005, True)
    Set result = calculator.Calculate("Tcompression", 3900000#, state): row = result.Summary
    TestTorsionNear "NM compression retains sign", row(1, 30), -100#
    TestTorsionNear "Signed sum is negative", row(1, 33), -1.5625
    TestTorsionAssert "Compressive sum has no tensile limit", IsEmpty(row(1, 35)) And CStr(row(1, 36)) = "OK"
    Set state = TestTorsionReadyPlane(section, 0.0005, False)
    Set result = calculator.Calculate("Tphysical", 3900000#, state): row = result.Summary
    TestTorsionNear "Outside physical flag still uses saved NM", row(1, 30), 100#
    TestTorsionNear "Outside physical flag signed sum", row(1, 33), 198.4375
    TestTorsionAssert "NM reasons stay in NM", Len(CStr(row(1, 39))) = 0 And CStr(row(1, 36)) = "OK"
    p.HoopResistance = 1000000#
    Set calculator = New CTorsionPostCalculator: calculator.Prepare section, p
    Set state = TestTorsionReadyPlane(section, -0.0005, True)
    Set result = calculator.Calculate("Tcompensation", 12000000#, state): row = result.Summary
    TestTorsionAssert "T only failure does not override signed sum", CStr(row(1, 15)) = "FAIL" And CStr(row(1, 36)) = "OK" And CStr(row(1, 38)) = "OK"
    Set section = TestTorsionHollow(): p.BarBand = 200#
    Set calculator = New CTorsionPostCalculator: calculator.Prepare section, p
    Set result = calculator.Calculate("Thollow", 3900000#, Nothing): row = result.Summary: bars = result.Bars
    TestTorsionNear "Hollow uses outer area only", row(1, 9), 400#
    TestTorsionAssert "Wide band excludes inner row", UBound(bars, 1) = 4
    total = 0#
    For i = 1 To UBound(bars, 1)
        total = total + CDbl(bars(i, 3))
        TestTorsionAssert "Outer bar ID " & CStr(i), Val(Mid$(CStr(bars(i, 2)), 2)) <= 4
    Next i
    TestTorsionNear "Hollow allocates all tie force", total, 39375#
    GoTo Finished
Failed:
    mTestFailed = mTestFailed + 1: mTestLog = mTestLog & vbCrLf & "RUNTIME: " & Err.Description
Finished:
    RunTorsionPostTests = "TorsionPost: Passed=" & CStr(mTestPassed) & "; Failed=" & CStr(mTestFailed) & mTestLog
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Даёт готовые напряжения 200000*epsilon с достоверным снимком материалов;
' флаг физической допустимости задаётся независимо от равновесия плоскости.
Private Function TestTorsionReadyPlane(ByVal section As CSectionModel, ByVal epsilon As Double, ByVal physical As Boolean) As CSectionStateResult
    Dim concrete As New CConcreteMaterialParameters, steel As New CSteelMaterialParameters
    concrete.Initialize 15.5, 1.1, 22#, 1.8, 32500#, 32500#, rbMc2:=14.6
    steel.Initialize 350#, 350#, 390#, 390#, 200000#, 200000#
    Dim provider As New CMaterialModelProvider: provider.TestInitializeFromParameters concrete, steel
    Dim spec As New CMaterialModelSpec: spec.Initialize "ULS(I)", "TwoLine", "Ignore", "TwoLine"
    Dim solver As New CSectionSolver
    solver.EvaluateStrainPlane section, provider.ConcreteMaterialFromSpec(spec), provider.SteelMaterialFromSpec(spec), epsilon, 0#, 0#
    TestTorsionAssert "Saved plane confirmed", solver.ConfirmEquilibrium(solver.Nint, solver.Mxint, solver.Myint)
    Dim state As New CSectionStateResult
    state.InitializeFromSolver sstStrengthState, cpStrength, spec, solver, Not physical, physical
    TestTorsionAssert "Saved material snapshot exists", state.HasMaterialSnapshot
    Set TestTorsionReadyPlane = state
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Внутренний ряд ближе к отверстию, наружный — к наружной бетонной границе.
Private Function TestTorsionHollow() As CSectionModel
    Dim section As CSectionModel: Set section = TestTorsionRectangle()
    section.Contours.AddContourLine "OB", -35#, -75#, -35#, 75#, loopID:="HOLE", loopRole:="Opening"
    section.Contours.AddContourLine "OL", -35#, 75#, 35#, 75#, loopID:="HOLE", loopRole:="Opening"
    section.Contours.AddContourLine "OT", 35#, 75#, 35#, -75#, loopID:="HOLE", loopRole:="Opening"
    section.Contours.AddContourLine "OR", 35#, -75#, -35#, -75#, loopID:="HOLE", loopRole:="Opening"
    Dim x As Long, y As Long, index As Long: index = 4
    For x = -1 To 1 Step 2
        For y = -1 To 1 Step 2
            index = index + 1
            section.AddRebarElement 40# * x, 80# * y, Sqr(400# / (4# * Atn(1#))), 100#, "A400", index
        Next y
    Next x
    Set TestTorsionHollow = section
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет сохранность каждой ячейки после объединения и выравнивания шапок;
' запускать только в собственной тестовой копии с корректной структурой Results.
Public Function RunTorsionPostOutputTests() As String
    mTestPassed = 0: mTestFailed = 0: mTestLog = vbNullString
    On Error GoTo Failed
    Dim settings As New CSystemSettingsReader, units As New CUnitSystem
    settings.LoadFromWorkbook ThisWorkbook: units.LoadFromSettings settings
    Dim p As TTorsionPostParameters
    p.StateName = "StrengthState": p.Cover = 16#: p.HoopDiameter = 8#: p.HoopSpacing = 150#
    p.SteelResistance = 270#: p.HoopResistance = 170#: p.BarBand = 25#: p.StrutAngle = 45#: p.ConcreteMethod = "None"
    Dim calculator As New CTorsionPostCalculator: calculator.Prepare TestTorsionRectangle(), p
    Dim results As New Collection: results.Add calculator.Calculate("Output01", 3900000#, Nothing)
    Dim snapshot As New CCalculationSnapshot, batch As New CBatchSectionCalculator
    SaveTorsionPostMatrices snapshot, results, batch, units
    Dim warning As String: warning = RestoreTorsionPostTables(ThisWorkbook, snapshot)
    TestTorsionAssert "Output has no warning", Len(warning) = 0
    Dim expected As Variant: expected = snapshot.Table(TORSION_POST_DATA_KEY)
    Dim target As Object: Set target = ThisWorkbook.Names("rngTorsionPostSummary").RefersToRange
    Dim column As Long
    For column = 1 To TORSION_POST_COLUMNS
        TestTorsionAssert "Output cell preserved " & CStr(column), CStr(target.Cells(5, column).Value2) = CStr(expected(1, column))
        If column < 39 Then TestTorsionAssert "Data cell unmerged " & CStr(column), Not CBool(target.Cells(5, column).MergeCells)
    Next column
    TestTorsionAssert "Header spans only title row", target.Cells(1, 1).MergeArea.Rows.Count = 1
    TestTorsionAssert "ID has general format", CStr(target.Cells(5, 1).NumberFormat) = "General"
    TestTorsionAssert "Small hoop area is visible", Val(Replace$(CStr(target.Cells(5, 16).Text), ",", ".")) > 0#
    GoTo Finished
Failed:
    mTestFailed = mTestFailed + 1: mTestLog = mTestLog & vbCrLf & "RUNTIME: " & Err.Description
Finished:
    RunTorsionPostOutputTests = "TorsionPostOutput: Passed=" & CStr(mTestPassed) & "; Failed=" & CStr(mTestFailed) & mTestLog
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Строит сплошной прямоугольник 200*300 и четыре равноплощадных стержня.
Private Function TestTorsionRectangle() As CSectionModel
    Dim section As New CSectionModel, x As Long, y As Long, index As Long
    For x = -1 To 1 Step 2
        For y = -1 To 1 Step 2
            index = index + 1
            section.AddConcreteElement 50# * x, 75# * y, 15000#, index, shapeType:="Rectangle", width:=100#, height:=150#
            section.AddRebarElement 65# * x, 115# * y, Sqr(400# / (4# * Atn(1#))), 100#, "A400", index
        Next y
    Next x
    section.Contours.AddContourLine "B", -100#, -150#, 100#, -150#
    section.Contours.AddContourLine "R", 100#, -150#, 100#, 150#
    section.Contours.AddContourLine "T", 100#, 150#, -100#, 150#
    section.Contours.AddContourLine "L", -100#, 150#, -100#, -150#
    Set TestTorsionRectangle = section
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Независимое сравнение с относительным допуском для аналитических величин.
Private Sub TestTorsionNear(ByVal name As String, ByVal actual As Variant, ByVal expected As Double)
    TestTorsionAssert name, Abs(CDbl(actual) - expected) <= 1E-9 * GeomMax(1#, Abs(expected))
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Записывает имена всех утверждений, чтобы первичный журнал был проверяемым.
Private Sub TestTorsionAssert(ByVal name As String, ByVal passed As Boolean)
    If passed Then mTestPassed = mTestPassed + 1 Else mTestFailed = mTestFailed + 1
    mTestLog = mTestLog & vbCrLf & IIf(passed, "PASS ", "FAIL ") & name
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет адрес исходной формулы толщины и защиту таблиц НДМ от ошибочно
' перемещённого собственного имени. Изменяет только собственную тестовую копию.
Public Function RunTorsionPostGuardTests() As String
    mTestPassed = 0: mTestFailed = 0: mTestLog = vbNullString
    Dim originalRef As String, originalCell As Variant, originalID As Variant, originalTorque As Variant, originalTorqueID As Variant
    Dim thickness As Object, originalThickness As Variant, originalLinked As Variant
    Dim loadRange As Object, torqueRange As Object, outputName As Object, sheet As Object
    On Error GoTo Failed
    Set outputName = ThisWorkbook.Names("rngTorsionPostSummary"): originalRef = outputName.RefersTo
    Set sheet = ThisWorkbook.Worksheets("Results"): originalCell = sheet.Range("A49").Formula
    sheet.Range("A49").Value2 = "NM_KEEP"
    outputName.RefersTo = "=Results!$A$45:$AV$79"
    Dim warning As String: warning = ClearTorsionPostBodies(ThisWorkbook)
    TestTorsionAssert "Overlapping clear rejected", Len(warning) > 0
    TestTorsionAssert "NM data survived bad own range", CStr(sheet.Range("A49").Value2) = "NM_KEEP"
    outputName.RefersTo = originalRef: sheet.Range("A49").Formula = originalCell
    Set loadRange = ThisWorkbook.Names("rngLoadCombinations").RefersToRange
    Set torqueRange = ThisWorkbook.Names("rngTorsionCombinations").RefersToRange
    originalTorqueID = torqueRange.Cells(2, 1).Formula: originalID = loadRange.Cells(2, 1).Formula: originalTorque = torqueRange.Cells(2, 2).Formula
    loadRange.Cells(2, 1).Value2 = "Guard01": torqueRange.Cells(2, 1).Value2 = "Guard01": torqueRange.Cells(2, 2).Value2 = 1#
    Dim table As Object, row As Long: Set table = ThisWorkbook.Names("rngTorsionSettings").RefersToRange
    For row = 2 To table.Rows.Count
        If CStr(table.Cells(row, 1).Value2) = "Torsion.WallThickness" Then Set thickness = table.Cells(row, 2): Exit For
    Next row
    originalThickness = thickness.Formula
    originalLinked = ThisWorkbook.Worksheets("Расчет").Range("ZZ2").Formula
    ThisWorkbook.Worksheets("Расчет").Range("ZZ2").Value2 = 1000#
    thickness.Formula = "='Расчет'!ZZ2"
    Dim settings As New CSystemSettingsReader, units As New CUnitSystem
    settings.LoadFromWorkbook ThisWorkbook: units.LoadFromSettings settings
    Dim snapshot As New CCalculationSnapshot, batch As New CBatchSectionCalculator
    warning = CaptureTorsionPostprocess(ThisWorkbook, TestTorsionRectangle(), batch, units, snapshot)
    Dim data As Variant: data = snapshot.Table(TORSION_POST_DATA_KEY)
    TestTorsionAssert "Geometric setting error retained", CStr(data(1, 38)) = "InputErr"
    TestTorsionAssert "Wall thickness key retained", InStr(1, CStr(data(1, 39)), "Torsion.WallThickness", vbBinaryCompare) > 0
    mTestLog = mTestLog & vbCrLf & "INFO wall diagnostic: " & CStr(data(1, 39))
    TestTorsionAssert "Wall thickness source chain retained", InStr(1, CStr(data(1, 39)), "ZZ2", vbBinaryCompare) > 0
    GoTo Finished
Failed:
    mTestFailed = mTestFailed + 1: mTestLog = mTestLog & vbCrLf & "RUNTIME: " & Err.Description
Finished:
    On Error Resume Next
    If Len(originalRef) > 0 Then outputName.RefersTo = originalRef: sheet.Range("A49").Formula = originalCell
    If Not thickness Is Nothing Then thickness.Formula = originalThickness: ThisWorkbook.Worksheets("Расчет").Range("ZZ2").Formula = originalLinked
    If Not loadRange Is Nothing Then loadRange.Cells(2, 1).Formula = originalID
    If Not torqueRange Is Nothing Then torqueRange.Cells(2, 1).Formula = originalTorqueID: torqueRange.Cells(2, 2).Formula = originalTorque
    On Error GoTo 0
    RunTorsionPostGuardTests = "TorsionPostGuard: Passed=" & CStr(mTestPassed) & "; Failed=" & CStr(mTestFailed) & mTestLog
End Function
