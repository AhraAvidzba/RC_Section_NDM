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
    p.StateName = "StrengthState": p.AxisOffset = 20#: p.HoopDiameter = 8#: p.HoopSpacing = 150#
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
    Dim axisA As New CTorsionHoopGeometry, axisB As New CTorsionHoopGeometry
    axisA.InitializeAutomatic TestTorsionRectangle(), 20#, 8#
    axisB.InitializeAutomatic TestTorsionRectangle(), 20#, 12#
    TestTorsionNear "Axis offset excludes diameter addition", axisA.EnclosedArea, 41600#
    TestTorsionNear "Same axis offset gives same area with other diameter", axisB.EnclosedArea, axisA.EnclosedArea
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
' Проверяет независимость доступных результатов от отказа только бетонной проверки.
Public Function RunTorsionPostAvailabilityTests() As String
    mTestPassed = 0: mTestFailed = 0: mTestLog = vbNullString
    On Error GoTo Failed
    Dim section As CSectionModel: Set section = TestTorsionRectangle()
    Dim p As TTorsionPostParameters
    p.StateName = "StrengthState": p.AxisOffset = 20#: p.HoopDiameter = 8#: p.HoopSpacing = 150#
    p.SteelResistance = 270#: p.HoopResistance = 170#: p.BarBand = 25#: p.StrutAngle = 45#
    p.ConcreteMethod = "None": p.ConcreteResistance = 10#: p.ConcreteEfficiency = 0.5
    Dim state As CSectionStateResult: Set state = TestTorsionReadyPlane(section, 0.0005, True)
    Dim calculator As New CTorsionPostCalculator, result As CTorsionPostResult, expected As Variant, actual As Variant
    calculator.Prepare section, p: Set result = calculator.Calculate("Base", 3900000#, state): expected = result.Summary
    Dim scenario As Variant, column As Long
    For Each scenario In Array("UnknownMethod", "ExcessiveThickness", "InvalidConcreteInput")
        p.ConcreteMethod = "Strut": p.WallThickness = 0#: p.ConcreteInputInvalid = False
        Select Case CStr(scenario)
            Case "UnknownMethod": p.ConcreteMethod = "Unknown"
            Case "ExcessiveThickness": p.WallThickness = 1000#
            Case "InvalidConcreteInput": p.ConcreteInputInvalid = True
        End Select
        Set calculator = New CTorsionPostCalculator: calculator.Prepare section, p
        Set result = calculator.Calculate(CStr(scenario), 3900000#, state): actual = result.Summary
        TestTorsionAssert CStr(scenario) & " concrete InputErr", CStr(actual(1, 28)) = "InputErr"
        TestTorsionAssert CStr(scenario) & " overall cannot be OK", CStr(actual(1, 38)) = "InputErr"
        TestTorsionAssert CStr(scenario) & " no invented concrete limit", IsEmpty(actual(1, 26)) And IsEmpty(actual(1, 27))
        TestTorsionAssert CStr(scenario) & " reason retained", Len(CStr(actual(1, 39))) > 0
        For column = 3 To 21
            TestTorsionAssert CStr(scenario) & " T result " & column, CStr(actual(1, column)) = CStr(expected(1, column))
        Next column
        For column = 29 To 37
            TestTorsionAssert CStr(scenario) & " NMT result " & column, CStr(actual(1, column)) = CStr(expected(1, column))
        Next column
    Next scenario
    p.ConcreteMethod = "None": p.ConcreteInputInvalid = False: p.WallThickness = 0#
    p.HoopDiameter = Sqr(400# / (4# * Atn(1#))): p.HoopSpacing = 210#: p.HoopResistance = p.SteelResistance
    Set calculator = New CTorsionPostCalculator: calculator.Prepare section, p
    Set result = calculator.Calculate("EqualSteelDensities45", 3900000#, state): actual = result.Summary
    TestTorsionNear "Equal density at 45 gives same longitudinal and hoop stress", actual(1, 12), CDbl(actual(1, 18))
    TestTorsionNear "Equal resistance gives same T reserve", actual(1, 14), CDbl(actual(1, 20))
    p.BarBand = 1000000#
    Set calculator = New CTorsionPostCalculator: calculator.Prepare section, p
    Set result = calculator.Calculate("WideBand", 3900000#, state): actual = result.Summary
    TestTorsionNear "Band wider than maximum dimension keeps eligible area", actual(1, 9), expected(1, 9)
    TestTorsionNear "Wide band keeps longitudinal force", actual(1, 8), expected(1, 8)
    TestTorsionAssert "Wide band is not clamped to geometry", CDbl(actual(1, 37)) = 1000000#
    Set section = TestTorsionHollow(): p.ConcreteMethod = "GorskyRectangle": p.GorskyCoefficient = 0.07
    Set calculator = New CTorsionPostCalculator: calculator.Prepare section, p
    Set result = calculator.Calculate("WrongShapeForGorsky", 3900000#, Nothing): actual = result.Summary
    TestTorsionAssert "Gorsky shape error only in concrete", CStr(actual(1, 28)) = "InputErr" And CStr(actual(1, 15)) <> "InputErr" And CStr(actual(1, 21)) <> "InputErr"
    Dim bars As Variant: bars = result.Bars
    TestTorsionAssert "Hollow wide band keeps outer row rule", UBound(bars, 1) = 4
    GoTo Finished
Failed:
    mTestFailed = mTestFailed + 1: mTestLog = mTestLog & vbCrLf & "RUNTIME: " & Err.Description
Finished:
    RunTorsionPostAvailabilityTests = "TorsionPostAvailability: Passed=" & mTestPassed & "; Failed=" & mTestFailed & mTestLog
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет живые формулы единиц и одинаковые результаты при физически
' равном вводе в mm/MPa/tf*m и m/kPa/kN*m. Исходный ввод восстанавливается.
Public Function RunTorsionPostUnitTests() As String
    mTestPassed = 0: mTestFailed = 0: mTestLog = vbNullString
    Dim unitRange As Object, settingsRange As Object, torqueRange As Object, loads As Object
    Dim originalUnits As Variant, originalSettings As Variant, originalTorques As Variant, originalLoads As Variant
    On Error GoTo Failed
    Set unitRange = ThisWorkbook.Names("rngUnitSettings").RefersToRange
    Set settingsRange = ThisWorkbook.Names("rngTorsionSettings").RefersToRange
    Set torqueRange = ThisWorkbook.Names("rngTorsionCombinations").RefersToRange
    Set loads = ThisWorkbook.Names("rngLoadCombinations").RefersToRange
    originalUnits = unitRange.Formula: originalSettings = settingsRange.Formula
    originalTorques = torqueRange.Formula: originalLoads = loads.Formula
    TestTorsionSetTableValue unitRange, "Length", "mm": TestTorsionSetTableValue unitRange, "Stress", "MPa"
    TestTorsionSetTableValue unitRange, "Moment", "tf*m"
    TestTorsionSetTableValue settingsRange, "Torsion.StateType", "StrengthState"
    TestTorsionSetTableValue settingsRange, "Torsion.HoopAxisOffset", 29#
    TestTorsionSetTableValue settingsRange, "Torsion.HoopDiameter", 8#
    TestTorsionSetTableValue settingsRange, "Torsion.HoopSpacing", 100#
    TestTorsionSetTableValue settingsRange, "Torsion.Rs", 350#
    TestTorsionSetTableValue settingsRange, "Torsion.Rsw", 300#
    TestTorsionSetTableValue settingsRange, "Torsion.BarBand", 40#
    TestTorsionSetTableValue settingsRange, "Torsion.StrutAngle", 45#
    TestTorsionSetTableValue settingsRange, "Torsion.ConcreteMethod", "Strut"
    TestTorsionSetTableValue settingsRange, "Torsion.WallThickness", 0#
    TestTorsionSetTableValue settingsRange, "Torsion.Rb", 17.5
    TestTorsionSetTableValue settingsRange, "Torsion.ConcreteEfficiency", 0.5
    TestTorsionSetTableValue settingsRange, "Torsion.GorskyCoefficient", 0.07
    torqueRange.Offset(1, 0).Resize(torqueRange.Rows.Count - 1, 2).ClearContents
    loads.Cells(2, 1).Value2 = "Units01": torqueRange.Cells(2, 1).Value2 = "Units01": torqueRange.Cells(2, 2).Value2 = 1#
    Application.Calculate
    Dim reader As New CSystemSettingsReader, units As New CUnitSystem, batch As New CBatchSectionCalculator
    reader.LoadFromWorkbook ThisWorkbook: units.LoadFromSettings reader
    Dim baselineSnapshot As New CCalculationSnapshot, convertedSnapshot As New CCalculationSnapshot, warning As String, expected As Variant, actual As Variant
    warning = CaptureTorsionPostprocess(ThisWorkbook, TestTorsionRectangle(), batch, units, baselineSnapshot)
    expected = baselineSnapshot.Table(TORSION_POST_DATA_KEY)
    TestTorsionAssert "Baseline concrete calculation available", CStr(expected(1, 28)) = "OK" Or CStr(expected(1, 28)) = "FAIL"
    Dim row As Long, key As String, unitKind As String
    For row = 2 To settingsRange.Rows.Count
        key = CStr(settingsRange.Cells(row, 1).Value2): unitKind = TestTorsionUnitCategory(key)
        If Len(unitKind) > 0 Then TestTorsionAssert key & " unit is formula", CBool(settingsRange.Cells(row, 3).HasFormula)
    Next row
    TestTorsionSetTableValue unitRange, "Length", "m": TestTorsionSetTableValue unitRange, "Stress", "kPa"
    TestTorsionSetTableValue unitRange, "Moment", "kN*m"
    TestTorsionSetTableValue settingsRange, "Torsion.HoopAxisOffset", 0.029
    TestTorsionSetTableValue settingsRange, "Torsion.HoopDiameter", 0.008
    TestTorsionSetTableValue settingsRange, "Torsion.HoopSpacing", 0.1
    TestTorsionSetTableValue settingsRange, "Torsion.BarBand", 0.04
    TestTorsionSetTableValue settingsRange, "Torsion.Rs", 350000#
    TestTorsionSetTableValue settingsRange, "Torsion.Rsw", 300000#
    TestTorsionSetTableValue settingsRange, "Torsion.Rb", 17500#
    torqueRange.Cells(2, 2).Value2 = 9.80665
    Application.Calculate
    For row = 2 To settingsRange.Rows.Count
        key = CStr(settingsRange.Cells(row, 1).Value2): unitKind = TestTorsionUnitCategory(key)
        If Len(unitKind) > 0 Then TestTorsionAssert key & " follows INPUT", CStr(settingsRange.Cells(row, 3).Value2) = IIf(unitKind = "Length", "m", "kPa")
    Next row
    TestTorsionAssert "Torque header follows INPUT Moment", CStr(torqueRange.Cells(1, 2).Value2) = "T, kN*m"
    Set reader = New CSystemSettingsReader: Set units = New CUnitSystem
    reader.LoadFromWorkbook ThisWorkbook: units.LoadFromSettings reader
    warning = CaptureTorsionPostprocess(ThisWorkbook, TestTorsionRectangle(), batch, units, convertedSnapshot)
    actual = convertedSnapshot.Table(TORSION_POST_DATA_KEY)
    Dim column As Long
    For column = 1 To TORSION_POST_COLUMNS
        If IsNumeric(expected(1, column)) And Not IsEmpty(expected(1, column)) Then
            TestTorsionNear "Equivalent INPUT column " & column, actual(1, column), CDbl(expected(1, column))
        Else
            TestTorsionAssert "Equivalent INPUT column " & column, CStr(actual(1, column)) = CStr(expected(1, column))
        End If
    Next column
    GoTo Finished
Failed:
    mTestFailed = mTestFailed + 1: mTestLog = mTestLog & vbCrLf & "RUNTIME: " & Err.Description
Finished:
    On Error Resume Next
    If IsArray(originalUnits) Then unitRange.Formula = originalUnits
    If IsArray(originalSettings) Then settingsRange.Formula = originalSettings
    If IsArray(originalTorques) Then torqueRange.Formula = originalTorques
    If IsArray(originalLoads) Then loads.Formula = originalLoads
    Application.Calculate
    RunTorsionPostUnitTests = "TorsionPostUnits: Passed=" & mTestPassed & "; Failed=" & mTestFailed & mTestLog
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Меняет значение по ключу только в отдельной нативной проверочной копии.
Private Sub TestTorsionSetTableValue(ByVal table As Object, ByVal key As String, ByVal value As Variant)
    Dim row As Long
    For row = 2 To table.Rows.Count
        If CStr(table.Cells(row, 1).Value2) = key Then table.Cells(row, 2).Value2 = value: Exit Sub
    Next row
    Err.Raise 5, , "Test setting not found: " & key
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Выделяет физические категории единиц без привязки к адресам листа.
Private Function TestTorsionUnitCategory(ByVal key As String) As String
    Select Case key
        Case "Torsion.HoopAxisOffset", "Torsion.HoopDiameter", "Torsion.HoopSpacing", "Torsion.BarBand", "Torsion.WallThickness": TestTorsionUnitCategory = "Length"
        Case "Torsion.Rs", "Torsion.Rsw", "Torsion.Rb": TestTorsionUnitCategory = "Stress"
    End Select
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
    p.StateName = "StrengthState": p.AxisOffset = 20#: p.HoopDiameter = 8#: p.HoopSpacing = 150#
    p.SteelResistance = 270#: p.HoopResistance = 170#: p.BarBand = 25#: p.StrutAngle = 45#: p.ConcreteMethod = "None"
    Dim calculator As New CTorsionPostCalculator: calculator.Prepare TestTorsionRectangle(), p
    Dim results As New Collection: results.Add calculator.Calculate("Output01", 3900000#, Nothing)
    Dim snapshot As New CCalculationSnapshot, batch As New CBatchSectionCalculator
    Dim torsionWriter As New CTorsionSummaryWriter
    torsionWriter.CaptureSummary snapshot, results, units
    Dim warning As String: warning = torsionWriter.RestoreSummary(ThisWorkbook, snapshot)
    TestTorsionAssert "Output has no warning", Len(warning) = 0
    Dim expected As Variant: expected = snapshot.Table(TORSION_POST_DATA_KEY)
    Dim target As Object: Set target = ThisWorkbook.Names("rngTorsionSummaryAnchor").RefersToRange.Offset(-4, 0).Resize(4 + ThisWorkbook.Names("rngLoadCombinations").RefersToRange.Rows.Count - 1, TORSION_POST_DISPLAY_COLUMNS)
    Dim column As Long, map As Variant: map = torsionWriter.DisplayColumns()
    For column = 1 To TORSION_POST_DISPLAY_COLUMNS
        If map(column - 1) > 0 Then TestTorsionAssert "Output cell preserved " & CStr(column), CStr(target.Cells(5, column).Value2) = CStr(expected(1, map(column - 1)))
        TestTorsionAssert "Data cell unmerged " & CStr(column), Not CBool(target.Cells(5, column).MergeCells)
    Next column
    TestTorsionAssert "Group title stays in one row", target.Cells(1, 9).MergeArea.Rows.Count = 1
    TestTorsionAssert "ID has text format", CStr(target.Cells(5, 1).NumberFormat) = "@"
    TestTorsionAssert "Flow has three significant digits", CStr(target.Cells(5, 14).NumberFormat) = "0.00E+00"
    TestTorsionAssert "Flow fits preserved column width", InStr(1, CStr(target.Cells(5, 14).Text), "#", vbBinaryCompare) = 0
    TestTorsionAssert "Small hoop area is visible", Val(Replace$(CStr(target.Cells(5, 23).Text), ",", ".")) > 0#
    Dim headers As Variant, notesPresent As Boolean, noQuestions As Boolean
    headers = target.Resize(4, TORSION_POST_DISPLAY_COLUMNS).Value2: notesPresent = True: noQuestions = True
    For column = 9 To TORSION_POST_DISPLAY_COLUMNS
        If Len(CStr(headers(3, column))) = 0 Then notesPresent = False
        If InStr(1, CStr(headers(3, column)) & CStr(headers(4, column)), "?", vbBinaryCompare) > 0 Then noQuestions = False
    Next column
    TestTorsionAssert "Russian notes for result columns", notesPresent
    TestTorsionAssert "Header glyphs have no question marks", noQuestions
    TestTorsionAssert "Moment unit follows short label", CStr(headers(4, 7)) = "T, " & units.OutputMomentUnit
    TestTorsionAssert "Area unit follows short label", CStr(headers(4, 11)) = "A_o, " & units.OutputAreaUnit
    TestTorsionAssert "Greek stress and inline unit", CStr(headers(4, 25)) = ChrW$(&H3C3) & "_sw, " & units.OutputStressUnit
    TestTorsionAssert "Notes use peer alignment", target.Cells(3, 9).HorizontalAlignment = -4131 And Not CBool(target.Cells(3, 9).Font.Bold) And Not CBool(target.Cells(3, 9).WrapText)
    TestTorsionAssert "Short labels do not wrap", Not CBool(target.Cells(4, 9).WrapText)
    TestTorsionAssert "Rebar labels", CStr(headers(4, 18)) = "Rebar" And CStr(headers(4, 36)) = "Rebar"
    TestTorsionAssert "Band is in longitudinal subgroup", target.Cells(2, 10).MergeArea.Column = target.Cells(2, 9).Column
    TestTorsionAssert "Detailed name remains one cell", ThisWorkbook.Names("rngTorsionSummaryAnchor").RefersToRange.Cells.CountLarge = 1
    TestTorsionAssert "Bar name remains one cell", ThisWorkbook.Names("rngTorsionBarForces").RefersToRange.Cells.CountLarge = 1
    TestTorsionAssert "Bar data has no border", ThisWorkbook.Names("rngTorsionBarForces").RefersToRange.Offset(1, 0).Borders(7).LineStyle = -4142
    Dim combined(1 To 42, 1 To 40) As Variant
    combined(13, 1) = "Output01": combined(13, 4) = "NM_KEEP": combined(13, 3) = "NM_REASON_KEEP"
    torsionWriter.ApplyBatchSummary combined, snapshot
    TestTorsionAssert "Torsion concrete status in summary", CStr(combined(13, 16)) = CStr(expected(1, 28))
    TestTorsionAssert "NMT steel status in summary", CStr(combined(13, 17)) = CStr(expected(1, 36))
    TestTorsionAssert "Hoop status in summary", CStr(combined(13, 18)) = CStr(expected(1, 21))
    TestTorsionAssert "NM status and reason unchanged", CStr(combined(13, 4)) = "NM_KEEP" And CStr(combined(13, 3)) = "NM_REASON_KEEP"
    TestTorsionNear "Hoop reserve in summary", combined(13, 32), expected(1, 20)
    Dim summaryWriter As New CBatchResultWriter, title As Object, originalTitle As Variant
    Set title = ThisWorkbook.Names("rngStrengthSummaryAnchor").RefersToRange.Offset(-5, 0)
    originalTitle = title.Formula
    title.Value2 = "NM_TITLE_KEEP"
    summaryWriter.RestoreSummaryValues ThisWorkbook, combined, True
    TestTorsionAssert "Summary restore preserves detailed title", CStr(title.Value2) = "NM_TITLE_KEEP"
    title.Formula = originalTitle
    Dim stabilityWriter As New CStabilitySummaryWriter
    stabilityWriter.ClearSummary ThisWorkbook, True
    TestTorsionAssert "Stability clear preserves torsion data", CStr(target.Cells(5, 19).Value2) = CStr(expected(1, 12))
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
    Dim thickness As Object, originalThickness As Variant, originalLinked As Variant, methodCell As Object, originalMethod As Variant
    Dim loadRange As Object, torqueRange As Object, outputName As Object, sheet As Object
    On Error GoTo Failed
    Set outputName = ThisWorkbook.Names("rngTorsionSummaryAnchor"): originalRef = outputName.RefersTo
    Set sheet = ThisWorkbook.Worksheets("Results"): originalCell = sheet.Range("A49").Formula
    sheet.Range("A49").Value2 = "NM_KEEP"
    outputName.RefersTo = "=Results!$A$49"
    Dim torsionWriter As New CTorsionSummaryWriter
    Dim warning As String: warning = torsionWriter.ClearSummary(ThisWorkbook)
    TestTorsionAssert "Overlapping clear rejected", Len(warning) > 0
    TestTorsionAssert "NM data survived bad own range", CStr(sheet.Range("A49").Value2) = "NM_KEEP"
    outputName.RefersTo = originalRef: sheet.Range("A49").Formula = originalCell
    Set loadRange = ThisWorkbook.Names("rngLoadCombinations").RefersToRange
    Set torqueRange = ThisWorkbook.Names("rngTorsionCombinations").RefersToRange
    originalTorqueID = torqueRange.Cells(2, 1).Formula: originalID = loadRange.Cells(2, 1).Formula: originalTorque = torqueRange.Cells(2, 2).Formula
    loadRange.Cells(2, 1).Value2 = "Guard01": torqueRange.Cells(2, 1).Value2 = "Guard01": torqueRange.Cells(2, 2).Value2 = 1#
    Dim table As Object, row As Long: Set table = ThisWorkbook.Names("rngTorsionSettings").RefersToRange
    For row = 2 To table.Rows.Count
        If CStr(table.Cells(row, 1).Value2) = "Torsion.WallThickness" Then Set thickness = table.Cells(row, 2)
        If CStr(table.Cells(row, 1).Value2) = "Torsion.ConcreteMethod" Then Set methodCell = table.Cells(row, 2)
    Next row
    originalMethod = methodCell.Formula: methodCell.Value2 = "Strut"
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
    TestTorsionAssert "Bad thickness preserves longitudinal output", IsNumeric(data(1, 8)) And CStr(data(1, 15)) <> "InputErr"
    TestTorsionAssert "Bad thickness preserves hoop output", IsNumeric(data(1, 18)) And CStr(data(1, 21)) <> "InputErr"
    TestTorsionAssert "One clarification in row diagnostic", UBound(Split(CStr(data(1, 39)), "T (кручение)")) = 1
    methodCell.Formula = "='Расчет'!ZZ2"
    Set snapshot = New CCalculationSnapshot
    warning = CaptureTorsionPostprocess(ThisWorkbook, TestTorsionRectangle(), batch, units, snapshot)
    data = snapshot.Table(TORSION_POST_DATA_KEY)
    TestTorsionAssert "Unknown method blocks only concrete", CStr(data(1, 28)) = "InputErr" And CStr(data(1, 15)) <> "InputErr" And CStr(data(1, 21)) <> "InputErr"
    TestTorsionAssert "Unknown method keeps source chain", InStr(1, CStr(data(1, 39)), "Torsion.ConcreteMethod", vbBinaryCompare) > 0 And InStr(1, CStr(data(1, 39)), "ZZ2", vbBinaryCompare) > 0
    Dim barValues As Variant: barValues = snapshot.Table(TORSION_BAR_DATA_KEY)
    TestTorsionAssert "Unknown method retains bar forces", UBound(barValues, 1) = 5
    methodCell.Value2 = "None": thickness.Formula = "=1/0"
    Set snapshot = New CCalculationSnapshot
    warning = CaptureTorsionPostprocess(ThisWorkbook, TestTorsionRectangle(), batch, units, snapshot)
    data = snapshot.Table(TORSION_POST_DATA_KEY)
    TestTorsionAssert "None does not read unused thickness error", CStr(data(1, 28)) = "N/A" And CStr(data(1, 15)) <> "InputErr" And InStr(1, CStr(data(1, 39)), "Torsion.WallThickness", vbBinaryCompare) = 0
    Dim message As String, formatted As String
    message = "Ошибка T (кручение). Причина T (кручение); проверьте T (кручение)."
    formatted = WorkbookDialogText(message)
    TestTorsionAssert "One clarification in combined dialog", UBound(Split(formatted, "T (кручение)")) = 1
    TestTorsionAssert "Repeated dialog formatting is stable", WorkbookDialogText(formatted) = formatted
    message = "Значение «T (кручение)». Ошибка T (кручение); причина T (кручение)."
    formatted = WorkbookDialogText(message)
    TestTorsionAssert "Quoted input clarification is preserved", InStr(1, formatted, "«T (кручение)»", vbBinaryCompare) > 0 And UBound(Split(formatted, "T (кручение)")) = 2
    TestTorsionAssert "NM message remains unchanged", WorkbookDialogText("Расчёт НДМ завершён.") = "Расчёт НДМ завершён."
    GoTo Finished
Failed:
    mTestFailed = mTestFailed + 1: mTestLog = mTestLog & vbCrLf & "RUNTIME: " & Err.Description
Finished:
    On Error Resume Next
    If Not methodCell Is Nothing Then methodCell.Formula = originalMethod
    If Len(originalRef) > 0 Then outputName.RefersTo = originalRef: sheet.Range("A49").Formula = originalCell
    If Not thickness Is Nothing Then thickness.Formula = originalThickness: ThisWorkbook.Worksheets("Расчет").Range("ZZ2").Formula = originalLinked
    If Not loadRange Is Nothing Then loadRange.Cells(2, 1).Formula = originalID
    If Not torqueRange Is Nothing Then torqueRange.Cells(2, 1).Formula = originalTorqueID: torqueRange.Cells(2, 2).Formula = originalTorque
    On Error GoTo 0
    RunTorsionPostGuardTests = "TorsionPostGuard: Passed=" & CStr(mTestPassed) & "; Failed=" & CStr(mTestFailed) & mTestLog
End Function
