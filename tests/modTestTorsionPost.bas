Attribute VB_Name = "modTestTorsionPost"
Option Explicit

' ==================== ДЛЯ ТЕСТОВ ====================
' Независимые аналитические числа прямоугольной фермы; не калибровка норм.
' Все данные синтетические и заданы во внутренних мм, Н, Н*мм, МПа.
Private mTestPassed As Long, mTestFailed As Long, mTestLog As String

' -------------------- ДЛЯ ТЕСТОВ --------------------
' При явном нуле NM сумма равна T; отсутствующее ненулевое состояние по-прежнему не подменяется нулём.
Public Function RunTorsionOnlyTests() As String
    mTestPassed = 0: mTestFailed = 0: mTestLog = vbNullString
    On Error GoTo Failed
    Dim section As CSectionModel: Set section = TestTorsionRectangle()
    Dim p As TTorsionPostParameters
    p.StateName = "StrengthState": p.AxisOffset = 20#: p.HoopDiameter = 8#: p.HoopSpacing = 150#
    p.SteelResistance = 270#: p.HoopResistance = 170#: p.BarBand = 25#: p.StrutAngle = 45#: p.ConcreteMethod = "None"
    Dim calculator As New CTorsionPostCalculator: calculator.Prepare section, p
    Dim result As CTorsionPostResult, values As Variant, bars As Variant, i As Long, total As Double
    Set result = calculator.Calculate("T_ONLY", 3900000#, Nothing, True)
    values = TestTorsionSummary(result): bars = TestTorsionBars(result)
    TestTorsionNear "Zero NM exactly", values(1, 30), 0#
    TestTorsionNear "T + zero NM", values(1, 33), 98.4375
    TestTorsionNear "Same governing T increment", values(1, 31), CDbl(values(1, 12))
    TestTorsionNear "Same reserve T and NMT", values(1, 35), CDbl(values(1, 14))
    TestTorsionAssert "Sum is available", CStr(values(1, 36)) = "OK" And CStr(values(1, 38)) = "OK"
    TestTorsionAssert "No missing state reason", Len(result.StateFailureReason) = 0 And Len(CStr(values(1, 39))) = 0
    For i = 1 To UBound(bars, 1): total = total + CDbl(bars(i, 4)): Next i
    TestTorsionNear "Force preserved", total, 39375#
    Set result = calculator.Calculate("MISSING_NM", 3900000#, Nothing)
    values = TestTorsionSummary(result)
    TestTorsionAssert "Missing nonzero NM is unavailable", CStr(values(1, 36)) = "N/A" And Len(result.StateFailureReason) > 0
    section.Contours.Clear
    Set calculator = New CTorsionPostCalculator
    Dim number As Long, reason As String
    On Error Resume Next
    calculator.Prepare section, p
    number = Err.Number: reason = Err.Description: Err.Clear
    On Error GoTo Failed
    TestTorsionAssert "Missing outer blocks only T preparation", number <> 0
    TestTorsionAssert "Missing outer cause is clear", InStr(1, reason, "нет достоверного наружного контура", vbTextCompare) > 0
    TestTorsionAssert "Missing outer gives action", InStr(1, reason, "снова сформируйте сечение", vbTextCompare) > 0
    TestTorsionAssert "Missing outer does not blame hoop inputs", Len(calculator.FailureSettingKeys) = 0
    GoTo Finished
Failed:
    mTestFailed = mTestFailed + 1: mTestLog = mTestLog & vbCrLf & "RUNTIME: " & Err.Description
Finished:
    RunTorsionOnlyTests = "TorsionOnly: Passed=" & CStr(mTestPassed) & "; Failed=" & CStr(mTestFailed) & mTestLog
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сверяет геометрию оси, тягу, силы стержней, хомут и два предела бетона.
Public Function RunTorsionPostTests() As String
    mTestPassed = 0: mTestFailed = 0: mTestLog = vbNullString
    On Error GoTo Failed
    Dim section As CSectionModel: Set section = TestTorsionRectangle()
    Dim p As TTorsionPostParameters
    p.StateName = "StrengthState": p.AxisOffset = 20#: p.HoopDiameter = 8#: p.HoopSpacing = 150#
    p.SteelResistance = 270#: p.HoopResistance = 170#: p.BarBand = 25#: p.StrutAngle = 45#
    p.ConcreteMethod = "Strut": p.ConcreteResistance = 10#: p.CharacteristicConcreteStrength = 20#: p.ConcretePartialFactor = 2#: p.LongitudinalAxisOffset = 20#
    Dim calculator As New CTorsionPostCalculator: calculator.Prepare section, p
    Dim result As CTorsionPostResult: Set result = calculator.Calculate("T01", 3900000#, Nothing)
    Dim row As Variant, bars As Variant: row = TestTorsionSummary(result): bars = TestTorsionBars(result)
    TestTorsionNear "Ao = 160*260", row(1, 4), 41600#
    TestTorsionNear "uo = 2*(160+260)", row(1, 5), 840#
    TestTorsionNear "q = T/(2Ao)", row(1, 7), 46.875
    TestTorsionNear "FT = q*uo", row(1, 8), 39375#
    TestTorsionNear "AsT = FT/Rs", row(1, 10), 145.833333333333
    TestTorsionNear "deltaSigma = FT/(4*As)", row(1, 12), 98.4375
    TestTorsionNear "HoopArea = pi*d*d/4", row(1, 16), 16# * 4# * Atn(1#)
    TestTorsionNear "HoopStress = q*s/Asw", row(1, 18), 46.875 * 150# / (16# * 4# * Atn(1#))
    TestTorsionNear "StrutThickness = max(A/u,2cl)", row(1, 23), 60#
    TestTorsionNear "StrutStress = T/(Ak*t) at 45 degrees", row(1, 24), 3900000# / (33600# * 60#)
    TestTorsionNear "StrutLimit = nu*fcd*Ak*t", row(1, 26), 11128320#
    Dim axisA As New CTorsionHoopGeometry, axisB As New CTorsionHoopGeometry
    axisA.InitializeAutomatic TestTorsionRectangle(), 20#, 8#
    axisB.InitializeAutomatic TestTorsionRectangle(), 20#, 12#
    TestTorsionNear "Axis offset excludes diameter addition", axisA.EnclosedArea, 41600#
    TestTorsionNear "Same axis offset gives same area with other diameter", axisB.EnclosedArea, axisA.EnclosedArea
    TestTorsionAssert "NoNM -> N/A, T retained", CStr(row(1, 36)) = "N/A" And CStr(row(1, 15)) = "OK"
    TestTorsionAssert "Compact bars with length last", UBound(bars, 2) = 5
    Dim i As Long, total As Double
    For i = 1 To UBound(bars, 1)
        total = total + CDbl(bars(i, 4))
        TestTorsionNear "BarForce" & CStr(i), bars(i, 4), 9843.75
        TestTorsionNear "Assigned perimeter length" & CStr(i), bars(i, 5), 210#
    Next i
    TestTorsionNear "Allocated force equilibrium", total, 39375#
    TestTorsionAssert "T determining bar owns its placement", CStr(row(1, 43)) = section.RebarComment(CLng(Mid$(CStr(row(1, 11)), 2)))
    Set result = calculator.Calculate("Tminus", -3900000#, Nothing): row = TestTorsionSummary(result)
    TestTorsionNear "T sign invariance", row(1, 8), total
    p.ConcreteMethod = "GorskyRectangle": p.GorskyCoefficient = 0.07
    Set calculator = New CTorsionPostCalculator: calculator.Prepare section, p
    Set result = calculator.Calculate("TG", 3900000#, Nothing): row = TestTorsionSummary(result)
    TestTorsionNear "Gorsky coefficient*c*b2*h", row(1, 26), 8400000#
    TestTorsionNear "Gorsky changes no longitudinal force", row(1, 8), 39375#
    p.ConcreteMethod = "None"
    Set calculator = New CTorsionPostCalculator: calculator.Prepare section, p
    Set result = calculator.Calculate("TnoConcrete", 3900000#, Nothing): row = TestTorsionSummary(result)
    TestTorsionAssert "Concrete disabled", IsEmpty(row(1, 27)) And CStr(row(1, 28)) = "N/A"
    Dim state As CSectionStateResult
    Set state = TestTorsionReadyPlane(section, -0.0005, True)
    Set result = calculator.Calculate("Tcompression", 3900000#, state): row = TestTorsionSummary(result)
    TestTorsionAssert "NMT determining bar owns its placement", CStr(row(1, 44)) = section.RebarComment(CLng(Mid$(CStr(row(1, 29)), 2)))
    TestTorsionNear "NM compression retains sign", row(1, 30), -100#
    TestTorsionNear "Signed sum is negative", row(1, 33), -1.5625
    TestTorsionAssert "Compressive sum has no tensile limit", IsEmpty(row(1, 35)) And CStr(row(1, 36)) = "OK"
    Set state = TestTorsionReadyPlane(section, 0.0005, False)
    Set result = calculator.Calculate("Tphysical", 3900000#, state): row = TestTorsionSummary(result)
    TestTorsionNear "Outside physical flag still uses saved NM", row(1, 30), 100#
    TestTorsionNear "Outside physical flag signed sum", row(1, 33), 198.4375
    TestTorsionAssert "NM reasons stay in NM", Len(CStr(row(1, 39))) = 0 And CStr(row(1, 36)) = "OK"
    p.HoopResistance = 1000000#
    Set calculator = New CTorsionPostCalculator: calculator.Prepare section, p
    Set state = TestTorsionReadyPlane(section, -0.0005, True)
    Set result = calculator.Calculate("Tcompensation", 12000000#, state): row = TestTorsionSummary(result)
    TestTorsionAssert "T only failure does not override signed sum", CStr(row(1, 15)) = "FAIL" And CStr(row(1, 36)) = "OK" And CStr(row(1, 38)) = "OK"
    Set section = TestTorsionHollow(): p.BarBand = 200#
    Set calculator = New CTorsionPostCalculator: calculator.Prepare section, p
    Set result = calculator.Calculate("Thollow", 3900000#, Nothing): row = TestTorsionSummary(result): bars = TestTorsionBars(result)
    TestTorsionNear "Hollow uses outer area only", row(1, 9), 400#
    TestTorsionAssert "Wide band excludes inner row", UBound(bars, 1) = 4
    total = 0#
    For i = 1 To UBound(bars, 1)
        total = total + CDbl(bars(i, 4))
        TestTorsionAssert "Outer bar ID " & CStr(i), Val(Mid$(CStr(bars(i, 2)), 2)) <= 4
    Next i
    TestTorsionNear "Hollow allocates all tie force", total, 39375#
    TestTypedTorsionResult
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
    p.ConcreteMethod = "None": p.ConcreteResistance = 10#: p.CharacteristicConcreteStrength = 20#: p.ConcretePartialFactor = 2#: p.LongitudinalAxisOffset = 20#
    Dim state As CSectionStateResult: Set state = TestTorsionReadyPlane(section, 0.0005, True)
    Dim calculator As New CTorsionPostCalculator, result As CTorsionPostResult, expected As Variant, actual As Variant
    calculator.Prepare section, p: Set result = calculator.Calculate("Base", 3900000#, state): expected = TestTorsionSummary(result)
    Dim scenario As Variant, column As Long
    For Each scenario In Array("UnknownMethod", "ExcessiveThickness", "InvalidConcreteInput")
        p.ConcreteMethod = "Strut": p.WallThickness = 0#: p.ConcreteInputInvalid = False
        Select Case CStr(scenario)
            Case "UnknownMethod": p.ConcreteMethod = "Unknown"
            Case "ExcessiveThickness": p.WallThickness = 1000#
            Case "InvalidConcreteInput": p.ConcreteInputInvalid = True
        End Select
        Set calculator = New CTorsionPostCalculator: calculator.Prepare section, p
        Set result = calculator.Calculate(CStr(scenario), 3900000#, state): actual = TestTorsionSummary(result)
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
    Set result = calculator.Calculate("EqualSteelDensities45", 3900000#, state): actual = TestTorsionSummary(result)
    TestTorsionNear "Equal density at 45 gives same longitudinal and hoop stress", actual(1, 12), CDbl(actual(1, 18))
    TestTorsionNear "Equal resistance gives same T reserve", actual(1, 14), CDbl(actual(1, 20))
    p.BarBand = 1000000#
    Set calculator = New CTorsionPostCalculator: calculator.Prepare section, p
    Set result = calculator.Calculate("WideBand", 3900000#, state): actual = TestTorsionSummary(result)
    TestTorsionNear "Band wider than maximum dimension keeps eligible area", actual(1, 9), expected(1, 9)
    TestTorsionNear "Wide band keeps longitudinal force", actual(1, 8), expected(1, 8)
    TestTorsionAssert "Wide band is not clamped to geometry", CDbl(actual(1, 37)) = 1000000#
    Set section = TestTorsionHollow(): p.ConcreteMethod = "GorskyRectangle": p.GorskyCoefficient = 0.07
    Set calculator = New CTorsionPostCalculator: calculator.Prepare section, p
    Set result = calculator.Calculate("WrongShapeForGorsky", 3900000#, Nothing): actual = TestTorsionSummary(result)
    TestTorsionAssert "Gorsky shape error only in concrete", CStr(actual(1, 28)) = "InputErr" And CStr(actual(1, 15)) <> "InputErr" And CStr(actual(1, 21)) <> "InputErr"
    Dim bars As Variant: bars = TestTorsionBars(result)
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
    TestTorsionSetTableValue settingsRange, "Torsion.Fck", 20#
    TestTorsionSetTableValue settingsRange, "Torsion.GammaC", 1.5
    TestTorsionSetTableValue settingsRange, "Torsion.RebarAxisOffset", 50#
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
    TestTorsionSetTableValue settingsRange, "Torsion.Fck", 20000#
    TestTorsionSetTableValue settingsRange, "Torsion.RebarAxisOffset", 0.05
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
    If IsArray(originalSettings) Then
        For row = 2 To settingsRange.Rows.Count
            If Left$(CStr(originalSettings(row, 1)), 8) = "Torsion." Then settingsRange.Cells(row, 2).Formula = originalSettings(row, 2)
        Next row
    End If
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
        Case "Torsion.HoopAxisOffset", "Torsion.HoopDiameter", "Torsion.HoopSpacing", "Torsion.BarBand", "Torsion.WallThickness", "Torsion.RebarAxisOffset": TestTorsionUnitCategory = "Length"
        Case "Torsion.Rs", "Torsion.Rsw", "Torsion.Rb", "Torsion.Fck": TestTorsionUnitCategory = "Stress"
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
    TestTorsionAssert "Small hoop area is visible", Val(Replace$(CStr(target.Cells(5, 24).Text), ",", ".")) > 0#
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
    TestTorsionAssert "Greek stress and inline unit", CStr(headers(4, 26)) = ChrW$(&H3C3) & "_sw, " & units.OutputStressUnit
    TestTorsionAssert "Notes use peer alignment", target.Cells(3, 9).HorizontalAlignment = -4131 And Not CBool(target.Cells(3, 9).Font.Bold) And Not CBool(target.Cells(3, 9).WrapText)
    TestTorsionAssert "Short labels do not wrap", Not CBool(target.Cells(4, 9).WrapText)
    TestTorsionAssert "Rebar labels", CStr(headers(4, 18)) = "Rebar" And CStr(headers(4, 40)) = "Rebar"
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
    TestTorsionAssert "Stability clear preserves torsion data", CStr(target.Cells(5, 20).Value2) = CStr(expected(1, 12))
    GoTo Finished
Failed:
    mTestFailed = mTestFailed + 1: mTestLog = mTestLog & vbCrLf & "RUNTIME: " & Err.Description
Finished:
    RunTorsionPostOutputTests = "TorsionPostOutput: Passed=" & CStr(mTestPassed) & "; Failed=" & CStr(mTestFailed) & mTestLog
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет длины в OUTPUT mm/m, точность двоичного файла и чтение трёхколоночных матриц.
' Только собственная проверочная книга: меняет таблицу T и создаёт временный файл рядом.
Public Function RunTorsionBarLengthTests() As String
    mTestPassed = 0: mTestFailed = 0: mTestLog = vbNullString
    On Error GoTo Failed
    Dim p As TTorsionPostParameters, calculator As CTorsionPostCalculator, units As CUnitSystem
    Dim settings As New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    p.AxisOffset = 20#: p.HoopDiameter = 8#: p.HoopSpacing = 150#: p.BarBand = 25#
    p.SteelResistance = 270#: p.HoopResistance = 170#: p.StrutAngle = 30#: p.ConcreteMethod = "None"
    Dim pass As Long, snapshot As CCalculationSnapshot, oldSnapshot As CCalculationSnapshot, writer As New CTorsionSummaryWriter
    Dim results As Collection, bars As Variant, data As Variant, headers As Variant, legacy() As Variant, row As Long, column As Long
    Dim anchor As Object: Set anchor = ThisWorkbook.Names("rngTorsionBarForces").RefersToRange
    Dim width As Double: width = anchor.Offset(0, 4).ColumnWidth
    Dim codec As New CCalculationSnapshotStore, indexed As CIndexedSnapshotFile, path As String, actual As Variant, warning As String
    For pass = 1 To 2
        Set units = New CUnitSystem
        If pass = 1 Then units.LoadFromSettings settings Else units.InitializeDefaults ' Штатная книга выводит mm; стандарт ядра для OUTPUT Length — m.
        Set calculator = New CTorsionPostCalculator: calculator.Prepare TestTorsionRectangle(), p
        Set results = New Collection: results.Add calculator.Calculate("Length01", 3900000#, Nothing)
        results.Add calculator.Calculate("Length02", 3900000#, Nothing) ' Два сочетания проверяют полный вывод строк, а не только первого блока стержней.
        Set snapshot = New CCalculationSnapshot: writer.CaptureSummary snapshot, results, units
        bars = snapshot.Table(TORSION_BAR_DATA_KEY): data = snapshot.Table(TORSION_POST_DATA_KEY)
        headers = snapshot.Table(TORSION_POST_HEADERS_KEY)
        TestTorsionAssert "Face header " & pass, CStr(bars(1, 3)) = "Грань"
        TestTorsionAssert "Length unit header " & pass, CStr(bars(1, 5)) = "l_i, " & units.OutputLengthUnit
        For row = 2 To UBound(bars, 1)
            TestTorsionAssert "Own placement " & pass & ":" & row, CStr(bars(row, 3)) = "Грань стержня " & CStr(bars(row, 2))
            TestTorsionNear "Output length " & pass & ":" & row, bars(row, 5), units.InternalLengthToOutput(210#)
            TestTorsionNear "Force matches length at 30 degrees " & pass & ":" & row, bars(row, 4), CDbl(data(1, 7)) * CDbl(bars(row, 5)) * CDbl(data(1, 6))
        Next row
        warning = writer.RestoreSummary(ThisWorkbook, snapshot)
        TestTorsionAssert "Five-column restore " & pass, Len(warning) = 0
        TestTorsionNear "Runtime preserves width " & pass, anchor.Offset(0, 4).ColumnWidth, width
        TestTorsionNear "Length cell " & pass, anchor.Offset(1, 4).Value2, CDbl(bars(2, 5))
        ReDim legacy(1 To UBound(bars, 1), 1 To 3)
        For row = 1 To UBound(bars, 1): For column = 1 To 2: legacy(row, column) = bars(row, column): Next column: legacy(row, 3) = bars(row, 4): Next row
        Set oldSnapshot = New CCalculationSnapshot
        oldSnapshot.AddTable TORSION_POST_HEADERS_KEY, headers
        oldSnapshot.AddTable TORSION_POST_DATA_KEY, data
        oldSnapshot.AddTable TORSION_BAR_DATA_KEY, legacy
        warning = writer.RestoreSummary(ThisWorkbook, oldSnapshot)
        TestTorsionAssert "Three-column restore " & pass, Len(warning) = 0
        TestTorsionAssert "Legacy saved unit " & pass, CStr(anchor.Offset(0, 4).Value2) = CStr(bars(1, 5))
        For row = 2 To UBound(bars, 1)
            TestTorsionNear "Legacy reconstructed length " & pass & ":" & row, anchor.Offset(row - 1, 4).Value2, CDbl(bars(row, 5))
            TestTorsionNear "Legacy force unchanged " & pass & ":" & row, anchor.Offset(row - 1, 3).Value2, CDbl(bars(row, 4))
        Next row
        ReDim legacy(1 To UBound(bars, 1), 1 To 4)
        For row = 1 To UBound(bars, 1)
            For column = 1 To 2: legacy(row, column) = bars(row, column): Next column: legacy(row, 3) = bars(row, 4)
            legacy(row, 4) = bars(row, 5)
        Next row
        Set oldSnapshot = New CCalculationSnapshot
        oldSnapshot.AddTable TORSION_POST_HEADERS_KEY, headers
        oldSnapshot.AddTable TORSION_POST_DATA_KEY, data
        oldSnapshot.AddTable TORSION_BAR_DATA_KEY, legacy
        warning = writer.RestoreSummary(ThisWorkbook, oldSnapshot)
        TestTorsionAssert "Four-column saved length restored " & pass, Len(warning) = 0
        TestTorsionNear "Four-column saved length unchanged " & pass, anchor.Offset(1, 4).Value2, CDbl(bars(2, 5))
        Dim elementHeader(1 To 1, 1 To 9) As Variant
        Dim elementLabels As Variant: elementLabels = Array("RunID", "LoadCase", "ProfileId", "StateType", "ElementID", "MaterialType", "Strain", "Stress", "PhysicalState")
        For column = 1 To 9: elementHeader(1, column) = elementLabels(column - 1): Next column
        snapshot.AddTable "rngNDMElementResults", elementHeader ' Контейнер v6 требует шапку и каталог элементов даже при проверке отдельной таблицы.
        path = ThisWorkbook.Path & Application.PathSeparator & "TestTorsionBarLength_" & CStr(pass) & ".ndm"
        Set indexed = New CIndexedSnapshotFile: indexed.WriteFile path, snapshot, codec
        TestTorsionAssert "File opened " & pass, indexed.OpenRead(path, codec)
        actual = indexed.Table(TORSION_BAR_DATA_KEY): ValidateSnapshotTableShape TORSION_BAR_DATA_KEY, actual
        For row = 1 To UBound(bars, 1): For column = 1 To 5
            TestTorsionAssert "File exact " & pass & ":" & row & ":" & column, CStr(actual(row, column)) = CStr(bars(row, column))
        Next column: Next row
        warning = writer.RestoreSummary(ThisWorkbook, indexed)
        TestTorsionAssert "File restore " & pass, Len(warning) = 0
        indexed.CloseFile: Kill path: path = vbNullString
        TestTorsionAssert "File both combinations retained " & pass, UBound(bars, 1) = 9 And CStr(anchor.Offset(8, 0).Value2) = "Length02"
        For row = 1 To UBound(bars, 1): For column = 1 To TORSION_BAR_COLUMNS
            TestTorsionAssert "File sheet exact " & pass & ":" & row & ":" & column, CStr(anchor.Offset(row - 1, column - 1).Value2) = CStr(bars(row, column))
        Next column: Next row
        warning = writer.ClearSummary(ThisWorkbook)
        TestTorsionAssert "Clear length data " & pass, Len(warning) = 0 And IsEmpty(anchor.Offset(1, 4).Value2)
    Next pass
    GoTo Finished
Failed:
    mTestFailed = mTestFailed + 1: mTestLog = mTestLog & vbCrLf & "RUNTIME: " & Err.Description
Finished:
    On Error Resume Next
    If Not indexed Is Nothing Then indexed.CloseFile
    If Len(path) > 0 Then Kill path
    On Error GoTo 0
    RunTorsionBarLengthTests = "TorsionBarLength: Passed=" & mTestPassed & "; Failed=" & mTestFailed & mTestLog
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Строит сплошной прямоугольник 200*300 и четыре равноплощадных стержня.
Private Function TestTorsionRectangle() As CSectionModel
    Dim section As New CSectionModel, x As Long, y As Long, index As Long
    For x = -1 To 1 Step 2
        For y = -1 To 1 Step 2
            index = index + 1
            section.AddConcreteElement 50# * x, 75# * y, 15000#, index, shapeType:="Rectangle", width:=100#, height:=150#
            section.AddRebarElement 65# * x, 115# * y, Sqr(400# / (4# * Atn(1#))), 100#, "A400", index, comment:="Грань стержня R" & CStr(index)
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
    TestTorsionAssert name, Abs(CDbl(actual) - expected) <= 0.000000001 * GeomMax(1#, Abs(expected))
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

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Различает адрес ошибки строки T и адрес настройки, включая повторное использование
' ошибки подготовки для нескольких сочетаний. Все изменения выполняются в тестовой копии и восстанавливаются.
Public Function RunTorsionPostDiagnosticTests() As String
    mTestPassed = 0: mTestFailed = 0: mTestLog = vbNullString
    Dim loads As Object, torques As Object, table As Object, axis As Object, linked As Object
    Dim originalLoads As Variant, originalTorques As Variant, originalAxis As Variant, originalLinked As Variant
    Dim row As Long, settings As New CSystemSettingsReader, units As New CUnitSystem
    Dim snapshot As CCalculationSnapshot, batch As New CBatchSectionCalculator, data As Variant
    Dim warning As String, reason As String, torqueAddress As String, axisAddress As String, identifierAddress As String
    On Error GoTo Failed
    Set loads = ThisWorkbook.Names("rngLoadCombinations").RefersToRange.Cells(2, 1).Resize(2, 1)
    Set torques = ThisWorkbook.Names("rngTorsionCombinations").RefersToRange
    Set table = ThisWorkbook.Names("rngTorsionSettings").RefersToRange
    For row = 2 To table.Rows.Count
        If CStr(table.Cells(row, 1).Value2) = "Torsion.HoopAxisOffset" Then Set axis = table.Cells(row, 2)
    Next row
    Set linked = ThisWorkbook.Worksheets("Расчет").Range("ZZ2")
    originalLoads = loads.Formula: originalTorques = torques.Formula
    originalAxis = axis.Formula: originalLinked = linked.Formula
    torqueAddress = torques.Worksheet.Name & "!" & torques.Cells(2, 2).Address(False, False)
    identifierAddress = torques.Worksheet.Name & "!" & torques.Cells(2, 1).Address(False, False)
    axisAddress = axis.Worksheet.Name & "!" & axis.Address(False, False)
    loads.Cells(1, 1).Value2 = "TestDiag1": loads.Cells(2, 1).Value2 = "TestDiag2"
    torques.Offset(1, 0).Resize(torques.Rows.Count - 1, 2).ClearContents
    torques.Cells(2, 1).Value2 = "TestDiag1": torques.Cells(2, 2).Value2 = 10#
    torques.Cells(3, 1).Value2 = "TestDiag2": torques.Cells(3, 2).Value2 = 20#
    linked.Value2 = "29+": axis.Formula = "='Расчет'!ZZ2"
    settings.LoadFromWorkbook ThisWorkbook: units.LoadFromSettings settings
    Set snapshot = New CCalculationSnapshot
    warning = CaptureTorsionPostprocess(ThisWorkbook, TestTorsionRectangle(), batch, units, snapshot)
    data = snapshot.Table(TORSION_POST_DATA_KEY)
    TestTorsionAssert "Both requested combinations preserved", UBound(data, 1) = 2
    For row = 1 To 2
        reason = CStr(data(row, 39))
        TestTorsionAssert "Setting failure status " & CStr(row), CStr(data(row, 38)) = "InputErr"
        TestTorsionAssert "Setting key retained " & CStr(row), InStr(1, reason, "Torsion.HoopAxisOffset", vbBinaryCompare) > 0
        TestTorsionAssert "Setting and formula source retained " & CStr(row), InStr(1, reason, axisAddress, vbBinaryCompare) > 0 And InStr(1, reason, "Расчет!ZZ2", vbBinaryCompare) > 0
        TestTorsionAssert "Valid torque is not blamed " & CStr(row), InStr(1, reason, torqueAddress, vbBinaryCompare) = 0 And InStr(1, reason, InputDiagnosticValueLine(10#), vbBinaryCompare) = 0
        TestTorsionAssert "Invalid setting value retained " & CStr(row), InStr(1, reason, InputDiagnosticValueLine("29+"), vbBinaryCompare) > 0
        TestTorsionAssert "Single action block " & CStr(row), UBound(Split(reason, "Возможное решение:")) = 1
    Next row
    torques.Cells(3, 2).Value2 = 0#: torques.Cells(2, 2).Value2 = "bad T"
    Set snapshot = New CCalculationSnapshot
    warning = CaptureTorsionPostprocess(ThisWorkbook, TestTorsionRectangle(), batch, units, snapshot)
    data = snapshot.Table(TORSION_POST_DATA_KEY): reason = CStr(data(1, 39))
    TestTorsionAssert "Invalid torque keeps its address", InStr(1, reason, torqueAddress, vbBinaryCompare) > 0
    TestTorsionAssert "Invalid torque does not blame settings", InStr(1, reason, "Torsion.HoopAxisOffset", vbBinaryCompare) = 0
    torques.Cells(2, 1).Value2 = "TestMissing": torques.Cells(2, 2).Value2 = 10#
    Set snapshot = New CCalculationSnapshot
    warning = CaptureTorsionPostprocess(ThisWorkbook, TestTorsionRectangle(), batch, units, snapshot)
    data = snapshot.Table(TORSION_POST_DATA_KEY): reason = CStr(data(1, 39))
    TestTorsionAssert "Missing ID keeps its address", InStr(1, reason, identifierAddress, vbBinaryCompare) > 0
    TestTorsionAssert "Missing ID does not blame torque", InStr(1, reason, torqueAddress, vbBinaryCompare) = 0
    torques.Cells(2, 2).Value2 = 0#
    Set snapshot = New CCalculationSnapshot
    warning = CaptureTorsionPostprocess(ThisWorkbook, TestTorsionRectangle(), batch, units, snapshot)
    data = snapshot.Table(TORSION_POST_DATA_KEY)
    TestTorsionAssert "Zero torque does not read invalid setting", Len(warning) = 0 And Len(CStr(data(1, 39))) = 0
    GoTo Finished
Failed:
    mTestFailed = mTestFailed + 1: mTestLog = mTestLog & vbCrLf & "RUNTIME: " & Err.Description
Finished:
    On Error Resume Next
    If IsArray(originalLoads) Then loads.Formula = originalLoads
    If IsArray(originalTorques) Then torques.Formula = originalTorques
    If Not axis Is Nothing Then axis.Formula = originalAxis
    If Not linked Is Nothing Then linked.Formula = originalLinked
    On Error GoTo 0
    RunTorsionPostDiagnosticTests = "TorsionPostDiagnostics: Passed=" & CStr(mTestPassed) & "; Failed=" & CStr(mTestFailed) & mTestLog
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Независимая размерностная проверка: подобные фермы, поворот/перенос,
' круг с точными дугами, сумма длин и тяги без задвоения угловых стержней.
Public Function RunTorsionGeometryInvariantTests() As String
    mTestPassed = 0: mTestFailed = 0: mTestLog = vbNullString
    On Error GoTo Failed
    Dim sizeFactor As Variant, angle As Variant, shift As Variant, sg As Long, p As TTorsionPostParameters
    Dim section As CSectionModel, calculator As CTorsionPostCalculator, result As CTorsionPostResult
    Dim data As Variant, bars As Variant, i As Long, lengthSum As Double, forceSum As Double, pi As Double
    pi = 4# * Atn(1#)
    For Each sizeFactor In Array(0.1, 1#, 10#)
        For Each angle In Array(0#, pi / 2#, 0.3)
            For Each shift In Array(0#, 1000000#)
                Set section = TestTorsionTransformedRectangle(CDbl(sizeFactor), CDbl(angle), CDbl(shift))
                p.AxisOffset = 20# * sizeFactor: p.HoopDiameter = 8# * sizeFactor: p.HoopSpacing = 150# * sizeFactor
                p.BarBand = 25# * sizeFactor: p.SteelResistance = 270#: p.HoopResistance = 170#
                p.StrutAngle = 45#: p.ConcreteMethod = "None"
                Set calculator = New CTorsionPostCalculator: calculator.Prepare section, p
                For sg = -1 To 1 Step 2
                    Set result = calculator.Calculate("INV", sg * 3900000# * sizeFactor ^ 3, Nothing)
                    data = TestTorsionSummary(result): bars = TestTorsionBars(result)
                    TestTorsionNear "Similar Ao", data(1, 4), 41600# * sizeFactor ^ 2
                    TestTorsionNear "Similar perimeter", data(1, 5), 840# * sizeFactor
                    TestTorsionNear "Similar stress", data(1, 12), 98.4375
                    TestTorsionAssert "Each physical bar once", UBound(bars, 1) = 4
                    lengthSum = 0#: forceSum = 0#
                    For i = 1 To UBound(bars, 1)
                        lengthSum = lengthSum + CDbl(bars(i, 5)): forceSum = forceSum + CDbl(bars(i, 4))
                        TestTorsionNear "Similar assigned length", bars(i, 5), 210# * sizeFactor
                        TestTorsionNear "Similar assigned force", bars(i, 4), 9843.75 * sizeFactor ^ 2
                        TestTorsionAssert "ID kept after transform", CStr(bars(i, 2)) = "R" & i
                    Next i
                    TestTorsionNear "Full perimeter no overlap", lengthSum, CDbl(data(1, 5))
                    TestTorsionNear "Full tie force no overlap", forceSum, CDbl(data(1, 8))
                Next sg
            Next shift
        Next angle
    Next sizeFactor
    Set section = New CSectionModel
    section.AddConcreteElement 0#, 0#, pi * 10000#, shapeType:="Circle", width:=200#, height:=200#
    section.Contours.AddContourArc "C1", 100#, 0#, -100#, 0#, 0#, 0#, 100#, pi
    section.Contours.AddContourArc "C2", -100#, 0#, 100#, 0#, 0#, 0#, 100#, pi
    For i = 0 To 3
        section.AddRebarElement 65# * Cos(i * pi / 2#), 65# * Sin(i * pi / 2#), Sqr(400# / pi), 100#, "A400", comment:="Окружность"
    Next i
    p.AxisOffset = 20#: p.HoopDiameter = 8#: p.HoopSpacing = 150#: p.BarBand = 25#
    Set calculator = New CTorsionPostCalculator: calculator.Prepare section, p
    Set result = calculator.Calculate("CIRCLE", 3900000#, Nothing): data = TestTorsionSummary(result): bars = TestTorsionBars(result)
    TestTorsionNear "Circle exact Ao", data(1, 4), pi * 6400#
    TestTorsionNear "Circle exact perimeter", data(1, 5), pi * 160#
    lengthSum = 0#: forceSum = 0#
    For i = 1 To UBound(bars, 1)
        TestTorsionNear "Circle quarter arc", bars(i, 5), pi * 40#
        lengthSum = lengthSum + CDbl(bars(i, 5)): forceSum = forceSum + CDbl(bars(i, 4))
    Next i
    TestTorsionNear "Circle length conservation", lengthSum, pi * 160#
    TestTorsionNear "Circle force conservation", forceSum, 3900000# / 80#
    GoTo Finished
Failed:
    mTestFailed = mTestFailed + 1: mTestLog = mTestLog & vbCrLf & "RUNTIME: " & Err.Description
Finished:
    RunTorsionGeometryInvariantTests = "TorsionGeometryInvariant: Passed=" & mTestPassed & "; Failed=" & mTestFailed & mTestLog
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Подобный прямоугольник с фактическим контуром; ожидаемые числа выводятся аналитически.
Private Function TestTorsionTransformedRectangle(ByVal sizeFactor As Double, ByVal angle As Double, ByVal shift As Double) As CSectionModel
    Dim section As New CSectionModel, i As Long, j As Long, x As Double, y As Double
    Dim xx As Double, yy As Double, xNext As Double, yNext As Double, corners As Variant, point As Variant, nextPoint As Variant
    corners = Array(Array(-100#, -150#), Array(100#, -150#), Array(100#, 150#), Array(-100#, 150#))
    section.AddConcreteElement shift, -shift, 60000# * sizeFactor ^ 2, shapeType:="Rectangle", width:=200# * sizeFactor, height:=300# * sizeFactor, rotation:=angle
    For i = 0 To 3
        point = corners(i): nextPoint = corners((i + 1) Mod 4)
        xx = shift + sizeFactor * (point(0) * Cos(angle) - point(1) * Sin(angle))
        yy = -shift + sizeFactor * (point(0) * Sin(angle) + point(1) * Cos(angle))
        xNext = shift + sizeFactor * (nextPoint(0) * Cos(angle) - nextPoint(1) * Sin(angle))
        yNext = -shift + sizeFactor * (nextPoint(0) * Sin(angle) + nextPoint(1) * Cos(angle))
        section.Contours.AddContourLine CStr(i), xx, yy, xNext, yNext
    Next i
    For i = -1 To 1 Step 2
        For j = -1 To 1 Step 2
            x = 65# * i: y = 115# * j
            section.AddRebarElement shift + sizeFactor * (x * Cos(angle) - y * Sin(angle)), _
                -shift + sizeFactor * (x * Sin(angle) + y * Cos(angle)), Sqr(400# / (4# * Atn(1#))) * sizeFactor, _
                100# * sizeFactor ^ 2, "A400", comment:="Грань прямоугольника"
        Next j
    Next i
    Set TestTorsionTransformedRectangle = section
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет прежний табличный IO-контракт через действующий адаптер вывода.
Private Function TestTorsionSummary(ByVal result As CTorsionPostResult) As Variant
    Dim writer As New CTorsionSummaryWriter
    TestTorsionSummary = writer.CreateSummaryData(result)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет владение данными, неизменяемость, независимую адресную копию и Empty/ноль в IO.
Private Sub TestTypedTorsionResult()
    Dim data As TTorsionPostData, policy As New CResultStatusPolicy
    data.HasTResults = True: data.HasCombinedResults = True
    data.CombinationID = "Typed": data.StateName = "StrengthState": data.Torque = 3900000#
    data.ShearFlow = 12.5: data.NormalStress = 0#: data.CombinedStress = -12#
    data.LongitudinalStatus = policy.OK: data.HoopStatus = policy.OK
    data.ConcreteMethod = "None": data.ConcreteStatus = policy.NotApplicable
    data.CombinedStatus = policy.OK: data.OverallStatus = policy.OK
    Dim identifiers(1 To 1) As String, placements(1 To 1) As String
    Dim forces(1 To 1) As Double, lengths(1 To 1) As Double
    identifiers(1) = "Rtyped": placements(1) = "B - верхняя внешняя грань, ряд 1"
    forces(1) = 123.456789012345: lengths(1) = 205#
    Dim result As New CTorsionPostResult, writer As New CTorsionSummaryWriter
    result.Initialize data, 1, identifiers, placements, forces, lengths, "Причина NM"
    data.Torque = -1#: identifiers(1) = "Changed": forces(1) = -1#
    TestTorsionAssert "Typed input values are copied", result.Torque = 3900000#
    TestTorsionAssert "Typed bar identifiers are copied", result.BarID(1) = "Rtyped"
    TestTorsionAssert "Typed bar forces are copied", result.BarForce(1) = 123.456789012345
    TestTorsionAssert "Typed bar placement retained", result.BarPlacement(1) = placements(1)
    TestTorsionAssert "Typed bar length retained", result.BarLength(1) = 205#
    Dim summary As Variant, bars As Variant
    summary = writer.CreateSummaryData(result): bars = writer.CreateBarData(result)
    TestTorsionAssert "Typed computed zero stays Double", VarType(summary(1, 30)) = vbDouble And CDbl(summary(1, 30)) = 0#
    TestTorsionAssert "Typed unbounded reserve stays Empty", IsEmpty(summary(1, 35))
    TestTorsionAssert "Typed absent concrete stress stays Empty", IsEmpty(summary(1, 24))
    TestTorsionAssert "Typed absent comment stays Empty", IsEmpty(summary(1, 39))
    bars(1, 4) = -1#
    TestTorsionAssert "Typed exported matrix cannot change result", result.BarForce(1) = 123.456789012345
    Dim addressed As CTorsionPostResult
    Set addressed = result.WithResultComment("Адресная диагностика")
    TestTorsionAssert "Typed addressed comment is independent", addressed.ResultComment = "Адресная диагностика" And Not result.HasResultComment
    TestTorsionAssert "Typed state failure reason retained", addressed.StateFailureReason = "Причина NM"
    TestTorsionAssert "Typed addressed bars retained", addressed.BarCount = 1 And addressed.BarForce(1) = result.BarForce(1)
    Dim copyData As Variant, column As Long
    copyData = writer.CreateSummaryData(addressed)
    For column = 1 To TORSION_POST_COLUMNS
        If column <> 39 Then TestTorsionAssert "Typed copy field " & CStr(column), _
            VarType(copyData(1, column)) = VarType(summary(1, column)) And CStr(copyData(1, column)) = CStr(summary(1, column))
    Next column
    Dim number As Long, ignored As Double
    On Error Resume Next
    result.InitializeFailure "Overwrite", "Недопустимая запись"
    number = Err.Number: Err.Clear
    On Error GoTo 0
    TestTorsionAssert "Typed immutable result rejects write", number = TORSION_RESULT_CONTRACT_ERROR
    TestTorsionAssert "Typed rejected write preserves values", result.CombinationID = "Typed" And result.Torque = 3900000#
    On Error Resume Next
    ignored = result.BarForce(0): number = Err.Number: Err.Clear
    On Error GoTo 0
    TestTorsionAssert "Typed invalid bar index has contract error", number = TORSION_RESULT_CONTRACT_ERROR
    TestTorsionAssert "Typed index error preserves result", result.BarForce(1) = 123.456789012345
    Dim failed As New CTorsionPostResult
    failed.InitializeFailure "Failure", "", True
    summary = writer.CreateSummaryData(failed)
    TestTorsionAssert "Typed failure has no numeric torque", IsEmpty(summary(1, 3))
    TestTorsionAssert "Typed explicit empty comment stays String", VarType(summary(1, 39)) = vbString And Len(CStr(summary(1, 39))) = 0
    TestTorsionAssert "Typed failure remains CalcErr", failed.OverallStatus = policy.CalcErr
    TestTorsionAssert "Typed failure has no bar matrix", IsEmpty(writer.CreateBarData(failed))
    Set addressed = failed.WithResultComment("Адрес ошибки")
    TestTorsionAssert "Typed failure copy retains no bars", addressed.BarCount = 0 And addressed.OverallStatus = policy.CalcErr
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Получает подробные строки только через сериализацию существующего writer.
Private Function TestTorsionBars(ByVal result As CTorsionPostResult) As Variant
    Dim writer As New CTorsionSummaryWriter
    TestTorsionBars = writer.CreateBarData(result)
End Function
