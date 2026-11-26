Attribute VB_Name = "modTestTorsionConcrete"
Option Explicit

' ==================== ДЛЯ ТЕСТОВ ====================
' Независимые числа эффективной стенки и сопротивления бетона при кручении.
' Пример ACI 445.1R-12, §9.4.3: 300x500, c_l=50, f_ck=20, gamma_c=1,5.
' Только синтетическая геометрия, без записи листов и без изменения реальных исходных данных.
Private mTestPassed As Long, mTestFailed As Long, mTestLog As String

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет формулы источника, круглые/полые контуры, поворот и независимые отказы.
Public Function RunTorsionConcreteSourceTests() As String
    mTestPassed = 0: mTestFailed = 0: mTestLog = vbNullString
    On Error GoTo Failed
    Dim query As New CSectionGeometryQuery, domain As CConcreteRegion, check As CTorsionConcreteCheck
    Dim p As TTorsionPostParameters, i As Long, expected As Double
    p.ConcreteMethod = "Strut": p.CharacteristicConcreteStrength = 20#: p.ConcretePartialFactor = 1.5
    p.LongitudinalAxisOffset = 50#: p.StrutAngle = 35#
    For i = 0 To 3
        Set domain = query.RectangleRegion(1234#, -567#, 300#, 500#, i * GEOM_PI / 7#)
        Set check = New CTorsionConcreteCheck: check.Prepare domain, p
        TestConcreteNear "Source thickness rotated " & i, check.Thickness, 100#
        TestConcreteNear "Source area rotated " & i, check.EffectiveArea, 80000#
        TestConcreteNear "Source design strength " & i, check.DesignStrength, 20# / 1.5
        TestConcreteNear "Source efficiency " & i, check.Efficiency, 0.552
        TestConcreteNear "Source strut resistance " & i, check.Resistance, 7.36
        expected = 58880000# * Sin(70# * GEOM_PI / 180#)
        TestConcreteNear "Source limit unrounded fcd " & i, check.Limit, expected
        TestConcreteNear "Source stress " & i, check.Stress(30000000#), 30000000# / (8000000# * Sin(70# * GEOM_PI / 180#))
        TestConcreteNear "Stress and torque reserves equal " & i, check.Resistance / check.Stress(30000000#), check.Limit / 30000000#
        TestConcreteNear "Torque sign " & i, check.Stress(-30000000#), check.Stress(30000000#)
    Next i
    Set domain = query.RectangleRegion(0#, 0#, 300#, 500#)
    p.ConcretePartialFactor = 20# / 13.3 ' В печатном численном примере f_cd округлено до 13,3 МПа.
    Set check = New CTorsionConcreteCheck: check.Prepare domain, p
    TestConcreteNear "Printed example rounded result 55.2 kNm", Round(check.Limit / 1000000#, 1), 55.2
    p.ConcretePartialFactor = 1.5: p.StrutAngle = 45#
    Set check = New CTorsionConcreteCheck: check.Prepare domain, p
    TestConcreteNear "45-degree limit", check.Limit, 58880000#
    p.CharacteristicConcreteStrength = 30#
    Set check = New CTorsionConcreteCheck: check.Prepare domain, p
    TestConcreteNear "fck30 efficiency", check.Efficiency, 0.528
    TestConcreteNear "fck30 fcd", check.DesignStrength, 20#
    TestConcreteNear "fck30 resistance", check.Resistance, 10.56
    p.StrutAngle = Atn(1# / 2.5) * 180# / GEOM_PI
    Set check = New CTorsionConcreteCheck: check.Prepare domain, p
    TestConcreteNear "Minimum admissible angle", check.Limit, 2# * 80000# * 100# * 10.56 / 2.9
    p.StrutAngle = 45#: p.CharacteristicConcreteStrength = 20#: p.LongitudinalAxisOffset = 20#
    Set domain = query.CircleRegion(0#, 0#, 100#)
    Set check = New CTorsionConcreteCheck: check.Prepare domain, p
    TestConcreteNear "Circle auto A/u", check.Thickness, 50#
    TestConcreteNear "Circle effective area analytic", check.EffectiveArea, GEOM_PI * 75# * 75#
    Dim outer As CConcreteRegion, opening As CConcreteRegion
    Set outer = query.RectangleRegion(0#, 0#, 300#, 500#)
    Set opening = query.RectangleRegion(0#, 0#, 200#, 400#)
    Set domain = query.SubtractContainedRegion(outer, opening)
    Set check = New CTorsionConcreteCheck: check.Prepare domain, p
    TestConcreteNear "Hollow actual wall upper limit", check.Thickness, 50#
    TestConcreteNear "Hollow area includes void", check.EffectiveArea, 112500#
    p.WallThickness = 45#
    Set check = New CTorsionConcreteCheck: check.Prepare domain, p
    TestConcreteNear "Explicit permitted hollow wall", check.Thickness, 45#
    TestConcreteNear "Explicit hollow effective area", check.EffectiveArea, 255# * 455#
    p.WallThickness = 51#: TestConcreteFailure domain, p, "Exceeds real wall", "фактическую"
    p.WallThickness = 39#: TestConcreteFailure domain, p, "Below twice cl", "удвоенного"
    p.WallThickness = 0#: p.LongitudinalAxisOffset = 30#: TestConcreteFailure domain, p, "No feasible hollow wall", "удвоенное"
    p.LongitudinalAxisOffset = 20#: p.StrutAngle = 60#: TestConcreteFailure domain, p, "Angle above source interval", "cot(theta)"
    p.StrutAngle = 20#: TestConcreteFailure domain, p, "Angle below source interval", "cot(theta)"
    p.StrutAngle = 45#: p.ConcretePartialFactor = 0.9: TestConcreteFailure domain, p, "gamma below one", "gamma_c"
    p.ConcretePartialFactor = 1.5: p.CharacteristicConcreteStrength = 250#: TestConcreteFailure domain, p, "Nonpositive nu", "f_ck"
    p.CharacteristicConcreteStrength = 20#: p.LongitudinalAxisOffset = 0#: TestConcreteFailure domain, p, "Zero cl", "Расстояние"
    p.LongitudinalAxisOffset = 50#: p.ConcreteMethod = "GorskyRectangle": p.ConcreteResistance = 10#: p.GorskyCoefficient = 0.07
    p.CharacteristicConcreteStrength = -1#: p.ConcretePartialFactor = -1#: p.WallThickness = -1#: p.StrutAngle = -1#
    Set domain = query.RectangleRegion(0#, 0#, 300#, 500#, GEOM_PI / 9#)
    Set check = New CTorsionConcreteCheck: check.Prepare domain, p
    TestConcreteNear "Gorsky ignores unused Strut fields", check.Limit, 31500000#
    p.ConcreteMethod = "None": Set check = New CTorsionConcreteCheck: check.Prepare Nothing, p
    TestConcreteNear "None reads no concrete fields", check.Limit, 0#
    p.ConcreteMethod = "Strut": p.CharacteristicConcreteStrength = 20#: p.ConcretePartialFactor = 1.5
    p.LongitudinalAxisOffset = 20#: p.WallThickness = 0#: p.StrutAngle = 45#
    Dim xy(1 To 6, 1 To 2) As Double
    xy(1, 1) = 0#: xy(1, 2) = 0#: xy(2, 1) = 600#: xy(2, 2) = 0#
    xy(3, 1) = 600#: xy(3, 2) = 200#: xy(4, 1) = 200#: xy(4, 2) = 200#
    xy(5, 1) = 200#: xy(5, 2) = 700#: xy(6, 1) = 0#: xy(6, 2) = 700#
    Set domain = query.PolygonRegion(xy)
    Set check = New CTorsionConcreteCheck: check.Prepare domain, p
    TestConcreteNear "L-section effective thickness", check.Thickness, 220000# / 2600#
    expected = 220000# - 2600# * (check.Thickness / 2#) + check.Thickness * check.Thickness
    TestConcreteNear "L-section exact inward area", check.EffectiveArea, expected
    GoTo Finished
Failed:
    mTestFailed = mTestFailed + 1: mTestLog = mTestLog & vbCrLf & "RUNTIME: " & Err.Description
Finished:
    RunTorsionConcreteSourceTests = "TorsionConcreteSource: Passed=" & mTestPassed & "; Failed=" & mTestFailed & mTestLog
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет чтение конкретного формата 39/43 колонок без фиктивных параметров EC2.
' Запускается только в собственной книге: меняет вывод T на Results.
Public Function RunTorsionConcreteLegacyTests() As String
    mTestPassed = 0: mTestFailed = 0: mTestLog = vbNullString
    On Error GoTo Failed
    Dim settings As New CSystemSettingsReader, units As New CUnitSystem, writer As New CTorsionSummaryWriter
    settings.LoadFromWorkbook ThisWorkbook: units.LoadFromSettings settings
    Dim oldData(1 To 1, 1 To 39) As Variant, oldHeaders(1 To 4, 1 To 43) As Variant, oldBars(1 To 2, 1 To 3) As Variant
    oldData(1, 1) = "Legacy01": oldData(1, 6) = 1#: oldData(1, 7) = 2#: oldData(1, 22) = "Strut"
    oldData(1, 24) = 3.5: oldData(1, 25) = 5#: oldData(1, 26) = 10#: oldData(1, 27) = 2#: oldData(1, 28) = "OK"
    oldData(1, 38) = "OK": oldData(1, 36) = "N/A"
    Dim current As Variant: current = writer.SummaryHeaders(units)
    Dim r As Long, c As Long, sourceColumn As Long
    For c = 1 To 43
        sourceColumn = c
        If c >= 19 And c <= 30 Then sourceColumn = c + 1
        If c >= 31 And c <= 36 Then sourceColumn = c + 4
        If c >= 37 Then sourceColumn = c + 5
        For r = 1 To 4: oldHeaders(r, c) = current(r, sourceColumn): Next r
    Next c
    oldBars(1, 1) = "LoadCase": oldBars(1, 2) = "Rebar": oldBars(1, 3) = "F_T, " & units.OutputForceUnit
    oldBars(2, 1) = "Legacy01": oldBars(2, 2) = "R1": oldBars(2, 3) = 420#
    Dim snapshot As New CCalculationSnapshot
    snapshot.AddTable TORSION_POST_HEADERS_KEY, oldHeaders: snapshot.AddTable TORSION_POST_DATA_KEY, oldData
    snapshot.AddTable TORSION_BAR_DATA_KEY, oldBars
    Dim warning As String: warning = writer.RestoreSummary(ThisWorkbook, snapshot)
    TestConcreteAssert "Legacy table restored", Len(warning) = 0
    Dim anchor As Object: Set anchor = ThisWorkbook.Names("rngTorsionSummaryAnchor").RefersToRange
    TestConcreteNear "Legacy stress retained", anchor.Offset(0, 34).Value2, 3.5
    TestConcreteNear "Legacy resistance retained", anchor.Offset(0, 35).Value2, 5#
    TestConcreteAssert "Legacy area not fabricated", IsEmpty(anchor.Offset(0, 31).Value2)
    TestConcreteAssert "Legacy fcd not fabricated", IsEmpty(anchor.Offset(0, 32).Value2)
    TestConcreteAssert "Legacy nu not fabricated", IsEmpty(anchor.Offset(0, 33).Value2)
    TestConcreteAssert "Legacy model identified", InStr(1, CStr(anchor.Offset(0, 1).Value2), "39 колонок", vbBinaryCompare) > 0
    TestConcreteAssert "Legacy stress unit after last comma", CStr(anchor.Offset(-1, 32).Value2) = "f_cd, " & units.OutputStressUnit
    TestConcreteNear "Legacy bar assigned length", ThisWorkbook.Names("rngTorsionBarForces").RefersToRange.Offset(1, 4).Value2, 210#
    GoTo Finished
Failed:
    mTestFailed = mTestFailed + 1: mTestLog = mTestLog & vbCrLf & "RUNTIME: " & Err.Description
Finished:
    RunTorsionConcreteLegacyTests = "TorsionConcreteLegacy: Passed=" & mTestPassed & "; Failed=" & mTestFailed & mTestLog
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет отказ по фактической причине и ключу активного исходного параметра.
Private Sub TestConcreteFailure(ByVal domain As CConcreteRegion, ByRef p As TTorsionPostParameters, ByVal label As String, ByVal reason As String)
    Dim check As New CTorsionConcreteCheck, message As String, number As Long
    On Error Resume Next
    check.Prepare domain, p: number = Err.Number: message = Err.Description
    On Error GoTo 0
    TestConcreteAssert label & " InputErr", number = vbObjectError + 5670
    TestConcreteAssert label & " clear reason", InStr(1, message, reason, vbTextCompare) > 0
    TestConcreteAssert label & " addressable setting", InStr(1, check.FailureSettingKeys, "Torsion.", vbBinaryCompare) > 0
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает с независимым аналитическим числом с относительным допуском 1e-9.
Private Sub TestConcreteNear(ByVal label As String, ByVal actual As Double, ByVal expected As Double)
    TestConcreteAssert label, Abs(actual - expected) <= 0.000000001 * GeomMax(Abs(expected), 0.000001)
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сохраняет каждое ожидание в первичном журнале, не заменяя ошибки общим счётчиком.
Private Sub TestConcreteAssert(ByVal label As String, ByVal passed As Boolean)
    If passed Then mTestPassed = mTestPassed + 1 Else mTestFailed = mTestFailed + 1
    mTestLog = mTestLog & vbCrLf & IIf(passed, "PASS ", "FAIL ") & label
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Выполняет один штатный расчёт до и после задания T в своей копии, для Excel/File.
' Сравнивает напряжения НДМ и проверяет фактическую запись новых параметров бетонного блока.
Public Function RunTorsionConcreteWorkbookTests(ByVal storageMode As String) As String
    mTestPassed = 0: mTestFailed = 0: mTestLog = vbNullString
    On Error GoTo Failed
    TestConcreteSet ThisWorkbook.Names("rngUnitSettings").RefersToRange, "Length", "mm"
    TestConcreteSet ThisWorkbook.Names("rngUnitSettings").RefersToRange, "Stress", "MPa"
    TestConcreteSet ThisWorkbook.Names("rngUnitSettings").RefersToRange, "Moment", "N*mm"
    modTestPerformance.TestConfigurePerformanceFixture "STRENGTH", 1, "Newton", "Circle", "Normal", "Yes"
    Dim circleRange As Object: Set circleRange = ThisWorkbook.Names("rngCircleGeometry").RefersToRange
    Dim inputTable As New CSectionInputTable, keyed As New CSectionInputReader
    inputTable.Initialize "rngCircleGeometry", circleRange.Value2: keyed.Initialize inputTable
    circleRange.Cells(keyed.RowIndex("Main", "Circle"), keyed.ColumnIndex("Main", "n")).Value2 = 12
    circleRange.Cells(keyed.RowIndex("Main", "Circle"), keyed.ColumnIndex("Main", "as")).Value2 = 50#
    circleRange.Cells(keyed.RowIndex("Main", "Circle"), keyed.ColumnIndex("Main", "d")).Value2 = 20#
    Dim system As Object: Set system = ThisWorkbook.Names("rngSystemSettings").RefersToRange
    TestConcreteSet system, "Results.SnapshotStorage", storageMode
    TestConcreteSet system, "Results.SnapshotName", "TestConcreteSource_" & storageMode
    TestConcreteSet system, "General.ExecutionReportEnabled", "No"
    TestConcreteSet system, "Plot.Enabled", "No"
    TestConcreteSet system, "Output.CombinationMode", "User"
    TestConcreteSet system, "Output.CombinationID", CStr(ThisWorkbook.Names("rngLoadCombinations").RefersToRange.Cells(2, 1).Value2)
    Dim settings As Object: Set settings = ThisWorkbook.Names("rngTorsionSettings").RefersToRange
    TestConcreteSet settings, "Torsion.ConcreteMethod", "Strut"
    TestConcreteSet settings, "Torsion.StateType", "StrengthState"
    TestConcreteSet settings, "Torsion.Fck", 20#
    TestConcreteSet settings, "Torsion.GammaC", 1.5
    TestConcreteSet settings, "Torsion.RebarAxisOffset", 50#
    TestConcreteSet settings, "Torsion.WallThickness", 0#
    TestConcreteSet settings, "Torsion.StrutAngle", 45#
    Dim torsion As Object: Set torsion = ThisWorkbook.Names("rngTorsionCombinations").RefersToRange
    torsion.Offset(1, 0).Resize(torsion.Rows.Count - 1, 2).ClearContents
    torsion.Cells(2, 1).Value2 = ThisWorkbook.Names("rngLoadCombinations").RefersToRange.Cells(2, 1).Value2
    torsion.Cells(2, 2).Value2 = 0#
    Dim message As String: message = FormSectionForWorkbook(ThisWorkbook)
    mTestLog = mTestLog & vbCrLf & "INFO form: " & message
    message = RunSectionCalculationForWorkbook(ThisWorkbook)
    mTestLog = mTestLog & vbCrLf & "INFO baseline: " & message
    Dim before As Variant: before = ReadAnchoredResultTable(ThisWorkbook, "rngNDMElementResults")
    TestConcreteAssert storageMode & " normal NM produced", UBound(before, 1) > 1
    torsion.Cells(2, 2).Value2 = 30000000#
    message = RunSectionCalculationForWorkbook(ThisWorkbook)
    mTestLog = mTestLog & vbCrLf & "INFO with T: " & message
    Dim after As Variant: after = ReadAnchoredResultTable(ThisWorkbook, "rngNDMElementResults")
    TestConcreteAssert storageMode & " NM row count unchanged", UBound(before, 1) = UBound(after, 1)
    Dim i As Long, same As Boolean: same = True
    For i = 2 To UBound(before, 1)
        If CStr(before(i, 5)) <> CStr(after(i, 5)) Or CStr(before(i, 6)) <> CStr(after(i, 6)) Or _
                CDbl(before(i, 7)) <> CDbl(after(i, 7)) Or CDbl(before(i, 8)) <> CDbl(after(i, 8)) Or _
                CStr(before(i, 9)) <> CStr(after(i, 9)) Then same = False
    Next i
    TestConcreteAssert storageMode & " exact NM strains and stresses unchanged", same
    Dim reader As New CSystemSettingsReader, units As New CUnitSystem
    reader.LoadFromWorkbook ThisWorkbook: units.LoadFromSettings reader
    Dim anchor As Object: Set anchor = ThisWorkbook.Names("rngTorsionSummaryAnchor").RefersToRange
    TestConcreteNear storageMode & " normal workflow fcd", anchor.Offset(0, 32).Value2, units.InternalStressToOutput(20# / 1.5)
    TestConcreteNear storageMode & " normal workflow nu", anchor.Offset(0, 33).Value2, 0.552
    TestConcreteAssert storageMode & " normal workflow Ak", CDbl(anchor.Offset(0, 31).Value2) > 0#
    TestConcreteAssert storageMode & " concrete status available", CStr(anchor.Offset(0, 38).Value2) = "OK" Or CStr(anchor.Offset(0, 38).Value2) = "FAIL"
    TestConcreteAssert storageMode & " NMT saved state available", CStr(anchor.Offset(0, 47).Value2) = "OK" Or CStr(anchor.Offset(0, 47).Value2) = "FAIL"
    Dim geometry As Variant, barMatrix As Variant, dict As Object, ownNames As Boolean
    geometry = ReadAnchoredResultTable(ThisWorkbook, "rngNDMSectionGeometry")
    Set dict = CreateObject("Scripting.Dictionary")
    For i = 2 To UBound(geometry, 1): dict(CStr(geometry(i, 2))) = CStr(geometry(i, 15)): Next i
    TestConcreteAssert storageMode & " T worst own face", Len(CStr(anchor.Offset(0, 18).Value2)) > 0 And CStr(anchor.Offset(0, 18).Value2) = CStr(dict(CStr(anchor.Offset(0, 17).Value2)))
    TestConcreteAssert storageMode & " NMT worst own face", Len(CStr(anchor.Offset(0, 40).Value2)) > 0 And CStr(anchor.Offset(0, 40).Value2) = CStr(dict(CStr(anchor.Offset(0, 39).Value2)))
    Dim store As New CCalculationSnapshotStore, saved As CCalculationSnapshot, matrix As Variant
    If storageMode = "File" Then
        Set saved = store.LoadFile(ThisWorkbook): barMatrix = saved.Table(TORSION_BAR_DATA_KEY)
    Else
        barMatrix = ReadPhysicalResultTable(ThisWorkbook, "rngTorsionBarForces")
    End If
    ownNames = True
    For i = 2 To UBound(barMatrix, 1)
        If Len(CStr(barMatrix(i, 3))) = 0 Or CStr(barMatrix(i, 3)) <> CStr(dict(CStr(barMatrix(i, 2)))) Then ownNames = False
    Next i
    TestConcreteAssert storageMode & " every bar owns its saved geometry face", ownNames And UBound(barMatrix, 1) = 13
    If storageMode = "File" Then
        Set saved = store.LoadFile(ThisWorkbook): matrix = saved.Table(TORSION_POST_DATA_KEY)
        TestConcreteAssert "File stores 44 columns", UBound(matrix, 2) = TORSION_POST_COLUMNS
        TestConcreteNear "File exact fcd", matrix(1, 41), CDbl(anchor.Offset(0, 32).Value2)
        TestConcreteNear "File exact nu", matrix(1, 42), 0.552
        Dim barAnchor As Object: Set barAnchor = ThisWorkbook.Names("rngTorsionBarForces").RefersToRange
        Dim column As Long, sheetBars As Variant
        sheetBars = barAnchor.Resize(UBound(barMatrix, 1), TORSION_BAR_COLUMNS).Value2
        For i = 1 To UBound(barMatrix, 1): For column = 1 To TORSION_BAR_COLUMNS
            TestConcreteAssert "File calculated bars exact " & i & ":" & column, CStr(sheetBars(i, column)) = CStr(barMatrix(i, column))
        Next column: Next i
        barAnchor.Resize(UBound(barMatrix, 1) + 2, TORSION_BAR_COLUMNS).Value2 = "Test stale bars"
        Dim savedBarAddress As String: savedBarAddress = ThisWorkbook.Names("rngTorsionBarForces").RefersTo
        ThisWorkbook.Names("rngTorsionBarForces").RefersTo = "='" & barAnchor.Worksheet.Name & "'!" & barAnchor.Worksheet.Cells(barAnchor.Worksheet.Rows.Count, barAnchor.Column).Address
        Dim overflowNumber As Long, overflowReason As String
        On Error Resume Next
        UpdateSectionResultsForWorkbook ThisWorkbook
        overflowNumber = Err.Number: overflowReason = Err.Description: Err.Clear
        ThisWorkbook.Names("rngTorsionBarForces").RefersTo = savedBarAddress
        On Error GoTo Failed
        TestConcreteAssert "File bar overflow diagnosed", overflowNumber <> 0 And InStr(overflowReason, "rngTorsionBarForces") > 0
        TestConcreteAssert "File bar overflow before clear", CStr(barAnchor.Value2) = "Test stale bars"
        UpdateSectionResultsForWorkbook ThisWorkbook ' Тот же вход, который использует кнопка обновления: источник только File, без решения НДМ.
        sheetBars = barAnchor.Resize(UBound(barMatrix, 1), TORSION_BAR_COLUMNS).Value2
        For i = 1 To UBound(barMatrix, 1): For column = 1 To TORSION_BAR_COLUMNS
            TestConcreteAssert "File refreshed bars exact " & i & ":" & column, CStr(sheetBars(i, column)) = CStr(barMatrix(i, column))
        Next column: Next i
        TestConcreteAssert "File stale bar tail cleared", IsEmpty(barAnchor.Offset(UBound(barMatrix, 1), 0).Value2)
        TestConcreteAssert "File NM details remain header only", IsEmpty(ThisWorkbook.Names("rngNDMElementResults").RefersToRange.Offset(1, 0).Value2)
    End If
    GoTo Finished
Failed:
    mTestFailed = mTestFailed + 1: mTestLog = mTestLog & vbCrLf & "RUNTIME: " & Err.Description
Finished:
    RunTorsionConcreteWorkbookTests = "TorsionConcreteWorkbook " & storageMode & ": Passed=" & mTestPassed & "; Failed=" & mTestFailed & mTestLog
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Меняет только используемый ключ собственной проверочной книги, без чтения служебных строк.
Private Sub TestConcreteSet(ByVal table As Object, ByVal key As String, ByVal value As Variant)
    Dim row As Long
    For row = 2 To table.Rows.Count
        If CStr(table.Cells(row, 1).Value2) = key Then table.Cells(row, 2).Value2 = value: Exit Sub
    Next row
    Err.Raise 5, , "Не найден тестовый ключ: " & key
End Sub
