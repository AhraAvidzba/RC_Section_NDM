Attribute VB_Name = "modTestGeometrySolidSnapshots"
Option Explicit

' ==========================================================================
' ДЛЯ ТЕСТОВ: точные снимки геометрии и прямого НДС до/после доработки архитектуры.
' ==========================================================================
Private Type TDoubleValue
    Value As Double
End Type
Private Type TDoubleWords
    Low As Long
    High As Long
End Type

' -------------------- ДЛЯ ТЕСТОВ --------------------
' битовое сравнение координат, площадей, контуров и выходных аннотаций десяти форм.
Public Function TestDumpSolidPhysicalBits() As String
    Dim config As Object, saved As Variant, fixture As Variant, result As String
    Dim settings As CSystemSettingsReader, units As CUnitSystem, model As CSectionModel, i As Long
    On Error GoTo Failed
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = TestFixtureSaveInputs(config)
    For Each fixture In TestFixtureNames()
        TestFixtureRestoreInputs config, saved
        modTestShapeArchitecture.TestConfigureShapeFixture CStr(fixture)
        Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
        Set units = New CUnitSystem: units.LoadFromSettings settings
        Set model = BuildWorkbookSectionModel(ThisWorkbook, settings, units)
        result = result & CStr(fixture) & "|COUNTS|" & CStr(model.ConcreteCount) & "|" & CStr(model.RebarCount) & vbCrLf
        For i = 1 To model.ConcreteCount
            result = result & CStr(fixture) & "|" & model.ConcreteID(i) & "|" & TestBits(model.ConcreteX(i)) & "|" & _
                TestBits(model.ConcreteY(i)) & "|" & TestBits(model.ConcreteArea(i)) & "|" & TestBits(model.ConcreteWidth(i)) & "|" & _
                TestBits(model.ConcreteHeight(i)) & "|" & TestBits(model.ConcreteRotation(i)) & "|" & _
                TestBits(model.ConcreteLocalIx(i)) & "|" & TestBits(model.ConcreteLocalIy(i)) & "|" & TestBits(model.ConcreteLocalIxy(i)) & vbCrLf
        Next i
        For i = 1 To model.RebarCount
            result = result & CStr(fixture) & "|" & model.RebarID(i) & "|" & model.RebarSourceName(i) & "|" & _
                TestBits(model.RebarX(i)) & "|" & TestBits(model.RebarY(i)) & "|" & TestBits(model.RebarDiameter(i)) & "|" & TestBits(model.RebarArea(i)) & vbCrLf
        Next i
        For i = 1 To model.Contours.Count
            With model.Contours
                result = result & CStr(fixture) & "|CONTOUR|" & .ContourID(i) & "|" & .LoopRole(i) & "|" & .LoopID(i) & "|" & _
                    TestBits(.StartX(i)) & "|" & TestBits(.StartY(i)) & "|" & TestBits(.EndX(i)) & "|" & TestBits(.EndY(i)) & "|" & _
                    TestBits(.CenterX(i)) & "|" & TestBits(.CenterY(i)) & "|" & TestBits(.Radius(i)) & "|" & TestBits(.SweepAngle(i)) & vbCrLf
            End With
        Next i
        For i = 1 To model.Annotations.Count
            With model.Annotations
                result = result & CStr(fixture) & "|ANNOTATION|" & .AnnotationID(i) & "|" & .AnnotationType(i) & "|" & .Text(i) & "|" & _
                    TestBits(.StartX(i)) & "|" & TestBits(.StartY(i)) & "|" & TestBits(.EndX(i)) & "|" & TestBits(.EndY(i)) & "|" & _
                    TestBits(.NormalX(i)) & "|" & TestBits(.NormalY(i)) & "|" & TestBits(.Value(i)) & vbCrLf
            End With
        Next i
    Next fixture
    GoTo Finished
Failed:
    result = result & "ERROR: " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Finished:
    If Not config Is Nothing And Not IsEmpty(saved) Then TestFixtureRestoreInputs config, saved
    TestDumpSolidPhysicalBits = result
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' реальные прямые решения трёх назначений материалов при N, N+Mx, N+My и N+Mx+My.
' Выход Double записывается без округления; NumFail и ошибки остаются видимыми в снимке.
Public Function TestDumpSolidDirectNdsBits(Optional ByVal restoreSaved As Boolean = False) As String
    Dim config As Object, saved As Variant, fixture As Variant, purpose As Variant, result As String
    Dim settings As CSystemSettingsReader, units As CUnitSystem, model As CSectionModel
    Dim props As CSectionPropertiesCalculator, provider As CMaterialModelProvider, solver As CSectionSolver
    Dim concrete As Object, steel As Object, n As Double, mx As Double, my As Double, loadIndex As Long
    Dim minX As Double, maxX As Double, minY As Double, maxY As Double
    On Error GoTo Failed
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = TestFixtureSaveInputs(config)
    For Each fixture In TestFixtureNames()
        TestFixtureRestoreInputs config, saved
        modTestShapeArchitecture.TestConfigureShapeFixture CStr(fixture)
        Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
        Set units = New CUnitSystem: units.LoadFromSettings settings
        Set model = BuildWorkbookSectionModel(ThisWorkbook, settings, units)
        If restoreSaved Then
            Dim writer As CNDMResultsWriter: Set writer = New CNDMResultsWriter
            writer.WriteGeometryPreview ThisWorkbook, model, PrepareSectionSnapshot(model), units
            Set model = ReadFormedSectionFromResults(ThisWorkbook)
        End If
        Set props = New CSectionPropertiesCalculator: props.CalculateConcrete model
        props.CalculateProjection model, 1#, 0#, False, minX, maxX
        props.CalculateProjection model, 0#, 1#, False, minY, maxY
        Set provider = New CMaterialModelProvider: provider.Initialize settings, units
        For Each purpose In Array(cpStrength, cpMcrc, cpCrackedNDS)
            Set concrete = provider.ConcreteMaterial(CLng(purpose)): Set steel = provider.SteelMaterial(CLng(purpose))
            For loadIndex = 0 To 3
                n = -2# * props.Area: mx = n * props.CentroidY: my = n * props.CentroidX
                If loadIndex = 1 Or loadIndex = 3 Then mx = mx + props.Area * (maxY - minY) * 0.2
                If loadIndex = 2 Or loadIndex = 3 Then my = my - props.Area * (maxX - minX) * 0.15
                Set solver = New CSectionSolver: solver.ApplySettings settings, units
                solver.Solve model, concrete, steel, n, mx, my
                result = result & CStr(fixture) & "|" & CStr(purpose) & "|" & CStr(loadIndex) & "|" & CStr(solver.Converged) & "|" & _
                    TestBits(solver.Epsilon0) & "|" & TestBits(solver.KappaX) & "|" & TestBits(solver.KappaY) & "|" & _
                    TestBits(solver.Nint) & "|" & TestBits(solver.Mxint) & "|" & TestBits(solver.Myint) & "|" & _
                    TestBits(solver.ResidualN) & "|" & TestBits(solver.ResidualMx) & "|" & TestBits(solver.ResidualMy) & vbCrLf
            Next loadIndex
        Next purpose
    Next fixture
    GoTo Finished
Failed:
    result = result & "ERROR: " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Finished:
    If Not config Is Nothing And Not IsEmpty(saved) Then TestFixtureRestoreInputs config, saved
    TestDumpSolidDirectNdsBits = result
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' те же 120 НДС после сохранения и минимального восстановления.
' Сравнение с DumpSolidDirectNdsBits проверяет физический результат без JSON
' и исходных аннотаций, а не только совпадение записанных координат.
Public Function TestDumpRestoredDirectNdsBits() As String
    TestDumpRestoredDirectNdsBits = TestDumpSolidDirectNdsBits(True)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Перечисляет десять фиксированных вариантов геометрии для побитового сопоставления снимков.
Private Function TestFixtureNames() As Variant
    TestFixtureNames = Array("Circle", "Rectangle", "LSection", "TwoRectangles", "RoundedSimple", "RoundedTapered", _
        "RoundedMixed", "RoundedMixedLeft", "HollowCentered", "HollowOffset")
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сохраняет все биты Double, включая знак нуля; сериализация не зависит от локали Excel.
Private Function TestBits(ByVal value As Double) As String
    Dim number As TDoubleValue, words As TDoubleWords
    number.Value = value: LSet words = number
    TestBits = Right$("00000000" & Hex$(words.High), 8) & Right$("00000000" & Hex$(words.Low), 8)
End Function
