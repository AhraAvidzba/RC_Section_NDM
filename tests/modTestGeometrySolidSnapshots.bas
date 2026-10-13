Attribute VB_Name = "modTestGeometrySolidSnapshots"
Option Explicit

' ДЛЯ ТЕСТОВ: точные снимки геометрии и прямого НДС до/после доработки архитектуры.
Private Type TDoubleValue
    Value As Double
End Type
Private Type TDoubleWords
    Low As Long
    High As Long
End Type

' ДЛЯ ТЕСТОВ: битовое сравнение координат, площадей, контуров и выходных аннотаций десяти форм.
Public Function DumpSolidPhysicalBits() As String
    Dim config As Object, saved As Variant, fixture As Variant, result As String
    Dim settings As CSystemSettingsReader, units As CUnitSystem, model As CSectionModel, i As Long
    On Error GoTo Failed
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = config.Formula
    For Each fixture In FixtureNames()
        config.Formula = saved
        modTestShapeArchitecture.ConfigureShapeFixture CStr(fixture)
        Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
        Set units = New CUnitSystem: units.LoadFromSettings settings
        Set model = BuildWorkbookSectionModel(ThisWorkbook, settings, units)
        result = result & CStr(fixture) & "|COUNTS|" & CStr(model.ConcreteCount) & "|" & CStr(model.RebarCount) & vbCrLf
        For i = 1 To model.ConcreteCount
            result = result & CStr(fixture) & "|" & model.ConcreteID(i) & "|" & Bits(model.ConcreteX(i)) & "|" & _
                Bits(model.ConcreteY(i)) & "|" & Bits(model.ConcreteArea(i)) & "|" & Bits(model.ConcreteWidth(i)) & "|" & _
                Bits(model.ConcreteHeight(i)) & "|" & Bits(model.ConcreteRotation(i)) & "|" & _
                Bits(model.ConcreteLocalIx(i)) & "|" & Bits(model.ConcreteLocalIy(i)) & "|" & Bits(model.ConcreteLocalIxy(i)) & vbCrLf
        Next i
        For i = 1 To model.RebarCount
            result = result & CStr(fixture) & "|" & model.RebarID(i) & "|" & model.RebarSourceName(i) & "|" & _
                Bits(model.RebarX(i)) & "|" & Bits(model.RebarY(i)) & "|" & Bits(model.RebarDiameter(i)) & "|" & Bits(model.RebarArea(i)) & vbCrLf
        Next i
        For i = 1 To model.Contours.Count
            With model.Contours
                result = result & CStr(fixture) & "|CONTOUR|" & .ContourID(i) & "|" & .LoopRole(i) & "|" & .LoopID(i) & "|" & _
                    Bits(.StartX(i)) & "|" & Bits(.StartY(i)) & "|" & Bits(.EndX(i)) & "|" & Bits(.EndY(i)) & "|" & _
                    Bits(.CenterX(i)) & "|" & Bits(.CenterY(i)) & "|" & Bits(.Radius(i)) & "|" & Bits(.SweepAngle(i)) & vbCrLf
            End With
        Next i
        For i = 1 To model.Annotations.Count
            With model.Annotations
                result = result & CStr(fixture) & "|ANNOTATION|" & .AnnotationID(i) & "|" & .AnnotationType(i) & "|" & .Text(i) & "|" & _
                    Bits(.StartX(i)) & "|" & Bits(.StartY(i)) & "|" & Bits(.EndX(i)) & "|" & Bits(.EndY(i)) & "|" & _
                    Bits(.NormalX(i)) & "|" & Bits(.NormalY(i)) & "|" & Bits(.Value(i)) & vbCrLf
            End With
        Next i
    Next fixture
    GoTo Finished
Failed:
    result = result & "ERROR: " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Finished:
    If Not config Is Nothing And Not IsEmpty(saved) Then config.Formula = saved
    DumpSolidPhysicalBits = result
End Function

' ДЛЯ ТЕСТОВ: реальные прямые решения трёх назначений материалов при N, N+Mx, N+My и N+Mx+My.
' Выход Double записывается без округления; NumFail и ошибки остаются видимыми в снимке.
Public Function DumpSolidDirectNdsBits() As String
    Dim config As Object, saved As Variant, fixture As Variant, purpose As Variant, result As String
    Dim settings As CSystemSettingsReader, units As CUnitSystem, model As CSectionModel
    Dim props As CSectionPropertiesCalculator, provider As CMaterialModelProvider, solver As CSectionSolver
    Dim concrete As Object, steel As Object, n As Double, mx As Double, my As Double, loadIndex As Long
    Dim minX As Double, maxX As Double, minY As Double, maxY As Double
    On Error GoTo Failed
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = config.Formula
    For Each fixture In FixtureNames()
        config.Formula = saved
        modTestShapeArchitecture.ConfigureShapeFixture CStr(fixture)
        Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
        Set units = New CUnitSystem: units.LoadFromSettings settings
        Set model = BuildWorkbookSectionModel(ThisWorkbook, settings, units)
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
                    Bits(solver.Epsilon0) & "|" & Bits(solver.KappaX) & "|" & Bits(solver.KappaY) & "|" & _
                    Bits(solver.Nint) & "|" & Bits(solver.Mxint) & "|" & Bits(solver.Myint) & "|" & _
                    Bits(solver.ResidualN) & "|" & Bits(solver.ResidualMx) & "|" & Bits(solver.ResidualMy) & vbCrLf
            Next loadIndex
        Next purpose
    Next fixture
    GoTo Finished
Failed:
    result = result & "ERROR: " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Finished:
    If Not config Is Nothing And Not IsEmpty(saved) Then config.Formula = saved
    DumpSolidDirectNdsBits = result
End Function

Private Function FixtureNames() As Variant
    FixtureNames = Array("Circle", "Rectangle", "LSection", "TwoRectangles", "RoundedSimple", "RoundedTapered", _
        "RoundedMixed", "RoundedMixedLeft", "HollowCentered", "HollowOffset")
End Function

' Сохраняет все биты Double, включая знак нуля; сериализация не зависит от локали Excel.
Private Function Bits(ByVal value As Double) As String
    Dim number As TDoubleValue, words As TDoubleWords
    number.Value = value: LSet words = number
    Bits = Right$("00000000" & Hex$(words.High), 8) & Right$("00000000" & Hex$(words.Low), 8)
End Function
