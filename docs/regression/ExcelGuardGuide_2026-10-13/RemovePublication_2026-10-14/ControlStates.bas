' ДЛЯ ТЕСТОВ: одинаковые НДС для четырёх семейств рецептов и трёх назначений материалов.
' Работает только в частной копии; результаты кодируются существующим точным Row.
Public Function DumpPublicationControlStates() As String
    Dim config As Object, saved As Variant, result As String, fixture As Variant, purpose As Variant
    Dim model As CSectionModel, settings As CSystemSettingsReader, units As CUnitSystem
    Dim props As CSectionPropertiesCalculator, provider As CMaterialModelProvider, solver As CSectionSolver
    Dim concrete As Object, steel As Object, i As Long, n As Double, mx As Double, my As Double
    Dim minX As Double, maxX As Double, minY As Double, maxY As Double, strain As Double, name As String
    Dim number As Long, reason As String
    On Error GoTo Failed
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = config.Formula
    For Each fixture In Array("Circle", "Rectangle", "RoundedTapered", "HollowOffset")
        Set model = FixtureModel(CStr(fixture))
        Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
        Set units = New CUnitSystem: units.LoadFromSettings settings
        Set props = New CSectionPropertiesCalculator: props.CalculateConcrete model
        props.CalculateProjection model, 1#, 0#, False, minX, maxX
        props.CalculateProjection model, 0#, 1#, False, minY, maxY
        n = -2# * props.Area
        mx = n * props.CentroidY + props.Area * (maxY - minY) * 0.2
        my = n * props.CentroidX - props.Area * (maxX - minX) * 0.15
        Set provider = New CMaterialModelProvider: provider.Initialize settings, units
        For Each purpose In Array(cpStrength, cpMcrc, cpCrackedNDS)
            Set concrete = provider.ConcreteMaterial(CLng(purpose)): Set steel = provider.SteelMaterial(CLng(purpose))
            Set solver = New CSectionSolver: solver.ApplySettings settings, units
            solver.Solve model, concrete, steel, n, mx, my
            name = CStr(fixture) & "." & CStr(purpose)
            If Not solver.Converged Then Err.Raise 5, "DumpPublicationControlStates", "Control NDS did not converge: " & name
            result = result & Row(name, "NDS", "state", Array(solver.Converged, n, mx, my, solver.Epsilon0, solver.KappaX, solver.KappaY, _
                solver.Nint, solver.Mxint, solver.Myint, solver.ResidualN, solver.ResidualMx, solver.ResidualMy))
            For i = 1 To model.ConcreteCount
                strain = solver.Epsilon0 + solver.KappaX * model.ConcreteY(i) + solver.KappaY * model.ConcreteX(i)
                result = result & Row(name, "CONCRETE_STATE", model.ConcreteID(i), Array(strain, concrete.StressAtStrain(strain)))
            Next i
            For i = 1 To model.RebarCount
                strain = solver.Epsilon0 + solver.KappaX * model.RebarY(i) + solver.KappaY * model.RebarX(i)
                result = result & Row(name, "REBAR_STATE", model.RebarID(i), Array(strain, steel.StressAtStrain(strain)))
            Next i
        Next purpose
    Next fixture
    GoTo Finished
Failed:
    number = Err.Number: reason = Err.Description
Finished:
    If Not config Is Nothing And Not IsEmpty(saved) Then config.Formula = saved
    If number Then Err.Raise number, "DumpPublicationControlStates", reason
    DumpPublicationControlStates = result
End Function
