Attribute VB_Name = "modTestShapeArchitecture"
Option Explicit

' ДЛЯ ТЕСТОВ: независимый снимок сетки, всей арматуры и всех аннотаций.
' Этот entrypoint выполняется и на книге до рефакторинга, и на новой книге.
' Вход восстанавливается; расчетные Results не записываются.
Public Function DumpShapeArchitectureSnapshot() As String
    Dim config As Object, saved As Variant, result As String, fixture As Variant
    On Error GoTo Failed
    Set config = ThisWorkbook.Worksheets("Config").UsedRange
    saved = config.Formula
    For Each fixture In Array("Circle", "Rectangle", "LSection", "TwoRectangles", _
            "RoundedSimple", "RoundedTapered", "RoundedMixed", "HollowCentered", "HollowOffset")
        result = result & ShapeSnapshot(CStr(fixture))
    Next fixture
    GoTo Finished
Failed:
    result = result & "FAIL: snapshot; " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Finished:
    If Not config Is Nothing And Not IsEmpty(saved) Then config.Formula = saved
    DumpShapeArchitectureSnapshot = result
End Function

Private Function ShapeSnapshot(ByVal fixture As String) As String
    Dim model As CSectionModel, properties As CSectionPropertiesCalculator, annotations As CSectionAnnotations
    Dim result As String, i As Long
    Set model = FixtureModel(fixture)
    Set properties = New CSectionPropertiesCalculator: properties.CalculateConcrete model
    result = Row(fixture, "PROPS", "concrete", Array(model.ConcreteCount, model.RebarCount, _
        properties.Area, properties.CentroidX, properties.CentroidY, properties.Ixc, properties.Iyc, properties.Ixyc))
    For i = 1 To model.RebarCount
        result = result & Row(fixture, "BAR", model.RebarID(i), Array(model.RebarX(i), model.RebarY(i), _
            model.RebarDiameter(i), model.RebarArea(i), model.RebarSteelClass(i), model.RebarComment(i)))
    Next i
    Set annotations = model.Annotations
    For i = 1 To annotations.Count
        result = result & Row(fixture, "ANNOTATION", annotations.AnnotationID(i), Array(annotations.AnnotationType(i), _
            annotations.StartX(i), annotations.StartY(i), annotations.EndX(i), annotations.EndY(i), _
            annotations.NormalX(i), annotations.NormalY(i), annotations.Text(i), annotations.Value(i), annotations.Comment(i)))
    Next i
    ShapeSnapshot = result
End Function

Private Function FixtureModel(ByVal fixture As String) As CSectionModel
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
        Case "RoundedSimple", "RoundedTapered", "RoundedMixed"
            shapeName = "RoundedRectangle": Set table = ThisWorkbook.Names("rngRoundedRectangleGeometry").RefersToRange
            table.Cells(3, 2).Value2 = 600# / factor: table.Cells(3, 3).Value2 = 400# / factor
            table.Cells(7, 2).Value2 = IIf(fixture = "RoundedTapered", "Tapered", "Simple")
            table.Cells(7, 3).Value2 = IIf(fixture = "RoundedSimple", "Simple", "Tapered")
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
    settings.LoadFromWorkbook ThisWorkbook: units.LoadFromSettings settings
    Set registry = New CSectionTypeRegistry
    Set FixtureModel = registry.BuildGeneratedModel(settings, units)
End Function

' ДЛЯ ТЕСТОВ: расчетный эталон до/после рефакторинга без изменения Results.
' Три назначения материалов и четыре направления нагрузки на каждую форму;
' затем общий пакет несущей способности, трещинообразования и ширины СП 35.
Public Function DumpShapeCalculationSnapshot(Optional ByVal crackCode As String = "SP35") As String
    Dim config As Object, saved As Variant, result As String, fixture As Variant
    On Error GoTo Failed
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
    For Each fixture In Array("Circle", "Rectangle", "LSection", "TwoRectangles", _
            "RoundedSimple", "RoundedTapered", "RoundedMixed", "HollowCentered", "HollowOffset")
        result = result & CalculationSnapshot(CStr(fixture))
    Next fixture
    GoTo Finished
Failed:
    result = result & "FAIL: calculation snapshot; " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Finished:
    If Not config Is Nothing And Not IsEmpty(saved) Then config.Formula = saved
    DumpShapeCalculationSnapshot = result
End Function

Private Function CalculationSnapshot(ByVal fixture As String) As String
    Dim model As CSectionModel, props As CSectionPropertiesCalculator
    Dim settings As CSystemSettingsReader, units As CUnitSystem, provider As CMaterialModelProvider
    Set model = FixtureModel(fixture)
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
                solver.MinConcreteStress, solver.MaxConcreteStress, solver.MinSteelStress, solver.MaxSteelStress))
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

Private Function Row(ByVal fixture As String, ByVal kind As String, ByVal id As String, ByVal values As Variant) As String
    Dim item As Variant
    Row = fixture & vbTab & kind & vbTab & id
    For Each item In values: Row = Row & vbTab & CStr(item): Next item
    Row = Row & vbCrLf
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
