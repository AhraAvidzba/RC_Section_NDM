Attribute VB_Name = "modTestDeclarativeExamples"
Option Explicit

' ДЛЯ ТЕСТОВ: программная рамка E09; общий конструктор без нового типа/таблиц.
Private mPassed As Long, mFailed As Long, mReport As String

Public Function RunRoundedFrameExampleTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    On Error GoTo Failed
    Dim definition As CSectionDefinition, region As CConcreteRegion, builder As CSectionBuilder, model As CSectionModel
    Dim query As CSectionGeometryQuery, x As Double, y As Double, mesh As TSectionMeshOptions
    Set definition = RoundedFrameDefinition: Set region = definition.ConcreteDomain
    Set query = New CSectionGeometryQuery: query.RegionCentroid region, x, y
    ' Суммы сохраняют заданные Double при ограничении литералов VBE 15 цифрами.
    Near "frame.area.independent", region.Area, 775879.644737231 + 1.16415321826935E-10
    Near "frame.centroid.x.independent", x, -11.5112278983836 - 1.59872115546023E-14
    Near "frame.centroid.y.independent", y, 4.60449115935345 - 3.5527136788005E-15
    Set builder = New CSectionBuilder
    mesh.StepX = 80#: mesh.StepY = 80#: mesh.MaterialID = 1: mesh.BoundarySubdivisions = 4
    Set model = builder.Build(definition, mesh)
    Check "frame.physical.bars", model.RebarCount = 40 And builder.Rules.CreatedPositions = 40
    Check "frame.no.overlap", builder.Rules.BarOverlaps = 0
    Check "frame.contours", model.Contours.Count = 16
    Check "frame.labels", model.AnnotationCount >= 4
    Dim solver As CSectionSolver, provider As CMaterialModelProvider, settings As CSystemSettingsReader, units As CUnitSystem
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    Set provider = New CMaterialModelProvider: provider.Initialize settings, units
    Set solver = New CSectionSolver
    solver.Solve model, provider.ConcreteMaterial(cpStrength), provider.SteelMaterial(cpStrength), -200000#, 10000000#, -20000000#
    Check "frame.solve", solver.Converged
    Check "frame.residual.n", Abs(solver.RelativeResidualN) < 0.0001
    Check "frame.residual.mx", Abs(solver.RelativeResidualMx) < 0.0001
    Check "frame.residual.my", Abs(solver.RelativeResidualMy) < 0.0001
    ' Механическая сетка имеет своё приближение; она не подменяется A области.
    Check "frame.mesh.exists", model.ConcreteCount > 0 And model.ConcreteMeshSource = "Generated"
    GoTo Finished
Failed:
    Check "frame.runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    RunRoundedFrameExampleTests = mReport & "TOTAL ROUNDED FRAME EXAMPLE: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

Public Function DumpRoundedFrameExamplePreview(ByVal outputDirectory As String) As String
    Dim builder As CSectionBuilder, model As CSectionModel, definition As CSectionDefinition, mesh As TSectionMeshOptions
    Set definition = RoundedFrameDefinition: Set builder = New CSectionBuilder
    mesh.StepX = 80#: mesh.StepY = 80#: mesh.MaterialID = 1: mesh.BoundarySubdivisions = 4
    Set model = builder.Build(definition, mesh)
    DumpRoundedFrameExamplePreview = modTestShapeArchitecture.PreparedPreviewSnapshot("RoundedFrame.E09", model, outputDirectory)
End Function

' ДЛЯ ТЕСТОВ: все входы уже в мм; 24 наружных и 16 внутренних стержней.
Public Function RoundedFrameDefinition() As CSectionDefinition
    Dim facade As CSectionDefinitionBuilder, path As CGeometryPath, options As CRebarRuleOptions
    Dim outerVertices As Variant, openingVertices As Variant, outerCorners As Variant, innerCorners As Variant
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "RoundedFrameExample"
    outerVertices = Array(Array(-600#, -400#), Array(600#, -400#), Array(600#, 400#), Array(-600#, 400#))
    openingVertices = Array(Array(-250#, -170#), Array(350#, -170#), Array(350#, 130#), Array(-250#, 130#))
    outerCorners = Array("Outer.BL", "Outer.BR", "Outer.TR", "Outer.TL")
    innerCorners = Array("Inner.BL", "Inner.BR", "Inner.TR", "Inner.TL")
    Set path = facade.RoundedPolygon("Outer", outerVertices, Array(80#, 80#, 80#, 80#), outerCorners)
    facade.Outer "Concrete", path.ID
    Set path = facade.RoundedPolygon("Opening", openingVertices, Array(40#, 40#, 40#, 40#), innerCorners)
    facade.Opening "Void", path.ID, "Concrete"
    Set path = facade.OffsetRoundedPolygon("Outer.Axis", outerVertices, Array(80#, 80#, 80#, 80#), outerCorners, 60#)
    Set options = New CRebarRuleOptions
    options.EndMode = ClosedUniform: options.PlacementCheck = CheckFullDisk: options.OverlapPolicy = RejectOverlap
    options.AnnotationGroup = "OuterBars": options.SourcePrefix = "Outer."
    facade.RebarByCount "OuterBars", path.ID, 24, 20#, options
    Set path = facade.OffsetRoundedPolygon("Inner.Axis", openingVertices, Array(40#, 40#, 40#, 40#), innerCorners, -45#)
    Set options = New CRebarRuleOptions
    options.EndMode = ClosedUniform: options.PlacementCheck = CheckFullDisk: options.OverlapPolicy = RejectOverlap
    options.AnnotationGroup = "InnerBars": options.SourcePrefix = "Inner."
    facade.RebarByCount "InnerBars", path.ID, 16, 16#, options
    facade.VerticalExtent "Height", "Outer", "H", -600#, -1#
    facade.HorizontalExtent "Width", "Outer", "B", -400#, -1#
    facade.VerticalExtent "OpeningHeight", "Opening", "Ho", 350#, 1#
    facade.HorizontalExtent "OpeningWidth", "Opening", "Bo", 130#, 1#
    facade.RebarLabel "OuterBars", "Outer.Axis", 0#, 1#
    facade.RebarLabel "InnerBars", "Inner.Axis", 0#, -1#
    Set RoundedFrameDefinition = facade.Finish
End Function

Private Sub Check(ByVal name As String, ByVal value As Boolean)
    If value Then mPassed = mPassed + 1 Else mFailed = mFailed + 1: mReport = mReport & "FAIL: " & name & vbCrLf
End Sub

Private Sub Near(ByVal name As String, ByVal actual As Double, ByVal expected As Double)
    Check name & "; actual=" & CStr(actual) & "; expected=" & CStr(expected), Abs(actual - expected) <= 0.000000001 * (1# + Abs(expected))
End Sub
