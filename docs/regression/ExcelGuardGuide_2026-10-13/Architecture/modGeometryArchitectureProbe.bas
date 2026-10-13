Attribute VB_Name = "modGeometryArchitectureProbe"
Option Explicit

' ДЛЯ ТЕСТОВ: владение объектами, защита публикации и ограничения общих контрактов.
Private mPassed As Long, mFailed As Long, mReport As String

' ДЛЯ ТЕСТОВ: проверяет копии входа и геометрии; отдельно фиксирует изменяемые единицы и позднюю проверку шага.
Public Function RunGeometryArchitectureProbe() As String
    mPassed = 0: mFailed = 0: mReport = ""
    On Error GoTo Failed
    Dim facade As CSectionDefinitionBuilder, boundary As CGeometryPath, axis As CGeometryPath
    Dim stored As CGeometryPath, definition As CSectionDefinition, options As CRebarRuleOptions
    Dim rule As TRebarRule, stations As Variant, errorCode As Long, region As CConcreteRegion
    Dim edges() As TRegionEdge, openings() As Boolean
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "OwnershipProbe"
    Set boundary = facade.RectanglePath("Boundary", 0#, 0#, 600#, 1000#)
    facade.Outer "Concrete", boundary.ID
    Set axis = facade.LinePath("Axis", 50#, 950#, 550#, 950#)
    Set options = New CRebarRuleOptions: options.StartSetback = 10#
    facade.RebarByCount "Row", axis.ID, 5, 20#, options
    options.StartSetback = 99#
    stations = Array(20#, 40#)
    facade.RebarAtStations "Explicit", axis.ID, stations, 10#
    stations(0) = 100#
    On Error Resume Next
    facade.AddPath boundary
    errorCode = Err.Number: Err.Clear
    On Error GoTo Failed
    Check "already registered path cannot be added twice", errorCode <> 0
    Set definition = facade.Finish
    Set stored = definition.Path("Boundary")
    Check "returned path is the registered object", stored Is boundary
    Check "two paths but only one concrete loop", definition.PathCount = 2 And definition.LoopCount = 1
    Check "auxiliary axis excluded from concrete area", Abs(definition.ConcreteDomain.Area - 600000#) < 0.000001
    definition.CopyRule 1, rule
    Check "options copied at declaration", rule.StartSetback = 10#
    definition.CopyRule 2, rule
    Check "explicit positions copied at declaration", rule.ExplicitStations(0) = 20#
    On Error Resume Next
    boundary.AddLine "Extra", 0#, 0#, 1#, 0#
    errorCode = Err.Number: Err.Clear
    On Error GoTo Failed
    Check "published path rejects mutation", errorCode <> 0
    On Error Resume Next
    facade.Anchor "AfterFinish", 0#, 0#
    errorCode = Err.Number: Err.Clear
    On Error GoTo Failed
    Check "finished facade rejects mutation", errorCode <> 0
    Set region = definition.ConcreteDomain
    region.SetGeometry edges, 0, 0, 0#, "Changed", 0, openings
    Check "returned domain is a deep geometric copy", Abs(definition.ConcreteDomain.Area - 600000#) < 0.000001
    ProbeUnits
    ProbeRuleValidation
    GoTo Finished
Failed:
    mFailed = mFailed + 1
    mReport = mReport & "RUNTIME ERROR: " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Finished:
    RunGeometryArchitectureProbe = mReport & "TOTAL ARCHITECTURE GUARDS: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

Private Sub ProbeUnits()
    Dim sheet As Object, settings As CSystemSettingsReader, units As CUnitSystem
    Dim context As CSectionRecipeContext, data(1 To 2, 1 To 3) As Variant
    Dim table As CSectionInputTable, cells(1 To 1, 1 To 1) As Variant, before As Double, after As Double
    Set sheet = ThisWorkbook.Worksheets.Add
    data(1, 1) = "Parameter": data(1, 2) = "Value": data(1, 3) = "Default"
    data(2, 1) = "Units.Length.Input": data(2, 2) = "cm": data(2, 3) = "cm"
    sheet.Range("A1:C2").Value2 = data
    Set settings = New CSystemSettingsReader: settings.LoadFromRange sheet.Range("A1:C2")
    Set units = New CUnitSystem: units.LoadFromSettings settings
    cells(1, 1) = 100#
    Set table = New CSectionInputTable: table.Initialize "Input", cells
    Set context = New CSectionRecipeContext: context.Initialize units
    context.AddTable table: context.Seal
    cells(1, 1) = 200#
    Check "input values copied independently", context.Table("Input").Number(1, 1) = 100#
    before = context.Units.InputLengthToInternal(100#)
    units.InitializeDefaults
    after = context.Units.InputLengthToInternal(100#)
    mReport = mReport & "OBSERVED SHARED_UNITS: before_mm=" & CStr(before) & "; after_mm=" & CStr(after) & "; sealed_context_changed=" & CStr(before <> after) & vbCrLf
End Sub

Private Sub ProbeRuleValidation()
    Dim facade As CSectionDefinitionBuilder, path As CGeometryPath, definition As CSectionDefinition
    Dim geometry As CGeometryRegion, executor As CRebarRuleExecutor, layout As CRebarLayout
    Dim errorCode As Long, description As String
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "InvalidStepProbe"
    Set path = facade.RectanglePath("Boundary", 0#, 0#, 600#, 1000#)
    facade.Outer "Concrete", path.ID
    Set path = facade.LinePath("Axis", 50#, 950#, 550#, 950#)
    facade.RebarByMaxStep "Row", path.ID, 0#, 20#
    Set definition = facade.Finish
    mReport = mReport & "OBSERVED LATE_RULE_VALIDATION: invalid_step_accepted_by_finish=" & CStr(definition.IsSealed) & vbCrLf
    Set geometry = New CGeometryRegion: geometry.Initialize definition.ConcreteDomain
    Set executor = New CRebarRuleExecutor
    On Error Resume Next
    Set layout = executor.Build(definition, geometry)
    errorCode = Err.Number: description = Err.Description: Err.Clear
    On Error GoTo 0
    Check "invalid step rejected by rule executor", errorCode <> 0
    mReport = mReport & "OBSERVED EXECUTOR_REJECTION: " & CStr(errorCode) & "; " & description & vbCrLf
End Sub

Private Sub Check(ByVal name As String, ByVal condition As Boolean)
    If condition Then
        mPassed = mPassed + 1: mReport = mReport & "PASS: " & name & vbCrLf
    Else
        mFailed = mFailed + 1: mReport = mReport & "FAIL: " & name & vbCrLf
    End If
End Sub

' ДЛЯ ТЕСТОВ: проверяет влияние аргумента подписи и сохранение старых привязок при повторном построении.
Public Function RunAnnotationOwnershipProbe() As String
    mPassed = 0: mFailed = 0: mReport = ""
    On Error GoTo Failed
    Dim first As String, second As String
    first = LabelSnapshot(50#): second = LabelSnapshot(5000#)
    mReport = mReport & "OBSERVED LABEL_AXIS_DISTANCE: 50_and_5000_produce_same_annotation=" & CStr(first = second) & vbCrLf
    ProbeChangedAnnotation
    GoTo Finished
Failed:
    mFailed = mFailed + 1
    mReport = mReport & "RUNTIME ERROR: " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Finished:
    RunAnnotationOwnershipProbe = mReport & "TOTAL ANNOTATION OWNERSHIP: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

Private Sub ProbeChangedAnnotation()
    Dim first As CSectionDefinition, second As CSectionDefinition, geometry As CGeometryRegion
    Dim executor As CRebarRuleExecutor, rebars As CRebarLayout, annotations As CSectionAnnotations
    Dim builder As CSectionAnnotationBuilder, actual As Double
    Set first = DefinitionWithLabel(1#): Set second = DefinitionWithLabel(-1#)
    Set geometry = New CGeometryRegion: geometry.Initialize first.ConcreteDomain
    Set executor = New CRebarRuleExecutor: Set rebars = executor.Build(first, geometry)
    Set annotations = New CSectionAnnotations: Set builder = New CSectionAnnotationBuilder
    builder.Build first, rebars, annotations, executor
    builder.Build second, rebars, annotations, executor
    actual = annotations.NormalY(1)
    mReport = mReport & "OBSERVED STALE_ANNOTATION: requested_normal_y=-1; actual_normal_y=" & CStr(actual) & "; stale_anchor=" & CStr(actual <> -1#) & vbCrLf
End Sub

Private Function DefinitionWithLabel(ByVal normalY As Double) As CSectionDefinition
    Dim facade As CSectionDefinitionBuilder, path As CGeometryPath, options As CRebarRuleOptions
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "RepeatLabelProbe"
    Set path = facade.RectanglePath("Boundary", 0#, 0#, 600#, 1000#)
    facade.Outer "Concrete", path.ID
    Set path = facade.LinePath("Axis", 50#, 950#, 550#, 950#)
    Set options = New CRebarRuleOptions: options.AnnotationGroup = "Top"
    facade.RebarByCount "Row", path.ID, 5, 20#, options
    facade.RebarRowLabel "Top", "Row", 0#, normalY
    Set DefinitionWithLabel = facade.Finish
End Function

Private Function LabelSnapshot(ByVal axisMetadata As Double) As String
    Dim facade As CSectionDefinitionBuilder, path As CGeometryPath, definition As CSectionDefinition
    Dim options As CRebarRuleOptions, geometry As CGeometryRegion, executor As CRebarRuleExecutor
    Dim rebars As CRebarLayout, annotations As CSectionAnnotations, builder As CSectionAnnotationBuilder
    Dim before As Long, i As Long
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "LabelProbe"
    Set path = facade.RectanglePath("Boundary", 0#, 0#, 600#, 1000#)
    facade.Outer "Concrete", path.ID
    Set path = facade.LinePath("Axis", 50#, 950#, 550#, 950#)
    Set options = New CRebarRuleOptions: options.AnnotationGroup = "Top"
    facade.RebarByCount "Row", path.ID, 5, 20#, options
    facade.RebarRowLabel "Top", "Row", 0#, 1#, axisMetadata
    Set definition = facade.Finish
    Set geometry = New CGeometryRegion: geometry.Initialize definition.ConcreteDomain
    Set executor = New CRebarRuleExecutor: Set rebars = executor.Build(definition, geometry)
    Set annotations = New CSectionAnnotations: Set builder = New CSectionAnnotationBuilder
    builder.Build definition, rebars, annotations, executor
    before = rebars.AnnotationCount
    Check "one initial annotation", annotations.Count = 1
    For i = 1 To annotations.Count
        LabelSnapshot = LabelSnapshot & annotations.AnnotationID(i) & "|" & annotations.Text(i) & "|" & _
            CStr(annotations.StartX(i)) & "|" & CStr(annotations.StartY(i)) & "|" & _
            CStr(annotations.EndX(i)) & "|" & CStr(annotations.EndY(i)) & "|" & _
            CStr(annotations.NormalX(i)) & "|" & CStr(annotations.NormalY(i)) & "|" & CStr(annotations.Value(i))
    Next i
    builder.Build definition, rebars, annotations, executor
    Check "repeated annotation build preserves physical bar count", rebars.Count = 5
    Check "repeated annotation build produces one label", annotations.Count = 1
    mReport = mReport & "OBSERVED ANNOTATION_SIDE_EFFECT: anchors_before=" & CStr(before) & "; anchors_after=" & CStr(rebars.AnnotationCount) & vbCrLf
End Function
