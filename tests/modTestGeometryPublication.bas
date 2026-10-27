Attribute VB_Name = "modTestGeometryPublication"
Option Explicit

' ==========================================================================
' ДЛЯ ТЕСТОВ: независимость аннотаций, снимка ввода и данных успешно собранной модели.
' ==========================================================================
Private mPassed As Long, mFailed As Long, mReport As String

' -------------------- ДЛЯ ТЕСТОВ --------------------
' повторное построение и смена направления подписи не меняют арматуру и не используют привязки предыдущего построения.
Public Function RunAnnotationIsolationTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    On Error GoTo Failed
    Dim first As CSectionDefinition, second As CSectionDefinition, geometry As CGeometryRegion
    Dim executor As CRebarRuleExecutor, layout As CRebarLayout, builder As CSectionAnnotationBuilder
    Dim annotations As CSectionAnnotations, before As String, initial As String, i As Long
    Set first = TestLabelDefinition(1#): Set second = TestLabelDefinition(-1#)
    Set geometry = New CGeometryRegion: geometry.Initialize first.ConcreteDomain
    Set executor = New CRebarRuleExecutor: Set layout = executor.Build(first, geometry)
    before = TestLayoutSnapshot(layout)
    Set builder = New CSectionAnnotationBuilder: Set annotations = New CSectionAnnotations
    builder.Build first, layout, annotations, executor
    initial = TestAnnotationSnapshot(annotations)
    TestCheck "labels.first.normal", annotations.NormalY(1) = 1#
    For i = 1 To 3
        builder.Build first, layout, annotations, executor
        TestCheck "labels.repeat." & CStr(i), TestAnnotationSnapshot(annotations) = initial
        TestCheck "bars.repeat." & CStr(i), TestLayoutSnapshot(layout) = before
    Next i
    builder.Build second, layout, annotations, executor
    TestCheck "labels.changed.normal", annotations.NormalY(1) = -1#
    TestCheck "labels.changed.count", annotations.Count = 1
    TestCheck "bars.changed.annotation", TestLayoutSnapshot(layout) = before
    builder.Build first, layout, annotations, executor
    TestCheck "labels.restored", TestAnnotationSnapshot(annotations) = initial
    TestCheck "bars.restored", TestLayoutSnapshot(layout) = before
    GoTo Finished
Failed:
    TestCheck "runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    RunAnnotationIsolationTests = mReport & "TOTAL ANNOTATION ISOLATION: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Контрольные определения различаются только направлением одной подписи; координаты и правила совпадают.
Private Function TestLabelDefinition(ByVal normalY As Double) As CSectionDefinition
    Dim facade As CSectionDefinitionBuilder, path As CGeometryPath, options As CRebarRuleOptions
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "LabelIsolation"
    Set path = facade.RectanglePath("Boundary", 0#, 0#, 600#, 1000#)
    facade.Outer "Concrete", path.ID
    Set path = facade.LinePath("Axis", 50#, 950#, 550#, 950#)
    Set options = New CRebarRuleOptions: options.AnnotationGroup = "Top"
    facade.RebarByCount "Row", path.ID, 5, 20#, options
    facade.RebarRowLabel "Top", "Row", 0#, normalY
    Set TestLabelDefinition = facade.Finish
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' завершённая копия единиц измерения не зависит от последующих изменений исходного
' объекта.
Public Function RunRecipeContextFreezeTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    On Error GoTo Failed
    Dim sheet As Object, values(1 To 2, 1 To 3) As Variant, cells(1 To 1, 1 To 1) As Variant
    Dim settings As CSystemSettingsReader, units As CUnitSystem, copy As CUnitSystem, captured As CUnitSystem
    Dim context As CSectionRecipeContext, table As CSectionInputTable, errorCode As Long
    Set sheet = ThisWorkbook.Worksheets.Add
    values(1, 1) = "Parameter": values(1, 2) = "Value": values(1, 3) = "Default"
    values(2, 1) = "Units.Length.Input": values(2, 2) = "cm": values(2, 3) = "cm"
    sheet.Range("A1:C2").Value2 = values
    Set settings = New CSystemSettingsReader: settings.LoadFromRange sheet.Range("A1:C2")
    Set units = New CUnitSystem: units.LoadFromSettings settings
    cells(1, 1) = 100#
    Set table = New CSectionInputTable: table.Initialize "Input", cells
    Set context = New CSectionRecipeContext: context.Initialize units
    context.AddTable table: context.Seal
    Set captured = context.Units
    TestCheck "units.independent.object", Not captured Is units
    TestCheck "units.readOnly", captured.IsReadOnly
    TestCheck "units.original.conversion", captured.InputLengthToInternal(100#) = 1000#
    units.InitializeDefaults
    TestCheck "units.external.reset", captured.InputLengthToInternal(100#) = 1000# And units.InputLengthToInternal(100#) = 100#
    On Error Resume Next
    captured.InitializeDefaults
    errorCode = Err.Number: Err.Clear
    On Error GoTo Failed
    TestCheck "units.reject.reset", errorCode <> 0
    On Error Resume Next
    captured.LoadFromSettings settings
    errorCode = Err.Number: Err.Clear
    On Error GoTo Failed
    TestCheck "units.reject.reload", errorCode <> 0
    TestCheck "units.rejected.calls.preserve.state", captured.InputLengthToInternal(100#) = 1000#
    cells(1, 1) = 200#
    TestCheck "table.original.value", context.Table("Input").Number(1, 1) = 100#
    TestCheck "table.original.units", context.Table("Input").Length(1, 1, captured) = 1000#
    Set copy = captured.Clone
    TestCheck "units.clone.mutable", Not copy.IsReadOnly
    copy.InitializeDefaults
    TestCheck "units.clone.independent", copy.InputLengthToInternal(100#) = 100# And captured.InputLengthToInternal(100#) = 1000#
    GoTo Finished
Failed:
    TestCheck "runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    If Not sheet Is Nothing Then
        Dim alerts As Boolean: alerts = Application.DisplayAlerts
        Application.DisplayAlerts = False: sheet.Delete: Application.DisplayAlerts = alerts
    End If
    RunRecipeContextFreezeTests = mReport & "TOTAL CONTEXT FREEZE: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' выдаётся тот же готовый объект с запретом изменений; массивы повторно не
' копируются.
Public Function RunPublishedBuildTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    On Error GoTo Failed
    Dim definition As CSectionDefinition, builder As CSectionBuilder, model As CSectionModel, meshOptions As TSectionMeshOptions
    Dim mesh As CFiberMeshBuilder, sameMesh As CFiberMeshBuilder, rebars As CRebarLayout, sameRebars As CRebarLayout
    Dim rules As CRebarRuleExecutor, sameRules As CRebarRuleExecutor, geometry As CGeometryRegion, unused As CRebarLayout
    Dim before As String, trace As String, errorCode As Long, count As Long, i As Long
    Set definition = TestLabelDefinition(1#): Set builder = New CSectionBuilder
    meshOptions.StepX = 100#: meshOptions.StepY = 100#: meshOptions.MaterialID = 1: meshOptions.BoundarySubdivisions = 1
    Set model = builder.Build(definition, meshOptions)
    Set mesh = builder.Mesh: Set sameMesh = builder.Mesh
    Set rebars = builder.Rebars: Set sameRebars = builder.Rebars
    Set rules = builder.Rules: Set sameRules = builder.Rules
    TestCheck "published.mesh.same.object", mesh Is sameMesh
    TestCheck "published.rebars.same.object", rebars Is sameRebars
    TestCheck "published.rules.same.object", rules Is sameRules
    TestCheck "published.readOnly", mesh.IsReadOnly And rebars.IsReadOnly And rules.IsReadOnly
    TestCheck "published.counts", mesh.FiberCount = model.ConcreteCount And rebars.Count = model.RebarCount
    For i = 1 To model.RebarCount
        TestCheck "published.bar." & CStr(i), model.RebarX(i) = rebars.X(i) And model.RebarY(i) = rebars.Y(i) And model.RebarDiameter(i) = rebars.Diameter(i)
    Next i
    before = TestLayoutSnapshot(rebars): trace = rules.DiagnosticsText
    Set geometry = New CGeometryRegion: geometry.Initialize definition.ConcreteDomain
    On Error Resume Next
    mesh.BuildMesh geometry, 50#, 50#, 1, 1
    errorCode = Err.Number: Err.Clear
    On Error GoTo Failed
    TestCheck "published.mesh.reject.rebuild", errorCode <> 0 And mesh.FiberCount = model.ConcreteCount
    On Error Resume Next
    rebars.AddBar "Intrusion", 100#, 100#, 10#, 0#, "Rebar"
    errorCode = Err.Number: Err.Clear
    On Error GoTo Failed
    TestCheck "published.rebars.reject.add", errorCode <> 0
    On Error Resume Next
    rebars.AddMembership 1, "Other"
    errorCode = Err.Number: Err.Clear
    On Error GoTo Failed
    TestCheck "published.rebars.reject.group", errorCode <> 0
    On Error Resume Next
    rebars.Clear
    errorCode = Err.Number: Err.Clear
    On Error GoTo Failed
    TestCheck "published.rebars.reject.clear", errorCode <> 0 And TestLayoutSnapshot(rebars) = before
    On Error Resume Next
    Set unused = rules.Build(definition, geometry)
    errorCode = Err.Number: Err.Clear
    On Error GoTo Failed
    TestCheck "published.rules.reject.rebuild", errorCode <> 0 And rules.DiagnosticsText = trace
    model.AddRebarElement 100#, 100#, 10#, 0#, "Rebar", 1, "AfterBuild"
    On Error Resume Next
    Set sameMesh = builder.Mesh
    errorCode = Err.Number: Err.Clear
    On Error GoTo Failed
    TestCheck "published.reject.changed.model.mesh", errorCode <> 0
    On Error Resume Next
    Set sameRebars = builder.Rebars
    errorCode = Err.Number: Err.Clear
    On Error GoTo Failed
    TestCheck "published.reject.changed.model.rebars", errorCode <> 0
    On Error Resume Next
    Set sameRules = builder.Rules
    errorCode = Err.Number: Err.Clear
    On Error GoTo Failed
    TestCheck "published.reject.changed.model.rules", errorCode <> 0
    Set model = builder.Build(definition, meshOptions)
    TestCheck "published.new.build", model.RebarCount = 5 And builder.Rebars.IsReadOnly
    TestCheck "published.old.reference.stable", TestLayoutSnapshot(rebars) = before
    GoTo Finished
Failed:
    TestCheck "runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    RunPublishedBuildTests = mReport & "TOTAL PUBLISHED BUILD: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' создание линии, её однократное добавление в описание сечения и запрет изменений
' готового описания.
Public Function RunDefinitionLifecycleTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    On Error GoTo Failed
    Dim facade As CSectionDefinitionBuilder, path As CGeometryPath, external As CGeometryPath
    Dim definition As CSectionDefinition, unused As CSectionDefinition, options As CRebarRuleOptions
    Dim builder As CSectionBuilder, model As CSectionModel, mesh As CFiberMeshBuilder
    Dim meshOptions As TSectionMeshOptions, errorCode As Long, reason As String
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "Lifecycle"
    Set path = facade.RectanglePath("Boundary", 0#, 0#, 600#, 1000#)
    TestCheck "path.creator.sealed", path.IsSealed
    On Error Resume Next
    facade.AddPath path
    errorCode = Err.Number: Err.Clear
    On Error GoTo Failed
    TestCheck "path.creator.already.registered", errorCode <> 0
    facade.Outer "Outer", path.ID
    Set external = New CGeometryPath: external.Initialize "External"
    external.AddLine "External.Segment", 50#, 950#, 550#, 950#: external.Seal
    facade.AddPath external
    Set options = New CRebarRuleOptions
    facade.RebarByCount "Row", external.ID, 5, 20#, options
    Set definition = facade.Finish
    TestCheck "definition.finish.sealed", definition.IsSealed
    TestCheck "definition.paths.once", definition.PathCount = 2
    On Error Resume Next
    Set unused = facade.Finish
    errorCode = Err.Number: Err.Clear
    On Error GoTo Failed
    TestCheck "definition.reject.second.finish", errorCode <> 0
    On Error Resume Next
    facade.AddPath external
    errorCode = Err.Number: Err.Clear
    On Error GoTo Failed
    TestCheck "definition.reject.add.after.finish", errorCode <> 0
    On Error Resume Next
    facade.Initialize "Again"
    errorCode = Err.Number: Err.Clear
    On Error GoTo Failed
    TestCheck "definition.reject.reinitialize", errorCode <> 0
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "InvalidStep"
    Set path = facade.RectanglePath("Boundary", 0#, 0#, 600#, 1000#): facade.Outer "Outer", path.ID
    Set path = facade.LinePath("Axis", 50#, 950#, 550#, 950#)
    facade.RebarByMaxStep "InvalidRow", path.ID, 0#, 20#, options
    Set definition = facade.Finish
    Set builder = New CSectionBuilder
    meshOptions.StepX = 100#: meshOptions.StepY = 100#: meshOptions.MaterialID = 1: meshOptions.BoundarySubdivisions = 1
    On Error Resume Next
    Set model = builder.Build(definition, meshOptions)
    errorCode = Err.Number: reason = Err.Description: Err.Clear
    On Error GoTo Failed
    TestCheck "build.reject.invalid.step", errorCode <> 0 And InStr(1, reason, "Шаг должен быть положительным", vbTextCompare) > 0
    On Error Resume Next
    Set mesh = builder.Mesh
    errorCode = Err.Number: Err.Clear
    On Error GoTo Failed
    TestCheck "build.failure.no.published.mesh", errorCode <> 0
    Set model = builder.Build(TestLabelDefinition(1#), meshOptions)
    TestCheck "build.recover.after.failure", model.RebarCount = 5 And builder.Mesh.IsReadOnly
    GoTo Finished
Failed:
    TestCheck "runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    RunDefinitionLifecycleTests = mReport & "TOTAL DEFINITION LIFECYCLE: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает все физические поля и группы раскладки; внутри нет данных аннотационного построителя.
Private Function TestLayoutSnapshot(ByVal layout As CRebarLayout) As String
    Dim i As Long, groups As Variant, group As Variant
    For i = 1 To layout.Count
        TestLayoutSnapshot = TestLayoutSnapshot & layout.BarID(i) & "|" & CStr(layout.X(i)) & "|" & CStr(layout.Y(i)) & "|" & _
            CStr(layout.Diameter(i)) & "|" & CStr(layout.Area(i)) & "|" & layout.SteelClass(i) & "|" & _
            layout.BarAnnotationGroupName(i) & "|" & layout.Comment(i)
        groups = layout.BarMemberships(i)
        For Each group In groups: TestLayoutSnapshot = TestLayoutSnapshot & "|" & CStr(group): Next group
        TestLayoutSnapshot = TestLayoutSnapshot & vbCrLf
    Next i
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Полный снимок выходных записей: координаты, нормали, подпись, значение, единицы и пояснение.
Private Function TestAnnotationSnapshot(ByVal annotations As CSectionAnnotations) As String
    Dim i As Long
    For i = 1 To annotations.Count
        TestAnnotationSnapshot = TestAnnotationSnapshot & annotations.AnnotationType(i) & "|" & annotations.AnnotationID(i) & "|" & _
            CStr(annotations.StartX(i)) & "|" & CStr(annotations.StartY(i)) & "|" & CStr(annotations.EndX(i)) & "|" & _
            CStr(annotations.EndY(i)) & "|" & CStr(annotations.NormalX(i)) & "|" & CStr(annotations.NormalY(i)) & "|" & _
            annotations.Text(i) & "|" & CStr(annotations.Value(i)) & "|" & annotations.Unit(i) & "|" & annotations.Comment(i) & vbCrLf
    Next i
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Учитывает отдельное условие проверки и сохраняет его имя в итоговом протоколе.
Private Sub TestCheck(ByVal name As String, ByVal condition As Boolean)
    If condition Then
        mPassed = mPassed + 1: mReport = mReport & "PASS: " & name & vbCrLf
    Else
        mFailed = mFailed + 1: mReport = mReport & "FAIL: " & name & vbCrLf
    End If
End Sub
