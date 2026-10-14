Attribute VB_Name = "modTestRoundedTrapezoidExample"
Option Explicit

' ДЛЯ ТЕСТОВ: установка компактного ввода и штатный workflow только fixture-книги.
' Геометрическая математика и расстановка стержней здесь отсутствуют.
Private mPassed As Long, mFailed As Long, mReport As String

' ДЛЯ ТЕСТОВ: полный минимальный класс из гайда исполняется через ISectionRecipe.
Public Function RunAudit02MinimalRecipeTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    On Error GoTo Failed
    Dim recipe As ISectionRecipe, context As CSectionRecipeContext, definition As CSectionDefinition
    Dim builder As CSectionBuilder, mesh As TSectionMeshOptions, model As CSectionModel
    Set recipe = New CSectionRecipeMySection: Set context = New CSectionRecipeContext
    Set definition = recipe.CreateDefinition(context)
    mesh.StepX = 50#: mesh.StepY = 50#: mesh.BoundarySubdivisions = 2: mesh.MaterialID = 1
    Set builder = New CSectionBuilder: Set model = builder.Build(definition, mesh)
    Check "minimal.area", definition.ConcreteDomain.Area = 600000#
    Check "minimal.bars", model.RebarCount = 5
    Check "minimal.annotations", model.Annotations.Count = 2
    Check "minimal.x.first", model.RebarX(1) = 50# And model.RebarY(1) = 950#
    Check "minimal.x.last", model.RebarX(5) = 550# And model.RebarY(5) = 950#
    GoTo Finished
Failed:
    Check "minimal.runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    RunAudit02MinimalRecipeTests = mReport & "TOTAL AUDIT02 MINIMAL RECIPE: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

Public Function RunExampleBuildTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    On Error GoTo Failed
    InstallExampleInput
    Dim settings As CSystemSettingsReader, units As CUnitSystem, builder As CSectionBuilder, model As CSectionModel
    Dim registry As CSectionTypeRegistry, definition As CSectionDefinition, context As CSectionRecipeContext
    Dim snapshot As CSectionInputSnapshotReader, table As Object, saved As Variant, oldArea As Double, changed As CSectionModel
    Dim recipe As CSectionRecipeRoundedTrapezoid, programmed As CSectionDefinition, mesh As TSectionMeshOptions
    Dim id As Variant, found As Boolean, i As Long
    Set registry = New CSectionTypeRegistry
    For Each id In registry.RegisteredRecipeTypes
        If CStr(id) = "RoundedTrapezoid" Then found = True
    Next id
    Check "registration.visible", found
    Check "registration.one.root", registry.RecipeTableNames("RoundedTrapezoid")(0) = "rngRoundedTrapezoidGeometry"
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    Set builder = New CSectionBuilder: Set model = BuildWorkbookSectionModel(ThisWorkbook, settings, units, builder)
    Check "build.physical.bars", model.RebarCount = 40 And builder.Rules.CreatedPositions = 40
    Check "build.no.overlaps", builder.Rules.BarOverlaps = 0
    Check "build.mesh", model.ConcreteCount > 0 And model.ConcreteMeshSource = "Generated"
    Check "build.contours", model.Contours.Count = 9
    Check "build.roles", model.Contours.LoopRole(1) = "Outer" And model.Contours.LoopRole(9) = "Opening"
    Set definition = ReadWorkbookSectionDefinition(ThisWorkbook, settings, units): oldArea = definition.ConcreteDomain.Area
    Check "geometry.area", oldArea > 700000# And oldArea < 800000#
    Set recipe = New CSectionRecipeRoundedTrapezoid: Set programmed = recipe.ProgramDefinition
    Near "program.same.area", programmed.ConcreteDomain.Area, oldArea
    mesh = ReadSectionMeshOptions(settings, units): Set changed = builder.Build(programmed, mesh)
    For i = 1 To 40
        Near "program.same.x." & CStr(i), changed.RebarX(i), model.RebarX(i)
        Near "program.same.y." & CStr(i), changed.RebarY(i), model.RebarY(i)
    Next i
    Set table = ThisWorkbook.Names("rngRoundedTrapezoidGeometry").RefersToRange: saved = table.Formula
    Set snapshot = New CSectionInputSnapshotReader
    Set context = snapshot.Capture(ThisWorkbook, registry.RecipeTableNames("RoundedTrapezoid"), units)
    table.Cells(2, 2).Value2 = 1400#
    Set programmed = registry.CreateDefinition("RoundedTrapezoid", context)
    Near "captured.immutable", programmed.ConcreteDomain.Area, oldArea
    settings.LoadFromWorkbook ThisWorkbook
    Set changed = BuildWorkbookSectionModel(ThisWorkbook, settings, units)
    Check "changed.same.count", changed.RebarCount = 40
    Set programmed = ReadWorkbookSectionDefinition(ThisWorkbook, settings, units)
    Check "changed.area", programmed.ConcreteDomain.Area > oldArea + 70000#
    Near "old.definition.immutable", definition.ConcreteDomain.Area, oldArea
    table.Formula = saved
    table.Cells(6, 2).Value2 = 1000#
    On Error Resume Next
    Set changed = BuildWorkbookSectionModel(ThisWorkbook, settings, units)
    i = Err.Number: id = Err.Description: Err.Clear
    On Error GoTo Failed
    Check "invalid.opening.address", i <> 0 And InStr(CStr(id), "Config!") > 0
    table.Formula = saved
    GoTo Finished
Failed:
    Check "example.runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    On Error Resume Next
    If Not table Is Nothing And Not IsEmpty(saved) Then table.Formula = saved
    On Error GoTo 0
    RunExampleBuildTests = mReport & "TOTAL ONE CLASS BUILD: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

Public Function RunExampleWorkflow(ByVal outputDirectory As String, Optional ByVal bottomWidth As Double = 1200#) As String
    mPassed = 0: mFailed = 0: mReport = ""
    On Error GoTo Failed
    InstallExampleInput
    ThisWorkbook.Names("rngRoundedTrapezoidGeometry").RefersToRange.Cells(2, 2).Value2 = bottomWidth
    Dim settings As CSystemSettingsReader, units As CUnitSystem, model As CSectionModel, text As String, started As Double
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    Set model = BuildWorkbookSectionModel(ThisWorkbook, settings, units)
    text = modTestShapeArchitecture.PreparedPreviewSnapshot("RoundedTrapezoid.Preview." & CStr(bottomWidth), model, outputDirectory)
    Check "preview.rendered", Len(text) > 0
    ConfigureSmoke
    started = Timer
    text = RunSectionCalculationForWorkbook(ThisWorkbook, False)
    mReport = mReport & "FULL_WORKFLOW_SECONDS: " & CStr(Timer - started) & vbCrLf & text & vbCrLf
    Check "smoke.completed", Len(text) > 0
    Dim savedModel As CSectionModel, data As CSectionPlotDataReader, chart As Object, summary As Object, count As Long
    Set savedModel = ReadSectionGeometryFromResults(ThisWorkbook)
    Check "smoke.saved.bars", savedModel.RebarCount = 40
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set data = New CSectionPlotDataReader: data.LoadFromWorkbook ThisWorkbook, settings
    Check "smoke.named.state", data.Count = model.ConcreteCount + 40
    Set chart = ThisWorkbook.Worksheets("Расчет").ChartObjects("chtNDMSectionPlot").Chart
    Check "smoke.scheme", chart.Shapes.Count > model.RebarCount
    Check "smoke.png", chart.Export(outputDirectory & Application.PathSeparator & "RoundedTrapezoid.Smoke." & CStr(bottomWidth) & ".png", "PNG")
    Set summary = ThisWorkbook.Names("rngBatchSummary").RefersToRange
    Dim row As Long
    For row = 1 To summary.Rows.Count
        If CStr(summary.Cells(row, 1).Value2) = "EXAMPLE" Then count = count + 1
    Next row
    Check "smoke.one.lc", count = 1
    GoTo Finished
Failed:
    Check "workflow.runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    RunExampleWorkflow = mReport & "TOTAL ONE CLASS WORKFLOW: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' ДЛЯ ТЕСТОВ: глобальные единицы и формулы компактного ввода в manual mode.
Public Function RunExampleInputTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    Dim calculation As Long, table As Object, baseline As Variant, unit As Variant, factor As Double, i As Long
    On Error GoTo Failed
    calculation = Application.Calculation
    InstallExampleInput
    Set table = ThisWorkbook.Names("rngRoundedTrapezoidGeometry").RefersToRange: baseline = table.Formula
    Dim settings As CSystemSettingsReader, units As CUnitSystem, definition As CSectionDefinition, original As CSectionDefinition
    Dim layout As CRebarLayout, originalLayout As CRebarLayout, reader As CSectionInputTable, reason As String, number As Long
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    Set original = ReadWorkbookSectionDefinition(ThisWorkbook, settings, units): Set originalLayout = FixtureAnnotatedLayout(original)
    For Each unit In Array("mm", "cm", "m")
        table.Formula = baseline: SetInputUnit "Length", CStr(unit)
        factor = 1#: If unit = "cm" Then factor = 10# Else If unit = "m" Then factor = 1000#
        For i = 2 To 12
            If i <> 9 And i <> 10 Then table.Cells(i, 2).Value2 = CDbl(baseline(i, 2)) / factor
        Next i
        settings.LoadFromWorkbook ThisWorkbook: units.LoadFromSettings settings
        Set definition = ReadWorkbookSectionDefinition(ThisWorkbook, settings, units): Set layout = FixtureAnnotatedLayout(definition)
        Near "units.area." & CStr(unit), definition.ConcreteDomain.Area, original.ConcreteDomain.Area
        Check "units.count." & CStr(unit), layout.Count = originalLayout.Count
        For i = 1 To 40
            Near "units.x." & CStr(unit) & "." & CStr(i), layout.X(i), originalLayout.X(i)
            Near "units.y." & CStr(unit) & "." & CStr(i), layout.Y(i), originalLayout.Y(i)
        Next i
    Next unit
    table.Formula = baseline: SetInputUnit "Length", "mm"
    Application.Calculation = xlCalculationManual
    table.Cells(2, 2).Formula = "=" & table.Cells(3, 2).Address & "*1.5": table.Calculate
    table.Cells(3, 2).Value2 = 1000#
    Near "manual.stale.fixture", CDbl(table.Cells(2, 2).Value2), 1200#
    settings.LoadFromWorkbook ThisWorkbook: units.LoadFromSettings settings
    Set definition = ReadWorkbookSectionDefinition(ThisWorkbook, settings, units)
    Near "manual.capture.refreshed", CDbl(table.Cells(2, 2).Value2), 1500#
    Dim recipe As CSectionRecipeRoundedTrapezoid, expected As CSectionDefinition
    Set recipe = New CSectionRecipeRoundedTrapezoid: Set expected = recipe.ProgramDefinition(1500#, 1000#)
    Near "manual.formula.same.definition", definition.ConcreteDomain.Area, expected.ConcreteDomain.Area
    table.Formula = baseline: table.Cells(7, 2).ClearContents
    On Error Resume Next
    Set definition = ReadWorkbookSectionDefinition(ThisWorkbook, settings, units)
    number = Err.Number: reason = Err.Description: Err.Clear
    On Error GoTo Failed
    Check "active.blank.address", number <> 0 And InStr(reason, "Config!") > 0
    GoTo Finished
Failed:
    Check "example.input.runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    On Error Resume Next
    If Not table Is Nothing And Not IsEmpty(baseline) Then table.Formula = baseline
    SetInputUnit "Length", "mm": Application.Calculation = calculation
    On Error GoTo 0
    RunExampleInputTests = mReport & "TOTAL ONE CLASS INPUT: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' ДЛЯ ТЕСТОВ: один корневой блок собственных параметров, без редактора примитивов.
Public Sub InstallExampleInput()
    Dim table As Object, sheet As Object, row As Long, values(1 To 12, 1 To 4) As Variant
    Set sheet = ThisWorkbook.Worksheets("Config")
    On Error Resume Next
    Set table = ThisWorkbook.Names("rngRoundedTrapezoidGeometry").RefersToRange
    On Error GoTo 0
    If table Is Nothing Then
        row = sheet.UsedRange.Row + sheet.UsedRange.Rows.Count + 3
        Set table = sheet.Cells(row, 1).Resize(12, 4)
        ThisWorkbook.Names.Add Name:="rngRoundedTrapezoidGeometry", RefersTo:="='Config'!" & table.Address
    End If
    Dim keys As Variant, defaults As Variant, i As Long
    keys = Array("BottomWidth", "TopWidth", "Height", "CornerRadius", "OpeningRadius", "OuterAxisDistance", "InnerAxisDistance", "OuterCount", "InnerCount", "OuterDiameter", "InnerDiameter")
    defaults = Array(1200#, 800#, 800#, 60#, 110#, 50#, 45#, 24, 16, 20#, 16#)
    values(1, 1) = "Параметр": values(1, 2) = "Значение": values(1, 3) = "Ед.": values(1, 4) = "Скругленная трапеция: пример"
    For i = 0 To 10
        values(i + 2, 1) = keys(i): values(i + 2, 2) = defaults(i)
        values(i + 2, 3) = "mm": If i = 7 Or i = 8 Then values(i + 2, 3) = "шт"
    Next i
    table.Value2 = values: table.Borders.LineStyle = xlContinuous: table.Columns(2).HorizontalAlignment = xlCenter
    For i = 2 To 12
        If i <> 9 And i <> 10 Then table.Cells(i, 3).Formula = "=INDEX(rngUnitSettings,MATCH(""Length"",INDEX(rngUnitSettings,,1),0),2)"
    Next i
    SetInputUnit "Length", "mm"
    modTestShapeArchitecture.ConfigureShapeFixture "Circle"
    SetSetting "Geometry.Type", "RoundedTrapezoid": SetSetting "Geometry.Source", "Generated"
    SetSetting "Mesh.StepX", "80": SetSetting "Mesh.StepY", "80": SetSetting "Mesh.BoundarySubdivisions", "4"
    modSectionRecipeSetup.UpdateSectionRecipeConfig
End Sub

Private Sub ConfigureSmoke()
    SetInputUnit "Force", "N": SetInputUnit "Moment", "N*mm"
    SetSetting "Plot.Enabled", "Yes": SetSetting "Plot.AutoUpdateAfterCalculation", "Yes": SetSetting "Plot.LoadCase", "Worst"
    SetProfile "Calculation.Strength.DirectState", "Yes": SetProfile "Calculation.Strength.Capacity", "No"
    SetProfile "Calculation.Crack.Width", "No": SetProfile "Calculation.Stability.Enabled", "No"
    SetProfile "Visualization.State", "StrengthState": SetProfile "Visualization.Quantity", "Stress"
    Dim loads As Object
    Set loads = ThisWorkbook.Names("rngLoadCombinations").RefersToRange
    loads.Offset(1).Resize(loads.Rows.Count - 1, loads.Columns.Count).ClearContents
    loads.Cells(2, 1).Value2 = "EXAMPLE": loads.Cells(2, 2).Value2 = 200000#
    loads.Cells(2, 3).Value2 = 0#: loads.Cells(2, 4).Value2 = 0#: loads.Cells(2, 5).Value2 = "PR1"
    loads.Cells(2, 6).Value2 = "Auto": loads.Cells(2, 7).Value2 = "Smoke нового рецепта"
End Sub

Private Sub SetInputUnit(ByVal key As String, ByVal value As String)
    Dim range As Object, table As CSectionInputTable
    Set range = ThisWorkbook.Names("rngUnitSettings").RefersToRange
    Set table = New CSectionInputTable: table.Initialize "rngUnitSettings", range.Value2
    range.Cells(table.RowByKey(key), 2).Value2 = value
End Sub

Private Sub SetSetting(ByVal key As String, ByVal value As String)
    Dim range As Object, table As CSectionInputTable
    Set range = ThisWorkbook.Names("rngSystemSettings").RefersToRange
    Set table = New CSectionInputTable: table.Initialize "rngSystemSettings", range.Value2
    range.Cells(table.RowByKey(key), 2).Value2 = value
End Sub

Private Sub SetProfile(ByVal key As String, ByVal value As String)
    Dim range As Object, row As Long, column As Long, found As Long
    Set range = ThisWorkbook.Names("rngCalculationProfiles").RefersToRange
    For row = 1 To range.Rows.Count
        For column = 3 To range.Columns.Count
            If CStr(range.Cells(row, column).Value2) = "PR1" Then found = column: Exit For
        Next column
        If found > 0 Then Exit For
    Next row
    If found = 0 Then Err.Raise 5, "SetProfile", "Не найден PR1."
    For row = 1 To range.Rows.Count
        If CStr(range.Cells(row, 2).Value2) = key Then range.Cells(row, found).Value2 = value: Exit Sub
    Next row
    Err.Raise 5, "SetProfile", "Не найден параметр " & key
End Sub

Private Sub Check(ByVal name As String, ByVal value As Boolean)
    If value Then mPassed = mPassed + 1 Else mFailed = mFailed + 1: mReport = mReport & "FAIL: " & name & vbCrLf
End Sub

Private Sub Near(ByVal name As String, ByVal actual As Double, ByVal expected As Double)
    Check name & "; actual=" & CStr(actual) & "; expected=" & CStr(expected), Abs(actual - expected) <= 0.000000001 * (1# + Abs(expected))
End Sub
