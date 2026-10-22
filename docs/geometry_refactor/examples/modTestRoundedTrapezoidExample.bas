Attribute VB_Name = "modTestRoundedTrapezoidExample"
Option Explicit

' ==========================================================================
' ДЛЯ ТЕСТОВ: задаёт короткий набор входных параметров и запускает обычный расчёт
' только в тестовой книге. Геометрические вычисления и расстановку стержней выполняют общие
' классы.
' ==========================================================================
Private mPassed As Long, mFailed As Long, mReport As String

' -------------------- ДЛЯ ТЕСТОВ --------------------
' полный минимальный класс из гайда исполняется через ISectionRecipe.
Public Function RunAudit02MinimalRecipeTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    On Error GoTo Failed
    Dim recipe As ISectionRecipe, context As CSectionRecipeContext, definition As CSectionDefinition
    Dim builder As CSectionBuilder, mesh As TSectionMeshOptions, model As CSectionModel
    Set recipe = New CSectionRecipeMySection: Set context = New CSectionRecipeContext
    Set definition = recipe.CreateDefinition(context)
    mesh.StepX = 50#: mesh.StepY = 50#: mesh.BoundarySubdivisions = 2: mesh.MaterialID = 1
    Set builder = New CSectionBuilder: Set model = builder.Build(definition, mesh)
    TestCheck "minimal.area", definition.ConcreteDomain.Area = 600000#
    TestCheck "minimal.bars", model.RebarCount = 5
    TestCheck "minimal.annotations", model.Annotations.Count = 2
    TestCheck "minimal.x.first", model.RebarX(1) = 50# And model.RebarY(1) = 950#
    TestCheck "minimal.x.last", model.RebarX(5) = 550# And model.RebarY(5) = 950#
    GoTo Finished
Failed:
    TestCheck "minimal.runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    RunAudit02MinimalRecipeTests = mReport & "TOTAL AUDIT02 MINIMAL RECIPE: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает построение учебной трапеции из таблицы и программных параметров; проверяет регистрацию, снимок входов, изменение геометрии и адрес ошибки отверстия.
Public Function RunExampleBuildTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    On Error GoTo Failed
    TestInstallExampleInput
    Dim settings As CSystemSettingsReader, units As CUnitSystem, builder As CSectionBuilder, model As CSectionModel
    Dim registry As CSectionTypeRegistry, definition As CSectionDefinition, context As CSectionRecipeContext
    Dim snapshot As CSectionInputSnapshotReader, table As Object, saved As Variant, oldArea As Double, changed As CSectionModel
    Dim recipe As CSectionRecipeRoundedTrapezoid, programmed As CSectionDefinition, mesh As TSectionMeshOptions
    Dim id As Variant, found As Boolean, i As Long
    Set registry = New CSectionTypeRegistry
    For Each id In registry.RegisteredRecipeTypes
        If CStr(id) = "RoundedTrapezoid" Then found = True
    Next id
    TestCheck "registration.visible", found
    TestCheck "registration.one.root", registry.RecipeTableNames("RoundedTrapezoid")(0) = "rngRoundedTrapezoidGeometry"
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    Set builder = New CSectionBuilder: Set model = BuildWorkbookSectionModel(ThisWorkbook, settings, units, builder)
    TestCheck "build.physical.bars", model.RebarCount = 40 And builder.Rules.CreatedPositions = 40
    TestCheck "build.no.overlaps", builder.Rules.BarOverlaps = 0
    TestCheck "build.mesh", model.ConcreteCount > 0 And model.ConcreteMeshSource = "Generated"
    TestCheck "build.contours", model.Contours.Count = 9
    TestCheck "build.roles", model.Contours.LoopRole(1) = "Outer" And model.Contours.LoopRole(9) = "Opening"
    Set definition = ReadWorkbookSectionDefinition(ThisWorkbook, settings, units): oldArea = definition.ConcreteDomain.Area
    TestCheck "geometry.area", oldArea > 700000# And oldArea < 800000#
    Set recipe = New CSectionRecipeRoundedTrapezoid: Set programmed = recipe.ProgramDefinition
    TestNear "program.same.area", programmed.ConcreteDomain.Area, oldArea
    mesh = ReadSectionMeshOptions(settings, units): Set changed = builder.Build(programmed, mesh)
    For i = 1 To 40
        TestNear "program.same.x." & CStr(i), changed.RebarX(i), model.RebarX(i)
        TestNear "program.same.y." & CStr(i), changed.RebarY(i), model.RebarY(i)
    Next i
    Set table = ThisWorkbook.Names("rngRoundedTrapezoidGeometry").RefersToRange: saved = table.Formula
    Set snapshot = New CSectionInputSnapshotReader
    Set context = snapshot.Capture(ThisWorkbook, registry.RecipeTableNames("RoundedTrapezoid"), units)
    table.Cells(2, 2).Value2 = 1400#
    Set programmed = registry.CreateDefinition("RoundedTrapezoid", context)
    TestNear "captured.immutable", programmed.ConcreteDomain.Area, oldArea
    settings.LoadFromWorkbook ThisWorkbook
    Set changed = BuildWorkbookSectionModel(ThisWorkbook, settings, units)
    TestCheck "changed.same.count", changed.RebarCount = 40
    Set programmed = ReadWorkbookSectionDefinition(ThisWorkbook, settings, units)
    TestCheck "changed.area", programmed.ConcreteDomain.Area > oldArea + 70000#
    TestNear "old.definition.immutable", definition.ConcreteDomain.Area, oldArea
    table.Formula = saved
    table.Cells(6, 2).Value2 = 1000#
    On Error Resume Next
    Set changed = BuildWorkbookSectionModel(ThisWorkbook, settings, units)
    i = Err.Number: id = Err.Description: Err.Clear
    On Error GoTo Failed
    TestCheck "invalid.opening.address", i <> 0 And InStr(CStr(id), "Config!") > 0
    table.Formula = saved
    GoTo Finished
Failed:
    TestCheck "example.runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    On Error Resume Next
    If Not table Is Nothing And Not IsEmpty(saved) Then table.Formula = saved
    On Error GoTo 0
    RunExampleBuildTests = mReport & "TOTAL ONE CLASS BUILD: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Формирует предварительную схему трапеции, выполняет одно сочетание штатной кнопки расчёта и проверяет сохранённую арматуру, НДС и экспорт PNG.
Public Function TestRunExampleWorkflow(ByVal outputDirectory As String, Optional ByVal bottomWidth As Double = 1200#) As String
    mPassed = 0: mFailed = 0: mReport = ""
    On Error GoTo Failed
    TestInstallExampleInput
    ThisWorkbook.Names("rngRoundedTrapezoidGeometry").RefersToRange.Cells(2, 2).Value2 = bottomWidth
    Dim settings As CSystemSettingsReader, units As CUnitSystem, model As CSectionModel, text As String, started As Double
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    Set model = BuildWorkbookSectionModel(ThisWorkbook, settings, units)
    text = modTestShapeArchitecture.TestPreparedPreviewSnapshot("RoundedTrapezoid.Preview." & CStr(bottomWidth), model, outputDirectory)
    TestCheck "preview.rendered", Len(text) > 0
    TestConfigureSmoke
    started = Timer
    text = RunSectionCalculationForWorkbook(ThisWorkbook, False)
    mReport = mReport & "FULL_WORKFLOW_SECONDS: " & CStr(Timer - started) & vbCrLf & text & vbCrLf
    TestCheck "smoke.completed", Len(text) > 0
    Dim savedModel As CSectionModel, data As CSectionPlotDataReader, chart As Object, summary As Object, count As Long
    Set savedModel = ReadSectionGeometryFromResults(ThisWorkbook)
    TestCheck "smoke.saved.bars", savedModel.RebarCount = 40
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set data = New CSectionPlotDataReader: data.LoadFromWorkbook ThisWorkbook, settings
    TestCheck "smoke.named.state", data.Count = model.ConcreteCount + 40
    Set chart = ThisWorkbook.Worksheets("Расчет").ChartObjects("chtNDMSectionPlot").Chart
    TestCheck "smoke.scheme", chart.Shapes.Count > model.RebarCount
    TestCheck "smoke.png", chart.Export(outputDirectory & Application.PathSeparator & "RoundedTrapezoid.Smoke." & CStr(bottomWidth) & ".png", "PNG")
    Set summary = ThisWorkbook.Names("rngBatchSummary").RefersToRange
    Dim row As Long
    For row = 1 To summary.Rows.Count
        If CStr(summary.Cells(row, 1).Value2) = "EXAMPLE" Then count = count + 1
    Next row
    TestCheck "smoke.one.lc", count = 1
    GoTo Finished
Failed:
    TestCheck "workflow.runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    TestRunExampleWorkflow = mReport & "TOTAL ONE CLASS WORKFLOW: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' проверяет единицы и обновление формул входа при ручном режиме пересчёта Excel.
Public Function RunExampleInputTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    Dim calculation As Long, table As Object, baseline As Variant, unit As Variant, factor As Double, i As Long
    On Error GoTo Failed
    calculation = Application.Calculation
    TestInstallExampleInput
    Set table = ThisWorkbook.Names("rngRoundedTrapezoidGeometry").RefersToRange: baseline = table.Formula
    Dim settings As CSystemSettingsReader, units As CUnitSystem, definition As CSectionDefinition, original As CSectionDefinition
    Dim layout As CRebarLayout, originalLayout As CRebarLayout, reader As CSectionInputTable, reason As String, number As Long
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    Set original = ReadWorkbookSectionDefinition(ThisWorkbook, settings, units): Set originalLayout = TestFixtureAnnotatedLayout(original)
    For Each unit In Array("mm", "cm", "m")
        table.Formula = baseline: TestSetInputUnit "Length", CStr(unit)
        factor = 1#: If unit = "cm" Then factor = 10# Else If unit = "m" Then factor = 1000#
        For i = 2 To 12
            If i <> 9 And i <> 10 Then table.Cells(i, 2).Value2 = CDbl(baseline(i, 2)) / factor ' Строки 9 и 10 содержат количества стержней, которые не переводятся вместе с длинами.
        Next i
        settings.LoadFromWorkbook ThisWorkbook: units.LoadFromSettings settings
        Set definition = ReadWorkbookSectionDefinition(ThisWorkbook, settings, units): Set layout = TestFixtureAnnotatedLayout(definition)
        TestNear "units.area." & CStr(unit), definition.ConcreteDomain.Area, original.ConcreteDomain.Area
        TestCheck "units.count." & CStr(unit), layout.Count = originalLayout.Count
        For i = 1 To 40
            TestNear "units.x." & CStr(unit) & "." & CStr(i), layout.X(i), originalLayout.X(i)
            TestNear "units.y." & CStr(unit) & "." & CStr(i), layout.Y(i), originalLayout.Y(i)
        Next i
    Next unit
    table.Formula = baseline: TestSetInputUnit "Length", "mm"
    Application.Calculation = xlCalculationManual
    table.Cells(2, 2).Formula = "=" & table.Cells(3, 2).Address & "*1.5": table.Calculate
    table.Cells(3, 2).Value2 = 1000#
    TestNear "manual.stale.fixture", CDbl(table.Cells(2, 2).Value2), 1200#
    settings.LoadFromWorkbook ThisWorkbook: units.LoadFromSettings settings
    Set definition = ReadWorkbookSectionDefinition(ThisWorkbook, settings, units)
    TestNear "manual.capture.refreshed", CDbl(table.Cells(2, 2).Value2), 1500#
    Dim recipe As CSectionRecipeRoundedTrapezoid, expected As CSectionDefinition
    Set recipe = New CSectionRecipeRoundedTrapezoid: Set expected = recipe.ProgramDefinition(1500#, 1000#)
    TestNear "manual.formula.same.definition", definition.ConcreteDomain.Area, expected.ConcreteDomain.Area
    table.Formula = baseline: table.Cells(7, 2).ClearContents
    On Error Resume Next
    Set definition = ReadWorkbookSectionDefinition(ThisWorkbook, settings, units)
    number = Err.Number: reason = Err.Description: Err.Clear
    On Error GoTo Failed
    TestCheck "active.blank.address", number <> 0 And InStr(reason, "Config!") > 0
    GoTo Finished
Failed:
    TestCheck "example.input.runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    On Error Resume Next
    If Not table Is Nothing And Not IsEmpty(baseline) Then table.Formula = baseline
    TestSetInputUnit "Length", "mm": Application.Calculation = calculation
    On Error GoTo 0
    RunExampleInputTests = mReport & "TOTAL ONE CLASS INPUT: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' один корневой блок собственных параметров, без редактора примитивов.
Public Sub TestInstallExampleInput()
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
    TestSetInputUnit "Length", "mm"
    modTestShapeArchitecture.TestConfigureShapeFixture "Circle"
    TestSetSetting "Geometry.Type", "RoundedTrapezoid": TestSetSetting "Geometry.Source", "Generated"
    TestSetSetting "Mesh.StepX", "80": TestSetSetting "Mesh.StepY", "80": TestSetSetting "Mesh.BoundarySubdivisions", "4"
    modSectionRecipeSetup.UpdateSectionRecipeTypes
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Готовит одно сочетание прочности с силой в Н, моментами в Н·мм и автоматической схемой; несущая способность, трещины и устойчивость отключены.
Private Sub TestConfigureSmoke()
    TestSetInputUnit "Force", "N": TestSetInputUnit "Moment", "N*mm"
    TestSetSetting "Plot.Enabled", "Yes": TestSetSetting "Plot.AutoUpdateAfterCalculation", "Yes": TestSetSetting "Plot.LoadCase", "Worst"
    TestSetProfile "Calculation.Strength.DirectState", "Yes": TestSetProfile "Calculation.Strength.Capacity", "No"
    TestSetProfile "Calculation.Crack.Width", "No": TestSetProfile "Calculation.Stability.Enabled", "No"
    TestSetProfile "Visualization.State", "StrengthState": TestSetProfile "Visualization.Quantity", "Stress"
    Dim loads As Object
    Set loads = ThisWorkbook.Names("rngLoadCombinations").RefersToRange
    loads.Offset(1).Resize(loads.Rows.Count - 1, loads.Columns.Count).ClearContents
    loads.Cells(2, 1).Value2 = "EXAMPLE": loads.Cells(2, 2).Value2 = 200000#
    loads.Cells(2, 3).Value2 = 0#: loads.Cells(2, 4).Value2 = 0#: loads.Cells(2, 5).Value2 = "PR1"
    loads.Cells(2, 6).Value2 = "Auto": loads.Cells(2, 7).Value2 = "Smoke нового рецепта"
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Меняет только единицу INPUT найденной строки таблицы единиц; сохранённые результаты не пересчитывает.
Private Sub TestSetInputUnit(ByVal key As String, ByVal value As String)
    Dim range As Object, table As CSectionInputTable
    Set range = ThisWorkbook.Names("rngUnitSettings").RefersToRange
    Set table = New CSectionInputTable: table.Initialize "rngUnitSettings", range.Value2
    range.Cells(table.RowByKey(key), 2).Value2 = value
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Меняет пользовательское значение системной настройки по ключу текущей именованной таблицы.
Private Sub TestSetSetting(ByVal key As String, ByVal value As String)
    Dim range As Object, table As CSectionInputTable
    Set range = ThisWorkbook.Names("rngSystemSettings").RefersToRange
    Set table = New CSectionInputTable: table.Initialize "rngSystemSettings", range.Value2
    range.Cells(table.RowByKey(key), 2).Value2 = value
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Находит столбец PR1 по заголовку и меняет указанную строку; отсутствие профиля или ключа является ошибкой подготовки.
Private Sub TestSetProfile(ByVal key As String, ByVal value As String)
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

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Учитывает результат условия и сохраняет имя неудачной проверки в итоговом протоколе.
Private Sub TestCheck(ByVal name As String, ByVal value As Boolean)
    If value Then mPassed = mPassed + 1 Else mFailed = mFailed + 1: mReport = mReport & "FAIL: " & name & vbCrLf
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает величины с допуском 10^-9·(1 + |ожидаемое|), сохраняя оба числа в протоколе.
Private Sub TestNear(ByVal name As String, ByVal actual As Double, ByVal expected As Double)
    TestCheck name & "; actual=" & CStr(actual) & "; expected=" & CStr(expected), Abs(actual - expected) <= 0.000000001 * (1# + Abs(expected))
End Sub
