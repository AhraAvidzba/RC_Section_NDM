Attribute VB_Name = "modTestRecipeSetbacks"
Option Explicit

' ДЛЯ ТЕСТОВ: реальные редактируемые отступы Config и их влияние на общий
' рецепт, включая фильтр проекций Opening. НДС в этом наборе не решается.
Private mPassed As Long, mFailed As Long, mReport As String

' ДЛЯ ТЕСТОВ: те же внутренние координаты при вводе отступов в мм/см/м.
Public Function RunRecipeSetbackUnitTests() As String
    Dim values As Variant, baseline As Variant, scaled As Variant, table As CSectionInputTable
    Dim sheet As Object, setting As Object, settings As CSystemSettingsReader, units As CUnitSystem
    Dim recipe As ISectionRecipe, definition As CSectionDefinition, expected As CRebarLayout, actual As CRebarLayout
    Dim main As Long, extra As Long, row As Long, column As Long, i As Long, item As Variant, factor As Double
    Dim alerts As Boolean, itemColumn As Variant
    On Error GoTo Failed
    mPassed = 0: mFailed = 0: mReport = "": alerts = Application.DisplayAlerts
    values = ThisWorkbook.Names("rngRoundedRectangleGeometry").RefersToRange.Value2
    Set table = New CSectionInputTable: table.Initialize "rngRoundedRectangleGeometry", values
    main = table.RowByKey("as", 2): extra = table.RowByKey("d2", 2)
    values(3, 2) = 800#: values(3, 3) = 250#
    values(7, 2) = "Simple": values(7, 3) = "Tapered"
    values(8, 2) = 0#: values(8, 3) = 150#
    values(9, 2) = 40#: values(9, 3) = 40#: values(10, 2) = 0#: values(10, 3) = 80#
    For row = main + 1 To main + 4
        values(row, 2) = 40#: values(row, 3) = 20#: values(row, 4) = 3
        values(row, 5) = 10#: values(row, 6) = 20#
    Next row
    For row = extra + 1 To extra + 4: values(row, 2) = 0#: values(row, 5) = 0#: Next row
    Set recipe = New CSectionRecipeRoundedRectangle
    Set expected = FixtureAnnotatedLayout(FixtureRecipeDefinition(recipe, "rngRoundedRectangleGeometry", values))
    baseline = values
    Set sheet = ThisWorkbook.Worksheets.Add: Set setting = sheet.Range("A1:C2")
    setting.Cells(1, 1).Value2 = "Key": setting.Cells(1, 2).Value2 = "Value": setting.Cells(1, 3).Value2 = "Unit"
    setting.Cells(2, 1).Value2 = "Units.Length.Input"
    For Each item In Array(Array("mm", 1#), Array("cm", 10#), Array("m", 1000#))
        factor = CDbl(item(1)): scaled = baseline
        For column = 2 To 3
            scaled(3, column) = CDbl(baseline(3, column)) / factor
            For row = 8 To 10: scaled(row, column) = CDbl(baseline(row, column)) / factor: Next row
        Next column
        For row = main + 1 To main + 4
            For Each itemColumn In Array(2, 3, 5, 6)
                scaled(row, CLng(itemColumn)) = CDbl(baseline(row, CLng(itemColumn))) / factor
            Next itemColumn
        Next row
        setting.Cells(2, 2).Value2 = CStr(item(0))
        Set settings = New CSystemSettingsReader: settings.LoadFromRange setting
        Set units = New CUnitSystem: units.LoadFromSettings settings
        Set definition = FixtureRecipeDefinition(recipe, "rngRoundedRectangleGeometry", scaled, units)
        Set actual = FixtureAnnotatedLayout(definition)
        Check "setback.units.count." & CStr(item(0)), actual.Count = expected.Count
        For i = 1 To expected.Count
            Near "setback.units.x." & CStr(item(0)) & "." & CStr(i), actual.X(i), expected.X(i)
            Near "setback.units.y." & CStr(item(0)) & "." & CStr(i), actual.Y(i), expected.Y(i)
        Next i
    Next item
    mReport = mReport & "TOTAL_RECIPE_SETBACK_UNITS: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
Restore:
    On Error Resume Next
    Application.DisplayAlerts = False
    If Not sheet Is Nothing Then sheet.Delete
    Application.DisplayAlerts = alerts
    On Error GoTo 0
    RunRecipeSetbackUnitTests = mReport
    Exit Function
Failed:
    mFailed = mFailed + 1: mReport = mReport & "FAIL: units.runtime; " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
    Resume Restore
End Function

' ДЛЯ ТЕСТОВ: повторное обновление сохраняет введённые отступы и Excel-state;
' принудительный отказ именованного диапазона происходит до изменения данных.
Public Function RunRecipeConfigMigrationTests() As String
    Dim table As Object, snapshot As CSectionInputTable, saved As Variant, row As Long, ids As Variant, choices As Object
    Dim namesBefore As Long, i As Long, number As Long, reference As String, expected As Variant, after As Variant
    Dim calculation As Long, events As Boolean, updating As Boolean, alerts As Boolean
    On Error GoTo Failed
    mPassed = 0: mFailed = 0: mReport = ""
    calculation = Application.Calculation: events = Application.EnableEvents
    updating = Application.ScreenUpdating: alerts = Application.DisplayAlerts
    Set table = ThisWorkbook.Names("rngRoundedRectangleGeometry").RefersToRange: saved = table.Formula
    Set snapshot = New CSectionInputTable: snapshot.Initialize "rngRoundedRectangleGeometry", table.Value2
    row = snapshot.RowByKey("as", 2) + 1
    table.Cells(row, 5).Value2 = 37.5: table.Cells(row, 6).Value2 = 11.25
    expected = table.Formula: namesBefore = ThisWorkbook.Names.Count
    Application.Calculation = xlCalculationManual: Application.EnableEvents = False
    Application.ScreenUpdating = False: Application.DisplayAlerts = False
    modSectionRecipeSetup.UpdateSectionRecipeConfig
    Near "migration.nonzero.start.saved", table.Cells(row, 5).Value2, 37.5
    Near "migration.nonzero.end.saved", table.Cells(row, 6).Value2, 11.25
    Check "migration.names.idempotent", ThisWorkbook.Names.Count = namesBefore
    after = table.Formula
    For i = 1 To UBound(expected, 1)
        Check "migration.row.unchanged." & CStr(i), SameInputRow(expected, after, i)
    Next i
    ids = RegisteredSectionRecipeTypes: Set choices = ThisWorkbook.Names("rngSectionRecipeTypes").RefersToRange
    Check "migration.registry.choice.count", choices.Rows.Count = UBound(ids) - LBound(ids) + 1
    For i = 1 To choices.Rows.Count
        Check "migration.registry.choice." & CStr(i), choices.Cells(i, 1).Value2 = ids(LBound(ids) + i - 1)
    Next i
    Check "migration.state.success", Application.Calculation = xlCalculationManual And Not Application.EnableEvents And _
        Not Application.ScreenUpdating And Not Application.DisplayAlerts
    reference = ThisWorkbook.Names("rngRoundedRectangleGeometry").RefersTo
    ThisWorkbook.Names("rngRoundedRectangleGeometry").RefersTo = "=#REF!"
    On Error Resume Next
    modSectionRecipeSetup.UpdateSectionRecipeConfig
    number = Err.Number: Err.Clear
    On Error GoTo Failed
    ThisWorkbook.Names("rngRoundedRectangleGeometry").RefersTo = reference: reference = ""
    Check "migration.invalid.name.rejected", number <> 0
    Check "migration.state.failure", Application.Calculation = xlCalculationManual And Not Application.EnableEvents And _
        Not Application.ScreenUpdating And Not Application.DisplayAlerts
    after = table.Formula
    For i = 1 To UBound(expected, 1)
        Check "migration.failure.row.unchanged." & CStr(i), SameInputRow(expected, after, i)
    Next i
    mReport = mReport & "TOTAL_RECIPE_CONFIG_MIGRATION: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
Restore:
    On Error Resume Next
    If Len(reference) > 0 Then ThisWorkbook.Names("rngRoundedRectangleGeometry").RefersTo = reference
    If Not table Is Nothing Then table.Formula = saved
    Application.Calculation = calculation: Application.EnableEvents = events
    Application.ScreenUpdating = updating: Application.DisplayAlerts = alerts
    On Error GoTo 0
    RunRecipeConfigMigrationTests = mReport
    Exit Function
Failed:
    mFailed = mFailed + 1: mReport = mReport & "FAIL: migration.runtime; " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
    Resume Restore
End Function

Private Function SameInputRow(ByVal first As Variant, ByVal second As Variant, ByVal row As Long) As Boolean
    Dim column As Long
    For column = 1 To UBound(first, 2)
        If VarType(first(row, column)) <> VarType(second(row, column)) Then Exit Function
        If CStr(first(row, column)) <> CStr(second(row, column)) Then Exit Function
    Next column
    SameInputRow = True
End Function

Public Function RunRecipeSetbackInputTests() As String
    Dim rounded As Object, hollow As Object, savedRounded As Variant, savedHollow As Variant
    Dim table As CSectionInputTable, row As Long, main As Long, def As CSectionDefinition
    Dim original As CRebarLayout, changed As CRebarLayout, region As CConcreteRegion
    Dim face As Variant, faces As Variant, helper As CRebarRuleExecutor, originalHelper As CRebarRuleExecutor, parameters As Variant
    Dim geometryArea As Double, layoutCount As Long, reference As String, reason As String, number As Long
    On Error GoTo Failed
    mPassed = 0: mFailed = 0: mReport = ""
    Set rounded = ThisWorkbook.Names("rngRoundedRectangleGeometry").RefersToRange
    Set hollow = ThisWorkbook.Names("rngHollowRectangleGeometry").RefersToRange
    savedRounded = rounded.Formula: savedHollow = hollow.Formula
    Set table = New CSectionInputTable: table.Initialize "rngRoundedRectangleGeometry", rounded.Value2
    main = table.RowByKey("as", 2): row = table.RowByAnyKey(Array("B - верхняя", "B.Top"), 1, main + 1, main + 4)
    Check "rounded.start.header", table.ColumnByCaption(main, "StartSetback") = 5
    Check "rounded.end.header", table.ColumnByCaption(main, "EndSetback") = 6
    Check "rounded.editable.centered", rounded.Cells(row, 5).HorizontalAlignment = xlCenter And rounded.Cells(row, 6).HorizontalAlignment = xlCenter
    face = Array(40#, 20#, 3, 0#, 0#, "Stacked", "Stacked", "EachBar", "EachBar")
    faces = Array(face, face, Array(40#, 20#, 5, 0#, 0#, "Stacked", "Stacked", "EachBar", "EachBar"), face)
    parameters = Array(800#, 250#, "Simple", "Tapered", 0#, 150#, 40#, 40#, 0#, 80#)
    rounded.Cells(main + 1, 5).Resize(4, 2).Value2 = 0#
    Set def = FixtureRoundedDefinition(parameters, faces): Set original = FixtureAnnotatedLayout(def)
    geometryArea = def.ConcreteDomain.Area: layoutCount = original.Count
    Near "rounded.zero.first", original.X(1), -360#
    rounded.Cells(row, 5).Value2 = 10#: rounded.Cells(row, 6).Value2 = 20#
    Set def = FixtureRoundedDefinition(parameters, faces): Set changed = FixtureAnnotatedLayout(def)
    Near "rounded.start.actual", changed.X(1), original.X(1) + 10#
    Near "rounded.end.actual", changed.X(5), original.X(5) - 20#
    Near "rounded.concrete.unchanged", def.ConcreteDomain.Area, geometryArea
    Check "rounded.count.unchanged", changed.Count = layoutCount
    Near "rounded.opposite.row.unchanged", changed.Y(6), original.Y(6)
    rounded.Cells(row, 5).ClearContents
    Set changed = Nothing: number = 0
    On Error Resume Next
    Set changed = FixtureAnnotatedLayout(FixtureRoundedDefinition(parameters, faces))
    number = Err.Number: reason = Err.Description: Err.Clear
    On Error GoTo Failed
    Check "rounded.active.empty.rejected", number <> 0 And changed Is Nothing
    rounded.Cells(row, 5).Value2 = -1#
    On Error Resume Next
    Set changed = FixtureAnnotatedLayout(FixtureRoundedDefinition(parameters, faces))
    number = Err.Number: reason = Err.Description: Err.Clear
    On Error GoTo Failed
    Check "rounded.active.negative.rejected", number <> 0 And InStr(1, reason, "[" & CStr(row) & ",5]", vbBinaryCompare) > 0
    rounded.Cells(row, 5).Value2 = CVErr(xlErrValue)
    faces(2) = Array(40#, 0#, 0, 0#, 0#, "Stacked", "Stacked", "EachBar", "EachBar")
    Set changed = FixtureAnnotatedLayout(FixtureRoundedDefinition(parameters, faces))
    Check "rounded.inactive.error.ignored", changed.Count = layoutCount - 5
    rounded.Formula = savedRounded

    Set table = New CSectionInputTable: table.Initialize "rngHollowRectangleGeometry", hollow.Value2
    main = table.RowByKey("as", 2)
    row = table.RowByAnyKey(Array("B_o - верхняя", "Opening.B.Top"), 1, main + 1, main + 8)
    Check "hollow.start.header", table.ColumnByCaption(main, "StartSetback") = 5
    Check "hollow.end.header", table.ColumnByCaption(main, "EndSetback") = 6
    hollow.Cells(main + 1, 5).Resize(8, 2).Value2 = 0#
    face = Array(40#, 20#, 6, 0#, 0#, "Stacked", "Stacked", "EachBar", "EachBar")
    faces = Array(face, face, Array(40#, 20#, 8, 0#, 0#, "Stacked", "Stacked", "EachBar", "EachBar"), _
        Array(40#, 20#, 8, 0#, 0#, "Stacked", "Stacked", "EachBar", "EachBar"), face, face, face, face)
    parameters = Array(500#, 800#, 0#, 200#, 500#, 0#)
    Set originalHelper = New CRebarRuleExecutor
    Set def = FixtureHollowDefinition(parameters, faces): Set original = FixtureAnnotatedLayout(def, originalHelper)
    geometryArea = def.ConcreteDomain.Area
    Check "hollow.zero.projected.count", GroupCount(original, "Opening.B.Top") = 6
    hollow.Cells(row, 5).Value2 = 75#: hollow.Cells(row, 6).Value2 = 15#
    Set helper = New CRebarRuleExecutor
    Set def = FixtureHollowDefinition(parameters, faces): Set changed = FixtureAnnotatedLayout(def, helper)
    Check "hollow.projection.filtered", GroupCount(changed, "Opening.B.Top") = 4
    Check "hollow.skips.diagnostic", helper.SkippedProjections - originalHelper.SkippedProjections = 2
    Check "hollow.positions.not.redistributed", ProjectedSubset(changed, original, "Opening.B.Top")
    Check "hollow.other.projections.unchanged", GroupCount(changed, "Opening.B.Bottom") = GroupCount(original, "Opening.B.Bottom")
    Near "hollow.concrete.unchanged", def.ConcreteDomain.Area, geometryArea
    hollow.Formula = savedHollow
    mReport = mReport & "TOTAL_RECIPE_SETBACK_INPUT: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
Restore:
    On Error Resume Next
    If Not rounded Is Nothing Then rounded.Formula = savedRounded
    If Not hollow Is Nothing Then hollow.Formula = savedHollow
    On Error GoTo 0
    RunRecipeSetbackInputTests = mReport
    Exit Function
Failed:
    mFailed = mFailed + 1: mReport = mReport & "FAIL: runtime; " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
    Resume Restore
End Function

Private Function GroupCount(ByVal layout As CRebarLayout, ByVal group As String) As Long
    Dim i As Long
    For i = 1 To layout.Count
        If InGroup(layout, i, group) Then GroupCount = GroupCount + 1
    Next i
End Function

Private Function ProjectedSubset(ByVal actual As CRebarLayout, ByVal original As CRebarLayout, ByVal group As String) As Boolean
    Dim i As Long, j As Long, found As Boolean
    For i = 1 To actual.Count
        If InGroup(actual, i, group) Then
            found = False
            For j = 1 To original.Count
                If InGroup(original, j, group) Then
                    If Abs(actual.X(i) - original.X(j)) < 0.000001 And Abs(actual.Y(i) - original.Y(j)) < 0.000001 Then found = True
                End If
            Next j
            If Not found Then Exit Function
        End If
    Next i
    ProjectedSubset = True
End Function

Private Function InGroup(ByVal layout As CRebarLayout, ByVal index As Long, ByVal group As String) As Boolean
    Dim membership As Variant
    For Each membership In layout.BarMemberships(index)
        If CStr(membership) = group Then InGroup = True: Exit Function
    Next membership
End Function

Private Sub Check(ByVal name As String, ByVal condition As Boolean)
    If condition Then
        mPassed = mPassed + 1: mReport = mReport & "OK: " & name & vbCrLf
    Else
        mFailed = mFailed + 1: mReport = mReport & "FAIL: " & name & vbCrLf
    End If
End Sub

Private Sub Near(ByVal name As String, ByVal actual As Double, ByVal expected As Double)
    Check name & "; actual=" & CStr(actual) & "; expected=" & CStr(expected), Abs(actual - expected) <= 0.000001
End Sub
