Attribute VB_Name = "modTestRecipeSetbacks"
Option Explicit

' ДЛЯ ТЕСТОВ: реальные редактируемые отступы Config и их влияние на общий
' рецепт, включая фильтр проекций Opening. НДС в этом наборе не решается.
Private mPassed As Long, mFailed As Long, mReport As String

' ДЛЯ ТЕСТОВ: те же внутренние координаты при вводе отступов в мм/см/м.
Public Function RunRecipeSetbackUnitTests() As String
    Dim values As Variant, baseline As Variant, scaled As Variant, table As CSectionInputReader
    Dim sheet As Object, setting As Object, settings As CSystemSettingsReader, units As CUnitSystem
    Dim recipe As ISectionRecipe, definition As CSectionDefinition, expected As CRebarLayout, actual As CRebarLayout
    Dim main As Long, extra As Long, row As Long, column As Long, i As Long, item As Variant, factor As Double
    Dim alerts As Boolean, itemColumn As Variant
    On Error GoTo Failed
    mPassed = 0: mFailed = 0: mReport = "": alerts = Application.DisplayAlerts
    values = ThisWorkbook.Names("rngRoundedRectangleGeometry").RefersToRange.Value2
    FixtureSetInput values, "Geom", "B", "val", 800#: FixtureSetInput values, "Geom", "H", "val", 250#
    FixtureSetInput values, "Detail", "Type", "left", "Simple": FixtureSetInput values, "Detail", "Type", "right", "Tapered"
    FixtureSetInput values, "Detail", "W", "left", 0#: FixtureSetInput values, "Detail", "W", "right", 150#
    FixtureSetInput values, "Detail", "R1", "left", 40#: FixtureSetInput values, "Detail", "R1", "right", 40#
    FixtureSetInput values, "Detail", "R2", "left", 0#: FixtureSetInput values, "Detail", "R2", "right", 80#
    Dim key As Variant, block As Variant, fields As Variant, field As Variant, rowKeys As Variant
    Set table = FixtureInput(values)
    For Each key In table.RowKeys("Main")
        FixtureSetInput values, "Main", CStr(key), "as", 40#: FixtureSetInput values, "Main", CStr(key), "d", 20#
        FixtureSetInput values, "Main", CStr(key), "n", 3
        FixtureSetInput values, "Main", CStr(key), "s0", 10#: FixtureSetInput values, "Main", CStr(key), "s1", 20#
        FixtureSetInput values, "Extra", CStr(key), "d2", 0#: FixtureSetInput values, "Extra", CStr(key), "d3", 0#
    Next key
    Set recipe = New CSectionRecipeRoundedRectangle
    Set expected = FixtureAnnotatedLayout(FixtureRecipeDefinition(recipe, "rngRoundedRectangleGeometry", values))
    baseline = values
    Set sheet = ThisWorkbook.Worksheets.Add: Set setting = sheet.Range("A1:C2")
    setting.Cells(1, 1).Value2 = "Key": setting.Cells(1, 2).Value2 = "Value": setting.Cells(1, 3).Value2 = "Unit"
    setting.Cells(2, 1).Value2 = "Units.Length.Input"
    For Each item In Array(Array("mm", 1#), Array("cm", 10#), Array("m", 1000#))
        factor = CDbl(item(1)): scaled = baseline
        For Each block In Array("Geom", "Detail", "Main", "Extra")
            Select Case CStr(block)
                Case "Geom": fields = Array("val")
                Case "Detail": fields = Array("left", "right")
                Case "Main": fields = Array("as", "d", "s0", "s1")
                Case "Extra": fields = Array("d2", "a2", "d3", "a3")
            End Select
            For Each key In table.RowKeys(CStr(block))
                If CStr(block) <> "Detail" Or CStr(key) <> "Type" Then
                    row = table.RowIndex(CStr(block), CStr(key))
                    For Each field In fields
                        column = table.ColumnIndex(CStr(block), CStr(field))
                        If IsNumeric(baseline(row, column)) Then scaled(row, column) = CDbl(baseline(row, column)) / factor
                    Next field
                End If
            Next key
        Next block
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

' ДЛЯ ТЕСТОВ: обновление реестра не перестраивает Config и не требует таблиц геометрии.
Public Function RunRecipeTypeListTests() As String
    Dim config As Object, saved As Variant, expected As Variant, list As Object, ids As Variant
    Dim originalList As String, originalCircle As String, brokenCircle As String, i As Long, count As Long, code As Long
    Dim calculation As Long, events As Boolean, updating As Boolean, alerts As Boolean
    On Error GoTo Failed
    mPassed = 0: mFailed = 0: mReport = ""
    calculation = Application.Calculation: events = Application.EnableEvents
    updating = Application.ScreenUpdating: alerts = Application.DisplayAlerts
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = config.Formula
    originalList = ThisWorkbook.Names("rngSectionRecipeTypes").RefersTo
    originalCircle = ThisWorkbook.Names("rngCircleGeometry").RefersTo
    Set list = ThisWorkbook.Names("rngSectionRecipeTypes").RefersToRange
    ids = RegisteredSectionRecipeTypes: count = UBound(ids) - LBound(ids) + 1
    Application.Calculation = xlCalculationManual: Application.EnableEvents = False
    Application.ScreenUpdating = False: Application.DisplayAlerts = False
    ' Даже повреждённая чужая таблица не участвует в обновлении списка.
    ThisWorkbook.Names("rngCircleGeometry").RefersTo = "=#REF!"
    brokenCircle = ThisWorkbook.Names("rngCircleGeometry").RefersTo
    list.Value2 = "stale choice": expected = config.Formula
    modSectionRecipeSetup.UpdateSectionRecipeTypes
    Check "types.config.outside.list", SameConfigOutsideTypeList(config, expected, list.Resize(count, 1))
    Check "types.current.selection", SameConfigOutsideTypeList(config, saved, list.Resize(count, 1))
    Check "types.unrelated.name", ThisWorkbook.Names("rngCircleGeometry").RefersTo = brokenCircle
    For i = 1 To count: Check "types.registered." & CStr(i), list.Cells(i, 1).Value2 = ids(LBound(ids) + i - 1): Next i
    Check "types.state.success", Application.Calculation = xlCalculationManual And Not Application.EnableEvents And Not Application.ScreenUpdating And Not Application.DisplayAlerts
    expected = config.Formula: modSectionRecipeSetup.UpdateSectionRecipeTypes
    Check "types.idempotent", SameConfigOutsideTypeList(config, expected, Nothing)
    ' Сокращённый список расширяется только на свободные ячейки своего столбца.
    list.Cells(count, 1).ClearContents
    ThisWorkbook.Names("rngSectionRecipeTypes").RefersTo = "=" & list.Resize(count - 1, 1).Address(True, True, xlA1, True)
    modSectionRecipeSetup.UpdateSectionRecipeTypes
    Check "types.grow", ThisWorkbook.Names("rngSectionRecipeTypes").RefersToRange.Rows.Count = count And list.Cells(count, 1).Value2 = ids(UBound(ids))
    ' Занятое продолжение отклоняется до записи значений и имени.
    ThisWorkbook.Names("rngSectionRecipeTypes").RefersTo = "=" & list.Resize(count - 1, 1).Address(True, True, xlA1, True)
    list.Cells(count, 1).Value2 = "user data": expected = config.Formula
    On Error Resume Next
    modSectionRecipeSetup.UpdateSectionRecipeTypes: code = Err.Number: Err.Clear
    On Error GoTo Failed
    Check "types.occupied.rejected", code <> 0
    Check "types.occupied.data", SameConfigOutsideTypeList(config, expected, Nothing)
    Check "types.occupied.name", ThisWorkbook.Names("rngSectionRecipeTypes").RefersToRange.Rows.Count = count - 1
    Check "types.state.failure", Application.Calculation = xlCalculationManual And Not Application.EnableEvents And Not Application.ScreenUpdating And Not Application.DisplayAlerts
    ' Лишняя прежняя запись удаляется только из принадлежащего списку диапазона.
    ThisWorkbook.Names("rngSectionRecipeTypes").RefersTo = "=" & list.Resize(count + 1, 1).Address(True, True, xlA1, True)
    list.Cells(count + 1, 1).Value2 = "removed type": modSectionRecipeSetup.UpdateSectionRecipeTypes
    Check "types.shrink", ThisWorkbook.Names("rngSectionRecipeTypes").RefersToRange.Rows.Count = count And IsEmpty(list.Cells(count + 1, 1).Value2)
    expected = config.Formula: ThisWorkbook.Names("rngSectionRecipeTypes").RefersTo = "=#REF!": code = 0
    On Error Resume Next
    modSectionRecipeSetup.UpdateSectionRecipeTypes: code = Err.Number: Err.Clear
    On Error GoTo Failed
    Check "types.missing.rejected", code <> 0
    Check "types.missing.data", SameConfigOutsideTypeList(config, expected, Nothing)
    GoTo Restore
Failed:
    Check "types.runtime; " & CStr(Err.Number) & "; " & Err.Description, False
Restore:
    On Error Resume Next
    If Len(originalList) > 0 Then ThisWorkbook.Names("rngSectionRecipeTypes").RefersTo = originalList
    If Len(originalCircle) > 0 Then ThisWorkbook.Names("rngCircleGeometry").RefersTo = originalCircle
    If Not config Is Nothing And Not IsEmpty(saved) Then config.Formula = saved
    Application.Calculation = calculation: Application.EnableEvents = events
    Application.ScreenUpdating = updating: Application.DisplayAlerts = alerts
    On Error GoTo 0
    RunRecipeTypeListTests = mReport & "TOTAL RECIPE TYPE LIST: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' Сравнивает весь Config, кроме явно разрешённого служебного списка.
Private Function SameConfigOutsideTypeList(ByVal config As Object, ByVal expected As Variant, ByVal list As Object) As Boolean
    Dim actual As Variant, row As Long, column As Long, sheetRow As Long, sheetColumn As Long, skip As Boolean
    actual = config.Formula
    For row = 1 To UBound(expected, 1)
        For column = 1 To UBound(expected, 2)
            skip = False
            If Not list Is Nothing Then
                sheetRow = config.Row + row - 1: sheetColumn = config.Column + column - 1
                skip = sheetColumn = list.Column And sheetRow >= list.Row And sheetRow < list.Row + list.Rows.Count
            End If
            If Not skip Then
                If VarType(actual(row, column)) <> VarType(expected(row, column)) Then Exit Function
                If CStr(actual(row, column)) <> CStr(expected(row, column)) Then Exit Function
            End If
        Next column
    Next row
    SameConfigOutsideTypeList = True
End Function

Public Function RunRecipeSetbackInputTests() As String
    Dim rounded As Object, hollow As Object, savedRounded As Variant, savedHollow As Variant
    Dim table As CSectionInputReader, row As Long, main As Long, def As CSectionDefinition
    Dim original As CRebarLayout, changed As CRebarLayout, region As CConcreteRegion
    Dim face As Variant, faces As Variant, helper As CRebarRuleExecutor, originalHelper As CRebarRuleExecutor, parameters As Variant
    Dim geometryArea As Double, layoutCount As Long, reference As String, reason As String, number As Long
    On Error GoTo Failed
    mPassed = 0: mFailed = 0: mReport = ""
    Set rounded = ThisWorkbook.Names("rngRoundedRectangleGeometry").RefersToRange
    Set hollow = ThisWorkbook.Names("rngHollowRectangleGeometry").RefersToRange
    savedRounded = rounded.Formula: savedHollow = hollow.Formula
    Set table = FixtureInput(rounded.Value2)
    row = table.RowIndex("Main", "Top"): main = row - 1
    Check "rounded.start.header", table.ColumnIndex("Main", "s0") = 6
    Check "rounded.end.header", table.ColumnIndex("Main", "s1") = 7
    Check "rounded.editable.centered", rounded.Cells(row, 6).HorizontalAlignment = xlCenter And rounded.Cells(row, 7).HorizontalAlignment = xlCenter
    face = Array(40#, 20#, 3, 0#, 0#, "Stacked", "Stacked", "EachBar", "EachBar")
    faces = Array(face, face, Array(40#, 20#, 5, 0#, 0#, "Stacked", "Stacked", "EachBar", "EachBar"), face)
    parameters = Array(800#, 250#, "Simple", "Tapered", 0#, 150#, 40#, 40#, 0#, 80#)
    rounded.Cells(main + 1, 6).Resize(4, 2).Value2 = 0#
    Set def = FixtureRoundedDefinition(parameters, faces): Set original = FixtureAnnotatedLayout(def)
    geometryArea = def.ConcreteDomain.Area: layoutCount = original.Count
    Near "rounded.zero.first", original.X(1), -360#
    rounded.Cells(row, 6).Value2 = 10#: rounded.Cells(row, 7).Value2 = 20#
    Set def = FixtureRoundedDefinition(parameters, faces): Set changed = FixtureAnnotatedLayout(def)
    Near "rounded.start.actual", changed.X(1), original.X(1) + 10#
    Near "rounded.end.actual", changed.X(5), original.X(5) - 20#
    Near "rounded.concrete.unchanged", def.ConcreteDomain.Area, geometryArea
    Check "rounded.count.unchanged", changed.Count = layoutCount
    Near "rounded.opposite.row.unchanged", changed.Y(6), original.Y(6)
    rounded.Cells(row, 6).ClearContents
    Set changed = Nothing: number = 0
    On Error Resume Next
    Set changed = FixtureAnnotatedLayout(FixtureRoundedDefinition(parameters, faces))
    number = Err.Number: reason = Err.Description: Err.Clear
    On Error GoTo Failed
    Check "rounded.active.empty.rejected", number <> 0 And changed Is Nothing
    rounded.Cells(row, 6).Value2 = -1#
    On Error Resume Next
    Set changed = FixtureAnnotatedLayout(FixtureRoundedDefinition(parameters, faces))
    number = Err.Number: reason = Err.Description: Err.Clear
    On Error GoTo Failed
    Check "rounded.active.negative.rejected", number <> 0 And InStr(1, reason, "[" & CStr(row) & ",6]", vbBinaryCompare) > 0
    rounded.Cells(row, 6).Value2 = CVErr(xlErrValue)
    faces(2) = Array(40#, 0#, 0, 0#, 0#, "Stacked", "Stacked", "EachBar", "EachBar")
    Set changed = FixtureAnnotatedLayout(FixtureRoundedDefinition(parameters, faces))
    Check "rounded.inactive.error.ignored", changed.Count = layoutCount - 5
    rounded.Formula = savedRounded

    Set table = FixtureInput(hollow.Value2)
    main = table.RowIndex("Main", "Top") - 1
    row = table.RowIndex("Main", "OpeningTop")
    Check "hollow.start.header", table.ColumnIndex("Main", "s0") = 6
    Check "hollow.end.header", table.ColumnIndex("Main", "s1") = 7
    hollow.Cells(main + 1, 6).Resize(8, 2).Value2 = 0#
    face = Array(40#, 20#, 6, 0#, 0#, "Stacked", "Stacked", "EachBar", "EachBar")
    faces = Array(face, face, Array(40#, 20#, 8, 0#, 0#, "Stacked", "Stacked", "EachBar", "EachBar"), _
        Array(40#, 20#, 8, 0#, 0#, "Stacked", "Stacked", "EachBar", "EachBar"), face, face, face, face)
    parameters = Array(500#, 800#, 0#, 200#, 500#, 0#)
    Set originalHelper = New CRebarRuleExecutor
    Set def = FixtureHollowDefinition(parameters, faces): Set original = FixtureAnnotatedLayout(def, originalHelper)
    geometryArea = def.ConcreteDomain.Area
    Check "hollow.zero.projected.count", GroupCount(original, "Opening.B.Top") = 6
    hollow.Cells(row, 6).Value2 = 75#: hollow.Cells(row, 7).Value2 = 15#
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
