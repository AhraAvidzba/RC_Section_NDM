Attribute VB_Name = "modTestProfileConfig"
Option Explicit

' ==========================================================================
' ДЛЯ ТЕСТОВ: профили и длительные нагрузки устойчивости
' ==========================================================================
' Проверяет ячейки PR1–PR4, выбор материалов и сохранённых результатов схемы.
' Ожидаемые материалы задаёт независимо; равновесие проверяет по элементам.
' Дополнительные нагрузки читает обычным способом. Ошибка строки должна
' сохраняться в результате её устойчивости. Все затронутые таблицы
' восстанавливаются также при ошибке; рабочие допуски не меняются.

Private Type TProfileStats
    Passed As Long
    Failed As Long
    Cases As Long ' Варианты поля/профиля, не число assertions.
    Report As String
End Type

' -------------------- ДЛЯ ТЕСТОВ --------------------
' ============================== ДЛЯ ТЕСТОВ ==============================

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Выполняет цельную приемку material spec, метаданных и визуализации профилей.
' Optional-счетчики включают те же assertions в штатную workbook suite.
Public Function RunAudit03ProfileConfigTests(Optional ByRef passed As Long = 0, _
        Optional ByRef failed As Long = 0) As String
    Dim stats As TProfileStats, names As Variant, saved(0 To 4) As Variant
    Dim source As Object, system As Object, unitsRange As Object, signs As Object, loads As Object
    Dim i As Long, section As CSectionModel, provider As CMaterialModelProvider
    Dim settings As CSystemSettingsReader, units As CUnitSystem
    On Error GoTo FailedRun
    names = Array("rngCalculationProfiles", "rngSystemSettings", "rngUnitSettings", _
        "rngSignConventionSettings", "rngLoadCombinations")
    For i = 0 To UBound(names)
        saved(i) = ThisWorkbook.Names.Item(CStr(names(i))).RefersToRange.Formula
    Next i
    Set source = ThisWorkbook.Names.Item("rngCalculationProfiles").RefersToRange
    Set system = ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange
    Set unitsRange = ThisWorkbook.Names.Item("rngUnitSettings").RefersToRange
    Set signs = ThisWorkbook.Names.Item("rngSignConventionSettings").RefersToRange
    Set loads = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange
    TestSetTableValue unitsRange, "Length", "mm", 2
    TestSetTableValue unitsRange, "Force", "N", 2
    TestSetTableValue unitsRange, "Moment", "N*mm", 2
    TestSetTableValue unitsRange, "Stress", "MPa", 2
    TestSetTableValue signs, "+N", "Tension", 2
    TestSetTableValue signs, "+Mx", "+Y tension", 2
    TestSetTableValue signs, "+My", "+X tension", 2
    TestSetTableValue system, "Calculation.ZeroMomentPerDepth", "0", 2
    TestSetTableValue system, "Load.ReferenceOffsetX", "0", 2
    TestSetTableValue system, "Load.ReferenceOffsetY", "0", 2
    TestSetTableValue system, "Solver.ToleranceN", "0.01", 2
    TestSetTableValue system, "Solver.ToleranceMx", "1", 2
    TestSetTableValue system, "Solver.ToleranceMy", "1", 2
    TestSetTableValue system, "SLS.Crack.PsiMode", "AlwaysCalc", 2
    TestSetTableValue system, "Capacity.SolutionStrategy", "Auto", 2
    TestSetTableValue system, "Stability.Code", "SP35", 2
    TestSetTableValue system, "Stability.ElementLength", "1000", 2
    TestSetTableValue system, "Stability.Mu1", "1", 2
    TestSetTableValue system, "Stability.Mu2", "1", 2
    TestSetTableValue system, "Stability.AccidentalEccentricityMode", "User", 2
    TestSetTableValue system, "Stability.AccidentalEccentricityPlanes", "BothPlanes", 2
    TestSetTableValue system, "Stability.AccidentalEccentricityUser1", "200", 2
    TestSetTableValue system, "Stability.AccidentalEccentricityUser2", "200", 2
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    TestBuildFixture section, provider, settings.GetBoolean("General.DiagramExtension", False)
    TestLogLine stats, "PROFILE_CONFIG: effectiveExtension=" & CStr(provider.DiagramExtensionEnabled)
    TestMaterialSpecs stats, source, loads, section, provider, settings, units
    TestStabilityValueSets stats, source, loads, section, provider, settings, units
    TestProfileMetadata stats, source, loads, section, provider, settings, units
    TestVisualization stats, source, system, loads, section, provider, settings, units
    TestInvalidProfileCells stats, source
    GoTo Restore
FailedRun:
    stats.Failed = stats.Failed + 1
    TestLogLine stats, "FAIL: profileConfig.runtime; " & CStr(Err.Number) & "; " & Err.Description
Restore:
    On Error Resume Next
    For i = 0 To UBound(names)
        ThisWorkbook.Names.Item(CStr(names(i))).RefersToRange.Formula = saved(i)
    Next i
    On Error GoTo 0
    TestLogLine stats, "TOTAL_AUDIT03_PROFILE_CONFIG: passed=" & CStr(stats.Passed) & _
        "; failed=" & CStr(stats.Failed) & "; cases=" & CStr(stats.Cases)
    passed = stats.Passed: failed = stats.Failed
    RunAudit03ProfileConfigTests = stats.Report
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет смену абсолютных адресов при переадресации именованных таблиц.
' Исходные диапазоны и их формулы остаются на месте; измененные Names и вход
' восстанавливаются до удаления созданного только для теста листа.
Public Function RunAudit03RelocatedInputTests(Optional ByRef passed As Long = 0, _
        Optional ByRef failed As Long = 0) As String
    Dim stats As TProfileStats, sheet As Object, source As Object, target As Object, cell As Object
    Dim names As Variant, refs() As String, baseline As Variant, field As Variant, fields As Collection
    Dim profiles As Object, loads As Object, savedProfiles As Variant, savedLoads As Variant
    Dim index As Long, position As Long, code As Long, reason As String, prefix As String, savedValue As Variant
    Dim settings As CSystemSettingsReader, units As CUnitSystem, section As CSectionModel
    Dim provider As CMaterialModelProvider, batch As CBatchSectionCalculator, oldAlerts As Boolean
    Dim caption As String, address As String, oldAddress As String, key As String
    Dim config As Object, savedConfig As Variant
    On Error GoTo FailedRun
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: savedConfig = TestFixtureSaveInputs(config)
    modTestShapeArchitecture.TestConfigureShapeFixture "Circle"
    modTestShapeArchitecture.TestConfigureShapeFixture "RoundedMixed"
    modTestShapeArchitecture.TestConfigureShapeFixture "HollowCentered"
    modTestShapeArchitecture.TestConfigureShapeFixture "TwoRectangles"
    names = Array("rngSystemSettings", "rngUnitSettings", "rngSignConventionSettings", _
        "rngConcreteMaterialParameters", "rngSteelMaterialParameters", "rngCircleGeometry", _
        "rngPointedSectionGeometry", "rngHollowRectangleGeometry", "rngRectanglesGeometry", _
        "rngPlotAnnotationSettings", "rngCalculationProfiles", "rngLoadCombinations")
    ReDim refs(0 To UBound(names))
    For index = 0 To UBound(names): refs(index) = ThisWorkbook.Names.Item(CStr(names(index))).RefersTo: Next index
    Set profiles = ThisWorkbook.Names.Item("rngCalculationProfiles").RefersToRange
    Set loads = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange
    savedProfiles = profiles.Formula: savedLoads = loads.Formula
    TestConfigureProfiles profiles: TestEnableChecks profiles, 1, "Yes", "Yes", "Yes", "Yes"
    TestSetLoad loads, 1, -50000#, 1000000#, 2000000#, "Auto"
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    TestBuildFixture section, provider, settings.GetBoolean("General.DiagramExtension", False)
    oldAlerts = Application.DisplayAlerts
    Set sheet = ThisWorkbook.Worksheets.Add: sheet.Name = "__AuditMovedInputs"
    For index = 0 To UBound(names)
        Set source = ThisWorkbook.Names.Item(CStr(names(index))).RefersToRange
        baseline = source.Value2
        Set fields = TestRelocatedFields(source, CStr(names(index)))
        For position = 1 To 2
            If position = 1 Then Set target = sheet.Cells(10, 5) Else Set target = sheet.Cells(800, 60)
            Set target = target.Resize(source.Rows.Count, source.Columns.Count)
            target.Value2 = baseline
            ThisWorkbook.Names.Item(CStr(names(index))).RefersTo = "='" & sheet.Name & "'!" & target.Address
            For Each field In fields
                Set cell = target.Cells(CLng(field(0)), CLng(field(1)))
                key = CStr(field(2)): savedValue = cell.Value2
                address = cell.Address(False, False)
                oldAddress = source.Cells(CLng(field(0)), CLng(field(1))).Address(False, False)
                prefix = "relocated." & CStr(names(index)) & ".p" & CStr(position) & "." & key
                TestLogLine stats, "RUN: " & prefix & "; actual=" & sheet.Name & "!" & address
                If (index >= 5 And index <= 8) Then
                    Dim inputTable As CSectionInputTable
                    Set inputTable = TestRelocatedRecipeTable(CStr(names(index)), units)
                    reason = inputTable.SourceReference(CLng(field(0)), CLng(field(1)))
                    TestCheck stats, prefix & ".captured", inputTable.Name = CStr(names(index))
                    TestCheck stats, prefix & ".locationSheet", InStr(1, reason, sheet.Name, vbBinaryCompare) > 0
                    TestCheck stats, prefix & ".locationAddress", InStr(1, reason, address, vbBinaryCompare) > 0
                    cell.Value2 = CVErr(xlErrDiv0)
                    code = TestRelocatedRecipeError(CStr(names(index)), reason)
                ElseIf index < 10 Then
                    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
                    reason = settings.InputErrorMessage(key, "Контроль адреса.", "Проверьте значение.")
                    TestCheck stats, prefix & ".registered", settings.HasKey(key)
                    TestCheck stats, prefix & ".locationSheet", InStr(1, reason, sheet.Name, vbBinaryCompare) > 0
                    TestCheck stats, prefix & ".locationAddress", InStr(1, reason, "ячейка " & address, vbBinaryCompare) > 0
                    ' Ошибки геометрических значений хранятся до обращения
                    ' активного потребителя. Проверяем его чтение по тому же
                    ' ключу и сохраненный адрес перемещенной ячейки.
                    cell.Value2 = CVErr(xlErrDiv0): code = TestSettingsLoadError(reason, key)
                ElseIf index = 10 Then
                    cell.Value2 = "NOT_A_VALID_CHOICE": code = TestCatalogError(reason)
                Else
                    cell.Value2 = "abc"
                    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
                    Set batch = TestExecuteFixture(section, provider, settings, units)
                    code = IIf(batch.ResultAt(1).Status = "InputErr", 1, 0)
                    reason = batch.ResultAt(1).OverallMeta.ResultComment
                    TestCheck stats, prefix & ".noSolve", batch.SolverCallCount = 0
                    TestCheckComments stats, prefix, batch, units
                End If
                TestCheck stats, prefix & ".rejected", code <> 0
                TestCheck stats, prefix & ".errorSheet", InStr(1, reason, sheet.Name, vbBinaryCompare) > 0
                TestCheck stats, prefix & ".errorAddress", InStr(1, reason, address, vbBinaryCompare) > 0
                TestCheck stats, prefix & ".oldAddressAbsent", InStr(1, reason, "ячейка " & oldAddress, vbBinaryCompare) = 0
                TestLogLine stats, "RELOCATED_INPUT_MESSAGE: " & prefix & "; " & reason
                cell.Value2 = savedValue
                If index = 0 And key = "Solver.ToleranceN" Then
                    cell.Value2 = "abc"
                    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
                    Set batch = TestExecuteFixture(section, provider, settings, units)
                    reason = batch.ResultAt(1).OverallMeta.ResultComment
                    TestCheck stats, prefix & ".batchInputErr", batch.ResultAt(1).Status = "InputErr"
                    TestCheck stats, prefix & ".batchAddress", InStr(1, reason, "ячейка " & address, vbBinaryCompare) > 0
                    TestCheck stats, prefix & ".batchNoSolve", batch.SolverCallCount = 0
                    TestCheckComments stats, prefix & ".batch", batch, units
                    cell.Value2 = savedValue
                End If
                stats.Cases = stats.Cases + 1
            Next field
            If CStr(names(index)) = "rngRectanglesGeometry" Then TestRelocatedRectanglesSelectors stats, target, position
            ThisWorkbook.Names.Item(CStr(names(index))).RefersTo = refs(index)
        Next position
    Next index
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    TestCheck stats, "relocated.recovery.settings", settings.HasKey("Solver.ToleranceN")
    Dim catalog As CCalculationProfileCatalog
    Set catalog = New CCalculationProfileCatalog: catalog.LoadFromWorkbook ThisWorkbook
    TestCheck stats, "relocated.recovery.profiles", catalog.Count = 4
    GoTo Restore
FailedRun:
    stats.Failed = stats.Failed + 1
    TestLogLine stats, "FAIL: relocated.runtime; " & CStr(Err.Number) & "; " & Err.Description
Restore:
    On Error Resume Next
    For index = 0 To UBound(names)
        If Len(refs(index)) > 0 Then ThisWorkbook.Names.Item(CStr(names(index))).RefersTo = refs(index)
    Next index
    profiles.Formula = savedProfiles: loads.Formula = savedLoads
    If Not config Is Nothing And IsArray(savedConfig) Then TestFixtureRestoreInputs config, savedConfig
    Application.DisplayAlerts = False
    If Not sheet Is Nothing Then sheet.Delete
    Application.DisplayAlerts = oldAlerts
    On Error GoTo 0
    TestLogLine stats, "TOTAL_AUDIT03_RELOCATED_INPUT: passed=" & CStr(stats.Passed) & _
        "; failed=" & CStr(stats.Failed) & "; cases=" & CStr(stats.Cases)
    passed = stats.Passed: failed = stats.Failed
    RunAudit03RelocatedInputTests = stats.Report
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет независимые выборы и ошибки формул обеих физических граней.
' Адрес ввода берется из перемещенной таблицы; другая грань не подменяет
' ошибочное значение, а исходные настройки восстанавливаются после пробы.
Private Sub TestRelocatedRectanglesSelectors(ByRef stats As TProfileStats, ByVal source As Object, ByVal position As Long)
    Dim reader As CSectionInputReader, pair As Variant, field As Variant, rowKey As Variant
    Dim first As Object, second As Object, cell As Object, column As Long, row As Long
    Dim savedFirst As Variant, savedSecond As Variant, choices As Variant, code As Long, reason As String
    Dim inputTable As CSectionInputTable, units As CUnitSystem, value As String, prefix As String
    Set units = New CUnitSystem: units.InitializeDefaults
    Set reader = TestFixtureInput(source.Value2)
    For Each pair In Array(Array("R1.Left", "R1.Right"), Array("R1.Top", "R1.Bottom"), _
            Array("R2.Left", "R2.Right"), Array("R2.Top", "R2.Bottom"))
        For Each field In Array("pos2", "bind2", "pos3", "bind3")
            column = reader.ColumnIndex("Extra", CStr(field))
            Set first = source.Cells(reader.RowIndex("Extra", CStr(pair(0))), column)
            Set second = source.Cells(reader.RowIndex("Extra", CStr(pair(1))), column)
            savedFirst = first.Value2: savedSecond = second.Value2
            choices = Array("EachBar", "EverySecondBar")
            If Left$(CStr(field), 3) = "pos" Then choices = Array("Stacked", "SideBySide")
            first.Value2 = choices(0): second.Value2 = choices(1)
            prefix = "relocated.rectangles.p" & CStr(position) & "." & CStr(pair(0)) & "." & CStr(field)
            Set inputTable = TestRelocatedRecipeTable("rngRectanglesGeometry", units)
            TestCheck stats, prefix & ".firstIndependent", inputTable.Choice(first.Row - source.Row + 1, column, choices) = choices(0)
            TestCheck stats, prefix & ".secondIndependent", inputTable.Choice(second.Row - source.Row + 1, column, choices) = choices(1)
            For Each rowKey In pair
                row = reader.RowIndex("Extra", CStr(rowKey)): Set cell = source.Cells(row, column)
                cell.Value2 = CVErr(xlErrDiv0)
                Set inputTable = TestRelocatedRecipeTable("rngRectanglesGeometry", units)
                code = 0: reason = ""
                On Error Resume Next
                value = inputTable.Choice(row, column, choices)
                code = Err.Number: reason = Err.Description: Err.Clear
                On Error GoTo 0
                TestCheck stats, prefix & ".formulaRejected." & CStr(rowKey), code <> 0
                TestCheck stats, prefix & ".formulaAddress." & CStr(rowKey), InStr(1, reason, cell.Address(False, False), vbBinaryCompare) > 0
                TestCheck stats, prefix & ".formulaSheet." & CStr(rowKey), InStr(1, reason, source.Worksheet.Name, vbBinaryCompare) > 0
                first.Value2 = choices(0): second.Value2 = choices(1)
            Next rowKey
            first.Value2 = savedFirst: second.Value2 = savedSecond
            stats.Cases = stats.Cases + 1
        Next field
    Next pair
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' изменение геометрии проверяется через копию входной таблицы.
' Общий класс системных настроек не разбирает поля отдельных рецептов.
Private Function TestRelocatedRecipeTable(ByVal rangeName As String, ByVal units As CUnitSystem) As CSectionInputTable
    Dim reader As CSectionInputSnapshotReader, context As CSectionRecipeContext
    Set reader = New CSectionInputSnapshotReader
    Set context = reader.Capture(ThisWorkbook, Array(rangeName), units)
    Set TestRelocatedRecipeTable = context.Table(rangeName)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' ошибочное поле читает настоящий выбранный рецепт, а адрес
' берётся из перемещённого именованного диапазона. Выбор формы восстанавливается.
Private Function TestRelocatedRecipeError(ByVal rangeName As String, ByRef reason As String) As Long
    Dim system As Object, cell As Object, row As Long, saved As Variant, typeID As String
    Dim settings As CSystemSettingsReader, units As CUnitSystem, definition As CSectionDefinition
    On Error GoTo Failed
    Select Case rangeName
        Case "rngCircleGeometry": typeID = "Circle"
        Case "rngPointedSectionGeometry": typeID = "PointedSection"
        Case "rngHollowRectangleGeometry": typeID = "HollowRectangle"
        Case "rngRectanglesGeometry": typeID = "Rectangles"
    End Select
    Set system = ThisWorkbook.Names("rngSystemSettings").RefersToRange
    For row = 2 To system.Rows.Count
        If CStr(system.Cells(row, 1).Value2) = "Geometry.Type" Then Set cell = system.Cells(row, 2): Exit For
    Next row
    If cell Is Nothing Then Err.Raise 5, , "Не найден Geometry.Type тестовой книги."
    saved = cell.Formula: cell.Value2 = typeID
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    Set definition = ReadWorkbookSectionDefinition(ThisWorkbook, settings, units)
    reason = vbNullString
    GoTo Restore
Failed:
    TestRelocatedRecipeError = Err.Number: reason = Err.Description
Restore:
    On Error Resume Next
    If Not cell Is Nothing Then cell.Formula = saved
    On Error GoTo 0
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Задаёт ожидаемые ключи независимо от приватных методов чтения.
' Положение полей определяется по заголовкам без фиксированных адресов.
Private Function TestRelocatedFields(ByVal source As Object, ByVal rangeName As String) As Collection
    Dim fields As New Collection, row As Long, i As Long, column As Long
    Select Case rangeName
        Case "rngSystemSettings"
            fields.Add Array(TestCaptionRow(source, "Solver.ToleranceN", 1), 2, "Solver.ToleranceN")
            fields.Add Array(TestCaptionRow(source, "SLS.Crack.Allowable", 1), 2, "SLS.Crack.Allowable")
        Case "rngUnitSettings"
            fields.Add Array(TestCaptionRow(source, "Force", 1), 2, "Units.Force.Input")
            fields.Add Array(TestCaptionRow(source, "Force", 1), 4, "Units.Force.Output")
        Case "rngSignConventionSettings"
            fields.Add Array(TestCaptionRow(source, "+N", 1), 2, "Sign.N.User")
        Case "rngConcreteMaterialParameters"
            fields.Add Array(TestCaptionRow(source, "Concrete.R.ULS(I)", 1), 2, "Concrete.Rb.ULS")
            fields.Add Array(TestCaptionRow(source, "Concrete.R.ULS(I)", 1), 3, "Concrete.Rbt.ULS")
        Case "rngSteelMaterialParameters"
            fields.Add Array(TestCaptionRow(source, "Steel.E", 1), 2, "Steel.Esc")
            fields.Add Array(TestCaptionRow(source, "Steel.E", 1), 3, "Steel.Es")
        Case "rngCircleGeometry", "rngPointedSectionGeometry", "rngHollowRectangleGeometry", "rngRectanglesGeometry"
            Dim reader As CSectionInputReader, key As Variant, selected As Variant
            Set reader = TestFixtureInput(source.Value2)
            Select Case rangeName
                Case "rngCircleGeometry": selected = Array("D")
                Case "rngPointedSectionGeometry": selected = Array("B", "H", "Left.R1")
                Case "rngHollowRectangleGeometry": selected = Array("B", "H", "Xo")
                Case Else: selected = Array("Shape", "H1", "B1")
            End Select
            For Each key In selected
                fields.Add Array(reader.RowIndex("General", CStr(key)), reader.ColumnIndex("General", "val"), rangeName & "." & CStr(key))
            Next key
            Select Case rangeName
                Case "rngCircleGeometry": key = "Circle"
                Case "rngRectanglesGeometry": key = "R1.Left"
                Case Else: key = "Left"
            End Select
            Dim rebarField As Variant
            For Each rebarField In Array("as", "d", "n")
                fields.Add Array(reader.RowIndex("Main", CStr(key)), reader.ColumnIndex("Main", CStr(rebarField)), rangeName & ".Main." & CStr(rebarField))
            Next rebarField
        Case "rngPlotAnnotationSettings"
            fields.Add Array(TestCaptionRow(source, "Enabled", 1), 2, "Plot.RebarLabels.Enabled")
            fields.Add Array(TestCaptionRow(source, "TextHeight", 1), 3, "Plot.Dimensions.TextHeight")
        Case "rngCalculationProfiles"
            For i = 1 To 4
                fields.Add Array(TestCaptionRow(source, "MaterialModel.Strength.ValueSet", 2), i + 2, "PR" & CStr(i) & ".MaterialModel.Strength.ValueSet")
                fields.Add Array(TestCaptionRow(source, "Visualization.StressPrecision", 2), i + 2, "PR" & CStr(i) & ".Visualization.StressPrecision")
            Next i
        Case "rngLoadCombinations"
            fields.Add Array(2, 2, "N")
    End Select
    Set TestRelocatedFields = fields
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Ищет подпись внутри фактической таблицы и не подменяет отсутствующую строку.
Private Function TestCaptionRow(ByVal source As Object, ByVal caption As String, ByVal column As Long) As Long
    Dim row As Long
    For row = 1 To source.Rows.Count
        If CStr(source.Cells(row, column).Value2) = caption Then TestCaptionRow = row: Exit Function
    Next row
    Err.Raise vbObjectError + 4501, "CaptionRow", "В тестовой таблице нет " & caption
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сохраняет фактическую ошибку загрузки или активного чтения указанного поля
' после перемещения таблицы. Геометрический CVErr не мешает другой форме,
' но при обращении к нему обязан выдавать исходный фактический адрес.
Private Function TestSettingsLoadError(ByRef reason As String, Optional ByVal readKey As String = vbNullString) As Long
    On Error GoTo Failed
    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Dim value As String
    If Len(readKey) > 0 Then value = settings.GetRawString(readKey)
    reason = vbNullString
    Exit Function
Failed:
    TestSettingsLoadError = Err.Number: reason = Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Меняет каждый материал каждой роли во всех четырех профильных колонках.
' Capacity/Pre/Post остаются физическими, текущие State используют effective Extension.
Private Sub TestMaterialSpecs(ByRef stats As TProfileStats, ByVal source As Object, ByVal loads As Object, _
        ByVal section As CSectionModel, ByVal provider As CMaterialModelProvider, _
        ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem)
    Dim profile As Long, role As Variant, field As Variant, value As Variant, prefix As String
    Dim batch As CBatchSectionCalculator, result As CCombinationResult, expected As CMaterialModelSpec
    For profile = 1 To 4
        For Each role In Array("Strength", "CrackInitiation", "CrackedState")
            For Each field In Array("ValueSet", "ConcreteDiagram", "ConcreteTension", "SteelDiagram")
                For Each value In TestMaterialVariants(CStr(role), CStr(field))
                    TestConfigureProfiles source
                    TestSetProfileValue source, profile, "MaterialModel." & CStr(role) & "." & CStr(field), value
                    If CStr(role) = "Strength" Then
                        TestEnableChecks source, profile, "Yes", "Yes", "No", "No"
                        TestSetLoad loads, profile, -50000#, 10000000#, 2000000#, "LambdaMxy"
                    Else
                        TestEnableChecks source, profile, "No", "No", "Yes", "No"
                        TestSetLoad loads, profile, 200000#, 0#, 0#, "LambdaN"
                    End If
                    prefix = "profileConfig.PR" & CStr(profile) & ".MaterialModel." & CStr(role) & _
                        "." & CStr(field) & "." & CStr(value)
                    TestLogLine stats, "RUN: " & prefix
                    Set batch = TestExecuteFixture(section, provider, settings, units)
                    Set result = batch.ResultAt(1)
                    Set expected = TestExpectedSpec(CStr(role), CStr(field), CStr(value))
                    If CStr(role) = "Strength" Then
                        TestCheckState stats, prefix & ".direct", result.StateRepository.FindState(sstStrengthState), section, provider, expected
                        TestCheckState stats, prefix & ".capacity", result.StrengthResult.Capacity.StateResult, section, provider, expected
                    ElseIf CStr(role) = "CrackInitiation" Then
                        TestCheckState stats, prefix & ".pre", result.CrackResult.Formation.PreCrackState, section, provider, expected
                        TestCheckState stats, prefix & ".post", result.CrackResult.Formation.PostCrackState, section, provider, TestExpectedSpec("CrackedState", "", "")
                    Else
                        TestCheckState stats, prefix & ".current", result.StateRepository.FindState(sstCrackedState), section, provider, expected
                        TestCheckState stats, prefix & ".post", result.CrackResult.Formation.PostCrackState, section, provider, expected
                    End If
                    TestCheckComments stats, prefix, batch, units
                    stats.Cases = stats.Cases + 1
                Next value
            Next field
        Next role
    Next profile
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Подтверждает эффект набора сопротивлений в табличной ветви СП 35.
' Модули общие для ULS/SLS; здесь Nult зависит от выбранных Rb и Rsc.
Private Sub TestStabilityValueSets(ByRef stats As TProfileStats, ByVal source As Object, ByVal loads As Object, _
        ByVal section As CSectionModel, ByVal provider As CMaterialModelProvider, _
        ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem)
    Dim profile As Long, value As Variant, batch As CBatchSectionCalculator, baseline As Double, prefix As String
    TestSetTableValue ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange, "Stability.AccidentalEccentricityUser1", "5", 2
    TestSetTableValue ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange, "Stability.AccidentalEccentricityUser2", "5", 2
    settings.LoadFromWorkbook ThisWorkbook
    For profile = 1 To 4
        For Each value In Array("ULS(I)", "SLS(II)")
            TestConfigureProfiles source
            TestEnableChecks source, profile, "No", "No", "No", "Yes"
            TestSetProfileValue source, profile, "MaterialModel.Stability.ValueSet", value
            TestSetLoad loads, profile, -50000#, 0#, 0#, "Auto"
            Set batch = TestExecuteFixture(section, provider, settings, units)
            prefix = "profileConfig.PR" & CStr(profile) & ".MaterialModel.Stability.ValueSet." & CStr(value)
            TestCheck stats, prefix & ".branch", batch.ResultAt(1).StabilityResult.Branch = "SP35-table"
            TestCheck stats, prefix & ".calculated", batch.ResultAt(1).StabilityMeta.Calculated
            TestCheck stats, prefix & ".noSolve", batch.SolverCallCount = 0
            TestCheck stats, prefix & ".positiveLimit", batch.ResultAt(1).StabilityResult.Nultimate1 > 0#
            If CStr(value) = "ULS(I)" Then
                baseline = batch.ResultAt(1).StabilityResult.Nultimate1
            Else
                TestCheck stats, prefix & ".resistanceEffect", batch.ResultAt(1).StabilityResult.Nultimate1 > baseline
            End If
            TestCheckComments stats, prefix, batch, units
            stats.Cases = stats.Cases + 1
        Next value
    Next profile
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Пользовательское имя реально разрешает ProfileId строки LC. Описание является
' видимым пояснением профиля, а не физическим коэффициентом или новым calculator-ом.
Private Sub TestProfileMetadata(ByRef stats As TProfileStats, ByVal source As Object, ByVal loads As Object, _
        ByVal section As CSectionModel, ByVal provider As CMaterialModelProvider, _
        ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem)
    Dim profile As Long, catalog As CCalculationProfileCatalog, batch As CBatchSectionCalculator
    Dim caption As String, description As String, prefix As String, state As CSectionStateResult, data As Variant
    Dim value As Variant, code As Long, reason As String, cell As Object
    For profile = 1 To 4
        TestConfigureProfiles source
        TestEnableChecks source, profile, "Yes", "No", "No", "No"
        caption = "Проверяемый профиль " & CStr(profile)
        description = "Пояснение профиля " & CStr(profile) & ": прямое НДС."
        TestSetProfileValue source, profile, "Profile.DisplayName", caption
        TestSetProfileValue source, profile, "Profile.Description", description
        TestSetLoad loads, profile, -50000#, 1000000#, 2000000#, "Auto"
        data = loads.Value2: data(2, 5) = caption: loads.Value2 = data
        Set catalog = New CCalculationProfileCatalog: catalog.LoadFromWorkbook ThisWorkbook
        Set batch = TestExecuteFixture(section, provider, settings, units)
        Set state = batch.ResultAt(1).StateRepository.FindState(sstStrengthState)
        prefix = "profileConfig.PR" & CStr(profile)
        TestCheck stats, prefix & ".Profile.DisplayName.reader", catalog.ProfileById(caption).ProfileId = "PR" & CStr(profile)
        TestCheck stats, prefix & ".Profile.DisplayName.consumer", batch.ResultAt(1).ProfileId = "PR" & CStr(profile) And Not state Is Nothing
        TestCheck stats, prefix & ".Profile.Description.metadata", catalog.ProfileById(caption).Description = description
        TestCheckState stats, prefix & ".metadataState", state, section, provider, TestExpectedSpec("Strength", "", "")
        TestCheckComments stats, prefix, batch, units
        stats.Cases = stats.Cases + 2
        For Each value In Array(vbNullString, "  Пояснение профиля  ", "Произвольное пояснение: НДС и трещины")
            TestSetProfileValue source, profile, "Profile.Description", value
            Set catalog = New CCalculationProfileCatalog: catalog.LoadFromWorkbook ThisWorkbook
            TestCheck stats, prefix & ".Profile.Description.optional", catalog.ProfileById(caption).Description = Trim$(CStr(value))
            TestCheck stats, prefix & ".Profile.Description.nameStable", catalog.ProfileById(caption).ProfileId = "PR" & CStr(profile)
            stats.Cases = stats.Cases + 1
        Next value
        Set cell = TestProfileCell(source, profile, "Profile.Description"): cell.Value2 = CVErr(xlErrValue)
        code = TestCatalogError(reason)
        TestCheck stats, prefix & ".Profile.Description.error.rejected", code <> 0
        TestCheck stats, prefix & ".Profile.Description.error.controlled", code <> 13 And code <> 6
        TestCheck stats, prefix & ".Profile.Description.error.key", InStr(1, reason, "Profile.Description", vbBinaryCompare) > 0
        TestCheck stats, prefix & ".Profile.Description.error.address", InStr(1, reason, cell.Worksheet.Name & "!" & cell.Address(False, False), vbBinaryCompare) > 0
        TestSetProfileValue source, profile, "Profile.Description", description
        code = TestCatalogError(reason): TestCheck stats, prefix & ".Profile.Description.recovery", code = 0
        stats.Cases = stats.Cases + 1
    Next profile
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Выбирает пять видов НДС и обе отображаемые величины при обычном чтении Results.
' Сверяет числа по полю деформаций и материалу, точность 0–10 — по тексту легенды.
' Схема не должна пересчитывать НДС или округлять сохранённые значения.
Private Sub TestVisualization(ByRef stats As TProfileStats, ByVal source As Object, ByVal system As Object, _
        ByVal loads As Object, ByVal section As CSectionModel, ByVal provider As CMaterialModelProvider, _
        ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem)
    Dim profile As Long, stateName As Variant, quantity As Variant, precision As Long, prefix As String
    Dim batch As CBatchSectionCalculator, writer As CNDMResultsWriter, reader As CSectionPlotDataReader
    Dim state As CSectionStateResult, solves As Long, before As Variant, firstValue As Double
    TestSetTableValue system, "Plot.LegendMode", "Common", 2
    TestSetTableValue system, "Plot.LegendEnabled", "Yes", 2
    TestSetTableValue system, "Plot.ResultLabelsEnabled", "Yes", 2
    TestSetTableValue system, "Plot.ResultLabelSpacing", "50", 2
    For profile = 1 To 4
        TestConfigureProfiles source
        TestEnableChecks source, profile, "Yes", "Yes", "Yes", "No"
        TestSetLoad loads, profile, 200000#, 0#, 0#, "LambdaN"
        Set batch = TestExecuteFixture(section, provider, settings, units)
        Set writer = New CNDMResultsWriter: writer.WriteResults ThisWorkbook, section, PrepareSectionSnapshot(section, provider), batch, units
        before = ThisWorkbook.Names.Item("rngNDMElementResults").RefersToRange.CurrentRegion.Value2
        solves = SectionEquilibriumSolveCount()
        TestSetTableValue system, "Plot.LoadCase", "PROFILE_CONFIG", 2
        settings.LoadFromWorkbook ThisWorkbook
        For Each stateName In Array("StrengthState", "CapacityState", "PreCrackState", "PostCrackState", "CrackedState")
            TestSetProfileValue source, profile, "Visualization.State", stateName
            Set state = batch.ResultAt(1).StateRepository.FindState(SectionStateTypeFromText(CStr(stateName)))
            For Each quantity In Array("Stress", "Strain")
                TestSetProfileValue source, profile, "Visualization.Quantity", quantity
                prefix = "profileConfig.PR" & CStr(profile) & ".Visualization." & CStr(stateName) & "." & CStr(quantity)
                Set reader = New CSectionPlotDataReader: reader.LoadFromWorkbook ThisWorkbook, settings
                TestCheck stats, prefix & ".state", reader.StateType = CStr(stateName)
                TestCheck stats, prefix & ".quantity", reader.VisualizationQuantity = CStr(quantity)
                TestCheck stats, prefix & ".elements", reader.Count = section.ConcreteCount + section.RebarCount
                TestCheck stats, prefix & ".exists", Not state Is Nothing
                If Not state Is Nothing Then
                    TestCheckClose stats, prefix & ".plane", reader.Epsilon0, state.Epsilon0, 0.000000000001
                    TestCheckVisualizationValues stats, prefix, reader, state, section, provider, units
                End If
                firstValue = reader.ResultValue(1)
                For precision = 0 To 10
                    TestSetProfileValue source, profile, "Visualization." & CStr(quantity) & "Precision", precision
                    Set reader = New CSectionPlotDataReader: reader.LoadFromWorkbook ThisWorkbook, settings
                    TestCheck stats, prefix & ".precision" & CStr(precision), reader.ResultPrecision = precision
                    TestCheckClose stats, prefix & ".valueUnrounded" & CStr(precision), reader.ResultValue(1), firstValue, 0#
                    TestCheckVisualizationLegend stats, prefix & ".precision" & CStr(precision), reader, settings
                Next precision
                stats.Cases = stats.Cases + 1
            Next quantity
        Next stateName
        TestCheck stats, "profileConfig.PR" & CStr(profile) & ".Visualization.noSolve", SectionEquilibriumSolveCount() = solves
        TestCheck stats, "profileConfig.PR" & CStr(profile) & ".Visualization.snapshotUnchanged", _
            TestTablesEqual(before, ThisWorkbook.Names.Item("rngNDMElementResults").RefersToRange.CurrentRegion.Value2)
    Next profile
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Ошибка/пустота обязательного значения должна назвать профиль, key и ячейку.
' Дробная точность не округляется; переполнение не выдается за технический сбой VBA.
Private Sub TestInvalidProfileCells(ByRef stats As TProfileStats, ByVal source As Object)
    Dim profile As Long, key As Variant, value As Variant, code As Long, reason As String, prefix As String
    Dim cell As Object, caseIndex As Long
    For profile = 1 To 4
        For Each key In Array("Visualization.StressPrecision", "Visualization.StrainPrecision")
            caseIndex = 0
            For Each value In Array("", "-1", "11", "0.25", "9.9", "2147483648", "1e100", "abc", CVErr(xlErrDiv0))
                TestConfigureProfiles source
                Set cell = TestProfileCell(source, profile, CStr(key)): cell.Value2 = value
                code = TestCatalogError(reason)
                prefix = "profileConfig.PR" & CStr(profile) & "." & CStr(key) & ".invalid" & CStr(caseIndex)
                TestCheck stats, prefix & ".rejected", code <> 0
                TestCheck stats, prefix & ".controlled", code <> 6 And code <> 13
                TestCheck stats, prefix & ".key", InStr(1, reason, CStr(key), vbBinaryCompare) > 0
                TestCheck stats, prefix & ".address", InStr(1, reason, cell.Worksheet.Name & "!" & cell.Address(False, False), vbBinaryCompare) > 0
                TestCheck stats, prefix & ".profile", InStr(1, reason, "PR" & CStr(profile), vbBinaryCompare) > 0
                caseIndex = caseIndex + 1: stats.Cases = stats.Cases + 1
            Next value
        Next key
        For Each key In Array("MaterialModel.Strength.ValueSet", "MaterialModel.Strength.ConcreteDiagram", _
                "MaterialModel.Strength.ConcreteTension", "MaterialModel.Strength.SteelDiagram", _
                "MaterialModel.CrackInitiation.ValueSet", "MaterialModel.CrackInitiation.ConcreteDiagram", _
                "MaterialModel.CrackInitiation.ConcreteTension", "MaterialModel.CrackInitiation.SteelDiagram", _
                "MaterialModel.CrackedState.ValueSet", "MaterialModel.CrackedState.ConcreteDiagram", _
                "MaterialModel.CrackedState.ConcreteTension", "MaterialModel.CrackedState.SteelDiagram", _
                "MaterialModel.Stability.ValueSet", "Visualization.State", "Visualization.Quantity")
            TestConfigureProfiles source
            Set cell = TestProfileCell(source, profile, CStr(key)): cell.Value2 = "NOT_A_VALID_CHOICE"
            code = TestCatalogError(reason)
            prefix = "profileConfig.PR" & CStr(profile) & "." & CStr(key) & ".invalidChoice"
            TestCheck stats, prefix & ".rejected", code <> 0
            TestCheck stats, prefix & ".key", InStr(1, reason, CStr(key), vbBinaryCompare) > 0
            TestCheck stats, prefix & ".address", InStr(1, reason, cell.Worksheet.Name & "!" & cell.Address(False, False), vbBinaryCompare) > 0
            stats.Cases = stats.Cases + 1
        Next key
    Next profile
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает фактическое НДС с независимо выбранной моделью материалов
' и проверяет равновесие. Из вклада арматуры вычитает замещённый бетон.
Private Sub TestCheckState(ByRef stats As TProfileStats, ByVal prefix As String, ByVal state As CSectionStateResult, _
        ByVal section As CSectionModel, ByVal provider As CMaterialModelProvider, ByVal expected As CMaterialModelSpec)
    TestCheck stats, prefix & ".exists", Not state Is Nothing
    If state Is Nothing Then Exit Sub
    TestCheck stats, prefix & ".spec", state.MaterialSpec.SpecKey = expected.SpecKey
    TestCheck stats, prefix & ".converged", state.Converged
    If Not state.Converged Then Exit Sub
    Dim concrete As CMaterialDiagram, steel As CMaterialDiagram, i As Long
    Dim strain As Double, force As Double, n As Double, mx As Double, my As Double
    If provider.DiagramExtensionEnabled And (state.StateType = sstStrengthState Or state.StateType = sstCrackedState) Then
        Set concrete = provider.ConcreteMaterialForEquilibriumFromSpec(expected)
        Set steel = provider.SteelMaterialForEquilibriumFromSpec(expected)
    Else
        Set concrete = provider.ConcreteMaterialFromSpec(expected)
        Set steel = provider.SteelMaterialFromSpec(expected)
    End If
    For i = 1 To section.ConcreteCount
        strain = state.Epsilon0 + state.KappaX * section.ConcreteY(i) + state.KappaY * section.ConcreteX(i)
        force = concrete.GetStress(strain) * section.ConcreteArea(i)
        n = n + force: mx = mx + force * section.ConcreteY(i): my = my + force * section.ConcreteX(i)
    Next i
    For i = 1 To section.RebarCount
        strain = state.Epsilon0 + state.KappaX * section.RebarY(i) + state.KappaY * section.RebarX(i)
        force = (steel.GetStress(strain) - concrete.GetStress(strain)) * section.RebarArea(i)
        n = n + force: mx = mx + force * section.RebarY(i): my = my + force * section.RebarX(i)
    Next i
    TestCheckClose stats, prefix & ".expectedMaterialN", n, state.TargetN, 0.01
    TestCheckClose stats, prefix & ".expectedMaterialMx", mx, state.TargetMx, 1#
    TestCheckClose stats, prefix & ".expectedMaterialMy", my, state.TargetMy, 1#
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет пояснения всех этапов и адресный контекст в таблицах. Форматирование
' может добавить сочетание и ссылки, но исходная причина и статус сохраняются.
Private Sub TestCheckComments(ByRef stats As TProfileStats, ByVal prefix As String, _
        ByVal batch As CBatchSectionCalculator, ByVal units As CUnitSystem)
    Dim writer As CBatchResultWriter, result As CCombinationResult, meta As CResultMeta
    Dim names As Variant, metas As Variant, i As Long, policy As CResultStatusPolicy, solves As Long
    Set result = batch.ResultAt(1): Set writer = New CBatchResultWriter
    solves = batch.SolverCallCount: writer.WriteSummary ThisWorkbook, batch, units
    TestCheck stats, prefix & ".writerNoSolve", batch.SolverCallCount = solves
    TestCheck stats, prefix & ".batchComment", CStr(ThisWorkbook.Names.Item("rngBatchSummary").RefersToRange.Offset(12, 2).Value2) = CalculationResultMessage(batch, 1, result.OverallMeta)
    names = Array("rngStrengthSummaryAnchor", "rngCrackSummaryAnchor", "rngStabilitySummaryAnchor")
    metas = Array(result.StrengthMeta, result.CrackSummaryMeta, result.StabilityMeta)
    Set policy = New CResultStatusPolicy
    For i = 0 To UBound(names)
        Set meta = metas(i)
        Dim displayed As String
        displayed = CStr(ThisWorkbook.Names.Item(CStr(names(i))).RefersToRange.Offset(0, 1).Value2)
        TestCheck stats, prefix & ".comment" & CStr(i), displayed = CalculationResultMessage(batch, 1, meta)
        If Len(meta.ResultComment) > 0 Then
            TestCheck stats, prefix & ".causePreserved" & CStr(i), InStr(1, displayed, meta.ResultComment, vbBinaryCompare) > 0
            TestCheck stats, prefix & ".combinationContext" & CStr(i), InStr(1, displayed, batch.CombinationID(1), vbBinaryCompare) > 0
            If Len(batch.SourceReference(1)) > 0 Then TestCheck stats, prefix & ".inputRow" & CStr(i), InStr(1, displayed, batch.SourceReference(1), vbBinaryCompare) > 0
        End If
        If policy.IsFailedOverall(policy.ExternalStatus(meta)) Then TestCheck stats, prefix & ".failureReason" & CStr(i), Len(Trim$(meta.ResultComment)) > 0
    Next i
    TestLogLine stats, "PROFILE_CONFIG_RESULT: " & prefix & "; status=" & result.Status & "; comment=" & result.OverallMeta.ResultComment
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Создает небольшую физическую модель; параметры материалов заданы явно,
' чтобы результат не зависел от оставленных соседней suite значений Config.
Private Sub TestBuildFixture(ByRef section As CSectionModel, ByRef provider As CMaterialModelProvider, ByVal extensionEnabled As Boolean)
    Dim geometry As CGeometryRegion, mesh As CFiberMeshBuilder, rebars As CRebarLayout
    Dim concrete As CConcreteMaterialParameters, steel As CSteelMaterialParameters
    Set geometry = TestFixtureBox(300#, 200#)
    Set mesh = New CFiberMeshBuilder: mesh.BuildMesh geometry, 30#, 20#, 1
    Set rebars = New CRebarLayout
    rebars.AddBar "A", -90#, -60#, 20#, 0#, "A400", "", geometry
    rebars.AddBar "B", 90#, -60#, 20#, 0#, "A400", "", geometry
    rebars.AddBar "C", -90#, 60#, 20#, 0#, "A400", "", geometry
    rebars.AddBar "D", 90#, 60#, 20#, 0#, "A400", "", geometry
    Set section = TestBuildGeneratedSectionModel(mesh, rebars, "Audit03ProfileConfig")
    Set concrete = New CConcreteMaterialParameters: concrete.Initialize 15.5, 1.1, 22#, 1.8, 32500#, 32500#, rbMc2:=14.6
    Set steel = New CSteelMaterialParameters: steel.Initialize 350#, 350#, 390#, 390#, 200000#, 200000#
    Set provider = New CMaterialModelProvider
    provider.TestInitializeFromParameters concrete, steel, diagramExtensionEnabled:=extensionEnabled
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Читает настоящую таблицу сочетаний и рассчитывает её в новом объекте.
' Диаграммы профилей выбирает обычный код без подставного решателя.
Private Function TestExecuteFixture(ByVal section As CSectionModel, ByVal provider As CMaterialModelProvider, _
        ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem, _
        Optional ByVal readDurationLoads As Boolean = False, _
        Optional ByVal referenceX As Double = 0#, Optional ByVal referenceY As Double = 0#) As CBatchSectionCalculator
    Dim catalog As CCalculationProfileCatalog, batch As CBatchSectionCalculator, reader As CLoadCombinationReader
    Set catalog = New CCalculationProfileCatalog: catalog.LoadFromWorkbook ThisWorkbook
    Set batch = New CBatchSectionCalculator: batch.Initialize section, provider
    Set batch.ProfileCatalog = catalog: batch.ApplySettings settings, units
    batch.SetSP35Table721 ThisWorkbook.Names.Item("rngSP35Table721").RefersToRange.Value2
    Set reader = New CLoadCombinationReader: reader.LoadFromWorkbook ThisWorkbook, batch, units
    If readDurationLoads Then LoadStabilityDurationLoadsFromWorkbook ThisWorkbook, batch, units
    If referenceX <> 0# Or referenceY <> 0# Then batch.ApplyLoadReference referenceX, referenceY, 0#, 0#
    batch.Execute
    Set TestExecuteFixture = batch
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Задает независимую базу трех моделей и полей визуализации для всех PR.
' Отключенные флаги включаются конкретным тестом, не глобальным default.
Private Sub TestConfigureProfiles(ByVal source As Object)
    Dim profile As Long, role As Variant, spec As CMaterialModelSpec
    For profile = 1 To 4
        TestSetProfileValue source, profile, "Profile.DisplayName", "Профиль " & CStr(profile)
        TestSetProfileValue source, profile, "Profile.Description", "Проверка параметров профиля."
        TestEnableChecks source, profile, "No", "No", "No", "No"
        TestSetProfileValue source, profile, "MaterialModel.Stability.ValueSet", "ULS(I)"
        For Each role In Array("Strength", "CrackInitiation", "CrackedState")
            Set spec = TestExpectedSpec(CStr(role), "", "")
            TestSetProfileValue source, profile, "MaterialModel." & CStr(role) & ".ValueSet", spec.ValueSetText
            TestSetProfileValue source, profile, "MaterialModel." & CStr(role) & ".ConcreteDiagram", spec.ConcreteDiagramText
            TestSetProfileValue source, profile, "MaterialModel." & CStr(role) & ".ConcreteTension", spec.ConcreteTensionText
            TestSetProfileValue source, profile, "MaterialModel." & CStr(role) & ".SteelDiagram", spec.SteelDiagramText
        Next role
        TestSetProfileValue source, profile, "Visualization.State", "StrengthState"
        TestSetProfileValue source, profile, "Visualization.Quantity", "Stress"
        TestSetProfileValue source, profile, "Visualization.StressPrecision", "1"
        TestSetProfileValue source, profile, "Visualization.StrainPrecision", "6"
    Next profile
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' ======================= ДЛЯ ТЕСТОВ: DURATION LOADS =======================

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет каждую из 90 ячеек дополнительных нагрузок на реальном расчете.
' Отдельно проверяет ошибочный ввод, изоляцию LC, единицы/знаки, перенос точки
' приложения и адреса после перемещения имени. Формулы/Names восстанавливаются.
Public Function RunAudit03DurationConfigTests(Optional ByRef passed As Long = 0, _
        Optional ByRef failed As Long = 0, Optional ByVal matchingOnly As Boolean = False, _
        Optional ByVal localeOnly As Boolean = False) As String
    Dim stats As TProfileStats, names As Variant, saved(0 To 5) As Variant, i As Long
    Dim profiles As Object, system As Object, loads As Object, duration As Object
    Dim settings As CSystemSettingsReader, units As CUnitSystem, section As CSectionModel
    Dim provider As CMaterialModelProvider, durationRef As String, sheet As Object
    Dim oldAlerts As Boolean, restoreNumber As Long, restoreReason As String
    On Error GoTo FailedRun
    names = Array("rngCalculationProfiles", "rngSystemSettings", "rngUnitSettings", _
        "rngSignConventionSettings", "rngLoadCombinations", "rngStabilityDurationLoads")
    For i = 0 To UBound(names): saved(i) = ThisWorkbook.Names.Item(CStr(names(i))).RefersToRange.Formula: Next i
    durationRef = ThisWorkbook.Names.Item("rngStabilityDurationLoads").RefersTo
    oldAlerts = Application.DisplayAlerts
    Set profiles = ThisWorkbook.Names.Item("rngCalculationProfiles").RefersToRange
    Set system = ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange
    Set loads = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange
    Set duration = ThisWorkbook.Names.Item("rngStabilityDurationLoads").RefersToRange
    TestConfigureProfiles profiles: TestEnableChecks profiles, 1, "No", "No", "No", "Yes"
    TestSetTableValue system, "Stability.Code", "SP63", 2
    TestSetTableValue system, "Stability.ElementLength", "1000", 2
    TestSetTableValue system, "Stability.Mu1", "1", 2: TestSetTableValue system, "Stability.Mu2", "1", 2
    TestSetTableValue system, "Stability.SystemType", "Determinate", 2
    TestSetTableValue system, "Stability.ZeroMomentEccentricitySign1", "1", 2
    TestSetTableValue system, "Stability.ZeroMomentEccentricitySign2", "1", 2
    TestSetTableValue system, "Stability.AccidentalEccentricityMode", "User", 2
    TestSetTableValue system, "Stability.AccidentalEccentricityPlanes", "BothPlanes", 2
    TestSetTableValue system, "Stability.AccidentalEccentricityUser1", "5", 2
    TestSetTableValue system, "Stability.AccidentalEccentricityUser2", "5", 2
    TestSetTableValue system, "Stability.PhiLMode", "Auto", 2
    TestSetTableValue system, "Calculation.ZeroMomentPerDepth", "0", 2
    TestSetTableValue system, "Load.ReferenceOffsetX", "0", 2: TestSetTableValue system, "Load.ReferenceOffsetY", "0", 2
    TestSetTableValue system, "Solver.ToleranceN", "0.01", 2
    TestSetTableValue system, "Solver.ToleranceMx", "1", 2: TestSetTableValue system, "Solver.ToleranceMy", "1", 2
    TestConfigureDurationUnits "N", "N*mm", "Tension", "+Y tension", "+X tension"
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    TestBuildDurationFixture section, provider, settings.GetBoolean("General.DiagramExtension", False)
    If localeOnly Then
        TestDurationInvalidRows stats, profiles, loads, duration, section, provider, settings, units, 1
        TestDurationLocaleValid stats, loads, duration, section, provider, settings, units
        TestDurationOverflow stats, loads, duration, section, provider, settings, units
        GoTo Restore
    End If
    TestDurationMatching stats, profiles, system, loads, duration, section, provider, settings, units, sheet
    If matchingOnly Then GoTo Restore
    TestDurationCells stats, profiles, loads, duration, section, provider, settings, units
    TestDurationUnits stats, loads, duration, section, provider, settings, units
    TestConfigureDurationUnits "N", "N*mm", "Tension", "+Y tension", "+X tension"
    settings.LoadFromWorkbook ThisWorkbook: units.LoadFromSettings settings
    TestDurationInvalidRows stats, profiles, loads, duration, section, provider, settings, units
    TestDurationOverflow stats, loads, duration, section, provider, settings, units
    TestDurationInactive stats, profiles, loads, duration, section, provider, settings, units
    TestDurationTableLinks stats, loads, duration, section, provider, settings, units
    TestDurationRangeContract stats, loads, duration, durationRef, section, provider, settings, units, sheet
    GoTo Restore
FailedRun:
    stats.Failed = stats.Failed + 1
    TestLogLine stats, "FAIL: durationConfig.runtime; " & CStr(Err.Number) & "; " & Err.Description
Restore:
    On Error Resume Next
    ThisWorkbook.Names.Item("rngStabilityDurationLoads").RefersTo = durationRef
    For i = 0 To UBound(names)
        Err.Clear: ThisWorkbook.Names.Item(CStr(names(i))).RefersToRange.Formula = saved(i)
        If Err.Number <> 0 Then restoreNumber = Err.Number: restoreReason = Err.Description
    Next i
    If Not sheet Is Nothing Then Application.DisplayAlerts = False: sheet.Delete
    Application.DisplayAlerts = oldAlerts
    On Error GoTo 0
    If restoreNumber <> 0 Then stats.Failed = stats.Failed + 1: TestLogLine stats, "FAIL: durationConfig.restore; " & CStr(restoreNumber) & "; " & restoreReason
    TestLogLine stats, "TOTAL_AUDIT03_DURATION_CONFIG: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed) & "; cases=" & CStr(stats.Cases)
    passed = stats.Passed: failed = stats.Failed: RunAudit03DurationConfigTests = stats.Report
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Для симметричного высокого прямоугольника главные оси совпадают с X/Y.
' Поэтому знаки и компоненты длительных моментов проверяются без production
' преобразователя осей; поставщик материалов остается общим с другими тестами.
Private Sub TestBuildDurationFixture(ByRef section As CSectionModel, ByRef provider As CMaterialModelProvider, ByVal extensionEnabled As Boolean)
    Dim geometry As CGeometryRegion, mesh As CFiberMeshBuilder, rebars As CRebarLayout
    TestBuildFixture section, provider, extensionEnabled
    Set geometry = TestFixtureBox(200#, 300#)
    Set mesh = New CFiberMeshBuilder: mesh.BuildMesh geometry, 20#, 30#, 1
    Set rebars = New CRebarLayout
    rebars.AddBar "A", -60#, -90#, 20#, 0#, "A400", "", geometry
    rebars.AddBar "B", 60#, -90#, 20#, 0#, "A400", "", geometry
    rebars.AddBar "C", -60#, 90#, 20#, 0#, "A400", "", geometry
    rebars.AddBar "D", 60#, 90#, 20#, 0#, "A400", "", geometry
    Set section = TestBuildGeneratedSectionModel(mesh, rebars, "Audit03DurationConfig")
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет каждую исходную строку и составляющую отдельно. Другие строки имеют
' иные номера, чтобы выявить ошибочное сопоставление по месту строки.
Private Sub TestDurationCells(ByRef stats As TProfileStats, ByVal profiles As Object, ByVal loads As Object, _
        ByVal duration As Object, ByVal section As CSectionModel, ByVal provider As CMaterialModelProvider, _
        ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem)
    Dim slot As Long, column As Long, testValue As Variant, variants As Variant, data As Variant
    Dim values As Variant, batch As CBatchSectionCalculator, code As Long, reason As String, prefix As String
    For slot = 1 To duration.Rows.Count - 1
        For column = 2 To 4
            Select Case column
                Case 2: variants = Array(-10000#, -70000#, 5000#, 0#, vbNullString)
                Case 3: variants = Array(-800000#, 1200000#, 0#, vbNullString)
                Case 4: variants = Array(-900000#, 2500000#, 0#, vbNullString)
            End Select
            For Each testValue In variants
                TestPrepareDurationLoads loads, duration, slot
                data = duration.Value2: data(slot + 1, column) = testValue: duration.Value2 = data
                values = Array(-25000#, 400000#, 600000#)
                If Len(CStr(testValue)) = 0 Then values(column - 2) = 0# Else values(column - 2) = CDbl(testValue)
                prefix = "durationConfig.Slot" & CStr(slot) & "." & TestDurationColumnKey(column) & ".value=" & CStr(testValue)
                TestLogLine stats, "RUN: " & prefix
                code = TestExecuteDuration(section, provider, settings, units, batch, reason)
                TestCheck stats, prefix & ".reader", code = 0
                If code = 0 Then
                    TestCheckDurationResult stats, prefix, batch, units, CDbl(values(0)), CDbl(values(1)), CDbl(values(2))
                    TestCheckComments stats, prefix, batch, units
                Else
                    TestLogLine stats, "DURATION_ERROR: " & prefix & "; " & CStr(code) & "; " & reason
                End If
                stats.Cases = stats.Cases + 1
            Next testValue
        Next column
    Next slot
    TestPrepareDurationLoads loads, duration, 1 ' Перенос длительной нагрузки должен использовать ту же точку, что полный LC.
    code = TestExecuteDuration(section, provider, settings, units, batch, reason, 40#, -25#)
    TestCheck stats, "durationConfig.offset.reader", code = 0
    If code = 0 Then TestCheckDurationResult stats, "durationConfig.offset", batch, units, -25000#, 1025000#, -400000#, 2250000#, 0#
    stats.Cases = stats.Cases + 1
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет эквивалентные физические нагрузки в N/kN/tf и N*mm/kN*m/tf*m,
' оба знака N и независимые знаки Mx/My. Ожидаемые масштабы заданы числами,
' а не вычисляются тем же InputForceToInternal, который проверяется.
Private Sub TestDurationUnits(ByRef stats As TProfileStats, ByVal loads As Object, ByVal duration As Object, _
        ByVal section As CSectionModel, ByVal provider As CMaterialModelProvider, _
        ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem)
    Dim mode As Long, signN As Long, signMx As Long, signMy As Long, data As Variant
    Dim forceUnit As String, momentUnit As String, forceScale As Double, momentScale As Double
    Dim nSign As String, mxSign As String, mySign As String, batch As CBatchSectionCalculator
    Dim code As Long, reason As String, prefix As String
    For mode = 1 To 3
        Select Case mode
            Case 1: forceUnit = "N": momentUnit = "N*mm": forceScale = 1#: momentScale = 1#
            Case 2: forceUnit = "kN": momentUnit = "kN*m": forceScale = 1000#: momentScale = 1000000#
            Case 3: forceUnit = "tf": momentUnit = "tf*m": forceScale = 9806.65: momentScale = 9806650#
        End Select
        For signN = -1 To 1 Step 2
            For signMx = -1 To 1 Step 2
                For signMy = -1 To 1 Step 2
                    If signN = 1 Then nSign = "Tension" Else nSign = "Compression"
                    If signMx = 1 Then mxSign = "+Y tension" Else mxSign = "-Y tension"
                    If signMy = 1 Then mySign = "+X tension" Else mySign = "-X tension"
                    TestConfigureDurationUnits forceUnit, momentUnit, nSign, mxSign, mySign
                    settings.LoadFromWorkbook ThisWorkbook: units.LoadFromSettings settings
                    TestPrepareDurationLoads loads, duration, 1
                    data = loads.Value2
                    data(2, 2) = -50000# / forceScale * signN
                    data(2, 3) = 1000000# / momentScale * signMx: data(2, 4) = 2000000# / momentScale * signMy
                    loads.Value2 = data
                    data = duration.Value2
                    data(2, 2) = -25000# / forceScale * signN
                    data(2, 3) = 400000# / momentScale * signMx: data(2, 4) = 600000# / momentScale * signMy
                    duration.Value2 = data
                    prefix = "durationConfig.units." & CStr(mode) & "." & CStr(signN) & "." & CStr(signMx) & "." & CStr(signMy)
                    TestLogLine stats, "RUN: " & prefix
                    code = TestExecuteDuration(section, provider, settings, units, batch, reason)
                    TestCheck stats, prefix & ".reader", code = 0
                    If code = 0 Then TestCheckDurationResult stats, prefix, batch, units, -25000#, 400000#, 600000#: TestCheckComments stats, prefix, batch, units
                    stats.Cases = stats.Cases + 1
                Next signMy
            Next signMx
        Next signN
    Next mode
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Ошибка каждой активной ячейки относится только к устойчивости данного LC.
' DirectState и следующая корректная строка должны быть рассчитаны; невыполненная
' устойчивость не получает Calculated=True или NumFail.
Private Sub TestDurationInvalidRows(ByRef stats As TProfileStats, ByVal profiles As Object, ByVal loads As Object, _
        ByVal duration As Object, ByVal section As CSectionModel, ByVal provider As CMaterialModelProvider, _
        ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem, Optional ByVal slotLimit As Long = 0)
    Dim slot As Long, column As Long, testValue As Variant, data As Variant, batch As CBatchSectionCalculator
    Dim code As Long, reason As String, prefix As String, cell As Object, errorCase As Long
    TestEnableChecks profiles, 1, "Yes", "No", "No", "Yes"
    If slotLimit = 0 Then slotLimit = duration.Rows.Count - 1
    For slot = 1 To slotLimit
        For column = 2 To 4
            For errorCase = 1 To IIf(slot = 1, 4, 2)
                TestPrepareDurationLoads loads, duration, slot
                Select Case errorCase
                    Case 1: testValue = "BAD_DURATION"
                    Case 2: testValue = CVErr(2015)
                    Case 3: testValue = "1 2"
                    Case 4: testValue = "12 34"
                End Select
                data = duration.Value2: data(slot + 1, column) = testValue: duration.Value2 = data
                TestAddDurationRecoveryLoad loads, duration, slot
                Set cell = duration.Cells(slot + 1, column)
                prefix = "durationConfig.Slot" & CStr(slot) & "." & TestDurationColumnKey(column) & ".error" & CStr(errorCase)
                TestLogLine stats, "RUN: " & prefix
                code = TestExecuteDuration(section, provider, settings, units, batch, reason)
                TestCheck stats, prefix & ".readerNoAbort", code = 0
                If code = 0 Then
                    TestCheckDurationInputError stats, prefix, batch, units, cell.Worksheet.Name & "!" & cell.Address(False, False)
                    TestCheck stats, prefix & ".directContinues", batch.ResultAt(1).DirectStateMeta.InternalStatus = rsSuccess
                    TestCheck stats, prefix & ".recovery", batch.Count = 2 And batch.ResultAt(2).StabilityMeta.InternalStatus = rsSuccess
                    TestCheck stats, prefix & ".recoveryDirect", batch.ResultAt(2).DirectStateMeta.InternalStatus = rsSuccess
                    TestCheck stats, prefix & ".recoveryBatchComment", _
                        CStr(ThisWorkbook.Names.Item("rngBatchSummary").RefersToRange.Offset(13, 2).Value2) = CalculationResultMessage(batch, 2, batch.ResultAt(2).OverallMeta)
                    TestCheck stats, prefix & ".recoveryOwnComment", _
                        CStr(ThisWorkbook.Names.Item("rngStabilitySummaryAnchor").RefersToRange.Offset(1, 1).Value2) = CalculationResultMessage(batch, 2, batch.ResultAt(2).StabilityMeta)
                    TestCheck stats, prefix & ".noReasonLeak", InStr(1, batch.ResultAt(2).StabilityMeta.ResultComment, _
                        cell.Worksheet.Name & "!" & cell.Address(False, False), vbBinaryCompare) = 0
                Else
                    TestLogLine stats, "DURATION_ERROR: " & prefix & "; " & CStr(code) & "; " & reason
                End If
                stats.Cases = stats.Cases + 1
            Next errorCase
        Next column
    Next slot
    TestEnableChecks profiles, 1, "No", "No", "No", "Yes"
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Число может помещаться в Double до перевода, но переполнить внутренние Н
' или Н*мм. Это адресная ошибка входа данной строки, а не NumFail solver-а.
Private Sub TestDurationOverflow(ByRef stats As TProfileStats, ByVal loads As Object, ByVal duration As Object, _
        ByVal section As CSectionModel, ByVal provider As CMaterialModelProvider, _
        ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem)
    Dim column As Long, data As Variant, batch As CBatchSectionCalculator, code As Long, reason As String, prefix As String, cell As Object
    TestConfigureDurationUnits "tf", "tf*m", "Tension", "+Y tension", "+X tension"
    settings.LoadFromWorkbook ThisWorkbook: units.LoadFromSettings settings
    For column = 2 To 4
        TestPrepareDurationLoads loads, duration, 1
        data = loads.Value2: data(2, 2) = -50000# / 9806.65
        data(2, 3) = 1000000# / 9806650#: data(2, 4) = 2000000# / 9806650#: loads.Value2 = data
        data = duration.Value2: data(2, 2) = -25000# / 9806.65
        data(2, 3) = 400000# / 9806650#: data(2, 4) = 600000# / 9806650#
        data(2, column) = 1E+305: duration.Value2 = data
        prefix = "durationConfig.overflow." & TestDurationColumnKey(column): TestLogLine stats, "RUN: " & prefix
        code = TestExecuteDuration(section, provider, settings, units, batch, reason)
        TestCheck stats, prefix & ".readerNoAbort", code = 0
        Set cell = duration.Cells(2, column)
        If code = 0 Then TestCheckDurationInputError stats, prefix, batch, units, cell.Worksheet.Name & "!" & cell.Address(False, False)
        stats.Cases = stats.Cases + 1
    Next column
    TestConfigureDurationUnits "N", "N*mm", "Tension", "+Y tension", "+X tension"
    settings.LoadFromWorkbook ThisWorkbook: units.LoadFromSettings settings
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет case-insensitive ID и существующий приоритет первой строки.
' Ошибка формулы в самом ID является структурной, ее адрес не выдумывается.
Private Sub TestDurationTableLinks(ByRef stats As TProfileStats, ByVal loads As Object, ByVal duration As Object, _
        ByVal section As CSectionModel, ByVal provider As CMaterialModelProvider, _
        ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem)
    Dim mode As Long, data As Variant, batch As CBatchSectionCalculator, code As Long, reason As String, prefix As String, cell As Object
    For mode = 1 To 4
        TestPrepareDurationLoads loads, duration, 1
        data = duration.Value2
        If mode = 1 Then data(2, 1) = "duration_1"
        If mode = 2 Then data(3, 1) = "DURATION_1": data(3, 2) = "BAD_DUPLICATE"
        If mode = 3 Then data(3, 1) = "DURATION_1": data(2, 2) = "BAD_FIRST"
        If mode = 4 Then data(3, 1) = CVErr(2015)
        duration.Value2 = data
        prefix = "durationConfig.links." & CStr(mode): TestLogLine stats, "RUN: " & prefix
        code = TestExecuteDuration(section, provider, settings, units, batch, reason)
        TestCheck stats, prefix & ".readerNoAbort", code = 0
        If code = 0 Then
            If mode <= 2 Then
                TestCheckDurationResult stats, prefix, batch, units, -25000#, 400000#, 600000#
                TestCheckComments stats, prefix, batch, units
            Else
                If mode = 3 Then Set cell = duration.Cells(2, 2) Else Set cell = duration.Cells(3, 1)
                TestCheckDurationInputError stats, prefix, batch, units, cell.Worksheet.Name & "!" & cell.Address(False, False)
            End If
        End If
        stats.Cases = stats.Cases + 1
    Next mode
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Неактивный профиль и неиспользуемый ID не должны зависеть от ошибочных
' чисел дополнительной строки. Проверяется также пустая строка без ID.
Private Sub TestDurationInactive(ByRef stats As TProfileStats, ByVal profiles As Object, ByVal loads As Object, _
        ByVal duration As Object, ByVal section As CSectionModel, ByVal provider As CMaterialModelProvider, _
        ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem)
    Dim mode As Long, data As Variant, batch As CBatchSectionCalculator, code As Long, reason As String, prefix As String
    For mode = 1 To 3
        TestPrepareDurationLoads loads, duration, 1
        TestEnableChecks profiles, 1, "Yes", "No", "No", "Yes"
        data = duration.Value2
        If mode = 1 Then TestEnableChecks profiles, 1, "Yes", "No", "No", "No"
        If mode = 2 Then data(3, 1) = "UNUSED_ID"
        If mode = 3 Then data(3, 1) = vbNullString
        If mode = 1 Then data(2, 2) = CVErr(2015) Else data(3, 2) = CVErr(2015)
        duration.Value2 = data
        prefix = "durationConfig.inactive." & CStr(mode): TestLogLine stats, "RUN: " & prefix
        code = TestExecuteDuration(section, provider, settings, units, batch, reason)
        TestCheck stats, prefix & ".readerNoAbort", code = 0
        If code = 0 Then
            TestCheck stats, prefix & ".direct", batch.ResultAt(1).DirectStateMeta.InternalStatus = rsSuccess
            If mode = 1 Then TestCheck stats, prefix & ".notRequested", batch.ResultAt(1).StabilityMeta.InternalStatus = rsNotRequested Else TestCheck stats, prefix & ".stability", batch.ResultAt(1).StabilityMeta.InternalStatus = rsSuccess
            TestCheckComments stats, prefix, batch, units
        End If
        stats.Cases = stats.Cases + 1
    Next mode
    TestEnableChecks profiles, 1, "No", "No", "No", "Yes"
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Непрочитанный активный named range не эквивалентен пустой допустимой строке.
' При переносе на другой лист ошибка обязана указывать новое место ввода.
Private Sub TestDurationRangeContract(ByRef stats As TProfileStats, ByVal loads As Object, ByVal duration As Object, _
        ByVal originalRef As String, ByVal section As CSectionModel, ByVal provider As CMaterialModelProvider, _
        ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem, ByRef sheet As Object)
    Dim code As Long, reason As String, batch As CBatchSectionCalculator, position As Long, column As Long
    Dim target As Object, data As Variant, prefix As String, cell As Object
    TestPrepareDurationLoads loads, duration, 1
    ThisWorkbook.Names.Item("rngStabilityDurationLoads").RefersTo = "=#REF!"
    code = TestExecuteDuration(section, provider, settings, units, batch, reason)
    TestCheck stats, "durationConfig.missing.readerNoAbort", code = 0
    If code = 0 Then TestCheckDurationInputError stats, "durationConfig.missing", batch, units, "rngStabilityDurationLoads"
    stats.Cases = stats.Cases + 1
    ThisWorkbook.Names.Item("rngStabilityDurationLoads").RefersTo = originalRef
    ThisWorkbook.Names.Item("rngStabilityDurationLoads").RefersTo = "='" & duration.Worksheet.Name & "'!" & duration.Resize(duration.Rows.Count, 3).Address
    code = TestExecuteDuration(section, provider, settings, units, batch, reason)
    TestCheck stats, "durationConfig.columns.readerNoAbort", code = 0
    If code = 0 Then TestCheckDurationInputError stats, "durationConfig.columns", batch, units, "rngStabilityDurationLoads"
    stats.Cases = stats.Cases + 1
    ThisWorkbook.Names.Item("rngStabilityDurationLoads").RefersTo = originalRef
    If sheet Is Nothing Then Set sheet = ThisWorkbook.Worksheets.Add: sheet.Name = "__AuditDurationInputs"
    For position = 1 To 2
        If position = 1 Then Set target = sheet.Cells(10, 5) Else Set target = sheet.Cells(800, 60)
        Set target = target.Resize(duration.Rows.Count, duration.Columns.Count)
        For column = 2 To 4
            TestPrepareDurationLoads loads, duration, 1
            data = duration.Value2: data(2, column) = CVErr(2015): target.Value2 = data
            ThisWorkbook.Names.Item("rngStabilityDurationLoads").RefersTo = "='" & sheet.Name & "'!" & target.Address
            Set cell = target.Cells(2, column)
            prefix = "durationConfig.relocated." & CStr(position) & "." & TestDurationColumnKey(column): TestLogLine stats, "RUN: " & prefix
            code = TestExecuteDuration(section, provider, settings, units, batch, reason)
            TestCheck stats, prefix & ".readerNoAbort", code = 0
            If code = 0 Then TestCheckDurationInputError stats, prefix, batch, units, sheet.Name & "!" & cell.Address(False, False)
            ThisWorkbook.Names.Item("rngStabilityDurationLoads").RefersTo = originalRef
            stats.Cases = stats.Cases + 1
        Next column
    Next position
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Задает один основной LC и 30 независимых duration ID. Пустой ID сам по себе
' не вводит ошибочный LC, а сопоставление производится по ID, не по номеру строки.
Private Sub TestPrepareDurationLoads(ByVal loads As Object, ByVal duration As Object, ByVal activeSlot As Long)
    Dim data As Variant, row As Long
    TestSetLoad loads, 1, -50000#, 1000000#, 2000000#, "Auto"
    data = loads.Value2: data(2, 1) = "DURATION_" & CStr(activeSlot): loads.Value2 = data
    data = duration.Value2
    For row = 2 To UBound(data, 1)
        data(row, 1) = "DURATION_" & CStr(row - 1)
        data(row, 2) = -25000#: data(row, 3) = 400000#: data(row, 4) = 600000#
    Next row
    duration.Value2 = data
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Добавляет следующую корректную строку, не совпадающую с поврежденным ID.
' Для неё явно задана строка с нулевой длительной частью.
Private Sub TestAddDurationRecoveryLoad(ByVal loads As Object, ByVal duration As Object, ByVal invalidSlot As Long)
    Dim data As Variant
    data = loads.Value2
    data(3, 1) = "DURATION_RECOVERY": data(3, 2) = -50000#
    data(3, 3) = 1000000#: data(3, 4) = 2000000#: data(3, 5) = "PR1": data(3, 6) = "Auto"
    loads.Value2 = data
    data = duration.Value2
    Dim row As Long
    row = IIf(invalidSlot = 1, 3, 2)
    data(row, 1) = "DURATION_RECOVERY": data(row, 2) = 0#: data(row, 3) = 0#: data(row, 4) = 0#
    duration.Value2 = data
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Меняет только таблицы единиц/знаков; все физические величины теста остаются
' одинаковыми. OUTPUT установлен явно, чтобы предыдущая suite не влияла на writer.
Private Sub TestConfigureDurationUnits(ByVal force As String, ByVal moment As String, ByVal signN As String, ByVal signMx As String, ByVal signMy As String)
    Dim source As Object
    Set source = ThisWorkbook.Names.Item("rngUnitSettings").RefersToRange
    TestSetTableValue source, "Length", "mm", 2: TestSetTableValue source, "Force", force, 2
    TestSetTableValue source, "Moment", moment, 2: TestSetTableValue source, "Stress", "MPa", 2
    TestSetTableValue source, "Force", force, 4: TestSetTableValue source, "Moment", moment, 4
    Set source = ThisWorkbook.Names.Item("rngSignConventionSettings").RefersToRange
    TestSetTableValue source, "+N", signN, 2: TestSetTableValue source, "+Mx", signMx, 2: TestSetTableValue source, "+My", signMy, 2
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сохраняет первоначальную ошибку чтения для сравнения до исправления.
' После исправления ошибка строки должна храниться в StabilityMeta, а не прерывать вызов.
Private Function TestExecuteDuration(ByVal section As CSectionModel, ByVal provider As CMaterialModelProvider, _
        ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem, _
        ByRef batch As CBatchSectionCalculator, ByRef reason As String, _
        Optional ByVal referenceX As Double = 0#, Optional ByVal referenceY As Double = 0#) As Long
    Set batch = Nothing: reason = vbNullString
    On Error GoTo Failed
    Set batch = TestExecuteFixture(section, provider, settings, units, True, referenceX, referenceY)
    Exit Function
Failed:
    TestExecuteDuration = Err.Number: reason = Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет подготовленные внутренние усилия и их фактический вывод. PhiL
' обязан реагировать на длительную часть; дополнительного State solve нет.
Private Sub TestCheckDurationResult(ByRef stats As TProfileStats, ByVal prefix As String, ByVal batch As CBatchSectionCalculator, _
        ByVal units As CUnitSystem, ByVal n As Double, ByVal mx As Double, ByVal my As Double, _
        Optional ByVal totalMx As Double = 1000000#, Optional ByVal totalMy As Double = 2000000#)
    Dim result As CStabilityResult, writer As CBatchResultWriter, source As Object
    Set result = batch.ResultAt(1).StabilityResult
    TestCheck stats, prefix & ".calculated", result.Meta.Calculated
    TestCheck stats, prefix & ".status", result.Meta.InternalStatus = rsSuccess
    TestCheckClose stats, prefix & ".N", result.SustainedN, -n, 0.01
    TestCheckClose stats, prefix & ".Mx", result.SustainedMoment1, mx, 0.01
    TestCheckClose stats, prefix & ".My", result.SustainedMoment2, my, 0.01
    TestCheckClose stats, prefix & ".PhiL1", result.PhiL1, TestDurationExpectedPhi(mx - n * 90#, totalMx + 50000# * 90#), 0.00000001
    TestCheckClose stats, prefix & ".PhiL2", result.PhiL2, TestDurationExpectedPhi(my - n * 60#, totalMy + 50000# * 60#), 0.00000001
    TestCheck stats, prefix & ".noSolve", batch.SolverCallCount = 0
    Set writer = New CBatchResultWriter: writer.WriteSummary ThisWorkbook, batch, units
    Set source = ThisWorkbook.Names.Item("rngStabilitySummaryAnchor").RefersToRange
    TestCheckClose stats, prefix & ".outputN", CDbl(source.Offset(0, 18).Value2), units.InternalForceToOutput(n), 0.00000001
    TestCheckClose stats, prefix & ".outputM1", CDbl(source.Offset(0, 19).Value2), units.InternalMomentMagnitudeToOutput(mx), 0.00000001
    TestCheckClose stats, prefix & ".outputM2", CDbl(source.Offset(0, 20).Value2), units.InternalMomentMagnitudeToOutput(my), 0.00000001
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Независимый oracle существующей формулы СП 63 для ненулевого полного
' момента относительно выбранного стержня; здесь не проверяется новая норма.
Private Function TestDurationExpectedPhi(ByVal sustainedMoment As Double, ByVal totalMoment As Double) As Double
    TestDurationExpectedPhi = 1# + sustainedMoment / totalMoment
    If TestDurationExpectedPhi < 1# Then TestDurationExpectedPhi = 1#
    If TestDurationExpectedPhi > 2# Then TestDurationExpectedPhi = 2#
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Ошибка ввода длительной нагрузки не является несходимостью.
' Полная причина с адресом должна сохраняться в результате устойчивости
' и передаваться в таблицу без повторного составления текста.
Private Sub TestCheckDurationInputError(ByRef stats As TProfileStats, ByVal prefix As String, ByVal batch As CBatchSectionCalculator, _
        ByVal units As CUnitSystem, ByVal location As String)
    Dim meta As CResultMeta, policy As CResultStatusPolicy
    Set meta = batch.ResultAt(1).StabilityMeta: Set policy = New CResultStatusPolicy
    TestCheck stats, prefix & ".typed", meta.InternalStatus = rsInvalidInput And meta.ResultCode = rcInvalidInput
    TestCheck stats, prefix & ".display", policy.ExternalStatus(meta) = "InputErr"
    TestCheck stats, prefix & ".notCalculated", Not meta.Calculated
    TestCheck stats, prefix & ".address", InStr(1, meta.ResultComment, location, vbBinaryCompare) > 0
    TestCheck stats, prefix & ".repair", InStr(1, meta.ResultComment, "исправ", vbTextCompare) > 0 Or InStr(1, meta.ResultComment, "восстанов", vbTextCompare) > 0
    TestCheckComments stats, prefix, batch, units
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Короткий машинный ключ компоненты не зависит от выбранной подписи единиц.
Private Function TestDurationColumnKey(ByVal column As Long) As String
    Select Case column
        Case 2: TestDurationColumnKey = "N"
        Case 3: TestDurationColumnKey = "Mx"
        Case 4: TestDurationColumnKey = "My"
    End Select
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Выбирает запросы существующих calculators без смешения с material spec.
Private Sub TestEnableChecks(ByVal source As Object, ByVal profile As Long, ByVal direct As String, _
        ByVal capacity As String, ByVal crack As String, ByVal stability As String)
    TestSetProfileValue source, profile, "Calculation.Strength.DirectState", direct
    TestSetProfileValue source, profile, "Calculation.Strength.Capacity", capacity
    TestSetProfileValue source, profile, "Calculation.Crack.Width", crack
    TestSetProfileValue source, profile, "Calculation.Stability.Enabled", stability
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Формирует один физический LC; остальные source slots остаются пустыми.
Private Sub TestSetLoad(ByVal source As Object, ByVal profile As Long, ByVal n As Double, _
        ByVal mx As Double, ByVal my As Double, ByVal path As String)
    Dim data As Variant, row As Long, column As Long
    data = source.Value2
    For row = 2 To UBound(data, 1)
        For column = 1 To UBound(data, 2): data(row, column) = vbNullString: Next column
    Next row
    data(2, 1) = "PROFILE_CONFIG": data(2, 2) = n: data(2, 3) = mx: data(2, 4) = my
    data(2, 5) = "PR" & CStr(profile): data(2, 6) = path: data(2, 7) = "Проверка полей профиля."
    source.Value2 = data
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает независимые ожидаемые тексты модели, не читая фактический State.
Private Function TestExpectedSpec(ByVal role As String, ByVal field As String, ByVal value As String) As CMaterialModelSpec
    Dim valueSet As String, concrete As String, tension As String, steel As String, spec As CMaterialModelSpec
    valueSet = "SLS(II)": concrete = "TwoLine": tension = "Ignore": steel = "TwoLine"
    If role = "Strength" Then valueSet = "ULS(I)"
    If role = "CrackInitiation" Then concrete = "ThreeLine": tension = "UseDiagram"
    Select Case field
        Case "ValueSet": valueSet = value
        Case "ConcreteDiagram": concrete = value
        Case "ConcreteTension": tension = value
        Case "SteelDiagram": steel = value
    End Select
    Set spec = New CMaterialModelSpec: spec.Initialize valueSet, concrete, tension, steel
    Set TestExpectedSpec = spec
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Перечисляет допустимые активные варианты роли; ограничения tensile mode
' принадлежат утвержденной постановке Pre/Post, а не новой тестовой формуле.
Private Function TestMaterialVariants(ByVal role As String, ByVal field As String) As Variant
    Select Case field
        Case "ValueSet": TestMaterialVariants = Array("ULS(I)", "SLS(II)")
        Case "ConcreteDiagram", "SteelDiagram": TestMaterialVariants = Array("TwoLine", "ThreeLine")
        Case "ConcreteTension"
            Select Case role
                Case "Strength": TestMaterialVariants = Array("Ignore", "UseDiagram")
                Case "CrackInitiation": TestMaterialVariants = Array("UseDiagram")
                Case "CrackedState": TestMaterialVariants = Array("Ignore")
            End Select
    End Select
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Находит фактическую профильную ячейку по key; ошибка fixture не скрывается.
Private Function TestProfileCell(ByVal source As Object, ByVal profile As Long, ByVal key As String) As Object
    Dim row As Long
    For row = 1 To source.Rows.Count
        If CStr(source.Cells(row, 2).Value2) = key Then Set TestProfileCell = source.Cells(row, profile + 2): Exit Function
    Next row
    Err.Raise vbObjectError + 4501, "ProfileCell", "В тестовой таблице нет " & key
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Записывает Variant без превращения CVErr или пустоты в строковый default.
Private Sub TestSetProfileValue(ByVal source As Object, ByVal profile As Long, ByVal key As String, ByVal value As Variant)
    TestProfileCell(source, profile, key).Value2 = value
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Находит scalar-настройку по фактическому ключу ее существующей таблицы.
Private Sub TestSetTableValue(ByVal source As Object, ByVal key As String, ByVal value As String, ByVal column As Long)
    Dim row As Long
    For row = 1 To source.Rows.Count
        If CStr(source.Cells(row, 1).Value2) = key Then source.Cells(row, column).Value2 = value: Exit Sub
    Next row
    Err.Raise vbObjectError + 4501, "SetTableValue", "В тестовой таблице нет " & key
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сохраняет фактическую контролируемую ошибку каталога, не назначая статус по тексту.
Private Function TestCatalogError(ByRef reason As String) As Long
    Dim catalog As CCalculationProfileCatalog
    reason = vbNullString
    On Error GoTo Failed
    Set catalog = New CCalculationProfileCatalog: catalog.LoadFromWorkbook ThisWorkbook
    Exit Function
Failed:
    TestCatalogError = Err.Number: reason = Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает сериализованные численные таблицы без нормализации или допусков.
Private Function TestTablesEqual(ByRef before As Variant, ByRef after As Variant) As Boolean
    Dim row As Long, column As Long
    If UBound(before, 1) <> UBound(after, 1) Or UBound(before, 2) <> UBound(after, 2) Then Exit Function
    For row = 1 To UBound(before, 1)
        For column = 1 To UBound(before, 2)
            If CStr(before(row, column)) <> CStr(after(row, column)) Then Exit Function
        Next column
    Next row
    TestTablesEqual = True
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сохраняет стабильный assertion ID и отдельный счетчик ошибок.
Private Sub TestCheck(ByRef stats As TProfileStats, ByVal id As String, ByVal condition As Boolean)
    If condition Then stats.Passed = stats.Passed + 1: TestLogLine stats, "OK: " & id Else stats.Failed = stats.Failed + 1: TestLogLine stats, "FAIL: " & id
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает с прежним компонентным допуском; precision не округляет actual.
Private Sub TestCheckClose(ByRef stats As TProfileStats, ByVal id As String, ByVal actual As Double, ByVal expected As Double, ByVal tolerance As Double)
    TestCheck stats, id & "; actual=" & Trim$(Str$(actual)) & "; expected=" & Trim$(Str$(expected)), Abs(actual - expected) <= tolerance
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Записывает диагностический checkpoint для внешнего watchdog без принятия PASS.
Private Sub TestLogLine(ByRef stats As TProfileStats, ByVal value As String)
    stats.Report = stats.Report & value & vbCrLf
    If Left$(value, 5) = "RUN: " Then
        Dim stream As Object
        Set stream = CreateObject("Scripting.FileSystemObject").OpenTextFile(ThisWorkbook.Path & "\RC_NDM_ui_test_progress.txt", 8, True, -1)
        stream.WriteLine Format$(Now, "yyyy-mm-dd hh:nn:ss") & " " & value: stream.Close
    End If
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' ==================== ДЛЯ ТЕСТОВ: ОФОРМЛЕНИЕ И SNAPSHOT ====================

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет настройки аннотаций через действительные ячейки и Chart.Shapes.
' Численные Results создаются один раз; изменение оформления не запускает
' равновесие и не меняет snapshot. Ошибочные и выключенные поля проверяются
' раздельно, исходные таблицы и переадресованный Name восстанавливаются.
Public Function RunAudit03PresentationConfigTests(Optional ByRef passed As Long = 0, _
        Optional ByRef failed As Long = 0) As String
    Dim stats As TProfileStats, names As Variant, saved(0 To 5) As Variant, i As Long
    Dim annotations As Object, profiles As Object, system As Object, loads As Object
    Dim settings As CSystemSettingsReader, units As CUnitSystem, section As CSectionModel
    Dim provider As CMaterialModelProvider, batch As CBatchSectionCalculator, writer As CNDMResultsWriter
    Dim reader As CSectionPlotDataReader, before As Variant, solves As Long, oldRef As String
    Dim sheet As Object, oldAlerts As Boolean, restoreNumber As Long, restoreReason As String
    On Error GoTo FailedRun
    names = Array("rngPlotAnnotationSettings", "rngCalculationProfiles", "rngSystemSettings", _
        "rngUnitSettings", "rngSignConventionSettings", "rngLoadCombinations")
    For i = 0 To UBound(names): saved(i) = ThisWorkbook.Names.Item(CStr(names(i))).RefersToRange.Formula: Next i
    oldRef = ThisWorkbook.Names.Item("rngPlotAnnotationSettings").RefersTo
    oldAlerts = Application.DisplayAlerts
    Set annotations = ThisWorkbook.Names.Item("rngPlotAnnotationSettings").RefersToRange
    Set profiles = ThisWorkbook.Names.Item("rngCalculationProfiles").RefersToRange
    Set system = ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange
    Set loads = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange
    TestConfigureProfiles profiles: TestEnableChecks profiles, 1, "Yes", "No", "No", "No"
    TestConfigureDurationUnits "N", "N*mm", "Tension", "+Y tension", "+X tension"
    TestSetTableValue system, "Calculation.ZeroMomentPerDepth", "0", 2
    TestSetTableValue system, "Load.ReferenceOffsetX", "0", 2: TestSetTableValue system, "Load.ReferenceOffsetY", "0", 2
    TestSetTableValue system, "Plot.LoadCase", "PROFILE_CONFIG", 2
    TestConfigureAnnotationFixture annotations
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    TestBuildFixture section, provider, settings.GetBoolean("General.DiagramExtension", False)
    section.Annotations.AddDimension "AuditDim", -100#, 0#, 100#, 0#, 0#, 1#, "AUDIT_DIM", 200#
    section.Annotations.AddRebarLabel "AuditRebar", -80#, -60#, 80#, -60#, 0#, -1#, "AUDIT_REBAR"
    TestSetLoad loads, 1, -50000#, 1000000#, 2000000#, "Auto"
    Set batch = TestExecuteFixture(section, provider, settings, units)
    Set writer = New CNDMResultsWriter: writer.WriteResults ThisWorkbook, section, PrepareSectionSnapshot(section, provider), batch, units
    Set reader = New CSectionPlotDataReader: reader.LoadFromWorkbook ThisWorkbook, settings
    before = ThisWorkbook.Names.Item("rngNDMElementResults").RefersToRange.CurrentRegion.Value2
    solves = SectionEquilibriumSolveCount()
    TestAnnotationLayoutValues stats, annotations
    TestAnnotationInputMatrix stats, annotations, reader
    TestAnnotationChartValues stats, annotations, reader
    TestPresentationSnapshotIntegrity stats, reader
    TestAnnotationRelocation stats, annotations, reader, sheet
    TestCheck stats, "presentationConfig.noSolve", SectionEquilibriumSolveCount() = solves
    TestCheck stats, "presentationConfig.snapshotUnchanged", TestTablesEqual(before, _
        ThisWorkbook.Names.Item("rngNDMElementResults").RefersToRange.CurrentRegion.Value2)
    GoTo Restore
FailedRun:
    stats.Failed = stats.Failed + 1
    TestLogLine stats, "FAIL: presentationConfig.runtime; " & CStr(Err.Number) & "; " & Err.Description
Restore:
    On Error Resume Next
    ThisWorkbook.Names.Item("rngPlotAnnotationSettings").RefersTo = oldRef
    For i = 0 To UBound(names)
        Err.Clear: ThisWorkbook.Names.Item(CStr(names(i))).RefersToRange.Formula = saved(i)
        If Err.Number <> 0 Then restoreNumber = Err.Number: restoreReason = Err.Description
    Next i
    If Not sheet Is Nothing Then Application.DisplayAlerts = False: sheet.Delete
    Application.DisplayAlerts = oldAlerts
    On Error GoTo 0
    If restoreNumber <> 0 Then stats.Failed = stats.Failed + 1: TestLogLine stats, "FAIL: presentationConfig.restore; " & CStr(restoreNumber) & "; " & restoreReason
    TestLogLine stats, "TOTAL_AUDIT03_PRESENTATION_CONFIG: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed) & "; cases=" & CStr(stats.Cases)
    passed = stats.Passed: failed = stats.Failed: RunAudit03PresentationConfigTests = stats.Report
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Задает известную базу оформления в существующей двухколоночной таблице.
' Символы "-" неприменимых полей не превращаются в пользовательские настройки.
Private Sub TestConfigureAnnotationFixture(ByVal source As Object)
    Dim column As Long, field As Variant
    For column = 2 To 3
        For Each field In Array("Enabled", "LineEnabled"): TestAnnotationCell(source, CStr(field), column).Value2 = "Yes": Next field
        TestAnnotationCell(source, "Placement", column).Value2 = "Outside"
        TestAnnotationCell(source, "Offset", column).Value2 = 50#
        TestAnnotationCell(source, "LineWeight", column).Value2 = 2#
        TestAnnotationCell(source, "TextUnits", column).Value2 = "pt"
        TestAnnotationCell(source, "TextHeight", column).Value2 = 13#
        TestAnnotationCell(source, "TextGap", column).Value2 = 9#
        TestAnnotationCell(source, "Color", column).Value2 = "20,30,90"
    Next column
    TestAnnotationCell(source, "LineEnabled", 3).Value2 = "-"
    TestAnnotationCell(source, "ExtensionLineWeight", 2).Value2 = "-"
    TestAnnotationCell(source, "ExtensionLineWeight", 3).Value2 = 1#
    TestAnnotationCell(source, "ArrowType", 2).Value2 = "-": TestAnnotationCell(source, "ArrowType", 3).Value2 = "Triangle"
    TestAnnotationCell(source, "ArrowSize", 2).Value2 = "-": TestAnnotationCell(source, "ArrowSize", 3).Value2 = "Medium"
    TestAnnotationCell(source, "ExtensionLineColor", 2).Value2 = "-"
    TestAnnotationCell(source, "ExtensionLineColor", 3).Value2 = "140,140,140"
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Находит действительную входную ячейку аннотации, не используя ее адрес Config.
Private Function TestAnnotationCell(ByVal source As Object, ByVal field As String, ByVal column As Long) As Object
    Dim row As Long
    For row = 2 To source.Rows.Count
        If CStr(source.Cells(row, 1).Value2) = field Then Set TestAnnotationCell = source.Cells(row, column): Exit Function
    Next row
    Err.Raise vbObjectError + 4502, "AnnotationCell", "В таблице аннотаций нет " & field
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Независимые численные ожидания: масштаб модели 0.43 pt/mm, visual-scale=1.
' Проверяются настоящий цвет 0, направление Inside, толщина и mm/pt, а также
' повторная инициализация того же layout без накопления старых элементов.
Private Sub TestAnnotationLayoutValues(ByRef stats As TProfileStats, ByVal source As Object)
    Dim column As Long, color As Variant, height As Variant, layout As CPlotAnnotationLayout
    Dim settings As CSystemSettingsReader, prefix As String, base As Long
    For column = 2 To 3
        For Each color In Array("0,0,0", "255,0,0", "0,255,0", "0,0,255")
            TestConfigureAnnotationFixture source
            TestAnnotationCell(source, "Color", column).Value2 = color
            TestAnnotationCell(source, "ExtensionLineColor", 3).Value2 = color
            Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
            Set layout = New CPlotAnnotationLayout: layout.ConfigureFromSettings settings
            layout.Initialize -500#, 500#, -500#, 500#, 0#, 0#, 640#, 430#, "mm"
            layout.AddDimension -100#, 0#, 100#, 0#, 0#, 1#, "DIM"
            layout.AddRebarLabel -100#, 0#, 100#, 0#, 0#, 1#, "REBAR"
            Select Case CStr(color)
                Case "0,0,0": base = 0
                Case "255,0,0": base = 255
                Case "0,255,0": base = 65280
                Case "0,0,255": base = 16711680
            End Select
            prefix = "presentationConfig." & CStr(column) & ".Color." & CStr(color)
            If column = 2 Then TestCheck stats, prefix & ".text", layout.Color(6) = base Else TestCheck stats, prefix & ".text", layout.Color(4) = base
            TestCheck stats, prefix & ".extension", layout.Color(1) = base And layout.Color(2) = base
            TestCheckClose stats, prefix & ".lineWeight", layout.Weight(3), 2#, 0#
            TestCheckClose stats, prefix & ".offset", layout.Y1(3), 193.5, 0.000000000001
            layout.Initialize -500#, 500#, -500#, 500#, 0#, 0#, 640#, 430#, "mm"
            TestCheck stats, prefix & ".reset", layout.Count = 0
            stats.Cases = stats.Cases + 1
        Next color
        For Each height In Array(13#, 20#)
            TestConfigureAnnotationFixture source
            TestAnnotationCell(source, "TextHeight", column).Value2 = height
            TestAnnotationCell(source, "TextUnits", column).Value2 = "mm"
            TestAnnotationCell(source, "Placement", column).Value2 = "Inside"
            Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
            Set layout = New CPlotAnnotationLayout: layout.ConfigureFromSettings settings
            layout.Initialize -500#, 500#, -500#, 500#, 0#, 0#, 640#, 430#, "mm"
            layout.AddDimension -100#, 0#, 100#, 0#, 0#, 1#, "DIM"
            layout.AddRebarLabel -100#, 0#, 100#, 0#, 0#, 1#, "REBAR"
            prefix = "presentationConfig." & CStr(column) & ".TextUnits.mm." & CStr(height)
            If column = 2 Then base = 6 Else base = 4
            TestCheckClose stats, prefix & ".font", layout.FontSize(base), CDbl(height) * 0.43, 0.000000000001
            If column = 2 Then
                TestCheck stats, prefix & ".Inside", layout.Y1(6) > 215#
            Else
                TestCheck stats, prefix & ".Inside", layout.Y1(4) > 184.5
            End If
            stats.Cases = stats.Cases + 1
        Next height
    Next column
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Матрица всех применимых полей: пустота/текст/Excel-ошибка и особые границы.
' Ошибка активного поля должна содержать key, текущую ячейку и исправление;
' тот же ошибочный параметр выключенной группы не должен мешать схеме.
Private Sub TestAnnotationInputMatrix(ByRef stats As TProfileStats, ByVal source As Object, ByVal reader As CSectionPlotDataReader)
    Dim column As Long, field As Variant, value As Variant, invalids As Variant, cell As Object
    Dim prefix As String, key As String, code As Long, reason As String, caseIndex As Long, savedKey As Variant
    For column = 2 To 3
        For Each field In Array("Enabled", "Placement", "Offset", "LineEnabled", "LineWeight", "ExtensionLineWeight", _
                "TextUnits", "TextHeight", "TextGap", "ArrowType", "ArrowSize", "Color", "ExtensionLineColor")
            TestConfigureAnnotationFixture source
            Set cell = TestAnnotationCell(source, CStr(field), column)
            If CStr(cell.Value2) <> "-" Then
                If column = 2 Then key = "Plot.RebarLabels." & CStr(field) Else key = "Plot.Dimensions." & CStr(field)
                invalids = Array(vbNullString, "abc", CVErr(xlErrDiv0))
                Select Case CStr(field)
                    Case "Offset", "TextGap": invalids = Array(vbNullString, "abc", CVErr(xlErrDiv0), -1#)
                    Case "TextHeight", "LineWeight", "ExtensionLineWeight": invalids = Array(vbNullString, "abc", CVErr(xlErrDiv0), -1#, 0#)
                    Case "Color", "ExtensionLineColor": invalids = Array(vbNullString, "abc", CVErr(xlErrDiv0), "-1,2,3", "256,2,3", "1.5,2,3")
                End Select
                caseIndex = 0
                For Each value In invalids
                    TestConfigureAnnotationFixture source: cell.Value2 = value
                    prefix = "presentationConfig." & key & ".invalid" & CStr(caseIndex)
                    TestLogLine stats, "RUN: " & prefix
                    code = TestPresentationDrawError(reader, reason)
                    TestCheck stats, prefix & ".rejected", code <> 0
                    If code <> 0 Then
                        TestCheck stats, prefix & ".controlled", code <> 6 And code <> 13 And code <> 9
                        TestCheck stats, prefix & ".key", InStr(1, reason, key, vbTextCompare) > 0
                        TestCheck stats, prefix & ".address", InStr(1, reason, "ячейка " & cell.Address(False, False), vbBinaryCompare) > 0
                        TestCheck stats, prefix & ".repair", InStr(1, reason, "Введите", vbTextCompare) > 0 Or InStr(1, reason, "Выберите", vbTextCompare) > 0 Or InStr(1, reason, "Исправьте", vbTextCompare) > 0
                    End If
                    If CStr(field) <> "Enabled" Then
                        TestAnnotationCell(source, "Enabled", column).Value2 = "No"
                        code = TestPresentationDrawError(reader, reason)
                        TestCheck stats, prefix & ".inactive", code = 0
                    End If
                    caseIndex = caseIndex + 1: stats.Cases = stats.Cases + 1
                Next value
                TestConfigureAnnotationFixture source
                savedKey = cell.Offset(0, 1 - column).Value2: cell.Offset(0, 1 - column).Value2 = "__MissingAnnotationField"
                If CStr(field) <> "Enabled" Then
                    If column = 2 Then TestAnnotationCell(source, "Enabled", 3).Value2 = "No" Else TestAnnotationCell(source, "Enabled", 2).Value2 = "No"
                End If
                code = TestPresentationDrawError(reader, reason)
                TestCheck stats, "presentationConfig." & key & ".missing", code <> 0
                If code <> 0 Then TestCheck stats, "presentationConfig." & key & ".missingKey", InStr(1, reason, CStr(field), vbTextCompare) > 0
                cell.Offset(0, 1 - column).Value2 = savedKey
                TestConfigureAnnotationFixture source: code = TestPresentationDrawError(reader, reason)
                TestCheck stats, "presentationConfig." & key & ".recovery", code = 0
            End If
        Next field
    Next column
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет реальные фигуры Excel: чёрный цвет не подменяется,
' а Enabled и LineEnabled убирают именно соответствующие элементы.
Private Sub TestAnnotationChartValues(ByRef stats As TProfileStats, ByVal source As Object, ByVal reader As CSectionPlotDataReader)
    Dim shape As Object, chart As Object, code As Long, reason As String, count As Long, arrows As Long
    TestConfigureAnnotationFixture source
    TestAnnotationCell(source, "Color", 2).Value2 = "0,0,0": TestAnnotationCell(source, "Color", 3).Value2 = "0,0,0"
    TestAnnotationCell(source, "ExtensionLineColor", 3).Value2 = "0,0,0"
    code = TestPresentationDrawError(reader, reason): TestCheck stats, "presentationConfig.Chart.black.draw", code = 0
    If code <> 0 Then Exit Sub
    Set chart = ThisWorkbook.Worksheets.Item("Расчет").ChartObjects.Item("chtNDMSectionPlot").Chart
    For Each shape In chart.Shapes
        If InStr(1, shape.Name, "AnnotationLine", vbBinaryCompare) > 0 Then
            count = count + 1: TestCheck stats, "presentationConfig.Chart.black.line" & CStr(count), shape.Line.ForeColor.RGB = 0
            If shape.Line.BeginArrowheadStyle <> 1 Then arrows = arrows + 1
        End If
    Next shape
    TestCheck stats, "presentationConfig.Chart.black.lines", count = 4
    TestCheck stats, "presentationConfig.Chart.black.arrows", arrows = 1
    TestCheck stats, "presentationConfig.Chart.black.image", chart.Export(ThisWorkbook.Path & "\Audit03_presentation_black.png", "PNG")
    TestLogLine stats, "PLOT_IMAGE: " & ThisWorkbook.Path & "\Audit03_presentation_black.png"
    TestAnnotationCell(source, "Enabled", 3).Value2 = "No": TestAnnotationCell(source, "LineEnabled", 2).Value2 = "No"
    code = TestPresentationDrawError(reader, reason): TestCheck stats, "presentationConfig.Chart.disabled.draw", code = 0
    count = 0
    For Each shape In chart.Shapes
        If InStr(1, shape.Name, "AnnotationLine", vbBinaryCompare) > 0 Then count = count + 1
    Next shape
    TestCheck stats, "presentationConfig.Chart.disabled.lines", count = 0
    TestConfigureAnnotationFixture source
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Перемещает существующий Name и проверяет ошибку активного цвета по новому
' адресу. Отрисовка не должна выдавать старый адрес или молча брать default.
Private Sub TestAnnotationRelocation(ByRef stats As TProfileStats, ByVal source As Object, _
        ByVal reader As CSectionPlotDataReader, ByRef sheet As Object)
    Dim target As Object, cell As Object, code As Long, reason As String, oldRef As String
    oldRef = ThisWorkbook.Names.Item("rngPlotAnnotationSettings").RefersTo
    TestConfigureAnnotationFixture source
    Set sheet = ThisWorkbook.Worksheets.Add: sheet.Name = "__AuditPlotInputs"
    Set target = sheet.Cells(25, 60).Resize(source.Rows.Count, source.Columns.Count): target.Value2 = source.Value2
    ThisWorkbook.Names.Item("rngPlotAnnotationSettings").RefersTo = "='" & sheet.Name & "'!" & target.Address
    Set cell = TestAnnotationCell(target, "Color", 3): cell.Value2 = "256,0,0"
    code = TestPresentationDrawError(reader, reason)
    TestCheck stats, "presentationConfig.relocated.rejected", code <> 0
    If code <> 0 Then
        TestCheck stats, "presentationConfig.relocated.sheet", InStr(1, reason, sheet.Name, vbBinaryCompare) > 0
        TestCheck stats, "presentationConfig.relocated.address", InStr(1, reason, "ячейка " & cell.Address(False, False), vbBinaryCompare) > 0
    End If
    ThisWorkbook.Names.Item("rngPlotAnnotationSettings").RefersTo = oldRef
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Строит обычную схему по готовым сохранённым данным.
' Номер ошибки записывает отдельно от текста; статус из пояснения не выводится.
Private Function TestPresentationDrawError(ByVal reader As CSectionPlotDataReader, ByRef reason As String) As Long
    Dim settings As CSystemSettingsReader, plotter As CSectionPlotter
    reason = vbNullString
    On Error GoTo Failed
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set plotter = New CSectionPlotter: plotter.Draw ThisWorkbook, reader, settings
    Exit Function
Failed:
    TestPresentationDrawError = Err.Number: reason = Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет целостность реальных таблиц Results: недостающая обязательная
' колонка и испорченная аннотация не выдаются за успешно прочитанную схему.
' Принимается единственная текущая шапка геометрии, пустая таблица
' semantic-аннотаций не блокирует саму расчетную геометрию.
Private Sub TestPresentationSnapshotIntegrity(ByRef stats As TProfileStats, ByRef reader As CSectionPlotDataReader)
    Dim geometry As Object, annotations As Object, oldGeometry As Variant, oldAnnotations As Variant
    Dim data As Variant, code As Long, reason As String, field As Variant, column As Long
    Dim savedNumber As Long, savedReason As String
    On Error GoTo Failed
    Set geometry = TestPresentationSnapshotRange("rngNDMSectionGeometry", 25)
    Set annotations = TestPresentationSnapshotRange("rngNDMSectionAnnotations", 13)
    oldGeometry = geometry.Formula: oldAnnotations = annotations.Formula
    data = geometry.Value2: data(1, 7) = "__MissingGeometryStatus": geometry.Value2 = data
    code = TestPresentationReaderError(reader, reason)
    TestCheck stats, "presentationConfig.snapshot.geometryHeader.rejected", code <> 0
    TestCheck stats, "presentationConfig.snapshot.geometryHeader.controlled", code <> 9 And code <> 13 And code <> 6
    If code <> 0 Then TestCheck stats, "presentationConfig.snapshot.geometryHeader.reason", InStr(1, reason, "GeometryInterpretationStatus", vbBinaryCompare) > 0
    data(1, 7) = "ShapeType": geometry.Value2 = data
    code = TestPresentationReaderError(reader, reason): TestCheck stats, "presentationConfig.snapshot.obsoleteShapeType.rejected", code <> 0
    geometry.Formula = oldGeometry
    For Each field In Array("StartX", "StartY", "EndX", "EndY", "OutsideNormalX", "OutsideNormalY", "Text")
        data = annotations.Value2
        column = TestPresentationColumn(data, CStr(field))
        data(2, column) = CVErr(xlErrValue): annotations.Value2 = data
        code = TestPresentationReaderError(reader, reason)
        TestCheck stats, "presentationConfig.snapshot.annotation." & CStr(field) & ".rejected", code <> 0
        If code <> 0 Then
            TestCheck stats, "presentationConfig.snapshot.annotation." & CStr(field) & ".controlled", code <> 9 And code <> 13 And code <> 6
            TestCheck stats, "presentationConfig.snapshot.annotation." & CStr(field) & ".reason", InStr(1, reason, CStr(field), vbBinaryCompare) > 0
            TestCheck stats, "presentationConfig.snapshot.annotation." & CStr(field) & ".address", _
                InStr(1, reason, annotations.Cells(2, column).Address(False, False), vbBinaryCompare) > 0
        End If
        annotations.Formula = oldAnnotations
        code = TestPresentationReaderError(reader, reason)
        TestCheck stats, "presentationConfig.snapshot.annotation." & CStr(field) & ".recovery", code = 0
        stats.Cases = stats.Cases + 1
    Next field
    data = annotations.Value2: column = TestPresentationColumn(data, "Text")
    data(2, column) = "  Подпись с пробелами  ": annotations.Value2 = data
    code = TestPresentationReaderError(reader, reason)
    TestCheck stats, "presentationConfig.snapshot.textWhitespace.reader", code = 0
    If code = 0 Then TestCheck stats, "presentationConfig.snapshot.textWhitespace.preserved", reader.AnnotationText(1) = "  Подпись с пробелами  "
    annotations.Formula = oldAnnotations
    data = annotations.Value2: data(1, 4) = "__MissingStartX": annotations.Value2 = data
    code = TestPresentationReaderError(reader, reason)
    TestCheck stats, "presentationConfig.snapshot.annotationHeader.rejected", code <> 0
    annotations.Formula = oldAnnotations
    annotations.Offset(1, 0).Resize(annotations.Rows.Count - 1, annotations.Columns.Count).ClearContents
    code = TestPresentationReaderError(reader, reason)
    TestCheck stats, "presentationConfig.snapshot.emptyAnnotations.reader", code = 0
    If code = 0 Then TestCheck stats, "presentationConfig.snapshot.emptyAnnotations.count", reader.AnnotationCount = 0
    GoTo Restore
Failed:
    savedNumber = Err.Number: savedReason = Err.Description
Restore:
    On Error Resume Next
    If Not geometry Is Nothing Then geometry.Formula = oldGeometry
    If Not annotations Is Nothing Then annotations.Formula = oldAnnotations
    On Error GoTo 0
    If savedNumber <> 0 Then stats.Failed = stats.Failed + 1: TestLogLine stats, "FAIL: presentationConfig.snapshot.runtime; " & CStr(savedNumber) & "; " & savedReason
    code = TestPresentationReaderError(reader, reason): TestCheck stats, "presentationConfig.snapshot.recovery", code = 0
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Изменяет только собственный блок от именованной начальной ячейки.
' CurrentRegion не используется из-за соседних служебных строк.
' Число столбцов задано независимо по проверяемой структуре.
Private Function TestPresentationSnapshotRange(ByVal name As String, ByVal columns As Long) As Object
    Dim anchor As Object, rows As Long
    Set anchor = ThisWorkbook.Names.Item(name).RefersToRange
    rows = 1
    Do While Len(CStr(anchor.Offset(rows, 0).Value2)) > 0: rows = rows + 1: Loop
    Set TestPresentationSnapshotRange = anchor.Resize(rows, columns)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Находит столбец эталонной таблицы по точному заголовку.
' Проверяемый рабочий ResultColumn для подготовки эталона не вызывается.
Private Function TestPresentationColumn(ByRef data As Variant, ByVal header As String) As Long
    Dim column As Long
    For column = 1 To UBound(data, 2)
        If CStr(data(1, column)) = header Then TestPresentationColumn = column: Exit Function
    Next column
    Err.Raise vbObjectError + 4503, "PresentationColumn", "В snapshot нет " & header
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает точную ошибку загрузки на новом reader-е и сохраняет объект для
' проверки последующего восстановления; пользовательские настройки не меняет.
Private Function TestPresentationReaderError(ByRef reader As CSectionPlotDataReader, ByRef reason As String) As Long
    Dim settings As CSystemSettingsReader
    reason = vbNullString
    On Error GoTo Failed
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set reader = New CSectionPlotDataReader: reader.LoadFromWorkbook ThisWorkbook, settings
    Exit Function
Failed:
    TestPresentationReaderError = Err.Number: reason = Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' независимо вычисляет выбранную величину каждого элемента
' по сохранённому полю деформаций и его диаграмме. Чтение Results не должно
' подменять напряжение деформацией или брать другое НДС.
Private Sub TestCheckVisualizationValues(ByRef stats As TProfileStats, ByVal prefix As String, _
        ByVal reader As CSectionPlotDataReader, ByVal state As CSectionStateResult, _
        ByVal section As CSectionModel, ByVal provider As CMaterialModelProvider, ByVal units As CUnitSystem)
    Dim concrete As CMaterialDiagram, steel As CMaterialDiagram, i As Long, bar As Long
    Dim x As Double, y As Double, expected As Double, strain As Double, diagram As CMaterialDiagram
    ' ExtensionUsed отмечает выход за физическую область, а не выбор диаграммы.
    ' Ожидаемые значения схемы считаем по фактическим диаграммам этого НДС.
    TestCheck stats, prefix & ".materialSnapshot", state.HasMaterialSnapshot
    If Not state.HasMaterialSnapshot Then Exit Sub
    Set concrete = state.ConcreteDiagramSnapshot
    Set steel = state.SteelDiagramSnapshot
    For i = 1 To reader.Count
        If i <= section.ConcreteCount Then
            x = section.ConcreteX(i): y = section.ConcreteY(i): Set diagram = concrete
        Else
            bar = i - section.ConcreteCount
            x = section.RebarX(bar): y = section.RebarY(bar): Set diagram = steel
        End If
        strain = state.Epsilon0 + state.KappaX * y + state.KappaY * x
        If reader.VisualizationQuantity = "Strain" Then
            expected = strain
        Else
            expected = units.InternalStressToOutput(diagram.GetStress(strain))
        End If
        TestCheckClose stats, prefix & ".element" & CStr(i), reader.ResultValue(i), expected, 0.000000000001
    Next i
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' проверяет именно текст растянутого края общей легенды, а не
' случайное совпадение подписи нуля в Chart. Положение определяется
' контрактом оформления; точность формируется независимым test-oracle.
Private Sub TestCheckVisualizationLegend(ByRef stats As TProfileStats, ByVal prefix As String, _
        ByVal reader As CSectionPlotDataReader, ByVal settings As CSystemSettingsReader)
    Dim plotter As CSectionPlotter, chartObject As Object, shape As Object, i As Long
    Dim maximum As Double, foundValue As Boolean, foundText As Boolean, text As String, pattern As String
    For i = 1 To reader.Count
        If reader.PhysicalState(i) = "Tension" Then
            If Not foundValue Or Abs(reader.ResultValue(i)) > Abs(maximum) Then maximum = reader.ResultValue(i)
            foundValue = True
        End If
    Next i
    TestCheck stats, prefix & ".legendFixtureTension", foundValue
    If Not foundValue Then Exit Sub
    pattern = "0": If reader.ResultPrecision > 0 Then pattern = pattern & "." & String$(reader.ResultPrecision, "0")
    text = Replace$(Format$(maximum, pattern), ",", ".")
    Set plotter = New CSectionPlotter: plotter.Draw ThisWorkbook, reader, settings
    Set chartObject = ThisWorkbook.Worksheets.Item("Расчет").ChartObjects.Item("chtNDMSectionPlot")
    For Each shape In chartObject.Chart.Shapes
        If Left$(shape.Name, Len("NDMPlot_Text")) = "NDMPlot_Text" Then
            If Abs(shape.Left - (chartObject.Width - 118#)) < 0.01 And Abs(shape.Top - 174#) < 0.01 Then
                foundText = True
                TestCheck stats, prefix & ".actualLegendText", CStr(shape.TextFrame.Characters().Text) = text
            End If
        End If
    Next shape
    TestCheck stats, prefix & ".actualLegendPresent", foundText
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' короткая проверка обязательного сопоставления номеров сочетаний.
Public Function RunStabilityDurationMatchingTests() As String
    RunStabilityDurationMatchingTests = RunAudit03DurationConfigTests(matchingOnly:=True)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет отсутствие строки отдельно от допустимых пустых и нулевых усилий.
' Для СП 35/63 проверяет статус, изоляцию сочетаний, отключённую устойчивость
' и фактический адрес таблицы после её переноса на другой лист.
Private Sub TestDurationMatching(ByRef stats As TProfileStats, ByVal profiles As Object, ByVal system As Object, _
        ByVal loads As Object, ByVal duration As Object, ByVal section As CSectionModel, ByVal provider As CMaterialModelProvider, _
        ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem, ByRef sheet As Object)
    Dim standard As Variant, mode As Long, data As Variant, loadValues As Variant, column As Long
    Dim batch As CBatchSectionCalculator, code As Long, reason As String, prefix As String, originalRef As String
    Dim location As String, target As Object, missing As Boolean, inactive As Boolean
    originalRef = ThisWorkbook.Names.Item("rngStabilityDurationLoads").RefersTo
    For Each standard In Array("SP63", "SP35")
        TestSetTableValue system, "Stability.Code", CStr(standard), 2
        settings.LoadFromWorkbook ThisWorkbook
        For mode = 1 To 12
            TestPrepareDurationLoads loads, duration, 1
            TestEnableChecks profiles, 1, "Yes", "No", "No", "Yes"
            TestEnableChecks profiles, 2, "Yes", "No", "No", "No"
            data = duration.Value2: missing = False: inactive = False
            Select Case mode
                Case 1, 6, 11, 12
                    data(2, 1) = "OTHER_NUMBER": missing = True
                Case 2
                    For column = 2 To 4: data(2, column) = vbNullString: Next column
                Case 3
                    For column = 2 To 4: data(2, column) = 0#: Next column
                Case 4
                    data(2, 1) = vbNullString: missing = True
                    For column = 2 To 4: data(2, column) = CVErr(2015): Next column
                Case 5, 10
                    data(2, 1) = "OTHER_NUMBER": inactive = True
                    For column = 2 To 4: data(2, column) = CVErr(2015): Next column
                    If mode = 5 Then TestEnableChecks profiles, 1, "Yes", "No", "No", "No"
                Case 7
                    For column = 2 To 4: data(2, column) = "=" & Chr$(34) & Chr$(34): Next column
                Case 8
                    data(2, 2) = vbNullString: data(2, 3) = 0#: data(2, 4) = vbNullString
                Case 9
                    data(2, 1) = " duration_1 "
                    For column = 2 To 4: data(2, column) = 0#: Next column
            End Select
            duration.Value2 = data
            If mode = 7 Then duration.Calculate
            If mode = 6 Or mode = 10 Then TestAddDurationRecoveryLoad loads, duration, 1
            If mode = 10 Then
                loadValues = loads.Value2: loadValues(2, 5) = "PR2": loads.Value2 = loadValues
            End If
            location = duration.Worksheet.Name & "!" & duration.Address(False, False)
            If mode >= 11 Then
                If sheet Is Nothing Then Set sheet = ThisWorkbook.Worksheets.Add: sheet.Name = "__AuditDurationInputs"
                If mode = 11 Then Set target = sheet.Cells(10, 5) Else Set target = sheet.Cells(800, 60)
                Set target = target.Resize(duration.Rows.Count, duration.Columns.Count)
                target.Value2 = duration.Value2
                ThisWorkbook.Names.Item("rngStabilityDurationLoads").RefersTo = "='" & sheet.Name & "'!" & target.Address
                location = sheet.Name & "!" & target.Address(False, False)
            End If
            prefix = "durationMatching." & CStr(standard) & ".case" & CStr(mode)
            TestLogLine stats, "RUN: " & prefix
            code = TestExecuteDuration(section, provider, settings, units, batch, reason)
            TestCheck stats, prefix & ".readerNoAbort", code = 0
            If code = 0 Then
                If missing Then
                    TestCheckDurationInputError stats, prefix, batch, units, location
                    TestCheck stats, prefix & ".combinationInReason", InStr(1, batch.ResultAt(1).StabilityMeta.ResultComment, "DURATION_1", vbBinaryCompare) > 0
                    TestCheck stats, prefix & ".missingReason", InStr(1, batch.ResultAt(1).StabilityMeta.ResultComment, "не найдена строка", vbBinaryCompare) > 0
                    TestCheck stats, prefix & ".overallInputErr", batch.ResultAt(1).Status = "InputErr"
                ElseIf inactive Then
                    TestCheck stats, prefix & ".notRequested", batch.ResultAt(1).StabilityMeta.InternalStatus = rsNotRequested
                    TestCheck stats, prefix & ".noInputErr", batch.ResultAt(1).Status <> "InputErr"
                    TestCheckComments stats, prefix, batch, units
                Else
                    TestCheck stats, prefix & ".stabilitySuccess", batch.ResultAt(1).StabilityMeta.InternalStatus = rsSuccess
                    TestCheck stats, prefix & ".calculated", batch.ResultAt(1).StabilityMeta.Calculated
                    TestCheckClose stats, prefix & ".sustainedNZero", batch.ResultAt(1).StabilityResult.SustainedN, 0#, 0.00000001
                    TestCheckClose stats, prefix & ".sustainedMxZero", batch.ResultAt(1).StabilityResult.SustainedMoment1, 0#, 0.00000001
                    TestCheckClose stats, prefix & ".sustainedMyZero", batch.ResultAt(1).StabilityResult.SustainedMoment2, 0#, 0.00000001
                    TestCheckComments stats, prefix, batch, units
                End If
                TestCheck stats, prefix & ".directContinues", batch.ResultAt(1).DirectStateMeta.InternalStatus = rsSuccess
                If mode = 6 Or mode = 10 Then
                    TestCheck stats, prefix & ".recoverySuccess", batch.Count = 2 And batch.ResultAt(2).StabilityMeta.InternalStatus = rsSuccess
                    TestCheck stats, prefix & ".recoveryNoInputErr", batch.ResultAt(2).Status <> "InputErr"
                End If
            End If
            ThisWorkbook.Names.Item("rngStabilityDurationLoads").RefersTo = originalRef
            stats.Cases = stats.Cases + 1
        Next mode
    Next standard
    TestEnableChecks profiles, 1, "No", "No", "No", "Yes"
    TestSetTableValue system, "Stability.Code", "SP63", 2
    settings.LoadFromWorkbook ThisWorkbook
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Выполняет только изменённую числовую границу duration; полный набор сохраняет прежний охват всех строк.
Public Function RunLocaleDurationTests(Optional ByRef passed As Long = 0, Optional ByRef failed As Long = 0) As String
    RunLocaleDurationTests = RunAudit03DurationConfigTests(passed, failed, False, True)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Оба десятичных написания дополнительных нагрузок дают прежние Н/Н*мм, знаки и пояснения.
Private Sub TestDurationLocaleValid(ByRef stats As TProfileStats, ByVal loads As Object, ByVal duration As Object, _
        ByVal section As CSectionModel, ByVal provider As CMaterialModelProvider, _
        ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem)
    Dim separator As Variant, column As Long, data As Variant, batch As CBatchSectionCalculator
    Dim code As Long, reason As String, prefix As String, oldUse As Boolean, oldDecimal As String, oldThousands As String, oldFormat As Variant
    oldUse = Application.UseSystemSeparators: oldDecimal = Application.DecimalSeparator: oldThousands = Application.ThousandsSeparator
    On Error GoTo Failed
    For Each separator In Array(".", ",")
        Application.UseSystemSeparators = False: Application.DecimalSeparator = CStr(separator)
        If separator = "." Then Application.ThousandsSeparator = "," Else Application.ThousandsSeparator = "."
        For column = 2 To 4
            TestPrepareDurationLoads loads, duration, 1
            data = duration.Value2: data(2, column) = CStr(data(2, column)) & CStr(separator) & "0"
            oldFormat = duration.Cells(2, column).NumberFormat
            duration.Cells(2, column).NumberFormat = "@": duration.Value2 = data
            duration.Cells(2, column).NumberFormat = oldFormat
            prefix = "durationLocale." & CStr(separator) & "." & column
            code = TestExecuteDuration(section, provider, settings, units, batch, reason)
            TestCheck stats, prefix & ".reader", code = 0
            If code = 0 Then TestCheckDurationResult stats, prefix, batch, units, -25000#, 400000#, 600000#: TestCheckComments stats, prefix, batch, units
            stats.Cases = stats.Cases + 1
        Next column
    Next separator
    GoTo Cleanup
Failed:
    TestCheck stats, "durationLocale.runtime." & Err.Number & "." & Err.Description, False
Cleanup:
    On Error Resume Next
    Application.UseSystemSeparators = False: Application.DecimalSeparator = oldDecimal
    Application.ThousandsSeparator = oldThousands: Application.UseSystemSeparators = oldUse
    On Error GoTo 0
End Sub
