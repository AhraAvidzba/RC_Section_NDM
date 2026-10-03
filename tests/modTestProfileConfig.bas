Attribute VB_Name = "modTestProfileConfig"
Option Explicit

' ==========================================================================
' ДЛЯ ТЕСТОВ: МАТЕРИАЛЬНЫЕ И ВИЗУАЛЬНЫЕ ПОЛЯ РАСЧЕТНЫХ ПРОФИЛЕЙ
' ==========================================================================
' Проверяет реальные ячейки PR1-PR4, передачу material spec в State/Search
' и выбор сохраненных результатов для схемы. Ожидаемая спецификация задается
' независимо от прочитанного профиля; равновесие пересчитывается по элементам.
' Тесты не вводят новую физику, не меняют допуски production и восстанавливают
' все затронутые входные таблицы даже после ошибки. Новых классов нет.

Private Type TProfileStats
    Passed As Long
    Failed As Long
    Cases As Long ' Варианты поля/профиля, не число assertions.
    Report As String
End Type

' ============================== ДЛЯ ТЕСТОВ ==============================

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
    SetTableValue unitsRange, "Length", "mm", 2
    SetTableValue unitsRange, "Force", "N", 2
    SetTableValue unitsRange, "Moment", "N*mm", 2
    SetTableValue unitsRange, "Stress", "MPa", 2
    SetTableValue signs, "+N", "Tension", 2
    SetTableValue signs, "+Mx", "+Y tension", 2
    SetTableValue signs, "+My", "+X tension", 2
    SetTableValue system, "Calculation.ZeroMomentPerDepth", "0", 2
    SetTableValue system, "Load.ReferenceOffsetX", "0", 2
    SetTableValue system, "Load.ReferenceOffsetY", "0", 2
    SetTableValue system, "Solver.ToleranceN", "0.01", 2
    SetTableValue system, "Solver.ToleranceMx", "1", 2
    SetTableValue system, "Solver.ToleranceMy", "1", 2
    SetTableValue system, "SLS.Crack.PsiMode", "AlwaysCalc", 2
    SetTableValue system, "Capacity.SolutionStrategy", "Auto", 2
    SetTableValue system, "Stability.Code", "SP35", 2
    SetTableValue system, "Stability.ElementLength", "1000", 2
    SetTableValue system, "Stability.Mu1", "1", 2
    SetTableValue system, "Stability.Mu2", "1", 2
    SetTableValue system, "Stability.AccidentalEccentricityMode", "User", 2
    SetTableValue system, "Stability.AccidentalEccentricityPlanes", "BothPlanes", 2
    SetTableValue system, "Stability.AccidentalEccentricityUser1", "200", 2
    SetTableValue system, "Stability.AccidentalEccentricityUser2", "200", 2
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    BuildFixture section, provider, settings.GetBoolean("General.DiagramExtension", False)
    LogLine stats, "PROFILE_CONFIG: effectiveExtension=" & CStr(provider.DiagramExtensionEnabled)
    TestMaterialSpecs stats, source, loads, section, provider, settings, units
    TestStabilityValueSets stats, source, loads, section, provider, settings, units
    TestProfileMetadata stats, source, loads, section, provider, settings, units
    TestVisualization stats, source, system, loads, section, provider, settings, units
    TestInvalidProfileCells stats, source
    GoTo Restore
FailedRun:
    stats.Failed = stats.Failed + 1
    LogLine stats, "FAIL: profileConfig.runtime; " & CStr(Err.Number) & "; " & Err.Description
Restore:
    On Error Resume Next
    For i = 0 To UBound(names)
        ThisWorkbook.Names.Item(CStr(names(i))).RefersToRange.Formula = saved(i)
    Next i
    On Error GoTo 0
    LogLine stats, "TOTAL_AUDIT03_PROFILE_CONFIG: passed=" & CStr(stats.Passed) & _
        "; failed=" & CStr(stats.Failed) & "; cases=" & CStr(stats.Cases)
    passed = stats.Passed: failed = stats.Failed
    RunAudit03ProfileConfigTests = stats.Report
End Function

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
    On Error GoTo FailedRun
    names = Array("rngSystemSettings", "rngUnitSettings", "rngSignConventionSettings", _
        "rngConcreteMaterialParameters", "rngSteelMaterialParameters", "rngCircleGeometry", _
        "rngRoundedRectangleGeometry", "rngHollowRectangleGeometry", "rngRectSetGeometry", _
        "rngPlotAnnotationSettings", "rngCalculationProfiles", "rngLoadCombinations")
    ReDim refs(0 To UBound(names))
    For index = 0 To UBound(names): refs(index) = ThisWorkbook.Names.Item(CStr(names(index))).RefersTo: Next index
    Set profiles = ThisWorkbook.Names.Item("rngCalculationProfiles").RefersToRange
    Set loads = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange
    savedProfiles = profiles.Formula: savedLoads = loads.Formula
    ConfigureProfiles profiles: EnableChecks profiles, 1, "Yes", "Yes", "Yes", "Yes"
    SetLoad loads, 1, -50000#, 1000000#, 2000000#, "Auto"
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    BuildFixture section, provider, settings.GetBoolean("General.DiagramExtension", False)
    oldAlerts = Application.DisplayAlerts
    Set sheet = ThisWorkbook.Worksheets.Add: sheet.Name = "__AuditMovedInputs"
    For index = 0 To UBound(names)
        Set source = ThisWorkbook.Names.Item(CStr(names(index))).RefersToRange
        baseline = source.Value2
        Set fields = RelocatedFields(source, CStr(names(index)))
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
                LogLine stats, "RUN: " & prefix & "; actual=" & sheet.Name & "!" & address
                If index < 10 Then
                    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
                    reason = settings.InputErrorMessage(key, "Контроль адреса.", "Проверьте значение.")
                    Check stats, prefix & ".registered", settings.HasKey(key)
                    Check stats, prefix & ".locationSheet", InStr(1, reason, sheet.Name, vbBinaryCompare) > 0
                    Check stats, prefix & ".locationAddress", InStr(1, reason, "ячейка " & address, vbBinaryCompare) > 0
                    cell.Value2 = CVErr(xlErrDiv0): code = SettingsLoadError(reason)
                ElseIf index = 10 Then
                    cell.Value2 = "NOT_A_VALID_CHOICE": code = CatalogError(reason)
                Else
                    cell.Value2 = "abc"
                    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
                    Set batch = ExecuteFixture(section, provider, settings, units)
                    code = IIf(batch.ResultAt(1).Status = "InputErr", 1, 0)
                    reason = batch.ResultAt(1).OverallMeta.ResultComment
                    Check stats, prefix & ".noSolve", batch.SolverCallCount = 0
                    CheckComments stats, prefix, batch, units
                End If
                Check stats, prefix & ".rejected", code <> 0
                Check stats, prefix & ".errorSheet", InStr(1, reason, sheet.Name, vbBinaryCompare) > 0
                Check stats, prefix & ".errorAddress", InStr(1, reason, address, vbBinaryCompare) > 0
                Check stats, prefix & ".oldAddressAbsent", InStr(1, reason, "ячейка " & oldAddress, vbBinaryCompare) = 0
                LogLine stats, "RELOCATED_INPUT_MESSAGE: " & prefix & "; " & reason
                cell.Value2 = savedValue
                If index = 0 And key = "Solver.ToleranceN" Then
                    cell.Value2 = "abc"
                    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
                    Set batch = ExecuteFixture(section, provider, settings, units)
                    reason = batch.ResultAt(1).OverallMeta.ResultComment
                    Check stats, prefix & ".batchInputErr", batch.ResultAt(1).Status = "InputErr"
                    Check stats, prefix & ".batchAddress", InStr(1, reason, "ячейка " & address, vbBinaryCompare) > 0
                    Check stats, prefix & ".batchNoSolve", batch.SolverCallCount = 0
                    CheckComments stats, prefix & ".batch", batch, units
                    cell.Value2 = savedValue
                End If
                stats.Cases = stats.Cases + 1
            Next field
            If index = 8 Then TestRelocatedSharedSelectors stats, target, position
            ThisWorkbook.Names.Item(CStr(names(index))).RefersTo = refs(index)
        Next position
    Next index
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Check stats, "relocated.recovery.settings", settings.HasKey("Solver.ToleranceN")
    Dim catalog As CCalculationProfileCatalog
    Set catalog = New CCalculationProfileCatalog: catalog.LoadFromWorkbook ThisWorkbook
    Check stats, "relocated.recovery.profiles", catalog.Count = 4
    GoTo Restore
FailedRun:
    stats.Failed = stats.Failed + 1
    LogLine stats, "FAIL: relocated.runtime; " & CStr(Err.Number) & "; " & Err.Description
Restore:
    On Error Resume Next
    For index = 0 To UBound(names)
        If Len(refs(index)) > 0 Then ThisWorkbook.Names.Item(CStr(names(index))).RefersTo = refs(index)
    Next index
    profiles.Formula = savedProfiles: loads.Formula = savedLoads
    Application.DisplayAlerts = False
    If Not sheet Is Nothing Then sheet.Delete
    Application.DisplayAlerts = oldAlerts
    On Error GoTo 0
    LogLine stats, "TOTAL_AUDIT03_RELOCATED_INPUT: passed=" & CStr(stats.Passed) & _
        "; failed=" & CStr(stats.Failed) & "; cases=" & CStr(stats.Cases)
    passed = stats.Passed: failed = stats.Failed
    RunAudit03RelocatedInputTests = stats.Report
End Function

' Проверяет конфликт двух общих выборов и ошибку формулы во второй строке.
' Адреса обоих вводов берутся из перемещенной таблицы; исходные значения
' восстанавливаются после каждой пробы, до перехода к следующей паре граней.
Private Sub TestRelocatedSharedSelectors(ByRef stats As TProfileStats, ByVal source As Object, ByVal position As Long)
    Dim face As Variant, column As Variant, row As Long, first As Object, second As Object
    Dim savedFirst As Variant, savedSecond As Variant, code As Long, reason As String, prefix As String
    row = CaptionRow(source, "H1", 1) + 14
    For Each face In Array("H1", "B1", "H2", "B2")
        For Each column In Array(3, 4, 6, 7)
            Set first = source.Cells(row, CLng(column)): Set second = source.Cells(row + 1, CLng(column))
            savedFirst = first.Value2: savedSecond = second.Value2
            prefix = "relocated.shared.p" & CStr(position) & "." & CStr(face) & ".c" & CStr(column)
            first.Value2 = "FIRST_CHOICE": second.Value2 = "SECOND_CHOICE"
            code = SettingsLoadError(reason)
            Check stats, prefix & ".conflictRejected", code = vbObjectError + 4317
            Check stats, prefix & ".firstAddress", InStr(1, reason, "ячейка " & first.Address(False, False), vbBinaryCompare) > 0
            Check stats, prefix & ".secondAddress", InStr(1, reason, "ячейка " & second.Address(False, False), vbBinaryCompare) > 0
            Check stats, prefix & ".sheet", InStr(1, reason, source.Worksheet.Name, vbBinaryCompare) > 0
            first.Value2 = savedFirst: second.Value2 = CVErr(xlErrDiv0)
            code = SettingsLoadError(reason)
            Check stats, prefix & ".formulaRejected", code <> 0
            Check stats, prefix & ".formulaAddress", InStr(1, reason, "ячейка " & second.Address(False, False), vbBinaryCompare) > 0
            Check stats, prefix & ".formulaSheet", InStr(1, reason, source.Worksheet.Name, vbBinaryCompare) > 0
            first.Value2 = savedFirst: second.Value2 = savedSecond
            stats.Cases = stats.Cases + 1
        Next column
        row = row + 2
    Next face
End Sub

' Задает ожидаемые ключи независимо от private-методов reader-а. Позиции
' находятся по шапкам исходной таблицы, а абсолютные адреса не зашиты.
Private Function RelocatedFields(ByVal source As Object, ByVal rangeName As String) As Collection
    Dim fields As New Collection, row As Long, i As Long, column As Long
    Select Case rangeName
        Case "rngSystemSettings"
            fields.Add Array(CaptionRow(source, "Solver.ToleranceN", 1), 2, "Solver.ToleranceN")
            fields.Add Array(CaptionRow(source, "SLS.Crack.Allowable", 1), 2, "SLS.Crack.Allowable")
        Case "rngUnitSettings"
            fields.Add Array(CaptionRow(source, "Force", 1), 2, "Units.Force.Input")
            fields.Add Array(CaptionRow(source, "Force", 1), 4, "Units.Force.Output")
        Case "rngSignConventionSettings"
            fields.Add Array(CaptionRow(source, "+N", 1), 2, "Sign.N.User")
        Case "rngConcreteMaterialParameters"
            fields.Add Array(CaptionRow(source, "Concrete.R.ULS(I)", 1), 2, "Concrete.Rb.ULS")
            fields.Add Array(CaptionRow(source, "Concrete.R.ULS(I)", 1), 3, "Concrete.Rbt.ULS")
        Case "rngSteelMaterialParameters"
            fields.Add Array(CaptionRow(source, "Steel.E", 1), 2, "Steel.Esc")
            fields.Add Array(CaptionRow(source, "Steel.E", 1), 3, "Steel.Es")
        Case "rngCircleGeometry"
            fields.Add Array(CaptionRow(source, "Circle.Diameter", 1), 2, "Circle.Diameter")
            fields.Add Array(CaptionRow(source, "Rebar.AxisDistance", 1), 2, "Rebar.AxisDistance")
        Case "rngRoundedRectangleGeometry"
            fields.Add Array(CaptionRow(source, "B", 2) + 1, 2, "RoundedRectangle.B")
            row = CaptionRow(source, "Грань", 1) + 1
            fields.Add Array(row, 2, "RoundedRectangle.H.Left.as")
            fields.Add Array(row, 3, "RoundedRectangle.H.Left.d")
            fields.Add Array(row, 4, "RoundedRectangle.H.Left.n")
        Case "rngHollowRectangleGeometry"
            fields.Add Array(CaptionRow(source, "HollowRectangle.InnerOffsetX", 1), 2, "HollowRectangle.InnerOffsetX")
            row = CaptionRow(source, "Грань", 1) + 1
            fields.Add Array(row, 2, "HollowRectangle.H.Left.as")
            fields.Add Array(row, 3, "HollowRectangle.H.Left.d")
        Case "rngRectSetGeometry"
            fields.Add Array(CaptionRow(source, "RectSet.SectionType", 1), 2, "RectSet.SectionType")
            row = CaptionRow(source, "H1", 1) + 4
            fields.Add Array(row, 2, "RectSet.H1.as_1")
            fields.Add Array(row, 3, "RectSet.H1.d_1")
            fields.Add Array(row, 4, "RectSet.H1.n_1")
        Case "rngPlotAnnotationSettings"
            fields.Add Array(CaptionRow(source, "Enabled", 1), 2, "Plot.RebarLabels.Enabled")
            fields.Add Array(CaptionRow(source, "TextHeight", 1), 3, "Plot.Dimensions.TextHeight")
        Case "rngCalculationProfiles"
            For i = 1 To 4
                fields.Add Array(CaptionRow(source, "MaterialModel.Strength.ValueSet", 2), i + 2, "PR" & CStr(i) & ".MaterialModel.Strength.ValueSet")
                fields.Add Array(CaptionRow(source, "Visualization.StressPrecision", 2), i + 2, "PR" & CStr(i) & ".Visualization.StressPrecision")
            Next i
        Case "rngLoadCombinations"
            fields.Add Array(2, 2, "N")
    End Select
    Set RelocatedFields = fields
End Function

' Ищет подпись внутри фактической таблицы и не подменяет отсутствующую строку.
Private Function CaptionRow(ByVal source As Object, ByVal caption As String, ByVal column As Long) As Long
    Dim row As Long
    For row = 1 To source.Rows.Count
        If CStr(source.Cells(row, column).Value2) = caption Then CaptionRow = row: Exit Function
    Next row
    Err.Raise vbObjectError + 4501, "CaptionRow", "В тестовой таблице нет " & caption
End Function

' Сохраняет фактическую ошибку чтения полной книги после перемещения таблицы.
Private Function SettingsLoadError(ByRef reason As String) As Long
    On Error GoTo Failed
    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    reason = vbNullString
    Exit Function
Failed:
    SettingsLoadError = Err.Number: reason = Err.Description
End Function

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
                For Each value In MaterialVariants(CStr(role), CStr(field))
                    ConfigureProfiles source
                    SetProfileValue source, profile, "MaterialModel." & CStr(role) & "." & CStr(field), value
                    If CStr(role) = "Strength" Then
                        EnableChecks source, profile, "Yes", "Yes", "No", "No"
                        SetLoad loads, profile, -50000#, 10000000#, 2000000#, "LambdaMxy"
                    Else
                        EnableChecks source, profile, "No", "No", "Yes", "No"
                        SetLoad loads, profile, 200000#, 0#, 0#, "LambdaN"
                    End If
                    prefix = "profileConfig.PR" & CStr(profile) & ".MaterialModel." & CStr(role) & _
                        "." & CStr(field) & "." & CStr(value)
                    LogLine stats, "RUN: " & prefix
                    Set batch = ExecuteFixture(section, provider, settings, units)
                    Set result = batch.ResultAt(1)
                    Set expected = ExpectedSpec(CStr(role), CStr(field), CStr(value))
                    If CStr(role) = "Strength" Then
                        CheckState stats, prefix & ".direct", result.StateRepository.FindState(sstStrengthState), section, provider, expected
                        CheckState stats, prefix & ".capacity", result.StrengthResult.Capacity.StateResult, section, provider, expected
                    ElseIf CStr(role) = "CrackInitiation" Then
                        CheckState stats, prefix & ".pre", result.CrackResult.Formation.PreCrackState, section, provider, expected
                        CheckState stats, prefix & ".post", result.CrackResult.Formation.PostCrackState, section, provider, ExpectedSpec("CrackedState", "", "")
                    Else
                        CheckState stats, prefix & ".current", result.StateRepository.FindState(sstCrackedState), section, provider, expected
                        CheckState stats, prefix & ".post", result.CrackResult.Formation.PostCrackState, section, provider, expected
                    End If
                    CheckComments stats, prefix, batch, units
                    stats.Cases = stats.Cases + 1
                Next value
            Next field
        Next role
    Next profile
End Sub

' Подтверждает эффект набора сопротивлений в табличной ветви СП 35.
' Модули общие для ULS/SLS; здесь Nult зависит от выбранных Rb и Rsc.
Private Sub TestStabilityValueSets(ByRef stats As TProfileStats, ByVal source As Object, ByVal loads As Object, _
        ByVal section As CSectionModel, ByVal provider As CMaterialModelProvider, _
        ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem)
    Dim profile As Long, value As Variant, batch As CBatchSectionCalculator, baseline As Double, prefix As String
    SetTableValue ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange, "Stability.AccidentalEccentricityUser1", "5", 2
    SetTableValue ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange, "Stability.AccidentalEccentricityUser2", "5", 2
    settings.LoadFromWorkbook ThisWorkbook
    For profile = 1 To 4
        For Each value In Array("ULS(I)", "SLS(II)")
            ConfigureProfiles source
            EnableChecks source, profile, "No", "No", "No", "Yes"
            SetProfileValue source, profile, "MaterialModel.Stability.ValueSet", value
            SetLoad loads, profile, -50000#, 0#, 0#, "Auto"
            Set batch = ExecuteFixture(section, provider, settings, units)
            prefix = "profileConfig.PR" & CStr(profile) & ".MaterialModel.Stability.ValueSet." & CStr(value)
            Check stats, prefix & ".branch", batch.ResultAt(1).StabilityResult.Branch = "SP35-table"
            Check stats, prefix & ".calculated", batch.ResultAt(1).StabilityMeta.Calculated
            Check stats, prefix & ".noSolve", batch.SolverCallCount = 0
            Check stats, prefix & ".positiveLimit", batch.ResultAt(1).StabilityResult.Nultimate1 > 0#
            If CStr(value) = "ULS(I)" Then
                baseline = batch.ResultAt(1).StabilityResult.Nultimate1
            Else
                Check stats, prefix & ".resistanceEffect", batch.ResultAt(1).StabilityResult.Nultimate1 > baseline
            End If
            CheckComments stats, prefix, batch, units
            stats.Cases = stats.Cases + 1
        Next value
    Next profile
End Sub

' Пользовательское имя реально разрешает ProfileId строки LC. Описание является
' видимым пояснением профиля, а не физическим коэффициентом или новым calculator-ом.
Private Sub TestProfileMetadata(ByRef stats As TProfileStats, ByVal source As Object, ByVal loads As Object, _
        ByVal section As CSectionModel, ByVal provider As CMaterialModelProvider, _
        ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem)
    Dim profile As Long, catalog As CCalculationProfileCatalog, batch As CBatchSectionCalculator
    Dim caption As String, description As String, prefix As String, state As CSectionStateResult, data As Variant
    For profile = 1 To 4
        ConfigureProfiles source
        EnableChecks source, profile, "Yes", "No", "No", "No"
        caption = "Проверяемый профиль " & CStr(profile)
        description = "Пояснение профиля " & CStr(profile) & ": прямое НДС."
        SetProfileValue source, profile, "Profile.DisplayName", caption
        SetProfileValue source, profile, "Profile.Description", description
        SetLoad loads, profile, -50000#, 1000000#, 2000000#, "Auto"
        data = loads.Value2: data(2, 5) = caption: loads.Value2 = data
        Set catalog = New CCalculationProfileCatalog: catalog.LoadFromWorkbook ThisWorkbook
        Set batch = ExecuteFixture(section, provider, settings, units)
        Set state = batch.ResultAt(1).StateRepository.FindState(sstStrengthState)
        prefix = "profileConfig.PR" & CStr(profile)
        Check stats, prefix & ".Profile.DisplayName.reader", catalog.ProfileById(caption).ProfileId = "PR" & CStr(profile)
        Check stats, prefix & ".Profile.DisplayName.consumer", batch.ResultAt(1).ProfileId = "PR" & CStr(profile) And Not state Is Nothing
        Check stats, prefix & ".Profile.Description.metadata", catalog.ProfileById(caption).Description = description
        CheckState stats, prefix & ".metadataState", state, section, provider, ExpectedSpec("Strength", "", "")
        CheckComments stats, prefix, batch, units
        stats.Cases = stats.Cases + 2
    Next profile
End Sub

' Выбирает все пять named states и обе величины в реальном snapshot-reader-е.
' Все допустимые точности 0..10 проверяются без пересчета и без округления Results.
Private Sub TestVisualization(ByRef stats As TProfileStats, ByVal source As Object, ByVal system As Object, _
        ByVal loads As Object, ByVal section As CSectionModel, ByVal provider As CMaterialModelProvider, _
        ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem)
    Dim profile As Long, stateName As Variant, quantity As Variant, precision As Long, prefix As String
    Dim batch As CBatchSectionCalculator, writer As CNDMResultsWriter, reader As CSectionPlotDataReader
    Dim state As CSectionStateResult, solves As Long, before As Variant, firstValue As Double
    For profile = 1 To 4
        ConfigureProfiles source
        EnableChecks source, profile, "Yes", "Yes", "Yes", "No"
        SetLoad loads, profile, 200000#, 0#, 0#, "LambdaN"
        Set batch = ExecuteFixture(section, provider, settings, units)
        Set writer = New CNDMResultsWriter: writer.WriteResults ThisWorkbook, section, provider, batch, units
        before = ThisWorkbook.Names.Item("rngNDMElementResults").RefersToRange.CurrentRegion.Value2
        solves = SectionEquilibriumSolveCount()
        SetTableValue system, "Plot.LoadCase", "PROFILE_CONFIG", 2
        settings.LoadFromWorkbook ThisWorkbook
        For Each stateName In Array("StrengthState", "CapacityState", "PreCrackState", "PostCrackState", "CrackedState")
            SetProfileValue source, profile, "Visualization.State", stateName
            Set state = batch.ResultAt(1).StateRepository.FindState(SectionStateTypeFromText(CStr(stateName)))
            For Each quantity In Array("Stress", "Strain")
                SetProfileValue source, profile, "Visualization.Quantity", quantity
                prefix = "profileConfig.PR" & CStr(profile) & ".Visualization." & CStr(stateName) & "." & CStr(quantity)
                Set reader = New CSectionPlotDataReader: reader.LoadFromWorkbook ThisWorkbook, settings
                Check stats, prefix & ".state", reader.StateType = CStr(stateName)
                Check stats, prefix & ".quantity", reader.VisualizationQuantity = CStr(quantity)
                Check stats, prefix & ".elements", reader.Count = section.ConcreteCount + section.RebarCount
                Check stats, prefix & ".exists", Not state Is Nothing
                If Not state Is Nothing Then
                    CheckClose stats, prefix & ".plane", reader.Epsilon0, state.Epsilon0, 0.000000000001
                End If
                firstValue = reader.ResultValue(1)
                For precision = 0 To 10
                    SetProfileValue source, profile, "Visualization." & CStr(quantity) & "Precision", precision
                    Set reader = New CSectionPlotDataReader: reader.LoadFromWorkbook ThisWorkbook, settings
                    Check stats, prefix & ".precision" & CStr(precision), reader.ResultPrecision = precision
                    CheckClose stats, prefix & ".valueUnrounded" & CStr(precision), reader.ResultValue(1), firstValue, 0#
                Next precision
                stats.Cases = stats.Cases + 1
            Next quantity
        Next stateName
        Check stats, "profileConfig.PR" & CStr(profile) & ".Visualization.noSolve", SectionEquilibriumSolveCount() = solves
        Check stats, "profileConfig.PR" & CStr(profile) & ".Visualization.snapshotUnchanged", _
            TablesEqual(before, ThisWorkbook.Names.Item("rngNDMElementResults").RefersToRange.CurrentRegion.Value2)
    Next profile
End Sub

' Ошибка/пустота обязательного значения должна назвать профиль, key и ячейку.
' Дробная точность не округляется; переполнение не выдается за технический сбой VBA.
Private Sub TestInvalidProfileCells(ByRef stats As TProfileStats, ByVal source As Object)
    Dim profile As Long, key As Variant, value As Variant, code As Long, reason As String, prefix As String
    Dim cell As Object, caseIndex As Long
    For profile = 1 To 4
        For Each key In Array("Visualization.StressPrecision", "Visualization.StrainPrecision")
            caseIndex = 0
            For Each value In Array("", "-1", "11", "0.25", "9.9", "2147483648", "1e100", "abc", CVErr(xlErrDiv0))
                ConfigureProfiles source
                Set cell = ProfileCell(source, profile, CStr(key)): cell.Value2 = value
                code = CatalogError(reason)
                prefix = "profileConfig.PR" & CStr(profile) & "." & CStr(key) & ".invalid" & CStr(caseIndex)
                Check stats, prefix & ".rejected", code <> 0
                Check stats, prefix & ".controlled", code <> 6 And code <> 13
                Check stats, prefix & ".key", InStr(1, reason, CStr(key), vbBinaryCompare) > 0
                Check stats, prefix & ".address", InStr(1, reason, cell.Worksheet.Name & "!" & cell.Address(False, False), vbBinaryCompare) > 0
                Check stats, prefix & ".profile", InStr(1, reason, "PR" & CStr(profile), vbBinaryCompare) > 0
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
            ConfigureProfiles source
            Set cell = ProfileCell(source, profile, CStr(key)): cell.Value2 = "NOT_A_VALID_CHOICE"
            code = CatalogError(reason)
            prefix = "profileConfig.PR" & CStr(profile) & "." & CStr(key) & ".invalidChoice"
            Check stats, prefix & ".rejected", code <> 0
            Check stats, prefix & ".key", InStr(1, reason, CStr(key), vbBinaryCompare) > 0
            Check stats, prefix & ".address", InStr(1, reason, cell.Worksheet.Name & "!" & cell.Address(False, False), vbBinaryCompare) > 0
            stats.Cases = stats.Cases + 1
        Next key
    Next profile
End Sub

' Сопоставляет фактический State с независимо выбранной моделью и равновесием.
' Из арматурного вклада вычитается замещенный бетон, как в утвержденном ядре.
Private Sub CheckState(ByRef stats As TProfileStats, ByVal prefix As String, ByVal state As CSectionStateResult, _
        ByVal section As CSectionModel, ByVal provider As CMaterialModelProvider, ByVal expected As CMaterialModelSpec)
    Check stats, prefix & ".exists", Not state Is Nothing
    If state Is Nothing Then Exit Sub
    Check stats, prefix & ".spec", state.MaterialSpec.SpecKey = expected.SpecKey
    Check stats, prefix & ".converged", state.Converged
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
    CheckClose stats, prefix & ".expectedMaterialN", n, state.TargetN, 0.01
    CheckClose stats, prefix & ".expectedMaterialMx", mx, state.TargetMx, 1#
    CheckClose stats, prefix & ".expectedMaterialMy", my, state.TargetMy, 1#
End Sub

' Проверяет готовые subtree-комментарии в реальных detailed/batch writers.
' Неуспех не может остаться без причины; writer не формирует новый текст.
Private Sub CheckComments(ByRef stats As TProfileStats, ByVal prefix As String, _
        ByVal batch As CBatchSectionCalculator, ByVal units As CUnitSystem)
    Dim writer As CBatchResultWriter, result As CCombinationResult, meta As CResultMeta
    Dim names As Variant, metas As Variant, i As Long, policy As CResultStatusPolicy, solves As Long
    Set result = batch.ResultAt(1): Set writer = New CBatchResultWriter
    solves = batch.SolverCallCount: writer.WriteSummary ThisWorkbook, batch, units
    Check stats, prefix & ".writerNoSolve", batch.SolverCallCount = solves
    Check stats, prefix & ".batchComment", CStr(ThisWorkbook.Names.Item("rngBatchSummary").RefersToRange.Offset(12, 2).Value2) = result.OverallMeta.ResultComment
    names = Array("rngStrengthSummaryAnchor", "rngCrackSummaryAnchor", "rngStabilitySummaryAnchor")
    metas = Array(result.StrengthMeta, result.CrackSummaryMeta, result.StabilityMeta)
    Set policy = New CResultStatusPolicy
    For i = 0 To UBound(names)
        Set meta = metas(i)
        Check stats, prefix & ".comment" & CStr(i), CStr(ThisWorkbook.Names.Item(CStr(names(i))).RefersToRange.Offset(0, 1).Value2) = meta.ResultComment
        If policy.IsFailedOverall(policy.ExternalStatus(meta)) Then Check stats, prefix & ".failureReason" & CStr(i), Len(Trim$(meta.ResultComment)) > 0
    Next i
    LogLine stats, "PROFILE_CONFIG_RESULT: " & prefix & "; status=" & result.Status & "; comment=" & result.OverallMeta.ResultComment
End Sub

' Создает небольшую физическую модель; параметры материалов заданы явно,
' чтобы результат не зависел от оставленных соседней suite значений Config.
Private Sub BuildFixture(ByRef section As CSectionModel, ByRef provider As CMaterialModelProvider, ByVal extensionEnabled As Boolean)
    Dim geometry As CGeometryRoundedRectangle, mesh As CFiberMeshBuilder, rebars As CRebarLayout
    Dim concrete As CConcreteMaterialParameters, steel As CSteelMaterialParameters
    Set geometry = New CGeometryRoundedRectangle: geometry.Initialize 300#, 200#, 0#, 0#, 0#, 0#
    Set mesh = New CFiberMeshBuilder: mesh.BuildMesh geometry, 30#, 20#, 1
    Set rebars = New CRebarLayout
    rebars.AddBar "A", -90#, -60#, 20#, 0#, "A400", "", geometry
    rebars.AddBar "B", 90#, -60#, 20#, 0#, "A400", "", geometry
    rebars.AddBar "C", -90#, 60#, 20#, 0#, "A400", "", geometry
    rebars.AddBar "D", 90#, 60#, 20#, 0#, "A400", "", geometry
    Set section = BuildGeneratedSectionModel(mesh, rebars, "Audit03ProfileConfig")
    Set concrete = New CConcreteMaterialParameters: concrete.Initialize 15.5, 1.1, 22#, 1.8, 32500#, 32500#, rbMc2:=14.6
    Set steel = New CSteelMaterialParameters: steel.Initialize 350#, 350#, 390#, 390#, 200000#, 200000#
    Set provider = New CMaterialModelProvider
    provider.InitializeFromParameters concrete, steel, diagramExtensionEnabled:=extensionEnabled
End Sub

' Пропускает реальную таблицу LC через reader и новый batch-контекст.
' Профильные диаграммы выбирает production, не тестовый fake solver.
Private Function ExecuteFixture(ByVal section As CSectionModel, ByVal provider As CMaterialModelProvider, _
        ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem) As CBatchSectionCalculator
    Dim catalog As CCalculationProfileCatalog, batch As CBatchSectionCalculator, reader As CLoadCombinationReader
    Set catalog = New CCalculationProfileCatalog: catalog.LoadFromWorkbook ThisWorkbook
    Set batch = New CBatchSectionCalculator: batch.Initialize section, provider
    Set batch.ProfileCatalog = catalog: batch.ApplySettings settings, units
    batch.SetSP35Table721 ThisWorkbook.Names.Item("rngSP35Table721").RefersToRange.Value2
    Set reader = New CLoadCombinationReader: reader.LoadFromWorkbook ThisWorkbook, batch, units
    batch.Execute
    Set ExecuteFixture = batch
End Function

' Задает независимую базу трех моделей и полей визуализации для всех PR.
' Отключенные флаги включаются конкретным тестом, не глобальным default.
Private Sub ConfigureProfiles(ByVal source As Object)
    Dim profile As Long, role As Variant, spec As CMaterialModelSpec
    For profile = 1 To 4
        SetProfileValue source, profile, "Profile.DisplayName", "Профиль " & CStr(profile)
        SetProfileValue source, profile, "Profile.Description", "Проверка параметров профиля."
        EnableChecks source, profile, "No", "No", "No", "No"
        SetProfileValue source, profile, "MaterialModel.Stability.ValueSet", "ULS(I)"
        For Each role In Array("Strength", "CrackInitiation", "CrackedState")
            Set spec = ExpectedSpec(CStr(role), "", "")
            SetProfileValue source, profile, "MaterialModel." & CStr(role) & ".ValueSet", spec.ValueSetText
            SetProfileValue source, profile, "MaterialModel." & CStr(role) & ".ConcreteDiagram", spec.ConcreteDiagramText
            SetProfileValue source, profile, "MaterialModel." & CStr(role) & ".ConcreteTension", spec.ConcreteTensionText
            SetProfileValue source, profile, "MaterialModel." & CStr(role) & ".SteelDiagram", spec.SteelDiagramText
        Next role
        SetProfileValue source, profile, "Visualization.State", "StrengthState"
        SetProfileValue source, profile, "Visualization.Quantity", "Stress"
        SetProfileValue source, profile, "Visualization.StressPrecision", "1"
        SetProfileValue source, profile, "Visualization.StrainPrecision", "6"
    Next profile
End Sub

' Выбирает запросы существующих calculators без смешения с material spec.
Private Sub EnableChecks(ByVal source As Object, ByVal profile As Long, ByVal direct As String, _
        ByVal capacity As String, ByVal crack As String, ByVal stability As String)
    SetProfileValue source, profile, "Calculation.Strength.DirectState", direct
    SetProfileValue source, profile, "Calculation.Strength.Capacity", capacity
    SetProfileValue source, profile, "Calculation.Crack.Width", crack
    SetProfileValue source, profile, "Calculation.Stability.Enabled", stability
End Sub

' Формирует один физический LC; остальные source slots остаются пустыми.
Private Sub SetLoad(ByVal source As Object, ByVal profile As Long, ByVal n As Double, _
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

' Возвращает независимые ожидаемые тексты модели, не читая фактический State.
Private Function ExpectedSpec(ByVal role As String, ByVal field As String, ByVal value As String) As CMaterialModelSpec
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
    Set ExpectedSpec = spec
End Function

' Перечисляет допустимые активные варианты роли; ограничения tensile mode
' принадлежат утвержденной постановке Pre/Post, а не новой тестовой формуле.
Private Function MaterialVariants(ByVal role As String, ByVal field As String) As Variant
    Select Case field
        Case "ValueSet": MaterialVariants = Array("ULS(I)", "SLS(II)")
        Case "ConcreteDiagram", "SteelDiagram": MaterialVariants = Array("TwoLine", "ThreeLine")
        Case "ConcreteTension"
            Select Case role
                Case "Strength": MaterialVariants = Array("Ignore", "UseDiagram")
                Case "CrackInitiation": MaterialVariants = Array("UseDiagram")
                Case "CrackedState": MaterialVariants = Array("Ignore")
            End Select
    End Select
End Function

' Находит фактическую профильную ячейку по key; ошибка fixture не скрывается.
Private Function ProfileCell(ByVal source As Object, ByVal profile As Long, ByVal key As String) As Object
    Dim row As Long
    For row = 1 To source.Rows.Count
        If CStr(source.Cells(row, 2).Value2) = key Then Set ProfileCell = source.Cells(row, profile + 2): Exit Function
    Next row
    Err.Raise vbObjectError + 4501, "ProfileCell", "В тестовой таблице нет " & key
End Function

' Записывает Variant без превращения CVErr или пустоты в строковый default.
Private Sub SetProfileValue(ByVal source As Object, ByVal profile As Long, ByVal key As String, ByVal value As Variant)
    ProfileCell(source, profile, key).Value2 = value
End Sub

' Находит scalar-настройку по фактическому ключу ее существующей таблицы.
Private Sub SetTableValue(ByVal source As Object, ByVal key As String, ByVal value As String, ByVal column As Long)
    Dim row As Long
    For row = 1 To source.Rows.Count
        If CStr(source.Cells(row, 1).Value2) = key Then source.Cells(row, column).Value2 = value: Exit Sub
    Next row
    Err.Raise vbObjectError + 4501, "SetTableValue", "В тестовой таблице нет " & key
End Sub

' Сохраняет фактическую контролируемую ошибку каталога, не назначая статус по тексту.
Private Function CatalogError(ByRef reason As String) As Long
    Dim catalog As CCalculationProfileCatalog
    reason = vbNullString
    On Error GoTo Failed
    Set catalog = New CCalculationProfileCatalog: catalog.LoadFromWorkbook ThisWorkbook
    Exit Function
Failed:
    CatalogError = Err.Number: reason = Err.Description
End Function

' Сравнивает сериализованные численные таблицы без нормализации или допусков.
Private Function TablesEqual(ByRef before As Variant, ByRef after As Variant) As Boolean
    Dim row As Long, column As Long
    If UBound(before, 1) <> UBound(after, 1) Or UBound(before, 2) <> UBound(after, 2) Then Exit Function
    For row = 1 To UBound(before, 1)
        For column = 1 To UBound(before, 2)
            If CStr(before(row, column)) <> CStr(after(row, column)) Then Exit Function
        Next column
    Next row
    TablesEqual = True
End Function

' Сохраняет стабильный assertion ID и отдельный счетчик ошибок.
Private Sub Check(ByRef stats As TProfileStats, ByVal id As String, ByVal condition As Boolean)
    If condition Then stats.Passed = stats.Passed + 1: LogLine stats, "OK: " & id Else stats.Failed = stats.Failed + 1: LogLine stats, "FAIL: " & id
End Sub

' Сравнивает с прежним компонентным допуском; precision не округляет actual.
Private Sub CheckClose(ByRef stats As TProfileStats, ByVal id As String, ByVal actual As Double, ByVal expected As Double, ByVal tolerance As Double)
    Check stats, id & "; actual=" & Trim$(Str$(actual)) & "; expected=" & Trim$(Str$(expected)), Abs(actual - expected) <= tolerance
End Sub

' Записывает диагностический checkpoint для внешнего watchdog без принятия PASS.
Private Sub LogLine(ByRef stats As TProfileStats, ByVal value As String)
    stats.Report = stats.Report & value & vbCrLf
    If Left$(value, 5) = "RUN: " Then
        Dim stream As Object
        Set stream = CreateObject("Scripting.FileSystemObject").OpenTextFile(ThisWorkbook.Path & "\RC_NDM_ui_test_progress.txt", 8, True, -1)
        stream.WriteLine Format$(Now, "yyyy-mm-dd hh:nn:ss") & " " & value: stream.Close
    End If
End Sub
