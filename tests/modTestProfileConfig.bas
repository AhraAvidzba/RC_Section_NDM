Attribute VB_Name = "modTestProfileConfig"
Option Explicit

' ==========================================================================
' ДЛЯ ТЕСТОВ: ПРОФИЛИ, АДРЕСНЫЙ ВВОД И НАГРУЗКИ УСТОЙЧИВОСТИ
' ==========================================================================
' Проверяет реальные ячейки PR1-PR4, передачу material spec в State/Search
' и выбор сохраненных результатов для схемы. Ожидаемая спецификация задается
' независимо от прочитанного профиля; равновесие пересчитывается по элементам.
' Дополнительные нагрузки проходят тот же workbook-reader, что кнопка расчета;
' ошибка активной строки должна оставаться в собственном результате устойчивости.
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
                    ' Ошибки геометрических значений хранятся до обращения
                    ' активного потребителя. Проверяем его чтение по тому же
                    ' ключу и сохраненный адрес перемещенной ячейки.
                    cell.Value2 = CVErr(xlErrDiv0): code = SettingsLoadError(reason, key)
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
            If index = 8 Then TestRelocatedRectSetSelectors stats, target, position
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

' Проверяет независимые выборы и ошибки формул обеих физических граней.
' Адрес ввода берется из перемещенной таблицы; другая грань не подменяет
' ошибочное значение, а исходные настройки восстанавливаются после пробы.
Private Sub TestRelocatedRectSetSelectors(ByRef stats As TProfileStats, ByVal source As Object, ByVal position As Long)
    Dim face As Variant, column As Variant, row As Long, first As Object, second As Object
    Dim savedFirst As Variant, savedSecond As Variant, code As Long, reason As String, prefix As String
    Dim settings As CSystemSettingsReader, choices As Variant, tail As String, key As String, value As String, side As Long, cell As Object
    row = CaptionRow(source, "H1", 1) + 14
    For Each face In Array("H1", "B1", "H2", "B2")
        For Each column In Array(3, 4, 6, 7)
            Set first = source.Cells(row, CLng(column)): Set second = source.Cells(row + 1, CLng(column))
            savedFirst = first.Value2: savedSecond = second.Value2
            prefix = "relocated.rectset.p" & CStr(position) & "." & CStr(face) & ".c" & CStr(column)
            Select Case CLng(column)
                Case 3: tail = "loc_2row": choices = Array("Stacked", "SideBySide")
                Case 4: tail = "bind_2row": choices = Array("EachBar", "EverySecondBar")
                Case 6: tail = "loc_3row": choices = Array("Stacked", "SideBySide")
                Case 7: tail = "bind_3row": choices = Array("EachBar", "EverySecondBar")
            End Select
            first.Value2 = choices(0): second.Value2 = choices(1)
            Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
            key = "RectSet." & CStr(face) & "." & tail
            Check stats, prefix & ".firstIndependent", settings.GetRequiredChoice(key & "_1", choices) = choices(0)
            Check stats, prefix & ".secondIndependent", settings.GetRequiredChoice(key & "_2", choices) = choices(1)
            For side = 1 To 2
                Set cell = source.Cells(row + side - 1, CLng(column))
                cell.Value2 = CVErr(xlErrDiv0)
                Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
                On Error Resume Next
                value = settings.GetRequiredChoice(key & "_" & CStr(side), choices)
                code = Err.Number: reason = Err.Description: Err.Clear
                On Error GoTo 0
                Check stats, prefix & ".formulaRejected." & CStr(side), code <> 0
                Check stats, prefix & ".formulaAddress." & CStr(side), InStr(1, reason, "ячейка " & cell.Address(False, False), vbBinaryCompare) > 0
                Check stats, prefix & ".formulaSheet." & CStr(side), InStr(1, reason, source.Worksheet.Name, vbBinaryCompare) > 0
                first.Value2 = choices(0): second.Value2 = choices(1)
            Next side
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

' Сохраняет фактическую ошибку загрузки или активного чтения указанного поля
' после перемещения таблицы. Геометрический CVErr не мешает другой форме,
' но при обращении к нему обязан выдавать исходный фактический адрес.
Private Function SettingsLoadError(ByRef reason As String, Optional ByVal readKey As String = vbNullString) As Long
    On Error GoTo Failed
    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Dim value As String
    If Len(readKey) > 0 Then value = settings.GetRawString(readKey)
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
    Dim value As Variant, code As Long, reason As String, cell As Object
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
        For Each value In Array(vbNullString, "  Пояснение профиля  ", "Произвольное пояснение: НДС и трещины")
            SetProfileValue source, profile, "Profile.Description", value
            Set catalog = New CCalculationProfileCatalog: catalog.LoadFromWorkbook ThisWorkbook
            Check stats, prefix & ".Profile.Description.optional", catalog.ProfileById(caption).Description = Trim$(CStr(value))
            Check stats, prefix & ".Profile.Description.nameStable", catalog.ProfileById(caption).ProfileId = "PR" & CStr(profile)
            stats.Cases = stats.Cases + 1
        Next value
        Set cell = ProfileCell(source, profile, "Profile.Description"): cell.Value2 = CVErr(xlErrValue)
        code = CatalogError(reason)
        Check stats, prefix & ".Profile.Description.error.rejected", code <> 0
        Check stats, prefix & ".Profile.Description.error.controlled", code <> 13 And code <> 6
        Check stats, prefix & ".Profile.Description.error.key", InStr(1, reason, "Profile.Description", vbBinaryCompare) > 0
        Check stats, prefix & ".Profile.Description.error.address", InStr(1, reason, cell.Worksheet.Name & "!" & cell.Address(False, False), vbBinaryCompare) > 0
        SetProfileValue source, profile, "Profile.Description", description
        code = CatalogError(reason): Check stats, prefix & ".Profile.Description.recovery", code = 0
        stats.Cases = stats.Cases + 1
    Next profile
End Sub

' Выбирает все пять named states и обе величины в реальном snapshot-reader-е.
' Сверяет значения по плоскости/материалу, а точности 0..10 - по настоящему
' тексту легенды Chart; схема не решает НДС и не округляет данные Results.
Private Sub TestVisualization(ByRef stats As TProfileStats, ByVal source As Object, ByVal system As Object, _
        ByVal loads As Object, ByVal section As CSectionModel, ByVal provider As CMaterialModelProvider, _
        ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem)
    Dim profile As Long, stateName As Variant, quantity As Variant, precision As Long, prefix As String
    Dim batch As CBatchSectionCalculator, writer As CNDMResultsWriter, reader As CSectionPlotDataReader
    Dim state As CSectionStateResult, solves As Long, before As Variant, firstValue As Double
    SetTableValue system, "Plot.LegendMode", "Common", 2
    SetTableValue system, "Plot.LegendEnabled", "Yes", 2
    SetTableValue system, "Plot.ResultLabelsEnabled", "Yes", 2
    SetTableValue system, "Plot.ResultLabelSpacing", "50", 2
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
                    CheckVisualizationValues stats, prefix, reader, state, section, provider, units
                End If
                firstValue = reader.ResultValue(1)
                For precision = 0 To 10
                    SetProfileValue source, profile, "Visualization." & CStr(quantity) & "Precision", precision
                    Set reader = New CSectionPlotDataReader: reader.LoadFromWorkbook ThisWorkbook, settings
                    Check stats, prefix & ".precision" & CStr(precision), reader.ResultPrecision = precision
                    CheckClose stats, prefix & ".valueUnrounded" & CStr(precision), reader.ResultValue(1), firstValue, 0#
                    CheckVisualizationLegend stats, prefix & ".precision" & CStr(precision), reader, settings
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

' ======================= ДЛЯ ТЕСТОВ: DURATION LOADS =======================

' Проверяет каждую из 90 ячеек дополнительных нагрузок на реальном расчете.
' Отдельно проверяет ошибочный ввод, изоляцию LC, единицы/знаки, перенос точки
' приложения и адреса после перемещения имени. Формулы/Names восстанавливаются.
Public Function RunAudit03DurationConfigTests(Optional ByRef passed As Long = 0, _
        Optional ByRef failed As Long = 0) As String
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
    ConfigureProfiles profiles: EnableChecks profiles, 1, "No", "No", "No", "Yes"
    SetTableValue system, "Stability.Code", "SP63", 2
    SetTableValue system, "Stability.ElementLength", "1000", 2
    SetTableValue system, "Stability.Mu1", "1", 2: SetTableValue system, "Stability.Mu2", "1", 2
    SetTableValue system, "Stability.SystemType", "Determinate", 2
    SetTableValue system, "Stability.ZeroMomentEccentricitySign1", "1", 2
    SetTableValue system, "Stability.ZeroMomentEccentricitySign2", "1", 2
    SetTableValue system, "Stability.AccidentalEccentricityMode", "User", 2
    SetTableValue system, "Stability.AccidentalEccentricityPlanes", "BothPlanes", 2
    SetTableValue system, "Stability.AccidentalEccentricityUser1", "5", 2
    SetTableValue system, "Stability.AccidentalEccentricityUser2", "5", 2
    SetTableValue system, "Stability.PhiLMode", "Auto", 2
    SetTableValue system, "Calculation.ZeroMomentPerDepth", "0", 2
    SetTableValue system, "Load.ReferenceOffsetX", "0", 2: SetTableValue system, "Load.ReferenceOffsetY", "0", 2
    SetTableValue system, "Solver.ToleranceN", "0.01", 2
    SetTableValue system, "Solver.ToleranceMx", "1", 2: SetTableValue system, "Solver.ToleranceMy", "1", 2
    ConfigureDurationUnits "N", "N*mm", "Tension", "+Y tension", "+X tension"
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    BuildDurationFixture section, provider, settings.GetBoolean("General.DiagramExtension", False)
    TestDurationCells stats, profiles, loads, duration, section, provider, settings, units
    TestDurationUnits stats, loads, duration, section, provider, settings, units
    ConfigureDurationUnits "N", "N*mm", "Tension", "+Y tension", "+X tension"
    settings.LoadFromWorkbook ThisWorkbook: units.LoadFromSettings settings
    TestDurationInvalidRows stats, profiles, loads, duration, section, provider, settings, units
    TestDurationOverflow stats, loads, duration, section, provider, settings, units
    TestDurationInactive stats, profiles, loads, duration, section, provider, settings, units
    TestDurationTableLinks stats, loads, duration, section, provider, settings, units
    TestDurationRangeContract stats, loads, duration, durationRef, section, provider, settings, units, sheet
    GoTo Restore
FailedRun:
    stats.Failed = stats.Failed + 1
    LogLine stats, "FAIL: durationConfig.runtime; " & CStr(Err.Number) & "; " & Err.Description
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
    If restoreNumber <> 0 Then stats.Failed = stats.Failed + 1: LogLine stats, "FAIL: durationConfig.restore; " & CStr(restoreNumber) & "; " & restoreReason
    LogLine stats, "TOTAL_AUDIT03_DURATION_CONFIG: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed) & "; cases=" & CStr(stats.Cases)
    passed = stats.Passed: failed = stats.Failed: RunAudit03DurationConfigTests = stats.Report
End Function

' Для симметричного высокого прямоугольника главные оси совпадают с X/Y.
' Поэтому знаки и компоненты длительных моментов проверяются без production
' преобразователя осей; поставщик материалов остается общим с другими тестами.
Private Sub BuildDurationFixture(ByRef section As CSectionModel, ByRef provider As CMaterialModelProvider, ByVal extensionEnabled As Boolean)
    Dim geometry As CGeometryRoundedRectangle, mesh As CFiberMeshBuilder, rebars As CRebarLayout
    BuildFixture section, provider, extensionEnabled
    Set geometry = New CGeometryRoundedRectangle: geometry.Initialize 200#, 300#, 0#, 0#, 0#, 0#
    Set mesh = New CFiberMeshBuilder: mesh.BuildMesh geometry, 20#, 30#, 1
    Set rebars = New CRebarLayout
    rebars.AddBar "A", -60#, -90#, 20#, 0#, "A400", "", geometry
    rebars.AddBar "B", 60#, -90#, 20#, 0#, "A400", "", geometry
    rebars.AddBar "C", -60#, 90#, 20#, 0#, "A400", "", geometry
    rebars.AddBar "D", 60#, 90#, 20#, 0#, "A400", "", geometry
    Set section = BuildGeneratedSectionModel(mesh, rebars, "Audit03DurationConfig")
End Sub

' Перебирает все source slots и каждую компоненту отдельно. Остальные строки
' имеют другие ID, чтобы тест выявлял неверное сопоставление по позиции.
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
                PrepareDurationLoads loads, duration, slot
                data = duration.Value2: data(slot + 1, column) = testValue: duration.Value2 = data
                values = Array(-25000#, 400000#, 600000#)
                If Len(CStr(testValue)) = 0 Then values(column - 2) = 0# Else values(column - 2) = CDbl(testValue)
                prefix = "durationConfig.Slot" & CStr(slot) & "." & DurationColumnKey(column) & ".value=" & CStr(testValue)
                LogLine stats, "RUN: " & prefix
                code = ExecuteDuration(section, provider, settings, units, batch, reason)
                Check stats, prefix & ".reader", code = 0
                If code = 0 Then
                    CheckDurationResult stats, prefix, batch, units, CDbl(values(0)), CDbl(values(1)), CDbl(values(2))
                    CheckComments stats, prefix, batch, units
                Else
                    LogLine stats, "DURATION_ERROR: " & prefix & "; " & CStr(code) & "; " & reason
                End If
                stats.Cases = stats.Cases + 1
            Next testValue
        Next column
    Next slot
    ' Перенос длительной нагрузки должен использовать ту же точку, что полный LC.
    PrepareDurationLoads loads, duration, 1
    code = ExecuteDuration(section, provider, settings, units, batch, reason, 40#, -25#)
    Check stats, "durationConfig.offset.reader", code = 0
    If code = 0 Then CheckDurationResult stats, "durationConfig.offset", batch, units, -25000#, 1025000#, -400000#, 2250000#, 0#
    stats.Cases = stats.Cases + 1
End Sub

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
                    ConfigureDurationUnits forceUnit, momentUnit, nSign, mxSign, mySign
                    settings.LoadFromWorkbook ThisWorkbook: units.LoadFromSettings settings
                    PrepareDurationLoads loads, duration, 1
                    data = loads.Value2
                    data(2, 2) = -50000# / forceScale * signN
                    data(2, 3) = 1000000# / momentScale * signMx: data(2, 4) = 2000000# / momentScale * signMy
                    loads.Value2 = data
                    data = duration.Value2
                    data(2, 2) = -25000# / forceScale * signN
                    data(2, 3) = 400000# / momentScale * signMx: data(2, 4) = 600000# / momentScale * signMy
                    duration.Value2 = data
                    prefix = "durationConfig.units." & CStr(mode) & "." & CStr(signN) & "." & CStr(signMx) & "." & CStr(signMy)
                    LogLine stats, "RUN: " & prefix
                    code = ExecuteDuration(section, provider, settings, units, batch, reason)
                    Check stats, prefix & ".reader", code = 0
                    If code = 0 Then CheckDurationResult stats, prefix, batch, units, -25000#, 400000#, 600000#: CheckComments stats, prefix, batch, units
                    stats.Cases = stats.Cases + 1
                Next signMy
            Next signMx
        Next signN
    Next mode
End Sub

' Ошибка каждой активной ячейки относится только к устойчивости данного LC.
' DirectState и следующая корректная строка должны быть рассчитаны; невыполненная
' устойчивость не получает Calculated=True или NumFail.
Private Sub TestDurationInvalidRows(ByRef stats As TProfileStats, ByVal profiles As Object, ByVal loads As Object, _
        ByVal duration As Object, ByVal section As CSectionModel, ByVal provider As CMaterialModelProvider, _
        ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem)
    Dim slot As Long, column As Long, testValue As Variant, data As Variant, batch As CBatchSectionCalculator
    Dim code As Long, reason As String, prefix As String, cell As Object, errorCase As Long
    EnableChecks profiles, 1, "Yes", "No", "No", "Yes"
    For slot = 1 To duration.Rows.Count - 1
        For column = 2 To 4
            For errorCase = 1 To 2
                PrepareDurationLoads loads, duration, slot
                If errorCase = 1 Then testValue = "BAD_DURATION" Else testValue = CVErr(2015)
                data = duration.Value2: data(slot + 1, column) = testValue: duration.Value2 = data
                AddDurationRecoveryLoad loads
                Set cell = duration.Cells(slot + 1, column)
                prefix = "durationConfig.Slot" & CStr(slot) & "." & DurationColumnKey(column) & ".error" & CStr(errorCase)
                LogLine stats, "RUN: " & prefix
                code = ExecuteDuration(section, provider, settings, units, batch, reason)
                Check stats, prefix & ".readerNoAbort", code = 0
                If code = 0 Then
                    CheckDurationInputError stats, prefix, batch, units, cell.Worksheet.Name & "!" & cell.Address(False, False)
                    Check stats, prefix & ".directContinues", batch.ResultAt(1).DirectStateMeta.InternalStatus = rsSuccess
                    Check stats, prefix & ".recovery", batch.Count = 2 And batch.ResultAt(2).StabilityMeta.InternalStatus = rsSuccess
                    Check stats, prefix & ".recoveryDirect", batch.ResultAt(2).DirectStateMeta.InternalStatus = rsSuccess
                    Check stats, prefix & ".recoveryBatchComment", _
                        CStr(ThisWorkbook.Names.Item("rngBatchSummary").RefersToRange.Offset(13, 2).Value2) = batch.ResultAt(2).OverallMeta.ResultComment
                    Check stats, prefix & ".recoveryOwnComment", _
                        CStr(ThisWorkbook.Names.Item("rngStabilitySummaryAnchor").RefersToRange.Offset(1, 1).Value2) = batch.ResultAt(2).StabilityMeta.ResultComment
                    Check stats, prefix & ".noReasonLeak", InStr(1, batch.ResultAt(2).StabilityMeta.ResultComment, _
                        cell.Worksheet.Name & "!" & cell.Address(False, False), vbBinaryCompare) = 0
                Else
                    LogLine stats, "DURATION_ERROR: " & prefix & "; " & CStr(code) & "; " & reason
                End If
                stats.Cases = stats.Cases + 1
            Next errorCase
        Next column
    Next slot
    EnableChecks profiles, 1, "No", "No", "No", "Yes"
End Sub

' Число может помещаться в Double до перевода, но переполнить внутренние Н
' или Н*мм. Это адресная ошибка входа данной строки, а не NumFail solver-а.
Private Sub TestDurationOverflow(ByRef stats As TProfileStats, ByVal loads As Object, ByVal duration As Object, _
        ByVal section As CSectionModel, ByVal provider As CMaterialModelProvider, _
        ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem)
    Dim column As Long, data As Variant, batch As CBatchSectionCalculator, code As Long, reason As String, prefix As String, cell As Object
    ConfigureDurationUnits "tf", "tf*m", "Tension", "+Y tension", "+X tension"
    settings.LoadFromWorkbook ThisWorkbook: units.LoadFromSettings settings
    For column = 2 To 4
        PrepareDurationLoads loads, duration, 1
        data = loads.Value2: data(2, 2) = -50000# / 9806.65
        data(2, 3) = 1000000# / 9806650#: data(2, 4) = 2000000# / 9806650#: loads.Value2 = data
        data = duration.Value2: data(2, 2) = -25000# / 9806.65
        data(2, 3) = 400000# / 9806650#: data(2, 4) = 600000# / 9806650#
        data(2, column) = 1E+305: duration.Value2 = data
        prefix = "durationConfig.overflow." & DurationColumnKey(column): LogLine stats, "RUN: " & prefix
        code = ExecuteDuration(section, provider, settings, units, batch, reason)
        Check stats, prefix & ".readerNoAbort", code = 0
        Set cell = duration.Cells(2, column)
        If code = 0 Then CheckDurationInputError stats, prefix, batch, units, cell.Worksheet.Name & "!" & cell.Address(False, False)
        stats.Cases = stats.Cases + 1
    Next column
    ConfigureDurationUnits "N", "N*mm", "Tension", "+Y tension", "+X tension"
    settings.LoadFromWorkbook ThisWorkbook: units.LoadFromSettings settings
End Sub

' Проверяет case-insensitive ID и существующий приоритет первой строки.
' Ошибка формулы в самом ID является структурной, ее адрес не выдумывается.
Private Sub TestDurationTableLinks(ByRef stats As TProfileStats, ByVal loads As Object, ByVal duration As Object, _
        ByVal section As CSectionModel, ByVal provider As CMaterialModelProvider, _
        ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem)
    Dim mode As Long, data As Variant, batch As CBatchSectionCalculator, code As Long, reason As String, prefix As String, cell As Object
    For mode = 1 To 4
        PrepareDurationLoads loads, duration, 1
        data = duration.Value2
        If mode = 1 Then data(2, 1) = "duration_1"
        If mode = 2 Then data(3, 1) = "DURATION_1": data(3, 2) = "BAD_DUPLICATE"
        If mode = 3 Then data(3, 1) = "DURATION_1": data(2, 2) = "BAD_FIRST"
        If mode = 4 Then data(3, 1) = CVErr(2015)
        duration.Value2 = data
        prefix = "durationConfig.links." & CStr(mode): LogLine stats, "RUN: " & prefix
        code = ExecuteDuration(section, provider, settings, units, batch, reason)
        Check stats, prefix & ".readerNoAbort", code = 0
        If code = 0 Then
            If mode <= 2 Then
                CheckDurationResult stats, prefix, batch, units, -25000#, 400000#, 600000#
                CheckComments stats, prefix, batch, units
            Else
                If mode = 3 Then Set cell = duration.Cells(2, 2) Else Set cell = duration.Cells(3, 1)
                CheckDurationInputError stats, prefix, batch, units, cell.Worksheet.Name & "!" & cell.Address(False, False)
            End If
        End If
        stats.Cases = stats.Cases + 1
    Next mode
End Sub

' Неактивный профиль и неиспользуемый ID не должны зависеть от ошибочных
' чисел дополнительной строки. Проверяется также пустая строка без ID.
Private Sub TestDurationInactive(ByRef stats As TProfileStats, ByVal profiles As Object, ByVal loads As Object, _
        ByVal duration As Object, ByVal section As CSectionModel, ByVal provider As CMaterialModelProvider, _
        ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem)
    Dim mode As Long, data As Variant, batch As CBatchSectionCalculator, code As Long, reason As String, prefix As String
    For mode = 1 To 3
        PrepareDurationLoads loads, duration, 1
        EnableChecks profiles, 1, "Yes", "No", "No", "Yes"
        data = duration.Value2
        If mode = 1 Then EnableChecks profiles, 1, "Yes", "No", "No", "No"
        If mode = 2 Then data(3, 1) = "UNUSED_ID"
        If mode = 3 Then data(3, 1) = vbNullString
        If mode = 1 Then data(2, 2) = CVErr(2015) Else data(3, 2) = CVErr(2015)
        duration.Value2 = data
        prefix = "durationConfig.inactive." & CStr(mode): LogLine stats, "RUN: " & prefix
        code = ExecuteDuration(section, provider, settings, units, batch, reason)
        Check stats, prefix & ".readerNoAbort", code = 0
        If code = 0 Then
            Check stats, prefix & ".direct", batch.ResultAt(1).DirectStateMeta.InternalStatus = rsSuccess
            If mode = 1 Then Check stats, prefix & ".notRequested", batch.ResultAt(1).StabilityMeta.InternalStatus = rsNotRequested Else Check stats, prefix & ".stability", batch.ResultAt(1).StabilityMeta.InternalStatus = rsSuccess
            CheckComments stats, prefix, batch, units
        End If
        stats.Cases = stats.Cases + 1
    Next mode
    EnableChecks profiles, 1, "No", "No", "No", "Yes"
End Sub

' Непрочитанный активный named range не эквивалентен пустой допустимой строке.
' При переносе на другой лист ошибка обязана указывать новое место ввода.
Private Sub TestDurationRangeContract(ByRef stats As TProfileStats, ByVal loads As Object, ByVal duration As Object, _
        ByVal originalRef As String, ByVal section As CSectionModel, ByVal provider As CMaterialModelProvider, _
        ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem, ByRef sheet As Object)
    Dim code As Long, reason As String, batch As CBatchSectionCalculator, position As Long, column As Long
    Dim target As Object, data As Variant, prefix As String, cell As Object
    PrepareDurationLoads loads, duration, 1
    ThisWorkbook.Names.Item("rngStabilityDurationLoads").RefersTo = "=#REF!"
    code = ExecuteDuration(section, provider, settings, units, batch, reason)
    Check stats, "durationConfig.missing.readerNoAbort", code = 0
    If code = 0 Then CheckDurationInputError stats, "durationConfig.missing", batch, units, "rngStabilityDurationLoads"
    stats.Cases = stats.Cases + 1
    ThisWorkbook.Names.Item("rngStabilityDurationLoads").RefersTo = originalRef
    ThisWorkbook.Names.Item("rngStabilityDurationLoads").RefersTo = "='" & duration.Worksheet.Name & "'!" & duration.Resize(duration.Rows.Count, 3).Address
    code = ExecuteDuration(section, provider, settings, units, batch, reason)
    Check stats, "durationConfig.columns.readerNoAbort", code = 0
    If code = 0 Then CheckDurationInputError stats, "durationConfig.columns", batch, units, "rngStabilityDurationLoads"
    stats.Cases = stats.Cases + 1
    ThisWorkbook.Names.Item("rngStabilityDurationLoads").RefersTo = originalRef
    Set sheet = ThisWorkbook.Worksheets.Add: sheet.Name = "__AuditDurationInputs"
    For position = 1 To 2
        If position = 1 Then Set target = sheet.Cells(10, 5) Else Set target = sheet.Cells(800, 60)
        Set target = target.Resize(duration.Rows.Count, duration.Columns.Count)
        For column = 2 To 4
            PrepareDurationLoads loads, duration, 1
            data = duration.Value2: data(2, column) = CVErr(2015): target.Value2 = data
            ThisWorkbook.Names.Item("rngStabilityDurationLoads").RefersTo = "='" & sheet.Name & "'!" & target.Address
            Set cell = target.Cells(2, column)
            prefix = "durationConfig.relocated." & CStr(position) & "." & DurationColumnKey(column): LogLine stats, "RUN: " & prefix
            code = ExecuteDuration(section, provider, settings, units, batch, reason)
            Check stats, prefix & ".readerNoAbort", code = 0
            If code = 0 Then CheckDurationInputError stats, prefix, batch, units, sheet.Name & "!" & cell.Address(False, False)
            ThisWorkbook.Names.Item("rngStabilityDurationLoads").RefersTo = originalRef
            stats.Cases = stats.Cases + 1
        Next column
    Next position
End Sub

' Задает один основной LC и 30 независимых duration ID. Пустой ID сам по себе
' не вводит ошибочный LC, а сопоставление производится по ID, не по номеру строки.
Private Sub PrepareDurationLoads(ByVal loads As Object, ByVal duration As Object, ByVal activeSlot As Long)
    Dim data As Variant, row As Long
    SetLoad loads, 1, -50000#, 1000000#, 2000000#, "Auto"
    data = loads.Value2: data(2, 1) = "DURATION_" & CStr(activeSlot): loads.Value2 = data
    data = duration.Value2
    For row = 2 To UBound(data, 1)
        data(row, 1) = "DURATION_" & CStr(row - 1)
        data(row, 2) = -25000#: data(row, 3) = 400000#: data(row, 4) = 600000#
    Next row
    duration.Value2 = data
End Sub

' Добавляет следующую корректную строку, не совпадающую с поврежденным ID.
' Длительная часть для нее отсутствует, что по действующему контракту допустимо.
Private Sub AddDurationRecoveryLoad(ByVal loads As Object)
    Dim data As Variant
    data = loads.Value2
    data(3, 1) = "DURATION_RECOVERY": data(3, 2) = -50000#
    data(3, 3) = 1000000#: data(3, 4) = 2000000#: data(3, 5) = "PR1": data(3, 6) = "Auto"
    loads.Value2 = data
End Sub

' Меняет только таблицы единиц/знаков; все физические величины теста остаются
' одинаковыми. OUTPUT установлен явно, чтобы предыдущая suite не влияла на writer.
Private Sub ConfigureDurationUnits(ByVal force As String, ByVal moment As String, ByVal signN As String, ByVal signMx As String, ByVal signMy As String)
    Dim source As Object
    Set source = ThisWorkbook.Names.Item("rngUnitSettings").RefersToRange
    SetTableValue source, "Length", "mm", 2: SetTableValue source, "Force", force, 2
    SetTableValue source, "Moment", moment, 2: SetTableValue source, "Stress", "MPa", 2
    SetTableValue source, "Force", force, 4: SetTableValue source, "Moment", moment, 4
    Set source = ThisWorkbook.Names.Item("rngSignConventionSettings").RefersToRange
    SetTableValue source, "+N", signN, 2: SetTableValue source, "+Mx", signMx, 2: SetTableValue source, "+My", signMy, 2
End Sub

' Сохраняет необработанный отказ reader-а для отрицательного evidence.
' После исправления ошибки строки должны находиться в StabilityMeta, не Err.
Private Function ExecuteDuration(ByVal section As CSectionModel, ByVal provider As CMaterialModelProvider, _
        ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem, _
        ByRef batch As CBatchSectionCalculator, ByRef reason As String, _
        Optional ByVal referenceX As Double = 0#, Optional ByVal referenceY As Double = 0#) As Long
    Set batch = Nothing: reason = vbNullString
    On Error GoTo Failed
    Set batch = ExecuteFixture(section, provider, settings, units, True, referenceX, referenceY)
    Exit Function
Failed:
    ExecuteDuration = Err.Number: reason = Err.Description
End Function

' Проверяет подготовленные внутренние усилия и их фактический вывод. PhiL
' обязан реагировать на длительную часть; дополнительного State solve нет.
Private Sub CheckDurationResult(ByRef stats As TProfileStats, ByVal prefix As String, ByVal batch As CBatchSectionCalculator, _
        ByVal units As CUnitSystem, ByVal n As Double, ByVal mx As Double, ByVal my As Double, _
        Optional ByVal totalMx As Double = 1000000#, Optional ByVal totalMy As Double = 2000000#)
    Dim result As CStabilityResult, writer As CBatchResultWriter, source As Object
    Set result = batch.ResultAt(1).StabilityResult
    Check stats, prefix & ".calculated", result.Meta.Calculated
    Check stats, prefix & ".status", result.Meta.InternalStatus = rsSuccess
    CheckClose stats, prefix & ".N", result.SustainedN, -n, 0.01
    CheckClose stats, prefix & ".Mx", result.SustainedMoment1, mx, 0.01
    CheckClose stats, prefix & ".My", result.SustainedMoment2, my, 0.01
    CheckClose stats, prefix & ".PhiL1", result.PhiL1, DurationExpectedPhi(mx - n * 90#, totalMx + 50000# * 90#), 0.00000001
    CheckClose stats, prefix & ".PhiL2", result.PhiL2, DurationExpectedPhi(my - n * 60#, totalMy + 50000# * 60#), 0.00000001
    Check stats, prefix & ".noSolve", batch.SolverCallCount = 0
    Set writer = New CBatchResultWriter: writer.WriteSummary ThisWorkbook, batch, units
    Set source = ThisWorkbook.Names.Item("rngStabilitySummaryAnchor").RefersToRange
    CheckClose stats, prefix & ".outputN", CDbl(source.Offset(0, 18).Value2), units.InternalForceToOutput(n), 0.00000001
    CheckClose stats, prefix & ".outputM1", CDbl(source.Offset(0, 19).Value2), units.InternalMomentMagnitudeToOutput(mx), 0.00000001
    CheckClose stats, prefix & ".outputM2", CDbl(source.Offset(0, 20).Value2), units.InternalMomentMagnitudeToOutput(my), 0.00000001
End Sub

' Независимый oracle существующей формулы СП 63 для ненулевого полного
' момента относительно выбранного стержня; здесь не проверяется новая норма.
Private Function DurationExpectedPhi(ByVal sustainedMoment As Double, ByVal totalMoment As Double) As Double
    DurationExpectedPhi = 1# + sustainedMoment / totalMoment
    If DurationExpectedPhi < 1# Then DurationExpectedPhi = 1#
    If DurationExpectedPhi > 2# Then DurationExpectedPhi = 2#
End Function

' Ввод длительной нагрузки не является численной несходимостью. Ошибка должна
' быть адресной, необрезанной в subtree и не склеиваться заново writer-ом.
Private Sub CheckDurationInputError(ByRef stats As TProfileStats, ByVal prefix As String, ByVal batch As CBatchSectionCalculator, _
        ByVal units As CUnitSystem, ByVal location As String)
    Dim meta As CResultMeta, policy As CResultStatusPolicy
    Set meta = batch.ResultAt(1).StabilityMeta: Set policy = New CResultStatusPolicy
    Check stats, prefix & ".typed", meta.InternalStatus = rsInvalidInput And meta.ResultCode = rcInvalidInput
    Check stats, prefix & ".display", policy.ExternalStatus(meta) = "InputErr"
    Check stats, prefix & ".notCalculated", Not meta.Calculated
    Check stats, prefix & ".address", InStr(1, meta.ResultComment, location, vbBinaryCompare) > 0
    Check stats, prefix & ".repair", InStr(1, meta.ResultComment, "исправ", vbTextCompare) > 0 Or InStr(1, meta.ResultComment, "восстанов", vbTextCompare) > 0
    CheckComments stats, prefix, batch, units
End Sub

' Короткий машинный ключ компоненты не зависит от выбранной подписи единиц.
Private Function DurationColumnKey(ByVal column As Long) As String
    Select Case column
        Case 2: DurationColumnKey = "N"
        Case 3: DurationColumnKey = "Mx"
        Case 4: DurationColumnKey = "My"
    End Select
End Function

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

' ==================== ДЛЯ ТЕСТОВ: ОФОРМЛЕНИЕ И SNAPSHOT ====================

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
    ConfigureProfiles profiles: EnableChecks profiles, 1, "Yes", "No", "No", "No"
    ConfigureDurationUnits "N", "N*mm", "Tension", "+Y tension", "+X tension"
    SetTableValue system, "Calculation.ZeroMomentPerDepth", "0", 2
    SetTableValue system, "Load.ReferenceOffsetX", "0", 2: SetTableValue system, "Load.ReferenceOffsetY", "0", 2
    SetTableValue system, "Plot.LoadCase", "PROFILE_CONFIG", 2
    ConfigureAnnotationFixture annotations
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    BuildFixture section, provider, settings.GetBoolean("General.DiagramExtension", False)
    section.Annotations.AddDimension "AuditDim", -100#, 0#, 100#, 0#, 0#, 1#, "AUDIT_DIM", 200#
    section.Annotations.AddRebarLabel "AuditRebar", -80#, -60#, 80#, -60#, 0#, -1#, "AUDIT_REBAR"
    SetLoad loads, 1, -50000#, 1000000#, 2000000#, "Auto"
    Set batch = ExecuteFixture(section, provider, settings, units)
    Set writer = New CNDMResultsWriter: writer.WriteResults ThisWorkbook, section, provider, batch, units
    Set reader = New CSectionPlotDataReader: reader.LoadFromWorkbook ThisWorkbook, settings
    before = ThisWorkbook.Names.Item("rngNDMElementResults").RefersToRange.CurrentRegion.Value2
    solves = SectionEquilibriumSolveCount()
    TestAnnotationLayoutValues stats, annotations
    TestAnnotationInputMatrix stats, annotations, reader
    TestAnnotationChartValues stats, annotations, reader
    TestPresentationSnapshotIntegrity stats, reader
    TestAnnotationRelocation stats, annotations, reader, sheet
    Check stats, "presentationConfig.noSolve", SectionEquilibriumSolveCount() = solves
    Check stats, "presentationConfig.snapshotUnchanged", TablesEqual(before, _
        ThisWorkbook.Names.Item("rngNDMElementResults").RefersToRange.CurrentRegion.Value2)
    GoTo Restore
FailedRun:
    stats.Failed = stats.Failed + 1
    LogLine stats, "FAIL: presentationConfig.runtime; " & CStr(Err.Number) & "; " & Err.Description
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
    If restoreNumber <> 0 Then stats.Failed = stats.Failed + 1: LogLine stats, "FAIL: presentationConfig.restore; " & CStr(restoreNumber) & "; " & restoreReason
    LogLine stats, "TOTAL_AUDIT03_PRESENTATION_CONFIG: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed) & "; cases=" & CStr(stats.Cases)
    passed = stats.Passed: failed = stats.Failed: RunAudit03PresentationConfigTests = stats.Report
End Function

' Задает известную базу оформления в существующей двухколоночной таблице.
' Символы "-" неприменимых полей не превращаются в пользовательские настройки.
Private Sub ConfigureAnnotationFixture(ByVal source As Object)
    Dim column As Long, field As Variant
    For column = 2 To 3
        For Each field In Array("Enabled", "LineEnabled"): AnnotationCell(source, CStr(field), column).Value2 = "Yes": Next field
        AnnotationCell(source, "Placement", column).Value2 = "Outside"
        AnnotationCell(source, "Offset", column).Value2 = 50#
        AnnotationCell(source, "LineWeight", column).Value2 = 2#
        AnnotationCell(source, "TextUnits", column).Value2 = "pt"
        AnnotationCell(source, "TextHeight", column).Value2 = 13#
        AnnotationCell(source, "TextGap", column).Value2 = 9#
        AnnotationCell(source, "Color", column).Value2 = "20,30,90"
    Next column
    AnnotationCell(source, "LineEnabled", 3).Value2 = "-"
    AnnotationCell(source, "ExtensionLineWeight", 2).Value2 = "-"
    AnnotationCell(source, "ExtensionLineWeight", 3).Value2 = 1#
    AnnotationCell(source, "ArrowType", 2).Value2 = "-": AnnotationCell(source, "ArrowType", 3).Value2 = "Triangle"
    AnnotationCell(source, "ArrowSize", 2).Value2 = "-": AnnotationCell(source, "ArrowSize", 3).Value2 = "Medium"
    AnnotationCell(source, "ExtensionLineColor", 2).Value2 = "-"
    AnnotationCell(source, "ExtensionLineColor", 3).Value2 = "140,140,140"
End Sub

' Находит действительную входную ячейку аннотации, не используя ее адрес Config.
Private Function AnnotationCell(ByVal source As Object, ByVal field As String, ByVal column As Long) As Object
    Dim row As Long
    For row = 2 To source.Rows.Count
        If CStr(source.Cells(row, 1).Value2) = field Then Set AnnotationCell = source.Cells(row, column): Exit Function
    Next row
    Err.Raise vbObjectError + 4502, "AnnotationCell", "В таблице аннотаций нет " & field
End Function

' Независимые численные ожидания: масштаб модели 0.43 pt/mm, visual-scale=1.
' Проверяются настоящий цвет 0, направление Inside, толщина и mm/pt, а также
' повторная инициализация того же layout без накопления старых элементов.
Private Sub TestAnnotationLayoutValues(ByRef stats As TProfileStats, ByVal source As Object)
    Dim column As Long, color As Variant, height As Variant, layout As CPlotAnnotationLayout
    Dim settings As CSystemSettingsReader, prefix As String, base As Long
    For column = 2 To 3
        For Each color In Array("0,0,0", "255,0,0", "0,255,0", "0,0,255")
            ConfigureAnnotationFixture source
            AnnotationCell(source, "Color", column).Value2 = color
            AnnotationCell(source, "ExtensionLineColor", 3).Value2 = color
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
            If column = 2 Then Check stats, prefix & ".text", layout.Color(6) = base Else Check stats, prefix & ".text", layout.Color(4) = base
            Check stats, prefix & ".extension", layout.Color(1) = base And layout.Color(2) = base
            CheckClose stats, prefix & ".lineWeight", layout.Weight(3), 2#, 0#
            CheckClose stats, prefix & ".offset", layout.Y1(3), 193.5, 0.000000000001
            layout.Initialize -500#, 500#, -500#, 500#, 0#, 0#, 640#, 430#, "mm"
            Check stats, prefix & ".reset", layout.Count = 0
            stats.Cases = stats.Cases + 1
        Next color
        For Each height In Array(13#, 20#)
            ConfigureAnnotationFixture source
            AnnotationCell(source, "TextHeight", column).Value2 = height
            AnnotationCell(source, "TextUnits", column).Value2 = "mm"
            AnnotationCell(source, "Placement", column).Value2 = "Inside"
            Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
            Set layout = New CPlotAnnotationLayout: layout.ConfigureFromSettings settings
            layout.Initialize -500#, 500#, -500#, 500#, 0#, 0#, 640#, 430#, "mm"
            layout.AddDimension -100#, 0#, 100#, 0#, 0#, 1#, "DIM"
            layout.AddRebarLabel -100#, 0#, 100#, 0#, 0#, 1#, "REBAR"
            prefix = "presentationConfig." & CStr(column) & ".TextUnits.mm." & CStr(height)
            If column = 2 Then base = 6 Else base = 4
            CheckClose stats, prefix & ".font", layout.FontSize(base), CDbl(height) * 0.43, 0.000000000001
            If column = 2 Then
                Check stats, prefix & ".Inside", layout.Y1(6) > 215#
            Else
                Check stats, prefix & ".Inside", layout.Y1(4) > 184.5
            End If
            stats.Cases = stats.Cases + 1
        Next height
    Next column
End Sub

' Матрица всех применимых полей: пустота/текст/Excel-ошибка и особые границы.
' Ошибка активного поля должна содержать key, текущую ячейку и исправление;
' тот же ошибочный параметр выключенной группы не должен мешать схеме.
Private Sub TestAnnotationInputMatrix(ByRef stats As TProfileStats, ByVal source As Object, ByVal reader As CSectionPlotDataReader)
    Dim column As Long, field As Variant, value As Variant, invalids As Variant, cell As Object
    Dim prefix As String, key As String, code As Long, reason As String, caseIndex As Long, savedKey As Variant
    For column = 2 To 3
        For Each field In Array("Enabled", "Placement", "Offset", "LineEnabled", "LineWeight", "ExtensionLineWeight", _
                "TextUnits", "TextHeight", "TextGap", "ArrowType", "ArrowSize", "Color", "ExtensionLineColor")
            ConfigureAnnotationFixture source
            Set cell = AnnotationCell(source, CStr(field), column)
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
                    ConfigureAnnotationFixture source: cell.Value2 = value
                    prefix = "presentationConfig." & key & ".invalid" & CStr(caseIndex)
                    LogLine stats, "RUN: " & prefix
                    code = PresentationDrawError(reader, reason)
                    Check stats, prefix & ".rejected", code <> 0
                    If code <> 0 Then
                        Check stats, prefix & ".controlled", code <> 6 And code <> 13 And code <> 9
                        Check stats, prefix & ".key", InStr(1, reason, key, vbTextCompare) > 0
                        Check stats, prefix & ".address", InStr(1, reason, "ячейка " & cell.Address(False, False), vbBinaryCompare) > 0
                        Check stats, prefix & ".repair", InStr(1, reason, "Введите", vbTextCompare) > 0 Or InStr(1, reason, "Выберите", vbTextCompare) > 0 Or InStr(1, reason, "Исправьте", vbTextCompare) > 0
                    End If
                    If CStr(field) <> "Enabled" Then
                        AnnotationCell(source, "Enabled", column).Value2 = "No"
                        code = PresentationDrawError(reader, reason)
                        Check stats, prefix & ".inactive", code = 0
                    End If
                    caseIndex = caseIndex + 1: stats.Cases = stats.Cases + 1
                Next value
                ConfigureAnnotationFixture source
                savedKey = cell.Offset(0, 1 - column).Value2: cell.Offset(0, 1 - column).Value2 = "__MissingAnnotationField"
                If CStr(field) <> "Enabled" Then
                    If column = 2 Then AnnotationCell(source, "Enabled", 3).Value2 = "No" Else AnnotationCell(source, "Enabled", 2).Value2 = "No"
                End If
                code = PresentationDrawError(reader, reason)
                Check stats, "presentationConfig." & key & ".missing", code <> 0
                If code <> 0 Then Check stats, "presentationConfig." & key & ".missingKey", InStr(1, reason, CStr(field), vbTextCompare) > 0
                cell.Offset(0, 1 - column).Value2 = savedKey
                ConfigureAnnotationFixture source: code = PresentationDrawError(reader, reason)
                Check stats, "presentationConfig." & key & ".recovery", code = 0
            End If
        Next field
    Next column
End Sub

' Проверяет собственно Excel Shape, а не только accessor layout: черный цвет
' не подменяется; Enabled и LineEnabled действительно убирают нужные элементы.
Private Sub TestAnnotationChartValues(ByRef stats As TProfileStats, ByVal source As Object, ByVal reader As CSectionPlotDataReader)
    Dim shape As Object, chart As Object, code As Long, reason As String, count As Long, arrows As Long
    ConfigureAnnotationFixture source
    AnnotationCell(source, "Color", 2).Value2 = "0,0,0": AnnotationCell(source, "Color", 3).Value2 = "0,0,0"
    AnnotationCell(source, "ExtensionLineColor", 3).Value2 = "0,0,0"
    code = PresentationDrawError(reader, reason): Check stats, "presentationConfig.Chart.black.draw", code = 0
    If code <> 0 Then Exit Sub
    Set chart = ThisWorkbook.Worksheets.Item("Расчет").ChartObjects.Item("chtNDMSectionPlot").Chart
    For Each shape In chart.Shapes
        If InStr(1, shape.Name, "AnnotationLine", vbBinaryCompare) > 0 Then
            count = count + 1: Check stats, "presentationConfig.Chart.black.line" & CStr(count), shape.Line.ForeColor.RGB = 0
            If shape.Line.BeginArrowheadStyle <> 1 Then arrows = arrows + 1
        End If
    Next shape
    Check stats, "presentationConfig.Chart.black.lines", count = 4
    Check stats, "presentationConfig.Chart.black.arrows", arrows = 1
    Check stats, "presentationConfig.Chart.black.image", chart.Export(ThisWorkbook.Path & "\Audit03_presentation_black.png", "PNG")
    LogLine stats, "PLOT_IMAGE: " & ThisWorkbook.Path & "\Audit03_presentation_black.png"
    AnnotationCell(source, "Enabled", 3).Value2 = "No": AnnotationCell(source, "LineEnabled", 2).Value2 = "No"
    code = PresentationDrawError(reader, reason): Check stats, "presentationConfig.Chart.disabled.draw", code = 0
    count = 0
    For Each shape In chart.Shapes
        If InStr(1, shape.Name, "AnnotationLine", vbBinaryCompare) > 0 Then count = count + 1
    Next shape
    Check stats, "presentationConfig.Chart.disabled.lines", count = 0
    ConfigureAnnotationFixture source
End Sub

' Перемещает существующий Name и проверяет ошибку активного цвета по новому
' адресу. Отрисовка не должна выдавать старый адрес или молча брать default.
Private Sub TestAnnotationRelocation(ByRef stats As TProfileStats, ByVal source As Object, _
        ByVal reader As CSectionPlotDataReader, ByRef sheet As Object)
    Dim target As Object, cell As Object, code As Long, reason As String, oldRef As String
    oldRef = ThisWorkbook.Names.Item("rngPlotAnnotationSettings").RefersTo
    ConfigureAnnotationFixture source
    Set sheet = ThisWorkbook.Worksheets.Add: sheet.Name = "__AuditPlotInputs"
    Set target = sheet.Cells(25, 60).Resize(source.Rows.Count, source.Columns.Count): target.Value2 = source.Value2
    ThisWorkbook.Names.Item("rngPlotAnnotationSettings").RefersTo = "='" & sheet.Name & "'!" & target.Address
    Set cell = AnnotationCell(target, "Color", 3): cell.Value2 = "256,0,0"
    code = PresentationDrawError(reader, reason)
    Check stats, "presentationConfig.relocated.rejected", code <> 0
    If code <> 0 Then
        Check stats, "presentationConfig.relocated.sheet", InStr(1, reason, sheet.Name, vbBinaryCompare) > 0
        Check stats, "presentationConfig.relocated.address", InStr(1, reason, "ячейка " & cell.Address(False, False), vbBinaryCompare) > 0
    End If
    ThisWorkbook.Names.Item("rngPlotAnnotationSettings").RefersTo = oldRef
End Sub

' Запускает production-рисование по уже подготовленному snapshot. Код ошибки
' сохраняется отдельно от текста: тест не выводит статус из комментария.
Private Function PresentationDrawError(ByVal reader As CSectionPlotDataReader, ByRef reason As String) As Long
    Dim settings As CSystemSettingsReader, plotter As CSectionPlotter
    reason = vbNullString
    On Error GoTo Failed
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set plotter = New CSectionPlotter: plotter.Draw ThisWorkbook, reader, settings
    Exit Function
Failed:
    PresentationDrawError = Err.Number: reason = Err.Description
End Function

' Проверяет целостность реальных таблиц Results: недостающая обязательная
' колонка и испорченная аннотация не выдаются за успешно прочитанную схему.
' ShapeType остается допустимым именем колонки геометрии, пустая таблица
' semantic-аннотаций не блокирует саму расчетную геометрию.
Private Sub TestPresentationSnapshotIntegrity(ByRef stats As TProfileStats, ByRef reader As CSectionPlotDataReader)
    Dim geometry As Object, annotations As Object, oldGeometry As Variant, oldAnnotations As Variant
    Dim data As Variant, code As Long, reason As String, field As Variant, column As Long
    Dim savedNumber As Long, savedReason As String
    On Error GoTo Failed
    Set geometry = PresentationSnapshotRange("rngNDMSectionGeometry", 25)
    Set annotations = PresentationSnapshotRange("rngNDMSectionAnnotations", 13)
    oldGeometry = geometry.Formula: oldAnnotations = annotations.Formula
    data = geometry.Value2: data(1, 7) = "__MissingGeometryStatus": geometry.Value2 = data
    code = PresentationReaderError(reader, reason)
    Check stats, "presentationConfig.snapshot.geometryHeader.rejected", code <> 0
    Check stats, "presentationConfig.snapshot.geometryHeader.controlled", code <> 9 And code <> 13 And code <> 6
    If code <> 0 Then Check stats, "presentationConfig.snapshot.geometryHeader.reason", InStr(1, reason, "GeometryInterpretationStatus", vbBinaryCompare) > 0
    data(1, 7) = "ShapeType": geometry.Value2 = data
    code = PresentationReaderError(reader, reason): Check stats, "presentationConfig.snapshot.ShapeType", code = 0
    geometry.Formula = oldGeometry
    For Each field In Array("StartX", "StartY", "EndX", "EndY", "OutsideNormalX", "OutsideNormalY", "Text")
        data = annotations.Value2
        column = PresentationColumn(data, CStr(field))
        data(2, column) = CVErr(xlErrValue): annotations.Value2 = data
        code = PresentationReaderError(reader, reason)
        Check stats, "presentationConfig.snapshot.annotation." & CStr(field) & ".rejected", code <> 0
        If code <> 0 Then
            Check stats, "presentationConfig.snapshot.annotation." & CStr(field) & ".controlled", code <> 9 And code <> 13 And code <> 6
            Check stats, "presentationConfig.snapshot.annotation." & CStr(field) & ".reason", InStr(1, reason, CStr(field), vbBinaryCompare) > 0
            Check stats, "presentationConfig.snapshot.annotation." & CStr(field) & ".address", _
                InStr(1, reason, annotations.Cells(2, column).Address(False, False), vbBinaryCompare) > 0
        End If
        annotations.Formula = oldAnnotations
        code = PresentationReaderError(reader, reason)
        Check stats, "presentationConfig.snapshot.annotation." & CStr(field) & ".recovery", code = 0
        stats.Cases = stats.Cases + 1
    Next field
    data = annotations.Value2: column = PresentationColumn(data, "Text")
    data(2, column) = "  Подпись с пробелами  ": annotations.Value2 = data
    code = PresentationReaderError(reader, reason)
    Check stats, "presentationConfig.snapshot.textWhitespace.reader", code = 0
    If code = 0 Then Check stats, "presentationConfig.snapshot.textWhitespace.preserved", reader.AnnotationText(1) = "  Подпись с пробелами  "
    annotations.Formula = oldAnnotations
    data = annotations.Value2: data(1, 4) = "__MissingStartX": annotations.Value2 = data
    code = PresentationReaderError(reader, reason)
    Check stats, "presentationConfig.snapshot.annotationHeader.rejected", code <> 0
    annotations.Formula = oldAnnotations
    annotations.Offset(1, 0).Resize(annotations.Rows.Count - 1, annotations.Columns.Count).ClearContents
    code = PresentationReaderError(reader, reason)
    Check stats, "presentationConfig.snapshot.emptyAnnotations.reader", code = 0
    If code = 0 Then Check stats, "presentationConfig.snapshot.emptyAnnotations.count", reader.AnnotationCount = 0
    GoTo Restore
Failed:
    savedNumber = Err.Number: savedReason = Err.Description
Restore:
    On Error Resume Next
    If Not geometry Is Nothing Then geometry.Formula = oldGeometry
    If Not annotations Is Nothing Then annotations.Formula = oldAnnotations
    On Error GoTo 0
    If savedNumber <> 0 Then stats.Failed = stats.Failed + 1: LogLine stats, "FAIL: presentationConfig.snapshot.runtime; " & CStr(savedNumber) & "; " & savedReason
    code = PresentationReaderError(reader, reason): Check stats, "presentationConfig.snapshot.recovery", code = 0
End Sub

' Ограничивает test-mutation собственным блоком от именованного якоря.
' CurrentRegion здесь непригоден: соседние Results-блоки могут соприкасаться
' служебными строками. Число колонок задано независимо по проверяемой схеме.
Private Function PresentationSnapshotRange(ByVal name As String, ByVal columns As Long) As Object
    Dim anchor As Object, rows As Long
    Set anchor = ThisWorkbook.Names.Item(name).RefersToRange
    rows = 1
    Do While Len(CStr(anchor.Offset(rows, 0).Value2)) > 0: rows = rows + 1: Loop
    Set PresentationSnapshotRange = anchor.Resize(rows, columns)
End Function

' Находит колонку самостоятельного test-oracle по буквальному заголовку.
' Это не production ResultColumn и не общий helper, скрывающий ошибку decoder-а.
Private Function PresentationColumn(ByRef data As Variant, ByVal header As String) As Long
    Dim column As Long
    For column = 1 To UBound(data, 2)
        If CStr(data(1, column)) = header Then PresentationColumn = column: Exit Function
    Next column
    Err.Raise vbObjectError + 4503, "PresentationColumn", "В snapshot нет " & header
End Function

' Возвращает точную ошибку загрузки на новом reader-е и сохраняет объект для
' проверки последующего восстановления; пользовательские настройки не меняет.
Private Function PresentationReaderError(ByRef reader As CSectionPlotDataReader, ByRef reason As String) As Long
    Dim settings As CSystemSettingsReader
    reason = vbNullString
    On Error GoTo Failed
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set reader = New CSectionPlotDataReader: reader.LoadFromWorkbook ThisWorkbook, settings
    Exit Function
Failed:
    PresentationReaderError = Err.Number: reason = Err.Description
End Function

' ДЛЯ ТЕСТОВ: независимо вычисляет выбранную величину каждого элемента
' из сохраненной плоскости и фактической диаграммы named-state. Reader не
' может незаметно подменить Stress на Strain или взять другое состояние.
Private Sub CheckVisualizationValues(ByRef stats As TProfileStats, ByVal prefix As String, _
        ByVal reader As CSectionPlotDataReader, ByVal state As CSectionStateResult, _
        ByVal section As CSectionModel, ByVal provider As CMaterialModelProvider, ByVal units As CUnitSystem)
    Dim concrete As CMaterialDiagram, steel As CMaterialDiagram, i As Long, bar As Long
    Dim x As Double, y As Double, expected As Double, strain As Double, diagram As CMaterialDiagram
    If state.ExtensionUsed Then
        Set concrete = provider.ConcreteMaterialForEquilibriumFromSpec(state.MaterialSpec)
        Set steel = provider.SteelMaterialForEquilibriumFromSpec(state.MaterialSpec)
    Else
        Set concrete = provider.ConcreteMaterialFromSpec(state.MaterialSpec)
        Set steel = provider.SteelMaterialFromSpec(state.MaterialSpec)
    End If
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
        CheckClose stats, prefix & ".element" & CStr(i), reader.ResultValue(i), expected, 0.000000000001
    Next i
End Sub

' ДЛЯ ТЕСТОВ: проверяет именно текст растянутого края общей легенды, а не
' случайное совпадение подписи нуля в Chart. Положение определяется
' контрактом оформления; точность формируется независимым test-oracle.
Private Sub CheckVisualizationLegend(ByRef stats As TProfileStats, ByVal prefix As String, _
        ByVal reader As CSectionPlotDataReader, ByVal settings As CSystemSettingsReader)
    Dim plotter As CSectionPlotter, chartObject As Object, shape As Object, i As Long
    Dim maximum As Double, foundValue As Boolean, foundText As Boolean, text As String, pattern As String
    For i = 1 To reader.Count
        If reader.PhysicalState(i) = "Tension" Then
            If Not foundValue Or Abs(reader.ResultValue(i)) > Abs(maximum) Then maximum = reader.ResultValue(i)
            foundValue = True
        End If
    Next i
    Check stats, prefix & ".legendFixtureTension", foundValue
    If Not foundValue Then Exit Sub
    pattern = "0": If reader.ResultPrecision > 0 Then pattern = pattern & "." & String$(reader.ResultPrecision, "0")
    text = Replace$(Format$(maximum, pattern), ",", ".")
    Set plotter = New CSectionPlotter: plotter.Draw ThisWorkbook, reader, settings
    Set chartObject = ThisWorkbook.Worksheets.Item("Расчет").ChartObjects.Item("chtNDMSectionPlot")
    For Each shape In chartObject.Chart.Shapes
        If Left$(shape.Name, Len("NDMPlot_Text")) = "NDMPlot_Text" Then
            If Abs(shape.Left - (chartObject.Width - 118#)) < 0.01 And Abs(shape.Top - 174#) < 0.01 Then
                foundText = True
                Check stats, prefix & ".actualLegendText", CStr(shape.TextFrame.Characters().Text) = text
            End If
        End If
    Next shape
    Check stats, prefix & ".actualLegendPresent", foundText
End Sub
