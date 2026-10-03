Attribute VB_Name = "modTestLoadTableConfig"
Option Explicit

' ==========================================================================
' ДЛЯ ТЕСТОВ: АДРЕСНОЕ ПОВЕДЕНИЕ ТАБЛИЦЫ СОЧЕТАНИЙ CONFIG
' ==========================================================================
' Проверяет все штатные строки rngLoadCombinations через настоящий reader,
' batch, конечные State/results и writers. Табличные индексы, комментарии и
' выбор профиля проверяются отдельно от физических усилий и lambda-пути.
' Временные значения и формулы Config восстанавливаются при любом исходе;
' модуль не добавляет расчетных классов и не меняет утвержденные формулы.

Private Type TLoadTableStats
    Passed As Long
    Failed As Long
    Cases As Long ' Отдельные строки LC, а не количество assertions.
    Report As String
End Type

Private Const FORCE_FACTOR As Double = 9806.65 ' Независимый эталон tf -> N.
Private Const MOMENT_FACTOR As Double = 9806650# ' Независимый эталон tf*m -> N*mm.
Private Const ROW_COUNT As Long = 30 ' Поадресная приемка штатного шаблона, не ограничение runtime.
Private mDiagnosticRun As Boolean ' Только отдельный диагностический entrypoint включает подробные solver logs.

' Выполняет целиком адресный блок LC. Optional-счетчики позволяют включить
' отчет в штатную suite без второго запуска и без подмены отдельных failures.
Public Function RunAudit03LoadTableTests(Optional ByRef passed As Long = 0, _
        Optional ByRef failed As Long = 0) As String
    Dim stats As TLoadTableStats
    Dim loads As Object, system As Object, profilesRange As Object, unitRange As Object, signRange As Object
    Dim savedLoads As Variant, savedSystem As Variant, savedProfiles As Variant
    Dim savedUnits As Variant, savedSigns As Variant, baseline As Variant
    Dim settings As CSystemSettingsReader, units As CUnitSystem
    Dim profiles As CCalculationProfileCatalog, section As CSectionModel, provider As CMaterialModelProvider
    Dim column As Long, row As Long, data As Variant, path As Variant, batch As CBatchSectionCalculator
    On Error GoTo FailedRun
    Set loads = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange
    Set system = ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange
    Set profilesRange = ThisWorkbook.Names.Item("rngCalculationProfiles").RefersToRange
    Set unitRange = ThisWorkbook.Names.Item("rngUnitSettings").RefersToRange
    Set signRange = ThisWorkbook.Names.Item("rngSignConventionSettings").RefersToRange
    savedLoads = loads.Formula: savedSystem = system.Formula: savedProfiles = profilesRange.Formula
    savedUnits = unitRange.Formula: savedSigns = signRange.Formula
    If loads.Rows.Count <> ROW_COUNT + 1 Or loads.Columns.Count <> 7 Then _
        Err.Raise vbObjectError + 4499, "RunAudit03LoadTableTests", "Адресный fixture требует штатные 30 строк и семь колонок."

    SetSetting system, "Calculation.ZeroMomentPerDepth", "0"
    If mDiagnosticRun Then SetSetting system, "General.ExecutionReportEnabled", "Yes"
    SetSetting system, "Solver.ToleranceN", FormatNumberInvariant(0.01 / FORCE_FACTOR)
    SetSetting system, "Solver.ToleranceMx", FormatNumberInvariant(1# / MOMENT_FACTOR)
    SetSetting system, "Solver.ToleranceMy", FormatNumberInvariant(1# / MOMENT_FACTOR)
    SetSetting unitRange, "Force", "tf", 2
    SetSetting unitRange, "Moment", "tf*m", 2
    SetSetting signRange, "+N", "Compression", 2
    SetSetting signRange, "+Mx", "+Y tension", 2
    SetSetting signRange, "+My", "+X tension", 2
    ConfigureProfiles profilesRange, False
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem
    units.LoadFromSettings settings
    Set profiles = New CCalculationProfileCatalog
    profiles.LoadFromWorkbook ThisWorkbook
    BuildPhysicalFixture section, provider, settings.GetBoolean("General.DiagramExtension", False)
    LogLine stats, "LOAD_TABLE_CONFIG: addresses=210; effectiveExtension=" & CStr(provider.DiagramExtensionEnabled)
    baseline = BaselineRows(loads)

    loads.Value2 = baseline
    LogLine stats, "RUN: loadTable.baseline"
    Set batch = ReadAndExecute(section, provider, profiles, settings, units)
    CheckValidRows stats, loads, batch, units, "baseline", 0, False
    batch.Execute
    CheckValidRows stats, loads, batch, units, "repeat", 0, False

    ' Пустая компонента нагрузки означает ноль; optional comment и path
    ' имеют собственные контракты, но не делают ненулевую строку свободной.
    For Each path In Array(2, 3, 4, 6, 7)
        column = CLng(path): data = baseline
        LogLine stats, "RUN: loadTable.blank." & CStr(column)
        For row = 2 To ROW_COUNT + 1
            data(row, column) = vbNullString
        Next row
        loads.Value2 = data
        Set batch = ReadAndExecute(section, provider, profiles, settings, units)
        CheckValidRows stats, loads, batch, units, "blank." & CStr(column), column, False
    Next path

    LogLine stats, "RUN: loadTable.invalid"
    CheckInvalidRows stats, loads, baseline, section, provider, profiles, settings, units
    LogLine stats, "RUN: loadTable.sparseAndReverse"
    CheckSparseAndReverse stats, loads, baseline, section, provider, profiles, settings, units

    ' В этом блоке пути действительно используются обоими Search consumers,
    ' а не проверяются только через строковый accessor описателя траектории.
    ConfigureProfiles profilesRange, True
    Set profiles = New CCalculationProfileCatalog
    profiles.LoadFromWorkbook ThisWorkbook
    For Each path In Array("Auto", "LambdaMx", "LambdaMy", "LambdaMxy", "LambdaN", "LambdaNMxy")
        LogLine stats, "RUN: loadTable.path." & CStr(path)
        data = baseline
        For row = 2 To ROW_COUNT + 1
            data(row, 6) = CStr(path)
        Next row
        loads.Value2 = data
        Set batch = ReadAndExecute(section, provider, profiles, settings, units)
        CheckActivePaths stats, loads, batch, units, CStr(path)
    Next path
    GoTo Restore
FailedRun:
    stats.Failed = stats.Failed + 1
    LogLine stats, "FAIL: loadTable.runtime; " & CStr(Err.Number) & "; " & Err.Description
Restore:
    On Error Resume Next
    If Not loads Is Nothing Then loads.Formula = savedLoads
    If Not system Is Nothing Then system.Formula = savedSystem
    If Not profilesRange Is Nothing Then profilesRange.Formula = savedProfiles
    If Not unitRange Is Nothing Then unitRange.Formula = savedUnits
    If Not signRange Is Nothing Then signRange.Formula = savedSigns
    On Error GoTo 0
    LogLine stats, "TOTAL_AUDIT03_LOAD_TABLE: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed) & "; cases=" & CStr(stats.Cases)
    passed = stats.Passed: failed = stats.Failed
    RunAudit03LoadTableTests = stats.Report
End Function

' Сохраняет реальные заголовки и задает уникальные данные каждой строки.
' Малые физические нагрузки позволяют отличить перестановку компонентов
' и строк от физического отказа; все четыре ProfileId проходят через каталог.
Private Function BaselineRows(ByVal source As Object) As Variant
    Dim data As Variant, slot As Long
    data = source.Value2
    For slot = 1 To ROW_COUNT
        data(slot + 1, 1) = "ATL" & Format$(slot, "00")
        data(slot + 1, 2) = -1# - CDbl(slot) / 100#
        data(slot + 1, 3) = 0.02 + CDbl(slot) / 1000#
        data(slot + 1, 4) = -0.01 - CDbl(slot) / 2000#
        data(slot + 1, 5) = "PR" & CStr((slot - 1) Mod 4 + 1)
        data(slot + 1, 6) = "Auto"
        data(slot + 1, 7) = "Комментарий строки " & CStr(slot) & "; проверка порядка"
    Next slot
    BaselineRows = data
End Function

' Готовит один небольшой бетонный mesh и симметричные стержни для всей серии.
' Материальные числа фиксированы независимо от пользовательского ввода;
' Extension сохраняет эффективное значение режима текущего прогона.
Private Sub BuildPhysicalFixture(ByRef section As CSectionModel, ByRef provider As CMaterialModelProvider, _
        ByVal extensionEnabled As Boolean)
    Dim geometry As CGeometryRoundedRectangle, mesh As CFiberMeshBuilder, rebars As CRebarLayout
    Set geometry = New CGeometryRoundedRectangle
    geometry.Initialize 300#, 200#, 0#, 0#, 0#, 0#
    Set mesh = New CFiberMeshBuilder
    mesh.BuildMesh geometry, 30#, 20#, 1
    Set rebars = New CRebarLayout
    rebars.AddBar "A", -90#, -60#, 20#, 0#, "A400", "", geometry
    rebars.AddBar "B", 90#, -60#, 20#, 0#, "A400", "", geometry
    rebars.AddBar "C", -90#, 60#, 20#, 0#, "A400", "", geometry
    rebars.AddBar "D", 90#, 60#, 20#, 0#, "A400", "", geometry
    Set section = BuildGeneratedSectionModel(mesh, rebars, "Audit03LoadTable")
    Dim concrete As CConcreteMaterialParameters, steel As CSteelMaterialParameters
    Set concrete = New CConcreteMaterialParameters
    concrete.Initialize 15.5, 1.1, 22#, 1.8, 32500#, 32500#, rbMc2:=14.6
    Set steel = New CSteelMaterialParameters
    steel.Initialize 350#, 350#, 390#, 390#, 200000#, 200000#
    Set provider = New CMaterialModelProvider
    provider.InitializeFromParameters concrete, steel, diagramExtensionEnabled:=extensionEnabled
End Sub

' Читает именно именованный диапазон Config, исполняет штатный batch и
' передает тому же writer готовые результаты без повторного решения НДС.
Private Function ReadAndExecute(ByVal section As CSectionModel, ByVal provider As CMaterialModelProvider, _
        ByVal profiles As CCalculationProfileCatalog, ByVal settings As CSystemSettingsReader, _
        ByVal units As CUnitSystem) As CBatchSectionCalculator
    Dim batch As CBatchSectionCalculator, reader As CLoadCombinationReader
    Set batch = New CBatchSectionCalculator
    batch.Initialize section, provider
    Set batch.ProfileCatalog = profiles
    batch.ApplySettings settings, units
    If mDiagnosticRun Then
        Dim report As CExecutionReport
        Set report = New CExecutionReport
        report.Initialize ThisWorkbook, settings
        Set batch.ExecutionReport = report
    End If
    Set reader = New CLoadCombinationReader
    reader.LoadFromWorkbook ThisWorkbook, batch, units
    batch.Execute
    Set ReadAndExecute = batch
End Function

' Проверяет единицы/знаки по независимым коэффициентам, target найденного
' НДС и соответствие конкретной строки Config конкретной строке Results.
' Повторный Execute и обратный порядок используют тот же физический oracle.
Private Sub CheckValidRows(ByRef stats As TLoadTableStats, ByVal source As Object, _
        ByVal batch As CBatchSectionCalculator, ByVal units As CUnitSystem, ByVal scenario As String, _
        ByVal blankColumn As Long, ByVal reversed As Boolean)
    Dim writer As CBatchResultWriter, summary As Object, point As CSectionStateResult
    Dim slot As Long, originalSlot As Long, prefix As String, n As Double, mx As Double, my As Double
    Dim solves As Long, expectedComment As String
    Set writer = New CBatchResultWriter
    solves = batch.SolverCallCount
    writer.WriteSummary ThisWorkbook, batch, units
    Set summary = ThisWorkbook.Names.Item("rngBatchSummary").RefersToRange
    Check stats, "loadTable." & scenario & ".count", batch.Count = ROW_COUNT
    Check stats, "loadTable." & scenario & ".writerNoSolve", batch.SolverCallCount = solves
    For slot = 1 To ROW_COUNT
        stats.Cases = stats.Cases + 1
        originalSlot = slot
        If reversed Then originalSlot = ROW_COUNT + 1 - slot
        prefix = "loadTable.Slot" & CStr(slot) & "." & scenario
        n = (1# + CDbl(originalSlot) / 100#) * FORCE_FACTOR
        mx = (0.02 + CDbl(originalSlot) / 1000#) * MOMENT_FACTOR
        my = (-0.01 - CDbl(originalSlot) / 2000#) * MOMENT_FACTOR
        If blankColumn = 2 Then n = 0#
        If blankColumn = 3 Then mx = 0#
        If blankColumn = 4 Then my = 0#
        expectedComment = "Комментарий строки " & CStr(originalSlot) & "; проверка порядка"
        If blankColumn = 7 Then expectedComment = vbNullString
        Check stats, prefix & ".CombinationID", batch.CombinationID(slot) = "ATL" & Format$(originalSlot, "00")
        Check stats, prefix & ".sourceSlot", batch.SourceDataOffset(slot) = slot
        CheckClose stats, prefix & ".N", batch.N(slot), n, 0.00000001
        CheckClose stats, prefix & ".Mx", batch.Mx(slot), mx, 0.00000001
        CheckClose stats, prefix & ".My", batch.My(slot), my, 0.00000001
        Check stats, prefix & ".ProfileId", batch.ResultAt(slot).ProfileId = "PR" & CStr((originalSlot - 1) Mod 4 + 1)
        Check stats, prefix & ".Comment", batch.CombinationName(slot) = expectedComment
        Check stats, prefix & ".directStatus", batch.ResultAt(slot).DirectStateMeta.InternalStatus = rsSuccess
        Set point = batch.ResultAt(slot).StrengthResult.DirectState.StateResult
        Check stats, prefix & ".stateExists", Not point Is Nothing
        If Not point Is Nothing Then
            CheckClose stats, prefix & ".targetN", point.TargetN, n, 0.00000001
            CheckClose stats, prefix & ".targetMx", point.TargetMx, mx, 0.00000001
            CheckClose stats, prefix & ".targetMy", point.TargetMy, my, 0.00000001
        End If
        Check stats, prefix & ".writerID", CStr(summary.Cells(slot + 12, 1).Value2) = batch.CombinationID(slot)
        Check stats, prefix & ".writerComment", CStr(summary.Cells(slot + 12, 2).Value2) = expectedComment
        Check stats, prefix & ".writerResultComment", CStr(summary.Cells(slot + 12, 3).Value2) = batch.ResultAt(slot).OverallMeta.ResultComment
        Check stats, prefix & ".writerStatus", CStr(summary.Cells(slot + 12, 4).Value2) = batch.ResultAt(slot).Status
    Next slot
End Sub

' Для каждой колонки проверяет ошибку формулы; дополнительно проверяет
' обязательные пустые ID, неизвестный путь, текст и overflow компонент.
' Ошибочные строки не решаются и сохраняют собственный адрес/причину в summary.
Private Sub CheckInvalidRows(ByRef stats As TLoadTableStats, ByVal source As Object, ByVal baseline As Variant, _
        ByVal section As CSectionModel, ByVal provider As CMaterialModelProvider, _
        ByVal profiles As CCalculationProfileCatalog, ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem)
    Dim column As Long, caseIndex As Long, slot As Long, data As Variant, batch As CBatchSectionCalculator
    Dim writer As CBatchResultWriter, summary As Object, prefix As String, reason As String, address As String
    For column = 1 To 7
        For caseIndex = 1 To 3
            If caseIndex = 3 And (column < 2 Or column > 4) Then Exit For
            If caseIndex = 2 And column = 7 Then Exit For
            data = baseline
            For slot = 1 To ROW_COUNT
                Select Case caseIndex
                    Case 1: data(slot + 1, column) = CVErr(xlErrNA)
                    Case 2
                        If column = 1 Or column = 5 Then data(slot + 1, column) = vbNullString Else data(slot + 1, column) = "unknown"
                    Case 3: data(slot + 1, column) = 1E+308
                End Select
            Next slot
            source.Value2 = data
            Set batch = ReadAndExecute(section, provider, profiles, settings, units)
            Set writer = New CBatchResultWriter
            writer.WriteSummary ThisWorkbook, batch, units
            Set summary = ThisWorkbook.Names.Item("rngBatchSummary").RefersToRange
            prefix = "loadTable.invalid.col" & CStr(column) & ".case" & CStr(caseIndex)
            Check stats, prefix & ".count", batch.Count = ROW_COUNT
            Check stats, prefix & ".noSolve", batch.SolverCallCount = 0
            For slot = 1 To ROW_COUNT
                stats.Cases = stats.Cases + 1
                reason = batch.ResultAt(slot).OverallMeta.ResultComment
                address = source.Worksheet.Name & "!" & source.Cells(slot + 1, column).Address(False, False)
                Check stats, prefix & ".Slot" & CStr(slot) & ".InputErr", batch.ResultAt(slot).Status = "InputErr"
                Check stats, prefix & ".Slot" & CStr(slot) & ".address", InStr(1, reason, address, vbBinaryCompare) > 0
                Check stats, prefix & ".Slot" & CStr(slot) & ".sourceSlot", batch.SourceDataOffset(slot) = slot
                Check stats, prefix & ".Slot" & CStr(slot) & ".writerComment", CStr(summary.Cells(slot + 12, 3).Value2) = reason
                Check stats, prefix & ".Slot" & CStr(slot) & ".writerStatus", CStr(summary.Cells(slot + 12, 4).Value2) = "InputErr"
                LogLine stats, "LOAD_TABLE_ERROR: " & prefix & ".Slot" & CStr(slot) & "; " & reason
            Next slot
        Next caseIndex
    Next column
End Sub

' Нулевые и полностью пустые строки пропускаются, но их места в Results
' остаются свободными. Перестановка всех строк не меняет физические ответы.
Private Sub CheckSparseAndReverse(ByRef stats As TLoadTableStats, ByVal source As Object, ByVal baseline As Variant, _
        ByVal section As CSectionModel, ByVal provider As CMaterialModelProvider, _
        ByVal profiles As CCalculationProfileCatalog, ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem)
    Dim data As Variant, slot As Long, column As Long, batch As CBatchSectionCalculator
    Dim writer As CBatchResultWriter, summary As Object, index As Long
    data = baseline
    For slot = 1 To ROW_COUNT
        If slot Mod 3 = 0 Then
            data(slot + 1, 2) = 0#: data(slot + 1, 3) = 0#: data(slot + 1, 4) = 0#
        ElseIf slot Mod 3 = 1 Then
            For column = 1 To 7
                data(slot + 1, column) = vbNullString
            Next column
        End If
    Next slot
    source.Value2 = data
    Set batch = ReadAndExecute(section, provider, profiles, settings, units)
    Check stats, "loadTable.sparse.count", batch.Count = 10
    Set writer = New CBatchResultWriter
    writer.WriteSummary ThisWorkbook, batch, units
    Set summary = ThisWorkbook.Names.Item("rngBatchSummary").RefersToRange
    For slot = 1 To ROW_COUNT
        If slot Mod 3 = 2 Then
            index = index + 1: stats.Cases = stats.Cases + 1
            Check stats, "loadTable.sparse.Slot" & CStr(slot) & ".sourceSlot", batch.SourceDataOffset(index) = slot
            Check stats, "loadTable.sparse.Slot" & CStr(slot) & ".writerID", CStr(summary.Cells(slot + 12, 1).Value2) = "ATL" & Format$(slot, "00")
        Else
            Check stats, "loadTable.sparse.Slot" & CStr(slot) & ".emptyOutput", Len(CStr(summary.Cells(slot + 12, 1).Value2)) = 0
        End If
    Next slot
    data = baseline
    For slot = 1 To ROW_COUNT
        For column = 1 To 7
            data(slot + 1, column) = baseline(ROW_COUNT + 2 - slot, column)
        Next column
    Next slot
    source.Value2 = data
    Set batch = ReadAndExecute(section, provider, profiles, settings, units)
    CheckValidRows stats, source, batch, units, "reverse", 0, True
End Sub

' Проверяет фактически найденные Capacity и Formation точки по независимому
' разложению усилий. Auto ожидает первый допустимый Mxy-путь этого fixture;
' фиксированный путь не должен незаметно подменяться другим.
Private Sub CheckActivePaths(ByRef stats As TLoadTableStats, ByVal source As Object, _
        ByVal batch As CBatchSectionCalculator, ByVal units As CUnitSystem, ByVal requested As String)
    Dim slot As Long, prefix As String, actual As String, lambda As Double, point As CSectionStateResult
    Dim writer As CBatchResultWriter, solves As Long, expected As String, descriptor As CLoadPathDescriptor
    Set descriptor = New CLoadPathDescriptor
    expected = requested
    If expected = "Auto" Then expected = "LambdaMxy"
    Set writer = New CBatchResultWriter
    solves = batch.SolverCallCount
    writer.WriteSummary ThisWorkbook, batch, units
    Check stats, "loadTable.path." & requested & ".count", batch.Count = ROW_COUNT
    Check stats, "loadTable.path." & requested & ".writerNoSolve", batch.SolverCallCount = solves
    For slot = 1 To ROW_COUNT
        stats.Cases = stats.Cases + 1
        prefix = "loadTable.Slot" & CStr(slot) & ".LoadPath." & requested
        actual = descriptor.NormalizeKey(batch.ResultAt(slot).StrengthResult.Capacity.PathResolved)
        Check stats, prefix & ".capacityPath", actual = expected
        Check stats, prefix & ".capacityStatus", batch.ResultAt(slot).CapacityMeta.InternalStatus = rsSuccess
        Set point = batch.ResultAt(slot).StrengthResult.Capacity.StateResult
        lambda = batch.ResultAt(slot).StrengthResult.Capacity.LambdaCapacity
        CheckPathPoint stats, prefix & ".capacityPoint", point, actual, lambda, batch.N(slot), batch.Mx(slot), batch.My(slot)
        actual = descriptor.NormalizeKey(batch.ResultAt(slot).CrackResult.Formation.FormationMethod)
        Check stats, prefix & ".formationPath", actual = expected
        Check stats, prefix & ".formationStatus", batch.ResultAt(slot).CrackResult.Formation.ResultMeta.InternalStatus = rsSuccess
        Set point = batch.ResultAt(slot).CrackResult.Formation.PreCrackState
        lambda = batch.ResultAt(slot).CrackResult.Formation.LambdaCrc
        CheckPathPoint stats, prefix & ".formationPoint", point, actual, lambda, batch.N(slot), batch.Mx(slot), batch.My(slot)
        Check stats, prefix & ".currentState", batch.ResultAt(slot).CrackResult.CurrentStateMeta.InternalStatus = rsSuccess
        Check stats, prefix & ".summaryComment", CStr(ThisWorkbook.Names.Item("rngBatchSummary").RefersToRange.Cells(slot + 12, 3).Value2) = batch.ResultAt(slot).OverallMeta.ResultComment
        LogLine stats, "LOAD_TABLE_PATH: " & prefix & _
            "; capacityStatus=" & CStr(batch.ResultAt(slot).CapacityMeta.InternalStatus) & _
            "; capacityCode=" & CStr(batch.ResultAt(slot).CapacityMeta.ResultCode) & _
            "; capacity=" & batch.ResultAt(slot).CapacityMeta.ResultComment & _
            "; strength=" & batch.ResultAt(slot).StrengthMeta.ResultComment & _
            "; crack=" & batch.ResultAt(slot).CrackResult.ResultMeta.ResultComment & _
            "; stability=" & batch.ResultAt(slot).StabilityMeta.ResultComment & _
            "; overall=" & batch.ResultAt(slot).OverallMeta.ResultComment
        If mDiagnosticRun And batch.ResultAt(slot).CapacityMeta.InternalStatus = rsNumericalFailure Then _
            LogLine stats, "LOAD_TABLE_CAPACITY_DIAGNOSTIC: " & prefix & vbCrLf & _
                batch.ResultAt(slot).StrengthResult.Capacity.DiagnosticLog
    Next slot
End Sub

' Проверяет сохраненный target и равновесие без повторного solve.
' Select Case задает независимый эталон пяти фиксированных траекторий.
Private Sub CheckPathPoint(ByRef stats As TLoadTableStats, ByVal prefix As String, _
        ByVal point As CSectionStateResult, ByVal path As String, ByVal lambda As Double, _
        ByVal n As Double, ByVal mx As Double, ByVal my As Double)
    Check stats, prefix & ".exists", Not point Is Nothing
    If point Is Nothing Then Exit Sub
    Select Case path
        Case "LambdaMx": mx = lambda * mx
        Case "LambdaMy": my = lambda * my
        Case "LambdaMxy": mx = lambda * mx: my = lambda * my
        Case "LambdaN": n = lambda * n
        Case "LambdaNMxy": n = lambda * n: mx = lambda * mx: my = lambda * my
        Case Else: Check stats, prefix & ".knownPath", False: Exit Sub
    End Select
    Check stats, prefix & ".converged", point.Converged
    Check stats, prefix & ".physicalRange", point.WithinPhysicalRange
    Check stats, prefix & ".noExtension", Not point.ExtensionUsed
    CheckClose stats, prefix & ".targetN", point.TargetN, n, 0.000001
    CheckClose stats, prefix & ".targetMx", point.TargetMx, mx, 0.000001
    CheckClose stats, prefix & ".targetMy", point.TargetMy, my, 0.000001
    CheckClose stats, prefix & ".equilibriumN", point.Nint, n, 0.01
    CheckClose stats, prefix & ".equilibriumMx", point.Mxint, mx, 1#
    CheckClose stats, prefix & ".equilibriumMy", point.Myint, my, 1#
End Sub

' Настраивает четыре реальные колонки профилей, не меняя их material specs
' и presentation-поля. Во второй фазе оба Search consumers становятся активны.
Private Sub ConfigureProfiles(ByVal source As Object, ByVal withSearch As Boolean)
    Dim row As Long, column As Long, key As String
    For row = 1 To source.Rows.Count
        key = CStr(source.Cells(row, 2).Value2)
        For column = 3 To 6
            Select Case key
                Case "Calculation.Strength.DirectState": source.Cells(row, column).Value2 = "Yes"
                Case "Calculation.Strength.Capacity", "Calculation.Crack.Width"
                    If withSearch Then source.Cells(row, column).Value2 = "Yes" Else source.Cells(row, column).Value2 = "No"
                Case "Calculation.Stability.Enabled": source.Cells(row, column).Value2 = "No"
            End Select
        Next column
    Next row
End Sub

' Находит настройку по фактическому ключу таблицы; отсутствующий ключ в fixture
' является ошибкой тестовой постановки и не заменяется другим default.
Private Sub SetSetting(ByVal source As Object, ByVal key As String, ByVal value As String, _
        Optional ByVal column As Long = 2)
    Dim row As Long
    For row = 1 To source.Rows.Count
        If StrComp(CStr(source.Cells(row, 1).Value2), key, vbTextCompare) = 0 Then
            source.Cells(row, column).Value2 = value
            Exit Sub
        End If
    Next row
    Err.Raise vbObjectError + 4499, "SetSetting", "В fixture не найден ключ " & key
End Sub

' Сохраняет отдельный проверяемый Boolean-контракт и его стабильный test-ID.
Private Sub Check(ByRef stats As TLoadTableStats, ByVal id As String, ByVal condition As Boolean)
    If condition Then
        stats.Passed = stats.Passed + 1
        LogLine stats, "OK: " & id
    Else
        stats.Failed = stats.Failed + 1
        LogLine stats, "FAIL: " & id
    End If
End Sub

' Сравнивает физические числа с независимым ожидаемым значением; в логе
' сохраняет actual/expected, чтобы последующие срезы не могли скрыть дрейф.
Private Sub CheckClose(ByRef stats As TLoadTableStats, ByVal id As String, _
        ByVal actual As Double, ByVal expected As Double, ByVal tolerance As Double)
    Check stats, id & "|actual=" & FormatNumberInvariant(actual) & "|expected=" & FormatNumberInvariant(expected), Abs(actual - expected) <= tolerance
End Sub

' Добавляет строку отчета; крупные этапы также сохраняются рядом с тестовой
' книгой, чтобы остановленный COM-вызов не терял место ошибки или зависания.
Private Sub LogLine(ByRef stats As TLoadTableStats, ByVal value As String)
    stats.Report = stats.Report & value & vbCrLf
    If Left$(value, 5) = "RUN: " Then
        Dim stream As Object
        Set stream = CreateObject("Scripting.FileSystemObject").OpenTextFile( _
            ThisWorkbook.Path & "\RC_NDM_ui_test_progress.txt", 8, True, -1)
        stream.WriteLine Format$(Now, "yyyy-mm-dd hh:nn:ss") & " " & value
        stream.Close
    End If
End Sub

' Сохраняет числовые параметры и actual/expected через точку, включая малые
' допуски в scientific notation; соседние private helpers других suites
' намеренно не используются как несуществующее общее production API.
Private Function FormatNumberInvariant(ByVal value As Double) As String
    FormatNumberInvariant = Trim$(Str$(value))
End Function

' ДЛЯ ТЕСТОВ: повторяет тот же адресный набор и те же строгие assertions,
' добавляя журнал solver-а только для несошедшейся Capacity. Численные
' настройки и expected values не заменяются ради диагностического прогона.
Public Function RunAudit03LoadTableDiagnosticTests() As String
    mDiagnosticRun = True
    RunAudit03LoadTableDiagnosticTests = RunAudit03LoadTableTests()
    mDiagnosticRun = False
End Function

' ДЛЯ ТЕСТОВ: повторяет только выгрузку одной готовой ошибочной строки, без
' solve. Выбор writer-а отделяет расход памяти summary от численных Search;
' соседние контрольные ячейки проверяют границы очистки и заливки заголовков.
Public Function RunAudit03SummaryWriterDiagnosticTests(Optional ByVal writerName As String = "Batch", _
        Optional ByVal repetitions As Long = 12, Optional ByRef passed As Long = 0, _
        Optional ByRef failed As Long = 0) As String
    Dim stats As TLoadTableStats, section As CSectionModel, provider As CMaterialModelProvider
    Dim settings As CSystemSettingsReader, units As CUnitSystem, batch As CBatchSectionCalculator
    Dim profiles As CCalculationProfileCatalog
    Dim strength As CStrengthSummaryWriter, crack As CCrackSummaryWriter, stability As CStabilitySummaryWriter
    Dim writer As CBatchResultWriter, anchors As Variant, heights As Variant
    Dim snapshot As CNDMResultsWriter, results As Object
    Dim guard As CExcelAppStateGuard
    Dim saved(1 To 4, 1 To 5) As Variant, cells(1 To 4) As Object, i As Long, iteration As Long
    Dim columnWidths() As Double, columnCount As Long, column As Long
    On Error GoTo FailedRun
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem
    units.LoadFromSettings settings
    If writerName = "GuardedBatch" Or writerName = "OriginalWidthsGuardedBatch" Then
        Set guard = New CExcelAppStateGuard
        guard.Enter ThisWorkbook.Application
    End If
    LogLine stats, "RUN: writerDiagnostic.application; screenUpdating=" & CStr(ThisWorkbook.Application.ScreenUpdating) & _
        "; calculation=" & CStr(ThisWorkbook.Application.Calculation)
    BuildPhysicalFixture section, provider, False
    Set batch = New CBatchSectionCalculator
    batch.Initialize section, provider
    Set profiles = New CCalculationProfileCatalog
    profiles.LoadFromWorkbook ThisWorkbook
    Set batch.ProfileCatalog = profiles
    batch.ApplySettings settings, units
    If writerName = "ThirtyRowsBatch" Then
        For i = 1 To ROW_COUNT
            batch.AddInvalidCombination "WRITER_DIAGNOSTIC_" & CStr(i), "PR1", "", "Диагностическая строка: нагрузка не задана."
        Next i
    Else
        batch.AddInvalidCombination "WRITER_DIAGNOSTIC", "PR1", "", "Диагностическая строка: нагрузка не задана."
    End If
    batch.Execute
    Check stats, "writerDiagnostic.noSolve", batch.SolverCallCount = 0
    Set writer = New CBatchResultWriter
    Set strength = New CStrengthSummaryWriter
    Set crack = New CCrackSummaryWriter
    Set stability = New CStabilitySummaryWriter
    Set snapshot = New CNDMResultsWriter
    Set results = ThisWorkbook.Worksheets.Item("Results")
    anchors = Array("rngStrengthSummaryAnchor", "rngCrackSummaryAnchor", "rngStabilitySummaryAnchor", "rngNDMElementResults")
    heights = Array(strength.HeaderRowsAboveAnchor, crack.HeaderRowsAboveAnchor, stability.HeaderRowsAboveAnchor, 1)
    For i = 1 To 4
        Set cells(i) = ThisWorkbook.Names.Item(CStr(anchors(i - 1))).RefersToRange.Offset(-CLng(heights(i - 1)) - 1, 199)
        saved(i, 1) = cells(i).Formula
        saved(i, 2) = cells(i).Interior.Pattern
        saved(i, 3) = cells(i).Interior.Color
        saved(i, 4) = cells(i).NumberFormat
        If i <= 3 Then saved(i, 5) = results.Columns.Item(i).ColumnWidth
        cells(i).Value2 = "WRITER_SCOPE_SENTINEL"
        cells(i).Interior.Color = RGB(7, 23, 47)
        cells(i).NumberFormat = "@"
        If i <= 3 And writerName <> "OriginalWidthsBatch" And writerName <> "OriginalWidthsGuardedBatch" And _
                writerName <> "ThirtyRowsBatch" Then results.Columns.Item(i).ColumnWidth = 12# + 3# * CDbl(i)
    Next i
    ' Контролируем и символьную, и физическую ширину всего рабочего блока,
    ' включая численные колонки, у которых смена шрифта меняла default-ширину.
    columnCount = stability.RequiredColumns
    ReDim columnWidths(1 To columnCount, 1 To 2)
    For column = 1 To columnCount
        columnWidths(column, 1) = results.Columns.Item(column).ColumnWidth
        columnWidths(column, 2) = results.Columns.Item(column).Width
    Next column
    For iteration = 1 To repetitions
        LogLine stats, "RUN: writerDiagnostic." & writerName & "." & CStr(iteration) & "; " & Audit03ExcelMemory()
        Select Case writerName
            Case "Batch", "GuardedBatch", "OriginalWidthsBatch", "OriginalWidthsGuardedBatch", "ThirtyRowsBatch"
                writer.ClearSummary ThisWorkbook
                writer.WriteSummary ThisWorkbook, batch, units, section
            Case "Strength"
                strength.ClearSummary ThisWorkbook
                strength.WriteSummary ThisWorkbook, batch, units, section
            Case "Crack"
                crack.ClearSummary ThisWorkbook
                crack.WriteSummary ThisWorkbook, batch, units
            Case "Stability"
                stability.ClearSummary ThisWorkbook
                stability.WriteSummary ThisWorkbook, batch, units
            Case "Snapshot"
                snapshot.ClearResults ThisWorkbook
                snapshot.WriteResults ThisWorkbook, section, provider, batch, units
            Case "Preview"
                snapshot.ClearResults ThisWorkbook
                snapshot.WriteGeometryPreview ThisWorkbook, section, units
            Case Else: Err.Raise vbObjectError + 4500, "RunAudit03SummaryWriterDiagnosticTests", "Неизвестный диагностический writer."
        End Select
        Check stats, "writerDiagnostic." & writerName & "." & CStr(iteration) & ".noSolve", batch.SolverCallCount = 0
    Next iteration
    LogLine stats, "RUN: writerDiagnostic." & writerName & ".completed; " & Audit03ExcelMemory()
    For i = 1 To 4
        Check stats, "writerDiagnostic." & writerName & ".scope" & CStr(i) & ".value", CStr(cells(i).Value2) = "WRITER_SCOPE_SENTINEL"
        Check stats, "writerDiagnostic." & writerName & ".scope" & CStr(i) & ".fill", cells(i).Interior.Color = RGB(7, 23, 47)
        Check stats, "writerDiagnostic." & writerName & ".scope" & CStr(i) & ".format", CStr(cells(i).NumberFormat) = "@"
        If i <= 3 Then
            If writerName = "OriginalWidthsBatch" Or writerName = "OriginalWidthsGuardedBatch" Or _
                    writerName = "ThirtyRowsBatch" Then
                Check stats, "writerDiagnostic." & writerName & ".width" & CStr(i), results.Columns.Item(i).ColumnWidth = saved(i, 5)
            Else
                Check stats, "writerDiagnostic." & writerName & ".width" & CStr(i), results.Columns.Item(i).ColumnWidth = 12# + 3# * CDbl(i)
            End If
        End If
    Next i
    For column = 1 To columnCount
        Check stats, "writerDiagnostic." & writerName & ".allWidths." & CStr(column), _
            results.Columns.Item(column).ColumnWidth = columnWidths(column, 1)
        Check stats, "writerDiagnostic." & writerName & ".allPoints." & CStr(column), _
            results.Columns.Item(column).Width = columnWidths(column, 2)
    Next column
    GoTo Restore
FailedRun:
    stats.Failed = stats.Failed + 1
    LogLine stats, "FAIL: writerDiagnostic.runtime; " & CStr(Err.Number) & "; " & Err.Description
Restore:
    On Error Resume Next
    For i = 1 To 4
        If Not cells(i) Is Nothing Then
            cells(i).Formula = saved(i, 1)
            cells(i).Interior.Color = saved(i, 3)
            cells(i).Interior.Pattern = saved(i, 2)
            cells(i).NumberFormat = saved(i, 4)
            If i <= 3 Then results.Columns.Item(i).ColumnWidth = saved(i, 5)
        End If
    Next i
    If Not guard Is Nothing Then guard.Restore
    On Error GoTo 0
    LogLine stats, "TOTAL_AUDIT03_WRITER_DIAGNOSTIC: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    passed = stats.Passed: failed = stats.Failed
    RunAudit03SummaryWriterDiagnosticTests = stats.Report
End Function

' ДЛЯ ТЕСТОВ: проверяет пользовательскую ширину в summary, трех подробных
' блоках, конечном снимке и геометрическом preview. Каждый consumer повторен
' дважды; временные ширины и соседние контрольные ячейки затем восстанавливаются.
Public Function RunAudit03ResultColumnWidthTests(Optional ByRef passed As Long = 0, _
        Optional ByRef failed As Long = 0) As String
    Dim name As Variant, childPassed As Long, childFailed As Long, report As String
    passed = 0: failed = 0
    For Each name In Array("Batch", "Strength", "Crack", "Stability", "Snapshot", "Preview")
        report = RunAudit03SummaryWriterDiagnosticTests(CStr(name), 2, childPassed, childFailed)
        RunAudit03ResultColumnWidthTests = RunAudit03ResultColumnWidthTests & report
        passed = passed + childPassed: failed = failed + childFailed
    Next name
    RunAudit03ResultColumnWidthTests = RunAudit03ResultColumnWidthTests & _
        "TOTAL_AUDIT03_RESULT_COLUMN_WIDTHS: passed=" & CStr(passed) & "; failed=" & CStr(failed) & vbCrLf
End Function

' ДЛЯ ТЕСТОВ: сохраняет память с PID, чтобы открытый пользователем Excel
' не подменял счетчики отдельного тестового процесса. Метод не влияет на solve.
Private Function Audit03ExcelMemory() As String
    Dim process As Object
    On Error GoTo Unavailable
    For Each process In GetObject("winmgmts:").ExecQuery("SELECT ProcessId,PrivatePageCount FROM Win32_Process WHERE Name='EXCEL.EXE'")
        If Len(Audit03ExcelMemory) > 0 Then Audit03ExcelMemory = Audit03ExcelMemory & "; "
        Audit03ExcelMemory = Audit03ExcelMemory & "pid=" & CStr(process.ProcessId) & "; privateBytes=" & CStr(process.PrivatePageCount)
    Next process
    Exit Function
Unavailable:
    Audit03ExcelMemory = "memory unavailable"
End Function
