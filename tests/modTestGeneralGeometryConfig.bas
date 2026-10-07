Attribute VB_Name = "modTestGeneralGeometryConfig"
Option Explicit

' ==========================================================================
' ДЛЯ ТЕСТОВ: ОБЩИЕ НАСТРОЙКИ, СЕТКА И ТОЧКА ПРИЛОЖЕНИЯ НАГРУЗКИ
' ==========================================================================
' Проверяет девять фактических настроек через их рабочие consumers: построение
' CSectionModel, сохраненный импорт, batch и перенос нагрузок Excel-adapter-ом.
' Ни геометрия, ни формулы solver-а не воспроизводятся в тестовом calculator-е.
' Изменения делаются только в изолированной книге runner-а; исходные формулы
' и привязка rngSystemSettings восстанавливаются даже после ошибки проверки.

Private Type TGeneralConfigStats
    Passed As Long
    Failed As Long
    Report As String
    Cases As Long ' Число consumer-сценариев отдельно от числа assertions.
End Type

' Выполняет поведенческие и адресные проверки в двух положениях таблицы.
' Второе положение специально выявляет сообщения с жестко заданным адресом.
Public Function RunAudit03GeneralGeometryConfigTests(Optional ByRef passed As Long = 0, _
        Optional ByRef failed As Long = 0) As String
    Dim stats As TGeneralConfigStats, system As Object, unitsRange As Object, profiles As Object, circleRange As Object
    Dim savedSystem As Variant, savedUnits As Variant, savedProfiles As Variant, savedCircle As Variant
    Dim shifted As Object, savedShifted As Variant, originalName As String, position As Long
    On Error GoTo FailedRun
    Set system = ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange
    Set unitsRange = ThisWorkbook.Names.Item("rngUnitSettings").RefersToRange
    Set profiles = ThisWorkbook.Names.Item("rngCalculationProfiles").RefersToRange
    Set circleRange = ThisWorkbook.Names.Item("rngCircleGeometry").RefersToRange
    savedSystem = system.Formula: savedUnits = unitsRange.Formula
    savedProfiles = profiles.Formula: savedCircle = circleRange.Formula
    originalName = ThisWorkbook.Names.Item("rngSystemSettings").RefersTo
    Set shifted = system.Worksheet.Range("CH800").Resize(system.Rows.Count, system.Columns.Count)
    savedShifted = shifted.Formula
    ConfigureGeneralFixture system, unitsRange, profiles, circleRange
    CheckGeometryAndMesh stats, system
    CheckGeneratedMeshUnits stats, system, unitsRange, circleRange
    CheckSavedImport stats, system, unitsRange
    CheckMomentFilter stats, system, unitsRange
    CheckReferenceOffsets stats, system, unitsRange
    CheckExtensionConsumer stats, system, profiles
    Dim prepared As Variant: prepared = system.Formula
    For position = 0 To 1
        If position = 1 Then
            shifted.Formula = prepared
            ThisWorkbook.Names.Item("rngSystemSettings").RefersTo = "='" & shifted.Worksheet.Name & "'!" & shifted.Address
            Set system = shifted
        End If
        CheckInvalidGeneralFields stats, system, position
    Next position
    GoTo Restore
FailedRun:
    stats.Failed = stats.Failed + 1
    LogLine stats, "FAIL: generalGeometry.runtime; " & CStr(Err.Number) & "; " & Err.Description
Restore:
    On Error GoTo FailedRestore
    If Len(originalName) > 0 Then ThisWorkbook.Names.Item("rngSystemSettings").RefersTo = originalName
    If Not IsEmpty(savedSystem) Then ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange.Formula = savedSystem
    If Not shifted Is Nothing Then shifted.Formula = savedShifted
    If Not unitsRange Is Nothing Then unitsRange.Formula = savedUnits
    If Not profiles Is Nothing Then profiles.Formula = savedProfiles
    If Not circleRange Is Nothing Then circleRange.Formula = savedCircle
    Check stats, "generalGeometry.restore.name", ThisWorkbook.Names.Item("rngSystemSettings").RefersTo = originalName
    GoTo Finish
FailedRestore:
    stats.Failed = stats.Failed + 1
    LogLine stats, "FAIL: generalGeometry.restore; " & CStr(Err.Number) & "; " & Err.Description
Finish:
    passed = stats.Passed: failed = stats.Failed
    LogLine stats, "GENERAL_GEOMETRY_CASES: " & CStr(stats.Cases)
    LogLine stats, "TOTAL_GENERAL_GEOMETRY_CONFIG: passed=" & CStr(passed) & "; failed=" & CStr(failed)
    RunAudit03GeneralGeometryConfigTests = stats.Report
End Function

' Задает малый армированный круг и прямое НДС без устойчивости/несущей.
' Варианты единиц меняются на самом Config, а не в локальных test getters.
Private Sub ConfigureGeneralFixture(ByVal system As Object, ByVal unitsRange As Object, _
        ByVal profiles As Object, ByVal circleRange As Object)
    Dim row As Long, value As String
    For row = 2 To unitsRange.Rows.Count
        Select Case CStr(unitsRange.Cells(row, 1).Value2)
            Case "Length": value = "mm"
            Case "Area": value = "mm2"
            Case "Force": value = "N"
            Case "Moment": value = "N*mm"
            Case "Stress": value = "MPa"
            Case "Curvature": value = "1/mm"
            Case Else: Err.Raise vbObjectError + 4499, "ConfigureGeneralFixture", "Неизвестная строка единиц."
        End Select
        unitsRange.Cells(row, 2).Value2 = value: unitsRange.Cells(row, 4).Value2 = value
    Next row
    SetValue system, "Geometry.Source", "Generated"
    SetValue system, "Geometry.Type", "Circle"
    SetValue system, "Mesh.StepX", 20#: SetValue system, "Mesh.StepY", 20#
    SetValue system, "Mesh.BoundarySubdivisions", 2
    SetValue system, "Load.ReferenceOffsetX", 0#: SetValue system, "Load.ReferenceOffsetY", 0#
    SetValue system, "Calculation.ZeroMomentPerDepth", 0#
    SetValue system, "General.DiagramExtension", "No"
    SetValue system, "General.ExecutionReportEnabled", "Yes"
    SetValue system, "Solver.ToleranceN", 0.01
    SetValue system, "Solver.ToleranceMx", 1#: SetValue system, "Solver.ToleranceMy", 1#
    SetValue system, "SLS.Crack.PsiMode", "User": SetValue system, "SLS.Crack.PsiS", 1#
    SetValue circleRange, "Circle.Diameter", 300#
    SetValue circleRange, "Rebar.AxisDistance", 30#: SetValue circleRange, "Rebar.Count", 12
    SetValue circleRange, "Rebar.Diameter", 16#
    SetValue circleRange, "Rebar.Diameter2", 0#: SetValue circleRange, "Rebar.Diameter3", 0#
    SetProfile profiles, "Calculation.Stability.Enabled", "No", "PR1"
    SetProfile profiles, "Calculation.Strength.DirectState", "Yes", "PR1"
    SetProfile profiles, "Calculation.Strength.Capacity", "No", "PR1"
    SetProfile profiles, "Calculation.Crack.Width", "No", "PR1"
    SetProfile profiles, "Calculation.Stability.Enabled", "No", "PR2"
    SetProfile profiles, "Calculation.Strength.DirectState", "No", "PR2"
    SetProfile profiles, "Calculation.Strength.Capacity", "No", "PR2"
    SetProfile profiles, "Calculation.Crack.Width", "Yes", "PR2"
End Sub

' Каждый тип реально строит бетон/арматуру и собственные semantic-аннотации.
' Изменения шагов проверяются отдельно по фактической площади элементов;
' subdivisions проверяется на кривой границе, где влияет на дискретизацию.
Private Sub CheckGeometryAndMesh(ByRef stats As TGeneralConfigStats, ByVal system As Object)
    Dim shape As Variant, model As CSectionModel, firstArea As Double, prefix As String
    For Each shape In Array("Circle", "RectSet", "RoundedRectangle", "HollowRectangle")
        SetValue system, "Geometry.Type", CStr(shape)
        Set model = BuildModel(stats)
        prefix = "generalGeometry.Geometry.Type." & CStr(shape)
        Check stats, prefix & ".concrete", model.ConcreteCount > 0
        Check stats, prefix & ".steel", model.RebarCount > 0
        Check stats, prefix & ".annotations", model.AnnotationCount > 0
        Check stats, prefix & ".generated", model.SourceType = "Generated"
        If firstArea > 0# Then Check stats, prefix & ".differentSection", Abs(ConcreteArea(model) - firstArea) > 1#
        firstArea = ConcreteArea(model)
        LogLine stats, "GENERAL_MODEL: " & prefix & "|count=" & CStr(model.ConcreteCount) & "|area=" & NumberText(firstArea)
    Next shape
    SetValue system, "Geometry.Type", "Circle"
    Dim baseline As CSectionModel, changed As CSectionModel, key As Variant
    Set baseline = BuildModel(stats)
    For Each key In Array("Mesh.StepX", "Mesh.StepY")
        SetValue system, CStr(key), 10#
        Set changed = BuildModel(stats)
        prefix = "generalGeometry." & CStr(key)
        Check stats, prefix & ".moreFibers", changed.ConcreteCount > baseline.ConcreteCount
        Check stats, prefix & ".cellArea", HasCellArea(changed, 200#)
        Check stats, prefix & ".steelUnchanged", changed.RebarCount = baseline.RebarCount
        CheckClose stats, prefix & ".areaAccuracy", ConcreteArea(changed), GEOM_PI * 150# ^ 2, 1000#
        SetValue system, CStr(key), 20#
    Next key
    SetValue system, "Mesh.StepX", 55#: SetValue system, "Mesh.StepY", 55#
    SetValue system, "Mesh.BoundarySubdivisions", 1
    Set baseline = BuildModel(stats)
    SetValue system, "Mesh.BoundarySubdivisions", 4
    Set changed = BuildModel(stats)
    ' Независимая квадратура окружности: 22 центра шага 55 мм и 378 центров
    ' шага 13.75 мм внутри x^2+y^2 <= 150^2, включая последние неполные клетки.
    CheckClose stats, "generalGeometry.Mesh.BoundarySubdivisions.coarseArea", ConcreteArea(baseline), 66550#, 0.00000001
    CheckClose stats, "generalGeometry.Mesh.BoundarySubdivisions.fineArea", ConcreteArea(changed), 71465.625, 0.00000001
    Check stats, "generalGeometry.Mesh.BoundarySubdivisions.boundaryChanged", _
        Abs(ConcreteArea(changed) - ConcreteArea(baseline)) > 1#
    Check stats, "generalGeometry.Mesh.BoundarySubdivisions.accuracy", _
        Abs(ConcreteArea(changed) - GEOM_PI * 150# ^ 2) < Abs(ConcreteArea(baseline) - GEOM_PI * 150# ^ 2)
    SetValue system, "Mesh.StepX", 20#: SetValue system, "Mesh.StepY", 20#
    SetValue system, "Mesh.BoundarySubdivisions", 2
End Sub

' Перевод всех активных длин Circle и шагов мм -> м сохраняет каждый
' фактический бетонный элемент и стержень. Не сравнивает две неподвижные
' таблицы: обе модели заново собираются штатным Generated-entrypoint.
Private Sub CheckGeneratedMeshUnits(ByRef stats As TGeneralConfigStats, ByVal system As Object, _
        ByVal unitRange As Object, ByVal circleRange As Object)
    Dim baseline As CSectionModel, changed As CSectionModel, savedCircle As Variant, i As Long, prefix As String
    savedCircle = circleRange.Formula
    Set baseline = BuildModel(stats)
    SetValue unitRange, "Length", "m"
    SetValue circleRange, "Circle.Diameter", 0.3: SetValue circleRange, "Rebar.AxisDistance", 0.03
    SetValue circleRange, "Rebar.Diameter", 0.016
    SetValue system, "Mesh.StepX", 0.02: SetValue system, "Mesh.StepY", 0.02
    Set changed = BuildModel(stats)
    Check stats, "generalGeometry.Mesh.StepX.unitsCount", changed.ConcreteCount = baseline.ConcreteCount
    Check stats, "generalGeometry.Mesh.StepY.unitsSteelCount", changed.RebarCount = baseline.RebarCount
    If changed.ConcreteCount = baseline.ConcreteCount Then
        For i = 1 To baseline.ConcreteCount
            prefix = "generalGeometry.Mesh.units.C" & CStr(i)
            CheckClose stats, prefix & ".x", changed.ConcreteX(i), baseline.ConcreteX(i), 0.00000001
            CheckClose stats, prefix & ".y", changed.ConcreteY(i), baseline.ConcreteY(i), 0.00000001
            CheckClose stats, prefix & ".area", changed.ConcreteArea(i), baseline.ConcreteArea(i), 0.00000001
        Next i
    End If
    If changed.RebarCount = baseline.RebarCount Then
        For i = 1 To baseline.RebarCount
            prefix = "generalGeometry.Mesh.units.R" & CStr(i)
            CheckClose stats, prefix & ".x", changed.RebarX(i), baseline.RebarX(i), 0.00000001
            CheckClose stats, prefix & ".y", changed.RebarY(i), baseline.RebarY(i), 0.00000001
            CheckClose stats, prefix & ".area", changed.RebarArea(i), baseline.RebarArea(i), 0.00000001
        Next i
    End If
    circleRange.Formula = savedCircle
    SetValue unitRange, "Length", "mm"
    SetValue system, "Mesh.StepX", 20#: SetValue system, "Mesh.StepY", 20#
End Sub

' Сохраненный импорт проходит настоящий importer/Results writer/readback.
' Generated-настройки могут быть невалидны при AutoCAD: повторный импорт и
' генерация не выполняются, координаты/площадь остаются внутренними мм/мм2.
Private Sub CheckSavedImport(ByRef stats As TGeneralConfigStats, ByVal system As Object, ByVal unitRange As Object)
    Dim settings As CSystemSettingsReader, units As CUnitSystem, importer As CAutoCADSectionModelImporter
    Dim regions As Collection, region As CFakeAcadRegion, imported As CSectionModel, restored As CSectionModel
    Dim writer As CNDMResultsWriter, saved As Variant, area As Double
    saved = system.Formula
    LoadReader settings, units
    Set regions = New Collection
    Set region = New CFakeAcadRegion: region.Initialize 40000#, 20#, 30#, 133333333#, 133333333#, 0#, "RC_CONCRETE", "CFG_C"
    regions.Add region
    Set region = New CFakeAcadRegion: region.Initialize 201#, 80#, 90#, 3215#, 3215#, 0#, "RC_REBAR", "CFG_R"
    regions.Add region
    Set importer = New CAutoCADSectionModelImporter
    Set imported = importer.ImportFromModelSpace(regions, "RC_CONCRETE", "RC_REBAR", "Rebar", 0#, units)
    Set writer = New CNDMResultsWriter: writer.WriteGeometryPreview ThisWorkbook, imported, PrepareSectionSnapshot(imported), units
    SetValue system, "Geometry.Source", "AutoCAD"
    SetValue system, "Geometry.Type", "Unknown"
    SetValue system, "Mesh.StepX", -1#: SetValue system, "Mesh.StepY", -1#
    SetValue system, "Mesh.BoundarySubdivisions", -1
    Set restored = BuildModel(stats)
    Check stats, "generalGeometry.Geometry.Source.AutoCAD.noGenerator", restored.ConcreteCount = 1 And restored.RebarCount = 1
    CheckClose stats, "generalGeometry.Geometry.Source.AutoCAD.x", restored.ConcreteX(1), 20#, 0.00000001
    CheckClose stats, "generalGeometry.Geometry.Source.AutoCAD.area", ConcreteArea(restored), 40000#, 0.00000001
    SetValue unitRange, "Length", "m"
    Set restored = BuildModel(stats)
    CheckClose stats, "generalGeometry.Geometry.Source.AutoCAD.changedInputUnit", restored.ConcreteX(1), 20#, 0.00000001
    CheckClose stats, "generalGeometry.Geometry.Source.AutoCAD.changedInputArea", ConcreteArea(restored), 40000#, 0.00000001
    SetValue unitRange, "Length", "mm"
    system.Formula = saved
End Sub

' Порог проверяется ниже/ровно/выше границы через действующие target moments
' batch и реальные НДС. Другой масштаб INPUT Moment/Length обязан сохранять
' внутренний порог; нулевая настройка отключает именно инженерный фильтр.
Private Sub CheckMomentFilter(ByRef stats As TGeneralConfigStats, ByVal system As Object, ByVal unitRange As Object)
    Dim section As CSectionModel, batch As CBatchSectionCalculator, point As CSectionStateResult
    Dim value As Variant, prefix As String, factor As Double
    Set section = BuildModel(stats)
    SetValue system, "Calculation.ZeroMomentPerDepth", 10#
    For Each value In Array(2999#, 3000#, 3001#)
        prefix = "generalGeometry.Calculation.ZeroMomentPerDepth." & NumberText(CDbl(value))
        Set batch = RunGeneralBatch(stats, section, -10000#, CDbl(value), CDbl(value), "PR1")
        factor = 1#: If CDbl(value) <= 3000# Then factor = 0#
        CheckClose stats, prefix & ".mx", batch.Mx(1), factor * CDbl(value), 0#
        CheckClose stats, prefix & ".my", batch.My(1), factor * CDbl(value), 0#
        Set point = batch.ResultAt(1).StrengthResult.DirectState.StateResult
        Check stats, prefix & ".state", point.Converged
    Next value
    SetValue system, "Calculation.ZeroMomentPerDepth", 0#
    Set batch = RunGeneralBatch(stats, section, -10000#, 2999#, -2999#, "PR1")
    CheckClose stats, "generalGeometry.Calculation.ZeroMomentPerDepth.disabled", batch.Mx(1), 2999#, 0#
    SetValue unitRange, "Length", "m": SetValue unitRange, "Moment", "kN*m"
    SetValue system, "Calculation.ZeroMomentPerDepth", 0.01
    Set batch = RunGeneralBatch(stats, section, -10000#, 3000#, 3001#, "PR1")
    CheckClose stats, "generalGeometry.Calculation.ZeroMomentPerDepth.unitsMx", batch.Mx(1), 0#, 0#
    CheckClose stats, "generalGeometry.Calculation.ZeroMomentPerDepth.unitsMy", batch.My(1), 3001#, 0#
    SetValue unitRange, "Length", "mm": SetValue unitRange, "Moment", "N*mm"
    SetValue system, "Calculation.ZeroMomentPerDepth", 0#
End Sub

' Сдвиги X/Y меняют моменты от N независимо и по утвержденному внутреннему
' знаку. Противоположный сдвиг и эквивалентные INPUT м сохраняют равновесие;
' ожидания получены из N*offset, а не из getters результата переноса.
Private Sub CheckReferenceOffsets(ByRef stats As TGeneralConfigStats, ByVal system As Object, ByVal unitRange As Object)
    Dim section As CSectionModel, batch As CBatchSectionCalculator, point As CSectionStateResult
    Dim baseline As CSectionStateResult, baselineSolves As Long
    Dim axis As Variant, value As Variant, mx As Double, my As Double, prefix As String
    Set section = BuildModel(stats)
    For Each axis In Array("X", "Y")
        For Each value In Array(-10#, 10#)
            SetValue system, "Load.ReferenceOffset" & CStr(axis), CDbl(value)
            Set batch = RunGeneralBatch(stats, section, -10000#, 200000#, 100000#, "PR1")
            Set point = batch.ResultAt(1).StrengthResult.DirectState.StateResult
            mx = 200000#: my = 100000#
            If CStr(axis) = "X" Then my = my - 10000# * CDbl(value)
            If CStr(axis) = "Y" Then mx = mx - 10000# * CDbl(value)
            prefix = "generalGeometry.Load.ReferenceOffset" & CStr(axis) & "." & NumberText(CDbl(value))
            CheckClose stats, prefix & ".mx", point.Mxint, mx, 1#
            CheckClose stats, prefix & ".my", point.Myint, my, 1#
            Check stats, prefix & ".converged", point.Converged
            CheckClose stats, prefix & ".n", point.Nint, -10000#, 0.01
        Next value
        SetValue system, "Load.ReferenceOffset" & CStr(axis), 0#
    Next axis
    SetValue unitRange, "Length", "m"
    SetValue system, "Load.ReferenceOffsetX", 0.01: SetValue system, "Load.ReferenceOffsetY", -0.01
    Set batch = RunGeneralBatch(stats, section, -10000#, 200000#, 100000#, "PR1")
    Set point = batch.ResultAt(1).StrengthResult.DirectState.StateResult
    CheckClose stats, "generalGeometry.Load.ReferenceOffsetX.units", point.Myint, 0#, 1#
    CheckClose stats, "generalGeometry.Load.ReferenceOffsetY.units", point.Mxint, 300000#, 1#
    SetValue unitRange, "Length", "mm"
    SetValue system, "Load.ReferenceOffsetX", 0#: SetValue system, "Load.ReferenceOffsetY", 0#
    Set batch = RunGeneralBatch(stats, section, -10000#, 200000#, 100000#, "PR1")
    Set baseline = batch.ResultAt(1).StrengthResult.DirectState.StateResult
    SetValue system, "Load.ReferenceOffsetX", 10#: SetValue system, "Load.ReferenceOffsetY", -10#
    Set batch = RunGeneralBatch(stats, section, -10000#, 100000#, 200000#, "PR1")
    Set point = batch.ResultAt(1).StrengthResult.DirectState.StateResult
    CheckClose stats, "generalGeometry.Load.ReferenceOffsetX.equivalentPlane", point.KappaY, baseline.KappaY, 0.000000000001
    CheckClose stats, "generalGeometry.Load.ReferenceOffsetY.equivalentPlane", point.KappaX, baseline.KappaX, 0.000000000001
    CheckClose stats, "generalGeometry.Load.ReferenceOffset.equivalentEps", point.Epsilon0, baseline.Epsilon0, 0.0000000001
    SetValue system, "Load.ReferenceOffsetX", 0#: SetValue system, "Load.ReferenceOffsetY", 0#
    Set batch = RunGeneralBatch(stats, section, 0#, 20000#, 10000#, "PR1")
    Set baseline = batch.ResultAt(1).StrengthResult.DirectState.StateResult: baselineSolves = batch.SolverCallCount
    SetValue system, "Load.ReferenceOffsetX", 10#: SetValue system, "Load.ReferenceOffsetY", -10#
    Set batch = RunGeneralBatch(stats, section, 0#, 20000#, 10000#, "PR1")
    Set point = batch.ResultAt(1).StrengthResult.DirectState.StateResult
    Check stats, "generalGeometry.Load.ReferenceOffsetX.inactiveNoN", point.Converged
    CheckClose stats, "generalGeometry.Load.ReferenceOffsetX.inactivePlane", point.KappaY, baseline.KappaY, 0.000000000001
    CheckClose stats, "generalGeometry.Load.ReferenceOffsetY.inactivePlane", point.KappaX, baseline.KappaX, 0.000000000001
    Check stats, "generalGeometry.Load.ReferenceOffset.inactiveSolveCount", batch.SolverCallCount = baselineSolves
    SetValue system, "Load.ReferenceOffsetX", 0#: SetValue system, "Load.ReferenceOffsetY", 0#
End Sub

' Глобальный Extension проверяется на текущем Strength/Cracked State и
' отдельной физической Capacity. Перегрузка не должна превращаться в OK
' прочности; техническая диаграмма помогает найти равновесие, не новый предел.
Private Sub CheckExtensionConsumer(ByRef stats As TGeneralConfigStats, ByVal system As Object, ByVal profiles As Object)
    Dim section As CSectionModel, batch As CBatchSectionCalculator, normal As CSectionStateResult
    Dim point As CSectionStateResult, profile As Variant, enabled As Variant, prefix As String
    Set section = BuildModel(stats)
    For Each profile In Array("PR1", "PR2")
        For Each enabled In Array("No", "Yes")
            SetValue system, "General.DiagramExtension", CStr(enabled)
            prefix = "generalGeometry.General.DiagramExtension." & CStr(profile) & "." & CStr(enabled)
            Set batch = RunGeneralBatch(stats, section, -10000#, 200000#, 100000#, CStr(profile))
            Set point = CurrentPoint(batch, CStr(profile))
            Check stats, prefix & ".normalConverged", Not point Is Nothing
            If Not point Is Nothing Then
                Check stats, prefix & ".normalPhysical", point.Converged And point.WithinPhysicalRange
                Check stats, prefix & ".normalNoExtension", Not point.ExtensionUsed
                If CStr(enabled) = "No" Then
                    Set normal = point
                Else
                    CheckClose stats, prefix & ".normalEps", point.Epsilon0, normal.Epsilon0, 0.0000000001
                    CheckClose stats, prefix & ".normalKx", point.KappaX, normal.KappaX, 0.000000000001
                    CheckClose stats, prefix & ".normalKy", point.KappaY, normal.KappaY, 0.000000000001
                End If
            End If
            Set batch = RunGeneralBatch(stats, section, -10000000#, 0#, 0#, CStr(profile))
            Set point = CurrentPoint(batch, CStr(profile))
            If CStr(enabled) = "Yes" Then
                Check stats, prefix & ".overloadState", Not point Is Nothing
                If Not point Is Nothing Then
                    Check stats, prefix & ".overloadConverged", point.Converged
                    Check stats, prefix & ".overloadExtended", point.ExtensionUsed And Not point.WithinPhysicalRange
                End If
                If CStr(profile) = "PR1" Then
                    Check stats, prefix & ".overloadFail", batch.ResultAt(1).DirectStateMeta.InternalStatus = rsCheckFailed
                Else
                    Check stats, prefix & ".overloadFail", batch.ResultAt(1).CrackCurrentStateMeta.InternalStatus = rsCheckFailed
                End If
            Else
                If CStr(profile) = "PR1" Then
                    Check stats, prefix & ".offNumFail", batch.ResultAt(1).DirectStateMeta.InternalStatus = rsNumericalFailure
                Else
                    Check stats, prefix & ".offNumFail", batch.ResultAt(1).CrackCurrentStateMeta.InternalStatus = rsNumericalFailure
                End If
            End If
        Next enabled
    Next profile
    SetProfile profiles, "Calculation.Strength.Capacity", "Yes", "PR1"
    Dim capacity As Double
    For Each enabled In Array("No", "Yes")
        SetValue system, "General.DiagramExtension", CStr(enabled)
        Set batch = RunGeneralBatch(stats, section, -10000#, 200000#, 100000#, "PR1")
        Set point = batch.ResultAt(1).StrengthResult.Capacity.StateResult
        Check stats, "generalGeometry.General.DiagramExtension.capacityPhysical." & CStr(enabled), point.Converged And point.WithinPhysicalRange And Not point.ExtensionUsed
        If CStr(enabled) = "No" Then
            capacity = batch.ResultAt(1).StrengthResult.Capacity.LambdaCapacity
        Else
            CheckClose stats, "generalGeometry.General.DiagramExtension.capacityInvariant", batch.ResultAt(1).StrengthResult.Capacity.LambdaCapacity, capacity, 0.0000001
        End If
    Next enabled
    SetProfile profiles, "Calculation.Strength.Capacity", "No", "PR1"
    SetValue system, "General.DiagramExtension", "No"
End Sub

' Невалидные значения и потерянные строки передаются именно владельцу поля.
' Диагностика обязана содержать текущий адрес/ключ и способ исправления;
' возврат defaults или нечисленный отказ считается отдельным падением.
Private Sub CheckInvalidGeneralFields(ByRef stats As TGeneralConfigStats, ByVal system As Object, ByVal position As Long)
    Dim keys As Variant, key As Variant, values As Variant, value As Variant, saved As Variant
    Dim target As Object, reason As String, prefix As String, index As Long
    keys = Array("General.DiagramExtension", "Calculation.ZeroMomentPerDepth", "Geometry.Source", "Geometry.Type", _
        "Mesh.StepX", "Mesh.StepY", "Mesh.BoundarySubdivisions", "Load.ReferenceOffsetX", "Load.ReferenceOffsetY")
    saved = system.Formula
    For Each key In keys
        values = Array(vbNullString, "TODO", "abc", CVErr(2015))
        If CStr(key) = "Mesh.StepX" Or CStr(key) = "Mesh.StepY" Then values = Array(vbNullString, "TODO", "abc", CVErr(2015), 0#, -1#)
        If CStr(key) = "Mesh.BoundarySubdivisions" Then values = Array(vbNullString, "TODO", "abc", CVErr(2015), 0#, -1#, 1.5)
        If CStr(key) = "Calculation.ZeroMomentPerDepth" Then values = Array(vbNullString, "TODO", "abc", CVErr(2015), -1#)
        index = 0
        For Each value In values
            system.Formula = saved: Set target = CellForKey(system, CStr(key), 2): target.Value2 = value
            prefix = "generalGeometry.invalid." & CStr(position) & "." & CStr(key) & "." & CStr(index)
            reason = ConsumeError(stats, CStr(key))
            Check stats, prefix & ".rejected", Len(reason) > 0
            Check stats, prefix & ".key", InStr(1, reason, CStr(key), vbTextCompare) > 0
            Check stats, prefix & ".address", InStr(1, reason, target.Address(False, False), vbTextCompare) > 0
            Check stats, prefix & ".action", InStr(1, reason, "введите", vbTextCompare) > 0 Or _
                InStr(1, reason, "выберите", vbTextCompare) > 0 Or InStr(1, reason, "исправ", vbTextCompare) > 0
            LogLine stats, "GENERAL_INPUT: " & prefix & "|" & reason
            index = index + 1
        Next value
        system.Formula = saved: Set target = CellForKey(system, CStr(key), 1): target.Value2 = "Missing." & CStr(key)
        reason = ConsumeError(stats, CStr(key))
        prefix = "generalGeometry.missing." & CStr(position) & "." & CStr(key)
        Check stats, prefix & ".rejected", Len(reason) > 0
        Check stats, prefix & ".key", InStr(1, reason, CStr(key), vbTextCompare) > 0
        system.Formula = saved
        reason = ConsumeError(stats, CStr(key))
        Check stats, "generalGeometry.recovery." & CStr(position) & "." & CStr(key), Len(reason) = 0
    Next key
    system.Formula = saved
End Sub

' Запускает нужный рабочий consumer без перехвата ошибки как NumFail.
' Batch-проверка читает typed meta, поскольку ApplySettings хранит InputErr.
Private Function ConsumeError(ByRef stats As TGeneralConfigStats, ByVal key As String) As String
    Dim settings As CSystemSettingsReader, units As CUnitSystem, model As CSectionModel
    Dim provider As CMaterialModelProvider, batch As CBatchSectionCalculator, profiles As CCalculationProfileCatalog
    stats.Cases = stats.Cases + 1
    On Error GoTo Rejected
    LoadReader settings, units
    If key = "General.DiagramExtension" Then
        Set provider = New CMaterialModelProvider: provider.Initialize settings, units
    ElseIf Left$(key, 9) = "Geometry." Or Left$(key, 5) = "Mesh." Then
        Set model = BuildWorkbookSectionModel(ThisWorkbook, settings, units)
    Else
        Set model = New CSectionModel
        model.AddConcreteElement 0#, 0#, 40000#, 1, shapeType:="Rectangle", width:=200#, height:=200#
        model.AddRebarElement 60#, 60#, 16#, 201#, "A400", 2
        Set provider = New CMaterialModelProvider: provider.Initialize settings, units
        Set batch = New CBatchSectionCalculator: batch.Initialize model, provider
        If key = "Calculation.ZeroMomentPerDepth" Then
            Set profiles = New CCalculationProfileCatalog: profiles.LoadFromWorkbook ThisWorkbook
            Set batch.ProfileCatalog = profiles
            batch.ApplySettings settings, units
            batch.AddCombination "INVALID_GENERAL", -10000#, 0#, 0#, "PR1", "", "Auto"
            batch.Execute
            If batch.ResultAt(1).WorkflowMeta.InternalStatus = rsInvalidInput Then ConsumeError = batch.ResultAt(1).WorkflowMeta.ResultComment
        Else
            Audit03ApplyLoadReferenceForTests model, settings, units, batch
        End If
    End If
    Exit Function
Rejected:
    ConsumeError = Err.Description
End Function

' Выполняет действительный batch с актуальными Config и профилями. Общий
' helper проверяет result-subtree, comments и все подробные writers.
Private Function RunGeneralBatch(ByRef stats As TGeneralConfigStats, ByVal section As CSectionModel, _
        ByVal nValue As Double, ByVal mxValue As Double, ByVal myValue As Double, ByVal profileId As String) As CBatchSectionCalculator
    Dim settings As CSystemSettingsReader, units As CUnitSystem, provider As CMaterialModelProvider
    Dim profiles As CCalculationProfileCatalog, batch As CBatchSectionCalculator, passed As Long, failed As Long
    LoadReader settings, units
    Set provider = New CMaterialModelProvider: provider.Initialize settings, units
    Set profiles = New CCalculationProfileCatalog: profiles.LoadFromWorkbook ThisWorkbook
    Set batch = New CBatchSectionCalculator: batch.Initialize section, provider
    Set batch.ProfileCatalog = profiles: batch.ApplySettings settings, units
    batch.AddCombination "GENERAL_CONFIG", nValue, mxValue, myValue, profileId, "Общие настройки Config", "Auto"
    Audit03ApplyLoadReferenceForTests section, settings, units, batch
    batch.Execute
    stats.Cases = stats.Cases + 1
    stats.Report = stats.Report & Audit03ValidateConfigBatchResults(batch, units, passed, failed)
    stats.Passed = stats.Passed + passed: stats.Failed = stats.Failed + failed
    LogLine stats, "GENERAL_BATCH: profile=" & profileId & "|N=" & NumberText(nValue) & "|Mx=" & NumberText(mxValue) & "|My=" & NumberText(myValue) & _
        "|status=" & batch.ResultAt(1).Status & "|solves=" & CStr(batch.SolverCallCount)
    Set RunGeneralBatch = batch
End Function

' Получает конечное именованное НДС; failed attempt не берется из repository
' как reusable state. Strength сохраняет собственную failure meta отдельно.
Private Function CurrentPoint(ByVal batch As CBatchSectionCalculator, ByVal profileId As String) As CSectionStateResult
    If profileId = "PR1" Then
        Set CurrentPoint = batch.ResultAt(1).StrengthResult.DirectState.StateResult
    Else
        Set CurrentPoint = batch.ResultAt(1).StateRepository.FindState(sstCrackedState)
    End If
End Function

' Строит модель актуальным Excel-entrypoint, не локальным geometry fixture.
Private Function BuildModel(ByRef stats As TGeneralConfigStats) As CSectionModel
    Dim settings As CSystemSettingsReader, units As CUnitSystem
    LoadReader settings, units
    Set BuildModel = BuildWorkbookSectionModel(ThisWorkbook, settings, units)
    stats.Cases = stats.Cases + 1
End Function

' Читает настоящие именованные таблицы и единицы единственным adapter-ом.
Private Sub LoadReader(ByRef settings As CSystemSettingsReader, ByRef units As CUnitSystem)
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
End Sub

' Независимый суммарный контроль площадей построенных бетонных элементов.
Private Function ConcreteArea(ByVal model As CSectionModel) As Double
    Dim i As Long
    For i = 1 To model.ConcreteCount: ConcreteArea = ConcreteArea + model.ConcreteArea(i): Next i
End Function

' Ищет полную внутреннюю ячейку с заданной площадью, не граничную долю.
Private Function HasCellArea(ByVal model As CSectionModel, ByVal area As Double) As Boolean
    Dim i As Long
    For i = 1 To model.ConcreteCount
        If Abs(model.ConcreteArea(i) - area) < 0.00000001 Then HasCellArea = True: Exit Function
    Next i
End Function

' Находит параметр по ключу в реальной таблице; ошибка самого fixture явная.
Private Function CellForKey(ByVal table As Object, ByVal key As String, ByVal column As Long) As Object
    Dim row As Long
    For row = 2 To table.Rows.Count
        If CStr(table.Cells(row, 1).Value2) = key Then Set CellForKey = table.Cells(row, column): Exit Function
    Next row
    Err.Raise vbObjectError + 4499, "CellForKey", "Не найден ключ Config " & key
End Function

' Меняет только значение найденного поля; восстановление сохраняет Formula.
Private Sub SetValue(ByVal table As Object, ByVal key As String, ByVal value As Variant)
    CellForKey(table, key, 2).Value2 = value
End Sub

' Выбирает столбец профиля по ID и строку по ключу во втором столбце.
Private Sub SetProfile(ByVal table As Object, ByVal key As String, ByVal value As String, ByVal profileId As String)
    Dim row As Long, column As Long, targetRow As Long, targetColumn As Long
    For row = 1 To table.Rows.Count
        If CStr(table.Cells(row, 2).Value2) = key Then targetRow = row
        For column = 3 To table.Columns.Count
            If CStr(table.Cells(row, column).Value2) = profileId Then targetColumn = column
        Next column
    Next row
    If targetRow = 0 Or targetColumn = 0 Then Err.Raise vbObjectError + 4499, "SetProfile", "Не найден параметр профиля " & key
    table.Cells(targetRow, targetColumn).Value2 = value
End Sub

' Учет assertions не прекращает набор после первого отрицательного результата.
Private Sub Check(ByRef stats As TGeneralConfigStats, ByVal name As String, ByVal condition As Boolean)
    If condition Then
        stats.Passed = stats.Passed + 1: LogLine stats, "OK: " & name
    Else
        stats.Failed = stats.Failed + 1: LogLine stats, "FAIL: " & name
    End If
End Sub

' Сохраняет actual/expected/tolerance до проверки, без округления допусков.
Private Sub CheckClose(ByRef stats As TGeneralConfigStats, ByVal name As String, ByVal actual As Double, ByVal expected As Double, ByVal tolerance As Double)
    LogLine stats, "GENERAL_NUMBER: " & name & "|actual=" & NumberText(actual) & "|expected=" & NumberText(expected) & "|tolerance=" & NumberText(tolerance)
    Check stats, name, Abs(actual - expected) <= tolerance
End Sub

' Формат журнала не зависит от десятичного разделителя текущего Windows.
Private Function NumberText(ByVal value As Double) As String
    NumberText = Replace(CStr(value), ",", ".")
End Function

' Отчет и живой progress используют общий принятый путь тестового runner-а.
Private Sub LogLine(ByRef stats As TGeneralConfigStats, ByVal line As String)
    stats.Report = stats.Report & line & vbCrLf
    Dim fileNumber As Integer: fileNumber = FreeFile
    Open ThisWorkbook.Path & Application.PathSeparator & "Audit03_Search_Progress.txt" For Output As #fileNumber
    Print #fileNumber, stats.Report;
    Close #fileNumber
End Sub
