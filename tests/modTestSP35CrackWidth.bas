Attribute VB_Name = "modTestSP35CrackWidth"
Option Explicit

' ==========================================================================
' Адресные тесты ширины нормальных трещин СП 35
' ==========================================================================
' Проверяют нормативные численные формулы и явные внутренние единицы.
' Эталоны получены из пособия (4.306)-(4.310) и таблицы 7.26 СП 35.
' Формульные и подготовительные сценарии принимают готовые плоскости НДС.
' Отдельные интеграционные сценарии запускают штатный batch в копии книги;
' статусы не подменяются, результаты сопоставляются с независимыми эталонами.

Private mPassed As Long
Private mFailed As Long
Private mReport As String

' ДЛЯ ТЕСТОВ: фиксирует результат сравнения без зависимости от Excel.
Private Sub Check(ByVal name As String, ByVal condition As Boolean)
    If condition Then
        mPassed = mPassed + 1: mReport = mReport & "OK: " & name & vbCrLf
    Else
        mFailed = mFailed + 1: mReport = mReport & "FAIL: " & name & vbCrLf
    End If
End Sub

' ДЛЯ ТЕСТОВ: сравнивает формулу с независимо заданным численным эталоном.
Private Sub CheckNear(ByVal name As String, ByVal actual As Double, ByVal expected As Double)
    Check name & "; actual=" & CStr(actual) & "; expected=" & CStr(expected), Abs(actual - expected) <= 0.000000000001
End Sub

' ДЛЯ ТЕСТОВ: материалы и спецификация не зависят от пользовательской книги.
' Текущие напряжения арматуры ниже площадки: независимый эталон sigma=Es*eps.
Private Function PipelineProvider() As CMaterialModelProvider
    Dim concrete As CConcreteMaterialParameters, steel As CSteelMaterialParameters, provider As CMaterialModelProvider
    Set concrete = New CConcreteMaterialParameters: concrete.Initialize 15.5, 1.1, 22#, 1.8, 32500#, 32500#, rbMc2:=14.6
    Set steel = New CSteelMaterialParameters: steel.Initialize 350#, 350#, 390#, 390#, 200000#, 200000#
    Set provider = New CMaterialModelProvider: provider.InitializeFromParameters concrete, steel
    Set PipelineProvider = provider
End Function

' ДЛЯ ТЕСТОВ: использует ту же спецификацию cracked-НДС, что штатный pipeline.
Private Function PipelineSpec() As CMaterialModelSpec
    Dim spec As CMaterialModelSpec: Set spec = New CMaterialModelSpec
    spec.Initialize "SLS(II)", "TwoLine", "Ignore", "TwoLine"
    Set PipelineSpec = spec
End Function

' ДЛЯ ТЕСТОВ: заданная плоскость лишь оценивается и подтверждается для своих
' усилий. Здесь нет поиска или итераций; Width обязан принять готовый State.
Private Function ReadyPlane(ByVal section As CSectionModel, ByVal provider As CMaterialModelProvider, _
        ByVal epsilon0 As Double, ByVal kappaX As Double, ByVal kappaY As Double) As CSectionStateResult
    Dim solver As CSectionSolver, result As CSectionStateResult, spec As CMaterialModelSpec
    Set spec = PipelineSpec: Set solver = New CSectionSolver
    solver.EvaluateStrainPlane section, provider.ConcreteMaterialFromSpec(spec), provider.SteelMaterialFromSpec(spec), epsilon0, kappaX, kappaY
    Check "SP35.readyPlane.confirmed", solver.ConfirmEquilibrium(solver.Nint, solver.Mxint, solver.Myint)
    Set result = New CSectionStateResult: result.InitializeFromSolver sstCrackedState, cpCrackedNDS, spec, solver, False, True, 0
    Set ReadyPlane = result
End Function

' ДЛЯ ТЕСТОВ: отсутствие Formation с машинной численной причиной должно
' оставаться отдельным результатом; Width СП 35 не получает NumFail/psi_s.
Private Function MissingFormation() As CCrackFormationResult
    Dim owner As CCrackFormationCalculator, meta As CResultMeta, result As CCrackFormationResult
    Set owner = New CCrackFormationCalculator: owner.LimitSearchSetFailure sfcNumericalFailure, "Контрольная причина отсутствия точки образования."
    Set meta = New CResultMeta: meta.SetSolverFailure sfcNumericalFailure, rkCrackFormation, "Контрольная причина отсутствия точки образования.", True
    Set result = New CCrackFormationResult: result.InitializeFromCalculator owner, meta
    Set MissingFormation = result
End Function

' ДЛЯ ТЕСТОВ: критический якорь выбирается по ширине, а не по максимальному
' напряжению. Группа 2 с меньшим sigma дает большее a_cr из-за своей зоны.
Private Sub TestReadyStatePipeline()
    Dim section As CSectionModel, provider As CMaterialModelProvider, state As CSectionStateResult
    Dim formation As CCrackFormationResult, load As CSectionLoadState, width As CCrackWidthCalculator
    Set section = RectangularSection: AddBar section, -190#, 390#, 20#: AddBar section, 100#, 300#, 10#
    Set provider = PipelineProvider: Set state = ReadyPlane(section, provider, -0.0002, 0.000002, 0#)
    Set formation = MissingFormation: Set load = New CSectionLoadState
    load.Initialize state.TargetN, state.TargetMx, state.TargetMy, 0#, 0#
    Set width = New CCrackWidthCalculator: width.StandardCode = "SP35": width.AllowableCrackWidth = 1#
    width.Calculate state, section, provider, PipelineSpec(), formation, load
    Check "SP35.pipeline.calculated", width.ResultMeta.Calculated And width.ResultMeta.InternalStatus = rsSuccess
    CheckNear "SP35.pipeline.criticalWidth", width.CrackWidth, 80# / 200000# * 1.5 * Sqr(192#) * 10#
    CheckNear "SP35.pipeline.criticalSigmaNotMaximum", width.SigmaS, 80#
    CheckNear "SP35.pipeline.criticalArea", width.Abt, 19200#
    CheckNear "SP35.pipeline.criticalRr", width.SP35ReinforcementRadius, 1920#
    Check "SP35.pipeline.criticalCandidate", width.SP35CriticalCandidate = 2
    Check "SP35.pipeline.actualBar", width.TensionRebarIds = "R2" And width.TensionRebarCount = 1
    Check "SP35.pipeline.noPsiSFallbackWarning", InStr(1, width.ResultMeta.ResultComment, "psi_s", vbTextCompare) = 0 And _
        InStr(1, width.ResultMeta.ResultComment, "Предупреждение", vbTextCompare) = 0
    Check "SP35.pipeline.formationCausePreserved", formation.ResultMeta.InternalStatus = rsNumericalFailure
    Check "SP35.pipeline.noSolverCalls", state.SolverCallCount = 0

    Dim result As CCrackWidthResult, region As CConcreteRegion, data As CSP35CrackData
    Set result = New CCrackWidthResult: result.InitializeFromCalculator width, width.ResultMeta, "SP35_TEST"
    Set region = result.InteractionRegion: Set data = result.SP35Data
    Check "SP35.result.owner", region.OwnerID = "SP35_TEST" And region.Standard = "SP35" And region.AnchorID = "G2"
    Check "SP35.result.criticalAnchor", result.CriticalAnchorGroupID = "G2" And result.StandardCode = "SP35"
    Check "SP35.result.revision", data.SectionRevision = section.Revision
    CheckNear "SP35.result.sameArea", region.Area, result.Abt
    Check "SP35.result.fiveBoundaryProbes", region.BoundaryProbeCount = 5
    Dim probes As Variant, savedProbe As Double
    probes = region.BoundaryProbes: savedProbe = CDbl(probes(3, 4)): probes(3, 4) = savedProbe + 1000#
    probes = result.InteractionRegion.BoundaryProbes
    CheckNear "SP35.result.probeSnapshotIndependent", CDbl(probes(3, 4)), savedProbe
    CheckNear "SP35.result.allCandidateWidths", data.CandidateCrackWidth(1), 116# / 200000# * 1.5 * Sqr(84.5) * 10#
    Dim number As Long
    On Error Resume Next
    data.Prepare section, 0.0005, 0#, 0#, LinearSteel(), 10#, 50#, "Max", "6d", True
    number = Err.Number: Err.Clear
    On Error GoTo 0
    Check "SP35.result.dataFrozen", number = vbObjectError + 5511
    number = 0
    On Error Resume Next
    data.StoreFormulaResult 1, 1#, 1#, 1#
    number = Err.Number: Err.Clear
    On Error GoTo 0
    Check "SP35.result.formulaDataFrozen", number = vbObjectError + 5511

    width.AllowableCrackWidth = 0.01
    width.Calculate state, section, provider, PipelineSpec(), formation, load
    Check "SP35.pipeline.exceeded", width.ResultMeta.InternalStatus = rsCheckFailed And width.ResultMeta.ResultCode = rcCrackWidthExceeded
    Check "SP35.pipeline.failureComment", InStr(1, width.ResultMeta.ResultComment, "превышает", vbTextCompare) > 0
    CheckNear "SP35.result.notChangedByNewCalculation", result.CrackWidth, 80# / 200000# * 1.5 * Sqr(192#) * 10#
    width.AllowableCrackWidth = 1#: width.SP35RebarProfile = "Smooth"
    width.Calculate state, section, provider, PipelineSpec(), formation, load
    CheckNear "SP35.pipeline.smooth", width.CrackWidth, 0.2688
    width.Calculate Nothing, section, provider, PipelineSpec(), formation, load
    Check "SP35.pipeline.missingStateBlocked", width.ResultMeta.InternalStatus = rsBlockedByDependency And Not width.ResultMeta.Calculated
    Check "SP35.pipeline.noOldRegion", width.InteractionRegion Is Nothing And width.SP35Data Is Nothing
End Sub

' ДЛЯ ТЕСТОВ: отсутствие растяжения не вызывает формулу; чистое растяжение
' использует всю чистую площадь с несколькими opening, без искусственной Н.О.
Private Sub TestCentralAndCompressedPipeline()
    Dim section As CSectionModel, provider As CMaterialModelProvider, state As CSectionStateResult
    Dim width As CCrackWidthCalculator, load As CSectionLoadState, formation As CCrackFormationResult
    Set section = RectangularSection: AddBar section, 0#, 350#, 20#: AddBar section, 0#, 100#, 20#
    AddOpening section, "A", -60#, 190#, 60#, 250#
    AddOpening section, "B", 80#, 300#, 100#, 320#
    Set provider = PipelineProvider: Set formation = MissingFormation: Set load = New CSectionLoadState
    Set width = New CCrackWidthCalculator: width.StandardCode = "SP35": width.AllowableCrackWidth = 1#
    Set state = ReadyPlane(section, provider, 0.0005, 0#, 0#)
    load.Initialize state.TargetN, state.TargetMx, state.TargetMy, 0#, 0#
    width.Calculate state, section, provider, PipelineSpec(), formation, load
    Check "SP35.central.calculated", width.ResultMeta.Calculated And width.CentralTensionBranch
    CheckNear "SP35.central.twoOpenings.netArea", width.Abt, 152400#
    CheckNear "SP35.central.actualDiaSum", width.SP35Data.CandidateBetaDiameterSum(1), 40#
    CheckNear "SP35.central.width", width.CrackWidth, 100# / 200000# * 1.5 * Sqr(381#) * 10#
    Set state = ReadyPlane(section, provider, -0.0005, 0#, 0#)
    load.Initialize state.TargetN, state.TargetMx, state.TargetMy, 0#, 0#
    width.Calculate state, section, provider, PipelineSpec(), formation, load
    Check "SP35.compressed.notApplicable", width.ResultMeta.InternalStatus = rsNotApplicable And Not width.ResultMeta.Calculated
    Check "SP35.compressed.noOldArea", width.InteractionRegion Is Nothing
End Sub

' ДЛЯ ТЕСТОВ: запускает Width штатным API, не численный поиск Formation.
' Независимые эталоны проверяют критический выбор, несколько пустот и статусы.
Public Function RunSP35PipelineTests() As String
    On Error GoTo Failed
    mPassed = 0: mFailed = 0: mReport = vbNullString
    TestReadyStatePipeline
    TestCentralAndCompressedPipeline
    GoTo Finished
Failed:
    Check "runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    mReport = mReport & "TOTAL_SP35_PIPELINE: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
    RunSP35PipelineTests = mReport
End Function

' ДЛЯ ТЕСТОВ: меняет конкретную настройку в собственной временной таблице.
' Вся таблица Config пользователя остается нетронутой; имя диапазона возвращается.
Private Sub SetWorkbookTestSetting(ByVal source As Object, ByVal key As String, ByVal value As Variant)
    Dim i As Long, emptyRow As Long
    For i = 2 To source.Rows.Count
        If CStr(source.Cells(i, 1).Value2) = key Then source.Cells(i, 2).Value2 = value: Exit Sub
        If emptyRow = 0 And Len(CStr(source.Cells(i, 1).Value2)) = 0 Then emptyRow = i
    Next i
    If emptyRow = 0 Then Err.Raise vbObjectError + 5520, "SP35 tests", "Нет места для тестовой настройки " & key
    source.Cells(emptyRow, 1).Value2 = key: source.Cells(emptyRow, 2).Value2 = value
End Sub

' ДЛЯ ТЕСТОВ: сетка и точный material-contour содержат две независимые пустоты.
' Чистая площадь известна независимо: 400*400 - 120*60 - 20*20 = 152400 мм2.
Private Function WorkbookTestSection() As CSectionModel
    Dim section As CSectionModel, i As Long, j As Long, x As Double, y As Double
    Set section = New CSectionModel: section.SourceType = "SP35WorkbookTest"
    For i = 0 To 19
        For j = 0 To 19
            x = -190# + 20# * i: y = 10# + 20# * j
            If Not (x > -60# And x < 60# And y > 190# And y < 250#) And _
                    Not (x > 80# And x < 100# And y > 300# And y < 320#) Then
                section.AddConcreteElement x, y, 400#, 1, , , "Rectangle", 20#, 20#, 0#, , 400# * 400# / 12#, 400# * 400# / 12#, 0#, True
            End If
        Next j
    Next i
    section.Contours.AddContourLine "CONTOUR_OUTER_1", -200#, 0#, 200#, 0#
    section.Contours.AddContourLine "CONTOUR_OUTER_2", 200#, 0#, 200#, 400#
    section.Contours.AddContourLine "CONTOUR_OUTER_3", 200#, 400#, -200#, 400#
    section.Contours.AddContourLine "CONTOUR_OUTER_4", -200#, 400#, -200#, 0#
    AddOpening section, "A", -60#, 190#, 60#, 250#
    AddOpening section, "B", 80#, 300#, 100#, 320#
    AddBar section, -100#, 100#, 25#: AddBar section, 100#, 100#, 25#
    AddBar section, -100#, 300#, 25#: AddBar section, 100#, 300#, 25#
    Set WorkbookTestSection = section
End Function

' ДЛЯ ТЕСТОВ: штатный batch решает два сочетания, writer пишет всю область и
' новую относительную шапку. Настоящий AutoCAD получает отдельные outer/opening
' из сохраненных в метрах Results; переключатели не вызывают повторный solve.
' ДЛЯ ТЕСТОВ: читает только annotation-блок по якорю и текущей схеме,
' не захватывая соседние таблицы через объединенные строки заголовков.
Private Function ReadSP35AnnotationTable(ByVal anchor As Object) As Variant
    Dim rows As Long
    rows = 1
    Do While Len(CStr(anchor.Offset(rows, 0).Value2)) > 0
        rows = rows + 1
    Loop
    ReadSP35AnnotationTable = anchor.Resize(rows, 22).Value2
End Function

' ДЛЯ ТЕСТОВ: проверяет штатный batch, относительный writer и реальный CAD export.
Public Function RunSP35WorkbookTests() As String
    Dim sheet As Object, settingsRange As Object, originalSettings As String, originalAnchor As String
    Dim values As Variant, expanded() As Variant, i As Long, j As Long, number As Long, description As String
    Dim section As CSectionModel, provider As CMaterialModelProvider, settings As CSystemSettingsReader
    Dim profiles As CCalculationProfileCatalog, profile As CCalculationProfile, batch As CBatchSectionCalculator
    Dim state As CSectionStateResult, width As CCrackWidthResult, units As CUnitSystem
    Dim writer As CNDMResultsWriter, summary As CCrackSummaryWriter, anchor As Object, annotations As Object
    Dim acad As Object, doc As Object, previousDoc As Object, entity As Object, countBefore As Long, totalArea As Double, maxArea As Double
    On Error GoTo Failed
    mPassed = 0: mFailed = 0: mReport = vbNullString
    originalSettings = ThisWorkbook.Names.Item("rngSystemSettings").RefersTo
    originalAnchor = ThisWorkbook.Names.Item("rngCrackSummaryAnchor").RefersTo
    values = ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange.Value2
    ReDim expanded(1 To UBound(values, 1) + 6, 1 To UBound(values, 2))
    For i = 1 To UBound(values, 1): For j = 1 To UBound(values, 2): expanded(i, j) = values(i, j): Next j: Next i
    Set sheet = ThisWorkbook.Worksheets.Add: sheet.Name = "SP35_TEST_SNAPSHOT"
    Set settingsRange = sheet.Cells(100, 1).Resize(UBound(expanded, 1), UBound(expanded, 2))
    settingsRange.Value2 = expanded
    ThisWorkbook.Names.Item("rngSystemSettings").RefersTo = "='" & sheet.Name & "'!" & settingsRange.Address
    ThisWorkbook.Names.Item("rngCrackSummaryAnchor").RefersTo = "='" & sheet.Name & "'!$C$20"
    SetWorkbookTestSetting settingsRange, "SLS.Crack.Code", "SP35"
    SetWorkbookTestSetting settingsRange, "SLS.Crack.SP35.GroupGapTolerance", 10#
    SetWorkbookTestSetting settingsRange, "SLS.Crack.SP35.RowTolerance", 50#
    SetWorkbookTestSetting settingsRange, "SLS.Crack.SP35.NeighborRatioLimit", 0.2
    SetWorkbookTestSetting settingsRange, "SLS.Crack.SP35.RadiusDiameterMode", "Max"
    SetWorkbookTestSetting settingsRange, "SLS.Crack.SP35.InteractionRadiusMode", "6d"
    SetWorkbookTestSetting settingsRange, "SLS.Crack.SP35.RebarProfile", "Periodic"
    SetWorkbookTestSetting settingsRange, "SLS.Crack.Allowable", 0.25
    SetWorkbookTestSetting settingsRange, "Calculation.ZeroMomentPerDepth", 0#
    SetWorkbookTestSetting settingsRange, "Solver.ToleranceN", 0.1
    SetWorkbookTestSetting settingsRange, "Solver.ToleranceMx", 10#
    SetWorkbookTestSetting settingsRange, "Solver.ToleranceMy", 10#
    Set section = WorkbookTestSection: Set provider = PipelineProvider
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set profiles = New CCalculationProfileCatalog: profiles.LoadFromWorkbook ThisWorkbook
    Set profile = profiles.ProfileById("PR1")
    profile.Initialize "PR1", "SP35 test", "Два сочетания с несколькими проемами.", False, False, True, False
    profile.SetCrackedStateSpec PipelineSpec()
    Dim formationSpec As CMaterialModelSpec: Set formationSpec = New CMaterialModelSpec
    formationSpec.Initialize "SLS(II)", "TwoLine", "UseDiagram", "TwoLine"
    profile.SetCrackInitiationSpec formationSpec
    Set batch = New CBatchSectionCalculator: batch.Initialize section, provider
    Set batch.ProfileCatalog = profiles: batch.ApplySettings settings
    Set state = ReadyPlane(section, provider, 0.0012, 0#, 0#)
    batch.AddCombination "SP35_A", state.TargetN, state.TargetMx, state.TargetMy, "PR1", "Контроль A", "Auto"
    Set state = ReadyPlane(section, provider, 0.0015, 0#, 0#)
    batch.AddCombination "SP35_B", state.TargetN, state.TargetMx, state.TargetMy, "PR1", "Контроль B", "Auto"
    Set state = ReadyPlane(section, provider, -0.0002, 0.000006, 0#)
    batch.AddCombination "SP35_C", state.TargetN, state.TargetMx, state.TargetMy, "PR1", "Контроль изгиба", "Auto"
    batch.Execute
    For i = 1 To 2
        Set width = batch.ResultAt(i).CrackResult.Width
        Check "SP35.workbook.calculated." & CStr(i) & "; " & batch.ResultAt(i).OverallMeta.ResultComment, width.ResultMeta.Calculated
        Check "SP35.workbook.noNumFail." & CStr(i), width.ResultMeta.InternalStatus <> rsNumericalFailure
        If Not width.InteractionRegion Is Nothing Then
            CheckNear "SP35.workbook.twoOpeningsArea." & CStr(i), width.InteractionRegion.Area, 152400#
            Check "SP35.workbook.threeLoops." & CStr(i), width.InteractionRegion.LoopCount = 3
        Else
            Check "SP35.workbook.missingRegion." & CStr(i), False
        End If
    Next i
    Check "SP35.workbook.OK", batch.ResultAt(1).CrackWidthMeta.InternalStatus = rsSuccess
    Check "SP35.workbook.FAIL", batch.ResultAt(2).CrackWidthMeta.InternalStatus = rsCheckFailed
    Set units = New CUnitSystem: units.InitializeDefaults
    Set summary = New CCrackSummaryWriter: summary.WriteSummary ThisWorkbook, batch, units
    Set anchor = ThisWorkbook.Names.Item("rngCrackSummaryAnchor").RefersToRange
    Check "SP35.writer.relativeHeader", anchor.Offset(-4, 30).Value2 = "нормальные трещины по СП 35"
    CheckNear "SP35.writer.areaOutputM2", CDbl(anchor.Offset(0, 30).Value2), 0.1524
    Check "SP35.writer.widthMetrePrecision", CStr(anchor.Offset(0, 40).NumberFormat) = "0.000000" And CStr(anchor.Offset(0, 40).Text) <> "0.000000"
    Check "SP35.writer.allowableMetrePrecision", CStr(anchor.Offset(0, 41).NumberFormat) = "0.000000"
    Check "SP35.writer.dataNoWrap", Not CBool(anchor.Resize(3, 72).WrapText)
    CheckSP35HeaderLayout anchor, units
    Check "SP35.writer.allBars", InStr(1, CStr(anchor.Offset(0, 36).Value2), "R1", vbBinaryCompare) > 0 And _
        InStr(1, CStr(anchor.Offset(0, 36).Value2), "R4", vbBinaryCompare) > 0
    Check "SP35.writer.inactiveSP63", anchor.Offset(0, 66).Value2 = "N/A"
    Check "SP35.writer.failureColor", anchor.Offset(1, 43).Interior.Color <> anchor.Offset(0, 43).Interior.Color
    Check "SP35.writer.emptySeparators", Len(CStr(anchor.Offset(0, 29).Value2)) = 0 And Len(CStr(anchor.Offset(0, 44).Value2)) = 0
    Set writer = New CNDMResultsWriter: writer.WriteResults ThisWorkbook, section, provider, batch, units
    Set annotations = ThisWorkbook.Names.Item("rngNDMSectionAnnotations").RefersToRange
    Check "SP35.annotation.ownerHeader", annotations.Offset(0, 13).Value2 = "CombinationID"
    Check "SP35.annotation.roleHeader", annotations.Offset(0, 17).Value2 = "LoopRole"
    Check "SP35.workbook.bendingCalculated; " & batch.ResultAt(3).OverallMeta.ResultComment, batch.ResultAt(3).CrackWidthMeta.Calculated
    values = ReadSP35AnnotationTable(annotations)
    Dim probeRows As Long
    For i = 2 To UBound(values, 1)
        If CStr(values(i, 2)) = "CRACK_PROBE" And CStr(values(i, 14)) = "SP35_C" Then probeRows = probeRows + 1
    Next i
    Check "SP35.annotation.fiveBoundaryProbes", probeRows = 5
    Set acad = ConnectToRunningAutoCAD()
    If acad.Documents.Count > 0 Then Set previousDoc = acad.ActiveDocument
    Set doc = acad.Documents.Add
    countBefore = doc.ModelSpace.Count
    Check "SP35.export.off", SP35ExportSavedCrackRegionForTests(doc, "SP35_A", "CrackedState", False) = 0 And doc.ModelSpace.Count = countBefore
    Check "SP35.export.strengthExcluded", SP35ExportSavedCrackRegionForTests(doc, "SP35_A", "StrengthState", True) = 0 And doc.ModelSpace.Count = countBefore
    Check "SP35.export.capacityExcluded", SP35ExportSavedCrackRegionForTests(doc, "SP35_A", "CapacityState", True) = 0 And doc.ModelSpace.Count = countBefore
    Check "SP35.export.missingLC", SP35ExportSavedCrackRegionForTests(doc, "MISSING", "CrackedState", True) = 0
    Check "SP35.export.threeLoops", SP35ExportSavedCrackRegionForTests(doc, "SP35_A", "CrackedState", True) = 3
    For i = countBefore To doc.ModelSpace.Count - 1
        Set entity = doc.ModelSpace.Item(i)
        Check "SP35.export.closed." & CStr(i - countBefore), CBool(entity.Closed)
        totalArea = totalArea + CDbl(entity.Area)
        If CDbl(entity.Area) > maxArea Then maxArea = CDbl(entity.Area)
    Next i
    Check "SP35.export.areaMM2; delta=" & CStr(2# * maxArea - totalArea - 152400#), _
        Abs(2# * maxArea - totalArea - 152400#) < 0.00001
    Check "SP35.export.FAILRegionAvailable", SP35ExportSavedCrackRegionForTests(doc, "SP35_B", "PostCrackState", True) = 3
    Check "SP35.export.preStateAllowed", SP35ExportSavedCrackRegionForTests(doc, "SP35_B", "PreCrackState", True) = 3
    values = ReadSP35AnnotationTable(annotations)
    For i = 2 To UBound(values, 1)
        If CStr(values(i, 14)) = "SP35_A" Then
            annotations.Offset(i - 1, 18).Value2 = CDbl(values(i, 19)) + 1#: Exit For
        End If
    Next i
    countBefore = doc.ModelSpace.Count
    Check "SP35.export.staleWholeRegionSkipped", SP35ExportSavedCrackRegionForTests(doc, "SP35_A", "CrackedState", True) = 0 And doc.ModelSpace.Count = countBefore
    ThisWorkbook.SaveCopyAs ThisWorkbook.Path & "\SP35_WorkbookSnapshot.xlsm"
    GoTo Restore
Failed:
    number = Err.Number: description = Err.Description
    Check "runtime: " & CStr(number) & "; " & description, False
Restore:
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close False
    If Not previousDoc Is Nothing Then previousDoc.Activate
    ThisWorkbook.Names.Item("rngSystemSettings").RefersTo = originalSettings
    ThisWorkbook.Names.Item("rngCrackSummaryAnchor").RefersTo = originalAnchor
    If Not sheet Is Nothing Then sheet.Delete
    On Error GoTo 0
    mReport = mReport & "TOTAL_SP35_WORKBOOK: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
    RunSP35WorkbookTests = mReport
End Function

' ДЛЯ ТЕСТОВ: Rr=90 мм=9 см; psi периодической арматуры равен 4.5 см,
' гладкой - 3.15 см. Итоговая ширина возвращается в мм, а не в см.
Private Sub TestNumericFormula(ByVal formula As CCrackWidthCalculator)
    Dim radius As Double, psi As Double
    radius = formula.SP35ReinforcementRadiusFromData(5400#, 60#)
    CheckNear "SP35.radius.mm", radius, 90#
    psi = formula.SP35PsiFromData(radius, "Periodic")
    CheckNear "SP35.periodic.psi.cm", psi, 4.5
    CheckNear "SP35.periodic.width.mm", formula.SP35CrackWidthFromData(100#, 200000#, psi), 0.0225
    psi = formula.SP35PsiFromData(radius, "Smooth")
    CheckNear "SP35.smooth.psi.cm", psi, 3.15
    CheckNear "SP35.smooth.width.mm", formula.SP35CrackWidthFromData(100#, 200000#, psi), 0.01575
    CheckNear "SP35.zeroArea.radius", formula.SP35ReinforcementRadiusFromData(0#, 60#), 0#
    CheckNear "SP35.zeroStress.width", formula.SP35CrackWidthFromData(0#, 200000#, 4.5), 0#
    CheckNear "SP35.actualDiameterSum", formula.SP35ReinforcementRadiusFromData(5400#, _
        0.75 * (32# + 25# + 20#) + 0.85 * (22# + 20#)), 5400# / 93.45
    CheckNear "SP35.beta.single", formula.SP35GroupBetaFromData(1), 1#
    CheckNear "SP35.beta.double", formula.SP35GroupBetaFromData(2), 0.85
    CheckNear "SP35.beta.triple", formula.SP35GroupBetaFromData(3), 0.75
    CheckNear "NDM.beta.four", formula.SP35GroupBetaFromData(4), 0.75
    CheckNear "SP35.radius.3d", formula.SP35InteractionMultiplierFromData("3d"), 3#
    CheckNear "SP35.radius.5d", formula.SP35InteractionMultiplierFromData("5d"), 5#
    CheckNear "SP35.radius.6d", formula.SP35InteractionMultiplierFromData("6d"), 6#
End Sub

' Сверяет относительные объединения, обозначения и оформление трех картинок ТЗ.
' Строка пояснений и данные не переносятся; верхние ярусы сохраняют переносы.
Private Sub CheckSP35HeaderLayout(ByVal anchor As Object, ByVal units As CUnitSystem)
    Dim top As Object: Set top = anchor.Offset(-4, 0)
    Dim merges As Variant, spec As Variant, expected As Object, cell As Object
    merges = Array(Array(1, 1, 4, 1), Array(1, 2, 4, 2), Array(1, 3, 4, 3), _
        Array(1, 4, 2, 6), Array(1, 7, 2, 9), Array(1, 11, 1, 25), _
        Array(2, 11, 2, 14), Array(2, 15, 2, 18), Array(2, 19, 2, 21), _
        Array(2, 22, 2, 24), Array(1, 26, 2, 29), Array(1, 31, 1, 44), _
        Array(2, 31, 2, 36), Array(2, 37, 2, 40), Array(2, 41, 2, 44), _
        Array(1, 46, 1, 67), Array(2, 46, 2, 51), Array(2, 52, 2, 54), _
        Array(2, 55, 2, 58), Array(2, 59, 2, 62), Array(2, 63, 2, 67), _
        Array(1, 69, 2, 72))
    Dim i As Long
    For i = LBound(merges) To UBound(merges)
        spec = merges(i)
        Set cell = top.Cells(CLng(spec(0)), CLng(spec(1)))
        Set expected = cell.Resize(CLng(spec(2)) - CLng(spec(0)) + 1, CLng(spec(3)) - CLng(spec(1)) + 1)
        Check "SP35.header.merge." & CStr(i), cell.MergeArea.Address = expected.Address
        Check "SP35.header.wrap." & CStr(i), CBool(cell.WrapText)
        Check "SP35.header.style." & CStr(i), cell.Font.Bold And cell.Font.Name = "Arial" And _
            cell.Font.Size = 9 And cell.Interior.Color = RGB(217, 217, 217) And _
            cell.HorizontalAlignment = -4108 And cell.VerticalAlignment = -4108
    Next i
    Check "SP35.header.notesNoWrap", Not CBool(top.Cells(3, 11).Resize(1, 62).WrapText)
    Check "SP35.header.labelsNoWrap", Not CBool(top.Cells(4, 11).Resize(1, 62).WrapText)
    Check "SP35.header.formationTitle", CStr(top.Cells(1, 11).Value2) = "Начало трещинообразования"
    Check "SP35.header.currentTitle", CStr(top.Cells(1, 26).Value2) = "равновесное состояние при заданных нагрузках"
    Check "SP35.header.stateColumn", CStr(top.Cells(4, 25).Value2) = "state" And Not CBool(top.Cells(2, 25).MergeCells)
    Dim labels As Variant
    labels = Array("Ar, " & units.OutputAreaUnit, ChrW$(&H3B2), "n", "d, " & units.OutputLengthUnit, _
        "Rr, " & units.OutputLengthUnit, ChrW$(&H3C8) & ", " & units.OutputLengthUnit, "rebars", "nrebars", _
        ChrW$(&H3C3) & "s, " & units.OutputStressUnit, "Es, " & units.OutputStressUnit, _
        "acrc, " & units.OutputLengthUnit, "acrc,ult, " & units.OutputLengthUnit, "acrc,ult/acrc", "статус")
    For i = LBound(labels) To UBound(labels)
        Check "SP35.header.label." & CStr(i), CStr(top.Cells(4, 31 + i).Value2) = CStr(labels(i))
    Next i
    Check "SP35.header.rebarsSubscript", CBool(top.Cells(4, 38).Characters(2, 6).Font.Subscript)
    Check "SP35.header.sigmaSubscript", CBool(top.Cells(4, 39).Characters(2, 1).Font.Subscript)
    Check "SP35.header.areaSubscript", CBool(top.Cells(4, 31).Characters(2, 1).Font.Subscript)
    Check "SP35.header.radiusCapitalNotSubscript", Not CBool(top.Cells(4, 35).Characters(1, 1).Font.Subscript)
    Check "SP35.header.radiusSuffixSubscript", CBool(top.Cells(4, 35).Characters(2, 1).Font.Subscript)
    Check "SP35.header.boundary", top.Cells(1, 31).MergeArea.Borders(7).LineStyle = 1 And _
        top.Cells(1, 31).MergeArea.Borders(8).LineStyle = 1 And _
        top.Cells(1, 31).MergeArea.Borders(10).LineStyle = 1
    For Each spec In Array(10, 30, 45, 68)
        Check "SP35.header.separator." & CStr(spec), top.Cells(1, CLng(spec)).Resize(4, 1).Interior.ColorIndex = -4142 And _
            top.Cells(1, CLng(spec)).Resize(4, 1).Borders.LineStyle = -4142
    Next spec
End Sub

' ДЛЯ ТЕСТОВ: невозможные численные входы и неизвестные варианты дают
' ошибку данных; чистая формула не возвращает rsNotApplicable/NumFail.
Private Sub TestInvalidData(ByVal formula As CCrackWidthCalculator)
    Dim i As Long, number As Long, description As String, unused As Double
    For i = 1 To 6
        On Error Resume Next
        Err.Clear
        Select Case i
            Case 1: unused = formula.SP35ReinforcementRadiusFromData(1#, 0#)
            Case 2: unused = formula.SP35PsiFromData(-1#, "Periodic")
            Case 3: unused = formula.SP35PsiFromData(90#, "Unknown")
            Case 4: unused = formula.SP35CrackWidthFromData(100#, 0#, 4.5)
            Case 5: unused = formula.SP35GroupBetaFromData(0)
            Case 6: unused = formula.SP35InteractionMultiplierFromData("D4")
        End Select
        number = Err.Number: description = Err.Description
        On Error GoTo 0
        Check "SP35.invalid." & CStr(i) & ".error", number <> 0
        Check "SP35.invalid." & CStr(i) & ".explanation", Len(description) > 0
    Next i
End Sub

' ДЛЯ ТЕСТОВ: запускает только чистые формульные проверки без тяжелого batch.
Public Function RunSP35FormulaTests() As String
    On Error GoTo Failed
    mPassed = 0: mFailed = 0: mReport = vbNullString
    Dim formula As CCrackWidthCalculator
    Set formula = New CCrackWidthCalculator
    TestNumericFormula formula
    TestInvalidData formula
    GoTo Finished
Failed:
    Check "runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    mReport = mReport & "TOTAL_SP35_FORMULA: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
    RunSP35FormulaTests = mReport
End Function

' ДЛЯ ТЕСТОВ: линейная диаграмма только для подготовки по заданной плоскости.
' Решатель не вызывается; stress = 200000*epsilon - независимый эталон.
Private Function LinearSteel() As CMaterialDiagram
    Dim strains(1 To 3) As Double, stresses(1 To 3) As Double, material As CMaterialDiagram
    strains(1) = -0.01: strains(2) = 0#: strains(3) = 0.01
    stresses(1) = -2000#: stresses(2) = 0#: stresses(3) = 2000#
    Set material = New CMaterialDiagram: material.InitializeFromArrays strains, stresses, 3
    Set LinearSteel = material
End Function

' ДЛЯ ТЕСТОВ: прямоугольник и его достоверный контур строятся независимо от
' clipping. Поворот распространяется на геометрию и готовую плоскость НДС.
Private Function RectangularSection(Optional ByVal angle As Double = 0#) As CSectionModel
    Dim section As CSectionModel, x As Variant, y As Variant, i As Long, j As Long
    Set section = New CSectionModel
    section.AddConcreteElement -200# * Sin(angle), 200# * Cos(angle), 160000#, 1, , , "Rectangle", 400#, 400#, angle
    x = Array(-200#, 200#, 200#, -200#): y = Array(0#, 0#, 400#, 400#)
    For i = 0 To 3
        j = (i + 1) Mod 4
        section.Contours.AddContourLine "CONTOUR_LINE_" & CStr(i + 1), _
            CDbl(x(i)) * Cos(angle) - CDbl(y(i)) * Sin(angle), CDbl(x(i)) * Sin(angle) + CDbl(y(i)) * Cos(angle), _
            CDbl(x(j)) * Cos(angle) - CDbl(y(j)) * Sin(angle), CDbl(x(j)) * Sin(angle) + CDbl(y(j)) * Cos(angle)
    Next i
    Set RectangularSection = section
End Function

' ДЛЯ ТЕСТОВ: добавляет реальный номинальный стержень с точной площадью.
Private Sub AddBar(ByVal section As CSectionModel, ByVal x As Double, ByVal y As Double, ByVal diameter As Double, Optional ByVal angle As Double = 0#)
    section.AddRebarElement x * Cos(angle) - y * Sin(angle), x * Sin(angle) + y * Cos(angle), diameter, GEOM_PI * diameter ^ 2 / 4#, "Rebar"
End Sub

' ДЛЯ ТЕСТОВ: сохраняет отдельный прямоугольный проем с собственным ключом
' контура. Несколько проемов никогда не объединяются общим идентификатором.
Private Sub AddOpening(ByVal section As CSectionModel, ByVal name As String, _
        ByVal xMin As Double, ByVal yMin As Double, ByVal xMax As Double, ByVal yMax As Double)
    Dim x As Variant, y As Variant, i As Long, j As Long
    x = Array(xMin, xMax, xMax, xMin): y = Array(yMin, yMin, yMax, yMax)
    For i = 0 To 3
        j = (i + 1) Mod 4
        section.Contours.AddContourLine "CAD_OPENING_" & name & "_1_" & CStr(i), CDbl(x(i)), CDbl(y(i)), CDbl(x(j)), CDbl(y(j)), vbNullString, name, "Opening"
    Next i
End Sub

' ДЛЯ ТЕСТОВ: A-B-C образуют одну транзитивную группу; нулевые и сжатые
' стержни не включаются. Эталон площади локального окна задан аналитически.
Private Sub TestConnectedGroups()
    Dim section As CSectionModel, data As CSP35CrackData, material As CMaterialDiagram, rows As Variant
    Set section = RectangularSection: Set material = LinearSteel
    AddBar section, -25#, 350#, 20#: AddBar section, 0#, 350#, 20#: AddBar section, 25#, 350#, 20#
    AddBar section, 0#, 50#, 20#: AddBar section, 0#, 100#, 20#
    Set data = New CSP35CrackData
    data.Prepare section, -0.0002, 0.000002, 0#, material, 10#, 50#, "Max", "6d", False
    Check "SP35.groups.transitive", data.GroupCount = 1 And data.GroupBarCount(1) = 3
    Check "SP35.groups.noCompressedOrZero", data.BarGroup(4) = 0 And data.BarGroup(5) = 0
    CheckNear "SP35.groups.centerX", data.GroupX(1), 0#
    CheckNear "SP35.groups.centerY", data.GroupY(1), 350#
    CheckNear "SP35.groups.betaDiameterSum", data.CandidateBetaDiameterSum(1), 45#
    CheckNear "SP35.groups.sigmaActual", data.CandidateSigmaS(1), 100#
    Check "SP35.groups.area", Abs(data.CandidateRegion(1).Area - 40800#) < 0.00001
    CheckNear "SP35.groups.sameSideAndRowDiameter", data.CandidateSideDiameter(1), data.CandidateRowDiameter(1)
    rows = data.CandidateRows(1)
    Check "SP35.groups.rowTopology", UBound(rows, 1) = 1 And CLng(rows(1, 1)) = 1 And CLng(rows(1, 2)) = 1
    data.Prepare section, 0.0005, 0#, 0#, material, 10#, 50#, "Max", "6d", True
    Check "SP35.central.oneCandidate", data.CandidateCount = 1
    Check "SP35.central.allConcrete", Abs(data.CandidateRegion(1).Area - 160000#) < 0.00001
    Check "SP35.central.allTensionBars", data.GroupCount = 3 And data.CandidateBetaDiameterSum(1) = 85#
End Sub

' ДЛЯ ТЕСТОВ: разные диаметры используют центр по As; Max/Min/Average
' меняют только построение радиуса, не реальную сумму beta*(10+20).
Private Sub TestMixedDiameters()
    Dim section As CSectionModel, data As CSP35CrackData, material As CMaterialDiagram
    Dim mode As Variant, expectedDiameter As Double, expectedArea As Double
    Set section = RectangularSection: Set material = LinearSteel
    AddBar section, 0#, 350#, 10#: AddBar section, 25#, 350#, 20#
    Set data = New CSP35CrackData
    For Each mode In Array("Max", "Min", "Average")
        Select Case CStr(mode)
            Case "Max": expectedDiameter = 20#: expectedArea = 13200#
            Case "Min": expectedDiameter = 10#: expectedArea = 4800#
            Case "Average": expectedDiameter = 15#: expectedArea = 8550#
        End Select
        data.Prepare section, -0.0002, 0.000002, 0#, material, 10#, 50#, CStr(mode), "3d", False
        Check "SP35.mixed." & CStr(mode) & ".oneGroup", data.GroupCount = 1
        CheckNear "SP35.mixed." & CStr(mode) & ".weightedCenter", data.GroupX(1), 20#
        CheckNear "SP35.mixed." & CStr(mode) & ".sideDiameter", data.CandidateSideDiameter(1), expectedDiameter
        CheckNear "SP35.mixed." & CStr(mode) & ".rowDiameter", data.CandidateRowDiameter(1), expectedDiameter
        CheckNear "SP35.mixed." & CStr(mode) & ".actualDiameters", data.CandidateBetaDiameterSum(1), 25.5
        Check "SP35.mixed." & CStr(mode) & ".area", Abs(data.CandidateRegion(1).Area - expectedArea) < 0.00001
    Next mode
End Sub

' ДЛЯ ТЕСТОВ: неполный ближайший к Н.О. ряд переносит опору на второй;
' исключенный конечной зоной ряд не остается в beta*nd и расшифровке.
Private Sub TestReferenceRows()
    Dim section As CSectionModel, data As CSP35CrackData, material As CMaterialDiagram, mode As Long, y As Double
    Dim rows As Variant
    Set material = LinearSteel: Set data = New CSP35CrackData
    For mode = 0 To 1
        Set section = RectangularSection
        AddBar section, -20#, 350#, 20#: AddBar section, 0#, 350#, 20#: AddBar section, 20#, 350#, 20#
        y = 250#: If mode = 1 Then y = 150#
        AddBar section, 10#, y, 10#
        data.Prepare section, -0.0002, 0.000002, 0#, material, 10#, 50#, "Max", "6d", False
        Check "SP35.rows." & CStr(mode) & ".twoGroups", data.GroupCount = 2
        CheckNear "SP35.rows." & CStr(mode) & ".referenceDiameter", data.CandidateRowDiameter(1), 20#
        Check "SP35.rows." & CStr(mode) & ".area", Abs(data.CandidateRegion(1).Area - 40800#) < 0.00001
        rows = data.CandidateRows(1)
        If mode = 0 Then
            Check "SP35.rows.incomplete.outerFullReferenceFirst", data.CandidateReferenceRow(1) = 1 And data.CandidateRowCount(1) = 2
            Check "SP35.rows.outerFirst", CLng(rows(1, 2)) = 1 And CLng(rows(2, 2)) = 2
            CheckNear "SP35.rows.incomplete.sum", data.CandidateBetaDiameterSum(1), 55#
        Else
            Check "SP35.rows.excluded.renumbered", data.CandidateReferenceRow(1) = 1 And data.CandidateRowCount(1) = 1 And UBound(rows, 1) = 1
            Check "SP35.rows.excluded.notAnAnchor", Not data.CandidateAvailable(2)
            CheckNear "SP35.rows.excluded.sum", data.CandidateBetaDiameterSum(1), 45#
        End If
    Next mode
End Sub

' ДЛЯ ТЕСТОВ: два opening дают местные вырезы, не удаляя бетон за ними.
' Если полоса остается связной, нижний ряд участвует по обычным правилам.
' Сквозной для полосы проем отдельно проверяет исключение другой стенки.
Private Sub TestMultipleOpeningWalls()
    Dim section As CSectionModel, data As CSP35CrackData, material As CMaterialDiagram
    Dim query As CSectionGeometryQuery, region As CConcreteRegion, i As Long, x As Variant, y As Variant, j As Long
    Set section = RectangularSection: Set material = LinearSteel: Set query = New CSectionGeometryQuery
    x = Array(-60#, 60#, 60#, -60#): y = Array(190#, 190#, 250#, 250#)
    For i = 0 To 3
        j = (i + 1) Mod 4
        section.Contours.AddContourLine "CAD_OPENING_FIRST_1_" & CStr(i), CDbl(x(i)), CDbl(y(i)), CDbl(x(j)), CDbl(y(j)), vbNullString, "FIRST", "Opening"
    Next i
    x = Array(80#, 100#, 100#, 80#): y = Array(300#, 300#, 320#, 320#)
    For i = 0 To 3
        j = (i + 1) Mod 4
        section.Contours.AddContourLine "CAD_OPENING_SECOND_1_" & CStr(i), CDbl(x(i)), CDbl(y(i)), CDbl(x(j)), CDbl(y(j)), vbNullString, "SECOND", "Opening"
    Next i
    AddBar section, 0#, 350#, 20#: AddBar section, 0#, 150#, 20#
    Set data = New CSP35CrackData
    data.Prepare section, -0.0002, 0.000002, 0#, material, 10#, 50#, "Max", "6d", False
    Set region = data.CandidateRegion(1)
    Check "SP35.openings.twoHoles.area", Abs(region.Area - 64400#) < 0.00001
    Check "SP35.openings.connectedLowerRow", query.ContainsPoint(region, 0#, 150#)
    Check "SP35.openings.firstHole", Not query.ContainsPoint(region, 0#, 220#)
    Check "SP35.openings.secondHole", Not query.ContainsPoint(region, 90#, 310#)
    Check "SP35.openings.noSecondShadow", query.ContainsPoint(region, 90#, 260#)
    Check "SP35.openings.fiveBaseNormals", region.BoundaryProbeCount = 5
    Dim probes As Variant
    probes = region.BoundaryProbes
    Check "SP35.openings.nearestQ", CBool(probes(3, 1)) And CDbl(probes(3, 3)) >= 250# - 0.000001
    Check "SP35.openings.outerP", CBool(probes(3, 1)) And Abs(CDbl(probes(3, 5)) - 400#) < 0.000001
    Check "SP35.openings.betweenProbesLocalCut", Not query.ContainsPoint(region, 90#, 310#) And query.ContainsPoint(region, 90#, 260#)
    CheckNear "SP35.openings.actualIncludedSum", data.CandidateBetaDiameterSum(1), 40#
    data.Prepare section, 0.0005, 0#, 0#, material, 10#, 50#, "Max", "6d", True
    Check "SP35.openings.central.fullNetArea", Abs(data.CandidateRegion(1).Area - 152400#) < 0.00001 And data.CandidateRegion(1).LoopCount = 3
    Check "SP35.openings.central.noArtificialProbes", data.CandidateRegion(1).BoundaryProbeCount = 0
    CheckNear "SP35.openings.central.allRebars", data.CandidateBetaDiameterSum(1), 40#

    Set section = RectangularSection
    AddOpening section, "BARRIER", -130#, 190#, 130#, 250#
    AddOpening section, "LOCAL", 80#, 300#, 100#, 320#
    AddBar section, 0#, 350#, 20#: AddBar section, 0#, 150#, 20#
    data.Prepare section, -0.0002, 0.000002, 0#, material, 10#, 50#, "Max", "6d", False
    Set region = data.CandidateRegion(1)
    Check "SP35.openings.barrier.area", Abs(region.Area - 35600#) < 0.00001
    Check "SP35.openings.barrier.noOppositeWall", Not query.ContainsPoint(region, 0#, 150#)
    Check "SP35.openings.barrier.noLocalShadow", query.ContainsPoint(region, 90#, 260#)
    CheckNear "SP35.openings.barrier.onlyUpperRow", data.CandidateBetaDiameterSum(1), 20#
End Sub

' ДЛЯ ТЕСТОВ: круг с двумя боковыми круглыми отверстиями воспроизводит
' ошибочное сужение полосы до зазора между ними. Независимый эталон площади
' состоит из сегмента наружного круга минус два круговых выреза; поворот
' проверяет, что ограничение не привязано к вертикальным линиям X/Y.
Private Sub TestCircularLocalOpeningCuts()
    Dim section As CSectionModel, data As CSP35CrackData, query As CSectionGeometryQuery, region As CConcreteRegion
    Dim angle As Variant, i As Long, candidate As Long, anchor As Long, prefix As String, expected As Double, capArea As Double
    Dim intervals As Variant, x1 As Double, y1 As Double, x2 As Double, y2 As Double, cx As Double, cy As Double
    Dim sweep As Double, loopID As Long, arcCount As Long, nx As Double, ny As Double
    capArea = 900# * (GEOM_PI / 2# - Atn(1# / Sqr(8#))) - 10# * Sqr(800#)
    expected = 40# * Sqr(8400#) + 10000# * Atn(0.4 / Sqr(0.84)) - 2# * capArea
    For Each angle In Array(0#, 0.47, -0.81)
        nx = -Sin(CDbl(angle)): ny = Cos(CDbl(angle))
        Set section = New CSectionModel
        section.AddConcreteElement 0#, 0#, GEOM_PI * 10000#, 1
        section.Contours.AddContourCircle "CONTOUR_CIRCLE", 0#, 0#, 100#
        section.Contours.AddContourCircle "CAD_OPENING_LEFT", -50# * ny + 40# * nx, 50# * nx + 40# * ny, _
            30#, vbNullString, "LEFT", "Opening"
        section.Contours.AddContourCircle "CAD_OPENING_RIGHT", 50# * ny + 40# * nx, -50# * nx + 40# * ny, _
            30#, vbNullString, "RIGHT", "Opening"
        AddBar section, 0#, 90#, 20#, CDbl(angle)
        AddBar section, 0#, 75#, 40# / 6#, CDbl(angle)
        Set data = New CSP35CrackData: Set query = New CSectionGeometryQuery
        data.Prepare section, 0#, 0.000002 * ny, 0.000002 * nx, LinearSteel(), 0#, 1#, "Max", "6d", False
        candidate = 0
        For i = 1 To data.CandidateCount
            anchor = data.CandidateAnchorGroup(i)
            If Abs(nx * data.GroupX(anchor) + ny * data.GroupY(anchor) - 75#) < 0.000001 Then candidate = i
        Next i
        prefix = "SP35.localCircle." & CStr(angle)
        Check prefix & ".innerAnchorFound", candidate > 0
        Set region = data.CandidateRegion(candidate)
        Check prefix & ".analyticArea", Abs(region.Area - expected) < 0.00001
        CheckNear prefix & ".sideRadius", data.CandidateSideRadius(candidate), 40#
        CheckNear prefix & ".normalRadius", data.CandidateNormalRadius(candidate), 120#
        Check prefix & ".fullWidthBelowHoles", query.ContainsPoint(region, 35# * ny + 5# * nx, -35# * nx + 5# * ny)
        Check prefix & ".localArcCut", Not query.ContainsPoint(region, 35# * ny + 40# * nx, -35# * nx + 40# * ny)
        Check prefix & ".outsideStrip", Not query.ContainsPoint(region, 41# * ny + 5# * nx, -41# * nx + 5# * ny)
        intervals = query.LineIntervals(region, 5# * nx, 5# * ny, ny, -nx)
        Check prefix & ".oneLowerInterval", UBound(intervals, 1) = 1
        CheckNear prefix & ".lowerLeft", CDbl(intervals(1, 1)), -40#
        CheckNear prefix & ".lowerRight", CDbl(intervals(1, 2)), 40#
        arcCount = 0
        For i = 1 To region.SegmentCount
            region.GetSegment i, x1, y1, x2, y2, cx, cy, sweep, loopID
            If Abs(sweep) > 0.0000000001 Then arcCount = arcCount + 1
            If CDbl(angle) = 0# Then mReport = mReport & "AR_LOCAL_SEGMENT:" & vbTab & CStr(x1) & vbTab & CStr(y1) & _
                vbTab & CStr(x2) & vbTab & CStr(y2) & vbTab & CStr(cx) & vbTab & CStr(cy) & vbTab & CStr(sweep) & vbCrLf
        Next i
        Check prefix & ".exactArcs", arcCount >= 4
        CheckNear prefix & ".actualBarDiameters", data.CandidateBetaDiameterSum(candidate), 20# + 40# / 6#
    Next angle
End Sub

' ДЛЯ ТЕСТОВ: общая геометрия и подготовка не зависят от глобальных X/Y.
' Нормаль Н.О. поворачивается вместе с сечением, чистая площадь неизменна.
Private Sub TestRotatedPreparation()
    Dim section As CSectionModel, data As CSP35CrackData, material As CMaterialDiagram, angle As Double
    angle = 0.47: Set section = RectangularSection(angle): Set material = LinearSteel
    AddBar section, -25#, 350#, 20#, angle: AddBar section, 0#, 350#, 20#, angle: AddBar section, 25#, 350#, 20#, angle
    Set data = New CSP35CrackData
    data.Prepare section, -0.0002, 0.000002 * Cos(angle), -0.000002 * Sin(angle), material, 10#, 50#, "Max", "6d", False
    Check "SP35.rotation.area", Abs(data.CandidateRegion(1).Area - 40800#) < 0.00001
    Check "SP35.rotation.sigma", Abs(data.CandidateSigmaS(1) - 100#) < 0.000000001
    CheckNear "SP35.rotation.betaDiameterSum", data.CandidateBetaDiameterSum(1), 45#
End Sub

' ДЛЯ ТЕСТОВ: минимальный воспроизводимый сценарий двух боковых отверстий.
' Позволяет отдельно проверить прежнюю ошибку и исправленный сохраненный код.
Public Function RunSP35LocalOpeningTests() As String
    mPassed = 0: mFailed = 0: mReport = vbNullString
    On Error GoTo Failed
    TestCircularLocalOpeningCuts
    GoTo Finished
Failed:
    Check "localOpening.runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    mReport = mReport & "TOTAL_SP35_LOCAL_OPENING: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
    RunSP35LocalOpeningTests = mReport
End Function

' ДЛЯ ТЕСТОВ: соседние группы смещены по нормали, поэтому длина между
' центрами отличается от проекции на Н.О. Проверяет обе стороны окна
' и сохранение размеров после поворота всей физической постановки.
Private Sub TestProjectedLateralSpacing()
    Dim section As CSectionModel, data As CSP35CrackData, query As CSectionGeometryQuery
    Dim region As CConcreteRegion, angle As Variant, minT As Double, maxT As Double, prefix As String
    For Each angle In Array(0#, 0.47, -0.81)
        Set section = RectangularSection(CDbl(angle))
        AddBar section, 0#, 300#, 20#, CDbl(angle)
        AddBar section, 100#, 330#, 20#, CDbl(angle)
        AddBar section, -90#, 320#, 20#, CDbl(angle)
        Set data = New CSP35CrackData
        data.Prepare section, -0.0002, 0.000002 * Cos(CDbl(angle)), -0.000002 * Sin(CDbl(angle)), _
            LinearSteel(), 10#, 50#, "Max", "6d", False
        prefix = "SP35.projectedSpacing." & CStr(angle)
        Check prefix & ".separateGroups", data.GroupCount = 3
        Check prefix & ".sameRow", data.GroupRow(1) = data.GroupRow(2) And data.GroupRow(1) = data.GroupRow(3)
        Set region = data.CandidateRegion(1): Set query = New CSectionGeometryQuery
        query.ProjectionBounds region, -Cos(CDbl(angle)), -Sin(CDbl(angle)), minT, maxT
        Check prefix & ".leftProjectionHalf", Abs(minT + 50#) < 0.000001
        Check prefix & ".rightProjectionHalf", Abs(maxT - 45#) < 0.000001
        Check prefix & ".notEuclideanHalf", Abs(minT + Sqr(100# ^ 2 + 30# ^ 2) / 2#) > 1#
        Check prefix & ".area", Abs(region.Area - 20900#) < 0.00001
    Next angle
End Sub

' ДЛЯ ТЕСТОВ: сосед другого ряда теперь ограничивает боковую границу,
' если t/L достигает порога. При отсутствии соседа номинальный радиус
' не уменьшается до грани на линии якоря; бетон пересекается с готовой полосой.
Private Sub TestLateralMaterialBoundary()
    Dim section As CSectionModel, data As CSP35CrackData, query As CSectionGeometryQuery, region As CConcreteRegion
    Dim nx As Double, ny As Double, minT As Double, maxT As Double, centerT As Double
    nx = 0.488832654432331: ny = 0.872377576488897
    Set section = RectangularSection
    AddBar section, 160#, 300#, 32#: AddBar section, 120#, 340#, 32#
    AddBar section, 160#, 202.5, 32#: AddBar section, 30#, 340#, 32#
    Set data = New CSP35CrackData
    data.Prepare section, -0.0002, 0.000002 * ny, 0.000002 * nx, LinearSteel(), 10#, 50#, "Max", "6d", False
    Check "SP35.materialBoundary.neighborOtherRow", data.GroupRow(1) <> data.GroupRow(3)
    Set region = data.CandidateRegion(1): Set query = New CSectionGeometryQuery
    query.ProjectionBounds region, -ny, nx, minT, maxT
    centerT = -ny * 160# + nx * 300#
    Check "SP35.materialBoundary.otherRowProjection", Abs((centerT - minT) - 97.5 * nx / 2#) < 0.000001
    Check "SP35.materialBoundary.projected27", Abs((maxT - centerT) - (40# * ny + 40# * nx) / 2#) < 0.000001
    Check "SP35.materialBoundary.notVerticalStepHalf", Abs((centerT - minT) - 97.5 / 2#) > 1#
    Set section = RectangularSection
    AddBar section, 160#, 300#, 32#
    data.Prepare section, -0.0002, 0.000002 * ny, 0.000002 * nx, LinearSteel(), 10#, 50#, "Max", "6d", False
    Set region = data.CandidateRegion(1)
    CheckNear "SP35.materialBoundary.nominalLeftUnchanged", data.CandidateLeftDistance(1), 192#
    CheckNear "SP35.materialBoundary.nominalRightUnchanged", data.CandidateRightDistance(1), 192#
    Check "SP35.materialBoundary.notClampedByAnchorLine", query.ContainsPoint(region, 195#, 150#)
    Check "SP35.materialBoundary.noOutsideConcrete", Not query.ContainsPoint(region, 201#, 150#)
End Sub

' ДЛЯ ТЕСТОВ: t/L проверяется отдельно для ближайшей по полному L группы
' с каждой стороны. RowTolerance меняет только разбиение рядов, не боковые
' расстояния. Проверяет обе ветки, равенство, отсутствие соседа, одинаковую
' проекцию, поворот и включение всех групп конечной области.
Private Sub TestNeighborRatioBoundaries()
    Dim section As CSectionModel, data As CSP35CrackData, angle As Variant, rowTolerance As Variant
    Dim ratio As Variant, prefix As String, expectedRight As Double, rows As Variant, region As CConcreteRegion
    Dim query As CSectionGeometryQuery, number As Long
    Set query = New CSectionGeometryQuery
    For Each angle In Array(0#, 0.47, -0.81)
        For Each rowTolerance In Array(0#, 50#, 200#)
            For Each ratio In Array(0.199, 0.201)
                Set section = RectangularSection(CDbl(angle))
                AddBar section, 0#, 250#, 20#, CDbl(angle)
                AddBar section, -100# * CDbl(ratio), 250# + Sqr(10000# * (1# - CDbl(ratio) ^ 2)), 5#, CDbl(angle)
                Set data = New CSP35CrackData
                data.Prepare section, -0.0002, 0.000002 * Cos(CDbl(angle)), -0.000002 * Sin(CDbl(angle)), _
                    LinearSteel(), 0#, CDbl(rowTolerance), "Max", "6d", False, 0.2
                expectedRight = 120#: If CDbl(ratio) > 0.2 Then expectedRight = 50# * CDbl(ratio)
                prefix = "SP35.neighborRatio." & CStr(angle) & "." & CStr(rowTolerance) & "." & CStr(ratio)
                CheckNear prefix & ".leftNoNeighbor", data.CandidateLeftDistance(1), 120#
                CheckNear prefix & ".right", data.CandidateRightDistance(1), expectedRight
                CheckNear prefix & ".storedThreshold", data.NeighborRatioLimit, 0.2
            Next ratio
        Next rowTolerance
    Next angle
    Set section = RectangularSection
    AddBar section, 0#, 250#, 20#: AddBar section, 30#, 290#, 20#
    Set data = New CSP35CrackData
    data.Prepare section, -0.0002, 0.000002, 0#, LinearSteel(), 0#, 10#, "Max", "6d", False, 0.6
    CheckNear "SP35.neighborRatio.equalThreshold", data.CandidateLeftDistance(1), 15#
    data.Prepare section, -0.0002, 0.000002, 0#, LinearSteel(), 0#, 10#, "Max", "6d", False, 1#
    CheckNear "SP35.neighborRatio.highThresholdFullRadius", data.CandidateLeftDistance(1), 120#
    data.Prepare section, -0.0002, 0.000002, 0#, LinearSteel(), 0#, 10#, "Max", "6d", False, 0#
    CheckNear "SP35.neighborRatio.zeroThresholdHalfProjection", data.CandidateLeftDistance(1), 15#

    Set section = RectangularSection
    AddBar section, 0#, 250#, 20#: AddBar section, 60#, 250#, 20#
    AddBar section, -20#, 350#, 20#: AddBar section, 5#, 380#, 20#
    data.Prepare section, -0.0002, 0.000002, 0#, LinearSteel(), 0#, 10#, "Max", "6d", False, 0.2
    CheckNear "SP35.neighborRatio.nearestByFullDistance", data.CandidateLeftDistance(1), 30#
    CheckNear "SP35.neighborRatio.independentRightFullRadius", data.CandidateRightDistance(1), 120#
    CheckNear "SP35.neighborRatio.allIncludedActualBars", data.CandidateBetaDiameterSum(1), 60#
    rows = data.CandidateRows(1)
    Check "SP35.neighborRatio.outerFirstGlobal", data.GroupRow(4) = 1 And data.GroupRow(3) = 2 And data.GroupRow(1) = 3
    Check "SP35.neighborRatio.outerFirstLocal", CLng(rows(1, 2)) = 3 And CLng(rows(2, 2)) = 2 And CLng(rows(3, 2)) = 1
    Set region = data.CandidateRegion(1)
    Check "SP35.neighborRatio.keepNeutralLine", Not query.ContainsPoint(region, 0#, 99#)

    Set section = RectangularSection
    AddBar section, 0#, 250#, 20#: AddBar section, 0#, 350#, 20#
    data.Prepare section, -0.0002, 0.000002, 0#, LinearSteel(), 0#, 10#, "Max", "6d", False, 0#
    CheckNear "SP35.neighborRatio.sameProjectionNotLeft", data.CandidateLeftDistance(1), 120#
    CheckNear "SP35.neighborRatio.sameProjectionNotRight", data.CandidateRightDistance(1), 120#
    On Error Resume Next
    data.Prepare section, -0.0002, 0.000002, 0#, LinearSteel(), 0#, 10#, "Max", "6d", False, 1.01
    number = Err.Number: Err.Clear
    On Error GoTo 0
    Check "SP35.neighborRatio.invalidThreshold", number <> 0 And data.CandidateCount = 0
End Sub

' ДЛЯ ТЕСТОВ: границы пользовательских допусков отличаются от малого
' геометрического epsilon; 10.1/50.1 не округляются до допустимых 10/50.
Private Sub TestPreparationToleranceBoundaries()
    Dim section As CSectionModel, data As CSP35CrackData, material As CMaterialDiagram
    Dim value As Variant, expected As Long, number As Long, description As String, members As Variant
    Set material = LinearSteel: Set data = New CSP35CrackData
    For Each value In Array(9.9, 10#, 10.1)
        Set section = RectangularSection
        AddBar section, 0#, 350#, 20#: AddBar section, 20# + CDbl(value), 350#, 20#
        data.Prepare section, -0.0002, 0.000002, 0#, material, 10#, 50#, "Max", "6d", False
        expected = 1: If CDbl(value) > 10# Then expected = 2
        Check "SP35.groupGap." & CStr(value), data.GroupCount = expected
    Next value
    For Each value In Array(49.9, 50#, 50.1)
        Set section = RectangularSection
        AddBar section, 0#, 250#, 20#: AddBar section, 100#, 250# + CDbl(value), 20#
        data.Prepare section, -0.0002, 0.000002, 0#, material, 10#, 50#, "Max", "6d", False
        Check "SP35.rowTolerance." & CStr(value), (data.GroupRow(1) = data.GroupRow(2)) = (CDbl(value) <= 50#)
    Next value
    members = data.GroupMembers(1)
    Check "SP35.snapshot.members", UBound(members) = 1 And data.BarID(CLng(members(1))) = "R1"
    CheckNear "SP35.snapshot.nominalDiameter", data.BarDiameter(CLng(members(1))), 20#
    data.Prepare section, -0.001, 0#, 0#, material, 10#, 50#, "Max", "6d", False
    Check "SP35.noTension.noCandidates", data.CandidateCount = 0 And data.GroupCount = 0
    On Error Resume Next
    data.Prepare section, 0.0005, 0#, 0#, material, -1#, 50#, "Max", "6d", True
    number = Err.Number: description = Err.Description: Err.Clear
    On Error GoTo 0
    Check "SP35.invalidTolerance.explained", number <> 0 And Len(description) > 20
    Check "SP35.invalidTolerance.noOldCandidates", data.CandidateCount = 0
End Sub

' ДЛЯ ТЕСТОВ: только подготовка групп/областей по готовым числам, без batch,
' Search и solver. Проверяет центральную/изгибную постановку и два проема.
Public Function RunSP35PreparationTests() As String
    mPassed = 0: mFailed = 0: mReport = vbNullString
    On Error GoTo Failed
    TestConnectedGroups
    TestMixedDiameters
    TestReferenceRows
    TestMultipleOpeningWalls
    TestCircularLocalOpeningCuts
    TestRotatedPreparation
    TestProjectedLateralSpacing
    TestLateralMaterialBoundary
    TestNeighborRatioBoundaries
    TestIndivisibleGroupMembership
    TestPreparationToleranceBoundaries
    GoTo Finished
Failed:
    Check "preparation.runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    mReport = mReport & "TOTAL_SP35_PREPARATION: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
    RunSP35PreparationTests = mReport
End Function

' ДЛЯ ТЕСТОВ: боковая граница проходит между стержнями исходной пары.
' Центр снаружи исключает обоих, внутри или на границе включает обоих.
' Поворот не меняет неделимость, полный beta/n и сумму реальных диаметров.
Private Sub TestIndivisibleGroupMembership()
    Dim section As CSectionModel, data As CSP35CrackData, query As CSectionGeometryQuery
    Dim angle As Variant, center As Variant, members As Variant, group As Long, prefix As String
    Dim region As CConcreteRegion, rows As Variant, expectedIncluded As Boolean
    Set query = New CSectionGeometryQuery
    For Each angle In Array(0#, 0.47, -0.81)
      For Each center In Array(130#, 115#, 120#)
        Set section = RectangularSection(CDbl(angle))
        AddBar section, 0#, 250#, 20#, CDbl(angle)
        AddBar section, CDbl(center) - 15#, 300#, 20#, CDbl(angle)
        AddBar section, CDbl(center) + 15#, 300#, 20#, CDbl(angle)
        Set data = New CSP35CrackData
        data.Prepare section, -0.0002, 0.000002 * Cos(CDbl(angle)), -0.000002 * Sin(CDbl(angle)), _
            LinearSteel(), 10#, 50#, "Max", "6d", False, 1#
        group = data.BarGroup(2): prefix = "SP35.indivisibleGroup." & CStr(angle) & "." & CStr(center)
        expectedIncluded = (CDbl(center) <= 120#)
        Check prefix & ".originalPair", data.GroupBarCount(group) = 2 And data.BarGroup(3) = group
        Set region = data.CandidateRegion(1): rows = data.CandidateRows(1)
        Check prefix & ".barInside", query.ContainsPoint(region, section.RebarX(2), section.RebarY(2))
        Check prefix & ".barOutside", Not query.ContainsPoint(region, section.RebarX(3), section.RebarY(3))
        Check prefix & ".centerDecision", query.ContainsPoint(region, data.GroupX(group), data.GroupY(group)) = expectedIncluded
        Check prefix & ".includedGroupCount", UBound(rows, 1) = IIf(expectedIncluded, 2, 1)
        members = data.GroupMembers(group)
        Check prefix & ".fullMembers", UBound(members) = 2 And CLng(members(1)) = 2 And CLng(members(2)) = 3
        CheckNear prefix & ".fullBeta", data.GroupBeta(group), 0.85
        CheckNear prefix & ".denominator", data.CandidateBetaDiameterSum(1), IIf(expectedIncluded, 54#, 20#)
        CheckCandidateMembership prefix, data, section, query
      Next center
    Next angle
End Sub

' ДЛЯ ТЕСТОВ: проверяет новое поле через настоящий reader/Width, включая
' неверные значения, неактивную ветку СП 63 и динамический адрес после
' перемещения таблицы. Все временные данные находятся в собственной копии.
Public Function RunSP35NeighborSettingsTests() As String
    mPassed = 0: mFailed = 0: mReport = vbNullString
    On Error GoTo Failed
    Dim original As String, sheet As Object, source As Object, values As Variant, sourceRow As Long
    Dim settings As CSystemSettingsReader, width As CCrackWidthCalculator, value As Variant, number As Long, description As String
    Dim key As String, address As String, oldAddress As String, row As Long, expectedValid As Boolean, prefix As String
    key = "SLS.Crack.SP35.NeighborRatioLimit"
    original = ThisWorkbook.Names.Item("rngSystemSettings").RefersTo
    values = ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange.Value2
    Set sheet = ThisWorkbook.Worksheets.Add: sheet.Name = "_SP35NeighborSettings"
    Set source = sheet.Cells(20, 3).Resize(UBound(values, 1) + 2, UBound(values, 2))
    source.Resize(UBound(values, 1), UBound(values, 2)).Value2 = values
    ThisWorkbook.Names.Item("rngSystemSettings").RefersTo = "='" & sheet.Name & "'!" & source.Address
    SetWorkbookTestSetting source, "SLS.Crack.Code", "SP35"
    SetWorkbookTestSetting source, key, 0.2
    For row = 2 To source.Rows.Count
        If CStr(source.Cells(row, 1).Value2) = key Then sourceRow = row: Exit For
    Next row
    address = source.Cells(sourceRow, 2).Address(False, False)
    For Each value In Array(0#, 0.2, 1#, -0.01, 1.01, vbNullString, "wrong")
        SetWorkbookTestSetting source, key, value
        On Error Resume Next
        Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
        Set width = New CCrackWidthCalculator: width.ApplySettings settings
        number = Err.Number: description = Err.Description: Err.Clear
        On Error GoTo Failed
        expectedValid = IsNumeric(value)
        If expectedValid Then expectedValid = (CDbl(value) >= 0# And CDbl(value) <= 1#)
        prefix = "SP35.neighborSetting." & CStr(value)
        If expectedValid Then
            Check prefix & ".accepted", number = 0
        Else
            Check prefix & ".rejected", number <> 0
            Check prefix & ".key", InStr(description, key) > 0
            Check prefix & ".location", InStr(description, address) > 0 And InStr(description, sheet.Name) > 0
        End If
    Next value
    SetWorkbookTestSetting source, "SLS.Crack.Code", "SP63"
    SetWorkbookTestSetting source, key, "wrong"
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set width = New CCrackWidthCalculator: width.ApplySettings settings
    Check "SP35.neighborSetting.inactiveIgnored", True
    SetWorkbookTestSetting source, "SLS.Crack.Code", "SP35"
    oldAddress = address
    source.Cut sheet.Cells(20, 10)
    Set source = ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange
    address = source.Cells(sourceRow, 2).Address(False, False)
    On Error Resume Next
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set width = New CCrackWidthCalculator: width.ApplySettings settings
    number = Err.Number: description = Err.Description: Err.Clear
    On Error GoTo Failed
    Check "SP35.neighborSetting.movedAddress", number <> 0 And InStr(description, address) > 0 And InStr(description, oldAddress) = 0
    source.Cells(sourceRow, 1).Value2 = "REMOVED_NEIGHBOR_RATIO"
    On Error Resume Next
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set width = New CCrackWidthCalculator: width.ApplySettings settings
    number = Err.Number: description = Err.Description: Err.Clear
    On Error GoTo Failed
    Check "SP35.neighborSetting.missingRequired", number <> 0 And InStr(description, key) > 0
    GoTo Restore
Failed:
    Check "neighborSetting.runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Restore:
    On Error Resume Next
    If Len(original) > 0 Then ThisWorkbook.Names.Item("rngSystemSettings").RefersTo = original
    If Not sheet Is Nothing Then sheet.Delete
    On Error GoTo 0
    mReport = mReport & "TOTAL_SP35_NEIGHBOR_SETTINGS: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
    RunSP35NeighborSettingsTests = mReport
End Function

' ДЛЯ ТЕСТОВ: повторяет текущие сочетания в памяти и независимо проверяет
' проекции боковых границ. Сохраняет центры/ряды групп для разбора конкретной
' области пользователя; Config и Results не изменяются.
Public Function RunSP35CurrentProjectionTests() As String
    mPassed = 0: mFailed = 0: mReport = vbNullString
    On Error GoTo Failed
    Dim settings As CSystemSettingsReader, units As CUnitSystem, section As CSectionModel
    Dim provider As CMaterialModelProvider, profiles As CCalculationProfileCatalog
    Dim batch As CBatchSectionCalculator, reader As CLoadCombinationReader
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    Set section = BuildWorkbookSectionModel(ThisWorkbook, settings, units)
    Set provider = New CMaterialModelProvider: provider.Initialize settings, units
    Set profiles = New CCalculationProfileCatalog: profiles.LoadFromWorkbook ThisWorkbook
    Set batch = New CBatchSectionCalculator: batch.Initialize section, provider
    Set batch.ProfileCatalog = profiles: batch.ApplySettings settings, units
    Set reader = New CLoadCombinationReader: reader.LoadFromWorkbook ThisWorkbook, batch, units
    Audit03ApplyLoadReferenceForTests section, settings, units, batch
    batch.Execute
    Dim result As CCrackWidthResult, data As CSP35CrackData, query As CSectionGeometryQuery
    Dim region As CConcreteRegion, i As Long, j As Long, candidate As Long, anchor As Long
    Dim tx As Double, ty As Double, centerT As Double, spacing As Double, left As Double, right As Double
    Dim minT As Double, maxT As Double, prefix As String, members As Variant, member As Variant, ids As String
    Dim distance As Double, nearestLeft As Double, nearestRight As Double, spacingLeft As Double, spacingRight As Double
    Set query = New CSectionGeometryQuery: query.Initialize section
    For i = 1 To batch.Count
        Set result = batch.ResultAt(i).CrackResult.Width: Set data = result.SP35Data
        If Not data Is Nothing Then
            If Not data.CentralTension Then
                tx = -data.NormalY: ty = data.NormalX
                mReport = mReport & "PROJECTION_NORMAL: " & CStr(data.NormalX) & "; " & CStr(data.NormalY) & vbCrLf
                For j = 1 To data.GroupCount
                    members = data.GroupMembers(j): ids = vbNullString
                    For Each member In members
                        If Len(ids) > 0 Then ids = ids & ","
                        ids = ids & data.BarID(CLng(member))
                    Next member
                    mReport = mReport & "PROJECTION_GROUP: G" & CStr(j) & "; bars=" & ids & "; row=" & CStr(data.GroupRow(j)) & _
                        "; x=" & CStr(data.GroupX(j)) & "; y=" & CStr(data.GroupY(j)) & vbCrLf
                Next j
                For candidate = 1 To data.CandidateCount
                    If data.CandidateAvailable(candidate) Then
                        anchor = data.CandidateAnchorGroup(candidate)
                        centerT = tx * data.GroupX(anchor) + ty * data.GroupY(anchor)
                        left = data.CandidateSideRadius(candidate): right = left
                        nearestLeft = 1E+100: nearestRight = 1E+100: spacingLeft = 0#: spacingRight = 0#
                        For j = 1 To data.GroupCount
                            If j <> anchor Then
                                spacing = tx * (data.GroupX(j) - data.GroupX(anchor)) + ty * (data.GroupY(j) - data.GroupY(anchor))
                                distance = Sqr((data.GroupX(j) - data.GroupX(anchor)) ^ 2 + (data.GroupY(j) - data.GroupY(anchor)) ^ 2)
                                If spacing < -0.000001 And distance < nearestLeft Then
                                    nearestLeft = distance: spacingLeft = -spacing
                                ElseIf spacing > 0.000001 And distance < nearestRight Then
                                    nearestRight = distance: spacingRight = spacing
                                End If
                            End If
                        Next j
                        If spacingLeft > 0# Then
                            If spacingLeft / nearestLeft >= data.NeighborRatioLimit And spacingLeft / 2# < left Then left = spacingLeft / 2#
                        End If
                        If spacingRight > 0# Then
                            If spacingRight / nearestRight >= data.NeighborRatioLimit And spacingRight / 2# < right Then right = spacingRight / 2#
                        End If
                        Set region = data.CandidateRegion(candidate): query.ProjectionBounds region, tx, ty, minT, maxT
                        prefix = "SP35.currentProjection." & CStr(i) & ".G" & CStr(anchor)
                        Check prefix & ".left", minT >= centerT - left - 0.000001
                        Check prefix & ".right", maxT <= centerT + right + 0.000001
                        CheckNear prefix & ".nominalLeft", data.CandidateLeftDistance(candidate), left
                        CheckNear prefix & ".nominalRight", data.CandidateRightDistance(candidate), right
                        mReport = mReport & "PROJECTION_WINDOW: " & prefix & "; left=" & CStr(left) & "; right=" & CStr(right) & _
                            "; actualLeft=" & CStr(centerT - minT) & "; actualRight=" & CStr(maxT - centerT) & vbCrLf
                    End If
                Next candidate
            End If
        End If
    Next i
    Check "SP35.currentProjection.coverage", mPassed > 0
    GoTo Finished
Failed:
    Check "currentProjection.runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    mReport = mReport & "TOTAL_SP35_CURRENT_PROJECTION: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
    RunSP35CurrentProjectionTests = mReport
End Function

' ДЛЯ ТЕСТОВ: симметричная сетка и восемь реальных стержней дают независимый
' эталон центрального растяжения и несколько групп/рядов при общем изгибе.
Private Function EndToEndSection() As CSectionModel
    Dim section As CSectionModel, i As Long, j As Long, x As Variant, y As Variant
    Set section = New CSectionModel: section.SourceType = "SP35EndToEnd"
    For i = 0 To 19
        For j = 0 To 19
            section.AddConcreteElement -190# + 20# * i, -190# + 20# * j, 400#, 1, , , _
                "Rectangle", 20#, 20#, 0#, , 400# * 400# / 12#, 400# * 400# / 12#, 0#, True
        Next j
    Next i
    section.Contours.AddContourLine "OUTER_1", -200#, -200#, 200#, -200#
    section.Contours.AddContourLine "OUTER_2", 200#, -200#, 200#, 200#
    section.Contours.AddContourLine "OUTER_3", 200#, 200#, -200#, 200#
    section.Contours.AddContourLine "OUTER_4", -200#, 200#, -200#, -200#
    For Each y In Array(-150#, -60#, 60#, 150#)
        For Each x In Array(-140#, 140#)
            AddBar section, CDbl(x), CDbl(y), 25#
        Next x
    Next y
    Set EndToEndSection = section
End Function

' ДЛЯ ТЕСТОВ: проверяет сохраненную HollowRectangle и настоящее сочетание
' пользователя в отдельной test-copy. Все стержни обязаны лежать в бетоне,
' а пригодный текущий НДС должен доходить до расчета ширины СП 35.
Public Function RunSP35SavedHollowTests() As String
    mPassed = 0: mFailed = 0: mReport = vbNullString
    On Error GoTo Failed
    Dim settings As CSystemSettingsReader, units As CUnitSystem, section As CSectionModel
    Dim query As CSectionGeometryQuery, region As CConcreteRegion, i As Long
    Dim provider As CMaterialModelProvider, profiles As CCalculationProfileCatalog
    Dim batch As CBatchSectionCalculator, reader As CLoadCombinationReader, width As CCrackWidthResult
    Dim writer As CCrackSummaryWriter, anchor As Object, listed As String, bar As Long, count As Long, group As Long
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    Set section = BuildWorkbookSectionModel(ThisWorkbook, settings, units)
    Check "savedHollow.source", section.RebarCount = 56 And section.Contours.Count = 16
    Set query = New CSectionGeometryQuery: query.Initialize section: Set region = query.ConcreteDomain
    For i = 1 To section.RebarCount
        Check "savedHollow.inConcrete." & section.RebarID(i) & "; x=" & Format$(section.RebarX(i), "0.00000000000000000") & _
            "; y=" & Format$(section.RebarY(i), "0.00000000000000000"), query.ContainsPoint(region, section.RebarX(i), section.RebarY(i))
    Next i
    Set provider = New CMaterialModelProvider: provider.Initialize settings, units
    Set profiles = New CCalculationProfileCatalog: profiles.LoadFromWorkbook ThisWorkbook
    Set batch = New CBatchSectionCalculator: batch.Initialize section, provider
    Set batch.ProfileCatalog = profiles: batch.ApplySettings settings, units
    Set reader = New CLoadCombinationReader: reader.LoadFromWorkbook ThisWorkbook, batch, units
    Audit03ApplyLoadReferenceForTests section, settings, units, batch
    batch.Execute
    Set writer = New CCrackSummaryWriter: writer.WriteSummary ThisWorkbook, batch, units
    Set anchor = ThisWorkbook.Names.Item("rngCrackSummaryAnchor").RefersToRange
    For i = 1 To batch.Count
        Set width = batch.ResultAt(i).CrackResult.Width
        Check "savedHollow.width.calculated; " & width.ResultMeta.ResultComment, width.ResultMeta.Calculated
        Check "savedHollow.width.noInputError", width.ResultMeta.InternalStatus <> rsInvalidInput
        CheckCandidateMembership "savedHollow", width.SP35Data, section, query
        listed = CStr(anchor.Offset(i - 1, 36).Value2): count = 0
        Set region = width.InteractionRegion
        For bar = 1 To section.RebarCount
            group = width.SP35Data.BarGroup(bar)
            If group > 0 Then
                If query.ContainsPoint(region, width.SP35Data.GroupX(group), width.SP35Data.GroupY(group)) Then
                    count = count + 1
                    Check "savedHollow.writer.member." & section.RebarID(bar), _
                        InStr(1, ", " & width.TensionRebarIds & ", ", ", " & section.RebarID(bar) & ", ", vbBinaryCompare) > 0 And _
                        InStr(1, Replace(Replace(Replace(Replace(listed, " | ", ", "), " ряд: ", ", "), "; ", ", "), "=", ", ") & ", ", ", " & section.RebarID(bar) & ", ", vbBinaryCompare) > 0
                End If
            End If
        Next bar
        Check "savedHollow.writer.count", CLng(anchor.Offset(i - 1, 37).Value2) = count And width.TensionRebarCount = count
        Check "savedHollow.writer.noWrap", Not CBool(anchor.Offset(i - 1, 36).WrapText)
        mReport = mReport & "SAVED_HOLLOW_REBARS: " & listed & "; count=" & CStr(count) & vbCrLf
        mReport = mReport & "SAVED_HOLLOW_WIDTH: " & CStr(width.CrackWidth) & "; " & width.ResultMeta.ResultComment & vbCrLf
    Next i
    GoTo Finished
Failed:
    Check "savedHollow.runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    mReport = mReport & "TOTAL_SP35_SAVED_HOLLOW: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
    RunSP35SavedHollowTests = mReport
End Function

' ДЛЯ ТЕСТОВ: принадлежность всех стержней группы сравнивается с положением
' ее исходного центра в конечной области. Проверяет неделимость всех кандидатов.
Private Sub CheckCandidateMembership(ByVal prefix As String, ByVal data As CSP35CrackData, _
        ByVal section As CSectionModel, ByVal query As CSectionGeometryQuery)
    Dim candidate As Long, bar As Long, row As Long, rows As Variant, members As Variant, member As Variant
    Dim region As CConcreteRegion, expected As Boolean, included As Boolean, group As Long
    Dim actualCount As Long, formula As CCrackWidthCalculator, expectedSum As Double, expectedCounts() As Long, ids As String
    If data Is Nothing Then Check prefix & ".dataPresent", False: Exit Sub
    Set formula = New CCrackWidthCalculator
    For candidate = 1 To data.CandidateCount
        If data.CandidateAvailable(candidate) Then
            Set region = data.CandidateRegion(candidate): rows = data.CandidateRows(candidate)
            ReDim expectedCounts(1 To data.GroupCount): expectedSum = 0#
            For bar = 1 To section.RebarCount
                group = data.BarGroup(bar)
                expected = False: included = False
                If group > 0 Then expected = query.ContainsPoint(region, data.GroupX(group), data.GroupY(group))
                If expected Then expectedCounts(group) = expectedCounts(group) + 1
                For row = LBound(rows, 1) To UBound(rows, 1)
                    members = data.GroupMembers(CLng(rows(row, 1)))
                    For Each member In members
                        If CLng(member) = bar Then included = True
                    Next member
                Next row
                Check prefix & ".membership.G" & CStr(candidate) & "." & section.RebarID(bar), included = expected
            Next bar
            For row = LBound(rows, 1) To UBound(rows, 1)
                group = CLng(rows(row, 1)): members = data.GroupMembers(group)
                actualCount = UBound(members) - LBound(members) + 1
                Check prefix & ".count.G" & CStr(candidate) & "." & CStr(group), _
                    actualCount = expectedCounts(group) And data.GroupBarCount(group) = expectedCounts(group)
                CheckNear prefix & ".beta.G" & CStr(candidate) & "." & CStr(group), _
                    data.GroupBeta(group), formula.SP35GroupBetaFromData(expectedCounts(group))
                For Each member In members
                    expectedSum = expectedSum + formula.SP35GroupBetaFromData(expectedCounts(group)) * section.RebarDiameter(CLng(member))
                Next member
            Next row
            CheckNear prefix & ".denominator.G" & CStr(candidate), data.CandidateBetaDiameterSum(candidate), expectedSum
            If prefix = "savedHollow" And candidate = 8 Then
                ids = vbNullString
                For bar = 1 To section.RebarCount
                    group = data.BarGroup(bar)
                    If group > 0 Then
                        If query.ContainsPoint(region, data.GroupX(group), data.GroupY(group)) Then
                            If Len(ids) > 0 Then ids = ids & ", "
                            ids = ids & section.RebarID(bar)
                        End If
                    End If
                Next bar
                mReport = mReport & "SAVED_HOLLOW_G8: " & ids & "; betaDiameterSum=" & CStr(expectedSum) & vbCrLf
            End If
        End If
    Next candidate
End Sub

' ДЛЯ ТЕСТОВ: сверяет готовую ширину с независимой записью нормативной
' формулы и ее сохраненной геометрией; исходные НДС здесь не решаются повторно.
Private Sub CheckEndToEndWidth(ByVal prefix As String, ByVal result As CCombinationResult, _
        ByVal section As CSectionModel, ByVal standard As String, ByVal caseIndex As Long)
    Dim width As CCrackWidthResult, region As CConcreteRegion, expected As Double
    Set width = result.CrackResult.Width
    If caseIndex = 6 Then
        Check prefix & ".compression.noWidth", width.ResultMeta.InternalStatus = rsNotApplicable And Not width.ResultMeta.Calculated
        Check prefix & ".compression.noPsiWarning", InStr(1, width.ResultMeta.ResultComment, "psi_s", vbTextCompare) = 0
        Exit Sub
    End If
    Check prefix & ".width.calculated; " & result.OverallMeta.ResultComment, width.ResultMeta.Calculated
    If Not width.ResultMeta.Calculated Then Exit Sub
    Check prefix & ".width.formulaStatus", width.ResultMeta.InternalStatus = rsSuccess Or width.ResultMeta.InternalStatus = rsCheckFailed
    Check prefix & ".width.standard", width.StandardCode = standard
    If standard = "SP35" Then
        expected = width.SigmaS / width.SteelEs * 1.5 * Sqr(width.SP35ReinforcementRadius / 10#) * 10#
        Dim data As CSP35CrackData, rows As Variant
        Set data = width.SP35Data
        Check prefix & ".groups.present", data.GroupCount > 1
        rows = data.CandidateRows(width.SP35CriticalCandidate)
        If caseIndex = 2 Then Check prefix & ".multiRow.present", UBound(rows, 1) > 1
        Check prefix & ".noSP63Warning", InStr(1, width.ResultMeta.ResultComment, "psi_s", vbTextCompare) = 0
    Else
        expected = width.Phi1 * width.Phi2 * width.Phi3 * width.PsiS * width.SigmaS / width.SteelEs * width.CrackSpacing
        Check prefix & ".psi.bounded", width.PsiS >= 0# And width.PsiS <= 1#
    End If
    Check prefix & ".width.independentFormula; delta=" & CStr(width.CrackWidth - expected), Abs(width.CrackWidth - expected) < 0.000000001
    Set region = width.InteractionRegion
    Check prefix & ".region.present", Not region Is Nothing
    If Not region Is Nothing Then
        Check prefix & ".region.sameArea", Abs(region.Area - width.Abt) < 0.000001
        Check prefix & ".region.areaBounds", region.Area > 0# And region.Area <= 160000# + 0.000001
        Check prefix & ".region.owner", region.OwnerID = Mid$(prefix, Len("E2E." & standard & ".") + 1)
    End If
    If caseIndex = 1 Then
        Check prefix & ".central.tension", width.CentralTensionBranch
        Check prefix & ".central.area", Abs(width.Abt - 160000#) < 0.000001
        Check prefix & ".central.sigma", Abs(width.SigmaS - 500000# / (8# * 3.14159265358979 * 25# ^ 2 / 4#)) < 0.000001
    End If
End Sub

' ДЛЯ ТЕСТОВ: законченные N/M/N+M/Mx+My/N+Mx+My и контроль сжатия проходят
' через batch, State/Search, обе ширины и writers. Переключение стандарта не
' меняет прочность/несущую/устойчивость; снимки сохраняются только в test-copy.
Public Function RunSP35EndToEndTests() As String
    Dim sheet As Object, settingsRange As Object, originalSettings As String, values As Variant
    Dim section As CSectionModel, provider As CMaterialModelProvider, settings As CSystemSettingsReader, units As CUnitSystem
    Dim profiles As CCalculationProfileCatalog, profile As CCalculationProfile, spec As CMaterialModelSpec
    Dim batch As CBatchSectionCalculator, result As CCombinationResult, state As CSectionStateResult
    Dim standard As Variant, i As Long, number As Long, description As String, prefix As String
    Dim baseline(1 To 6, 1 To 7) As Variant, names As Variant, forces As Variant, momentsX As Variant, momentsY As Variant
    Dim summary As CCrackSummaryWriter, writer As CNDMResultsWriter, batchWriter As CBatchResultWriter, anchor As Object, presentationUnits As CUnitSystem
    On Error GoTo Failed
    mPassed = 0: mFailed = 0: mReport = vbNullString
    originalSettings = ThisWorkbook.Names.Item("rngSystemSettings").RefersTo
    values = ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange.Value2
    Set sheet = ThisWorkbook.Worksheets.Add: sheet.Name = "SP35_E2E_CONFIG"
    Set settingsRange = sheet.Cells(1, 1).Resize(UBound(values, 1), UBound(values, 2)): settingsRange.Value2 = values
    ThisWorkbook.Names.Item("rngSystemSettings").RefersTo = "='" & sheet.Name & "'!" & settingsRange.Address
    SetWorkbookTestSetting settingsRange, "Calculation.ZeroMomentPerDepth", 0#
    SetWorkbookTestSetting settingsRange, "Solver.ToleranceN", 0.1
    SetWorkbookTestSetting settingsRange, "Solver.ToleranceMx", 10#
    SetWorkbookTestSetting settingsRange, "Solver.ToleranceMy", 10#
    SetWorkbookTestSetting settingsRange, "SLS.Crack.Allowable", 0.3
    SetWorkbookTestSetting settingsRange, "SLS.Crack.SP35.NeighborRatioLimit", 0.2
    SetWorkbookTestSetting settingsRange, "Stability.Code", "SP63"
    SetWorkbookTestSetting settingsRange, "Stability.ElementLength", 1000#
    SetWorkbookTestSetting settingsRange, "Stability.Mu1", 1#
    SetWorkbookTestSetting settingsRange, "Stability.Mu2", 1#
    Set section = EndToEndSection: Set provider = PipelineProvider
    Set units = New CUnitSystem: units.InitializeDefaults
    Set profiles = New CCalculationProfileCatalog: profiles.LoadFromWorkbook ThisWorkbook
    Set profile = profiles.ProfileById("PR1")
    profile.Initialize "PR1", "E2E", "Полный интеграционный расчет в тестовой копии.", True, True, True, True
    Set spec = New CMaterialModelSpec: spec.Initialize "ULS(I)", "TwoLine", "Ignore", "TwoLine": profile.SetStrengthSpec spec
    profile.SetCrackedStateSpec PipelineSpec()
    spec.Initialize "SLS(II)", "TwoLine", "UseDiagram", "TwoLine": profile.SetCrackInitiationSpec spec
    profile.SetStabilityValueSet "ULS(I)": profile.SetVisualization "CrackedState", "Stress"
    names = Array("N", "M", "N_M", "Mx_My", "N_Mx_My", "N_COMPRESSION")
    forces = Array(500000#, 0#, -100000#, 0#, -100000#, -200000#)
    momentsX = Array(0#, 120000000#, 120000000#, 80000000#, 80000000#, 0#)
    momentsY = Array(0#, 0#, 0#, 60000000#, 60000000#, 0#)
    For Each standard In Array("SP63", "SP35")
        SetWorkbookTestSetting settingsRange, "SLS.Crack.Code", CStr(standard)
        Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
        Set batch = New CBatchSectionCalculator: batch.Initialize section, provider
        Set batch.ProfileCatalog = profiles: batch.ApplySettings settings, units
        For i = 1 To 6
            batch.AddCombination CStr(names(i - 1)), CDbl(forces(i - 1)), CDbl(momentsX(i - 1)), CDbl(momentsY(i - 1)), "PR1", "Контроль " & CStr(names(i - 1)), "Auto"
        Next i
        batch.Execute
        For i = 1 To 6
            prefix = "E2E." & CStr(standard) & "." & CStr(names(i - 1)): Set result = batch.ResultAt(i)
            mReport = mReport & "RESULT: " & prefix & "; status=" & result.Status & "; comment=" & result.OverallMeta.ResultComment & vbCrLf
            Set state = result.StateRepository.FindState(sstCrackedState)
            Check prefix & ".current.present", Not state Is Nothing
            If Not state Is Nothing Then
                mReport = mReport & "STATE: " & prefix & "; N=" & CStr(state.TargetN) & "; Mx=" & CStr(state.TargetMx) & "; My=" & CStr(state.TargetMy) & _
                    "; eps0=" & CStr(state.Epsilon0) & "; kx=" & CStr(state.KappaX) & "; ky=" & CStr(state.KappaY) & vbCrLf
                Check prefix & ".current.converged", state.Converged And state.WithinPhysicalRange
                Check prefix & ".equilibrium.N", Abs(state.Nint - CDbl(forces(i - 1))) <= 0.100001
                ' При сжатии штатная устойчивость добавляет эксцентриситет и
                ' увеличивает моменты: проверяем равновесие принятой DesignLoad,
                ' а не требуем от последующего State исходные моменты таблицы.
                Check prefix & ".equilibrium.Mx", Abs(state.Mxint - batch.CrackReferenceMx(i)) <= 10.000001
                Check prefix & ".equilibrium.My", Abs(state.Myint - batch.CrackReferenceMy(i)) <= 10.000001
                Check prefix & ".target.Mx", Abs(state.TargetMx - batch.CrackReferenceMx(i)) < 0.000001
                Check prefix & ".target.My", Abs(state.TargetMy - batch.CrackReferenceMy(i)) < 0.000001
                If i = 1 Then Check prefix & ".central.noNA", Abs(state.KappaX) < 0.0000000001 And Abs(state.KappaY) < 0.0000000001
            End If
            CheckEndToEndWidth prefix, result, section, CStr(standard), i
            Check prefix & ".strength.success", result.DirectStateMeta.InternalStatus = rsSuccess
            Check prefix & ".capacity.success", result.CapacityMeta.InternalStatus = rsSuccess
            If standard = "SP63" Then
                baseline(i, 1) = result.DirectStateMeta.InternalStatus: baseline(i, 2) = result.CapacityMeta.InternalStatus
                baseline(i, 3) = result.StrengthResult.Capacity.LambdaCapacity
                baseline(i, 4) = result.StabilityMeta.InternalStatus: baseline(i, 5) = result.StabilityResult.Ncr1
                baseline(i, 6) = result.StabilityResult.Ncr2: baseline(i, 7) = result.StabilityResult.DesignN
            Else
                Check prefix & ".strength.unchanged", result.DirectStateMeta.InternalStatus = baseline(i, 1)
                Check prefix & ".capacity.statusUnchanged", result.CapacityMeta.InternalStatus = baseline(i, 2)
                CheckNear prefix & ".capacity.lambdaUnchanged", result.StrengthResult.Capacity.LambdaCapacity, CDbl(baseline(i, 3))
                Check prefix & ".stability.statusUnchanged", result.StabilityMeta.InternalStatus = baseline(i, 4)
                CheckNear prefix & ".stability.Ncr1Unchanged", result.StabilityResult.Ncr1, CDbl(baseline(i, 5))
                CheckNear prefix & ".stability.Ncr2Unchanged", result.StabilityResult.Ncr2, CDbl(baseline(i, 6))
                CheckNear prefix & ".stability.designNUnchanged", result.StabilityResult.DesignN, CDbl(baseline(i, 7))
            End If
        Next i
        ' Выводим обычные пользовательские единицы книги. Solver-эталоны выше
        ' используют отдельный INTERNAL-адаптер и не зависят от этого оформления.
        Set presentationUnits = New CUnitSystem: presentationUnits.LoadFromSettings settings
        Set summary = New CCrackSummaryWriter: summary.WriteSummary ThisWorkbook, batch, presentationUnits
        Set batchWriter = New CBatchResultWriter: batchWriter.WriteSummary ThisWorkbook, batch, presentationUnits
        Set writer = New CNDMResultsWriter: writer.WriteResults ThisWorkbook, section, provider, batch, presentationUnits
        Set anchor = ThisWorkbook.Names.Item("rngCrackSummaryAnchor").RefersToRange
        For i = 1 To 6
            Check "E2E.writer." & CStr(standard) & "." & CStr(i) & ".comment", CStr(anchor.Offset(i - 1, 1).Value2) = batch.ResultAt(i).CrackSummaryMeta.ResultComment
        Next i
        ThisWorkbook.SaveCopyAs ThisWorkbook.Path & "\E2E_" & CStr(standard) & ".xlsm"
    Next standard
    GoTo Restore
Failed:
    number = Err.Number: description = Err.Description
    Check "runtime: " & CStr(number) & "; " & description, False
Restore:
    On Error Resume Next
    ThisWorkbook.Names.Item("rngSystemSettings").RefersTo = originalSettings
    If Not sheet Is Nothing Then sheet.Delete
    On Error GoTo 0
    mReport = mReport & "TOTAL_SP35_END_TO_END: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
    RunSP35EndToEndTests = mReport
End Function
