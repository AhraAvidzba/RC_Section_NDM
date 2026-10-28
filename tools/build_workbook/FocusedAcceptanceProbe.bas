' ==========================================================================
' ДЛЯ ТЕСТОВ — вспомогательный фрагмент частной приёмки; в выпуск не включается.
' ==========================================================================

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сохраняет ссылки на фактически рассчитанные модель и пакет частной книги для независимых проверок после штатного workflow.
Public Sub TestCaptureFocusedCalculation(ByVal section As CSectionModel, ByVal batch As CBatchSectionCalculator)
    Set mFocusedSection = section
    Set mFocusedBatch = batch
End Sub
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает состав выбранных нагрузок и формы без решения НДС, поиска предела и записи Results.
Public Function TestDescribeFocusedAcceptance(ByVal shapeName As String, ByVal groupName As String) As String
    Dim oldRebars As Variant
    oldRebars = ThisWorkbook.Worksheets("Сечения").UsedRange.Formula
    On Error GoTo Failed
    Dim oldSystem As Variant, settings As CSystemSettingsReader, units As CUnitSystem
    Dim section As CSectionModel, cases As Collection, item As Variant, output As String
    oldSystem = ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange.Formula
    TestFocusedSelectGeometry shapeName
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    Set section = BuildWorkbookSectionModel(ThisWorkbook, settings, units)
    Set cases = TestFocusedSelectedLoads(section, groupName)
    output = "FOCUSED_MANIFEST|shape=" & shapeName & "|group=" & groupName & "|cases=" & CStr(cases.Count) & _
        "|fibers=" & CStr(section.ConcreteCount) & "|rebars=" & CStr(section.RebarCount) & vbCrLf
    For Each item In cases
        output = output & "FOCUSED_INPUT|id=" & CStr(item(0)) & "|N=" & TestFormatNumberInvariant(CDbl(item(1))) & _
            "|Mx=" & TestFormatNumberInvariant(CDbl(item(2))) & "|My=" & TestFormatNumberInvariant(CDbl(item(3))) & _
            "|path=" & CStr(item(4)) & vbCrLf
    Next item
    output = output & "FOCUSED_MANIFEST_COMPLETE: no calculation executed" & vbCrLf
    GoTo Restore
Failed:
    output = output & "RUNTIME ERROR: " & CStr(Err.Number) & "; " & Err.Description
Restore:
    If IsArray(oldRebars) Then ThisWorkbook.Worksheets("Сечения").UsedRange.Formula = oldRebars
    If IsArray(oldSystem) Then ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange.Formula = oldSystem
    TestDescribeFocusedAcceptance = output
End Function
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Выполняет выбранные нагрузки штатным workflow, проверяет состояния и вывод, затем восстанавливает исходные таблицы частной книги.
Public Function TestRunFocusedAcceptanceCases(ByVal shapeName As String, ByVal groupName As String, _
        Optional ByVal selectedLoadName As String = "") As String
    On Error GoTo Failed
    Dim oldRebars As Variant
    oldRebars = ThisWorkbook.Worksheets("Сечения").UsedRange.Formula
    Dim stats As TBatchTestStats, oldSystem As Variant, oldProfiles As Variant, oldLoads As Variant
    Dim systemRange As Object, profileRange As Object, loadRange As Object
    Dim settings As CSystemSettingsReader, units As CUnitSystem, section As CSectionModel
    Dim cases As Collection, item As Variant, values() As Variant, path As String
    Dim start As Double, elapsed As Double, message As String, completed As Long, maxSeconds As Double
    Dim report As String, result As CCombinationResult, meta As Variant, prefix As String
    Set systemRange = ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange
    Set profileRange = ThisWorkbook.Names.Item("rngCalculationProfiles").RefersToRange
    Set loadRange = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange
    oldSystem = systemRange.Formula: oldProfiles = profileRange.Formula: oldLoads = loadRange.Formula
    TestFocusedSelectGeometry shapeName
    TestSetSystemSetting "Load.ReferenceOffsetX", "0"
    TestSetSystemSetting "Load.ReferenceOffsetY", "0"
    TestSetSystemSetting "Calculation.ZeroMomentPerDepth", "0"
    TestSetSystemSetting "General.ExecutionReportEnabled", "Yes"
    TestSetSystemSetting "SLS.Crack.Code", "SP35"
    TestSetSystemSetting "Plot.Enabled", "Yes"
    TestSetSystemSetting "Plot.AutoUpdateAfterCalculation", "Yes"
    modTestPlotConfig.TestSetOutputCombination ThisWorkbook, "Worst"
    TestSetProfileValue "Calculation.Strength.DirectState", "PR3", "Yes"
    TestSetProfileValue "Calculation.Strength.Capacity", "PR3", "Yes"
    TestSetProfileValue "Calculation.Crack.Width", "PR3", "Yes"
    TestSetProfileValue "Calculation.Stability.Enabled", "PR3", "No"
    TestSetProfileValue "Visualization.State", "PR3", "StrengthState"
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    TestSetSystemSetting "Load.ReferenceOffsetX", TestFormatNumberInvariant(10# / units.InputLengthToInternal(1#))
    TestSetSystemSetting "Load.ReferenceOffsetY", TestFormatNumberInvariant(-7# / units.InputLengthToInternal(1#))
    settings.LoadFromWorkbook ThisWorkbook
    Set section = BuildWorkbookSectionModel(ThisWorkbook, settings, units)
    Set cases = TestFocusedSelectedLoads(section, groupName)
    If Len(selectedLoadName) > 0 Then
        Dim selected As New Collection
        For Each item In cases
            If CStr(item(0)) = selectedLoadName Then selected.Add item
        Next item
        If selected.Count <> 1 Then Err.Raise 5, , "Selected focused load was not found uniquely."
        Set cases = selected
    End If
    TestAppendLine stats, "FOCUSED_SETUP|shape=" & shapeName & "|group=" & groupName & "|mode=" & _
        settings.GetRequiredString("General.DiagramExtension") & "|cases=" & CStr(cases.Count) & _
        "|fibers=" & CStr(section.ConcreteCount) & "|rebars=" & CStr(section.RebarCount) & "|offsetX=10|offsetY=-7|profile=PR3|code=SP35"
    For Each item In cases
        prefix = "focused." & shapeName & "." & CStr(item(0))
        path = CStr(item(4)): If path <> "Auto" Then path = ChrW$(&H3BB) & "*" & path
        ReDim values(1 To loadRange.Rows.Count - 1, 1 To loadRange.Columns.Count)
        values(1, 1) = CStr(item(0))
        values(1, 2) = CDbl(item(1)) / units.InputForceToInternal(1#)
        values(1, 3) = CDbl(item(2)) / units.InputMomentMxToInternal(1#)
        values(1, 4) = CDbl(item(3)) / units.InputMomentMyToInternal(1#)
        values(1, 5) = "PR3": values(1, 6) = path
        loadRange.Offset(1, 0).Resize(loadRange.Rows.Count - 1, loadRange.Columns.Count).Value2 = values
        Set mFocusedBatch = Nothing: Set mFocusedSection = Nothing
        ResetSectionEquilibriumSolveCount
        start = Timer
        If CDbl(item(1)) = 0# And CDbl(item(2)) = 0# And CDbl(item(3)) = 0# Then
            TestFocusedCheckZeroInput stats, prefix
        Else
            message = TestFixtureRunSectionWorkflow(ThisWorkbook, False)
        End If
        elapsed = Timer - start: If elapsed < 0# Then elapsed = elapsed + 86400#
        If elapsed > maxSeconds Then maxSeconds = elapsed
        TestAssertTrue stats, prefix & ".within10Seconds", elapsed <= 10#
        TestAppendLine stats, "FOCUSED_TIME|shape=" & shapeName & "|id=" & CStr(item(0)) & _
            "|elapsedSec=" & TestFormatNumberInvariant(elapsed) & "|solverCalls=" & CStr(SectionEquilibriumSolveCount())
        If Not mFocusedBatch Is Nothing Then
            TestAssertTrue stats, prefix & ".oneCombination", mFocusedBatch.Count = 1
            TestAssertTrue stats, prefix & ".preparationReleased", mFocusedBatch.TestRunPreparationReleased
            ' ULS и SLS имеют отдельные точные кэши; устойчивость в этом профиле отключена.
            TestAssertTrue stats, prefix & ".mechanicalNoRepeatedMaterialPreparation", _
                mFocusedBatch.TestMechanicalPreparationBuildCount > 0 And mFocusedBatch.TestMechanicalPreparationBuildCount <= 2
            TestAssertTrue stats, prefix & ".crackDomainPreparedOnce", mFocusedBatch.TestCrackDomainBuildCount <= 1
            TestAppendLine stats, "FOCUSED_PREPARATION|id=" & CStr(item(0)) & _
                "|mechanicalBuilds=" & CStr(mFocusedBatch.TestMechanicalPreparationBuildCount) & _
                "|crackDomainBuilds=" & CStr(mFocusedBatch.TestCrackDomainBuildCount)
            Set result = mFocusedBatch.ResultAt(1)
            TestAudit03CheckResultComments stats, mFocusedBatch, 1
            Dim cx As Double, cy As Double, area As Double
            mFocusedSection.ConcreteMeshCentroid cx, cy, area
            TestAudit03CheckCapacityPath stats, result.StrengthResult.Capacity, prefix, CStr(item(4)), _
                Array(item(1), item(2), item(3)), cx, cy, True
            TestAudit03CheckFormationPath stats, result.CrackResult.Formation, prefix, path, Array(item(1), item(2), item(3))
            TestFocusedCheckStates stats, result, settings, units, prefix
            TestAppendLine stats, TestFocusedComparisonRecord(result, CStr(item(0)))
            For Each meta In Array(result.DirectStateMeta, result.CapacityMeta, result.CrackFormationMeta, _
                    result.CrackCurrentStateMeta, result.CrackWidthMeta, result.LongitudinalCrackMeta, result.OverallMeta)
                TestAssertTrue stats, prefix & ".noInputOrInternalError." & CStr(meta.ResultKind), _
                    meta.InternalStatus <> rsInvalidInput And meta.InternalStatus <> rsInvalidConfiguration And _
                    meta.InternalStatus <> rsInternalError
                If meta.InternalStatus = rsNumericalFailure Then TestAppendLine stats, _
                    "FOCUSED_UNRESOLVED|shape=" & shapeName & "|id=" & CStr(item(0)) & _
                    "|kind=" & CStr(meta.ResultKind) & "|reason=" & meta.ResultComment
            Next meta
            TestFocusedCheckOutputRow stats, "rngStrengthSummaryAnchor", CStr(item(0)), prefix
            TestFocusedCheckOutputRow stats, "rngCrackSummaryAnchor", CStr(item(0)), prefix
            TestAppendLine stats, "FOCUSED_RESULT|id=" & CStr(item(0)) & "|overall=" & result.Status & _
                "|capacity=" & result.StrengthResult.Capacity.Status & "|cracks=" & result.CrackSummaryStatus
        ElseIf CStr(item(0)) <> "zero" Then
            TestAssertTrue stats, prefix & ".capturedCalculation", False
        End If
        completed = completed + 1
        TestAudit03SaveMatrixProgress stats.Report
    Next item
    TestAssertTrue stats, "focused.complete", completed = cases.Count
    TestAppendLine stats, "TOTAL_FOCUSED_ACCEPTANCE: shape=" & shapeName & "; group=" & groupName & _
        "; cases=" & CStr(completed) & "; maxSeconds=" & TestFormatNumberInvariant(maxSeconds) & _
        "; passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    GoTo Restore
Failed:
    stats.Failed = stats.Failed + 1
    TestAppendLine stats, "FAIL: focused.runtime; " & CStr(Err.Number) & "; " & Err.Description
Restore:
    On Error GoTo RestoreFailed
    If IsArray(oldRebars) Then ThisWorkbook.Worksheets("Сечения").UsedRange.Formula = oldRebars
    If IsArray(oldSystem) Then systemRange.Formula = oldSystem
    If IsArray(oldProfiles) Then profileRange.Formula = oldProfiles
    If IsArray(oldLoads) Then loadRange.Formula = oldLoads
    Set mFocusedBatch = Nothing: Set mFocusedSection = Nothing
    TestRunFocusedAcceptanceCases = stats.Report
    Exit Function
RestoreFailed:
    TestRunFocusedAcceptanceCases = stats.Report & "RESTORE ERROR: " & CStr(Err.Number) & "; " & Err.Description
End Function
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сериализует typed-статусы, коды и численные результаты для сравнения сборок без подмены статуса текстом.
Private Function TestFocusedComparisonRecord(ByVal result As CCombinationResult, ByVal id As String) As String
    Dim output As String, meta As Variant, numbers As Variant, i As Long
    output = "FOCUSED_COMPARE|id=" & id & "|path=" & result.StrengthResult.Capacity.PathResolved
    For Each meta In Array(result.DirectStateMeta, result.CapacityMeta, result.CrackFormationMeta, _
            result.CrackCurrentStateMeta, result.CrackWidthMeta, result.LongitudinalCrackMeta, result.OverallMeta)
        output = output & "|meta" & CStr(meta.ResultKind) & "=" & CStr(meta.InternalStatus) & ":" & CStr(meta.ResultCode)
    Next meta
    numbers = Array(result.StrengthResult.Capacity.LambdaCapacity, result.StrengthResult.Capacity.NUltimate, _
        result.StrengthResult.Capacity.MxUltimate, result.StrengthResult.Capacity.MyUltimate, _
        result.CrackResult.Formation.LambdaCrc, result.CrackResult.Formation.Ncrc, result.CrackResult.Formation.Mcrc, _
        result.CrackResult.Width.CrackWidth, result.CrackResult.Width.Utilization, result.CrackResult.Width.CrackSpacing, _
        result.CrackResult.Longitudinal.Utilization, result.GoverningUtil)
    For i = LBound(numbers) To UBound(numbers)
        output = output & "|number" & CStr(i) & "=" & Replace$(Trim$(Str$(CDbl(numbers(i)))), ",", ".")
    Next i
    TestFocusedComparisonRecord = output
End Function
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Запускает одно именованное сочетание встроенной формы через общий focused workflow и его проверки.
Public Function TestRunFocusedSingleLoad(ByVal shapeName As String, ByVal loadName As String) As String
    TestRunFocusedSingleLoad = TestRunFocusedAcceptanceCases(shapeName, "Generated", loadName)
End Function
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Выбирает сохранённый импорт либо зарегистрированную форму и задаёт независимый контрольный набор армирования этой формы.
Private Sub TestFocusedSelectGeometry(ByVal shapeName As String)
    Select Case shapeName
        Case "SavedWorkbook": TestSetSystemSetting "Geometry.Source", "AutoCAD"
        Case "Rectangles", "Circle", "PointedSection", "HollowRectangle"
            TestSetSystemSetting "Geometry.Source", "Generated"
            TestSetSystemSetting "Geometry.Type", shapeName
            ' Общая таблица не хранит четыре независимых набора армирования.
            ' Перед каждой формой задаём явный контрольный набор прежних исходных данных.
            TestFixtureResetReinforcement shapeName
        Case Else: Err.Raise 5, "FocusedSelectGeometry", "Unknown focused geometry: " & shapeName
    End Select
End Sub
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Отбирает нормированные нагрузки по устойчивым именам, отдельно задаёт шесть LoadPath и проверяет полноту группы.
Private Function TestFocusedSelectedLoads(ByVal section As CSectionModel, ByVal groupName As String) As Collection
    Dim result As New Collection, item As Variant, path As Variant, names As Variant
    names = Array("axialSmall-1", "axialSmall1", "Mx-1", "Mx1", "My-1", "My1", "mixedC1_-1", "mixedT-1_1")
    If groupName = "Paths" Then
        For Each item In TestAudit03NormalizedLoads(section, "Light")
            If CStr(item(0)) = "mixedC1_-1" Then
                For Each path In Array("Auto", "Mx", "My", "Mxy", "N", "NMxy")
                    result.Add Array("path" & CStr(path), item(1), item(2), item(3), CStr(path))
                Next path
                Exit For
            End If
        Next item
        If result.Count <> 6 Then Err.Raise 5, , "Incomplete LoadPath selection."
    ElseIf groupName = "Generated" Or groupName = "Saved" Then
        If groupName = "Saved" Then names = Array("axialSmall-1", "axialSmall1", "Mx-1", "Mx1", "My-1", "My1", _
            "mixedC1_-1", "mixedT-1_1", "zero", "tinyMx1", "biaxial1_-1_10", "biaxial1_-1_1000000", "biaxial1_-1_" & CStr(0.000001))
        For Each item In TestAudit03NormalizedLoads(section, "Light")
            If TestFocusedNameIn(CStr(item(0)), names) Then result.Add Array(item(0), item(1), item(2), item(3), "Auto")
        Next item
        If groupName = "Saved" Then
            For Each item In TestAudit03NormalizedLoads(section, "Stress")
                If TestFocusedNameIn(CStr(item(0)), Array("axialC" & CStr(0.95), "axialC1", "axialC" & CStr(1.05), "axialC2", _
                        "overloadMixed1", "overloadMixed1000000")) Then _
                    result.Add Array(item(0), item(1), item(2), item(3), "Auto")
            Next item
            result.Add Array("N500", -500# * 9806.65, 0#, 0#, "Auto")
            result.Add Array("N1000", -1000# * 9806.65, 0#, 0#, "Auto")
            result.Add Array("N2098Plateau", -2098.694 * 9806.65, 0#, 0#, "Auto")
            result.Add Array("N500TinyMoment", -500# * 9806.65, 1#, 0#, "Auto")
            If result.Count <> 23 Then Err.Raise 5, , "Incomplete saved-load selection: " & CStr(result.Count)
        Else
            If result.Count <> 8 Then Err.Raise 5, , "Incomplete generated-load selection."
        End If
    Else
        Err.Raise 5, "FocusedSelectedLoads", "Unknown focused group: " & groupName
    End If
    Set TestFocusedSelectedLoads = result
End Function
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет точное совпадение имени нагрузки с одним из заданных имён, сохраняя исходный регистр отбора.
Private Function TestFocusedNameIn(ByVal name As String, ByVal names As Variant) As Boolean
    Dim item As Variant
    For Each item In names
        If name = CStr(item) Then TestFocusedNameIn = True: Exit Function
    Next item
End Function
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Требует штатного пропуска полностью нулевой нагрузки без solver и изменения прежних результатов или исходной таблицы.
Private Sub TestFocusedCheckZeroInput(ByRef stats As TBatchTestStats, ByVal prefix As String)
    Dim resultRange As Object, loadRange As Object
    Set resultRange = ThisWorkbook.Worksheets("Results").UsedRange
    Set loadRange = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange
    Dim resultsBefore As Variant, loadsBefore As Variant
    resultsBefore = resultRange.Formula: loadsBefore = loadRange.Formula
    Dim message As String
    message = TestFixtureRunSectionWorkflow(ThisWorkbook, False)
    TestAssertTrue stats, prefix & ".zeroSkippedWithMessage", _
        InStr(1, message, "Расчёт не запускался:", vbBinaryCompare) = 1 And _
        InStr(1, message, "Прежние результаты сохранены.", vbBinaryCompare) > 0
    TestAssertTrue stats, prefix & ".zeroNoSolverCalls", SectionEquilibriumSolveCount() = 0
    TestAssertTrue stats, prefix & ".zeroNoCapturedCalculation", mFocusedBatch Is Nothing
    TestAssertTrue stats, prefix & ".zeroResultsPreserved", TestFocusedCellMatricesEqual(resultsBefore, resultRange.Formula)
    TestAssertTrue stats, prefix & ".zeroLoadsPreserved", TestFocusedCellMatricesEqual(loadsBefore, loadRange.Formula)
End Sub
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Побуквенно сравнивает формулы и точные значения двух прямоугольных снимков, включая пустые ячейки и ошибки Excel.
Private Function TestFocusedCellMatricesEqual(ByVal expected As Variant, ByVal actual As Variant) As Boolean
    If LBound(expected, 1) <> LBound(actual, 1) Or UBound(expected, 1) <> UBound(actual, 1) Or _
            LBound(expected, 2) <> LBound(actual, 2) Or UBound(expected, 2) <> UBound(actual, 2) Then Exit Function
    Dim row As Long, column As Long, expectedValue As Variant, actualValue As Variant
    For row = LBound(expected, 1) To UBound(expected, 1)
        For column = LBound(expected, 2) To UBound(expected, 2)
            expectedValue = expected(row, column): actualValue = actual(row, column)
            If VarType(expectedValue) <> VarType(actualValue) Then Exit Function
            If IsError(expectedValue) Or VarType(expectedValue) = vbString Then
                If StrComp(CStr(expectedValue), CStr(actualValue), vbBinaryCompare) <> 0 Then Exit Function
            ElseIf Not IsNull(expectedValue) Then
                If expectedValue <> actualValue Then Exit Function
            End If
        Next column
    Next row
    TestFocusedCellMatricesEqual = True
End Function
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Независимо интегрирует сохранённые диаграммы и плоскость деформаций, проверяя равновесие в WCS и относительно бетонного центра.
Private Sub TestFocusedCheckStates(ByRef stats As TBatchTestStats, ByVal result As CCombinationResult, _
        ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem, ByVal prefix As String)
    Dim i As Long, state As CSectionStateResult, probe As CSectionSolver
    Dim n As Double, mx As Double, my As Double, cx As Double, cy As Double, area As Double
    Dim toleranceN As Double, toleranceMx As Double, toleranceMy As Double, extensionEnabled As Boolean
    toleranceN = Abs(units.InputForceToInternal(settings.GetDouble("Solver.ToleranceN")))
    toleranceMx = Abs(units.InputMomentMxToInternal(settings.GetDouble("Solver.ToleranceMx")))
    toleranceMy = Abs(units.InputMomentMyToInternal(settings.GetDouble("Solver.ToleranceMy")))
    extensionEnabled = settings.GetRequiredBoolean("General.DiagramExtension")
    mFocusedSection.ConcreteMeshCentroid cx, cy, area
    For i = 1 To result.StateRepository.StateCount
        Set state = result.StateRepository.StateAt(i)
        TestAppendLine stats, "FOCUSED_STATE|id=" & prefix & "|type=" & state.StateTypeText & _
            "|converged=" & CStr(state.Converged) & "|physical=" & CStr(state.WithinPhysicalRange) & _
            "|extension=" & CStr(state.ExtensionUsed) & "|status=" & state.Status & _
            "|N=" & TestFocusedExactNumber(state.TargetN) & "|Mx=" & TestFocusedExactNumber(state.TargetMx) & _
            "|My=" & TestFocusedExactNumber(state.TargetMy) & "|eps0=" & TestFocusedExactNumber(state.Epsilon0) & _
            "|kx=" & TestFocusedExactNumber(state.KappaX) & "|ky=" & TestFocusedExactNumber(state.KappaY)
        If state.Converged Then
            TestAssertTrue stats, prefix & ".snapshot." & CStr(i), state.HasMaterialSnapshot
            Set probe = New CSectionSolver
            probe.EvaluateStrainPlane mFocusedSection, state.ConcreteDiagramSnapshot, state.SteelDiagramSnapshot, _
                state.Epsilon0, state.KappaX, state.KappaY
            TestAssertClose stats, prefix & ".originalN." & CStr(i), probe.Nint, state.TargetN, toleranceN
            TestAssertClose stats, prefix & ".originalMx." & CStr(i), probe.Mxint, state.TargetMx, toleranceMx
            TestAssertClose stats, prefix & ".originalMy." & CStr(i), probe.Myint, state.TargetMy, toleranceMy
            probe.NumericalEfforts n, mx, my
            TestAssertClose stats, prefix & ".centerN." & CStr(i), n, state.TargetN, toleranceN
            TestAssertClose stats, prefix & ".centerMx." & CStr(i), mx, state.TargetMx - cy * state.TargetN, toleranceMx
            TestAssertClose stats, prefix & ".centerMy." & CStr(i), my, state.TargetMy - cx * state.TargetN, toleranceMy
            TestAssertTrue stats, prefix & ".physicalStatus." & CStr(i), _
                state.WithinPhysicalRange Or state.InternalStatus = rsCheckFailed
        Else
            TestAssertTrue stats, prefix & ".failureNotPhysical." & CStr(i), Not state.WithinPhysicalRange
        End If
        If Not extensionEnabled Then TestAssertTrue stats, prefix & ".noHiddenExtension." & CStr(i), Not state.ExtensionUsed
    Next i
End Sub
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сохраняет все биты Double в F64-записи через два слова для независимого контроля без десятичной потери точности.
Private Function TestFocusedExactNumber(ByVal value As Double) As String
    Dim number As TFocusedDouble, words As TFocusedWords
    number.Value = value: LSet words = number
    TestFocusedExactNumber = "F64:" & Right$("00000000" & Hex$(words.High), 8) & Right$("00000000" & Hex$(words.Low), 8)
End Function
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет, что первая строка итоговой таблицы относится к фактически рассчитанному сочетанию.
Private Sub TestFocusedCheckOutputRow(ByRef stats As TBatchTestStats, ByVal rangeName As String, _
        ByVal id As String, ByVal prefix As String)
    Dim anchor As Object
    Set anchor = ThisWorkbook.Names.Item(rangeName).RefersToRange.Cells(1, 1)
    TestAssertEquals stats, prefix & ".output." & rangeName, CStr(anchor.Value2), id
End Sub
