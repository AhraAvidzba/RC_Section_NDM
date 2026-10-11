' ДЛЯ ТЕСТОВ: добавляется только в modTestBatchCalculation независимой копии.
' Использует штатную кнопку расчета, не меняя solver, материалы или допуски.

Public Sub CaptureFocusedCalculation(ByVal section As CSectionModel, ByVal batch As CBatchSectionCalculator)
    Set mFocusedSection = section
    Set mFocusedBatch = batch
End Sub

' ДЛЯ ТЕСТОВ: фиксирует состав без запуска НДС, Search и записи Results.
Public Function DescribeFocusedAcceptance(ByVal shapeName As String, ByVal groupName As String) As String
    On Error GoTo Failed
    Dim oldSystem As Variant, settings As CSystemSettingsReader, units As CUnitSystem
    Dim section As CSectionModel, cases As Collection, item As Variant, output As String
    oldSystem = ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange.Formula
    FocusedSelectGeometry shapeName
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    Set section = BuildWorkbookSectionModel(ThisWorkbook, settings, units)
    Set cases = FocusedSelectedLoads(section, groupName)
    output = "FOCUSED_MANIFEST|shape=" & shapeName & "|group=" & groupName & "|cases=" & CStr(cases.Count) & _
        "|fibers=" & CStr(section.ConcreteCount) & "|rebars=" & CStr(section.RebarCount) & vbCrLf
    For Each item In cases
        output = output & "FOCUSED_INPUT|id=" & CStr(item(0)) & "|N=" & FormatNumberInvariant(CDbl(item(1))) & _
            "|Mx=" & FormatNumberInvariant(CDbl(item(2))) & "|My=" & FormatNumberInvariant(CDbl(item(3))) & _
            "|path=" & CStr(item(4)) & vbCrLf
    Next item
    output = output & "FOCUSED_MANIFEST_COMPLETE: no calculation executed" & vbCrLf
    GoTo Restore
Failed:
    output = output & "RUNTIME ERROR: " & CStr(Err.Number) & "; " & Err.Description
Restore:
    If IsArray(oldSystem) Then ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange.Formula = oldSystem
    DescribeFocusedAcceptance = output
End Function

' ДЛЯ ТЕСТОВ: одна форма и выбранные нагрузки, без матрицы всех настроек.
Public Function RunFocusedAcceptanceCases(ByVal shapeName As String, ByVal groupName As String, _
        Optional ByVal selectedLoadName As String = "") As String
    On Error GoTo Failed
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
    FocusedSelectGeometry shapeName
    SetSystemSetting "Load.ReferenceOffsetX", "0"
    SetSystemSetting "Load.ReferenceOffsetY", "0"
    SetSystemSetting "Calculation.ZeroMomentPerDepth", "0"
    SetSystemSetting "General.ExecutionReportEnabled", "Yes"
    SetSystemSetting "SLS.Crack.Code", "SP35"
    SetSystemSetting "Plot.Enabled", "Yes"
    SetSystemSetting "Plot.AutoUpdateAfterCalculation", "Yes"
    SetSystemSetting "Plot.LoadCase", "Worst"
    SetProfileValue "Calculation.Strength.DirectState", "PR3", "Yes"
    SetProfileValue "Calculation.Strength.Capacity", "PR3", "Yes"
    SetProfileValue "Calculation.Crack.Width", "PR3", "Yes"
    SetProfileValue "Calculation.Stability.Enabled", "PR3", "No"
    SetProfileValue "Visualization.State", "PR3", "StrengthState"
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    SetSystemSetting "Load.ReferenceOffsetX", FormatNumberInvariant(10# / units.InputLengthToInternal(1#))
    SetSystemSetting "Load.ReferenceOffsetY", FormatNumberInvariant(-7# / units.InputLengthToInternal(1#))
    settings.LoadFromWorkbook ThisWorkbook
    Set section = BuildWorkbookSectionModel(ThisWorkbook, settings, units)
    Set cases = FocusedSelectedLoads(section, groupName)
    If Len(selectedLoadName) > 0 Then
        Dim selected As New Collection
        For Each item In cases
            If CStr(item(0)) = selectedLoadName Then selected.Add item
        Next item
        If selected.Count <> 1 Then Err.Raise 5, , "Selected focused load was not found uniquely."
        Set cases = selected
    End If
    AppendLine stats, "FOCUSED_SETUP|shape=" & shapeName & "|group=" & groupName & "|mode=" & _
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
            FocusedCheckZeroInput stats, prefix
        Else
            message = RunSectionCalculationForWorkbook(ThisWorkbook, False)
        End If
        elapsed = Timer - start: If elapsed < 0# Then elapsed = elapsed + 86400#
        If elapsed > maxSeconds Then maxSeconds = elapsed
        AssertTrue stats, prefix & ".within10Seconds", elapsed <= 10#
        AppendLine stats, "FOCUSED_TIME|shape=" & shapeName & "|id=" & CStr(item(0)) & _
            "|elapsedSec=" & FormatNumberInvariant(elapsed) & "|solverCalls=" & CStr(SectionEquilibriumSolveCount())
        If Not mFocusedBatch Is Nothing Then
            AssertTrue stats, prefix & ".oneCombination", mFocusedBatch.Count = 1
            AssertTrue stats, prefix & ".preparationReleased", mFocusedBatch.RunPreparationReleased
            ' ULS и SLS имеют отдельные точные кэши; устойчивость в этом профиле отключена.
            AssertTrue stats, prefix & ".mechanicalNoRepeatedMaterialPreparation", _
                mFocusedBatch.MechanicalPreparationBuildCount > 0 And mFocusedBatch.MechanicalPreparationBuildCount <= 2
            AssertTrue stats, prefix & ".crackDomainPreparedOnce", mFocusedBatch.CrackDomainBuildCount <= 1
            AppendLine stats, "FOCUSED_PREPARATION|id=" & CStr(item(0)) & _
                "|mechanicalBuilds=" & CStr(mFocusedBatch.MechanicalPreparationBuildCount) & _
                "|crackDomainBuilds=" & CStr(mFocusedBatch.CrackDomainBuildCount)
            Set result = mFocusedBatch.ResultAt(1)
            Audit03CheckResultComments stats, mFocusedBatch, 1
            Dim cx As Double, cy As Double, area As Double
            mFocusedSection.ConcreteMeshCentroid cx, cy, area
            Audit03CheckCapacityPath stats, result.StrengthResult.Capacity, prefix, CStr(item(4)), _
                Array(item(1), item(2), item(3)), cx, cy, True
            Audit03CheckFormationPath stats, result.CrackResult.Formation, prefix, path, Array(item(1), item(2), item(3))
            FocusedCheckStates stats, result, settings, units, prefix
            For Each meta In Array(result.DirectStateMeta, result.CapacityMeta, result.CrackFormationMeta, _
                    result.CrackCurrentStateMeta, result.CrackWidthMeta, result.LongitudinalCrackMeta, result.OverallMeta)
                AssertTrue stats, prefix & ".noInputOrInternalError." & CStr(meta.ResultKind), _
                    meta.InternalStatus <> rsInvalidInput And meta.InternalStatus <> rsInvalidConfiguration And _
                    meta.InternalStatus <> rsInternalError
                If meta.InternalStatus = rsNumericalFailure Then AppendLine stats, _
                    "FOCUSED_UNRESOLVED|shape=" & shapeName & "|id=" & CStr(item(0)) & _
                    "|kind=" & CStr(meta.ResultKind) & "|reason=" & meta.ResultComment
            Next meta
            FocusedCheckOutputRow stats, "rngStrengthSummaryAnchor", CStr(item(0)), prefix
            FocusedCheckOutputRow stats, "rngCrackSummaryAnchor", CStr(item(0)), prefix
            AppendLine stats, "FOCUSED_RESULT|id=" & CStr(item(0)) & "|overall=" & result.Status & _
                "|capacity=" & result.StrengthResult.Capacity.Status & "|cracks=" & result.CrackSummaryStatus
        ElseIf CStr(item(0)) <> "zero" Then
            AssertTrue stats, prefix & ".capturedCalculation", False
        End If
        completed = completed + 1
        Audit03SaveMatrixProgress stats.Report
    Next item
    AssertTrue stats, "focused.complete", completed = cases.Count
    AppendLine stats, "TOTAL_FOCUSED_ACCEPTANCE: shape=" & shapeName & "; group=" & groupName & _
        "; cases=" & CStr(completed) & "; maxSeconds=" & FormatNumberInvariant(maxSeconds) & _
        "; passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    GoTo Restore
Failed:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: focused.runtime; " & CStr(Err.Number) & "; " & Err.Description
Restore:
    On Error GoTo RestoreFailed
    If IsArray(oldSystem) Then systemRange.Formula = oldSystem
    If IsArray(oldProfiles) Then profileRange.Formula = oldProfiles
    If IsArray(oldLoads) Then loadRange.Formula = oldLoads
    Set mFocusedBatch = Nothing: Set mFocusedSection = Nothing
    RunFocusedAcceptanceCases = stats.Report
    Exit Function
RestoreFailed:
    RunFocusedAcceptanceCases = stats.Report & "RESTORE ERROR: " & CStr(Err.Number) & "; " & Err.Description
End Function

' ДЛЯ ТЕСТОВ: воспроизводит один замер с теми же данными и проверками.
Public Function RunFocusedSingleLoad(ByVal shapeName As String, ByVal loadName As String) As String
    RunFocusedSingleLoad = RunFocusedAcceptanceCases(shapeName, "Generated", loadName)
End Function

' ДЛЯ ТЕСТОВ: генерируемые формы читаются настоящим реестром из Config.
Private Sub FocusedSelectGeometry(ByVal shapeName As String)
    Select Case shapeName
        Case "SavedWorkbook": SetSystemSetting "Geometry.Source", "AutoCAD"
        Case "RectSet", "Circle", "RoundedRectangle", "HollowRectangle"
            SetSystemSetting "Geometry.Source", "Generated"
            SetSystemSetting "Geometry.Type", shapeName
        Case Else: Err.Raise 5, "FocusedSelectGeometry", "Unknown focused geometry: " & shapeName
    End Select
End Sub

' ДЛЯ ТЕСТОВ: выбирает нагрузки по именам, не по положению в старой матрице.
Private Function FocusedSelectedLoads(ByVal section As CSectionModel, ByVal groupName As String) As Collection
    Dim result As New Collection, item As Variant, path As Variant, names As Variant
    names = Array("axialSmall-1", "axialSmall1", "Mx-1", "Mx1", "My-1", "My1", "mixedC1_-1", "mixedT-1_1")
    If groupName = "Paths" Then
        For Each item In Audit03NormalizedLoads(section, "Light")
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
        For Each item In Audit03NormalizedLoads(section, "Light")
            If FocusedNameIn(CStr(item(0)), names) Then result.Add Array(item(0), item(1), item(2), item(3), "Auto")
        Next item
        If groupName = "Saved" Then
            For Each item In Audit03NormalizedLoads(section, "Stress")
                If FocusedNameIn(CStr(item(0)), Array("axialC" & CStr(0.95), "axialC1", "axialC" & CStr(1.05), "axialC2", _
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
    Set FocusedSelectedLoads = result
End Function

Private Function FocusedNameIn(ByVal name As String, ByVal names As Variant) As Boolean
    Dim item As Variant
    For Each item In names
        If name = CStr(item) Then FocusedNameIn = True: Exit Function
    Next item
End Function

' ДЛЯ ТЕСТОВ: reader намеренно не запускает расчет для полностью нулевой строки.
Private Sub FocusedCheckZeroInput(ByRef stats As TBatchTestStats, ByVal prefix As String)
    On Error GoTo Rejected
    Dim message As String
    message = RunSectionCalculationForWorkbook(ThisWorkbook, False)
    AssertTrue stats, prefix & ".zeroRejected", False
    Exit Sub
Rejected:
    Dim errorNumber As Long
    errorNumber = Err.Number: Err.Clear
    AssertTrue stats, prefix & ".zeroRejectedBeforeSolve", errorNumber = vbObjectError + 4101
    AssertTrue stats, prefix & ".zeroNoSolverCalls", SectionEquilibriumSolveCount() = 0
End Sub

' ДЛЯ ТЕСТОВ: повторное интегрирование по фактическим диаграммам каждого state.
Private Sub FocusedCheckStates(ByRef stats As TBatchTestStats, ByVal result As CCombinationResult, _
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
        AppendLine stats, "FOCUSED_STATE|id=" & prefix & "|type=" & state.StateTypeText & _
            "|converged=" & CStr(state.Converged) & "|physical=" & CStr(state.WithinPhysicalRange) & _
            "|extension=" & CStr(state.ExtensionUsed) & "|status=" & state.Status & _
            "|N=" & FormatNumberInvariant(state.TargetN) & "|Mx=" & FormatNumberInvariant(state.TargetMx) & _
            "|My=" & FormatNumberInvariant(state.TargetMy) & "|eps0=" & FormatNumberInvariant(state.Epsilon0) & _
            "|kx=" & FormatNumberInvariant(state.KappaX) & "|ky=" & FormatNumberInvariant(state.KappaY)
        If state.Converged Then
            AssertTrue stats, prefix & ".snapshot." & CStr(i), state.HasMaterialSnapshot
            Set probe = New CSectionSolver
            probe.EvaluateStrainPlane mFocusedSection, state.ConcreteDiagramSnapshot, state.SteelDiagramSnapshot, _
                state.Epsilon0, state.KappaX, state.KappaY
            AssertClose stats, prefix & ".originalN." & CStr(i), probe.Nint, state.TargetN, toleranceN
            AssertClose stats, prefix & ".originalMx." & CStr(i), probe.Mxint, state.TargetMx, toleranceMx
            AssertClose stats, prefix & ".originalMy." & CStr(i), probe.Myint, state.TargetMy, toleranceMy
            probe.NumericalEfforts n, mx, my
            AssertClose stats, prefix & ".centerN." & CStr(i), n, state.TargetN, toleranceN
            AssertClose stats, prefix & ".centerMx." & CStr(i), mx, state.TargetMx - cy * state.TargetN, toleranceMx
            AssertClose stats, prefix & ".centerMy." & CStr(i), my, state.TargetMy - cx * state.TargetN, toleranceMy
            AssertTrue stats, prefix & ".physicalStatus." & CStr(i), _
                state.WithinPhysicalRange Or state.InternalStatus = rsCheckFailed
        Else
            AssertTrue stats, prefix & ".failureNotPhysical." & CStr(i), Not state.WithinPhysicalRange
        End If
        If Not extensionEnabled Then AssertTrue stats, prefix & ".noHiddenExtension." & CStr(i), Not state.ExtensionUsed
    Next i
End Sub

Private Sub FocusedCheckOutputRow(ByRef stats As TBatchTestStats, ByVal rangeName As String, _
        ByVal id As String, ByVal prefix As String)
    Dim anchor As Object
    Set anchor = ThisWorkbook.Names.Item(rangeName).RefersToRange.Cells(1, 1)
    AssertEquals stats, prefix & ".output." & rangeName, CStr(anchor.Value2), id
End Sub
