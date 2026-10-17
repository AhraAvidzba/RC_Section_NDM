' ДЛЯ ТЕСТОВ: один штатный batch из 15 разных сочетаний, включая N, M и N+Mx+My.
Public Function RunCoreRegressionBatch15() As String
    On Error GoTo Failed
    Dim stats As TBatchTestStats, oldSystem As Variant, oldProfiles As Variant, oldLoads As Variant
    Dim systemRange As Object, profileRange As Object, loadRange As Object
    Dim settings As CSystemSettingsReader, units As CUnitSystem, section As CSectionModel
    Dim allCases As Collection, cases As New Collection, item As Variant, names As Variant
    Dim values() As Variant, i As Long, started As Double, elapsed As Double, message As String
    Dim result As CCombinationResult
    Set systemRange = ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange
    Set profileRange = ThisWorkbook.Names.Item("rngCalculationProfiles").RefersToRange
    Set loadRange = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange
    oldSystem = systemRange.Formula: oldProfiles = profileRange.Formula: oldLoads = loadRange.Formula
    FocusedSelectGeometry "SavedWorkbook"
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
    Set allCases = FocusedSelectedLoads(section, "Saved")
    names = Array("axialSmall-1", "axialSmall1", "Mx-1", "Mx1", "My-1", "My1", "mixedC1_-1", "mixedT-1_1", _
        "tinyMx1", "biaxial1_-1_10", "biaxial1_-1_" & CStr(0.000001), "N500", "N1000", "N500TinyMoment", "N2098Plateau")
    For Each item In allCases
        If FocusedNameIn(CStr(item(0)), names) Then cases.Add item
    Next item
    If cases.Count <> 15 Then Err.Raise 5, , "Batch must contain exactly 15 selected loads."
    ReDim values(1 To loadRange.Rows.Count - 1, 1 To loadRange.Columns.Count)
    i = 0
    For Each item In cases
        i = i + 1
        values(i, 1) = CStr(item(0))
        values(i, 2) = CDbl(item(1)) / units.InputForceToInternal(1#)
        values(i, 3) = CDbl(item(2)) / units.InputMomentMxToInternal(1#)
        values(i, 4) = CDbl(item(3)) / units.InputMomentMyToInternal(1#)
        values(i, 5) = "PR3": values(i, 6) = "Auto"
    Next item
    loadRange.Offset(1, 0).Resize(loadRange.Rows.Count - 1, loadRange.Columns.Count).Value2 = values
    Set mFocusedBatch = Nothing: Set mFocusedSection = Nothing
    ResetSectionEquilibriumSolveCount
    started = Timer
    message = RunSectionCalculationForWorkbook(ThisWorkbook, False)
    elapsed = Timer - started: If elapsed < 0# Then elapsed = elapsed + 86400#
    AssertTrue stats, "plateau.batch15.capture", Not mFocusedBatch Is Nothing
    If Not mFocusedBatch Is Nothing Then
        AssertTrue stats, "plateau.batch15.count", mFocusedBatch.Count = 15
        AssertTrue stats, "plateau.batch15.preparationReleased", mFocusedBatch.RunPreparationReleased
        AssertTrue stats, "plateau.batch15.noRepeatedMaterialPreparation", mFocusedBatch.MechanicalPreparationBuildCount <= 2
        AssertTrue stats, "plateau.batch15.oneCrackDomain", mFocusedBatch.CrackDomainBuildCount <= 1
        AppendLine stats, "CORE_BATCH_TIME|seconds=" & FormatNumberInvariant(elapsed)
        ResetSectionEquilibriumSolveCount
        For i = 1 To mFocusedBatch.Count
            Set result = mFocusedBatch.ResultAt(i)
            item = cases(i)
            AppendLine stats, FocusedComparisonRecord(result, "batch15." & CStr(item(0)))
            FocusedCheckStates stats, result, settings, units, "plateau.batch15." & CStr(item(0))
            Audit03CheckResultComments stats, mFocusedBatch, i
        Next i
    End If
    AppendLine stats, "TOTAL_CORE_BATCH15: seconds=" & FormatNumberInvariant(elapsed) & _
        "; passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    GoTo Restore
Failed:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: plateau.batch15.runtime; " & CStr(Err.Number) & "; " & Err.Description
Restore:
    If IsArray(oldSystem) Then systemRange.Formula = oldSystem
    If IsArray(oldProfiles) Then profileRange.Formula = oldProfiles
    If IsArray(oldLoads) Then loadRange.Formula = oldLoads
    Set mFocusedBatch = Nothing: Set mFocusedSection = Nothing
    RunCoreRegressionBatch15 = stats.Report
End Function

' ДЛЯ ТЕСТОВ: два полных смешанных расчета, только в независимой копии.
Public Function RunFocusedSpecialCase(ByVal caseName As String) As String
    On Error GoTo Failed
    Dim stats As TBatchTestStats, systemRange As Object, profileRange As Object, loadRange As Object
    Dim oldSystem As Variant, oldProfiles As Variant, oldLoads As Variant, oldSteel As Variant, values() As Variant
    Dim steelRange As Object
    Dim durationRange As Object, oldDuration As Variant, durationValues() As Variant
    Dim settings As CSystemSettingsReader, units As CUnitSystem, section As CSectionModel
    Dim provider As CMaterialModelProvider, profiles As CCalculationProfileCatalog, frame As CSectionNumericalFrame
    Dim concrete As CMaterialDiagram, steel As CMaterialDiagram, evaluator As CSectionSolver
    Dim result As CCombinationResult, state As CSectionStateResult, item As Variant, meta As Variant
    Dim n As Double, mx As Double, my As Double, started As Double, elapsed As Double, message As String
    Set systemRange = ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange
    Set profileRange = ThisWorkbook.Names.Item("rngCalculationProfiles").RefersToRange
    Set loadRange = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange
    oldSystem = systemRange.Formula: oldProfiles = profileRange.Formula: oldLoads = loadRange.Formula
    Set steelRange = ThisWorkbook.Names.Item("rngSteelMaterialParameters").RefersToRange
    oldSteel = steelRange.Formula
    Set durationRange = ThisWorkbook.Names.Item("rngStabilityDurationLoads").RefersToRange
    oldDuration = durationRange.Formula
    FocusedSelectGeometry "SavedWorkbook"
    SetSystemSetting "Load.ReferenceOffsetX", "0": SetSystemSetting "Load.ReferenceOffsetY", "0"
    SetSystemSetting "Calculation.ZeroMomentPerDepth", "0"
    SetSystemSetting "General.ExecutionReportEnabled", "Yes"
    SetSystemSetting "Plot.Enabled", "Yes": SetSystemSetting "Plot.AutoUpdateAfterCalculation", "Yes"
    SetSystemSetting "Plot.LoadCase", "Worst": SetSystemSetting "SLS.Crack.Code", "SP35"
    SetProfileValue "Calculation.Strength.DirectState", "PR3", "Yes"
    SetProfileValue "Calculation.Strength.Capacity", "PR3", "Yes"
    SetProfileValue "Calculation.Crack.Width", "PR3", "Yes"
    SetProfileValue "Calculation.Stability.Enabled", "PR3", IIf(caseName = "Stability", "Yes", "No")
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    Select Case caseName
        Case "Stability"
            SetSystemSetting "Stability.Code", "SP35"
            SetSystemSetting "Stability.ElementLength", FormatNumberInvariant(1000# / units.InputLengthToInternal(1#))
            SetSystemSetting "Stability.Mu1", "1": SetSystemSetting "Stability.Mu2", "1"
            SetSystemSetting "Stability.AccidentalEccentricityMode", "User"
            SetSystemSetting "Stability.AccidentalEccentricityPlanes", "BothPlanes"
            SetSystemSetting "Stability.AccidentalEccentricityUser1", FormatNumberInvariant(5# / units.InputLengthToInternal(1#))
            SetSystemSetting "Stability.AccidentalEccentricityUser2", FormatNumberInvariant(5# / units.InputLengthToInternal(1#))
            SetProfileValue "MaterialModel.Stability.ValueSet", "PR3", "ULS(I)"
        Case "AsymmetricSteel"
            For Each item In Array(Array("Steel.R.ULS(I)", 2, 250#), Array("Steel.R.ULS(I)", 3, 400#), _
                    Array("Steel.R.SLS(II)", 2, 280#), Array("Steel.R.SLS(II)", 3, 450#), _
                    Array("Steel.E", 2, 190000#), Array("Steel.E", 3, 210000#))
                FocusedSetSteelValue steelRange, CStr(item(0)), CLng(item(1)), CDbl(item(2)) / units.InputStressToInternal(1#)
            Next item
            FocusedSetSteelValue steelRange, "Steel.TwoLine.Es2", 2, 0.02
            FocusedSetSteelValue steelRange, "Steel.TwoLine.Es2", 3, 0.03
            FocusedSetSteelValue steelRange, "Steel.ThreeLine.Es2", 2, 0.018
            FocusedSetSteelValue steelRange, "Steel.ThreeLine.Es2", 3, 0.022
            SetProfileValue "MaterialModel.Strength.SteelDiagram", "PR3", "TwoLine"
            SetProfileValue "MaterialModel.CrackInitiation.SteelDiagram", "PR3", "TwoLine"
            SetProfileValue "MaterialModel.CrackedState.SteelDiagram", "PR3", "TwoLine"
        Case Else
            Err.Raise 5, "RunFocusedSpecialCase", "Unknown focused special case."
    End Select
    settings.LoadFromWorkbook ThisWorkbook
    Set section = BuildWorkbookSectionModel(ThisWorkbook, settings, units)
    Set provider = New CMaterialModelProvider: provider.Initialize settings, units
    Set profiles = New CCalculationProfileCatalog: profiles.LoadFromWorkbook ThisWorkbook
    Set concrete = provider.ConcreteMaterialForEquilibriumFromSpec(profiles.ProfileById("PR3").StrengthSpec)
    Set steel = provider.SteelMaterialForEquilibriumFromSpec(profiles.ProfileById("PR3").StrengthSpec)
    If caseName = "AsymmetricSteel" Then
        AssertClose stats, "special.steel.compressionLimit", steel.PhysicalCompressionStrain, -0.02, 0#
        AssertClose stats, "special.steel.tensionLimit", steel.PhysicalTensionStrain, 0.03, 0#
        ' На плато к сопротивлению прибавляется alpha * E * длина участка после текучести.
        AssertClose stats, "special.steel.compressionStress", steel.GetStress(-0.02), _
            -250# - settings.GetRequiredDouble("General.PlateauSlopeFactor") * 190000# * (0.02 - 250# / 190000#), 0.00000001
        AssertClose stats, "special.steel.tensionStress", steel.GetStress(0.03), _
            400# + settings.GetRequiredDouble("General.PlateauSlopeFactor") * 210000# * (0.03 - 400# / 210000#), 0.00000001
    End If
    Set frame = New CSectionNumericalFrame: frame.Initialize section
    Set evaluator = New CSectionSolver
    evaluator.EvaluatePhysicalCenterPlane section, concrete, steel, -0.00015, _
        0.0007 / frame.LengthScale, -0.0005 / frame.LengthScale
    n = evaluator.Nint: mx = frame.CenterMx(n, evaluator.Mxint): my = frame.CenterMy(n, evaluator.Myint)
    ReDim values(1 To loadRange.Rows.Count - 1, 1 To loadRange.Columns.Count)
    values(1, 1) = caseName: values(1, 2) = n / units.InputForceToInternal(1#)
    values(1, 3) = mx / units.InputMomentMxToInternal(1#): values(1, 4) = my / units.InputMomentMyToInternal(1#)
    values(1, 5) = "PR3": values(1, 6) = "Auto"
    loadRange.Offset(1, 0).Resize(loadRange.Rows.Count - 1, loadRange.Columns.Count).Value2 = values
    If caseName = "Stability" Then
        ' Совпадающий номер обязателен; нулевые постоянные нагрузки допустимы.
        ReDim durationValues(1 To durationRange.Rows.Count - 1, 1 To durationRange.Columns.Count)
        durationValues(1, 1) = caseName
        durationValues(1, 2) = 0#: durationValues(1, 3) = 0#: durationValues(1, 4) = 0#
        durationRange.Offset(1, 0).Resize(durationRange.Rows.Count - 1, durationRange.Columns.Count).Value2 = durationValues
    End If
    AppendLine stats, "SPECIAL_INPUT|case=" & caseName & "|N=" & FormatNumberInvariant(n) & _
        "|Mx=" & FormatNumberInvariant(mx) & "|My=" & FormatNumberInvariant(my)
    Set mFocusedBatch = Nothing: Set mFocusedSection = Nothing
    ResetSectionEquilibriumSolveCount: started = Timer
    message = RunSectionCalculationForWorkbook(ThisWorkbook, False)
    elapsed = Timer - started: If elapsed < 0# Then elapsed = elapsed + 86400#
    AssertTrue stats, "special.within10Seconds", elapsed <= 10#
    AssertTrue stats, "special.capture", Not mFocusedBatch Is Nothing
    If Not mFocusedBatch Is Nothing Then
        Set result = mFocusedBatch.ResultAt(1)
        ' Проверка трещин может законно дать FAIL при равновесном НДС.
        ' Состав общего статуса и причины каждого результата проверяются ниже.
        AssertTrue stats, "special.overall.typed", result.Status = "OK" Or result.Status = "FAIL"
        AssertTrue stats, "special.preparationReleased", mFocusedBatch.RunPreparationReleased
        AssertTrue stats, "special.crackDomainOnce", mFocusedBatch.CrackDomainBuildCount <= 1
        AssertTrue stats, "special.mechanicalNoRepeat", mFocusedBatch.MechanicalPreparationBuildCount <= IIf(caseName = "Stability", 3, 2)
        If caseName = "Stability" Then AssertEquals stats, "special.stability", result.StabilityResult.Status, "OK"
        Audit03CheckResultComments stats, mFocusedBatch, 1
        FocusedCheckStates stats, result, settings, units, "special." & caseName
        For Each meta In Array(result.DirectStateMeta, result.CapacityMeta, result.CrackFormationMeta, _
                result.CrackCurrentStateMeta, result.CrackWidthMeta, result.LongitudinalCrackMeta, result.OverallMeta)
            AssertTrue stats, "special.noNumericalOrInputError." & CStr(meta.ResultKind), _
                meta.InternalStatus <> rsNumericalFailure And meta.InternalStatus <> rsInvalidInput And _
                meta.InternalStatus <> rsInvalidConfiguration And meta.InternalStatus <> rsInternalError
        Next meta
        FocusedCheckOutputRow stats, "rngStrengthSummaryAnchor", caseName, "special"
        FocusedCheckOutputRow stats, "rngCrackSummaryAnchor", caseName, "special"
    End If
    AppendLine stats, "TOTAL_FOCUSED_SPECIAL: case=" & caseName & "; elapsedSec=" & FormatNumberInvariant(elapsed) & _
        "; passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    GoTo Restore
Failed:
    AssertTrue stats, "special.runtime." & CStr(Err.Number) & "." & Err.Description, False
Restore:
    On Error GoTo RestoreFailed
    If IsArray(oldSystem) Then systemRange.Formula = oldSystem
    If IsArray(oldProfiles) Then profileRange.Formula = oldProfiles
    If IsArray(oldLoads) Then loadRange.Formula = oldLoads
    If IsArray(oldSteel) Then steelRange.Formula = oldSteel
    If IsArray(oldDuration) Then durationRange.Formula = oldDuration
    Set mFocusedBatch = Nothing: Set mFocusedSection = Nothing
    RunFocusedSpecialCase = stats.Report
    Exit Function
RestoreFailed:
    RunFocusedSpecialCase = stats.Report & "RESTORE ERROR: " & CStr(Err.Number) & "; " & Err.Description
End Function

Private Sub FocusedSetSteelValue(ByVal table As Object, ByVal key As String, ByVal column As Long, ByVal value As Double)
    Dim row As Long
    For row = 2 To table.Rows.Count
        If StrComp(CStr(table.Cells(row, 1).Value2), key, vbTextCompare) = 0 Then
            table.Cells(row, column).Value2 = value
            Exit Sub
        End If
    Next row
    Err.Raise 5, "FocusedSetSteelValue", "Steel material row not found: " & key
End Sub

' ДЛЯ ТЕСТОВ: фактическая геометрия и диаграммы PR3 для независимых границ усилий.
Public Function DumpFocusedForceBounds(Optional ByVal shapeName As String = "SavedWorkbook") As String
    On Error GoTo Failed
    Dim settings As CSystemSettingsReader, units As CUnitSystem, section As CSectionModel
    Dim profiles As CCalculationProfileCatalog, provider As CMaterialModelProvider, spec As CMaterialModelSpec
    Dim concrete As CMaterialDiagram, steel As CMaterialDiagram, mode As Variant, label As Variant
    Dim output As String, x As Double, y As Double, i As Long, oldSystem As Variant
    oldSystem = ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange.Formula
    FocusedSelectGeometry shapeName
    SetSystemSetting "SLS.Crack.Code", "SP35"
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    Set section = BuildWorkbookSectionModel(ThisWorkbook, settings, units)
    Set profiles = New CCalculationProfileCatalog: profiles.LoadFromWorkbook ThisWorkbook
    CalculateConcreteSectionCentroid section, x, y
    output = "CENTER|x=" & FocusedBoundNumber(x) & "|y=" & FocusedBoundNumber(y) & vbCrLf
    output = output & "ACCEPTANCE_TOLERANCE|N=" & FocusedBoundNumber(Abs(units.InputForceToInternal(settings.GetDouble("Solver.ToleranceN")))) & _
        "|Mx=" & FocusedBoundNumber(Abs(units.InputMomentMxToInternal(settings.GetDouble("Solver.ToleranceMx")))) & _
        "|My=" & FocusedBoundNumber(Abs(units.InputMomentMyToInternal(settings.GetDouble("Solver.ToleranceMy")))) & vbCrLf
    For i = 1 To section.ConcreteCount
        output = output & "FIBER|kind=C|x=" & FocusedBoundNumber(section.ConcreteX(i) - x) & _
            "|y=" & FocusedBoundNumber(section.ConcreteY(i) - y) & "|area=" & FocusedBoundNumber(section.ConcreteArea(i)) & vbCrLf
    Next i
    For i = 1 To section.RebarCount
        output = output & "FIBER|kind=S|x=" & FocusedBoundNumber(section.RebarX(i) - x) & _
            "|y=" & FocusedBoundNumber(section.RebarY(i) - y) & "|area=" & FocusedBoundNumber(section.RebarArea(i)) & vbCrLf
    Next i
    For Each mode In Array("No", "Yes")
        SetSystemSetting "General.DiagramExtension", CStr(mode)
        settings.LoadFromWorkbook ThisWorkbook
        Set provider = New CMaterialModelProvider: provider.Initialize settings, units
        For Each label In Array("PR1", "PR2", "PR3")
            Select Case CStr(label)
                Case "PR1": Set spec = profiles.ProfileById("PR3").StrengthSpec
                Case "PR2": Set spec = profiles.ProfileById("PR3").CrackedStateSpec
                Case "PR3": Set spec = profiles.ProfileById("PR3").CrackInitiationSpec
            End Select
            Set concrete = provider.ConcreteMaterialForEquilibriumFromSpec(spec)
            Set steel = provider.SteelMaterialForEquilibriumFromSpec(spec)
            FocusedBoundPoints output, CStr(mode), CStr(label), "C", concrete
            FocusedBoundPoints output, CStr(mode), CStr(label), "S", steel
        Next label
    Next mode
    GoTo Restore
Failed:
    output = "RUNTIME ERROR: " & CStr(Err.Number) & "; " & Err.Description
Restore:
    If IsArray(oldSystem) Then ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange.Formula = oldSystem
    DumpFocusedForceBounds = output
End Function

Private Sub FocusedBoundPoints(ByRef output As String, ByVal mode As String, ByVal profile As String, _
        ByVal kind As String, ByVal material As CMaterialDiagram)
    Dim i As Long
    output = output & "MATERIAL_FLAGS|mode=" & mode & "|profile=" & profile & "|kind=" & kind & _
        "|compressionExtension=" & CStr(material.HasCompressionExtension) & "|tensionExtension=" & CStr(material.HasTensionExtension) & _
        "|compressionLimit=" & FocusedBoundNumber(material.PhysicalCompressionStrain) & _
        "|tensionLimit=" & FocusedBoundNumber(material.PhysicalTensionStrain) & vbCrLf
    For i = 1 To material.PointCount
        output = output & "MATERIAL|mode=" & mode & "|profile=" & profile & "|kind=" & kind & _
            "|strain=" & FocusedBoundNumber(material.PointStrain(i)) & "|stress=" & FocusedBoundNumber(material.PointStress(i)) & vbCrLf
    Next i
End Sub

Private Function FocusedBoundNumber(ByVal value As Double) As String
    If mFocusedBinaryDump Then
        FocusedBoundNumber = FocusedExactNumber(value)
    Else
        FocusedBoundNumber = Replace$(Format$(value, "0.0000000000000000E+00"), ",", ".")
    End If
End Function

' ДЛЯ ТЕСТОВ: сохраняет все биты Double для независимой проверки больших усилий.
Public Function DumpFocusedForceBoundsExact(Optional ByVal shapeName As String = "SavedWorkbook") As String
    mFocusedBinaryDump = True
    DumpFocusedForceBoundsExact = DumpFocusedForceBounds(shapeName)
    mFocusedBinaryDump = False
End Function
