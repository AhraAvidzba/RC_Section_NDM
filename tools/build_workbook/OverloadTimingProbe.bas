' ==========================================================================
' ДЛЯ ТЕСТОВ — вспомогательный фрагмент частной приёмки; в выпуск не включается.
' ==========================================================================

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Измеряет полный штатный расчёт перегрузки и выдаёт счётчики, статусы и диагностику предельного поиска.
Public Function TestRunOverloadTimingCase(ByVal specification As String) As String
    On Error GoTo Failed
    Dim parts As Variant, settings As CSystemSettingsReader, units As CUnitSystem
    Dim loads As Object, values() As Variant, started As Double, elapsed As Double
    Dim message As String, result As CCombinationResult, capacity As CCapacityResult
    parts = Split(specification, "|")
    If UBound(parts) <> 6 Then Err.Raise 5, , "Expected seven timing arguments."
    TestFocusedSelectGeometry CStr(parts(0))
    TestSetSystemSetting "Load.ReferenceOffsetX", "0"
    TestSetSystemSetting "Load.ReferenceOffsetY", "0"
    TestSetSystemSetting "Calculation.ZeroMomentPerDepth", "0"
    TestSetSystemSetting "General.ExecutionReportEnabled", "Yes"
    TestSetSystemSetting "General.NonCriticalMessagesEnabled", "No"
    TestSetSystemSetting "General.DiagramExtension", "Yes"
    TestSetSystemSetting "Capacity.SolutionStrategy", CStr(parts(6))
    TestSetProfileValue "Calculation.Strength.DirectState", "PR3", "Yes"
    TestSetProfileValue "Calculation.Strength.Capacity", "PR3", "Yes"
    TestSetProfileValue "Calculation.Crack.Width", "PR3", "No"
    TestSetProfileValue "Calculation.Stability.Enabled", "PR3", "No"
    TestSetProfileValue "Visualization.State", "PR3", "StrengthState"
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    Set loads = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange
    ReDim values(1 To loads.Rows.Count - 1, 1 To loads.Columns.Count)
    values(1, 1) = CStr(parts(1))
    values(1, 2) = Val(CStr(parts(2))) / units.InputForceToInternal(1#)
    values(1, 3) = Val(CStr(parts(3))) / units.InputMomentMxToInternal(1#)
    values(1, 4) = Val(CStr(parts(4))) / units.InputMomentMyToInternal(1#)
    values(1, 5) = "PR3": values(1, 6) = CStr(parts(5))
    loads.Offset(1, 0).Resize(loads.Rows.Count - 1, loads.Columns.Count).Value2 = values
    Set mFocusedBatch = Nothing: Set mFocusedSection = Nothing
    started = Timer
    message = TestFixtureRunSectionWorkflow(ThisWorkbook, False)
    elapsed = Timer - started: If elapsed < 0# Then elapsed = elapsed + 86400#
    If mFocusedBatch Is Nothing Then Err.Raise 5, , "Calculation did not reach batch completion: " & message
    Set result = mFocusedBatch.ResultAt(1)
    Set capacity = result.StrengthResult.Capacity
    TestRunOverloadTimingCase = "OVERLOAD|id=" & CStr(parts(1)) & "|seconds=" & TestFormatNumberInvariant(elapsed) & _
        "|fibers=" & CStr(mFocusedSection.ConcreteCount) & "|bars=" & CStr(mFocusedSection.RebarCount) & _
        "|solves=" & CStr(SectionEquilibriumSolveCount()) & "|iterations=" & CStr(SectionEquilibriumIterationCount()) & _
        "|ultimateStarts=" & CStr(UltimateStrainStartCount()) & "|ultimateIterations=" & CStr(UltimateStrainIterationCount()) & _
        "|overall=" & result.Status & "|direct=" & result.StrengthResult.DirectState.Status & _
        "|capacity=" & capacity.Status & "|lambda=" & TestFormatNumberInvariant(capacity.LambdaCapacity) & _
        "|method=" & capacity.SolutionMethod & "|path=" & capacity.PathResolved & _
        "|plateauFactor=" & settings.GetRequiredString("General.PlateauSlopeFactor") & vbCrLf & _
        "CAPACITY_DIAGNOSTICS" & vbCrLf & capacity.DiagnosticLog & vbCrLf & "MESSAGE" & vbCrLf & message
    Exit Function
Failed:
    TestRunOverloadTimingCase = "RUNTIME ERROR: " & CStr(Err.Number) & "; " & Err.Description
End Function
