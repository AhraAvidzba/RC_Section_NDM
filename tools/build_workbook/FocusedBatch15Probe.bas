' ДЛЯ ТЕСТОВ: полная запись 15 сочетаний и контроль повторного использования подготовки.
Public Function RunFocusedBatch15() As String
    On Error GoTo Failed
    Dim stats As TBatchTestStats, body As String, result As CCombinationResult
    Dim settings As CSystemSettingsReader, units As CUnitSystem, i As Long, meta As Variant
    Set mFocusedBatch = Nothing: Set mFocusedSection = Nothing
    body = RunSavedCrackSearchProbe("batch15")
    AssertTrue stats, "batch15.completed", InStr(1, body, "TOTAL_CRACK_SEARCH_PROBE: completed=True", vbBinaryCompare) > 0
    AssertTrue stats, "batch15.capture", Not mFocusedBatch Is Nothing
    If Not mFocusedBatch Is Nothing Then
        AssertTrue stats, "batch15.count", mFocusedBatch.Count = 15
        AssertTrue stats, "batch15.preparationReleased", mFocusedBatch.RunPreparationReleased
        AssertTrue stats, "batch15.mechanicalNoRepeat", mFocusedBatch.MechanicalPreparationBuildCount > 0 And _
            mFocusedBatch.MechanicalPreparationBuildCount <= 2
        AssertTrue stats, "batch15.crackDomainNoRepeat", mFocusedBatch.CrackDomainBuildCount <= 1
        AppendLine stats, "BATCH15_PREPARATION|mechanicalBuilds=" & CStr(mFocusedBatch.MechanicalPreparationBuildCount) & _
            "|crackDomainBuilds=" & CStr(mFocusedBatch.CrackDomainBuildCount)
        Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
        Set units = New CUnitSystem: units.LoadFromSettings settings
        For i = 1 To mFocusedBatch.Count
            Set result = mFocusedBatch.ResultAt(i)
            Audit03CheckResultComments stats, mFocusedBatch, i
            FocusedCheckStates stats, result, settings, units, "batch15." & CStr(i)
            For Each meta In Array(result.DirectStateMeta, result.CapacityMeta, result.CrackFormationMeta, _
                    result.CrackCurrentStateMeta, result.CrackWidthMeta, result.LongitudinalCrackMeta, result.OverallMeta)
                AssertTrue stats, "batch15.noNumericalOrInputError." & CStr(i) & "." & CStr(meta.ResultKind), _
                    meta.InternalStatus <> rsNumericalFailure And meta.InternalStatus <> rsInvalidInput And _
                    meta.InternalStatus <> rsInvalidConfiguration And meta.InternalStatus <> rsInternalError
            Next meta
        Next i
    End If
    GoTo Finished
Failed:
    AssertTrue stats, "batch15.runtime." & CStr(Err.Number) & "." & Err.Description, False
Finished:
    Set mFocusedBatch = Nothing: Set mFocusedSection = Nothing
    RunFocusedBatch15 = body & stats.Report & "TOTAL_FOCUSED_BATCH15: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
End Function

' ДЛЯ ТЕСТОВ: четыре показательных нагрузки для каждой встроенной формы.
Public Function RunFocusedRepresentativeGeometry() As String
    Dim shape As Variant, loadName As Variant, body As String, one As String, count As Long, failures As Long
    For Each shape In Array("Rectangles", "Circle", "PointedSection", "HollowRectangle")
        For Each loadName In Array("axialSmall-1", "Mx1", "mixedC1_-1", "mixedT-1_1")
            one = RunFocusedSingleLoad(CStr(shape), CStr(loadName))
            body = body & one & vbCrLf
            If InStr(1, one, "FAIL:", vbBinaryCompare) > 0 Or _
                    InStr(1, one, "FOCUSED_UNRESOLVED", vbBinaryCompare) > 0 Or _
                    InStr(1, one, "RESTORE ERROR", vbBinaryCompare) > 0 Then failures = failures + 1
            If InStr(1, one, "TOTAL_FOCUSED_ACCEPTANCE:", vbBinaryCompare) = 0 Then failures = failures + 1
            count = count + 1
        Next loadName
    Next shape
    RunFocusedRepresentativeGeometry = body & "TOTAL_FOCUSED_REPRESENTATIVE_GEOMETRY: cases=" & CStr(count) & _
        "; failed=" & CStr(failures)
End Function
