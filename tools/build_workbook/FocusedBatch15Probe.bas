' ==========================================================================
' ДЛЯ ТЕСТОВ — вспомогательный фрагмент частной приёмки; в выпуск не включается.
' ==========================================================================

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Архивный entrypoint: исходный SavedCrackSearchProbe отсутствует; актуальные 15 сочетаний проверяет CoreRegressionBatchProbe.
Public Function TestRunFocusedBatch15() As String
    Err.Raise vbObjectError + 5917, "TestRunFocusedBatch15", "Архивная проверка требует исходный SavedCrackSearchProbe.bas от 09.10.2026, которого нет в репозитории. Используйте tools/build_workbook/Run-FocusedAcceptance.ps1 для действующей полной приёмки с пакетом из 15 сочетаний."
#If False Then ' Фиксированное тело описывает ожидания batch15 и отключено без SavedCrackSearchProbe; фиктивная реализация зависимости не используется.
    On Error GoTo Failed
    Dim stats As TBatchTestStats, body As String, result As CCombinationResult
    Dim settings As CSystemSettingsReader, units As CUnitSystem, i As Long, meta As Variant
    Set mFocusedBatch = Nothing: Set mFocusedSection = Nothing
    body = RunSavedCrackSearchProbe("batch15")
    TestAssertTrue stats, "batch15.completed", InStr(1, body, "TOTAL_CRACK_SEARCH_PROBE: completed=True", vbBinaryCompare) > 0
    TestAssertTrue stats, "batch15.capture", Not mFocusedBatch Is Nothing
    If Not mFocusedBatch Is Nothing Then
        TestAssertTrue stats, "batch15.count", mFocusedBatch.Count = 15
        TestAssertTrue stats, "batch15.preparationReleased", mFocusedBatch.TestRunPreparationReleased
        TestAssertTrue stats, "batch15.mechanicalNoRepeat", mFocusedBatch.TestMechanicalPreparationBuildCount > 0 And _
            mFocusedBatch.TestMechanicalPreparationBuildCount <= 2
        TestAssertTrue stats, "batch15.crackDomainNoRepeat", mFocusedBatch.TestCrackDomainBuildCount <= 1
        TestAppendLine stats, "BATCH15_PREPARATION|mechanicalBuilds=" & CStr(mFocusedBatch.TestMechanicalPreparationBuildCount) & _
            "|crackDomainBuilds=" & CStr(mFocusedBatch.TestCrackDomainBuildCount)
        Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
        Set units = New CUnitSystem: units.LoadFromSettings settings
        For i = 1 To mFocusedBatch.Count
            Set result = mFocusedBatch.ResultAt(i)
            TestAudit03CheckResultComments stats, mFocusedBatch, i
            TestFocusedCheckStates stats, result, settings, units, "batch15." & CStr(i)
            For Each meta In Array(result.DirectStateMeta, result.CapacityMeta, result.CrackFormationMeta, _
                    result.CrackCurrentStateMeta, result.CrackWidthMeta, result.LongitudinalCrackMeta, result.OverallMeta)
                TestAssertTrue stats, "batch15.noNumericalOrInputError." & CStr(i) & "." & CStr(meta.ResultKind), _
                    meta.InternalStatus <> rsNumericalFailure And meta.InternalStatus <> rsInvalidInput And _
                    meta.InternalStatus <> rsInvalidConfiguration And meta.InternalStatus <> rsInternalError
            Next meta
        Next i
    End If
    GoTo Finished
Failed:
    TestAssertTrue stats, "batch15.runtime." & CStr(Err.Number) & "." & Err.Description, False
Finished:
    Set mFocusedBatch = Nothing: Set mFocusedSection = Nothing
    TestRunFocusedBatch15 = body & stats.Report & "TOTAL_FOCUSED_BATCH15: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
#End If
End Function
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет четыре характерные нагрузки каждой встроенной формы через общий focused entrypoint и сохраняет все отказы.
Public Function TestRunFocusedRepresentativeGeometry() As String
    Dim shape As Variant, loadName As Variant, body As String, one As String, count As Long, failures As Long
    For Each shape In Array("Rectangles", "Circle", "PointedSection", "HollowRectangle")
        For Each loadName In Array("axialSmall-1", "Mx1", "mixedC1_-1", "mixedT-1_1")
            one = TestRunFocusedSingleLoad(CStr(shape), CStr(loadName))
            body = body & one & vbCrLf
            If InStr(1, one, "FAIL:", vbBinaryCompare) > 0 Or _
                    InStr(1, one, "FOCUSED_UNRESOLVED", vbBinaryCompare) > 0 Or _
                    InStr(1, one, "RESTORE ERROR", vbBinaryCompare) > 0 Then failures = failures + 1
            If InStr(1, one, "TOTAL_FOCUSED_ACCEPTANCE:", vbBinaryCompare) = 0 Then failures = failures + 1
            count = count + 1
        Next loadName
    Next shape
    TestRunFocusedRepresentativeGeometry = body & "TOTAL_FOCUSED_REPRESENTATIVE_GEOMETRY: cases=" & CStr(count) & _
        "; failed=" & CStr(failures)
End Function
