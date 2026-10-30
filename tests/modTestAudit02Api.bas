Attribute VB_Name = "modTestAudit02Api"
Option Explicit

' ==========================================================================
' ДЛЯ ТЕСТОВ: описание сечения, параметры арматуры, технические имена и совместимость размеров.
' ==========================================================================
Private mPassed As Long, mFailed As Long, mReport As String

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет разбиение итогового окна, вложенных ошибок и длинного пути отчёта;
' расчёт не запускается, исходные причины и численные значения не подменяются.
Public Function RunProgramMessagePresentationTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    On Error GoTo Failed
    Dim reason As String, message As String, path As String, text As String, expected As String
    reason = "CombinationID=2: Продольные трещины: Проверка продольных трещин не проходит: сжатие бетона выше допускаемого. " & _
        "Максимальное сжатие бетона 17.550 МПа; допускаемое Rb,mc2 = 14.600 МПа.; Устойчивость: " & _
        "Для сочетания 2 в таблице «Нагрузки для устойчивости» (rngStabilityDurationLoads, Config!P70:S100) не найдена строка с этим номером. " & _
        "Добавьте строку или исправьте номер сочетания. Если дополнительных нагрузок нет, оставьте N, Mx и My пустыми либо укажите 0. " & _
        "(статус InputErr) Исходная строка сочетания: Config!H72:N72."
    message = CalculationCompletionMessage(30, 29, reason, "")
    TestCheck "presentation.summary.heading", Left$(message, Len("Расчет завершен с ошибками ввода." & vbCrLf & vbCrLf)) = "Расчет завершен с ошибками ввода." & vbCrLf & vbCrLf
    TestCheck "presentation.summary.counts", InStr(message, "Обработано сочетаний: 30." & vbCrLf & "Ошибок ввода: 29.") > 0
    TestCheck "presentation.summary.action", InStr(message, vbCrLf & vbCrLf & "Результаты и пояснения всех проверок — на листе Results.") > 0
    TestCheck "presentation.summary.first", InStr(message, vbCrLf & vbCrLf & "Показана только первая ошибка ввода:" & vbCrLf & "Сочетание 2:") > 0
    TestCheck "presentation.summary.longitudinal", InStr(message, vbCrLf & vbCrLf & "Продольные трещины:" & vbCrLf & "Проверка продольных трещин") > 0
    TestCheck "presentation.summary.stability", InStr(message, vbCrLf & vbCrLf & "Устойчивость:" & vbCrLf & "Для сочетания 2") > 0
    TestCheck "presentation.summary.values", InStr(message, "17.550") > 0 And InStr(message, "14.600") > 0 And InStr(message, "Rb,mc2") > 0
    TestCheck "presentation.summary.addresses", InStr(message, "rngStabilityDurationLoads, Config!P70:S100") > 0 And InStr(message, "Исходная строка сочетания:" & vbCrLf & "Config!H72:N72.") > 0
    TestCheck "presentation.summary.reason.unchanged", Left$(reason, Len("CombinationID=2:")) = "CombinationID=2:"
    TestCheck "presentation.summary.idempotent", WorkbookDialogText(message) = message
    path = "C:\Проверки НДС\Расчёт.2026\Сечение_" & ChrW$(&H3B1) & "_2\RC_Section_NDM_execution_report.txt"
    message = WorkbookMessageParagraph(message, "Пошаговый отчет сохранен:" & vbCrLf & path)
    message = WorkbookMessageParagraph(message, "Время расчёта: 9.50 с.")
    TestCheck "presentation.report.paragraph", InStr(message, vbCrLf & vbCrLf & "Пошаговый отчет сохранен:" & vbCrLf & path & vbCrLf & vbCrLf & "Время расчёта: 9.50 с.") > 0
    TestCheck "presentation.report.path.exact", InStr(message, path) > 0
    TestCheck "presentation.report.idempotent", WorkbookDialogText(message) = message
    mReport = mReport & "MESSAGE_LAYOUT_EXAMPLE:" & vbCrLf & message & vbCrLf & "END_MESSAGE_LAYOUT_EXAMPLE" & vbCrLf
    text = WorkbookErrorMessage("Расчёт не выполнен", 12345, "Probe.Stage.2", "Неизвестный вариант. Проверьте " & ChrW$(&H3BB) & " = 0.125 и " & ChrW$(&H3B5) & " = 1.0E-8.")
    TestCheck "presentation.runtime.blocks", InStr(text, "Расчёт не выполнен:" & vbCrLf & vbCrLf & "Неизвестный вариант." & vbCrLf & vbCrLf & "Проверьте") > 0
    TestCheck "presentation.runtime.technical", InStr(text, vbCrLf & vbCrLf & "Источник ошибки: Probe.Stage.2." & vbCrLf & "Код ошибки: 12345.") > 0
    TestCheck "presentation.runtime.math", InStr(text, ChrW$(&H3BB) & " = 0.125") > 0 And InStr(text, ChrW$(&H3B5) & " = 1.0E-8") > 0
    TestCheck "presentation.runtime.idempotent", WorkbookDialogText(text) = text
    text = WorkbookErrorMessage("Расчёт не выполнен", 91, "Reader", "")
    TestCheck "presentation.runtime.empty.reason", InStr(text, "без описания причины") > 0 And InStr(text, "Код ошибки: 91.") > 0
    text = WorkbookErrorMessage("Расчёт не выполнен", 17, "Settings", "Настройка Solver.Method (лист Config, ячейка B15): Неизвестный метод. Выберите Newton или Secant.")
    TestCheck "presentation.settings.address", InStr(text, "Адреса ячеек:" & vbCrLf & "Config!B15") > 0
    TestCheck "presentation.settings.solution", InStr(text, "Возможное решение:" & vbCrLf & "Выберите Newton или Secant.") > 0
    TestCheck "presentation.settings.idempotent", WorkbookDialogText(text) = text
    text = SectionRecipeErrorText("Неверный отступ. Уменьшите значение.", "Сечения!C123", "Арматура: верхний ряд.")
    expected = text
    TestCheck "presentation.geometry.stable", WorkbookDialogText(text) = expected
    text = CalculationCompletionMessage(2, 0, "", "LC:2." & ChrW$(&H3B1))
    TestCheck "presentation.success.blocks", text = "Расчет завершен." & vbCrLf & vbCrLf & "Обработано сочетаний: 2." & vbCrLf & vbCrLf & "Результаты и пояснения всех проверок — на листе Results." & vbCrLf & vbCrLf & "Определяющее сочетание: LC:2." & ChrW$(&H3B1)
    TestCheck "presentation.empty.detail", InStr(CalculationCompletionMessage(2, 1, "", ""), "Показана только первая ошибка ввода:") = 0
    TestCheck "presentation.empty.paragraph", WorkbookMessageParagraph("Текст", "") = "Текст" And WorkbookMessageParagraph("", "Текст") = "Текст"
    text = WorkbookDialogText(vbCrLf & vbLf & "Первый" & vbCr & vbCrLf & vbLf & "Второй" & vbCrLf)
    TestCheck "presentation.newlines.normalized", text = "Первый" & vbCrLf & vbCrLf & "Второй"
    TestCheck "presentation.blank", WorkbookDialogText(vbCrLf & " " & vbLf) = ""
    text = "Введённое значение: ""Прочность: 2. Проверьте имя"". Выберите допустимый вариант."
    TestCheck "presentation.quoted.value", InStr(WorkbookDialogText(text), """Прочность: 2. Проверьте имя""") > 0
    text = "Слой «Устойчивость: 3. Проверьте имя» не найден. Откройте чертёж."
    TestCheck "presentation.quoted.layer", InStr(WorkbookDialogText(text), "«Устойчивость: 3. Проверьте имя»") > 0
    TestGeometryWarningParagraphs
    GoTo Finish
Failed:
    TestCheck "presentation.runtime", False
    mReport = mReport & "ERROR: " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Finish:
    RunProgramMessagePresentationTests = mReport & "TOTAL MESSAGE PRESENTATION: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет все типизированные статусы каждого этапа, смешанный отказ и InputErr,
' а также настоящий пакет с результатом проверки и четырьмя ошибками ввода.
' Повторное формирование сообщения не должно вызывать solver.
Public Function RunInputOnlyDialogTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    On Error GoTo Failed
    Dim stage As Long, status As Long, result As CCombinationResult, meta As CResultMeta, selected As CResultMeta
    Dim before As String, expected As Boolean, prefix As String, batch As CBatchSectionCalculator, text As String
    Dim settings As CSystemSettingsReader, units As CUnitSystem, provider As CMaterialModelProvider
    Dim profiles As CCalculationProfileCatalog, section As CSectionModel, i As Long, solves As Long, started As Double
    Dim geometry As CGeometryRegion, mesh As CFiberMeshBuilder, rebars As CRebarLayout
    Dim requiresAttention As Boolean, requiresNumerical As Boolean
    For stage = 1 To 8
        For status = rsUnset To rsInternalError
            Set result = New CCombinationResult: Set meta = New CResultMeta
            meta.SetResult status, rcInvalidInput, rkGeneric, "Причина InputErr/FAIL: " & CStr(stage) & "/" & CStr(status)
            meta.SourceKeys = "Solver.Method"
            TestSetInputDialogMeta result, stage, meta
            before = result.OverallMeta.DiagnosticText
            Set selected = result.InputErrorMeta
            expected = (status = rsInvalidInput Or status = rsInvalidConfiguration)
            prefix = "inputDialog.stage." & CStr(stage) & ".status." & CStr(status)
            TestCheck prefix & ".classification", (Not selected Is Nothing) = expected
            If Not selected Is Nothing Then
                TestCheck prefix & ".reason", selected.ResultComment = meta.ResultComment
                TestCheck prefix & ".keys", selected.SourceKeys = meta.SourceKeys
                selected.ResultComment = "Изменена только копия"
                TestCheck prefix & ".independent", result.InputErrorMeta.ResultComment = meta.ResultComment
            End If
            TestCheck prefix & ".unchanged", result.OverallMeta.DiagnosticText = before
        Next status
    Next stage
    Set result = New CCombinationResult: Set meta = New CResultMeta
    meta.SetResult rsCheckFailed, rcPhysicalLimitExceeded, rkDirectState, "Физические пределы превышены."
    result.SetDirectStateMeta meta
    meta.SetResult rsInvalidInput, rcInvalidInput, rkCapacity, "Для выбранного LoadPath соответствующая нагрузка должна быть ненулевой."
    TestSetInputDialogMeta result, 3, meta
    before = result.OverallMeta.ResultComment
    TestCheck "inputDialog.mixed.results.physical", InStr(before, "Физические пределы") > 0
    TestCheck "inputDialog.mixed.results.input", InStr(before, "LoadPath") > 0
    TestCheck "inputDialog.mixed.own.reason", result.InputErrorMeta.ResultComment = meta.ResultComment
    TestCheck "inputDialog.mixed.no.physical", InStr(result.InputErrorMeta.ResultComment, "Физические пределы") = 0
    TestCheck "inputDialog.mixed.results.unchanged", result.OverallMeta.ResultComment = before
    meta.SetResult rsInternalError, rcInternalError, rkStability, "Техническая остановка."
    result.SetStabilityMeta meta
    TestCheck "inputDialog.mixed.calcErr.retained", result.Status = "CalcErr" And Not result.InputErrorMeta Is Nothing
    TestRequiredAttentionMatrix

    Set batch = New CBatchSectionCalculator
    requiresAttention = True: text = CalculationBatchMessage(batch, , requiresAttention)
    TestCheck "inputDialog.empty.neutral", InStr(text, "Обработано сочетаний: 0.") > 0 And InStr(text, "Ошибка") = 0
    TestCheck "inputDialog.empty.no.stale.warning", Not requiresAttention
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    Set provider = New CMaterialModelProvider: provider.Initialize settings, units
    Set profiles = New CCalculationProfileCatalog: profiles.LoadFromWorkbook ThisWorkbook, False
    profiles.ProfileById("PR1").Initialize "PR1", "Проверка вывода", "", True, False, False, False
    Set geometry = TestFixtureBox(300#, 200#): Set mesh = New CFiberMeshBuilder
    mesh.BuildMesh geometry, 30#, 20#, 1
    Set rebars = New CRebarLayout
    rebars.AddBar "B1", -90#, -60#, 20#, 0#, "A400", "", geometry
    rebars.AddBar "B2", 90#, -60#, 20#, 0#, "A400", "", geometry
    rebars.AddBar "B3", -90#, 60#, 20#, 0#, "A400", "", geometry
    rebars.AddBar "B4", 90#, 60#, 20#, 0#, "A400", "", geometry
    Set section = TestBuildGeneratedSectionModel(mesh, rebars, "Test.InputDialog")
    batch.Initialize section, provider: Set batch.ProfileCatalog = profiles: batch.ApplySettings settings, units
    batch.AddCombination "CHECK_ONLY", 90000000#, 0#, 0#, "PR1", "", "lambda*N"
    batch.Execute: text = CalculationBatchMessage(batch, , requiresAttention)
    mReport = mReport & "CHECK_ONLY_ACTUAL_META: " & batch.ResultAt(1).OverallMeta.DiagnosticText & vbCrLf
    requiresNumerical = (batch.ResultAt(1).DirectStateMeta.InternalStatus = rsNumericalFailure)
    TestCheck "inputDialog.largeLoad.actualOutcome", requiresNumerical Or batch.ResultAt(1).DirectStateMeta.InternalStatus = rsCheckFailed
    TestCheck "inputDialog.largeLoad.typedAttention", requiresAttention = requiresNumerical
    TestCheck "inputDialog.largeLoad.noInputError", InStr(text, "Ошибок ввода:") = 0
    If requiresNumerical Then
        TestCheck "inputDialog.largeLoad.numericalNotice", InStr(text, "CHECK_ONLY") > 0 And InStr(text, "Первая численная неудача:") > 0
    Else
        TestCheck "inputDialog.largeLoad.physicalDetailsOnlyInResults", InStr(text, "CHECK_ONLY") = 0
    End If
    For i = 1 To 4
        batch.AddInvalidCombination "USER_" & CStr(i), "PR1", "", "Введите числовую нагрузку " & CStr(i) & ".", "Auto", i, "Лист ввода!AA" & CStr(40 + i) & ":AG" & CStr(40 + i)
    Next i
    batch.Execute: solves = SectionEquilibriumSolveCount()
    before = batch.ResultAt(1).OverallMeta.DiagnosticText
    text = CalculationBatchMessage(batch, , requiresAttention)
    TestCheck "inputDialog.batch.typed.warning", requiresAttention
    TestCheck "inputDialog.batch.count", InStr(text, "Ошибок ввода: 4.") > 0
    TestCheck "inputDialog.batch.first.notice", InStr(text, "Показана только первая ошибка ввода:") > 0
    TestCheck "inputDialog.batch.first.reason", InStr(text, "Введите числовую нагрузку 1.") > 0
    TestCheck "inputDialog.batch.first.address", InStr(text, "Лист ввода!AA41:AG41") > 0
    TestCheck "inputDialog.batch.only.first", InStr(text, "USER_2") = 0 And InStr(text, "нагрузку 2.") = 0
    TestCheck "inputDialog.batch.check.only.in.results", batch.ResultAt(1).Status <> "InputErr" And Len(batch.ResultAt(1).OverallMeta.ResultComment) > 0
    TestCheck "inputDialog.batch.largeLoad.sameOutcome", (batch.ResultAt(1).DirectStateMeta.InternalStatus = rsNumericalFailure) = requiresNumerical
    TestCheck "inputDialog.batch.largeLoad.typedNotice", (InStr(text, "CHECK_ONLY") > 0) = requiresNumerical
    TestCheck "inputDialog.batch.largeLoad.noAggregateComment", InStr(text, batch.ResultAt(1).OverallMeta.ResultComment) = 0
    TestCheck "inputDialog.batch.results.link", InStr(text, "на листе Results") > 0
    TestCheck "inputDialog.batch.results.unchanged", batch.ResultAt(1).OverallMeta.DiagnosticText = before
    TestCheck "inputDialog.batch.no.solver", SectionEquilibriumSolveCount() = solves
    mReport = mReport & "INPUT_ONLY_DIALOG_EXAMPLE:" & vbCrLf & text & vbCrLf & "END_INPUT_ONLY_DIALOG_EXAMPLE" & vbCrLf
    started = Timer
    For i = 1 To 100
        text = CalculationBatchMessage(batch)
    Next i
    mReport = mReport & "INPUT_ONLY_DIALOG_100_SUMMARIES_SECONDS: " & CStr(Timer - started) & vbCrLf
    GoTo Finish
Failed:
    TestCheck "inputDialog.runtime", False
    mReport = mReport & "ERROR: " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Finish:
    RunInputOnlyDialogTests = mReport & "TOTAL INPUT ONLY DIALOG: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function


' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает typed-операции с неизменным внешним словарём на всей матрице
' 11статусов ×20причин ×7видов; смена комментария не меняет классификацию.
Public Function RunTypedStatusPolicyTests(Optional ByRef passedCount As Long = 0, Optional ByRef failedCount As Long = 0) As String
    mPassed = 0: mFailed = 0: mReport = ""
    On Error GoTo Failed
    Dim policy As CResultStatusPolicy: Set policy = New CResultStatusPolicy
    Dim meta As CResultMeta, copy As CResultMeta, status As Long, code As Long, kind As Long
    Dim expected As Variant, priorities As Variant, expectedStatus As String, priority As Long
    Dim text As Variant, prefix As String, finished As Boolean, problem As Boolean
    expected = Array("CalcErr", "N/A", "N/A", "N/A", "OK", "FAIL", "NumFail", "InputErr", "InputErr", "OK", "CalcErr")
    priorities = Array(6, 0, 0, 0, 1, 2, 4, 5, 5, 1, 6)
    For status = rsUnset To rsInternalError
        For code = rcNone To rcSolverDidNotReturnState
            For kind = rkGeneric To rkStability
                expectedStatus = CStr(expected(status)): priority = CLng(priorities(status))
                If status = rsCheckFailed And (kind = rkCapacity Or kind = rkCrackFormation) And _
                        (code = rcInitialStateBeyondLimit Or code = rcCriterionNotReached) Then
                    expectedStatus = "BaseFail": priority = 3
                End If
                finished = (priority >= 1 And priority <= 3): problem = (priority >= 2)
                For Each text In Array("Пояснение InputErr NumFail FAIL", "Localized description", "")
                    Set meta = New CResultMeta: meta.SetResult status, code, kind, CStr(text), "diagnostic"
                    meta.SourceKeys = "One.Key;Shared.Key"
                    prefix = "typed.matrix." & status & "." & code & "." & kind & "." & Len(CStr(text))
                    TestCheck prefix & ".classification", policy.ExternalStatus(meta) = expectedStatus And _
                        policy.MetaPriority(meta) = priority And policy.IsMetaFinished(meta) = finished And _
                        policy.IsMetaFailedOverall(meta) = problem And policy.MetaRequiresAttention(meta) = (priority >= 4)
                    Set copy = policy.WorstResultMeta(rkGeneric, meta)
                    TestCheck prefix & ".preserved", copy.InternalStatus = meta.InternalStatus And copy.ResultCode = meta.ResultCode And _
                        copy.ResultKind = meta.ResultKind And copy.ResultComment = CStr(text) And copy.DiagnosticDetails = "diagnostic" And _
                        copy.SourceKeys = meta.SourceKeys And copy.Applies = meta.Applies And copy.Calculated = meta.Calculated
                    copy.ResultComment = "Changed clone"
                    TestCheck prefix & ".independent", meta.ResultComment = CStr(text)
                Next text
            Next kind
        Next code
    Next status
    Dim warning As CResultMeta, success As CResultMeta
    Set success = New CResultMeta: success.SetResult rsSuccess, rcCheckPassed, rkDirectState, "success"
    success.SourceKeys = "One.Key;Shared.Key"
    Set warning = New CResultMeta: warning.SetResult rsSuccessWithWarning, rcSearchBoundReached, rkCrackWidth, "warning"
    warning.SourceKeys = "Two.Key;shared.key"
    Set copy = policy.WorstResultMeta(rkGeneric, success, warning)
    TestCheck "typed.warning.forward", copy.InternalStatus = rsSuccessWithWarning And copy.ResultComment = "warning" And copy.SourceKeys = "One.Key;Shared.Key;Two.Key"
    Set copy = policy.WorstResultMeta(rkGeneric, warning, success)
    TestCheck "typed.warning.reverse", copy.InternalStatus = rsSuccessWithWarning And copy.ResultCode = rcSearchBoundReached
    Set meta = Nothing
    TestCheck "typed.missing", policy.MetaPriority(meta) = 6 And Not policy.IsMetaFinished(meta) And policy.IsMetaFailedOverall(meta) And policy.MetaRequiresAttention(meta)
    GoTo Finish
Failed:
    TestCheck "typed.runtime", False
    mReport = mReport & "RUNTIME ERROR: " & Err.Number & "; " & Err.Source & "; " & Err.Description & vbCrLf
Finish:
    passedCount = mPassed: failedCount = mFailed
    RunTypedStatusPolicyTests = mReport & "TOTAL TYPED STATUS POLICY: passed=" & mPassed & "; failed=" & mFailed & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет три категории обязательного внимания на каждом источнике результата,
' включая пустое/изменённое пояснение, адрес, копирование и отсутствие нового solver.
Private Sub TestRequiredAttentionMatrix()
    Dim stage As Long, status As Long, reason As Variant, result As CCombinationResult, meta As CResultMeta
    Dim batch As CBatchSectionCalculator, text As String, inputFlag As Boolean, attention As Boolean
    Dim expectedAttention As Boolean, expectedInput As Boolean, prefix As String, before As String, solves As Long
    solves = SectionEquilibriumSolveCount()
    For stage = 1 To 8
        For status = rsUnset To rsInternalError
            expectedInput = (status = rsInvalidInput Or status = rsInvalidConfiguration)
            expectedAttention = expectedInput Or status = rsInternalError Or status = rsNumericalFailure Or status = rsUnset
            For Each reason In Array("Причина FAIL InputErr", "Localized text", "")
                Set result = New CCombinationResult: Set meta = New CResultMeta
                meta.SetResult status, rcInternalError, rkGeneric, CStr(reason)
                meta.SourceKeys = "Solver.Method"
                TestSetInputDialogMeta result, stage, meta
                Set batch = New CBatchSectionCalculator
                batch.AddCombination "TYPED_CASE", -1#, 0#, 0#, "PR1", "", "Auto", 1, "TestInputs!A1:G1"
                batch.TestStoreResultForPresentation 1, result
                before = result.OverallMeta.DiagnosticText
                inputFlag = Not expectedInput: attention = Not expectedAttention
                text = CalculationBatchMessage(batch, inputFlag, attention)
                prefix = "attention." & stage & "." & status & "." & Len(CStr(reason))
                TestCheck prefix & ".flags", inputFlag = expectedInput And attention = expectedAttention
                TestCheck prefix & ".state", result.OverallMeta.DiagnosticText = before
                If expectedAttention Then
                    TestCheck prefix & ".address", InStr(1, text, "TestInputs!A1:G1", vbBinaryCompare) > 0
                    If Len(CStr(reason)) > 0 Then TestCheck prefix & ".ownReason", InStr(1, text, CStr(reason), vbBinaryCompare) > 0
                Else
                    TestCheck prefix & ".neutral", InStr(1, text, "TYPED_CASE", vbBinaryCompare) = 0
                End If
            Next reason
        Next status
    Next stage
    TestCheck "attention.no.solver", SectionEquilibriumSolveCount() = solves
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Заполняет один результат через штатные контейнеры; solver и разбор текста не используются.
Private Sub TestSetInputDialogMeta(ByVal result As CCombinationResult, ByVal stage As Long, ByVal meta As CResultMeta)
    Dim capacity As CCapacityResult, search As CLimitSearchResult, path As CLoadPathDescriptor
    Dim formation As CCrackFormationResult, width As CCrackWidthResult, longitudinal As CLongitudinalCrackResult
    Dim crack As CCrackResult
    Dim uncalculated As CResultMeta
    Select Case stage
        Case 1: result.SetWorkflowMeta meta
        Case 2: result.SetDirectStateMeta meta
        Case 3
            meta.SetResult meta.InternalStatus, meta.ResultCode, rkCapacity, meta.ResultComment
            meta.SourceKeys = "Solver.Method"
            Set search = New CLimitSearchResult: search.Initialize meta, "Auto", "LoadMultiplier", False, 0#, 0#, 0#, 0#, Nothing, "", ""
            Set path = New CLoadPathDescriptor: Set capacity = New CCapacityResult
            capacity.InitializeFromSearch search, path, False, 0#, 0#, Nothing, 0#, 0#
            result.StoreCapacityResult capacity
        Case 4, 5, 6, 7
            Set crack = New CCrackResult
            Select Case stage
                Case 4
                    Set uncalculated = New CResultMeta
                    uncalculated.SetResult meta.InternalStatus, meta.ResultCode, meta.ResultKind, meta.ResultComment, "", meta.Applies, False
                    uncalculated.SourceKeys = meta.SourceKeys
                    Set formation = New CCrackFormationResult: formation.InitializeUncalculated uncalculated
                    crack.Initialize formation, Nothing, Nothing, Nothing
                Case 5: crack.Initialize Nothing, meta, Nothing, Nothing
                Case 6
                    Set width = New CCrackWidthResult: width.InitializeFromCalculator Nothing, meta
                    crack.Initialize Nothing, Nothing, width, Nothing
                Case 7
                    Set longitudinal = New CLongitudinalCrackResult: longitudinal.Initialize meta, 0#, 0#, 0#
                    crack.Initialize Nothing, Nothing, Nothing, longitudinal
            End Select
            result.StoreCrackAggregateResult crack
        Case 8: result.SetStabilityMeta meta
    End Select
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Прогоняет все извлечённые сообщения рабочих компонентов через настоящий
' форматтер: повторный вызов устойчив, численные токены не теряются и не меняются.
Public Function RunProgramMessageCorpusTests(ByVal corpusPath As String) As String
    mPassed = 0: mFailed = 0: mReport = ""
    On Error GoTo Failed
    Dim file As Object, source As String, formatted As String, index As Long, expression As Object
    Set expression = CreateObject("VBScript.RegExp")
    expression.Global = True: expression.Pattern = "[+-]?[0-9]+([.,][0-9]+)?([Ee][+-]?[0-9]+)?"
    Set file = CreateObject("Scripting.FileSystemObject").OpenTextFile(corpusPath, 1, False, -1)
    Do Until file.AtEndOfStream
        source = file.ReadLine: index = index + 1
        formatted = WorkbookDialogText(source)
        TestCheck "presentation.corpus.idempotent." & CStr(index), WorkbookDialogText(formatted) = formatted
        TestCheck "presentation.corpus.numbers." & CStr(index), TestMessageNumberTokensEqual(source, formatted, expression)
        TestCheck "presentation.corpus.no.extra.blank." & CStr(index), InStr(formatted, vbCrLf & vbCrLf & vbCrLf) = 0
    Loop
    GoTo Finish
Failed:
    TestCheck "presentation.corpus.runtime", False
    mReport = mReport & "ERROR: " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Finish:
    If Not file Is Nothing Then file.Close
    RunProgramMessageCorpusTests = mReport & "MESSAGE_CORPUS_CASES: " & CStr(index) & vbCrLf & _
        "TOTAL MESSAGE CORPUS: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает числа с учётом повторов, но без требования одинакового порядка:
' форматтер настройки переносит адрес перед причиной, сохраняя оба значения.
Private Function TestMessageNumberTokensEqual(ByVal before As String, ByVal after As String, ByVal expression As Object) As Boolean
    Dim counts As Object, item As Object, key As Variant
    Set counts = CreateObject("Scripting.Dictionary")
    For Each item In expression.Execute(before)
        key = item.Value
        If Not counts.Exists(key) Then counts.Add key, 0
        counts(key) = counts(key) + 1
    Next item
    For Each item In expression.Execute(after)
        key = item.Value
        If Not counts.Exists(key) Then Exit Function
        counts(key) = counts(key) - 1
    Next item
    For Each key In counts.Keys
        If counts(key) <> 0 Then Exit Function
    Next key
    TestMessageNumberTokensEqual = True
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' повторное использование spec, атомарная регистрация и политика имён без зависимости от типа.
Public Function RunAudit02PolygonApiTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    On Error GoTo Failed
    Dim polygon As CPolygonSpec, crossed As CPolygonSpec, query As CSectionGeometryQuery
    Dim facade As CSectionDefinitionBuilder, path As CGeometryPath, inside As CGeometryPath, outside As CGeometryPath
    Dim definition As CSectionDefinition, contours As CSectionContours, before As String, errorCode As Long, i As Long
    Dim region As CConcreteRegion, n As Variant, angle As Double, x As Double, y As Double
    Set query = New CSectionGeometryQuery
    For Each n In Array(4, 5, 6, 9)
        Set polygon = New CPolygonSpec
        For i = 1 To CLng(n)
            angle = 2# * GEOM_PI * CDbl(i - 1) / CDbl(n)
            polygon.AddVertex "V" & CStr(i), 100# * Cos(angle), 100# * Sin(angle), i - 1#, "Vertex." & CStr(i)
        Next i
        polygon.Seal
        Set path = query.RoundedPolygon("Many." & CStr(n), polygon)
        TestCheck "P01.kernel." & CStr(n), path.IsClosed And path.IsSealed And path.SegmentCount = 2 * CLng(n) - 1
    Next n
    Set polygon = TestRectangleSpec(20#): before = TestSpecSnapshot(polygon)
    Set path = query.RoundedPolygon("Base", polygon)
    Set inside = query.OffsetRoundedPolygon("Inward", polygon, 5#)
    Set outside = query.OffsetRoundedPolygon("Outward", polygon, -5#)
    TestCheck "P05.spec.unchanged", TestSpecSnapshot(polygon) = before
    TestCheck "P05.paths.independent", Not path Is inside And Not inside Is outside
    Set region = query.RegionFromPath(path)
    TestCheck "P02.radius.geometry", Abs(region.Area - (10000# - 4# * 400# + GEOM_PI * 400#)) < 0.000001
    Set crossed = New CPolygonSpec
    crossed.AddVertex "A", 0#, 0#: crossed.AddVertex "B", 100#, 100#
    crossed.AddVertex "C", 0#, 100#: crossed.AddVertex "D", 100#, 0#: crossed.Seal
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "Atomic"
    On Error Resume Next
    Set path = facade.RoundedPolygon("Boundary", crossed)
    errorCode = Err.Number: Err.Clear
    On Error GoTo Failed
    TestCheck "P06.invalid.topology.rejected", errorCode <> 0
    Set path = facade.RoundedPolygon("Boundary", polygon)
    On Error Resume Next
    facade.AddPath path
    errorCode = Err.Number: Err.Clear
    On Error GoTo Failed
    TestCheck "P07.creator.registers.once", errorCode <> 0
    facade.Outer "Outer", path.ID, namingPolicy:=SequentialByKind, namingPrefix:="OUT_A"
    Set path = facade.RectanglePath("Other", 200#, 0#, 100#, 100#)
    facade.Outer "OtherOuter", path.ID, namingPolicy:=SequentialByKind, namingPrefix:="OUT_B"
    Set path = New CGeometryPath: path.Initialize "External"
    path.AddLine "External.Edge", 1#, 1#, 20#, 20#: path.Seal: facade.AddPath path
    Set definition = facade.Finish: Set contours = definition.MaterialContours
    TestCheck "P06.no.partial.path", definition.PathCount = 3
    TestCheck "P08.external.path", definition.Path("External").Length = Sqr(19# * 19# * 2#)
    TestCheck "G10.multiple.prefixes", contours.ContourID(1) = "OUT_A_ARC_1" And contours.ContourID(9) = "OUT_B_LINE_1"
    TestCheck "G10.loop.roles", contours.LoopID(1) = "Outer" And contours.LoopID(9) = "OtherOuter" And contours.LoopRole(9) = "Outer"
    On Error Resume Next
    definition.AnchorCoordinates "A", x, y
    errorCode = Err.Number: Err.Clear
    On Error GoTo Failed
    TestCheck "A03.no.implicit.anchor", errorCode <> 0
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "NamingConflict"
    Set path = facade.RectanglePath("Boundary", 0#, 0#, 100#, 100#)
    On Error Resume Next
    facade.Outer "Outer", path.ID, contourIDs:=Array("A", "B", "C", "D"), namingPolicy:=SequentialByKind
    errorCode = Err.Number: Err.Clear
    On Error GoTo Failed
    TestCheck "G10.explicit.policy.conflict", errorCode <> 0
    facade.Outer "Outer", path.ID, namingPolicy:=SequentialByKind
    Set path = facade.RectanglePath("Other", 200#, 0#, 100#, 100#)
    On Error Resume Next
    facade.Outer "Other", path.ID, namingPolicy:=SequentialByKind
    errorCode = Err.Number: Err.Clear
    On Error GoTo Failed
    TestCheck "G10.duplicate.prefix.rejected", errorCode <> 0
    GoTo Finished
Failed:
    TestCheck "runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    RunAudit02PolygonApiTests = mReport & "TOTAL AUDIT02 POLYGON API: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' независимые отступы, контактные группы, знаки и добавка 2,5 мм на каждое соседство.
Public Function RunIndependentRowOffsetBuilderTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    On Error GoTo Failed
    Dim cases As Variant, item As Variant, firstLocation As Variant, lastLocation As Variant
    Dim contact As CRebarContactBuilder, defaults As CRebarRuleOptions
    Dim second As CRebarRuleOptions, third As CRebarRuleOptions, copy As CRebarRuleOptions
    Dim expectedNormal As Double, expectedTangent As Double, groupBase As Double, intermediate As Double, tag As String
    cases = Array(Array(0#, 0#), Array(100#, 200#), Array(200#, 100#), Array(100#, 100#), Array(0#, 100#), Array(100#, 0#))
    For Each item In cases
        For Each firstLocation In Array("Stacked", "SideBySide")
            For Each lastLocation In Array("Stacked", "SideBySide")
                Set defaults = New CRebarRuleOptions
                Set contact = New CRebarContactBuilder: contact.Initialize 32#, -1#, 1#, defaults
                Set second = contact.AppendRow(20#, CStr(firstLocation), "EverySecondBar", "Offsets[2]", CDbl(item(0)))
                Set third = contact.AppendRow(16#, CStr(lastLocation), "EachBar", "Offsets[3]", CDbl(item(1)))
                tag = CStr(item(0)) & "/" & CStr(item(1)) & "/" & CStr(firstLocation) & "/" & CStr(lastLocation)
                expectedNormal = CDbl(item(1)): expectedTangent = 0#: intermediate = 0#: groupBase = 0#
                If CDbl(item(1)) = 0# Then groupBase = 32#
                If CDbl(item(1)) = CDbl(item(0)) Then
                    If CDbl(item(1)) > 0# Then
                        groupBase = 20#
                    ElseIf CStr(firstLocation) = CStr(lastLocation) Then
                        intermediate = 22.5
                    End If
                End If
                If groupBase > 0# Then
                    If CStr(lastLocation) = "Stacked" Then
                        expectedNormal = expectedNormal + groupBase / 2# + intermediate + 8# + 2.5
                    Else
                        expectedTangent = groupBase / 2# + intermediate + 8# + 2.5
                    End If
                End If
                TestCheck "offsets.normal." & tag, third.NormalOffset = -expectedNormal
                TestCheck "offsets.tangent." & tag, third.TangentOffset = expectedTangent
                TestCheck "offsets.binding." & tag, second.EverySecond And Not third.EverySecond
                Set copy = third.Clone
                TestCheck "offsets.clone." & tag, copy.ValidateRowOffset = third.ValidateRowOffset And copy.NormalOffset = third.NormalOffset
            Next lastLocation
        Next firstLocation
    Next item
    Set contact = New CRebarContactBuilder: contact.Initialize 32#, 1#, -1#, defaults
    Set second = contact.AppendRow(20#, "Stacked", "EachBar", "Tolerance[2]", 100#)
    Set third = contact.AppendRow(16#, "Stacked", "EachBar", "Tolerance[3]", 100# + GEOM_TOLERANCE / 2#)
    TestCheck "offsets.tolerance.same.group", Abs(third.NormalOffset - 120.5) < GEOM_TOLERANCE
    Set contact = New CRebarContactBuilder: contact.Initialize 32#, 1#, -1#, defaults
    Set third = contact.AppendRow(16#, "SideBySide", "EverySecondBar", "DisabledSecond[3]", 100#)
    TestCheck "offsets.disabled.second", third.NormalOffset = 100# And third.TangentOffset = 0# And third.EverySecond
    Set contact = New CRebarContactBuilder: contact.Initialize 32#, 1#, -1#, defaults
    Set third = contact.AppendRow(16#, "SideBySide", "EachBar", "DisabledSecondZero[3]")
    TestCheck "offsets.profile.disabled.second", third.NormalOffset = 0# And third.TangentOffset = -26.5
    GoTo Finished
Failed:
    TestCheck "offsets.runtime; " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    RunIndependentRowOffsetBuilderTests = mReport & "TOTAL INDEPENDENT ROW OFFSET BUILDER: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' контактный helper меняет только компоненты смещения и EverySecond.
Public Function RunAudit02OptionsTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    On Error GoTo Failed
    Dim options As CRebarRuleOptions, errorCode As Long, axis As Variant, sign As Variant
    Set options = New CRebarRuleOptions
    options.StartSetback = 100#: options.EndSetback = 150#: options.EndMode = ExcludeBoth
    options.PositionOrder = FromPathMiddle: options.IncomingTangentAtJoint = True
    options.SourceNormalSign = -1#: options.SuccessfulProjectionOrdinal = True
    options.SkipMissingProjection = True: options.NewPublicationRoot = True
    options.AnnotationGroup = "Group" ' В одной подписи группы суммируем стержни каждого диаметра, включая все её ряды.
    options.Publication = PublishPerSourcePosition ' В списке стержней нумеруем сначала основной, затем связанные с ним; координаты не меняются.
    options.Coincidence = KeepPublishedOwner ' При совпадении центров сохраняем стержень, добавленный первым.
    TestCheck "R08.publication.only", options.StartSetback = 100# And options.EndMode = ExcludeBoth And options.SourceNormalSign = -1#
    For Each axis In Array(RowNormal, RowTangent)
        For Each sign In Array(-1#, 1#)
            options.ConfigureDependentContact 20#, 16#, CLng(axis), CDbl(sign), False
            TestCheck "R03.contact.18." & CStr(axis) & "." & CStr(sign), _
                IIf(axis = RowNormal, options.NormalOffset, options.TangentOffset) = sign * 18# And _
                IIf(axis = RowNormal, options.TangentOffset, options.NormalOffset) = 0#
            options.ConfigureDependentContact 20#, 12#, CLng(axis), CDbl(sign), True, 16#
            TestCheck "R04.contact.32." & CStr(axis) & "." & CStr(sign), IIf(axis = RowNormal, options.NormalOffset, options.TangentOffset) = sign * 32#
            options.ConfigureDependentContact 20#, 12#, CLng(axis), CDbl(sign), False
            TestCheck "R04.other.direction.16." & CStr(axis) & "." & CStr(sign), IIf(axis = RowNormal, options.NormalOffset, options.TangentOffset) = sign * 16#
        Next sign
    Next axis
    TestCheck "R08.independent.fields", options.StartSetback = 100# And options.EndSetback = 150# And options.EndMode = ExcludeBoth And _
        options.PositionOrder = FromPathMiddle And options.IncomingTangentAtJoint And options.SuccessfulProjectionOrdinal And _
        options.SkipMissingProjection And options.NewPublicationRoot And options.SourceNormalSign = -1# And _
        options.AnnotationGroup = "Group" And options.Publication = PublishPerSourcePosition And options.Coincidence = KeepPublishedOwner
    On Error Resume Next
    options.ConfigureDependentContact 20#, 12#, RowNormal, 0#, False
    errorCode = Err.Number: Err.Clear
    On Error GoTo Failed
    TestCheck "R08.reject.invalid.sign", errorCode <> 0 And options.TangentOffset = 16# And options.NormalOffset = 0#
    GoTo Finished
Failed:
    TestCheck "runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    RunAudit02OptionsTests = mReport & "TOTAL AUDIT02 OPTIONS: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' положительные/неотрицательные длины сохраняют обязательность ввода и его источник.
Public Function RunAudit02InputTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    On Error GoTo Failed
    Dim values(1 To 4, 1 To 1) As Variant, addresses(1 To 4, 1 To 1) As Variant
    Dim table As CSectionInputTable, units As CUnitSystem, number As Double, errorCode As Long, reason As String, i As Long
    values(1, 1) = 25#: values(2, 1) = 0#: values(3, 1) = -1#: values(4, 1) = "-"
    For i = 1 To 4: addresses(i, 1) = "Config!A" & CStr(i): Next i
    Set table = New CSectionInputTable: table.Initialize "Input", values, addresses
    Set units = New CUnitSystem: units.InitializeDefaults
    TestCheck "input.positive", table.PositiveLength(1, 1, units) = 25#
    TestCheck "input.nonnegative.zero", table.NonNegativeLength(2, 1, units) = 0#
    For i = 2 To 4
        On Error Resume Next
        number = table.PositiveLength(i, 1, units)
        errorCode = Err.Number: reason = Err.Description: Err.Clear
        On Error GoTo Failed
        TestCheck "input.reject.address." & CStr(i), errorCode <> 0 And InStr(reason, "Config!A" & CStr(i)) > 0
    Next i
    GoTo Finished
Failed:
    TestCheck "runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    RunAudit02InputTests = mReport & "TOTAL AUDIT02 INPUT: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' три линейных постановки, действительная грань и смешанная длина пути.
Public Function RunAudit02AnnotationTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    On Error GoTo Failed
    Dim facade As CSectionDefinitionBuilder, definition As CSectionDefinition, path As CGeometryPath, polygon As CPolygonSpec
    Dim builder As CSectionAnnotationBuilder, annotations As CSectionAnnotations, rebars As CRebarLayout, edge As TRegionEdge
    Dim x As Double, y As Double, errorCode As Long
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "Dimensions"
    Set path = facade.RectanglePath("Boundary", -100#, -100#, 1000#, 1000#): facade.Outer "Outer", path.ID
    facade.Anchor "A", 0#, 0#: facade.Anchor "B", 300#, 400#
    facade.Dimension "H", "A", "B", "X", "Horizontal"
    facade.Dimension "V", "A", "B", "Y", "Vertical"
    facade.Dimension "L", "A", "B", "L", "Aligned"
    Set polygon = TestRectangleSpec(20#): Set path = facade.RoundedPolygon("Filleted", polygon)
    facade.PathAnchor "Touch", path.ID, "Edge.A", 0#
    Set path = New CGeometryPath: path.Initialize "Mixed"
    path.AddLine "Line", 0#, 0#, 100#, 0#: path.AddArc "Arc", 100#, 50#, 50#, -GEOM_PI / 2#, GEOM_PI / 2#: path.Seal
    facade.AddPath path: facade.PathLength "PL", path.ID, "A", "PL"
    Set definition = facade.Finish
    Set builder = New CSectionAnnotationBuilder: Set annotations = New CSectionAnnotations: Set rebars = New CRebarLayout
    builder.Build definition, rebars, annotations
    TestCheck "A01.horizontal.300", annotations.Value(1) = 300#
    TestCheck "A01.vertical.400", annotations.Value(2) = 400#
    TestCheck "A01.aligned.500", annotations.Value(3) = 500#
    TestCheck "A02.aligned.endpoints", annotations.StartX(3) = 0# And annotations.EndX(3) = 300# And annotations.EndY(3) = 400#
    TestCheck "A02.aligned.normal", annotations.NormalX(3) = -0.8 And annotations.NormalY(3) = 0.6
    definition.AnchorCoordinates "Touch", x, y
    TestCheck "A03.virtual.not.touch", x <> polygon.X(1) Or y <> polygon.Y(1)
    TestCheck "A04.mixed.length", Abs(annotations.Value(4) - (100# + 25# * GEOM_PI)) < 0.000001
    TestCheck "A04.not.chord", Abs(annotations.Value(4) - Sqr(150# * 150# + 50# * 50#)) > 10#
    builder.Build definition, rebars, annotations
    TestCheck "A05.repeat", annotations.Count = 4 And rebars.Count = 0
    GoTo Finished
Failed:
    TestCheck "runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    RunAudit02AnnotationTests = mReport & "TOTAL AUDIT02 ANNOTATIONS: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' полные примеры для комментарийного гайда и реальная наклонная грань.
Public Function RunAudit02GuideExamplesTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    On Error GoTo Failed
    Dim definition As CSectionDefinition, builder As CSectionBuilder, model As CSectionModel
    Dim mesh As TSectionMeshOptions, path As CGeometryPath, edge As TRegionEdge, i As Long
    Dim annotations As CSectionAnnotations, annotationBuilder As CSectionAnnotationBuilder, rebars As CRebarLayout
    Dim facade As CSectionDefinitionBuilder, polygon As CPolygonSpec, errorCode As Long, reason As String
    mesh.StepX = 50#: mesh.StepY = 50#: mesh.BoundarySubdivisions = 2: mesh.MaterialID = 1
    Set definition = TestBuildAudit02DemoDefinition()
    Set builder = New CSectionBuilder: Set model = builder.Build(definition, mesh)
    TestCheck "D01.demo.bars", model.RebarCount = 15
    Set path = definition.Path("Boundary"): path.CopySegment path.SegmentIndex("Edge.TR"), edge
    For i = 1 To model.Annotations.Count
        If model.Annotations.AnnotationID(i) = "DIM_SLOPE" Then
            TestCheck "A02.slope.start", model.Annotations.StartX(i) = edge.X1 And model.Annotations.StartY(i) = edge.Y1
            TestCheck "A02.slope.end", model.Annotations.EndX(i) = edge.X2 And model.Annotations.EndY(i) = edge.Y2
            TestCheck "A02.slope.length", Abs(model.Annotations.Value(i) - Sqr((edge.X2 - edge.X1) ^ 2 + (edge.Y2 - edge.Y1) ^ 2)) < 0.000001
            TestCheck "A02.slope.parallel", Abs((edge.X2 - edge.X1) * model.Annotations.NormalX(i) + _
                (edge.Y2 - edge.Y1) * model.Annotations.NormalY(i)) < 0.000001
        End If
    Next i
    Set definition = TestBuildAudit02OperationsDefinition()
    Set model = builder.Build(definition, mesh)
    TestCheck "D01.operations.build", model.RebarCount > 20 And model.ConcreteCount > 0
    Set annotations = New CSectionAnnotations: Set annotationBuilder = New CSectionAnnotationBuilder
    annotationBuilder.Build definition, builder.Rebars, annotations, builder.Rules
    TestCheck "A05.annotations.repeat", annotations.Count = model.Annotations.Count And builder.Rebars.Count = model.RebarCount
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "Zero"
    Set path = facade.RectanglePath("Boundary", 0#, 0#, 100#, 100#): facade.Outer "Outer", path.ID
    facade.Anchor "A", 20#, 20#: facade.Anchor "B", 20#, 20#
    facade.Source "ZERO", "Config!TestZero": facade.Dimension "ZERO", "A", "B", "L", "Aligned"
    Set definition = facade.Finish: Set rebars = New CRebarLayout
    On Error Resume Next
    annotationBuilder.Build definition, rebars, annotations
    errorCode = Err.Number: reason = Err.Description: Err.Clear
    On Error GoTo Failed
    TestCheck "A06.zero.rejected.address", errorCode <> 0 And InStr(reason, "ZERO") > 0 And InStr(reason, "Config!TestZero") > 0
    TestCheck "A06.no.replacement", annotations.Count = 0
    Set polygon = New CPolygonSpec
    polygon.AddVertex "TL", -300#, 500#, 300#: polygon.AddVertex "TR", 300#, 500#, 300#
    polygon.AddVertex "BR", 300#, -500#, 300#: polygon.AddVertex "BL", -300#, -500#, 300#: polygon.Seal
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "CollapsedDimension"
    Set path = facade.RoundedPolygon("OuterPath", polygon): facade.Outer "Outer", path.ID
    Set path = facade.OffsetRoundedPolygon("OffsetPath", polygon, 50#)
    facade.PathAnchor "A", path.ID, "Edge.TL", 0#: facade.PathAnchor "B", path.ID, "Edge.TL", 1#
    facade.Source "COLLAPSED", "Config!TestCollapsed": facade.Dimension "COLLAPSED", "A", "B", "L", "Aligned"
    Set definition = facade.Finish
    On Error Resume Next
    annotationBuilder.Build definition, rebars, annotations
    errorCode = Err.Number: reason = Err.Description: Err.Clear
    On Error GoTo Failed
    TestCheck "A06.collapsed.rejected.address", errorCode <> 0 And InStr(reason, "COLLAPSED") > 0 And InStr(reason, "Config!TestCollapsed") > 0
    TestCheck "A06.collapsed.no.replacement", annotations.Count = 0
    GoTo Finished
Failed:
    TestCheck "runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    RunAudit02GuideExamplesTests = mReport & "TOTAL AUDIT02 GUIDE EXAMPLES: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' пример из Audit02, без подбора достаточного армирования.
Private Function TestBuildAudit02DemoDefinition() As CSectionDefinition
    Dim facade As CSectionDefinitionBuilder, polygon As CPolygonSpec
    Dim shapePath As CGeometryPath, offsetPath As CGeometryPath, rightAxis As CGeometryPath
    Dim options As CRebarRuleOptions, b As Double, h As Double, w As Double
    b = 600#: h = 1000#: w = 200#
    Set polygon = New CPolygonSpec
    polygon.AddVertex "TL", -b / 2#, h / 2#, 80#, outgoingEdgeID:="Edge.TL" ' Верхняя грань для привязки бокового ряда, мм.
    polygon.AddVertex "TR", b / 2#, h / 2#, 120#, outgoingEdgeID:="Edge.TR" ' Наклонная грань для размерной аннотации, мм.
    polygon.AddVertex "RightNose", b / 2# + w, 0#, 80#
    polygon.AddVertex "BR", b / 2#, -h / 2#, 120#, outgoingEdgeID:="Edge.BR" ' Нижняя грань для конца бокового ряда, мм.
    polygon.AddVertex "BL", -b / 2#, -h / 2#, 80#, outgoingEdgeID:="Edge.BL" ' Левая грань от нижнего угла к верхнему, мм.
    polygon.Seal
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "Audit02Demo"
    Set shapePath = facade.RoundedPolygon("Boundary", polygon)
    facade.Outer "OUTER", shapePath.ID
    Set offsetPath = facade.OffsetRoundedPolygon("Offset.H.Right", polygon, 50#)
    Set rightAxis = facade.Chain("H.Right", offsetPath.ID, "Edge.TL", 1#, "Edge.BR", 0#)
    Set options = New CRebarRuleOptions
    options.AnnotationGroup = "H.Right" ' В одной подписи группы суммируем стержни каждого диаметра, включая все её ряды.
    options.Publication = PublishPerSourcePosition ' В списке стержней нумеруем сначала основной, затем связанные с ним; координаты не меняются.
    options.Coincidence = KeepPublishedOwner ' При совпадении центров сохраняем стержень, добавленный первым.
    options.EndMode = ExcludeBoth: options.StartSetback = 100#: options.EndSetback = 150#
    options.IncomingTangentAtJoint = True
    facade.RebarByCount "H.Right.Main", rightAxis.ID, 15, 20#, options
    facade.RebarRowExtentLabel "H.Right", "H.Right.Main", 1#, 0#, True
    facade.PathAnchor "Slope.A", shapePath.ID, "Edge.TR", 0#
    facade.PathAnchor "Slope.B", shapePath.ID, "Edge.TR", 1#
    facade.Dimension "DIM_SLOPE", "Slope.A", "Slope.B", "L", "Aligned"
    facade.HorizontalDistance "DIM_B", "B", -b / 2#, b / 2#, h / 2#, 1#
    facade.VerticalExtent "DIM_H", shapePath.ID, "H", -b / 2#, -1#
    Set TestBuildAudit02DemoDefinition = facade.Finish
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' независимые примитивы, signed-left offset, распределения и проекция.
Private Function TestBuildAudit02OperationsDefinition() As CSectionDefinition
    Dim facade As CSectionDefinitionBuilder, path As CGeometryPath, options As CRebarRuleOptions
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "OperationsExample"
    Set path = facade.AdjacentRectangles("Boundary", Array(Array(0#, 0#, 1000#, 1000#), Array(1000#, 0#, 200#, 600#)))
    facade.Outer "Outer", path.ID
    Set path = facade.CirclePath("Hole", 900#, 800#, 50#): facade.Opening "Opening", path.ID, "Outer"
    Set path = facade.LinePath("MainAxis", 100#, 100#, 700#, 100#)
    Set path = facade.ArcPath("Arc", 500#, 600#, 100#, 0#, GEOM_PI / 2#)
    Set path = facade.Offset("ProjectedAxis", "MainAxis", 100#)
    Set path = facade.Transform("Rotated", "MainAxis", 900#, 100#, GEOM_PI / 2#, True)
    Set path = facade.ConstrainLine("Limited", "MainAxis", 1#, 0#, 150#, 1#, 0#, 650#)
    Set path = facade.LinePath("MaxStepAxis", 100#, 300#, 700#, 300#)
    Set path = facade.LinePath("FixedStepAxis", 100#, 400#, 700#, 400#)
    Set path = facade.LinePath("ExplicitAxis", 100#, 500#, 700#, 500#)
    facade.PathGroup "TwoLines", Array("MaxStepAxis", "FixedStepAxis")
    Set options = New CRebarRuleOptions
    options.AnnotationGroup = "Bottom" ' В одной подписи группы суммируем стержни каждого диаметра, включая все её ряды.
    options.Publication = PublishPerSourcePosition ' В списке стержней нумеруем сначала основной, затем связанные с ним; координаты не меняются.
    options.Coincidence = KeepPublishedOwner ' При совпадении центров сохраняем стержень, добавленный первым.
    facade.RebarByCount "Main", "MainAxis", 5, 20#, options
    Set options = New CRebarRuleOptions
    options.AnnotationGroup = "Bottom" ' В одной подписи группы суммируем стержни каждого диаметра, включая все её ряды.
    options.Publication = PublishPerSourcePosition ' В списке стержней нумеруем сначала основной, затем связанные с ним; координаты не меняются.
    options.Coincidence = KeepPublishedOwner ' При совпадении центров сохраняем стержень, добавленный первым.
    options.ConfigureDependentContact 20#, 16#, RowNormal, 1#, False
    facade.RebarDependent "Contact", "Main", 16#, options
    Set options = New CRebarRuleOptions
    options.NormalOffset = 40#: options.TangentOffset = 10#
    facade.RebarDependent "Shifted", "Main", 12#, options
    Set options = New CRebarRuleOptions: options.DirectionY = 1#
    facade.RebarProjected "Projected", "ProjectedAxis", "Main", 16#, options
    facade.RebarByMaxStep "MaxStep", "MaxStepAxis", 150#, 16#
    facade.RebarByFixedStep "FixedStep", "FixedStepAxis", 150#, 16#
    facade.RebarAtStations "Explicit", "ExplicitAxis", Array(0#, 150#, 600#), 12#
    facade.Anchor "A", 100#, 100#: facade.Anchor "B", 700#, 500#
    facade.TransformAnchor "MovedA", "A", 0#, 700#
    facade.Dimension "H", "A", "B", "X", "Horizontal"
    facade.Dimension "V", "A", "B", "Y", "Vertical"
    facade.Dimension "L", "A", "B", "L", "Aligned"
    facade.PathLength "Length", "TwoLines", "MovedA", "SumL"
    facade.Radius "Radius", "Arc", "Arc", 0.5, "R"
    facade.Diameter "Diameter", "Hole", "Hole", 0.5, "D"
    facade.Text "Note", "MovedA", "Учебный пример"
    facade.RebarRowLabel "Bottom", "Main", 0#, -1#
    Set TestBuildAudit02OperationsDefinition = facade.Finish
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' реальная Excel-схема наклонного размера и проверка повторного Draw.
Public Function TestDumpAudit02DemoPreview(ByVal outputDirectory As String) As String
    Dim definition As CSectionDefinition, builder As CSectionBuilder, mesh As TSectionMeshOptions, model As CSectionModel
    Set definition = TestBuildAudit02DemoDefinition(): Set builder = New CSectionBuilder
    mesh.StepX = 50#: mesh.StepY = 50#: mesh.BoundarySubdivisions = 2: mesh.MaterialID = 1
    Set model = builder.Build(definition, mesh)
    Dim first As String, second As String
    first = modTestShapeArchitecture.TestPreparedPreviewSnapshot("Audit02Demo", model, outputDirectory)
    second = modTestShapeArchitecture.TestPreparedPreviewSnapshot("Audit02Demo", model, outputDirectory)
    If first <> second Then Err.Raise vbObjectError + 5490, "modTestAudit02Api", "Повторный Draw изменил данные/Shapes."
    TestDumpAudit02DemoPreview = first
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' настройка контактного ряда отвергает нулевые диаметры,
' а при отказе не изменяет переданные параметры.
Public Function RunAudit02CloseoutContactTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    On Error GoTo Failed
    Dim options As CRebarRuleOptions, cases As Variant, item As Variant, code As Long, reason As String
    Set options = New CRebarRuleOptions
    options.SourceReference = "Program.Contact": options.StartSetback = 100#: options.EndSetback = 150#
    options.AnnotationGroup = "Group" ' В одной подписи группы суммируем стержни каждого диаметра, включая все её ряды.
    options.Publication = PublishPerSourcePosition ' В списке стержней нумеруем сначала основной, затем связанные с ним; координаты не меняются.
    options.Coincidence = KeepPublishedOwner ' При совпадении центров сохраняем стержень, добавленный первым.
    options.NormalOffset = 12#: options.TangentOffset = 34#: options.EverySecond = True
    cases = Array(Array(0#, 16#, RowNormal, 1#, 0#), Array(-20#, 16#, RowNormal, 1#, 0#), _
        Array(20#, 0#, RowNormal, 1#, 0#), Array(20#, -16#, RowNormal, 1#, 0#), _
        Array(20#, 16#, RowNormal, 1#, -1#), Array(20#, 16#, 0, 1#, 0#), Array(20#, 16#, RowNormal, 0#, 0#))
    For Each item In cases
        On Error Resume Next
        options.ConfigureDependentContact CDbl(item(0)), CDbl(item(1)), CLng(item(2)), CDbl(item(3)), False, CDbl(item(4))
        code = Err.Number: reason = Err.Description: Err.Clear
        On Error GoTo Failed
        TestCheck "contact.reject.address." & CStr(item(0)) & "." & CStr(item(1)) & "." & CStr(item(2)) & "." & CStr(item(3)), _
            code <> 0 And InStr(reason, "Program.Contact") > 0
        TestCheck "contact.reject.atomic", options.NormalOffset = 12# And options.TangentOffset = 34# And options.EverySecond
        TestCheck "contact.unrelated.fields", options.StartSetback = 100# And options.EndSetback = 150# And _
            options.AnnotationGroup = "Group" And options.Publication = PublishPerSourcePosition And options.Coincidence = KeepPublishedOwner
    Next item
    options.ConfigureDependentContact 20#, 16#, RowNormal, -1#, False
    TestCheck "contact.valid.18", options.NormalOffset = -18# And options.TangentOffset = 0# And Not options.EverySecond
    options.ConfigureDependentContact 20#, 12#, RowTangent, 1#, True, 16#
    TestCheck "contact.valid.32", options.NormalOffset = 0# And options.TangentOffset = 32# And options.EverySecond
    GoTo Finished
Failed:
    TestCheck "runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    RunAudit02CloseoutContactTests = mReport & "TOTAL AUDIT02 CLOSEOUT CONTACT: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' адресная матрица размеров, неактивных полей, знаков и порядка диагностик Rectangles.
Public Function RunAudit02CloseoutRectanglesInputTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    Dim records As String
    records = TestCloseoutRectanglesInputMatrix()
    RunAudit02CloseoutRectanglesInputTests = mReport & "TOTAL AUDIT02 CLOSEOUT RECTANGLES INPUT: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' полный снимок значений/ошибок той же матрицы для побайтной сверки baseline/candidate.
Public Function TestDumpAudit02CloseoutInputContract() As String
    mPassed = 0: mFailed = 0: mReport = ""
    TestDumpAudit02CloseoutInputContract = TestCloseoutRectanglesInputMatrix()
    If mFailed > 0 Then Err.Raise vbObjectError + 5491, "modTestAudit02Api", mReport
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' только три схемы Rectangles, без полного набора UI/ON-OFF.
Public Function TestDumpAudit02CloseoutRectanglesPreview(ByVal outputDirectory As String) As String
    Dim fixture As Variant
    For Each fixture In Array("Rectangle", "LSection", "TwoRectangles")
        TestDumpAudit02CloseoutRectanglesPreview = TestDumpAudit02CloseoutRectanglesPreview & _
            modTestShapeArchitecture.TestDumpShapePreviewSnapshot(outputDirectory, CStr(fixture))
    Next fixture
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Записываем номер ошибки, её источник и полный текст с фактическим адресом ячейки.
' Сравнение должно выявлять любое различие.
Private Function TestCloseoutRectanglesInputMatrix() As String
    Dim keyed As CSectionInputReader, fieldKeys As Variant
    Dim config As Object, saved As Variant, range As Object, block As Variant
    Dim settings As CSystemSettingsReader, units As CUnitSystem, definition As CSectionDefinition
    Dim reader As CSectionInputSnapshotReader, context As CSectionRecipeContext, table As CSectionInputTable
    Dim fixture As Variant, valueKind As Variant, unit As Variant, aliasName As Variant
    Dim row As Long, column As Long, maxColumn As Long, offsetRow As Long, shapeRow As Long
    Dim code As Long, source As String, reason As String, expectedAddress As String, area As Double
    Dim result As String, factor As Double, pair As Variant, number As Double, otherCode As Long, otherReason As String
    fieldKeys = Array("H1", "B1", "H2", "B2")
    On Error GoTo Failed
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = TestFixtureSaveInputs(config)
    For Each unit In Array("mm", "cm", "m")
        TestFixtureRestoreInputs config, saved
        ThisWorkbook.Names("rngUnitSettings").RefersToRange.Cells(2, 2).Value2 = CStr(unit)
        For Each fixture In Array("Rectangle", "LSection", "TwoRectangles")
            modTestShapeArchitecture.TestConfigureShapeFixture "Circle"
            modTestShapeArchitecture.TestConfigureShapeFixture CStr(fixture)
            Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
            Set units = New CUnitSystem: units.LoadFromSettings settings: factor = units.InputLengthToInternal(1#)
            Set range = ThisWorkbook.Names("rngRectanglesGeometry").RefersToRange: block = range.Formula
            Set reader = New CSectionInputSnapshotReader
            Set context = reader.Capture(ThisWorkbook, Array("rngRectanglesGeometry"), units): Set table = context.Table("rngRectanglesGeometry"): Set keyed = context.Reader("rngRectanglesGeometry")
            offsetRow = keyed.RowIndex("General", "XOffset")
            shapeRow = keyed.RowIndex("General", "Shape")
            Set definition = TestRectanglesInputOutcome(settings, units, code, source, reason)
            TestCheck "rect.valid." & CStr(unit) & "." & CStr(fixture), code = 0
            If code <> 0 Then GoTo Failed
            area = definition.ConcreteDomain.Area
            result = result & TestInputContractRow(CStr(unit), CStr(fixture), "base", code, source, reason, area)
            maxColumn = 4: If CStr(fixture) = "Rectangle" Then maxColumn = 2
            For column = 1 To maxColumn
                expectedAddress = keyed.SourceReference("General", CStr(fieldKeys(column - 1)))
                For Each valueKind In Array("positive", "zero", "negative", "empty", "text")
                    range.Formula = block
                    Select Case CStr(valueKind)
                        Case "zero": TestFixtureInputCell(range, "General", CStr(fieldKeys(column - 1))).Value2 = 0#
                        Case "negative": TestFixtureInputCell(range, "General", CStr(fieldKeys(column - 1))).Value2 = -1# / factor
                        Case "empty": TestFixtureInputCell(range, "General", CStr(fieldKeys(column - 1))).ClearContents
                        Case "text": TestFixtureInputCell(range, "General", CStr(fieldKeys(column - 1))).Value2 = "bad dimension"
                    End Select
                    Set definition = TestRectanglesInputOutcome(settings, units, code, source, reason)
                    If CStr(valueKind) = "positive" Then
                        TestCheck "rect.positive." & CStr(unit) & "." & CStr(fixture) & "." & CStr(column), code = 0
                        TestCheck "rect.unit.once", definition.ConcreteDomain.Area = area
                    Else
                        TestCheck "rect.reject.address." & CStr(unit) & "." & CStr(fixture) & "." & CStr(column) & "." & CStr(valueKind), _
                            code <> 0 And InStr(reason, expectedAddress) > 0
                    End If
                    result = result & TestInputContractRow(CStr(unit), CStr(fixture), CStr(column) & "." & CStr(valueKind), code, source, reason, 0#)
                Next valueKind
            Next column
            range.Formula = block ' У OneRect нижние размеры и смещение не используются; у обоих вариантов TwoRect смещение активно.
            If CStr(fixture) = "Rectangle" Then
                TestFixtureInputCell(range, "General", "H2").Value2 = "inactive H2": TestFixtureInputCell(range, "General", "B2").Value2 = CVErr(xlErrValue)
            End If
            If CStr(fixture) = "Rectangle" Then
                range.Cells(offsetRow, 3).Value2 = "inactive offset"
                Set definition = TestRectanglesInputOutcome(settings, units, code, source, reason)
                TestCheck "rect.inactive.fields." & CStr(unit) & "." & CStr(fixture), code = 0 And definition.ConcreteDomain.Area = area
                result = result & TestInputContractRow(CStr(unit), CStr(fixture), "inactive", code, source, reason, 0#)
            Else
                For Each pair In Array(-100#, 0#, 100#)
                    range.Formula = block: range.Cells(offsetRow, 3).Value2 = CDbl(pair) / factor
                    Set definition = TestRectanglesInputOutcome(settings, units, code, source, reason)
                    TestCheck "rect.signed.offset." & CStr(unit) & "." & CStr(pair), code = 0
                    result = result & TestInputContractRow(CStr(unit), CStr(fixture), "offset." & CStr(pair), code, source, reason, 0#)
                Next pair
            End If
            ' Во взаимосвязанных полях сначала проверяется, что второе поле числовое,
            ' и только затем — положительность первого.
            For Each pair In Array(1, 3)
                If CLng(pair) <= maxColumn Then
                    range.Formula = block: TestFixtureInputCell(range, "General", CStr(fieldKeys(CLng(pair) - 1))).Value2 = 0#
                    TestFixtureInputCell(range, "General", CStr(fieldKeys(CLng(pair)))).Value2 = "bad dimension"
                    Set definition = TestRectanglesInputOutcome(settings, units, code, source, reason)
                    TestCheck "rect.pair.numeric.before.positive." & CStr(unit) & "." & CStr(fixture) & "." & CStr(pair), _
                        code <> 0 And InStr(reason, keyed.SourceReference("General", CStr(fieldKeys(CLng(pair))))) > 0
                    result = result & TestInputContractRow(CStr(unit), CStr(fixture), "pair." & CStr(pair), code, source, reason, 0#)
                    Set context = reader.Capture(ThisWorkbook, Array("rngRectanglesGeometry"), units): Set table = context.Table("rngRectanglesGeometry"): Set keyed = context.Reader("rngRectanglesGeometry")
                    On Error Resume Next
                    number = keyed.PositiveLength("General", CStr(fieldKeys(CLng(pair) - 1)))
                    otherCode = Err.Number: otherReason = Err.Description: Err.Clear
                    On Error GoTo Failed
                    TestCheck "rect.direct.helper.changes.first.address", otherCode <> 0 And _
                        InStr(otherReason, keyed.SourceReference("General", CStr(fieldKeys(CLng(pair) - 1)))) > 0 And otherReason <> reason
                End If
            Next pair
            range.Formula = block
            ' Составные условия остаются у рецепта и указывают соответствующее поле.
            If CStr(fixture) <> "Rectangle" Then
                range.Cells(offsetRow, 3).Value2 = TestFixtureInputCell(range, "General", "B2").Value2
                Set definition = TestRectanglesInputOutcome(settings, units, code, source, reason)
                TestCheck "rect.adjacency.address", code <> 0 And InStr(reason, table.SourceReference(offsetRow, 3)) > 0
                result = result & TestInputContractRow(CStr(unit), CStr(fixture), "nonadjacent", code, source, reason, 0#)
            End If
            For Each aliasName In Array("Rectangle", "LSection", "TwoRectangles", "Rectangles", "TwoRectangle", "TwoRectanglesOffset", "OffsetTwoRectangles")
                range.Formula = block: range.Cells(shapeRow, 3).Value2 = CStr(aliasName)
                Set definition = TestRectanglesInputOutcome(settings, units, code, source, reason)
                TestCheck "rect.removed.choice." & CStr(unit) & "." & CStr(aliasName), code <> 0 And InStr(reason, keyed.SourceReference("General", "Shape")) > 0
                result = result & TestInputContractRow(CStr(unit), CStr(fixture), CStr(aliasName), code, source, reason, 0#)
            Next aliasName
            range.Formula = block
        Next fixture
    Next unit
    GoTo Finished
Failed:
    TestCheck "rect.runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    If Not config Is Nothing And Not IsEmpty(saved) Then TestFixtureRestoreInputs config, saved
    TestCloseoutRectanglesInputMatrix = result
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Вызывает обычное чтение описания сечения и возвращает полученную ошибку.
' Ожидаемый номер и текст проверяет вызывающий тест.
Private Function TestRectanglesInputOutcome(ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem, _
        ByRef code As Long, ByRef source As String, ByRef reason As String) As CSectionDefinition
    code = 0: source = "": reason = ""
    On Error GoTo Invalid
    Set TestRectanglesInputOutcome = ReadWorkbookSectionDefinition(ThisWorkbook, settings, units)
    Exit Function
Invalid:
    code = Err.Number: source = Err.Source: reason = Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Записывает единицы, вариант ввода, исходную ошибку и площадь в одну строку контрольного протокола.
Private Function TestInputContractRow(ByVal unit As String, ByVal fixture As String, ByVal name As String, _
        ByVal code As Long, ByVal source As String, ByVal reason As String, ByVal area As Double) As String
    TestInputContractRow = unit & "|" & fixture & "|" & name & "|" & CStr(code) & "|" & source & "|" & reason & "|" & CStr(area) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Исходный квадрат в заданном CCW-порядке; имена не являются автоматическими anchors.

Private Function TestRectangleSpec(ByVal radius As Double) As CPolygonSpec
    Dim polygon As CPolygonSpec: Set polygon = New CPolygonSpec
    polygon.AddVertex "A", 0#, 0#, radius, "Input.A"
    polygon.AddVertex "B", 100#, 0#, radius, "Input.B"
    polygon.AddVertex "C", 100#, 100#, radius, "Input.C"
    polygon.AddVertex "D", 0#, 100#, radius, "Input.D"
    polygon.Seal: Set TestRectangleSpec = polygon
End Function
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сохраняет порядок, координаты, радиусы и адреса всех вершин для проверки неизменности описателя.
Private Function TestSpecSnapshot(ByVal polygon As CPolygonSpec) As String
    Dim i As Long
    For i = 1 To polygon.Count
        TestSpecSnapshot = TestSpecSnapshot & polygon.VertexID(i) & "|" & CStr(polygon.X(i)) & "|" & CStr(polygon.Y(i)) & _
            "|" & CStr(polygon.FilletRadius(i)) & "|" & polygon.SourceReference(i) & vbCrLf
    Next i
End Function
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Учитывает отдельное условие проверки и сохраняет его имя в итоговом протоколе.
Private Sub TestCheck(ByVal name As String, ByVal condition As Boolean)
    If condition Then
        mPassed = mPassed + 1: mReport = mReport & "PASS: " & name & vbCrLf
    Else
        mFailed = mFailed + 1: mReport = mReport & "FAIL: " & name & vbCrLf
    End If
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' контактные последовательности, независимые настройки и сохранение состояния после ошибки.
Public Function RunRebarContactBuilderTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    On Error GoTo Failed
    TestContactSequence
    TestContactDefaults
    TestContactSigns
    TestContactFailures
    GoTo Finished
Failed:
    TestCheck "contactBuilder.runtime; " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    RunRebarContactBuilderTests = mReport & "TOTAL REBAR CONTACT BUILDER: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' пример 32 + 20 + 16 по нормали и 12 вдоль линии; все ряды относятся к одному Main.
Private Sub TestContactSequence()
    Dim contact As CRebarContactBuilder, defaults As CRebarRuleOptions, row As CRebarRuleOptions
    Set defaults = New CRebarRuleOptions
    Set contact = New CRebarContactBuilder
    contact.Initialize 32#, 1#, 1#, defaults
    Set row = contact.AppendRow(20#, "Stacked", "EverySecondBar", "ContactTest[1,2]")
    TestCheck "contactBuilder.example.first", row.NormalOffset = 28.5 And row.TangentOffset = 0#
    TestCheck "contactBuilder.example.binding", row.EverySecond
    TestCheck "contactBuilder.example.source", row.SourceReference = "ContactTest[1,2]"
    Set row = contact.AppendRow(16#, "Stacked", "EachBar", "ContactTest[2,2]")
    TestCheck "contactBuilder.example.second", row.NormalOffset = 49# And row.TangentOffset = 0#
    TestCheck "contactBuilder.example.nominal", Not row.EverySecond
    Set row = contact.AppendRow(12#, "SideBySide", "EachBar", "ContactTest[3,2]")
    TestCheck "contactBuilder.example.third", row.TangentOffset = 24.5 And row.NormalOffset = 0#
    Set row = contact.AppendRow(10#, "Stacked", "EverySecondBar", "ContactTest[4,2]")
    TestCheck "contactBuilder.example.fourth", row.NormalOffset = 64.5 And row.TangentOffset = 0#
    Set row = contact.AppendRow(8#, "SideBySide", "EachBar", "ContactTest[5,2]")
    TestCheck "contactBuilder.example.fifth", row.TangentOffset = 37# And row.NormalOffset = 0#
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' defaults, его снимок и каждый результат независимы; индивидуальные поля можно менять.
Private Sub TestContactDefaults()
    Dim contact As CRebarContactBuilder, defaults As CRebarRuleOptions, first As CRebarRuleOptions, second As CRebarRuleOptions, copy As CRebarRuleOptions
    Set defaults = New CRebarRuleOptions
    defaults.StartSetback = 7#: defaults.EndSetback = 9#: defaults.EndMode = ExcludeBoth
    defaults.SourcePrefix = "Extra.": defaults.SourceSuffix = ".Default": defaults.Comment = "Общие настройки"
    defaults.AnnotationGroup = "ExtraGroup": defaults.Publication = PublishPerSourcePosition: defaults.Coincidence = KeepPublishedOwner
    defaults.NormalOffset = 90#: defaults.TangentOffset = -90#: defaults.EverySecond = True
    defaults.PlacementCheck = CheckFullDisk: defaults.IncomingTangentAtJoint = True
    defaults.DirectionX = 2#: defaults.DirectionY = -3#: defaults.SourceNormalSign = -1#
    defaults.SuccessfulProjectionOrdinal = True: defaults.NewPublicationRoot = True: defaults.SkipMissingProjection = True
    defaults.PositionOrder = FromPathMiddle: defaults.OverlapPolicy = RejectOverlap
    defaults.SourceReference = "DefaultsTest[1,2]"
    Set copy = defaults.Clone
    TestCheck "contactBuilder.clone.StartSetback", copy.StartSetback = defaults.StartSetback
    TestCheck "contactBuilder.clone.EndSetback", copy.EndSetback = defaults.EndSetback
    TestCheck "contactBuilder.clone.EndMode", copy.EndMode = defaults.EndMode
    TestCheck "contactBuilder.clone.DistributionScope", copy.DistributionScope = defaults.DistributionScope
    TestCheck "contactBuilder.clone.RemainderPlacement", copy.RemainderPlacement = defaults.RemainderPlacement
    TestCheck "contactBuilder.clone.TangentOffset", copy.TangentOffset = defaults.TangentOffset
    TestCheck "contactBuilder.clone.NormalOffset", copy.NormalOffset = defaults.NormalOffset
    TestCheck "contactBuilder.clone.EverySecond", copy.EverySecond = defaults.EverySecond
    TestCheck "contactBuilder.clone.IncomingTangentAtJoint", copy.IncomingTangentAtJoint = defaults.IncomingTangentAtJoint
    TestCheck "contactBuilder.clone.DirectionX", copy.DirectionX = defaults.DirectionX
    TestCheck "contactBuilder.clone.DirectionY", copy.DirectionY = defaults.DirectionY
    TestCheck "contactBuilder.clone.SourceNormalSign", copy.SourceNormalSign = defaults.SourceNormalSign
    TestCheck "contactBuilder.clone.SuccessfulProjectionOrdinal", copy.SuccessfulProjectionOrdinal = defaults.SuccessfulProjectionOrdinal
    TestCheck "contactBuilder.clone.NewPublicationRoot", copy.NewPublicationRoot = defaults.NewPublicationRoot
    TestCheck "contactBuilder.clone.SkipMissingProjection", copy.SkipMissingProjection = defaults.SkipMissingProjection
    TestCheck "contactBuilder.clone.Publication", copy.Publication = defaults.Publication
    TestCheck "contactBuilder.clone.PositionOrder", copy.PositionOrder = defaults.PositionOrder
    TestCheck "contactBuilder.clone.Coincidence", copy.Coincidence = defaults.Coincidence
    TestCheck "contactBuilder.clone.PlacementCheck", copy.PlacementCheck = defaults.PlacementCheck
    TestCheck "contactBuilder.clone.OverlapPolicy", copy.OverlapPolicy = defaults.OverlapPolicy
    TestCheck "contactBuilder.clone.SourcePrefix", copy.SourcePrefix = defaults.SourcePrefix
    TestCheck "contactBuilder.clone.SourceSuffix", copy.SourceSuffix = defaults.SourceSuffix
    TestCheck "contactBuilder.clone.AnnotationGroup", copy.AnnotationGroup = defaults.AnnotationGroup
    TestCheck "contactBuilder.clone.Comment", copy.Comment = defaults.Comment
    TestCheck "contactBuilder.clone.SourceReference", copy.SourceReference = defaults.SourceReference
    copy.StartSetback = 111#: copy.AnnotationGroup = "CopyOnly"
    TestCheck "contactBuilder.clone.independent", defaults.StartSetback = 7# And defaults.AnnotationGroup = "ExtraGroup"
    Set contact = New CRebarContactBuilder
    contact.Initialize 32#, 1#, -1#, defaults
    defaults.StartSetback = 999#: defaults.EndMode = IncludeStart: defaults.SourcePrefix = "Changed."
    defaults.AnnotationGroup = "Changed": defaults.PlacementCheck = CheckCenter
    defaults.SourceReference = "ChangedDefaults"
    Set first = contact.AppendRow(20#, "Stacked", "EachBar", "FirstTest[1,2]")
    TestCheck "contactBuilder.snapshot.setbacks", first.StartSetback = 7# And first.EndSetback = 9#
    TestCheck "contactBuilder.snapshot.endMode", first.EndMode = ExcludeBoth
    TestCheck "contactBuilder.snapshot.names", first.SourcePrefix = "Extra." And first.SourceSuffix = ".Default"
    TestCheck "contactBuilder.snapshot.policies", first.AnnotationGroup = "ExtraGroup" And first.PlacementCheck = CheckFullDisk
    TestCheck "contactBuilder.snapshot.source", first.SourceReference = "FirstTest[1,2]"
    TestCheck "contactBuilder.snapshot.overrides", first.NormalOffset = 28.5 And first.TangentOffset = 0# And Not first.EverySecond
    first.StartSetback = 123#: first.AnnotationGroup = "FirstOnly": first.SourceSuffix = "_R2": first.Comment = "Индивидуальный ряд"
    first.NormalOffset = -1000#: first.PlacementCheck = CheckCenter
    Set second = contact.AppendRow(16#, "Stacked", "EverySecondBar", "SecondTest[1,2]")
    TestCheck "contactBuilder.rows.independent", second.StartSetback = 7# And second.AnnotationGroup = "ExtraGroup"
    TestCheck "contactBuilder.rows.names", second.SourceSuffix = ".Default" And second.Comment = "Общие настройки"
    TestCheck "contactBuilder.rows.nominal", second.NormalOffset = 49# And second.TangentOffset = 0# And second.EverySecond
    TestCheck "contactBuilder.rows.placement", second.PlacementCheck = CheckFullDisk
    TestCheck "contactBuilder.rows.firstRetained", first.SourceSuffix = "_R2" And first.Comment = "Индивидуальный ряд" And first.NormalOffset = -1000#
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' все сочетания знаков, перемежающиеся направления и EverySecondBar.
Private Sub TestContactSigns()
    Dim contact As CRebarContactBuilder, defaults As CRebarRuleOptions, row As CRebarRuleOptions, normal As Variant, tangent As Variant
    Set defaults = New CRebarRuleOptions
    For Each normal In Array(-1#, 1#)
        For Each tangent In Array(-1#, 1#)
            Set contact = New CRebarContactBuilder
            contact.Initialize 32#, CDbl(normal), CDbl(tangent), defaults
            Set row = contact.AppendRow(20#, "Stacked", "EachBar", "Signs[1,2]")
            TestCheck "contactBuilder.sign.normal.first", row.NormalOffset = CDbl(normal) * 28.5 And row.TangentOffset = 0#
            Set row = contact.AppendRow(12#, "SideBySide", "EverySecondBar", "Signs[2,2]")
            TestCheck "contactBuilder.sign.tangent.first", row.TangentOffset = CDbl(tangent) * 24.5 And row.NormalOffset = 0#
            Set row = contact.AppendRow(16#, "Stacked", "EverySecondBar", "Signs[3,2]")
            TestCheck "contactBuilder.sign.normal.next", row.NormalOffset = CDbl(normal) * 49# And row.EverySecond
            Set row = contact.AppendRow(8#, "SideBySide", "EachBar", "Signs[4,2]")
            TestCheck "contactBuilder.sign.tangent.next", row.TangentOffset = CDbl(tangent) * 37# And Not row.EverySecond
        Next tangent
    Next normal
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' ошибки не расходуют диаметр и не меняют независимый накопитель второго направления.
Private Sub TestContactFailures()
    Dim contact As CRebarContactBuilder, defaults As CRebarRuleOptions, row As CRebarRuleOptions, scenario As Long
    Set defaults = New CRebarRuleOptions: defaults.SourceReference = "InitializeTest[1,2]"
    Set contact = New CRebarContactBuilder
    TestCheck "contactBuilder.uninitialized", TestContactAppendRejected(contact, 20#, "Stacked", "EachBar")
    For scenario = 0 To 4
        Set contact = New CRebarContactBuilder
        TestCheck "contactBuilder.initialize.reject." & CStr(scenario), TestContactInitializeRejected(contact, defaults, scenario)
        contact.Initialize 32#, 1#, -1#, defaults
        Set row = contact.AppendRow(20#, "Stacked", "EachBar", "AfterInitializeTest[1,2]")
        TestCheck "contactBuilder.initialize.recover." & CStr(scenario), row.NormalOffset = 28.5
    Next scenario
    For scenario = 0 To 6
        Set contact = New CRebarContactBuilder
        contact.Initialize 32#, 1#, -1#, defaults
        Set row = contact.AppendRow(20#, "Stacked", "EachBar", "BeforeFailure[1,2]")
        Select Case scenario
            Case 0: TestCheck "contactBuilder.reject.location", TestContactAppendRejected(contact, 100#, "Unknown", "EachBar")
            Case 1: TestCheck "contactBuilder.reject.emptyLocation", TestContactAppendRejected(contact, 100#, "", "EachBar")
            Case 2: TestCheck "contactBuilder.reject.binding", TestContactAppendRejected(contact, 100#, "Stacked", "Unknown")
            Case 3: TestCheck "contactBuilder.reject.emptyBinding", TestContactAppendRejected(contact, 100#, "SideBySide", "")
            Case 4: TestCheck "contactBuilder.reject.zero", TestContactAppendRejected(contact, 0#, "Stacked", "EachBar")
            Case 5: TestCheck "contactBuilder.reject.negative", TestContactAppendRejected(contact, -10#, "SideBySide", "EachBar")
            Case 6: TestCheck "contactBuilder.reject.reinitialize", TestContactInitializeRejected(contact, defaults, 5)
        End Select
        Set row = contact.AppendRow(16#, "Stacked", "EverySecondBar", "AfterFailure[1,2]")
        TestCheck "contactBuilder.failure.normal." & CStr(scenario), row.NormalOffset = 49#
        Set row = contact.AppendRow(12#, "SideBySide", "EachBar", "AfterFailure[2,2]")
        TestCheck "contactBuilder.failure.tangent." & CStr(scenario), row.TangentOffset = -24.5
    Next scenario
    For scenario = 0 To 1
        Set contact = New CRebarContactBuilder
        contact.Initialize 1#, 1#, -1#, defaults
        Dim huge As Double
        huge = 1E+308: If scenario = 1 Then huge = 1.4E+308
        Set row = contact.AppendRow(huge, "Stacked", "EachBar", "HugeTest[1,2]")
        TestCheck "contactBuilder.overflow.reject." & CStr(scenario), TestContactAppendRejected(contact, huge, "Stacked", "EachBar")
        Set row = contact.AppendRow(1#, "Stacked", "EachBar", "HugeRecovery[1,2]")
        TestCheck "contactBuilder.overflow.recover." & CStr(scenario), row.NormalOffset = huge
    Next scenario
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' ожидает адресную ошибку AppendRow без регистрации правила и без размещения стержней.
Private Function TestContactAppendRejected(ByVal contact As CRebarContactBuilder, ByVal diameter As Double, ByVal location As String, ByVal binding As String) As Boolean
    Dim row As CRebarRuleOptions
    On Error GoTo Rejected
    Set row = contact.AppendRow(diameter, location, binding, "FailureTest[1,2]")
    Exit Function
Rejected:
    TestContactAppendRejected = Err.Number <> 0 And InStr(1, Err.Description, "FailureTest[1,2]", vbBinaryCompare) > 0
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' неверная и повторная инициализация не должны частично менять основной ряд.
Private Function TestContactInitializeRejected(ByVal contact As CRebarContactBuilder, ByVal defaults As CRebarRuleOptions, ByVal scenario As Long) As Boolean
    On Error GoTo Rejected
    Select Case scenario
        Case 0: contact.Initialize 0#, 1#, 1#, defaults
        Case 1: contact.Initialize -10#, 1#, 1#, defaults
        Case 2: contact.Initialize 32#, 0#, 1#, defaults
        Case 3: contact.Initialize 32#, 1#, 2#, defaults
        Case 4: contact.Initialize 32#, 1#, 1#, Nothing
        Case 5: contact.Initialize 64#, -1#, 1#, defaults
    End Select
    Exit Function
Rejected:
    TestContactInitializeRejected = Err.Number <> 0
    If scenario <> 4 Then TestContactInitializeRejected = TestContactInitializeRejected And InStr(1, Err.Description, defaults.SourceReference, vbBinaryCompare) > 0
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' реальные координаты, фильтр исходных позиций и запрет новых коллизий общим исполнителем.
Public Function RunIndependentRowOffsetGeometryTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    On Error GoTo Failed
    Dim kind As Long, b2 As Variant, b3 As Variant, definition As CSectionDefinition, layout As CRebarLayout
    Dim executor As CRebarRuleExecutor, i As Long, secondCount As Long, thirdCount As Long, radius As Double
    For kind = 0 To 1
        For Each b2 In Array("EachBar", "EverySecondBar")
            For Each b3 In Array("EachBar", "EverySecondBar")
                Set definition = TestOffsetGeometryDefinition(kind, 100#, 100#, CStr(b2), CStr(b3))
                Set executor = New CRebarRuleExecutor
                Set layout = executor.Build(definition, TestFixtureGeometry(definition.ConcreteDomain))
                secondCount = 0: thirdCount = 0
                For i = 1 To executor.PositionCount
                    If executor.PositionRuleID(i) = "Row2" Then
                        secondCount = secondCount + 1
                        If kind = 0 Then radius = layout.Y(executor.PositionPhysicalOwner(i)) Else radius = Sqr(layout.X(executor.PositionPhysicalOwner(i)) ^ 2 + layout.Y(executor.PositionPhysicalOwner(i)) ^ 2)
                        TestCheck "offset.geometry.row2." & CStr(kind) & "." & CStr(i), Abs(radius - IIf(kind = 0, 100#, 200#)) < 0.000001
                    ElseIf executor.PositionRuleID(i) = "Row3" Then
                        thirdCount = thirdCount + 1
                        If kind = 0 Then radius = layout.Y(executor.PositionPhysicalOwner(i)) Else radius = Sqr(layout.X(executor.PositionPhysicalOwner(i)) ^ 2 + layout.Y(executor.PositionPhysicalOwner(i)) ^ 2)
                        TestCheck "offset.geometry.row3." & CStr(kind) & "." & CStr(i), Abs(radius - IIf(kind = 0, 120.5, 179.5)) < 0.000001
                    End If
                Next i
                TestCheck "offset.geometry.count2", secondCount = IIf(CStr(b2) = "EachBar", 5, 3)
                TestCheck "offset.geometry.count3", thirdCount = IIf(CStr(b3) = "EachBar", 5, 3)
            Next b3
        Next b2
    Next kind
    For kind = 0 To 1
        Set definition = TestOffsetGeometryDefinition(kind, 0#, 0#)
        Set executor = New CRebarRuleExecutor
        Set layout = executor.Build(definition, TestFixtureGeometry(definition.ConcreteDomain))
        For i = 1 To executor.PositionCount
            If executor.PositionRuleID(i) = "Row2" Or executor.PositionRuleID(i) = "Row3" Then
                If kind = 0 Then radius = layout.Y(executor.PositionPhysicalOwner(i)) Else radius = 300# - Sqr(layout.X(executor.PositionPhysicalOwner(i)) ^ 2 + layout.Y(executor.PositionPhysicalOwner(i)) ^ 2)
                If executor.PositionRuleID(i) = "Row2" Then
                    TestCheck "offset.profile.zero.row2", Abs(radius - 28.5) < 0.000001 And layout.Diameter(executor.PositionPhysicalOwner(i)) = 20#
                Else
                    TestCheck "offset.profile.zero.row3", Abs(radius - 49#) < 0.000001 And layout.Diameter(executor.PositionPhysicalOwner(i)) = 16#
                End If
            End If
        Next i
    Next kind
    TestCheck "offset.overlap.rejected", TestOffsetGeometryRejected(TestOffsetGeometryDefinition(0, 100#, 105#), "Поперечные сечения стержней")
    Set layout = TestFixtureLayout(TestOffsetGeometryDefinition(0, 100#, 118#))
    TestCheck "offset.touch.allowed", layout.Count = 15
    TestCheck "offset.disk.outside", TestOffsetGeometryRejected(TestOffsetGeometryDefinition(0, 1499#, 100#), "целиком")
    TestCheck "offset.arc.center", TestOffsetGeometryRejected(TestOffsetGeometryDefinition(1, 300#, 0#), "контур исчез целиком") ' У полной окружности после исчезновения радиуса не остаётся сторон для замыкания.
    TestCheck "offset.arc.inversion", TestOffsetGeometryRejected(TestOffsetGeometryDefinition(1, 301#, 0#), "контур исчез целиком")
    TestCheck "offset.arc.contact.collapse", TestOffsetGeometryRejected(TestOffsetGeometryDefinition(1, 290#, 290#), "контур исчез целиком")
    For kind = 0 To 1
        TestCheck "offset.coincidence.order." & CStr(kind), TestOffsetGeometryRejected(TestOffsetGeometryDefinition(0, 100#, 200#, "EachBar", "EachBar", kind + 1), "совпал")
        TestCheck "offset.shared.upgrade.order." & CStr(kind), TestOffsetGeometryRejected(TestOffsetSharedJointDefinition(kind, True), "Поперечные сечения стержней")
        Set layout = TestFixtureLayout(TestOffsetSharedJointDefinition(kind, False))
        TestCheck "offset.shared.joint.allowed." & CStr(kind), layout.Count = 2
    Next kind
    GoTo Finished
Failed:
    TestCheck "offset.geometry.runtime; " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    RunIndependentRowOffsetGeometryTests = mReport & "TOTAL INDEPENDENT ROW OFFSET GEOMETRY: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' пять исходных позиций и два связанных с Main ряда; формулы координат выполняет рабочий код.
Private Function TestOffsetGeometryDefinition(ByVal kind As Long, ByVal a2 As Double, ByVal a3 As Double, _
        Optional ByVal b2 As String = "EachBar", Optional ByVal b3 As String = "EachBar", Optional ByVal intruder As Long = 0) As CSectionDefinition
    Dim facade As CSectionDefinitionBuilder, path As CGeometryPath, options As CRebarRuleOptions
    Dim contact As CRebarContactBuilder, sign As Double
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "OffsetGeometryTest"
    If kind = 0 Then Set path = facade.RectanglePath("Boundary", -500#, -500#, 2000#, 2000#) Else Set path = facade.CirclePath("Boundary", 0#, 0#, 500#)
    facade.Outer "OUTER", path.ID
    If intruder = 1 Then TestAddOffsetIntruder facade, 100#
    If kind = 0 Then Set path = facade.LinePath("Axis", 0#, 0#, 400#, 0#) Else Set path = facade.CirclePath("Axis", 0#, 0#, 300#)
    If kind = 1 Then path.ValidateNormalTransfer 0#, 10#
    Set options = New CRebarRuleOptions: options.Coincidence = KeepPublishedOwner
    If kind = 1 Then options.EndMode = IncludeStart
    facade.RebarByCount "Main", path.ID, 5, 32#, options
    facade.RebarContour "Main", "OUTER", "Boundary", IIf(kind = 0, 500#, 200#), "Row2", "Row3"
    Set options = New CRebarRuleOptions: options.SourceReference = "OffsetTest[2,3]": options.Coincidence = KeepPublishedOwner
    Set contact = New CRebarContactBuilder: contact.Initialize 32#, 1#, 1#, options
    Set options = contact.AppendRow(20#, "Stacked", b2, "OffsetTest[2,3]", a2)
    facade.RebarDependent "Row2", "Main", 20#, options
    Set options = contact.AppendRow(16#, "Stacked", b3, "OffsetTest[3,3]", a3)
    facade.RebarDependent "Row3", "Main", 16#, options
    If intruder = 2 Then TestAddOffsetIntruder facade, 100#
    Set TestOffsetGeometryDefinition = facade.Finish
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' независимый стержень на месте нового ряда; его нельзя скрыть слиянием.
Private Sub TestAddOffsetIntruder(ByVal facade As CSectionDefinitionBuilder, ByVal y As Double)
    Dim path As CGeometryPath, options As CRebarRuleOptions
    Set path = facade.LinePath("IntruderPath", 0#, y, 400#, y)
    Set options = New CRebarRuleOptions: options.Coincidence = KeepPublishedOwner
    facade.RebarByCount "Intruder", path.ID, 5, 20#, options
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' два Main делят конец соседних линий; новое участие должно усилить проверки общего стержня.
Private Function TestOffsetSharedJointDefinition(ByVal reversed As Long, ByVal neighbor As Boolean) As CSectionDefinition
    Dim facade As CSectionDefinitionBuilder, path As CGeometryPath, options As CRebarRuleOptions
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "SharedOffsetJoint"
    Set path = facade.RectanglePath("Boundary", -500#, -500#, 1000#, 1000#, "Bottom", "Right", "Top", "Left"): facade.Outer "OUTER", path.ID
    If neighbor And reversed = 1 Then TestAddOffsetJointNeighbor facade
    Set path = facade.LinePath("FirstPath", 0#, 0#, 400#, 0#)
    Set options = New CRebarRuleOptions: options.Coincidence = KeepPublishedOwner
    facade.RebarAtStations "FirstMain", path.ID, Array(0#), 32#, options
    facade.RebarContour "FirstMain", "OUTER", "Boundary", 500#, "LegacyJoint", "", Array("Bottom")
    options.NormalOffset = 100#: options.SourceReference = "LegacyJoint[2,3]"
    facade.RebarDependent "LegacyJoint", "FirstMain", 20#, options
    Set path = facade.LinePath("SecondPath", -400#, 0#, 0#, 0#)
    Set options = New CRebarRuleOptions: options.Coincidence = KeepPublishedOwner
    facade.RebarAtStations "SecondMain", path.ID, Array(400#), 32#, options
    facade.RebarContour "SecondMain", "OUTER", "Boundary", 500#, "OffsetJoint", "", Array("Bottom")
    options.NormalOffset = 100#: options.ValidateRowOffset = True: options.SourceReference = "OffsetJoint[2,3]"
    facade.RebarDependent "OffsetJoint", "SecondMain", 20#, options
    If neighbor And reversed = 0 Then TestAddOffsetJointNeighbor facade
    Set TestOffsetSharedJointDefinition = facade.Finish
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' сосед пересекает общий дополнительный стержень, но не основные ряды.
Private Sub TestAddOffsetJointNeighbor(ByVal facade As CSectionDefinitionBuilder)
    Dim path As CGeometryPath, options As CRebarRuleOptions
    Set path = facade.LinePath("NeighborPath", 5#, 100#, 100#, 100#)
    Set options = New CRebarRuleOptions: options.Coincidence = KeepPublishedOwner
    facade.RebarAtStations "Neighbor", path.ID, Array(0#), 8#, options
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' ожидает понятную адресную ошибку общего исполнителя.
Private Function TestOffsetGeometryRejected(ByVal definition As CSectionDefinition, ByVal expected As String) As Boolean
    Dim layout As CRebarLayout
    On Error GoTo Rejected
    Set layout = TestFixtureLayout(definition)
    Exit Function
Rejected:
    TestOffsetGeometryRejected = InStr(1, Err.Description, expected, vbTextCompare) > 0 And _
        (InStr(1, Err.Description, "OffsetTest[", vbBinaryCompare) > 0 Or InStr(1, Err.Description, "OffsetJoint[", vbBinaryCompare) > 0)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' настоящий ввод четырёх рецептов, единицы, пустые и формульные отступы.
' Каждая попытка строит арматуру общим исполнителем; исходный Config восстанавливается.
Public Function RunIndependentRowOffsetWorkbookTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    Dim config As Object, saved As Variant, calculation As Long, fixture As Variant, item As Variant
    Dim unitName As Variant, offsets As Variant, settings As CSystemSettingsReader, units As CUnitSystem
    Dim definition As CSectionDefinition, bars As CRebarLayout, expected As Variant, actual As Variant
    Dim target As Object, extra As Long, main As Long, cell As Object, bad As Variant, i As Long, inputTable As CSectionInputTable
    Dim description As String, number As Long, prefix As String, rowCount As Long
    On Error GoTo Failed
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = TestFixtureSaveInputs(config)
    calculation = Application.Calculation: Application.Calculation = -4135
    offsets = Array(Array(0#, 0#), Array(100#, 0#), Array(0#, 200#), Array(100#, 200#), Array(100#, 100#))
    For Each fixture In Array("Circle", "Rectangle", "RoundedMixed", "HollowOffset")
        For Each item In offsets
            expected = Empty
            For Each unitName In Array("mm", "cm", "m")
                TestConfigureOffsetWorkbook CStr(fixture), CStr(unitName), CDbl(item(0)), CDbl(item(1))
                Set definition = TestReadOffsetWorkbookDefinition()
                Set bars = TestFixtureLayout(definition): actual = TestOffsetBarsSnapshot(bars)
                prefix = "offset.workbook." & CStr(fixture) & "." & CStr(item(0)) & "." & CStr(item(1)) & "." & CStr(unitName)
                TestCheck prefix & ".three.rows", bars.Count >= 15
                If IsEmpty(expected) Then expected = actual Else TestCheck prefix & ".units", TestOffsetBarsEqual(expected, actual)
            Next unitName
        Next item
        TestConfigureOffsetWorkbook CStr(fixture), "mm", 100#, 118#
        Set target = TestOffsetWorkbookTable(CStr(fixture))
        TestOffsetWorkbookRows target, CStr(fixture), main, extra
        Set cell = TestOffsetRebarCell(extra, 4)
        For Each bad In Array(-1#, "abc", True, CVErr(2015))
            cell.Value2 = bad: number = 0: description = ""
            On Error Resume Next
            Set definition = TestReadOffsetWorkbookDefinition(): number = Err.Number: description = Err.Description
            Err.Clear: On Error GoTo Failed
            TestCheck "offset.input.reject." & CStr(fixture) & "." & CStr(VarType(bad)), number <> 0
            TestCheck "offset.input.address." & CStr(fixture) & "." & CStr(VarType(bad)), InStr(description, cell.Address(False, False)) > 0 And InStr(description, cell.Worksheet.Name & "!") > 0
        Next bad
        cell.ClearContents: Set definition = TestReadOffsetWorkbookDefinition(): Set bars = TestFixtureLayout(definition)
        actual = TestOffsetBarsSnapshot(bars)
        cell.Value2 = 0#: Set definition = TestReadOffsetWorkbookDefinition(): Set bars = TestFixtureLayout(definition)
        TestCheck "offset.input.blank.zero." & CStr(fixture), TestOffsetBarsEqual(actual, TestOffsetBarsSnapshot(bars))
        TestOffsetRebarCell(extra, 3).Value2 = 0#: cell.Value2 = "inactive"
        Set definition = TestReadOffsetWorkbookDefinition(): Set bars = TestFixtureLayout(definition)
        TestCheck "offset.input.disabled.second." & CStr(fixture), bars.Count >= 10
        TestOffsetRebarCell(extra, 7).Value2 = 0#: TestOffsetRebarCell(extra, 8).Value2 = CVErr(2015)
        Set definition = TestReadOffsetWorkbookDefinition(): Set bars = TestFixtureLayout(definition)
        TestCheck "offset.input.disabled.extras." & CStr(fixture), bars.Count >= 5
    Next fixture
    TestConfigureOffsetWorkbook "Circle", "mm", 100#, 200#
    Set units = New CUnitSystem: units.InitializeDefaults
    Set target = TestOffsetWorkbookTable("Circle")
    Set inputTable = New CSectionInputTable: inputTable.Initialize "rngCircleGeometry", target.Value2
    TestCheck "offset.input.direct.reader", inputTable.OptionalRowOffset(TestFixtureInput(target.Value2).RowIndex("Extra", "Circle"), TestFixtureInput(target.Value2).ColumnIndex("Extra", "a2"), units) = 100#
    Set cell = TestFixtureCircleFieldCell(ThisWorkbook, "Rebar.AxisOffset2")
    cell.Formula = "=" & TestFixtureCircleFieldCell(ThisWorkbook, "Circle.Diameter").Address() & "/10"
    cell.Calculate ' Явно готовим кэш 100 при D=1000; далее низкоуровневое чтение не должно его пересчитывать.
    TestCheck "offset.input.formula.manual.seeded", CDbl(cell.Value2) = 100#
    TestFixtureCircleFieldCell(ThisWorkbook, "Circle.Diameter").Value2 = 1200#
    TestCheck "offset.input.formula.manual.stale", CDbl(cell.Value2) = 100#
    Set definition = TestReadOffsetWorkbookDefinition(): Set bars = TestFixtureLayout(definition)
    actual = TestOffsetBarsSnapshot(bars)
    TestCheck "offset.input.formula.manual.cacheUnchanged", CDbl(cell.Value2) = 100#
    TestCheck "offset.input.formula.manual.modeUnchanged", Application.Calculation = xlCalculationManual
    mReport = mReport & "OFFSET_MANUAL_FORMULA: stage=cached; a2=" & CStr(cell.Value2) & vbCrLf
    Application.Calculate ' Один явный пересчёт графа соответствует границе штатного формирования; Capture читает готовый кэш.
    TestCheck "offset.input.formula.explicit.value", CDbl(cell.Value2) = 120#
    Set definition = TestReadOffsetWorkbookDefinition(): Set bars = TestFixtureLayout(definition)
    expected = TestOffsetBarsSnapshot(bars)
    TestCheck "offset.input.formula.explicit.changedLayout", Not TestOffsetBarsEqual(actual, expected)
    mReport = mReport & "OFFSET_MANUAL_FORMULA: stage=explicitCalculate; a2=" & CStr(cell.Value2) & vbCrLf
    cell.Value2 = 100#: Set definition = TestReadOffsetWorkbookDefinition(): Set bars = TestFixtureLayout(definition)
    TestCheck "offset.input.formula.manual.cachedLayout", TestOffsetBarsEqual(actual, TestOffsetBarsSnapshot(bars)) ' Проверяем полный состав, порядок, координаты и диаметры при исходном a2=100.
    cell.Value2 = 120#: Set definition = TestReadOffsetWorkbookDefinition(): Set bars = TestFixtureLayout(definition)
    TestCheck "offset.input.formula.manual.recalculated", TestOffsetBarsEqual(expected, TestOffsetBarsSnapshot(bars))
    TestCheck "offset.input.formula.explicit.modeUnchanged", Application.Calculation = xlCalculationManual
    GoTo Finished
Failed:
    TestCheck "offset.workbook.runtime; " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    If Not config Is Nothing And Not IsEmpty(saved) Then TestFixtureRestoreInputs config, saved
    Application.Calculation = calculation
    RunIndependentRowOffsetWorkbookTests = mReport & "TOTAL INDEPENDENT ROW OFFSET WORKBOOK: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' задаёт достаточно просторные сечения; отступы и диаметры даны в мм.
' Открытый ряд берём по верхней стороне; в пустотелом сечении добавляем ряд у отверстия.
Private Sub TestConfigureOffsetWorkbook(ByVal fixture As String, ByVal unitName As String, ByVal a2 As Double, ByVal a3 As Double)
    Dim unitTable As Object, target As Object, factor As Double, row As Long, key As Variant
    Dim keyed As CSectionInputReader, active As Boolean
    Set unitTable = ThisWorkbook.Names("rngUnitSettings").RefersToRange
    For row = 1 To unitTable.Rows.Count
        If CStr(unitTable.Cells(row, 1).Value2) = "Length" Then unitTable.Cells(row, 2).Value2 = unitName
    Next row
    factor = 1#: If unitName = "cm" Then factor = 10#
    If unitName = "m" Then factor = 1000#
    modTestShapeArchitecture.TestConfigureShapeFixture fixture
    Set target = TestOffsetWorkbookTable(fixture)
    Select Case fixture
        Case "Circle"
            TestFixtureInputCell(target, "General", "D").Value2 = 1000# / factor
            TestFixtureInputCell(target, "Main", "Circle", "as").Value2 = 60# / factor
            TestFixtureInputCell(target, "Main", "Circle", "n").Value2 = 12
            TestFixtureInputCell(target, "Main", "Circle", "d").Value2 = 32# / factor
        Case "Rectangle"
            TestFixtureInputCell(target, "General", "H1").Value2 = 1200# / factor: TestFixtureInputCell(target, "General", "B1").Value2 = 1200# / factor
        Case "RoundedMixed"
            TestFixtureInputCell(target, "General", "B").Value2 = 1600# / factor: TestFixtureInputCell(target, "General", "H").Value2 = 1600# / factor
            For Each key In Array("left", "right")
                TestFixtureInputCell(target, "Detail", "W", CStr(key)).Value2 = 400# / factor
                TestFixtureInputCell(target, "Detail", "R1", CStr(key)).Value2 = 300# / factor
                TestFixtureInputCell(target, "Detail", "R2", CStr(key)).Value2 = 300# / factor
            Next key
        Case Else
            TestFixtureInputCell(target, "General", "H").Value2 = 2400# / factor: TestFixtureInputCell(target, "General", "B").Value2 = 2400# / factor
            TestFixtureInputCell(target, "Detail", "Ho").Value2 = 800# / factor: TestFixtureInputCell(target, "Detail", "Bo").Value2 = 800# / factor
    End Select
    Select Case fixture
        Case "Circle": Set keyed = TestFixtureInput(TestFixtureRecipeValues("rngCircleGeometry"))
        Case "Rectangle": Set keyed = TestFixtureInput(TestFixtureRecipeValues("rngRectanglesGeometry"))
        Case "RoundedMixed": Set keyed = TestFixtureInput(TestFixtureRecipeValues("rngPointedSectionGeometry"))
        Case Else: Set keyed = TestFixtureInput(TestFixtureRecipeValues("rngHollowRectangleGeometry"))
    End Select
    For Each key In keyed.RowKeys("Main")
        active = (CStr(key) = "Circle" Or CStr(key) = "Top" Or CStr(key) = "R1.Top" Or CStr(key) = "OpeningTop")
        If fixture <> "Circle" Then
            TestFixtureInputCell(target, "Main", CStr(key), "as").Value2 = 60# / factor
            TestFixtureInputCell(target, "Main", CStr(key), "d").Value2 = IIf(active, 32# / factor, 0#)
            If Left$(CStr(key), 7) <> "Opening" Then TestFixtureInputCell(target, "Main", CStr(key), "n").Value2 = IIf(active, 5, 0)
            If fixture = "Rectangle" Then
                TestFixtureInputCell(target, "Main", CStr(key), "s0").Value2 = 150# / factor
                TestFixtureInputCell(target, "Main", CStr(key), "s1").Value2 = 150# / factor
            End If
        End If
        TestFixtureInputCell(target, "Extra", CStr(key), "d2").Value2 = IIf(active, 20# / factor, 0#)
        TestFixtureInputCell(target, "Extra", CStr(key), "d3").Value2 = IIf(active, 16# / factor, 0#)
        TestFixtureInputCell(target, "Extra", CStr(key), "a2").Value2 = a2 / factor
        TestFixtureInputCell(target, "Extra", CStr(key), "a3").Value2 = a3 / factor
        TestFixtureInputCell(target, "Extra", CStr(key), "bind2").Value2 = "EachBar"
        TestFixtureInputCell(target, "Extra", CStr(key), "bind3").Value2 = "EachBar"
        TestFixtureInputCell(target, "Extra", CStr(key), "pos2").Value2 = "Stacked"
        TestFixtureInputCell(target, "Extra", CStr(key), "pos3").Value2 = "Stacked"
    Next key
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' получает действующую таблицу выбранного рецепта по штатному имени.
Private Function TestOffsetWorkbookTable(ByVal fixture As String) As Object
    Dim name As String
    Select Case fixture
        Case "Circle": name = "rngCircleGeometry"
        Case "Rectangle": name = "rngRectanglesGeometry"
        Case "RoundedMixed": name = "rngPointedSectionGeometry"
        Case Else: name = "rngHollowRectangleGeometry"
    End Select
    Set TestOffsetWorkbookTable = ThisWorkbook.Names(name).RefersToRange
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' строки основного и дополнительного армирования тестируемой верхней стороны.
Private Sub TestOffsetWorkbookRows(ByVal target As Object, ByVal fixture As String, ByRef main As Long, ByRef extra As Long)
    Dim keyed As CSectionInputReader, key As String, typeID As String
    Select Case fixture
        Case "Circle": typeID = "Circle": key = "Circle"
        Case "Rectangle": typeID = "Rectangles": key = "R1.Top"
        Case "RoundedMixed": typeID = "PointedSection": key = "Top"
        Case Else: typeID = "HollowRectangle": key = "Top"
    End Select
    TestFixtureSelectRecipeType ThisWorkbook, typeID
    Set keyed = TestFixtureInput(target.Value2)
    main = keyed.RowIndex("Main", key): extra = keyed.RowIndex("Extra", key)
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' чтение тем же снимком входа, которым пользуется рабочий сборщик.
Private Function TestReadOffsetWorkbookDefinition() As CSectionDefinition
    Dim settings As CSystemSettingsReader, units As CUnitSystem
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    Set TestReadOffsetWorkbookDefinition = TestFixtureWorkbookDefinition(settings.GetString("Geometry.Type"), ThisWorkbook, units)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' независимая численная запись координат и диаметров в порядке публикации.
Private Function TestOffsetBarsSnapshot(ByVal bars As CRebarLayout) As Variant
    Dim values() As Double, i As Long
    ReDim values(1 To bars.Count, 1 To 3)
    For i = 1 To bars.Count
        values(i, 1) = bars.X(i): values(i, 2) = bars.Y(i): values(i, 3) = bars.Diameter(i)
    Next i
    TestOffsetBarsSnapshot = values
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' допускает только погрешность пересчёта единиц, без округления входа.
Private Function TestOffsetBarsEqual(ByVal first As Variant, ByVal second As Variant) As Boolean
    Dim i As Long, j As Long
    If UBound(first, 1) <> UBound(second, 1) Then Exit Function
    For i = 1 To UBound(first, 1)
        For j = 1 To 3
            If Abs(first(i, j) - second(i, j)) > 0.00000001 Then Exit Function
        Next j
    Next i
    TestOffsetBarsEqual = True
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' настоящие схемы Results и расчётные снимки новых рядов, без изменения рабочей книги.
Public Function TestDumpIndependentRowOffsetExamples(ByVal outputDirectory As String) As String
    Dim config As Object, saved As Variant, result As String, fixtures As Variant, fixture As String
    Dim titles As Variant, a2 As Double, a3 As Double, index As Long, target As Object, main As Long, extra As Long
    Dim settings As CSystemSettingsReader, units As CUnitSystem, builder As CSectionBuilder, model As CSectionModel
    Dim calculation As Long, events As Boolean, updating As Boolean, plot As String
    On Error GoTo Failed
    calculation = Application.Calculation: events = Application.EnableEvents: updating = Application.ScreenUpdating
    Application.EnableEvents = False: Application.ScreenUpdating = True
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = TestFixtureSaveInputs(config)
    fixtures = Array("Circle", "Circle", "Rectangle", "RoundedMixed", "HollowOffset", "Rectangle", "Circle")
    titles = Array("Circle_TwoRows_100", "Circle_ThreeRows_100_100", "Rectangle_ThreeRows_100_200", _
        "Rounded_Sloped_60_100", "Hollow_Opening_100_100_SideBySide", "Rectangle_100_100_EverySecond", "Circle_Arc_ThreeRows_100_200")
    For index = 0 To UBound(fixtures)
        fixture = CStr(fixtures(index)): a2 = 100#: a3 = 100#
        If index = 2 Or index = 6 Then a3 = 200#
        If index = 3 Then a2 = 60#: a3 = 100#
        TestConfigureOffsetWorkbook fixture, "mm", a2, a3
        Set target = TestOffsetWorkbookTable(fixture): TestOffsetWorkbookRows target, fixture, main, extra
        If index = 0 Then TestOffsetRebarCell(extra, 7).Value2 = 0#
        If index = 3 Then
            TestOffsetRebarCell(main, 4).Value2 = 0#: TestOffsetRebarCell(main, 5).Value2 = 0
            TestOffsetRebarCell(extra, 3).Value2 = 0#: TestOffsetRebarCell(extra, 7).Value2 = 0#
            TestFixtureInputCell(target, "Main", "Left", "d").Value2 = 32#: TestFixtureInputCell(target, "Main", "Left", "n").Value2 = 5
            TestFixtureInputCell(target, "Extra", "Left", "d2").Value2 = 20#: TestFixtureInputCell(target, "Extra", "Left", "d3").Value2 = 16#
        End If
        If index = 6 Then
            TestFixtureCircleFieldCell(ThisWorkbook, "Rebar.StartSetback").Value2 = 150#
            TestFixtureCircleFieldCell(ThisWorkbook, "Rebar.EndSetback").Value2 = 250#
            TestFixtureInputCell(ThisWorkbook.Names("rngCircleGeometry").RefersToRange, "Main", "Circle", "endMode").Value2 = "IncludeBoth"
        End If
        If index = 4 Or index = 5 Then
            TestOffsetRebarCell(extra, 10).Value2 = "SideBySide": TestOffsetRebarCell(extra, 9).Value2 = "EverySecondBar"
            If index = 4 Then
                TestOffsetRebarCell(extra + 4, 10).Value2 = "SideBySide": TestOffsetRebarCell(extra + 4, 9).Value2 = "EverySecondBar"
            End If
        End If
        Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
        Set units = New CUnitSystem: units.LoadFromSettings settings
        Set builder = New CSectionBuilder: Set model = BuildWorkbookSectionModel(ThisWorkbook, settings, units, builder)
        result = result & "EXAMPLE: " & CStr(titles(index)) & "; bars=" & CStr(builder.Rebars.Count) & vbCrLf
        result = result & modTestShapeArchitecture.TestPreparedShapeSnapshot(CStr(titles(index)), model, builder.Mesh, builder.Rebars)
        If index = 1 Or index = 2 Or index = 4 Or index = 6 Then result = result & modTestShapeArchitecture.TestPreparedCalculationSnapshot(CStr(titles(index)), model)
        plot = modTestShapeArchitecture.TestPreparedPreviewSnapshot(CStr(titles(index)), model, outputDirectory)
        result = result & plot
        If index = 4 Then TestExportOffsetPreviewLarge outputDirectory & Application.PathSeparator & "Hollow_Detail_4x.png"
    Next index
    GoTo Finished
Failed:
    result = result & "FAIL: offset.examples; " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Finished:
    If Not config Is Nothing And Not IsEmpty(saved) Then TestFixtureRestoreInputs config, saved
    Application.Calculation = calculation: Application.EnableEvents = events: Application.ScreenUpdating = updating
    TestDumpIndependentRowOffsetExamples = result
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Excel повторно выводит векторный снимок готовой диаграммы в крупный PNG.
' Координаты стержней и фигуры исходной диаграммы не меняются.
Private Sub TestExportOffsetPreviewLarge(ByVal path As String)
    Dim container As Object, scratch As Object, canvas As Object, picture As Object
    Dim width As Double, height As Double
    Set container = ThisWorkbook.Worksheets("Расчет").ChartObjects("chtNDMSectionPlot")
    width = container.Width * 4#: height = container.Height * 4#
    container.Activate: container.Chart.CopyPicture 1, -4147
    Set scratch = ThisWorkbook.Worksheets.Add
    On Error GoTo Failed
    Set canvas = scratch.ChartObjects.Add(0#, 0#, width, height)
    canvas.Activate: canvas.Chart.Paste
    Set picture = canvas.Chart.Shapes(1)
    picture.LockAspectRatio = 0: picture.Left = 0#: picture.Top = 0#: picture.Width = width: picture.Height = height
    If Not canvas.Chart.Export(path, "PNG") Then Err.Raise vbObjectError + 5294, "modTestAudit02Api", "Не сохранён крупный снимок схемы."
    Application.DisplayAlerts = False: scratch.Delete
    Exit Sub
Failed:
    Dim number As Long, reason As String
    number = Err.Number: reason = Err.Description
    Application.DisplayAlerts = False: scratch.Delete
    Err.Raise number, "modTestAudit02Api", reason
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' смещение по касательной скругления и центр подписи асимметричного ряда у отверстия.
Public Function RunOffsetOpeningAnnotationTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    Dim config As Object, saved As Variant, target As Object, main As Long, extra As Long
    Dim settings As CSystemSettingsReader, units As CUnitSystem, builder As CSectionBuilder, model As CSectionModel
    Dim i As Long, source As Long, second As Long, third As Long, minX As Double, maxX As Double, maxY As Double
    Dim dx As Double, dy As Double, nx As Double, ny As Double, found As Boolean, annotations As CSectionAnnotations
    On Error GoTo Failed
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = TestFixtureSaveInputs(config)
    TestConfigureOffsetWorkbook "HollowOffset", "mm", 100#, 100#
    Set target = TestOffsetWorkbookTable("HollowOffset"): TestOffsetWorkbookRows target, "HollowOffset", main, extra
    TestOffsetRebarCell(extra + 4, 10).Value2 = "SideBySide": TestOffsetRebarCell(extra + 4, 9).Value2 = "EverySecondBar"
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    Set builder = New CSectionBuilder: Set model = BuildWorkbookSectionModel(ThisWorkbook, settings, units, builder)
    minX = 1E+100: maxX = -1E+100: maxY = -1E+100
    For i = 1 To model.RebarCount
        Select Case model.RebarSourceName(i)
            Case "Opening.B.Top.Main.2": source = i
            Case "Opening.B.Top.Row2.1": second = i
            Case "Opening.B.Top.Row3.1": third = i
        End Select
        If Left$(model.RebarSourceName(i), 19) = "Opening.B.Top.Main." Then
            minX = GeomMin(minX, model.RebarX(i)): maxX = GeomMax(maxX, model.RebarX(i))
            maxY = GeomMax(maxY, model.RebarY(i))
        End If
    Next i
    TestCheck "opening.offset.source.positions", source > 0 And second > 0 And third > 0
    If source = 0 Or second = 0 Or third = 0 Then GoTo Finished
    nx = model.RebarX(second) - model.RebarX(source): ny = model.RebarY(second) - model.RebarY(source)
    dx = model.RebarX(third) - model.RebarX(second): dy = model.RebarY(third) - model.RebarY(second)
    TestCheck "opening.offset.normal100", Abs(Sqr(nx * nx + ny * ny) - 100#) < 0.0000001
    TestCheck "opening.offset.profile.gap2.5", Abs(Sqr(dx * dx + dy * dy) - 20.5) < 0.0000001
    TestCheck "opening.offset.tangent.perpendicular", Abs(nx * dx + ny * dy) < 0.0000001
    TestCheck "opening.offset.corner.components", Abs(dx - 17.0453280954174) < 0.0000001 And Abs(dy - 11.3888888888889) < 0.0000001
    Set annotations = model.Annotations
    For i = 1 To annotations.Count
        If annotations.AnnotationID(i) = "REBAR_Opening.B.Top" Then
            found = True
            TestCheck "opening.annotation.full.row.start", Abs(annotations.StartX(i) - minX) < 0.0000001
            TestCheck "opening.annotation.full.row.end", Abs(annotations.EndX(i) - maxX) < 0.0000001
            TestCheck "opening.annotation.full.row.center", Abs((annotations.StartX(i) + annotations.EndX(i)) / 2#) < 0.0000001
            TestCheck "opening.annotation.extreme.support", Abs(annotations.StartY(i) - maxY) < 0.0000001 And Abs(annotations.EndY(i) - maxY) < 0.0000001
        End If
    Next i
    TestCheck "opening.annotation.found", found
    GoTo Finished
Failed:
    TestCheck "opening.annotation.runtime; " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    If Not config Is Nothing And Not IsEmpty(saved) Then TestFixtureRestoreInputs config, saved
    RunOffsetOpeningAnnotationTests = mReport & "TOTAL OFFSET OPENING ANNOTATION: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' полный круг и открытая дуга, полученная концевыми отступами;
' координаты проверяются по независимо заданным точкам полуокружности.
Public Function RunCircleSetbackTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    Dim config As Object, saved As Variant, unitName As Variant, factor As Double, table As Object
    Dim bars As CRebarLayout, original As Variant, point As Variant, expected As Variant, i As Long, column As Variant
    Dim definition As CSectionDefinition, bad As Variant, number As Long, reason As String
    Dim facade As CSectionDefinitionBuilder, path As CGeometryPath, x As Double, y As Double, tx As Double, ty As Double
    Dim executor As CRebarRuleExecutor, location As Variant, binding As Variant, secondCount As Long, thirdCount As Long, owner As Long
    On Error GoTo Failed
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = TestFixtureSaveInputs(config)
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "SliceTest"
    Set path = facade.CirclePath("Full", 0#, 0#, 100#)
    path.AtDistance 0#, x, y, tx, ty
    TestCheck "circle.setback.generic.start.unchanged", Abs(x - 100#) < 0.0000001 And Abs(y) < 0.0000001 And ty > 0#
    Set path = facade.Slice("Half", path.ID, 50# * GEOM_PI, 150# * GEOM_PI)
    TestCheck "circle.setback.shared.slice", Not path.IsClosed And Abs(path.Length - 100# * GEOM_PI) < 0.0000001
    expected = Array(Array(100#, 0#), Array(50#, -86.6025403784439), Array(-50#, -86.6025403784439), Array(-100#, 0#))
    For Each unitName In Array("mm", "cm", "m")
        TestConfigureOffsetWorkbook "Circle", CStr(unitName), 0#, 0#
        factor = 1#: If CStr(unitName) = "cm" Then factor = 10#
        If CStr(unitName) = "m" Then factor = 1000#
        TestFixtureCircleFieldCell(ThisWorkbook, "Circle.Diameter").Value2 = 300# / factor
        TestFixtureCircleFieldCell(ThisWorkbook, "Rebar.AxisDistance").Value2 = 50# / factor
        TestFixtureCircleFieldCell(ThisWorkbook, "Rebar.Count").Value2 = 4
        TestFixtureCircleFieldCell(ThisWorkbook, "Rebar.Diameter").Value2 = 16# / factor
        TestFixtureCircleFieldCell(ThisWorkbook, "Rebar.Diameter2").Value2 = 0#
        TestFixtureCircleFieldCell(ThisWorkbook, "Rebar.Diameter3").Value2 = 0#
        TestFixtureCircleFieldCell(ThisWorkbook, "Rebar.StartSetback").Value2 = 0#
        TestFixtureCircleFieldCell(ThisWorkbook, "Rebar.EndSetback").Value2 = 0#
        TestFixtureInputCell(ThisWorkbook.Names("rngCircleGeometry").RefersToRange, "Main", "Circle", "endMode").Value2 = "IncludeStart"
        Set definition = TestReadOffsetWorkbookDefinition(): Set bars = TestFixtureLayout(definition)
        original = TestOffsetBarsSnapshot(bars)
        TestCheck "circle.setback.full.closed", definition.Path("Circle.Axis").IsClosed And bars.Count = 4
        TestCheck "circle.setback.full.first.top", Abs(bars.X(1)) < 0.0000001 And Abs(bars.Y(1) - 100#) < 0.0000001
        TestCheck "circle.setback.full.clockwise", Abs(bars.X(2) - 100#) < 0.0000001 And Abs(bars.Y(2)) < 0.0000001
        TestFixtureInputCell(ThisWorkbook.Names("rngCircleGeometry").RefersToRange, "Main", "Circle", "endMode").Value2 = "IncludeBoth"
        TestCheck "circle.setback.full.bottom.left", Abs(bars.X(3)) < 0.0000001 And Abs(bars.Y(3) + 100#) < 0.0000001 And Abs(bars.X(4) + 100#) < 0.0000001 And Abs(bars.Y(4)) < 0.0000001
        For Each column In Array("Rebar.StartSetback", "Rebar.EndSetback")
            TestFixtureCircleFieldCell(ThisWorkbook, CStr(column)).Value2 = 50# * GEOM_PI / factor
        Next column
        Set definition = TestReadOffsetWorkbookDefinition(): Set bars = TestFixtureLayout(definition)
        TestCheck "circle.setback.arc.open", Not definition.Path("Circle.Axis").IsClosed And bars.Count = 4
        For i = 1 To 4
            point = expected(i - 1)
            TestCheck "circle.setback.arc.point." & CStr(unitName) & "." & CStr(i), Abs(bars.X(i) - CDbl(point(0))) < 0.0000001 And Abs(bars.Y(i) - CDbl(point(1))) < 0.0000001 And bars.Diameter(i) = 16#
        Next i
        TestFixtureCircleFieldCell(ThisWorkbook, "Rebar.EndSetback").Value2 = 0#
        Set definition = TestReadOffsetWorkbookDefinition(): Set bars = TestFixtureLayout(definition)
        TestCheck "circle.setback.start.only", Abs(bars.X(1) - 100#) < 0.0000001 And Abs(bars.Y(1)) < 0.0000001 And Abs(bars.Y(4) - 100#) < 0.0000001
        TestFixtureCircleFieldCell(ThisWorkbook, "Rebar.StartSetback").Value2 = 0#
        TestFixtureCircleFieldCell(ThisWorkbook, "Rebar.EndSetback").Value2 = 50# * GEOM_PI / factor
        Set definition = TestReadOffsetWorkbookDefinition(): Set bars = TestFixtureLayout(definition)
        TestCheck "circle.setback.end.only", Abs(bars.X(1)) < 0.0000001 And Abs(bars.Y(1) - 100#) < 0.0000001 And Abs(bars.X(4) + 100#) < 0.0000001
        TestFixtureCircleFieldCell(ThisWorkbook, "Rebar.StartSetback").Value2 = 0#
        TestFixtureCircleFieldCell(ThisWorkbook, "Rebar.EndSetback").Value2 = 0#
        TestFixtureInputCell(ThisWorkbook.Names("rngCircleGeometry").RefersToRange, "Main", "Circle", "endMode").Value2 = "IncludeStart"
        Set definition = TestReadOffsetWorkbookDefinition(): Set bars = TestFixtureLayout(definition)
        TestCheck "circle.setback.return.zero", TestOffsetBarsEqual(original, TestOffsetBarsSnapshot(bars))
        For Each column In Array("Rebar.StartSetback", "Rebar.EndSetback")
            Set table = TestFixtureCircleFieldCell(ThisWorkbook, CStr(column))
            For Each bad In Array(-1#, "bad", True, CVErr(2015), 200# * GEOM_PI / factor)
                table.Value2 = bad: number = 0: reason = ""
                On Error Resume Next
                Set definition = TestReadOffsetWorkbookDefinition()
                number = Err.Number: reason = Err.Description: Err.Clear: On Error GoTo Failed
                TestCheck "circle.setback.invalid." & CStr(column), number <> 0 And InStr(reason, table.Worksheet.Name & "!") > 0
            Next bad
            table.Value2 = 0#
        Next column
        TestFixtureCircleFieldCell(ThisWorkbook, "Rebar.Diameter2").Value2 = 12# / factor
        TestFixtureCircleFieldCell(ThisWorkbook, "Rebar.Diameter3").Value2 = 10# / factor
        TestFixtureCircleFieldCell(ThisWorkbook, "Rebar.AxisOffset2").Value2 = 20# / factor
        TestFixtureCircleFieldCell(ThisWorkbook, "Rebar.AxisOffset3").Value2 = 20# / factor
        For Each location In Array("Stacked", "SideBySide")
            TestFixtureCircleFieldCell(ThisWorkbook, "Rebar.Loc3row").Value2 = location
            For Each binding In Array("EachBar", "EverySecondBar")
                TestFixtureCircleFieldCell(ThisWorkbook, "Rebar.Bind3row").Value2 = binding
                Set definition = TestReadOffsetWorkbookDefinition(): Set executor = New CRebarRuleExecutor
                Set bars = executor.Build(definition, TestFixtureGeometry(definition.ConcreteDomain))
                secondCount = 0: thirdCount = 0
                For i = 1 To executor.PositionCount
                    owner = executor.PositionPhysicalOwner(i)
                    If executor.PositionRuleID(i) = "Circle.Row2" Then
                        secondCount = secondCount + 1
                        TestCheck "circle.clockwise.row2.inward", Abs(Sqr(bars.X(owner) ^ 2 + bars.Y(owner) ^ 2) - 80#) < 0.0000001
                    ElseIf executor.PositionRuleID(i) = "Circle.Row3" Then
                        thirdCount = thirdCount + 1
                        If thirdCount = 1 Then
                            TestCheck "circle.clockwise.row3.top", Abs(bars.X(owner) - IIf(location = "Stacked", 0#, 13.5)) < 0.0000001 And Abs(bars.Y(owner) - IIf(location = "Stacked", 66.5, 80#)) < 0.0000001
                        ElseIf thirdCount = 2 And binding = "EverySecondBar" Then
                            TestCheck "circle.clockwise.everySecond.bottom", Abs(bars.X(owner) + IIf(location = "Stacked", 0#, 13.5)) < 0.0000001 And Abs(bars.Y(owner) + IIf(location = "Stacked", 66.5, 80#)) < 0.0000001
                        End If
                    End If
                Next i
                TestCheck "circle.clockwise.row.counts", secondCount = 4 And thirdCount = IIf(binding = "EachBar", 4, 2)
            Next binding
        Next location
    Next unitName
    GoTo Finished
Failed:
    TestCheck "circle.setback.runtime; " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    If Not config Is Nothing And Not IsEmpty(saved) Then TestFixtureRestoreInputs config, saved
    RunCircleSetbackTests = mReport & "TOTAL CIRCLE SETBACKS: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' встреча рядов наружной грани и отверстия в тонкой стенке;
' отдельно проверяет диск, центр которого в бетоне, а край заходит в отверстие.
Public Function RunOffsetWallCollisionTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    Dim config As Object, saved As Variant, target As Object, main As Long, extra As Long
    Dim definition As CSectionDefinition, bars As CRebarLayout, scenario As Long, number As Long, reason As String
    On Error GoTo Failed
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = TestFixtureSaveInputs(config)
    For scenario = 0 To 2
        TestConfigureOffsetWorkbook "HollowOffset", "mm", 100#, 100#
        Set target = TestOffsetWorkbookTable("HollowOffset"): TestOffsetWorkbookRows target, "HollowOffset", main, extra
        TestFixtureInputCell(target, "General", "H", "val").Value2 = 1000#: TestFixtureInputCell(target, "Detail", "Ho", "val").Value2 = 324#: TestFixtureInputCell(target, "Detail", "Ro", "val").Value2 = 0#
        TestFixtureInputCell(target, "Detail", "Xo", "val").Value2 = 0#: TestFixtureInputCell(target, "Detail", "Yo", "val").Value2 = 0#
        TestOffsetRebarCell(extra, 7).Value2 = 0#: TestOffsetRebarCell(extra + 4, 7).Value2 = 0#
        TestOffsetRebarCell(extra + 4, 3).Value2 = 16#
        If scenario = 1 Then TestFixtureInputCell(target, "General", "H", "val").Value2 = 992#
        If scenario = 2 Then
            TestOffsetRebarCell(extra, 3).Value2 = 0#
            TestOffsetRebarCell(extra + 4, 3).Value2 = 200#: TestOffsetRebarCell(extra + 4, 4).Value2 = 1#
        End If
        Set bars = Nothing: number = 0: reason = ""
        On Error Resume Next
        Set definition = TestReadOffsetWorkbookDefinition(): Set bars = TestFixtureLayout(definition)
        number = Err.Number: reason = Err.Description: Err.Clear: On Error GoTo Failed
        mReport = mReport & "OBSERVED wall." & CStr(scenario) & ": " & CStr(number) & "; " & reason & vbCrLf
        Select Case scenario
            Case 0
                TestCheck "offset.wall.touch.allowed", number = 0
                If Not bars Is Nothing Then TestCheck "offset.wall.touch.both.families", bars.Count = 16
            Case 1
                TestCheck "offset.wall.overlap.rejected", number <> 0 And InStr(reason, "Поперечные сечения стержней") > 0 And InStr(reason, "Сечения!") > 0
            Case 2
                TestCheck "offset.wall.disk.opening.rejected", number <> 0 And InStr(reason, "целиком") > 0 And InStr(reason, "Сечения!") > 0
        End Select
    Next scenario
    GoTo Finished
Failed:
    TestCheck "offset.wall.runtime; " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    If Not config Is Nothing And Not IsEmpty(saved) Then TestFixtureRestoreInputs config, saved
    RunOffsetWallCollisionTests = mReport & "TOTAL OFFSET WALL COLLISION: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' находит ячейку дополнительного ряда в таблице выбранного типа на листе «Сечения».
Private Function TestOffsetRebarCell(ByVal row As Long, ByVal column As Long) As Object
    Set TestOffsetRebarCell = TestFixtureSelectedInput().Cells(row, column)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' адреса сообщений после переноса таблицы, числовая причина ошибки
' концевых отступов и сохранение допустимого размещения одного стержня в точке.
Public Function RunProgramMessageTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    Dim sheet As Object, alerts As Boolean, previousSheet As Object, source As Object
    Dim reader As CLoadCombinationReader, batch As CBatchSectionCalculator, meta As CResultMeta, copiedMeta As CResultMeta
    Dim settings As CSystemSettingsReader, message As String, errorNumber As Long
    Dim facade As CSectionDefinitionBuilder, path As CGeometryPath, options As CRebarRuleOptions
    Dim definition As CSectionDefinition, executor As CRebarRuleExecutor, layout As CRebarLayout
    Dim diagnostic As CGeometryDiagnostic, rectangle As CGeometryPath
    Dim profile As CCalculationProfile, locations As Object, policy As CResultStatusPolicy
    Dim otherMeta As CResultMeta, merged As CResultMeta, mesh As CFiberMeshBuilder, reason As String
    Dim searchResult As CLimitSearchResult, inputTable As CSectionInputTable, units As CUnitSystem
    Dim values(1 To 1, 1 To 1) As Variant, addresses(1 To 1, 1 To 1) As Variant, lengthValue As Double
    On Error GoTo Failed
    Set previousSheet = ThisWorkbook.ActiveSheet: alerts = Application.DisplayAlerts
    Set sheet = ThisWorkbook.Worksheets.Add
    Set source = sheet.Range("AA20:AG21")
    source.Rows(1).Value2 = Array("CombinationID", "N", "Mx", "My", "ProfileId", "LoadPath", "Comment")
    source.Rows(2).Value2 = Array("Message.LC", 1#, 2#, 3#, "Message.Profile", "Auto", "")
    Set batch = New CBatchSectionCalculator: Set reader = New CLoadCombinationReader
    reader.LoadFromRange source, batch
    TestCheck "messages.source.count", batch.Count = 1
    TestCheck "messages.source.relocated", batch.SourceReference(1) = sheet.Name & "!AA21:AG21"
    Set meta = New CResultMeta
    meta.SetResult rsNumericalFailure, rcNumericalFailure, rkDirectState, "Равновесие не найдено."
    message = CalculationResultMessage(batch, 1, meta)
    TestCheck "messages.result.combination", InStr(message, "Message.LC") > 0
    TestCheck "messages.result.address", InStr(message, sheet.Name & "!AA21:AG21") > 0
    TestCheck "messages.result.reason", InStr(message, "Равновесие не найдено.") > 0
    TestCheck "messages.result.statusUnchanged", meta.InternalStatus = rsNumericalFailure And meta.ResultCode = rcNumericalFailure
    meta.ResultComment = "": TestCheck "messages.empty.staysEmpty", CalculationResultMessage(batch, 1, meta) = ""
    Set source = sheet.Range("AA30:AD31")
    source.Rows(1).Value2 = Array("Key", "Value", "Default", "Comment")
    source.Rows(2).Value2 = Array("Solver.Method", "WrongMethod", "Newton", "")
    Set settings = New CSystemSettingsReader: settings.LoadFromRange source
    message = settings.InputErrorMessage("Solver.Method", "Неизвестный метод.", "Выберите Newton или Secant.")
    TestCheck "messages.setting.address", InStr(message, sheet.Name) > 0 And InStr(message, "AB31") > 0
    mReport = mReport & "SETTING_MESSAGE: " & message & vbCrLf
    TestCheck "messages.setting.value", InStr(message, "WrongMethod") > 0
    TestCheck "messages.setting.action", InStr(message, "Выберите Newton или Secant.") > 0
    batch.ApplySettings settings
    meta.SetSolverFailure sfcInvalidConfiguration, rkDirectState, "Выберите метод равновесия.", False, "Solver.Method"
    Set copiedMeta = meta.Clone
    TestCheck "messages.keys.clone", copiedMeta.SourceKeys = "Solver.Method"
    message = CalculationResultMessage(batch, 1, copiedMeta)
    TestCheck "messages.keys.actualAddress", InStr(message, sheet.Name) > 0 And InStr(message, "AB31") > 0
    TestCheck "messages.keys.noStatusParsing", copiedMeta.InternalStatus = rsInvalidConfiguration And copiedMeta.ResultCode = rcInvalidConfiguration
    copiedMeta.SetNotRequested rkDirectState
    TestCheck "messages.keys.clearedForNewResult", copiedMeta.SourceKeys = ""
    Set policy = New CResultStatusPolicy: Set otherMeta = New CResultMeta
    otherMeta.SetSolverFailure sfcInvalidConfiguration, rkCapacity, "Другой параметр.", False, "Capacity.MaxLambda;Solver.Method"
    Set merged = policy.WorstResultMeta(rkGeneric, meta, otherMeta)
    TestCheck "messages.keys.merge.first", InStr(merged.SourceKeys, "Solver.Method") > 0
    TestCheck "messages.keys.merge.second", InStr(merged.SourceKeys, "Capacity.MaxLambda") > 0
    TestCheck "messages.keys.merge.noDuplicates", UBound(Split(merged.SourceKeys, ";")) = 1
    Set searchResult = New CLimitSearchResult
    searchResult.Initialize meta, "Auto", "LoadMultiplier", False, 0#, 0#, 0#, 0#, Nothing, "InvalidInput", "", vbNullString, False
    searchResult.AddLoadPathHistory "Первая попытка не завершена.", "lambda*N", True
    TestCheck "messages.keys.pathHistory", searchResult.Meta.SourceKeys = "Solver.Method"
    Set profile = New CCalculationProfile
    profile.Initialize "Message.Profile", "", "", True, False, False
    Set locations = CreateObject("Scripting.Dictionary")
    locations.Add "Profile.DisplayName", sheet.Name & "!AZ50"
    profile.SetInputLocations locations
    TestCheck "messages.profile.invalid", Not profile.ValidateForCalculation(reason)
    TestCheck "messages.profile.address", InStr(reason, sheet.Name & "!AZ50") > 0
    TestCheck "messages.profile.action", InStr(reason, "имя профиля") > 0
    message = WorkbookErrorMessage("Расчёт остановлен", 12345, "Probe.Stage", "")
    TestCheck "messages.runtime.code", InStr(message, "12345") > 0
    TestCheck "messages.runtime.stage", InStr(message, "Probe.Stage") > 0
    TestCheck "messages.runtime.reason", InStr(message, "без описания причины") > 0

    Set facade = New CSectionDefinitionBuilder: facade.Initialize "Messages"
    Set rectangle = facade.RectanglePath("Boundary", 0#, 0#, 1000#, 1000#)
    facade.Outer "OUTER", rectangle.ID
    Set path = facade.LinePath("Source", 100#, 400#, 100#, 600#)
    facade.RebarByCount "Source.Main", path.ID, 3, 10#
    Set path = facade.LinePath("Target", 200#, 430#, 200#, 570#)
    Set options = New CRebarRuleOptions
    options.StartSetback = 80#: options.EndSetback = 60#: options.DirectionX = 1#
    options.SourceReference = sheet.Name & "!AA40; " & sheet.Name & "!AB40"
    options.Comment = "Арматура справа от отверстия"
    facade.RebarProjected "Target.Main", path.ID, "Source.Main", 10#, options
    Set definition = facade.Finish: Set executor = New CRebarRuleExecutor
    On Error Resume Next
    Set layout = executor.Build(definition, TestFixtureGeometry(definition.ConcreteDomain))
    errorNumber = Err.Number: message = Err.Description: Err.Clear
    On Error GoTo Failed
    Set diagnostic = executor.LastDiagnostic
    TestCheck "messages.interval.error", errorNumber = vbObjectError + 5530
    TestCheck "messages.interval.code", diagnostic.Code = "WORK_INTERVAL_EMPTY"
    TestCheck "messages.interval.source", InStr(message, sheet.Name & "!AA40") > 0 And InStr(message, sheet.Name & "!AB40") > 0
    TestCheck "messages.interval.numbers", InStr(message, "140") > 0 And InStr(message, "80") > 0 And InStr(message, "60") > 0
    TestCheck "messages.interval.caption", InStr(message, "Арматура справа от отверстия") > 0
    TestCheck "messages.interval.action", InStr(message, "Уменьшите концевые отступы") > 0
    TestCheck "messages.interval.details", Abs(diagnostic.Detail1 - 140#) < 0.000001 And Abs(diagnostic.Detail2 - 140#) < 0.000001
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "PointMessage"
    Set rectangle = facade.RectanglePath("Boundary", 0#, 0#, 1000#, 1000#)
    facade.Outer "OUTER", rectangle.ID
    Set path = facade.LinePath("Axis", 200#, 430#, 200#, 570#)
    facade.RebarByCount "Point.Main", path.ID, 1, 10#, options
    Set layout = TestFixtureLayout(facade.Finish)
    TestCheck "messages.interval.oneBarStillAllowed", layout.Count = 1
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "ExplicitPointMessage"
    Set rectangle = facade.RectanglePath("Boundary", 0#, 0#, 1000#, 1000#)
    facade.Outer "OUTER", rectangle.ID
    Set path = facade.LinePath("Axis", 200#, 430#, 200#, 570#)
    facade.RebarAtStations "Point.Main", path.ID, Array(80#), 10#, options
    Set layout = TestFixtureLayout(facade.Finish)
    TestCheck "messages.interval.explicitPointStillAllowed", layout.Count = 1
    Set mesh = New CFiberMeshBuilder
    On Error Resume Next
    mesh.BuildMesh TestFixtureGeometry(definition.ConcreteDomain), 0.00000000001, 100#, 1, 1, sheet.Name & "!AC60"
    errorNumber = Err.Number: message = Err.Description: Err.Clear
    On Error GoTo Failed
    TestCheck "messages.mesh.existingCode", errorNumber = vbObjectError + 2108
    TestCheck "messages.mesh.address", InStr(message, sheet.Name & "!AC60") > 0
    source.Rows(2).Value2 = Array("Units.Length.Input", "m", "mm", "")
    settings.LoadFromRange source
    Set units = New CUnitSystem: units.LoadFromSettings settings
    values(1, 1) = 1E+308: addresses(1, 1) = sheet.Name & "!AD70"
    Set inputTable = New CSectionInputTable: inputTable.Initialize "Message.Geometry", values, addresses
    On Error Resume Next
    lengthValue = inputTable.Length(1, 1, units)
    errorNumber = Err.Number: message = Err.Description: Err.Clear
    On Error GoTo Failed
    TestCheck "messages.length.overflowCode", errorNumber = vbObjectError + 5525
    TestCheck "messages.length.overflowAddress", InStr(message, sheet.Name & "!AD70") > 0
    TestCheck "messages.length.overflowAction", InStr(message, "Уменьшите значение") > 0
    GoTo Restore
Failed:
    TestCheck "messages.runtime", False
    mReport = mReport & "ERROR: " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Restore:
    On Error Resume Next
    Application.DisplayAlerts = False
    If Not previousSheet Is Nothing Then previousSheet.Activate
    If Not sheet Is Nothing Then sheet.Delete
    Application.DisplayAlerts = alerts
    On Error GoTo 0
    RunProgramMessageTests = mReport & "TOTAL PROGRAM MESSAGES: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function


' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет реальное предупреждение после счётчиков, несколько причин и точность цитируемых пользовательских строк.
Private Sub TestGeometryWarningParagraphs()
    Dim warning As String, message As String, text As String, expected As String
    warning = "Предупреждение: пропущено совпавших стержней: 8. В каждой совпавшей позиции сохранён первый добавленный стержень с его диаметром и материалом. " & _
        "На схеме и в расчёте учтены фактически созданные стержни. Подробности и адреса полей приведены в Geometry.RecipeDiagnostics на Results."
    message = "Сечение сформировано и сохранено на Results." & vbCrLf & vbCrLf & _
        "Бетонных элементов: 1232." & vbCrLf & "Стержней арматуры: 20."
    text = WorkbookDialogText(WorkbookMessageParagraph(message, warning))
    expected = message & vbCrLf & vbCrLf & "Предупреждение:" & vbCrLf & "пропущено совпавших стержней: 8." & vbCrLf & vbCrLf & _
        "В каждой совпавшей позиции сохранён первый добавленный стержень с его диаметром и материалом. На схеме и в расчёте учтены фактически созданные стержни." & vbCrLf & vbCrLf & _
        "Подробности и адреса полей:" & vbCrLf & "Geometry.RecipeDiagnostics на Results."
    TestCheck "presentation.geometry.warning.blocks", text = expected
    TestCheck "presentation.geometry.warning.single.newline", WorkbookDialogText(message & vbCrLf & warning) = expected
    TestCheck "presentation.geometry.warning.idempotent", WorkbookDialogText(text) = text
    TestCheck "presentation.geometry.warning.raw.compact", InStr(warning, vbCrLf) = 0
    mReport = mReport & "GEOMETRY_WARNING_LAYOUT:" & vbCrLf & text & vbCrLf & "END_GEOMETRY_WARNING_LAYOUT" & vbCrLf
    warning = "Предупреждение: после смещения не размещено дополнительных стержней: 2. Их положения по нормали оказались вне полного смещённого контура. " & _
        "Подробности и адреса полей приведены в Geometry.RecipeDiagnostics на Results." & vbCrLf & _
        "Предупреждение: объединено дополнительных стержней соседних граней: 3. В расчёте и подписях учтён только первый добавленный стержень с его координатами и диаметром." & vbCrLf & _
        "По настройке Mesh.MinElementArea пропущено стержней: 4. Площадь каждого из них меньше заданного порога; в расчёте и подписях они не учитываются."
    text = WorkbookDialogText(warning)
    TestCheck "presentation.geometry.warning.offset.reason", InStr(text, "2." & vbCrLf & vbCrLf & "Их положения по нормали") > 0
    TestCheck "presentation.geometry.warning.neighbor.reason", InStr(text, "3." & vbCrLf & vbCrLf & "В расчёте и подписях") > 0
    TestCheck "presentation.geometry.warning.area.reason", InStr(text, "4." & vbCrLf & vbCrLf & "Площадь каждого из них") > 0
    TestCheck "presentation.geometry.warning.several.idempotent", WorkbookDialogText(text) = text
    TestCheck "presentation.geometry.warning.no.extra.blanks", InStr(text, vbCrLf & vbCrLf & vbCrLf) = 0
    warning = "«Предупреждение: 8. В каждой совпавшей позиции имя. Подробности и адреса полей приведены в Имя»"
    TestCheck "presentation.geometry.warning.quoted", WorkbookDialogText(warning) = warning
End Sub
