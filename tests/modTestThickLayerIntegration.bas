Attribute VB_Name = "modTestThickLayerIntegration"
Option Explicit

#If VBA7 Then
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef value As Currency) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef value As Currency) As Long
#Else
Private Declare Function QueryPerformanceCounter Lib "kernel32" (ByRef value As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (ByRef value As Currency) As Long
#End If

' ==========================================================================
' ДЛЯ ТЕСТОВ: Непрерывное MTLI-интегрирование R/K по точным контурам
' ==========================================================================
' Аналитические интегралы не зависят от квадратуры проверяемого класса.
' Проверяет отверстия, перенос, поворот, круговые дуги, узкую сжатую полосу,
' переходы материала и производные усилий. Это не полная приёмка книги.

Private mPassed As Long, mFailed As Long, mReport As String
Private Const PI As Double = 3.1415926535897931

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Повторяет связанные численные проверки на окончательных исходниках.
' Полные ON/OFF и интерфейсный набор не входят в этот состав.
Public Function RunThickLayerFinalDirectedTests() As String
    On Error GoTo Failed
    Dim report As String
    report = RunThickLayerSolverTests()
    report = report & RunThickLayerCapacityTests()
    report = report & RunThickLayerEdgeTests()
    report = report & RunThickLayerOracleTests()
    report = report & RunThickLayerRelatedTests()
    RunThickLayerFinalDirectedTests = report & "TOTAL_MTLI_FINAL_DIRECTED: finished=5"
    Exit Function
Failed:
    RunThickLayerFinalDirectedTests = report & "RUNTIME ERROR: " & Err.Number & "; " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Экспортирует R/K для независимого интегрирования площади в Python.
' Эталон не использует обход контура или формулы проверяемого класса.
Public Function RunThickLayerOracleTests() As String
    On Error GoTo Failed
    mPassed = 0: mFailed = 0: mReport = vbNullString
    Dim query As CSectionGeometryQuery: Set query = New CSectionGeometryQuery
    Dim material As CMaterialDiagram: Set material = New CMaterialDiagram
    Dim strains(1 To 7) As Double, stresses(1 To 7) As Double
    strains(1) = -0.004: strains(2) = -0.002: strains(3) = -0.001
    strains(5) = 0.00005: strains(6) = 0.0001: strains(7) = 0.004
    stresses(1) = -40#: stresses(2) = -30#: stresses(3) = -20#: stresses(5) = 2#
    material.InitializeFromArrays strains, stresses, 7
    Dim kind As Long, testCase As Long, row As Long, column As Long
    Dim region As CConcreteRegion, hole As CConcreteRegion, integrator As CThickLayerIntegrator
    Dim epsilon As Double, kx As Double, ky As Double, record As String
    For kind = 1 To 4
        Set region = query.RectangleRegion(0#, 0#, 200#, 400#)
        If kind = 2 Then
            Set hole = query.RectangleRegion(0#, 0#, 100#, 200#)
            Set region = query.SubtractContainedRegion(region, hole)
        End If
        If kind >= 3 Then Set region = query.CircleRegion(0#, 0#, 100#)
        If kind = 4 Then
            Set hole = query.CircleRegion(0#, 0#, 50#)
            Set region = query.SubtractContainedRegion(region, hole)
        End If
        Set integrator = New CThickLayerIntegrator: integrator.Initialize region, 0#, 0#
        For testCase = 1 To 4
            epsilon = -0.00031: kx = 0.000013: ky = -0.000011
            If testCase = 2 Then epsilon = 0.000025: kx = 0#: ky = 0#
            If testCase = 3 Then epsilon = -0.0013: kx = -0.000007: ky = 0.000023
            If testCase = 4 Then epsilon = 0.0009: kx = 0.000011: ky = 0.000003
            integrator.Evaluate material, epsilon, kx, ky
            record = "DATA_MTLI_ORACLE|" & kind & "|" & testCase & "|" & TestNumber(epsilon) & "|" & TestNumber(kx) & "|" & TestNumber(ky)
            For row = 1 To 3: record = record & "|" & TestNumber(integrator.Force(row)): Next row
            For row = 1 To 3: For column = 1 To 3
                record = record & "|" & TestNumber(integrator.Tangent(row, column))
            Next column: Next row
            mReport = mReport & record & vbCrLf
            TestTrue "OracleExport_" & kind & "_" & testCase, integrator.EvaluationCount > 0
        Next testCase
    Next kind
    RunThickLayerOracleTests = mReport & "TOTAL_MTLI_ORACLE: passed=" & mPassed & "; failed=" & mFailed
    Exit Function
Failed:
    RunThickLayerOracleTests = mReport & "RUNTIME ERROR: " & Err.Number & "; " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Запускает существующие адресные проверки затронутого ядра при выключенном
' эксперименте. Состав не включает полную ON/OFF или Excel UI приёмку.
Public Function RunThickLayerRelatedTests() As String
    On Error GoTo Failed
    Dim report As String
    report = modTestSectionSolver.RunCenteredNumericalTests()
    report = report & modTestSectionSolver.RunContinuationStateTests()
    report = report & modTestCapacitySolver.RunCommonStrainLimitTests()
    report = report & modTestCapacitySolver.RunAxialCapacityRegressionTests()
    report = report & modTestCapacitySolver.RunUltimateFrameRegressionTests()
    report = report & modTestCrackWidth.RunFormationOutcomeTests()
    report = report & modTestCalculationSnapshot.RunPreparedSnapshotReuseTests()
    report = report & modTestBatchCalculation.RunAudit02RepositoryContextTests()
    RunThickLayerRelatedTests = report & "TOTAL_MTLI_RELATED: finished=8"
    Exit Function
Failed:
    RunThickLayerRelatedTests = report & "RUNTIME ERROR: " & Err.Number & "; " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет четыре численных варианта полной команды на двух сочетаниях.
' Геометрия создаётся один раз до таймера; команда включает штатный вывод,
' файл снимка и схему. Все исходные данные находятся в частной книге.
Public Function RunThickLayerWorkflowSmokeTests() As String
    RunThickLayerWorkflowSmokeTests = TestWorkflowSeries(True)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Выполняет сравнение в переменном порядке на двух размерах сетки:
' три повтора для небольшой сетки, один для густой (оценочный замер).
' Включает прочность, трещины и устойчивость. Это адресный benchmark,
' а не полная приёмка ON/OFF или интерфейса Excel.
Public Function RunThickLayerWorkflowBenchmarkTests() As String
    RunThickLayerWorkflowBenchmarkTests = TestWorkflowSeries(False)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Выполняет те же исходные задачи на исходном релизном VBA. Отдельный
' сборщик пробы импортирует только функции этого benchmark, без MTLI-кода.
Public Function RunThickLayerWorkflowBaselineBenchmarkTests() As String
    RunThickLayerWorkflowBaselineBenchmarkTests = TestWorkflowSeries(False, True)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сохраняет первичные таблицы каждого запуска отдельным UTF-8 файлом.
' Параметры физической модели, нагрузки и оформление одинаковы для вариантов.
Private Function TestWorkflowSeries(ByVal smoke As Boolean, Optional ByVal baselineOnly As Boolean = False) As String
    On Error GoTo Failed
    mPassed = 0: mFailed = 0: mReport = vbNullString
    Dim family As String, caseNumber As Long, caseCount As Long, repeats As Long, repeat As Long, position As Long, firstRepeat As Long
    Dim variantNumber As Long, variantCount As Long, relative As Double, stepSize As Double, method As String
    caseCount = 2: repeats = 3: variantCount = 8
    If smoke Then caseCount = 1: repeats = 1: variantCount = 4
    If baselineOnly Then variantCount = 1
    For caseNumber = 1 To caseCount
        firstRepeat = 0: repeats = 3
        If smoke Then repeats = 1
        If caseNumber = 2 Then firstRepeat = 1: repeats = 1 ' Густая сетка: одиночный полный замер без утверждения о статистической устойчивости времени.
        family = "CAPACITY": stepSize = 25#
        If Not smoke Then
            family = "MIXED"
            If caseNumber = 2 Then stepSize = 5#
            If caseNumber = 2 Then family = "CAPACITY_CRACK_STABILITY_SP63"
        End If
        modTestPerformance.TestConfigurePerformanceFixture family, 2, "Newton", "Rectangles", "Normal", "Yes"
        If Not smoke Then TestWorkflowCrackedLoads
        TestWorkbookSetting "General.NonCriticalMessagesEnabled", "No"
        TestWorkbookSetting "General.ExecutionReportEnabled", "No"
        TestWorkbookSetting "Results.SnapshotStorage", "File"
        TestWorkbookSetting "Results.SnapshotName", "MTLI_Workflow_" & caseNumber
        TestWorkbookSetting "Mesh.StepMode", "User"
        TestWorkbookSetting "Mesh.StepX", CStr(stepSize): TestWorkbookSetting "Mesh.StepY", CStr(stepSize)
        TestWorkbookSetting "Plot.Enabled", "Yes": TestWorkbookSetting "Plot.AutoUpdateAfterCalculation", "Yes"
        TestWorkbookSetting "General.ReformatResultsEnabled", "Yes"
        Dim torque As Object: Set torque = ThisWorkbook.Names("rngTorsionCombinations").RefersToRange
        torque.Offset(1, 0).Resize(torque.Rows.Count - 1, torque.Columns.Count).ClearContents
        Dim formed As String, geometryStore As CCalculationSnapshotStore
        Set geometryStore = New CCalculationSnapshotStore
        If Len(Dir$(geometryStore.SnapshotPath(ThisWorkbook))) = 0 Then
            formed = modWorkbookCalculation.FormSectionForWorkbook(ThisWorkbook)
        Else
            Dim existingGeometry As CCalculationSnapshot: Set existingGeometry = geometryStore.LoadFile(ThisWorkbook)
            existingGeometry.Validate
            formed = "Геометрия взята из отдельно проверенного снимка; вне таймера команды."
            Set existingGeometry = Nothing
        End If
        mReport = mReport & "WORKFLOW_FORMATION|" & caseNumber & "|" & Replace(formed, vbCrLf, " ") & vbCrLf
        Dim experiment As Object: Set experiment = ThisWorkbook.Names("rngExperimentalNumericsSettings").RefersToRange
        For repeat = firstRepeat To repeats
            For position = 0 To variantCount - 1
                variantNumber = (position + repeat) Mod variantCount
                method = "Discrete": If variantNumber Mod 2 = 1 Then method = "MTLI"
                relative = 0#: If variantNumber >= 2 Then relative = 0.001
                If variantNumber >= 4 Then relative = 0.01
                experiment.Cells(2, 2).Value2 = method: experiment.Cells(3, 2).Value2 = relative
                TestWorkbookSetting "Capacity.ToleranceStrain", "0.00001"
                TestWorkbookSetting "Capacity.ToleranceLambda", "0.01"
                TestWorkbookSetting "SLS.Crack.ToleranceLambda", "0.01"
                If variantNumber >= 6 Then
                    TestWorkbookSetting "Capacity.ToleranceStrain", "0.000035"
                    TestWorkbookSetting "Capacity.ToleranceLambda", "0.02"
                    TestWorkbookSetting "SLS.Crack.ToleranceLambda", "0.02"
                End If
                Dim seconds As Double, message As String, outputFile As String, started As Double
                started = TestWorkflowNow()
                message = modWorkbookCalculation.RunSectionCalculationForWorkbook(ThisWorkbook, False)
                seconds = TestWorkflowNow() - started
                outputFile = "Workflow_" & caseNumber & "_" & variantNumber & "_" & repeat & ".txt"
                mReport = mReport & "DATA_MTLI_WORKFLOW|" & caseNumber & "|" & variantNumber & "|" & repeat & "|" & TestNumber(seconds) & "|" & _
                    TestNumber(relative) & "|" & method & "|" & SectionEquilibriumSolveCount() & "|" & SectionEquilibriumIterationCount() & "|" & outputFile & vbCrLf
                mReport = mReport & "WORKFLOW_MESSAGE|" & caseNumber & "|" & variantNumber & "|" & repeat & "|" & Replace(message, vbCrLf, " ") & vbCrLf
                TestCaptureWorkflowSnapshot ThisWorkbook.Path & Application.PathSeparator & outputFile
                TestTrue "Workflow_" & caseNumber & "_" & variantNumber & "_" & repeat, SectionEquilibriumSolveCount() > 0
            Next position
        Next repeat
    Next caseNumber
    TestWorkflowSeries = mReport & "TOTAL_MTLI_WORKFLOW: passed=" & mPassed & "; failed=" & mFailed
    Exit Function
Failed:
    TestWorkflowSeries = mReport & "RUNTIME ERROR: " & Err.Number & "; " & Err.Source & "; " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Нагрузки вызывают реальное раскрытие трещин, а не только прохождение
' нерастрескавшегося случая. Их единицы и знаки преобразует общий CUnitSystem.
Private Sub TestWorkflowCrackedLoads()
    Dim settings As CSystemSettingsReader, units As CUnitSystem
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    Dim loads As Object, duration As Object, i As Long, n As Double, mx As Double, my As Double
    Set loads = ThisWorkbook.Names("rngLoadCombinations").RefersToRange
    Set duration = ThisWorkbook.Names("rngStabilityDurationLoads").RefersToRange
    For i = 1 To 2
        n = -150000#: mx = 800000000#: my = 200000000#
        If i = 2 Then n = -100000#: mx = -700000000#: my = 150000000#
        loads.Cells(i + 1, 2).Value2 = n / units.InputForceToInternal(1#)
        loads.Cells(i + 1, 3).Value2 = mx / units.InputMomentMxToInternal(1#)
        loads.Cells(i + 1, 4).Value2 = my / units.InputMomentMyToInternal(1#)
        loads.Cells(i + 1, 6).Value2 = "lambda*Mxy"
        duration.Cells(i + 1, 2).Value2 = n * 0.5 / units.InputForceToInternal(1#)
        duration.Cells(i + 1, 3).Value2 = mx * 0.5 / units.InputMomentMxToInternal(1#)
        duration.Cells(i + 1, 4).Value2 = my * 0.5 / units.InputMomentMyToInternal(1#)
    Next i
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Читает только уже сохранённые результаты; время их экспорта находится
' за пределами измерения пользовательской команды.
Private Sub TestCaptureWorkflowSnapshot(ByVal outputPath As String)
    Dim store As CCalculationSnapshotStore: Set store = New CCalculationSnapshotStore
    Dim snapshot As CCalculationSnapshot: Set snapshot = store.LoadFile(ThisWorkbook)
    snapshot.Validate
    Dim stream As Object: Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2: stream.Charset = "utf-8": stream.Open
    Dim key As Variant, values As Variant, row As Long, column As Long, tokens() As String, text As String
    For Each key In Array("Summary", "Strength.Headers", "Strength.Data", "Cracks.Headers", "Cracks.Data", "Stability.Labels", "Stability.Data", "Calculation.Profiles")
        values = snapshot.Table(CStr(key))
        stream.WriteText "TABLE|" & key & "|" & UBound(values, 1) & "|" & UBound(values, 2) & vbCrLf
        ReDim tokens(1 To UBound(values, 2))
        For row = 1 To UBound(values, 1)
            For column = 1 To UBound(values, 2)
                If IsError(values(row, column)) Then
                    text = "ERROR"
                ElseIf IsEmpty(values(row, column)) Then
                    text = ""
                Else
                    text = Replace(Replace(CStr(values(row, column)), vbCr, "\r"), vbLf, "\n")
                End If
                tokens(column) = "CELL|" & key & "|" & row & "|" & column & "|" & VarType(values(row, column)) & "|" & text
            Next column
            stream.WriteText Join(tokens, vbCrLf) & vbCrLf
        Next row
    Next key
    stream.SaveToFile outputPath, 2: stream.Close
    Dim geometry As Variant: geometry = snapshot.Table("rngNDMSectionGeometry")
    mReport = mReport & "WORKFLOW_SNAPSHOT|" & snapshot.RunID & "|" & UBound(geometry, 1) - 1 & "|" & UBound(geometry, 2) & vbCrLf
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Находит строку общей настройки по ключу в текущей именованной таблице.
Private Sub TestWorkbookSetting(ByVal key As String, ByVal value As String)
    Dim table As Object: Set table = ThisWorkbook.Names("rngSystemSettings").RefersToRange
    Dim row As Long
    For row = 2 To table.Rows.Count
        If CStr(table.Cells(row, 1).Value2) = key Then table.Cells(row, 2).Value2 = value: Exit Sub
    Next row
    Err.Raise 5, "modTestThickLayerIntegration", "Не найдена настройка " & key
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает монотонное время высокого разрешения для коротких запусков.
Private Function TestWorkflowNow() As Double
    Dim counter As Currency, frequency As Currency
    If QueryPerformanceCounter(counter) = 0 Or QueryPerformanceFrequency(frequency) = 0 Then Err.Raise 5, , "Высокоточный таймер недоступен."
    TestWorkflowNow = counter / frequency
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет настройки, восстановление, продолжения и нелинейные случаи,
' которые не охватывает сравнение трёх обычных задач изгиба.
Public Function RunThickLayerEdgeTests() As String
    On Error GoTo Failed
    mPassed = 0: mFailed = 0: mReport = vbNullString
    TestIntegrationSettings
    TestContinuousBoundaryAndSnapshot
    TestExperimentalContextIdentity
    TestSignedStateMatrix
    TestAxialLoadMultiplier
    RunThickLayerEdgeTests = mReport & "TOTAL_MTLI_EDGE: passed=" & mPassed & "; failed=" & mFailed
    Exit Function
Failed:
    RunThickLayerEdgeTests = mReport & "RUNTIME ERROR: " & Err.Number & "; " & Err.Source & "; " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Различает соседние Double допуска, даже если CStr выдаёт одну строку,
' и запрещает смешивать дискретное и непрерывное НДС в общем хранилище.
Private Sub TestExperimentalContextIdentity()
    Dim section As CSectionModel: Set section = TestExperimentSection(1, 10#)
    Dim concrete As CConcreteMaterialParameters: Set concrete = New CConcreteMaterialParameters
    concrete.Initialize 15.5, 1.1, 22#, 1.8, 32500#, 32500#, rbMc2:=14.6
    Dim steel As CSteelMaterialParameters: Set steel = New CSteelMaterialParameters
    steel.Initialize 350#, 350#, 390#, 390#, 200000#, 200000#
    Dim materials As CMaterialModelProvider: Set materials = New CMaterialModelProvider
    materials.TestInitializeFromParameters concrete, steel
    Dim repository As CStateRepository: Set repository = New CStateRepository
    Dim provider As CStateProvider: Set provider = New CStateProvider
    provider.Initialize section, materials, repository
    provider.UseThickLayerIntegration = True: provider.RelativeEquilibriumTolerance = 0.001
    Dim spec As CMaterialModelSpec: Set spec = New CMaterialModelSpec
    spec.Initialize "ULS(I)", "ThreeLine", "Ignore", "TwoLine"
    Dim request As CStateRequest: Set request = New CStateRequest
    request.Initialize sstStrengthState, cpStrength, spec, -100000#, 4000000#, 3000000#, True, False
    Dim state As CSectionStateResult: Set state = provider.GetOrSolve(request)
    TestTrue "ExactContext_InitialConverged", state.Converged
    Set state = provider.GetOrSolve(request)
    TestTrue "ExactContext_UnchangedReused", provider.LastStateWasReused
    Dim nextRelative As Double: nextRelative = 0.001 + 2.17E-19
    TestTrue "ExactContext_CStrCollisionFixture", nextRelative <> 0.001 And CStr(nextRelative) = CStr(0.001)
    provider.RelativeEquilibriumTolerance = nextRelative
    Set state = provider.GetOrSolve(request)
    TestTrue "ExactContext_NeighbourNotReused", state.Converged And Not provider.LastStateWasReused
    Set state = provider.GetOrSolve(request)
    TestTrue "ExactContext_NeighbourRepeatReused", provider.LastStateWasReused
    provider.UseThickLayerIntegration = False
    Set state = provider.GetOrSolve(request)
    TestTrue "ExactContext_ModeNotMixed", state.Converged And Not provider.LastStateWasReused
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Читает действительную таблицу через общий reader и проверяет адресные
' отказы неизвестного способа и недопустимой точности, восстанавливая ввод.
Private Sub TestIntegrationSettings()
    Dim table As Object: Set table = ThisWorkbook.Names("rngExperimentalNumericsSettings").RefersToRange
    Dim oldValues As Variant: oldValues = table.Value2
    Dim settings As CSystemSettingsReader: Set settings = New CSystemSettingsReader
    Dim useMTLI As Boolean, relative As Double, message As String, errorNumber As Long
    table.Cells(2, 2).Value2 = "MTLI": table.Cells(3, 2).Value2 = 0.001
    settings.LoadFromWorkbook ThisWorkbook: settings.ReadIntegrationSettings useMTLI, relative
    TestTrue "Settings_CommonReader_MTLI", useMTLI
    TestClose "Settings_CommonReader_Relative", relative, 0.001, 0#, 0#
    Dim solver As CSectionSolver: Set solver = New CSectionSolver
    solver.ApplySettings settings
    TestTrue "Settings_Solver_MTLI", solver.UseThickLayerIntegration
    table.Cells(2, 2).Value2 = "OtherMethod": settings.LoadFromWorkbook ThisWorkbook
    On Error Resume Next
    settings.ReadIntegrationSettings useMTLI, relative
    errorNumber = Err.Number: message = Err.Description: Err.Clear
    On Error GoTo 0
    TestTrue "Settings_UnknownMethod", errorNumber <> 0 And InStr(message, "Solver.IntegrationMethod") > 0 And _
        Len(settings.InputLocation("Solver.IntegrationMethod")) > 0 And InStr(message, settings.InputLocation("Solver.IntegrationMethod")) > 0
    mReport = mReport & "SETTINGS_DIAGNOSTIC: " & message & vbCrLf
    table.Cells(2, 2).Value2 = "Discrete": table.Cells(3, 2).Value2 = 0.2: settings.LoadFromWorkbook ThisWorkbook
    On Error Resume Next
    settings.ReadIntegrationSettings useMTLI, relative
    errorNumber = Err.Number: message = Err.Description: Err.Clear
    On Error GoTo 0
    TestTrue "Settings_OutOfRange", errorNumber <> 0 And InStr(message, "Solver.RelativeEquilibriumTolerance") > 0
    table.Value2 = oldValues
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Только граница выходит в продолжение: все центры ячеек ещё физически
' допустимы. Проверяет диапазон, признак extension и восстановление НДС.
Private Sub TestContinuousBoundaryAndSnapshot()
    Dim section As CSectionModel: Set section = TestExperimentSection(1, 5#)
    Dim e(1 To 4) As Double, s(1 To 4) As Double
    e(1) = -0.01: e(2) = -0.0035: e(3) = 0#: e(4) = 0.025
    s(1) = -100#: s(2) = -35#
    Dim concrete As CMaterialDiagram, steel As CMaterialDiagram
    Set concrete = New CMaterialDiagram: concrete.InitializeFromArrays e, s, 4, -0.0035, 0.025, True, False, -0.0035
    Set steel = TestNonlinearSteel()
    Dim solver As CSectionSolver: Set solver = TestNewSolver(True, 0#)
    solver.EvaluateStrainPlane section, concrete, steel, -0.003301, 0.000001, 0#
    TestTrue "Boundary_Confirm", solver.ConfirmEquilibrium(solver.Nint, solver.Mxint, solver.Myint)
    TestClose "Boundary_ExactMinimum", solver.MinConcreteStrain, -0.003501, 0#, 0.000000000001
    Dim runner As CStateSolutionRunner: Set runner = New CStateSolutionRunner
    TestTrue "Boundary_Extension", runner.StateUsesExtension(section, solver, concrete, steel)
    TestTrue "Boundary_PhysicalRange", Not runner.StateWithinPhysicalRange(section, solver, concrete, steel)
    Dim state As CSectionStateResult, spec As CMaterialModelSpec, restored As CSectionSolver
    Set state = New CSectionStateResult: Set spec = New CMaterialModelSpec
    spec.Initialize "ULS(I)", "ThreeLine", "Ignore", "TwoLine"
    state.InitializeFromSolver sstStrengthState, cpStrength, spec, solver, True, False
    runner.UseThickLayerIntegration = True
    runner.ToleranceN = 1#: runner.ToleranceMx = 1000#: runner.ToleranceMy = 1000#
    Set restored = runner.RestoreSnapshot(section, concrete, steel, state)
    TestTrue "Snapshot_Restored", Not restored Is Nothing
    If Not restored Is Nothing Then
        TestClose "Snapshot_N", restored.Nint, solver.Nint, 0#, 0.000001
        TestClose "Snapshot_Mx", restored.Mxint, solver.Mxint, 0#, 0.001
        TestClose "Snapshot_Minimum", restored.MinConcreteStrain, solver.MinConcreteStrain, 0#, 0#
    End If
    solver.ClearResult
    TestClose "Boundary_ClearMinimum", solver.MinConcreteStrain, 0#, 0#, 0#
    TestClose "Boundary_ClearMaximum", solver.MaxConcreteStrain, 0#, 0#, 0#
    solver.EvaluateStrainPlane section, concrete, steel, 0.001, 0#, 0#
    TestClose "Boundary_UniformTensionMinimum", solver.MinConcreteStrain, 0.001, 0#, 0.000000000001
    TestClose "Boundary_UniformTensionMaximum", solver.MaxConcreteStrain, 0.001, 0#, 0.000000000001
    section.Contours.Clear
    solver.Solve section, concrete, steel, -10000#, 0#, 0#
    TestTrue "MissingContour_Stopped", Not solver.Converged And solver.FailureCode = sfcInvalidInput And InStr(solver.StopReason, "контур") > 0
    TestClose "MissingContour_NoStaleMinimum", solver.MinConcreteStrain, 0#, 0#, 0#
    TestClose "MissingContour_NoStaleMaximum", solver.MaxConcreteStrain, 0#, 0#, 0#
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет знаки, осевое растяжение, чистое сжатие, круглую геометрию,
' независимый Secant и асимметричную сталь. Ход решения не сравнивается.
Private Sub TestSignedStateMatrix()
    Dim concrete As CMaterialDiagram, steel As CMaterialDiagram
    Set concrete = TestNonlinearConcrete(): Set steel = TestAsymmetricSteel()
    Dim kind As Long, loadCase As Long, method As Variant, mode As Long, precision As Long
    Dim section As CSectionModel, probe As CSectionSolver, baseline As CSectionSolver, solver As CSectionSolver
    Dim eps As Double, kx As Double, ky As Double, relative As Double, prefix As String
    For kind = 1 To 4
        Set section = TestExperimentSection(kind, 2#)
        For loadCase = 1 To 4
            eps = -0.0008: kx = 0#: ky = 0#
            If loadCase = 2 Then eps = 0.001
            If loadCase = 3 Then eps = -0.0003: kx = 0.000008: ky = -0.000004
            If loadCase = 4 Then eps = -0.0005: kx = -0.000005: ky = 0.000007
            Set probe = New CSectionSolver: probe.EvaluateStrainPlane section, concrete, steel, eps, kx, ky
            Set baseline = TestNewSolver(False, 0#): baseline.Solve section, concrete, steel, probe.Nint, probe.Mxint, probe.Myint
            TestTrue "SignedBaseline_" & kind & "_" & loadCase, baseline.Converged
            For Each method In Array("Newton", "Secant")
                For mode = 0 To 1
                    For precision = 0 To 1
                        relative = 0#: If precision = 1 Then relative = 0.001
                        Set solver = TestNewSolver(mode = 1, relative): solver.SolverMethod = CStr(method)
                        solver.Solve section, concrete, steel, probe.Nint, probe.Mxint, probe.Myint
                        prefix = "Signed_" & kind & "_" & loadCase & "_" & method & "_" & mode & "_" & precision
                        TestTrue prefix & "_Converged", solver.Converged
                        If baseline.Converged And solver.Converged Then
                            TestClose prefix & "_eps", solver.Epsilon0, baseline.Epsilon0, 0.02, 0.00000001
                            TestStateField prefix, section, baseline, solver
                        Else
                            mReport = mReport & "SIGNED_DIAGNOSTIC: " & solver.StopReason & vbCrLf
                        End If
                    Next precision
                Next mode
            Next method
        Next loadCase
    Next kind
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' При почти нулевой кривизне относительная ошибка её коэффициента не
' определяет качество НДС. Контролирует максимальное изменение деформации
' во всех фактических бетонных элементах и стержнях, не больше 2% от
' максимальной исходной деформации, с абсолютным запасом 1e-10.
Private Sub TestStateField(ByVal prefix As String, ByVal section As CSectionModel, ByVal baseline As CSectionSolver, ByVal candidate As CSectionSolver)
    Dim i As Long, x As Double, y As Double, area As Double, referenceStrain As Double, candidateStrain As Double
    Dim maximumReference As Double, maximumDifference As Double
    For i = 1 To section.ConcreteCount + section.RebarCount
        If i <= section.ConcreteCount Then
            section.ConcreteElementForSolver i, x, y, area
        Else
            section.RebarElementForSolver i - section.ConcreteCount, x, y, area
        End If
        referenceStrain = baseline.StrainAt(x, y): candidateStrain = candidate.StrainAt(x, y)
        If Abs(referenceStrain) > maximumReference Then maximumReference = Abs(referenceStrain)
        If Abs(candidateStrain - referenceStrain) > maximumDifference Then maximumDifference = Abs(candidateStrain - referenceStrain)
    Next i
    TestTrue prefix & "_Field", maximumDifference <= 0.02 * maximumReference + 0.0000000001
    mReport = mReport & "DATA_MTLI_FIELD|" & prefix & "|" & TestNumber(maximumDifference) & "|" & TestNumber(maximumReference) & vbCrLf
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Осевые пределы сверяет с независимой формулой равномерного поля:
' N=Aс*sigmaс+As*(sigmas-sigmaс). Включает растяжение и сжатие коробки и круга.
Private Sub TestAxialLoadMultiplier()
    Dim kind As Long, mode As Long, sign As Long, section As CSectionModel
    Dim concrete As CMaterialDiagram, steel As CMaterialDiagram, cap As CCapacitySolver
    Set concrete = TestNonlinearConcrete(): Set steel = TestAsymmetricSteel()
    Dim expected As Double, area As Double, steelArea As Double, i As Long, strain As Double
    Dim mx As Double, my As Double, force As Double, x As Double, y As Double, elementArea As Double, lambda As Double
    For kind = 1 To 4
        Set section = TestExperimentSection(kind, 2#)
        area = 0#: steelArea = 0#
        For i = 1 To section.ConcreteCount: area = area + section.ConcreteArea(i): Next i
        For i = 1 To section.RebarCount: steelArea = steelArea + section.RebarArea(i): Next i
        For sign = -1 To 1 Step 2
            strain = -0.0035: If sign = 1 Then strain = 0.025
            expected = area * concrete.GetStress(strain) + steelArea * (steel.GetStress(strain) - concrete.GetStress(strain))
            mx = 0#: my = 0#
            If kind = 3 Then
                For i = 1 To section.ConcreteCount + section.RebarCount
                    If i <= section.ConcreteCount Then
                        section.ConcreteElementForSolver i, x, y, elementArea
                        force = elementArea * concrete.GetStress(strain)
                    Else
                        section.RebarElementForSolver i - section.ConcreteCount, x, y, elementArea
                        force = elementArea * (steel.GetStress(strain) - concrete.GetStress(strain))
                    End If
                    mx = mx + force * y: my = my + force * x
                Next i
            End If
            lambda = expected / (sign * 500000#)
            For mode = 0 To 1
                Set cap = TestNewCapacity(mode = 1, 1)
                cap.SolveByLoadPathMultiplier section, concrete, steel, 0#, sign * 500000#, 0#, mx / lambda, 0#, my / lambda
                TestTrue "AxialCap_" & kind & "_" & sign & "_" & mode, cap.Converged
                If cap.Converged Then TestClose "AxialExact_" & kind & "_" & sign & "_" & mode, cap.NUltimate, expected, 0.02, 1#
            Next mode
        Next sign
    Next kind
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Задаёт разную прочность и разные наклоны стали при растяжении и сжатии.
Private Function TestAsymmetricSteel() As CMaterialDiagram
    Dim e(1 To 4) As Double, s(1 To 4) As Double
    e(1) = -0.025: e(2) = -0.0015: e(3) = 0.00175: e(4) = 0.025
    s(1) = -320#: s(2) = -300#: s(3) = 350#: s(4) = 380#
    Set TestAsymmetricSteel = New CMaterialDiagram: TestAsymmetricSteel.InitializeFromArrays e, s, 4, -0.025, 0.025
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает несущую способность: постоянное N с масштабированием моментов
' и пропорциональное масштабирование N+Mx+My. Число итераций не является
' критерием приёмки; каждый конечный результат сравнивается с базовым.
Public Function RunThickLayerCapacityTests() As String
    On Error GoTo Failed
    mPassed = 0: mFailed = 0: mReport = vbNullString
    Dim concrete As CMaterialDiagram, steel As CMaterialDiagram
    Set concrete = TestNonlinearConcrete(): Set steel = TestNonlinearSteel()
    Dim kind As Long, mode As Long, precision As Long, path As Long
    Dim section As CSectionModel, baseline As CCapacitySolver, cap As CCapacitySolver
    Dim t0 As Double, relative As Double, noffset As Double, nbase As Double
    For kind = 1 To 3
        Set section = TestExperimentSection(kind, 5#)
        For path = 1 To 2
            noffset = -500000#: nbase = 0#
            If path = 2 Then noffset = 0#: nbase = -300000#
            Set baseline = TestNewCapacity(False, 0)
            baseline.SolveByUltimateLoadPath section, concrete, steel, noffset, nbase, 0#, 80000000#, 0#, 30000000#
            TestTrue "CapacityBaseline_" & kind & "_" & path, baseline.Converged
            For precision = 0 To 2
                For mode = 0 To 1
                    Set cap = TestNewCapacity(mode = 1, precision)
                    t0 = Timer
                    cap.SolveByUltimateLoadPath section, concrete, steel, noffset, nbase, 0#, 80000000#, 0#, 30000000#
                    mReport = mReport & "DATA_MTLI_CAPACITY|" & kind & "|" & path & "|" & mode & "|" & precision & "|" & _
                        TestNumber(Timer - t0) & "|" & cap.Iterations & "|" & cap.Converged & "|" & _
                        TestNumber(cap.LambdaUltimate) & "|" & TestNumber(cap.NUltimate) & "|" & _
                        TestNumber(cap.MxUltimate) & "|" & TestNumber(cap.MyUltimate) & "|" & cap.TestCriticalElement & "|" & cap.SolutionMethod & vbCrLf
                    TestTrue "CapacitySolve_" & kind & "_" & path & "_" & mode & "_" & precision, cap.Converged
                    If cap.Converged And baseline.Converged Then
                        TestClose "CapacityLambda_" & kind & "_" & path & "_" & mode & "_" & precision, cap.LambdaUltimate, baseline.LambdaUltimate, 0.02, 0.0000001
                        TestClose "CapacityN_" & kind & "_" & path & "_" & mode & "_" & precision, cap.NUltimate, baseline.NUltimate, 0.02, 1#
                        TestClose "CapacityMx_" & kind & "_" & path & "_" & mode & "_" & precision, cap.MxUltimate, baseline.MxUltimate, 0.02, 1000#
                        TestClose "CapacityMy_" & kind & "_" & path & "_" & mode & "_" & precision, cap.MyUltimate, baseline.MyUltimate, 0.02, 1000#
                    Else
                        mReport = mReport & "CAPACITY_DIAGNOSTIC: " & cap.StopReason & vbCrLf
                    End If
                Next mode
            Next precision
        Next path
    Next kind
    RunThickLayerCapacityTests = mReport & "TOTAL_MTLI_CAPACITY: passed=" & mPassed & "; failed=" & mFailed
    Exit Function
Failed:
    RunThickLayerCapacityTests = mReport & "RUNTIME ERROR: " & Err.Number & "; " & Err.Source & "; " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Базовая точность и два ослабленных набора заданы явно; физические пределы
' и диаграммы общие. Изменения точности не выдаются за изменение материала.
Private Function TestNewCapacity(ByVal mtli As Boolean, ByVal precision As Long) As CCapacitySolver
    Set TestNewCapacity = New CCapacitySolver
    With TestNewCapacity
        .UseThickLayerIntegration = mtli
        .ConcreteCompressionLimit = -0.0035: .ConcreteTensionLimitEnabled = False: .SteelStrainLimit = 0.025
        .SolverBaseLoadSteps = 1: .SolverMaxIterations = 80: .MaxRetries = 2
        .SolverToleranceN = 1#: .SolverToleranceMx = 1000#: .SolverToleranceMy = 1000#
        .StrainTolerance = 0.00001: .LambdaTolerance = 0.01
        .InitialLambdaStep = 1#: .MaxLambda = 64#: .DiagnosticsEnabled = False
        If precision = 1 Then .RelativeEquilibriumTolerance = 0.001
        If precision = 2 Then
            .RelativeEquilibriumTolerance = 0.01: .StrainTolerance = 0.000035: .LambdaTolerance = 0.02
        End If
    End With
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает реальный CSectionSolver на исходных дискретных нагрузках.
' Путь итераций свободен; проверяются конечные НДС и численные невязки.
' Записывает все варианты, включая неуспешные, до завершения набора.
Public Function RunThickLayerSolverTests() As String
    On Error GoTo Failed
    Dim kernelReport As String: kernelReport = RunThickLayerIntegrationTests()
    If InStr(kernelReport, "RUNTIME ERROR") > 0 Then RunThickLayerSolverTests = kernelReport: Exit Function
    Dim concrete As CMaterialDiagram, steel As CMaterialDiagram
    Set concrete = TestNonlinearConcrete(): Set steel = TestNonlinearSteel()
    Dim kind As Long, mode As Long, precision As Long, repeat As Long
    Dim section As CSectionModel, baseline As CSectionSolver, solver As CSectionSolver, probe As CSectionSolver
    Dim n As Double, mx As Double, my As Double, t0 As Double, relative As Double
    For kind = 1 To 3
        Set section = TestExperimentSection(kind, 2#)
        Set probe = New CSectionSolver
        probe.EvaluateStrainPlane section, concrete, steel, -0.0003, 0.000008, -0.000004
        n = probe.Nint: mx = probe.Mxint: my = probe.Myint
        Set baseline = TestNewSolver(False, 0#)
        baseline.Solve section, concrete, steel, n, mx, my
        TestTrue "Baseline_Converged_" & kind, baseline.Converged
        For repeat = 1 To 3
            For precision = 0 To 5
                Select Case precision
                    Case 0: relative = 0#
                    Case 1: relative = 0.000001
                    Case 2: relative = 0.0001
                    Case 3: relative = 0.001
                    Case 4: relative = 0.005
                    Case 5: relative = 0.01
                End Select
                For mode = 0 To 1
                    Set solver = TestNewSolver(mode = 1, relative)
                    t0 = Timer: solver.Solve section, concrete, steel, n, mx, my
                    mReport = mReport & "DATA_MTLI_SOLVE|" & kind & "|" & mode & "|" & TestNumber(relative) & "|" & repeat & "|" & _
                        TestNumber(Timer - t0) & "|" & solver.Iterations & "|" & solver.Converged & "|" & _
                        TestNumber(solver.Nint) & "|" & TestNumber(solver.Mxint) & "|" & TestNumber(solver.Myint) & "|" & _
                        TestNumber(solver.Epsilon0) & "|" & TestNumber(solver.KappaX) & "|" & TestNumber(solver.KappaY) & "|" & _
                        TestNumber(solver.ResidualN) & "|" & TestNumber(solver.ResidualMx) & "|" & TestNumber(solver.ResidualMy) & vbCrLf
                    TestTrue "Solve_" & kind & "_" & mode & "_" & precision & "_" & repeat, solver.Converged
                    If baseline.Converged And solver.Converged Then
                        TestClose "Eps_" & kind & "_" & mode & "_" & precision & "_" & repeat, solver.Epsilon0, baseline.Epsilon0, 0.02, 0.000000001
                        TestClose "Kx_" & kind & "_" & mode & "_" & precision & "_" & repeat, solver.KappaX, baseline.KappaX, 0.02, 0.00000000001
                        TestClose "Ky_" & kind & "_" & mode & "_" & precision & "_" & repeat, solver.KappaY, baseline.KappaY, 0.02, 0.00000000001
                    End If
                Next mode
            Next precision
        Next repeat
    Next kind
    RunThickLayerSolverTests = mReport & "TOTAL_MTLI_SOLVER: passed=" & mPassed & "; failed=" & mFailed
    Exit Function
Failed:
    RunThickLayerSolverTests = mReport & "RUNTIME ERROR: " & Err.Number & "; " & Err.Source & "; " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Настраивает только численный способ и точность; материалы и нагрузки общие.
Private Function TestNewSolver(ByVal mtli As Boolean, ByVal relative As Double) As CSectionSolver
    Set TestNewSolver = New CSectionSolver
    With TestNewSolver
        .UseThickLayerIntegration = mtli: .RelativeEquilibriumTolerance = relative
        .LoadSteps = 1: .MaxIterations = 80: .ToleranceN = 1#: .ToleranceMx = 1000#: .ToleranceMy = 1000#
        .DiagnosticsEnabled = False
    End With
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Создаёт прямоугольник, коробку или Г-сечение с одинаковым шагом точечной
' сетки. Контуры передаёт общее ядро. Арматура не размещается в отверстиях.
Private Function TestExperimentSection(ByVal kind As Long, ByVal stepSize As Double) As CSectionModel
    Dim result As CSectionModel: Set result = New CSectionModel: result.Clear
    Dim query As CSectionGeometryQuery: Set query = New CSectionGeometryQuery
    Dim region As CConcreteRegion, opening As CConcreteRegion
    Set region = query.RectangleRegion(0#, 0#, 200#, 400#)
    If kind = 2 Then
        Set opening = query.RectangleRegion(0#, 0#, 80#, 200#)
        Set region = query.SubtractContainedRegion(region, opening)
    ElseIf kind = 3 Then
        Set opening = query.RectangleRegion(50#, 100#, 100#, 200#)
        Set region = query.SubtractRegion(region, opening)
    ElseIf kind = 4 Then
        Set region = query.CircleRegion(0#, 0#, 100#)
    End If
    result.Contours.AddMaterialRegion region, "MTLITest", "Точная граница экспериментальной модели."
    Dim x As Double, y As Double, i As Long, included As Boolean
    For x = -100# + stepSize / 2# To 100# - stepSize / 2# Step stepSize
        For y = -200# + stepSize / 2# To 200# - stepSize / 2# Step stepSize
            included = True
            If kind = 2 Then included = Not (Abs(x) < 40# And Abs(y) < 100#)
            If kind = 3 Then included = Not (x > 0# And y > 0#)
            If kind = 4 Then included = (x * x + y * y < 10000#)
            If included Then result.AddConcreteElement x, y, stepSize * stepSize, width:=stepSize, height:=stepSize
        Next y
    Next x
    If kind = 4 Then
        For i = 0 To 11: result.AddRebarElement 80# * Cos(i * PI / 6#), 80# * Sin(i * PI / 6#), 20#, 0#, "A400": Next i
    Else
        For i = 0 To 5
            y = -180# + 72# * i
            result.AddRebarElement -80#, y, 20#, 0#, "A400"
            If kind <> 3 Or y < 0# Then result.AddRebarElement 80#, y, 20#, 0#, "A400"
        Next i
    End If
    Set TestExperimentSection = result
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Задаёт кусочно-линейный бетон с отключённым растяжением для сравнения интеграторов.
Private Function TestNonlinearConcrete() As CMaterialDiagram
    Dim e(1 To 5) As Double, s(1 To 5) As Double
    e(1) = -0.004: e(2) = -0.002: e(3) = -0.001: e(4) = 0#: e(5) = 0.1
    s(1) = -36#: s(2) = -30#: s(3) = -20#
    Set TestNonlinearConcrete = New CMaterialDiagram: TestNonlinearConcrete.InitializeFromArrays e, s, 5, -0.0035
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Задаёт симметричную сталь с ненулевым наклоном за пределом текучести.
Private Function TestNonlinearSteel() As CMaterialDiagram
    Dim e(1 To 4) As Double, s(1 To 4) As Double
    e(1) = -0.1: e(2) = -0.00175: e(3) = 0.00175: e(4) = 0.1
    s(1) = -369.65: s(2) = -350#: s(3) = 350#: s(4) = 369.65
    Set TestNonlinearSteel = New CMaterialDiagram: TestNonlinearSteel.InitializeFromArrays e, s, 4, -0.025, 0.025
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Выводит числа с точкой; разделитель полей журнала — вертикальная черта.
Private Function TestNumber(ByVal value As Double) As String
    TestNumber = Replace(CStr(value), ",", ".")
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Запускает адресные проверки нового интегратора без изменения листов.
Public Function RunThickLayerIntegrationTests() As String
    On Error GoTo Failed
    mPassed = 0: mFailed = 0: mReport = vbNullString
    Dim query As CSectionGeometryQuery: Set query = New CSectionGeometryQuery
    Dim region As CConcreteRegion, hole As CConcreteRegion
    Set region = query.RectangleRegion(0#, 0#, 200#, 400#)
    TestLinearRegion region, 80000#, 200# * 400# ^ 3 / 12#, 400# * 200# ^ 3 / 12#, "Rectangle"
    Set hole = query.RectangleRegion(0#, 0#, 100#, 200#)
    Set region = query.SubtractContainedRegion(region, hole)
    TestLinearRegion region, 60000#, (200# * 400# ^ 3 - 100# * 200# ^ 3) / 12#, _
        (400# * 200# ^ 3 - 200# * 100# ^ 3) / 12#, "HollowRectangle"
    Set region = query.CircleRegion(0#, 0#, 100#)
    TestLinearRegion region, PI * 10000#, PI * 100# ^ 4 / 4#, PI * 100# ^ 4 / 4#, "Circle"
    Set hole = query.CircleRegion(0#, 0#, 50#)
    Set region = query.SubtractContainedRegion(region, hole)
    TestLinearRegion region, PI * 7500#, PI * (100# ^ 4 - 50# ^ 4) / 4#, PI * (100# ^ 4 - 50# ^ 4) / 4#, "Ring"
    TestTranslationAndRotation query
    TestNarrowCompression query
    TestMaterialTransitions query
    RunThickLayerIntegrationTests = mReport & "TOTAL_MTLI_INTEGRATION: passed=" & mPassed & "; failed=" & mFailed
    Exit Function
Failed:
    RunThickLayerIntegrationTests = mReport & "RUNTIME ERROR: " & Err.Number & "; " & Err.Source & "; " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет N=EA*eps, M=EI*k и шесть независимых компонент K.
Private Sub TestLinearRegion(ByVal region As CConcreteRegion, ByVal area As Double, ByVal ix As Double, ByVal iy As Double, ByVal label As String)
    Dim material As CMaterialDiagram: Set material = TestLinearMaterial()
    Dim integrator As CThickLayerIntegrator: Set integrator = New CThickLayerIntegrator
    integrator.Initialize region, 0#, 0#
    Dim i As Long, eps As Double, kx As Double, ky As Double
    For i = 0 To 4
        eps = -0.0004 + i * 0.0002: kx = i * 0.0000005: ky = (2 - i) * 0.0000007
        If i = 0 Then ky = 0#
        integrator.Evaluate material, eps, kx, ky
        TestClose label & "_N_" & i, integrator.Force(1), 30000# * area * eps, 0.000000001, 0.00000001
        TestClose label & "_Mx_" & i, integrator.Force(2), 30000# * ix * kx, 0.000000001, 0.000001
        TestClose label & "_My_" & i, integrator.Force(3), 30000# * iy * ky, 0.000000001, 0.000001
        TestClose label & "_K11_" & i, integrator.Tangent(1, 1), 30000# * area, 0.000000001, 0.000001
        TestClose label & "_K22_" & i, integrator.Tangent(2, 2), 30000# * ix, 0.000000001, 0.001
        TestClose label & "_K33_" & i, integrator.Tangent(3, 3), 30000# * iy, 0.000000001, 0.001
        TestClose label & "_K12_" & i, integrator.Tangent(1, 2), 0#, 0#, 0.001
        TestClose label & "_K13_" & i, integrator.Tangent(1, 3), 0#, 0#, 0.001
        TestClose label & "_K23_" & i, integrator.Tangent(2, 3), 0#, 0#, 0.1
        TestTrue label & "_Points_" & i, integrator.EvaluationCount > 0
    Next i
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сопоставляет повёрнутую линейную область с аналитическим тензором инерции.
Private Sub TestTranslationAndRotation(ByVal query As CSectionGeometryQuery)
    Dim region As CConcreteRegion, integrator As CThickLayerIntegrator
    Set region = query.RectangleRegion(1000000#, -2000000#, 200#, 400#, 0.37)
    Set integrator = New CThickLayerIntegrator: integrator.Initialize region, 1000000#, -2000000#
    Dim material As CMaterialDiagram: Set material = TestLinearMaterial()
    Dim ix As Double, iy As Double, ixy As Double, a As Double, b As Double
    a = 200# * 400# ^ 3 / 12#: b = 400# * 200# ^ 3 / 12#
    ix = a * Cos(0.37) ^ 2 + b * Sin(0.37) ^ 2
    iy = b * Cos(0.37) ^ 2 + a * Sin(0.37) ^ 2
    ixy = (b - a) * Sin(0.37) * Cos(0.37)
    integrator.Evaluate material, -0.0002, 0.000001, -0.0000004
    TestClose "TranslatedRotated_N", integrator.Force(1), -480000#, 0.000000001, 0.000001
    TestClose "TranslatedRotated_Mx", integrator.Force(2), 30000# * (ix * 0.000001 - ixy * 0.0000004), 0.000000001, 0.001
    TestClose "TranslatedRotated_My", integrator.Force(3), 30000# * (ixy * 0.000001 - iy * 0.0000004), 0.000000001, 0.001
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Полоса толщиной 0,1 мм не должна исчезнуть между точками квадратуры.
Private Sub TestNarrowCompression(ByVal query As CSectionGeometryQuery)
    Dim material As CMaterialDiagram: Set material = TestCompressionMaterial()
    Dim region As CConcreteRegion: Set region = query.RectangleRegion(0#, 0#, 200#, 400#)
    Dim integrator As CThickLayerIntegrator: Set integrator = New CThickLayerIntegrator
    integrator.Initialize region, 0#, 0#: integrator.Evaluate material, 0.001999, 0.00001, 0#
    TestClose "NarrowCompression_N", integrator.Force(1), -0.3, 0.000000001, 0.000000001
    TestClose "NarrowCompression_Mx", integrator.Force(2), 59.99, 0.000000001, 0.00000001
    TestClose "NarrowCompression_MinEps", integrator.MinStrain, -0.000001, 0#, 0.00000000000001
    TestClose "NarrowCompression_MaxEps", integrator.MaxStrain, 0.003999, 0#, 0.00000000000001
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет общую производную R по конечным разностям на обеих сторонах
' диаграммы; у круга также сравнивает экстремумы с аналитическим радиусом.
Private Sub TestMaterialTransitions(ByVal query As CSectionGeometryQuery)
    Dim strains(1 To 7) As Double, stresses(1 To 7) As Double
    strains(1) = -0.004: strains(2) = -0.002: strains(3) = -0.001: strains(4) = 0#
    strains(5) = 0.00005: strains(6) = 0.0001: strains(7) = 0.004
    stresses(1) = -40#: stresses(2) = -30#: stresses(3) = -20#: stresses(4) = 0#
    stresses(5) = 2#: stresses(6) = 0#: stresses(7) = 0#
    Dim material As CMaterialDiagram: Set material = New CMaterialDiagram
    material.InitializeFromArrays strains, stresses, 7
    Dim integrator As CThickLayerIntegrator: Set integrator = New CThickLayerIntegrator
    Dim kind As Long, region As CConcreteRegion, p(1 To 3) As Double, h As Double, row As Long, col As Long
    Dim expected(1 To 3, 1 To 3) As Double, plus(1 To 3) As Double, minus(1 To 3) As Double
    For kind = 1 To 2
        If kind = 1 Then Set region = query.RectangleRegion(0#, 0#, 200#, 400#) Else Set region = query.CircleRegion(0#, 0#, 100#)
        integrator.Initialize region, 0#, 0#
        p(1) = -0.00031: p(2) = 0.000013: p(3) = -0.000011
        integrator.Evaluate material, p(1), p(2), p(3)
        For row = 1 To 3: For col = 1 To 3: expected(row, col) = integrator.Tangent(row, col): Next col: Next row
        If kind = 2 Then
            TestClose "Circle_MinEps", integrator.MinStrain, p(1) - 100# * Sqr(p(2) ^ 2 + p(3) ^ 2), 0#, 0.000000000001
            TestClose "Circle_MaxEps", integrator.MaxStrain, p(1) + 100# * Sqr(p(2) ^ 2 + p(3) ^ 2), 0#, 0.000000000001
        End If
        For col = 1 To 3
            h = 0.0000000001: If col > 1 Then h = 0.000000000001
            p(col) = p(col) + h: integrator.Evaluate material, p(1), p(2), p(3)
            For row = 1 To 3: plus(row) = integrator.Force(row): Next row
            p(col) = p(col) - 2# * h: integrator.Evaluate material, p(1), p(2), p(3)
            For row = 1 To 3: minus(row) = integrator.Force(row): Next row
            p(col) = p(col) + h
            For row = 1 To 3
                TestClose "Derivative_" & kind & "_" & row & "_" & col, (plus(row) - minus(row)) / (2# * h), expected(row, col), 0.00001, 1#
            Next row
        Next col
    Next kind
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Материал с одинаковым наклоном 30000 МПа на всём диапазоне проверки.
Private Function TestLinearMaterial() As CMaterialDiagram
    Dim strains(1 To 2) As Double, stresses(1 To 2) As Double
    strains(1) = -1#: strains(2) = 1#: stresses(1) = -30000#: stresses(2) = 30000#
    Set TestLinearMaterial = New CMaterialDiagram: TestLinearMaterial.InitializeFromArrays strains, stresses, 2
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Линейное сжатие при полном отключении сопротивления растяжению.
Private Function TestCompressionMaterial() As CMaterialDiagram
    Dim strains(1 To 3) As Double, stresses(1 To 3) As Double
    strains(1) = -1#: strains(2) = 0#: strains(3) = 1#: stresses(1) = -30000#
    Set TestCompressionMaterial = New CMaterialDiagram: TestCompressionMaterial.InitializeFromArrays strains, stresses, 3
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет численный результат по заранее выбранному абсолютному и относительному допускам.
Private Sub TestClose(ByVal label As String, ByVal actual As Double, ByVal expected As Double, ByVal relative As Double, ByVal absolute As Double)
    TestTrue label, Abs(actual - expected) <= absolute + relative * Abs(expected)
    If Abs(actual - expected) > absolute + relative * Abs(expected) Then mReport = mReport & "VALUES: actual=" & actual & "; expected=" & expected & vbCrLf
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Записывает каждую проверку, сохраняя все отказы до окончания набора.
Private Sub TestTrue(ByVal label As String, ByVal accepted As Boolean)
    If accepted Then mPassed = mPassed + 1 Else mFailed = mFailed + 1
    mReport = mReport & IIf(accepted, "PASS: ", "FAIL: ") & label & vbCrLf
End Sub
