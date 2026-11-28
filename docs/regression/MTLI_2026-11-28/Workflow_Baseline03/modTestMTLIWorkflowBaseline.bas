Attribute VB_Name = "modTestMTLIWorkflowBaseline"
Option Explicit

' ДЛЯ ТЕСТОВ: единая проба пользовательской команды на релизном VBA.

#If VBA7 Then
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef value As Currency) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef value As Currency) As Long
#Else
Private Declare Function QueryPerformanceCounter Lib "kernel32" (ByRef value As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (ByRef value As Currency) As Long
#End If

Private mPassed As Long, mFailed As Long, mReport As String

Public Function RunThickLayerWorkflowBaselineBenchmarkTests() As String
    RunThickLayerWorkflowBaselineBenchmarkTests = TestWorkflowSeries(False, True)
End Function

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

Private Sub TestWorkbookSetting(ByVal key As String, ByVal value As String)
    Dim table As Object: Set table = ThisWorkbook.Names("rngSystemSettings").RefersToRange
    Dim row As Long
    For row = 2 To table.Rows.Count
        If CStr(table.Cells(row, 1).Value2) = key Then table.Cells(row, 2).Value2 = value: Exit Sub
    Next row
    Err.Raise 5, "modTestThickLayerIntegration", "Не найдена настройка " & key
End Sub

Private Function TestWorkflowNow() As Double
    Dim counter As Currency, frequency As Currency
    If QueryPerformanceCounter(counter) = 0 Or QueryPerformanceFrequency(frequency) = 0 Then Err.Raise 5, , "Высокоточный таймер недоступен."
    TestWorkflowNow = counter / frequency
End Function

Private Function TestNumber(ByVal value As Double) As String
    TestNumber = Replace(CStr(value), ",", ".")
End Function

Private Sub TestTrue(ByVal label As String, ByVal accepted As Boolean)
    If accepted Then mPassed = mPassed + 1 Else mFailed = mFailed + 1
    mReport = mReport & IIf(accepted, "PASS: ", "FAIL: ") & label & vbCrLf
End Sub