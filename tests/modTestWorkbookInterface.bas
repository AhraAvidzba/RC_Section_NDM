Attribute VB_Name = "modTestWorkbookInterface"
Option Explicit

' ==========================================================================
' ДЛЯ ТЕСТОВ: структура книги и пользовательский интерфейс
' ==========================================================================
' Проверяет именованные диапазоны, выпадающие списки, листы и макросы
' для воспроизводимой сборки книги.

#If VBA7 Then
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Связывает полученный COM HWND с PID, чтобы не трогать прежний пользовательский Excel.
Private Declare PtrSafe Function TestExcelWindowProcessId Lib "user32" Alias "GetWindowThreadProcessId" (ByVal hwnd As LongPtr, ByRef processID As Long) As Long
#Else
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Тот же Win32-контракт для 32-разрядного VBA без LongPtr.
Private Declare Function TestExcelWindowProcessId Lib "user32" Alias "GetWindowThreadProcessId" (ByVal hwnd As Long, ByRef processID As Long) As Long
#End If

Private Type TUiTestStats
    Passed As Long
    Failed As Long
    Report As String
End Type

Private Const TEST_PLOT_FRAME_LEFT As Double = 28#
Private Const TEST_PLOT_FRAME_TOP As Double = 36#
Private Const TEST_PLOT_FRAME_WIDTH_MARGIN As Double = 180#
Private Const TEST_PLOT_FRAME_HEIGHT_MARGIN As Double = 72#
Private mAudit03StopBeforeLoadTable As Boolean ' Только диагностический прогон сохраняет книгу на границе крупных тестовых блоков.
Private mAudit03StopBeforeUnitSign As Boolean ' Диагностический снимок входных тестов отделяет состояние книги от накопления ресурсов Excel.

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Выполняет COM-регрессии структуры Config/Results, workbook-макросов и схемы.
' Проверка снимков и export-data не является приемкой фактической записи DWG.
Public Function RunWorkbookInterfaceTests() As String
    On Error GoTo Failed

    Dim stats As TUiTestStats
    Dim t0 As Double
    Dim suiteSheet As Object
    Set suiteSheet = ThisWorkbook.Application.ActiveSheet
    t0 = Timer

    TestAppendLine stats, "RUN: TestButtons"
    TestButtons stats
    TestLoadCombinationsOnConfig stats
    TestAppendLine stats, "RUN: TestSingleCombinationSkipsBlankRows"
    TestSingleCombinationSkipsBlankRows stats
    TestLoadCombinationRangeMinimumRows stats
    TestConfigDropdownChoices stats
    TestConfigDropdownMeaning stats
    TestConfigSettingsRightBorder stats
    TestAutoCADGeometryContourWarning stats
    TestAppendLine stats, "RUN: TestAudit03ReaderContract"
    TestAudit03ReaderContract stats
    TestAppendLine stats, "RUN: TestAudit03InputContracts; " & TestAudit02ExcelMemory()
    TestAudit03InputContracts stats
    TestAppendLine stats, "RUN: TestAudit03ProfileInputContracts; " & TestAudit02ExcelMemory()
    TestAudit03ProfileInputContracts stats
    TestAppendLine stats, "RUN: TestAudit03NumericSettingsInputContracts; " & TestAudit02ExcelMemory()
    TestAudit03NumericSettingsInputContracts stats
    TestAppendLine stats, "RUN: TestAudit03SettingErrorMessages; " & TestAudit02ExcelMemory()
    TestAudit03SettingErrorMessages stats
    TestAppendLine stats, "RUN: TestAudit03RequiredTableMessages; " & TestAudit02ExcelMemory()
    TestAudit03RequiredTableMessages stats
    Dim settingsTablePassed As Long, settingsTableFailed As Long
    stats.Report = stats.Report & modTestConfiguration.RunAudit03SettingsTableGuardTests(settingsTablePassed, settingsTableFailed)
    stats.Passed = stats.Passed + settingsTablePassed
    stats.Failed = stats.Failed + settingsTableFailed
    Dim multiAreaPassed As Long, multiAreaFailed As Long
    stats.Report = stats.Report & modTestConfiguration.RunAudit03MultiAreaInputTests(multiAreaPassed, multiAreaFailed)
    stats.Passed = stats.Passed + multiAreaPassed
    stats.Failed = stats.Failed + multiAreaFailed
    Dim generalRunPassed As Long, generalRunFailed As Long
    stats.Report = stats.Report & modTestConfiguration.RunAudit03GeneralRunConfigTests(generalRunPassed, generalRunFailed)
    stats.Passed = stats.Passed + generalRunPassed
    stats.Failed = stats.Failed + generalRunFailed
    TestAppendLine stats, "RUN: TestAudit03ConfigConversionMessages; " & TestAudit02ExcelMemory()
    TestAudit03ConfigConversionMessages stats
    TestAppendLine stats, "RUN: TestAudit03UnitSignConsumers; " & TestAudit02ExcelMemory()
    If mAudit03StopBeforeUnitSign Then
        TestAppendLine stats, "TOTAL_AUDIT03_UI_INPUT_PREFIX: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
        RunWorkbookInterfaceTests = stats.Report
        Exit Function
    End If
    ' Массовые Config-прогоны не требуют отображения постоянно меняющегося
    ' Results. Активный лист suite возвращаем перед окончательным отчетом.
    ThisWorkbook.Worksheets.Item("Config").Activate
    TestAudit03UnitSignConsumers stats
    TestAppendLine stats, "RUN: TestAudit03UnitSignChoices; " & TestAudit02ExcelMemory()
    TestAudit03UnitSignChoices stats
    TestAppendLine stats, "RUN: TestAudit03UnitSignEquivalence; " & TestAudit02ExcelMemory()
    TestAudit03UnitSignEquivalence stats
    TestAppendLine stats, "RUN: TestAudit03InputUnitConsumers; " & TestAudit02ExcelMemory()
    TestAudit03InputUnitConsumers stats
    TestAppendLine stats, "RUN: Audit03 Rectangles selectors; " & TestAudit02ExcelMemory()
    TestRectanglesIndependentSelectors stats
    TestRectanglesIndependentSelectorLayout stats
    TestRectanglesIndependentSelectorEffects stats
    TestRectanglesIndependentThirdRows stats
    Dim circleConfigPassed As Long, circleConfigFailed As Long
    TestAppendLine stats, "RUN: Audit03 Circle Config; " & TestAudit02ExcelMemory()
    stats.Report = stats.Report & modTestGeometryConfig.RunAudit03CircleConfigTests(circleConfigPassed, circleConfigFailed)
    stats.Passed = stats.Passed + circleConfigPassed
    stats.Failed = stats.Failed + circleConfigFailed
    Dim circleInputPassed As Long, circleInputFailed As Long
    TestAppendLine stats, "RUN: Audit03 Circle rebar input; " & TestAudit02ExcelMemory()
    stats.Report = stats.Report & modTestGeometryConfig.RunAudit03CircleRebarInputTests(circleInputPassed, circleInputFailed)
    stats.Passed = stats.Passed + circleInputPassed
    stats.Failed = stats.Failed + circleInputFailed
    Dim shapeConfigPassed As Long, shapeConfigFailed As Long
    TestAppendLine stats, "RUN: Audit03 Shape Config; " & TestAudit02ExcelMemory()
    stats.Report = stats.Report & modTestGeometryConfig.RunAudit03ShapeConfigTests(shapeConfigPassed, shapeConfigFailed)
    stats.Passed = stats.Passed + shapeConfigPassed
    stats.Failed = stats.Failed + shapeConfigFailed
    Dim rebarCounterPassed As Long, rebarCounterFailed As Long
    TestAppendLine stats, "RUN: Audit03 Rebar counters; " & TestAudit02ExcelMemory()
    stats.Report = stats.Report & modTestGeometryConfig.RunAudit03RebarCounterTests(rebarCounterPassed, rebarCounterFailed)
    stats.Passed = stats.Passed + rebarCounterPassed
    stats.Failed = stats.Failed + rebarCounterFailed
    Dim requiredGeometryPassed As Long, requiredGeometryFailed As Long
    TestAppendLine stats, "RUN: Audit03 Required geometry; " & TestAudit02ExcelMemory()
    stats.Report = stats.Report & modTestGeometryConfig.RunAudit03RequiredGeometryInputTests(requiredGeometryPassed, requiredGeometryFailed)
    stats.Passed = stats.Passed + requiredGeometryPassed
    stats.Failed = stats.Failed + requiredGeometryFailed
    Dim rebarFieldPassed As Long, rebarFieldFailed As Long
    TestAppendLine stats, "RUN: Audit03 Rebar fields; " & TestAudit02ExcelMemory()
    stats.Report = stats.Report & modTestGeometryConfig.RunAudit03RebarFieldConfigTests(rebarFieldPassed, rebarFieldFailed)
    stats.Passed = stats.Passed + rebarFieldPassed
    stats.Failed = stats.Failed + rebarFieldFailed
    If mAudit03StopBeforeLoadTable Then
        TestAppendLine stats, "TOTAL_AUDIT03_UI_PRE_LOAD_TABLE: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
        RunWorkbookInterfaceTests = stats.Report
        Exit Function
    End If
    Dim loadTablePassed As Long, loadTableFailed As Long
    TestAppendLine stats, "RUN: Audit03 Load table; " & TestAudit02ExcelMemory()
    stats.Report = stats.Report & modTestLoadTableConfig.RunAudit03LoadTableTests(loadTablePassed, loadTableFailed)
    TestAppendLine stats, "RUN: Audit03 Load table completed; " & TestAudit02ExcelMemory()
    stats.Passed = stats.Passed + loadTablePassed
    stats.Failed = stats.Failed + loadTableFailed
    Dim profileScopePassed As Long, profileScopeFailed As Long
    TestAppendLine stats, "RUN: RunAudit03ProfileFailureScopeTests"
    stats.Report = stats.Report & modTestLoadTableConfig.RunAudit03ProfileFailureScopeTests(profileScopePassed, profileScopeFailed)
    stats.Passed = stats.Passed + profileScopePassed
    stats.Failed = stats.Failed + profileScopeFailed
    Dim resultWidthsPassed As Long, resultWidthsFailed As Long
    stats.Report = stats.Report & modTestLoadTableConfig.RunAudit03ResultColumnWidthTests(resultWidthsPassed, resultWidthsFailed)
    stats.Passed = stats.Passed + resultWidthsPassed
    stats.Failed = stats.Failed + resultWidthsFailed
    Dim profileConfigPassed As Long, profileConfigFailed As Long
    TestAppendLine stats, "RUN: RunAudit03ProfileConfigTests"
    stats.Report = stats.Report & modTestProfileConfig.RunAudit03ProfileConfigTests(profileConfigPassed, profileConfigFailed)
    stats.Passed = stats.Passed + profileConfigPassed
    stats.Failed = stats.Failed + profileConfigFailed
    Dim relocatedPassed As Long, relocatedFailed As Long
    TestAppendLine stats, "RUN: RunAudit03RelocatedInputTests"
    stats.Report = stats.Report & modTestProfileConfig.RunAudit03RelocatedInputTests(relocatedPassed, relocatedFailed)
    stats.Passed = stats.Passed + relocatedPassed
    stats.Failed = stats.Failed + relocatedFailed
    Dim durationPassed As Long, durationFailed As Long
    TestAppendLine stats, "RUN: RunAudit03DurationConfigTests"
    stats.Report = stats.Report & modTestProfileConfig.RunAudit03DurationConfigTests(durationPassed, durationFailed)
    stats.Passed = stats.Passed + durationPassed
    stats.Failed = stats.Failed + durationFailed
    Dim presentationPassed As Long, presentationFailed As Long
    TestAppendLine stats, "RUN: RunAudit03PresentationConfigTests"
    stats.Report = stats.Report & modTestProfileConfig.RunAudit03PresentationConfigTests(presentationPassed, presentationFailed)
    stats.Passed = stats.Passed + presentationPassed
    stats.Failed = stats.Failed + presentationFailed
    TestAppendLine stats, "RUN: TestPartialCombinationIsInvalid"
    TestPartialCombinationIsInvalid stats
    TestInvalidProfileIdDoesNotRunPlot stats
    TestCalculationRequiresFormedSection stats
    TestResultsFormatting stats
    TestFormButtonBuildsGeneratedSection stats
    TestAutoCADImporterTreatsDrawingUnitsAsMillimeters stats
    TestAudit03InputAreaImportFilter stats
    TestAudit03ImportedSnapshotUnitChanges stats
    TestAudit03GeometrySnapshotContract stats
    TestAudit03ReadLifecycleContracts stats
    Dim plotConfigPassed As Long, plotConfigFailed As Long
    stats.Report = stats.Report & modTestPlotConfig.RunAudit03GeneralPlotTests(plotConfigPassed, plotConfigFailed)
    stats.Passed = stats.Passed + plotConfigPassed
    stats.Failed = stats.Failed + plotConfigFailed
    TestAppendLine stats, "RUN: TestAutoCADPreviewWritesAndDrawsBoundsDimensions"
    TestAutoCADPreviewWritesAndDrawsBoundsDimensions stats
    TestAnnotationDimensionTextRoundsInMillimeters stats
    TestPlotOverlayCoordinatesMatchResults stats
    TestAppendLine stats, "RUN: TestPlotUsesSavedGeometryAfterSourceChange"
    TestPlotUsesSavedGeometryAfterSourceChange stats
    TestAppendLine stats, "RUN: TestGeneratedDirectStateWorstStillDrawsFirstCalculatedLC"
    TestGeneratedDirectStateWorstStillDrawsFirstCalculatedLC stats
    TestAppendLine stats, "RUN: TestProfileDrivenPlotUsesSnapshotState"
    TestProfileDrivenPlotUsesSnapshotState stats
    TestAudit02SavedResultsIgnoreMaterialChanges stats
    TestAppendLine stats, "RUN: TestMissingProfileStateDrawsGeometryOnly"
    TestMissingProfileStateDrawsGeometryOnly stats
    TestMissingProfileStateDrawsGeometryOnly stats, "File"
    TestAutoCADCalculationMessageUsesSavedGeometry stats
    TestBlankMomentDefaultsToZeroAndZeroLoadsAreSkipped stats
    TestAppendLine stats, "RUN: TestCircleWorkbookRunWritesResults"
    TestCircleWorkbookRunWritesResults stats
    TestAppendLine stats, "RUN: TestThirtyCombinationsWithFiveStatesWriteSnapshot"
    TestThirtyCombinationsWithFiveStatesWriteSnapshot stats
    TestDynamicLoadCombinationRangeAndLayoutGuard stats
    TestAppendLine stats, "RUN: TestExecutionReportFile"
    TestExecutionReportFile stats
    TestExcelApplicationStateGuardRestoresSettings stats
    TestAppendLine stats, "RUN: TestRectanglesWorkbookRunWritesResults"
    TestRectanglesWorkbookRunWritesResults stats
    TestAppendLine stats, "RUN: TestRectanglesMomentUltimateStrainWorkbookPath"
    TestRectanglesMomentUltimateStrainWorkbookPath stats
    TestStrengthSummaryUsesOutputCurvatureUnit stats
    TestAppendLine stats, "RUN: TestRectanglesPureBendingUltimateStrainWorkbookPath"
    TestRectanglesPureBendingUltimateStrainWorkbookPath stats
    TestAppendLine stats, "RUN: TestRectanglesPureBendingDirectStateWorkbookPath"
    TestRectanglesPureBendingDirectStateWorkbookPath stats
    TestAppendLine stats, "RUN: TestRectanglesAxialTensionExtensionFromWorkbookSettings"
    TestRectanglesAxialTensionExtensionFromWorkbookSettings stats
    TestAutoCADExportUsesSharedLoadReference stats
    TestAppendLine stats, "RUN: TestGoverningCombinationWritesDetailedResults"
    TestGoverningCombinationWritesDetailedResults stats
    TestCapacitySearchMethodValidation stats
    TestSolverToleranceUnitLabels stats
    TestCapacitySettingsUnitLabels stats
    TestRebarInputValidationDoesNotUseHiddenDefaults stats
    TestClearResultsKeepsInputs stats

    ' Настройки CAD проверяются без подключения к приложению. Настоящий DWG
    ' остается отдельным явным тестом с guard-ом собственных документов.
    Dim autoCADPassed As Long, autoCADFailed As Long
    TestAppendLine stats, "RUN: RunAudit03AutoCADConfigTests"
    stats.Report = stats.Report & modTestAutoCADConfig.RunAudit03AutoCADConfigTests(autoCADPassed, autoCADFailed)
    stats.Passed = stats.Passed + autoCADPassed
    stats.Failed = stats.Failed + autoCADFailed

    suiteSheet.Activate
    TestAssertTrue stats, "ui.suite.activeSheetRestored", ThisWorkbook.Application.ActiveSheet Is suiteSheet
    TestAppendLine stats, "TOTAL_WORKBOOK_UI: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed) & _
        "; elapsedSec=" & TestFormatNumberInvariant(Timer - t0)
    RunWorkbookInterfaceTests = stats.Report
    Exit Function

Failed:
    Dim failureNumber As Long, failureSource As String, failureDescription As String
    failureNumber = Err.Number: failureSource = Err.Source: failureDescription = Err.Description
    On Error Resume Next
    If Not suiteSheet Is Nothing Then suiteSheet.Activate
    On Error GoTo 0
    RunWorkbookInterfaceTests = stats.Report & "RUNTIME ERROR: " & CStr(failureNumber) & _
        "; source=" & failureSource & "; description=" & failureDescription
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' ДЛЯ ТЕСТОВ
' Запускает тяжелый ручной сценарий: 20 LC, пять named-state на LC и мелкая
' сетка Г-сечения. Он не входит в обычный RunWorkbookInterfaceTests, потому что
' нужен только для проверки больших snapshot-ов и лимита Excel Chart на series.
Public Function RunLargeSnapshotPlotStressTest() As String
    On Error GoTo Failed

    Dim stats As TUiTestStats
    Dim t0 As Double
    t0 = Timer

    TestAudit03ExcelGuardContracts stats
    TestLargeSnapshotPlotStress stats

    TestAppendLine stats, "TOTAL_LARGE_SNAPSHOT_PLOT: passed=" & CStr(stats.Passed) & _
        "; failed=" & CStr(stats.Failed) & "; elapsedSec=" & TestFormatNumberInvariant(Timer - t0)
    RunLargeSnapshotPlotStressTest = stats.Report
    Exit Function

Failed:
    RunLargeSnapshotPlotStressTest = stats.Report & "RUNTIME ERROR: " & CStr(Err.Number) & _
        "; source=" & Err.Source & "; description=" & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет, что таблица сочетаний живет на Config рядом с настройками,
' и не выводится отдельной параллельной таблицей на лист Расчет.
Private Sub TestLoadCombinationsOnConfig(ByRef stats As TUiTestStats)
    Dim loads As Object
    Set loads = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange

    TestAssertTrue stats, "ui.loads.layout.sheet", loads.Worksheet.Name = "Config"
    TestAssertTrue stats, "ui.loads.layout.titleRowExists", loads.Row > 1
    TestAssertTrue stats, "ui.loads.layout.defaultRows30", loads.Rows.Count = 31
    TestAssertTrue stats, "ui.loads.layout.title", _
        InStr(1, CStr(loads.Worksheet.Cells.Item(loads.Row - 1, loads.Column).Value2), "Сочетания нагрузок", vbTextCompare) > 0
    Dim pathOptions As Variant
    pathOptions = Array("Auto", ChrW$(&H3BB) & "*Mx", ChrW$(&H3BB) & "*My", _
        ChrW$(&H3BB) & "*Mxy", ChrW$(&H3BB) & "*N", ChrW$(&H3BB) & "*NMxy")
    TestAssertTrue stats, "ui.loads.layout.capacityPathValidation", _
        TestLoadCombinationValidationHasOptions(6, pathOptions)
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет наличие пяти кнопок, привязанные OnAction-макросы и размещение
' вне области печати; надписи должны быть центрированы в самих shapes.
Private Sub TestButtons(ByRef stats As TUiTestStats)
    Dim calc As Object
    Set calc = ThisWorkbook.Worksheets.Item("Расчет")

    Dim runButton As Object
    Dim clearButton As Object
    Dim importButton As Object
    Dim acadButton As Object
    Dim plotButton As Object
    Set runButton = calc.Shapes.Item("btnRunSectionCalculation")
    Set importButton = calc.Shapes.Item("btnImportGeometryFromAutoCAD")
    Set clearButton = calc.Shapes.Item("btnClearAutoCADDrawing")
    Set acadButton = calc.Shapes.Item("btnExportStressToAutoCAD")
    Set plotButton = calc.Shapes.Item("btnUpdateSectionPlot")

    TestAssertTrue stats, "ui.buttons.run.exists", Not runButton Is Nothing
    TestAssertTrue stats, "ui.buttons.import.exists", Not importButton Is Nothing
    TestAssertTrue stats, "ui.buttons.clear.exists", Not clearButton Is Nothing
    TestAssertTrue stats, "ui.buttons.autocad.exists", Not acadButton Is Nothing
    TestAssertTrue stats, "ui.buttons.plot.exists", Not plotButton Is Nothing
    TestAssertTrue stats, "ui.buttons.order", importButton.Top < runButton.Top And runButton.Top < plotButton.Top And _
        plotButton.Top < acadButton.Top And acadButton.Top < clearButton.Top
    TestAssertTrue stats, "ui.buttons.form.caption", importButton.TextFrame.Characters.Text = "Сформировать сечение"
    TestAssertTrue stats, "ui.buttons.run.macro", InStr(1, runButton.OnAction, "RunSectionCalculation", vbTextCompare) > 0
    TestAssertTrue stats, "ui.buttons.import.macro", InStr(1, importButton.OnAction, "FormSection", vbTextCompare) > 0
    TestAssertTrue stats, "ui.buttons.clear.macro", InStr(1, clearButton.OnAction, "ClearAutoCADDrawing", vbTextCompare) > 0
    TestAssertTrue stats, "ui.buttons.autocad.macro", InStr(1, acadButton.OnAction, "ExportSectionStressToAutoCAD", vbTextCompare) > 0
    TestAssertTrue stats, "ui.buttons.plot.macro", InStr(1, plotButton.OnAction, "UpdateSectionPlot", vbTextCompare) > 0
    TestAssertTrue stats, "ui.buttons.outsidePrint", runButton.Left > calc.Range("AJ1").Left And importButton.Left > calc.Range("AJ1").Left And clearButton.Left > calc.Range("AJ1").Left And acadButton.Left > calc.Range("AJ1").Left And plotButton.Left > calc.Range("AJ1").Left
    TestAssertTrue stats, "ui.buttons.textCentered", TestButtonTextIsCentered(runButton) And TestButtonTextIsCentered(importButton) And _
        TestButtonTextIsCentered(clearButton) And TestButtonTextIsCentered(acadButton) And TestButtonTextIsCentered(plotButton)
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет оба выравнивания надписи кнопки; размер/положение фигуры этим не оцениваются.
Private Function TestButtonTextIsCentered(ByVal buttonShape As Object) As Boolean
    TestButtonTextIsCentered = (buttonShape.TextFrame.HorizontalAlignment = -4108 And _
        buttonShape.TextFrame.VerticalAlignment = -4108)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет пользовательское правило: количество стержней не подставляется из
' кода, а локальное n=0 или d=0 просто отключает соответствующую линию.
Private Sub TestRebarInputValidationDoesNotUseHiddenDefaults(ByRef stats As TUiTestStats)
    TestPrepareCircleInput
    TestSetSystemSetting "Rebar.Diameter", vbNullString

    Dim errorText As String
    errorText = TestBuildCurrentWorkbookSectionError()
    TestAssertTrue stats, "ui.circle.blankDiameter.noHiddenDefaults", _
        InStr(1, errorText, "не получен ни один стержень", vbTextCompare) > 0

    TestPrepareRectanglesInput
    TestSetSystemSetting "Rectangles.H1.d_1", vbNullString

    Dim rectanglesSection As CSectionModel
    Set rectanglesSection = TestBuildCurrentWorkbookSection()
    TestAssertTrue stats, "ui.rectangles.blankDiameterWithCount.skipsLine", rectanglesSection.RebarCount = 7
    TestAssertTrue stats, "ui.rectangles.blankDiameterWithCount.dimensionsRemain", rectanglesSection.AnnotationCount >= 4

    TestPrepareRectanglesInput
    TestSetSystemSetting "Rectangles.H1.n_1", "0"
    TestSetSystemSetting "Rectangles.H1.d_1", vbNullString

    Set rectanglesSection = TestBuildCurrentWorkbookSection()
    TestAssertTrue stats, "ui.rectangles.zeroCountBlankDiameter.skipsLine", rectanglesSection.RebarCount = 7
    TestAssertTrue stats, "ui.rectangles.zeroCountBlankDiameter.dimensionsRemain", rectanglesSection.AnnotationCount >= 4

    TestSetSystemSetting "Geometry.Type", "PointedSection"
    TestSetSystemSetting "PointedSection.H.Left.n", vbNullString
    TestSetSystemSetting "PointedSection.H.Right.n", vbNullString
    TestSetSystemSetting "PointedSection.B.Top.n", vbNullString
    TestSetSystemSetting "PointedSection.B.Bottom.n", vbNullString

    errorText = TestBuildCurrentWorkbookSectionError()
    TestAssertTrue stats, "ui.rounded.blankCounts.noHiddenDefaults", _
        InStr(1, errorText, "не получен ни один стержень", vbTextCompare) > 0

    TestPrepareCircleInput
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' заданные имена опциональных контуров отсутствуют в fixture;
' импорт расчетной сетки должен продолжать работать без дополнительных слоев.
Private Function TestAudit03AbsentContourLayers() As Collection
    Set TestAudit03AbsentContourLayers = New Collection
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Собирает модель через тот же путь, которым пользуется кнопка расчета:
' Config -> CSystemSettingsReader -> CUnitSystem -> CSectionTypeRegistry.
Private Function TestBuildCurrentWorkbookSection() As CSectionModel
    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook

    Dim units As CUnitSystem
    Set units = New CUnitSystem
    units.LoadFromSettings settings

    Set TestBuildCurrentWorkbookSection = BuildWorkbookSectionModel(ThisWorkbook, settings, units)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает текст ошибки при сборке модели или пустую строку, если модель
' построилась. Нужна для тестов пользовательской валидации Config без остановки
' всего набора workbook-тестов.
Private Function TestBuildCurrentWorkbookSectionError() As String
    On Error GoTo GotError
    Dim section As CSectionModel
    Set section = TestBuildCurrentWorkbookSection()
    TestBuildCurrentWorkbookSectionError = vbNullString
    Exit Function

GotError:
    TestBuildCurrentWorkbookSectionError = Err.Description
    Err.Clear
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет, что включенный общий флаг создает человекочитаемый txt-отчет
' рядом с книгой и не мешает обычному расчетному сценарию.
Private Sub TestExecutionReportFile(ByRef stats As TUiTestStats)
    TestPrepareCircleInput
    TestSetSystemSetting "General.ExecutionReportEnabled", "Yes"

    Dim reportPath As String
    reportPath = ThisWorkbook.Path & "\RC_Section_NDM_execution_report.txt"
    TestDeleteFileIfExists reportPath

    Dim message As String
    message = TestFixtureRunSectionWorkflow(ThisWorkbook, False)

    TestAssertTrue stats, "ui.executionReport.message", InStr(1, message, "Пошаговый отчет сохранен", vbTextCompare) > 0
    TestAssertTrue stats, "ui.executionReport.fileExists", TestFileExists(reportPath)
    Dim reportText As String
    reportText = TestReadTextFile(reportPath)
    TestAssertTrue stats, "ui.executionReport.content", InStr(1, reportText, "ОТЧЕТ ВЫПОЛНЕНИЯ RC SECTION NDM", vbTextCompare) > 0 And _
        InStr(1, reportText, "Старт расчета", vbTextCompare) > 0 And _
        InStr(1, reportText, "Итерации прямого CSectionSolver", vbTextCompare) > 0 And _
        InStr(1, reportText, "Финальное сообщение", vbTextCompare) > 0

    TestSetSystemSetting "General.ExecutionReportEnabled", "No"
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' проверяет точный текст трех видов диагностики через публичный отчет,
' без решения НДС. Пробелы, регистр и специальные целые строки проверяются строго.
Public Function RunExecutionReportFormattingTests() As String
    Dim stats As TUiTestStats, settings As CSystemSettingsReader, report As CExecutionReport
    Dim oldEnabled As Boolean, captured As Boolean
    On Error GoTo Failed
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    oldEnabled = settings.GetRequiredBoolean("General.ExecutionReportEnabled"): captured = True
    TestSetSystemSetting "General.ExecutionReportEnabled", "Yes"
    settings.LoadFromWorkbook ThisWorkbook
    Set report = New CExecutionReport: report.Initialize ThisWorkbook, settings

    Dim axial As String, axialRu As String, redirected As String
    axial = "ultimatePath pureAxial redirected=LoadMultiplier"
    axialRu = "Силовая траектория UltimateStrain перенаправлена на LoadMultiplier."
    redirected = "Силовая траектория lambda*N перенаправлена на LoadMultiplier, потому что UltimateStrain для нее вырождается."
    Dim bodies As Variant, expected As Variant, titles As Variant
    bodies = Array(axial, axialRu, axial & vbCrLf & axialRu, "prefix " & axial, axial & " suffix", _
        axial & vbLf & axial, UCase$(axial), " " & vbCrLf & "  " & axial & "  " & vbCrLf & " ", _
        "lambda_crc_path_iter=2; iter=3", "rN=0; converged=True; converged=False", _
        "Кириллица: " & ChrW(&H3BB) & "; R2; мм" & ChrW(&HB2), _
        "searchMethod=Brent; method=Bisection; iterations=3; probe=1; residualMx=0; residualMy=0", _
        axial & vbCrLf & "prefix " & axial & vbCrLf & axial, _
        "initialGuess=warmStartSolver", "initialGuess=directStateActiveSet", "   ", _
        "ultimateTrustRegion attempt=2; rho=0.5", _
        "ultimateTrustRegion flat trial rejected; original recovery remains available")
    expected = Array(redirected, redirected, redirected & vbCrLf & redirected, "prefix " & axial, axial & " suffix", _
        axial & vbLf & axial, UCase$(axial), redirected, _
        "итерация UltimateStrain по lambda-траектории lambda_crc=2; итерация=3", _
        "невязка N=0; сошлось=да; сошлось=нет", "Кириллица: " & ChrW(&H3BB) & "; R2; мм" & ChrW(&HB2), _
        "метод поиска=Brent; метод=Bisection; итераций=3; точка lambda=1; невязка Mx=0; невязка My=0", _
        redirected & vbCrLf & "prefix " & axial & vbCrLf & redirected, _
        "initialGuess=предыдущее близкое решение", "initialGuess=активная зона прямого НДС", vbNullString, _
        "Резервный шаг на активном пределе: проба=2; соответствие прогноза=0.5", _
        "Резервная проба на плато отклонена; прежний подбор шага сохранен.")
    titles = Array("Диагностика расчета трещин", "Сценарий прямого StateSolution", "Диагностика поиска несущей способности")
    Dim kind As Long, i As Long, title As String, body As String
    For kind = 0 To UBound(titles)
        For i = 0 To UBound(bodies)
            title = CStr(titles(kind)) & " FORMAT_CASE_" & CStr(kind) & "_" & Format$(i, "00")
            body = CStr(bodies(i))
            report.AddDiagnosticBlock title, body
            TestAssertTrue stats, "report.format.inputUnchanged." & CStr(kind) & "." & CStr(i), _
                StrComp(body, CStr(bodies(i)), vbBinaryCompare) = 0
            report.AddStep "FORMAT_END_" & CStr(kind) & "_" & Format$(i, "00")
        Next i
    Next kind
    report.Save "FORMAT_COMPLETE"

    Dim text As String, startAt As Long, endAt As Long, actual As String, wanted As String
    text = TestReadTextFile(report.FilePath)
    For kind = 0 To UBound(titles)
        For i = 0 To UBound(bodies)
            title = CStr(titles(kind)) & " FORMAT_CASE_" & CStr(kind) & "_" & Format$(i, "00")
            startAt = InStr(1, text, title, vbBinaryCompare)
            wanted = CStr(expected(i))
            If i = 13 And kind = 0 Then wanted = "Для повторных НДС использовалась стартовая плоскость из предыдущего близкого решения; повторов: 1."
            If i = 13 And kind = 1 Then wanted = "Стартовая плоскость: использовано предыдущее близкое решение; повторов: 1."
            If i = 14 And kind = 1 Then wanted = "Стартовая плоскость: подобрана по активной зоне сечения; повторов: 1."
            If Len(wanted) = 0 Then
                TestAssertTrue stats, "report.format.emptySkipped." & CStr(kind), startAt = 0
            Else
                TestAssertTrue stats, "report.format.title." & CStr(kind) & "." & CStr(i), startAt > 0
                If startAt > 0 Then
                    startAt = startAt + Len(title) + Len(vbCrLf)
                    endAt = InStr(startAt, text, "FORMAT_END_" & CStr(kind) & "_" & Format$(i, "00"), vbBinaryCompare)
                    TestAssertTrue stats, "report.format.end." & CStr(kind) & "." & CStr(i), endAt > 0
                    If endAt > 0 Then
                        endAt = InStrRev(text, vbCrLf, endAt, vbBinaryCompare)
                        actual = Mid$(text, startAt, endAt - startAt)
                        TestAssertTrue stats, "report.format.exact." & CStr(kind) & "." & CStr(i), StrComp(actual, wanted, vbBinaryCompare) = 0
                        If StrComp(actual, wanted, vbBinaryCompare) <> 0 Then TestAppendLine stats, "actual=" & actual & "; expected=" & wanted
                    End If
                End If
            End If
        Next i
    Next kind
    GoTo Finished
Failed:
    TestAssertTrue stats, "report.format.runtime." & CStr(Err.Number) & "." & Err.Description, False
Finished:
    On Error Resume Next
    If captured Then
        If oldEnabled Then TestSetSystemSetting "General.ExecutionReportEnabled", "Yes" Else TestSetSystemSetting "General.ExecutionReportEnabled", "No"
    End If
    On Error GoTo 0
    RunExecutionReportFormattingTests = stats.Report & "TOTAL_REPORT_FORMATTING: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет защиту пользовательского сценария расчета от лишней работы Excel.
' Макрос временно отключает события, экран, предупреждения и автопересчет,
' но после завершения должен вернуть настройки текущего Excel.Application
' ровно в исходное состояние.
Private Sub TestExcelApplicationStateGuardRestoresSettings(ByRef stats As TUiTestStats)
    TestPrepareCircleInput
    TestSetSystemSetting "General.ExecutionReportEnabled", "No"

    Dim app As Object
    Set app = ThisWorkbook.Application

    Dim oldCalculation As Variant
    Dim oldEnableEvents As Boolean
    Dim oldScreenUpdating As Boolean
    Dim oldDisplayAlerts As Boolean
    oldCalculation = app.Calculation
    oldEnableEvents = app.EnableEvents
    oldScreenUpdating = app.ScreenUpdating
    oldDisplayAlerts = app.DisplayAlerts

    Dim message As String
    message = TestFixtureRunSectionWorkflow(ThisWorkbook, False)

    TestAssertTrue stats, "ui.excelGuard.run", InStr(1, message, "Расчет", vbTextCompare) > 0
    TestAssertTrue stats, "ui.excelGuard.calculation", app.Calculation = oldCalculation
    TestAssertTrue stats, "ui.excelGuard.events", app.EnableEvents = oldEnableEvents
    TestAssertTrue stats, "ui.excelGuard.screen", app.ScreenUpdating = oldScreenUpdating
    TestAssertTrue stats, "ui.excelGuard.alerts", app.DisplayAlerts = oldDisplayAlerts
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет общий перенос нагрузки к бетонному центру, используемый batch
' и export-данными. AutoCAD COM здесь не вызывается, DWG не создается.
Private Sub TestAutoCADExportUsesSharedLoadReference(ByRef stats As TUiTestStats)
    TestPrepareRectanglesInput
    TestSetSystemSetting "Load.ReferenceOffsetX", "0"
    TestSetSystemSetting "Load.ReferenceOffsetY", "0"

    Dim loads As Object
    Set loads = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange
    loads.Cells.Item(2, 2).Value2 = -1000000#
    loads.Cells.Item(2, 3).Value2 = 0#
    loads.Cells.Item(2, 4).Value2 = 0#

    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook

    Dim units As CUnitSystem
    Set units = New CUnitSystem
    units.LoadFromSettings settings

    Dim materialProvider As CMaterialModelProvider
    Set materialProvider = New CMaterialModelProvider
    materialProvider.Initialize settings, units

    Dim section As CSectionModel
    Set section = BuildWorkbookSectionModel(ThisWorkbook, settings, units)
    Dim props As CSectionPropertiesCalculator
    Set props = New CSectionPropertiesCalculator
    props.CalculateConcrete section

    Dim batch As CBatchSectionCalculator
    Set batch = New CBatchSectionCalculator
    batch.Initialize section, materialProvider
    batch.ApplySettings settings, units
    Dim profiles As CCalculationProfileCatalog
    Set profiles = New CCalculationProfileCatalog
    profiles.LoadFromWorkbook ThisWorkbook
    Set batch.ProfileCatalog = profiles

    Dim reader As CLoadCombinationReader
    Set reader = New CLoadCombinationReader
    reader.LoadFromWorkbook ThisWorkbook, batch, units
    batch.ApplyLoadReference props.CentroidX, props.CentroidY, props.CentroidX, props.CentroidY
    batch.Execute

    TestAssertTrue stats, "ui.autocad.reference.converged", _
        batch.ResultAt(1).StateRepository.FindState(sstCrackedState).Converged
    TestAssertTrue stats, "ui.autocad.reference.point", Abs(batch.LoadReferenceX) > 0.000001 Or Abs(batch.LoadReferenceY) > 0.000001
    TestAssertClose stats, "ui.autocad.reference.concreteCenterX", props.CentroidX, batch.LoadReferenceX, 0.000001
    TestAssertClose stats, "ui.autocad.reference.concreteCenterY", props.CentroidY, batch.LoadReferenceY, 0.000001
    Dim expectedLoad As CSectionLoadState
    Set expectedLoad = New CSectionLoadState
    expectedLoad.Initialize batch.N(1), batch.UserMx(1), batch.UserMy(1), props.CentroidX, props.CentroidY
    TestAssertClose stats, "ui.autocad.reference.mxTransfer", batch.Mx(1), expectedLoad.InternalMx, 0.000001
    TestAssertClose stats, "ui.autocad.reference.myTransfer", batch.My(1), expectedLoad.InternalMy, 0.000001
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Читает один LC среди пустых строк: свободные строки не становятся
' ошибочными сочетаниями и не добавляют расчетов в пакет.
Private Sub TestSingleCombinationSkipsBlankRows(ByRef stats As TUiTestStats)
    TestPrepareCircleInput

    Dim batch As CBatchSectionCalculator
    Set batch = TestBuildUiBatch()

    Dim reader As CLoadCombinationReader
    Set reader = New CLoadCombinationReader
    reader.LoadFromWorkbook ThisWorkbook, batch

    Dim t0 As Double
    t0 = Timer
    batch.Execute

    TestAssertTrue stats, "ui.loads.single.count", batch.Count = 1
    TestAssertTrue stats, "ui.loads.single.id", batch.CombinationID(1) = "LC1"
    TestAssertTrue stats, "ui.loads.single.elapsed", (Timer - t0) < 20#
    TestAssertTrue stats, "ui.loads.single.noBlankInvalid", InStr(1, batch.DiagnosticLog, "InputErr", vbTextCompare) = 0
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет нижнюю допустимую высоту rngLoadCombinations: шапка плюс одна строка LC.
Private Sub TestLoadCombinationRangeMinimumRows(ByRef stats As TUiTestStats)
    On Error GoTo Failed

    Dim app As Object
    Set app = ThisWorkbook.Application
    Dim oldDisplayAlerts As Boolean
    oldDisplayAlerts = app.DisplayAlerts
    app.DisplayAlerts = False

    On Error Resume Next
    ThisWorkbook.Worksheets.Item("__tmpMinLoads").Delete
    On Error GoTo Failed

    Dim tempSheet As Object
    Set tempSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets.Item(ThisWorkbook.Worksheets.Count))
    tempSheet.Name = "__tmpMinLoads"

    Dim reader As CLoadCombinationReader
    Set reader = New CLoadCombinationReader

    Dim batch As CBatchSectionCalculator
    Set batch = TestBuildUiBatch()
    FillLoadCombinationTestRange tempSheet.Range("A1:G2"), 1, "PR1", "LC_MIN_"
    reader.LoadFromRange tempSheet.Range("A1:G2"), batch
    TestAssertTrue stats, "ui.loads.minRows.oneDataRow", batch.Count = 1

    Dim errorNumber As Long
    Dim errorText As String
    Set batch = TestBuildUiBatch()
    On Error Resume Next
    reader.LoadFromRange tempSheet.Range("A4:G4"), batch
    errorNumber = Err.Number
    errorText = Err.Description
    Err.Clear
    On Error GoTo Failed
    TestAssertTrue stats, "ui.loads.minRows.headerOnlyRejected", errorNumber <> 0 And _
        InStr(1, errorText, "минимум одну строку", vbTextCompare) > 0

CleanUp:
    On Error Resume Next
    If Not tempSheet Is Nothing Then tempSheet.Delete
    app.DisplayAlerts = oldDisplayAlerts
    On Error GoTo 0
    Exit Sub

Failed:
    stats.Failed = stats.Failed + 1
    TestAppendLine stats, "FAIL: ui.loads.minRows; err=" & CStr(Err.Number) & "; " & Err.Description
    Resume CleanUp
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Удаляет ID из строки с заданными усилиями: строка сохраняется как InputErr,
' а пользовательский макрос выводит понятный итог ошибки ввода.
Private Sub TestPartialCombinationIsInvalid(ByRef stats As TUiTestStats)
    TestPrepareCircleInput
    Dim loads As Object
    Set loads = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange
    loads.Cells.Item(2, 1).ClearContents

    Dim batch As CBatchSectionCalculator
    Set batch = TestBuildUiBatch()

    Dim reader As CLoadCombinationReader
    Set reader = New CLoadCombinationReader
    reader.LoadFromWorkbook ThisWorkbook, batch
    batch.Execute

    TestAssertTrue stats, "ui.loads.partial.count", batch.Count = 1
    TestAssertTrue stats, "ui.loads.partial.invalid", batch.ResultAt(1).Status = "InputErr"

    Dim message As String
    message = TestFixtureRunSectionWorkflow(ThisWorkbook, False)
    TestAssertTrue stats, "ui.loads.partial.message", InStr(1, message, "ошиб", vbTextCompare) > 0 And _
        InStr(1, message, "InputErr", vbTextCompare) > 0
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' При пустом ProfileId даже включенное автообновление не запускает вывод НДС.
' Проверяет сообщение InputErr и отсутствие строк element-results.
Private Sub TestInvalidProfileIdDoesNotRunPlot(ByRef stats As TUiTestStats)
    TestPrepareCircleInput
    TestSetSystemSetting "Plot.AutoUpdateAfterCalculation", "Yes"

    Dim loads As Object
    Set loads = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange
    loads.Cells.Item(2, 5).ClearContents

    Dim message As String
    message = TestFixtureRunSectionWorkflow(ThisWorkbook, False)

    TestAssertTrue stats, "ui.loads.invalidProfileId.message", _
        InStr(1, message, "InputErr", vbTextCompare) > 0 And _
        InStr(1, message, "результатов элементов", vbTextCompare) = 0
    TestAssertTrue stats, "ui.loads.invalidProfileId.noElementRows", _
        TestResultTableRowCount("rngNDMElementResults") = 1
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет, что расчет в режиме AutoCAD не запускает скрытый импорт.
' Пользователь должен сначала явно нажать отдельную кнопку импорта; иначе
' расчет останавливается до очистки старых Results и до обращения к AutoCAD.
Private Sub TestCalculationRequiresFormedSection(ByRef stats As TUiTestStats)
    TestPrepareCircleInput
    TestSetSystemSetting "Geometry.Source", "AutoCAD"
    ClearSectionResultsForWorkbook ThisWorkbook

    Dim message As String
    On Error Resume Next
    message = RunSectionCalculationForWorkbook(ThisWorkbook, False)
    Dim description As String
    description = Err.Description
    Err.Clear
    On Error GoTo 0

    TestAssertTrue stats, "ui.autocad.manualImport.required", _
        InStr(1, description, "нет сформированного сечения", vbTextCompare) > 0 And _
        InStr(1, description, "Сформировать сечение", vbTextCompare) > 0
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' При Generated первая кнопка строит сечение по таблице на листе «Сечения».
' Доступ к AutoCAD и расчёт НДС для этого не требуются.
Private Sub TestFormButtonBuildsGeneratedSection(ByRef stats As TUiTestStats)
    TestPrepareCircleInput
    TestSetSystemSetting "Geometry.Source", "Generated"

    Dim message As String
    On Error Resume Next
    message = FormSectionForWorkbook(ThisWorkbook)
    Dim description As String
    description = Err.Description
    Err.Clear
    On Error GoTo 0

    TestAssertTrue stats, "ui.form.generated.completed", Len(description) = 0 And InStr(message, "Сечение сформировано") > 0
    TestAssertTrue stats, "ui.form.generated.snapshot", ResultsGeometrySource(ThisWorkbook) = "Generated"
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет, что live AutoCAD-import не масштабирует координаты DWG через
' пользовательские INPUT-единицы. По контракту проекта AutoCAD-чертеж
' импортируется как миллиметры, иначе смена Units.Length.Input сдвинула бы
' весь imported snapshot относительно исходного чертежа.
Private Sub TestAutoCADImporterTreatsDrawingUnitsAsMillimeters(ByRef stats As TUiTestStats)
    TestPrepareCircleInput
    TestSetSystemSetting "Units.Length.Input", "m"
    TestSetSystemSetting "Units.Area.Input", "m2"

    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook

    Dim units As CUnitSystem
    Set units = New CUnitSystem
    units.LoadFromSettings settings

    Dim concreteRegion As CTestFakeAcadRegion
    Set concreteRegion = New CTestFakeAcadRegion
    concreteRegion.TestInitialize 2500#, 1000#, 2000#, 2500# * 2500# / 12#, _
        2500# * 2500# / 12#, 0#, "Concrete", "C1"

    Dim rebarRegion As CTestFakeAcadRegion
    Set rebarRegion = New CTestFakeAcadRegion
    rebarRegion.TestInitialize GEOM_PI * 25# * 25# / 4#, 1100#, 1900#, 1#, 1#, 0#, _
        "Reinf", "R1"

    Dim modelSpace As Collection
    Set modelSpace = New Collection
    modelSpace.Add concreteRegion
    modelSpace.Add rebarRegion

    Dim importer As CAutoCADSectionModelImporter
    Set importer = New CAutoCADSectionModelImporter

    Dim section As CSectionModel
    Set section = importer.ImportFromModelSpace(modelSpace, "Concrete", "Reinf", "Rebar", 0.000001, units)

    TestAssertClose stats, "ui.autocad.importUnits.concreteX", section.ConcreteX(1), 1000#, 0.000001
    TestAssertClose stats, "ui.autocad.importUnits.concreteY", section.ConcreteY(1), 2000#, 0.000001
    TestAssertClose stats, "ui.autocad.importUnits.concreteArea", section.ConcreteArea(1), 2500#, 0.000001
    TestAssertClose stats, "ui.autocad.importUnits.localIx", section.ConcreteLocalIx(1), _
        2500# * 2500# / 12#, 0.001
    TestAssertClose stats, "ui.autocad.importUnits.rebarX", section.RebarX(1), 1100#, 0.000001
    TestAssertClose stats, "ui.autocad.importUnits.rebarDiameter", section.RebarDiameter(1), 25#, 0.000001

    TestSetSystemSetting "Units.Length.Input", "mm"
    TestSetSystemSetting "Units.Area.Input", "mm2"
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет, что подпись размера округляется до целого миллиметра до перевода
' в пользовательскую единицу вывода. При Units.Length.Output = m размер 517 мм
' должен отображаться как 0.517 m, а не как 1 m или 0 m.
Private Sub TestAnnotationDimensionTextRoundsInMillimeters(ByRef stats As TUiTestStats)
    Dim oldOutputLength As String
    oldOutputLength = TestGetSystemSetting("Units.Length.Output")

    On Error GoTo RestoreAndFail
    TestPrepareCircleInput
    TestSetSystemSetting "Geometry.Source", "AutoCAD"
    TestSetSystemSetting "Units.Length.Output", "m"

    Dim section As CSectionModel
    Set section = New CSectionModel
    section.SourceType = "AutoCADImport"
    section.AddConcreteElement 50#, 50#, 10000#, 1, vbNullString, vbNullString, "Rectangle", 100#, 100#
    section.AddConcreteElement 250#, 50#, 10000#, 1, vbNullString, vbNullString, "Rectangle", 100#, 100#
    section.AddConcreteElement 450#, 50#, 7200#, 1, vbNullString, vbNullString, _
        "Rectangle", 120#, 60#, GEOM_PI / 6#

    Dim writer As CNDMResultsWriter
    Set writer = New CNDMResultsWriter
    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook
    Dim units As CUnitSystem
    Set units = New CUnitSystem
    units.LoadFromSettings settings
    writer.WriteGeometryPreview ThisWorkbook, section, PrepareSectionSnapshot(section), units

    Dim annotationData As Variant
    annotationData = TestResultTable("rngNDMSectionAnnotations")
    TestAssertTextEquals stats, "ui.autocad.preview.roundedWidthTextMeters", _
        TestAnnotationTextByID(annotationData, "DIM_AUTO_BOUNDS_B"), ChrW$(&H2248) & " 0.517 m"
    TestAssertTextEquals stats, "ui.autocad.preview.roundedHeightTextMeters", _
        TestAnnotationTextByID(annotationData, "DIM_AUTO_BOUNDS_H"), ChrW$(&H2248) & " 0.112 m"

Restore:
    TestSetSystemSetting "Units.Length.Output", oldOutputLength
    Exit Sub

RestoreAndFail:
    stats.Failed = stats.Failed + 1
    TestAppendLine stats, "FAIL: ui.autocad.preview.roundedDimensionTextMeters; " & Err.Description
    Resume Restore
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет preview импортированной геометрии без реального AutoCAD.
' Важна вся цепочка: CSectionModel(AutoCADImport) -> Results annotations ->
' CSectionPlotDataReader -> CSectionPlotter -> Chart.Shapes размерных линий.
Private Sub TestAutoCADPreviewWritesAndDrawsBoundsDimensions(ByRef stats As TUiTestStats)
    TestPrepareCircleInput
    TestSetSystemSetting "Geometry.Source", "AutoCAD"

    Dim section As CSectionModel
    Set section = New CSectionModel
    section.SourceType = "AutoCADImport"
    section.AddConcreteElement 50#, 50#, 10000#, 1, vbNullString, vbNullString, "Rectangle", 100#, 100#
    section.AddConcreteElement 250#, 50#, 10000#, 1, vbNullString, vbNullString, "Rectangle", 100#, 100#
    section.AddConcreteElement 450#, 50#, 7200#, 1, vbNullString, vbNullString, _
        "Rectangle", 120#, 60#, GEOM_PI / 6#
    section.AddConcreteElement 250#, 50#, 2500#, 1, vbNullString, vbNullString, _
        "EquivalentSquare", 0#, 0#
    section.AddRebarElement 50#, 50#, 20#, 0#, "Rebar"
    section.AddRebarElement 250#, 80#, 20#, 0#, "Rebar"
    section.AddRebarElement 450#, 80#, 32#, 0#, "Rebar"
    section.ApplyAverageRotationToEquivalentAreaFallbacks

    Dim writer As CNDMResultsWriter
    Set writer = New CNDMResultsWriter
    writer.WriteGeometryPreview ThisWorkbook, section, PrepareSectionSnapshot(section)

    Dim annotationData As Variant
    annotationData = TestResultTable("rngNDMSectionAnnotations")
    TestAssertTrue stats, "ui.autocad.preview.boundsAnnotations", TestCountAnnotationType(annotationData, "DIMENSION") = 2
    TestAssertTrue stats, "ui.autocad.preview.rebarAnnotation", TestCountAnnotationType(annotationData, "REBAR_GLOBAL_ANNOTATION") = 1
    TestAssertTextEquals stats, "ui.autocad.preview.rebarText", _
        TestAnnotationTextByID(annotationData, "REBAR_AUTO_ALL"), "ALL: 1" & ChrW$(&H2205) & "32+2" & ChrW$(&H2205) & "20"
    TestAssertTrue stats, "ui.autocad.preview.approxText", _
        InStr(1, CStr(annotationData(2, TestResultHeaderColumn(annotationData, "Text"))), ChrW$(&H2248), vbTextCompare) > 0
    TestAssertTextEquals stats, "ui.autocad.preview.roundedWidthText", _
        TestAnnotationTextByID(annotationData, "DIM_AUTO_BOUNDS_B"), ChrW$(&H2248) & " 517 mm"
    TestAssertTextEquals stats, "ui.autocad.preview.roundedHeightText", _
        TestAnnotationTextByID(annotationData, "DIM_AUTO_BOUNDS_H"), ChrW$(&H2248) & " 112 mm"
    TestAssertTrue stats, "ui.autocad.preview.russianComment", _
        InStr(1, CStr(annotationData(2, TestResultHeaderColumn(annotationData, "Comment"))), "Приблизительная", vbTextCompare) > 0
    TestAssertClose stats, "ui.autocad.preview.boundsWidthUsesRotation", _
        TestAnnotationValueByID(annotationData, "DIM_AUTO_BOUNDS_B"), _
        450# + 0.5 * (120# * Cos(GEOM_PI / 6#) + 60# * Sin(GEOM_PI / 6#)), 0.001
    TestAssertClose stats, "ui.autocad.preview.boundsHeightUsesRotation", _
        TestAnnotationValueByID(annotationData, "DIM_AUTO_BOUNDS_H"), _
        120# * Sin(GEOM_PI / 6#) + 60# * Cos(GEOM_PI / 6#), 0.001

    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook

    Dim reader As CSectionPlotDataReader
    Set reader = New CSectionPlotDataReader
    reader.LoadGeometryPreviewFromWorkbook ThisWorkbook, settings
    TestAssertTrue stats, "ui.autocad.preview.readerAnnotations", reader.AnnotationCount = 3

    UpdateSectionPlotForWorkbook ThisWorkbook
    TestAssertTrue stats, "ui.autocad.preview.dimensionShapes", TestCountPlotShapes("AnnotationLine") > 0
    TestAssertTrue stats, "ui.autocad.preview.rebarTextShape", TestPlotShapeTextContains("ALL: 1")
    TestAssertTrue stats, "ui.autocad.preview.rotatedElementShape", _
        TestPlotShapeRotationExists("ElementConcrete", -30#, 0.5)
    Dim expectedFallbackDegrees As Double
    expectedFallbackDegrees = -0.5 * Atn(Sin(GEOM_PI / 3#) / (2# + Cos(GEOM_PI / 3#))) * 180# / GEOM_PI
    TestAssertTrue stats, "ui.autocad.preview.squareFallbackRotation", _
        TestPlotShapeRotationExists("ElementConcrete", expectedFallbackDegrees, 0.5)
End Sub


' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет не только наличие осей/точки, но и их взаимное положение.
' Для режима Transformed центр главных осей берется из Transformed.Centroid,
' а точка нагрузки - из LoadReference. Их экранный сдвиг должен совпадать с
' расчетным сдвигом из Results после одного общего model-to-chart масштаба.
Private Sub TestPlotOverlayCoordinatesMatchResults(ByRef stats As TUiTestStats)
    Dim oldAxisLabelsEnabled As String
    Dim oldAxisLabelsFontSize As String
    oldAxisLabelsEnabled = TestGetSystemSetting("Plot.AxisLabelsEnabled")
    oldAxisLabelsFontSize = TestGetSystemSetting("Plot.AxisLabelsFontSize")

    TestPrepareUserRectanglesMomentUltimateInput
    TestSetSystemSetting "Plot.PrincipalAxesMode", "Transformed"
    TestSetSystemSetting "Plot.LoadApplicationPointEnabled", "Yes"
    TestSetSystemSetting "Plot.AxisLabelsEnabled", "Yes"
    TestSetSystemSetting "Plot.AxisLabelsFontSize", "9"
    TestSetSystemSetting "Plot.ResultLabelsEnabled", "No"
    TestSetSystemSetting "Plot.AutoUpdateAfterCalculation", "Yes"

    Dim message As String
    message = TestFixtureRunSectionWorkflow(ThisWorkbook, False)
    TestAssertTrue stats, "ui.plot.overlay.run", InStr(1, message, "Расчет завершен", vbTextCompare) > 0

    Dim loadX As Double
    Dim loadY As Double
    Dim centroidX As Double
    Dim centroidY As Double
    loadX = TestResultsLengthPropertyMm("ALL", "LoadReferenceX")
    loadY = TestResultsLengthPropertyMm("ALL", "LoadReferenceY")
    centroidX = TestResultsLengthPropertyMm("ALL", "Transformed.CentroidX")
    centroidY = TestResultsLengthPropertyMm("ALL", "Transformed.CentroidY")

    Dim loadPoint As Object
    Set loadPoint = TestFirstGeneratedPlotShape("LoadPoint")
    TestAssertTrue stats, "ui.plot.overlay.shapes", TestCountGeneratedPlotShapes("Principal1") > 0 And _
        TestCountGeneratedPlotShapes("Principal2") > 0 And Not loadPoint Is Nothing
    If TestCountGeneratedPlotShapes("Principal1") <= 0 Or TestCountGeneratedPlotShapes("Principal2") <= 0 Or loadPoint Is Nothing Then
        TestSetSystemSetting "Plot.AxisLabelsEnabled", oldAxisLabelsEnabled
        TestSetSystemSetting "Plot.AxisLabelsFontSize", oldAxisLabelsFontSize
        Exit Sub
    End If
    TestAssertTrue stats, "ui.plot.overlay.chartLayer", TestCountPlotShapes("Principal1") > 0 And _
        TestCountPlotShapes("Principal2") > 0 And TestCountPlotShapes("LoadPoint") > 0 And _
        TestCountWorksheetPlotShapes("Principal1") = 0 And TestCountWorksheetPlotShapes("Principal2") = 0 And _
        TestCountWorksheetPlotShapes("LoadPoint") = 0
    TestAssertTrue stats, "ui.plot.axisLabels.exists", TestCountPlotShapes("AxisLabelX") = 1 And TestCountPlotShapes("AxisLabelY") = 1
    TestAssertTrue stats, "ui.plot.axisLabels.text", TestPlotShapeTextContains("Ось X") And TestPlotShapeTextContains("Ось Y")
    TestAssertTrue stats, "ui.plot.axisLabels.italic", TestPlotShapeTextItalic("AxisLabelX") And TestPlotShapeTextItalic("AxisLabelY")
    TestAssertClose stats, "ui.plot.axisLabels.fontSize", TestPlotShapeTextFontSize("AxisLabelX"), 9#, 0.01
    TestAssertTrue stats, "ui.plot.axisLabels.insideFrame", TestPlotShapesInsideStableFrame("AxisLabelX", 0.5) And _
        TestPlotShapesInsideStableFrame("AxisLabelY", 0.5)

    Dim loadCenterX As Double
    Dim loadCenterY As Double
    TestAssertTrue stats, "ui.plot.overlay.loadCenter", TestGeneratedPlotShapeCenterAverage("LoadPoint", loadCenterX, loadCenterY)

    Dim plot As Object
    Set plot = ThisWorkbook.Worksheets.Item("Расчет").ChartObjects("chtNDMSectionPlot")
    TestAssertTrue stats, "ui.plot.overlay.clippedAxes", TestPlotShapesInsideStableFrame("Principal1", 1#) And _
        TestPlotShapesInsideStableFrame("Principal2", 1#)

    Dim minX As Double
    Dim maxX As Double
    Dim minY As Double
    Dim maxY As Double
    minX = CDbl(plot.Chart.Axes(1).MinimumScale)
    maxX = CDbl(plot.Chart.Axes(1).MaximumScale)
    minY = CDbl(plot.Chart.Axes(2).MinimumScale)
    maxY = CDbl(plot.Chart.Axes(2).MaximumScale)

    Dim scaleValue As Double
    scaleValue = (CDbl(plot.Width) - TEST_PLOT_FRAME_WIDTH_MARGIN) / (maxX - minX)
    If (CDbl(plot.Height) - TEST_PLOT_FRAME_HEIGHT_MARGIN) / (maxY - minY) < scaleValue Then _
        scaleValue = (CDbl(plot.Height) - TEST_PLOT_FRAME_HEIGHT_MARGIN) / (maxY - minY)

    Dim expectedCentroidX As Double
    Dim expectedCentroidY As Double
    Dim expectedLoadX As Double
    Dim expectedLoadY As Double
    expectedCentroidX = CDbl(plot.Left) + TEST_PLOT_FRAME_LEFT + (centroidX - minX) / (maxX - minX) * _
        (CDbl(plot.Width) - TEST_PLOT_FRAME_WIDTH_MARGIN)
    expectedCentroidY = CDbl(plot.Top) + TEST_PLOT_FRAME_TOP + (CDbl(plot.Height) - TEST_PLOT_FRAME_HEIGHT_MARGIN) - _
        (centroidY - minY) / (maxY - minY) * (CDbl(plot.Height) - TEST_PLOT_FRAME_HEIGHT_MARGIN)
    expectedLoadX = CDbl(plot.Left) + TEST_PLOT_FRAME_LEFT + (loadX - minX) / (maxX - minX) * _
        (CDbl(plot.Width) - TEST_PLOT_FRAME_WIDTH_MARGIN)
    expectedLoadY = CDbl(plot.Top) + TEST_PLOT_FRAME_TOP + (CDbl(plot.Height) - TEST_PLOT_FRAME_HEIGHT_MARGIN) - _
        (loadY - minY) / (maxY - minY) * (CDbl(plot.Height) - TEST_PLOT_FRAME_HEIGHT_MARGIN)

    TestAssertTrue stats, "ui.plot.overlay.principal1ThroughCentroid", _
        TestPlotShapeBoundsContainAbsolutePoint("Principal1", expectedCentroidX, expectedCentroidY, 1.5)
    TestAssertTrue stats, "ui.plot.overlay.principal2ThroughCentroid", _
        TestPlotShapeBoundsContainAbsolutePoint("Principal2", expectedCentroidX, expectedCentroidY, 1.5)
    TestAssertClose stats, "ui.plot.overlay.loadX", loadCenterX, expectedLoadX, 0.8
    TestAssertClose stats, "ui.plot.overlay.loadY", loadCenterY, expectedLoadY, 0.8
    TestAssertClose stats, "ui.plot.overlay.loadDx", loadCenterX - expectedCentroidX, (loadX - centroidX) * scaleValue, 0.8
    TestAssertClose stats, "ui.plot.overlay.loadDy", loadCenterY - expectedCentroidY, -(loadY - centroidY) * scaleValue, 0.8

    Dim beforeMoveX As Double
    Dim beforeMoveY As Double
    Dim afterMoveX As Double
    Dim afterMoveY As Double
    Dim oldLeft As Double
    Dim oldTop As Double
    oldLeft = CDbl(plot.Left)
    oldTop = CDbl(plot.Top)
    TestAssertTrue stats, "ui.plot.overlay.move.before", TestGeneratedPlotShapeCenterAverage("LoadPoint", beforeMoveX, beforeMoveY)
    plot.Left = oldLeft + 17#
    plot.Top = oldTop + 11#
    TestAssertTrue stats, "ui.plot.overlay.move.after", TestGeneratedPlotShapeCenterAverage("LoadPoint", afterMoveX, afterMoveY)
    TestAssertClose stats, "ui.plot.overlay.move.dx", afterMoveX - beforeMoveX, 17#, 0.3
    TestAssertClose stats, "ui.plot.overlay.move.dy", afterMoveY - beforeMoveY, 11#, 0.3
    plot.Left = oldLeft
    plot.Top = oldTop
    TestSetSystemSetting "Plot.AxisLabelsEnabled", oldAxisLabelsEnabled
    TestSetSystemSetting "Plot.AxisLabelsFontSize", oldAxisLabelsFontSize
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет, что выбор будущего источника не меняет уже сформированное сечение.
' Только команда формирования заменяет сохранённую геометрию и её схему.
Private Sub TestPlotUsesSavedGeometryAfterSourceChange(ByRef stats As TUiTestStats)
    TestPrepareCircleInput
    TestSetSystemSetting "Geometry.Source", "AutoCAD"
    Dim section As CSectionModel: Set section = New CSectionModel
    section.SourceType = "AutoCADImport"
    section.AddConcreteElement 50#, 50#, 10000#, 1, vbNullString, vbNullString, "Rectangle", 100#, 100#
    section.AddRebarElement 50#, 50#, 20#, 0#, "Rebar"
    Dim writer As CNDMResultsWriter: Set writer = New CNDMResultsWriter
    writer.WriteGeometryPreview ThisWorkbook, section, PrepareSectionSnapshot(section)
    UpdateSectionPlotForWorkbook ThisWorkbook
    TestAssertTrue stats, "ui.plot.preview.title", TestPlotVisibleTitleContains("Сформированное сечение: AutoCADImport")
    Dim geometry As Variant, annotations As Variant, actual As Variant, solves As Long
    geometry = TestResultTable("rngNDMSectionGeometry"): annotations = TestResultTable("rngNDMSectionAnnotations")
    solves = SectionEquilibriumSolveCount()
    TestSetSystemSetting "Geometry.Source", "Generated"
    UpdateSectionPlotForWorkbook ThisWorkbook
    TestAssertTrue stats, "ui.plot.sourceChange.savedTitle", TestPlotVisibleTitleContains("Сформированное сечение: AutoCADImport")
    actual = TestResultTable("rngNDMSectionGeometry")
    TestAssertTrue stats, "ui.plot.sourceChange.savedGeometry", TestAudit02SnapshotTablesEqual(geometry, actual)
    actual = TestResultTable("rngNDMSectionAnnotations")
    TestAssertTrue stats, "ui.plot.sourceChange.savedAnnotations", TestAudit02SnapshotTablesEqual(annotations, actual)
    TestAssertTrue stats, "ui.plot.sourceChange.noSolve", SectionEquilibriumSolveCount() = solves
    Dim message As String: message = FormSectionForWorkbook(ThisWorkbook)
    TestAssertTrue stats, "ui.plot.sourceChange.formationReplacesSource", ResultsGeometrySource(ThisWorkbook) <> "AutoCADImport"
    TestAssertTrue stats, "ui.plot.sourceChange.newTitle", TestPlotVisibleTitleContains("Сформированное сечение: " & ResultsGeometrySource(ThisWorkbook))
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет DirectState-сценарий без определяющего сочетания по прочности.
' При Plot.LoadCase = Worst схема должна показать первый рассчитанный LC из
' Results, а не оставлять AutoCAD-preview и не очищаться до пустого окна.
Private Sub TestGeneratedDirectStateWorstStillDrawsFirstCalculatedLC(ByRef stats As TUiTestStats)
    TestPrepareCircleInput
    TestSetSystemSetting "Geometry.Source", "Generated"
    TestSetSystemSetting "Plot.LoadCase", "Worst"

    Dim message As String
    message = TestFixtureRunSectionWorkflow(ThisWorkbook, False)

    TestAssertTrue stats, "ui.plot.generatedWorst.message", InStr(1, message, "Расчет завершен", vbTextCompare) > 0
    TestAssertTrue stats, "ui.plot.generatedWorst.drawsCalculatedLc", TestPlotVisibleTitleContains("LC1")
    TestAssertTrue stats, "ui.plot.generatedWorst.noImportTitle", _
        Not TestPlotVisibleTitleContains("Импортированная геометрия AutoCAD")
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет, что схема после расчета выбирает состояние через профиль, а
' численные Stress/Strain берет из snapshot Results.
Private Sub TestProfileDrivenPlotUsesSnapshotState(ByRef stats As TUiTestStats)
    TestPrepareCircleInput
    TestSetSystemSetting "Plot.AutoUpdateAfterCalculation", "Yes"
    TestSetSystemSetting "Plot.LoadCase", "Worst"

    Dim message As String
    message = TestFixtureRunSectionWorkflow(ThisWorkbook, False)

    TestAssertTrue stats, "ui.profilePlot.message", InStr(1, message, "Расчет завершен", vbTextCompare) > 0
    TestAssertTrue stats, "ui.profilePlot.hasElementStateRows", TestResultTableRowCount("rngNDMElementResults") > 1
    TestAssertTrue stats, "ui.profilePlot.drawsLoadCase", TestPlotVisibleTitleContains("LC1")
    TestAssertTrue stats, "ui.profilePlot.noImportTitle", _
        Not TestPlotVisibleTitleContains("Импортированная геометрия AutoCAD")
    TestAssertTrue stats, "ui.profilePlot.commonLoadReference", _
        Len(TestResultsPropertyValue("ALL", "LoadReferenceX")) > 0 And Len(TestResultsPropertyValue("ALL", "LoadReferenceY")) > 0

End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет, что отсутствие выбранного StateType не превращается в popup-only
' ошибку. Схема должна показать сохраненную геометрию Results и крупную
' подпись под сечением, что запрошенного состояния в snapshot нет.
Private Sub TestMissingProfileStateDrawsGeometryOnly(ByRef stats As TUiTestStats, Optional ByVal storage As String = "Excel")
    Dim originalStorage As String, originalName As String, path As String, label As String
    originalStorage = TestGetSystemSetting("Results.SnapshotStorage")
    originalName = TestGetSystemSetting("Results.SnapshotName")
    label = "ui.plot.missingState": If storage = "File" Then label = label & ".File"
    On Error GoTo Failed
    TestSetSystemSetting "Results.SnapshotStorage", storage
    If storage = "File" Then
        TestSetSystemSetting "Results.SnapshotName", "TestMissingState_" & Format$(Now, "yyyymmdd_hhnnss") & ".ndm"
        Dim store As CCalculationSnapshotStore: Set store = New CCalculationSnapshotStore
        path = store.SnapshotPath(ThisWorkbook)
        If TestFileExists(path) Then path = vbNullString: Err.Raise 5, , "Файл текущего теста уже существует."
    End If
    TestPrepareCircleInput
    TestSetSystemSetting "Plot.AutoUpdateAfterCalculation", "No"
    TestSetSystemSetting "Plot.LoadCase", "LC1"

    Dim message As String
    message = TestFixtureRunSectionWorkflow(ThisWorkbook, False)
    TestAssertTrue stats, label & ".run", InStr(1, message, "Расчет завершен", vbTextCompare) > 0

    TestSetProfileSetting "PR2", "Visualization.State", "CapacityState"
    Dim solves As Long: solves = SectionEquilibriumSolveCount()
    Dim errorDescription As String
    On Error Resume Next
    UpdateSectionPlotForWorkbook ThisWorkbook
    errorDescription = Err.Description
    On Error GoTo Failed

    TestAssertTrue stats, label & ".noError", Len(errorDescription) = 0
    TestAssertTrue stats, label & ".geometryTitle", TestPlotVisibleTitleContains("Геометрия расчетного сечения")
    TestAssertTrue stats, label & ".note", TestPlotShapeTextContains("Запрашиваемое состояние")
    TestAssertTrue stats, label & ".plotNoSolve", SectionEquilibriumSolveCount() = solves
    UpdateSectionResultsForWorkbook ThisWorkbook
    TestAssertTrue stats, label & ".buttonGeometryTitle", TestPlotVisibleTitleContains("Геометрия расчетного сечения")
    TestAssertTrue stats, label & ".buttonNote", TestPlotShapeTextContains("Запрашиваемое состояние")
    TestAssertTrue stats, label & ".buttonNoSolve", SectionEquilibriumSolveCount() = solves
Finished:
    If Len(path) > 0 Then
        If TestFileExists(path) And store.HasFileReference(ThisWorkbook) Then
            TestSetSystemSetting "Results.SnapshotStorage", "Excel"
            UpdateSectionResultsForWorkbook ThisWorkbook
        End If
        TestDeleteFileIfExists path
    End If
    TestSetSystemSetting "Results.SnapshotStorage", originalStorage
    TestSetSystemSetting "Results.SnapshotName", originalName
    Exit Sub
Failed:
    TestAssertTrue stats, label & ".runtime." & CStr(Err.Number) & "." & Err.Description, False
    Resume Finished
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет, что кнопка расчета в режиме AutoCAD использует уже сохраненную
' геометрию Results. Макрос не должен повторно импортировать Region и не должен
' повторять сведения об импорте или создании сетки: они относятся только
' к команде формирования сечения.
Private Sub TestAutoCADCalculationMessageUsesSavedGeometry(ByRef stats As TUiTestStats)
    TestPrepareCircleInput

    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook

    Dim units As CUnitSystem
    Set units = New CUnitSystem
    units.LoadFromSettings settings

    Dim section As CSectionModel
    Set section = New CSectionModel
    section.SourceType = "AutoCADImport"
    section.ConcreteMeshSource = "AutoCADContours"
    section.AddConcreteElement 950#, -750#, 10000#, 1, vbNullString, vbNullString, "Rectangle", 100#, 100#
    section.AddConcreteElement 1050#, -750#, 10000#, 1, vbNullString, vbNullString, "Rectangle", 100#, 100#
    section.AddConcreteElement 950#, -650#, 10000#, 1, vbNullString, vbNullString, "Rectangle", 100#, 100#
    section.AddConcreteElement 1050#, -650#, 10000#, 1, vbNullString, vbNullString, "Rectangle", 100#, 100#
    section.AddRebarElement 930#, -770#, 20#, 0#, "Rebar"
    section.AddRebarElement 1070#, -770#, 20#, 0#, "Rebar"
    section.AddRebarElement 930#, -630#, 20#, 0#, "Rebar"
    section.AddRebarElement 1070#, -630#, 20#, 0#, "Rebar"

    Dim writer As CNDMResultsWriter
    Set writer = New CNDMResultsWriter
    writer.WriteGeometryPreview ThisWorkbook, section, PrepareSectionSnapshot(section), units

    Dim beforeGeometry As Variant
    beforeGeometry = TestResultTable("rngNDMSectionGeometry")
    Dim beforeConcreteX As Double
    Dim beforeConcreteY As Double
    Dim beforeRebarX As Double
    Dim beforeRebarY As Double
    beforeConcreteX = TestGeometryResultValue(beforeGeometry, "C1", "X")
    beforeConcreteY = TestGeometryResultValue(beforeGeometry, "C1", "Y")
    beforeRebarX = TestGeometryResultValue(beforeGeometry, "R1", "X")
    beforeRebarY = TestGeometryResultValue(beforeGeometry, "R1", "Y")

    TestSetSystemSetting "Geometry.Source", "AutoCAD"
    Dim loads As Object
    Set loads = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange
    TestClearDataRows loads
    loads.Cells.Item(2, 1).Value2 = "LC1"
    loads.Cells.Item(2, 2).Value2 = 0#
    loads.Cells.Item(2, 3).Value2 = -1000000#
    loads.Cells.Item(2, 4).Value2 = -750000#
    loads.Cells.Item(2, 5).Value2 = "PR1"
    loads.Cells.Item(2, 6).Value2 = ChrW$(&H3BB) & "*Mxy"
    loads.Cells.Item(2, 7).Value2 = "saved geometry"

    Dim message As String
    message = TestFixtureRunSectionWorkflow(ThisWorkbook, False)

    TestAssertTrue stats, "ui.autocad.run.savedGeometry.message", _
        InStr(1, message, "Расчет завершен", vbTextCompare) > 0
    TestAssertTrue stats, "ui.autocad.run.savedGeometry.noImportText", _
        InStr(1, message, "AutoCAD", vbTextCompare) = 0
    TestAssertTrue stats, "ui.autocad.run.savedGeometry.noMeshNotice", _
        InStr(1, message, CONTOUR_MESH_GENERATION_NOTICE, vbTextCompare) = 0

    Dim afterGeometry As Variant
    afterGeometry = TestResultTable("rngNDMSectionGeometry")
    TestAssertClose stats, "ui.autocad.run.savedGeometry.concreteX", _
        TestGeometryResultValue(afterGeometry, "C1", "X"), beforeConcreteX, 0.000001
    TestAssertClose stats, "ui.autocad.run.savedGeometry.concreteY", _
        TestGeometryResultValue(afterGeometry, "C1", "Y"), beforeConcreteY, 0.000001
    TestAssertClose stats, "ui.autocad.run.savedGeometry.rebarX", _
        TestGeometryResultValue(afterGeometry, "R1", "X"), beforeRebarX, 0.000001
    TestAssertClose stats, "ui.autocad.run.savedGeometry.rebarY", _
        TestGeometryResultValue(afterGeometry, "R1", "Y"), beforeRebarY, 0.000001
    Dim afterAnnotations As Variant
    afterAnnotations = TestResultTable("rngNDMSectionAnnotations")
    TestAssertTextEquals stats, "ui.autocad.run.savedGeometry.rebarText", _
        TestAnnotationTextByID(afterAnnotations, "REBAR_AUTO_ALL"), "ALL: 4" & ChrW$(&H2205) & "20"
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет разрешенный пустой My как 0 и пропуск достоверно нулевого LC.
' Пропуск нуля не должен скрывать ошибку или менять число ненулевых сочетаний.
Private Sub TestBlankMomentDefaultsToZeroAndZeroLoadsAreSkipped(ByRef stats As TUiTestStats)
    TestPrepareCircleInput
    Dim loads As Object
    Set loads = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange
    loads.Cells.Item(2, 4).ClearContents
    loads.Cells.Item(3, 1).Value2 = "ZERO"
    loads.Cells.Item(3, 2).Value2 = 0#
    loads.Cells.Item(3, 3).Value2 = 0#
    loads.Cells.Item(3, 4).Value2 = 0#
    loads.Cells.Item(3, 5).Value2 = "PR1"

    Dim batch As CBatchSectionCalculator
    Set batch = TestBuildUiBatch()

    Dim reader As CLoadCombinationReader
    Set reader = New CLoadCombinationReader
    reader.LoadFromWorkbook ThisWorkbook, batch
    batch.Execute

    TestAssertTrue stats, "ui.loads.blankMoment.count", batch.Count = 1
    TestAssertClose stats, "ui.loads.blankMoment.myZero", batch.UserMy(1), 0#, 0.0000001
    TestAssertTrue stats, "ui.loads.blankMoment.valid", InStr(1, batch.ResultAt(1).Status, "InputErr", vbTextCompare) = 0
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Запускает полный workbook-сценарий круга и сверяет сохраненные таблицы,
' шапки, плоскость НДС, units и отсутствие дублированной геометрии элементов.
Private Sub TestCircleWorkbookRunWritesResults(ByRef stats As TUiTestStats)
    TestPrepareCircleInput
    Dim message As String
    message = TestFixtureRunSectionWorkflow(ThisWorkbook, False)

    TestAssertTrue stats, "ui.run.message", InStr(1, message, "Расчет завершен", vbTextCompare) > 0
    Dim summary As Object
    Set summary = ThisWorkbook.Names.Item("rngBatchSummary").RefersToRange
    Dim firstDataRow As Long
    firstDataRow = TestBatchSummaryStartRow() + 12
    TestAssertTrue stats, "ui.run.capacity.na", CStr(summary.Worksheet.Cells.Item(firstDataRow, 7).Value2) = "N/A"
    TestAssertTrue stats, "ui.run.direct.status", Len(CStr(summary.Worksheet.Cells.Item(firstDataRow, 6).Value2)) > 0
    Dim strengthAnchor As Object
    Set strengthAnchor = ThisWorkbook.Names.Item("rngStrengthSummaryAnchor").RefersToRange
    TestAssertTrue stats, "ui.run.strainPlane", IsNumeric(strengthAnchor.Worksheet.Cells.Item(strengthAnchor.Row, 15).Value2) And _
        IsNumeric(strengthAnchor.Worksheet.Cells.Item(strengthAnchor.Row, 16).Value2) And _
        IsNumeric(strengthAnchor.Worksheet.Cells.Item(strengthAnchor.Row, 17).Value2)
    Dim crackAnchor As Object
    Set crackAnchor = ThisWorkbook.Names.Item("rngCrackSummaryAnchor").RefersToRange
    TestAssertTrue stats, "ui.run.crack", Len(CStr(crackAnchor.Worksheet.Cells.Item(crackAnchor.Row, 3).Value2)) > 0 And _
        Len(CStr(crackAnchor.Worksheet.Cells.Item(crackAnchor.Row, 18).Value2)) > 0
    Dim sys As Object
    Set sys = ThisWorkbook.Worksheets.Item("Config")
    TestAssertTrue stats, "ui.run.system.noLegacyMainInput", Not TestWorkbookNameExists("rngMainInput")
    TestAssertTrue stats, "ui.run.system.materialDiagramControls", _
        TestMaterialDiagramControlsHeadingExists(sys)
    Dim resultsSheet As Object
    Set resultsSheet = ThisWorkbook.Worksheets.Item("Results")
    TestAssertTrue stats, "ui.strengthSummary.currentDepths", IsNumeric(resultsSheet.Cells.Item(strengthAnchor.Row, 26).Value2) And IsNumeric(resultsSheet.Cells.Item(strengthAnchor.Row, 27).Value2)
    TestAssertTrue stats, "ui.strengthSummary.direct.noCapacityDepths", Len(CStr(resultsSheet.Cells.Item(strengthAnchor.Row, 46).Value2)) = 0 And Len(CStr(resultsSheet.Cells.Item(strengthAnchor.Row, 47).Value2)) = 0
    TestAssertTrue stats, "ui.results.elements.header", CStr(ThisWorkbook.Names.Item("rngNDMElementResults").RefersToRange.Value2) = "RunID"
    Dim elementResults As Variant
    elementResults = TestResultTable("rngNDMElementResults")
    TestAssertTrue stats, "ui.results.elements.rows", UBound(elementResults, 1) > 1
    TestAssertTrue stats, "ui.results.elements.noCombinationIndex", TestResultHeaderColumn(elementResults, "CombinationIndex") = 0
    TestAssertTrue stats, "ui.results.elements.materialType", TestResultHeaderColumn(elementResults, "MaterialType") > 0
    TestAssertTrue stats, "ui.results.elements.units", TestResultHeaderColumn(elementResults, "Stress, MPa") > 0
    TestAssertTrue stats, "ui.results.elements.noGeometryDup", TestResultHeaderColumn(elementResults, "X, mm") = 0
    TestAssertTrue stats, "ui.results.elements.noPlaneDup", TestResultHeaderColumn(elementResults, "Epsilon0") = 0
    TestAssertTrue stats, "ui.results.elements.physicalStateAlign", _
        ThisWorkbook.Names.Item("rngNDMElementResults").RefersToRange.Offset(1, TestResultHeaderColumn(elementResults, "PhysicalState") - 1).HorizontalAlignment = -4152
    TestAssertTrue stats, "ui.results.geometry.header", CStr(ThisWorkbook.Names.Item("rngNDMSectionGeometry").RefersToRange.Value2) = "RunID"
    Dim geometryResults As Variant
    geometryResults = TestResultTable("rngNDMSectionGeometry")
    TestAssertTrue stats, "ui.results.geometry.rows", UBound(geometryResults, 1) > 1
    TestAssertTrue stats, "ui.results.strength.anchor", ThisWorkbook.Names.Item("rngStrengthSummaryAnchor").RefersToRange.Row = 49 And ThisWorkbook.Names.Item("rngStrengthSummaryAnchor").RefersToRange.Column = 1
    TestAssertTrue stats, "ui.results.crack.anchor", ThisWorkbook.Names.Item("rngCrackSummaryAnchor").RefersToRange.Row = 85 And ThisWorkbook.Names.Item("rngCrackSummaryAnchor").RefersToRange.Column = 1
    TestAssertTrue stats, "ui.results.stability.anchor", ThisWorkbook.Names.Item("rngStabilitySummaryAnchor").RefersToRange.Row = 122 And ThisWorkbook.Names.Item("rngStabilitySummaryAnchor").RefersToRange.Column = 1
    TestAssertTrue stats, "ui.results.geometry.position", ThisWorkbook.Names.Item("rngNDMSectionGeometry").RefersToRange.Row = 156 And ThisWorkbook.Names.Item("rngNDMSectionGeometry").RefersToRange.Column = 1
    TestAssertTrue stats, "ui.results.contours.position", ThisWorkbook.Names.Item("rngNDMSectionContours").RefersToRange.Row = 156 And ThisWorkbook.Names.Item("rngNDMSectionContours").RefersToRange.Column = 51
    TestAssertTrue stats, "ui.results.properties.position", ThisWorkbook.Names.Item("rngNDMSectionProperties").RefersToRange.Row = 156 And ThisWorkbook.Names.Item("rngNDMSectionProperties").RefersToRange.Column = 19
    TestAssertTrue stats, "ui.results.materialDiagrams.position", ThisWorkbook.Names.Item("rngNDMMaterialDiagrams").RefersToRange.Row = 156 And ThisWorkbook.Names.Item("rngNDMMaterialDiagrams").RefersToRange.Column = 27
    TestAssertTrue stats, "ui.results.annotations.position", ThisWorkbook.Names.Item("rngNDMSectionAnnotations").RefersToRange.Row = 156 And ThisWorkbook.Names.Item("rngNDMSectionAnnotations").RefersToRange.Column = 70
    TestAssertTrue stats, "ui.results.geometry.noSource", TestResultHeaderColumn(geometryResults, "SourceName") = 0
    TestAssertTrue stats, "ui.results.geometry.noMaterialClass", TestResultHeaderColumn(geometryResults, "MaterialClass") = 0
    TestAssertTrue stats, "ui.results.properties.header", CStr(ThisWorkbook.Names.Item("rngNDMSectionProperties").RefersToRange.Value2) = "RunID"
    TestAssertTrue stats, "ui.results.properties.hasStateEpsilon0", TestResultsPropertyExists("LC1", "State.CrackedState.Epsilon0")
    TestAssertTrue stats, "ui.results.properties.hasStateNeutralLine", _
        TestResultsPropertyExists("LC1", "State.CrackedState.NeutralLine.Angle") And _
        TestResultsPropertyExists("LC1", "State.CrackedState.NeutralLine.SignedDistance")
    TestAssertTrue stats, "ui.results.properties.noLegacyPlane", _
        Not TestResultsPropertyExists("LC1", "Epsilon0") And _
        Not TestResultsPropertyExists("LC1", "KappaX") And _
        Not TestResultsPropertyExists("LC1", "KappaY")
    TestAssertTrue stats, "ui.results.properties.hasBounds", TestResultsPropertyExists("ALL", "Bounds.MinX")
    TestAssertTrue stats, "ui.results.properties.commonLoadReference", _
        TestResultsPropertyExists("ALL", "LoadReferenceX") And TestResultsPropertyExists("ALL", "LoadReferenceY")
    TestAssertTrue stats, "ui.results.properties.hasTransformedAxes", _
        TestResultsPropertyExists("ALL", "Transformed.CentroidX") And _
        TestResultsPropertyExists("ALL", "Transformed.CentroidY") And _
        TestResultsPropertyExists("ALL", "Transformed.PrincipalAngle") And _
        TestResultsPropertyExists("ALL", "Transformed.PrincipalI1") And _
        TestResultsPropertyExists("ALL", "Transformed.PrincipalI2")
    TestAssertTrue stats, "ui.results.properties.noLegacyTransformedAliases", _
        Not TestResultsPropertyExists("ALL", "CentroidX") And _
        Not TestResultsPropertyExists("ALL", "CentroidY") And _
        Not TestResultsPropertyExists("ALL", "PrincipalAngle") And _
        Not TestResultsPropertyExists("ALL", "PrincipalI1") And _
        Not TestResultsPropertyExists("ALL", "PrincipalI2")
    TestAssertTrue stats, "ui.results.properties.noLcLoadReference", _
        Not TestResultsPropertyExists("LC1", "LoadReferenceX") And Not TestResultsPropertyExists("LC1", "LoadReferenceY")
    TestAssertTransformedAreaUsesElasticModuli stats
    TestAssertTrue stats, "ui.results.annotations.header", CStr(ThisWorkbook.Names.Item("rngNDMSectionAnnotations").RefersToRange.Value2) = "RunID"
    TestAssertTrue stats, "ui.results.annotations.rows", TestResultTableRowCount("rngNDMSectionAnnotations") > 1
    TestAssertTrue stats, "ui.results.materialDiagrams.header", CStr(ThisWorkbook.Names.Item("rngNDMMaterialDiagrams").RefersToRange.Value2) = "RunID"
    Dim materialDiagrams As Variant
    materialDiagrams = TestResultTable("rngNDMMaterialDiagrams")
    TestAssertTrue stats, "ui.results.materialDiagrams.rows", UBound(materialDiagrams, 1) > 1
    TestAssertTrue stats, "ui.results.materialDiagrams.stateType", TestResultHeaderColumn(materialDiagrams, "StateType") > 0
    TestAssertTrue stats, "ui.results.materialDiagrams.role", TestResultHeaderColumn(materialDiagrams, "MaterialModelRole") > 0
    TestAssertTrue stats, "ui.results.materialDiagrams.diagramId", TestResultHeaderColumn(materialDiagrams, "DiagramId") > 0
    TestAssertTrue stats, "ui.results.materialDiagrams.spec", TestResultHeaderColumn(materialDiagrams, "MaterialModelSpec") > 0
    TestAssertTrue stats, "ui.results.materialDiagrams.mode", TestResultHeaderColumn(materialDiagrams, "DiagramMode") > 0
    TestAssertTrue stats, "ui.results.materialDiagrams.noPurpose", TestResultHeaderColumn(materialDiagrams, "Purpose") = 0
    TestAssertTrue stats, "ui.results.materialDiagrams.noLoadCase", TestResultHeaderColumn(materialDiagrams, "LoadCase") = 0
    TestAssertTrue stats, "ui.results.materialDiagrams.usedStates", _
        TestMaterialDiagramIdsMatchStateProperties(materialDiagrams)
    TestAssertTrue stats, "ui.plot.chart.created", TestPlotChartExists()
    TestAssertTrue stats, "ui.plot.title.comment", TestPlotVisibleTitleContains("(ui test)")
    TestAssertTrue stats, "ui.plot.title.overlay", _
        TestPlotShapeTextContains("Схема сечения") And TestPlotShapeTextContains("(ui test)")

    Dim profiles As Object
    Set profiles = ThisWorkbook.Names.Item("rngCalculationProfiles").RefersToRange
    TestAssertTrue stats, "ui.profiles.topHeader.noDuplicateParameter", Len(Trim$(CStr(profiles.Cells.Item(1, 1).Value2))) = 0
    TestAssertTrue stats, "ui.profiles.lowerHeader.parameter", CStr(profiles.Cells.Item(2, 1).Value2) = "Параметр"
    TestAssertTrue stats, "ui.profiles.pr3.default", _
        Len(TestProfileSettingValue("PR3", "Profile.DisplayName")) > 0 And _
        TestProfileSettingValue("PR3", "Calculation.Strength.DirectState") = "Yes" And _
        TestProfileSettingValue("PR3", "Calculation.Crack.Width") = "Yes"
    TestAssertTrue stats, "ui.profiles.pr4.default", _
        Len(TestProfileSettingValue("PR4", "Profile.DisplayName")) > 0 And _
        TestProfileSettingValue("PR4", "Visualization.Quantity") = "Strain" And _
        TestProfileSettingValue("PR4", "Visualization.StrainPrecision") = "6"
    TestAssertTrue stats, "ui.profiles.visualizationStressPrecision.default", _
        TestProfileSettingValue("PR1", "Visualization.StressPrecision") = "1"
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет полный предельный snapshot: 30 сочетаний, каждое с пятью
' конечными named-state. Для устойчивого получения PreCrackState/PostCrackState
' используется чистый изгиб: сжатие может подавить образование нормальной
' трещины и тогда эти состояния физически не обязаны появляться.
Private Sub TestThirtyCombinationsWithFiveStatesWriteSnapshot(ByRef stats As TUiTestStats)
    TestPrepareUserRectanglesMomentUltimateInput
    TestPrepareFullStateProfile "PR3"
    TestSetSystemSetting "SLS.Crack.PsiMode", "Auto"
    TestSetSystemSetting "SLS.Crack.Allowable", "0.000001"
    TestSetSystemSetting "Plot.AutoUpdateAfterCalculation", "No"
    TestSetSystemSetting "Plot.LoadCase", "LC_FULL_01"

    Dim loads As Object
    Set loads = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange
    TestClearDataRows loads

    Dim rowIndex As Long
    For rowIndex = 1 To 30
        loads.Cells.Item(rowIndex + 1, 1).Value2 = "LC_FULL_" & Format$(rowIndex, "00")
        loads.Cells.Item(rowIndex + 1, 2).Value2 = 0#
        loads.Cells.Item(rowIndex + 1, 3).Value2 = 50#
        loads.Cells.Item(rowIndex + 1, 4).ClearContents
        loads.Cells.Item(rowIndex + 1, 5).Value2 = "PR3"
        loads.Cells.Item(rowIndex + 1, 6).Value2 = ChrW$(&H3BB) & "*Mx"
        loads.Cells.Item(rowIndex + 1, 7).Value2 = "full snapshot " & CStr(rowIndex)
    Next rowIndex

    Dim message As String
    message = TestFixtureRunSectionWorkflow(ThisWorkbook, False)

    TestAssertTrue stats, "ui.results.fullSnapshot.message", _
        InStr(1, message, "Расчет завершен", vbTextCompare) > 0

    Dim geometryRowCount As Long
    geometryRowCount = TestResultTableRowCount("rngNDMSectionGeometry") - 1
    TestAssertTrue stats, "ui.results.fullSnapshot.geometry143", geometryRowCount = 143

    Dim elementResults As Variant
    elementResults = TestResultTable("rngNDMElementResults")
    Dim expectedElementRows As Long
    expectedElementRows = 1 + 30 * 5 * geometryRowCount
    TestAssertTrue stats, "ui.results.fullSnapshot.elementRows", _
        UBound(elementResults, 1) = expectedElementRows

    TestAssertElementStateRows stats, elementResults, "LC_FULL_01", "StrengthState", geometryRowCount
    TestAssertElementStateRows stats, elementResults, "LC_FULL_01", "CapacityState", geometryRowCount
    TestAssertElementStateRows stats, elementResults, "LC_FULL_01", "CrackedState", geometryRowCount
    TestAssertElementStateRows stats, elementResults, "LC_FULL_01", "PreCrackState", geometryRowCount
    TestAssertElementStateRows stats, elementResults, "LC_FULL_01", "PostCrackState", geometryRowCount
    TestAssertElementStateRows stats, elementResults, "LC_FULL_30", "StrengthState", geometryRowCount
    TestAssertElementStateRows stats, elementResults, "LC_FULL_30", "CapacityState", geometryRowCount
    TestAssertElementStateRows stats, elementResults, "LC_FULL_30", "CrackedState", geometryRowCount
    TestAssertElementStateRows stats, elementResults, "LC_FULL_30", "PreCrackState", geometryRowCount
    TestAssertElementStateRows stats, elementResults, "LC_FULL_30", "PostCrackState", geometryRowCount

    TestAssertTrue stats, "ui.results.fullSnapshot.propertiesRows", _
        UBound(TestResultTable("rngNDMSectionProperties"), 1) >= 1 + 55 + 30 * (23 + 5 * 10)
    Dim common As Variant, commonOnly As Boolean, commonRow As Long
    common = ReadPhysicalResultTable(ThisWorkbook, "rngNDMSectionProperties"): commonOnly = True
    For commonRow = 2 To UBound(common, 1)
        If StrComp(CStr(common(commonRow, 2)), "ALL", vbTextCompare) <> 0 Then commonOnly = False
    Next commonRow
    TestAssertTrue stats, "ui.results.fullSnapshot.readableLoadCaseProperties", Not commonOnly
    TestAssertTrue stats, "ui.results.fullSnapshot.materialRows", _
        TestResultTableRowCount("rngNDMMaterialDiagrams") > 1
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет, что высота rngLoadCombinations управляет числом LC, а расчет
' заранее останавливается, если текущие якоря Results не оставляют места для
' всех строк, шапок и выводимых столбцов.
Private Sub TestDynamicLoadCombinationRangeAndLayoutGuard(ByRef stats As TUiTestStats)
    On Error GoTo Failed

    TestPrepareCircleInput
    TestSetSystemSetting "Plot.AutoUpdateAfterCalculation", "No"

    Dim originalRefersTo As String
    originalRefersTo = ThisWorkbook.Names.Item("rngLoadCombinations").RefersTo

    Dim app As Object
    Set app = ThisWorkbook.Application
    Dim oldDisplayAlerts As Boolean
    oldDisplayAlerts = app.DisplayAlerts
    app.DisplayAlerts = False

    On Error Resume Next
    ThisWorkbook.Worksheets.Item("__tmpDynamicLoads").Delete
    On Error GoTo Failed

    Dim tempSheet As Object
    Set tempSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets.Item(ThisWorkbook.Worksheets.Count))
    tempSheet.Name = "__tmpDynamicLoads"

    Dim tempRange As Object
    Set tempRange = tempSheet.Range("A1:G46")
    FillLoadCombinationTestRange tempRange, 45, "PR1", "LC_DYN_"
    ThisWorkbook.Names.Item("rngLoadCombinations").RefersTo = "=" & tempRange.Address(True, True, 1, True)

    Dim batch As CBatchSectionCalculator
    Set batch = TestBuildUiBatch()
    Dim reader As CLoadCombinationReader
    Set reader = New CLoadCombinationReader
    reader.LoadFromWorkbook ThisWorkbook, batch
    TestAssertTrue stats, "ui.loads.dynamicRange.count45", batch.Count = 45

    Dim errorNumber As Long
    Dim errorText As String
    On Error Resume Next
    Dim message As String
    message = TestFixtureRunSectionWorkflow(ThisWorkbook, False)
    errorNumber = Err.Number
    errorText = Err.Description
    Err.Clear
    On Error GoTo Failed

    TestAssertTrue stats, "ui.results.layoutGuard.blocksCalculation", errorNumber <> 0
    TestAssertTrue stats, "ui.results.layoutGuard.message", _
        InStr(1, errorText, "не помещаются", vbTextCompare) > 0 And _
        InStr(1, errorText, "Расчет не запущен", vbTextCompare) > 0 And _
        InStr(1, errorText, "Что сделать", vbTextCompare) > 0 And _
        InStr(1, errorText, "Вставьте", vbTextCompare) > 0 And _
        InStr(1, errorText, "на листе Results", vbTextCompare) > 0 And _
        InStr(1, errorText, "над строкой", vbTextCompare) > 0 And _
        InStr(1, errorText, "над строкой 44", vbTextCompare) > 0 And _
        InStr(1, errorText, "над строкой 80", vbTextCompare) > 0 And _
        InStr(1, errorText, "над строкой 116", vbTextCompare) > 0 And _
        InStr(1, errorText, "над строкой 45", vbTextCompare) = 0 And _
        InStr(1, errorText, "над строкой 81", vbTextCompare) = 0 And _
        InStr(1, errorText, "над строкой 117", vbTextCompare) = 0 And _
        InStr(1, errorText, "на листе Config", vbTextCompare) > 0 And _
        InStr(1, errorText, vbCrLf & "1)", vbBinaryCompare) > 0 And _
        InStr(1, errorText, vbCrLf & "2)", vbBinaryCompare) > 0 And _
        InStr(1, errorText, "rngStrengthSummaryAnchor", vbTextCompare) > 0 And _
        InStr(1, errorText, "rngCrackSummaryAnchor", vbTextCompare) > 0 And _
        InStr(1, errorText, "rngStabilitySummaryAnchor", vbTextCompare) > 0 And _
        InStr(1, errorText, "rngLoadCombinations", vbTextCompare) > 0

CleanUp:
    On Error Resume Next
    ThisWorkbook.Names.Item("rngLoadCombinations").RefersTo = originalRefersTo
    If Not tempSheet Is Nothing Then tempSheet.Delete
    app.DisplayAlerts = oldDisplayAlerts
    On Error GoTo 0
    Exit Sub

Failed:
    stats.Failed = stats.Failed + 1
    TestAppendLine stats, "FAIL: ui.results.layoutGuard; err=" & CStr(Err.Number) & "; " & Err.Description
    Resume CleanUp
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Заполняет временную таблицу сочетаний с тем же frontend-контрактом, что и
' rngLoadCombinations. Используется только для проверки динамической высоты
' именованного диапазона, не затрагивая рабочий блок Config.
Private Sub FillLoadCombinationTestRange(ByVal target As Object, ByVal combinationCount As Long, _
        ByVal profileId As String, ByVal idPrefix As String)
    target.ClearContents
    target.Cells.Item(1, 1).Value2 = "CombinationID"
    target.Cells.Item(1, 2).Value2 = "N"
    target.Cells.Item(1, 3).Value2 = "Mx"
    target.Cells.Item(1, 4).Value2 = "My"
    target.Cells.Item(1, 5).Value2 = "ProfileId"
    target.Cells.Item(1, 6).Value2 = "LoadPath"
    target.Cells.Item(1, 7).Value2 = "Comment"

    Dim rowIndex As Long
    For rowIndex = 1 To combinationCount
        target.Cells.Item(rowIndex + 1, 1).Value2 = idPrefix & Format$(rowIndex, "00")
        target.Cells.Item(rowIndex + 1, 2).Value2 = -100000#
        target.Cells.Item(rowIndex + 1, 3).Value2 = -1000000#
        target.Cells.Item(rowIndex + 1, 4).Value2 = 0#
        target.Cells.Item(rowIndex + 1, 5).Value2 = profileId
        target.Cells.Item(rowIndex + 1, 6).Value2 = ChrW$(&H3BB) & "*Mx"
        target.Cells.Item(rowIndex + 1, 7).Value2 = "dynamic range " & CStr(rowIndex)
    Next rowIndex
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' ДЛЯ ТЕСТОВ
' Проверяет тяжелую схему с мелкой сеткой. Расчетные строки остаются полными,
' а plotter обязан рисовать элементы поэлементно Shape-ами до безопасного
' предела и не раздувать число Excel Chart series.
Private Sub TestLargeSnapshotPlotStress(ByRef stats As TUiTestStats)
    TestPrepareUserRectanglesMomentUltimateInput
    TestPrepareFullStateProfile "PR3"
    TestSetSystemSetting "Mesh.StepMode", "User"
    TestSetSystemSetting "Mesh.StepX", "10"
    TestSetSystemSetting "Mesh.StepY", "10"
    TestSetSystemSetting "Mesh.BoundarySubdivisions", "1"
    TestSetSystemSetting "SLS.Crack.PsiMode", "Auto"
    TestSetSystemSetting "SLS.Crack.Allowable", "0.000001"
    TestSetSystemSetting "Plot.AutoUpdateAfterCalculation", "Yes"
    TestSetSystemSetting "Plot.LoadCase", "LC_BIG_01"

    Dim loads As Object
    Set loads = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange
    TestClearDataRows loads

    Dim rowIndex As Long
    For rowIndex = 1 To 20
        loads.Cells.Item(rowIndex + 1, 1).Value2 = "LC_BIG_" & Format$(rowIndex, "00")
        loads.Cells.Item(rowIndex + 1, 2).Value2 = 200#
        loads.Cells.Item(rowIndex + 1, 3).Value2 = 50#
        loads.Cells.Item(rowIndex + 1, 4).ClearContents
        loads.Cells.Item(rowIndex + 1, 5).Value2 = "PR3"
        loads.Cells.Item(rowIndex + 1, 6).Value2 = ChrW$(&H3BB) & "*Mx"
        loads.Cells.Item(rowIndex + 1, 7).Value2 = "large snapshot " & CStr(rowIndex)
    Next rowIndex

    Dim message As String
    message = TestFixtureRunSectionWorkflow(ThisWorkbook, False)

    Dim geometryRowCount As Long
    geometryRowCount = TestResultTableRowCount("rngNDMSectionGeometry") - 1

    Dim elementResults As Variant
    elementResults = TestResultTable("rngNDMElementResults")

    Dim seriesCount As Long
    seriesCount = TestPlotSeriesCount()
    Dim elementShapeCount As Long
    elementShapeCount = TestCountPlotShapes("Element")

    TestAppendLine stats, "INFO: ui.largeSnapshot.geometryRows=" & CStr(geometryRowCount) & _
        "; elementResultRows=" & CStr(UBound(elementResults, 1)) & _
        "; plotSeries=" & CStr(seriesCount) & _
        "; elementShapes=" & CStr(elementShapeCount)

    TestAssertTrue stats, "ui.largeSnapshot.message", _
        InStr(1, message, "Расчет завершен", vbTextCompare) > 0
    TestAssertTrue stats, "ui.largeSnapshot.geometryLarge", geometryRowCount > 2500
    TestAssertTrue stats, "ui.largeSnapshot.elementRows", _
        UBound(elementResults, 1) = 1 + 20 * 5 * geometryRowCount
    TestAssertTrue stats, "ui.largeSnapshot.plotCreated", TestPlotChartExists()
    TestAssertTrue stats, "ui.largeSnapshot.elementShapes", elementShapeCount >= geometryRowCount
    TestAssertTrue stats, "ui.largeSnapshot.plotSeriesLimit", seriesCount < 256
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Запускает Rectangles через пользовательский workbook-entrypoint и проверяет
' фактические Results, блок диаграмм и независимую таблицу армирования этого типа.
Private Sub TestRectanglesWorkbookRunWritesResults(ByRef stats As TUiTestStats)
    TestPrepareRectanglesInput
    Dim message As String
    message = TestFixtureRunSectionWorkflow(ThisWorkbook, False)

    TestAssertTrue stats, "ui.rectangles.message", InStr(1, message, "завершен", vbTextCompare) > 0
    Dim summary As Object
    Set summary = ThisWorkbook.Names.Item("rngBatchSummary").RefersToRange
    TestAssertTrue stats, "ui.rectangles.result.status", Len(CStr(summary.Worksheet.Cells.Item(TestBatchSummaryStartRow() + 12, 1).Value2)) > 0

    Dim sys As Object
    Set sys = ThisWorkbook.Worksheets.Item("Config")
    TestAssertTrue stats, "ui.rectangles.system.perSectionInput", Not TestWorkbookNameExists("rngRebarInput")
    TestAssertTrue stats, "ui.rectangles.system.materialDiagramControls", _
        TestMaterialDiagramControlsHeadingExists(sys)
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет пользовательский сценарий Г-сечения с N + Mx и выбранной
' траекторией lambda*Mx через полный путь книги. Этот тест защищает быстрый
' UltimateStrain-путь для обычного изгибного расчета: после универсализации
' LoadPath он не должен уходить в тяжелую общую residual-систему.
Private Sub TestRectanglesMomentUltimateStrainWorkbookPath(ByRef stats As TUiTestStats)
    TestPrepareUserRectanglesMomentUltimateInput

    Dim message As String
    message = TestFixtureRunSectionWorkflow(ThisWorkbook, False)
    ThisWorkbook.Application.CalculateFull

    Dim resultsSheet As Object
    Set resultsSheet = ThisWorkbook.Worksheets.Item("Results")
    Dim firstRow As Long
    firstRow = TestStrengthSummaryStartRow()

    Dim capacityStatus As String
    Dim solutionMethod As String
    capacityStatus = CStr(resultsSheet.Cells.Item(firstRow, 49).Value2)
    solutionMethod = CStr(resultsSheet.Cells.Item(firstRow, 37).Value2)

    TestAppendLine stats, "INFO: ui.rectangles.momentUltimate capacityStatus=" & capacityStatus & _
        "; solutionMethod=" & solutionMethod & _
        "; lambda=" & CStr(resultsSheet.Cells.Item(firstRow, 38).Value2)

    TestAssertTrue stats, "ui.rectangles.momentUltimate.message", InStr(1, message, "Расчет завершен", vbTextCompare) > 0
    TestAssertTextEquals stats, "ui.rectangles.momentUltimate.capacityOk", capacityStatus, "OK"
    TestAssertTextEquals stats, "ui.rectangles.momentUltimate.method", solutionMethod, "UltimateStrain"
    TestAssertTrue stats, "ui.rectangles.momentUltimate.lambda", _
        CDbl(resultsSheet.Cells.Item(firstRow, 38).Value2) > 0#
    TestAssertTrue stats, "ui.rectangles.momentUltimate.mxult", _
        Abs(CDbl(resultsSheet.Cells.Item(firstRow, 40).Value2)) > 0#
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет, что подробный блок прочности не держит кривизны в 1/мм
' жестко, а использует OUTPUT-единицу из Units.Curvature.Output.
Private Sub TestStrengthSummaryUsesOutputCurvatureUnit(ByRef stats As TUiTestStats)
    Dim oldOutputCurvature As String
    oldOutputCurvature = TestGetSystemSetting("Units.Curvature.Output")

    On Error GoTo RestoreAndFail
    TestPrepareUserRectanglesMomentUltimateInput
    TestSetSystemSetting "Units.Curvature.Output", "1/m"

    Dim message As String
    message = TestFixtureRunSectionWorkflow(ThisWorkbook, False)
    ThisWorkbook.Application.CalculateFull

    Dim resultsSheet As Object
    Set resultsSheet = ThisWorkbook.Worksheets.Item("Results")
    Dim firstRow As Long
    firstRow = TestStrengthSummaryStartRow()
    Dim labelRow As Long
    labelRow = firstRow - 1

    TestAssertTrue stats, "ui.strength.curvatureOutput.message", _
        InStr(1, message, "Расчет завершен", vbTextCompare) > 0
    TestAssertTextEquals stats, "ui.strength.curvatureOutput.snapshotUnit", _
        TestResultsPropertyValue("ALL", "Output.CurvatureUnit"), "1/m"
    TestAssertTrue stats, "ui.strength.curvatureOutput.stateHeaderKx", _
        InStr(1, CStr(resultsSheet.Cells.Item(labelRow, 16).Value2), "1/m", vbTextCompare) > 0
    TestAssertTrue stats, "ui.strength.curvatureOutput.stateHeaderKy", _
        InStr(1, CStr(resultsSheet.Cells.Item(labelRow, 17).Value2), "1/m", vbTextCompare) > 0
    TestAssertTrue stats, "ui.strength.curvatureOutput.capacityHeaderKx", _
        InStr(1, CStr(resultsSheet.Cells.Item(labelRow, 33).Value2), "1/m", vbTextCompare) > 0
    TestAssertTrue stats, "ui.strength.curvatureOutput.capacityHeaderKy", _
        InStr(1, CStr(resultsSheet.Cells.Item(labelRow, 34).Value2), "1/m", vbTextCompare) > 0

    TestAssertClose stats, "ui.strength.curvatureOutput.stateKxValue", _
        CDbl(resultsSheet.Cells.Item(firstRow, 16).Value2), _
        CDbl(TestResultsPropertyValue("LC_MX", "State.StrengthState.KappaX")), 0.000000000001
    TestAssertClose stats, "ui.strength.curvatureOutput.capacityKxValue", _
        CDbl(resultsSheet.Cells.Item(firstRow, 33).Value2), _
        CDbl(TestResultsPropertyValue("LC_MX", "State.CapacityState.KappaX")), 0.000000000001

Restore:
    TestSetSystemSetting "Units.Curvature.Output", oldOutputCurvature
    Exit Sub

RestoreAndFail:
    stats.Failed = stats.Failed + 1
    TestAppendLine stats, "FAIL: ui.strength.curvatureOutput; " & Err.Description
    Resume Restore
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет чистый изгиб Г-сечения по полному Excel-пути.
' Здесь N намеренно равен нулю: программа должна найти прямое НДС и
' предельный момент без перехода в осевую или аварийную ветку.
Private Sub TestRectanglesPureBendingUltimateStrainWorkbookPath(ByRef stats As TUiTestStats)
    TestPrepareUserRectanglesMomentUltimateInput

    Dim loads As Object
    Set loads = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange
    loads.Cells.Item(2, 2).Value2 = 0#
    loads.Cells.Item(2, 7).Value2 = "pure bending ultimate regression"

    Dim message As String
    message = TestFixtureRunSectionWorkflow(ThisWorkbook, False)
    ThisWorkbook.Application.CalculateFull

    Dim resultsSheet As Object
    Set resultsSheet = ThisWorkbook.Worksheets.Item("Results")
    Dim firstRow As Long
    firstRow = TestStrengthSummaryStartRow()

    Dim directStatus As String
    Dim capacityStatus As String
    Dim solutionMethod As String
    directStatus = CStr(resultsSheet.Cells.Item(firstRow, 30).Value2)
    capacityStatus = CStr(resultsSheet.Cells.Item(firstRow, 49).Value2)
    solutionMethod = CStr(resultsSheet.Cells.Item(firstRow, 37).Value2)

    TestAppendLine stats, "INFO: ui.rectangles.pureBending direct=" & directStatus & _
        "; capacity=" & capacityStatus & _
        "; solutionMethod=" & solutionMethod & _
        "; lambda=" & CStr(resultsSheet.Cells.Item(firstRow, 38).Value2)

    TestAssertTrue stats, "ui.rectangles.pureBending.message", InStr(1, message, "Расчет завершен", vbTextCompare) > 0
    TestAssertTextEquals stats, "ui.rectangles.pureBending.directOk", directStatus, "OK"
    TestAssertTextEquals stats, "ui.rectangles.pureBending.capacityOk", capacityStatus, "OK"
    TestAssertTextEquals stats, "ui.rectangles.pureBending.method", solutionMethod, "UltimateStrain"
    TestAssertTrue stats, "ui.rectangles.pureBending.lambda", _
        CDbl(resultsSheet.Cells.Item(firstRow, 38).Value2) > 0#
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет чистый изгиб Г-сечения по профилю PR1. Такой профиль запрашивает
' и прямое НДС, и несущую способность, поэтому обе ветви должны проходить
' через полный workbook-path без специальных обходов.
Private Sub TestRectanglesPureBendingDirectStateWorkbookPath(ByRef stats As TUiTestStats)
    TestPrepareUserRectanglesMomentUltimateInput

    Dim loads As Object
    Set loads = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange
    loads.Cells.Item(2, 2).Value2 = 0#
    loads.Cells.Item(2, 7).Value2 = "pure bending direct regression"

    Dim message As String
    message = TestFixtureRunSectionWorkflow(ThisWorkbook, False)
    ThisWorkbook.Application.CalculateFull

    Dim resultsSheet As Object
    Set resultsSheet = ThisWorkbook.Worksheets.Item("Results")
    Dim firstRow As Long
    firstRow = TestStrengthSummaryStartRow()

    Dim directStatus As String
    Dim capacityStatus As String
    directStatus = CStr(resultsSheet.Cells.Item(firstRow, 30).Value2)
    capacityStatus = CStr(resultsSheet.Cells.Item(firstRow, 49).Value2)

    TestAppendLine stats, "INFO: ui.rectangles.pureBendingDirect direct=" & directStatus & _
        "; capacity=" & capacityStatus

    TestAssertTrue stats, "ui.rectangles.pureBendingDirect.message", InStr(1, message, "Расчет завершен", vbTextCompare) > 0
    TestAssertTextEquals stats, "ui.rectangles.pureBendingDirect.directOk", directStatus, "OK"
    TestAssertTextEquals stats, "ui.rectangles.pureBendingDirect.capacityOk", capacityStatus, "OK"
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет пользовательский сценарий с сильным осевым растяжением Г-сечения
' через настоящий workbook-path: Config -> CUnitSystem -> batch -> Results.
' Это важно, потому что знак N и единицы tf здесь проходят ровно тем же путем,
' что и при нажатии кнопки "Расчет" в книге.
Private Sub TestRectanglesAxialTensionExtensionFromWorkbookSettings(ByRef stats As TUiTestStats)
    TestPrepareUserRectanglesAxialTensionInput

    Dim reportPath As String
    reportPath = ThisWorkbook.Path & "\RC_Section_NDM_execution_report.txt"
    TestDeleteFileIfExists reportPath

    Dim message As String
    message = TestFixtureRunSectionWorkflow(ThisWorkbook, False)

    Dim resultsSheet As Object
    Set resultsSheet = ThisWorkbook.Worksheets.Item("Results")
    Dim safeSummaryRow As Long
    Dim overSummaryRow As Long
    Dim safeCrackRow As Long
    Dim overCrackRow As Long
    safeSummaryRow = TestBatchSummaryRowByCombination(resultsSheet, "LC_SAFE")
    overSummaryRow = TestBatchSummaryRowByCombination(resultsSheet, "LC_OVER")
    safeCrackRow = TestCrackSummaryRowByCombination(resultsSheet, "LC_SAFE")
    overCrackRow = TestCrackSummaryRowByCombination(resultsSheet, "LC_OVER")

    Dim safeOverall As String
    Dim safeCrack As String
    Dim overOverall As String
    Dim overCrack As String
    Dim overCrackExtUsed As String
    Dim overCrackEquilibrium As String
    Dim overLongitudinal As String
    Dim overExtension As String
    safeOverall = CStr(resultsSheet.Cells.Item(safeSummaryRow, 4).Value2)
    safeCrack = CStr(resultsSheet.Cells.Item(safeCrackRow, 3).Value2)
    overOverall = CStr(resultsSheet.Cells.Item(overSummaryRow, 4).Value2)
    overCrackExtUsed = CStr(resultsSheet.Cells.Item(overCrackRow, 28).Value2)
    overCrackEquilibrium = CStr(resultsSheet.Cells.Item(overCrackRow, 29).Value2)
    overCrack = CStr(resultsSheet.Cells.Item(overCrackRow, 67).Value2)
    overLongitudinal = CStr(resultsSheet.Cells.Item(overCrackRow, 72).Value2)
    overExtension = TestResultsPropertyValue("LC_OVER", "ExtensionUsed")

    TestAppendLine stats, "INFO: ui.rectangles.axial795 overall=" & safeOverall & _
        "; crack=" & safeCrack
    TestAppendLine stats, "INFO: ui.rectangles.axial900 overall=" & overOverall & _
        "; crack=" & overCrack & _
        "; extensionUsed=" & overExtension

    TestAssertTrue stats, "ui.rectangles.axial795.overallCalculated", _
        safeOverall = "OK" Or safeOverall = "FAIL"
    TestAssertTrue stats, "ui.rectangles.axial795.crackCalculated", _
        safeCrack = "OK" Or safeCrack = "FAIL"

    TestAssertTrue stats, "ui.rectangles.axial900.finishedWithoutExtension", _
        overOverall = "FAIL" Or overOverall = "NumFail"
    TestAssertTrue stats, "ui.rectangles.axial900.crackNoNumFail", _
        overCrack <> "NumFail" And overCrack <> "InputErr"
    TestAssertTextEquals stats, "ui.rectangles.axial900.crackedStateExtUsed", _
        overCrackExtUsed, "yes"
    TestAssertTextEquals stats, "ui.rectangles.axial900.crackedStateEquilibrium", _
        overCrackEquilibrium, "FAIL"
    TestAssertTextEquals stats, "ui.rectangles.axial900.crackWidthSkipped", _
        overCrack, "N/A"
    TestAssertTextEquals stats, "ui.rectangles.axial900.longitudinalSkipped", _
        overLongitudinal, "N/A"
    TestAssertTrue stats, "ui.rectangles.axial900.message", InStr(1, message, "Расчет завершен", vbTextCompare) > 0
    TestAssertTextEquals stats, "ui.rectangles.axial900.extensionSnapshot", overExtension, "True"
    TestAssertTrue stats, "ui.rectangles.axial900.crackedStateExtensionDiagram", _
        TestMaterialDiagramHasMode("CrackedState", "StateExtension")
    TestAssertTrue stats, "ui.rectangles.axial900.reportCreated", TestFileExists(reportPath)

    TestSetSystemSetting "General.ExecutionReportEnabled", "No"
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Решает безопасное и более тяжелое LC с критерием governing StrengthCapacity.
' Сверяет выбранный ID с запасами в настоящей сводке и итоговым сообщением.
Private Sub TestGoverningCombinationWritesDetailedResults(ByRef stats As TUiTestStats)
    TestPrepareCircleInput
    TestSetSystemSetting "General.WorstCombinationCriterion", "StrengthCapacity"
    TestSetSystemSetting "Capacity.SolutionStrategy", "Auto"
    TestSetSystemSetting "Capacity.MaxLambda", "1024"
    TestSetSystemSetting "Plot.LoadCase", "Worst"
    TestSetProfileSetting "PR1", "Calculation.Strength.DirectState", "Yes"
    TestSetProfileSetting "PR1", "Calculation.Strength.Capacity", "Yes"
    TestSetProfileSetting "PR1", "Calculation.Crack.Width", "No"
    TestSetProfileSetting "PR1", "Calculation.Stability.Enabled", "No"

    Dim loads As Object
    Set loads = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange
    loads.Cells.Item(2, 1).Value2 = "LC_SAFE"
    loads.Cells.Item(2, 2).Value2 = -100000#
    loads.Cells.Item(2, 3).Value2 = -1000000#
    loads.Cells.Item(2, 4).Value2 = 0#
    loads.Cells.Item(2, 5).Value2 = "PR1"
    loads.Cells.Item(2, 6).Value2 = ChrW$(&H3BB) & "*Mxy"
    loads.Cells.Item(2, 7).Value2 = "less severe"

    loads.Cells.Item(3, 1).Value2 = "LC_GOV"
    loads.Cells.Item(3, 2).Value2 = -100000#
    loads.Cells.Item(3, 3).Value2 = -8000000#
    loads.Cells.Item(3, 4).Value2 = 0#
    loads.Cells.Item(3, 5).Value2 = "PR1"
    loads.Cells.Item(3, 6).Value2 = ChrW$(&H3BB) & "*Mxy"
    loads.Cells.Item(3, 7).Value2 = "governing"

    Dim message As String
    message = TestFixtureRunSectionWorkflow(ThisWorkbook, False)
    ThisWorkbook.Application.CalculateFull

    Dim resultsSheet As Object
    Set resultsSheet = ThisWorkbook.Worksheets.Item("Results")
    Dim summaryRow As Long
    summaryRow = TestBatchSummaryStartRow()
    Dim governingID As String
    governingID = CStr(resultsSheet.Cells.Item(summaryRow, 1).Value2)
    Dim expectedID As String
    expectedID = TestExpectedGoverningByLowestStrengthSafety(resultsSheet, summaryRow)
    TestAppendLine stats, "INFO: ui.governing actual=" & governingID & "; expected=" & expectedID

    TestAssertTrue stats, "ui.governing.id", governingID = expectedID
    TestAssertTrue stats, "ui.governing.message", InStr(1, message, governingID, vbTextCompare) > 0
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Независимо выбирает LC по минимальному положительному запасу в записанном summary.
' Пустые строки пропускаются; производственный выбор governing здесь не вызывается.
Private Function TestExpectedGoverningByLowestStrengthSafety(ByVal resultsSheet As Object, ByVal summaryRow As Long) As String
    Dim rowIndex As Long
    Dim bestSafety As Double
    For rowIndex = summaryRow + 12 To summaryRow + 30
        If Len(Trim$(CStr(resultsSheet.Cells.Item(rowIndex, 1).Value2))) > 0 Then
            Dim safetyValue As Double
            safetyValue = TestStrengthSafetyForSummaryRow(resultsSheet, rowIndex)
            If safetyValue > 0# And (bestSafety = 0# Or safetyValue < bestSafety) Then
                bestSafety = safetyValue
                TestExpectedGoverningByLowestStrengthSafety = CStr(resultsSheet.Cells.Item(rowIndex, 1).Value2)
            End If
        End If
    Next rowIndex
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Читает численный запас подробной строки; пустое/текстовое поле не участвует в минимуме.
Private Function TestStrengthSafetyForSummaryRow(ByVal resultsSheet As Object, ByVal rowIndex As Long) As Double
    If IsNumeric(resultsSheet.Cells.Item(rowIndex, 17).Value2) Then _
        TestStrengthSafetyForSummaryRow = CDbl(resultsSheet.Cells.Item(rowIndex, 17).Value2)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сопоставляет LC с подробным выводом и собирает модуль из записанных Mx/My.
' Отсутствующий LC возвращает ноль, который направленный тест отличает от нужной точки.
Private Function TestMomentUltimateForCombination(ByVal resultsSheet As Object, ByVal summaryRow As Long, ByVal combinationID As String) As Double
    Dim rowIndex As Long
    For rowIndex = summaryRow + 12 To summaryRow + 30
        If StrComp(CStr(resultsSheet.Cells.Item(rowIndex, 1).Value2), combinationID, vbTextCompare) = 0 Then
            Dim mxUltimate As Double
            Dim myUltimate As Double
            Dim strengthRow As Long
            strengthRow = TestStrengthSummaryRowByCombination(resultsSheet, combinationID)
            If strengthRow > 0 Then
                If IsNumeric(resultsSheet.Cells.Item(strengthRow, 33).Value2) Then mxUltimate = CDbl(resultsSheet.Cells.Item(strengthRow, 33).Value2)
                If IsNumeric(resultsSheet.Cells.Item(strengthRow, 34).Value2) Then myUltimate = CDbl(resultsSheet.Cells.Item(strengthRow, 34).Value2)
            End If
            TestMomentUltimateForCombination = Sqr(mxUltimate * mxUltimate + myUltimate * myUltimate)
            Exit Function
        End If
    Next rowIndex
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Координаты трех summary берутся из именованных anchors, не из фиксированной верстки.
Private Function TestBatchSummaryStartRow() As Long
    TestBatchSummaryStartRow = ThisWorkbook.Names.Item("rngBatchSummary").RefersToRange.Row
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Находит фактическую строку LC в batch summary с учетом промежутков исходной таблицы.
' Ноль обозначает отсутствие ID, а не первую строку результатов.
Private Function TestBatchSummaryRowByCombination(ByVal resultsSheet As Object, ByVal combinationID As String) As Long
    Dim anchorRow As Long
    anchorRow = TestBatchSummaryStartRow()
    Dim rowIndex As Long
    For rowIndex = anchorRow + 12 To anchorRow + 200
        If StrComp(CStr(resultsSheet.Cells.Item(rowIndex, 1).Value2), combinationID, vbTextCompare) = 0 Then
            TestBatchSummaryRowByCombination = rowIndex
            Exit Function
        End If
    Next rowIndex
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает строку strength anchor для остальных readback-проверок.
Private Function TestStrengthSummaryStartRow() As Long
    TestStrengthSummaryStartRow = ThisWorkbook.Names.Item("rngStrengthSummaryAnchor").RefersToRange.Row
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает строку crack anchor для остальных readback-проверок.
Private Function TestCrackSummaryStartRow() As Long
    TestCrackSummaryStartRow = ThisWorkbook.Names.Item("rngCrackSummaryAnchor").RefersToRange.Row
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Ищет ID в подробной таблице прочности; пропуски между сочетаниями не уплотняются.
Private Function TestStrengthSummaryRowByCombination(ByVal resultsSheet As Object, ByVal combinationID As String) As Long
    Dim anchorRow As Long
    anchorRow = TestStrengthSummaryStartRow()
    Dim rowIndex As Long
    For rowIndex = anchorRow To anchorRow + 200
        If StrComp(CStr(resultsSheet.Cells.Item(rowIndex, 1).Value2), combinationID, vbTextCompare) = 0 Then
            TestStrengthSummaryRowByCombination = rowIndex
            Exit Function
        End If
    Next rowIndex
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Ищет ID в подробной таблице трещин независимо от позиции в batch summary.
Private Function TestCrackSummaryRowByCombination(ByVal resultsSheet As Object, ByVal combinationID As String) As Long
    Dim anchorRow As Long
    anchorRow = TestCrackSummaryStartRow()
    Dim rowIndex As Long
    For rowIndex = anchorRow To anchorRow + 200
        If StrComp(CStr(resultsSheet.Cells.Item(rowIndex, 1).Value2), combinationID, vbTextCompare) = 0 Then
            TestCrackSummaryRowByCombination = rowIndex
            Exit Function
        End If
    Next rowIndex
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Определяет колонку по точному тексту первой строки массива; отсутствующий заголовок дает ноль.
Private Function TestResultHeaderColumn(ByRef data As Variant, ByVal headerText As String) As Long
    Dim colIndex As Long
    For colIndex = 1 To UBound(data, 2)
        If StrComp(CStr(data(1, colIndex)), headerText, vbTextCompare) = 0 Then
            TestResultHeaderColumn = colIndex
            Exit Function
        End If
    Next colIndex
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет, что конкретное состояние конкретного LC записано для каждого
' расчетного элемента сечения. Это защищает snapshot от частичной записи.
Private Sub TestAssertElementStateRows(ByRef stats As TUiTestStats, ByRef data As Variant, _
        ByVal loadCase As String, ByVal stateType As String, ByVal expectedCount As Long)
    Dim count As Long
    count = TestElementStateRowCount(data, loadCase, stateType)
    TestAssertTrue stats, "ui.results.fullSnapshot." & loadCase & "." & stateType, count = expectedCount
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Считает строки rngNDMElementResults для пары LoadCase + StateType.
Private Function TestElementStateRowCount(ByRef data As Variant, ByVal loadCase As String, _
        ByVal stateType As String) As Long
    Dim loadCaseColumn As Long
    Dim stateTypeColumn As Long
    loadCaseColumn = TestResultHeaderColumn(data, "LoadCase")
    stateTypeColumn = TestResultHeaderColumn(data, "StateType")
    If loadCaseColumn = 0 Or stateTypeColumn = 0 Then Exit Function

    Dim rowIndex As Long
    For rowIndex = 2 To UBound(data, 1)
        If StrComp(CStr(data(rowIndex, loadCaseColumn)), loadCase, vbTextCompare) = 0 And _
                StrComp(CStr(data(rowIndex, stateTypeColumn)), stateType, vbTextCompare) = 0 Then
            TestElementStateRowCount = TestElementStateRowCount + 1
        End If
    Next rowIndex
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает максимальное по модулю Stress для сохраненного named-state.
' Тест защищает snapshot от рассинхрона: если StateSolution найден через
' numerical extension, строки элементов должны соответствовать той же диаграмме.
Private Function TestMaxAbsElementStress(ByRef data As Variant, ByVal loadCase As String, _
        ByVal stateType As String) As Double
    Dim loadCaseColumn As Long
    Dim stateTypeColumn As Long
    Dim stressColumn As Long
    loadCaseColumn = TestResultHeaderColumn(data, "LoadCase")
    stateTypeColumn = TestResultHeaderColumn(data, "StateType")
    stressColumn = TestResultHeaderColumnByBaseName(data, "Stress")
    If loadCaseColumn = 0 Or stateTypeColumn = 0 Or stressColumn = 0 Then Exit Function

    Dim rowIndex As Long
    Dim stressValue As Double
    For rowIndex = 2 To UBound(data, 1)
        If StrComp(CStr(data(rowIndex, loadCaseColumn)), loadCase, vbTextCompare) = 0 And _
                StrComp(CStr(data(rowIndex, stateTypeColumn)), stateType, vbTextCompare) = 0 Then
            stressValue = Abs(CDbl(data(rowIndex, stressColumn)))
            If stressValue > TestMaxAbsElementStress Then TestMaxAbsElementStress = stressValue
        End If
    Next rowIndex
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет, есть ли в диагностической таблице диаграмм хотя бы одна диаграмма
' заданного named-state с указанным режимом Physical/StateExtension.
Private Function TestMaterialDiagramHasMode(ByVal stateTypeText As String, ByVal diagramModeText As String) As Boolean
    Dim materialDiagrams As Variant
    materialDiagrams = TestResultTable("rngNDMMaterialDiagrams")

    Dim stateCol As Long
    Dim modeCol As Long
    stateCol = TestResultHeaderColumn(materialDiagrams, "StateType")
    modeCol = TestResultHeaderColumn(materialDiagrams, "DiagramMode")
    If stateCol = 0 Or modeCol = 0 Then Exit Function

    Dim rowIndex As Long
    For rowIndex = 2 To UBound(materialDiagrams, 1)
        If StrComp(CStr(materialDiagrams(rowIndex, stateCol)), stateTypeText, vbTextCompare) = 0 And _
                StrComp(CStr(materialDiagrams(rowIndex, modeCol)), diagramModeText, vbTextCompare) = 0 Then
            TestMaterialDiagramHasMode = True
            Exit Function
        End If
    Next rowIndex
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет, что каталог фактических диаграмм связан с named-state metadata.
' Диаграммы пишутся один раз по DiagramId, а состояния ссылаются на них через
' State.*.ConcreteDiagramId и State.*.RebarDiagramId в rngNDMSectionProperties.
Private Function TestMaterialDiagramIdsMatchStateProperties(ByRef materialDiagrams As Variant) As Boolean
    Dim diagramIds As Object
    Set diagramIds = CreateObject("Scripting.Dictionary")
    diagramIds.CompareMode = vbTextCompare

    Dim pointKeys As Object
    Set pointKeys = CreateObject("Scripting.Dictionary")
    pointKeys.CompareMode = vbTextCompare

    Dim diagramCol As Long
    Dim pointCol As Long
    diagramCol = TestResultHeaderColumn(materialDiagrams, "DiagramId")
    pointCol = TestResultHeaderColumn(materialDiagrams, "PointIndex")
    If diagramCol = 0 Or pointCol = 0 Then Exit Function

    Dim rowIndex As Long
    For rowIndex = 2 To UBound(materialDiagrams, 1)
        Dim diagramId As String
        diagramId = Trim$(CStr(materialDiagrams(rowIndex, diagramCol)))
        If Len(diagramId) = 0 Then Exit Function
        If Not diagramIds.Exists(diagramId) Then diagramIds.Add diagramId, True

        Dim pointKey As String
        pointKey = diagramId & "|" & CStr(materialDiagrams(rowIndex, pointCol))
        If pointKeys.Exists(pointKey) Then Exit Function
        pointKeys.Add pointKey, True
    Next rowIndex

    Dim usedIds As Object
    Set usedIds = CreateObject("Scripting.Dictionary")
    usedIds.CompareMode = vbTextCompare

    Dim props As Variant
    props = TestResultTable("rngNDMSectionProperties")
    Dim parameterCol As Long
    Dim valueCol As Long
    parameterCol = TestResultHeaderColumn(props, "Parameter")
    valueCol = TestResultHeaderColumn(props, "Value")
    If parameterCol = 0 Or valueCol = 0 Then Exit Function

    For rowIndex = 2 To UBound(props, 1)
        Dim parameterName As String
        parameterName = CStr(props(rowIndex, parameterCol))
        If Right$(parameterName, Len("ConcreteDiagramId")) = "ConcreteDiagramId" Or _
                Right$(parameterName, Len("RebarDiagramId")) = "RebarDiagramId" Then
            diagramId = Trim$(CStr(props(rowIndex, valueCol)))
            If Len(diagramId) = 0 Then Exit Function
            If Not diagramIds.Exists(diagramId) Then Exit Function
            If Not usedIds.Exists(diagramId) Then usedIds.Add diagramId, True
        End If
    Next rowIndex

    Dim keyVariant As Variant
    For Each keyVariant In diagramIds.Keys
        If Not usedIds.Exists(CStr(keyVariant)) Then Exit Function
    Next keyVariant

    TestMaterialDiagramIdsMatchStateProperties = True
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сверяет фактические validation-списки расчетов, геометрии, plot и export,
' динамические ссылки и default labels. Наличие списка не доказывает поведение каждой опции.
Private Sub TestCapacitySearchMethodValidation(ByRef stats As TUiTestStats)
    TestAssertTrue stats, "ui.validation.CapacitySolutionStrategy", _
        TestSystemSettingValidationHasOptions("Capacity.SolutionStrategy", Array("Auto", "UltimateStrain", "LoadMultiplier"))
    TestAssertTrue stats, "ui.validation.nonCriticalMessages", _
        TestSystemSettingValidationHasOptions("General.NonCriticalMessagesEnabled", Array("Yes", "No"))
    TestAssertTrue stats, "ui.validation.capacitySearchMethod", _
        TestSystemSettingValidationHasOptions("Capacity.SearchMethod", Array("Bisection", "Brent", "Secant"))
    TestAssertTrue stats, "ui.validation.crackCoverDistanceMode", _
        TestSystemSettingValidationHasOptions("SLS.Crack.CoverDistanceMode", Array("NearestContour", "GlobalExtreme"))
    TestAssertTrue stats, "ui.validation.universalLoadPath", TestUniversalLoadPathValidationHasOptions()
    TestAssertTrue stats, "ui.validation.autocadLabelMode", _
        TestSystemSettingValidationHasOptions("AutoCAD.Export.LabelMode", Array("ValuesOnly", "NamesAndValues"))
    TestAssertTrue stats, "ui.validation.autocadNeutralLine", _
        TestSystemSettingValidationHasOptions("AutoCAD.Export.NeutralLineEnabled", Array("Yes", "No"))
    TestAssertTrue stats, "ui.validation.autocadPrincipalAxes", _
        TestSystemSettingValidationHasOptions("AutoCAD.Export.PrincipalAxesMode", Array("Transformed", "Concrete", "None"))
    TestAssertTrue stats, "ui.validation.autocadLoadPoint", _
        TestSystemSettingValidationHasOptions("AutoCAD.Export.LoadPointEnabled", Array("Yes", "No"))
    TestAssertTrue stats, "ui.validation.autocadContour", _
        TestSystemSettingValidationHasOptions("AutoCAD.Export.ContourEnabled", Array("Yes", "No"))
    TestAssertTrue stats, "ui.validation.autocadCombination", TestAutoCADCombinationValidationIsDynamic()
    TestAssertTrue stats, "ui.validation.plotLoadCase", TestPlotLoadCaseValidationIsDynamic()
    TestAssertTrue stats, "ui.validation.loadProfileId", TestLoadProfileValidationUsesDisplayNames()
    TestAssertTrue stats, "ui.validation.loadProfileId.noFormula", TestLoadProfileFirstValueHasNoFormula()
    TestAssertTrue stats, "ui.validation.plotLabels", _
        TestSystemSettingValidationHasOptions("Plot.ResultLabelsEnabled", Array("Yes", "No"))
    TestAssertTrue stats, "ui.validation.plotPrincipalAxes", _
        TestSystemSettingValidationHasOptions("Plot.PrincipalAxesMode", Array("Transformed", "Concrete", "None"))
    TestAssertTrue stats, "ui.validation.plotAxisLabels", _
        TestSystemSettingValidationHasOptions("Plot.AxisLabelsEnabled", Array("Yes", "No"))
    TestAssertClose stats, "ui.validation.plotAxisLabelsFontSize.default", Val(CStr(TestGetSystemSetting("Plot.AxisLabelsFontSize"))), 7.5, 0.000000001
    TestAssertTrue stats, "ui.validation.plotContour", _
        TestSystemSettingValidationHasOptions("Plot.ContourEnabled", Array("Yes", "No"))
    TestAssertTrue stats, "ui.validation.plotRebarAnnotationEnabled", _
        TestPlotAnnotationValidationHasOptions("Enabled", 2, Array("Yes", "No"))
    TestAssertTrue stats, "ui.validation.plotDimensionEnabled", _
        TestPlotAnnotationValidationHasOptions("Enabled", 3, Array("Yes", "No"))
    TestAssertTrue stats, "ui.validation.plotRebarPlacement", _
        TestPlotAnnotationValidationHasOptions("Placement", 2, Array("Outside", "Inside"))
    TestAssertTrue stats, "ui.validation.plotDimensionPlacement", _
        TestPlotAnnotationValidationHasOptions("Placement", 3, Array("Outside", "Inside"))
    TestAssertTrue stats, "ui.validation.plotRebarTextUnits", _
        TestPlotAnnotationValidationHasOptions("TextUnits", 2, Array("mm", "pt"))
    TestAssertTrue stats, "ui.validation.plotDimensionTextUnits", _
        TestPlotAnnotationValidationHasOptions("TextUnits", 3, Array("mm", "pt"))
    TestAssertTrue stats, "ui.validation.plotTextHeightUnitDynamic", _
        TestPlotAnnotationTextUnitCellIsDynamic("TextHeight")
    TestAssertTrue stats, "ui.validation.plotTextGapUnitDynamic", _
        TestPlotAnnotationTextUnitCellIsDynamic("TextGap")
    TestAssertClose stats, "ui.validation.plotTextHeight.defaultRebar", TestPlotAnnotationSettingValue("TextHeight", 2), 13#, 0.000000001
    TestAssertClose stats, "ui.validation.plotTextHeight.defaultDimension", TestPlotAnnotationSettingValue("TextHeight", 3), 13#, 0.000000001
    TestAssertClose stats, "ui.validation.plotTextGap.defaultRebar", TestPlotAnnotationSettingValue("TextGap", 2), 9#, 0.000000001
    TestAssertClose stats, "ui.validation.plotTextGap.defaultDimension", TestPlotAnnotationSettingValue("TextGap", 3), 9#, 0.000000001
    TestAssertTrue stats, "ui.validation.plotRebarLineEnabled", _
        TestPlotAnnotationValidationHasOptions("LineEnabled", 2, Array("Yes", "No"))
    TestAssertTrue stats, "ui.validation.plotDimensionArrowType", _
        TestPlotAnnotationValidationHasOptions("ArrowType", 3, Array("Triangle", "Stealth", "Diamond", "Oval", "Open"))
    TestAssertTrue stats, "ui.validation.plotDimensionArrowSize", _
        TestPlotAnnotationValidationHasOptions("ArrowSize", 3, Array("Small", "Medium", "Wide"))
    TestAssertTrue stats, "ui.validation.geometrySource", _
        TestSystemSettingValidationHasOptions("Geometry.Source", Array("Generated", "AutoCAD"))
    TestAssertTrue stats, "ui.validation.units.forceInput", _
        TestAnySettingValidationHasOptions("Force", Array("N", "kN", "tf"))
    TestAssertTrue stats, "ui.validation.units.momentOutput", _
        TestAnySettingValidationHasOptions("Moment", Array("N*mm", "kN*m", "tf*m"))
    TestAssertTrue stats, "ui.validation.sign.n", _
        TestAnySettingValidationHasOptions("+N", Array("Tension", "Compression"))
    TestSetSystemSetting "Geometry.Type", "Circle"
    TestAssertTrue stats, "ui.validation.circleLoc2row", _
        TestAnySettingValidationHasOptions("Rebar.Loc2row", Array("Stacked", "SideBySide"))
    TestAssertTrue stats, "ui.validation.circleLoc3row", _
        TestAnySettingValidationHasOptions("Rebar.Loc3row", Array("Stacked", "SideBySide"))
    TestSetSystemSetting "Geometry.Type", "Rectangles"
    TestAssertTrue stats, "ui.validation.RectanglesSectionType", _
        TestRectanglesParameterValidationHasOptions("Rectangles.SectionType", Array("TwoRect", "OneRect"))
    TestAssertTrue stats, "ui.validation.rectanglesLoc2row", _
        TestRectanglesExtraValidationHasOptions("H1 - левая", "положение", 1, Array("Stacked", "SideBySide"))
    TestAssertTrue stats, "ui.validation.rectanglesBind2row", _
        TestRectanglesExtraValidationHasOptions("H1 - левая", "привязка", 1, Array("EachBar", "EverySecondBar"))
    Dim reader As CSystemSettingsReader
    Set reader = New CSystemSettingsReader
    TestSetSystemSetting "Geometry.Type", "Circle"
    reader.LoadFromWorkbook ThisWorkbook
    TestAssertTrue stats, "ui.circle.key.rebarDiameter2.recipeOwned", Not reader.HasKey("Rebar.Diameter2")
    TestAssertTrue stats, "ui.circle.key.rebarDiameter3.recipeOwned", Not reader.HasKey("Rebar.Diameter3")
    TestAssertTrue stats, "ui.circle.key.rebarLoc2row.recipeOwned", Not reader.HasKey("Rebar.Loc2row")
    TestAssertTrue stats, "ui.circle.key.rebarLoc3row.recipeOwned", Not reader.HasKey("Rebar.Loc3row")
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Подготавливает круговой fixture и проверяет реальную validation выбора LC для AutoCAD.
' Список должен содержать Worst/LC1 и ссылаться на формулы, а не быть статической строкой.
Private Function TestAutoCADCombinationValidationIsDynamic() As Boolean
    On Error GoTo Failed
    TestPrepareCircleInput

    Dim settings As Object
    Set settings = ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange
    Dim rowIndex As Long
    For rowIndex = 2 To settings.Rows.Count
        If StrComp(CStr(settings.Cells.Item(rowIndex, 1).Value2), "AutoCAD.Export.CombinationID", vbTextCompare) = 0 Then
            Dim formulaText As String
            formulaText = CStr(settings.Cells.Item(rowIndex, 2).Validation.Formula1)
            If Left$(formulaText, 1) <> "=" Then Exit Function

            Dim listRange As Object
            Set listRange = ThisWorkbook.Worksheets.Item("Config").Range(Mid$(formulaText, 2))
            listRange.Calculate ' Данные сочетаний изменены тестом; обновляем формулы списка и при ручном пересчёте.
            TestAutoCADCombinationValidationIsDynamic = (CStr(listRange.Cells.Item(1, 1).Value2) = "Worst" And _
                CStr(listRange.Cells.Item(2, 1).Value2) = "LC1" And listRange.Cells.Item(2, 1).HasFormula)
            Exit Function
        End If
    Next rowIndex
Failed:
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Применяет ту же проверку динамического списка к селектору LC Excel-схемы.
Private Function TestPlotLoadCaseValidationIsDynamic() As Boolean
    TestPlotLoadCaseValidationIsDynamic = TestLoadCaseValidationIsDynamicForSetting("Plot.LoadCase")
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' После подготовки fixture читает validation указанного Config-поля и зависимый диапазон.
' Ошибка/отсутствие диапазона означает непрохождение теста, а не создание новой validation.
Private Function TestLoadCaseValidationIsDynamicForSetting(ByVal settingKey As String) As Boolean
    On Error GoTo Failed
    TestPrepareCircleInput

    Dim settings As Object
    Set settings = ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange
    Dim rowIndex As Long
    For rowIndex = 2 To settings.Rows.Count
        If StrComp(CStr(settings.Cells.Item(rowIndex, 1).Value2), settingKey, vbTextCompare) = 0 Then
            Dim formulaText As String
            formulaText = CStr(settings.Cells.Item(rowIndex, 2).Validation.Formula1)
            If Left$(formulaText, 1) <> "=" Then Exit Function

            Dim listRange As Object
            Set listRange = ThisWorkbook.Worksheets.Item("Config").Range(Mid$(formulaText, 2))
            listRange.Calculate ' Данные сочетаний изменены тестом; обновляем формулы списка и при ручном пересчёте.
            TestLoadCaseValidationIsDynamicForSetting = (CStr(listRange.Cells.Item(1, 1).Value2) = "Worst" And _
                CStr(listRange.Cells.Item(2, 1).Value2) = "LC1" And listRange.Cells.Item(2, 1).HasFormula)
            Exit Function
        End If
    Next rowIndex
Failed:
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет пару LC/Parameter в сохраненной таблице свойств Results по именам колонок.
' Расчет не запускается; отсутствующий снимок или колонка дает False.
Private Function TestResultsPropertyExists(ByVal loadCase As String, ByVal parameter As String) As Boolean
    On Error GoTo Failed
    Dim data As Variant
    data = TestResultTable("rngNDMSectionProperties")
    Dim colLC As Long: colLC = TestResultHeaderColumn(data, "LoadCase")
    Dim colParam As Long: colParam = TestResultHeaderColumn(data, "Parameter")
    Dim rowIndex As Long
    For rowIndex = 2 To UBound(data, 1)
        If StrComp(CStr(data(rowIndex, colLC)), loadCase, vbTextCompare) = 0 And _
                StrComp(CStr(data(rowIndex, colParam)), parameter, vbTextCompare) = 0 Then
            TestResultsPropertyExists = True
            Exit Function
        End If
    Next rowIndex
Failed:
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает значение свойства из расчетного snapshot Results.
' Используется в UI-тестах, где важно проверить не только видимую Summary,
' но и машинные признаки выбранного LC, например ExtensionUsed.
Private Function TestResultsPropertyValue(ByVal loadCase As String, ByVal parameter As String) As String
    On Error GoTo Failed
    Dim data As Variant
    data = TestResultTable("rngNDMSectionProperties")
    Dim colLC As Long: colLC = TestResultHeaderColumn(data, "LoadCase")
    Dim colParam As Long: colParam = TestResultHeaderColumn(data, "Parameter")
    Dim colValue As Long: colValue = TestResultHeaderColumn(data, "Value")
    Dim rowIndex As Long
    For rowIndex = 2 To UBound(data, 1)
        If StrComp(CStr(data(rowIndex, colLC)), loadCase, vbTextCompare) = 0 And _
                StrComp(CStr(data(rowIndex, colParam)), parameter, vbTextCompare) = 0 Then
            TestResultsPropertyValue = CStr(data(rowIndex, colValue))
            Exit Function
        End If
    Next rowIndex
Failed:
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает длину из rngNDMSectionProperties во внутренних мм.
' Results хранит числовые значения в пользовательских OUTPUT-единицах последнего
' расчета, а Excel-схема строится во внутренних координатах модели.
Private Function TestResultsLengthPropertyMm(ByVal loadCase As String, ByVal parameter As String) As Double
    TestResultsLengthPropertyMm = CDbl(TestResultsPropertyValue(loadCase, parameter)) * TestResultsOutputLengthFactorToMm()
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Читает сохраненную в Results единицу вывода длины и возвращает множитель к мм.
' Тест намеренно опирается на snapshot Results, а не на текущий Config: пользователь
' может поменять настройки после расчета, но уже нарисованная схема должна
' соответствовать именно сохраненному расчетному снимку.
Private Function TestResultsOutputLengthFactorToMm() As Double
    Select Case LCase$(Trim$(TestResultsPropertyValue("ALL", "Output.LengthUnit")))
        Case "mm"
            TestResultsOutputLengthFactorToMm = 1#
        Case "cm"
            TestResultsOutputLengthFactorToMm = 10#
        Case "m"
            TestResultsOutputLengthFactorToMm = 1000#
        Case Else
            TestResultsOutputLengthFactorToMm = 1#
    End Select
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает числовое поле элемента из rngNDMSectionGeometry.
' Так UI-тест проверяет сохранность импортированных координат между
' geometry-preview и расчетным snapshot после очистки Results.
Private Function TestGeometryResultValue(ByRef data As Variant, ByVal elementID As String, _
        ByVal headerText As String) As Double
    Dim colID As Long
    Dim colValue As Long
    colID = TestResultHeaderColumn(data, "ElementID")
    colValue = TestResultHeaderColumn(data, headerText)
    If colValue = 0 Then colValue = TestResultHeaderColumnByBaseName(data, headerText)
    If colID = 0 Or colValue = 0 Then Err.Raise vbObjectError + 4850, "GeometryResultValue", _
        "В rngNDMSectionGeometry не найдены нужные заголовки."

    Dim rowIndex As Long
    For rowIndex = 2 To UBound(data, 1)
        If StrComp(CStr(data(rowIndex, colID)), elementID, vbTextCompare) = 0 Then
            TestGeometryResultValue = CDbl(data(rowIndex, colValue))
            Exit Function
        End If
    Next rowIndex

    Err.Raise vbObjectError + 4851, "GeometryResultValue", _
        "В rngNDMSectionGeometry не найден элемент " & elementID & "."
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Ищет колонку по имени до запятой в заголовке вида "X, mm".
Private Function TestResultHeaderColumnByBaseName(ByRef data As Variant, ByVal headerText As String) As Long
    Dim colIndex As Long
    For colIndex = 1 To UBound(data, 2)
        Dim actualHeader As String
        actualHeader = CStr(data(1, colIndex))
        If InStr(1, actualHeader, ",", vbTextCompare) > 0 Then _
            actualHeader = Trim$(Left$(actualHeader, InStr(1, actualHeader, ",", vbTextCompare) - 1))
        If StrComp(actualHeader, headerText, vbTextCompare) = 0 Then
            TestResultHeaderColumnByBaseName = colIndex
            Exit Function
        End If
    Next colIndex
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет, что справочные характеристики приведенного сечения в Results
' считаются через обычный модульный коэффициент Es/Eb. Это важно для
' устойчивости и ручной проверки геометрии: вид диаграммы TwoLine/ThreeLine
' не должен менять Ared.
Private Sub TestAssertTransformedAreaUsesElasticModuli(ByRef stats As TUiTestStats)
    Dim concreteArea As Double
    Dim transformedArea As Double
    Dim rebarArea As Double
    Dim expectedArea As Double
    Dim r1Plus As Double
    Dim r1Minus As Double
    Dim r2Plus As Double
    Dim r2Minus As Double
    Dim xPlus As Double
    Dim xMinus As Double
    Dim yPlus As Double
    Dim yMinus As Double

    concreteArea = CDbl(TestResultsPropertyValue("ALL", "Concrete.Area"))
    transformedArea = CDbl(TestResultsPropertyValue("ALL", "Transformed.Area"))
    rebarArea = 8# * GEOM_PI * 20# * 20# / 4#
    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook
    Dim units As CUnitSystem
    Set units = New CUnitSystem
    units.LoadFromSettings settings
    expectedArea = concreteArea + units.InternalAreaToOutput((200000# / 32500# - 1#) * rebarArea)

    TestAssertClose stats, "ui.results.properties.transformedArea.moduli", _
        transformedArea, expectedArea, 0.001
    r1Plus = CDbl(TestResultsPropertyValue("ALL", "Transformed.CoreDistance1Plus"))
    r1Minus = CDbl(TestResultsPropertyValue("ALL", "Transformed.CoreDistance1Minus"))
    r2Plus = CDbl(TestResultsPropertyValue("ALL", "Transformed.CoreDistance2Plus"))
    r2Minus = CDbl(TestResultsPropertyValue("ALL", "Transformed.CoreDistance2Minus"))
    TestAssertTrue stats, "ui.results.properties.transformedCoreDistance.positive", _
        r1Plus > 0# And r1Minus > 0# And r2Plus > 0# And r2Minus > 0#
    TestAssertClose stats, "ui.results.properties.transformedCoreDistance.axis1Sym", _
        r1Plus, r1Minus, 0.001
    TestAssertClose stats, "ui.results.properties.transformedCoreDistance.axis2Sym", _
        r2Plus, r2Minus, 0.001
    xPlus = CDbl(TestResultsPropertyValue("ALL", "Transformed.CoreDistanceXPlus"))
    xMinus = CDbl(TestResultsPropertyValue("ALL", "Transformed.CoreDistanceXMinus"))
    yPlus = CDbl(TestResultsPropertyValue("ALL", "Transformed.CoreDistanceYPlus"))
    yMinus = CDbl(TestResultsPropertyValue("ALL", "Transformed.CoreDistanceYMinus"))
    TestAssertTrue stats, "ui.results.properties.transformedCoreDistanceXY.positive", _
        xPlus > 0# And xMinus > 0# And yPlus > 0# And yMinus > 0#
    TestAssertTrue stats, "ui.results.properties.concreteCoreDistanceXY.exists", _
        Len(TestResultsPropertyValue("ALL", "Concrete.CoreDistanceXPlus")) > 0 And _
        Len(TestResultsPropertyValue("ALL", "Concrete.CoreDistanceYMinus")) > 0
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет наличие штатного ChartObject схемы на листе Расчет без его создания.
Private Function TestPlotChartExists() As Boolean
    On Error GoTo Failed
    Dim chartObject As Object
    Set chartObject = ThisWorkbook.Worksheets.Item("Расчет").ChartObjects("chtNDMSectionPlot")
    TestPlotChartExists = Not chartObject Is Nothing
Failed:
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет видимый Shape-заголовок существующей схемы.
Private Function TestPlotVisibleTitleContains(ByVal expectedText As String) As Boolean
    On Error GoTo Failed
    Dim chartObject As Object
    Set chartObject = ThisWorkbook.Worksheets.Item("Расчет").ChartObjects("chtNDMSectionPlot")

    Dim shapeIndex As Long
    For shapeIndex = 1 To chartObject.Chart.Shapes.Count
        If StrComp(chartObject.Chart.Shapes.Item(shapeIndex).Name, "NDMPlot_Title", vbTextCompare) = 0 Then
            TestPlotVisibleTitleContains = _
                (InStr(1, chartObject.Chart.Shapes.Item(shapeIndex).TextFrame.Characters().Text, expectedText, vbTextCompare) > 0)
            Exit Function
        End If
    Next shapeIndex
Failed:
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Считает подходящие фигуры внутри chart; при отсутствии схемы возвращает ноль.
Private Function TestCountPlotShapes(ByVal nameFragment As String) As Long
    On Error GoTo Failed
    Dim chartObject As Object
    Set chartObject = ThisWorkbook.Worksheets.Item("Расчет").ChartObjects("chtNDMSectionPlot")

    Dim shapeIndex As Long
    For shapeIndex = 1 To chartObject.Chart.Shapes.Count
        If InStr(1, chartObject.Chart.Shapes.Item(shapeIndex).Name, nameFragment, vbTextCompare) > 0 Then
            TestCountPlotShapes = TestCountPlotShapes + 1
        End If
    Next shapeIndex
Failed:
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Ищет тестируемую фигуру сначала в chart, затем на листе; отсутствие возвращает Nothing.
Private Function TestFirstGeneratedPlotShape(ByVal nameFragment As String) As Object
    On Error GoTo Failed
    Dim chartObject As Object
    Set chartObject = ThisWorkbook.Worksheets.Item("Расчет").ChartObjects("chtNDMSectionPlot")

    Dim shapeIndex As Long
    For shapeIndex = 1 To chartObject.Chart.Shapes.Count
        If InStr(1, chartObject.Chart.Shapes.Item(shapeIndex).Name, nameFragment, vbTextCompare) > 0 Then
            Set TestFirstGeneratedPlotShape = chartObject.Chart.Shapes.Item(shapeIndex)
            Exit Function
        End If
    Next shapeIndex

    Dim sheet As Object
    Set sheet = ThisWorkbook.Worksheets.Item("Расчет")
    For shapeIndex = 1 To sheet.Shapes.Count
        If InStr(1, sheet.Shapes.Item(shapeIndex).Name, nameFragment, vbTextCompare) > 0 Then
            Set TestFirstGeneratedPlotShape = sheet.Shapes.Item(shapeIndex)
            Exit Function
        End If
    Next shapeIndex
Failed:
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Центры фигур в единицах Excel points; для chart абсолютное смещение добавляется вызывающим кодом.
Private Function TestShapeCenterX(ByVal shapeObject As Object) As Double
    TestShapeCenterX = CDbl(shapeObject.Left) + CDbl(shapeObject.Width) / 2#
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает вертикальную середину фактической фигуры Excel для сравнения положения подписи со схемой.
Private Function TestShapeCenterY(ByVal shapeObject As Object) As Double
    TestShapeCenterY = CDbl(shapeObject.Top) + CDbl(shapeObject.Height) / 2#
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Учитывает оба контейнера generated-фигур: chart и лист.
Private Function TestCountGeneratedPlotShapes(ByVal nameFragment As String) As Long
    TestCountGeneratedPlotShapes = TestCountPlotShapes(nameFragment) + TestCountWorksheetPlotShapes(nameFragment)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Собирает средний абсолютный центр фигур двух контейнеров с учетом смещения chart.
' Вызывающий тест передает нулевые накопители; False означает отсутствие подходящих фигур.
Private Function TestGeneratedPlotShapeCenterAverage(ByVal nameFragment As String, ByRef centerX As Double, ByRef centerY As Double) As Boolean
    On Error GoTo Failed
    Dim chartObject As Object
    Set chartObject = ThisWorkbook.Worksheets.Item("Расчет").ChartObjects("chtNDMSectionPlot")

    Dim shapeIndex As Long
    Dim count As Long
    For shapeIndex = 1 To chartObject.Chart.Shapes.Count
        If InStr(1, chartObject.Chart.Shapes.Item(shapeIndex).Name, nameFragment, vbTextCompare) > 0 Then
            centerX = centerX + CDbl(chartObject.Left) + TestShapeCenterX(chartObject.Chart.Shapes.Item(shapeIndex))
            centerY = centerY + CDbl(chartObject.Top) + TestShapeCenterY(chartObject.Chart.Shapes.Item(shapeIndex))
            count = count + 1
        End If
    Next shapeIndex

    Dim sheet As Object
    Set sheet = ThisWorkbook.Worksheets.Item("Расчет")
    For shapeIndex = 1 To sheet.Shapes.Count
        If InStr(1, sheet.Shapes.Item(shapeIndex).Name, nameFragment, vbTextCompare) > 0 Then
            centerX = centerX + TestShapeCenterX(sheet.Shapes.Item(shapeIndex))
            centerY = centerY + TestShapeCenterY(sheet.Shapes.Item(shapeIndex))
            count = count + 1
        End If
    Next shapeIndex

    If count <= 0 Then Exit Function
    centerX = centerX / count
    centerY = centerY / count
    TestGeneratedPlotShapeCenterAverage = True
Failed:
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет границы каждой подходящей chart-фигуры относительно устойчивой рамки.
' Для PASS нужна хотя бы одна фигура; допуск задан тестом в Excel points.
Private Function TestPlotShapesInsideStableFrame(ByVal nameFragment As String, ByVal tolerance As Double) As Boolean
    On Error GoTo Failed
    Dim chartObject As Object
    Set chartObject = ThisWorkbook.Worksheets.Item("Расчет").ChartObjects("chtNDMSectionPlot")

    Dim leftBound As Double
    Dim topBound As Double
    Dim rightBound As Double
    Dim bottomBound As Double
    leftBound = TEST_PLOT_FRAME_LEFT - tolerance
    topBound = TEST_PLOT_FRAME_TOP - tolerance
    rightBound = TEST_PLOT_FRAME_LEFT + CDbl(chartObject.Width) - TEST_PLOT_FRAME_WIDTH_MARGIN + tolerance
    bottomBound = TEST_PLOT_FRAME_TOP + CDbl(chartObject.Height) - TEST_PLOT_FRAME_HEIGHT_MARGIN + tolerance

    Dim shapeIndex As Long
    Dim matched As Boolean
    For shapeIndex = 1 To chartObject.Chart.Shapes.Count
        If InStr(1, chartObject.Chart.Shapes.Item(shapeIndex).Name, nameFragment, vbTextCompare) > 0 Then
            matched = True
            If CDbl(chartObject.Chart.Shapes.Item(shapeIndex).Left) < leftBound Then Exit Function
            If CDbl(chartObject.Chart.Shapes.Item(shapeIndex).Top) < topBound Then Exit Function
            If CDbl(chartObject.Chart.Shapes.Item(shapeIndex).Left) + _
                    CDbl(chartObject.Chart.Shapes.Item(shapeIndex).Width) > rightBound Then Exit Function
            If CDbl(chartObject.Chart.Shapes.Item(shapeIndex).Top) + _
                    CDbl(chartObject.Chart.Shapes.Item(shapeIndex).Height) > bottomBound Then Exit Function
        End If
    Next shapeIndex
    TestPlotShapesInsideStableFrame = matched
Failed:
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет попадание абсолютной точки листа в границы хотя бы одной chart-фигуры.
' Смещение ChartObject учитывается отдельно от локальных координат фигуры.
Private Function TestPlotShapeBoundsContainAbsolutePoint(ByVal nameFragment As String, _
        ByVal pointX As Double, ByVal pointY As Double, ByVal tolerance As Double) As Boolean
    On Error GoTo Failed
    Dim chartObject As Object
    Set chartObject = ThisWorkbook.Worksheets.Item("Расчет").ChartObjects("chtNDMSectionPlot")

    Dim shapeIndex As Long
    Dim leftValue As Double
    Dim topValue As Double
    Dim rightValue As Double
    Dim bottomValue As Double
    For shapeIndex = 1 To chartObject.Chart.Shapes.Count
        If InStr(1, chartObject.Chart.Shapes.Item(shapeIndex).Name, nameFragment, vbTextCompare) > 0 Then
            leftValue = CDbl(chartObject.Left) + CDbl(chartObject.Chart.Shapes.Item(shapeIndex).Left) - tolerance
            topValue = CDbl(chartObject.Top) + CDbl(chartObject.Chart.Shapes.Item(shapeIndex).Top) - tolerance
            rightValue = CDbl(chartObject.Left) + CDbl(chartObject.Chart.Shapes.Item(shapeIndex).Left) + _
                CDbl(chartObject.Chart.Shapes.Item(shapeIndex).Width) + tolerance
            bottomValue = CDbl(chartObject.Top) + CDbl(chartObject.Chart.Shapes.Item(shapeIndex).Top) + _
                CDbl(chartObject.Chart.Shapes.Item(shapeIndex).Height) + tolerance
            If pointX >= leftValue And pointX <= rightValue And _
                    pointY >= topValue And pointY <= bottomValue Then
                TestPlotShapeBoundsContainAbsolutePoint = True
                Exit Function
            End If
        End If
    Next shapeIndex
Failed:
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' считает только Shapes схемы с префиксом NDMPlot_ на листе Расчет.
' Объекты Chart и пользовательские фигуры не включаются в проверку очистки.
Private Function TestCountWorksheetPlotShapes(ByVal nameFragment As String) As Long
    On Error GoTo Failed
    Dim sheet As Object
    Set sheet = ThisWorkbook.Worksheets.Item("Расчет")

    Dim shapeIndex As Long
    For shapeIndex = 1 To sheet.Shapes.Count
        If Left$(sheet.Shapes.Item(shapeIndex).Name, Len("NDMPlot_")) = "NDMPlot_" And _
                InStr(1, sheet.Shapes.Item(shapeIndex).Name, nameFragment, vbTextCompare) > 0 Then
            TestCountWorksheetPlotShapes = TestCountWorksheetPlotShapes + 1
        End If
    Next shapeIndex
Failed:
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет, что среди Shapes схемы есть объект с нужным углом поворота.
' Excel может хранить один и тот же угол как -30 или 330 градусов, поэтому
' сравниваем минимальную круговую разницу.
Private Function TestPlotShapeRotationExists(ByVal nameFragment As String, _
        ByVal expectedDegrees As Double, ByVal toleranceDegrees As Double) As Boolean
    On Error GoTo Failed
    Dim chartObject As Object
    Set chartObject = ThisWorkbook.Worksheets.Item("Расчет").ChartObjects("chtNDMSectionPlot")

    Dim shapeIndex As Long
    For shapeIndex = 1 To chartObject.Chart.Shapes.Count
        If InStr(1, chartObject.Chart.Shapes.Item(shapeIndex).Name, nameFragment, vbTextCompare) > 0 Then
            If TestAngleDistanceDegrees(CDbl(chartObject.Chart.Shapes.Item(shapeIndex).Rotation), expectedDegrees) <= toleranceDegrees Then
                TestPlotShapeRotationExists = True
                Exit Function
            End If
        End If
    Next shapeIndex
Failed:
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает минимальную разницу между углами в градусах с учетом периода 360.
Private Function TestAngleDistanceDegrees(ByVal actualDegrees As Double, ByVal expectedDegrees As Double) As Double
    Dim diff As Double
    diff = Abs(TestNormalizeDegrees(actualDegrees) - TestNormalizeDegrees(expectedDegrees))
    If diff > 180# Then diff = 360# - diff
    TestAngleDistanceDegrees = diff
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Нормализует угол к диапазону 0...360 для устойчивого сравнения Excel Shapes.
Private Function TestNormalizeDegrees(ByVal angleDegrees As Double) As Double
    Do While angleDegrees < 0#
        angleDegrees = angleDegrees + 360#
    Loop
    Do While angleDegrees >= 360#
        angleDegrees = angleDegrees - 360#
    Loop
    TestNormalizeDegrees = angleDegrees
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Ищет текст среди Shape-подписей текущей схемы.
' Тесты используют это для предупреждений, которые рисуются внутри ChartObject,
' а не выводятся отдельным окном Excel.
Private Function TestPlotShapeTextContains(ByVal expectedText As String) As Boolean
    On Error GoTo Failed
    Dim chartObject As Object
    Set chartObject = ThisWorkbook.Worksheets.Item("Расчет").ChartObjects("chtNDMSectionPlot")

    Dim shapeIndex As Long
    For shapeIndex = 1 To chartObject.Chart.Shapes.Count
        Dim textValue As String
        textValue = vbNullString
        Err.Clear
        On Error Resume Next
        textValue = chartObject.Chart.Shapes.Item(shapeIndex).TextFrame.Characters().Text
        On Error GoTo Failed
        If InStr(1, textValue, expectedText, vbTextCompare) > 0 Then
            TestPlotShapeTextContains = True
            Exit Function
        End If
    Next shapeIndex
Failed:
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет курсив у текстового Shape схемы по фрагменту имени.
Private Function TestPlotShapeTextItalic(ByVal nameFragment As String) As Boolean
    On Error GoTo Failed
    Dim chartObject As Object
    Set chartObject = ThisWorkbook.Worksheets.Item("Расчет").ChartObjects("chtNDMSectionPlot")

    Dim shapeIndex As Long
    For shapeIndex = 1 To chartObject.Chart.Shapes.Count
        If InStr(1, chartObject.Chart.Shapes.Item(shapeIndex).Name, nameFragment, vbTextCompare) > 0 Then
            TestPlotShapeTextItalic = (chartObject.Chart.Shapes.Item(shapeIndex).TextFrame.Characters().Font.Italic <> 0)
            Exit Function
        End If
    Next shapeIndex
Failed:
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает размер шрифта текстового Shape схемы по фрагменту имени.
Private Function TestPlotShapeTextFontSize(ByVal nameFragment As String) As Double
    On Error GoTo Failed
    Dim chartObject As Object
    Set chartObject = ThisWorkbook.Worksheets.Item("Расчет").ChartObjects("chtNDMSectionPlot")

    Dim shapeIndex As Long
    For shapeIndex = 1 To chartObject.Chart.Shapes.Count
        If InStr(1, chartObject.Chart.Shapes.Item(shapeIndex).Name, nameFragment, vbTextCompare) > 0 Then
            TestPlotShapeTextFontSize = CDbl(chartObject.Chart.Shapes.Item(shapeIndex).TextFrame.Characters().Font.Size)
            Exit Function
        End If
    Next shapeIndex
Failed:
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' ДЛЯ ТЕСТОВ
' Возвращает число рядов данных на текущей схеме Excel.
Private Function TestPlotSeriesCount() As Long
    On Error GoTo Failed
    Dim chartObject As Object
    Set chartObject = ThisWorkbook.Worksheets.Item("Расчет").ChartObjects("chtNDMSectionPlot")
    TestPlotSeriesCount = chartObject.Chart.SeriesCollection.Count
    Exit Function

Failed:
    TestPlotSeriesCount = 0
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' считает строки сохраненной semantic-таблицы нужного типа.
' Заголовок пропускается; тест проверяет данные Results, не текущий генератор.
Private Function TestCountAnnotationType(ByRef annotationData As Variant, ByVal annotationType As String) As Long
    On Error GoTo Failed
    Dim rowIndex As Long
    For rowIndex = 2 To UBound(annotationData, 1)
        If StrComp(CStr(annotationData(rowIndex, 2)), annotationType, vbTextCompare) = 0 Then
            TestCountAnnotationType = TestCountAnnotationType + 1
        End If
    Next rowIndex
Failed:
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' ДЛЯ ТЕСТОВ
' Возвращает видимую подпись semantic-аннотации из Results по ее ID.
Private Function TestAnnotationTextByID(ByRef annotationData As Variant, ByVal annotationID As String) As String
    On Error GoTo Failed
    Dim colID As Long
    Dim colText As Long
    colID = TestResultHeaderColumn(annotationData, "AnnotationID")
    colText = TestResultHeaderColumn(annotationData, "Text")

    Dim rowIndex As Long
    For rowIndex = 2 To UBound(annotationData, 1)
        If StrComp(CStr(annotationData(rowIndex, colID)), annotationID, vbTextCompare) = 0 Then
            TestAnnotationTextByID = CStr(annotationData(rowIndex, colText))
            Exit Function
        End If
    Next rowIndex

Failed:
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' ДЛЯ ТЕСТОВ
' Возвращает численное значение semantic-аннотации из Results по ее ID.
Private Function TestAnnotationValueByID(ByRef annotationData As Variant, ByVal annotationID As String) As Double
    On Error GoTo Failed
    Dim colID As Long
    Dim colValue As Long
    colID = TestResultHeaderColumn(annotationData, "AnnotationID")
    colValue = TestResultHeaderColumn(annotationData, "Value")

    Dim rowIndex As Long
    For rowIndex = 2 To UBound(annotationData, 1)
        If StrComp(CStr(annotationData(rowIndex, colID)), annotationID, vbTextCompare) = 0 Then
            TestAnnotationValueByID = CDbl(annotationData(rowIndex, colValue))
            Exit Function
        End If
    Next rowIndex

Failed:
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Очищает все расчетные snapshot-блоки пользовательским macro, проверяя
' пустые результаты и сохранение исходного Geometry.Type на Config.
Private Sub TestClearResultsKeepsInputs(ByRef stats As TUiTestStats)
    TestPrepareCircleInput
    ThisWorkbook.Names.Item("rngNDMElementResults").RefersToRange.Value2 = "RunID"
    ThisWorkbook.Names.Item("rngNDMSectionGeometry").RefersToRange.Value2 = "RunID"
    ThisWorkbook.Names.Item("rngNDMSectionProperties").RefersToRange.Value2 = "RunID"
    ThisWorkbook.Names.Item("rngNDMSectionAnnotations").RefersToRange.Value2 = "RunID"
    ThisWorkbook.Names.Item("rngNDMMaterialDiagrams").RefersToRange.Value2 = "RunID"
    ClearSectionResultsForWorkbook ThisWorkbook
    TestAssertTrue stats, "ui.clear.results.elements", Len(CStr(ThisWorkbook.Names.Item("rngNDMElementResults").RefersToRange.Value2)) = 0
    TestAssertTrue stats, "ui.clear.results.geometry", Len(CStr(ThisWorkbook.Names.Item("rngNDMSectionGeometry").RefersToRange.Value2)) = 0
    TestAssertTrue stats, "ui.clear.results.properties", Len(CStr(ThisWorkbook.Names.Item("rngNDMSectionProperties").RefersToRange.Value2)) = 0
    TestAssertTrue stats, "ui.clear.results.annotations", Len(CStr(ThisWorkbook.Names.Item("rngNDMSectionAnnotations").RefersToRange.Value2)) = 0
    TestAssertTrue stats, "ui.clear.results.materialDiagrams", Len(CStr(ThisWorkbook.Names.Item("rngNDMMaterialDiagrams").RefersToRange.Value2)) = 0
    TestAssertTrue stats, "ui.clear.input", CStr(TestGetSystemSetting("Geometry.Type")) = "Circle"
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Изменяет ячейку fixture по каноническому ключу с учетом специализированных таблиц.
' Неизвестный ключ вызывает ошибку теста, а не создает новую настройку.
Private Sub TestSetSystemSetting(ByVal key As String, ByVal value As String)
    Dim geometryCell As Object
    Set geometryCell = TestFixtureConfigFieldCell(key)
    If Not geometryCell Is Nothing Then
        geometryCell.Value2 = value: Exit Sub
    End If
    If TestTrySetUnitOrSignSetting(key, value) Then Exit Sub

    Dim settings As Object
    Dim ranges As Variant
    ranges = TestSettingsRangeSearchOrder()

    Dim rangeIndex As Long
    For rangeIndex = LBound(ranges) To UBound(ranges)
        Set settings = ThisWorkbook.Names.Item(CStr(ranges(rangeIndex))).RefersToRange

        Dim rowIndex As Long
        For rowIndex = 2 To settings.Rows.Count
            If StrComp(CStr(settings.Cells.Item(rowIndex, 1).Value2), key, vbTextCompare) = 0 Then
                settings.Cells.Item(rowIndex, 2).Value2 = value
                Exit Sub
            End If
        Next rowIndex
    Next rangeIndex


    Err.Raise vbObjectError + 4210, "modTestWorkbookInterface", "System setting not found: " & key
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Меняет одну ячейку расчетного профиля в тестовой книге.
' Диапазон профилей имеет две строки шапки, поэтому ProfileId ищется
' динамически, как это делает реальный reader.
Private Sub TestSetProfileSetting(ByVal profileId As String, ByVal key As String, ByVal value As String)
    Dim profiles As Object
    Set profiles = ThisWorkbook.Names.Item("rngCalculationProfiles").RefersToRange

    Dim profileColumn As Long
    Dim rowIndex As Long
    Dim colIndex As Long
    For rowIndex = 1 To profiles.Rows.Count
        For colIndex = 3 To profiles.Columns.Count
            If StrComp(Trim$(CStr(profiles.Cells.Item(rowIndex, colIndex).Value2)), profileId, vbTextCompare) = 0 Then
                profileColumn = colIndex
                Exit For
            End If
        Next colIndex
        If profileColumn > 0 Then Exit For
    Next rowIndex
    If profileColumn = 0 Then Err.Raise vbObjectError + 4212, "modTestWorkbookInterface", "ProfileId not found: " & profileId

    For rowIndex = 1 To profiles.Rows.Count
        If StrComp(Trim$(CStr(profiles.Cells.Item(rowIndex, 2).Value2)), key, vbTextCompare) = 0 Then
            profiles.Cells.Item(rowIndex, profileColumn).Value2 = value
            Exit Sub
        End If
    Next rowIndex

    Err.Raise vbObjectError + 4213, "modTestWorkbookInterface", "Profile setting not found: " & key
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Читает ячейку расчетного профиля по ProfileId и Key. Это нужно тестам,
' чтобы не зависеть от физического номера строки в rngCalculationProfiles.
Private Function TestProfileSettingValue(ByVal profileId As String, ByVal key As String) As String
    Dim profiles As Object
    Set profiles = ThisWorkbook.Names.Item("rngCalculationProfiles").RefersToRange

    Dim profileColumn As Long
    Dim rowIndex As Long
    Dim colIndex As Long
    For rowIndex = 1 To profiles.Rows.Count
        For colIndex = 3 To profiles.Columns.Count
            If StrComp(Trim$(CStr(profiles.Cells.Item(rowIndex, colIndex).Value2)), profileId, vbTextCompare) = 0 Then
                profileColumn = colIndex
                Exit For
            End If
        Next colIndex
        If profileColumn > 0 Then Exit For
    Next rowIndex
    If profileColumn = 0 Then Exit Function

    For rowIndex = 1 To profiles.Rows.Count
        If StrComp(Trim$(CStr(profiles.Cells.Item(rowIndex, 2).Value2)), key, vbTextCompare) = 0 Then
            TestProfileSettingValue = CStr(profiles.Cells.Item(rowIndex, profileColumn).Value2)
            Exit Function
        End If
    Next rowIndex
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Направляет тестовый ввод единиц и знаков в их отдельные именованные таблицы.
' Возвращает False для другого типа ключа, чтобы продолжить поиск в Config.
Private Function TestTrySetUnitOrSignSetting(ByVal key As String, ByVal value As String) As Boolean
    Dim target As Object
    Dim rowIndex As Long
    Dim columnIndex As Long

    If TestUnitSettingAddress(key, rowIndex, columnIndex) Then
        Set target = ThisWorkbook.Names.Item("rngUnitSettings").RefersToRange
        target.Cells.Item(rowIndex, columnIndex).Value2 = value
        TestTrySetUnitOrSignSetting = True
        Exit Function
    End If

    If TestSignSettingAddress(key, rowIndex, columnIndex) Then
        Set target = ThisWorkbook.Names.Item("rngSignConventionSettings").RefersToRange
        target.Cells.Item(rowIndex, columnIndex).Value2 = value
        TestTrySetUnitOrSignSetting = True
    End If
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Разбирает ключ Units.Quantity.Side в строку физической величины и колонку стороны.
' Неизвестная величина или сторона не трактуется как допустимая настройка.
Private Function TestUnitSettingAddress(ByVal key As String, ByRef rowIndex As Long, ByRef columnIndex As Long) As Boolean
    If InStr(1, key, "Units.", vbTextCompare) <> 1 Then Exit Function

    Dim quantity As String
    Dim sideName As String
    Dim tail As String
    tail = Mid$(key, 7)
    Dim dotPos As Long
    dotPos = InStr(1, tail, ".", vbTextCompare)
    If dotPos <= 0 Then Exit Function
    quantity = Left$(tail, dotPos - 1)
    sideName = Mid$(tail, dotPos + 1)

    Select Case LCase$(quantity)
        Case "length": rowIndex = 2
        Case "area": rowIndex = 3
        Case "force": rowIndex = 4
        Case "moment": rowIndex = 5
        Case "stress": rowIndex = 6
        Case "curvature": rowIndex = 7
        Case Else: Exit Function
    End Select

    Select Case LCase$(sideName)
        Case "input": columnIndex = 2
        Case "internal": columnIndex = 3
        Case "output": columnIndex = 4
        Case Else: Exit Function
    End Select
    TestUnitSettingAddress = True
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сопоставляет ключ Sign.N/Mx/My.User/Internal с отдельной таблицей знаков.
' Индексы возвращаются только для поддержанного формата ключа.
Private Function TestSignSettingAddress(ByVal key As String, ByRef rowIndex As Long, ByRef columnIndex As Long) As Boolean
    If InStr(1, key, "Sign.", vbTextCompare) <> 1 Then Exit Function

    Dim quantity As String
    Dim sideName As String
    Dim tail As String
    tail = Mid$(key, 6)
    Dim dotPos As Long
    dotPos = InStr(1, tail, ".", vbTextCompare)
    If dotPos <= 0 Then Exit Function
    quantity = Left$(tail, dotPos - 1)
    sideName = Mid$(tail, dotPos + 1)

    Select Case LCase$(quantity)
        Case "n": rowIndex = 2
        Case "mx": rowIndex = 3
        Case "my": rowIndex = 4
        Case Else: Exit Function
    End Select

    Select Case LCase$(sideName)
        Case "user": columnIndex = 2
        Case "internal": columnIndex = 3
        Case Else: Exit Function
    End Select
    TestSignSettingAddress = True
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Читает контрольный параметр из его действующей именованной таблицы.
Private Function TestGetSystemSetting(ByVal key As String) As String
    Dim geometryCell As Object
    Set geometryCell = TestFixtureConfigFieldCell(key)
    If Not geometryCell Is Nothing Then
        TestGetSystemSetting = CStr(geometryCell.Value2): Exit Function
    End If
    Dim directValue As String
    If TestTryGetUnitOrSignSetting(key, directValue) Then
        TestGetSystemSetting = directValue
        Exit Function
    End If

    Dim settings As Object
    Dim ranges As Variant
    ranges = TestSettingsRangeSearchOrder()

    Dim rangeIndex As Long
    For rangeIndex = LBound(ranges) To UBound(ranges)
        Set settings = ThisWorkbook.Names.Item(CStr(ranges(rangeIndex))).RefersToRange

        Dim rowIndex As Long
        For rowIndex = 2 To settings.Rows.Count
            If StrComp(CStr(settings.Cells.Item(rowIndex, 1).Value2), key, vbTextCompare) = 0 Then
                TestGetSystemSetting = CStr(settings.Cells.Item(rowIndex, 2).Value2)
                Exit Function
            End If
        Next rowIndex
    Next rangeIndex

    Err.Raise vbObjectError + 4211, "modTestWorkbookInterface", "System setting not found: " & key
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет наличие имени книги без падения теста на ошибке Excel.
' Используется для контроля, что старые входные диапазоны не вернулись после
' перестройки листа Config.
Private Function TestWorkbookNameExists(ByVal nameText As String) As Boolean
    On Error Resume Next
    Dim nm As Object
    Set nm = ThisWorkbook.Names.Item(nameText)
    TestWorkbookNameExists = (Err.Number = 0 And Not nm Is Nothing)
    Err.Clear
    On Error GoTo 0
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Читает настройки из специальных таблиц единиц и знаков. Эти диапазоны не
' имеют обычного формата Key/Value, поэтому используют ту же адресацию, что и
' SetSystemSetting при записи тестового значения.
Private Function TestTryGetUnitOrSignSetting(ByVal key As String, ByRef value As String) As Boolean
    Dim target As Object
    Dim rowIndex As Long
    Dim columnIndex As Long

    If TestUnitSettingAddress(key, rowIndex, columnIndex) Then
        Set target = ThisWorkbook.Names.Item("rngUnitSettings").RefersToRange
        value = CStr(target.Cells.Item(rowIndex, columnIndex).Value2)
        TestTryGetUnitOrSignSetting = True
        Exit Function
    End If

    If TestSignSettingAddress(key, rowIndex, columnIndex) Then
        Set target = ThisWorkbook.Names.Item("rngSignConventionSettings").RefersToRange
        value = CStr(target.Cells.Item(rowIndex, columnIndex).Value2)
        TestTryGetUnitOrSignSetting = True
    End If
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Ставит выбранную геометрию раньше неактивных таблиц с одинаковыми именами параметров.
' Остальные системные, материальные и презентационные блоки остаются доступны.
Private Function TestSettingsRangeSearchOrder() As Variant
    Dim geometryType As String
    geometryType = TestSystemGeometryType()
    If StrComp(geometryType, "Rectangles", vbTextCompare) = 0 Then
        TestSettingsRangeSearchOrder = Array("rngUnitSettings", "rngSignConventionSettings", "rngSystemSettings", "rngConcreteMaterialParameters", "rngSteelMaterialParameters", "rngCalculationProfiles", "rngPlotAnnotationSettings", "rngRectanglesGeometry", "rngCircleGeometry", "rngPointedSectionGeometry", "rngHollowRectangleGeometry")
    ElseIf StrComp(geometryType, "PointedSection", vbTextCompare) = 0 Then
        TestSettingsRangeSearchOrder = Array("rngUnitSettings", "rngSignConventionSettings", "rngSystemSettings", "rngConcreteMaterialParameters", "rngSteelMaterialParameters", "rngCalculationProfiles", "rngPlotAnnotationSettings", "rngPointedSectionGeometry", "rngCircleGeometry", "rngRectanglesGeometry", "rngHollowRectangleGeometry")
    ElseIf StrComp(geometryType, "HollowRectangle", vbTextCompare) = 0 Then
        TestSettingsRangeSearchOrder = Array("rngUnitSettings", "rngSignConventionSettings", "rngSystemSettings", "rngConcreteMaterialParameters", "rngSteelMaterialParameters", "rngCalculationProfiles", "rngPlotAnnotationSettings", "rngHollowRectangleGeometry", "rngPointedSectionGeometry", "rngCircleGeometry", "rngRectanglesGeometry")
    Else
        TestSettingsRangeSearchOrder = Array("rngUnitSettings", "rngSignConventionSettings", "rngSystemSettings", "rngConcreteMaterialParameters", "rngSteelMaterialParameters", "rngCalculationProfiles", "rngPlotAnnotationSettings", "rngCircleGeometry", "rngPointedSectionGeometry", "rngHollowRectangleGeometry", "rngRectanglesGeometry")
    End If
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Читает тип геометрии для поиска тестовых ячеек без построения модели.
' При отсутствии тестовых данных инфраструктура использует Circle.
Private Function TestSystemGeometryType() As String
    On Error GoTo Failed
    Dim settings As Object
    Set settings = ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange

    Dim rowIndex As Long
    For rowIndex = 2 To settings.Rows.Count
        If StrComp(CStr(settings.Cells.Item(rowIndex, 1).Value2), "Geometry.Type", vbTextCompare) = 0 Then
            TestSystemGeometryType = CStr(settings.Cells.Item(rowIndex, 2).Value2)
            Exit Function
        End If
    Next rowIndex
Failed:
    TestSystemGeometryType = "Circle"
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Группа wrappers проверяет validation в выбранном блоке, а не только текст ячейки.
' Реальные списки допустимых вариантов сверяются адресными helpers ниже.
Private Function TestSystemSettingValidationHasOptions(ByVal key As String, ByVal expectedOptions As Variant) As Boolean
    TestSystemSettingValidationHasOptions = TestSettingValidationHasOptionsInRange(ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange, key, expectedOptions)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' проверяет реальный общий dropdown сочетания, включая Auto.
' Пустой ввод разрешен; удаленная отдельная настройка Formation не используется.
Private Function TestUniversalLoadPathValidationHasOptions() As Boolean
    On Error GoTo Failed
    Dim loads As Object, cell As Object
    Set loads = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange
    If CStr(loads.Cells(1, 6).Value2) <> "LoadPath" Then Exit Function
    Set cell = loads.Cells(2, 6)
    If cell.Validation.Type <> xlValidateList Or Not cell.Validation.IgnoreBlank Then Exit Function
    Dim expected As Variant
    expected = Array("Auto", ChrW$(&H3BB) & "*Mx", ChrW$(&H3BB) & "*My", _
        ChrW$(&H3BB) & "*Mxy", ChrW$(&H3BB) & "*N", ChrW$(&H3BB) & "*NMxy")
    TestUniversalLoadPathValidationHasOptions = TestValidationCellHasOptions(cell, expected)
Failed:
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Ищет настройку в rngSystemSettings и проверяет, что ее колонка "Ед."
' является формулой к нужной строке rngUnitSettings. Видимое значение
' может быть разным после выбора пользователем единиц, поэтому тестируем
' источник данных формулы, а не конкретный текст вроде tf или kN*m.
Private Function TestSystemSettingUnitCellReferencesQuantity(ByVal key As String, ByVal quantity As String) As Boolean
    On Error GoTo Failed

    Dim settings As Object
    Set settings = ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange

    Dim rowIndex As Long
    For rowIndex = 2 To settings.Rows.Count
        If StrComp(CStr(settings.Cells.Item(rowIndex, 1).Value2), key, vbTextCompare) = 0 Then
            Dim formulaText As String
            formulaText = CStr(settings.Cells.Item(rowIndex, 3).Formula)
            If Left$(formulaText, 1) <> "=" Then Exit Function

            TestSystemSettingUnitCellReferencesQuantity = _
                InStr(1, formulaText, "rngUnitSettings", vbTextCompare) > 0 And _
                InStr(1, formulaText, """" & quantity & """", vbTextCompare) > 0
            Exit Function
        End If
    Next rowIndex

Failed:
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Читает сохранённую таблицу от её именованной начальной ячейки до конца данных.
' Новый расчёт и текущие исходные размеры сечения не используются.
Private Function TestResultTable(ByVal rangeName As String) As Variant
    If rangeName = "rngNDMSectionProperties" Or rangeName = "rngNDMSectionAnnotations" Then
        TestResultTable = ReadAnchoredResultTable(ThisWorkbook, rangeName)
        Exit Function
    End If
    Dim anchor As Object
    Set anchor = ThisWorkbook.Names.Item(rangeName).RefersToRange

    Dim columnCount As Long
    columnCount = TestResultTableColumnCount(anchor)
    If columnCount <= 0 Then Exit Function

    Dim rowCount As Long
    rowCount = TestResultTableRowCount(rangeName)
    If rowCount <= 0 Then rowCount = 1

    TestResultTable = anchor.Resize(rowCount, columnCount).Value2
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Определяет число непрерывно заполненных строк сохранённой таблицы
' по первому ключевому столбцу до первой пустой строки.
Private Function TestResultTableRowCount(ByVal rangeName As String) As Long
    Dim anchor As Object
    Set anchor = ThisWorkbook.Names.Item(rangeName).RefersToRange

    Dim rowOffset As Long
    For rowOffset = 0 To 1048575 - anchor.Row
        If Len(Trim$(CStr(anchor.Offset(rowOffset, 0).Value2))) = 0 Then Exit For
        TestResultTableRowCount = TestResultTableRowCount + 1
    Next rowOffset
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Определяет ширину сохранённой таблицы с учётом соседнего именованного начала
' справа. Чтение ограничено своим блоком Results.
Private Function TestResultTableColumnCount(ByVal anchor As Object) As Long
    Dim maxColumns As Long
    maxColumns = TestColumnsUntilNextResultAnchor(anchor)
    If maxColumns <= 0 Then maxColumns = 256

    Dim colOffset As Long
    For colOffset = 0 To maxColumns - 1
        If Len(Trim$(CStr(anchor.Offset(0, colOffset).Value2))) = 0 Then Exit For
        TestResultTableColumnCount = TestResultTableColumnCount + 1
    Next colOffset
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет, что единица настройки собрана формулой сразу из двух строк
' rngUnitSettings. Это нужно для величин вида момент/длина, где видимый текст
' меняется вместе с выбранной пользователем INPUT-системой единиц.
Private Function TestSystemSettingUnitCellReferencesQuantities(ByVal key As String, _
        ByVal quantity1 As String, ByVal quantity2 As String) As Boolean
    On Error GoTo Failed

    Dim settings As Object
    Set settings = ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange

    Dim rowIndex As Long
    For rowIndex = 2 To settings.Rows.Count
        If StrComp(CStr(settings.Cells.Item(rowIndex, 1).Value2), key, vbTextCompare) = 0 Then
            Dim formulaText As String
            formulaText = CStr(settings.Cells.Item(rowIndex, 3).Formula)
            If Left$(formulaText, 1) <> "=" Then Exit Function

            TestSystemSettingUnitCellReferencesQuantities = _
                InStr(1, formulaText, "rngUnitSettings", vbTextCompare) > 0 And _
                InStr(1, formulaText, """" & quantity1 & """", vbTextCompare) > 0 And _
                InStr(1, formulaText, """" & quantity2 & """", vbTextCompare) > 0
            Exit Function
        End If
    Next rowIndex

Failed:
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Ограничивает чтение таблицы Results ближайшим соседним именованным
' диапазоном справа. Блоки Results стоят на одной строке и не обязаны иметь
' пустой столбец между собой, поэтому простого CurrentRegion здесь мало.
Private Function TestColumnsUntilNextResultAnchor(ByVal anchor As Object) As Long
    Dim bestDelta As Long
    bestDelta = 0

    Dim nm As Object
    For Each nm In ThisWorkbook.Names
        Dim candidate As Object
        On Error Resume Next
        Set candidate = nm.RefersToRange
        If Err.Number <> 0 Then
            Err.Clear
            Set candidate = Nothing
        End If
        On Error GoTo 0

        If Not candidate Is Nothing Then
            If candidate.Worksheet.Name = anchor.Worksheet.Name And candidate.Row = anchor.Row Then
                Dim delta As Long
                delta = candidate.Column - anchor.Column
                If delta > 0 Then
                    If bestDelta = 0 Or delta < bestDelta Then bestDelta = delta
                End If
            End If
        End If
    Next nm

    TestColumnsUntilNextResultAnchor = bestDelta
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Ищет validation ключа в реально существующих таблицах выбранной конфигурации.
' Успех требует совпадения полного списка вариантов, а не только наличия dropdown.
Private Function TestAnySettingValidationHasOptions(ByVal key As String, ByVal expectedOptions As Variant) As Boolean
    Dim geometryCell As Object
    Set geometryCell = TestFixtureConfigFieldCell(key)
    If Not geometryCell Is Nothing Then
        TestAnySettingValidationHasOptions = TestValidationCellHasOptions(geometryCell, expectedOptions)
        Exit Function
    End If
    Dim ranges As Variant
    ranges = TestSettingsRangeSearchOrder()

    Dim rangeIndex As Long
    For rangeIndex = LBound(ranges) To UBound(ranges)
        If TestSettingValidationHasOptionsInRange(ThisWorkbook.Names.Item(CStr(ranges(rangeIndex))).RefersToRange, key, expectedOptions) Then
            TestAnySettingValidationHasOptions = True
            Exit Function
        End If
    Next rangeIndex
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сверяет formula-linked validation с ожидаемым списком, включая его порядок и длину.
' Ошибка чтения списка означает неуспешную проверку интерфейса.
Private Function TestSettingValidationHasOptionsInRange(ByVal settings As Object, ByVal key As String, ByVal expectedOptions As Variant) As Boolean
    On Error GoTo Failed

    Dim rowIndex As Long
    For rowIndex = 2 To settings.Rows.Count
        If StrComp(CStr(settings.Cells.Item(rowIndex, 1).Value2), key, vbTextCompare) = 0 Then
            TestSettingValidationHasOptionsInRange = TestValidationCellHasOptions(settings.Cells.Item(rowIndex, 2), expectedOptions)
            Exit Function
        End If
    Next rowIndex
Failed:
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет список вариантов выбранной группы аннотаций в ее собственной колонке.
' Это исключает случайный успех за счет настройки соседней группы.
Private Function TestPlotAnnotationValidationHasOptions(ByVal rowName As String, ByVal valueColumn As Long, ByVal expectedOptions As Variant) As Boolean
    On Error GoTo Failed

    Dim settings As Object
    Set settings = ThisWorkbook.Names.Item("rngPlotAnnotationSettings").RefersToRange

    Dim rowIndex As Long
    For rowIndex = 2 To settings.Rows.Count
        If StrComp(CStr(settings.Cells.Item(rowIndex, 1).Value2), rowName, vbTextCompare) = 0 Then
            TestPlotAnnotationValidationHasOptions = TestValidationCellHasOptions(settings.Cells.Item(rowIndex, valueColumn), expectedOptions)
            Exit Function
        End If
    Next rowIndex
Failed:
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Находит строку параметра Rectangles и проверяет validation ее ячейки значения.
Private Function TestRectanglesParameterValidationHasOptions(ByVal parameterName As String, ByVal expectedOptions As Variant) As Boolean
    On Error GoTo Failed
    TestRectanglesParameterValidationHasOptions = TestValidationCellHasOptions(TestFixtureConfigFieldCell(parameterName), expectedOptions)
Failed:
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет выпадающие списки в табличной части дополнительных рядов Rectangles.
' Helper ищет строку грани и нужный заголовок, поэтому вставка строк выше
' блока геометрии не меняет предмет проверки.
Private Function TestRectanglesExtraValidationHasOptions(ByVal faceCaption As String, ByVal headerCaption As String, _
        ByVal occurrenceIndex As Long, ByVal expectedOptions As Variant) As Boolean
    On Error GoTo Failed
    Dim field As String, face As String, side As Long, key As String
    face = Left$(faceCaption, 2): side = 1
    If InStr(faceCaption, "правая") Or InStr(faceCaption, "нижняя") Then side = 2
    field = "bind_": If InStr(1, headerCaption, "Положение", vbTextCompare) Then field = "loc_"
    key = "Rectangles." & face & "." & field & CStr(occurrenceIndex + 1) & "row_" & CStr(side)
    TestRectanglesExtraValidationHasOptions = TestValidationCellHasOptions(TestFixtureConfigFieldCell(key), expectedOptions)
Failed:
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сверяет dropdown первой строки LC с полным допустимым набором вариантов.
' Проверяется actual validation ячейки, а не текст каталога сборки.
Private Function TestLoadCombinationValidationHasOptions(ByVal valueColumn As Long, ByVal expectedOptions As Variant) As Boolean
    On Error GoTo Failed

    Dim loads As Object
    Set loads = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange
    TestLoadCombinationValidationHasOptions = TestValidationCellHasOptions(loads.Cells.Item(2, valueColumn), expectedOptions)
    Exit Function
Failed:
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Список сочетаний должен содержать вычисленные формулами названия всех профилей.
' Смена имени профиля обновляет список без изменения VBA.
Private Function TestLoadProfileValidationUsesDisplayNames() As Boolean
    On Error GoTo Failed

    Dim loads As Object
    Set loads = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange

    Dim formulaText As String
    formulaText = CStr(loads.Cells.Item(2, 5).Validation.Formula1)
    If Left$(formulaText, 1) <> "=" Then Exit Function

    Dim listRange As Object
    Set listRange = loads.Worksheet.Range(Mid$(formulaText, 2))

    Dim profiles As Object
    Set profiles = ThisWorkbook.Names.Item("rngCalculationProfiles").RefersToRange
    Dim displayRow As Long
    For displayRow = 1 To profiles.Rows.Count
        If CStr(profiles.Cells.Item(displayRow, 2).Value2) = "Profile.DisplayName" Then Exit For
    Next displayRow
    If displayRow > profiles.Rows.Count Then Exit Function

    Dim i As Long
    For i = 1 To listRange.Cells.Count
        If Not CBool(listRange.Cells.Item(i, 1).HasFormula) Then Exit Function
        If CStr(listRange.Cells.Item(i, 1).Value2) <> CStr(profiles.Cells.Item(displayRow, i + 2).Value2) Then Exit Function
    Next i
    TestLoadProfileValidationUsesDisplayNames = True
    Exit Function
Failed:
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет, что профиль в строке LC остается пользовательским вводом, не формулой.
Private Function TestLoadProfileFirstValueHasNoFormula() As Boolean
    On Error GoTo Failed

    Dim loads As Object
    Set loads = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange
    TestLoadProfileFirstValueHasNoFormula = Not CBool(loads.Cells.Item(2, 5).HasFormula)
    Exit Function
Failed:
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет точный состав и порядок списка допустимых значений ячейки.
' Если список не читается или задан неверным способом, проверка даёт False.
Private Function TestValidationCellHasOptions(ByVal target As Object, ByVal expectedOptions As Variant) As Boolean
    On Error GoTo Failed

    Dim actual As Variant, i As Long
    actual = TestValidationCellOptions(target)
    If UBound(actual) - LBound(actual) <> UBound(expectedOptions) - LBound(expectedOptions) Then Exit Function
    For i = LBound(expectedOptions) To UBound(expectedOptions)
        If CStr(actual(i - LBound(expectedOptions) + LBound(actual))) <> CStr(expectedOptions(i)) Then Exit Function
    Next i
    TestValidationCellHasOptions = True
    Exit Function
Failed:
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Собирает тестовый круг и арматуру из настоящего Config, подключает материалы,
' solver-options и каталог профилей. Reader сочетаний затем вызывается отдельно.
Private Function TestBuildUiBatch() As CBatchSectionCalculator
    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook

    Dim units As CUnitSystem, definition As CSectionDefinition, geom As CGeometryRegion
    Set units = New CUnitSystem: units.LoadFromSettings settings
    Set definition = TestFixtureWorkbookDefinition("Circle", ThisWorkbook, units)
    Set geom = TestFixtureGeometry(definition.ConcreteDomain)
    Dim mesh As CFiberMeshBuilder: Set mesh = New CFiberMeshBuilder
    mesh.BuildMesh geom, 20#, 20#, 1
    Dim rebars As CRebarLayout: Set rebars = TestFixtureLayout(definition)
    Dim materialProvider As CMaterialModelProvider
    Set materialProvider = New CMaterialModelProvider
    materialProvider.Initialize settings, units

    Dim section As CSectionModel
    Set section = TestBuildGeneratedSectionModel(mesh, rebars, "WorkbookInterface")

    Dim batch As CBatchSectionCalculator
    Set batch = New CBatchSectionCalculator
    batch.Initialize section, materialProvider
    batch.ApplySettings settings, units
    Dim profiles As CCalculationProfileCatalog
    Set profiles = New CCalculationProfileCatalog
    profiles.LoadFromWorkbook ThisWorkbook
    Set batch.ProfileCatalog = profiles
    Set TestBuildUiBatch = batch
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Готовит круговой SLS fixture через реальные ячейки Config и одну строку LC.
' Единицы/знаки приведены к внутренним; вызывающий тест работает на отдельной книге.
Private Sub TestPrepareCircleInput()
    TestSetSystemSetting "Units.Force.Input", "N"
    TestSetSystemSetting "Units.Moment.Input", "N*mm"
    TestSetSystemSetting "Units.Length.Input", "mm"
    TestSetSystemSetting "Units.Area.Input", "mm2"
    TestSetSystemSetting "Units.Stress.Input", "MPa"
    TestSetSystemSetting "Units.Curvature.Input", "1/mm"
    TestSetSystemSetting "Sign.N.User", "Tension"
    TestSetSystemSetting "Sign.Mx.User", "+Y tension"
    TestSetSystemSetting "Sign.My.User", "+X tension"
    TestSetSystemSetting "Geometry.Source", "Generated"
    TestSetSystemSetting "Geometry.Type", "Circle"
    TestFixtureResetReinforcement "Circle"
    TestSetSystemSetting "Circle.Diameter", "300"
    TestSetSystemSetting "Plot.AutoUpdateAfterCalculation", "Yes"
    TestSetSystemSetting "Plot.LoadCase", "LC1"
    TestSetSystemSetting "Rebar.AxisDistance", "40"
    TestSetSystemSetting "Rebar.Count", "8"
    TestSetSystemSetting "Rebar.Diameter", "20"
    TestSetSystemSetting "SLS.Crack.Allowable", "0.3"
    TestSetProfileSetting "PR2", "Visualization.State", "CrackedState"

    Dim loads As Object
    Set loads = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange
    TestClearDataRows loads
    loads.Cells.Item(2, 1).Value2 = "LC1"
    loads.Cells.Item(2, 2).Value2 = -100000#
    loads.Cells.Item(2, 3).Value2 = -4000000#
    loads.Cells.Item(2, 4).Value2 = -3000000#
    loads.Cells.Item(2, 5).Value2 = "PR2"
    loads.Cells.Item(2, 6).Value2 = ChrW$(&H3BB) & "*Mxy"
    loads.Cells.Item(2, 7).Value2 = "ui test"
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Готовит Rectangles fixture, активные грани и нагрузку через обычные таблицы Config.
' Неактивные грани отключаются явно, чтобы не зависеть от прежних пользовательских данных.
Private Sub TestPrepareRectanglesInput()
    TestSetSystemSetting "Units.Force.Input", "N"
    TestSetSystemSetting "Units.Moment.Input", "N*mm"
    TestSetSystemSetting "Units.Length.Input", "mm"
    TestSetSystemSetting "Units.Area.Input", "mm2"
    TestSetSystemSetting "Units.Stress.Input", "MPa"
    TestSetSystemSetting "Units.Curvature.Input", "1/mm"
    TestSetSystemSetting "Sign.N.User", "Tension"
    TestSetSystemSetting "Sign.Mx.User", "+Y tension"
    TestSetSystemSetting "Sign.My.User", "+X tension"
    TestSetSystemSetting "Geometry.Source", "Generated"
    TestSetSystemSetting "Geometry.Type", "Rectangles"
    TestFixtureResetReinforcement "Rectangles"
    TestSetSystemSetting "Rectangles.SectionType", "TwoRect"
    TestSetSystemSetting "Rectangles.UpperOffsetX", "0"
    TestSetSystemSetting "Plot.LoadCase", "LC_L"
    TestSetSystemSetting "Mesh.StepMode", "User"
    TestSetSystemSetting "Mesh.StepX", "40"
    TestSetSystemSetting "Mesh.StepY", "40"
    TestSetSystemSetting "Mesh.BoundarySubdivisions", "2"
    TestSetSystemSetting "Rectangles.B1", "160"
    TestSetSystemSetting "Rectangles.H1", "280"
    TestSetSystemSetting "Rectangles.B2", "360"
    TestSetSystemSetting "Rectangles.H2", "140"
    TestSetSystemSetting "Rectangles.H1.as_1", "40"
    TestSetSystemSetting "Rectangles.H1.as_2", "40"
    TestSetSystemSetting "Rectangles.H1.d_1", "16"
    TestSetSystemSetting "Rectangles.H1.d_2", "16"
    TestSetSystemSetting "Rectangles.H1.n_1", "3"
    TestSetSystemSetting "Rectangles.H1.n_2", "2"
    TestSetSystemSetting "Rectangles.H1.StartOffset1", "40"
    TestSetSystemSetting "Rectangles.H1.EndOffset1", "40"
    TestSetSystemSetting "Rectangles.H1.StartOffset2", "40"
    TestSetSystemSetting "Rectangles.H1.EndOffset2", "40"
    TestSetSystemSetting "Rectangles.H2.as_1", "40"
    TestSetSystemSetting "Rectangles.H2.as_2", "40"
    TestSetSystemSetting "Rectangles.H2.d_1", "16"
    TestSetSystemSetting "Rectangles.H2.d_2", "16"
    TestSetSystemSetting "Rectangles.H2.n_1", "0"
    TestSetSystemSetting "Rectangles.H2.n_2", "0"
    TestSetSystemSetting "Rectangles.H2.StartOffset1", "40"
    TestSetSystemSetting "Rectangles.H2.EndOffset1", "40"
    TestSetSystemSetting "Rectangles.H2.StartOffset2", "40"
    TestSetSystemSetting "Rectangles.H2.EndOffset2", "40"
    TestSetSystemSetting "Rectangles.B1.as_1", "40"
    TestSetSystemSetting "Rectangles.B1.as_2", "40"
    TestSetSystemSetting "Rectangles.B1.d_1", "16"
    TestSetSystemSetting "Rectangles.B1.d_2", "16"
    TestSetSystemSetting "Rectangles.B1.n_1", "2"
    TestSetSystemSetting "Rectangles.B1.n_2", "1"
    TestSetSystemSetting "Rectangles.B1.StartOffset1", "20"
    TestSetSystemSetting "Rectangles.B1.EndOffset1", "20"
    TestSetSystemSetting "Rectangles.B1.StartOffset2", "20"
    TestSetSystemSetting "Rectangles.B1.EndOffset2", "20"
    TestSetSystemSetting "Rectangles.B2.as_1", "40"
    TestSetSystemSetting "Rectangles.B2.as_2", "40"
    TestSetSystemSetting "Rectangles.B2.d_1", "16"
    TestSetSystemSetting "Rectangles.B2.d_2", "16"
    TestSetSystemSetting "Rectangles.B2.n_1", "2"
    TestSetSystemSetting "Rectangles.B2.n_2", "0"
    TestSetSystemSetting "Rectangles.B2.StartOffset1", "60"
    TestSetSystemSetting "Rectangles.B2.EndOffset1", "60"
    TestSetSystemSetting "Rectangles.B2.StartOffset2", "60"
    TestSetSystemSetting "Rectangles.B2.EndOffset2", "60"
    TestSetSystemSetting "SLS.Crack.Allowable", "0.3"

    Dim loads As Object
    Set loads = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange
    TestClearDataRows loads
    loads.Cells.Item(2, 1).Value2 = "LC_L"
    loads.Cells.Item(2, 2).Value2 = -80000#
    loads.Cells.Item(2, 3).Value2 = -1500000#
    loads.Cells.Item(2, 4).Value2 = -1000000#
    loads.Cells.Item(2, 5).Value2 = "PR2"
    loads.Cells.Item(2, 6).Value2 = ChrW$(&H3BB) & "*Mxy"
    loads.Cells.Item(2, 7).Value2 = "rectangles ui test"
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Настраивает книгу под типовой прочностной расчет Г-сечения:
' пользователь задает N и Mx, а несущая способность ищется увеличением Mx.
' Здесь специально выбран Capacity.SolutionStrategy = UltimateStrain, чтобы проверить,
' что быстрый изгибный путь работает без fallback на LoadMultiplier.
Private Sub TestPrepareUserRectanglesMomentUltimateInput()
    TestSetSystemSetting "Units.Force.Input", "tf"
    TestSetSystemSetting "Units.Moment.Input", "tf*m"
    TestSetSystemSetting "Units.Length.Input", "mm"
    TestSetSystemSetting "Units.Area.Input", "mm2"
    TestSetSystemSetting "Units.Stress.Input", "MPa"
    TestSetSystemSetting "Units.Curvature.Input", "1/mm"
    TestSetSystemSetting "Sign.N.User", "Compression"
    TestSetSystemSetting "Sign.Mx.User", "+Y tension"
    TestSetSystemSetting "Sign.My.User", "+X tension"
    TestSetSystemSetting "Geometry.Source", "Generated"
    TestSetSystemSetting "Geometry.Type", "Rectangles"
    TestSetSystemSetting "Capacity.SolutionStrategy", "UltimateStrain"
    TestSetSystemSetting "Capacity.MaxLambda", "64"
    TestSetSystemSetting "Capacity.ToleranceStrain", "0.00001"
    TestSetSystemSetting "Capacity.SolverMaxIterations", "60"
    TestSetSystemSetting "Solver.Method", "Newton"
    TestSetSystemSetting "General.DiagramExtension", "Yes"
    TestSetSystemSetting "Solver.MaxIterations", "80"
    TestSetSystemSetting "Solver.LoadSteps", "1"
    TestSetSystemSetting "Mesh.StepMode", "User"
    TestSetSystemSetting "Mesh.StepX", "50"
    TestSetSystemSetting "Mesh.StepY", "50"
    TestSetSystemSetting "Mesh.BoundarySubdivisions", "1"
    TestSetSystemSetting "Load.ReferenceOffsetX", "0"
    TestSetSystemSetting "Load.ReferenceOffsetY", "0"
    TestSetSystemSetting "Plot.LoadCase", "LC_MX"
    TestSetSystemSetting "General.ExecutionReportEnabled", "No"

    TestSetSystemSetting "Rectangles.H1", "550"
    TestSetSystemSetting "Rectangles.B1", "250"
    TestSetSystemSetting "Rectangles.H2", "250"
    TestSetSystemSetting "Rectangles.B2", "600"

    TestSetUserRectanglesMainRow "H1", 5, 5
    TestSetUserRectanglesMainRow "B1", 2, 2
    TestSetUserRectanglesMainRow "H2", 2, 2
    TestSetUserRectanglesMainRow "B2", 5, 5
    TestClearUserRectanglesExtraRows

    Dim loads As Object
    Set loads = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange
    TestClearDataRows loads
    loads.Cells.Item(2, 1).Value2 = "LC_MX"
    loads.Cells.Item(2, 2).Value2 = 200#
    loads.Cells.Item(2, 3).Value2 = 50#
    loads.Cells.Item(2, 4).ClearContents
    loads.Cells.Item(2, 5).Value2 = "PR1"
    loads.Cells.Item(2, 6).Value2 = ChrW$(&H3BB) & "*Mx"
    loads.Cells.Item(2, 7).Value2 = "moment ultimate regression"
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Настраивает книгу ровно под пользовательский пример с Г-сечением и осевым
' растяжением в tf. Это дополняет unit-тест batch-слоя проверкой полного
' Excel-пути: пользовательский знак N, единицы, чтение Config и запись Results.
Private Sub TestPrepareUserRectanglesAxialTensionInput()
    TestSetSystemSetting "Units.Force.Input", "tf"
    TestSetSystemSetting "Units.Moment.Input", "tf*m"
    TestSetSystemSetting "Units.Length.Input", "mm"
    TestSetSystemSetting "Units.Area.Input", "mm2"
    TestSetSystemSetting "Units.Stress.Input", "MPa"
    TestSetSystemSetting "Units.Curvature.Input", "1/mm"
    TestSetSystemSetting "Sign.N.User", "Compression"
    TestSetSystemSetting "Sign.Mx.User", "+Y tension"
    TestSetSystemSetting "Sign.My.User", "+X tension"
    TestSetSystemSetting "Geometry.Source", "Generated"
    TestSetSystemSetting "Geometry.Type", "Rectangles"
    TestSetSystemSetting "Solver.Method", "Newton"
    TestSetSystemSetting "General.DiagramExtension", "Yes"
    TestSetSystemSetting "Solver.MaxIterations", "80"
    TestSetSystemSetting "Solver.LoadSteps", "1"
    TestSetSystemSetting "Mesh.StepMode", "User"
    TestSetSystemSetting "Mesh.StepX", "50"
    TestSetSystemSetting "Mesh.StepY", "50"
    TestSetSystemSetting "Mesh.BoundarySubdivisions", "1"
    TestSetSystemSetting "Load.ReferenceOffsetX", "0"
    TestSetSystemSetting "Load.ReferenceOffsetY", "0"
    TestSetSystemSetting "Plot.LoadCase", "LC_OVER"
    TestSetSystemSetting "General.ExecutionReportEnabled", "Yes"

    TestSetSystemSetting "Rectangles.H1", "550"
    TestSetSystemSetting "Rectangles.B1", "250"
    TestSetSystemSetting "Rectangles.H2", "250"
    TestSetSystemSetting "Rectangles.B2", "600"

    TestSetUserRectanglesMainRow "H1", 5, 5
    TestSetUserRectanglesMainRow "B1", 2, 2
    TestSetUserRectanglesMainRow "H2", 2, 2
    TestSetUserRectanglesMainRow "B2", 5, 5
    TestClearUserRectanglesExtraRows

    Dim loads As Object
    Set loads = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange
    TestClearDataRows loads
    loads.Cells.Item(2, 1).Value2 = "LC_SAFE"
    loads.Cells.Item(2, 2).Value2 = -795#
    loads.Cells.Item(2, 3).ClearContents
    loads.Cells.Item(2, 4).ClearContents
    loads.Cells.Item(2, 5).Value2 = "PR2"
    loads.Cells.Item(2, 6).Value2 = ChrW$(&H3BB) & "*N"
    loads.Cells.Item(2, 7).Value2 = "inside physical range"

    loads.Cells.Item(3, 1).Value2 = "LC_OVER"
    loads.Cells.Item(3, 2).Value2 = -900#
    loads.Cells.Item(3, 3).ClearContents
    loads.Cells.Item(3, 4).ClearContents
    loads.Cells.Item(3, 5).Value2 = "PR2"
    loads.Cells.Item(3, 6).Value2 = ChrW$(&H3BB) & "*N"
    loads.Cells.Item(3, 7).Value2 = "uses extension"
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Включает сохранение прочностных и трещинных состояний выбранного тестового профиля.
' Каждая роль материала задается явно; Visualization выбирает уже сохраненный State.
Private Sub TestPrepareFullStateProfile(ByVal profileId As String)
    TestSetProfileSetting profileId, "Profile.DisplayName", "Полный snapshot"
    TestSetProfileSetting profileId, "Profile.Description", "Тестовая запись всех состояний"
    TestSetProfileSetting profileId, "Calculation.Strength.DirectState", "Yes"
    TestSetProfileSetting profileId, "Calculation.Strength.Capacity", "Yes"
    TestSetProfileSetting profileId, "Calculation.Crack.Width", "Yes"
    TestSetProfileSetting profileId, "MaterialModel.Strength.ValueSet", "ULS(I)"
    TestSetProfileSetting profileId, "MaterialModel.Strength.ConcreteDiagram", "TwoLine"
    TestSetProfileSetting profileId, "MaterialModel.Strength.ConcreteTension", "Ignore"
    TestSetProfileSetting profileId, "MaterialModel.Strength.SteelDiagram", "TwoLine"
    TestSetProfileSetting profileId, "MaterialModel.CrackInitiation.ValueSet", "SLS(II)"
    TestSetProfileSetting profileId, "MaterialModel.CrackInitiation.ConcreteDiagram", "ThreeLine"
    TestSetProfileSetting profileId, "MaterialModel.CrackInitiation.ConcreteTension", "UseDiagram"
    TestSetProfileSetting profileId, "MaterialModel.CrackInitiation.SteelDiagram", "TwoLine"
    TestSetProfileSetting profileId, "MaterialModel.CrackedState.ValueSet", "SLS(II)"
    TestSetProfileSetting profileId, "MaterialModel.CrackedState.ConcreteDiagram", "TwoLine"
    TestSetProfileSetting profileId, "MaterialModel.CrackedState.ConcreteTension", "Ignore"
    TestSetProfileSetting profileId, "MaterialModel.CrackedState.SteelDiagram", "TwoLine"
    TestSetProfileSetting profileId, "Visualization.State", "StrengthState"
    TestSetProfileSetting profileId, "Visualization.Quantity", "Stress"
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Заполняет обе стороны грани Rectangles фиксированными диаметрами и отступами.
' Количество первых рядов задается входами; дополнительные ряды очищаются отдельно.
Private Sub TestSetUserRectanglesMainRow(ByVal faceName As String, ByVal count1 As Long, ByVal count2 As Long)
    TestSetSystemSetting "Rectangles." & faceName & ".as_1", "40"
    TestSetSystemSetting "Rectangles." & faceName & ".as_2", "40"
    TestSetSystemSetting "Rectangles." & faceName & ".d_1", "32"
    TestSetSystemSetting "Rectangles." & faceName & ".d_2", "32"
    TestSetSystemSetting "Rectangles." & faceName & ".n_1", CStr(count1)
    TestSetSystemSetting "Rectangles." & faceName & ".n_2", CStr(count2)
    TestSetSystemSetting "Rectangles." & faceName & ".StartOffset1", "80"
    TestSetSystemSetting "Rectangles." & faceName & ".EndOffset1", "80"
    TestSetSystemSetting "Rectangles." & faceName & ".StartOffset2", "80"
    TestSetSystemSetting "Rectangles." & faceName & ".EndOffset2", "80"
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Отключает дополнительные ряды всех сторон пустыми диаметрами.
' Расположение и привязку задаёт явно для воспроизводимости следующих проверок.
Private Sub TestClearUserRectanglesExtraRows()
    Dim faces As Variant
    faces = Array("H1", "B1", "H2", "B2")
    Dim i As Long
    For i = LBound(faces) To UBound(faces)
        TestSetSystemSetting "Rectangles." & CStr(faces(i)) & ".d_2row_1", vbNullString
        TestSetSystemSetting "Rectangles." & CStr(faces(i)) & ".d_2row_2", vbNullString
        TestSetSystemSetting "Rectangles." & CStr(faces(i)) & ".d_3row_1", vbNullString
        TestSetSystemSetting "Rectangles." & CStr(faces(i)) & ".d_3row_2", vbNullString
        TestSetSystemSetting "Rectangles." & CStr(faces(i)) & ".loc_2row_1", "Stacked"
        TestSetSystemSetting "Rectangles." & CStr(faces(i)) & ".loc_3row_1", "Stacked"
        TestSetSystemSetting "Rectangles." & CStr(faces(i)) & ".bind_2row_1", "EachBar"
        TestSetSystemSetting "Rectangles." & CStr(faces(i)) & ".bind_3row_1", "EachBar"
        TestSetSystemSetting "Rectangles." & CStr(faces(i)) & ".loc_2row_2", "Stacked"
        TestSetSystemSetting "Rectangles." & CStr(faces(i)) & ".loc_3row_2", "Stacked"
        TestSetSystemSetting "Rectangles." & CStr(faces(i)) & ".bind_2row_2", "EachBar"
        TestSetSystemSetting "Rectangles." & CStr(faces(i)) & ".bind_3row_2", "EachBar"
    Next i
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Очищает только строки данных тестового диапазона, сохраняя его строку заголовков.
Private Sub TestClearDataRows(ByVal target As Object)
    Dim rowIndex As Long
    Dim colIndex As Long
    For rowIndex = 2 To target.Rows.Count
        For colIndex = 1 To target.Columns.Count
            target.Cells.Item(rowIndex, colIndex).ClearContents
        Next colIndex
    Next rowIndex
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет, что единицы у допусков равновесия не зашиты текстом.
' Эти значения вводятся пользователем в выбранных INPUT-единицах, поэтому
' колонка "Ед." должна ссылаться формулой на rngUnitSettings. Иначе при
' переходе, например, на tf*m можно случайно получить гигантский допуск
' момента и принять несбалансированное состояние как сошедшееся.
Private Sub TestSolverToleranceUnitLabels(ByRef stats As TUiTestStats)
    TestAssertTrue stats, "ui.units.solverToleranceN.dynamic", _
        TestSystemSettingUnitCellReferencesQuantity("Solver.ToleranceN", "Force")
    TestAssertTrue stats, "ui.units.solverToleranceMx.dynamic", _
        TestSystemSettingUnitCellReferencesQuantity("Solver.ToleranceMx", "Moment")
    TestAssertTrue stats, "ui.units.solverToleranceMy.dynamic", _
        TestSystemSettingUnitCellReferencesQuantity("Solver.ToleranceMy", "Moment")
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет, что настройки поиска несущей способности не выглядят как
' недооформленный блок: у безразмерных коэффициентов стоит "-", у счетчиков -
' "шт". Размерных величин в этом блоке сейчас нет, поэтому формулы единиц
' здесь не нужны; важно именно не оставлять пустую колонку "Ед.".
Private Sub TestCapacitySettingsUnitLabels(ByRef stats As TUiTestStats)
    TestAssertTextEquals stats, "ui.units.Capacity.SolutionStrategy", TestSystemSettingUnitText("Capacity.SolutionStrategy"), "-"
    TestAssertTextEquals stats, "ui.units.capacity.searchMethod", TestSystemSettingUnitText("Capacity.SearchMethod"), "-"
    TestAssertTextEquals stats, "ui.units.capacity.initialLambda", TestSystemSettingUnitText("Capacity.InitialLambda"), "-"
    TestAssertTextEquals stats, "ui.units.capacity.toleranceLambda", TestSystemSettingUnitText("Capacity.ToleranceLambda"), "-"
    TestAssertTextEquals stats, "ui.units.capacity.maxRetries", TestSystemSettingUnitText("Capacity.MaxRetries"), "шт"
    TestAssertTextEquals stats, "ui.units.capacity.baseLoadSteps", TestSystemSettingUnitText("Capacity.BaseLoadSteps"), "шт"
    TestAssertTextEquals stats, "ui.units.capacity.toleranceStrain", TestSystemSettingUnitText("Capacity.ToleranceStrain"), "-"
    TestAssertTextEquals stats, "ui.units.capacity.maxLambda", TestSystemSettingUnitText("Capacity.MaxLambda"), "-"
    TestAssertTextEquals stats, "ui.units.capacity.solverIterations", TestSystemSettingUnitText("Capacity.SolverMaxIterations"), "шт"
    TestAssertTrue stats, "ui.units.meshStepX.dynamic", _
        TestSystemSettingUnitCellReferencesQuantity("Mesh.StepX", "Length")
    TestAssertTrue stats, "ui.units.meshStepY.dynamic", _
        TestSystemSettingUnitCellReferencesQuantity("Mesh.StepY", "Length")
    TestAssertTrue stats, "ui.units.zeroMomentPerDepth.dynamic", _
        TestSystemSettingUnitCellReferencesQuantities("Calculation.ZeroMomentPerDepth", "Moment", "Length")
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает фактическую подпись единиц настройки для проверки собранной книги.
' Формульную зависимость от выбранных единиц проверяет отдельный helper.
Private Function TestSystemSettingUnitText(ByVal key As String) As String
    On Error GoTo Failed

    Dim settings As Object
    Set settings = ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange

    Dim rowIndex As Long
    For rowIndex = 2 To settings.Rows.Count
        If StrComp(CStr(settings.Cells.Item(rowIndex, 1).Value2), key, vbTextCompare) = 0 Then
            TestSystemSettingUnitText = Trim$(CStr(settings.Cells.Item(rowIndex, 3).Value2))
            Exit Function
        End If
    Next rowIndex

Failed:
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Меняет единицы текста обеих групп и проверяет пересчет формульной подписи.
' Исходные значения восстанавливаются и после неуспешного чтения fixture.
Private Function TestPlotAnnotationTextUnitCellIsDynamic(ByVal rowName As String) As Boolean
    On Error GoTo Failed

    Dim settings As Object
    Set settings = ThisWorkbook.Names.Item("rngPlotAnnotationSettings").RefersToRange

    Dim textUnitsRow As Long, targetRow As Long, rowIndex As Long
    For rowIndex = 2 To settings.Rows.Count
        If StrComp(CStr(settings.Cells.Item(rowIndex, 1).Value2), "TextUnits", vbTextCompare) = 0 Then textUnitsRow = rowIndex
        If StrComp(CStr(settings.Cells.Item(rowIndex, 1).Value2), rowName, vbTextCompare) = 0 Then targetRow = rowIndex
    Next rowIndex
    If textUnitsRow = 0 Or targetRow = 0 Then Exit Function
    If Left$(CStr(settings.Cells.Item(targetRow, 4).Formula), 1) <> "=" Then Exit Function

    Dim oldRebarUnits As Variant, oldDimensionUnits As Variant
    Dim valuesSaved As Boolean
    oldRebarUnits = settings.Cells.Item(textUnitsRow, 2).Value2
    oldDimensionUnits = settings.Cells.Item(textUnitsRow, 3).Value2
    valuesSaved = True

    settings.Cells.Item(textUnitsRow, 2).Value2 = "pt"
    settings.Cells.Item(textUnitsRow, 3).Value2 = "pt"
    settings.Worksheet.Calculate
    TestPlotAnnotationTextUnitCellIsDynamic = (CStr(settings.Cells.Item(targetRow, 4).Value2) = "pt")

CleanUp:
    If Not valuesSaved Then Exit Function
    settings.Cells.Item(textUnitsRow, 2).Value2 = oldRebarUnits
    settings.Cells.Item(textUnitsRow, 3).Value2 = oldDimensionUnits
    settings.Worksheet.Calculate
    Exit Function
Failed:
    If valuesSaved Then Resume CleanUp
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Читает численную настройку заданной группы аннотаций для независимой UI-проверки.
' Метод не строит схему и не запускает расчет модели.
Private Function TestPlotAnnotationSettingValue(ByVal rowName As String, ByVal valueColumn As Long) As Double
    On Error GoTo Failed

    Dim settings As Object
    Set settings = ThisWorkbook.Names.Item("rngPlotAnnotationSettings").RefersToRange

    Dim rowIndex As Long
    For rowIndex = 2 To settings.Rows.Count
        If StrComp(CStr(settings.Cells.Item(rowIndex, 1).Value2), rowName, vbTextCompare) = 0 Then
            TestPlotAnnotationSettingValue = CDbl(settings.Cells.Item(rowIndex, valueColumn).Value2)
            Exit Function
        End If
    Next rowIndex

Failed:
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает пользовательскую подпись без учета регистра, сохраняя оба текста при отказе.
Private Sub TestAssertTextEquals(ByRef stats As TUiTestStats, ByVal name As String, _
        ByVal actual As String, ByVal expected As String)
    If StrComp(actual, expected, vbTextCompare) = 0 Then
        stats.Passed = stats.Passed + 1
        TestAppendLine stats, "OK: " & name & "; actual=" & actual
    Else
        stats.Failed = stats.Failed + 1
        TestAppendLine stats, "FAIL: " & name & "; actual=" & actual & "; expected=" & expected
    End If
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Группа файловых helpers работает с отчетами, созданными текущим UI-тестом.
' Удаление ниже допускается только для известного пути тестового артефакта.
Private Function TestFileExists(ByVal path As String) As Boolean
    TestFileExists = CreateObject("Scripting.FileSystemObject").FileExists(path)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Удаляет существующий отчет по известному пути текущего UI-теста;
' отсутствие файла допустимо, произвольные пользовательские пути не передаются.
Private Sub TestDeleteFileIfExists(ByVal path As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(path) Then fso.DeleteFile path, True
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Читает настоящий Unicode-отчет для сравнения с выводимыми комментариями результата.
Private Function TestReadTextFile(ByVal path As String) As String
    Dim stream As Object
    Set stream = CreateObject("Scripting.FileSystemObject").OpenTextFile(path, 1, False, -1)
    TestReadTextFile = stream.ReadAll
    stream.Close
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Группа assertions ведет счет проверок actual интерфейса и протоколирует отказ.
' Проверки чисел ниже используют фиксированный допуск вызывающего сценария.
Private Sub TestAssertTrue(ByRef stats As TUiTestStats, ByVal name As String, ByVal condition As Boolean)
    If condition Then
        stats.Passed = stats.Passed + 1
        TestAppendLine stats, "OK: " & name
    Else
        stats.Failed = stats.Failed + 1
        TestAppendLine stats, "FAIL: " & name
    End If
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сверяет абсолютное отклонение UI-величины, сохраняя actual/expected в отчете.
Private Sub TestAssertClose(ByRef stats As TUiTestStats, ByVal name As String, ByVal actual As Double, _
        ByVal expected As Double, ByVal tolerance As Double)
    If Abs(actual - expected) <= tolerance Then
        stats.Passed = stats.Passed + 1
        TestAppendLine stats, "OK: " & name & "; actual=" & TestFormatNumberInvariant(actual) & _
            "; expected=" & TestFormatNumberInvariant(expected)
    Else
        stats.Failed = stats.Failed + 1
        TestAppendLine stats, "FAIL: " & name & "; actual=" & TestFormatNumberInvariant(actual) & _
            "; expected=" & TestFormatNumberInvariant(expected)
    End If
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Добавляет диагностическую строку к протоколу интерфейсного набора.
Private Sub TestAppendLine(ByRef stats As TUiTestStats, ByVal text As String)
    stats.Report = stats.Report & text & vbCrLf
    ' ДЛЯ ТЕСТОВ: сохраняем текущий этап вне COM-вызова. При зависании или
    ' ошибке Excel runner сможет определить тест, не обращаясь к busy Excel.
    If Left$(text, 5) = "RUN: " Then
        Dim stream As Object
        Set stream = CreateObject("Scripting.FileSystemObject").OpenTextFile( _
            ThisWorkbook.Path & "\RC_NDM_ui_test_progress.txt", 8, True, -1)
        stream.WriteLine Format$(Now, "yyyy-mm-dd hh:nn:ss") & " " & text
        stream.Close
    End If
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Форматирует число для текстового протокола с десятичной точкой; исходные расчётные значения не меняются.
Private Function TestFormatNumberInvariant(ByVal value As Double) As String
    TestFormatNumberInvariant = Replace$(Format$(value, "0.############"), ",", ".")
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' ДЛЯ ТЕСТОВ
' Изолирует один полный расчет книги от остальных UI-сценариев. Диагностика
' используется для поиска роста памяти; основной regression-набор не заменяет.
Public Function TestRunAudit02SingleCalculationDiagnostic() As String
    On Error GoTo Failed
    Dim stats As TUiTestStats
    TestSetSystemSetting "General.ExecutionReportEnabled", "No"
    TestCircleWorkbookRunWritesResults stats
    TestRunAudit02SingleCalculationDiagnostic = stats.Report & _
        "TOTAL_UI_DIAGNOSTIC: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    Exit Function
Failed:
    TestRunAudit02SingleCalculationDiagnostic = stats.Report & "RUNTIME ERROR: " & CStr(Err.Number) & "; " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' ДЛЯ ТЕСТОВ
' Изолирует запись execution_report на свежей книге, чтобы отличить расходы
' отчета от накопления данных между последовательными UI-расчетами.
Public Function TestRunAudit02ReportDiagnostic() As String
    On Error GoTo Failed
    Dim stats As TUiTestStats
    TestExecutionReportFile stats
    TestRunAudit02ReportDiagnostic = stats.Report & _
        "TOTAL_UI_DIAGNOSTIC: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    Exit Function
Failed:
    TestRunAudit02ReportDiagnostic = stats.Report & "RUNTIME ERROR: " & CStr(Err.Number) & "; " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' повторяет расчётную часть обычного запуска без записи таблиц
' и схемы. Возвращает число решений и время без изменения физических ожиданий.
Public Function TestRunAudit02BatchOnlyDiagnostic() As String
    TestRunAudit02BatchOnlyDiagnostic = TestRunAudit02WorkbookPhases(False)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' измеряет последовательную запись таблиц по одному готовому результату.
' Каждая отметка содержит время и память собственного тестового Excel.
Public Function TestRunAudit02WriterPhasesDiagnostic() As String
    TestRunAudit02WriterPhasesDiagnostic = TestRunAudit02WorkbookPhases(True)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' ДЛЯ ТЕСТОВ
' Общая подготовка независимого диагностического сценария без основного UI.
Private Function TestRunAudit02WorkbookPhases(ByVal writeResults As Boolean) As String
    On Error GoTo Failed
    TestPrepareCircleInput
    TestSetSystemSetting "General.ExecutionReportEnabled", "No"
    Dim started As Double
    started = Timer
    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook
    Dim units As CUnitSystem
    Set units = New CUnitSystem
    units.LoadFromSettings settings
    Dim materials As CMaterialModelProvider
    Set materials = New CMaterialModelProvider
    materials.Initialize settings, units
    Dim section As CSectionModel
    Set section = BuildWorkbookSectionModel(ThisWorkbook, settings, units)
    Dim props As CSectionPropertiesCalculator
    Set props = New CSectionPropertiesCalculator
    props.CalculateConcrete section
    Dim batch As CBatchSectionCalculator
    Set batch = New CBatchSectionCalculator
    batch.Initialize section, materials
    batch.ApplySettings settings, units
    Dim profiles As CCalculationProfileCatalog
    Set profiles = New CCalculationProfileCatalog
    profiles.LoadFromWorkbook ThisWorkbook
    Set batch.ProfileCatalog = profiles
    Dim reader As CLoadCombinationReader
    Set reader = New CLoadCombinationReader
    reader.LoadFromWorkbook ThisWorkbook, batch, units
    batch.ApplyLoadReference props.CentroidX, props.CentroidY, props.CentroidX, props.CentroidY
    batch.Execute
    TestRunAudit02WorkbookPhases = "INFO: batchOnly; status=" & batch.ResultAt(1).Status & _
        "; elapsedSec=" & TestFormatNumberInvariant(Timer - started) & "; solves=" & CStr(batch.SolverCallCount)
    If writeResults Then
        Dim stats As TUiTestStats
        TestAppendLine stats, "RUN: batch " & TestAudit02ExcelMemory()
        Dim strengthWriter As CStrengthSummaryWriter
        Set strengthWriter = New CStrengthSummaryWriter
        strengthWriter.WriteSummary ThisWorkbook, batch, units
        TestAppendLine stats, "RUN: strength " & TestAudit02ExcelMemory()
        Dim crackWriter As CCrackSummaryWriter
        Set crackWriter = New CCrackSummaryWriter
        crackWriter.WriteSummary ThisWorkbook, batch, units
        TestAppendLine stats, "RUN: crack " & TestAudit02ExcelMemory()
        Dim stabilityWriter As CStabilitySummaryWriter
        Set stabilityWriter = New CStabilitySummaryWriter
        stabilityWriter.WriteSummary ThisWorkbook, batch, units
        TestAppendLine stats, "RUN: stability " & TestAudit02ExcelMemory()
        Dim summaryWriter As CBatchResultWriter
        Set summaryWriter = New CBatchResultWriter
        summaryWriter.WriteSummary ThisWorkbook, batch, units, section
        TestAppendLine stats, "RUN: summary " & TestAudit02ExcelMemory()
        Dim ndmWriter As CNDMResultsWriter
        Set ndmWriter = New CNDMResultsWriter
        ndmWriter.WriteResults ThisWorkbook, section, PrepareSectionSnapshot(section, materials), batch, units
        TestAppendLine stats, "RUN: NDM " & TestAudit02ExcelMemory()
        TestRunAudit02WorkbookPhases = TestRunAudit02WorkbookPhases & vbCrLf & stats.Report
    End If
    Exit Function
Failed:
    TestRunAudit02WorkbookPhases = "RUNTIME ERROR: " & CStr(Err.Number) & "; " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' ДЛЯ ТЕСТОВ
' Читает память процесса через WMI, не меняя настройки или состояние Excel.
Private Function TestAudit02ExcelMemory() As String
    Dim process As Object
    For Each process In GetObject("winmgmts:").ExecQuery( _
            "SELECT ProcessId, WorkingSetSize, PrivatePageCount FROM Win32_Process WHERE Name='EXCEL.EXE'")
        If Len(TestAudit02ExcelMemory) > 0 Then TestAudit02ExcelMemory = TestAudit02ExcelMemory & "; "
        TestAudit02ExcelMemory = TestAudit02ExcelMemory & "pid=" & CStr(process.ProcessId) & _
            "; workingSet=" & CStr(process.WorkingSetSize) & "; privateBytes=" & CStr(process.PrivatePageCount)
    Next process
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' выполняет неизмененные UI/Config-проверки до адресного LC-блока.
' Runner может сохранить полученную книгу и воспроизвести LC на новом Excel,
' отличив зависимость от содержимого Results от накопления памяти процесса.
Public Function RunAudit03PreLoadTableDiagnosticTests() As String
    mAudit03StopBeforeLoadTable = True
    RunAudit03PreLoadTableDiagnosticTests = RunWorkbookInterfaceTests()
    mAudit03StopBeforeLoadTable = False
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' сохраняет неизмененный входной UI-prefix перед unit/sign-блоком.
' Повтор этого блока на сохраненной копии в свежем Excel отличает влияние
' Config/Results от ресурсов, накопленных предшествующими тестовыми вызовами.
Public Function RunAudit03InputPrefixDiagnosticTests() As String
    mAudit03StopBeforeUnitSign = True
    RunAudit03InputPrefixDiagnosticTests = RunWorkbookInterfaceTests()
    mAudit03StopBeforeUnitSign = False
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' записывает численные Results без обновления схемы.
' Помогает определить, на каком этапе растёт память: расчёт, таблицы или диаграмма.
Public Function TestRunAudit02NoPlotDiagnostic() As String
    On Error GoTo Failed
    TestPrepareCircleInput
    TestSetSystemSetting "General.ExecutionReportEnabled", "No"
    TestSetSystemSetting "Plot.AutoUpdateAfterCalculation", "No"
    TestRunAudit02NoPlotDiagnostic = TestFixtureRunSectionWorkflow(ThisWorkbook, False)
    Exit Function
Failed:
    TestRunAudit02NoPlotDiagnostic = "RUNTIME ERROR: " & CStr(Err.Number) & "; " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' ДЛЯ ТЕСТОВ
' Проверяет сохраненные Results после смены материальных настроек без пересчета.
' Отдельный entrypoint позволяет повторить проверку на независимой книге.
Public Function RunAudit02SavedSnapshotTests() As String
    On Error GoTo Failed
    Dim stats As TUiTestStats
    TestAudit02SavedResultsIgnoreMaterialChanges stats
    RunAudit02SavedSnapshotTests = stats.Report & _
        "TOTAL_AUDIT02_SNAPSHOT: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    Exit Function
Failed:
    RunAudit02SavedSnapshotTests = stats.Report & "RUNTIME ERROR: " & CStr(Err.Number) & "; " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' ДЛЯ ТЕСТОВ
' Меняет Config после завершенного расчета, но не его сериализованный снимок.
' Reader и повторная схема обязаны сохранить напряжения, деформации, признаки
' State и таблицы диаграмм; чтение геометрии для AutoCAD также не решает НДС.
Private Sub TestAudit02SavedResultsIgnoreMaterialChanges(ByRef stats As TUiTestStats)
    TestPrepareCircleInput
    TestSetSystemSetting "Plot.AutoUpdateAfterCalculation", "No"
    Dim materialKeys As Variant
    materialKeys = Array("General.DiagramExtension", "Concrete.E", "Concrete.R.ULS(I)", _
        "Concrete.R.SLS(II)", "Steel.E", "Steel.R.ULS(I)", "Steel.R.SLS(II)")
    Dim oldValues(0 To 6) As String
    Dim i As Long
    For i = 0 To UBound(materialKeys)
        oldValues(i) = CStr(TestGetSystemSetting(CStr(materialKeys(i))))
    Next i
    Dim oldQuantity As String
    oldQuantity = TestProfileSettingValue("PR2", "Visualization.Quantity")
    On Error GoTo RestoreFailed

    Dim message As String
    message = TestFixtureRunSectionWorkflow(ThisWorkbook, False)
    TestAssertTrue stats, "audit02.saved.run", InStr(1, message, "Расчет завершен", vbTextCompare) > 0
    Dim tables As Variant
    tables = Array("rngNDMSectionGeometry", "rngNDMSectionContours", "rngNDMElementResults", "rngNDMSectionProperties", _
        "rngNDMSectionAnnotations", "rngNDMMaterialDiagrams", "rngBatchSummary")
    Dim beforeTables() As Variant
    ReDim beforeTables(0 To UBound(tables))
    For i = 0 To UBound(tables)
        beforeTables(i) = TestResultTable(CStr(tables(i)))
    Next i
    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook
    Dim beforeStress As CSectionPlotDataReader
    Set beforeStress = New CSectionPlotDataReader
    TestSetProfileSetting "PR2", "Visualization.Quantity", "Stress"
    beforeStress.LoadFromWorkbook ThisWorkbook, settings
    TestSetSystemSetting "AutoCAD.Export.CombinationID", "LC1"
    Dim beforeExport As String
    beforeExport = Audit02ReadExportSnapshotForTests(ThisWorkbook)
    Dim beforeStrain As CSectionPlotDataReader
    Set beforeStrain = New CSectionPlotDataReader
    TestSetProfileSetting "PR2", "Visualization.Quantity", "Strain"
    beforeStrain.LoadFromWorkbook ThisWorkbook, settings
    Dim solves As Long
    solves = SectionEquilibriumSolveCount()

    If StrComp(oldValues(0), "Yes", vbTextCompare) = 0 Then
        TestSetSystemSetting "General.DiagramExtension", "No"
    Else
        TestSetSystemSetting "General.DiagramExtension", "Yes"
    End If
    For i = 1 To UBound(materialKeys)
        TestSetSystemSetting CStr(materialKeys(i)), TestFormatNumberInvariant(CDbl(oldValues(i)) * 0.5)
    Next i
    settings.LoadFromWorkbook ThisWorkbook
    Dim afterData As CSectionPlotDataReader
    Set afterData = New CSectionPlotDataReader
    TestSetProfileSetting "PR2", "Visualization.Quantity", "Stress"
    afterData.LoadFromWorkbook ThisWorkbook, settings
    TestAssertAudit02PlotSnapshot stats, "stress", beforeStress, afterData
    TestAssertTextEquals stats, "audit02.saved.exportSnapshot", _
        Audit02ReadExportSnapshotForTests(ThisWorkbook), beforeExport
    TestSetProfileSetting "PR2", "Visualization.Quantity", "Strain"
    afterData.LoadFromWorkbook ThisWorkbook, settings
    TestAssertAudit02PlotSnapshot stats, "strain", beforeStrain, afterData
    UpdateSectionPlotForWorkbook ThisWorkbook
    Dim exportedSection As CSectionModel
    Set exportedSection = ReadSectionGeometryFromResults(ThisWorkbook)
    TestAssertTrue stats, "audit02.saved.exportGeometry", exportedSection.ConcreteCount > 0
    TestAssertTrue stats, "audit02.saved.noSolve", SectionEquilibriumSolveCount() = solves
    Dim afterTable As Variant
    For i = 0 To UBound(tables)
        afterTable = TestResultTable(CStr(tables(i)))
        TestAssertTrue stats, "audit02.saved.table." & CStr(tables(i)), _
            TestAudit02SnapshotTablesEqual(beforeTables(i), afterTable)
    Next i
    GoTo RestoreSettings
RestoreFailed:
    Dim failureText As String
    failureText = CStr(Err.Number) & "; " & Err.Description
    TestAssertTrue stats, "audit02.saved.runtime." & failureText, False
RestoreSettings:
    For i = 0 To UBound(materialKeys)
        TestSetSystemSetting CStr(materialKeys(i)), oldValues(i)
    Next i
    TestSetProfileSetting "PR2", "Visualization.Quantity", oldQuantity
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' ДЛЯ ТЕСТОВ
' Сравнивает сериализованную раскраску и плоскость без повторной оценки материала.
Private Sub TestAssertAudit02PlotSnapshot(ByRef stats As TUiTestStats, ByVal quantity As String, _
        ByVal beforeData As CSectionPlotDataReader, ByVal afterData As CSectionPlotDataReader)
    Dim prefix As String
    prefix = "audit02.saved." & quantity
    TestAssertTrue stats, prefix & ".nonempty", beforeData.Count > 0
    TestAssertTrue stats, prefix & ".count", beforeData.Count = afterData.Count
    TestAssertTextEquals stats, prefix & ".state", afterData.StateType, beforeData.StateType
    TestAssertTextEquals stats, prefix & ".status", afterData.DirectStateStatus, beforeData.DirectStateStatus
    TestAssertTrue stats, prefix & ".extension", afterData.ExtensionUsed = beforeData.ExtensionUsed
    TestAssertClose stats, prefix & ".epsilon0", afterData.Epsilon0, beforeData.Epsilon0, 0#
    TestAssertClose stats, prefix & ".kappaX", afterData.KappaX, beforeData.KappaX, 0#
    TestAssertClose stats, prefix & ".kappaY", afterData.KappaY, beforeData.KappaY, 0#
    Dim i As Long
    For i = 1 To beforeData.Count
        TestAssertTextEquals stats, prefix & ".id." & CStr(i), afterData.ElementID(i), beforeData.ElementID(i)
        TestAssertTextEquals stats, prefix & ".physical." & CStr(i), afterData.PhysicalState(i), beforeData.PhysicalState(i)
        TestAssertClose stats, prefix & ".value." & CStr(i), afterData.ResultValue(i), beforeData.ResultValue(i), 0#
    Next i
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' сравнивает все ячейки двух сохранённых наборов данных,
' включая характеристики материалов и фактический режим продолжения диаграмм.
Private Function TestAudit02SnapshotTablesEqual(ByRef beforeTable As Variant, ByRef afterTable As Variant) As Boolean
    If UBound(beforeTable, 1) <> UBound(afterTable, 1) Then Exit Function
    If UBound(beforeTable, 2) <> UBound(afterTable, 2) Then Exit Function
    Dim r As Long
    Dim c As Long
    For r = 1 To UBound(beforeTable, 1)
        For c = 1 To UBound(beforeTable, 2)
            If CStr(beforeTable(r, c)) <> CStr(afterTable(r, c)) Then Exit Function
        Next c
    Next r
    TestAudit02SnapshotTablesEqual = True
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' ========================== ДЛЯ ТЕСТОВ ==========================
' Проверяет контракт reader-а через реальные диапазоны независимой книги.
' Временный лист удаляется даже при ошибке; исходные Config-ячейки не меняются.
Public Function RunAudit03ReaderTests() As String
    Dim stats As TUiTestStats
    TestAudit03ReaderContract stats
    TestAppendLine stats, "TOTAL_AUDIT03_READER: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunAudit03ReaderTests = stats.Report
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Выполняет reader-кейсы в общей статистике suite и восстанавливает временный лист.
' Отдельный entrypoint и полный UI-набор используют те же assertions без дублей.
Private Sub TestAudit03ReaderContract(ByRef stats As TUiTestStats)
    Dim sheet As Object
    Dim oldAlerts As Boolean
    oldAlerts = Application.DisplayAlerts
    On Error GoTo Failed
    Set sheet = ThisWorkbook.Worksheets.Add
    sheet.Name = "__Audit03Reader"
    Dim columns As Variant
    For Each columns In Array(5, 6, 7, 9)
        TestAudit03ReaderWidth stats, sheet, CLng(columns)
    Next columns
    TestAudit03ReaderRows stats, sheet
    TestAudit03ReaderTinyUnits stats, sheet
    GoTo CleanUp
Failed:
    TestAssertTrue stats, "audit03.reader.entry.runtime." & CStr(Err.Number) & "." & Err.Description, False
CleanUp:
    On Error Resume Next
    Application.DisplayAlerts = False
    If Not sheet Is Nothing Then sheet.Delete
    Application.DisplayAlerts = oldAlerts
    On Error GoTo 0
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет 5/6/7+ колонок: необязательный comment, path и исходный пропуск.
' Ошибка структуры четырех колонок должна быть контролируемой, не ошибкой массива.
Private Sub TestAudit03ReaderWidth(ByRef stats As TUiTestStats, ByVal sheet As Object, ByVal columns As Long)
    On Error GoTo Failed
    sheet.Cells.ClearContents
    Dim data() As Variant
    ReDim data(1 To 3, 1 To columns)
    data(1, 1) = "CombinationID": data(1, 2) = "N": data(1, 3) = "Mx"
    data(1, 4) = "My": data(1, 5) = "ProfileId"
    data(3, 1) = "WIDTH_" & CStr(columns): data(3, 2) = -10000#
    data(3, 3) = 10000000#: data(3, 4) = 20000000#: data(3, 5) = "PR1"
    Dim expectedComment As String
    If columns = 6 Then data(3, 6) = "Комментарий шести колонок"
    If columns >= 7 Then
        data(3, 6) = "LambdaMx"
        data(3, 7) = "Комментарий семи колонок"
    End If
    If columns = 6 Then expectedComment = CStr(data(3, 6))
    If columns >= 7 Then expectedComment = CStr(data(3, 7))
    sheet.Range("A1").Resize(3, columns).Value2 = data
    Dim reader As CLoadCombinationReader
    Set reader = New CLoadCombinationReader
    Dim batch As CBatchSectionCalculator
    Set batch = TestBuildUiBatch()
    reader.LoadFromRange sheet.Range("A1").Resize(3, columns), batch
    Dim prefix As String
    prefix = "audit03.reader.width." & CStr(columns)
    TestAssertTrue stats, prefix & ".count", batch.Count = 1
    If batch.Count <> 1 Then Exit Sub
    TestAssertTextEquals stats, prefix & ".id", batch.CombinationID(1), "WIDTH_" & CStr(columns)
    TestAssertTextEquals stats, prefix & ".profile", batch.InputProfileId(1), "PR1"
    TestAssertTextEquals stats, prefix & ".comment", batch.CombinationName(1), expectedComment
    TestAssertTrue stats, prefix & ".sourceSlot", batch.SourceDataOffset(1) = 2
    TestAssertClose stats, prefix & ".n", batch.N(1), -10000#, 0#
    TestAssertClose stats, prefix & ".mx", batch.UserMx(1), 10000000#, 0#
    TestAssertClose stats, prefix & ".my", batch.UserMy(1), 20000000#, 0#
    If columns >= 7 Then
        batch.Execute
        TestAssertTextEquals stats, prefix & ".capacityPath", batch.ResultAt(1).StrengthResult.Capacity.PathResolved, "LambdaMx"
    End If
    Dim errorNumber As Long
    On Error Resume Next
    reader.LoadFromRange sheet.Range("A1:D3"), batch
    errorNumber = Err.Number
    Err.Clear
    On Error GoTo Failed
    TestAssertTrue stats, prefix & ".shortRangeControlled", errorNumber = vbObjectError + 3955
    Exit Sub
Failed:
    TestAssertTrue stats, "audit03.reader.width." & CStr(columns) & ".runtime." & CStr(Err.Number) & "." & Err.Description, False
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Нечисловые N/M, ошибки формул и metadata не теряются как нулевые LC.
' Числовая строка/формула и последующий корректный LC остаются в результатах.
Private Sub TestAudit03ReaderRows(ByRef stats As TUiTestStats, ByVal sheet As Object)
    On Error GoTo Failed
    sheet.Cells.ClearContents
    Dim data(1 To 13, 1 To 7) As Variant
    Dim r As Long
    For r = 2 To 13
        data(r, 1) = "ROW_" & CStr(r)
        data(r, 2) = 0#: data(r, 3) = 0#: data(r, 4) = 0#
        data(r, 5) = "PR1"
    Next r
    data(2, 2) = "abc": data(3, 3) = "abc": data(4, 4) = "abc"
    data(5, 2) = CVErr(xlErrDiv0)
    data(6, 2) = "   "
    For r = 1 To 7
        data(7, r) = Empty
    Next r
    data(8, 2) = CStr(42.5): data(8, 6) = "LambdaN"
    data(9, 6) = "LambdaN"
    data(10, 1) = " ": data(10, 2) = 1000#
    data(11, 2) = 1000#: data(11, 5) = CVErr(xlErrNA)
    data(12, 2) = 1000#: data(12, 6) = CVErr(xlErrRef)
    data(13, 1) = "VALID_LAST": data(13, 2) = -1000#: data(13, 3) = 10000000#
    data(13, 6) = "LambdaMxy"
    sheet.Range("A1:G13").Value2 = data
    sheet.Range("B5").Formula = "=1/0"
    sheet.Range("B5").Calculate
    sheet.Range("B8").NumberFormat = "@"
    sheet.Range("B8").Value2 = CStr(42.5)
    TestAssertTrue stats, "audit03.reader.rows.stringFixture", VarType(sheet.Range("B8").Value2) = vbString
    sheet.Range("B9").Formula = "=2+3"
    sheet.Range("B9").Calculate
    Dim reader As CLoadCombinationReader
    Set reader = New CLoadCombinationReader
    Dim batch As CBatchSectionCalculator
    Set batch = TestBuildUiBatch()
    reader.LoadFromRange sheet.Range("A1:G13"), batch
    batch.SetSP35Table721 ThisWorkbook.Names.Item("rngSP35Table721").RefersToRange.Value2
    TestAssertTrue stats, "audit03.reader.rows.count", batch.Count = 10
    If batch.Count <> 10 Then Exit Sub
    TestAssertTrue stats, "audit03.reader.rows.sourceGap", batch.SourceDataOffset(5) = 7
    TestAssertTextEquals stats, "audit03.reader.rows.lastId", batch.CombinationID(10), "VALID_LAST"
    TestAssertTrue stats, "audit03.reader.rows.lastSlot", batch.SourceDataOffset(10) = 12
    TestAssertClose stats, "audit03.reader.rows.numericString", batch.N(5), 42.5, 0#
    TestAssertClose stats, "audit03.reader.rows.numericFormula", batch.N(6), 5#, 0#
    batch.Execute
    For r = 1 To 4
        TestAssertTextEquals stats, "audit03.reader.rows.invalid." & CStr(r), batch.ResultAt(r).Status, "InputErr"
        TestAssertTrue stats, "audit03.reader.rows.comment." & CStr(r), Len(batch.ResultAt(r).OverallMeta.ResultComment) > 0
    Next r
    For r = 7 To 9
        TestAssertTextEquals stats, "audit03.reader.rows.invalid." & CStr(r), batch.ResultAt(r).Status, "InputErr"
    Next r
    TestAppendLine stats, "COMMENT: audit03.reader.rows.valid; " & batch.ResultAt(10).OverallMeta.ResultComment
    TestAssertTrue stats, "audit03.reader.rows.validReachedResults", batch.ResultAt(10).Status <> "InputErr"
    Exit Sub
Failed:
    TestAssertTrue stats, "audit03.reader.rows.runtime." & CStr(Err.Number) & "." & Err.Description, False
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Один tiny физический вектор не исчезает при смене N/N*mm на kN/kN*m.
' Переполнение единичного пересчета остается ошибочной строкой, следующая читается.
Private Sub TestAudit03ReaderTinyUnits(ByRef stats As TUiTestStats, ByVal sheet As Object)
    On Error GoTo Failed
    sheet.Cells.ClearContents
    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    Dim units As CUnitSystem
    Dim reader As CLoadCombinationReader
    Set reader = New CLoadCombinationReader
    Dim batch As CBatchSectionCalculator
    Dim mode As Long
    For mode = 0 To 1
        sheet.Range("K1").Value2 = "Key": sheet.Range("L1").Value2 = "Value"
        sheet.Range("K2").Value2 = "Units.Force.Input": sheet.Range("L2").Value2 = IIf(mode = 0, "N", "kN")
        sheet.Range("K3").Value2 = "Units.Moment.Input": sheet.Range("L3").Value2 = IIf(mode = 0, "N*mm", "kN*m")
        settings.LoadFromRange sheet.Range("K1:M3")
        Set units = New CUnitSystem
        units.LoadFromSettings settings
        sheet.Range("A2").Value2 = "TINY_N": sheet.Range("E2").Value2 = "PR1"
        sheet.Range("A3").Value2 = "TINY_M": sheet.Range("E3").Value2 = "PR1"
        sheet.Range("B2").Value2 = 0.00000000001 / (1000# ^ mode)
        sheet.Range("C3").Value2 = -0.00000000001 / (1000000# ^ mode)
        Set batch = New CBatchSectionCalculator
        reader.LoadFromRange sheet.Range("A1:G3"), batch, units
        TestAssertTrue stats, "audit03.reader.tiny." & CStr(mode) & ".count", batch.Count = 2
        If batch.Count = 2 Then
            TestAssertClose stats, "audit03.reader.tiny." & CStr(mode) & ".n", batch.N(1), 0.00000000001, 1E-26
            TestAssertClose stats, "audit03.reader.tiny." & CStr(mode) & ".m", batch.UserMx(2), -0.00000000001, 1E-26
        End If
    Next mode
    sheet.Range("B2").Value2 = 1E+308
    sheet.Range("A3").Value2 = "AFTER_OVERFLOW": sheet.Range("B3").Value2 = 1#
    Set batch = TestBuildUiBatch()
    reader.LoadFromRange sheet.Range("A1:G3"), batch, units
    TestAssertTrue stats, "audit03.reader.overflow.count", batch.Count = 2
    If batch.Count = 2 Then
        batch.Execute
        TestAssertTextEquals stats, "audit03.reader.overflow.status", batch.ResultAt(1).Status, "InputErr"
        TestAssertTrue stats, "audit03.reader.overflow.comment", Len(batch.ResultAt(1).OverallMeta.ResultComment) > 0
        TestAssertTextEquals stats, "audit03.reader.overflow.next", batch.CombinationID(2), "AFTER_OVERFLOW"
    End If
    Exit Sub
Failed:
    TestAssertTrue stats, "audit03.reader.tiny.runtime." & CStr(Err.Number) & "." & Err.Description, False
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' ========================== ДЛЯ ТЕСТОВ ==========================
' Проверяет неверный ввод через реальные Range и публичные reader/build API.
' Метод не изменяет пользовательские Config-ячейки и возвращает отдельный отчет.
Public Function RunAudit03InputContractTests() As String
    Dim stats As TUiTestStats
    TestAudit03InputContracts stats
    TestAppendLine stats, "TOTAL_AUDIT03_INPUT_CONTRACTS: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunAudit03InputContractTests = stats.Report
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Различает недопустимую структуру, значение и размер сетки. Все ветви должны
' давать адресную ошибку ввода, а не runtime 6/9/13/91 или скрытый default.
Private Sub TestAudit03InputContracts(ByRef stats As TUiTestStats)
    Dim sheet As Object, oldAlerts As Boolean
    oldAlerts = Application.DisplayAlerts
    On Error GoTo Failed
    Set sheet = ThisWorkbook.Worksheets.Add
    sheet.Name = "__Audit03Input"
    sheet.Range("A1").Value2 = "Key": sheet.Range("B1").Value2 = "Value": sheet.Range("C1").Value2 = "Unit"
    sheet.Range("A2").Value2 = "Audit03.Test"
    Dim settings As CSystemSettingsReader, value As Variant
    Set settings = New CSystemSettingsReader
    For Each value In Array("Yes", "True", "1", "Да", "No", "False", "0", "Нет")
        sheet.Range("B2").Value2 = CStr(value)
        settings.LoadFromRange sheet.Range("A1:C2")
        TestAssertTrue stats, "audit03.input.boolean.valid." & CStr(value), _
            settings.GetBoolean("Audit03.Test") = (value = "Yes" Or value = "True" Or value = "1" Or value = "Да")
    Next value
    TestAssertTrue stats, "audit03.input.boolean.missingTrue", settings.GetBoolean("Audit03.Absent", True)
    TestAssertTrue stats, "audit03.input.boolean.missingFalse", Not settings.GetBoolean("Audit03.Absent", False)
    For Each value In Array("Maybe", "TODO", "")
        sheet.Range("B2").Value2 = CStr(value)
        TestAudit03AssertInputError stats, sheet.Range("A1:C2"), "Boolean", _
            "audit03.input.boolean.invalid." & CStr(value), vbObjectError + 4314
    Next value
    For Each value In Array(-2147483648#, 0#, 17#, 2147483647#)
        sheet.Range("B2").Value2 = CDbl(value)
        settings.LoadFromRange sheet.Range("A1:C2")
        TestAssertClose stats, "audit03.input.long.valid." & CStr(value), CDbl(settings.GetLong("Audit03.Test")), CDbl(value), 0#
    Next value
    For Each value In Array(1.25, -1.25, 1E+20)
        sheet.Range("B2").Value2 = CDbl(value)
        TestAudit03AssertInputError stats, sheet.Range("A1:C2"), "Long", _
            "audit03.input.long.invalid." & CStr(value), vbObjectError + 4315
    Next value
    sheet.Range("A2").Value2 = "Solver.MaxIterations": sheet.Range("B2").Value2 = 1.25
    TestAudit03AssertInputError stats, sheet.Range("A1:C2"), "Solver", "audit03.input.consumer.integer", vbObjectError + 4315
    sheet.Range("A2").Value2 = "Mesh.BoundarySubdivisions": sheet.Range("B2").Value2 = 0
    TestAudit03AssertInputError stats, sheet.Range("A1:C2"), "MeshRegistry", "audit03.input.consumer.boundary", vbObjectError + 4139
    For Each value In Array("", "TODO")
        sheet.Range("B2").Value2 = CStr(value)
        TestAudit03AssertInputError stats, sheet.Range("A1:C2"), "MeshRegistry", _
            "audit03.input.consumer.boundaryMissing." & CStr(value), vbObjectError + 4139
    Next value
    TestAudit03AssertInputError stats, sheet.Range("A1"), "Settings", "audit03.input.settings.scalar", vbObjectError + 4316
    TestAudit03AssertInputError stats, sheet.Range("A1:B2"), "Settings", "audit03.input.settings.narrow", vbObjectError + 4316
    TestAudit03AssertInputError stats, Nothing, "Settings", "audit03.input.settings.nothing", vbObjectError + 4316
    TestAudit03AssertInputError stats, sheet.Range("A1"), "Profiles", "audit03.input.profiles.scalar", vbObjectError + 3987
    TestAudit03AssertInputError stats, Nothing, "MeshNothing", "audit03.input.mesh.nothing", vbObjectError + 2107
    TestAudit03AssertInputError stats, Nothing, "MeshAxis", "audit03.input.mesh.axis", vbObjectError + 2108
    TestAudit03AssertInputError stats, Nothing, "MeshProduct", "audit03.input.mesh.product", vbObjectError + 2108
    TestAudit03AssertInputError stats, Nothing, "MeshBoundary", "audit03.input.mesh.boundary", vbObjectError + 2109
    TestAudit03AssertInputError stats, Nothing, "MeshSubcellProduct", "audit03.input.mesh.subcellProduct", vbObjectError + 2108
    GoTo CleanUp
Failed:
    TestAssertTrue stats, "audit03.input.runtime." & CStr(Err.Number) & "." & Err.Description, False
CleanUp:
    On Error Resume Next
    Application.DisplayAlerts = False
    If Not sheet Is Nothing Then sheet.Delete
    Application.DisplayAlerts = oldAlerts
    On Error GoTo 0
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет точный номер и полный текст ошибки публичного метода.
' Перехватывает только ожидаемый ошибочный ввод текущего теста.
Private Sub TestAudit03AssertInputError(ByRef stats As TUiTestStats, ByVal source As Object, _
        ByVal operation As String, ByVal prefix As String, ByVal expectedCode As Long)
    Dim description As String, actualCode As Long
    actualCode = TestAudit03CaptureInputError(source, operation, description)
    TestAssertTrue stats, prefix & ".typedError", actualCode = expectedCode
    If actualCode <> 0 Then TestAssertTrue stats, prefix & ".reason", Len(description) > 0
    TestAppendLine stats, "INPUT: " & prefix & "; error=" & CStr(actualCode) & "; expected=" & CStr(expectedCode) & "; reason=" & description
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Запускает отдельную неверную операцию без маскировки ошибок следующего теста.
' Сетка с непредставимым числом ячеек должна отказать до выделения массива/цикла.
Private Function TestAudit03CaptureInputError(ByVal source As Object, ByVal operation As String, _
        ByRef description As String) As Long
    On Error GoTo ExpectedError
    Dim settings As CSystemSettingsReader, profiles As CCalculationProfileCatalog
    Dim mesh As CFiberMeshBuilder, geometry As CGeometryRegion
    Dim solver As CSectionSolver, registry As CSectionTypeRegistry
    Dim capacity As CCapacitySolver, formation As CCrackFormationCalculator
    Dim flag As Boolean, integerValue As Long
    Select Case operation
        Case "Settings", "Boolean", "Long", "Solver", "MeshRegistry", "Capacity", "Formation"
            Set settings = New CSystemSettingsReader
            settings.LoadFromRange source
            If operation = "Boolean" Then flag = settings.GetBoolean("Audit03.Test", True)
            If operation = "Long" Then integerValue = settings.GetLong("Audit03.Test")
            If operation = "Solver" Then
                Set solver = New CSectionSolver
                solver.ApplySettings settings
            End If
            If operation = "Capacity" Then
                Set capacity = New CCapacitySolver
                capacity.ApplySettings settings
            End If
            If operation = "Formation" Then
                Set formation = New CCrackFormationCalculator
                formation.ApplySettings settings
            End If
            If operation = "MeshRegistry" Then
                integerValue = ReadMeshBoundarySubdivisions(settings)
            End If
        Case "Profiles"
            Set profiles = New CCalculationProfileCatalog
            profiles.LoadFromRange source
        Case Else
            Set mesh = New CFiberMeshBuilder
            Set geometry = TestFixtureBox(100#, 100#)
            Select Case operation
                Case "MeshNothing": mesh.BuildMesh Nothing, 25#, 25#
                Case "MeshAxis": mesh.BuildMesh geometry, 1E-308, 25#
                Case "MeshProduct": mesh.BuildMesh geometry, 0.001, 0.001
                Case "MeshBoundary": mesh.BuildMesh geometry, 25#, 25#, 1, 0
                Case "MeshSubcellProduct": mesh.BuildMesh geometry, 25#, 25#, 1, 2147483647
                Case Else: Err.Raise vbObjectError + 4499, "Audit03CaptureInputError", "Неизвестный тестовый сценарий ввода."
            End Select
    End Select
    Exit Function
ExpectedError:
    TestAudit03CaptureInputError = Err.Number
    description = Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' проверяет настоящий reader профилей независимо от общего
' SettingsReader. Неверный переключатель не должен молча отключать расчет.
Public Function RunAudit03ProfileInputTests() As String
    On Error GoTo Failed
    Dim stats As TUiTestStats
    TestAudit03ProfileInputContracts stats
    TestAppendLine stats, "TOTAL_AUDIT03_PROFILE_INPUT: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunAudit03ProfileInputTests = stats.Report
    Exit Function
Failed:
    RunAudit03ProfileInputTests = stats.Report & "RUNTIME ERROR: " & CStr(Err.Number) & "; " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' копирует фактическую таблицу профилей на временный лист.
' Проверяет все Boolean aliases и поврежденные/отсутствующие ячейки каждого
' переключателя; исходные Config, формулы и validation не изменяются.
Private Sub TestAudit03ProfileInputContracts(ByRef stats As TUiTestStats)
    Dim source As Object, sheet As Object, target As Object, data As Variant
    Dim oldAlerts As Boolean, key As Variant, value As Variant, row As Long, column As Long
    Dim keyRow As Long, profileColumn As Long, profiles As CCalculationProfileCatalog
    Dim profile As CCalculationProfile, actualCode As Long, description As String, caseIndex As Long
    oldAlerts = Application.DisplayAlerts
    On Error GoTo Failed
    Set source = ThisWorkbook.Names.Item("rngCalculationProfiles").RefersToRange
    data = source.Value2
    Set sheet = ThisWorkbook.Worksheets.Add
    sheet.Name = "__Audit03ProfileInput"
    Set target = sheet.Range("A1").Resize(source.Rows.Count, source.Columns.Count)
    target.Value2 = data
    For row = 1 To UBound(data, 1)
        For column = 3 To UBound(data, 2)
            If CStr(data(row, column)) = "PR1" Then profileColumn = column
        Next column
        If profileColumn > 0 Then Exit For
    Next row
    If profileColumn = 0 Then Err.Raise vbObjectError + 4499, "TestAudit03ProfileInputContracts", "В fixture не найден PR1."
    For Each key In Array("Calculation.Strength.DirectState", "Calculation.Strength.Capacity", _
            "Calculation.Crack.Width", "Calculation.Stability.Enabled")
        keyRow = 0
        For row = 1 To UBound(data, 1)
            If CStr(data(row, 2)) = CStr(key) Then keyRow = row
        Next row
        If keyRow = 0 Then Err.Raise vbObjectError + 4499, "TestAudit03ProfileInputContracts", "В fixture не найден " & CStr(key)
        For Each value In Array("Yes", "True", "1", "Да", "No", "False", "0", "Нет")
            target.Value2 = data
            target.Cells.Item(keyRow, profileColumn).Value2 = CStr(value)
            Set profiles = New CCalculationProfileCatalog
            profiles.LoadFromRange target
            Set profile = profiles.ProfileById("PR1")
            TestAssertTrue stats, "audit03.profile.boolean." & CStr(key) & "." & CStr(value), _
                TestAudit03ProfileFlag(profile, CStr(key)) = (value = "Yes" Or value = "True" Or value = "1" Or value = "Да")
        Next value
        caseIndex = 0
        For Each value In Array("Maybe", "TODO", "", CVErr(xlErrNA))
            caseIndex = caseIndex + 1
            target.Value2 = data
            target.Cells.Item(keyRow, profileColumn).Value2 = value
            description = vbNullString
            actualCode = TestAudit03CaptureInputError(target, "Profiles", description)
            TestAssertTrue stats, "audit03.profile.invalid." & CStr(key) & "." & CStr(caseIndex), actualCode = vbObjectError + 3988
            TestAssertTrue stats, "audit03.profile.reason." & CStr(key) & "." & CStr(caseIndex), _
                InStr(1, description, CStr(key), vbBinaryCompare) > 0 And InStr(1, description, "PR1", vbBinaryCompare) > 0
            TestAppendLine stats, "PROFILE_INPUT: key=" & CStr(key) & "; case=" & CStr(caseIndex) & _
                "; error=" & CStr(actualCode) & "; reason=" & description
        Next value
        target.Value2 = data
        target.Cells.Item(keyRow, 2).Value2 = "__Removed." & CStr(key)
        description = vbNullString
        actualCode = TestAudit03CaptureInputError(target, "Profiles", description)
        TestAssertTrue stats, "audit03.profile.missing." & CStr(key), actualCode = vbObjectError + 3988
        TestAssertTrue stats, "audit03.profile.missingReason." & CStr(key), InStr(1, description, CStr(key), vbBinaryCompare) > 0
    Next key
    GoTo CleanUp
Failed:
    stats.Failed = stats.Failed + 1
    TestAppendLine stats, "FAIL: audit03.profile.runtime; " & CStr(Err.Number) & "; " & Err.Description
CleanUp:
    On Error Resume Next
    Application.DisplayAlerts = False
    If Not sheet Is Nothing Then sheet.Delete
    Application.DisplayAlerts = oldAlerts
    On Error GoTo 0
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' читает конечный флаг профиля, а не повторяет строковый parser.
' Поэтому тест выявляет потерю настройки между ячейкой и CCalculationProfile.
Private Function TestAudit03ProfileFlag(ByVal profile As CCalculationProfile, ByVal key As String) As Boolean
    Select Case key
        Case "Calculation.Strength.DirectState": TestAudit03ProfileFlag = profile.StrengthDirectStateEnabled
        Case "Calculation.Strength.Capacity": TestAudit03ProfileFlag = profile.StrengthCapacityEnabled
        Case "Calculation.Crack.Width": TestAudit03ProfileFlag = profile.CrackWidthEnabled
        Case "Calculation.Stability.Enabled": TestAudit03ProfileFlag = profile.StabilityEnabled
        Case Else: Err.Raise vbObjectError + 4499, "Audit03ProfileFlag", "Неизвестный тестовый переключатель профиля."
    End Select
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' отдельный entrypoint для одинакового контрпримера до и после
' исправления. Настройки передаются через настоящий Range и ApplySettings.
Public Function RunAudit03NumericSettingsInputTests() As String
    Dim stats As TUiTestStats
    TestAudit03NumericSettingsInputContracts stats
    TestAppendLine stats, "TOTAL_AUDIT03_NUMERIC_SETTINGS_INPUT: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunAudit03NumericSettingsInputTests = stats.Report
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' явно пустая численная настройка или TODO не заменяется значением
' по умолчанию. Отсутствующий необязательный ключ и пустой диаметр
' имеют собственные правила чтения и проверяются отдельно.
Private Sub TestAudit03NumericSettingsInputContracts(ByRef stats As TUiTestStats)
    Dim sheet As Object, oldAlerts As Boolean, operationValue As Variant, keyValue As Variant
    Dim key As String, operation As String, value As Variant, actualCode As Long
    Dim description As String, expectedCode As Long, settings As CSystemSettingsReader, inputRange As Object
    oldAlerts = Application.DisplayAlerts
    On Error GoTo Failed
    Set sheet = ThisWorkbook.Worksheets.Add
    sheet.Name = "__Audit03NumericInput"
    sheet.Range("A1").Value2 = "Key": sheet.Range("B1").Value2 = "Value": sheet.Range("C1").Value2 = "Unit"
    For Each operationValue In Array("Solver", "Capacity", "Formation")
        operation = CStr(operationValue)
        Set inputRange = sheet.Range("A1:C2")
        If operation = "Formation" Then
            ' Минимальный Config Formation содержит обязательные селекторы;
            ' отсутствие проверяемого Solver-key остается отдельным сценарием.
            sheet.Range("A3").Value2 = "SLS.Crack.InitiationSolutionStrategy": sheet.Range("B3").Value2 = "Auto"
            sheet.Range("A4").Value2 = "SLS.Crack.ToleranceLambda": sheet.Range("B4").Value2 = 0.01
            Set inputRange = sheet.Range("A1:C4")
        End If
        For Each keyValue In TestAudit03NumericSettingsKeys(operation)
            key = CStr(keyValue)
            sheet.Range("A2").Value2 = key
            For Each value In Array("", "TODO", "abc", CVErr(xlErrNA))
                sheet.Range("B2").Value2 = value
                If VarType(value) = vbError Then
                    expectedCode = vbObjectError + 4309
                ElseIf Len(CStr(value)) = 0 Or CStr(value) = "TODO" Then
                    expectedCode = vbObjectError + 4310
                Else
                    expectedCode = vbObjectError + 4312
                End If
                description = vbNullString
                actualCode = TestAudit03CaptureInputError(inputRange, operation, description)
                TestAssertTrue stats, "audit03.numeric." & operation & "." & key & ".error", actualCode = expectedCode
                TestAssertTrue stats, "audit03.numeric." & operation & "." & key & ".reason", InStr(1, description, key, vbBinaryCompare) > 0
                TestAppendLine stats, "NUMERIC_INPUT: operation=" & operation & "|key=" & key & _
                    "|value=" & CStr(value) & "|error=" & CStr(actualCode) & "|reason=" & description
            Next value
            sheet.Range("B2").Value2 = 80#
            actualCode = TestAudit03CaptureInputError(inputRange, operation, description)
            TestAssertTrue stats, "audit03.numeric." & operation & "." & key & ".parsable", actualCode = 0
            sheet.Range("A2").Value2 = "Audit03.Unrelated"
            actualCode = TestAudit03CaptureInputError(inputRange, operation, description)
            TestAssertTrue stats, "audit03.numeric." & operation & ".optionalAbsent", actualCode = 0
        Next keyValue
    Next operationValue
    sheet.Range("A2").Value2 = "Rebar.Diameter2": sheet.Range("B2").Value2 = vbNullString
    Set settings = New CSystemSettingsReader
    settings.LoadFromRange sheet.Range("A1:C2")
    TestAssertClose stats, "audit03.numeric.optionalDiameterBlank", settings.GetDouble("Rebar.Diameter2"), 0#, 0#
    TestAssertClose stats, "audit03.numeric.optionalKeyDefault", settings.GetDouble("Audit03.Absent", 17#), 17#, 0#
    TestAudit03BatchNumericSettingsInput stats
    GoTo CleanUp
Failed:
    stats.Failed = stats.Failed + 1
    TestAppendLine stats, "FAIL: audit03.numeric.runtime; " & CStr(Err.Number) & "; " & Err.Description
CleanUp:
    On Error Resume Next
    Application.DisplayAlerts = False
    If Not sheet Is Nothing Then sheet.Delete
    Application.DisplayAlerts = oldAlerts
    On Error GoTo 0
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' проверяет сохранение отрицательного знака после реального
' преобразования tf/tf*m/кривизны. Capacity, Formation и Batch проходят свои
' API до solver-а; ошибка не доказывается только чтением settings/getter-а.
Public Function RunAudit03UnitSignConsumerTests() As String
    Dim stats As TUiTestStats
    TestAppendLine stats, "RUN: Audit03 unit-sign consumers"
    TestAudit03UnitSignConsumers stats
    TestAppendLine stats, "TOTAL_AUDIT03_UNIT_SIGN_CONSUMERS: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunAudit03UnitSignConsumerTests = stats.Report
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' меняет четыре поля на контролируемой копии Config и гарантированно
' восстанавливает формулы. Внутренний знак сжатия не делает отрицательный
' допуск положительным; численные итерации с такой конфигурацией недопустимы.
Private Sub TestAudit03UnitSignConsumers(ByRef stats As TUiTestStats)
    Dim target As Object, savedFormula As Variant, settings As CSystemSettingsReader
    Dim unitRange As Object, signRange As Object, savedUnits As Variant, savedSigns As Variant
    Dim key As Variant, row As Long, valueCell As Object, units As CUnitSystem
    Dim section As CSectionModel, materials As CMaterialModelProvider, i As Long
    Dim policy As CResultStatusPolicy
    On Error GoTo Failed
    Set target = ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange
    savedFormula = target.Formula
    Set unitRange = ThisWorkbook.Names.Item("rngUnitSettings").RefersToRange
    Set signRange = ThisWorkbook.Names.Item("rngSignConventionSettings").RefersToRange
    savedUnits = unitRange.Formula: savedSigns = signRange.Formula
    TestSetSystemSetting "Units.Force.Input", "tf"
    TestSetSystemSetting "Units.Moment.Input", "tf*m"
    TestSetSystemSetting "Units.Curvature.Input", "1/m"
    TestSetSystemSetting "Sign.N.User", "Compression"
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem
    units.LoadFromSettings settings
    Set policy = New CResultStatusPolicy
    Set materials = New CMaterialModelProvider
    materials.Initialize settings
    Set section = New CSectionModel
    For i = 0 To 3
        section.AddConcreteElement -50# + 100# * (i Mod 2), -50# + 100# * (i \ 2), 10000#
        section.AddRebarElement -40# + 80# * (i Mod 2), -40# + 80# * (i \ 2), 20#, 314.159265358979, "A400"
    Next i
    Dim strengthSpec As CMaterialModelSpec, currentSpec As CMaterialModelSpec, formationSpec As CMaterialModelSpec
    Set strengthSpec = New CMaterialModelSpec
    strengthSpec.Initialize "ULS(I)", "ThreeLine", "Ignore", "TwoLine"
    Set currentSpec = New CMaterialModelSpec
    currentSpec.Initialize "SLS(II)", "TwoLine", "Ignore", "TwoLine"
    Set formationSpec = New CMaterialModelSpec
    formationSpec.Initialize "SLS(II)", "ThreeLine", "UseDiagram", "TwoLine"
    Dim concrete As Object, steel As Object, load As CSectionLoadState
    Set concrete = materials.ConcreteMaterialFromSpec(strengthSpec)
    Set steel = materials.SteelMaterialFromSpec(strengthSpec)
    Set load = New CSectionLoadState
    load.Initialize -100000#, 50000000#, 10000000#, 0#, 0#
    For Each key In Array("Solver.ToleranceN", "Solver.ToleranceMx", "Solver.ToleranceMy")
        target.Formula = savedFormula
        Set valueCell = Nothing
        For row = 2 To target.Rows.Count
            If CStr(target.Cells.Item(row, 1).Value2) = CStr(key) Then
                Set valueCell = target.Cells.Item(row, 2)
                Exit For
            End If
        Next row
        If valueCell Is Nothing Then Err.Raise vbObjectError + 4499, "TestAudit03UnitSignConsumers", "Не найден ключ " & CStr(key)
        valueCell.Value2 = -1#
        Set settings = New CSystemSettingsReader
        settings.LoadFromWorkbook ThisWorkbook
        Dim cap As CCapacitySolver, formation As CCrackFormationCalculator, formed As CCrackFormationResult
        Dim batch As CBatchSectionCalculator, result As CCombinationResult, prefix As String
        prefix = "audit03.unitSign." & CStr(key)
        TestAppendLine stats, "RUN: " & prefix & ".capacity"
        Set cap = New CCapacitySolver
        cap.ApplySettings settings, units
        cap.SolveByLoadPathMultiplier section, concrete, steel, 0#, load.N, 0#, load.InternalMx, 0#, load.InternalMy
        TestAssertTrue stats, prefix & ".capacityTyped", cap.FailureCode = sfcInvalidConfiguration
        TestAssertTrue stats, prefix & ".capacityHasDiagnosticSolver", Not cap.LastSolver Is Nothing
        If Not cap.LastSolver Is Nothing Then
            TestAssertTrue stats, prefix & ".capacityNoIterations", cap.LastSolver.Iterations = 0
        End If
        TestAssertTrue stats, prefix & ".capacityNoRetries", cap.TestRetryCount = 0
        TestAssertTrue stats, prefix & ".capacityReason", Len(cap.StopReason) > 0
        TestAppendLine stats, "RUN: " & prefix & ".formation"
        Set formation = New CCrackFormationCalculator
        formation.ApplySettings settings, units
        formation.CrackFormationPath = "LambdaNMxy"
        formation.CrackFormationSolutionStrategy = "LoadMultiplier"
        Set formed = formation.CheckFormation(section, materials, currentSpec, formationSpec, _
            load.N, load.InternalMx, load.InternalMy, load, 0#, 0#)
        TestAssertTrue stats, prefix & ".formationTyped", formed.ResultMeta.InternalStatus = rsInvalidConfiguration
        TestAssertTrue stats, prefix & ".formationNoPoint", Not formed.HasLimitPoint
        TestAssertTrue stats, prefix & ".formationReason", Len(formed.ResultMeta.ResultComment) > 0
        TestAppendLine stats, "RUN: " & prefix & ".batch"
        Set batch = TestBuildUiBatch()
        TestAppendLine stats, "RUN: " & prefix & ".batch.built; " & TestAudit02ExcelMemory()
        batch.ApplySettings settings, units
        TestAppendLine stats, "RUN: " & prefix & ".batch.settingsApplied; " & TestAudit02ExcelMemory()
        batch.AddCombination "UNIT_SIGN", load.N, load.InternalMx, load.InternalMy, "PR1", vbNullString
        TestAppendLine stats, "RUN: " & prefix & ".batch.execute; " & TestAudit02ExcelMemory()
        batch.Execute
        TestAppendLine stats, "RUN: " & prefix & ".batch.executed; " & TestAudit02ExcelMemory()
        Set result = batch.ResultAt(1)
        TestAssertTrue stats, prefix & ".batchTyped", result.DirectStateMeta.InternalStatus = rsInvalidConfiguration
        TestAssertTextEquals stats, prefix & ".batchExternal", policy.ExternalStatus(result.DirectStateMeta), "InputErr"
        TestAssertTrue stats, prefix & ".batchReason", Len(result.DirectStateMeta.ResultComment) > 0
        TestAppendLine stats, "UNIT_SIGN_CONSUMER: key=" & CStr(key) & "|capacity=" & cap.StopReason & _
            "|formation=" & formed.ResultMeta.ResultComment & "|batch=" & result.DirectStateMeta.ResultComment
    Next key
    GoTo CleanUp
Failed:
    stats.Failed = stats.Failed + 1
    TestAppendLine stats, "FAIL: audit03.unitSign.runtime; " & CStr(Err.Number) & "; " & Err.Description
CleanUp:
    On Error Resume Next
    If Not target Is Nothing Then target.Formula = savedFormula
    If Not unitRange Is Nothing Then unitRange.Formula = savedUnits
    If Not signRange Is Nothing Then signRange.Formula = savedSigns
    On Error GoTo 0
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' полный список численных ключей каждого ApplySettings.
' Общие Solver-поля и Capacity overrides проверяются по реальным маршрутам,
' а не только прямым вызовом GetDouble. Инженерные диапазоны проверяются отдельно.
Private Function TestAudit03NumericSettingsKeys(ByVal operation As String) As Variant
    If operation = "Capacity" Then
        TestAudit03NumericSettingsKeys = Array("Capacity.InitialLambda", "Capacity.MaxLambda", "Capacity.ToleranceLambda", _
        "Capacity.ToleranceStrain", "Capacity.MaxRetries", "Capacity.BaseLoadSteps", _
        "Capacity.SolverMaxIterations", "Solver.ToleranceN", "Solver.ToleranceMx", _
        "Solver.ToleranceMy", "Solver.SecantMaxRestarts")
    ElseIf operation = "Batch" Then
        TestAudit03NumericSettingsKeys = Array("Solver.MaxIterations", "Solver.LoadSteps", "Solver.ToleranceN", _
        "Solver.ToleranceMx", "Solver.ToleranceMy", "Solver.SecantMaxRestarts", _
        "Capacity.InitialLambda", "Capacity.MaxLambda", "Capacity.ToleranceLambda", _
        "Capacity.ToleranceStrain", "Capacity.MaxRetries", "Capacity.BaseLoadSteps", _
        "Capacity.SolverMaxIterations")
    Else
        TestAudit03NumericSettingsKeys = Array("Solver.MaxIterations", "Solver.LoadSteps", "Solver.ToleranceN", _
        "Solver.ToleranceMx", "Solver.ToleranceMy", "Solver.SecantMaxRestarts")
    End If
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Batch сохраняет ошибку ApplySettings, а не выбрасывает ее наружу.
' Используем полный Config с обязательными параметрами устойчивости; проверяем
' InputErr, отсутствие solve и возможность следующего запуска после восстановления.
Private Sub TestAudit03BatchNumericSettingsInput(ByRef stats As TUiTestStats)
    Dim target As Object, savedFormula As Variant, batch As CBatchSectionCalculator
    Dim settings As CSystemSettingsReader, key As Variant, value As Variant, row As Long
    Dim valueCell As Object
    On Error GoTo Failed
    Set target = ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange
    savedFormula = target.Formula
    Set batch = TestBuildUiBatch()
    batch.AddCombination "NUMERIC_INPUT", 0#, 0#, 0#, "PR1", "Проверка ввода численных настроек"
    For Each key In TestAudit03NumericSettingsKeys("Batch")
        Set valueCell = Nothing
        For row = 2 To target.Rows.Count
            If CStr(target.Cells.Item(row, 1).Value2) = CStr(key) Then
                Set valueCell = target.Cells.Item(row, 2)
                Exit For
            End If
        Next row
        If valueCell Is Nothing Then Err.Raise vbObjectError + 4499, _
            "TestAudit03BatchNumericSettingsInput", "Не найден тестируемый ключ " & CStr(key)
        For Each value In Array("", "TODO", "abc", CVErr(xlErrNA))
            target.Formula = savedFormula
            valueCell.Value2 = value
            TestAudit03CheckBatchNumericInput stats, batch, CStr(key), value
        Next value
    Next key
    target.Formula = savedFormula
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook
    batch.ApplySettings settings
    batch.Execute
    TestAssertTrue stats, "audit03.numeric.Batch.restoredStatus", batch.ResultAt(1).Status <> "InputErr"
    TestAssertTrue stats, "audit03.numeric.Batch.restoredSolve", batch.ResultAt(1).DirectStateMeta.Calculated
    GoTo CleanUp
Failed:
    stats.Failed = stats.Failed + 1
    TestAppendLine stats, "FAIL: audit03.numeric.Batch.runtime; " & CStr(Err.Number) & "; " & Err.Description
CleanUp:
    On Error Resume Next
    If Not target Is Nothing Then target.Formula = savedFormula
    On Error GoTo 0
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' различает ошибку ячейки при чтении и сохранённый отказ расчёта
' сочетаний. Неверный численный ввод не должен запускать поиск равновесия.
Private Sub TestAudit03CheckBatchNumericInput(ByRef stats As TUiTestStats, _
        ByVal batch As CBatchSectionCalculator, ByVal key As String, ByVal value As Variant)
    Dim settings As CSystemSettingsReader, result As CCombinationResult, prefix As String
    Dim errorNumber As Long, description As String
    prefix = "audit03.numeric.Batch." & key & "." & CStr(value)
    On Error GoTo ReaderFailed
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook
    On Error GoTo Failed
    batch.ApplySettings settings
    batch.Execute
    Set result = batch.ResultAt(1)
    TestAssertTextEquals stats, prefix & ".status", result.Status, "InputErr"
    TestAssertTrue stats, prefix & ".typed", result.OverallMeta.InternalStatus = rsInvalidInput
    TestAssertTrue stats, prefix & ".reason", InStr(1, result.OverallMeta.ResultComment, key, vbBinaryCompare) > 0
    TestAssertTrue stats, prefix & ".noDirectSolve", Not result.DirectStateMeta.Calculated
    TestAssertTrue stats, prefix & ".noCapacitySearch", Not result.CapacityMeta.Calculated
    TestAppendLine stats, "NUMERIC_INPUT: operation=Batch|key=" & key & "|value=" & CStr(value) & _
        "|status=" & result.Status & "|reason=" & result.OverallMeta.ResultComment
    Exit Sub
ReaderFailed:
    errorNumber = Err.Number: description = Err.Description
    TestAssertTrue stats, prefix & ".readerError", VarType(value) = vbError And errorNumber = vbObjectError + 4309
    TestAssertTrue stats, prefix & ".readerReason", InStr(1, description, key, vbBinaryCompare) > 0
    TestAppendLine stats, "NUMERIC_INPUT: operation=BatchReader|key=" & key & "|value=" & CStr(value) & _
        "|error=" & CStr(errorNumber) & "|reason=" & description
    Exit Sub
Failed:
    stats.Failed = stats.Failed + 1
    TestAppendLine stats, "FAIL: " & prefix & ".runtime; " & CStr(Err.Number) & "; " & Err.Description
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' адресный gate независимых настроек дополнительных рядов Rectangles.
' Проверяет чтение, списки, влияние на одну грань и ошибки активного ввода.
Public Function RunRectanglesIndependentSelectorTests() As String
    Dim stats As TUiTestStats
    TestRectanglesIndependentSelectors stats
    TestRectanglesIndependentSelectorLayout stats
    TestRectanglesIndependentSelectorEffects stats
    TestRectanglesIndependentThirdRows stats
    TestAppendLine stats, "TOTAL_RECTANGLES_INDEPENDENT: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunRectanglesIndependentSelectorTests = stats.Report
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' разные значения противоположных граней сохраняются в своих
' ячейках снимка. Чтение первой строки не подменяет значение второй строки.
Private Sub TestRectanglesIndependentSelectors(ByRef stats As TUiTestStats)
    Dim values As Variant, reader As CSectionInputReader, side As Variant, field As Variant
    Dim firstValue As String, secondValue As String, pair As Variant
    On Error GoTo Failed
    values = TestFixtureRecipeValues("rngRectanglesGeometry")
    For Each pair In Array(Array("R1.Left", "R1.Right"), Array("R1.Top", "R1.Bottom"), _
            Array("R2.Left", "R2.Right"), Array("R2.Top", "R2.Bottom"))
        For Each field In Array("pos2", "bind2", "pos3", "bind3")
            firstValue = "EachBar": secondValue = "EverySecondBar"
            If Left$(CStr(field), 3) = "pos" Then firstValue = "Stacked": secondValue = "SideBySide"
            TestFixtureSetInput values, "Extra", CStr(pair(0)), CStr(field), firstValue
            TestFixtureSetInput values, "Extra", CStr(pair(1)), CStr(field), secondValue
            Set reader = TestFixtureInput(values)
            TestAssertTextEquals stats, CStr(pair(0)) & "." & CStr(field), reader.Text("Extra", CStr(pair(0)), CStr(field)), firstValue
            TestAssertTextEquals stats, CStr(pair(1)) & "." & CStr(field), reader.Text("Extra", CStr(pair(1)), CStr(field)), secondValue
            For Each side In pair
                TestFixtureSetInput values, "Extra", CStr(side), CStr(field), Empty
                Set reader = TestFixtureInput(values)
                TestAssertTextEquals stats, CStr(side) & "." & CStr(field) & ".empty", reader.Text("Extra", CStr(side), CStr(field), ""), ""
            Next side
        Next field
    Next pair
    Exit Sub
Failed:
    stats.Failed = stats.Failed + 1
    TestAppendLine stats, "FAIL: rectangles.independent.reader; " & CStr(Err.Number) & "; " & Err.Description
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' все 32 селектора фактического Config не объединены, имеют
' вертикальный список выбора и читаются по адресу своей физической грани.
Private Sub TestRectanglesIndependentSelectorLayout(ByRef stats As TUiTestStats)
    Dim target As Object, cell As Object, reader As CSectionInputReader
    Dim side As Variant, field As Variant, prefix As String, config As Object, saved As Variant
    On Error GoTo Failed
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = TestFixtureSaveInputs(config)
    modTestShapeArchitecture.TestConfigureShapeFixture "TwoRectangles"
    Set target = ThisWorkbook.Names("rngRectanglesGeometry").RefersToRange
    For Each side In Array("R1.Top", "R1.Bottom", "R1.Left", "R1.Right", "R2.Top", "R2.Bottom", "R2.Left", "R2.Right")
        For Each field In Array("pos2", "bind2", "pos3", "bind3")
            Set cell = TestFixtureInputCell(target, "Extra", CStr(side), CStr(field))
            prefix = "rectangles.independent.layout." & CStr(side) & "." & CStr(field)
            TestAssertTrue stats, prefix & ".unmerged", Not cell.MergeCells
            TestAssertTrue stats, prefix & ".dropdown", cell.Validation.Type = 3 And cell.Validation.InCellDropdown
            TestAssertTrue stats, prefix & ".inlineList", InStr(1, cell.Validation.Formula1, Application.International(xlListSeparator), vbBinaryCompare) > 0
            TestAssertTrue stats, prefix & ".centered", cell.HorizontalAlignment = xlCenter And cell.VerticalAlignment = xlCenter
            Set reader = TestFixtureInput(target.Value2)
            TestAssertTextEquals stats, prefix & ".reader", reader.Text("Extra", CStr(side), CStr(field)), CStr(cell.Value2)
        Next field
    Next side
    GoTo Finished
Failed:
    stats.Failed = stats.Failed + 1
    TestAppendLine stats, "FAIL: rectangles.independent.layout; " & CStr(Err.Number) & "; " & Err.Description
Finished:
    If Not config Is Nothing And IsArray(saved) Then TestFixtureRestoreInputs config, saved
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' меняет каждый из 32 вариантов размещения через обычное чтение
' таблиц и рецепт. Меняются только три дополнительных стержня выбранной стороны;
' шесть стержней противоположной стороны сохраняются.
Private Sub TestRectanglesIndependentSelectorEffects(ByRef stats As TUiTestStats)
    Dim target As Object, config As Object, saved As Variant, cell As Object, units As CUnitSystem
    Dim baseline As CRebarLayout, changed As CRebarLayout, inactive As CRebarLayout
    Dim pair As Variant, field As Variant, side As Long, rowKey As Variant, diameter As String
    Dim prefix As String, alternative As String, original As String, changedBars As Long
    Dim i As Long, oppositeStart As Long, number As Long, reason As String, invalid As Variant
    On Error GoTo Failed
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = TestFixtureSaveInputs(config)
    Set units = New CUnitSystem: units.InitializeDefaults
    For Each pair In Array(Array("R1.Left", "R1.Right"), Array("R1.Top", "R1.Bottom"), _
            Array("R2.Left", "R2.Right"), Array("R2.Top", "R2.Bottom"))
        For Each field In Array("pos2", "bind2", "pos3", "bind3")
            For side = 0 To 1
                modTestShapeArchitecture.TestConfigureShapeFixture "LSection"
                Set target = ThisWorkbook.Names("rngRectanglesGeometry").RefersToRange
                TestFixtureInputCell(target, "General", "H1").Value2 = 300#: TestFixtureInputCell(target, "General", "B1").Value2 = 200#
                TestFixtureInputCell(target, "General", "H2").Value2 = 200#: TestFixtureInputCell(target, "General", "B2").Value2 = 500#
                For Each rowKey In Array("R1.Top", "R1.Bottom", "R1.Left", "R1.Right", "R2.Top", "R2.Bottom", "R2.Left", "R2.Right")
                    TestFixtureInputCell(target, "Main", CStr(rowKey), "n").Value2 = 0
                    TestFixtureInputCell(target, "Main", CStr(rowKey), "d").Value2 = 0#
                    TestFixtureInputCell(target, "Extra", CStr(rowKey), "d2").Value2 = 0#
                    TestFixtureInputCell(target, "Extra", CStr(rowKey), "d3").Value2 = 0#
                Next rowKey
                diameter = "d" & Right$(CStr(field), 1)
                For Each rowKey In pair
                    TestFixtureInputCell(target, "Main", CStr(rowKey), "as").Value2 = 20#
                    TestFixtureInputCell(target, "Main", CStr(rowKey), "d").Value2 = 10#
                    TestFixtureInputCell(target, "Main", CStr(rowKey), "n").Value2 = 3
                    TestFixtureInputCell(target, "Main", CStr(rowKey), "s0").Value2 = 30#
                    TestFixtureInputCell(target, "Main", CStr(rowKey), "s1").Value2 = 30#
                    TestFixtureInputCell(target, "Extra", CStr(rowKey), diameter).Value2 = 8#
                Next rowKey
                original = "EachBar": alternative = "EverySecondBar"
                If Left$(CStr(field), 3) = "pos" Then original = "Stacked": alternative = "SideBySide"
                Set cell = TestFixtureInputCell(target, "Extra", CStr(pair(side)), CStr(field))
                prefix = "rectangles.independent.effect." & CStr(pair(side)) & "." & CStr(field)
                Set baseline = TestFixtureLayout(TestFixtureWorkbookDefinition("Rectangles", ThisWorkbook, units))
                TestAssertTrue stats, prefix & ".baseline", baseline.Count = 12
                cell.Value2 = alternative
                Set changed = TestFixtureLayout(TestFixtureWorkbookDefinition("Rectangles", ThisWorkbook, units))
                If Left$(CStr(field), 3) = "pos" Then
                    changedBars = 0
                    For i = 1 To baseline.Count
                        If Abs(baseline.X(i) - changed.X(i)) + Abs(baseline.Y(i) - changed.Y(i)) > 0.00000001 Then changedBars = changedBars + 1
                    Next i
                    TestAssertTrue stats, prefix & ".oneSideMoved", changedBars = 3 And changed.Count = 12
                Else
                    TestAssertTrue stats, prefix & ".oneSideBound", changed.Count = 11
                End If
                oppositeStart = 1: If side = 0 Then oppositeStart = 7
                For i = oppositeStart To oppositeStart + 5
                    TestAssertTrue stats, prefix & ".opposite" & CStr(i), TestRectanglesLayoutContainsBar(changed, baseline, i)
                Next i
                For Each invalid In Array(vbNullString, "InvalidChoice", CVErr(xlErrDiv0))
                    cell.Value2 = invalid
                    number = 0: reason = ""
                    On Error Resume Next
                    Set changed = TestFixtureLayout(TestFixtureWorkbookDefinition("Rectangles", ThisWorkbook, units))
                    number = Err.Number: reason = Err.Description: Err.Clear
                    On Error GoTo Failed
                    TestAssertTrue stats, prefix & ".invalidRejected." & TypeName(invalid), number <> 0
                    TestAssertTrue stats, prefix & ".invalidAddress." & TypeName(invalid), InStr(1, reason, cell.Worksheet.Name & "!" & cell.Address(False, False), vbBinaryCompare) > 0
                Next invalid
                TestFixtureInputCell(target, "Extra", CStr(pair(side)), diameter).Value2 = 0#
                Set inactive = TestFixtureLayout(TestFixtureWorkbookDefinition("Rectangles", ThisWorkbook, units))
                TestAssertTrue stats, prefix & ".inactiveIgnored", inactive.Count = 9
                cell.Value2 = original
                Set changed = TestFixtureLayout(TestFixtureWorkbookDefinition("Rectangles", ThisWorkbook, units))
                TestAssertTrue stats, prefix & ".inactiveSameCount", changed.Count = inactive.Count
            Next side
        Next field
    Next pair
    GoTo Finished
Failed:
    stats.Failed = stats.Failed + 1
    TestAppendLine stats, "FAIL: rectangles.independent.effect; " & CStr(Err.Number) & "; " & Err.Description
Finished:
    If Not config Is Nothing And IsArray(saved) Then TestFixtureRestoreInputs config, saved
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' третий ряд учитывает направление и диаметр второго ряда только
' собственной стороны. Совпадение направлений справа не вызывает перескок
' слева, где направления различаются; последующее изменение слева изолировано.
Private Sub TestRectanglesIndependentThirdRows(ByRef stats As TUiTestStats)
    Dim recipe As ISectionRecipe, layout As CRebarLayout, changed As CRebarLayout
    Dim values As Variant, rowKey As Variant, field As Variant
    On Error GoTo Failed
    Set recipe = New CSectionRecipeRectangles
    values = TestFixtureRecipeValues("rngRectanglesGeometry")
    TestFixtureSetInput values, "General", "Shape", "val", "OneRect"
    TestFixtureSetInput values, "General", "H1", "val", 300#: TestFixtureSetInput values, "General", "B1", "val", 200#
    For Each rowKey In Array("R1.Top", "R1.Bottom", "R1.Left", "R1.Right", "R2.Top", "R2.Bottom", "R2.Left", "R2.Right")
        TestFixtureSetInput values, "Main", CStr(rowKey), "as", 50#: TestFixtureSetInput values, "Main", CStr(rowKey), "d", 20#
        TestFixtureSetInput values, "Main", CStr(rowKey), "n", 0
        TestFixtureSetInput values, "Main", CStr(rowKey), "s0", 100#: TestFixtureSetInput values, "Main", CStr(rowKey), "s1", 100#
        For Each field In Array("d2", "a2", "d3", "a3"): TestFixtureSetInput values, "Extra", CStr(rowKey), CStr(field), 0#: Next field
        For Each field In Array("pos2", "pos3"): TestFixtureSetInput values, "Extra", CStr(rowKey), CStr(field), "Stacked": Next field
        For Each field In Array("bind2", "bind3"): TestFixtureSetInput values, "Extra", CStr(rowKey), CStr(field), "EachBar": Next field
    Next rowKey
    TestFixtureSetInput values, "Main", "R1.Left", "n", 1: TestFixtureSetInput values, "Main", "R1.Right", "n", 1
    TestFixtureSetInput values, "Extra", "R1.Left", "d2", 18#: TestFixtureSetInput values, "Extra", "R1.Left", "d3", 28#
    TestFixtureSetInput values, "Extra", "R1.Left", "pos2", "SideBySide"
    TestFixtureSetInput values, "Extra", "R1.Right", "d2", 22#: TestFixtureSetInput values, "Extra", "R1.Right", "d3", 34#
    Set layout = TestFixtureLayout(TestFixtureRecipeDefinition(recipe, "rngRectanglesGeometry", values))
    TestAssertTrue stats, "rectangles.thirdRows.count", layout.Count = 6
    TestAssertTrue stats, "rectangles.thirdRows.leftNoSkip", Abs(layout.X(3) - layout.X(1) - ((20# + 28#) / 2# + 2.5)) < 0.00000001 ' К каждой приставленной паре добавляются 2,5 мм для арматуры периодического профиля.
    TestAssertTrue stats, "rectangles.thirdRows.rightSkip", Abs(layout.X(4) - layout.X(6) - ((20# + 22#) / 2# + 2.5 + (22# + 34#) / 2# + 2.5)) < 0.00000001
    TestFixtureSetInput values, "Extra", "R1.Left", "pos3", "SideBySide"
    Set changed = TestFixtureLayout(TestFixtureRecipeDefinition(recipe, "rngRectanglesGeometry", values))
    TestAssertTrue stats, "rectangles.thirdRows.leftSkip", Abs(Abs(changed.Y(3) - changed.Y(1)) - ((20# + 18#) / 2# + 2.5 + (18# + 28#) / 2# + 2.5)) < 0.00000001
    TestAssertTrue stats, "rectangles.thirdRows.rightUnchanged", TestRectanglesLayoutContainsBar(changed, layout, 6)
    Exit Sub
Failed:
    stats.Failed = stats.Failed + 1
    TestAppendLine stats, "FAIL: rectangles.thirdRows; " & CStr(Err.Number) & "; " & Err.Description
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' связывает столбец пользовательской таблицы с независимым ключом.
Private Function TestRectanglesSelectorTail(ByVal column As Long) As String
    Select Case column
        Case 5: TestRectanglesSelectorTail = "loc_2row"
        Case 4: TestRectanglesSelectorTail = "bind_2row"
        Case 9: TestRectanglesSelectorTail = "loc_3row"
        Case 8: TestRectanglesSelectorTail = "bind_3row"
    End Select
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' ищет исходный стержень в новой раскладке без зависимости от
' перенумерации после изменения привязки дополнительных рядов.
Private Function TestRectanglesLayoutContainsBar(ByVal actual As CRebarLayout, ByVal original As CRebarLayout, ByVal index As Long) As Boolean
    Dim i As Long
    For i = 1 To actual.Count
        If Abs(actual.X(i) - original.X(index)) + Abs(actual.Y(i) - original.Y(index)) + _
                Abs(actual.Diameter(i) - original.Diameter(index)) < 0.00000001 Then
            TestRectanglesLayoutContainsBar = True
            Exit Function
        End If
    Next i
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' запускает поведенческую проверку всех пятнадцати пользовательских
' селекторов единиц/знаков через фактический Config, включая ошибку и recovery.
Public Function RunAudit03UnitSignChoiceTests() As String
    Dim stats As TUiTestStats
    TestAudit03UnitSignChoices stats
    TestAppendLine stats, "TOTAL_AUDIT03_UNIT_SIGN_CHOICES: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunAudit03UnitSignChoiceTests = stats.Report
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' сверяет каждый вариант с независимым коэффициентом пересчета,
' а не только с обратным преобразованием того же адаптера. Ошибочный выбор
' не может молча стать default; исходные формулы обеих таблиц восстанавливаются.
Private Sub TestAudit03UnitSignChoices(ByRef stats As TUiTestStats)
    Dim unitRange As Object, signRange As Object, savedUnits As Variant, savedSigns As Variant
    Dim quantities As Variant, unitQuantities As Variant, choices As Variant, factors As Variant, side As Variant
    Dim q As Long, optionIndex As Long, rowIndex As Long, columnIndex As Long, signIndex As Long
    Dim cell As Object, settings As CSystemSettingsReader, units As CUnitSystem
    Dim key As String, prefix As String, factor As Double, oldValue As Variant
    Dim invalid As Variant, errorNumber As Long, errorDescription As String
    Dim actual As Double, expected As Double, value As Variant, labelRange As Object
    On Error GoTo Failed
    Set unitRange = ThisWorkbook.Names.Item("rngUnitSettings").RefersToRange
    Set signRange = ThisWorkbook.Names.Item("rngSignConventionSettings").RefersToRange
    savedUnits = unitRange.Formula: savedSigns = signRange.Formula
    signRange.Cells(2, 2).Value2 = "Tension"
    signRange.Cells(3, 2).Value2 = "+Y tension"
    signRange.Cells(4, 2).Value2 = "+X tension"
    quantities = Array("Length", "Area", "Force", "Moment", "Stress", "Curvature")
    unitQuantities = quantities
    choices = Array(Array("mm", "cm", "m"), Array("mm2", "cm2", "m2"), _
        Array("N", "kN", "tf"), Array("N*mm", "kN*m", "tf*m"), _
        Array("Pa", "kPa", "MPa", "kgf/cm2", "tf/m2"), Array("1/mm", "1/m"))
    factors = Array(Array(1#, 10#, 1000#), Array(1#, 100#, 1000000#), _
        Array(1#, 1000#, 9806.65), Array(1#, 1000000#, 9806650#), _
        Array(0.000001, 0.001, 1#, 0.0980665, 0.00980665), Array(1#, 0.001))

    For q = 0 To UBound(quantities)
        For Each side In Array("Input", "Output")
            key = "Units." & CStr(quantities(q)) & "." & CStr(side)
            If Not TestUnitSettingAddress(key, rowIndex, columnIndex) Then Err.Raise vbObjectError + 4499, , key
            Set cell = unitRange.Cells(rowIndex, columnIndex)
            oldValue = cell.Formula
            TestAssertTrue stats, "audit03.unitChoice." & key & ".validation", TestValidationCellHasOptions(cell, choices(q))
            For optionIndex = 0 To UBound(choices(q))
                cell.Value2 = choices(q)(optionIndex)
                Set settings = New CSystemSettingsReader
                settings.LoadFromWorkbook ThisWorkbook
                Set units = New CUnitSystem
                units.LoadFromSettings settings
                factor = CDbl(factors(q)(optionIndex))
                prefix = "audit03.unitChoice." & key & "." & CStr(choices(q)(optionIndex))
                For Each value In Array(-123.456, 0#, 123.456)
                    actual = TestAudit03UnitConversion(units, q, CStr(side), CDbl(value))
                    If CStr(side) = "Input" Then expected = CDbl(value) * factor Else expected = CDbl(value) / factor
                    TestAssertClose stats, prefix & ".value" & CStr(value), actual, expected, _
                        0.000000000001 * (1# + Abs(expected))
                Next value
                If q = 0 Then
                    If CStr(side) = "Input" Then
                        actual = units.InputFourthPowerLengthToInternal(2#): expected = 2# * factor ^ 4
                    Else
                        actual = units.InternalFourthPowerLengthToOutput(2#): expected = 2# / factor ^ 4
                        TestAssertClose stats, prefix & ".reverseL4", units.OutputFourthPowerLengthToInternal(2#), _
                            2# * factor ^ 4, 0.000000000001 * (1# + 2# * factor ^ 4)
                    End If
                    TestAssertClose stats, prefix & ".L4", actual, expected, 0.000000000001 * (1# + Abs(expected))
                End If
            Next optionIndex
            For Each invalid In Array("", " ", "TODO", "unknown-unit", "0", CVErr(xlErrValue))
                cell.Value2 = invalid
                On Error Resume Next
                Err.Clear
                Set settings = New CSystemSettingsReader
                settings.LoadFromWorkbook ThisWorkbook
                If Err.Number = 0 Then
                    Set units = New CUnitSystem
                    units.LoadFromSettings settings
                End If
                errorNumber = Err.Number: errorDescription = Err.Description
                On Error GoTo Failed
                TestAssertTrue stats, "audit03.unitChoice." & key & ".invalid." & CStr(VarType(invalid)) & "." & CStr(errorNumber), _
                    errorNumber = vbObjectError + 4530 + q Or errorNumber = vbObjectError + 4309
                TestAssertTrue stats, "audit03.unitChoice." & key & ".invalid.reason", _
                    InStr(1, errorDescription, "Units." & CStr(quantities(q)), vbTextCompare) > 0
                TestAssertTrue stats, "audit03.unitChoice." & key & ".invalid.address", _
                    InStr(1, errorDescription, cell.Address(False, False), vbTextCompare) > 0
                TestAssertTrue stats, "audit03.unitChoice." & key & ".invalid.action", _
                    InStr(1, errorDescription, "Выберите", vbTextCompare) > 0 Or InStr(1, errorDescription, "Исправьте", vbTextCompare) > 0
            Next invalid
            cell.Formula = oldValue
            Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
            Set units = New CUnitSystem: units.LoadFromSettings settings
            TestAssertTrue stats, "audit03.unitChoice." & key & ".recovery", Len(settings.GetRawString(key)) > 0
        Next side
    Next q

    unitRange.Cells(2, 2).Value2 = "mm"
    unitRange.Cells(4, 2).Value2 = "tf": unitRange.Cells(4, 4).Value2 = "tf"
    unitRange.Cells(5, 2).Value2 = "tf*m": unitRange.Cells(5, 4).Value2 = "tf*m"
    unitRange.Cells(6, 2).Value2 = "MPa"
    choices = Array(Array("Tension", "Compression"), Array("+Y tension", "-Y tension"), Array("+X tension", "-X tension"))
    quantities = Array("N", "Mx", "My")
    For signIndex = 0 To 2
        key = "Sign." & CStr(quantities(signIndex)) & ".User"
        Set cell = signRange.Cells(signIndex + 2, 2)
        oldValue = cell.Formula
        TestAssertTrue stats, "audit03.signChoice." & key & ".validation", TestValidationCellHasOptions(cell, choices(signIndex))
        For optionIndex = 0 To 1
            cell.Value2 = choices(signIndex)(optionIndex)
            Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
            Set units = New CUnitSystem: units.LoadFromSettings settings
            factor = 1# - 2# * optionIndex
            If signIndex = 0 Then
                expected = 123# * 9806.65 * factor
                actual = units.InputForceToInternal(123#)
                TestAssertClose stats, "audit03.signChoice.N.output." & CStr(optionIndex), _
                    units.InternalForceToOutput(9806.65), factor, 0.000000000001
            ElseIf signIndex = 1 Then
                expected = 123# * 9806650# * factor
                actual = units.InputMomentMxToInternal(123#)
                TestAssertClose stats, "audit03.signChoice.Mx.output." & CStr(optionIndex), _
                    units.InternalMomentMxToOutput(9806650#), factor, 0.000000000001
            Else
                expected = 123# * 9806650# * factor
                actual = units.InputMomentMyToInternal(123#)
                TestAssertClose stats, "audit03.signChoice.My.output." & CStr(optionIndex), _
                    units.InternalMomentMyToOutput(9806650#), factor, 0.000000000001
            End If
            TestAssertClose stats, "audit03.signChoice." & key & ".input." & CStr(optionIndex), actual, expected, 0.00001
            TestAssertClose stats, "audit03.signChoice." & key & ".stressNotFlipped", units.InputStressToInternal(-12#), -12#, 0.000000000001
            TestAssertClose stats, "audit03.signChoice." & key & ".magnitudeNotFlipped", units.InternalMomentMagnitudeToOutput(9806650#), 1#, 0.000000000001
            TestAssertClose stats, "audit03.signChoice." & key & ".thresholdNotFlipped", units.InputMomentPerLengthToInternal(1#), 9806650#, 0.000001
        Next optionIndex
        For Each invalid In Array("", " ", "TODO", "unknown-sign", "0", CVErr(xlErrValue))
            cell.Value2 = invalid
            On Error Resume Next
            Err.Clear
            Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
            If Err.Number = 0 Then
                Set units = New CUnitSystem: units.LoadFromSettings settings
            End If
            errorNumber = Err.Number: errorDescription = Err.Description
            On Error GoTo Failed
            TestAssertTrue stats, "audit03.signChoice." & key & ".invalid." & CStr(errorNumber), _
                errorNumber = vbObjectError + 4536 + signIndex Or errorNumber = vbObjectError + 4309
            TestAssertTrue stats, "audit03.signChoice." & key & ".invalid.reason", InStr(1, errorDescription, key, vbTextCompare) > 0
            TestAssertTrue stats, "audit03.signChoice." & key & ".invalid.address", _
                InStr(1, errorDescription, cell.Address(False, False), vbTextCompare) > 0
            TestAssertTrue stats, "audit03.signChoice." & key & ".invalid.action", _
                InStr(1, errorDescription, "Выберите", vbTextCompare) > 0 Or InStr(1, errorDescription, "Исправьте", vbTextCompare) > 0
        Next invalid
        cell.Formula = oldValue
        Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
        Set units = New CUnitSystem: units.LoadFromSettings settings
        TestAssertTrue stats, "audit03.signChoice." & key & ".recovery", Len(settings.GetRawString(key)) > 0
    Next signIndex
    ' Поврежденная подпись строки не должна превращать обязательную настройку
    ' полной книги в отсутствующий optional key с внутренним default.
    For q = 0 To 8
        If q < 6 Then
            Set labelRange = unitRange.Cells(q + 2, 1)
            key = "Units." & CStr(unitQuantities(q)) & ".Input"
        Else
            Set labelRange = signRange.Cells(q - 4, 1)
            key = "Sign." & CStr(quantities(q - 6)) & ".User"
        End If
        oldValue = labelRange.Formula
        For Each invalid In Array("", "unknown-quantity")
            labelRange.Value2 = invalid
            On Error Resume Next
            Err.Clear
            Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
            If Err.Number = 0 Then
                Set units = New CUnitSystem: units.LoadFromSettings settings
            End If
            errorNumber = Err.Number: errorDescription = Err.Description
            On Error GoTo Failed
            TestAssertTrue stats, "audit03.unitSignMissing." & key & ".inputError", errorNumber = vbObjectError + 4313
            TestAssertTrue stats, "audit03.unitSignMissing." & key & ".reason", InStr(1, errorDescription, key, vbTextCompare) > 0
        Next invalid
        labelRange.Formula = oldValue
        Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
        Set units = New CUnitSystem: units.LoadFromSettings settings
        TestAssertTrue stats, "audit03.unitSignMissing." & key & ".recovery", settings.HasKey(key)
    Next q
    GoTo Restore
Failed:
    stats.Failed = stats.Failed + 1
    TestAppendLine stats, "FAIL: audit03.unitSignChoices.runtime; " & CStr(Err.Number) & "; " & Err.Description
Restore:
    On Error Resume Next
    If Not unitRange Is Nothing Then unitRange.Formula = savedUnits
    If Not signRange Is Nothing Then signRange.Formula = savedSigns
    On Error GoTo 0
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' выбирает именно проверяемое преобразование. Независимые
' коэффициенты и ожидаемые значения задаются в тесте, не читаются из CUnitSystem.
Private Function TestAudit03UnitConversion(ByVal units As CUnitSystem, ByVal quantity As Long, _
        ByVal side As String, ByVal value As Double) As Double
    If side = "Input" Then
        Select Case quantity
            Case 0: TestAudit03UnitConversion = units.InputLengthToInternal(value)
            Case 1: TestAudit03UnitConversion = units.InputAreaToInternal(value)
            Case 2: TestAudit03UnitConversion = units.InputForceToInternal(value)
            Case 3: TestAudit03UnitConversion = units.InputMomentMxToInternal(value)
            Case 4: TestAudit03UnitConversion = units.InputStressToInternal(value)
            Case 5: TestAudit03UnitConversion = units.InputCurvatureToInternal(value)
        End Select
    Else
        Select Case quantity
            Case 0: TestAudit03UnitConversion = units.InternalLengthToOutput(value)
            Case 1: TestAudit03UnitConversion = units.InternalAreaToOutput(value)
            Case 2: TestAudit03UnitConversion = units.InternalForceToOutput(value)
            Case 3: TestAudit03UnitConversion = units.InternalMomentMxToOutput(value)
            Case 4: TestAudit03UnitConversion = units.InternalStressToOutput(value)
            Case 5: TestAudit03UnitConversion = units.InternalCurvatureToOutput(value)
        End Select
    End If
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' проверяет эквивалентные физические LC при всех восьми знаковых
' соглашениях и девяти парах единиц силы/момента, с фактической записью Results.
Public Function RunAudit03UnitSignEquivalenceTests() As String
    Dim stats As TUiTestStats
    TestAudit03UnitSignEquivalence stats
    TestAppendLine stats, "TOTAL_AUDIT03_UNIT_SIGN_EQUIVALENCE: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunAudit03UnitSignEquivalenceTests = stats.Report
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' один внутренний section используется во всех 72 вариантах.
' Внешние LC читаются обычным reader-ом; сравниваются все сохраненные named-state,
' напряжения, геометрия, физические признаки и metadata. Output-only перевод
' выполняется после solve и не должен запускать ни одного дополнительного НДС.
' Массовый табличный прогон не отображает меняющийся Results: Excel накапливает
' ресурсы от его многократного оформления. Исходный активный лист возвращается;
' короткая stage-диагностика отдельно сохраняет режим воспроизведения этого сбоя.
Private Sub TestAudit03UnitSignEquivalence(ByRef stats As TUiTestStats, Optional ByVal stopAfterCases As Long = 0)
    Dim unitRange As Object, signRange As Object, systemRange As Object, profileRange As Object, loads As Object
    Dim savedUnits As Variant, savedSigns As Variant, savedSystem As Variant, savedProfiles As Variant, savedLoads As Variant
    Dim forceNames As Variant, momentNames As Variant, forceFactors As Variant, momentFactors As Variant
    Dim lengthNames As Variant, areaNames As Variant, lengthFactors As Variant, areaFactors As Variant
    Dim stressNames As Variant, stressFactors As Variant
    Dim nValues As Variant, mxValues As Variant, myValues As Variant, signN As Double, signMx As Double, signMy As Double
    Dim settings As CSystemSettingsReader, units As CUnitSystem, provider As CMaterialModelProvider
    Dim profiles As CCalculationProfileCatalog, reader As CLoadCombinationReader, batch As CBatchSectionCalculator
    Dim section As CSectionModel, mesh As CFiberMeshBuilder, geom As CGeometryRegion, rebars As CRebarLayout
    Dim facade As CSectionDefinitionBuilder, snapshot As CNDMResultsWriter, summary As CBatchResultWriter
    Dim baselineElements As Variant, baselineGeometry As Variant, baselineProperties As Variant
    Dim actual As Variant, signs As Long, forceIndex As Long, momentIndex As Long, caseIndex As Long, i As Long
    Dim lengthFactor As Double, areaFactor As Double, stressFactor As Double, curvatureFactor As Double
    Dim solveCount As Long, prefix As String, savedSheet As Object
    On Error GoTo Failed
    Set savedSheet = ThisWorkbook.Application.ActiveSheet
    If stopAfterCases = 0 Then ThisWorkbook.Worksheets.Item("Config").Activate
    Set unitRange = ThisWorkbook.Names.Item("rngUnitSettings").RefersToRange
    Set signRange = ThisWorkbook.Names.Item("rngSignConventionSettings").RefersToRange
    Set systemRange = ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange
    Set profileRange = ThisWorkbook.Names.Item("rngCalculationProfiles").RefersToRange
    Set loads = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange
    savedUnits = unitRange.Formula: savedSigns = signRange.Formula: savedSystem = systemRange.Formula
    savedProfiles = profileRange.Formula: savedLoads = loads.Formula
    unitRange.Cells(2, 2).Value2 = "mm": unitRange.Cells(2, 4).Value2 = "mm"
    unitRange.Cells(3, 2).Value2 = "mm2": unitRange.Cells(3, 4).Value2 = "mm2"
    unitRange.Cells(6, 2).Value2 = "MPa": unitRange.Cells(6, 4).Value2 = "MPa"
    unitRange.Cells(7, 2).Value2 = "1/mm": unitRange.Cells(7, 4).Value2 = "1/mm"
    TestSetSystemSetting "Calculation.ZeroMomentPerDepth", "0"
    TestSetSystemSetting "Load.ReferenceOffsetX", "10": TestSetSystemSetting "Load.ReferenceOffsetY", "-7"
    TestSetSystemSetting "General.ExecutionReportEnabled", "No"
    TestSetProfileSetting "PR1", "Calculation.Strength.DirectState", "Yes"
    TestSetProfileSetting "PR1", "Calculation.Strength.Capacity", "No"
    TestSetProfileSetting "PR1", "Calculation.Crack.Width", "No"
    TestSetProfileSetting "PR1", "Calculation.Stability.Enabled", "No"
    TestSetProfileSetting "PR2", "Calculation.Strength.DirectState", "No"
    TestSetProfileSetting "PR2", "Calculation.Strength.Capacity", "No"
    TestSetProfileSetting "PR2", "Calculation.Crack.Width", "Yes"
    TestSetProfileSetting "PR2", "Calculation.Stability.Enabled", "No"
    Set geom = TestFixtureDisk(300#)
    Set mesh = New CFiberMeshBuilder: mesh.BuildMesh geom, 20#, 20#, 1
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "UnitSignPaths"
    Set rebars = TestFixtureClosedCountRow(facade.CirclePath("Boundary", 0#, 0#, 150#), _
        facade.CirclePath("Axis", 0#, 0#, 110#), 8, 20#, "Rebar", "Circle", "Круг: основной ряд арматуры")
    Set section = TestBuildGeneratedSectionModel(mesh, rebars, "Audit03UnitSign")
    Set snapshot = New CNDMResultsWriter: Set summary = New CBatchResultWriter
    forceNames = Array("N", "kN", "tf"): forceFactors = Array(1#, 1000#, 9806.65)
    momentNames = Array("N*mm", "kN*m", "tf*m"): momentFactors = Array(1#, 1000000#, 9806650#)
    lengthNames = Array("mm", "cm", "m"): areaNames = Array("mm2", "cm2", "m2")
    lengthFactors = Array(1#, 10#, 1000#): areaFactors = Array(1#, 100#, 1000000#)
    stressNames = Array("MPa", "Pa", "kPa", "kgf/cm2", "tf/m2")
    stressFactors = Array(1#, 0.000001, 0.001, 0.0980665, 0.00980665)
    nValues = Array(-100000#, 30000#, -100000#, 30000#)
    mxValues = Array(-4000000#, 3000000#, -4000000#, 3000000#)
    myValues = Array(-3000000#, -2000000#, -3000000#, -2000000#)
    For signs = 0 To 7
        signN = IIf((signs And 1) = 0, 1#, -1#)
        signMx = IIf((signs And 2) = 0, 1#, -1#)
        signMy = IIf((signs And 4) = 0, 1#, -1#)
        signRange.Cells(2, 2).Value2 = IIf(signN > 0#, "Tension", "Compression")
        signRange.Cells(3, 2).Value2 = IIf(signMx > 0#, "+Y tension", "-Y tension")
        signRange.Cells(4, 2).Value2 = IIf(signMy > 0#, "+X tension", "-X tension")
        For forceIndex = 0 To 2
            For momentIndex = 0 To 2
                prefix = "audit03.unitEquivalent.s" & CStr(signs) & ".f" & CStr(forceIndex) & ".m" & CStr(momentIndex)
                TestAppendLine stats, "RUN: " & prefix & "; " & TestAudit02ExcelMemory()
                unitRange.Cells(4, 2).Value2 = forceNames(forceIndex)
                unitRange.Cells(5, 2).Value2 = momentNames(momentIndex)
                TestSetSystemSetting "Solver.ToleranceN", CStr(0.1 / CDbl(forceFactors(forceIndex)))
                TestSetSystemSetting "Solver.ToleranceMx", CStr(1# / CDbl(momentFactors(momentIndex)))
                TestSetSystemSetting "Solver.ToleranceMy", CStr(1# / CDbl(momentFactors(momentIndex)))
                TestClearDataRows loads
                For i = 0 To 3
                    loads.Cells(i + 2, 1).Value2 = "UNIT" & CStr(i + 1)
                    loads.Cells(i + 2, 2).Value2 = CDbl(nValues(i)) / CDbl(forceFactors(forceIndex)) * signN
                    loads.Cells(i + 2, 3).Value2 = CDbl(mxValues(i)) / CDbl(momentFactors(momentIndex)) * signMx
                    loads.Cells(i + 2, 4).Value2 = CDbl(myValues(i)) / CDbl(momentFactors(momentIndex)) * signMy
                    loads.Cells(i + 2, 5).Value2 = IIf(i < 2, "PR1", "PR2")
                    loads.Cells(i + 2, 6).Value2 = ChrW$(&H3BB) & "*NMxy"
                Next i
                Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
                Set units = New CUnitSystem: units.LoadFromSettings settings
                Set provider = New CMaterialModelProvider: provider.Initialize settings, units
                Set profiles = New CCalculationProfileCatalog: profiles.LoadFromWorkbook ThisWorkbook
                Set batch = New CBatchSectionCalculator
                batch.Initialize section, provider: batch.ApplySettings settings, units
                Set batch.ProfileCatalog = profiles
                Set reader = New CLoadCombinationReader: reader.LoadFromWorkbook ThisWorkbook, batch, units
                batch.ApplyLoadReference 10#, -7#, 0#, 0#
                TestAssertTrue stats, prefix & ".count", batch.Count = 4
                For i = 0 To 3
                    TestAssertClose stats, prefix & ".N." & CStr(i), batch.N(i + 1), CDbl(nValues(i)), 0.000001
                    TestAssertClose stats, prefix & ".Mx." & CStr(i), batch.UserMx(i + 1), CDbl(mxValues(i)), 0.000001
                    TestAssertClose stats, prefix & ".My." & CStr(i), batch.UserMy(i + 1), CDbl(myValues(i)), 0.000001
                Next i
                batch.Execute
                If stopAfterCases > 0 Then TestAppendLine stats, "RUN: " & prefix & ".batch.done; " & TestAudit02ExcelMemory()
                solveCount = batch.SolverCallCount
                TestAssertTrue stats, prefix & ".solved", solveCount > 0
                unitRange.Cells(2, 4).Value2 = lengthNames(caseIndex Mod 3) ' Выходные единицы меняются после получения всех физических состояний.
                unitRange.Cells(3, 4).Value2 = areaNames((caseIndex \ 3) Mod 3)
                unitRange.Cells(4, 4).Value2 = forceNames((caseIndex \ 9) Mod 3)
                unitRange.Cells(5, 4).Value2 = momentNames((caseIndex \ 3) Mod 3)
                unitRange.Cells(6, 4).Value2 = stressNames(caseIndex Mod 5)
                unitRange.Cells(7, 4).Value2 = IIf((caseIndex Mod 2) = 0, "1/mm", "1/m")
                lengthFactor = CDbl(lengthFactors(caseIndex Mod 3))
                areaFactor = CDbl(areaFactors((caseIndex \ 3) Mod 3))
                stressFactor = CDbl(stressFactors(caseIndex Mod 5))
                curvatureFactor = IIf((caseIndex Mod 2) = 0, 1#, 0.001)
                Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
                units.LoadFromSettings settings
                snapshot.WriteResults ThisWorkbook, section, PrepareSectionSnapshot(section, provider), batch, units
                If stopAfterCases > 0 Then TestAppendLine stats, "RUN: " & prefix & ".snapshot.done; " & TestAudit02ExcelMemory()
                summary.WriteSummary ThisWorkbook, batch, units, section
                If stopAfterCases > 0 Then TestAppendLine stats, "RUN: " & prefix & ".summary.done; " & TestAudit02ExcelMemory()
                TestAudit03AssertAutoCADMmGeometry stats, prefix, section
                If stopAfterCases > 0 Then TestAppendLine stats, "RUN: " & prefix & ".readback.done; " & TestAudit02ExcelMemory()
                TestAssertTrue stats, prefix & ".outputDoesNotSolve", solveCount = batch.SolverCallCount
                TestAssertTextEquals stats, prefix & ".outputForceUnit", TestResultsPropertyValue("ALL", "Output.ForceUnit"), units.OutputForceUnit
                TestAssertTextEquals stats, prefix & ".outputStressUnit", TestResultsPropertyValue("ALL", "Output.StressUnit"), CStr(stressNames(caseIndex Mod 5))
                If caseIndex = 0 Then
                    baselineElements = TestResultTable("rngNDMElementResults")
                    baselineGeometry = TestResultTable("rngNDMSectionGeometry")
                    baselineProperties = TestResultTable("rngNDMSectionProperties")
                    TestAssertTrue stats, prefix & ".strengthState", TestResultsPropertyExists("UNIT1", "State.StrengthState.Epsilon0")
                    TestAssertTrue stats, prefix & ".crackedState", TestResultsPropertyExists("UNIT3", "State.CrackedState.Epsilon0")
                Else
                    actual = TestResultTable("rngNDMElementResults")
                    TestAudit03CompareUnitSnapshot stats, prefix & ".elements", baselineElements, actual, "Elements", _
                        lengthFactor, areaFactor, stressFactor, curvatureFactor, _
                        CDbl(forceFactors((caseIndex \ 9) Mod 3)), CDbl(momentFactors((caseIndex \ 3) Mod 3)), signN, signMx, signMy
                    actual = TestResultTable("rngNDMSectionGeometry")
                    TestAudit03CompareUnitSnapshot stats, prefix & ".geometry", baselineGeometry, actual, "Geometry", _
                        lengthFactor, areaFactor, stressFactor, curvatureFactor, _
                        CDbl(forceFactors((caseIndex \ 9) Mod 3)), CDbl(momentFactors((caseIndex \ 3) Mod 3)), signN, signMx, signMy
                    actual = TestResultTable("rngNDMSectionProperties")
                    TestAudit03CompareUnitSnapshot stats, prefix & ".properties", baselineProperties, actual, "Properties", _
                        lengthFactor, areaFactor, stressFactor, curvatureFactor, _
                        CDbl(forceFactors((caseIndex \ 9) Mod 3)), CDbl(momentFactors((caseIndex \ 3) Mod 3)), signN, signMx, signMy
                End If
                caseIndex = caseIndex + 1
                If stopAfterCases > 0 Then
                    If caseIndex >= stopAfterCases Then GoTo Restore
                End If
            Next momentIndex
        Next forceIndex
    Next signs
    TestAssertTrue stats, "audit03.unitEquivalent.allCases", caseIndex = 72
    GoTo Restore
Failed:
    stats.Failed = stats.Failed + 1
    TestAppendLine stats, "FAIL: audit03.unitEquivalent.runtime; " & prefix & "; " & CStr(Err.Number) & "; " & Err.Description
Restore:
    On Error Resume Next
    If Not unitRange Is Nothing Then unitRange.Formula = savedUnits
    If Not signRange Is Nothing Then signRange.Formula = savedSigns
    If Not systemRange Is Nothing Then systemRange.Formula = savedSystem
    If Not profileRange Is Nothing Then profileRange.Formula = savedProfiles
    If Not loads Is Nothing Then loads.Formula = savedLoads
    If stopAfterCases = 0 Then
        If Not savedSheet Is Nothing Then savedSheet.Activate
    End If
    On Error GoTo 0
    If stopAfterCases = 0 Then TestAssertTrue stats, "audit03.unitEquivalent.activeSheetRestored", _
        ThisWorkbook.Application.ActiveSheet Is savedSheet
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' два первых полных unit/sign-варианта локализуют расход памяти
' между расчетом, snapshot writer, summary writer и чтением AutoCAD-данных.
' Это диагностический срез, не замена штатных 72 вариантов и их приемки.
Public Function RunAudit03UnitSignStageDiagnosticTests() As String
    Dim stats As TUiTestStats
    TestAudit03UnitSignEquivalence stats, 2
    TestAppendLine stats, "TOTAL_UNIT_SIGN_STAGE_DIAGNOSTIC: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunAudit03UnitSignStageDiagnosticTests = stats.Report
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' повторяет те же два варианта с изоляцией режима приложения.
' No/Yes меняют только ScreenUpdating, Config проверяет влияние активного листа,
' Guard использует штатный быстрый режим. Все исходные настройки возвращаются.
Public Function RunAudit03UnitSignStageScreenDiagnosticTests(Optional ByVal screenMode As String = "No") As String
    Dim savedScreen As Boolean, stats As TUiTestStats, guard As CExcelAppStateGuard, savedSheet As Object
    savedScreen = ThisWorkbook.Application.ScreenUpdating
    Set savedSheet = ThisWorkbook.Application.ActiveSheet
    On Error GoTo Failed
    Select Case screenMode
        Case "No", "Yes"
            ThisWorkbook.Application.ScreenUpdating = (screenMode = "Yes")
        Case "Config"
            ThisWorkbook.Worksheets.Item("Config").Activate
        Case "Guard"
            Set guard = New CExcelAppStateGuard
            guard.Enter ThisWorkbook.Application
        Case Else
            Err.Raise vbObjectError + 4496, "RunAudit03UnitSignStageScreenDiagnosticTests", "Неизвестный режим диагностического теста."
    End Select
    TestAppendLine stats, "RUN: unitStage.screenUpdating=" & CStr(ThisWorkbook.Application.ScreenUpdating)
    TestAudit03UnitSignEquivalence stats, 2
    GoTo Restore
Failed:
    stats.Failed = stats.Failed + 1
    TestAppendLine stats, "FAIL: unitStage.screen.runtime; " & CStr(Err.Number) & "; " & Err.Description
Restore:
    If Not guard Is Nothing Then guard.Restore
    savedSheet.Activate
    ThisWorkbook.Application.ScreenUpdating = savedScreen
    TestAssertTrue stats, "unitStage.screen.restored", ThisWorkbook.Application.ScreenUpdating = savedScreen
    TestAppendLine stats, "TOTAL_UNIT_SIGN_STAGE_SCREEN: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunAudit03UnitSignStageScreenDiagnosticTests = stats.Report
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' нормализует только выводимые численные поля независимыми
' коэффициентами. Strain и PhysicalState не меняются от пользовательских
' знаков; знак момента применяется к усилиям, но не к внутренней KappaX/Y.
Private Sub TestAudit03CompareUnitSnapshot(ByRef stats As TUiTestStats, ByVal prefix As String, _
        ByRef baseline As Variant, ByRef actual As Variant, ByVal kind As String, _
        ByVal lengthFactor As Double, ByVal areaFactor As Double, ByVal stressFactor As Double, _
        ByVal curvatureFactor As Double, ByVal forceFactor As Double, ByVal momentFactor As Double, _
        ByVal signN As Double, ByVal signMx As Double, ByVal signMy As Double, _
        Optional ByVal compareInertiaRoundoff As Boolean = False)
    Dim row As Long, column As Long, factor As Double, parameter As String, tolerance As Double
    Dim same As Boolean, detail As String, expected As Double, converted As Double, absoluteLimit As Double
    same = (UBound(baseline, 1) = UBound(actual, 1) And UBound(baseline, 2) = UBound(actual, 2))
    If Not same Then detail = "размер таблицы" Else detail = vbNullString
    If same Then
        For row = 2 To UBound(baseline, 1)
            For column = 2 To UBound(baseline, 2)
                factor = 1#: tolerance = 0.000000001
                If kind = "Elements" Then
                    If column = 8 Then factor = stressFactor
                    If column = 7 Then tolerance = 0.0000000001
                ElseIf kind = "Geometry" Then
                    Select Case column
                        Case 4, 5, 8, 9, 10: factor = lengthFactor
                        Case 6: factor = areaFactor
                        Case 12, 13, 14: factor = lengthFactor ^ 4
                    End Select
                Else
                    If column = 5 Or column = 6 Then GoTo NextColumn
                    parameter = CStr(baseline(row, 3))
                    If Left$(parameter, 7) = "Output." Then GoTo NextColumn
                    If column = 4 Then
                        Select Case CStr(baseline(row, 5))
                            Case "mm": factor = lengthFactor
                            Case "mm2": factor = areaFactor
                            Case "mm4": factor = lengthFactor ^ 4
                            Case "MPa": factor = stressFactor
                            Case "1/mm": factor = curvatureFactor: tolerance = 0.000000000001
                            Case "N": factor = forceFactor * signN
                            Case "N*mm"
                                If Left$(parameter, 2) = "My" Then factor = momentFactor * signMy Else factor = momentFactor * signMx
                        End Select
                    End If
                End If
                If IsNumeric(baseline(row, column)) And IsNumeric(actual(row, column)) Then
                    expected = CDbl(baseline(row, column)): converted = CDbl(actual(row, column)) * factor
                    If compareInertiaRoundoff And kind = "Properties" And column = 4 Then
                        Dim axisPrefix As String, opposite As String
                        axisPrefix = Split(parameter, ".")(0): opposite = vbNullString
                        Select Case Mid$(parameter, Len(axisPrefix) + 2)
                            Case "CoreDistance1Plus", "CoreDistance2Plus": opposite = Left$(parameter, Len(parameter) - 4) & "Minus"
                            Case "CoreDistance1Minus", "CoreDistance2Minus": opposite = Left$(parameter, Len(parameter) - 5) & "Plus"
                        End Select
                        If Len(opposite) > 0 Then
                            If TestAudit03PrincipalAxisSign(baseline, actual, axisPrefix) < 0# Then converted = TestAudit03PropertyNumber(actual, opposite) * factor
                        End If
                    End If
                    absoluteLimit = tolerance * (1# + Abs(expected))
                    If kind = "Properties" And column = 4 Then
                        If CStr(baseline(row, 5)) = "1/mm" Then
                            ' Около нуля сохраняется абсолютный порог 1e-12.
                            ' Относительное округление большого коэффициента
                            ' согласовано с проверкой деформаций Elements (1e-10).
                            absoluteLimit = 0.000000000001 + 0.0000000001 * Abs(expected)
                            If compareInertiaRoundoff Then
                                ' ДЛЯ ТЕСТОВ: INPUT tf/м2 даёт относительное
                                ' округление кривизны 2.73e-10 на плато в обеих
                                ' книгах. Используем общий предел unit-сравнения
                                ' 1e-9; абсолютный порог и проверки деформаций,
                                ' усилий, статусов и геометрии остаются прежними.
                                absoluteLimit = 0.000000000001 + 0.000000001 * Abs(expected)
                            End If
                        End If
                    End If
                    If compareInertiaRoundoff And column = 4 And kind = "Properties" Then
                        If parameter = "Transformed.Ixy" Or parameter = "Transformed.Ixyc" Then
                            Dim roundoffLimit As Double
                            roundoffLimit = TestAudit03InertiaRoundoffLimit(baseline, parameter)
                            If Abs(converted - expected) > absoluteLimit Then
                                TestAppendLine stats, "INERTIA_ROUNDOFF: " & prefix & "; parameter=" & parameter & _
                                    "; expected=" & CStr(expected) & "; actual=" & CStr(converted) & "; limit=" & CStr(roundoffLimit)
                            End If
                            If roundoffLimit > absoluteLimit Then absoluteLimit = roundoffLimit
                        End If
                    End If
                    ' Главная ось — прямая без выделенного положительного направления.
                    ' Поворот на pi допустим; поворот на pi/2 переставляет оси и не принимается.
                    If compareInertiaRoundoff And column = 4 And kind = "Properties" And _
                            (parameter = "Concrete.PrincipalAngle" Or parameter = "Transformed.PrincipalAngle") Then
                        If Abs(Sin(converted - expected)) <= 0.000000001 Then converted = expected
                    End If
                    If Abs(converted - expected) > absoluteLimit Then same = False
                ElseIf CStr(baseline(row, column)) <> CStr(actual(row, column)) Then
                    same = False
                End If
                If Not same Then
                    detail = "row=" & CStr(row) & "; col=" & CStr(column) & "; parameter=" & parameter & _
                        "; expected=" & CStr(baseline(row, column)) & "; actual=" & CStr(actual(row, column)) & "; factor=" & CStr(factor)
                    Exit For
                End If
NextColumn:
            Next column
            If Not same Then Exit For
        Next row
    End If
    TestAssertTrue stats, prefix & ".equivalent; " & detail, same
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Около нулевого Ixy сравниваем unit-roundtrip относительно sqrt(Ix*Iy).
' Инженерный относительный порог 1e-9 совпадает с общим сравнением Properties:
' машинный шум не требует дополнительного прохода расчёта геометрии.
' Правило включается только для Transformed.Ixy/Ixyc при compareInertiaRoundoff=True;
' пороги solver/прочности/трещин/устойчивости и исторический unit-sign тест не меняются.
Private Function TestAudit03InertiaRoundoffLimit(ByRef properties As Variant, ByVal parameter As String) As Double
    Dim ix As Double, iy As Double, row As Long, suffix As String
    If parameter = "Transformed.Ixyc" Then suffix = "c"
    For row = 2 To UBound(properties, 1)
        If CStr(properties(row, 3)) = "Transformed.Ix" & suffix Then ix = CDbl(properties(row, 4))
        If CStr(properties(row, 3)) = "Transformed.Iy" & suffix Then iy = CDbl(properties(row, 4))
    Next row
    TestAudit03InertiaRoundoffLimit = 0.000000001 * Sqr(Abs(ix)) * Sqr(Abs(iy))
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' проверяет обычное восстановление геометрии для экспорта AutoCAD
' при всех выходных единицах. Координаты и диаметры должны возвращаться
' в мм, площади — в мм2, моменты инерции — в мм4.
Private Sub TestAudit03AssertAutoCADMmGeometry(ByRef stats As TUiTestStats, ByVal prefix As String, _
        ByVal expected As CSectionModel)
    Dim actual As CSectionModel, beforeSolves As Long, same As Boolean, i As Long, j As Long
    Dim values As Variant, reference As Variant, detail As String
    beforeSolves = SectionEquilibriumSolveCount()
    Set actual = ReadSectionGeometryFromResults(ThisWorkbook)
    same = (actual.ConcreteCount = expected.ConcreteCount And actual.RebarCount = expected.RebarCount)
    If Not same Then detail = "число элементов"
    If same Then
        For i = 1 To expected.ConcreteCount + expected.RebarCount
            If i <= expected.ConcreteCount Then
                values = Array(actual.ConcreteX(i), actual.ConcreteY(i), actual.ConcreteArea(i), _
                    actual.ConcreteWidth(i), actual.ConcreteHeight(i), actual.ConcreteRotation(i), _
                    actual.ConcreteLocalIx(i), actual.ConcreteLocalIy(i), actual.ConcreteLocalIxy(i))
                reference = Array(expected.ConcreteX(i), expected.ConcreteY(i), expected.ConcreteArea(i), _
                    expected.ConcreteWidth(i), expected.ConcreteHeight(i), expected.ConcreteRotation(i), _
                    expected.ConcreteLocalIx(i), expected.ConcreteLocalIy(i), expected.ConcreteLocalIxy(i))
            Else
                Dim bar As Long
                bar = i - expected.ConcreteCount
                values = Array(actual.RebarX(bar), actual.RebarY(bar), actual.RebarArea(bar), actual.RebarDiameter(bar))
                reference = Array(expected.RebarX(bar), expected.RebarY(bar), expected.RebarArea(bar), expected.RebarDiameter(bar))
            End If
            For j = 0 To UBound(values)
                If Abs(CDbl(values(j)) - CDbl(reference(j))) > 0.000000001 * (1# + Abs(CDbl(reference(j)))) Then
                    same = False
                    detail = "element=" & CStr(i) & "; field=" & CStr(j) & _
                        "; expected=" & CStr(reference(j)) & "; actual=" & CStr(values(j))
                    Exit For
                End If
            Next j
            If Not same Then Exit For
        Next i
    End If
    TestAssertTrue stats, prefix & ".autoCADGeometryMm; " & detail, same
    TestAssertTrue stats, prefix & ".autoCADDoesNotSolve", SectionEquilibriumSolveCount() = beforeSolves
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' отдельный сквозной gate INPUT длины, напряжения и кривизны.
' Все сочетания единиц проходят обычный макрос расчета, а не ручной solver.
Public Function RunAudit03InputUnitConsumerTests() As String
    Dim stats As TUiTestStats
    TestAudit03InputUnitConsumers stats
    TestAppendLine stats, "TOTAL_AUDIT03_INPUT_UNIT_CONSUMERS: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunAudit03InputUnitConsumerTests = stats.Report
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' четыре формы и 30 эквивалентных систем INPUT для каждой.
' Геометрия, материалы, допуски, crack width и устойчивость строятся из Config;
' сравниваются численные снимки и все комментарии подробных output-блоков.
' Измененные таблицы и активный лист восстанавливаются даже после ошибки;
' массовый прогон не отображает Results во время каждой табличной записи.
Private Sub TestAudit03InputUnitConsumers(ByRef stats As TUiTestStats)
    TestAudit03LimitCommentComparison stats
    Dim rangeNames As Variant, original As Collection, prepared As Collection, name As Variant
    Dim target As Object, loads As Object, lengthNames As Variant, lengthFactors As Variant
    Dim stressNames As Variant, stressFactors As Variant, shape As Variant
    Dim lengthIndex As Long, stressIndex As Long, curvatureIndex As Long, i As Long, caseCount As Long
    Dim baselineElements As Variant, baselineGeometry As Variant, baselineProperties As Variant
    Dim baselineDiagrams As Variant, baselineStrength As Variant, baselineCrack As Variant, baselineStability As Variant
    Dim actual As Variant, prefix As String, message As String, caseFailures As Long, savedSheet As Object
    Dim principalSign As Double
    On Error GoTo Failed
    Set savedSheet = ThisWorkbook.Application.ActiveSheet
    ThisWorkbook.Worksheets.Item("Config").Activate
    rangeNames = Array("rngSystemSettings", "rngUnitSettings", "rngSignConventionSettings", _
        "rngCalculationProfiles", "rngLoadCombinations", "rngConcreteMaterialParameters", _
        "rngSteelMaterialParameters", "rngCircleGeometry", "rngRectanglesGeometry", _
        "rngPointedSectionGeometry", "rngHollowRectangleGeometry")
    Set original = New Collection
    For Each name In rangeNames
        original.Add ThisWorkbook.Names.Item(CStr(name)).RefersToRange.Formula
    Next name
    For Each name In Array("Length", "Area", "Force", "Moment", "Stress", "Curvature")
        Select Case CStr(name)
            Case "Length": message = "mm"
            Case "Area": message = "mm2"
            Case "Force": message = "N"
            Case "Moment": message = "N*mm"
            Case "Stress": message = "MPa"
            Case "Curvature": message = "1/mm"
        End Select
        TestSetSystemSetting "Units." & CStr(name) & ".Input", message
        TestSetSystemSetting "Units." & CStr(name) & ".Output", message
    Next name
    TestSetSystemSetting "Sign.N.User", "Tension"
    TestSetSystemSetting "Sign.Mx.User", "+Y tension"
    TestSetSystemSetting "Sign.My.User", "+X tension"
    TestSetSystemSetting "Geometry.Source", "Generated"
    TestSetSystemSetting "General.ExecutionReportEnabled", "No"
    TestSetSystemSetting "Plot.AutoUpdateAfterCalculation", "No"
    TestSetSystemSetting "Calculation.ZeroMomentPerDepth", "0.001"
    TestSetSystemSetting "Load.ReferenceOffsetX", "10"
    TestSetSystemSetting "Load.ReferenceOffsetY", "-7"
    TestSetSystemSetting "Mesh.StepMode", "User"
    TestSetSystemSetting "Mesh.StepX", "50": TestSetSystemSetting "Mesh.StepY", "50"
    TestSetSystemSetting "Mesh.BoundarySubdivisions", "1"
    TestSetSystemSetting "Solver.ToleranceN", "0.1"
    TestSetSystemSetting "Solver.ToleranceMx", "1": TestSetSystemSetting "Solver.ToleranceMy", "1"
    TestSetSystemSetting "Capacity.SolutionStrategy", "LoadMultiplier"
    TestSetSystemSetting "Capacity.SearchMethod", "Bisection"
    TestSetSystemSetting "SLS.Crack.InitiationSolutionStrategy", "Auto"
    TestSetSystemSetting "SLS.Crack.Allowable", "0.3"
    TestSetSystemSetting "Stability.ElementLength", "8000"
    For Each name In Array("PR1", "PR2")
        TestSetProfileSetting CStr(name), "Calculation.Strength.DirectState", IIf(CStr(name) = "PR1", "Yes", "No")
        TestSetProfileSetting CStr(name), "Calculation.Strength.Capacity", IIf(CStr(name) = "PR1", "Yes", "No")
        TestSetProfileSetting CStr(name), "Calculation.Crack.Width", IIf(CStr(name) = "PR2", "Yes", "No")
        TestSetProfileSetting CStr(name), "Calculation.Stability.Enabled", "Yes"
    Next name
    Set loads = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange
    TestClearDataRows loads
    For i = 0 To 3
        loads.Cells(i + 2, 1).Value2 = "INPUT" & CStr(i + 1)
        loads.Cells(i + 2, 2).Value2 = IIf((i Mod 2) = 0, -100000#, 30000#)
        loads.Cells(i + 2, 3).Value2 = IIf((i Mod 2) = 0, 30000000#, -20000000#)
        loads.Cells(i + 2, 4).Value2 = IIf((i Mod 2) = 0, -20000000#, 15000000#)
        ' Для трещин сравниваем найденное НДС при эксплуатационных моментах,
        ' а не последние приближения поиска, который закончился без равновесия.
        If i >= 2 Then
            loads.Cells(i + 2, 3).Value2 = loads.Cells(i + 2, 3).Value2 / 10#
            loads.Cells(i + 2, 4).Value2 = loads.Cells(i + 2, 4).Value2 / 10#
        End If
        loads.Cells(i + 2, 5).Value2 = IIf(i < 2, "PR1", "PR2")
        loads.Cells(i + 2, 6).Value2 = ChrW$(&H3BB) & "*NMxy"
    Next i
    Set prepared = New Collection
    For Each name In rangeNames
        prepared.Add ThisWorkbook.Names.Item(CStr(name)).RefersToRange.Formula
    Next name
    lengthNames = Array("mm", "cm", "m"): lengthFactors = Array(1#, 10#, 1000#)
    stressNames = Array("MPa", "Pa", "kPa", "kgf/cm2", "tf/m2")
    stressFactors = Array(1#, 0.000001, 0.001, 0.0980665, 0.00980665)
    For Each shape In Array("Circle", "Rectangles", "PointedSection", "HollowRectangle")
        For lengthIndex = 0 To 2
            For stressIndex = 0 To 4
                For curvatureIndex = 0 To 1
                    For i = 0 To UBound(rangeNames)
                        ThisWorkbook.Names.Item(CStr(rangeNames(i))).RefersToRange.Formula = prepared.Item(i + 1)
                    Next i
                    TestSetSystemSetting "Geometry.Type", CStr(shape)
                    TestSetSystemSetting "Units.Length.Input", CStr(lengthNames(lengthIndex))
                    TestSetSystemSetting "Units.Stress.Input", CStr(stressNames(stressIndex))
                    TestSetSystemSetting "Units.Curvature.Input", IIf(curvatureIndex = 0, "1/mm", "1/m")
                    TestAudit03RescaleInputTables CDbl(lengthFactors(lengthIndex)), CDbl(stressFactors(stressIndex))
                    prefix = "audit03.inputUnits." & CStr(shape) & ".l" & CStr(lengthIndex) & _
                        ".s" & CStr(stressIndex) & ".k" & CStr(curvatureIndex)
                    TestAppendLine stats, "RUN: " & prefix
                    caseFailures = stats.Failed
                    message = TestFixtureRunSectionWorkflow(ThisWorkbook, False)
                    TestAssertTrue stats, prefix & ".completed", Len(message) > 0
                    TestAssertTrue stats, prefix & ".strengthState", TestResultsPropertyExists("INPUT1", "State.StrengthState.Epsilon0")
                    TestAssertTrue stats, prefix & ".capacityState", TestResultsPropertyExists("INPUT1", "State.CapacityState.Epsilon0")
                    TestAssertTrue stats, prefix & ".crackedState", TestResultsPropertyExists("INPUT3", "State.CrackedState.Epsilon0")
                    TestAssertTextEquals stats, prefix & ".crackedState.INPUT3.ok", TestResultsPropertyValue("INPUT3", "State.CrackedState.Status"), "OK"
                    TestAssertTextEquals stats, prefix & ".crackedState.INPUT4.ok", TestResultsPropertyValue("INPUT4", "State.CrackedState.Status"), "OK"
                    If lengthIndex = 0 And stressIndex = 0 And curvatureIndex = 0 Then
                        baselineElements = TestResultTable("rngNDMElementResults")
                        baselineGeometry = TestResultTable("rngNDMSectionGeometry")
                        baselineProperties = TestResultTable("rngNDMSectionProperties")
                        baselineDiagrams = TestResultTable("rngNDMMaterialDiagrams")
                        baselineStrength = ThisWorkbook.Names.Item("rngStrengthSummaryAnchor").RefersToRange.Resize(4, 49).Value2
                        baselineCrack = ThisWorkbook.Names.Item("rngCrackSummaryAnchor").RefersToRange.Resize(4, 72).Value2
                        baselineStability = ThisWorkbook.Names.Item("rngStabilitySummaryAnchor").RefersToRange.Resize(4, 84).Value2
                    Else
                        actual = TestResultTable("rngNDMElementResults")
                        TestAudit03CompareUnitSnapshot stats, prefix & ".elements", baselineElements, actual, "Elements", 1#, 1#, 1#, 1#, 1#, 1#, 1#, 1#, 1#
                        actual = TestResultTable("rngNDMSectionGeometry")
                        TestAudit03CompareUnitSnapshot stats, prefix & ".geometry", baselineGeometry, actual, "Geometry", 1#, 1#, 1#, 1#, 1#, 1#, 1#, 1#, 1#
                        actual = TestResultTable("rngNDMSectionProperties")
                        principalSign = TestAudit03PrincipalAxisSign(baselineProperties, actual)
                        TestAudit03CompareUnitSnapshot stats, prefix & ".properties", baselineProperties, actual, "Properties", 1#, 1#, 1#, 1#, 1#, 1#, 1#, 1#, 1#, True
                        actual = TestResultTable("rngNDMMaterialDiagrams")
                        TestAudit03ComparePlainSnapshot stats, prefix & ".diagrams", baselineDiagrams, actual, 2
                        actual = ThisWorkbook.Names.Item("rngStrengthSummaryAnchor").RefersToRange.Resize(4, 49).Value2
                        TestAudit03ComparePlainSnapshot stats, prefix & ".strength", baselineStrength, actual, 1
                        actual = ThisWorkbook.Names.Item("rngCrackSummaryAnchor").RefersToRange.Resize(4, 72).Value2
                        TestAudit03ComparePlainSnapshot stats, prefix & ".crack", baselineCrack, actual, 1, True
                        actual = ThisWorkbook.Names.Item("rngStabilitySummaryAnchor").RefersToRange.Resize(4, 84).Value2
                        TestAudit03ComparePlainSnapshot stats, prefix & ".stability", baselineStability, actual, 1, False, True, principalSign
                    End If
                    caseCount = caseCount + 1
                    TestAppendLine stats, "INPUT_UNIT_CASE: " & prefix & "; newFailures=" & CStr(stats.Failed - caseFailures)
                Next curvatureIndex
            Next stressIndex
        Next lengthIndex
    Next shape
    TestAssertTrue stats, "audit03.inputUnits.allCases", caseCount = 120
    GoTo Restore
Failed:
    stats.Failed = stats.Failed + 1
    TestAppendLine stats, "FAIL: audit03.inputUnits.runtime; " & prefix & "; " & CStr(Err.Number) & "; " & Err.Description
Restore:
    On Error Resume Next
    If Not original Is Nothing Then
        For i = 0 To original.Count - 1
            ThisWorkbook.Names.Item(CStr(rangeNames(i))).RefersToRange.Formula = original.Item(i + 1)
        Next i
    End If
    If Not savedSheet Is Nothing Then savedSheet.Activate
    On Error GoTo 0
    TestAssertTrue stats, "audit03.inputUnits.activeSheetRestored", ThisWorkbook.Application.ActiveSheet Is savedSheet
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' изменяет только размерные числовые поля согласно видимым
' подтаблицам Config. Счетчики, enum, деформации, пустые диаметры и формулы
' единиц не масштабируются. Коэффициенты независимы от CUnitSystem.
Private Sub TestAudit03RescaleInputTables(ByVal lengthFactor As Double, ByVal stressFactor As Double)
    Dim target As Object, name As Variant, row As Long, column As Variant, key As String
    For Each name In Array("rngSystemSettings")
        Set target = ThisWorkbook.Names.Item(CStr(name)).RefersToRange
        For row = 2 To target.Rows.Count
            key = CStr(target.Cells(row, 1).Value2)
            Select Case key
                Case "Circle.Diameter", "Rebar.AxisDistance", "Rebar.Diameter", "Rebar.Diameter2", "Rebar.Diameter3", _
                        "Mesh.StepX", "Mesh.StepY", "Load.ReferenceOffsetX", "Load.ReferenceOffsetY", _
                        "Stability.ElementLength", "Stability.AccidentalEccentricityUser1", _
                        "Stability.AccidentalEccentricityUser2", "SLS.Crack.Allowable", "Plot.ResultLabelSpacing", _
                        "SLS.Crack.SP35.GroupGapTolerance", "SLS.Crack.SP35.RowTolerance"
                    TestAudit03DivideNumericInput target.Cells(row, 2), lengthFactor
            End Select
        Next row
    Next name
    TestSetSystemSetting "Calculation.ZeroMomentPerDepth", CStr(0.001 * lengthFactor) ' Порог имеет размерность момент/длина; входные единицы момента не меняются.
    Dim fieldReader As CSectionInputReader, fieldKey As Variant, fieldRow As Long, fieldColumn As Long
    For Each name In Array("rngCircleGeometry", "rngRectanglesGeometry", "rngPointedSectionGeometry", "rngHollowRectangleGeometry")
        Set target = ThisWorkbook.Names.Item(CStr(name)).RefersToRange
        Set fieldReader = TestFixtureInput(target.Value2)
        For Each fieldKey In fieldReader.RowKeys("General")
            If CStr(fieldKey) <> "Shape" And InStr(CStr(fieldKey), ".Type") = 0 Then
                TestAudit03DivideNumericInput TestFixtureInputCell(target, "General", CStr(fieldKey)), lengthFactor
            End If
        Next fieldKey
    Next name
    Dim geometryName As Variant, blockName As Variant
    For Each geometryName In Array("rngCircleGeometry", "rngRectanglesGeometry", "rngPointedSectionGeometry", "rngHollowRectangleGeometry")
        Set target = ThisWorkbook.Names(CStr(geometryName)).RefersToRange
        Set fieldReader = TestFixtureInput(target.Value2)
        For Each blockName In Array("Main", "Extra")
            For Each fieldKey In fieldReader.RowKeys(CStr(blockName))
                If CStr(fieldKey) <> "Units" Then
                    If CStr(blockName) = "Main" Then
                        For Each column In Array("as", "d", "s0", "s1")
                            TestAudit03DivideNumericInput TestFixtureInputCell(target, CStr(blockName), CStr(fieldKey), CStr(column)), lengthFactor
                        Next column
                    Else
                        For Each column In Array("d2", "a2", "d3", "a3")
                            TestAudit03DivideNumericInput TestFixtureInputCell(target, CStr(blockName), CStr(fieldKey), CStr(column)), lengthFactor
                        Next column
                    End If
                End If
            Next fieldKey
        Next blockName
    Next geometryName
    For Each name In Array("rngConcreteMaterialParameters", "rngSteelMaterialParameters")
        Set target = ThisWorkbook.Names.Item(CStr(name)).RefersToRange
        For row = 2 To target.Rows.Count
            key = CStr(target.Cells(row, 1).Value2)
            If InStr(key, ".R") > 0 Or key = "Concrete.E" Or key = "Steel.E" Then
                For column = 2 To 3: TestAudit03DivideNumericInput target.Cells(row, column), stressFactor: Next column
            End If
        Next row
    Next name
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' сохраняет пустой ввод и служебный прочерк; масштабирует число
' в ячейке fixture, не вызывая production-преобразования и не меняя validation.
Private Sub TestAudit03DivideNumericInput(ByVal cell As Object, ByVal factor As Double)
    If IsNumeric(cell.Value2) And Len(CStr(cell.Value2)) > 0 Then cell.Value2 = CDbl(cell.Value2) / factor
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' сравнивает фактические output-таблицы при одинаковых OUTPUT
' единицах. Тексты и статусы совпадают точно; числа допускают округление 1e-9.
' Только при сравнении комментариев о трещинах можно отдельно сверить lambda
' с допуском 1e-12; допуски расчётного решателя при этом не меняются.
Private Sub TestAudit03ComparePlainSnapshot(ByRef stats As TUiTestStats, ByVal prefix As String, _
        ByRef baseline As Variant, ByRef actual As Variant, ByVal firstColumn As Long, _
        Optional ByVal compareCommentLambda As Boolean = False, _
        Optional ByVal compareStability As Boolean = False, Optional ByVal principalSign As Double = 1#)
    Dim row As Long, column As Long, same As Boolean, detail As String, converted As Double
    same = (UBound(baseline, 1) = UBound(actual, 1) And UBound(baseline, 2) = UBound(actual, 2))
    If Not same Then detail = "размер таблицы"
    If same Then
        For row = 1 To UBound(baseline, 1)
            For column = firstColumn To UBound(baseline, 2)
                If IsNumeric(baseline(row, column)) And IsNumeric(actual(row, column)) Then
                    converted = CDbl(actual(row, column))
                    If compareStability Then
                        Select Case column
                            Case 7, 8, 17, 18, 20, 21, 30, 31, 45, 46, 62, 63, 74, 75
                                converted = converted * principalSign
                        End Select
                    End If
                    same = Abs(converted - CDbl(baseline(row, column))) <= _
                        0.000000001 * (1# + Abs(CDbl(baseline(row, column))))
                ElseIf compareStability And column = 2 Then
                    same = TestAudit03SameStabilityComment(CStr(baseline(row, column)), CStr(actual(row, column)))
                ElseIf compareCommentLambda And column = 2 Then
                    same = TestAudit03SameLimitComment(CStr(baseline(row, column)), CStr(actual(row, column)))
                Else
                    same = (CStr(baseline(row, column)) = CStr(actual(row, column)))
                End If
                If Not same Then
                    detail = "row=" & CStr(row) & "; col=" & CStr(column) & _
                        "; expected=" & CStr(baseline(row, column)) & "; actual=" & CStr(actual(row, column))
                    Exit For
                End If
            Next column
            If Not same Then Exit For
        Next row
    End If
    TestAssertTrue stats, prefix & ".equivalent; " & detail, same
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' число lambda внутри комментария сравнивает как число с допуском
' 1e-12 от его величины; весь окружающий текст должен совпасть побуквенно.
' Это учитывает последние разряды Double после смены единиц, сохраняя проверку причины и статуса.
Private Function TestAudit03SameLimitComment(ByVal expected As String, ByVal actual As String) As Boolean
    On Error GoTo Different
    If expected = actual Then
        TestAudit03SameLimitComment = True
        Exit Function
    End If
    Const marker As String = "lambda = "
    Dim firstExpected As Long, firstActual As Long, lastExpected As Long, lastActual As Long
    firstExpected = InStr(1, expected, marker, vbBinaryCompare)
    firstActual = InStr(1, actual, marker, vbBinaryCompare)
    If firstExpected = 0 Or firstActual = 0 Then Exit Function
    firstExpected = firstExpected + Len(marker): firstActual = firstActual + Len(marker)
    If InStr(firstExpected, expected, marker, vbBinaryCompare) > 0 Or _
        InStr(firstActual, actual, marker, vbBinaryCompare) > 0 Then Exit Function
    lastExpected = InStr(firstExpected, expected, ";", vbBinaryCompare)
    lastActual = InStr(firstActual, actual, ";", vbBinaryCompare)
    If lastExpected <= firstExpected Or lastActual <= firstActual Then Exit Function
    If Left$(expected, firstExpected - 1) & Mid$(expected, lastExpected) <> _
        Left$(actual, firstActual - 1) & Mid$(actual, lastActual) Then Exit Function
    Dim expectedNumber As String, actualNumber As String, reference As Double
    expectedNumber = Mid$(expected, firstExpected, lastExpected - firstExpected)
    actualNumber = Mid$(actual, firstActual, lastActual - firstActual)
    If Not IsNumeric(expectedNumber) Or Not IsNumeric(actualNumber) Then Exit Function
    reference = CDbl(expectedNumber)
    TestAudit03SameLimitComment = Abs(CDbl(actualNumber) - reference) <= 0.000000000001 * (1# + Abs(reference))
    Exit Function
Different:
    TestAudit03SameLimitComment = False
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' сравнение комментария допускает только округление lambda;
' изменение причины, другого числа, недопустимое число или заметное изменение lambda отклоняется.
Private Sub TestAudit03LimitCommentComparison(ByRef stats As TUiTestStats)
    Dim prefix As String, suffix As String, baseline As String
    prefix = "Предел достигнут, при lambda = ": suffix = "; точка трещинообразования не определяется."
    baseline = prefix & "115.570626695522" & suffix
    TestAssertTrue stats, "commentLambda.identical", TestAudit03SameLimitComment(baseline, baseline)
    TestAssertTrue stats, "commentLambda.rectRoundoff", TestAudit03SameLimitComment(baseline, prefix & "115.570626695521" & suffix)
    TestAssertTrue stats, "commentLambda.hollowRoundoff", TestAudit03SameLimitComment(prefix & "101.47179846739" & suffix, prefix & "101.471798467389" & suffix)
    TestAssertTrue stats, "commentLambda.exponent", TestAudit03SameLimitComment(prefix & "100" & suffix, prefix & "1E+02" & suffix)
    TestAssertTrue stats, "commentLambda.nearZero", TestAudit03SameLimitComment(prefix & "0" & suffix, prefix & "0.0000000000001" & suffix)
    TestAssertTrue stats, "commentLambda.signedZero", TestAudit03SameLimitComment(prefix & "0" & suffix, prefix & "-0" & suffix)
    TestAssertTrue stats, "commentLambda.changedNumberRejected", Not TestAudit03SameLimitComment(baseline, prefix & "115.570627" & suffix)
    TestAssertTrue stats, "commentLambda.smallRealChangeRejected", Not TestAudit03SameLimitComment(prefix & "1" & suffix, prefix & "1.00000001" & suffix)
    TestAssertTrue stats, "commentLambda.changedPrefixRejected", Not TestAudit03SameLimitComment(baseline, "Другой статус: " & baseline)
    TestAssertTrue stats, "commentLambda.changedSuffixRejected", Not TestAudit03SameLimitComment(baseline, prefix & "115.570626695521" & "; другая причина.")
    TestAssertTrue stats, "commentLambda.otherNumberRejected", Not TestAudit03SameLimitComment("Случай 1: " & baseline, "Случай 2: " & prefix & "115.570626695521" & suffix)
    TestAssertTrue stats, "commentLambda.plainTextExact", TestAudit03SameLimitComment("InputErr", "InputErr")
    TestAssertTrue stats, "commentLambda.statusRejected", Not TestAudit03SameLimitComment("InputErr", "NotCracked")
    TestAssertTrue stats, "commentLambda.missingMarkerRejected", Not TestAudit03SameLimitComment(baseline, "lambda отсутствует")
    TestAssertTrue stats, "commentLambda.missingEndRejected", Not TestAudit03SameLimitComment(baseline, prefix & "115.570626695521")
    TestAssertTrue stats, "commentLambda.invalidNumberRejected", Not TestAudit03SameLimitComment(baseline, prefix & "abc" & suffix)
    TestAssertTrue stats, "commentLambda.emptyNumberRejected", Not TestAudit03SameLimitComment(baseline, prefix & suffix)
    TestAssertTrue stats, "commentLambda.repeatedMarkerRejected", Not TestAudit03SameLimitComment(baseline & " lambda = 1;", prefix & "115.570626695521" & suffix & " lambda = 1;")
    TestAssertTrue stats, "commentLambda.overflowRejected", Not TestAudit03SameLimitComment(prefix & "1E+308" & suffix, prefix & "-1E+308" & suffix)
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' проверяет INPUT-площадь и численный порог на том же Config ->
' ModelSpace pipeline, который вызывает live import, но без подключения DWG.
Public Function RunAudit03InputAreaImportFilterTests() As String
    Dim stats As TUiTestStats
    TestAudit03InputAreaImportFilter stats
    TestAppendLine stats, "TOTAL_AUDIT03_INPUT_AREA_IMPORT: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunAudit03InputAreaImportFilterTests = stats.Report
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' mm2/cm2/m2 должны одинаково отфильтровать Region около порога.
' Саму DWG-геометрию не масштабируем. Нулевой порог, ошибочный ввод и потеря
' обязательного key проверяются отдельно; измененные ячейки восстанавливаются.
Private Sub TestAudit03InputAreaImportFilter(ByRef stats As TUiTestStats)
    Dim unitRange As Object, systemRange As Object, savedUnits As Variant, savedSystem As Variant
    Dim modelSpace As Collection, region As CTestFakeAcadRegion, importer As CAutoCADSectionModelImporter
    Dim section As CSectionModel, disconnected As CSectionModel
    Dim settings As CSystemSettingsReader, units As CUnitSystem
    Dim unitIndex As Long, caseIndex As Long, factor As Double, threshold As Double, rawValue As Double
    Dim row As Long, minAreaRow As Long, expectedConcrete As Long, expectedRebar As Long
    Dim unitNames As Variant, factors As Variant, areas As Variant, invalid As Variant, invalidIndex As Long
    Dim errorNumber As Long, description As String, prefix As String, solveCount As Long
    On Error GoTo Failed
    Set unitRange = ThisWorkbook.Names.Item("rngUnitSettings").RefersToRange
    Set systemRange = ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange
    savedUnits = unitRange.Formula: savedSystem = systemRange.Formula
    For row = 2 To systemRange.Rows.Count
        If CStr(systemRange.Cells(row, 1).Value2) = "AutoCAD.Import.MinArea" Then minAreaRow = row
    Next row
    If minAreaRow = 0 Then Err.Raise vbObjectError + 4602, "TestAudit03InputAreaImportFilter", "Не найден порог площади импорта."
    TestSetSystemSetting "Units.Length.Input", "m"
    TestSetSystemSetting "AutoCAD.Common.ConcreteLayer", "Concrete"
    TestSetSystemSetting "AutoCAD.Common.RebarLayer", "Reinf"
    Set modelSpace = New Collection
    Set region = New CTestFakeAcadRegion
    region.TestInitialize 10000#, 1000#, 2000#, 10000# * 10000# / 12#, 10000# * 10000# / 12#, 0#, "Concrete", "C_MAIN"
    modelSpace.Add region
    Set region = New CTestFakeAcadRegion
    region.TestInitialize 2#, 1005#, 2005#, 4# / 12#, 4# / 12#, 0#, "Concrete", "C_SMALL"
    modelSpace.Add region
    areas = Array(24.999, 25#, 25.001)
    For row = 0 To UBound(areas)
        Set region = New CTestFakeAcadRegion
        region.TestInitialize CDbl(areas(row)), 980# + row * 20#, 1980#, 1#, 1#, 0#, "Reinf", "R" & CStr(row + 1)
        modelSpace.Add region
    Next row
    Set importer = New CAutoCADSectionModelImporter
    unitNames = Array("mm2", "cm2", "m2")
    factors = Array(1#, 100#, 1000000#)
    solveCount = SectionEquilibriumSolveCount()
    For unitIndex = 0 To 2
        factor = CDbl(factors(unitIndex))
        TestSetSystemSetting "Units.Area.Input", CStr(unitNames(unitIndex))
        For caseIndex = 0 To 2
            Select Case caseIndex
                Case 0: threshold = 0#: expectedConcrete = 2: expectedRebar = 3
                Case 1: threshold = 25#: expectedConcrete = 1: expectedRebar = 2
                Case 2: threshold = 24.998: expectedConcrete = 1: expectedRebar = 3
            End Select
            rawValue = threshold / factor
            systemRange.Cells(minAreaRow, 2).Value2 = rawValue
            Set settings = New CSystemSettingsReader
            settings.LoadFromWorkbook ThisWorkbook
            Set units = New CUnitSystem
            units.LoadFromSettings settings
            Set section = importer.ImportConfiguredModelSpace(modelSpace, settings, units, TestAudit03AbsentContourLayers())
            prefix = "audit03.inputArea.unit" & CStr(unitIndex) & ".case" & CStr(caseIndex)
            TestAssertTrue stats, prefix & ".concreteCount", section.ConcreteCount = expectedConcrete
            TestAssertTrue stats, prefix & ".rebarCount", section.RebarCount = expectedRebar
            TestAssertClose stats, prefix & ".xMm", section.ConcreteX(1), 1000#, 0.000001
            TestAssertClose stats, prefix & ".yMm", section.ConcreteY(1), 2000#, 0.000001
            TestAssertClose stats, prefix & ".areaMm2", section.ConcreteArea(1), 10000#, 0.000001
            TestAssertClose stats, prefix & ".inertiaMm4", section.ConcreteLocalIx(1), 10000# * 10000# / 12#, 0.001
            If caseIndex = 1 Then
                TestAssertTrue stats, prefix & ".inclusiveBoundary", section.RebarSourceHandle(1) = "R2"
                TestAssertClose stats, prefix & ".boundaryArea", section.RebarArea(1), 25#, 0.000001
                TestAssertClose stats, prefix & ".boundaryX", section.RebarX(1), 1000#, 0.000001
                Set section = importer.ImportConfiguredModelSpace(modelSpace, settings, Nothing, TestAudit03AbsentContourLayers())
                TestAssertTrue stats, prefix & ".implicitUnitsRead", section.ConcreteCount = expectedConcrete And section.RebarCount = expectedRebar
                If unitIndex > 0 Then
                    Set disconnected = importer.ImportFromModelSpace(modelSpace, "Concrete", "Reinf", "Rebar", rawValue, units)
                    TestAssertTrue stats, prefix & ".detectDisconnectedUnits", disconnected.ConcreteCount <> section.ConcreteCount And _
                        disconnected.RebarCount <> section.RebarCount
                End If
            End If
            TestAppendLine stats, "AREA_FILTER: input=" & CStr(unitNames(unitIndex)) & "|raw=" & CStr(rawValue) & _
                "|thresholdMm2=" & CStr(threshold) & "|concrete=" & CStr(section.ConcreteCount) & "|rebar=" & CStr(section.RebarCount)
        Next caseIndex
    Next unitIndex
    TestSetSystemSetting "Units.Area.Input", "m2"
    invalid = Array(vbNullString, "TODO", "abc", CVErr(2042), -1#, "1e308")
    For invalidIndex = 0 To UBound(invalid)
        systemRange.Cells(minAreaRow, 2).Value2 = invalid(invalidIndex)
        TestAudit03CaptureImportError modelSpace, errorNumber, description
        prefix = "audit03.inputArea.invalid" & CStr(invalidIndex)
        TestAssertTrue stats, prefix & ".rejected", errorNumber <> 0
        TestAssertTrue stats, prefix & ".address", InStr(1, description, "AutoCAD.Import.MinArea", vbBinaryCompare) > 0
        TestAssertTrue stats, prefix & ".actualCell", InStr(1, description, systemRange.Cells(minAreaRow, 2).Address(False, False), vbTextCompare) > 0
        TestAssertTrue stats, prefix & ".action", InStr(1, description, "Введите", vbTextCompare) > 0 Or _
            InStr(1, description, "Исправьте", vbTextCompare) > 0 Or InStr(1, description, "Уменьшите", vbTextCompare) > 0
        TestAppendLine stats, "AREA_INVALID: case=" & CStr(invalidIndex) & "|error=" & CStr(errorNumber) & "|reason=" & description
    Next invalidIndex
    systemRange.Cells(minAreaRow, 2).Value2 = 25.002 / 1000000#
    TestAudit03CaptureImportError modelSpace, errorNumber, description
    TestAssertTrue stats, "audit03.inputArea.emptyRebar.rejected", errorNumber <> 0
    TestAssertTrue stats, "audit03.inputArea.emptyRebar.reason", InStr(1, description, "арматурных", vbBinaryCompare) > 0 And _
        InStr(1, description, "AutoCAD.Import.MinArea", vbBinaryCompare) > 0
    systemRange.Cells(minAreaRow, 2).Value2 = 10001# / 1000000#
    TestAudit03CaptureImportError modelSpace, errorNumber, description
    TestAssertTrue stats, "audit03.inputArea.emptyConcrete.rejected", errorNumber <> 0
    TestAssertTrue stats, "audit03.inputArea.emptyConcrete.reason", InStr(1, description, "бетонных", vbBinaryCompare) > 0 And _
        InStr(1, description, "AutoCAD.Import.MinArea", vbBinaryCompare) > 0
    systemRange.Cells(minAreaRow, 1).Value2 = "__MissingImportMinimumArea"
    TestAudit03CaptureImportError modelSpace, errorNumber, description
    TestAssertTrue stats, "audit03.inputArea.missing.rejected", errorNumber <> 0
    TestAssertTrue stats, "audit03.inputArea.missing.address", InStr(1, description, "AutoCAD.Import.MinArea", vbBinaryCompare) > 0
    TestAssertTrue stats, "audit03.inputArea.noStateSolve", SectionEquilibriumSolveCount() = solveCount
    GoTo Restore
Failed:
    TestAssertTrue stats, "audit03.inputArea.runtime." & CStr(Err.Number) & "." & Err.Description, False
Restore:
    On Error Resume Next
    If Not unitRange Is Nothing Then unitRange.Formula = savedUnits
    If Not systemRange Is Nothing Then systemRange.Formula = savedSystem
    On Error GoTo 0
    If Not unitRange Is Nothing Then
        If IsArray(savedUnits) Then
            invalid = unitRange.Formula
            TestAudit03ComparePlainSnapshot stats, "audit03.inputArea.restoreUnits", savedUnits, invalid, 1
        End If
    End If
    If Not systemRange Is Nothing Then
        If IsArray(savedSystem) Then
            invalid = systemRange.Formula
            TestAudit03ComparePlainSnapshot stats, "audit03.inputArea.restoreSettings", savedSystem, invalid, 1
        End If
    End If
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' собирает настройки заново после изменения ячейки и сохраняет
' только фактическую ошибку общего импорта. Ошибка не превращается в default.
Private Sub TestAudit03CaptureImportError(ByVal modelSpace As Object, ByRef errorNumber As Long, ByRef description As String)
    Dim settings As CSystemSettingsReader, units As CUnitSystem, importer As CAutoCADSectionModelImporter, section As CSectionModel
    errorNumber = 0: description = vbNullString
    On Error GoTo Failed
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem
    units.LoadFromSettings settings
    Set importer = New CAutoCADSectionModelImporter
    Set section = importer.ImportConfiguredModelSpace(modelSpace, settings, units, TestAudit03AbsentContourLayers())
    Exit Sub
Failed:
    errorNumber = Err.Number: description = Err.Description
    Err.Clear
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' воспроизводит импорт -> смена INPUT/OUTPUT -> кнопка расчета.
' Использует общий импорт существующих fake Region и реальный snapshot/pipeline.
Public Function RunAudit03ImportedSnapshotUnitChangeTests() As String
    Dim stats As TUiTestStats
    TestAudit03ImportedSnapshotUnitChanges stats
    TestAppendLine stats, "TOTAL_AUDIT03_IMPORTED_UNIT_CHANGES: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunAudit03ImportedSnapshotUnitChangeTests = stats.Report
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' все девять переходов OUTPUT-длины mm/cm/m должны сохранить
' импортное сечение. INPUT и знаки меняются вместе с независимым переводом
' чисел; заголовки старого snapshot остаются источником его размерностей.
' Непереведенное число нагрузки отдельно должно получить новый физический смысл.
Private Sub TestAudit03ImportedSnapshotUnitChanges(ByRef stats As TUiTestStats)
    Dim rangeNames As Variant, original As Collection, name As Variant, i As Long, beforeIndex As Long, afterIndex As Long
    Dim lengthNames As Variant, areaNames As Variant, lengthFactors As Variant, areaFactors As Variant
    Dim forceNames As Variant, forceFactors As Variant, momentNames As Variant, momentFactors As Variant
    Dim stressNames As Variant, stressFactors As Variant, lengthFactor As Double, stressFactor As Double
    Dim forceFactor As Double, momentFactor As Double, signN As Double, signMx As Double, signMy As Double
    Dim settings As CSystemSettingsReader, units As CUnitSystem, section As CSectionModel, restored As CSectionModel
    Dim writer As CNDMResultsWriter, loads As Object, preview As Variant, actual As Variant
    Dim baselineElements As Variant, baselineGeometry As Variant, baselineProperties As Variant
    Dim prefix As String, message As String, solveCount As Long, errorNumber As Long, description As String
    Dim smallStrain As Double, largeStrain As Double
    On Error GoTo Failed
    rangeNames = Array("rngSystemSettings", "rngUnitSettings", "rngSignConventionSettings", _
        "rngCalculationProfiles", "rngLoadCombinations", "rngConcreteMaterialParameters", "rngSteelMaterialParameters", _
        "rngCircleGeometry", "rngRectanglesGeometry", "rngPointedSectionGeometry", "rngHollowRectangleGeometry")
    Set original = New Collection
    For Each name In rangeNames
        original.Add ThisWorkbook.Names.Item(CStr(name)).RefersToRange.Formula
    Next name
    lengthNames = Array("mm", "cm", "m"): lengthFactors = Array(1#, 10#, 1000#)
    areaNames = Array("mm2", "cm2", "m2"): areaFactors = Array(1#, 100#, 1000000#)
    forceNames = Array("N", "kN", "tf"): forceFactors = Array(1#, 1000#, 9806.65)
    momentNames = Array("N*mm", "kN*m", "tf*m"): momentFactors = Array(1#, 1000000#, 9806650#)
    stressNames = Array("MPa", "kPa", "Pa"): stressFactors = Array(1#, 0.001, 0.000001)
    Set writer = New CNDMResultsWriter
    For beforeIndex = 0 To 2
        For afterIndex = 0 To 2
            For i = 0 To UBound(rangeNames)
                ThisWorkbook.Names.Item(CStr(rangeNames(i))).RefersToRange.Formula = original(i + 1)
            Next i
            TestPrepareCircleInput
            TestSetSystemSetting "General.ExecutionReportEnabled", "No"
            TestSetSystemSetting "Plot.AutoUpdateAfterCalculation", "No"
            TestSetSystemSetting "Calculation.ZeroMomentPerDepth", "0"
            TestSetSystemSetting "Load.ReferenceOffsetX", "0": TestSetSystemSetting "Load.ReferenceOffsetY", "0"
            TestSetSystemSetting "Solver.ToleranceN", "0.001"
            TestSetSystemSetting "Solver.ToleranceMx", "0.001": TestSetSystemSetting "Solver.ToleranceMy", "0.001"
            TestSetProfileSetting "PR1", "Calculation.Strength.DirectState", "Yes"
            TestSetProfileSetting "PR1", "Calculation.Strength.Capacity", "No"
            TestSetProfileSetting "PR1", "Calculation.Crack.Width", "No"
            TestSetProfileSetting "PR1", "Calculation.Stability.Enabled", "No"
            TestSetProfileSetting "PR1", "Visualization.State", "StrengthState"
            TestSetSystemSetting "AutoCAD.Import.MinArea", "0.000001"
            TestSetSystemSetting "AutoCAD.Common.ConcreteLayer", "Concrete"
            TestSetSystemSetting "AutoCAD.Common.RebarLayer", "Reinf"
            For Each name In Array("Force", "Moment", "Stress", "Curvature")
                TestSetSystemSetting "Units." & CStr(name) & ".Output", TestGetSystemSetting("Units." & CStr(name) & ".Input")
            Next name
            TestSetSystemSetting "Units.Length.Output", CStr(lengthNames(beforeIndex))
            TestSetSystemSetting "Units.Area.Output", CStr(areaNames(beforeIndex))
            Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
            Set units = New CUnitSystem: units.LoadFromSettings settings
            Set section = TestAudit03ImportSnapshotFixture(settings, units)
            writer.WriteGeometryPreview ThisWorkbook, section, PrepareSectionSnapshot(section), units
            preview = TestResultTable("rngNDMSectionGeometry")
            prefix = "audit03.importUnits.before" & CStr(beforeIndex) & ".after" & CStr(afterIndex)
            TestAssertClose stats, prefix & ".previewX", TestGeometryResultValue(preview, "C1", "X"), 950# / CDbl(lengthFactors(beforeIndex)), 0.000001
            TestAssertClose stats, prefix & ".previewArea", TestGeometryResultValue(preview, "C1", "Area"), 10000# / CDbl(areaFactors(beforeIndex)), 0.000001
            TestAssertTextEquals stats, prefix & ".previewHeader", CStr(preview(1, TestResultHeaderColumnByBaseName(preview, "X"))), "X, " & CStr(lengthNames(beforeIndex))
            TestAssertTextEquals stats, prefix & ".previewAreaHeader", CStr(preview(1, TestResultHeaderColumnByBaseName(preview, "Area"))), "Area, " & CStr(areaNames(beforeIndex))
            TestAudit03AssertAutoCADMmGeometry stats, prefix & ".beforeChange", section

            lengthFactor = CDbl(lengthFactors(afterIndex)): stressFactor = CDbl(stressFactors(afterIndex)) ' Меняем текущие единицы, сохраняя прежнюю импортированную геометрию и её единицы.
            forceFactor = CDbl(forceFactors(afterIndex)): momentFactor = CDbl(momentFactors(afterIndex))
            TestAudit03RescaleInputTables lengthFactor, stressFactor
            TestSetSystemSetting "Calculation.ZeroMomentPerDepth", "0"
            TestSetSystemSetting "Units.Length.Input", CStr(lengthNames(afterIndex))
            TestSetSystemSetting "Units.Area.Input", CStr(areaNames(afterIndex))
            TestSetSystemSetting "Units.Force.Input", CStr(forceNames(afterIndex))
            TestSetSystemSetting "Units.Moment.Input", CStr(momentNames(afterIndex))
            TestSetSystemSetting "Units.Stress.Input", CStr(stressNames(afterIndex))
            TestSetSystemSetting "Units.Curvature.Input", IIf(afterIndex = 0, "1/mm", "1/m")
            TestSetSystemSetting "Units.Length.Output", CStr(lengthNames(afterIndex))
            TestSetSystemSetting "Units.Area.Output", CStr(areaNames(afterIndex))
            TestSetSystemSetting "Units.Force.Output", CStr(forceNames(afterIndex))
            TestSetSystemSetting "Units.Moment.Output", CStr(momentNames(afterIndex))
            TestSetSystemSetting "Units.Stress.Output", CStr(stressNames(afterIndex))
            TestSetSystemSetting "Units.Curvature.Output", IIf(afterIndex = 0, "1/mm", "1/m")
            signN = IIf(afterIndex = 1, -1#, 1#): signMx = IIf(afterIndex = 1, -1#, 1#): signMy = IIf(afterIndex = 2, -1#, 1#)
            TestSetSystemSetting "Sign.N.User", IIf(signN > 0#, "Tension", "Compression")
            TestSetSystemSetting "Sign.Mx.User", IIf(signMx > 0#, "+Y tension", "-Y tension")
            TestSetSystemSetting "Sign.My.User", IIf(signMy > 0#, "+X tension", "-X tension")
            TestSetSystemSetting "Solver.ToleranceN", CStr(0.001 / forceFactor)
            TestSetSystemSetting "Solver.ToleranceMx", CStr(0.001 / momentFactor)
            TestSetSystemSetting "Solver.ToleranceMy", CStr(0.001 / momentFactor)
            TestSetSystemSetting "Geometry.Source", "AutoCAD"
            ' Эти параметры остановили бы новый импорт, но сохраненную модель
            ' при расчете не должны ни перечитывать из DWG, ни фильтровать заново.
            TestSetSystemSetting "AutoCAD.Import.MinArea", "1e100"
            TestSetSystemSetting "AutoCAD.Common.ConcreteLayer", "__NoNewConcreteImport"
            TestSetSystemSetting "AutoCAD.Common.RebarLayer", "__NoNewRebarImport"
            Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
            Set units = New CUnitSystem: units.LoadFromSettings settings
            solveCount = SectionEquilibriumSolveCount()
            Set restored = BuildWorkbookSectionModel(ThisWorkbook, settings, units)
            TestAssertTextEquals stats, prefix & ".source", restored.SourceType, "AutoCADImport"
            TestAssertClose stats, prefix & ".restoredX", restored.ConcreteX(1), 950#, 0.000001
            TestAssertClose stats, prefix & ".restoredArea", restored.ConcreteArea(1), 10000#, 0.000001
            TestAssertTrue stats, prefix & ".restoreNoSolve", SectionEquilibriumSolveCount() = solveCount
            actual = TestResultTable("rngNDMSectionGeometry")
            TestAudit03ComparePlainSnapshot stats, prefix & ".snapshotUnchangedBeforeSolve", preview, actual, 1
            TestAudit03AssertAutoCADMmGeometry stats, prefix & ".afterChangeBeforeSolve", section

            Set loads = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange
            TestClearDataRows loads
            loads.Cells(2, 1).Value2 = "LC_IMPORT"
            loads.Cells(2, 2).Value2 = -100000# / forceFactor * signN
            loads.Cells(2, 3).Value2 = 1000000# / momentFactor * signMx
            loads.Cells(2, 4).Value2 = -750000# / momentFactor * signMy
            loads.Cells(2, 5).Value2 = "PR1": loads.Cells(2, 6).Value2 = ChrW$(&H3BB) & "*Mxy"
            message = TestFixtureRunSectionWorkflow(ThisWorkbook, False)
            TestAppendLine stats, "IMPORTED_CALCULATION_MESSAGE: " & prefix & "|" & message
            TestAssertTrue stats, prefix & ".run", InStr(1, message, "Расчет завершен. Обработано сочетаний: 1.", vbTextCompare) = 1
            TestAssertTrue stats, prefix & ".noImportNotice", InStr(1, message, "импорт", vbTextCompare) = 0 And InStr(1, message, "AutoCAD", vbTextCompare) = 0 ' Сообщения импорта допустимы только при формировании сечения.
            TestAssertTextEquals stats, prefix & ".directState", TestResultsPropertyValue("LC_IMPORT", "DirectStateStatus"), "OK"
            TestAssertClose stats, prefix & ".equilibriumN", CDbl(TestResultsPropertyValue("LC_IMPORT", "Nint")) * forceFactor * signN, -100000#, 0.001
            TestAssertTextEquals stats, prefix & ".savedSource", ResultsGeometrySource(ThisWorkbook), "AutoCADImport"
            TestAssertTrue stats, prefix & ".hasState", TestResultTableRowCount("rngNDMElementResults") > 1
            TestAssertClose stats, prefix & ".loadN", CDbl(TestResultsPropertyValue("LC_IMPORT", "N")) * forceFactor * signN, -100000#, 0.001
            TestAudit03AssertAutoCADMmGeometry stats, prefix & ".afterSolve", section
            actual = TestResultTable("rngNDMSectionGeometry")
            TestAssertTextEquals stats, prefix & ".savedHeader", CStr(actual(1, TestResultHeaderColumnByBaseName(actual, "X"))), "X, " & CStr(lengthNames(beforeIndex))
            TestAssertTextEquals stats, prefix & ".savedAreaHeader", CStr(actual(1, TestResultHeaderColumnByBaseName(actual, "Area"))), "Area, " & CStr(areaNames(beforeIndex))
            TestAssertClose stats, prefix & ".savedX", TestGeometryResultValue(actual, "C1", "X"), 950# / CDbl(lengthFactors(beforeIndex)), 0.000001
            TestAudit03ComparePlainSnapshot stats, prefix & ".geometryPreservedAfterSolve", preview, actual, 2
            actual = TestResultTable("rngNDMElementResults")
            If beforeIndex = 0 And afterIndex = 0 Then
                baselineElements = actual
                baselineGeometry = TestResultTable("rngNDMSectionGeometry")
                baselineProperties = TestResultTable("rngNDMSectionProperties")
            Else
                TestAudit03CompareUnitSnapshot stats, prefix & ".elements", baselineElements, actual, "Elements", _
                    lengthFactor, CDbl(areaFactors(afterIndex)), stressFactor, IIf(afterIndex = 0, 1#, 0.001), forceFactor, momentFactor, signN, signMx, signMy
                actual = TestResultTable("rngNDMSectionGeometry")
                TestAudit03CompareUnitSnapshot stats, prefix & ".geometry", baselineGeometry, actual, "Geometry", _
                    CDbl(lengthFactors(beforeIndex)), CDbl(areaFactors(beforeIndex)), stressFactor, 1#, forceFactor, momentFactor, signN, signMx, signMy
                actual = TestResultTable("rngNDMSectionProperties")
                TestAudit03CompareUnitSnapshot stats, prefix & ".properties", baselineProperties, actual, "Properties", _
                    lengthFactor, CDbl(areaFactors(afterIndex)), stressFactor, IIf(afterIndex = 0, 1#, 0.001), forceFactor, momentFactor, signN, signMx, signMy, True
            End If
            TestAppendLine stats, "IMPORTED_UNIT_CHANGE: before=" & CStr(lengthNames(beforeIndex)) & "|after=" & CStr(lengthNames(afterIndex)) & _
                "|inputForce=" & CStr(forceNames(afterIndex)) & "|inputStress=" & CStr(stressNames(afterIndex)) & "|source=" & ResultsGeometrySource(ThisWorkbook)
        Next afterIndex
    Next beforeIndex

    ' Смена INPUT без изменения числа не сохраняет физическую нагрузку.
    ' Геометрия при этом остается той же, обе осевые задачи здесь линейны.
    loads.Cells(2, 2).Value2 = -100#: loads.Cells(2, 3).Value2 = 0#: loads.Cells(2, 4).Value2 = 0#
    TestSetSystemSetting "Units.Force.Input", "N"
    TestSetSystemSetting "Solver.ToleranceN", "0.001"
    message = TestFixtureRunSectionWorkflow(ThisWorkbook, False)
    TestAssertClose stats, "audit03.importUnits.rawUnchanged.oldN", CDbl(TestResultsPropertyValue("LC_IMPORT", "N")) * 9806.65, -100#, 0.001
    smallStrain = CDbl(TestResultsPropertyValue("LC_IMPORT", "State.StrengthState.Epsilon0"))
    TestAppendLine stats, "IMPORTED_RAW_LOAD: stage=N; epsilon0=" & CStr(smallStrain) & _
        "; kappaX=" & TestResultsPropertyValue("LC_IMPORT", "State.StrengthState.KappaX") & _
        "; kappaY=" & TestResultsPropertyValue("LC_IMPORT", "State.StrengthState.KappaY") & _
        "; Nint=" & TestResultsPropertyValue("LC_IMPORT", "Nint")
    TestSetSystemSetting "Units.Force.Input", "kN"
    TestSetSystemSetting "Solver.ToleranceN", "0.000001"
    message = TestFixtureRunSectionWorkflow(ThisWorkbook, False)
    TestAssertClose stats, "audit03.importUnits.rawUnchanged.number", CDbl(loads.Cells(2, 2).Value2), -100#, 0#
    TestAssertClose stats, "audit03.importUnits.rawUnchanged.newN", CDbl(TestResultsPropertyValue("LC_IMPORT", "N")) * 9806.65, -100000#, 0.001
    TestAssertClose stats, "audit03.importUnits.rawUnchanged.equilibrium", CDbl(TestResultsPropertyValue("LC_IMPORT", "Nint")) * 9806.65, -100000#, 0.001
    largeStrain = CDbl(TestResultsPropertyValue("LC_IMPORT", "State.StrengthState.Epsilon0"))
    TestAppendLine stats, "IMPORTED_RAW_LOAD: stage=kN; epsilon0=" & CStr(largeStrain) & _
        "; kappaX=" & TestResultsPropertyValue("LC_IMPORT", "State.StrengthState.KappaX") & _
        "; kappaY=" & TestResultsPropertyValue("LC_IMPORT", "State.StrengthState.KappaY") & _
        "; Nint=" & TestResultsPropertyValue("LC_IMPORT", "Nint")
    TestAssertTrue stats, "audit03.importUnits.rawUnchanged.compression", smallStrain < 0# And largeStrain < 0#
    TestAssertClose stats, "audit03.importUnits.rawUnchanged.strainRatio", largeStrain / smallStrain, 1000#, 0.000001
    TestAudit03AssertAutoCADMmGeometry stats, "audit03.importUnits.rawUnchanged.geometry", section

    preview = TestResultTable("rngNDMSectionGeometry")
    solveCount = SectionEquilibriumSolveCount()
    TestSetSystemSetting "Units.Length.Output", "ft"
    On Error Resume Next
    message = TestFixtureRunSectionWorkflow(ThisWorkbook, False)
    errorNumber = Err.Number: description = Err.Description
    Err.Clear
    On Error GoTo Failed
    TestAssertTrue stats, "audit03.importUnits.invalidOutput.rejected", errorNumber <> 0
    TestAssertTrue stats, "audit03.importUnits.invalidOutput.address", InStr(1, description, "Units.Length.Output", vbBinaryCompare) > 0
    TestAssertTrue stats, "audit03.importUnits.invalidOutput.noSolve", SectionEquilibriumSolveCount() = solveCount
    actual = TestResultTable("rngNDMSectionGeometry")
    TestAudit03ComparePlainSnapshot stats, "audit03.importUnits.invalidOutput.snapshotPreserved", preview, actual, 1
    TestSetSystemSetting "Units.Length.Output", "m"
    message = TestFixtureRunSectionWorkflow(ThisWorkbook, False)
    TestAppendLine stats, "IMPORTED_CALCULATION_MESSAGE: invalidOutput.recovery|" & message
    TestAssertTrue stats, "audit03.importUnits.invalidOutput.recovery", InStr(1, message, "Расчет завершен. Обработано сочетаний: 1.", vbTextCompare) = 1
    TestAssertTrue stats, "audit03.importUnits.invalidOutput.noImportNotice", InStr(1, message, "импорт", vbTextCompare) = 0 And InStr(1, message, "AutoCAD", vbTextCompare) = 0 ' Повторный расчёт не повторяет сообщения этапа импорта.
    TestAssertTextEquals stats, "audit03.importUnits.invalidOutput.directState", TestResultsPropertyValue("LC_IMPORT", "DirectStateStatus"), "OK"
    TestAssertClose stats, "audit03.importUnits.invalidOutput.equilibriumN", CDbl(TestResultsPropertyValue("LC_IMPORT", "Nint")) * 9806.65, -100000#, 0.001 ' Здесь сохраняются выходные tf и положительное растяжение из последнего набора.
    TestAssertTextEquals stats, "audit03.importUnits.invalidOutput.savedSource", ResultsGeometrySource(ThisWorkbook), "AutoCADImport"
    TestAudit03AssertAutoCADMmGeometry stats, "audit03.importUnits.invalidOutput.recoveredGeometry", section
    GoTo Restore
Failed:
    TestAssertTrue stats, "audit03.importUnits.runtime." & CStr(Err.Number) & "." & Err.Description, False
Restore:
    On Error Resume Next
    If Not original Is Nothing Then
        For i = 0 To original.Count - 1
            ThisWorkbook.Names.Item(CStr(rangeNames(i))).RefersToRange.Formula = original(i + 1)
        Next i
    End If
    On Error GoTo 0
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' четыре квадратных бетонных Region и четыре арматурных Region
' проходят общий importer с Config. Смещенный центр проверяет перенос момента.
Private Function TestAudit03ImportSnapshotFixture(ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem) As CSectionModel
    Dim modelSpace As Collection, region As CTestFakeAcadRegion, importer As CAutoCADSectionModelImporter
    Dim x As Double, y As Double, i As Long
    Set modelSpace = New Collection
    For i = 0 To 3
        x = 950# + (i Mod 2) * 100#: y = -750# + (i \ 2) * 100#
        Set region = New CTestFakeAcadRegion
        region.TestInitialize 10000#, x, y, 10000# * 10000# / 12#, 10000# * 10000# / 12#, 0#, "Concrete", "C" & CStr(i + 1)
        modelSpace.Add region
        x = 930# + (i Mod 2) * 140#: y = -770# + (i \ 2) * 140#
        Set region = New CTestFakeAcadRegion
        region.TestInitialize GEOM_PI * 20# ^ 2 / 4#, x, y, 1#, 1#, 0#, "Reinf", "R" & CStr(i + 1)
        modelSpace.Add region
    Next i
    Set importer = New CAutoCADSectionModelImporter
    Set TestAudit03ImportSnapshotFixture = importer.ImportConfiguredModelSpace(modelSpace, settings, units, TestAudit03AbsentContourLayers())
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' ====================== ДЛЯ ТЕСТОВ: ПОНЯТНЫЕ ОШИБКИ CONFIG ======================

' -------------------- ДЛЯ ТЕСТОВ --------------------
' отдельный gate адресов, причин и действий в диагностике ввода.
' Полный UI suite запускает те же проверки, а не отдельную упрощенную реализацию.
Public Function RunAudit03SettingErrorMessageTests() As String
    Dim stats As TUiTestStats
    TestAudit03SettingErrorMessages stats
    TestAudit03RequiredTableMessages stats
    TestAudit03ConfigConversionMessages stats
    TestAudit03UnitSignChoices stats
    TestAppendLine stats, "TOTAL_AUDIT03_SETTING_ERROR_MESSAGES: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunAudit03SettingErrorMessageTests = stats.Report
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' обязательные строки берутся из фактического Config. Материальные
' alias-ключи отдельно проверяются по сжатой/растянутой ячейке. После мутаций
' восстанавливаются формулы; проверка не заменяет производственные defaults.
Private Sub TestAudit03SettingErrorMessages(ByRef stats As TUiTestStats)
    Dim systemRange As Object, concreteRange As Object, steelRange As Object, loads As Object
    Dim savedSystem As Variant, savedConcrete As Variant, savedSteel As Variant, savedLoads As Variant
    Dim settings As CSystemSettingsReader, cell As Object, keyCell As Object, target As Object
    Dim keys As Variant, entry As Variant, invalid As Variant, key As String, row As Long, i As Long
    Dim numericValue As Boolean, oldValue As Variant, description As String, code As Long, prefix As String
    Dim result As Double, text As String, concrete As CConcreteMaterialParameters, steel As CSteelMaterialParameters
    On Error GoTo Failed
    Set systemRange = ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange
    Set concreteRange = ThisWorkbook.Names.Item("rngConcreteMaterialParameters").RefersToRange
    Set steelRange = ThisWorkbook.Names.Item("rngSteelMaterialParameters").RefersToRange
    Set loads = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange
    savedSystem = systemRange.Formula: savedConcrete = concreteRange.Formula
    savedSteel = steelRange.Formula: savedLoads = loads.Formula
    keys = Array("SLS.Crack.Allowable", "SLS.Crack.Phi1", "SLS.Crack.Phi2", "SLS.Crack.Phi3", "SLS.Crack.PsiS", _
        "SLS.Crack.Phi3Mode", "SLS.Crack.PsiMode", "SLS.Crack.SigmaSCrcAveragingMode", "SLS.Crack.TensionZoneMode", _
        "SLS.Crack.CoverDistanceMode", "SLS.Crack.InitiationSolutionStrategy", _
        "Stability.Code", "Stability.ElementLength", "Stability.Mu1", "Stability.Mu2", "Stability.SystemType", _
        "Stability.ZeroMomentEccentricitySign1", "Stability.ZeroMomentEccentricitySign2", "Stability.PhiLMode", _
        "Stability.AccidentalEccentricityMode", "Stability.AccidentalEccentricityPlanes", _
        "Stability.AccidentalEccentricityUser1", "Stability.AccidentalEccentricityUser2", "Stability.SP63.Ks", _
        "Stability.SP63.DeltaEMin", "Stability.SP63.DeltaEMax", "Stability.SP35.PhiP", "Stability.SP35.NOverNcrLimit", _
        "AutoCAD.Import.MinArea", "AutoCAD.Common.ConcreteLayer", "AutoCAD.Common.RebarLayer")
    For i = 0 To UBound(keys)
        key = CStr(keys(i))
        Set keyCell = Nothing
        For row = 2 To systemRange.Rows.Count
            If CStr(systemRange.Cells(row, 1).Value2) = key Then Set keyCell = systemRange.Cells(row, 1): Exit For
        Next row
        If keyCell Is Nothing Then Err.Raise 5, , "Не найдена обязательная строка " & key
        Set cell = keyCell.Offset(0, 1): oldValue = cell.Formula
        numericValue = IsNumeric(cell.Value2)
        For Each invalid In Array("", "TODO", "abc", CVErr(xlErrValue))
            text = vbNullString
            If Not IsError(invalid) Then text = CStr(invalid)
            If numericValue Or IsError(invalid) Or text <> "abc" Then
                cell.Value2 = invalid
                On Error Resume Next
                Err.Clear
                Set settings = New CSystemSettingsReader: settings.LoadFromRange systemRange
                If Err.Number = 0 Then
                    If numericValue Then result = settings.GetRequiredDouble(key) Else text = settings.GetRequiredString(key)
                End If
                code = Err.Number: description = Err.Description
                On Error GoTo Failed
                prefix = "audit03.settingMessage.system." & key & "." & CStr(VarType(invalid))
                TestAssertTrue stats, prefix & ".error", code <> 0
                TestAudit03AssertSettingMessage stats, prefix, description, key, cell
            End If
        Next invalid
        cell.Formula = oldValue
        keyCell.Value2 = key & ".REMOVED"
        On Error Resume Next
        Err.Clear
        Set settings = New CSystemSettingsReader: settings.LoadFromRange systemRange
        If numericValue Then result = settings.GetRequiredDouble(key) Else text = settings.GetRequiredString(key)
        code = Err.Number: description = Err.Description
        On Error GoTo Failed
        TestAssertTrue stats, "audit03.settingMessage.missing." & key & ".error", code <> 0
        TestAudit03AssertSettingMessage stats, "audit03.settingMessage.missing." & key, description, key, Nothing
        TestAssertTrue stats, "audit03.settingMessage.missing." & key & ".restore", InStr(1, description, "Восстановите строку", vbTextCompare) > 0
        keyCell.Value2 = key
    Next i

    keys = Array(Array("Concrete.Rb.ULS", "Concrete.R.ULS(I)", 2), Array("Concrete.Rbt.ULS", "Concrete.R.ULS(I)", 3), _
        Array("Concrete.Rb.SLS", "Concrete.R.SLS(II)", 2), Array("Concrete.Rbt.SLS", "Concrete.R.SLS(II)", 3), _
        Array("Concrete.Rb.mc2", "Concrete.Rb.mc2", 2), Array("Concrete.Eb", "Concrete.E", 2), Array("Concrete.Ebt", "Concrete.E", 3), _
        Array("Concrete.Eb1Red", "Concrete.TwoLine.Eb1Red", 2), Array("Concrete.Ebt1Red", "Concrete.TwoLine.Eb1Red", 3), _
        Array("Concrete.Eb0", "Concrete.ThreeLine.Eb0", 2), Array("Concrete.Ebt0", "Concrete.ThreeLine.Eb0", 3), _
        Array("Concrete.Eb2", "Concrete.TwoThreeLine.Eb2", 2), Array("Concrete.Ebt2", "Concrete.TwoThreeLine.Eb2", 3), _
        Array("Steel.Rsc.ULS", "Steel.R.ULS(I)", 2), Array("Steel.Rs.ULS", "Steel.R.ULS(I)", 3), _
        Array("Steel.Rsc.SLS", "Steel.R.SLS(II)", 2), Array("Steel.Rs.SLS", "Steel.R.SLS(II)", 3), _
        Array("Steel.Esc", "Steel.E", 2), Array("Steel.Es", "Steel.E", 3), _
        Array("Steel.TwoLine.Esc2", "Steel.TwoLine.Es2", 2), Array("Steel.TwoLine.Es2", "Steel.TwoLine.Es2", 3), _
        Array("Steel.ThreeLine.Esc2", "Steel.ThreeLine.Es2", 2), Array("Steel.ThreeLine.Es2", "Steel.ThreeLine.Es2", 3)) ' Входной alias материала должен указывать на правильную сторону диаграммы.
    For Each entry In keys
        key = CStr(entry(0))
        If Left$(key, 9) = "Concrete." Then Set target = concreteRange Else Set target = steelRange
        Set cell = Nothing
        For row = 2 To target.Rows.Count
            If CStr(target.Cells(row, 1).Value2) = CStr(entry(1)) Then Set cell = target.Cells(row, CLng(entry(2))): Exit For
        Next row
        If cell Is Nothing Then Err.Raise 5, , "Не найдена материальная строка " & CStr(entry(1))
        oldValue = cell.Formula
        For Each invalid In Array("", "TODO", "abc", 0#, CVErr(xlErrValue))
            cell.Value2 = invalid
            On Error Resume Next
            Err.Clear
            Set settings = New CSystemSettingsReader: settings.LoadFromRange target
            If Err.Number = 0 Then
                If Left$(key, 9) = "Concrete." Then
                    Set concrete = New CConcreteMaterialParameters: concrete.LoadFromSettings settings
                Else
                    Set steel = New CSteelMaterialParameters: steel.LoadFromSettings settings
                End If
            End If
            code = Err.Number: description = Err.Description
            On Error GoTo Failed
            prefix = "audit03.settingMessage.material." & key & "." & CStr(VarType(invalid))
            TestAssertTrue stats, prefix & ".error", code <> 0
            ' Excel-error возникает до разворачивания alias, поэтому там видимое имя строки.
            If IsError(invalid) Then text = CStr(entry(1)) Else text = key
            TestAudit03AssertSettingMessage stats, prefix, description, text, cell
            cell.Formula = oldValue
        Next invalid
    Next entry

    TestSetSystemSetting "Geometry.Source", "Generated" ' Проверяем не только исключение reader-а, но и сообщение обычного workbook-сценария.
    TestSetSystemSetting "Geometry.Type", "Circle"
    TestSetSystemSetting "General.ExecutionReportEnabled", "No"
    TestSetSystemSetting "SLS.Crack.Allowable", ""
    loads.Offset(1, 0).Resize(loads.Rows.Count - 1, loads.Columns.Count).ClearContents
    loads.Cells(2, 1).Value2 = "INPUT_MESSAGE": loads.Cells(2, 2).Value2 = 1#
    loads.Cells(2, 3).Value2 = 0#: loads.Cells(2, 4).Value2 = 0#: loads.Cells(2, 5).Value2 = "PR2"
    text = TestFixtureRunSectionWorkflow(ThisWorkbook)
    For row = 2 To systemRange.Rows.Count
        If CStr(systemRange.Cells(row, 1).Value2) = "SLS.Crack.Allowable" Then Set cell = systemRange.Cells(row, 2): Exit For
    Next row
    TestAudit03AssertSettingMessage stats, "audit03.settingMessage.workbook.message", text, "SLS.Crack.Allowable", cell
    description = CStr(ThisWorkbook.Names.Item("rngBatchSummary").RefersToRange.Offset(12, 2).Value2)
    TestAudit03AssertSettingMessage stats, "audit03.settingMessage.workbook.results", description, "SLS.Crack.Allowable", cell
    GoTo Restore
Failed:
    stats.Failed = stats.Failed + 1
    TestAppendLine stats, "FAIL: audit03.settingMessage.runtime; " & CStr(Err.Number) & "; " & Err.Description
Restore:
    On Error Resume Next
    If Not systemRange Is Nothing Then systemRange.Formula = savedSystem
    If Not concreteRange Is Nothing Then concreteRange.Formula = savedConcrete
    If Not steelRange Is Nothing Then steelRange.Formula = savedSteel
    If Not loads Is Nothing Then loads.Formula = savedLoads
    On Error GoTo 0
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' проверяет четыре независимых компонента видимой диагностики.
' Сообщение пишется в evidence без назначения статуса по его тексту.
Private Sub TestAudit03AssertSettingMessage(ByRef stats As TUiTestStats, ByVal prefix As String, _
        ByVal description As String, ByVal key As String, ByVal cell As Object)
    TestAssertTrue stats, prefix & ".key", InStr(1, description, key, vbTextCompare) > 0
    TestAssertTrue stats, prefix & ".sheet", InStr(1, description, "Config", vbTextCompare) > 0
    If Not cell Is Nothing Then TestAssertTrue stats, prefix & ".address", InStr(1, description, cell.Address(False, False), vbTextCompare) > 0
    TestAssertTrue stats, prefix & ".action", InStr(1, description, "Введите", vbTextCompare) > 0 Or _
        InStr(1, description, "Заполните", vbTextCompare) > 0 Or InStr(1, description, "Исправьте", vbTextCompare) > 0
    TestAppendLine stats, "SETTING_MESSAGE: " & prefix & "|" & description
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' повреждает только ссылку обязательной таблицы и проверяет
' понятное действие вместо технической COM-ошибки. Имя и исходная ссылка
' восстанавливаются даже при исключении; пользовательские ячейки не меняются.
Private Sub TestAudit03RequiredTableMessages(ByRef stats As TUiTestStats)
    Dim rangeName As Variant, name As Object, savedReference As String
    Dim settings As CSystemSettingsReader, code As Long, description As String, prefix As String
    On Error GoTo Failed
    For Each rangeName In Array("rngUnitSettings", "rngSignConventionSettings", _
            "rngSystemSettings", "rngPlotAnnotationSettings")
        Set name = ThisWorkbook.Names.Item(CStr(rangeName))
        savedReference = name.RefersTo
        name.RefersTo = "=#REF!"
        On Error Resume Next
        Err.Clear
        Set settings = New CSystemSettingsReader
        settings.LoadFromWorkbook ThisWorkbook
        code = Err.Number: description = Err.Description
        On Error GoTo Failed
        name.RefersTo = savedReference
        savedReference = vbNullString
        prefix = "audit03.settingMessage.requiredTable." & CStr(rangeName)
        TestAssertTrue stats, prefix & ".code", code = vbObjectError + 4316
        TestAssertTrue stats, prefix & ".sheet", InStr(1, description, "Config", vbTextCompare) > 0
        TestAssertTrue stats, prefix & ".name", InStr(1, description, CStr(rangeName), vbTextCompare) > 0
        TestAssertTrue stats, prefix & ".action", InStr(1, description, "Восстановите таблицу", vbTextCompare) > 0 And _
            InStr(1, description, "диспетчере имен", vbTextCompare) > 0
        TestAppendLine stats, "SETTING_MESSAGE: " & prefix & "|" & description
        Set settings = New CSystemSettingsReader
        settings.LoadFromWorkbook ThisWorkbook
        TestAssertTrue stats, prefix & ".recovery", settings.HasKey("Units.Length.Input")
    Next rangeName
    Exit Sub
Failed:
    stats.Failed = stats.Failed + 1
    TestAppendLine stats, "FAIL: audit03.settingMessage.requiredTable.runtime; " & CStr(Err.Number) & "; " & Err.Description
    On Error Resume Next
    If Not name Is Nothing Then
        If Len(savedReference) > 0 Then name.RefersTo = savedReference
    End If
    On Error GoTo 0
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' отделяет переполнение перевода единиц от ошибки чтения Double.
' Каждая настройка дает InputErr до solver-а с собственной ячейкой и действием;
' восстановление Config позволяет повторный запуск того же batch без ошибки.
Private Sub TestAudit03ConfigConversionMessages(ByRef stats As TUiTestStats)
    Dim target As Object, unitRange As Object, savedSettings As Variant, savedUnits As Variant
    Dim key As Variant, row As Long, cell As Object, unitRow As Long, prefix As String
    Dim settings As CSystemSettingsReader, units As CUnitSystem, batch As CBatchSectionCalculator
    Dim value As Double, description As String
    On Error GoTo Failed
    Set target = ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange
    Set unitRange = ThisWorkbook.Names.Item("rngUnitSettings").RefersToRange
    savedSettings = target.Formula: savedUnits = unitRange.Formula
    ' ДЛЯ ТЕСТОВ: готовим неизменяемую модель до смены единиц ввода.
    ' Проверка переполнения настройки не требует новой геометрии размером
    ' в тысячи раз больше и сетки с неизменным внутренним шагом 20 мм.
    ' Контрольный круг формируется по собственной таблице листа «Сечения» до изменения единиц.
    TestSetSystemSetting "Geometry.Type", "Circle"
    Set batch = TestBuildUiBatch()
    batch.AddCombination "CONFIG_OVERFLOW", 0#, 0#, 0#, "PR1", "Переполнение пересчета настройки"
    For unitRow = 2 To unitRange.Rows.Count
        Select Case CStr(unitRange.Cells(unitRow, 1).Value2)
            Case "Length": unitRange.Cells(unitRow, 2).Value2 = "m"
            Case "Force": unitRange.Cells(unitRow, 2).Value2 = "tf"
            Case "Moment": unitRange.Cells(unitRow, 2).Value2 = "tf*m"
        End Select
    Next unitRow
    For Each key In Array("SLS.Crack.Allowable", "Stability.ElementLength", _
            "Stability.AccidentalEccentricityUser1", "Stability.AccidentalEccentricityUser2", _
            "Solver.ToleranceN", "Solver.ToleranceMx", "Solver.ToleranceMy", "Calculation.ZeroMomentPerDepth")
        target.Formula = savedSettings
        Set cell = Nothing
        For row = 2 To target.Rows.Count
            If CStr(target.Cells(row, 1).Value2) = CStr(key) Then Set cell = target.Cells(row, 2): Exit For
        Next row
        If cell Is Nothing Then Err.Raise 5, , "Не найдена проверяемая настройка " & CStr(key)
        cell.Value2 = "1e308"
        Set settings = New CSystemSettingsReader
        settings.LoadFromWorkbook ThisWorkbook
        value = settings.GetRequiredDouble(CStr(key))
        prefix = "audit03.settingMessage.conversion." & CStr(key)
        TestAssertTrue stats, prefix & ".readableDouble", value > 1E+307
        Set units = New CUnitSystem: units.LoadFromSettings settings
        batch.ApplySettings settings, units
        batch.Execute
        TestAssertTextEquals stats, prefix & ".status", batch.ResultAt(1).Status, "InputErr"
        TestAssertTrue stats, prefix & ".noSolve", batch.SolverCallCount = 0
        description = batch.ResultAt(1).OverallMeta.ResultComment
        TestAudit03AssertSettingMessage stats, prefix, description, CStr(key), cell
        TestAssertTrue stats, prefix & ".units", InStr(1, description, "INPUT", vbTextCompare) > 0
        target.Formula = savedSettings
        settings.LoadFromWorkbook ThisWorkbook
        batch.ApplySettings settings, units
        batch.Execute
        TestAssertTrue stats, prefix & ".recovery", batch.ResultAt(1).Status <> "InputErr"
    Next key
    GoTo Restore
Failed:
    stats.Failed = stats.Failed + 1
    TestAppendLine stats, "FAIL: audit03.settingMessage.conversion.runtime; " & CStr(Err.Number) & "; " & Err.Description
Restore:
    On Error Resume Next
    If Not target Is Nothing Then target.Formula = savedSettings
    If Not unitRange Is Nothing Then unitRange.Formula = savedUnits
    On Error GoTo 0
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' ==================== ДЛЯ ТЕСТОВ: ВОССТАНОВЛЕНИЕ ГЕОМЕТРИИ ====================

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет восстановление модели и чтение схемы без AutoCAD и решения НДС.
' Временная книга содержит только заданные численные результаты
' и закрывается без сохранения.
Public Function RunAudit03GeometrySnapshotTests() As String
    Dim stats As TUiTestStats
    TestAudit03GeometrySnapshotContract stats
    TestAppendLine stats, "TOTAL_AUDIT03_GEOMETRY_SNAPSHOT: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunAudit03GeometrySnapshotTests = stats.Report
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Пустой сохранённый набор версии 1 требует нового обязательного блока;
' проверка выполняется в отдельной временной книге.
Private Sub TestInitializeEmptyContourFixture(ByVal workbook As Object, ByVal sheet As Object)
    workbook.Names.Add Name:="rngNDMSectionContours", RefersTo:="='" & sheet.Name & "'!$GR$5"
    Dim writer As CNDMResultsWriter, contours As CSectionContours
    Set writer = New CNDMResultsWriter: Set contours = New CSectionContours
    writer.WriteSectionContours workbook, contours, "FIXTURE"
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет дробные числа, mm/cm/m, поврежденные поля, перенос к последней
' строке листа и комментарий LC за прежней границей 1000. Текущий Config не
' может менять физическую геометрию, уже сохраненную с собственной размерностью.
Private Sub TestAudit03GeometrySnapshotContract(ByRef stats As TUiTestStats)
    Dim fixture As Object, sheet As Object, anchor As Object, properties As Object, source As Object
    Dim data As Variant, prop(1 To 3, 1 To 4) As Variant, configuration(1 To 2, 1 To 3) As Variant
    Dim results(1 To 3, 1 To 7) As Variant, operation As Variant
    Dim settings As CSystemSettingsReader, reader As CSectionPlotDataReader, model As CSectionModel
    Dim unitIndex As Long, position As Long, field As Variant, badValue As Variant, prefix As String, badIndex As Long
    Dim units As Variant, factors As Variant, fields As Variant, badValues As Variant
    Dim errorNumber As Long, description As String, solveCount As Long, row As Long, fieldName As Variant, cases As Long
    Dim summary(1 To 1005, 1 To 2) As Variant, comments As Object, before As Variant, actual As Variant
    On Error GoTo Failed
    solveCount = SectionEquilibriumSolveCount()
    Set fixture = Application.Workbooks.Add(-4167)
    Set sheet = fixture.Worksheets(1): sheet.Name = "Snapshot"
    TestInitializeEmptyContourFixture fixture, sheet
    Set anchor = sheet.Range("A5")
    fixture.Names.Add Name:="rngNDMSectionGeometry", RefersTo:="=Snapshot!" & anchor.Address
    Set properties = sheet.Range("R5")
    prop(1, 1) = "LoadCase": prop(1, 2) = "Parameter": prop(1, 3) = "Value": prop(1, 4) = "Unit"
    prop(2, 1) = "ALL": prop(2, 2) = "Bounds.MaxX": prop(2, 3) = 100#: prop(2, 4) = "mm"
    prop(3, 1) = "LATE": prop(3, 2) = "ProfileId": prop(3, 3) = "PR1": prop(3, 4) = "-"
    properties.Resize(3, 4).Value2 = prop
    fixture.Names.Add Name:="rngNDMSectionProperties", RefersTo:="=Snapshot!" & properties.Address
    configuration(1, 1) = "Параметр": configuration(1, 2) = "Значение": configuration(1, 3) = "Комментарий"
    configuration(2, 1) = "Plot.LoadCase": configuration(2, 2) = "LATE"
    sheet.Range("AO5").Resize(2, 3).Value2 = configuration
    Set settings = New CSystemSettingsReader: settings.LoadFromRange sheet.Range("AO5").Resize(2, 3)
    Set reader = New CSectionPlotDataReader
    Set source = ThisWorkbook.Names.Item("rngCalculationProfiles").RefersToRange
    sheet.Range("AF5").Resize(source.Rows.Count, source.Columns.Count).Value2 = source.Value2
    For row = 2 To source.Rows.Count
        If CStr(source.Cells(row, 2).Value2) = "Visualization.State" Then sheet.Range("AF5").Cells(row, 3).Value2 = "StrengthState"
        If CStr(source.Cells(row, 2).Value2) = "Visualization.Quantity" Then sheet.Range("AF5").Cells(row, 3).Value2 = "Stress"
    Next row
    fixture.Names.Add Name:="rngCalculationProfiles", RefersTo:="=Snapshot!" & sheet.Range("AF5").Resize(source.Rows.Count, source.Columns.Count).Address
    results(1, 1) = "LoadCase": results(1, 2) = "ProfileId": results(1, 3) = "StateType": results(1, 4) = "ElementID"
    results(1, 5) = "Strain": results(1, 6) = "Stress, MPa": results(1, 7) = "PhysicalState"
    results(2, 1) = "LATE": results(2, 2) = "PR1": results(2, 3) = "StrengthState": results(2, 4) = "C1"
    results(2, 5) = -0.0001: results(2, 6) = -1.25: results(2, 7) = "Compression"
    results(3, 1) = "LATE": results(3, 2) = "PR1": results(3, 3) = "StrengthState": results(3, 4) = "R1"
    results(3, 5) = 0.000015: results(3, 6) = 2.75: results(3, 7) = "Tension"
    sheet.Range("R15").Resize(3, 7).Value2 = results
    fixture.Names.Add Name:="rngNDMElementResults", RefersTo:="=Snapshot!$R$15"
    units = Array("mm", "cm", "m"): factors = Array(1#, 10#, 1000#)
    For position = 0 To 1
        If position = 0 Then Set anchor = sheet.Range("A5") Else Set anchor = sheet.Cells(sheet.Rows.Count - 2, 1)
        fixture.Names.Item("rngNDMSectionGeometry").RefersTo = "=Snapshot!" & anchor.Address
        For unitIndex = 0 To 2
            cases = cases + 1
            data = TestAudit03GeometrySnapshotArray(CStr(units(unitIndex)), CDbl(factors(unitIndex)))
            anchor.Resize(3, 15).Value2 = data
            prefix = "audit03.geometrySnapshot." & CStr(units(unitIndex)) & ".position" & CStr(position)
            Set model = ReadSectionGeometryFromResults(fixture, "AutoCADImport")
            TestAudit03CheckRestoredGeometry stats, prefix & ".model", model
            reader.LoadGeometryPreviewFromWorkbook fixture, settings
            TestAssertTrue stats, prefix & ".plot.count", reader.Count = 2
            If reader.Count = 2 Then
                TestAssertClose stats, prefix & ".plot.width", reader.Width(1), 123.75, 0.000000001
                TestAssertClose stats, prefix & ".plot.height", reader.Height(1), 87.5, 0.000000001
                TestAssertClose stats, prefix & ".plot.rotation", reader.Rotation(1), 0.3125, 0.000000001
                TestAssertClose stats, prefix & ".plot.rebarDiameter", reader.Diameter(2), 12.25, 0.000000001
            End If
            reader.LoadFromWorkbook fixture, settings
            TestAssertTrue stats, prefix & ".state.count", reader.Count = 2
            If reader.Count = 2 Then
                TestAssertClose stats, prefix & ".state.concreteStress", reader.ResultValue(1), -1.25, 0#
                TestAssertClose stats, prefix & ".state.rebarStress", reader.ResultValue(2), 2.75, 0#
            End If
        Next unitIndex
    Next position
    Set anchor = sheet.Range("A5")
    fixture.Names.Item("rngNDMSectionGeometry").RefersTo = "=Snapshot!" & anchor.Address
    data = TestAudit03GeometrySnapshotArray("mm", 1#)
    fields = Array(3, 4, 5, 8, 9, 10, 11, 12, 13, 14)
    badValues = Array("12oops", "abc", CVErr(2015), True)
    For Each field In fields
        If CLng(field) = 10 Then row = 2 Else row = 1
        For badIndex = 0 To UBound(badValues)
            cases = cases + 1
            badValue = badValues(badIndex)
            anchor.Resize(3, 15).Value2 = data
            anchor.Offset(row, CLng(field) - 1).Value2 = badValue
            prefix = "audit03.geometrySnapshot.invalid." & CStr(data(1, CLng(field))) & "." & CStr(badIndex)
            fieldName = Split(CStr(data(1, CLng(field))), ",")
            TestAudit03CaptureGeometrySnapshotError fixture, errorNumber, description
            TestAssertTrue stats, prefix & ".rejected", errorNumber <> 0
            TestAssertTrue stats, prefix & ".field", InStr(1, description, CStr(fieldName(0)), vbTextCompare) > 0
            TestAssertTrue stats, prefix & ".address", InStr(1, description, "Snapshot!" & anchor.Offset(row, CLng(field) - 1).Address(False, False), vbTextCompare) > 0
            TestAssertTrue stats, prefix & ".action", InStr(1, description, "Сформировать сечение", vbTextCompare) > 0
            If CLng(field) < 12 Then
                For Each operation In Array("Preview", "State")
                    TestAudit03CapturePlotGeometryError fixture, settings, reader, errorNumber, description, CStr(operation) = "State"
                    TestAssertTrue stats, prefix & ".plot." & CStr(operation) & ".rejected", errorNumber <> 0
                    TestAssertTrue stats, prefix & ".plot." & CStr(operation) & ".field", InStr(1, description, CStr(fieldName(0)), vbTextCompare) > 0
                    TestAssertTrue stats, prefix & ".plot." & CStr(operation) & ".address", InStr(1, description, "Snapshot!" & anchor.Offset(row, CLng(field) - 1).Address(False, False), vbTextCompare) > 0
                    TestAssertTrue stats, prefix & ".plot." & CStr(operation) & ".action", InStr(1, description, "Повторите", vbTextCompare) > 0
                Next operation
            End If
            anchor.Resize(3, 15).Value2 = data
            Set model = ReadSectionGeometryFromResults(fixture, "AutoCADImport")
            TestAudit03CheckRestoredGeometry stats, prefix & ".recovery", model
        Next badIndex
    Next field
    For Each field In Array(3, 4, 5, 10)
        cases = cases + 1
        If CLng(field) = 10 Then row = 2 Else row = 1
        anchor.Resize(3, 15).Value2 = data
        anchor.Offset(row, CLng(field) - 1).ClearContents
        prefix = "audit03.geometrySnapshot.requiredBlank." & CStr(data(1, CLng(field)))
        TestAudit03CaptureGeometrySnapshotError fixture, errorNumber, description
        TestAssertTrue stats, prefix & ".model", errorNumber = vbObjectError + 4367
        TestAudit03CapturePlotGeometryError fixture, settings, reader, errorNumber, description
        TestAssertTrue stats, prefix & ".preview", errorNumber = vbObjectError + 4719
        TestAudit03CapturePlotGeometryError fixture, settings, reader, errorNumber, description, True
        TestAssertTrue stats, prefix & ".state", errorNumber = vbObjectError + 4719
    Next field
    ' Готовая сетка хранит размеры, угол и собственные инерции каждого элемента.
    ' Потерянное значение не восстанавливаем молча по приближённой оболочке.
    For Each field In Array(8, 9, 11, 12, 13, 14)
        anchor.Resize(3, 15).Value2 = data
        anchor.Offset(1, CLng(field) - 1).ClearContents
        cases = cases + 1
        TestAudit03CaptureGeometrySnapshotError fixture, errorNumber, description
        prefix = "audit03.geometrySnapshot.missingGeometry." & CStr(field)
        TestAssertTrue stats, prefix & ".rejected", errorNumber <> 0
        TestAssertTrue stats, prefix & ".address", InStr(description, "Snapshot!" & anchor.Offset(1, CLng(field) - 1).Address(False, False)) > 0
        TestAssertTrue stats, prefix & ".action", InStr(description, "Сформировать сечение") > 0
        anchor.Resize(3, 15).Value2 = data
        Set model = ReadSectionGeometryFromResults(fixture, "AutoCADImport")
        TestAudit03CheckRestoredGeometry stats, prefix & ".recovery", model
    Next field
    anchor.Resize(3, 15).Value2 = data
    cases = cases + 1
    anchor.Offset(1, 7).NumberFormat = "@": anchor.Offset(1, 7).Value2 = "123.75"
    Set model = ReadSectionGeometryFromResults(fixture, "AutoCADImport")
    TestAssertClose stats, "audit03.geometrySnapshot.numericText", model.ConcreteWidth(1), 123.75, 0#
    anchor.Offset(1, 7).NumberFormat = "General"
    anchor.Resize(3, 15).Value2 = data
    cases = cases + 1
    anchor.Offset(0, 6).Value2 = "WrongShapeHeader"
    TestAudit03CaptureGeometrySnapshotError fixture, errorNumber, description
    TestAssertTrue stats, "audit03.geometrySnapshot.missingShape.controlled", errorNumber = vbObjectError + 4355
    TestAssertTrue stats, "audit03.geometrySnapshot.missingShape.named", InStr(1, description, "GeometryInterpretationStatus", vbTextCompare) > 0
    anchor.Offset(0, 6).Value2 = "ShapeType"
    cases = cases + 1
    TestAudit03CaptureGeometrySnapshotError fixture, errorNumber, description
    TestAssertTrue stats, "audit03.geometrySnapshot.obsoleteShapeHeader.rejected", errorNumber = vbObjectError + 4355
    anchor.Resize(3, 15).Value2 = data
    Set comments = sheet.Range("R100")
    fixture.Names.Add Name:="rngBatchSummary", RefersTo:="=Snapshot!" & comments.Address
    For row = 1 To 1005
        summary(row, 1) = "OTHER_" & CStr(row): summary(row, 2) = "Другой комментарий"
    Next row
    summary(1005, 1) = "LATE": summary(1005, 2) = "Комментарий позднего сочетания"
    cases = cases + 1
    comments.Offset(11, 0).Resize(1005, 2).Value2 = summary
    before = anchor.Resize(3, 15).Value2
    reader.LoadGeometryOnlyForMissingState fixture, settings, "Нет расчетного состояния в test fixture"
    TestAssertTextEquals stats, "audit03.geometrySnapshot.lateComment", reader.LoadCaseComment, "Комментарий позднего сочетания"
    comments.Offset(12, 0).Resize(2, 2).ClearContents ' Сводка сохраняет пустые позиции исходных строк LC; они не заканчивают таблицу.
    reader.LoadGeometryOnlyForMissingState fixture, settings, "Нет расчетного состояния в test fixture"
    TestAssertTextEquals stats, "workingReview.plotComment.afterBlankSlots", reader.LoadCaseComment, "Комментарий позднего сочетания"
    actual = anchor.Resize(3, 15).Value2
    TestAssertTrue stats, "audit03.geometrySnapshot.unchanged", TestAudit02SnapshotTablesEqual(before, actual)
    ' Ширина таблицы тоже должна учитывать реальную границу листа, если
    ' заголовок последнего поля расположен ровно в последнем столбце Excel.
    Set anchor = sheet.Cells(5, sheet.Columns.Count - 14)
    cases = cases + 1
    fixture.Names.Item("rngNDMSectionGeometry").RefersTo = "=Snapshot!" & anchor.Address
    anchor.Resize(3, 15).Value2 = data
    TestAudit03CaptureGeometrySnapshotError fixture, errorNumber, description
    TestAssertTrue stats, "audit03.geometrySnapshot.lastColumn.model", errorNumber = 0
    TestAudit03CapturePlotGeometryError fixture, settings, reader, errorNumber, description
    TestAssertTrue stats, "audit03.geometrySnapshot.lastColumn.plot", errorNumber = 0
    TestAssertTrue stats, "audit03.geometrySnapshot.noSolve", SectionEquilibriumSolveCount() = solveCount
    GoTo Restore
Failed:
    stats.Failed = stats.Failed + 1
    TestAppendLine stats, "FAIL: audit03.geometrySnapshot.runtime; " & CStr(Err.Number) & "; " & Err.Description
Restore:
    On Error Resume Next
    If Not fixture Is Nothing Then fixture.Close False
    On Error GoTo 0
    TestAppendLine stats, "CASES_GEOMETRY_SNAPSHOT: independentGeometryFixtures=1; inputVariants=" & CStr(cases) & "; equilibriumCases=0"
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сохраняет ошибку предварительного просмотра и повторяет чтение тем же объектом.
' После исправления входных данных прежние элементы должны быть очищены.
Private Sub TestAudit03CapturePlotGeometryError(ByVal workbook As Object, ByVal settings As CSystemSettingsReader, _
        ByVal reader As CSectionPlotDataReader, ByRef number As Long, ByRef description As String, _
        Optional ByVal readState As Boolean = False)
    number = 0: description = vbNullString
    On Error GoTo Failed
    If readState Then reader.LoadFromWorkbook workbook, settings Else reader.LoadGeometryPreviewFromWorkbook workbook, settings
    Exit Sub
Failed:
    number = Err.Number: description = Err.Description
    Err.Clear
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Задает дробные параметры для проверки сохранности в трех системах OUTPUT.
' A и оболочка согласованы, I заданы отдельно для контроля восстановления;
' это не live Region или инженерный solve. Масштаб fixture независим от reader-а.
Private Function TestAudit03GeometrySnapshotArray(ByVal unit As String, ByVal factor As Double) As Variant
    Dim data(1 To 3, 1 To 15) As Variant, headers As Variant, column As Long
    headers = Array("RunID", "MaterialType", "X, " & unit, "Y, " & unit, "Area, " & unit & "2", _
        "ElementID", "GeometryInterpretationStatus", "Width, " & unit, "Height, " & unit, _
        "Diameter, " & unit, "Rotation, rad", "LocalIx, " & unit & "4", "LocalIy, " & unit & "4", "LocalIxy, " & unit & "4", "Comment")
    For column = 1 To 15: data(1, column) = headers(column - 1): Next column
    data(2, 1) = 46023#: data(2, 2) = "Concrete": data(2, 3) = -23.125 / factor: data(2, 4) = 17.625 / factor
    data(2, 5) = 10828.125 / factor ^ 2: data(2, 6) = "C1": data(2, 7) = "Rectangle"
    data(2, 8) = 123.75 / factor: data(2, 9) = 87.5 / factor: data(2, 10) = 0#: data(2, 11) = 0.3125
    data(2, 12) = 654321.125 / factor ^ 4: data(2, 13) = 987654.375 / factor ^ 4: data(2, 14) = -23456.625 / factor ^ 4
    data(2, 15) = "Дробные параметры Region"
    data(3, 1) = 46023#: data(3, 2) = "Rebar": data(3, 3) = 12.375 / factor: data(3, 4) = -8.125 / factor
    data(3, 5) = 117.8581 / factor ^ 2: data(3, 6) = "R1": data(3, 7) = "Circle": data(3, 10) = 12.25 / factor
    data(3, 15) = "Дробная арматура"
    TestAudit03GeometrySnapshotArray = data
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сверяет расчетные числа модели, включая инерции с ненулевым signed Ixy.
' Оболочка и площадь имеют разный смысл; сохраненные I не вычисляются заново.
Private Sub TestAudit03CheckRestoredGeometry(ByRef stats As TUiTestStats, ByVal prefix As String, ByVal model As CSectionModel)
    TestAssertTrue stats, prefix & ".concreteCount", model.ConcreteCount = 1
    TestAssertTrue stats, prefix & ".rebarCount", model.RebarCount = 1
    TestAssertClose stats, prefix & ".x", model.ConcreteX(1), -23.125, 0.000000001
    TestAssertClose stats, prefix & ".y", model.ConcreteY(1), 17.625, 0.000000001
    TestAssertClose stats, prefix & ".area", model.ConcreteArea(1), 10828.125, 0.000000001
    TestAssertClose stats, prefix & ".width", model.ConcreteWidth(1), 123.75, 0.000000001
    TestAssertClose stats, prefix & ".height", model.ConcreteHeight(1), 87.5, 0.000000001
    TestAssertClose stats, prefix & ".rotation", model.ConcreteRotation(1), 0.3125, 0.000000001
    TestAssertClose stats, prefix & ".ix", model.ConcreteLocalIx(1), 654321.125, 0.000000001
    TestAssertClose stats, prefix & ".iy", model.ConcreteLocalIy(1), 987654.375, 0.000000001
    TestAssertClose stats, prefix & ".ixy", model.ConcreteLocalIxy(1), -23456.625, 0.000000001
    If model.RebarCount = 1 Then
        TestAssertClose stats, prefix & ".rebarX", model.RebarX(1), 12.375, 0.000000001
        TestAssertClose stats, prefix & ".rebarY", model.RebarY(1), -8.125, 0.000000001
        TestAssertClose stats, prefix & ".rebarArea", model.RebarArea(1), 117.8581, 0.000000001
        TestAssertClose stats, prefix & ".rebarDiameter", model.RebarDiameter(1), 12.25, 0.000000001
    End If
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сохраняет номер и текст ошибки публичного чтения без их трактовки.
' Отдельный перехват позволяет проверить повторную работу после исправления данных.
Private Sub TestAudit03CaptureGeometrySnapshotError(ByVal workbook As Object, ByRef number As Long, ByRef description As String)
    Dim model As CSectionModel
    number = 0: description = vbNullString
    On Error GoTo Failed
    Set model = ReadSectionGeometryFromResults(workbook, "AutoCADImport")
    Exit Sub
Failed:
    number = Err.Number: description = Err.Description
    Err.Clear
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' ==================== ДЛЯ ТЕСТОВ: ОШИБКА ЧТЕНИЯ И ПОВТОРНЫЙ ВВОД ====================

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Запускает направленную проверку незавершенной инициализации и поврежденных
' чисел сохраненного снимка. Использует только собственную временную книгу,
' не ищет равновесие и не меняет исходный Config/Results основной книги.
Public Function RunAudit03ReadLifecycleTests() As String
    Dim stats As TUiTestStats
    TestAudit03ReadLifecycleContracts stats
    TestAppendLine stats, "TOTAL_AUDIT03_READ_LIFECYCLE: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunAudit03ReadLifecycleTests = stats.Report
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет valid -> invalid -> valid на тех же экземплярах Spec, Catalog
' и PlotDataReader. Старая корректная копия результата должна оставаться
' независимой, но новый неуспешный ввод не должен выглядеть как готовые данные.
Private Sub TestAudit03ReadLifecycleContracts(ByRef stats As TUiTestStats)
    Dim fixture As Object, sheet As Object, profiles As Object, source As Object, target As Object
    Dim spec As CMaterialModelSpec, savedSpec As CMaterialModelSpec, copied As CMaterialModelSpec
    Dim provider As CMaterialModelProvider, diagram As CMaterialDiagram, baselineDiagram As CMaterialDiagram
    Dim catalog As CCalculationProfileCatalog, profile As CCalculationProfile, savedProfile As CCalculationProfile
    Dim settings As CSystemSettingsReader, reader As CSectionPlotDataReader
    Dim arguments As Variant, index As Long, field As Long, row As Long, column As Long, position As Long
    Dim bad As Variant, badIndex As Long, mode As Long, code As Long, reason As String, prefix As String
    Dim baseline As Variant, actual As Variant, values As Variant, keys As Variant, badValues As Variant
    Dim props(1 To 22, 1 To 4) As Variant, results(1 To 3, 1 To 7) As Variant
    Dim configuration(1 To 2, 1 To 3) As Variant, solveCount As Long, cases As Long
    On Error GoTo Failed
    solveCount = SectionEquilibriumSolveCount()
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set provider = New CMaterialModelProvider: provider.Initialize settings
    Set spec = New CMaterialModelSpec
    For field = 0 To 3
        arguments = Array("SLS(II)", "ThreeLine", "UseDiagram", "TwoLine")
        spec.Initialize CStr(arguments(0)), CStr(arguments(1)), CStr(arguments(2)), CStr(arguments(3))
        Set savedSpec = spec.Clone: Set baselineDiagram = provider.ConcreteMaterialFromSpec(savedSpec)
        arguments(field) = "INVALID_SPEC"
        On Error Resume Next
        Err.Clear
        spec.Initialize CStr(arguments(0)), CStr(arguments(1)), CStr(arguments(2)), CStr(arguments(3))
        code = Err.Number: reason = Err.Description
        On Error GoTo Failed
        prefix = "audit03.readLifecycle.spec.field" & CStr(field): cases = cases + 1
        TestAssertTrue stats, prefix & ".error", code <> 0
        TestAssertTrue stats, prefix & ".incomplete", Not spec.IsComplete
        TestAssertTextEquals stats, prefix & ".key", spec.SpecKey, "<empty>"
        Set copied = spec.Clone
        TestAssertTrue stats, prefix & ".cloneIncomplete", Not copied.IsComplete
        On Error Resume Next
        Err.Clear
        Set diagram = Nothing: Set diagram = provider.ConcreteMaterialFromSpec(spec)
        code = Err.Number
        On Error GoTo Failed
        TestAssertTrue stats, prefix & ".providerRejected", code = vbObjectError + 3251
        TestAssertTrue stats, prefix & ".noDiagram", diagram Is Nothing
        TestAssertTextEquals stats, prefix & ".snapshot", savedSpec.SpecKey, "SLS(II)|ThreeLine|UseDiagram|TwoLine"
        spec.Initialize "ULS(I)", "TwoLine", "Ignore", "ThreeLine"
        Set diagram = provider.ConcreteMaterialFromSpec(spec)
        TestAssertTrue stats, prefix & ".recovery", spec.IsComplete
        TestAssertTextEquals stats, prefix & ".recoveryKey", spec.SpecKey, "ULS(I)|TwoLine|Ignore|ThreeLine"
        TestAssertTrue stats, prefix & ".distinctDiagram", Not diagram Is baselineDiagram
        TestAppendLine stats, "READ_LIFECYCLE_SPEC: " & prefix & "|" & reason
    Next field

    Set fixture = Application.Workbooks.Add(-4167)
    Set sheet = fixture.Worksheets(1): sheet.Name = "ReadLifecycle"
    Set source = ThisWorkbook.Names.Item("rngCalculationProfiles").RefersToRange
    Set profiles = sheet.Range("AF5").Resize(source.Rows.Count, source.Columns.Count)
    profiles.Value2 = source.Value2: baseline = profiles.Value2
    Set catalog = New CCalculationProfileCatalog: catalog.LoadFromRange profiles
    Set savedProfile = catalog.ProfileById("PR1")
    keys = Array("Calculation.Strength.DirectState", "MaterialModel.Strength.ConcreteDiagram", _
        "Visualization.State", "Visualization.StressPrecision")
    For position = 0 To 1
        If position = 0 Then Set target = profiles Else Set target = sheet.Range("CH800").Resize(source.Rows.Count, source.Columns.Count)
        target.Value2 = baseline
        For Each bad In Array("BAD_SELECTOR", CVErr(2015), "")
            For index = 0 To UBound(keys)
                For row = 2 To target.Rows.Count
                    If CStr(target.Cells(row, 2).Value2) = CStr(keys(index)) Then Exit For
                Next row
                If row > target.Rows.Count Then Err.Raise 5, , "Нет test-параметра " & CStr(keys(index))
                column = 6 ' Ошибка в последнем PR4 после уже прочитанных PR1-PR3.
                target.Cells(row, column).Value2 = bad
                On Error Resume Next
                Err.Clear: catalog.LoadFromRange target
                code = Err.Number: reason = Err.Description
                On Error GoTo Failed
                prefix = "audit03.readLifecycle.catalog.position" & CStr(position) & ".field" & CStr(index) & ".type" & CStr(VarType(bad))
                cases = cases + 1
                TestAssertTrue stats, prefix & ".error", code <> 0
                TestAssertTrue stats, prefix & ".address", InStr(1, reason, sheet.Name & "!" & target.Cells(row, column).Address(False, False), vbBinaryCompare) > 0
                TestAssertTrue stats, prefix & ".empty", catalog.Count = 0
                TestAssertTrue stats, prefix & ".noPR1", Not catalog.HasProfile("PR1")
                TestAssertTrue stats, prefix & ".noPR4", Not catalog.HasProfile("PR4")
                On Error Resume Next
                Err.Clear: Set profile = Nothing: Set profile = catalog.ProfileById("PR1")
                code = Err.Number
                On Error GoTo Failed
                TestAssertTrue stats, prefix & ".lookupRejected", code = vbObjectError + 3984
                TestAssertTrue stats, prefix & ".noPublishedProfile", profile Is Nothing
                TestAssertTextEquals stats, prefix & ".oldSnapshot", savedProfile.ProfileId, "PR1"
                target.Value2 = baseline: catalog.LoadFromRange target
                TestAssertTrue stats, prefix & ".recoveryCount", catalog.Count = 4
                TestAssertTrue stats, prefix & ".recoveryPR4", catalog.HasProfile("PR4")
                TestAppendLine stats, "READ_LIFECYCLE_CATALOG: " & prefix & "|" & reason
            Next index
        Next bad
    Next position

    fixture.Names.Add Name:="rngCalculationProfiles", RefersTo:="=ReadLifecycle!" & profiles.Address
    For row = 2 To profiles.Rows.Count
        If CStr(profiles.Cells(row, 2).Value2) = "Visualization.State" Then profiles.Cells(row, 3).Value2 = "StrengthState"
    Next row
    values = TestAudit03GeometrySnapshotArray("mm", 1#): sheet.Range("A5").Resize(3, 15).Value2 = values
    fixture.Names.Add Name:="rngNDMSectionGeometry", RefersTo:="=ReadLifecycle!$A$5"
    TestInitializeEmptyContourFixture fixture, sheet
    results(1, 1) = "LoadCase": results(1, 2) = "ProfileId": results(1, 3) = "StateType": results(1, 4) = "ElementID"
    results(1, 5) = "Strain": results(1, 6) = "Stress, MPa": results(1, 7) = "PhysicalState"
    results(2, 1) = "LIFE": results(2, 2) = "PR1": results(2, 3) = "StrengthState": results(2, 4) = "C1"
    results(2, 5) = -0.0001: results(2, 6) = -1.25: results(2, 7) = "Compression"
    results(3, 1) = "LIFE": results(3, 2) = "PR1": results(3, 3) = "StrengthState": results(3, 4) = "R1"
    results(3, 5) = 0.000015: results(3, 6) = 2.75: results(3, 7) = "Tension"
    sheet.Range("R35").Resize(3, 7).Value2 = results
    fixture.Names.Add Name:="rngNDMElementResults", RefersTo:="=ReadLifecycle!$R$35"
    configuration(1, 1) = "Параметр": configuration(1, 2) = "Значение": configuration(1, 3) = "Комментарий"
    configuration(2, 1) = "Plot.LoadCase": configuration(2, 2) = "LIFE"
    sheet.Range("AO70").Resize(2, 3).Value2 = configuration
    Set settings = New CSystemSettingsReader: settings.LoadFromRange sheet.Range("AO70").Resize(2, 3)
    Set reader = New CSectionPlotDataReader
    keys = Array("Bounds.MinX", "Bounds.MaxX", "Bounds.MinY", "Bounds.MaxY", _
        "Concrete.CentroidX", "Concrete.CentroidY", "Concrete.PrincipalAngle", _
        "Transformed.CentroidX", "Transformed.CentroidY", "Transformed.PrincipalAngle", _
        "LoadReferenceX", "LoadReferenceY", "State.StrengthState.Epsilon0", _
        "State.StrengthState.KappaX", "State.StrengthState.KappaY")
    props(1, 1) = "LoadCase": props(1, 2) = "Parameter": props(1, 3) = "Value": props(1, 4) = "Unit"
    For index = 0 To UBound(keys)
        row = index + 2: props(row, 1) = "ALL": If index >= 12 Then props(row, 1) = "LIFE"
        props(row, 2) = keys(index): props(row, 3) = 0.125: props(row, 4) = "mm"
        If InStr(1, CStr(keys(index)), "Angle", vbTextCompare) > 0 Or index = 12 Then props(row, 4) = "-"
        If index >= 13 Then props(row, 4) = "1/mm"
    Next index
    props(17, 1) = "LIFE": props(17, 2) = "ProfileId": props(17, 3) = "PR1": props(17, 4) = "-"
    props(18, 1) = "LIFE": props(18, 2) = "State.StrengthState.ExtensionUsed": props(18, 3) = "False": props(18, 4) = "-"
    props(19, 1) = "LIFE": props(19, 2) = "State.StrengthState.Status": props(19, 3) = "OK": props(19, 4) = "-"
    props(20, 1) = "OTHER": props(20, 2) = "State.StrengthState.Epsilon0": props(20, 3) = CVErr(2015): props(20, 4) = "-"
    props(21, 1) = "LIFE": props(21, 2) = "State.CapacityState.Epsilon0": props(21, 3) = CVErr(2015): props(21, 4) = "-"
    props(22, 1) = "ALL": props(22, 2) = "Output.LengthUnit": props(22, 3) = "mm": props(22, 4) = "-"
    badValues = Array("123oops", CVErr(2015), True, "", "1e309")
    For position = 0 To 1
        If position = 0 Then Set target = sheet.Range("R5") Else Set target = sheet.Range("DC800")
        target.Resize(22, 4).Value2 = props
        fixture.Names.Add Name:="rngNDMSectionProperties", RefersTo:="=ReadLifecycle!" & target.Address
        For mode = 0 To 2
            TestAudit03ReadLifecycleLoad reader, fixture, settings, mode, code, reason
            TestAssertTrue stats, "audit03.readLifecycle.reader.baseline.p" & CStr(position) & ".m" & CStr(mode), code = 0 And reader.Count = 2
            For index = 0 To UBound(keys)
                If mode = 1 And index >= 12 Then Exit For ' Предварительный просмотр импортированной геометрии не использует поле деформаций выбранного НДС.
                For badIndex = 0 To 4
                    bad = badValues(badIndex)
                    row = index + 2: target.Cells(row, 3).Value2 = bad
                    TestAudit03ReadLifecycleLoad reader, fixture, settings, mode, code, reason
                    prefix = "audit03.readLifecycle.reader.p" & CStr(position) & ".m" & CStr(mode) & ".field" & CStr(index) & ".bad" & CStr(badIndex)
                    cases = cases + 1
                    TestAssertTrue stats, prefix & ".error", code = vbObjectError + 4719
                    TestAssertTrue stats, prefix & ".field", InStr(1, reason, CStr(keys(index)), vbTextCompare) > 0
                    TestAssertTrue stats, prefix & ".address", InStr(1, reason, sheet.Name & "!" & target.Cells(row, 3).Address(False, False), vbBinaryCompare) > 0
                    TestAssertTrue stats, prefix & ".action", InStr(1, reason, "Повторите", vbTextCompare) > 0
                    TestAssertTrue stats, prefix & ".empty", reader.Count = 0 And reader.AnnotationCount = 0
                    TestAssertTrue stats, prefix & ".clearedPlane", reader.Epsilon0 = 0# And reader.KappaX = 0# And reader.KappaY = 0#
                    TestAssertTrue stats, prefix & ".clearedSelection", Len(reader.LoadCase) = 0 And Len(reader.ProfileId) = 0
                    TestAppendLine stats, "READ_LIFECYCLE_SNAPSHOT: " & prefix & "|" & reason
                    target.Cells(row, 3).Value2 = props(row, 3)
                    TestAudit03ReadLifecycleLoad reader, fixture, settings, mode, code, reason
                    TestAssertTrue stats, prefix & ".recovery", code = 0 And reader.Count = 2
                    If mode = 0 Then TestAssertClose stats, prefix & ".recoveryPlane", reader.Epsilon0, 0.125, 0#
                Next badIndex
            Next index
            sheet.Range("A5").Cells(3, 3).Value2 = "late-invalid" ' Поздняя ошибка после AppendElement не публикует первый элемент как весь снимок.
            TestAudit03ReadLifecycleLoad reader, fixture, settings, mode, code, reason
            prefix = "audit03.readLifecycle.reader.late.p" & CStr(position) & ".m" & CStr(mode)
            TestAssertTrue stats, prefix & ".error", code = vbObjectError + 4719
            TestAssertTrue stats, prefix & ".empty", reader.Count = 0 And reader.AnnotationCount = 0
            sheet.Range("A5").Resize(3, 15).Value2 = values
            TestAudit03ReadLifecycleLoad reader, fixture, settings, mode, code, reason
            TestAssertTrue stats, prefix & ".recovery", code = 0 And reader.Count = 2
        Next mode
        For index = 0 To UBound(keys)
            target.Cells(index + 2, 3).Value2 = "0.125"
        Next index
        TestAudit03ReadLifecycleLoad reader, fixture, settings, 0, code, reason
        TestAssertTrue stats, "audit03.readLifecycle.numericText.p" & CStr(position), code = 0 And reader.Count = 2
        TestAssertClose stats, "audit03.readLifecycle.numericText.epsilon.p" & CStr(position), reader.Epsilon0, 0.125, 0#
        TestAssertClose stats, "audit03.readLifecycle.numericText.kappaX.p" & CStr(position), reader.KappaX, 0.125, 0#
        TestAssertClose stats, "audit03.readLifecycle.numericText.kappaY.p" & CStr(position), reader.KappaY, 0.125, 0#
        target.Resize(22, 4).Value2 = props
        For column = 5 To 6
            For badIndex = 0 To 4
                sheet.Range("R35").Cells(3, column).Value2 = badValues(badIndex)
                TestAudit03ReadLifecycleLoad reader, fixture, settings, 0, code, reason
                prefix = "audit03.readLifecycle.element.p" & CStr(position) & ".column" & CStr(column) & ".bad" & CStr(badIndex)
                cases = cases + 1
                TestAssertTrue stats, prefix & ".error", code = vbObjectError + 4719
                TestAssertTrue stats, prefix & ".field", InStr(1, reason, CStr(results(1, column)), vbTextCompare) > 0
                TestAssertTrue stats, prefix & ".address", InStr(1, reason, sheet.Name & "!" & sheet.Range("R35").Cells(3, column).Address(False, False), vbBinaryCompare) > 0
                TestAssertTrue stats, prefix & ".empty", reader.Count = 0
                TestAppendLine stats, "READ_LIFECYCLE_ELEMENT: " & prefix & "|" & reason
                sheet.Range("R35").Resize(3, 7).Value2 = results
                TestAudit03ReadLifecycleLoad reader, fixture, settings, 0, code, reason
                TestAssertTrue stats, prefix & ".recovery", code = 0 And reader.Count = 2
            Next badIndex
        Next column
        sheet.Range("R35").Cells(2, 5).Value2 = "-0.0001"
        sheet.Range("R35").Cells(3, 6).Value2 = "2.75"
        TestAudit03ReadLifecycleLoad reader, fixture, settings, 0, code, reason
        TestAssertTrue stats, "audit03.readLifecycle.element.numericText.p" & CStr(position), code = 0 And reader.Count = 2
        TestAssertClose stats, "audit03.readLifecycle.element.numericValue.p" & CStr(position), reader.ResultValue(2), 2.75, 0#
        sheet.Range("R35").Resize(3, 7).Value2 = results
        actual = target.Resize(22, 4).Value2
        TestAudit03ComparePlainSnapshot stats, "audit03.readLifecycle.snapshotUnchanged.p" & CStr(position), props, actual, 1
    Next position
    TestAssertTrue stats, "audit03.readLifecycle.noSolve", SectionEquilibriumSolveCount() = solveCount
    TestAppendLine stats, "READ_LIFECYCLE_CASES: variants=" & CStr(cases) & "; geometryFixtures=1; equilibriumCases=0"
    GoTo Cleanup
Failed:
    TestAssertTrue stats, "audit03.readLifecycle.runtime." & CStr(Err.Number) & "." & Err.Description, False
Cleanup:
    On Error Resume Next
    If Not fixture Is Nothing Then fixture.Close False
    On Error GoTo 0
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сохраняет номер и текст ошибки выбранного режима чтения. Повторное чтение
' выполняется на том же объекте без подстановки других Results или Config.
Private Sub TestAudit03ReadLifecycleLoad(ByVal reader As CSectionPlotDataReader, ByVal workbook As Object, _
        ByVal settings As CSystemSettingsReader, ByVal mode As Long, ByRef code As Long, ByRef reason As String)
    code = 0: reason = vbNullString
    On Error GoTo Failed
    Select Case mode
        Case 0: reader.LoadFromWorkbook workbook, settings
        Case 1: reader.LoadGeometryPreviewFromWorkbook workbook, settings
        Case 2: reader.LoadGeometryOnlyForMissingState workbook, settings
    End Select
    Exit Sub
Failed:
    code = Err.Number: reason = Err.Description
    Err.Clear
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' ==================== ДЛЯ ТЕСТОВ: ОБЩИЕ НАСТРОЙКИ EXCEL-СХЕМЫ ====================

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сохраняет направленный entrypoint интерфейсного набора; сами проверки
' готового snapshot и layout находятся в отдельном стандартном test-модуле.
Public Function RunAudit03GeneralPlotTests() As String
    RunAudit03GeneralPlotTests = modTestPlotConfig.RunAudit03GeneralPlotTests()
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' ==================== ДЛЯ ТЕСТОВ: СОСТОЯНИЕ EXCEL.APPLICATION ====================

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет восстановление настроек собственного настоящего Excel.Application,
' включая отказ записи Calculation без открытых книг. Подставные классы
' и пользовательский Excel в этой проверке не используются.
Public Function RunAudit03ExcelGuardTests() As String
    Dim stats As TUiTestStats
    TestExcelGuardLocalCancellation stats
    TestAudit03ExcelGuardContracts stats
    TestExcelGuardFaultRecovery stats
    TestAppendLine stats, "TOTAL_AUDIT03_EXCEL_GUARD: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunAudit03ExcelGuardTests = stats.Report
End Function


' -------------------- ДЛЯ ТЕСТОВ --------------------
' Снимает PID и время создания всех прежних Excel до CreateObject; отказ WMI
' запрещает тесту менять видимость, состояние или закрывать неизвестный процесс.
Private Function TestExcelProcessSnapshot() As Object
    Dim result As Object, process As Object
    Set result = CreateObject("Scripting.Dictionary")
    For Each process In GetObject("winmgmts:").ExecQuery("SELECT ProcessId,CreationDate FROM Win32_Process WHERE Name='EXCEL.EXE'")
        result.Add CStr(process.ProcessId), CStr(process.CreationDate)
    Next process
    Set TestExcelProcessSnapshot = result
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Получает текущую идентичность конкретного COM-окна; пустой результат означает
' отсутствие доказательства владения, поэтому cleanup не получает разрешения Quit.
Private Function TestExcelProcessIdentity(ByVal excel As Object, ByRef processID As Long) As String
    Dim process As Object
    processID = 0
    TestExcelWindowProcessId excel.hwnd, processID
    If processID <= 0 Then Exit Function
    For Each process In GetObject("winmgmts:").ExecQuery("SELECT CreationDate FROM Win32_Process WHERE Name='EXCEL.EXE' AND ProcessId=" & CStr(processID))
        TestExcelProcessIdentity = CStr(process.CreationDate)
    Next process
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Восемь сочетаний Boolean и три режима пересчета проверяются через тот же
' guard, что используется workbook-сценарием. После failed Enter проверяем
' состояние сразу, затем явный Restore, повторный Enter и деструктор.
Private Sub TestAudit03ExcelGuardContracts(ByRef stats As TUiTestStats)
    Dim otherExcel As Object, book As Object, guard As CExcelAppStateGuard, owned As Boolean
    Dim mode As Variant, flags As Long, screen As Boolean, events As Boolean, alerts As Boolean
    Dim code As Long, source As String, reason As String, nativeCode As Long, nativeSource As String, nativeReason As String
    Dim prefix As String, solveCount As Long, existing As Object, ownedPID As Long, checkedPID As Long, ownedStart As String, currentStart As String
    On Error GoTo Failed
    solveCount = SectionEquilibriumSolveCount()
    Set existing = TestExcelProcessSnapshot()
    Set otherExcel = CreateObject("Excel.Application")
    ownedStart = TestExcelProcessIdentity(otherExcel, ownedPID)
    If ownedPID <= 0 Or Len(ownedStart) = 0 Then Err.Raise 5, "TestAudit03ExcelGuardContracts", "Не удалось доказать PID/время создания отдельного Excel."
    If existing.Exists(CStr(ownedPID)) Then Err.Raise 5, "TestAudit03ExcelGuardContracts", "CreateObject вернул ранее открытый Excel; его состояние не менялось."
    owned = True ' Только новый PID с проверенным временем создания получает право изменять свойства и завершаться.
    TestAppendLine stats, "EXCEL_GUARD_OWNED: pid=" & ownedPID & "; creationDate=" & ownedStart
    otherExcel.Visible = False
    TestAssertTrue stats, "audit03.excelGuard.emptyApplication", otherExcel.Workbooks.Count = 0
    If otherExcel.Workbooks.Count <> 0 Then Err.Raise 5, "TestAudit03ExcelGuardContracts", "Новый Excel уже содержит книги: проверка пустого приложения остановлена."
    On Error Resume Next
    Err.Clear: otherExcel.Calculation = xlCalculationManual
    nativeCode = Err.Number: nativeSource = Err.Source: nativeReason = Err.Description
    On Error GoTo Failed
    TestAssertTrue stats, "audit03.excelGuard.nativeFault", nativeCode <> 0
    Set guard = New CExcelAppStateGuard
    For flags = 0 To 7
        screen = ((flags And 1) <> 0): events = ((flags And 2) <> 0): alerts = ((flags And 4) <> 0)
        otherExcel.ScreenUpdating = screen: otherExcel.EnableEvents = events: otherExcel.DisplayAlerts = alerts
        On Error Resume Next
        Err.Clear: guard.Enter otherExcel
        code = Err.Number: source = Err.Source: reason = Err.Description
        On Error GoTo Failed
        prefix = "audit03.excelGuard.failure.f" & CStr(flags)
        TestAssertTrue stats, prefix & ".originalError", code = nativeCode And source = nativeSource And reason = nativeReason
        TestAssertTrue stats, prefix & ".immediateScreen", otherExcel.ScreenUpdating = screen
        TestAssertTrue stats, prefix & ".immediateEvents", otherExcel.EnableEvents = events
        TestAssertTrue stats, prefix & ".immediateAlerts", otherExcel.DisplayAlerts = alerts
        guard.Restore
        TestAssertTrue stats, prefix & ".restore", otherExcel.ScreenUpdating = screen And otherExcel.EnableEvents = events And otherExcel.DisplayAlerts = alerts
        TestAppendLine stats, "EXCEL_GUARD_ERROR: " & prefix & "|code=" & CStr(code) & "|source=" & source & "|" & reason
        otherExcel.ScreenUpdating = screen: otherExcel.EnableEvents = events: otherExcel.DisplayAlerts = alerts ' Возвращаем тестовую копию в заданное состояние также при ожидаемой неудаче.
    Next flags
    Set book = otherExcel.Workbooks.Add(-4167)
    For Each mode In Array(xlCalculationAutomatic, xlCalculationManual, xlCalculationSemiautomatic)
        For flags = 0 To 7
            screen = ((flags And 1) <> 0): events = ((flags And 2) <> 0): alerts = ((flags And 4) <> 0)
            otherExcel.Calculation = mode
            otherExcel.ScreenUpdating = screen: otherExcel.EnableEvents = events: otherExcel.DisplayAlerts = alerts
            prefix = "audit03.excelGuard.success.m" & CStr(mode) & ".f" & CStr(flags)
            guard.Enter otherExcel
            TestAssertTrue stats, prefix & ".entered", Not otherExcel.ScreenUpdating And Not otherExcel.EnableEvents And Not otherExcel.DisplayAlerts And otherExcel.Calculation = xlCalculationManual
            guard.Enter otherExcel
            guard.Restore
            TestAssertTrue stats, prefix & ".restored", otherExcel.ScreenUpdating = screen And otherExcel.EnableEvents = events And otherExcel.DisplayAlerts = alerts And otherExcel.Calculation = mode
            guard.Restore
            TestAssertTrue stats, prefix & ".idempotent", otherExcel.ScreenUpdating = screen And otherExcel.EnableEvents = events And otherExcel.DisplayAlerts = alerts And otherExcel.Calculation = mode
            guard.Enter otherExcel: Set guard = Nothing
            TestAssertTrue stats, prefix & ".terminate", otherExcel.ScreenUpdating = screen And otherExcel.EnableEvents = events And otherExcel.DisplayAlerts = alerts And otherExcel.Calculation = mode
            Set guard = New CExcelAppStateGuard
        Next flags
    Next mode
    TestAssertTrue stats, "audit03.excelGuard.noSolve", SectionEquilibriumSolveCount() = solveCount
    TestAppendLine stats, "EXCEL_GUARD_CASES: failedEnterVariants=8; validEnterVariants=24; equilibriumCases=0"
    GoTo Cleanup
Failed:
    TestAssertTrue stats, "audit03.excelGuard.runtime." & CStr(Err.Number) & "." & Err.Description, False
Cleanup:
    On Error Resume Next
    If owned Then
        Err.Clear: currentStart = TestExcelProcessIdentity(otherExcel, checkedPID)
        If Err.Number = 0 And checkedPID = ownedPID And currentStart = ownedStart Then
            If Not guard Is Nothing Then guard.Restore
            Set guard = Nothing
            otherExcel.DisplayAlerts = False
            If Not book Is Nothing Then book.Close False
            otherExcel.Quit
        Else
            TestAssertTrue stats, "audit03.excelGuard.cleanup.identityChanged", False
        End If
    End If
    Set book = Nothing: Set otherExcel = Nothing
    On Error GoTo 0
End Sub


' -------------------- ДЛЯ ТЕСТОВ --------------------
' отказы каждого свойства при входе/выходе, включая ошибку 18.
Private Sub TestExcelGuardFaultRecovery(ByRef stats As TUiTestStats)
    Dim app As CTestExcelApplicationState, guard As CExcelAppStateGuard
    Dim name As Variant, code As Long, reason As String, source As String
    On Error GoTo Failed
    For Each name In Array("EnableCancelKey", "ScreenUpdating", "EnableEvents", "DisplayAlerts", "Calculation")
        Set app = New CTestExcelApplicationState: Set guard = New CExcelAppStateGuard
        app.GetterFault = CStr(name)
        On Error Resume Next
        Err.Clear: guard.Enter app
        code = Err.Number
        On Error GoTo Failed
        app.GetterFault = vbNullString
        TestAssertTrue stats, "guard.getter." & name, code = 1004 And TestFakeExcelRestored(app)

        app.FaultProperty = CStr(name): app.FaultsRemaining = 1: app.FaultNumber = 18
        On Error Resume Next
        Err.Clear: guard.Enter app
        code = Err.Number: source = Err.Source: reason = Err.Description
        On Error GoTo Failed
        TestAssertTrue stats, "guard.enter.originalInterrupt." & name, code = 18 And source = "CTestExcelApplicationState" And reason = "set " & name
        TestAssertTrue stats, "guard.enter.restored." & name, TestFakeExcelRestored(app)

        app.FaultsRemaining = 0: guard.Enter app
        app.FaultsRemaining = 1: guard.Restore True
        TestAssertTrue stats, "guard.restore.transient." & name, TestFakeExcelRestored(app)

        guard.Enter app
        app.FaultsRemaining = -1
        On Error Resume Next
        Err.Clear: guard.Restore True
        code = Err.Number
        On Error GoTo Failed
        TestAssertTrue stats, "guard.restore.persistentReported." & name, code = vbObjectError + 4181
        On Error Resume Next
        Err.Clear: guard.Enter app
        code = Err.Number
        On Error GoTo Failed
        TestAssertTrue stats, "guard.restore.noSnapshotOverwrite." & name, code = vbObjectError + 4181
        app.FaultsRemaining = 0: guard.Restore True
        TestAssertTrue stats, "guard.restore.laterRecovery." & name, TestFakeExcelRestored(app)
        guard.Restore True

        guard.Enter app
        app.FaultsRemaining = 1
        Set guard = Nothing
        TestAssertTrue stats, "guard.restore.destructor." & name, TestFakeExcelRestored(app)
    Next name
    Exit Sub
Failed:
    TestAssertTrue stats, "guard.faultRecovery.runtime." & CStr(Err.Number) & "." & Err.Description, False
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет восстановление всех заданных свойств имитируемого Excel после успешного выхода или ошибки.
Private Function TestFakeExcelRestored(ByVal app As CTestExcelApplicationState) As Boolean
    TestFakeExcelRestored = app.ScreenUpdating And app.EnableEvents And app.DisplayAlerts And _
        app.Calculation = xlCalculationAutomatic And app.EnableCancelKey = xlInterrupt
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Режим прерывания проверяется внутри выполняющего VBA экземпляра Excel:
' чужой Application между COM-вызовами простаивает и сбрасывает его в xlInterrupt.
Private Sub TestExcelGuardLocalCancellation(ByRef stats As TUiTestStats)
    Dim app As Object, guard As CExcelAppStateGuard, mode As Variant, saved As Long
    Set app = ThisWorkbook.Application: saved = app.EnableCancelKey
    On Error GoTo Failed
    TestAppendLine stats, "GUARD_LOCAL_HOST: globalHwnd=" & CStr(Application.Hwnd) & "; workbookHwnd=" & CStr(app.Hwnd) & "; books=" & CStr(app.Workbooks.Count)
    For Each mode In Array(xlDisabled, xlInterrupt, xlErrorHandler)
        app.EnableCancelKey = mode
        Set guard = New CExcelAppStateGuard
        guard.Enter app
        TestAssertTrue stats, "guard.localCancel.handled." & CStr(mode), app.EnableCancelKey = xlErrorHandler
        guard.Restore True
        TestAssertTrue stats, "guard.localCancel.restored." & CStr(mode), app.EnableCancelKey = mode
    Next mode
    GoTo Done
Failed:
    TestAssertTrue stats, "guard.localCancel.runtime." & CStr(Err.Number) & "." & Err.Description, False
Done:
    If Not guard Is Nothing Then guard.Restore
    app.EnableCancelKey = saved
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' штатный расчёт и отказ ввода проходят тот же общий cleanup.
Public Function RunExcelGuardWorkflowTests() As String
    Dim stats As TUiTestStats, app As Object, code As Long, message As String
    Dim oldCancel As Long
    On Error GoTo Failed
    Set app = ThisWorkbook.Application
    oldCancel = app.EnableCancelKey
    TestAppendLine stats, "GUARD_WORKFLOW_HOST: globalHwnd=" & CStr(Application.Hwnd) & "; workbookHwnd=" & CStr(app.Hwnd) & "; books=" & CStr(app.Workbooks.Count) & "; name=" & ThisWorkbook.Name
    TestSetSystemSetting "Circle.Diameter", "400" ' Фиксированный небольшой пример, независимый от пользовательской сетки.
    TestSetSystemSetting "Mesh.StepMode", "User"
    TestSetSystemSetting "Mesh.StepX", "50": TestSetSystemSetting "Mesh.StepY", "50"
    TestSetSystemSetting "Mesh.BoundarySubdivisions", "1"
    TestSetSystemSetting "Rebar.Diameter2", "0": TestSetSystemSetting "Rebar.Diameter3", "0"
    TestExcelApplicationStateGuardRestoresSettings stats
    TestAssertTrue stats, "guard.workflow.cancelRestored", app.EnableCancelKey = oldCancel
    Dim oldEvents As Boolean, oldScreen As Boolean, oldAlerts As Boolean, oldCalculation As Long
    oldEvents = app.EnableEvents: oldScreen = app.ScreenUpdating
    oldAlerts = app.DisplayAlerts: oldCalculation = app.Calculation
    TestSetSystemSetting "Geometry.Type", "Invalid.Guard.Test"
    On Error Resume Next
    Err.Clear: message = FormSectionForWorkbook(ThisWorkbook)
    code = Err.Number
    On Error GoTo Failed
    TestAssertTrue stats, "guard.workflow.invalidInput", code <> 0
    TestAssertTrue stats, "guard.workflow.failureRestored", app.EnableEvents = oldEvents And _
        app.ScreenUpdating = oldScreen And app.DisplayAlerts = oldAlerts And _
        app.Calculation = oldCalculation And app.EnableCancelKey = oldCancel
    GoTo Done
Failed:
    TestAssertTrue stats, "guard.workflow.runtime." & CStr(Err.Number) & "." & Err.Description, False
Done:
    RunExcelGuardWorkflowTests = stats.Report & "TOTAL_EXCEL_GUARD_WORKFLOW: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' ================== ДЛЯ ТЕСТОВ: СПИСКИ И РАМКА CONFIG ==================

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет действующие списки в сохраненной книге, а не только формулы
' validation из сборщика. Каждое значение источника должно приниматься
' Excel как отдельный вариант; исходная формула ячейки возвращается сразу.
Public Function RunConfigPresentationTests() As String
    Dim stats As TUiTestStats
    TestConfigDropdownChoices stats
    TestConfigDropdownMeaning stats
    TestConfigSettingsRightBorder stats
    TestRectanglesIndependentSelectorLayout stats
    TestAutoCADCommonSettingsLayout stats
    TestAutoCADGeometryContourWarning stats
    RunConfigPresentationTests = stats.Report & "TOTAL_CONFIG_PRESENTATION: passed=" & _
        CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет предупреждение в инструкции подготовки геометрии, включая
' риски завышения/занижения площади и красное выделение всего абзаца.
' Адреса не фиксированы: после пересборки справки ищутся ее заголовки.
Private Sub TestAutoCADGeometryContourWarning(ByRef stats As TUiTestStats)
    Dim guide As Object, warning As Object, heading As Object, nextHeading As Object
    Dim row As Long, text As String, cell As Object, used As Object, values As Variant, value As String, noteText As String
    On Error GoTo Failed
    Set guide = ThisWorkbook.Worksheets("Справка")
    Set used = guide.UsedRange: values = used.Value2 ' Читаем и скрытые строки одним массивом, независимо от настроек Find.
    For row = 1 To UBound(values, 1)
        value = CStr(values(row, 1))
        If InStr(1, value, "ВНИМАНИЕ: для расчета раскрытия трещин", vbTextCompare) = 1 Then Set warning = guide.Cells(used.Row + row - 1, used.Column)
        If value = "3. Точный контур и несколько отверстий" Then Set heading = guide.Cells(used.Row + row - 1, used.Column)
        If value = "4. Импорт и контроль модели" Then Set nextHeading = guide.Cells(used.Row + row - 1, used.Column)
        If Left$(value, Len("2. Задан только наружный контур.")) = "2. Задан только наружный контур." Then
            Dim noteRow As Long
            For noteRow = row To UBound(values, 1)
                If Left$(CStr(values(noteRow, 1)), Len("3. Заданы наружный контур")) = "3. Заданы наружный контур" Then Exit For
                noteText = noteText & " " & CStr(values(noteRow, 1))
            Next noteRow
        End If
    Next row
    TestAssertTrue stats, "guide.autoCAD.contourWarning.exists", Not warning Is Nothing And Not heading Is Nothing And Not nextHeading Is Nothing
    If warning Is Nothing Or heading Is Nothing Or nextHeading Is Nothing Then Exit Sub
    TestAssertTrue stats, "guide.autoCAD.contourWarning.location", warning.Row > heading.Row And warning.Row < nextHeading.Row
    For row = warning.Row To nextHeading.Row - 1
        Set cell = guide.Cells(row, 1)
        If CLng(cell.Font.Color) <> 255 Then Exit For
        text = text & " " & CStr(cell.Value2)
        TestAssertTrue stats, "guide.autoCAD.contourWarning.redBold." & CStr(row), CBool(cell.Font.Bold)
    Next row
    TestAssertTrue stats, "guide.autoCAD.contourWarning.recommendation", InStr(1, text, "настоятельно рекомендуется", vbTextCompare) > 0
    TestAssertTrue stats, "guide.autoCAD.contourWarning.armature", InStr(1, text, "площадь взаимодействия для трещин", vbTextCompare) > 0
    TestAssertTrue stats, "guide.autoCAD.contourWarning.risks", InStr(1, text, "завышены или занижены", vbTextCompare) > 0
    ' Четыре правила описываются один раз в примечаниях первого раздела,
    ' а не дублируются в подразделе точного контура.
    TestAssertTrue stats, "guide.autoCAD.knownOuter.solid", InStr(1, noteText, "сплошной бетон", vbTextCompare) > 0
    TestAssertTrue stats, "guide.autoCAD.knownOuter.noMeshVoids", InStr(1, noteText, "пустоты сетки их не заменят", vbTextCompare) > 0
    Exit Sub
Failed:
    TestAssertTrue stats, "guide.autoCAD.contourWarning.runtime." & CStr(Err.Number) & "." & Err.Description, False
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет единственный набор общих слоев и размещение экспортной области
' в каталоге. Ожидания независимы от SettingsCatalog и действуют после сборки.
Private Sub TestAutoCADCommonSettingsLayout(ByRef stats As TUiTestStats)
    Dim range As Object, values As Variant, row As Long, key As String, subsection As String, section As String
    Dim counts As Object, keyRows As Object, expected As Variant
    On Error GoTo Failed
    Set range = ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange
    values = range.Value2: Set counts = CreateObject("Scripting.Dictionary"): Set keyRows = CreateObject("Scripting.Dictionary")
    For row = 2 To range.Rows.Count
        key = CStr(values(row, 1))
        TestAssertTrue stats, "config.subsection.russian." & CStr(row), key <> "[Common]"
        If key = "[AutoCAD]" Then
            section = "AutoCAD": subsection = vbNullString
        ElseIf Left$(key, 1) = "[" Then
            subsection = key
        ElseIf Left$(key, 8) = "AutoCAD." Then
            If Not counts.Exists(key) Then counts.Add key, 0
            counts(key) = CLng(counts(key)) + 1: keyRows(key) = row
            If Left$(key, 15) = "AutoCAD.Common." Then
                TestAssertTrue stats, "config.autoCAD.common." & key, section = "AutoCAD" And subsection = "[Общие настройки]"
            ElseIf key = "AutoCAD.Export.CrackInteractionLayer" Then
                TestAssertTrue stats, "config.autoCAD.crack.exportOnly", subsection = "[Экспорт]"
            End If
        End If
    Next row
    For Each expected In Array("AutoCAD.Common.ConcreteLayer", "AutoCAD.Common.RebarLayer", _
            "AutoCAD.Common.SectionContourLayer", "AutoCAD.Common.OpeningContourLayer", "AutoCAD.Export.CrackInteractionLayer")
        TestAssertTrue stats, "config.autoCAD.unique." & CStr(expected), counts.Exists(CStr(expected))
        If counts.Exists(CStr(expected)) Then TestAssertTrue stats, "config.autoCAD.once." & CStr(expected), CLng(counts(CStr(expected))) = 1
    Next expected
    For Each expected In Array("AutoCAD.Import.ConcreteLayer", "AutoCAD.Import.RebarLayer", _
            "AutoCAD.Layer.Concrete", "AutoCAD.Layer.Rebar", "AutoCAD.Common.CrackInteractionLayer")
        TestAssertTrue stats, "config.autoCAD.removed." & CStr(expected), Not counts.Exists(CStr(expected))
    Next expected
    If keyRows.Exists("AutoCAD.Export.CrackInteractionLayer") And keyRows.Exists("AutoCAD.Layer.ConcreteTension") Then _
        TestAssertTrue stats, "config.autoCAD.crack.beforeConcreteTension", keyRows("AutoCAD.Export.CrackInteractionLayer") + 1 = keyRows("AutoCAD.Layer.ConcreteTension")
    Exit Sub
Failed:
    TestAssertTrue stats, "config.autoCAD.layout.runtime." & CStr(Err.Number) & "." & Err.Description, False
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет смысл вариантов, а не только способность Excel открыть список.
' Независимые ожидаемые наборы не позволяют пройти тесту при подмене
' Stacked/SideBySide значениями mm/pt или при смешении иных настроек.
Private Sub TestConfigDropdownMeaning(ByRef stats As TUiTestStats)
    Dim table As Object, cell As Object, source As Object, name As Variant, key As String, expected As String, actual As String
    Dim row As Long, column As Long, headerRow As Long, checked As Long, formula As String, validationType As Long
    On Error GoTo Failed
    For Each name In Array("rngSystemSettings", "rngUnitSettings", "rngSignConventionSettings", _
            "rngCalculationProfiles", "rngPlotAnnotationSettings", "rngCircleGeometry", _
            "rngRectanglesGeometry", "rngPointedSectionGeometry", "rngHollowRectangleGeometry", "rngLoadCombinations")
        Set table = ThisWorkbook.Names.Item(CStr(name)).RefersToRange
        For row = 2 To table.Rows.Count
            For column = 2 To table.Columns.Count
                Set cell = table.Cells(row, column)
                validationType = 0
                On Error Resume Next
                validationType = cell.Validation.Type
                On Error GoTo Failed
                If validationType <> xlValidateList Then GoTo NextMeaningCell
                key = CStr(table.Cells(row, 1).Value2)
                If CStr(name) = "rngCalculationProfiles" Then key = CStr(table.Cells(row, 2).Value2)
                expected = TestExpectedConfigDropdownOptions(key)
                If CStr(name) = "rngLoadCombinations" Then
                    If column = 6 Then expected = "Auto|" & ChrW(&H3BB) & "*Mx|" & ChrW(&H3BB) & "*My|" & ChrW(&H3BB) & "*Mxy|" & ChrW(&H3BB) & "*N|" & ChrW(&H3BB) & "*NMxy"
                End If
                If InStr(1, CStr(name), "Geometry", vbTextCompare) > 0 And Len(expected) = 0 Then
                    For headerRow = row - 1 To 1 Step -1
                        key = LCase$(CStr(table.Cells(headerRow, column).Value2))
                        If key = "положение" Then expected = "Stacked|SideBySide": Exit For
                        If key = "привязка" Then expected = "EachBar|EverySecondBar": Exit For
                    Next headerRow
                End If
                If CStr(name) Like "rng*Geometry" Then
                    For headerRow = row - 1 To 1 Step -1
                        key = LCase$(CStr(table.Cells(headerRow, column).Value2))
                        If key = "pos2" Or key = "pos3" Then expected = "Stacked|SideBySide": Exit For
                        If key = "bind2" Or key = "bind3" Then expected = "EachBar|EverySecondBar": Exit For
                    Next headerRow
                End If
                If Len(expected) > 0 Then
                    actual = Join(TestValidationCellOptions(cell), "|")
                    TestAssertTextEquals stats, "config.dropdown.meaning." & CStr(name) & "." & CStr(row) & "." & CStr(column), actual, expected
                    checked = checked + 1
                End If
NextMeaningCell:
            Next column
        Next row
    Next name
    TestAssertTrue stats, "config.dropdown.meaning.coverage", checked > 150
    TestAppendLine stats, "CONFIG_DROPDOWN_MEANING: checked=" & CStr(checked)
    Exit Sub
Failed:
    TestAssertTrue stats, "config.dropdown.meaning.runtime." & CStr(Err.Number) & "." & Err.Description, False
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает независимый oracle допустимых вариантов конкретной настройки.
' Динамические профили и LC проверяются отдельно; здесь нет адресов helper-ячеек.
Private Function TestExpectedConfigDropdownOptions(ByVal key As String) As String
    Select Case key
        Case "Geometry.Source": TestExpectedConfigDropdownOptions = "Generated|AutoCAD"
        Case "General.WorstCombinationCriterion": TestExpectedConfigDropdownOptions = "StrengthStrain|StrengthCapacity|Cracks|Stability"
        Case "Stability.Code", "SLS.Crack.Code": TestExpectedConfigDropdownOptions = "SP63|SP35"
        Case "Stability.SystemType": TestExpectedConfigDropdownOptions = "Determinate|Indeterminate"
        Case "Stability.ZeroMomentEccentricitySign1", "Stability.ZeroMomentEccentricitySign2": TestExpectedConfigDropdownOptions = "1|-1"
        Case "Stability.AccidentalEccentricityMode": TestExpectedConfigDropdownOptions = "AutoWithL|AutoWith" & ChrW(&H3BC) & "L|User"
        Case "Stability.AccidentalEccentricityPlanes": TestExpectedConfigDropdownOptions = "OnlyMomentPlane|BothPlanes"
        Case "Stability.PhiLMode": TestExpectedConfigDropdownOptions = "Auto|PhiL2"
        Case "Geometry.Type": TestExpectedConfigDropdownOptions = Join(RegisteredSectionRecipeTypes(), "|")
        Case "Solver.Method": TestExpectedConfigDropdownOptions = "Newton|Secant"
        Case "Capacity.SolutionStrategy", "SLS.Crack.InitiationSolutionStrategy": TestExpectedConfigDropdownOptions = "Auto|UltimateStrain|LoadMultiplier"
        Case "Capacity.SearchMethod": TestExpectedConfigDropdownOptions = "Bisection|Brent|Secant"
        Case "SLS.Crack.SP35.RadiusDiameterMode": TestExpectedConfigDropdownOptions = "Max|Min|Average"
        Case "SLS.Crack.SP35.InteractionRadiusMode": TestExpectedConfigDropdownOptions = "3d|5d|6d"
        Case "SLS.Crack.SP35.RebarProfile": TestExpectedConfigDropdownOptions = "Periodic|Smooth"
        Case "SLS.Crack.Phi3Mode": TestExpectedConfigDropdownOptions = "Auto|User"
        Case "SLS.Crack.PsiMode": TestExpectedConfigDropdownOptions = "User|Auto|AlwaysCalc"
        Case "SLS.Crack.SigmaSCrcAveragingMode": TestExpectedConfigDropdownOptions = "AllSelected|TensionOnly"
        Case "SLS.Crack.TensionZoneMode": TestExpectedConfigDropdownOptions = "Effective|FullTension"
        Case "SLS.Crack.CoverDistanceMode": TestExpectedConfigDropdownOptions = "NearestContour|GlobalExtreme"
        Case "AutoCAD.Export.LabelMode": TestExpectedConfigDropdownOptions = "ValuesOnly|NamesAndValues"
        Case "AutoCAD.Export.ContourFormat": TestExpectedConfigDropdownOptions = "Region|Polyline"
        Case "AutoCAD.Export.PrincipalAxesMode", "Plot.PrincipalAxesMode": TestExpectedConfigDropdownOptions = "Transformed|Concrete|None"
        Case "Plot.LegendMode": TestExpectedConfigDropdownOptions = "Separate|Common"
        Case "General.ExecutionReportEnabled", "General.NonCriticalMessagesEnabled", "General.ReformatResultsEnabled", "General.DiagramExtension", _
                "AutoCAD.Export.NeutralLineEnabled", "AutoCAD.Export.LoadPointEnabled", "AutoCAD.Export.ContourEnabled", _
                "AutoCAD.Export.ExportCrackInteractionContour", "Plot.Enabled", "Plot.AutoUpdateAfterCalculation", _
                "Plot.ResultGradient", "Plot.ResultLabelsEnabled", "Plot.NeutralLineEnabled", "Plot.LoadApplicationPointEnabled", _
                "Plot.AxisLabelsEnabled", "Plot.ContourEnabled", "Plot.LegendEnabled", "Calculation.Strength.DirectState", _
                "Calculation.Strength.Capacity", "Calculation.Crack.Width", "Calculation.Stability.Enabled", "Enabled", "LineEnabled"
            TestExpectedConfigDropdownOptions = "Yes|No"
        Case "Length": TestExpectedConfigDropdownOptions = "mm|cm|m"
        Case "Area": TestExpectedConfigDropdownOptions = "mm2|cm2|m2"
        Case "Force": TestExpectedConfigDropdownOptions = "N|kN|tf"
        Case "Moment": TestExpectedConfigDropdownOptions = "N*mm|kN*m|tf*m"
        Case "Stress": TestExpectedConfigDropdownOptions = "Pa|kPa|MPa|kgf/cm2|tf/m2"
        Case "Curvature": TestExpectedConfigDropdownOptions = "1/mm|1/m"
        Case "+N": TestExpectedConfigDropdownOptions = "Tension|Compression"
        Case "+Mx": TestExpectedConfigDropdownOptions = "+Y tension|-Y tension"
        Case "+My": TestExpectedConfigDropdownOptions = "+X tension|-X tension"
        Case "MaterialModel.Stability.ValueSet", "MaterialModel.Strength.ValueSet", "MaterialModel.CrackInitiation.ValueSet", "MaterialModel.CrackedState.ValueSet"
            TestExpectedConfigDropdownOptions = "ULS(I)|SLS(II)"
        Case "MaterialModel.Strength.ConcreteDiagram", "MaterialModel.CrackInitiation.ConcreteDiagram", "MaterialModel.CrackedState.ConcreteDiagram", _
                "MaterialModel.Strength.SteelDiagram", "MaterialModel.CrackInitiation.SteelDiagram", "MaterialModel.CrackedState.SteelDiagram"
            TestExpectedConfigDropdownOptions = "TwoLine|ThreeLine"
        Case "MaterialModel.Strength.ConcreteTension": TestExpectedConfigDropdownOptions = "Ignore|UseDiagram"
        Case "Visualization.State": TestExpectedConfigDropdownOptions = "StrengthState|CapacityState|PreCrackState|PostCrackState|CrackedState"
        Case "Visualization.Quantity": TestExpectedConfigDropdownOptions = "Stress|Strain"
        Case "Placement": TestExpectedConfigDropdownOptions = "Outside|Inside"
        Case "TextUnits": TestExpectedConfigDropdownOptions = "mm|pt"
        Case "ArrowType": TestExpectedConfigDropdownOptions = "Triangle|Stealth|Diamond|Oval|Open"
        Case "ArrowSize": TestExpectedConfigDropdownOptions = "Small|Medium|Wide"
        Case "Rebar.Loc2row", "Rebar.Loc3row": TestExpectedConfigDropdownOptions = "Stacked|SideBySide"
        Case "Rebar.Bind2row", "Rebar.Bind3row": TestExpectedConfigDropdownOptions = "EachBar|EverySecondBar"
        Case "Rectangles.SectionType": TestExpectedConfigDropdownOptions = "TwoRect|OneRect"
    End Select
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет все list-ячейки Config и по одному представителю каждого
' источника. Ссылочные списки не зависят от локального разделителя Excel;
' проверка Validation.Value выявляет ситуацию, когда вся строка стала пунктом.
Private Sub TestConfigDropdownChoices(ByRef stats As TUiTestStats)
    Dim sheet As Object, cells As Object, cell As Object, source As Object, optionCell As Object
    Dim checked As Object, formula As String, prefix As String, saved As Variant
    Dim sample As Object, options As Long, lists As Long, targets As Long
    On Error GoTo Failed
    Set sheet = ThisWorkbook.Worksheets.Item("Config")
    Set cells = sheet.Cells.SpecialCells(xlCellTypeAllValidation)
    Set checked = CreateObject("Scripting.Dictionary")
    For Each cell In cells.Cells
        If cell.MergeCells Then
            If cell.Address <> cell.MergeArea.Cells(1, 1).Address Then GoTo NextCell
        End If
        If cell.Validation.Type = xlValidateList Then
            targets = targets + 1
            formula = CStr(cell.Validation.Formula1)
            prefix = "config.dropdown." & cell.Address(False, False)
            TestAssertTrue stats, prefix & ".validSource", Len(formula) > 0
            TestAssertTrue stats, prefix & ".visible", cell.Validation.InCellDropdown
            If Left$(formula, 1) <> "=" And Not checked.Exists(formula) Then
                checked.Add formula, True
                saved = cell.Formula
                Dim inlineOption As Variant
                For Each inlineOption In TestValidationCellOptions(cell)
                    cell.Value2 = CStr(inlineOption)
                    TestAssertTrue stats, prefix & ".accepts." & CStr(inlineOption), cell.Validation.Value
                Next inlineOption
                cell.Value2 = "__INVALID_DROPDOWN_OPTION__"
                TestAssertTrue stats, prefix & ".rejectsUnknown", Not cell.Validation.Value
                cell.Formula = saved
                lists = lists + 1
            End If
            If Left$(formula, 1) = "=" And Not checked.Exists(formula) Then
                checked.Add formula, True
                Set source = sheet.Range(Mid$(formula, 2))
                TestAssertTrue stats, prefix & ".vertical", source.Columns.Count = 1
                Set sample = cell
                saved = sample.Formula
                options = 0
                For Each optionCell In source.Cells
                    If Not IsError(optionCell.Value2) Then
                        If Len(CStr(optionCell.Value2)) > 0 Then
                            options = options + 1
                            sample.Value2 = optionCell.Value2
                            TestAssertTrue stats, prefix & ".accepts." & CStr(options), sample.Validation.Value
                        End If
                    Else
                        TestAssertTrue stats, prefix & ".sourceError", False
                    End If
                Next optionCell
                sample.Value2 = "__INVALID_DROPDOWN_OPTION__"
                TestAssertTrue stats, prefix & ".rejectsUnknown", Not sample.Validation.Value
                sample.Formula = saved
                Set sample = Nothing
                TestAssertTrue stats, prefix & ".notEmpty", options > 0
                lists = lists + 1
            End If
        End If
NextCell:
    Next cell
    TestAssertTrue stats, "config.dropdown.targetsPresent", targets > 0
    TestAppendLine stats, "CONFIG_DROPDOWNS: targets=" & CStr(targets) & "; rangeSources=" & CStr(lists)
    Exit Sub
Failed:
    Dim code As Long, reason As String
    code = Err.Number: reason = Err.Description
    On Error Resume Next
    If Not sample Is Nothing Then sample.Formula = saved
    On Error GoTo 0
    TestAssertTrue stats, "config.dropdown.runtime." & CStr(code) & "." & reason, False
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет правую грань каждой строки общей таблицы, включая последнюю
' колонку Справка. Адреса берутся из имени, а не из фиксированных букв/строк.
Private Sub TestConfigSettingsRightBorder(ByRef stats As TUiTestStats)
    Dim settings As Object, cell As Object, row As Long
    On Error GoTo Failed
    Set settings = ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange
    For row = 1 To settings.Rows.Count
        Set cell = settings.Cells(row, settings.Columns.Count)
        If cell.MergeCells Then Set cell = cell.MergeArea
        TestAssertTrue stats, "config.settings.rightBorder." & CStr(row), _
            cell.Borders(xlEdgeRight).LineStyle = xlDash And _
            cell.Borders(xlEdgeRight).Weight = xlThin And _
            cell.Borders(xlEdgeRight).Color = RGB(0, 0, 0)
    Next row
    Exit Sub
Failed:
    TestAssertTrue stats, "config.settings.rightBorder.runtime." & CStr(Err.Number) & "." & Err.Description, False
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' проверка ожиданий новой структуры Config до долгого расчётного набора.
Public Function RunSharedConfigPreflightTests() As String
    Dim stats As TUiTestStats, config As Object, saved As Variant, passed As Long, failed As Long
    On Error GoTo FailedRun
    TestSharedConfigProgress "START", stats
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = TestFixtureSaveInputs(config)
    TestSharedConfigProgress "ComparisonContract", stats
    stats.Report = stats.Report & RunConfigComparisonContractTests(passed, failed)
    stats.Passed = stats.Passed + passed: stats.Failed = stats.Failed + failed
    TestSharedConfigProgress "DropdownChoices", stats
    TestConfigDropdownChoices stats
    TestSharedConfigProgress "DropdownMeaning", stats
    TestConfigDropdownMeaning stats
    TestSharedConfigProgress "SettingsRightBorder", stats
    TestConfigSettingsRightBorder stats
    TestSharedConfigProgress "IndependentSelectors", stats
    TestRectanglesIndependentSelectors stats
    TestSharedConfigProgress "IndependentSelectorLayout", stats
    TestRectanglesIndependentSelectorLayout stats
    TestSharedConfigProgress "IndependentSelectorEffects", stats
    TestRectanglesIndependentSelectorEffects stats
    TestSharedConfigProgress "IndependentThirdRows", stats
    TestRectanglesIndependentThirdRows stats
    TestSharedConfigProgress "ShapeConfig", stats
    stats.Report = stats.Report & modTestGeometryConfig.RunAudit03ShapeConfigTests(passed, failed)
    stats.Passed = stats.Passed + passed: stats.Failed = stats.Failed + failed
    TestSharedConfigProgress "RebarCounter", stats
    stats.Report = stats.Report & modTestGeometryConfig.RunAudit03RebarCounterTests(passed, failed)
    stats.Passed = stats.Passed + passed: stats.Failed = stats.Failed + failed
    TestSharedConfigProgress "RebarFieldConfig", stats
    stats.Report = stats.Report & modTestGeometryConfig.RunAudit03RebarFieldConfigTests(passed, failed)
    stats.Passed = stats.Passed + passed: stats.Failed = stats.Failed + failed
    GoTo Finished
FailedRun:
    stats.Failed = stats.Failed + 1
    TestAppendLine stats, "FAIL: config.preflight; " & CStr(Err.Number) & "; " & Err.Description
Finished:
    TestSharedConfigProgress "RestoreInputs", stats
    If Not config Is Nothing And IsArray(saved) Then TestFixtureRestoreInputs config, saved
    TestSharedConfigProgress "FINISHED", stats
    RunSharedConfigPreflightTests = stats.Report & "TOTAL_SHARED_CONFIG_PREFLIGHT: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' сохраняет завершённые проверки и начало следующего этапа.
' Журнал помогает отличить долгую проверку от остановки; отметка этапа не означает PASS.
Private Sub TestSharedConfigProgress(ByVal stage As String, ByRef stats As TUiTestStats)
    On Error Resume Next
    Dim fso As Object, stream As Object, line As String, timeline As String
    line = Format$(Now, "yyyy-mm-dd hh:nn:ss") & " | " & stage & _
        " | passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set stream = fso.CreateTextFile(ThisWorkbook.Path & "\Config_Preflight_Progress.txt", True, True)
    stream.Write stats.Report & line & vbCrLf: stream.Close
    timeline = ThisWorkbook.Path & "\Config_Preflight_Timeline.txt"
    If stage = "START" Then
        Set stream = fso.CreateTextFile(timeline, True, True)
    Else
        Set stream = fso.OpenTextFile(timeline, 8, True, -1)
    End If
    stream.WriteLine line: stream.Close
    On Error GoTo 0
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' противоположные направления одной главной оси меняют знак её проекций.
Private Function TestAudit03PrincipalAxisSign(ByRef baseline As Variant, ByRef actual As Variant, Optional ByVal sectionPrefix As String = "Transformed") As Double
    Dim row As Long, delta As Double
    TestAudit03PrincipalAxisSign = 1#
    For row = 2 To UBound(baseline, 1)
        If CStr(baseline(row, 3)) = sectionPrefix & ".PrincipalAngle" Then
            delta = CDbl(actual(row, 4)) - CDbl(baseline(row, 4))
            If Abs(Sin(delta)) <= 0.000000001 And Cos(delta) < 0# Then TestAudit03PrincipalAxisSign = -1#
            Exit Function
        End If
    Next row
    Err.Raise 5, , "В контрольном снимке отсутствует угол главных осей приведённого сечения."
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' читает направленное ядровое расстояние из фактического снимка.
Private Function TestAudit03PropertyNumber(ByRef values As Variant, ByVal key As String) As Double
    Dim row As Long
    For row = 2 To UBound(values, 1)
        If CStr(values(row, 3)) = key Then TestAudit03PropertyNumber = CDbl(values(row, 4)): Exit Function
    Next row
    Err.Raise 5, , "В контрольном снимке отсутствует параметр " & key
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' причина и все числа комментария сохраняются; допускается только ошибка Double.
Private Function TestAudit03SameStabilityComment(ByVal expected As String, ByVal actual As String) As Boolean
    On Error GoTo Different
    If expected = actual Then TestAudit03SameStabilityComment = True: Exit Function
    Dim expression As Object, first As Object, second As Object, i As Long, reference As Double
    Set expression = CreateObject("VBScript.RegExp")
    expression.Global = True: expression.Pattern = "[-+]?[0-9]+([.][0-9]+)?([Ee][-+]?[0-9]+)?"
    If expression.Replace(expected, "#") <> expression.Replace(actual, "#") Then Exit Function
    Set first = expression.Execute(expected): Set second = expression.Execute(actual)
    If first.Count <> second.Count Then Exit Function
    For i = 0 To first.Count - 1
        reference = CDbl(Replace(first(i).Value, ".", Application.International(xlDecimalSeparator)))
        If Abs(CDbl(Replace(second(i).Value, ".", Application.International(xlDecimalSeparator))) - reference) > _
                0.000000000001 * (1# + Abs(reference)) Then Exit Function
    Next i
    TestAudit03SameStabilityComment = True
Different:
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' сравнение не скрывает смену причины, статуса или значимого числа.
Public Function RunConfigComparisonContractTests(Optional ByRef passed As Long = 0, Optional ByRef failed As Long = 0) As String
    Dim stats As TUiTestStats, first As Variant, second As Variant
    TestAssertTrue stats, "comparison.comment.roundoff", TestAudit03SameStabilityComment("Запас = 0.86662649246776; eta = 5.20101263390298.", "Запас = 0.866626492467761; eta = 5.20101263390296.")
    TestAssertTrue stats, "comparison.comment.numberChange", Not TestAudit03SameStabilityComment("eta = 5.201", "eta = 5.202")
    TestAssertTrue stats, "comparison.comment.reasonChange", Not TestAudit03SameStabilityComment("Проходит: 1", "Не проходит: 1")
    TestAssertTrue stats, "comparison.comment.status", Not TestAudit03SameStabilityComment("OK", "NumFail")
    ReDim first(1 To 2, 1 To 4): ReDim second(1 To 2, 1 To 4)
    first(2, 3) = "Transformed.PrincipalAngle": second(2, 3) = first(2, 3)
    first(2, 4) = -GEOM_PI / 2#: second(2, 4) = GEOM_PI / 2#
    TestAssertTrue stats, "comparison.axis.opposite", TestAudit03PrincipalAxisSign(first, second) = -1#
    second(2, 4) = 0#
    TestAssertTrue stats, "comparison.axis.swapNotReversed", TestAudit03PrincipalAxisSign(first, second) = 1#
    passed = stats.Passed: failed = stats.Failed
    RunConfigComparisonContractTests = stats.Report & "TOTAL_CONFIG_COMPARISON_CONTRACT: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' читает варианты короткого списка либо значения его диапазона.
Private Function TestValidationCellOptions(ByVal cell As Object) As Variant
    Dim formula As String, source As Object, item As Object, values() As String, i As Long
    formula = CStr(cell.Validation.Formula1)
    If Left$(formula, 1) <> "=" Then
        TestValidationCellOptions = Split(formula, Application.International(xlListSeparator))
        Exit Function
    End If
    Set source = cell.Worksheet.Range(Mid$(formula, 2))
    ReDim values(0 To source.Cells.Count - 1)
    For Each item In source.Cells
        values(i) = CStr(item.Value2): i = i + 1
    Next item
    TestValidationCellOptions = values
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' проверяет таблицу сочетаний после переноса, не меняя её данных.
Public Function RunSavedLoadTableLayoutTests() As String
    Dim stats As TUiTestStats
    TestLoadCombinationsOnConfig stats
    RunSavedLoadTableLayoutTests = stats.Report & "TOTAL_LOAD_TABLE_LAYOUT: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' находит шапку контрольных диаграмм независимо от положения таблицы.
Private Function TestMaterialDiagramControlsHeadingExists(ByVal sheet As Object) As Boolean
    Dim heading As Object
    Set heading = sheet.UsedRange.Find(What:="Контрольные точки диаграмм", LookIn:=xlValues, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
    TestMaterialDiagramControlsHeadingExists = Not heading Is Nothing
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' повторяет сценарии восьми устаревших проверок после изменения таблиц и сообщений.
' В этом модуле находятся пять проверок; оставшиеся три выполняет GeometryReview.
Public Function RunCorrectedUiExpectationsTests() As String
    Dim stats As TUiTestStats, config As Object, saved As Variant
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = TestFixtureSaveInputs(config)
    On Error GoTo Failed
    modTestShapeArchitecture.TestConfigureShapeFixture "LSection" ' Готовим геометрию и ряды явно, независимо от сохранённых пользовательских значений.
    modTestShapeArchitecture.TestConfigureShapeFixture "RoundedSimple"
    TestFixtureResetReinforcement "Circle"
    TestSetSystemSetting "Circle.Diameter", "300"
    TestCircleWorkbookRunWritesResults stats
    TestRectanglesWorkbookRunWritesResults stats
    TestCapacitySearchMethodValidation stats
    TestRebarInputValidationDoesNotUseHiddenDefaults stats
    GoTo Restore
Failed:
    TestAssertTrue stats, "ui.corrected.runtime; " & CStr(Err.Number) & "; " & Err.Description, False
Restore:
    On Error GoTo FailedRestore
    TestFixtureRestoreInputs config, saved
    GoTo Finished
FailedRestore:
    TestAssertTrue stats, "ui.corrected.restore; " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    RunCorrectedUiExpectationsTests = stats.Report & "TOTAL_CORRECTED_UI_EXPECTATIONS: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' ДЛЯ ТЕСТОВ
' Направленная проверка пяти команд без длительного перебора всех полей Config.
Public Function RunSectionWorkflowUiTests() As String
    Dim stats As TUiTestStats, saved As Variant
    saved = TestFixtureSaveInputs(ThisWorkbook.Worksheets("Config").UsedRange)
    On Error GoTo Failed
    TestButtons stats
    TestCalculationRequiresFormedSection stats
    TestFormButtonBuildsGeneratedSection stats
Done:
    TestFixtureRestoreInputs ThisWorkbook.Worksheets("Config").UsedRange, saved
    TestAppendLine stats, "TOTAL_WORKFLOW_UI: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunSectionWorkflowUiTests = stats.Report
    Exit Function
Failed:
    TestAssertTrue stats, "unexpected.workflow." & Err.Number & "." & Err.Description, False
    Resume Done
End Function


' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает весь блок без служебной даты запуска. Читает уже готовые массивы,
' поэтому размер сечения не увеличивает число обращений к Excel.
Private Function TestWorkflowSameTable(ByVal before As Variant, ByVal after As Variant) As Boolean
    On Error GoTo Different
    If UBound(before, 1) <> UBound(after, 1) Or UBound(before, 2) <> UBound(after, 2) Then Exit Function
    Dim row As Long, column As Long
    For row = 1 To UBound(before, 1)
        For column = 2 To UBound(before, 2)
            If IsError(before(row, column)) Or IsError(after(row, column)) Then Exit Function
            If IsNumeric(before(row, column)) And IsNumeric(after(row, column)) Then
                If CDbl(before(row, column)) <> CDbl(after(row, column)) Then Exit Function
            ElseIf CStr(before(row, column)) <> CStr(after(row, column)) Then
                Exit Function
            End If
        Next column
    Next row
    TestWorkflowSameTable = True
Different:
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Доказывает, что расчёт использует снимок, а не изменённые настройки генератора
' или чертёж; повторные расчёты и смена OUTPUT не перезаписывают координаты.
Private Sub TestWorkflowFrozenGeometry(ByRef stats As TUiTestStats)
    TestPrepareCircleInput
    TestSetSystemSetting "Circle.Diameter", "300"
    TestFixtureResetReinforcement "Circle"
    TestSetSystemSetting "Plot.Enabled", "Yes"
    TestSetSystemSetting "General.ExecutionReportEnabled", "No"
    Dim message As String, solves As Long
    solves = SectionEquilibriumSolveCount()
    message = FormSectionForWorkbook(ThisWorkbook)
    TestAssertTrue stats, "workflow.form.noSolver", SectionEquilibriumSolveCount() = solves
    TestAssertTrue stats, "workflow.form.autoPreview", TestPlotVisibleTitleContains("Сформированное сечение")
    TestAssertTrue stats, "workflow.form.elapsedMessage", InStr(message, "Время формирования сечения:") > 0
    Dim properties As Variant, propertyRow As Long, storedWarning As String
    properties = TestResultTable("rngNDMSectionProperties")
    storedWarning = "Предупреждение: контрольное сообщение формирования."
    For propertyRow = 2 To UBound(properties, 1)
        If CStr(properties(propertyRow, 3)) = "Geometry.RecipeDiagnostics" Then properties(propertyRow, 4) = storedWarning
    Next propertyRow
    ThisWorkbook.Names("rngNDMSectionProperties").RefersToRange.Resize(UBound(properties, 1), UBound(properties, 2)).Value2 = properties
    Dim geometry As Variant, contours As Variant
    geometry = TestResultTable("rngNDMSectionGeometry"): contours = TestResultTable("rngNDMSectionContours")
    UpdateSectionPlotForWorkbook ThisWorkbook
    TestAssertTrue stats, "workflow.preview.noSolver", SectionEquilibriumSolveCount() = solves
    TestSetSystemSetting "Geometry.Source", "AutoCAD"
    TestSetSystemSetting "Geometry.Type", "Invalid.Unused.Geometry"
    TestSetSystemSetting "Circle.Diameter", "INVALID"
    Dim calculationStarted As Double, actualElapsed As Double, savedElapsed As Double
    calculationStarted = Timer
    message = RunSectionCalculationForWorkbook(ThisWorkbook, False)
    actualElapsed = Timer - calculationStarted: If actualElapsed < 0# Then actualElapsed = actualElapsed + 86400#
    savedElapsed = CDbl(ThisWorkbook.Names("rngBatchSummary").RefersToRange.Offset(3, 0).Value2)
    TestAssertClose stats, "workflow.calculate.elapsed", savedElapsed, actualElapsed, 0.1
    TestAssertTrue stats, "workflow.calculate.elapsedMessage", InStr(message, "Время расчёта:") > 0
    TestAssertTrue stats, "workflow.calculate.noFormationMessages", InStr(message, storedWarning) = 0
    TestAssertTrue stats, "workflow.calculate.savedFormationDiagnostic", CStr(TestResultsPropertyValue("ALL", "Geometry.RecipeDiagnostics")) = storedWarning
    TestAssertTrue stats, "workflow.calculate.frozenSource", InStr(message, "Расчет завершен") > 0 And ResultsGeometrySource(ThisWorkbook) = "Generated"
    TestAssertTrue stats, "workflow.calculate.geometryExact", TestWorkflowSameTable(geometry, TestResultTable("rngNDMSectionGeometry"))
    TestAssertTrue stats, "workflow.calculate.contoursExact", TestWorkflowSameTable(contours, TestResultTable("rngNDMSectionContours"))
    Dim model As CSectionModel: Set model = ReadFormedSectionFromResults(ThisWorkbook)
    TestAssertTrue stats, "workflow.calculate.eightBars", model.RebarCount = 8
    Dim n As Double: n = CDbl(TestResultsPropertyValue("LC1", "N"))
    TestSetSystemSetting "Units.Length.Output", "m"
    TestSetSystemSetting "Units.Area.Output", "m2"
    message = RunSectionCalculationForWorkbook(ThisWorkbook, False)
    TestAssertTrue stats, "workflow.repeat.completed", InStr(message, "Расчет завершен") > 0
    TestAssertTrue stats, "workflow.repeat.geometryExact", TestWorkflowSameTable(geometry, TestResultTable("rngNDMSectionGeometry"))
    TestAssertTrue stats, "workflow.repeat.contoursExact", TestWorkflowSameTable(contours, TestResultTable("rngNDMSectionContours"))
    TestAssertClose stats, "workflow.repeat.force", CDbl(TestResultsPropertyValue("LC1", "N")), n, 0#
    TestSetSystemSetting "Plot.LegendMode", "INVALID"
    message = RunSectionCalculationForWorkbook(ThisWorkbook, False)
    TestAssertTrue stats, "workflow.plotError.warning", InStr(message, "Расчет завершен") > 0 And InStr(message, "Схема не обновлена") > 0
    TestAssertClose stats, "workflow.plotError.force", CDbl(TestResultsPropertyValue("LC1", "N")), n, 0#
    TestSetSystemSetting "Plot.LegendMode", "Separate"
    TestSetProfileSetting "PR2", "Visualization.State", "INVALID"
    message = RunSectionCalculationForWorkbook(ThisWorkbook, False)
    TestAssertTrue stats, "workflow.visualizationError.warning", InStr(message, "Расчет завершен") > 0 And InStr(message, "Схема не обновлена") > 0
    TestAssertClose stats, "workflow.visualizationError.force", CDbl(TestResultsPropertyValue("LC1", "N")), n, 0#
    TestSetProfileSetting "PR2", "Visualization.State", "CrackedState"
    TestSetSystemSetting "Geometry.Source", "Generated"
    Dim code As Long, reason As String
    On Error Resume Next
    message = FormSectionForWorkbook(ThisWorkbook)
    code = Err.Number: reason = Err.Description
    On Error GoTo 0
    TestAssertTrue stats, "workflow.failedForm.reason", code <> 0 And InStr(reason, "Возможное решение:") > 0
    TestAssertTrue stats, "workflow.failedForm.geometryPreserved", TestWorkflowSameTable(geometry, TestResultTable("rngNDMSectionGeometry"))
    TestAssertTrue stats, "workflow.failedForm.contoursPreserved", TestWorkflowSameTable(contours, TestResultTable("rngNDMSectionContours"))
    TestSetSystemSetting "Geometry.Type", "Circle"
    TestSetSystemSetting "Circle.Diameter", "400"
    Dim loads As Object: Set loads = ThisWorkbook.Names("rngLoadCombinations").RefersToRange
    loads.Cells(2, 2).Value2 = "INVALID_UNUSED_LOAD"
    solves = SectionEquilibriumSolveCount()
    message = FormSectionForWorkbook(ThisWorkbook)
    TestAssertTrue stats, "workflow.reform.independentLoads", InStr(message, "Сечение сформировано") > 0 And SectionEquilibriumSolveCount() = solves
    TestAssertTrue stats, "workflow.reform.changed", Not TestWorkflowSameTable(geometry, TestResultTable("rngNDMSectionGeometry"))
    Dim results As Variant: results = TestResultTable("rngNDMElementResults")
    TestAssertTrue stats, "workflow.reform.clearsStates", UBound(results, 1) = 1
    UpdateSectionPlotForWorkbook ThisWorkbook
    TestAssertTrue stats, "workflow.reform.previewOnly", SectionEquilibriumSolveCount() = solves
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Доказывает независимость численных состояний от JSON и исходных подписей.
' Удаление необязательного столбца не должно прерывать запись готового НДС.
Private Sub TestWorkflowOptionalAnnotations(ByRef stats As TUiTestStats)
    TestPrepareCircleInput
    TestSetSystemSetting "Circle.Diameter", "300"
    TestFixtureResetReinforcement "Circle"
    TestSetSystemSetting "Plot.Enabled", "Yes"
    TestSetSystemSetting "General.ExecutionReportEnabled", "No"
    Dim message As String, states As Variant, annotations As Variant, geometry As Object, target As Object
    message = FormSectionForWorkbook(ThisWorkbook)
    message = RunSectionCalculationForWorkbook(ThisWorkbook, False)
    states = TestResultTable("rngNDMElementResults")
    Set geometry = ThisWorkbook.Names("rngNDMSectionGeometry").RefersToRange
    Dim data As Variant: data = TestResultTable("rngNDMSectionGeometry")
    geometry.Offset(1, 15).Resize(UBound(data, 1) - 1, 1).Value2 = CVErr(xlErrValue)
    geometry.Offset(1, 14).Resize(UBound(data, 1) - 1, 1).Value2 = CVErr(xlErrValue)
    geometry.Offset(0, 15).ClearContents
    Set target = ThisWorkbook.Names("rngNDMSectionAnnotations").RefersToRange
    target.Offset(0, 23).ClearContents
    message = RunSectionCalculationForWorkbook(ThisWorkbook, False)
    TestAssertTrue stats, "workflow.optional.completed", InStr(message, "Расчет завершен") > 0
    TestAssertTrue stats, "workflow.optional.statesExact", TestWorkflowSameTable(states, TestResultTable("rngNDMElementResults"))
    TestAssertTrue stats, "workflow.optional.plotAvailable", InStr(message, "Схема не обновлена") = 0
    annotations = TestResultTable("rngNDMSectionAnnotations")
    target.Resize(UBound(annotations, 1), UBound(annotations, 2)).ClearContents
    message = RunSectionCalculationForWorkbook(ThisWorkbook, False)
    TestAssertTrue stats, "workflow.noAnnotations.completed", InStr(message, "Расчет завершен") > 0
    TestAssertTrue stats, "workflow.noAnnotations.statesExact", TestWorkflowSameTable(states, TestResultTable("rngNDMElementResults"))
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' ДЛЯ ТЕСТОВ
' Проверяет расчёт Circle/Rectangles и неизменность снимка между командами.
Public Function RunSectionWorkflowCalculationTests() As String
    Dim stats As TUiTestStats, saved As Variant
    saved = TestFixtureSaveInputs(ThisWorkbook.Worksheets("Config").UsedRange)
    On Error GoTo Failed
    TestCircleWorkbookRunWritesResults stats
    TestRectanglesWorkbookRunWritesResults stats
    TestWorkflowFrozenGeometry stats
    TestWorkflowOptionalAnnotations stats
    TestNoActiveLoadsPreserveResults stats
Done:
    TestFixtureRestoreInputs ThisWorkbook.Worksheets("Config").UsedRange, saved
    TestAppendLine stats, "TOTAL_WORKFLOW_CALCULATION: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunSectionWorkflowCalculationTests = stats.Report
    Exit Function
Failed:
    TestAssertTrue stats, "unexpected.workflowCalculation." & Err.Number & "." & Err.Description, False
    Resume Done
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' ДЛЯ ТЕСТОВ
' Проверяет автоматическую схему после формирования и независимость от НДС.
Public Function RunFormationPreviewTests() As String
    Dim stats As TUiTestStats, saved As Variant, message As String, solves As Long
    saved = TestFixtureSaveInputs(ThisWorkbook.Worksheets("Config").UsedRange)
    On Error GoTo Failed
    TestPrepareCircleInput
    TestSetSystemSetting "General.ExecutionReportEnabled", "No"
    TestSetSystemSetting "Plot.Enabled", "Yes"
    TestSetSystemSetting "Plot.AutoUpdateAfterCalculation", "No"
    solves = SectionEquilibriumSolveCount()
    message = FormSectionForWorkbook(ThisWorkbook)
    TestAssertTrue stats, "formPreview.saved", InStr(message, "Сечение сформировано") > 0
    TestAssertTrue stats, "formPreview.drawn", TestPlotVisibleTitleContains("Сформированное сечение")
    TestAssertTrue stats, "formPreview.annotations", TestCountPlotShapes("AnnotationLine") > 0
    TestAssertTrue stats, "formPreview.noSolver", SectionEquilibriumSolveCount() = solves
    TestSetSystemSetting "Plot.LegendMode", "INVALID"
    message = FormSectionForWorkbook(ThisWorkbook)
    TestAssertTrue stats, "formPreview.error.warning", InStr(message, "Сечение сформировано") > 0 And InStr(message, "Схема не обновлена") > 0
    TestAssertTrue stats, "formPreview.error.reason", InStr(message, "Plot.LegendMode") > 0
    Dim section As CSectionModel: Set section = ReadFormedSectionFromResults(ThisWorkbook)
    TestAssertTrue stats, "formPreview.error.snapshot", section.ConcreteCount > 0 And section.RebarCount = 8
    TestAssertTrue stats, "formPreview.error.noSolver", SectionEquilibriumSolveCount() = solves
    TestSetSystemSetting "Plot.Enabled", "No"
    message = FormSectionForWorkbook(ThisWorkbook)
    TestAssertTrue stats, "formPreview.disabled.saved", InStr(message, "Сечение сформировано") > 0 And InStr(message, "Схема не обновлена") = 0
    TestAssertTrue stats, "formPreview.disabled.noDrawing", Not TestPlotVisibleTitleContains("Сформированное сечение")
    TestAssertTrue stats, "formPreview.disabled.noSolver", SectionEquilibriumSolveCount() = solves
Done:
    TestFixtureRestoreInputs ThisWorkbook.Worksheets("Config").UsedRange, saved
    TestAppendLine stats, "TOTAL_FORMATION_PREVIEW: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunFormationPreviewTests = stats.Report
    Exit Function
Failed:
    TestAssertTrue stats, "unexpected.formPreview." & Err.Number & "." & Err.Description, False
    Resume Done
End Function


' -------------------- ДЛЯ ТЕСТОВ --------------------
' групповая палитра и сохранение высоты строк Results.
Public Function RunResultBlockOutputTests() As String
    Dim stats As TUiTestStats, fixture As Object, sheet As Object, target As Object
    Dim values(1 To 120, 1 To 11) As Variant, rows(1 To 120) As Long
    Dim statuses As Variant, row As Long, column As Variant, expected As Long, message As String
    On Error GoTo Failed
    Set fixture = Application.Workbooks.Add(-4167): Set sheet = fixture.Worksheets(1)
    Set target = sheet.Range("AZ10:BJ129")
    target.Interior.Color = RGB(210, 220, 230): target.Font.Bold = True
    target.RowHeight = 18.25: sheet.Range("AY10").Value2 = "KEEP"
    statuses = Array("OK", "Fail", "BaseFail", "NumFail", "InputErr", "CalcErr", "N/A", "", "Cracked", "NotCracked")
    For row = 1 To 120
        rows(row) = row
        For Each column In Array(2, 4, 11)
            values(row, CLng(column)) = statuses((row + CLng(column)) Mod 10)
        Next column
        values(row, 1) = row
    Next row
    target.Value2 = values
    ApplyTableStatusFills target, rows, Array(2, 4, 11), True, 4
    For row = 1 To 120
        For Each column In Array(2, 4, 11)
            expected = RGB(210, 220, 230)
            If Len(CStr(values(row, CLng(column)))) > 0 Then
                If CLng(column) <> 4 Or (values(row, 4) <> "Cracked" And values(row, 4) <> "NotCracked") Then
                    expected = StatusFillColor(CStr(values(row, CLng(column))))
                End If
            End If
            TestAssertTrue stats, "block.color." & row & "." & column, target.Cells(row, CLng(column)).Interior.Color = expected
        Next column
    Next row
    TestAssertTrue stats, "block.unrelatedColumn", target.Columns(1).Interior.Color = RGB(210, 220, 230)
    TestAssertTrue stats, "block.fontUnchanged", target.Font.Bold = True
    TestAssertTrue stats, "block.heightUnchanged", target.RowHeight = 18.25
    TestAssertTrue stats, "block.outsideUnchanged", sheet.Range("AY10").Value2 = "KEEP"
    fixture.Close False: Set fixture = Nothing
    TestPrepareCircleInput
    TestSetSystemSetting "General.ExecutionReportEnabled", "No"
    ThisWorkbook.Worksheets("Results").Rows.RowHeight = 17.25
    message = FormSectionForWorkbook(ThisWorkbook)
    TestAssertTrue stats, "block.formed", InStr(message, "Сечение сформировано") > 0
    TestAssertTrue stats, "block.formationKeepsHeight", ThisWorkbook.Worksheets("Results").Rows.RowHeight = 17.25
    TestAssertTrue stats, "block.geometryJsonRightAligned", ThisWorkbook.Names.Item("rngNDMSectionGeometry").RefersToRange.Offset(1, 15).HorizontalAlignment = -4152
    TestAssertTrue stats, "block.annotationJsonRightAligned", ThisWorkbook.Names.Item("rngNDMSectionAnnotations").RefersToRange.Offset(1, 23).HorizontalAlignment = -4152
    message = RunSectionCalculationForWorkbook(ThisWorkbook)
    TestAssertTrue stats, "block.calculated", InStr(message, "Расчет завершен") > 0
    TestAssertTrue stats, "block.calculationKeepsHeight", ThisWorkbook.Worksheets("Results").Rows.RowHeight = 17.25
    GoTo Finished
Failed:
    TestAssertTrue stats, "block.runtime." & Err.Number & "." & Err.Description, False
Finished:
    On Error Resume Next
    If Not fixture Is Nothing Then fixture.Close False
    ThisWorkbook.Worksheets("Results").Rows.RowHeight = 15
    On Error GoTo 0
    RunResultBlockOutputTests = stats.Report & "TOTAL_BLOCK_OUTPUT: passed=" & stats.Passed & "; failed=" & stats.Failed
End Function


' -------------------- ДЛЯ ТЕСТОВ --------------------
' ДЛЯ ТЕСТОВ
' Проверяет быстрый вывод, изменение единиц и статусов, восстановление шапок
' и очистку старых сочетаний без изменения численных результатов.
Public Function RunResultsFormattingTests() As String
    Dim stats As TUiTestStats
    TestResultsFormatting stats
    TestAppendLine stats, "TOTAL_RESULTS_FORMATTING: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunResultsFormattingTests = stats.Report
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Меняет единицы, число сочетаний и заголовки только в отдельной тестовой книге.
Private Sub TestResultsFormatting(ByRef stats As TUiTestStats)
    Dim saved As Variant, message As String
    Dim states As Variant, properties As Variant, target As Object, section As CSectionModel
    Dim solves As Long, number As Long, row As Long, values As Variant
    saved = TestFixtureSaveInputs(ThisWorkbook.Worksheets("Config").UsedRange)
    On Error GoTo Failed
    Dim settings As CSystemSettingsReader: Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook
    Dim extension As String: extension = settings.GetRequiredString("General.DiagramExtension")
    TestPrepareCircleInput
    TestSetSystemSetting "Mesh.StepMode", "User"
    TestSetSystemSetting "Mesh.StepX", "30": TestSetSystemSetting "Mesh.StepY", "30"
    TestSetSystemSetting "Mesh.BoundarySubdivisions", "2"
    TestSetSystemSetting "Plot.Enabled", "No"
    TestSetSystemSetting "General.ExecutionReportEnabled", "Yes"
    TestSetSystemSetting "General.ReformatResultsEnabled", "Yes"
    TestConfigurePerformanceFixture "MIXED", 6, "Newton", extension:=extension
    TestAssertTrue stats, "formatting.dropdown", TestSystemSettingValidationHasOptions("General.ReformatResultsEnabled", Array("Yes", "No"))
    message = FormSectionForWorkbook(ThisWorkbook)
    message = RunSectionCalculationForWorkbook(ThisWorkbook, False)
    states = TestResultTable("rngNDMElementResults")
    properties = TestResultTable("rngNDMSectionProperties")
    Dim name As Variant
    For Each name In Array("rngStrengthSummaryAnchor", "rngCrackSummaryAnchor", "rngStabilitySummaryAnchor")
        ThisWorkbook.Names.Item(CStr(name)).RefersToRange.Font.Size = 11.25
    Next name
    ThisWorkbook.Names.Item("rngBatchSummary").RefersToRange.Offset(12, 1).Font.Size = 11.25
    TestSetSystemSetting "General.ReformatResultsEnabled", "No"
    message = RunSectionCalculationForWorkbook(ThisWorkbook, False)
    TestAssertTrue stats, "formatting.completed", InStr(message, "Расчет завершен") > 0
    For Each name In Array("rngStrengthSummaryAnchor", "rngCrackSummaryAnchor", "rngStabilitySummaryAnchor")
        TestAssertTrue stats, "formatting.reusesBody." & CStr(name), CDbl(ThisWorkbook.Names.Item(CStr(name)).RefersToRange.Font.Size) = 11.25
    Next name
    TestAssertTrue stats, "formatting.reusesSummary", CDbl(ThisWorkbook.Names.Item("rngBatchSummary").RefersToRange.Offset(12, 1).Font.Size) = 11.25
    TestAssertTrue stats, "formatting.elementStatesExact", TestWorkflowSameTable(states, TestResultTable("rngNDMElementResults"))
    TestAssertTrue stats, "formatting.propertiesExact", TestWorkflowSameTable(properties, TestResultTable("rngNDMSectionProperties"))
    Set target = ThisWorkbook.Names.Item("rngStrengthSummaryAnchor").RefersToRange
    TestAssertTrue stats, "formatting.strengthTitle", Len(CStr(target.Offset(-5, 0).Value2)) > 0
    TestAssertTrue stats, "formatting.strengthMergedHeader", target.Offset(-4, 0).MergeCells
    TestAssertTrue stats, "formatting.strengthDataText", CStr(target.NumberFormat) = "@"
    TestSetSystemSetting "Units.Force.Output", "kN"
    TestSetSystemSetting "Units.Length.Output", "mm"
    TestSetSystemSetting "Units.Stress.Output", "kPa"
    message = RunSectionCalculationForWorkbook(ThisWorkbook, False)
    TestAssertTrue stats, "formatting.strengthForceUnit", InStr(CStr(target.Offset(-1, 3).Value2), "kN") > 0
    Set target = ThisWorkbook.Names.Item("rngStabilitySummaryAnchor").RefersToRange
    TestAssertTrue stats, "formatting.stabilityForceUnit", InStr(CStr(target.Offset(-1, 5).Value2), "kN") > 0
    Set target = ThisWorkbook.Names.Item("rngNDMElementResults").RefersToRange
    TestAssertTrue stats, "formatting.elementStressUnit", CStr(target.Offset(0, 7).Value2) = "Stress, kPa"
    TestSetProfileSetting "PR1", "Calculation.Strength.Capacity", "No"
    message = RunSectionCalculationForWorkbook(ThisWorkbook, False)
    Set target = ThisWorkbook.Names.Item("rngBatchSummary").RefersToRange.Offset(12, 6)
    TestAssertTrue stats, "formatting.changedStatus", CStr(target.Value2) = "N/A"
    TestAssertTrue stats, "formatting.changedStatusColor", CLng(target.Interior.Color) = StatusFillColor("N/A")
    TestConfigurePerformanceFixture "MIXED", 1, "Newton", extension:=extension
    message = RunSectionCalculationForWorkbook(ThisWorkbook, False)
    Set target = ThisWorkbook.Names.Item("rngBatchSummary").RefersToRange
    TestAssertTrue stats, "formatting.oldCombinationsCleared", Len(CStr(target.Offset(13, 0).Value2)) = 0
    TestAssertTrue stats, "formatting.oldStatusColorCleared", CLng(target.Offset(13, 3).Interior.ColorIndex) = -4142
    Set target = ThisWorkbook.Names.Item("rngStrengthSummaryAnchor").RefersToRange
    Dim strengthWriter As CStrengthSummaryWriter: Set strengthWriter = New CStrengthSummaryWriter
    Dim lastDataRow As Long: lastDataRow = strengthWriter.RequiredDataRowsForWorkbook(ThisWorkbook, 1)
    target.Offset(lastDataRow - 1, 0).NumberFormat = "General"
    message = RunSectionCalculationForWorkbook(ThisWorkbook, False)
    TestAssertTrue stats, "formatting.newRowsFormatted", CStr(target.Offset(lastDataRow - 1, 0).NumberFormat) = "@"
    target.Offset(-4, 0).Resize(4, strengthWriter.RequiredColumns).UnMerge
    target.Offset(-4, 0).Resize(4, strengthWriter.RequiredColumns).ClearContents
    message = RunSectionCalculationForWorkbook(ThisWorkbook, False)
    TestAssertTrue stats, "formatting.missingHeaderRestored", CStr(target.Offset(-4, 0).Value2) = "Combination ID"
    TestAssertTrue stats, "formatting.headerMergesRestored", target.Offset(-4, 0).MergeCells
    states = TestResultTable("rngNDMElementResults")
    solves = SectionEquilibriumSolveCount()
    TestSetSystemSetting "General.ReformatResultsEnabled", "INVALID"
    On Error Resume Next
    message = RunSectionCalculationForWorkbook(ThisWorkbook, False)
    number = Err.Number: Err.Clear
    On Error GoTo Failed
    TestAssertTrue stats, "formatting.invalidRejected", number <> 0
    TestAssertTrue stats, "formatting.invalidNoSolve", SectionEquilibriumSolveCount() = solves
    TestAssertTrue stats, "formatting.invalidPreservesResults", TestWorkflowSameTable(states, TestResultTable("rngNDMElementResults"))
    TestSetSystemSetting "General.ReformatResultsEnabled", "No"
    message = FormSectionForWorkbook(ThisWorkbook)
    Set section = ReadFormedSectionFromResults(ThisWorkbook)
    TestAssertTrue stats, "formatting.formationModel", section.ConcreteCount > 0 And section.RebarCount = 8
    TestAssertTrue stats, "formatting.formationClearsStates", UBound(TestResultTable("rngNDMElementResults"), 1) = 1
    values = TestResultTable("rngNDMSectionProperties")
    Set target = ThisWorkbook.Names.Item("rngNDMSectionProperties").RefersToRange
    For row = 2 To UBound(values, 1)
        TestAssertTrue stats, "formatting.propertyNoWrap." & row, Not CBool(target.Offset(row - 1, 3).WrapText)
    Next row
Done:
    TestFixtureRestoreInputs ThisWorkbook.Worksheets("Config").UsedRange, saved
    Exit Sub
Failed:
    TestAssertTrue stats, "unexpected.resultsFormatting." & Err.Number & "." & Err.Description, False
    Resume Done
End Sub


' -------------------- ДЛЯ ТЕСТОВ --------------------
' ДЛЯ ТЕСТОВ
' Повторяет проверки импортированных подписей и независимости схемы от будущего ввода.
Public Function RunPreviewSnapshotContractTests() As String
    Dim stats As TUiTestStats
    On Error GoTo Failed
    TestAutoCADPreviewWritesAndDrawsBoundsDimensions stats
    TestPlotUsesSavedGeometryAfterSourceChange stats
    GoTo Finished
Failed:
    TestAssertTrue stats, "previewSnapshot.runtime." & CStr(Err.Number) & "." & Err.Description, False
Finished:
    RunPreviewSnapshotContractTests = stats.Report & "TOTAL_PREVIEW_SNAPSHOT_CONTRACT: passed=" & stats.Passed & "; failed=" & stats.Failed
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' ДЛЯ ТЕСТОВ
' Проверяет итог расчёта сохранённой импортированной сетки без повторения сведений об импорте.
Public Function RunImportMessageScopeTests() As String
    Dim stats As TUiTestStats, saved As Variant
    saved = TestFixtureSaveInputs(ThisWorkbook.Worksheets("Config").UsedRange)
    On Error GoTo Failed
    TestAutoCADCalculationMessageUsesSavedGeometry stats
    GoTo Finished
Failed:
    TestAssertTrue stats, "importMessageScope.runtime." & CStr(Err.Number) & "." & Err.Description, False
Finished:
    TestFixtureRestoreInputs ThisWorkbook.Worksheets("Config").UsedRange, saved
    RunImportMessageScopeTests = stats.Report & "TOTAL_IMPORT_MESSAGE_SCOPE: passed=" & stats.Passed & "; failed=" & stats.Failed
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Измеряет только чтение готовых 100/1000/10000 аннотаций, без формирования,
' НДС и рисунка. Запускается одинаково до и после изменения закрытого резерва;
' проверка всех полей выполняется вне измеряемого участка.
Public Function RunTestAnnotationReadBenchmark() As String
    Dim stats As TUiTestStats, fixture As Object, sheet As Object, anchor As Object
    Dim reader As CSectionPlotDataReader, settings As CSystemSettingsReader
    Dim configuration(1 To 2, 1 To 3) As Variant, properties(1 To 2, 1 To 4) As Variant
    Dim sizes As Variant, size As Variant, data As Variant, row As Long, column As Long, repetition As Long
    Dim started As Double, elapsed As Double, errorNumber As Long, description As String
    Dim wasAlerts As Boolean: wasAlerts = Application.DisplayAlerts
    On Error GoTo Failed
    Set fixture = Application.Workbooks.Add(-4167)
    Set sheet = fixture.Worksheets(1): sheet.Name = "AnnotationBench"
    TestInitializeEmptyContourFixture fixture, sheet
    fixture.Names.Add Name:="rngNDMSectionGeometry", RefersTo:="='AnnotationBench'!$A$5"
    sheet.Range("A5").Resize(3, 15).Value2 = TestAudit03GeometrySnapshotArray("mm", 1#)
    properties(1, 1) = "LoadCase": properties(1, 2) = "Parameter": properties(1, 3) = "Value": properties(1, 4) = "Unit"
    properties(2, 1) = "ALL": properties(2, 2) = "Bounds.MaxX": properties(2, 3) = 100#: properties(2, 4) = "mm"
    sheet.Range("R5").Resize(2, 4).Value2 = properties
    fixture.Names.Add Name:="rngNDMSectionProperties", RefersTo:="='AnnotationBench'!$R$5"
    configuration(1, 1) = "Параметр": configuration(1, 2) = "Значение": configuration(1, 3) = "Комментарий"
    configuration(2, 1) = "Plot.LoadCase": configuration(2, 2) = "Worst"
    sheet.Range("AO5").Resize(2, 3).Value2 = configuration
    Set settings = New CSystemSettingsReader: settings.LoadFromRange sheet.Range("AO5").Resize(2, 3)
    Set anchor = sheet.Range("A30")
    fixture.Names.Add Name:="rngNDMSectionAnnotations", RefersTo:="='AnnotationBench'!$A$30"
    Set reader = New CSectionPlotDataReader
    sizes = Array(100, 1000, 10000, 100) ' Последний меньший набор проверяет очистку хвоста после большого.
    For Each size In sizes
        anchor.Resize(10002, 24).ClearContents
        data = TestAnnotationBenchmarkArray(CLng(size))
        anchor.Resize(CLng(size) + 1, 24).Value2 = data
        For repetition = 1 To 3
            started = Timer
            reader.LoadGeometryPreviewFromWorkbook fixture, settings
            elapsed = Timer - started: If elapsed < 0# Then elapsed = elapsed + 86400#
            TestAppendLine stats, "ANNOTATION_BENCH: count=" & size & "; repetition=" & repetition & "; seconds=" & TestFormatNumberInvariant(elapsed)
        Next repetition
        TestAssertTrue stats, "annotationBench.count." & size, reader.AnnotationCount = CLng(size)
        Dim valuesMatch As Boolean: valuesMatch = True
        For row = 1 To CLng(size)
            If reader.AnnotationType(row) <> "TEXT_ANNOTATION" Or reader.AnnotationText(row) <> "ROW_" & row Then valuesMatch = False
            If reader.AnnotationStartX(row) <> row * 0.25 Or reader.AnnotationStartY(row) <> -row * 0.5 Then valuesMatch = False
            If reader.AnnotationEndX(row) <> row + 1# Or reader.AnnotationEndY(row) <> row + 2# Then valuesMatch = False
            If reader.AnnotationNormalX(row) <> 1# Or reader.AnnotationNormalY(row) <> 0# Or reader.AnnotationSweepAngle(row) <> 0# Then valuesMatch = False
        Next row
        TestAssertTrue stats, "annotationBench.orderAndEveryField." & size, valuesMatch
    Next size
    anchor.Cells(51, 4).Value2 = CVErr(2015) ' Ошибка в середине набора должна очистить частично прочитанные массивы.
    On Error Resume Next
    reader.LoadGeometryPreviewFromWorkbook fixture, settings
    errorNumber = Err.Number: description = Err.Description: Err.Clear
    On Error GoTo Failed
    TestAssertTrue stats, "annotationBench.invalidNumber.rejected", errorNumber <> 0
    TestAssertTrue stats, "annotationBench.invalidNumber.address", InStr(1, description, "D80", vbTextCompare) > 0
    TestAssertTrue stats, "annotationBench.failureClearsState", reader.Count = 0 And reader.AnnotationCount = 0
    anchor.Resize(101, 24).Value2 = data
    reader.LoadGeometryPreviewFromWorkbook fixture, settings
    TestAssertTrue stats, "annotationBench.recovery", reader.Count = 2 And reader.AnnotationCount = 100 And reader.AnnotationText(100) = "ROW_100"
    GoTo Finished
Failed:
    TestAssertTrue stats, "annotationBench.runtime." & CStr(Err.Number) & "." & Err.Description, False
Finished:
    On Error Resume Next
    Application.DisplayAlerts = False
    If Not fixture Is Nothing Then fixture.Close False
    Application.DisplayAlerts = wasAlerts
    On Error GoTo 0
    RunTestAnnotationReadBenchmark = stats.Report & "TOTAL_ANNOTATION_BENCH: passed=" & stats.Passed & "; failed=" & stats.Failed
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Готовит независимые координаты и подписи без обращения к закрытым массивам
' reader-а. Все дроби двоичные, поэтому проверка равенства не требует допуска.
Private Function TestAnnotationBenchmarkArray(ByVal count As Long) As Variant
    Dim data() As Variant, headers As Variant, row As Long, column As Long
    ReDim data(1 To count + 1, 1 To 24)
    headers = Array("RunID", "AnnotationType", "AnnotationID", "StartX", "StartY", "EndX", "EndY", "OutsideNormalX", "OutsideNormalY", "Text", "Value", "Unit", "Comment", "CombinationID", "Standard", "AnchorID", "LoopID", "LoopRole", "SectionRevision", "RegionArea", "AreaUnit", "GeometrySource", "CoordinateUnit", "ExtraData")
    For column = 1 To 24: data(1, column) = headers(column - 1): Next column
    For row = 1 To count
        data(row + 1, 1) = 46023#: data(row + 1, 2) = "TEXT_ANNOTATION": data(row + 1, 3) = "T" & row
        data(row + 1, 4) = row * 0.25: data(row + 1, 5) = -row * 0.5
        data(row + 1, 6) = row + 1#: data(row + 1, 7) = row + 2#
        data(row + 1, 8) = 1#: data(row + 1, 9) = 0#: data(row + 1, 10) = "ROW_" & row
        data(row + 1, 23) = "mm"
    Next row
    TestAnnotationBenchmarkArray = data
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает подготовку графического layout до/после резерва без Excel Chart.
' Проверяет каждый элемент, пустой и повторный Initialize, меньший набор
' после большого и сохранность предыдущего состояния при отказе до инициализации.
Public Function RunTestAnnotationLayoutBenchmark() As String
    Dim stats As TUiTestStats, layout As CPlotAnnotationLayout, size As Variant, row As Long, repetition As Long
    Dim started As Double, elapsed As Double, valuesMatch As Boolean, number As Long
    Set layout = New CPlotAnnotationLayout
    On Error GoTo Failed
    For Each size In Array(100, 1000, 10000, 100)
        For repetition = 1 To 3
            layout.Initialize 0#, 20000#, 0#, 20000#, 0#, 0#, 640#, 430#
            TestAssertTrue stats, "layoutBench.initializeZero." & size, layout.Count = 0
            started = Timer
            For row = 1 To CLng(size)
                layout.AddPointLabel row, row, "ROW_" & row
            Next row
            elapsed = Timer - started: If elapsed < 0# Then elapsed = elapsed + 86400#
            TestAppendLine stats, "LAYOUT_BENCH: count=" & size & "; repetition=" & repetition & "; seconds=" & TestFormatNumberInvariant(elapsed)
        Next repetition
        TestAssertTrue stats, "layoutBench.count." & size, layout.Count = CLng(size)
        valuesMatch = True
        For row = 1 To CLng(size)
            If layout.ItemKind(row) <> "TEXT" Or layout.Text(row) <> "ROW_" & row Then valuesMatch = False
            If Abs(layout.X1(row) - (row * 640# / 20000# - 45#)) > 0.000000001 Then valuesMatch = False
            If Abs(layout.Y1(row) - (430# - row * 430# / 20000# - 4.5)) > 0.000000001 Then valuesMatch = False
            If layout.X2(row) <> 90# Or layout.Y2(row) <> 9# Or layout.FontSize(row) <> 5# Then valuesMatch = False
            If layout.Rotation(row) <> 0# Or layout.Weight(row) <> 0# Or Not layout.Bold(row) Or layout.Arrowheads(row) Then valuesMatch = False
            If layout.Color(row) <> RGB(20, 30, 90) Then valuesMatch = False
        Next row
        TestAssertTrue stats, "layoutBench.orderAndEveryField." & size, valuesMatch
    Next size
    On Error Resume Next
    layout.Initialize 0#, 1#, 0#, 1#, 0#, 0#, 640#, 430#, "mm", 0.5
    number = Err.Number: Err.Clear
    On Error GoTo Failed
    TestAssertTrue stats, "layoutBench.invalidInitialize.rejected", number <> 0
    TestAssertTrue stats, "layoutBench.invalidInitialize.preservesOldState", layout.Count = 100 And layout.Text(100) = "ROW_100"
    layout.Initialize 0#, 20000#, 0#, 20000#, 0#, 0#, 640#, 430#
    TestAssertTrue stats, "layoutBench.emptyAfterReuse", layout.Count = 0
    layout.AddPointLabel 1#, 1#, "RECOVERED"
    TestAssertTrue stats, "layoutBench.recovered", layout.Count = 1 And layout.Text(1) = "RECOVERED"
    GoTo Finished
Failed:
    TestAssertTrue stats, "layoutBench.runtime." & CStr(Err.Number) & "." & Err.Description, False
Finished:
    RunTestAnnotationLayoutBenchmark = stats.Report & "TOTAL_LAYOUT_BENCH: passed=" & stats.Passed & "; failed=" & stats.Failed
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Запускает два микрозамера в одинаковом ручном режиме без событий и обновления
' экрана. Время подготовки/записи fixture и восстановления Excel не входит
' во внутренние интервалы; метод не измеряет полный пользовательский сценарий.
Public Function RunTestAnnotationControlledBenchmark() As String
    Dim oldCalculation As Long, oldEvents As Boolean, oldScreenUpdating As Boolean, oldAlerts As Boolean
    Dim report As String, number As Long, description As String, oldBooks As Long
    oldCalculation = Application.Calculation: oldEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating: oldAlerts = Application.DisplayAlerts
    oldBooks = Application.Workbooks.Count
    On Error GoTo Failed
    Application.EnableEvents = False: Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    report = "BENCH_CONTROL: calculation=" & Application.Calculation & "; events=" & Application.EnableEvents & _
        "; screenUpdating=" & Application.ScreenUpdating & "; openBooks=" & oldBooks & _
        "; calculationState=" & Application.CalculationState & vbCrLf
    report = report & RunTestAnnotationReadBenchmark() & vbCrLf & RunTestAnnotationLayoutBenchmark() & vbCrLf
    GoTo Finished
Failed:
    number = Err.Number: description = Err.Description
    report = report & "BENCH_CONTROL_ERROR=" & number & "; " & description & vbCrLf
Finished:
    On Error Resume Next
    Application.Calculation = oldCalculation: Application.DisplayAlerts = oldAlerts
    Application.EnableEvents = oldEvents: Application.ScreenUpdating = oldScreenUpdating
    On Error GoTo 0
    report = report & "BENCH_CONTROL_RESTORE: calculation=" & Application.Calculation & "; events=" & Application.EnableEvents & _
        "; screenUpdating=" & Application.ScreenUpdating & "; openBooks=" & Application.Workbooks.Count & vbCrLf
    RunTestAnnotationControlledBenchmark = report
End Function


' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сохраняет ровно пять свойств Excel, которыми владеет CExcelAppStateGuard.
Private Function TestFreshnessAppState(ByVal app As Object) As Variant
    TestFreshnessAppState = Array(app.EnableCancelKey, app.ScreenUpdating, app.EnableEvents, app.DisplayAlerts, app.Calculation)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет точное восстановление исходного состояния после штатного и ошибочного действия.
Private Sub TestFreshnessAssertState(ByRef stats As TUiTestStats, ByVal app As Object, ByVal expected As Variant, ByVal label As String)
    TestAssertTrue stats, label & ".cancel", app.EnableCancelKey = expected(0)
    TestAssertTrue stats, label & ".screen", app.ScreenUpdating = expected(1)
    TestAssertTrue stats, label & ".events", app.EnableEvents = expected(2)
    TestAssertTrue stats, label & ".alerts", app.DisplayAlerts = expected(3)
    TestAssertTrue stats, label & ".calculation", app.Calculation = expected(4)
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Находит реальную ячейку ключа Config, не привязывая тест к номеру строки листа.
Private Function TestFreshnessSystemCell(ByVal key As String) As Object
    Dim area As Object, row As Long, values As Variant
    Set area = ThisWorkbook.Names("rngSystemSettings").RefersToRange: values = area.Value2
    For row = 2 To UBound(values, 1)
        If CStr(values(row, 1)) = key Then Set TestFreshnessSystemCell = area.Cells(row, 2): Exit Function
    Next row
    Err.Raise 5, "TestFreshnessSystemCell", "Нет ключа " & key
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Находит ячейку параметра в текущей широкой таблице профилей: ключи по строкам,
' ProfileId по столбцам; старую длинную структуру fixture здесь не имитирует.
Private Function TestFreshnessProfileCell(ByVal profileID As String, ByVal key As String) As Object
    Dim area As Object, values As Variant, row As Long, column As Long, profileColumn As Long
    Set area = ThisWorkbook.Names("rngCalculationProfiles").RefersToRange: values = area.Value2
    For row = 1 To UBound(values, 1)
        For column = 3 To UBound(values, 2)
            If CStr(values(row, column)) = profileID Then profileColumn = column: Exit For
        Next column
        If profileColumn > 0 Then Exit For
    Next row
    If profileColumn = 0 Then Err.Raise 5, "TestFreshnessProfileCell", "Нет профиля " & profileID
    For row = 1 To UBound(values, 1)
        If CStr(values(row, 2)) = key Then Set TestFreshnessProfileCell = area.Cells(row, profileColumn): Exit Function
    Next row
    Err.Raise 5, "TestFreshnessProfileCell", "Нет параметра профиля " & key
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Вызывает настоящую границу действия и сохраняет исходный номер отказа.
' Никакую ошибку тест не заменяет успехом: ожидаемый код проверяется отдельно.
Private Function TestFreshnessInvoke(ByVal form As Boolean, ByVal workbook As Object, _
        ByRef message As String, ByRef informationEnabled As Boolean) As Long
    On Error GoTo Failed
    If form Then
        message = FormSectionForWorkbook(workbook, -1#, informationEnabled)
    Else
        message = RunSectionCalculationForWorkbook(workbook, False, -1#, informationEnabled)
    End If
    Exit Function
Failed:
    TestFreshnessInvoke = Err.Number: message = Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет одну явную актуализацию до любых имён и при выключенных событиях.
' Режим и остальные свойства должны совпасть с исходными после возврата из действия.
Private Sub TestFreshnessAssertAction(ByRef stats As TUiTestStats, ByVal proxy As CTestWorkflowWorkbook, _
        ByVal observed As CTestExcelApplicationState, ByVal expected As Variant, ByVal label As String)
    TestAssertTrue stats, label & ".oneApplicationCalculate", observed.TestCalculateCalls = 1
    TestAssertTrue stats, label & ".calculateInsideGuard", observed.TestCalculateGuarded
    TestAssertTrue stats, label & ".noInputBeforeCalculate", proxy.TestNamesBeforeCalculate = 0
    TestFreshnessAssertState stats, observed.TestNativeApplication, expected, label & ".restored"
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' На настоящей приватной книге проверяет свежесть межлистовой цепи и границы
' Form/Calculate для трёх исходных режимов. Экспорт, схема и очистка не получают
' Application.Calculate; снимок геометрии не меняется при расчёте после смены Config.
Public Function RunTestCalculationFreshnessWorkflow() As String
    Dim stats As TUiTestStats, saved As Variant, app As Object, initial As Variant, expected As Variant
    Dim helper As Object, mode As Variant, action As Variant, fault As Variant, prefix As String, label As String
    Dim observed As CTestExcelApplicationState, proxy As CTestWorkflowWorkbook
    Dim message As String, code As Long, informationEnabled As Boolean, solves As Long, beforeGeometry As Variant
    Dim target As Object, loads As Object, geometryCell As Object, profileCell As Object, flagCell As Object
    Set app = ThisWorkbook.Application: initial = TestFreshnessAppState(app)
    saved = TestFixtureSaveInputs(ThisWorkbook.Worksheets("Config").UsedRange)
    On Error GoTo Failed
    app.EnableEvents = False: app.DisplayAlerts = False
    Set helper = ThisWorkbook.Worksheets.Add
    prefix = "='" & Replace$(helper.Name, "'", "''") & "'!$C$1"
    Set loads = ThisWorkbook.Names("rngLoadCombinations").RefersToRange
    For Each mode In Array(xlCalculationAutomatic, xlCalculationManual, xlCalculationSemiautomatic)
        TestPrepareCircleInput
        TestSetSystemSetting "General.ExecutionReportEnabled", "No"
        TestSetSystemSetting "Plot.Enabled", "No"
        TestSetSystemSetting "Units.Force.Output", "N"
        TestSetSystemSetting "Mesh.StepMode", "User"
        TestSetSystemSetting "Mesh.StepX", "60": TestSetSystemSetting "Mesh.StepY", "60"
        TestSetSystemSetting "Mesh.BoundarySubdivisions", "1"
        TestSetSystemSetting "Solver.ToleranceN", "0.0001"
        TestSetProfileSetting "PR1", "Calculation.Strength.DirectState", "Yes"
        TestSetProfileSetting "PR1", "Calculation.Strength.Capacity", "No"
        TestSetProfileSetting "PR1", "Calculation.Crack.Width", "No"
        TestSetProfileSetting "PR1", "Calculation.Stability.Enabled", "No"
        TestSetProfileSetting "PR1", "Visualization.State", "StrengthState"
        loads.Cells(2, 5).Value2 = "PR1": loads.Cells(2, 3).Value2 = 0#: loads.Cells(2, 4).Value2 = 0#
        Set geometryCell = TestFixtureConfigFieldCell("Circle.Diameter")
        Set profileCell = TestFreshnessProfileCell("PR1", "Calculation.Strength.DirectState")
        Set flagCell = TestFreshnessSystemCell("General.NonCriticalMessagesEnabled")
        app.Calculation = xlCalculationManual
        helper.Range("A1").Value2 = 1#: helper.Range("B1").Formula = "=A1*3": helper.Range("C1").Formula = "=B1+1"
        geometryCell.Formula = prefix & "*10+200"
        flagCell.Formula = "=IF(" & Mid$(prefix, 2) & ">10,""Yes"",""No"")"
        app.Calculate ' Достоверный старый кэш240/No; дальнейшее изменение не вычисляется вручную тестом.
        helper.Range("A1").Value2 = 5#
        app.Calculation = mode: app.ScreenUpdating = (mode = xlCalculationAutomatic)
        app.EnableEvents = (mode <> xlCalculationManual): app.DisplayAlerts = (mode = xlCalculationSemiautomatic)
        app.EnableCancelKey = xlInterrupt: expected = TestFreshnessAppState(app)
        Set observed = New CTestExcelApplicationState: Set observed.TestNativeApplication = app
        Set proxy = New CTestWorkflowWorkbook: proxy.TestInitialize ThisWorkbook, observed
        solves = SectionEquilibriumSolveCount(): informationEnabled = False
        code = TestFreshnessInvoke(True, proxy, message, informationEnabled)
        label = "freshness.form." & CStr(mode)
        TestAssertTrue stats, label & ".success." & message, code = 0
        TestFreshnessAssertAction stats, proxy, observed, expected, label
        TestAssertClose stats, label & ".geometryFormula", CDbl(geometryCell.Value2), 360#, 0#
        TestAssertTrue stats, label & ".sameFreshUiFlag", informationEnabled
        TestAssertTrue stats, label & ".noSolver", SectionEquilibriumSolveCount() = solves
        beforeGeometry = TestResultTable("rngNDMSectionGeometry")

        app.Calculation = xlCalculationManual: app.EnableEvents = False
        helper.Range("A1").Value2 = 1#
        loads.Cells(2, 2).Formula = prefix & "*-1000"
        profileCell.Formula = "=IF(" & Mid$(prefix, 2) & ">10,""Yes"",""No"")"
        app.Calculate ' Старый кэшN=-4000 и профильDirectState=No; ожидается обновление обоих reader-ов.
        helper.Range("A1").Value2 = 6#
        TestSetSystemSetting "Geometry.Source", "AutoCAD": TestSetSystemSetting "Geometry.Type", "INVALID_UNUSED_RECIPE"
        app.Calculation = mode: app.EnableEvents = expected(2)
        Set observed = New CTestExcelApplicationState: Set observed.TestNativeApplication = app
        proxy.TestInitialize ThisWorkbook, observed
        code = TestFreshnessInvoke(False, proxy, message, informationEnabled)
        label = "freshness.calculate." & CStr(mode)
        TestAssertTrue stats, label & ".success." & message, code = 0 And InStr(message, "Расчет завершен") > 0
        TestFreshnessAssertAction stats, proxy, observed, expected, label
        TestAssertClose stats, label & ".freshLoad", CDbl(TestResultsPropertyValue("LC1", "N")), -19000#, 0.000001
        TestAssertTrue stats, label & ".freshProfileSolved", SectionEquilibriumSolveCount() > 0 ' Execute сбрасывает счётчики перед текущим пакетом; сравнение с предыдущим пакетом неверно.
        TestAssertTrue stats, label & ".directStateFinished", TestResultsPropertyValue("LC1", "DirectStateStatus") = "OK"
        Dim epsilonText As String, forceText As String
        epsilonText = TestResultsPropertyValue("LC1", "State.StrengthState.Epsilon0")
        forceText = TestResultsPropertyValue("LC1", "Nint")
        TestAssertTrue stats, label & ".savedStrainPresent", IsNumeric(epsilonText)
        If IsNumeric(epsilonText) Then TestAssertTrue stats, label & ".compressionStrain", CDbl(epsilonText) < 0#
        TestAssertTrue stats, label & ".savedEquilibriumPresent", IsNumeric(forceText)
        If IsNumeric(forceText) Then TestAssertClose stats, label & ".savedEquilibrium", CDbl(forceText), -19000#, 0.0001 ' Тот же явно заданный абсолютный допуск N, без ослабления равновесия ради успешного теста.
        TestAssertTrue stats, label & ".savedGeometryUnchanged", TestWorkflowSameTable(beforeGeometry, TestResultTable("rngNDMSectionGeometry"))

        observed.TestCalculateCalls = 0
        UpdateSectionPlotForWorkbook proxy ' Plot.Enabled=No: ранний выход не требует актуализации входов.
        UpdateSectionGeometryPreviewForWorkbook proxy
        TestAssertTrue stats, label & ".disabledPlotNoCalculate", observed.TestCalculateCalls = 0
        TestSetSystemSetting "Plot.Enabled", "Yes"
        solves = SectionEquilibriumSolveCount()
        message = Audit02ReadExportSnapshotForTests(proxy) ' Тот же read-path экспорта без запуска реального AutoCAD.
        UpdateSectionPlotForWorkbook proxy
        TestAssertTrue stats, label & ".exportPlotNoCalculate", observed.TestCalculateCalls = 0
        TestAssertTrue stats, label & ".exportPlotNoSolve", SectionEquilibriumSolveCount() = solves
        TestFreshnessAssertState stats, app, expected, label & ".exportPlotRestored"
        TestSetSystemSetting "Plot.Enabled", "No"

        For Each action In Array(True, False)
            Set observed = New CTestExcelApplicationState: Set observed.TestNativeApplication = app
            proxy.TestInitialize ThisWorkbook, observed: proxy.TestNamesFault = 1004
            code = TestFreshnessInvoke(CBool(action), proxy, message, informationEnabled)
            label = "freshness.expectedError." & mode & "." & CStr(action)
            TestAssertTrue stats, label & ".addressedReason", code = vbObjectError + 4316 And InStr(message, "rngUnitSettings") > 0
            TestFreshnessAssertAction stats, proxy, observed, expected, label
            For Each fault In Array(9, 18)
                Set observed = New CTestExcelApplicationState: Set observed.TestNativeApplication = app
                observed.TestCalculateFault = CLng(fault): proxy.TestInitialize ThisWorkbook, observed
                code = TestFreshnessInvoke(CBool(action), proxy, message, informationEnabled)
                label = "freshness.unexpectedError." & mode & "." & CStr(action) & "." & fault
                TestAssertTrue stats, label & ".originalNumber", code = CLng(fault)
                TestAssertTrue stats, label & ".noInputRead", proxy.TestNameReads = 0
                TestFreshnessAssertAction stats, proxy, observed, expected, label
            Next fault
        Next action
        For Each action In Array(True, False)
            code = TestFreshnessInvoke(CBool(action), Nothing, message, informationEnabled)
            TestAssertTrue stats, "freshness.nullWorkbook.rejected." & mode & "." & CStr(action), code <> 0
            TestFreshnessAssertState stats, app, expected, "freshness.nullWorkbook." & mode & "." & CStr(action)
        Next action
        Set observed = New CTestExcelApplicationState: Set observed.TestNativeApplication = app
        proxy.TestInitialize ThisWorkbook, observed
        ClearSectionResultsForWorkbook proxy
        TestAssertTrue stats, "freshness.clearNoCalculate." & mode, observed.TestCalculateCalls = 0
        TestFreshnessAssertState stats, app, expected, "freshness.clearRestored." & mode
        app.EnableEvents = False
    Next mode
    GoTo Finished
Failed:
    TestAssertTrue stats, "freshness.runtime." & Err.Number & "." & Err.Description, False
Finished:
    On Error Resume Next
    app.EnableEvents = False: app.DisplayAlerts = False: app.Calculation = xlCalculationManual
    TestFixtureRestoreInputs ThisWorkbook.Worksheets("Config").UsedRange, saved
    If Not helper Is Nothing Then helper.Delete
    app.Calculation = initial(4): app.DisplayAlerts = initial(3): app.EnableEvents = initial(2)
    app.ScreenUpdating = initial(1): app.EnableCancelKey = initial(0)
    On Error GoTo 0
    TestFreshnessAssertState stats, app, initial, "freshness.testCleanup"
    RunTestCalculationFreshnessWorkflow = stats.Report & "TOTAL_FRESHNESS_WORKFLOW: passed=" & stats.Passed & "; failed=" & stats.Failed
End Function


' -------------------- ДЛЯ ТЕСТОВ --------------------
' Запускает затронутые UI-проверки на независимой кандидатной книге.
Public Function TestRunSnapshotUiFailureProbe() As String
    Dim stats As TUiTestStats
    On Error GoTo Failed
    TestMissingProfileStateDrawsGeometryOnly stats
    TestMissingProfileStateDrawsGeometryOnly stats, "File"
    TestThirtyCombinationsWithFiveStatesWriteSnapshot stats
Done:
    TestAppendLine stats, "TOTAL SNAPSHOT UI PROBE: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    TestRunSnapshotUiFailureProbe = stats.Report
    Exit Function
Failed:
    TestAssertTrue stats, "probe.runtime." & CStr(Err.Number) & "." & Err.Description, False
    Resume Done
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Полностью отключённые LC не запускают решатель и сохраняют показанный снимок в обоих хранилищах.
Public Function RunLoadActivationWorkflowTests() As String
    Dim stats As TUiTestStats
    TestNoActiveLoadsPreserveResults stats
    RunLoadActivationWorkflowTests = stats.Report & "TOTAL LOAD ACTIVATION WORKFLOW: passed=" & stats.Passed & "; failed=" & stats.Failed & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет штатную команду без активных LC; входит также в общий workbook-набор.
Private Sub TestNoActiveLoadsPreserveResults(ByRef stats As TUiTestStats)
    Dim saved As Object, names As Variant, name As Variant, mode As Variant
    Set saved = CreateObject("Scripting.Dictionary")
    names = Array("rngSystemSettings", "rngCalculationProfiles", "rngUnitSettings", "rngSignConventionSettings", _
        "rngLoadCombinations", "rngStabilityDurationLoads", "rngCircleGeometry", "rngRectanglesGeometry", "rngPointedSectionGeometry", "rngHollowRectangleGeometry")
    On Error GoTo Failed
    For Each name In names: saved.Add CStr(name), ThisWorkbook.Names(CStr(name)).RefersToRange.Formula: Next name
    Dim store As CCalculationSnapshotStore: Set store = New CCalculationSnapshotStore
    Dim checksum As CFileChecksum: Set checksum = New CFileChecksum
    Dim file As String, oldHash As String, message As String, before As String, testCase As Long, loads As Object
    For Each mode In Array("Excel", "File")
        TestPrepareCircleInput
        TestSetSystemSetting "General.ExecutionReportEnabled", "No"
        TestSetSystemSetting "Results.SnapshotStorage", CStr(mode)
        TestSetSystemSetting "Results.SnapshotName", "TestLoadActivation_" & Format$(Now, "yyyymmdd_hhnnss") & ".ndm"
        message = FormSectionForWorkbook(ThisWorkbook)
        message = RunSectionCalculationForWorkbook(ThisWorkbook)
        before = modTestPerformance.TestPerformanceResultsFingerprint()
        If CStr(mode) = "File" Then file = store.SnapshotPath(ThisWorkbook): oldHash = checksum.FileSHA256(file)
        Set loads = ThisWorkbook.Names("rngLoadCombinations").RefersToRange
        For testCase = 0 To 1
            If testCase = 0 Then
                loads.Cells(2, 1).ClearContents
            Else
                loads.Cells(2, 1).Value2 = "ZERO": loads.Cells(2, 2).Resize(1, 3).Value2 = 0#
                loads.Cells(2, 5).Value2 = CVErr(xlErrNA): loads.Cells(2, 6).Value2 = "unknown"
            End If
            message = RunSectionCalculationForWorkbook(ThisWorkbook)
            TestAssertTrue stats, "loadActivation." & CStr(mode) & ".message." & testCase, InStr(1, message, "Расчёт не запускался", vbBinaryCompare) > 0
            TestAssertTrue stats, "loadActivation." & CStr(mode) & ".resultsPreserved." & testCase, before = modTestPerformance.TestPerformanceResultsFingerprint()
            TestAssertTrue stats, "loadActivation." & CStr(mode) & ".scopeReleased." & testCase, TestCalculationSnapshotReadReleased()
            If CStr(mode) = "File" Then TestAssertTrue stats, "loadActivation.File.bytesPreserved." & testCase, oldHash = checksum.FileSHA256(file)
        Next testCase
        If Len(file) > 0 Then Kill file: file = vbNullString
    Next mode
    GoTo Restore
Failed:
    TestAssertTrue stats, "loadActivation.workflow.runtime." & Err.Description, False
Restore:
    On Error Resume Next
    If Len(file) > 0 Then Kill file
    For Each name In saved.Keys: ThisWorkbook.Names(CStr(name)).RefersToRange.Formula = saved.Item(name): Next name
    On Error GoTo 0
End Sub
