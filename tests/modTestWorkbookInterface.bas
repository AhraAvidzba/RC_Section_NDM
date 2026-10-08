Attribute VB_Name = "modTestWorkbookInterface"
Option Explicit

' ==========================================================================
' Тесты структуры книги и пользовательского интерфейса
' ==========================================================================
' Модуль проверяет именованные диапазоны, выпадающие списки, листы и макросы
' workbook-слоя, чтобы сборка книги оставалась воспроизводимой.

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

' Выполняет COM-регрессии структуры Config/Results, workbook-макросов и схемы.
' Проверка снимков и export-data не является приемкой фактической записи DWG.
Public Function RunWorkbookInterfaceTests() As String
    On Error GoTo Failed

    Dim stats As TUiTestStats
    Dim t0 As Double
    Dim suiteSheet As Object
    Set suiteSheet = ThisWorkbook.Application.ActiveSheet
    t0 = Timer

    AppendLine stats, "RUN: TestButtons"
    TestButtons stats
    TestLoadCombinationsOnConfig stats
    AppendLine stats, "RUN: TestSingleCombinationSkipsBlankRows"
    TestSingleCombinationSkipsBlankRows stats
    TestLoadCombinationRangeMinimumRows stats
    TestConfigDropdownChoices stats
    TestConfigDropdownMeaning stats
    TestConfigSettingsRightBorder stats
    TestAutoCADGeometryContourWarning stats
    AppendLine stats, "RUN: TestAudit03ReaderContract"
    TestAudit03ReaderContract stats
    AppendLine stats, "RUN: TestAudit03InputContracts; " & Audit02ExcelMemory()
    TestAudit03InputContracts stats
    AppendLine stats, "RUN: TestAudit03ProfileInputContracts; " & Audit02ExcelMemory()
    TestAudit03ProfileInputContracts stats
    AppendLine stats, "RUN: TestAudit03NumericSettingsInputContracts; " & Audit02ExcelMemory()
    TestAudit03NumericSettingsInputContracts stats
    AppendLine stats, "RUN: TestAudit03SettingErrorMessages; " & Audit02ExcelMemory()
    TestAudit03SettingErrorMessages stats
    AppendLine stats, "RUN: TestAudit03RequiredTableMessages; " & Audit02ExcelMemory()
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
    AppendLine stats, "RUN: TestAudit03ConfigConversionMessages; " & Audit02ExcelMemory()
    TestAudit03ConfigConversionMessages stats
    AppendLine stats, "RUN: TestAudit03UnitSignConsumers; " & Audit02ExcelMemory()
    If mAudit03StopBeforeUnitSign Then
        AppendLine stats, "TOTAL_AUDIT03_UI_INPUT_PREFIX: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
        RunWorkbookInterfaceTests = stats.Report
        Exit Function
    End If
    ' Массовые Config-прогоны не требуют отображения постоянно меняющегося
    ' Results. Активный лист suite возвращаем перед окончательным отчетом.
    ThisWorkbook.Worksheets.Item("Config").Activate
    TestAudit03UnitSignConsumers stats
    AppendLine stats, "RUN: TestAudit03UnitSignChoices; " & Audit02ExcelMemory()
    TestAudit03UnitSignChoices stats
    AppendLine stats, "RUN: TestAudit03UnitSignEquivalence; " & Audit02ExcelMemory()
    TestAudit03UnitSignEquivalence stats
    AppendLine stats, "RUN: TestAudit03InputUnitConsumers; " & Audit02ExcelMemory()
    TestAudit03InputUnitConsumers stats
    AppendLine stats, "RUN: Audit03 RectSet selectors; " & Audit02ExcelMemory()
    TestRectSetIndependentSelectors stats
    TestRectSetIndependentSelectorLayout stats
    TestRectSetIndependentSelectorEffects stats
    TestRectSetIndependentThirdRows stats
    Dim circleConfigPassed As Long, circleConfigFailed As Long
    AppendLine stats, "RUN: Audit03 Circle Config; " & Audit02ExcelMemory()
    stats.Report = stats.Report & modTestGeometryConfig.RunAudit03CircleConfigTests(circleConfigPassed, circleConfigFailed)
    stats.Passed = stats.Passed + circleConfigPassed
    stats.Failed = stats.Failed + circleConfigFailed
    Dim circleInputPassed As Long, circleInputFailed As Long
    AppendLine stats, "RUN: Audit03 Circle rebar input; " & Audit02ExcelMemory()
    stats.Report = stats.Report & modTestGeometryConfig.RunAudit03CircleRebarInputTests(circleInputPassed, circleInputFailed)
    stats.Passed = stats.Passed + circleInputPassed
    stats.Failed = stats.Failed + circleInputFailed
    Dim shapeConfigPassed As Long, shapeConfigFailed As Long
    AppendLine stats, "RUN: Audit03 Shape Config; " & Audit02ExcelMemory()
    stats.Report = stats.Report & modTestGeometryConfig.RunAudit03ShapeConfigTests(shapeConfigPassed, shapeConfigFailed)
    stats.Passed = stats.Passed + shapeConfigPassed
    stats.Failed = stats.Failed + shapeConfigFailed
    Dim rebarCounterPassed As Long, rebarCounterFailed As Long
    AppendLine stats, "RUN: Audit03 Rebar counters; " & Audit02ExcelMemory()
    stats.Report = stats.Report & modTestGeometryConfig.RunAudit03RebarCounterTests(rebarCounterPassed, rebarCounterFailed)
    stats.Passed = stats.Passed + rebarCounterPassed
    stats.Failed = stats.Failed + rebarCounterFailed
    Dim requiredGeometryPassed As Long, requiredGeometryFailed As Long
    AppendLine stats, "RUN: Audit03 Required geometry; " & Audit02ExcelMemory()
    stats.Report = stats.Report & modTestGeometryConfig.RunAudit03RequiredGeometryInputTests(requiredGeometryPassed, requiredGeometryFailed)
    stats.Passed = stats.Passed + requiredGeometryPassed
    stats.Failed = stats.Failed + requiredGeometryFailed
    Dim rebarFieldPassed As Long, rebarFieldFailed As Long
    AppendLine stats, "RUN: Audit03 Rebar fields; " & Audit02ExcelMemory()
    stats.Report = stats.Report & modTestGeometryConfig.RunAudit03RebarFieldConfigTests(rebarFieldPassed, rebarFieldFailed)
    stats.Passed = stats.Passed + rebarFieldPassed
    stats.Failed = stats.Failed + rebarFieldFailed
    If mAudit03StopBeforeLoadTable Then
        AppendLine stats, "TOTAL_AUDIT03_UI_PRE_LOAD_TABLE: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
        RunWorkbookInterfaceTests = stats.Report
        Exit Function
    End If
    Dim loadTablePassed As Long, loadTableFailed As Long
    AppendLine stats, "RUN: Audit03 Load table; " & Audit02ExcelMemory()
    stats.Report = stats.Report & modTestLoadTableConfig.RunAudit03LoadTableTests(loadTablePassed, loadTableFailed)
    AppendLine stats, "RUN: Audit03 Load table completed; " & Audit02ExcelMemory()
    stats.Passed = stats.Passed + loadTablePassed
    stats.Failed = stats.Failed + loadTableFailed
    Dim profileScopePassed As Long, profileScopeFailed As Long
    AppendLine stats, "RUN: RunAudit03ProfileFailureScopeTests"
    stats.Report = stats.Report & modTestLoadTableConfig.RunAudit03ProfileFailureScopeTests(profileScopePassed, profileScopeFailed)
    stats.Passed = stats.Passed + profileScopePassed
    stats.Failed = stats.Failed + profileScopeFailed
    Dim resultWidthsPassed As Long, resultWidthsFailed As Long
    stats.Report = stats.Report & modTestLoadTableConfig.RunAudit03ResultColumnWidthTests(resultWidthsPassed, resultWidthsFailed)
    stats.Passed = stats.Passed + resultWidthsPassed
    stats.Failed = stats.Failed + resultWidthsFailed
    Dim profileConfigPassed As Long, profileConfigFailed As Long
    AppendLine stats, "RUN: RunAudit03ProfileConfigTests"
    stats.Report = stats.Report & modTestProfileConfig.RunAudit03ProfileConfigTests(profileConfigPassed, profileConfigFailed)
    stats.Passed = stats.Passed + profileConfigPassed
    stats.Failed = stats.Failed + profileConfigFailed
    Dim relocatedPassed As Long, relocatedFailed As Long
    AppendLine stats, "RUN: RunAudit03RelocatedInputTests"
    stats.Report = stats.Report & modTestProfileConfig.RunAudit03RelocatedInputTests(relocatedPassed, relocatedFailed)
    stats.Passed = stats.Passed + relocatedPassed
    stats.Failed = stats.Failed + relocatedFailed
    Dim durationPassed As Long, durationFailed As Long
    AppendLine stats, "RUN: RunAudit03DurationConfigTests"
    stats.Report = stats.Report & modTestProfileConfig.RunAudit03DurationConfigTests(durationPassed, durationFailed)
    stats.Passed = stats.Passed + durationPassed
    stats.Failed = stats.Failed + durationFailed
    Dim presentationPassed As Long, presentationFailed As Long
    AppendLine stats, "RUN: RunAudit03PresentationConfigTests"
    stats.Report = stats.Report & modTestProfileConfig.RunAudit03PresentationConfigTests(presentationPassed, presentationFailed)
    stats.Passed = stats.Passed + presentationPassed
    stats.Failed = stats.Failed + presentationFailed
    AppendLine stats, "RUN: TestPartialCombinationIsInvalid"
    TestPartialCombinationIsInvalid stats
    TestInvalidProfileIdDoesNotRunPlot stats
    TestAutoCADSourceRequiresManualImport stats
    TestAutoCADImportButtonRejectsGeneratedSource stats
    TestAutoCADImporterTreatsDrawingUnitsAsMillimeters stats
    TestAudit03InputAreaImportFilter stats
    TestAudit03ImportedSnapshotUnitChanges stats
    TestAudit03GeometrySnapshotContract stats
    TestAudit03ReadLifecycleContracts stats
    Dim plotConfigPassed As Long, plotConfigFailed As Long
    stats.Report = stats.Report & modTestPlotConfig.RunAudit03GeneralPlotTests(plotConfigPassed, plotConfigFailed)
    stats.Passed = stats.Passed + plotConfigPassed
    stats.Failed = stats.Failed + plotConfigFailed
    AppendLine stats, "RUN: TestAutoCADPreviewWritesAndDrawsBoundsDimensions"
    TestAutoCADPreviewWritesAndDrawsBoundsDimensions stats
    TestAnnotationDimensionTextRoundsInMillimeters stats
    TestPlotOverlayCoordinatesMatchResults stats
    AppendLine stats, "RUN: TestGeneratedSourceDoesNotReuseAutoCADPreview"
    TestGeneratedSourceDoesNotReuseAutoCADPreview stats
    AppendLine stats, "RUN: TestGeneratedDirectStateWorstStillDrawsFirstCalculatedLC"
    TestGeneratedDirectStateWorstStillDrawsFirstCalculatedLC stats
    AppendLine stats, "RUN: TestProfileDrivenPlotUsesSnapshotState"
    TestProfileDrivenPlotUsesSnapshotState stats
    TestAudit02SavedResultsIgnoreMaterialChanges stats
    AppendLine stats, "RUN: TestMissingProfileStateDrawsGeometryOnly"
    TestMissingProfileStateDrawsGeometryOnly stats
    TestAutoCADCalculationMessageUsesSavedGeometry stats
    TestBlankMomentDefaultsToZeroAndZeroLoadsAreSkipped stats
    AppendLine stats, "RUN: TestCircleWorkbookRunWritesResults"
    TestCircleWorkbookRunWritesResults stats
    AppendLine stats, "RUN: TestThirtyCombinationsWithFiveStatesWriteSnapshot"
    TestThirtyCombinationsWithFiveStatesWriteSnapshot stats
    TestDynamicLoadCombinationRangeAndLayoutGuard stats
    AppendLine stats, "RUN: TestExecutionReportFile"
    TestExecutionReportFile stats
    TestExcelApplicationStateGuardRestoresSettings stats
    AppendLine stats, "RUN: TestRectSetWorkbookRunWritesResults"
    TestRectSetWorkbookRunWritesResults stats
    AppendLine stats, "RUN: TestRectSetMomentUltimateStrainWorkbookPath"
    TestRectSetMomentUltimateStrainWorkbookPath stats
    TestStrengthSummaryUsesOutputCurvatureUnit stats
    AppendLine stats, "RUN: TestRectSetPureBendingUltimateStrainWorkbookPath"
    TestRectSetPureBendingUltimateStrainWorkbookPath stats
    AppendLine stats, "RUN: TestRectSetPureBendingDirectStateWorkbookPath"
    TestRectSetPureBendingDirectStateWorkbookPath stats
    AppendLine stats, "RUN: TestRectSetAxialTensionExtensionFromWorkbookSettings"
    TestRectSetAxialTensionExtensionFromWorkbookSettings stats
    TestAutoCADExportUsesSharedLoadReference stats
    AppendLine stats, "RUN: TestGoverningCombinationWritesDetailedResults"
    TestGoverningCombinationWritesDetailedResults stats
    TestCapacitySearchMethodValidation stats
    TestSolverToleranceUnitLabels stats
    TestCapacitySettingsUnitLabels stats
    TestRebarInputValidationDoesNotUseHiddenDefaults stats
    TestClearResultsKeepsInputs stats

    ' Настройки CAD проверяются без подключения к приложению. Настоящий DWG
    ' остается отдельным явным тестом с guard-ом собственных документов.
    Dim autoCADPassed As Long, autoCADFailed As Long
    AppendLine stats, "RUN: RunAudit03AutoCADConfigTests"
    stats.Report = stats.Report & modTestAutoCADConfig.RunAudit03AutoCADConfigTests(autoCADPassed, autoCADFailed)
    stats.Passed = stats.Passed + autoCADPassed
    stats.Failed = stats.Failed + autoCADFailed

    suiteSheet.Activate
    AssertTrue stats, "ui.suite.activeSheetRestored", ThisWorkbook.Application.ActiveSheet Is suiteSheet
    AppendLine stats, "TOTAL_WORKBOOK_UI: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed) & _
        "; elapsedSec=" & FormatNumberInvariant(Timer - t0)
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

    AppendLine stats, "TOTAL_LARGE_SNAPSHOT_PLOT: passed=" & CStr(stats.Passed) & _
        "; failed=" & CStr(stats.Failed) & "; elapsedSec=" & FormatNumberInvariant(Timer - t0)
    RunLargeSnapshotPlotStressTest = stats.Report
    Exit Function

Failed:
    RunLargeSnapshotPlotStressTest = stats.Report & "RUNTIME ERROR: " & CStr(Err.Number) & _
        "; source=" & Err.Source & "; description=" & Err.Description
End Function

' Проверяет, что таблица сочетаний живет на Config рядом с настройками,
' и не выводится отдельной параллельной таблицей на лист Расчет.
Private Sub TestLoadCombinationsOnConfig(ByRef stats As TUiTestStats)
    Dim loads As Object
    Set loads = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange

    AssertTrue stats, "ui.loads.layout.sheet", loads.Worksheet.Name = "Config"
    AssertTrue stats, "ui.loads.layout.position", loads.Row = 3 And loads.Column = 15
    AssertTrue stats, "ui.loads.layout.defaultRows30", loads.Rows.Count = 31
    AssertTrue stats, "ui.loads.layout.title", _
        InStr(1, CStr(loads.Worksheet.Cells.Item(2, 15).Value2), "Сочетания нагрузок", vbTextCompare) > 0
    Dim pathOptions As Variant
    pathOptions = Array("Auto", ChrW$(&H3BB) & "*Mx", ChrW$(&H3BB) & "*My", _
        ChrW$(&H3BB) & "*Mxy", ChrW$(&H3BB) & "*N", ChrW$(&H3BB) & "*NMxy")
    AssertTrue stats, "ui.loads.layout.capacityPathValidation", _
        LoadCombinationValidationHasOptions(6, pathOptions)
End Sub

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

    AssertTrue stats, "ui.buttons.run.exists", Not runButton Is Nothing
    AssertTrue stats, "ui.buttons.import.exists", Not importButton Is Nothing
    AssertTrue stats, "ui.buttons.clear.exists", Not clearButton Is Nothing
    AssertTrue stats, "ui.buttons.autocad.exists", Not acadButton Is Nothing
    AssertTrue stats, "ui.buttons.plot.exists", Not plotButton Is Nothing
    AssertTrue stats, "ui.buttons.run.macro", InStr(1, runButton.OnAction, "RunSectionCalculation", vbTextCompare) > 0
    AssertTrue stats, "ui.buttons.import.macro", InStr(1, importButton.OnAction, "ImportGeometryFromAutoCAD", vbTextCompare) > 0
    AssertTrue stats, "ui.buttons.clear.macro", InStr(1, clearButton.OnAction, "ClearAutoCADDrawing", vbTextCompare) > 0
    AssertTrue stats, "ui.buttons.autocad.macro", InStr(1, acadButton.OnAction, "ExportSectionStressToAutoCAD", vbTextCompare) > 0
    AssertTrue stats, "ui.buttons.plot.macro", InStr(1, plotButton.OnAction, "UpdateSectionPlot", vbTextCompare) > 0
    AssertTrue stats, "ui.buttons.outsidePrint", runButton.Left > calc.Range("AJ1").Left And importButton.Left > calc.Range("AJ1").Left And clearButton.Left > calc.Range("AJ1").Left And acadButton.Left > calc.Range("AJ1").Left And plotButton.Left > calc.Range("AJ1").Left
    AssertTrue stats, "ui.buttons.textCentered", ButtonTextIsCentered(runButton) And ButtonTextIsCentered(importButton) And _
        ButtonTextIsCentered(clearButton) And ButtonTextIsCentered(acadButton) And ButtonTextIsCentered(plotButton)
End Sub

' Проверяет оба выравнивания надписи кнопки; размер/положение фигуры этим не оцениваются.
Private Function ButtonTextIsCentered(ByVal buttonShape As Object) As Boolean
    ButtonTextIsCentered = (buttonShape.TextFrame.HorizontalAlignment = -4108 And _
        buttonShape.TextFrame.VerticalAlignment = -4108)
End Function

' Проверяет пользовательское правило: количество стержней не подставляется из
' кода, а локальное n=0 или d=0 просто отключает соответствующую линию.
Private Sub TestRebarInputValidationDoesNotUseHiddenDefaults(ByRef stats As TUiTestStats)
    PrepareCircleInput
    SetSystemSetting "Rebar.Diameter", vbNullString

    Dim errorText As String
    errorText = BuildCurrentWorkbookSectionError()
    AssertTrue stats, "ui.circle.blankDiameter.noHiddenDefaults", _
        InStr(1, errorText, "не задан ни один стержень", vbTextCompare) > 0

    PrepareRectSetInput
    SetSystemSetting "RectSet.H1.d_1", vbNullString

    Dim rectsetSection As CSectionModel
    Set rectsetSection = BuildCurrentWorkbookSection()
    AssertTrue stats, "ui.rectset.blankDiameterWithCount.skipsLine", rectsetSection.RebarCount = 7
    AssertTrue stats, "ui.rectset.blankDiameterWithCount.dimensionsRemain", rectsetSection.AnnotationCount >= 4

    PrepareRectSetInput
    SetSystemSetting "RectSet.H1.n_1", "0"
    SetSystemSetting "RectSet.H1.d_1", vbNullString

    Set rectsetSection = BuildCurrentWorkbookSection()
    AssertTrue stats, "ui.rectset.zeroCountBlankDiameter.skipsLine", rectsetSection.RebarCount = 7
    AssertTrue stats, "ui.rectset.zeroCountBlankDiameter.dimensionsRemain", rectsetSection.AnnotationCount >= 4

    SetSystemSetting "Geometry.Type", "RoundedRectangle"
    SetSystemSetting "RoundedRectangle.H.Left.n", vbNullString
    SetSystemSetting "RoundedRectangle.H.Right.n", vbNullString
    SetSystemSetting "RoundedRectangle.B.Top.n", vbNullString
    SetSystemSetting "RoundedRectangle.B.Bottom.n", vbNullString

    errorText = BuildCurrentWorkbookSectionError()
    AssertTrue stats, "ui.rounded.blankCounts.noHiddenDefaults", _
        InStr(1, errorText, "не задан ни один стержень", vbTextCompare) > 0

    PrepareCircleInput
End Sub

' ДЛЯ ТЕСТОВ: заданные имена опциональных контуров отсутствуют в fixture;
' импорт расчетной сетки должен продолжать работать без дополнительных слоев.
Private Function Audit03AbsentContourLayers() As Collection
    Set Audit03AbsentContourLayers = New Collection
End Function

' Собирает модель через тот же путь, которым пользуется кнопка расчета:
' Config -> CSystemSettingsReader -> CUnitSystem -> CSectionTypeRegistry.
Private Function BuildCurrentWorkbookSection() As CSectionModel
    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook

    Dim units As CUnitSystem
    Set units = New CUnitSystem
    units.LoadFromSettings settings

    Set BuildCurrentWorkbookSection = BuildWorkbookSectionModel(ThisWorkbook, settings, units)
End Function

' Возвращает текст ошибки при сборке модели или пустую строку, если модель
' построилась. Нужна для тестов пользовательской валидации Config без остановки
' всего набора workbook-тестов.
Private Function BuildCurrentWorkbookSectionError() As String
    On Error GoTo GotError
    Dim section As CSectionModel
    Set section = BuildCurrentWorkbookSection()
    BuildCurrentWorkbookSectionError = vbNullString
    Exit Function

GotError:
    BuildCurrentWorkbookSectionError = Err.Description
    Err.Clear
End Function

' Проверяет, что включенный общий флаг создает человекочитаемый txt-отчет
' рядом с книгой и не мешает обычному расчетному сценарию.
Private Sub TestExecutionReportFile(ByRef stats As TUiTestStats)
    PrepareCircleInput
    SetSystemSetting "General.ExecutionReportEnabled", "Yes"

    Dim reportPath As String
    reportPath = ThisWorkbook.Path & "\RC_Section_NDM_execution_report.txt"
    DeleteFileIfExists reportPath

    Dim message As String
    message = RunSectionCalculationForWorkbook(ThisWorkbook, False)

    AssertTrue stats, "ui.executionReport.message", InStr(1, message, "Пошаговый отчет сохранен", vbTextCompare) > 0
    AssertTrue stats, "ui.executionReport.fileExists", FileExists(reportPath)
    Dim reportText As String
    reportText = ReadTextFile(reportPath)
    AssertTrue stats, "ui.executionReport.content", InStr(1, reportText, "ОТЧЕТ ВЫПОЛНЕНИЯ RC SECTION NDM", vbTextCompare) > 0 And _
        InStr(1, reportText, "Старт расчета", vbTextCompare) > 0 And _
        InStr(1, reportText, "Итерации прямого CSectionSolver", vbTextCompare) > 0 And _
        InStr(1, reportText, "Финальное сообщение", vbTextCompare) > 0

    SetSystemSetting "General.ExecutionReportEnabled", "No"
End Sub

' Проверяет защиту пользовательского сценария расчета от лишней работы Excel.
' Макрос временно отключает события, экран, предупреждения и автопересчет,
' но после завершения должен вернуть настройки текущего Excel.Application
' ровно в исходное состояние.
Private Sub TestExcelApplicationStateGuardRestoresSettings(ByRef stats As TUiTestStats)
    PrepareCircleInput
    SetSystemSetting "General.ExecutionReportEnabled", "No"

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
    message = RunSectionCalculationForWorkbook(ThisWorkbook, False)

    AssertTrue stats, "ui.excelGuard.run", InStr(1, message, "Расчет", vbTextCompare) > 0
    AssertTrue stats, "ui.excelGuard.calculation", app.Calculation = oldCalculation
    AssertTrue stats, "ui.excelGuard.events", app.EnableEvents = oldEnableEvents
    AssertTrue stats, "ui.excelGuard.screen", app.ScreenUpdating = oldScreenUpdating
    AssertTrue stats, "ui.excelGuard.alerts", app.DisplayAlerts = oldDisplayAlerts
End Sub

' Проверяет общий перенос нагрузки к бетонному центру, используемый batch
' и export-данными. AutoCAD COM здесь не вызывается, DWG не создается.
Private Sub TestAutoCADExportUsesSharedLoadReference(ByRef stats As TUiTestStats)
    PrepareRectSetInput
    SetSystemSetting "Load.ReferenceOffsetX", "0"
    SetSystemSetting "Load.ReferenceOffsetY", "0"

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

    AssertTrue stats, "ui.autocad.reference.converged", _
        batch.ResultAt(1).StateRepository.FindState(sstCrackedState).Converged
    AssertTrue stats, "ui.autocad.reference.point", Abs(batch.LoadReferenceX) > 0.000001 Or Abs(batch.LoadReferenceY) > 0.000001
    AssertClose stats, "ui.autocad.reference.concreteCenterX", props.CentroidX, batch.LoadReferenceX, 0.000001
    AssertClose stats, "ui.autocad.reference.concreteCenterY", props.CentroidY, batch.LoadReferenceY, 0.000001
    Dim expectedLoad As CSectionLoadState
    Set expectedLoad = New CSectionLoadState
    expectedLoad.Initialize batch.N(1), batch.UserMx(1), batch.UserMy(1), props.CentroidX, props.CentroidY
    AssertClose stats, "ui.autocad.reference.mxTransfer", batch.Mx(1), expectedLoad.InternalMx, 0.000001
    AssertClose stats, "ui.autocad.reference.myTransfer", batch.My(1), expectedLoad.InternalMy, 0.000001
End Sub

' Читает один LC среди пустых строк: свободные строки не становятся
' ошибочными сочетаниями и не добавляют расчетов в пакет.
Private Sub TestSingleCombinationSkipsBlankRows(ByRef stats As TUiTestStats)
    PrepareCircleInput

    Dim batch As CBatchSectionCalculator
    Set batch = BuildUiBatch()

    Dim reader As CLoadCombinationReader
    Set reader = New CLoadCombinationReader
    reader.LoadFromWorkbook ThisWorkbook, batch

    Dim t0 As Double
    t0 = Timer
    batch.Execute

    AssertTrue stats, "ui.loads.single.count", batch.Count = 1
    AssertTrue stats, "ui.loads.single.id", batch.CombinationID(1) = "LC1"
    AssertTrue stats, "ui.loads.single.elapsed", (Timer - t0) < 20#
    AssertTrue stats, "ui.loads.single.noBlankInvalid", InStr(1, batch.DiagnosticLog, "InputErr", vbTextCompare) = 0
End Sub

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
    Set batch = BuildUiBatch()
    FillLoadCombinationTestRange tempSheet.Range("A1:G2"), 1, "PR1", "LC_MIN_"
    reader.LoadFromRange tempSheet.Range("A1:G2"), batch
    AssertTrue stats, "ui.loads.minRows.oneDataRow", batch.Count = 1

    Dim errorNumber As Long
    Dim errorText As String
    Set batch = BuildUiBatch()
    On Error Resume Next
    reader.LoadFromRange tempSheet.Range("A4:G4"), batch
    errorNumber = Err.Number
    errorText = Err.Description
    Err.Clear
    On Error GoTo Failed
    AssertTrue stats, "ui.loads.minRows.headerOnlyRejected", errorNumber <> 0 And _
        InStr(1, errorText, "минимум одну строку", vbTextCompare) > 0

CleanUp:
    On Error Resume Next
    If Not tempSheet Is Nothing Then tempSheet.Delete
    app.DisplayAlerts = oldDisplayAlerts
    On Error GoTo 0
    Exit Sub

Failed:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: ui.loads.minRows; err=" & CStr(Err.Number) & "; " & Err.Description
    Resume CleanUp
End Sub

' Удаляет ID из строки с заданными усилиями: строка сохраняется как InputErr,
' а пользовательский макрос выводит понятный итог ошибки ввода.
Private Sub TestPartialCombinationIsInvalid(ByRef stats As TUiTestStats)
    PrepareCircleInput
    Dim loads As Object
    Set loads = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange
    loads.Cells.Item(2, 1).ClearContents

    Dim batch As CBatchSectionCalculator
    Set batch = BuildUiBatch()

    Dim reader As CLoadCombinationReader
    Set reader = New CLoadCombinationReader
    reader.LoadFromWorkbook ThisWorkbook, batch
    batch.Execute

    AssertTrue stats, "ui.loads.partial.count", batch.Count = 1
    AssertTrue stats, "ui.loads.partial.invalid", batch.ResultAt(1).Status = "InputErr"

    Dim message As String
    message = RunSectionCalculationForWorkbook(ThisWorkbook, False)
    AssertTrue stats, "ui.loads.partial.message", InStr(1, message, "ошиб", vbTextCompare) > 0 And _
        InStr(1, message, "InputErr", vbTextCompare) > 0
End Sub

' При пустом ProfileId даже включенное автообновление не запускает вывод НДС.
' Проверяет сообщение InputErr и отсутствие строк element-results.
Private Sub TestInvalidProfileIdDoesNotRunPlot(ByRef stats As TUiTestStats)
    PrepareCircleInput
    SetSystemSetting "Plot.AutoUpdateAfterCalculation", "Yes"

    Dim loads As Object
    Set loads = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange
    loads.Cells.Item(2, 5).ClearContents

    Dim message As String
    message = RunSectionCalculationForWorkbook(ThisWorkbook, False)

    AssertTrue stats, "ui.loads.invalidProfileId.message", _
        InStr(1, message, "InputErr", vbTextCompare) > 0 And _
        InStr(1, message, "результатов элементов", vbTextCompare) = 0
    AssertTrue stats, "ui.loads.invalidProfileId.noElementRows", _
        ResultTableRowCount("rngNDMElementResults") = 1
End Sub

' Проверяет, что расчет в режиме AutoCAD не запускает скрытый импорт.
' Пользователь должен сначала явно нажать отдельную кнопку импорта; иначе
' расчет останавливается до очистки старых Results и до обращения к AutoCAD.
Private Sub TestAutoCADSourceRequiresManualImport(ByRef stats As TUiTestStats)
    PrepareCircleInput
    SetSystemSetting "Geometry.Source", "AutoCAD"
    ClearSectionResultsForWorkbook ThisWorkbook

    Dim message As String
    On Error Resume Next
    message = RunSectionCalculationForWorkbook(ThisWorkbook, False)
    Dim description As String
    description = Err.Description
    Err.Clear
    On Error GoTo 0

    AssertTrue stats, "ui.autocad.manualImport.required", _
        InStr(1, description, "предварительно импортированной", vbTextCompare) > 0 And _
        InStr(1, description, "Импортировать геометрию", vbTextCompare) > 0
End Sub

' Проверяет защиту отдельной кнопки импорта: если пользователь оставил
' Geometry.Source = Generated, макрос должен объяснить, что импорт недоступен
' в этом режиме, и не должен пытаться подключаться к AutoCAD.
Private Sub TestAutoCADImportButtonRejectsGeneratedSource(ByRef stats As TUiTestStats)
    PrepareCircleInput
    SetSystemSetting "Geometry.Source", "Generated"

    Dim message As String
    On Error Resume Next
    message = ImportGeometryFromAutoCADForWorkbook(ThisWorkbook)
    Dim description As String
    description = Err.Description
    Err.Clear
    On Error GoTo 0

    AssertTrue stats, "ui.autocad.import.generatedRejected", _
        InStr(1, description, "Geometry.Source = AutoCAD", vbTextCompare) > 0
End Sub

' Проверяет, что live AutoCAD-import не масштабирует координаты DWG через
' пользовательские INPUT-единицы. По контракту проекта AutoCAD-чертеж
' импортируется как миллиметры, иначе смена Units.Length.Input сдвинула бы
' весь imported snapshot относительно исходного чертежа.
Private Sub TestAutoCADImporterTreatsDrawingUnitsAsMillimeters(ByRef stats As TUiTestStats)
    PrepareCircleInput
    SetSystemSetting "Units.Length.Input", "m"
    SetSystemSetting "Units.Area.Input", "m2"

    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook

    Dim units As CUnitSystem
    Set units = New CUnitSystem
    units.LoadFromSettings settings

    Dim concreteRegion As CFakeAcadRegion
    Set concreteRegion = New CFakeAcadRegion
    concreteRegion.Initialize 2500#, 1000#, 2000#, 2500# * 2500# / 12#, _
        2500# * 2500# / 12#, 0#, "Concrete", "C1"

    Dim rebarRegion As CFakeAcadRegion
    Set rebarRegion = New CFakeAcadRegion
    rebarRegion.Initialize GEOM_PI * 25# * 25# / 4#, 1100#, 1900#, 1#, 1#, 0#, _
        "Reinf", "R1"

    Dim modelSpace As Collection
    Set modelSpace = New Collection
    modelSpace.Add concreteRegion
    modelSpace.Add rebarRegion

    Dim importer As CAutoCADSectionModelImporter
    Set importer = New CAutoCADSectionModelImporter

    Dim section As CSectionModel
    Set section = importer.ImportFromModelSpace(modelSpace, "Concrete", "Reinf", "Rebar", 0.000001, units)

    AssertClose stats, "ui.autocad.importUnits.concreteX", section.ConcreteX(1), 1000#, 0.000001
    AssertClose stats, "ui.autocad.importUnits.concreteY", section.ConcreteY(1), 2000#, 0.000001
    AssertClose stats, "ui.autocad.importUnits.concreteArea", section.ConcreteArea(1), 2500#, 0.000001
    AssertClose stats, "ui.autocad.importUnits.localIx", section.ConcreteLocalIx(1), _
        2500# * 2500# / 12#, 0.001
    AssertClose stats, "ui.autocad.importUnits.rebarX", section.RebarX(1), 1100#, 0.000001
    AssertClose stats, "ui.autocad.importUnits.rebarDiameter", section.RebarDiameter(1), 25#, 0.000001

    SetSystemSetting "Units.Length.Input", "mm"
    SetSystemSetting "Units.Area.Input", "mm2"
End Sub

' Проверяет, что подпись размера округляется до целого миллиметра до перевода
' в пользовательскую единицу вывода. При Units.Length.Output = m размер 517 мм
' должен отображаться как 0.517 m, а не как 1 m или 0 m.
Private Sub TestAnnotationDimensionTextRoundsInMillimeters(ByRef stats As TUiTestStats)
    Dim oldOutputLength As String
    oldOutputLength = GetSystemSetting("Units.Length.Output")

    On Error GoTo RestoreAndFail
    PrepareCircleInput
    SetSystemSetting "Geometry.Source", "AutoCAD"
    SetSystemSetting "Units.Length.Output", "m"

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
    annotationData = ResultTable("rngNDMSectionAnnotations")
    AssertTextEquals stats, "ui.autocad.preview.roundedWidthTextMeters", _
        AnnotationTextByID(annotationData, "DIM_AUTO_BOUNDS_B"), ChrW$(&H2248) & " 0.517 m"
    AssertTextEquals stats, "ui.autocad.preview.roundedHeightTextMeters", _
        AnnotationTextByID(annotationData, "DIM_AUTO_BOUNDS_H"), ChrW$(&H2248) & " 0.112 m"

Restore:
    SetSystemSetting "Units.Length.Output", oldOutputLength
    Exit Sub

RestoreAndFail:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: ui.autocad.preview.roundedDimensionTextMeters; " & Err.Description
    Resume Restore
End Sub

' Проверяет preview импортированной геометрии без реального AutoCAD.
' Важна вся цепочка: CSectionModel(AutoCADImport) -> Results annotations ->
' CSectionPlotDataReader -> CSectionPlotter -> Chart.Shapes размерных линий.
Private Sub TestAutoCADPreviewWritesAndDrawsBoundsDimensions(ByRef stats As TUiTestStats)
    PrepareCircleInput
    SetSystemSetting "Geometry.Source", "AutoCAD"

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
    annotationData = ResultTable("rngNDMSectionAnnotations")
    AssertTrue stats, "ui.autocad.preview.boundsAnnotations", CountAnnotationType(annotationData, "DIMENSION") = 2
    AssertTrue stats, "ui.autocad.preview.rebarAnnotation", CountAnnotationType(annotationData, "REBAR_ANNOTATION") = 1
    AssertTextEquals stats, "ui.autocad.preview.rebarText", _
        AnnotationTextByID(annotationData, "REBAR_AUTO_ALL"), "ALL: 1" & ChrW$(&H2205) & "32+2" & ChrW$(&H2205) & "20"
    AssertTrue stats, "ui.autocad.preview.approxText", _
        InStr(1, CStr(annotationData(2, ResultHeaderColumn(annotationData, "Text"))), ChrW$(&H2248), vbTextCompare) > 0
    AssertTextEquals stats, "ui.autocad.preview.roundedWidthText", _
        AnnotationTextByID(annotationData, "DIM_AUTO_BOUNDS_B"), ChrW$(&H2248) & " 517 mm"
    AssertTextEquals stats, "ui.autocad.preview.roundedHeightText", _
        AnnotationTextByID(annotationData, "DIM_AUTO_BOUNDS_H"), ChrW$(&H2248) & " 112 mm"
    AssertTrue stats, "ui.autocad.preview.russianComment", _
        InStr(1, CStr(annotationData(2, ResultHeaderColumn(annotationData, "Comment"))), "Приблизительная", vbTextCompare) > 0
    AssertClose stats, "ui.autocad.preview.boundsWidthUsesRotation", _
        AnnotationValueByID(annotationData, "DIM_AUTO_BOUNDS_B"), _
        450# + 0.5 * (120# * Cos(GEOM_PI / 6#) + 60# * Sin(GEOM_PI / 6#)), 0.001
    AssertClose stats, "ui.autocad.preview.boundsHeightUsesRotation", _
        AnnotationValueByID(annotationData, "DIM_AUTO_BOUNDS_H"), _
        120# * Sin(GEOM_PI / 6#) + 60# * Cos(GEOM_PI / 6#), 0.001

    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook

    Dim reader As CSectionPlotDataReader
    Set reader = New CSectionPlotDataReader
    reader.LoadGeometryPreviewFromWorkbook ThisWorkbook, settings
    AssertTrue stats, "ui.autocad.preview.readerAnnotations", reader.AnnotationCount = 3

    UpdateSectionPlotForWorkbook ThisWorkbook
    AssertTrue stats, "ui.autocad.preview.dimensionShapes", CountPlotShapes("AnnotationLine") > 0
    AssertTrue stats, "ui.autocad.preview.rebarTextShape", PlotShapeTextContains("ALL: 1")
    AssertTrue stats, "ui.autocad.preview.rotatedElementShape", _
        PlotShapeRotationExists("ElementConcrete", -30#, 0.5)
    Dim expectedFallbackDegrees As Double
    expectedFallbackDegrees = -0.5 * Atn(Sin(GEOM_PI / 3#) / (2# + Cos(GEOM_PI / 3#))) * 180# / GEOM_PI
    AssertTrue stats, "ui.autocad.preview.squareFallbackRotation", _
        PlotShapeRotationExists("ElementConcrete", expectedFallbackDegrees, 0.5)
End Sub


' Проверяет не только наличие осей/точки, но и их взаимное положение.
' Для режима Transformed центр главных осей берется из Transformed.Centroid,
' а точка нагрузки - из LoadReference. Их экранный сдвиг должен совпадать с
' расчетным сдвигом из Results после одного общего model-to-chart масштаба.
Private Sub TestPlotOverlayCoordinatesMatchResults(ByRef stats As TUiTestStats)
    Dim oldAxisLabelsEnabled As String
    Dim oldAxisLabelsFontSize As String
    oldAxisLabelsEnabled = GetSystemSetting("Plot.AxisLabelsEnabled")
    oldAxisLabelsFontSize = GetSystemSetting("Plot.AxisLabelsFontSize")

    PrepareUserRectSetMomentUltimateInput
    SetSystemSetting "Plot.PrincipalAxesMode", "Transformed"
    SetSystemSetting "Plot.LoadApplicationPointEnabled", "Yes"
    SetSystemSetting "Plot.AxisLabelsEnabled", "Yes"
    SetSystemSetting "Plot.AxisLabelsFontSize", "9"
    SetSystemSetting "Plot.ResultLabelsEnabled", "No"
    SetSystemSetting "Plot.AutoUpdateAfterCalculation", "Yes"

    Dim message As String
    message = RunSectionCalculationForWorkbook(ThisWorkbook, False)
    AssertTrue stats, "ui.plot.overlay.run", InStr(1, message, "Расчет завершен", vbTextCompare) > 0

    Dim loadX As Double
    Dim loadY As Double
    Dim centroidX As Double
    Dim centroidY As Double
    loadX = ResultsLengthPropertyMm("ALL", "LoadReferenceX")
    loadY = ResultsLengthPropertyMm("ALL", "LoadReferenceY")
    centroidX = ResultsLengthPropertyMm("ALL", "Transformed.CentroidX")
    centroidY = ResultsLengthPropertyMm("ALL", "Transformed.CentroidY")

    Dim loadPoint As Object
    Set loadPoint = FirstGeneratedPlotShape("LoadPoint")
    AssertTrue stats, "ui.plot.overlay.shapes", CountGeneratedPlotShapes("Principal1") > 0 And _
        CountGeneratedPlotShapes("Principal2") > 0 And Not loadPoint Is Nothing
    If CountGeneratedPlotShapes("Principal1") <= 0 Or CountGeneratedPlotShapes("Principal2") <= 0 Or loadPoint Is Nothing Then
        SetSystemSetting "Plot.AxisLabelsEnabled", oldAxisLabelsEnabled
        SetSystemSetting "Plot.AxisLabelsFontSize", oldAxisLabelsFontSize
        Exit Sub
    End If
    AssertTrue stats, "ui.plot.overlay.chartLayer", CountPlotShapes("Principal1") > 0 And _
        CountPlotShapes("Principal2") > 0 And CountPlotShapes("LoadPoint") > 0 And _
        CountWorksheetPlotShapes("Principal1") = 0 And CountWorksheetPlotShapes("Principal2") = 0 And _
        CountWorksheetPlotShapes("LoadPoint") = 0
    AssertTrue stats, "ui.plot.axisLabels.exists", CountPlotShapes("AxisLabelX") = 1 And CountPlotShapes("AxisLabelY") = 1
    AssertTrue stats, "ui.plot.axisLabels.text", PlotShapeTextContains("Ось X") And PlotShapeTextContains("Ось Y")
    AssertTrue stats, "ui.plot.axisLabels.italic", PlotShapeTextItalic("AxisLabelX") And PlotShapeTextItalic("AxisLabelY")
    AssertClose stats, "ui.plot.axisLabels.fontSize", PlotShapeTextFontSize("AxisLabelX"), 9#, 0.01
    AssertTrue stats, "ui.plot.axisLabels.insideFrame", PlotShapesInsideStableFrame("AxisLabelX", 0.5) And _
        PlotShapesInsideStableFrame("AxisLabelY", 0.5)

    Dim loadCenterX As Double
    Dim loadCenterY As Double
    AssertTrue stats, "ui.plot.overlay.loadCenter", GeneratedPlotShapeCenterAverage("LoadPoint", loadCenterX, loadCenterY)

    Dim plot As Object
    Set plot = ThisWorkbook.Worksheets.Item("Расчет").ChartObjects("chtNDMSectionPlot")
    AssertTrue stats, "ui.plot.overlay.clippedAxes", PlotShapesInsideStableFrame("Principal1", 1#) And _
        PlotShapesInsideStableFrame("Principal2", 1#)

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

    AssertTrue stats, "ui.plot.overlay.principal1ThroughCentroid", _
        PlotShapeBoundsContainAbsolutePoint("Principal1", expectedCentroidX, expectedCentroidY, 1.5)
    AssertTrue stats, "ui.plot.overlay.principal2ThroughCentroid", _
        PlotShapeBoundsContainAbsolutePoint("Principal2", expectedCentroidX, expectedCentroidY, 1.5)
    AssertClose stats, "ui.plot.overlay.loadX", loadCenterX, expectedLoadX, 0.8
    AssertClose stats, "ui.plot.overlay.loadY", loadCenterY, expectedLoadY, 0.8
    AssertClose stats, "ui.plot.overlay.loadDx", loadCenterX - expectedCentroidX, (loadX - centroidX) * scaleValue, 0.8
    AssertClose stats, "ui.plot.overlay.loadDy", loadCenterY - expectedCentroidY, -(loadY - centroidY) * scaleValue, 0.8

    Dim beforeMoveX As Double
    Dim beforeMoveY As Double
    Dim afterMoveX As Double
    Dim afterMoveY As Double
    Dim oldLeft As Double
    Dim oldTop As Double
    oldLeft = CDbl(plot.Left)
    oldTop = CDbl(plot.Top)
    AssertTrue stats, "ui.plot.overlay.move.before", GeneratedPlotShapeCenterAverage("LoadPoint", beforeMoveX, beforeMoveY)
    plot.Left = oldLeft + 17#
    plot.Top = oldTop + 11#
    AssertTrue stats, "ui.plot.overlay.move.after", GeneratedPlotShapeCenterAverage("LoadPoint", afterMoveX, afterMoveY)
    AssertClose stats, "ui.plot.overlay.move.dx", afterMoveX - beforeMoveX, 17#, 0.3
    AssertClose stats, "ui.plot.overlay.move.dy", afterMoveY - beforeMoveY, 11#, 0.3
    plot.Left = oldLeft
    plot.Top = oldTop
    SetSystemSetting "Plot.AxisLabelsEnabled", oldAxisLabelsEnabled
    SetSystemSetting "Plot.AxisLabelsFontSize", oldAxisLabelsFontSize
End Sub

' Проверяет, что AutoCAD-preview не используется как запасная схема
' после переключения Geometry.Source обратно на Generated. Иначе пользователь
' видит подпись "Импортированная геометрия AutoCAD" у уже generated-сценария.
Private Sub TestGeneratedSourceDoesNotReuseAutoCADPreview(ByRef stats As TUiTestStats)
    PrepareCircleInput
    SetSystemSetting "Geometry.Source", "AutoCAD"

    Dim section As CSectionModel
    Set section = New CSectionModel
    section.SourceType = "AutoCADImport"
    section.AddConcreteElement 50#, 50#, 10000#, 1, vbNullString, vbNullString, "Rectangle", 100#, 100#
    section.AddRebarElement 50#, 50#, 20#, 0#, "Rebar"

    Dim writer As CNDMResultsWriter
    Set writer = New CNDMResultsWriter
    writer.WriteGeometryPreview ThisWorkbook, section, PrepareSectionSnapshot(section)

    UpdateSectionPlotForWorkbook ThisWorkbook
    AssertTrue stats, "ui.plot.preview.title", PlotVisibleTitleContains("Импортированная геометрия AutoCAD")

    SetSystemSetting "Geometry.Source", "Generated"
    Dim errorDescription As String
    On Error Resume Next
    UpdateSectionPlotForWorkbook ThisWorkbook
    errorDescription = Err.Description
    On Error GoTo 0

    AssertTrue stats, "ui.plot.generated.noPreviewFallback.error", _
        InStr(1, errorDescription, "Выполните расчет", vbTextCompare) > 0
    AssertTrue stats, "ui.plot.generated.noPreviewFallback.title", _
        Not PlotVisibleTitleContains("Импортированная геометрия AutoCAD")
End Sub

' Проверяет DirectState-сценарий без определяющего сочетания по прочности.
' При Plot.LoadCase = Worst схема должна показать первый рассчитанный LC из
' Results, а не оставлять AutoCAD-preview и не очищаться до пустого окна.
Private Sub TestGeneratedDirectStateWorstStillDrawsFirstCalculatedLC(ByRef stats As TUiTestStats)
    PrepareCircleInput
    SetSystemSetting "Geometry.Source", "Generated"
    SetSystemSetting "Plot.LoadCase", "Worst"

    Dim message As String
    message = RunSectionCalculationForWorkbook(ThisWorkbook, False)

    AssertTrue stats, "ui.plot.generatedWorst.message", InStr(1, message, "Расчет завершен", vbTextCompare) > 0
    AssertTrue stats, "ui.plot.generatedWorst.drawsCalculatedLc", PlotVisibleTitleContains("LC1")
    AssertTrue stats, "ui.plot.generatedWorst.noImportTitle", _
        Not PlotVisibleTitleContains("Импортированная геометрия AutoCAD")
End Sub

' Проверяет, что схема после расчета выбирает состояние через профиль, а
' численные Stress/Strain берет из snapshot Results.
Private Sub TestProfileDrivenPlotUsesSnapshotState(ByRef stats As TUiTestStats)
    PrepareCircleInput
    SetSystemSetting "Plot.AutoUpdateAfterCalculation", "Yes"
    SetSystemSetting "Plot.LoadCase", "Worst"

    Dim message As String
    message = RunSectionCalculationForWorkbook(ThisWorkbook, False)

    AssertTrue stats, "ui.profilePlot.message", InStr(1, message, "Расчет завершен", vbTextCompare) > 0
    AssertTrue stats, "ui.profilePlot.hasElementStateRows", ResultTableRowCount("rngNDMElementResults") > 1
    AssertTrue stats, "ui.profilePlot.drawsLoadCase", PlotVisibleTitleContains("LC1")
    AssertTrue stats, "ui.profilePlot.noImportTitle", _
        Not PlotVisibleTitleContains("Импортированная геометрия AutoCAD")
    AssertTrue stats, "ui.profilePlot.commonLoadReference", _
        Len(ResultsPropertyValue("ALL", "LoadReferenceX")) > 0 And Len(ResultsPropertyValue("ALL", "LoadReferenceY")) > 0

End Sub

' Проверяет, что отсутствие выбранного StateType не превращается в popup-only
' ошибку. Схема должна показать сохраненную геометрию Results и крупную
' подпись под сечением, что запрошенного состояния в snapshot нет.
Private Sub TestMissingProfileStateDrawsGeometryOnly(ByRef stats As TUiTestStats)
    PrepareCircleInput
    SetSystemSetting "Plot.AutoUpdateAfterCalculation", "No"
    SetSystemSetting "Plot.LoadCase", "LC1"

    Dim message As String
    message = RunSectionCalculationForWorkbook(ThisWorkbook, False)
    AssertTrue stats, "ui.plot.missingState.run", InStr(1, message, "Расчет завершен", vbTextCompare) > 0

    SetProfileSetting "PR2", "Visualization.State", "CapacityState"

    Dim errorDescription As String
    On Error Resume Next
    UpdateSectionPlotForWorkbook ThisWorkbook
    errorDescription = Err.Description
    On Error GoTo 0

    AssertTrue stats, "ui.plot.missingState.noError", Len(errorDescription) = 0
    AssertTrue stats, "ui.plot.missingState.geometryTitle", PlotVisibleTitleContains("Геометрия расчетного сечения")
    AssertTrue stats, "ui.plot.missingState.note", PlotShapeTextContains("Запрашиваемое состояние")
End Sub

' Проверяет, что кнопка расчета в режиме AutoCAD использует уже сохраненную
' геометрию Results. Макрос не должен повторно импортировать Region и не должен
' показывать строку "Импортировано из AutoCAD", потому что это действие относится
' только к отдельной кнопке импорта геометрии.
Private Sub TestAutoCADCalculationMessageUsesSavedGeometry(ByRef stats As TUiTestStats)
    PrepareCircleInput

    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook

    Dim units As CUnitSystem
    Set units = New CUnitSystem
    units.LoadFromSettings settings

    Dim section As CSectionModel
    Set section = New CSectionModel
    section.SourceType = "AutoCADImport"
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
    beforeGeometry = ResultTable("rngNDMSectionGeometry")
    Dim beforeConcreteX As Double
    Dim beforeConcreteY As Double
    Dim beforeRebarX As Double
    Dim beforeRebarY As Double
    beforeConcreteX = GeometryResultValue(beforeGeometry, "C1", "X")
    beforeConcreteY = GeometryResultValue(beforeGeometry, "C1", "Y")
    beforeRebarX = GeometryResultValue(beforeGeometry, "R1", "X")
    beforeRebarY = GeometryResultValue(beforeGeometry, "R1", "Y")

    SetSystemSetting "Geometry.Source", "AutoCAD"
    Dim loads As Object
    Set loads = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange
    ClearDataRows loads
    loads.Cells.Item(2, 1).Value2 = "LC1"
    loads.Cells.Item(2, 2).Value2 = 0#
    loads.Cells.Item(2, 3).Value2 = -1000000#
    loads.Cells.Item(2, 4).Value2 = -750000#
    loads.Cells.Item(2, 5).Value2 = "PR1"
    loads.Cells.Item(2, 6).Value2 = ChrW$(&H3BB) & "*Mxy"
    loads.Cells.Item(2, 7).Value2 = "saved geometry"

    Dim message As String
    message = RunSectionCalculationForWorkbook(ThisWorkbook, False)

    AssertTrue stats, "ui.autocad.run.savedGeometry.message", _
        InStr(1, message, "Расчет импортированной из AutoCAD геометрии завершен", vbTextCompare) > 0
    AssertTrue stats, "ui.autocad.run.savedGeometry.noImportText", _
        InStr(1, message, "Импортировано из AutoCAD", vbTextCompare) = 0

    Dim afterGeometry As Variant
    afterGeometry = ResultTable("rngNDMSectionGeometry")
    AssertClose stats, "ui.autocad.run.savedGeometry.concreteX", _
        GeometryResultValue(afterGeometry, "C1", "X"), beforeConcreteX, 0.000001
    AssertClose stats, "ui.autocad.run.savedGeometry.concreteY", _
        GeometryResultValue(afterGeometry, "C1", "Y"), beforeConcreteY, 0.000001
    AssertClose stats, "ui.autocad.run.savedGeometry.rebarX", _
        GeometryResultValue(afterGeometry, "R1", "X"), beforeRebarX, 0.000001
    AssertClose stats, "ui.autocad.run.savedGeometry.rebarY", _
        GeometryResultValue(afterGeometry, "R1", "Y"), beforeRebarY, 0.000001
    Dim afterAnnotations As Variant
    afterAnnotations = ResultTable("rngNDMSectionAnnotations")
    AssertTextEquals stats, "ui.autocad.run.savedGeometry.rebarText", _
        AnnotationTextByID(afterAnnotations, "REBAR_AUTO_ALL"), "ALL: 4" & ChrW$(&H2205) & "20"
End Sub

' Проверяет разрешенный пустой My как 0 и пропуск достоверно нулевого LC.
' Пропуск нуля не должен скрывать ошибку или менять число ненулевых сочетаний.
Private Sub TestBlankMomentDefaultsToZeroAndZeroLoadsAreSkipped(ByRef stats As TUiTestStats)
    PrepareCircleInput
    Dim loads As Object
    Set loads = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange
    loads.Cells.Item(2, 4).ClearContents
    loads.Cells.Item(3, 1).Value2 = "ZERO"
    loads.Cells.Item(3, 2).Value2 = 0#
    loads.Cells.Item(3, 3).Value2 = 0#
    loads.Cells.Item(3, 4).Value2 = 0#
    loads.Cells.Item(3, 5).Value2 = "PR1"

    Dim batch As CBatchSectionCalculator
    Set batch = BuildUiBatch()

    Dim reader As CLoadCombinationReader
    Set reader = New CLoadCombinationReader
    reader.LoadFromWorkbook ThisWorkbook, batch
    batch.Execute

    AssertTrue stats, "ui.loads.blankMoment.count", batch.Count = 1
    AssertClose stats, "ui.loads.blankMoment.myZero", batch.UserMy(1), 0#, 0.0000001
    AssertTrue stats, "ui.loads.blankMoment.valid", InStr(1, batch.ResultAt(1).Status, "InputErr", vbTextCompare) = 0
End Sub

' Запускает полный workbook-сценарий круга и сверяет сохраненные таблицы,
' шапки, плоскость НДС, units и отсутствие дублированной геометрии элементов.
Private Sub TestCircleWorkbookRunWritesResults(ByRef stats As TUiTestStats)
    PrepareCircleInput
    Dim message As String
    message = RunSectionCalculationForWorkbook(ThisWorkbook, False)

    AssertTrue stats, "ui.run.message", InStr(1, message, "Расчет завершен", vbTextCompare) > 0
    Dim summary As Object
    Set summary = ThisWorkbook.Names.Item("rngBatchSummary").RefersToRange
    Dim firstDataRow As Long
    firstDataRow = BatchSummaryStartRow() + 12
    AssertTrue stats, "ui.run.capacity.na", CStr(summary.Worksheet.Cells.Item(firstDataRow, 7).Value2) = "N/A"
    AssertTrue stats, "ui.run.direct.status", Len(CStr(summary.Worksheet.Cells.Item(firstDataRow, 6).Value2)) > 0
    Dim strengthAnchor As Object
    Set strengthAnchor = ThisWorkbook.Names.Item("rngStrengthSummaryAnchor").RefersToRange
    AssertTrue stats, "ui.run.strainPlane", IsNumeric(strengthAnchor.Worksheet.Cells.Item(strengthAnchor.Row, 15).Value2) And _
        IsNumeric(strengthAnchor.Worksheet.Cells.Item(strengthAnchor.Row, 16).Value2) And _
        IsNumeric(strengthAnchor.Worksheet.Cells.Item(strengthAnchor.Row, 17).Value2)
    Dim crackAnchor As Object
    Set crackAnchor = ThisWorkbook.Names.Item("rngCrackSummaryAnchor").RefersToRange
    AssertTrue stats, "ui.run.crack", Len(CStr(crackAnchor.Worksheet.Cells.Item(crackAnchor.Row, 3).Value2)) > 0 And _
        Len(CStr(crackAnchor.Worksheet.Cells.Item(crackAnchor.Row, 18).Value2)) > 0
    Dim sys As Object
    Set sys = ThisWorkbook.Worksheets.Item("Config")
    AssertTrue stats, "ui.run.system.noLegacyMainInput", Not WorkbookNameExists("rngMainInput")
    AssertTrue stats, "ui.run.system.materialDiagramControls", _
        InStr(1, CStr(sys.Cells.Item(1, 35).Value2), "Контрольные точки диаграмм", vbTextCompare) > 0
    Dim resultsSheet As Object
    Set resultsSheet = ThisWorkbook.Worksheets.Item("Results")
    AssertTrue stats, "ui.strengthSummary.currentDepths", IsNumeric(resultsSheet.Cells.Item(strengthAnchor.Row, 26).Value2) And IsNumeric(resultsSheet.Cells.Item(strengthAnchor.Row, 27).Value2)
    AssertTrue stats, "ui.strengthSummary.direct.noCapacityDepths", Len(CStr(resultsSheet.Cells.Item(strengthAnchor.Row, 46).Value2)) = 0 And Len(CStr(resultsSheet.Cells.Item(strengthAnchor.Row, 47).Value2)) = 0
    AssertTrue stats, "ui.results.elements.header", CStr(ThisWorkbook.Names.Item("rngNDMElementResults").RefersToRange.Value2) = "RunID"
    Dim elementResults As Variant
    elementResults = ResultTable("rngNDMElementResults")
    AssertTrue stats, "ui.results.elements.rows", UBound(elementResults, 1) > 1
    AssertTrue stats, "ui.results.elements.noCombinationIndex", ResultHeaderColumn(elementResults, "CombinationIndex") = 0
    AssertTrue stats, "ui.results.elements.materialType", ResultHeaderColumn(elementResults, "MaterialType") > 0
    AssertTrue stats, "ui.results.elements.units", ResultHeaderColumn(elementResults, "Stress, MPa") > 0
    AssertTrue stats, "ui.results.elements.noGeometryDup", ResultHeaderColumn(elementResults, "X, mm") = 0
    AssertTrue stats, "ui.results.elements.noPlaneDup", ResultHeaderColumn(elementResults, "Epsilon0") = 0
    AssertTrue stats, "ui.results.elements.physicalStateAlign", _
        ThisWorkbook.Names.Item("rngNDMElementResults").RefersToRange.Offset(1, ResultHeaderColumn(elementResults, "PhysicalState") - 1).HorizontalAlignment = -4152
    AssertTrue stats, "ui.results.geometry.header", CStr(ThisWorkbook.Names.Item("rngNDMSectionGeometry").RefersToRange.Value2) = "RunID"
    Dim geometryResults As Variant
    geometryResults = ResultTable("rngNDMSectionGeometry")
    AssertTrue stats, "ui.results.geometry.rows", UBound(geometryResults, 1) > 1
    AssertTrue stats, "ui.results.strength.anchor", ThisWorkbook.Names.Item("rngStrengthSummaryAnchor").RefersToRange.Row = 49 And ThisWorkbook.Names.Item("rngStrengthSummaryAnchor").RefersToRange.Column = 1
    AssertTrue stats, "ui.results.crack.anchor", ThisWorkbook.Names.Item("rngCrackSummaryAnchor").RefersToRange.Row = 85 And ThisWorkbook.Names.Item("rngCrackSummaryAnchor").RefersToRange.Column = 1
    AssertTrue stats, "ui.results.stability.anchor", ThisWorkbook.Names.Item("rngStabilitySummaryAnchor").RefersToRange.Row = 122 And ThisWorkbook.Names.Item("rngStabilitySummaryAnchor").RefersToRange.Column = 1
    AssertTrue stats, "ui.results.geometry.position", ThisWorkbook.Names.Item("rngNDMSectionGeometry").RefersToRange.Row = 156 And ThisWorkbook.Names.Item("rngNDMSectionGeometry").RefersToRange.Column = 12
    AssertTrue stats, "ui.results.contours.position", ThisWorkbook.Names.Item("rngNDMSectionContours").RefersToRange.Row = 156 And ThisWorkbook.Names.Item("rngNDMSectionContours").RefersToRange.Column = 29
    AssertTrue stats, "ui.results.properties.position", ThisWorkbook.Names.Item("rngNDMSectionProperties").RefersToRange.Row = 156 And ThisWorkbook.Names.Item("rngNDMSectionProperties").RefersToRange.Column = 48
    AssertTrue stats, "ui.results.materialDiagrams.position", ThisWorkbook.Names.Item("rngNDMMaterialDiagrams").RefersToRange.Row = 156 And ThisWorkbook.Names.Item("rngNDMMaterialDiagrams").RefersToRange.Column = 56
    AssertTrue stats, "ui.results.annotations.position", ThisWorkbook.Names.Item("rngNDMSectionAnnotations").RefersToRange.Row = 156 And ThisWorkbook.Names.Item("rngNDMSectionAnnotations").RefersToRange.Column = 69
    AssertTrue stats, "ui.results.geometry.noSource", ResultHeaderColumn(geometryResults, "SourceName") = 0
    AssertTrue stats, "ui.results.geometry.noMaterialClass", ResultHeaderColumn(geometryResults, "MaterialClass") = 0
    AssertTrue stats, "ui.results.properties.header", CStr(ThisWorkbook.Names.Item("rngNDMSectionProperties").RefersToRange.Value2) = "RunID"
    AssertTrue stats, "ui.results.properties.hasStateEpsilon0", ResultsPropertyExists("LC1", "State.CrackedState.Epsilon0")
    AssertTrue stats, "ui.results.properties.hasStateNeutralLine", _
        ResultsPropertyExists("LC1", "State.CrackedState.NeutralLine.Angle") And _
        ResultsPropertyExists("LC1", "State.CrackedState.NeutralLine.SignedDistance")
    AssertTrue stats, "ui.results.properties.noLegacyPlane", _
        Not ResultsPropertyExists("LC1", "Epsilon0") And _
        Not ResultsPropertyExists("LC1", "KappaX") And _
        Not ResultsPropertyExists("LC1", "KappaY")
    AssertTrue stats, "ui.results.properties.hasBounds", ResultsPropertyExists("ALL", "Bounds.MinX")
    AssertTrue stats, "ui.results.properties.commonLoadReference", _
        ResultsPropertyExists("ALL", "LoadReferenceX") And ResultsPropertyExists("ALL", "LoadReferenceY")
    AssertTrue stats, "ui.results.properties.hasTransformedAxes", _
        ResultsPropertyExists("ALL", "Transformed.CentroidX") And _
        ResultsPropertyExists("ALL", "Transformed.CentroidY") And _
        ResultsPropertyExists("ALL", "Transformed.PrincipalAngle") And _
        ResultsPropertyExists("ALL", "Transformed.PrincipalI1") And _
        ResultsPropertyExists("ALL", "Transformed.PrincipalI2")
    AssertTrue stats, "ui.results.properties.noLegacyTransformedAliases", _
        Not ResultsPropertyExists("ALL", "CentroidX") And _
        Not ResultsPropertyExists("ALL", "CentroidY") And _
        Not ResultsPropertyExists("ALL", "PrincipalAngle") And _
        Not ResultsPropertyExists("ALL", "PrincipalI1") And _
        Not ResultsPropertyExists("ALL", "PrincipalI2")
    AssertTrue stats, "ui.results.properties.noLcLoadReference", _
        Not ResultsPropertyExists("LC1", "LoadReferenceX") And Not ResultsPropertyExists("LC1", "LoadReferenceY")
    AssertTransformedAreaUsesElasticModuli stats
    AssertTrue stats, "ui.results.annotations.header", CStr(ThisWorkbook.Names.Item("rngNDMSectionAnnotations").RefersToRange.Value2) = "RunID"
    AssertTrue stats, "ui.results.annotations.rows", ResultTableRowCount("rngNDMSectionAnnotations") > 1
    AssertTrue stats, "ui.results.materialDiagrams.header", CStr(ThisWorkbook.Names.Item("rngNDMMaterialDiagrams").RefersToRange.Value2) = "RunID"
    Dim materialDiagrams As Variant
    materialDiagrams = ResultTable("rngNDMMaterialDiagrams")
    AssertTrue stats, "ui.results.materialDiagrams.rows", UBound(materialDiagrams, 1) > 1
    AssertTrue stats, "ui.results.materialDiagrams.stateType", ResultHeaderColumn(materialDiagrams, "StateType") > 0
    AssertTrue stats, "ui.results.materialDiagrams.role", ResultHeaderColumn(materialDiagrams, "MaterialModelRole") > 0
    AssertTrue stats, "ui.results.materialDiagrams.diagramId", ResultHeaderColumn(materialDiagrams, "DiagramId") > 0
    AssertTrue stats, "ui.results.materialDiagrams.spec", ResultHeaderColumn(materialDiagrams, "MaterialModelSpec") > 0
    AssertTrue stats, "ui.results.materialDiagrams.mode", ResultHeaderColumn(materialDiagrams, "DiagramMode") > 0
    AssertTrue stats, "ui.results.materialDiagrams.noPurpose", ResultHeaderColumn(materialDiagrams, "Purpose") = 0
    AssertTrue stats, "ui.results.materialDiagrams.noLoadCase", ResultHeaderColumn(materialDiagrams, "LoadCase") = 0
    AssertTrue stats, "ui.results.materialDiagrams.usedStates", _
        MaterialDiagramIdsMatchStateProperties(materialDiagrams)
    AssertTrue stats, "ui.plot.chart.created", PlotChartExists()
    AssertTrue stats, "ui.plot.title.comment", PlotVisibleTitleContains("(ui test)")
    AssertTrue stats, "ui.plot.title.overlay", _
        PlotShapeTextContains("Схема сечения") And PlotShapeTextContains("(ui test)")

    Dim profiles As Object
    Set profiles = ThisWorkbook.Names.Item("rngCalculationProfiles").RefersToRange
    AssertTrue stats, "ui.profiles.topHeader.noDuplicateParameter", Len(Trim$(CStr(profiles.Cells.Item(1, 1).Value2))) = 0
    AssertTrue stats, "ui.profiles.lowerHeader.parameter", CStr(profiles.Cells.Item(2, 1).Value2) = "Параметр"
    AssertTrue stats, "ui.profiles.pr3.default", _
        Len(ProfileSettingValue("PR3", "Profile.DisplayName")) > 0 And _
        ProfileSettingValue("PR3", "Calculation.Strength.DirectState") = "Yes" And _
        ProfileSettingValue("PR3", "Calculation.Crack.Width") = "Yes"
    AssertTrue stats, "ui.profiles.pr4.default", _
        Len(ProfileSettingValue("PR4", "Profile.DisplayName")) > 0 And _
        ProfileSettingValue("PR4", "Visualization.Quantity") = "Strain" And _
        ProfileSettingValue("PR4", "Visualization.StrainPrecision") = "6"
    AssertTrue stats, "ui.profiles.visualizationStressPrecision.default", _
        ProfileSettingValue("PR1", "Visualization.StressPrecision") = "1"
End Sub

' Проверяет полный предельный snapshot: 30 сочетаний, каждое с пятью
' конечными named-state. Для устойчивого получения PreCrackState/PostCrackState
' используется чистый изгиб: сжатие может подавить образование нормальной
' трещины и тогда эти состояния физически не обязаны появляться.
Private Sub TestThirtyCombinationsWithFiveStatesWriteSnapshot(ByRef stats As TUiTestStats)
    PrepareUserRectSetMomentUltimateInput
    PrepareFullStateProfile "PR3"
    SetSystemSetting "SLS.Crack.PsiMode", "Auto"
    SetSystemSetting "SLS.Crack.Allowable", "0.000001"
    SetSystemSetting "Plot.AutoUpdateAfterCalculation", "No"
    SetSystemSetting "Plot.LoadCase", "LC_FULL_01"

    Dim loads As Object
    Set loads = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange
    ClearDataRows loads

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
    message = RunSectionCalculationForWorkbook(ThisWorkbook, False)

    AssertTrue stats, "ui.results.fullSnapshot.message", _
        InStr(1, message, "Расчет завершен", vbTextCompare) > 0

    Dim geometryRowCount As Long
    geometryRowCount = ResultTableRowCount("rngNDMSectionGeometry") - 1
    AssertTrue stats, "ui.results.fullSnapshot.geometry143", geometryRowCount = 143

    Dim elementResults As Variant
    elementResults = ResultTable("rngNDMElementResults")
    Dim expectedElementRows As Long
    expectedElementRows = 1 + 30 * 5 * geometryRowCount
    AssertTrue stats, "ui.results.fullSnapshot.elementRows", _
        UBound(elementResults, 1) = expectedElementRows

    AssertElementStateRows stats, elementResults, "LC_FULL_01", "StrengthState", geometryRowCount
    AssertElementStateRows stats, elementResults, "LC_FULL_01", "CapacityState", geometryRowCount
    AssertElementStateRows stats, elementResults, "LC_FULL_01", "CrackedState", geometryRowCount
    AssertElementStateRows stats, elementResults, "LC_FULL_01", "PreCrackState", geometryRowCount
    AssertElementStateRows stats, elementResults, "LC_FULL_01", "PostCrackState", geometryRowCount
    AssertElementStateRows stats, elementResults, "LC_FULL_30", "StrengthState", geometryRowCount
    AssertElementStateRows stats, elementResults, "LC_FULL_30", "CapacityState", geometryRowCount
    AssertElementStateRows stats, elementResults, "LC_FULL_30", "CrackedState", geometryRowCount
    AssertElementStateRows stats, elementResults, "LC_FULL_30", "PreCrackState", geometryRowCount
    AssertElementStateRows stats, elementResults, "LC_FULL_30", "PostCrackState", geometryRowCount

    AssertTrue stats, "ui.results.fullSnapshot.propertiesRows", _
        ResultTableRowCount("rngNDMSectionProperties") >= 1 + 55 + 30 * (23 + 5 * 10)
    AssertTrue stats, "ui.results.fullSnapshot.materialRows", _
        ResultTableRowCount("rngNDMMaterialDiagrams") > 1
End Sub

' Проверяет, что высота rngLoadCombinations управляет числом LC, а расчет
' заранее останавливается, если текущие якоря Results не оставляют места для
' всех строк, шапок и выводимых столбцов.
Private Sub TestDynamicLoadCombinationRangeAndLayoutGuard(ByRef stats As TUiTestStats)
    On Error GoTo Failed

    PrepareCircleInput
    SetSystemSetting "Plot.AutoUpdateAfterCalculation", "No"

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
    Set batch = BuildUiBatch()
    Dim reader As CLoadCombinationReader
    Set reader = New CLoadCombinationReader
    reader.LoadFromWorkbook ThisWorkbook, batch
    AssertTrue stats, "ui.loads.dynamicRange.count45", batch.Count = 45

    Dim errorNumber As Long
    Dim errorText As String
    On Error Resume Next
    Dim message As String
    message = RunSectionCalculationForWorkbook(ThisWorkbook, False)
    errorNumber = Err.Number
    errorText = Err.Description
    Err.Clear
    On Error GoTo Failed

    AssertTrue stats, "ui.results.layoutGuard.blocksCalculation", errorNumber <> 0
    AssertTrue stats, "ui.results.layoutGuard.message", _
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
    AppendLine stats, "FAIL: ui.results.layoutGuard; err=" & CStr(Err.Number) & "; " & Err.Description
    Resume CleanUp
End Sub

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

' ДЛЯ ТЕСТОВ
' Проверяет тяжелую схему с мелкой сеткой. Расчетные строки остаются полными,
' а plotter обязан рисовать элементы поэлементно Shape-ами до безопасного
' предела и не раздувать число Excel Chart series.
Private Sub TestLargeSnapshotPlotStress(ByRef stats As TUiTestStats)
    PrepareUserRectSetMomentUltimateInput
    PrepareFullStateProfile "PR3"
    SetSystemSetting "Mesh.StepX", "10"
    SetSystemSetting "Mesh.StepY", "10"
    SetSystemSetting "Mesh.BoundarySubdivisions", "1"
    SetSystemSetting "SLS.Crack.PsiMode", "Auto"
    SetSystemSetting "SLS.Crack.Allowable", "0.000001"
    SetSystemSetting "Plot.AutoUpdateAfterCalculation", "Yes"
    SetSystemSetting "Plot.LoadCase", "LC_BIG_01"

    Dim loads As Object
    Set loads = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange
    ClearDataRows loads

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
    message = RunSectionCalculationForWorkbook(ThisWorkbook, False)

    Dim geometryRowCount As Long
    geometryRowCount = ResultTableRowCount("rngNDMSectionGeometry") - 1

    Dim elementResults As Variant
    elementResults = ResultTable("rngNDMElementResults")

    Dim seriesCount As Long
    seriesCount = PlotSeriesCount()
    Dim elementShapeCount As Long
    elementShapeCount = CountPlotShapes("Element")

    AppendLine stats, "INFO: ui.largeSnapshot.geometryRows=" & CStr(geometryRowCount) & _
        "; elementResultRows=" & CStr(UBound(elementResults, 1)) & _
        "; plotSeries=" & CStr(seriesCount) & _
        "; elementShapes=" & CStr(elementShapeCount)

    AssertTrue stats, "ui.largeSnapshot.message", _
        InStr(1, message, "Расчет завершен", vbTextCompare) > 0
    AssertTrue stats, "ui.largeSnapshot.geometryLarge", geometryRowCount > 2500
    AssertTrue stats, "ui.largeSnapshot.elementRows", _
        UBound(elementResults, 1) = 1 + 20 * 5 * geometryRowCount
    AssertTrue stats, "ui.largeSnapshot.plotCreated", PlotChartExists()
    AssertTrue stats, "ui.largeSnapshot.elementShapes", elementShapeCount >= geometryRowCount
    AssertTrue stats, "ui.largeSnapshot.plotSeriesLimit", seriesCount < 256
End Sub

' Запускает RectSet через пользовательский workbook-entrypoint и проверяет
' фактический Results, блок диаграмм и отсутствие отдельного ручного RebarInput.
Private Sub TestRectSetWorkbookRunWritesResults(ByRef stats As TUiTestStats)
    PrepareRectSetInput
    Dim message As String
    message = RunSectionCalculationForWorkbook(ThisWorkbook, False)

    AssertTrue stats, "ui.rectset.message", InStr(1, message, "завершен", vbTextCompare) > 0
    Dim summary As Object
    Set summary = ThisWorkbook.Names.Item("rngBatchSummary").RefersToRange
    AssertTrue stats, "ui.rectset.result.status", Len(CStr(summary.Worksheet.Cells.Item(BatchSummaryStartRow() + 12, 1).Value2)) > 0

    Dim sys As Object
    Set sys = ThisWorkbook.Worksheets.Item("Config")
    AssertTrue stats, "ui.rectset.system.noLegacyRebarInput", Not WorkbookNameExists("rngRebarInput")
    AssertTrue stats, "ui.rectset.system.materialDiagramControls", _
        InStr(1, CStr(sys.Cells.Item(1, 35).Value2), "Контрольные точки диаграмм", vbTextCompare) > 0
End Sub

' Проверяет пользовательский сценарий Г-сечения с N + Mx и выбранной
' траекторией lambda*Mx через полный путь книги. Этот тест защищает быстрый
' UltimateStrain-путь для обычного изгибного расчета: после универсализации
' LoadPath он не должен уходить в тяжелую общую residual-систему.
Private Sub TestRectSetMomentUltimateStrainWorkbookPath(ByRef stats As TUiTestStats)
    PrepareUserRectSetMomentUltimateInput

    Dim message As String
    message = RunSectionCalculationForWorkbook(ThisWorkbook, False)
    ThisWorkbook.Application.CalculateFull

    Dim resultsSheet As Object
    Set resultsSheet = ThisWorkbook.Worksheets.Item("Results")
    Dim firstRow As Long
    firstRow = StrengthSummaryStartRow()

    Dim capacityStatus As String
    Dim solutionMethod As String
    capacityStatus = CStr(resultsSheet.Cells.Item(firstRow, 49).Value2)
    solutionMethod = CStr(resultsSheet.Cells.Item(firstRow, 37).Value2)

    AppendLine stats, "INFO: ui.rectset.momentUltimate capacityStatus=" & capacityStatus & _
        "; solutionMethod=" & solutionMethod & _
        "; lambda=" & CStr(resultsSheet.Cells.Item(firstRow, 38).Value2)

    AssertTrue stats, "ui.rectset.momentUltimate.message", InStr(1, message, "Расчет завершен", vbTextCompare) > 0
    AssertTextEquals stats, "ui.rectset.momentUltimate.capacityOk", capacityStatus, "OK"
    AssertTextEquals stats, "ui.rectset.momentUltimate.method", solutionMethod, "UltimateStrain"
    AssertTrue stats, "ui.rectset.momentUltimate.lambda", _
        CDbl(resultsSheet.Cells.Item(firstRow, 38).Value2) > 0#
    AssertTrue stats, "ui.rectset.momentUltimate.mxult", _
        Abs(CDbl(resultsSheet.Cells.Item(firstRow, 40).Value2)) > 0#
End Sub

' Проверяет, что подробный блок прочности не держит кривизны в 1/мм
' жестко, а использует OUTPUT-единицу из Units.Curvature.Output.
Private Sub TestStrengthSummaryUsesOutputCurvatureUnit(ByRef stats As TUiTestStats)
    Dim oldOutputCurvature As String
    oldOutputCurvature = GetSystemSetting("Units.Curvature.Output")

    On Error GoTo RestoreAndFail
    PrepareUserRectSetMomentUltimateInput
    SetSystemSetting "Units.Curvature.Output", "1/m"

    Dim message As String
    message = RunSectionCalculationForWorkbook(ThisWorkbook, False)
    ThisWorkbook.Application.CalculateFull

    Dim resultsSheet As Object
    Set resultsSheet = ThisWorkbook.Worksheets.Item("Results")
    Dim firstRow As Long
    firstRow = StrengthSummaryStartRow()
    Dim labelRow As Long
    labelRow = firstRow - 1

    AssertTrue stats, "ui.strength.curvatureOutput.message", _
        InStr(1, message, "Расчет завершен", vbTextCompare) > 0
    AssertTextEquals stats, "ui.strength.curvatureOutput.snapshotUnit", _
        ResultsPropertyValue("ALL", "Output.CurvatureUnit"), "1/m"
    AssertTrue stats, "ui.strength.curvatureOutput.stateHeaderKx", _
        InStr(1, CStr(resultsSheet.Cells.Item(labelRow, 16).Value2), "1/m", vbTextCompare) > 0
    AssertTrue stats, "ui.strength.curvatureOutput.stateHeaderKy", _
        InStr(1, CStr(resultsSheet.Cells.Item(labelRow, 17).Value2), "1/m", vbTextCompare) > 0
    AssertTrue stats, "ui.strength.curvatureOutput.capacityHeaderKx", _
        InStr(1, CStr(resultsSheet.Cells.Item(labelRow, 33).Value2), "1/m", vbTextCompare) > 0
    AssertTrue stats, "ui.strength.curvatureOutput.capacityHeaderKy", _
        InStr(1, CStr(resultsSheet.Cells.Item(labelRow, 34).Value2), "1/m", vbTextCompare) > 0

    AssertClose stats, "ui.strength.curvatureOutput.stateKxValue", _
        CDbl(resultsSheet.Cells.Item(firstRow, 16).Value2), _
        CDbl(ResultsPropertyValue("LC_MX", "State.StrengthState.KappaX")), 0.000000000001
    AssertClose stats, "ui.strength.curvatureOutput.capacityKxValue", _
        CDbl(resultsSheet.Cells.Item(firstRow, 33).Value2), _
        CDbl(ResultsPropertyValue("LC_MX", "State.CapacityState.KappaX")), 0.000000000001

Restore:
    SetSystemSetting "Units.Curvature.Output", oldOutputCurvature
    Exit Sub

RestoreAndFail:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: ui.strength.curvatureOutput; " & Err.Description
    Resume Restore
End Sub

' Проверяет чистый изгиб Г-сечения по полному Excel-пути.
' Здесь N намеренно равен нулю: программа должна найти прямое НДС и
' предельный момент без перехода в осевую или аварийную ветку.
Private Sub TestRectSetPureBendingUltimateStrainWorkbookPath(ByRef stats As TUiTestStats)
    PrepareUserRectSetMomentUltimateInput

    Dim loads As Object
    Set loads = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange
    loads.Cells.Item(2, 2).Value2 = 0#
    loads.Cells.Item(2, 7).Value2 = "pure bending ultimate regression"

    Dim message As String
    message = RunSectionCalculationForWorkbook(ThisWorkbook, False)
    ThisWorkbook.Application.CalculateFull

    Dim resultsSheet As Object
    Set resultsSheet = ThisWorkbook.Worksheets.Item("Results")
    Dim firstRow As Long
    firstRow = StrengthSummaryStartRow()

    Dim directStatus As String
    Dim capacityStatus As String
    Dim solutionMethod As String
    directStatus = CStr(resultsSheet.Cells.Item(firstRow, 30).Value2)
    capacityStatus = CStr(resultsSheet.Cells.Item(firstRow, 49).Value2)
    solutionMethod = CStr(resultsSheet.Cells.Item(firstRow, 37).Value2)

    AppendLine stats, "INFO: ui.rectset.pureBending direct=" & directStatus & _
        "; capacity=" & capacityStatus & _
        "; solutionMethod=" & solutionMethod & _
        "; lambda=" & CStr(resultsSheet.Cells.Item(firstRow, 38).Value2)

    AssertTrue stats, "ui.rectset.pureBending.message", InStr(1, message, "Расчет завершен", vbTextCompare) > 0
    AssertTextEquals stats, "ui.rectset.pureBending.directOk", directStatus, "OK"
    AssertTextEquals stats, "ui.rectset.pureBending.capacityOk", capacityStatus, "OK"
    AssertTextEquals stats, "ui.rectset.pureBending.method", solutionMethod, "UltimateStrain"
    AssertTrue stats, "ui.rectset.pureBending.lambda", _
        CDbl(resultsSheet.Cells.Item(firstRow, 38).Value2) > 0#
End Sub

' Проверяет чистый изгиб Г-сечения по профилю PR1. Такой профиль запрашивает
' и прямое НДС, и несущую способность, поэтому обе ветви должны проходить
' через полный workbook-path без специальных обходов.
Private Sub TestRectSetPureBendingDirectStateWorkbookPath(ByRef stats As TUiTestStats)
    PrepareUserRectSetMomentUltimateInput

    Dim loads As Object
    Set loads = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange
    loads.Cells.Item(2, 2).Value2 = 0#
    loads.Cells.Item(2, 7).Value2 = "pure bending direct regression"

    Dim message As String
    message = RunSectionCalculationForWorkbook(ThisWorkbook, False)
    ThisWorkbook.Application.CalculateFull

    Dim resultsSheet As Object
    Set resultsSheet = ThisWorkbook.Worksheets.Item("Results")
    Dim firstRow As Long
    firstRow = StrengthSummaryStartRow()

    Dim directStatus As String
    Dim capacityStatus As String
    directStatus = CStr(resultsSheet.Cells.Item(firstRow, 30).Value2)
    capacityStatus = CStr(resultsSheet.Cells.Item(firstRow, 49).Value2)

    AppendLine stats, "INFO: ui.rectset.pureBendingDirect direct=" & directStatus & _
        "; capacity=" & capacityStatus

    AssertTrue stats, "ui.rectset.pureBendingDirect.message", InStr(1, message, "Расчет завершен", vbTextCompare) > 0
    AssertTextEquals stats, "ui.rectset.pureBendingDirect.directOk", directStatus, "OK"
    AssertTextEquals stats, "ui.rectset.pureBendingDirect.capacityOk", capacityStatus, "OK"
End Sub

' Проверяет пользовательский сценарий с сильным осевым растяжением Г-сечения
' через настоящий workbook-path: Config -> CUnitSystem -> batch -> Results.
' Это важно, потому что знак N и единицы tf здесь проходят ровно тем же путем,
' что и при нажатии кнопки "Расчет" в книге.
Private Sub TestRectSetAxialTensionExtensionFromWorkbookSettings(ByRef stats As TUiTestStats)
    PrepareUserRectSetAxialTensionInput

    Dim reportPath As String
    reportPath = ThisWorkbook.Path & "\RC_Section_NDM_execution_report.txt"
    DeleteFileIfExists reportPath

    Dim message As String
    message = RunSectionCalculationForWorkbook(ThisWorkbook, False)

    Dim resultsSheet As Object
    Set resultsSheet = ThisWorkbook.Worksheets.Item("Results")
    Dim safeSummaryRow As Long
    Dim overSummaryRow As Long
    Dim safeCrackRow As Long
    Dim overCrackRow As Long
    safeSummaryRow = BatchSummaryRowByCombination(resultsSheet, "LC_SAFE")
    overSummaryRow = BatchSummaryRowByCombination(resultsSheet, "LC_OVER")
    safeCrackRow = CrackSummaryRowByCombination(resultsSheet, "LC_SAFE")
    overCrackRow = CrackSummaryRowByCombination(resultsSheet, "LC_OVER")

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
    overExtension = ResultsPropertyValue("LC_OVER", "ExtensionUsed")

    AppendLine stats, "INFO: ui.rectset.axial795 overall=" & safeOverall & _
        "; crack=" & safeCrack
    AppendLine stats, "INFO: ui.rectset.axial900 overall=" & overOverall & _
        "; crack=" & overCrack & _
        "; extensionUsed=" & overExtension

    AssertTrue stats, "ui.rectset.axial795.overallCalculated", _
        safeOverall = "OK" Or safeOverall = "FAIL"
    AssertTrue stats, "ui.rectset.axial795.crackCalculated", _
        safeCrack = "OK" Or safeCrack = "FAIL"

    AssertTrue stats, "ui.rectset.axial900.finishedWithoutExtension", _
        overOverall = "FAIL" Or overOverall = "NumFail"
    AssertTrue stats, "ui.rectset.axial900.crackNoNumFail", _
        overCrack <> "NumFail" And overCrack <> "InputErr"
    AssertTextEquals stats, "ui.rectset.axial900.crackedStateExtUsed", _
        overCrackExtUsed, "yes"
    AssertTextEquals stats, "ui.rectset.axial900.crackedStateEquilibrium", _
        overCrackEquilibrium, "FAIL"
    AssertTextEquals stats, "ui.rectset.axial900.crackWidthSkipped", _
        overCrack, "N/A"
    AssertTextEquals stats, "ui.rectset.axial900.longitudinalSkipped", _
        overLongitudinal, "N/A"
    AssertTrue stats, "ui.rectset.axial900.message", InStr(1, message, "Расчет завершен", vbTextCompare) > 0
    AssertTextEquals stats, "ui.rectset.axial900.extensionSnapshot", overExtension, "True"
    AssertTrue stats, "ui.rectset.axial900.crackedStateExtensionDiagram", _
        MaterialDiagramHasMode("CrackedState", "StateExtension")
    AssertTrue stats, "ui.rectset.axial900.reportCreated", FileExists(reportPath)

    SetSystemSetting "General.ExecutionReportEnabled", "No"
End Sub

' Решает безопасное и более тяжелое LC с критерием governing StrengthCapacity.
' Сверяет выбранный ID с запасами в настоящей сводке и итоговым сообщением.
Private Sub TestGoverningCombinationWritesDetailedResults(ByRef stats As TUiTestStats)
    PrepareCircleInput
    SetSystemSetting "General.WorstCombinationCriterion", "StrengthCapacity"
    SetSystemSetting "Capacity.SolutionStrategy", "Auto"
    SetSystemSetting "Capacity.MaxLambda", "1024"
    SetSystemSetting "Plot.LoadCase", "Worst"
    SetProfileSetting "PR1", "Calculation.Strength.DirectState", "Yes"
    SetProfileSetting "PR1", "Calculation.Strength.Capacity", "Yes"
    SetProfileSetting "PR1", "Calculation.Crack.Width", "No"
    SetProfileSetting "PR1", "Calculation.Stability.Enabled", "No"

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
    message = RunSectionCalculationForWorkbook(ThisWorkbook, False)
    ThisWorkbook.Application.CalculateFull

    Dim resultsSheet As Object
    Set resultsSheet = ThisWorkbook.Worksheets.Item("Results")
    Dim summaryRow As Long
    summaryRow = BatchSummaryStartRow()
    Dim governingID As String
    governingID = CStr(resultsSheet.Cells.Item(summaryRow, 1).Value2)
    Dim expectedID As String
    expectedID = ExpectedGoverningByLowestStrengthSafety(resultsSheet, summaryRow)
    AppendLine stats, "INFO: ui.governing actual=" & governingID & "; expected=" & expectedID

    AssertTrue stats, "ui.governing.id", governingID = expectedID
    AssertTrue stats, "ui.governing.message", InStr(1, message, governingID, vbTextCompare) > 0
End Sub

' Независимо выбирает LC по минимальному положительному запасу в записанном summary.
' Пустые строки пропускаются; производственный выбор governing здесь не вызывается.
Private Function ExpectedGoverningByLowestStrengthSafety(ByVal resultsSheet As Object, ByVal summaryRow As Long) As String
    Dim rowIndex As Long
    Dim bestSafety As Double
    For rowIndex = summaryRow + 12 To summaryRow + 30
        If Len(Trim$(CStr(resultsSheet.Cells.Item(rowIndex, 1).Value2))) > 0 Then
            Dim safetyValue As Double
            safetyValue = StrengthSafetyForSummaryRow(resultsSheet, rowIndex)
            If safetyValue > 0# And (bestSafety = 0# Or safetyValue < bestSafety) Then
                bestSafety = safetyValue
                ExpectedGoverningByLowestStrengthSafety = CStr(resultsSheet.Cells.Item(rowIndex, 1).Value2)
            End If
        End If
    Next rowIndex
End Function

' Читает численный запас подробной строки; пустое/текстовое поле не участвует в минимуме.
Private Function StrengthSafetyForSummaryRow(ByVal resultsSheet As Object, ByVal rowIndex As Long) As Double
    If IsNumeric(resultsSheet.Cells.Item(rowIndex, 17).Value2) Then _
        StrengthSafetyForSummaryRow = CDbl(resultsSheet.Cells.Item(rowIndex, 17).Value2)
End Function

' Сопоставляет LC с подробным выводом и собирает модуль из записанных Mx/My.
' Отсутствующий LC возвращает ноль, который направленный тест отличает от нужной точки.
Private Function MomentUltimateForCombination(ByVal resultsSheet As Object, ByVal summaryRow As Long, ByVal combinationID As String) As Double
    Dim rowIndex As Long
    For rowIndex = summaryRow + 12 To summaryRow + 30
        If StrComp(CStr(resultsSheet.Cells.Item(rowIndex, 1).Value2), combinationID, vbTextCompare) = 0 Then
            Dim mxUltimate As Double
            Dim myUltimate As Double
            Dim strengthRow As Long
            strengthRow = StrengthSummaryRowByCombination(resultsSheet, combinationID)
            If strengthRow > 0 Then
                If IsNumeric(resultsSheet.Cells.Item(strengthRow, 33).Value2) Then mxUltimate = CDbl(resultsSheet.Cells.Item(strengthRow, 33).Value2)
                If IsNumeric(resultsSheet.Cells.Item(strengthRow, 34).Value2) Then myUltimate = CDbl(resultsSheet.Cells.Item(strengthRow, 34).Value2)
            End If
            MomentUltimateForCombination = Sqr(mxUltimate * mxUltimate + myUltimate * myUltimate)
            Exit Function
        End If
    Next rowIndex
End Function

' Координаты трех summary берутся из именованных anchors, не из фиксированной верстки.
Private Function BatchSummaryStartRow() As Long
    BatchSummaryStartRow = ThisWorkbook.Names.Item("rngBatchSummary").RefersToRange.Row
End Function

' Находит фактическую строку LC в batch summary с учетом промежутков исходной таблицы.
' Ноль обозначает отсутствие ID, а не первую строку результатов.
Private Function BatchSummaryRowByCombination(ByVal resultsSheet As Object, ByVal combinationID As String) As Long
    Dim anchorRow As Long
    anchorRow = BatchSummaryStartRow()
    Dim rowIndex As Long
    For rowIndex = anchorRow + 12 To anchorRow + 200
        If StrComp(CStr(resultsSheet.Cells.Item(rowIndex, 1).Value2), combinationID, vbTextCompare) = 0 Then
            BatchSummaryRowByCombination = rowIndex
            Exit Function
        End If
    Next rowIndex
End Function

' Возвращает строку strength anchor для остальных readback-проверок.
Private Function StrengthSummaryStartRow() As Long
    StrengthSummaryStartRow = ThisWorkbook.Names.Item("rngStrengthSummaryAnchor").RefersToRange.Row
End Function

' Возвращает строку crack anchor для остальных readback-проверок.
Private Function CrackSummaryStartRow() As Long
    CrackSummaryStartRow = ThisWorkbook.Names.Item("rngCrackSummaryAnchor").RefersToRange.Row
End Function

' Ищет ID в подробной таблице прочности; пропуски между сочетаниями не уплотняются.
Private Function StrengthSummaryRowByCombination(ByVal resultsSheet As Object, ByVal combinationID As String) As Long
    Dim anchorRow As Long
    anchorRow = StrengthSummaryStartRow()
    Dim rowIndex As Long
    For rowIndex = anchorRow To anchorRow + 200
        If StrComp(CStr(resultsSheet.Cells.Item(rowIndex, 1).Value2), combinationID, vbTextCompare) = 0 Then
            StrengthSummaryRowByCombination = rowIndex
            Exit Function
        End If
    Next rowIndex
End Function

' Ищет ID в подробной таблице трещин независимо от позиции в batch summary.
Private Function CrackSummaryRowByCombination(ByVal resultsSheet As Object, ByVal combinationID As String) As Long
    Dim anchorRow As Long
    anchorRow = CrackSummaryStartRow()
    Dim rowIndex As Long
    For rowIndex = anchorRow To anchorRow + 200
        If StrComp(CStr(resultsSheet.Cells.Item(rowIndex, 1).Value2), combinationID, vbTextCompare) = 0 Then
            CrackSummaryRowByCombination = rowIndex
            Exit Function
        End If
    Next rowIndex
End Function

' Определяет колонку по точному тексту первой строки массива; отсутствующий заголовок дает ноль.
Private Function ResultHeaderColumn(ByRef data As Variant, ByVal headerText As String) As Long
    Dim colIndex As Long
    For colIndex = 1 To UBound(data, 2)
        If StrComp(CStr(data(1, colIndex)), headerText, vbTextCompare) = 0 Then
            ResultHeaderColumn = colIndex
            Exit Function
        End If
    Next colIndex
End Function

' Проверяет, что конкретное состояние конкретного LC записано для каждого
' расчетного элемента сечения. Это защищает snapshot от частичной записи.
Private Sub AssertElementStateRows(ByRef stats As TUiTestStats, ByRef data As Variant, _
        ByVal loadCase As String, ByVal stateType As String, ByVal expectedCount As Long)
    Dim count As Long
    count = ElementStateRowCount(data, loadCase, stateType)
    AssertTrue stats, "ui.results.fullSnapshot." & loadCase & "." & stateType, count = expectedCount
End Sub

' Считает строки rngNDMElementResults для пары LoadCase + StateType.
Private Function ElementStateRowCount(ByRef data As Variant, ByVal loadCase As String, _
        ByVal stateType As String) As Long
    Dim loadCaseColumn As Long
    Dim stateTypeColumn As Long
    loadCaseColumn = ResultHeaderColumn(data, "LoadCase")
    stateTypeColumn = ResultHeaderColumn(data, "StateType")
    If loadCaseColumn = 0 Or stateTypeColumn = 0 Then Exit Function

    Dim rowIndex As Long
    For rowIndex = 2 To UBound(data, 1)
        If StrComp(CStr(data(rowIndex, loadCaseColumn)), loadCase, vbTextCompare) = 0 And _
                StrComp(CStr(data(rowIndex, stateTypeColumn)), stateType, vbTextCompare) = 0 Then
            ElementStateRowCount = ElementStateRowCount + 1
        End If
    Next rowIndex
End Function

' Возвращает максимальное по модулю Stress для сохраненного named-state.
' Тест защищает snapshot от рассинхрона: если StateSolution найден через
' numerical extension, строки элементов должны соответствовать той же диаграмме.
Private Function MaxAbsElementStress(ByRef data As Variant, ByVal loadCase As String, _
        ByVal stateType As String) As Double
    Dim loadCaseColumn As Long
    Dim stateTypeColumn As Long
    Dim stressColumn As Long
    loadCaseColumn = ResultHeaderColumn(data, "LoadCase")
    stateTypeColumn = ResultHeaderColumn(data, "StateType")
    stressColumn = ResultHeaderColumnByBaseName(data, "Stress")
    If loadCaseColumn = 0 Or stateTypeColumn = 0 Or stressColumn = 0 Then Exit Function

    Dim rowIndex As Long
    Dim stressValue As Double
    For rowIndex = 2 To UBound(data, 1)
        If StrComp(CStr(data(rowIndex, loadCaseColumn)), loadCase, vbTextCompare) = 0 And _
                StrComp(CStr(data(rowIndex, stateTypeColumn)), stateType, vbTextCompare) = 0 Then
            stressValue = Abs(CDbl(data(rowIndex, stressColumn)))
            If stressValue > MaxAbsElementStress Then MaxAbsElementStress = stressValue
        End If
    Next rowIndex
End Function

' Проверяет, есть ли в диагностической таблице диаграмм хотя бы одна диаграмма
' заданного named-state с указанным режимом Physical/StateExtension.
Private Function MaterialDiagramHasMode(ByVal stateTypeText As String, ByVal diagramModeText As String) As Boolean
    Dim materialDiagrams As Variant
    materialDiagrams = ResultTable("rngNDMMaterialDiagrams")

    Dim stateCol As Long
    Dim modeCol As Long
    stateCol = ResultHeaderColumn(materialDiagrams, "StateType")
    modeCol = ResultHeaderColumn(materialDiagrams, "DiagramMode")
    If stateCol = 0 Or modeCol = 0 Then Exit Function

    Dim rowIndex As Long
    For rowIndex = 2 To UBound(materialDiagrams, 1)
        If StrComp(CStr(materialDiagrams(rowIndex, stateCol)), stateTypeText, vbTextCompare) = 0 And _
                StrComp(CStr(materialDiagrams(rowIndex, modeCol)), diagramModeText, vbTextCompare) = 0 Then
            MaterialDiagramHasMode = True
            Exit Function
        End If
    Next rowIndex
End Function

' Проверяет, что каталог фактических диаграмм связан с named-state metadata.
' Диаграммы пишутся один раз по DiagramId, а состояния ссылаются на них через
' State.*.ConcreteDiagramId и State.*.RebarDiagramId в rngNDMSectionProperties.
Private Function MaterialDiagramIdsMatchStateProperties(ByRef materialDiagrams As Variant) As Boolean
    Dim diagramIds As Object
    Set diagramIds = CreateObject("Scripting.Dictionary")
    diagramIds.CompareMode = vbTextCompare

    Dim pointKeys As Object
    Set pointKeys = CreateObject("Scripting.Dictionary")
    pointKeys.CompareMode = vbTextCompare

    Dim diagramCol As Long
    Dim pointCol As Long
    diagramCol = ResultHeaderColumn(materialDiagrams, "DiagramId")
    pointCol = ResultHeaderColumn(materialDiagrams, "PointIndex")
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
    props = ResultTable("rngNDMSectionProperties")
    Dim parameterCol As Long
    Dim valueCol As Long
    parameterCol = ResultHeaderColumn(props, "Parameter")
    valueCol = ResultHeaderColumn(props, "Value")
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

    MaterialDiagramIdsMatchStateProperties = True
End Function

' Сверяет фактические validation-списки расчетов, геометрии, plot и export,
' динамические ссылки и default labels. Наличие списка не доказывает поведение каждой опции.
Private Sub TestCapacitySearchMethodValidation(ByRef stats As TUiTestStats)
    AssertTrue stats, "ui.validation.CapacitySolutionStrategy", _
        SystemSettingValidationHasOptions("Capacity.SolutionStrategy", Array("Auto", "UltimateStrain", "LoadMultiplier"))
    AssertTrue stats, "ui.validation.nonCriticalMessages", _
        SystemSettingValidationHasOptions("General.NonCriticalMessagesEnabled", Array("Yes", "No"))
    AssertTrue stats, "ui.validation.capacitySearchMethod", _
        SystemSettingValidationHasOptions("Capacity.SearchMethod", Array("Bisection", "Brent", "Secant"))
    AssertTrue stats, "ui.validation.crackCoverDistanceMode", _
        SystemSettingValidationHasOptions("SLS.Crack.CoverDistanceMode", Array("NearestContour", "GlobalExtreme"))
    AssertTrue stats, "ui.validation.universalLoadPath", UniversalLoadPathValidationHasOptions()
    AssertTrue stats, "ui.validation.autocadLabelMode", _
        SystemSettingValidationHasOptions("AutoCAD.Export.LabelMode", Array("ValuesOnly", "NamesAndValues"))
    AssertTrue stats, "ui.validation.autocadNeutralLine", _
        SystemSettingValidationHasOptions("AutoCAD.Export.NeutralLineEnabled", Array("Yes", "No"))
    AssertTrue stats, "ui.validation.autocadPrincipalAxes", _
        SystemSettingValidationHasOptions("AutoCAD.Export.PrincipalAxesMode", Array("Transformed", "Concrete", "None"))
    AssertTrue stats, "ui.validation.autocadLoadPoint", _
        SystemSettingValidationHasOptions("AutoCAD.Export.LoadPointEnabled", Array("Yes", "No"))
    AssertTrue stats, "ui.validation.autocadContour", _
        SystemSettingValidationHasOptions("AutoCAD.Export.ContourEnabled", Array("Yes", "No"))
    AssertTrue stats, "ui.validation.autocadCombination", AutoCADCombinationValidationIsDynamic()
    AssertTrue stats, "ui.validation.plotLoadCase", PlotLoadCaseValidationIsDynamic()
    AssertTrue stats, "ui.validation.loadProfileId", LoadProfileValidationUsesDisplayNames()
    AssertTrue stats, "ui.validation.loadProfileId.noFormula", LoadProfileFirstValueHasNoFormula()
    AssertTrue stats, "ui.validation.plotLabels", _
        SystemSettingValidationHasOptions("Plot.ResultLabelsEnabled", Array("Yes", "No"))
    AssertTrue stats, "ui.validation.plotPrincipalAxes", _
        SystemSettingValidationHasOptions("Plot.PrincipalAxesMode", Array("Transformed", "Concrete", "None"))
    AssertTrue stats, "ui.validation.plotAxisLabels", _
        SystemSettingValidationHasOptions("Plot.AxisLabelsEnabled", Array("Yes", "No"))
    AssertClose stats, "ui.validation.plotAxisLabelsFontSize.default", Val(CStr(GetSystemSetting("Plot.AxisLabelsFontSize"))), 7.5, 0.000000001
    AssertTrue stats, "ui.validation.plotContour", _
        SystemSettingValidationHasOptions("Plot.ContourEnabled", Array("Yes", "No"))
    AssertTrue stats, "ui.validation.plotRebarAnnotationEnabled", _
        PlotAnnotationValidationHasOptions("Enabled", 2, Array("Yes", "No"))
    AssertTrue stats, "ui.validation.plotDimensionEnabled", _
        PlotAnnotationValidationHasOptions("Enabled", 3, Array("Yes", "No"))
    AssertTrue stats, "ui.validation.plotRebarPlacement", _
        PlotAnnotationValidationHasOptions("Placement", 2, Array("Outside", "Inside"))
    AssertTrue stats, "ui.validation.plotDimensionPlacement", _
        PlotAnnotationValidationHasOptions("Placement", 3, Array("Outside", "Inside"))
    AssertTrue stats, "ui.validation.plotRebarTextUnits", _
        PlotAnnotationValidationHasOptions("TextUnits", 2, Array("mm", "pt"))
    AssertTrue stats, "ui.validation.plotDimensionTextUnits", _
        PlotAnnotationValidationHasOptions("TextUnits", 3, Array("mm", "pt"))
    AssertTrue stats, "ui.validation.plotTextHeightUnitDynamic", _
        PlotAnnotationTextUnitCellIsDynamic("TextHeight")
    AssertTrue stats, "ui.validation.plotTextGapUnitDynamic", _
        PlotAnnotationTextUnitCellIsDynamic("TextGap")
    AssertClose stats, "ui.validation.plotTextHeight.defaultRebar", PlotAnnotationSettingValue("TextHeight", 2), 13#, 0.000000001
    AssertClose stats, "ui.validation.plotTextHeight.defaultDimension", PlotAnnotationSettingValue("TextHeight", 3), 13#, 0.000000001
    AssertClose stats, "ui.validation.plotTextGap.defaultRebar", PlotAnnotationSettingValue("TextGap", 2), 9#, 0.000000001
    AssertClose stats, "ui.validation.plotTextGap.defaultDimension", PlotAnnotationSettingValue("TextGap", 3), 9#, 0.000000001
    AssertTrue stats, "ui.validation.plotRebarLineEnabled", _
        PlotAnnotationValidationHasOptions("LineEnabled", 2, Array("Yes", "No"))
    AssertTrue stats, "ui.validation.plotDimensionArrowType", _
        PlotAnnotationValidationHasOptions("ArrowType", 3, Array("Triangle", "Stealth", "Diamond", "Oval", "Open"))
    AssertTrue stats, "ui.validation.plotDimensionArrowSize", _
        PlotAnnotationValidationHasOptions("ArrowSize", 3, Array("Small", "Medium", "Wide"))
    AssertTrue stats, "ui.validation.geometrySource", _
        SystemSettingValidationHasOptions("Geometry.Source", Array("Generated", "AutoCAD"))
    AssertTrue stats, "ui.validation.units.forceInput", _
        AnySettingValidationHasOptions("Force", Array("N", "kN", "tf"))
    AssertTrue stats, "ui.validation.units.momentOutput", _
        AnySettingValidationHasOptions("Moment", Array("N*mm", "kN*m", "tf*m"))
    AssertTrue stats, "ui.validation.sign.n", _
        AnySettingValidationHasOptions("+N", Array("Tension", "Compression"))
    SetSystemSetting "Geometry.Type", "Circle"
    AssertTrue stats, "ui.validation.circleLoc2row", _
        AnySettingValidationHasOptions("Rebar.Loc2row", Array("Stacked", "SideBySide"))
    AssertTrue stats, "ui.validation.circleLoc3row", _
        AnySettingValidationHasOptions("Rebar.Loc3row", Array("Stacked", "SideBySide"))
    SetSystemSetting "Geometry.Type", "RectSet"
    AssertTrue stats, "ui.validation.rectSetSectionType", _
        RectSetParameterValidationHasOptions("RectSet.SectionType", Array("Rectangle", "LSection", "TwoRectangles"))
    AssertTrue stats, "ui.validation.rectsetLoc2row", _
        RectSetExtraValidationHasOptions("H1 - левая", "положение", 1, Array("Stacked", "SideBySide"))
    AssertTrue stats, "ui.validation.rectsetBind2row", _
        RectSetExtraValidationHasOptions("H1 - левая", "привязка", 1, Array("EachBar", "EverySecondBar"))
    Dim reader As CSystemSettingsReader
    Set reader = New CSystemSettingsReader
    SetSystemSetting "Geometry.Type", "Circle"
    reader.LoadFromWorkbook ThisWorkbook
    AssertTrue stats, "ui.circle.key.rebarDiameter2", reader.HasKey("Rebar.Diameter2")
    AssertTrue stats, "ui.circle.key.rebarDiameter3", reader.HasKey("Rebar.Diameter3")
    AssertTrue stats, "ui.circle.key.rebarLoc2row", reader.HasKey("Rebar.Loc2row")
    AssertTrue stats, "ui.circle.key.rebarLoc3row", reader.HasKey("Rebar.Loc3row")
End Sub

' Подготавливает круговой fixture и проверяет реальную validation выбора LC для AutoCAD.
' Список должен содержать Worst/LC1 и ссылаться на формулы, а не быть статической строкой.
Private Function AutoCADCombinationValidationIsDynamic() As Boolean
    On Error GoTo Failed
    PrepareCircleInput

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
            AutoCADCombinationValidationIsDynamic = (CStr(listRange.Cells.Item(1, 1).Value2) = "Worst" And _
                CStr(listRange.Cells.Item(2, 1).Value2) = "LC1" And listRange.Cells.Item(2, 1).HasFormula)
            Exit Function
        End If
    Next rowIndex
Failed:
End Function

' Применяет ту же проверку динамического списка к селектору LC Excel-схемы.
Private Function PlotLoadCaseValidationIsDynamic() As Boolean
    PlotLoadCaseValidationIsDynamic = LoadCaseValidationIsDynamicForSetting("Plot.LoadCase")
End Function

' После подготовки fixture читает validation указанного Config-поля и зависимый диапазон.
' Ошибка/отсутствие диапазона означает непрохождение теста, а не создание новой validation.
Private Function LoadCaseValidationIsDynamicForSetting(ByVal settingKey As String) As Boolean
    On Error GoTo Failed
    PrepareCircleInput

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
            LoadCaseValidationIsDynamicForSetting = (CStr(listRange.Cells.Item(1, 1).Value2) = "Worst" And _
                CStr(listRange.Cells.Item(2, 1).Value2) = "LC1" And listRange.Cells.Item(2, 1).HasFormula)
            Exit Function
        End If
    Next rowIndex
Failed:
End Function

' Проверяет пару LC/Parameter в сохраненной таблице свойств Results по именам колонок.
' Расчет не запускается; отсутствующий снимок или колонка дает False.
Private Function ResultsPropertyExists(ByVal loadCase As String, ByVal parameter As String) As Boolean
    On Error GoTo Failed
    Dim data As Variant
    data = ResultTable("rngNDMSectionProperties")
    Dim colLC As Long: colLC = ResultHeaderColumn(data, "LoadCase")
    Dim colParam As Long: colParam = ResultHeaderColumn(data, "Parameter")
    Dim rowIndex As Long
    For rowIndex = 2 To UBound(data, 1)
        If StrComp(CStr(data(rowIndex, colLC)), loadCase, vbTextCompare) = 0 And _
                StrComp(CStr(data(rowIndex, colParam)), parameter, vbTextCompare) = 0 Then
            ResultsPropertyExists = True
            Exit Function
        End If
    Next rowIndex
Failed:
End Function

' Возвращает значение свойства из расчетного snapshot Results.
' Используется в UI-тестах, где важно проверить не только видимую Summary,
' но и машинные признаки выбранного LC, например ExtensionUsed.
Private Function ResultsPropertyValue(ByVal loadCase As String, ByVal parameter As String) As String
    On Error GoTo Failed
    Dim data As Variant
    data = ResultTable("rngNDMSectionProperties")
    Dim colLC As Long: colLC = ResultHeaderColumn(data, "LoadCase")
    Dim colParam As Long: colParam = ResultHeaderColumn(data, "Parameter")
    Dim colValue As Long: colValue = ResultHeaderColumn(data, "Value")
    Dim rowIndex As Long
    For rowIndex = 2 To UBound(data, 1)
        If StrComp(CStr(data(rowIndex, colLC)), loadCase, vbTextCompare) = 0 And _
                StrComp(CStr(data(rowIndex, colParam)), parameter, vbTextCompare) = 0 Then
            ResultsPropertyValue = CStr(data(rowIndex, colValue))
            Exit Function
        End If
    Next rowIndex
Failed:
End Function

' Возвращает длину из rngNDMSectionProperties во внутренних мм.
' Results хранит числовые значения в пользовательских OUTPUT-единицах последнего
' расчета, а Excel-схема строится во внутренних координатах модели.
Private Function ResultsLengthPropertyMm(ByVal loadCase As String, ByVal parameter As String) As Double
    ResultsLengthPropertyMm = CDbl(ResultsPropertyValue(loadCase, parameter)) * ResultsOutputLengthFactorToMm()
End Function

' Читает сохраненную в Results единицу вывода длины и возвращает множитель к мм.
' Тест намеренно опирается на snapshot Results, а не на текущий Config: пользователь
' может поменять настройки после расчета, но уже нарисованная схема должна
' соответствовать именно сохраненному расчетному снимку.
Private Function ResultsOutputLengthFactorToMm() As Double
    Select Case LCase$(Trim$(ResultsPropertyValue("ALL", "Output.LengthUnit")))
        Case "mm"
            ResultsOutputLengthFactorToMm = 1#
        Case "cm"
            ResultsOutputLengthFactorToMm = 10#
        Case "m"
            ResultsOutputLengthFactorToMm = 1000#
        Case Else
            ResultsOutputLengthFactorToMm = 1#
    End Select
End Function

' Возвращает числовое поле элемента из rngNDMSectionGeometry.
' Так UI-тест проверяет сохранность импортированных координат между
' geometry-preview и расчетным snapshot после очистки Results.
Private Function GeometryResultValue(ByRef data As Variant, ByVal elementID As String, _
        ByVal headerText As String) As Double
    Dim colID As Long
    Dim colValue As Long
    colID = ResultHeaderColumn(data, "ElementID")
    colValue = ResultHeaderColumn(data, headerText)
    If colValue = 0 Then colValue = ResultHeaderColumnByBaseName(data, headerText)
    If colID = 0 Or colValue = 0 Then Err.Raise vbObjectError + 4850, "GeometryResultValue", _
        "В rngNDMSectionGeometry не найдены нужные заголовки."

    Dim rowIndex As Long
    For rowIndex = 2 To UBound(data, 1)
        If StrComp(CStr(data(rowIndex, colID)), elementID, vbTextCompare) = 0 Then
            GeometryResultValue = CDbl(data(rowIndex, colValue))
            Exit Function
        End If
    Next rowIndex

    Err.Raise vbObjectError + 4851, "GeometryResultValue", _
        "В rngNDMSectionGeometry не найден элемент " & elementID & "."
End Function

' Ищет колонку по имени до запятой в заголовке вида "X, mm".
Private Function ResultHeaderColumnByBaseName(ByRef data As Variant, ByVal headerText As String) As Long
    Dim colIndex As Long
    For colIndex = 1 To UBound(data, 2)
        Dim actualHeader As String
        actualHeader = CStr(data(1, colIndex))
        If InStr(1, actualHeader, ",", vbTextCompare) > 0 Then _
            actualHeader = Trim$(Left$(actualHeader, InStr(1, actualHeader, ",", vbTextCompare) - 1))
        If StrComp(actualHeader, headerText, vbTextCompare) = 0 Then
            ResultHeaderColumnByBaseName = colIndex
            Exit Function
        End If
    Next colIndex
End Function

' Проверяет, что справочные характеристики приведенного сечения в Results
' считаются через обычный модульный коэффициент Es/Eb. Это важно для
' устойчивости и ручной проверки геометрии: вид диаграммы TwoLine/ThreeLine
' не должен менять Ared.
Private Sub AssertTransformedAreaUsesElasticModuli(ByRef stats As TUiTestStats)
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

    concreteArea = CDbl(ResultsPropertyValue("ALL", "Concrete.Area"))
    transformedArea = CDbl(ResultsPropertyValue("ALL", "Transformed.Area"))
    rebarArea = 8# * GEOM_PI * 20# * 20# / 4#
    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook
    Dim units As CUnitSystem
    Set units = New CUnitSystem
    units.LoadFromSettings settings
    expectedArea = concreteArea + units.InternalAreaToOutput((200000# / 32500# - 1#) * rebarArea)

    AssertClose stats, "ui.results.properties.transformedArea.moduli", _
        transformedArea, expectedArea, 0.001
    r1Plus = CDbl(ResultsPropertyValue("ALL", "Transformed.CoreDistance1Plus"))
    r1Minus = CDbl(ResultsPropertyValue("ALL", "Transformed.CoreDistance1Minus"))
    r2Plus = CDbl(ResultsPropertyValue("ALL", "Transformed.CoreDistance2Plus"))
    r2Minus = CDbl(ResultsPropertyValue("ALL", "Transformed.CoreDistance2Minus"))
    AssertTrue stats, "ui.results.properties.transformedCoreDistance.positive", _
        r1Plus > 0# And r1Minus > 0# And r2Plus > 0# And r2Minus > 0#
    AssertClose stats, "ui.results.properties.transformedCoreDistance.axis1Sym", _
        r1Plus, r1Minus, 0.001
    AssertClose stats, "ui.results.properties.transformedCoreDistance.axis2Sym", _
        r2Plus, r2Minus, 0.001
    xPlus = CDbl(ResultsPropertyValue("ALL", "Transformed.CoreDistanceXPlus"))
    xMinus = CDbl(ResultsPropertyValue("ALL", "Transformed.CoreDistanceXMinus"))
    yPlus = CDbl(ResultsPropertyValue("ALL", "Transformed.CoreDistanceYPlus"))
    yMinus = CDbl(ResultsPropertyValue("ALL", "Transformed.CoreDistanceYMinus"))
    AssertTrue stats, "ui.results.properties.transformedCoreDistanceXY.positive", _
        xPlus > 0# And xMinus > 0# And yPlus > 0# And yMinus > 0#
    AssertTrue stats, "ui.results.properties.concreteCoreDistanceXY.exists", _
        Len(ResultsPropertyValue("ALL", "Concrete.CoreDistanceXPlus")) > 0 And _
        Len(ResultsPropertyValue("ALL", "Concrete.CoreDistanceYMinus")) > 0
End Sub

' Проверяет наличие штатного ChartObject схемы на листе Расчет без его создания.
Private Function PlotChartExists() As Boolean
    On Error GoTo Failed
    Dim chartObject As Object
    Set chartObject = ThisWorkbook.Worksheets.Item("Расчет").ChartObjects("chtNDMSectionPlot")
    PlotChartExists = Not chartObject Is Nothing
Failed:
End Function

' Проверяет видимый Shape-заголовок существующей схемы.
Private Function PlotVisibleTitleContains(ByVal expectedText As String) As Boolean
    On Error GoTo Failed
    Dim chartObject As Object
    Set chartObject = ThisWorkbook.Worksheets.Item("Расчет").ChartObjects("chtNDMSectionPlot")

    Dim shapeIndex As Long
    For shapeIndex = 1 To chartObject.Chart.Shapes.Count
        If StrComp(chartObject.Chart.Shapes.Item(shapeIndex).Name, "NDMPlot_Title", vbTextCompare) = 0 Then
            PlotVisibleTitleContains = _
                (InStr(1, chartObject.Chart.Shapes.Item(shapeIndex).TextFrame.Characters().Text, expectedText, vbTextCompare) > 0)
            Exit Function
        End If
    Next shapeIndex
Failed:
End Function

' Считает подходящие фигуры внутри chart; при отсутствии схемы возвращает ноль.
Private Function CountPlotShapes(ByVal nameFragment As String) As Long
    On Error GoTo Failed
    Dim chartObject As Object
    Set chartObject = ThisWorkbook.Worksheets.Item("Расчет").ChartObjects("chtNDMSectionPlot")

    Dim shapeIndex As Long
    For shapeIndex = 1 To chartObject.Chart.Shapes.Count
        If InStr(1, chartObject.Chart.Shapes.Item(shapeIndex).Name, nameFragment, vbTextCompare) > 0 Then
            CountPlotShapes = CountPlotShapes + 1
        End If
    Next shapeIndex
Failed:
End Function

' Ищет тестируемую фигуру сначала в chart, затем на листе; отсутствие возвращает Nothing.
Private Function FirstGeneratedPlotShape(ByVal nameFragment As String) As Object
    On Error GoTo Failed
    Dim chartObject As Object
    Set chartObject = ThisWorkbook.Worksheets.Item("Расчет").ChartObjects("chtNDMSectionPlot")

    Dim shapeIndex As Long
    For shapeIndex = 1 To chartObject.Chart.Shapes.Count
        If InStr(1, chartObject.Chart.Shapes.Item(shapeIndex).Name, nameFragment, vbTextCompare) > 0 Then
            Set FirstGeneratedPlotShape = chartObject.Chart.Shapes.Item(shapeIndex)
            Exit Function
        End If
    Next shapeIndex

    Dim sheet As Object
    Set sheet = ThisWorkbook.Worksheets.Item("Расчет")
    For shapeIndex = 1 To sheet.Shapes.Count
        If InStr(1, sheet.Shapes.Item(shapeIndex).Name, nameFragment, vbTextCompare) > 0 Then
            Set FirstGeneratedPlotShape = sheet.Shapes.Item(shapeIndex)
            Exit Function
        End If
    Next shapeIndex
Failed:
End Function

' Центры фигур в единицах Excel points; для chart абсолютное смещение добавляется вызывающим кодом.
Private Function ShapeCenterX(ByVal shapeObject As Object) As Double
    ShapeCenterX = CDbl(shapeObject.Left) + CDbl(shapeObject.Width) / 2#
End Function

Private Function ShapeCenterY(ByVal shapeObject As Object) As Double
    ShapeCenterY = CDbl(shapeObject.Top) + CDbl(shapeObject.Height) / 2#
End Function

' Учитывает оба контейнера generated-фигур: chart и лист.
Private Function CountGeneratedPlotShapes(ByVal nameFragment As String) As Long
    CountGeneratedPlotShapes = CountPlotShapes(nameFragment) + CountWorksheetPlotShapes(nameFragment)
End Function

' Собирает средний абсолютный центр фигур двух контейнеров с учетом смещения chart.
' Вызывающий тест передает нулевые накопители; False означает отсутствие подходящих фигур.
Private Function GeneratedPlotShapeCenterAverage(ByVal nameFragment As String, ByRef centerX As Double, ByRef centerY As Double) As Boolean
    On Error GoTo Failed
    Dim chartObject As Object
    Set chartObject = ThisWorkbook.Worksheets.Item("Расчет").ChartObjects("chtNDMSectionPlot")

    Dim shapeIndex As Long
    Dim count As Long
    For shapeIndex = 1 To chartObject.Chart.Shapes.Count
        If InStr(1, chartObject.Chart.Shapes.Item(shapeIndex).Name, nameFragment, vbTextCompare) > 0 Then
            centerX = centerX + CDbl(chartObject.Left) + ShapeCenterX(chartObject.Chart.Shapes.Item(shapeIndex))
            centerY = centerY + CDbl(chartObject.Top) + ShapeCenterY(chartObject.Chart.Shapes.Item(shapeIndex))
            count = count + 1
        End If
    Next shapeIndex

    Dim sheet As Object
    Set sheet = ThisWorkbook.Worksheets.Item("Расчет")
    For shapeIndex = 1 To sheet.Shapes.Count
        If InStr(1, sheet.Shapes.Item(shapeIndex).Name, nameFragment, vbTextCompare) > 0 Then
            centerX = centerX + ShapeCenterX(sheet.Shapes.Item(shapeIndex))
            centerY = centerY + ShapeCenterY(sheet.Shapes.Item(shapeIndex))
            count = count + 1
        End If
    Next shapeIndex

    If count <= 0 Then Exit Function
    centerX = centerX / count
    centerY = centerY / count
    GeneratedPlotShapeCenterAverage = True
Failed:
End Function

' Проверяет границы каждой подходящей chart-фигуры относительно устойчивой рамки.
' Для PASS нужна хотя бы одна фигура; допуск задан тестом в Excel points.
Private Function PlotShapesInsideStableFrame(ByVal nameFragment As String, ByVal tolerance As Double) As Boolean
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
    PlotShapesInsideStableFrame = matched
Failed:
End Function

' Проверяет попадание абсолютной точки листа в границы хотя бы одной chart-фигуры.
' Смещение ChartObject учитывается отдельно от локальных координат фигуры.
Private Function PlotShapeBoundsContainAbsolutePoint(ByVal nameFragment As String, _
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
                PlotShapeBoundsContainAbsolutePoint = True
                Exit Function
            End If
        End If
    Next shapeIndex
Failed:
End Function

' ДЛЯ ТЕСТОВ: считает только Shapes схемы с префиксом NDMPlot_ на листе Расчет.
' Объекты Chart и пользовательские фигуры не включаются в проверку очистки.
Private Function CountWorksheetPlotShapes(ByVal nameFragment As String) As Long
    On Error GoTo Failed
    Dim sheet As Object
    Set sheet = ThisWorkbook.Worksheets.Item("Расчет")

    Dim shapeIndex As Long
    For shapeIndex = 1 To sheet.Shapes.Count
        If Left$(sheet.Shapes.Item(shapeIndex).Name, Len("NDMPlot_")) = "NDMPlot_" And _
                InStr(1, sheet.Shapes.Item(shapeIndex).Name, nameFragment, vbTextCompare) > 0 Then
            CountWorksheetPlotShapes = CountWorksheetPlotShapes + 1
        End If
    Next shapeIndex
Failed:
End Function

' Проверяет, что среди Shapes схемы есть объект с нужным углом поворота.
' Excel может хранить один и тот же угол как -30 или 330 градусов, поэтому
' сравниваем минимальную круговую разницу.
Private Function PlotShapeRotationExists(ByVal nameFragment As String, _
        ByVal expectedDegrees As Double, ByVal toleranceDegrees As Double) As Boolean
    On Error GoTo Failed
    Dim chartObject As Object
    Set chartObject = ThisWorkbook.Worksheets.Item("Расчет").ChartObjects("chtNDMSectionPlot")

    Dim shapeIndex As Long
    For shapeIndex = 1 To chartObject.Chart.Shapes.Count
        If InStr(1, chartObject.Chart.Shapes.Item(shapeIndex).Name, nameFragment, vbTextCompare) > 0 Then
            If AngleDistanceDegrees(CDbl(chartObject.Chart.Shapes.Item(shapeIndex).Rotation), expectedDegrees) <= toleranceDegrees Then
                PlotShapeRotationExists = True
                Exit Function
            End If
        End If
    Next shapeIndex
Failed:
End Function

' Возвращает минимальную разницу между углами в градусах с учетом периода 360.
Private Function AngleDistanceDegrees(ByVal actualDegrees As Double, ByVal expectedDegrees As Double) As Double
    Dim diff As Double
    diff = Abs(NormalizeDegrees(actualDegrees) - NormalizeDegrees(expectedDegrees))
    If diff > 180# Then diff = 360# - diff
    AngleDistanceDegrees = diff
End Function

' Нормализует угол к диапазону 0...360 для устойчивого сравнения Excel Shapes.
Private Function NormalizeDegrees(ByVal angleDegrees As Double) As Double
    Do While angleDegrees < 0#
        angleDegrees = angleDegrees + 360#
    Loop
    Do While angleDegrees >= 360#
        angleDegrees = angleDegrees - 360#
    Loop
    NormalizeDegrees = angleDegrees
End Function

' Ищет текст среди Shape-подписей текущей схемы.
' Тесты используют это для предупреждений, которые рисуются внутри ChartObject,
' а не выводятся отдельным окном Excel.
Private Function PlotShapeTextContains(ByVal expectedText As String) As Boolean
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
            PlotShapeTextContains = True
            Exit Function
        End If
    Next shapeIndex
Failed:
End Function

' Проверяет курсив у текстового Shape схемы по фрагменту имени.
Private Function PlotShapeTextItalic(ByVal nameFragment As String) As Boolean
    On Error GoTo Failed
    Dim chartObject As Object
    Set chartObject = ThisWorkbook.Worksheets.Item("Расчет").ChartObjects("chtNDMSectionPlot")

    Dim shapeIndex As Long
    For shapeIndex = 1 To chartObject.Chart.Shapes.Count
        If InStr(1, chartObject.Chart.Shapes.Item(shapeIndex).Name, nameFragment, vbTextCompare) > 0 Then
            PlotShapeTextItalic = (chartObject.Chart.Shapes.Item(shapeIndex).TextFrame.Characters().Font.Italic <> 0)
            Exit Function
        End If
    Next shapeIndex
Failed:
End Function

' Возвращает размер шрифта текстового Shape схемы по фрагменту имени.
Private Function PlotShapeTextFontSize(ByVal nameFragment As String) As Double
    On Error GoTo Failed
    Dim chartObject As Object
    Set chartObject = ThisWorkbook.Worksheets.Item("Расчет").ChartObjects("chtNDMSectionPlot")

    Dim shapeIndex As Long
    For shapeIndex = 1 To chartObject.Chart.Shapes.Count
        If InStr(1, chartObject.Chart.Shapes.Item(shapeIndex).Name, nameFragment, vbTextCompare) > 0 Then
            PlotShapeTextFontSize = CDbl(chartObject.Chart.Shapes.Item(shapeIndex).TextFrame.Characters().Font.Size)
            Exit Function
        End If
    Next shapeIndex
Failed:
End Function

' ДЛЯ ТЕСТОВ
' Возвращает число рядов данных на текущей схеме Excel.
Private Function PlotSeriesCount() As Long
    On Error GoTo Failed
    Dim chartObject As Object
    Set chartObject = ThisWorkbook.Worksheets.Item("Расчет").ChartObjects("chtNDMSectionPlot")
    PlotSeriesCount = chartObject.Chart.SeriesCollection.Count
    Exit Function

Failed:
    PlotSeriesCount = 0
End Function

' ДЛЯ ТЕСТОВ: считает строки сохраненной semantic-таблицы нужного типа.
' Заголовок пропускается; тест проверяет данные Results, не текущий генератор.
Private Function CountAnnotationType(ByRef annotationData As Variant, ByVal annotationType As String) As Long
    On Error GoTo Failed
    Dim rowIndex As Long
    For rowIndex = 2 To UBound(annotationData, 1)
        If StrComp(CStr(annotationData(rowIndex, 2)), annotationType, vbTextCompare) = 0 Then
            CountAnnotationType = CountAnnotationType + 1
        End If
    Next rowIndex
Failed:
End Function

' ДЛЯ ТЕСТОВ
' Возвращает видимую подпись semantic-аннотации из Results по ее ID.
Private Function AnnotationTextByID(ByRef annotationData As Variant, ByVal annotationID As String) As String
    On Error GoTo Failed
    Dim colID As Long
    Dim colText As Long
    colID = ResultHeaderColumn(annotationData, "AnnotationID")
    colText = ResultHeaderColumn(annotationData, "Text")

    Dim rowIndex As Long
    For rowIndex = 2 To UBound(annotationData, 1)
        If StrComp(CStr(annotationData(rowIndex, colID)), annotationID, vbTextCompare) = 0 Then
            AnnotationTextByID = CStr(annotationData(rowIndex, colText))
            Exit Function
        End If
    Next rowIndex

Failed:
End Function

' ДЛЯ ТЕСТОВ
' Возвращает численное значение semantic-аннотации из Results по ее ID.
Private Function AnnotationValueByID(ByRef annotationData As Variant, ByVal annotationID As String) As Double
    On Error GoTo Failed
    Dim colID As Long
    Dim colValue As Long
    colID = ResultHeaderColumn(annotationData, "AnnotationID")
    colValue = ResultHeaderColumn(annotationData, "Value")

    Dim rowIndex As Long
    For rowIndex = 2 To UBound(annotationData, 1)
        If StrComp(CStr(annotationData(rowIndex, colID)), annotationID, vbTextCompare) = 0 Then
            AnnotationValueByID = CDbl(annotationData(rowIndex, colValue))
            Exit Function
        End If
    Next rowIndex

Failed:
End Function

' Очищает все расчетные snapshot-блоки пользовательским macro, проверяя
' пустые результаты и сохранение исходного Geometry.Type на Config.
Private Sub TestClearResultsKeepsInputs(ByRef stats As TUiTestStats)
    PrepareCircleInput
    ThisWorkbook.Names.Item("rngNDMElementResults").RefersToRange.Value2 = "RunID"
    ThisWorkbook.Names.Item("rngNDMSectionGeometry").RefersToRange.Value2 = "RunID"
    ThisWorkbook.Names.Item("rngNDMSectionProperties").RefersToRange.Value2 = "RunID"
    ThisWorkbook.Names.Item("rngNDMSectionAnnotations").RefersToRange.Value2 = "RunID"
    ThisWorkbook.Names.Item("rngNDMMaterialDiagrams").RefersToRange.Value2 = "RunID"
    ClearSectionResultsForWorkbook ThisWorkbook
    AssertTrue stats, "ui.clear.results.elements", Len(CStr(ThisWorkbook.Names.Item("rngNDMElementResults").RefersToRange.Value2)) = 0
    AssertTrue stats, "ui.clear.results.geometry", Len(CStr(ThisWorkbook.Names.Item("rngNDMSectionGeometry").RefersToRange.Value2)) = 0
    AssertTrue stats, "ui.clear.results.properties", Len(CStr(ThisWorkbook.Names.Item("rngNDMSectionProperties").RefersToRange.Value2)) = 0
    AssertTrue stats, "ui.clear.results.annotations", Len(CStr(ThisWorkbook.Names.Item("rngNDMSectionAnnotations").RefersToRange.Value2)) = 0
    AssertTrue stats, "ui.clear.results.materialDiagrams", Len(CStr(ThisWorkbook.Names.Item("rngNDMMaterialDiagrams").RefersToRange.Value2)) = 0
    AssertTrue stats, "ui.clear.input", CStr(GetSystemSetting("Geometry.Type")) = "Circle"
End Sub

' Изменяет ячейку fixture по каноническому ключу с учетом специализированных таблиц.
' Неизвестный ключ вызывает ошибку теста, а не создает новую настройку.
Private Sub SetSystemSetting(ByVal key As String, ByVal value As String)
    If TrySetUnitOrSignSetting(key, value) Then Exit Sub

    Dim settings As Object
    Dim ranges As Variant
    ranges = SettingsRangeSearchOrder()

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

    If TrySetRectSetFaceSetting(key, value) Then Exit Sub
    If TrySetRoundedRectangleSetting(key, value) Then Exit Sub
    If TrySetHollowRectangleSetting(key, value) Then Exit Sub

    Err.Raise vbObjectError + 4210, "modTestWorkbookInterface", "System setting not found: " & key
End Sub

' Меняет одну ячейку расчетного профиля в тестовой книге.
' Диапазон профилей имеет две строки шапки, поэтому ProfileId ищется
' динамически, как это делает реальный reader.
Private Sub SetProfileSetting(ByVal profileId As String, ByVal key As String, ByVal value As String)
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

' Читает ячейку расчетного профиля по ProfileId и Key. Это нужно тестам,
' чтобы не зависеть от физического номера строки в rngCalculationProfiles.
Private Function ProfileSettingValue(ByVal profileId As String, ByVal key As String) As String
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
            ProfileSettingValue = CStr(profiles.Cells.Item(rowIndex, profileColumn).Value2)
            Exit Function
        End If
    Next rowIndex
End Function

' Направляет тестовый ввод единиц и знаков в их отдельные именованные таблицы.
' Возвращает False для другого типа ключа, чтобы продолжить поиск в Config.
Private Function TrySetUnitOrSignSetting(ByVal key As String, ByVal value As String) As Boolean
    Dim target As Object
    Dim rowIndex As Long
    Dim columnIndex As Long

    If UnitSettingAddress(key, rowIndex, columnIndex) Then
        Set target = ThisWorkbook.Names.Item("rngUnitSettings").RefersToRange
        target.Cells.Item(rowIndex, columnIndex).Value2 = value
        TrySetUnitOrSignSetting = True
        Exit Function
    End If

    If SignSettingAddress(key, rowIndex, columnIndex) Then
        Set target = ThisWorkbook.Names.Item("rngSignConventionSettings").RefersToRange
        target.Cells.Item(rowIndex, columnIndex).Value2 = value
        TrySetUnitOrSignSetting = True
    End If
End Function

' Разбирает ключ Units.Quantity.Side в строку физической величины и колонку стороны.
' Неизвестная величина или сторона не трактуется как допустимая настройка.
Private Function UnitSettingAddress(ByVal key As String, ByRef rowIndex As Long, ByRef columnIndex As Long) As Boolean
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
    UnitSettingAddress = True
End Function

' Сопоставляет ключ Sign.N/Mx/My.User/Internal с отдельной таблицей знаков.
' Индексы возвращаются только для поддержанного формата ключа.
Private Function SignSettingAddress(ByVal key As String, ByRef rowIndex As Long, ByRef columnIndex As Long) As Boolean
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
    SignSettingAddress = True
End Function

' Записывает распознанное поле RectSet в видимую таблицу геометрии и армирования.
' Отказ распознавания не меняет книгу и позволяет проверить другие формы.
Private Function TrySetRectSetFaceSetting(ByVal key As String, ByVal value As String) As Boolean
    Dim target As Object
    Set target = ThisWorkbook.Names.Item("rngRectSetGeometry").RefersToRange

    Dim rowIndex As Long
    Dim columnIndex As Long
    If Not RectSetSettingAddress(key, rowIndex, columnIndex) Then Exit Function
    target.Cells.Item(rowIndex, columnIndex).Value2 = value
    TrySetRectSetFaceSetting = True
End Function

' Сопоставляет размеры, форму, смещение и параметры рядов с ячейками RectSet.
' Положение граней определяется helpers ниже; это адреса fixture, не расчет координат.
Private Function RectSetSettingAddress(ByVal key As String, ByRef rowIndex As Long, ByRef columnIndex As Long) As Boolean
    If StrComp(key, "RectSet.SectionType", vbTextCompare) = 0 Then
        rowIndex = 3
        columnIndex = 2
        RectSetSettingAddress = True
        Exit Function
    ElseIf StrComp(key, "RectSet.UpperOffsetX", vbTextCompare) = 0 Then
        rowIndex = 4
        columnIndex = 2
        RectSetSettingAddress = True
        Exit Function
    End If

    Dim faceName As String
    If InStr(1, key, "RectSet.H1", vbTextCompare) = 1 Then
        faceName = "H1"
    ElseIf InStr(1, key, "RectSet.B1", vbTextCompare) = 1 Then
        faceName = "B1"
    ElseIf InStr(1, key, "RectSet.H2", vbTextCompare) = 1 Then
        faceName = "H2"
    ElseIf InStr(1, key, "RectSet.B2", vbTextCompare) = 1 Then
        faceName = "B2"
    Else
        Exit Function
    End If

    If StrComp(key, "RectSet." & faceName, vbTextCompare) = 0 Then
        rowIndex = 8
        columnIndex = RectSetGeometryColumn(faceName)
    ElseIf StrComp(key, "RectSet." & faceName & ".as_1", vbTextCompare) = 0 Then
        rowIndex = RectSetMainRow(faceName, 1): columnIndex = 2
    ElseIf StrComp(key, "RectSet." & faceName & ".as_2", vbTextCompare) = 0 Then
        rowIndex = RectSetMainRow(faceName, 2): columnIndex = 2
    ElseIf StrComp(key, "RectSet." & faceName & ".d_1", vbTextCompare) = 0 Then
        rowIndex = RectSetMainRow(faceName, 1): columnIndex = 3
    ElseIf StrComp(key, "RectSet." & faceName & ".d_2", vbTextCompare) = 0 Then
        rowIndex = RectSetMainRow(faceName, 2): columnIndex = 3
    ElseIf StrComp(key, "RectSet." & faceName & ".n_1", vbTextCompare) = 0 Then
        rowIndex = RectSetMainRow(faceName, 1): columnIndex = 4
    ElseIf StrComp(key, "RectSet." & faceName & ".n_2", vbTextCompare) = 0 Then
        rowIndex = RectSetMainRow(faceName, 2): columnIndex = 4
    ElseIf StrComp(key, "RectSet." & faceName & ".StartOffset1", vbTextCompare) = 0 Then
        rowIndex = RectSetMainRow(faceName, 1): columnIndex = 5
    ElseIf StrComp(key, "RectSet." & faceName & ".EndOffset1", vbTextCompare) = 0 Then
        rowIndex = RectSetMainRow(faceName, 1): columnIndex = 6
    ElseIf StrComp(key, "RectSet." & faceName & ".StartOffset2", vbTextCompare) = 0 Then
        rowIndex = RectSetMainRow(faceName, 2): columnIndex = 5
    ElseIf StrComp(key, "RectSet." & faceName & ".EndOffset2", vbTextCompare) = 0 Then
        rowIndex = RectSetMainRow(faceName, 2): columnIndex = 6
    ElseIf StrComp(key, "RectSet." & faceName & ".d_2row_1", vbTextCompare) = 0 Then
        rowIndex = RectSetExtraRow(faceName, 1): columnIndex = 2
    ElseIf StrComp(key, "RectSet." & faceName & ".d_2row_2", vbTextCompare) = 0 Then
        rowIndex = RectSetExtraRow(faceName, 2): columnIndex = 2
    ElseIf StrComp(key, "RectSet." & faceName & ".d_3row_1", vbTextCompare) = 0 Then
        rowIndex = RectSetExtraRow(faceName, 1): columnIndex = 5
    ElseIf StrComp(key, "RectSet." & faceName & ".d_3row_2", vbTextCompare) = 0 Then
        rowIndex = RectSetExtraRow(faceName, 2): columnIndex = 5
    ElseIf StrComp(Left$(key, Len(key) - 1), "RectSet." & faceName & ".loc_2row_", vbTextCompare) = 0 Then
        rowIndex = RectSetExtraRow(faceName, CLng(Right$(key, 1))): columnIndex = 3
    ElseIf StrComp(Left$(key, Len(key) - 1), "RectSet." & faceName & ".loc_3row_", vbTextCompare) = 0 Then
        rowIndex = RectSetExtraRow(faceName, CLng(Right$(key, 1))): columnIndex = 6
    ElseIf StrComp(Left$(key, Len(key) - 1), "RectSet." & faceName & ".bind_2row_", vbTextCompare) = 0 Then
        rowIndex = RectSetExtraRow(faceName, CLng(Right$(key, 1))): columnIndex = 4
    ElseIf StrComp(Left$(key, Len(key) - 1), "RectSet." & faceName & ".bind_3row_", vbTextCompare) = 0 Then
        rowIndex = RectSetExtraRow(faceName, CLng(Right$(key, 1))): columnIndex = 7
    Else
        Exit Function
    End If
    RectSetSettingAddress = True
End Function

' Группа адресных helpers сохраняет порядок граней H1/B1/H2/B2 и двух сторон.
' Основные ряды и дополнительные ряды находятся в разных блоках таблицы.
Private Function RectSetGeometryColumn(ByVal faceName As String) As Long
    Select Case UCase$(faceName)
        Case "H1": RectSetGeometryColumn = 1
        Case "B1": RectSetGeometryColumn = 2
        Case "H2": RectSetGeometryColumn = 3
        Case "B2": RectSetGeometryColumn = 4
    End Select
End Function

Private Function RectSetMainRow(ByVal faceName As String, ByVal sideIndex As Long) As Long
    RectSetMainRow = 10 + RectSetFaceOrdinal(faceName, sideIndex)
End Function

Private Function RectSetExtraRow(ByVal faceName As String, ByVal sideIndex As Long) As Long
    RectSetExtraRow = 20 + RectSetFaceOrdinal(faceName, sideIndex)
End Function

' Переводит пару грань/сторона в порядковый номер строки fixture-таблицы:
' H1, B1, H2, B2 занимают по две строки, стороны сохраняют порядок 1/2.
Private Function RectSetFaceOrdinal(ByVal faceName As String, ByVal sideIndex As Long) As Long
    Select Case UCase$(faceName)
        Case "H1": RectSetFaceOrdinal = sideIndex
        Case "B1": RectSetFaceOrdinal = 2 + sideIndex
        Case "H2": RectSetFaceOrdinal = 4 + sideIndex
        Case "B2": RectSetFaceOrdinal = 6 + sideIndex
    End Select
End Function

' Меняет одну ячейку табличного блока RoundedRectangle в тестовой книге.
' Адресация учитывает видимые блоки размеров, граней и дополнительных рядов.
Private Function TrySetRoundedRectangleSetting(ByVal key As String, ByVal value As String) As Boolean
    Dim target As Object
    Set target = ThisWorkbook.Names.Item("rngRoundedRectangleGeometry").RefersToRange

    Dim rowIndex As Long
    Dim columnIndex As Long
    If Not RoundedRectangleSettingAddress(key, rowIndex, columnIndex) Then Exit Function
    target.Cells.Item(rowIndex, columnIndex).Value2 = value
    TrySetRoundedRectangleSetting = True
End Function

' Возвращает адрес пользовательского параметра внутри табличного диапазона
' RoundedRectangle. Номера строк привязаны к layout-у SettingsCatalog, а не к
' старым внутренним ключам, которых пользователь в этой таблице не видит.
Private Function RoundedRectangleSettingAddress(ByVal key As String, _
        ByRef rowIndex As Long, ByRef columnIndex As Long) As Boolean
    Select Case LCase$(key)
        Case "roundedrectangle.b"
            rowIndex = 3: columnIndex = 2
        Case "roundedrectangle.h"
            rowIndex = 3: columnIndex = 3
        Case "roundedrectangle.left.type"
            rowIndex = 7: columnIndex = 2
        Case "roundedrectangle.right.type"
            rowIndex = 7: columnIndex = 3
        Case "roundedrectangle.left.w"
            rowIndex = 8: columnIndex = 2
        Case "roundedrectangle.right.w"
            rowIndex = 8: columnIndex = 3
        Case "roundedrectangle.left.r1"
            rowIndex = 9: columnIndex = 2
        Case "roundedrectangle.right.r1"
            rowIndex = 9: columnIndex = 3
        Case "roundedrectangle.left.r2"
            rowIndex = 10: columnIndex = 2
        Case "roundedrectangle.right.r2"
            rowIndex = 10: columnIndex = 3
        Case Else
            If Not RoundedRectangleFaceSettingAddress(key, rowIndex, columnIndex) Then Exit Function
    End Select
    RoundedRectangleSettingAddress = True
End Function

' Возвращает адрес параметра арматуры по одной из четырех видимых граней:
' H.Left, H.Right, B.Top или B.Bottom.
Private Function RoundedRectangleFaceSettingAddress(ByVal key As String, _
        ByRef rowIndex As Long, ByRef columnIndex As Long) As Boolean
    Dim prefix As String
    Dim faceOrdinal As Long
    prefix = RoundedRectangleFacePrefix(key, faceOrdinal)
    If Len(prefix) = 0 Then Exit Function

    Dim tail As String
    tail = Mid$(key, Len(prefix) + 2)
    Select Case LCase$(tail)
        Case "as"
            rowIndex = 12 + faceOrdinal: columnIndex = 2
        Case "d"
            rowIndex = 12 + faceOrdinal: columnIndex = 3
        Case "n"
            rowIndex = 12 + faceOrdinal: columnIndex = 4
        Case "d_2row"
            rowIndex = 18 + faceOrdinal: columnIndex = 2
        Case "loc_2row"
            rowIndex = 18 + faceOrdinal: columnIndex = 3
        Case "bind_2row"
            rowIndex = 18 + faceOrdinal: columnIndex = 4
        Case "d_3row"
            rowIndex = 18 + faceOrdinal: columnIndex = 5
        Case "loc_3row"
            rowIndex = 18 + faceOrdinal: columnIndex = 6
        Case "bind_3row"
            rowIndex = 18 + faceOrdinal: columnIndex = 7
        Case Else
            Exit Function
    End Select
    RoundedRectangleFaceSettingAddress = True
End Function

' Нормализует имя грани RoundedRectangle и возвращает ее порядковый номер в
' блоках основного и дополнительного армирования.
Private Function RoundedRectangleFacePrefix(ByVal key As String, ByRef faceOrdinal As Long) As String
    If InStr(1, key, "RoundedRectangle.H.Left.", vbTextCompare) = 1 Then
        RoundedRectangleFacePrefix = "RoundedRectangle.H.Left"
        faceOrdinal = 1
    ElseIf InStr(1, key, "RoundedRectangle.H.Right.", vbTextCompare) = 1 Then
        RoundedRectangleFacePrefix = "RoundedRectangle.H.Right"
        faceOrdinal = 2
    ElseIf InStr(1, key, "RoundedRectangle.B.Top.", vbTextCompare) = 1 Then
        RoundedRectangleFacePrefix = "RoundedRectangle.B.Top"
        faceOrdinal = 3
    ElseIf InStr(1, key, "RoundedRectangle.B.Bottom.", vbTextCompare) = 1 Then
        RoundedRectangleFacePrefix = "RoundedRectangle.B.Bottom"
        faceOrdinal = 4
    End If
End Function

' Меняет одну ячейку таблицы HollowRectangle. Таблица содержит наружный
' контур, Opening и восемь смысловых граней, поэтому адреса задаются не
' обычным Key/Value-списком, а видимой структурой блока Config.
Private Function TrySetHollowRectangleSetting(ByVal key As String, ByVal value As String) As Boolean
    Dim target As Object
    Set target = ThisWorkbook.Names.Item("rngHollowRectangleGeometry").RefersToRange

    Dim rowIndex As Long
    Dim columnIndex As Long
    If Not HollowRectangleSettingAddress(key, rowIndex, columnIndex) Then Exit Function
    target.Cells.Item(rowIndex, columnIndex).Value2 = value
    TrySetHollowRectangleSetting = True
End Function

' Определяет ячейку размеров/смещения отверстия или передает ключ адресатору грани.
' Распознавание не меняет значение, позволяя отдельно проверять чтение и запись.
Private Function HollowRectangleSettingAddress(ByVal key As String, _
        ByRef rowIndex As Long, ByRef columnIndex As Long) As Boolean
    Select Case LCase$(key)
        Case "hollowrectangle.inneroffsetx"
            rowIndex = 3: columnIndex = 2
        Case "hollowrectangle.inneroffsety"
            rowIndex = 4: columnIndex = 2
        Case "hollowrectangle.h"
            rowIndex = 7: columnIndex = 1
        Case "hollowrectangle.b"
            rowIndex = 7: columnIndex = 2
        Case "hollowrectangle.r"
            rowIndex = 7: columnIndex = 3
        Case "hollowrectangle.openingh"
            rowIndex = 7: columnIndex = 4
        Case "hollowrectangle.openingb"
            rowIndex = 7: columnIndex = 5
        Case "hollowrectangle.openingr"
            rowIndex = 7: columnIndex = 6
        Case Else
            If Not HollowRectangleFaceSettingAddress(key, rowIndex, columnIndex) Then Exit Function
    End Select
    HollowRectangleSettingAddress = True
End Function

' Сопоставляет арматурный параметр с наружной или внутренней гранью HollowRectangle.
' Основные и дополнительные ряды адресуются отдельно; неверный параметр отклоняется.
Private Function HollowRectangleFaceSettingAddress(ByVal key As String, _
        ByRef rowIndex As Long, ByRef columnIndex As Long) As Boolean
    Dim prefix As String
    Dim faceOrdinal As Long
    prefix = HollowRectangleFacePrefix(key, faceOrdinal)
    If Len(prefix) = 0 Then Exit Function

    Dim tail As String
    tail = Mid$(key, Len(prefix) + 2)
    Select Case LCase$(tail)
        Case "as"
            rowIndex = 10 + faceOrdinal: columnIndex = 2
        Case "d"
            rowIndex = 10 + faceOrdinal: columnIndex = 3
        Case "n"
            rowIndex = 10 + faceOrdinal: columnIndex = 4
        Case "d_2row"
            rowIndex = 20 + faceOrdinal: columnIndex = 2
        Case "loc_2row"
            rowIndex = 20 + faceOrdinal: columnIndex = 3
        Case "bind_2row"
            rowIndex = 20 + faceOrdinal: columnIndex = 4
        Case "d_3row"
            rowIndex = 20 + faceOrdinal: columnIndex = 5
        Case "loc_3row"
            rowIndex = 20 + faceOrdinal: columnIndex = 6
        Case "bind_3row"
            rowIndex = 20 + faceOrdinal: columnIndex = 7
        Case Else
            Exit Function
    End Select
    HollowRectangleFaceSettingAddress = True
End Function

' Возвращает канонический префикс и номер одной из восьми граней сечения.
' Четыре грани отверстия не смешиваются с наружными сторонами.
Private Function HollowRectangleFacePrefix(ByVal key As String, ByRef faceOrdinal As Long) As String
    If InStr(1, key, "HollowRectangle.H.Left.", vbTextCompare) = 1 Then
        HollowRectangleFacePrefix = "HollowRectangle.H.Left": faceOrdinal = 1
    ElseIf InStr(1, key, "HollowRectangle.H.Right.", vbTextCompare) = 1 Then
        HollowRectangleFacePrefix = "HollowRectangle.H.Right": faceOrdinal = 2
    ElseIf InStr(1, key, "HollowRectangle.B.Top.", vbTextCompare) = 1 Then
        HollowRectangleFacePrefix = "HollowRectangle.B.Top": faceOrdinal = 3
    ElseIf InStr(1, key, "HollowRectangle.B.Bottom.", vbTextCompare) = 1 Then
        HollowRectangleFacePrefix = "HollowRectangle.B.Bottom": faceOrdinal = 4
    ElseIf InStr(1, key, "HollowRectangle.Opening.H.Left.", vbTextCompare) = 1 Then
        HollowRectangleFacePrefix = "HollowRectangle.Opening.H.Left": faceOrdinal = 5
    ElseIf InStr(1, key, "HollowRectangle.Opening.H.Right.", vbTextCompare) = 1 Then
        HollowRectangleFacePrefix = "HollowRectangle.Opening.H.Right": faceOrdinal = 6
    ElseIf InStr(1, key, "HollowRectangle.Opening.B.Top.", vbTextCompare) = 1 Then
        HollowRectangleFacePrefix = "HollowRectangle.Opening.B.Top": faceOrdinal = 7
    ElseIf InStr(1, key, "HollowRectangle.Opening.B.Bottom.", vbTextCompare) = 1 Then
        HollowRectangleFacePrefix = "HollowRectangle.Opening.B.Bottom": faceOrdinal = 8
    End If
End Function

' Читает значение fixture из специальных таблиц или упорядоченных диапазонов Config.
' Отсутствие ключа вызывает ошибку теста, вместо подмены прочитанного значения default.
Private Function GetSystemSetting(ByVal key As String) As String
    Dim directValue As String
    If TryGetUnitOrSignSetting(key, directValue) Then
        GetSystemSetting = directValue
        Exit Function
    End If

    Dim settings As Object
    Dim ranges As Variant
    ranges = SettingsRangeSearchOrder()

    Dim rangeIndex As Long
    For rangeIndex = LBound(ranges) To UBound(ranges)
        Set settings = ThisWorkbook.Names.Item(CStr(ranges(rangeIndex))).RefersToRange

        Dim rowIndex As Long
        For rowIndex = 2 To settings.Rows.Count
            If StrComp(CStr(settings.Cells.Item(rowIndex, 1).Value2), key, vbTextCompare) = 0 Then
                GetSystemSetting = CStr(settings.Cells.Item(rowIndex, 2).Value2)
                Exit Function
            End If
        Next rowIndex
    Next rangeIndex

    Err.Raise vbObjectError + 4211, "modTestWorkbookInterface", "System setting not found: " & key
End Function

' Проверяет наличие имени книги без падения теста на ошибке Excel.
' Используется для контроля, что старые входные диапазоны не вернулись после
' перестройки листа Config.
Private Function WorkbookNameExists(ByVal nameText As String) As Boolean
    On Error Resume Next
    Dim nm As Object
    Set nm = ThisWorkbook.Names.Item(nameText)
    WorkbookNameExists = (Err.Number = 0 And Not nm Is Nothing)
    Err.Clear
    On Error GoTo 0
End Function

' Читает настройки из специальных таблиц единиц и знаков. Эти диапазоны не
' имеют обычного формата Key/Value, поэтому используют ту же адресацию, что и
' SetSystemSetting при записи тестового значения.
Private Function TryGetUnitOrSignSetting(ByVal key As String, ByRef value As String) As Boolean
    Dim target As Object
    Dim rowIndex As Long
    Dim columnIndex As Long

    If UnitSettingAddress(key, rowIndex, columnIndex) Then
        Set target = ThisWorkbook.Names.Item("rngUnitSettings").RefersToRange
        value = CStr(target.Cells.Item(rowIndex, columnIndex).Value2)
        TryGetUnitOrSignSetting = True
        Exit Function
    End If

    If SignSettingAddress(key, rowIndex, columnIndex) Then
        Set target = ThisWorkbook.Names.Item("rngSignConventionSettings").RefersToRange
        value = CStr(target.Cells.Item(rowIndex, columnIndex).Value2)
        TryGetUnitOrSignSetting = True
    End If
End Function

' Ставит выбранную геометрию раньше неактивных таблиц с одинаковыми именами параметров.
' Остальные системные, материальные и презентационные блоки остаются доступны.
Private Function SettingsRangeSearchOrder() As Variant
    Dim geometryType As String
    geometryType = SystemGeometryType()
    If StrComp(geometryType, "RectSet", vbTextCompare) = 0 Then
        SettingsRangeSearchOrder = Array("rngUnitSettings", "rngSignConventionSettings", "rngSystemSettings", "rngConcreteMaterialParameters", "rngSteelMaterialParameters", "rngCalculationProfiles", "rngPlotAnnotationSettings", "rngRectSetGeometry", "rngCircleGeometry", "rngRoundedRectangleGeometry", "rngHollowRectangleGeometry")
    ElseIf StrComp(geometryType, "RoundedRectangle", vbTextCompare) = 0 Then
        SettingsRangeSearchOrder = Array("rngUnitSettings", "rngSignConventionSettings", "rngSystemSettings", "rngConcreteMaterialParameters", "rngSteelMaterialParameters", "rngCalculationProfiles", "rngPlotAnnotationSettings", "rngRoundedRectangleGeometry", "rngCircleGeometry", "rngRectSetGeometry", "rngHollowRectangleGeometry")
    ElseIf StrComp(geometryType, "HollowRectangle", vbTextCompare) = 0 Then
        SettingsRangeSearchOrder = Array("rngUnitSettings", "rngSignConventionSettings", "rngSystemSettings", "rngConcreteMaterialParameters", "rngSteelMaterialParameters", "rngCalculationProfiles", "rngPlotAnnotationSettings", "rngHollowRectangleGeometry", "rngRoundedRectangleGeometry", "rngCircleGeometry", "rngRectSetGeometry")
    Else
        SettingsRangeSearchOrder = Array("rngUnitSettings", "rngSignConventionSettings", "rngSystemSettings", "rngConcreteMaterialParameters", "rngSteelMaterialParameters", "rngCalculationProfiles", "rngPlotAnnotationSettings", "rngCircleGeometry", "rngRoundedRectangleGeometry", "rngHollowRectangleGeometry", "rngRectSetGeometry")
    End If
End Function

' Читает тип геометрии для адресации тестовых ячеек, не строя модель сечения.
' Circle используется только как fallback инфраструктуры при отсутствии fixture.
Private Function SystemGeometryType() As String
    On Error GoTo Failed
    Dim settings As Object
    Set settings = ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange

    Dim rowIndex As Long
    For rowIndex = 2 To settings.Rows.Count
        If StrComp(CStr(settings.Cells.Item(rowIndex, 1).Value2), "Geometry.Type", vbTextCompare) = 0 Then
            SystemGeometryType = CStr(settings.Cells.Item(rowIndex, 2).Value2)
            Exit Function
        End If
    Next rowIndex
Failed:
    SystemGeometryType = "Circle"
End Function

' Группа wrappers проверяет validation в выбранном блоке, а не только текст ячейки.
' Реальные списки допустимых вариантов сверяются адресными helpers ниже.
Private Function SystemSettingValidationHasOptions(ByVal key As String, ByVal expectedOptions As Variant) As Boolean
    SystemSettingValidationHasOptions = SettingValidationHasOptionsInRange(ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange, key, expectedOptions)
End Function

' ДЛЯ ТЕСТОВ: проверяет реальный общий dropdown сочетания, включая Auto.
' Пустой ввод разрешен; удаленная отдельная настройка Formation не используется.
Private Function UniversalLoadPathValidationHasOptions() As Boolean
    On Error GoTo Failed
    Dim loads As Object, cell As Object, options As Object, item As Variant
    Set loads = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange
    If CStr(loads.Cells(1, 6).Value2) <> "LoadPath" Then Exit Function
    Set cell = loads.Cells(2, 6)
    If cell.Validation.Type <> xlValidateList Or Not cell.Validation.IgnoreBlank Then Exit Function
    Set options = loads.Worksheet.Range(Mid$(CStr(cell.Validation.Formula1), 2))
    Dim expected As Variant, index As Long
    expected = Array("Auto", ChrW$(&H3BB) & "*Mx", ChrW$(&H3BB) & "*My", _
        ChrW$(&H3BB) & "*Mxy", ChrW$(&H3BB) & "*N", ChrW$(&H3BB) & "*NMxy")
    If options.Cells.Count <> UBound(expected) + 1 Then Exit Function
    For index = 0 To UBound(expected)
        If CStr(options.Cells(index + 1, 1).Value2) <> CStr(expected(index)) Then Exit Function
    Next index
    UniversalLoadPathValidationHasOptions = True
Failed:
End Function

' Ищет настройку в rngSystemSettings и проверяет, что ее колонка "Ед."
' является формулой к нужной строке rngUnitSettings. Видимое значение
' может быть разным после выбора пользователем единиц, поэтому тестируем
' источник данных формулы, а не конкретный текст вроде tf или kN*m.
Private Function SystemSettingUnitCellReferencesQuantity(ByVal key As String, ByVal quantity As String) As Boolean
    On Error GoTo Failed

    Dim settings As Object
    Set settings = ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange

    Dim rowIndex As Long
    For rowIndex = 2 To settings.Rows.Count
        If StrComp(CStr(settings.Cells.Item(rowIndex, 1).Value2), key, vbTextCompare) = 0 Then
            Dim formulaText As String
            formulaText = CStr(settings.Cells.Item(rowIndex, 3).Formula)
            If Left$(formulaText, 1) <> "=" Then Exit Function

            SystemSettingUnitCellReferencesQuantity = _
                InStr(1, formulaText, "rngUnitSettings", vbTextCompare) > 0 And _
                InStr(1, formulaText, """" & quantity & """", vbTextCompare) > 0
            Exit Function
        End If
    Next rowIndex

Failed:
End Function

' Читает сохраненный output-блок от именованного anchor до конца его данных.
' Не вызывает solver и не использует текущие геометрические исходные данные.
Private Function ResultTable(ByVal rangeName As String) As Variant
    Dim anchor As Object
    Set anchor = ThisWorkbook.Names.Item(rangeName).RefersToRange

    Dim columnCount As Long
    columnCount = ResultTableColumnCount(anchor)
    If columnCount <= 0 Then Exit Function

    Dim rowCount As Long
    rowCount = ResultTableRowCount(rangeName)
    If rowCount <= 0 Then rowCount = 1

    ResultTable = anchor.Resize(rowCount, columnCount).Value2
End Function

' Определяет высоту непрерывного сохраненного блока по первой ключевой колонке.
' Пустая строка завершает именно табличный snapshot, не таблицу входных LC.
Private Function ResultTableRowCount(ByVal rangeName As String) As Long
    Dim anchor As Object
    Set anchor = ThisWorkbook.Names.Item(rangeName).RefersToRange

    Dim rowOffset As Long
    For rowOffset = 0 To 1048575 - anchor.Row
        If Len(Trim$(CStr(anchor.Offset(rowOffset, 0).Value2))) = 0 Then Exit For
        ResultTableRowCount = ResultTableRowCount + 1
    Next rowOffset
End Function

' Находит число колонок output-блока с учетом следующего соседнего anchor.
' Helpers ниже ограничивают чтение своим блоком Results без захвата чужих данных.
Private Function ResultTableColumnCount(ByVal anchor As Object) As Long
    Dim maxColumns As Long
    maxColumns = ColumnsUntilNextResultAnchor(anchor)
    If maxColumns <= 0 Then maxColumns = 256

    Dim colOffset As Long
    For colOffset = 0 To maxColumns - 1
        If Len(Trim$(CStr(anchor.Offset(0, colOffset).Value2))) = 0 Then Exit For
        ResultTableColumnCount = ResultTableColumnCount + 1
    Next colOffset
End Function

' Проверяет, что единица настройки собрана формулой сразу из двух строк
' rngUnitSettings. Это нужно для величин вида момент/длина, где видимый текст
' меняется вместе с выбранной пользователем INPUT-системой единиц.
Private Function SystemSettingUnitCellReferencesQuantities(ByVal key As String, _
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

            SystemSettingUnitCellReferencesQuantities = _
                InStr(1, formulaText, "rngUnitSettings", vbTextCompare) > 0 And _
                InStr(1, formulaText, """" & quantity1 & """", vbTextCompare) > 0 And _
                InStr(1, formulaText, """" & quantity2 & """", vbTextCompare) > 0
            Exit Function
        End If
    Next rowIndex

Failed:
End Function

' Ограничивает чтение таблицы Results ближайшим соседним именованным
' диапазоном справа. Блоки Results стоят на одной строке и не обязаны иметь
' пустой столбец между собой, поэтому простого CurrentRegion здесь мало.
Private Function ColumnsUntilNextResultAnchor(ByVal anchor As Object) As Long
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

    ColumnsUntilNextResultAnchor = bestDelta
End Function

' Ищет validation ключа в реально существующих таблицах выбранной конфигурации.
' Успех требует совпадения полного списка вариантов, а не только наличия dropdown.
Private Function AnySettingValidationHasOptions(ByVal key As String, ByVal expectedOptions As Variant) As Boolean
    Dim ranges As Variant
    ranges = SettingsRangeSearchOrder()

    Dim rangeIndex As Long
    For rangeIndex = LBound(ranges) To UBound(ranges)
        If SettingValidationHasOptionsInRange(ThisWorkbook.Names.Item(CStr(ranges(rangeIndex))).RefersToRange, key, expectedOptions) Then
            AnySettingValidationHasOptions = True
            Exit Function
        End If
    Next rangeIndex
End Function

' Сверяет formula-linked validation с ожидаемым списком, включая его порядок и длину.
' Ошибка чтения списка означает неуспешную проверку интерфейса.
Private Function SettingValidationHasOptionsInRange(ByVal settings As Object, ByVal key As String, ByVal expectedOptions As Variant) As Boolean
    On Error GoTo Failed

    Dim rowIndex As Long
    For rowIndex = 2 To settings.Rows.Count
        If StrComp(CStr(settings.Cells.Item(rowIndex, 1).Value2), key, vbTextCompare) = 0 Then
            Dim formulaText As String
            formulaText = CStr(settings.Cells.Item(rowIndex, 2).Validation.Formula1)
            If Left$(formulaText, 1) <> "=" Then Exit Function

            Dim listRange As Object
            Set listRange = ThisWorkbook.Worksheets.Item("Config").Range(Mid$(formulaText, 2))
            Dim i As Long
            If listRange.Cells.Count <> (UBound(expectedOptions) - LBound(expectedOptions) + 1) Then Exit Function
            For i = LBound(expectedOptions) To UBound(expectedOptions)
                If CStr(listRange.Cells.Item(i - LBound(expectedOptions) + 1, 1).Value2) <> CStr(expectedOptions(i)) Then Exit Function
            Next i
            SettingValidationHasOptionsInRange = True
            Exit Function
        End If
    Next rowIndex
Failed:
End Function

' Проверяет список вариантов выбранной группы аннотаций в ее собственной колонке.
' Это исключает случайный успех за счет настройки соседней группы.
Private Function PlotAnnotationValidationHasOptions(ByVal rowName As String, ByVal valueColumn As Long, ByVal expectedOptions As Variant) As Boolean
    On Error GoTo Failed

    Dim settings As Object
    Set settings = ThisWorkbook.Names.Item("rngPlotAnnotationSettings").RefersToRange

    Dim rowIndex As Long
    For rowIndex = 2 To settings.Rows.Count
        If StrComp(CStr(settings.Cells.Item(rowIndex, 1).Value2), rowName, vbTextCompare) = 0 Then
            Dim formulaText As String
            formulaText = CStr(settings.Cells.Item(rowIndex, valueColumn).Validation.Formula1)
            If Left$(formulaText, 1) <> "=" Then Exit Function

            Dim listRange As Object
            Set listRange = ThisWorkbook.Worksheets.Item("Config").Range(Mid$(formulaText, 2))
            Dim i As Long
            If listRange.Cells.Count <> (UBound(expectedOptions) - LBound(expectedOptions) + 1) Then Exit Function
            For i = LBound(expectedOptions) To UBound(expectedOptions)
                If CStr(listRange.Cells.Item(i - LBound(expectedOptions) + 1, 1).Value2) <> CStr(expectedOptions(i)) Then Exit Function
            Next i
            PlotAnnotationValidationHasOptions = True
            Exit Function
        End If
    Next rowIndex
Failed:
End Function

' Находит строку параметра RectSet и проверяет validation ее ячейки значения.
Private Function RectSetParameterValidationHasOptions(ByVal parameterName As String, ByVal expectedOptions As Variant) As Boolean
    On Error GoTo Failed

    Dim settings As Object
    Set settings = ThisWorkbook.Names.Item("rngRectSetGeometry").RefersToRange

    Dim rowIndex As Long
    For rowIndex = 1 To settings.Rows.Count
        If StrComp(CStr(settings.Cells.Item(rowIndex, 1).Value2), parameterName, vbTextCompare) = 0 Then
            RectSetParameterValidationHasOptions = ValidationCellHasOptions(settings.Cells.Item(rowIndex, 2), expectedOptions)
            Exit Function
        End If
    Next rowIndex
    Exit Function
Failed:
End Function

' Проверяет выпадающие списки в табличной части дополнительных рядов RectSet.
' Helper ищет строку грани и нужный заголовок, поэтому вставка строк выше
' блока геометрии не меняет предмет проверки.
Private Function RectSetExtraValidationHasOptions(ByVal faceCaption As String, ByVal headerCaption As String, _
        ByVal occurrenceIndex As Long, ByVal expectedOptions As Variant) As Boolean
    On Error GoTo Failed

    Dim settings As Object
    Set settings = ThisWorkbook.Names.Item("rngRectSetGeometry").RefersToRange

    Dim headerRow As Long
    Dim rowIndex As Long
    For rowIndex = 1 To settings.Rows.Count
        If StrComp(CStr(settings.Cells.Item(rowIndex, 1).Value2), "Грань", vbTextCompare) = 0 And _
                StrComp(CStr(settings.Cells.Item(rowIndex, 2).Value2), "d2", vbTextCompare) = 0 Then
            headerRow = rowIndex
            Exit For
        End If
    Next rowIndex
    If headerRow = 0 Then Exit Function

    Dim valueColumn As Long
    Dim foundCount As Long
    Dim columnIndex As Long
    For columnIndex = 1 To settings.Columns.Count
        If StrComp(CStr(settings.Cells.Item(headerRow, columnIndex).Value2), headerCaption, vbTextCompare) = 0 Then
            foundCount = foundCount + 1
            If foundCount = occurrenceIndex Then
                valueColumn = columnIndex
                Exit For
            End If
        End If
    Next columnIndex
    If valueColumn = 0 Then Exit Function

    For rowIndex = headerRow + 1 To settings.Rows.Count
        If StrComp(CStr(settings.Cells.Item(rowIndex, 1).Value2), faceCaption, vbTextCompare) = 0 Then
            RectSetExtraValidationHasOptions = ValidationCellHasOptions(settings.Cells.Item(rowIndex, valueColumn), expectedOptions)
            Exit Function
        End If
    Next rowIndex
    Exit Function
Failed:
End Function

' Сверяет dropdown первой строки LC с полным допустимым набором вариантов.
' Проверяется actual validation ячейки, а не текст каталога сборки.
Private Function LoadCombinationValidationHasOptions(ByVal valueColumn As Long, ByVal expectedOptions As Variant) As Boolean
    On Error GoTo Failed

    Dim loads As Object
    Set loads = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange
    LoadCombinationValidationHasOptions = ValidationCellHasOptions(loads.Cells.Item(2, valueColumn), expectedOptions)
    Exit Function
Failed:
End Function

' Проверяет, что dropdown LC содержит формульные display names всех профилей.
' Измененное имя профиля должно попадать в список без пересборки исходного VBA.
Private Function LoadProfileValidationUsesDisplayNames() As Boolean
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
    LoadProfileValidationUsesDisplayNames = True
    Exit Function
Failed:
End Function

' Проверяет, что профиль в строке LC остается пользовательским вводом, не формулой.
Private Function LoadProfileFirstValueHasNoFormula() As Boolean
    On Error GoTo Failed

    Dim loads As Object
    Set loads = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange
    LoadProfileFirstValueHasNoFormula = Not CBool(loads.Cells.Item(2, 5).HasFormula)
    Exit Function
Failed:
End Function

' Проверяет точный состав и порядок actual validation-list одной ячейки fixture.
' Нечитаемый или неформульный список возвращает False, не пропускает assertion.
Private Function ValidationCellHasOptions(ByVal target As Object, ByVal expectedOptions As Variant) As Boolean
    On Error GoTo Failed

    Dim formulaText As String
    formulaText = CStr(target.Validation.Formula1)
    If Left$(formulaText, 1) <> "=" Then Exit Function

    Dim listRange As Object
    Set listRange = target.Worksheet.Range(Mid$(formulaText, 2))
    Dim i As Long
    If listRange.Cells.Count <> (UBound(expectedOptions) - LBound(expectedOptions) + 1) Then Exit Function
    For i = LBound(expectedOptions) To UBound(expectedOptions)
        If CStr(listRange.Cells.Item(i - LBound(expectedOptions) + 1, 1).Value2) <> CStr(expectedOptions(i)) Then Exit Function
    Next i
    ValidationCellHasOptions = True
    Exit Function
Failed:
End Function

' Собирает тестовый круг и арматуру из настоящего Config, подключает материалы,
' solver-options и каталог профилей. Reader сочетаний затем вызывается отдельно.
Private Function BuildUiBatch() As CBatchSectionCalculator
    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook

    Dim geom As CGeometryCircle
    Set geom = New CGeometryCircle
    geom.InitializeByDiameter settings.GetDouble("Circle.Diameter", 300#), _
        0#, _
        0#

    Dim mesh As CFiberMeshBuilder
    Set mesh = New CFiberMeshBuilder
    mesh.BuildMesh geom, 20#, 20#, 1

    Dim builder As CCircleRebarLayoutBuilder
    Set builder = New CCircleRebarLayoutBuilder
    Dim rebars As CRebarLayout
    Set rebars = builder.Build(settings.GetDouble("Circle.Diameter", 300#), _
        0#, _
        0#, _
        settings.GetDouble("Rebar.AxisDistance", 40#), _
        settings.GetLong("Rebar.Count", 0), _
        settings.GetDouble("Rebar.Diameter", 20#), _
        "Rebar")
    Dim materialProvider As CMaterialModelProvider
    Set materialProvider = New CMaterialModelProvider
    materialProvider.Initialize settings

    Dim section As CSectionModel
    Set section = BuildGeneratedSectionModel(mesh, rebars, "WorkbookInterface")

    Dim batch As CBatchSectionCalculator
    Set batch = New CBatchSectionCalculator
    batch.Initialize section, materialProvider
    batch.ApplySettings settings
    Dim profiles As CCalculationProfileCatalog
    Set profiles = New CCalculationProfileCatalog
    profiles.LoadFromWorkbook ThisWorkbook
    Set batch.ProfileCatalog = profiles
    Set BuildUiBatch = batch
End Function

' Готовит круговой SLS fixture через реальные ячейки Config и одну строку LC.
' Единицы/знаки приведены к внутренним; вызывающий тест работает на отдельной книге.
Private Sub PrepareCircleInput()
    SetSystemSetting "Units.Force.Input", "N"
    SetSystemSetting "Units.Moment.Input", "N*mm"
    SetSystemSetting "Units.Length.Input", "mm"
    SetSystemSetting "Units.Area.Input", "mm2"
    SetSystemSetting "Units.Stress.Input", "MPa"
    SetSystemSetting "Units.Curvature.Input", "1/mm"
    SetSystemSetting "Sign.N.User", "Tension"
    SetSystemSetting "Sign.Mx.User", "+Y tension"
    SetSystemSetting "Sign.My.User", "+X tension"
    SetSystemSetting "Geometry.Source", "Generated"
    SetSystemSetting "Geometry.Type", "Circle"
    SetSystemSetting "Plot.AutoUpdateAfterCalculation", "Yes"
    SetSystemSetting "Plot.LoadCase", "LC1"
    SetSystemSetting "Rebar.AxisDistance", "40"
    SetSystemSetting "Rebar.Count", "8"
    SetSystemSetting "Rebar.Diameter", "20"
    SetSystemSetting "SLS.Crack.Allowable", "0.3"
    SetProfileSetting "PR2", "Visualization.State", "CrackedState"

    Dim loads As Object
    Set loads = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange
    ClearDataRows loads
    loads.Cells.Item(2, 1).Value2 = "LC1"
    loads.Cells.Item(2, 2).Value2 = -100000#
    loads.Cells.Item(2, 3).Value2 = -4000000#
    loads.Cells.Item(2, 4).Value2 = -3000000#
    loads.Cells.Item(2, 5).Value2 = "PR2"
    loads.Cells.Item(2, 6).Value2 = ChrW$(&H3BB) & "*Mxy"
    loads.Cells.Item(2, 7).Value2 = "ui test"
End Sub

' Готовит RectSet fixture, активные грани и нагрузку через обычные таблицы Config.
' Неактивные грани отключаются явно, чтобы не зависеть от прежних пользовательских данных.
Private Sub PrepareRectSetInput()
    SetSystemSetting "Units.Force.Input", "N"
    SetSystemSetting "Units.Moment.Input", "N*mm"
    SetSystemSetting "Units.Length.Input", "mm"
    SetSystemSetting "Units.Area.Input", "mm2"
    SetSystemSetting "Units.Stress.Input", "MPa"
    SetSystemSetting "Units.Curvature.Input", "1/mm"
    SetSystemSetting "Sign.N.User", "Tension"
    SetSystemSetting "Sign.Mx.User", "+Y tension"
    SetSystemSetting "Sign.My.User", "+X tension"
    SetSystemSetting "Geometry.Source", "Generated"
    SetSystemSetting "Geometry.Type", "RectSet"
    SetSystemSetting "Plot.LoadCase", "LC_L"
    SetSystemSetting "Mesh.StepX", "40"
    SetSystemSetting "Mesh.StepY", "40"
    SetSystemSetting "Mesh.BoundarySubdivisions", "2"
    SetSystemSetting "RectSet.B1", "160"
    SetSystemSetting "RectSet.H1", "280"
    SetSystemSetting "RectSet.B2", "360"
    SetSystemSetting "RectSet.H2", "140"
    SetSystemSetting "RectSet.H1.as_1", "40"
    SetSystemSetting "RectSet.H1.as_2", "40"
    SetSystemSetting "RectSet.H1.d_1", "16"
    SetSystemSetting "RectSet.H1.d_2", "16"
    SetSystemSetting "RectSet.H1.n_1", "3"
    SetSystemSetting "RectSet.H1.n_2", "2"
    SetSystemSetting "RectSet.H1.StartOffset1", "40"
    SetSystemSetting "RectSet.H1.EndOffset1", "40"
    SetSystemSetting "RectSet.H1.StartOffset2", "40"
    SetSystemSetting "RectSet.H1.EndOffset2", "40"
    SetSystemSetting "RectSet.H2.as_1", "40"
    SetSystemSetting "RectSet.H2.as_2", "40"
    SetSystemSetting "RectSet.H2.d_1", "16"
    SetSystemSetting "RectSet.H2.d_2", "16"
    SetSystemSetting "RectSet.H2.n_1", "0"
    SetSystemSetting "RectSet.H2.n_2", "0"
    SetSystemSetting "RectSet.H2.StartOffset1", "40"
    SetSystemSetting "RectSet.H2.EndOffset1", "40"
    SetSystemSetting "RectSet.H2.StartOffset2", "40"
    SetSystemSetting "RectSet.H2.EndOffset2", "40"
    SetSystemSetting "RectSet.B1.as_1", "40"
    SetSystemSetting "RectSet.B1.as_2", "40"
    SetSystemSetting "RectSet.B1.d_1", "16"
    SetSystemSetting "RectSet.B1.d_2", "16"
    SetSystemSetting "RectSet.B1.n_1", "2"
    SetSystemSetting "RectSet.B1.n_2", "1"
    SetSystemSetting "RectSet.B1.StartOffset1", "20"
    SetSystemSetting "RectSet.B1.EndOffset1", "20"
    SetSystemSetting "RectSet.B1.StartOffset2", "20"
    SetSystemSetting "RectSet.B1.EndOffset2", "20"
    SetSystemSetting "RectSet.B2.as_1", "40"
    SetSystemSetting "RectSet.B2.as_2", "40"
    SetSystemSetting "RectSet.B2.d_1", "16"
    SetSystemSetting "RectSet.B2.d_2", "16"
    SetSystemSetting "RectSet.B2.n_1", "2"
    SetSystemSetting "RectSet.B2.n_2", "0"
    SetSystemSetting "RectSet.B2.StartOffset1", "60"
    SetSystemSetting "RectSet.B2.EndOffset1", "60"
    SetSystemSetting "RectSet.B2.StartOffset2", "60"
    SetSystemSetting "RectSet.B2.EndOffset2", "60"
    SetSystemSetting "SLS.Crack.Allowable", "0.3"

    Dim loads As Object
    Set loads = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange
    ClearDataRows loads
    loads.Cells.Item(2, 1).Value2 = "LC_L"
    loads.Cells.Item(2, 2).Value2 = -80000#
    loads.Cells.Item(2, 3).Value2 = -1500000#
    loads.Cells.Item(2, 4).Value2 = -1000000#
    loads.Cells.Item(2, 5).Value2 = "PR2"
    loads.Cells.Item(2, 6).Value2 = ChrW$(&H3BB) & "*Mxy"
    loads.Cells.Item(2, 7).Value2 = "rectset ui test"
End Sub

' Настраивает книгу под типовой прочностной расчет Г-сечения:
' пользователь задает N и Mx, а несущая способность ищется увеличением Mx.
' Здесь специально выбран Capacity.SolutionStrategy = UltimateStrain, чтобы проверить,
' что быстрый изгибный путь работает без fallback на LoadMultiplier.
Private Sub PrepareUserRectSetMomentUltimateInput()
    SetSystemSetting "Units.Force.Input", "tf"
    SetSystemSetting "Units.Moment.Input", "tf*m"
    SetSystemSetting "Units.Length.Input", "mm"
    SetSystemSetting "Units.Area.Input", "mm2"
    SetSystemSetting "Units.Stress.Input", "MPa"
    SetSystemSetting "Units.Curvature.Input", "1/mm"
    SetSystemSetting "Sign.N.User", "Compression"
    SetSystemSetting "Sign.Mx.User", "+Y tension"
    SetSystemSetting "Sign.My.User", "+X tension"
    SetSystemSetting "Geometry.Source", "Generated"
    SetSystemSetting "Geometry.Type", "RectSet"
    SetSystemSetting "Capacity.SolutionStrategy", "UltimateStrain"
    SetSystemSetting "Capacity.MaxLambda", "64"
    SetSystemSetting "Capacity.ToleranceStrain", "0.00001"
    SetSystemSetting "Capacity.SolverMaxIterations", "60"
    SetSystemSetting "Solver.Method", "Newton"
    SetSystemSetting "General.DiagramExtension", "Yes"
    SetSystemSetting "Solver.MaxIterations", "80"
    SetSystemSetting "Solver.LoadSteps", "1"
    SetSystemSetting "Mesh.StepX", "50"
    SetSystemSetting "Mesh.StepY", "50"
    SetSystemSetting "Mesh.BoundarySubdivisions", "1"
    SetSystemSetting "Load.ReferenceOffsetX", "0"
    SetSystemSetting "Load.ReferenceOffsetY", "0"
    SetSystemSetting "Plot.LoadCase", "LC_MX"
    SetSystemSetting "General.ExecutionReportEnabled", "No"

    SetSystemSetting "RectSet.H1", "550"
    SetSystemSetting "RectSet.B1", "250"
    SetSystemSetting "RectSet.H2", "250"
    SetSystemSetting "RectSet.B2", "600"

    SetUserRectSetMainRow "H1", 5, 5
    SetUserRectSetMainRow "B1", 2, 2
    SetUserRectSetMainRow "H2", 2, 2
    SetUserRectSetMainRow "B2", 5, 5
    ClearUserRectSetExtraRows

    Dim loads As Object
    Set loads = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange
    ClearDataRows loads
    loads.Cells.Item(2, 1).Value2 = "LC_MX"
    loads.Cells.Item(2, 2).Value2 = 200#
    loads.Cells.Item(2, 3).Value2 = 50#
    loads.Cells.Item(2, 4).ClearContents
    loads.Cells.Item(2, 5).Value2 = "PR1"
    loads.Cells.Item(2, 6).Value2 = ChrW$(&H3BB) & "*Mx"
    loads.Cells.Item(2, 7).Value2 = "moment ultimate regression"
End Sub

' Настраивает книгу ровно под пользовательский пример с Г-сечением и осевым
' растяжением в tf. Это дополняет unit-тест batch-слоя проверкой полного
' Excel-пути: пользовательский знак N, единицы, чтение Config и запись Results.
Private Sub PrepareUserRectSetAxialTensionInput()
    SetSystemSetting "Units.Force.Input", "tf"
    SetSystemSetting "Units.Moment.Input", "tf*m"
    SetSystemSetting "Units.Length.Input", "mm"
    SetSystemSetting "Units.Area.Input", "mm2"
    SetSystemSetting "Units.Stress.Input", "MPa"
    SetSystemSetting "Units.Curvature.Input", "1/mm"
    SetSystemSetting "Sign.N.User", "Compression"
    SetSystemSetting "Sign.Mx.User", "+Y tension"
    SetSystemSetting "Sign.My.User", "+X tension"
    SetSystemSetting "Geometry.Source", "Generated"
    SetSystemSetting "Geometry.Type", "RectSet"
    SetSystemSetting "Solver.Method", "Newton"
    SetSystemSetting "General.DiagramExtension", "Yes"
    SetSystemSetting "Solver.MaxIterations", "80"
    SetSystemSetting "Solver.LoadSteps", "1"
    SetSystemSetting "Mesh.StepX", "50"
    SetSystemSetting "Mesh.StepY", "50"
    SetSystemSetting "Mesh.BoundarySubdivisions", "1"
    SetSystemSetting "Load.ReferenceOffsetX", "0"
    SetSystemSetting "Load.ReferenceOffsetY", "0"
    SetSystemSetting "Plot.LoadCase", "LC_OVER"
    SetSystemSetting "General.ExecutionReportEnabled", "Yes"

    SetSystemSetting "RectSet.H1", "550"
    SetSystemSetting "RectSet.B1", "250"
    SetSystemSetting "RectSet.H2", "250"
    SetSystemSetting "RectSet.B2", "600"

    SetUserRectSetMainRow "H1", 5, 5
    SetUserRectSetMainRow "B1", 2, 2
    SetUserRectSetMainRow "H2", 2, 2
    SetUserRectSetMainRow "B2", 5, 5
    ClearUserRectSetExtraRows

    Dim loads As Object
    Set loads = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange
    ClearDataRows loads
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

' Включает сохранение прочностных и трещинных состояний выбранного тестового профиля.
' Каждая роль материала задается явно; Visualization выбирает уже сохраненный State.
Private Sub PrepareFullStateProfile(ByVal profileId As String)
    SetProfileSetting profileId, "Profile.DisplayName", "Полный snapshot"
    SetProfileSetting profileId, "Profile.Description", "Тестовая запись всех состояний"
    SetProfileSetting profileId, "Calculation.Strength.DirectState", "Yes"
    SetProfileSetting profileId, "Calculation.Strength.Capacity", "Yes"
    SetProfileSetting profileId, "Calculation.Crack.Width", "Yes"
    SetProfileSetting profileId, "MaterialModel.Strength.ValueSet", "ULS(I)"
    SetProfileSetting profileId, "MaterialModel.Strength.ConcreteDiagram", "TwoLine"
    SetProfileSetting profileId, "MaterialModel.Strength.ConcreteTension", "Ignore"
    SetProfileSetting profileId, "MaterialModel.Strength.SteelDiagram", "TwoLine"
    SetProfileSetting profileId, "MaterialModel.CrackInitiation.ValueSet", "SLS(II)"
    SetProfileSetting profileId, "MaterialModel.CrackInitiation.ConcreteDiagram", "ThreeLine"
    SetProfileSetting profileId, "MaterialModel.CrackInitiation.ConcreteTension", "UseDiagram"
    SetProfileSetting profileId, "MaterialModel.CrackInitiation.SteelDiagram", "TwoLine"
    SetProfileSetting profileId, "MaterialModel.CrackedState.ValueSet", "SLS(II)"
    SetProfileSetting profileId, "MaterialModel.CrackedState.ConcreteDiagram", "TwoLine"
    SetProfileSetting profileId, "MaterialModel.CrackedState.ConcreteTension", "Ignore"
    SetProfileSetting profileId, "MaterialModel.CrackedState.SteelDiagram", "TwoLine"
    SetProfileSetting profileId, "Visualization.State", "StrengthState"
    SetProfileSetting profileId, "Visualization.Quantity", "Stress"
End Sub

' Заполняет обе стороны грани RectSet фиксированными диаметрами и отступами.
' Количество первых рядов задается входами; дополнительные ряды очищаются отдельно.
Private Sub SetUserRectSetMainRow(ByVal faceName As String, ByVal count1 As Long, ByVal count2 As Long)
    SetSystemSetting "RectSet." & faceName & ".as_1", "40"
    SetSystemSetting "RectSet." & faceName & ".as_2", "40"
    SetSystemSetting "RectSet." & faceName & ".d_1", "32"
    SetSystemSetting "RectSet." & faceName & ".d_2", "32"
    SetSystemSetting "RectSet." & faceName & ".n_1", CStr(count1)
    SetSystemSetting "RectSet." & faceName & ".n_2", CStr(count2)
    SetSystemSetting "RectSet." & faceName & ".StartOffset1", "80"
    SetSystemSetting "RectSet." & faceName & ".EndOffset1", "80"
    SetSystemSetting "RectSet." & faceName & ".StartOffset2", "80"
    SetSystemSetting "RectSet." & faceName & ".EndOffset2", "80"
End Sub

' Отключает дополнительные ряды всех граней пустыми диаметрами по контракту Config.
' Режимы размещения/привязки задаются явно для воспроизводимости последующих тестов.
Private Sub ClearUserRectSetExtraRows()
    Dim faces As Variant
    faces = Array("H1", "B1", "H2", "B2")
    Dim i As Long
    For i = LBound(faces) To UBound(faces)
        SetSystemSetting "RectSet." & CStr(faces(i)) & ".d_2row_1", vbNullString
        SetSystemSetting "RectSet." & CStr(faces(i)) & ".d_2row_2", vbNullString
        SetSystemSetting "RectSet." & CStr(faces(i)) & ".d_3row_1", vbNullString
        SetSystemSetting "RectSet." & CStr(faces(i)) & ".d_3row_2", vbNullString
        SetSystemSetting "RectSet." & CStr(faces(i)) & ".loc_2row_1", "Stacked"
        SetSystemSetting "RectSet." & CStr(faces(i)) & ".loc_3row_1", "Stacked"
        SetSystemSetting "RectSet." & CStr(faces(i)) & ".bind_2row_1", "EachBar"
        SetSystemSetting "RectSet." & CStr(faces(i)) & ".bind_3row_1", "EachBar"
        SetSystemSetting "RectSet." & CStr(faces(i)) & ".loc_2row_2", "Stacked"
        SetSystemSetting "RectSet." & CStr(faces(i)) & ".loc_3row_2", "Stacked"
        SetSystemSetting "RectSet." & CStr(faces(i)) & ".bind_2row_2", "EachBar"
        SetSystemSetting "RectSet." & CStr(faces(i)) & ".bind_3row_2", "EachBar"
    Next i
End Sub

' Очищает только строки данных тестового диапазона, сохраняя его строку заголовков.
Private Sub ClearDataRows(ByVal target As Object)
    Dim rowIndex As Long
    Dim colIndex As Long
    For rowIndex = 2 To target.Rows.Count
        For colIndex = 1 To target.Columns.Count
            target.Cells.Item(rowIndex, colIndex).ClearContents
        Next colIndex
    Next rowIndex
End Sub

' Проверяет, что единицы у допусков равновесия не зашиты текстом.
' Эти значения вводятся пользователем в выбранных INPUT-единицах, поэтому
' колонка "Ед." должна ссылаться формулой на rngUnitSettings. Иначе при
' переходе, например, на tf*m можно случайно получить гигантский допуск
' момента и принять несбалансированное состояние как сошедшееся.
Private Sub TestSolverToleranceUnitLabels(ByRef stats As TUiTestStats)
    AssertTrue stats, "ui.units.solverToleranceN.dynamic", _
        SystemSettingUnitCellReferencesQuantity("Solver.ToleranceN", "Force")
    AssertTrue stats, "ui.units.solverToleranceMx.dynamic", _
        SystemSettingUnitCellReferencesQuantity("Solver.ToleranceMx", "Moment")
    AssertTrue stats, "ui.units.solverToleranceMy.dynamic", _
        SystemSettingUnitCellReferencesQuantity("Solver.ToleranceMy", "Moment")
End Sub

' Проверяет, что настройки поиска несущей способности не выглядят как
' недооформленный блок: у безразмерных коэффициентов стоит "-", у счетчиков -
' "шт". Размерных величин в этом блоке сейчас нет, поэтому формулы единиц
' здесь не нужны; важно именно не оставлять пустую колонку "Ед.".
Private Sub TestCapacitySettingsUnitLabels(ByRef stats As TUiTestStats)
    AssertTextEquals stats, "ui.units.Capacity.SolutionStrategy", SystemSettingUnitText("Capacity.SolutionStrategy"), "-"
    AssertTextEquals stats, "ui.units.capacity.searchMethod", SystemSettingUnitText("Capacity.SearchMethod"), "-"
    AssertTextEquals stats, "ui.units.capacity.initialLambda", SystemSettingUnitText("Capacity.InitialLambda"), "-"
    AssertTextEquals stats, "ui.units.capacity.toleranceLambda", SystemSettingUnitText("Capacity.ToleranceLambda"), "-"
    AssertTextEquals stats, "ui.units.capacity.maxRetries", SystemSettingUnitText("Capacity.MaxRetries"), "шт"
    AssertTextEquals stats, "ui.units.capacity.baseLoadSteps", SystemSettingUnitText("Capacity.BaseLoadSteps"), "шт"
    AssertTextEquals stats, "ui.units.capacity.toleranceStrain", SystemSettingUnitText("Capacity.ToleranceStrain"), "-"
    AssertTextEquals stats, "ui.units.capacity.maxLambda", SystemSettingUnitText("Capacity.MaxLambda"), "-"
    AssertTextEquals stats, "ui.units.capacity.solverIterations", SystemSettingUnitText("Capacity.SolverMaxIterations"), "шт"
    AssertTrue stats, "ui.units.meshStepX.dynamic", _
        SystemSettingUnitCellReferencesQuantity("Mesh.StepX", "Length")
    AssertTrue stats, "ui.units.meshStepY.dynamic", _
        SystemSettingUnitCellReferencesQuantity("Mesh.StepY", "Length")
    AssertTrue stats, "ui.units.zeroMomentPerDepth.dynamic", _
        SystemSettingUnitCellReferencesQuantities("Calculation.ZeroMomentPerDepth", "Moment", "Length")
End Sub

' Возвращает фактическую подпись единиц настройки для проверки собранной книги.
' Формульную зависимость от выбранных единиц проверяет отдельный helper.
Private Function SystemSettingUnitText(ByVal key As String) As String
    On Error GoTo Failed

    Dim settings As Object
    Set settings = ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange

    Dim rowIndex As Long
    For rowIndex = 2 To settings.Rows.Count
        If StrComp(CStr(settings.Cells.Item(rowIndex, 1).Value2), key, vbTextCompare) = 0 Then
            SystemSettingUnitText = Trim$(CStr(settings.Cells.Item(rowIndex, 3).Value2))
            Exit Function
        End If
    Next rowIndex

Failed:
End Function

' Меняет единицы текста обеих групп и проверяет пересчет формульной подписи.
' Исходные значения восстанавливаются и после неуспешного чтения fixture.
Private Function PlotAnnotationTextUnitCellIsDynamic(ByVal rowName As String) As Boolean
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
    PlotAnnotationTextUnitCellIsDynamic = (CStr(settings.Cells.Item(targetRow, 4).Value2) = "pt")

CleanUp:
    If Not valuesSaved Then Exit Function
    settings.Cells.Item(textUnitsRow, 2).Value2 = oldRebarUnits
    settings.Cells.Item(textUnitsRow, 3).Value2 = oldDimensionUnits
    settings.Worksheet.Calculate
    Exit Function
Failed:
    If valuesSaved Then Resume CleanUp
End Function

' Читает численную настройку заданной группы аннотаций для независимой UI-проверки.
' Метод не строит схему и не запускает расчет модели.
Private Function PlotAnnotationSettingValue(ByVal rowName As String, ByVal valueColumn As Long) As Double
    On Error GoTo Failed

    Dim settings As Object
    Set settings = ThisWorkbook.Names.Item("rngPlotAnnotationSettings").RefersToRange

    Dim rowIndex As Long
    For rowIndex = 2 To settings.Rows.Count
        If StrComp(CStr(settings.Cells.Item(rowIndex, 1).Value2), rowName, vbTextCompare) = 0 Then
            PlotAnnotationSettingValue = CDbl(settings.Cells.Item(rowIndex, valueColumn).Value2)
            Exit Function
        End If
    Next rowIndex

Failed:
End Function

' Сравнивает пользовательскую подпись без учета регистра, сохраняя оба текста при отказе.
Private Sub AssertTextEquals(ByRef stats As TUiTestStats, ByVal name As String, _
        ByVal actual As String, ByVal expected As String)
    If StrComp(actual, expected, vbTextCompare) = 0 Then
        stats.Passed = stats.Passed + 1
        AppendLine stats, "OK: " & name & "; actual=" & actual
    Else
        stats.Failed = stats.Failed + 1
        AppendLine stats, "FAIL: " & name & "; actual=" & actual & "; expected=" & expected
    End If
End Sub

' Группа файловых helpers работает с отчетами, созданными текущим UI-тестом.
' Удаление ниже допускается только для известного пути тестового артефакта.
Private Function FileExists(ByVal path As String) As Boolean
    FileExists = CreateObject("Scripting.FileSystemObject").FileExists(path)
End Function

' Удаляет существующий отчет по известному пути текущего UI-теста;
' отсутствие файла допустимо, произвольные пользовательские пути не передаются.
Private Sub DeleteFileIfExists(ByVal path As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(path) Then fso.DeleteFile path, True
End Sub

' Читает настоящий Unicode-отчет для сравнения с выводимыми комментариями результата.
Private Function ReadTextFile(ByVal path As String) As String
    Dim stream As Object
    Set stream = CreateObject("Scripting.FileSystemObject").OpenTextFile(path, 1, False, -1)
    ReadTextFile = stream.ReadAll
    stream.Close
End Function

' Группа assertions ведет счет проверок actual интерфейса и протоколирует отказ.
' Проверки чисел ниже используют фиксированный допуск вызывающего сценария.
Private Sub AssertTrue(ByRef stats As TUiTestStats, ByVal name As String, ByVal condition As Boolean)
    If condition Then
        stats.Passed = stats.Passed + 1
        AppendLine stats, "OK: " & name
    Else
        stats.Failed = stats.Failed + 1
        AppendLine stats, "FAIL: " & name
    End If
End Sub

' Сверяет абсолютное отклонение UI-величины, сохраняя actual/expected в отчете.
Private Sub AssertClose(ByRef stats As TUiTestStats, ByVal name As String, ByVal actual As Double, _
        ByVal expected As Double, ByVal tolerance As Double)
    If Abs(actual - expected) <= tolerance Then
        stats.Passed = stats.Passed + 1
        AppendLine stats, "OK: " & name & "; actual=" & FormatNumberInvariant(actual) & _
            "; expected=" & FormatNumberInvariant(expected)
    Else
        stats.Failed = stats.Failed + 1
        AppendLine stats, "FAIL: " & name & "; actual=" & FormatNumberInvariant(actual) & _
            "; expected=" & FormatNumberInvariant(expected)
    End If
End Sub

' Добавляет диагностическую строку к протоколу интерфейсного набора.
Private Sub AppendLine(ByRef stats As TUiTestStats, ByVal text As String)
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

Private Function FormatNumberInvariant(ByVal value As Double) As String
    FormatNumberInvariant = Replace$(Format$(value, "0.############"), ",", ".")
End Function

' ДЛЯ ТЕСТОВ
' Изолирует один полный расчет книги от остальных UI-сценариев. Диагностика
' используется для поиска роста памяти; основной regression-набор не заменяет.
Public Function RunAudit02SingleCalculationDiagnostic() As String
    On Error GoTo Failed
    Dim stats As TUiTestStats
    SetSystemSetting "General.ExecutionReportEnabled", "No"
    TestCircleWorkbookRunWritesResults stats
    RunAudit02SingleCalculationDiagnostic = stats.Report & _
        "TOTAL_UI_DIAGNOSTIC: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    Exit Function
Failed:
    RunAudit02SingleCalculationDiagnostic = stats.Report & "RUNTIME ERROR: " & CStr(Err.Number) & "; " & Err.Description
End Function

' ДЛЯ ТЕСТОВ
' Изолирует запись execution_report на свежей книге, чтобы отличить расходы
' отчета от накопления данных между последовательными UI-расчетами.
Public Function RunAudit02ReportDiagnostic() As String
    On Error GoTo Failed
    Dim stats As TUiTestStats
    TestExecutionReportFile stats
    RunAudit02ReportDiagnostic = stats.Report & _
        "TOTAL_UI_DIAGNOSTIC: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    Exit Function
Failed:
    RunAudit02ReportDiagnostic = stats.Report & "RUNTIME ERROR: " & CStr(Err.Number) & "; " & Err.Description
End Function

' ДЛЯ ТЕСТОВ
' Повторяет расчетную часть workbook-сценария без writer-ов и схемы.
' Возвращает счетчик solve и время; физических ожиданий не изменяет.
Public Function RunAudit02BatchOnlyDiagnostic() As String
    RunAudit02BatchOnlyDiagnostic = RunAudit02WorkbookPhases(False)
End Function

' ДЛЯ ТЕСТОВ
' Измеряет последовательные writer-ы на одинаковом каноническом результате.
' Каждая отметка содержит время и память единственного тестового Excel.
Public Function RunAudit02WriterPhasesDiagnostic() As String
    RunAudit02WriterPhasesDiagnostic = RunAudit02WorkbookPhases(True)
End Function

' ДЛЯ ТЕСТОВ
' Общая подготовка независимого диагностического сценария без основного UI.
Private Function RunAudit02WorkbookPhases(ByVal writeResults As Boolean) As String
    On Error GoTo Failed
    PrepareCircleInput
    SetSystemSetting "General.ExecutionReportEnabled", "No"
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
    RunAudit02WorkbookPhases = "INFO: batchOnly; status=" & batch.ResultAt(1).Status & _
        "; elapsedSec=" & FormatNumberInvariant(Timer - started) & "; solves=" & CStr(batch.SolverCallCount)
    If writeResults Then
        Dim stats As TUiTestStats
        AppendLine stats, "RUN: batch " & Audit02ExcelMemory()
        Dim strengthWriter As CStrengthSummaryWriter
        Set strengthWriter = New CStrengthSummaryWriter
        strengthWriter.WriteSummary ThisWorkbook, batch, units
        AppendLine stats, "RUN: strength " & Audit02ExcelMemory()
        Dim crackWriter As CCrackSummaryWriter
        Set crackWriter = New CCrackSummaryWriter
        crackWriter.WriteSummary ThisWorkbook, batch, units
        AppendLine stats, "RUN: crack " & Audit02ExcelMemory()
        Dim stabilityWriter As CStabilitySummaryWriter
        Set stabilityWriter = New CStabilitySummaryWriter
        stabilityWriter.WriteSummary ThisWorkbook, batch, units
        AppendLine stats, "RUN: stability " & Audit02ExcelMemory()
        Dim summaryWriter As CBatchResultWriter
        Set summaryWriter = New CBatchResultWriter
        summaryWriter.WriteSummary ThisWorkbook, batch, units, section
        AppendLine stats, "RUN: summary " & Audit02ExcelMemory()
        Dim ndmWriter As CNDMResultsWriter
        Set ndmWriter = New CNDMResultsWriter
        ndmWriter.WriteResults ThisWorkbook, section, PrepareSectionSnapshot(section, materials), batch, units
        AppendLine stats, "RUN: NDM " & Audit02ExcelMemory()
        RunAudit02WorkbookPhases = RunAudit02WorkbookPhases & vbCrLf & stats.Report
    End If
    Exit Function
Failed:
    RunAudit02WorkbookPhases = "RUNTIME ERROR: " & CStr(Err.Number) & "; " & Err.Description
End Function

' ДЛЯ ТЕСТОВ
' Читает память процесса через WMI, не меняя настройки или состояние Excel.
Private Function Audit02ExcelMemory() As String
    Dim process As Object
    For Each process In GetObject("winmgmts:").ExecQuery( _
            "SELECT ProcessId, WorkingSetSize, PrivatePageCount FROM Win32_Process WHERE Name='EXCEL.EXE'")
        If Len(Audit02ExcelMemory) > 0 Then Audit02ExcelMemory = Audit02ExcelMemory & "; "
        Audit02ExcelMemory = Audit02ExcelMemory & "pid=" & CStr(process.ProcessId) & _
            "; workingSet=" & CStr(process.WorkingSetSize) & "; privateBytes=" & CStr(process.PrivatePageCount)
    Next process
End Function

' ДЛЯ ТЕСТОВ: выполняет неизмененные UI/Config-проверки до адресного LC-блока.
' Runner может сохранить полученную книгу и воспроизвести LC на новом Excel,
' отличив зависимость от содержимого Results от накопления памяти процесса.
Public Function RunAudit03PreLoadTableDiagnosticTests() As String
    mAudit03StopBeforeLoadTable = True
    RunAudit03PreLoadTableDiagnosticTests = RunWorkbookInterfaceTests()
    mAudit03StopBeforeLoadTable = False
End Function

' ДЛЯ ТЕСТОВ: сохраняет неизмененный входной UI-prefix перед unit/sign-блоком.
' Повтор этого блока на сохраненной копии в свежем Excel отличает влияние
' Config/Results от ресурсов, накопленных предшествующими тестовыми вызовами.
Public Function RunAudit03InputPrefixDiagnosticTests() As String
    mAudit03StopBeforeUnitSign = True
    RunAudit03InputPrefixDiagnosticTests = RunWorkbookInterfaceTests()
    mAudit03StopBeforeUnitSign = False
End Function

' ДЛЯ ТЕСТОВ
' Записывает численные Results без обновления схемы; нужна для локализации
' роста памяти между расчетом, табличным выводом и построением Chart.
Public Function RunAudit02NoPlotDiagnostic() As String
    On Error GoTo Failed
    PrepareCircleInput
    SetSystemSetting "General.ExecutionReportEnabled", "No"
    SetSystemSetting "Plot.AutoUpdateAfterCalculation", "No"
    RunAudit02NoPlotDiagnostic = RunSectionCalculationForWorkbook(ThisWorkbook, False)
    Exit Function
Failed:
    RunAudit02NoPlotDiagnostic = "RUNTIME ERROR: " & CStr(Err.Number) & "; " & Err.Description
End Function

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

' ДЛЯ ТЕСТОВ
' Меняет Config после завершенного расчета, но не его сериализованный снимок.
' Reader и повторная схема обязаны сохранить напряжения, деформации, признаки
' State и таблицы диаграмм; чтение геометрии для AutoCAD также не решает НДС.
Private Sub TestAudit02SavedResultsIgnoreMaterialChanges(ByRef stats As TUiTestStats)
    PrepareCircleInput
    SetSystemSetting "Plot.AutoUpdateAfterCalculation", "No"
    Dim materialKeys As Variant
    materialKeys = Array("General.DiagramExtension", "Concrete.E", "Concrete.R.ULS(I)", _
        "Concrete.R.SLS(II)", "Steel.E", "Steel.R.ULS(I)", "Steel.R.SLS(II)")
    Dim oldValues(0 To 6) As String
    Dim i As Long
    For i = 0 To UBound(materialKeys)
        oldValues(i) = CStr(GetSystemSetting(CStr(materialKeys(i))))
    Next i
    Dim oldQuantity As String
    oldQuantity = ProfileSettingValue("PR2", "Visualization.Quantity")
    On Error GoTo RestoreFailed

    Dim message As String
    message = RunSectionCalculationForWorkbook(ThisWorkbook, False)
    AssertTrue stats, "audit02.saved.run", InStr(1, message, "Расчет завершен", vbTextCompare) > 0
    Dim tables As Variant
    tables = Array("rngNDMSectionGeometry", "rngNDMSectionContours", "rngNDMElementResults", "rngNDMSectionProperties", _
        "rngNDMSectionAnnotations", "rngNDMMaterialDiagrams", "rngBatchSummary")
    Dim beforeTables() As Variant
    ReDim beforeTables(0 To UBound(tables))
    For i = 0 To UBound(tables)
        beforeTables(i) = ResultTable(CStr(tables(i)))
    Next i
    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook
    Dim beforeStress As CSectionPlotDataReader
    Set beforeStress = New CSectionPlotDataReader
    SetProfileSetting "PR2", "Visualization.Quantity", "Stress"
    beforeStress.LoadFromWorkbook ThisWorkbook, settings
    SetSystemSetting "AutoCAD.Export.CombinationID", "LC1"
    Dim beforeExport As String
    beforeExport = Audit02ReadExportSnapshotForTests(ThisWorkbook)
    Dim beforeStrain As CSectionPlotDataReader
    Set beforeStrain = New CSectionPlotDataReader
    SetProfileSetting "PR2", "Visualization.Quantity", "Strain"
    beforeStrain.LoadFromWorkbook ThisWorkbook, settings
    Dim solves As Long
    solves = SectionEquilibriumSolveCount()

    If StrComp(oldValues(0), "Yes", vbTextCompare) = 0 Then
        SetSystemSetting "General.DiagramExtension", "No"
    Else
        SetSystemSetting "General.DiagramExtension", "Yes"
    End If
    For i = 1 To UBound(materialKeys)
        SetSystemSetting CStr(materialKeys(i)), FormatNumberInvariant(CDbl(oldValues(i)) * 0.5)
    Next i
    settings.LoadFromWorkbook ThisWorkbook
    Dim afterData As CSectionPlotDataReader
    Set afterData = New CSectionPlotDataReader
    SetProfileSetting "PR2", "Visualization.Quantity", "Stress"
    afterData.LoadFromWorkbook ThisWorkbook, settings
    AssertAudit02PlotSnapshot stats, "stress", beforeStress, afterData
    AssertTextEquals stats, "audit02.saved.exportSnapshot", _
        Audit02ReadExportSnapshotForTests(ThisWorkbook), beforeExport
    SetProfileSetting "PR2", "Visualization.Quantity", "Strain"
    afterData.LoadFromWorkbook ThisWorkbook, settings
    AssertAudit02PlotSnapshot stats, "strain", beforeStrain, afterData
    UpdateSectionPlotForWorkbook ThisWorkbook
    Dim exportedSection As CSectionModel
    Set exportedSection = ReadSectionGeometryFromResults(ThisWorkbook)
    AssertTrue stats, "audit02.saved.exportGeometry", exportedSection.ConcreteCount > 0
    AssertTrue stats, "audit02.saved.noSolve", SectionEquilibriumSolveCount() = solves
    Dim afterTable As Variant
    For i = 0 To UBound(tables)
        afterTable = ResultTable(CStr(tables(i)))
        AssertTrue stats, "audit02.saved.table." & CStr(tables(i)), _
            Audit02SnapshotTablesEqual(beforeTables(i), afterTable)
    Next i
    GoTo RestoreSettings
RestoreFailed:
    Dim failureText As String
    failureText = CStr(Err.Number) & "; " & Err.Description
    AssertTrue stats, "audit02.saved.runtime." & failureText, False
RestoreSettings:
    For i = 0 To UBound(materialKeys)
        SetSystemSetting CStr(materialKeys(i)), oldValues(i)
    Next i
    SetProfileSetting "PR2", "Visualization.Quantity", oldQuantity
End Sub

' ДЛЯ ТЕСТОВ
' Сравнивает сериализованную раскраску и плоскость без повторной оценки материала.
Private Sub AssertAudit02PlotSnapshot(ByRef stats As TUiTestStats, ByVal quantity As String, _
        ByVal beforeData As CSectionPlotDataReader, ByVal afterData As CSectionPlotDataReader)
    Dim prefix As String
    prefix = "audit02.saved." & quantity
    AssertTrue stats, prefix & ".nonempty", beforeData.Count > 0
    AssertTrue stats, prefix & ".count", beforeData.Count = afterData.Count
    AssertTextEquals stats, prefix & ".state", afterData.StateType, beforeData.StateType
    AssertTextEquals stats, prefix & ".status", afterData.DirectStateStatus, beforeData.DirectStateStatus
    AssertTrue stats, prefix & ".extension", afterData.ExtensionUsed = beforeData.ExtensionUsed
    AssertClose stats, prefix & ".epsilon0", afterData.Epsilon0, beforeData.Epsilon0, 0#
    AssertClose stats, prefix & ".kappaX", afterData.KappaX, beforeData.KappaX, 0#
    AssertClose stats, prefix & ".kappaY", afterData.KappaY, beforeData.KappaY, 0#
    Dim i As Long
    For i = 1 To beforeData.Count
        AssertTextEquals stats, prefix & ".id." & CStr(i), afterData.ElementID(i), beforeData.ElementID(i)
        AssertTextEquals stats, prefix & ".physical." & CStr(i), afterData.PhysicalState(i), beforeData.PhysicalState(i)
        AssertClose stats, prefix & ".value." & CStr(i), afterData.ResultValue(i), beforeData.ResultValue(i), 0#
    Next i
End Sub

' ДЛЯ ТЕСТОВ
' Сравнивает все ячейки двух снимков, включая model spec и effective Extension.
Private Function Audit02SnapshotTablesEqual(ByRef beforeTable As Variant, ByRef afterTable As Variant) As Boolean
    If UBound(beforeTable, 1) <> UBound(afterTable, 1) Then Exit Function
    If UBound(beforeTable, 2) <> UBound(afterTable, 2) Then Exit Function
    Dim r As Long
    Dim c As Long
    For r = 1 To UBound(beforeTable, 1)
        For c = 1 To UBound(beforeTable, 2)
            If CStr(beforeTable(r, c)) <> CStr(afterTable(r, c)) Then Exit Function
        Next c
    Next r
    Audit02SnapshotTablesEqual = True
End Function

' ========================== ДЛЯ ТЕСТОВ ==========================
' Проверяет контракт reader-а через реальные диапазоны независимой книги.
' Временный лист удаляется даже при ошибке; исходные Config-ячейки не меняются.
Public Function RunAudit03ReaderTests() As String
    Dim stats As TUiTestStats
    TestAudit03ReaderContract stats
    AppendLine stats, "TOTAL_AUDIT03_READER: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunAudit03ReaderTests = stats.Report
End Function

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
    AssertTrue stats, "audit03.reader.entry.runtime." & CStr(Err.Number) & "." & Err.Description, False
CleanUp:
    On Error Resume Next
    Application.DisplayAlerts = False
    If Not sheet Is Nothing Then sheet.Delete
    Application.DisplayAlerts = oldAlerts
    On Error GoTo 0
End Sub

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
    Set batch = BuildUiBatch()
    reader.LoadFromRange sheet.Range("A1").Resize(3, columns), batch
    Dim prefix As String
    prefix = "audit03.reader.width." & CStr(columns)
    AssertTrue stats, prefix & ".count", batch.Count = 1
    If batch.Count <> 1 Then Exit Sub
    AssertTextEquals stats, prefix & ".id", batch.CombinationID(1), "WIDTH_" & CStr(columns)
    AssertTextEquals stats, prefix & ".profile", batch.InputProfileId(1), "PR1"
    AssertTextEquals stats, prefix & ".comment", batch.CombinationName(1), expectedComment
    AssertTrue stats, prefix & ".sourceSlot", batch.SourceDataOffset(1) = 2
    AssertClose stats, prefix & ".n", batch.N(1), -10000#, 0#
    AssertClose stats, prefix & ".mx", batch.UserMx(1), 10000000#, 0#
    AssertClose stats, prefix & ".my", batch.UserMy(1), 20000000#, 0#
    If columns >= 7 Then
        batch.Execute
        AssertTextEquals stats, prefix & ".capacityPath", batch.ResultAt(1).StrengthResult.Capacity.PathResolved, "LambdaMx"
    End If
    Dim errorNumber As Long
    On Error Resume Next
    reader.LoadFromRange sheet.Range("A1:D3"), batch
    errorNumber = Err.Number
    Err.Clear
    On Error GoTo Failed
    AssertTrue stats, prefix & ".shortRangeControlled", errorNumber = vbObjectError + 3955
    Exit Sub
Failed:
    AssertTrue stats, "audit03.reader.width." & CStr(columns) & ".runtime." & CStr(Err.Number) & "." & Err.Description, False
End Sub

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
    AssertTrue stats, "audit03.reader.rows.stringFixture", VarType(sheet.Range("B8").Value2) = vbString
    sheet.Range("B9").Formula = "=2+3"
    sheet.Range("B9").Calculate
    Dim reader As CLoadCombinationReader
    Set reader = New CLoadCombinationReader
    Dim batch As CBatchSectionCalculator
    Set batch = BuildUiBatch()
    reader.LoadFromRange sheet.Range("A1:G13"), batch
    batch.SetSP35Table721 ThisWorkbook.Names.Item("rngSP35Table721").RefersToRange.Value2
    AssertTrue stats, "audit03.reader.rows.count", batch.Count = 10
    If batch.Count <> 10 Then Exit Sub
    AssertTrue stats, "audit03.reader.rows.sourceGap", batch.SourceDataOffset(5) = 7
    AssertTextEquals stats, "audit03.reader.rows.lastId", batch.CombinationID(10), "VALID_LAST"
    AssertTrue stats, "audit03.reader.rows.lastSlot", batch.SourceDataOffset(10) = 12
    AssertClose stats, "audit03.reader.rows.numericString", batch.N(5), 42.5, 0#
    AssertClose stats, "audit03.reader.rows.numericFormula", batch.N(6), 5#, 0#
    batch.Execute
    For r = 1 To 4
        AssertTextEquals stats, "audit03.reader.rows.invalid." & CStr(r), batch.ResultAt(r).Status, "InputErr"
        AssertTrue stats, "audit03.reader.rows.comment." & CStr(r), Len(batch.ResultAt(r).OverallMeta.ResultComment) > 0
    Next r
    For r = 7 To 9
        AssertTextEquals stats, "audit03.reader.rows.invalid." & CStr(r), batch.ResultAt(r).Status, "InputErr"
    Next r
    AppendLine stats, "COMMENT: audit03.reader.rows.valid; " & batch.ResultAt(10).OverallMeta.ResultComment
    AssertTrue stats, "audit03.reader.rows.validReachedResults", batch.ResultAt(10).Status <> "InputErr"
    Exit Sub
Failed:
    AssertTrue stats, "audit03.reader.rows.runtime." & CStr(Err.Number) & "." & Err.Description, False
End Sub

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
        AssertTrue stats, "audit03.reader.tiny." & CStr(mode) & ".count", batch.Count = 2
        If batch.Count = 2 Then
            AssertClose stats, "audit03.reader.tiny." & CStr(mode) & ".n", batch.N(1), 0.00000000001, 0.00000000000000000000000001
            AssertClose stats, "audit03.reader.tiny." & CStr(mode) & ".m", batch.UserMx(2), -0.00000000001, 0.00000000000000000000000001
        End If
    Next mode
    sheet.Range("B2").Value2 = 1E+308
    sheet.Range("A3").Value2 = "AFTER_OVERFLOW": sheet.Range("B3").Value2 = 1#
    Set batch = BuildUiBatch()
    reader.LoadFromRange sheet.Range("A1:G3"), batch, units
    AssertTrue stats, "audit03.reader.overflow.count", batch.Count = 2
    If batch.Count = 2 Then
        batch.Execute
        AssertTextEquals stats, "audit03.reader.overflow.status", batch.ResultAt(1).Status, "InputErr"
        AssertTrue stats, "audit03.reader.overflow.comment", Len(batch.ResultAt(1).OverallMeta.ResultComment) > 0
        AssertTextEquals stats, "audit03.reader.overflow.next", batch.CombinationID(2), "AFTER_OVERFLOW"
    End If
    Exit Sub
Failed:
    AssertTrue stats, "audit03.reader.tiny.runtime." & CStr(Err.Number) & "." & Err.Description, False
End Sub

' ========================== ДЛЯ ТЕСТОВ ==========================
' Проверяет неверный ввод через реальные Range и публичные reader/build API.
' Метод не изменяет пользовательские Config-ячейки и возвращает отдельный отчет.
Public Function RunAudit03InputContractTests() As String
    Dim stats As TUiTestStats
    TestAudit03InputContracts stats
    AppendLine stats, "TOTAL_AUDIT03_INPUT_CONTRACTS: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunAudit03InputContractTests = stats.Report
End Function

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
        AssertTrue stats, "audit03.input.boolean.valid." & CStr(value), _
            settings.GetBoolean("Audit03.Test") = (value = "Yes" Or value = "True" Or value = "1" Or value = "Да")
    Next value
    AssertTrue stats, "audit03.input.boolean.missingTrue", settings.GetBoolean("Audit03.Absent", True)
    AssertTrue stats, "audit03.input.boolean.missingFalse", Not settings.GetBoolean("Audit03.Absent", False)
    For Each value In Array("Maybe", "TODO", "")
        sheet.Range("B2").Value2 = CStr(value)
        Audit03AssertInputError stats, sheet.Range("A1:C2"), "Boolean", _
            "audit03.input.boolean.invalid." & CStr(value), vbObjectError + 4314
    Next value
    For Each value In Array(-2147483648#, 0#, 17#, 2147483647#)
        sheet.Range("B2").Value2 = CDbl(value)
        settings.LoadFromRange sheet.Range("A1:C2")
        AssertClose stats, "audit03.input.long.valid." & CStr(value), CDbl(settings.GetLong("Audit03.Test")), CDbl(value), 0#
    Next value
    For Each value In Array(1.25, -1.25, 1E+20)
        sheet.Range("B2").Value2 = CDbl(value)
        Audit03AssertInputError stats, sheet.Range("A1:C2"), "Long", _
            "audit03.input.long.invalid." & CStr(value), vbObjectError + 4315
    Next value
    sheet.Range("A2").Value2 = "Solver.LineSearchEnabled": sheet.Range("B2").Value2 = "Maybe"
    Audit03AssertInputError stats, sheet.Range("A1:C2"), "Solver", "audit03.input.consumer.boolean", vbObjectError + 4314
    sheet.Range("A2").Value2 = "Solver.MaxIterations": sheet.Range("B2").Value2 = 1.25
    Audit03AssertInputError stats, sheet.Range("A1:C2"), "Solver", "audit03.input.consumer.integer", vbObjectError + 4315
    sheet.Range("A2").Value2 = "Mesh.BoundarySubdivisions": sheet.Range("B2").Value2 = 0
    Audit03AssertInputError stats, sheet.Range("A1:C2"), "MeshRegistry", "audit03.input.consumer.boundary", vbObjectError + 4139
    For Each value In Array("", "TODO")
        sheet.Range("B2").Value2 = CStr(value)
        Audit03AssertInputError stats, sheet.Range("A1:C2"), "MeshRegistry", _
            "audit03.input.consumer.boundaryMissing." & CStr(value), vbObjectError + 4139
    Next value
    Audit03AssertInputError stats, sheet.Range("A1"), "Settings", "audit03.input.settings.scalar", vbObjectError + 4316
    Audit03AssertInputError stats, sheet.Range("A1:B2"), "Settings", "audit03.input.settings.narrow", vbObjectError + 4316
    Audit03AssertInputError stats, Nothing, "Settings", "audit03.input.settings.nothing", vbObjectError + 4316
    Audit03AssertInputError stats, sheet.Range("A1"), "Profiles", "audit03.input.profiles.scalar", vbObjectError + 3987
    Audit03AssertInputError stats, Nothing, "MeshNothing", "audit03.input.mesh.nothing", vbObjectError + 2107
    Audit03AssertInputError stats, Nothing, "MeshAxis", "audit03.input.mesh.axis", vbObjectError + 2108
    Audit03AssertInputError stats, Nothing, "MeshProduct", "audit03.input.mesh.product", vbObjectError + 2108
    Audit03AssertInputError stats, Nothing, "MeshBoundary", "audit03.input.mesh.boundary", vbObjectError + 2109
    Audit03AssertInputError stats, Nothing, "MeshSubcellProduct", "audit03.input.mesh.subcellProduct", vbObjectError + 2108
    GoTo CleanUp
Failed:
    AssertTrue stats, "audit03.input.runtime." & CStr(Err.Number) & "." & Err.Description, False
CleanUp:
    On Error Resume Next
    Application.DisplayAlerts = False
    If Not sheet Is Nothing Then sheet.Delete
    Application.DisplayAlerts = oldAlerts
    On Error GoTo 0
End Sub

' Проверяет точный контракт публичного API; описание ошибки сохраняется в логе
' для содержательной ревизии. Перехват относится только к ожидаемому test input.
Private Sub Audit03AssertInputError(ByRef stats As TUiTestStats, ByVal source As Object, _
        ByVal operation As String, ByVal prefix As String, ByVal expectedCode As Long)
    Dim description As String, actualCode As Long
    actualCode = Audit03CaptureInputError(source, operation, description)
    AssertTrue stats, prefix & ".typedError", actualCode = expectedCode
    If actualCode <> 0 Then AssertTrue stats, prefix & ".reason", Len(description) > 0
    AppendLine stats, "INPUT: " & prefix & "; error=" & CStr(actualCode) & "; expected=" & CStr(expectedCode) & "; reason=" & description
End Sub

' Запускает отдельную неверную операцию без маскировки ошибок следующего теста.
' Сетка с непредставимым числом ячеек должна отказать до выделения массива/цикла.
Private Function Audit03CaptureInputError(ByVal source As Object, ByVal operation As String, _
        ByRef description As String) As Long
    On Error GoTo ExpectedError
    Dim settings As CSystemSettingsReader, profiles As CCalculationProfileCatalog
    Dim mesh As CFiberMeshBuilder, geometry As CGeometryRoundedRectangle
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
                Set registry = New CSectionTypeRegistry
                integerValue = registry.MeshBoundarySubdivisions(settings)
            End If
        Case "Profiles"
            Set profiles = New CCalculationProfileCatalog
            profiles.LoadFromRange source
        Case Else
            Set mesh = New CFiberMeshBuilder
            Set geometry = New CGeometryRoundedRectangle
            geometry.Initialize 100#, 100#, 0#, 0#, 0#, 0#
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
    Audit03CaptureInputError = Err.Number
    description = Err.Description
End Function

' ДЛЯ ТЕСТОВ: проверяет настоящий reader профилей независимо от общего
' SettingsReader. Неверный переключатель не должен молча отключать расчет.
Public Function RunAudit03ProfileInputTests() As String
    On Error GoTo Failed
    Dim stats As TUiTestStats
    TestAudit03ProfileInputContracts stats
    AppendLine stats, "TOTAL_AUDIT03_PROFILE_INPUT: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunAudit03ProfileInputTests = stats.Report
    Exit Function
Failed:
    RunAudit03ProfileInputTests = stats.Report & "RUNTIME ERROR: " & CStr(Err.Number) & "; " & Err.Description
End Function

' ДЛЯ ТЕСТОВ: копирует фактическую таблицу профилей на временный лист.
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
            AssertTrue stats, "audit03.profile.boolean." & CStr(key) & "." & CStr(value), _
                Audit03ProfileFlag(profile, CStr(key)) = (value = "Yes" Or value = "True" Or value = "1" Or value = "Да")
        Next value
        caseIndex = 0
        For Each value In Array("Maybe", "TODO", "", CVErr(xlErrNA))
            caseIndex = caseIndex + 1
            target.Value2 = data
            target.Cells.Item(keyRow, profileColumn).Value2 = value
            description = vbNullString
            actualCode = Audit03CaptureInputError(target, "Profiles", description)
            AssertTrue stats, "audit03.profile.invalid." & CStr(key) & "." & CStr(caseIndex), actualCode = vbObjectError + 3988
            AssertTrue stats, "audit03.profile.reason." & CStr(key) & "." & CStr(caseIndex), _
                InStr(1, description, CStr(key), vbBinaryCompare) > 0 And InStr(1, description, "PR1", vbBinaryCompare) > 0
            AppendLine stats, "PROFILE_INPUT: key=" & CStr(key) & "; case=" & CStr(caseIndex) & _
                "; error=" & CStr(actualCode) & "; reason=" & description
        Next value
        target.Value2 = data
        target.Cells.Item(keyRow, 2).Value2 = "__Removed." & CStr(key)
        description = vbNullString
        actualCode = Audit03CaptureInputError(target, "Profiles", description)
        AssertTrue stats, "audit03.profile.missing." & CStr(key), actualCode = vbObjectError + 3988
        AssertTrue stats, "audit03.profile.missingReason." & CStr(key), InStr(1, description, CStr(key), vbBinaryCompare) > 0
    Next key
    GoTo CleanUp
Failed:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: audit03.profile.runtime; " & CStr(Err.Number) & "; " & Err.Description
CleanUp:
    On Error Resume Next
    Application.DisplayAlerts = False
    If Not sheet Is Nothing Then sheet.Delete
    Application.DisplayAlerts = oldAlerts
    On Error GoTo 0
End Sub

' ДЛЯ ТЕСТОВ: читает конечный флаг профиля, а не повторяет строковый parser.
' Поэтому тест выявляет потерю настройки между ячейкой и CCalculationProfile.
Private Function Audit03ProfileFlag(ByVal profile As CCalculationProfile, ByVal key As String) As Boolean
    Select Case key
        Case "Calculation.Strength.DirectState": Audit03ProfileFlag = profile.StrengthDirectStateEnabled
        Case "Calculation.Strength.Capacity": Audit03ProfileFlag = profile.StrengthCapacityEnabled
        Case "Calculation.Crack.Width": Audit03ProfileFlag = profile.CrackWidthEnabled
        Case "Calculation.Stability.Enabled": Audit03ProfileFlag = profile.StabilityEnabled
        Case Else: Err.Raise vbObjectError + 4499, "Audit03ProfileFlag", "Неизвестный тестовый переключатель профиля."
    End Select
End Function

' ДЛЯ ТЕСТОВ: отдельный entrypoint для одинакового контрпримера до и после
' исправления. Настройки передаются через настоящий Range и ApplySettings.
Public Function RunAudit03NumericSettingsInputTests() As String
    Dim stats As TUiTestStats
    TestAudit03NumericSettingsInputContracts stats
    AppendLine stats, "TOTAL_AUDIT03_NUMERIC_SETTINGS_INPUT: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunAudit03NumericSettingsInputTests = stats.Report
End Function

' ДЛЯ ТЕСТОВ: явная пустая/TODO численная настройка не должна возвращать
' default. Отсутствующий optional key и пустой optional диаметр имеют другой
' контракт; проверяем их отдельно без изменения пользовательского Config.
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
            Set inputRange = sheet.Range("A1:C3")
        End If
        For Each keyValue In Audit03NumericSettingsKeys(operation)
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
                actualCode = Audit03CaptureInputError(inputRange, operation, description)
                AssertTrue stats, "audit03.numeric." & operation & "." & key & ".error", actualCode = expectedCode
                AssertTrue stats, "audit03.numeric." & operation & "." & key & ".reason", InStr(1, description, key, vbBinaryCompare) > 0
                AppendLine stats, "NUMERIC_INPUT: operation=" & operation & "|key=" & key & _
                    "|value=" & CStr(value) & "|error=" & CStr(actualCode) & "|reason=" & description
            Next value
            sheet.Range("B2").Value2 = 80#
            actualCode = Audit03CaptureInputError(inputRange, operation, description)
            AssertTrue stats, "audit03.numeric." & operation & "." & key & ".parsable", actualCode = 0
            sheet.Range("A2").Value2 = "Audit03.Unrelated"
            actualCode = Audit03CaptureInputError(inputRange, operation, description)
            AssertTrue stats, "audit03.numeric." & operation & ".optionalAbsent", actualCode = 0
        Next keyValue
    Next operationValue
    sheet.Range("A2").Value2 = "Rebar.Diameter2": sheet.Range("B2").Value2 = vbNullString
    Set settings = New CSystemSettingsReader
    settings.LoadFromRange sheet.Range("A1:C2")
    AssertClose stats, "audit03.numeric.optionalDiameterBlank", settings.GetDouble("Rebar.Diameter2"), 0#, 0#
    AssertClose stats, "audit03.numeric.optionalKeyDefault", settings.GetDouble("Audit03.Absent", 17#), 17#, 0#
    TestAudit03BatchNumericSettingsInput stats
    GoTo CleanUp
Failed:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: audit03.numeric.runtime; " & CStr(Err.Number) & "; " & Err.Description
CleanUp:
    On Error Resume Next
    Application.DisplayAlerts = False
    If Not sheet Is Nothing Then sheet.Delete
    Application.DisplayAlerts = oldAlerts
    On Error GoTo 0
End Sub

' ДЛЯ ТЕСТОВ: проверяет сохранение отрицательного знака после реального
' преобразования tf/tf*m/кривизны. Capacity, Formation и Batch проходят свои
' API до solver-а; ошибка не доказывается только чтением settings/getter-а.
Public Function RunAudit03UnitSignConsumerTests() As String
    Dim stats As TUiTestStats
    AppendLine stats, "RUN: Audit03 unit-sign consumers"
    TestAudit03UnitSignConsumers stats
    AppendLine stats, "TOTAL_AUDIT03_UNIT_SIGN_CONSUMERS: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunAudit03UnitSignConsumerTests = stats.Report
End Function

' ДЛЯ ТЕСТОВ: меняет четыре поля на контролируемой копии Config и гарантированно
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
    SetSystemSetting "Units.Force.Input", "tf"
    SetSystemSetting "Units.Moment.Input", "tf*m"
    SetSystemSetting "Units.Curvature.Input", "1/m"
    SetSystemSetting "Sign.N.User", "Compression"
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
    For Each key In Array("Solver.ToleranceN", "Solver.ToleranceMx", "Solver.ToleranceMy", "Solver.MaxDeltaKappa")
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
        AppendLine stats, "RUN: " & prefix & ".capacity"
        Set cap = New CCapacitySolver
        cap.ApplySettings settings, units
        cap.SolveByLoadPathMultiplier section, concrete, steel, 0#, load.N, 0#, load.InternalMx, 0#, load.InternalMy
        AssertTrue stats, prefix & ".capacityTyped", cap.FailureCode = sfcInvalidConfiguration
        AssertTrue stats, prefix & ".capacityHasDiagnosticSolver", Not cap.LastSolver Is Nothing
        If Not cap.LastSolver Is Nothing Then
            AssertTrue stats, prefix & ".capacityNoIterations", cap.LastSolver.Iterations = 0
        End If
        AssertTrue stats, prefix & ".capacityNoRetries", cap.RetryCount = 0
        AssertTrue stats, prefix & ".capacityReason", Len(cap.StopReason) > 0
        AppendLine stats, "RUN: " & prefix & ".formation"
        Set formation = New CCrackFormationCalculator
        formation.ApplySettings settings, units
        formation.CrackFormationPath = "LambdaNMxy"
        formation.CrackFormationSolutionStrategy = "LoadMultiplier"
        Set formed = formation.CheckFormation(section, materials, currentSpec, formationSpec, _
            load.N, load.InternalMx, load.InternalMy, load, 0#, 0#)
        AssertTrue stats, prefix & ".formationTyped", formed.ResultMeta.InternalStatus = rsInvalidConfiguration
        AssertTrue stats, prefix & ".formationNoPoint", Not formed.HasLimitPoint
        AssertTrue stats, prefix & ".formationReason", Len(formed.ResultMeta.ResultComment) > 0
        AppendLine stats, "RUN: " & prefix & ".batch"
        Set batch = BuildUiBatch()
        AppendLine stats, "RUN: " & prefix & ".batch.built; " & Audit02ExcelMemory()
        batch.ApplySettings settings, units
        AppendLine stats, "RUN: " & prefix & ".batch.settingsApplied; " & Audit02ExcelMemory()
        batch.AddCombination "UNIT_SIGN", load.N, load.InternalMx, load.InternalMy, "PR1", vbNullString
        AppendLine stats, "RUN: " & prefix & ".batch.execute; " & Audit02ExcelMemory()
        batch.Execute
        AppendLine stats, "RUN: " & prefix & ".batch.executed; " & Audit02ExcelMemory()
        Set result = batch.ResultAt(1)
        AssertTrue stats, prefix & ".batchTyped", result.DirectStateMeta.InternalStatus = rsInvalidConfiguration
        AssertTextEquals stats, prefix & ".batchExternal", policy.ExternalStatus(result.DirectStateMeta), "InputErr"
        AssertTrue stats, prefix & ".batchReason", Len(result.DirectStateMeta.ResultComment) > 0
        AppendLine stats, "UNIT_SIGN_CONSUMER: key=" & CStr(key) & "|capacity=" & cap.StopReason & _
            "|formation=" & formed.ResultMeta.ResultComment & "|batch=" & result.DirectStateMeta.ResultComment
    Next key
    GoTo CleanUp
Failed:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: audit03.unitSign.runtime; " & CStr(Err.Number) & "; " & Err.Description
CleanUp:
    On Error Resume Next
    If Not target Is Nothing Then target.Formula = savedFormula
    If Not unitRange Is Nothing Then unitRange.Formula = savedUnits
    If Not signRange Is Nothing Then signRange.Formula = savedSigns
    On Error GoTo 0
End Sub

' ДЛЯ ТЕСТОВ: полный список численных ключей каждого ApplySettings.
' Общие Solver-поля и Capacity overrides проверяются по реальным маршрутам,
' а не только прямым вызовом GetDouble. Инженерные диапазоны проверяются отдельно.
Private Function Audit03NumericSettingsKeys(ByVal operation As String) As Variant
    If operation = "Capacity" Then
        Audit03NumericSettingsKeys = Array("Capacity.InitialLambda", "Capacity.MaxLambda", _
            "Capacity.ToleranceLambda", "Capacity.ToleranceStrain", "Capacity.MaxRetries", _
            "Capacity.BaseLoadSteps", "Capacity.SolverMaxIterations", "Solver.ToleranceN", _
            "Solver.ToleranceMx", "Solver.ToleranceMy", "Solver.DampingInitial", _
            "Solver.MinLineSearchAlpha", "Solver.MaxDeltaEpsilon0", "Solver.MaxDeltaKappa", _
            "Solver.SecantMaxRestarts", "Solver.SecantMinStepNorm")
    ElseIf operation = "Batch" Then
        Audit03NumericSettingsKeys = Array("Solver.MaxIterations", "Solver.LoadSteps", _
            "Solver.ToleranceN", "Solver.ToleranceMx", "Solver.ToleranceMy", "Solver.DampingInitial", _
            "Solver.MinLineSearchAlpha", "Solver.MaxDeltaEpsilon0", "Solver.MaxDeltaKappa", _
            "Solver.SecantMaxRestarts", "Solver.SecantMinStepNorm", "Capacity.InitialLambda", _
            "Capacity.MaxLambda", "Capacity.ToleranceLambda", "Capacity.ToleranceStrain", _
            "Capacity.MaxRetries", "Capacity.BaseLoadSteps", "Capacity.SolverMaxIterations")
    Else
        Audit03NumericSettingsKeys = Array("Solver.MaxIterations", "Solver.LoadSteps", _
            "Solver.ToleranceN", "Solver.ToleranceMx", "Solver.ToleranceMy", "Solver.DampingInitial", _
            "Solver.MinLineSearchAlpha", "Solver.MaxDeltaEpsilon0", "Solver.MaxDeltaKappa", _
            "Solver.SecantMaxRestarts", "Solver.SecantMinStepNorm")
    End If
End Function

' ДЛЯ ТЕСТОВ: Batch сохраняет ошибку ApplySettings, а не выбрасывает ее наружу.
' Используем полный Config с обязательными параметрами устойчивости; проверяем
' InputErr, отсутствие solve и возможность следующего запуска после восстановления.
Private Sub TestAudit03BatchNumericSettingsInput(ByRef stats As TUiTestStats)
    Dim target As Object, savedFormula As Variant, batch As CBatchSectionCalculator
    Dim settings As CSystemSettingsReader, key As Variant, value As Variant, row As Long
    Dim valueCell As Object
    On Error GoTo Failed
    Set target = ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange
    savedFormula = target.Formula
    Set batch = BuildUiBatch()
    batch.AddCombination "NUMERIC_INPUT", 0#, 0#, 0#, "PR1", "Проверка ввода численных настроек"
    For Each key In Audit03NumericSettingsKeys("Batch")
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
            Audit03CheckBatchNumericInput stats, batch, CStr(key), value
        Next value
    Next key
    target.Formula = savedFormula
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook
    batch.ApplySettings settings
    batch.Execute
    AssertTrue stats, "audit03.numeric.Batch.restoredStatus", batch.ResultAt(1).Status <> "InputErr"
    AssertTrue stats, "audit03.numeric.Batch.restoredSolve", batch.ResultAt(1).DirectStateMeta.Calculated
    GoTo CleanUp
Failed:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: audit03.numeric.Batch.runtime; " & CStr(Err.Number) & "; " & Err.Description
CleanUp:
    On Error Resume Next
    If Not target Is Nothing Then target.Formula = savedFormula
    On Error GoTo 0
End Sub

' ДЛЯ ТЕСТОВ: различает раннюю ошибку Excel-ячейки у reader-а и сохраненный
' отказ Batch. Ни один неверный численный ввод не должен запускать equilibrium.
Private Sub Audit03CheckBatchNumericInput(ByRef stats As TUiTestStats, _
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
    AssertTextEquals stats, prefix & ".status", result.Status, "InputErr"
    AssertTrue stats, prefix & ".typed", result.OverallMeta.InternalStatus = rsInvalidInput
    AssertTrue stats, prefix & ".reason", InStr(1, result.OverallMeta.ResultComment, key, vbBinaryCompare) > 0
    AssertTrue stats, prefix & ".noDirectSolve", Not result.DirectStateMeta.Calculated
    AssertTrue stats, prefix & ".noCapacitySearch", Not result.CapacityMeta.Calculated
    AppendLine stats, "NUMERIC_INPUT: operation=Batch|key=" & key & "|value=" & CStr(value) & _
        "|status=" & result.Status & "|reason=" & result.OverallMeta.ResultComment
    Exit Sub
ReaderFailed:
    errorNumber = Err.Number: description = Err.Description
    AssertTrue stats, prefix & ".readerError", VarType(value) = vbError And errorNumber = vbObjectError + 4309
    AssertTrue stats, prefix & ".readerReason", InStr(1, description, key, vbBinaryCompare) > 0
    AppendLine stats, "NUMERIC_INPUT: operation=BatchReader|key=" & key & "|value=" & CStr(value) & _
        "|error=" & CStr(errorNumber) & "|reason=" & description
    Exit Sub
Failed:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: " & prefix & ".runtime; " & CStr(Err.Number) & "; " & Err.Description
End Sub

' ДЛЯ ТЕСТОВ: адресный gate независимых настроек дополнительных рядов RectSet.
' Проверяет чтение, списки, влияние на одну грань и ошибки активного ввода.
Public Function RunRectSetIndependentSelectorTests() As String
    Dim stats As TUiTestStats
    TestRectSetIndependentSelectors stats
    TestRectSetIndependentSelectorLayout stats
    TestRectSetIndependentSelectorEffects stats
    TestRectSetIndependentThirdRows stats
    AppendLine stats, "TOTAL_RECTSET_INDEPENDENT: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunRectSetIndependentSelectorTests = stats.Report
End Function

' ДЛЯ ТЕСТОВ: разные значения противоположных граней сохраняются как отдельные
' ключи. Никакой выбор первой строки не подменяет значение второй строки.
Private Sub TestRectSetIndependentSelectors(ByRef stats As TUiTestStats)
    Dim sheet As Object, target As Object, source As Object, settings As CSystemSettingsReader
    Dim face As Variant, column As Variant, pairIndex As Long, topRow As Long, side As Long
    Dim firstValue As String, secondValue As String, tail As String, key As String
    Dim savedAlerts As Boolean
    On Error GoTo Failed
    savedAlerts = Application.DisplayAlerts
    Set source = ThisWorkbook.Names.Item("rngRectSetGeometry").RefersToRange
    Set sheet = ThisWorkbook.Worksheets.Add
    Set target = sheet.Range("A1").Resize(source.Rows.Count, source.Columns.Count)
    target.Value2 = source.Value2
    For Each face In Array("H1", "B1", "H2", "B2")
        topRow = 21 + pairIndex * 2
        For Each column In Array(3, 4, 6, 7)
            tail = RectSetSelectorTail(CLng(column))
            If CLng(column) = 3 Or CLng(column) = 6 Then
                firstValue = "Stacked": secondValue = "SideBySide"
            Else
                firstValue = "EachBar": secondValue = "EverySecondBar"
            End If
            target.Cells(topRow, CLng(column)).Value2 = firstValue
            target.Cells(topRow + 1, CLng(column)).Value2 = secondValue
            Set settings = New CSystemSettingsReader
            settings.LoadFromRange target
            key = "RectSet." & CStr(face) & "." & tail
            AssertTextEquals stats, key & ".first", settings.GetString(key & "_1"), firstValue
            AssertTextEquals stats, key & ".second", settings.GetString(key & "_2"), secondValue
            For side = 1 To 2
                target.Cells(topRow + side - 1, CLng(column)).ClearContents
                settings.LoadFromRange target
                AssertTextEquals stats, key & ".empty" & CStr(side), settings.GetString(key & "_" & CStr(side)), vbNullString
                target.Cells(topRow, CLng(column)).Value2 = firstValue
                target.Cells(topRow + 1, CLng(column)).Value2 = secondValue
            Next side
        Next column
        pairIndex = pairIndex + 1
    Next face
    GoTo Restore
Failed:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: rectset.independent.reader; " & CStr(Err.Number) & "; " & Err.Description
Restore:
    On Error Resume Next
    Application.DisplayAlerts = False
    If Not sheet Is Nothing Then sheet.Delete
    Application.DisplayAlerts = savedAlerts
    On Error GoTo 0
End Sub

' ДЛЯ ТЕСТОВ: все 32 селектора фактического Config не объединены, имеют
' вертикальный список выбора и читаются по адресу своей физической грани.
Private Sub TestRectSetIndependentSelectorLayout(ByRef stats As TUiTestStats)
    Dim target As Object, cell As Object, settings As CSystemSettingsReader
    Dim face As Variant, column As Variant, pairIndex As Long, topRow As Long, side As Long
    Dim key As String, prefix As String
    On Error GoTo Failed
    Set target = ThisWorkbook.Names.Item("rngRectSetGeometry").RefersToRange
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook
    For Each face In Array("H1", "B1", "H2", "B2")
        topRow = 21 + pairIndex * 2
        For Each column In Array(3, 4, 6, 7)
            For side = 1 To 2
                key = "RectSet." & CStr(face) & "." & RectSetSelectorTail(CLng(column)) & "_" & CStr(side)
                prefix = "rectset.independent.layout." & key
                Set cell = target.Cells(topRow + side - 1, CLng(column))
                AssertTrue stats, prefix & ".unmerged", Not cell.MergeCells
                AssertTrue stats, prefix & ".dropdown", cell.Validation.Type = 3 And cell.Validation.InCellDropdown
                AssertTrue stats, prefix & ".rangeList", Left$(cell.Validation.Formula1, 1) = "="
                AssertTrue stats, prefix & ".centered", cell.HorizontalAlignment = xlCenter And cell.VerticalAlignment = xlCenter
                AssertTextEquals stats, prefix & ".reader", settings.GetString(key), CStr(cell.Value2)
            Next side
        Next column
        pairIndex = pairIndex + 1
    Next face
    Exit Sub
Failed:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: rectset.independent.layout; " & CStr(Err.Number) & "; " & Err.Description
End Sub

' ДЛЯ ТЕСТОВ: меняем каждый из 32 селекторов через Range -> reader -> builder.
' Меняются только три дополнительных стержня выбранной стороны; шесть
' стержней противоположной стороны сохраняют координаты и диаметры.
' Ошибочный/пустой активный ввод дает адресную ошибку, отключенный не читается.
Private Sub TestRectSetIndependentSelectorEffects(ByRef stats As TUiTestStats)
    Dim source As Object, sheet As Object, target As Object, settings As CSystemSettingsReader
    Dim units As CUnitSystem, builder As CRectSetRebarLayoutBuilder
    Dim baseline As CRebarLayout, changed As CRebarLayout, inactive As CRebarLayout
    Dim face As Variant, column As Variant, pairIndex As Long, topRow As Long, mainRow As Long
    Dim row As Long, i As Long, side As Long, selectedRow As Long, diameterColumn As Long
    Dim prefix As String, alternative As String, original As String, changedBars As Long
    Dim oppositeStart As Long, number As Long, reason As String, savedAlerts As Boolean
    Dim invalid As Variant
    On Error GoTo Failed
    savedAlerts = Application.DisplayAlerts
    Set source = ThisWorkbook.Names.Item("rngRectSetGeometry").RefersToRange
    Set sheet = ThisWorkbook.Worksheets.Add
    Set target = sheet.Range("A1").Resize(source.Rows.Count, source.Columns.Count)
    Set units = New CUnitSystem: units.InitializeDefaults
    Set builder = New CRectSetRebarLayoutBuilder
    For Each face In Array("H1", "B1", "H2", "B2")
        topRow = 21 + pairIndex * 2: mainRow = 11 + pairIndex * 2
        For Each column In Array(3, 4, 6, 7)
            For side = 1 To 2
                target.Value2 = source.Value2
                target.Cells(3, 2).Value2 = "LSection": target.Cells(4, 2).Value2 = 0#
                target.Cells(8, 1).Value2 = 300#: target.Cells(8, 2).Value2 = 200#
                target.Cells(8, 3).Value2 = 200#: target.Cells(8, 4).Value2 = 500#
                For row = 11 To 18
                    target.Cells(row, 3).Value2 = 0#: target.Cells(row, 4).Value2 = 0
                Next row
                For row = 21 To 28
                    target.Cells(row, 2).Value2 = 0#: target.Cells(row, 5).Value2 = 0#
                    target.Cells(row, 3).Value2 = "Stacked": target.Cells(row, 6).Value2 = "Stacked"
                    target.Cells(row, 4).Value2 = "EachBar": target.Cells(row, 7).Value2 = "EachBar"
                Next row
                For row = mainRow To mainRow + 1
                    target.Cells(row, 2).Value2 = 20#
                    target.Cells(row, 3).Value2 = 10#: target.Cells(row, 4).Value2 = 3
                    target.Cells(row, 5).Value2 = 30#: target.Cells(row, 6).Value2 = 30#
                Next row
                If CLng(column) <= 4 Then diameterColumn = 2 Else diameterColumn = 5
                target.Cells(topRow, diameterColumn).Value2 = 8#
                target.Cells(topRow + 1, diameterColumn).Value2 = 8#
                selectedRow = topRow + side - 1
                If CLng(column) = 3 Or CLng(column) = 6 Then
                    original = "Stacked": alternative = "SideBySide"
                Else
                    original = "EachBar": alternative = "EverySecondBar"
                End If
                prefix = "rectset.independent.effect." & CStr(face) & "." & CStr(column) & "." & CStr(side)
                Set settings = New CSystemSettingsReader: settings.LoadFromRange target
                Set baseline = builder.BuildFromSettings(settings, units)
                AssertTrue stats, prefix & ".baseline", baseline.Count = 12
                target.Cells(selectedRow, CLng(column)).Value2 = alternative
                settings.LoadFromRange target
                Set changed = builder.BuildFromSettings(settings, units)
                If CLng(column) = 3 Or CLng(column) = 6 Then
                    changedBars = 0
                    For i = 1 To baseline.Count
                        If Abs(baseline.X(i) - changed.X(i)) + Abs(baseline.Y(i) - changed.Y(i)) > 0.00000001 Then changedBars = changedBars + 1
                    Next i
                    AssertTrue stats, prefix & ".oneSideMoved", changedBars = 3 And changed.Count = 12
                Else
                    AssertTrue stats, prefix & ".oneSideBound", changed.Count = 11
                End If
                If side = 1 Then oppositeStart = 7 Else oppositeStart = 1
                For i = oppositeStart To oppositeStart + 5
                    AssertTrue stats, prefix & ".opposite" & CStr(i), RectSetLayoutContainsBar(changed, baseline, i)
                Next i
                For Each invalid In Array(vbNullString, "InvalidChoice", CVErr(xlErrDiv0))
                    target.Cells(selectedRow, CLng(column)).Value2 = invalid
                    settings.LoadFromRange target
                    On Error Resume Next
                    Set changed = builder.BuildFromSettings(settings, units)
                    number = Err.Number: reason = Err.Description: Err.Clear
                    On Error GoTo Failed
                    AssertTrue stats, prefix & ".invalidRejected." & TypeName(invalid), number <> 0
                    AssertTrue stats, prefix & ".invalidAddress." & TypeName(invalid), _
                        InStr(1, reason, "ячейка " & target.Cells(selectedRow, CLng(column)).Address(False, False), vbBinaryCompare) > 0
                Next invalid
                target.Cells(selectedRow, diameterColumn).Value2 = 0#
                settings.LoadFromRange target
                Set inactive = builder.BuildFromSettings(settings, units)
                AssertTrue stats, prefix & ".inactiveIgnored", inactive.Count = 9
                target.Cells(selectedRow, CLng(column)).Value2 = original
                settings.LoadFromRange target
                Set changed = builder.BuildFromSettings(settings, units)
                AssertTrue stats, prefix & ".inactiveSameCount", changed.Count = inactive.Count
                For i = 1 To inactive.Count
                    AssertTrue stats, prefix & ".inactiveSame" & CStr(i), RectSetLayoutContainsBar(changed, inactive, i)
                Next i
            Next side
        Next column
        pairIndex = pairIndex + 1
    Next face
    GoTo Restore
Failed:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: rectset.independent.effect; " & CStr(Err.Number) & "; " & Err.Description
Restore:
    On Error Resume Next
    Application.DisplayAlerts = False
    If Not sheet Is Nothing Then sheet.Delete
    Application.DisplayAlerts = savedAlerts
    On Error GoTo 0
End Sub

' ДЛЯ ТЕСТОВ: третий ряд учитывает направление и диаметр второго ряда только
' собственной стороны. Совпадение направлений справа не вызывает перескок
' слева, где направления различаются; последующее изменение слева изолировано.
Private Sub TestRectSetIndependentThirdRows(ByRef stats As TUiTestStats)
    Dim builder As CRectSetRebarLayoutBuilder, layout As CRebarLayout, changed As CRebarLayout
    Dim face As Variant, emptyFace As Variant
    On Error GoTo Failed
    Set builder = New CRectSetRebarLayoutBuilder
    emptyFace = Array(50#, 50#, 20#, 20#, 0, 0, 100#, 100#, 100#, 100#)
    face = Array(50#, 50#, 20#, 20#, 1, 1, 100#, 100#, 100#, 100#, 18#, 22#, 28#, 34#, _
        "SideBySide", "Stacked", "EachBar", "EachBar", "Stacked", "Stacked", "EachBar", "EachBar")
    Set layout = builder.Build(200#, 300#, 0#, 0#, 0#, 0#, face, emptyFace, emptyFace, emptyFace, "Rebar", 0#, "Rectangle")
    AssertTrue stats, "rectset.thirdRows.count", layout.Count = 6
    AssertTrue stats, "rectset.thirdRows.leftNoSkip", Abs(layout.X(3) - layout.X(1) - 24#) < 0.00000001
    AssertTrue stats, "rectset.thirdRows.rightSkip", Abs(layout.X(4) - layout.X(6) - 49#) < 0.00000001
    face(15) = "SideBySide"
    Set changed = builder.Build(200#, 300#, 0#, 0#, 0#, 0#, face, emptyFace, emptyFace, emptyFace, "Rebar", 0#, "Rectangle")
    AssertTrue stats, "rectset.thirdRows.leftSkip", Abs(Abs(changed.Y(3) - changed.Y(1)) - 42#) < 0.00000001
    AssertTrue stats, "rectset.thirdRows.rightUnchanged", RectSetLayoutContainsBar(changed, layout, 6)
    Exit Sub
Failed:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: rectset.thirdRows; " & CStr(Err.Number) & "; " & Err.Description
End Sub

' ДЛЯ ТЕСТОВ: связывает столбец пользовательской таблицы с независимым ключом.
Private Function RectSetSelectorTail(ByVal column As Long) As String
    Select Case column
        Case 3: RectSetSelectorTail = "loc_2row"
        Case 4: RectSetSelectorTail = "bind_2row"
        Case 6: RectSetSelectorTail = "loc_3row"
        Case 7: RectSetSelectorTail = "bind_3row"
    End Select
End Function

' ДЛЯ ТЕСТОВ: ищет исходный стержень в новой раскладке без зависимости от
' перенумерации после изменения привязки дополнительных рядов.
Private Function RectSetLayoutContainsBar(ByVal actual As CRebarLayout, ByVal original As CRebarLayout, ByVal index As Long) As Boolean
    Dim i As Long
    For i = 1 To actual.Count
        If Abs(actual.X(i) - original.X(index)) + Abs(actual.Y(i) - original.Y(index)) + _
                Abs(actual.Diameter(i) - original.Diameter(index)) < 0.00000001 Then
            RectSetLayoutContainsBar = True
            Exit Function
        End If
    Next i
End Function

' ДЛЯ ТЕСТОВ: запускает поведенческую проверку всех пятнадцати пользовательских
' селекторов единиц/знаков через фактический Config, включая ошибку и recovery.
Public Function RunAudit03UnitSignChoiceTests() As String
    Dim stats As TUiTestStats
    TestAudit03UnitSignChoices stats
    AppendLine stats, "TOTAL_AUDIT03_UNIT_SIGN_CHOICES: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunAudit03UnitSignChoiceTests = stats.Report
End Function

' ДЛЯ ТЕСТОВ: сверяет каждый вариант с независимым коэффициентом пересчета,
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
            If Not UnitSettingAddress(key, rowIndex, columnIndex) Then Err.Raise vbObjectError + 4499, , key
            Set cell = unitRange.Cells(rowIndex, columnIndex)
            oldValue = cell.Formula
            AssertTrue stats, "audit03.unitChoice." & key & ".validation", ValidationCellHasOptions(cell, choices(q))
            For optionIndex = 0 To UBound(choices(q))
                cell.Value2 = choices(q)(optionIndex)
                Set settings = New CSystemSettingsReader
                settings.LoadFromWorkbook ThisWorkbook
                Set units = New CUnitSystem
                units.LoadFromSettings settings
                factor = CDbl(factors(q)(optionIndex))
                prefix = "audit03.unitChoice." & key & "." & CStr(choices(q)(optionIndex))
                For Each value In Array(-123.456, 0#, 123.456)
                    actual = Audit03UnitConversion(units, q, CStr(side), CDbl(value))
                    If CStr(side) = "Input" Then expected = CDbl(value) * factor Else expected = CDbl(value) / factor
                    AssertClose stats, prefix & ".value" & CStr(value), actual, expected, _
                        0.000000000001 * (1# + Abs(expected))
                Next value
                If q = 0 Then
                    If CStr(side) = "Input" Then
                        actual = units.InputFourthPowerLengthToInternal(2#): expected = 2# * factor ^ 4
                    Else
                        actual = units.InternalFourthPowerLengthToOutput(2#): expected = 2# / factor ^ 4
                        AssertClose stats, prefix & ".reverseL4", units.OutputFourthPowerLengthToInternal(2#), _
                            2# * factor ^ 4, 0.000000000001 * (1# + 2# * factor ^ 4)
                    End If
                    AssertClose stats, prefix & ".L4", actual, expected, 0.000000000001 * (1# + Abs(expected))
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
                AssertTrue stats, "audit03.unitChoice." & key & ".invalid." & CStr(VarType(invalid)) & "." & CStr(errorNumber), _
                    errorNumber = vbObjectError + 4530 + q Or errorNumber = vbObjectError + 4309
                AssertTrue stats, "audit03.unitChoice." & key & ".invalid.reason", _
                    InStr(1, errorDescription, "Units." & CStr(quantities(q)), vbTextCompare) > 0
                AssertTrue stats, "audit03.unitChoice." & key & ".invalid.address", _
                    InStr(1, errorDescription, cell.Address(False, False), vbTextCompare) > 0
                AssertTrue stats, "audit03.unitChoice." & key & ".invalid.action", _
                    InStr(1, errorDescription, "Выберите", vbTextCompare) > 0 Or InStr(1, errorDescription, "Исправьте", vbTextCompare) > 0
            Next invalid
            cell.Formula = oldValue
            Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
            Set units = New CUnitSystem: units.LoadFromSettings settings
            AssertTrue stats, "audit03.unitChoice." & key & ".recovery", Len(settings.GetRawString(key)) > 0
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
        AssertTrue stats, "audit03.signChoice." & key & ".validation", ValidationCellHasOptions(cell, choices(signIndex))
        For optionIndex = 0 To 1
            cell.Value2 = choices(signIndex)(optionIndex)
            Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
            Set units = New CUnitSystem: units.LoadFromSettings settings
            factor = 1# - 2# * optionIndex
            If signIndex = 0 Then
                expected = 123# * 9806.65 * factor
                actual = units.InputForceToInternal(123#)
                AssertClose stats, "audit03.signChoice.N.output." & CStr(optionIndex), _
                    units.InternalForceToOutput(9806.65), factor, 0.000000000001
            ElseIf signIndex = 1 Then
                expected = 123# * 9806650# * factor
                actual = units.InputMomentMxToInternal(123#)
                AssertClose stats, "audit03.signChoice.Mx.output." & CStr(optionIndex), _
                    units.InternalMomentMxToOutput(9806650#), factor, 0.000000000001
            Else
                expected = 123# * 9806650# * factor
                actual = units.InputMomentMyToInternal(123#)
                AssertClose stats, "audit03.signChoice.My.output." & CStr(optionIndex), _
                    units.InternalMomentMyToOutput(9806650#), factor, 0.000000000001
            End If
            AssertClose stats, "audit03.signChoice." & key & ".input." & CStr(optionIndex), actual, expected, 0.00001
            AssertClose stats, "audit03.signChoice." & key & ".stressNotFlipped", units.InputStressToInternal(-12#), -12#, 0.000000000001
            AssertClose stats, "audit03.signChoice." & key & ".magnitudeNotFlipped", units.InternalMomentMagnitudeToOutput(9806650#), 1#, 0.000000000001
            AssertClose stats, "audit03.signChoice." & key & ".thresholdNotFlipped", units.InputMomentPerLengthToInternal(1#), 9806650#, 0.000001
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
            AssertTrue stats, "audit03.signChoice." & key & ".invalid." & CStr(errorNumber), _
                errorNumber = vbObjectError + 4536 + signIndex Or errorNumber = vbObjectError + 4309
            AssertTrue stats, "audit03.signChoice." & key & ".invalid.reason", InStr(1, errorDescription, key, vbTextCompare) > 0
            AssertTrue stats, "audit03.signChoice." & key & ".invalid.address", _
                InStr(1, errorDescription, cell.Address(False, False), vbTextCompare) > 0
            AssertTrue stats, "audit03.signChoice." & key & ".invalid.action", _
                InStr(1, errorDescription, "Выберите", vbTextCompare) > 0 Or InStr(1, errorDescription, "Исправьте", vbTextCompare) > 0
        Next invalid
        cell.Formula = oldValue
        Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
        Set units = New CUnitSystem: units.LoadFromSettings settings
        AssertTrue stats, "audit03.signChoice." & key & ".recovery", Len(settings.GetRawString(key)) > 0
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
            AssertTrue stats, "audit03.unitSignMissing." & key & ".inputError", errorNumber = vbObjectError + 4313
            AssertTrue stats, "audit03.unitSignMissing." & key & ".reason", InStr(1, errorDescription, key, vbTextCompare) > 0
        Next invalid
        labelRange.Formula = oldValue
        Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
        Set units = New CUnitSystem: units.LoadFromSettings settings
        AssertTrue stats, "audit03.unitSignMissing." & key & ".recovery", settings.HasKey(key)
    Next q
    GoTo Restore
Failed:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: audit03.unitSignChoices.runtime; " & CStr(Err.Number) & "; " & Err.Description
Restore:
    On Error Resume Next
    If Not unitRange Is Nothing Then unitRange.Formula = savedUnits
    If Not signRange Is Nothing Then signRange.Formula = savedSigns
    On Error GoTo 0
End Sub

' ДЛЯ ТЕСТОВ: выбирает именно проверяемое преобразование. Независимые
' коэффициенты и ожидаемые значения задаются в тесте, не читаются из CUnitSystem.
Private Function Audit03UnitConversion(ByVal units As CUnitSystem, ByVal quantity As Long, _
        ByVal side As String, ByVal value As Double) As Double
    If side = "Input" Then
        Select Case quantity
            Case 0: Audit03UnitConversion = units.InputLengthToInternal(value)
            Case 1: Audit03UnitConversion = units.InputAreaToInternal(value)
            Case 2: Audit03UnitConversion = units.InputForceToInternal(value)
            Case 3: Audit03UnitConversion = units.InputMomentMxToInternal(value)
            Case 4: Audit03UnitConversion = units.InputStressToInternal(value)
            Case 5: Audit03UnitConversion = units.InputCurvatureToInternal(value)
        End Select
    Else
        Select Case quantity
            Case 0: Audit03UnitConversion = units.InternalLengthToOutput(value)
            Case 1: Audit03UnitConversion = units.InternalAreaToOutput(value)
            Case 2: Audit03UnitConversion = units.InternalForceToOutput(value)
            Case 3: Audit03UnitConversion = units.InternalMomentMxToOutput(value)
            Case 4: Audit03UnitConversion = units.InternalStressToOutput(value)
            Case 5: Audit03UnitConversion = units.InternalCurvatureToOutput(value)
        End Select
    End If
End Function

' ДЛЯ ТЕСТОВ: проверяет эквивалентные физические LC при всех восьми знаковых
' соглашениях и девяти парах единиц силы/момента, с фактической записью Results.
Public Function RunAudit03UnitSignEquivalenceTests() As String
    Dim stats As TUiTestStats
    TestAudit03UnitSignEquivalence stats
    AppendLine stats, "TOTAL_AUDIT03_UNIT_SIGN_EQUIVALENCE: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunAudit03UnitSignEquivalenceTests = stats.Report
End Function

' ДЛЯ ТЕСТОВ: один внутренний section используется во всех 72 вариантах.
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
    Dim section As CSectionModel, mesh As CFiberMeshBuilder, geom As CGeometryCircle, rebars As CRebarLayout
    Dim rebarBuilder As CCircleRebarLayoutBuilder, snapshot As CNDMResultsWriter, summary As CBatchResultWriter
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
    SetSystemSetting "Calculation.ZeroMomentPerDepth", "0"
    SetSystemSetting "Load.ReferenceOffsetX", "10": SetSystemSetting "Load.ReferenceOffsetY", "-7"
    SetSystemSetting "General.ExecutionReportEnabled", "No"
    SetSystemSetting "Solver.MaxDeltaKappa", "0.0001"
    SetProfileSetting "PR1", "Calculation.Strength.DirectState", "Yes"
    SetProfileSetting "PR1", "Calculation.Strength.Capacity", "No"
    SetProfileSetting "PR1", "Calculation.Crack.Width", "No"
    SetProfileSetting "PR1", "Calculation.Stability.Enabled", "No"
    SetProfileSetting "PR2", "Calculation.Strength.DirectState", "No"
    SetProfileSetting "PR2", "Calculation.Strength.Capacity", "No"
    SetProfileSetting "PR2", "Calculation.Crack.Width", "Yes"
    SetProfileSetting "PR2", "Calculation.Stability.Enabled", "No"
    Set geom = New CGeometryCircle: geom.InitializeByDiameter 300#
    Set mesh = New CFiberMeshBuilder: mesh.BuildMesh geom, 20#, 20#, 1
    Set rebarBuilder = New CCircleRebarLayoutBuilder
    Set rebars = rebarBuilder.Build(300#, 0#, 0#, 40#, 8, 20#, "Rebar")
    Set section = BuildGeneratedSectionModel(mesh, rebars, "Audit03UnitSign")
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
                AppendLine stats, "RUN: " & prefix & "; " & Audit02ExcelMemory()
                unitRange.Cells(4, 2).Value2 = forceNames(forceIndex)
                unitRange.Cells(5, 2).Value2 = momentNames(momentIndex)
                SetSystemSetting "Solver.ToleranceN", CStr(0.1 / CDbl(forceFactors(forceIndex)))
                SetSystemSetting "Solver.ToleranceMx", CStr(1# / CDbl(momentFactors(momentIndex)))
                SetSystemSetting "Solver.ToleranceMy", CStr(1# / CDbl(momentFactors(momentIndex)))
                ClearDataRows loads
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
                AssertTrue stats, prefix & ".count", batch.Count = 4
                For i = 0 To 3
                    AssertClose stats, prefix & ".N." & CStr(i), batch.N(i + 1), CDbl(nValues(i)), 0.000001
                    AssertClose stats, prefix & ".Mx." & CStr(i), batch.UserMx(i + 1), CDbl(mxValues(i)), 0.000001
                    AssertClose stats, prefix & ".My." & CStr(i), batch.UserMy(i + 1), CDbl(myValues(i)), 0.000001
                Next i
                batch.Execute
                If stopAfterCases > 0 Then AppendLine stats, "RUN: " & prefix & ".batch.done; " & Audit02ExcelMemory()
                solveCount = batch.SolverCallCount
                AssertTrue stats, prefix & ".solved", solveCount > 0
                ' Выбор OUTPUT меняем уже после получения всех физических состояний.
                unitRange.Cells(2, 4).Value2 = lengthNames(caseIndex Mod 3)
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
                If stopAfterCases > 0 Then AppendLine stats, "RUN: " & prefix & ".snapshot.done; " & Audit02ExcelMemory()
                summary.WriteSummary ThisWorkbook, batch, units, section
                If stopAfterCases > 0 Then AppendLine stats, "RUN: " & prefix & ".summary.done; " & Audit02ExcelMemory()
                Audit03AssertAutoCADMmGeometry stats, prefix, section
                If stopAfterCases > 0 Then AppendLine stats, "RUN: " & prefix & ".readback.done; " & Audit02ExcelMemory()
                AssertTrue stats, prefix & ".outputDoesNotSolve", solveCount = batch.SolverCallCount
                AssertTextEquals stats, prefix & ".outputForceUnit", ResultsPropertyValue("ALL", "Output.ForceUnit"), units.OutputForceUnit
                AssertTextEquals stats, prefix & ".outputStressUnit", ResultsPropertyValue("ALL", "Output.StressUnit"), CStr(stressNames(caseIndex Mod 5))
                If caseIndex = 0 Then
                    baselineElements = ResultTable("rngNDMElementResults")
                    baselineGeometry = ResultTable("rngNDMSectionGeometry")
                    baselineProperties = ResultTable("rngNDMSectionProperties")
                    AssertTrue stats, prefix & ".strengthState", ResultsPropertyExists("UNIT1", "State.StrengthState.Epsilon0")
                    AssertTrue stats, prefix & ".crackedState", ResultsPropertyExists("UNIT3", "State.CrackedState.Epsilon0")
                Else
                    actual = ResultTable("rngNDMElementResults")
                    Audit03CompareUnitSnapshot stats, prefix & ".elements", baselineElements, actual, "Elements", _
                        lengthFactor, areaFactor, stressFactor, curvatureFactor, _
                        CDbl(forceFactors((caseIndex \ 9) Mod 3)), CDbl(momentFactors((caseIndex \ 3) Mod 3)), signN, signMx, signMy
                    actual = ResultTable("rngNDMSectionGeometry")
                    Audit03CompareUnitSnapshot stats, prefix & ".geometry", baselineGeometry, actual, "Geometry", _
                        lengthFactor, areaFactor, stressFactor, curvatureFactor, _
                        CDbl(forceFactors((caseIndex \ 9) Mod 3)), CDbl(momentFactors((caseIndex \ 3) Mod 3)), signN, signMx, signMy
                    actual = ResultTable("rngNDMSectionProperties")
                    Audit03CompareUnitSnapshot stats, prefix & ".properties", baselineProperties, actual, "Properties", _
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
    AssertTrue stats, "audit03.unitEquivalent.allCases", caseIndex = 72
    GoTo Restore
Failed:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: audit03.unitEquivalent.runtime; " & prefix & "; " & CStr(Err.Number) & "; " & Err.Description
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
    If stopAfterCases = 0 Then AssertTrue stats, "audit03.unitEquivalent.activeSheetRestored", _
        ThisWorkbook.Application.ActiveSheet Is savedSheet
End Sub

' ДЛЯ ТЕСТОВ: два первых полных unit/sign-варианта локализуют расход памяти
' между расчетом, snapshot writer, summary writer и чтением AutoCAD-данных.
' Это диагностический срез, не замена штатных 72 вариантов и их приемки.
Public Function RunAudit03UnitSignStageDiagnosticTests() As String
    Dim stats As TUiTestStats
    TestAudit03UnitSignEquivalence stats, 2
    AppendLine stats, "TOTAL_UNIT_SIGN_STAGE_DIAGNOSTIC: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunAudit03UnitSignStageDiagnosticTests = stats.Report
End Function

' ДЛЯ ТЕСТОВ: повторяет те же два варианта с изоляцией режима приложения.
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
    AppendLine stats, "RUN: unitStage.screenUpdating=" & CStr(ThisWorkbook.Application.ScreenUpdating)
    TestAudit03UnitSignEquivalence stats, 2
    GoTo Restore
Failed:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: unitStage.screen.runtime; " & CStr(Err.Number) & "; " & Err.Description
Restore:
    If Not guard Is Nothing Then guard.Restore
    savedSheet.Activate
    ThisWorkbook.Application.ScreenUpdating = savedScreen
    AssertTrue stats, "unitStage.screen.restored", ThisWorkbook.Application.ScreenUpdating = savedScreen
    AppendLine stats, "TOTAL_UNIT_SIGN_STAGE_SCREEN: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunAudit03UnitSignStageScreenDiagnosticTests = stats.Report
End Function

' ДЛЯ ТЕСТОВ: нормализует только выводимые численные поля независимыми
' коэффициентами. Strain и PhysicalState не меняются от пользовательских
' знаков; знак момента применяется к усилиям, но не к внутренней KappaX/Y.
Private Sub Audit03CompareUnitSnapshot(ByRef stats As TUiTestStats, ByVal prefix As String, _
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
                    absoluteLimit = tolerance * (1# + Abs(expected))
                    If compareInertiaRoundoff And column = 4 And kind = "Properties" Then
                        If parameter = "Transformed.Ixy" Or parameter = "Transformed.Ixyc" Then
                            Dim roundoffLimit As Double
                            roundoffLimit = Audit03InertiaRoundoffLimit(baseline, parameter)
                            If Abs(converted - expected) > absoluteLimit Then
                                AppendLine stats, "INERTIA_ROUNDOFF: " & prefix & "; parameter=" & parameter & _
                                    "; expected=" & CStr(expected) & "; actual=" & CStr(converted) & "; limit=" & CStr(roundoffLimit)
                            End If
                            If roundoffLimit > absoluteLimit Then absoluteLimit = roundoffLimit
                        End If
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
    AssertTrue stats, prefix & ".equivalent; " & detail, same
End Sub

' ДЛЯ ТЕСТОВ: около нулевого Ixy сравниваем round-trip округление относительно
' sqrt(Ix*Iy), а не самого исчезающе малого Ixy. Восемь машинных epsilon
' ограничивают только ошибку представления модулей; solver tolerance не меняется.
' Исторический unit-sign тест не включает это дополнительное правило.
Private Function Audit03InertiaRoundoffLimit(ByRef properties As Variant, ByVal parameter As String) As Double
    Dim ix As Double, iy As Double, row As Long, suffix As String
    If parameter = "Transformed.Ixyc" Then suffix = "c"
    For row = 2 To UBound(properties, 1)
        If CStr(properties(row, 3)) = "Transformed.Ix" & suffix Then ix = CDbl(properties(row, 4))
        If CStr(properties(row, 3)) = "Transformed.Iy" & suffix Then iy = CDbl(properties(row, 4))
    Next row
    Audit03InertiaRoundoffLimit = 8# * (2# ^ -52#) * Sqr(Abs(ix)) * Sqr(Abs(iy))
End Function

' ДЛЯ ТЕСТОВ: проверяет реальный reader геометрии AutoCAD export при всех
' OUTPUT единицах. Он обязан вернуть координаты/диаметры в мм, площади в мм2
' и инерции в мм4 из сохраненных заголовков Results, не вызывая новый solve.
Private Sub Audit03AssertAutoCADMmGeometry(ByRef stats As TUiTestStats, ByVal prefix As String, _
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
    AssertTrue stats, prefix & ".autoCADGeometryMm; " & detail, same
    AssertTrue stats, prefix & ".autoCADDoesNotSolve", SectionEquilibriumSolveCount() = beforeSolves
End Sub

' ДЛЯ ТЕСТОВ: отдельный сквозной gate INPUT длины, напряжения и кривизны.
' Все сочетания единиц проходят обычный макрос расчета, а не ручной solver.
Public Function RunAudit03InputUnitConsumerTests() As String
    Dim stats As TUiTestStats
    TestAudit03InputUnitConsumers stats
    AppendLine stats, "TOTAL_AUDIT03_INPUT_UNIT_CONSUMERS: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunAudit03InputUnitConsumerTests = stats.Report
End Function

' ДЛЯ ТЕСТОВ: четыре формы и 30 эквивалентных систем INPUT для каждой.
' Геометрия, материалы, допуски, crack width и устойчивость строятся из Config;
' сравниваются численные снимки и все комментарии подробных output-блоков.
' Измененные таблицы и активный лист восстанавливаются даже после ошибки;
' массовый прогон не отображает Results во время каждой табличной записи.
Private Sub TestAudit03InputUnitConsumers(ByRef stats As TUiTestStats)
    Dim rangeNames As Variant, original As Collection, prepared As Collection, name As Variant
    Dim target As Object, loads As Object, lengthNames As Variant, lengthFactors As Variant
    Dim stressNames As Variant, stressFactors As Variant, shape As Variant
    Dim lengthIndex As Long, stressIndex As Long, curvatureIndex As Long, i As Long, caseCount As Long
    Dim baselineElements As Variant, baselineGeometry As Variant, baselineProperties As Variant
    Dim baselineDiagrams As Variant, baselineStrength As Variant, baselineCrack As Variant, baselineStability As Variant
    Dim actual As Variant, prefix As String, message As String, caseFailures As Long, savedSheet As Object
    On Error GoTo Failed
    Set savedSheet = ThisWorkbook.Application.ActiveSheet
    ThisWorkbook.Worksheets.Item("Config").Activate
    rangeNames = Array("rngSystemSettings", "rngUnitSettings", "rngSignConventionSettings", _
        "rngCalculationProfiles", "rngLoadCombinations", "rngConcreteMaterialParameters", _
        "rngSteelMaterialParameters", "rngCircleGeometry", "rngRectSetGeometry", _
        "rngRoundedRectangleGeometry", "rngHollowRectangleGeometry")
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
        SetSystemSetting "Units." & CStr(name) & ".Input", message
        SetSystemSetting "Units." & CStr(name) & ".Output", message
    Next name
    SetSystemSetting "Sign.N.User", "Tension"
    SetSystemSetting "Sign.Mx.User", "+Y tension"
    SetSystemSetting "Sign.My.User", "+X tension"
    SetSystemSetting "Geometry.Source", "Generated"
    SetSystemSetting "General.ExecutionReportEnabled", "No"
    SetSystemSetting "Plot.AutoUpdateAfterCalculation", "No"
    SetSystemSetting "Calculation.ZeroMomentPerDepth", "0.001"
    SetSystemSetting "Load.ReferenceOffsetX", "10"
    SetSystemSetting "Load.ReferenceOffsetY", "-7"
    SetSystemSetting "Mesh.StepX", "50": SetSystemSetting "Mesh.StepY", "50"
    SetSystemSetting "Mesh.BoundarySubdivisions", "1"
    SetSystemSetting "Solver.ToleranceN", "0.1"
    SetSystemSetting "Solver.ToleranceMx", "1": SetSystemSetting "Solver.ToleranceMy", "1"
    SetSystemSetting "Solver.MaxDeltaKappa", "0.00005"
    SetSystemSetting "Capacity.SolutionStrategy", "LoadMultiplier"
    SetSystemSetting "Capacity.SearchMethod", "Bisection"
    SetSystemSetting "SLS.Crack.InitiationSolutionStrategy", "Auto"
    SetSystemSetting "SLS.Crack.Allowable", "0.3"
    SetSystemSetting "Stability.ElementLength", "8000"
    For Each name In Array("PR1", "PR2")
        SetProfileSetting CStr(name), "Calculation.Strength.DirectState", IIf(CStr(name) = "PR1", "Yes", "No")
        SetProfileSetting CStr(name), "Calculation.Strength.Capacity", IIf(CStr(name) = "PR1", "Yes", "No")
        SetProfileSetting CStr(name), "Calculation.Crack.Width", IIf(CStr(name) = "PR2", "Yes", "No")
        SetProfileSetting CStr(name), "Calculation.Stability.Enabled", "Yes"
    Next name
    Set loads = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange
    ClearDataRows loads
    For i = 0 To 3
        loads.Cells(i + 2, 1).Value2 = "INPUT" & CStr(i + 1)
        loads.Cells(i + 2, 2).Value2 = IIf((i Mod 2) = 0, -100000#, 30000#)
        loads.Cells(i + 2, 3).Value2 = IIf((i Mod 2) = 0, 30000000#, -20000000#)
        loads.Cells(i + 2, 4).Value2 = IIf((i Mod 2) = 0, -20000000#, 15000000#)
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
    For Each shape In Array("Circle", "RectSet", "RoundedRectangle", "HollowRectangle")
        For lengthIndex = 0 To 2
            For stressIndex = 0 To 4
                For curvatureIndex = 0 To 1
                    For i = 0 To UBound(rangeNames)
                        ThisWorkbook.Names.Item(CStr(rangeNames(i))).RefersToRange.Formula = prepared.Item(i + 1)
                    Next i
                    SetSystemSetting "Geometry.Type", CStr(shape)
                    SetSystemSetting "Units.Length.Input", CStr(lengthNames(lengthIndex))
                    SetSystemSetting "Units.Stress.Input", CStr(stressNames(stressIndex))
                    SetSystemSetting "Units.Curvature.Input", IIf(curvatureIndex = 0, "1/mm", "1/m")
                    Audit03RescaleInputTables CDbl(lengthFactors(lengthIndex)), CDbl(stressFactors(stressIndex))
                    SetSystemSetting "Solver.MaxDeltaKappa", CStr(0.00005 / IIf(curvatureIndex = 0, 1#, 0.001))
                    prefix = "audit03.inputUnits." & CStr(shape) & ".l" & CStr(lengthIndex) & _
                        ".s" & CStr(stressIndex) & ".k" & CStr(curvatureIndex)
                    AppendLine stats, "RUN: " & prefix
                    caseFailures = stats.Failed
                    message = RunSectionCalculationForWorkbook(ThisWorkbook, False)
                    AssertTrue stats, prefix & ".completed", Len(message) > 0
                    AssertTrue stats, prefix & ".strengthState", ResultsPropertyExists("INPUT1", "State.StrengthState.Epsilon0")
                    AssertTrue stats, prefix & ".capacityState", ResultsPropertyExists("INPUT1", "State.CapacityState.Epsilon0")
                    AssertTrue stats, prefix & ".crackedState", ResultsPropertyExists("INPUT3", "State.CrackedState.Epsilon0")
                    If lengthIndex = 0 And stressIndex = 0 And curvatureIndex = 0 Then
                        baselineElements = ResultTable("rngNDMElementResults")
                        baselineGeometry = ResultTable("rngNDMSectionGeometry")
                        baselineProperties = ResultTable("rngNDMSectionProperties")
                        baselineDiagrams = ResultTable("rngNDMMaterialDiagrams")
                        baselineStrength = ThisWorkbook.Names.Item("rngStrengthSummaryAnchor").RefersToRange.Resize(4, 49).Value2
                        baselineCrack = ThisWorkbook.Names.Item("rngCrackSummaryAnchor").RefersToRange.Resize(4, 72).Value2
                        baselineStability = ThisWorkbook.Names.Item("rngStabilitySummaryAnchor").RefersToRange.Resize(4, 84).Value2
                    Else
                        actual = ResultTable("rngNDMElementResults")
                        Audit03CompareUnitSnapshot stats, prefix & ".elements", baselineElements, actual, "Elements", 1#, 1#, 1#, 1#, 1#, 1#, 1#, 1#, 1#
                        actual = ResultTable("rngNDMSectionGeometry")
                        Audit03CompareUnitSnapshot stats, prefix & ".geometry", baselineGeometry, actual, "Geometry", 1#, 1#, 1#, 1#, 1#, 1#, 1#, 1#, 1#
                        actual = ResultTable("rngNDMSectionProperties")
                        Audit03CompareUnitSnapshot stats, prefix & ".properties", baselineProperties, actual, "Properties", 1#, 1#, 1#, 1#, 1#, 1#, 1#, 1#, 1#, True
                        actual = ResultTable("rngNDMMaterialDiagrams")
                        Audit03ComparePlainSnapshot stats, prefix & ".diagrams", baselineDiagrams, actual, 2
                        actual = ThisWorkbook.Names.Item("rngStrengthSummaryAnchor").RefersToRange.Resize(4, 49).Value2
                        Audit03ComparePlainSnapshot stats, prefix & ".strength", baselineStrength, actual, 1
                        actual = ThisWorkbook.Names.Item("rngCrackSummaryAnchor").RefersToRange.Resize(4, 72).Value2
                        Audit03ComparePlainSnapshot stats, prefix & ".crack", baselineCrack, actual, 1
                        actual = ThisWorkbook.Names.Item("rngStabilitySummaryAnchor").RefersToRange.Resize(4, 84).Value2
                        Audit03ComparePlainSnapshot stats, prefix & ".stability", baselineStability, actual, 1
                    End If
                    caseCount = caseCount + 1
                    AppendLine stats, "INPUT_UNIT_CASE: " & prefix & "; newFailures=" & CStr(stats.Failed - caseFailures)
                Next curvatureIndex
            Next stressIndex
        Next lengthIndex
    Next shape
    AssertTrue stats, "audit03.inputUnits.allCases", caseCount = 120
    GoTo Restore
Failed:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: audit03.inputUnits.runtime; " & prefix & "; " & CStr(Err.Number) & "; " & Err.Description
Restore:
    On Error Resume Next
    If Not original Is Nothing Then
        For i = 0 To original.Count - 1
            ThisWorkbook.Names.Item(CStr(rangeNames(i))).RefersToRange.Formula = original.Item(i + 1)
        Next i
    End If
    If Not savedSheet Is Nothing Then savedSheet.Activate
    On Error GoTo 0
    AssertTrue stats, "audit03.inputUnits.activeSheetRestored", ThisWorkbook.Application.ActiveSheet Is savedSheet
End Sub

' ДЛЯ ТЕСТОВ: изменяет только размерные числовые поля согласно видимым
' подтаблицам Config. Счетчики, enum, деформации, пустые диаметры и формулы
' единиц не масштабируются. Коэффициенты независимы от CUnitSystem.
Private Sub Audit03RescaleInputTables(ByVal lengthFactor As Double, ByVal stressFactor As Double)
    Dim target As Object, name As Variant, row As Long, column As Variant, key As String
    For Each name In Array("rngCircleGeometry", "rngSystemSettings")
        Set target = ThisWorkbook.Names.Item(CStr(name)).RefersToRange
        For row = 2 To target.Rows.Count
            key = CStr(target.Cells(row, 1).Value2)
            Select Case key
                Case "Circle.Diameter", "Rebar.AxisDistance", "Rebar.Diameter", "Rebar.Diameter2", "Rebar.Diameter3", _
                        "Mesh.StepX", "Mesh.StepY", "Load.ReferenceOffsetX", "Load.ReferenceOffsetY", _
                        "Stability.ElementLength", "Stability.AccidentalEccentricityUser1", _
                        "Stability.AccidentalEccentricityUser2", "SLS.Crack.Allowable", "Plot.ResultLabelSpacing", _
                        "SLS.Crack.SP35.GroupGapTolerance", "SLS.Crack.SP35.RowTolerance"
                    Audit03DivideNumericInput target.Cells(row, 2), lengthFactor
            End Select
        Next row
    Next name
    ' Порог имеет размерность момент/длина: моментные INPUT единицы неизменны.
    SetSystemSetting "Calculation.ZeroMomentPerDepth", CStr(0.001 * lengthFactor)
    Set target = ThisWorkbook.Names.Item("rngRectSetGeometry").RefersToRange
    Audit03DivideNumericInput target.Cells(4, 2), lengthFactor
    For column = 1 To 4: Audit03DivideNumericInput target.Cells(8, column), lengthFactor: Next column
    For row = 11 To 18
        For Each column In Array(2, 3, 5, 6): Audit03DivideNumericInput target.Cells(row, CLng(column)), lengthFactor: Next column
    Next row
    For row = 21 To 28
        For Each column In Array(2, 5): Audit03DivideNumericInput target.Cells(row, CLng(column)), lengthFactor: Next column
    Next row
    Set target = ThisWorkbook.Names.Item("rngRoundedRectangleGeometry").RefersToRange
    For column = 2 To 3: Audit03DivideNumericInput target.Cells(3, column), lengthFactor: Next column
    For row = 8 To 10
        For column = 2 To 3: Audit03DivideNumericInput target.Cells(row, column), lengthFactor: Next column
    Next row
    For row = 13 To 16
        For column = 2 To 3: Audit03DivideNumericInput target.Cells(row, column), lengthFactor: Next column
    Next row
    For row = 19 To 22
        For Each column In Array(2, 5): Audit03DivideNumericInput target.Cells(row, CLng(column)), lengthFactor: Next column
    Next row
    Set target = ThisWorkbook.Names.Item("rngHollowRectangleGeometry").RefersToRange
    For row = 3 To 4: Audit03DivideNumericInput target.Cells(row, 2), lengthFactor: Next row
    For column = 1 To 6: Audit03DivideNumericInput target.Cells(7, column), lengthFactor: Next column
    For row = 11 To 18
        For column = 2 To 3: Audit03DivideNumericInput target.Cells(row, column), lengthFactor: Next column
    Next row
    For row = 21 To 28
        For Each column In Array(2, 5): Audit03DivideNumericInput target.Cells(row, CLng(column)), lengthFactor: Next column
    Next row
    For Each name In Array("rngConcreteMaterialParameters", "rngSteelMaterialParameters")
        Set target = ThisWorkbook.Names.Item(CStr(name)).RefersToRange
        For row = 2 To target.Rows.Count
            key = CStr(target.Cells(row, 1).Value2)
            If InStr(key, ".R") > 0 Or key = "Concrete.E" Or key = "Steel.E" Then
                For column = 2 To 3: Audit03DivideNumericInput target.Cells(row, column), stressFactor: Next column
            End If
        Next row
    Next name
End Sub

' ДЛЯ ТЕСТОВ: сохраняет пустой ввод и служебный прочерк; масштабирует число
' в ячейке fixture, не вызывая production-преобразования и не меняя validation.
Private Sub Audit03DivideNumericInput(ByVal cell As Object, ByVal factor As Double)
    If IsNumeric(cell.Value2) And Len(CStr(cell.Value2)) > 0 Then cell.Value2 = CDbl(cell.Value2) / factor
End Sub

' ДЛЯ ТЕСТОВ: сравнивает фактические output-таблицы при одинаковых OUTPUT
' единицах. Тексты и статусы должны совпадать точно; числа сравниваются с
' относительной погрешностью округления 1e-9, без изменения solver tolerance.
Private Sub Audit03ComparePlainSnapshot(ByRef stats As TUiTestStats, ByVal prefix As String, _
        ByRef baseline As Variant, ByRef actual As Variant, ByVal firstColumn As Long)
    Dim row As Long, column As Long, same As Boolean, detail As String
    same = (UBound(baseline, 1) = UBound(actual, 1) And UBound(baseline, 2) = UBound(actual, 2))
    If Not same Then detail = "размер таблицы"
    If same Then
        For row = 1 To UBound(baseline, 1)
            For column = firstColumn To UBound(baseline, 2)
                If IsNumeric(baseline(row, column)) And IsNumeric(actual(row, column)) Then
                    same = Abs(CDbl(actual(row, column)) - CDbl(baseline(row, column))) <= _
                        0.000000001 * (1# + Abs(CDbl(baseline(row, column))))
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
    AssertTrue stats, prefix & ".equivalent; " & detail, same
End Sub

' ДЛЯ ТЕСТОВ: проверяет INPUT-площадь и численный порог на том же Config ->
' ModelSpace pipeline, который вызывает live import, но без подключения DWG.
Public Function RunAudit03InputAreaImportFilterTests() As String
    Dim stats As TUiTestStats
    TestAudit03InputAreaImportFilter stats
    AppendLine stats, "TOTAL_AUDIT03_INPUT_AREA_IMPORT: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunAudit03InputAreaImportFilterTests = stats.Report
End Function

' ДЛЯ ТЕСТОВ: mm2/cm2/m2 должны одинаково отфильтровать Region около порога.
' Саму DWG-геометрию не масштабируем. Нулевой порог, ошибочный ввод и потеря
' обязательного key проверяются отдельно; измененные ячейки восстанавливаются.
Private Sub TestAudit03InputAreaImportFilter(ByRef stats As TUiTestStats)
    Dim unitRange As Object, systemRange As Object, savedUnits As Variant, savedSystem As Variant
    Dim modelSpace As Collection, region As CFakeAcadRegion, importer As CAutoCADSectionModelImporter
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
    SetSystemSetting "Units.Length.Input", "m"
    SetSystemSetting "AutoCAD.Common.ConcreteLayer", "Concrete"
    SetSystemSetting "AutoCAD.Common.RebarLayer", "Reinf"
    Set modelSpace = New Collection
    Set region = New CFakeAcadRegion
    region.Initialize 10000#, 1000#, 2000#, 10000# * 10000# / 12#, 10000# * 10000# / 12#, 0#, "Concrete", "C_MAIN"
    modelSpace.Add region
    Set region = New CFakeAcadRegion
    region.Initialize 2#, 1005#, 2005#, 4# / 12#, 4# / 12#, 0#, "Concrete", "C_SMALL"
    modelSpace.Add region
    areas = Array(24.999, 25#, 25.001)
    For row = 0 To UBound(areas)
        Set region = New CFakeAcadRegion
        region.Initialize CDbl(areas(row)), 980# + row * 20#, 1980#, 1#, 1#, 0#, "Reinf", "R" & CStr(row + 1)
        modelSpace.Add region
    Next row
    Set importer = New CAutoCADSectionModelImporter
    unitNames = Array("mm2", "cm2", "m2")
    factors = Array(1#, 100#, 1000000#)
    solveCount = SectionEquilibriumSolveCount()
    For unitIndex = 0 To 2
        factor = CDbl(factors(unitIndex))
        SetSystemSetting "Units.Area.Input", CStr(unitNames(unitIndex))
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
            Set section = importer.ImportConfiguredModelSpace(modelSpace, settings, units, Audit03AbsentContourLayers())
            prefix = "audit03.inputArea.unit" & CStr(unitIndex) & ".case" & CStr(caseIndex)
            AssertTrue stats, prefix & ".concreteCount", section.ConcreteCount = expectedConcrete
            AssertTrue stats, prefix & ".rebarCount", section.RebarCount = expectedRebar
            AssertClose stats, prefix & ".xMm", section.ConcreteX(1), 1000#, 0.000001
            AssertClose stats, prefix & ".yMm", section.ConcreteY(1), 2000#, 0.000001
            AssertClose stats, prefix & ".areaMm2", section.ConcreteArea(1), 10000#, 0.000001
            AssertClose stats, prefix & ".inertiaMm4", section.ConcreteLocalIx(1), 10000# * 10000# / 12#, 0.001
            If caseIndex = 1 Then
                AssertTrue stats, prefix & ".inclusiveBoundary", section.RebarSourceHandle(1) = "R2"
                AssertClose stats, prefix & ".boundaryArea", section.RebarArea(1), 25#, 0.000001
                AssertClose stats, prefix & ".boundaryX", section.RebarX(1), 1000#, 0.000001
                Set section = importer.ImportConfiguredModelSpace(modelSpace, settings, Nothing, Audit03AbsentContourLayers())
                AssertTrue stats, prefix & ".implicitUnitsRead", section.ConcreteCount = expectedConcrete And section.RebarCount = expectedRebar
                If unitIndex > 0 Then
                    Set disconnected = importer.ImportFromModelSpace(modelSpace, "Concrete", "Reinf", "Rebar", rawValue, units)
                    AssertTrue stats, prefix & ".detectDisconnectedUnits", disconnected.ConcreteCount <> section.ConcreteCount And _
                        disconnected.RebarCount <> section.RebarCount
                End If
            End If
            AppendLine stats, "AREA_FILTER: input=" & CStr(unitNames(unitIndex)) & "|raw=" & CStr(rawValue) & _
                "|thresholdMm2=" & CStr(threshold) & "|concrete=" & CStr(section.ConcreteCount) & "|rebar=" & CStr(section.RebarCount)
        Next caseIndex
    Next unitIndex
    SetSystemSetting "Units.Area.Input", "m2"
    invalid = Array(vbNullString, "TODO", "abc", CVErr(2042), -1#, "1e308")
    For invalidIndex = 0 To UBound(invalid)
        systemRange.Cells(minAreaRow, 2).Value2 = invalid(invalidIndex)
        Audit03CaptureImportError modelSpace, errorNumber, description
        prefix = "audit03.inputArea.invalid" & CStr(invalidIndex)
        AssertTrue stats, prefix & ".rejected", errorNumber <> 0
        AssertTrue stats, prefix & ".address", InStr(1, description, "AutoCAD.Import.MinArea", vbBinaryCompare) > 0
        AssertTrue stats, prefix & ".actualCell", InStr(1, description, systemRange.Cells(minAreaRow, 2).Address(False, False), vbTextCompare) > 0
        AssertTrue stats, prefix & ".action", InStr(1, description, "Введите", vbTextCompare) > 0 Or _
            InStr(1, description, "Исправьте", vbTextCompare) > 0 Or InStr(1, description, "Уменьшите", vbTextCompare) > 0
        AppendLine stats, "AREA_INVALID: case=" & CStr(invalidIndex) & "|error=" & CStr(errorNumber) & "|reason=" & description
    Next invalidIndex
    systemRange.Cells(minAreaRow, 2).Value2 = 25.002 / 1000000#
    Audit03CaptureImportError modelSpace, errorNumber, description
    AssertTrue stats, "audit03.inputArea.emptyRebar.rejected", errorNumber <> 0
    AssertTrue stats, "audit03.inputArea.emptyRebar.reason", InStr(1, description, "арматурных", vbBinaryCompare) > 0 And _
        InStr(1, description, "AutoCAD.Import.MinArea", vbBinaryCompare) > 0
    systemRange.Cells(minAreaRow, 2).Value2 = 10001# / 1000000#
    Audit03CaptureImportError modelSpace, errorNumber, description
    AssertTrue stats, "audit03.inputArea.emptyConcrete.rejected", errorNumber <> 0
    AssertTrue stats, "audit03.inputArea.emptyConcrete.reason", InStr(1, description, "бетонных", vbBinaryCompare) > 0 And _
        InStr(1, description, "AutoCAD.Import.MinArea", vbBinaryCompare) > 0
    systemRange.Cells(minAreaRow, 1).Value2 = "__MissingImportMinimumArea"
    Audit03CaptureImportError modelSpace, errorNumber, description
    AssertTrue stats, "audit03.inputArea.missing.rejected", errorNumber <> 0
    AssertTrue stats, "audit03.inputArea.missing.address", InStr(1, description, "AutoCAD.Import.MinArea", vbBinaryCompare) > 0
    AssertTrue stats, "audit03.inputArea.noStateSolve", SectionEquilibriumSolveCount() = solveCount
    GoTo Restore
Failed:
    AssertTrue stats, "audit03.inputArea.runtime." & CStr(Err.Number) & "." & Err.Description, False
Restore:
    On Error Resume Next
    If Not unitRange Is Nothing Then unitRange.Formula = savedUnits
    If Not systemRange Is Nothing Then systemRange.Formula = savedSystem
    On Error GoTo 0
    If Not unitRange Is Nothing Then
        If IsArray(savedUnits) Then
            invalid = unitRange.Formula
            Audit03ComparePlainSnapshot stats, "audit03.inputArea.restoreUnits", savedUnits, invalid, 1
        End If
    End If
    If Not systemRange Is Nothing Then
        If IsArray(savedSystem) Then
            invalid = systemRange.Formula
            Audit03ComparePlainSnapshot stats, "audit03.inputArea.restoreSettings", savedSystem, invalid, 1
        End If
    End If
End Sub

' ДЛЯ ТЕСТОВ: собирает настройки заново после изменения ячейки и сохраняет
' только фактическую ошибку общего импорта. Ошибка не превращается в default.
Private Sub Audit03CaptureImportError(ByVal modelSpace As Object, ByRef errorNumber As Long, ByRef description As String)
    Dim settings As CSystemSettingsReader, units As CUnitSystem, importer As CAutoCADSectionModelImporter, section As CSectionModel
    errorNumber = 0: description = vbNullString
    On Error GoTo Failed
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem
    units.LoadFromSettings settings
    Set importer = New CAutoCADSectionModelImporter
    Set section = importer.ImportConfiguredModelSpace(modelSpace, settings, units, Audit03AbsentContourLayers())
    Exit Sub
Failed:
    errorNumber = Err.Number: description = Err.Description
    Err.Clear
End Sub

' ДЛЯ ТЕСТОВ: воспроизводит импорт -> смена INPUT/OUTPUT -> кнопка расчета.
' Использует общий импорт существующих fake Region и реальный snapshot/pipeline.
Public Function RunAudit03ImportedSnapshotUnitChangeTests() As String
    Dim stats As TUiTestStats
    TestAudit03ImportedSnapshotUnitChanges stats
    AppendLine stats, "TOTAL_AUDIT03_IMPORTED_UNIT_CHANGES: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunAudit03ImportedSnapshotUnitChangeTests = stats.Report
End Function

' ДЛЯ ТЕСТОВ: все девять переходов OUTPUT-длины mm/cm/m должны сохранить
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
        "rngCircleGeometry", "rngRectSetGeometry", "rngRoundedRectangleGeometry", "rngHollowRectangleGeometry")
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
            PrepareCircleInput
            SetSystemSetting "General.ExecutionReportEnabled", "No"
            SetSystemSetting "Plot.AutoUpdateAfterCalculation", "No"
            SetSystemSetting "Calculation.ZeroMomentPerDepth", "0"
            SetSystemSetting "Load.ReferenceOffsetX", "0": SetSystemSetting "Load.ReferenceOffsetY", "0"
            SetSystemSetting "Solver.MaxDeltaKappa", "0"
            SetSystemSetting "Solver.ToleranceN", "0.001"
            SetSystemSetting "Solver.ToleranceMx", "0.001": SetSystemSetting "Solver.ToleranceMy", "0.001"
            SetProfileSetting "PR1", "Calculation.Strength.DirectState", "Yes"
            SetProfileSetting "PR1", "Calculation.Strength.Capacity", "No"
            SetProfileSetting "PR1", "Calculation.Crack.Width", "No"
            SetProfileSetting "PR1", "Calculation.Stability.Enabled", "No"
            SetProfileSetting "PR1", "Visualization.State", "StrengthState"
            SetSystemSetting "AutoCAD.Import.MinArea", "0.000001"
            SetSystemSetting "AutoCAD.Common.ConcreteLayer", "Concrete"
            SetSystemSetting "AutoCAD.Common.RebarLayer", "Reinf"
            For Each name In Array("Force", "Moment", "Stress", "Curvature")
                SetSystemSetting "Units." & CStr(name) & ".Output", GetSystemSetting("Units." & CStr(name) & ".Input")
            Next name
            SetSystemSetting "Units.Length.Output", CStr(lengthNames(beforeIndex))
            SetSystemSetting "Units.Area.Output", CStr(areaNames(beforeIndex))
            Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
            Set units = New CUnitSystem: units.LoadFromSettings settings
            Set section = Audit03ImportSnapshotFixture(settings, units)
            writer.WriteGeometryPreview ThisWorkbook, section, PrepareSectionSnapshot(section), units
            preview = ResultTable("rngNDMSectionGeometry")
            prefix = "audit03.importUnits.before" & CStr(beforeIndex) & ".after" & CStr(afterIndex)
            AssertClose stats, prefix & ".previewX", GeometryResultValue(preview, "C1", "X"), 950# / CDbl(lengthFactors(beforeIndex)), 0.000001
            AssertClose stats, prefix & ".previewArea", GeometryResultValue(preview, "C1", "Area"), 10000# / CDbl(areaFactors(beforeIndex)), 0.000001
            AssertTextEquals stats, prefix & ".previewHeader", CStr(preview(1, ResultHeaderColumnByBaseName(preview, "X"))), "X, " & CStr(lengthNames(beforeIndex))
            AssertTextEquals stats, prefix & ".previewAreaHeader", CStr(preview(1, ResultHeaderColumnByBaseName(preview, "Area"))), "Area, " & CStr(areaNames(beforeIndex))
            Audit03AssertAutoCADMmGeometry stats, prefix & ".beforeChange", section

            ' Меняем текущие единицы, но не переписываем импортированный snapshot.
            lengthFactor = CDbl(lengthFactors(afterIndex)): stressFactor = CDbl(stressFactors(afterIndex))
            forceFactor = CDbl(forceFactors(afterIndex)): momentFactor = CDbl(momentFactors(afterIndex))
            Audit03RescaleInputTables lengthFactor, stressFactor
            SetSystemSetting "Calculation.ZeroMomentPerDepth", "0"
            SetSystemSetting "Units.Length.Input", CStr(lengthNames(afterIndex))
            SetSystemSetting "Units.Area.Input", CStr(areaNames(afterIndex))
            SetSystemSetting "Units.Force.Input", CStr(forceNames(afterIndex))
            SetSystemSetting "Units.Moment.Input", CStr(momentNames(afterIndex))
            SetSystemSetting "Units.Stress.Input", CStr(stressNames(afterIndex))
            SetSystemSetting "Units.Curvature.Input", IIf(afterIndex = 0, "1/mm", "1/m")
            SetSystemSetting "Units.Length.Output", CStr(lengthNames(afterIndex))
            SetSystemSetting "Units.Area.Output", CStr(areaNames(afterIndex))
            SetSystemSetting "Units.Force.Output", CStr(forceNames(afterIndex))
            SetSystemSetting "Units.Moment.Output", CStr(momentNames(afterIndex))
            SetSystemSetting "Units.Stress.Output", CStr(stressNames(afterIndex))
            SetSystemSetting "Units.Curvature.Output", IIf(afterIndex = 0, "1/mm", "1/m")
            signN = IIf(afterIndex = 1, -1#, 1#): signMx = IIf(afterIndex = 1, -1#, 1#): signMy = IIf(afterIndex = 2, -1#, 1#)
            SetSystemSetting "Sign.N.User", IIf(signN > 0#, "Tension", "Compression")
            SetSystemSetting "Sign.Mx.User", IIf(signMx > 0#, "+Y tension", "-Y tension")
            SetSystemSetting "Sign.My.User", IIf(signMy > 0#, "+X tension", "-X tension")
            SetSystemSetting "Solver.ToleranceN", CStr(0.001 / forceFactor)
            SetSystemSetting "Solver.ToleranceMx", CStr(0.001 / momentFactor)
            SetSystemSetting "Solver.ToleranceMy", CStr(0.001 / momentFactor)
            SetSystemSetting "Geometry.Source", "AutoCAD"
            ' Эти параметры остановили бы новый импорт, но сохраненную модель
            ' при расчете не должны ни перечитывать из DWG, ни фильтровать заново.
            SetSystemSetting "AutoCAD.Import.MinArea", "1e100"
            SetSystemSetting "AutoCAD.Common.ConcreteLayer", "__NoNewConcreteImport"
            SetSystemSetting "AutoCAD.Common.RebarLayer", "__NoNewRebarImport"
            Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
            Set units = New CUnitSystem: units.LoadFromSettings settings
            solveCount = SectionEquilibriumSolveCount()
            Set restored = BuildWorkbookSectionModel(ThisWorkbook, settings, units)
            AssertTextEquals stats, prefix & ".source", restored.SourceType, "AutoCADImport"
            AssertClose stats, prefix & ".restoredX", restored.ConcreteX(1), 950#, 0.000001
            AssertClose stats, prefix & ".restoredArea", restored.ConcreteArea(1), 10000#, 0.000001
            AssertTrue stats, prefix & ".restoreNoSolve", SectionEquilibriumSolveCount() = solveCount
            actual = ResultTable("rngNDMSectionGeometry")
            Audit03ComparePlainSnapshot stats, prefix & ".snapshotUnchangedBeforeSolve", preview, actual, 1
            Audit03AssertAutoCADMmGeometry stats, prefix & ".afterChangeBeforeSolve", section

            Set loads = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange
            ClearDataRows loads
            loads.Cells(2, 1).Value2 = "LC_IMPORT"
            loads.Cells(2, 2).Value2 = -100000# / forceFactor * signN
            loads.Cells(2, 3).Value2 = 1000000# / momentFactor * signMx
            loads.Cells(2, 4).Value2 = -750000# / momentFactor * signMy
            loads.Cells(2, 5).Value2 = "PR1": loads.Cells(2, 6).Value2 = ChrW$(&H3BB) & "*Mxy"
            message = RunSectionCalculationForWorkbook(ThisWorkbook, False)
            AssertTrue stats, prefix & ".run", InStr(1, message, "Расчет импортированной", vbTextCompare) > 0
            AssertTrue stats, prefix & ".hasState", ResultTableRowCount("rngNDMElementResults") > 1
            AssertClose stats, prefix & ".loadN", CDbl(ResultsPropertyValue("LC_IMPORT", "N")) * forceFactor * signN, -100000#, 0.001
            Audit03AssertAutoCADMmGeometry stats, prefix & ".afterSolve", section
            actual = ResultTable("rngNDMSectionGeometry")
            AssertTextEquals stats, prefix & ".newHeader", CStr(actual(1, ResultHeaderColumnByBaseName(actual, "X"))), "X, " & CStr(lengthNames(afterIndex))
            AssertTextEquals stats, prefix & ".newAreaHeader", CStr(actual(1, ResultHeaderColumnByBaseName(actual, "Area"))), "Area, " & CStr(areaNames(afterIndex))
            AssertClose stats, prefix & ".newX", GeometryResultValue(actual, "C1", "X"), 950# / lengthFactor, 0.000001
            actual = ResultTable("rngNDMElementResults")
            If beforeIndex = 0 And afterIndex = 0 Then
                baselineElements = actual
                baselineGeometry = ResultTable("rngNDMSectionGeometry")
                baselineProperties = ResultTable("rngNDMSectionProperties")
            Else
                Audit03CompareUnitSnapshot stats, prefix & ".elements", baselineElements, actual, "Elements", _
                    lengthFactor, CDbl(areaFactors(afterIndex)), stressFactor, IIf(afterIndex = 0, 1#, 0.001), forceFactor, momentFactor, signN, signMx, signMy
                actual = ResultTable("rngNDMSectionGeometry")
                Audit03CompareUnitSnapshot stats, prefix & ".geometry", baselineGeometry, actual, "Geometry", _
                    lengthFactor, CDbl(areaFactors(afterIndex)), stressFactor, 1#, forceFactor, momentFactor, signN, signMx, signMy
                actual = ResultTable("rngNDMSectionProperties")
                Audit03CompareUnitSnapshot stats, prefix & ".properties", baselineProperties, actual, "Properties", _
                    lengthFactor, CDbl(areaFactors(afterIndex)), stressFactor, IIf(afterIndex = 0, 1#, 0.001), forceFactor, momentFactor, signN, signMx, signMy, True
            End If
            AppendLine stats, "IMPORTED_UNIT_CHANGE: before=" & CStr(lengthNames(beforeIndex)) & "|after=" & CStr(lengthNames(afterIndex)) & _
                "|inputForce=" & CStr(forceNames(afterIndex)) & "|inputStress=" & CStr(stressNames(afterIndex)) & "|source=" & ResultsGeometrySource(ThisWorkbook)
        Next afterIndex
    Next beforeIndex

    ' Смена INPUT без изменения числа не сохраняет физическую нагрузку.
    ' Геометрия при этом остается той же, обе осевые задачи здесь линейны.
    loads.Cells(2, 2).Value2 = -100#: loads.Cells(2, 3).Value2 = 0#: loads.Cells(2, 4).Value2 = 0#
    SetSystemSetting "Units.Force.Input", "N"
    SetSystemSetting "Solver.ToleranceN", "0.001"
    message = RunSectionCalculationForWorkbook(ThisWorkbook, False)
    AssertClose stats, "audit03.importUnits.rawUnchanged.oldN", CDbl(ResultsPropertyValue("LC_IMPORT", "N")) * 9806.65, -100#, 0.001
    smallStrain = CDbl(ResultsPropertyValue("LC_IMPORT", "State.StrengthState.Epsilon0"))
    AppendLine stats, "IMPORTED_RAW_LOAD: stage=N; epsilon0=" & CStr(smallStrain) & _
        "; kappaX=" & ResultsPropertyValue("LC_IMPORT", "State.StrengthState.KappaX") & _
        "; kappaY=" & ResultsPropertyValue("LC_IMPORT", "State.StrengthState.KappaY") & _
        "; Nint=" & ResultsPropertyValue("LC_IMPORT", "Nint")
    SetSystemSetting "Units.Force.Input", "kN"
    SetSystemSetting "Solver.ToleranceN", "0.000001"
    message = RunSectionCalculationForWorkbook(ThisWorkbook, False)
    AssertClose stats, "audit03.importUnits.rawUnchanged.number", CDbl(loads.Cells(2, 2).Value2), -100#, 0#
    AssertClose stats, "audit03.importUnits.rawUnchanged.newN", CDbl(ResultsPropertyValue("LC_IMPORT", "N")) * 9806.65, -100000#, 0.001
    AssertClose stats, "audit03.importUnits.rawUnchanged.equilibrium", CDbl(ResultsPropertyValue("LC_IMPORT", "Nint")) * 9806.65, -100000#, 0.001
    largeStrain = CDbl(ResultsPropertyValue("LC_IMPORT", "State.StrengthState.Epsilon0"))
    AppendLine stats, "IMPORTED_RAW_LOAD: stage=kN; epsilon0=" & CStr(largeStrain) & _
        "; kappaX=" & ResultsPropertyValue("LC_IMPORT", "State.StrengthState.KappaX") & _
        "; kappaY=" & ResultsPropertyValue("LC_IMPORT", "State.StrengthState.KappaY") & _
        "; Nint=" & ResultsPropertyValue("LC_IMPORT", "Nint")
    AssertTrue stats, "audit03.importUnits.rawUnchanged.compression", smallStrain < 0# And largeStrain < 0#
    AssertClose stats, "audit03.importUnits.rawUnchanged.strainRatio", largeStrain / smallStrain, 1000#, 0.000001
    Audit03AssertAutoCADMmGeometry stats, "audit03.importUnits.rawUnchanged.geometry", section

    preview = ResultTable("rngNDMSectionGeometry")
    solveCount = SectionEquilibriumSolveCount()
    SetSystemSetting "Units.Length.Output", "ft"
    On Error Resume Next
    message = RunSectionCalculationForWorkbook(ThisWorkbook, False)
    errorNumber = Err.Number: description = Err.Description
    Err.Clear
    On Error GoTo Failed
    AssertTrue stats, "audit03.importUnits.invalidOutput.rejected", errorNumber <> 0
    AssertTrue stats, "audit03.importUnits.invalidOutput.address", InStr(1, description, "Units.Length.Output", vbBinaryCompare) > 0
    AssertTrue stats, "audit03.importUnits.invalidOutput.noSolve", SectionEquilibriumSolveCount() = solveCount
    actual = ResultTable("rngNDMSectionGeometry")
    Audit03ComparePlainSnapshot stats, "audit03.importUnits.invalidOutput.snapshotPreserved", preview, actual, 1
    SetSystemSetting "Units.Length.Output", "m"
    message = RunSectionCalculationForWorkbook(ThisWorkbook, False)
    AssertTrue stats, "audit03.importUnits.invalidOutput.recovery", InStr(1, message, "Расчет импортированной", vbTextCompare) > 0
    Audit03AssertAutoCADMmGeometry stats, "audit03.importUnits.invalidOutput.recoveredGeometry", section
    GoTo Restore
Failed:
    AssertTrue stats, "audit03.importUnits.runtime." & CStr(Err.Number) & "." & Err.Description, False
Restore:
    On Error Resume Next
    If Not original Is Nothing Then
        For i = 0 To original.Count - 1
            ThisWorkbook.Names.Item(CStr(rangeNames(i))).RefersToRange.Formula = original(i + 1)
        Next i
    End If
    On Error GoTo 0
End Sub

' ДЛЯ ТЕСТОВ: четыре квадратных бетонных Region и четыре арматурных Region
' проходят общий importer с Config. Смещенный центр проверяет перенос момента.
Private Function Audit03ImportSnapshotFixture(ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem) As CSectionModel
    Dim modelSpace As Collection, region As CFakeAcadRegion, importer As CAutoCADSectionModelImporter
    Dim x As Double, y As Double, i As Long
    Set modelSpace = New Collection
    For i = 0 To 3
        x = 950# + (i Mod 2) * 100#: y = -750# + (i \ 2) * 100#
        Set region = New CFakeAcadRegion
        region.Initialize 10000#, x, y, 10000# * 10000# / 12#, 10000# * 10000# / 12#, 0#, "Concrete", "C" & CStr(i + 1)
        modelSpace.Add region
        x = 930# + (i Mod 2) * 140#: y = -770# + (i \ 2) * 140#
        Set region = New CFakeAcadRegion
        region.Initialize GEOM_PI * 20# ^ 2 / 4#, x, y, 1#, 1#, 0#, "Reinf", "R" & CStr(i + 1)
        modelSpace.Add region
    Next i
    Set importer = New CAutoCADSectionModelImporter
    Set Audit03ImportSnapshotFixture = importer.ImportConfiguredModelSpace(modelSpace, settings, units, Audit03AbsentContourLayers())
End Function

' ====================== ДЛЯ ТЕСТОВ: ПОНЯТНЫЕ ОШИБКИ CONFIG ======================

' ДЛЯ ТЕСТОВ: отдельный gate адресов, причин и действий в диагностике ввода.
' Полный UI suite запускает те же проверки, а не отдельную упрощенную реализацию.
Public Function RunAudit03SettingErrorMessageTests() As String
    Dim stats As TUiTestStats
    TestAudit03SettingErrorMessages stats
    TestAudit03RequiredTableMessages stats
    TestAudit03ConfigConversionMessages stats
    TestAudit03UnitSignChoices stats
    AppendLine stats, "TOTAL_AUDIT03_SETTING_ERROR_MESSAGES: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunAudit03SettingErrorMessageTests = stats.Report
End Function

' ДЛЯ ТЕСТОВ: обязательные строки берутся из фактического Config. Материальные
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
                AssertTrue stats, prefix & ".error", code <> 0
                Audit03AssertSettingMessage stats, prefix, description, key, cell
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
        AssertTrue stats, "audit03.settingMessage.missing." & key & ".error", code <> 0
        Audit03AssertSettingMessage stats, "audit03.settingMessage.missing." & key, description, key, Nothing
        AssertTrue stats, "audit03.settingMessage.missing." & key & ".restore", InStr(1, description, "Восстановите строку", vbTextCompare) > 0
        keyCell.Value2 = key
    Next i

    ' Входной alias материала должен указывать на правильную сторону диаграммы.
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
        Array("Steel.ThreeLine.Esc2", "Steel.ThreeLine.Es2", 2), Array("Steel.ThreeLine.Es2", "Steel.ThreeLine.Es2", 3))
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
            AssertTrue stats, prefix & ".error", code <> 0
            ' Excel-error возникает до разворачивания alias, поэтому там видимое имя строки.
            If IsError(invalid) Then text = CStr(entry(1)) Else text = key
            Audit03AssertSettingMessage stats, prefix, description, text, cell
            cell.Formula = oldValue
        Next invalid
    Next entry

    ' Проверяем не только исключение reader-а, но и сообщение обычного workbook-сценария.
    SetSystemSetting "Geometry.Source", "Generated"
    SetSystemSetting "Geometry.Type", "Circle"
    SetSystemSetting "General.ExecutionReportEnabled", "No"
    SetSystemSetting "SLS.Crack.Allowable", ""
    loads.Offset(1, 0).Resize(loads.Rows.Count - 1, loads.Columns.Count).ClearContents
    loads.Cells(2, 1).Value2 = "INPUT_MESSAGE": loads.Cells(2, 2).Value2 = 1#
    loads.Cells(2, 3).Value2 = 0#: loads.Cells(2, 4).Value2 = 0#: loads.Cells(2, 5).Value2 = "PR2"
    text = RunSectionCalculationForWorkbook(ThisWorkbook)
    For row = 2 To systemRange.Rows.Count
        If CStr(systemRange.Cells(row, 1).Value2) = "SLS.Crack.Allowable" Then Set cell = systemRange.Cells(row, 2): Exit For
    Next row
    Audit03AssertSettingMessage stats, "audit03.settingMessage.workbook.message", text, "SLS.Crack.Allowable", cell
    description = CStr(ThisWorkbook.Names.Item("rngBatchSummary").RefersToRange.Offset(12, 2).Value2)
    Audit03AssertSettingMessage stats, "audit03.settingMessage.workbook.results", description, "SLS.Crack.Allowable", cell
    GoTo Restore
Failed:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: audit03.settingMessage.runtime; " & CStr(Err.Number) & "; " & Err.Description
Restore:
    On Error Resume Next
    If Not systemRange Is Nothing Then systemRange.Formula = savedSystem
    If Not concreteRange Is Nothing Then concreteRange.Formula = savedConcrete
    If Not steelRange Is Nothing Then steelRange.Formula = savedSteel
    If Not loads Is Nothing Then loads.Formula = savedLoads
    On Error GoTo 0
End Sub

' ДЛЯ ТЕСТОВ: проверяет четыре независимых компонента видимой диагностики.
' Сообщение пишется в evidence без назначения статуса по его тексту.
Private Sub Audit03AssertSettingMessage(ByRef stats As TUiTestStats, ByVal prefix As String, _
        ByVal description As String, ByVal key As String, ByVal cell As Object)
    AssertTrue stats, prefix & ".key", InStr(1, description, key, vbTextCompare) > 0
    AssertTrue stats, prefix & ".sheet", InStr(1, description, "Config", vbTextCompare) > 0
    If Not cell Is Nothing Then AssertTrue stats, prefix & ".address", InStr(1, description, cell.Address(False, False), vbTextCompare) > 0
    AssertTrue stats, prefix & ".action", InStr(1, description, "Введите", vbTextCompare) > 0 Or _
        InStr(1, description, "Заполните", vbTextCompare) > 0 Or InStr(1, description, "Исправьте", vbTextCompare) > 0
    AppendLine stats, "SETTING_MESSAGE: " & prefix & "|" & description
End Sub

' ДЛЯ ТЕСТОВ: повреждает только ссылку обязательной таблицы и проверяет
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
        AssertTrue stats, prefix & ".code", code = vbObjectError + 4316
        AssertTrue stats, prefix & ".sheet", InStr(1, description, "Config", vbTextCompare) > 0
        AssertTrue stats, prefix & ".name", InStr(1, description, CStr(rangeName), vbTextCompare) > 0
        AssertTrue stats, prefix & ".action", InStr(1, description, "Восстановите таблицу", vbTextCompare) > 0 And _
            InStr(1, description, "диспетчере имен", vbTextCompare) > 0
        AppendLine stats, "SETTING_MESSAGE: " & prefix & "|" & description
        Set settings = New CSystemSettingsReader
        settings.LoadFromWorkbook ThisWorkbook
        AssertTrue stats, prefix & ".recovery", settings.HasKey("Units.Length.Input")
    Next rangeName
    Exit Sub
Failed:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: audit03.settingMessage.requiredTable.runtime; " & CStr(Err.Number) & "; " & Err.Description
    On Error Resume Next
    If Not name Is Nothing Then
        If Len(savedReference) > 0 Then name.RefersTo = savedReference
    End If
    On Error GoTo 0
End Sub

' ДЛЯ ТЕСТОВ: отделяет переполнение перевода единиц от ошибки чтения Double.
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
        AssertTrue stats, prefix & ".readableDouble", value > 1E+307
        Set units = New CUnitSystem: units.LoadFromSettings settings
        Set batch = BuildUiBatch()
        batch.ApplySettings settings, units
        batch.AddCombination "CONFIG_OVERFLOW", 0#, 0#, 0#, "PR1", "Переполнение пересчета настройки"
        batch.Execute
        AssertTextEquals stats, prefix & ".status", batch.ResultAt(1).Status, "InputErr"
        AssertTrue stats, prefix & ".noSolve", batch.SolverCallCount = 0
        description = batch.ResultAt(1).OverallMeta.ResultComment
        Audit03AssertSettingMessage stats, prefix, description, CStr(key), cell
        AssertTrue stats, prefix & ".units", InStr(1, description, "INPUT", vbTextCompare) > 0
        target.Formula = savedSettings
        settings.LoadFromWorkbook ThisWorkbook
        batch.ApplySettings settings, units
        batch.Execute
        AssertTrue stats, prefix & ".recovery", batch.ResultAt(1).Status <> "InputErr"
    Next key
    GoTo Restore
Failed:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: audit03.settingMessage.conversion.runtime; " & CStr(Err.Number) & "; " & Err.Description
Restore:
    On Error Resume Next
    If Not target Is Nothing Then target.Formula = savedSettings
    If Not unitRange Is Nothing Then unitRange.Formula = savedUnits
    On Error GoTo 0
End Sub

' ==================== ДЛЯ ТЕСТОВ: ВОССТАНОВЛЕНИЕ ГЕОМЕТРИИ ====================

' Проверяет публичное восстановление модели и reader схемы без AutoCAD/solve.
' Временная книга содержит только заданные численные snapshots и закрывается
' без сохранения; корректность размеров и инерций проверяется независимо.
Public Function RunAudit03GeometrySnapshotTests() As String
    Dim stats As TUiTestStats
    TestAudit03GeometrySnapshotContract stats
    AppendLine stats, "TOTAL_AUDIT03_GEOMETRY_SNAPSHOT: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunAudit03GeometrySnapshotTests = stats.Report
End Function

' Новый обязательный блок пустого v1-снимка в изолированных reader-fixtures.
Private Sub InitializeEmptyContourFixture(ByVal workbook As Object, ByVal sheet As Object)
    workbook.Names.Add Name:="rngNDMSectionContours", RefersTo:="='" & sheet.Name & "'!$GR$5"
    Dim writer As CNDMResultsWriter, contours As CSectionContours
    Set writer = New CNDMResultsWriter: Set contours = New CSectionContours
    writer.WriteSectionContours workbook, contours, "FIXTURE"
End Sub

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
    InitializeEmptyContourFixture fixture, sheet
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
            data = Audit03GeometrySnapshotArray(CStr(units(unitIndex)), CDbl(factors(unitIndex)))
            anchor.Resize(3, 15).Value2 = data
            prefix = "audit03.geometrySnapshot." & CStr(units(unitIndex)) & ".position" & CStr(position)
            Set model = ReadSectionGeometryFromResults(fixture, "AutoCADImport")
            Audit03CheckRestoredGeometry stats, prefix & ".model", model
            reader.LoadGeometryPreviewFromWorkbook fixture, settings
            AssertTrue stats, prefix & ".plot.count", reader.Count = 2
            If reader.Count = 2 Then
                AssertClose stats, prefix & ".plot.width", reader.Width(1), 123.75, 0.000000001
                AssertClose stats, prefix & ".plot.height", reader.Height(1), 87.5, 0.000000001
                AssertClose stats, prefix & ".plot.rotation", reader.Rotation(1), 0.3125, 0.000000001
                AssertClose stats, prefix & ".plot.rebarDiameter", reader.Diameter(2), 12.25, 0.000000001
            End If
            reader.LoadFromWorkbook fixture, settings
            AssertTrue stats, prefix & ".state.count", reader.Count = 2
            If reader.Count = 2 Then
                AssertClose stats, prefix & ".state.concreteStress", reader.ResultValue(1), -1.25, 0#
                AssertClose stats, prefix & ".state.rebarStress", reader.ResultValue(2), 2.75, 0#
            End If
        Next unitIndex
    Next position
    Set anchor = sheet.Range("A5")
    fixture.Names.Item("rngNDMSectionGeometry").RefersTo = "=Snapshot!" & anchor.Address
    data = Audit03GeometrySnapshotArray("mm", 1#)
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
            Audit03CaptureGeometrySnapshotError fixture, errorNumber, description
            AssertTrue stats, prefix & ".rejected", errorNumber <> 0
            AssertTrue stats, prefix & ".field", InStr(1, description, CStr(fieldName(0)), vbTextCompare) > 0
            AssertTrue stats, prefix & ".address", InStr(1, description, "Snapshot!" & anchor.Offset(row, CLng(field) - 1).Address(False, False), vbTextCompare) > 0
            AssertTrue stats, prefix & ".action", InStr(1, description, "Повторите", vbTextCompare) > 0
            If CLng(field) < 12 Then
                For Each operation In Array("Preview", "State")
                    Audit03CapturePlotGeometryError fixture, settings, reader, errorNumber, description, CStr(operation) = "State"
                    AssertTrue stats, prefix & ".plot." & CStr(operation) & ".rejected", errorNumber <> 0
                    AssertTrue stats, prefix & ".plot." & CStr(operation) & ".field", InStr(1, description, CStr(fieldName(0)), vbTextCompare) > 0
                    AssertTrue stats, prefix & ".plot." & CStr(operation) & ".address", InStr(1, description, "Snapshot!" & anchor.Offset(row, CLng(field) - 1).Address(False, False), vbTextCompare) > 0
                    AssertTrue stats, prefix & ".plot." & CStr(operation) & ".action", InStr(1, description, "Повторите", vbTextCompare) > 0
                Next operation
            End If
            anchor.Resize(3, 15).Value2 = data
            Set model = ReadSectionGeometryFromResults(fixture, "AutoCADImport")
            Audit03CheckRestoredGeometry stats, prefix & ".recovery", model
        Next badIndex
    Next field
    For Each field In Array(3, 4, 5, 10)
        cases = cases + 1
        If CLng(field) = 10 Then row = 2 Else row = 1
        anchor.Resize(3, 15).Value2 = data
        anchor.Offset(row, CLng(field) - 1).ClearContents
        prefix = "audit03.geometrySnapshot.requiredBlank." & CStr(data(1, CLng(field)))
        Audit03CaptureGeometrySnapshotError fixture, errorNumber, description
        AssertTrue stats, prefix & ".model", errorNumber = vbObjectError + 4367
        Audit03CapturePlotGeometryError fixture, settings, reader, errorNumber, description
        AssertTrue stats, prefix & ".preview", errorNumber = vbObjectError + 4719
        Audit03CapturePlotGeometryError fixture, settings, reader, errorNumber, description, True
        AssertTrue stats, prefix & ".state", errorNumber = vbObjectError + 4719
    Next field
    ' Пустые необязательные размеры/инерции остаются допустимым отсутствием
    ' геометрической оболочки; это не текст, ошибочно распознанный как число.
    anchor.Resize(3, 15).Value2 = data
    cases = cases + 1
    For Each field In Array(8, 9, 11, 12, 13, 14): anchor.Offset(1, CLng(field) - 1).ClearContents: Next field
    Set model = ReadSectionGeometryFromResults(fixture, "AutoCADImport")
    AssertTrue stats, "audit03.geometrySnapshot.optionalBlank", model.ConcreteCount = 1 And model.RebarCount = 1
    anchor.Resize(3, 15).Value2 = data
    cases = cases + 1
    anchor.Offset(1, 7).NumberFormat = "@": anchor.Offset(1, 7).Value2 = "123.75"
    Set model = ReadSectionGeometryFromResults(fixture, "AutoCADImport")
    AssertClose stats, "audit03.geometrySnapshot.numericText", model.ConcreteWidth(1), 123.75, 0#
    anchor.Offset(1, 7).NumberFormat = "General"
    anchor.Resize(3, 15).Value2 = data
    cases = cases + 1
    anchor.Offset(0, 6).Value2 = "WrongShapeHeader"
    Audit03CaptureGeometrySnapshotError fixture, errorNumber, description
    AssertTrue stats, "audit03.geometrySnapshot.missingShape.controlled", errorNumber = vbObjectError + 4355
    AssertTrue stats, "audit03.geometrySnapshot.missingShape.named", InStr(1, description, "GeometryInterpretationStatus", vbTextCompare) > 0
    anchor.Offset(0, 6).Value2 = "ShapeType"
    cases = cases + 1
    Audit03CaptureGeometrySnapshotError fixture, errorNumber, description
    AssertTrue stats, "audit03.geometrySnapshot.obsoleteShapeHeader.rejected", errorNumber = vbObjectError + 4355
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
    AssertTextEquals stats, "audit03.geometrySnapshot.lateComment", reader.LoadCaseComment, "Комментарий позднего сочетания"
    actual = anchor.Resize(3, 15).Value2
    AssertTrue stats, "audit03.geometrySnapshot.unchanged", Audit02SnapshotTablesEqual(before, actual)
    ' Ширина таблицы тоже должна учитывать реальную границу листа, если
    ' заголовок последнего поля расположен ровно в последнем столбце Excel.
    Set anchor = sheet.Cells(5, sheet.Columns.Count - 14)
    cases = cases + 1
    fixture.Names.Item("rngNDMSectionGeometry").RefersTo = "=Snapshot!" & anchor.Address
    anchor.Resize(3, 15).Value2 = data
    Audit03CaptureGeometrySnapshotError fixture, errorNumber, description
    AssertTrue stats, "audit03.geometrySnapshot.lastColumn.model", errorNumber = 0
    Audit03CapturePlotGeometryError fixture, settings, reader, errorNumber, description
    AssertTrue stats, "audit03.geometrySnapshot.lastColumn.plot", errorNumber = 0
    AssertTrue stats, "audit03.geometrySnapshot.noSolve", SectionEquilibriumSolveCount() = solveCount
    GoTo Restore
Failed:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: audit03.geometrySnapshot.runtime; " & CStr(Err.Number) & "; " & Err.Description
Restore:
    On Error Resume Next
    If Not fixture Is Nothing Then fixture.Close False
    On Error GoTo 0
    AppendLine stats, "CASES_GEOMETRY_SNAPSHOT: independentGeometryFixtures=1; inputVariants=" & CStr(cases) & "; equilibriumCases=0"
End Sub

' Сохраняет ошибку preview через тот же экземпляр reader-а. После исправления
' fixture этот reader должен загрузить данные заново без прежних элементов.
Private Sub Audit03CapturePlotGeometryError(ByVal workbook As Object, ByVal settings As CSystemSettingsReader, _
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

' Задает дробные параметры для проверки сохранности в трех системах OUTPUT.
' A и оболочка согласованы, I заданы отдельно для контроля восстановления;
' это не live Region или инженерный solve. Масштаб fixture независим от reader-а.
Private Function Audit03GeometrySnapshotArray(ByVal unit As String, ByVal factor As Double) As Variant
    Dim data(1 To 3, 1 To 15) As Variant, headers As Variant, column As Long
    headers = Array("ElementID", "MaterialType", "X, " & unit, "Y, " & unit, "Area, " & unit & "2", _
        "MaterialID", "GeometryInterpretationStatus", "Width, " & unit, "Height, " & unit, _
        "Diameter, " & unit, "Rotation, rad", "LocalIx, " & unit & "4", "LocalIy, " & unit & "4", "LocalIxy, " & unit & "4", "Comment")
    For column = 1 To 15: data(1, column) = headers(column - 1): Next column
    data(2, 1) = "C1": data(2, 2) = "Concrete": data(2, 3) = -23.125 / factor: data(2, 4) = 17.625 / factor
    data(2, 5) = 10828.125 / factor ^ 2: data(2, 6) = 1: data(2, 7) = "Rectangle"
    data(2, 8) = 123.75 / factor: data(2, 9) = 87.5 / factor: data(2, 10) = 0#: data(2, 11) = 0.3125
    data(2, 12) = 654321.125 / factor ^ 4: data(2, 13) = 987654.375 / factor ^ 4: data(2, 14) = -23456.625 / factor ^ 4
    data(2, 15) = "Дробные параметры Region"
    data(3, 1) = "R1": data(3, 2) = "Rebar": data(3, 3) = 12.375 / factor: data(3, 4) = -8.125 / factor
    data(3, 5) = 117.8581 / factor ^ 2: data(3, 6) = 1: data(3, 7) = "Circle": data(3, 10) = 12.25 / factor
    data(3, 15) = "Дробная арматура"
    Audit03GeometrySnapshotArray = data
End Function

' Сверяет расчетные числа модели, включая инерции с ненулевым signed Ixy.
' Оболочка и площадь имеют разный смысл; сохраненные I не вычисляются заново.
Private Sub Audit03CheckRestoredGeometry(ByRef stats As TUiTestStats, ByVal prefix As String, ByVal model As CSectionModel)
    AssertTrue stats, prefix & ".concreteCount", model.ConcreteCount = 1
    AssertTrue stats, prefix & ".rebarCount", model.RebarCount = 1
    AssertClose stats, prefix & ".x", model.ConcreteX(1), -23.125, 0.000000001
    AssertClose stats, prefix & ".y", model.ConcreteY(1), 17.625, 0.000000001
    AssertClose stats, prefix & ".area", model.ConcreteArea(1), 10828.125, 0.000000001
    AssertClose stats, prefix & ".width", model.ConcreteWidth(1), 123.75, 0.000000001
    AssertClose stats, prefix & ".height", model.ConcreteHeight(1), 87.5, 0.000000001
    AssertClose stats, prefix & ".rotation", model.ConcreteRotation(1), 0.3125, 0.000000001
    AssertClose stats, prefix & ".ix", model.ConcreteLocalIx(1), 654321.125, 0.000000001
    AssertClose stats, prefix & ".iy", model.ConcreteLocalIy(1), 987654.375, 0.000000001
    AssertClose stats, prefix & ".ixy", model.ConcreteLocalIxy(1), -23456.625, 0.000000001
    If model.RebarCount = 1 Then
        AssertClose stats, prefix & ".rebarX", model.RebarX(1), 12.375, 0.000000001
        AssertClose stats, prefix & ".rebarY", model.RebarY(1), -8.125, 0.000000001
        AssertClose stats, prefix & ".rebarArea", model.RebarArea(1), 117.8581, 0.000000001
        AssertClose stats, prefix & ".rebarDiameter", model.RebarDiameter(1), 12.25, 0.000000001
    End If
End Sub

' Сохраняет фактическую ошибку публичного чтения, не интерпретируя ее текст.
' Перехват локален test fixture и позволяет проверить recovery после ошибки.
Private Sub Audit03CaptureGeometrySnapshotError(ByVal workbook As Object, ByRef number As Long, ByRef description As String)
    Dim model As CSectionModel
    number = 0: description = vbNullString
    On Error GoTo Failed
    Set model = ReadSectionGeometryFromResults(workbook, "AutoCADImport")
    Exit Sub
Failed:
    number = Err.Number: description = Err.Description
    Err.Clear
End Sub

' ==================== ДЛЯ ТЕСТОВ: ОШИБКА ЧТЕНИЯ И ПОВТОРНЫЙ ВВОД ====================

' Запускает направленную проверку незавершенной инициализации и поврежденных
' чисел сохраненного снимка. Использует только собственную временную книгу,
' не ищет равновесие и не меняет исходный Config/Results основной книги.
Public Function RunAudit03ReadLifecycleTests() As String
    Dim stats As TUiTestStats
    TestAudit03ReadLifecycleContracts stats
    AppendLine stats, "TOTAL_AUDIT03_READ_LIFECYCLE: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunAudit03ReadLifecycleTests = stats.Report
End Function

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
        AssertTrue stats, prefix & ".error", code <> 0
        AssertTrue stats, prefix & ".incomplete", Not spec.IsComplete
        AssertTextEquals stats, prefix & ".key", spec.SpecKey, "<empty>"
        Set copied = spec.Clone
        AssertTrue stats, prefix & ".cloneIncomplete", Not copied.IsComplete
        On Error Resume Next
        Err.Clear
        Set diagram = Nothing: Set diagram = provider.ConcreteMaterialFromSpec(spec)
        code = Err.Number
        On Error GoTo Failed
        AssertTrue stats, prefix & ".providerRejected", code = vbObjectError + 3251
        AssertTrue stats, prefix & ".noDiagram", diagram Is Nothing
        AssertTextEquals stats, prefix & ".snapshot", savedSpec.SpecKey, "SLS(II)|ThreeLine|UseDiagram|TwoLine"
        spec.Initialize "ULS(I)", "TwoLine", "Ignore", "ThreeLine"
        Set diagram = provider.ConcreteMaterialFromSpec(spec)
        AssertTrue stats, prefix & ".recovery", spec.IsComplete
        AssertTextEquals stats, prefix & ".recoveryKey", spec.SpecKey, "ULS(I)|TwoLine|Ignore|ThreeLine"
        AssertTrue stats, prefix & ".distinctDiagram", Not diagram Is baselineDiagram
        AppendLine stats, "READ_LIFECYCLE_SPEC: " & prefix & "|" & reason
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
                AssertTrue stats, prefix & ".error", code <> 0
                AssertTrue stats, prefix & ".address", InStr(1, reason, sheet.Name & "!" & target.Cells(row, column).Address(False, False), vbBinaryCompare) > 0
                AssertTrue stats, prefix & ".empty", catalog.Count = 0
                AssertTrue stats, prefix & ".noPR1", Not catalog.HasProfile("PR1")
                AssertTrue stats, prefix & ".noPR4", Not catalog.HasProfile("PR4")
                On Error Resume Next
                Err.Clear: Set profile = Nothing: Set profile = catalog.ProfileById("PR1")
                code = Err.Number
                On Error GoTo Failed
                AssertTrue stats, prefix & ".lookupRejected", code = vbObjectError + 3984
                AssertTrue stats, prefix & ".noPublishedProfile", profile Is Nothing
                AssertTextEquals stats, prefix & ".oldSnapshot", savedProfile.ProfileId, "PR1"
                target.Value2 = baseline: catalog.LoadFromRange target
                AssertTrue stats, prefix & ".recoveryCount", catalog.Count = 4
                AssertTrue stats, prefix & ".recoveryPR4", catalog.HasProfile("PR4")
                AppendLine stats, "READ_LIFECYCLE_CATALOG: " & prefix & "|" & reason
            Next index
        Next bad
    Next position

    fixture.Names.Add Name:="rngCalculationProfiles", RefersTo:="=ReadLifecycle!" & profiles.Address
    For row = 2 To profiles.Rows.Count
        If CStr(profiles.Cells(row, 2).Value2) = "Visualization.State" Then profiles.Cells(row, 3).Value2 = "StrengthState"
    Next row
    values = Audit03GeometrySnapshotArray("mm", 1#): sheet.Range("A5").Resize(3, 15).Value2 = values
    fixture.Names.Add Name:="rngNDMSectionGeometry", RefersTo:="=ReadLifecycle!$A$5"
    InitializeEmptyContourFixture fixture, sheet
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
            Audit03ReadLifecycleLoad reader, fixture, settings, mode, code, reason
            AssertTrue stats, "audit03.readLifecycle.reader.baseline.p" & CStr(position) & ".m" & CStr(mode), code = 0 And reader.Count = 2
            For index = 0 To UBound(keys)
                If mode = 1 And index >= 12 Then Exit For ' AutoCAD-preview не использует плоскость выбранного State.
                For badIndex = 0 To 4
                    bad = badValues(badIndex)
                    row = index + 2: target.Cells(row, 3).Value2 = bad
                    Audit03ReadLifecycleLoad reader, fixture, settings, mode, code, reason
                    prefix = "audit03.readLifecycle.reader.p" & CStr(position) & ".m" & CStr(mode) & ".field" & CStr(index) & ".bad" & CStr(badIndex)
                    cases = cases + 1
                    AssertTrue stats, prefix & ".error", code = vbObjectError + 4719
                    AssertTrue stats, prefix & ".field", InStr(1, reason, CStr(keys(index)), vbTextCompare) > 0
                    AssertTrue stats, prefix & ".address", InStr(1, reason, sheet.Name & "!" & target.Cells(row, 3).Address(False, False), vbBinaryCompare) > 0
                    AssertTrue stats, prefix & ".action", InStr(1, reason, "Повторите", vbTextCompare) > 0
                    AssertTrue stats, prefix & ".empty", reader.Count = 0 And reader.AnnotationCount = 0
                    AssertTrue stats, prefix & ".clearedPlane", reader.Epsilon0 = 0# And reader.KappaX = 0# And reader.KappaY = 0#
                    AssertTrue stats, prefix & ".clearedSelection", Len(reader.LoadCase) = 0 And Len(reader.ProfileId) = 0
                    AppendLine stats, "READ_LIFECYCLE_SNAPSHOT: " & prefix & "|" & reason
                    target.Cells(row, 3).Value2 = props(row, 3)
                    Audit03ReadLifecycleLoad reader, fixture, settings, mode, code, reason
                    AssertTrue stats, prefix & ".recovery", code = 0 And reader.Count = 2
                    If mode = 0 Then AssertClose stats, prefix & ".recoveryPlane", reader.Epsilon0, 0.125, 0#
                Next badIndex
            Next index
            ' Поздняя ошибка после AppendElement не публикует первый элемент как весь снимок.
            sheet.Range("A5").Cells(3, 3).Value2 = "late-invalid"
            Audit03ReadLifecycleLoad reader, fixture, settings, mode, code, reason
            prefix = "audit03.readLifecycle.reader.late.p" & CStr(position) & ".m" & CStr(mode)
            AssertTrue stats, prefix & ".error", code = vbObjectError + 4719
            AssertTrue stats, prefix & ".empty", reader.Count = 0 And reader.AnnotationCount = 0
            sheet.Range("A5").Resize(3, 15).Value2 = values
            Audit03ReadLifecycleLoad reader, fixture, settings, mode, code, reason
            AssertTrue stats, prefix & ".recovery", code = 0 And reader.Count = 2
        Next mode
        For index = 0 To UBound(keys)
            target.Cells(index + 2, 3).Value2 = "0.125"
        Next index
        Audit03ReadLifecycleLoad reader, fixture, settings, 0, code, reason
        AssertTrue stats, "audit03.readLifecycle.numericText.p" & CStr(position), code = 0 And reader.Count = 2
        AssertClose stats, "audit03.readLifecycle.numericText.epsilon.p" & CStr(position), reader.Epsilon0, 0.125, 0#
        AssertClose stats, "audit03.readLifecycle.numericText.kappaX.p" & CStr(position), reader.KappaX, 0.125, 0#
        AssertClose stats, "audit03.readLifecycle.numericText.kappaY.p" & CStr(position), reader.KappaY, 0.125, 0#
        target.Resize(22, 4).Value2 = props
        For column = 5 To 6
            For badIndex = 0 To 4
                sheet.Range("R35").Cells(3, column).Value2 = badValues(badIndex)
                Audit03ReadLifecycleLoad reader, fixture, settings, 0, code, reason
                prefix = "audit03.readLifecycle.element.p" & CStr(position) & ".column" & CStr(column) & ".bad" & CStr(badIndex)
                cases = cases + 1
                AssertTrue stats, prefix & ".error", code = vbObjectError + 4719
                AssertTrue stats, prefix & ".field", InStr(1, reason, CStr(results(1, column)), vbTextCompare) > 0
                AssertTrue stats, prefix & ".address", InStr(1, reason, sheet.Name & "!" & sheet.Range("R35").Cells(3, column).Address(False, False), vbBinaryCompare) > 0
                AssertTrue stats, prefix & ".empty", reader.Count = 0
                AppendLine stats, "READ_LIFECYCLE_ELEMENT: " & prefix & "|" & reason
                sheet.Range("R35").Resize(3, 7).Value2 = results
                Audit03ReadLifecycleLoad reader, fixture, settings, 0, code, reason
                AssertTrue stats, prefix & ".recovery", code = 0 And reader.Count = 2
            Next badIndex
        Next column
        sheet.Range("R35").Cells(2, 5).Value2 = "-0.0001"
        sheet.Range("R35").Cells(3, 6).Value2 = "2.75"
        Audit03ReadLifecycleLoad reader, fixture, settings, 0, code, reason
        AssertTrue stats, "audit03.readLifecycle.element.numericText.p" & CStr(position), code = 0 And reader.Count = 2
        AssertClose stats, "audit03.readLifecycle.element.numericValue.p" & CStr(position), reader.ResultValue(2), 2.75, 0#
        sheet.Range("R35").Resize(3, 7).Value2 = results
        actual = target.Resize(22, 4).Value2
        Audit03ComparePlainSnapshot stats, "audit03.readLifecycle.snapshotUnchanged.p" & CStr(position), props, actual, 1
    Next position
    AssertTrue stats, "audit03.readLifecycle.noSolve", SectionEquilibriumSolveCount() = solveCount
    AppendLine stats, "READ_LIFECYCLE_CASES: variants=" & CStr(cases) & "; geometryFixtures=1; equilibriumCases=0"
    GoTo Cleanup
Failed:
    AssertTrue stats, "audit03.readLifecycle.runtime." & CStr(Err.Number) & "." & Err.Description, False
Cleanup:
    On Error Resume Next
    If Not fixture Is Nothing Then fixture.Close False
    On Error GoTo 0
End Sub

' Сохраняет typed VBA-ошибку конкретного публичного reader-режима. Повторное
' чтение выполняется на том же объекте; helper не подменяет snapshot/Config.
Private Sub Audit03ReadLifecycleLoad(ByVal reader As CSectionPlotDataReader, ByVal workbook As Object, _
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

' ==================== ДЛЯ ТЕСТОВ: ОБЩИЕ НАСТРОЙКИ EXCEL-СХЕМЫ ====================

' Сохраняет направленный entrypoint интерфейсного набора; сами проверки
' готового snapshot и layout находятся в отдельном стандартном test-модуле.
Public Function RunAudit03GeneralPlotTests() As String
    RunAudit03GeneralPlotTests = modTestPlotConfig.RunAudit03GeneralPlotTests()
End Function

' ==================== ДЛЯ ТЕСТОВ: СОСТОЯНИЕ EXCEL.APPLICATION ====================

' Проверяет восстановление настоящего отдельного Excel.Application, включая
' реальный отказ Calculation setter без открытых книг. Fake-классы и
' пользовательский экземпляр Excel не используются.
Public Function RunAudit03ExcelGuardTests() As String
    Dim stats As TUiTestStats
    TestAudit03ExcelGuardContracts stats
    AppendLine stats, "TOTAL_AUDIT03_EXCEL_GUARD: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunAudit03ExcelGuardTests = stats.Report
End Function

' Восемь сочетаний Boolean и три режима пересчета проверяются через тот же
' guard, что используется workbook-сценарием. После failed Enter проверяем
' состояние сразу, затем явный Restore, повторный Enter и деструктор.
Private Sub TestAudit03ExcelGuardContracts(ByRef stats As TUiTestStats)
    Dim otherExcel As Object, book As Object, guard As CExcelAppStateGuard, owned As Boolean
    Dim mode As Variant, flags As Long, screen As Boolean, events As Boolean, alerts As Boolean
    Dim code As Long, source As String, reason As String, nativeCode As Long, nativeSource As String, nativeReason As String
    Dim prefix As String, solveCount As Long
    On Error GoTo Failed
    solveCount = SectionEquilibriumSolveCount()
    Set otherExcel = CreateObject("Excel.Application")
    owned = (otherExcel.Hwnd <> Application.Hwnd)
    If Not owned Then Err.Raise 5, "TestAudit03ExcelGuardContracts", "Не создан отдельный Excel.Application для теста."
    otherExcel.Visible = False
    AssertTrue stats, "audit03.excelGuard.emptyApplication", otherExcel.Workbooks.Count = 0
    On Error Resume Next
    Err.Clear: otherExcel.Calculation = xlCalculationManual
    nativeCode = Err.Number: nativeSource = Err.Source: nativeReason = Err.Description
    On Error GoTo Failed
    AssertTrue stats, "audit03.excelGuard.nativeFault", nativeCode <> 0
    Set guard = New CExcelAppStateGuard
    For flags = 0 To 7
        screen = ((flags And 1) <> 0): events = ((flags And 2) <> 0): alerts = ((flags And 4) <> 0)
        otherExcel.ScreenUpdating = screen: otherExcel.EnableEvents = events: otherExcel.DisplayAlerts = alerts
        On Error Resume Next
        Err.Clear: guard.Enter otherExcel
        code = Err.Number: source = Err.Source: reason = Err.Description
        On Error GoTo Failed
        prefix = "audit03.excelGuard.failure.f" & CStr(flags)
        AssertTrue stats, prefix & ".originalError", code = nativeCode And source = nativeSource And reason = nativeReason
        AssertTrue stats, prefix & ".immediateScreen", otherExcel.ScreenUpdating = screen
        AssertTrue stats, prefix & ".immediateEvents", otherExcel.EnableEvents = events
        AssertTrue stats, prefix & ".immediateAlerts", otherExcel.DisplayAlerts = alerts
        guard.Restore
        AssertTrue stats, prefix & ".restore", otherExcel.ScreenUpdating = screen And otherExcel.EnableEvents = events And otherExcel.DisplayAlerts = alerts
        AppendLine stats, "EXCEL_GUARD_ERROR: " & prefix & "|code=" & CStr(code) & "|source=" & source & "|" & reason
        ' Возвращаем fixture в заданное состояние и при negative реализации.
        otherExcel.ScreenUpdating = screen: otherExcel.EnableEvents = events: otherExcel.DisplayAlerts = alerts
    Next flags
    Set book = otherExcel.Workbooks.Add(-4167)
    For Each mode In Array(xlCalculationAutomatic, xlCalculationManual, xlCalculationSemiautomatic)
        For flags = 0 To 7
            screen = ((flags And 1) <> 0): events = ((flags And 2) <> 0): alerts = ((flags And 4) <> 0)
            otherExcel.Calculation = mode
            otherExcel.ScreenUpdating = screen: otherExcel.EnableEvents = events: otherExcel.DisplayAlerts = alerts
            prefix = "audit03.excelGuard.success.m" & CStr(mode) & ".f" & CStr(flags)
            guard.Enter otherExcel
            AssertTrue stats, prefix & ".entered", Not otherExcel.ScreenUpdating And Not otherExcel.EnableEvents And Not otherExcel.DisplayAlerts And otherExcel.Calculation = xlCalculationManual
            guard.Enter otherExcel
            guard.Restore
            AssertTrue stats, prefix & ".restored", otherExcel.ScreenUpdating = screen And otherExcel.EnableEvents = events And otherExcel.DisplayAlerts = alerts And otherExcel.Calculation = mode
            guard.Restore
            AssertTrue stats, prefix & ".idempotent", otherExcel.ScreenUpdating = screen And otherExcel.EnableEvents = events And otherExcel.DisplayAlerts = alerts And otherExcel.Calculation = mode
            guard.Enter otherExcel: Set guard = Nothing
            AssertTrue stats, prefix & ".terminate", otherExcel.ScreenUpdating = screen And otherExcel.EnableEvents = events And otherExcel.DisplayAlerts = alerts And otherExcel.Calculation = mode
            Set guard = New CExcelAppStateGuard
        Next flags
    Next mode
    AssertTrue stats, "audit03.excelGuard.noSolve", SectionEquilibriumSolveCount() = solveCount
    AppendLine stats, "EXCEL_GUARD_CASES: failedEnterVariants=8; validEnterVariants=24; equilibriumCases=0"
    GoTo Cleanup
Failed:
    AssertTrue stats, "audit03.excelGuard.runtime." & CStr(Err.Number) & "." & Err.Description, False
Cleanup:
    On Error Resume Next
    If Not guard Is Nothing Then guard.Restore
    Set guard = Nothing
    If owned Then
        otherExcel.DisplayAlerts = False
        If Not book Is Nothing Then book.Close False
        otherExcel.Quit
    End If
    Set book = Nothing: Set otherExcel = Nothing
    On Error GoTo 0
End Sub

' ================== ДЛЯ ТЕСТОВ: СПИСКИ И РАМКА CONFIG ==================

' Проверяет действующие списки в сохраненной книге, а не только формулы
' validation из сборщика. Каждое значение источника должно приниматься
' Excel как отдельный вариант; исходная формула ячейки возвращается сразу.
Public Function RunConfigPresentationTests() As String
    Dim stats As TUiTestStats
    TestConfigDropdownChoices stats
    TestConfigDropdownMeaning stats
    TestConfigSettingsRightBorder stats
    TestRectSetIndependentSelectorLayout stats
    TestAutoCADCommonSettingsLayout stats
    TestAutoCADGeometryContourWarning stats
    RunConfigPresentationTests = stats.Report & "TOTAL_CONFIG_PRESENTATION: passed=" & _
        CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
End Function

' Проверяет предупреждение в инструкции подготовки геометрии, включая
' риски завышения/занижения площади и красное выделение всего абзаца.
' Адреса не фиксированы: после пересборки справки ищутся ее заголовки.
Private Sub TestAutoCADGeometryContourWarning(ByRef stats As TUiTestStats)
    Dim guide As Object, warning As Object, heading As Object, nextHeading As Object
    Dim row As Long, text As String, cell As Object, used As Object, values As Variant, value As String, sectionText As String
    On Error GoTo Failed
    Set guide = ThisWorkbook.Worksheets("Справка")
    ' Читаем и скрытые строки одним массивом, независимо от настроек Find.
    Set used = guide.UsedRange: values = used.Value2
    For row = 1 To UBound(values, 1)
        value = CStr(values(row, 1))
        If InStr(1, value, "ВНИМАНИЕ: для расчета раскрытия трещин", vbTextCompare) = 1 Then Set warning = guide.Cells(used.Row + row - 1, used.Column)
        If value = "2. Точный контур и несколько отверстий" Then Set heading = guide.Cells(used.Row + row - 1, used.Column)
        If value = "3. Единицы, поворот и плоскость" Then Set nextHeading = guide.Cells(used.Row + row - 1, used.Column)
    Next row
    AssertTrue stats, "guide.autoCAD.contourWarning.exists", Not warning Is Nothing And Not heading Is Nothing And Not nextHeading Is Nothing
    If warning Is Nothing Or heading Is Nothing Or nextHeading Is Nothing Then Exit Sub
    AssertTrue stats, "guide.autoCAD.contourWarning.location", warning.Row > heading.Row And warning.Row < nextHeading.Row
    For row = warning.Row To nextHeading.Row - 1
        Set cell = guide.Cells(row, 1)
        If CLng(cell.Font.Color) <> 255 Then Exit For
        text = text & " " & CStr(cell.Value2)
        AssertTrue stats, "guide.autoCAD.contourWarning.redBold." & CStr(row), CBool(cell.Font.Bold)
    Next row
    AssertTrue stats, "guide.autoCAD.contourWarning.recommendation", InStr(1, text, "настоятельно рекомендуется", vbTextCompare) > 0
    AssertTrue stats, "guide.autoCAD.contourWarning.armature", InStr(1, text, "взаимодействия бетона с растянутой арматурой", vbTextCompare) > 0
    AssertTrue stats, "guide.autoCAD.contourWarning.risks", InStr(1, text, "завышены или занижены", vbTextCompare) > 0
    For row = heading.Row + 1 To nextHeading.Row - 1
        sectionText = sectionText & " " & CStr(guide.Cells(row, 1).Value2)
    Next row
    AssertTrue stats, "guide.autoCAD.knownOuter.solid", InStr(1, sectionText, "бетон внутри контура считается сплошным", vbTextCompare) > 0
    AssertTrue stats, "guide.autoCAD.knownOuter.noMeshVoids", InStr(1, sectionText, "Пустоты по сетке не восстанавливаются", vbTextCompare) > 0
    Exit Sub
Failed:
    AssertTrue stats, "guide.autoCAD.contourWarning.runtime." & CStr(Err.Number) & "." & Err.Description, False
End Sub

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
        AssertTrue stats, "config.subsection.russian." & CStr(row), key <> "[Common]"
        If key = "[AutoCAD]" Then
            section = "AutoCAD": subsection = vbNullString
        ElseIf Left$(key, 1) = "[" Then
            subsection = key
        ElseIf Left$(key, 8) = "AutoCAD." Then
            If Not counts.Exists(key) Then counts.Add key, 0
            counts(key) = CLng(counts(key)) + 1: keyRows(key) = row
            If Left$(key, 15) = "AutoCAD.Common." Then
                AssertTrue stats, "config.autoCAD.common." & key, section = "AutoCAD" And subsection = "[Общие настройки]"
            ElseIf key = "AutoCAD.Export.CrackInteractionLayer" Then
                AssertTrue stats, "config.autoCAD.crack.exportOnly", subsection = "[Экспорт]"
            End If
        End If
    Next row
    For Each expected In Array("AutoCAD.Common.ConcreteLayer", "AutoCAD.Common.RebarLayer", _
            "AutoCAD.Common.SectionContourLayer", "AutoCAD.Common.OpeningContourLayer", "AutoCAD.Export.CrackInteractionLayer")
        AssertTrue stats, "config.autoCAD.unique." & CStr(expected), counts.Exists(CStr(expected))
        If counts.Exists(CStr(expected)) Then AssertTrue stats, "config.autoCAD.once." & CStr(expected), CLng(counts(CStr(expected))) = 1
    Next expected
    For Each expected In Array("AutoCAD.Import.ConcreteLayer", "AutoCAD.Import.RebarLayer", _
            "AutoCAD.Layer.Concrete", "AutoCAD.Layer.Rebar", "AutoCAD.Common.CrackInteractionLayer")
        AssertTrue stats, "config.autoCAD.removed." & CStr(expected), Not counts.Exists(CStr(expected))
    Next expected
    If keyRows.Exists("AutoCAD.Export.CrackInteractionLayer") And keyRows.Exists("AutoCAD.Layer.ConcreteTension") Then _
        AssertTrue stats, "config.autoCAD.crack.beforeConcreteTension", keyRows("AutoCAD.Export.CrackInteractionLayer") + 1 = keyRows("AutoCAD.Layer.ConcreteTension")
    Exit Sub
Failed:
    AssertTrue stats, "config.autoCAD.layout.runtime." & CStr(Err.Number) & "." & Err.Description, False
End Sub

' Проверяет смысл вариантов, а не только способность Excel открыть список.
' Независимые ожидаемые наборы не позволяют пройти тесту при подмене
' Stacked/SideBySide значениями mm/pt или при смешении иных настроек.
Private Sub TestConfigDropdownMeaning(ByRef stats As TUiTestStats)
    Dim table As Object, cell As Object, source As Object, name As Variant, key As String, expected As String, actual As String
    Dim row As Long, column As Long, headerRow As Long, checked As Long, formula As String, validationType As Long
    On Error GoTo Failed
    For Each name In Array("rngSystemSettings", "rngUnitSettings", "rngSignConventionSettings", _
            "rngCalculationProfiles", "rngPlotAnnotationSettings", "rngCircleGeometry", _
            "rngRectSetGeometry", "rngRoundedRectangleGeometry", "rngHollowRectangleGeometry", "rngLoadCombinations")
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
                expected = ExpectedConfigDropdownOptions(key)
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
                If Len(expected) > 0 Then
                    formula = CStr(cell.Validation.Formula1)
                    Set source = cell.Worksheet.Range(Mid$(formula, 2))
                    actual = vbNullString
                    For Each cell In source.Cells
                        If Len(actual) > 0 Then actual = actual & "|"
                        actual = actual & CStr(cell.Value2)
                    Next cell
                    AssertTextEquals stats, "config.dropdown.meaning." & CStr(name) & "." & CStr(row) & "." & CStr(column), actual, expected
                    checked = checked + 1
                End If
NextMeaningCell:
            Next column
        Next row
    Next name
    AssertTrue stats, "config.dropdown.meaning.coverage", checked > 150
    AppendLine stats, "CONFIG_DROPDOWN_MEANING: checked=" & CStr(checked)
    Exit Sub
Failed:
    AssertTrue stats, "config.dropdown.meaning.runtime." & CStr(Err.Number) & "." & Err.Description, False
End Sub

' Возвращает независимый oracle допустимых вариантов конкретной настройки.
' Динамические профили и LC проверяются отдельно; здесь нет адресов helper-ячеек.
Private Function ExpectedConfigDropdownOptions(ByVal key As String) As String
    Select Case key
        Case "Geometry.Source": ExpectedConfigDropdownOptions = "Generated|AutoCAD"
        Case "General.WorstCombinationCriterion": ExpectedConfigDropdownOptions = "StrengthStrain|StrengthCapacity|Cracks|Stability"
        Case "Stability.Code", "SLS.Crack.Code": ExpectedConfigDropdownOptions = "SP63|SP35"
        Case "Stability.SystemType": ExpectedConfigDropdownOptions = "Determinate|Indeterminate"
        Case "Stability.ZeroMomentEccentricitySign1", "Stability.ZeroMomentEccentricitySign2": ExpectedConfigDropdownOptions = "1|-1"
        Case "Stability.AccidentalEccentricityMode": ExpectedConfigDropdownOptions = "AutoWithL|AutoWith" & ChrW(&H3BC) & "L|User"
        Case "Stability.AccidentalEccentricityPlanes": ExpectedConfigDropdownOptions = "OnlyMomentPlane|BothPlanes"
        Case "Stability.PhiLMode": ExpectedConfigDropdownOptions = "Auto|PhiL2"
        Case "Geometry.Type": ExpectedConfigDropdownOptions = "RoundedRectangle|HollowRectangle|Circle|RectSet"
        Case "Solver.Method": ExpectedConfigDropdownOptions = "Newton|Secant"
        Case "Capacity.SolutionStrategy", "SLS.Crack.InitiationSolutionStrategy": ExpectedConfigDropdownOptions = "Auto|UltimateStrain|LoadMultiplier"
        Case "Capacity.SearchMethod": ExpectedConfigDropdownOptions = "Bisection|Brent|Secant"
        Case "SLS.Crack.SP35.RadiusDiameterMode": ExpectedConfigDropdownOptions = "Max|Min|Average"
        Case "SLS.Crack.SP35.InteractionRadiusMode": ExpectedConfigDropdownOptions = "3d|5d|6d"
        Case "SLS.Crack.SP35.RebarProfile": ExpectedConfigDropdownOptions = "Periodic|Smooth"
        Case "SLS.Crack.Phi3Mode": ExpectedConfigDropdownOptions = "Auto|User"
        Case "SLS.Crack.PsiMode": ExpectedConfigDropdownOptions = "User|Auto|AlwaysCalc"
        Case "SLS.Crack.SigmaSCrcAveragingMode": ExpectedConfigDropdownOptions = "AllSelected|TensionOnly"
        Case "SLS.Crack.TensionZoneMode": ExpectedConfigDropdownOptions = "Effective|FullTension"
        Case "SLS.Crack.CoverDistanceMode": ExpectedConfigDropdownOptions = "NearestContour|GlobalExtreme"
        Case "AutoCAD.Export.LabelMode": ExpectedConfigDropdownOptions = "ValuesOnly|NamesAndValues"
        Case "AutoCAD.Export.PrincipalAxesMode", "Plot.PrincipalAxesMode": ExpectedConfigDropdownOptions = "Transformed|Concrete|None"
        Case "Plot.LegendMode": ExpectedConfigDropdownOptions = "Separate|Common"
        Case "General.ExecutionReportEnabled", "General.NonCriticalMessagesEnabled", "General.DiagramExtension", "Solver.LineSearchEnabled", _
                "AutoCAD.Export.NeutralLineEnabled", "AutoCAD.Export.LoadPointEnabled", "AutoCAD.Export.ContourEnabled", _
                "AutoCAD.Export.ExportCrackInteractionContour", "Plot.Enabled", "Plot.AutoUpdateAfterCalculation", _
                "Plot.ResultGradient", "Plot.ResultLabelsEnabled", "Plot.NeutralLineEnabled", "Plot.LoadApplicationPointEnabled", _
                "Plot.AxisLabelsEnabled", "Plot.ContourEnabled", "Plot.LegendEnabled", "Calculation.Strength.DirectState", _
                "Calculation.Strength.Capacity", "Calculation.Crack.Width", "Calculation.Stability.Enabled", "Enabled", "LineEnabled"
            ExpectedConfigDropdownOptions = "Yes|No"
        Case "Length": ExpectedConfigDropdownOptions = "mm|cm|m"
        Case "Area": ExpectedConfigDropdownOptions = "mm2|cm2|m2"
        Case "Force": ExpectedConfigDropdownOptions = "N|kN|tf"
        Case "Moment": ExpectedConfigDropdownOptions = "N*mm|kN*m|tf*m"
        Case "Stress": ExpectedConfigDropdownOptions = "Pa|kPa|MPa|kgf/cm2|tf/m2"
        Case "Curvature": ExpectedConfigDropdownOptions = "1/mm|1/m"
        Case "+N": ExpectedConfigDropdownOptions = "Tension|Compression"
        Case "+Mx": ExpectedConfigDropdownOptions = "+Y tension|-Y tension"
        Case "+My": ExpectedConfigDropdownOptions = "+X tension|-X tension"
        Case "MaterialModel.Stability.ValueSet", "MaterialModel.Strength.ValueSet", "MaterialModel.CrackInitiation.ValueSet", "MaterialModel.CrackedState.ValueSet"
            ExpectedConfigDropdownOptions = "ULS(I)|SLS(II)"
        Case "MaterialModel.Strength.ConcreteDiagram", "MaterialModel.CrackInitiation.ConcreteDiagram", "MaterialModel.CrackedState.ConcreteDiagram", _
                "MaterialModel.Strength.SteelDiagram", "MaterialModel.CrackInitiation.SteelDiagram", "MaterialModel.CrackedState.SteelDiagram"
            ExpectedConfigDropdownOptions = "TwoLine|ThreeLine"
        Case "MaterialModel.Strength.ConcreteTension": ExpectedConfigDropdownOptions = "Ignore|UseDiagram"
        Case "Visualization.State": ExpectedConfigDropdownOptions = "StrengthState|CapacityState|PreCrackState|PostCrackState|CrackedState"
        Case "Visualization.Quantity": ExpectedConfigDropdownOptions = "Stress|Strain"
        Case "Placement": ExpectedConfigDropdownOptions = "Outside|Inside"
        Case "TextUnits": ExpectedConfigDropdownOptions = "mm|pt"
        Case "ArrowType": ExpectedConfigDropdownOptions = "Triangle|Stealth|Diamond|Oval|Open"
        Case "ArrowSize": ExpectedConfigDropdownOptions = "Small|Medium|Wide"
        Case "Rebar.Loc2row", "Rebar.Loc3row": ExpectedConfigDropdownOptions = "Stacked|SideBySide"
        Case "RectSet.SectionType": ExpectedConfigDropdownOptions = "Rectangle|LSection|TwoRectangles"
    End Select
End Function

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
            AssertTrue stats, prefix & ".rangeSource", Left$(formula, 1) = "="
            AssertTrue stats, prefix & ".visible", cell.Validation.InCellDropdown
            If Left$(formula, 1) = "=" And Not checked.Exists(formula) Then
                checked.Add formula, True
                Set source = sheet.Range(Mid$(formula, 2))
                AssertTrue stats, prefix & ".vertical", source.Columns.Count = 1
                Set sample = cell
                saved = sample.Formula
                options = 0
                For Each optionCell In source.Cells
                    If Not IsError(optionCell.Value2) Then
                        If Len(CStr(optionCell.Value2)) > 0 Then
                            options = options + 1
                            sample.Value2 = optionCell.Value2
                            AssertTrue stats, prefix & ".accepts." & CStr(options), sample.Validation.Value
                        End If
                    Else
                        AssertTrue stats, prefix & ".sourceError", False
                    End If
                Next optionCell
                sample.Value2 = "__INVALID_DROPDOWN_OPTION__"
                AssertTrue stats, prefix & ".rejectsUnknown", Not sample.Validation.Value
                sample.Formula = saved
                Set sample = Nothing
                AssertTrue stats, prefix & ".notEmpty", options > 0
                lists = lists + 1
            End If
        End If
NextCell:
    Next cell
    AssertTrue stats, "config.dropdown.targetsPresent", targets > 0
    AppendLine stats, "CONFIG_DROPDOWNS: targets=" & CStr(targets) & "; rangeSources=" & CStr(lists)
    Exit Sub
Failed:
    Dim code As Long, reason As String
    code = Err.Number: reason = Err.Description
    On Error Resume Next
    If Not sample Is Nothing Then sample.Formula = saved
    On Error GoTo 0
    AssertTrue stats, "config.dropdown.runtime." & CStr(code) & "." & reason, False
End Sub

' Проверяет правую грань каждой строки общей таблицы, включая последнюю
' колонку Справка. Адреса берутся из имени, а не из фиксированных букв/строк.
Private Sub TestConfigSettingsRightBorder(ByRef stats As TUiTestStats)
    Dim settings As Object, cell As Object, row As Long
    On Error GoTo Failed
    Set settings = ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange
    For row = 1 To settings.Rows.Count
        Set cell = settings.Cells(row, settings.Columns.Count)
        If cell.MergeCells Then Set cell = cell.MergeArea
        AssertTrue stats, "config.settings.rightBorder." & CStr(row), _
            cell.Borders(xlEdgeRight).LineStyle = xlDash And _
            cell.Borders(xlEdgeRight).Weight = xlThin And _
            cell.Borders(xlEdgeRight).Color = RGB(0, 0, 0)
    Next row
    Exit Sub
Failed:
    AssertTrue stats, "config.settings.rightBorder.runtime." & CStr(Err.Number) & "." & Err.Description, False
End Sub
