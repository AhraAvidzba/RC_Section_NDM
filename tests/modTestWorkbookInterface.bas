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

' Запускает связанный набор операций и возвращает пользователю итоговый статус выполнения.
Public Function RunWorkbookInterfaceTests() As String
    On Error GoTo Failed

    Dim stats As TUiTestStats
    Dim t0 As Double
    t0 = Timer

    TestButtons stats
    TestLoadCombinationsOnConfig stats
    TestSingleCombinationSkipsBlankRows stats
    TestPartialCombinationIsInvalid stats
    TestInvalidCalculationTypeDoesNotRunPlot stats
    TestAutoCADSourceRequiresManualImport stats
    TestAutoCADImportButtonRejectsGeneratedSource stats
    TestAutoCADPreviewWritesAndDrawsBoundsDimensions stats
    TestGeneratedSourceDoesNotReuseAutoCADPreview stats
    TestGeneratedDirectStateWorstStillDrawsFirstCalculatedLC stats
    TestCapacityOnlyDrawsGeometryWithoutStateResults stats
    TestAutoCADCalculationMessageUsesSavedGeometry stats
    TestBlankMomentDefaultsToZeroAndZeroLoadsAreSkipped stats
    TestCircleWorkbookRunWritesResults stats
    TestExecutionReportFile stats
    TestExcelApplicationStateGuardRestoresSettings stats
    TestLShapeWorkbookRunWritesResults stats
    TestLShapeMomentUltimateStrainWorkbookPath stats
    TestLShapePureBendingUltimateStrainWorkbookPath stats
    TestLShapePureBendingDirectStateWorkbookPath stats
    TestLShapeAxialTensionExtensionFromWorkbookSettings stats
    TestAutoCADExportUsesSharedLoadReference stats
    TestGoverningCombinationWritesDetailedResults stats
    TestCapacitySearchMethodValidation stats
    TestSolverToleranceUnitLabels stats
    TestCapacitySettingsUnitLabels stats
    TestBlankDiametersDisableGeneratedRebars stats
    TestClearResultsKeepsInputs stats

    AppendLine stats, "TOTAL_WORKBOOK_UI: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed) & _
        "; elapsedSec=" & FormatNumberInvariant(Timer - t0)
    RunWorkbookInterfaceTests = stats.Report
    Exit Function

Failed:
    RunWorkbookInterfaceTests = stats.Report & "RUNTIME ERROR: " & CStr(Err.Number) & _
        "; source=" & Err.Source & "; description=" & Err.Description
End Function

' Проверяет, что таблица сочетаний живет на Config рядом с настройками,
' и не выводится отдельной параллельной таблицей на лист Расчет.
Private Sub TestLoadCombinationsOnConfig(ByRef stats As TUiTestStats)
    Dim loads As Object
    Set loads = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange

    AssertTrue stats, "ui.loads.layout.sheet", loads.Worksheet.Name = "Config"
    AssertTrue stats, "ui.loads.layout.position", loads.Row = 3 And loads.Column = 15
    AssertTrue stats, "ui.loads.layout.title", _
        InStr(1, CStr(loads.Worksheet.Cells.Item(2, 15).Value2), "Сочетания нагрузок", vbTextCompare) > 0
    Dim pathOptions As Variant
    pathOptions = Array(ChrW$(&H3BB) & "*Mx", ChrW$(&H3BB) & "*My", _
        ChrW$(&H3BB) & "*Mxy", ChrW$(&H3BB) & "*N", ChrW$(&H3BB) & "*NMxy")
    AssertTrue stats, "ui.loads.layout.capacityPathValidation", _
        LoadCombinationValidationHasOptions(6, pathOptions)
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
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

Private Function ButtonTextIsCentered(ByVal buttonShape As Object) As Boolean
    ButtonTextIsCentered = (buttonShape.TextFrame.HorizontalAlignment = -4108 And _
        buttonShape.TextFrame.VerticalAlignment = -4108)
End Function

' Проверяет пользовательское правило: пустой диаметр арматуры означает
' отсутствие ряда, а не подстановку типового default-диаметра из кода.
Private Sub TestBlankDiametersDisableGeneratedRebars(ByRef stats As TUiTestStats)
    PrepareCircleInput
    SetSystemSetting "Rebar.Diameter", vbNullString

    Dim circleSection As CSectionModel
    Set circleSection = BuildCurrentWorkbookSection()
    AssertTrue stats, "ui.circle.blankDiameter.noRebar", circleSection.RebarCount = 0
    AssertTrue stats, "ui.circle.blankDiameter.dimensionsRemain", circleSection.AnnotationCount >= 2

    PrepareLShapeInput
    SetSystemSetting "LShape.H1.d_1", vbNullString

    Dim lshapeSection As CSectionModel
    Set lshapeSection = BuildCurrentWorkbookSection()
    AssertTrue stats, "ui.lshape.blankFaceDiameter.skipsLine", lshapeSection.RebarCount = 7
    AssertTrue stats, "ui.lshape.blankFaceDiameter.dimensionsRemain", lshapeSection.AnnotationCount >= 4

    SetSystemSetting "Geometry.Type", "RoundedRectangle"
    Dim roundedSection As CSectionModel
    Set roundedSection = BuildCurrentWorkbookSection()
    AssertTrue stats, "ui.rounded.noAutoRebar.noError", roundedSection.RebarCount = 0
    AssertTrue stats, "ui.rounded.noAutoRebar.dimensionsRemain", roundedSection.AnnotationCount >= 2

    PrepareCircleInput
End Sub

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

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestAutoCADExportUsesSharedLoadReference(ByRef stats As TUiTestStats)
    PrepareLShapeInput
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

    Dim reader As CLoadCombinationReader
    Set reader = New CLoadCombinationReader
    reader.LoadFromWorkbook ThisWorkbook, batch, units
    batch.ApplyLoadReference props.CentroidX, props.CentroidY, props.CentroidX, props.CentroidY
    batch.Execute

    AssertTrue stats, "ui.autocad.reference.converged", batch.StateConverged(1)
    AssertTrue stats, "ui.autocad.reference.point", Abs(batch.LoadReferenceX) > 0.000001 Or Abs(batch.LoadReferenceY) > 0.000001
    AssertClose stats, "ui.autocad.reference.concreteCenterX", props.CentroidX, batch.LoadReferenceX, 0.000001
    AssertClose stats, "ui.autocad.reference.concreteCenterY", props.CentroidY, batch.LoadReferenceY, 0.000001
    AssertClose stats, "ui.autocad.reference.mxTransfer", batch.Mx(1), batch.N(1) * props.CentroidY, 0.000001
    AssertClose stats, "ui.autocad.reference.myTransfer", batch.My(1), batch.N(1) * props.CentroidX, 0.000001
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
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

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
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
    AssertTrue stats, "ui.loads.partial.invalid", batch.Status(1) = "InputErr"

    Dim message As String
    message = RunSectionCalculationForWorkbook(ThisWorkbook, False)
    AssertTrue stats, "ui.loads.partial.message", InStr(1, message, "ошиб", vbTextCompare) > 0 And _
        InStr(1, message, "InputErr", vbTextCompare) > 0
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestInvalidCalculationTypeDoesNotRunPlot(ByRef stats As TUiTestStats)
    PrepareCircleInput
    SetSystemSetting "Plot.AutoUpdateAfterCalculation", "Yes"

    Dim loads As Object
    Set loads = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange
    loads.Cells.Item(2, 5).ClearContents

    Dim message As String
    message = RunSectionCalculationForWorkbook(ThisWorkbook, False)

    AssertTrue stats, "ui.loads.invalidCalculationType.message", _
        InStr(1, message, "InputErr", vbTextCompare) > 0 And _
        InStr(1, message, "результатов элементов", vbTextCompare) = 0
    AssertTrue stats, "ui.loads.invalidCalculationType.noElementRows", _
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
    section.AddRebarElement 50#, 50#, 20#, 0#, "Ribbed"

    Dim writer As CNDMResultsWriter
    Set writer = New CNDMResultsWriter
    writer.WriteGeometryPreview ThisWorkbook, section

    Dim annotationData As Variant
    annotationData = ResultTable("rngNDMSectionAnnotations")
    AssertTrue stats, "ui.autocad.preview.boundsAnnotations", CountAnnotationType(annotationData, "DIMENSION") = 2
    AssertTrue stats, "ui.autocad.preview.approxText", _
        InStr(1, CStr(annotationData(2, ResultHeaderColumn(annotationData, "Text"))), ChrW$(&H2248), vbTextCompare) > 0
    AssertTrue stats, "ui.autocad.preview.russianComment", _
        InStr(1, CStr(annotationData(2, ResultHeaderColumn(annotationData, "Comment"))), "Приблизительная", vbTextCompare) > 0

    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook

    Dim reader As CSectionPlotDataReader
    Set reader = New CSectionPlotDataReader
    reader.LoadGeometryPreviewFromWorkbook ThisWorkbook, settings
    AssertTrue stats, "ui.autocad.preview.readerAnnotations", reader.AnnotationCount = 2

    UpdateSectionPlotForWorkbook ThisWorkbook
    AssertTrue stats, "ui.autocad.preview.dimensionShapes", CountPlotShapes("AnnotationLine") > 0
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
    section.AddRebarElement 50#, 50#, 20#, 0#, "Ribbed"

    Dim writer As CNDMResultsWriter
    Set writer = New CNDMResultsWriter
    writer.WriteGeometryPreview ThisWorkbook, section

    UpdateSectionPlotForWorkbook ThisWorkbook
    AssertTrue stats, "ui.plot.preview.title", PlotChartTitleContains("Импортированная геометрия AutoCAD")

    SetSystemSetting "Geometry.Source", "Generated"
    Dim errorDescription As String
    On Error Resume Next
    UpdateSectionPlotForWorkbook ThisWorkbook
    errorDescription = Err.Description
    On Error GoTo 0

    AssertTrue stats, "ui.plot.generated.noPreviewFallback.error", _
        InStr(1, errorDescription, "Выполните расчет", vbTextCompare) > 0
    AssertTrue stats, "ui.plot.generated.noPreviewFallback.title", _
        Not PlotChartTitleContains("Импортированная геометрия AutoCAD")
End Sub

' Проверяет DirectState-сценарий без определяющего сочетания по прочности.
' При Plot.LoadCase = Worst схема должна показать первый рассчитанный LC из
' Results, а не оставлять AutoCAD-preview и не очищаться до пустого окна.
Private Sub TestGeneratedDirectStateWorstStillDrawsFirstCalculatedLC(ByRef stats As TUiTestStats)
    PrepareCircleInput
    SetSystemSetting "Geometry.Source", "Generated"
    SetSystemSetting "Calculation.Mode", "DirectState"
    SetSystemSetting "Plot.LoadCase", "Worst"

    Dim message As String
    message = RunSectionCalculationForWorkbook(ThisWorkbook, False)

    AssertTrue stats, "ui.plot.generatedWorst.message", InStr(1, message, "Расчет завершен", vbTextCompare) > 0
    AssertTrue stats, "ui.plot.generatedWorst.drawsCalculatedLc", PlotChartTitleContains("LC1")
    AssertTrue stats, "ui.plot.generatedWorst.noImportTitle", _
        Not PlotChartTitleContains("Импортированная геометрия AutoCAD")
End Sub

' Проверяет режим CapacityOnly на уровне книги.
' Расчет сохраняет геометрию и capacity-результаты, но не записывает
' поэлементные Stress/Strain. Схема должна строиться как geometry-only,
' чтобы Excel и AutoCAD показывали согласованную картину без НДС.
Private Sub TestCapacityOnlyDrawsGeometryWithoutStateResults(ByRef stats As TUiTestStats)
    PrepareCircleInput
    SetSystemSetting "Calculation.Mode", "CapacityOnly"
    SetSystemSetting "Plot.AutoUpdateAfterCalculation", "Yes"
    SetSystemSetting "Plot.LoadCase", "Worst"

    Dim message As String
    message = RunSectionCalculationForWorkbook(ThisWorkbook, False)

    AssertTrue stats, "ui.capacityOnly.message", InStr(1, message, "Расчет завершен", vbTextCompare) > 0
    AssertTextEquals stats, "ui.capacityOnly.directNA", _
        CStr(ThisWorkbook.Worksheets.Item("Results").Cells.Item(BatchSummaryStartRow() + 9, _
        BatchSummaryColumnByHeader("DirectStateStatus")).Value2), "N/A"
    AssertTrue stats, "ui.capacityOnly.noElementStateRows", ResultTableRowCount("rngNDMElementResults") = 1
    AssertTrue stats, "ui.capacityOnly.plotGeometryTitle", PlotChartTitleContains("CapacityOnly: геометрия без НДС")
    AssertTrue stats, "ui.capacityOnly.plotNoLoadCaseTitle", Not PlotChartTitleContains("при загружении")
    AssertTrue stats, "ui.capacityOnly.plotNoImportTitle", _
        Not PlotChartTitleContains("Импортированная геометрия AutoCAD")
    AssertTrue stats, "ui.capacityOnly.commonLoadReference", _
        Len(ResultsPropertyValue("ALL", "LoadReferenceX")) > 0 And Len(ResultsPropertyValue("ALL", "LoadReferenceY")) > 0

    SetSystemSetting "Calculation.Mode", "FullCapacity"
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
    Set section = BuildWorkbookSectionModel(ThisWorkbook, settings, units)
    section.SourceType = "AutoCADImport"
    section.Annotations.Clear

    Dim writer As CNDMResultsWriter
    Set writer = New CNDMResultsWriter
    writer.WriteGeometryPreview ThisWorkbook, section, units

    SetSystemSetting "Geometry.Source", "AutoCAD"
    ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange.Cells.Item(2, 5).Value2 = "Group1"

    Dim message As String
    message = RunSectionCalculationForWorkbook(ThisWorkbook, False)

    AssertTrue stats, "ui.autocad.run.savedGeometry.message", _
        InStr(1, message, "Расчет импортированной из AutoCAD геометрии завершен", vbTextCompare) > 0
    AssertTrue stats, "ui.autocad.run.savedGeometry.noImportText", _
        InStr(1, message, "Импортировано из AutoCAD", vbTextCompare) = 0
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestBlankMomentDefaultsToZeroAndZeroLoadsAreSkipped(ByRef stats As TUiTestStats)
    PrepareCircleInput
    Dim loads As Object
    Set loads = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange
    loads.Cells.Item(2, 4).ClearContents
    loads.Cells.Item(3, 1).Value2 = "ZERO"
    loads.Cells.Item(3, 2).Value2 = 0#
    loads.Cells.Item(3, 3).Value2 = 0#
    loads.Cells.Item(3, 4).Value2 = 0#
    loads.Cells.Item(3, 5).Value2 = "Group1"

    Dim batch As CBatchSectionCalculator
    Set batch = BuildUiBatch()

    Dim reader As CLoadCombinationReader
    Set reader = New CLoadCombinationReader
    reader.LoadFromWorkbook ThisWorkbook, batch
    batch.Execute

    AssertTrue stats, "ui.loads.blankMoment.count", batch.Count = 1
    AssertClose stats, "ui.loads.blankMoment.myZero", batch.UserMy(1), 0#, 0.0000001
    AssertTrue stats, "ui.loads.blankMoment.valid", InStr(1, batch.Status(1), "InputErr", vbTextCompare) = 0
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestCircleWorkbookRunWritesResults(ByRef stats As TUiTestStats)
    PrepareCircleInput
    Dim message As String
    message = RunSectionCalculationForWorkbook(ThisWorkbook, False)

    AssertTrue stats, "ui.run.message", InStr(1, message, "Расчет завершен", vbTextCompare) > 0
    Dim summary As Object
    Set summary = ThisWorkbook.Names.Item("rngBatchSummary").RefersToRange
    Dim firstDataRow As Long
    firstDataRow = BatchSummaryStartRow() + 9
    AssertTrue stats, "ui.run.capacity.na", CStr(summary.Worksheet.Cells.Item(firstDataRow, 16).Value2) = "N/A"
    AssertTrue stats, "ui.run.direct.status", Len(CStr(summary.Worksheet.Cells.Item(firstDataRow, 5).Value2)) > 0
    AssertTrue stats, "ui.run.deformations", IsNumeric(summary.Worksheet.Cells.Item(firstDataRow, 6).Value2) And _
        IsNumeric(summary.Worksheet.Cells.Item(firstDataRow, 9).Value2)
    AssertTrue stats, "ui.run.crack", Len(CStr(summary.Worksheet.Cells.Item(firstDataRow, 28).Value2)) > 0
    Dim sys As Object
    Set sys = ThisWorkbook.Worksheets.Item("Config")
    AssertTrue stats, "ui.run.system.noRebarTable", Len(CStr(sys.Cells.Item(130, 1).Value2)) = 0
    AssertTrue stats, "ui.run.system.materialDiagramControls", _
        InStr(1, CStr(sys.Cells.Item(1, 35).Value2), "Контрольные точки диаграмм", vbTextCompare) > 0
    Dim resultsSheet As Object
    Set resultsSheet = ThisWorkbook.Worksheets.Item("Results")
    Dim summaryRow As Long
    summaryRow = BatchSummaryStartRow()
    AssertTrue stats, "ui.batchSummary.currentDepths", IsNumeric(resultsSheet.Cells.Item(summaryRow + 9, 10).Value2) And IsNumeric(resultsSheet.Cells.Item(summaryRow + 9, 11).Value2)
    AssertTrue stats, "ui.batchSummary.direct.noCapacityDepths", Len(CStr(resultsSheet.Cells.Item(summaryRow + 9, 23).Value2)) = 0 And Len(CStr(resultsSheet.Cells.Item(summaryRow + 9, 24).Value2)) = 0
    AssertTrue stats, "ui.results.elements.header", CStr(ThisWorkbook.Names.Item("rngNDMElementResults").RefersToRange.Value2) = "RunID"
    Dim elementResults As Variant
    elementResults = ResultTable("rngNDMElementResults")
    AssertTrue stats, "ui.results.elements.rows", UBound(elementResults, 1) > 1
    AssertTrue stats, "ui.results.elements.noCombinationIndex", ResultHeaderColumn(elementResults, "CombinationIndex") = 0
    AssertTrue stats, "ui.results.elements.units", ResultHeaderColumn(elementResults, "Stress, MPa") > 0
    AssertTrue stats, "ui.results.elements.noGeometryDup", ResultHeaderColumn(elementResults, "X, mm") = 0
    AssertTrue stats, "ui.results.elements.noPlaneDup", ResultHeaderColumn(elementResults, "Epsilon0") = 0
    AssertTrue stats, "ui.results.elements.physicalStateAlign", _
        ThisWorkbook.Names.Item("rngNDMElementResults").RefersToRange.Offset(1, ResultHeaderColumn(elementResults, "PhysicalState") - 1).HorizontalAlignment = -4152
    AssertTrue stats, "ui.results.geometry.header", CStr(ThisWorkbook.Names.Item("rngNDMSectionGeometry").RefersToRange.Value2) = "RunID"
    Dim geometryResults As Variant
    geometryResults = ResultTable("rngNDMSectionGeometry")
    AssertTrue stats, "ui.results.geometry.rows", UBound(geometryResults, 1) > 1
    AssertTrue stats, "ui.results.geometry.position", ThisWorkbook.Names.Item("rngNDMSectionGeometry").RefersToRange.Row = 32 And ThisWorkbook.Names.Item("rngNDMSectionGeometry").RefersToRange.Column = 9
    AssertTrue stats, "ui.results.properties.position", ThisWorkbook.Names.Item("rngNDMSectionProperties").RefersToRange.Row = 32 And ThisWorkbook.Names.Item("rngNDMSectionProperties").RefersToRange.Column = 26
    AssertTrue stats, "ui.results.annotations.position", ThisWorkbook.Names.Item("rngNDMSectionAnnotations").RefersToRange.Row = 32 And ThisWorkbook.Names.Item("rngNDMSectionAnnotations").RefersToRange.Column = 34
    AssertTrue stats, "ui.results.materialDiagrams.position", ThisWorkbook.Names.Item("rngNDMMaterialDiagrams").RefersToRange.Row = 32 And ThisWorkbook.Names.Item("rngNDMMaterialDiagrams").RefersToRange.Column = 50
    AssertTrue stats, "ui.results.geometry.noSource", ResultHeaderColumn(geometryResults, "SourceName") = 0
    AssertTrue stats, "ui.results.geometry.noMaterialClass", ResultHeaderColumn(geometryResults, "MaterialClass") = 0
    AssertTrue stats, "ui.results.properties.header", CStr(ThisWorkbook.Names.Item("rngNDMSectionProperties").RefersToRange.Value2) = "RunID"
    AssertTrue stats, "ui.results.properties.hasEpsilon0", ResultsPropertyExists("LC1", "Epsilon0")
    AssertTrue stats, "ui.results.properties.hasBounds", ResultsPropertyExists("ALL", "Bounds.MinX")
    AssertTrue stats, "ui.results.properties.commonLoadReference", _
        ResultsPropertyExists("ALL", "LoadReferenceX") And ResultsPropertyExists("ALL", "LoadReferenceY")
    AssertTrue stats, "ui.results.properties.noLcLoadReference", _
        Not ResultsPropertyExists("LC1", "LoadReferenceX") And Not ResultsPropertyExists("LC1", "LoadReferenceY")
    AssertTrue stats, "ui.results.annotations.header", CStr(ThisWorkbook.Names.Item("rngNDMSectionAnnotations").RefersToRange.Value2) = "RunID"
    AssertTrue stats, "ui.results.annotations.rows", ResultTableRowCount("rngNDMSectionAnnotations") > 1
    AssertTrue stats, "ui.results.materialDiagrams.header", CStr(ThisWorkbook.Names.Item("rngNDMMaterialDiagrams").RefersToRange.Value2) = "RunID"
    AssertTrue stats, "ui.results.materialDiagrams.rows", ResultTableRowCount("rngNDMMaterialDiagrams") > 1
    AssertTrue stats, "ui.plot.chart.created", PlotChartExists()
    AssertTrue stats, "ui.plot.title.comment", PlotChartTitleContains("(ui test)")
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestLShapeWorkbookRunWritesResults(ByRef stats As TUiTestStats)
    PrepareLShapeInput
    Dim message As String
    message = RunSectionCalculationForWorkbook(ThisWorkbook, False)

    AssertTrue stats, "ui.lshape.message", InStr(1, message, "завершен", vbTextCompare) > 0
    Dim summary As Object
    Set summary = ThisWorkbook.Names.Item("rngBatchSummary").RefersToRange
    AssertTrue stats, "ui.lshape.result.status", Len(CStr(summary.Worksheet.Cells.Item(BatchSummaryStartRow() + 9, 1).Value2)) > 0

    Dim sys As Object
    Set sys = ThisWorkbook.Worksheets.Item("Config")
    AssertTrue stats, "ui.lshape.system.noRebarTable", Len(CStr(sys.Cells.Item(130, 1).Value2)) = 0
    AssertTrue stats, "ui.lshape.system.materialDiagramControls", _
        InStr(1, CStr(sys.Cells.Item(1, 35).Value2), "Контрольные точки диаграмм", vbTextCompare) > 0
End Sub

' Проверяет пользовательский сценарий Г-сечения с N + Mx и выбранной
' траекторией lambda*Mx через полный путь книги. Этот тест защищает быстрый
' UltimateStrain-путь для обычного изгибного расчета: после универсализации
' CapacityLoadPath он не должен уходить в тяжелую общую residual-систему.
Private Sub TestLShapeMomentUltimateStrainWorkbookPath(ByRef stats As TUiTestStats)
    PrepareUserLShapeMomentUltimateInput

    Dim message As String
    message = RunSectionCalculationForWorkbook(ThisWorkbook, False)
    ThisWorkbook.Application.CalculateFull

    Dim resultsSheet As Object
    Set resultsSheet = ThisWorkbook.Worksheets.Item("Results")
    Dim firstRow As Long
    firstRow = BatchSummaryStartRow() + 9

    Dim capacityStatus As String
    Dim solutionMethod As String
    capacityStatus = CStr(resultsSheet.Cells.Item(firstRow, BatchSummaryColumnByHeader("CapacityStatus")).Value2)
    solutionMethod = CStr(resultsSheet.Cells.Item(firstRow, BatchSummaryColumnByHeader("CapacitySolutionMethod")).Value2)

    AppendLine stats, "INFO: ui.lshape.momentUltimate capacityStatus=" & capacityStatus & _
        "; solutionMethod=" & solutionMethod & _
        "; lambda=" & CStr(resultsSheet.Cells.Item(firstRow, BatchSummaryColumnByHeader("lambdaUltimate")).Value2)

    AssertTrue stats, "ui.lshape.momentUltimate.message", InStr(1, message, "Расчет завершен", vbTextCompare) > 0
    AssertTextEquals stats, "ui.lshape.momentUltimate.capacityOk", capacityStatus, "OK"
    AssertTextEquals stats, "ui.lshape.momentUltimate.method", solutionMethod, "UltimateStrain"
    AssertTrue stats, "ui.lshape.momentUltimate.lambda", _
        CDbl(resultsSheet.Cells.Item(firstRow, BatchSummaryColumnByHeader("lambdaUltimate")).Value2) > 0#
    AssertTrue stats, "ui.lshape.momentUltimate.mxult", _
        Abs(CDbl(resultsSheet.Cells.Item(firstRow, BatchSummaryColumnByHeader("Mxult")).Value2)) > 0#
End Sub

' Проверяет чистый изгиб Г-сечения по полному Excel-пути.
' Здесь N намеренно равен нулю: программа должна найти прямое НДС и
' предельный момент без перехода в осевую или аварийную ветку.
Private Sub TestLShapePureBendingUltimateStrainWorkbookPath(ByRef stats As TUiTestStats)
    PrepareUserLShapeMomentUltimateInput

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
    firstRow = BatchSummaryStartRow() + 9

    Dim directStatus As String
    Dim capacityStatus As String
    Dim solutionMethod As String
    directStatus = CStr(resultsSheet.Cells.Item(firstRow, BatchSummaryColumnByHeader("DirectStateStatus")).Value2)
    capacityStatus = CStr(resultsSheet.Cells.Item(firstRow, BatchSummaryColumnByHeader("CapacityStatus")).Value2)
    solutionMethod = CStr(resultsSheet.Cells.Item(firstRow, BatchSummaryColumnByHeader("CapacitySolutionMethod")).Value2)

    AppendLine stats, "INFO: ui.lshape.pureBending direct=" & directStatus & _
        "; capacity=" & capacityStatus & _
        "; solutionMethod=" & solutionMethod & _
        "; lambda=" & CStr(resultsSheet.Cells.Item(firstRow, BatchSummaryColumnByHeader("lambdaUltimate")).Value2)

    AssertTrue stats, "ui.lshape.pureBending.message", InStr(1, message, "Расчет завершен", vbTextCompare) > 0
    AssertTextEquals stats, "ui.lshape.pureBending.directOk", directStatus, "OK"
    AssertTextEquals stats, "ui.lshape.pureBending.capacityOk", capacityStatus, "OK"
    AssertTextEquals stats, "ui.lshape.pureBending.method", solutionMethod, "UltimateStrain"
    AssertTrue stats, "ui.lshape.pureBending.lambda", _
        CDbl(resultsSheet.Cells.Item(firstRow, BatchSummaryColumnByHeader("lambdaUltimate")).Value2) > 0#
End Sub

' Проверяет чистый изгиб Г-сечения без расчета несущей способности.
' Это защищает именно прямой StateSolution: даже если FullCapacity выключен,
' solver должен найти НДС для простого Mx при N=0, а не зависеть от ранее
' найденной предельной capacity-плоскости.
Private Sub TestLShapePureBendingDirectStateWorkbookPath(ByRef stats As TUiTestStats)
    PrepareUserLShapeMomentUltimateInput
    SetSystemSetting "Calculation.Mode", "DirectState"

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
    firstRow = BatchSummaryStartRow() + 9

    Dim directStatus As String
    Dim capacityStatus As String
    directStatus = CStr(resultsSheet.Cells.Item(firstRow, BatchSummaryColumnByHeader("DirectStateStatus")).Value2)
    capacityStatus = CStr(resultsSheet.Cells.Item(firstRow, BatchSummaryColumnByHeader("CapacityStatus")).Value2)

    AppendLine stats, "INFO: ui.lshape.pureBendingDirect direct=" & directStatus & _
        "; capacity=" & capacityStatus

    AssertTrue stats, "ui.lshape.pureBendingDirect.message", InStr(1, message, "Расчет завершен", vbTextCompare) > 0
    AssertTextEquals stats, "ui.lshape.pureBendingDirect.directOk", directStatus, "OK"
    AssertTextEquals stats, "ui.lshape.pureBendingDirect.capacityNA", capacityStatus, "N/A"
End Sub

' Проверяет пользовательский сценарий с сильным осевым растяжением Г-сечения
' через настоящий workbook-path: Config -> CUnitSystem -> batch -> Results.
' Это важно, потому что знак N и единицы tf здесь проходят ровно тем же путем,
' что и при нажатии кнопки "Расчет" в книге.
Private Sub TestLShapeAxialTensionExtensionFromWorkbookSettings(ByRef stats As TUiTestStats)
    PrepareUserLShapeAxialTensionInput

    Dim reportPath As String
    reportPath = ThisWorkbook.Path & "\RC_Section_NDM_execution_report.txt"
    DeleteFileIfExists reportPath

    Dim message As String
    message = RunSectionCalculationForWorkbook(ThisWorkbook, False)

    Dim resultsSheet As Object
    Set resultsSheet = ThisWorkbook.Worksheets.Item("Results")
    Dim firstRow As Long
    firstRow = BatchSummaryStartRow() + 9

    Dim safeOverall As String
    Dim safeDirect As String
    Dim overOverall As String
    Dim overDirect As String
    Dim overCrack As String
    Dim overExtension As String
    safeOverall = CStr(resultsSheet.Cells.Item(firstRow, 1).Value2)
    safeDirect = CStr(resultsSheet.Cells.Item(firstRow, 5).Value2)
    overOverall = CStr(resultsSheet.Cells.Item(firstRow + 1, 1).Value2)
    overDirect = CStr(resultsSheet.Cells.Item(firstRow + 1, 5).Value2)
    overCrack = CStr(resultsSheet.Cells.Item(firstRow + 1, BatchSummaryColumnByHeader("CrackStatus")).Value2)
    overExtension = ResultsPropertyValue("LC_OVER", "ExtensionUsed")

    AppendLine stats, "INFO: ui.lshape.axial795 overall=" & safeOverall & _
        "; direct=" & safeDirect
    AppendLine stats, "INFO: ui.lshape.axial900 overall=" & overOverall & _
        "; direct=" & overDirect & "; crack=" & overCrack & _
        "; extensionUsed=" & overExtension

    AssertTextEquals stats, "ui.lshape.axial795.directOk", safeDirect, "OK"

    AssertTextEquals stats, "ui.lshape.axial900.fail", overOverall, "FAIL"
    AssertTextEquals stats, "ui.lshape.axial900.directFail", overDirect, "FAIL"
    AssertTextEquals stats, "ui.lshape.axial900.crackSkipped", overCrack, "N/A"
    AssertTrue stats, "ui.lshape.axial900.message", InStr(1, message, "Расчет завершен", vbTextCompare) > 0
    AssertTextEquals stats, "ui.lshape.axial900.extensionSnapshot", overExtension, "True"
    AssertTrue stats, "ui.lshape.axial900.reportCreated", FileExists(reportPath)

    SetSystemSetting "General.ExecutionReportEnabled", "No"
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestGoverningCombinationWritesDetailedResults(ByRef stats As TUiTestStats)
    PrepareCircleInput
    SetSystemSetting "Calculation.Mode", "FullCapacity"
    SetSystemSetting "Capacity.SolutionStrategy", "LoadMultiplier"
    SetSystemSetting "Capacity.ToleranceLambda", "0.05"
    SetSystemSetting "Capacity.MaxLambda", "10"
    SetSystemSetting "Plot.LoadCase", "Worst"

    Dim loads As Object
    Set loads = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange
    loads.Cells.Item(2, 1).Value2 = "LC_SAFE"
    loads.Cells.Item(2, 2).Value2 = -100000#
    loads.Cells.Item(2, 3).Value2 = -1000000#
    loads.Cells.Item(2, 4).Value2 = 0#
    loads.Cells.Item(2, 5).Value2 = "Group1"
    loads.Cells.Item(2, 6).Value2 = ChrW$(&H3BB) & "*Mxy"
    loads.Cells.Item(2, 7).Value2 = "less severe"

    loads.Cells.Item(3, 1).Value2 = "LC_GOV"
    loads.Cells.Item(3, 2).Value2 = -100000#
    loads.Cells.Item(3, 3).Value2 = -8000000#
    loads.Cells.Item(3, 4).Value2 = 0#
    loads.Cells.Item(3, 5).Value2 = "Group1"
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
    governingID = CStr(resultsSheet.Cells.Item(summaryRow + 1, 5).Value2)
    Dim expectedID As String
    expectedID = ExpectedGoverningByLowestStrengthSafety(resultsSheet, summaryRow)

    AssertTrue stats, "ui.governing.id", governingID = expectedID
    AssertTrue stats, "ui.governing.message", InStr(1, message, governingID, vbTextCompare) > 0
End Sub

Private Function ExpectedGoverningByLowestStrengthSafety(ByVal resultsSheet As Object, ByVal summaryRow As Long) As String
    Dim rowIndex As Long
    Dim bestSafety As Double
    For rowIndex = summaryRow + 9 To summaryRow + 28
        If Len(Trim$(CStr(resultsSheet.Cells.Item(rowIndex, 2).Value2))) > 0 Then
            Dim safetyValue As Double
            safetyValue = StrengthSafetyForSummaryRow(resultsSheet, rowIndex)
            If safetyValue > 0# And (bestSafety = 0# Or safetyValue < bestSafety) Then
                bestSafety = safetyValue
                ExpectedGoverningByLowestStrengthSafety = CStr(resultsSheet.Cells.Item(rowIndex, 2).Value2)
            End If
        End If
    Next rowIndex
End Function

Private Function StrengthSafetyForSummaryRow(ByVal resultsSheet As Object, ByVal rowIndex As Long) As Double
    If IsNumeric(resultsSheet.Cells.Item(rowIndex, 25).Value2) Then _
        StrengthSafetyForSummaryRow = CDbl(resultsSheet.Cells.Item(rowIndex, 25).Value2)
End Function

Private Function MomentUltimateForCombination(ByVal resultsSheet As Object, ByVal summaryRow As Long, ByVal combinationID As String) As Double
    Dim rowIndex As Long
    For rowIndex = summaryRow + 9 To summaryRow + 28
        If StrComp(CStr(resultsSheet.Cells.Item(rowIndex, 2).Value2), combinationID, vbTextCompare) = 0 Then
            Dim mxUltimate As Double
            Dim myUltimate As Double
            If IsNumeric(resultsSheet.Cells.Item(rowIndex, 21).Value2) Then mxUltimate = CDbl(resultsSheet.Cells.Item(rowIndex, 21).Value2)
            If IsNumeric(resultsSheet.Cells.Item(rowIndex, 22).Value2) Then myUltimate = CDbl(resultsSheet.Cells.Item(rowIndex, 22).Value2)
            MomentUltimateForCombination = Sqr(mxUltimate * mxUltimate + myUltimate * myUltimate)
            Exit Function
        End If
    Next rowIndex
End Function

Private Function BatchSummaryStartRow() As Long
    BatchSummaryStartRow = ThisWorkbook.Names.Item("rngBatchSummary").RefersToRange.Row
End Function

' Ищет колонку rngBatchSummary по началу текста заголовка.
' В сводке часто добавляются новые расчетные поля, поэтому UI-тесты не
' должны зависеть от номера столбца: проверяем именно смысловую
' колонку, которую видит пользователь.
Private Function BatchSummaryColumnByHeader(ByVal headerPrefix As String) As Long
    Dim summary As Object
    Set summary = ThisWorkbook.Names.Item("rngBatchSummary").RefersToRange

    Dim colIndex As Long
    For colIndex = 1 To summary.Columns.Count
        If InStr(1, CStr(summary.Cells.Item(9, colIndex).Value2), headerPrefix, vbTextCompare) = 1 Then
            BatchSummaryColumnByHeader = colIndex
            Exit Function
        End If
    Next colIndex
End Function

Private Function ResultHeaderColumn(ByRef data As Variant, ByVal headerText As String) As Long
    Dim colIndex As Long
    For colIndex = 1 To UBound(data, 2)
        If StrComp(CStr(data(1, colIndex)), headerText, vbTextCompare) = 0 Then
            ResultHeaderColumn = colIndex
            Exit Function
        End If
    Next colIndex
End Function

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestCapacitySearchMethodValidation(ByRef stats As TUiTestStats)
    AssertTrue stats, "ui.validation.calculationMode", _
        SystemSettingValidationHasOptions("Calculation.Mode", Array("DirectState", "FullCapacity", "CapacityOnly"))
    AssertTrue stats, "ui.validation.CapacitySolutionStrategy", _
        SystemSettingValidationHasOptions("Capacity.SolutionStrategy", Array("Auto", "UltimateStrain", "LoadMultiplier"))
    AssertTrue stats, "ui.validation.capacitySearchMethod", _
        SystemSettingValidationHasOptions("Capacity.SearchMethod", Array("Bisection", "Brent", "Secant"))
    AssertTrue stats, "ui.validation.capacityScope", _
        SystemSettingValidationHasOptions("Capacity.CalculationScope", Array("Group1Only", "Group1+2"))
    AssertTrue stats, "ui.validation.autocadLabelMode", _
        SystemSettingValidationHasOptions("AutoCAD.Export.LabelMode", Array("ValuesOnly", "NamesAndValues"))
    AssertTrue stats, "ui.validation.autocadResultType", _
        SystemSettingValidationHasOptions("AutoCAD.Export.ResultType", Array("Stress", "Strain"))
    AssertTrue stats, "ui.validation.autocadNeutralLine", _
        SystemSettingValidationHasOptions("AutoCAD.Export.NeutralLineEnabled", Array("Yes", "No"))
    AssertTrue stats, "ui.validation.autocadPrincipalAxes", _
        SystemSettingValidationHasOptions("AutoCAD.Export.PrincipalAxesEnabled", Array("Yes", "No"))
    AssertTrue stats, "ui.validation.autocadLoadPoint", _
        SystemSettingValidationHasOptions("AutoCAD.Export.LoadPointEnabled", Array("Yes", "No"))
    AssertTrue stats, "ui.validation.autocadCombination", AutoCADCombinationValidationIsDynamic()
    AssertTrue stats, "ui.validation.plotLoadCase", PlotLoadCaseValidationIsDynamic()
    AssertTrue stats, "ui.validation.loadCalculationType", _
        LoadCombinationValidationHasOptions(5, Array("Group1", "Group2"))
    AssertTrue stats, "ui.validation.plotResultType", _
        SystemSettingValidationHasOptions("Plot.ResultType", Array("Stress", "Strain"))
    AssertTrue stats, "ui.validation.plotLabels", _
        SystemSettingValidationHasOptions("Plot.ResultLabelsEnabled", Array("Yes", "No"))
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
    SetSystemSetting "Geometry.Type", "LShape"
    AssertTrue stats, "ui.validation.lshapeLoc2row", _
        LShapeAdditionalValidationHasOptions(16, 3, Array("Stacked", "SideBySide"))
    AssertTrue stats, "ui.validation.lshapeBind2row", _
        LShapeAdditionalValidationHasOptions(16, 4, Array("EachBar", "EverySecondBar"))
    Dim reader As CSystemSettingsReader
    Set reader = New CSystemSettingsReader
    SetSystemSetting "Geometry.Type", "Circle"
    reader.LoadFromWorkbook ThisWorkbook
    AssertTrue stats, "ui.circle.key.rebarDiameter2", reader.HasKey("Rebar.Diameter2")
    AssertTrue stats, "ui.circle.key.rebarDiameter3", reader.HasKey("Rebar.Diameter3")
    AssertTrue stats, "ui.circle.key.rebarLoc2row", reader.HasKey("Rebar.Loc2row")
    AssertTrue stats, "ui.circle.key.rebarLoc3row", reader.HasKey("Rebar.Loc3row")
End Sub

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

Private Function PlotLoadCaseValidationIsDynamic() As Boolean
    PlotLoadCaseValidationIsDynamic = LoadCaseValidationIsDynamicForSetting("Plot.LoadCase")
End Function

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

Private Function PlotChartExists() As Boolean
    On Error GoTo Failed
    Dim chartObject As Object
    Set chartObject = ThisWorkbook.Worksheets.Item("Расчет").ChartObjects("chtNDMSectionPlot")
    PlotChartExists = Not chartObject Is Nothing
Failed:
End Function

' Проверяет текст заголовка существующей схемы, чтобы не запускать повторную отрисовку только ради проверки подписи.
Private Function PlotChartTitleContains(ByVal expectedText As String) As Boolean
    On Error GoTo Failed
    Dim chartObject As Object
    Set chartObject = ThisWorkbook.Worksheets.Item("Расчет").ChartObjects("chtNDMSectionPlot")
    PlotChartTitleContains = InStr(1, chartObject.Chart.ChartTitle.Text, expectedText, vbTextCompare) > 0
Failed:
End Function

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

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
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

    If TrySetLShapeFaceSetting(key, value) Then Exit Sub

    Err.Raise vbObjectError + 4210, "modTestWorkbookInterface", "System setting not found: " & key
End Sub

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

Private Function TrySetLShapeFaceSetting(ByVal key As String, ByVal value As String) As Boolean
    Dim target As Object
    Set target = ThisWorkbook.Names.Item("rngLShapeGeometry").RefersToRange

    Dim rowIndex As Long
    Dim columnIndex As Long
    If Not LShapeSettingAddress(key, rowIndex, columnIndex) Then Exit Function
    target.Cells.Item(rowIndex, columnIndex).Value2 = value
    TrySetLShapeFaceSetting = True
End Function

Private Function LShapeSettingAddress(ByVal key As String, ByRef rowIndex As Long, ByRef columnIndex As Long) As Boolean
    Dim faceName As String
    If InStr(1, key, "LShape.H1", vbTextCompare) = 1 Then
        faceName = "H1"
    ElseIf InStr(1, key, "LShape.B1", vbTextCompare) = 1 Then
        faceName = "B1"
    ElseIf InStr(1, key, "LShape.H2", vbTextCompare) = 1 Then
        faceName = "H2"
    ElseIf InStr(1, key, "LShape.B2", vbTextCompare) = 1 Then
        faceName = "B2"
    Else
        Exit Function
    End If

    If StrComp(key, "LShape." & faceName, vbTextCompare) = 0 Then
        rowIndex = 3
        columnIndex = LShapeGeometryColumn(faceName)
    ElseIf StrComp(key, "LShape." & faceName & ".as_1", vbTextCompare) = 0 Then
        rowIndex = LShapeMainRow(faceName, 1): columnIndex = 2
    ElseIf StrComp(key, "LShape." & faceName & ".as_2", vbTextCompare) = 0 Then
        rowIndex = LShapeMainRow(faceName, 2): columnIndex = 2
    ElseIf StrComp(key, "LShape." & faceName & ".d_1", vbTextCompare) = 0 Then
        rowIndex = LShapeMainRow(faceName, 1): columnIndex = 3
    ElseIf StrComp(key, "LShape." & faceName & ".d_2", vbTextCompare) = 0 Then
        rowIndex = LShapeMainRow(faceName, 2): columnIndex = 3
    ElseIf StrComp(key, "LShape." & faceName & ".n_1", vbTextCompare) = 0 Then
        rowIndex = LShapeMainRow(faceName, 1): columnIndex = 4
    ElseIf StrComp(key, "LShape." & faceName & ".n_2", vbTextCompare) = 0 Then
        rowIndex = LShapeMainRow(faceName, 2): columnIndex = 4
    ElseIf StrComp(key, "LShape." & faceName & ".StartOffset1", vbTextCompare) = 0 Then
        rowIndex = LShapeMainRow(faceName, 1): columnIndex = 5
    ElseIf StrComp(key, "LShape." & faceName & ".EndOffset1", vbTextCompare) = 0 Then
        rowIndex = LShapeMainRow(faceName, 1): columnIndex = 6
    ElseIf StrComp(key, "LShape." & faceName & ".StartOffset2", vbTextCompare) = 0 Then
        rowIndex = LShapeMainRow(faceName, 2): columnIndex = 5
    ElseIf StrComp(key, "LShape." & faceName & ".EndOffset2", vbTextCompare) = 0 Then
        rowIndex = LShapeMainRow(faceName, 2): columnIndex = 6
    ElseIf StrComp(key, "LShape." & faceName & ".d_2row_1", vbTextCompare) = 0 Then
        rowIndex = LShapeExtraRow(faceName, 1): columnIndex = 2
    ElseIf StrComp(key, "LShape." & faceName & ".d_2row_2", vbTextCompare) = 0 Then
        rowIndex = LShapeExtraRow(faceName, 2): columnIndex = 2
    ElseIf StrComp(key, "LShape." & faceName & ".d_3row_1", vbTextCompare) = 0 Then
        rowIndex = LShapeExtraRow(faceName, 1): columnIndex = 5
    ElseIf StrComp(key, "LShape." & faceName & ".d_3row_2", vbTextCompare) = 0 Then
        rowIndex = LShapeExtraRow(faceName, 2): columnIndex = 5
    ElseIf StrComp(key, "LShape." & faceName & ".loc_2row", vbTextCompare) = 0 Then
        rowIndex = LShapeExtraRow(faceName, 1): columnIndex = 3
    ElseIf StrComp(key, "LShape." & faceName & ".loc_3row", vbTextCompare) = 0 Then
        rowIndex = LShapeExtraRow(faceName, 1): columnIndex = 6
    ElseIf StrComp(key, "LShape." & faceName & ".bind_2row", vbTextCompare) = 0 Then
        rowIndex = LShapeExtraRow(faceName, 1): columnIndex = 4
    ElseIf StrComp(key, "LShape." & faceName & ".bind_3row", vbTextCompare) = 0 Then
        rowIndex = LShapeExtraRow(faceName, 1): columnIndex = 7
    Else
        Exit Function
    End If
    LShapeSettingAddress = True
End Function

Private Function LShapeGeometryColumn(ByVal faceName As String) As Long
    Select Case UCase$(faceName)
        Case "H1": LShapeGeometryColumn = 1
        Case "B1": LShapeGeometryColumn = 2
        Case "H2": LShapeGeometryColumn = 3
        Case "B2": LShapeGeometryColumn = 4
    End Select
End Function

Private Function LShapeMainRow(ByVal faceName As String, ByVal sideIndex As Long) As Long
    LShapeMainRow = 5 + LShapeFaceOrdinal(faceName, sideIndex)
End Function

Private Function LShapeExtraRow(ByVal faceName As String, ByVal sideIndex As Long) As Long
    LShapeExtraRow = 15 + LShapeFaceOrdinal(faceName, sideIndex)
End Function

Private Function LShapeFaceOrdinal(ByVal faceName As String, ByVal sideIndex As Long) As Long
    Select Case UCase$(faceName)
        Case "H1": LShapeFaceOrdinal = sideIndex
        Case "B1": LShapeFaceOrdinal = 2 + sideIndex
        Case "H2": LShapeFaceOrdinal = 4 + sideIndex
        Case "B2": LShapeFaceOrdinal = 6 + sideIndex
    End Select
End Function

Private Function GetSystemSetting(ByVal key As String) As String
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

Private Function SettingsRangeSearchOrder() As Variant
    Dim geometryType As String
    geometryType = SystemGeometryType()
    If StrComp(geometryType, "LShape", vbTextCompare) = 0 Then
        SettingsRangeSearchOrder = Array("rngUnitSettings", "rngSignConventionSettings", "rngSystemSettings", "rngConcreteMaterialParameters", "rngSteelMaterialParameters", "rngCalculationDiagramSettings", "rngPlotAnnotationSettings", "rngLShapeGeometry", "rngCircleGeometry", "rngRoundedRectangleGeometry")
    ElseIf StrComp(geometryType, "RoundedRectangle", vbTextCompare) = 0 Then
        SettingsRangeSearchOrder = Array("rngUnitSettings", "rngSignConventionSettings", "rngSystemSettings", "rngConcreteMaterialParameters", "rngSteelMaterialParameters", "rngCalculationDiagramSettings", "rngPlotAnnotationSettings", "rngRoundedRectangleGeometry", "rngCircleGeometry", "rngLShapeGeometry")
    Else
        SettingsRangeSearchOrder = Array("rngUnitSettings", "rngSignConventionSettings", "rngSystemSettings", "rngConcreteMaterialParameters", "rngSteelMaterialParameters", "rngCalculationDiagramSettings", "rngPlotAnnotationSettings", "rngCircleGeometry", "rngRoundedRectangleGeometry", "rngLShapeGeometry")
    End If
End Function

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

Private Function SystemSettingValidationHasOptions(ByVal key As String, ByVal expectedOptions As Variant) As Boolean
    SystemSettingValidationHasOptions = SettingValidationHasOptionsInRange(ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange, key, expectedOptions)
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

Private Function ResultTableRowCount(ByVal rangeName As String) As Long
    Dim anchor As Object
    Set anchor = ThisWorkbook.Names.Item(rangeName).RefersToRange

    Dim rowOffset As Long
    For rowOffset = 0 To 1048575 - anchor.Row
        If Len(Trim$(CStr(anchor.Offset(rowOffset, 0).Value2))) = 0 Then Exit For
        ResultTableRowCount = ResultTableRowCount + 1
    Next rowOffset
End Function

Private Function ResultTableColumnCount(ByVal anchor As Object) As Long
    Dim colOffset As Long
    For colOffset = 0 To 255
        If Len(Trim$(CStr(anchor.Offset(0, colOffset).Value2))) = 0 Then Exit For
        ResultTableColumnCount = ResultTableColumnCount + 1
    Next colOffset
End Function

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

Private Function LShapeAdditionalValidationHasOptions(ByVal rowIndex As Long, ByVal valueColumn As Long, ByVal expectedOptions As Variant) As Boolean
    On Error GoTo Failed

    Dim settings As Object
    Set settings = ThisWorkbook.Names.Item("rngLShapeGeometry").RefersToRange
    LShapeAdditionalValidationHasOptions = ValidationCellHasOptions(settings.Cells.Item(rowIndex, valueColumn), expectedOptions)
    Exit Function
Failed:
End Function

Private Function LoadCombinationValidationHasOptions(ByVal valueColumn As Long, ByVal expectedOptions As Variant) As Boolean
    On Error GoTo Failed

    Dim loads As Object
    Set loads = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange
    LoadCombinationValidationHasOptions = ValidationCellHasOptions(loads.Cells.Item(2, valueColumn), expectedOptions)
    Exit Function
Failed:
End Function

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

' Создает расчетный или интерфейсный объект из нормализованных исходных данных и локальных настроек.
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
        settings.GetLong("Rebar.Count", 8), _
        settings.GetDouble("Rebar.Diameter", 20#), _
        settings.GetString("Steel.RebarProfile", "Ribbed"))
    Dim materialProvider As CMaterialModelProvider
    Set materialProvider = New CMaterialModelProvider
    materialProvider.Initialize settings

    Dim section As CSectionModel
    Set section = BuildGeneratedSectionModel(mesh, rebars, "WorkbookInterface")

    Dim batch As CBatchSectionCalculator
    Set batch = New CBatchSectionCalculator
    batch.Initialize section, materialProvider
    batch.ApplySettings settings
    Set BuildUiBatch = batch
End Function

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
    SetSystemSetting "Calculation.Mode", "DirectState"
    SetSystemSetting "Plot.LoadCase", "LC1"
    SetSystemSetting "Steel.RebarProfile", "Ribbed"
    SetSystemSetting "Rebar.AxisDistance", "40"
    SetSystemSetting "Rebar.Count", "8"
    SetSystemSetting "Rebar.Diameter", "20"
    SetSystemSetting "SLS.Crack.Allowable", "0.3"

    Dim loads As Object
    Set loads = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange
    ClearDataRows loads
    loads.Cells.Item(2, 1).Value2 = "LC1"
    loads.Cells.Item(2, 2).Value2 = -100000#
    loads.Cells.Item(2, 3).Value2 = -4000000#
    loads.Cells.Item(2, 4).Value2 = -3000000#
    loads.Cells.Item(2, 5).Value2 = "Group2"
    loads.Cells.Item(2, 6).Value2 = ChrW$(&H3BB) & "*Mxy"
    loads.Cells.Item(2, 7).Value2 = "ui test"
End Sub

Private Sub PrepareLShapeInput()
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
    SetSystemSetting "Geometry.Type", "LShape"
    SetSystemSetting "Calculation.Mode", "DirectState"
    SetSystemSetting "Plot.LoadCase", "LC_L"
    SetSystemSetting "Mesh.Step", "40"
    SetSystemSetting "Mesh.BoundarySubdivisions", "2"
    SetSystemSetting "LShape.B1", "160"
    SetSystemSetting "LShape.H1", "280"
    SetSystemSetting "LShape.B2", "360"
    SetSystemSetting "LShape.H2", "140"
    SetSystemSetting "LShape.H1.as_1", "40"
    SetSystemSetting "LShape.H1.as_2", "40"
    SetSystemSetting "LShape.H1.d_1", "16"
    SetSystemSetting "LShape.H1.d_2", "16"
    SetSystemSetting "LShape.H1.n_1", "3"
    SetSystemSetting "LShape.H1.n_2", "2"
    SetSystemSetting "LShape.H1.StartOffset1", "40"
    SetSystemSetting "LShape.H1.EndOffset1", "40"
    SetSystemSetting "LShape.H1.StartOffset2", "40"
    SetSystemSetting "LShape.H1.EndOffset2", "40"
    SetSystemSetting "LShape.H2.as_1", "40"
    SetSystemSetting "LShape.H2.as_2", "40"
    SetSystemSetting "LShape.H2.d_1", "16"
    SetSystemSetting "LShape.H2.d_2", "16"
    SetSystemSetting "LShape.H2.n_1", "0"
    SetSystemSetting "LShape.H2.n_2", "0"
    SetSystemSetting "LShape.H2.StartOffset1", "40"
    SetSystemSetting "LShape.H2.EndOffset1", "40"
    SetSystemSetting "LShape.H2.StartOffset2", "40"
    SetSystemSetting "LShape.H2.EndOffset2", "40"
    SetSystemSetting "LShape.B1.as_1", "40"
    SetSystemSetting "LShape.B1.as_2", "40"
    SetSystemSetting "LShape.B1.d_1", "16"
    SetSystemSetting "LShape.B1.d_2", "16"
    SetSystemSetting "LShape.B1.n_1", "2"
    SetSystemSetting "LShape.B1.n_2", "1"
    SetSystemSetting "LShape.B1.StartOffset1", "20"
    SetSystemSetting "LShape.B1.EndOffset1", "20"
    SetSystemSetting "LShape.B1.StartOffset2", "20"
    SetSystemSetting "LShape.B1.EndOffset2", "20"
    SetSystemSetting "LShape.B2.as_1", "40"
    SetSystemSetting "LShape.B2.as_2", "40"
    SetSystemSetting "LShape.B2.d_1", "16"
    SetSystemSetting "LShape.B2.d_2", "16"
    SetSystemSetting "LShape.B2.n_1", "2"
    SetSystemSetting "LShape.B2.n_2", "0"
    SetSystemSetting "LShape.B2.StartOffset1", "60"
    SetSystemSetting "LShape.B2.EndOffset1", "60"
    SetSystemSetting "LShape.B2.StartOffset2", "60"
    SetSystemSetting "LShape.B2.EndOffset2", "60"
    SetSystemSetting "SLS.Crack.Allowable", "0.3"

    Dim loads As Object
    Set loads = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange
    ClearDataRows loads
    loads.Cells.Item(2, 1).Value2 = "LC_L"
    loads.Cells.Item(2, 2).Value2 = -80000#
    loads.Cells.Item(2, 3).Value2 = -1500000#
    loads.Cells.Item(2, 4).Value2 = -1000000#
    loads.Cells.Item(2, 5).Value2 = "Group2"
    loads.Cells.Item(2, 6).Value2 = ChrW$(&H3BB) & "*Mxy"
    loads.Cells.Item(2, 7).Value2 = "lshape ui test"
End Sub

' Настраивает книгу под типовой прочностной расчет Г-сечения:
' пользователь задает N и Mx, а несущая способность ищется увеличением Mx.
' Здесь специально выбран Capacity.SolutionStrategy = UltimateStrain, чтобы проверить,
' что быстрый изгибный путь работает без fallback на LoadMultiplier.
Private Sub PrepareUserLShapeMomentUltimateInput()
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
    SetSystemSetting "Geometry.Type", "LShape"
    SetSystemSetting "Calculation.Mode", "FullCapacity"
    SetSystemSetting "Capacity.CalculationScope", "Group1Only"
    SetSystemSetting "Capacity.SolutionStrategy", "UltimateStrain"
    SetSystemSetting "Capacity.MaxLambda", "64"
    SetSystemSetting "Capacity.ToleranceStrain", "0.00001"
    SetSystemSetting "Capacity.SolverMaxIterations", "60"
    SetSystemSetting "Solver.Method", "Newton"
    SetSystemSetting "Solver.DirectState.DiagramExtension", "Yes"
    SetSystemSetting "Solver.MaxIterations", "80"
    SetSystemSetting "Solver.LoadSteps", "1"
    SetSystemSetting "Mesh.Step", "50"
    SetSystemSetting "Mesh.BoundarySubdivisions", "1"
    SetSystemSetting "Load.ReferenceOffsetX", "0"
    SetSystemSetting "Load.ReferenceOffsetY", "0"
    SetSystemSetting "Plot.LoadCase", "LC_MX"
    SetSystemSetting "General.ExecutionReportEnabled", "No"

    SetSystemSetting "LShape.H1", "550"
    SetSystemSetting "LShape.B1", "250"
    SetSystemSetting "LShape.H2", "250"
    SetSystemSetting "LShape.B2", "600"

    SetUserLShapeMainRow "H1", 5, 5
    SetUserLShapeMainRow "B1", 2, 2
    SetUserLShapeMainRow "H2", 2, 2
    SetUserLShapeMainRow "B2", 5, 5
    ClearUserLShapeExtraRows

    Dim loads As Object
    Set loads = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange
    ClearDataRows loads
    loads.Cells.Item(2, 1).Value2 = "LC_MX"
    loads.Cells.Item(2, 2).Value2 = 200#
    loads.Cells.Item(2, 3).Value2 = 50#
    loads.Cells.Item(2, 4).ClearContents
    loads.Cells.Item(2, 5).Value2 = "Group1"
    loads.Cells.Item(2, 6).Value2 = ChrW$(&H3BB) & "*Mx"
    loads.Cells.Item(2, 7).Value2 = "moment ultimate regression"
End Sub

' Настраивает книгу ровно под пользовательский пример с Г-сечением и осевым
' растяжением в tf. Это дополняет unit-тест batch-слоя проверкой полного
' Excel-пути: пользовательский знак N, единицы, чтение Config и запись Results.
Private Sub PrepareUserLShapeAxialTensionInput()
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
    SetSystemSetting "Geometry.Type", "LShape"
    SetSystemSetting "Calculation.Mode", "DirectState"
    SetSystemSetting "Solver.Method", "Newton"
    SetSystemSetting "Solver.DirectState.DiagramExtension", "Yes"
    SetSystemSetting "Solver.MaxIterations", "80"
    SetSystemSetting "Solver.LoadSteps", "1"
    SetSystemSetting "Mesh.Step", "50"
    SetSystemSetting "Mesh.BoundarySubdivisions", "1"
    SetSystemSetting "Load.ReferenceOffsetX", "0"
    SetSystemSetting "Load.ReferenceOffsetY", "0"
    SetSystemSetting "Plot.LoadCase", "LC_OVER"
    SetSystemSetting "General.ExecutionReportEnabled", "Yes"

    SetSystemSetting "LShape.H1", "550"
    SetSystemSetting "LShape.B1", "250"
    SetSystemSetting "LShape.H2", "250"
    SetSystemSetting "LShape.B2", "600"

    SetUserLShapeMainRow "H1", 5, 5
    SetUserLShapeMainRow "B1", 2, 2
    SetUserLShapeMainRow "H2", 2, 2
    SetUserLShapeMainRow "B2", 5, 5
    ClearUserLShapeExtraRows

    Dim loads As Object
    Set loads = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange
    ClearDataRows loads
    loads.Cells.Item(2, 1).Value2 = "LC_SAFE"
    loads.Cells.Item(2, 2).Value2 = -795#
    loads.Cells.Item(2, 3).ClearContents
    loads.Cells.Item(2, 4).ClearContents
    loads.Cells.Item(2, 5).Value2 = "Group2"
    loads.Cells.Item(2, 6).Value2 = ChrW$(&H3BB) & "*N"
    loads.Cells.Item(2, 7).Value2 = "inside physical range"

    loads.Cells.Item(3, 1).Value2 = "LC_OVER"
    loads.Cells.Item(3, 2).Value2 = -900#
    loads.Cells.Item(3, 3).ClearContents
    loads.Cells.Item(3, 4).ClearContents
    loads.Cells.Item(3, 5).Value2 = "Group2"
    loads.Cells.Item(3, 6).Value2 = ChrW$(&H3BB) & "*N"
    loads.Cells.Item(3, 7).Value2 = "uses extension"
End Sub

Private Sub SetUserLShapeMainRow(ByVal faceName As String, ByVal count1 As Long, ByVal count2 As Long)
    SetSystemSetting "LShape." & faceName & ".as_1", "40"
    SetSystemSetting "LShape." & faceName & ".as_2", "40"
    SetSystemSetting "LShape." & faceName & ".d_1", "32"
    SetSystemSetting "LShape." & faceName & ".d_2", "32"
    SetSystemSetting "LShape." & faceName & ".n_1", CStr(count1)
    SetSystemSetting "LShape." & faceName & ".n_2", CStr(count2)
    SetSystemSetting "LShape." & faceName & ".StartOffset1", "80"
    SetSystemSetting "LShape." & faceName & ".EndOffset1", "80"
    SetSystemSetting "LShape." & faceName & ".StartOffset2", "80"
    SetSystemSetting "LShape." & faceName & ".EndOffset2", "80"
End Sub

Private Sub ClearUserLShapeExtraRows()
    Dim faces As Variant
    faces = Array("H1", "B1", "H2", "B2")
    Dim i As Long
    For i = LBound(faces) To UBound(faces)
        SetSystemSetting "LShape." & CStr(faces(i)) & ".d_2row_1", vbNullString
        SetSystemSetting "LShape." & CStr(faces(i)) & ".d_2row_2", vbNullString
        SetSystemSetting "LShape." & CStr(faces(i)) & ".d_3row_1", vbNullString
        SetSystemSetting "LShape." & CStr(faces(i)) & ".d_3row_2", vbNullString
        SetSystemSetting "LShape." & CStr(faces(i)) & ".loc_2row", "Stacked"
        SetSystemSetting "LShape." & CStr(faces(i)) & ".loc_3row", "Stacked"
        SetSystemSetting "LShape." & CStr(faces(i)) & ".bind_2row", "EachBar"
        SetSystemSetting "LShape." & CStr(faces(i)) & ".bind_3row", "EachBar"
    Next i
End Sub

' Очищает накопленное состояние перед новым расчетом или повторным формированием вывода.
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
    AssertTextEquals stats, "ui.units.capacity.mode", SystemSettingUnitText("Calculation.Mode"), "-"
    AssertTextEquals stats, "ui.units.capacity.scope", SystemSettingUnitText("Capacity.CalculationScope"), "-"
    AssertTextEquals stats, "ui.units.Capacity.SolutionStrategy", SystemSettingUnitText("Capacity.SolutionStrategy"), "-"
    AssertTextEquals stats, "ui.units.capacity.searchMethod", SystemSettingUnitText("Capacity.SearchMethod"), "-"
    AssertTextEquals stats, "ui.units.capacity.initialLambda", SystemSettingUnitText("Capacity.InitialLambda"), "-"
    AssertTextEquals stats, "ui.units.capacity.toleranceLambda", SystemSettingUnitText("Capacity.ToleranceLambda"), "-"
    AssertTextEquals stats, "ui.units.capacity.maxRetries", SystemSettingUnitText("Capacity.MaxRetries"), "шт"
    AssertTextEquals stats, "ui.units.capacity.baseLoadSteps", SystemSettingUnitText("Capacity.BaseLoadSteps"), "шт"
    AssertTextEquals stats, "ui.units.capacity.toleranceStrain", SystemSettingUnitText("Capacity.ToleranceStrain"), "-"
    AssertTextEquals stats, "ui.units.capacity.maxLambda", SystemSettingUnitText("Capacity.MaxLambda"), "-"
    AssertTextEquals stats, "ui.units.capacity.solverIterations", SystemSettingUnitText("Capacity.SolverMaxIterations"), "шт"
End Sub

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

Private Function FileExists(ByVal path As String) As Boolean
    FileExists = CreateObject("Scripting.FileSystemObject").FileExists(path)
End Function

Private Sub DeleteFileIfExists(ByVal path As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(path) Then fso.DeleteFile path, True
End Sub

Private Function ReadTextFile(ByVal path As String) As String
    Dim stream As Object
    Set stream = CreateObject("Scripting.FileSystemObject").OpenTextFile(path, 1, False, -1)
    ReadTextFile = stream.ReadAll
    stream.Close
End Function

Private Sub AssertTrue(ByRef stats As TUiTestStats, ByVal name As String, ByVal condition As Boolean)
    If condition Then
        stats.Passed = stats.Passed + 1
        AppendLine stats, "OK: " & name
    Else
        stats.Failed = stats.Failed + 1
        AppendLine stats, "FAIL: " & name
    End If
End Sub

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

Private Sub AppendLine(ByRef stats As TUiTestStats, ByVal text As String)
    stats.Report = stats.Report & text & vbCrLf
End Sub

Private Function FormatNumberInvariant(ByVal value As Double) As String
    FormatNumberInvariant = Replace$(Format$(value, "0.############"), ",", ".")
End Function












