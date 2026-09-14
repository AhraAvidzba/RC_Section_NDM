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

' Запускает связанный набор операций и возвращает пользователю итоговый статус выполнения.
Public Function RunWorkbookInterfaceTests() As String
    On Error GoTo Failed

    Dim stats As TUiTestStats
    Dim t0 As Double
    t0 = Timer

    TestButtons stats
    TestLoadCombinationsOnConfig stats
    TestSingleCombinationSkipsBlankRows stats
    TestLoadCombinationRangeMinimumRows stats
    TestPartialCombinationIsInvalid stats
    TestInvalidProfileIdDoesNotRunPlot stats
    TestAutoCADSourceRequiresManualImport stats
    TestAutoCADImportButtonRejectsGeneratedSource stats
    TestAutoCADImporterTreatsDrawingUnitsAsMillimeters stats
    TestAutoCADPreviewWritesAndDrawsBoundsDimensions stats
    TestPlotClearsLegacyWorksheetShapes stats
    TestPlotOverlayCoordinatesMatchResults stats
    TestGeneratedSourceDoesNotReuseAutoCADPreview stats
    TestGeneratedDirectStateWorstStillDrawsFirstCalculatedLC stats
    TestProfileDrivenPlotUsesSnapshotState stats
    TestMissingProfileStateDrawsGeometryOnly stats
    TestAutoCADCalculationMessageUsesSavedGeometry stats
    TestBlankMomentDefaultsToZeroAndZeroLoadsAreSkipped stats
    TestCircleWorkbookRunWritesResults stats
    TestTwentyCombinationsWithFiveStatesWriteSnapshot stats
    TestDynamicLoadCombinationRangeAndLayoutGuard stats
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

' ДЛЯ ТЕСТОВ
' Запускает тяжелый ручной сценарий: 20 LC, пять named-state на LC и мелкая
' сетка Г-сечения. Он не входит в обычный RunWorkbookInterfaceTests, потому что
' нужен только для проверки больших snapshot-ов и лимита Excel Chart на series.
Public Function RunLargeSnapshotPlotStressTest() As String
    On Error GoTo Failed

    Dim stats As TUiTestStats
    Dim t0 As Double
    t0 = Timer

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
    Dim profiles As CCalculationProfileCatalog
    Set profiles = New CCalculationProfileCatalog
    profiles.LoadFromWorkbook ThisWorkbook
    Set batch.ProfileCatalog = profiles

    Dim reader As CLoadCombinationReader
    Set reader = New CLoadCombinationReader
    reader.LoadFromWorkbook ThisWorkbook, batch, units
    batch.ApplyLoadReference props.CentroidX, props.CentroidY, props.CentroidX, props.CentroidY
    batch.Execute

    AssertTrue stats, "ui.autocad.reference.converged", batch.StateConverged(1)
    AssertTrue stats, "ui.autocad.reference.point", Abs(batch.LoadReferenceX) > 0.000001 Or Abs(batch.LoadReferenceY) > 0.000001
    AssertClose stats, "ui.autocad.reference.concreteCenterX", props.CentroidX, batch.LoadReferenceX, 0.000001
    AssertClose stats, "ui.autocad.reference.concreteCenterY", props.CentroidY, batch.LoadReferenceY, 0.000001
    Dim expectedLoad As CSectionLoadState
    Set expectedLoad = New CSectionLoadState
    expectedLoad.Initialize batch.N(1), batch.UserMx(1), batch.UserMy(1), props.CentroidX, props.CentroidY
    AssertClose stats, "ui.autocad.reference.mxTransfer", batch.Mx(1), expectedLoad.InternalMx, 0.000001
    AssertClose stats, "ui.autocad.reference.myTransfer", batch.My(1), expectedLoad.InternalMy, 0.000001
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
    section.ApplyAverageRotationToEquivalentAreaFallbacks

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
    AssertTrue stats, "ui.autocad.preview.readerAnnotations", reader.AnnotationCount = 2

    UpdateSectionPlotForWorkbook ThisWorkbook
    AssertTrue stats, "ui.autocad.preview.dimensionShapes", CountPlotShapes("AnnotationLine") > 0
    AssertTrue stats, "ui.autocad.preview.rotatedElementShape", _
        PlotShapeRotationExists("ElementConcrete", -30#, 0.5)
    Dim expectedFallbackDegrees As Double
    expectedFallbackDegrees = -0.5 * Atn(Sin(GEOM_PI / 3#) / (2# + Cos(GEOM_PI / 3#))) * 180# / GEOM_PI
    AssertTrue stats, "ui.autocad.preview.squareFallbackRotation", _
        PlotShapeRotationExists("ElementConcrete", expectedFallbackDegrees, 0.5)
End Sub

' Проверяет, что новая Chart-схема удаляет старые листовые NDMPlot_*
' объекты. Такие Shapes могли остаться от прежней реализации и визуально
' сдвигать точку нагрузки или оси относительно актуальной схемы внутри Chart.
Private Sub TestPlotClearsLegacyWorksheetShapes(ByRef stats As TUiTestStats)
    PrepareCircleInput
    SetSystemSetting "Geometry.Source", "AutoCAD"

    Dim section As CSectionModel
    Set section = New CSectionModel
    section.SourceType = "AutoCADImport"
    section.AddConcreteElement 50#, 50#, 10000#, 1, vbNullString, vbNullString, "Rectangle", 100#, 100#

    Dim writer As CNDMResultsWriter
    Set writer = New CNDMResultsWriter
    writer.WriteGeometryPreview ThisWorkbook, section

    Dim calc As Object
    Set calc = ThisWorkbook.Worksheets.Item("Расчет")
    Dim legacyShape As Object
    Set legacyShape = calc.Shapes.AddShape(1, 10#, 10#, 20#, 20#)
    legacyShape.Name = "NDMPlot_LegacyWorksheetShapeForTest"

    AssertTrue stats, "ui.plot.legacyWorksheetShape.created", _
        CountWorksheetPlotShapes("LegacyWorksheetShapeForTest") = 1

    UpdateSectionPlotForWorkbook ThisWorkbook

    AssertTrue stats, "ui.plot.legacyWorksheetShape.removed", _
        CountWorksheetPlotShapes("LegacyWorksheetShapeForTest") = 0
    AssertTrue stats, "ui.plot.legacyWorksheetShape.chartStillDraws", CountPlotShapes("AnnotationLine") > 0
End Sub

' Проверяет не только наличие осей/точки, но и их взаимное положение.
' Для режима Transformed центр главных осей берется из Transformed.Centroid,
' а точка нагрузки - из LoadReference. Их экранный сдвиг должен совпадать с
' расчетным сдвигом из Results после одного общего model-to-chart масштаба.
Private Sub TestPlotOverlayCoordinatesMatchResults(ByRef stats As TUiTestStats)
    PrepareUserLShapeMomentUltimateInput
    SetSystemSetting "Plot.PrincipalAxesMode", "Transformed"
    SetSystemSetting "Plot.LoadApplicationPointEnabled", "Yes"
    SetSystemSetting "Plot.ResultLabelsEnabled", "No"
    SetSystemSetting "Plot.AutoUpdateAfterCalculation", "Yes"

    Dim message As String
    message = RunSectionCalculationForWorkbook(ThisWorkbook, False)
    AssertTrue stats, "ui.plot.overlay.run", InStr(1, message, "Расчет завершен", vbTextCompare) > 0

    Dim loadX As Double
    Dim loadY As Double
    Dim centroidX As Double
    Dim centroidY As Double
    loadX = CDbl(ResultsPropertyValue("ALL", "LoadReferenceX"))
    loadY = CDbl(ResultsPropertyValue("ALL", "LoadReferenceY"))
    centroidX = CDbl(ResultsPropertyValue("ALL", "Transformed.CentroidX"))
    centroidY = CDbl(ResultsPropertyValue("ALL", "Transformed.CentroidY"))

    Dim loadPoint As Object
    Set loadPoint = FirstGeneratedPlotShape("LoadPoint")
    AssertTrue stats, "ui.plot.overlay.shapes", CountGeneratedPlotShapes("Principal1") > 0 And _
        CountGeneratedPlotShapes("Principal2") > 0 And Not loadPoint Is Nothing
    If CountGeneratedPlotShapes("Principal1") <= 0 Or CountGeneratedPlotShapes("Principal2") <= 0 Or loadPoint Is Nothing Then Exit Sub
    AssertTrue stats, "ui.plot.overlay.chartLayer", CountPlotShapes("Principal1") > 0 And _
        CountPlotShapes("Principal2") > 0 And CountPlotShapes("LoadPoint") > 0 And _
        CountWorksheetPlotShapes("Principal1") = 0 And CountWorksheetPlotShapes("Principal2") = 0 And _
        CountWorksheetPlotShapes("LoadPoint") = 0

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
    writer.WriteGeometryPreview ThisWorkbook, section

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
    writer.WriteGeometryPreview ThisWorkbook, section, units

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
    loads.Cells.Item(3, 5).Value2 = "PR1"

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
    Dim crackAnchor As Object
    Set crackAnchor = ThisWorkbook.Names.Item("rngCrackSummaryAnchor").RefersToRange
    AssertTrue stats, "ui.run.crack", Len(CStr(crackAnchor.Worksheet.Cells.Item(crackAnchor.Row, 2).Value2)) > 0 And _
        Len(CStr(crackAnchor.Worksheet.Cells.Item(crackAnchor.Row, 18).Value2)) > 0
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
    AssertTrue stats, "ui.results.strength.anchor", ThisWorkbook.Names.Item("rngStrengthSummaryAnchor").RefersToRange.Row = 36 And ThisWorkbook.Names.Item("rngStrengthSummaryAnchor").RefersToRange.Column = 1
    AssertTrue stats, "ui.results.crack.anchor", ThisWorkbook.Names.Item("rngCrackSummaryAnchor").RefersToRange.Row = 62 And ThisWorkbook.Names.Item("rngCrackSummaryAnchor").RefersToRange.Column = 1
    AssertTrue stats, "ui.results.stability.anchor", ThisWorkbook.Names.Item("rngStabilitySummaryAnchor").RefersToRange.Row = 89 And ThisWorkbook.Names.Item("rngStabilitySummaryAnchor").RefersToRange.Column = 1
    AssertTrue stats, "ui.results.geometry.position", ThisWorkbook.Names.Item("rngNDMSectionGeometry").RefersToRange.Row = 113 And ThisWorkbook.Names.Item("rngNDMSectionGeometry").RefersToRange.Column = 12
    AssertTrue stats, "ui.results.properties.position", ThisWorkbook.Names.Item("rngNDMSectionProperties").RefersToRange.Row = 113 And ThisWorkbook.Names.Item("rngNDMSectionProperties").RefersToRange.Column = 29
    AssertTrue stats, "ui.results.materialDiagrams.position", ThisWorkbook.Names.Item("rngNDMMaterialDiagrams").RefersToRange.Row = 113 And ThisWorkbook.Names.Item("rngNDMMaterialDiagrams").RefersToRange.Column = 37
    AssertTrue stats, "ui.results.annotations.position", ThisWorkbook.Names.Item("rngNDMSectionAnnotations").RefersToRange.Row = 113 And ThisWorkbook.Names.Item("rngNDMSectionAnnotations").RefersToRange.Column = 50
    AssertTrue stats, "ui.results.geometry.noSource", ResultHeaderColumn(geometryResults, "SourceName") = 0
    AssertTrue stats, "ui.results.geometry.noMaterialClass", ResultHeaderColumn(geometryResults, "MaterialClass") = 0
    AssertTrue stats, "ui.results.properties.header", CStr(ThisWorkbook.Names.Item("rngNDMSectionProperties").RefersToRange.Value2) = "RunID"
    AssertTrue stats, "ui.results.properties.hasStateEpsilon0", ResultsPropertyExists("LC1", "State.CrackedState.Epsilon0")
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
        ProfileSettingValue("PR4", "Visualization.Quantity") = "Strain"
End Sub

' Проверяет полный предельный snapshot: 20 сочетаний, каждое с пятью
' конечными named-state. Для устойчивого получения Before/AfterMcrcState
' используется чистый изгиб: сжатие может подавить образование нормальной
' трещины и тогда эти состояния физически не обязаны появляться.
Private Sub TestTwentyCombinationsWithFiveStatesWriteSnapshot(ByRef stats As TUiTestStats)
    PrepareUserLShapeMomentUltimateInput
    PrepareFullStateProfile "PR3"
    SetSystemSetting "SLS.Crack.PsiMode", "Auto"
    SetSystemSetting "SLS.Crack.Allowable", "0.000001"
    SetSystemSetting "Plot.AutoUpdateAfterCalculation", "No"
    SetSystemSetting "Plot.LoadCase", "LC_FULL_01"

    Dim loads As Object
    Set loads = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange
    ClearDataRows loads

    Dim rowIndex As Long
    For rowIndex = 1 To 20
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
    expectedElementRows = 1 + 20 * 5 * geometryRowCount
    AssertTrue stats, "ui.results.fullSnapshot.elementRows", _
        UBound(elementResults, 1) = expectedElementRows

    AssertElementStateRows stats, elementResults, "LC_FULL_01", "StrengthState", geometryRowCount
    AssertElementStateRows stats, elementResults, "LC_FULL_01", "CapacityState", geometryRowCount
    AssertElementStateRows stats, elementResults, "LC_FULL_01", "CrackedState", geometryRowCount
    AssertElementStateRows stats, elementResults, "LC_FULL_01", "BeforeMcrcState", geometryRowCount
    AssertElementStateRows stats, elementResults, "LC_FULL_01", "AfterMcrcState", geometryRowCount
    AssertElementStateRows stats, elementResults, "LC_FULL_20", "StrengthState", geometryRowCount
    AssertElementStateRows stats, elementResults, "LC_FULL_20", "CapacityState", geometryRowCount
    AssertElementStateRows stats, elementResults, "LC_FULL_20", "CrackedState", geometryRowCount
    AssertElementStateRows stats, elementResults, "LC_FULL_20", "BeforeMcrcState", geometryRowCount
    AssertElementStateRows stats, elementResults, "LC_FULL_20", "AfterMcrcState", geometryRowCount

    AssertTrue stats, "ui.results.fullSnapshot.propertiesRows", _
        ResultTableRowCount("rngNDMSectionProperties") >= 1 + 55 + 20 * (23 + 5 * 10)
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
    Set tempRange = tempSheet.Range("A1:G22")
    FillLoadCombinationTestRange tempRange, 21, "PR1", "LC_DYN_"
    ThisWorkbook.Names.Item("rngLoadCombinations").RefersTo = "=" & tempRange.Address(True, True, 1, True)

    Dim batch As CBatchSectionCalculator
    Set batch = BuildUiBatch()
    Dim reader As CLoadCombinationReader
    Set reader = New CLoadCombinationReader
    reader.LoadFromWorkbook ThisWorkbook, batch
    AssertTrue stats, "ui.loads.dynamicRange.count21", batch.Count = 21

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
        InStr(1, errorText, "не хватает места", vbTextCompare) > 0 And _
        InStr(1, errorText, "Расчет не запущен", vbTextCompare) > 0 And _
        InStr(1, errorText, "пустые строки", vbTextCompare) > 0 And _
        InStr(1, errorText, "rngBatchSummary", vbTextCompare) > 0 And _
        InStr(1, errorText, "rngStrengthSummaryAnchor", vbTextCompare) > 0 And _
        InStr(1, errorText, "rngCrackSummaryAnchor", vbTextCompare) > 0 And _
        InStr(1, errorText, "rngStabilitySummaryAnchor", vbTextCompare) > 0

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
    target.Cells.Item(1, 6).Value2 = "CapacityLoadPath"
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
    PrepareUserLShapeMomentUltimateInput
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

' Проверяет чистый изгиб Г-сечения по профилю PR1. Такой профиль запрашивает
' и прямое НДС, и несущую способность, поэтому обе ветви должны проходить
' через полный workbook-path без специальных обходов.
Private Sub TestLShapePureBendingDirectStateWorkbookPath(ByRef stats As TUiTestStats)
    PrepareUserLShapeMomentUltimateInput

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
    AssertTextEquals stats, "ui.lshape.pureBendingDirect.capacityOk", capacityStatus, "OK"
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
    AssertTrue stats, "ui.lshape.axial900.extensionStressSnapshot", _
        MaxAbsElementStress(ResultTable("rngNDMElementResults"), "LC_OVER", "CrackedState") > 390.1
    AssertTrue stats, "ui.lshape.axial900.reportCreated", FileExists(reportPath)

    SetSystemSetting "General.ExecutionReportEnabled", "No"
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestGoverningCombinationWritesDetailedResults(ByRef stats As TUiTestStats)
    PrepareCircleInput
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

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestCapacitySearchMethodValidation(ByRef stats As TUiTestStats)
    AssertTrue stats, "ui.validation.CapacitySolutionStrategy", _
        SystemSettingValidationHasOptions("Capacity.SolutionStrategy", Array("Auto", "UltimateStrain", "LoadMultiplier"))
    AssertTrue stats, "ui.validation.nonCriticalMessages", _
        SystemSettingValidationHasOptions("General.NonCriticalMessagesEnabled", Array("Yes", "No"))
    AssertTrue stats, "ui.validation.capacitySearchMethod", _
        SystemSettingValidationHasOptions("Capacity.SearchMethod", Array("Bisection", "Brent", "Secant"))
    AssertTrue stats, "ui.validation.crackCoverDistanceMode", _
        SystemSettingValidationHasOptions("SLS.Crack.CoverDistanceMode", Array("NearestContour", "GlobalExtreme"))
    AssertTrue stats, "ui.validation.crackInitiationLoadPath", _
        SystemSettingValidationHasOptions("SLS.Crack.InitiationLoadPath", Array("Auto", ChrW$(&H3BB) & "*Mxy", ChrW$(&H3BB) & "*N", ChrW$(&H3BB) & "*NMxy"))
    AssertTrue stats, "ui.validation.autocadLabelMode", _
        SystemSettingValidationHasOptions("AutoCAD.Export.LabelMode", Array("ValuesOnly", "NamesAndValues"))
    AssertTrue stats, "ui.validation.autocadNeutralLine", _
        SystemSettingValidationHasOptions("AutoCAD.Export.NeutralLineEnabled", Array("Yes", "No"))
    AssertTrue stats, "ui.validation.autocadPrincipalAxes", _
        SystemSettingValidationHasOptions("AutoCAD.Export.PrincipalAxesMode", Array("Transformed", "Concrete", "None"))
    AssertTrue stats, "ui.validation.autocadLoadPoint", _
        SystemSettingValidationHasOptions("AutoCAD.Export.LoadPointEnabled", Array("Yes", "No"))
    AssertTrue stats, "ui.validation.autocadCombination", AutoCADCombinationValidationIsDynamic()
    AssertTrue stats, "ui.validation.plotLoadCase", PlotLoadCaseValidationIsDynamic()
    AssertTrue stats, "ui.validation.loadProfileId", _
        LoadCombinationValidationHasOptions(5, Array("PR1", "PR2", "PR3", "PR4"))
    AssertTrue stats, "ui.validation.plotLabels", _
        SystemSettingValidationHasOptions("Plot.ResultLabelsEnabled", Array("Yes", "No"))
    AssertTrue stats, "ui.validation.plotPrincipalAxes", _
        SystemSettingValidationHasOptions("Plot.PrincipalAxesMode", Array("Transformed", "Concrete", "None"))
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

    concreteArea = CDbl(ResultsPropertyValue("ALL", "Concrete.Area"))
    transformedArea = CDbl(ResultsPropertyValue("ALL", "Transformed.Area"))
    rebarArea = 8# * GEOM_PI * 20# * 20# / 4#
    expectedArea = concreteArea + (200000# / 32500# - 1#) * rebarArea

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
End Sub

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

Private Function ShapeCenterX(ByVal shapeObject As Object) As Double
    ShapeCenterX = CDbl(shapeObject.Left) + CDbl(shapeObject.Width) / 2#
End Function

Private Function ShapeCenterY(ByVal shapeObject As Object) As Double
    ShapeCenterY = CDbl(shapeObject.Top) + CDbl(shapeObject.Height) / 2#
End Function

Private Function CountGeneratedPlotShapes(ByVal nameFragment As String) As Long
    CountGeneratedPlotShapes = CountPlotShapes(nameFragment) + CountWorksheetPlotShapes(nameFragment)
End Function

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
        SettingsRangeSearchOrder = Array("rngUnitSettings", "rngSignConventionSettings", "rngSystemSettings", "rngConcreteMaterialParameters", "rngSteelMaterialParameters", "rngCalculationProfiles", "rngPlotAnnotationSettings", "rngLShapeGeometry", "rngCircleGeometry", "rngRoundedRectangleGeometry")
    ElseIf StrComp(geometryType, "RoundedRectangle", vbTextCompare) = 0 Then
        SettingsRangeSearchOrder = Array("rngUnitSettings", "rngSignConventionSettings", "rngSystemSettings", "rngConcreteMaterialParameters", "rngSteelMaterialParameters", "rngCalculationProfiles", "rngPlotAnnotationSettings", "rngRoundedRectangleGeometry", "rngCircleGeometry", "rngLShapeGeometry")
    Else
        SettingsRangeSearchOrder = Array("rngUnitSettings", "rngSignConventionSettings", "rngSystemSettings", "rngConcreteMaterialParameters", "rngSteelMaterialParameters", "rngCalculationProfiles", "rngPlotAnnotationSettings", "rngCircleGeometry", "rngRoundedRectangleGeometry", "rngLShapeGeometry")
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
    SetSystemSetting "Plot.LoadCase", "LC_L"
    SetSystemSetting "Mesh.StepX", "40"
    SetSystemSetting "Mesh.StepY", "40"
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
    loads.Cells.Item(2, 5).Value2 = "PR2"
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
    SetSystemSetting "Capacity.SolutionStrategy", "UltimateStrain"
    SetSystemSetting "Capacity.MaxLambda", "64"
    SetSystemSetting "Capacity.ToleranceStrain", "0.00001"
    SetSystemSetting "Capacity.SolverMaxIterations", "60"
    SetSystemSetting "Solver.Method", "Newton"
    SetSystemSetting "Solver.DirectState.DiagramExtension", "Yes"
    SetSystemSetting "Solver.MaxIterations", "80"
    SetSystemSetting "Solver.LoadSteps", "1"
    SetSystemSetting "Mesh.StepX", "50"
    SetSystemSetting "Mesh.StepY", "50"
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
    loads.Cells.Item(2, 5).Value2 = "PR1"
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
    SetSystemSetting "Solver.Method", "Newton"
    SetSystemSetting "Solver.DirectState.DiagramExtension", "Yes"
    SetSystemSetting "Solver.MaxIterations", "80"
    SetSystemSetting "Solver.LoadSteps", "1"
    SetSystemSetting "Mesh.StepX", "50"
    SetSystemSetting "Mesh.StepY", "50"
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












