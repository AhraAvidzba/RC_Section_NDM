Attribute VB_Name = "modWorkbookCalculation"
Option Explicit

' ==========================================================================
' Точки входа Excel-макросов
' ==========================================================================
' Модуль связывает кнопки книги с архитектурными слоями: чтение Config,
' построение CSectionModel, запуск batch-расчета, запись Results, обновление
' схемы и экспорт в AutoCAD. Сложная математика остается в классах solver-ов,
' а этот модуль держит пользовательский сценарий целиком.

Public Sub RunSectionCalculation()
    On Error GoTo Failed
    Dim message As String
    message = RunSectionCalculationForWorkbook(ThisWorkbook, True)
    If InStr(1, message, "ошиб", vbTextCompare) > 0 Or InStr(1, message, "InputErr", vbTextCompare) > 0 Then
        MsgBox message, vbExclamation, "RC Section NDM"
    ElseIf NonCriticalMessagesEnabled(ThisWorkbook) Then
        MsgBox message, vbInformation, "RC Section NDM"
    End If
    Exit Sub

Failed:
    MsgBox "Расчет не выполнен: " & Err.Description, vbExclamation, "RC Section NDM"
End Sub

' Очищает накопленное состояние перед новым расчетом или повторным формированием вывода.
Public Sub ClearSectionResults()
    On Error GoTo Failed
    ClearSectionResultsForWorkbook ThisWorkbook
    If NonCriticalMessagesEnabled(ThisWorkbook) Then
        MsgBox "Результаты и диагностика очищены. Исходные данные не изменены.", vbInformation, "RC Section NDM"
    End If
    Exit Sub

Failed:
    MsgBox "Не удалось очистить результаты: " & Err.Description, vbExclamation, "RC Section NDM"
End Sub

Public Sub UpdateSectionPlot()
    On Error GoTo Failed
    UpdateSectionPlotForWorkbook ThisWorkbook
    If NonCriticalMessagesEnabled(ThisWorkbook) Then
        MsgBox "Схема сечения обновлена по последнему расчетному снимку Results.", vbInformation, "RC Section NDM"
    End If
    Exit Sub

Failed:
    MsgBox "Схема не обновлена: " & Err.Description, vbExclamation, "RC Section NDM"
End Sub

Public Sub ImportGeometryFromAutoCAD()
    On Error GoTo Failed
    Dim message As String
    message = ImportGeometryFromAutoCADForWorkbook(ThisWorkbook)
    If NonCriticalMessagesEnabled(ThisWorkbook) Then MsgBox message, vbInformation, "RC Section NDM"
    Exit Sub

Failed:
    MsgBox "Импорт геометрии из AutoCAD не выполнен: " & Err.Description, vbExclamation, "RC Section NDM"
End Sub

Public Sub UpdateSectionPlotForWorkbook(ByVal workbook As Object, Optional ByVal raiseIfNoData As Boolean = True)
    If workbook Is Nothing Then Err.Raise vbObjectError + 4140, "UpdateSectionPlotForWorkbook", "Книга Excel не передана."

    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook workbook

    If Not settings.GetBoolean("Plot.Enabled", True) Then Exit Sub

    Dim reader As CSectionPlotDataReader
    Set reader = New CSectionPlotDataReader
    Dim plotLoadErrorNumber As Long
    Dim plotLoadErrorDescription As String
    If Not TryLoadFullPlotReader(reader, workbook, settings, plotLoadErrorNumber, plotLoadErrorDescription) Then
        If plotLoadErrorNumber = vbObjectError + 4706 Then
            Set reader = New CSectionPlotDataReader
            reader.LoadGeometryOnlyForMissingState workbook, settings, plotLoadErrorDescription
        Else
            If Not CanUseGeometryPreview(workbook, settings) Then
                ClearSectionPlotForNoData workbook, _
                    "Схема не обновлена: в Results нет выбранного состояния НДС."
                If raiseIfNoData Then
                    Err.Raise vbObjectError + 4144, "UpdateSectionPlotForWorkbook", _
                        "В Results нет выбранного состояния для схемы. Выполните расчет или проверьте ProfileId и Visualization.State."
                End If
                Exit Sub
            End If

            Set reader = New CSectionPlotDataReader
            reader.LoadGeometryPreviewFromWorkbook workbook, settings
        End If
    End If

    Dim plotter As CSectionPlotter
    Set plotter = New CSectionPlotter
    plotter.Draw workbook, reader, settings
End Sub

' Сначала пытаемся построить обычную схему по расчетным результатам.
' Если расчета еще нет, но после AutoCAD-import уже сохранена геометрия,
' возвращаем False: вызывающий код построит geometry-preview с текущими
' настройками аннотаций, не запуская solver и не меняя Results.
Private Function TryLoadFullPlotReader(ByVal reader As CSectionPlotDataReader, _
        ByVal workbook As Object, ByVal settings As CSystemSettingsReader, _
        ByRef errorNumber As Long, ByRef errorDescription As String) As Boolean
    On Error GoTo Failed
    errorNumber = 0
    errorDescription = vbNullString
    reader.LoadFromWorkbook workbook, settings
    TryLoadFullPlotReader = True
    Exit Function

Failed:
    If Err.Number = vbObjectError + 4702 Or Err.Number = vbObjectError + 4705 Or _
            Err.Number = vbObjectError + 4706 Or Err.Number = vbObjectError + 4710 Then
        errorNumber = Err.Number
        errorDescription = Err.Description
        TryLoadFullPlotReader = False
    Else
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Function

' Проверяет, можно ли вместо полноценной расчетной схемы показать preview
' импортированной AutoCAD-геометрии. Generated-сценарий сюда не допускается:
' иначе AutoCAD-preview может остаться на листе после переключения
' настроек на автоматическую генерацию сечения.
Private Function CanUseGeometryPreview(ByVal workbook As Object, ByVal settings As CSystemSettingsReader) As Boolean
    On Error GoTo Failed
    If workbook Is Nothing Then Exit Function
    If settings Is Nothing Then Exit Function
    If StrComp(settings.GetRawString("Geometry.Source", "Generated"), "AutoCAD", vbTextCompare) <> 0 Then Exit Function

    CanUseGeometryPreview = (StrComp(ResultsGeometrySource(workbook), "AutoCADImport", vbTextCompare) = 0)
    Exit Function
Failed:
End Function

' Очищает содержимое существующего ChartObject, но не удаляет само окно схемы.
' Это нужно, когда расчет не дал ни одного доступного состояния LC: пользователь
' не должен видеть подпись или AutoCAD-preview схему как будто она
' относится к текущим Generated-настройкам.
Private Sub ClearSectionPlotForNoData(ByVal workbook As Object, ByVal titleText As String)
    On Error GoTo Done
    Dim plotter As CSectionPlotter
    Set plotter = New CSectionPlotter
    plotter.ClearExisting workbook, titleText
Done:
End Sub

Public Sub UpdateSectionGeometryPreviewForWorkbook(ByVal workbook As Object)
    If workbook Is Nothing Then Err.Raise vbObjectError + 4141, "UpdateSectionGeometryPreviewForWorkbook", "Книга Excel не передана."

    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook workbook

    Dim reader As CSectionPlotDataReader
    Set reader = New CSectionPlotDataReader
    reader.LoadGeometryPreviewFromWorkbook workbook, settings

    Dim plotter As CSectionPlotter
    Set plotter = New CSectionPlotter
    plotter.Draw workbook, reader, settings
End Sub

' Импортирует AutoCAD Region в CSectionModel и сохраняет geometry-only snapshot.
' Расчет при Geometry.Source = AutoCAD затем использует этот снимок, а не
' обращается к AutoCAD повторно.
Public Function ImportGeometryFromAutoCADForWorkbook(ByVal workbook As Object) As String
    If workbook Is Nothing Then Err.Raise vbObjectError + 4142, "ImportGeometryFromAutoCADForWorkbook", "Книга Excel не передана."

    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook workbook

    If StrComp(settings.GetRawString("Geometry.Source", "Generated"), "AutoCAD", vbTextCompare) <> 0 Then
        Err.Raise vbObjectError + 4143, "ImportGeometryFromAutoCADForWorkbook", _
            "Для импорта выберите Geometry.Source = AutoCAD на листе Config."
    End If

    Dim units As CUnitSystem
    Set units = New CUnitSystem
    units.LoadFromSettings settings

    Dim importer As CAutoCADSectionModelImporter
    Set importer = New CAutoCADSectionModelImporter

    Dim section As CSectionModel
    Set section = importer.ImportFromActiveDocument(settings, units)
    ' Согласуем geometry-preview с расчетным путем: изотропные Region без
    ' собственного угла получают среднюю ориентацию распознанной сетки до
    ' записи snapshot, как это уже происходит при чтении Results для расчета.
    section.ApplyAverageRotationToEquivalentAreaFallbacks

    Dim writer As CNDMResultsWriter
    Set writer = New CNDMResultsWriter
    writer.WriteGeometryPreview workbook, section, units

    UpdateSectionGeometryPreviewForWorkbook workbook

    ImportGeometryFromAutoCADForWorkbook = "Геометрия успешно импортирована из AutoCAD." & vbCrLf & _
        "Бетонных Region: " & CStr(section.ConcreteCount) & "; арматурных Region: " & _
        CStr(section.RebarCount) & "." & vbCrLf & _
        "На схеме показаны только импортированные элементы для визуального контроля."
End Function

' Возвращает пользовательское решение о показе обычных информационных окон.
' Ошибки и предупреждения эта настройка не гасит: она нужна только для
' сообщений об успешно завершенных действиях, которые могут мешать серии запусков.
Public Function NonCriticalMessagesEnabled(ByVal workbook As Object) As Boolean
    On Error GoTo DefaultEnabled
    If workbook Is Nothing Then GoTo DefaultEnabled

    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook workbook

    NonCriticalMessagesEnabled = settings.GetBoolean("General.NonCriticalMessagesEnabled", True)
    Exit Function

DefaultEnabled:
    NonCriticalMessagesEnabled = True
End Function

' Запускает связанный набор операций и возвращает пользователю итоговый статус выполнения.
Public Function RunSectionCalculationForWorkbook(ByVal workbook As Object, Optional ByVal showMessages As Boolean = False) As String
    On Error GoTo Failed
    If workbook Is Nothing Then Err.Raise vbObjectError + 4100, "RunSectionCalculationForWorkbook", "Книга Excel не передана."

    Dim excelGuard As CExcelAppStateGuard
    Set excelGuard = New CExcelAppStateGuard
    excelGuard.Enter workbook.Application

    Dim runStart As Double
    runStart = Timer

    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook workbook

    Dim report As CExecutionReport
    Set report = New CExecutionReport
    report.Initialize workbook, settings
    report.AddSection "Старт расчета"
    report.AddStep "Запущен макрос расчета сечения."
    report.AddStep "Лист Config прочитан."
    report.AddValue "Загружено настроек", CStr(settings.KeyCount)
    report.AddValue "Дубликатов настроек", CStr(settings.DuplicateCount)
    report.AddValue "Источник геометрии", settings.GetRawString("Geometry.Source", "Generated")
    report.AddValue "Тип геометрии", settings.GetRawString("Geometry.Type", "-")
    report.AddValue "Профили расчета", "rngCalculationProfiles"
    
    Dim units As CUnitSystem
    Set units = New CUnitSystem
    report.AddSection "Единицы и знаки"
    report.AddStep "Создан CUnitSystem."
    units.LoadFromSettings settings
    report.AddValue "INPUT Length/Area/Force/Moment/Stress/Curvature", _
        settings.GetString("Units.Length.Input", "-") & "; " & _
        settings.GetString("Units.Area.Input", "-") & "; " & _
        settings.GetString("Units.Force.Input", "-") & "; " & _
        settings.GetString("Units.Moment.Input", "-") & "; " & _
        settings.GetString("Units.Stress.Input", "-") & "; " & _
        settings.GetString("Units.Curvature.Input", "-")
    report.AddValue "OUTPUT Length/Area/Force/Moment/Stress/Curvature", _
        units.OutputLengthUnit & "; " & units.OutputAreaUnit & "; " & _
        units.OutputForceUnit & "; " & units.OutputMomentUnit & "; " & _
        units.OutputStressUnit & "; " & units.OutputCurvatureUnit
    report.AddValue "Знаки OUTPUT N/Mx/My", units.OutputSignN & "; " & units.OutputSignMx & "; " & units.OutputSignMy

    Dim section As CSectionModel
    report.AddSection "Геометрия и материалы"
    report.AddStep "Начато построение CSectionModel."
    report.AddValue "Geometry.Source", settings.GetRawString("Geometry.Source", "Generated")
    report.AddValue "Geometry.Type", settings.GetRawString("Geometry.Type", "-")
    report.AddValue "Mesh.StepX", settings.GetRawString("Mesh.StepX", settings.GetRawString("Mesh.Step", "-"))
    report.AddValue "Mesh.StepY", settings.GetRawString("Mesh.StepY", settings.GetRawString("Mesh.StepX", settings.GetRawString("Mesh.Step", "-")))
    report.AddValue "Mesh.BoundarySubdivisions", settings.GetRawString("Mesh.BoundarySubdivisions", "-")
    Set section = BuildWorkbookSectionModel(workbook, settings, units)
    report.AddStep "CSectionModel построен."
    report.AddValue "Бетонных элементов", CStr(section.ConcreteCount)
    report.AddValue "Стержней арматуры", CStr(section.RebarCount)
    report.AddValue "Semantic-аннотаций", CStr(section.AnnotationCount)

    Dim materialProvider As CMaterialModelProvider
    Set materialProvider = New CMaterialModelProvider
    report.AddStep "Начато построение material provider."
    materialProvider.Initialize settings, units
    report.AddStep "Построены материальные модели Strength, CrackInitiation и CrackedNDS по параметрам Config."

    Dim profiles As CCalculationProfileCatalog
    Set profiles = New CCalculationProfileCatalog
    profiles.LoadFromWorkbook workbook
    report.AddStep "Загружен каталог расчетных профилей."
    report.AddValue "Профилей в rngCalculationProfiles", CStr(profiles.Count)

    Dim batch As CBatchSectionCalculator
    Set batch = New CBatchSectionCalculator
    report.AddStep "Создан CBatchSectionCalculator."
    batch.Initialize section, materialProvider
    Set batch.ProfileCatalog = profiles
    report.AddStep "Batch инициализирован моделью сечения и материалами."
    batch.ApplySettings settings, units
    report.AddStep "Настройки solver/capacity/crack применены к batch."
    Set batch.ExecutionReport = report

    Dim reader As CLoadCombinationReader
    Set reader = New CLoadCombinationReader
    report.AddSection "Сочетания нагрузок"
    report.AddStep "Начато чтение rngLoadCombinations."
    reader.LoadFromWorkbook workbook, batch, units
    If batch.Count = 0 Then Err.Raise vbObjectError + 4101, "RunSectionCalculationForWorkbook", "Не задано ни одного сочетания нагрузок."
    report.AddValue "Прочитано сочетаний", CStr(batch.Count)
    report.AddBlock "Список сочетаний", CombinationListForReport(batch)
    LoadStabilityDurationLoadsFromWorkbook workbook, batch, units
    batch.SetSP35Table721 ReadSP35Table721FromWorkbook(workbook)
    report.AddStep "Прочитаны нагрузки и таблицы для расчета устойчивости."

    report.AddSection "Точка приложения нагрузки"
    report.AddStep "Расчет центра тяжести бетонного сечения для пользовательской точки нагрузки."
    ApplyLoadReferenceFromSettings section, settings, units, batch
    report.AddValue "Точка приложения нагрузки X", FormatReportNumber(batch.LoadReferenceX) & " мм"
    report.AddValue "Точка приложения нагрузки Y", FormatReportNumber(batch.LoadReferenceY) & " мм"

    report.AddSection "Проверка вывода Results"
    ValidateResultsOutputLayout workbook, section, batch, profiles
    report.AddStep "Проверена раскладка Results для " & CStr(batch.Count) & " сочетаний."

    report.AddSection "Подготовка книги"
    report.AddStep "Начата очистка старых результатов."
    ClearSectionResultsForWorkbook workbook
    report.AddStep "Очищены rngBatchSummary и таблицы расчетного снимка Results."

    report.AddSection "Расчет сочетаний"
    batch.Execute
    report.AddValue "Время batch-расчета", FormatReportNumber(batch.ElapsedSeconds) & " с"

    Dim summaryWriter As CBatchResultWriter
    Set summaryWriter = New CBatchResultWriter
    report.AddSection "Запись результатов"
    report.AddStep "Начата запись batch summary."
    summaryWriter.WriteSummary workbook, batch, units, section
    report.AddStep "Сводка batch summary записана на лист Results."

    Dim ndmWriter As CNDMResultsWriter
    Set ndmWriter = New CNDMResultsWriter
    report.AddStep "Начата запись расчетного snapshot NDM."
    ndmWriter.WriteResults workbook, section, materialProvider, batch, units
    report.AddStep "Расчетный снимок NDM записан на лист Results."

    If settings.GetBoolean("Plot.AutoUpdateAfterCalculation", True) And batch.StateAvailableCount > 0 Then
        report.AddSection "Схема"
        report.AddStep "Начато обновление схемы сечения по Results."
        UpdateSectionPlotForWorkbook workbook, False
        report.AddStep "Схема сечения обновлена по сохраненному снимку Results."
    Else
        report.AddSection "Схема"
        If settings.GetBoolean("Plot.AutoUpdateAfterCalculation", True) Then
            ClearSectionPlotForNoData workbook, _
                "Схема не обновлена: нет доступных расчетных состояний LC."
            report.AddStep "Схема очищена: нет доступных named-state в Results."
        Else
            report.AddStep "Автообновление схемы пропущено: Plot.AutoUpdateAfterCalculation = No."
        End If
    End If

    Dim totalElapsedSeconds As Double
    totalElapsedSeconds = ElapsedSecondsFrom(runStart)
    summaryWriter.UpdateElapsedSeconds workbook, totalElapsedSeconds
    report.AddValue "Полное время выполнения макроса", FormatReportNumber(totalElapsedSeconds) & " с"

    report.AddSection "Финальное сообщение"
    report.AddStep "Формируется пользовательское сообщение о завершении."
    RunSectionCalculationForWorkbook = BuildCalculationMessage(section, settings, batch)
    If report.Enabled Then
        report.AddStep "Отчет сохраняется в txt-файл рядом с книгой."
        report.Save RunSectionCalculationForWorkbook
        summaryWriter.UpdateElapsedSeconds workbook, ElapsedSecondsFrom(runStart)
        RunSectionCalculationForWorkbook = RunSectionCalculationForWorkbook & vbCrLf & _
            "Пошаговый отчет сохранен: " & report.FilePath
    End If
    CalculateResultsSheet workbook
    excelGuard.Restore
    Exit Function

Failed:
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    If Not report Is Nothing Then
        report.AddError "Расчет остановлен", errorDescription
        report.Save "Расчет не выполнен: " & errorDescription
    End If
    If Not excelGuard Is Nothing Then excelGuard.Restore
    Err.Raise errorNumber, errorSource, errorDescription
End Function

' Проверяет, поместятся ли все блоки Results до запуска solver-а.
' Учитываются шапки над якорями, фактические ширины writer-ов и нижние
' snapshot-блоки, между которыми по принятому правилу должны оставаться
' два пустых столбца. Если раскладка тесная, расчет не запускается.
Private Sub ValidateResultsOutputLayout(ByVal workbook As Object, ByVal section As CSectionModel, _
        ByVal batch As CBatchSectionCalculator, ByVal profiles As CCalculationProfileCatalog)
    If workbook Is Nothing Then Err.Raise vbObjectError + 4160, "ValidateResultsOutputLayout", "Книга Excel не передана."
    If section Is Nothing Then Err.Raise vbObjectError + 4161, "ValidateResultsOutputLayout", "Модель сечения не передана."
    If batch Is Nothing Then Err.Raise vbObjectError + 4162, "ValidateResultsOutputLayout", "Пакетный расчетчик не передан."

    Const SNAPSHOT_GAP_COLUMNS As Long = 2
    Const RESULTS_TABLE_GAP_ROWS As Long = 2

    Dim issues As Collection
    Set issues = New Collection

    Dim summaryWriter As CBatchResultWriter
    Dim strengthWriter As CStrengthSummaryWriter
    Dim crackWriter As CCrackSummaryWriter
    Dim stabilityWriter As CStabilitySummaryWriter
    Dim ndmWriter As CNDMResultsWriter
    Set summaryWriter = New CBatchResultWriter
    Set strengthWriter = New CStrengthSummaryWriter
    Set crackWriter = New CCrackSummaryWriter
    Set stabilityWriter = New CStabilitySummaryWriter
    Set ndmWriter = New CNDMResultsWriter

    Dim summaryAnchor As Object
    Dim strengthAnchor As Object
    Dim crackAnchor As Object
    Dim stabilityAnchor As Object
    Dim elementAnchor As Object
    On Error GoTo MissingAnchor
    Set summaryAnchor = workbook.Names.Item("rngBatchSummary").RefersToRange(1, 1)
    Set strengthAnchor = workbook.Names.Item("rngStrengthSummaryAnchor").RefersToRange
    Set crackAnchor = workbook.Names.Item("rngCrackSummaryAnchor").RefersToRange
    Set stabilityAnchor = workbook.Names.Item("rngStabilitySummaryAnchor").RefersToRange
    Set elementAnchor = workbook.Names.Item("rngNDMElementResults").RefersToRange
    On Error GoTo 0

    Dim ws As Object
    Set ws = summaryAnchor.Worksheet
    If Not (strengthAnchor.Worksheet Is ws) Or Not (crackAnchor.Worksheet Is ws) Or _
            Not (stabilityAnchor.Worksheet Is ws) Or Not (elementAnchor.Worksheet Is ws) Then
        AddLayoutIssue issues, "rngBatchSummary, rngStrengthSummaryAnchor, rngCrackSummaryAnchor, rngStabilitySummaryAnchor и rngNDMElementResults должны находиться на одном листе Results."
    End If

    Dim summaryRows As Long
    Dim summaryCols As Long
    summaryRows = summaryWriter.RequiredSummaryOutputRows(batch.Count)
    summaryCols = summaryWriter.RequiredSummaryOutputColumns()
    CheckFootprintWithinSheet issues, ws, "rngBatchSummary", _
        summaryAnchor.Row, summaryAnchor.Column, summaryRows, summaryCols

    Dim strengthTopRow As Long
    Dim strengthRows As Long
    Dim strengthCols As Long
    strengthTopRow = strengthAnchor.Row - strengthWriter.HeaderRowsAboveAnchor
    strengthRows = strengthWriter.RequiredRowsForWorkbook(workbook, batch.Count)
    strengthCols = strengthWriter.RequiredColumns()
    If strengthTopRow < 1 Then
        AddLayoutIssue issues, "rngStrengthSummaryAnchor расположен слишком высоко: над ним нет места для шапки таблицы прочности."
    Else
        CheckFootprintWithinSheet issues, ws, "rngStrengthSummaryAnchor с шапкой прочности", _
            strengthTopRow, strengthAnchor.Column, strengthRows, strengthCols
    End If

    Dim crackTopRow As Long
    Dim crackRows As Long
    Dim crackCols As Long
    crackTopRow = crackAnchor.Row - crackWriter.HeaderRowsAboveAnchor
    crackRows = crackWriter.RequiredRowsForWorkbook(workbook, batch.Count)
    crackCols = crackWriter.RequiredColumns()
    If crackTopRow < 1 Then
        AddLayoutIssue issues, "rngCrackSummaryAnchor расположен слишком высоко: над ним нет места для шапки таблицы трещин."
    Else
        CheckFootprintWithinSheet issues, ws, "rngCrackSummaryAnchor с шапкой трещин", _
            crackTopRow, crackAnchor.Column, crackRows, crackCols
    End If

    Dim stabilityTopRow As Long
    Dim stabilityRows As Long
    Dim stabilityCols As Long
    stabilityTopRow = stabilityAnchor.Row - stabilityWriter.HeaderRowsAboveAnchor
    stabilityRows = stabilityWriter.RequiredRowsForWorkbook(workbook, batch.Count)
    stabilityCols = stabilityWriter.RequiredColumns()
    If stabilityTopRow < 1 Then
        AddLayoutIssue issues, "rngStabilitySummaryAnchor расположен слишком высоко: над ним нет места для шапки таблицы устойчивости."
    Else
        CheckFootprintWithinSheet issues, ws, "rngStabilitySummaryAnchor с шапкой устойчивости", _
            stabilityTopRow, stabilityAnchor.Column, stabilityRows, stabilityCols
    End If

    Dim summaryBottomRow As Long
    Dim strengthBottomRow As Long
    Dim crackBottomRow As Long
    Dim stabilityBottomRow As Long
    Dim snapshotTitleRow As Long
    summaryBottomRow = summaryAnchor.Row + summaryRows - 1
    strengthBottomRow = strengthTopRow + strengthRows - 1
    crackBottomRow = crackTopRow + crackRows - 1
    stabilityBottomRow = stabilityTopRow + stabilityRows - 1
    snapshotTitleRow = elementAnchor.Row - 2
    If summaryBottomRow + RESULTS_TABLE_GAP_ROWS >= strengthTopRow Then
        AddLayoutIssue issues, "Между rngBatchSummary и шапкой rngStrengthSummaryAnchor нужны " & _
            CStr(RESULTS_TABLE_GAP_ROWS) & " пустые строки: для " & _
            CStr(batch.Count) & " LC она занимает строки до " & CStr(summaryBottomRow) & _
            ", шапка прочности начинается со строки " & CStr(strengthTopRow) & _
            ". Опустите rngStrengthSummaryAnchor или сократите число строк rngLoadCombinations."
    End If
    If strengthBottomRow + RESULTS_TABLE_GAP_ROWS >= crackTopRow Then
        AddLayoutIssue issues, "Между rngStrengthSummaryAnchor и шапкой rngCrackSummaryAnchor нужны " & _
            CStr(RESULTS_TABLE_GAP_ROWS) & " пустые строки: для " & _
            CStr(batch.Count) & " LC таблица прочности занимает строки до " & CStr(strengthBottomRow) & _
            ", шапка трещин начинается со строки " & CStr(crackTopRow) & _
            ". Опустите rngCrackSummaryAnchor или сократите число строк rngLoadCombinations."
    End If
    If crackBottomRow + RESULTS_TABLE_GAP_ROWS >= stabilityTopRow Then
        AddLayoutIssue issues, "Между rngCrackSummaryAnchor и шапкой rngStabilitySummaryAnchor нужны " & _
            CStr(RESULTS_TABLE_GAP_ROWS) & " пустые строки: для " & _
            CStr(batch.Count) & " LC таблица трещин занимает строки до " & CStr(crackBottomRow) & _
            ", шапка устойчивости начинается со строки " & CStr(stabilityTopRow) & _
            ". Опустите rngStabilitySummaryAnchor или сократите число строк rngLoadCombinations."
    End If
    If stabilityBottomRow + RESULTS_TABLE_GAP_ROWS >= snapshotTitleRow Then
        AddLayoutIssue issues, "Между rngStabilitySummaryAnchor и общей строкой нижнего snapshot нужны " & _
            CStr(RESULTS_TABLE_GAP_ROWS) & " пустые строки: для " & _
            CStr(batch.Count) & " LC она занимает строки до " & CStr(stabilityBottomRow) & _
            ", а общий заголовок snapshot начинается со строки " & CStr(snapshotTitleRow) & _
            " перед rngNDMElementResults на строке " & CStr(elementAnchor.Row) & _
            ". Опустите rngNDMElementResults и нижние snapshot-якоря или сократите число строк rngLoadCombinations."
    End If

    Dim estimatedStateCount As Long
    Dim elementCount As Long
    estimatedStateCount = EstimatedNamedStateCountForOutput(batch, profiles)
    elementCount = section.ConcreteCount + section.RebarCount
    ValidateSnapshotOutputLayout issues, workbook, ws, ndmWriter, elementCount, _
        estimatedStateCount, batch.Count, section.AnnotationCount, SNAPSHOT_GAP_COLUMNS

    If issues.Count > 0 Then
        Err.Raise vbObjectError + 4163, "ValidateResultsOutputLayout", _
            ResultsOutputLayoutMessage(batch.Count, issues)
    End If
    Exit Sub

MissingAnchor:
    Err.Raise vbObjectError + 4164, "ValidateResultsOutputLayout", _
        "На листе Results не найден один из обязательных якорей вывода: rngBatchSummary, rngStrengthSummaryAnchor, rngCrackSummaryAnchor, rngStabilitySummaryAnchor или rngNDMElementResults."
End Sub

' Проверяет нижние соседние блоки Results по строкам листа и по ширине.
Private Sub ValidateSnapshotOutputLayout(ByVal issues As Collection, ByVal workbook As Object, _
        ByVal ws As Object, ByVal ndmWriter As CNDMResultsWriter, ByVal elementCount As Long, _
        ByVal estimatedStateCount As Long, ByVal combinationCount As Long, _
        ByVal annotationCount As Long, ByVal requiredGapColumns As Long)
    Dim names(1 To 5) As String
    Dim rows(1 To 5) As Long
    Dim cols(1 To 5) As Long
    names(1) = "rngNDMElementResults"
    names(2) = "rngNDMSectionGeometry"
    names(3) = "rngNDMSectionProperties"
    names(4) = "rngNDMMaterialDiagrams"
    names(5) = "rngNDMSectionAnnotations"

    rows(1) = ndmWriter.EstimatedElementResultRows(elementCount, estimatedStateCount)
    rows(2) = 1 + elementCount
    rows(3) = ndmWriter.EstimatedSectionPropertyRows(combinationCount, estimatedStateCount)
    rows(4) = 1 + MaxLong(1, estimatedStateCount) * 14
    rows(5) = 1 + MaxLong(2, annotationCount)

    Dim index As Long
    Dim previousRightColumn As Long
    previousRightColumn = 0
    For index = 1 To 5
        cols(index) = ndmWriter.OutputColumnCount(names(index))

        Dim anchor As Object
        On Error GoTo MissingSnapshotAnchor
        Set anchor = workbook.Names.Item(names(index)).RefersToRange
        On Error GoTo 0

        If Not (anchor.Worksheet Is ws) Then
            AddLayoutIssue issues, names(index) & " должен находиться на том же листе Results, что и остальные блоки вывода."
        End If
        CheckFootprintWithinSheet issues, ws, names(index), anchor.Row, anchor.Column, rows(index), cols(index)

        If previousRightColumn > 0 Then
            Dim gapColumns As Long
            gapColumns = anchor.Column - previousRightColumn - 1
            If gapColumns < requiredGapColumns Then
                AddLayoutIssue issues, "Между нижними блоками Results должно быть не меньше " & _
                    CStr(requiredGapColumns) & " пустых столбцов. Перед " & names(index) & _
                    " сейчас " & CStr(gapColumns) & ". Раздвиньте якоря вправо."
            End If
        End If
        previousRightColumn = anchor.Column + cols(index) - 1
    Next index
    Exit Sub

MissingSnapshotAnchor:
    AddLayoutIssue issues, "На листе Results не найден якорь " & names(index) & "."
    On Error GoTo 0
End Sub

' Проверяет, что прямоугольник вывода не выходит за пределы листа Excel.
Private Sub CheckFootprintWithinSheet(ByVal issues As Collection, ByVal ws As Object, _
        ByVal blockName As String, ByVal firstRow As Long, ByVal firstColumn As Long, _
        ByVal rowCount As Long, ByVal columnCount As Long)
    If rowCount < 1 Then rowCount = 1
    If columnCount < 1 Then columnCount = 1
    If firstRow < 1 Or firstColumn < 1 Then
        AddLayoutIssue issues, blockName & " имеет некорректный левый верхний угол вывода."
        Exit Sub
    End If

    If firstRow + rowCount - 1 > ws.Rows.Count Then
        AddLayoutIssue issues, blockName & " требует " & CStr(rowCount) & _
            " строк от строки " & CStr(firstRow) & ", но ниже не хватает строк листа."
    End If
    If firstColumn + columnCount - 1 > ws.Columns.Count Then
        AddLayoutIssue issues, blockName & " требует " & CStr(columnCount) & _
            " столбцов от столбца " & CStr(firstColumn) & ", но справа не хватает столбцов листа."
    End If
End Sub

' Оценивает максимальное число поэлементных named-state строк до расчета.
Private Function EstimatedNamedStateCountForOutput(ByVal batch As CBatchSectionCalculator, _
        ByVal profiles As CCalculationProfileCatalog) As Long
    If batch Is Nothing Then Exit Function

    Dim i As Long
    For i = 1 To batch.Count
        EstimatedNamedStateCountForOutput = EstimatedNamedStateCountForOutput + _
            EstimatedNamedStateCountForProfile(batch.ProfileId(i), profiles)
    Next i
End Function

' Возвращает верхнюю оценку named-state для профиля: трещины дают CrackedState
' и две возможные Mcrc-точки, даже если в конкретном LC они могут не понадобиться.
Private Function EstimatedNamedStateCountForProfile(ByVal profileId As String, _
        ByVal profiles As CCalculationProfileCatalog) As Long
    On Error GoTo UnknownProfile
    If profiles Is Nothing Then Exit Function
    If Not profiles.HasProfile(profileId) Then Exit Function

    Dim profile As CCalculationProfile
    Set profile = profiles.ProfileById(profileId)
    If profile.StrengthDirectStateEnabled Then EstimatedNamedStateCountForProfile = EstimatedNamedStateCountForProfile + 1
    If profile.StrengthCapacityEnabled Then EstimatedNamedStateCountForProfile = EstimatedNamedStateCountForProfile + 1
    If profile.CrackWidthEnabled Then EstimatedNamedStateCountForProfile = EstimatedNamedStateCountForProfile + 3
    Exit Function

UnknownProfile:
End Function

Private Sub AddLayoutIssue(ByVal issues As Collection, ByVal text As String)
    If issues Is Nothing Then Exit Sub
    issues.Add text
End Sub

Private Function ResultsOutputLayoutMessage(ByVal combinationCount As Long, ByVal issues As Collection) As String
    ResultsOutputLayoutMessage = "Results: не хватает места для " & _
        CStr(combinationCount) & " сочетаний. Расчет не запущен." & _
        vbCrLf & "Сдвиньте нужные якоря: rngStrengthSummaryAnchor, rngCrackSummaryAnchor, rngStabilitySummaryAnchor, rngNDMElementResults, rngNDMSectionGeometry, rngNDMSectionProperties, rngNDMMaterialDiagrams, rngNDMSectionAnnotations; либо уменьшите rngLoadCombinations." & _
        vbCrLf & JoinCollectionLines(issues)
End Function

Private Function MaxLong(ByVal a As Long, ByVal b As Long) As Long
    If a > b Then MaxLong = a Else MaxLong = b
End Function

' Пересчитывает только лист Results после блочной записи формул.
' Во время расчета Excel работает в ручном режиме, поэтому точечный пересчет
' нужен, чтобы пользователь сразу видел актуальные коэффициенты запаса, но
' чужие открытые книги не пересчитывались на каждом шаге вывода.
Private Sub CalculateResultsSheet(ByVal workbook As Object)
    On Error GoTo SafeExit
    workbook.Worksheets.Item("Results").Calculate
SafeExit:
End Sub

' Создает расчетный или интерфейсный объект из нормализованных исходных данных и локальных настроек.
Private Function BuildCalculationMessage(ByVal section As CSectionModel, ByVal settings As CSystemSettingsReader, _
        ByVal batch As CBatchSectionCalculator) As String
    Dim calculationCaption As String
    calculationCaption = "Расчет"
    If IsAutoCADSnapshotCalculation(section, settings) Then
        calculationCaption = "Расчет импортированной из AutoCAD геометрии"
    End If

    If batch.InvalidInputCount > 0 Then
        BuildCalculationMessage = calculationCaption & " завершен с ошибками ввода. Обработано сочетаний: " & _
            CStr(batch.Count) & "; ошибок ввода: " & CStr(batch.InvalidInputCount) & "." & vbCrLf & _
            "Проверьте строки со статусом InputErr на листе Results." & vbCrLf & _
            "Первая ошибка: " & batch.FirstInvalidInputMessage
    Else
        BuildCalculationMessage = calculationCaption & " завершен. Обработано сочетаний: " & CStr(batch.Count) & _
            ". Определяющее сочетание: " & batch.WorstCombinationID
    End If
End Function

' Отличает расчет обычной Generated-модели от расчета AutoCAD-snapshot.
' Здесь не выполняется импорт: функция только выбирает человеческий текст
' сообщения после того, как модель уже построена через BuildWorkbookSectionModel.
Private Function IsAutoCADSnapshotCalculation(ByVal section As CSectionModel, ByVal settings As CSystemSettingsReader) As Boolean
    If section Is Nothing Then Exit Function
    If settings Is Nothing Then Exit Function
    If StrComp(settings.GetRawString("Geometry.Source", "Generated"), "AutoCAD", vbTextCompare) <> 0 Then Exit Function

    IsAutoCADSnapshotCalculation = (StrComp(section.SourceType, "AutoCADImport", vbTextCompare) = 0)
End Function

' Собирает компактный список всех сочетаний, которые reader реально добавил
' в batch. В отчет попадают уже внутренние усилия после пересчета единиц и
' пользовательской системы знаков, потому что именно с ними работает solver.
Private Function CombinationListForReport(ByVal batch As CBatchSectionCalculator) As String
    If batch Is Nothing Then Exit Function

    Dim lines As Collection
    Set lines = New Collection

    Dim i As Long
    For i = 1 To batch.Count
        lines.Add CStr(i) & ". " & batch.CombinationID(i) & _
            "; profile=" & batch.ProfileId(i) & _
            "; N=" & FormatReportNumber(batch.N(i)) & " Н" & _
            "; Mx=" & FormatReportNumber(batch.Mx(i)) & " Н*мм" & _
            "; My=" & FormatReportNumber(batch.My(i)) & " Н*мм" & _
            "; comment=" & batch.CombinationName(i)
    Next i

    CombinationListForReport = JoinCollectionLines(lines)
End Function

' Склеивает строки коллекции через переносы. Это маленькая локальная утилита
' для текстового отчета, чтобы не плодить ручную конкатенацию с vbCrLf.
Private Function JoinCollectionLines(ByVal lines As Collection) As String
    If lines Is Nothing Then Exit Function
    If lines.Count = 0 Then Exit Function

    Dim parts() As String
    ReDim parts(0 To lines.Count - 1)

    Dim i As Long
    For i = 1 To lines.Count
        parts(i - 1) = CStr(lines.Item(i))
    Next i

    JoinCollectionLines = Join(parts, vbCrLf)
End Function

Private Sub ApplyLoadReferenceFromSettings(ByVal section As CSectionModel, _
        ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem, _
        ByVal batch As CBatchSectionCalculator)
    Dim referenceX As Double
    Dim referenceY As Double
    CalculateConcreteSectionCentroid section, referenceX, referenceY

    batch.ApplyLoadReference referenceX + units.InputLengthToInternal(settings.GetDouble("Load.ReferenceOffsetX", 0#)), _
        referenceY + units.InputLengthToInternal(settings.GetDouble("Load.ReferenceOffsetY", 0#)), referenceX, referenceY
End Sub

' Читает дополнительную таблицу нагрузок для устойчивости. Excel-слой сразу
' приводит силы и моменты к внутренним единицам, чтобы расчетный batch не
' обращался к листам и не знал пользовательских единиц.
Private Sub LoadStabilityDurationLoadsFromWorkbook(ByVal workbook As Object, _
        ByVal batch As CBatchSectionCalculator, ByVal units As CUnitSystem)
    On Error GoTo MissingRange
    If workbook Is Nothing Then Exit Sub
    If batch Is Nothing Then Exit Sub

    Dim range As Object
    Set range = workbook.Names.Item("rngStabilityDurationLoads").RefersToRange
    On Error GoTo 0
    Dim values As Variant
    values = range.Value2

    batch.ClearStabilityDurationLoads

    Dim rowIndex As Long
    For rowIndex = 2 To UBound(values, 1)
        Dim combinationID As String
        combinationID = Trim$(CStr(values(rowIndex, 1)))
        If Len(combinationID) > 0 Then
            batch.AddStabilityDurationLoad combinationID, _
                units.InputForceToInternal(NumericCellOrZero(values(rowIndex, 2))), _
                units.InputMomentMxToInternal(NumericCellOrZero(values(rowIndex, 3))), _
                units.InputMomentMyToInternal(NumericCellOrZero(values(rowIndex, 4)))
        End If
    Next rowIndex
    Exit Sub

MissingRange:
    batch.ClearStabilityDurationLoads
End Sub

' Возвращает таблицу 7.21 СП 35 как обычный массив Variant. Дальше она живет
' только в памяти и передается в CStabilityCalculator через batch.
Private Function ReadSP35Table721FromWorkbook(ByVal workbook As Object) As Variant
    On Error GoTo MissingRange
    If workbook Is Nothing Then Exit Function

    Dim range As Object
    Set range = workbook.Names.Item("rngSP35Table721").RefersToRange
    On Error GoTo 0
    ReadSP35Table721FromWorkbook = range.Value2
    Exit Function

MissingRange:
End Function

' Превращает пустую ячейку дополнительной таблицы нагрузок в 0. Ошибочные
' значения оставляем ошибкой исходных данных Excel, чтобы они не маскировались.
Private Function NumericCellOrZero(ByVal value As Variant) As Double
    If IsEmpty(value) Or IsNull(value) Then Exit Function
    If VarType(value) = vbString Then
        If Len(Trim$(CStr(value))) = 0 Then Exit Function
    End If
    NumericCellOrZero = CDbl(value)
End Function

' Возвращает центр тяжести бетонной части сечения.
' Это базовая точка пользовательских нагрузок: Load.ReferenceOffsetX/Y
' откладываются именно от центра бетона, а не от приведенного сечения
' бетон + арматура.
Public Sub CalculateConcreteSectionCentroid(ByVal section As CSectionModel, _
        ByRef referenceX As Double, ByRef referenceY As Double)
    Dim props As CSectionPropertiesCalculator
    Set props = New CSectionPropertiesCalculator
    props.CalculateConcrete section
    referenceX = props.CentroidX
    referenceY = props.CentroidY
End Sub

' Возвращает центр тяжести приведенного сечения бетон + арматура.
' Метод остается для справочных геометрических характеристик, главных осей и
' регрессионных проверок, но не используется как база пользовательского offset.
Public Sub CalculateTransformedSectionCentroid(ByVal section As CSectionModel, _
        ByVal concrete As Object, ByVal steel As Object, ByRef referenceX As Double, ByRef referenceY As Double)
    Dim props As CSectionPropertiesCalculator
    Set props = New CSectionPropertiesCalculator
    props.CalculateTransformed section, concrete, steel
    referenceX = props.CentroidX
    referenceY = props.CentroidY
End Sub

' Очищает накопленное состояние перед новым расчетом или повторным формированием вывода.
Public Sub ClearSectionResultsForWorkbook(ByVal workbook As Object)
    If workbook Is Nothing Then Err.Raise vbObjectError + 4110, "ClearSectionResultsForWorkbook", "Книга Excel не передана."
    Dim summaryWriter As CBatchResultWriter
    Set summaryWriter = New CBatchResultWriter
    summaryWriter.ClearSummary workbook
    Dim ndmWriter As CNDMResultsWriter
    Set ndmWriter = New CNDMResultsWriter
    ndmWriter.ClearResults workbook
End Sub

Public Function ReadWorkbookGeometry(ByVal workbook As Object, ByVal settings As CSystemSettingsReader, Optional ByVal units As CUnitSystem = Nothing) As ISectionGeometry
    If units Is Nothing Then
        Set units = New CUnitSystem
        units.InitializeDefaults
    End If
    Dim registry As CSectionTypeRegistry
    Set registry = New CSectionTypeRegistry
    Set ReadWorkbookGeometry = registry.CreateGeometry(settings, units)
End Function

' Создает расчетный или интерфейсный объект из нормализованных исходных данных и локальных настроек.
Public Function BuildWorkbookSectionModel(ByVal workbook As Object, ByVal settings As CSystemSettingsReader, Optional ByVal units As CUnitSystem = Nothing) As CSectionModel
    If units Is Nothing Then
        Set units = New CUnitSystem
        units.InitializeDefaults
    End If
    If workbook Is Nothing Then Err.Raise vbObjectError + 4130, "BuildWorkbookSectionModel", "Книга Excel не передана."
    If settings Is Nothing Then Err.Raise vbObjectError + 4131, "BuildWorkbookSectionModel", "Настройки Config не переданы."

    Dim geometrySource As String
    geometrySource = Trim$(settings.GetRawString("Geometry.Source", "Generated"))
    If Len(geometrySource) = 0 Then Err.Raise vbObjectError + 4132, "BuildWorkbookSectionModel", _
        "Geometry.Source должен быть Generated или AutoCAD."

    If StrComp(geometrySource, "Generated", vbTextCompare) = 0 Then
        Dim registry As CSectionTypeRegistry
        Set registry = New CSectionTypeRegistry
        Set BuildWorkbookSectionModel = registry.BuildGeneratedModel(settings, units)
    ElseIf StrComp(geometrySource, "AutoCAD", vbTextCompare) = 0 Then
        If StrComp(ResultsGeometrySource(workbook), "AutoCADImport", vbTextCompare) <> 0 Then
            Err.Raise vbObjectError + 4134, "BuildWorkbookSectionModel", _
                "Выбрано Geometry.Source = AutoCAD, но на листе Results нет предварительно импортированной AutoCAD-геометрии. " & _
                "Сначала нажмите кнопку ""Импортировать геометрию из AutoCAD""."
        End If
        Set BuildWorkbookSectionModel = ReadSectionGeometryFromResults(workbook, "AutoCADImport")
    Else
        Err.Raise vbObjectError + 4133, "BuildWorkbookSectionModel", _
            "Geometry.Source должен быть Generated или AutoCAD."
    End If
End Function

Public Function WorkbookMeshBoundarySubdivisions(ByVal settings As CSystemSettingsReader) As Long
    Dim registry As CSectionTypeRegistry
    Set registry = New CSectionTypeRegistry
    WorkbookMeshBoundarySubdivisions = registry.MeshBoundarySubdivisions(settings)
End Function

Public Function ReadWorkbookRebars(ByVal workbook As Object, ByVal geometry As ISectionGeometry, ByVal settings As CSystemSettingsReader, Optional ByVal units As CUnitSystem = Nothing) As CRebarLayout
    If units Is Nothing Then
        Set units = New CUnitSystem
        units.InitializeDefaults
    End If
    Dim registry As CSectionTypeRegistry
    Set registry = New CSectionTypeRegistry
    Set ReadWorkbookRebars = registry.CreateRebars(geometry, settings, units)
End Function

' Форматирует числа для человекочитаемого txt-отчета без зависимости от
' десятичного разделителя Windows. Это локальная утилита Excel-слоя: она
' нужна только для протокола запуска и не участвует в расчетных формулах.
Private Function FormatReportNumber(ByVal value As Double) As String
    FormatReportNumber = Replace$(Format$(value, "0.############"), ",", ".")
End Function

' Считает продолжительность пользовательского сценария с учетом возможного
' перехода Timer через полночь.
Private Function ElapsedSecondsFrom(ByVal startTimer As Double) As Double
    Dim nowTimer As Double
    nowTimer = Timer
    If nowTimer >= startTimer Then
        ElapsedSecondsFrom = nowTimer - startTimer
    Else
        ElapsedSecondsFrom = 86400# - startTimer + nowTimer
    End If
End Function
