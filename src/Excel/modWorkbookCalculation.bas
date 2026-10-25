Attribute VB_Name = "modWorkbookCalculation"
Option Explicit

' ==========================================================================
' Макросы кнопок и выполнение расчёта в Excel
' ==========================================================================
' Связывает отдельное формирование сечения, чтение Config, расчёт сочетаний и запись Results,
' а также обновление схемы. Геометрию и расчётные алгоритмы
' выполняют соответствующие классы; модуль задаёт порядок действий в книге.

Private Const RESULTS_TABLE_GAP_ROWS As Long = 2 ' Минимум пустых строк между крупными таблицами Results.

' Запускает полный сценарий кнопки расчета для текущей книги и показывает итог.
' Чтение, расчёт и запись выполняет RunSectionCalculationForWorkbook; ошибка выводится
' независимо от настройки необязательных информационных сообщений.
Public Sub RunSectionCalculation()
    On Error GoTo Failed
    Dim startedAt As Double: startedAt = Timer
    Dim informationEnabled As Boolean
    Dim message As String
    message = RunSectionCalculationForWorkbook(ThisWorkbook, True, startedAt, informationEnabled)
    If InStr(1, message, "ошиб", vbTextCompare) > 0 Or InStr(1, message, "InputErr", vbTextCompare) > 0 Or _
            InStr(1, message, "Предупреждение:", vbTextCompare) > 0 Then
        ShowWorkbookMessage message, vbExclamation, "RC Section NDM"
    ElseIf informationEnabled Then
        ShowWorkbookMessage message, vbInformation, "RC Section NDM"
    End If
    Exit Sub

Failed:
    ShowWorkbookMessage WorkbookErrorMessage("Расчет не выполнен", Err.Number, Err.Source, Err.Description), vbExclamation, "RC Section NDM"
End Sub

' Очищает Results по кнопке, сохраняя Config и таблицу нагрузок.
' Вызывает общую очистку указанной книги. Итоговое сообщение зависит
' от настройки; причина ошибки показывается всегда.
Public Sub ClearSectionResults()
    On Error GoTo Failed
    Dim informationEnabled As Boolean
    informationEnabled = NonCriticalMessagesEnabled(ThisWorkbook)
    ClearSectionResultsForWorkbook ThisWorkbook
    If informationEnabled Then
        ShowWorkbookMessage "Результаты и диагностика очищены. Исходные данные не изменены.", vbInformation, "RC Section NDM"
    End If
    Exit Sub

Failed:
    ShowWorkbookMessage WorkbookErrorMessage("Не удалось очистить результаты", Err.Number, Err.Source, Err.Description), vbExclamation, "RC Section NDM"
End Sub

' Обновляет схему по сохранённым Results и показывает причину ошибки.
' Новый расчёт не запускается; текущие исходные данные не заменяют сохранённые.
Public Sub UpdateSectionPlot()
    On Error GoTo Failed
    Dim informationEnabled As Boolean
    informationEnabled = NonCriticalMessagesEnabled(ThisWorkbook)
    UpdateSectionResultsForWorkbook ThisWorkbook
    If informationEnabled Then
        ShowWorkbookMessage "Схема и подробные результаты восстановлены из выбранного снимка без повторного расчёта.", vbInformation, "RC Section NDM"
    End If
    Exit Sub

Failed:
    ShowWorkbookMessage WorkbookErrorMessage("Схема не обновлена", Err.Number, Err.Source, Err.Description), vbExclamation, "RC Section NDM"
End Sub

' Формирует сечение по настройке Geometry.Source и сохраняет его для расчёта.
' Кнопка не запускает решение НДС; ошибки и предупреждения показываются всегда.
Public Sub FormSection()
    On Error GoTo Failed
    Dim startedAt As Double: startedAt = Timer
    Dim informationEnabled As Boolean
    Dim message As String
    message = FormSectionForWorkbook(ThisWorkbook, startedAt, informationEnabled)
    If InStr(1, message, "Предупреждени", vbTextCompare) > 0 Then
        ShowWorkbookMessage message, vbExclamation, "RC Section NDM"
    ElseIf informationEnabled Then
        ShowWorkbookMessage message, vbInformation, "RC Section NDM"
    End If
    Exit Sub

Failed:
    ShowWorkbookMessage WorkbookErrorMessage("Сечение не сформировано", Err.Number, Err.Source, Err.Description), vbExclamation, "RC Section NDM"
End Sub

' Читает выбранное НДС из Results и перерисовывает схему указанной книги.
' Если расчётных результатов ещё нет, может показать сформированную
' геометрию. НДС и таблицы Results при этом не меняются.
Public Sub UpdateSectionPlotForWorkbook(ByVal workbook As Object, Optional ByVal raiseIfNoData As Boolean = True)
    If workbook Is Nothing Then Err.Raise vbObjectError + 4140, "UpdateSectionPlotForWorkbook", "Книга Excel не передана."

    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook workbook

    If Not settings.GetRequiredBoolean("Plot.Enabled") Then Exit Sub

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

' Пытается прочитать данные обычной схемы с результатами расчёта.
' Если сечение сформировано, но ещё не рассчитано, возвращает False:
' вызывающий код покажет её с текущими настройками аннотаций без расчёта НДС.
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

' Разрешает показать сформированное сечение, пока результатов НДС ещё нет.
' Источник берётся из снимка: переключение Config само по себе его не заменяет.
Private Function CanUseGeometryPreview(ByVal workbook As Object, ByVal settings As CSystemSettingsReader) As Boolean
    On Error GoTo Failed
    If workbook Is Nothing Then Exit Function
    If settings Is Nothing Then Exit Function
    CanUseGeometryPreview = (Len(ResultsGeometrySource(workbook)) > 0 And _
        Len(ReadLoadCasePropertyText(workbook, "ALL", "LoadReferenceX")) = 0)
    Exit Function
Failed:
End Function

' Очищает содержимое существующего ChartObject без удаления окна схемы.
' Прежний рисунок не должен выглядеть как результат нового запуска,
' если пригодные данные отсутствуют или обновление схемы ещё не выполнено.
Private Sub ClearSectionPlotForNoData(ByVal workbook As Object, ByVal titleText As String)
    On Error GoTo Done
    Dim plotter As CSectionPlotter
    Set plotter = New CSectionPlotter
    plotter.ClearExisting workbook, titleText
Done:
End Sub

' Показывает сформированное сечение до расчёта НДС.
' Оформление читает из Config, координаты и характеристики — из Results.
Public Sub UpdateSectionGeometryPreviewForWorkbook(ByVal workbook As Object)
    If workbook Is Nothing Then Err.Raise vbObjectError + 4141, "UpdateSectionGeometryPreviewForWorkbook", "Книга Excel не передана."

    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook workbook

    If Not settings.GetRequiredBoolean("Plot.Enabled") Then Exit Sub

    Dim reader As CSectionPlotDataReader
    Set reader = New CSectionPlotDataReader
    reader.LoadGeometryPreviewFromWorkbook workbook, settings

    Dim plotter As CSectionPlotter
    Set plotter = New CSectionPlotter
    plotter.Draw workbook, reader, settings
End Sub


' Формирует сечение выбранным способом и публикует полный снимок в Results.
' Не читает сочетания и не выполняет расчёт; прежний снимок сохраняется,
' если получение новой геометрии или подготовка её характеристик завершились ошибкой.
Public Function FormSectionForWorkbook(ByVal workbook As Object, Optional ByVal startedAt As Double = -1#, _
        Optional ByRef informationEnabled As Boolean = False) As String
    On Error GoTo Failed
    If startedAt < 0# Then startedAt = Timer
    If workbook Is Nothing Then Err.Raise vbObjectError + 4142, "FormSectionForWorkbook", "Книга Excel не передана."
    Dim guard As CExcelAppStateGuard: Set guard = New CExcelAppStateGuard
    guard.Enter workbook.Application
    workbook.Application.Calculate ' Один пересчёт всех открытых книг этого Excel: актуализирует внешние формульные зависимости при отключённых событиях, до чтения входов.
    Dim settings As CSystemSettingsReader: Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook workbook
    informationEnabled = settings.GetRequiredBoolean("General.NonCriticalMessagesEnabled") ' Кнопка получает проверенный флаг из того же снимка; ошибка обнаруживается до изменения Results.
    Dim units As CUnitSystem: Set units = New CUnitSystem
    units.LoadFromSettings settings
    Dim fullFormatting As Boolean: fullFormatting = settings.GetRequiredBoolean("General.ReformatResultsEnabled")
    Dim section As CSectionModel, importer As CAutoCADSectionModelImporter
    Dim source As String: source = settings.GetRequiredChoice("Geometry.Source", Array("Generated", "AutoCAD"))
    If StrComp(source, "Generated", vbTextCompare) = 0 Then
        Set section = BuildWorkbookSectionModel(workbook, settings, units)
    Else
        Set importer = New CAutoCADSectionModelImporter
        Set section = importer.ImportFromActiveDocument(settings, units)
        section.ApplyAverageRotationToEquivalentAreaFallbacks
    End If
    Dim snapshot As CSectionPropertiesCalculator: Set snapshot = PrepareSectionSnapshot(section)
    Dim summary As CBatchResultWriter: Set summary = New CBatchResultWriter
    Dim writer As CNDMResultsWriter: Set writer = New CNDMResultsWriter
    Dim issues As Collection: Set issues = New Collection
    ValidateSnapshotOutputLayout issues, workbook, workbook.Worksheets.Item("Results"), writer, _
        section.ConcreteCount + section.RebarCount, 0, 0, section.AnnotationCount, section.Contours.Count, 2
    If issues.Count > 0 Then Err.Raise vbObjectError + 4163, "FormSectionForWorkbook", ResultsOutputLayoutMessage(0, issues)
    summary.ClearSummary workbook, fullFormatting
    writer.WriteGeometryPreview workbook, section, snapshot, units, fullFormatting
    ClearSectionPlotForNoData workbook, "Сечение сформировано. Схема не обновлена."
    If importer Is Nothing Then
        FormSectionForWorkbook = "Сечение сформировано и сохранено в Results." & vbCrLf & vbCrLf & _
            "Бетонных элементов: " & CStr(section.ConcreteCount) & "." & vbCrLf & _
            "Стержней арматуры: " & CStr(section.RebarCount) & "."
    Else
        FormSectionForWorkbook = importer.ResultMessage(section) & vbCrLf & vbCrLf & _
            "Сечение сформировано и сохранено в Results. Для расчёта нажмите «Выполнить расчёт»."
    End If
    If Len(section.GeometryWarningText) > 0 Then FormSectionForWorkbook = FormSectionForWorkbook & vbCrLf & section.GeometryWarningText
    CalculateResultsSheet workbook
    Dim plotErrorNumber As Long, plotError As String
    On Error Resume Next
    Err.Clear
    UpdateSectionGeometryPreviewForWorkbook workbook ' Оформление использует уже сохранённый снимок и не запускает расчёт. Ошибка схемы не отменяет успешное формирование сечения.
    plotErrorNumber = Err.Number: plotError = Err.Description
    Err.Clear
    On Error GoTo Failed
    If plotErrorNumber <> 0 Then
        FormSectionForWorkbook = FormSectionForWorkbook & vbCrLf & vbCrLf & _
            "Предупреждение:" & vbCrLf & "Схема не обновлена. Сечение сохранено в Results." & vbCrLf & vbCrLf & _
            SectionRecipeErrorText(plotError) & vbCrLf & vbCrLf & _
            "После исправления настроек оформления нажмите «Обновить схему и результаты»."
    End If
    guard.Restore True
    FormSectionForWorkbook = FormSectionForWorkbook & vbCrLf & vbCrLf & _
        "Время формирования сечения: " & Format$(ElapsedSecondsFrom(startedAt), "0.00") & " с."
    Exit Function
Failed:
    Dim number As Long, origin As String, description As String
    number = Err.Number: origin = Err.Source: description = Err.Description
    If Not guard Is Nothing Then guard.Restore
    description = SectionFormationFailureText(description, source, settings)
    Err.Raise number, origin, description
End Function

' Возвращает пользовательское решение о показе обычных информационных окон.
' Ошибки и предупреждения эта настройка не гасит: она нужна только для
' сообщений об успешно завершенных действиях, которые могут мешать серии запусков.
' Ошибка чтения остается адресной ошибкой Config; она не включает сообщения
' по умолчанию и обрабатывается внешним обработчиком кнопки как обычная ошибка.
Public Function NonCriticalMessagesEnabled(ByVal workbook As Object) As Boolean
    If workbook Is Nothing Then Err.Raise vbObjectError + 4140, "NonCriticalMessagesEnabled", "Книга Excel не передана."

    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook workbook

    NonCriticalMessagesEnabled = settings.GetRequiredBoolean("General.NonCriticalMessagesEnabled")
End Function

' Читает сформированное сечение из Results и текущие расчётные настройки.
' Подготавливает материалы один раз, рассчитывает сочетания и записывает Results.
' Возвращает итоговое сообщение; временные настройки Excel сохраняет CExcelAppStateGuard.
Public Function RunSectionCalculationForWorkbook(ByVal workbook As Object, Optional ByVal showMessages As Boolean = False, _
        Optional ByVal startedAt As Double = -1#, Optional ByRef informationEnabled As Boolean = False) As String
    On Error GoTo Failed
    Dim runStart As Double
    runStart = startedAt
    If runStart < 0# Then runStart = Timer
    If workbook Is Nothing Then Err.Raise vbObjectError + 4100, "RunSectionCalculationForWorkbook", "Книга Excel не передана."

    Dim excelGuard As CExcelAppStateGuard
    Set excelGuard = New CExcelAppStateGuard
    excelGuard.Enter workbook.Application
    workbook.Application.Calculate ' Один пересчёт всех открытых книг этого Excel: актуализирует внешние формульные зависимости при отключённых событиях, до чтения входов.

    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook workbook
    informationEnabled = settings.GetRequiredBoolean("General.NonCriticalMessagesEnabled") ' Кнопка получает проверенный флаг из того же снимка; ошибка обнаруживается до изменения Results.
    Dim fullFormatting As Boolean: fullFormatting = settings.GetRequiredBoolean("General.ReformatResultsEnabled")

    Dim report As CExecutionReport
    Set report = New CExecutionReport
    report.Initialize workbook, settings
    report.AddSection "Старт расчета"
    report.AddStep "Запущен макрос расчета сечения."
    report.AddStep "Лист Config прочитан."
    report.AddValue "Загружено настроек", CStr(settings.KeyCount)
    report.AddValue "Дубликатов настроек", CStr(settings.DuplicateCount)
    report.AddValue "Геометрия", "Сформированное сечение из Results"
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
    report.AddStep "Начато чтение сформированного сечения из Results."
    Dim ndmWriter As CNDMResultsWriter: Set ndmWriter = New CNDMResultsWriter
    Set section = ReadFormedSectionFromResults(workbook, False, ndmWriter)
    report.AddStep "CSectionModel прочитан из сохранённого снимка Results без генерации и импорта."
    report.AddValue "Источник сохранённого сечения", section.SourceType
    report.AddValue "Происхождение бетонной сетки", section.ConcreteMeshSource
    report.AddValue "Бетонных элементов", CStr(section.ConcreteCount)
    report.AddValue "Стержней арматуры", CStr(section.RebarCount)

    Dim materialProvider As CMaterialModelProvider
    Set materialProvider = New CMaterialModelProvider
    report.AddStep "Начато построение поставщика материалов."
    materialProvider.Initialize settings, units
    report.AddStep "Построены материальные модели Strength, CrackInitiation и CrackedNDS по параметрам Config."

    Dim profiles As CCalculationProfileCatalog
    Set profiles = New CCalculationProfileCatalog
    profiles.LoadFromWorkbook workbook, False
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
    If batch.Count = 0 Then
        RunSectionCalculationForWorkbook = "Расчёт не запускался: нет сочетаний с заполненным CombinationID и ненулевыми нагрузками. Прежние результаты сохранены."
        report.AddStep RunSectionCalculationForWorkbook
        report.Save RunSectionCalculationForWorkbook
        excelGuard.Restore True
        Exit Function
    End If
    report.AddValue "Прочитано сочетаний", CStr(batch.Count)
    report.AddBlock "Список сочетаний", CombinationListForReport(batch, units)
    LoadStabilityDurationLoadsFromWorkbook workbook, batch, units
    Dim sp35Table As Variant, sp35TableInputError As String
    sp35Table = ReadSP35Table721FromWorkbook(workbook, sp35TableInputError)
    batch.SetSP35Table721 sp35Table, sp35TableInputError
    report.AddStep "Прочитаны нагрузки и таблицы для расчета устойчивости."

    report.AddSection "Точка приложения нагрузки"
    report.AddStep "Расчет центра тяжести бетонного сечения для пользовательской точки нагрузки."
    ApplyLoadReferenceFromSettings section, settings, units, batch
    report.AddValue "Точка приложения нагрузки X", ReportLengthText(batch.LoadReferenceX, units)
    report.AddValue "Точка приложения нагрузки Y", ReportLengthText(batch.LoadReferenceY, units)

    report.AddSection "Проверка вывода Results"
    OfferExternalCalculationSnapshot workbook, CDbl(section.ConcreteCount + section.RebarCount) * EstimatedNamedStateCountForOutput(batch, profiles), showMessages
    Dim snapshotStore As CCalculationSnapshotStore: Set snapshotStore = New CCalculationSnapshotStore
    If snapshotStore.StorageMode(workbook) = "File" Then
        Dim validatedSnapshotPath As String: validatedSnapshotPath = snapshotStore.SnapshotPath(workbook)
    End If
    ValidateResultsOutputLayout workbook, section, batch, profiles
    report.AddStep "Проверена раскладка Results для " & CStr(batch.Count) & " сочетаний."

    report.AddSection "Подготовка книги"
    report.AddStep "Сформированное сечение и прежние результаты сохраняются до завершения вычислений."

    report.AddSection "Расчет сочетаний"
    Dim savedProfileValues As Variant: savedProfileValues = workbook.Names.Item("rngCalculationProfiles").RefersToRange.Value2 ' Замораживает исходный профиль до вычислений, без чтения Config при подготовке файла.
    batch.Execute
    report.AddValue "Время batch-расчета", FormatReportNumber(batch.ElapsedSeconds) & " с"
    report.AddValue "Количество стартов UltimateStrain", CStr(batch.UltimateStrainStartCount)
    report.AddValue "Количество итераций НДС", CStr(batch.SolverIterationCount)
    report.AddValue "Количество итераций UltimateStrain", CStr(batch.UltimateStrainIterationCount)

    Dim summaryWriter As CBatchResultWriter
    Set summaryWriter = New CBatchResultWriter
    report.AddSection "Запись результатов"
    Dim savedSnapshot As CCalculationSnapshot: Set savedSnapshot = New CCalculationSnapshot
    Set summaryWriter.CaptureSnapshot = savedSnapshot
    summaryWriter.WriteSummary workbook, batch, units, section, fullFormatting
    Set ndmWriter.CaptureSnapshot = savedSnapshot
    ndmWriter.WriteResults workbook, section, PrepareSectionSnapshot(section, materialProvider), batch, units, True, fullFormatting
    savedSnapshot.AddTable "Calculation.Profiles", savedProfileValues
    savedSnapshot.Validate
    report.AddStep "Результаты подготовлены; сохранение в выбранном режиме и вывод из памяти."
    Dim plotWarning As String
    plotWarning = SaveAndDisplayCalculationSnapshot(workbook, savedSnapshot)
    report.AddStep "Подробные таблицы и схема заполнены из готовых данных без повторного решения. Полный файл снимка сохраняется только в режиме File."

    Dim totalElapsedSeconds As Double
    totalElapsedSeconds = ElapsedSecondsFrom(runStart)
    summaryWriter.UpdateElapsedSeconds workbook, totalElapsedSeconds
    report.AddValue "Время до записи итогового отчёта", FormatReportNumber(totalElapsedSeconds) & " с"

    report.AddSection "Финальное сообщение"
    report.AddStep "Формируется пользовательское сообщение о завершении."
    RunSectionCalculationForWorkbook = BuildCalculationMessage(batch)
    If Len(plotWarning) > 0 Then RunSectionCalculationForWorkbook = RunSectionCalculationForWorkbook & vbCrLf & vbCrLf & plotWarning
    If report.Enabled Then
        report.AddStep "Отчет сохраняется в txt-файл рядом с книгой."
        report.Save RunSectionCalculationForWorkbook
        summaryWriter.UpdateElapsedSeconds workbook, ElapsedSecondsFrom(runStart)
        RunSectionCalculationForWorkbook = WorkbookMessageParagraph(RunSectionCalculationForWorkbook, _
            "Пошаговый отчет сохранен:" & vbCrLf & report.FilePath)
    End If
    CalculateResultsSheet workbook
    savedSnapshot.ElapsedSeconds = ElapsedSecondsFrom(runStart)
    snapshotStore.UpdateSavedElapsed workbook, savedSnapshot
    summaryWriter.UpdateElapsedSeconds workbook, savedSnapshot.ElapsedSeconds
    excelGuard.Restore True
    RunSectionCalculationForWorkbook = RunSectionCalculationForWorkbook & vbCrLf & vbCrLf & _
        "Время расчёта: " & Format$(ElapsedSecondsFrom(runStart), "0.00") & " с."
    Exit Function

Failed:
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = ResultOutputErrorDescription(Err.Description)
    If Not excelGuard Is Nothing Then excelGuard.Restore ' Интерфейс возвращаем до записи отчёта: недоступный каталог отчёта не должен оставлять события и экран выключенными или подменять ошибку.
    On Error Resume Next
    If Not report Is Nothing Then
        report.AddError "Расчет остановлен", errorDescription
        report.Save "Расчет не выполнен: " & errorDescription
    End If
    On Error GoTo 0
    Err.Raise errorNumber, errorSource, errorDescription
End Function


' До расчёта проверяет, хватает ли места таблицам Results. Учитывает заголовки,
' ширину таблиц, блоки сохранённой геометрии и состояний, а также две пустые
' строки между крупными блоками. При нехватке места расчёт не запускается.
Private Sub ValidateResultsOutputLayout(ByVal workbook As Object, ByVal section As CSectionModel, _
        ByVal batch As CBatchSectionCalculator, ByVal profiles As CCalculationProfileCatalog)
    If workbook Is Nothing Then Err.Raise vbObjectError + 4160, "ValidateResultsOutputLayout", "Книга Excel не передана."
    If section Is Nothing Then Err.Raise vbObjectError + 4161, "ValidateResultsOutputLayout", "Модель сечения не передана."
    If batch Is Nothing Then Err.Raise vbObjectError + 4162, "ValidateResultsOutputLayout", "Пакетный расчетчик не передан."

    Const SNAPSHOT_GAP_COLUMNS As Long = 2

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

    Dim outputCombinationRows As Long
    outputCombinationRows = summaryWriter.RequiredDataRowsForWorkbook(workbook, batch.Count)

    Dim summaryRows As Long
    Dim summaryCols As Long
    summaryRows = summaryWriter.RequiredSummaryOutputRowsForWorkbook(workbook, batch.Count)
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
        AddRowsGapIssue issues, "rngStrengthSummaryAnchor", summaryBottomRow, strengthTopRow, strengthTopRow - 1
    End If
    If strengthBottomRow + RESULTS_TABLE_GAP_ROWS >= crackTopRow Then
        AddRowsGapIssue issues, "rngCrackSummaryAnchor", strengthBottomRow, crackTopRow, crackTopRow - 1
    End If
    If crackBottomRow + RESULTS_TABLE_GAP_ROWS >= stabilityTopRow Then
        AddRowsGapIssue issues, "rngStabilitySummaryAnchor", crackBottomRow, stabilityTopRow, stabilityTopRow - 1
    End If
    If stabilityBottomRow + RESULTS_TABLE_GAP_ROWS >= snapshotTitleRow Then
        AddRowsGapIssue issues, "rngNDMElementResults и расположенных ниже таблиц Results", stabilityBottomRow, snapshotTitleRow, snapshotTitleRow
    End If

    Dim estimatedStateCount As Long
    Dim elementCount As Long
    estimatedStateCount = EstimatedNamedStateCountForOutput(batch, profiles)
    elementCount = section.ConcreteCount + section.RebarCount
    ValidateSnapshotOutputLayout issues, workbook, ws, ndmWriter, elementCount, _
        estimatedStateCount, batch.Count, section.AnnotationCount, section.Contours.Count, SNAPSHOT_GAP_COLUMNS

    If issues.Count > 0 Then
        Err.Raise vbObjectError + 4163, "ValidateResultsOutputLayout", _
            ResultsOutputLayoutMessage(outputCombinationRows, issues)
    End If
    Exit Sub

MissingAnchor:
    Err.Raise vbObjectError + 4164, "ValidateResultsOutputLayout", _
        MissingResultsAnchorMessage(workbook)
End Sub

' Перечисляет отсутствующие именованные ячейки начала таблиц Results,
' без которых нельзя заранее проверить размещение блоков вывода.
Private Function MissingResultsAnchorMessage(ByVal workbook As Object) As String
    Dim requiredNames As Variant
    requiredNames = Array( _
        "rngBatchSummary", _
        "rngStrengthSummaryAnchor", _
        "rngCrackSummaryAnchor", _
        "rngStabilitySummaryAnchor", _
        "rngNDMElementResults")

    Dim missing As String
    Dim index As Long
    For index = LBound(requiredNames) To UBound(requiredNames)
        If Not WorkbookNameHasRange(workbook, CStr(requiredNames(index))) Then
            If Len(missing) > 0 Then missing = missing & ", "
            missing = missing & CStr(requiredNames(index))
        End If
    Next index

    If Len(missing) = 0 Then
        MissingResultsAnchorMessage = _
            "На листе Results не удалось прочитать один из обязательных якорей вывода. " & _
            "Проверьте именованные диапазоны rngBatchSummary, rngStrengthSummaryAnchor, " & _
            "rngCrackSummaryAnchor, rngStabilitySummaryAnchor и rngNDMElementResults."
    ElseIf InStr(1, missing, ",", vbBinaryCompare) > 0 Then
        MissingResultsAnchorMessage = _
            "На листе Results не найдены обязательные якоря вывода: " & missing & ". " & _
            "Добавьте эти именованные диапазоны на лист Results."
    Else
        MissingResultsAnchorMessage = _
            "На листе Results не найден обязательный якорь вывода: " & missing & ". " & _
            "Добавьте именованный диапазон " & missing & " на лист Results."
    End If
End Function

' Проверяет, что имя существует в книге и действительно ссылается на диапазон.
Private Function WorkbookNameHasRange(ByVal workbook As Object, ByVal rangeName As String) As Boolean
    On Error GoTo MissingName
    Dim anchor As Object
    Set anchor = workbook.Names.Item(rangeName).RefersToRange
    WorkbookNameHasRange = Not anchor Is Nothing
    Exit Function

MissingName:
    WorkbookNameHasRange = False
End Function

' Проверяет расстояния между соседними нижними таблицами Results и их ширину.
Private Sub ValidateSnapshotOutputLayout(ByVal issues As Collection, ByVal workbook As Object, _
        ByVal ws As Object, ByVal ndmWriter As CNDMResultsWriter, ByVal elementCount As Long, _
        ByVal estimatedStateCount As Long, ByVal combinationCount As Long, _
        ByVal annotationCount As Long, ByVal contourCount As Long, ByVal requiredGapColumns As Long)
    Dim names(1 To 6) As String
    Dim rows(1 To 6) As Double
    Dim cols(1 To 6) As Long
    names(1) = "rngNDMElementResults"
    names(2) = "rngNDMSectionGeometry"
    names(3) = "rngNDMSectionContours"
    names(4) = "rngNDMSectionProperties"
    names(5) = "rngNDMMaterialDiagrams"
    names(6) = "rngNDMSectionAnnotations"

    rows(1) = ndmWriter.EstimatedElementResultRows(elementCount, estimatedStateCount)
    Dim store As CCalculationSnapshotStore: Set store = New CCalculationSnapshotStore
    Dim fileMode As Boolean: fileMode = False
    If combinationCount > 0 Then fileMode = store.StorageMode(workbook) = "File"
    If fileMode Then rows(1) = 1 ' Таблица состояний на лист не выводится; остальные технические таблицы остаются обязательными.
    rows(2) = 1 + elementCount
    rows(3) = 1 + contourCount
    rows(4) = ndmWriter.EstimatedSectionPropertyRows(combinationCount, estimatedStateCount)
    rows(5) = 1 + MaxLong(1, estimatedStateCount) * 14
    rows(6) = 1 + MaxLong(2, annotationCount)

    Dim index As Long
    Dim anchors(1 To 6) As Object
    For index = 1 To 6
        cols(index) = ndmWriter.OutputColumnCount(names(index))

        Dim anchor As Object
        On Error GoTo MissingSnapshotAnchor
        Set anchor = workbook.Names.Item(names(index)).RefersToRange
        On Error GoTo 0

        If Not (anchor.Worksheet Is ws) Then
            AddLayoutIssue issues, names(index) & " должен находиться на том же листе Results, что и остальные блоки вывода."
        End If
        CheckFootprintWithinSheet issues, ws, names(index), anchor.Row, anchor.Column, rows(index), cols(index)
        If index = 1 And Not fileMode Then
            If Not ResultRowsFitSheet(anchor.Row, rows(index), ws.Rows.Count) Then AddLayoutIssue issues, ExternalSnapshotStorageAdvice(workbook)
        End If

        Set anchors(index) = anchor
        Dim other As Long, gapColumns As Long
        For other = 1 To index - 1
            If anchor.Column >= anchors(other).Column Then
                gapColumns = anchor.Column - anchors(other).Column - cols(other)
            Else
                gapColumns = anchors(other).Column - anchor.Column - cols(index)
            End If
            If gapColumns < requiredGapColumns Then AddLayoutIssue issues, _
                "Между " & names(other) & " и " & names(index) & " требуется не менее " & CStr(requiredGapColumns) & " пустых столбцов."
        Next other
    Next index
    Exit Sub

MissingSnapshotAnchor:
    AddLayoutIssue issues, "На листе Results не найден якорь " & names(index) & "."
    On Error GoTo 0
End Sub

' Проверяет, что прямоугольник вывода не выходит за пределы листа Excel.
Private Sub CheckFootprintWithinSheet(ByVal issues As Collection, ByVal ws As Object, _
        ByVal blockName As String, ByVal firstRow As Long, ByVal firstColumn As Long, _
        ByVal rowCount As Double, ByVal columnCount As Long)
    If rowCount < 1 Then rowCount = 1
    If columnCount < 1 Then columnCount = 1
    If firstRow < 1 Or firstColumn < 1 Then
        AddLayoutIssue issues, blockName & " имеет некорректный левый верхний угол вывода."
        Exit Sub
    End If

    If Not ResultRowsFitSheet(firstRow, rowCount, ws.Rows.Count) Then AddLayoutIssue issues, ResultRowLimitDescription(blockName, firstRow, rowCount, ws.Rows.Count)
    If firstColumn + columnCount - 1 > ws.Columns.Count Then
        AddLayoutIssue issues, blockName & " требует " & CStr(columnCount) & _
            " столбцов от столбца " & CStr(firstColumn) & ", но справа не хватает столбцов листа."
    End If
End Sub

' Оценивает максимальное число сохраняемых расчётных состояний для таблицы результатов
' элементов.
Private Function EstimatedNamedStateCountForOutput(ByVal batch As CBatchSectionCalculator, _
        ByVal profiles As CCalculationProfileCatalog) As Long
    If batch Is Nothing Then Exit Function

    Dim i As Long
    For i = 1 To batch.Count
        EstimatedNamedStateCountForOutput = EstimatedNamedStateCountForOutput + _
            EstimatedNamedStateCountForProfile(batch.InputProfileId(i), profiles)
    Next i
End Function

' Оценивает максимальное число состояний для профиля. Расчёт трещин может дать
' НДС с трещинами и две точки образования трещины, даже если конкретному сочетанию
' не понадобятся все три состояния.
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

' Добавляет конкретную проблему размещения Results в общий список исправлений.
' Отсутствующий список допускается у служебного вызова и не создает новую коллекцию.
Private Sub AddLayoutIssue(ByVal issues As Collection, ByVal text As String)
    If issues Is Nothing Then Exit Sub
    issues.Add text
End Sub

' Добавляет рекомендацию о числе строк между соседними таблицами Results.
' rowsToInsert — недостающее число строк до начала следующей таблицы;
' insertBeforeRow — строка синего заголовка, перед которой следует вставить строки.
Private Sub AddRowsGapIssue(ByVal issues As Collection, ByVal rangeName As String, _
        ByVal previousBottomRow As Long, ByVal nextTopRow As Long, ByVal insertBeforeRow As Long)
    Dim rowsToInsert As Long
    rowsToInsert = previousBottomRow + RESULTS_TABLE_GAP_ROWS + 1 - nextTopRow
    If rowsToInsert < 1 Then rowsToInsert = 1
    If insertBeforeRow < 1 Then insertBeforeRow = nextTopRow

    AddLayoutIssue issues, "На листе Results вставьте " & CStr(rowsToInsert) & " " & RowsWord(rowsToInsert) & _
        " над строкой " & CStr(insertBeforeRow) & ", чтобы опустить диапазон " & rangeName & "."
End Sub

' Возвращает короткое русское склонение слова "строка" для пользовательской подсказки.
Private Function RowsWord(ByVal count As Long) As String
    Dim lastTwo As Long
    lastTwo = Abs(count) Mod 100
    If lastTwo >= 11 And lastTwo <= 14 Then
        RowsWord = "строк"
        Exit Function
    End If

    Select Case Abs(count) Mod 10
        Case 1
            RowsWord = "строку"
        Case 2, 3, 4
            RowsWord = "строки"
        Case Else
            RowsWord = "строк"
    End Select
End Function

' Объясняет отказ до расчета, когда блокам Results не хватает строк.
' Включает уже подготовленные адресные рекомендации, не теряя число сочетаний.
Private Function ResultsOutputLayoutMessage(ByVal combinationCount As Long, ByVal issues As Collection) As String
    ResultsOutputLayoutMessage = "Все сочетания (" & CStr(combinationCount) & _
        ") не помещаются в текущую раскладку или пределы листа Results. Расчет не запущен." & _
        vbCrLf & "Что сделать:" & _
        vbCrLf & NumberedCollectionLines(issues) & _
        vbCrLf & CStr(issues.Count + 1) & ") Либо на листе Config уменьшите число строк в rngLoadCombinations."
End Function

' Собирает список действий компактным нумерованным перечнем для итогового окна.
Private Function NumberedCollectionLines(ByVal items As Collection) As String
    Dim lines As String
    Dim i As Long
    For i = 1 To items.Count
        If Len(lines) > 0 Then lines = lines & vbCrLf
        lines = lines & CStr(i) & ") " & CStr(items.Item(i))
    Next i
    NumberedCollectionLines = lines
End Function

' Выбирает большее число при резервировании строк под сохранённые состояния и аннотации.
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

' Готовит итоговое сообщение о запуске, отдельно выделяя ошибки входных строк.
' Берёт инженерные статусы из готовых результатов расчёта сочетаний.
Private Function BuildCalculationMessage(ByVal batch As CBatchSectionCalculator) As String
    If batch.InvalidInputCount > 0 Then
        BuildCalculationMessage = CalculationCompletionMessage(batch.Count, batch.InvalidInputCount, _
            batch.FirstInvalidInputMessage, vbNullString)
    Else
        BuildCalculationMessage = CalculationCompletionMessage(batch.Count, 0, vbNullString, batch.WorstCombinationID)
    End If
End Function

' Формирует список всех сочетаний, принятых для расчёта.
' В отчёте усилия показываются в пользовательских выходных единицах
' и с пользовательскими знаками, как на листе Results.
Private Function CombinationListForReport(ByVal batch As CBatchSectionCalculator, ByVal units As CUnitSystem) As String
    If batch Is Nothing Then Exit Function

    Dim lines As Collection
    Set lines = New Collection

    Dim i As Long
    For i = 1 To batch.Count
        lines.Add CStr(i) & ". " & batch.CombinationID(i) & _
            "; profile=" & batch.InputProfileId(i) & _
            "; N=" & ReportForceText(batch.N(i), units) & _
            "; Mx=" & ReportMxText(batch.Mx(i), units) & _
            "; My=" & ReportMyText(batch.My(i), units) & _
            "; comment=" & batch.CombinationName(i)
    Next i

    CombinationListForReport = JoinCollectionLines(lines)
End Function

' Форматирует длину для отчёта выполнения в тех же единицах, что и лист Results.
Private Function ReportLengthText(ByVal value As Double, ByVal units As CUnitSystem) As String
    If units Is Nothing Then
        ReportLengthText = FormatReportNumber(value) & " мм"
    Else
        ReportLengthText = FormatReportNumber(units.InternalLengthToOutput(value)) & " " & units.OutputLengthUnit
    End If
End Function

' Форматирует продольную силу с учетом выбранной пользователем системы знаков.
Private Function ReportForceText(ByVal value As Double, ByVal units As CUnitSystem) As String
    If units Is Nothing Then
        ReportForceText = FormatReportNumber(value) & " Н"
    Else
        ReportForceText = FormatReportNumber(units.InternalForceToOutput(value)) & " " & units.OutputForceUnit
    End If
End Function

' Форматирует Mx с тем же знаком и единицей, что в подробном выводе Results.
Private Function ReportMxText(ByVal value As Double, ByVal units As CUnitSystem) As String
    If units Is Nothing Then
        ReportMxText = FormatReportNumber(value) & " Н*мм"
    Else
        ReportMxText = FormatReportNumber(units.InternalMomentMxToOutput(value)) & " " & units.OutputMomentUnit
    End If
End Function

' Форматирует My с тем же знаком и единицей, что в подробном выводе Results.
Private Function ReportMyText(ByVal value As Double, ByVal units As CUnitSystem) As String
    If units Is Nothing Then
        ReportMyText = FormatReportNumber(value) & " Н*мм"
    Else
        ReportMyText = FormatReportNumber(units.InternalMomentMyToOutput(value)) & " " & units.OutputMomentUnit
    End If
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

' Задает пользовательскую точку относительно центра бетонного сечения.
' Смещения переводятся через CUnitSystem; приведенный центр здесь не используется.
Private Sub ApplyLoadReferenceFromSettings(ByVal section As CSectionModel, _
        ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem, _
        ByVal batch As CBatchSectionCalculator)
    Dim referenceX As Double
    Dim referenceY As Double
    Dim offsetX As Double, offsetY As Double
    Dim key As String ' Текущая координата для адресной диагностики переполнения.
    CalculateConcreteSectionCentroid section, referenceX, referenceY

    key = "Load.ReferenceOffsetX"
    offsetX = settings.GetRequiredDouble(key)
    On Error GoTo InvalidConversion
    offsetX = units.InputLengthToInternal(offsetX)
    offsetX = referenceX + offsetX
    On Error GoTo 0
    key = "Load.ReferenceOffsetY"
    offsetY = settings.GetRequiredDouble(key)
    On Error GoTo InvalidConversion
    offsetY = units.InputLengthToInternal(offsetY)
    offsetY = referenceY + offsetY
    On Error GoTo 0
    batch.ApplyLoadReference offsetX, offsetY, referenceX, referenceY
    Exit Sub
InvalidConversion:
    If Err.Number <> 6 Then Err.Raise Err.Number, Err.Source, Err.Description
    Err.Raise vbObjectError + 4139, "ApplyLoadReferenceFromSettings", _
        settings.InputErrorMessage(key, "Смещение слишком велико для получения координаты точки нагрузки в мм.", _
            "Уменьшите модуль числа и проверьте INPUT-единицу длины.")
End Sub

' Читает дополнительную таблицу длительных нагрузок для расчёта устойчивости
' и сразу преобразует усилия во внутренние единицы. Обрабатывает только нужные
' сочетания; ошибку строки сохраняет для результата соответствующего сочетания.
Public Sub LoadStabilityDurationLoadsFromWorkbook(ByVal workbook As Object, _
        ByVal batch As CBatchSectionCalculator, ByVal units As CUnitSystem)
    If workbook Is Nothing Then Err.Raise vbObjectError + 4113, _
        "LoadStabilityDurationLoadsFromWorkbook", "Не передана книга для чтения дополнительных нагрузок устойчивости."
    If batch Is Nothing Then Err.Raise vbObjectError + 4113, _
        "LoadStabilityDurationLoadsFromWorkbook", "Не передан расчетный пакет для дополнительных нагрузок устойчивости."

    batch.ClearStabilityDurationLoads
    Dim activeIDs As Object, index As Long
    Set activeIDs = CreateObject("Scripting.Dictionary")
    activeIDs.CompareMode = vbTextCompare
    For index = 1 To batch.Count
        If batch.StabilityRequested(index) Then activeIDs(batch.CombinationID(index)) = False
    Next index
    If activeIDs.Count = 0 Then Exit Sub
    If units Is Nothing Then Err.Raise vbObjectError + 4113, _
        "LoadStabilityDurationLoadsFromWorkbook", "Не передана система единиц для дополнительных нагрузок устойчивости."

    On Error GoTo MissingRange
    Dim range As Object
    Set range = workbook.Names.Item("rngStabilityDurationLoads").RefersToRange
    On Error GoTo 0
    If range.Rows.Count < 2 Or range.Columns.Count < 4 Then
        RecordDurationTableError batch, activeIDs, _
            "Таблица дополнительных нагрузок устойчивости rngStabilityDurationLoads (" & _
            range.Worksheet.Name & "!" & range.Address(False, False) & _
            ") должна содержать заголовок, хотя бы одну строку и четыре столбца: Combination ID, N, Mx, My. " & _
            "Восстановите границы именованного диапазона в диспетчере имен Excel."
        Exit Sub
    End If
    Dim values As Variant
    values = range.Value2

    Dim rowIndex As Long
    For rowIndex = 2 To UBound(values, 1)
        If IsError(values(rowIndex, 1)) Or IsNull(values(rowIndex, 1)) Then
            RecordDurationTableError batch, activeIDs, _
                "В таблице rngStabilityDurationLoads ссылка на сочетание (" & _
                range.Worksheet.Name & "!" & range.Cells(rowIndex, 1).Address(False, False) & _
                ") содержит ошибку. Исправьте формулу или Combination ID; без этой ссылки нельзя сопоставить дополнительные нагрузки."
            Exit Sub
        End If
        Dim combinationID As String
        combinationID = Trim$(CStr(values(rowIndex, 1)))
        If activeIDs.Exists(combinationID) Then
            activeIDs(combinationID) = True
            Dim component As Long, loads(2 To 4) As Double, rowError As String, componentError As String
            rowError = vbNullString
            For component = 2 To 4
                componentError = vbNullString
                If Not TryReadDurationLoad(values(rowIndex, component), component, units, loads(component), componentError) Then
                    If Len(rowError) > 0 Then rowError = rowError & "; "
                    rowError = rowError & DurationLoadComponentName(component) & " (" & _
                        range.Worksheet.Name & "!" & range.Cells(rowIndex, component).Address(False, False) & "): " & componentError
                End If
            Next component
            If Len(rowError) > 0 Then
                batch.AddInvalidStabilityDurationLoad combinationID, _
                    "Дополнительные нагрузки устойчивости для сочетания " & combinationID & _
                    " в таблице rngStabilityDurationLoads: " & rowError & _
                    " Исправьте указанные ячейки; пустое значение допускается и означает ноль."
            Else
                batch.AddStabilityDurationLoad combinationID, loads(2), loads(3), loads(4)
            End If
        End If
    Next rowIndex
    Dim id As Variant
    For Each id In activeIDs.Keys ' Пустые усилия допустимы только в существующей строке нужного сочетания.
        If Not activeIDs(id) Then
            batch.AddInvalidStabilityDurationLoad CStr(id), _
                "Для сочетания " & CStr(id) & " в таблице «Нагрузки для устойчивости» (rngStabilityDurationLoads, " & _
                range.Worksheet.Name & "!" & range.Address(False, False) & ") не найдена строка с этим номером. " & _
                "Добавьте строку или исправьте номер сочетания. Если дополнительной нагрузки нет, оставьте N, Mx и My пустыми либо укажите 0."
        End If
    Next id
    Exit Sub

MissingRange:
    Dim errorNumber As Long, description As String
    errorNumber = Err.Number: description = Err.Description
    If errorNumber <> 1004 Then Err.Raise errorNumber, "LoadStabilityDurationLoadsFromWorkbook", description
    RecordDurationTableError batch, activeIDs, _
        "Для запрошенной проверки устойчивости не удалось прочитать таблицу дополнительных нагрузок rngStabilityDurationLoads. " & _
        "Восстановите ссылку этого имени в диспетчере имен Excel. Отсутствующая таблица не заменяется нулевыми нагрузками."
End Sub

' Читает одно числовое значение и выполняет централизованный перевод знаков
' и единиц. Пустота допустима; Null, ошибка формулы, текст и переполнение
' получают причину ввода. Неожиданная программная ошибка не маскируется.
Private Function TryReadDurationLoad(ByVal value As Variant, ByVal component As Long, _
        ByVal units As CUnitSystem, ByRef internalValue As Double, ByRef reason As String) As Boolean
    internalValue = 0#: reason = vbNullString
    If IsError(value) Then reason = "ячейка содержит ошибку формулы Excel.": Exit Function
    If IsNull(value) Then reason = "значение не определено.": Exit Function
    If IsEmpty(value) Then TryReadDurationLoad = True: Exit Function
    If VarType(value) = vbString Then
        If Len(Trim$(CStr(value))) = 0 Then TryReadDurationLoad = True: Exit Function
    End If
    If Not IsNumeric(value) Then reason = "вместо числа задан текст или другой недопустимый тип значения.": Exit Function
    On Error GoTo InvalidNumber
    Select Case component
        Case 2: internalValue = units.InputForceToInternal(CDbl(value))
        Case 3: internalValue = units.InputMomentMxToInternal(CDbl(value))
        Case 4: internalValue = units.InputMomentMyToInternal(CDbl(value))
        Case Else: Err.Raise vbObjectError + 4113, "TryReadDurationLoad", "Неизвестная компонента дополнительной нагрузки устойчивости."
    End Select
    TryReadDurationLoad = True
    Exit Function
InvalidNumber:
    Dim errorNumber As Long, description As String
    errorNumber = Err.Number: description = Err.Description
    If errorNumber <> 6 And errorNumber <> 13 Then Err.Raise errorNumber, "TryReadDurationLoad", description
    reason = "число невозможно прочитать или перевести в выбранные единицы INPUT без переполнения."
End Function

' Возвращает название составляющей нагрузки в дополнительной таблице из четырёх столбцов.
' Не определяет адрес ячейки или пользовательские единицы.
Private Function DurationLoadComponentName(ByVal component As Long) As String
    Select Case component
        Case 2: DurationLoadComponentName = "N"
        Case 3: DurationLoadComponentName = "Mx"
        Case 4: DurationLoadComponentName = "My"
    End Select
End Function

' Сохраняет структурную ошибку таблицы для всех реально запрошенных проверок
' устойчивости. Ранее прочитанные строки сбрасываются: частичный ввод при
' поврежденном диапазоне не должен выдаваться за корректную нулевую нагрузку.
Private Sub RecordDurationTableError(ByVal batch As CBatchSectionCalculator, ByVal activeIDs As Object, ByVal reason As String)
    batch.ClearStabilityDurationLoads
    Dim id As Variant
    For Each id In activeIDs.Keys: batch.AddInvalidStabilityDurationLoad CStr(id), reason: Next id
End Sub

' Возвращает таблицу 7.21 СП 35 как массив и причину неверного ввода.
' Сообщение содержит адрес именованного диапазона таблицы. CStabilityCalculator
' проверяет применимость её значений. Ошибка относится только к той главной
' плоскости, для которой выбран табличный метод СП 35.
Public Function ReadSP35Table721FromWorkbook(ByVal workbook As Object, ByRef inputError As String) As Variant
    inputError = vbNullString
    If workbook Is Nothing Then Err.Raise vbObjectError + 4135, _
        "ReadSP35Table721FromWorkbook", "Для чтения таблицы СП 35 не передана книга."
    On Error GoTo MissingRange
    Dim range As Object
    Set range = workbook.Names.Item("rngSP35Table721").RefersToRange
    On Error GoTo 0
    If range.Areas.Count <> 1 Then
        inputError = "Таблица 7.21 СП 35: rngSP35Table721 должен ссылаться на один прямоугольный диапазон. Исправьте имя диапазона."
        Exit Function
    End If

    Dim data As Variant, validator As CStabilityCalculator
    Dim row As Long, column As Long, reason As String, address As String
    data = range.Value2
    Set validator = New CStabilityCalculator
    If Not validator.TryValidateSP35Table721(data, row, column, reason) Then
        address = range.Worksheet.Name & "!"
        If row > 0 And column > 0 Then
            address = address & range.Cells(row, column).Address(False, False)
        Else
            address = address & range.Address(False, False)
        End If
        inputError = "Таблица 7.21 СП 35, " & address & ": " & reason
    End If
    ReadSP35Table721FromWorkbook = data
    Exit Function

MissingRange:
    inputError = "Не найден диапазон rngSP35Table721 с таблицей 7.21 СП 35. " & _
        "Восстановите именованный диапазон таблицы на Config; резервные коэффициенты не используются."
End Function

' Возвращает центр тяжести бетонной части сечения.
' Это базовая точка пользовательских нагрузок: Load.ReferenceOffsetX/Y
' откладываются именно от центра бетона, а не от приведенного сечения
' бетон + арматура.
Public Sub CalculateConcreteSectionCentroid(ByVal section As CSectionModel, _
        ByRef referenceX As Double, ByRef referenceY As Double)
    If section Is Nothing Then Err.Raise vbObjectError + 2200, "CSectionPropertiesCalculator", "Модель сечения не передана."
    If section.ConcreteCount <= 0 Then Err.Raise vbObjectError + 2201, "CSectionPropertiesCalculator", "В модели сечения нет бетонных элементов."
    Dim area As Double
    section.ConcreteMeshCentroid referenceX, referenceY, area
    If area <= GEOM_TOLERANCE Then Err.Raise vbObjectError + 2202, "CSectionPropertiesCalculator", "Площадь бетонного сечения равна нулю."
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

' Очищает сводки, подробные результаты и сохранённые таблицы через классы записи.
' Настройки и исходные данные не меняет; расчёт и импорт не запускает.
Public Sub ClearSectionResultsForWorkbook(ByVal workbook As Object)
    If workbook Is Nothing Then Err.Raise vbObjectError + 4110, "ClearSectionResultsForWorkbook", "Книга Excel не передана."
    Dim summaryWriter As CBatchResultWriter
    Set summaryWriter = New CBatchResultWriter
    summaryWriter.ClearSummary workbook
    Dim ndmWriter As CNDMResultsWriter
    Set ndmWriter = New CNDMResultsWriter
    ndmWriter.ClearResults workbook
End Sub

' Читает зарегистрированные таблицы выбранного рецепта и создаёт единое описание сечения.
' Готовое описание не содержит ссылок на объекты Excel.
Public Function ReadWorkbookSectionDefinition(ByVal workbook As Object, ByVal settings As CSystemSettingsReader, Optional ByVal units As CUnitSystem = Nothing) As CSectionDefinition
    If workbook Is Nothing Or settings Is Nothing Then SectionRecipeError "ReadWorkbookSectionDefinition", "Книга и настройки Config не переданы."
    If units Is Nothing Then
        Set units = New CUnitSystem: units.LoadFromSettings settings
    End If
    Dim registry As CSectionTypeRegistry, reader As CSectionInputSnapshotReader
    Dim context As CSectionRecipeContext, typeID As String
    Set registry = New CSectionTypeRegistry: typeID = registry.SelectedRecipeType(settings)
    Set reader = New CSectionInputSnapshotReader
    Set context = reader.Capture(workbook, registry.RecipeTableNames(typeID), units)
    Set ReadWorkbookSectionDefinition = registry.CreateDefinition(typeID, context)
End Function

' На этапе формирования строит параметрическую модель по таблице выбранного
' рецепта на листе «Сечения». Ветка AutoCAD читает только ранее сохранённую
' модель и не импортирует чертёж. Штатная кнопка расчёта читает Results напрямую.
Public Function BuildWorkbookSectionModel(ByVal workbook As Object, ByVal settings As CSystemSettingsReader, _
        Optional ByVal units As CUnitSystem = Nothing, Optional ByRef generatedBuilder As CSectionBuilder = Nothing) As CSectionModel
    If workbook Is Nothing Then Err.Raise vbObjectError + 4130, "BuildWorkbookSectionModel", "Книга Excel не передана."
    If settings Is Nothing Then Err.Raise vbObjectError + 4131, "BuildWorkbookSectionModel", "Настройки Config не переданы."
    If units Is Nothing Then
        Set units = New CUnitSystem: units.LoadFromSettings settings
    End If

    Dim geometrySource As String
    geometrySource = settings.GetRequiredChoice("Geometry.Source", Array("Generated", "AutoCAD"))

    If StrComp(geometrySource, "Generated", vbTextCompare) = 0 Then
        Dim definition As CSectionDefinition, mesh As TSectionMeshOptions
        Set definition = ReadWorkbookSectionDefinition(workbook, settings, units)
        mesh = ReadSectionMeshOptions(settings, units)
        If generatedBuilder Is Nothing Then Set generatedBuilder = New CSectionBuilder
        Set BuildWorkbookSectionModel = generatedBuilder.Build(definition, mesh)
    ElseIf StrComp(geometrySource, "AutoCAD", vbTextCompare) = 0 Then
        Set generatedBuilder = Nothing
        If StrComp(ResultsGeometrySource(workbook), "AutoCADImport", vbTextCompare) <> 0 Then
            Err.Raise vbObjectError + 4134, "BuildWorkbookSectionModel", _
                "Выбрано Geometry.Source = AutoCAD, но на листе Results нет предварительно импортированной AutoCAD-геометрии. " & _
                "Сначала нажмите кнопку ""Сформировать сечение""."
        End If
        Set BuildWorkbookSectionModel = ReadSectionGeometryFromResults(workbook, "AutoCADImport")
    Else
        Err.Raise vbObjectError + 4133, "BuildWorkbookSectionModel", _
            "Geometry.Source должен быть Generated или AutoCAD."
    End If
End Function

' ДЛЯ СБОРКИ: тот же реестр является источником штатного списка Geometry.Type.
Public Function RegisteredSectionTypes() As Variant
    Dim registry As CSectionTypeRegistry: Set registry = New CSectionTypeRegistry
    RegisteredSectionTypes = registry.RegisteredRecipeTypes
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

' Подготавливает характеристики сечения и аннотации до записи результатов.
' Геометрические характеристики вычисляет CSectionPropertiesCalculator,
' а размеры и подписи импортированной модели — CSectionAnnotationBuilder.
' Текущие Config и Results не читаются; решение НДС не выполняется.
Public Function PrepareSectionSnapshot(ByVal section As CSectionModel, _
        Optional ByVal provider As CMaterialModelProvider = Nothing) As CSectionPropertiesCalculator
    Dim snapshot As CSectionPropertiesCalculator
    Set snapshot = New CSectionPropertiesCalculator
    snapshot.PrepareSnapshot section, provider
    Dim annotations As CSectionAnnotationBuilder: Set annotations = New CSectionAnnotationBuilder
    annotations.PrepareImportedSection section, snapshot.SnapshotBounds
    Set PrepareSectionSnapshot = snapshot
End Function

' ==========================================================================
' ДЛЯ ТЕСТОВ
' ==========================================================================

' -------------------- ДЛЯ ТЕСТОВ --------------------
' перенос пользовательской точки нагрузки
' Вызывает обычное чтение Config и перенос нагрузки. Проверяемые преобразования
' единиц, центр сечения и формула переноса не повторяются в тесте.
Public Sub Audit03ApplyLoadReferenceForTests(ByVal section As CSectionModel, _
        ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem, _
        ByVal batch As CBatchSectionCalculator)
    ApplyLoadReferenceFromSettings section, settings, units, batch
End Sub
