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
    Else
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
    MsgBox "Результаты и диагностика очищены. Исходные данные не изменены.", vbInformation, "RC Section NDM"
    Exit Sub

Failed:
    MsgBox "Не удалось очистить результаты: " & Err.Description, vbExclamation, "RC Section NDM"
End Sub

Public Sub UpdateSectionPlot()
    On Error GoTo Failed
    UpdateSectionPlotForWorkbook ThisWorkbook
    MsgBox "Схема сечения обновлена по последнему расчетному снимку Results.", vbInformation, "RC Section NDM"
    Exit Sub

Failed:
    MsgBox "Схема не обновлена: " & Err.Description, vbExclamation, "RC Section NDM"
End Sub

Public Sub ImportGeometryFromAutoCAD()
    On Error GoTo Failed
    Dim message As String
    message = ImportGeometryFromAutoCADForWorkbook(ThisWorkbook)
    MsgBox message, vbInformation, "RC Section NDM"
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
    If Not TryLoadFullPlotReader(reader, workbook, settings) Then
        If Not CanUseGeometryPreview(workbook, settings) Then
            ClearSectionPlotForNoData workbook, _
                "Схема не обновлена: нет расчетных данных для текущего Geometry.Source."
            If raiseIfNoData Then
                Err.Raise vbObjectError + 4144, "UpdateSectionPlotForWorkbook", _
                    "На листе Results нет расчетных данных для текущей схемы. Выполните расчет."
            End If
            Exit Sub
        End If

        Set reader = New CSectionPlotDataReader
        reader.LoadGeometryPreviewFromWorkbook workbook, settings
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
        ByVal workbook As Object, ByVal settings As CSystemSettingsReader) As Boolean
    On Error GoTo Failed
    reader.LoadFromWorkbook workbook, settings
    TryLoadFullPlotReader = True
    Exit Function

Failed:
    If Err.Number = vbObjectError + 4702 Or Err.Number = vbObjectError + 4705 Then
        TryLoadFullPlotReader = False
    Else
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Function

' Проверяет, можно ли вместо полноценной расчетной схемы показать preview
' импортированной AutoCAD-геометрии. Generated-сценарий сюда не допускается:
' иначе старый AutoCAD-preview может остаться на листе после переключения
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
' не должен видеть старую подпись или старую AutoCAD-preview схему как будто она
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

    Dim writer As CNDMResultsWriter
    Set writer = New CNDMResultsWriter
    writer.WriteGeometryPreview workbook, section, units

    UpdateSectionGeometryPreviewForWorkbook workbook

    ImportGeometryFromAutoCADForWorkbook = "Геометрия успешно импортирована из AutoCAD." & vbCrLf & _
        "Бетонных Region: " & CStr(section.ConcreteCount) & "; арматурных Region: " & _
        CStr(section.RebarCount) & "." & vbCrLf & _
        "На схеме показаны только импортированные элементы для визуального контроля."
End Function

' Запускает связанный набор операций и возвращает пользователю итоговый статус выполнения.
Public Function RunSectionCalculationForWorkbook(ByVal workbook As Object, Optional ByVal showMessages As Boolean = False) As String
    On Error GoTo Failed
    If workbook Is Nothing Then Err.Raise vbObjectError + 4100, "RunSectionCalculationForWorkbook", "Книга Excel не передана."

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
    report.AddValue "Режим расчета", settings.GetRawString("Calculation.Mode", "-")
    report.AddValue "Расчет трещин", settings.GetRawString("SLS.Crack.Enabled", "-")
    
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
    report.AddValue "Mesh.Step", settings.GetRawString("Mesh.Step", "-")
    report.AddValue "Mesh.BoundarySubdivisions", settings.GetRawString("Mesh.BoundarySubdivisions", "-")
    Set section = BuildWorkbookSectionModel(workbook, settings, units)
    report.AddStep "CSectionModel построен."
    report.AddValue "Бетонных элементов", CStr(section.ConcreteCount)
    report.AddValue "Стержней арматуры", CStr(section.RebarCount)
    report.AddValue "Semantic-аннотаций", CStr(section.AnnotationCount)

    report.AddSection "Подготовка книги"
    report.AddStep "Начата очистка старых результатов."
    ClearSectionResultsForWorkbook workbook
    report.AddStep "Очищены rngResultSection, rngBatchSummary и таблицы расчетного снимка Results."

    Dim materialProvider As CMaterialModelProvider
    Set materialProvider = New CMaterialModelProvider
    report.AddStep "Начато построение material provider."
    materialProvider.Initialize settings, units
    report.AddStep "Построены материалные модели Strength, Mcrc и CrackedNDS по параметрам Config."

    Dim batch As CBatchSectionCalculator
    Set batch = New CBatchSectionCalculator
    report.AddStep "Создан CBatchSectionCalculator."
    batch.Initialize section, materialProvider
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

    report.AddSection "Точка приложения нагрузки"
    report.AddStep "Расчет центра тяжести бетонного сечения для пользовательской точки нагрузки."
    ApplyLoadReferenceFromSettings section, settings, units, batch
    report.AddValue "Точка приложения нагрузки X", FormatReportNumber(batch.LoadReferenceX) & " мм"
    report.AddValue "Точка приложения нагрузки Y", FormatReportNumber(batch.LoadReferenceY) & " мм"

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

    report.AddStep "Начата запись итогового блока определяющего сочетания на лист Расчет."
    WriteGoverningCombinationResults workbook, section, materialProvider, settings, units, batch
    report.AddStep "Итоговый блок определяющего сочетания записан на лист Расчет."

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
            report.AddStep "Схема очищена от старого изображения: нет доступных состояний LC."
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
    Err.Raise errorNumber, errorSource, errorDescription
End Function

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
            ". Определяющее сочетание: " & batch.GoverningCombinationID
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
            "; type=" & batch.CalculationType(i) & _
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
    ClearResultRange workbook.Names.Item("rngResultSection").RefersToRange
    Dim summaryWriter As CBatchResultWriter
    Set summaryWriter = New CBatchResultWriter
    summaryWriter.ClearSummary workbook
    Dim ndmWriter As CNDMResultsWriter
    Set ndmWriter = New CNDMResultsWriter
    ndmWriter.ClearResults workbook
End Sub

' Очищает накопленное состояние перед новым расчетом или повторным формированием вывода.
Private Sub ClearResultRange(ByVal target As Object)
    target.Offset(0, 0).Resize(target.Rows.Count, 1).ClearContents
    target.Offset(0, 4).Resize(target.Rows.Count, 1).ClearContents
    target.Offset(0, 8).Resize(target.Rows.Count, 1).ClearContents
    target.Offset(0, 12).Resize(target.Rows.Count, 1).ClearContents
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

Private Sub WriteGoverningCombinationResults(ByVal workbook As Object, ByVal section As CSectionModel, _
        ByVal materialProvider As CMaterialModelProvider, _
        ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem, ByVal batch As CBatchSectionCalculator)
    If batch Is Nothing Then Err.Raise vbObjectError + 4123, "WriteGoverningCombinationResults", "Результаты пакетного расчета отсутствуют."
    If batch.GoverningCombinationIndex <= 0 Then Exit Sub

    Dim index As Long
    index = batch.GoverningCombinationIndex
    If Not batch.StateAvailable(index) Then Exit Sub

    Dim currentPurpose As ECalculationPurpose
    currentPurpose = MaterialPurposeForCalculationType(batch.CalculationType(index))

    WriteCapacityAndCrackResults workbook, section, materialProvider, currentPurpose, settings, _
        units, batch.N(index), batch.UserMx(index), batch.UserMy(index), _
        batch.LoadReferenceX, batch.LoadReferenceY, batch.CapacityLoadPath(index)
End Sub

Private Sub WriteCapacityAndCrackResults(ByVal workbook As Object, ByVal section As CSectionModel, _
        ByVal materialProvider As CMaterialModelProvider, ByVal currentPurpose As ECalculationPurpose, _
        ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem, _
        ByVal nValue As Double, ByVal userMxValue As Double, ByVal userMyValue As Double, _
        ByVal referenceX As Double, ByVal referenceY As Double, _
        Optional ByVal capacityLoadPath As String = vbNullString)

    Dim writer As CCapacityResultWriter
    Set writer = New CCapacityResultWriter

    Dim concrete As Object
    Dim steel As Object
    Set concrete = materialProvider.ConcreteStateMaterial(currentPurpose)
    Set steel = materialProvider.SteelStateMaterial(currentPurpose)

    Dim internalMxValue As Double
    Dim internalMyValue As Double
    internalMxValue = userMxValue + nValue * referenceY
    internalMyValue = userMyValue + nValue * referenceX

    Dim concreteCentroidX As Double
    Dim concreteCentroidY As Double
    CalculateConcreteSectionCentroid section, concreteCentroidX, concreteCentroidY

    Dim centroidMxForCrack As Double
    Dim centroidMyForCrack As Double
    ' Проверка центрального растяжения в расчете трещин выполняется относительно
    ' центра тяжести бетонного сечения. Если пользователь задал чистое N, но
    ' сместил точку нагрузки через Load.ReferenceOffset, это уже внецентренное
    ' состояние и центральная ветка трещин не включается.
    centroidMxForCrack = userMxValue + nValue * (referenceY - concreteCentroidY)
    centroidMyForCrack = userMyValue + nValue * (referenceX - concreteCentroidX)

    Dim calculationMode As String
    calculationMode = settings.GetString("Calculation.Mode", "FullCapacity")

    Select Case LCase$(Trim$(calculationMode))
        Case "directstate"
            WriteDirectStateAndCrackResults workbook, section, materialProvider, currentPurpose, settings, units, _
                nValue, internalMxValue, internalMyValue, centroidMxForCrack, centroidMyForCrack, writer
            Exit Sub
        Case "fullcapacity"
        Case Else
            Err.Raise vbObjectError + 4124, "WriteCapacityAndCrackResults", _
                "Calculation.Mode должен быть DirectState или FullCapacity."
    End Select

    Dim capacity As CCapacitySolver
    Set capacity = New CCapacitySolver
    capacity.ApplySettings settings, units
    ApplyCapacityLimitsFromProvider capacity, materialProvider, cpStrength
    capacityLoadPath = NormalizedCapacityLoadPathForResult(capacityLoadPath, nValue, userMxValue, userMyValue)

    Dim nOffset As Double
    Dim nBase As Double
    Dim mxOffset As Double
    Dim mxBase As Double
    Dim myOffset As Double
    Dim myBase As Double
    BuildCapacityLoadPathForResult capacityLoadPath, nValue, userMxValue, userMyValue, referenceX, referenceY, _
        nOffset, nBase, mxOffset, mxBase, myOffset, myBase

    Dim forceOnlyPath As Boolean
    ' Одиночный вывод capacity повторяет batch-логику: если пользовательские
    ' моменты равны нулю, λ*N и λ*NMxy считаются силовой траекторией N. Сам
    ' выбранный CapacityLoadPath не меняется, меняется только численный метод.
    forceOnlyPath = ((capacityLoadPath = "LambdaN" Or capacityLoadPath = "LambdaNMxy") And _
        Abs(userMxValue) <= 0.000000001 And Abs(userMyValue) <= 0.000000001)
    Select Case LCase$(Trim$(settings.GetRawString("Capacity.SolutionStrategy", vbNullString)))
        Case "auto"
            If forceOnlyPath Then
                capacity.SolveByLoadPathMultiplier section, materialProvider.ConcreteMaterial(cpStrength), _
                    materialProvider.SteelMaterial(cpStrength), nOffset, nBase, mxOffset, mxBase, myOffset, myBase, True
            Else
                capacity.SolveByAutoLoadPath section, materialProvider.ConcreteMaterial(cpStrength), _
                    materialProvider.SteelMaterial(cpStrength), nOffset, nBase, mxOffset, mxBase, myOffset, myBase
            End If
        Case "ultimatestrain"
            If forceOnlyPath Then
                capacity.SolveByLoadPathMultiplier section, materialProvider.ConcreteMaterial(cpStrength), _
                    materialProvider.SteelMaterial(cpStrength), nOffset, nBase, mxOffset, mxBase, myOffset, myBase, True
            Else
                capacity.SolveByUltimateLoadPath section, materialProvider.ConcreteMaterial(cpStrength), _
                    materialProvider.SteelMaterial(cpStrength), nOffset, nBase, mxOffset, mxBase, myOffset, myBase
            End If
        Case "loadmultiplier"
            capacity.SolveByLoadPathMultiplier section, materialProvider.ConcreteMaterial(cpStrength), _
                materialProvider.SteelMaterial(cpStrength), nOffset, nBase, mxOffset, mxBase, myOffset, myBase, _
                forceOnlyPath
        Case Else
            Err.Raise vbObjectError + 4125, "WriteCapacityAndCrackResults", _
                "Capacity.SolutionStrategy должен быть Auto, LoadMultiplier или UltimateStrain."
    End Select
    writer.WriteCapacityResult workbook, capacity, units
    Dim service As CSectionSolver
    Set service = New CSectionSolver
    service.ApplySettings settings, units
    service.Solve section, concrete, steel, nValue, internalMxValue, internalMyValue

    If service.Converged Then
        Dim crack As CCrackWidthCalculator
        Set crack = New CCrackWidthCalculator
        crack.ApplySettings settings, units
        If currentPurpose = cpCrackedNDS And CrackCalculationEnabled(settings) And _
                Not ServiceUsesExtension(section, service, concrete, steel) And _
                ServiceWithinPhysicalRange(section, service, concrete, steel) Then
            crack.Calculate service, section, materialProvider, currentPurpose, _
                nValue, internalMxValue, internalMyValue, _
                centroidMxForCrack, centroidMyForCrack
            writer.WriteCrackResult workbook, crack, units
        End If
    End If
End Sub

Private Sub WriteDirectStateAndCrackResults(ByVal workbook As Object, ByVal section As CSectionModel, _
        ByVal materialProvider As CMaterialModelProvider, ByVal currentPurpose As ECalculationPurpose, _
        ByVal settings As CSystemSettingsReader, _
        ByVal units As CUnitSystem, _
        ByVal nValue As Double, ByVal mxValue As Double, ByVal myValue As Double, _
        ByVal centroidMxForCrack As Double, ByVal centroidMyForCrack As Double, ByVal writer As CCapacityResultWriter)

    Dim concrete As Object
    Dim steel As Object
    Set concrete = materialProvider.ConcreteStateMaterial(currentPurpose)
    Set steel = materialProvider.SteelStateMaterial(currentPurpose)

    Dim service As CSectionSolver
    Set service = New CSectionSolver
    service.ApplySettings settings, units
    service.Solve section, concrete, steel, nValue, mxValue, myValue

    writer.WriteDirectSectionResult workbook, service, units
    
    

    If service.Converged And currentPurpose = cpCrackedNDS And CrackCalculationEnabled(settings) And _
            Not ServiceUsesExtension(section, service, concrete, steel) And _
            ServiceWithinPhysicalRange(section, service, concrete, steel) Then
        Dim crack As CCrackWidthCalculator
        Set crack = New CCrackWidthCalculator
        crack.ApplySettings settings, units
        crack.Calculate service, section, materialProvider, currentPurpose, _
            nValue, mxValue, myValue, _
            centroidMxForCrack, centroidMyForCrack
        writer.WriteCrackResult workbook, crack, units
    End If
End Sub

' Нормализует capacity-траекторию для одиночного блока результата на листе
' "Расчет". Здесь повторяется только пограничная логика Excel-слоя: сама
' предельная задача дальше уходит в универсальный CCapacitySolver.
Private Function NormalizedCapacityLoadPathForResult(ByVal capacityLoadPath As String, _
        ByVal nValue As Double, ByVal userMxValue As Double, ByVal userMyValue As Double) As String
    Dim value As String
    value = LCase$(Trim$(capacityLoadPath))
    value = Replace$(value, " ", vbNullString)
    value = Replace$(value, ChrW$(&H3BB), "lambda")
    value = Replace$(value, "λ", "lambda")

    Select Case value
        Case "lambda*mx", "lambdamx", "mx"
            NormalizedCapacityLoadPathForResult = "LambdaMx"
        Case "lambda*my", "lambdamy", "my"
            NormalizedCapacityLoadPathForResult = "LambdaMy"
        Case "lambda*mxy", "lambdamxy", "mxy"
            NormalizedCapacityLoadPathForResult = "LambdaMxy"
        Case "lambda*n", "lambdan", "nload"
            NormalizedCapacityLoadPathForResult = "LambdaN"
        Case "lambda*nmxy", "lambdanmxy", "nmxy", "all"
            NormalizedCapacityLoadPathForResult = "LambdaNMxy"
        Case Else
            If Sqr(userMxValue * userMxValue + userMyValue * userMyValue) > 0.000000001 Then
                NormalizedCapacityLoadPathForResult = "LambdaMxy"
            ElseIf Abs(nValue) > 0.000000001 Then
                NormalizedCapacityLoadPathForResult = "LambdaN"
            End If
    End Select
End Function

' Собирает математические Offset/Base для повторного вывода capacity на лист
' "Расчет". Логика та же, что в CBatchSectionCalculator: пользовательские
' Mx/My отделены от моментов, появившихся из-за смещения точки приложения N.
Private Sub BuildCapacityLoadPathForResult(ByVal loadPath As String, _
        ByVal nValue As Double, ByVal userMxValue As Double, ByVal userMyValue As Double, _
        ByVal referenceX As Double, ByVal referenceY As Double, _
        ByRef nOffset As Double, ByRef nBase As Double, _
        ByRef mxOffset As Double, ByRef mxBase As Double, _
        ByRef myOffset As Double, ByRef myBase As Double)
    Dim mxFromN As Double
    Dim myFromN As Double
    mxFromN = nValue * referenceY
    myFromN = nValue * referenceX

    Select Case loadPath
        Case "LambdaMx"
            nOffset = nValue
            mxOffset = mxFromN
            mxBase = userMxValue
            myOffset = myFromN + userMyValue
        Case "LambdaMy"
            nOffset = nValue
            mxOffset = mxFromN + userMxValue
            myOffset = myFromN
            myBase = userMyValue
        Case "LambdaMxy"
            nOffset = nValue
            mxOffset = mxFromN
            mxBase = userMxValue
            myOffset = myFromN
            myBase = userMyValue
        Case "LambdaN"
            nBase = nValue
            mxOffset = userMxValue
            mxBase = mxFromN
            myOffset = userMyValue
            myBase = myFromN
        Case "LambdaNMxy"
            nBase = nValue
            mxBase = userMxValue + mxFromN
            myBase = userMyValue + myFromN
    End Select
End Sub

' Выбирает расчетную цель материала для повторного вывода одного LC на лист "Расчет".
' Это тот же смысл, что и в batch: I группа использует Strength, II группа -
' CrackedNDS с диаграммами II группы и неработающим растянутым бетоном.
Private Function MaterialPurposeForCalculationType(ByVal calculationType As String) As ECalculationPurpose
    MaterialPurposeForCalculationType = StateBasePurposeForCalculationType(calculationType)
End Function

' Передает CCapacitySolver пределы деформаций из material provider-а. Эти величины
' больше не читаются как отдельные настройки Capacity.*Limit: источник истины -
' автоматически построенная диаграмма материала для Strength.
Private Sub ApplyCapacityLimitsFromProvider(ByVal capacity As CCapacitySolver, _
        ByVal materialProvider As CMaterialModelProvider, ByVal purpose As ECalculationPurpose)
    capacity.ConcreteCompressionLimit = materialProvider.ConcreteCompressionLimit(purpose)
    capacity.ConcreteTensionLimit = materialProvider.ConcreteTensionLimit(purpose)
    capacity.ConcreteTensionLimitEnabled = materialProvider.ConcreteTensionLimitEnabled(purpose)
    capacity.SteelStrainLimit = MaxDouble(Abs(materialProvider.SteelCompressionLimit(purpose)), _
        Abs(materialProvider.SteelTensionLimit(purpose)))
End Sub

' Проверяет, использовал ли прямой solve техническое продолжение диаграммы.
' Это локальная защита итогового блока листа "Расчет"; основной batch пишет
' тот же признак в Results snapshot и именно его используют схема/AutoCAD.
Private Function ServiceUsesExtension(ByVal section As CSectionModel, ByVal solver As CSectionSolver, _
        ByVal concrete As Object, ByVal steel As Object) As Boolean
    Dim i As Long
    Dim strain As Double

    For i = 1 To section.ConcreteCount
        strain = solver.Epsilon0 + solver.KappaX * section.ConcreteY(i) + solver.KappaY * section.ConcreteX(i)
        If concrete.IsInExtensionRange(strain) Then
            ServiceUsesExtension = True
            Exit Function
        End If
    Next i

    For i = 1 To section.RebarCount
        strain = solver.Epsilon0 + solver.KappaX * section.RebarY(i) + solver.KappaY * section.RebarX(i)
        If steel.IsInExtensionRange(strain) Then
            ServiceUsesExtension = True
            Exit Function
        End If
    Next i
End Function

' Проверяет физические пределы прямого НДС в итоговом блоке листа "Расчет".
' Batch делает такую же проверку перед записью Results. Здесь она нужна, чтобы
' повторный вывод одного LC не запускал расчет трещин по состоянию, которое
' формально сошлось, но уже находится за физическими eps_ult.
Private Function ServiceWithinPhysicalRange(ByVal section As CSectionModel, ByVal solver As CSectionSolver, _
        ByVal concrete As Object, ByVal steel As Object) As Boolean
    Dim i As Long
    Dim strain As Double

    For i = 1 To section.ConcreteCount
        strain = solver.Epsilon0 + solver.KappaX * section.ConcreteY(i) + solver.KappaY * section.ConcreteX(i)
        If Not concrete.IsInPhysicalRange(strain) Then Exit Function
    Next i

    For i = 1 To section.RebarCount
        strain = solver.Epsilon0 + solver.KappaX * section.RebarY(i) + solver.KappaY * section.RebarX(i)
        If Not steel.IsInPhysicalRange(strain) Then Exit Function
    Next i

    ServiceWithinPhysicalRange = True
End Function

Private Function CrackCalculationEnabled(ByVal settings As CSystemSettingsReader) As Boolean
    CrackCalculationEnabled = settings.GetBoolean("SLS.Crack.Enabled", True)
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

' Локальный максимум для передачи асимметричных пределов арматуры в старый
' контракт CCapacitySolver, который пока принимает одно абсолютное значение.
Private Function MaxDouble(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then MaxDouble = a Else MaxDouble = b
End Function


















