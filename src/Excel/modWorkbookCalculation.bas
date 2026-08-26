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
    If InStr(1, message, "ошиб", vbTextCompare) > 0 Or InStr(1, message, "InvalidInput", vbTextCompare) > 0 Then
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

Public Sub UpdateSectionPlotForWorkbook(ByVal workbook As Object)
    If workbook Is Nothing Then Err.Raise vbObjectError + 4140, "UpdateSectionPlotForWorkbook", "Книга Excel не передана."

    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook workbook

    If Not settings.GetBoolean("Plot.Enabled", True) Then Exit Sub

    Dim reader As CSectionPlotDataReader
    Set reader = New CSectionPlotDataReader
    If Not TryLoadFullPlotReader(reader, workbook, settings) Then
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
    report.AddStep "Расчет приведенного центра сечения для пользовательской точки нагрузки."
    ApplyLoadReferenceFromSettings section, materialProvider.ConcreteMaterial("Strength"), _
        materialProvider.SteelMaterial("Strength"), settings, units, batch
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
        UpdateSectionPlotForWorkbook workbook
        report.AddStep "Схема сечения обновлена по сохраненному снимку Results."
    Else
        report.AddSection "Схема"
        report.AddStep "Автообновление схемы пропущено: Plot.AutoUpdateAfterCalculation = No или нет доступных состояний LC."
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
    Dim prefix As String
    prefix = ImportedSectionMessage(section, settings)

    If batch.InvalidInputCount > 0 Then
        BuildCalculationMessage = prefix & "Расчет завершен с ошибками ввода. Обработано сочетаний: " & _
            CStr(batch.Count) & "; ошибок ввода: " & CStr(batch.InvalidInputCount) & "." & vbCrLf & _
            "Проверьте строки со статусом InvalidInput на листе Results." & vbCrLf & _
            "Первая ошибка: " & batch.FirstInvalidInputMessage
    Else
        BuildCalculationMessage = prefix & "Расчет завершен. Обработано сочетаний: " & CStr(batch.Count) & _
            ". Определяющее сочетание: " & batch.GoverningCombinationID
    End If
End Function

Private Function ImportedSectionMessage(ByVal section As CSectionModel, ByVal settings As CSystemSettingsReader) As String
    If section Is Nothing Then Exit Function
    If settings Is Nothing Then Exit Function

    If StrComp(settings.GetRawString("Geometry.Source", "Generated"), "AutoCAD", vbTextCompare) = 0 Then
        ImportedSectionMessage = "Импортировано из AutoCAD: бетонных Region: " & _
            CStr(section.ConcreteCount) & "; арматурных Region: " & CStr(section.RebarCount) & "." & vbCrLf
    End If
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
        ByVal concrete As Object, ByVal steel As Object, ByVal settings As CSystemSettingsReader, _
        ByVal units As CUnitSystem, _
        ByVal batch As CBatchSectionCalculator)
    Dim referenceX As Double
    Dim referenceY As Double
    CalculateTransformedSectionCentroid section, concrete, steel, referenceX, referenceY

    batch.ApplyLoadReference referenceX + units.InputLengthToInternal(settings.GetDouble("Load.ReferenceOffsetX", 0#)), _
        referenceY + units.InputLengthToInternal(settings.GetDouble("Load.ReferenceOffsetY", 0#)), referenceX, referenceY
End Sub

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

    Dim currentPurpose As String
    currentPurpose = MaterialPurposeForCalculationType(batch.CalculationType(index))

    WriteCapacityAndCrackResults workbook, section, materialProvider, currentPurpose, settings, _
        units, batch.N(index), batch.UserMx(index), batch.UserMy(index), _
        batch.LoadReferenceX, batch.LoadReferenceY
End Sub

Private Sub WriteCapacityAndCrackResults(ByVal workbook As Object, ByVal section As CSectionModel, _
        ByVal materialProvider As CMaterialModelProvider, ByVal currentPurpose As String, _
        ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem, _
        ByVal nValue As Double, ByVal userMxValue As Double, ByVal userMyValue As Double, _
        ByVal referenceX As Double, ByVal referenceY As Double)

    Dim writer As CCapacityResultWriter
    Set writer = New CCapacityResultWriter

    Dim concrete As Object
    Dim steel As Object
    Set concrete = materialProvider.ConcreteMaterial(currentPurpose)
    Set steel = materialProvider.SteelMaterial(currentPurpose)

    Dim internalMxValue As Double
    Dim internalMyValue As Double
    internalMxValue = userMxValue + nValue * referenceY
    internalMyValue = userMyValue + nValue * referenceX

    Dim centroidX As Double
    Dim centroidY As Double
    CalculateTransformedSectionCentroid section, concrete, steel, centroidX, centroidY

    Dim centroidMxForCrack As Double
    Dim centroidMyForCrack As Double
    ' Проверка центрального растяжения в расчете трещин выполняется относительно
    ' центра тяжести, поэтому здесь учитываем только эксцентриситет точки
    ' приложения нагрузки от этого центра, а не абсолютные координаты модели.
    centroidMxForCrack = userMxValue + nValue * (referenceY - centroidY)
    centroidMyForCrack = userMyValue + nValue * (referenceX - centroidX)

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
    ApplyCapacityLimitsFromProvider capacity, materialProvider, "Strength"
    If Sqr(userMxValue * userMxValue + userMyValue * userMyValue) > 0.000000001 Then
        Select Case LCase$(Trim$(settings.GetRawString("Capacity.Method", vbNullString)))
            Case "ultimatestrain"
                capacity.SolveByUltimateStrain section, materialProvider.ConcreteMaterial("Strength"), _
                    materialProvider.SteelMaterial("Strength"), nValue, userMxValue, userMyValue, _
                    nValue * referenceY, nValue * referenceX
            Case "loadmultiplier"
                capacity.SolveByLoadMultiplier section, materialProvider.ConcreteMaterial("Strength"), _
                    materialProvider.SteelMaterial("Strength"), nValue, userMxValue, userMyValue, _
                    nValue * referenceY, nValue * referenceX
            Case Else
                Err.Raise vbObjectError + 4125, "WriteCapacityAndCrackResults", _
                    "Capacity.Method должен быть LoadMultiplier или UltimateStrain."
        End Select
    End If
    writer.WriteCapacityResult workbook, capacity, units
    Dim service As CSectionSolver
    Set service = New CSectionSolver
    service.ApplySettings settings, units
    service.Solve section, concrete, steel, nValue, internalMxValue, internalMyValue

    If service.Converged Then
        Dim crack As CCrackWidthCalculator
        Set crack = New CCrackWidthCalculator
        crack.ApplySettings settings, units
        If CrackCalculationEnabled(settings) Then
            crack.Calculate service, section, materialProvider, currentPurpose, _
                nValue, internalMxValue, internalMyValue, _
                centroidMxForCrack, centroidMyForCrack
            writer.WriteCrackResult workbook, crack, units
        End If
    End If
End Sub

Private Sub WriteDirectStateAndCrackResults(ByVal workbook As Object, ByVal section As CSectionModel, _
        ByVal materialProvider As CMaterialModelProvider, ByVal currentPurpose As String, _
        ByVal settings As CSystemSettingsReader, _
        ByVal units As CUnitSystem, _
        ByVal nValue As Double, ByVal mxValue As Double, ByVal myValue As Double, _
        ByVal centroidMxForCrack As Double, ByVal centroidMyForCrack As Double, ByVal writer As CCapacityResultWriter)

    Dim concrete As Object
    Dim steel As Object
    Set concrete = materialProvider.ConcreteMaterial(currentPurpose)
    Set steel = materialProvider.SteelMaterial(currentPurpose)

    Dim service As CSectionSolver
    Set service = New CSectionSolver
    service.ApplySettings settings, units
    service.Solve section, concrete, steel, nValue, mxValue, myValue

    writer.WriteDirectSectionResult workbook, service, units
    
    

    If service.Converged And CrackCalculationEnabled(settings) Then
        Dim crack As CCrackWidthCalculator
        Set crack = New CCrackWidthCalculator
        crack.ApplySettings settings, units
        crack.Calculate service, section, materialProvider, currentPurpose, _
            nValue, mxValue, myValue, _
            centroidMxForCrack, centroidMyForCrack
        writer.WriteCrackResult workbook, crack, units
    End If
End Sub

' Выбирает расчетную цель материала для повторного вывода одного LC на лист "Расчет".
' Это тот же смысл, что и в batch: I группа использует Strength, II группа -
' CrackedNDS с диаграммами II группы и неработающим растянутым бетоном.
Private Function MaterialPurposeForCalculationType(ByVal calculationType As String) As String
    Select Case LCase$(Trim$(calculationType))
        Case "group2", "2", "sls", "crack", "crackonly"
            MaterialPurposeForCalculationType = "CrackedNDS"
        Case Else
            MaterialPurposeForCalculationType = "Strength"
    End Select
End Function

' Передает CCapacitySolver пределы деформаций из material provider-а. Эти величины
' больше не читаются как отдельные настройки Capacity.*Limit: источник истины -
' автоматически построенная диаграмма материала для Strength.
Private Sub ApplyCapacityLimitsFromProvider(ByVal capacity As CCapacitySolver, _
        ByVal materialProvider As CMaterialModelProvider, ByVal purpose As String)
    capacity.ConcreteCompressionLimit = materialProvider.ConcreteCompressionLimit(purpose)
    capacity.ConcreteTensionLimit = materialProvider.ConcreteTensionLimit(purpose)
    capacity.ConcreteTensionLimitEnabled = materialProvider.ConcreteTensionLimitEnabled(purpose)
    capacity.SteelStrainLimit = MaxDouble(Abs(materialProvider.SteelCompressionLimit(purpose)), _
        Abs(materialProvider.SteelTensionLimit(purpose)))
End Sub

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

















