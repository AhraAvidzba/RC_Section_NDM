Attribute VB_Name = "modTestSectionWorkflow"
Option Explicit

' ==========================================================================
' ДЛЯ ТЕСТОВ: Проверки независимых этапов формирования, расчёта и оформления сечения
' ==========================================================================
' Выполняются только на отдельной тестовой копии книги. Проверяют полный
' снимок модели, неизменность координат и отсутствие скрытой генерации.

Private Type TWorkflowStats
    Passed As Long ' Количество успешных проверок.
    Failed As Long ' Количество несовпадений.
    Report As String ' Подробный журнал проверок.
End Type

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает условие и добавляет именованный результат в журнал.
Private Sub TestCheck(ByRef stats As TWorkflowStats, ByVal name As String, ByVal condition As Boolean)
    If condition Then
        stats.Passed = stats.Passed + 1: stats.Report = stats.Report & "OK: " & name & vbCrLf
    Else
        stats.Failed = stats.Failed + 1: stats.Report = stats.Report & "FAIL: " & name & vbCrLf
    End If
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает физические величины с допуском значительно меньше расчётного.
Private Sub TestNear(ByRef stats As TWorkflowStats, ByVal name As String, ByVal actual As Double, ByVal expected As Double)
    TestCheck stats, name, Abs(actual - expected) <= 0.000000000001 * (1# + Abs(expected))
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Меняет настройку по ключу в тестовой копии, сохраняя независимость от адресов.
Private Sub TestSetting(ByVal key As String, ByVal value As Variant)
    Dim area As Object, row As Long, values As Variant
    Set area = ThisWorkbook.Names("rngSystemSettings").RefersToRange
    values = area.Value2
    For row = 1 To UBound(values, 1)
        If CStr(values(row, 1)) = key Then area.Cells(row, 2).Value2 = value: Exit Sub
    Next row
    Err.Raise 5, "modTestSectionWorkflow", "Не найдена настройка " & key
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Выбирает выходную единицу по строке общей таблицы единиц.
Private Sub TestOutputUnit(ByVal quantity As String, ByVal value As String)
    Dim area As Object, row As Long, values As Variant
    Set area = ThisWorkbook.Names("rngUnitSettings").RefersToRange
    values = area.Value2
    For row = 1 To UBound(values, 1)
        If CStr(values(row, 1)) = quantity Then area.Cells(row, 4).Value2 = value: Exit Sub
    Next row
    Err.Raise 5, "modTestSectionWorkflow", "Не найдена единица " & quantity
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Создаёт контрольную модель с исходными материалами, именами и связями.
' Размеры и численные значения заданы независимо от генератора рецептов.
Private Function TestExampleModel() As CSectionModel
    Dim model As CSectionModel: Set model = New CSectionModel
    model.SourceType = "AutoCADImport": model.ConcreteMeshSource = "AutoCADRegions"
    model.AddConcreteElement 0#, 0#, 10000#, 2, "00123", "1E2", "Rectangle", 100#, 100#, 0.25, "Основная область", 0#, 0#, 0#, True
    model.AddConcreteElement 60#, 0#, 2500#, 3, "Бетон 2", "1B", "Region", 0#, 0#, 0#, "Область с реальными инерциями", 100000#, 120000#, 7000#, False
    model.AddRebarElement -20#, -20#, 20#, 314.159265358979, "A500Б", 4, "Стержень 1", "2A", "Главная позиция"
    model.AddRebarElement 20#, 20#, 16#, 201.061929829747, "A400", 5, "Стержень 2", "2B", "Дополнительная позиция"
    model.AppendRebarRecipeTrace 1, "Main|A", "Axis;1", "Main.1"
    model.AppendRebarRecipeTrace 1, "Row2", "Axis;2", "Main.2"
    model.GeometryDiagnosticsText = "Предупреждение: контрольное сообщение." & vbCrLf & "Позиции сохранены."
    model.Contours.AddContourLine "B1", -100#, -100#, 100#, -100#, "Низ", "OUTER", "Outer", "SOURCE"
    model.Contours.AddContourLine "B2", 100#, -100#, 100#, 100#, "Право", "OUTER", "Outer", "SOURCE"
    model.Contours.AddContourLine "B3", 100#, 100#, -100#, 100#, "Верх", "OUTER", "Outer", "SOURCE"
    model.Contours.AddContourLine "B4", -100#, 100#, -100#, -100#, "Лево", "OUTER", "Outer", "SOURCE"
    model.Annotations.AddDimension "DIM", 0#, 100#, 85#, 100#, 0#, 1#, "B", 123.456789, "Точное значение без округления"
    model.Annotations.AddRebarLabel "LABEL", -20#, -30#, 20#, -30#, 0#, -1#, "2" & ChrW$(&H2205) & "20", "Подпись фактической арматуры"
    model.ApplyAverageRotationToEquivalentAreaFallbacks
    Set TestExampleModel = model
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Записывает и восстанавливает модель при разных выходных единицах.
' Проверяет данные, которые прежнее чтение только изображения теряло.
Private Sub TestCompleteModel(ByRef stats As TWorkflowStats)
    Dim unit As Variant, expected As CSectionModel, actual As CSectionModel
    Dim settings As CSystemSettingsReader, units As CUnitSystem, writer As CNDMResultsWriter
    For Each unit In Array("mm", "cm", "m")
        TestOutputUnit "Length", CStr(unit)
        TestOutputUnit "Area", "cm2"
        Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
        Set units = New CUnitSystem: units.LoadFromSettings settings
        Set expected = TestExampleModel(): Set writer = New CNDMResultsWriter
        writer.WriteGeometryPreview ThisWorkbook, expected, PrepareSectionSnapshot(expected), units
        Set actual = ReadFormedSectionFromResults(ThisWorkbook, True)
        stats.Report = stats.Report & "WORKFLOW_SNAPSHOT_READ: unit=" & CStr(unit) & "; " & TestSectionSnapshotReadDiagnostics() & vbCrLf
        TestCheck stats, "snapshot." & unit & ".bulkReads", TestSectionSnapshotReadDiagnostics() = "tables=5;released=True"
        TestCheck stats, "snapshot." & unit & ".counts", actual.ConcreteCount = 2 And actual.RebarCount = 2 And actual.AnnotationCount = 5
        TestCheck stats, "snapshot." & unit & ".source", actual.SourceType = expected.SourceType And actual.ConcreteMeshSource = expected.ConcreteMeshSource
        TestCheck stats, "snapshot." & unit & ".diagnostics", actual.GeometryDiagnosticsText = expected.GeometryDiagnosticsText
        TestCheck stats, "snapshot." & unit & ".contours", actual.Contours.Count = 4
        Dim i As Long
        For i = 1 To 2
            TestNear stats, "snapshot." & unit & ".concrete.x." & i, actual.ConcreteX(i), expected.ConcreteX(i)
            TestNear stats, "snapshot." & unit & ".concrete.area." & i, actual.ConcreteArea(i), expected.ConcreteArea(i)
            TestNear stats, "snapshot." & unit & ".concrete.Ix." & i, actual.ConcreteLocalIx(i), expected.ConcreteLocalIx(i)
            TestNear stats, "snapshot." & unit & ".concrete.Ixy." & i, actual.ConcreteLocalIxy(i), expected.ConcreteLocalIxy(i)
            TestCheck stats, "snapshot." & unit & ".concrete.material." & i, actual.ConcreteMaterialID(i) = expected.ConcreteMaterialID(i)
            TestCheck stats, "snapshot." & unit & ".concrete.name." & i, actual.ConcreteSourceName(i) = expected.ConcreteSourceName(i)
            TestCheck stats, "snapshot." & unit & ".concrete.handle." & i, actual.ConcreteSourceHandle(i) = expected.ConcreteSourceHandle(i)
            TestCheck stats, "snapshot." & unit & ".steel." & i, actual.RebarSteelClass(i) = expected.RebarSteelClass(i)
            TestCheck stats, "snapshot." & unit & ".rebar.name." & i, actual.RebarSourceName(i) = expected.RebarSourceName(i)
            TestCheck stats, "snapshot." & unit & ".rebar.handle." & i, actual.RebarSourceHandle(i) = expected.RebarSourceHandle(i)
            TestCheck stats, "snapshot." & unit & ".material." & i, actual.RebarMaterialID(i) = expected.RebarMaterialID(i)
            TestCheck stats, "snapshot." & unit & ".trace." & i, actual.RebarRecipeTrace(i) = expected.RebarRecipeTrace(i)
            TestCheck stats, "snapshot." & unit & ".comment." & i, actual.RebarComment(i) = expected.RebarComment(i)
            TestNear stats, "snapshot." & unit & ".rebar.x." & i, actual.RebarX(i), expected.RebarX(i)
            TestNear stats, "snapshot." & unit & ".rebar.area." & i, actual.RebarArea(i), expected.RebarArea(i)
            TestCheck stats, "snapshot." & unit & ".annotation.text." & i, actual.Annotations.Text(i) = expected.Annotations.Text(i)
            TestCheck stats, "snapshot." & unit & ".annotation.unit." & i, actual.Annotations.Unit(i) = expected.Annotations.Unit(i)
            TestNear stats, "snapshot." & unit & ".annotation.value." & i, actual.Annotations.Value(i), expected.Annotations.Value(i)
        Next i
    Next unit
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет, что оформление ошибок сохраняет причину, слои и отдельные шаги.
Private Sub TestFormationMessages(ByRef stats As TWorkflowStats)
    Dim settings As CSystemSettingsReader: Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook
    Dim text As String
    text = SectionFormationFailureText("На слоях импорта не найдена геометрия сечения.", "AutoCAD", settings)
    TestCheck stats, "message.cause", InStr(text, "не найдена геометрия") > 0
    TestCheck stats, "message.layers.paragraph", InStr(text, vbCrLf & vbCrLf & "Слои, указанные на Config:" & vbCrLf) > 0
    TestCheck stats, "message.solution.paragraph", InStr(text, vbCrLf & vbCrLf & "Возможное решение:" & vbCrLf) > 0
    TestCheck stats, "message.generated", InStr(text, "Geometry.Source = Generated") > 0
    TestCheck stats, "message.sections", InStr(text, vbCrLf & "2. На листе «Сечения»") > 0
    TestCheck stats, "message.retry", InStr(text, vbCrLf & "3. Нажмите «Сформировать сечение».") > 0
    text = WorkbookErrorMessage("Сечение не сформировано", vbObjectError + 4405, "Import", text)
    TestCheck stats, "message.operation.paragraph", Left$(text, Len("Сечение не сформировано:" & vbCrLf & vbCrLf)) = "Сечение не сформировано:" & vbCrLf & vbCrLf
End Sub


' -------------------- ДЛЯ ТЕСТОВ --------------------
' Повреждённые сведения не заменяются значениями по умолчанию. После ошибки
' кэш очищается, а перенос именованной таблицы не меняет результат чтения.
Private Sub TestSnapshotValidation(ByRef stats As TWorkflowStats)
    Dim model As CSectionModel, writer As CNDMResultsWriter, units As CUnitSystem
    Dim settings As CSystemSettingsReader, table As Object, annotations As Object
    Dim original As Variant, annotationData As Variant, changed As Variant, mode As Long
    Dim code As Long, reason As String, oldReference As String
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    Set writer = New CNDMResultsWriter: Set model = TestExampleModel()
    writer.WriteGeometryPreview ThisWorkbook, model, PrepareSectionSnapshot(model), units
    Set table = ThisWorkbook.Names("rngNDMSectionGeometry").RefersToRange
    Set annotations = ThisWorkbook.Names("rngNDMSectionAnnotations").RefersToRange
    original = ReadAnchoredResultTable(ThisWorkbook, "rngNDMSectionGeometry")
    annotationData = ReadAnchoredResultTable(ThisWorkbook, "rngNDMSectionAnnotations")
    For mode = 0 To 5
        changed = original
        Select Case mode
            Case 0: changed(2, 6) = 0
            Case 1: changed(3, 2) = changed(2, 2)
            Case 2: changed(2, 1) = "OTHER_SNAPSHOT"
            Case 3: changed(2, 8) = False
            Case 4: changed(2, 14) = "not a number"
            Case 5: changed(4, 10) = -10#
        End Select
        table.Resize(UBound(changed, 1), UBound(changed, 2)).Value2 = changed
        Set model = Nothing
        On Error Resume Next
        Set model = ReadFormedSectionFromResults(ThisWorkbook)
        code = Err.Number: reason = Err.Description: Err.Clear
        On Error GoTo 0
        TestCheck stats, "validation." & mode & ".reject", code <> 0 And model Is Nothing
        TestCheck stats, "validation." & mode & ".action", InStr(reason, "Сформировать сечение") > 0
        TestCheck stats, "validation." & mode & ".released", InStr(TestSectionSnapshotReadDiagnostics(), "released=True") > 0
        annotations.Resize(UBound(annotationData, 1), UBound(annotationData, 2)).Value2 = annotationData
    Next mode
    table.Resize(UBound(original, 1), UBound(original, 2)).Value2 = original
    oldReference = ThisWorkbook.Names("rngNDMSectionGeometry").RefersTo
    Dim moved As Object: Set moved = table.Worksheet.Range("QZ156")
    moved.Resize(UBound(original, 1), UBound(original, 2)).Value2 = original
    ThisWorkbook.Names("rngNDMSectionGeometry").RefersTo = "=Results!" & moved.Address
    On Error GoTo RestoreName
    Set model = ReadFormedSectionFromResults(ThisWorkbook)
    TestCheck stats, "validation.moved.geometry", model.ConcreteCount = 2 And model.RebarCount = 2
RestoreName:
    code = Err.Number: reason = Err.Description
    ThisWorkbook.Names("rngNDMSectionGeometry").RefersTo = oldReference
    moved.Resize(UBound(original, 1), UBound(original, 2)).ClearContents
    If code <> 0 Then Err.Raise code, "TestSnapshotValidation", reason
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Подтверждает отсутствие тихой генерации при отсутствии сформированного снимка.
Private Sub TestMissingSnapshot(ByRef stats As TWorkflowStats)
    Dim writer As CNDMResultsWriter: Set writer = New CNDMResultsWriter
    writer.ClearResults ThisWorkbook
    Dim model As CSectionModel, number As Long, reason As String
    On Error Resume Next
    Set model = ReadFormedSectionFromResults(ThisWorkbook)
    number = Err.Number: reason = Err.Description
    On Error GoTo 0
    TestCheck stats, "missing.reject", number <> 0 And model Is Nothing
    TestCheck stats, "missing.action", InStr(reason, "Сформировать сечение") > 0
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' ДЛЯ ТЕСТОВ
' Быстрая приёмка сохранения полного сечения и понятных сообщений без расчёта НДС.
Public Function RunWorkflowSnapshotTests() As String
    Dim stats As TWorkflowStats, saved As Variant
    saved = TestFixtureSaveInputs(ThisWorkbook.Worksheets("Config").UsedRange)
    On Error GoTo Failed
    TestCompleteModel stats
    TestFormationMessages stats
    TestSnapshotValidation stats
    TestMissingSnapshot stats
Done:
    TestFixtureRestoreInputs ThisWorkbook.Worksheets("Config").UsedRange, saved
    stats.Report = stats.Report & "TOTAL_WORKFLOW_SNAPSHOT: passed=" & stats.Passed & "; failed=" & stats.Failed & vbCrLf
    RunWorkflowSnapshotTests = stats.Report
    Exit Function
Failed:
    TestCheck stats, "unexpected." & Err.Number & "." & Err.Description, False
    Resume Done
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' ДЛЯ ТЕСТОВ
' Замеряет только чтение готового снимка; подготовка и запись в замер не входят.
' Четыре обязательные таблицы (свойства, элементы, общие контуры и границы элементов)
' читаются по одному разу независимо от количества стержней; аннотации дают пятое чтение.
Public Function RunWorkflowReadPerformanceTests() As String
    Dim stats As TWorkflowStats, saved As Variant, count As Variant
    saved = TestFixtureSaveInputs(ThisWorkbook.Worksheets("Config").UsedRange)
    On Error GoTo Failed
    TestOutputUnit "Length", "mm": TestOutputUnit "Area", "mm2"
    Dim settings As CSystemSettingsReader: Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook
    Dim units As CUnitSystem: Set units = New CUnitSystem
    units.LoadFromSettings settings
    Dim writer As CNDMResultsWriter: Set writer = New CNDMResultsWriter
    For Each count In Array(100, 1000, 5000)
        Dim model As CSectionModel: Set model = New CSectionModel
        model.SourceType = "Generated": model.ConcreteMeshSource = "Generated"
        model.ReserveElements 1, CLng(count)
        model.AddConcreteElement 0#, 0#, 1000000#, 1, , , "Rectangle", 1000#, 1000#
        Dim i As Long
        For i = 1 To CLng(count)
            model.AddRebarElement CDbl(i Mod 100) - 50#, CDbl(i \ 100) - 25#, 10#, 0#, "A400"
        Next i
        writer.WriteGeometryPreview ThisWorkbook, model, PrepareSectionSnapshot(model), units
        Dim start As Double, elapsed As Double, actual As CSectionModel, pass As Long
        For pass = 1 To 3
            start = Timer
            Set actual = ReadFormedSectionFromResults(ThisWorkbook)
            elapsed = Timer - start: If elapsed < 0# Then elapsed = elapsed + 86400#
            TestCheck stats, "read." & count & "." & pass & ".count", actual.RebarCount = CLng(count)
            TestCheck stats, "read." & count & "." & pass & ".tables", TestSectionSnapshotReadDiagnostics() = "tables=4;released=True"
            TestNear stats, "read." & count & "." & pass & ".lastX", actual.RebarX(CLng(count)), model.RebarX(CLng(count))
            stats.Report = stats.Report & "WORKFLOW_READ: rebars=" & count & "; seconds=" & _
                Replace$(CStr(elapsed), ",", ".") & "; " & TestSectionSnapshotReadDiagnostics() & vbCrLf
        Next pass
    Next count
Done:
    TestFixtureRestoreInputs ThisWorkbook.Worksheets("Config").UsedRange, saved
    stats.Report = stats.Report & "TOTAL_WORKFLOW_READ: passed=" & stats.Passed & "; failed=" & stats.Failed & vbCrLf
    RunWorkflowReadPerformanceTests = stats.Report
    Exit Function
Failed:
    TestCheck stats, "unexpected.read." & Err.Number & "." & Err.Description, False
    Resume Done
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет минимальное и полное восстановление, независимость JSON и
' аннотаций от расчётной геометрии и верхнюю привязку импортированной ALL.
Public Function RunCompactSnapshotTests() As String
    Dim stats As TWorkflowStats, saved As Variant, unit As Variant
    saved = TestFixtureSaveInputs(ThisWorkbook.Worksheets("Config").UsedRange)
    On Error GoTo Failed
    Dim settings As CSystemSettingsReader, units As CUnitSystem, writer As CNDMResultsWriter
    Dim expected As CSectionModel, actual As CSectionModel, full As CSectionModel
    Dim expectedProperties As CSectionPropertiesCalculator, actualProperties As CSectionPropertiesCalculator
    Dim geometry As Object, annotations As Object, original As Variant, changed As Variant, mode As Long, code As Long, reason As String
    For Each unit In Array("mm", "cm", "m")
        TestOutputUnit "Length", CStr(unit): TestOutputUnit "Area", "cm2"
        Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
        Set units = New CUnitSystem: units.LoadFromSettings settings
        Set expected = TestExampleModel(): Set writer = New CNDMResultsWriter
        writer.WriteGeometryPreview ThisWorkbook, expected, PrepareSectionSnapshot(expected), units
        Set actual = ReadFormedSectionFromResults(ThisWorkbook)
        stats.Report = stats.Report & "WORKFLOW_COMPACT_READ: mode=minimum; unit=" & CStr(unit) & "; " & TestSectionSnapshotReadDiagnostics() & vbCrLf
        TestCheck stats, "compact." & unit & ".minimumTables", TestSectionSnapshotReadDiagnostics() = "tables=4;released=True"
        TestCheck stats, "compact." & unit & ".noPresentation", actual.AnnotationCount = 0 And Len(actual.GeometryDiagnosticsText) = 0 And Len(actual.RebarRecipeTrace(1)) = 0
        Set full = ReadFormedSectionFromResults(ThisWorkbook, True)
        stats.Report = stats.Report & "WORKFLOW_COMPACT_READ: mode=full; unit=" & CStr(unit) & "; " & TestSectionSnapshotReadDiagnostics() & vbCrLf
        TestCheck stats, "compact." & unit & ".fullTables", TestSectionSnapshotReadDiagnostics() = "tables=5;released=True"
        TestCheck stats, "compact." & unit & ".trace", full.RebarRecipeTrace(1) = expected.RebarRecipeTrace(1)
        TestCheck stats, "compact." & unit & ".source", full.ConcreteSourceName(1) = "00123" And full.ConcreteSourceHandle(1) = "1E2"
        TestCheck stats, "compact." & unit & ".annotation", full.Annotations.Text(1) = "B" And full.Annotations.Unit(1) = "Length"
        TestNear stats, "compact." & unit & ".rawDimension", full.Annotations.Value(1), 123.456789
        Set expectedProperties = PrepareSectionSnapshot(expected)
        Set actualProperties = PrepareSectionSnapshot(actual)
        TestNear stats, "compact." & unit & ".area", actualProperties.Area, expectedProperties.Area
        TestNear stats, "compact." & unit & ".Ix", actualProperties.Ix, expectedProperties.Ix
        TestNear stats, "compact." & unit & ".Iy", actualProperties.Iy, expectedProperties.Iy
        TestNear stats, "compact." & unit & ".Ixy", actualProperties.Ixy, expectedProperties.Ixy
        TestNear stats, "compact." & unit & ".cellWidth", actual.ConcreteWidth(2), expected.ConcreteWidth(2)
        TestNear stats, "compact." & unit & ".cellAngle", actual.ConcreteRotation(2), expected.ConcreteRotation(2)
        Dim savedAnnotations As Variant, readyDimension As String
        savedAnnotations = ReadAnchoredResultTable(ThisWorkbook, "rngNDMSectionAnnotations")
        readyDimension = CStr(savedAnnotations(2, 10))
        ThisWorkbook.Names("rngNDMSectionGeometry").RefersToRange.Offset(1, 15).Resize(4, 1).ClearContents
        ThisWorkbook.Names("rngNDMSectionAnnotations").RefersToRange.Offset(1, 23).Resize(expected.AnnotationCount, 1).ClearContents
        Set full = ReadFormedSectionFromResults(ThisWorkbook, True)
        writer.WriteGeometryPreview ThisWorkbook, full, PrepareSectionSnapshot(full), units
        savedAnnotations = ReadAnchoredResultTable(ThisWorkbook, "rngNDMSectionAnnotations")
        TestCheck stats, "compact." & unit & ".readyTextWithoutJson", CStr(savedAnnotations(2, 10)) = readyDimension
        writer.WriteGeometryPreview ThisWorkbook, expected, PrepareSectionSnapshot(expected), units ' Следующая проверка начинает с исходной записи со всеми дополнениями.
        Set geometry = ThisWorkbook.Names("rngNDMSectionGeometry").RefersToRange
        Set annotations = ThisWorkbook.Names("rngNDMSectionAnnotations").RefersToRange
        original = ReadAnchoredResultTable(ThisWorkbook, "rngNDMSectionGeometry")
        geometry.Offset(0, 14).Resize(1, 2).ClearContents
        Set actual = ReadFormedSectionFromResults(ThisWorkbook)
        TestCheck stats, "compact." & unit & ".mandatoryColumnsOnly", actual.ConcreteCount = 2 And actual.RebarCount = 2
        Set full = ReadFormedSectionFromResults(ThisWorkbook, True)
        TestCheck stats, "compact." & unit & ".fullWithoutOptionalColumns", full.ConcreteCount = 2 And full.RebarCount = 2
        geometry.Resize(UBound(original, 1), UBound(original, 2)).Value2 = original
        annotations.Resize(1 + expected.AnnotationCount, 24).ClearContents
        geometry.Offset(1, 15).Resize(4, 1).Value2 = "не JSON"
        geometry.Offset(1, 14).Resize(4, 1).Value2 = CVErr(xlErrValue)
        Set actual = ReadFormedSectionFromResults(ThisWorkbook)
        TestCheck stats, "compact." & unit & ".optionalErrorsIgnored", actual.ConcreteCount = 2 And actual.RebarCount = 2 And actual.AnnotationCount = 0
        Set actualProperties = PrepareSectionSnapshot(actual)
        TestNear stats, "compact." & unit & ".optionalIx", actualProperties.Ix, expectedProperties.Ix
        geometry.Resize(UBound(original, 1), UBound(original, 2)).Value2 = original
        For mode = 0 To 4
            changed = original
            Select Case mode
                Case 0: changed(2, 4) = ""
                Case 1: changed(2, 6) = -1#
                Case 2: changed(2, 12) = 0#
                Case 3: changed(2, 14) = 1E+30
                Case 4: changed(3, 2) = "C1"
            End Select
            geometry.Resize(UBound(changed, 1), UBound(changed, 2)).Value2 = changed
            Set actual = Nothing
            On Error Resume Next
            Set actual = ReadFormedSectionFromResults(ThisWorkbook)
            code = Err.Number: reason = Err.Description: Err.Clear
            On Error GoTo Failed
            TestCheck stats, "compact." & unit & ".invalid." & mode, code <> 0 And actual Is Nothing
            TestCheck stats, "compact." & unit & ".address." & mode, InStr(reason, "Results!") > 0 And InStr(reason, "Сформировать сечение") > 0
        Next mode
        geometry.Resize(UBound(original, 1), UBound(original, 2)).Value2 = original
    Next unit
    Dim text As String, info As Object
    text = "Кавычки "" и обратная черта \" & vbCrLf & ChrW$(&H2205) & "20" & vbTab & "конец"
    Set info = ReadSectionOptionalJson(SectionOptionalJson("Text", text))
    TestCheck stats, "compact.json.unicode", SectionOptionalText(info, "Text") = text
    For Each unit In Array("{", "{""schema"":""1"",""a"":""x"",""a"":""y""}", "{""schema"":""2""}", "{""schema"":""1""} trailing")
        code = 0
        On Error Resume Next
        Set info = ReadSectionOptionalJson(CStr(unit), "Results!AA157")
        code = Err.Number: reason = Err.Description: Err.Clear
        On Error GoTo Failed
        TestCheck stats, "compact.json.reject." & CStr(unit), code <> 0 And InStr(reason, "Results!AA157") > 0
    Next unit
    TestOutputUnit "Length", "mm"
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    Dim withContour As Long, i As Long, px(1 To 4) As Double, py(1 To 4) As Double, x As Double, y As Double
    Dim reader As CSectionPlotDataReader, found As Boolean, peak As Double
    For withContour = 0 To 1
        Set expected = New CSectionModel: expected.SourceType = "AutoCADImport"
        expected.AddConcreteElement 0#, 0#, 180000#, 1, , , "Rectangle", 600#, 300#, 0.35
        expected.AddRebarElement 0#, 0#, 20#, 0#, "A500"
        If withContour = 1 Then
            For i = 1 To 4
                Select Case i
                    Case 1: x = -300#: y = -150#
                    Case 2: x = 300#: y = -150#
                    Case 3: x = 300#: y = 150#
                    Case 4: x = -300#: y = 150#
                End Select
                px(i) = x * Cos(0.35) - y * Sin(0.35): py(i) = x * Sin(0.35) + y * Cos(0.35)
            Next i
            For i = 1 To 4
                expected.Contours.AddContourLine "E" & i, px(i), py(i), px(i Mod 4 + 1), py(i Mod 4 + 1), , "OUTER", "Outer"
            Next i
        End If
        writer.WriteGeometryPreview ThisWorkbook, expected, PrepareSectionSnapshot(expected), units
        Set reader = New CSectionPlotDataReader: reader.LoadGeometryPreviewFromWorkbook ThisWorkbook, settings
        peak = 300# * Sin(0.35) + 150# * Cos(0.35): found = False
        For i = 1 To reader.AnnotationCount
            If reader.AnnotationType(i) = "REBAR_GLOBAL_ANNOTATION" Then
                found = True
                TestNear stats, "compact.ALL.peak." & withContour, reader.AnnotationStartY(i), peak
                TestNear stats, "compact.ALL.noLocalShift." & withContour, reader.AnnotationBoundaryDistance(i), 0#
            End If
        Next i
        TestCheck stats, "compact.ALL.present." & withContour, found
    Next withContour
Done:
    TestFixtureRestoreInputs ThisWorkbook.Worksheets("Config").UsedRange, saved
    RunCompactSnapshotTests = stats.Report & "TOTAL_COMPACT_SNAPSHOT: passed=" & stats.Passed & "; failed=" & stats.Failed & vbCrLf
    Exit Function
Failed:
    TestCheck stats, "unexpected.compact." & Err.Number & "." & Err.Description, False
    Resume Done
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Доказывает ранний выход очистки без ModelSpace и сохранение ошибок
' недоступного документа вместо ложного сообщения об отсутствии слоёв.
Public Function RunCleanupLayerPreflightTests() As String
    Dim stats As TWorkflowStats, document As CTestFakeAcadDocument, layers As Object, settings As CSystemSettingsReader
    On Error GoTo Failed
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set layers = CreateObject("Scripting.Dictionary")
    Set document = New CTestFakeAcadDocument: document.TestInitializeCleanupLayers layers
    TestCheck stats, "cleanup.noLayers.zeroDeleted", Audit03CleanupAutoCADDocumentForTests(document, settings) = 0
    TestCheck stats, "cleanup.noLayers.noModelSpace", document.TestModelSpaceReads = 0
    TestCheck stats, "cleanup.noLayers.oneLayerTableRead", document.TestGeometryReads = 1
    Dim code As Long
    document.TestInitialize True
    On Error Resume Next
    Audit03CleanupAutoCADDocumentForTests document, settings
    code = Err.Number: Err.Clear
    On Error GoTo Failed
    TestCheck stats, "cleanup.inaccessibleDocument.error", code <> 0 And document.TestModelSpaceReads = 0
Done:
    RunCleanupLayerPreflightTests = stats.Report & "TOTAL_CLEANUP_LAYER_PREFLIGHT: passed=" & stats.Passed & "; failed=" & stats.Failed & vbCrLf
    Exit Function
Failed:
    TestCheck stats, "unexpected.cleanup." & Err.Number & "." & Err.Description, False
    Resume Done
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' При исчезнувшем прямом участке указываются только он и два соседа,
' определяющие его концы. Адрес противоположной грани не добавляется.
Public Function RunOffsetAddressTests() As String
    Dim stats As TWorkflowStats, edges(1 To 4) As TRegionEdge, distances(1 To 4) As Double
    Dim references(1 To 4) As String, values As Variant, i As Long, code As Long, reason As String
    Dim offset As CGeometryVariableOffset, region As CConcreteRegion
    edges(1).X1 = 0#: edges(1).Y1 = 0#: edges(1).X2 = 300#: edges(1).Y2 = 0#
    edges(2).X1 = 300#: edges(2).Y1 = 0#: edges(2).X2 = 300#: edges(2).Y2 = 500#
    edges(3).X1 = 300#: edges(3).Y1 = 500#: edges(3).X2 = 0#: edges(3).Y2 = 500#
    edges(4).X1 = 0#: edges(4).Y1 = 500#: edges(4).X2 = 0#: edges(4).Y2 = 0#
    distances(1) = 300#: distances(2) = 20#: distances(3) = 300#: distances(4) = 20#
    For i = 1 To 4: references(i) = "Сечения!D" & CStr(70 + i): Next i
    values = references: Set offset = New CGeometryVariableOffset
    On Error Resume Next
    Set region = offset.Build(edges, distances, 4, Join(references, "; "), values)
    code = Err.Number: reason = Err.Description: Err.Clear
    On Error GoTo Failed
    TestCheck stats, "offsetAddress.failure", code <> 0 And region Is Nothing
    TestCheck stats, "offsetAddress.reason", InStr(reason, "прямой участок") > 0
    TestCheck stats, "offsetAddress.localNeighbours", InStr(reason, references(1)) > 0 And InStr(reason, references(2)) > 0 And InStr(reason, references(3)) > 0
    TestCheck stats, "offsetAddress.noOpposite", InStr(reason, references(4)) = 0
    TestCheck stats, "offsetAddress.noDuplicateParagraph", UBound(Split(reason, "Адреса ячеек:")) = 1
Done:
    RunOffsetAddressTests = stats.Report & "TOTAL_OFFSET_ADDRESSES: passed=" & stats.Passed & "; failed=" & stats.Failed & vbCrLf
    Exit Function
Failed:
    TestCheck stats, "unexpected.offsetAddress." & Err.Number & "." & Err.Description, False
    Resume Done
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Адресная приёмка компактного снимка и затронутых пользовательских команд.
' Полный ON/OFF и обязательный часовой план этим набором не заменяются.
Public Function RunCompactAcceptanceTests() As String
    Dim report As String
    report = RunCompactSnapshotTests()
    report = report & RunWorkflowSnapshotTests()
    report = report & RunCleanupLayerPreflightTests()
    report = report & RunOffsetAddressTests()
    report = report & RunSectionWorkflowCalculationTests()
    report = report & RunFormationPreviewTests()
    report = report & RunRecipeAnnotationSnapshotTests()
    report = report & RunImportedDimensionTests()
    report = report & RunAnnotationSectionScaleTests()
    RunCompactAcceptanceTests = report
End Function
