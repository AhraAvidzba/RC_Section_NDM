Attribute VB_Name = "modTestSectionWorkflow"
Option Explicit

' ==========================================================================
' Проверки независимых этапов формирования, расчёта и оформления сечения
' ==========================================================================
' Выполняются только на отдельной тестовой копии книги. Проверяют полный
' снимок модели, неизменность координат и отсутствие скрытой генерации.

Private Type TWorkflowStats
    Passed As Long ' Количество успешных проверок.
    Failed As Long ' Количество несовпадений.
    Report As String ' Подробный журнал проверок.
End Type

' Сравнивает условие и добавляет именованный результат в журнал.
Private Sub Check(ByRef stats As TWorkflowStats, ByVal name As String, ByVal condition As Boolean)
    If condition Then
        stats.Passed = stats.Passed + 1: stats.Report = stats.Report & "OK: " & name & vbCrLf
    Else
        stats.Failed = stats.Failed + 1: stats.Report = stats.Report & "FAIL: " & name & vbCrLf
    End If
End Sub

' Сравнивает физические величины с допуском значительно меньше расчётного.
Private Sub Near(ByRef stats As TWorkflowStats, ByVal name As String, ByVal actual As Double, ByVal expected As Double)
    Check stats, name, Abs(actual - expected) <= 0.000000000001 * (1# + Abs(expected))
End Sub

' Меняет настройку по ключу в тестовой копии, сохраняя независимость от адресов.
Private Sub Setting(ByVal key As String, ByVal value As Variant)
    Dim area As Object, row As Long, values As Variant
    Set area = ThisWorkbook.Names("rngSystemSettings").RefersToRange
    values = area.Value2
    For row = 1 To UBound(values, 1)
        If CStr(values(row, 1)) = key Then area.Cells(row, 2).Value2 = value: Exit Sub
    Next row
    Err.Raise 5, "modTestSectionWorkflow", "Не найдена настройка " & key
End Sub

' Выбирает выходную единицу по строке общей таблицы единиц.
Private Sub OutputUnit(ByVal quantity As String, ByVal value As String)
    Dim area As Object, row As Long, values As Variant
    Set area = ThisWorkbook.Names("rngUnitSettings").RefersToRange
    values = area.Value2
    For row = 1 To UBound(values, 1)
        If CStr(values(row, 1)) = quantity Then area.Cells(row, 4).Value2 = value: Exit Sub
    Next row
    Err.Raise 5, "modTestSectionWorkflow", "Не найдена единица " & quantity
End Sub

' Создаёт контрольную модель с исходными материалами, именами и связями.
' Размеры и численные значения заданы независимо от генератора рецептов.
Private Function ExampleModel() As CSectionModel
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
    Set ExampleModel = model
End Function

' Записывает и восстанавливает модель при разных выходных единицах.
' Проверяет данные, которые прежнее чтение только изображения теряло.
Private Sub TestCompleteModel(ByRef stats As TWorkflowStats)
    Dim unit As Variant, expected As CSectionModel, actual As CSectionModel
    Dim settings As CSystemSettingsReader, units As CUnitSystem, writer As CNDMResultsWriter
    For Each unit In Array("mm", "cm", "m")
        OutputUnit "Length", CStr(unit)
        OutputUnit "Area", "cm2"
        Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
        Set units = New CUnitSystem: units.LoadFromSettings settings
        Set expected = ExampleModel(): Set writer = New CNDMResultsWriter
        writer.WriteGeometryPreview ThisWorkbook, expected, PrepareSectionSnapshot(expected), units
        Set actual = ReadFormedSectionFromResults(ThisWorkbook)
        Check stats, "snapshot." & unit & ".bulkReads", SectionSnapshotReadDiagnostics() = "tables=5;released=True"
        Check stats, "snapshot." & unit & ".counts", actual.ConcreteCount = 2 And actual.RebarCount = 2 And actual.AnnotationCount = 5
        Check stats, "snapshot." & unit & ".source", actual.SourceType = expected.SourceType And actual.ConcreteMeshSource = expected.ConcreteMeshSource
        Check stats, "snapshot." & unit & ".diagnostics", actual.GeometryDiagnosticsText = expected.GeometryDiagnosticsText
        Check stats, "snapshot." & unit & ".contours", actual.Contours.Count = 4
        Dim i As Long
        For i = 1 To 2
            Near stats, "snapshot." & unit & ".concrete.x." & i, actual.ConcreteX(i), expected.ConcreteX(i)
            Near stats, "snapshot." & unit & ".concrete.area." & i, actual.ConcreteArea(i), expected.ConcreteArea(i)
            Near stats, "snapshot." & unit & ".concrete.Ix." & i, actual.ConcreteLocalIx(i), expected.ConcreteLocalIx(i)
            Near stats, "snapshot." & unit & ".concrete.Ixy." & i, actual.ConcreteLocalIxy(i), expected.ConcreteLocalIxy(i)
            Check stats, "snapshot." & unit & ".concrete.material." & i, actual.ConcreteMaterialID(i) = expected.ConcreteMaterialID(i)
            Check stats, "snapshot." & unit & ".concrete.name." & i, actual.ConcreteSourceName(i) = expected.ConcreteSourceName(i)
            Check stats, "snapshot." & unit & ".concrete.handle." & i, actual.ConcreteSourceHandle(i) = expected.ConcreteSourceHandle(i)
            Check stats, "snapshot." & unit & ".steel." & i, actual.RebarSteelClass(i) = expected.RebarSteelClass(i)
            Check stats, "snapshot." & unit & ".rebar.name." & i, actual.RebarSourceName(i) = expected.RebarSourceName(i)
            Check stats, "snapshot." & unit & ".rebar.handle." & i, actual.RebarSourceHandle(i) = expected.RebarSourceHandle(i)
            Check stats, "snapshot." & unit & ".material." & i, actual.RebarMaterialID(i) = expected.RebarMaterialID(i)
            Check stats, "snapshot." & unit & ".trace." & i, actual.RebarRecipeTrace(i) = expected.RebarRecipeTrace(i)
            Check stats, "snapshot." & unit & ".comment." & i, actual.RebarComment(i) = expected.RebarComment(i)
            Near stats, "snapshot." & unit & ".rebar.x." & i, actual.RebarX(i), expected.RebarX(i)
            Near stats, "snapshot." & unit & ".rebar.area." & i, actual.RebarArea(i), expected.RebarArea(i)
            Check stats, "snapshot." & unit & ".annotation.text." & i, actual.Annotations.Text(i) = expected.Annotations.Text(i)
            Check stats, "snapshot." & unit & ".annotation.unit." & i, actual.Annotations.Unit(i) = expected.Annotations.Unit(i)
            Near stats, "snapshot." & unit & ".annotation.value." & i, actual.Annotations.Value(i), expected.Annotations.Value(i)
        Next i
    Next unit
End Sub

' Проверяет, что оформление ошибок сохраняет причину, слои и отдельные шаги.
Private Sub TestFormationMessages(ByRef stats As TWorkflowStats)
    Dim settings As CSystemSettingsReader: Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook
    Dim text As String
    text = SectionFormationFailureText("На слоях импорта не найдена геометрия сечения.", "AutoCAD", settings)
    Check stats, "message.cause", InStr(text, "не найдена геометрия") > 0
    Check stats, "message.layers.paragraph", InStr(text, vbCrLf & vbCrLf & "Слои, указанные на Config:" & vbCrLf) > 0
    Check stats, "message.solution.paragraph", InStr(text, vbCrLf & vbCrLf & "Возможное решение:" & vbCrLf) > 0
    Check stats, "message.generated", InStr(text, "Geometry.Source = Generated") > 0
    Check stats, "message.sections", InStr(text, vbCrLf & "2. На листе «Сечения»") > 0
    Check stats, "message.retry", InStr(text, vbCrLf & "3. Нажмите «Сформировать сечение».") > 0
    text = WorkbookErrorMessage("Сечение не сформировано", vbObjectError + 4405, "Import", text)
    Check stats, "message.operation.paragraph", Left$(text, Len("Сечение не сформировано:" & vbCrLf & vbCrLf)) = "Сечение не сформировано:" & vbCrLf & vbCrLf
End Sub


' Повреждённые сведения не заменяются значениями по умолчанию. После ошибки
' кэш очищается, а перенос именованной таблицы не меняет результат чтения.
Private Sub TestSnapshotValidation(ByRef stats As TWorkflowStats)
    Dim model As CSectionModel, writer As CNDMResultsWriter, units As CUnitSystem
    Dim settings As CSystemSettingsReader, table As Object, annotations As Object
    Dim original As Variant, annotationData As Variant, changed As Variant, mode As Long
    Dim code As Long, reason As String, oldReference As String
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    Set writer = New CNDMResultsWriter: Set model = ExampleModel()
    writer.WriteGeometryPreview ThisWorkbook, model, PrepareSectionSnapshot(model), units
    Set table = ThisWorkbook.Names("rngNDMSectionModelMetadata").RefersToRange
    Set annotations = ThisWorkbook.Names("rngNDMSectionAnnotations").RefersToRange
    original = ReadAnchoredResultTable(ThisWorkbook, "rngNDMSectionModelMetadata")
    annotationData = ReadAnchoredResultTable(ThisWorkbook, "rngNDMSectionAnnotations")
    For mode = 0 To 5
        changed = original
        Select Case mode
            Case 0: changed(2, 6) = 0
            Case 1: changed(3, 3) = changed(2, 3)
            Case 2: changed(2, 1) = "OTHER_SNAPSHOT"
            Case 3
                table.Offset(1, 12).NumberFormat = "@"
                changed(2, 13) = "False"
            Case 4: changed(1, 14) = "WrongHeader"
            Case 5: annotations.Offset(1, 0).Value2 = "OTHER_SNAPSHOT"
        End Select
        table.Resize(UBound(changed, 1), UBound(changed, 2)).Value2 = changed
        Set model = Nothing
        On Error Resume Next
        Set model = ReadFormedSectionFromResults(ThisWorkbook)
        code = Err.Number: reason = Err.Description: Err.Clear
        On Error GoTo 0
        Check stats, "validation." & mode & ".reject", code <> 0 And model Is Nothing
        Check stats, "validation." & mode & ".action", InStr(reason, "Сформировать сечение") > 0
        Check stats, "validation." & mode & ".released", InStr(SectionSnapshotReadDiagnostics(), "released=True") > 0
        table.Offset(1, 12).NumberFormat = "General"
        annotations.Resize(UBound(annotationData, 1), UBound(annotationData, 2)).Value2 = annotationData
    Next mode
    table.Resize(UBound(original, 1), UBound(original, 2)).Value2 = original
    oldReference = ThisWorkbook.Names("rngNDMSectionModelMetadata").RefersTo
    Dim moved As Object: Set moved = table.Worksheet.Range("QZ156")
    moved.Resize(UBound(original, 1), UBound(original, 2)).Value2 = original
    ThisWorkbook.Names("rngNDMSectionModelMetadata").RefersTo = "=Results!" & moved.Address
    On Error GoTo RestoreName
    Set model = ReadFormedSectionFromResults(ThisWorkbook)
    Check stats, "validation.moved.metadata", model.ConcreteCount = 2 And model.RebarCount = 2
RestoreName:
    code = Err.Number: reason = Err.Description
    ThisWorkbook.Names("rngNDMSectionModelMetadata").RefersTo = oldReference
    moved.Resize(UBound(original, 1), UBound(original, 2)).ClearContents
    If code <> 0 Then Err.Raise code, "TestSnapshotValidation", reason
End Sub

' Подтверждает отсутствие тихой генерации при отсутствии сформированного снимка.
Private Sub TestMissingSnapshot(ByRef stats As TWorkflowStats)
    Dim writer As CNDMResultsWriter: Set writer = New CNDMResultsWriter
    writer.ClearResults ThisWorkbook
    Dim model As CSectionModel, number As Long, reason As String
    On Error Resume Next
    Set model = ReadFormedSectionFromResults(ThisWorkbook)
    number = Err.Number: reason = Err.Description
    On Error GoTo 0
    Check stats, "missing.reject", number <> 0 And model Is Nothing
    Check stats, "missing.action", InStr(reason, "Сформировать сечение") > 0
End Sub

' ДЛЯ ТЕСТОВ
' Быстрая приёмка сохранения полного сечения и понятных сообщений без расчёта НДС.
Public Function RunWorkflowSnapshotTests() As String
    Dim stats As TWorkflowStats, saved As Variant
    saved = FixtureSaveInputs(ThisWorkbook.Worksheets("Config").UsedRange)
    On Error GoTo Failed
    TestCompleteModel stats
    TestFormationMessages stats
    TestSnapshotValidation stats
    TestMissingSnapshot stats
Done:
    FixtureRestoreInputs ThisWorkbook.Worksheets("Config").UsedRange, saved
    stats.Report = stats.Report & "TOTAL_WORKFLOW_SNAPSHOT: passed=" & stats.Passed & "; failed=" & stats.Failed & vbCrLf
    RunWorkflowSnapshotTests = stats.Report
    Exit Function
Failed:
    Check stats, "unexpected." & Err.Number & "." & Err.Description, False
    Resume Done
End Function

' ДЛЯ ТЕСТОВ
' Замеряет только чтение готового снимка; подготовка и запись в замер не входят.
' Все обращения к данным проверяются по счётчику блочных чтений.
Public Function RunWorkflowReadPerformanceTests() As String
    Dim stats As TWorkflowStats, saved As Variant, count As Variant
    saved = FixtureSaveInputs(ThisWorkbook.Worksheets("Config").UsedRange)
    On Error GoTo Failed
    OutputUnit "Length", "mm": OutputUnit "Area", "mm2"
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
            Check stats, "read." & count & "." & pass & ".count", actual.RebarCount = CLng(count)
            Check stats, "read." & count & "." & pass & ".tables", SectionSnapshotReadDiagnostics() = "tables=5;released=True"
            Near stats, "read." & count & "." & pass & ".lastX", actual.RebarX(CLng(count)), model.RebarX(CLng(count))
            stats.Report = stats.Report & "WORKFLOW_READ: rebars=" & count & "; seconds=" & _
                Replace$(CStr(elapsed), ",", ".") & "; " & SectionSnapshotReadDiagnostics() & vbCrLf
        Next pass
    Next count
Done:
    FixtureRestoreInputs ThisWorkbook.Worksheets("Config").UsedRange, saved
    stats.Report = stats.Report & "TOTAL_WORKFLOW_READ: passed=" & stats.Passed & "; failed=" & stats.Failed & vbCrLf
    RunWorkflowReadPerformanceTests = stats.Report
    Exit Function
Failed:
    Check stats, "unexpected.read." & Err.Number & "." & Err.Description, False
    Resume Done
End Function
