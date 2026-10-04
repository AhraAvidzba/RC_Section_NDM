Attribute VB_Name = "modTestAutoCADConfig"
Option Explicit

' ==========================================================================
' ДЛЯ ТЕСТОВ: CONFIG И РАБОЧИЕ CONSUMERS AUTOCAD
' ==========================================================================
' Проверяет настройки через штатный importer, подготовку экспортного текста,
' выбор цветов/слоев и очистку ModelSpace. Fixtures не выдаются за DWG-тест:
' реальный Autodesk AutoCAD проверяется отдельным entrypoint-ом.
' Исходные формулы Config и положение именованной таблицы восстанавливаются.

Private Type TCadStats
    Passed As Long
    Failed As Long
    Cases As Long ' Число независимых вызовов consumer-а, не assertions.
    Report As String
End Type

' Выполняет проверки значений, отсутствующих строк и динамических адресов
' в двух положениях Config, не подключаясь к AutoCAD и не меняя DWG.
Public Function RunAudit03AutoCADConfigTests(Optional ByRef passed As Long = 0, _
        Optional ByRef failed As Long = 0) As String
    Dim stats As TCadStats, system As Object, shifted As Object, profiles As Object
    Dim saved As Variant, savedShifted As Variant, savedProfiles As Variant
    Dim originalName As String, prepared As Variant, position As Long
    On Error GoTo FailedRun
    Set system = ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange
    Set profiles = ThisWorkbook.Names.Item("rngCalculationProfiles").RefersToRange
    saved = system.Formula: savedProfiles = profiles.Formula
    originalName = ThisWorkbook.Names.Item("rngSystemSettings").RefersTo
    Set shifted = system.Worksheet.Range("CR800").Resize(system.Rows.Count, system.Columns.Count)
    savedShifted = shifted.Formula
    ConfigureFixture system
    PrepareResults stats, system, profiles
    prepared = system.Formula
    For position = 0 To 1
        If position = 1 Then
            shifted.Formula = prepared
            ThisWorkbook.Names.Item("rngSystemSettings").RefersTo = "='" & shifted.Worksheet.Name & "'!" & shifted.Address
            Set system = shifted
        End If
        CheckPresentation stats, system, position
        CheckImporter stats, system, position
        CheckCleanup stats, system, position
        CheckInvalidFields stats, system, position
        CheckLayerContracts stats, system, position
    Next position
    GoTo Restore
FailedRun:
    stats.Failed = stats.Failed + 1
    LogLine stats, "FAIL: autoCADConfig.runtime; " & CStr(Err.Number) & "; " & Err.Description
Restore:
    On Error GoTo FailedRestore
    If Len(originalName) > 0 Then ThisWorkbook.Names.Item("rngSystemSettings").RefersTo = originalName
    If Not IsEmpty(saved) Then ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange.Formula = saved
    If Not shifted Is Nothing Then shifted.Formula = savedShifted
    If Not profiles Is Nothing Then profiles.Formula = savedProfiles
    Check stats, "autoCADConfig.restore.name", ThisWorkbook.Names.Item("rngSystemSettings").RefersTo = originalName
    GoTo Finish
FailedRestore:
    stats.Failed = stats.Failed + 1
    LogLine stats, "FAIL: autoCADConfig.restore; " & Err.Description
Finish:
    passed = stats.Passed: failed = stats.Failed
    LogLine stats, "AUTOCAD_CONFIG_CASES: " & CStr(stats.Cases)
    LogLine stats, "TOTAL_AUTOCAD_CONFIG: passed=" & CStr(passed) & "; failed=" & CStr(failed)
    RunAudit03AutoCADConfigTests = stats.Report
End Function

' Разделяет слои геометрии, подписей и контура; цвета задаются индексами
' ACI 1..255, совместимыми с сущностями и вновь создаваемыми Layer.
Private Sub ConfigureFixture(ByVal system As Object)
    Dim keys As Variant, values As Variant, i As Long
    keys = CadKeys()
    values = Array("CAD_CFG", "Yes", "Transformed", "Yes", "Yes", "NamesAndValues", _
        "AUDIT_C", "AUDIT_R", "AUDIT_CONTOUR", "AUDIT_CT", "AUDIT_CC", "AUDIT_RT", "AUDIT_RC", _
        9, 5, 1, 6, 8, "AUDIT_IC", "AUDIT_IR")
    For i = 0 To UBound(keys): SetValue system, CStr(keys(i)), values(i): Next i
    SetValue system, "AutoCAD.Import.MinArea", 0#
    SetValue system, "General.NonCriticalMessagesEnabled", "No"
    SetValue system, "Calculation.ZeroMomentPerDepth", 0#
End Sub

' Сохраняет два настоящих прямых НДС с разными ID и напряжениями.
' Проверка выбора сочетания затем читает этот snapshot без нового solve.
Private Sub PrepareResults(ByRef stats As TCadStats, ByVal system As Object, ByVal profiles As Object)
    Dim settings As CSystemSettingsReader, units As CUnitSystem, section As CSectionModel
    Dim provider As CMaterialModelProvider, batch As CBatchSectionCalculator, writer As CNDMResultsWriter
    Dim key As Variant, x As Long, y As Long, elementIndex As Long
    For Each key In Array("Calculation.Stability.Enabled", "Calculation.Strength.Capacity", "Calculation.Crack.Width")
        SetProfile profiles, CStr(key), "No"
    Next key
    SetProfile profiles, "Calculation.Strength.DirectState", "Yes"
    SetProfile profiles, "Visualization.State", "StrengthState"
    SetProfile profiles, "Visualization.Quantity", "Stress"
    SetProfile profiles, "MaterialModel.Strength.ConcreteTension", "UseDiagram"
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    Set section = New CSectionModel
    For x = -1 To 1 Step 2
        For y = -1 To 1 Step 2
            elementIndex = elementIndex + 1
            section.AddConcreteElement 50# * x, 50# * y, 10000#, elementIndex, shapeType:="Rectangle", width:=100#, height:=100#
            section.AddRebarElement 60# * x, 60# * y, 16#, GEOM_PI * 64#, "A400", elementIndex
        Next y
    Next x
    section.Annotations.AddContourLine "CAD_BOTTOM", -100#, -100#, 100#, -100#
    section.Annotations.AddContourLine "CAD_RIGHT", 100#, -100#, 100#, 100#
    section.Annotations.AddContourLine "CAD_TOP", 100#, 100#, -100#, 100#
    section.Annotations.AddContourLine "CAD_LEFT", -100#, 100#, -100#, -100#
    Set provider = New CMaterialModelProvider: provider.Initialize settings, units
    Set batch = New CBatchSectionCalculator: batch.Initialize section, provider
    Dim catalog As CCalculationProfileCatalog
    Set catalog = New CCalculationProfileCatalog: catalog.LoadFromWorkbook ThisWorkbook
    Set batch.ProfileCatalog = catalog: batch.ApplySettings settings, units
    batch.AddCombination "CAD_CFG", 0#, 100000#, 50000#, "PR1", "AutoCAD Config fixture", "Auto"
    batch.AddCombination "CAD_CFG_2", 0#, 200000#, 100000#, "PR1", "AutoCAD Config fixture", "Auto"
    batch.Execute
    LogLine stats, "AUTOCAD_SNAPSHOT: " & batch.ResultAt(1).OverallMeta.ResultComment
    LogLine stats, "AUTOCAD_SNAPSHOT: " & batch.ResultAt(2).OverallMeta.ResultComment
    Check stats, "autoCADConfig.snapshot.state1", batch.ResultAt(1).DirectStateMeta.InternalStatus = rsSuccess
    Check stats, "autoCADConfig.snapshot.state2", batch.ResultAt(2).DirectStateMeta.InternalStatus = rsSuccess
    Set writer = New CNDMResultsWriter: writer.WriteResults ThisWorkbook, section, provider, batch, units
End Sub

' Сравнивает фактические подписи и выбор цвета/слоя всех физических состояний,
' а не только значения getters. Флаги требуют отдельного настоящего DWG gate.
Private Sub CheckPresentation(ByRef stats As TCadStats, ByVal system As Object, ByVal position As Long)
    Dim settings As CSystemSettingsReader, result As Object, material As Variant, physical As Variant
    Dim key As String, expectedLayer As String, expectedColor As Long, prefix As String
    Set settings = Reader()
    For Each material In Array("Concrete", "Rebar")
        For Each physical In Array("Tension", "Compression", "NearZero")
            Set result = Audit03AutoCADPresentationForTests(settings, CStr(material), CStr(physical), "E1", 12.3456, 2)
            stats.Cases = stats.Cases + 1
            prefix = "autoCADConfig.presentation." & CStr(position) & "." & CStr(material) & "." & CStr(physical)
            If CStr(material) = "Concrete" Then expectedLayer = "AUDIT_C" Else expectedLayer = "AUDIT_R"
            Check stats, prefix & ".geometry", CStr(result("GeometryLayer")) = expectedLayer
            key = CStr(material) & CStr(physical)
            If CStr(physical) = "NearZero" Then key = "Neutral"
            expectedColor = CLng(CellForKey(system, "AutoCAD.Color." & key, 2).Value2)
            Check stats, prefix & ".color", CLng(result("Color")) = expectedColor
            key = CStr(material) & "Compression"
            If CStr(physical) = "Tension" Then key = CStr(material) & "Tension"
            Check stats, prefix & ".annotation", CStr(result("AnnotationLayer")) = CStr(CellForKey(system, "AutoCAD.Layer." & key, 2).Value2)
            Check stats, prefix & ".label", CStr(result("Label")) = "E1 12.35"
        Next physical
    Next material
    Dim value As Variant, flag As Variant, saved As Variant, first As String, second As String, solves As Long
    saved = system.Formula: solves = SectionEquilibriumSolveCount()
    SetValue system, "AutoCAD.Export.LabelMode", "ValuesOnly"
    Set result = Audit03AutoCADPresentationForTests(Reader(), "Concrete", "Tension", "E1", 12.3456, 2)
    Check stats, "autoCADConfig.label.values." & CStr(position), CStr(result("Label")) = "12.35"
    For Each value In Array("None", "Concrete", "Transformed")
        SetValue system, "AutoCAD.Export.PrincipalAxesMode", CStr(value)
        Set result = Audit03AutoCADPresentationForTests(Reader(), "Rebar", "Compression", "R1", 1#, 0)
        Check stats, "autoCADConfig.axes." & CStr(position) & "." & CStr(value), CStr(result("AxesMode")) = CStr(value)
    Next value
    For Each flag In Array("NeutralLine", "LoadPoint", "Contour")
        For Each value In Array("No", "Yes")
            SetValue system, "AutoCAD.Export." & CStr(flag) & "Enabled", CStr(value)
            Set result = Audit03AutoCADPresentationForTests(Reader(), "Rebar", "Compression", "R1", 1#, 0)
            Check stats, "autoCADConfig.flag." & CStr(position) & "." & CStr(flag) & "." & CStr(value), CBool(result(CStr(flag))) = (CStr(value) = "Yes")
        Next value
    Next flag
    For Each value In Array(1, 255)
        SetValue system, "AutoCAD.Color.RebarTension", value
        Set result = Audit03AutoCADPresentationForTests(Reader(), "Rebar", "Tension", "R1", 1#, 0)
        Check stats, "autoCADConfig.color.bound." & CStr(position) & "." & CStr(value), CLng(result("Color")) = CLng(value)
    Next value
    SetValue system, "AutoCAD.Export.CombinationID", "CAD_CFG"
    first = Audit02ReadExportSnapshotForTests(ThisWorkbook)
    SetValue system, "AutoCAD.Export.CombinationID", "CAD_CFG_2"
    second = Audit02ReadExportSnapshotForTests(ThisWorkbook)
    Check stats, "autoCADConfig.combination.explicit." & CStr(position), Left$(first, 8) = "CAD_CFG|" And Left$(second, 10) = "CAD_CFG_2|"
    Check stats, "autoCADConfig.combination.changed." & CStr(position), first <> second
    Check stats, "autoCADConfig.combination.valuesChanged." & CStr(position), _
        Mid$(first, InStr(first, vbCrLf) + 2) <> Mid$(second, InStr(second, vbCrLf) + 2)
    SetValue system, "AutoCAD.Export.CombinationID", "Worst"
    first = Audit02ReadExportSnapshotForTests(ThisWorkbook)
    Check stats, "autoCADConfig.combination.worst." & CStr(position), Left$(first, 7) = "CAD_CFG"
    Check stats, "autoCADConfig.presentation.noSolve." & CStr(position), SectionEquilibriumSolveCount() = solves
    system.Formula = saved
End Sub

' Выбор каждого импортного слоя меняет реальный набор Region и координаты;
' геометрия источника остается в мм независимо от пользовательского INPUT.
Private Sub CheckImporter(ByRef stats As TCadStats, ByVal system As Object, ByVal position As Long)
    Dim model As CSectionModel, importer As CAutoCADSectionModelImporter, regions As Collection
    Dim settings As CSystemSettingsReader, prefix As String, saved As Variant
    saved = system.Formula: Set regions = ImportRegions(): Set importer = New CAutoCADSectionModelImporter
    Set settings = Reader(): Set model = importer.ImportConfiguredModelSpace(regions, settings)
    prefix = "autoCADConfig.import." & CStr(position)
    Check stats, prefix & ".counts", model.ConcreteCount = 1 And model.RebarCount = 1
    CheckClose stats, prefix & ".concreteMm", model.ConcreteX(1), 20#
    CheckClose stats, prefix & ".steelMm", model.RebarX(1), 80#
    SetValue system, "AutoCAD.Import.ConcreteLayer", "AUDIT_IC2"
    SetValue system, "AutoCAD.Import.RebarLayer", "AUDIT_IR2"
    Set model = importer.ImportConfiguredModelSpace(regions, Reader())
    CheckClose stats, prefix & ".changedConcrete", model.ConcreteX(1), 120#
    CheckClose stats, prefix & ".changedSteel", model.RebarX(1), 180#
    stats.Cases = stats.Cases + 2: system.Formula = saved
End Sub

' Рабочая очистка удаляет линии оформления, но сохраняет Region и чужие слои.
' Region на annotation-слое защищен независимо от настроек импортных слоев.
Private Sub CheckCleanup(ByRef stats As TCadStats, ByVal system As Object, ByVal position As Long)
    Dim ms As Object, region As CFakeAcadRegion, line As CFakeAcadLine, foreign As CFakeAcadLine, deleted As Long
    Set ms = CreateObject("Scripting.Dictionary")
    Set region = New CFakeAcadRegion: region.Initialize 100#, 0#, 0#, 100#, 100#, 0#, "AUDIT_CT", "KEEP_REGION"
    Set line = New CFakeAcadLine: line.Initialize 0#, 0#, 1#, 0#, "AUDIT_CT"
    Set foreign = New CFakeAcadLine: foreign.Initialize 0#, 0#, 1#, 0#, "FOREIGN"
    ms.Add 0, region: ms.Add 1, line: ms.Add 2, foreign
    deleted = Audit03CleanupAutoCADModelSpaceForTests(ms, Reader())
    Check stats, "autoCADConfig.cleanup.region." & CStr(position), Not region.Deleted
    Check stats, "autoCADConfig.cleanup.line." & CStr(position), line.Deleted
    Check stats, "autoCADConfig.cleanup.foreign." & CStr(position), Not foreign.Deleted
    Check stats, "autoCADConfig.cleanup.count." & CStr(position), deleted = 1
    stats.Cases = stats.Cases + 1
End Sub

' Невалидные поля проходят соответствующий consumer в двух положениях Config.
' Ошибка должна объяснять причину, текущую ячейку и требуемое исправление.
Private Sub CheckInvalidFields(ByRef stats As TCadStats, ByVal system As Object, ByVal position As Long)
    Dim key As Variant, values As Variant, value As Variant, target As Object, saved As Variant
    Dim reason As String, prefix As String, index As Long
    saved = system.Formula
    For Each key In CadKeys()
        values = Array(vbNullString, "TODO", "abc", CVErr(2015))
        If InStr(1, CStr(key), ".Layer.", vbTextCompare) > 0 Or InStr(1, CStr(key), "Layer", vbTextCompare) > 0 Then _
            values = Array(vbNullString, "TODO", "bad/name", "bad,name", String$(256, "a"), "bad" & vbLf & "name", CVErr(2015))
        If InStr(1, CStr(key), ".Color.", vbTextCompare) > 0 Then _
            values = Array(vbNullString, "TODO", "abc", CVErr(2015), -1, 0, 256, 1.5, 2147483648#)
        index = 0
        For Each value In values
            system.Formula = saved: Set target = CellForKey(system, CStr(key), 2): target.Value2 = value
            reason = ConsumeError(stats, CStr(key))
            prefix = "autoCADConfig.invalid." & CStr(position) & "." & CStr(key) & "." & CStr(index)
            Check stats, prefix & ".rejected", Len(reason) > 0
            Check stats, prefix & ".key", InStr(1, reason, CStr(key), vbTextCompare) > 0
            Check stats, prefix & ".address", InStr(1, reason, target.Address(False, False), vbTextCompare) > 0
            Check stats, prefix & ".action", InStr(1, reason, "введите", vbTextCompare) > 0 Or InStr(1, reason, "выберите", vbTextCompare) > 0 Or InStr(1, reason, "исправ", vbTextCompare) > 0
            LogLine stats, "AUTOCAD_INPUT: " & prefix & "|" & reason
            index = index + 1
        Next value
        system.Formula = saved: CellForKey(system, CStr(key), 1).Value2 = "Missing." & CStr(key)
        reason = ConsumeError(stats, CStr(key))
        Check stats, "autoCADConfig.missing." & CStr(position) & "." & CStr(key), Len(reason) > 0 And InStr(1, reason, CStr(key), vbTextCompare) > 0
        system.Formula = saved
        Check stats, "autoCADConfig.recovery." & CStr(position) & "." & CStr(key), Len(ConsumeError(stats, CStr(key))) = 0
    Next key
    system.Formula = saved
End Sub

' Проверяет межполевые конфликты, границы имени и неактивный contour-consumer.
' Очистка зависит только от слоев, поэтому поврежденный цвет ее не блокирует.
Private Sub CheckLayerContracts(ByRef stats As TCadStats, ByVal system As Object, ByVal position As Long)
    Dim saved As Variant, key As Variant, name As Variant, reason As String, target As Object, result As Object
    saved = system.Formula
    For Each key In Array("AutoCAD.Layer.Concrete", "AutoCAD.Layer.Rebar")
        For Each name In Array("AUDIT_CT", "AUDIT_CONTOUR", "RC_NDM_Axes", "RC_NDM_LoadPoint", "RC_NDM_NeutralLine", "RC_NDM_Warnings")
            Set target = CellForKey(system, CStr(key), 2): target.Value2 = CStr(name)
            reason = ConsumeError(stats, CStr(key))
            Check stats, "autoCADConfig.layerCollision." & CStr(position) & "." & CStr(key) & "." & CStr(name), _
                InStr(1, reason, CStr(key), vbTextCompare) > 0 And InStr(1, reason, target.Address(False, False), vbTextCompare) > 0
            system.Formula = saved
        Next name
    Next key
    SetValue system, "AutoCAD.Import.RebarLayer", "audit_ic"
    reason = ConsumeError(stats, "AutoCAD.Import.RebarLayer")
    Check stats, "autoCADConfig.import.sameLayer." & CStr(position), InStr(1, reason, "совпадают", vbTextCompare) > 0
    system.Formula = saved
    SetValue system, "AutoCAD.Export.ContourEnabled", "No"
    SetValue system, "AutoCAD.Layer.Contour", "bad/name"
    Check stats, "autoCADConfig.contour.inactive." & CStr(position), Len(ConsumeError(stats, "AutoCAD.Export.ContourEnabled")) = 0
    system.Formula = saved
    SetValue system, "AutoCAD.Color.RebarTension", "TODO"
    CheckCleanup stats, system, position
    system.Formula = saved
    For Each name In Array("Слой бетона длинный", String$(255, "a"))
        SetValue system, "AutoCAD.Layer.Concrete", CStr(name)
        Set result = Audit03AutoCADPresentationForTests(Reader(), "Concrete", "Tension", "C1", 1#, 2)
        Check stats, "autoCADConfig.layer.validBoundary." & CStr(position) & "." & CStr(Len(CStr(name))), result("GeometryLayer") = CStr(name)
        system.Formula = saved
    Next name
End Sub

' Не превращает ошибки Config в численные статусы: сохраняет рабочее описание.
' Snapshot и presentation являются штатными consumers, а не тестовым parser-ом.
Private Function ConsumeError(ByRef stats As TCadStats, ByVal key As String) As String
    Dim result As Object, model As CSectionModel, importer As CAutoCADSectionModelImporter, text As String
    stats.Cases = stats.Cases + 1
    On Error GoTo Rejected
    If Left$(key, 15) = "AutoCAD.Import." Then
        Set importer = New CAutoCADSectionModelImporter
        Set model = importer.ImportConfiguredModelSpace(ImportRegions(), Reader())
    ElseIf key = "AutoCAD.Export.CombinationID" Then
        text = Audit02ReadExportSnapshotForTests(ThisWorkbook)
    Else
        Set result = Audit03AutoCADPresentationForTests(Reader(), "Concrete", "Tension", "C1", 1#, 2)
    End If
    Exit Function
Rejected:
    ConsumeError = Err.Description
End Function

' Один список соответствует двадцати еще не принятым editable CAD-полям.
Private Function CadKeys() As Variant
    CadKeys = Array("AutoCAD.Export.CombinationID", "AutoCAD.Export.NeutralLineEnabled", _
        "AutoCAD.Export.PrincipalAxesMode", "AutoCAD.Export.LoadPointEnabled", "AutoCAD.Export.ContourEnabled", _
        "AutoCAD.Export.LabelMode", "AutoCAD.Layer.Concrete", "AutoCAD.Layer.Rebar", "AutoCAD.Layer.Contour", _
        "AutoCAD.Layer.ConcreteTension", "AutoCAD.Layer.ConcreteCompression", "AutoCAD.Layer.RebarTension", "AutoCAD.Layer.RebarCompression", _
        "AutoCAD.Color.ConcreteTension", "AutoCAD.Color.ConcreteCompression", "AutoCAD.Color.RebarTension", "AutoCAD.Color.RebarCompression", _
        "AutoCAD.Color.Neutral", "AutoCAD.Import.ConcreteLayer", "AutoCAD.Import.RebarLayer")
End Function

' Четыре реальные Region-fixtures позволяют отличить выбор каждого слоя.
Private Function ImportRegions() As Collection
    Dim result As Collection, region As CFakeAcadRegion, i As Long, layers As Variant
    Set result = New Collection: layers = Array("AUDIT_IC", "AUDIT_IR", "AUDIT_IC2", "AUDIT_IR2")
    For i = 0 To 3
        Set region = New CFakeAcadRegion
        If i Mod 2 = 0 Then
            region.Initialize 40000#, 20# + 50# * i, 30#, 133333333#, 133333333#, 0#, CStr(layers(i)), "C" & CStr(i)
        Else
            region.Initialize 201#, 30# + 50# * i, 90#, 3215#, 3215#, 0#, CStr(layers(i)), "R" & CStr(i)
        End If
        result.Add region
    Next i
    Set ImportRegions = result
End Function

' Читает текущую привязку Config; перенесенная таблица не кешируется.
Private Function Reader() As CSystemSettingsReader
    Set Reader = New CSystemSettingsReader: Reader.LoadFromWorkbook ThisWorkbook
End Function

' Ищет ключ и возвращает настоящую ячейку для изменения и проверки адреса.
Private Function CellForKey(ByVal table As Object, ByVal key As String, ByVal column As Long) As Object
    Dim row As Long
    For row = 2 To table.Rows.Count
        If CStr(table.Cells(row, 1).Value2) = key Then Set CellForKey = table.Cells(row, column): Exit Function
    Next row
    Err.Raise vbObjectError + 4499, "CellForKey", "В fixture отсутствует Config " & key
End Function

' Меняет значение без изменения самого ключа или структуры таблицы.
Private Sub SetValue(ByVal table As Object, ByVal key As String, ByVal value As Variant)
    CellForKey(table, key, 2).Value2 = value
End Sub

' Ищет PR1 в шапке таблицы профилей, не предполагая фиксированный столбец.
Private Sub SetProfile(ByVal table As Object, ByVal key As String, ByVal value As String)
    Dim row As Long, column As Long, targetRow As Long, targetColumn As Long
    For row = 1 To table.Rows.Count
        If CStr(table.Cells(row, 2).Value2) = key Then targetRow = row
        For column = 3 To table.Columns.Count
            If CStr(table.Cells(row, column).Value2) = "PR1" Then targetColumn = column
        Next column
    Next row
    If targetRow = 0 Or targetColumn = 0 Then Err.Raise vbObjectError + 4499, "SetProfile", "Не найден PR1 " & key
    table.Cells(targetRow, targetColumn).Value2 = value
End Sub

' Накапливает все assertions, чтобы отрицательный baseline не обрывал набор.
Private Sub Check(ByRef stats As TCadStats, ByVal name As String, ByVal condition As Boolean)
    If condition Then
        stats.Passed = stats.Passed + 1: LogLine stats, "OK: " & name
    Else
        stats.Failed = stats.Failed + 1: LogLine stats, "FAIL: " & name
    End If
End Sub

' Сравнивает точные координаты fixtures в мм и сохраняет исходные числа.
Private Sub CheckClose(ByRef stats As TCadStats, ByVal name As String, ByVal actual As Double, ByVal expected As Double)
    LogLine stats, "AUTOCAD_NUMBER: " & name & "|actual=" & CStr(actual) & "|expected=" & CStr(expected) & "|tolerance=0.00000001"
    Check stats, name, Abs(actual - expected) <= 0.00000001
End Sub

' Возвращаемый отчет не зависит от доступности дополнительного файла progress.
Private Sub LogLine(ByRef stats As TCadStats, ByVal line As String)
    stats.Report = stats.Report & line & vbCrLf
End Sub

' Создает только собственный новый DWG в разрешенном каталоге; все проверки
' проходят реальный CAD/экспорт/import. Чужие открытые документы запрещены
' ранним guard-ом. HWND связывает тест с явно запущенным пользователем CAD.
Public Function RunAudit03RealAutoCADTests(ByVal drawingPath As String, ByVal expectedHwnd As String) As String
    Dim stats As TCadStats, system As Object, profiles As Object, unitsRange As Object
    Dim saved As Variant, savedProfiles As Variant, savedUnits As Variant, savedSigns As Variant, acad As Object, doc As Object, signs As Object
    Dim value As Variant, key As Variant, added As Long, solves As Long, layer As Object, copyRegion As Object
    On Error GoTo FailedRun
    Set acad = ConnectToRunningAutoCAD()
    If CStr(acad.HWND) <> expectedHwnd Then Err.Raise vbObjectError + 4499, "RunAudit03RealAutoCADTests", "Изменилась идентичность тестового AutoCAD."
    If acad.Documents.Count <> 0 Then Err.Raise vbObjectError + 4499, "RunAudit03RealAutoCADTests", "В AutoCAD открыт чужой или предыдущий документ; тест его не изменяет."
    If InStr(1, drawingPath, "\docs\regression\Audit03\", vbTextCompare) = 0 Then Err.Raise vbObjectError + 4499, "RunAudit03RealAutoCADTests", "DWG должен находиться в каталоге Audit03."
    If Len(Dir$(drawingPath)) > 0 Then Err.Raise vbObjectError + 4499, "RunAudit03RealAutoCADTests", "Тестовый DWG уже существует; перезапись запрещена."
    Set system = ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange
    Set profiles = ThisWorkbook.Names.Item("rngCalculationProfiles").RefersToRange
    Set unitsRange = ThisWorkbook.Names.Item("rngUnitSettings").RefersToRange
    Set signs = ThisWorkbook.Names.Item("rngSignConventionSettings").RefersToRange
    saved = system.Formula: savedProfiles = profiles.Formula: savedUnits = unitsRange.Formula: savedSigns = signs.Formula
    ConfigureFixture system: PrepareResults stats, system, profiles
    LogLine stats, "REAL_SNAPSHOT: " & Audit02ReadExportSnapshotForTests(ThisWorkbook)
    solves = SectionEquilibriumSolveCount()
    Set doc = acad.Documents.Add()
    WaitForCAD acad
    LogLine stats, "REAL_AUTOCAD: " & acad.FullName & "|version=" & acad.Version & "|HWND=" & expectedHwnd
    Set layer = doc.Layers.Add("AUDIT_C"): layer.Color = 7
    Audit03EnsureAutoCADLayerForTests doc, "AUDIT_C", 1
    Check stats, "realAutoCAD.layer.existingColor", layer.Color = 7
    Audit03EnsureAutoCADLayerForTests doc, "AUDIT_ACI1", 1
    Audit03EnsureAutoCADLayerForTests doc, "AUDIT_ACI255", 255
    Check stats, "realAutoCAD.layer.color1", doc.Layers.Item("AUDIT_ACI1").Color = 1
    Check stats, "realAutoCAD.layer.color255", doc.Layers.Item("AUDIT_ACI255").Color = 255
    Dim aci As Long
    Set layer = doc.Layers.Item("AUDIT_ACI1")
    For aci = 1 To 255
        layer.Color = aci
        Check stats, "realAutoCAD.layer.fullACI." & CStr(aci), layer.Color = aci
    Next aci
    Audit03EnsureAutoCADLayerForTests doc, "Слой бетона длинный", 4
    Audit03EnsureAutoCADLayerForTests doc, String$(255, "a"), 4
    Check stats, "realAutoCAD.layer.unicode", doc.Layers.Item("Слой бетона длинный").Name = "Слой бетона длинный"
    Check stats, "realAutoCAD.layer.length255", Len(doc.Layers.Item(String$(255, "a")).Name) = 255
    added = Audit03ExportAutoCADDocumentForTests(doc)
    Check stats, "realAutoCAD.export.added", added > 16
    Check stats, "realAutoCAD.export.regions", CountEntities(doc, "AcDbRegion") = 8
    Check stats, "realAutoCAD.export.concrete", CountEntities(doc, "AcDbRegion", "AUDIT_C") = 4
    Check stats, "realAutoCAD.export.rebar", CountEntities(doc, "AcDbRegion", "AUDIT_R") = 4
    Check stats, "realAutoCAD.export.axes", CountEntities(doc, "", "RC_NDM_Axes") > 0
    Check stats, "realAutoCAD.export.loadPoint", CountEntities(doc, "", "RC_NDM_LoadPoint") > 0
    Check stats, "realAutoCAD.export.neutralLine", CountEntities(doc, "", "RC_NDM_NeutralLine") > 0
    Check stats, "realAutoCAD.export.contour", CountEntities(doc, "", "AUDIT_CONTOUR") > 0
    CheckNativeColors stats, doc, system
    CheckNativeImport stats, doc, system
    For Each key In Array("NeutralLine", "LoadPoint", "Contour", "PrincipalAxes")
        ClearOwnModelSpace doc
        If CStr(key) = "PrincipalAxes" Then
            SetValue system, "AutoCAD.Export.PrincipalAxesMode", "None"
        Else
            SetValue system, "AutoCAD.Export." & CStr(key) & "Enabled", "No"
        End If
        added = Audit03ExportAutoCADDocumentForTests(doc)
        Dim disabledLayer As String
        Select Case CStr(key)
            Case "Contour": disabledLayer = "AUDIT_CONTOUR"
            Case "PrincipalAxes": disabledLayer = "RC_NDM_Axes"
            Case Else: disabledLayer = "RC_NDM_" & CStr(key)
        End Select
        Check stats, "realAutoCAD.disabled." & CStr(key), CountEntities(doc, "", disabledLayer) = 0
        Check stats, "realAutoCAD.disabled.geometry." & CStr(key), CountEntities(doc, "AcDbRegion") = 8
        ConfigureFixture system
    Next key
    For Each value In Array("ValuesOnly", "NamesAndValues")
        ClearOwnModelSpace doc: SetValue system, "AutoCAD.Export.LabelMode", CStr(value)
        added = Audit03ExportAutoCADDocumentForTests(doc)
        CheckNativeLabels stats, doc, CStr(value)
    Next value
    ' Меняем только текущий Config после сохранения Results: размеры DWG
    ' должны остаться 100/200 мм, а не превратиться в метры или миллиметры повторно.
    SetValue unitsRange, "Length", "m": CellForKey(unitsRange, "Length", 4).Value2 = "m"
    SetValue unitsRange, "Area", "m2": CellForKey(unitsRange, "Area", 4).Value2 = "m2"
    SetValue signs, "+N", "Tension": SetValue signs, "+Mx", "-Y tension": SetValue signs, "+My", "-X tension"
    ClearOwnModelSpace doc: added = Audit03ExportAutoCADDocumentForTests(doc)
    CheckNativeImport stats, doc, system
    CheckNativeColors stats, doc, system
    Dim presentationPath As String
    presentationPath = Left$(drawingPath, Len(drawingPath) - 4) & "_presentation.dwg"
    If Len(Dir$(presentationPath)) > 0 Then Err.Raise vbObjectError + 4499, "RunAudit03RealAutoCADTests", "Тестовый presentation-DWG уже существует; перезапись запрещена."
    doc.SaveAs presentationPath
    LogLine stats, "REAL_DWG_PRESENTATION: " & doc.FullName & "|entities=" & CStr(doc.ModelSpace.Count)
    Dim entity As Object
    For Each entity In doc.ModelSpace
        If entity.ObjectName = "AcDbRegion" Then
            Set copyRegion = entity.Copy(): copyRegion.Layer = "AUDIT_CT": Exit For
        End If
    Next entity
    Dim deleted As Long: deleted = Audit03CleanupAutoCADModelSpaceForTests(doc.ModelSpace, Reader())
    Check stats, "realAutoCAD.cleanup.deleted", deleted > 0
    Check stats, "realAutoCAD.cleanup.allRegionsKept", CountEntities(doc, "AcDbRegion") = 9
    Check stats, "realAutoCAD.cleanup.regionOnAnnotation", CountEntities(doc, "AcDbRegion", "AUDIT_CT") = 1
    Check stats, "realAutoCAD.export.noSolve", SectionEquilibriumSolveCount() = solves
    doc.SaveAs drawingPath
    Check stats, "realAutoCAD.saved", Len(Dir$(drawingPath)) > 0
    LogLine stats, "REAL_DWG: " & doc.FullName & "|entities=" & CStr(doc.ModelSpace.Count)
    doc.Close False: Set doc = Nothing
    Set doc = acad.Documents.Open(drawingPath): WaitForCAD acad
    Check stats, "realAutoCAD.reopen.regions", CountEntities(doc, "AcDbRegion") = 9
    CheckNativeImport stats, doc, system
    GoTo Restore
FailedRun:
    stats.Failed = stats.Failed + 1
    LogLine stats, "FAIL: realAutoCAD.runtime; " & CStr(Err.Number) & "; " & Err.Description
Restore:
    On Error GoTo FailedRestore
    If Not doc Is Nothing Then WaitForCAD acad: doc.Close False
    If Not system Is Nothing Then system.Formula = saved
    If Not profiles Is Nothing Then profiles.Formula = savedProfiles
    If Not unitsRange Is Nothing Then unitsRange.Formula = savedUnits
    If Not signs Is Nothing Then signs.Formula = savedSigns
    GoTo Finish
FailedRestore:
    stats.Failed = stats.Failed + 1: LogLine stats, "FAIL: realAutoCAD.restore; " & Err.Description
Finish:
    LogLine stats, "TOTAL_REAL_AUTOCAD: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunAudit03RealAutoCADTests = stats.Report
End Function

' Ждет завершения CAD-инициализации нового документа ограниченное время.
' Busy не становится ложным отказом геометрии; после timeout тест прекращается.
Private Sub WaitForCAD(ByVal acad As Object)
    Dim started As Single, elapsed As Double: started = Timer
    Do
        If acad.GetAcadState().IsQuiescent Then Exit Sub
        DoEvents: Application.Wait DateAdd("s", 1, Now)
        elapsed = Timer - started: If elapsed < 0# Then elapsed = elapsed + 86400#
        If elapsed > 30# Then Err.Raise vbObjectError + 4499, "WaitForCAD", "AutoCAD не завершил текущую команду за 30 секунд."
    Loop
End Sub

' Считает настоящие сущности по COM ObjectName/Layer, не по flags Config.
Private Function CountEntities(ByVal doc As Object, Optional ByVal objectName As String = "", _
        Optional ByVal layerName As String = "") As Long
    Dim entity As Object
    For Each entity In doc.ModelSpace
        If (Len(objectName) = 0 Or StrComp(entity.ObjectName, objectName, vbTextCompare) = 0) And _
                (Len(layerName) = 0 Or StrComp(entity.Layer, layerName, vbTextCompare) = 0) Then CountEntities = CountEntities + 1
    Next entity
End Function

' Удаляет только содержимое собственного нового test-document между сценариями;
' этот helper не вызывается ни для пользовательского ActiveDocument, ни из UI.
Private Sub ClearOwnModelSpace(ByVal doc As Object)
    Dim i As Long
    For i = doc.ModelSpace.Count - 1 To 0 Step -1: doc.ModelSpace.Item(i).Delete: Next i
End Sub

' Читает экспортированные настоящие Region рабочим importer-ом. Контролирует
' точные мм/мм2 и раскладку вместо проверки только существования DWG-файла.
Private Sub CheckNativeImport(ByRef stats As TCadStats, ByVal doc As Object, ByVal system As Object)
    SetValue system, "AutoCAD.Import.ConcreteLayer", "AUDIT_C"
    SetValue system, "AutoCAD.Import.RebarLayer", "AUDIT_R"
    Dim importer As CAutoCADSectionModelImporter, model As CSectionModel, i As Long
    Set importer = New CAutoCADSectionModelImporter
    Set model = importer.ImportConfiguredModelSpace(doc.ModelSpace, Reader())
    Check stats, "realAutoCAD.import.counts", model.ConcreteCount = 4 And model.RebarCount = 4
    For i = 1 To model.ConcreteCount
        CheckClose stats, "realAutoCAD.import.concreteArea." & CStr(i), model.ConcreteArea(i), 10000#
        CheckClose stats, "realAutoCAD.import.concreteX." & CStr(i), Abs(model.ConcreteX(i)), 50#
        CheckClose stats, "realAutoCAD.import.concreteY." & CStr(i), Abs(model.ConcreteY(i)), 50#
    Next i
    For i = 1 To model.RebarCount
        CheckClose stats, "realAutoCAD.import.rebarArea." & CStr(i), model.RebarArea(i), GEOM_PI * 64#
        CheckClose stats, "realAutoCAD.import.rebarX." & CStr(i), Abs(model.RebarX(i)), 60#
        CheckClose stats, "realAutoCAD.import.rebarY." & CStr(i), Abs(model.RebarY(i)), 60#
    Next i
End Sub

' Проверяет текст именно экспортных подписей элементов; осевые подписи
' и маркеры нагрузки не включаются в количество восемь расчетных labels.
Private Sub CheckNativeLabels(ByRef stats As TCadStats, ByVal doc As Object, ByVal mode As String)
    Dim entity As Object, count As Long, text As String, hasName As Boolean
    For Each entity In doc.ModelSpace
        If entity.ObjectName = "AcDbText" And Left$(entity.Layer, 6) = "AUDIT_" Then
            text = entity.TextString: count = count + 1
            hasName = (Left$(text, 1) = "C" Or Left$(text, 1) = "R")
            Check stats, "realAutoCAD.label." & mode & "." & CStr(count), hasName = (mode = "NamesAndValues")
        End If
    Next entity
    Check stats, "realAutoCAD.label.count." & mode, count = 8
End Sub

' Независимо сопоставляет цвет настоящего Region/текста с PhysicalState
' сохраненного snapshot и текущими пятью ACI. Измененные знаки N/Mx/My
' не должны перекрасить физическое растяжение в сжатие.
Private Sub CheckNativeColors(ByRef stats As TCadStats, ByVal doc As Object, ByVal system As Object)
    Dim states As Object, lines As Variant, parts As Variant, i As Long, entity As Object
    Dim elementID As String, physical As String, material As String, colorKey As String, center As Variant, index As Long
    Set states = CreateObject("Scripting.Dictionary"): states.CompareMode = vbTextCompare
    lines = Split(Audit02ReadExportSnapshotForTests(ThisWorkbook), vbCrLf)
    For i = 1 To UBound(lines)
        parts = Split(CStr(lines(i)), "|")
        If UBound(parts) = 2 Then states.Add CStr(parts(0)), CStr(parts(2))
    Next i
    For Each entity In doc.ModelSpace
        elementID = vbNullString
        If entity.ObjectName = "AcDbRegion" And (entity.Layer = "AUDIT_C" Or entity.Layer = "AUDIT_R") Then
            center = entity.Centroid: index = 1
            If center(LBound(center)) > 0# Then index = index + 2
            If center(LBound(center) + 1) > 0# Then index = index + 1
            If entity.Layer = "AUDIT_C" Then elementID = "C" & CStr(index) Else elementID = "R" & CStr(index)
        ElseIf entity.ObjectName = "AcDbText" And Left$(entity.Layer, 6) = "AUDIT_" Then
            parts = Split(CStr(entity.TextString), " ")
            If states.Exists(CStr(parts(0))) Then elementID = CStr(parts(0))
        End If
        If Len(elementID) > 0 Then
            physical = CStr(states(elementID)): material = "Concrete"
            If Left$(elementID, 1) = "R" Then material = "Rebar"
            Select Case physical
                Case "Compression", "Tension": colorKey = material & physical
                Case Else: colorKey = "Neutral"
            End Select
            Check stats, "realAutoCAD.color." & elementID & "." & entity.ObjectName, _
                entity.Color = CLng(CellForKey(system, "AutoCAD.Color." & colorKey, 2).Value2)
        End If
    Next entity
End Sub
