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
    ProgressPath As String ' Только native CAD: журнал до завершения COM-вызова.
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
        CheckSettingEffects stats, system, position
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
    Set model = importer.ImportConfiguredModelSpace(regions, Reader())
    CheckClose stats, prefix & ".changedConcrete", model.ConcreteX(1), 120#
    CheckClose stats, prefix & ".concreteLayerKeepsSteel", model.RebarX(1), 80#
    system.Formula = saved
    SetValue system, "AutoCAD.Import.RebarLayer", "AUDIT_IR2"
    Set model = importer.ImportConfiguredModelSpace(regions, Reader())
    CheckClose stats, prefix & ".changedSteel", model.RebarX(1), 180#
    CheckClose stats, prefix & ".steelLayerKeepsConcrete", model.ConcreteX(1), 20#
    stats.Cases = stats.Cases + 3: system.Formula = saved
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

' Меняет каждый экспортный слой и цвет независимо. Ожидается изменение
' именно предназначенного поля оформления; чувствительность не доказывается
' одним чтением default или сравнением полного snapshot с другим ID.
Private Sub CheckSettingEffects(ByRef stats As TCadStats, ByVal system As Object, ByVal position As Long)
    Dim saved As Variant, result As Object, keys As Variant, materials As Variant
    Dim states As Variant, fields As Variant, i As Long, name As String
    saved = system.Formula
    keys = Array("Concrete", "Rebar", "Contour", "ConcreteTension", _
        "ConcreteCompression", "RebarTension", "RebarCompression")
    materials = Array("Concrete", "Rebar", "Concrete", "Concrete", "Concrete", "Rebar", "Rebar")
    states = Array("Tension", "Tension", "Tension", "Tension", "Compression", "Tension", "Compression")
    fields = Array("GeometryLayer", "GeometryLayer", "ContourLayer", _
        "AnnotationLayer", "AnnotationLayer", "AnnotationLayer", "AnnotationLayer")
    For i = 0 To UBound(keys)
        name = "AUDIT_CHANGED_" & CStr(position) & "_" & CStr(i)
        SetValue system, "AutoCAD.Layer." & CStr(keys(i)), name
        Set result = Audit03AutoCADPresentationForTests(Reader(), CStr(materials(i)), CStr(states(i)), "E1", 1#, 2)
        stats.Cases = stats.Cases + 1
        Check stats, "autoCADConfig.effect." & CStr(position) & ".AutoCAD.Layer." & CStr(keys(i)), _
            CStr(result(CStr(fields(i)))) = name
        system.Formula = saved
    Next i
    keys = Array("ConcreteTension", "ConcreteCompression", "RebarTension", "RebarCompression", "Neutral")
    materials = Array("Concrete", "Concrete", "Rebar", "Rebar", "Concrete")
    states = Array("Tension", "Compression", "Tension", "Compression", "NearZero")
    Dim color As Variant
    For i = 0 To UBound(keys)
        For Each color In Array(1, 255)
            SetValue system, "AutoCAD.Color." & CStr(keys(i)), color
            Set result = Audit03AutoCADPresentationForTests(Reader(), CStr(materials(i)), CStr(states(i)), "E1", 1#, 2)
            stats.Cases = stats.Cases + 1
            Check stats, "autoCADConfig.effect." & CStr(position) & ".AutoCAD.Color." & CStr(keys(i)) & "." & CStr(color), _
                CLng(result("Color")) = CLng(color)
            system.Formula = saved
        Next color
    Next i
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
    If Len(stats.ProgressPath) = 0 Then Exit Sub
    Dim fileNumber As Integer
    On Error GoTo ProgressUnavailable
    fileNumber = FreeFile
    Open stats.ProgressPath For Append As #fileNumber
    Print #fileNumber, line
    Close #fileNumber
    Exit Sub
ProgressUnavailable:
    On Error Resume Next
    If fileNumber > 0 Then Close #fileNumber
    On Error GoTo 0
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
    stats.ProgressPath = drawingPath & ".progress.txt"
    LogLine stats, "REAL_STAGE: prepare workbook snapshot"
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
    LogLine stats, "REAL_STAGE: reopen saved DWG"
    Set doc = acad.Documents.Open(drawingPath)
    LogLine stats, "REAL_STAGE: wait for reopened document"
    WaitForCAD acad
    LogLine stats, "REAL_STAGE: verify reopened regions"
    Check stats, "realAutoCAD.reopen.regions", CountEntities(doc, "AcDbRegion") = 9
    CheckNativeImport stats, doc, system
    LogLine stats, "REAL_STAGE: independent central inertia probes"
    CheckNativeRegionInertia stats, doc
    LogLine stats, "REAL_STAGE: imported shape -> Results -> exported Region"
    CheckNativeShapeRoundTrip stats, doc, system, unitsRange, drawingPath
    LogLine stats, "REAL_STAGE: square edge and circular Region average orientation"
    CheckNativeIsotropicRoundTrip stats, doc, system, unitsRange, drawingPath
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

' Проверяет центральные A/I настоящего повернутого Region независимо от
' выбора кандидатов importer-а. Перенос по X/Y не меняет центральный тензор;
' оба знака угла и произведения координат выявляют ошибочную эвристику Ixy.
Private Sub CheckNativeRegionInertia(ByRef stats As TCadStats, ByVal doc As Object)
    Dim angleSign As Long, position As Long, angle As Double, cx As Double, cy As Double
    Dim expectedIx As Double, expectedIy As Double, expectedIxy As Double
    Dim ix As Double, iy As Double, ixy As Double, model As CSectionModel
    Dim importer As CAutoCADSectionModelImporter, region As Object, layerName As String
    Dim raw As Variant, rawProduct As Double, prefix As String
    Set importer = New CAutoCADSectionModelImporter
    For angleSign = -1 To 1 Step 2
        angle = CDbl(angleSign) * GEOM_PI / 6#
        expectedIx = 2160000# * Cos(angle) ^ 2 + 8640000# * Sin(angle) ^ 2
        expectedIy = 8640000# * Cos(angle) ^ 2 + 2160000# * Sin(angle) ^ 2
        expectedIxy = (8640000# - 2160000#) * Sin(angle) * Cos(angle)
        For position = 0 To 2
            cx = 0#: cy = 0#
            If position > 0 Then cx = 20#: cy = -10#
            If position = 2 Then cy = 10#
            layerName = "AUDIT_INERTIA_" & CStr(angleSign + 1) & "_" & CStr(position)
            LogLine stats, "REAL_STAGE: create " & layerName
            Audit03EnsureAutoCADLayerForTests doc, layerName, 7
            Set region = NativeRectangleRegion(doc, 120#, 60#, angle, cx, cy)
            region.Layer = layerName
            raw = region.ProductOfInertia
            If IsArray(raw) Then rawProduct = CDbl(raw(LBound(raw))) Else rawProduct = CDbl(raw)
            prefix = "realAutoCAD.inertia." & CStr(angleSign) & "." & CStr(position)
            LogLine stats, "REAL_INERTIA: " & prefix & "|cx=" & CStr(cx) & "; cy=" & CStr(cy) & _
                "; rawProduct=" & NativeNumber(rawProduct) & "; expectedCentralIxy=" & NativeNumber(expectedIxy)
            Set model = importer.ImportFromModelSpace(doc.ModelSpace, layerName, "AUDIT_R", "A400", 0#)
            Check stats, prefix & ".count", model.ConcreteCount = 1
            If model.ConcreteCount = 1 Then
                model.ConcreteLocalInertiaComponents 1, ix, iy, ixy
                CheckNativeInertiaClose stats, prefix & ".area", model.ConcreteArea(1), 7200#
                CheckNativeInertiaClose stats, prefix & ".x", model.ConcreteX(1), cx
                CheckNativeInertiaClose stats, prefix & ".y", model.ConcreteY(1), cy
                CheckNativeInertiaClose stats, prefix & ".Ix", ix, expectedIx
                CheckNativeInertiaClose stats, prefix & ".Iy", iy, expectedIy
                CheckNativeInertiaClose stats, prefix & ".Ixy", ixy, expectedIxy
            End If
        Next position
    Next angleSign
End Sub

' Создает тестовый прямоугольник по независимо повернутым четырем вершинам.
' AddRegion получает замкнутый Polyline; временная исходная линия удаляется
' только из собственного документа теста после создания Region.
Private Function NativeRectangleRegion(ByVal doc As Object, ByVal width As Double, ByVal height As Double, _
        ByVal angle As Double, ByVal cx As Double, ByVal cy As Double) As Object
    Dim points(0 To 7) As Double, px As Variant, py As Variant, i As Long
    px = Array(-width / 2#, width / 2#, width / 2#, -width / 2#)
    py = Array(-height / 2#, -height / 2#, height / 2#, height / 2#)
    For i = 0 To 3
        points(2 * i) = cx + CDbl(px(i)) * Cos(angle) - CDbl(py(i)) * Sin(angle)
        points(2 * i + 1) = cy + CDbl(px(i)) * Sin(angle) + CDbl(py(i)) * Cos(angle)
    Next i
    Dim line As Object, curves(0 To 0) As Object, regions As Variant
    Set line = doc.ModelSpace.AddLightWeightPolyline(points)
    line.Closed = True
    Set curves(0) = line
    regions = doc.ModelSpace.AddRegion(curves)
    Set NativeRectangleRegion = regions(LBound(regions))
    line.Delete
End Function

' Отделяет численную погрешность CAD-интеграции от инженерного расхождения.
' Абсолютный допуск 1e-6 фиксирован до запуска и не меняет старые oracle ядра.
Private Sub CheckNativeInertiaClose(ByRef stats As TCadStats, ByVal name As String, ByVal actual As Double, ByVal expected As Double)
    LogLine stats, "AUTOCAD_INERTIA_NUMBER: " & name & "|actual=" & NativeNumber(actual) & _
        "|expected=" & NativeNumber(expected) & "|tolerance=0.000001"
    Check stats, name, Abs(actual - expected) <= 0.000001
End Sub

' Сохраняет значение CAD в журнале без зависимости от десятичного разделителя
' Excel. Этот helper только форматирует диагностику, не округляет oracle.
Private Function NativeNumber(ByVal value As Double) As String
    NativeNumber = Replace$(CStr(value), ",", ".")
End Function

' Проверяет весь рабочий snapshot/export путь для повернутого прямоугольника,
' треугольника и Г-образного Region. Oracle интегрирует исходный polygon,
' а не использует восстановленные стороны CSectionModel как собственный эталон.
' Обе ориентации повторяются при OUTPUT мм/м; все DWG принадлежат этому тесту.
Private Sub CheckNativeShapeRoundTrip(ByRef stats As TCadStats, ByVal doc As Object, _
        ByVal system As Object, ByVal unitsRange As Object, ByVal drawingPath As String)
    Dim shape As Long, signValue As Long, unitMode As Long, i As Long, j As Long
    Dim px As Variant, py As Variant, x() As Double, y() As Double, angle As Double
    Dim region As Object, bar As Object, entity As Object, exported As Object
    Dim importer As CAutoCADSectionModelImporter, model As CSectionModel
    Dim area As Double, cx As Double, cy As Double, ix As Double, iy As Double, ixy As Double
    Dim actualArea As Double, actualX As Double, actualY As Double
    Dim actualIx As Double, actualIy As Double, actualIxy As Double
    Dim prefix As String, path As String, solves As Long, added As Long
    Set importer = New CAutoCADSectionModelImporter
    ConfigureFixture system
    SetValue system, "AutoCAD.Export.PrincipalAxesMode", "None"
    SetValue system, "AutoCAD.Export.NeutralLineEnabled", "No"
    SetValue system, "AutoCAD.Export.LoadPointEnabled", "No"
    SetValue system, "AutoCAD.Export.ContourEnabled", "No"
    For unitMode = 0 To 1
        If unitMode = 0 Then
            SetValue unitsRange, "Length", "mm": CellForKey(unitsRange, "Length", 4).Value2 = "mm"
            SetValue unitsRange, "Area", "mm2": CellForKey(unitsRange, "Area", 4).Value2 = "mm2"
        Else
            SetValue unitsRange, "Length", "m": CellForKey(unitsRange, "Length", 4).Value2 = "m"
            SetValue unitsRange, "Area", "m2": CellForKey(unitsRange, "Area", 4).Value2 = "m2"
        End If
        For shape = 0 To 2
            Select Case shape
                Case 0
                    px = Array(-60#, 60#, 60#, -60#): py = Array(-30#, -30#, 30#, 30#)
                Case 1
                    px = Array(-60#, 60#, 0#): py = Array(-30#, -30#, 30#)
                Case 2
                    px = Array(-60#, 60#, 60#, 0#, 0#, -60#)
                    py = Array(-40#, -40#, -10#, -10#, 40#, 40#)
            End Select
            For signValue = -1 To 1 Step 2
                prefix = "realAutoCAD.shape." & CStr(unitMode) & "." & CStr(shape) & "." & CStr(signValue)
                LogLine stats, "REAL_STAGE: " & prefix
                ClearOwnModelSpace doc
                angle = CDbl(signValue) * GEOM_PI / 6#
                ReDim x(0 To UBound(px)): ReDim y(0 To UBound(py))
                For i = 0 To UBound(px)
                    x(i) = 20# + CDbl(px(i)) * Cos(angle) - CDbl(py(i)) * Sin(angle)
                    y(i) = -10# + CDbl(px(i)) * Sin(angle) + CDbl(py(i)) * Cos(angle)
                Next i
                NativePolygonInertia x, y, area, cx, cy, ix, iy, ixy
                Audit03EnsureAutoCADLayerForTests doc, "AUDIT_SHAPE_IN", 7
                Audit03EnsureAutoCADLayerForTests doc, "AUDIT_SHAPE_R", 7
                Set region = NativePolygonRegion(doc, x, y): region.Layer = "AUDIT_SHAPE_IN"
                For i = -1 To 1 Step 2
                    For j = 0 To 1
                        actualX = 20# + 5# * i * Cos(angle) - (-25# + 10# * j) * Sin(angle)
                        actualY = -10# + 5# * i * Sin(angle) + (-25# + 10# * j) * Cos(angle)
                        Set bar = NativeRectangleRegion(doc, 2#, 2#, angle, actualX, actualY)
                        bar.Layer = "AUDIT_SHAPE_R"
                    Next j
                Next i
                Set model = importer.ImportFromModelSpace(doc.ModelSpace, "AUDIT_SHAPE_IN", "AUDIT_SHAPE_R", "A400", 0#)
                CheckNativeInertiaClose stats, prefix & ".import.area", model.ConcreteArea(1), area
                CheckNativeInertiaClose stats, prefix & ".import.Ix", model.ConcreteLocalIx(1), ix
                CheckNativeInertiaClose stats, prefix & ".import.Iy", model.ConcreteLocalIy(1), iy
                CheckNativeInertiaClose stats, prefix & ".import.Ixy", model.ConcreteLocalIxy(1), ixy
                If shape = 0 Then
                    Check stats, prefix & ".sourceRectangle", model.ConcreteShapeType(1) = "Rectangle"
                Else
                    Check stats, prefix & ".sourceEquivalent", model.ConcreteShapeType(1) = "Equivalent rectangle"
                End If
                PrepareNativeImportedSnapshot stats, system, model, prefix
                solves = SectionEquilibriumSolveCount()
                added = Audit03ExportAutoCADDocumentForTests(doc)
                Check stats, prefix & ".export.noSolve", SectionEquilibriumSolveCount() = solves
                Check stats, prefix & ".export.added", added >= 10
                Check stats, prefix & ".export.regionCount", CountEntities(doc, "AcDbRegion", "AUDIT_C") = 1
                Set exported = Nothing
                For Each entity In doc.ModelSpace
                    If entity.ObjectName = "AcDbRegion" And entity.Layer = "AUDIT_C" Then Set exported = entity: Exit For
                Next entity
                If exported Is Nothing Then Err.Raise vbObjectError + 4499, "CheckNativeShapeRoundTrip", "Экспортированный бетонный Region не найден."
                ReadNativeCentralInertia exported, actualArea, actualX, actualY, actualIx, actualIy, actualIxy
                CheckNativeInertiaClose stats, prefix & ".export.area", actualArea, area
                CheckNativeInertiaClose stats, prefix & ".export.x", actualX, cx
                CheckNativeInertiaClose stats, prefix & ".export.y", actualY, cy
                CheckNativePrincipalShape stats, prefix, ix, iy, ixy, actualIx, actualIy, actualIxy
                If shape = 0 Then
                    CheckNativeInertiaClose stats, prefix & ".rectangle.Ix", actualIx, ix
                    CheckNativeInertiaClose stats, prefix & ".rectangle.Iy", actualIy, iy
                    CheckNativeInertiaClose stats, prefix & ".rectangle.Ixy", actualIxy, ixy
                End If
                CheckNativeRectangleEdges stats, prefix, exported
                path = Left$(drawingPath, Len(drawingPath) - 4) & "_shape_" & CStr(unitMode) & "_" & CStr(shape) & "_" & CStr(signValue) & ".dwg"
                If Len(Dir$(path)) > 0 Then Err.Raise vbObjectError + 4499, "CheckNativeShapeRoundTrip", "Тестовый DWG формы уже существует; перезапись запрещена."
                doc.SaveAs path
                LogLine stats, "REAL_SHAPE_DWG: " & doc.FullName & "|shape=" & model.ConcreteShapeType(1) & _
                    "; sourceArea=" & NativeNumber(area) & "; exportArea=" & NativeNumber(actualArea) & _
                    "; rotation=" & NativeNumber(model.ConcreteRotation(1))
            Next signValue
        Next shape
    Next unitMode
End Sub

' Проверяет изотропные Region на реальных сущностях: квадрат берет направление
' своей прямой грани, круг берет среднее остальных ориентированных элементов.
' Oracle читает исходные ребра отдельно от importer-а; круг не участвует в среднем.
Private Sub CheckNativeIsotropicRoundTrip(ByRef stats As TCadStats, ByVal doc As Object, _
        ByVal system As Object, ByVal unitsRange As Object, ByVal drawingPath As String)
    Dim scenario As Long, unitMode As Long, i As Long, j As Long, circleIndex As Long
    Dim squareIndex As Long, squareAngle As Double, expected As Double, angle As Double, circleX As Double
    Dim sumCos As Double, sumSin As Double, region As Object, circleRegion As Object, bar As Object
    Dim model As CSectionModel, importer As CAutoCADSectionModelImporter, entity As Object
    Dim exportedCircle As Object, exportedSquare As Object, added As Long, solves As Long
    Dim prefix As String, path As String, squareHandle As String, circleHandle As String
    Dim area As Double, cx As Double, cy As Double, ix As Double, iy As Double, ixy As Double
    ConfigureFixture system
    SetValue system, "AutoCAD.Export.PrincipalAxesMode", "None"
    SetValue system, "AutoCAD.Export.NeutralLineEnabled", "No"
    SetValue system, "AutoCAD.Export.LoadPointEnabled", "No"
    SetValue system, "AutoCAD.Export.ContourEnabled", "No"
    Set importer = New CAutoCADSectionModelImporter
    For unitMode = 0 To 1
        If unitMode = 0 Then
            SetValue unitsRange, "Length", "mm": CellForKey(unitsRange, "Length", 4).Value2 = "mm"
            SetValue unitsRange, "Area", "mm2": CellForKey(unitsRange, "Area", 4).Value2 = "mm2"
        Else
            SetValue unitsRange, "Length", "m": CellForKey(unitsRange, "Length", 4).Value2 = "m"
            SetValue unitsRange, "Area", "m2": CellForKey(unitsRange, "Area", 4).Value2 = "m2"
        End If
        For scenario = 0 To 4
            prefix = "realAutoCAD.isotropic." & CStr(unitMode) & "." & CStr(scenario)
            LogLine stats, "REAL_STAGE: " & prefix
            ClearOwnModelSpace doc
            sumCos = 0#: sumSin = 0#: squareHandle = vbNullString
            If scenario <= 1 Then
                angle = GEOM_PI / 6#: If scenario = 1 Then angle = 0#
                Set region = NativeRectangleRegion(doc, 60#, 60#, angle, 0#, 0#)
                region.Layer = "AUDIT_SHAPE_IN": squareHandle = region.Handle
                squareAngle = NativeFirstEdgeAngle(region)
                sumCos = Cos(2# * squareAngle): sumSin = Sin(2# * squareAngle)
                If scenario = 1 Then
                    Set region = NativeRectangleRegion(doc, 60#, 120#, GEOM_PI / 6#, 200#, 0#)
                    region.Layer = "AUDIT_SHAPE_IN"
                    sumCos = sumCos + Cos(GEOM_PI / 3#): sumSin = sumSin + Sin(GEOM_PI / 3#)
                End If
            ElseIf scenario = 2 Or scenario = 4 Then
                angle = 89# * GEOM_PI / 180#: If scenario = 4 Then angle = GEOM_PI / 4#
                For i = -1 To 1 Step 2
                    Set region = NativeRectangleRegion(doc, 60#, 120#, CDbl(i) * angle, 100# + 100# * i, 0#)
                    region.Layer = "AUDIT_SHAPE_IN"
                    sumCos = sumCos + Cos(2# * CDbl(i) * angle)
                    sumSin = sumSin + Sin(2# * CDbl(i) * angle)
                Next i
            End If
            expected = 0#
            If Abs(sumCos) > 0.000000001 Or Abs(sumSin) > 0.000000001 Then
                expected = Atn(sumSin / IIf(Abs(sumCos) < 0.000000001, 0.000000001, sumCos))
                If sumCos < 0# Then expected = expected + GEOM_PI
                expected = expected / 2#
            End If
            circleX = 400#: If scenario = 3 Then circleX = 0#
            Set circleRegion = NativeCircleRegion(doc, circleX, 0#, 20#)
            circleRegion.Layer = "AUDIT_SHAPE_IN": circleHandle = circleRegion.Handle
            For i = -1 To 1 Step 2
                For j = -1 To 1 Step 2
                    Set bar = NativeRectangleRegion(doc, 2#, 2#, 0#, circleX + 5# * i, 5# * j)
                    bar.Layer = "AUDIT_SHAPE_R"
                Next j
            Next i
            Set model = importer.ImportFromModelSpace(doc.ModelSpace, "AUDIT_SHAPE_IN", "AUDIT_SHAPE_R", "A400", 0#)
            circleIndex = 0: squareIndex = 0
            For i = 1 To model.ConcreteCount
                If model.ConcreteSourceHandle(i) = circleHandle Then circleIndex = i
                If Len(squareHandle) > 0 And model.ConcreteSourceHandle(i) = squareHandle Then squareIndex = i
            Next i
            Check stats, prefix & ".circle.found", circleIndex > 0
            If circleIndex = 0 Then Err.Raise vbObjectError + 4499, "CheckNativeIsotropicRoundTrip", "Круговой Region потерян при импорте."
            If squareIndex > 0 Then CheckNativeDirection stats, prefix & ".square.import", model.ConcreteRotation(squareIndex), squareAngle, 4#
            CheckNativeDirection stats, prefix & ".average.sources", model.AverageKnownConcreteElementRotation(), expected, 2#
            model.ApplyAverageRotationToEquivalentAreaFallbacks
            CheckNativeDirection stats, prefix & ".circle.assigned", model.ConcreteRotation(circleIndex), expected, 2#
            model.ApplyAverageRotationToEquivalentAreaFallbacks
            CheckNativeDirection stats, prefix & ".circle.repeated", model.ConcreteRotation(circleIndex), expected, 2#
            If squareIndex > 0 Then CheckNativeDirection stats, prefix & ".square.retained", model.ConcreteRotation(squareIndex), squareAngle, 4#
            PrepareNativeImportedSnapshot stats, system, model, prefix
            solves = SectionEquilibriumSolveCount(): added = Audit03ExportAutoCADDocumentForTests(doc)
            Check stats, prefix & ".export.noSolve", SectionEquilibriumSolveCount() = solves
            Check stats, prefix & ".export.regionCount", CountEntities(doc, "AcDbRegion", "AUDIT_C") = model.ConcreteCount
            Set exportedCircle = Nothing: Set exportedSquare = Nothing
            For Each entity In doc.ModelSpace
                If entity.ObjectName = "AcDbRegion" And entity.Layer = "AUDIT_C" Then
                    ReadNativeCentralInertia entity, area, cx, cy, ix, iy, ixy
                    If Abs(cx - circleX) < 0.000001 And Abs(cy) < 0.000001 Then Set exportedCircle = entity
                    If Abs(cx) < 0.000001 And Abs(cy) < 0.000001 Then Set exportedSquare = entity
                End If
            Next entity
            Check stats, prefix & ".circle.exported", Not exportedCircle Is Nothing
            If exportedCircle Is Nothing Then Err.Raise vbObjectError + 4499, "CheckNativeIsotropicRoundTrip", "Оболочка круга потеряна при экспорте."
            ReadNativeCentralInertia exportedCircle, area, cx, cy, ix, iy, ixy
            CheckNativeInertiaClose stats, prefix & ".circle.export.area", area, GEOM_PI * 400#
            CheckNativeDirection stats, prefix & ".circle.export.edge", NativeFirstEdgeAngle(exportedCircle), expected, 4#
            CheckNativeRectangleEdges stats, prefix & ".circle.export", exportedCircle
            If squareIndex > 0 Then
                Check stats, prefix & ".square.exported", Not exportedSquare Is Nothing
                If exportedSquare Is Nothing Then Err.Raise vbObjectError + 4499, "CheckNativeIsotropicRoundTrip", "Оболочка квадрата потеряна при экспорте."
                ReadNativeCentralInertia exportedSquare, area, cx, cy, ix, iy, ixy
                CheckNativeInertiaClose stats, prefix & ".square.export.area", area, 3600#
                CheckNativeDirection stats, prefix & ".square.export.edge", NativeFirstEdgeAngle(exportedSquare), squareAngle, 4#
                CheckNativeRectangleEdges stats, prefix & ".square.export", exportedSquare
            End If
            path = Left$(drawingPath, Len(drawingPath) - 4) & "_isotropic_" & CStr(unitMode) & "_" & CStr(scenario) & ".dwg"
            If Len(Dir$(path)) > 0 Then Err.Raise vbObjectError + 4499, "CheckNativeIsotropicRoundTrip", "Тестовый DWG уже существует; перезапись запрещена."
            doc.SaveAs path
            LogLine stats, "REAL_ISOTROPIC_DWG: " & doc.FullName & "|expected=" & NativeNumber(expected) & "; squareEdge=" & NativeNumber(squareAngle)
        Next scenario
    Next unitMode
End Sub

' Создает круговой Region без прямых граней, чтобы проверять именно средний
' угол сетки, а не подсказку из полилинии. Все исходные кривые принадлежат тесту.
Private Function NativeCircleRegion(ByVal doc As Object, ByVal x As Double, ByVal y As Double, ByVal radius As Double) As Object
    Dim center(0 To 2) As Double, circleEntity As Object, curves(0 To 0) As Object, regions As Variant
    center(0) = x: center(1) = y
    Set circleEntity = doc.ModelSpace.AddCircle(center, radius): Set curves(0) = circleEntity
    regions = doc.ModelSpace.AddRegion(curves): Set NativeCircleRegion = regions(LBound(regions))
    circleEntity.Delete
End Function

' Независимо читает угол первого прямого ребра настоящего Region; не вызывает
' importer. Копия и результаты Explode удаляются из собственного документа.
Private Function NativeFirstEdgeAngle(ByVal region As Object) As Double
    Dim copy As Object, pieces As Variant, line As Variant, p As Variant, q As Variant
    Dim dx As Double, dy As Double, found As Boolean, angle As Double
    Set copy = region.Copy(): pieces = copy.Explode
    For Each line In pieces
        If Not found And line.ObjectName = "AcDbLine" Then
            p = line.StartPoint: q = line.EndPoint: dx = q(0) - p(0): dy = q(1) - p(1)
            If Abs(dx) > 0.000000001 Or Abs(dy) > 0.000000001 Then
                If Abs(dx) < 0.000000001 Then
                    angle = Sgn(dy) * GEOM_PI / 2#
                Else
                    angle = Atn(dy / dx): If dx < 0# Then angle = angle + GEOM_PI
                End If
                found = True
            End If
        End If
        line.Delete
    Next line
    copy.Delete
    If Not found Then Err.Raise vbObjectError + 4499, "NativeFirstEdgeAngle", "В тестовой оболочке нет прямого ребра."
    NativeFirstEdgeAngle = angle
End Function

' Сравнивает направления без ложного отказа из-за эквивалентных углов:
' multiplier=2 для оси, multiplier=4 для взаимозаменяемых граней квадрата.
Private Sub CheckNativeDirection(ByRef stats As TCadStats, ByVal prefix As String, _
        ByVal actual As Double, ByVal expected As Double, ByVal multiplier As Double)
    CheckNativeInertiaClose stats, prefix & ".cos", Cos(multiplier * actual), Cos(multiplier * expected)
    CheckNativeInertiaClose stats, prefix & ".sin", Sin(multiplier * actual), Sin(multiplier * expected)
End Sub

' Получает настоящее допустимое НДС маленькой осевой нагрузки и записывает
' импортированную модель существующим writer-ом. Экспорт ниже читает только
' сохраненный Results; этот setup не подменяет snapshot готовой тестовой строкой.
Private Sub PrepareNativeImportedSnapshot(ByRef stats As TCadStats, ByVal system As Object, _
        ByVal model As CSectionModel, ByVal prefix As String)
    Dim settings As CSystemSettingsReader, units As CUnitSystem, provider As CMaterialModelProvider
    Dim batch As CBatchSectionCalculator, catalog As CCalculationProfileCatalog, writer As CNDMResultsWriter
    SetValue system, "AutoCAD.Export.CombinationID", "CAD_SHAPE"
    Set settings = Reader(): Set units = New CUnitSystem: units.LoadFromSettings settings
    Set provider = New CMaterialModelProvider: provider.Initialize settings, units
    Set catalog = New CCalculationProfileCatalog: catalog.LoadFromWorkbook ThisWorkbook
    Set batch = New CBatchSectionCalculator: batch.Initialize model, provider
    Set batch.ProfileCatalog = catalog: batch.ApplySettings settings, units
    batch.AddCombination "CAD_SHAPE", -1000#, 0#, 0#, "PR1", "Native imported shape", "Auto"
    batch.Execute
    LogLine stats, "REAL_SHAPE_STATE: " & prefix & "|" & batch.ResultAt(1).OverallMeta.ResultComment
    Check stats, prefix & ".snapshot.state", batch.ResultAt(1).DirectStateMeta.InternalStatus = rsSuccess
    If batch.ResultAt(1).DirectStateMeta.InternalStatus <> rsSuccess Then _
        Err.Raise vbObjectError + 4499, "PrepareNativeImportedSnapshot", "Не получено допустимое НДС импортированной формы: " & batch.ResultAt(1).DirectStateMeta.ResultComment
    Set writer = New CNDMResultsWriter: writer.WriteResults ThisWorkbook, model, provider, batch, units
End Sub

' Создает фактический замкнутый polygon Region в WCS-плоскости собственного
' test-document. Исходная Polyline после преобразования не остается в ModelSpace.
Private Function NativePolygonRegion(ByVal doc As Object, ByRef x() As Double, ByRef y() As Double) As Object
    Dim points() As Double, i As Long, line As Object, curves(0 To 0) As Object, regions As Variant
    ReDim points(0 To 2 * (UBound(x) + 1) - 1)
    For i = 0 To UBound(x): points(2 * i) = x(i): points(2 * i + 1) = y(i): Next i
    Set line = doc.ModelSpace.AddLightWeightPolyline(points): line.Closed = True
    Set curves(0) = line: regions = doc.ModelSpace.AddRegion(curves)
    Set NativePolygonRegion = regions(LBound(regions))
    line.Delete
End Function

' Независимая квадратура простого polygon по ориентированным ребрам.
' Сначала получаем интегралы относительно начала, затем снимаем перенос
' к центру площади. Здесь нет чтения CSectionModel или CAD MomentOfInertia.
Private Sub NativePolygonInertia(ByRef x() As Double, ByRef y() As Double, _
        ByRef area As Double, ByRef cx As Double, ByRef cy As Double, _
        ByRef ix As Double, ByRef iy As Double, ByRef ixy As Double)
    Dim i As Long, j As Long, cross As Double
    area = 0#: cx = 0#: cy = 0#: ix = 0#: iy = 0#: ixy = 0#
    For i = 0 To UBound(x)
        j = (i + 1) Mod (UBound(x) + 1)
        cross = x(i) * y(j) - x(j) * y(i)
        area = area + cross / 2#
        cx = cx + (x(i) + x(j)) * cross
        cy = cy + (y(i) + y(j)) * cross
        ix = ix + (y(i) ^ 2 + y(i) * y(j) + y(j) ^ 2) * cross / 12#
        iy = iy + (x(i) ^ 2 + x(i) * x(j) + x(j) ^ 2) * cross / 12#
        ixy = ixy + (2# * x(i) * y(i) + x(i) * y(j) + x(j) * y(i) + 2# * x(j) * y(j)) * cross / 24#
    Next i
    If area <= 0# Then Err.Raise vbObjectError + 4499, "NativePolygonInertia", "Тестовый polygon имеет неверную ориентированную площадь."
    cx = cx / (6# * area): cy = cy / (6# * area)
    ix = ix - area * cy ^ 2: iy = iy - area * cx ^ 2: ixy = ixy - area * cx * cy
End Sub

' Читает центральный тензор непосредственно из настоящего Region, независимо
' от importer-а. Знак отрицательной XY-компоненты AutoCAD подтвержден отдельным
' native oracle; геометрические значения здесь всегда миллиметровые.
Private Sub ReadNativeCentralInertia(ByVal region As Object, ByRef area As Double, _
        ByRef cx As Double, ByRef cy As Double, ByRef ix As Double, ByRef iy As Double, ByRef ixy As Double)
    Dim center As Variant, moments As Variant, product As Variant
    area = region.Area: center = region.Centroid
    cx = center(LBound(center)): cy = center(LBound(center) + 1)
    moments = region.MomentOfInertia: product = region.ProductOfInertia
    ix = moments(LBound(moments)) - area * cy ^ 2
    iy = moments(LBound(moments) + 1) - area * cx ^ 2
    If IsArray(product) Then
        ixy = -CDbl(product(LBound(product))) - area * cx * cy
    Else
        ixy = -CDbl(product) - area * cx * cy
    End If
End Sub

' Сравнивает отношение собственных инерций и беззнаковое направление оси.
' Нормализованные (Ix-Iy, -2Ixy) задают cos(2a)/sin(2a), поэтому проверка
' допускает ту же ось плюс 180 градусов, но не ее зеркальное направление.
Private Sub CheckNativePrincipalShape(ByRef stats As TCadStats, ByVal prefix As String, _
        ByVal ix As Double, ByVal iy As Double, ByVal ixy As Double, _
        ByVal ex As Double, ByVal ey As Double, ByVal exy As Double)
    Dim delta As Double, exportDelta As Double, ratio As Double, exportRatio As Double
    delta = Sqr((ix - iy) ^ 2 + 4# * ixy ^ 2)
    exportDelta = Sqr((ex - ey) ^ 2 + 4# * exy ^ 2)
    ratio = (ix + iy + delta) / (ix + iy - delta)
    exportRatio = (ex + ey + exportDelta) / (ex + ey - exportDelta)
    CheckNativeInertiaClose stats, prefix & ".principal.ratio", exportRatio, ratio
    CheckNativeInertiaClose stats, prefix & ".principal.cos2", (ex - ey) / exportDelta, (ix - iy) / delta
    CheckNativeInertiaClose stats, prefix & ".principal.sin2", -2# * exy / exportDelta, -2# * ixy / delta
End Sub

' Проверяет экспортированную оболочку как четыре прямых взаимно перпендикулярных
' ребра. Даже для polygon исходника export должен дать rectangle, не исходный
' triangle/L-contour. Временные Explode-сущности удаляются только из own DWG.
Private Sub CheckNativeRectangleEdges(ByRef stats As TCadStats, ByVal prefix As String, ByVal region As Object)
    Dim copy As Object, pieces As Variant, line As Variant, p As Variant, q As Variant
    Dim vx As Double, vy As Double, wx As Double, wy As Double, first As Boolean, count As Long
    Set copy = region.Copy(): pieces = copy.Explode
    For Each line In pieces
        count = count + 1
        Check stats, prefix & ".rectangle.line." & CStr(count), line.ObjectName = "AcDbLine"
        If line.ObjectName = "AcDbLine" Then
            p = line.StartPoint: q = line.EndPoint
            wx = q(0) - p(0): wy = q(1) - p(1)
            If first Then
                Check stats, prefix & ".rectangle.edgeDirection." & CStr(count), _
                    Abs(vx * wy - vy * wx) <= 0.000001 Or Abs(vx * wx + vy * wy) <= 0.000001
            Else
                vx = wx: vy = wy: first = True
            End If
        End If
        line.Delete
    Next line
    copy.Delete
    Check stats, prefix & ".rectangle.fourEdges", count = 4
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
