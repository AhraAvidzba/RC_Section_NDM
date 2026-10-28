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

' -------------------- ДЛЯ ТЕСТОВ --------------------
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
    TestConfigureFixture system
    TestPrepareResults stats, system, profiles
    prepared = system.Formula
    For position = 0 To 1
        If position = 1 Then
            shifted.Formula = prepared
            ThisWorkbook.Names.Item("rngSystemSettings").RefersTo = "='" & shifted.Worksheet.Name & "'!" & shifted.Address
            Set system = shifted
        End If
        TestCheckPresentation stats, system, position
        TestCheckImporter stats, system, position
        TestCheckCleanup stats, system, position
        TestCheckSettingEffects stats, system, position
        TestCheckInvalidFields stats, system, position
        TestCheckLayerContracts stats, system, position
    Next position
    GoTo Restore
FailedRun:
    stats.Failed = stats.Failed + 1
    TestLogLine stats, "FAIL: autoCADConfig.runtime; " & CStr(Err.Number) & "; " & Err.Description
Restore:
    On Error GoTo FailedRestore
    If Len(originalName) > 0 Then ThisWorkbook.Names.Item("rngSystemSettings").RefersTo = originalName
    If Not IsEmpty(saved) Then ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange.Formula = saved
    If Not shifted Is Nothing Then shifted.Formula = savedShifted
    If Not profiles Is Nothing Then profiles.Formula = savedProfiles
    TestCheck stats, "autoCADConfig.restore.name", ThisWorkbook.Names.Item("rngSystemSettings").RefersTo = originalName
    GoTo Finish
FailedRestore:
    stats.Failed = stats.Failed + 1
    TestLogLine stats, "FAIL: autoCADConfig.restore; " & Err.Description
Finish:
    passed = stats.Passed: failed = stats.Failed
    TestLogLine stats, "AUTOCAD_CONFIG_CASES: " & CStr(stats.Cases)
    TestLogLine stats, "TOTAL_AUTOCAD_CONFIG: passed=" & CStr(passed) & "; failed=" & CStr(failed)
    RunAudit03AutoCADConfigTests = stats.Report
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Разделяет слои геометрии, подписей и контура; цвета задаются индексами
' ACI 1..255, совместимыми с сущностями и вновь создаваемыми Layer.
Private Sub TestConfigureFixture(ByVal system As Object)
    Dim keys As Variant, values As Variant, i As Long
    keys = TestCadKeys()
    values = Array("CAD_CFG", "Yes", "Transformed", "Yes", "Yes", "NamesAndValues", _
        "AUDIT_C", "AUDIT_R", "AUDIT_CONTOUR", "AUDIT_CT", "AUDIT_CC", "AUDIT_RT", "AUDIT_RC", _
        9, 5, 1, 6, 8)
    For i = 0 To UBound(keys): TestSetValue system, CStr(keys(i)), values(i): Next i
    TestSetValue system, "Mesh.MinElementArea", 0#
    TestSetValue system, "Mesh.CellGeometry", "Adaptive"
    TestSetValue system, "General.NonCriticalMessagesEnabled", "No"
    TestSetValue system, "Calculation.ZeroMomentPerDepth", 0#
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сохраняет два рассчитанных НДС с разными номерами сочетаний и напряжениями.
' Проверка выбора сочетания читает эти данные без нового расчёта.
Private Sub TestPrepareResults(ByRef stats As TCadStats, ByVal system As Object, ByVal profiles As Object)
    Dim settings As CSystemSettingsReader, units As CUnitSystem, section As CSectionModel
    Dim provider As CMaterialModelProvider, batch As CBatchSectionCalculator, writer As CNDMResultsWriter
    Dim summaryWriter As CBatchResultWriter
    Dim key As Variant, x As Long, y As Long, elementIndex As Long
    For Each key In Array("Calculation.Stability.Enabled", "Calculation.Strength.Capacity", "Calculation.Crack.Width")
        TestSetProfile profiles, CStr(key), "No"
    Next key
    TestSetProfile profiles, "Calculation.Strength.DirectState", "Yes"
    TestSetProfile profiles, "Visualization.State", "StrengthState"
    TestSetProfile profiles, "Visualization.Quantity", "Stress"
    TestSetProfile profiles, "MaterialModel.Strength.ConcreteTension", "UseDiagram"
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
    section.Contours.AddContourLine "CAD_BOTTOM", -100#, -100#, 100#, -100#
    section.Contours.AddContourLine "CAD_RIGHT", 100#, -100#, 100#, 100#
    section.Contours.AddContourLine "CAD_TOP", 100#, 100#, -100#, 100#
    section.Contours.AddContourLine "CAD_LEFT", -100#, 100#, -100#, -100#
    Set provider = New CMaterialModelProvider: provider.Initialize settings, units
    Set batch = New CBatchSectionCalculator: batch.Initialize section, provider
    Dim catalog As CCalculationProfileCatalog
    Set catalog = New CCalculationProfileCatalog: catalog.LoadFromWorkbook ThisWorkbook
    Set batch.ProfileCatalog = catalog: batch.ApplySettings settings, units
    batch.AddCombination "CAD_CFG", 0#, 100000#, 50000#, "PR1", "AutoCAD Config fixture", "Auto"
    batch.AddCombination "CAD_CFG_2", 0#, 200000#, 100000#, "PR1", "AutoCAD Config fixture", "Auto"
    batch.Execute
    TestLogLine stats, "AUTOCAD_SNAPSHOT: " & batch.ResultAt(1).OverallMeta.ResultComment
    TestLogLine stats, "AUTOCAD_SNAPSHOT: " & batch.ResultAt(2).OverallMeta.ResultComment
    TestCheck stats, "autoCADConfig.snapshot.state1", batch.ResultAt(1).DirectStateMeta.InternalStatus = rsSuccess
    TestCheck stats, "autoCADConfig.snapshot.state2", batch.ResultAt(2).DirectStateMeta.InternalStatus = rsSuccess
    ' Определяющее сочетание Worst должно относиться к текущим тестовым результатам,
    ' а не к сохранённому пользовательскому расчёту или другому тесту.
    Set summaryWriter = New CBatchResultWriter: summaryWriter.WriteSummary ThisWorkbook, batch
    Set writer = New CNDMResultsWriter: writer.WriteResults ThisWorkbook, section, PrepareSectionSnapshot(section, provider), batch, units
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает фактические подписи и выбор цвета/слоя всех физических состояний,
' а не только значения getters. Флаги требуют отдельного настоящего DWG gate.
Private Sub TestCheckPresentation(ByRef stats As TCadStats, ByVal system As Object, ByVal position As Long)
    Dim settings As CSystemSettingsReader, result As Object, material As Variant, physical As Variant
    Dim key As String, expectedLayer As String, expectedColor As Long, prefix As String
    Set settings = TestReader()
    For Each material In Array("Concrete", "Rebar")
        For Each physical In Array("Tension", "Compression", "NearZero")
            Set result = Audit03AutoCADPresentationForTests(settings, CStr(material), CStr(physical), "E1", 12.3456, 2)
            stats.Cases = stats.Cases + 1
            prefix = "autoCADConfig.presentation." & CStr(position) & "." & CStr(material) & "." & CStr(physical)
            If CStr(material) = "Concrete" Then expectedLayer = "AUDIT_C" Else expectedLayer = "AUDIT_R"
            TestCheck stats, prefix & ".geometry", CStr(result("GeometryLayer")) = expectedLayer
            key = CStr(material) & CStr(physical)
            If CStr(physical) = "NearZero" Then key = "Neutral"
            expectedColor = CLng(TestCellForKey(system, "AutoCAD.Color." & key, 2).Value2)
            TestCheck stats, prefix & ".color", CLng(result("Color")) = expectedColor
            key = CStr(material) & "Compression"
            If CStr(physical) = "Tension" Then key = CStr(material) & "Tension"
            TestCheck stats, prefix & ".annotation", CStr(result("AnnotationLayer")) = CStr(TestCellForKey(system, "AutoCAD.Layer." & key, 2).Value2)
            TestCheck stats, prefix & ".label", CStr(result("Label")) = "E1 12.35"
        Next physical
    Next material
    Dim value As Variant, flag As Variant, saved As Variant, first As String, second As String, solves As Long
    saved = system.Formula: solves = SectionEquilibriumSolveCount()
    TestSetValue system, "AutoCAD.Export.LabelMode", "ValuesOnly"
    Set result = Audit03AutoCADPresentationForTests(TestReader(), "Concrete", "Tension", "E1", 12.3456, 2)
    TestCheck stats, "autoCADConfig.label.values." & CStr(position), CStr(result("Label")) = "12.35"
    For Each value In Array("None", "Concrete", "Transformed")
        TestSetValue system, "AutoCAD.Export.PrincipalAxesMode", CStr(value)
        Set result = Audit03AutoCADPresentationForTests(TestReader(), "Rebar", "Compression", "R1", 1#, 0)
        TestCheck stats, "autoCADConfig.axes." & CStr(position) & "." & CStr(value), CStr(result("AxesMode")) = CStr(value)
    Next value
    For Each flag In Array("NeutralLine", "LoadPoint", "Contour")
        For Each value In Array("No", "Yes")
            TestSetValue system, "AutoCAD.Export." & CStr(flag) & "Enabled", CStr(value)
            Set result = Audit03AutoCADPresentationForTests(TestReader(), "Rebar", "Compression", "R1", 1#, 0)
            TestCheck stats, "autoCADConfig.flag." & CStr(position) & "." & CStr(flag) & "." & CStr(value), CBool(result(CStr(flag))) = (CStr(value) = "Yes")
        Next value
    Next flag
    For Each value In Array(1, 255)
        TestSetValue system, "AutoCAD.Color.RebarTension", value
        Set result = Audit03AutoCADPresentationForTests(TestReader(), "Rebar", "Tension", "R1", 1#, 0)
        TestCheck stats, "autoCADConfig.color.bound." & CStr(position) & "." & CStr(value), CLng(result("Color")) = CLng(value)
    Next value
    modTestPlotConfig.TestSetOutputCombinationInTable system, "CAD_CFG"
    first = Audit02ReadExportSnapshotForTests(ThisWorkbook)
    modTestPlotConfig.TestSetOutputCombinationInTable system, "CAD_CFG_2"
    second = Audit02ReadExportSnapshotForTests(ThisWorkbook)
    TestCheck stats, "autoCADConfig.combination.explicit." & CStr(position), Left$(first, 8) = "CAD_CFG|" And Left$(second, 10) = "CAD_CFG_2|"
    TestCheck stats, "autoCADConfig.combination.changed." & CStr(position), first <> second
    TestCheck stats, "autoCADConfig.combination.valuesChanged." & CStr(position), _
        Mid$(first, InStr(first, vbCrLf) + 2) <> Mid$(second, InStr(second, vbCrLf) + 2)
    modTestPlotConfig.TestSetOutputCombinationInTable system, "Worst"
    first = Audit02ReadExportSnapshotForTests(ThisWorkbook)
    TestCheck stats, "autoCADConfig.combination.worst." & CStr(position), Left$(first, 7) = "CAD_CFG"
    TestCheck stats, "autoCADConfig.presentation.noSolve." & CStr(position), SectionEquilibriumSolveCount() = solves
    system.Formula = saved
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Выбор каждого слоя импорта меняет набор читаемых Region и их координаты.
' Геометрия чертежа остаётся в мм независимо от пользовательских входных единиц.
Private Sub TestCheckImporter(ByRef stats As TCadStats, ByVal system As Object, ByVal position As Long)
    Dim model As CSectionModel, importer As CAutoCADSectionModelImporter, regions As Collection
    Dim settings As CSystemSettingsReader, prefix As String, saved As Variant
    saved = system.Formula: Set regions = TestImportRegions(): Set importer = New CAutoCADSectionModelImporter
    Set settings = TestReader(): Set model = importer.ImportConfiguredModelSpace(regions, settings, Nothing, TestEmptyContourLayers())
    prefix = "autoCADConfig.import." & CStr(position)
    TestCheck stats, prefix & ".counts", model.ConcreteCount = 1 And model.RebarCount = 1
    TestCheckClose stats, prefix & ".concreteMm", model.ConcreteX(1), 20#
    TestCheckClose stats, prefix & ".steelMm", model.RebarX(1), 80#
    Dim emptyLayers As Collection, layer As CTestFakeAcadContour, layerName As Variant
    Set emptyLayers = New Collection
    For Each layerName In Array(settings.GetRawString("AutoCAD.Common.SectionContourLayer"), settings.GetRawString("AutoCAD.Common.OpeningContourLayer"))
        If Len(CStr(layerName)) > 0 Then
            Set layer = New CTestFakeAcadContour: emptyLayers.Add layer, CStr(layerName)
        End If
    Next layerName
    Set model = importer.ImportConfiguredModelSpace(regions, settings, Nothing, emptyLayers)
    TestCheck stats, prefix & ".existingEmptyLayers.meshKept", model.ConcreteCount = 1 And model.RebarCount = 1
    TestCheck stats, prefix & ".existingEmptyLayers.noInventedContour", model.Contours.Count = 0
    TestCheck stats, prefix & ".existingEmptyLayers.comment", InStr(importer.ImportComment, "бетонной сетке") > 0
    TestSetValue system, "AutoCAD.Common.ConcreteLayer", "AUDIT_IC2"
    Set model = importer.ImportConfiguredModelSpace(regions, TestReader(), Nothing, TestEmptyContourLayers())
    TestCheckClose stats, prefix & ".changedConcrete", model.ConcreteX(1), 120#
    TestCheckClose stats, prefix & ".concreteLayerKeepsSteel", model.RebarX(1), 80#
    Dim presentation As Object
    Set presentation = Audit03AutoCADPresentationForTests(TestReader(), "Concrete", "Tension", "C1", 1#, 2)
    TestCheck stats, prefix & ".concreteSharedWithExport", CStr(presentation("GeometryLayer")) = "AUDIT_IC2"
    system.Formula = saved
    TestSetValue system, "AutoCAD.Common.RebarLayer", "AUDIT_IR2"
    Set model = importer.ImportConfiguredModelSpace(regions, TestReader(), Nothing, TestEmptyContourLayers())
    TestCheckClose stats, prefix & ".changedSteel", model.RebarX(1), 180#
    TestCheckClose stats, prefix & ".steelLayerKeepsConcrete", model.ConcreteX(1), 20#
    Set presentation = Audit03AutoCADPresentationForTests(TestReader(), "Rebar", "Tension", "R1", 1#, 2)
    TestCheck stats, prefix & ".rebarSharedWithExport", CStr(presentation("GeometryLayer")) = "AUDIT_IR2"
    stats.Cases = stats.Cases + 4: system.Formula = saved
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Рабочая очистка удаляет оформление, но сохраняет всю геометрию общих слоев,
' в том числе помеченные копии контуров. Region защищен даже на annotation-слое.
Private Sub TestCheckCleanup(ByRef stats As TCadStats, ByVal system As Object, ByVal position As Long)
    Dim ms As Object, region As CTestFakeAcadRegion, line As CTestFakeAcadLine, foreign As CTestFakeAcadLine, deleted As Long
    Dim settings As CSystemSettingsReader, contour As CTestFakeAcadContour, curves As Collection
    Dim protected As Collection, layer As Variant, owned As Long, i As Long
    Set ms = CreateObject("Scripting.Dictionary")
    Set region = New CTestFakeAcadRegion: region.TestInitialize 100#, 0#, 0#, 100#, 100#, 0#, "AUDIT_CT", "KEEP_REGION"
    Set line = New CTestFakeAcadLine: line.TestInitialize 0#, 0#, 1#, 0#, "AUDIT_CT"
    Set foreign = New CTestFakeAcadLine: foreign.TestInitialize 0#, 0#, 1#, 0#, "FOREIGN"
    ms.Add 0, region: ms.Add 1, line: ms.Add 2, foreign
    Set settings = TestReader(): Set protected = New Collection
    For Each layer In Array("AUDIT_C", "AUDIT_R", "AUDIT_CONTOUR", settings.GetRequiredString("AutoCAD.Common.OpeningContourLayer"))
        For owned = 0 To 1
            Set contour = New CTestFakeAcadContour
            contour.TestInitializeLoop "AcDbPolyline", CStr(layer), "KEEP_GEOMETRY", curves
            contour.TestSetLegacyContourTag (owned = 1)
            protected.Add contour: ms.Add ms.Count, contour
        Next owned
    Next layer
    deleted = Audit03CleanupAutoCADModelSpaceForTests(ms, settings)
    TestCheck stats, "autoCADConfig.cleanup.region." & CStr(position), Not region.TestDeleted
    TestCheck stats, "autoCADConfig.cleanup.line." & CStr(position), line.TestDeleted
    TestCheck stats, "autoCADConfig.cleanup.foreign." & CStr(position), Not foreign.TestDeleted
    For i = 1 To protected.Count
        Set contour = protected(i)
        TestCheck stats, "autoCADConfig.cleanup.protected." & CStr(position) & "." & CStr(i), Not contour.TestDeleted
    Next i
    TestCheck stats, "autoCADConfig.cleanup.count." & CStr(position), deleted = 1
    ' Общая граница имеет приоритет, даже если пользователь назначил ей
    ' тот же слой, что и подписи: на этом слое не удаляется ни один объект.
    Dim saved As Variant: saved = system.Formula
    TestSetValue system, "AutoCAD.Common.SectionContourLayer", "AUDIT_CT"
    Set line = New CTestFakeAcadLine: line.TestInitialize 0#, 0#, 1#, 0#, "AUDIT_CT"
    Set contour = New CTestFakeAcadContour
    contour.TestInitializeLoop "AcDbPolyline", "AUDIT_CT", "KEEP_SHARED", curves
    contour.TestSetLegacyContourTag True
    Set ms = CreateObject("Scripting.Dictionary"): ms.Add 0, line: ms.Add 1, contour
    deleted = Audit03CleanupAutoCADModelSpaceForTests(ms, TestReader())
    TestCheck stats, "autoCADConfig.cleanup.sharedLayer." & CStr(position), deleted = 0 And Not line.TestDeleted And Not contour.TestDeleted
    system.Formula = saved
    stats.Cases = stats.Cases + 2
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Меняет каждый экспортный слой и цвет по отдельности. Проверяет соответствующее
' поле оформления: чтения значения по умолчанию или сравнения разных сочетаний
' недостаточно для подтверждения работы настройки.
Private Sub TestCheckSettingEffects(ByRef stats As TCadStats, ByVal system As Object, ByVal position As Long)
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
        Dim settingKey As String
        settingKey = "AutoCAD.Layer." & CStr(keys(i))
        If CStr(keys(i)) = "Concrete" Or CStr(keys(i)) = "Rebar" Then settingKey = "AutoCAD.Common." & CStr(keys(i)) & "Layer"
        If CStr(keys(i)) = "Contour" Then settingKey = "AutoCAD.Common.SectionContourLayer"
        TestSetValue system, settingKey, name
        Set result = Audit03AutoCADPresentationForTests(TestReader(), CStr(materials(i)), CStr(states(i)), "E1", 1#, 2)
        stats.Cases = stats.Cases + 1
        TestCheck stats, "autoCADConfig.effect." & CStr(position) & "." & settingKey, _
            CStr(result(CStr(fields(i)))) = name
        system.Formula = saved
    Next i
    keys = Array("ConcreteTension", "ConcreteCompression", "RebarTension", "RebarCompression", "Neutral")
    materials = Array("Concrete", "Concrete", "Rebar", "Rebar", "Concrete")
    states = Array("Tension", "Compression", "Tension", "Compression", "NearZero")
    Dim color As Variant
    For i = 0 To UBound(keys)
        For Each color In Array(1, 255)
            TestSetValue system, "AutoCAD.Color." & CStr(keys(i)), color
            Set result = Audit03AutoCADPresentationForTests(TestReader(), CStr(materials(i)), CStr(states(i)), "E1", 1#, 2)
            stats.Cases = stats.Cases + 1
            TestCheck stats, "autoCADConfig.effect." & CStr(position) & ".AutoCAD.Color." & CStr(keys(i)) & "." & CStr(color), _
                CLng(result("Color")) = CLng(color)
            system.Formula = saved
        Next color
    Next i
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Невалидные поля проходят соответствующий consumer в двух положениях Config.
' Ошибка должна объяснять причину, текущую ячейку и требуемое исправление.
Private Sub TestCheckInvalidFields(ByRef stats As TCadStats, ByVal system As Object, ByVal position As Long)
    Dim key As Variant, values As Variant, value As Variant, target As Object, saved As Variant
    Dim reason As String, prefix As String, index As Long
    saved = system.Formula
    For Each key In TestCadKeys()
        values = Array(vbNullString, "TODO", "abc", CVErr(2015))
        If InStr(1, CStr(key), ".Layer.", vbTextCompare) > 0 Or InStr(1, CStr(key), "Layer", vbTextCompare) > 0 Then _
            values = Array(vbNullString, "TODO", "bad/name", "bad,name", String$(256, "a"), "bad" & vbLf & "name", CVErr(2015))
        If InStr(1, CStr(key), ".Color.", vbTextCompare) > 0 Then _
            values = Array(vbNullString, "TODO", "abc", CVErr(2015), -1, 0, 256, 1.5, 2147483648#)
        index = 0
        For Each value In values
            system.Formula = saved: Set target = TestCellForKey(system, CStr(key), 2): target.Value2 = value
            reason = TestConsumeError(stats, CStr(key))
            prefix = "autoCADConfig.invalid." & CStr(position) & "." & CStr(key) & "." & CStr(index)
            TestCheck stats, prefix & ".rejected", Len(reason) > 0
            TestCheck stats, prefix & ".key", InStr(1, reason, CStr(key), vbTextCompare) > 0
            TestCheck stats, prefix & ".address", InStr(1, reason, target.Address(False, False), vbTextCompare) > 0
            TestCheck stats, prefix & ".action", InStr(1, reason, "введите", vbTextCompare) > 0 Or InStr(1, reason, "выберите", vbTextCompare) > 0 Or InStr(1, reason, "исправ", vbTextCompare) > 0
            TestLogLine stats, "AUTOCAD_INPUT: " & prefix & "|" & reason
            index = index + 1
        Next value
        system.Formula = saved: TestCellForKey(system, CStr(key), 1).Value2 = "Missing." & CStr(key)
        reason = TestConsumeError(stats, CStr(key))
        TestCheck stats, "autoCADConfig.missing." & CStr(position) & "." & CStr(key), Len(reason) > 0 And InStr(1, reason, CStr(key), vbTextCompare) > 0
        system.Formula = saved
        TestCheck stats, "autoCADConfig.recovery." & CStr(position) & "." & CStr(key), Len(TestConsumeError(stats, CStr(key))) = 0
    Next key
    system.Formula = saved
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет межполевые конфликты, границы имени и неактивный contour-consumer.
' Очистка зависит только от слоев, поэтому поврежденный цвет ее не блокирует.
Private Sub TestCheckLayerContracts(ByRef stats As TCadStats, ByVal system As Object, ByVal position As Long)
    Dim saved As Variant, key As Variant, name As Variant, reason As String, target As Object, result As Object
    saved = system.Formula
    For Each key In Array("AutoCAD.Common.ConcreteLayer", "AutoCAD.Common.RebarLayer")
        For Each name In Array("AUDIT_CT", "AUDIT_CONTOUR", "RC_NDM_Axes", "RC_NDM_LoadPoint", "RC_NDM_NeutralLine", "RC_NDM_Warnings")
            Set target = TestCellForKey(system, CStr(key), 2): target.Value2 = CStr(name)
            reason = TestConsumeError(stats, CStr(key))
            TestCheck stats, "autoCADConfig.layerCollision." & CStr(position) & "." & CStr(key) & "." & CStr(name), _
                InStr(1, reason, CStr(key), vbTextCompare) > 0 And InStr(1, reason, target.Address(False, False), vbTextCompare) > 0
            system.Formula = saved
        Next name
    Next key
    TestSetValue system, "AutoCAD.Common.RebarLayer", "audit_c"
    On Error Resume Next
    Dim importer As CAutoCADSectionModelImporter, model As CSectionModel
    Set importer = New CAutoCADSectionModelImporter
    Set model = importer.ImportConfiguredModelSpace(TestImportRegions(), TestReader(), Nothing, TestEmptyContourLayers())
    reason = Err.Description: Err.Clear
    On Error GoTo 0
    TestCheck stats, "autoCADConfig.import.sameLayer." & CStr(position), InStr(1, reason, "совпадают", vbTextCompare) > 0
    system.Formula = saved
    TestSetValue system, "AutoCAD.Export.ContourEnabled", "No"
    TestSetValue system, "AutoCAD.Common.SectionContourLayer", "bad/name"
    TestCheck stats, "autoCADConfig.contour.inactive." & CStr(position), Len(TestConsumeError(stats, "AutoCAD.Export.ContourEnabled")) = 0
    system.Formula = saved
    TestSetValue system, "AutoCAD.Color.RebarTension", "TODO"
    TestCheckCleanup stats, system, position
    system.Formula = saved
    For Each name In Array("Слой бетона длинный", String$(255, "a"))
        TestSetValue system, "AutoCAD.Common.ConcreteLayer", CStr(name)
        Set result = Audit03AutoCADPresentationForTests(TestReader(), "Concrete", "Tension", "C1", 1#, 2)
        TestCheck stats, "autoCADConfig.layer.validBoundary." & CStr(position) & "." & CStr(Len(CStr(name))), result("GeometryLayer") = CStr(name)
        system.Formula = saved
    Next name
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет фактический отказ обычного чтения результатов и оформления.
' Ошибка Config должна сохранять свою причину, а не превращаться в численную несходимость.
Private Function TestConsumeError(ByRef stats As TCadStats, ByVal key As String) As String
    Dim result As Object, model As CSectionModel, importer As CAutoCADSectionModelImporter, text As String
    stats.Cases = stats.Cases + 1
    On Error GoTo Rejected
    If Left$(key, 15) = "AutoCAD.Import." Then
        Set importer = New CAutoCADSectionModelImporter
        Set model = importer.ImportConfiguredModelSpace(TestImportRegions(), TestReader(), Nothing, TestEmptyContourLayers())
    ElseIf key = "Output.CombinationID" Then
        TestSetValue ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange, "Output.CombinationMode", "User"
        text = Audit02ReadExportSnapshotForTests(ThisWorkbook)
    Else
        Set result = Audit03AutoCADPresentationForTests(TestReader(), "Concrete", "Tension", "C1", 1#, 2)
    End If
    If key = "AutoCAD.Common.ConcreteLayer" Or key = "AutoCAD.Common.RebarLayer" Then
        Set importer = New CAutoCADSectionModelImporter
        Set model = importer.ImportConfiguredModelSpace(TestImportRegions(), TestReader(), Nothing, TestEmptyContourLayers())
    End If
    Exit Function
Rejected:
    TestConsumeError = Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Один список содержит самостоятельные CAD-поля без дублирования общих
' слоев. Изменение общего слоя проверяется и импортером, и экспортом.
Private Function TestCadKeys() As Variant
    TestCadKeys = Array("Output.CombinationID", "AutoCAD.Export.NeutralLineEnabled", _
        "AutoCAD.Export.PrincipalAxesMode", "AutoCAD.Export.LoadPointEnabled", "AutoCAD.Export.ContourEnabled", _
        "AutoCAD.Export.LabelMode", "AutoCAD.Common.ConcreteLayer", "AutoCAD.Common.RebarLayer", "AutoCAD.Common.SectionContourLayer", _
        "AutoCAD.Layer.ConcreteTension", "AutoCAD.Layer.ConcreteCompression", "AutoCAD.Layer.RebarTension", "AutoCAD.Layer.RebarCompression", _
        "AutoCAD.Color.ConcreteTension", "AutoCAD.Color.ConcreteCompression", "AutoCAD.Color.RebarTension", "AutoCAD.Color.RebarCompression", _
        "AutoCAD.Color.Neutral")
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Четыре реальные Region-fixtures позволяют отличить выбор каждого слоя.
Private Function TestImportRegions() As Collection
    Dim result As Collection, region As CTestFakeAcadRegion, i As Long, layers As Variant
    Set result = New Collection: layers = Array("AUDIT_C", "AUDIT_R", "AUDIT_IC2", "AUDIT_IR2")
    For i = 0 To 3
        Set region = New CTestFakeAcadRegion
        If i Mod 2 = 0 Then
            region.TestInitialize 40000#, 20# + 50# * i, 30#, 133333333#, 133333333#, 0#, CStr(layers(i)), "C" & CStr(i)
        Else
            region.TestInitialize 201#, 30# + 50# * i, 90#, 3215#, 3215#, 0#, CStr(layers(i)), "R" & CStr(i)
        End If
        result.Add region
    Next i
    Set TestImportRegions = result
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Читает текущую привязку Config; перенесенная таблица не кешируется.
Private Function TestReader() As CSystemSettingsReader
    Set TestReader = New CSystemSettingsReader: TestReader.LoadFromWorkbook ThisWorkbook
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' в тестовом чертеже нет контурных слоёв. Их имена в Config
' не должны препятствовать импорту Region сетки с четырёх заданных слоёв.
Private Function TestEmptyContourLayers() As Collection
    Set TestEmptyContourLayers = New Collection
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Ищет ключ и возвращает настоящую ячейку для изменения и проверки адреса.
Private Function TestCellForKey(ByVal table As Object, ByVal key As String, ByVal column As Long) As Object
    Dim row As Long
    For row = 2 To table.Rows.Count
        If CStr(table.Cells(row, 1).Value2) = key Then Set TestCellForKey = table.Cells(row, column): Exit Function
    Next row
    Err.Raise vbObjectError + 4499, "CellForKey", "В fixture отсутствует Config " & key
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Меняет значение без изменения самого ключа или структуры таблицы.
Private Sub TestSetValue(ByVal table As Object, ByVal key As String, ByVal value As Variant)
    TestCellForKey(table, key, 2).Value2 = value
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Ищет PR1 в шапке таблицы профилей, не предполагая фиксированный столбец.
Private Sub TestSetProfile(ByVal table As Object, ByVal key As String, ByVal value As String)
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

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Накапливает все assertions, чтобы отрицательный baseline не обрывал набор.
Private Sub TestCheck(ByRef stats As TCadStats, ByVal name As String, ByVal condition As Boolean)
    If condition Then
        stats.Passed = stats.Passed + 1: TestLogLine stats, "OK: " & name
    Else
        stats.Failed = stats.Failed + 1: TestLogLine stats, "FAIL: " & name
    End If
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает точные координаты fixtures в мм и сохраняет исходные числа.
Private Sub TestCheckClose(ByRef stats As TCadStats, ByVal name As String, ByVal actual As Double, ByVal expected As Double)
    TestLogLine stats, "AUTOCAD_NUMBER: " & name & "|actual=" & CStr(actual) & "|expected=" & CStr(expected) & "|tolerance=0.00000001"
    TestCheck stats, name, Abs(actual - expected) <= 0.00000001
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращаемый отчет не зависит от доступности дополнительного файла progress.
Private Sub TestLogLine(ByRef stats As TCadStats, ByVal line As String)
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

' -------------------- ДЛЯ ТЕСТОВ --------------------
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
    TestLogLine stats, "REAL_STAGE: prepare workbook snapshot"
    Set system = ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange
    Set profiles = ThisWorkbook.Names.Item("rngCalculationProfiles").RefersToRange
    Set unitsRange = ThisWorkbook.Names.Item("rngUnitSettings").RefersToRange
    Set signs = ThisWorkbook.Names.Item("rngSignConventionSettings").RefersToRange
    saved = system.Formula: savedProfiles = profiles.Formula: savedUnits = unitsRange.Formula: savedSigns = signs.Formula
    TestConfigureFixture system: TestPrepareResults stats, system, profiles
    TestLogLine stats, "REAL_SNAPSHOT: " & Audit02ReadExportSnapshotForTests(ThisWorkbook)
    solves = SectionEquilibriumSolveCount()
    Set doc = acad.Documents.Add()
    TestWaitForCAD acad
    TestLogLine stats, "REAL_AUTOCAD: " & acad.FullName & "|version=" & acad.Version & "|HWND=" & expectedHwnd
    Set layer = doc.Layers.Add("AUDIT_C"): layer.Color = 7
    Audit03EnsureAutoCADLayerForTests doc, "AUDIT_C", 1
    TestCheck stats, "realAutoCAD.layer.existingColor", layer.Color = 7
    Audit03EnsureAutoCADLayerForTests doc, "AUDIT_ACI1", 1
    Audit03EnsureAutoCADLayerForTests doc, "AUDIT_ACI255", 255
    TestCheck stats, "realAutoCAD.layer.color1", doc.Layers.Item("AUDIT_ACI1").Color = 1
    TestCheck stats, "realAutoCAD.layer.color255", doc.Layers.Item("AUDIT_ACI255").Color = 255
    Dim aci As Long
    Set layer = doc.Layers.Item("AUDIT_ACI1")
    For aci = 1 To 255
        layer.Color = aci
        TestCheck stats, "realAutoCAD.layer.fullACI." & CStr(aci), layer.Color = aci
    Next aci
    Audit03EnsureAutoCADLayerForTests doc, "Слой бетона длинный", 4
    Audit03EnsureAutoCADLayerForTests doc, String$(255, "a"), 4
    TestCheck stats, "realAutoCAD.layer.unicode", doc.Layers.Item("Слой бетона длинный").Name = "Слой бетона длинный"
    TestCheck stats, "realAutoCAD.layer.length255", Len(doc.Layers.Item(String$(255, "a")).Name) = 255
    added = Audit03ExportAutoCADDocumentForTests(doc)
    TestCheck stats, "realAutoCAD.export.added", added > 16
    TestCheck stats, "realAutoCAD.export.regions", TestCountEntities(doc, "AcDbRegion") = 8
    TestCheck stats, "realAutoCAD.export.concrete", TestCountEntities(doc, "AcDbRegion", "AUDIT_C") = 4
    TestCheck stats, "realAutoCAD.export.rebar", TestCountEntities(doc, "AcDbRegion", "AUDIT_R") = 4
    TestCheck stats, "realAutoCAD.export.axes", TestCountEntities(doc, "", "RC_NDM_Axes") > 0
    TestCheck stats, "realAutoCAD.export.loadPoint", TestCountEntities(doc, "", "RC_NDM_LoadPoint") > 0
    TestCheck stats, "realAutoCAD.export.neutralLine", TestCountEntities(doc, "", "RC_NDM_NeutralLine") > 0
    TestCheck stats, "realAutoCAD.export.contour", TestCountEntities(doc, "", "AUDIT_CONTOUR") > 0
    TestCheckNativeColors stats, doc, system
    TestCheckNativeImport stats, doc, system
    For Each key In Array("NeutralLine", "LoadPoint", "Contour", "PrincipalAxes")
        TestClearOwnModelSpace doc
        If CStr(key) = "PrincipalAxes" Then
            TestSetValue system, "AutoCAD.Export.PrincipalAxesMode", "None"
        Else
            TestSetValue system, "AutoCAD.Export." & CStr(key) & "Enabled", "No"
        End If
        added = Audit03ExportAutoCADDocumentForTests(doc)
        Dim disabledLayer As String
        Select Case CStr(key)
            Case "Contour": disabledLayer = "AUDIT_CONTOUR"
            Case "PrincipalAxes": disabledLayer = "RC_NDM_Axes"
            Case Else: disabledLayer = "RC_NDM_" & CStr(key)
        End Select
        TestCheck stats, "realAutoCAD.disabled." & CStr(key), TestCountEntities(doc, "", disabledLayer) = 0
        TestCheck stats, "realAutoCAD.disabled.geometry." & CStr(key), TestCountEntities(doc, "AcDbRegion") = 8
        TestConfigureFixture system
    Next key
    For Each value In Array("ValuesOnly", "NamesAndValues")
        TestClearOwnModelSpace doc: TestSetValue system, "AutoCAD.Export.LabelMode", CStr(value)
        added = Audit03ExportAutoCADDocumentForTests(doc)
        TestCheckNativeLabels stats, doc, CStr(value)
    Next value
    ' Меняем только текущий Config после сохранения Results: размеры DWG
    ' должны остаться 100/200 мм, а не превратиться в метры или миллиметры повторно.
    TestSetValue unitsRange, "Length", "m": TestCellForKey(unitsRange, "Length", 4).Value2 = "m"
    TestSetValue unitsRange, "Area", "m2": TestCellForKey(unitsRange, "Area", 4).Value2 = "m2"
    TestSetValue signs, "+N", "Tension": TestSetValue signs, "+Mx", "-Y tension": TestSetValue signs, "+My", "-X tension"
    TestClearOwnModelSpace doc: added = Audit03ExportAutoCADDocumentForTests(doc)
    TestCheckNativeImport stats, doc, system
    TestCheckNativeColors stats, doc, system
    Dim presentationPath As String
    presentationPath = Left$(drawingPath, Len(drawingPath) - 4) & "_presentation.dwg"
    If Len(Dir$(presentationPath)) > 0 Then Err.Raise vbObjectError + 4499, "RunAudit03RealAutoCADTests", "Тестовый presentation-DWG уже существует; перезапись запрещена."
    doc.SaveAs presentationPath
    TestLogLine stats, "REAL_DWG_PRESENTATION: " & doc.FullName & "|entities=" & CStr(doc.ModelSpace.Count)
    Dim entity As Object
    For Each entity In doc.ModelSpace
        If entity.ObjectName = "AcDbRegion" Then
            Set copyRegion = entity.Copy(): copyRegion.Layer = "AUDIT_CT": Exit For
        End If
    Next entity
    Dim deleted As Long: deleted = Audit03CleanupAutoCADModelSpaceForTests(doc.ModelSpace, TestReader())
    TestCheck stats, "realAutoCAD.cleanup.deleted", deleted > 0
    TestCheck stats, "realAutoCAD.cleanup.allRegionsKept", TestCountEntities(doc, "AcDbRegion") = 9
    TestCheck stats, "realAutoCAD.cleanup.regionOnAnnotation", TestCountEntities(doc, "AcDbRegion", "AUDIT_CT") = 1
    TestCheck stats, "realAutoCAD.export.noSolve", SectionEquilibriumSolveCount() = solves
    doc.SaveAs drawingPath
    TestCheck stats, "realAutoCAD.saved", Len(Dir$(drawingPath)) > 0
    TestLogLine stats, "REAL_DWG: " & doc.FullName & "|entities=" & CStr(doc.ModelSpace.Count)
    doc.Close False: Set doc = Nothing
    TestLogLine stats, "REAL_STAGE: reopen saved DWG"
    Set doc = acad.Documents.Open(drawingPath)
    TestLogLine stats, "REAL_STAGE: wait for reopened document"
    TestWaitForCAD acad
    TestLogLine stats, "REAL_STAGE: verify reopened regions"
    TestCheck stats, "realAutoCAD.reopen.regions", TestCountEntities(doc, "AcDbRegion") = 9
    TestCheckNativeImport stats, doc, system
    TestLogLine stats, "REAL_STAGE: independent central inertia probes"
    TestCheckNativeRegionInertia stats, doc
    TestLogLine stats, "REAL_STAGE: imported shape -> Results -> exported Region"
    TestCheckNativeShapeRoundTrip stats, doc, system, unitsRange, drawingPath
    TestLogLine stats, "REAL_STAGE: square edge and circular Region average orientation"
    TestCheckNativeIsotropicRoundTrip stats, doc, system, unitsRange, drawingPath
    GoTo Restore
FailedRun:
    stats.Failed = stats.Failed + 1
    TestLogLine stats, "FAIL: realAutoCAD.runtime; " & CStr(Err.Number) & "; " & Err.Description
Restore:
    On Error GoTo FailedRestore
    If Not doc Is Nothing Then TestWaitForCAD acad: doc.Close False
    If Not system Is Nothing Then system.Formula = saved
    If Not profiles Is Nothing Then profiles.Formula = savedProfiles
    If Not unitsRange Is Nothing Then unitsRange.Formula = savedUnits
    If Not signs Is Nothing Then signs.Formula = savedSigns
    GoTo Finish
FailedRestore:
    stats.Failed = stats.Failed + 1: TestLogLine stats, "FAIL: realAutoCAD.restore; " & Err.Description
Finish:
    TestLogLine stats, "TOTAL_REAL_AUTOCAD: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunAudit03RealAutoCADTests = stats.Report
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Ждет завершения CAD-инициализации нового документа ограниченное время.
' Busy не становится ложным отказом геометрии; после timeout тест прекращается.
Private Sub TestWaitForCAD(ByVal acad As Object)
    Dim started As Single, elapsed As Double: started = Timer
    Do
        If acad.GetAcadState().IsQuiescent Then Exit Sub
        DoEvents: Application.Wait DateAdd("s", 1, Now)
        elapsed = Timer - started: If elapsed < 0# Then elapsed = elapsed + 86400#
        If elapsed > 30# Then Err.Raise vbObjectError + 4499, "WaitForCAD", "AutoCAD не завершил текущую команду за 30 секунд."
    Loop
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Считает настоящие сущности по COM ObjectName/Layer, не по flags Config.
Private Function TestCountEntities(ByVal doc As Object, Optional ByVal objectName As String = "", _
        Optional ByVal layerName As String = "") As Long
    Dim entity As Object
    For Each entity In doc.ModelSpace
        If (Len(objectName) = 0 Or StrComp(entity.ObjectName, objectName, vbTextCompare) = 0) And _
                (Len(layerName) = 0 Or StrComp(entity.Layer, layerName, vbTextCompare) = 0) Then TestCountEntities = TestCountEntities + 1
    Next entity
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Удаляет только содержимое собственного нового test-document между сценариями;
' этот helper не вызывается ни для пользовательского ActiveDocument, ни из UI.
Private Sub TestClearOwnModelSpace(ByVal doc As Object)
    Dim i As Long
    For i = doc.ModelSpace.Count - 1 To 0 Step -1: doc.ModelSpace.Item(i).Delete: Next i
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Читает экспортированные настоящие Region рабочим importer-ом. Контролирует
' точные мм/мм2 и раскладку вместо проверки только существования DWG-файла.
Private Sub TestCheckNativeImport(ByRef stats As TCadStats, ByVal doc As Object, ByVal system As Object)
    TestSetValue system, "AutoCAD.Common.ConcreteLayer", "AUDIT_C"
    TestSetValue system, "AutoCAD.Common.RebarLayer", "AUDIT_R"
    Dim importer As CAutoCADSectionModelImporter, model As CSectionModel, i As Long
    Set importer = New CAutoCADSectionModelImporter
    Set model = importer.ImportConfiguredModelSpace(doc.ModelSpace, TestReader())
    TestCheck stats, "realAutoCAD.import.counts", model.ConcreteCount = 4 And model.RebarCount = 4
    For i = 1 To model.ConcreteCount
        TestCheckClose stats, "realAutoCAD.import.concreteArea." & CStr(i), model.ConcreteArea(i), 10000#
        TestCheckClose stats, "realAutoCAD.import.concreteX." & CStr(i), Abs(model.ConcreteX(i)), 50#
        TestCheckClose stats, "realAutoCAD.import.concreteY." & CStr(i), Abs(model.ConcreteY(i)), 50#
    Next i
    For i = 1 To model.RebarCount
        TestCheckClose stats, "realAutoCAD.import.rebarArea." & CStr(i), model.RebarArea(i), GEOM_PI * 64#
        TestCheckClose stats, "realAutoCAD.import.rebarX." & CStr(i), Abs(model.RebarX(i)), 60#
        TestCheckClose stats, "realAutoCAD.import.rebarY." & CStr(i), Abs(model.RebarY(i)), 60#
    Next i
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет центральные площадь и моменты инерции повёрнутого Region.
' Перенос объекта не меняет центральные моменты инерции. Оба знака угла
' и произведения координат выявляют неверное восстановление Ixy.
Private Sub TestCheckNativeRegionInertia(ByRef stats As TCadStats, ByVal doc As Object)
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
            TestLogLine stats, "REAL_STAGE: create " & layerName
            Audit03EnsureAutoCADLayerForTests doc, layerName, 7
            Set region = TestNativeRectangleRegion(doc, 120#, 60#, angle, cx, cy)
            region.Layer = layerName
            raw = region.ProductOfInertia
            If IsArray(raw) Then rawProduct = CDbl(raw(LBound(raw))) Else rawProduct = CDbl(raw)
            prefix = "realAutoCAD.inertia." & CStr(angleSign) & "." & CStr(position)
            TestLogLine stats, "REAL_INERTIA: " & prefix & "|cx=" & CStr(cx) & "; cy=" & CStr(cy) & _
                "; rawProduct=" & TestNativeNumber(rawProduct) & "; expectedCentralIxy=" & TestNativeNumber(expectedIxy)
            Set model = importer.ImportFromModelSpace(doc.ModelSpace, layerName, "AUDIT_R", "A400", 0#)
            TestCheck stats, prefix & ".count", model.ConcreteCount = 1
            If model.ConcreteCount = 1 Then
                model.ConcreteLocalInertiaComponents 1, ix, iy, ixy
                TestCheckNativeInertiaClose stats, prefix & ".area", model.ConcreteArea(1), 7200#
                TestCheckNativeInertiaClose stats, prefix & ".x", model.ConcreteX(1), cx
                TestCheckNativeInertiaClose stats, prefix & ".y", model.ConcreteY(1), cy
                TestCheckNativeInertiaClose stats, prefix & ".Ix", ix, expectedIx
                TestCheckNativeInertiaClose stats, prefix & ".Iy", iy, expectedIy
                TestCheckNativeInertiaClose stats, prefix & ".Ixy", ixy, expectedIxy
            End If
        Next position
    Next angleSign
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Создает тестовый прямоугольник по независимо повернутым четырем вершинам.
' AddRegion получает замкнутый Polyline; временная исходная линия удаляется
' только из собственного документа теста после создания Region.
Private Function TestNativeRectangleRegion(ByVal doc As Object, ByVal width As Double, ByVal height As Double, _
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
    Set TestNativeRectangleRegion = regions(LBound(regions))
    line.Delete
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Отделяет численную погрешность CAD-интеграции от инженерного расхождения.
' Абсолютный допуск 1e-6 фиксирован до запуска и не меняет независимые эталоны ядра.
Private Sub TestCheckNativeInertiaClose(ByRef stats As TCadStats, ByVal name As String, ByVal actual As Double, ByVal expected As Double)
    TestLogLine stats, "AUTOCAD_INERTIA_NUMBER: " & name & "|actual=" & TestNativeNumber(actual) & _
        "|expected=" & TestNativeNumber(expected) & "|tolerance=0.000001"
    TestCheck stats, name, Abs(actual - expected) <= 0.000001
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сохраняет значение CAD в журнале без зависимости от десятичного разделителя
' Excel. Этот helper только форматирует диагностику, не округляет oracle.
Private Function TestNativeNumber(ByVal value As Double) As String
    TestNativeNumber = Replace$(CStr(value), ",", ".")
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет сохранение и экспорт повёрнутого прямоугольника, треугольника
' и Г-образного Region. Эталон независимо вычисляется по исходному многоугольнику,
' а не по восстановленным размерам модели. Оба направления поворота проверяются
' с выходными единицами мм и м в собственных тестовых чертежах.
Private Sub TestCheckNativeShapeRoundTrip(ByRef stats As TCadStats, ByVal doc As Object, _
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
    TestConfigureFixture system
    TestSetValue system, "AutoCAD.Export.PrincipalAxesMode", "None"
    TestSetValue system, "AutoCAD.Export.NeutralLineEnabled", "No"
    TestSetValue system, "AutoCAD.Export.LoadPointEnabled", "No"
    TestSetValue system, "AutoCAD.Export.ContourEnabled", "No"
    For unitMode = 0 To 1
        If unitMode = 0 Then
            TestSetValue unitsRange, "Length", "mm": TestCellForKey(unitsRange, "Length", 4).Value2 = "mm"
            TestSetValue unitsRange, "Area", "mm2": TestCellForKey(unitsRange, "Area", 4).Value2 = "mm2"
        Else
            TestSetValue unitsRange, "Length", "m": TestCellForKey(unitsRange, "Length", 4).Value2 = "m"
            TestSetValue unitsRange, "Area", "m2": TestCellForKey(unitsRange, "Area", 4).Value2 = "m2"
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
                TestLogLine stats, "REAL_STAGE: " & prefix
                TestClearOwnModelSpace doc
                angle = CDbl(signValue) * GEOM_PI / 6#
                ReDim x(0 To UBound(px)): ReDim y(0 To UBound(py))
                For i = 0 To UBound(px)
                    x(i) = 20# + CDbl(px(i)) * Cos(angle) - CDbl(py(i)) * Sin(angle)
                    y(i) = -10# + CDbl(px(i)) * Sin(angle) + CDbl(py(i)) * Cos(angle)
                Next i
                TestNativePolygonInertia x, y, area, cx, cy, ix, iy, ixy
                Audit03EnsureAutoCADLayerForTests doc, "AUDIT_SHAPE_IN", 7
                Audit03EnsureAutoCADLayerForTests doc, "AUDIT_SHAPE_R", 7
                Set region = TestNativePolygonRegion(doc, x, y): region.Layer = "AUDIT_SHAPE_IN"
                For i = -1 To 1 Step 2
                    For j = 0 To 1
                        actualX = 20# + 5# * i * Cos(angle) - (-25# + 10# * j) * Sin(angle)
                        actualY = -10# + 5# * i * Sin(angle) + (-25# + 10# * j) * Cos(angle)
                        Set bar = TestNativeRectangleRegion(doc, 2#, 2#, angle, actualX, actualY)
                        bar.Layer = "AUDIT_SHAPE_R"
                    Next j
                Next i
                Set model = importer.ImportFromModelSpace(doc.ModelSpace, "AUDIT_SHAPE_IN", "AUDIT_SHAPE_R", "A400", 0#)
                TestCheckNativeInertiaClose stats, prefix & ".import.area", model.ConcreteArea(1), area
                TestCheckNativeInertiaClose stats, prefix & ".import.Ix", model.ConcreteLocalIx(1), ix
                TestCheckNativeInertiaClose stats, prefix & ".import.Iy", model.ConcreteLocalIy(1), iy
                TestCheckNativeInertiaClose stats, prefix & ".import.Ixy", model.ConcreteLocalIxy(1), ixy
                If shape = 0 Then
                    TestCheck stats, prefix & ".sourceRectangle", model.ConcreteShapeType(1) = "Rectangle"
                Else
                    TestCheck stats, prefix & ".sourceEquivalent", model.ConcreteShapeType(1) = "Equivalent rectangle"
                End If
                TestPrepareNativeImportedSnapshot stats, system, model, prefix
                solves = SectionEquilibriumSolveCount()
                added = Audit03ExportAutoCADDocumentForTests(doc)
                TestCheck stats, prefix & ".export.noSolve", SectionEquilibriumSolveCount() = solves
                TestCheck stats, prefix & ".export.added", added >= 10
                TestCheck stats, prefix & ".export.regionCount", TestCountEntities(doc, "AcDbRegion", "AUDIT_C") = 1
                Set exported = Nothing
                For Each entity In doc.ModelSpace
                    If entity.ObjectName = "AcDbRegion" And entity.Layer = "AUDIT_C" Then Set exported = entity: Exit For
                Next entity
                If exported Is Nothing Then Err.Raise vbObjectError + 4499, "CheckNativeShapeRoundTrip", "Экспортированный бетонный Region не найден."
                TestReadNativeCentralInertia exported, actualArea, actualX, actualY, actualIx, actualIy, actualIxy
                TestCheckNativeInertiaClose stats, prefix & ".export.area", actualArea, area
                TestCheckNativeInertiaClose stats, prefix & ".export.x", actualX, cx
                TestCheckNativeInertiaClose stats, prefix & ".export.y", actualY, cy
                TestCheckNativePrincipalShape stats, prefix, ix, iy, ixy, actualIx, actualIy, actualIxy
                If shape = 0 Then
                    TestCheckNativeInertiaClose stats, prefix & ".rectangle.Ix", actualIx, ix
                    TestCheckNativeInertiaClose stats, prefix & ".rectangle.Iy", actualIy, iy
                    TestCheckNativeInertiaClose stats, prefix & ".rectangle.Ixy", actualIxy, ixy
                End If
                TestCheckNativeRectangleEdges stats, prefix, exported
                path = Left$(drawingPath, Len(drawingPath) - 4) & "_shape_" & CStr(unitMode) & "_" & CStr(shape) & "_" & CStr(signValue) & ".dwg"
                If Len(Dir$(path)) > 0 Then Err.Raise vbObjectError + 4499, "CheckNativeShapeRoundTrip", "Тестовый DWG формы уже существует; перезапись запрещена."
                doc.SaveAs path
                TestLogLine stats, "REAL_SHAPE_DWG: " & doc.FullName & "|shape=" & model.ConcreteShapeType(1) & _
                    "; sourceArea=" & TestNativeNumber(area) & "; exportArea=" & TestNativeNumber(actualArea) & _
                    "; rotation=" & TestNativeNumber(model.ConcreteRotation(1))
            Next signValue
        Next shape
    Next unitMode
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет изотропные Region на реальных сущностях: квадрат берет направление
' своей прямой грани, круг берет среднее остальных ориентированных элементов.
' Oracle читает исходные ребра отдельно от importer-а; круг не участвует в среднем.
Private Sub TestCheckNativeIsotropicRoundTrip(ByRef stats As TCadStats, ByVal doc As Object, _
        ByVal system As Object, ByVal unitsRange As Object, ByVal drawingPath As String)
    Dim scenario As Long, unitMode As Long, i As Long, j As Long, circleIndex As Long
    Dim squareIndex As Long, squareAngle As Double, expected As Double, angle As Double, circleX As Double
    Dim sumCos As Double, sumSin As Double, region As Object, circleRegion As Object, bar As Object
    Dim model As CSectionModel, importer As CAutoCADSectionModelImporter, entity As Object
    Dim exportedCircle As Object, exportedSquare As Object, added As Long, solves As Long
    Dim prefix As String, path As String, squareHandle As String, circleHandle As String
    Dim area As Double, cx As Double, cy As Double, ix As Double, iy As Double, ixy As Double
    TestConfigureFixture system
    TestSetValue system, "AutoCAD.Export.PrincipalAxesMode", "None"
    TestSetValue system, "AutoCAD.Export.NeutralLineEnabled", "No"
    TestSetValue system, "AutoCAD.Export.LoadPointEnabled", "No"
    TestSetValue system, "AutoCAD.Export.ContourEnabled", "No"
    Set importer = New CAutoCADSectionModelImporter
    For unitMode = 0 To 1
        If unitMode = 0 Then
            TestSetValue unitsRange, "Length", "mm": TestCellForKey(unitsRange, "Length", 4).Value2 = "mm"
            TestSetValue unitsRange, "Area", "mm2": TestCellForKey(unitsRange, "Area", 4).Value2 = "mm2"
        Else
            TestSetValue unitsRange, "Length", "m": TestCellForKey(unitsRange, "Length", 4).Value2 = "m"
            TestSetValue unitsRange, "Area", "m2": TestCellForKey(unitsRange, "Area", 4).Value2 = "m2"
        End If
        For scenario = 0 To 4
            prefix = "realAutoCAD.isotropic." & CStr(unitMode) & "." & CStr(scenario)
            TestLogLine stats, "REAL_STAGE: " & prefix
            TestClearOwnModelSpace doc
            sumCos = 0#: sumSin = 0#: squareHandle = vbNullString
            If scenario <= 1 Then
                angle = GEOM_PI / 6#: If scenario = 1 Then angle = 0#
                Set region = TestNativeRectangleRegion(doc, 60#, 60#, angle, 0#, 0#)
                region.Layer = "AUDIT_SHAPE_IN": squareHandle = region.Handle
                squareAngle = TestNativeFirstEdgeAngle(region)
                sumCos = Cos(2# * squareAngle): sumSin = Sin(2# * squareAngle)
                If scenario = 1 Then
                    Set region = TestNativeRectangleRegion(doc, 60#, 120#, GEOM_PI / 6#, 200#, 0#)
                    region.Layer = "AUDIT_SHAPE_IN"
                    sumCos = sumCos + Cos(GEOM_PI / 3#): sumSin = sumSin + Sin(GEOM_PI / 3#)
                End If
            ElseIf scenario = 2 Or scenario = 4 Then
                angle = 89# * GEOM_PI / 180#: If scenario = 4 Then angle = GEOM_PI / 4#
                For i = -1 To 1 Step 2
                    Set region = TestNativeRectangleRegion(doc, 60#, 120#, CDbl(i) * angle, 100# + 100# * i, 0#)
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
            Set circleRegion = TestNativeCircleRegion(doc, circleX, 0#, 20#)
            circleRegion.Layer = "AUDIT_SHAPE_IN": circleHandle = circleRegion.Handle
            For i = -1 To 1 Step 2
                For j = -1 To 1 Step 2
                    Set bar = TestNativeRectangleRegion(doc, 2#, 2#, 0#, circleX + 5# * i, 5# * j)
                    bar.Layer = "AUDIT_SHAPE_R"
                Next j
            Next i
            Set model = importer.ImportFromModelSpace(doc.ModelSpace, "AUDIT_SHAPE_IN", "AUDIT_SHAPE_R", "A400", 0#)
            circleIndex = 0: squareIndex = 0
            For i = 1 To model.ConcreteCount
                If model.ConcreteSourceHandle(i) = circleHandle Then circleIndex = i
                If Len(squareHandle) > 0 And model.ConcreteSourceHandle(i) = squareHandle Then squareIndex = i
            Next i
            TestCheck stats, prefix & ".circle.found", circleIndex > 0
            If circleIndex = 0 Then Err.Raise vbObjectError + 4499, "CheckNativeIsotropicRoundTrip", "Круговой Region потерян при импорте."
            If squareIndex > 0 Then TestCheckNativeDirection stats, prefix & ".square.import", model.ConcreteRotation(squareIndex), squareAngle, 4#
            TestCheckNativeDirection stats, prefix & ".average.sources", model.AverageKnownConcreteElementRotation(), expected, 2#
            model.ApplyAverageRotationToEquivalentAreaFallbacks
            TestCheckNativeDirection stats, prefix & ".circle.assigned", model.ConcreteRotation(circleIndex), expected, 2#
            model.ApplyAverageRotationToEquivalentAreaFallbacks
            TestCheckNativeDirection stats, prefix & ".circle.repeated", model.ConcreteRotation(circleIndex), expected, 2#
            If squareIndex > 0 Then TestCheckNativeDirection stats, prefix & ".square.retained", model.ConcreteRotation(squareIndex), squareAngle, 4#
            TestPrepareNativeImportedSnapshot stats, system, model, prefix
            solves = SectionEquilibriumSolveCount(): added = Audit03ExportAutoCADDocumentForTests(doc)
            TestCheck stats, prefix & ".export.noSolve", SectionEquilibriumSolveCount() = solves
            TestCheck stats, prefix & ".export.regionCount", TestCountEntities(doc, "AcDbRegion", "AUDIT_C") = model.ConcreteCount
            Set exportedCircle = Nothing: Set exportedSquare = Nothing
            For Each entity In doc.ModelSpace
                If entity.ObjectName = "AcDbRegion" And entity.Layer = "AUDIT_C" Then
                    TestReadNativeCentralInertia entity, area, cx, cy, ix, iy, ixy
                    If Abs(cx - circleX) < 0.000001 And Abs(cy) < 0.000001 Then Set exportedCircle = entity
                    If Abs(cx) < 0.000001 And Abs(cy) < 0.000001 Then Set exportedSquare = entity
                End If
            Next entity
            TestCheck stats, prefix & ".circle.exported", Not exportedCircle Is Nothing
            If exportedCircle Is Nothing Then Err.Raise vbObjectError + 4499, "CheckNativeIsotropicRoundTrip", "Оболочка круга потеряна при экспорте."
            TestReadNativeCentralInertia exportedCircle, area, cx, cy, ix, iy, ixy
            TestCheckNativeInertiaClose stats, prefix & ".circle.export.area", area, GEOM_PI * 400#
            TestCheckNativeDirection stats, prefix & ".circle.export.edge", TestNativeFirstEdgeAngle(exportedCircle), expected, 4#
            TestCheckNativeRectangleEdges stats, prefix & ".circle.export", exportedCircle
            If squareIndex > 0 Then
                TestCheck stats, prefix & ".square.exported", Not exportedSquare Is Nothing
                If exportedSquare Is Nothing Then Err.Raise vbObjectError + 4499, "CheckNativeIsotropicRoundTrip", "Оболочка квадрата потеряна при экспорте."
                TestReadNativeCentralInertia exportedSquare, area, cx, cy, ix, iy, ixy
                TestCheckNativeInertiaClose stats, prefix & ".square.export.area", area, 3600#
                TestCheckNativeDirection stats, prefix & ".square.export.edge", TestNativeFirstEdgeAngle(exportedSquare), squareAngle, 4#
                TestCheckNativeRectangleEdges stats, prefix & ".square.export", exportedSquare
            End If
            path = Left$(drawingPath, Len(drawingPath) - 4) & "_isotropic_" & CStr(unitMode) & "_" & CStr(scenario) & ".dwg"
            If Len(Dir$(path)) > 0 Then Err.Raise vbObjectError + 4499, "CheckNativeIsotropicRoundTrip", "Тестовый DWG уже существует; перезапись запрещена."
            doc.SaveAs path
            TestLogLine stats, "REAL_ISOTROPIC_DWG: " & doc.FullName & "|expected=" & TestNativeNumber(expected) & "; squareEdge=" & TestNativeNumber(squareAngle)
        Next scenario
    Next unitMode
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Создает круговой Region без прямых граней, чтобы проверять именно средний
' угол сетки, а не подсказку из полилинии. Все исходные кривые принадлежат тесту.
Private Function TestNativeCircleRegion(ByVal doc As Object, ByVal x As Double, ByVal y As Double, ByVal radius As Double) As Object
    Dim center(0 To 2) As Double, circleEntity As Object, curves(0 To 0) As Object, regions As Variant
    center(0) = x: center(1) = y
    Set circleEntity = doc.ModelSpace.AddCircle(center, radius): Set curves(0) = circleEntity
    regions = doc.ModelSpace.AddRegion(curves): Set TestNativeCircleRegion = regions(LBound(regions))
    circleEntity.Delete
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Независимо читает угол первого прямого ребра настоящего Region; не вызывает
' importer. Копия и результаты Explode удаляются из собственного документа.
Private Function TestNativeFirstEdgeAngle(ByVal region As Object) As Double
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
    TestNativeFirstEdgeAngle = angle
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает направления без ложного отказа из-за эквивалентных углов:
' multiplier=2 для оси, multiplier=4 для взаимозаменяемых граней квадрата.
Private Sub TestCheckNativeDirection(ByRef stats As TCadStats, ByVal prefix As String, _
        ByVal actual As Double, ByVal expected As Double, ByVal multiplier As Double)
    TestCheckNativeInertiaClose stats, prefix & ".cos", Cos(multiplier * actual), Cos(multiplier * expected)
    TestCheckNativeInertiaClose stats, prefix & ".sin", Sin(multiplier * actual), Sin(multiplier * expected)
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Рассчитывает допустимое НДС от малой осевой нагрузки и сохраняет импортированную
' модель обычным классом записи Results. Экспорт читает эти сохранённые результаты;
' готовые строки результатов в тесте не подставляются.
Private Sub TestPrepareNativeImportedSnapshot(ByRef stats As TCadStats, ByVal system As Object, _
        ByVal model As CSectionModel, ByVal prefix As String)
    Dim settings As CSystemSettingsReader, units As CUnitSystem, provider As CMaterialModelProvider
    Dim batch As CBatchSectionCalculator, catalog As CCalculationProfileCatalog, writer As CNDMResultsWriter
    Dim referenceX As Double, referenceY As Double
    modTestPlotConfig.TestSetOutputCombinationInTable system, "CAD_SHAPE"
    Set settings = TestReader(): Set units = New CUnitSystem: units.LoadFromSettings settings
    Set provider = New CMaterialModelProvider: provider.Initialize settings, units
    Set catalog = New CCalculationProfileCatalog: catalog.LoadFromWorkbook ThisWorkbook
    Set batch = New CBatchSectionCalculator: batch.Initialize model, provider
    Set batch.ProfileCatalog = catalog: batch.ApplySettings settings, units
    CalculateConcreteSectionCentroid model, referenceX, referenceY
    batch.ApplyLoadReference referenceX, referenceY, referenceX, referenceY
    batch.AddCombination "CAD_SHAPE", -1000#, 0#, 0#, "PR1", "Native imported shape", "Auto"
    batch.Execute
    TestLogLine stats, "REAL_SHAPE_STATE: " & prefix & "|" & batch.ResultAt(1).OverallMeta.ResultComment
    TestCheck stats, prefix & ".snapshot.state", batch.ResultAt(1).DirectStateMeta.InternalStatus = rsSuccess
    If batch.ResultAt(1).DirectStateMeta.InternalStatus <> rsSuccess Then _
        Err.Raise vbObjectError + 4499, "PrepareNativeImportedSnapshot", "Не получено допустимое НДС импортированной формы: " & batch.ResultAt(1).DirectStateMeta.ResultComment
    Set writer = New CNDMResultsWriter: writer.WriteResults ThisWorkbook, model, PrepareSectionSnapshot(model, provider), batch, units
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Создает фактический замкнутый polygon Region в WCS-плоскости собственного
' test-document. Исходная Polyline после преобразования не остается в ModelSpace.
Private Function TestNativePolygonRegion(ByVal doc As Object, ByRef x() As Double, ByRef y() As Double) As Object
    Dim points() As Double, i As Long, line As Object, curves(0 To 0) As Object, regions As Variant
    ReDim points(0 To 2 * (UBound(x) + 1) - 1)
    For i = 0 To UBound(x): points(2 * i) = x(i): points(2 * i + 1) = y(i): Next i
    Set line = doc.ModelSpace.AddLightWeightPolyline(points): line.Closed = True
    Set curves(0) = line: regions = doc.ModelSpace.AddRegion(curves)
    Set TestNativePolygonRegion = regions(LBound(regions))
    line.Delete
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Независимая квадратура простого polygon по ориентированным ребрам.
' Сначала получаем интегралы относительно начала, затем снимаем перенос
' к центру площади. Здесь нет чтения CSectionModel или CAD MomentOfInertia.
Private Sub TestNativePolygonInertia(ByRef x() As Double, ByRef y() As Double, _
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

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Читает центральный тензор непосредственно из настоящего Region, независимо
' от importer-а. Знак отрицательной XY-компоненты AutoCAD подтвержден отдельным
' native oracle; геометрические значения здесь всегда миллиметровые.
Private Sub TestReadNativeCentralInertia(ByVal region As Object, ByRef area As Double, _
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

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает отношение собственных инерций и беззнаковое направление оси.
' Нормализованные (Ix-Iy, -2Ixy) задают cos(2a)/sin(2a), поэтому проверка
' допускает ту же ось плюс 180 градусов, но не ее зеркальное направление.
Private Sub TestCheckNativePrincipalShape(ByRef stats As TCadStats, ByVal prefix As String, _
        ByVal ix As Double, ByVal iy As Double, ByVal ixy As Double, _
        ByVal ex As Double, ByVal ey As Double, ByVal exy As Double)
    Dim delta As Double, exportDelta As Double, ratio As Double, exportRatio As Double
    delta = Sqr((ix - iy) ^ 2 + 4# * ixy ^ 2)
    exportDelta = Sqr((ex - ey) ^ 2 + 4# * exy ^ 2)
    ratio = (ix + iy + delta) / (ix + iy - delta)
    exportRatio = (ex + ey + exportDelta) / (ex + ey - exportDelta)
    TestCheckNativeInertiaClose stats, prefix & ".principal.ratio", exportRatio, ratio
    TestCheckNativeInertiaClose stats, prefix & ".principal.cos2", (ex - ey) / exportDelta, (ix - iy) / delta
    TestCheckNativeInertiaClose stats, prefix & ".principal.sin2", -2# * exy / exportDelta, -2# * ixy / delta
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет экспортированную оболочку как четыре прямых взаимно перпендикулярных
' ребра. Даже для polygon исходника export должен дать rectangle, не исходный
' triangle/L-contour. Временные Explode-сущности удаляются только из own DWG.
Private Sub TestCheckNativeRectangleEdges(ByRef stats As TCadStats, ByVal prefix As String, ByVal region As Object)
    Dim copy As Object, pieces As Variant, line As Variant, p As Variant, q As Variant
    Dim vx As Double, vy As Double, wx As Double, wy As Double, first As Boolean, count As Long
    Set copy = region.Copy(): pieces = copy.Explode
    For Each line In pieces
        count = count + 1
        TestCheck stats, prefix & ".rectangle.line." & CStr(count), line.ObjectName = "AcDbLine"
        If line.ObjectName = "AcDbLine" Then
            p = line.StartPoint: q = line.EndPoint
            wx = q(0) - p(0): wy = q(1) - p(1)
            If first Then
                TestCheck stats, prefix & ".rectangle.edgeDirection." & CStr(count), _
                    Abs(vx * wy - vy * wx) <= 0.000001 Or Abs(vx * wx + vy * wy) <= 0.000001
            Else
                vx = wx: vy = wy: first = True
            End If
        End If
        line.Delete
    Next line
    copy.Delete
    TestCheck stats, prefix & ".rectangle.fourEdges", count = 4
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает центральное сжатие одного круга после переноса координат модели.
' Использует обычные Config, профили и расчёт сочетаний; точку нагрузки задаёт
' в бетонном центре. Сохраняет заданную нагрузку и фактический результат с причиной отказа.
Public Function RunAudit03CircleTranslationStateTests() As String
    Dim stats As TCadStats, system As Object, profiles As Object, saved As Variant, savedProfiles As Variant
    Dim settings As CSystemSettingsReader, units As CUnitSystem, provider As CMaterialModelProvider
    Dim catalog As CCalculationProfileCatalog, batch As CBatchSectionCalculator, model As CSectionModel
    Dim state As CSectionStateResult, report As CExecutionReport, method As Variant, center As Variant
    Dim i As Long, j As Long, area As Double, prefix As String, result As CCombinationResult
    Dim referenceX As Double, referenceY As Double, baselineStrain As Double
    On Error GoTo FailedRun
    Set system = ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange
    Set profiles = ThisWorkbook.Names.Item("rngCalculationProfiles").RefersToRange
    saved = system.Formula: savedProfiles = profiles.Formula
    TestConfigureFixture system: TestPrepareResults stats, system, profiles
    TestSetValue system, "General.ExecutionReportEnabled", "Yes"
    For Each method In Array("Newton", "Secant")
        TestSetValue system, "Solver.Method", CStr(method)
        For Each center In Array(0#, 20#, 400#, 10000#)
            prefix = "audit03.circleTranslation." & CStr(method) & "." & TestNativeNumber(CDbl(center))
            Set model = New CSectionModel: model.SourceType = "AutoCADImport"
            area = GEOM_PI * 400#
            model.AddConcreteElement CDbl(center), 0#, area, 1, "circle", "CIRCLE", "Region", _
                0#, 0#, 0#, vbNullString, GEOM_PI * 20# ^ 4 / 4#, GEOM_PI * 20# ^ 4 / 4#, 0#
            For i = -1 To 1 Step 2
                For j = -1 To 1 Step 2
                    model.AddRebarElement CDbl(center) + 5# * i, 5# * j, Sqr(16# / GEOM_PI), 4#, "A400"
                Next j
            Next i
            Set settings = TestReader(): Set units = New CUnitSystem: units.LoadFromSettings settings
            Set provider = New CMaterialModelProvider: provider.Initialize settings, units
            Set catalog = New CCalculationProfileCatalog: catalog.LoadFromWorkbook ThisWorkbook
            Set report = New CExecutionReport: report.Initialize ThisWorkbook, settings
            Set batch = New CBatchSectionCalculator: batch.Initialize model, provider
            Set batch.ExecutionReport = report: Set batch.ProfileCatalog = catalog
            batch.ApplySettings settings, units
            CalculateConcreteSectionCentroid model, referenceX, referenceY
            batch.ApplyLoadReference referenceX, referenceY, referenceX, referenceY
            batch.AddCombination "CAD_CIRCLE", -1000#, 0#, 0#, "PR1", "Translation invariant circle", "Auto"
            batch.Execute
            Set result = batch.ResultAt(1)
            Set state = result.StrengthResult.DirectState.StateResult
            TestCheck stats, prefix & ".statePresent", Not state Is Nothing
            If state Is Nothing Then Err.Raise vbObjectError + 4499, "RunAudit03CircleTranslationStateTests", "Не возвращен диагностический State."
            TestLogLine stats, "CIRCLE_TRANSLATION_RESULT: " & prefix & "|status=" & result.StrengthResult.DirectState.Status & _
                "; internal=" & CStr(state.InternalStatus) & "; code=" & CStr(state.ResultCode) & _
                "; comment=" & state.ResultMeta.ResultComment & "; eps0=" & TestNativeNumber(state.Epsilon0) & _
                "; kx=" & TestNativeNumber(state.KappaX) & "; ky=" & TestNativeNumber(state.KappaY) & _
                "; targetN=" & TestNativeNumber(state.TargetN) & "; targetMx=" & TestNativeNumber(state.TargetMx) & _
                "; targetMy=" & TestNativeNumber(state.TargetMy)
            TestLogLine stats, "CIRCLE_TRANSLATION_DIAGNOSTIC: " & prefix & vbCrLf & state.DiagnosticLog
            TestCheck stats, prefix & ".success", state.InternalStatus = rsSuccess
            If state.Converged Then
                TestCheckNativeInertiaClose stats, prefix & ".N", state.Nint, -1000#
                TestCheckNativeInertiaClose stats, prefix & ".Mx", state.Mxint, 0#
                TestCheckNativeInertiaClose stats, prefix & ".My", state.Myint, -1000# * CDbl(center)
                If CDbl(center) = 0# Then baselineStrain = state.Epsilon0
                TestCheckNativeInertiaClose stats, prefix & ".centerStrain", _
                    state.Epsilon0 + state.KappaY * CDbl(center), baselineStrain
                TestCheckNativeInertiaClose stats, prefix & ".kx", state.KappaX, 0#
                TestCheckNativeInertiaClose stats, prefix & ".ky", state.KappaY, 0#
            End If
        Next center
    Next method
    GoTo Restore
FailedRun:
    stats.Failed = stats.Failed + 1: TestLogLine stats, "FAIL: audit03.circleTranslation.runtime; " & CStr(Err.Number) & "; " & Err.Description
Restore:
    On Error GoTo FailedRestore
    If Not system Is Nothing Then system.Formula = saved
    If Not profiles Is Nothing Then profiles.Formula = savedProfiles
    GoTo Finish
FailedRestore:
    stats.Failed = stats.Failed + 1: TestLogLine stats, "FAIL: audit03.circleTranslation.restore; " & Err.Description
Finish:
    TestLogLine stats, "TOTAL_CIRCLE_TRANSLATION: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunAudit03CircleTranslationStateTests = stats.Report
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет текст именно экспортных подписей элементов; осевые подписи
' и маркеры нагрузки не включаются в количество восемь расчетных labels.
Private Sub TestCheckNativeLabels(ByRef stats As TCadStats, ByVal doc As Object, ByVal mode As String)
    Dim entity As Object, count As Long, text As String, hasName As Boolean
    For Each entity In doc.ModelSpace
        If entity.ObjectName = "AcDbText" And Left$(entity.Layer, 6) = "AUDIT_" Then
            text = entity.TextString: count = count + 1
            hasName = (Left$(text, 1) = "C" Or Left$(text, 1) = "R")
            TestCheck stats, "realAutoCAD.label." & mode & "." & CStr(count), hasName = (mode = "NamesAndValues")
        End If
    Next entity
    TestCheck stats, "realAutoCAD.label.count." & mode, count = 8
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает цвета настоящих Region и текстов с физическим состоянием из Results
' и текущими пятью цветами ACI. Смена пользовательских знаков N, Mx и My
' не должна превращать цвет физического растяжения в цвет сжатия.
Private Sub TestCheckNativeColors(ByRef stats As TCadStats, ByVal doc As Object, ByVal system As Object)
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
            TestCheck stats, "realAutoCAD.color." & elementID & "." & entity.ObjectName, _
                entity.Color = CLng(TestCellForKey(system, "AutoCAD.Color." & colorKey, 2).Value2)
        End If
    Next entity
End Sub
