Attribute VB_Name = "modTestAudit02Snapshots"
Option Explicit

' ==========================================================================
' ДЛЯ ТЕСТОВ: сравнение построения и направленные расчёты всех групп сечений.
' ==========================================================================
#If VBA7 Then
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef value As Currency) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef value As Currency) As Long
#Else
Private Declare Function QueryPerformanceCounter Lib "kernel32" (ByRef value As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (ByRef value As Currency) As Long
#End If
Private Type TDoubleValue
    Value As Double
End Type
Private Type TDoubleWords
    Low As Long
    High As Long
End Type

' -------------------- ДЛЯ ТЕСТОВ --------------------
' один прогрев и пять замеров полного построения модели каждого варианта геометрии.
Public Function TestDumpAudit02BuildPerformance() As String
    Dim config As Object, saved As Variant, fixture As Variant, i As Long, result As String
    Dim settings As CSystemSettingsReader, units As CUnitSystem, model As CSectionModel
    Dim started As Currency, finished As Currency, frequency As Currency
    On Error GoTo Failed
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = TestFixtureSaveInputs(config)
    If QueryPerformanceFrequency(frequency) = 0 Then Err.Raise vbObjectError + 5490, , "Недоступен точный таймер."
    For Each fixture In Array("Circle", "Rectangle", "LSection", "TwoRectangles", "RoundedSimple", _
            "RoundedTapered", "RoundedMixed", "RoundedMixedLeft", "HollowCentered", "HollowOffset")
        TestFixtureRestoreInputs config, saved: modTestShapeArchitecture.TestConfigureShapeFixture CStr(fixture)
        Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
        Set units = New CUnitSystem: units.LoadFromSettings settings
        Set model = BuildWorkbookSectionModel(ThisWorkbook, settings, units)
        For i = 1 To 5
            QueryPerformanceCounter started
            Set model = BuildWorkbookSectionModel(ThisWorkbook, settings, units)
            QueryPerformanceCounter finished
            result = result & CStr(fixture) & "|" & CStr(i) & "|" & CStr((finished - started) / frequency) & _
                "|" & CStr(model.ConcreteCount) & "|" & CStr(model.RebarCount) & vbCrLf
        Next i
    Next fixture
    GoTo Finished
Failed:
    result = result & "ERROR: " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Finished:
    If Not config Is Nothing And Not IsEmpty(saved) Then TestFixtureRestoreInputs config, saved
    TestDumpAudit02BuildPerformance = result
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' два сочетания четырёх групп сечений с НДС, несущей способностью,
' устойчивостью и трещинами. Сохраняет все биты чисел Double, статусы
' и пояснения отдельно. Настройку продолжения диаграмм не меняет.
Public Function TestDumpAudit02Calculations(Optional ByVal crackCode As String = "SP35") As String
    Dim config As Object, saved As Variant, fixture As Variant, result As String
    Dim profiles As Object, row As Long, key As String, settings As CSystemSettingsReader, units As CUnitSystem
    Dim model As CSectionModel, properties As CSectionPropertiesCalculator, provider As CMaterialModelProvider
    Dim catalog As CCalculationProfileCatalog, batch As CBatchSectionCalculator, item As CCombinationResult
    Dim minX As Double, maxX As Double, minY As Double, maxY As Double, i As Long, meta As CResultMeta
    On Error GoTo Failed
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = TestFixtureSaveInputs(config)
    For Each fixture In Array("Circle", "LSection", "RoundedTapered", "HollowOffset")
        TestFixtureRestoreInputs config, saved
        modTestShapeArchitecture.TestConfigureShapeFixture "Circle"
        modTestShapeArchitecture.TestConfigureShapeFixture CStr(fixture)
        Set profiles = ThisWorkbook.Names("rngCalculationProfiles").RefersToRange
        For row = 2 To profiles.Rows.Count
            key = CStr(profiles.Cells(row, 2).Value2)
            Select Case key
                Case "Calculation.Strength.DirectState", "Calculation.Strength.Capacity", "Calculation.Stability.Enabled", _
                     "Calculation.Crack.Width", "Calculation.Crack.Longitudinal"
                    profiles.Cells(row, 3).Value2 = "Yes"
            End Select
        Next row
        TestSetSetting "SLS.Crack.Code", crackCode
        Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
        Set units = New CUnitSystem: units.LoadFromSettings settings
        Set model = BuildWorkbookSectionModel(ThisWorkbook, settings, units)
        Set properties = New CSectionPropertiesCalculator: properties.PrepareSnapshot model
        properties.CalculateProjection model, 1#, 0#, False, minX, maxX
        properties.CalculateProjection model, 0#, 1#, False, minY, maxY
        Set provider = New CMaterialModelProvider: provider.Initialize settings, units
        Set catalog = New CCalculationProfileCatalog: catalog.LoadFromWorkbook ThisWorkbook
        Set batch = New CBatchSectionCalculator: batch.Initialize model, provider
        Set batch.ProfileCatalog = catalog: batch.ApplySettings settings, units
        batch.ApplyLoadReference properties.CentroidX, properties.CentroidY, 0#, 0#
        batch.AddCombination "COMP", -2# * properties.Area, 0#, 0#, "PR1", "Audit02 control", "Auto"
        batch.AddCombination "BEND", -0.2 * properties.Area, properties.Area * (maxY - minY) * 1.6, _
            -properties.Area * (maxX - minX) * 0.6, "PR1", "Audit02 control", "Auto"
        batch.Execute
        For i = 1 To batch.Count
            Set item = batch.ResultAt(i)
            result = result & CStr(fixture) & "|" & CStr(i) & "|STATUS|" & item.Status & "|" & item.NormalCrackStatus & _
                "|" & CStr(item.DirectStateMeta.InternalStatus) & "|" & CStr(item.CapacityMeta.InternalStatus) & _
                "|" & CStr(item.StabilityMeta.InternalStatus) & "|" & CStr(item.CrackFormationMeta.InternalStatus) & _
                "|" & CStr(item.CrackWidthMeta.InternalStatus) & "|" & CStr(item.LongitudinalCrackMeta.InternalStatus) & vbCrLf
            result = result & CStr(fixture) & "|" & CStr(i) & "|VALUES|" & TestBits(item.StrengthResult.Capacity.LambdaCapacity) & _
                "|" & TestBits(item.CrackResult.Formation.LambdaCrc) & "|" & TestBits(item.CrackResult.Width.CrackWidth) & _
                "|" & TestBits(item.CrackResult.Width.Abt) & "|" & TestBits(item.StabilityResult.CriticalForce) & _
                "|" & TestBits(item.StabilityResult.Eta1) & "|" & TestBits(item.StabilityResult.Eta2) & _
                "|" & TestBits(item.StabilityResult.DesignN) & "|" & TestBits(item.StabilityResult.DesignMx) & "|" & TestBits(item.StabilityResult.DesignMy) & vbCrLf
            Set meta = item.OverallMeta
            result = result & CStr(fixture) & "|" & CStr(i) & "|COMMENT|" & Replace$(meta.ResultComment, vbCrLf, " / ") & vbCrLf
            result = result & CStr(fixture) & "|" & CStr(i) & "|STABILITY|" & Replace$(item.StabilityMeta.ResultComment, vbCrLf, " / ") & vbCrLf
        Next i
    Next fixture
    GoTo Finished
Failed:
    result = result & "ERROR: " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Finished:
    If Not config Is Nothing And Not IsEmpty(saved) Then TestFixtureRestoreInputs config, saved
    TestDumpAudit02Calculations = result
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Изменяет только адресный параметр в проверочной копии, без фиксированной координаты Config.
Private Sub TestSetSetting(ByVal key As String, ByVal value As String)
    Dim table As Object, row As Long
    Set table = ThisWorkbook.Names("rngSystemSettings").RefersToRange
    For row = 2 To table.Rows.Count
        If CStr(table.Cells(row, 1).Value2) = key Then table.Cells(row, 2).Value2 = value: Exit Sub
    Next row
    Err.Raise vbObjectError + 5491, , "Нет параметра " & key
End Sub
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает все 64 бита Double в шестнадцатеричном виде, без округления для отображения.
Private Function TestBits(ByVal value As Double) As String
    Dim number As TDoubleValue, words As TDoubleWords
    number.Value = value: LSet words = number ' Перенос байтов в два Long сохраняет IEEE 754, в отличие от строкового округления числа.
    TestBits = Right$("00000000" & Hex$(words.High), 8) & Right$("00000000" & Hex$(words.Low), 8)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' сохраняет каждое поле правил всех семейств до и после упрощения рецептов.
' Повторно выполняет рецепт с тем же снимком ввода: ID, порядок, адреса, политики
' и контактные смещения должны совпадать, включая отключённые дополнительные ряды.
Public Function TestDumpRecipeRuleContracts(Optional ByVal includePaths As Boolean = False) As String
    Dim config As Object, saved As Variant, fixture As Variant, scenario As Long, result As String
    Dim settings As CSystemSettingsReader, units As CUnitSystem, reader As CSectionInputSnapshotReader
    Dim registry As CSectionTypeRegistry, context As CSectionRecipeContext, definition As CSectionDefinition
    Dim typeID As String, first As String, second As String
    On Error GoTo Failed
    Set config = ThisWorkbook.Worksheets("Config").UsedRange
    saved = TestFixtureSaveInputs(config)
    For Each fixture In Array("Circle", "Rectangle", "LSection", "TwoRectangles", "RoundedSimple", _
            "RoundedTapered", "RoundedMixed", "RoundedMixedLeft", "HollowCentered", "HollowOffset")
        For scenario = 0 To 5
            TestFixtureRestoreInputs config, saved
            modTestShapeArchitecture.TestConfigureShapeFixture "Circle"
            modTestShapeArchitecture.TestConfigureShapeFixture CStr(fixture)
            Select Case CStr(fixture)
                Case "Circle"
                    TestConfigureCircleRuleContract scenario
                    typeID = "Circle"
                Case "Rectangle", "LSection", "TwoRectangles"
                    modTestShapeArchitecture.TestConfigureRectanglesScenario CStr(fixture), scenario
                    typeID = "Rectangles"
                Case "RoundedSimple", "RoundedTapered", "RoundedMixed", "RoundedMixedLeft"
                    modTestShapeArchitecture.TestConfigureRoundedScenario scenario
                    typeID = "PointedSection"
                Case "HollowCentered", "HollowOffset"
                    modTestShapeArchitecture.TestConfigureHollowScenario scenario
                    typeID = "HollowRectangle"
            End Select
            ThisWorkbook.Worksheets("Config").Calculate
            Set settings = New CSystemSettingsReader
            settings.LoadFromWorkbook ThisWorkbook
            Set units = New CUnitSystem
            units.LoadFromSettings settings
            Set registry = New CSectionTypeRegistry
            Set reader = New CSectionInputSnapshotReader
            Set context = reader.Capture(ThisWorkbook, registry.RecipeTableNames(typeID), units)
            Set definition = registry.CreateDefinition(typeID, context)
            If includePaths Then first = TestRecipePathFields(definition) Else first = TestRecipeRuleFields(definition)
            Set definition = registry.CreateDefinition(typeID, context)
            If includePaths Then second = TestRecipePathFields(definition) Else second = TestRecipeRuleFields(definition)
            If first <> second Then Err.Raise vbObjectError + 5492, , "Повторное описание изменило правила: " & CStr(fixture)
            result = result & "CASE|" & CStr(fixture) & "|" & CStr(scenario) & vbCrLf & first
        Next scenario
    Next fixture
    GoTo Finished
Failed:
    result = result & "ERROR: " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Finished:
    If Not config Is Nothing And Not IsEmpty(saved) Then TestFixtureRestoreInputs config, saved
    TestDumpRecipeRuleContracts = result
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' все используемые линии штатных рецептов, их точная геометрия, имена и исходные параметры.
Public Function TestDumpRecipePathContracts() As String
    TestDumpRecipePathContracts = TestDumpRecipeRuleContracts(True)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' обходим линии правил и аннотаций, затем их исходные линии без изменения определения.
Private Function TestRecipePathFields(ByVal definition As CSectionDefinition) As String
    Dim names As Object, queue As Collection, rule As TRebarRule, annotation As TSectionAnnotation
    Dim i As Long, j As Long, id As Variant, ids As Variant, path As CGeometryPath, edge As TRegionEdge, value As Variant, result As String
    Set names = CreateObject("Scripting.Dictionary"): names.CompareMode = vbBinaryCompare
    Set queue = New Collection
    For i = 1 To definition.RuleCount
        definition.CopyRule i, rule
        If rule.Enabled And rule.Mode <> DependentOffset Then
            ids = definition.TargetPaths(rule.Target)
            For Each id In ids
                If Not names.Exists(CStr(id)) Then names.Add CStr(id), True: queue.Add CStr(id)
            Next id
        End If
    Next i
    For i = 1 To definition.AnnotationCount
        definition.CopyAnnotation i, annotation
        Select Case annotation.Kind
            Case "VerticalExtent", "HorizontalExtent", "RebarLabel", "PathLength"
                ids = definition.TargetPaths(annotation.Target)
            Case "Radius", "Diameter"
                ids = Array(annotation.Target)
            Case Else
                ids = Array()
        End Select
        For Each id In ids
            If Not names.Exists(CStr(id)) Then names.Add CStr(id), True: queue.Add CStr(id)
        Next id
    Next i
    i = 1
    Do While i <= queue.Count
        Set path = definition.Path(CStr(queue(i)))
        result = result & "PATH|" & path.ID & "|" & CStr(path.IsClosed) & "|" & CStr(path.SegmentCount) & "|" & TestBits(path.Length) & vbCrLf
        For j = 1 To path.SegmentCount
            result = result & "SEGMENT|" & CStr(j) & "|" & path.SegmentID(j) & "|" & path.SourcePathID(j) & "|" & path.SourceSegmentID(j)
            path.CopySegment j, edge
            For Each value In Array(path.SourceStartParameter(j), path.SourceEndParameter(j), edge.X1, edge.Y1, edge.X2, edge.Y2, edge.CenterX, edge.CenterY, edge.Sweep)
                result = result & "|" & TestBits(CDbl(value))
            Next value
            result = result & vbCrLf
            id = path.SourcePathID(j)
            If Not names.Exists(CStr(id)) Then names.Add CStr(id), True: queue.Add CStr(id)
        Next j
        i = i + 1
    Loop
    TestRecipePathFields = result
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Готовит варианты круговых рядов: отдельно второй и третий, оба направления,
' нечётные позиции и недопустимый текст в полях отключённого ряда.
Private Sub TestConfigureCircleRuleContract(ByVal scenario As Long)
    Dim table As Object, settings As CSystemSettingsReader, units As CUnitSystem, factor As Double
    Dim row As Long, key As String
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem
    units.LoadFromSettings settings
    factor = units.InputLengthToInternal(1#)
    Set table = ThisWorkbook.Names("rngCircleGeometry").RefersToRange
    Dim fields As Variant, field As Variant, inputCell As Object
    fields = Array("Rebar.Diameter2", "Rebar.Diameter3", "Rebar.Loc2row", "Rebar.Loc3row", "Rebar.Bind2row", "Rebar.Bind3row")
    For Each field In fields
        key = CStr(field): Set inputCell = TestFixtureCircleFieldCell(ThisWorkbook, key)
        Select Case key
            Case "Rebar.Diameter2": inputCell.Value2 = IIf(scenario = 0 Or scenario = 2, 0#, 16# / factor)
            Case "Rebar.Diameter3": inputCell.Value2 = IIf(scenario = 0 Or scenario = 1, 0#, 12# / factor)
            Case "Rebar.Loc2row": inputCell.Value2 = IIf(scenario >= 4, "SideBySide", "Stacked")
            Case "Rebar.Loc3row": inputCell.Value2 = IIf(scenario = 4, "SideBySide", "Stacked")
            Case "Rebar.Bind2row": inputCell.Value2 = IIf(scenario Mod 2 = 0, "EverySecondBar", "EachBar")
            Case "Rebar.Bind3row": inputCell.Value2 = IIf(scenario Mod 2 = 0, "EachBar", "EverySecondBar")
        End Select
        If scenario = 0 And (key = "Rebar.Loc2row" Or key = "Rebar.Loc3row" Or _
                key = "Rebar.Bind2row" Or key = "Rebar.Bind3row") Then inputCell.Value2 = "inactive value"
    Next field
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сериализует все поля TRebarRule без округления Double и без отбора «важных» настроек.
' Сравнение с исходной книгой проверяет также значения по умолчанию независимых options.
Private Function TestRecipeRuleFields(ByVal definition As CSectionDefinition) As String
    Dim rule As TRebarRule, index As Long, field As Variant, value As Variant, result As String
    For index = 1 To definition.RuleCount
        definition.CopyRule index, rule
        For Each field In Array(rule.RuleID, rule.Target, rule.Mode, rule.DistributionScope, rule.Enabled, _
                rule.Diameter, rule.SteelClass, rule.Count, rule.Step, rule.StartSetback, rule.EndSetback, _
                rule.EndMode, rule.RemainderPlacement, rule.SourceRuleID, rule.TangentOffset, rule.NormalOffset, _
                rule.EverySecond, rule.IncomingTangentAtJoint, rule.DirectionX, rule.DirectionY, rule.SourceNormalSign, _
                rule.SuccessfulProjectionOrdinal, rule.NewPublicationRoot, rule.SkipMissingProjection, _
                rule.Publication, rule.PositionOrder, rule.Coincidence, rule.PlacementCheck, rule.OverlapPolicy, _
                rule.SourcePrefix, rule.SourceSuffix, rule.AnnotationGroup, rule.Comment, rule.SourceReference)
            If VarType(field) = vbDouble Then value = TestBits(CDbl(field)) Else value = CStr(field)
            result = result & "|" & Replace$(Replace$(CStr(value), vbCr, "%0D"), vbLf, "%0A")
        Next field
        result = result & "|stations.type=" & CStr(VarType(rule.ExplicitStations))
        If IsArray(rule.ExplicitStations) Then
            For Each field In rule.ExplicitStations
                result = result & "|" & TestBits(CDbl(field))
            Next field
        End If
        result = result & vbCrLf
    Next index
    TestRecipeRuleFields = result
End Function
