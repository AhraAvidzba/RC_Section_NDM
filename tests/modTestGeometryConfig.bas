Attribute VB_Name = "modTestGeometryConfig"
Option Explicit

' ==========================================================================
' ДЛЯ ТЕСТОВ: параметры геометрии через фактический Config
' ==========================================================================
' Проверяет передачу пользовательских геометрических полей от ячеек через
' общие настройки/снимок/рецепт к готовой форме
' и раскладке. Не ищет равновесие и не подменяет инженерный расчет fixtures.
' Измененные таблицы восстанавливаются по исходным формулам даже при ошибке.

Private Type TGeometryConfigStats
    Passed As Long
    Failed As Long
    Report As String
End Type

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет представимость количества до размещения и выделения массивов.
' Только два max-Long случая имеют корректный отступ и воспроизводят Overflow;
' остальные намеренно используют неверный отступ, чтобы сравниваемые реализации завершались
' безопасно, без миллиардного цикла или многогигабайтного массива.
Public Function RunAudit03RebarCounterTests(Optional ByRef passed As Long = 0, _
        Optional ByRef failed As Long = 0) As String
    Dim allConfig As Object, savedConfig As Variant
    Set allConfig = ThisWorkbook.Worksheets("Config").UsedRange: savedConfig = TestFixtureSaveInputs(allConfig)
    Dim stats As TGeometryConfigStats, shapeName As Variant, scenario As Long
    Dim bars As CRebarLayout, number As Long, description As String, expected As Long
    For Each shapeName In Array("Rectangles", "PointedSection", "HollowRectangle")
        expected = vbObjectError + 5530
        For scenario = 0 To 2
            Set bars = Nothing: number = 0: description = vbNullString
            On Error Resume Next
            Set bars = TestCounterFixture(CStr(shapeName), scenario)
            number = Err.Number: description = Err.Description
            Err.Clear
            On Error GoTo 0
            TestCheck stats, "rebarCounter." & CStr(shapeName) & "." & CStr(scenario) & ".typed", number = expected
            TestCheck stats, "rebarCounter." & CStr(shapeName) & "." & CStr(scenario) & ".message", _
                InStr(1, description, ThisWorkbook.Names("rng" & CStr(shapeName) & "Geometry").RefersToRange.Worksheet.Name, vbTextCompare) > 0 And InStr(1, Replace(description, "ё", "е"), "счетчик", vbTextCompare) > 0
            stats.Report = stats.Report & "REBAR_COUNTER_FAILURE: shape=" & CStr(shapeName) & "; scenario=" & CStr(scenario) & _
                "; error=" & CStr(number) & "; comment=" & description & vbCrLf
        Next scenario
        On Error GoTo FailedRun
        Set bars = TestCounterFixture(CStr(shapeName), 3)
        TestCheck stats, "rebarCounter." & CStr(shapeName) & ".disabled", bars.Count = 0
        Set bars = TestCounterFixture(CStr(shapeName), 4)
        TestCheck stats, "rebarCounter." & CStr(shapeName) & ".recovery", bars.Count = 3
    Next shapeName
    Set bars = TestCounterFixture("Rectangles", 5)
    TestCheck stats, "rebarCounter.Rectangles.inactiveLower", bars.Count = 0
    TestCheckRuleCounterBinding stats, 5, 16#, True, 11, False
    TestCheckRuleCounterBinding stats, 6, 16#, False, 15, False
    TestCheckRuleCounterBinding stats, 2147483647, 16#, True, 0, True
    TestCheckRuleCounterBinding stats, 2147483647, 0#, True, 0, False
    GoTo Finish
FailedRun:
    stats.Failed = stats.Failed + 1
    stats.Report = stats.Report & "FAIL: rebarCounter.runtime; " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Finish:
    passed = stats.Passed: failed = stats.Failed
    TestFixtureRestoreInputs allConfig, savedConfig
    RunAudit03RebarCounterTests = stats.Report & "TOTAL_REBAR_COUNTER: passed=" & CStr(passed) & "; failed=" & CStr(failed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет нечётное количество, отключённые ряды и переполнение обычным
' исполнителем правил. Недопустимо большое количество отклоняется до выделения позиций.
Private Sub TestCheckRuleCounterBinding(ByRef stats As TGeometryConfigStats, ByVal count As Long, _
        ByVal diameter As Double, ByVal thirdEverySecond As Boolean, ByVal expected As Long, ByVal overflow As Boolean)
    Dim facade As CSectionDefinitionBuilder, path As CGeometryPath, options As CRebarRuleOptions
    Dim definition As CSectionDefinition, geometry As CGeometryRegion, executor As CRebarRuleExecutor
    Dim bars As CRebarLayout, number As Long, reason As String, label As String
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "CounterBinding"
    Set path = facade.CirclePath("Outer", 0#, 0#, 10000#): facade.Outer "OUTER", path.ID
    Set path = facade.LinePath("Axis", -100#, 0#, 100#, 0#)
    Set options = New CRebarRuleOptions: options.SourceReference = "CounterBinding.Main"
    facade.RebarByCount "Main", path.ID, count, diameter, options
    Set options = New CRebarRuleOptions: options.EverySecond = True: options.NormalOffset = 30#
    facade.RebarDependent "Second", "Main", 12#, options
    Set options = New CRebarRuleOptions: options.EverySecond = thirdEverySecond: options.NormalOffset = 60#
    facade.RebarDependent "Third", "Main", 10#, options
    Set definition = facade.Finish: Set geometry = New CGeometryRegion: geometry.Initialize definition.ConcreteDomain
    Set executor = New CRebarRuleExecutor
    On Error Resume Next
    Set bars = executor.Build(definition, geometry): number = Err.Number: reason = Err.Description: Err.Clear
    On Error GoTo 0
    label = "rebarCounter.binding." & CStr(count) & ".d" & CStr(diameter)
    If overflow Then
        TestCheck stats, label & ".typed", number = vbObjectError + 5530
        TestCheck stats, label & ".beforeAllocation", executor.PositionCount = 0 And InStr(Replace(reason, "ё", "е"), "счетчик") > 0
    Else
        TestCheck stats, label & ".success", number = 0
        If Not bars Is Nothing Then TestCheck stats, label & ".physical", bars.Count = expected
    End If
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' общая защита счетчиков через реальный рецепт и named-ввод.
' Сценарии не размещают миллиарды позиций: preflight обязан отказать до массива.
Private Function TestCounterFixture(ByVal shapeName As String, ByVal scenario As Long) As CRebarLayout
    Dim config As Object, saved As Variant, events As Boolean, fixture As String, root As Object, names As Variant
    Dim settings As CSystemSettingsReader, units As CUnitSystem, registry As CSectionTypeRegistry
    Dim reader As CSectionInputSnapshotReader, context As CSectionRecipeContext, table As CSectionInputTable
    Dim keyed As CSectionInputReader, key As String
    Dim main As Long, extra As Long, faceCount As Long, i As Long, count1 As Long, count2 As Long, factor As Double
    Dim definition As CSectionDefinition, geometry As CGeometryRegion, executor As CRebarRuleExecutor
    Dim number As Long, source As String, reason As String
    On Error GoTo Failed
    events = Application.EnableEvents: Application.EnableEvents = False
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = TestFixtureSaveInputs(config)
    Select Case shapeName
        Case "Rectangles": fixture = "Rectangle"
        Case "PointedSection": fixture = "RoundedSimple"
        Case "HollowRectangle": fixture = "HollowCentered"
    End Select
    modTestShapeArchitecture.TestConfigureShapeFixture fixture
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings: factor = units.InputLengthToInternal(1#)
    Set registry = New CSectionTypeRegistry: names = registry.RecipeTableNames(shapeName)
    Set root = ThisWorkbook.Names(CStr(names(0))).RefersToRange
    Set reader = New CSectionInputSnapshotReader: Set context = reader.Capture(ThisWorkbook, names, units)
    Set root = ThisWorkbook.Names(CStr(names(0))).RefersToRange
    Set table = context.Table(CStr(names(0)))
    Set keyed = context.Reader(CStr(names(0)))
    key = "Top": If shapeName = "Rectangles" Then key = "R1.Top"
    main = keyed.RowIndex("Main", key) - 1: extra = keyed.RowIndex("Extra", key) - 1
    faceCount = 8: If shapeName = "PointedSection" Then faceCount = 4
    For i = 1 To faceCount
        root.Cells(main + i, 3).Value2 = 40# / factor
        root.Cells(main + i, 4).Value2 = 0#
        If shapeName <> "HollowRectangle" Or i <= 4 Then root.Cells(main + i, 5).Value2 = 0
        root.Cells(extra + i, 3).Value2 = 0#: root.Cells(extra + i, 7).Value2 = 0#
        root.Cells(extra + i, 4).Value2 = 0#: root.Cells(extra + i, 8).Value2 = 0#
    Next i
    count1 = 2147483647
    If scenario = 1 Or scenario = 2 Then count1 = 1100000000
    If scenario = 1 Then count2 = 1100000000
    If scenario = 4 Then count1 = 3
    If scenario = 5 Then main = main + 4 ' Нижние грани Rectangle не используются.
    root.Cells(main + 1, 5).Value2 = count1
    If scenario <> 3 Then root.Cells(main + 1, 4).Value2 = 16# / factor
    If count2 > 0 Then root.Cells(main + 2, 5).Value2 = count2: root.Cells(main + 2, 4).Value2 = 16# / factor
    If scenario = 2 Then
        root.Cells(extra + 1, 3).Value2 = 16# / factor
        root.Cells(extra + 1, 6).Value2 = "Stacked": root.Cells(extra + 1, 5).Value2 = "EachBar"
    End If
    Set definition = ReadWorkbookSectionDefinition(ThisWorkbook, settings, units)
    Set geometry = New CGeometryRegion: geometry.Initialize definition.ConcreteDomain
    Set executor = New CRebarRuleExecutor: Set TestCounterFixture = executor.Build(definition, geometry)
    GoTo Finished
Failed:
    number = Err.Number: source = Err.Source: reason = Err.Description
Finished:
    On Error Resume Next
    If Not config Is Nothing And Not IsEmpty(saved) Then TestFixtureRestoreInputs config, saved
    Application.EnableEvents = events
    On Error GoTo 0
    If number <> 0 Then Err.Raise number, source, reason
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Выполняет активные/неактивные варианты восьми параметров круга, границы
' и восстановление после ошибочного ввода. Счетчики доступны полному UI-suite.
Public Function RunAudit03CircleConfigTests(Optional ByRef passed As Long = 0, _
        Optional ByRef failed As Long = 0) As String
    Dim allConfig As Object, savedConfig As Variant
    Set allConfig = ThisWorkbook.Worksheets("Config").UsedRange: savedConfig = TestFixtureSaveInputs(allConfig)
    Dim stats As TGeometryConfigStats
    Dim systemRange As Object, circleRange As Object, unitRange As Object
    Dim savedSystem As Variant, savedCircle As Variant, savedUnits As Variant
    On Error GoTo FailedRun
    Set systemRange = ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange
    Set circleRange = ThisWorkbook.Names.Item("rngCircleGeometry").RefersToRange
    Set unitRange = ThisWorkbook.Names.Item("rngUnitSettings").RefersToRange
    savedSystem = systemRange.Formula: savedCircle = circleRange.Formula: savedUnits = unitRange.Formula
    TestSetKey systemRange, "Geometry.Type", "Circle"
    unitRange.Cells(2, 2).Value2 = "mm"
    TestConfigureCircle circleRange
    TestCheckCircleEffects stats, circleRange
    TestCheckCircleBoundaries stats, circleRange
    TestCheckCircleInactive stats, systemRange, circleRange
    GoTo Restore
FailedRun:
    stats.Failed = stats.Failed + 1
    stats.Report = stats.Report & "FAIL: circleConfig.runtime; " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Restore:
    On Error GoTo FailedRestore
    If Not systemRange Is Nothing Then systemRange.Formula = savedSystem
    If Not circleRange Is Nothing Then circleRange.Formula = savedCircle
    If Not unitRange Is Nothing Then unitRange.Formula = savedUnits
    If Not systemRange Is Nothing Then TestCheck stats, "circleConfig.restore.system", TestSameFormula(systemRange.Formula, savedSystem)
    If Not circleRange Is Nothing Then TestCheck stats, "circleConfig.restore.circle", TestSameFormula(circleRange.Formula, savedCircle)
    If Not unitRange Is Nothing Then TestCheck stats, "circleConfig.restore.units", TestSameFormula(unitRange.Formula, savedUnits)
    GoTo Finish
FailedRestore:
    stats.Failed = stats.Failed + 1
    stats.Report = stats.Report & "FAIL: circleConfig.restore.runtime; " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Finish:
    passed = stats.Passed: failed = stats.Failed
    TestFixtureRestoreInputs allConfig, savedConfig
    RunAudit03CircleConfigTests = stats.Report & "TOTAL_CIRCLE_CONFIG: passed=" & CStr(passed) & "; failed=" & CStr(failed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет размеры и режимы трех некруглых форм через реальные ячейки,
' включая переключение единиц, неактивные поля и неверные количества арматуры.
' Все измененные таблицы восстанавливаются, чтобы тест не менял рабочий ввод.
Public Function RunAudit03ShapeConfigTests(Optional ByRef passed As Long = 0, _
        Optional ByRef failed As Long = 0) As String
    Dim allConfig As Object, savedConfig As Variant
    Set allConfig = ThisWorkbook.Worksheets("Config").UsedRange: savedConfig = TestFixtureSaveInputs(allConfig)
    Dim stats As TGeometryConfigStats
    Dim systemRange As Object, unitRange As Object, circleRange As Object
    Dim target As Object, rangeName As Variant, shapeName As String, shapeNames As Variant
    Dim savedSystem As Variant, savedUnits As Variant, savedCircle As Variant
    Dim savedShapes(0 To 2) As Variant, shapeRanges(0 To 2) As Object, index As Long
    On Error GoTo FailedRun
    Set systemRange = ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange
    Set unitRange = ThisWorkbook.Names.Item("rngUnitSettings").RefersToRange
    Set circleRange = ThisWorkbook.Names.Item("rngCircleGeometry").RefersToRange
    savedSystem = systemRange.Formula: savedUnits = unitRange.Formula: savedCircle = circleRange.Formula
    For Each rangeName In Array("rngRectanglesGeometry", "rngPointedSectionGeometry", "rngHollowRectangleGeometry")
        Set shapeRanges(index) = ThisWorkbook.Names.Item(CStr(rangeName)).RefersToRange
        savedShapes(index) = shapeRanges(index).Formula
        index = index + 1
    Next rangeName
    unitRange.Cells(2, 2).Value2 = "mm"
    TestConfigureCircle circleRange
    shapeNames = Array("Rectangles", "PointedSection", "HollowRectangle")
    For index = 0 To 2
        Set target = shapeRanges(index)
        shapeName = CStr(shapeNames(index))
        TestSetKey systemRange, "Geometry.Type", shapeName
        TestCheckShapeEffects stats, target, shapeName
        TestCheckShapeInactiveModes stats, target, shapeName
        TestCheckShapeUnits stats, target, unitRange, shapeName
        TestSetKey systemRange, "Geometry.Type", "Circle"
        TestCheckShapeInactive stats, target, shapeName
    Next index
    TestSetKey systemRange, "Geometry.Type", "HollowRectangle"
    TestCheckHollowCounts stats, shapeRanges(2)
    GoTo Restore
FailedRun:
    stats.Failed = stats.Failed + 1
    stats.Report = stats.Report & "FAIL: shapeConfig.runtime; " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Restore:
    On Error GoTo FailedRestore
    For index = 0 To 2
        If Not shapeRanges(index) Is Nothing Then
            shapeRanges(index).Formula = savedShapes(index)
            TestCheck stats, "shapeConfig.restore." & CStr(index), TestSameFormula(shapeRanges(index).Formula, savedShapes(index))
        End If
    Next index
    If Not systemRange Is Nothing Then systemRange.Formula = savedSystem
    If Not unitRange Is Nothing Then unitRange.Formula = savedUnits
    If Not circleRange Is Nothing Then circleRange.Formula = savedCircle
    If Not systemRange Is Nothing Then TestCheck stats, "shapeConfig.restore.system", TestSameFormula(systemRange.Formula, savedSystem)
    If Not unitRange Is Nothing Then TestCheck stats, "shapeConfig.restore.units", TestSameFormula(unitRange.Formula, savedUnits)
    If Not circleRange Is Nothing Then TestCheck stats, "shapeConfig.restore.circle", TestSameFormula(circleRange.Formula, savedCircle)
    GoTo Finish
FailedRestore:
    stats.Failed = stats.Failed + 1
    stats.Report = stats.Report & "FAIL: shapeConfig.restore.runtime; " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Finish:
    passed = stats.Passed: failed = stats.Failed
    TestFixtureRestoreInputs allConfig, savedConfig
    RunAudit03ShapeConfigTests = stats.Report & "TOTAL_SHAPE_CONFIG: passed=" & CStr(passed) & "; failed=" & CStr(failed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Описывает фактические поля трёх таблиц листа «Сечения», а не поиск
' строк с искусственными ключами. Поля: ключ, строка, колонка, база, изменение,
' недопустимое значение. Эти позиции проверены по действующему reader-у.
Private Function TestShapeFields(ByVal shapeName As String) As Variant
    Select Case shapeName
        Case "Rectangles"
            TestShapeFields = Array(Array("Rectangles.H1", 5, 3, 300#, 350#, -1#), _
                Array("Rectangles.B1", 6, 3, 200#, 240#, -1#), _
                Array("Rectangles.H2", 7, 3, 250#, 300#, -1#), _
                Array("Rectangles.B2", 8, 3, 600#, 650#, -1#), _
                Array("Rectangles.UpperOffsetX", 9, 3, -50#, -70#, 10000#), _
                Array("Rectangles.SectionType", 4, 3, "TwoRect", "OneRect", "Unknown"))
        Case "PointedSection"
            TestShapeFields = Array(Array("PointedSection.B", 4, 3, 600#, 650#, -1#), _
                Array("PointedSection.H", 5, 3, 400#, 450#, -1#), _
                Array("PointedSection.Left.W", 6, 3, 100#, 130#, -1#), _
                Array("PointedSection.Right.W", 9, 3, 100#, 130#, -1#), _
                Array("PointedSection.Left.R1", 7, 3, 30#, 45#, -1#), _
                Array("PointedSection.Right.R1", 10, 3, 30#, 45#, -1#), _
                Array("PointedSection.Left.R2", 8, 3, 20#, 35#, -1#), _
                Array("PointedSection.Right.R2", 11, 3, 20#, 35#, -1#))
        Case "HollowRectangle"
            TestShapeFields = Array(Array("HollowRectangle.H", 5, 3, 1200#, 1300#, -1#), _
                Array("HollowRectangle.B", 4, 3, 1000#, 1100#, -1#), _
                Array("HollowRectangle.R", 6, 3, 60#, 80#, -1#), _
                Array("HollowRectangle.OpeningH", 8, 3, 500#, 550#, -1#), _
                Array("HollowRectangle.OpeningB", 7, 3, 400#, 450#, -1#), _
                Array("HollowRectangle.OpeningR", 9, 3, 30#, 40#, -1#), _
                Array("HollowRectangle.InnerOffsetX", 10, 3, 40#, 60#, 10000#), _
                Array("HollowRectangle.InnerOffsetY", 11, 3, 50#, 70#, 10000#))
        Case Else
            Err.Raise vbObjectError + 5961, "modTestGeometryConfig", "Неизвестная форма fixture: " & shapeName
    End Select
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Задает физически допустимую форму. Множитель переводит только длины,
' поэтому mm и cm fixtures описывают одно сечение, а enum не изменяются.
Private Sub TestConfigureShape(ByVal target As Object, ByVal shapeName As String, Optional ByVal lengthScale As Double = 1#)
    Dim item As Variant, value As Variant
    For Each item In TestShapeFields(shapeName)
        value = item(3)
        If IsNumeric(value) Then value = CDbl(value) * lengthScale
        target.Cells(CLng(item(1)), CLng(item(2))).Value2 = value
    Next item
    TestFixtureResetReinforcement shapeName
    TestConfigureRequiredGeometryRebars target, shapeName, lengthScale
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' ==========================================================================
' ОБЯЗАТЕЛЬНЫЙ И УСЛОВНО НЕАКТИВНЫЙ ВВОД АРМАТУРЫ КРУГА
' ==========================================================================
' Проверяет реальные ячейки круга, отсутствие обязательных строк и ошибки
' формул. Отдельно проверяет отключенные ряды и невыбранную форму: их значения
' не участвуют в раскладке. Весь исходный ввод восстанавливается по Formula.
Public Function RunAudit03CircleRebarInputTests(Optional ByRef passed As Long = 0, _
        Optional ByRef failed As Long = 0) As String
    Dim allConfig As Object, savedConfig As Variant
    Set allConfig = ThisWorkbook.Worksheets("Config").UsedRange: savedConfig = TestFixtureSaveInputs(allConfig)
    Dim stats As TGeometryConfigStats, ranges(0 To 3) As Object, saved(0 To 3) As Variant
    Dim names As Variant, keys As Variant, key As Variant, bad As Variant, badValues As Variant
    Dim i As Long, cases As Long, cell As Object, savedKey As Variant, label As Object
    Dim baseline As Variant, prefix As String, offKey As Variant, offValue As Variant, ignored As Variant
    names = Array("rngSystemSettings", "rngUnitSettings", "rngCircleGeometry", "rngRectanglesGeometry")
    keys = Array("Circle.Diameter", "Rebar.AxisDistance", "Rebar.Count", "Rebar.Diameter", _
        "Rebar.Diameter2", "Rebar.Diameter3", "Rebar.Loc2row", "Rebar.Loc3row")
    On Error GoTo FailedRun
    For i = 0 To 3
        Set ranges(i) = ThisWorkbook.Names.Item(CStr(names(i))).RefersToRange
        saved(i) = ranges(i).Formula
    Next i
    TestSetKey ranges(0), "Geometry.Source", "Generated"
    TestSetKey ranges(0), "Geometry.Type", "Circle"
    ranges(1).Cells(2, 2).Value2 = "mm"
    TestConfigureCircle ranges(2)
    baseline = TestRequiredGeometryRouteSnapshot("Circle", "Rebars")
    For Each key In keys
        TestConfigureCircle ranges(2)
        If InStr(1, CStr(key), "Loc", vbBinaryCompare) > 0 Then
            badValues = Array(vbNullString, "Wrong", CVErr(2015))
        ElseIf CStr(key) = "Rebar.Count" Then
            badValues = Array(-1#, True, 0.5, 2147483648#, "abc", CVErr(2015)) ' Один стержень теперь допустим; логическое значение по-прежнему не является количеством.
        ElseIf CStr(key) = "Circle.Diameter" Then
            badValues = Array(vbNullString, -1#, 0#, "abc", CVErr(2015))
        ElseIf CStr(key) = "Rebar.AxisDistance" Then
            badValues = Array(vbNullString, -1#, 0#, 500#, "abc", CVErr(2015))
        Else
            badValues = Array(-1#, "abc", CVErr(2015))
        End If
        Set cell = TestCircleInputCell(ranges(2), CStr(key))
        i = 0
        For Each bad In badValues
            i = i + 1
            TestCheckCircleInputAttempt stats, cell, bad, "circleRebar.active." & CStr(key) & "." & CStr(i), _
                "Circle", baseline, True, cases
        Next bad

        ' Пустой диаметр является допустимым выключением; исчезнувшая строка
        ' имеет другой смысл и не должна незаметно брать шаблонный диаметр.
        If CStr(key) = "Circle.Diameter" Then
            Set label = cell.Offset(0, -2)
        Else
            Set label = cell.Offset(-2, 0)
        End If
        savedKey = label.Formula
        label.Value2 = "Audit03.AbsentCircleField"
        TestCheckCircleMissingInput stats, CStr(key), cases, CStr(savedKey)
        label.Formula = savedKey
        TestCheck stats, "circleRebar.missing." & CStr(key) & ".recovery", _
            TestSameGeometryRouteSnapshot(baseline, TestRequiredGeometryRouteSnapshot("Circle", "Rebars"))
    Next key

    For Each offKey In Array("Rebar.Count", "Rebar.Diameter", "Rebar.Diameter2", "Rebar.Diameter3")
        For Each offValue In Array(vbNullString, 0#)
            TestConfigureCircle ranges(2): TestSetKey ranges(2), CStr(offKey), offValue
            baseline = TestRequiredGeometryRouteSnapshot("Circle", "Rebars")
            If CStr(offKey) = "Rebar.Count" Or CStr(offKey) = "Rebar.Diameter" Then
                ignored = Array("Rebar.AxisDistance", "Rebar.Diameter2", "Rebar.Diameter3", "Rebar.Loc2row", "Rebar.Loc3row")
            ElseIf CStr(offKey) = "Rebar.Diameter2" Then
                ignored = Array("Rebar.Loc2row")
            Else
                ignored = Array("Rebar.Loc3row")
            End If
            For Each key In ignored
                Set cell = TestCircleInputCell(ranges(2), CStr(key)): i = 0
                For Each bad In Array(vbNullString, "abc", CVErr(2015), -1#)
                    i = i + 1
                    prefix = "circleRebar.inactive." & CStr(offKey) & "." & CStr(offValue) & "." & CStr(key) & "." & CStr(i)
                    TestCheckCircleInputAttempt stats, cell, bad, prefix, "Circle", baseline, False, cases
                Next bad
            Next key
        Next offValue
    Next offKey

    TestConfigureRebarFieldFixture ranges(3), "Rectangles", True
    TestSetKey ranges(0), "Geometry.Type", "Rectangles"
    baseline = TestRequiredGeometryRouteSnapshot("Rectangles", "Rebars")
    ' Выбран рецепт Rectangles; повреждённые поля отдельной таблицы Circle не должны влиять на его сборку.
    For Each key In Array("Circle.Diameter")
        Set cell = TestCircleInputCell(ranges(2), CStr(key)): i = 0
        For Each bad In Array(vbNullString, "abc", CVErr(2015))
            i = i + 1
            TestCheckCircleInputAttempt stats, cell, bad, "circleRebar.otherShape." & CStr(key) & "." & CStr(i), _
                "Rectangles", baseline, False, cases
        Next bad
    Next key
    GoTo Restore
FailedRun:
    stats.Failed = stats.Failed + 1
    stats.Report = stats.Report & "FAIL: circleRebar.runtime; " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Restore:
    On Error GoTo FailedRestore
    For i = 0 To 3
        If Not ranges(i) Is Nothing Then
            ranges(i).Formula = saved(i)
            TestCheck stats, "circleRebar.restore." & CStr(names(i)), TestSameFormula(ranges(i).Formula, saved(i))
        End If
    Next i
    GoTo Finish
FailedRestore:
    stats.Failed = stats.Failed + 1
    stats.Report = stats.Report & "FAIL: circleRebar.restore.runtime; " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Finish:
    passed = stats.Passed: failed = stats.Failed
    TestFixtureRestoreInputs allConfig, savedConfig
    RunAudit03CircleRebarInputTests = stats.Report & "TOTAL_CIRCLE_REBAR_INPUT: passed=" & CStr(passed) & _
        "; failed=" & CStr(failed) & "; cases=" & CStr(cases) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Находит фактическую ячейку значения по ключу, без типовых адресов Config.
' Это позволяет проверить адресную диагностику после перемещения таблицы.
Private Function TestCircleInputCell(ByVal target As Object, ByVal key As String) As Object
    Set TestCircleInputCell = TestFixtureCircleFieldCell(target.Worksheet.Parent, key, target)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет один ввод через полный production-маршрут. Ошибка активного поля
' должна иметь адрес и инструкцию исправления; неактивное поле обязано оставить
' весь снимок неизменным. Затем проверяется восстановление исходной формулы.
Private Sub TestCheckCircleInputAttempt(ByRef stats As TGeometryConfigStats, ByVal cell As Object, _
        ByVal value As Variant, ByVal prefix As String, ByVal shape As String, ByVal baseline As Variant, _
        ByVal mustReject As Boolean, ByRef cases As Long)
    Dim savedValue As Variant, number As Long, description As String, actual As Variant
    savedValue = cell.Formula: cell.Value2 = value: cases = cases + 1
    TestRequiredGeometryProgress prefix, stats.Report
    On Error Resume Next
    actual = TestRequiredGeometryRouteSnapshot(shape, "Rebars")
    number = Err.Number: description = Err.Description
    Err.Clear: On Error GoTo 0
    If mustReject Then
        TestCheck stats, prefix & ".reject", number <> 0
        TestCheck stats, prefix & ".address", InStr(1, description, cell.Address(False, False), vbTextCompare) > 0
        TestCheck stats, prefix & ".repair", InStr(1, description, "Введите", vbTextCompare) > 0 Or _
            InStr(1, description, "Выберите", vbTextCompare) > 0 Or InStr(1, description, "Исправ", vbTextCompare) > 0
    Else
        TestCheck stats, prefix & ".noInterference", number = 0
        If number = 0 Then TestCheck stats, prefix & ".snapshot", TestSameGeometryRouteSnapshot(baseline, actual)
    End If
    stats.Report = stats.Report & "CIRCLE_REBAR_INPUT: " & prefix & "; error=" & CStr(number) & "; " & description & vbCrLf
    cell.Formula = savedValue
    TestCheck stats, prefix & ".recovery", TestSameGeometryRouteSnapshot(baseline, TestRequiredGeometryRouteSnapshot(shape, "Rebars"))
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет структурно отсутствующий активный параметр отдельно от пустой
' ячейки. Сообщение должно назвать ключ и способ восстановить ввод, а не default.
Private Sub TestCheckCircleMissingInput(ByRef stats As TGeometryConfigStats, ByVal key As String, ByRef cases As Long, ByVal expectedCaption As String)
    Dim actual As Variant, number As Long, description As String
    cases = cases + 1: TestRequiredGeometryProgress "circleRebar.missing." & key, stats.Report
    On Error Resume Next
    actual = TestRequiredGeometryRouteSnapshot("Circle", "Rebars")
    number = Err.Number: description = Err.Description
    Err.Clear: On Error GoTo 0
    TestCheck stats, "circleRebar.missing." & key & ".reject", number <> 0
    TestCheck stats, "circleRebar.missing." & key & ".key", InStr(1, description, expectedCaption, vbBinaryCompare) > 0
    TestCheck stats, "circleRebar.missing." & key & ".repair", InStr(1, description, "Восстанов", vbTextCompare) > 0 Or _
        InStr(1, description, "Введите", vbTextCompare) > 0 Or InStr(1, description, "Выберите", vbTextCompare) > 0
    stats.Report = stats.Report & "CIRCLE_MISSING_INPUT: " & key & "; error=" & CStr(number) & "; " & description & vbCrLf
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет бетонную область общего описания сечения без размещения стержней.
Private Function TestReadShape() As CGeometryRegion
    Dim settings As CSystemSettingsReader, units As CUnitSystem, definition As CSectionDefinition
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    Set definition = ReadWorkbookSectionDefinition(ThisWorkbook, settings, units)
    Set TestReadShape = New CGeometryRegion: TestReadShape.Initialize definition.ConcreteDomain
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сохраняет вычисленные площадь, центр и габариты. Непригодная геометрия — ошибка
' подготовки теста; исходные размеры формы не подставляются вместо её характеристик.
Private Function TestShapeSnapshot(ByVal geom As CGeometryRegion) As Variant
    Dim message As String, available As Boolean, area As Double, X As Double, Y As Double
    If Not geom.IsValid(message) Then Err.Raise vbObjectError + 5962, "modTestGeometryConfig", message
    area = geom.AnalyticalArea(available)
    If Not available Then Err.Raise vbObjectError + 5963, "modTestGeometryConfig", "Нет аналитической площади fixture."
    geom.AnalyticalCentroid available, X, Y
    If Not available Then Err.Raise vbObjectError + 5964, "modTestGeometryConfig", "Нет центра тяжести fixture."
    TestShapeSnapshot = Array(area, X, Y, geom.MinX, geom.MaxX, geom.MinY, geom.MaxY)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает все численные характеристики в фиксированном абсолютном допуске.
' Локальный допуск теста не меняет точность mesh/solver пользовательского расчета.
Private Function TestSameShape(ByVal first As Variant, ByVal second As Variant) As Boolean
    Dim index As Long
    For index = 0 To 6
        If Abs(CDbl(first(index)) - CDbl(second(index))) > 0.000001 Then Exit Function
    Next index
    TestSameShape = True
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Для каждого поля проверяет изменение готовой формы, отклонение неверного
' значения и повторный нормальный ввод. Числа журналируются для последующей сверки.
Private Sub TestCheckShapeEffects(ByRef stats As TGeometryConfigStats, ByVal target As Object, ByVal shapeName As String)
    Dim item As Variant, baseline As Variant, changed As Variant, geom As CGeometryRegion
    Dim errorNumber As Long, description As String, valid As Boolean, prefix As String
    TestConfigureShape target, shapeName
    baseline = TestShapeSnapshot(TestReadShape())
    TestCheckShapeBaseline stats, shapeName, baseline
    For Each item In TestShapeFields(shapeName)
        prefix = "shapeConfig." & CStr(item(0))
        TestConfigureShape target, shapeName
        target.Cells(CLng(item(1)), CLng(item(2))).Value2 = item(4)
        changed = TestShapeSnapshot(TestReadShape())
        TestCheck stats, prefix & ".effect", Not TestSameShape(baseline, changed)
        stats.Report = stats.Report & "SHAPE_CONFIG_EFFECT: key=" & CStr(item(0)) & "; area=" & CStr(changed(0)) & _
            "; cx=" & CStr(changed(1)) & "; cy=" & CStr(changed(2)) & vbCrLf
        target.Cells(CLng(item(1)), CLng(item(2))).Value2 = item(5)
        errorNumber = 0: description = vbNullString: valid = False
        On Error Resume Next
        Set geom = TestReadShape()
        errorNumber = Err.Number: description = Err.Description
        If errorNumber = 0 Then valid = geom.IsValid(description)
        Err.Clear
        On Error GoTo 0
        TestCheck stats, prefix & ".invalid", Not valid And Len(description) > 0
        TestConfigureShape target, shapeName
        TestCheck stats, prefix & ".recovery", TestSameShape(baseline, TestShapeSnapshot(TestReadShape()))
    Next item
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Независимо проверяет площадь и центр. Скругления представлены точными
' дугами: потеря площади четырех углов равна (4 - pi) * R^2.
Private Sub TestCheckShapeBaseline(ByRef stats As TGeometryConfigStats, ByVal shapeName As String, ByVal actual As Variant)
    Dim area As Double, openingArea As Double, cornerLoss As Double
    If shapeName = "Rectangles" Then
        area = 200# * 300# + 600# * 250#
        TestNear stats, "shapeConfig.Rectangles.oracle.area", CDbl(actual(0)), area
        TestNear stats, "shapeConfig.Rectangles.oracle.cx", CDbl(actual(1)), (600# * 250# * 300# + 200# * 300# * 50#) / area
        TestNear stats, "shapeConfig.Rectangles.oracle.cy", CDbl(actual(2)), (600# * 250# * 125# + 200# * 300# * 400#) / area
    ElseIf shapeName = "HollowRectangle" Then
        cornerLoss = 4# - GEOM_PI
        openingArea = 400# * 500# - cornerLoss * 30# ^ 2
        area = 1000# * 1200# - cornerLoss * 60# ^ 2 - openingArea
        TestNear stats, "shapeConfig.HollowRectangle.oracle.area", CDbl(actual(0)), area
        TestNear stats, "shapeConfig.HollowRectangle.oracle.cx", CDbl(actual(1)), -openingArea * 40# / area
        TestNear stats, "shapeConfig.HollowRectangle.oracle.cy", CDbl(actual(2)), -openingArea * 50# / area
    Else
        TestNear stats, "shapeConfig.PointedSection.oracle.cx", CDbl(actual(1)), 0#
        TestNear stats, "shapeConfig.PointedSection.oracle.cy", CDbl(actual(2)), 0#
    End If
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет именно условно выключенные поля формы: нижний прямоугольник и
' смещение, а также W/R2 простых боковых граней не меняют активную геометрию.
Private Sub TestCheckShapeInactiveModes(ByRef stats As TGeometryConfigStats, ByVal target As Object, ByVal shapeName As String)
    Dim baseline As Variant, item As Variant
    TestConfigureShape target, shapeName
    If shapeName = "Rectangles" Then
        target.Cells(4, 3).Value2 = "OneRect"
        baseline = TestShapeSnapshot(TestReadShape())
        For Each item In Array(Array(7, 3), Array(8, 3), Array(9, 3))
            target.Cells(CLng(item(0)), CLng(item(1))).Value2 = 100#
            TestCheck stats, "shapeConfig.Rectangles.Rectangle.inactive." & CStr(item(0)) & "." & CStr(item(1)), _
                TestSameShape(baseline, TestShapeSnapshot(TestReadShape()))
        Next item
        TestConfigureShape target, shapeName
        target.Cells(4, 3).Value2 = "TwoRect"
        baseline = TestShapeSnapshot(TestReadShape())
        target.Cells(9, 3).Value2 = 100#
        TestCheck stats, "shapeConfig.Rectangles.TwoRect.active.offset", Not TestSameShape(baseline, TestShapeSnapshot(TestReadShape()))
    ElseIf shapeName = "PointedSection" Then
        target.Cells(6, 3).Value2 = 0#: target.Cells(9, 3).Value2 = 0#
        baseline = TestShapeSnapshot(TestReadShape())
        For Each item In Array(Array(8, 3), Array(11, 3))
            target.Cells(CLng(item(0)), CLng(item(1))).Value2 = 70#
            TestCheck stats, "shapeConfig.PointedSection.W0.inactive." & CStr(item(0)) & "." & CStr(item(1)), _
                TestSameShape(baseline, TestShapeSnapshot(TestReadShape()))
        Next item
    End If
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Смена INPUT mm на cm с явным переводом всех длин обязана сохранять
' внутреннюю форму. Это не меняет контракт импорта AutoCAD и его миллиметры.
Private Sub TestCheckShapeUnits(ByRef stats As TGeometryConfigStats, ByVal target As Object, _
        ByVal unitRange As Object, ByVal shapeName As String)
    Dim baseline As Variant, baselineBars As Variant
    unitRange.Cells(2, 2).Value2 = "mm"
    TestConfigureShape target, shapeName
    baseline = TestShapeSnapshot(TestReadShape())
    baselineBars = TestRequiredGeometryRouteSnapshot(shapeName, "Rebars")
    unitRange.Cells(2, 2).Value2 = "cm"
    TestConfigureShape target, shapeName, 0.1
    TestCheck stats, "shapeConfig." & shapeName & ".units.cm", TestSameShape(baseline, TestShapeSnapshot(TestReadShape()))
    TestCheck stats, "shapeConfig." & shapeName & ".units.cm.rebars", TestSameGeometryRouteSnapshot(baselineBars, TestRequiredGeometryRouteSnapshot(shapeName, "Rebars"))
    unitRange.Cells(2, 2).Value2 = "mm"
    TestConfigureShape target, shapeName
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Рецепт публикует пути армирования: переводим отступы/диаметры вместе с B/H/R.


' -------------------- ДЛЯ ТЕСТОВ --------------------
' В другом Geometry.Type корректные изменения невыбранной таблицы не должны
' влиять на расчетную форму Circle. Неверные значения здесь не маскируются.
Private Sub TestCheckShapeInactive(ByRef stats As TGeometryConfigStats, ByVal target As Object, ByVal shapeName As String)
    Dim item As Variant, baseline As Variant
    TestConfigureShape target, shapeName
    TestConfigureCircle ThisWorkbook.Names("rngCircleGeometry").RefersToRange
    baseline = TestShapeSnapshot(TestReadShape())
    For Each item In TestShapeFields(shapeName)
        TestConfigureShape target, shapeName
        TestConfigureCircle ThisWorkbook.Names("rngCircleGeometry").RefersToRange
        target.Cells(CLng(item(1)), CLng(item(2))).Value2 = item(4)
        TestCheck stats, "shapeConfig." & CStr(item(0)) & ".inactive.type", TestSameShape(baseline, TestShapeSnapshot(TestReadShape()))
    Next item
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет все четыре пользовательских счетчика наружной арматуры. Текст,
' дробь, отрицательное/непредставимое число и ошибка Excel не выключают грань
' молча; сообщение должно назвать Config, фактическую ячейку и настройку.
' Служебные количества внутренних граней не являются пользовательским вводом.
Private Sub TestCheckHollowCounts(ByRef stats As TGeometryConfigStats, ByVal target As Object)
    Dim rowIndex As Long, index As Long, bad As Variant, geom As CGeometryRegion, bars As CRebarLayout
    Dim errorNumber As Long, description As String, face As String, prefix As String, faceNames As Variant
    faceNames = Array("B.Top", "B.Bottom", "H.Left", "H.Right")
    TestConfigureRebarFieldFixture target, "HollowRectangle", False
    Dim sides As Variant, cell As Object
    sides = Array("Top", "Bottom", "Left", "Right")
    For rowIndex = 0 To 3: TestFixtureInputCell(target, "Main", CStr(sides(rowIndex)), "n").Value2 = 3: Next rowIndex
    For rowIndex = 0 To 3
        Set cell = TestFixtureInputCell(target, "Main", CStr(sides(rowIndex)), "n")
        face = CStr(faceNames(rowIndex))
        cell.Value2 = 4
        TestReadGeometry geom, bars
        TestCheck stats, "shapeConfig.hollowCount." & face & ".effect", TestFaceBarCount(bars, face) = 4
        cell.Value2 = 3
        For Each bad In Array("abc", -1#, 0.5, CVErr(2015), 2147483648#)
            index = index + 1
            cell.Value2 = bad
            errorNumber = 0: description = vbNullString
            On Error Resume Next
            TestReadGeometry geom, bars
            errorNumber = Err.Number: description = Err.Description
            Err.Clear
            On Error GoTo 0
            prefix = "shapeConfig.hollowCount." & face & "." & CStr(index)
            TestCheck stats, prefix & ".reject", errorNumber <> 0
            TestCheck stats, prefix & ".location", InStr(1, description, cell.Worksheet.Name & "!", vbTextCompare) > 0 And _
                InStr(1, description, cell.Address(False, False), vbTextCompare) > 0
            TestCheck stats, prefix & ".parameter", InStr(1, description, CStr(sides(rowIndex)), vbBinaryCompare) > 0 And _
                (InStr(1, description, "Число стержней", vbBinaryCompare) > 0 Or InStr(1, description, "число", vbBinaryCompare) > 0 Or InStr(1, description, "Excel", vbBinaryCompare) > 0)
            stats.Report = stats.Report & "HOLLOW_COUNT_ERROR: " & prefix & "; error=" & CStr(errorNumber) & "; " & description & vbCrLf
            cell.Value2 = 3
            TestReadGeometry geom, bars
            TestCheck stats, prefix & ".recovery", TestFaceBarCount(bars, face) = 3
        Next bad
        cell.Value2 = vbNullString
        TestReadGeometry geom, bars
        TestCheck stats, "shapeConfig.hollowCount." & face & ".blank.off", TestFaceBarCount(bars, face) = 0
        cell.Value2 = 0
        TestReadGeometry geom, bars
        TestCheck stats, "shapeConfig.hollowCount." & face & ".zero.off", TestFaceBarCount(bars, face) = 0
        cell.Value2 = 3
    Next rowIndex
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Считает стержни только выбранной стороны, исключая проекции на отверстие
' и соседние стороны. Наличие арматуры в других местах не доказывает действие настройки
' количества.
Private Function TestFaceBarCount(ByVal bars As CRebarLayout, ByVal face As String) As Long
    Dim index As Long
    For index = 1 To bars.Count
        If StrComp(bars.BarAnnotationGroupName(index), face, vbBinaryCompare) = 0 Then TestFaceBarCount = TestFaceBarCount + 1
    Next index
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Задает достаточное защитное расстояние и три активных ряда. Каждая следующая
' проверка меняет только одно поле относительно этого физически допустимого набора.
Private Sub TestConfigureCircle(ByVal target As Object)
    TestSetKey target, "Circle.Diameter", 1000#
    TestSetKey target, "Rebar.AxisDistance", 50#
    TestSetKey target, "Rebar.Count", 8
    TestSetKey target, "Rebar.Diameter", 16#
    TestSetKey target, "Rebar.Diameter2", 12#
    TestSetKey target, "Rebar.Diameter3", 10#
    TestSetKey target, "Rebar.Loc2row", "Stacked"
    TestSetKey target, "Rebar.Loc3row", "Stacked"
    TestSetKey target, "Rebar.Bind2row", "EachBar"
    TestSetKey target, "Rebar.Bind3row", "EachBar"
    TestSetKey target, "Rebar.AxisOffset2", 0#: TestSetKey target, "Rebar.AxisOffset3", 0#
    TestSetKey target, "Rebar.StartSetback", 0#: TestSetKey target, "Rebar.EndSetback", 0#
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Первая позиция находится над центром, обход идёт по часовой стрелке.
' Контактное расстояние включает сумму радиусов и 2,5 мм на каждую соседнюю пару.
Private Sub TestCheckCircleEffects(ByRef stats As TGeometryConfigStats, ByVal target As Object)
    Dim key As Variant, geom As CGeometryRegion, bars As CRebarLayout
    For Each key In Array("Circle.Diameter", "Rebar.AxisDistance", "Rebar.Count", "Rebar.Diameter", _
            "Rebar.Diameter2", "Rebar.Diameter3", "Rebar.Loc2row", "Rebar.Loc3row")
        TestConfigureCircle target
        Select Case CStr(key)
            Case "Circle.Diameter": TestSetKey target, CStr(key), 1100#
            Case "Rebar.AxisDistance": TestSetKey target, CStr(key), 60#
            Case "Rebar.Count": TestSetKey target, CStr(key), 10
            Case "Rebar.Diameter": TestSetKey target, CStr(key), 20#
            Case "Rebar.Diameter2": TestSetKey target, CStr(key), 20#
            Case "Rebar.Diameter3": TestSetKey target, CStr(key), 14#
            Case Else: TestSetKey target, CStr(key), "SideBySide"
        End Select
        TestReadGeometry geom, bars
        Dim prefix As String
        prefix = "circleConfig.effect." & CStr(key)
        Select Case CStr(key)
            Case "Circle.Diameter"
                TestNear stats, prefix & ".concreteExtent", geom.MaxX, 550#
                TestNear stats, prefix & ".barAxis", bars.Y(1), 500#
            Case "Rebar.AxisDistance"
                TestNear stats, prefix & ".barAxis", bars.Y(1), 440#
                TestNear stats, prefix & ".row2Axis", bars.Y(2), 423.5
            Case "Rebar.Count"
                TestCheck stats, prefix & ".allRows", bars.Count = 30
                TestNear stats, prefix & ".secondAngleX", bars.X(4), 450# * Sin(2# * GEOM_PI / 10#)
            Case "Rebar.Diameter"
                TestNear stats, prefix & ".diameter", bars.Diameter(1), 20#
                TestNear stats, prefix & ".row2Axis", bars.Y(2), 431.5
            Case "Rebar.Diameter2"
                TestNear stats, prefix & ".diameter", bars.Diameter(2), 20#
                TestNear stats, prefix & ".row3Skip", bars.Y(3), 412#
            Case "Rebar.Diameter3"
                TestNear stats, prefix & ".diameter", bars.Diameter(3), 14#
                TestNear stats, prefix & ".row3Axis", bars.Y(3), 418#
            Case "Rebar.Loc2row"
                TestNear stats, prefix & ".radialY", bars.Y(2), 450#
                TestNear stats, prefix & ".clockwiseX", bars.X(2), 16.5
                TestNear stats, prefix & ".row3NoSkip", bars.Y(3), 434.5
            Case "Rebar.Loc3row"
                TestNear stats, prefix & ".radialY", bars.Y(3), 450#
                TestNear stats, prefix & ".clockwiseX", bars.X(3), 15.5
                TestNear stats, prefix & ".row2Unchanged", bars.Y(2), 433.5
        End Select
        stats.Report = stats.Report & "CIRCLE_CONFIG_EFFECT: key=" & CStr(key) & "; bars=" & CStr(bars.Count) & vbCrLf
    Next key
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет ошибочные численные/enum значения на том же reader/registry-маршруте,
' затем обычную раскладку. Пустые/нулевые диаметры сохраняют утвержденное выключение ряда.
Private Sub TestCheckCircleBoundaries(ByRef stats As TGeometryConfigStats, ByVal target As Object)
    Dim key As Variant, value As Variant, errorNumber As Long, description As String
    Dim geom As CGeometryRegion, bars As CRebarLayout
    For Each key In Array("Circle.Diameter", "Rebar.AxisDistance", "Rebar.Count", "Rebar.Diameter", _
            "Rebar.Diameter2", "Rebar.Diameter3", "Rebar.Loc2row", "Rebar.Loc3row")
        TestConfigureCircle target
        If InStr(1, CStr(key), "Loc", vbBinaryCompare) > 0 Then value = "Unknown" Else value = -1
        TestSetKey target, CStr(key), value
        errorNumber = 0: description = vbNullString
        On Error Resume Next
        TestReadGeometry geom, bars
        errorNumber = Err.Number: description = Err.Description
        Err.Clear
        On Error GoTo 0
        TestCheck stats, "circleConfig.invalid." & CStr(key), errorNumber <> 0 And Len(description) > 0
        TestConfigureCircle target
        TestReadGeometry geom, bars
        TestCheck stats, "circleConfig.recovery." & CStr(key), bars.Count = 24
    Next key
    For Each key In Array("Rebar.Diameter", "Rebar.Diameter2", "Rebar.Diameter3")
        For Each value In Array(vbNullString, 0#)
            TestConfigureCircle target
            TestSetKey target, CStr(key), value
            TestReadGeometry geom, bars
            If CStr(key) = "Rebar.Diameter" Then
                TestCheck stats, "circleConfig.optional." & CStr(key) & "." & CStr(value), bars.Count = 0
            Else
                TestCheck stats, "circleConfig.optional." & CStr(key) & "." & CStr(value), bars.Count = 16
            End If
        Next value
    Next key
    TestConfigureCircle target
    TestSetKey target, "Rebar.Count", 0
    TestReadGeometry geom, bars
    TestCheck stats, "circleConfig.count.zero", bars.Count = 0
    TestSetKey target, "Rebar.Count", 2
    TestReadGeometry geom, bars
    TestCheck stats, "circleConfig.count.minimum", bars.Count = 6
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' При выбранном Rectangles диаметр бетонного круга не влияет на его форму и арматуру.
' Общие параметры армирования принадлежат выбранному типу и здесь не считаются неактивными.
Private Sub TestCheckCircleInactive(ByRef stats As TGeometryConfigStats, ByVal systemRange As Object, ByVal target As Object)
    TestConfigureRebarFieldFixture ThisWorkbook.Names("rngRectanglesGeometry").RefersToRange, "Rectangles", True
    TestSetKey systemRange, "Geometry.Type", "Rectangles"
    Dim originalGeom As CGeometryRegion, originalBars As CRebarLayout
    Dim geom As CGeometryRegion, bars As CRebarLayout, key As Variant, index As Long, equal As Boolean
    TestReadGeometry originalGeom, originalBars
    For Each key In Array("Circle.Diameter")
        If InStr(1, CStr(key), "Loc", vbBinaryCompare) > 0 Then
            TestSetKey target, CStr(key), "SideBySide"
        Else
            TestSetKey target, CStr(key), 30#
        End If
        TestReadGeometry geom, bars
        equal = geom.MinX = originalGeom.MinX And geom.MaxX = originalGeom.MaxX And _
            geom.MinY = originalGeom.MinY And geom.MaxY = originalGeom.MaxY And bars.Count = originalBars.Count
        If bars.Count = originalBars.Count Then
            For index = 1 To bars.Count
                If bars.X(index) <> originalBars.X(index) Or bars.Y(index) <> originalBars.Y(index) Or _
                        bars.Diameter(index) <> originalBars.Diameter(index) Then equal = False
            Next index
        End If
        TestCheck stats, "circleConfig.inactive." & CStr(key), equal
    Next key
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Читает текущие ячейки заново и использует штатные единицы/registry;
' прямые вызовы чистого Build здесь не могли бы обнаружить обрыв Config-передачи.
Private Sub TestReadGeometry(ByRef geom As CGeometryRegion, ByRef bars As CRebarLayout)
    Dim settings As CSystemSettingsReader, units As CUnitSystem, definition As CSectionDefinition, executor As CRebarRuleExecutor
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    Set definition = ReadWorkbookSectionDefinition(ThisWorkbook, settings, units)
    Set geom = New CGeometryRegion: geom.Initialize definition.ConcreteDomain
    Set executor = New CRebarRuleExecutor: Set bars = executor.Build(definition, geom)
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Меняет существующую ячейку по ключу таблицы. Отсутствующий ключ означает
' ошибку подготовки теста; новая настройка вместо него не создаётся.
Private Sub TestSetKey(ByVal target As Object, ByVal key As String, ByVal value As Variant)
    If Left$(key, 6) = "Rebar." Or Left$(key, 7) = "Circle." Then
        Dim circleCell As Object
        Set circleCell = TestFixtureCircleFieldCell(target.Worksheet.Parent, key, target): circleCell.Value2 = value: Exit Sub
    End If
    Dim rowIndex As Long
    For rowIndex = 2 To target.Rows.Count
        If CStr(target.Cells(rowIndex, 1).Value2) = key Then
            target.Cells(rowIndex, 2).Value2 = value
            Exit Sub
        End If
    Next rowIndex
    Err.Raise vbObjectError + 5960, "modTestGeometryConfig", "Не найдена строка fixture: " & key
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает восстановленные двумерные Formula-массивы без округления,
' включая пустые значения; формы исходных диапазонов не изменяются тестом.
Private Function TestSameFormula(ByVal actual As Variant, ByVal expected As Variant) As Boolean
    Dim rowIndex As Long, columnIndex As Long
    For rowIndex = LBound(expected, 1) To UBound(expected, 1)
        For columnIndex = LBound(expected, 2) To UBound(expected, 2)
            If CStr(actual(rowIndex, columnIndex)) <> CStr(expected(rowIndex, columnIndex)) Then Exit Function
        Next columnIndex
    Next rowIndex
    TestSameFormula = True
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Записывает абсолютный допуск 1e-6 мм и оба сравниваемых значения.
' Локальный допуск теста не изменяет настройки рабочего решателя.
Private Sub TestNear(ByRef stats As TGeometryConfigStats, ByVal name As String, ByVal actual As Double, ByVal expected As Double)
    TestCheck stats, name, Abs(actual - expected) <= 0.000001
    stats.Report = stats.Report & "CIRCLE_CONFIG_NUMBER: " & name & "; actual=" & CStr(actual) & "; expected=" & CStr(expected) & vbCrLf
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сохраняет счет и строку каждого assertion, не останавливая набор при несовпадении.
Private Sub TestCheck(ByRef stats As TGeometryConfigStats, ByVal name As String, ByVal condition As Boolean)
    If condition Then
        stats.Passed = stats.Passed + 1
        stats.Report = stats.Report & "OK: " & name & vbCrLf
    Else
        stats.Failed = stats.Failed + 1
        stats.Report = stats.Report & "FAIL: " & name & vbCrLf
    End If
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' ======================================================================
' ДЛЯ ТЕСТОВ: ОБЯЗАТЕЛЬНЫЙ ВВОД АКТИВНОЙ ГЕОМЕТРИИ
' ======================================================================

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет бетонную область и арматуру общего описания четырёх форм.
' Пропущенные пользовательские значения не подставляются. После серии,
' в том числе после ошибки, восстанавливаются все исходные формулы таблиц.
Public Function RunAudit03RequiredGeometryInputTests(Optional ByRef passed As Long = 0, _
        Optional ByRef failed As Long = 0) As String
    Dim allConfig As Object, savedConfig As Variant
    Set allConfig = ThisWorkbook.Worksheets("Config").UsedRange: savedConfig = TestFixtureSaveInputs(allConfig)
    Dim stats As TGeometryConfigStats, ranges(0 To 5) As Object, saved(0 To 5) As Variant
    Dim names As Variant, shapes As Variant, shape As String, fields As Variant
    Dim i As Long, j As Long, item As Variant, route As Variant, bad As Variant
    Dim cell As Object, baseline As Variant, current As Variant, description As String
    Dim errorNumber As Long, prefix As String, caseIndex As Long
    names = Array("rngSystemSettings", "rngUnitSettings", "rngCircleGeometry", _
        "rngRectanglesGeometry", "rngPointedSectionGeometry", "rngHollowRectangleGeometry")
    shapes = Array("Circle", "Rectangles", "PointedSection", "HollowRectangle")
    On Error GoTo FailedRun
    For i = 0 To 5
        Set ranges(i) = ThisWorkbook.Names.Item(CStr(names(i))).RefersToRange
        saved(i) = ranges(i).Formula
    Next i
    ranges(1).Cells(2, 2).Value2 = "mm"
    For i = 0 To 3
        shape = CStr(shapes(i))
        TestSetKey ranges(0), "Geometry.Type", shape
        If shape = "Circle" Then
            TestConfigureCircle ranges(i + 2)
            Set cell = TestRequiredGeometryKeyCell(ranges(i + 2), "Circle.Diameter")
            fields = Array(Array("Circle.Diameter", cell.Row - ranges(i + 2).Row + 1, cell.Column - ranges(i + 2).Column + 1, CDbl(cell.Value2)))
        Else
            TestConfigureShape ranges(i + 2), shape
            fields = TestShapeFields(shape)
            TestConfigureRequiredGeometryRebars ranges(i + 2), shape
        End If
        For Each route In Array("Geometry", "Rebars")
            baseline = TestRequiredGeometryRouteSnapshot(shape, CStr(route))
            For Each item In fields
                Set cell = ranges(i + 2).Cells(CLng(item(1)), CLng(item(2)))
                caseIndex = 0
                For Each bad In Array(vbNullString, "TODO", "abc", CVErr(2015))
                    caseIndex = caseIndex + 1
                    cell.Value2 = bad
                    prefix = "requiredGeometry." & CStr(item(0)) & "." & CStr(route) & "." & CStr(caseIndex)
                    TestRequiredGeometryProgress prefix, stats.Report
                    errorNumber = 0: description = vbNullString
                    On Error Resume Next
                    current = TestRequiredGeometryRouteSnapshot(shape, CStr(route))
                    errorNumber = Err.Number: description = Err.Description
                    Err.Clear
                    On Error GoTo FailedRun
                    TestCheck stats, prefix & ".reject", errorNumber <> 0
                    TestCheck stats, prefix & ".namedRange", InStr(1, description, CStr(names(i + 2)), vbBinaryCompare) > 0
                    TestCheck stats, prefix & ".address", InStr(1, description, cell.Worksheet.Name & "!", vbTextCompare) > 0 And _
                        InStr(1, description, cell.Address(False, False), vbTextCompare) > 0
                    stats.Report = stats.Report & "REQUIRED_GEOMETRY_ERROR: " & prefix & "; error=" & _
                        CStr(errorNumber) & "; " & description & vbCrLf
                    cell.Value2 = item(3)
                    current = TestRequiredGeometryRouteSnapshot(shape, CStr(route))
                    TestCheck stats, prefix & ".recovery", TestSameGeometryRouteSnapshot(baseline, current)
                Next bad
            Next item
        Next route
    Next i
    TestCheckRequiredGeometryInactive stats, ranges(0), ranges(3), ranges(4)
    TestCheckRequiredGeometryCommon stats, ranges(0), ranges(2)
    GoTo Restore
FailedRun:
    stats.Failed = stats.Failed + 1
    stats.Report = stats.Report & "FAIL: requiredGeometry.runtime; " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Restore:
    On Error GoTo FailedRestore
    For j = 0 To 5
        If Not ranges(j) Is Nothing Then
            ranges(j).Formula = saved(j)
            TestCheck stats, "requiredGeometry.restore." & CStr(names(j)), TestSameFormula(ranges(j).Formula, saved(j))
        End If
    Next j
    GoTo Finish
FailedRestore:
    stats.Failed = stats.Failed + 1
    stats.Report = stats.Report & "FAIL: requiredGeometry.restore.runtime; " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Finish:
    passed = stats.Passed: failed = stats.Failed
    TestFixtureRestoreInputs allConfig, savedConfig
    RunAudit03RequiredGeometryInputTests = stats.Report & "TOTAL_REQUIRED_GEOMETRY: passed=" & _
        CStr(passed) & "; failed=" & CStr(failed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Настраивает допустимые стержни для измененных размеров fixture. Не меняет
' тестируемые геометрические поля и не зависит от исходного армирования книги.
Private Sub TestConfigureRequiredGeometryRebars(ByVal target As Object, ByVal shape As String, Optional ByVal lengthScale As Double = 1#)
    Dim keyed As CSectionInputReader, key As Variant
    Set keyed = TestFixtureInput(target.Value2)
    For Each key In keyed.RowKeys("Main")
        TestFixtureInputCell(target, "Main", CStr(key), "as").Value2 = IIf(shape = "Rectangles", 20#, 40#) * lengthScale
        TestFixtureInputCell(target, "Main", CStr(key), "d").Value2 = 16# * lengthScale
        If Left$(CStr(key), 7) <> "Opening" Then TestFixtureInputCell(target, "Main", CStr(key), "n").Value2 = 3
        If shape = "Rectangles" Then
            TestFixtureInputCell(target, "Main", CStr(key), "s0").Value2 = 20# * lengthScale: TestFixtureInputCell(target, "Main", CStr(key), "s1").Value2 = 20# * lengthScale
        End If
        TestFixtureInputCell(target, "Extra", CStr(key), "d2").Value2 = 0#: TestFixtureInputCell(target, "Extra", CStr(key), "d3").Value2 = 0#
    Next key
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Читает полный Config и выбранный производственный вход, не вызывая фабрику
' формы перед builder-ом: так тест обнаруживает его собственные defaults.
' Снимок арматуры содержит все координаты/диаметры, а не только общий count.
Private Function TestRequiredGeometryRouteSnapshot(ByVal shape As String, ByVal route As String) As Variant
    Dim geom As CGeometryRegion, bars As CRebarLayout
    TestReadGeometry geom, bars
    If route = "Geometry" Then
        TestRequiredGeometryRouteSnapshot = TestShapeSnapshot(geom)
        Exit Function
    End If
    Dim values() As Double, i As Long
    ReDim values(0 To bars.Count * 3)
    values(0) = bars.Count
    For i = 1 To bars.Count
        values(i * 3 - 2) = bars.X(i): values(i * 3 - 1) = bars.Y(i): values(i * 3) = bars.Diameter(i)
    Next i
    TestRequiredGeometryRouteSnapshot = values
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает расчетные snapshots в том же абсолютном геометрическом допуске,
' что и shape-tests; неодинаковый размер массива означает различие раскладки.
Private Function TestSameGeometryRouteSnapshot(ByVal first As Variant, ByVal second As Variant) As Boolean
    If UBound(first) <> UBound(second) Then Exit Function
    Dim i As Long
    For i = LBound(first) To UBound(first)
        If Abs(CDbl(first(i)) - CDbl(second(i))) > 0.000001 Then Exit Function
    Next i
    TestSameGeometryRouteSnapshot = True
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Находит фактическую ячейку параметра в таблице. Отсутствующая строка —
' ошибка теста; значение по умолчанию не подставляется.
Private Function TestRequiredGeometryKeyCell(ByVal target As Object, ByVal key As String) As Object
    If Left$(key, 6) = "Rebar." Or Left$(key, 7) = "Circle." Then
        Set TestRequiredGeometryKeyCell = TestFixtureCircleFieldCell(target.Worksheet.Parent, key, target): Exit Function
    End If
    Dim item As Variant, shape As String, rowIndex As Long
    shape = Split(key, ".")(0)
    If shape <> "Rectangles" And shape <> "PointedSection" And shape <> "HollowRectangle" Then
        For rowIndex = 2 To target.Rows.Count
            If CStr(target.Cells(rowIndex, 1).Value2) = key Then Set TestRequiredGeometryKeyCell = target.Cells(rowIndex, 2): Exit Function
        Next rowIndex
        Err.Raise vbObjectError + 5965, , "Не найдено контрольное поле " & key
    End If
    For Each item In TestShapeFields(shape)
        If CStr(item(0)) = key Then Set TestRequiredGeometryKeyCell = target.Cells(CLng(item(1)), CLng(item(2))): Exit Function
    Next item
    Err.Raise vbObjectError + 5965, , "Не найдено контрольное поле " & key
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет неактивные размеры Rectangle и радиусы R2 сторон без выступа. Эти значения
' не участвуют в форме/раскладке, поэтому не должны требовать числового ввода.
' Ошибки формул Excel здесь не подменяются пустым значением reader-а.
Private Sub TestCheckRequiredGeometryInactive(ByRef stats As TGeometryConfigStats, ByVal systemRange As Object, _
        ByVal rectRange As Object, ByVal roundedRange As Object)
    Dim route As Variant, item As Variant, value As Variant, baseline As Variant, current As Variant
    Dim description As String, errorNumber As Long, prefix As String
    TestSetKey systemRange, "Geometry.Type", "Rectangles"
    TestConfigureShape rectRange, "Rectangles": TestConfigureRequiredGeometryRebars rectRange, "Rectangles"
    rectRange.Cells(4, 3).Value2 = "OneRect"
    For Each route In Array("Geometry", "Rebars")
        baseline = TestRequiredGeometryRouteSnapshot("Rectangles", CStr(route))
        For Each item In Array(Array(7, 3), Array(8, 3), Array(9, 3))
            For Each value In Array(vbNullString, "TODO", "abc")
                rectRange.Cells(CLng(item(0)), CLng(item(1))).Value2 = value
                errorNumber = 0: description = vbNullString
                On Error Resume Next
                current = TestRequiredGeometryRouteSnapshot("Rectangles", CStr(route))
                errorNumber = Err.Number: description = Err.Description
                Err.Clear: On Error GoTo 0
                prefix = "requiredGeometry.inactive.Rectangle." & CStr(route) & "." & CStr(item(0)) & "." & CStr(item(1)) & "." & CStr(value)
                TestCheck stats, prefix, errorNumber = 0
                If errorNumber = 0 Then TestCheck stats, prefix & ".snapshot", TestSameGeometryRouteSnapshot(baseline, current)
                stats.Report = stats.Report & "INACTIVE_GEOMETRY: " & prefix & "; error=" & CStr(errorNumber) & "; " & description & vbCrLf
            Next value
            rectRange.Cells(CLng(item(0)), CLng(item(1))).Value2 = 100#
        Next item
    Next route
    TestSetKey systemRange, "Geometry.Type", "PointedSection"
    TestConfigureShape roundedRange, "PointedSection": TestConfigureRequiredGeometryRebars roundedRange, "PointedSection"
    roundedRange.Cells(6, 3).Value2 = 0#: roundedRange.Cells(9, 3).Value2 = 0#
    For Each route In Array("Geometry", "Rebars")
        baseline = TestRequiredGeometryRouteSnapshot("PointedSection", CStr(route))
        For Each item In Array(Array(8, 3), Array(11, 3))
            For Each value In Array(vbNullString, "TODO", "abc")
                roundedRange.Cells(CLng(item(0)), CLng(item(1))).Value2 = value
                errorNumber = 0: description = vbNullString
                On Error Resume Next
                current = TestRequiredGeometryRouteSnapshot("PointedSection", CStr(route))
                errorNumber = Err.Number: description = Err.Description
                Err.Clear: On Error GoTo 0
                prefix = "requiredGeometry.inactive.W0." & CStr(route) & "." & CStr(item(0)) & "." & CStr(item(1)) & "." & CStr(value)
                TestCheck stats, prefix, errorNumber = 0
                If errorNumber = 0 Then TestCheck stats, prefix & ".snapshot", TestSameGeometryRouteSnapshot(baseline, current)
                stats.Report = stats.Report & "INACTIVE_GEOMETRY: " & prefix & "; error=" & CStr(errorNumber) & "; " & description & vbCrLf
            Next value
            roundedRange.Cells(CLng(item(0)), CLng(item(1))).Value2 = 0#
        Next item
    Next route
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Обязательные общие поля Geometry и ручные шаги в режиме User проверяет штатная сборка модели.
' В missing-варианте временно меняется ключ, а не удаляются пользовательские
' строки; recovery восстанавливает его до следующего сценария.
Private Sub TestCheckRequiredGeometryCommon(ByRef stats As TGeometryConfigStats, ByVal systemRange As Object, ByVal circleRange As Object)
    Dim key As Variant, bad As Variant, cell As Object, savedValue As Variant
    Dim settings As CSystemSettingsReader, units As CUnitSystem, model As CSectionModel
    Dim errorNumber As Long, description As String, prefix As String, caseIndex As Long
    TestSetKey systemRange, "Geometry.Type", "Circle": TestSetKey systemRange, "Geometry.Source", "Generated"
    TestSetKey systemRange, "Mesh.StepMode", "User" ' Ручные шаги проверяются как активные входы; внешний тест возвращает исходные формулы, включая режим Auto.
    TestSetKey systemRange, "Mesh.CellGeometry", "Adaptive" ' Число уровней границы — активный обязательный вход Adaptive; внешний suite восстанавливает исходный режим сетки.
    TestConfigureCircle circleRange
    For Each key In Array("Geometry.Type", "Geometry.Source", "Mesh.StepX", "Mesh.StepY", "Mesh.BoundarySubdivisions")
        Set cell = TestRequiredGeometryKeyCell(systemRange, CStr(key))
        savedValue = cell.Formula: caseIndex = 0
        For Each bad In Array(vbNullString, "TODO", "abc", CVErr(2015), "__MISSING_KEY__")
            caseIndex = caseIndex + 1
            If Not IsError(bad) Then
                If CStr(bad) = "__MISSING_KEY__" Then cell.Offset(0, -1).Value2 = "Audit03.Hidden." & CStr(key) Else cell.Value2 = bad
            Else
                cell.Value2 = bad
            End If
            errorNumber = 0: description = vbNullString
            TestRequiredGeometryProgress "requiredGeometry.common." & CStr(key) & "." & CStr(caseIndex), stats.Report
            On Error Resume Next
            Set model = TestRequiredGeometryModelFromWorkbook()
            errorNumber = Err.Number: description = Err.Description
            Err.Clear: On Error GoTo 0
            prefix = "requiredGeometry.common." & CStr(key) & "." & CStr(caseIndex)
            TestCheck stats, prefix & ".reject", errorNumber <> 0
            TestCheck stats, prefix & ".key", InStr(1, description, CStr(key), vbBinaryCompare) > 0
            TestCheck stats, prefix & ".location", InStr(1, description, cell.Worksheet.Name & "!", vbTextCompare) > 0 Or _
                InStr(1, description, "лист " & cell.Worksheet.Name, vbTextCompare) > 0 ' Reader общих настроек называет лист и ячейку словами; reader геометрии использует Sheet!Cell.
            If caseIndex < 5 Then TestCheck stats, prefix & ".address", InStr(1, description, cell.Address(False, False), vbTextCompare) > 0
            stats.Report = stats.Report & "REQUIRED_GEOMETRY_ERROR: " & prefix & "; error=" & CStr(errorNumber) & "; " & description & vbCrLf
            cell.Offset(0, -1).Value2 = key: cell.Formula = savedValue
            Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
            Set units = New CUnitSystem: units.LoadFromSettings settings
            Set model = BuildWorkbookSectionModel(ThisWorkbook, settings, units)
            TestCheck stats, prefix & ".recovery", model.ConcreteCount > 0 And model.RebarCount = 24
        Next bad
    Next key
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Выполняет обычное чтение, преобразование единиц и построение без подавления
' ошибок между шагами. После первой ошибки последующие шаги не выполняются;
' вызывающий тест сохраняет причину отказа всего вызова.
Private Function TestRequiredGeometryModelFromWorkbook() As CSectionModel
    Dim settings As CSystemSettingsReader, units As CUnitSystem
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    Set TestRequiredGeometryModelFromWorkbook = BuildWorkbookSectionModel(ThisWorkbook, settings, units)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сохраняет уже выполненные assertions и следующий сценарий рядом с временной
' книгой. Watchdog забирает этот журнал при timeout; отметка START не считается
' успешной проверкой. Ошибка записи диагностики не изменяет результат теста.
Private Sub TestRequiredGeometryProgress(ByVal nextCase As String, ByVal report As String)
    On Error Resume Next
    Dim fso As Object, stream As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set stream = fso.CreateTextFile(ThisWorkbook.Path & "\Audit03_Search_Progress.txt", True, True)
    stream.Write report & "START_REQUIRED_GEOMETRY: " & nextCase & vbCrLf
    stream.Close
    On Error GoTo 0
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' ==========================================================================
' ПОАДРЕСНЫЕ ПАРАМЕТРЫ ГРАНЕЙ И ДОПОЛНИТЕЛЬНЫХ РЯДОВ
' ==========================================================================
' Читает каждую физическую грань через настоящий Config/registry pipeline.
' Прямоугольные контуры дают независимый oracle нормали, диаметра и количества;
' арматура Opening проверяется в направлении от пустоты в бетон. Все таблицы
' восстанавливаются целиком; этот набор не решает НДС и не назначает его статус.
Public Function RunAudit03RebarFieldConfigTests(Optional ByRef passed As Long = 0, _
        Optional ByRef failed As Long = 0) As String
    Dim allConfig As Object, savedConfig As Variant
    Set allConfig = ThisWorkbook.Worksheets("Config").UsedRange: savedConfig = TestFixtureSaveInputs(allConfig)
    Dim stats As TGeometryConfigStats, ranges(0 To 5) As Object, saved(0 To 5) As Variant
    Dim names As Variant, shapes As Variant, shape As String, face As Variant, field As Variant
    Dim i As Long, cases As Long, target As Object
    names = Array("rngSystemSettings", "rngUnitSettings", "rngRectanglesGeometry", _
        "rngPointedSectionGeometry", "rngHollowRectangleGeometry", "rngCircleGeometry")
    shapes = Array("Rectangles", "PointedSection", "HollowRectangle")
    On Error GoTo FailedRun
    For i = 0 To 5
        Set ranges(i) = ThisWorkbook.Names.Item(CStr(names(i))).RefersToRange
        saved(i) = ranges(i).Formula
    Next i
    ranges(1).Cells(2, 2).Value2 = "mm"
    TestSetKey ranges(0), "Geometry.Source", "Generated"
    For i = 0 To 2
        shape = CStr(shapes(i)): Set target = ranges(i + 2)
        TestSetKey ranges(0), "Geometry.Type", shape
        For Each face In TestRebarFieldFaces(shape)
            For Each field In TestRebarFaceFields(shape, face)
                TestConfigureRebarFieldFixture target, shape, CBool(field(3))
                TestCheckRebarFieldEffect stats, target, shape, face, field, cases
                TestCheckRebarFieldInvalid stats, target, shape, face, field, cases
            Next field
        Next face
    Next i
    TestCheckInactiveRebarFields stats, ranges(0), ranges(5), ranges(2), ranges(3), ranges(4), cases
    GoTo Restore
FailedRun:
    stats.Failed = stats.Failed + 1
    stats.Report = stats.Report & "FAIL: rebarFields.runtime; " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Restore:
    On Error GoTo FailedRestore
    For i = 0 To 5
        If Not ranges(i) Is Nothing Then
            ranges(i).Formula = saved(i)
            TestCheck stats, "rebarFields.restore." & CStr(names(i)), TestSameFormula(ranges(i).Formula, saved(i))
        End If
    Next i
    GoTo Finish
FailedRestore:
    stats.Failed = stats.Failed + 1
    stats.Report = stats.Report & "FAIL: rebarFields.restore.runtime; " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Finish:
    passed = stats.Passed: failed = stats.Failed
    TestFixtureRestoreInputs allConfig, savedConfig
    RunAudit03RebarFieldConfigTests = stats.Report & "TOTAL_REBAR_FIELD_CONFIG: passed=" & CStr(passed) & _
        "; failed=" & CStr(failed) & "; cases=" & CStr(cases) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет неактивность на настоящем workbook-маршруте, а не только в Build.
' Общие loc/bind Rectangles не считаются неактивными, пока активна вторая сторона;
' для этого сценария выключаются оба первых ряда соответствующей пары граней.
' as/отступы отсутствующего ряда здесь не объявляются неактивными автоматически:
' например cover Opening может ограничивать проекции соседних стержней.
Private Sub TestCheckInactiveRebarFields(ByRef stats As TGeometryConfigStats, ByVal systemRange As Object, _
        ByVal circleRange As Object, ByVal rectRange As Object, ByVal roundedRange As Object, _
        ByVal hollowRange As Object, ByRef cases As Long)
    Dim shape As Variant, face As Variant, field As Variant, mode As Variant, target As Object
    Dim kind As String, rowIndex As Long, pairRow As Long
    For Each shape In Array("Rectangles", "PointedSection", "HollowRectangle")
        Select Case CStr(shape)
            Case "Rectangles": Set target = rectRange
            Case "PointedSection": Set target = roundedRange
            Case "HollowRectangle": Set target = hollowRange
        End Select
        TestSetKey systemRange, "Geometry.Type", CStr(shape)
        For Each face In TestRebarFieldFaces(CStr(shape))
            For Each field In TestRebarFaceFields(CStr(shape), face)
                kind = CStr(field(0))
                If Left$(kind, 3) = "loc" Or Left$(kind, 4) = "bind" Then
                    TestConfigureRebarFieldFixture target, CStr(shape), False
                    TestCheckInactiveRebarCell stats, target, CStr(shape), face, field, "extraRowsOff", cases
                End If
                ' as остаётся параметром общего контура при n=0/d=0; пустой as равен нулю.
                ' Концевые отступы отключённого наружного ряда и параметры отсутствующих extras не используются.
                Dim inactiveOffset As Boolean
                inactiveOffset = (kind = "start" Or kind = "end")
                If CStr(shape) = "HollowRectangle" And Left$(CStr(face(0)), 8) = "Opening." Then inactiveOffset = False
                If kind = "as" Or kind = "d2" Or kind = "d3" Or Left$(kind, 3) = "loc" Or Left$(kind, 4) = "bind" Or inactiveOffset Then
                    For Each mode In Array("diameterOff", "countOff")
                        If CStr(mode) <> "countOff" Or Left$(CStr(face(0)), 8) <> "Opening." Then
                            TestConfigureRebarFieldFixture target, CStr(shape), True
                            rowIndex = CLng(face(1))
                            If CStr(mode) = "countOff" Then
                                TestRebarFieldCell(target, rowIndex, 5).Value2 = 0
                            Else
                                TestRebarFieldCell(target, rowIndex, 4).Value2 = 0#
                            End If
                            If CStr(shape) = "Rectangles" Then
                                pairRow = rowIndex + 1
                                If (rowIndex - 5) Mod 2 <> 0 Then pairRow = rowIndex - 1
                                If CStr(mode) = "countOff" Then
                                    TestRebarFieldCell(target, pairRow, 5).Value2 = 0
                                Else
                                    TestRebarFieldCell(target, pairRow, 4).Value2 = 0#
                                End If
                            End If
                            TestCheckInactiveRebarCell stats, target, CStr(shape), face, field, CStr(mode), cases
                        End If
                    Next mode
                End If
                If CStr(shape) = "Rectangles" And CLng(face(1)) >= 9 Then
                    TestConfigureRebarFieldFixture target, "Rectangles", True
                    TestFixtureInputCell(target, "General", "Shape").Value2 = "OneRect"
                    TestCheckInactiveRebarCell stats, target, "Rectangles", face, field, "lowerRectangleOff", cases
                End If
            Next field
        Next face
        modTestShapeArchitecture.TestConfigureShapeFixture "Circle" ' Ошибка в армировании невыбранного типа не влияет на круг.
        face = TestRebarFieldFaces(CStr(shape))(0)
        field = TestRebarFaceFields(CStr(shape), face)(0)
        TestCheckInactiveRebarCell stats, target, "Circle", face, field, "otherType", cases
    Next shape
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет поле отключённого ряда: as остаётся отступом общего контура, пустой равен нулю.
' Остальные неиспользуемые поля не вмешиваются в раскладку. Текст/ошибка формулы/отрицательный as
' должны отклоняться с адресом; после каждого случая восстанавливается полный снимок стержней.
Private Sub TestCheckInactiveRebarCell(ByRef stats As TGeometryConfigStats, ByVal target As Object, _
        ByVal activeShape As String, ByVal face As Variant, ByVal field As Variant, _
        ByVal mode As String, ByRef cases As Long)
    Dim cell As Object, savedValue As Variant, bad As Variant, baseline As Variant, current As Variant
    Dim prefix As String, number As Long, description As String, index As Long
    Dim contourOffset As Boolean, zeroSnapshot As Variant, badValues As Variant
    Set cell = TestRebarFieldCell(target, CLng(field(1)), CLng(field(2)))
    savedValue = cell.Formula
    baseline = TestRequiredGeometryRouteSnapshot(activeShape, "Rebars")
    contourOffset = (CStr(field(0)) = "as" And (mode = "diameterOff" Or mode = "countOff"))
    If contourOffset Then
        cell.Value2 = 0#: zeroSnapshot = TestRequiredGeometryRouteSnapshot(activeShape, "Rebars")
        cell.Formula = savedValue
        badValues = Array(vbNullString, 0#, 50#, -1#, "abc", CVErr(2015)) ' Fixture задаёт as=50 мм; первые три случая проверяют допустимый ввод отключённой стороны.
    Else
        badValues = Array(vbNullString, "abc", CVErr(2015))
    End If
    For Each bad In badValues
        index = index + 1: cases = cases + 1
        prefix = "rebarFields.inactive." & mode & "." & CStr(face(0)) & "." & CStr(field(0)) & "." & CStr(index)
        cell.Value2 = bad: number = 0: description = vbNullString
        TestRequiredGeometryProgress prefix, stats.Report
        On Error Resume Next
        current = TestRequiredGeometryRouteSnapshot(activeShape, "Rebars")
        number = Err.Number: description = Err.Description
        Err.Clear: On Error GoTo 0
        If contourOffset Then
            If index <= 3 Then
                TestCheck stats, prefix & ".accepted", number = 0
                If number = 0 Then
                    If index <= 2 Then
                        TestCheck stats, prefix & ".blankEqualsZero", TestSameGeometryRouteSnapshot(zeroSnapshot, current)
                    Else
                        TestCheck stats, prefix & ".givenAsRestored", TestSameGeometryRouteSnapshot(baseline, current)
                    End If
                End If
            Else
                TestCheck stats, prefix & ".reject", number <> 0
                TestCheck stats, prefix & ".address", InStr(1, description, cell.Address(False, False), vbTextCompare) > 0
            End If
        Else
            TestCheck stats, prefix & ".noInterference", number = 0
            If number = 0 Then TestCheck stats, prefix & ".snapshot", TestSameGeometryRouteSnapshot(baseline, current)
        End If
        stats.Report = stats.Report & "INACTIVE_REBAR_FIELD: " & prefix & "; error=" & CStr(number) & "; " & description & vbCrLf
        cell.Formula = savedValue
        current = TestRequiredGeometryRouteSnapshot(activeShape, "Rebars")
        TestCheck stats, prefix & ".recovery", TestSameGeometryRouteSnapshot(baseline, current)
    Next bad
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Описывает физические грани: имя layout-группы, строка первого/дополнительных
' рядов, нормаль внутрь бетона и касательная обхода. Это независимые сведения
' прямоугольного fixture, а не вызов production-функции смещения ряда.
Private Function TestRebarFieldFaces(ByVal shape As String) As Variant
    Select Case shape
        Case "Rectangles"
            TestRebarFieldFaces = Array(Array("H1.as_1", 7, 20, 1#, 0#, 0#, 1#), _
                Array("H1.as_2", 8, 21, -1#, 0#, 0#, -1#), _
                Array("B1.as_1", 5, 18, 0#, -1#, 1#, 0#), _
                Array("B1.as_2", 6, 19, 0#, 1#, -1#, 0#), _
                Array("H2.as_1", 11, 24, 1#, 0#, 0#, 1#), _
                Array("H2.as_2", 12, 25, -1#, 0#, 0#, -1#), _
                Array("B2.as_1", 9, 22, 0#, -1#, 1#, 0#), _
                Array("B2.as_2", 10, 23, 0#, 1#, -1#, 0#))
        Case "PointedSection"
            TestRebarFieldFaces = Array(Array("H.Left", 7, 20, 1#, 0#, 0#, 1#), _
                Array("H.Right", 8, 21, -1#, 0#, 0#, -1#), _
                Array("B.Top", 5, 18, 0#, -1#, 1#, 0#), _
                Array("B.Bottom", 6, 19, 0#, 1#, -1#, 0#))
        Case "HollowRectangle"
            TestRebarFieldFaces = Array(Array("H.Left", 7, 20, 1#, 0#, 0#, 1#), _
                Array("H.Right", 8, 21, -1#, 0#, 0#, -1#), _
                Array("B.Top", 5, 18, 0#, -1#, 1#, 0#), _
                Array("B.Bottom", 6, 19, 0#, 1#, -1#, 0#), _
                Array("Opening.H.Left", 11, 24, -1#, 0#, 0#, 1#), _
                Array("Opening.H.Right", 12, 25, 1#, 0#, 0#, -1#), _
                Array("Opening.B.Top", 9, 22, 0#, 1#, 1#, 0#), _
                Array("Opening.B.Bottom", 10, 23, 0#, -1#, -1#, 0#))
        Case Else
            Err.Raise vbObjectError + 5966, "modTestGeometryConfig", "Неизвестная форма арматурного fixture: " & shape
    End Select
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Перечисляет настоящие input-ячейки одной грани: тип поля, относительная
' строка/колонка и необходимость активных дополнительных рядов. Количество
' Opening отсутствует по контракту и не превращается в искусственный ввод.
Private Function TestRebarFaceFields(ByVal shape As String, ByVal face As Variant) As Variant
    Dim mainRow As Long, extraRow As Long
    mainRow = CLng(face(1)): extraRow = CLng(face(2))
    If shape = "Rectangles" Then
        TestRebarFaceFields = Array(Array("as", mainRow, 3, False), Array("d", mainRow, 4, False), _
            Array("n", mainRow, 5, False), Array("start", mainRow, 6, False), Array("end", mainRow, 7, False), _
            Array("d2", extraRow, 3, True), Array("loc2", extraRow, 6, True), Array("bind2", extraRow, 5, True), _
            Array("d3", extraRow, 7, True), Array("loc3", extraRow, 10, True), Array("bind3", extraRow, 9, True))
    ElseIf Left$(CStr(face(0)), 8) = "Opening." Then
        TestRebarFaceFields = Array(Array("as", mainRow, 3, False), Array("d", mainRow, 4, False), _
            Array("d2", extraRow, 3, True), Array("loc2", extraRow, 6, True), Array("bind2", extraRow, 5, True), _
            Array("d3", extraRow, 7, True), Array("loc3", extraRow, 10, True), Array("bind3", extraRow, 9, True))
    Else
        TestRebarFaceFields = Array(Array("as", mainRow, 3, False), Array("d", mainRow, 4, False), Array("n", mainRow, 5, False), _
            Array("d2", extraRow, 3, True), Array("loc2", extraRow, 6, True), Array("bind2", extraRow, 5, True), _
            Array("d3", extraRow, 7, True), Array("loc3", extraRow, 10, True), Array("bind3", extraRow, 9, True))
    End If
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Задает достаточно просторные прямые контуры и пять базовых стержней каждой
' грани. Дополнительные ряды при необходимости имеют разные диаметры 12/10.
' Концевые отступы 100 мм оставляют запас для контактного смещения до 35 мм:
' здесь проверяем поля ввода, а пропуски у углов — в RunFullBoundaryOffsetTests.
Private Sub TestConfigureRebarFieldFixture(ByVal target As Object, ByVal shape As String, ByVal extras As Boolean)
    TestFixtureSelectRecipeType ThisWorkbook, shape
    TestConfigureShape target, shape
    Dim face As Variant, mainRow As Long, extraRow As Long, column As Variant
    Select Case shape
        Case "Rectangles"
            target.Cells(4, 3).Value2 = "TwoRect": TestFixtureInputCell(target, "General", "XOffset").Value2 = 0#
            For Each column In Array("H1", "B1", "H2", "B2"): TestFixtureInputCell(target, "General", CStr(column)).Value2 = 800#: Next column
        Case "PointedSection"
            target.Cells(4, 3).Value2 = 800#: target.Cells(5, 3).Value2 = 800#
            For column = 2 To 3
                TestFixtureInputCell(target, "Detail", "W", IIf(column = 2, "left", "right")).Value2 = 0#
                TestFixtureInputCell(target, "Detail", "R1", IIf(column = 2, "left", "right")).Value2 = 0#: TestFixtureInputCell(target, "Detail", "R2", IIf(column = 2, "left", "right")).Value2 = 0#
            Next column
        Case "HollowRectangle"
            TestFixtureInputCell(target, "General", "Xo").Value2 = 0#: TestFixtureInputCell(target, "General", "Yo").Value2 = 0#
            TestFixtureInputCell(target, "General", "H").Value2 = 1600#: TestFixtureInputCell(target, "General", "B").Value2 = 1600#
            TestFixtureInputCell(target, "General", "R").Value2 = 0#: TestFixtureInputCell(target, "General", "Ro").Value2 = 0#
            TestFixtureInputCell(target, "General", "Ho").Value2 = 1000#: TestFixtureInputCell(target, "General", "Bo").Value2 = 1000#
    End Select
    For Each face In TestRebarFieldFaces(shape)
        mainRow = CLng(face(1)): extraRow = CLng(face(2))
        TestRebarFieldCell(target, extraRow, 4).Value2 = 0#: TestRebarFieldCell(target, extraRow, 8).Value2 = 0#
        TestRebarFieldCell(target, mainRow, 3).Value2 = 50#: TestRebarFieldCell(target, mainRow, 4).Value2 = 16#
        If Left$(CStr(face(0)), 8) <> "Opening." Then
            TestRebarFieldCell(target, mainRow, 5).Value2 = 5
            TestRebarFieldCell(target, mainRow, 6).Value2 = 100#: TestRebarFieldCell(target, mainRow, 7).Value2 = 100#
        End If
        If extras Then
            TestRebarFieldCell(target, extraRow, 3).Value2 = 12#: TestRebarFieldCell(target, extraRow, 7).Value2 = 10#
        Else
            TestRebarFieldCell(target, extraRow, 3).Value2 = 0#: TestRebarFieldCell(target, extraRow, 7).Value2 = 0#
        End If
        TestRebarFieldCell(target, extraRow, 6).Value2 = "Stacked"
        TestRebarFieldCell(target, extraRow, 10).Value2 = "Stacked"
        TestRebarFieldCell(target, extraRow, 5).Value2 = "EachBar"
        TestRebarFieldCell(target, extraRow, 9).Value2 = "EachBar"
    Next face
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Меняет одно поле и проверяет предсказуемый физический эффект. Для as/t
' проверяется подписанное перемещение, для d/n независимые diameter/count,
' для loc/bind положение и число дополнительных стержней относительно первого.
Private Sub TestCheckRebarFieldEffect(ByRef stats As TGeometryConfigStats, ByVal target As Object, _
        ByVal shape As String, ByVal face As Variant, ByVal field As Variant, ByRef cases As Long)
    Dim geom As CGeometryRegion, baseline As CRebarLayout, bars As CRebarLayout, cell As Object
    Dim before As Variant, after As Variant, kind As String, prefix As String, value As Variant
    Dim rowNumber As Long, expectedCount As Long, expectedDistance As Double
    kind = CStr(field(0)): Set cell = TestRebarFieldCell(target, CLng(field(1)), CLng(field(2)))
    prefix = "rebarFields." & shape & "." & CStr(face(0)) & "." & kind
    TestReadGeometry geom, baseline
    before = TestRebarFieldRowSummary(baseline, CStr(face(0)), 1)
    TestCheck stats, prefix & ".activeBaseline", CLng(before(0)) > 1
    Select Case kind
        Case "as": value = 60#
        Case "d": value = 20#
        Case "n": value = 7
        Case "start", "end": value = 120#
        Case "d2": value = 18#
        Case "d3": value = 20#
        Case "loc2", "loc3": value = "SideBySide"
        Case "bind2", "bind3": value = "EverySecondBar"
    End Select
    cell.Value2 = value: cases = cases + 1
    TestRequiredGeometryProgress prefix & ".effect", stats.Report
    TestReadGeometry geom, bars
    after = TestRebarFieldRowSummary(bars, CStr(face(0)), 1)
    Select Case kind
        Case "as"
            TestNear stats, prefix & ".normalX", CDbl(after(2)) - CDbl(before(2)), 10# * CDbl(face(3))
            TestNear stats, prefix & ".normalY", CDbl(after(3)) - CDbl(before(3)), 10# * CDbl(face(4))
            TestNear stats, prefix & ".count", CDbl(after(0)), CDbl(before(0))
        Case "d"
            TestNear stats, prefix & ".diameter", CDbl(after(1)), 20#
            TestNear stats, prefix & ".count", CDbl(after(0)), CDbl(before(0))
            TestNear stats, prefix & ".xUnchanged", CDbl(after(2)), CDbl(before(2))
            TestNear stats, prefix & ".yUnchanged", CDbl(after(3)), CDbl(before(3))
        Case "n"
            TestNear stats, prefix & ".count", CDbl(after(0)), 7#
        Case "start", "end"
            expectedDistance = 10#: If kind = "end" Then expectedDistance = -10#
            TestNear stats, prefix & ".tangentX", CDbl(after(2)) - CDbl(before(2)), expectedDistance * CDbl(face(5))
            TestNear stats, prefix & ".tangentY", CDbl(after(3)) - CDbl(before(3)), expectedDistance * CDbl(face(6))
            TestNear stats, prefix & ".count", CDbl(after(0)), CDbl(before(0))
        Case Else
            rowNumber = 2: If Right$(kind, 1) = "3" Then rowNumber = 3
            after = TestRebarFieldRowSummary(bars, CStr(face(0)), rowNumber)
            expectedCount = CLng(before(0))
            If Left$(kind, 4) = "bind" Then expectedCount = (expectedCount + 1) \ 2
            TestNear stats, prefix & ".count", CDbl(after(0)), CDbl(expectedCount)
            expectedDistance = 16.5: If rowNumber = 3 Then expectedDistance = 30# ' В каждой приставленной паре к сумме радиусов добавлено 2,5 мм.
            If kind = "d2" Then expectedDistance = 19.5
            If kind = "d3" Then expectedDistance = 35#
            If Left$(kind, 3) = "loc" Then
                expectedDistance = 15.5
                If rowNumber = 2 Then expectedDistance = 16.5
            End If
            If kind = "d2" Or kind = "d3" Then TestNear stats, prefix & ".diameter", CDbl(after(1)), CDbl(value)
            TestCheckRebarRowOffsets stats, prefix, bars, CStr(face(0)), rowNumber, face, _
                expectedDistance, Left$(kind, 3) <> "loc"
    End Select
    If kind = "d" Or kind = "n" Or kind = "d2" Or kind = "d3" Then
        cell.ClearContents: cases = cases + 1
        TestReadGeometry geom, bars
        rowNumber = 1: If kind = "d2" Then rowNumber = 2
        If kind = "d3" Then rowNumber = 3
        after = TestRebarFieldRowSummary(bars, CStr(face(0)), rowNumber)
        TestNear stats, prefix & ".blankDisables", CDbl(after(0)), 0#
        cell.Value2 = 0#: cases = cases + 1
        TestReadGeometry geom, bars
        after = TestRebarFieldRowSummary(bars, CStr(face(0)), rowNumber)
        TestNear stats, prefix & ".zeroDisables", CDbl(after(0)), 0#
    End If
    TestConfigureRebarFieldFixture target, shape, CBool(field(3))
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет активный ошибочный ввод и восстановление того же reader/builder
' маршрута. Пустой s0/s1 должен совпасть с явным нулём; отрицательный, текст и ошибка отклоняются.
' Для активного as и loc/bind пустота остаётся ошибкой; допустимая пустота d/n проверена отдельно.
Private Sub TestCheckRebarFieldInvalid(ByRef stats As TGeometryConfigStats, ByVal target As Object, _
        ByVal shape As String, ByVal face As Variant, ByVal field As Variant, ByRef cases As Long)
    Dim bad As Variant, badValues As Variant, kind As String, prefix As String, cell As Object
    Dim savedValue As Variant, geom As CGeometryRegion, bars As CRebarLayout, baseline As Variant, current As Variant
    Dim number As Long, description As String, index As Long, zeroSnapshot As Variant
    kind = CStr(field(0)): Set cell = TestRebarFieldCell(target, CLng(field(1)), CLng(field(2)))
    savedValue = cell.Formula: baseline = TestRequiredGeometryRouteSnapshot(shape, "Rebars")
    If Left$(kind, 3) = "loc" Or Left$(kind, 4) = "bind" Then
        badValues = Array(vbNullString, "Wrong", CVErr(2015))
    ElseIf kind = "as" Or kind = "start" Or kind = "end" Then
        badValues = Array(vbNullString, -1#, "abc", CVErr(2015))
    Else
        badValues = Array(-1#, "abc", CVErr(2015))
    End If
    For Each bad In badValues
        index = index + 1: cases = cases + 1
        prefix = "rebarFields." & shape & "." & CStr(face(0)) & "." & kind & ".invalid" & CStr(index)
        cell.Value2 = bad: number = 0: description = vbNullString
        TestRequiredGeometryProgress prefix, stats.Report
        On Error Resume Next
        TestReadGeometry geom, bars
        number = Err.Number: description = Err.Description
        Err.Clear: On Error GoTo 0
        If (kind = "start" Or kind = "end") And index = 1 Then
            TestCheck stats, prefix & ".blankAccepted", number = 0
            If number = 0 Then
                current = TestRequiredGeometryRouteSnapshot(shape, "Rebars")
                cell.Value2 = 0#: zeroSnapshot = TestRequiredGeometryRouteSnapshot(shape, "Rebars")
                TestCheck stats, prefix & ".blankEqualsZero", TestSameGeometryRouteSnapshot(current, zeroSnapshot)
            End If
        Else
            TestCheck stats, prefix & ".reject", number <> 0
            TestCheck stats, prefix & ".address", InStr(1, description, cell.Address(False, False), vbTextCompare) > 0
            TestCheck stats, prefix & ".repair", InStr(1, description, "Введите", vbTextCompare) > 0 Or _
                InStr(1, description, "Выберите", vbTextCompare) > 0 Or InStr(1, description, "Исправ", vbTextCompare) > 0
        End If
        stats.Report = stats.Report & "REBAR_FIELD_ERROR: " & prefix & "; error=" & CStr(number) & "; " & description & vbCrLf
        cell.Formula = savedValue
        current = TestRequiredGeometryRouteSnapshot(shape, "Rebars")
        TestCheck stats, prefix & ".recovery", TestSameGeometryRouteSnapshot(baseline, current)
    Next bad
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Находит поле по ключам; номера в контрольных наборах обозначают стороны,
' а не физические строки текущего листа. Размеры General здесь не изменяются.
Private Function TestRebarFieldCell(ByVal target As Object, ByVal rowIndex As Long, ByVal column As Long) As Object
    Dim shape As String, face As Variant, group As String, key As String, block As String, fields As Variant
    shape = CStr(target.Name.Name)
    shape = Mid$(shape, InStrRev(shape, "!") + 1)
    shape = Mid$(shape, 4, Len(shape) - 11)
    For Each face In TestRebarFieldFaces(shape)
        block = ""
        If rowIndex = CLng(face(1)) Then block = "Main"
        If rowIndex = CLng(face(2)) Then block = "Extra"
        If Len(block) > 0 Then
            group = CStr(face(0))
            If shape = "Rectangles" Then
                key = "R" & Mid$(group, 2, 1) & "."
                If Left$(group, 1) = "H" Then
                    key = key & IIf(Right$(group, 1) = "1", "Left", "Right")
                Else
                    key = key & IIf(Right$(group, 1) = "1", "Top", "Bottom")
                End If
            Else
                key = Mid$(group, InStrRev(group, ".") + 1)
                If Left$(group, 8) = "Opening." Then key = "Opening" & key
            End If
            If block = "Main" Then
                fields = Array("key", "name", "as", "d", "n", "s0", "s1", "endMode", "note")
            Else
                fields = Array("key", "name", "d2", "a2", "bind2", "pos2", "d3", "a3", "bind3", "pos3", "note")
            End If
            Set TestRebarFieldCell = TestFixtureInputCell(target, block, key, CStr(fields(column - 1)))
            Exit Function
        End If
    Next face
    Err.Raise 5, , "Неизвестная сторона контрольного набора: " & CStr(rowIndex)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает количество, средний диаметр и средние координаты стержней заданного
' ряда выбранной стороны. Имя row_N нужно только для группировки арматуры,
' а не для определения статусов или разбора ошибки.
Private Function TestRebarFieldRowSummary(ByVal bars As CRebarLayout, ByVal group As String, ByVal rowNumber As Long) As Variant
    Dim count As Long, diameter As Double, X As Double, Y As Double, i As Long
    For i = 1 To bars.Count
        If bars.BarAnnotationGroupName(i) = group Then
            If TestRebarFieldBelongsToRow(bars.BarID(i), group, rowNumber) Then
                count = count + 1: diameter = diameter + bars.Diameter(i)
                X = X + bars.X(i): Y = Y + bars.Y(i)
            End If
        End If
    Next i
    If count > 0 Then diameter = diameter / count: X = X / count: Y = Y / count
    TestRebarFieldRowSummary = Array(count, diameter, X, Y)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' ряд определяется по ID декларативного правила, а не по
' удалённому формату имён legacy-раскладчика. Расчётные R1/R2 здесь не используются.
Private Function TestRebarFieldBelongsToRow(ByVal barID As String, ByVal group As String, ByVal rowNumber As Long) As Boolean
    Dim prefix As String
    prefix = group & ".Row" & CStr(rowNumber) & "."
    If rowNumber = 1 Then prefix = group & ".Main."
    TestRebarFieldBelongsToRow = (Left$(barID, Len(prefix)) = prefix)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Находит для каждого дополнительного стержня ближайший базовый стержень
' своей грани. Просторный fixture исключает неоднозначность: шаг первого ряда
' больше любого смещения. Проверяются расстояние и компонента нормали, включая
' противоположное направление у Opening; production offset helper не вызывается.
Private Sub TestCheckRebarRowOffsets(ByRef stats As TGeometryConfigStats, ByVal prefix As String, _
        ByVal bars As CRebarLayout, ByVal group As String, ByVal rowNumber As Long, ByVal face As Variant, _
        ByVal expectedDistance As Double, ByVal stacked As Boolean)
    Dim i As Long, j As Long, closest As Long, distance As Double, minimum As Double, ordinal As Long
    Dim dx As Double, dy As Double, expectedNormal As Double
    If stacked Then expectedNormal = expectedDistance
    For i = 1 To bars.Count
        If bars.BarAnnotationGroupName(i) = group And _
                TestRebarFieldBelongsToRow(bars.BarID(i), group, rowNumber) Then
            closest = 0
            For j = 1 To bars.Count
                If bars.BarAnnotationGroupName(j) = group And TestRebarFieldBelongsToRow(bars.BarID(j), group, 1) Then
                    distance = Sqr((bars.X(i) - bars.X(j)) ^ 2 + (bars.Y(i) - bars.Y(j)) ^ 2)
                    If closest = 0 Then
                        closest = j: minimum = distance
                    ElseIf distance < minimum Then
                        closest = j: minimum = distance
                    End If
                End If
            Next j
            ordinal = ordinal + 1
            TestCheck stats, prefix & ".base" & CStr(ordinal), closest > 0
            If closest > 0 Then
                dx = bars.X(i) - bars.X(closest): dy = bars.Y(i) - bars.Y(closest)
                TestNear stats, prefix & ".distance" & CStr(ordinal), minimum, expectedDistance
                TestNear stats, prefix & ".normal" & CStr(ordinal), dx * CDbl(face(3)) + dy * CDbl(face(4)), expectedNormal
            End If
        End If
    Next i
    TestCheck stats, prefix & ".offsetsObserved", ordinal > 0
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' короткая проверка адресов и действия полей после перестройки таблиц.
' Проверяет основное и дополнительное армирование, отверстие и невыбранный тип;
' полный перебор всех сторон остаётся в RunAudit03RebarFieldConfigTests.
Public Function RunSectionRebarInputSmokeTests() As String
    Dim stats As TGeometryConfigStats, config As Object, saved As Variant, target As Object
    Dim shape As Variant, face As Variant, field As Variant, cases As Long, group As String
    On Error GoTo Failed
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = TestFixtureSaveInputs(config)
    ThisWorkbook.Names("rngUnitSettings").RefersToRange.Cells(2, 2).Value2 = "mm"
    For Each shape In Array("Rectangles", "PointedSection", "HollowRectangle")
        Set target = ThisWorkbook.Names("rng" & CStr(shape) & "Geometry").RefersToRange
        For Each face In TestRebarFieldFaces(CStr(shape))
            group = CStr(face(0))
            If group = "H1.as_1" Or group = "H.Left" Or group = "Opening.H.Left" Then
                For Each field In TestRebarFaceFields(CStr(shape), face)
                    Select Case CStr(field(0))
                        Case "as", "n", "d2", "bind3"
                            TestConfigureRebarFieldFixture target, CStr(shape), CBool(field(3))
                            TestCheckRebarFieldEffect stats, target, CStr(shape), face, field, cases
                            TestCheckRebarFieldInvalid stats, target, CStr(shape), face, field, cases
                    End Select
                Next field
            End If
        Next face
        modTestShapeArchitecture.TestConfigureShapeFixture "Circle"
        face = TestRebarFieldFaces(CStr(shape))(0)
        field = TestRebarFaceFields(CStr(shape), face)(0)
        TestCheckInactiveRebarCell stats, target, "Circle", face, field, "otherType", cases
    Next shape
    GoTo Finished
Failed:
    TestCheck stats, "sectionInput.smoke.runtime; " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    If Not config Is Nothing And IsArray(saved) Then TestFixtureRestoreInputs config, saved
    RunSectionRebarInputSmokeTests = stats.Report & "TOTAL_SECTION_REBAR_INPUT_SMOKE: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed) & "; cases=" & CStr(cases) & vbCrLf
End Function
