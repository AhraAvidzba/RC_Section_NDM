Attribute VB_Name = "modTestGeometryConfig"
Option Explicit

' ==========================================================================
' ДЛЯ ТЕСТОВ: параметры геометрии через фактический Config
' ==========================================================================
' Проверяет передачу пользовательских геометрических полей от ячеек через
' CSystemSettingsReader/CUnitSystem и CSectionTypeRegistry к готовой форме
' и раскладке. Не ищет равновесие и не подменяет инженерный расчет fixtures.
' Измененные таблицы восстанавливаются по исходным формулам даже при ошибке.

Private Type TGeometryConfigStats
    Passed As Long
    Failed As Long
    Report As String
End Type

' Проверяет представимость количества до размещения и выделения массивов.
' Только два max-Long случая имеют корректный отступ и воспроизводят Overflow;
' остальные намеренно используют неверный отступ, чтобы старый код завершался
' безопасно, без миллиардного цикла или многогигабайтного массива.
Public Function RunAudit03RebarCounterTests(Optional ByRef passed As Long = 0, _
        Optional ByRef failed As Long = 0) As String
    Dim stats As TGeometryConfigStats, shapeName As Variant, scenario As Long
    Dim bars As CRebarLayout, number As Long, description As String, expected As Long
    For Each shapeName In Array("RectSet", "RoundedRectangle", "HollowRectangle")
        Select Case CStr(shapeName)
            Case "RectSet": expected = vbObjectError + 3171
            Case "RoundedRectangle": expected = vbObjectError + 3230
            Case "HollowRectangle": expected = vbObjectError + 3250
        End Select
        For scenario = 0 To 2
            Set bars = Nothing: number = 0: description = vbNullString
            On Error Resume Next
            Set bars = CounterFixture(CStr(shapeName), scenario)
            number = Err.Number: description = Err.Description
            Err.Clear
            On Error GoTo 0
            Check stats, "rebarCounter." & CStr(shapeName) & "." & CStr(scenario) & ".typed", number = expected
            Check stats, "rebarCounter." & CStr(shapeName) & "." & CStr(scenario) & ".message", _
                InStr(1, description, "Config", vbTextCompare) > 0 And InStr(1, description, "счетчиков", vbTextCompare) > 0
            stats.Report = stats.Report & "REBAR_COUNTER_FAILURE: shape=" & CStr(shapeName) & "; scenario=" & CStr(scenario) & _
                "; error=" & CStr(number) & "; comment=" & description & vbCrLf
        Next scenario
        On Error GoTo FailedRun
        Set bars = CounterFixture(CStr(shapeName), 3)
        Check stats, "rebarCounter." & CStr(shapeName) & ".disabled", bars.Count = 0
        Set bars = CounterFixture(CStr(shapeName), 4)
        Check stats, "rebarCounter." & CStr(shapeName) & ".recovery", bars.Count = 3
    Next shapeName
    Set bars = CounterFixture("RectSet", 5)
    Check stats, "rebarCounter.RectSet.inactiveLower", bars.Count = 0
    Near stats, "rebarCounter.binding.odd", RebarRequestedPositionCount(5, 16#, 12#, 10#, "EverySecondBar", "EverySecondBar"), 11#
    Near stats, "rebarCounter.binding.even", RebarRequestedPositionCount(6, 16#, 12#, 10#, "EverySecondBar", "EachBar"), 15#
    Near stats, "rebarCounter.binding.maxLong", RebarRequestedPositionCount(2147483647, 16#, 12#, 0#, "EverySecondBar", "EachBar"), 3221225471#
    Near stats, "rebarCounter.binding.disabled", RebarRequestedPositionCount(2147483647, 0#, 12#, 10#, "EachBar", "EachBar"), 0#
    GoTo Finish
FailedRun:
    stats.Failed = stats.Failed + 1
    stats.Report = stats.Report & "FAIL: rebarCounter.runtime; " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Finish:
    passed = stats.Passed: failed = stats.Failed
    RunAudit03RebarCounterTests = stats.Report & "TOTAL_REBAR_COUNTER: passed=" & CStr(passed) & "; failed=" & CStr(failed) & vbCrLf
End Function

' Создает один и тот же безопасный крайний ввод для трех штатных builders.
' Сценарии: max Long, сумма двух граней, сумма рядов, выключенная грань,
' обычные три стержня и неактивные нижние грани режима Rectangle.
Private Function CounterFixture(ByVal shapeName As String, ByVal scenario As Long) As CRebarLayout
    Dim count1 As Long, count2 As Long, axisDistance As Double, diameter As Double, row2 As Double
    count1 = 2147483647: axisDistance = 0#: diameter = 16#
    Select Case scenario
        Case 0
            If shapeName <> "RectSet" Then axisDistance = 40#
        Case 1: count1 = 1100000000: count2 = 1100000000
        Case 2: count1 = 1100000000: row2 = 16#
        Case 3: diameter = 0#: axisDistance = 40#
        Case 4: count1 = 3: axisDistance = 40#
        Case 5: axisDistance = 40#
    End Select
    Dim inactive As Variant, active As Variant, second As Variant
    If shapeName = "RectSet" Then
        inactive = CounterRectFace(0, 0, 40#, 0#, 0#)
        active = CounterRectFace(count1, count2, axisDistance, diameter, row2)
        Dim rectBuilder As CRectSetRebarLayoutBuilder
        Set rectBuilder = New CRectSetRebarLayoutBuilder
        If scenario = 5 Then
            Set CounterFixture = rectBuilder.Build(600#, 400#, 600#, 400#, 0#, 0#, inactive, active, inactive, active, "Rebar", 0#, "Rectangle")
        Else
            Set CounterFixture = rectBuilder.Build(600#, 400#, 600#, 400#, 0#, 0#, active, inactive, inactive, inactive, "Rebar", 0#, "Rectangle")
        End If
    Else
        inactive = CounterCompactFace(0, 40#, 0#, 0#)
        active = CounterCompactFace(count1, axisDistance, diameter, row2)
        second = CounterCompactFace(count2, axisDistance, diameter, 0#)
        If shapeName = "RoundedRectangle" Then
            Dim rounded As CGeometryRoundedRectangle, roundedBuilder As CRoundedRectRebarLayoutBuilder
            Set rounded = New CGeometryRoundedRectangle: rounded.Initialize 600#, 400#
            Set roundedBuilder = New CRoundedRectRebarLayoutBuilder
            Set CounterFixture = roundedBuilder.Build(rounded, active, second, inactive, inactive, "Rebar")
        Else
            Dim hollow As CGeometryHollowRectangle, hollowBuilder As CHollowRectRebarLayoutBuilder
            Set hollow = New CGeometryHollowRectangle: hollow.Initialize 1000#, 1200#, 60#, 400#, 500#, 30#
            Set hollowBuilder = New CHollowRectRebarLayoutBuilder
            Set CounterFixture = hollowBuilder.Build(hollow, inactive, inactive, active, second, inactive, inactive, inactive, inactive, "Rebar")
        End If
    End If
End Function

' Заполняет компактную грань Rounded/Hollow: первый ряд, дополнительные ряды
' и их положение/привязка. Неактивные диаметры остаются нулевыми.
Private Function CounterCompactFace(ByVal count As Long, ByVal axisDistance As Double, _
        ByVal diameter As Double, ByVal row2 As Double) As Variant
    CounterCompactFace = Array(axisDistance, diameter, count, row2, 0#, "Stacked", "Stacked", "EachBar", "EachBar")
End Function

' Заполняет две линии одной грани RectSet без концевых отступов; второй ряд
' относится только к первой линии. Нижние грани можно передать неактивными.
Private Function CounterRectFace(ByVal count1 As Long, ByVal count2 As Long, ByVal axisDistance As Double, _
        ByVal diameter As Double, ByVal row2 As Double) As Variant
    CounterRectFace = Array(axisDistance, axisDistance, diameter, diameter, count1, count2, _
        0#, 0#, 0#, 0#, row2, 0#, 0#, 0#, "Stacked", "Stacked", "EachBar", "EachBar")
End Function

' Выполняет активные/неактивные варианты восьми параметров круга, границы
' и восстановление после ошибочного ввода. Счетчики доступны полному UI-suite.
Public Function RunAudit03CircleConfigTests(Optional ByRef passed As Long = 0, _
        Optional ByRef failed As Long = 0) As String
    Dim stats As TGeometryConfigStats
    Dim systemRange As Object, circleRange As Object, unitRange As Object
    Dim savedSystem As Variant, savedCircle As Variant, savedUnits As Variant
    On Error GoTo FailedRun
    Set systemRange = ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange
    Set circleRange = ThisWorkbook.Names.Item("rngCircleGeometry").RefersToRange
    Set unitRange = ThisWorkbook.Names.Item("rngUnitSettings").RefersToRange
    savedSystem = systemRange.Formula: savedCircle = circleRange.Formula: savedUnits = unitRange.Formula
    SetKey systemRange, "Geometry.Type", "Circle"
    unitRange.Cells(2, 2).Value2 = "mm"
    ConfigureCircle circleRange
    CheckCircleEffects stats, circleRange
    CheckCircleBoundaries stats, circleRange
    CheckCircleInactive stats, systemRange, circleRange
    GoTo Restore
FailedRun:
    stats.Failed = stats.Failed + 1
    stats.Report = stats.Report & "FAIL: circleConfig.runtime; " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Restore:
    On Error GoTo FailedRestore
    If Not systemRange Is Nothing Then systemRange.Formula = savedSystem
    If Not circleRange Is Nothing Then circleRange.Formula = savedCircle
    If Not unitRange Is Nothing Then unitRange.Formula = savedUnits
    If Not systemRange Is Nothing Then Check stats, "circleConfig.restore.system", SameFormula(systemRange.Formula, savedSystem)
    If Not circleRange Is Nothing Then Check stats, "circleConfig.restore.circle", SameFormula(circleRange.Formula, savedCircle)
    If Not unitRange Is Nothing Then Check stats, "circleConfig.restore.units", SameFormula(unitRange.Formula, savedUnits)
    GoTo Finish
FailedRestore:
    stats.Failed = stats.Failed + 1
    stats.Report = stats.Report & "FAIL: circleConfig.restore.runtime; " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Finish:
    passed = stats.Passed: failed = stats.Failed
    RunAudit03CircleConfigTests = stats.Report & "TOTAL_CIRCLE_CONFIG: passed=" & CStr(passed) & "; failed=" & CStr(failed) & vbCrLf
End Function

' Проверяет размеры и режимы трех некруглых форм через реальные ячейки,
' включая переключение единиц, неактивные поля и неверные количества арматуры.
' Все измененные таблицы восстанавливаются, чтобы тест не менял рабочий ввод.
Public Function RunAudit03ShapeConfigTests(Optional ByRef passed As Long = 0, _
        Optional ByRef failed As Long = 0) As String
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
    For Each rangeName In Array("rngRectSetGeometry", "rngRoundedRectangleGeometry", "rngHollowRectangleGeometry")
        Set shapeRanges(index) = ThisWorkbook.Names.Item(CStr(rangeName)).RefersToRange
        savedShapes(index) = shapeRanges(index).Formula
        index = index + 1
    Next rangeName
    unitRange.Cells(2, 2).Value2 = "mm"
    ConfigureCircle circleRange
    shapeNames = Array("RectSet", "RoundedRectangle", "HollowRectangle")
    For index = 0 To 2
        Set target = shapeRanges(index)
        shapeName = CStr(shapeNames(index))
        SetKey systemRange, "Geometry.Type", shapeName
        CheckShapeEffects stats, target, shapeName
        CheckShapeInactiveModes stats, target, shapeName
        CheckShapeUnits stats, target, unitRange, shapeName
        SetKey systemRange, "Geometry.Type", "Circle"
        CheckShapeInactive stats, target, shapeName
    Next index
    SetKey systemRange, "Geometry.Type", "HollowRectangle"
    CheckHollowCounts stats, shapeRanges(2)
    GoTo Restore
FailedRun:
    stats.Failed = stats.Failed + 1
    stats.Report = stats.Report & "FAIL: shapeConfig.runtime; " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Restore:
    On Error GoTo FailedRestore
    For index = 0 To 2
        If Not shapeRanges(index) Is Nothing Then
            shapeRanges(index).Formula = savedShapes(index)
            Check stats, "shapeConfig.restore." & CStr(index), SameFormula(shapeRanges(index).Formula, savedShapes(index))
        End If
    Next index
    If Not systemRange Is Nothing Then systemRange.Formula = savedSystem
    If Not unitRange Is Nothing Then unitRange.Formula = savedUnits
    If Not circleRange Is Nothing Then circleRange.Formula = savedCircle
    If Not systemRange Is Nothing Then Check stats, "shapeConfig.restore.system", SameFormula(systemRange.Formula, savedSystem)
    If Not unitRange Is Nothing Then Check stats, "shapeConfig.restore.units", SameFormula(unitRange.Formula, savedUnits)
    If Not circleRange Is Nothing Then Check stats, "shapeConfig.restore.circle", SameFormula(circleRange.Formula, savedCircle)
    GoTo Finish
FailedRestore:
    stats.Failed = stats.Failed + 1
    stats.Report = stats.Report & "FAIL: shapeConfig.restore.runtime; " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Finish:
    passed = stats.Passed: failed = stats.Failed
    RunAudit03ShapeConfigTests = stats.Report & "TOTAL_SHAPE_CONFIG: passed=" & CStr(passed) & "; failed=" & CStr(failed) & vbCrLf
End Function

' Описывает фактические позиции размеров в трех таблицах Config, а не поиск
' строк с искусственными ключами. Поля: ключ, строка, колонка, база, изменение,
' недопустимое значение. Эти позиции проверены по действующему reader-у.
Private Function ShapeFields(ByVal shapeName As String) As Variant
    Select Case shapeName
        Case "RectSet"
            ShapeFields = Array(Array("RectSet.H1", 8, 1, 300#, 350#, -1#), _
                Array("RectSet.B1", 8, 2, 200#, 240#, -1#), _
                Array("RectSet.H2", 8, 3, 250#, 300#, -1#), _
                Array("RectSet.B2", 8, 4, 600#, 650#, -1#), _
                Array("RectSet.UpperOffsetX", 4, 2, -50#, -70#, 10000#), _
                Array("RectSet.SectionType", 3, 2, "TwoRectangles", "Rectangle", "Unknown"))
        Case "RoundedRectangle"
            ShapeFields = Array(Array("RoundedRectangle.B", 3, 2, 600#, 650#, -1#), _
                Array("RoundedRectangle.H", 3, 3, 400#, 450#, -1#), _
                Array("RoundedRectangle.Left.Type", 7, 2, "Tapered", "Simple", "Unknown"), _
                Array("RoundedRectangle.Right.Type", 7, 3, "Tapered", "Simple", "Unknown"), _
                Array("RoundedRectangle.Left.W", 8, 2, 100#, 130#, -1#), _
                Array("RoundedRectangle.Right.W", 8, 3, 100#, 130#, -1#), _
                Array("RoundedRectangle.Left.R1", 9, 2, 30#, 45#, -1#), _
                Array("RoundedRectangle.Right.R1", 9, 3, 30#, 45#, -1#), _
                Array("RoundedRectangle.Left.R2", 10, 2, 20#, 35#, -1#), _
                Array("RoundedRectangle.Right.R2", 10, 3, 20#, 35#, -1#))
        Case "HollowRectangle"
            ShapeFields = Array(Array("HollowRectangle.H", 7, 1, 1200#, 1300#, -1#), _
                Array("HollowRectangle.B", 7, 2, 1000#, 1100#, -1#), _
                Array("HollowRectangle.R", 7, 3, 60#, 80#, -1#), _
                Array("HollowRectangle.OpeningH", 7, 4, 500#, 550#, -1#), _
                Array("HollowRectangle.OpeningB", 7, 5, 400#, 450#, -1#), _
                Array("HollowRectangle.OpeningR", 7, 6, 30#, 40#, -1#), _
                Array("HollowRectangle.InnerOffsetX", 3, 2, 40#, 60#, 10000#), _
                Array("HollowRectangle.InnerOffsetY", 4, 2, 50#, 70#, 10000#))
        Case Else
            Err.Raise vbObjectError + 5961, "modTestGeometryConfig", "Неизвестная форма fixture: " & shapeName
    End Select
End Function

' Задает физически допустимую форму. Множитель переводит только длины,
' поэтому mm и cm fixtures описывают одно сечение, а enum не изменяются.
Private Sub ConfigureShape(ByVal target As Object, ByVal shapeName As String, Optional ByVal lengthScale As Double = 1#)
    Dim item As Variant, value As Variant
    For Each item In ShapeFields(shapeName)
        value = item(3)
        If IsNumeric(value) Then value = CDbl(value) * lengthScale
        target.Cells(CLng(item(1)), CLng(item(2))).Value2 = value
    Next item
End Sub

' ==========================================================================
' ДЛЯ ТЕСТОВ: ОБЯЗАТЕЛЬНЫЙ И УСЛОВНО НЕАКТИВНЫЙ ВВОД АРМАТУРЫ КРУГА
' ==========================================================================
' Проверяет реальные ячейки круга, отсутствие обязательных строк и ошибки
' формул. Отдельно проверяет отключенные ряды и невыбранную форму: их значения
' не участвуют в раскладке. Весь исходный ввод восстанавливается по Formula.
Public Function RunAudit03CircleRebarInputTests(Optional ByRef passed As Long = 0, _
        Optional ByRef failed As Long = 0) As String
    Dim stats As TGeometryConfigStats, ranges(0 To 3) As Object, saved(0 To 3) As Variant
    Dim names As Variant, keys As Variant, key As Variant, bad As Variant, badValues As Variant
    Dim i As Long, cases As Long, cell As Object, savedKey As Variant
    Dim baseline As Variant, prefix As String, offKey As Variant, offValue As Variant, ignored As Variant
    names = Array("rngSystemSettings", "rngUnitSettings", "rngCircleGeometry", "rngRectSetGeometry")
    keys = Array("Circle.Diameter", "Rebar.AxisDistance", "Rebar.Count", "Rebar.Diameter", _
        "Rebar.Diameter2", "Rebar.Diameter3", "Rebar.Loc2row", "Rebar.Loc3row")
    On Error GoTo FailedRun
    For i = 0 To 3
        Set ranges(i) = ThisWorkbook.Names.Item(CStr(names(i))).RefersToRange
        saved(i) = ranges(i).Formula
    Next i
    SetKey ranges(0), "Geometry.Source", "Generated"
    SetKey ranges(0), "Geometry.Type", "Circle"
    ranges(1).Cells(2, 2).Value2 = "mm"
    ConfigureCircle ranges(2)
    baseline = RequiredGeometryRouteSnapshot("Circle", "Rebars")
    For Each key In keys
        ConfigureCircle ranges(2)
        If InStr(1, CStr(key), "Loc", vbBinaryCompare) > 0 Then
            badValues = Array(vbNullString, "Wrong", CVErr(2015))
        ElseIf CStr(key) = "Rebar.Count" Then
            badValues = Array(-1#, 1#, 0.5, 2147483648#, "abc", CVErr(2015))
        ElseIf CStr(key) = "Circle.Diameter" Then
            badValues = Array(vbNullString, -1#, 0#, "abc", CVErr(2015))
        ElseIf CStr(key) = "Rebar.AxisDistance" Then
            badValues = Array(vbNullString, -1#, 0#, 500#, "abc", CVErr(2015))
        Else
            badValues = Array(-1#, "abc", CVErr(2015))
        End If
        Set cell = CircleInputCell(ranges(2), CStr(key))
        i = 0
        For Each bad In badValues
            i = i + 1
            CheckCircleInputAttempt stats, cell, bad, "circleRebar.active." & CStr(key) & "." & CStr(i), _
                "Circle", baseline, True, cases
        Next bad

        ' Пустой диаметр является допустимым выключением; исчезнувшая строка
        ' имеет другой смысл и не должна незаметно брать шаблонный диаметр.
        savedKey = cell.Offset(0, -1).Formula
        cell.Offset(0, -1).Value2 = "Audit03.AbsentCircleField"
        CheckCircleMissingInput stats, CStr(key), cases
        cell.Offset(0, -1).Formula = savedKey
        Check stats, "circleRebar.missing." & CStr(key) & ".recovery", _
            SameGeometryRouteSnapshot(baseline, RequiredGeometryRouteSnapshot("Circle", "Rebars"))
    Next key

    For Each offKey In Array("Rebar.Count", "Rebar.Diameter", "Rebar.Diameter2", "Rebar.Diameter3")
        For Each offValue In Array(vbNullString, 0#)
            ConfigureCircle ranges(2): SetKey ranges(2), CStr(offKey), offValue
            baseline = RequiredGeometryRouteSnapshot("Circle", "Rebars")
            If CStr(offKey) = "Rebar.Count" Or CStr(offKey) = "Rebar.Diameter" Then
                ignored = Array("Rebar.AxisDistance", "Rebar.Diameter2", "Rebar.Diameter3", "Rebar.Loc2row", "Rebar.Loc3row")
            ElseIf CStr(offKey) = "Rebar.Diameter2" Then
                ignored = Array("Rebar.Loc2row")
            Else
                ignored = Array("Rebar.Loc3row")
            End If
            For Each key In ignored
                Set cell = CircleInputCell(ranges(2), CStr(key)): i = 0
                For Each bad In Array(vbNullString, "abc", CVErr(2015), -1#)
                    i = i + 1
                    prefix = "circleRebar.inactive." & CStr(offKey) & "." & CStr(offValue) & "." & CStr(key) & "." & CStr(i)
                    CheckCircleInputAttempt stats, cell, bad, prefix, "Circle", baseline, False, cases
                Next bad
            Next key
        Next offValue
    Next offKey

    ConfigureRebarFieldFixture ranges(3), "RectSet", True
    SetKey ranges(0), "Geometry.Type", "RectSet"
    ConfigureCircle ranges(2): baseline = RequiredGeometryRouteSnapshot("RectSet", "Rebars")
    For Each key In keys
        Set cell = CircleInputCell(ranges(2), CStr(key)): i = 0
        For Each bad In Array(vbNullString, "abc", CVErr(2015))
            i = i + 1
            CheckCircleInputAttempt stats, cell, bad, "circleRebar.otherShape." & CStr(key) & "." & CStr(i), _
                "RectSet", baseline, False, cases
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
            Check stats, "circleRebar.restore." & CStr(names(i)), SameFormula(ranges(i).Formula, saved(i))
        End If
    Next i
    GoTo Finish
FailedRestore:
    stats.Failed = stats.Failed + 1
    stats.Report = stats.Report & "FAIL: circleRebar.restore.runtime; " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Finish:
    passed = stats.Passed: failed = stats.Failed
    RunAudit03CircleRebarInputTests = stats.Report & "TOTAL_CIRCLE_REBAR_INPUT: passed=" & CStr(passed) & _
        "; failed=" & CStr(failed) & "; cases=" & CStr(cases) & vbCrLf
End Function

' Находит фактическую ячейку значения по ключу, без типовых адресов Config.
' Это позволяет проверить адресную диагностику после перемещения таблицы.
Private Function CircleInputCell(ByVal target As Object, ByVal key As String) As Object
    Dim i As Long
    For i = 2 To target.Rows.Count
        If CStr(target.Cells(i, 1).Value2) = key Then
            Set CircleInputCell = target.Cells(i, 2)
            Exit Function
        End If
    Next i
    Err.Raise vbObjectError + 5960, "modTestGeometryConfig", "Не найдена строка fixture: " & key
End Function

' Проверяет один ввод через полный production-маршрут. Ошибка активного поля
' должна иметь адрес и инструкцию исправления; неактивное поле обязано оставить
' весь снимок неизменным. Затем проверяется восстановление исходной формулы.
Private Sub CheckCircleInputAttempt(ByRef stats As TGeometryConfigStats, ByVal cell As Object, _
        ByVal value As Variant, ByVal prefix As String, ByVal shape As String, ByVal baseline As Variant, _
        ByVal mustReject As Boolean, ByRef cases As Long)
    Dim savedValue As Variant, number As Long, description As String, actual As Variant
    savedValue = cell.Formula: cell.Value2 = value: cases = cases + 1
    RequiredGeometryProgress prefix, stats.Report
    On Error Resume Next
    actual = RequiredGeometryRouteSnapshot(shape, "Rebars")
    number = Err.Number: description = Err.Description
    Err.Clear: On Error GoTo 0
    If mustReject Then
        Check stats, prefix & ".reject", number <> 0
        Check stats, prefix & ".address", InStr(1, description, cell.Address(False, False), vbTextCompare) > 0
        Check stats, prefix & ".repair", InStr(1, description, "Введите", vbTextCompare) > 0 Or _
            InStr(1, description, "Выберите", vbTextCompare) > 0 Or InStr(1, description, "Исправ", vbTextCompare) > 0
    Else
        Check stats, prefix & ".noInterference", number = 0
        If number = 0 Then Check stats, prefix & ".snapshot", SameGeometryRouteSnapshot(baseline, actual)
    End If
    stats.Report = stats.Report & "CIRCLE_REBAR_INPUT: " & prefix & "; error=" & CStr(number) & "; " & description & vbCrLf
    cell.Formula = savedValue
    Check stats, prefix & ".recovery", SameGeometryRouteSnapshot(baseline, RequiredGeometryRouteSnapshot(shape, "Rebars"))
End Sub

' Проверяет структурно отсутствующий активный параметр отдельно от пустой
' ячейки. Сообщение должно назвать ключ и способ восстановить ввод, а не default.
Private Sub CheckCircleMissingInput(ByRef stats As TGeometryConfigStats, ByVal key As String, ByRef cases As Long)
    Dim actual As Variant, number As Long, description As String
    cases = cases + 1: RequiredGeometryProgress "circleRebar.missing." & key, stats.Report
    On Error Resume Next
    actual = RequiredGeometryRouteSnapshot("Circle", "Rebars")
    number = Err.Number: description = Err.Description
    Err.Clear: On Error GoTo 0
    Check stats, "circleRebar.missing." & key & ".reject", number <> 0
    Check stats, "circleRebar.missing." & key & ".key", InStr(1, description, key, vbBinaryCompare) > 0
    Check stats, "circleRebar.missing." & key & ".repair", InStr(1, description, "Восстанов", vbTextCompare) > 0 Or _
        InStr(1, description, "Введите", vbTextCompare) > 0 Or InStr(1, description, "Выберите", vbTextCompare) > 0
    stats.Report = stats.Report & "CIRCLE_MISSING_INPUT: " & key & "; error=" & CStr(number) & "; " & description & vbCrLf
End Sub

' Читает форму штатным reader/unit/registry pipeline без создания арматуры:
' проверка одного размера не должна зависеть от чужой раскладки стержней.
Private Function ReadShape() As CGeometryRegion
    Dim settings As CSystemSettingsReader, units As CUnitSystem, registry As CSectionTypeRegistry
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    Set registry = New CSectionTypeRegistry
    Set ReadShape = registry.CreateGeometry(settings, units)
End Function

' Сохраняет расчетные площадь, центр и габариты. Отказ геометрии считается
' ошибкой fixture; снимок не подменяет форму ее исходными getter-параметрами.
Private Function ShapeSnapshot(ByVal geom As CGeometryRegion) As Variant
    Dim message As String, available As Boolean, area As Double, X As Double, Y As Double
    If Not geom.IsValid(message) Then Err.Raise vbObjectError + 5962, "modTestGeometryConfig", message
    area = geom.AnalyticalArea(available)
    If Not available Then Err.Raise vbObjectError + 5963, "modTestGeometryConfig", "Нет аналитической площади fixture."
    geom.AnalyticalCentroid available, X, Y
    If Not available Then Err.Raise vbObjectError + 5964, "modTestGeometryConfig", "Нет центра тяжести fixture."
    ShapeSnapshot = Array(area, X, Y, geom.MinX, geom.MaxX, geom.MinY, geom.MaxY)
End Function

' Сравнивает все численные характеристики в фиксированном абсолютном допуске.
' Локальный допуск теста не меняет точность mesh/solver пользовательского расчета.
Private Function SameShape(ByVal first As Variant, ByVal second As Variant) As Boolean
    Dim index As Long
    For index = 0 To 6
        If Abs(CDbl(first(index)) - CDbl(second(index))) > 0.000001 Then Exit Function
    Next index
    SameShape = True
End Function

' Для каждого поля проверяет изменение готовой формы, отклонение неверного
' значения и повторный нормальный ввод. Числа журналируются для последующей сверки.
Private Sub CheckShapeEffects(ByRef stats As TGeometryConfigStats, ByVal target As Object, ByVal shapeName As String)
    Dim item As Variant, baseline As Variant, changed As Variant, geom As CGeometryRegion
    Dim errorNumber As Long, description As String, valid As Boolean, prefix As String
    ConfigureShape target, shapeName
    baseline = ShapeSnapshot(ReadShape())
    CheckShapeBaseline stats, shapeName, baseline
    For Each item In ShapeFields(shapeName)
        prefix = "shapeConfig." & CStr(item(0))
        ConfigureShape target, shapeName
        target.Cells(CLng(item(1)), CLng(item(2))).Value2 = item(4)
        changed = ShapeSnapshot(ReadShape())
        Check stats, prefix & ".effect", Not SameShape(baseline, changed)
        stats.Report = stats.Report & "SHAPE_CONFIG_EFFECT: key=" & CStr(item(0)) & "; area=" & CStr(changed(0)) & _
            "; cx=" & CStr(changed(1)) & "; cy=" & CStr(changed(2)) & vbCrLf
        target.Cells(CLng(item(1)), CLng(item(2))).Value2 = item(5)
        errorNumber = 0: description = vbNullString: valid = False
        On Error Resume Next
        Set geom = ReadShape()
        errorNumber = Err.Number: description = Err.Description
        If errorNumber = 0 Then valid = geom.IsValid(description)
        Err.Clear
        On Error GoTo 0
        Check stats, prefix & ".invalid", Not valid And Len(description) > 0
        ConfigureShape target, shapeName
        Check stats, prefix & ".recovery", SameShape(baseline, ShapeSnapshot(ReadShape()))
    Next item
End Sub

' Независимо проверяет площадь и центр. Скругления представлены точными
' дугами: потеря площади четырех углов равна (4 - pi) * R^2.
Private Sub CheckShapeBaseline(ByRef stats As TGeometryConfigStats, ByVal shapeName As String, ByVal actual As Variant)
    Dim area As Double, openingArea As Double, cornerLoss As Double
    If shapeName = "RectSet" Then
        area = 200# * 300# + 600# * 250#
        Near stats, "shapeConfig.RectSet.oracle.area", CDbl(actual(0)), area
        Near stats, "shapeConfig.RectSet.oracle.cx", CDbl(actual(1)), (600# * 250# * 300# + 200# * 300# * 50#) / area
        Near stats, "shapeConfig.RectSet.oracle.cy", CDbl(actual(2)), (600# * 250# * 125# + 200# * 300# * 400#) / area
    ElseIf shapeName = "HollowRectangle" Then
        cornerLoss = 4# - GEOM_PI
        openingArea = 400# * 500# - cornerLoss * 30# ^ 2
        area = 1000# * 1200# - cornerLoss * 60# ^ 2 - openingArea
        Near stats, "shapeConfig.HollowRectangle.oracle.area", CDbl(actual(0)), area
        Near stats, "shapeConfig.HollowRectangle.oracle.cx", CDbl(actual(1)), -openingArea * 40# / area
        Near stats, "shapeConfig.HollowRectangle.oracle.cy", CDbl(actual(2)), -openingArea * 50# / area
    Else
        Near stats, "shapeConfig.RoundedRectangle.oracle.cx", CDbl(actual(1)), 0#
        Near stats, "shapeConfig.RoundedRectangle.oracle.cy", CDbl(actual(2)), 0#
    End If
End Sub

' Проверяет именно условно выключенные поля формы: нижний прямоугольник и
' смещение, а также W/R2 простых боковых граней не меняют активную геометрию.
Private Sub CheckShapeInactiveModes(ByRef stats As TGeometryConfigStats, ByVal target As Object, ByVal shapeName As String)
    Dim baseline As Variant, item As Variant
    ConfigureShape target, shapeName
    If shapeName = "RectSet" Then
        target.Cells(3, 2).Value2 = "Rectangle"
        baseline = ShapeSnapshot(ReadShape())
        For Each item In Array(Array(8, 3), Array(8, 4), Array(4, 2))
            target.Cells(CLng(item(0)), CLng(item(1))).Value2 = 100#
            Check stats, "shapeConfig.RectSet.Rectangle.inactive." & CStr(item(0)) & "." & CStr(item(1)), _
                SameShape(baseline, ShapeSnapshot(ReadShape()))
        Next item
        ConfigureShape target, shapeName
        target.Cells(3, 2).Value2 = "LSection"
        baseline = ShapeSnapshot(ReadShape())
        target.Cells(4, 2).Value2 = 100#
        Check stats, "shapeConfig.RectSet.LSection.inactive.offset", SameShape(baseline, ShapeSnapshot(ReadShape()))
    ElseIf shapeName = "RoundedRectangle" Then
        target.Cells(7, 2).Value2 = "Simple": target.Cells(7, 3).Value2 = "Simple"
        baseline = ShapeSnapshot(ReadShape())
        For Each item In Array(Array(8, 2), Array(8, 3), Array(10, 2), Array(10, 3))
            target.Cells(CLng(item(0)), CLng(item(1))).Value2 = 70#
            Check stats, "shapeConfig.RoundedRectangle.Simple.inactive." & CStr(item(0)) & "." & CStr(item(1)), _
                SameShape(baseline, ShapeSnapshot(ReadShape()))
        Next item
    End If
End Sub

' Смена INPUT mm на cm с явным переводом всех длин обязана сохранять
' внутреннюю форму. Это не меняет контракт импорта AutoCAD и его миллиметры.
Private Sub CheckShapeUnits(ByRef stats As TGeometryConfigStats, ByVal target As Object, _
        ByVal unitRange As Object, ByVal shapeName As String)
    Dim baseline As Variant
    unitRange.Cells(2, 2).Value2 = "mm"
    ConfigureShape target, shapeName
    baseline = ShapeSnapshot(ReadShape())
    unitRange.Cells(2, 2).Value2 = "cm"
    ConfigureShape target, shapeName, 0.1
    Check stats, "shapeConfig." & shapeName & ".units.cm", SameShape(baseline, ShapeSnapshot(ReadShape()))
    unitRange.Cells(2, 2).Value2 = "mm"
    ConfigureShape target, shapeName
End Sub

' В другом Geometry.Type корректные изменения невыбранной таблицы не должны
' влиять на расчетную форму Circle. Неверные значения здесь не маскируются.
Private Sub CheckShapeInactive(ByRef stats As TGeometryConfigStats, ByVal target As Object, ByVal shapeName As String)
    Dim item As Variant, baseline As Variant
    ConfigureShape target, shapeName
    baseline = ShapeSnapshot(ReadShape())
    For Each item In ShapeFields(shapeName)
        ConfigureShape target, shapeName
        target.Cells(CLng(item(1)), CLng(item(2))).Value2 = item(4)
        Check stats, "shapeConfig." & CStr(item(0)) & ".inactive.type", SameShape(baseline, ShapeSnapshot(ReadShape()))
    Next item
End Sub

' Проверяет все четыре пользовательских счетчика наружной арматуры. Текст,
' дробь, отрицательное/непредставимое число и ошибка Excel не выключают грань
' молча; сообщение должно назвать Config, фактическую ячейку и настройку.
' Служебные количества внутренних граней не являются пользовательским вводом.
Private Sub CheckHollowCounts(ByRef stats As TGeometryConfigStats, ByVal target As Object)
    Dim rowIndex As Long, index As Long, bad As Variant, geom As CGeometryRegion, bars As CRebarLayout
    Dim errorNumber As Long, description As String, face As String, prefix As String, faceNames As Variant
    faceNames = Array("H.Left", "H.Right", "B.Top", "B.Bottom")
    ConfigureShape target, "HollowRectangle"
    For rowIndex = 11 To 18
        target.Cells(rowIndex, 2).Value2 = 40#: target.Cells(rowIndex, 3).Value2 = 16#
        If rowIndex <= 14 Then target.Cells(rowIndex, 4).Value2 = 3 Else target.Cells(rowIndex, 4).Value2 = "auto"
    Next rowIndex
    For rowIndex = 21 To 28
        target.Cells(rowIndex, 2).Value2 = 0#: target.Cells(rowIndex, 5).Value2 = 0#
        target.Cells(rowIndex, 3).Value2 = "Stacked": target.Cells(rowIndex, 6).Value2 = "Stacked"
        target.Cells(rowIndex, 4).Value2 = "EachBar": target.Cells(rowIndex, 7).Value2 = "EachBar"
    Next rowIndex
    For rowIndex = 11 To 14
        face = CStr(faceNames(rowIndex - 11))
        target.Cells(rowIndex, 4).Value2 = 4
        ReadGeometry geom, bars
        Check stats, "shapeConfig.hollowCount." & face & ".effect", FaceBarCount(bars, face) = 4
        target.Cells(rowIndex, 4).Value2 = 3
        For Each bad In Array("abc", -1#, 0.5, CVErr(2015), 2147483648#)
            index = index + 1
            target.Cells(rowIndex, 4).Value2 = bad
            errorNumber = 0: description = vbNullString
            On Error Resume Next
            ReadGeometry geom, bars
            errorNumber = Err.Number: description = Err.Description
            Err.Clear
            On Error GoTo 0
            prefix = "shapeConfig.hollowCount." & face & "." & CStr(index)
            Check stats, prefix & ".reject", errorNumber <> 0
            Check stats, prefix & ".location", InStr(1, description, "Config", vbTextCompare) > 0 And _
                InStr(1, description, target.Cells(rowIndex, 4).Address(False, False), vbTextCompare) > 0
            Check stats, prefix & ".key", InStr(1, description, "HollowRectangle." & face & ".n", vbBinaryCompare) > 0
            stats.Report = stats.Report & "HOLLOW_COUNT_ERROR: " & prefix & "; error=" & CStr(errorNumber) & "; " & description & vbCrLf
            target.Cells(rowIndex, 4).Value2 = 3
            ReadGeometry geom, bars
            Check stats, prefix & ".recovery", FaceBarCount(bars, face) = 3
        Next bad
        target.Cells(rowIndex, 4).Value2 = vbNullString
        ReadGeometry geom, bars
        Check stats, "shapeConfig.hollowCount." & face & ".blank.off", FaceBarCount(bars, face) = 0
        target.Cells(rowIndex, 4).Value2 = 0
        ReadGeometry geom, bars
        Check stats, "shapeConfig.hollowCount." & face & ".zero.off", FaceBarCount(bars, face) = 0
        target.Cells(rowIndex, 4).Value2 = 3
    Next rowIndex
End Sub

' Считает только стержни нужной смысловой грани, не включая ее проекции
' на отверстие или соседние грани; общий ненулевой layout не доказывает эффект n.
Private Function FaceBarCount(ByVal bars As CRebarLayout, ByVal face As String) As Long
    Dim index As Long
    For index = 1 To bars.Count
        If StrComp(bars.BarAnnotationGroupName(index), face, vbBinaryCompare) = 0 Then FaceBarCount = FaceBarCount + 1
    Next index
End Function

' Задает достаточное защитное расстояние и три активных ряда. Каждая следующая
' проверка меняет только одно поле относительно этого физически допустимого набора.
Private Sub ConfigureCircle(ByVal target As Object)
    SetKey target, "Circle.Diameter", 1000#
    SetKey target, "Rebar.AxisDistance", 50#
    SetKey target, "Rebar.Count", 8
    SetKey target, "Rebar.Diameter", 16#
    SetKey target, "Rebar.Diameter2", 12#
    SetKey target, "Rebar.Diameter3", 10#
    SetKey target, "Rebar.Loc2row", "Stacked"
    SetKey target, "Rebar.Loc3row", "Stacked"
End Sub

' Oracle первой позиции не использует builder: ось основного ряда лежит на +X,
' Stacked сдвигает внутрь на полусумму диаметров, третий ряд перескакивает второй.
Private Sub CheckCircleEffects(ByRef stats As TGeometryConfigStats, ByVal target As Object)
    Dim key As Variant, geom As CGeometryRegion, bars As CRebarLayout
    For Each key In Array("Circle.Diameter", "Rebar.AxisDistance", "Rebar.Count", "Rebar.Diameter", _
            "Rebar.Diameter2", "Rebar.Diameter3", "Rebar.Loc2row", "Rebar.Loc3row")
        ConfigureCircle target
        Select Case CStr(key)
            Case "Circle.Diameter": SetKey target, CStr(key), 1100#
            Case "Rebar.AxisDistance": SetKey target, CStr(key), 60#
            Case "Rebar.Count": SetKey target, CStr(key), 10
            Case "Rebar.Diameter": SetKey target, CStr(key), 20#
            Case "Rebar.Diameter2": SetKey target, CStr(key), 20#
            Case "Rebar.Diameter3": SetKey target, CStr(key), 14#
            Case Else: SetKey target, CStr(key), "SideBySide"
        End Select
        ReadGeometry geom, bars
        Dim prefix As String
        prefix = "circleConfig.effect." & CStr(key)
        Select Case CStr(key)
            Case "Circle.Diameter"
                Near stats, prefix & ".concreteExtent", geom.MaxX, 550#
                Near stats, prefix & ".barAxis", bars.X(1), 500#
            Case "Rebar.AxisDistance"
                Near stats, prefix & ".barAxis", bars.X(1), 440#
                Near stats, prefix & ".row2Axis", bars.X(2), 426#
            Case "Rebar.Count"
                Check stats, prefix & ".allRows", bars.Count = 30
                Near stats, prefix & ".secondAngleX", bars.X(4), 450# * Cos(2# * GEOM_PI / 10#)
            Case "Rebar.Diameter"
                Near stats, prefix & ".diameter", bars.Diameter(1), 20#
                Near stats, prefix & ".row2Axis", bars.X(2), 434#
            Case "Rebar.Diameter2"
                Near stats, prefix & ".diameter", bars.Diameter(2), 20#
                Near stats, prefix & ".row3Skip", bars.X(3), 417#
            Case "Rebar.Diameter3"
                Near stats, prefix & ".diameter", bars.Diameter(3), 14#
                Near stats, prefix & ".row3Axis", bars.X(3), 423#
            Case "Rebar.Loc2row"
                Near stats, prefix & ".tangentX", bars.X(2), 450#
                Near stats, prefix & ".clockwiseY", bars.Y(2), -14#
                Near stats, prefix & ".row3NoSkip", bars.X(3), 437#
            Case "Rebar.Loc3row"
                Near stats, prefix & ".tangentX", bars.X(3), 450#
                Near stats, prefix & ".clockwiseY", bars.Y(3), -13#
                Near stats, prefix & ".row2Unchanged", bars.X(2), 436#
        End Select
        stats.Report = stats.Report & "CIRCLE_CONFIG_EFFECT: key=" & CStr(key) & "; bars=" & CStr(bars.Count) & vbCrLf
    Next key
End Sub

' Проверяет ошибочные численные/enum значения на том же reader/registry-маршруте,
' затем обычную раскладку. Пустые/нулевые диаметры сохраняют утвержденное выключение ряда.
Private Sub CheckCircleBoundaries(ByRef stats As TGeometryConfigStats, ByVal target As Object)
    Dim key As Variant, value As Variant, errorNumber As Long, description As String
    Dim geom As CGeometryRegion, bars As CRebarLayout
    For Each key In Array("Circle.Diameter", "Rebar.AxisDistance", "Rebar.Count", "Rebar.Diameter", _
            "Rebar.Diameter2", "Rebar.Diameter3", "Rebar.Loc2row", "Rebar.Loc3row")
        ConfigureCircle target
        If InStr(1, CStr(key), "Loc", vbBinaryCompare) > 0 Then value = "Unknown" Else value = -1
        SetKey target, CStr(key), value
        errorNumber = 0: description = vbNullString
        On Error Resume Next
        ReadGeometry geom, bars
        errorNumber = Err.Number: description = Err.Description
        Err.Clear
        On Error GoTo 0
        Check stats, "circleConfig.invalid." & CStr(key), errorNumber <> 0 And Len(description) > 0
        ConfigureCircle target
        ReadGeometry geom, bars
        Check stats, "circleConfig.recovery." & CStr(key), bars.Count = 24
    Next key
    For Each key In Array("Rebar.Diameter", "Rebar.Diameter2", "Rebar.Diameter3")
        For Each value In Array(vbNullString, 0#)
            ConfigureCircle target
            SetKey target, CStr(key), value
            ReadGeometry geom, bars
            If CStr(key) = "Rebar.Diameter" Then
                Check stats, "circleConfig.optional." & CStr(key) & "." & CStr(value), bars.Count = 0
            Else
                Check stats, "circleConfig.optional." & CStr(key) & "." & CStr(value), bars.Count = 16
            End If
        Next value
    Next key
    ConfigureCircle target
    SetKey target, "Rebar.Count", 0
    ReadGeometry geom, bars
    Check stats, "circleConfig.count.zero", bars.Count = 0
    SetKey target, "Rebar.Count", 2
    ReadGeometry geom, bars
    Check stats, "circleConfig.count.minimum", bars.Count = 6
End Sub

' При другом Geometry.Type изменение каждого поля круга не должно менять
' форму и арматуру выбранного RectSet. Проверяем весь снимок, а не один счетчик.
Private Sub CheckCircleInactive(ByRef stats As TGeometryConfigStats, ByVal systemRange As Object, ByVal target As Object)
    SetKey systemRange, "Geometry.Type", "RectSet"
    Dim originalGeom As CGeometryRegion, originalBars As CRebarLayout
    Dim geom As CGeometryRegion, bars As CRebarLayout, key As Variant, index As Long, equal As Boolean
    ReadGeometry originalGeom, originalBars
    For Each key In Array("Circle.Diameter", "Rebar.AxisDistance", "Rebar.Count", "Rebar.Diameter", _
            "Rebar.Diameter2", "Rebar.Diameter3", "Rebar.Loc2row", "Rebar.Loc3row")
        ConfigureCircle target
        If InStr(1, CStr(key), "Loc", vbBinaryCompare) > 0 Then
            SetKey target, CStr(key), "SideBySide"
        Else
            SetKey target, CStr(key), 30#
        End If
        ReadGeometry geom, bars
        equal = geom.MinX = originalGeom.MinX And geom.MaxX = originalGeom.MaxX And _
            geom.MinY = originalGeom.MinY And geom.MaxY = originalGeom.MaxY And bars.Count = originalBars.Count
        If bars.Count = originalBars.Count Then
            For index = 1 To bars.Count
                If bars.X(index) <> originalBars.X(index) Or bars.Y(index) <> originalBars.Y(index) Or _
                        bars.Diameter(index) <> originalBars.Diameter(index) Then equal = False
            Next index
        End If
        Check stats, "circleConfig.inactive." & CStr(key), equal
    Next key
End Sub

' Читает текущие ячейки заново и использует штатные единицы/registry;
' прямые вызовы чистого Build здесь не могли бы обнаружить обрыв Config-передачи.
Private Sub ReadGeometry(ByRef geom As CGeometryRegion, ByRef bars As CRebarLayout)
    Dim settings As CSystemSettingsReader, units As CUnitSystem, registry As CSectionTypeRegistry
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    Set registry = New CSectionTypeRegistry
    Set geom = registry.CreateGeometry(settings, units)
    Set bars = registry.CreateRebars(geom, settings, units)
End Sub

' Меняет реальную ячейку значения по ключу таблицы; отсутствие ключа является
' ошибкой fixture, а не скрытым созданием новой настройки или default.
Private Sub SetKey(ByVal target As Object, ByVal key As String, ByVal value As Variant)
    Dim rowIndex As Long
    For rowIndex = 2 To target.Rows.Count
        If CStr(target.Cells(rowIndex, 1).Value2) = key Then
            target.Cells(rowIndex, 2).Value2 = value
            Exit Sub
        End If
    Next rowIndex
    Err.Raise vbObjectError + 5960, "modTestGeometryConfig", "Не найдена строка fixture: " & key
End Sub

' Сравнивает восстановленные двумерные Formula-массивы без округления,
' включая пустые значения; формы исходных диапазонов не изменяются тестом.
Private Function SameFormula(ByVal actual As Variant, ByVal expected As Variant) As Boolean
    Dim rowIndex As Long, columnIndex As Long
    For rowIndex = LBound(expected, 1) To UBound(expected, 1)
        For columnIndex = LBound(expected, 2) To UBound(expected, 2)
            If CStr(actual(rowIndex, columnIndex)) <> CStr(expected(rowIndex, columnIndex)) Then Exit Function
        Next columnIndex
    Next rowIndex
    SameFormula = True
End Function

' Фиксирует абсолютный геометрический допуск 1e-6 мм и оба значения oracle;
' этот допуск не меняет настройки рабочего solver-а.
Private Sub Near(ByRef stats As TGeometryConfigStats, ByVal name As String, ByVal actual As Double, ByVal expected As Double)
    Check stats, name, Abs(actual - expected) <= 0.000001
    stats.Report = stats.Report & "CIRCLE_CONFIG_NUMBER: " & name & "; actual=" & CStr(actual) & "; expected=" & CStr(expected) & vbCrLf
End Sub

' Сохраняет счет и строку каждого assertion, не останавливая набор при несовпадении.
Private Sub Check(ByRef stats As TGeometryConfigStats, ByVal name As String, ByVal condition As Boolean)
    If condition Then
        stats.Passed = stats.Passed + 1
        stats.Report = stats.Report & "OK: " & name & vbCrLf
    Else
        stats.Failed = stats.Failed + 1
        stats.Report = stats.Report & "FAIL: " & name & vbCrLf
    End If
End Sub

' ======================================================================
' ДЛЯ ТЕСТОВ: ОБЯЗАТЕЛЬНЫЙ ВВОД АКТИВНОЙ ГЕОМЕТРИИ
' ======================================================================

' Проверяет реальные таблицы четырех форм двумя независимыми входами:
' фабрика геометрии и BuildFromSettings арматуры. Не заполняет пропущенные
' значения вместо пользователя; все Formula-массивы восстанавливает после
' серии, включая неудачу. Количество assertions доступно полному UI-suite.
Public Function RunAudit03RequiredGeometryInputTests(Optional ByRef passed As Long = 0, _
        Optional ByRef failed As Long = 0) As String
    Dim stats As TGeometryConfigStats, ranges(0 To 5) As Object, saved(0 To 5) As Variant
    Dim names As Variant, shapes As Variant, shape As String, fields As Variant
    Dim i As Long, j As Long, item As Variant, route As Variant, bad As Variant
    Dim cell As Object, baseline As Variant, current As Variant, description As String
    Dim errorNumber As Long, prefix As String, caseIndex As Long
    names = Array("rngSystemSettings", "rngUnitSettings", "rngCircleGeometry", _
        "rngRectSetGeometry", "rngRoundedRectangleGeometry", "rngHollowRectangleGeometry")
    shapes = Array("Circle", "RectSet", "RoundedRectangle", "HollowRectangle")
    On Error GoTo FailedRun
    For i = 0 To 5
        Set ranges(i) = ThisWorkbook.Names.Item(CStr(names(i))).RefersToRange
        saved(i) = ranges(i).Formula
    Next i
    ranges(1).Cells(2, 2).Value2 = "mm"
    For i = 0 To 3
        shape = CStr(shapes(i))
        SetKey ranges(0), "Geometry.Type", shape
        If shape = "Circle" Then
            ConfigureCircle ranges(i + 2)
            Set cell = RequiredGeometryKeyCell(ranges(i + 2), "Circle.Diameter")
            fields = Array(Array("Circle.Diameter", cell.Row - ranges(i + 2).Row + 1, 2, CDbl(cell.Value2)))
        Else
            ConfigureShape ranges(i + 2), shape
            fields = ShapeFields(shape)
            ConfigureRequiredGeometryRebars ranges(i + 2), shape
        End If
        For Each route In Array("Geometry", "Rebars")
            baseline = RequiredGeometryRouteSnapshot(shape, CStr(route))
            For Each item In fields
                Set cell = ranges(i + 2).Cells(CLng(item(1)), CLng(item(2)))
                caseIndex = 0
                For Each bad In Array(vbNullString, "TODO", "abc", CVErr(2015))
                    caseIndex = caseIndex + 1
                    cell.Value2 = bad
                    prefix = "requiredGeometry." & CStr(item(0)) & "." & CStr(route) & "." & CStr(caseIndex)
                    RequiredGeometryProgress prefix, stats.Report
                    errorNumber = 0: description = vbNullString
                    On Error Resume Next
                    current = RequiredGeometryRouteSnapshot(shape, CStr(route))
                    errorNumber = Err.Number: description = Err.Description
                    Err.Clear
                    On Error GoTo FailedRun
                    Check stats, prefix & ".reject", errorNumber <> 0
                    Check stats, prefix & ".key", InStr(1, description, CStr(item(0)), vbBinaryCompare) > 0
                    Check stats, prefix & ".address", InStr(1, description, "Config", vbTextCompare) > 0 And _
                        InStr(1, description, cell.Address(False, False), vbTextCompare) > 0
                    stats.Report = stats.Report & "REQUIRED_GEOMETRY_ERROR: " & prefix & "; error=" & _
                        CStr(errorNumber) & "; " & description & vbCrLf
                    cell.Value2 = item(3)
                    current = RequiredGeometryRouteSnapshot(shape, CStr(route))
                    Check stats, prefix & ".recovery", SameGeometryRouteSnapshot(baseline, current)
                Next bad
            Next item
        Next route
    Next i
    CheckRequiredGeometryInactive stats, ranges(0), ranges(3), ranges(4)
    CheckRequiredGeometryCommon stats, ranges(0), ranges(2)
    GoTo Restore
FailedRun:
    stats.Failed = stats.Failed + 1
    stats.Report = stats.Report & "FAIL: requiredGeometry.runtime; " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Restore:
    On Error GoTo FailedRestore
    For j = 0 To 5
        If Not ranges(j) Is Nothing Then
            ranges(j).Formula = saved(j)
            Check stats, "requiredGeometry.restore." & CStr(names(j)), SameFormula(ranges(j).Formula, saved(j))
        End If
    Next j
    GoTo Finish
FailedRestore:
    stats.Failed = stats.Failed + 1
    stats.Report = stats.Report & "FAIL: requiredGeometry.restore.runtime; " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Finish:
    passed = stats.Passed: failed = stats.Failed
    RunAudit03RequiredGeometryInputTests = stats.Report & "TOTAL_REQUIRED_GEOMETRY: passed=" & _
        CStr(passed) & "; failed=" & CStr(failed) & vbCrLf
End Function

' Настраивает допустимые стержни для измененных размеров fixture. Не меняет
' тестируемые геометрические поля и не зависит от исходного армирования книги.
Private Sub ConfigureRequiredGeometryRebars(ByVal target As Object, ByVal shape As String)
    Dim rowIndex As Long
    If shape = "RectSet" Then
        For rowIndex = 11 To 18
            target.Cells(rowIndex, 2).Value2 = 20#: target.Cells(rowIndex, 3).Value2 = 16#
            target.Cells(rowIndex, 4).Value2 = 3
            target.Cells(rowIndex, 5).Value2 = 20#: target.Cells(rowIndex, 6).Value2 = 20#
        Next rowIndex
        For rowIndex = 21 To 28
            target.Cells(rowIndex, 2).Value2 = 0#: target.Cells(rowIndex, 5).Value2 = 0#
        Next rowIndex
    Else
        If shape = "RoundedRectangle" Then rowIndex = 13 Else rowIndex = 11
        Dim firstRow As Long, lastRow As Long, extraFirst As Long
        firstRow = rowIndex: lastRow = firstRow + 3
        If shape = "RoundedRectangle" Then extraFirst = firstRow + 6 Else extraFirst = firstRow + 10
        If shape = "HollowRectangle" Then lastRow = firstRow + 7
        For rowIndex = firstRow To lastRow
            target.Cells(rowIndex, 2).Value2 = 40#: target.Cells(rowIndex, 3).Value2 = 16#
            If rowIndex < firstRow + 4 Then target.Cells(rowIndex, 4).Value2 = 3
        Next rowIndex
        For rowIndex = extraFirst To extraFirst + lastRow - firstRow
            target.Cells(rowIndex, 2).Value2 = 0#: target.Cells(rowIndex, 5).Value2 = 0#
        Next rowIndex
    End If
End Sub

' Читает полный Config и выбранный производственный вход, не вызывая фабрику
' формы перед builder-ом: так тест обнаруживает его собственные defaults.
' Снимок арматуры содержит все координаты/диаметры, а не только общий count.
Private Function RequiredGeometryRouteSnapshot(ByVal shape As String, ByVal route As String) As Variant
    Dim settings As CSystemSettingsReader, units As CUnitSystem, registry As CSectionTypeRegistry
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    Set registry = New CSectionTypeRegistry
    If route = "Geometry" Then
        RequiredGeometryRouteSnapshot = ShapeSnapshot(registry.CreateGeometry(settings, units))
        Exit Function
    End If
    Dim circleBuilder As CCircleRebarLayoutBuilder, rect As CRectSetRebarLayoutBuilder
    Dim rounded As CRoundedRectRebarLayoutBuilder, hollow As CHollowRectRebarLayoutBuilder
    Dim bars As CRebarLayout
    Select Case shape
        Case "Circle": Set circleBuilder = New CCircleRebarLayoutBuilder: Set bars = circleBuilder.BuildFromSettings(settings, units)
        Case "RectSet": Set rect = New CRectSetRebarLayoutBuilder: Set bars = rect.BuildFromSettings(settings, units)
        Case "RoundedRectangle": Set rounded = New CRoundedRectRebarLayoutBuilder: Set bars = rounded.BuildFromSettings(settings, units)
        Case "HollowRectangle": Set hollow = New CHollowRectRebarLayoutBuilder: Set bars = hollow.BuildFromSettings(settings, units)
    End Select
    Dim values() As Double, i As Long
    ReDim values(0 To bars.Count * 3)
    values(0) = bars.Count
    For i = 1 To bars.Count
        values(i * 3 - 2) = bars.X(i): values(i * 3 - 1) = bars.Y(i): values(i * 3) = bars.Diameter(i)
    Next i
    RequiredGeometryRouteSnapshot = values
End Function

' Сравнивает расчетные snapshots в том же абсолютном геометрическом допуске,
' что и shape-tests; неодинаковый размер массива означает различие раскладки.
Private Function SameGeometryRouteSnapshot(ByVal first As Variant, ByVal second As Variant) As Boolean
    If UBound(first) <> UBound(second) Then Exit Function
    Dim i As Long
    For i = LBound(first) To UBound(first)
        If Abs(CDbl(first(i)) - CDbl(second(i))) > 0.000001 Then Exit Function
    Next i
    SameGeometryRouteSnapshot = True
End Function

' Возвращает фактическую ячейку ключевой таблицы. Потеря строки fixture
' является ошибкой теста, а не добавлением отсутствующего Config/default.
Private Function RequiredGeometryKeyCell(ByVal target As Object, ByVal key As String) As Object
    Dim rowIndex As Long
    For rowIndex = 2 To target.Rows.Count
        If CStr(target.Cells(rowIndex, 1).Value2) = key Then
            Set RequiredGeometryKeyCell = target.Cells(rowIndex, 2)
            Exit Function
        End If
    Next rowIndex
    Err.Raise vbObjectError + 5965, "modTestGeometryConfig", "Не найдена строка fixture: " & key
End Function

' Проверяет неактивные размеры Rectangle и W/R2 сторон Simple. Эти значения
' не участвуют в форме/раскладке, поэтому не должны требовать числового ввода.
' Ошибки формул Excel здесь не подменяются пустым значением reader-а.
Private Sub CheckRequiredGeometryInactive(ByRef stats As TGeometryConfigStats, ByVal systemRange As Object, _
        ByVal rectRange As Object, ByVal roundedRange As Object)
    Dim route As Variant, item As Variant, value As Variant, baseline As Variant, current As Variant
    Dim description As String, errorNumber As Long, prefix As String
    SetKey systemRange, "Geometry.Type", "RectSet"
    ConfigureShape rectRange, "RectSet": ConfigureRequiredGeometryRebars rectRange, "RectSet"
    rectRange.Cells(3, 2).Value2 = "Rectangle"
    For Each route In Array("Geometry", "Rebars")
        baseline = RequiredGeometryRouteSnapshot("RectSet", CStr(route))
        For Each item In Array(Array(8, 3), Array(8, 4), Array(4, 2))
            For Each value In Array(vbNullString, "TODO", "abc")
                rectRange.Cells(CLng(item(0)), CLng(item(1))).Value2 = value
                errorNumber = 0: description = vbNullString
                On Error Resume Next
                current = RequiredGeometryRouteSnapshot("RectSet", CStr(route))
                errorNumber = Err.Number: description = Err.Description
                Err.Clear: On Error GoTo 0
                prefix = "requiredGeometry.inactive.Rectangle." & CStr(route) & "." & CStr(item(0)) & "." & CStr(item(1)) & "." & CStr(value)
                Check stats, prefix, errorNumber = 0
                If errorNumber = 0 Then Check stats, prefix & ".snapshot", SameGeometryRouteSnapshot(baseline, current)
                stats.Report = stats.Report & "INACTIVE_GEOMETRY: " & prefix & "; error=" & CStr(errorNumber) & "; " & description & vbCrLf
            Next value
            rectRange.Cells(CLng(item(0)), CLng(item(1))).Value2 = 100#
        Next item
    Next route
    SetKey systemRange, "Geometry.Type", "RoundedRectangle"
    ConfigureShape roundedRange, "RoundedRectangle": ConfigureRequiredGeometryRebars roundedRange, "RoundedRectangle"
    roundedRange.Cells(7, 2).Value2 = "Simple": roundedRange.Cells(7, 3).Value2 = "Simple"
    For Each route In Array("Geometry", "Rebars")
        baseline = RequiredGeometryRouteSnapshot("RoundedRectangle", CStr(route))
        For Each item In Array(Array(8, 2), Array(8, 3), Array(10, 2), Array(10, 3))
            For Each value In Array(vbNullString, "TODO", "abc")
                roundedRange.Cells(CLng(item(0)), CLng(item(1))).Value2 = value
                errorNumber = 0: description = vbNullString
                On Error Resume Next
                current = RequiredGeometryRouteSnapshot("RoundedRectangle", CStr(route))
                errorNumber = Err.Number: description = Err.Description
                Err.Clear: On Error GoTo 0
                prefix = "requiredGeometry.inactive.Simple." & CStr(route) & "." & CStr(item(0)) & "." & CStr(item(1)) & "." & CStr(value)
                Check stats, prefix, errorNumber = 0
                If errorNumber = 0 Then Check stats, prefix & ".snapshot", SameGeometryRouteSnapshot(baseline, current)
                stats.Report = stats.Report & "INACTIVE_GEOMETRY: " & prefix & "; error=" & CStr(errorNumber) & "; " & description & vbCrLf
            Next value
            roundedRange.Cells(CLng(item(0)), CLng(item(1))).Value2 = 0#
        Next item
    Next route
End Sub

' Обязательные общие Geometry/mesh поля проверяет штатная сборка модели.
' В missing-варианте временно меняется ключ, а не удаляются пользовательские
' строки; recovery восстанавливает его до следующего сценария.
Private Sub CheckRequiredGeometryCommon(ByRef stats As TGeometryConfigStats, ByVal systemRange As Object, ByVal circleRange As Object)
    Dim key As Variant, bad As Variant, cell As Object, savedValue As Variant
    Dim settings As CSystemSettingsReader, units As CUnitSystem, model As CSectionModel
    Dim errorNumber As Long, description As String, prefix As String, caseIndex As Long
    SetKey systemRange, "Geometry.Type", "Circle": SetKey systemRange, "Geometry.Source", "Generated"
    ConfigureCircle circleRange
    For Each key In Array("Geometry.Type", "Geometry.Source", "Mesh.StepX", "Mesh.StepY", "Mesh.BoundarySubdivisions")
        Set cell = RequiredGeometryKeyCell(systemRange, CStr(key))
        savedValue = cell.Formula: caseIndex = 0
        For Each bad In Array(vbNullString, "TODO", "abc", CVErr(2015), "__MISSING_KEY__")
            caseIndex = caseIndex + 1
            If Not IsError(bad) Then
                If CStr(bad) = "__MISSING_KEY__" Then cell.Offset(0, -1).Value2 = "Audit03.Hidden." & CStr(key) Else cell.Value2 = bad
            Else
                cell.Value2 = bad
            End If
            errorNumber = 0: description = vbNullString
            RequiredGeometryProgress "requiredGeometry.common." & CStr(key) & "." & CStr(caseIndex), stats.Report
            On Error Resume Next
            Set model = RequiredGeometryModelFromWorkbook()
            errorNumber = Err.Number: description = Err.Description
            Err.Clear: On Error GoTo 0
            prefix = "requiredGeometry.common." & CStr(key) & "." & CStr(caseIndex)
            Check stats, prefix & ".reject", errorNumber <> 0
            Check stats, prefix & ".key", InStr(1, description, CStr(key), vbBinaryCompare) > 0
            Check stats, prefix & ".location", InStr(1, description, "Config", vbTextCompare) > 0
            If caseIndex < 5 Then Check stats, prefix & ".address", InStr(1, description, cell.Address(False, False), vbTextCompare) > 0
            stats.Report = stats.Report & "REQUIRED_GEOMETRY_ERROR: " & prefix & "; error=" & CStr(errorNumber) & "; " & description & vbCrLf
            cell.Offset(0, -1).Value2 = key: cell.Formula = savedValue
            Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
            Set units = New CUnitSystem: units.LoadFromSettings settings
            Set model = BuildWorkbookSectionModel(ThisWorkbook, settings, units)
            Check stats, prefix & ".recovery", model.ConcreteCount > 0 And model.RebarCount = 24
        Next bad
    Next key
End Sub

' Выполняет один обычный production-маршрут без Resume Next между reader,
' units и builder. Первая ошибка ввода сохраняется, следующие шаги при ней
' не вызываются; внешний тест перехватывает единственный вызов целиком.
Private Function RequiredGeometryModelFromWorkbook() As CSectionModel
    Dim settings As CSystemSettingsReader, units As CUnitSystem
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    Set RequiredGeometryModelFromWorkbook = BuildWorkbookSectionModel(ThisWorkbook, settings, units)
End Function

' Сохраняет уже выполненные assertions и следующий сценарий рядом с временной
' книгой. Watchdog забирает этот журнал при timeout; отметка START не считается
' успешной проверкой. Ошибка записи диагностики не изменяет результат теста.
Private Sub RequiredGeometryProgress(ByVal nextCase As String, ByVal report As String)
    On Error Resume Next
    Dim fso As Object, stream As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set stream = fso.CreateTextFile(ThisWorkbook.Path & "\Audit03_Search_Progress.txt", True, True)
    stream.Write report & "START_REQUIRED_GEOMETRY: " & nextCase & vbCrLf
    stream.Close
    On Error GoTo 0
End Sub

' ==========================================================================
' ДЛЯ ТЕСТОВ: ПОАДРЕСНЫЕ ПАРАМЕТРЫ ГРАНЕЙ И ДОПОЛНИТЕЛЬНЫХ РЯДОВ
' ==========================================================================
' Читает каждую физическую грань через настоящий Config/registry pipeline.
' Прямоугольные контуры дают независимый oracle нормали, диаметра и количества;
' арматура Opening проверяется в направлении от пустоты в бетон. Все таблицы
' восстанавливаются целиком; этот набор не решает НДС и не назначает его статус.
Public Function RunAudit03RebarFieldConfigTests(Optional ByRef passed As Long = 0, _
        Optional ByRef failed As Long = 0) As String
    Dim stats As TGeometryConfigStats, ranges(0 To 5) As Object, saved(0 To 5) As Variant
    Dim names As Variant, shapes As Variant, shape As String, face As Variant, field As Variant
    Dim i As Long, cases As Long, target As Object
    names = Array("rngSystemSettings", "rngUnitSettings", "rngRectSetGeometry", _
        "rngRoundedRectangleGeometry", "rngHollowRectangleGeometry", "rngCircleGeometry")
    shapes = Array("RectSet", "RoundedRectangle", "HollowRectangle")
    On Error GoTo FailedRun
    For i = 0 To 5
        Set ranges(i) = ThisWorkbook.Names.Item(CStr(names(i))).RefersToRange
        saved(i) = ranges(i).Formula
    Next i
    ranges(1).Cells(2, 2).Value2 = "mm"
    SetKey ranges(0), "Geometry.Source", "Generated"
    For i = 0 To 2
        shape = CStr(shapes(i)): Set target = ranges(i + 2)
        SetKey ranges(0), "Geometry.Type", shape
        For Each face In RebarFieldFaces(shape)
            For Each field In RebarFaceFields(shape, face)
                ConfigureRebarFieldFixture target, shape, CBool(field(3))
                CheckRebarFieldEffect stats, target, shape, face, field, cases
                CheckRebarFieldInvalid stats, target, shape, face, field, cases
            Next field
        Next face
    Next i
    CheckInactiveRebarFields stats, ranges(0), ranges(5), ranges(2), ranges(3), ranges(4), cases
    GoTo Restore
FailedRun:
    stats.Failed = stats.Failed + 1
    stats.Report = stats.Report & "FAIL: rebarFields.runtime; " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Restore:
    On Error GoTo FailedRestore
    For i = 0 To 5
        If Not ranges(i) Is Nothing Then
            ranges(i).Formula = saved(i)
            Check stats, "rebarFields.restore." & CStr(names(i)), SameFormula(ranges(i).Formula, saved(i))
        End If
    Next i
    GoTo Finish
FailedRestore:
    stats.Failed = stats.Failed + 1
    stats.Report = stats.Report & "FAIL: rebarFields.restore.runtime; " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Finish:
    passed = stats.Passed: failed = stats.Failed
    RunAudit03RebarFieldConfigTests = stats.Report & "TOTAL_REBAR_FIELD_CONFIG: passed=" & CStr(passed) & _
        "; failed=" & CStr(failed) & "; cases=" & CStr(cases) & vbCrLf
End Function

' Проверяет неактивность на настоящем workbook-маршруте, а не только в Build.
' Общие loc/bind RectSet не считаются неактивными, пока активна вторая сторона;
' для этого сценария выключаются оба первых ряда соответствующей пары граней.
' as/отступы отсутствующего ряда здесь не объявляются неактивными автоматически:
' например cover Opening может ограничивать проекции соседних стержней.
Private Sub CheckInactiveRebarFields(ByRef stats As TGeometryConfigStats, ByVal systemRange As Object, _
        ByVal circleRange As Object, ByVal rectRange As Object, ByVal roundedRange As Object, _
        ByVal hollowRange As Object, ByRef cases As Long)
    Dim shape As Variant, face As Variant, field As Variant, mode As Variant, target As Object
    Dim kind As String, rowIndex As Long, pairRow As Long
    For Each shape In Array("RectSet", "RoundedRectangle", "HollowRectangle")
        Select Case CStr(shape)
            Case "RectSet": Set target = rectRange
            Case "RoundedRectangle": Set target = roundedRange
            Case "HollowRectangle": Set target = hollowRange
        End Select
        SetKey systemRange, "Geometry.Type", CStr(shape)
        For Each face In RebarFieldFaces(CStr(shape))
            For Each field In RebarFaceFields(CStr(shape), face)
                kind = CStr(field(0))
                If Left$(kind, 3) = "loc" Or Left$(kind, 4) = "bind" Then
                    ConfigureRebarFieldFixture target, CStr(shape), False
                    CheckInactiveRebarCell stats, target, CStr(shape), face, field, "extraRowsOff", cases
                End If
                ' Отступы отключенной наружной линии не нужны. Исключение:
                ' as Opening ограничивает продолжение соседних проекций.
                Dim inactiveOffset As Boolean
                inactiveOffset = (kind = "as" Or kind = "start" Or kind = "end")
                If CStr(shape) = "HollowRectangle" And Left$(CStr(face(0)), 8) = "Opening." Then inactiveOffset = False
                If kind = "d2" Or kind = "d3" Or Left$(kind, 3) = "loc" Or Left$(kind, 4) = "bind" Or inactiveOffset Then
                    For Each mode In Array("diameterOff", "countOff")
                        If CStr(mode) <> "countOff" Or Left$(CStr(face(0)), 8) <> "Opening." Then
                            ConfigureRebarFieldFixture target, CStr(shape), True
                            rowIndex = CLng(face(1))
                            If CStr(mode) = "countOff" Then
                                target.Cells(rowIndex, 4).Value2 = 0
                            Else
                                target.Cells(rowIndex, 3).Value2 = 0#
                            End If
                            If CStr(shape) = "RectSet" Then
                                pairRow = rowIndex + 1
                                If (rowIndex - 11) Mod 2 <> 0 Then pairRow = rowIndex - 1
                                If CStr(mode) = "countOff" Then
                                    target.Cells(pairRow, 4).Value2 = 0
                                Else
                                    target.Cells(pairRow, 3).Value2 = 0#
                                End If
                            End If
                            CheckInactiveRebarCell stats, target, CStr(shape), face, field, CStr(mode), cases
                        End If
                    Next mode
                End If
                If CStr(shape) = "RectSet" And CLng(face(1)) >= 15 Then
                    ConfigureRebarFieldFixture target, "RectSet", True
                    target.Cells(3, 2).Value2 = "Rectangle"
                    CheckInactiveRebarCell stats, target, "RectSet", face, field, "lowerRectangleOff", cases
                End If
            Next field
        Next face
        ' Другая форма не потребляет ни размеры, ни армирование этой таблицы.
        ' Корректный Circle задается независимо от редактируемой грани.
        ConfigureRebarFieldFixture target, CStr(shape), True
        ConfigureCircle circleRange
        SetKey systemRange, "Geometry.Type", "Circle"
        For Each face In RebarFieldFaces(CStr(shape))
            For Each field In RebarFaceFields(CStr(shape), face)
                CheckInactiveRebarCell stats, target, "Circle", face, field, _
                    "otherShape." & CStr(shape), cases
            Next field
        Next face
    Next shape
End Sub

' Меняет только одну неиспользуемую ячейку и сравнивает полный расчетный layout
' с независимым снимком до мутации. Ошибка формулы проверяется отдельно от текста;
' после каждой попытки исходная Formula восстанавливается до следующего чтения.
Private Sub CheckInactiveRebarCell(ByRef stats As TGeometryConfigStats, ByVal target As Object, _
        ByVal activeShape As String, ByVal face As Variant, ByVal field As Variant, _
        ByVal mode As String, ByRef cases As Long)
    Dim cell As Object, savedValue As Variant, bad As Variant, baseline As Variant, current As Variant
    Dim prefix As String, number As Long, description As String, index As Long
    Set cell = RebarFieldCell(target, CLng(field(1)), CLng(field(2)))
    savedValue = cell.Formula
    baseline = RequiredGeometryRouteSnapshot(activeShape, "Rebars")
    For Each bad In Array(vbNullString, "abc", CVErr(2015))
        index = index + 1: cases = cases + 1
        prefix = "rebarFields.inactive." & mode & "." & CStr(face(0)) & "." & CStr(field(0)) & "." & CStr(index)
        cell.Value2 = bad: number = 0: description = vbNullString
        RequiredGeometryProgress prefix, stats.Report
        On Error Resume Next
        current = RequiredGeometryRouteSnapshot(activeShape, "Rebars")
        number = Err.Number: description = Err.Description
        Err.Clear: On Error GoTo 0
        Check stats, prefix & ".noInterference", number = 0
        If number = 0 Then Check stats, prefix & ".snapshot", SameGeometryRouteSnapshot(baseline, current)
        stats.Report = stats.Report & "INACTIVE_REBAR_FIELD: " & prefix & "; error=" & CStr(number) & "; " & description & vbCrLf
        cell.Formula = savedValue
        current = RequiredGeometryRouteSnapshot(activeShape, "Rebars")
        Check stats, prefix & ".recovery", SameGeometryRouteSnapshot(baseline, current)
    Next bad
End Sub

' Описывает физические грани: имя layout-группы, строка первого/дополнительных
' рядов, нормаль внутрь бетона и касательная обхода. Это независимые сведения
' прямоугольного fixture, а не вызов production-функции смещения ряда.
Private Function RebarFieldFaces(ByVal shape As String) As Variant
    Select Case shape
        Case "RectSet"
            RebarFieldFaces = Array(Array("H1.as_1", 11, 21, 1#, 0#, 0#, 1#), _
                Array("H1.as_2", 12, 22, -1#, 0#, 0#, -1#), _
                Array("B1.as_1", 13, 23, 0#, -1#, 1#, 0#), _
                Array("B1.as_2", 14, 24, 0#, 1#, -1#, 0#), _
                Array("H2.as_1", 15, 25, 1#, 0#, 0#, 1#), _
                Array("H2.as_2", 16, 26, -1#, 0#, 0#, -1#), _
                Array("B2.as_1", 17, 27, 0#, -1#, 1#, 0#), _
                Array("B2.as_2", 18, 28, 0#, 1#, -1#, 0#))
        Case "RoundedRectangle"
            RebarFieldFaces = Array(Array("H.Left", 13, 19, 1#, 0#, 0#, 1#), _
                Array("H.Right", 14, 20, -1#, 0#, 0#, -1#), _
                Array("B.Top", 15, 21, 0#, -1#, 1#, 0#), _
                Array("B.Bottom", 16, 22, 0#, 1#, -1#, 0#))
        Case "HollowRectangle"
            RebarFieldFaces = Array(Array("H.Left", 11, 21, 1#, 0#, 0#, 1#), _
                Array("H.Right", 12, 22, -1#, 0#, 0#, -1#), _
                Array("B.Top", 13, 23, 0#, -1#, 1#, 0#), _
                Array("B.Bottom", 14, 24, 0#, 1#, -1#, 0#), _
                Array("Opening.H.Left", 15, 25, -1#, 0#, 0#, 1#), _
                Array("Opening.H.Right", 16, 26, 1#, 0#, 0#, -1#), _
                Array("Opening.B.Top", 17, 27, 0#, 1#, 1#, 0#), _
                Array("Opening.B.Bottom", 18, 28, 0#, -1#, -1#, 0#))
        Case Else
            Err.Raise vbObjectError + 5966, "modTestGeometryConfig", "Неизвестная форма арматурного fixture: " & shape
    End Select
End Function

' Перечисляет настоящие input-ячейки одной грани: тип поля, относительная
' строка/колонка и необходимость активных дополнительных рядов. Количество
' Opening отсутствует по контракту и не превращается в искусственный ввод.
Private Function RebarFaceFields(ByVal shape As String, ByVal face As Variant) As Variant
    Dim mainRow As Long, extraRow As Long
    mainRow = CLng(face(1)): extraRow = CLng(face(2))
    If shape = "RectSet" Then
        RebarFaceFields = Array(Array("as", mainRow, 2, False), Array("d", mainRow, 3, False), _
            Array("n", mainRow, 4, False), Array("start", mainRow, 5, False), Array("end", mainRow, 6, False), _
            Array("d2", extraRow, 2, True), Array("loc2", extraRow, 3, True), Array("bind2", extraRow, 4, True), _
            Array("d3", extraRow, 5, True), Array("loc3", extraRow, 6, True), Array("bind3", extraRow, 7, True))
    ElseIf Left$(CStr(face(0)), 8) = "Opening." Then
        RebarFaceFields = Array(Array("as", mainRow, 2, False), Array("d", mainRow, 3, False), _
            Array("d2", extraRow, 2, True), Array("loc2", extraRow, 3, True), Array("bind2", extraRow, 4, True), _
            Array("d3", extraRow, 5, True), Array("loc3", extraRow, 6, True), Array("bind3", extraRow, 7, True))
    Else
        RebarFaceFields = Array(Array("as", mainRow, 2, False), Array("d", mainRow, 3, False), Array("n", mainRow, 4, False), _
            Array("d2", extraRow, 2, True), Array("loc2", extraRow, 3, True), Array("bind2", extraRow, 4, True), _
            Array("d3", extraRow, 5, True), Array("loc3", extraRow, 6, True), Array("bind3", extraRow, 7, True))
    End If
End Function

' Задает достаточно просторные прямые контуры и пять базовых стержней каждой
' грани. Дополнительные ряды при необходимости имеют разные диаметры 12/10;
' стержни не касаются соседних линий, поэтому count-oracle не зависит от дублей.
Private Sub ConfigureRebarFieldFixture(ByVal target As Object, ByVal shape As String, ByVal extras As Boolean)
    ConfigureShape target, shape
    Dim face As Variant, mainRow As Long, extraRow As Long, column As Long
    Select Case shape
        Case "RectSet"
            target.Cells(3, 2).Value2 = "TwoRectangles": target.Cells(4, 2).Value2 = 0#
            For column = 1 To 4: target.Cells(8, column).Value2 = 800#: Next column
        Case "RoundedRectangle"
            target.Cells(3, 2).Value2 = 800#: target.Cells(3, 3).Value2 = 800#
            For column = 2 To 3
                target.Cells(7, column).Value2 = "Simple": target.Cells(8, column).Value2 = 0#
                target.Cells(9, column).Value2 = 0#: target.Cells(10, column).Value2 = 0#
            Next column
        Case "HollowRectangle"
            target.Cells(3, 2).Value2 = 0#: target.Cells(4, 2).Value2 = 0#
            target.Cells(7, 1).Value2 = 1600#: target.Cells(7, 2).Value2 = 1600#: target.Cells(7, 3).Value2 = 0#
            target.Cells(7, 4).Value2 = 1000#: target.Cells(7, 5).Value2 = 1000#: target.Cells(7, 6).Value2 = 0#
    End Select
    For Each face In RebarFieldFaces(shape)
        mainRow = CLng(face(1)): extraRow = CLng(face(2))
        target.Cells(mainRow, 2).Value2 = 50#: target.Cells(mainRow, 3).Value2 = 16#
        If Left$(CStr(face(0)), 8) <> "Opening." Then target.Cells(mainRow, 4).Value2 = 5
        If shape = "RectSet" Then
            target.Cells(mainRow, 5).Value2 = 80#: target.Cells(mainRow, 6).Value2 = 80#
        End If
        If extras Then
            target.Cells(extraRow, 2).Value2 = 12#: target.Cells(extraRow, 5).Value2 = 10#
        Else
            target.Cells(extraRow, 2).Value2 = 0#: target.Cells(extraRow, 5).Value2 = 0#
        End If
        RebarFieldCell(target, extraRow, 3).Value2 = "Stacked"
        RebarFieldCell(target, extraRow, 6).Value2 = "Stacked"
        RebarFieldCell(target, extraRow, 4).Value2 = "EachBar"
        RebarFieldCell(target, extraRow, 7).Value2 = "EachBar"
    Next face
End Sub

' Меняет одно поле и проверяет предсказуемый физический эффект. Для as/t
' проверяется подписанное перемещение, для d/n независимые diameter/count,
' для loc/bind положение и число дополнительных стержней относительно первого.
Private Sub CheckRebarFieldEffect(ByRef stats As TGeometryConfigStats, ByVal target As Object, _
        ByVal shape As String, ByVal face As Variant, ByVal field As Variant, ByRef cases As Long)
    Dim geom As CGeometryRegion, baseline As CRebarLayout, bars As CRebarLayout, cell As Object
    Dim before As Variant, after As Variant, kind As String, prefix As String, value As Variant
    Dim rowNumber As Long, expectedCount As Long, expectedDistance As Double
    kind = CStr(field(0)): Set cell = RebarFieldCell(target, CLng(field(1)), CLng(field(2)))
    prefix = "rebarFields." & shape & "." & CStr(face(0)) & "." & kind
    ReadGeometry geom, baseline
    before = RebarFieldRowSummary(baseline, CStr(face(0)), 1)
    Check stats, prefix & ".activeBaseline", CLng(before(0)) > 1
    Select Case kind
        Case "as": value = 60#
        Case "d": value = 20#
        Case "n": value = 7
        Case "start", "end": value = 100#
        Case "d2": value = 18#
        Case "d3": value = 20#
        Case "loc2", "loc3": value = "SideBySide"
        Case "bind2", "bind3": value = "EverySecondBar"
    End Select
    cell.Value2 = value: cases = cases + 1
    RequiredGeometryProgress prefix & ".effect", stats.Report
    ReadGeometry geom, bars
    after = RebarFieldRowSummary(bars, CStr(face(0)), 1)
    Select Case kind
        Case "as"
            Near stats, prefix & ".normalX", CDbl(after(2)) - CDbl(before(2)), 10# * CDbl(face(3))
            Near stats, prefix & ".normalY", CDbl(after(3)) - CDbl(before(3)), 10# * CDbl(face(4))
            Near stats, prefix & ".count", CDbl(after(0)), CDbl(before(0))
        Case "d"
            Near stats, prefix & ".diameter", CDbl(after(1)), 20#
            Near stats, prefix & ".count", CDbl(after(0)), CDbl(before(0))
            Near stats, prefix & ".xUnchanged", CDbl(after(2)), CDbl(before(2))
            Near stats, prefix & ".yUnchanged", CDbl(after(3)), CDbl(before(3))
        Case "n"
            Near stats, prefix & ".count", CDbl(after(0)), 7#
        Case "start", "end"
            expectedDistance = 10#: If kind = "end" Then expectedDistance = -10#
            Near stats, prefix & ".tangentX", CDbl(after(2)) - CDbl(before(2)), expectedDistance * CDbl(face(5))
            Near stats, prefix & ".tangentY", CDbl(after(3)) - CDbl(before(3)), expectedDistance * CDbl(face(6))
            Near stats, prefix & ".count", CDbl(after(0)), CDbl(before(0))
        Case Else
            rowNumber = 2: If Right$(kind, 1) = "3" Then rowNumber = 3
            after = RebarFieldRowSummary(bars, CStr(face(0)), rowNumber)
            expectedCount = CLng(before(0))
            If Left$(kind, 4) = "bind" Then expectedCount = (expectedCount + 1) \ 2
            Near stats, prefix & ".count", CDbl(after(0)), CDbl(expectedCount)
            expectedDistance = 14#: If rowNumber = 3 Then expectedDistance = 25#
            If kind = "d2" Then expectedDistance = 17#
            If kind = "d3" Then expectedDistance = 30#
            If Left$(kind, 3) = "loc" Then
                expectedDistance = 13#
                If rowNumber = 2 Then expectedDistance = 14#
            End If
            If kind = "d2" Or kind = "d3" Then Near stats, prefix & ".diameter", CDbl(after(1)), CDbl(value)
            CheckRebarRowOffsets stats, prefix, bars, CStr(face(0)), rowNumber, face, _
                expectedDistance, Left$(kind, 3) <> "loc"
    End Select
    If kind = "d" Or kind = "n" Or kind = "d2" Or kind = "d3" Then
        cell.ClearContents: cases = cases + 1
        ReadGeometry geom, bars
        rowNumber = 1: If kind = "d2" Then rowNumber = 2
        If kind = "d3" Then rowNumber = 3
        after = RebarFieldRowSummary(bars, CStr(face(0)), rowNumber)
        Near stats, prefix & ".blankDisables", CDbl(after(0)), 0#
        cell.Value2 = 0#: cases = cases + 1
        ReadGeometry geom, bars
        after = RebarFieldRowSummary(bars, CStr(face(0)), rowNumber)
        Near stats, prefix & ".zeroDisables", CDbl(after(0)), 0#
    End If
    ConfigureRebarFieldFixture target, shape, CBool(field(3))
End Sub

' Проверяет активный ошибочный ввод и восстановление того же reader/builder
' маршрута. Допустимая пустота d/n уже проверена отдельно; для обязательных
' as/t/активных loc/bind пустота не должна незаметно выбирать размер/режим.
Private Sub CheckRebarFieldInvalid(ByRef stats As TGeometryConfigStats, ByVal target As Object, _
        ByVal shape As String, ByVal face As Variant, ByVal field As Variant, ByRef cases As Long)
    Dim bad As Variant, badValues As Variant, kind As String, prefix As String, cell As Object
    Dim savedValue As Variant, geom As CGeometryRegion, bars As CRebarLayout, baseline As Variant, current As Variant
    Dim number As Long, description As String, index As Long
    kind = CStr(field(0)): Set cell = RebarFieldCell(target, CLng(field(1)), CLng(field(2)))
    savedValue = cell.Formula: baseline = RequiredGeometryRouteSnapshot(shape, "Rebars")
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
        RequiredGeometryProgress prefix, stats.Report
        On Error Resume Next
        ReadGeometry geom, bars
        number = Err.Number: description = Err.Description
        Err.Clear: On Error GoTo 0
        Check stats, prefix & ".reject", number <> 0
        Check stats, prefix & ".address", InStr(1, description, cell.Address(False, False), vbTextCompare) > 0
        Check stats, prefix & ".repair", InStr(1, description, "Введите", vbTextCompare) > 0 Or _
            InStr(1, description, "Выберите", vbTextCompare) > 0 Or InStr(1, description, "Исправ", vbTextCompare) > 0
        stats.Report = stats.Report & "REBAR_FIELD_ERROR: " & prefix & "; error=" & CStr(number) & "; " & description & vbCrLf
        cell.Formula = savedValue
        current = RequiredGeometryRouteSnapshot(shape, "Rebars")
        Check stats, prefix & ".recovery", SameGeometryRouteSnapshot(baseline, current)
    Next bad
End Sub

' Возвращает фактический anchor ввода: объединенный RectSet-селектор остается
' одним значением для обеих сторон грани. Тест не записывает скрытое значение
' в follower-ячейку и проверяет в сообщении именно адрес доступного пользователю ввода.
Private Function RebarFieldCell(ByVal target As Object, ByVal rowIndex As Long, ByVal column As Long) As Object
    Dim cell As Object
    Set cell = target.Cells(rowIndex, column)
    If cell.MergeCells Then Set cell = cell.MergeArea.Cells(1, 1)
    Set RebarFieldCell = cell
End Function

' Возвращает count, средний диаметр и средние X/Y конкретного ряда конкретной
' смысловой грани. Техническое имя row_N используется только для группировки
' generated layout, не для назначения статусов или разбора комментария ошибки.
Private Function RebarFieldRowSummary(ByVal bars As CRebarLayout, ByVal group As String, ByVal rowNumber As Long) As Variant
    Dim count As Long, diameter As Double, X As Double, Y As Double, i As Long
    For i = 1 To bars.Count
        If bars.BarAnnotationGroupName(i) = group Then
            If InStr(1, bars.BarID(i), ".row_" & CStr(rowNumber) & "-", vbBinaryCompare) > 0 Then
                count = count + 1: diameter = diameter + bars.Diameter(i)
                X = X + bars.X(i): Y = Y + bars.Y(i)
            End If
        End If
    Next i
    If count > 0 Then diameter = diameter / count: X = X / count: Y = Y / count
    RebarFieldRowSummary = Array(count, diameter, X, Y)
End Function

' Находит для каждого дополнительного стержня ближайший базовый стержень
' своей грани. Просторный fixture исключает неоднозначность: шаг первого ряда
' больше любого смещения. Проверяются расстояние и компонента нормали, включая
' противоположное направление у Opening; production offset helper не вызывается.
Private Sub CheckRebarRowOffsets(ByRef stats As TGeometryConfigStats, ByVal prefix As String, _
        ByVal bars As CRebarLayout, ByVal group As String, ByVal rowNumber As Long, ByVal face As Variant, _
        ByVal expectedDistance As Double, ByVal stacked As Boolean)
    Dim i As Long, j As Long, closest As Long, distance As Double, minimum As Double, ordinal As Long
    Dim dx As Double, dy As Double, expectedNormal As Double
    If stacked Then expectedNormal = expectedDistance
    For i = 1 To bars.Count
        If bars.BarAnnotationGroupName(i) = group And _
                InStr(1, bars.BarID(i), ".row_" & CStr(rowNumber) & "-", vbBinaryCompare) > 0 Then
            closest = 0
            For j = 1 To bars.Count
                If bars.BarAnnotationGroupName(j) = group And InStr(1, bars.BarID(j), ".row_1-", vbBinaryCompare) > 0 Then
                    distance = Sqr((bars.X(i) - bars.X(j)) ^ 2 + (bars.Y(i) - bars.Y(j)) ^ 2)
                    If closest = 0 Then
                        closest = j: minimum = distance
                    ElseIf distance < minimum Then
                        closest = j: minimum = distance
                    End If
                End If
            Next j
            ordinal = ordinal + 1
            Check stats, prefix & ".base" & CStr(ordinal), closest > 0
            If closest > 0 Then
                dx = bars.X(i) - bars.X(closest): dy = bars.Y(i) - bars.Y(closest)
                Near stats, prefix & ".distance" & CStr(ordinal), minimum, expectedDistance
                Near stats, prefix & ".normal" & CStr(ordinal), dx * CDbl(face(3)) + dy * CDbl(face(4)), expectedNormal
            End If
        End If
    Next i
    Check stats, prefix & ".offsetsObserved", ordinal > 0
End Sub
