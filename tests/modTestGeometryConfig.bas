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

' Читает форму штатным reader/unit/registry pipeline без создания арматуры:
' проверка одного размера не должна зависеть от чужой раскладки стержней.
Private Function ReadShape() As ISectionGeometry
    Dim settings As CSystemSettingsReader, units As CUnitSystem, registry As CSectionTypeRegistry
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    Set registry = New CSectionTypeRegistry
    Set ReadShape = registry.CreateGeometry(settings, units)
End Function

' Сохраняет расчетные площадь, центр и габариты. Отказ геометрии считается
' ошибкой fixture; снимок не подменяет форму ее исходными getter-параметрами.
Private Function ShapeSnapshot(ByVal geom As ISectionGeometry) As Variant
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
    Dim item As Variant, baseline As Variant, changed As Variant, geom As ISectionGeometry
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

' Независимо проверяет площадь/центр двух простых fixtures. Для HollowRectangle
' учитывает утвержденный polygon-контракт: каждая четверть дуги содержит 12
' хорд; это не подмена текущей геометрии точной круглой дугой.
Private Sub CheckShapeBaseline(ByRef stats As TGeometryConfigStats, ByVal shapeName As String, ByVal actual As Variant)
    Dim area As Double, openingArea As Double, cornerLoss As Double
    If shapeName = "RectSet" Then
        area = 200# * 300# + 600# * 250#
        Near stats, "shapeConfig.RectSet.oracle.area", CDbl(actual(0)), area
        Near stats, "shapeConfig.RectSet.oracle.cx", CDbl(actual(1)), (600# * 250# * 300# + 200# * 300# * 50#) / area
        Near stats, "shapeConfig.RectSet.oracle.cy", CDbl(actual(2)), (600# * 250# * 125# + 200# * 300# * 400#) / area
    ElseIf shapeName = "HollowRectangle" Then
        cornerLoss = 4# - 24# * Sin(GEOM_PI / 24#)
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
    Dim rowIndex As Long, index As Long, bad As Variant, geom As ISectionGeometry, bars As CRebarLayout
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
    Dim key As Variant, geom As ISectionGeometry, bars As CRebarLayout
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
    Dim geom As ISectionGeometry, bars As CRebarLayout
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
    Dim originalGeom As ISectionGeometry, originalBars As CRebarLayout
    Dim geom As ISectionGeometry, bars As CRebarLayout, key As Variant, index As Long, equal As Boolean
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
Private Sub ReadGeometry(ByRef geom As ISectionGeometry, ByRef bars As CRebarLayout)
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
