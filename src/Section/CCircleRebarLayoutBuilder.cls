VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCircleRebarLayoutBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Автоматическая раскладка арматуры круглого сечения
' ==========================================================================
' Builder расставляет первый, второй и третий ряды стержней по окружности:
' Stacked уводит дополнительный ряд к центру, SideBySide смещает по касательной
' по часовой стрелке. Если первого ряда нет, дополнительные ряды не создаются.

Private mAngleStep As Double
Private mCentroidX As Double
Private mCentroidY As Double
Private mAxisRadius As Double
Private Const REBAR_STEEL_CLASS As String = "Rebar" ' Техническая метка типа элемента: расчетная арматура, не коэффициент трещин.

' Строит круговую раскладку по таблице rngCircleGeometry.
' Пустой или нулевой Rebar.Diameter означает, что первый ряд отсутствует;
' дополнительные ряды без первого ряда не создаются, потому что их положение
' задается относительно стержней основного ряда.
Public Function BuildFromSettings(ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem, _
        Optional ByVal centerX As Double = 0#, Optional ByVal centerY As Double = 0#) As CRebarLayout
    If settings Is Nothing Then Err.Raise vbObjectError + 3162, "CCircleRebarLayoutBuilder", "Настройки Config не переданы."
    If units Is Nothing Then Err.Raise vbObjectError + 3163, "CCircleRebarLayoutBuilder", "Система единиц не передана."

    Set BuildFromSettings = Build( _
        units.InputLengthToInternal(settings.GetDouble("Circle.Diameter", 300#)), _
        centerX, _
        centerY, _
        units.InputLengthToInternal(settings.GetDouble("Rebar.AxisDistance", 40#)), _
        settings.GetLong("Rebar.Count", 8), _
        OptionalDiameterFromSettings(settings, units, "Rebar.Diameter", 20#), _
        REBAR_STEEL_CLASS, _
        OptionalDiameterFromSettings(settings, units, "Rebar.Diameter2", 0#), _
        OptionalDiameterFromSettings(settings, units, "Rebar.Diameter3", 0#), _
        settings.GetString("Rebar.Loc2row", "Stacked"), _
        settings.GetString("Rebar.Loc3row", "Stacked"))
End Function

' Возвращает диаметр из Config без подстановки default-а для пустой ячейки.
' Такое правило важно именно для арматуры: пустое значение означает "не
' создавать ряд", а не "взять типовой диаметр".
Private Function OptionalDiameterFromSettings(ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem, _
        ByVal key As String, ByVal defaultValue As Double) As Double
    If Not settings.HasKey(key) Then
        OptionalDiameterFromSettings = units.InputLengthToInternal(defaultValue)
        Exit Function
    End If

    Dim rawValue As String
    rawValue = Trim$(settings.GetRawString(key, vbNullString))
    If Len(rawValue) = 0 Or UCase$(rawValue) = "TODO" Then
        OptionalDiameterFromSettings = 0#
    Else
        OptionalDiameterFromSettings = units.InputLengthToInternal(settings.GetDouble(key, defaultValue))
    End If
End Function

' Создает раскладку из уже нормализованных чисел.
' Метод не читает Excel и поэтому удобен для расчетных тестов и внутренних
' сценариев, где параметры круга подготовлены заранее.
Public Function Build(ByVal diameter As Double, ByVal centerX As Double, ByVal centerY As Double, _
        ByVal axisDistance As Double, ByVal barCount As Long, ByVal barDiameter As Double, _
        ByVal steelClass As String, Optional ByVal row2Diameter As Double = 0#, _
        Optional ByVal row3Diameter As Double = 0#, Optional ByVal row2Location As String = "Stacked", _
        Optional ByVal row3Location As String = "Stacked") As CRebarLayout

    ValidateInputs diameter, axisDistance, barCount, barDiameter, row2Diameter, row3Diameter, row2Location, row3Location

    Dim geom As CGeometryCircle
    Set geom = New CGeometryCircle
    geom.InitializeByDiameter diameter, centerX, centerY

    Dim layout As CRebarLayout
    Set layout = New CRebarLayout

    mCentroidX = centerX
    mCentroidY = centerY
    mAxisRadius = 0#
    mAngleStep = 0#

    If barCount = 0 Or barDiameter = 0# Then
        Set Build = layout
        Exit Function
    End If

    mAxisRadius = diameter / 2# - axisDistance
    mAngleStep = 2# * GEOM_PI / CDbl(barCount)
    layout.AddAnnotationAnchor "Circle", centerX - mAxisRadius, centerY + mAxisRadius, _
        centerX + mAxisRadius, centerY + mAxisRadius, 0#, 1#, axisDistance

    Dim i As Long
    Dim alpha As Double
    Dim xCoord As Double
    Dim yCoord As Double
    Dim sx As Double
    Dim sy As Double
    For i = 0 To barCount - 1
        alpha = mAngleStep * CDbl(i)
        xCoord = centerX + mAxisRadius * Cos(alpha)
        yCoord = centerY + mAxisRadius * Sin(alpha)
        layout.AddBar "B" & CStr(i + 1), xCoord, yCoord, barDiameter, 0#, steelClass, _
            "Круг: основной ряд арматуры", geom, "Circle"
        AddAdditionalCircleBar layout, geom, "B" & CStr(i + 1) & "_R2", centerX, centerY, xCoord, yCoord, _
            alpha, barDiameter, row2Diameter, row2Location, 0#, steelClass
        AddAdditionalCircleBar layout, geom, "B" & CStr(i + 1) & "_R3", centerX, centerY, xCoord, yCoord, _
            alpha, barDiameter, row3Diameter, row3Location, RowSkipDiameter(row2Diameter, row2Location, row3Location), steelClass
        sx = sx + xCoord
        sy = sy + yCoord
    Next i

    Dim expectedCount As Long
    expectedCount = barCount
    If row2Diameter > 0# Then expectedCount = expectedCount + barCount
    If row3Diameter > 0# Then expectedCount = expectedCount + barCount
    If layout.Count <> expectedCount Then Err.Raise vbObjectError + 3160, "CCircleRebarLayoutBuilder", "Количество созданных стержней не совпадает с заданным."

    mCentroidX = sx / CDbl(barCount)
    mCentroidY = sy / CDbl(barCount)
    If Abs(mCentroidX - centerX) > 0.000001 Or Abs(mCentroidY - centerY) > 0.000001 Then
        Err.Raise vbObjectError + 3161, "CCircleRebarLayoutBuilder", "Центр тяжести первого ряда арматуры не совпадает с центром круга."
    End If

    Set Build = layout
End Function

Private Sub AddAdditionalCircleBar(ByVal layout As CRebarLayout, ByVal geom As CGeometryCircle, _
        ByVal barName As String, ByVal centerX As Double, ByVal centerY As Double, _
        ByVal baseX As Double, ByVal baseY As Double, ByVal alpha As Double, _
        ByVal baseDiameter As Double, ByVal addDiameter As Double, ByVal location As String, _
        ByVal skipDiameter As Double, ByVal steelClass As String)
    If addDiameter <= 0# Then Exit Sub

    Dim offsetDistance As Double
    offsetDistance = baseDiameter / 2# + skipDiameter + addDiameter / 2#

    Dim dirX As Double
    Dim dirY As Double
    If IsStackedLocation(location) Then
        dirX = -Cos(alpha)
        dirY = -Sin(alpha)
    Else
        dirX = Sin(alpha)
        dirY = -Cos(alpha)
    End If

    Dim xCoord As Double
    Dim yCoord As Double
    xCoord = baseX + offsetDistance * dirX
    yCoord = baseY + offsetDistance * dirY
    ValidateBarInsideCircle geom, centerX, centerY, xCoord, yCoord, addDiameter
    layout.AddBar barName, xCoord, yCoord, addDiameter, 0#, steelClass, _
        "Круг: дополнительный ряд арматуры", geom, "Circle"
End Sub

Private Function RowSkipDiameter(ByVal row2Diameter As Double, ByVal row2Location As String, ByVal row3Location As String) As Double
    If row2Diameter > 0# And StrComp(Trim$(row2Location), Trim$(row3Location), vbTextCompare) = 0 Then
        RowSkipDiameter = row2Diameter
    Else
        RowSkipDiameter = 0#
    End If
End Function

' Проверяет входные данные и прерывает выполнение понятной ошибкой, если расчетный контракт нарушен.
Private Sub ValidateInputs(ByVal diameter As Double, ByVal axisDistance As Double, _
        ByVal barCount As Long, ByVal barDiameter As Double, ByVal row2Diameter As Double, _
        ByVal row3Diameter As Double, ByVal row2Location As String, ByVal row3Location As String)
    If diameter <= 0# Then Err.Raise vbObjectError + 3150, "CCircleRebarLayoutBuilder", "Диаметр круга должен быть положительным."
    If barCount < 0 Then Err.Raise vbObjectError + 3151, "CCircleRebarLayoutBuilder", "Количество стержней первого ряда не может быть отрицательным."
    If barDiameter < 0# Then Err.Raise vbObjectError + 3152, "CCircleRebarLayoutBuilder", "Диаметр арматуры первого ряда не может быть отрицательным."
    If row2Diameter < 0# Then Err.Raise vbObjectError + 3155, "CCircleRebarLayoutBuilder", "Диаметр второго ряда не может быть отрицательным."
    If row3Diameter < 0# Then Err.Raise vbObjectError + 3156, "CCircleRebarLayoutBuilder", "Диаметр третьего ряда не может быть отрицательным."

    If barCount = 0 Or barDiameter = 0# Then Exit Sub

    If barCount < 2 Then Err.Raise vbObjectError + 3151, "CCircleRebarLayoutBuilder", "Количество стержней первого ряда должно быть не менее 2 или равно 0."
    If row2Diameter > 0# And Not IsSupportedRowLocation(row2Location) Then Err.Raise vbObjectError + 3157, "CCircleRebarLayoutBuilder", "Rebar.Loc2row должен быть SideBySide или Stacked."
    If row3Diameter > 0# And Not IsSupportedRowLocation(row3Location) Then Err.Raise vbObjectError + 3158, "CCircleRebarLayoutBuilder", "Rebar.Loc3row должен быть SideBySide или Stacked."
    If axisDistance <= barDiameter / 2# Then Err.Raise vbObjectError + 3153, "CCircleRebarLayoutBuilder", "Расстояние до оси стержня должно быть больше половины диаметра арматуры."
    If axisDistance >= diameter / 2# Then Err.Raise vbObjectError + 3154, "CCircleRebarLayoutBuilder", "Расстояние до оси стержня должно быть меньше половины диаметра круга."
End Sub

' Проверяет входные данные и прерывает выполнение понятной ошибкой, если расчетный контракт нарушен.
Private Sub ValidateBarInsideCircle(ByVal geom As CGeometryCircle, ByVal centerX As Double, ByVal centerY As Double, _
        ByVal xCoord As Double, ByVal yCoord As Double, ByVal barDiameter As Double)
    Dim dx As Double
    Dim dy As Double
    dx = xCoord - centerX
    dy = yCoord - centerY
    If Sqr(dx * dx + dy * dy) + barDiameter / 2# > geom.Radius + GEOM_TOLERANCE Then
        Err.Raise vbObjectError + 3159, "CCircleRebarLayoutBuilder", "Стержень дополнительного ряда не помещается внутри круглого сечения."
    End If
End Sub

Private Function IsSupportedRowLocation(ByVal value As String) As Boolean
    IsSupportedRowLocation = IsStackedLocation(value) Or StrComp(Trim$(value), "SideBySide", vbTextCompare) = 0
End Function

Private Function IsStackedLocation(ByVal value As String) As Boolean
    IsStackedLocation = (StrComp(Trim$(value), "Stacked", vbTextCompare) = 0)
End Function

Public Property Get CentroidX() As Double
    CentroidX = mCentroidX
End Property

Public Property Get CentroidY() As Double
    CentroidY = mCentroidY
End Property

' ==========================================================================
' ДЛЯ ТЕСТОВ
' ==========================================================================
' Эти диагностические свойства нужны проверкам круговой раскладки: тесты
' сверяют угловой шаг и расчетный радиус осей стержней с ожидаемой геометрией.

Public Property Get AngleStep() As Double
    AngleStep = mAngleStep
End Property

Public Property Get AxisRadius() As Double
    AxisRadius = mAxisRadius
End Property

