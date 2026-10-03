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
        units.InputLengthToInternal(settings.GetRequiredDouble("Circle.Diameter")), _
        centerX, _
        centerY, _
        units.InputLengthToInternal(settings.GetDouble("Rebar.AxisDistance", 40#)), _
        settings.GetLong("Rebar.Count", 0), _
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

    ' Сбрасываем диагностические величины последнего построения. Их читают
    ' тесты и справочная логика, но на расчетную модель они не влияют.
    mCentroidX = centerX
    mCentroidY = centerY
    mAxisRadius = 0#
    mAngleStep = 0#

    If barCount = 0 Or barDiameter = 0# Then
        Set Build = layout
        Exit Function
    End If

    ' Суммируем активные ряды в Double до построения. Счетчики раскладки
    ' и обходы имеют тип Long; оставляем представимым также шаг Next.
    Dim expectedCountValue As Double
    expectedCountValue = CDbl(barCount)
    If row2Diameter > 0# Then expectedCountValue = expectedCountValue + CDbl(barCount)
    If row3Diameter > 0# Then expectedCountValue = expectedCountValue + CDbl(barCount)
    If expectedCountValue > 2147483646# Then
        Err.Raise vbObjectError + 3164, "CCircleRebarLayoutBuilder", _
            "На листе Config в блоке круглого сечения уменьшите Rebar.Count или число активных рядов: суммарное количество стержней превышает диапазон счетчиков программы."
    End If
    Dim expectedCount As Long
    expectedCount = CLng(expectedCountValue)

    mAxisRadius = diameter / 2# - axisDistance
    mAngleStep = 2# * GEOM_PI / CDbl(barCount)

    ' Anchor подписи арматуры идет по хорде первого ряда, чтобы схема могла
    ' вывести одну групповую надпись вместо подписи каждого стержня.
    layout.AddAnnotationAnchor "Circle", centerX - mAxisRadius, centerY + mAxisRadius, _
        centerX + mAxisRadius, centerY + mAxisRadius, 0#, 1#, axisDistance

    Dim i As Long
    Dim alpha As Double ' Угол текущего стержня первого ряда от глобальной +X.
    Dim xCoord As Double ' Координата оси текущего стержня.
    Dim yCoord As Double
    Dim sx As Double ' Суммы координат нужны только для контрольной проверки центра первого ряда.
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

    If layout.Count <> expectedCount Then Err.Raise vbObjectError + 3160, "CCircleRebarLayoutBuilder", "Количество созданных стержней не совпадает с заданным."

    mCentroidX = sx / CDbl(barCount)
    mCentroidY = sy / CDbl(barCount)
    If Abs(mCentroidX - centerX) > 0.000001 Or Abs(mCentroidY - centerY) > 0.000001 Then
        Err.Raise vbObjectError + 3161, "CCircleRebarLayoutBuilder", "Центр тяжести первого ряда арматуры не совпадает с центром круга."
    End If

    Set Build = layout
End Function

' Добавляет второй или третий ряд относительно одного стержня первого ряда.
' Stacked смещает стержень к центру круга, SideBySide - по касательной по
' часовой стрелке. Перед записью проверяется, что новый круг целиком остается
' внутри бетонного контура.
Private Sub AddAdditionalCircleBar(ByVal layout As CRebarLayout, ByVal geom As CGeometryCircle, _
        ByVal barName As String, ByVal centerX As Double, ByVal centerY As Double, _
        ByVal baseX As Double, ByVal baseY As Double, ByVal alpha As Double, _
        ByVal baseDiameter As Double, ByVal addDiameter As Double, ByVal location As String, _
        ByVal skipDiameter As Double, ByVal steelClass As String)
    If addDiameter <= 0# Then Exit Sub

    Dim offsetDistance As Double ' Расстояние между осями соседних стержней с учетом возможного пропущенного ряда.
    offsetDistance = baseDiameter / 2# + skipDiameter + addDiameter / 2#

    Dim dirX As Double ' Единичное направление смещения дополнительного стержня.
    Dim dirY As Double
    If IsStackedLocation(location) Then
        ' Радиальное направление внутрь: от стержня первого ряда к центру.
        dirX = -Cos(alpha)
        dirY = -Sin(alpha)
    Else
        ' Касательная по часовой стрелке в локальной точке окружности.
        dirX = Sin(alpha)
        dirY = -Cos(alpha)
    End If

    Dim xCoord As Double ' Координаты оси добавляемого стержня.
    Dim yCoord As Double
    xCoord = baseX + offsetDistance * dirX
    yCoord = baseY + offsetDistance * dirY
    ValidateBarInsideCircle geom, centerX, centerY, xCoord, yCoord, addDiameter
    layout.AddBar barName, xCoord, yCoord, addDiameter, 0#, steelClass, _
        "Круг: дополнительный ряд арматуры", geom, "Circle"
End Sub

' Возвращает диаметр промежуточного ряда, который надо "перепрыгнуть" третьему
' ряду. Это нужно, когда второй и третий ряды имеют одинаковую схему смещения:
' третий ряд ставится за вторым, а не накладывается на него.
Private Function RowSkipDiameter(ByVal row2Diameter As Double, ByVal row2Location As String, ByVal row3Location As String) As Double
    If row2Diameter > 0# And StrComp(Trim$(row2Location), Trim$(row3Location), vbTextCompare) = 0 Then
        RowSkipDiameter = row2Diameter
    Else
        RowSkipDiameter = 0#
    End If
End Function

' Проверяет размеры круга, число стержней и положение активных рядов до
' построения координат. Нулевой первый ряд выключает зависимые ряды;
' активный ряд должен иметь защитное расстояние и помещаться по радиусу.
Private Sub ValidateInputs(ByVal diameter As Double, ByVal axisDistance As Double, _
        ByVal barCount As Long, ByVal barDiameter As Double, ByVal row2Diameter As Double, _
        ByVal row3Diameter As Double, ByVal row2Location As String, ByVal row3Location As String)
    If diameter <= 0# Then Err.Raise vbObjectError + 3150, "CCircleRebarLayoutBuilder", "Circle.Diameter должен быть положительным."
    If barCount < 0 Then Err.Raise vbObjectError + 3151, "CCircleRebarLayoutBuilder", "Rebar.Count не может быть отрицательным."
    If barDiameter < 0# Then Err.Raise vbObjectError + 3152, "CCircleRebarLayoutBuilder", "Rebar.Diameter не может быть отрицательным."
    If row2Diameter < 0# Then Err.Raise vbObjectError + 3155, "CCircleRebarLayoutBuilder", "Rebar.Diameter2 не может быть отрицательным."
    If row3Diameter < 0# Then Err.Raise vbObjectError + 3156, "CCircleRebarLayoutBuilder", "Rebar.Diameter3 не может быть отрицательным."

    If barCount = 0 Or barDiameter = 0# Then Exit Sub

    If barCount < 2 Then Err.Raise vbObjectError + 3151, "CCircleRebarLayoutBuilder", "Rebar.Count должен быть равен 0 или быть не меньше 2."
    If row2Diameter > 0# And Not IsSupportedRowLocation(row2Location) Then Err.Raise vbObjectError + 3157, "CCircleRebarLayoutBuilder", "Rebar.Loc2row должен быть SideBySide или Stacked."
    If row3Diameter > 0# And Not IsSupportedRowLocation(row3Location) Then Err.Raise vbObjectError + 3158, "CCircleRebarLayoutBuilder", "Rebar.Loc3row должен быть SideBySide или Stacked."
    If axisDistance <= barDiameter / 2# Then Err.Raise vbObjectError + 3153, "CCircleRebarLayoutBuilder", "Rebar.AxisDistance должен быть больше Rebar.Diameter/2."
    If axisDistance >= diameter / 2# Then Err.Raise vbObjectError + 3154, "CCircleRebarLayoutBuilder", "Rebar.AxisDistance должен быть меньше Circle.Diameter/2."
End Sub

' Проверяет размещение всего круглого стержня, а не только его центра.
' Выход дополнительного ряда за наружную окружность отклоняется с причиной
' до добавления стержня в раскладку.
Private Sub ValidateBarInsideCircle(ByVal geom As CGeometryCircle, ByVal centerX As Double, ByVal centerY As Double, _
        ByVal xCoord As Double, ByVal yCoord As Double, ByVal barDiameter As Double)
    Dim dx As Double
    Dim dy As Double
    dx = xCoord - centerX
    dy = yCoord - centerY
    If Sqr(dx * dx + dy * dy) + barDiameter / 2# > geom.Radius + GEOM_TOLERANCE Then
        Err.Raise vbObjectError + 3159, "CCircleRebarLayoutBuilder", "Стержень дополнительного ряда Rebar.Diameter2/3 не помещается внутри круглого сечения."
    End If
End Sub

' Проверяет, что текстовый режим дополнительного ряда относится к известным
' пользовательским вариантам. Проверка отделена от Build, чтобы одинаково
' обслуживать второй и третий ряд.
Private Function IsSupportedRowLocation(ByVal value As String) As Boolean
    IsSupportedRowLocation = IsStackedLocation(value) Or StrComp(Trim$(value), "SideBySide", vbTextCompare) = 0
End Function

' Возвращает True для режима Stacked; отдельная функция нужна, чтобы не
' размазывать регистронезависимое сравнение по нескольким расчетным веткам.
Private Function IsStackedLocation(ByVal value As String) As Boolean
    IsStackedLocation = (StrComp(Trim$(value), "Stacked", vbTextCompare) = 0)
End Function

' Центр тяжести первого ряда арматуры после последнего Build. Используется
' тестами как контроль равномерности круговой раскладки.
Public Property Get CentroidX() As Double
    CentroidX = mCentroidX
End Property

' Y-координата центра первого ряда арматуры после последнего Build.
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

' Радиус линии осей первого ряда после последнего построения круговой раскладки.
Public Property Get AxisRadius() As Double
    AxisRadius = mAxisRadius
End Property

