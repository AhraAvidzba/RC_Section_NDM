VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CRoundedRectRebarLayoutBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Автоматическая раскладка арматуры RoundedRectangle
' ==========================================================================
' Builder читает таблицу скругленного прямоугольника: общие размеры B/H,
' левую/правую боковую геометрию Simple/Tapered и четыре смысловые линии
' арматуры H-left/H-right/B-top/B-bottom. На выходе создается обычный
' CRebarLayout; расчетное ядро дальше работает с ним так же, как с Circle или
' RectSet.

Private Const REBAR_STEEL_CLASS As String = "Rebar" ' Техническая метка материала арматуры в generated-модели.

Private mRequestedCount As Long
Private mSkippedDuplicateCount As Long

' Строит раскладку по настройкам Config, не обращаясь к Excel напрямую.
Public Function BuildFromSettings(ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem) As CRebarLayout
    If settings Is Nothing Then Err.Raise vbObjectError + 3220, "CRoundedRectRebarLayoutBuilder", "Настройки Config не переданы."
    If units Is Nothing Then Err.Raise vbObjectError + 3221, "CRoundedRectRebarLayoutBuilder", "Система единиц не передана."

    Dim geom As CGeometryRoundedRectangle
    Set geom = New CGeometryRoundedRectangle
    Dim leftType As String, rightType As String, leftW As Double, rightW As Double, leftR2 As Double, rightR2 As Double
    leftType = settings.GetRequiredChoice("RoundedRectangle.Left.Type", Array("Simple", "Tapered"))
    rightType = settings.GetRequiredChoice("RoundedRectangle.Right.Type", Array("Simple", "Tapered"))
    ' Для Simple поля W/R2 не участвуют в форме и не должны требовать числа.
    If leftType = "Tapered" Then
        leftW = units.InputLengthToInternal(settings.GetRequiredDouble("RoundedRectangle.Left.W"))
        leftR2 = units.InputLengthToInternal(settings.GetRequiredDouble("RoundedRectangle.Left.R2"))
    End If
    If rightType = "Tapered" Then
        rightW = units.InputLengthToInternal(settings.GetRequiredDouble("RoundedRectangle.Right.W"))
        rightR2 = units.InputLengthToInternal(settings.GetRequiredDouble("RoundedRectangle.Right.R2"))
    End If
    geom.InitializeSides _
        units.InputLengthToInternal(settings.GetRequiredDouble("RoundedRectangle.B")), _
        units.InputLengthToInternal(settings.GetRequiredDouble("RoundedRectangle.H")), _
        leftType, rightType, leftW, rightW, _
        units.InputLengthToInternal(settings.GetRequiredDouble("RoundedRectangle.Left.R1")), _
        units.InputLengthToInternal(settings.GetRequiredDouble("RoundedRectangle.Right.R1")), leftR2, rightR2

    Set BuildFromSettings = Build(geom, _
        ReadFaceSettingsFromConfig(settings, units, "H.Left"), _
        ReadFaceSettingsFromConfig(settings, units, "H.Right"), _
        ReadFaceSettingsFromConfig(settings, units, "B.Top"), _
        ReadFaceSettingsFromConfig(settings, units, "B.Bottom"), _
        REBAR_STEEL_CLASS)
End Function

' Создает раскладку для уже подготовленной геометрии. Этот метод нужен тестам
' и сохраняет shape-specific математику внутри builder-а.
Public Function Build(ByVal geom As CGeometryRoundedRectangle, _
        ByVal hLeftSettings As Variant, ByVal hRightSettings As Variant, _
        ByVal bTopSettings As Variant, ByVal bBottomSettings As Variant, _
        ByVal steelClass As String) As CRebarLayout
    If geom Is Nothing Then Err.Raise vbObjectError + 3222, "CRoundedRectRebarLayoutBuilder", "Геометрия RoundedRectangle не передана."

    Dim message As String
    If Not geom.IsValid(message) Then Err.Raise vbObjectError + 3223, "CRoundedRectRebarLayoutBuilder", message

    ' Сумма в Double проверяется до Long-выражения n+1, циклов и размещения.
    Dim requestedPositions As Double
    requestedPositions = FaceRequestedPositions(hLeftSettings) + FaceRequestedPositions(hRightSettings) + _
        FaceRequestedPositions(bTopSettings) + FaceRequestedPositions(bBottomSettings)
    If requestedPositions > GEOM_MAX_REBAR_COUNT Then Err.Raise vbObjectError + 3230, "CRoundedRectRebarLayoutBuilder", _
        "На листе Config в блоке скругленного прямоугольника уменьшите количество стержней n или число активных рядов: суммарное количество запрошенных позиций превышает диапазон счетчиков программы."

    Dim layout As CRebarLayout
    Set layout = New CRebarLayout
    mRequestedCount = 0
    mSkippedDuplicateCount = 0

    Dim nextIndex As Long
    nextIndex = 1
    AddHorizontalFace layout, geom, "B.Top", True, bTopSettings, steelClass, nextIndex
    AddHorizontalFace layout, geom, "B.Bottom", False, bBottomSettings, steelClass, nextIndex
    AddSideFace layout, geom, "H.Left", True, hLeftSettings, steelClass, nextIndex
    AddSideFace layout, geom, "H.Right", False, hRightSettings, steelClass, nextIndex

    Set Build = layout
End Function

' Оценивает запрошенные позиции одной активной грани без координат и дублей;
' нулевой первый ряд выключает также зависящие от него дополнительные ряды.
Private Function FaceRequestedPositions(ByVal settings As Variant) As Double
    FaceRequestedPositions = RebarRequestedPositionCount(FaceCount(settings), FaceDiameter(settings), _
        FaceRow2Diameter(settings), FaceRow3Diameter(settings), FaceRow2Binding(settings), FaceRow3Binding(settings))
End Function

' Читает только параметры активной грани и активных дополнительных рядов.
' Нулевой/пустой n или d выключает линию; as и selectors активного ряда
' обязательны. Ошибки ввода содержат фактический адрес, а не типовой default.
Private Function ReadFaceSettingsFromConfig(ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem, _
        ByVal faceName As String) As Variant
    Dim prefix As String, countKey As String, count As Long, diameter As Double, axisDistance As Double
    Dim row2 As Double, row3 As Double, loc2 As String, loc3 As String, bind2 As String, bind3 As String
    prefix = "RoundedRectangle." & faceName & ".": countKey = prefix & "n"
    If Not settings.HasKey(countKey) Then Err.Raise vbObjectError + 4310, "CRoundedRectRebarLayoutBuilder", _
        settings.InputErrorMessage(countKey, "Строка количества стержней отсутствует.", "Восстановите строку; введите целое количество или 0 для выключенной грани.")
    count = settings.GetLong(countKey, 0)
    If count < 0 Then Err.Raise vbObjectError + 3230, "CRoundedRectRebarLayoutBuilder", _
        settings.InputErrorMessage(countKey, RoundedFaceDisplayName(faceName) & ": количество не может быть отрицательным.", "Введите целое неотрицательное количество стержней.")
    diameter = OptionalDiameterFromSettings(settings, units, prefix & "d", 3231)
    loc2 = "Stacked": loc3 = "Stacked": bind2 = "EachBar": bind3 = "EachBar"
    If count > 0 And diameter > 0# Then
        axisDistance = units.InputLengthToInternal(settings.GetRequiredDouble(prefix & "as"))
        If axisDistance <= diameter / 2# Then Err.Raise vbObjectError + 3232, "CRoundedRectRebarLayoutBuilder", _
            settings.InputErrorMessage(prefix & "as", RoundedFaceDisplayName(faceName) & ": расстояние до оси стержня должно быть больше d/2.", "Введите расстояние as больше половины диаметра первого ряда.")
        row2 = OptionalDiameterFromSettings(settings, units, prefix & "d_2row", 3233)
        row3 = OptionalDiameterFromSettings(settings, units, prefix & "d_3row", 3233)
        If row2 > 0# Then
            loc2 = settings.GetRequiredChoice(prefix & "loc_2row", Array("Stacked", "SideBySide"))
            bind2 = settings.GetRequiredChoice(prefix & "bind_2row", Array("EachBar", "EverySecondBar"))
        End If
        If row3 > 0# Then
            loc3 = settings.GetRequiredChoice(prefix & "loc_3row", Array("Stacked", "SideBySide"))
            bind3 = settings.GetRequiredChoice(prefix & "bind_3row", Array("EachBar", "EverySecondBar"))
        End If
    End If
    ReadFaceSettingsFromConfig = Array(axisDistance, diameter, count, row2, row3, loc2, loc3, bind2, bind3)
End Function

' Читает существующую ячейку диаметра: пустота/TODO означают отсутствие ряда.
' Потерянный ключ и отрицательное число дают адресную ошибку ввода; errorCode
' сохраняет инженерный код основной либо дополнительной арматуры этой формы.
Private Function OptionalDiameterFromSettings(ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem, _
        ByVal key As String, ByVal errorCode As Long) As Double
    If Not settings.HasKey(key) Then Err.Raise vbObjectError + 4310, "CRoundedRectRebarLayoutBuilder", _
        settings.InputErrorMessage(key, "Строка диаметра отсутствует.", "Восстановите строку; введите диаметр или оставьте ячейку пустой для отсутствующего ряда.")
    Dim rawValue As String
    rawValue = Trim$(settings.GetRawString(key, vbNullString))
    If Len(rawValue) = 0 Or UCase$(rawValue) = "TODO" Then Exit Function
    Dim value As Double
    value = settings.GetRequiredDouble(key)
    If value < 0# Then Err.Raise vbObjectError + errorCode, "CRoundedRectRebarLayoutBuilder", _
        settings.InputErrorMessage(key, "Диаметр стержней не может быть отрицательным.", "Введите неотрицательный диаметр; 0 или пустая ячейка выключают ряд.")
    OptionalDiameterFromSettings = units.InputLengthToInternal(value)
End Function

' Добавляет один ряд арматуры вдоль верхней или нижней смысловой грани B.
' Геометрия возвращает уже смещенную внутрь линию между касательными точками
' соседних скруглений R1. В отличие от H-граней, B-ряд занимает весь этот
' прямой участок: первый и последний стержни ставятся в начальную и конечную
' точки линии, а промежуточные равномерно делят оставшуюся длину.
Private Sub AddHorizontalFace(ByVal layout As CRebarLayout, ByVal geom As CGeometryRoundedRectangle, _
        ByVal faceName As String, ByVal isTop As Boolean, ByVal settings As Variant, _
        ByVal steelClass As String, ByRef nextIndex As Long)
    Dim barCount As Long
    Dim diameter As Double
    Dim axisDistance As Double
    axisDistance = FaceAs(settings)
    diameter = FaceDiameter(settings)
    barCount = FaceCount(settings)
    ValidateFace faceName, axisDistance, diameter, barCount, FaceRow2Diameter(settings), FaceRow3Diameter(settings), _
        FaceRow2Location(settings), FaceRow3Location(settings), FaceRow2Binding(settings), FaceRow3Binding(settings)
    If barCount = 0 Or diameter = 0# Then Exit Sub
    mRequestedCount = mRequestedCount + barCount

    Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
    Dim normalX As Double, normalY As Double
    geom.GetHorizontalRebarLine isTop, axisDistance, x1, y1, x2, y2, normalX, normalY

    Dim lengthValue As Double
    lengthValue = Sqr((x2 - x1) ^ 2 + (y2 - y1) ^ 2)
    If lengthValue <= GEOM_TOLERANCE Then Err.Raise vbObjectError + 3224, "CRoundedRectRebarLayoutBuilder", _
        RoundedFaceDisplayName(faceName) & ": длина линии раскладки должна быть больше нуля."

    layout.AddAnnotationAnchor faceName, x1, y1, x2, y2, normalX, normalY, axisDistance

    Dim tx As Double, ty As Double
    tx = (x2 - x1) / lengthValue
    ty = (y2 - y1) / lengthValue
    Dim inwardX As Double, inwardY As Double
    inwardX = -normalX
    inwardY = -normalY

    Dim i As Long
    For i = 1 To barCount
        Dim t As Double
        If barCount = 1 Then
            t = 0.5
        Else
            t = CDbl(i - 1) / CDbl(barCount - 1)
        End If
        Dim xCoord As Double, yCoord As Double
        xCoord = x1 + t * (x2 - x1)
        yCoord = y1 + t * (y2 - y1)
        AddBarIfAvailable layout, geom, faceName, 1, xCoord, yCoord, diameter, steelClass, nextIndex
        AddAdditionalRows layout, geom, faceName, xCoord, yCoord, diameter, settings, i, tx, ty, inwardX, inwardY, steelClass, nextIndex
    Next i
End Sub

' Добавляет арматуру вдоль левой или правой смысловой грани H.
' Путь H идет по внутренней линии, параллельной соответствующей боковой части
' контура: для Simple это вертикальная вставка со скругленными концами, для
' Tapered - верхнее скругление R1, наклонная часть, носовой радиус R2,
' нижняя наклонная часть и нижнее скругление R1. Начало и конец пути совпадают
' с точками, где горизонтальные прямые B переходят в свои скругления.
'
' Стержни H не ставятся непосредственно в эти конечные точки: линия длиной L
' делится на n+1 равных интервалов, поэтому первый и последний стержни имеют
' одинаковый отступ от начала и конца, равный принятому шагу L/(n+1).
Private Sub AddSideFace(ByVal layout As CRebarLayout, ByVal geom As CGeometryRoundedRectangle, _
        ByVal faceName As String, ByVal isLeft As Boolean, ByVal settings As Variant, _
        ByVal steelClass As String, ByRef nextIndex As Long)
    Dim barCount As Long
    Dim diameter As Double
    Dim axisDistance As Double
    axisDistance = FaceAs(settings)
    diameter = FaceDiameter(settings)
    barCount = FaceCount(settings)
    ValidateFace faceName, axisDistance, diameter, barCount, FaceRow2Diameter(settings), FaceRow3Diameter(settings), _
        FaceRow2Location(settings), FaceRow3Location(settings), FaceRow2Binding(settings), FaceRow3Binding(settings)
    If barCount = 0 Or diameter = 0# Then Exit Sub
    mRequestedCount = mRequestedCount + barCount

    Dim px() As Double
    Dim py() As Double
    geom.GetSideRebarPath isLeft, axisDistance, px, py

    Dim totalLength As Double
    totalLength = PolylineLength(px, py)
    If totalLength <= GEOM_TOLERANCE Then Err.Raise vbObjectError + 3225, "CRoundedRectRebarLayoutBuilder", _
        RoundedFaceDisplayName(faceName) & ": длина боковой линии раскладки должна быть больше нуля."

    Dim outwardX As Double, outwardY As Double
    If isLeft Then outwardX = -1# Else outwardX = 1#

    Dim anchorX As Double
    Dim anchorMinY As Double
    Dim anchorMaxY As Double
    Dim anchorInitialized As Boolean
    Dim spacing As Double
    spacing = totalLength / CDbl(barCount + 1)

    Dim i As Long
    For i = 1 To barCount
        Dim distanceAlong As Double
        distanceAlong = spacing * CDbl(i)
        Dim xCoord As Double, yCoord As Double, tx As Double, ty As Double
        PointAtDistance px, py, distanceAlong, xCoord, yCoord, tx, ty

        AccumulateSideLabelAnchor isLeft, xCoord, yCoord, anchorX, anchorMinY, anchorMaxY, anchorInitialized

        Dim inwardX As Double, inwardY As Double
        If isLeft Then
            inwardX = -ty: inwardY = tx
        Else
            inwardX = ty: inwardY = -tx
        End If
        NormalizeVector inwardX, inwardY

        AddBarIfAvailable layout, geom, faceName, 1, xCoord, yCoord, diameter, steelClass, nextIndex
        AddAdditionalRows layout, geom, faceName, xCoord, yCoord, diameter, settings, i, tx, ty, inwardX, inwardY, steelClass, nextIndex
    Next i
    If anchorInitialized Then layout.AddAnnotationAnchor faceName, anchorX, anchorMaxY, anchorX, anchorMinY, outwardX, outwardY, axisDistance
End Sub

' Собирает вертикальный anchor подписи по фактическим стержням боковой грани.
' Для криволинейных и tapered-граней это важнее, чем хорда между первым и
' последним стержнем: видимый отступ должен считаться от крайней арматуры.
Private Sub AccumulateSideLabelAnchor(ByVal isLeft As Boolean, ByVal xCoord As Double, ByVal yCoord As Double, _
        ByRef anchorX As Double, ByRef anchorMinY As Double, ByRef anchorMaxY As Double, ByRef initialized As Boolean)
    If Not initialized Then
        anchorX = xCoord
        anchorMinY = yCoord
        anchorMaxY = yCoord
        initialized = True
    Else
        If isLeft Then
            If xCoord < anchorX Then anchorX = xCoord
        Else
            If xCoord > anchorX Then anchorX = xCoord
        End If
        If yCoord < anchorMinY Then anchorMinY = yCoord
        If yCoord > anchorMaxY Then anchorMaxY = yCoord
    End If
End Sub

' Добавляет второй и третий ряды для уже найденного стержня первого ряда.
' Stacked смещает стержень внутрь сечения по нормали к текущей линии
' раскладки, SideBySide - вдоль касательной к этой линии. Для криволинейной
' H-грани касательная и нормаль берутся в фактической точке первого ряда,
' поэтому дополнительные ряды следуют форме боковой траектории.
Private Sub AddAdditionalRows(ByVal layout As CRebarLayout, ByVal geom As CGeometryRoundedRectangle, _
        ByVal faceName As String, ByVal baseX As Double, ByVal baseY As Double, ByVal row1Diameter As Double, _
        ByVal settings As Variant, ByVal baseBarOrdinal As Long, ByVal tangentX As Double, ByVal tangentY As Double, _
        ByVal inwardX As Double, ByVal inwardY As Double, ByVal steelClass As String, ByRef nextIndex As Long)
    Dim row2Diameter As Double
    Dim row3Diameter As Double
    row2Diameter = FaceRow2Diameter(settings)
    row3Diameter = FaceRow3Diameter(settings)

    If row2Diameter > 0# And ShouldAddAdditionalRow(baseBarOrdinal, FaceRow2Binding(settings)) Then
        Dim row2X As Double, row2Y As Double
        OffsetAdditionalRow baseX, baseY, row1Diameter, row2Diameter, 0#, FaceRow2Location(settings), _
            tangentX, tangentY, inwardX, inwardY, row2X, row2Y
        AddBarIfAvailable layout, geom, faceName, 2, row2X, row2Y, row2Diameter, steelClass, nextIndex
    End If

    If row3Diameter > 0# And ShouldAddAdditionalRow(baseBarOrdinal, FaceRow3Binding(settings)) Then
        Dim skippedDiameter As Double
        skippedDiameter = 0#
        If row2Diameter > 0# And StrComp(FaceRow2Location(settings), FaceRow3Location(settings), vbTextCompare) = 0 Then skippedDiameter = row2Diameter
        Dim row3X As Double, row3Y As Double
        OffsetAdditionalRow baseX, baseY, row1Diameter, row3Diameter, skippedDiameter, FaceRow3Location(settings), _
            tangentX, tangentY, inwardX, inwardY, row3X, row3Y
        AddBarIfAvailable layout, geom, faceName, 3, row3X, row3Y, row3Diameter, steelClass, nextIndex
    End If
End Sub

' Возвращает координату дополнительного стержня относительно базового стержня.
' Расстояние между осями считается по сумме радиусов стержней; если третий ряд
' ставится в том же направлении, что и второй, диаметр второго ряда добавляется
' как пропущенный промежуток между первым и третьим рядами.
Private Sub OffsetAdditionalRow(ByVal baseX As Double, ByVal baseY As Double, _
        ByVal row1Diameter As Double, ByVal targetDiameter As Double, ByVal skippedDiameter As Double, _
        ByVal locationMode As String, ByVal tangentX As Double, ByVal tangentY As Double, _
        ByVal inwardX As Double, ByVal inwardY As Double, ByRef xCoord As Double, ByRef yCoord As Double)
    Dim distance As Double
    distance = row1Diameter / 2# + skippedDiameter + targetDiameter / 2#
    If StrComp(locationMode, "SideBySide", vbTextCompare) = 0 Then
        xCoord = baseX + tangentX * distance
        yCoord = baseY + tangentY * distance
    Else
        xCoord = baseX + inwardX * distance
        yCoord = baseY + inwardY * distance
    End If
End Sub

' Проверяет, нужно ли привязывать дополнительный ряд к конкретному стержню
' первого ряда. EachBar дублирует каждый стержень, EverySecondBar - только
' нечетные номера, чтобы можно было разреженно поставить второй/третий ряд.
Private Function ShouldAddAdditionalRow(ByVal baseBarOrdinal As Long, ByVal bindingMode As String) As Boolean
    If StrComp(bindingMode, "EverySecondBar", vbTextCompare) = 0 Then
        ShouldAddAdditionalRow = ((baseBarOrdinal Mod 2) = 1)
    Else
        ShouldAddAdditionalRow = True
    End If
End Function

' Добавляет стержень, если в этой точке еще нет ранее созданного стержня.
' Такая защита нужна в местах, где линии B и H могут геометрически сблизиться:
' пользователь видит запрошенное число отдельно по граням, а расчетная модель
' не получает две арматуры с одинаковыми координатами.
Private Sub AddBarIfAvailable(ByVal layout As CRebarLayout, ByVal geom As CGeometryRoundedRectangle, _
        ByVal faceName As String, ByVal rowNumber As Long, ByVal xCoord As Double, ByVal yCoord As Double, _
        ByVal diameter As Double, ByVal steelClass As String, ByRef nextIndex As Long)
    If HasBarAt(layout, xCoord, yCoord) Then
        mSkippedDuplicateCount = mSkippedDuplicateCount + 1
        Exit Sub
    End If
    layout.AddBar faceName & ".row_" & CStr(rowNumber) & "-" & CStr(nextIndex), _
        xCoord, yCoord, diameter, 0#, steelClass, _
        "RoundedRectangle: " & RoundedFaceDisplayName(faceName) & ", ряд " & CStr(rowNumber), geom, faceName
    nextIndex = nextIndex + 1
End Sub

' Проверяет пользовательскую строку арматуры одной смысловой грани.
' Пустой/нулевой диаметр или n=0 означает "линию не создавать"; ошибки
' выдаются только для активной линии или явно некорректных дополнительных рядов.
Private Sub ValidateFace(ByVal faceName As String, ByVal axisDistance As Double, ByVal diameter As Double, _
        ByVal barCount As Long, ByVal row2Diameter As Double, ByVal row3Diameter As Double, _
        ByVal row2Location As String, ByVal row3Location As String, ByVal row2Binding As String, ByVal row3Binding As String)
    If barCount < 0 Then Err.Raise vbObjectError + 3230, "CRoundedRectRebarLayoutBuilder", RoundedFaceDisplayName(faceName) & ": количество стержней n должно быть нулевым или положительным."
    If diameter < 0# Then Err.Raise vbObjectError + 3231, "CRoundedRectRebarLayoutBuilder", RoundedFaceDisplayName(faceName) & ": диаметр d стержней первого ряда не может быть отрицательным."
    If barCount = 0 Or diameter = 0# Then Exit Sub
    If axisDistance <= diameter / 2# Then Err.Raise vbObjectError + 3232, "CRoundedRectRebarLayoutBuilder", RoundedFaceDisplayName(faceName) & ": расстояние as до оси стержня должно быть больше d/2."
    If row2Diameter < 0# Or row3Diameter < 0# Then Err.Raise vbObjectError + 3233, "CRoundedRectRebarLayoutBuilder", RoundedFaceDisplayName(faceName) & ": диаметры дополнительных рядов d2 и d3 должны быть неотрицательными."
    If row2Diameter > 0# And Not IsSupportedRowLocation(row2Location) Then Err.Raise vbObjectError + 3234, "CRoundedRectRebarLayoutBuilder", RoundedFaceDisplayName(faceName) & ": положение второго ряда должно быть SideBySide или Stacked."
    If row3Diameter > 0# And Not IsSupportedRowLocation(row3Location) Then Err.Raise vbObjectError + 3235, "CRoundedRectRebarLayoutBuilder", RoundedFaceDisplayName(faceName) & ": положение третьего ряда должно быть SideBySide или Stacked."
    If row2Diameter > 0# And Not IsSupportedRowBinding(row2Binding) Then Err.Raise vbObjectError + 3236, "CRoundedRectRebarLayoutBuilder", RoundedFaceDisplayName(faceName) & ": привязка второго ряда должна быть EachBar или EverySecondBar."
    If row3Diameter > 0# And Not IsSupportedRowBinding(row3Binding) Then Err.Raise vbObjectError + 3237, "CRoundedRectRebarLayoutBuilder", RoundedFaceDisplayName(faceName) & ": привязка третьего ряда должна быть EachBar или EverySecondBar."
End Sub

' Поддерживаемые способы размещения дополнительного ряда относительно первого.
Private Function IsSupportedRowLocation(ByVal value As String) As Boolean
    IsSupportedRowLocation = (StrComp(value, "SideBySide", vbTextCompare) = 0 Or StrComp(value, "Stacked", vbTextCompare) = 0)
End Function

' Поддерживаемые способы привязки дополнительного ряда к стержням первого ряда.
Private Function IsSupportedRowBinding(ByVal value As String) As Boolean
    IsSupportedRowBinding = (StrComp(value, "EachBar", vbTextCompare) = 0 Or StrComp(value, "EverySecondBar", vbTextCompare) = 0)
End Function

' Считает фактическую длину ломаной/дискретизированной кривой, по которой
' раскладывается H-арматура. Длина нужна именно по пути, а не по хорде между
' концами, чтобы шаг по tapered-части и дугам был равномерным.
Private Function PolylineLength(ByRef px() As Double, ByRef py() As Double) As Double
    Dim i As Long
    For i = LBound(px) To UBound(px) - 1
        PolylineLength = PolylineLength + Sqr((px(i + 1) - px(i)) ^ 2 + (py(i + 1) - py(i)) ^ 2)
    Next i
End Function

' Находит точку и локальную касательную на заданном расстоянии от начала
' бокового пути. Этим методом H-арматура получает равномерный шаг по фактической
' траектории, а дополнительные ряды - правильное локальное направление.
Private Sub PointAtDistance(ByRef px() As Double, ByRef py() As Double, ByVal targetDistance As Double, _
        ByRef xCoord As Double, ByRef yCoord As Double, ByRef tangentX As Double, ByRef tangentY As Double)
    Dim accumulated As Double
    Dim i As Long
    For i = LBound(px) To UBound(px) - 1
        Dim segLen As Double
        segLen = Sqr((px(i + 1) - px(i)) ^ 2 + (py(i + 1) - py(i)) ^ 2)
        If segLen <= GEOM_TOLERANCE Then GoTo NextSegment
        If accumulated + segLen >= targetDistance - GEOM_TOLERANCE Then
            Dim t As Double
            t = (targetDistance - accumulated) / segLen
            If t < 0# Then t = 0#
            If t > 1# Then t = 1#
            xCoord = px(i) + t * (px(i + 1) - px(i))
            yCoord = py(i) + t * (py(i + 1) - py(i))
            tangentX = (px(i + 1) - px(i)) / segLen
            tangentY = (py(i + 1) - py(i)) / segLen
            Exit Sub
        End If
        accumulated = accumulated + segLen
NextSegment:
    Next i
    xCoord = px(UBound(px))
    yCoord = py(UBound(py))
    tangentX = 0#
    tangentY = -1#
End Sub

' Нормирует локальную нормаль или касательную перед смещением дополнительных рядов.
Private Sub NormalizeVector(ByRef xValue As Double, ByRef yValue As Double)
    Dim lenValue As Double
    lenValue = Sqr(xValue * xValue + yValue * yValue)
    If lenValue <= GEOM_TOLERANCE Then Exit Sub
    xValue = xValue / lenValue
    yValue = yValue / lenValue
End Sub

' Проверяет совпадение координат с уже добавленными стержнями с геометрическим
' допуском. Используется только внутри автоматической раскладки, чтобы убрать
' дубли на сопряжениях граней.
Private Function HasBarAt(ByVal layout As CRebarLayout, ByVal xCoord As Double, ByVal yCoord As Double) As Boolean
    Dim i As Long
    For i = 1 To layout.Count
        If Abs(layout.X(i) - xCoord) <= GEOM_TOLERANCE And Abs(layout.Y(i) - yCoord) <= GEOM_TOLERANCE Then
            HasBarAt = True
            Exit Function
        End If
    Next i
End Function

' Возвращает расстояние as из компактного массива настроек одной грани.
Private Function FaceAs(ByVal settings As Variant) As Double
    FaceAs = CDbl(settings(LBound(settings)))
End Function

' Возвращает диаметр первого ряда; ноль означает, что грань не армируется.
Private Function FaceDiameter(ByVal settings As Variant) As Double
    FaceDiameter = CDbl(settings(LBound(settings) + 1))
End Function

' Возвращает количество стержней первого ряда на смысловой грани.
Private Function FaceCount(ByVal settings As Variant) As Long
    FaceCount = CLng(settings(LBound(settings) + 2))
End Function

' Возвращает диаметр второго ряда, если он задан пользователем.
Private Function FaceRow2Diameter(ByVal settings As Variant) As Double
    FaceRow2Diameter = CDbl(settings(LBound(settings) + 3))
End Function

' Возвращает диаметр третьего ряда, если он задан пользователем.
Private Function FaceRow3Diameter(ByVal settings As Variant) As Double
    FaceRow3Diameter = CDbl(settings(LBound(settings) + 4))
End Function

' Возвращает способ размещения второго ряда относительно первого.
Private Function FaceRow2Location(ByVal settings As Variant) As String
    FaceRow2Location = CStr(settings(LBound(settings) + 5))
End Function

' Возвращает способ размещения третьего ряда относительно первого.
Private Function FaceRow3Location(ByVal settings As Variant) As String
    FaceRow3Location = CStr(settings(LBound(settings) + 6))
End Function

' Возвращает правило привязки второго ряда к стержням первого ряда.
Private Function FaceRow2Binding(ByVal settings As Variant) As String
    FaceRow2Binding = CStr(settings(LBound(settings) + 7))
End Function

' Возвращает правило привязки третьего ряда к стержням первого ряда.
Private Function FaceRow3Binding(ByVal settings As Variant) As String
    FaceRow3Binding = CStr(settings(LBound(settings) + 8))
End Function

' Переводит внутренний идентификатор смысловой грани RoundedRectangle в
' короткое пользовательское имя для ошибок и комментариев Results.
Private Function RoundedFaceDisplayName(ByVal faceName As String) As String
    Select Case LCase$(Trim$(faceName))
        Case "h.left": RoundedFaceDisplayName = "H - левая грань"
        Case "h.right": RoundedFaceDisplayName = "H - правая грань"
        Case "b.top": RoundedFaceDisplayName = "B - верхняя грань"
        Case "b.bottom": RoundedFaceDisplayName = "B - нижняя грань"
        Case Else: RoundedFaceDisplayName = faceName
    End Select
End Function

' Количество стержней первого ряда, которое пользователь запросил по всем граням.
Public Property Get RequestedCount() As Long
    RequestedCount = mRequestedCount
End Property

' Количество стержней, не добавленных из-за совпадения с уже созданной арматурой.
Public Property Get SkippedDuplicateCount() As Long
    SkippedDuplicateCount = mSkippedDuplicateCount
End Property
