VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CRectSetRebarLayoutBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Автоматическая раскладка арматуры RectSet-сечения
' ==========================================================================
' Builder читает таблицу граней H1/B1/H2/B2, расставляет стержни первого ряда
' и при необходимости добавляет второй/третий ряды. В режиме Rectangle активны
' только грани верхнего прямоугольника H1/B1; LSection формирует Г-сечение;
' в TwoRectangles верхний прямоугольник может быть смещен по X. Builder также
' формирует semantic-группы для подписей вида 6Ø32 + 3Ø20.

Private mStepAlong As Double
Private mPerimeter As Double
Private mCentroidX As Double
Private mCentroidY As Double
Private mRequestedCount As Long
Private mSkippedDuplicateCount As Long
Private Const REBAR_STEEL_CLASS As String = "Rebar" ' Техническая метка типа элемента: расчетная арматура, не профиль для phi2.

' Строит раскладку прямо по таблице rngRectSetGeometry.
' Registry только выбирает, что активна форма RectSet, а вся структура граней,
' строк d/n/as/t и дополнительных рядов остается здесь, рядом с логикой
' расстановки стержней и формирования rebar annotation anchors.
Public Function BuildFromSettings(ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem) As CRebarLayout
    If settings Is Nothing Then Err.Raise vbObjectError + 3190, "CRectSetRebarLayoutBuilder", "Настройки Config не переданы."
    If units Is Nothing Then Err.Raise vbObjectError + 3191, "CRectSetRebarLayoutBuilder", "Система единиц не передана."

    Set BuildFromSettings = Build( _
        units.InputLengthToInternal(settings.GetDouble("RectSet.B1", 250#)), _
        units.InputLengthToInternal(settings.GetDouble("RectSet.H1", 550#)), _
        units.InputLengthToInternal(settings.GetDouble("RectSet.B2", 600#)), _
        units.InputLengthToInternal(settings.GetDouble("RectSet.H2", 250#)), _
        0#, _
        0#, _
        ReadFaceSettingsFromConfig(settings, units, "H1"), _
        ReadFaceSettingsFromConfig(settings, units, "H2"), _
        ReadFaceSettingsFromConfig(settings, units, "B1"), _
        ReadFaceSettingsFromConfig(settings, units, "B2"), _
        REBAR_STEEL_CLASS, _
        units.InputLengthToInternal(settings.GetDouble("RectSet.UpperOffsetX", 0#)), _
        settings.GetString("RectSet.SectionType", "LSection"))
End Function

' Читает параметры одной смысловой грани RectSet-сечения.
' Для диаметров действует особое правило ввода: если ячейка существует, но
' оставлена пустой, это не default, а явное отсутствие арматуры на данной линии.
Private Function ReadFaceSettingsFromConfig(ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem, _
        ByVal faceName As String) As Variant
    ReadFaceSettingsFromConfig = Array( _
        units.InputLengthToInternal(settings.GetDouble("RectSet." & faceName & ".as_1", 40#)), _
        units.InputLengthToInternal(settings.GetDouble("RectSet." & faceName & ".as_2", 40#)), _
        OptionalDiameterFromSettings(settings, units, "RectSet." & faceName & ".d_1", 32#), _
        OptionalDiameterFromSettings(settings, units, "RectSet." & faceName & ".d_2", 32#), _
        settings.GetLong("RectSet." & faceName & ".n_1", 0), _
        settings.GetLong("RectSet." & faceName & ".n_2", 0), _
        units.InputLengthToInternal(settings.GetDouble("RectSet." & faceName & ".StartOffset1", 80#)), _
        units.InputLengthToInternal(settings.GetDouble("RectSet." & faceName & ".EndOffset1", 80#)), _
        units.InputLengthToInternal(settings.GetDouble("RectSet." & faceName & ".StartOffset2", 80#)), _
        units.InputLengthToInternal(settings.GetDouble("RectSet." & faceName & ".EndOffset2", 80#)), _
        OptionalDiameterFromSettings(settings, units, "RectSet." & faceName & ".d_2row_1", 0#), _
        OptionalDiameterFromSettings(settings, units, "RectSet." & faceName & ".d_2row_2", 0#), _
        OptionalDiameterFromSettings(settings, units, "RectSet." & faceName & ".d_3row_1", 0#), _
        OptionalDiameterFromSettings(settings, units, "RectSet." & faceName & ".d_3row_2", 0#), _
        settings.GetString("RectSet." & faceName & ".loc_2row", "Stacked"), _
        settings.GetString("RectSet." & faceName & ".loc_3row", "Stacked"), _
        settings.GetString("RectSet." & faceName & ".bind_2row", "EachBar"), _
        settings.GetString("RectSet." & faceName & ".bind_3row", "EachBar"))
End Function

' Возвращает диаметр из Config с учетом пользовательского смысла пустой ячейки.
' Если ключа вообще нет, используется default для программных вызовов builder-а.
' Если ключ есть, но значение пустое или TODO, диаметр считается нулевым.
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
' Метод оставлен как чистая расчетная точка входа для тестов и для сценариев,
' где геометрические параметры уже подготовлены без чтения Config.
Public Function Build(ByVal upperWidth As Double, ByVal upperHeight As Double, _
        ByVal lowerWidth As Double, ByVal lowerHeight As Double, _
        ByVal originX As Double, ByVal originY As Double, _
        ByVal h1Settings As Variant, ByVal h2Settings As Variant, _
        ByVal b1Settings As Variant, ByVal b2Settings As Variant, _
        ByVal steelClass As String, Optional ByVal upperOffsetX As Double = 0#, _
        Optional ByVal sectionType As String = "LSection") As CRebarLayout

    ValidateGeometry upperWidth, upperHeight, lowerWidth, lowerHeight, upperOffsetX, sectionType

    Dim geom As CGeometryRectSet
    Set geom = New CGeometryRectSet
    geom.Initialize upperWidth, upperHeight, lowerWidth, lowerHeight, originX, originY, upperOffsetX, sectionType

    Dim layout As CRebarLayout
    Set layout = New CRebarLayout

    Dim sx As Double
    Dim sy As Double
    Dim nextIndex As Long
    nextIndex = 1
    mPerimeter = RebarDiagnosticPerimeter(geom, upperWidth, upperHeight, lowerWidth, lowerHeight)
    mStepAlong = 0#
    mRequestedCount = 0
    mSkippedDuplicateCount = 0

    Dim upperX As Double
    Dim upperY As Double
    upperX = geom.UpperMinX
    upperY = geom.UpperMinY

    AddVerticalPair layout, geom, "H1", upperX, upperY, upperHeight, _
        upperWidth, h1Settings, b1Settings, steelClass, nextIndex, sx, sy
    AddHorizontalPair layout, geom, "B1", upperX, upperY, upperWidth, _
        upperHeight, b1Settings, h1Settings, steelClass, nextIndex, sx, sy
    If geom.HasLowerRectangle Then
        AddVerticalPair layout, geom, "H2", originX, originY, lowerHeight, _
            lowerWidth, h2Settings, b2Settings, steelClass, nextIndex, sx, sy
        AddHorizontalPair layout, geom, "B2", originX, originY, lowerWidth, _
            lowerHeight, b2Settings, h2Settings, steelClass, nextIndex, sx, sy
    End If

    If layout.Count > 0 Then
        mCentroidX = sx / CDbl(layout.Count)
        mCentroidY = sy / CDbl(layout.Count)
    Else
        mCentroidX = 0#
        mCentroidY = 0#
    End If
    Set Build = layout
End Function

' Передает размеры и смещение верхнего прямоугольника общему владельцу
' геометрии RectSet. Раскладка не дублирует правила допустимости формы;
' при отказе возвращается инженерное пояснение CGeometryRectSet.
Private Sub ValidateGeometry(ByVal b1 As Double, ByVal h1 As Double, ByVal b2 As Double, ByVal h2 As Double, _
        Optional ByVal upperOffsetX As Double = 0#, Optional ByVal sectionType As String = "LSection")
    Dim geom As CGeometryRectSet
    Dim message As String
    Set geom = New CGeometryRectSet
    geom.Initialize b1, h1, b2, h2, 0#, 0#, upperOffsetX, sectionType
    If Not geom.IsValid(message) Then Err.Raise vbObjectError + 3170, "CRectSetRebarLayoutBuilder", message
End Sub

' Возвращает диагностическую длину контура для тестов и отчетности builder-а.
' Для LSection используется сумма периметров двух прямоугольников, потому что
' диагностическая величина описывает исходные линии раскладки арматуры, а не
' внешний контур объединенной фигуры. Для остальных режимов берется фактический контур.
Private Function RebarDiagnosticPerimeter(ByVal geom As CGeometryRectSet, ByVal upperWidth As Double, _
        ByVal upperHeight As Double, ByVal lowerWidth As Double, ByVal lowerHeight As Double) As Double
    If StrComp(geom.SectionType, "LSection", vbTextCompare) = 0 Then
        RebarDiagnosticPerimeter = 2# * (upperWidth + upperHeight) + 2# * (lowerWidth + lowerHeight)
        Exit Function
    End If

    Dim px() As Double
    Dim py() As Double
    geom.GetExtremePoints px, py

    Dim i As Long
    Dim nextIndex As Long
    For i = LBound(px) To UBound(px)
        If i = UBound(px) Then
            nextIndex = LBound(px)
        Else
            nextIndex = i + 1
        End If
        RebarDiagnosticPerimeter = RebarDiagnosticPerimeter + Sqr((px(nextIndex) - px(i)) ^ 2 + (py(nextIndex) - py(i)) ^ 2)
    Next i
End Function

' Добавляет две линии арматуры на вертикальной смысловой грани H.
' Левая линия идет снизу вверх, правая - сверху вниз: такой порядок сохраняет
' единую часовую нумерацию граней и согласует t нач./t кон. с пользовательской
' справкой по RectSet.
Private Sub AddVerticalPair(ByVal layout As CRebarLayout, ByVal geom As ISectionGeometry, _
        ByVal faceName As String, ByVal xStart As Double, ByVal yStart As Double, _
        ByVal heightValue As Double, ByVal widthValue As Double, _
        ByVal verticalSettings As Variant, ByVal horizontalSettings As Variant, _
        ByVal steelClass As String, ByRef nextIndex As Long, ByRef sx As Double, ByRef sy As Double)
    AddFaceBars layout, geom, faceName & ".as_1", _
        xStart + FaceAs1(verticalSettings), yStart, _
        xStart + FaceAs1(verticalSettings), yStart + heightValue, heightValue, _
        FaceLineSettings(verticalSettings, 1), steelClass, -1#, 0#, 0#, -1#, 1#, 0#, nextIndex, sx, sy
    AddFaceBars layout, geom, faceName & ".as_2", _
        xStart + widthValue - FaceAs2(verticalSettings), yStart + heightValue, _
        xStart + widthValue - FaceAs2(verticalSettings), yStart, heightValue, _
        FaceLineSettings(verticalSettings, 2), steelClass, 1#, 0#, 0#, -1#, -1#, 0#, nextIndex, sx, sy
End Sub

' Добавляет две линии арматуры на горизонтальной смысловой грани B.
' Верхняя линия идет слева направо, нижняя - справа налево, чтобы обход
' прямоугольника оставался часовым и пользовательские отступы t читались
' одинаково для всех граней.
Private Sub AddHorizontalPair(ByVal layout As CRebarLayout, ByVal geom As ISectionGeometry, _
        ByVal faceName As String, ByVal xStart As Double, ByVal yStart As Double, _
        ByVal widthValue As Double, ByVal heightValue As Double, _
        ByVal horizontalSettings As Variant, ByVal verticalSettings As Variant, _
        ByVal steelClass As String, ByRef nextIndex As Long, ByRef sx As Double, ByRef sy As Double)
    AddFaceBars layout, geom, faceName & ".as_1", _
        xStart, yStart + heightValue - FaceAs1(horizontalSettings), _
        xStart + widthValue, yStart + heightValue - FaceAs1(horizontalSettings), widthValue, _
        FaceLineSettings(horizontalSettings, 1), steelClass, 0#, 1#, 1#, 0#, 0#, -1#, nextIndex, sx, sy
    AddFaceBars layout, geom, faceName & ".as_2", _
        xStart + widthValue, yStart + FaceAs2(horizontalSettings), _
        xStart, yStart + FaceAs2(horizontalSettings), widthValue, _
        FaceLineSettings(horizontalSettings, 2), steelClass, 0#, -1#, 1#, 0#, 0#, 1#, nextIndex, sx, sy
End Sub

' Раскладывает один ряд первого уровня по прямой линии грани и сразу добавляет
' связанные второй/третий ряды. Линия задается начальной/конечной точками, а
' side/stack-направления передаются явно, потому что для H и B они различаются.
Private Sub AddFaceBars(ByVal layout As CRebarLayout, ByVal geom As ISectionGeometry, _
        ByVal faceName As String, ByVal x1 As Double, ByVal y1 As Double, _
        ByVal x2 As Double, ByVal y2 As Double, ByVal faceLength As Double, _
        ByVal faceSettings As Variant, ByVal steelClass As String, _
        ByVal outsideNormalX As Double, ByVal outsideNormalY As Double, _
        ByVal sideX As Double, ByVal sideY As Double, ByVal stackX As Double, ByVal stackY As Double, _
        ByRef nextIndex As Long, ByRef sx As Double, ByRef sy As Double)
    Dim axisDistance As Double ' as: расстояние от бетонной грани до оси первого ряда.
    Dim barCount As Long ' n: число стержней первого ряда на этой линии.
    Dim barDiameter As Double ' d: диаметр первого ряда.
    Dim startOffset As Double ' t нач.: отступ первого стержня от начала линии обхода.
    Dim endOffset As Double ' t кон.: отступ последнего стержня от конца линии обхода.
    Dim row2Diameter As Double ' Диаметр дополнительного второго ряда.
    Dim row3Diameter As Double ' Диаметр дополнительного третьего ряда.
    Dim row2Location As String ' Stacked/SideBySide для второго ряда.
    Dim row3Location As String ' Stacked/SideBySide для третьего ряда.
    Dim row2Binding As String ' EachBar/EverySecondBar для второго ряда.
    Dim row3Binding As String ' EachBar/EverySecondBar для третьего ряда.
    barCount = LineBarCount(faceSettings)
    If barCount = 0 Then Exit Sub

    axisDistance = LineAxisDistance(faceSettings)
    barDiameter = LineBarDiameter(faceSettings)
    startOffset = LineStartOffset(faceSettings)
    endOffset = LineEndOffset(faceSettings)
    row2Diameter = LineRow2Diameter(faceSettings)
    row3Diameter = LineRow3Diameter(faceSettings)
    row2Location = LineRow2Location(faceSettings)
    row3Location = LineRow3Location(faceSettings)
    row2Binding = LineRow2Binding(faceSettings)
    row3Binding = LineRow3Binding(faceSettings)

    Dim lineLabel As String
    lineLabel = RebarLineDisplayName(faceName)

    If barCount < 0 Then Err.Raise vbObjectError + 3171, "CRectSetRebarLayoutBuilder", lineLabel & ": количество стержней n должно быть нулевым или положительным."
    If barDiameter < 0# Then Err.Raise vbObjectError + 3172, "CRectSetRebarLayoutBuilder", lineLabel & ": диаметр d стержней первого ряда не может быть отрицательным."
    If barCount = 0 Or barDiameter = 0# Then Exit Sub
    If axisDistance <= barDiameter / 2# Then Err.Raise vbObjectError + 3173, "CRectSetRebarLayoutBuilder", lineLabel & ": расстояние as до оси стержня должно быть больше d/2."
    If startOffset < 0# Or endOffset < 0# Then Err.Raise vbObjectError + 3174, "CRectSetRebarLayoutBuilder", lineLabel & ": отступы t нач. и t кон. должны быть неотрицательными."
    ValidateAdditionalRows lineLabel, row2Diameter, row3Diameter, row2Location, row3Location, row2Binding, row3Binding

    Dim usableLength As Double
    usableLength = faceLength - startOffset - endOffset
    If usableLength < -GEOM_TOLERANCE Then Err.Raise vbObjectError + 3175, "CRectSetRebarLayoutBuilder", lineLabel & ": сумма отступов t нач. + t кон. превышает длину грани."

    Dim stepAlong As Double
    If barCount > 1 Then
        stepAlong = usableLength / CDbl(barCount - 1)
    Else
        stepAlong = 0#
    End If
    If stepAlong > mStepAlong Then mStepAlong = stepAlong
    mRequestedCount = mRequestedCount + barCount

    Dim distances() As Double ' Расстояния стержней первого ряда от начала линии.
    ReDim distances(1 To barCount)
    Dim i As Long
    If barCount = 1 Then
        distances(1) = startOffset + usableLength / 2#
    Else
        For i = 1 To barCount
            distances(i) = startOffset + CDbl(i - 1) * stepAlong
        Next i
    End If

    AddFaceAnnotationAnchor layout, faceName, x1, y1, x2, y2, faceLength, _
        distances(1), distances(barCount), outsideNormalX, outsideNormalY, axisDistance

    Dim used() As Boolean
    ReDim used(1 To barCount)
    For i = 1 To barCount
        Dim chosen As Long ' Номер еще не использованной позиции, ближайшей к середине линии.
        chosen = ClosestUnusedDistanceIndex(distances, used, faceLength)
        used(chosen) = True

        Dim t As Double
        If faceLength <= GEOM_TOLERANCE Then
            t = 0#
        Else
            t = distances(chosen) / faceLength
        End If

        Dim xCoord As Double
        Dim yCoord As Double
        xCoord = x1 + t * (x2 - x1)
        yCoord = y1 + t * (y2 - y1)
        AddBarIfAvailable layout, geom, faceName, 1, xCoord, yCoord, barDiameter, steelClass, nextIndex, sx, sy
        AddAdditionalRows layout, geom, faceName, xCoord, yCoord, barDiameter, row2Diameter, row3Diameter, _
            row2Location, row3Location, row2Binding, row3Binding, chosen, sideX, sideY, stackX, stackY, steelClass, nextIndex, sx, sy
    Next i
End Sub

' Сохраняет semantic-anchor для групповой подписи арматуры на схеме.
' Anchor берется не по полной бетонной грани, а между первым и последним
' фактически поставленным стержнем, чтобы подпись визуально относилась к ряду.
Private Sub AddFaceAnnotationAnchor(ByVal layout As CRebarLayout, ByVal faceName As String, _
        ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double, _
        ByVal faceLength As Double, ByVal startDistance As Double, ByVal endDistance As Double, _
        ByVal normalX As Double, ByVal normalY As Double, ByVal axisDistance As Double)
    Dim startT As Double, endT As Double
    If faceLength <= GEOM_TOLERANCE Then
        startT = 0#
        endT = 1#
    Else
        startT = startDistance / faceLength
        endT = endDistance / faceLength
    End If
    If endT < startT Then endT = startT

    layout.AddAnnotationAnchor faceName, _
        x1 + startT * (x2 - x1), y1 + startT * (y2 - y1), _
        x1 + endT * (x2 - x1), y1 + endT * (y2 - y1), _
        normalX, normalY, axisDistance
End Sub

' Добавляет второй и третий ряды для выбранного стержня первого ряда.
' Если третий ряд имеет то же направление, что и второй, он "перепрыгивает"
' второй ряд на его диаметр, чтобы стержни не наложились.
Private Sub AddAdditionalRows(ByVal layout As CRebarLayout, ByVal geom As ISectionGeometry, _
        ByVal faceName As String, ByVal baseX As Double, ByVal baseY As Double, ByVal row1Diameter As Double, _
        ByVal row2Diameter As Double, ByVal row3Diameter As Double, ByVal row2Location As String, ByVal row3Location As String, _
        ByVal row2Binding As String, ByVal row3Binding As String, ByVal baseBarOrdinal As Long, _
        ByVal sideX As Double, ByVal sideY As Double, ByVal stackX As Double, ByVal stackY As Double, _
        ByVal steelClass As String, ByRef nextIndex As Long, ByRef sx As Double, ByRef sy As Double)
    Dim row2X As Double
    Dim row2Y As Double
    If row2Diameter > 0# And ShouldAddAdditionalRow(baseBarOrdinal, row2Binding) Then
        OffsetAdditionalRowFromFirst baseX, baseY, row1Diameter, row2Diameter, 0#, row2Location, _
            sideX, sideY, stackX, stackY, row2X, row2Y
        AddBarIfAvailable layout, geom, faceName, 2, row2X, row2Y, row2Diameter, steelClass, nextIndex, sx, sy
    End If

    Dim row3X As Double
    Dim row3Y As Double
    If row3Diameter > 0# And ShouldAddAdditionalRow(baseBarOrdinal, row3Binding) Then
        Dim skippedDiameter As Double
        skippedDiameter = 0#
        If row2Diameter > 0# And StrComp(row2Location, row3Location, vbTextCompare) = 0 Then skippedDiameter = row2Diameter
        OffsetAdditionalRowFromFirst baseX, baseY, row1Diameter, row3Diameter, skippedDiameter, row3Location, _
            sideX, sideY, stackX, stackY, row3X, row3Y
        AddBarIfAvailable layout, geom, faceName, 3, row3X, row3Y, row3Diameter, steelClass, nextIndex, sx, sy
    End If
End Sub

' Проверяет правило привязки дополнительного ряда к стержням первого ряда.
' EverySecondBar оставляет только нечетные номера, считая от начала линии.
Private Function ShouldAddAdditionalRow(ByVal baseBarOrdinal As Long, ByVal bindingMode As String) As Boolean
    If StrComp(bindingMode, "EverySecondBar", vbTextCompare) = 0 Then
        ShouldAddAdditionalRow = ((baseBarOrdinal Mod 2) = 1)
    Else
        ShouldAddAdditionalRow = True
    End If
End Function

' Считает координаты дополнительного стержня от стержня первого ряда.
' SideBySide смещает вдоль линии грани, Stacked - внутрь сечения по нормали,
' а расстояние между осями равно сумме радиусов с учетом пропущенного ряда.
Private Sub OffsetAdditionalRowFromFirst(ByVal baseX As Double, ByVal baseY As Double, _
        ByVal row1Diameter As Double, ByVal targetDiameter As Double, ByVal skippedDiameter As Double, ByVal locationMode As String, _
        ByVal sideX As Double, ByVal sideY As Double, ByVal stackX As Double, ByVal stackY As Double, _
        ByRef xCoord As Double, ByRef yCoord As Double)
    Dim distance As Double
    distance = row1Diameter / 2# + skippedDiameter + targetDiameter / 2#

    If StrComp(locationMode, "SideBySide", vbTextCompare) = 0 Then
        xCoord = baseX + sideX * distance
        yCoord = baseY + sideY * distance
    Else
        xCoord = baseX + stackX * distance
        yCoord = baseY + stackY * distance
    End If
End Sub

' Добавляет стержень только если в этой координате еще нет арматуры.
' Это защищает углы и пересечения прямоугольников от двойного учета одной и той
' же физической позиции, когда пользователь задал соседние линии с совпадением.
Private Sub AddBarIfAvailable(ByVal layout As CRebarLayout, ByVal geom As ISectionGeometry, _
        ByVal faceName As String, ByVal rowNumber As Long, ByVal xCoord As Double, ByVal yCoord As Double, _
        ByVal diameter As Double, ByVal steelClass As String, ByRef nextIndex As Long, ByRef sx As Double, ByRef sy As Double)
    If HasBarAt(layout, xCoord, yCoord) Then
        mSkippedDuplicateCount = mSkippedDuplicateCount + 1
    Else
        layout.AddBar faceName & ".row_" & CStr(rowNumber) & "-" & CStr(nextIndex), _
            xCoord, yCoord, diameter, 0#, steelClass, _
            "RectSet: " & RebarLineDisplayName(faceName) & ", ряд " & CStr(rowNumber), geom, faceName
        nextIndex = nextIndex + 1
        sx = sx + xCoord
        sy = sy + yCoord
    End If
End Sub

' Возвращает as одной линии из компактного массива настроек линии.
Private Function LineAxisDistance(ByVal lineSettings As Variant) As Double
    LineAxisDistance = CDbl(lineSettings(LBound(lineSettings)))
End Function

' Возвращает n - число стержней первого ряда линии.
Private Function LineBarCount(ByVal lineSettings As Variant) As Long
    LineBarCount = CLng(lineSettings(LBound(lineSettings) + 1))
End Function

' Возвращает d - диаметр первого ряда линии.
Private Function LineBarDiameter(ByVal lineSettings As Variant) As Double
    LineBarDiameter = CDbl(lineSettings(LBound(lineSettings) + 2))
End Function

' Возвращает t нач. для первого стержня линии.
Private Function LineStartOffset(ByVal lineSettings As Variant) As Double
    LineStartOffset = CDbl(lineSettings(LBound(lineSettings) + 3))
End Function

' Возвращает t кон. для последнего стержня линии.
Private Function LineEndOffset(ByVal lineSettings As Variant) As Double
    LineEndOffset = CDbl(lineSettings(LBound(lineSettings) + 4))
End Function

' Возвращает диаметр второго ряда; отсутствие optional поля означает, что ряда нет.
Private Function LineRow2Diameter(ByVal lineSettings As Variant) As Double
    If UBound(lineSettings) - LBound(lineSettings) < 5 Then Exit Function
    LineRow2Diameter = CDbl(lineSettings(LBound(lineSettings) + 5))
End Function

' Возвращает диаметр третьего ряда; отсутствие optional поля означает, что ряда нет.
Private Function LineRow3Diameter(ByVal lineSettings As Variant) As Double
    If UBound(lineSettings) - LBound(lineSettings) < 6 Then Exit Function
    LineRow3Diameter = CDbl(lineSettings(LBound(lineSettings) + 6))
End Function

' Возвращает способ размещения второго ряда, по умолчанию Stacked.
Private Function LineRow2Location(ByVal lineSettings As Variant) As String
    LineRow2Location = "Stacked"
    If UBound(lineSettings) - LBound(lineSettings) < 7 Then Exit Function
    If Len(Trim$(CStr(lineSettings(LBound(lineSettings) + 7)))) > 0 Then LineRow2Location = Trim$(CStr(lineSettings(LBound(lineSettings) + 7)))
End Function

' Возвращает способ размещения третьего ряда, по умолчанию Stacked.
Private Function LineRow3Location(ByVal lineSettings As Variant) As String
    LineRow3Location = "Stacked"
    If UBound(lineSettings) - LBound(lineSettings) < 8 Then Exit Function
    If Len(Trim$(CStr(lineSettings(LBound(lineSettings) + 8)))) > 0 Then LineRow3Location = Trim$(CStr(lineSettings(LBound(lineSettings) + 8)))
End Function

' Возвращает правило привязки второго ряда к стержням первого ряда.
Private Function LineRow2Binding(ByVal lineSettings As Variant) As String
    LineRow2Binding = "EachBar"
    If UBound(lineSettings) - LBound(lineSettings) < 9 Then Exit Function
    If Len(Trim$(CStr(lineSettings(LBound(lineSettings) + 9)))) > 0 Then LineRow2Binding = Trim$(CStr(lineSettings(LBound(lineSettings) + 9)))
End Function

' Возвращает правило привязки третьего ряда к стержням первого ряда.
Private Function LineRow3Binding(ByVal lineSettings As Variant) As String
    LineRow3Binding = "EachBar"
    If UBound(lineSettings) - LBound(lineSettings) < 10 Then Exit Function
    If Len(Trim$(CStr(lineSettings(LBound(lineSettings) + 10)))) > 0 Then LineRow3Binding = Trim$(CStr(lineSettings(LBound(lineSettings) + 10)))
End Function

' Возвращает пользовательскую подпись стороны арматуры для сообщений и Results.
' Техническое имя H2.as_1 сохраняется для annotation-групп, но в ошибках
' и комментариях показывается конкретная грань: левая/правая/верхняя/нижняя.
Private Function RebarLineDisplayName(ByVal faceName As String) As String
    Dim baseName As String
    Dim sideIndex As Long
    If Right$(faceName, 5) = ".as_1" Then
        baseName = Left$(faceName, Len(faceName) - 5)
        sideIndex = 1
    ElseIf Right$(faceName, 5) = ".as_2" Then
        baseName = Left$(faceName, Len(faceName) - 5)
        sideIndex = 2
    Else
        RebarLineDisplayName = faceName
        Exit Function
    End If

    RebarLineDisplayName = baseName & " - " & RectSetSideDisplayName(baseName, sideIndex)
End Function

' Переводит сторону RectSet в обычные слова пользователя.
' Для H-граней стороны читаются как левая/правая, для B-граней - верхняя/нижняя.
Private Function RectSetSideDisplayName(ByVal faceBaseName As String, ByVal sideIndex As Long) As String
    Dim prefix As String
    prefix = UCase$(Left$(faceBaseName, 1))
    If prefix = "H" Then
        If sideIndex = 1 Then
            RectSetSideDisplayName = "левая грань"
        Else
            RectSetSideDisplayName = "правая грань"
        End If
    ElseIf prefix = "B" Then
        If sideIndex = 1 Then
            RectSetSideDisplayName = "верхняя грань"
        Else
            RectSetSideDisplayName = "нижняя грань"
        End If
    Else
        RectSetSideDisplayName = "сторона " & CStr(sideIndex)
    End If
End Function

' Проверяет неотрицательные диаметры и режимы активных дополнительных рядов.
' Нулевой диаметр означает отсутствие ряда, поэтому его положение и привязка
' не участвуют в построении и не требуют активного значения selector-а.
Private Sub ValidateAdditionalRows(ByVal faceName As String, ByVal row2Diameter As Double, ByVal row3Diameter As Double, _
        ByVal row2Location As String, ByVal row3Location As String, ByVal row2Binding As String, ByVal row3Binding As String)
    If row2Diameter < 0# Or row3Diameter < 0# Then Err.Raise vbObjectError + 3180, "CRectSetRebarLayoutBuilder", faceName & ": диаметры дополнительных рядов d_2row и d_3row должны быть неотрицательными."
    If row2Diameter > 0# And Not IsSupportedRowLocation(row2Location) Then Err.Raise vbObjectError + 3182, "CRectSetRebarLayoutBuilder", faceName & ": loc_2row должен быть SideBySide или Stacked."
    If row3Diameter > 0# And Not IsSupportedRowLocation(row3Location) Then Err.Raise vbObjectError + 3183, "CRectSetRebarLayoutBuilder", faceName & ": loc_3row должен быть SideBySide или Stacked."
    If row2Diameter > 0# And Not IsSupportedRowBinding(row2Binding) Then Err.Raise vbObjectError + 3184, "CRectSetRebarLayoutBuilder", faceName & ": bind_2row должен быть EachBar или EverySecondBar."
    If row3Diameter > 0# And Not IsSupportedRowBinding(row3Binding) Then Err.Raise vbObjectError + 3185, "CRectSetRebarLayoutBuilder", faceName & ": bind_3row должен быть EachBar или EverySecondBar."
End Sub

' Проверяет пользовательское значение loc_2row/loc_3row.
Private Function IsSupportedRowLocation(ByVal value As String) As Boolean
    IsSupportedRowLocation = (StrComp(value, "SideBySide", vbTextCompare) = 0 Or StrComp(value, "Stacked", vbTextCompare) = 0)
End Function

' Проверяет пользовательское значение bind_2row/bind_3row.
Private Function IsSupportedRowBinding(ByVal value As String) As Boolean
    IsSupportedRowBinding = (StrComp(value, "EachBar", vbTextCompare) = 0 Or StrComp(value, "EverySecondBar", vbTextCompare) = 0)
End Function

' Собирает настройки одной стороны смысловой грани в компактный массив линии.
' Это позволяет общей процедуре AddFaceBars одинаково работать с _1 и _2.
Private Function FaceLineSettings(ByVal faceSettings As Variant, ByVal lineIndex As Long) As Variant
    If lineIndex = 1 Then
        FaceLineSettings = Array(FaceAs1(faceSettings), FaceCount1(faceSettings), FaceDiameter1(faceSettings), _
            FaceStartOffset1(faceSettings), FaceEndOffset1(faceSettings), _
            FaceRow2Diameter1(faceSettings), FaceRow3Diameter1(faceSettings), FaceRow2Location(faceSettings), FaceRow3Location(faceSettings), _
            FaceRow2Binding(faceSettings), FaceRow3Binding(faceSettings))
    Else
        FaceLineSettings = Array(FaceAs2(faceSettings), FaceCount2(faceSettings), FaceDiameter2(faceSettings), _
            FaceStartOffset2(faceSettings), FaceEndOffset2(faceSettings), _
            FaceRow2Diameter2(faceSettings), FaceRow3Diameter2(faceSettings), FaceRow2Location(faceSettings), FaceRow3Location(faceSettings), _
            FaceRow2Binding(faceSettings), FaceRow3Binding(faceSettings))
    End If
End Function

' Возвращает as для стороны _1: левая для H или верхняя для B.
Private Function FaceAs1(ByVal faceSettings As Variant) As Double
    FaceAs1 = CDbl(faceSettings(LBound(faceSettings)))
End Function

' Возвращает as для стороны _2: правая для H или нижняя для B.
Private Function FaceAs2(ByVal faceSettings As Variant) As Double
    FaceAs2 = CDbl(faceSettings(LBound(faceSettings) + 1))
End Function

' Возвращает d первого ряда для стороны _1.
Private Function FaceDiameter1(ByVal faceSettings As Variant) As Double
    FaceDiameter1 = CDbl(faceSettings(LBound(faceSettings) + 2))
End Function

' Возвращает d первого ряда для стороны _2.
Private Function FaceDiameter2(ByVal faceSettings As Variant) As Double
    FaceDiameter2 = CDbl(faceSettings(LBound(faceSettings) + 3))
End Function

' Возвращает n первого ряда для стороны _1.
Private Function FaceCount1(ByVal faceSettings As Variant) As Long
    FaceCount1 = CLng(faceSettings(LBound(faceSettings) + 4))
End Function

' Возвращает n первого ряда для стороны _2.
Private Function FaceCount2(ByVal faceSettings As Variant) As Long
    FaceCount2 = CLng(faceSettings(LBound(faceSettings) + 5))
End Function

' Возвращает t нач. для стороны _1.
Private Function FaceStartOffset1(ByVal faceSettings As Variant) As Double
    FaceStartOffset1 = CDbl(faceSettings(LBound(faceSettings) + 6))
End Function

' Возвращает t кон. для стороны _1.
Private Function FaceEndOffset1(ByVal faceSettings As Variant) As Double
    FaceEndOffset1 = CDbl(faceSettings(LBound(faceSettings) + 7))
End Function

' Возвращает t нач. для стороны _2.
Private Function FaceStartOffset2(ByVal faceSettings As Variant) As Double
    FaceStartOffset2 = CDbl(faceSettings(LBound(faceSettings) + 8))
End Function

' Возвращает t кон. для стороны _2.
Private Function FaceEndOffset2(ByVal faceSettings As Variant) As Double
    FaceEndOffset2 = CDbl(faceSettings(LBound(faceSettings) + 9))
End Function

' Возвращает диаметр второго ряда для стороны _1, если он задан.
Private Function FaceRow2Diameter1(ByVal faceSettings As Variant) As Double
    If UBound(faceSettings) - LBound(faceSettings) < 10 Then Exit Function
    FaceRow2Diameter1 = CDbl(faceSettings(LBound(faceSettings) + 10))
End Function

' Возвращает диаметр второго ряда для стороны _2, если он задан.
Private Function FaceRow2Diameter2(ByVal faceSettings As Variant) As Double
    If UBound(faceSettings) - LBound(faceSettings) < 11 Then Exit Function
    FaceRow2Diameter2 = CDbl(faceSettings(LBound(faceSettings) + 11))
End Function

' Возвращает диаметр третьего ряда для стороны _1, если он задан.
Private Function FaceRow3Diameter1(ByVal faceSettings As Variant) As Double
    If UBound(faceSettings) - LBound(faceSettings) < 12 Then Exit Function
    FaceRow3Diameter1 = CDbl(faceSettings(LBound(faceSettings) + 12))
End Function

' Возвращает диаметр третьего ряда для стороны _2, если он задан.
Private Function FaceRow3Diameter2(ByVal faceSettings As Variant) As Double
    If UBound(faceSettings) - LBound(faceSettings) < 13 Then Exit Function
    FaceRow3Diameter2 = CDbl(faceSettings(LBound(faceSettings) + 13))
End Function

' Возвращает общий способ размещения второго ряда для грани.
Private Function FaceRow2Location(ByVal faceSettings As Variant) As String
    FaceRow2Location = "Stacked"
    If UBound(faceSettings) - LBound(faceSettings) < 14 Then Exit Function
    If Len(Trim$(CStr(faceSettings(LBound(faceSettings) + 14)))) > 0 Then FaceRow2Location = Trim$(CStr(faceSettings(LBound(faceSettings) + 14)))
End Function

' Возвращает общий способ размещения третьего ряда для грани.
Private Function FaceRow3Location(ByVal faceSettings As Variant) As String
    FaceRow3Location = "Stacked"
    If UBound(faceSettings) - LBound(faceSettings) < 15 Then Exit Function
    If Len(Trim$(CStr(faceSettings(LBound(faceSettings) + 15)))) > 0 Then FaceRow3Location = Trim$(CStr(faceSettings(LBound(faceSettings) + 15)))
End Function

' Возвращает правило привязки второго ряда для грани.
Private Function FaceRow2Binding(ByVal faceSettings As Variant) As String
    FaceRow2Binding = "EachBar"
    If UBound(faceSettings) - LBound(faceSettings) < 16 Then Exit Function
    If Len(Trim$(CStr(faceSettings(LBound(faceSettings) + 16)))) > 0 Then FaceRow2Binding = Trim$(CStr(faceSettings(LBound(faceSettings) + 16)))
End Function

' Возвращает правило привязки третьего ряда для грани.
Private Function FaceRow3Binding(ByVal faceSettings As Variant) As String
    FaceRow3Binding = "EachBar"
    If UBound(faceSettings) - LBound(faceSettings) < 17 Then Exit Function
    If Len(Trim$(CStr(faceSettings(LBound(faceSettings) + 17)))) > 0 Then FaceRow3Binding = Trim$(CStr(faceSettings(LBound(faceSettings) + 17)))
End Function

' Выбирает следующую позицию первого ряда, начиная с ближайшей к середине линии.
' Такой порядок уменьшает шанс, что из-за дублей на углах сначала будут заняты
' только крайние точки и групповой anchor получится визуально смещенным.
Private Function ClosestUnusedDistanceIndex(ByRef distances() As Double, ByRef used() As Boolean, ByVal faceLength As Double) As Long
    Dim bestDistance As Double
    bestDistance = 1E+99
    Dim i As Long
    For i = LBound(distances) To UBound(distances)
        If Not used(i) Then
            Dim currentDistance As Double
            currentDistance = Abs(distances(i) - faceLength / 2#)
            If currentDistance < bestDistance Then
                bestDistance = currentDistance
                ClosestUnusedDistanceIndex = i
            End If
        End If
    Next i
End Function

' Проверяет, есть ли уже стержень в той же точке с геометрическим допуском.
Private Function HasBarAt(ByVal layout As CRebarLayout, ByVal xCoord As Double, ByVal yCoord As Double) As Boolean
    Dim i As Long
    For i = 1 To layout.Count
        If Abs(layout.X(i) - xCoord) <= GEOM_TOLERANCE And Abs(layout.Y(i) - yCoord) <= GEOM_TOLERANCE Then
            HasBarAt = True
            Exit Function
        End If
    Next i
End Function

' X-координата центра тяжести созданной арматуры после последнего Build.
Public Property Get CentroidX() As Double
    CentroidX = mCentroidX
End Property

' Y-координата центра тяжести созданной арматуры после последнего Build.
Public Property Get CentroidY() As Double
    CentroidY = mCentroidY
End Property

' Количество стержней, пропущенных из-за совпадения координат.
Public Property Get SkippedDuplicateCount() As Long
    SkippedDuplicateCount = mSkippedDuplicateCount
End Property

' ==========================================================================
' ДЛЯ ТЕСТОВ
' ==========================================================================
' Эти свойства раскрывают внутреннюю геометрию раскладки Г-образного сечения.
' Тесты используют их, чтобы проверить длину обхода и шаг расстановки без
' анализа уже созданных координат каждого стержня.

' Максимальный шаг по прямым линиям первого ряда, найденный в последнем Build.
Public Property Get StepAlong() As Double
    StepAlong = mStepAlong
End Property

' Диагностическая длина линий раскладки, использованная в последнем Build.
Public Property Get Perimeter() As Double
    Perimeter = mPerimeter
End Property

' Количество стержней первого ряда, запрошенное пользователем до удаления дублей.
Public Property Get RequestedCount() As Long
    RequestedCount = mRequestedCount
End Property

