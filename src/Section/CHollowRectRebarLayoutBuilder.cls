VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CHollowRectRebarLayoutBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Автоматическая раскладка арматуры HollowRectangle
' ==========================================================================
' Builder читает таблицу пустотелого скругленного прямоугольника и создает
' обычный CRebarLayout. Наружные грани H/B получают заданное пользователем
' количество стержней, а внутренние грани Opening заполняются проекциями
' стержней соответствующей наружной грани по нормали к линии раскладки.
'
' Для внешних граней "внутрь" направлено к телу бетонного кольца от наружного
' контура. Для внутренних граней Opening "внутрь" направлено от пустоты в
' бетон, то есть в противоположную сторону относительно нормали в отверстие.

Private Const REBAR_STEEL_CLASS As String = "Rebar" ' Материал продольной арматуры generated-модели.

' Строит геометрию и раскладку по настройкам Config.
Public Function BuildFromSettings(ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem) As CRebarLayout
    If settings Is Nothing Then Err.Raise vbObjectError + 3240, "CHollowRectRebarLayoutBuilder", "Настройки Config не переданы."
    If units Is Nothing Then Err.Raise vbObjectError + 3241, "CHollowRectRebarLayoutBuilder", "Система единиц не передана."

    Dim geom As CGeometryHollowRectangle
    Set geom = New CGeometryHollowRectangle
    geom.Initialize _
        units.InputLengthToInternal(settings.GetDouble("HollowRectangle.B", 500#)), _
        units.InputLengthToInternal(settings.GetDouble("HollowRectangle.H", 800#)), _
        units.InputLengthToInternal(settings.GetDouble("HollowRectangle.R", 180#)), _
        units.InputLengthToInternal(settings.GetDouble("HollowRectangle.OpeningB", 200#)), _
        units.InputLengthToInternal(settings.GetDouble("HollowRectangle.OpeningH", 500#)), _
        units.InputLengthToInternal(settings.GetDouble("HollowRectangle.OpeningR", 30#)), _
        units.InputLengthToInternal(settings.GetDouble("HollowRectangle.InnerOffsetX", 0#)), _
        units.InputLengthToInternal(settings.GetDouble("HollowRectangle.InnerOffsetY", 0#))

    Set BuildFromSettings = Build(geom, _
        ReadFaceSettingsFromConfig(settings, units, "H.Left", "40", "20"), _
        ReadFaceSettingsFromConfig(settings, units, "H.Right", "40", "20"), _
        ReadFaceSettingsFromConfig(settings, units, "B.Top", "40", "20"), _
        ReadFaceSettingsFromConfig(settings, units, "B.Bottom", "40", "20"), _
        ReadFaceSettingsFromConfig(settings, units, "Opening.H.Left", "40", "20"), _
        ReadFaceSettingsFromConfig(settings, units, "Opening.H.Right", "40", "20"), _
        ReadFaceSettingsFromConfig(settings, units, "Opening.B.Top", "40", "20"), _
        ReadFaceSettingsFromConfig(settings, units, "Opening.B.Bottom", "40", "20"), _
        REBAR_STEEL_CLASS)
End Function

' Создает CRebarLayout для уже подготовленной HollowRectangle-геометрии.
' Метод открыт отдельно, чтобы тесты могли проверять геометрию без Excel.
Public Function Build(ByVal geom As CGeometryHollowRectangle, _
        ByVal hLeftSettings As Variant, ByVal hRightSettings As Variant, _
        ByVal bTopSettings As Variant, ByVal bBottomSettings As Variant, _
        ByVal openingHLeftSettings As Variant, ByVal openingHRightSettings As Variant, _
        ByVal openingBTopSettings As Variant, ByVal openingBBottomSettings As Variant, _
        ByVal steelClass As String) As CRebarLayout
    If geom Is Nothing Then Err.Raise vbObjectError + 3242, "CHollowRectRebarLayoutBuilder", "Геометрия HollowRectangle не передана."

    Dim message As String
    If Not geom.IsValid(message) Then Err.Raise vbObjectError + 3243, "CHollowRectRebarLayoutBuilder", message

    Dim layout As CRebarLayout
    Set layout = New CRebarLayout

    Dim nextIndex As Long
    nextIndex = 1

    Dim topX() As Double, topY() As Double
    Dim bottomX() As Double, bottomY() As Double
    Dim leftX() As Double, leftY() As Double
    Dim rightX() As Double, rightY() As Double
    Dim topNX() As Double, topNY() As Double
    Dim bottomNX() As Double, bottomNY() As Double
    Dim leftNX() As Double, leftNY() As Double
    Dim rightNX() As Double, rightNY() As Double
    Dim topCount As Long, bottomCount As Long, leftCount As Long, rightCount As Long

    ' Сначала создаем наружный ряд: Opening потом берет от него координаты
    ' и проектирует их на свои линии раскладки.
    AddHorizontalFace layout, geom, "B.Top", False, True, bTopSettings, steelClass, nextIndex, topX, topY, topNX, topNY, topCount
    AddHorizontalFace layout, geom, "B.Bottom", False, False, bBottomSettings, steelClass, nextIndex, bottomX, bottomY, bottomNX, bottomNY, bottomCount
    AddSideFace layout, geom, "H.Left", False, True, hLeftSettings, steelClass, nextIndex, leftX, leftY, leftNX, leftNY, leftCount
    AddSideFace layout, geom, "H.Right", False, False, hRightSettings, steelClass, nextIndex, rightX, rightY, rightNX, rightNY, rightCount

    AddProjectedHorizontalFace layout, geom, "Opening.B.Top", True, openingBTopSettings, _
        openingHLeftSettings, openingHRightSettings, topX, topY, topNX, topNY, topCount, steelClass, nextIndex
    AddProjectedHorizontalFace layout, geom, "Opening.B.Bottom", False, openingBBottomSettings, _
        openingHLeftSettings, openingHRightSettings, bottomX, bottomY, bottomNX, bottomNY, bottomCount, steelClass, nextIndex
    AddProjectedSideFace layout, geom, "Opening.H.Left", True, openingHLeftSettings, _
        openingBTopSettings, openingBBottomSettings, leftX, leftY, leftNX, leftNY, leftCount, steelClass, nextIndex
    AddProjectedSideFace layout, geom, "Opening.H.Right", False, openingHRightSettings, _
        openingBTopSettings, openingBBottomSettings, rightX, rightY, rightNX, rightNY, rightCount, steelClass, nextIndex

    Set Build = layout
End Function

' Читает компактную строку параметров одной смысловой грани.
Private Function ReadFaceSettingsFromConfig(ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem, _
        ByVal faceName As String, ByVal defaultAs As String, ByVal defaultD As String) As Variant
    ReadFaceSettingsFromConfig = Array( _
        units.InputLengthToInternal(settings.GetDouble("HollowRectangle." & faceName & ".as", CDbl(defaultAs))), _
        OptionalDiameterFromSettings(settings, units, "HollowRectangle." & faceName & ".d", CDbl(defaultD)), _
        settings.GetLong("HollowRectangle." & faceName & ".n", 0), _
        OptionalDiameterFromSettings(settings, units, "HollowRectangle." & faceName & ".d_2row", 0#), _
        OptionalDiameterFromSettings(settings, units, "HollowRectangle." & faceName & ".d_3row", 0#), _
        settings.GetString("HollowRectangle." & faceName & ".loc_2row", "Stacked"), _
        settings.GetString("HollowRectangle." & faceName & ".loc_3row", "Stacked"), _
        settings.GetString("HollowRectangle." & faceName & ".bind_2row", "EachBar"), _
        settings.GetString("HollowRectangle." & faceName & ".bind_3row", "EachBar"))
End Function

' Пустой диаметр в таблице означает, что ряд выключен.
Private Function OptionalDiameterFromSettings(ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem, _
        ByVal key As String, ByVal defaultValue As Double) As Double
    If Not settings.HasKey(key) Then
        OptionalDiameterFromSettings = units.InputLengthToInternal(defaultValue)
        Exit Function
    End If

    Dim rawValue As String
    rawValue = Trim$(settings.GetRawString(key, vbNullString))
    If Len(rawValue) = 0 Or UCase$(rawValue) = "TODO" Or rawValue = "-" Then
        OptionalDiameterFromSettings = 0#
    Else
        OptionalDiameterFromSettings = units.InputLengthToInternal(settings.GetDouble(key, defaultValue))
    End If
End Function

' Добавляет верхнюю или нижнюю наружную грань B. Для HollowRectangle это не
' только прямой горизонтальный участок: линия раскладки идет по верхней или
' нижней траектории с учетом радиусов сопряжения.
Private Sub AddHorizontalFace(ByVal layout As CRebarLayout, ByVal geom As CGeometryHollowRectangle, _
        ByVal faceName As String, ByVal openingFace As Boolean, ByVal isTop As Boolean, _
        ByVal settings As Variant, ByVal steelClass As String, ByRef nextIndex As Long, _
        ByRef barX() As Double, ByRef barY() As Double, ByRef barNormalX() As Double, _
        ByRef barNormalY() As Double, ByRef barCountOut As Long)
    Dim axisDistance As Double, diameter As Double, barCount As Long
    axisDistance = FaceAs(settings)
    diameter = FaceDiameter(settings)
    barCount = FaceCount(settings)
    ValidateFace faceName, axisDistance, diameter, barCount, settings
    If barCount = 0 Or diameter = 0# Then Exit Sub

    Dim px() As Double, py() As Double
    geom.GetHorizontalRebarPath openingFace, isTop, axisDistance, px, py

    Dim lengthValue As Double
    lengthValue = PolylineLength(px, py)
    If lengthValue <= GEOM_TOLERANCE Then Err.Raise vbObjectError + 3244, "CHollowRectRebarLayoutBuilder", _
        HollowFaceDisplayName(faceName) & ": длина линии раскладки должна быть больше нуля."

    Dim normalX As Double, normalY As Double
    HorizontalAnnotationNormal openingFace, isTop, normalX, normalY

    Dim centroidAvailable As Boolean
    Dim referenceX As Double, referenceY As Double
    geom.AnalyticalCentroid centroidAvailable, referenceX, referenceY
    If Not centroidAvailable Then referenceY = 0#

    Dim spacing As Double
    spacing = lengthValue / CDbl(barCount + 1)
    Dim anchorStartX As Double, anchorStartY As Double
    Dim anchorEndX As Double, anchorEndY As Double
    Dim anchorProjection As Double
    Dim anchorInitialized As Boolean

    Dim i As Long
    For i = 1 To barCount
        Dim xCoord As Double, yCoord As Double, tx As Double, ty As Double
        PointAtDistance px, py, spacing * CDbl(i), xCoord, yCoord, tx, ty
        Dim inwardX As Double, inwardY As Double
        ConcreteInwardVector geom, xCoord, yCoord, tx, ty, axisDistance, inwardX, inwardY

        If AddBarIfAvailable(layout, geom, faceName, 1, xCoord, yCoord, diameter, steelClass, nextIndex) Then
            AccumulateBLabelAnchor isTop, referenceY, xCoord, yCoord, _
                anchorStartX, anchorStartY, anchorEndX, anchorEndY, anchorProjection, anchorInitialized
            RecordBarProjection barX, barY, barNormalX, barNormalY, barCountOut, xCoord, yCoord, inwardX, inwardY
            AddAdditionalRows layout, geom, faceName, xCoord, yCoord, diameter, settings, i, tx, ty, inwardX, inwardY, steelClass, nextIndex
        End If
    Next i
    If anchorInitialized Then layout.AddAnnotationAnchor faceName, anchorStartX, anchorStartY, anchorEndX, anchorEndY, normalX, normalY, axisDistance
End Sub

' Добавляет левую или правую наружную грань H. Для HollowRectangle H-линия
' является прямым вертикальным участком между верхним и нижним сопряжением.
Private Sub AddSideFace(ByVal layout As CRebarLayout, ByVal geom As CGeometryHollowRectangle, _
        ByVal faceName As String, ByVal openingFace As Boolean, ByVal isLeft As Boolean, _
        ByVal settings As Variant, ByVal steelClass As String, ByRef nextIndex As Long, _
        ByRef barX() As Double, ByRef barY() As Double, ByRef barNormalX() As Double, _
        ByRef barNormalY() As Double, ByRef barCountOut As Long)
    Dim axisDistance As Double, diameter As Double, barCount As Long
    axisDistance = FaceAs(settings)
    diameter = FaceDiameter(settings)
    barCount = FaceCount(settings)
    ValidateFace faceName, axisDistance, diameter, barCount, settings
    If barCount = 0 Or diameter = 0# Then Exit Sub

    Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
    Dim normalX As Double, normalY As Double
    geom.GetVerticalRebarLine openingFace, isLeft, axisDistance, x1, y1, x2, y2, normalX, normalY

    Dim totalLength As Double
    totalLength = Sqr((x2 - x1) ^ 2 + (y2 - y1) ^ 2)
    If totalLength <= GEOM_TOLERANCE Then Err.Raise vbObjectError + 3245, "CHollowRectRebarLayoutBuilder", _
        HollowFaceDisplayName(faceName) & ": длина боковой линии раскладки должна быть больше нуля."

    layout.AddAnnotationAnchor faceName, x1, y1, x2, y2, normalX, normalY, axisDistance

    Dim tx As Double, ty As Double
    tx = (x2 - x1) / totalLength
    ty = (y2 - y1) / totalLength
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

        If AddBarIfAvailable(layout, geom, faceName, 1, xCoord, yCoord, diameter, steelClass, nextIndex) Then
            RecordBarProjection barX, barY, barNormalX, barNormalY, barCountOut, xCoord, yCoord, inwardX, inwardY
            AddAdditionalRows layout, geom, faceName, xCoord, yCoord, diameter, settings, i, tx, ty, inwardX, inwardY, steelClass, nextIndex
        End If
    Next i
End Sub

' Строит внутреннюю горизонтальную B-грань Opening по X-проекциям наружных
' B-стержней на криволинейную внутреннюю траекторию.
Private Sub AddProjectedHorizontalFace(ByVal layout As CRebarLayout, ByVal geom As CGeometryHollowRectangle, _
        ByVal faceName As String, ByVal isTop As Boolean, ByVal settings As Variant, _
        ByVal openingHLeftSettings As Variant, ByVal openingHRightSettings As Variant, _
        ByRef sourceX() As Double, ByRef sourceY() As Double, ByRef sourceNormalX() As Double, _
        ByRef sourceNormalY() As Double, ByVal sourceCount As Long, ByVal steelClass As String, ByRef nextIndex As Long)
    Dim axisDistance As Double, diameter As Double
    axisDistance = FaceAs(settings)
    diameter = FaceDiameter(settings)
    ValidateOpeningFace faceName, axisDistance, diameter, settings
    If sourceCount = 0 Or diameter = 0# Then Exit Sub

    Dim normalX As Double, normalY As Double
    HorizontalAnnotationNormal True, isTop, normalX, normalY

    Dim straightOpeningLine As Boolean
    straightOpeningLine = (geom.OpeningR <= GEOM_TOLERANCE)

    Dim px() As Double, py() As Double
    Dim lineX1 As Double, lineY1 As Double, lineX2 As Double, lineY2 As Double
    Dim lineNormalX As Double, lineNormalY As Double
    Dim lineLength As Double
    If straightOpeningLine Then
        ' При R_o=0 смысловая B-линия Opening - это один прямой offset-отрезок.
        ' Лучи крайних наружных стержней не должны цепляться за соседние H-сегменты
        ' прямоугольного offset-контура.
        geom.GetHorizontalRebarLine True, isTop, axisDistance, _
            lineX1, lineY1, lineX2, lineY2, lineNormalX, lineNormalY
        ExtendOpeningBLineBySideCover openingHLeftSettings, openingHRightSettings, lineX1, lineX2
        lineLength = Sqr((lineX2 - lineX1) ^ 2 + (lineY2 - lineY1) ^ 2)
        If lineLength <= GEOM_TOLERANCE Then Exit Sub
    Else
        geom.GetHorizontalRebarPath True, isTop, axisDistance, px, py
        lineLength = PolylineLength(px, py)
        If lineLength <= GEOM_TOLERANCE Then Exit Sub
    End If

    Dim centroidAvailable As Boolean
    Dim referenceX As Double, referenceY As Double
    geom.AnalyticalCentroid centroidAvailable, referenceX, referenceY
    If Not centroidAvailable Then referenceY = 0#

    Dim i As Long
    Dim ordinal As Long
    Dim anchorStartX As Double, anchorStartY As Double
    Dim anchorEndX As Double, anchorEndY As Double
    Dim anchorProjection As Double
    Dim anchorInitialized As Boolean
    For i = 1 To sourceCount
        Dim xCoord As Double, yCoord As Double, tx As Double, ty As Double
        Dim projected As Boolean
        If straightOpeningLine Then
            projected = ProjectRayToSegment(sourceX(i), sourceY(i), sourceNormalX(i), sourceNormalY(i), _
                lineX1, lineY1, lineX2, lineY2, xCoord, yCoord)
            If projected Then
                tx = (lineX2 - lineX1) / lineLength
                ty = (lineY2 - lineY1) / lineLength
            End If
        Else
            projected = ProjectRayToPolyline(sourceX(i), sourceY(i), sourceNormalX(i), sourceNormalY(i), px, py, xCoord, yCoord, tx, ty)
        End If

        If projected Then
            ordinal = ordinal + 1
            Dim inwardX As Double, inwardY As Double
            ConcreteInwardVector geom, xCoord, yCoord, tx, ty, axisDistance, inwardX, inwardY
            If AddBarIfAvailable(layout, geom, faceName, 1, xCoord, yCoord, diameter, steelClass, nextIndex) Then
                AccumulateBLabelAnchor isTop, referenceY, xCoord, yCoord, _
                    anchorStartX, anchorStartY, anchorEndX, anchorEndY, anchorProjection, anchorInitialized
                AddAdditionalRows layout, geom, faceName, xCoord, yCoord, diameter, settings, ordinal, tx, ty, inwardX, inwardY, steelClass, nextIndex
            End If
        End If
    Next i
    If anchorInitialized Then layout.AddAnnotationAnchor faceName, anchorStartX, anchorStartY, anchorEndX, anchorEndY, normalX, normalY, axisDistance
End Sub

' Строит внутреннюю вертикальную H-грань Opening по Y-проекциям наружных
' H-стержней на прямую внутреннюю линию.
Private Sub AddProjectedSideFace(ByVal layout As CRebarLayout, ByVal geom As CGeometryHollowRectangle, _
        ByVal faceName As String, ByVal isLeft As Boolean, ByVal settings As Variant, _
        ByVal openingBTopSettings As Variant, ByVal openingBBottomSettings As Variant, _
        ByRef sourceX() As Double, ByRef sourceY() As Double, ByRef sourceNormalX() As Double, _
        ByRef sourceNormalY() As Double, ByVal sourceCount As Long, ByVal steelClass As String, ByRef nextIndex As Long)
    Dim axisDistance As Double, diameter As Double
    axisDistance = FaceAs(settings)
    diameter = FaceDiameter(settings)
    ValidateOpeningFace faceName, axisDistance, diameter, settings
    If sourceCount = 0 Or diameter = 0# Then Exit Sub

    Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
    Dim normalX As Double, normalY As Double
    geom.GetVerticalRebarLine True, isLeft, axisDistance, x1, y1, x2, y2, normalX, normalY
    If geom.OpeningR <= GEOM_TOLERANCE Then
        ExtendOpeningHLineByTopBottomCover openingBTopSettings, openingBBottomSettings, y1, y2
    End If
    layout.AddAnnotationAnchor faceName, x1, y1, x2, y2, normalX, normalY, axisDistance

    Dim lengthValue As Double
    lengthValue = Sqr((x2 - x1) ^ 2 + (y2 - y1) ^ 2)
    If lengthValue <= GEOM_TOLERANCE Then Exit Sub
    Dim tx As Double, ty As Double
    tx = (x2 - x1) / lengthValue
    ty = (y2 - y1) / lengthValue
    Dim inwardX As Double, inwardY As Double
    inwardX = -normalX
    inwardY = -normalY

    Dim i As Long
    Dim ordinal As Long
    For i = 1 To sourceCount
        Dim xCoord As Double, yCoord As Double
        If ProjectRayToSegment(sourceX(i), sourceY(i), sourceNormalX(i), sourceNormalY(i), x1, y1, x2, y2, xCoord, yCoord) Then
            ordinal = ordinal + 1
            If AddBarIfAvailable(layout, geom, faceName, 1, xCoord, yCoord, diameter, steelClass, nextIndex) Then
                AddAdditionalRows layout, geom, faceName, xCoord, yCoord, diameter, settings, ordinal, tx, ty, inwardX, inwardY, steelClass, nextIndex
            End If
        End If
    Next i
End Sub

' Добавляет второй и третий ряды относительно найденного первого ряда.
Private Sub AddAdditionalRows(ByVal layout As CRebarLayout, ByVal geom As CGeometryHollowRectangle, _
        ByVal faceName As String, ByVal baseX As Double, ByVal baseY As Double, ByVal row1Diameter As Double, _
        ByVal settings As Variant, ByVal baseBarOrdinal As Long, ByVal tangentX As Double, ByVal tangentY As Double, _
        ByVal inwardX As Double, ByVal inwardY As Double, ByVal steelClass As String, ByRef nextIndex As Long)
    Dim row2Diameter As Double, row3Diameter As Double
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

' Вычисляет координату стержня второго или третьего ряда относительно стержня
' первого ряда. Stacked уводит ряд в бетон по нормали, SideBySide - вдоль
' касательной к линии раскладки; skippedDiameter нужен, когда третий ряд
' должен "перепрыгнуть" уже существующий второй ряд.
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

' Добавляет стержень в layout только если в этой же точке еще нет стержня.
' Это защищает углы и пересечения проекций Opening от дублей: один и тот же
' физический стержень может быть достигнут из соседних смысловых граней.
Private Function AddBarIfAvailable(ByVal layout As CRebarLayout, ByVal geom As CGeometryHollowRectangle, _
        ByVal faceName As String, ByVal rowNumber As Long, ByVal xCoord As Double, ByVal yCoord As Double, _
        ByVal diameter As Double, ByVal steelClass As String, ByRef nextIndex As Long) As Boolean
    If HasBarAt(layout, xCoord, yCoord) Then Exit Function
    layout.AddBar faceName & ".row_" & CStr(rowNumber) & "-" & CStr(nextIndex), _
        xCoord, yCoord, diameter, 0#, steelClass, _
        "HollowRectangle: " & HollowFaceDisplayName(faceName) & ", ряд " & CStr(rowNumber), geom, faceName
    nextIndex = nextIndex + 1
    AddBarIfAvailable = True
End Function

' Проверяет параметры наружной грани, где пользователь явно задает количество
' стержней n. Нулевая грань разрешена и просто пропускается, но отрицательное
' количество или некорректные диаметры/ряды считаются ошибкой ввода.
Private Sub ValidateFace(ByVal faceName As String, ByVal axisDistance As Double, ByVal diameter As Double, _
        ByVal barCount As Long, ByVal settings As Variant)
    If barCount < 0 Then Err.Raise vbObjectError + 3250, "CHollowRectRebarLayoutBuilder", HollowFaceDisplayName(faceName) & ": количество стержней n должно быть нулевым или положительным."
    ValidateOpeningFace faceName, axisDistance, diameter, settings
    If barCount = 0 Or diameter = 0# Then Exit Sub
End Sub

' Проверяет общие параметры одной грани: as, диаметр первого ряда и параметры
' дополнительных рядов. Метод используется и для Opening, где количество
' стержней не задается пользователем, а получается проекцией наружной грани.
Private Sub ValidateOpeningFace(ByVal faceName As String, ByVal axisDistance As Double, ByVal diameter As Double, ByVal settings As Variant)
    If diameter < 0# Then Err.Raise vbObjectError + 3251, "CHollowRectRebarLayoutBuilder", HollowFaceDisplayName(faceName) & ": диаметр d стержней первого ряда не может быть отрицательным."
    If diameter = 0# Then Exit Sub
    If axisDistance <= diameter / 2# Then Err.Raise vbObjectError + 3252, "CHollowRectRebarLayoutBuilder", HollowFaceDisplayName(faceName) & ": расстояние as до оси стержня должно быть больше d/2."
    If FaceRow2Diameter(settings) < 0# Or FaceRow3Diameter(settings) < 0# Then Err.Raise vbObjectError + 3253, "CHollowRectRebarLayoutBuilder", HollowFaceDisplayName(faceName) & ": диаметры дополнительных рядов d2 и d3 должны быть неотрицательными."
    If FaceRow2Diameter(settings) > 0# And Not IsSupportedRowLocation(FaceRow2Location(settings)) Then Err.Raise vbObjectError + 3254, "CHollowRectRebarLayoutBuilder", HollowFaceDisplayName(faceName) & ": положение второго ряда должно быть SideBySide или Stacked."
    If FaceRow3Diameter(settings) > 0# And Not IsSupportedRowLocation(FaceRow3Location(settings)) Then Err.Raise vbObjectError + 3255, "CHollowRectRebarLayoutBuilder", HollowFaceDisplayName(faceName) & ": положение третьего ряда должно быть SideBySide или Stacked."
    If FaceRow2Diameter(settings) > 0# And Not IsSupportedRowBinding(FaceRow2Binding(settings)) Then Err.Raise vbObjectError + 3256, "CHollowRectRebarLayoutBuilder", HollowFaceDisplayName(faceName) & ": привязка второго ряда должна быть EachBar или EverySecondBar."
    If FaceRow3Diameter(settings) > 0# And Not IsSupportedRowBinding(FaceRow3Binding(settings)) Then Err.Raise vbObjectError + 3257, "CHollowRectRebarLayoutBuilder", HollowFaceDisplayName(faceName) & ": привязка третьего ряда должна быть EachBar или EverySecondBar."
End Sub

' Возвращает True только для поддерживаемых режимов положения дополнительного
' ряда: рядом по касательной или внутрь бетона по нормали.
Private Function IsSupportedRowLocation(ByVal value As String) As Boolean
    IsSupportedRowLocation = (StrComp(value, "SideBySide", vbTextCompare) = 0 Or StrComp(value, "Stacked", vbTextCompare) = 0)
End Function

' Возвращает True только для поддерживаемых режимов привязки дополнительного
' ряда к первому ряду: к каждому стержню или через один.
Private Function IsSupportedRowBinding(ByVal value As String) As Boolean
    IsSupportedRowBinding = (StrComp(value, "EachBar", vbTextCompare) = 0 Or StrComp(value, "EverySecondBar", vbTextCompare) = 0)
End Function

' Решает, создавать ли дополнительный стержень для текущего стержня первого
' ряда. Для EverySecondBar используются не индексы layout, а порядковый номер
' базового стержня на этой конкретной смысловой грани.
Private Function ShouldAddAdditionalRow(ByVal baseBarOrdinal As Long, ByVal bindingMode As String) As Boolean
    If StrComp(bindingMode, "EverySecondBar", vbTextCompare) = 0 Then
        ShouldAddAdditionalRow = ((baseBarOrdinal Mod 2) = 1)
    Else
        ShouldAddAdditionalRow = True
    End If
End Function

' Запоминает координаты и нормаль уже созданного наружного стержня. Позже эти
' массивы используются, чтобы проецировать наружную раскладку на внутренний
' контур Opening и автоматически получать число внутренних стержней.
Private Sub RecordBarProjection(ByRef barX() As Double, ByRef barY() As Double, _
        ByRef barNormalX() As Double, ByRef barNormalY() As Double, ByRef barCount As Long, _
        ByVal xCoord As Double, ByVal yCoord As Double, ByVal normalX As Double, ByVal normalY As Double)
    barCount = barCount + 1
    ReDim Preserve barX(1 To barCount)
    ReDim Preserve barY(1 To barCount)
    ReDim Preserve barNormalX(1 To barCount)
    ReDim Preserve barNormalY(1 To barCount)
    barX(barCount) = xCoord
    barY(barCount) = yCoord
    barNormalX(barCount) = normalX
    barNormalY(barCount) = normalY
End Sub

' Продлевает прямую B-линию Opening до углов offset-прямоугольника. При R_o=0
' внутренняя горизонтальная линия смещается от отверстия на свой as_B, но ее
' концы должны доходить до вертикальных offset-линий Opening.H.Left/Right,
' то есть дополнительно уходить на соответствующие as_H влево и вправо.
Private Sub ExtendOpeningBLineBySideCover(ByVal openingHLeftSettings As Variant, ByVal openingHRightSettings As Variant, _
        ByRef lineX1 As Double, ByRef lineX2 As Double)
    Dim leftCover As Double
    Dim rightCover As Double
    leftCover = FaceAs(openingHLeftSettings)
    rightCover = FaceAs(openingHRightSettings)

    If lineX1 <= lineX2 Then
        lineX1 = lineX1 - leftCover
        lineX2 = lineX2 + rightCover
    Else
        lineX1 = lineX1 + rightCover
        lineX2 = lineX2 - leftCover
    End If
End Sub

' Продлевает прямую H-линию Opening до углов offset-прямоугольника. При R_o=0
' внутренняя вертикальная линия смещается от отверстия на свой as_H, а ее
' концы доходят до горизонтальных offset-линий Opening.B.Top/Bottom.
Private Sub ExtendOpeningHLineByTopBottomCover(ByVal openingBTopSettings As Variant, ByVal openingBBottomSettings As Variant, _
        ByRef lineY1 As Double, ByRef lineY2 As Double)
    Dim topCover As Double
    Dim bottomCover As Double
    topCover = FaceAs(openingBTopSettings)
    bottomCover = FaceAs(openingBBottomSettings)

    If lineY1 <= lineY2 Then
        lineY1 = lineY1 - bottomCover
        lineY2 = lineY2 + topCover
    Else
        lineY1 = lineY1 + topCover
        lineY2 = lineY2 - bottomCover
    End If
End Sub

' Собирает anchor подписи B-группы по стержням, наиболее удаленным от центра
' сечения по вертикали. Так Offset подписи отсчитывается от дальнего видимого
' стержня, а не от хорды всей криволинейной траектории раскладки.
Private Sub AccumulateBLabelAnchor(ByVal isTop As Boolean, ByVal referenceY As Double, _
        ByVal xCoord As Double, ByVal yCoord As Double, ByRef startX As Double, ByRef startY As Double, _
        ByRef endX As Double, ByRef endY As Double, ByRef bestProjection As Double, ByRef initialized As Boolean)
    Dim projection As Double
    If isTop Then
        projection = yCoord - referenceY
    Else
        projection = referenceY - yCoord
    End If

    If Not initialized Or projection > bestProjection + GEOM_TOLERANCE Then
        startX = xCoord
        endX = xCoord
        startY = yCoord
        endY = yCoord
        bestProjection = projection
        initialized = True
    ElseIf Abs(projection - bestProjection) <= GEOM_TOLERANCE Then
        If xCoord < startX Then startX = xCoord
        If xCoord > endX Then endX = xCoord
        If isTop Then
            If yCoord > startY Then startY = yCoord
        Else
            If yCoord < startY Then startY = yCoord
        End If
        endY = startY
    End If
End Sub

' Пересекает луч из наружного стержня по нормали к его линии раскладки с
' криволинейной B-траекторией Opening. Если нормаль проходит мимо offset-линии,
' внутренний стержень для этого наружного стержня не создается.
Private Function ProjectRayToPolyline(ByVal sourceX As Double, ByVal sourceY As Double, _
        ByVal dirX As Double, ByVal dirY As Double, ByRef px() As Double, ByRef py() As Double, _
        ByRef xCoord As Double, ByRef yCoord As Double, ByRef tangentX As Double, ByRef tangentY As Double) As Boolean
    Dim i As Long
    Dim bestRayT As Double
    bestRayT = 0#
    For i = LBound(px) To UBound(px) - 1
        Dim hitX As Double, hitY As Double, hitT As Double, segT As Double
        If IntersectRayWithSegment(sourceX, sourceY, dirX, dirY, px(i), py(i), px(i + 1), py(i + 1), hitX, hitY, hitT, segT) Then
            If Not ProjectRayToPolyline Or hitT < bestRayT Then
                bestRayT = hitT
                xCoord = hitX
                yCoord = hitY
                Dim segLen As Double
                segLen = Sqr((px(i + 1) - px(i)) ^ 2 + (py(i + 1) - py(i)) ^ 2)
                If segLen > GEOM_TOLERANCE Then
                    tangentX = (px(i + 1) - px(i)) / segLen
                    tangentY = (py(i + 1) - py(i)) / segLen
                End If
                ProjectRayToPolyline = True
            End If
        End If
    Next i
End Function

' Пересекает нормаль наружного H-стержня с прямой H-линией Opening.
Private Function ProjectRayToSegment(ByVal sourceX As Double, ByVal sourceY As Double, _
        ByVal dirX As Double, ByVal dirY As Double, ByVal x1 As Double, ByVal y1 As Double, _
        ByVal x2 As Double, ByVal y2 As Double, ByRef xCoord As Double, ByRef yCoord As Double) As Boolean
    Dim rayT As Double, segT As Double
    ProjectRayToSegment = IntersectRayWithSegment(sourceX, sourceY, dirX, dirY, x1, y1, x2, y2, xCoord, yCoord, rayT, segT)
End Function

' Общая геометрия пересечения: луч source + t*dir, t>=0, и отрезок a+s*(b-a),
' 0<=s<=1. Возвращает только реальные пересечения, без "дотягивания" до линии.
Private Function IntersectRayWithSegment(ByVal sourceX As Double, ByVal sourceY As Double, _
        ByVal dirX As Double, ByVal dirY As Double, ByVal x1 As Double, ByVal y1 As Double, _
        ByVal x2 As Double, ByVal y2 As Double, ByRef hitX As Double, ByRef hitY As Double, _
        ByRef rayT As Double, ByRef segmentT As Double) As Boolean
    Dim segX As Double, segY As Double
    segX = x2 - x1
    segY = y2 - y1

    Dim denom As Double
    denom = dirX * segY - dirY * segX
    If Abs(denom) <= GEOM_TOLERANCE Then Exit Function

    Dim dx As Double, dy As Double
    dx = x1 - sourceX
    dy = y1 - sourceY
    rayT = (dx * segY - dy * segX) / denom
    segmentT = (dx * dirY - dy * dirX) / denom

    If rayT < -GEOM_TOLERANCE Then Exit Function
    If segmentT < -GEOM_TOLERANCE Or segmentT > 1# + GEOM_TOLERANCE Then Exit Function
    If rayT < 0# Then rayT = 0#
    If segmentT < 0# Then segmentT = 0#
    If segmentT > 1# Then segmentT = 1#

    hitX = sourceX + dirX * rayT
    hitY = sourceY + dirY * rayT
    IntersectRayWithSegment = True
End Function

' Возвращает нормаль для подписи B-группы. Для Opening нормаль направлена в
' сторону пустоты, чтобы режим Placement работал относительно внутренней грани.
Private Sub HorizontalAnnotationNormal(ByVal openingFace As Boolean, ByVal isTop As Boolean, _
        ByRef normalX As Double, ByRef normalY As Double)
    normalX = 0#
    If openingFace Then
        If isTop Then normalY = -1# Else normalY = 1#
    Else
        If isTop Then normalY = 1# Else normalY = -1#
    End If
End Sub

' Выбирает нормаль второго/третьего ряда в сторону бетона. Это надежнее, чем
' вручную угадывать знак для верхней/нижней и внутренней/наружной кривой:
' контрольная точка должна оставаться внутри бетонного кольца.
Private Sub ConcreteInwardVector(ByVal geom As CGeometryHollowRectangle, ByVal xCoord As Double, ByVal yCoord As Double, _
        ByVal tangentX As Double, ByVal tangentY As Double, ByVal axisDistance As Double, _
        ByRef inwardX As Double, ByRef inwardY As Double)
    Dim nx1 As Double, ny1 As Double
    Dim nx2 As Double, ny2 As Double
    nx1 = -tangentY: ny1 = tangentX
    nx2 = tangentY: ny2 = -tangentX
    NormalizeVector nx1, ny1
    NormalizeVector nx2, ny2

    ' Пробуем точку глубже, чем собственный as: стержень уже стоит внутри бетона,
    ' поэтому проверка на 1 мм могла принять наружное направление за внутреннее.
    Dim probe As Double
    probe = WorksheetMax(1#, Abs(axisDistance) * 1.25)
    If geom.ContainsPoint(xCoord + nx1 * probe, yCoord + ny1 * probe) Then
        inwardX = nx1
        inwardY = ny1
    Else
        inwardX = nx2
        inwardY = ny2
    End If
End Sub

' Считает суммарную длину ломаной траектории раскладки. Дуги к этому моменту
' уже дискретизированы геометрией в последовательность точек, поэтому здесь
' используется обычная сумма длин соседних отрезков.
Private Function PolylineLength(ByRef px() As Double, ByRef py() As Double) As Double
    Dim i As Long
    For i = LBound(px) To UBound(px) - 1
        PolylineLength = PolylineLength + Sqr((px(i + 1) - px(i)) ^ 2 + (py(i + 1) - py(i)) ^ 2)
    Next i
End Function

' Возвращает точку на ломаной на заданном расстоянии от начала и касательный
' единичный вектор в этой точке. Метод нужен для равномерной раскладки по
' криволинейным B-граням, где простой линейной интерполяции по X/Y недостаточно.
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

' Нормирует двумерный вектор на месте. Нулевой или почти нулевой вектор не
' меняется, чтобы не создавать деление на численный шум.
Private Sub NormalizeVector(ByRef xValue As Double, ByRef yValue As Double)
    Dim lenValue As Double
    lenValue = Sqr(xValue * xValue + yValue * yValue)
    If lenValue <= GEOM_TOLERANCE Then Exit Sub
    xValue = xValue / lenValue
    yValue = yValue / lenValue
End Sub

' Проверяет, есть ли уже стержень с такими координатами. Используется перед
' добавлением угловых и проецируемых стержней, чтобы разные грани не создавали
' совпадающие расчетные элементы арматуры.
Private Function HasBarAt(ByVal layout As CRebarLayout, ByVal xCoord As Double, ByVal yCoord As Double) As Boolean
    Dim i As Long
    For i = 1 To layout.Count
        If Abs(layout.X(i) - xCoord) <= GEOM_TOLERANCE And Abs(layout.Y(i) - yCoord) <= GEOM_TOLERANCE Then
            HasBarAt = True
            Exit Function
        End If
    Next i
End Function

' Достает as из компактного массива настроек одной грани.
Private Function FaceAs(ByVal settings As Variant) As Double
    FaceAs = CDbl(settings(LBound(settings)))
End Function

' Достает диаметр первого ряда из компактного массива настроек одной грани.
Private Function FaceDiameter(ByVal settings As Variant) As Double
    FaceDiameter = CDbl(settings(LBound(settings) + 1))
End Function

' Достает заданное пользователем количество стержней первого ряда.
Private Function FaceCount(ByVal settings As Variant) As Long
    FaceCount = CLng(settings(LBound(settings) + 2))
End Function

' Достает диаметр второго ряда; ноль означает, что ряд выключен.
Private Function FaceRow2Diameter(ByVal settings As Variant) As Double
    FaceRow2Diameter = CDbl(settings(LBound(settings) + 3))
End Function

' Достает диаметр третьего ряда; ноль означает, что ряд выключен.
Private Function FaceRow3Diameter(ByVal settings As Variant) As Double
    FaceRow3Diameter = CDbl(settings(LBound(settings) + 4))
End Function

' Достает режим положения второго ряда относительно первого.
Private Function FaceRow2Location(ByVal settings As Variant) As String
    FaceRow2Location = CStr(settings(LBound(settings) + 5))
End Function

' Достает режим положения третьего ряда относительно первого.
Private Function FaceRow3Location(ByVal settings As Variant) As String
    FaceRow3Location = CStr(settings(LBound(settings) + 6))
End Function

' Достает режим привязки второго ряда к стержням первого ряда.
Private Function FaceRow2Binding(ByVal settings As Variant) As String
    FaceRow2Binding = CStr(settings(LBound(settings) + 7))
End Function

' Достает режим привязки третьего ряда к стержням первого ряда.
Private Function FaceRow3Binding(ByVal settings As Variant) As String
    FaceRow3Binding = CStr(settings(LBound(settings) + 8))
End Function

' Локальный минимум без обращения к WorksheetFunction в горячей геометрии.
Private Function WorksheetMin(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then WorksheetMin = a Else WorksheetMin = b
End Function

' Локальный максимум без обращения к WorksheetFunction в горячей геометрии.
Private Function WorksheetMax(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then WorksheetMax = a Else WorksheetMax = b
End Function

' Переводит внутренний ключ грани в видимое имя для ошибок и комментариев.
Private Function HollowFaceDisplayName(ByVal faceName As String) As String
    Select Case LCase$(Trim$(faceName))
        Case "h.left": HollowFaceDisplayName = "H - левая наружная грань"
        Case "h.right": HollowFaceDisplayName = "H - правая наружная грань"
        Case "b.top": HollowFaceDisplayName = "B - верхняя наружная грань"
        Case "b.bottom": HollowFaceDisplayName = "B - нижняя наружная грань"
        Case "opening.h.left": HollowFaceDisplayName = "H_o - левая грань Opening"
        Case "opening.h.right": HollowFaceDisplayName = "H_o - правая грань Opening"
        Case "opening.b.top": HollowFaceDisplayName = "B_o - верхняя грань Opening"
        Case "opening.b.bottom": HollowFaceDisplayName = "B_o - нижняя грань Opening"
        Case Else: HollowFaceDisplayName = faceName
    End Select
End Function




