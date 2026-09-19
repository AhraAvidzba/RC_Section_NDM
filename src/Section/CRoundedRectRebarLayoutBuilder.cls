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
' Builder читает новую таблицу скругленного прямоугольника: общие размеры B/H,
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
    geom.InitializeSides _
        units.InputLengthToInternal(settings.GetDouble("RoundedRectangle.B", 300#)), _
        units.InputLengthToInternal(settings.GetDouble("RoundedRectangle.H", 200#)), _
        settings.GetString("RoundedRectangle.Left.Type", "Simple"), _
        settings.GetString("RoundedRectangle.Right.Type", "Simple"), _
        units.InputLengthToInternal(settings.GetDouble("RoundedRectangle.Left.W", 0#)), _
        units.InputLengthToInternal(settings.GetDouble("RoundedRectangle.Right.W", 0#)), _
        units.InputLengthToInternal(settings.GetDouble("RoundedRectangle.Left.R1", 0#)), _
        units.InputLengthToInternal(settings.GetDouble("RoundedRectangle.Right.R1", 0#)), _
        units.InputLengthToInternal(settings.GetDouble("RoundedRectangle.Left.R2", 0#)), _
        units.InputLengthToInternal(settings.GetDouble("RoundedRectangle.Right.R2", 0#))

    Set BuildFromSettings = Build(geom, _
        ReadFaceSettingsFromConfig(settings, units, "H.Left", "40", "32", "5"), _
        ReadFaceSettingsFromConfig(settings, units, "H.Right", "40", "32", "5"), _
        ReadFaceSettingsFromConfig(settings, units, "B.Top", "40", "32", "8"), _
        ReadFaceSettingsFromConfig(settings, units, "B.Bottom", "40", "32", "8"), _
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

' Читает параметры одной строки арматуры RoundedRectangle.
Private Function ReadFaceSettingsFromConfig(ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem, _
        ByVal faceName As String, ByVal defaultAs As String, ByVal defaultD As String, ByVal defaultN As String) As Variant
    ReadFaceSettingsFromConfig = Array( _
        units.InputLengthToInternal(settings.GetDouble("RoundedRectangle." & faceName & ".as", CDbl(defaultAs))), _
        OptionalDiameterFromSettings(settings, units, "RoundedRectangle." & faceName & ".d", CDbl(defaultD)), _
        settings.GetLong("RoundedRectangle." & faceName & ".n", CLng(defaultN)), _
        OptionalDiameterFromSettings(settings, units, "RoundedRectangle." & faceName & ".d_2row", 0#), _
        OptionalDiameterFromSettings(settings, units, "RoundedRectangle." & faceName & ".d_3row", 0#), _
        settings.GetString("RoundedRectangle." & faceName & ".loc_2row", "Stacked"), _
        settings.GetString("RoundedRectangle." & faceName & ".loc_3row", "Stacked"), _
        settings.GetString("RoundedRectangle." & faceName & ".bind_2row", "EachBar"), _
        settings.GetString("RoundedRectangle." & faceName & ".bind_3row", "EachBar"))
End Function

' Пустой диаметр означает отсутствие ряда, а не значение по умолчанию.
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
    Dim i As Long
    For i = 1 To barCount
        Dim distanceAlong As Double
        distanceAlong = totalLength * CDbl(i) / CDbl(barCount + 1)
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

Private Function ShouldAddAdditionalRow(ByVal baseBarOrdinal As Long, ByVal bindingMode As String) As Boolean
    If StrComp(bindingMode, "EverySecondBar", vbTextCompare) = 0 Then
        ShouldAddAdditionalRow = ((baseBarOrdinal Mod 2) = 1)
    Else
        ShouldAddAdditionalRow = True
    End If
End Function

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

Private Function IsSupportedRowLocation(ByVal value As String) As Boolean
    IsSupportedRowLocation = (StrComp(value, "SideBySide", vbTextCompare) = 0 Or StrComp(value, "Stacked", vbTextCompare) = 0)
End Function

Private Function IsSupportedRowBinding(ByVal value As String) As Boolean
    IsSupportedRowBinding = (StrComp(value, "EachBar", vbTextCompare) = 0 Or StrComp(value, "EverySecondBar", vbTextCompare) = 0)
End Function

Private Function PolylineLength(ByRef px() As Double, ByRef py() As Double) As Double
    Dim i As Long
    For i = LBound(px) To UBound(px) - 1
        PolylineLength = PolylineLength + Sqr((px(i + 1) - px(i)) ^ 2 + (py(i + 1) - py(i)) ^ 2)
    Next i
End Function

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

Private Sub NormalizeVector(ByRef xValue As Double, ByRef yValue As Double)
    Dim lenValue As Double
    lenValue = Sqr(xValue * xValue + yValue * yValue)
    If lenValue <= GEOM_TOLERANCE Then Exit Sub
    xValue = xValue / lenValue
    yValue = yValue / lenValue
End Sub

Private Function HasBarAt(ByVal layout As CRebarLayout, ByVal xCoord As Double, ByVal yCoord As Double) As Boolean
    Dim i As Long
    For i = 1 To layout.Count
        If Abs(layout.X(i) - xCoord) <= GEOM_TOLERANCE And Abs(layout.Y(i) - yCoord) <= GEOM_TOLERANCE Then
            HasBarAt = True
            Exit Function
        End If
    Next i
End Function

Private Function FaceAs(ByVal settings As Variant) As Double
    FaceAs = CDbl(settings(LBound(settings)))
End Function

Private Function FaceDiameter(ByVal settings As Variant) As Double
    FaceDiameter = CDbl(settings(LBound(settings) + 1))
End Function

Private Function FaceCount(ByVal settings As Variant) As Long
    FaceCount = CLng(settings(LBound(settings) + 2))
End Function

Private Function FaceRow2Diameter(ByVal settings As Variant) As Double
    FaceRow2Diameter = CDbl(settings(LBound(settings) + 3))
End Function

Private Function FaceRow3Diameter(ByVal settings As Variant) As Double
    FaceRow3Diameter = CDbl(settings(LBound(settings) + 4))
End Function

Private Function FaceRow2Location(ByVal settings As Variant) As String
    FaceRow2Location = CStr(settings(LBound(settings) + 5))
End Function

Private Function FaceRow3Location(ByVal settings As Variant) As String
    FaceRow3Location = CStr(settings(LBound(settings) + 6))
End Function

Private Function FaceRow2Binding(ByVal settings As Variant) As String
    FaceRow2Binding = CStr(settings(LBound(settings) + 7))
End Function

Private Function FaceRow3Binding(ByVal settings As Variant) As String
    FaceRow3Binding = CStr(settings(LBound(settings) + 8))
End Function

Private Function RoundedFaceDisplayName(ByVal faceName As String) As String
    Select Case LCase$(Trim$(faceName))
        Case "h.left": RoundedFaceDisplayName = "H - левая грань"
        Case "h.right": RoundedFaceDisplayName = "H - правая грань"
        Case "b.top": RoundedFaceDisplayName = "B - верхняя грань"
        Case "b.bottom": RoundedFaceDisplayName = "B - нижняя грань"
        Case Else: RoundedFaceDisplayName = faceName
    End Select
End Function

Public Property Get RequestedCount() As Long
    RequestedCount = mRequestedCount
End Property

Public Property Get SkippedDuplicateCount() As Long
    SkippedDuplicateCount = mSkippedDuplicateCount
End Property
