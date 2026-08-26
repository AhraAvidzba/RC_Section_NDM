VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLShapeRebarLayoutBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Автоматическая раскладка арматуры Г-образного сечения
' ==========================================================================
' Builder читает таблицу граней H1/B1/H2/B2, расставляет стержни первого ряда
' и при необходимости добавляет второй/третий ряды. Он также формирует
' semantic-группы, по которым затем строятся подписи вида 6Ø32 + 3Ø20.

Private mStepAlong As Double
Private mPerimeter As Double
Private mCentroidX As Double
Private mCentroidY As Double
Private mRequestedCount As Long
Private mSkippedDuplicateCount As Long

' Создает расчетный или интерфейсный объект из нормализованных исходных данных и локальных настроек.
Public Function Build(ByVal upperWidth As Double, ByVal upperHeight As Double, _
        ByVal lowerWidth As Double, ByVal lowerHeight As Double, _
        ByVal originX As Double, ByVal originY As Double, _
        ByVal h1Settings As Variant, ByVal h2Settings As Variant, _
        ByVal b1Settings As Variant, ByVal b2Settings As Variant, _
        ByVal steelClass As String) As CRebarLayout

    ValidateGeometry upperWidth, upperHeight, lowerWidth, lowerHeight

    Dim geom As CGeometryLShape
    Set geom = New CGeometryLShape
    geom.Initialize upperWidth, upperHeight, lowerWidth, lowerHeight, originX, originY

    Dim layout As CRebarLayout
    Set layout = New CRebarLayout

    Dim sx As Double
    Dim sy As Double
    Dim nextIndex As Long
    nextIndex = 1
    mPerimeter = 2# * (upperWidth + upperHeight) + 2# * (lowerWidth + lowerHeight)
    mStepAlong = 0#
    mRequestedCount = 0
    mSkippedDuplicateCount = 0

    AddVerticalPair layout, geom, "H1", originX, originY + lowerHeight, upperHeight, _
        upperWidth, h1Settings, b1Settings, steelClass, nextIndex, sx, sy
    AddHorizontalPair layout, geom, "B1", originX, originY + lowerHeight, upperWidth, _
        upperHeight, b1Settings, h1Settings, steelClass, nextIndex, sx, sy
    AddVerticalPair layout, geom, "H2", originX, originY, lowerHeight, _
        lowerWidth, h2Settings, b2Settings, steelClass, nextIndex, sx, sy
    AddHorizontalPair layout, geom, "B2", originX, originY, lowerWidth, _
        lowerHeight, b2Settings, h2Settings, steelClass, nextIndex, sx, sy

    If layout.Count <= 0 Then Err.Raise vbObjectError + 3179, "CLShapeRebarLayoutBuilder", "Для Г-сечения должна быть задана арматура хотя бы у одной грани."
    mCentroidX = sx / CDbl(layout.Count)
    mCentroidY = sy / CDbl(layout.Count)
    Set Build = layout
End Function

' Проверяет входные данные и прерывает выполнение понятной ошибкой, если расчетный контракт нарушен.
Private Sub ValidateGeometry(ByVal b1 As Double, ByVal h1 As Double, ByVal b2 As Double, ByVal h2 As Double)
    Dim geom As CGeometryLShape
    Dim message As String
    Set geom = New CGeometryLShape
    geom.Initialize b1, h1, b2, h2
    If Not geom.IsValid(message) Then Err.Raise vbObjectError + 3170, "CLShapeRebarLayoutBuilder", message
End Sub

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

Private Sub AddFaceBars(ByVal layout As CRebarLayout, ByVal geom As ISectionGeometry, _
        ByVal faceName As String, ByVal x1 As Double, ByVal y1 As Double, _
        ByVal x2 As Double, ByVal y2 As Double, ByVal faceLength As Double, _
        ByVal faceSettings As Variant, ByVal steelClass As String, _
        ByVal outsideNormalX As Double, ByVal outsideNormalY As Double, _
        ByVal sideX As Double, ByVal sideY As Double, ByVal stackX As Double, ByVal stackY As Double, _
        ByRef nextIndex As Long, ByRef sx As Double, ByRef sy As Double)
    Dim axisDistance As Double
    Dim barCount As Long
    Dim barDiameter As Double
    Dim startOffset As Double
    Dim endOffset As Double
    Dim row2Diameter As Double
    Dim row3Diameter As Double
    Dim row2Location As String
    Dim row3Location As String
    Dim row2Binding As String
    Dim row3Binding As String
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

    If barCount < 0 Then Err.Raise vbObjectError + 3171, "CLShapeRebarLayoutBuilder", faceName & " количество стержней должно быть нулевым или положительным."
    If barDiameter <= 0# Then Err.Raise vbObjectError + 3172, "CLShapeRebarLayoutBuilder", faceName & " диаметр стержня первого ряда должен быть положительным."
    If axisDistance <= barDiameter / 2# Then Err.Raise vbObjectError + 3173, "CLShapeRebarLayoutBuilder", faceName & " расстояние до оси стержня должно быть больше половины диаметра."
    If startOffset < 0# Or endOffset < 0# Then Err.Raise vbObjectError + 3174, "CLShapeRebarLayoutBuilder", faceName & " отступы должны быть неотрицательными."
    ValidateAdditionalRows faceName, row2Diameter, row3Diameter, row2Location, row3Location, row2Binding, row3Binding

    Dim usableLength As Double
    usableLength = faceLength - startOffset - endOffset
    If usableLength < -GEOM_TOLERANCE Then Err.Raise vbObjectError + 3175, "CLShapeRebarLayoutBuilder", faceName & " сумма отступов превышает длину грани."

    Dim stepAlong As Double
    If barCount > 1 Then
        stepAlong = usableLength / CDbl(barCount - 1)
    Else
        stepAlong = 0#
    End If
    If stepAlong > mStepAlong Then mStepAlong = stepAlong
    mRequestedCount = mRequestedCount + barCount

    Dim distances() As Double
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
        Dim chosen As Long
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

Private Function ShouldAddAdditionalRow(ByVal baseBarOrdinal As Long, ByVal bindingMode As String) As Boolean
    If StrComp(bindingMode, "EverySecondBar", vbTextCompare) = 0 Then
        ShouldAddAdditionalRow = ((baseBarOrdinal Mod 2) = 1)
    Else
        ShouldAddAdditionalRow = True
    End If
End Function

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

Private Sub AddBarIfAvailable(ByVal layout As CRebarLayout, ByVal geom As ISectionGeometry, _
        ByVal faceName As String, ByVal rowNumber As Long, ByVal xCoord As Double, ByVal yCoord As Double, _
        ByVal diameter As Double, ByVal steelClass As String, ByRef nextIndex As Long, ByRef sx As Double, ByRef sy As Double)
    If HasBarAt(layout, xCoord, yCoord) Then
        mSkippedDuplicateCount = mSkippedDuplicateCount + 1
    Else
        layout.AddBar faceName & ".row_" & CStr(rowNumber) & "-" & CStr(nextIndex), _
            xCoord, yCoord, diameter, 0#, steelClass, "auto lshape " & faceName & " row " & CStr(rowNumber), geom, faceName
        nextIndex = nextIndex + 1
        sx = sx + xCoord
        sy = sy + yCoord
    End If
End Sub

Private Function LineAxisDistance(ByVal lineSettings As Variant) As Double
    LineAxisDistance = CDbl(lineSettings(LBound(lineSettings)))
End Function

Private Function LineBarCount(ByVal lineSettings As Variant) As Long
    LineBarCount = CLng(lineSettings(LBound(lineSettings) + 1))
End Function

Private Function LineBarDiameter(ByVal lineSettings As Variant) As Double
    LineBarDiameter = CDbl(lineSettings(LBound(lineSettings) + 2))
End Function

Private Function LineStartOffset(ByVal lineSettings As Variant) As Double
    LineStartOffset = CDbl(lineSettings(LBound(lineSettings) + 3))
End Function

Private Function LineEndOffset(ByVal lineSettings As Variant) As Double
    LineEndOffset = CDbl(lineSettings(LBound(lineSettings) + 4))
End Function

Private Function LineRow2Diameter(ByVal lineSettings As Variant) As Double
    If UBound(lineSettings) - LBound(lineSettings) < 5 Then Exit Function
    LineRow2Diameter = CDbl(lineSettings(LBound(lineSettings) + 5))
End Function

Private Function LineRow3Diameter(ByVal lineSettings As Variant) As Double
    If UBound(lineSettings) - LBound(lineSettings) < 6 Then Exit Function
    LineRow3Diameter = CDbl(lineSettings(LBound(lineSettings) + 6))
End Function

Private Function LineRow2Location(ByVal lineSettings As Variant) As String
    LineRow2Location = "Stacked"
    If UBound(lineSettings) - LBound(lineSettings) < 7 Then Exit Function
    If Len(Trim$(CStr(lineSettings(LBound(lineSettings) + 7)))) > 0 Then LineRow2Location = Trim$(CStr(lineSettings(LBound(lineSettings) + 7)))
End Function

Private Function LineRow3Location(ByVal lineSettings As Variant) As String
    LineRow3Location = "Stacked"
    If UBound(lineSettings) - LBound(lineSettings) < 8 Then Exit Function
    If Len(Trim$(CStr(lineSettings(LBound(lineSettings) + 8)))) > 0 Then LineRow3Location = Trim$(CStr(lineSettings(LBound(lineSettings) + 8)))
End Function

Private Function LineRow2Binding(ByVal lineSettings As Variant) As String
    LineRow2Binding = "EachBar"
    If UBound(lineSettings) - LBound(lineSettings) < 9 Then Exit Function
    If Len(Trim$(CStr(lineSettings(LBound(lineSettings) + 9)))) > 0 Then LineRow2Binding = Trim$(CStr(lineSettings(LBound(lineSettings) + 9)))
End Function

Private Function LineRow3Binding(ByVal lineSettings As Variant) As String
    LineRow3Binding = "EachBar"
    If UBound(lineSettings) - LBound(lineSettings) < 10 Then Exit Function
    If Len(Trim$(CStr(lineSettings(LBound(lineSettings) + 10)))) > 0 Then LineRow3Binding = Trim$(CStr(lineSettings(LBound(lineSettings) + 10)))
End Function

' Проверяет входные данные и прерывает выполнение понятной ошибкой, если расчетный контракт нарушен.
Private Sub ValidateAdditionalRows(ByVal faceName As String, ByVal row2Diameter As Double, ByVal row3Diameter As Double, _
        ByVal row2Location As String, ByVal row3Location As String, ByVal row2Binding As String, ByVal row3Binding As String)
    If row2Diameter < 0# Or row3Diameter < 0# Then Err.Raise vbObjectError + 3180, "CLShapeRebarLayoutBuilder", faceName & " Диаметр дополнительного ряда должен быть неотрицательным."
    If row2Diameter > 0# And Not IsSupportedRowLocation(row2Location) Then Err.Raise vbObjectError + 3182, "CLShapeRebarLayoutBuilder", faceName & " loc_2row должен быть SideBySide или Stacked."
    If row3Diameter > 0# And Not IsSupportedRowLocation(row3Location) Then Err.Raise vbObjectError + 3183, "CLShapeRebarLayoutBuilder", faceName & " loc_3row должен быть SideBySide или Stacked."
    If row2Diameter > 0# And Not IsSupportedRowBinding(row2Binding) Then Err.Raise vbObjectError + 3184, "CLShapeRebarLayoutBuilder", faceName & " bind_2row должен быть EachBar или EverySecondBar."
    If row3Diameter > 0# And Not IsSupportedRowBinding(row3Binding) Then Err.Raise vbObjectError + 3185, "CLShapeRebarLayoutBuilder", faceName & " bind_3row должен быть EachBar или EverySecondBar."
End Sub

Private Function IsSupportedRowLocation(ByVal value As String) As Boolean
    IsSupportedRowLocation = (StrComp(value, "SideBySide", vbTextCompare) = 0 Or StrComp(value, "Stacked", vbTextCompare) = 0)
End Function

Private Function IsSupportedRowBinding(ByVal value As String) As Boolean
    IsSupportedRowBinding = (StrComp(value, "EachBar", vbTextCompare) = 0 Or StrComp(value, "EverySecondBar", vbTextCompare) = 0)
End Function

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

Private Function FaceAs1(ByVal faceSettings As Variant) As Double
    FaceAs1 = CDbl(faceSettings(LBound(faceSettings)))
End Function

Private Function FaceAs2(ByVal faceSettings As Variant) As Double
    FaceAs2 = CDbl(faceSettings(LBound(faceSettings) + 1))
End Function

Private Function FaceDiameter1(ByVal faceSettings As Variant) As Double
    FaceDiameter1 = CDbl(faceSettings(LBound(faceSettings) + 2))
End Function

Private Function FaceDiameter2(ByVal faceSettings As Variant) As Double
    FaceDiameter2 = CDbl(faceSettings(LBound(faceSettings) + 3))
End Function

Private Function FaceCount1(ByVal faceSettings As Variant) As Long
    FaceCount1 = CLng(faceSettings(LBound(faceSettings) + 4))
End Function

Private Function FaceCount2(ByVal faceSettings As Variant) As Long
    FaceCount2 = CLng(faceSettings(LBound(faceSettings) + 5))
End Function

Private Function FaceStartOffset1(ByVal faceSettings As Variant) As Double
    FaceStartOffset1 = CDbl(faceSettings(LBound(faceSettings) + 6))
End Function

Private Function FaceEndOffset1(ByVal faceSettings As Variant) As Double
    FaceEndOffset1 = CDbl(faceSettings(LBound(faceSettings) + 7))
End Function

Private Function FaceStartOffset2(ByVal faceSettings As Variant) As Double
    FaceStartOffset2 = CDbl(faceSettings(LBound(faceSettings) + 8))
End Function

Private Function FaceEndOffset2(ByVal faceSettings As Variant) As Double
    FaceEndOffset2 = CDbl(faceSettings(LBound(faceSettings) + 9))
End Function

Private Function FaceRow2Diameter1(ByVal faceSettings As Variant) As Double
    If UBound(faceSettings) - LBound(faceSettings) < 10 Then Exit Function
    FaceRow2Diameter1 = CDbl(faceSettings(LBound(faceSettings) + 10))
End Function

Private Function FaceRow2Diameter2(ByVal faceSettings As Variant) As Double
    If UBound(faceSettings) - LBound(faceSettings) < 11 Then Exit Function
    FaceRow2Diameter2 = CDbl(faceSettings(LBound(faceSettings) + 11))
End Function

Private Function FaceRow3Diameter1(ByVal faceSettings As Variant) As Double
    If UBound(faceSettings) - LBound(faceSettings) < 12 Then Exit Function
    FaceRow3Diameter1 = CDbl(faceSettings(LBound(faceSettings) + 12))
End Function

Private Function FaceRow3Diameter2(ByVal faceSettings As Variant) As Double
    If UBound(faceSettings) - LBound(faceSettings) < 13 Then Exit Function
    FaceRow3Diameter2 = CDbl(faceSettings(LBound(faceSettings) + 13))
End Function

Private Function FaceRow2Location(ByVal faceSettings As Variant) As String
    FaceRow2Location = "Stacked"
    If UBound(faceSettings) - LBound(faceSettings) < 14 Then Exit Function
    If Len(Trim$(CStr(faceSettings(LBound(faceSettings) + 14)))) > 0 Then FaceRow2Location = Trim$(CStr(faceSettings(LBound(faceSettings) + 14)))
End Function

Private Function FaceRow3Location(ByVal faceSettings As Variant) As String
    FaceRow3Location = "Stacked"
    If UBound(faceSettings) - LBound(faceSettings) < 15 Then Exit Function
    If Len(Trim$(CStr(faceSettings(LBound(faceSettings) + 15)))) > 0 Then FaceRow3Location = Trim$(CStr(faceSettings(LBound(faceSettings) + 15)))
End Function

Private Function FaceRow2Binding(ByVal faceSettings As Variant) As String
    FaceRow2Binding = "EachBar"
    If UBound(faceSettings) - LBound(faceSettings) < 16 Then Exit Function
    If Len(Trim$(CStr(faceSettings(LBound(faceSettings) + 16)))) > 0 Then FaceRow2Binding = Trim$(CStr(faceSettings(LBound(faceSettings) + 16)))
End Function

Private Function FaceRow3Binding(ByVal faceSettings As Variant) As String
    FaceRow3Binding = "EachBar"
    If UBound(faceSettings) - LBound(faceSettings) < 17 Then Exit Function
    If Len(Trim$(CStr(faceSettings(LBound(faceSettings) + 17)))) > 0 Then FaceRow3Binding = Trim$(CStr(faceSettings(LBound(faceSettings) + 17)))
End Function

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

Private Function HasBarAt(ByVal layout As CRebarLayout, ByVal xCoord As Double, ByVal yCoord As Double) As Boolean
    Dim i As Long
    For i = 1 To layout.Count
        If Abs(layout.X(i) - xCoord) <= GEOM_TOLERANCE And Abs(layout.Y(i) - yCoord) <= GEOM_TOLERANCE Then
            HasBarAt = True
            Exit Function
        End If
    Next i
End Function

Public Property Get CentroidX() As Double
    CentroidX = mCentroidX
End Property

Public Property Get CentroidY() As Double
    CentroidY = mCentroidY
End Property

Public Property Get SkippedDuplicateCount() As Long
    SkippedDuplicateCount = mSkippedDuplicateCount
End Property

' ==========================================================================
' ДЛЯ ТЕСТОВ
' ==========================================================================
' Эти свойства раскрывают внутреннюю геометрию раскладки Г-образного сечения.
' Тесты используют их, чтобы проверить длину обхода и шаг расстановки без
' анализа уже созданных координат каждого стержня.

Public Property Get StepAlong() As Double
    StepAlong = mStepAlong
End Property

Public Property Get Perimeter() As Double
    Perimeter = mPerimeter
End Property

Public Property Get RequestedCount() As Long
    RequestedCount = mRequestedCount
End Property

