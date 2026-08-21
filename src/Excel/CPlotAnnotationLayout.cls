VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CPlotAnnotationLayout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const ITEM_LINE As String = "LINE"
Private Const ITEM_TEXT As String = "TEXT"

Private mMinX As Double
Private mMaxX As Double
Private mMinY As Double
Private mMaxY As Double
Private mPlotLeft As Double
Private mPlotTop As Double
Private mPlotWidth As Double
Private mPlotHeight As Double
Private mScaleX As Double
Private mScaleY As Double
Private mLengthUnitFactorToMm As Double
Private mVisualScale As Double
Private mDimensionFontSize As Double
Private mRebarFontSize As Double
Private mDimensionBaseGap As Double
Private mRebarBaseGap As Double
Private mDimensionLaneGap As Double
Private mRebarLaneGap As Double
Private mBoxPadding As Double
Private mDimensionLineWeight As Double
Private mExtensionLineWeight As Double
Private mRebarLineWeight As Double
Private mDimensionTextHeightBase As Double
Private mRebarTextHeightBase As Double
Private mDimensionBaseGapBase As Double
Private mRebarBaseGapBase As Double
Private mDimensionLineWeightBase As Double
Private mExtensionLineWeightBase As Double
Private mRebarLineWeightBase As Double
Private mDimensionTextGapBase As Double
Private mRebarTextGapBase As Double
Private mDimensionPlacement As String
Private mRebarPlacement As String
Private mRebarLineEnabled As Boolean
Private mDimensionColor As Long
Private mDimensionExtensionColor As Long
Private mRebarColor As Long

Private mCount As Long
Private mKind() As String
Private mX1() As Double
Private mY1() As Double
Private mX2() As Double
Private mY2() As Double
Private mText() As String
Private mRotation() As Double
Private mColor() As Long
Private mWeight() As Double
Private mBold() As Boolean
Private mArrowheads() As Boolean
Private mFontSize() As Double

Private mOccupiedCount As Long
Private mOccLeft() As Double
Private mOccTop() As Double
Private mOccRight() As Double
Private mOccBottom() As Double

Public Sub ConfigureFromSettings(ByVal settings As CSystemSettingsReader)
    mDimensionTextHeightBase = settings.GetDouble("Plot.Dimensions.TextHeight", 8.5)
    mRebarTextHeightBase = settings.GetDouble("Plot.RebarLabels.TextHeight", 9#)
    mDimensionBaseGapBase = settings.GetDouble("Plot.Dimensions.Offset", 42#)
    mRebarBaseGapBase = settings.GetDouble("Plot.RebarLabels.Offset", 10#)
    mDimensionLineWeightBase = settings.GetDouble("Plot.Dimensions.LineWeight", 1.75)
    mExtensionLineWeightBase = settings.GetDouble("Plot.Dimensions.ExtensionLineWeight", 0.85)
    mRebarLineWeightBase = settings.GetDouble("Plot.RebarLabels.LineWeight", 1.35)
    mDimensionTextGapBase = settings.GetDouble("Plot.Dimensions.TextGap", 7#)
    mRebarTextGapBase = settings.GetDouble("Plot.RebarLabels.TextGap", 7#)
    mDimensionPlacement = NormalizePlacement(settings.GetString("Plot.Dimensions.Placement", "Outside"))
    mRebarPlacement = NormalizePlacement(settings.GetString("Plot.RebarLabels.Placement", "Outside"))
    mRebarLineEnabled = settings.GetBoolean("Plot.RebarLabels.LineEnabled", True)
    mDimensionColor = ParseRgb(settings.GetString("Plot.Dimensions.Color", "20,30,90"), RGB(20, 30, 90))
    mDimensionExtensionColor = ParseRgb(settings.GetString("Plot.Dimensions.ExtensionLineColor", "140,140,140"), RGB(140, 140, 140))
    mRebarColor = ParseRgb(settings.GetString("Plot.RebarLabels.Color", "20,30,90"), RGB(20, 30, 90))
End Sub

Public Sub Initialize(ByVal minX As Double, ByVal maxX As Double, ByVal minY As Double, ByVal maxY As Double, _
        ByVal plotLeft As Double, ByVal plotTop As Double, ByVal plotWidth As Double, ByVal plotHeight As Double, _
        Optional ByVal lengthUnit As String = "mm")
    mMinX = minX
    mMaxX = maxX
    mMinY = minY
    mMaxY = maxY
    mPlotLeft = plotLeft
    mPlotTop = plotTop
    mPlotWidth = plotWidth
    mPlotHeight = plotHeight
    If mMaxX = mMinX Then mMaxX = mMinX + 1#
    If mMaxY = mMinY Then mMaxY = mMinY + 1#
    mScaleX = mPlotWidth / (mMaxX - mMinX)
    mScaleY = mPlotHeight / (mMaxY - mMinY)
    mLengthUnitFactorToMm = LengthFactorToMm(lengthUnit)
    InitializeVisualMetrics plotWidth, plotHeight
End Sub

Private Sub InitializeVisualMetrics(ByVal plotWidth As Double, ByVal plotHeight As Double)
    Dim sx As Double, sy As Double
    sx = plotWidth / 640#
    sy = plotHeight / 430#
    mVisualScale = ClampDouble(MinDouble(sx, sy), 0.65, 1.8)

    If mDimensionTextHeightBase <= 0# Then mDimensionTextHeightBase = 8.5
    If mRebarTextHeightBase <= 0# Then mRebarTextHeightBase = 9#
    If mDimensionBaseGapBase <= 0# Then mDimensionBaseGapBase = 42#
    If mRebarBaseGapBase <= 0# Then mRebarBaseGapBase = 10#
    If mDimensionLineWeightBase <= 0# Then mDimensionLineWeightBase = 1.75
    If mExtensionLineWeightBase <= 0# Then mExtensionLineWeightBase = 0.85
    If mRebarLineWeightBase <= 0# Then mRebarLineWeightBase = 1.35
    If mDimensionTextGapBase <= 0# Then mDimensionTextGapBase = 7#
    If mRebarTextGapBase <= 0# Then mRebarTextGapBase = 7#
    If Len(mDimensionPlacement) = 0 Then mDimensionPlacement = "Outside"
    If Len(mRebarPlacement) = 0 Then mRebarPlacement = "Outside"
    If mDimensionColor = 0 Then mDimensionColor = RGB(20, 30, 90)
    If mDimensionExtensionColor = 0 Then mDimensionExtensionColor = RGB(140, 140, 140)
    If mRebarColor = 0 Then mRebarColor = RGB(20, 30, 90)

    mDimensionFontSize = ClampDouble(ModelMmToChartPoints(mDimensionTextHeightBase), 5#, 28#)
    mRebarFontSize = ClampDouble(ModelMmToChartPoints(mRebarTextHeightBase), 5#, 28#)
    mDimensionBaseGap = ClampDouble(ModelMmToChartPoints(mDimensionBaseGapBase), 4#, 180#)
    mRebarBaseGap = ClampDouble(ModelMmToChartPoints(mRebarBaseGapBase), 2#, 120#)
    mDimensionLaneGap = 24# * mVisualScale
    mRebarLaneGap = 18# * mVisualScale
    mBoxPadding = 5# * mVisualScale
    mDimensionLineWeight = mDimensionLineWeightBase * mVisualScale
    mExtensionLineWeight = mExtensionLineWeightBase * mVisualScale
    mRebarLineWeight = mRebarLineWeightBase * mVisualScale
End Sub

Public Sub AddDimension(ByVal startX As Double, ByVal startY As Double, ByVal endX As Double, ByVal endY As Double, _
        ByVal normalX As Double, ByVal normalY As Double, ByVal offset As Double, ByVal textValue As String)
    Dim dirX As Double, dirY As Double
    dirX = endX - startX
    dirY = endY - startY
    NormalizeModelVector dirX, dirY, 1#, 0#
    NormalizeModelVector normalX, normalY, -dirY, dirX

    Dim normalPxX As Double, normalPxY As Double
    PixelUnitVector normalX, normalY, normalPxX, normalPxY

    Dim baseOffset As Double
    baseOffset = mDimensionBaseGap

    Dim sx As Double, sy As Double, ex As Double, ey As Double
    sx = ModelToChartX(startX) + normalPxX * baseOffset
    sy = ModelToChartY(startY) + normalPxY * baseOffset
    ex = ModelToChartX(endX) + normalPxX * baseOffset
    ey = ModelToChartY(endY) + normalPxY * baseOffset

    Dim rotation As Double
    rotation = AnnotationRotation(dirX, dirY)

    Dim textNormalPxX As Double, textNormalPxY As Double
    textNormalPxX = normalPxX
    textNormalPxY = normalPxY
    ApplyPlacement textNormalPxX, textNormalPxY, mDimensionPlacement

    Dim textLeft As Double, textTop As Double, textWidth As Double, textHeight As Double
    DimensionTextBox sx, sy, ex, ey, textNormalPxX, textNormalPxY, rotation, textLeft, textTop, textWidth, textHeight

    AddLine ModelToChartX(startX), ModelToChartY(startY), sx, sy, mDimensionExtensionColor, mExtensionLineWeight, False
    AddLine ModelToChartX(endX), ModelToChartY(endY), ex, ey, mDimensionExtensionColor, mExtensionLineWeight, False
    AddLine sx, sy, ex, ey, mDimensionColor, mDimensionLineWeight, True
    AddText textLeft, textTop, textWidth, textHeight, textValue, rotation, mDimensionColor, True, mDimensionFontSize
    AddOccupied textLeft, textTop, textWidth, textHeight
End Sub

Public Sub AddRebarLabel(ByVal startX As Double, ByVal startY As Double, ByVal endX As Double, ByVal endY As Double, _
        ByVal normalX As Double, ByVal normalY As Double, ByVal offset As Double, ByVal textValue As String)
    Dim dirX As Double, dirY As Double
    dirX = endX - startX
    dirY = endY - startY
    NormalizeModelVector dirX, dirY, 1#, 0#
    NormalizeModelVector normalX, normalY, -dirY, dirX

    Dim normalPxX As Double, normalPxY As Double
    PixelUnitVector normalX, normalY, normalPxX, normalPxY
    ApplyPlacement normalPxX, normalPxY, mRebarPlacement

    Dim dirPxX As Double, dirPxY As Double
    PixelUnitVector dirX, dirY, dirPxX, dirPxY

    Dim baseOffset As Double
    baseOffset = mRebarBaseGap
    If StrComp(mRebarPlacement, "Outside", vbTextCompare) = 0 And offset > 0# Then
        baseOffset = baseOffset + ModelDistanceToChartPoints(offset)
    End If

    Dim anchorX As Double, anchorY As Double
    anchorX = 0.5 * (startX + endX)
    anchorY = 0.5 * (startY + endY)

    Dim labelOffset As Double
    labelOffset = baseOffset

    Dim anchorPxX As Double, anchorPxY As Double
    anchorPxX = ModelToChartX(anchorX) + normalPxX * labelOffset
    anchorPxY = ModelToChartY(anchorY) + normalPxY * labelOffset

    Dim rotation As Double
    rotation = AnnotationRotation(dirX, dirY)

    Dim textLeft As Double, textTop As Double, textWidth As Double, textHeight As Double
    RebarTextBox anchorPxX, anchorPxY, normalPxX, normalPxY, rotation, textLeft, textTop, textWidth, textHeight

    Dim halfLength As Double
    halfLength = MaxDouble(26# * mVisualScale, 0.22 * Sqr((ModelToChartX(endX) - ModelToChartX(startX)) * _
        (ModelToChartX(endX) - ModelToChartX(startX)) + (ModelToChartY(endY) - ModelToChartY(startY)) * _
        (ModelToChartY(endY) - ModelToChartY(startY))))
    If mRebarLineEnabled Then
        AddLine anchorPxX - dirPxX * halfLength, anchorPxY - dirPxY * halfLength, _
            anchorPxX + dirPxX * halfLength, anchorPxY + dirPxY * halfLength, mRebarColor, mRebarLineWeight, False
    End If
    AddText textLeft, textTop, textWidth, textHeight, textValue, rotation, mRebarColor, True, mRebarFontSize
    AddOccupied textLeft, textTop, textWidth, textHeight
End Sub

Private Sub DimensionTextBox(ByVal sx As Double, ByVal sy As Double, ByVal ex As Double, ByVal ey As Double, _
        ByVal normalPxX As Double, ByVal normalPxY As Double, ByVal rotation As Double, _
        ByRef leftPos As Double, ByRef topPos As Double, ByRef widthValue As Double, ByRef heightValue As Double)
    If rotation = 0# Then
        widthValue = 150# * mVisualScale
        heightValue = MaxDouble(18# * mVisualScale, mDimensionFontSize * 1.65)
    Else
        widthValue = 150# * mVisualScale
        heightValue = MaxDouble(18# * mVisualScale, mDimensionFontSize * 1.65)
    End If

    Dim textGap As Double
    textGap = ClampDouble(ModelMmToChartPoints(mDimensionTextGapBase), 0#, 80#)

    Dim midX As Double, midY As Double
    midX = 0.5 * (sx + ex) + normalPxX * textGap
    midY = 0.5 * (sy + ey) + normalPxY * textGap

    leftPos = midX - widthValue / 2#
    topPos = midY - heightValue / 2#
End Sub

Private Sub RebarTextBox(ByVal anchorPxX As Double, ByVal anchorPxY As Double, ByVal normalPxX As Double, _
        ByVal normalPxY As Double, ByVal rotation As Double, ByRef leftPos As Double, ByRef topPos As Double, _
        ByRef widthValue As Double, ByRef heightValue As Double)
    If rotation = 0# Then
        widthValue = 150# * mVisualScale
        heightValue = MaxDouble(18# * mVisualScale, mRebarFontSize * 1.65)
    Else
        widthValue = 150# * mVisualScale
        heightValue = MaxDouble(18# * mVisualScale, mRebarFontSize * 1.65)
    End If

    Dim textGap As Double
    textGap = ClampDouble(ModelMmToChartPoints(mRebarTextGapBase), 0#, 80#)

    Dim centerX As Double, centerY As Double
    centerX = anchorPxX + normalPxX * textGap
    centerY = anchorPxY + normalPxY * textGap

    leftPos = centerX - widthValue / 2#
    topPos = centerY - heightValue / 2#
End Sub

Private Function AnnotationRotation(ByVal dirX As Double, ByVal dirY As Double) As Double
    If Abs(dirY) > Abs(dirX) Then AnnotationRotation = 90# Else AnnotationRotation = 0#
End Function

Private Function NormalizePlacement(ByVal value As String) As String
    If StrComp(Trim$(value), "Inside", vbTextCompare) = 0 Then
        NormalizePlacement = "Inside"
    Else
        NormalizePlacement = "Outside"
    End If
End Function

Private Sub ApplyPlacement(ByRef normalPxX As Double, ByRef normalPxY As Double, ByVal placement As String)
    If StrComp(placement, "Inside", vbTextCompare) = 0 Then
        normalPxX = -normalPxX
        normalPxY = -normalPxY
    End If
End Sub

Private Function ParseRgb(ByVal textValue As String, ByVal fallbackColor As Long) As Long
    On Error GoTo Failed
    Dim parts As Variant
    parts = Split(Trim$(textValue), ",")
    If UBound(parts) <> 2 Then GoTo Failed
    ParseRgb = RGB(CLng(Trim$(CStr(parts(0)))), CLng(Trim$(CStr(parts(1)))), CLng(Trim$(CStr(parts(2)))))
    Exit Function
Failed:
    ParseRgb = fallbackColor
End Function

Private Sub NormalizeModelVector(ByRef x As Double, ByRef y As Double, ByVal fallbackX As Double, ByVal fallbackY As Double)
    Dim lengthValue As Double
    lengthValue = Sqr(x * x + y * y)
    If lengthValue <= 0# Then
        x = fallbackX
        y = fallbackY
        lengthValue = Sqr(x * x + y * y)
    End If
    If lengthValue <= 0# Then
        x = 1#
        y = 0#
    Else
        x = x / lengthValue
        y = y / lengthValue
    End If
End Sub

Private Sub PixelUnitVector(ByVal modelX As Double, ByVal modelY As Double, ByRef pixelX As Double, ByRef pixelY As Double)
    pixelX = modelX * mScaleX
    pixelY = -modelY * mScaleY
    Dim lengthValue As Double
    lengthValue = Sqr(pixelX * pixelX + pixelY * pixelY)
    If lengthValue <= 0# Then
        pixelX = 1#
        pixelY = 0#
    Else
        pixelX = pixelX / lengthValue
        pixelY = pixelY / lengthValue
    End If
End Sub

Private Function ModelToChartX(ByVal xCoord As Double) As Double
    ModelToChartX = mPlotLeft + (xCoord - mMinX) / (mMaxX - mMinX) * mPlotWidth
End Function

Private Function ModelToChartY(ByVal yCoord As Double) As Double
    ModelToChartY = mPlotTop + mPlotHeight - (yCoord - mMinY) / (mMaxY - mMinY) * mPlotHeight
End Function

Private Function ModelScale() As Double
    ModelScale = MinDouble(mScaleX, mScaleY)
End Function

Private Function ModelMmToChartPoints(ByVal valueMm As Double) As Double
    If mLengthUnitFactorToMm <= 0# Then mLengthUnitFactorToMm = 1#
    ModelMmToChartPoints = valueMm / mLengthUnitFactorToMm * ModelScale()
End Function

Private Function ModelDistanceToChartPoints(ByVal valueInModelUnits As Double) As Double
    ModelDistanceToChartPoints = valueInModelUnits * ModelScale()
End Function

Private Function LengthFactorToMm(ByVal unitName As String) As Double
    Select Case LCase$(Trim$(unitName))
        Case "mm": LengthFactorToMm = 1#
        Case "cm": LengthFactorToMm = 10#
        Case "m": LengthFactorToMm = 1000#
        Case Else: LengthFactorToMm = 1#
    End Select
End Function

Private Sub AddLine(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double, _
        ByVal colorValue As Long, ByVal weight As Double, ByVal hasArrowheads As Boolean)
    AppendItem ITEM_LINE, x1, y1, x2, y2, vbNullString, 0#, colorValue, weight, False, hasArrowheads
End Sub

Private Sub AddText(ByVal leftPos As Double, ByVal topPos As Double, ByVal widthValue As Double, _
        ByVal heightValue As Double, ByVal textValue As String, ByVal rotation As Double, _
        ByVal colorValue As Long, ByVal isBold As Boolean, ByVal fontSize As Double)
    AppendItem ITEM_TEXT, leftPos, topPos, widthValue, heightValue, textValue, rotation, colorValue, 0#, isBold, False, fontSize
End Sub

Private Sub AppendItem(ByVal kind As String, ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double, _
        ByVal textValue As String, ByVal rotation As Double, ByVal colorValue As Long, ByVal weight As Double, _
        ByVal isBold As Boolean, ByVal hasArrowheads As Boolean, Optional ByVal fontSize As Double = 0#)
    mCount = mCount + 1
    ReDim Preserve mKind(1 To mCount)
    ReDim Preserve mX1(1 To mCount)
    ReDim Preserve mY1(1 To mCount)
    ReDim Preserve mX2(1 To mCount)
    ReDim Preserve mY2(1 To mCount)
    ReDim Preserve mText(1 To mCount)
    ReDim Preserve mRotation(1 To mCount)
    ReDim Preserve mColor(1 To mCount)
    ReDim Preserve mWeight(1 To mCount)
    ReDim Preserve mBold(1 To mCount)
    ReDim Preserve mArrowheads(1 To mCount)
    ReDim Preserve mFontSize(1 To mCount)

    mKind(mCount) = kind
    mX1(mCount) = x1
    mY1(mCount) = y1
    mX2(mCount) = x2
    mY2(mCount) = y2
    mText(mCount) = textValue
    mRotation(mCount) = rotation
    mColor(mCount) = colorValue
    mWeight(mCount) = weight
    mBold(mCount) = isBold
    mArrowheads(mCount) = hasArrowheads
    mFontSize(mCount) = fontSize
End Sub

Private Function Collides(ByVal leftPos As Double, ByVal topPos As Double, ByVal widthValue As Double, _
        ByVal heightValue As Double) As Boolean
    Dim paddedLeft As Double, paddedTop As Double, paddedRight As Double, paddedBottom As Double
    paddedLeft = leftPos - mBoxPadding
    paddedTop = topPos - mBoxPadding
    paddedRight = leftPos + widthValue + mBoxPadding
    paddedBottom = topPos + heightValue + mBoxPadding

    Dim i As Long
    For i = 1 To mOccupiedCount
        If paddedLeft < mOccRight(i) And paddedRight > mOccLeft(i) And _
                paddedTop < mOccBottom(i) And paddedBottom > mOccTop(i) Then
            Collides = True
            Exit Function
        End If
    Next i
End Function

Private Sub AddOccupied(ByVal leftPos As Double, ByVal topPos As Double, ByVal widthValue As Double, _
        ByVal heightValue As Double)
    mOccupiedCount = mOccupiedCount + 1
    ReDim Preserve mOccLeft(1 To mOccupiedCount)
    ReDim Preserve mOccTop(1 To mOccupiedCount)
    ReDim Preserve mOccRight(1 To mOccupiedCount)
    ReDim Preserve mOccBottom(1 To mOccupiedCount)
    mOccLeft(mOccupiedCount) = leftPos - mBoxPadding
    mOccTop(mOccupiedCount) = topPos - mBoxPadding
    mOccRight(mOccupiedCount) = leftPos + widthValue + mBoxPadding
    mOccBottom(mOccupiedCount) = topPos + heightValue + mBoxPadding
End Sub

Private Function ClampDouble(ByVal value As Double, ByVal minValue As Double, ByVal maxValue As Double) As Double
    If value < minValue Then
        ClampDouble = minValue
    ElseIf value > maxValue Then
        ClampDouble = maxValue
    Else
        ClampDouble = value
    End If
End Function

Private Function MaxDouble(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then MaxDouble = a Else MaxDouble = b
End Function

Private Function MinDouble(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then MinDouble = a Else MinDouble = b
End Function

Public Property Get Count() As Long: Count = mCount: End Property
Public Property Get ItemKind(ByVal index As Long) As String: ItemKind = mKind(index): End Property
Public Property Get X1(ByVal index As Long) As Double: X1 = mX1(index): End Property
Public Property Get Y1(ByVal index As Long) As Double: Y1 = mY1(index): End Property
Public Property Get X2(ByVal index As Long) As Double: X2 = mX2(index): End Property
Public Property Get Y2(ByVal index As Long) As Double: Y2 = mY2(index): End Property
Public Property Get Text(ByVal index As Long) As String: Text = mText(index): End Property
Public Property Get Rotation(ByVal index As Long) As Double: Rotation = mRotation(index): End Property
Public Property Get Color(ByVal index As Long) As Long: Color = mColor(index): End Property
Public Property Get Weight(ByVal index As Long) As Double: Weight = mWeight(index): End Property
Public Property Get Bold(ByVal index As Long) As Boolean: Bold = mBold(index): End Property
Public Property Get Arrowheads(ByVal index As Long) As Boolean: Arrowheads = mArrowheads(index): End Property
Public Property Get FontSize(ByVal index As Long) As Double: FontSize = mFontSize(index): End Property
Public Property Get VisualScale() As Double: VisualScale = mVisualScale: End Property
