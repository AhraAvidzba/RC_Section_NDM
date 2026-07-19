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

Private mAngleStep As Double
Private mCentroidX As Double
Private mCentroidY As Double
Private mAxisRadius As Double

Public Function Build(ByVal diameter As Double, ByVal centerX As Double, ByVal centerY As Double, _
        ByVal axisDistance As Double, ByVal barCount As Long, ByVal barDiameter As Double, _
        ByVal steelClass As String) As CRebarLayout

    ValidateInputs diameter, axisDistance, barCount, barDiameter

    Dim geom As CGeometryCircle
    Set geom = New CGeometryCircle
    geom.InitializeByDiameter diameter, centerX, centerY

    mAxisRadius = diameter / 2# - axisDistance
    mAngleStep = 2# * GEOM_PI / CDbl(barCount)

    Dim layout As CRebarLayout
    Set layout = New CRebarLayout

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
        layout.AddBar "B" & CStr(i + 1), xCoord, yCoord, barDiameter, 0#, steelClass, "auto circle", geom
        sx = sx + xCoord
        sy = sy + yCoord
    Next i

    If layout.Count <> barCount Then Err.Raise vbObjectError + 3160, "CCircleRebarLayoutBuilder", "Generated rebar count mismatch."

    mCentroidX = sx / CDbl(barCount)
    mCentroidY = sy / CDbl(barCount)
    If Abs(mCentroidX - centerX) > 0.000001 Or Abs(mCentroidY - centerY) > 0.000001 Then
        Err.Raise vbObjectError + 3161, "CCircleRebarLayoutBuilder", "Generated rebar centroid does not match circle center."
    End If

    Set Build = layout
End Function

Private Sub ValidateInputs(ByVal diameter As Double, ByVal axisDistance As Double, _
        ByVal barCount As Long, ByVal barDiameter As Double)
    If diameter <= 0# Then Err.Raise vbObjectError + 3150, "CCircleRebarLayoutBuilder", "Circle diameter must be positive."
    If barCount < 2 Then Err.Raise vbObjectError + 3151, "CCircleRebarLayoutBuilder", "Rebar count must be at least 2."
    If barDiameter <= 0# Then Err.Raise vbObjectError + 3152, "CCircleRebarLayoutBuilder", "Rebar diameter must be positive."
    If axisDistance <= barDiameter / 2# Then Err.Raise vbObjectError + 3153, "CCircleRebarLayoutBuilder", "Axis distance must be greater than half rebar diameter."
    If axisDistance >= diameter / 2# Then Err.Raise vbObjectError + 3154, "CCircleRebarLayoutBuilder", "Axis distance must be less than half circle diameter."
End Sub

Public Property Get AngleStep() As Double
    AngleStep = mAngleStep
End Property

Public Property Get CentroidX() As Double
    CentroidX = mCentroidX
End Property

Public Property Get CentroidY() As Double
    CentroidY = mCentroidY
End Property

Public Property Get AxisRadius() As Double
    AxisRadius = mAxisRadius
End Property
