VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CGeometryPropertiesCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mArea As Double
Private mStaticMomentX As Double
Private mStaticMomentY As Double
Private mCentroidX As Double
Private mCentroidY As Double
Private mIx As Double
Private mIy As Double
Private mIxy As Double
Private mIxc As Double
Private mIyc As Double
Private mIxyc As Double
Private mPrincipalI1 As Double
Private mPrincipalI2 As Double
Private mPrincipalAngleRad As Double
Private mRadiusX As Double
Private mRadiusY As Double
Private mPrincipalRadius1 As Double
Private mPrincipalRadius2 As Double
Private mLastCalculationSeconds As Double

Public Sub CalculateFromMesh(ByVal mesh As CFiberMeshBuilder)
    If mesh.FiberCount <= 0 Then Err.Raise vbObjectError + 2200, "CGeometryPropertiesCalculator", "????? ??????? ????."

    Clear

    Dim t0 As Double
    t0 = Timer

    Dim i As Long
    Dim a As Double
    Dim x As Double
    Dim y As Double

    For i = 1 To mesh.FiberCount
        a = mesh.FiberArea(i) * mesh.FiberFillFactor(i)
        x = mesh.FiberX(i)
        y = mesh.FiberY(i)

        mArea = mArea + a
        mStaticMomentX = mStaticMomentX + a * y
        mStaticMomentY = mStaticMomentY + a * x
        mIx = mIx + a * y * y
        mIy = mIy + a * x * x
        mIxy = mIxy + a * x * y
    Next i

    If mArea <= GEOM_TOLERANCE Then Err.Raise vbObjectError + 2201, "CGeometryPropertiesCalculator", "????????? ??????? ??????? ????? ????."

    mCentroidX = mStaticMomentY / mArea
    mCentroidY = mStaticMomentX / mArea

    mIxc = mIx - mArea * mCentroidY * mCentroidY
    mIyc = mIy - mArea * mCentroidX * mCentroidX
    mIxyc = mIxy - mArea * mCentroidX * mCentroidY

    Dim avgI As Double
    Dim radius As Double
    avgI = (mIxc + mIyc) / 2#
    radius = Sqr(((mIxc - mIyc) / 2#) ^ 2 + mIxyc ^ 2)
    mPrincipalI1 = avgI + radius
    mPrincipalI2 = avgI - radius

    If Abs(mIxyc) <= GEOM_TOLERANCE And Abs(mIxc - mIyc) <= GEOM_TOLERANCE Then
        mPrincipalAngleRad = 0#
    Else
        mPrincipalAngleRad = 0.5 * Atn2(-2# * mIxyc, mIyc - mIxc)
    End If

    mRadiusX = Sqr(GeomMax(mIxc, 0#) / mArea)
    mRadiusY = Sqr(GeomMax(mIyc, 0#) / mArea)
    mPrincipalRadius1 = Sqr(GeomMax(mPrincipalI1, 0#) / mArea)
    mPrincipalRadius2 = Sqr(GeomMax(mPrincipalI2, 0#) / mArea)

    mLastCalculationSeconds = Timer - t0
End Sub

Public Sub Clear()
    mArea = 0#
    mStaticMomentX = 0#
    mStaticMomentY = 0#
    mCentroidX = 0#
    mCentroidY = 0#
    mIx = 0#
    mIy = 0#
    mIxy = 0#
    mIxc = 0#
    mIyc = 0#
    mIxyc = 0#
    mPrincipalI1 = 0#
    mPrincipalI2 = 0#
    mPrincipalAngleRad = 0#
    mRadiusX = 0#
    mRadiusY = 0#
    mPrincipalRadius1 = 0#
    mPrincipalRadius2 = 0#
    mLastCalculationSeconds = 0#
End Sub

Public Property Get Area() As Double
    Area = mArea
End Property

Public Property Get StaticMomentX() As Double
    StaticMomentX = mStaticMomentX
End Property

Public Property Get StaticMomentY() As Double
    StaticMomentY = mStaticMomentY
End Property

Public Property Get CentroidX() As Double
    CentroidX = mCentroidX
End Property

Public Property Get CentroidY() As Double
    CentroidY = mCentroidY
End Property

Public Property Get Ix() As Double
    Ix = mIx
End Property

Public Property Get Iy() As Double
    Iy = mIy
End Property

Public Property Get Ixy() As Double
    Ixy = mIxy
End Property

Public Property Get Ixc() As Double
    Ixc = mIxc
End Property

Public Property Get Iyc() As Double
    Iyc = mIyc
End Property

Public Property Get Ixyc() As Double
    Ixyc = mIxyc
End Property

Public Property Get PrincipalI1() As Double
    PrincipalI1 = mPrincipalI1
End Property

Public Property Get PrincipalI2() As Double
    PrincipalI2 = mPrincipalI2
End Property

Public Property Get PrincipalAngleRad() As Double
    PrincipalAngleRad = mPrincipalAngleRad
End Property

Public Property Get RadiusX() As Double
    RadiusX = mRadiusX
End Property

Public Property Get RadiusY() As Double
    RadiusY = mRadiusY
End Property

Public Property Get PrincipalRadius1() As Double
    PrincipalRadius1 = mPrincipalRadius1
End Property

Public Property Get PrincipalRadius2() As Double
    PrincipalRadius2 = mPrincipalRadius2
End Property

Public Property Get LastCalculationSeconds() As Double
    LastCalculationSeconds = mLastCalculationSeconds
End Property

Private Function Atn2(ByVal y As Double, ByVal x As Double) As Double
    If Abs(x) <= GEOM_TOLERANCE Then
        If y > 0# Then
            Atn2 = GEOM_PI / 2#
        ElseIf y < 0# Then
            Atn2 = -GEOM_PI / 2#
        Else
            Atn2 = 0#
        End If
    ElseIf x > 0# Then
        Atn2 = Atn(y / x)
    ElseIf y >= 0# Then
        Atn2 = Atn(y / x) + GEOM_PI
    Else
        Atn2 = Atn(y / x) - GEOM_PI
    End If
End Function









