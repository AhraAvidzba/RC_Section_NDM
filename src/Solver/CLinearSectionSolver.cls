VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLinearSectionSolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mConverged As Boolean
Private mStopReason As String
Private mEpsilon0 As Double
Private mKappaX As Double
Private mKappaY As Double
Private mNint As Double
Private mMxint As Double
Private mMyint As Double
Private mResidualN As Double
Private mResidualMx As Double
Private mResidualMy As Double
Private mRelativeResidualN As Double
Private mRelativeResidualMx As Double
Private mRelativeResidualMy As Double
Private mMinConcreteStrain As Double
Private mMaxConcreteStrain As Double
Private mRebarStrain() As Double
Private mRebarStress() As Double
Private mRebarCount As Long
Private mK(1 To 3, 1 To 3) As Double
Private mMatrixBuildSeconds As Double
Private mSolveSeconds As Double
Private mBackCalcSeconds As Double
Private mDeterminant As Double
Private mSystemResidual As Double

Public Sub Solve(ByVal mesh As CFiberMeshBuilder, ByVal rebars As CRebarLayout, _
        ByVal concreteModulus As Double, ByVal steelModulus As Double, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double)

    ClearResult
    If mesh.FiberCount <= 0 Then Err.Raise vbObjectError + 3300, "CLinearSectionSolver", "Concrete fiber mesh is empty."
    If concreteModulus <= 0# Then Err.Raise vbObjectError + 3301, "CLinearSectionSolver", "Concrete modulus must be positive."
    If steelModulus <= 0# Then Err.Raise vbObjectError + 3302, "CLinearSectionSolver", "Steel modulus must be positive."

    Dim t0 As Double
    t0 = Timer
    BuildStiffness mesh, rebars, concreteModulus, steelModulus
    mMatrixBuildSeconds = Timer - t0

    t0 = Timer
    Dim system As CLinearSystem3x3
    Set system = New CLinearSystem3x3
    If Not system.Solve(mK(1, 1), mK(1, 2), mK(1, 3), _
            mK(2, 1), mK(2, 2), mK(2, 3), _
            mK(3, 1), mK(3, 2), mK(3, 3), _
            targetN, targetMx, targetMy) Then
        mConverged = False
        mStopReason = system.StopReason
        mDeterminant = system.Determinant
        mSystemResidual = system.AbsoluteResidual
        Exit Sub
    End If

    mEpsilon0 = system.X1
    mKappaX = system.X2
    mKappaY = system.X3
    mDeterminant = system.Determinant
    mSystemResidual = system.AbsoluteResidual
    mSolveSeconds = Timer - t0

    t0 = Timer
    BackCalculate mesh, rebars, concreteModulus, steelModulus, targetN, targetMx, targetMy
    mBackCalcSeconds = Timer - t0

    mConverged = True
    mStopReason = "Solved linear elastic system"
End Sub

Private Sub BuildStiffness(ByVal mesh As CFiberMeshBuilder, ByVal rebars As CRebarLayout, _
        ByVal concreteModulus As Double, ByVal steelModulus As Double)

    Dim i As Long
    For i = 1 To 3
        mK(i, 1) = 0#
        mK(i, 2) = 0#
        mK(i, 3) = 0#
    Next i

    For i = 1 To mesh.FiberCount
        AddStiffnessContribution concreteModulus * mesh.FiberArea(i) * mesh.FiberFillFactor(i), mesh.FiberX(i), mesh.FiberY(i)
    Next i

    Dim effectiveSteelModulus As Double
    effectiveSteelModulus = steelModulus - concreteModulus
    If Not rebars Is Nothing Then
        For i = 1 To rebars.Count
            AddStiffnessContribution effectiveSteelModulus * rebars.Area(i), rebars.X(i), rebars.Y(i)
        Next i
    End If
End Sub

Private Sub AddStiffnessContribution(ByVal axialStiffness As Double, ByVal xCoord As Double, ByVal yCoord As Double)
    mK(1, 1) = mK(1, 1) + axialStiffness
    mK(1, 2) = mK(1, 2) + axialStiffness * yCoord
    mK(1, 3) = mK(1, 3) + axialStiffness * xCoord
    mK(2, 1) = mK(2, 1) + axialStiffness * yCoord
    mK(2, 2) = mK(2, 2) + axialStiffness * yCoord * yCoord
    mK(2, 3) = mK(2, 3) + axialStiffness * yCoord * xCoord
    mK(3, 1) = mK(3, 1) + axialStiffness * xCoord
    mK(3, 2) = mK(3, 2) + axialStiffness * xCoord * yCoord
    mK(3, 3) = mK(3, 3) + axialStiffness * xCoord * xCoord
End Sub

Private Sub BackCalculate(ByVal mesh As CFiberMeshBuilder, ByVal rebars As CRebarLayout, _
        ByVal concreteModulus As Double, ByVal steelModulus As Double, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double)

    Dim i As Long
    Dim strain As Double
    Dim force As Double
    mNint = 0#: mMxint = 0#: mMyint = 0#
    mMinConcreteStrain = 1E+100
    mMaxConcreteStrain = -1E+100

    For i = 1 To mesh.FiberCount
        strain = StrainAt(mesh.FiberX(i), mesh.FiberY(i))
        If strain < mMinConcreteStrain Then mMinConcreteStrain = strain
        If strain > mMaxConcreteStrain Then mMaxConcreteStrain = strain
        force = concreteModulus * strain * mesh.FiberArea(i) * mesh.FiberFillFactor(i)
        mNint = mNint + force
        mMxint = mMxint + force * mesh.FiberY(i)
        mMyint = mMyint + force * mesh.FiberX(i)
    Next i

    If Not rebars Is Nothing Then
        mRebarCount = rebars.Count
        If mRebarCount > 0 Then
            ReDim mRebarStrain(1 To mRebarCount)
            ReDim mRebarStress(1 To mRebarCount)
        End If
        Dim effectiveSteelModulus As Double
        effectiveSteelModulus = steelModulus - concreteModulus
        For i = 1 To mRebarCount
            strain = StrainAt(rebars.X(i), rebars.Y(i))
            mRebarStrain(i) = strain
            mRebarStress(i) = steelModulus * strain
            force = effectiveSteelModulus * strain * rebars.Area(i)
            mNint = mNint + force
            mMxint = mMxint + force * rebars.Y(i)
            mMyint = mMyint + force * rebars.X(i)
        Next i
    End If

    mResidualN = mNint - targetN
    mResidualMx = mMxint - targetMx
    mResidualMy = mMyint - targetMy
    mRelativeResidualN = RelativeResidual(mResidualN, targetN)
    mRelativeResidualMx = RelativeResidual(mResidualMx, targetMx)
    mRelativeResidualMy = RelativeResidual(mResidualMy, targetMy)
End Sub

Private Function StrainAt(ByVal xCoord As Double, ByVal yCoord As Double) As Double
    StrainAt = mEpsilon0 + mKappaX * yCoord + mKappaY * xCoord
End Function

Private Function RelativeResidual(ByVal residual As Double, ByVal target As Double) As Double
    If Abs(target) > 1E-12 Then
        RelativeResidual = Abs(residual / target)
    Else
        RelativeResidual = Abs(residual)
    End If
End Function

Public Sub ClearResult()
    Dim i As Long
    mConverged = False
    mStopReason = vbNullString
    mEpsilon0 = 0#: mKappaX = 0#: mKappaY = 0#
    mNint = 0#: mMxint = 0#: mMyint = 0#
    mResidualN = 0#: mResidualMx = 0#: mResidualMy = 0#
    mRelativeResidualN = 0#: mRelativeResidualMx = 0#: mRelativeResidualMy = 0#
    mMinConcreteStrain = 0#: mMaxConcreteStrain = 0#
    Erase mRebarStrain
    Erase mRebarStress
    mRebarCount = 0
    For i = 1 To 3
        mK(i, 1) = 0#: mK(i, 2) = 0#: mK(i, 3) = 0#
    Next i
    mMatrixBuildSeconds = 0#: mSolveSeconds = 0#: mBackCalcSeconds = 0#
    mDeterminant = 0#: mSystemResidual = 0#
End Sub

Public Property Get Converged() As Boolean
    Converged = mConverged
End Property

Public Property Get StopReason() As String
    StopReason = mStopReason
End Property

Public Property Get Epsilon0() As Double
    Epsilon0 = mEpsilon0
End Property

Public Property Get KappaX() As Double
    KappaX = mKappaX
End Property

Public Property Get KappaY() As Double
    KappaY = mKappaY
End Property

Public Property Get Nint() As Double
    Nint = mNint
End Property

Public Property Get Mxint() As Double
    Mxint = mMxint
End Property

Public Property Get Myint() As Double
    Myint = mMyint
End Property

Public Property Get ResidualN() As Double
    ResidualN = mResidualN
End Property

Public Property Get ResidualMx() As Double
    ResidualMx = mResidualMx
End Property

Public Property Get ResidualMy() As Double
    ResidualMy = mResidualMy
End Property

Public Property Get RelativeResidualN() As Double
    RelativeResidualN = mRelativeResidualN
End Property

Public Property Get RelativeResidualMx() As Double
    RelativeResidualMx = mRelativeResidualMx
End Property

Public Property Get RelativeResidualMy() As Double
    RelativeResidualMy = mRelativeResidualMy
End Property

Public Property Get MinConcreteStrain() As Double
    MinConcreteStrain = mMinConcreteStrain
End Property

Public Property Get MaxConcreteStrain() As Double
    MaxConcreteStrain = mMaxConcreteStrain
End Property

Public Property Get RebarCount() As Long
    RebarCount = mRebarCount
End Property

Public Property Get RebarStrain(ByVal index As Long) As Double
    If index < 1 Or index > mRebarCount Then Err.Raise vbObjectError + 3303, "CLinearSectionSolver", "Rebar result index out of range."
    RebarStrain = mRebarStrain(index)
End Property

Public Property Get RebarStress(ByVal index As Long) As Double
    If index < 1 Or index > mRebarCount Then Err.Raise vbObjectError + 3304, "CLinearSectionSolver", "Rebar result index out of range."
    RebarStress = mRebarStress(index)
End Property

Public Property Get Stiffness(ByVal rowIndex As Long, ByVal columnIndex As Long) As Double
    If rowIndex < 1 Or rowIndex > 3 Or columnIndex < 1 Or columnIndex > 3 Then Err.Raise vbObjectError + 3305, "CLinearSectionSolver", "Stiffness index out of range."
    Stiffness = mK(rowIndex, columnIndex)
End Property

Public Property Get MatrixBuildSeconds() As Double
    MatrixBuildSeconds = mMatrixBuildSeconds
End Property

Public Property Get SolveSeconds() As Double
    SolveSeconds = mSolveSeconds
End Property

Public Property Get BackCalcSeconds() As Double
    BackCalcSeconds = mBackCalcSeconds
End Property

Public Property Get Determinant() As Double
    Determinant = mDeterminant
End Property

Public Property Get SystemResidual() As Double
    SystemResidual = mSystemResidual
End Property






