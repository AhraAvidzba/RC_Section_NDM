VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSectionSolver"
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
Private mMaxAbsSteelStrain As Double
Private mIterations As Long
Private mLoadStepsCompleted As Long
Private mK(1 To 3, 1 To 3) As Double
Private mMaxIterations As Long
Private mLoadSteps As Long
Private mToleranceN As Double
Private mToleranceMx As Double
Private mToleranceMy As Double
Private mLineSearchEnabled As Boolean
Private mDampingInitial As Double
Private mMinLineSearchAlpha As Double
Private mMaxDeltaEpsilon0 As Double
Private mMaxDeltaKappa As Double
Private mDiagnosticsEnabled As Boolean
Private mDiagnosticLog As String
Private mLastAlpha As Double
Private mLastDeltaScale As Double
Private mUseInitialState As Boolean
Private mInitialEpsilon0 As Double
Private mInitialKappaX As Double
Private mInitialKappaY As Double
Private mUseInitialLoads As Boolean
Private mInitialTargetN As Double
Private mInitialTargetMx As Double
Private mInitialTargetMy As Double

Private Sub Class_Initialize()
    mMaxIterations = 40
    mLoadSteps = 5
    mToleranceN = 1#
    mToleranceMx = 1000#
    mToleranceMy = 1000#
    mLineSearchEnabled = True
    mDampingInitial = 1#
    mMinLineSearchAlpha = 0.03125
    mMaxDeltaEpsilon0 = 0.0005
    mMaxDeltaKappa = 0.00001
    mDiagnosticsEnabled = True
End Sub

Public Sub ApplySettings(ByVal settings As CSystemSettingsReader)
    If settings Is Nothing Then Exit Sub
    mMaxIterations = settings.GetLong("Solver.MaxIterations", mMaxIterations)
    mLoadSteps = settings.GetLong("Solver.LoadSteps", mLoadSteps)
    mToleranceN = settings.GetDouble("Solver.ToleranceN", mToleranceN)
    mToleranceMx = settings.GetDouble("Solver.ToleranceMx", mToleranceMx)
    mToleranceMy = settings.GetDouble("Solver.ToleranceMy", mToleranceMy)
    mLineSearchEnabled = settings.GetBoolean("Solver.LineSearchEnabled", mLineSearchEnabled)
    mDampingInitial = settings.GetDouble("Solver.DampingInitial", mDampingInitial)
    mMinLineSearchAlpha = settings.GetDouble("Solver.MinLineSearchAlpha", mMinLineSearchAlpha)
    mMaxDeltaEpsilon0 = settings.GetDouble("Solver.MaxDeltaEpsilon0", mMaxDeltaEpsilon0)
    mMaxDeltaKappa = settings.GetDouble("Solver.MaxDeltaKappa", mMaxDeltaKappa)
    mDiagnosticsEnabled = settings.GetBoolean("Solver.DiagnosticsEnabled", mDiagnosticsEnabled)
End Sub

Public Sub Solve(ByVal mesh As CFiberMeshBuilder, ByVal rebars As CRebarLayout, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double)

    ClearResult
    ValidateInputs mesh, concreteMaterial, steelMaterial
    If mUseInitialState Then
        mEpsilon0 = mInitialEpsilon0
        mKappaX = mInitialKappaX
        mKappaY = mInitialKappaY
    End If

    Dim stepIndex As Long
    Dim stepN As Double
    Dim stepMx As Double
    Dim stepMy As Double
    Dim startN As Double
    Dim startMx As Double
    Dim startMy As Double

    If mUseInitialLoads Then
        startN = mInitialTargetN
        startMx = mInitialTargetMx
        startMy = mInitialTargetMy
    End If

    For stepIndex = 1 To mLoadSteps
        stepN = startN + (targetN - startN) * stepIndex / mLoadSteps
        stepMx = startMx + (targetMx - startMx) * stepIndex / mLoadSteps
        stepMy = startMy + (targetMy - startMy) * stepIndex / mLoadSteps

        If Not SolveLoadStep(mesh, rebars, concreteMaterial, steelMaterial, stepIndex, stepN, stepMx, stepMy) Then
            Exit Sub
        End If
        mLoadStepsCompleted = stepIndex
    Next stepIndex

    EvaluateState mesh, rebars, concreteMaterial, steelMaterial, mEpsilon0, mKappaX, mKappaY, True
    SetResiduals targetN, targetMx, targetMy
    mConverged = IsConverged(targetN, targetMx, targetMy)
    If mConverged Then
        mStopReason = "Converged by full Newton with load stepping."
    Else
        mStopReason = "Final residual check failed."
    End If
End Sub

Private Function SolveLoadStep(ByVal mesh As CFiberMeshBuilder, ByVal rebars As CRebarLayout, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal stepIndex As Long, ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double) As Boolean

    Dim iter As Long
    For iter = 1 To mMaxIterations
        mIterations = mIterations + 1
        EvaluateState mesh, rebars, concreteMaterial, steelMaterial, mEpsilon0, mKappaX, mKappaY, True
        SetResiduals targetN, targetMx, targetMy
        If IsConverged(targetN, targetMx, targetMy) Then
            AppendDiagnostics stepIndex, iter, 0#, 0#, targetN, targetMx, targetMy, "converged"
            SolveLoadStep = True
            Exit Function
        End If

        Dim system As CLinearSystem3x3
        Set system = New CLinearSystem3x3
        If Not system.Solve(mK(1, 1), mK(1, 2), mK(1, 3), _
                mK(2, 1), mK(2, 2), mK(2, 3), _
                mK(3, 1), mK(3, 2), mK(3, 3), _
                -mResidualN, -mResidualMx, -mResidualMy) Then
            mConverged = False
            mStopReason = "Tangent matrix solve failed: " & system.StopReason
            AppendDiagnostics stepIndex, iter, 0#, 0#, targetN, targetMx, targetMy, mStopReason
            SolveLoadStep = False
            Exit Function
        End If

        Dim deltaEps As Double
        Dim deltaKx As Double
        Dim deltaKy As Double
        deltaEps = system.X1
        deltaKx = system.X2
        deltaKy = system.X3

        mLastDeltaScale = LimitIncrement(deltaEps, deltaKx, deltaKy)
        mLastAlpha = AcceptStep(mesh, rebars, concreteMaterial, steelMaterial, targetN, targetMx, targetMy, deltaEps, deltaKx, deltaKy)
        AppendDiagnostics stepIndex, iter, mLastAlpha, mLastDeltaScale, targetN, targetMx, targetMy, "iteration"
        If mLastAlpha <= 0# Then
            mConverged = False
            mStopReason = "Line search failed to improve residual."
            SolveLoadStep = False
            Exit Function
        End If
    Next iter

    mConverged = False
    mStopReason = "Maximum iterations reached at load step " & CStr(stepIndex) & "."
    SolveLoadStep = False
End Function

Private Function AcceptStep(ByVal mesh As CFiberMeshBuilder, ByVal rebars As CRebarLayout, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double, _
        ByVal deltaEps As Double, ByVal deltaKx As Double, ByVal deltaKy As Double) As Double

    Dim baseNorm As Double
    baseNorm = ResidualNorm(mResidualN, mResidualMx, mResidualMy, targetN, targetMx, targetMy)

    Dim alpha As Double
    alpha = mDampingInitial
    If alpha <= 0# Or alpha > 1# Then alpha = 1#

    Do While alpha >= mMinLineSearchAlpha
        Dim trialEps As Double
        Dim trialKx As Double
        Dim trialKy As Double
        trialEps = mEpsilon0 + alpha * deltaEps
        trialKx = mKappaX + alpha * deltaKx
        trialKy = mKappaY + alpha * deltaKy

        Dim trialN As Double
        Dim trialMx As Double
        Dim trialMy As Double
        EvaluateStateInto mesh, rebars, concreteMaterial, steelMaterial, trialEps, trialKx, trialKy, False, trialN, trialMx, trialMy
        Dim trialNorm As Double
        trialNorm = ResidualNorm(trialN - targetN, trialMx - targetMx, trialMy - targetMy, targetN, targetMx, targetMy)

        If (Not mLineSearchEnabled) Or trialNorm <= baseNorm Or baseNorm <= 0.000000000001 Then
            mEpsilon0 = trialEps
            mKappaX = trialKx
            mKappaY = trialKy
            AcceptStep = alpha
            Exit Function
        End If

        alpha = alpha / 2#
    Loop

    AcceptStep = 0#
End Function

Private Function LimitIncrement(ByRef deltaEps As Double, ByRef deltaKx As Double, ByRef deltaKy As Double) As Double
    Dim scaleFactor As Double
    scaleFactor = 1#

    If mMaxDeltaEpsilon0 > 0# And Abs(deltaEps) > mMaxDeltaEpsilon0 Then
        scaleFactor = MinDouble(scaleFactor, mMaxDeltaEpsilon0 / Abs(deltaEps))
    End If
    If mMaxDeltaKappa > 0# And Abs(deltaKx) > mMaxDeltaKappa Then
        scaleFactor = MinDouble(scaleFactor, mMaxDeltaKappa / Abs(deltaKx))
    End If
    If mMaxDeltaKappa > 0# And Abs(deltaKy) > mMaxDeltaKappa Then
        scaleFactor = MinDouble(scaleFactor, mMaxDeltaKappa / Abs(deltaKy))
    End If

    deltaEps = deltaEps * scaleFactor
    deltaKx = deltaKx * scaleFactor
    deltaKy = deltaKy * scaleFactor
    LimitIncrement = scaleFactor
End Function

Private Sub EvaluateState(ByVal mesh As CFiberMeshBuilder, ByVal rebars As CRebarLayout, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal eps0 As Double, ByVal kx As Double, ByVal ky As Double, ByVal buildTangent As Boolean)
    EvaluateStateInto mesh, rebars, concreteMaterial, steelMaterial, eps0, kx, ky, buildTangent, mNint, mMxint, mMyint
End Sub

Private Sub EvaluateStateInto(ByVal mesh As CFiberMeshBuilder, ByVal rebars As CRebarLayout, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal eps0 As Double, ByVal kx As Double, ByVal ky As Double, ByVal buildTangent As Boolean, _
        ByRef nInt As Double, ByRef mxInt As Double, ByRef myInt As Double)

    nInt = 0#: mxInt = 0#: myInt = 0#
    If buildTangent Then ClearTangent
    If buildTangent Then
        mMinConcreteStrain = 1E+100
        mMaxConcreteStrain = -1E+100
        mMaxAbsSteelStrain = 0#
    End If

    Dim i As Long
    Dim strain As Double
    Dim stress As Double
    Dim tangent As Double
    Dim force As Double
    Dim area As Double

    For i = 1 To mesh.FiberCount
        strain = eps0 + kx * mesh.FiberY(i) + ky * mesh.FiberX(i)
        If buildTangent Then
            If strain < mMinConcreteStrain Then mMinConcreteStrain = strain
            If strain > mMaxConcreteStrain Then mMaxConcreteStrain = strain
        End If
        stress = concreteMaterial.GetStress(strain)
        tangent = concreteMaterial.GetTangentModulus(strain)
        area = mesh.FiberArea(i) * mesh.FiberFillFactor(i)
        force = stress * area
        nInt = nInt + force
        mxInt = mxInt + force * mesh.FiberY(i)
        myInt = myInt + force * mesh.FiberX(i)
        If buildTangent Then AddTangent tangent * area, mesh.FiberX(i), mesh.FiberY(i)
    Next i

    If Not rebars Is Nothing Then
        For i = 1 To rebars.Count
            strain = eps0 + kx * rebars.Y(i) + ky * rebars.X(i)
            If buildTangent Then
                If Abs(strain) > mMaxAbsSteelStrain Then mMaxAbsSteelStrain = Abs(strain)
            End If
            Dim steelStress As Double
            Dim steelTangent As Double
            Dim concreteStress As Double
            Dim concreteTangent As Double
            steelStress = steelMaterial.GetStress(strain)
            steelTangent = steelMaterial.GetTangentModulus(strain)
            concreteStress = concreteMaterial.GetStress(strain)
            concreteTangent = concreteMaterial.GetTangentModulus(strain)
            stress = steelStress - concreteStress
            tangent = steelTangent - concreteTangent
            force = stress * rebars.Area(i)
            nInt = nInt + force
            mxInt = mxInt + force * rebars.Y(i)
            myInt = myInt + force * rebars.X(i)
            If buildTangent Then AddTangent tangent * rebars.Area(i), rebars.X(i), rebars.Y(i)
        Next i
    End If
End Sub

Private Sub ClearTangent()
    Dim i As Long
    For i = 1 To 3
        mK(i, 1) = 0#: mK(i, 2) = 0#: mK(i, 3) = 0#
    Next i
End Sub

Private Sub AddTangent(ByVal axialStiffness As Double, ByVal xCoord As Double, ByVal yCoord As Double)
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

Private Sub SetResiduals(ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double)
    mResidualN = mNint - targetN
    mResidualMx = mMxint - targetMx
    mResidualMy = mMyint - targetMy
    mRelativeResidualN = RelativeResidual(mResidualN, targetN)
    mRelativeResidualMx = RelativeResidual(mResidualMx, targetMx)
    mRelativeResidualMy = RelativeResidual(mResidualMy, targetMy)
End Sub

Private Function IsConverged(ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double) As Boolean
    IsConverged = Abs(mResidualN) <= mToleranceN And _
        Abs(mResidualMx) <= mToleranceMx And _
        Abs(mResidualMy) <= mToleranceMy
End Function

Private Function ResidualNorm(ByVal resN As Double, ByVal resMx As Double, ByVal resMy As Double, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double) As Double
    ResidualNorm = Abs(resN) / MaxDouble(mToleranceN, Abs(targetN) * 0.000000001) + _
        Abs(resMx) / MaxDouble(mToleranceMx, Abs(targetMx) * 0.000000001) + _
        Abs(resMy) / MaxDouble(mToleranceMy, Abs(targetMy) * 0.000000001)
End Function

Private Function RelativeResidual(ByVal residual As Double, ByVal target As Double) As Double
    If Abs(target) > 0.000000000001 Then
        RelativeResidual = Abs(residual / target)
    Else
        RelativeResidual = Abs(residual)
    End If
End Function

Private Sub AppendDiagnostics(ByVal stepIndex As Long, ByVal iter As Long, ByVal alpha As Double, _
        ByVal deltaScale As Double, ByVal targetN As Double, ByVal targetMx As Double, _
        ByVal targetMy As Double, ByVal state As String)
    If Not mDiagnosticsEnabled Then Exit Sub
    mDiagnosticLog = mDiagnosticLog & "step=" & CStr(stepIndex) & _
        "; iter=" & CStr(iter) & _
        "; state=" & state & _
        "; eps0=" & FormatNumberInvariant(mEpsilon0) & _
        "; kappaX=" & FormatNumberInvariant(mKappaX) & _
        "; kappaY=" & FormatNumberInvariant(mKappaY) & _
        "; RN=" & FormatNumberInvariant(mResidualN) & _
        "; RMx=" & FormatNumberInvariant(mResidualMx) & _
        "; RMy=" & FormatNumberInvariant(mResidualMy) & _
        "; alpha=" & FormatNumberInvariant(alpha) & _
        "; deltaScale=" & FormatNumberInvariant(deltaScale) & vbCrLf
End Sub

Private Sub ValidateInputs(ByVal mesh As CFiberMeshBuilder, ByVal concreteMaterial As Object, ByVal steelMaterial As Object)
    If mesh Is Nothing Then Err.Raise vbObjectError + 3400, "CSectionSolver", "Concrete mesh is missing."
    If mesh.FiberCount <= 0 Then Err.Raise vbObjectError + 3401, "CSectionSolver", "Concrete mesh is empty."
    If concreteMaterial Is Nothing Then Err.Raise vbObjectError + 3402, "CSectionSolver", "Concrete material is missing."
    If steelMaterial Is Nothing Then Err.Raise vbObjectError + 3403, "CSectionSolver", "Steel material is missing."
    If mMaxIterations <= 0 Then Err.Raise vbObjectError + 3404, "CSectionSolver", "MaxIterations must be positive."
    If mLoadSteps <= 0 Then Err.Raise vbObjectError + 3405, "CSectionSolver", "LoadSteps must be positive."
End Sub

Public Sub ClearResult()
    mConverged = False
    mStopReason = vbNullString
    mEpsilon0 = 0#: mKappaX = 0#: mKappaY = 0#
    mNint = 0#: mMxint = 0#: mMyint = 0#
    mResidualN = 0#: mResidualMx = 0#: mResidualMy = 0#
    mRelativeResidualN = 0#: mRelativeResidualMx = 0#: mRelativeResidualMy = 0#
    mMinConcreteStrain = 0#: mMaxConcreteStrain = 0#: mMaxAbsSteelStrain = 0#
    mIterations = 0
    mLoadStepsCompleted = 0
    mDiagnosticLog = vbNullString
    mLastAlpha = 0#
    mLastDeltaScale = 0#
    ClearTangent
End Sub

Public Sub SetInitialState(ByVal epsilon0 As Double, ByVal kappaX As Double, ByVal kappaY As Double)
    mUseInitialState = True
    mInitialEpsilon0 = epsilon0
    mInitialKappaX = kappaX
    mInitialKappaY = kappaY
End Sub

Public Sub SetInitialLoadState(ByVal epsilon0 As Double, ByVal kappaX As Double, ByVal kappaY As Double, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double)
    SetInitialState epsilon0, kappaX, kappaY
    mUseInitialLoads = True
    mInitialTargetN = targetN
    mInitialTargetMx = targetMx
    mInitialTargetMy = targetMy
End Sub

Public Sub ClearInitialState()
    mUseInitialState = False
    mInitialEpsilon0 = 0#: mInitialKappaX = 0#: mInitialKappaY = 0#
    mUseInitialLoads = False
    mInitialTargetN = 0#: mInitialTargetMx = 0#: mInitialTargetMy = 0#
End Sub

Private Function MinDouble(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then MinDouble = a Else MinDouble = b
End Function

Private Function MaxDouble(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then MaxDouble = a Else MaxDouble = b
End Function

Private Function FormatNumberInvariant(ByVal value As Double) As String
    FormatNumberInvariant = Replace$(Format$(value, "0.############"), ",", ".")
End Function

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

Public Property Get MaxAbsSteelStrain() As Double
    MaxAbsSteelStrain = mMaxAbsSteelStrain
End Property

Public Property Get Iterations() As Long
    Iterations = mIterations
End Property

Public Property Get LoadStepsCompleted() As Long
    LoadStepsCompleted = mLoadStepsCompleted
End Property

Public Property Get Tangent(ByVal rowIndex As Long, ByVal columnIndex As Long) As Double
    Tangent = mK(rowIndex, columnIndex)
End Property

Public Property Get DiagnosticLog() As String
    DiagnosticLog = mDiagnosticLog
End Property

Public Property Let MaxIterations(ByVal value As Long)
    mMaxIterations = value
End Property

Public Property Let LoadSteps(ByVal value As Long)
    mLoadSteps = value
End Property

Public Property Let ToleranceN(ByVal value As Double)
    mToleranceN = value
End Property

Public Property Let ToleranceMx(ByVal value As Double)
    mToleranceMx = value
End Property

Public Property Let ToleranceMy(ByVal value As Double)
    mToleranceMy = value
End Property

Public Property Let LineSearchEnabled(ByVal value As Boolean)
    mLineSearchEnabled = value
End Property

Public Property Let DiagnosticsEnabled(ByVal value As Boolean)
    mDiagnosticsEnabled = value
End Property

Public Property Let MaxDeltaEpsilon0(ByVal value As Double)
    mMaxDeltaEpsilon0 = value
End Property

Public Property Let MaxDeltaKappa(ByVal value As Double)
    mMaxDeltaKappa = value
End Property












