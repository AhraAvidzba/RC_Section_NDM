VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCapacitySolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const STATE_ACCEPTABLE As String = "Acceptable"
Private Const STATE_CONCRETE_LIMIT As String = "ConcreteStrainLimit"
Private Const STATE_STEEL_LIMIT As String = "SteelStrainLimit"
Private Const STATE_NUMERICAL_FAILURE As String = "NumericalFailure"
Private Const STATE_SINGULAR_TANGENT As String = "SingularTangent"
Private Const STATE_INVALID_INPUT As String = "InvalidInput"

Private mConverged As Boolean
Private mStopReason As String
Private mLimitState As String
Private mUpperLimitState As String
Private mLambdaUltimate As Double
Private mMomentUltimate As Double
Private mMxUltimate As Double
Private mMyUltimate As Double
Private mSafetyFactor As Double
Private mUtilization As Double
Private mIterations As Long
Private mRetryCount As Long
Private mLastSolver As CSectionSolver
Private mLastAcceptableSolver As CSectionSolver
Private mLastAcceptableN As Double
Private mLastAcceptableMx As Double
Private mLastAcceptableMy As Double
Private mConcreteCompressionLimit As Double
Private mSteelStrainLimit As Double
Private mLambdaTolerance As Double
Private mMaxLambda As Double
Private mInitialLambdaStep As Double
Private mMaxRetries As Long
Private mSolverBaseLoadSteps As Long
Private mSolverMaxIterations As Long
Private mSolverToleranceN As Double
Private mSolverToleranceMx As Double
Private mSolverToleranceMy As Double
Private mSolverMaxDeltaEpsilon0 As Double
Private mSolverMaxDeltaKappa As Double
Private mMomentEquilibriumResidual As Double
Private mDiagnosticLog As String

Private Sub Class_Initialize()
    mConcreteCompressionLimit = -0.0035
    mSteelStrainLimit = 0.025
    mLambdaTolerance = 0.001
    mMaxLambda = 100#
    mInitialLambdaStep = 1#
    mMaxRetries = 4
    mSolverBaseLoadSteps = 8
    mSolverMaxIterations = 60
    mSolverToleranceN = 5#
    mSolverToleranceMx = 5000#
    mSolverToleranceMy = 5000#
    mSolverMaxDeltaEpsilon0 = 0.0003
    mSolverMaxDeltaKappa = 0.000005
End Sub

Public Sub ApplySettings(ByVal settings As CSystemSettingsReader)
    If settings Is Nothing Then Exit Sub
    mInitialLambdaStep = settings.GetDouble("Capacity.InitialLambda", mInitialLambdaStep)
    mMaxLambda = settings.GetDouble("Capacity.MaxLambda", mMaxLambda)
    mLambdaTolerance = settings.GetDouble("Capacity.ToleranceLambda", mLambdaTolerance)
    mMaxRetries = settings.GetLong("Capacity.MaxRetries", mMaxRetries)
    mSolverBaseLoadSteps = settings.GetLong("Capacity.BaseLoadSteps", mSolverBaseLoadSteps)
    mSolverMaxIterations = settings.GetLong("Capacity.SolverMaxIterations", mSolverMaxIterations)
    mSolverToleranceN = settings.GetDouble("Solver.ToleranceN", mSolverToleranceN)
    mSolverToleranceMx = settings.GetDouble("Solver.ToleranceMx", mSolverToleranceMx)
    mSolverToleranceMy = settings.GetDouble("Solver.ToleranceMy", mSolverToleranceMy)
    mSolverMaxDeltaEpsilon0 = settings.GetDouble("Solver.MaxDeltaEpsilon0", mSolverMaxDeltaEpsilon0)
    mSolverMaxDeltaKappa = settings.GetDouble("Solver.MaxDeltaKappa", mSolverMaxDeltaKappa)
    mConcreteCompressionLimit = settings.GetDouble("Capacity.ConcreteCompressionLimit", mConcreteCompressionLimit)
    mSteelStrainLimit = settings.GetDouble("Capacity.SteelStrainLimit", mSteelStrainLimit)
End Sub

Public Sub SolveByLoadMultiplier(ByVal mesh As CFiberMeshBuilder, ByVal rebars As CRebarLayout, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal nSpecified As Double, ByVal mxBase As Double, ByVal myBase As Double)
    ClearResult

    If Not ValidateCapacityInputs(mesh, concreteMaterial, steelMaterial, mxBase, myBase) Then Exit Sub

    Dim lowLambda As Double
    Dim highLambda As Double
    Dim probeState As String
    lowLambda = 0#
    highLambda = mInitialLambdaStep
    If highLambda <= 0# Then highLambda = 1#

    probeState = ProbeWithRetries(mesh, rebars, concreteMaterial, steelMaterial, nSpecified, 0#, 0#, lowLambda)
    If probeState <> STATE_ACCEPTABLE Then
        mLimitState = probeState
        mStopReason = "Initial N-only state is not acceptable: " & probeState & "."
        Exit Sub
    End If
    RecordAcceptableState mLastSolver, nSpecified, 0#, 0#

    Dim bracketFound As Boolean
    Do While highLambda <= mMaxLambda
        probeState = ProbeWithRetries(mesh, rebars, concreteMaterial, steelMaterial, nSpecified, highLambda * mxBase, highLambda * myBase, highLambda)
        If probeState = STATE_ACCEPTABLE Then
            lowLambda = highLambda
            RecordAcceptableState mLastSolver, nSpecified, highLambda * mxBase, highLambda * myBase
            highLambda = highLambda * 2#
        ElseIf IsPhysicalLimit(probeState) Then
            mUpperLimitState = probeState
            bracketFound = True
            Exit Do
        ElseIf IsNumericalFailure(probeState) Then
            highLambda = ReduceLambdaAfterNumericalFailure(lowLambda, highLambda, probeState)
            If highLambda <= lowLambda Then Exit Sub
        Else
            mLimitState = probeState
            mStopReason = "Capacity search stopped: " & probeState & "."
            Exit Sub
        End If
    Loop

    If Not bracketFound Then
        If Len(mLimitState) = 0 Then mLimitState = STATE_NUMERICAL_FAILURE
        If Len(mStopReason) = 0 Then mStopReason = "Physical upper lambda was not found before MaxLambda."
        mLambdaUltimate = lowLambda
        Exit Sub
    End If

    Dim midLambda As Double
    Do While (highLambda - lowLambda) > mLambdaTolerance
        midLambda = (lowLambda + highLambda) / 2#
        probeState = ProbeWithRetries(mesh, rebars, concreteMaterial, steelMaterial, nSpecified, midLambda * mxBase, midLambda * myBase, midLambda)
        If probeState = STATE_ACCEPTABLE Then
            lowLambda = midLambda
            RecordAcceptableState mLastSolver, nSpecified, midLambda * mxBase, midLambda * myBase
        ElseIf IsPhysicalLimit(probeState) Then
            highLambda = midLambda
            mUpperLimitState = probeState
        ElseIf IsNumericalFailure(probeState) Then
            mLimitState = probeState
            mStopReason = "Numerical failure occurred inside physical bracket; previous converged lambda retained."
            mLambdaUltimate = lowLambda
            If Not mLastAcceptableSolver Is Nothing Then Set mLastSolver = mLastAcceptableSolver
            Exit Sub
        Else
            mLimitState = probeState
            mStopReason = "Capacity bisection stopped: " & probeState & "."
            Exit Sub
        End If
    Loop

    mLambdaUltimate = lowLambda
    mMxUltimate = mLambdaUltimate * mxBase
    mMyUltimate = mLambdaUltimate * myBase
    mMomentUltimate = Sqr(mMxUltimate * mMxUltimate + mMyUltimate * mMyUltimate)
    If Not mLastAcceptableSolver Is Nothing Then
        mMomentEquilibriumResidual = Sqr((mLastAcceptableSolver.Mxint - mMxUltimate) ^ 2 + _
            (mLastAcceptableSolver.Myint - mMyUltimate) ^ 2)
    End If
    mSafetyFactor = mLambdaUltimate
    If mLambdaUltimate > 0# Then mUtilization = 1# / mLambdaUltimate
    If Not mLastAcceptableSolver Is Nothing Then Set mLastSolver = mLastAcceptableSolver
    mLimitState = mUpperLimitState
    mConverged = True
    mStopReason = "Capacity search converged by bisection with physical upper state: " & mUpperLimitState & "."
End Sub

Private Function ValidateCapacityInputs(ByVal mesh As CFiberMeshBuilder, ByVal concreteMaterial As Object, _
        ByVal steelMaterial As Object, ByVal mxBase As Double, ByVal myBase As Double) As Boolean
    If mesh Is Nothing Then
        SetInvalidInput "Concrete mesh is missing."
    ElseIf mesh.FiberCount <= 0 Then
        SetInvalidInput "Concrete mesh is empty."
    ElseIf concreteMaterial Is Nothing Then
        SetInvalidInput "Concrete material is missing."
    ElseIf steelMaterial Is Nothing Then
        SetInvalidInput "Steel material is missing."
    ElseIf Sqr(mxBase * mxBase + myBase * myBase) <= 0.000000001 Then
        SetInvalidInput "Base moment vector must be non-zero."
    ElseIf mLambdaTolerance <= 0# Or mMaxLambda <= 0# Or mMaxRetries < 0 Or mSolverBaseLoadSteps <= 0 Or mSolverMaxIterations <= 0 Then
        SetInvalidInput "Capacity solver settings are invalid."
    Else
        ValidateCapacityInputs = True
    End If
End Function

Private Sub SetInvalidInput(ByVal reason As String)
    mLimitState = STATE_INVALID_INPUT
    mStopReason = reason
End Sub

Private Function ProbeWithRetries(ByVal mesh As CFiberMeshBuilder, ByVal rebars As CRebarLayout, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal nSpecified As Double, ByVal mx As Double, ByVal my As Double, ByVal lambdaValue As Double) As String

    Dim attempt As Long
    Dim probeState As String
    For attempt = 0 To mMaxRetries
        probeState = ProbeOnce(mesh, rebars, concreteMaterial, steelMaterial, nSpecified, mx, my, lambdaValue, attempt)
        If Not IsNumericalFailure(probeState) Then
            ProbeWithRetries = probeState
            Exit Function
        End If
        If attempt < mMaxRetries Then mRetryCount = mRetryCount + 1
    Next attempt

    ProbeWithRetries = probeState
End Function

Private Function ProbeOnce(ByVal mesh As CFiberMeshBuilder, ByVal rebars As CRebarLayout, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal nSpecified As Double, ByVal mx As Double, ByVal my As Double, _
        ByVal lambdaValue As Double, ByVal attempt As Long) As String

    On Error GoTo Failed
    mIterations = mIterations + 1

    Dim solver As CSectionSolver
    Set solver = New CSectionSolver
    solver.LoadSteps = mSolverBaseLoadSteps * (2 ^ attempt)
    solver.MaxIterations = mSolverMaxIterations
    solver.ToleranceN = mSolverToleranceN
    solver.ToleranceMx = mSolverToleranceMx
    solver.ToleranceMy = mSolverToleranceMy
    solver.MaxDeltaEpsilon0 = mSolverMaxDeltaEpsilon0
    solver.MaxDeltaKappa = mSolverMaxDeltaKappa / (2 ^ attempt)
    If Not mLastAcceptableSolver Is Nothing Then
        solver.SetInitialLoadState mLastAcceptableSolver.Epsilon0, mLastAcceptableSolver.KappaX, mLastAcceptableSolver.KappaY, _
            mLastAcceptableN, mLastAcceptableMx, mLastAcceptableMy
    End If

    solver.Solve mesh, rebars, concreteMaterial, steelMaterial, nSpecified, mx, my
    Set mLastSolver = solver

    If Not solver.Converged Then
        ProbeOnce = ClassifyNumericalFailure(solver.StopReason)
    ElseIf solver.MinConcreteStrain < mConcreteCompressionLimit Then
        ProbeOnce = STATE_CONCRETE_LIMIT
    ElseIf solver.MaxAbsSteelStrain > mSteelStrainLimit Then
        ProbeOnce = STATE_STEEL_LIMIT
    Else
        ProbeOnce = STATE_ACCEPTABLE
    End If

    AppendDiagnostic solver, mx, my, lambdaValue, attempt, ProbeOnce
    Exit Function

Failed:
    ProbeOnce = STATE_INVALID_INPUT
    mLimitState = STATE_INVALID_INPUT
    mStopReason = "Capacity probe failed: " & Err.Description
End Function

Private Function ClassifyNumericalFailure(ByVal reason As String) As String
    If InStr(1, reason, "Tangent matrix", vbTextCompare) > 0 Or _
            InStr(1, reason, "singular", vbTextCompare) > 0 Or _
            InStr(1, reason, "solve failed", vbTextCompare) > 0 Then
        ClassifyNumericalFailure = STATE_SINGULAR_TANGENT
    Else
        ClassifyNumericalFailure = STATE_NUMERICAL_FAILURE
    End If
End Function

Private Sub RecordAcceptableState(ByVal solver As CSectionSolver, ByVal nValue As Double, ByVal mxValue As Double, ByVal myValue As Double)
    Set mLastAcceptableSolver = solver
    mLastAcceptableN = nValue
    mLastAcceptableMx = mxValue
    mLastAcceptableMy = myValue
End Sub

Private Function ReduceLambdaAfterNumericalFailure(ByVal lowLambda As Double, ByVal highLambda As Double, ByVal probeState As String) As Double
    Dim reduced As Double
    reduced = lowLambda + (highLambda - lowLambda) / 2#
    If (reduced - lowLambda) <= mLambdaTolerance Then
        mLimitState = probeState
        mStopReason = "Numerical failure before a physical upper boundary; last converged lambda retained."
        mLambdaUltimate = lowLambda
        If Not mLastAcceptableSolver Is Nothing Then Set mLastSolver = mLastAcceptableSolver
        ReduceLambdaAfterNumericalFailure = lowLambda
    Else
        AppendSearchDiagnostic "numerical failure; reducing lambda high from " & FormatNumberInvariant(highLambda) & _
            " to " & FormatNumberInvariant(reduced) & "; state=" & probeState
        ReduceLambdaAfterNumericalFailure = reduced
    End If
End Function

Private Function IsPhysicalLimit(ByVal state As String) As Boolean
    IsPhysicalLimit = (state = STATE_CONCRETE_LIMIT Or state = STATE_STEEL_LIMIT)
End Function

Private Function IsNumericalFailure(ByVal state As String) As Boolean
    IsNumericalFailure = (state = STATE_NUMERICAL_FAILURE Or state = STATE_SINGULAR_TANGENT)
End Function

Private Sub AppendDiagnostic(ByVal solver As CSectionSolver, ByVal mx As Double, ByVal my As Double, _
        ByVal lambdaValue As Double, ByVal attempt As Long, ByVal probeState As String)
    mDiagnosticLog = mDiagnosticLog & "probe=" & CStr(mIterations) & _
        "; attempt=" & CStr(attempt) & _
        "; lambda=" & FormatNumberInvariant(lambdaValue) & _
        "; Mx=" & FormatNumberInvariant(mx) & _
        "; My=" & FormatNumberInvariant(my) & _
        "; converged=" & CStr(solver.Converged) & _
        "; state=" & probeState & _
        "; stop=" & solver.StopReason & _
        "; minConcreteStrain=" & FormatNumberInvariant(solver.MinConcreteStrain) & _
        "; maxAbsSteelStrain=" & FormatNumberInvariant(solver.MaxAbsSteelStrain) & vbCrLf
End Sub

Private Sub AppendSearchDiagnostic(ByVal message As String)
    mDiagnosticLog = mDiagnosticLog & "search=" & message & vbCrLf
End Sub

Private Sub ClearResult()
    mConverged = False
    mStopReason = vbNullString
    mLimitState = vbNullString
    mUpperLimitState = vbNullString
    mLambdaUltimate = 0#
    mMomentUltimate = 0#
    mMxUltimate = 0#: mMyUltimate = 0#
    mSafetyFactor = 0#
    mUtilization = 0#
    mIterations = 0
    mRetryCount = 0
    mMomentEquilibriumResidual = 0#
    mLastAcceptableN = 0#: mLastAcceptableMx = 0#: mLastAcceptableMy = 0#
    Set mLastSolver = Nothing
    Set mLastAcceptableSolver = Nothing
    mDiagnosticLog = vbNullString
End Sub

Private Function FormatNumberInvariant(ByVal value As Double) As String
    FormatNumberInvariant = Replace$(Format$(value, "0.############"), ",", ".")
End Function

Public Property Get Converged() As Boolean
    Converged = mConverged
End Property

Public Property Get StopReason() As String
    StopReason = mStopReason
End Property

Public Property Get LimitState() As String
    LimitState = mLimitState
End Property

Public Property Get UpperLimitState() As String
    UpperLimitState = mUpperLimitState
End Property

Public Property Get LambdaUltimate() As Double
    LambdaUltimate = mLambdaUltimate
End Property

Public Property Get MomentUltimate() As Double
    MomentUltimate = mMomentUltimate
End Property

Public Property Get MxUltimate() As Double
    MxUltimate = mMxUltimate
End Property

Public Property Get MyUltimate() As Double
    MyUltimate = mMyUltimate
End Property

Public Property Get SafetyFactor() As Double
    SafetyFactor = mSafetyFactor
End Property

Public Property Get Utilization() As Double
    Utilization = mUtilization
End Property

Public Property Get Iterations() As Long
    Iterations = mIterations
End Property

Public Property Get RetryCount() As Long
    RetryCount = mRetryCount
End Property

Public Property Get MomentEquilibriumResidual() As Double
    MomentEquilibriumResidual = mMomentEquilibriumResidual
End Property

Public Property Get DiagnosticLog() As String
    DiagnosticLog = mDiagnosticLog
End Property

Public Property Get LastSolver() As CSectionSolver
    Set LastSolver = mLastSolver
End Property

Public Property Let ConcreteCompressionLimit(ByVal value As Double)
    mConcreteCompressionLimit = value
End Property

Public Property Let SteelStrainLimit(ByVal value As Double)
    mSteelStrainLimit = value
End Property

Public Property Let LambdaTolerance(ByVal value As Double)
    mLambdaTolerance = value
End Property

Public Property Let MaxLambda(ByVal value As Double)
    mMaxLambda = value
End Property

Public Property Let InitialLambdaStep(ByVal value As Double)
    mInitialLambdaStep = value
End Property

Public Property Let MaxRetries(ByVal value As Long)
    mMaxRetries = value
End Property

Public Property Let SolverBaseLoadSteps(ByVal value As Long)
    mSolverBaseLoadSteps = value
End Property

Public Property Let SolverMaxIterations(ByVal value As Long)
    mSolverMaxIterations = value
End Property

Public Property Let SolverToleranceN(ByVal value As Double)
    mSolverToleranceN = value
End Property

Public Property Let SolverToleranceMx(ByVal value As Double)
    mSolverToleranceMx = value
End Property

Public Property Let SolverToleranceMy(ByVal value As Double)
    mSolverToleranceMy = value
End Property

Public Property Let SolverMaxDeltaEpsilon0(ByVal value As Double)
    mSolverMaxDeltaEpsilon0 = value
End Property

Public Property Let SolverMaxDeltaKappa(ByVal value As Double)
    mSolverMaxDeltaKappa = value
End Property







