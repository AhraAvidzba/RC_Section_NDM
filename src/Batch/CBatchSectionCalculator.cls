VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CBatchSectionCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_COMBINATIONS As Long = 20
Private Const STATUS_OK As String = "OK"
Private Const STATUS_INVALID_INPUT As String = "InvalidInput"
Private Const STATUS_SKIPPED As String = "SkippedZeroBaseMoment"
Private Const STATUS_NUMERICAL_FAILURE As String = "NumericalFailure"
Private Const STATUS_NOT_CALCULATED As String = "NotCalculated"

Private mMesh As CFiberMeshBuilder
Private mRebars As CRebarLayout
Private mConcreteMaterial As Object
Private mSteelMaterial As Object
Private mCount As Long

Private mCombinationID(1 To MAX_COMBINATIONS) As String
Private mN(1 To MAX_COMBINATIONS) As Double
Private mMx(1 To MAX_COMBINATIONS) As Double
Private mMy(1 To MAX_COMBINATIONS) As Double
Private mCalculationType(1 To MAX_COMBINATIONS) As String
Private mDurationType(1 To MAX_COMBINATIONS) As String
Private mComment(1 To MAX_COMBINATIONS) As String
Private mInputValid(1 To MAX_COMBINATIONS) As Boolean
Private mInputError(1 To MAX_COMBINATIONS) As String

Private mStatus(1 To MAX_COMBINATIONS) As String
Private mCapacityStatus(1 To MAX_COMBINATIONS) As String
Private mCrackStatus(1 To MAX_COMBINATIONS) As String
Private mLambdaCapacity(1 To MAX_COMBINATIONS) As Double
Private mUtilCapacity(1 To MAX_COMBINATIONS) As Double
Private mCrackWidth(1 To MAX_COMBINATIONS) As Double
Private mCrackUtil(1 To MAX_COMBINATIONS) As Double
Private mGoverningUtil(1 To MAX_COMBINATIONS) As Double
Private mGoverningCombinationIndex As Long
Private mElapsedSeconds As Double
Private mDiagnosticLog As String

Private mCrackAllowable As Double
Private mCrackSpacing As Double
Private mCrackStrainFactor As Double
Private mCrackDurationFactor As Double
Private mCrackEnabled As Boolean
Private mCalculationMode As String
Private mSolverLoadSteps As Long
Private mSolverMaxIterations As Long
Private mSolverToleranceN As Double
Private mSolverToleranceMx As Double
Private mSolverToleranceMy As Double
Private mSolverDiagnosticsEnabled As Boolean
Private mSolverMaxDeltaEpsilon0 As Double
Private mSolverMaxDeltaKappa As Double
Private mCapacityInitialLambda As Double
Private mCapacityMaxLambda As Double
Private mCapacityToleranceLambda As Double
Private mCapacityMaxRetries As Long
Private mCapacityBaseLoadSteps As Long
Private mCapacitySolverMaxIterations As Long
Private mCapacityConcreteCompressionLimit As Double
Private mCapacitySteelStrainLimit As Double

Private Sub Class_Initialize()
    mCrackAllowable = 0.3
    mCrackSpacing = 200#
    mCrackStrainFactor = 1#
    mCrackDurationFactor = 1#
    mCalculationMode = "FullCapacity"
    mCrackEnabled = True
    mSolverLoadSteps = 8
    mSolverMaxIterations = 80
    mSolverToleranceN = 5#
    mSolverToleranceMx = 5000#
    mSolverToleranceMy = 5000#
    mSolverDiagnosticsEnabled = True
    mSolverMaxDeltaEpsilon0 = 0.0005
    mSolverMaxDeltaKappa = 0.00001
    mCapacityInitialLambda = 1#
    mCapacityMaxLambda = 64#
    mCapacityToleranceLambda = 0.01
    mCapacityMaxRetries = 0
    mCapacityBaseLoadSteps = 1
    mCapacitySolverMaxIterations = 60
    mCapacityConcreteCompressionLimit = -0.0035
    mCapacitySteelStrainLimit = 0.025
End Sub

Public Sub Initialize(ByVal mesh As CFiberMeshBuilder, ByVal rebars As CRebarLayout, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object)
    If mesh Is Nothing Then Err.Raise vbObjectError + 3900, "CBatchSectionCalculator", "Mesh is missing."
    If rebars Is Nothing Then Err.Raise vbObjectError + 3901, "CBatchSectionCalculator", "Rebar layout is missing."
    If concreteMaterial Is Nothing Then Err.Raise vbObjectError + 3902, "CBatchSectionCalculator", "Concrete material is missing."
    If steelMaterial Is Nothing Then Err.Raise vbObjectError + 3903, "CBatchSectionCalculator", "Steel material is missing."
    Set mMesh = mesh
    Set mRebars = rebars
    Set mConcreteMaterial = concreteMaterial
    Set mSteelMaterial = steelMaterial
End Sub

Public Sub ApplySettings(ByVal settings As CSystemSettingsReader)
    If settings Is Nothing Then Exit Sub
    mCrackAllowable = settings.GetDouble("CrackWidth.Allowable", mCrackAllowable)
    mCrackSpacing = settings.GetDouble("CrackWidth.CrackSpacing", mCrackSpacing)
    mCrackStrainFactor = settings.GetDouble("CrackWidth.StrainFactor", mCrackStrainFactor)
    mCrackDurationFactor = settings.GetDouble("CrackWidth.DurationFactor", mCrackDurationFactor)
    mCalculationMode = settings.GetString("Calculation.Mode", mCalculationMode)
    mCrackEnabled = settings.GetBoolean("CrackWidth.Enabled", mCrackEnabled)
    mSolverLoadSteps = settings.GetLong("Solver.LoadSteps", mSolverLoadSteps)
    mSolverMaxIterations = settings.GetLong("Solver.MaxIterations", mSolverMaxIterations)
    mSolverToleranceN = settings.GetDouble("Solver.ToleranceN", mSolverToleranceN)
    mSolverToleranceMx = settings.GetDouble("Solver.ToleranceMx", mSolverToleranceMx)
    mSolverToleranceMy = settings.GetDouble("Solver.ToleranceMy", mSolverToleranceMy)
    mSolverDiagnosticsEnabled = settings.GetBoolean("Solver.DiagnosticsEnabled", mSolverDiagnosticsEnabled)
    mSolverMaxDeltaEpsilon0 = settings.GetDouble("Solver.MaxDeltaEpsilon0", mSolverMaxDeltaEpsilon0)
    mSolverMaxDeltaKappa = settings.GetDouble("Solver.MaxDeltaKappa", mSolverMaxDeltaKappa)
    mCapacityInitialLambda = settings.GetDouble("Capacity.InitialLambda", mCapacityInitialLambda)
    mCapacityMaxLambda = settings.GetDouble("Capacity.MaxLambda", mCapacityMaxLambda)
    mCapacityToleranceLambda = settings.GetDouble("Capacity.ToleranceLambda", mCapacityToleranceLambda)
    mCapacityMaxRetries = settings.GetLong("Capacity.MaxRetries", mCapacityMaxRetries)
    mCapacityBaseLoadSteps = settings.GetLong("Capacity.BaseLoadSteps", mCapacityBaseLoadSteps)
    mCapacitySolverMaxIterations = settings.GetLong("Capacity.SolverMaxIterations", mCapacitySolverMaxIterations)
    mCapacityConcreteCompressionLimit = settings.GetDouble("Capacity.ConcreteCompressionLimit", mCapacityConcreteCompressionLimit)
    mCapacitySteelStrainLimit = settings.GetDouble("Capacity.SteelStrainLimit", mCapacitySteelStrainLimit)
End Sub

Public Sub ClearCombinations()
    mCount = 0
    mGoverningCombinationIndex = 0
    mElapsedSeconds = 0#
    mDiagnosticLog = vbNullString
End Sub

Public Sub AddCombination(ByVal combinationID As String, ByVal nValue As Double, ByVal mxValue As Double, _
        ByVal myValue As Double, ByVal calculationType As String, ByVal durationType As String, ByVal comment As String)
    AddCombinationInternal combinationID, nValue, mxValue, myValue, calculationType, durationType, comment, True, vbNullString
End Sub

Public Sub AddInvalidCombination(ByVal combinationID As String, ByVal calculationType As String, _
        ByVal durationType As String, ByVal comment As String, ByVal reason As String)
    AddCombinationInternal combinationID, 0#, 0#, 0#, calculationType, durationType, comment, False, reason
End Sub

Private Sub AddCombinationInternal(ByVal combinationID As String, ByVal nValue As Double, ByVal mxValue As Double, _
        ByVal myValue As Double, ByVal calculationType As String, ByVal durationType As String, _
        ByVal comment As String, ByVal inputValid As Boolean, ByVal inputError As String)
    If mCount >= MAX_COMBINATIONS Then Err.Raise vbObjectError + 3911, "CBatchSectionCalculator", "Maximum 20 combinations are supported."
    mCount = mCount + 1
    If Len(Trim$(combinationID)) = 0 Then combinationID = "ROW" & CStr(mCount)
    mCombinationID(mCount) = combinationID
    mN(mCount) = nValue
    mMx(mCount) = mxValue
    mMy(mCount) = myValue
    mCalculationType(mCount) = calculationType
    mDurationType(mCount) = durationType
    mComment(mCount) = comment
    mInputValid(mCount) = inputValid
    mInputError(mCount) = inputError
End Sub

Public Sub Execute()
    ValidateReady
    Dim t0 As Double
    t0 = Timer
    mGoverningCombinationIndex = 0
    mDiagnosticLog = vbNullString

    Dim i As Long
    For i = 1 To mCount
        ClearResultRow i
        RunCombination i
        If mInputValid(i) Then
            If mGoverningCombinationIndex = 0 Then
                mGoverningCombinationIndex = i
            ElseIf mGoverningUtil(i) > mGoverningUtil(mGoverningCombinationIndex) Then
                mGoverningCombinationIndex = i
            End If
        End If
    Next i
    mElapsedSeconds = Timer - t0
End Sub

Private Sub RunCombination(ByVal index As Long)
    On Error GoTo Failed
    If Not mInputValid(index) Then
        MarkInvalid index, mInputError(index)
        Exit Sub
    End If

    mStatus(index) = STATUS_OK
    If IsFullCapacityMode Then
        RunCapacity index
    Else
        StoreNotCalculatedCapacity index
    End If
    RunCrack index

    mGoverningUtil(index) = MaxDouble(mUtilCapacity(index), mCrackUtil(index))
    AppendDiagnostic index
    Exit Sub

Failed:
    MarkInvalid index, Err.Description
End Sub

Private Sub RunCapacity(ByVal index As Long)
    If Sqr(mMx(index) * mMx(index) + mMy(index) * mMy(index)) <= 0.000000001 Then
        StoreSkippedCapacity index
        Exit Sub
    End If

    Dim cap As CCapacitySolver
    Set cap = New CCapacitySolver
    ConfigureCapacity cap
    cap.SolveByLoadMultiplier mMesh, mRebars, mConcreteMaterial, mSteelMaterial, mN(index), mMx(index), mMy(index)
    StoreCapacity index, cap
End Sub

Private Sub RunCrack(ByVal index As Long)
    If IsLinearMatrixMode Then
        mCrackStatus(index) = STATUS_NOT_CALCULATED
        Exit Sub
    End If

    If Not mCrackEnabled Then
        mCrackStatus(index) = STATUS_NOT_CALCULATED
        Exit Sub
    End If

    Dim service As CSectionSolver
    Set service = New CSectionSolver
    ConfigureSectionSolver service
    service.Solve mMesh, mRebars, mConcreteMaterial, mSteelMaterial, mN(index), mMx(index), mMy(index)
    If service.Converged Then
        Dim crack As CCrackWidthCalculator
        Set crack = New CCrackWidthCalculator
        crack.AllowableCrackWidth = mCrackAllowable
        crack.CrackSpacing = mCrackSpacing
        crack.StrainFactor = mCrackStrainFactor
        crack.DurationFactor = mCrackDurationFactor
        crack.Calculate service, mRebars, mSteelMaterial
        mCrackStatus(index) = IIf(crack.Converged, STATUS_OK, crack.StopReason)
        mCrackWidth(index) = crack.CrackWidth
        mCrackUtil(index) = crack.Utilization
    Else
        mCrackStatus(index) = STATUS_NUMERICAL_FAILURE & ": " & service.StopReason
        SetRowStatusIfOk index, mCrackStatus(index)
    End If
End Sub

Private Function IsFullCapacityMode() As Boolean
    IsFullCapacityMode = (StrComp(mCalculationMode, "FullCapacity", vbTextCompare) = 0)
End Function

Private Function IsLinearMatrixMode() As Boolean
    IsLinearMatrixMode = (StrComp(mCalculationMode, "LinearMatrix", vbTextCompare) = 0)
End Function

Private Sub StoreNotCalculatedCapacity(ByVal index As Long)
    mCapacityStatus(index) = STATUS_NOT_CALCULATED
End Sub

Private Sub StoreSkippedCapacity(ByVal index As Long)
    mCapacityStatus(index) = STATUS_SKIPPED
End Sub

Private Sub StoreCapacity(ByVal index As Long, ByVal cap As CCapacitySolver)
    Dim statusText As String
    statusText = cap.LimitState
    If Len(statusText) = 0 Then statusText = cap.StopReason
    If cap.Converged Then statusText = STATUS_OK & ": " & cap.LimitState

    mCapacityStatus(index) = statusText
    mLambdaCapacity(index) = cap.LambdaUltimate
    If cap.Converged Then mUtilCapacity(index) = cap.Utilization

    If Not cap.Converged Then SetRowStatusIfOk index, statusText
End Sub

Private Sub MarkInvalid(ByVal index As Long, ByVal reason As String)
    If Len(Trim$(reason)) = 0 Then reason = "Invalid load combination."
    mStatus(index) = STATUS_INVALID_INPUT & ": " & reason
    mCapacityStatus(index) = STATUS_INVALID_INPUT
    mCrackStatus(index) = STATUS_INVALID_INPUT
    mGoverningUtil(index) = 0#
    AppendDiagnostic index
End Sub

Private Sub SetRowStatusIfOk(ByVal index As Long, ByVal statusText As String)
    If mStatus(index) = STATUS_OK Or Len(mStatus(index)) = 0 Then mStatus(index) = statusText
End Sub

Private Sub ConfigureCapacity(ByVal cap As CCapacitySolver)
    cap.ConcreteCompressionLimit = mCapacityConcreteCompressionLimit
    cap.SteelStrainLimit = mCapacitySteelStrainLimit
    cap.LambdaTolerance = mCapacityToleranceLambda
    cap.MaxLambda = mCapacityMaxLambda
    cap.InitialLambdaStep = mCapacityInitialLambda
    cap.MaxRetries = mCapacityMaxRetries
    cap.SolverBaseLoadSteps = mCapacityBaseLoadSteps
    cap.SolverMaxIterations = mCapacitySolverMaxIterations
    cap.SolverToleranceN = mSolverToleranceN
    cap.SolverToleranceMx = mSolverToleranceMx
    cap.SolverToleranceMy = mSolverToleranceMy
    cap.SolverMaxDeltaEpsilon0 = mSolverMaxDeltaEpsilon0
    cap.SolverMaxDeltaKappa = mSolverMaxDeltaKappa
End Sub

Private Sub ConfigureSectionSolver(ByVal solver As CSectionSolver)
    solver.LoadSteps = mSolverLoadSteps
    solver.MaxIterations = mSolverMaxIterations
    solver.ToleranceN = mSolverToleranceN
    solver.ToleranceMx = mSolverToleranceMx
    solver.ToleranceMy = mSolverToleranceMy
    solver.MaxDeltaEpsilon0 = mSolverMaxDeltaEpsilon0
    solver.MaxDeltaKappa = mSolverMaxDeltaKappa
    solver.DiagnosticsEnabled = mSolverDiagnosticsEnabled
End Sub

Private Sub ValidateReady()
    If mMesh Is Nothing Then Err.Raise vbObjectError + 3920, "CBatchSectionCalculator", "Batch calculator is not initialized."
    If mCount <= 0 Then Err.Raise vbObjectError + 3921, "CBatchSectionCalculator", "No load combinations."
End Sub

Private Sub ClearResultRow(ByVal index As Long)
    mStatus(index) = vbNullString
    mCapacityStatus(index) = vbNullString
    mCrackStatus(index) = vbNullString
    mLambdaCapacity(index) = 0#
    mUtilCapacity(index) = 0#
    mCrackWidth(index) = 0#: mCrackUtil(index) = 0#: mGoverningUtil(index) = 0#
End Sub

Private Sub AppendDiagnostic(ByVal index As Long)
    mDiagnosticLog = mDiagnosticLog & "combination=" & mCombinationID(index) & _
        "; status=" & mStatus(index) & _
        "; CapacityStatus=" & mCapacityStatus(index) & _
        "; CrackStatus=" & mCrackStatus(index) & _
        "; governingUtil=" & FormatNumberInvariant(mGoverningUtil(index)) & vbCrLf
End Sub

Private Function MaxDouble(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then MaxDouble = a Else MaxDouble = b
End Function

Private Function FormatNumberInvariant(ByVal value As Double) As String
    FormatNumberInvariant = Replace$(Format$(value, "0.############"), ",", ".")
End Function

Public Property Get Count() As Long
    Count = mCount
End Property

Public Property Get CombinationID(ByVal index As Long) As String
    CombinationID = mCombinationID(index)
End Property

Public Property Get CombinationName(ByVal index As Long) As String
    CombinationName = mComment(index)
End Property

Public Property Get Status(ByVal index As Long) As String
    Status = mStatus(index)
End Property

Public Property Get CapacityStatus(ByVal index As Long) As String
    CapacityStatus = mCapacityStatus(index)
End Property

Public Property Get CrackStatus(ByVal index As Long) As String
    CrackStatus = mCrackStatus(index)
End Property

Public Property Get LambdaCapacity(ByVal index As Long) As Double
    LambdaCapacity = mLambdaCapacity(index)
End Property

Public Property Get UtilizationCapacity(ByVal index As Long) As Double
    UtilizationCapacity = mUtilCapacity(index)
End Property

Public Property Get CrackWidth(ByVal index As Long) As Double
    CrackWidth = mCrackWidth(index)
End Property

Public Property Get CrackUtilization(ByVal index As Long) As Double
    CrackUtilization = mCrackUtil(index)
End Property

Public Property Get GoverningUtilization(ByVal index As Long) As Double
    GoverningUtilization = mGoverningUtil(index)
End Property

Public Property Get GoverningCombinationIndex() As Long
    GoverningCombinationIndex = mGoverningCombinationIndex
End Property

Public Property Get GoverningCombinationID() As String
    If mGoverningCombinationIndex > 0 Then GoverningCombinationID = mCombinationID(mGoverningCombinationIndex)
End Property

Public Property Get GoverningCombinationName() As String
    If mGoverningCombinationIndex > 0 Then GoverningCombinationName = mComment(mGoverningCombinationIndex)
End Property

Public Property Get ElapsedSeconds() As Double
    ElapsedSeconds = mElapsedSeconds
End Property

Public Property Get DiagnosticLog() As String
    DiagnosticLog = mDiagnosticLog
End Property






