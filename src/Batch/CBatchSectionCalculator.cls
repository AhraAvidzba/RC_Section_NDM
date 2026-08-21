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
Private Const STATUS_STRENGTH_PASSED As String = "StrengthPassed"
Private Const STATUS_STRENGTH_FAILED As String = "StrengthFailed"

Private mSection As CSectionModel
Private mConcreteMaterial As Object
Private mSteelMaterial As Object
Private mCount As Long

Private mCombinationID(1 To MAX_COMBINATIONS) As String
Private mN(1 To MAX_COMBINATIONS) As Double
Private mMx(1 To MAX_COMBINATIONS) As Double
Private mMy(1 To MAX_COMBINATIONS) As Double
Private mUserMx(1 To MAX_COMBINATIONS) As Double
Private mUserMy(1 To MAX_COMBINATIONS) As Double
Private mCalculationType(1 To MAX_COMBINATIONS) As String
Private mDurationType(1 To MAX_COMBINATIONS) As String
Private mComment(1 To MAX_COMBINATIONS) As String
Private mInputValid(1 To MAX_COMBINATIONS) As Boolean
Private mInputError(1 To MAX_COMBINATIONS) As String
Private mLoadReferenceX As Double
Private mLoadReferenceY As Double

Private mStatus(1 To MAX_COMBINATIONS) As String
Private mCapacityStatus(1 To MAX_COMBINATIONS) As String
Private mCapacityLimitState(1 To MAX_COMBINATIONS) As String
Private mStrengthCheckStatus(1 To MAX_COMBINATIONS) As String
Private mCrackStatus(1 To MAX_COMBINATIONS) As String
Private mLambdaCapacity(1 To MAX_COMBINATIONS) As Double
Private mUtilCapacity(1 To MAX_COMBINATIONS) As Double
Private mCrackWidth(1 To MAX_COMBINATIONS) As Double
Private mCrackUtil(1 To MAX_COMBINATIONS) As Double
Private mStrainSafetyFactor(1 To MAX_COMBINATIONS) As Double
Private mGoverningUtil(1 To MAX_COMBINATIONS) As Double
Private mGoverningCombinationIndex As Long
Private mElapsedSeconds As Double
Private mDiagnosticLog As String
Private mStateAvailable(1 To MAX_COMBINATIONS) As Boolean
Private mStateConverged(1 To MAX_COMBINATIONS) As Boolean
Private mEpsilon0(1 To MAX_COMBINATIONS) As Double
Private mKappaX(1 To MAX_COMBINATIONS) As Double
Private mKappaY(1 To MAX_COMBINATIONS) As Double
Private mNint(1 To MAX_COMBINATIONS) As Double
Private mMxint(1 To MAX_COMBINATIONS) As Double
Private mMyint(1 To MAX_COMBINATIONS) As Double
Private mResidualN(1 To MAX_COMBINATIONS) As Double
Private mResidualMx(1 To MAX_COMBINATIONS) As Double
Private mResidualMy(1 To MAX_COMBINATIONS) As Double
Private mStateIterations(1 To MAX_COMBINATIONS) As Long

Private mCrackAllowable As Double
Private mCrackSpacing As Double
Private mCrackStrainFactor As Double
Private mCrackDurationFactor As Double
Private mCrackEnabled As Boolean
Private mCalculationMode As String
Private mSolverMethod As String
Private mSolverLoadSteps As Long
Private mSolverMaxIterations As Long
Private mSolverToleranceN As Double
Private mSolverToleranceMx As Double
Private mSolverToleranceMy As Double
Private mSolverMaxDeltaEpsilon0 As Double
Private mSolverMaxDeltaKappa As Double
Private mSolverSecantMaxRestarts As Long
Private mSolverSecantMinStepNorm As Double
Private mCapacityInitialLambda As Double
Private mCapacityMethod As String
Private mCapacitySearchMethod As String
Private mCapacityMaxLambda As Double
Private mCapacityToleranceLambda As Double
Private mCapacityToleranceStrain As Double
Private mCapacityMaxRetries As Long
Private mCapacityBaseLoadSteps As Long
Private mCapacitySolverMaxIterations As Long
Private mCapacityConcreteCompressionLimit As Double
Private mCapacityConcreteTensionLimit As Double
Private mCapacitySteelStrainLimit As Double
Private mConcreteTensionMode As String

Private Sub Class_Initialize()
    mCrackAllowable = 0.3
    mCrackSpacing = 200#
    mCrackStrainFactor = 1#
    mCrackDurationFactor = 1#
    mCalculationMode = "FullCapacity"
    mSolverMethod = "Newton"
    mCrackEnabled = True
    mSolverLoadSteps = 8
    mSolverMaxIterations = 80
    mSolverToleranceN = 5#
    mSolverToleranceMx = 5000#
    mSolverToleranceMy = 5000#
    mSolverMaxDeltaEpsilon0 = 0.0005
    mSolverMaxDeltaKappa = 0.00001
    mSolverSecantMaxRestarts = 2
    mSolverSecantMinStepNorm = 0.000000000001
    mCapacityInitialLambda = 1#
    mCapacityMethod = "UltimateStrain"
    mCapacitySearchMethod = "Bisection"
    mCapacityMaxLambda = 64#
    mCapacityToleranceLambda = 0.01
    mCapacityToleranceStrain = 0.00001
    mCapacityMaxRetries = 0
    mCapacityBaseLoadSteps = 1
    mCapacitySolverMaxIterations = 60
    mCapacityConcreteCompressionLimit = -0.0035
    mCapacityConcreteTensionLimit = 0.00015
    mCapacitySteelStrainLimit = 0.025
    mConcreteTensionMode = "Ignore"
End Sub

Public Sub Initialize(ByVal section As CSectionModel, ByVal concreteMaterial As Object, ByVal steelMaterial As Object)
    If section Is Nothing Then Err.Raise vbObjectError + 3900, "CBatchSectionCalculator", "Модель сечения не передана."
    If section.ConcreteCount <= 0 Then Err.Raise vbObjectError + 3901, "CBatchSectionCalculator", "В модели сечения нет бетонных элементов."
    If concreteMaterial Is Nothing Then Err.Raise vbObjectError + 3902, "CBatchSectionCalculator", "Материал бетона не передан."
    If steelMaterial Is Nothing Then Err.Raise vbObjectError + 3903, "CBatchSectionCalculator", "Материал арматуры не передан."
    Set mSection = section
    Set mConcreteMaterial = concreteMaterial
    Set mSteelMaterial = steelMaterial
End Sub

Public Sub ApplySettings(ByVal settings As CSystemSettingsReader, Optional ByVal units As CUnitSystem = Nothing)
    If settings Is Nothing Then Exit Sub
    mCrackAllowable = settings.GetDouble("CrackWidth.Allowable", mCrackAllowable)
    mCrackSpacing = settings.GetDouble("CrackWidth.CrackSpacing", mCrackSpacing)
    mCrackStrainFactor = settings.GetDouble("CrackWidth.StrainFactor", mCrackStrainFactor)
    mCrackDurationFactor = settings.GetDouble("CrackWidth.DurationFactor", mCrackDurationFactor)
    mCalculationMode = settings.GetString("Calculation.Mode", mCalculationMode)
    If settings.HasKey("Solver.Method") Then
        mSolverMethod = settings.GetRawString("Solver.Method", vbNullString)
    Else
        mSolverMethod = "Newton"
    End If
    mCrackEnabled = settings.GetBoolean("CrackWidth.Enabled", mCrackEnabled)
    mSolverLoadSteps = settings.GetLong("Solver.LoadSteps", mSolverLoadSteps)
    mSolverMaxIterations = settings.GetLong("Solver.MaxIterations", mSolverMaxIterations)
    mSolverToleranceN = settings.GetDouble("Solver.ToleranceN", mSolverToleranceN)
    mSolverToleranceMx = settings.GetDouble("Solver.ToleranceMx", mSolverToleranceMx)
    mSolverToleranceMy = settings.GetDouble("Solver.ToleranceMy", mSolverToleranceMy)
    mSolverMaxDeltaEpsilon0 = settings.GetDouble("Solver.MaxDeltaEpsilon0", mSolverMaxDeltaEpsilon0)
    mSolverMaxDeltaKappa = settings.GetDouble("Solver.MaxDeltaKappa", mSolverMaxDeltaKappa)
    mSolverSecantMaxRestarts = settings.GetLong("Solver.SecantMaxRestarts", mSolverSecantMaxRestarts)
    mSolverSecantMinStepNorm = settings.GetDouble("Solver.SecantMinStepNorm", mSolverSecantMinStepNorm)
    mCapacityInitialLambda = settings.GetDouble("Capacity.InitialLambda", mCapacityInitialLambda)
    mCapacityMethod = settings.GetString("Capacity.Method", mCapacityMethod)
    If settings.HasKey("Capacity.SearchMethod") Then
        mCapacitySearchMethod = settings.GetRawString("Capacity.SearchMethod", vbNullString)
    Else
        mCapacitySearchMethod = "Bisection"
    End If
    mCapacityMaxLambda = settings.GetDouble("Capacity.MaxLambda", mCapacityMaxLambda)
    mCapacityToleranceLambda = settings.GetDouble("Capacity.ToleranceLambda", mCapacityToleranceLambda)
    mCapacityToleranceStrain = settings.GetDouble("Capacity.ToleranceStrain", mCapacityToleranceStrain)
    mCapacityMaxRetries = settings.GetLong("Capacity.MaxRetries", mCapacityMaxRetries)
    mCapacityBaseLoadSteps = settings.GetLong("Capacity.BaseLoadSteps", mCapacityBaseLoadSteps)
    mCapacitySolverMaxIterations = settings.GetLong("Capacity.SolverMaxIterations", mCapacitySolverMaxIterations)
    mCapacityConcreteCompressionLimit = settings.GetDouble("Capacity.ConcreteCompressionLimit", mCapacityConcreteCompressionLimit)
    mCapacityConcreteTensionLimit = settings.GetDouble("Capacity.ConcreteTensionLimit", mCapacityConcreteTensionLimit)
    mCapacitySteelStrainLimit = settings.GetDouble("Capacity.SteelStrainLimit", mCapacitySteelStrainLimit)
    mConcreteTensionMode = settings.GetString("Concrete.TensionMode", mConcreteTensionMode)
    If Not units Is Nothing Then
        mCrackAllowable = units.InputLengthToInternal(mCrackAllowable)
        mCrackSpacing = units.InputLengthToInternal(mCrackSpacing)
        mSolverToleranceN = Abs(units.InputForceToInternal(mSolverToleranceN))
        mSolverToleranceMx = Abs(units.InputMomentMxToInternal(mSolverToleranceMx))
        mSolverToleranceMy = Abs(units.InputMomentMyToInternal(mSolverToleranceMy))
        mSolverMaxDeltaKappa = Abs(units.InputCurvatureToInternal(mSolverMaxDeltaKappa))
    End If
End Sub

Public Sub ClearCombinations()
    mCount = 0
    mGoverningCombinationIndex = 0
    mElapsedSeconds = 0#
    mDiagnosticLog = vbNullString
    mLoadReferenceX = 0#
    mLoadReferenceY = 0#
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
    If mCount >= MAX_COMBINATIONS Then Err.Raise vbObjectError + 3911, "CBatchSectionCalculator", "РџРѕРґРґРµСЂР¶РёРІР°РµС‚СЃСЏ РЅРµ Р±РѕР»РµРµ 20 СЃРѕС‡РµС‚Р°РЅРёР№ РЅР°РіСЂСѓР·РѕРє."
    mCount = mCount + 1
    If Len(Trim$(combinationID)) = 0 Then combinationID = "ROW" & CStr(mCount)
    mCombinationID(mCount) = combinationID
    mN(mCount) = nValue
    mMx(mCount) = mxValue
    mMy(mCount) = myValue
    mCalculationType(mCount) = calculationType
    mDurationType(mCount) = durationType
    mComment(mCount) = comment
    mUserMx(mCount) = mxValue
    mUserMy(mCount) = myValue
    mInputValid(mCount) = inputValid
    mInputError(mCount) = inputError
End Sub

Public Sub ApplyLoadReference(ByVal referenceX As Double, ByVal referenceY As Double)
    mLoadReferenceX = referenceX
    mLoadReferenceY = referenceY

    Dim i As Long
    For i = 1 To mCount
        If mInputValid(i) Then
            mMx(i) = mUserMx(i) + mN(i) * referenceY
            mMy(i) = mUserMy(i) + mN(i) * referenceX
        End If
    Next i
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
            ElseIf IsBetterGoverningCandidate(i, mGoverningCombinationIndex) Then
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
    If Not IsSupportedCalculationMode() Then
        MarkInvalid index, "Calculation.Mode РґРѕР»Р¶РµРЅ Р±С‹С‚СЊ DirectState РёР»Рё FullCapacity."
        Exit Sub
    End If

    mStatus(index) = STATUS_OK
    If IsFullCapacityMode Then
        RunCapacity index
    Else
        StoreNotCalculatedCapacity index
    End If
    RunCrack index

    mGoverningUtil(index) = MaxDouble(MaxDouble(mUtilCapacity(index), StrainUtilization(index)), mCrackUtil(index))
    mStatus(index) = AggregateStatus(index)
    AppendDiagnostic index
    Exit Sub

Failed:
    MarkInvalid index, Err.Description
End Sub

Private Sub RunCapacity(ByVal index As Long)
    If Sqr(mUserMx(index) * mUserMx(index) + mUserMy(index) * mUserMy(index)) <= 0.000000001 Then
        StoreSkippedCapacity index
        Exit Sub
    End If

    Dim cap As CCapacitySolver
    Set cap = New CCapacitySolver
    ConfigureCapacity cap
    Select Case NormalizedCapacityMethod()
        Case "UltimateStrain"
            cap.SolveByUltimateStrain mSection, mConcreteMaterial, mSteelMaterial, _
                mN(index), mUserMx(index), mUserMy(index), mN(index) * mLoadReferenceY, mN(index) * mLoadReferenceX
        Case "LoadMultiplier"
            cap.SolveByLoadMultiplier mSection, mConcreteMaterial, mSteelMaterial, _
                mN(index), mUserMx(index), mUserMy(index), mN(index) * mLoadReferenceY, mN(index) * mLoadReferenceX
        Case Else
            mCapacityStatus(index) = STATUS_INVALID_INPUT & ": Capacity.Method РґРѕР»Р¶РµРЅ Р±С‹С‚СЊ LoadMultiplier РёР»Рё UltimateStrain."
            mCapacityLimitState(index) = STATUS_INVALID_INPUT
            mStrengthCheckStatus(index) = STATUS_INVALID_INPUT
            Exit Sub
    End Select
    StoreCapacity index, cap
End Sub

Private Sub RunCrack(ByVal index As Long)
    Dim service As CSectionSolver
    Set service = New CSectionSolver
    ConfigureSectionSolver service
    service.Solve mSection, mConcreteMaterial, mSteelMaterial, mN(index), mMx(index), mMy(index)
    StoreSectionState index, service
    mStrainSafetyFactor(index) = CalculateStrainSafetyFactor(service)
    If service.Converged Then
        If Not mCrackEnabled Then
            mCrackStatus(index) = STATUS_NOT_CALCULATED
            Exit Sub
        End If
        Dim crack As CCrackWidthCalculator
        Set crack = New CCrackWidthCalculator
        crack.AllowableCrackWidth = mCrackAllowable
        crack.CrackSpacing = mCrackSpacing
        crack.StrainFactor = mCrackStrainFactor
        crack.DurationFactor = mCrackDurationFactor
        crack.Calculate service, mSection, mSteelMaterial
        mCrackStatus(index) = IIf(crack.Converged, STATUS_OK, crack.StopReason)
        mCrackWidth(index) = crack.CrackWidth
        mCrackUtil(index) = crack.Utilization
    Else
        mCrackStatus(index) = STATUS_NUMERICAL_FAILURE & ": " & service.StopReason
    End If
End Sub

Private Function IsFullCapacityMode() As Boolean
    IsFullCapacityMode = (StrComp(mCalculationMode, "FullCapacity", vbTextCompare) = 0)
End Function

Private Function IsDirectStateMode() As Boolean
    IsDirectStateMode = (StrComp(mCalculationMode, "DirectState", vbTextCompare) = 0)
End Function

Private Function IsSupportedCalculationMode() As Boolean
    IsSupportedCalculationMode = IsDirectStateMode() Or IsFullCapacityMode()
End Function

Private Function NormalizedCapacityMethod() As String
    Select Case LCase$(Trim$(mCapacityMethod))
        Case "loadmultiplier"
            NormalizedCapacityMethod = "LoadMultiplier"
        Case "ultimatestrain"
            NormalizedCapacityMethod = "UltimateStrain"
    End Select
End Function

Private Sub StoreNotCalculatedCapacity(ByVal index As Long)
    mCapacityStatus(index) = STATUS_NOT_CALCULATED
    mCapacityLimitState(index) = STATUS_NOT_CALCULATED
    mStrengthCheckStatus(index) = STATUS_NOT_CALCULATED
End Sub

Private Sub StoreSkippedCapacity(ByVal index As Long)
    mCapacityStatus(index) = STATUS_SKIPPED
    mCapacityLimitState(index) = STATUS_SKIPPED
    mStrengthCheckStatus(index) = STATUS_NOT_CALCULATED
End Sub

Private Sub StoreCapacity(ByVal index As Long, ByVal cap As CCapacitySolver)
    Dim statusText As String
    If cap.Converged Then
        statusText = STATUS_OK
    Else
        statusText = cap.LimitState
        If Len(cap.StopReason) > 0 Then statusText = statusText & ": " & cap.StopReason
    End If

    mCapacityStatus(index) = statusText
    mCapacityLimitState(index) = cap.LimitState
    mLambdaCapacity(index) = cap.LambdaUltimate
    If cap.LambdaUltimate > 0# Then mUtilCapacity(index) = 1# / cap.LambdaUltimate
    If cap.Converged Then mUtilCapacity(index) = cap.Utilization
    mStrengthCheckStatus(index) = StrengthStatusFromLambda(cap)
End Sub

Private Sub MarkInvalid(ByVal index As Long, ByVal reason As String)
    If Len(Trim$(reason)) = 0 Then reason = "РќРµРєРѕСЂСЂРµРєС‚РЅРѕРµ СЃРѕС‡РµС‚Р°РЅРёРµ РЅР°РіСЂСѓР·РѕРє."
    mStatus(index) = STATUS_INVALID_INPUT & ": " & reason
    mCapacityStatus(index) = STATUS_INVALID_INPUT
    mCapacityLimitState(index) = STATUS_INVALID_INPUT
    mStrengthCheckStatus(index) = STATUS_INVALID_INPUT
    mCrackStatus(index) = STATUS_INVALID_INPUT
    mGoverningUtil(index) = 0#
    AppendDiagnostic index
End Sub

Private Sub ConfigureCapacity(ByVal cap As CCapacitySolver)
    cap.ConcreteCompressionLimit = mCapacityConcreteCompressionLimit
    cap.ConcreteTensionLimit = mCapacityConcreteTensionLimit
    cap.SteelStrainLimit = mCapacitySteelStrainLimit
    cap.LambdaTolerance = mCapacityToleranceLambda
    cap.StrainTolerance = mCapacityToleranceStrain
    cap.SearchMethod = mCapacitySearchMethod
    cap.MaxLambda = mCapacityMaxLambda
    cap.InitialLambdaStep = mCapacityInitialLambda
    cap.MaxRetries = mCapacityMaxRetries
    cap.SolverBaseLoadSteps = mCapacityBaseLoadSteps
    cap.SolverMaxIterations = mCapacitySolverMaxIterations
    cap.SolverToleranceN = mSolverToleranceN
    cap.SolverToleranceMx = mSolverToleranceMx
    cap.SolverToleranceMy = mSolverToleranceMy
    cap.SolverMethod = mSolverMethod
    cap.SolverMaxDeltaEpsilon0 = mSolverMaxDeltaEpsilon0
    cap.SolverMaxDeltaKappa = mSolverMaxDeltaKappa
    cap.SolverSecantMaxRestarts = mSolverSecantMaxRestarts
    cap.SolverSecantMinStepNorm = mSolverSecantMinStepNorm
End Sub

Private Sub ConfigureSectionSolver(ByVal solver As CSectionSolver)
    solver.LoadSteps = mSolverLoadSteps
    solver.SolverMethod = mSolverMethod
    solver.MaxIterations = mSolverMaxIterations
    solver.ToleranceN = mSolverToleranceN
    solver.ToleranceMx = mSolverToleranceMx
    solver.ToleranceMy = mSolverToleranceMy
    solver.MaxDeltaEpsilon0 = mSolverMaxDeltaEpsilon0
    solver.MaxDeltaKappa = mSolverMaxDeltaKappa
    solver.SecantMaxRestarts = mSolverSecantMaxRestarts
    solver.SecantMinStepNorm = mSolverSecantMinStepNorm
End Sub

Private Sub ValidateReady()
    If mSection Is Nothing Then Err.Raise vbObjectError + 3920, "CBatchSectionCalculator", "Пакетный расчетчик не инициализирован моделью сечения."
    If mCount <= 0 Then Err.Raise vbObjectError + 3921, "CBatchSectionCalculator", "РќРµ Р·Р°РґР°РЅРѕ РЅРё РѕРґРЅРѕРіРѕ СЃРѕС‡РµС‚Р°РЅРёСЏ РЅР°РіСЂСѓР·РѕРє."
End Sub

Private Sub ClearResultRow(ByVal index As Long)
    mStatus(index) = vbNullString
    mCapacityStatus(index) = vbNullString
    mCapacityLimitState(index) = vbNullString
    mStrengthCheckStatus(index) = vbNullString
    mCrackStatus(index) = vbNullString
    mLambdaCapacity(index) = 0#
    mUtilCapacity(index) = 0#
    mCrackWidth(index) = 0#: mCrackUtil(index) = 0#: mStrainSafetyFactor(index) = 0#: mGoverningUtil(index) = 0#
    mStateAvailable(index) = False
    mStateConverged(index) = False
    mEpsilon0(index) = 0#: mKappaX(index) = 0#: mKappaY(index) = 0#
    mNint(index) = 0#: mMxint(index) = 0#: mMyint(index) = 0#
    mResidualN(index) = 0#: mResidualMx(index) = 0#: mResidualMy(index) = 0#
    mStateIterations(index) = 0
End Sub

Private Sub StoreSectionState(ByVal index As Long, ByVal solver As CSectionSolver)
    mStateAvailable(index) = True
    mStateConverged(index) = solver.Converged
    mEpsilon0(index) = solver.Epsilon0
    mKappaX(index) = solver.KappaX
    mKappaY(index) = solver.KappaY
    mNint(index) = solver.Nint
    mMxint(index) = solver.Mxint
    mMyint(index) = solver.Myint
    mResidualN(index) = solver.ResidualN
    mResidualMx(index) = solver.ResidualMx
    mResidualMy(index) = solver.ResidualMy
    mStateIterations(index) = solver.Iterations
End Sub

Private Sub AppendDiagnostic(ByVal index As Long)
    mDiagnosticLog = mDiagnosticLog & "combination=" & mCombinationID(index) & _
        "; status=" & mStatus(index) & _
        "; CapacityStatus=" & mCapacityStatus(index) & _
        "; StrengthCheckStatus=" & mStrengthCheckStatus(index) & _
        "; CrackStatus=" & mCrackStatus(index) & _
        "; governingUtil=" & FormatNumberInvariant(mGoverningUtil(index)) & vbCrLf
End Sub

Private Function AggregateStatus(ByVal index As Long) As String
    If mStrengthCheckStatus(index) = STATUS_STRENGTH_FAILED Then
        AggregateStatus = STATUS_STRENGTH_FAILED
    ElseIf mStrengthCheckStatus(index) = STATUS_STRENGTH_PASSED Then
        If IsNumericalStatus(mCrackStatus(index)) Then
            AggregateStatus = "CrackNumericalFailure"
        Else
            AggregateStatus = STATUS_OK
        End If
    ElseIf IsInvalidStatus(mCapacityStatus(index)) Or IsInvalidStatus(mStrengthCheckStatus(index)) Then
        AggregateStatus = mCapacityStatus(index)
    ElseIf StrainUtilization(index) > 1# Then
        AggregateStatus = "StrainLimitExceeded"
    ElseIf IsNumericalStatus(mCapacityStatus(index)) Then
        AggregateStatus = mCapacityStatus(index)
    ElseIf IsNumericalStatus(mCrackStatus(index)) Then
        AggregateStatus = mCrackStatus(index)
    Else
        AggregateStatus = STATUS_OK
    End If
End Function

Private Function IsNumericalStatus(ByVal statusText As String) As Boolean
    IsNumericalStatus = (InStr(1, statusText, STATUS_NUMERICAL_FAILURE, vbTextCompare) > 0 Or _
        InStr(1, statusText, "SingularTangent", vbTextCompare) > 0)
End Function

Private Function IsInvalidStatus(ByVal statusText As String) As Boolean
    IsInvalidStatus = (InStr(1, statusText, STATUS_INVALID_INPUT, vbTextCompare) > 0)
End Function

Private Function StrengthStatusFromLambda(ByVal cap As CCapacitySolver) As String
    If Not cap.Converged Then
        StrengthStatusFromLambda = STATUS_NOT_CALCULATED
    ElseIf cap.LambdaUltimate <= 0# Then
        StrengthStatusFromLambda = STATUS_NOT_CALCULATED
    ElseIf cap.LambdaUltimate < 1# Then
        StrengthStatusFromLambda = STATUS_STRENGTH_FAILED
    Else
        StrengthStatusFromLambda = STATUS_STRENGTH_PASSED
    End If
End Function

Private Function MaxDouble(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then MaxDouble = a Else MaxDouble = b
End Function

Private Function CalculateStrainSafetyFactor(ByVal solver As CSectionSolver) As Double
    Dim maxUtilization As Double
    maxUtilization = 0#

    If mCapacityConcreteCompressionLimit < 0# And solver.MinConcreteStrain < 0# Then
        maxUtilization = MaxDouble(maxUtilization, Abs(solver.MinConcreteStrain) / Abs(mCapacityConcreteCompressionLimit))
    End If

    If StrComp(mConcreteTensionMode, "UseDiagram", vbTextCompare) = 0 Then
        If mCapacityConcreteTensionLimit > 0# And solver.MaxConcreteStrain > 0# Then
            maxUtilization = MaxDouble(maxUtilization, solver.MaxConcreteStrain / mCapacityConcreteTensionLimit)
        End If
    End If

    If mCapacitySteelStrainLimit > 0# And solver.MaxAbsSteelStrain > 0# Then
        maxUtilization = MaxDouble(maxUtilization, solver.MaxAbsSteelStrain / mCapacitySteelStrainLimit)
    End If

    If maxUtilization > 0# Then CalculateStrainSafetyFactor = 1# / maxUtilization
End Function

Private Function StrainUtilization(ByVal index As Long) As Double
    If mStrainSafetyFactor(index) > 0# Then StrainUtilization = 1# / mStrainSafetyFactor(index)
End Function

Private Function IsBetterGoverningCandidate(ByVal candidateIndex As Long, ByVal currentIndex As Long) As Boolean
    Dim candidateLambda As Double
    Dim currentLambda As Double
    candidateLambda = mLambdaCapacity(candidateIndex)
    currentLambda = mLambdaCapacity(currentIndex)

    If candidateLambda > 0# And currentLambda > 0# Then
        IsBetterGoverningCandidate = (candidateLambda < currentLambda)
    ElseIf candidateLambda > 0# Then
        IsBetterGoverningCandidate = True
    ElseIf currentLambda > 0# Then
        IsBetterGoverningCandidate = False
    Else
        IsBetterGoverningCandidate = (mGoverningUtil(candidateIndex) > mGoverningUtil(currentIndex))
    End If
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

Public Property Get N(ByVal index As Long) As Double
    N = mN(index)
End Property

Public Property Get Mx(ByVal index As Long) As Double
    Mx = mMx(index)
End Property

Public Property Get My(ByVal index As Long) As Double
    My = mMy(index)
End Property

Public Property Get UserMx(ByVal index As Long) As Double
    UserMx = mUserMx(index)
End Property

Public Property Get UserMy(ByVal index As Long) As Double
    UserMy = mUserMy(index)
End Property

Public Property Get LoadReferenceX() As Double
    LoadReferenceX = mLoadReferenceX
End Property

Public Property Get LoadReferenceY() As Double
    LoadReferenceY = mLoadReferenceY
End Property

Public Property Get Status(ByVal index As Long) As String
    Status = mStatus(index)
End Property

Public Property Get CapacityStatus(ByVal index As Long) As String
    CapacityStatus = mCapacityStatus(index)
End Property

Public Property Get CapacityLimitState(ByVal index As Long) As String
    CapacityLimitState = mCapacityLimitState(index)
End Property

Public Property Get StrengthCheckStatus(ByVal index As Long) As String
    StrengthCheckStatus = mStrengthCheckStatus(index)
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

Public Property Get StrainSafetyFactor(ByVal index As Long) As Double
    StrainSafetyFactor = mStrainSafetyFactor(index)
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

Public Property Get InvalidInputCount() As Long
    Dim i As Long
    For i = 1 To mCount
        If IsInvalidStatus(mStatus(i)) Then InvalidInputCount = InvalidInputCount + 1
    Next i
End Property

Public Property Get FirstInvalidInputMessage() As String
    Dim i As Long
    For i = 1 To mCount
        If IsInvalidStatus(mStatus(i)) Then
            FirstInvalidInputMessage = mStatus(i)
            Exit Property
        End If
    Next i
End Property

Public Property Get StateAvailable(ByVal index As Long) As Boolean
    StateAvailable = mStateAvailable(index)
End Property

Public Property Get StateConverged(ByVal index As Long) As Boolean
    StateConverged = mStateConverged(index)
End Property

Public Property Get Epsilon0(ByVal index As Long) As Double
    Epsilon0 = mEpsilon0(index)
End Property

Public Property Get KappaX(ByVal index As Long) As Double
    KappaX = mKappaX(index)
End Property

Public Property Get KappaY(ByVal index As Long) As Double
    KappaY = mKappaY(index)
End Property

Public Property Get Nint(ByVal index As Long) As Double
    Nint = mNint(index)
End Property

Public Property Get Mxint(ByVal index As Long) As Double
    Mxint = mMxint(index)
End Property

Public Property Get Myint(ByVal index As Long) As Double
    Myint = mMyint(index)
End Property

Public Property Get ResidualN(ByVal index As Long) As Double
    ResidualN = mResidualN(index)
End Property

Public Property Get ResidualMx(ByVal index As Long) As Double
    ResidualMx = mResidualMx(index)
End Property

Public Property Get ResidualMy(ByVal index As Long) As Double
    ResidualMy = mResidualMy(index)
End Property

Public Property Get StateIterations(ByVal index As Long) As Long
    StateIterations = mStateIterations(index)
End Property

