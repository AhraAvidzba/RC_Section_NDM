VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCrackWidthCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mConverged As Boolean
Private mStopReason As String
Private mTensionRebarCount As Long
Private mMaxSteelStrain As Double
Private mMaxSteelStress As Double
Private mCrackWidth As Double
Private mAllowableCrackWidth As Double
Private mUtilization As Double
Private mCrackSpacing As Double
Private mStrainFactor As Double
Private mDurationFactor As Double
Private mDiagnosticLog As String

Private Sub Class_Initialize()
    mAllowableCrackWidth = 0.3
    mCrackSpacing = 200#
    mStrainFactor = 1#
    mDurationFactor = 1#
End Sub

Public Sub ApplySettings(ByVal settings As CSystemSettingsReader)
    If settings Is Nothing Then Exit Sub
    mAllowableCrackWidth = settings.GetDouble("CrackWidth.Allowable", mAllowableCrackWidth)
    mCrackSpacing = settings.GetDouble("CrackWidth.CrackSpacing", mCrackSpacing)
    mStrainFactor = settings.GetDouble("CrackWidth.StrainFactor", mStrainFactor)
    mDurationFactor = settings.GetDouble("CrackWidth.DurationFactor", mDurationFactor)
End Sub

Public Sub Calculate(ByVal solver As CSectionSolver, ByVal rebars As CRebarLayout, ByVal steelMaterial As Object)
    ClearResult
    ValidateInputs solver, rebars, steelMaterial

    Dim i As Long
    Dim strain As Double
    Dim stress As Double
    For i = 1 To rebars.Count
        strain = solver.Epsilon0 + solver.KappaX * rebars.Y(i) + solver.KappaY * rebars.X(i)
        stress = steelMaterial.GetStress(strain)
        AppendBarDiagnostic rebars.BarID(i), strain, stress
        If strain > 0# Then
            mTensionRebarCount = mTensionRebarCount + 1
            If strain > mMaxSteelStrain Then
                mMaxSteelStrain = strain
                mMaxSteelStress = stress
            End If
        End If
    Next i

    If mTensionRebarCount = 0 Then
        mConverged = True
        mStopReason = "No tensile reinforcement in the supplied cracked-section state."
        Exit Sub
    End If

    mCrackWidth = mMaxSteelStrain * mCrackSpacing * mStrainFactor * mDurationFactor
    If mAllowableCrackWidth > 0# Then mUtilization = mCrackWidth / mAllowableCrackWidth
    mConverged = True
    mStopReason = "Crack width calculated from supplied SLS section state."
End Sub

Private Sub ValidateInputs(ByVal solver As CSectionSolver, ByVal rebars As CRebarLayout, ByVal steelMaterial As Object)
    If solver Is Nothing Then Err.Raise vbObjectError + 3700, "CCrackWidthCalculator", "Section solver result is missing."
    If Not solver.Converged Then Err.Raise vbObjectError + 3701, "CCrackWidthCalculator", "Section solver result is not converged."
    If rebars Is Nothing Then Err.Raise vbObjectError + 3702, "CCrackWidthCalculator", "Rebar layout is missing."
    If rebars.Count <= 0 Then Err.Raise vbObjectError + 3703, "CCrackWidthCalculator", "Rebar layout is empty."
    If steelMaterial Is Nothing Then Err.Raise vbObjectError + 3704, "CCrackWidthCalculator", "Steel material is missing."
    If mAllowableCrackWidth <= 0# Then Err.Raise vbObjectError + 3705, "CCrackWidthCalculator", "Allowable crack width must be positive."
    If mCrackSpacing < 0# Or mStrainFactor < 0# Or mDurationFactor < 0# Then Err.Raise vbObjectError + 3706, "CCrackWidthCalculator", "Crack parameters must be non-negative."
End Sub

Private Sub AppendBarDiagnostic(ByVal barName As String, ByVal strain As Double, ByVal stress As Double)
    mDiagnosticLog = mDiagnosticLog & "bar=" & barName & _
        "; strain=" & FormatNumberInvariant(strain) & _
        "; stress=" & FormatNumberInvariant(stress) & vbCrLf
End Sub

Private Sub ClearResult()
    mConverged = False
    mStopReason = vbNullString
    mTensionRebarCount = 0
    mMaxSteelStrain = 0#
    mMaxSteelStress = 0#
    mCrackWidth = 0#
    mUtilization = 0#
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

Public Property Get TensionRebarCount() As Long
    TensionRebarCount = mTensionRebarCount
End Property

Public Property Get MaxSteelStrain() As Double
    MaxSteelStrain = mMaxSteelStrain
End Property

Public Property Get MaxSteelStress() As Double
    MaxSteelStress = mMaxSteelStress
End Property

Public Property Get CrackWidth() As Double
    CrackWidth = mCrackWidth
End Property

Public Property Get AllowableCrackWidth() As Double
    AllowableCrackWidth = mAllowableCrackWidth
End Property

Public Property Get Utilization() As Double
    Utilization = mUtilization
End Property

Public Property Get DiagnosticLog() As String
    DiagnosticLog = mDiagnosticLog
End Property

Public Property Let AllowableCrackWidth(ByVal value As Double)
    mAllowableCrackWidth = value
End Property

Public Property Let CrackSpacing(ByVal value As Double)
    mCrackSpacing = value
End Property

Public Property Let StrainFactor(ByVal value As Double)
    mStrainFactor = value
End Property

Public Property Let DurationFactor(ByVal value As Double)
    mDurationFactor = value
End Property


