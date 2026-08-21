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

Public Sub ApplySettings(ByVal settings As CSystemSettingsReader, Optional ByVal units As CUnitSystem = Nothing)
    If settings Is Nothing Then Exit Sub
    mAllowableCrackWidth = settings.GetDouble("CrackWidth.Allowable", mAllowableCrackWidth)
    mCrackSpacing = settings.GetDouble("CrackWidth.CrackSpacing", mCrackSpacing)
    mStrainFactor = settings.GetDouble("CrackWidth.StrainFactor", mStrainFactor)
    mDurationFactor = settings.GetDouble("CrackWidth.DurationFactor", mDurationFactor)
    If Not units Is Nothing Then
        mAllowableCrackWidth = units.InputLengthToInternal(mAllowableCrackWidth)
        mCrackSpacing = units.InputLengthToInternal(mCrackSpacing)
    End If
End Sub

Public Sub Calculate(ByVal solver As CSectionSolver, ByVal section As CSectionModel, ByVal steelMaterial As Object)
    ClearResult
    ValidateInputs solver, section, steelMaterial

    Dim i As Long
    Dim strain As Double
    Dim stress As Double
    For i = 1 To section.RebarCount
        strain = solver.Epsilon0 + solver.KappaX * section.RebarY(i) + solver.KappaY * section.RebarX(i)
        stress = steelMaterial.GetStress(strain)
        AppendBarDiagnostic section.RebarID(i), strain, stress
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
        mStopReason = "Р’ РїРµСЂРµРґР°РЅРЅРѕРј СЃРѕСЃС‚РѕСЏРЅРёРё СЃ С‚СЂРµС‰РёРЅРѕР№ РЅРµС‚ СЂР°СЃС‚СЏРЅСѓС‚РѕР№ Р°СЂРјР°С‚СѓСЂС‹."
        Exit Sub
    End If

    mCrackWidth = mMaxSteelStrain * mCrackSpacing * mStrainFactor * mDurationFactor
    If mAllowableCrackWidth > 0# Then mUtilization = mCrackWidth / mAllowableCrackWidth
    mConverged = True
    mStopReason = "РЁРёСЂРёРЅР° СЂР°СЃРєСЂС‹С‚РёСЏ С‚СЂРµС‰РёРЅ СЂР°СЃСЃС‡РёС‚Р°РЅР° РїРѕ РїРµСЂРµРґР°РЅРЅРѕРјСѓ СЃРѕСЃС‚РѕСЏРЅРёСЋ СЃРµС‡РµРЅРёСЏ."
End Sub

Private Sub ValidateInputs(ByVal solver As CSectionSolver, ByVal section As CSectionModel, ByVal steelMaterial As Object)
    If solver Is Nothing Then Err.Raise vbObjectError + 3700, "CCrackWidthCalculator", "Р РµР·СѓР»СЊС‚Р°С‚ СЂРµС€Р°С‚РµР»СЏ СЃРµС‡РµРЅРёСЏ РЅРµ РїРµСЂРµРґР°РЅ."
    If Not solver.Converged Then Err.Raise vbObjectError + 3701, "CCrackWidthCalculator", "Р РµР·СѓР»СЊС‚Р°С‚ СЂРµС€Р°С‚РµР»СЏ СЃРµС‡РµРЅРёСЏ РЅРµ СЃРѕС€РµР»СЃСЏ."
    If section Is Nothing Then Err.Raise vbObjectError + 3702, "CCrackWidthCalculator", "Section model is missing."
    If section.RebarCount <= 0 Then Err.Raise vbObjectError + 3703, "CCrackWidthCalculator", "Section model has no rebar elements."
    If steelMaterial Is Nothing Then Err.Raise vbObjectError + 3704, "CCrackWidthCalculator", "РњР°С‚РµСЂРёР°Р» Р°СЂРјР°С‚СѓСЂС‹ РЅРµ РїРµСЂРµРґР°РЅ."
    If mAllowableCrackWidth <= 0# Then Err.Raise vbObjectError + 3705, "CCrackWidthCalculator", "Р”РѕРїСѓСЃС‚РёРјР°СЏ С€РёСЂРёРЅР° СЂР°СЃРєСЂС‹С‚РёСЏ С‚СЂРµС‰РёРЅ РґРѕР»Р¶РЅР° Р±С‹С‚СЊ РїРѕР»РѕР¶РёС‚РµР»СЊРЅРѕР№."
    If mCrackSpacing < 0# Or mStrainFactor < 0# Or mDurationFactor < 0# Then Err.Raise vbObjectError + 3706, "CCrackWidthCalculator", "РџР°СЂР°РјРµС‚СЂС‹ СЂР°СЃС‡РµС‚Р° С‚СЂРµС‰РёРЅ РґРѕР»Р¶РЅС‹ Р±С‹С‚СЊ РЅРµРѕС‚СЂРёС†Р°С‚РµР»СЊРЅС‹РјРё."
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



