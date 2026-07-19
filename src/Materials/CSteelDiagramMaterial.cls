VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSteelDiagramMaterial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mStrains() As Double
Private mStresses() As Double
Private mPointCount As Long
Private mInitialized As Boolean

Public Sub InitializeFromSettings(ByVal settings As CSystemSettingsReader)
    InitializeFromThreeTensionPoints _
        settings.GetDouble("Steel.Point1.Eps", 0.00175), _
        settings.GetDouble("Steel.Point1.Stress", 350#), _
        settings.GetDouble("Steel.Point2.Eps", 0.01), _
        settings.GetDouble("Steel.Point2.Stress", 350#), _
        settings.GetDouble("Steel.Point3.Eps", 0.025), _
        settings.GetDouble("Steel.Point3.Stress", 350#)
End Sub

Public Sub InitializeFromTwoTensionPoints(ByVal eps1 As Double, ByVal stress1 As Double, _
        ByVal eps2 As Double, ByVal stress2 As Double)
    If eps1 <= 0# Then Err.Raise vbObjectError + 3120, "CSteelDiagramMaterial", "first tension strain must be positive."
    If eps2 <= eps1 Then Err.Raise vbObjectError + 3121, "CSteelDiagramMaterial", "second tension strain must be greater than first."
    If stress1 <= 0# Or stress2 <= 0# Then Err.Raise vbObjectError + 3122, "CSteelDiagramMaterial", "Steel tension stresses must be positive."

    ReDim mStrains(1 To 3)
    ReDim mStresses(1 To 3)
    mStrains(1) = 0#: mStresses(1) = 0#
    mStrains(2) = eps1: mStresses(2) = stress1
    mStrains(3) = eps2: mStresses(3) = stress2
    mPointCount = 3
    mInitialized = True
End Sub

Public Sub Initialize(ByVal eps1 As Double, ByVal stress1 As Double, ByVal eps2 As Double)
    InitializeFromTwoTensionPoints eps1, stress1, eps2, stress1
End Sub

Public Sub InitializeFromThreeTensionPoints(ByVal eps1 As Double, ByVal stress1 As Double, _
        ByVal eps2 As Double, ByVal stress2 As Double, ByVal eps3 As Double, ByVal stress3 As Double)
    If eps1 <= 0# Then Err.Raise vbObjectError + 3120, "CSteelDiagramMaterial", "first tension strain must be positive."
    If eps2 <= eps1 Then Err.Raise vbObjectError + 3121, "CSteelDiagramMaterial", "second tension strain must be greater than first."
    If eps3 <= eps2 Then Err.Raise vbObjectError + 3124, "CSteelDiagramMaterial", "third tension strain must be greater than second."
    If stress1 <= 0# Or stress2 <= 0# Or stress3 <= 0# Then Err.Raise vbObjectError + 3122, "CSteelDiagramMaterial", "Steel tension stresses must be positive."

    ReDim mStrains(1 To 4)
    ReDim mStresses(1 To 4)
    mStrains(1) = 0#: mStresses(1) = 0#
    mStrains(2) = eps1: mStresses(2) = stress1
    mStrains(3) = eps2: mStresses(3) = stress2
    mStrains(4) = eps3: mStresses(4) = stress3
    mPointCount = 4
    mInitialized = True
End Sub

Public Property Get UltimateTensionStrain() As Double
    EnsureInitialized
    UltimateTensionStrain = mStrains(mPointCount)
End Property

Public Function GetStress(ByVal strain As Double) As Double
    EnsureInitialized

    Dim signValue As Double
    signValue = SignOf(strain)
    GetStress = signValue * PositiveStress(Abs(strain))
End Function

Public Function GetTangentModulus(ByVal strain As Double) As Double
    EnsureInitialized
    GetTangentModulus = PositiveTangent(Abs(strain))
End Function

Public Function GetSecantModulus(ByVal strain As Double) As Double
    EnsureInitialized
    If Abs(strain) <= 0.000000000001 Then
        GetSecantModulus = GetTangentModulus(strain)
    Else
        GetSecantModulus = GetStress(strain) / strain
    End If
End Function

Private Function PositiveStress(ByVal strain As Double) As Double
    If strain >= mStrains(mPointCount) Then
        PositiveStress = mStresses(mPointCount)
    Else
        Dim i As Long
        For i = 1 To mPointCount - 1
            If strain >= mStrains(i) And strain <= mStrains(i + 1) Then
                PositiveStress = mStresses(i) + (mStresses(i + 1) - mStresses(i)) * _
                    (strain - mStrains(i)) / (mStrains(i + 1) - mStrains(i))
                Exit Function
            End If
        Next i
    End If
End Function

Private Function PositiveTangent(ByVal strain As Double) As Double
    If strain >= mStrains(mPointCount) Then Exit Function
    Dim i As Long
    For i = 1 To mPointCount - 1
        If strain >= mStrains(i) And strain <= mStrains(i + 1) Then
            PositiveTangent = (mStresses(i + 1) - mStresses(i)) / (mStrains(i + 1) - mStrains(i))
            Exit Function
        End If
    Next i
End Function

Private Function SignOf(ByVal value As Double) As Double
    If value < 0# Then
        SignOf = -1#
    ElseIf value > 0# Then
        SignOf = 1#
    End If
End Function

Private Sub EnsureInitialized()
    If Not mInitialized Then Err.Raise vbObjectError + 3123, "CSteelDiagramMaterial", "Material is not initialized."
End Sub
