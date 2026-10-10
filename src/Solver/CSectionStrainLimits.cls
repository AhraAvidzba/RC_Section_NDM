VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSectionStrainLimits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Общий критерий предельной плоскости: бетон в сжатии/растяжении и арматура
' в обоих направлениях. Оценивает фактические деформации, не прогноз усилий.
Private mConcreteCompression As Double, mConcreteTension As Double
Private mSteelCompression As Double, mSteelTension As Double
Private mTensionEnabled As Boolean
Private mUtilization As Double, mStrain As Double
Private mElement As String, mState As String, mDescription As String

' Сжатый предел бетона отрицателен; пределы арматуры передаются по модулю.
' Задача трещинообразования передает свой растянутый критерий вместо предела
' несущей способности. Остальные физические ограничения остаются теми же.
Public Sub Configure(ByVal concreteCompression As Double, ByVal concreteTension As Double, _
        ByVal steelCompression As Double, ByVal steelTension As Double, _
        Optional ByVal concreteTensionEnabled As Boolean = True)
    mConcreteCompression = concreteCompression
    mConcreteTension = concreteTension
    mSteelCompression = steelCompression
    mSteelTension = steelTension
    mTensionEnabled = concreteTensionEnabled
End Sub

' Максимальное отношение деформации к своему пределу определяет ближайшую
' границу. При равенстве сохраняется прежний порядок: бетон, затем арматура.
Public Sub Evaluate(ByVal section As CSectionModel, ByVal solver As CSectionSolver)
    mUtilization = 0#: mStrain = 0#: mElement = vbNullString
    mState = "NumericalFailure": mDescription = vbNullString
    Dim i As Long, strain As Double
    For i = 1 To section.ConcreteCount
        strain = solver.StrainAt(section.ConcreteX(i), section.ConcreteY(i))
        If strain < 0# And mConcreteCompression < 0# Then
            StoreIfLarger strain / mConcreteCompression, strain, section.ConcreteID(i) & ":Compression", _
                "ConcreteStrainLimit", "сжатие бетона"
        ElseIf strain > 0# And mTensionEnabled And mConcreteTension > 0# Then
            StoreIfLarger strain / mConcreteTension, strain, section.ConcreteID(i) & ":Tension", _
                "ConcreteTensionStrainLimit", "растяжение бетона"
        End If
    Next i
    For i = 1 To section.RebarCount
        strain = solver.StrainAt(section.RebarX(i), section.RebarY(i))
        If strain < 0# And mSteelCompression > 0# Then
            StoreIfLarger Abs(strain) / mSteelCompression, strain, section.RebarID(i) & ":Steel", _
                "SteelStrainLimit", "сжатие арматуры"
        ElseIf strain >= 0# And mSteelTension > 0# Then
            StoreIfLarger Abs(strain) / mSteelTension, strain, section.RebarID(i) & ":Steel", _
                "SteelStrainLimit", "растяжение арматуры"
        End If
    Next i
End Sub

Private Sub StoreIfLarger(ByVal value As Double, ByVal strain As Double, _
        ByVal element As String, ByVal state As String, ByVal description As String)
    If value <= mUtilization Then Exit Sub
    mUtilization = value: mStrain = strain: mElement = element
    mState = state: mDescription = description
End Sub

Public Property Get Utilization() As Double
    Utilization = mUtilization
End Property

Public Property Get CriticalStrain() As Double
    CriticalStrain = mStrain
End Property

Public Property Get CriticalElement() As String
    CriticalElement = mElement
End Property

Public Property Get LimitState() As String
    LimitState = mState
End Property

Public Property Get Description() As String
    Description = mDescription
End Property
