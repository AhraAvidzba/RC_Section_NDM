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
    Dim i As Long, strain As Double, epsilon As Double, kx As Double, ky As Double
    Dim originX As Double, originY As Double, x As Double, y As Double, area As Double
    Dim cMin As Double, cMax As Double, sMin As Double, sMax As Double
    Dim cMinIndex As Long, cMaxIndex As Long, sMinIndex As Long, sMaxIndex As Long
    If Not solver.TryGetStrainExtrema(section, cMin, cMax, cMinIndex, cMaxIndex, _
            sMin, sMax, sMinIndex, sMaxIndex) Then
        ' Измененная модель или еще не оцененная плоскость требует обычного
        ' обхода. Чужие и устаревшие экстремумы никогда не используются.
        solver.NumericalStrainPlane epsilon, kx, ky, originX, originY
        For i = 1 To section.ConcreteCount
            section.ConcreteElementForSolver i, x, y, area
            strain = epsilon + kx * (y - originY) + ky * (x - originX)
            EvaluateConcrete section, i, strain
        Next i
        For i = 1 To section.RebarCount
            section.RebarElementForSolver i, x, y, area
            strain = epsilon + kx * (y - originY) + ky * (x - originX)
            EvaluateSteel section, i, strain
        Next i
        Exit Sub
    End If
    ' При равных отношениях действует прежний порядок элементов, а не
    ' приоритет сжатия или растяжения; бетон по-прежнему предшествует стали.
    If cMinIndex <= cMaxIndex Then
        EvaluateConcrete section, cMinIndex, cMin
        EvaluateConcrete section, cMaxIndex, cMax
    Else
        EvaluateConcrete section, cMaxIndex, cMax
        EvaluateConcrete section, cMinIndex, cMin
    End If
    If sMinIndex <= sMaxIndex Then
        EvaluateSteel section, sMinIndex, sMin
        EvaluateSteel section, sMaxIndex, sMax
    Else
        EvaluateSteel section, sMaxIndex, sMax
        EvaluateSteel section, sMinIndex, sMin
    End If
End Sub

' Проверяет экстремумы сжатия и активного растяжения бетона по подготовленным предельным
' деформациям.
Private Sub EvaluateConcrete(ByVal section As CSectionModel, ByVal index As Long, ByVal strain As Double)
    If index = 0 Then Exit Sub
    Dim ratio As Double
    If strain < 0# And mConcreteCompression < 0# Then
        ratio = strain / mConcreteCompression
        If ratio > mUtilization Then StoreIfLarger ratio, strain, section.ConcreteID(index) & ":Compression", _
            "ConcreteStrainLimit", "сжатие бетона"
    ElseIf strain > 0# And mTensionEnabled And mConcreteTension > 0# Then
        ratio = strain / mConcreteTension
        If ratio > mUtilization Then StoreIfLarger ratio, strain, section.ConcreteID(index) & ":Tension", _
            "ConcreteTensionStrainLimit", "растяжение бетона"
    End If
End Sub

' Проверяет сжатие и растяжение арматуры, сохраняя критический стержень и использование предела.
Private Sub EvaluateSteel(ByVal section As CSectionModel, ByVal index As Long, ByVal strain As Double)
    If index = 0 Then Exit Sub
    Dim ratio As Double
    If strain < 0# And mSteelCompression > 0# Then
        ratio = Abs(strain) / mSteelCompression
        If ratio > mUtilization Then StoreIfLarger ratio, strain, section.RebarID(index) & ":Steel", _
            "SteelStrainLimit", "сжатие арматуры"
    ElseIf strain >= 0# And mSteelTension > 0# Then
        ratio = Abs(strain) / mSteelTension
        If ratio > mUtilization Then StoreIfLarger ratio, strain, section.RebarID(index) & ":Steel", _
            "SteelStrainLimit", "растяжение арматуры"
    End If
End Sub

' Обновляет критический элемент только при большем использовании предельной деформации.
Private Sub StoreIfLarger(ByVal value As Double, ByVal strain As Double, _
        ByVal element As String, ByVal state As String, ByVal description As String)
    If value <= mUtilization Then Exit Sub
    mUtilization = value: mStrain = strain: mElement = element
    mState = state: mDescription = description
End Sub

' Сохранённый коэффициент использования проверяемого предела; смысл предела задаёт результат
' этого класса.
Public Property Get Utilization() As Double
    Utilization = mUtilization
End Property

' Деформация критического элемента, задающего максимальное использование предела.
Public Property Get CriticalStrain() As Double
    CriticalStrain = mStrain
End Property

' Идентификатор критического бетонного элемента или стержня арматуры.
Public Property Get CriticalElement() As String
    CriticalElement = mElement
End Property

' Признак достигнутого физического предела по проверяемому критерию.
Public Property Get LimitState() As String
    LimitState = mState
End Property

' Сохранённое объяснение критического физического ограничения.
Public Property Get Description() As String
    Description = mDescription
End Property
