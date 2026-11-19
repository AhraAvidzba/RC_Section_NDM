VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLoadZeroFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Независимая проверка практически нулевой продольной силы и моментов
' ==========================================================================
' Обнуляет N до построения траектории по минимальному бетонному габариту,
' а физические компоненты моментов в пользовательском вводе или главных
' плоскостях устойчивости. Служебные слагаемые переноса N*deltaX и N*deltaY сохраняются;
' вызывающий код не должен фильтровать ими смену начала координат. Порог применяется до выбора расчёта
' и начального приближения; допуски численного равновесия задаются отдельно.
' Ntol = Calculation.ZeroAxialForcePerDepth * min(hx, hy).
' Малая N сохраняется, если хотя бы один её момент от эксцентриситета
' относительно бетонного центра превышает порог соответствующей плоскости.
' Для каждой плоскости Mtol = Calculation.ZeroMomentPerDepth * h, где h —
' соответствующий габарит бетонного сечения. При Abs(M) <= Mtol момент обнуляется.
' Завышенный порог может исключить значимый момент, поэтому задаётся пользователем.

Private mMomentPerDepth As Double ' Внутреннее значение порога: Н*мм/мм.
Private mAxialForcePerDepth As Double ' Независимый коэффициент порога N: Н/мм; не применяется к пробам lambda.

' Принимает независимые пороги момента и силы на единицу габарита во внутренних единицах.
Public Sub Initialize(ByVal momentPerDepth As Double, Optional ByVal axialForcePerDepth As Double = 0#)
    If momentPerDepth < 0# Then Err.Raise vbObjectError + 6010, _
        "CLoadZeroFilter", "Calculation.ZeroMomentPerDepth не может быть отрицательным."
    If axialForcePerDepth < 0# Then Err.Raise vbObjectError + 6011, _
        "CLoadZeroFilter", "Calculation.ZeroAxialForcePerDepth не может быть отрицательным."
    mMomentPerDepth = momentPerDepth
    mAxialForcePerDepth = axialForcePerDepth
End Sub

' Возвращает порог N по меньшему габариту бетонного сечения; нулевой коэффициент отключает фильтр силы.
Public Function AxialForceTolerance(ByVal concreteDepthX As Double, ByVal concreteDepthY As Double) As Double
    If mAxialForcePerDepth <= 0# Or concreteDepthX <= 0# Or concreteDepthY <= 0# Then Exit Function
    Dim minimumDepth As Double: minimumDepth = concreteDepthX
    If concreteDepthY < minimumDepth Then minimumDepth = concreteDepthY
    On Error GoTo Overflow
    AxialForceTolerance = mAxialForcePerDepth * minimumDepth
    Exit Function
Overflow:
    Err.Raise vbObjectError + 6012, "CLoadZeroFilter", "Порог N по минимальному габариту сечения выходит за допустимый числовой диапазон."
End Function

' Проверяет исходную силу до построения пути, одинаково для сжатия и растяжения; равенство порогу включает зануление.
Public Function IsAxialForceEffectivelyZero(ByVal forceValue As Double, ByVal concreteDepthX As Double, ByVal concreteDepthY As Double, _
        Optional ByVal offsetX As Double = 0#, Optional ByVal offsetY As Double = 0#) As Boolean
    If Abs(forceValue) > AxialForceTolerance(concreteDepthX, concreteDepthY) Then Exit Function
    If Not IsAxialEccentricMomentZero(forceValue, offsetY, concreteDepthY) Then Exit Function
    If Not IsAxialEccentricMomentZero(forceValue, offsetX, concreteDepthX) Then Exit Function
    IsAxialForceEffectivelyZero = True
End Function

' Сравнивает N*e с порогом без умножения больших чисел; e отсчитывается от бетонного центра, не от начала координат.
Private Function IsAxialEccentricMomentZero(ByVal forceValue As Double, ByVal offsetValue As Double, ByVal concreteDepth As Double) As Boolean
    If forceValue = 0# Or offsetValue = 0# Then
        IsAxialEccentricMomentZero = True
    ElseIf Abs(offsetValue) <= 1# Then
        IsAxialEccentricMomentZero = Abs(forceValue * offsetValue) <= MomentTolerance(concreteDepth)
    Else
        IsAxialEccentricMomentZero = Abs(forceValue) <= MomentTolerance(concreteDepth) / Abs(offsetValue)
    End If
End Function

' Зануляет только исходную N; вызывающий код затем рассчитывает её моменты переноса и lambda-траекторию.
Public Function NormalizeAxialForce(ByVal forceValue As Double, ByVal concreteDepthX As Double, ByVal concreteDepthY As Double, _
        Optional ByVal offsetX As Double = 0#, Optional ByVal offsetY As Double = 0#) As Double
    If IsAxialForceEffectivelyZero(forceValue, concreteDepthX, concreteDepthY, offsetX, offsetY) Then
        NormalizeAxialForce = 0#
    Else
        NormalizeAxialForce = forceValue
    End If
End Function

' Возвращает порог момента для конкретного бетонного габарита h.
Public Function MomentTolerance(ByVal concreteDepth As Double) As Double
    If mMomentPerDepth <= 0# Or concreteDepth <= 0# Then Exit Function
    On Error GoTo Overflow
    MomentTolerance = mMomentPerDepth * concreteDepth
    Exit Function
Overflow:
    Err.Raise vbObjectError + 6013, "CLoadZeroFilter", "Порог момента по габариту сечения выходит за допустимый числовой диапазон."
End Function

' Проверяет инженерный ноль момента для заданной плоскости.
Public Function IsMomentEffectivelyZero(ByVal momentValue As Double, ByVal concreteDepth As Double) As Boolean
    Dim toleranceValue As Double
    toleranceValue = MomentTolerance(concreteDepth)
    If toleranceValue <= 0# Then
        IsMomentEffectivelyZero = (momentValue = 0#)
    Else
        IsMomentEffectivelyZero = (Abs(momentValue) <= toleranceValue)
    End If
End Function

' Зануляет момент, модуль которого не превышает инженерный порог для данной плоскости.
Public Function NormalizeMoment(ByVal momentValue As Double, ByVal concreteDepth As Double) As Double
    If IsMomentEffectivelyZero(momentValue, concreteDepth) Then
        NormalizeMoment = 0#
    Else
        NormalizeMoment = momentValue
    End If
End Function

' Внутреннее значение порога: Н*мм/мм.
Public Property Get MomentPerDepth() As Double
    MomentPerDepth = mMomentPerDepth
End Property
