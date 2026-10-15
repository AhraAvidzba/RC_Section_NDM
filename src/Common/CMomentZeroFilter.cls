VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMomentZeroFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Проверка практически нулевых моментов
' ==========================================================================
' Обнуляет моменты, меньшие заданного инженерного порога, после переноса нагрузки,
' поворота осей или учёта устойчивости. Порог применяется до выбора расчёта
' и начального приближения; допуски численного равновесия задаются отдельно.
' Для каждой плоскости Mtol = Calculation.ZeroMomentPerDepth * h, где h —
' соответствующий габарит бетонного сечения. При Abs(M) <= Mtol момент обнуляется.
' Завышенный порог может исключить значимый момент, поэтому задаётся пользователем.

Private mMomentPerDepth As Double ' Внутреннее значение порога: Н*мм/мм.

' Принимает порог момента на единицу габарита во внутренних единицах.
Public Sub Initialize(ByVal momentPerDepth As Double)
    If momentPerDepth < 0# Then Err.Raise vbObjectError + 6010, _
        "CMomentZeroFilter", "Calculation.ZeroMomentPerDepth не может быть отрицательным."
    mMomentPerDepth = momentPerDepth
End Sub

' Возвращает порог момента для конкретного бетонного габарита h.
Public Function MomentTolerance(ByVal concreteDepth As Double) As Double
    If mMomentPerDepth <= 0# Or concreteDepth <= 0# Then Exit Function
    MomentTolerance = mMomentPerDepth * concreteDepth
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

' Зануляет момент, если он меньше инженерного порога для данной плоскости.
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
