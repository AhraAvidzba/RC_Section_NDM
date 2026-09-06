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
' Инженерный фильтр практически нулевого момента
' ==========================================================================
' CMomentZeroFilter отделяет физически значимый малый момент от численного
' мусора после переноса точки приложения нагрузки, поворота осей или фильтра
' устойчивости. Это не solver tolerance: внутренние малые допуски решателей
' остаются для математики итераций, а этот класс используется до выбора
' расчетной ветки и начального приближения.
'
' Настройка задается как момент на единицу габарита сечения. Для каждой
' плоскости порог считается по бетонному габариту:
'
'     Mtol = Calculation.ZeroMomentPerDepth * h
'
' Если Abs(M) <= Mtol, момент зануляется. Слишком большое значение настройки
' может обнулить уже инженерно значимый малый момент, поэтому пользователь
' должен выбирать его осознанно.

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

Public Property Get MomentPerDepth() As Double
    MomentPerDepth = mMomentPerDepth
End Property
