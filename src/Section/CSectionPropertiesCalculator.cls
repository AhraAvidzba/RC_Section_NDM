VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSectionPropertiesCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Геометрические свойства сечения
' ==========================================================================
' Класс считает площадь, центр тяжести, моменты инерции и главные центральные
' оси для двух инженерных представлений одной CSectionModel.
'
' CalculateConcrete работает только по бетонным конечным элементам. Этот центр
' используется как базовая точка задания пользовательских нагрузок, потому что
' внешние расчетные модели обычно привязаны к бетонному поперечному сечению без
' учета продольной арматуры.
'
' CalculateTransformed считает приведенное сечение бетон + арматура. Эти
' свойства нужны для справочного вывода, главных осей и контроля расчетной
' модели, но больше не являются базой для Load.ReferenceOffset.

Private mArea As Double ' Площадь текущего представления: бетонная или приведенная к бетону, мм2.
Private mStaticMomentX As Double ' Статический момент относительно оси X: сумма A*y.
Private mStaticMomentY As Double ' Статический момент относительно оси Y: сумма A*x.
Private mCentroidX As Double ' X центра тяжести текущего представления сечения.
Private mCentroidY As Double ' Y центра тяжести текущего представления сечения.
Private mIx As Double ' Момент инерции относительно исходной оси X.
Private mIy As Double ' Момент инерции относительно исходной оси Y.
Private mIxy As Double ' Центробежный момент инерции относительно исходных осей.
Private mIxc As Double ' Центральный момент инерции относительно оси X через центр тяжести.
Private mIyc As Double ' Центральный момент инерции относительно оси Y через центр тяжести.
Private mIxyc As Double ' Центральный центробежный момент инерции.
Private mPrincipalI1 As Double ' Больший главный центральный момент инерции.
Private mPrincipalI2 As Double ' Меньший главный центральный момент инерции.
Private mPrincipalAngleRad As Double ' Угол поворота главных осей относительно исходной системы, рад.
Private mRadiusX As Double ' Радиус инерции относительно центральной оси X.
Private mRadiusY As Double ' Радиус инерции относительно центральной оси Y.
Private mPrincipalRadius1 As Double ' Радиус инерции по главной оси I1.
Private mPrincipalRadius2 As Double ' Радиус инерции по главной оси I2.
Private mLastCalculationSeconds As Double ' Время последнего расчета свойств для диагностики производительности.

' Считает геометрические характеристики только бетонной части сечения.
' Такой расчет не зависит от диаграмм материалов и продольной арматуры. Он
' используется для Load.ReferenceOffset: нулевое смещение означает, что N, Mx и
' My заданы относительно центра тяжести бетонного сечения.
Public Sub CalculateConcrete(ByVal section As CSectionModel)
    If section Is Nothing Then Err.Raise vbObjectError + 2200, "CSectionPropertiesCalculator", "Модель сечения не передана."
    If section.ConcreteCount <= 0 Then Err.Raise vbObjectError + 2201, "CSectionPropertiesCalculator", "В модели сечения нет бетонных элементов."

    Clear

    Dim t0 As Double
    t0 = Timer

    Dim i As Long
    Dim a As Double ' Площадь бетонного элемента, полностью входящая в дискретную модель.
    Dim x As Double ' X центра тяжести бетонного элемента.
    Dim y As Double ' Y центра тяжести бетонного элемента.

    ' Проходим по бетонной сетке и накапливаем обычные геометрические моменты.
    ' Для справочных свойств частичное восстановление контура не требуется:
    ' используются те же дискретные бетонные элементы, что и в НДМ.
    For i = 1 To section.ConcreteCount
        a = section.ConcreteArea(i)
        x = section.ConcreteX(i)
        y = section.ConcreteY(i)

        mArea = mArea + a
        mStaticMomentX = mStaticMomentX + a * y
        mStaticMomentY = mStaticMomentY + a * x
        mIx = mIx + a * y * y
        mIy = mIy + a * x * x
        mIxy = mIxy + a * x * y
    Next i

    If mArea <= GEOM_TOLERANCE Then Err.Raise vbObjectError + 2202, "CSectionPropertiesCalculator", "Площадь бетонного сечения равна нулю."

    FinalizeProperties t0
End Sub

' Считает свойства приведенного сечения бетон + арматура.
' Бетон и арматура приводятся через начальный секущий модуль в малой сжатой
' деформации. Метод нужен для справочных осей и инерций, но не определяет
' пользовательскую точку приложения нагрузки.
Public Sub CalculateTransformed(ByVal section As CSectionModel, _
        ByVal concrete As Object, ByVal steel As Object)
    Const probeStrain As Double = -0.000001

    If section Is Nothing Then Err.Raise vbObjectError + 2203, "CSectionPropertiesCalculator", "Модель сечения не передана."
    If section.ConcreteCount <= 0 Then Err.Raise vbObjectError + 2204, "CSectionPropertiesCalculator", "В модели сечения нет бетонных элементов."
    If concrete Is Nothing Then Err.Raise vbObjectError + 2205, "CSectionPropertiesCalculator", "Материал бетона не передан."
    If steel Is Nothing Then Err.Raise vbObjectError + 2206, "CSectionPropertiesCalculator", "Материал арматуры не передан."

    Clear

    Dim t0 As Double
    t0 = Timer

    Dim concreteModulus As Double ' Начальная жесткость бетона, базовый модуль для приведения.
    Dim steelModulus As Double ' Начальная жесткость арматуры, из нее получаем n = Es / Eb.
    concreteModulus = concrete.GetStress(probeStrain) / probeStrain
    steelModulus = steel.GetStress(probeStrain) / probeStrain
    If Abs(concreteModulus) <= GEOM_TOLERANCE Then Err.Raise vbObjectError + 2208, _
        "CSectionPropertiesCalculator", "Начальный модуль бетона для приведенного сечения равен нулю."

    Dim steelModularRatio As Double
    steelModularRatio = steelModulus / concreteModulus

    Dim i As Long
    For i = 1 To section.ConcreteCount
        AddWeightedPoint section.ConcreteArea(i), _
            section.ConcreteX(i), section.ConcreteY(i)
    Next i

    If section.RebarCount > 0 Then
        Dim replacementAreaFactor As Double ' Добавочная площадь арматуры в бетонном приведении: n - 1.
        replacementAreaFactor = steelModularRatio - 1#
        For i = 1 To section.RebarCount
            AddWeightedPoint replacementAreaFactor * section.RebarArea(i), section.RebarX(i), section.RebarY(i)
        Next i
    End If

    If mArea <= GEOM_TOLERANCE Then Err.Raise vbObjectError + 2207, "CSectionPropertiesCalculator", "Площадь приведенного сечения равна нулю."

    FinalizeProperties t0
End Sub

' Добавляет точку с весовой площадью в накопленные статические и инерционные
' характеристики. Для бетонного сечения вес равен площади, для приведенного -
' площади, пересчитанной к бетонной базе через модульное отношение n = Es/Eb.
Private Sub AddWeightedPoint(ByVal weight As Double, ByVal x As Double, ByVal y As Double)
    mArea = mArea + weight
    mStaticMomentX = mStaticMomentX + weight * y
    mStaticMomentY = mStaticMomentY + weight * x
    mIx = mIx + weight * y * y
    mIy = mIy + weight * x * x
    mIxy = mIxy + weight * x * y
End Sub

' Завершает расчет: переводит моменты к центру тяжести, находит главные оси и
' радиусы инерции. Формулы одинаковы для бетонного и приведенного представления,
' потому что к этому моменту все различие уже учтено в накопленных весах.
Private Sub FinalizeProperties(ByVal t0 As Double)
    mCentroidX = mStaticMomentY / mArea
    mCentroidY = mStaticMomentX / mArea

    mIxc = mIx - mArea * mCentroidY * mCentroidY
    mIyc = mIy - mArea * mCentroidX * mCentroidX
    mIxyc = mIxy - mArea * mCentroidX * mCentroidY

    Dim avgI As Double
    Dim radius As Double
    avgI = (mIxc + mIyc) / 2#
    radius = Sqr(((mIxc - mIyc) / 2#) ^ 2 + mIxyc ^ 2)
    mPrincipalI1 = avgI + radius
    mPrincipalI2 = avgI - radius

    Dim inertiaRef As Double
    inertiaRef = GeomMax(GeomMax(Abs(mIxc), Abs(mIyc)), 1#)

    If Abs(mIxyc) / inertiaRef <= 0.000000001 And Abs(mIxc - mIyc) / inertiaRef <= 0.000000001 Then
        mPrincipalAngleRad = 0#
    Else
        mPrincipalAngleRad = 0.5 * Atn2(-2# * mIxyc, mIyc - mIxc)
    End If

    mRadiusX = Sqr(GeomMax(mIxc, 0#) / mArea)
    mRadiusY = Sqr(GeomMax(mIyc, 0#) / mArea)
    mPrincipalRadius1 = Sqr(GeomMax(mPrincipalI1, 0#) / mArea)
    mPrincipalRadius2 = Sqr(GeomMax(mPrincipalI2, 0#) / mArea)

    mLastCalculationSeconds = Timer - t0
End Sub

' Очищает накопленное состояние перед новым расчетом или повторным формированием вывода.
Public Sub Clear()
    mArea = 0#
    mStaticMomentX = 0#
    mStaticMomentY = 0#
    mCentroidX = 0#
    mCentroidY = 0#
    mIx = 0#
    mIy = 0#
    mIxy = 0#
    mIxc = 0#
    mIyc = 0#
    mIxyc = 0#
    mPrincipalI1 = 0#
    mPrincipalI2 = 0#
    mPrincipalAngleRad = 0#
    mRadiusX = 0#
    mRadiusY = 0#
    mPrincipalRadius1 = 0#
    mPrincipalRadius2 = 0#
    mLastCalculationSeconds = 0#
End Sub

Public Property Get Area() As Double
    Area = mArea
End Property

Public Property Get StaticMomentX() As Double
    StaticMomentX = mStaticMomentX
End Property

Public Property Get StaticMomentY() As Double
    StaticMomentY = mStaticMomentY
End Property

Public Property Get CentroidX() As Double
    CentroidX = mCentroidX
End Property

Public Property Get CentroidY() As Double
    CentroidY = mCentroidY
End Property

Public Property Get Ix() As Double
    Ix = mIx
End Property

Public Property Get Iy() As Double
    Iy = mIy
End Property

Public Property Get Ixy() As Double
    Ixy = mIxy
End Property

Public Property Get Ixc() As Double
    Ixc = mIxc
End Property

Public Property Get Iyc() As Double
    Iyc = mIyc
End Property

Public Property Get Ixyc() As Double
    Ixyc = mIxyc
End Property

Public Property Get PrincipalI1() As Double
    PrincipalI1 = mPrincipalI1
End Property

Public Property Get PrincipalI2() As Double
    PrincipalI2 = mPrincipalI2
End Property

Public Property Get PrincipalAngleRad() As Double
    PrincipalAngleRad = mPrincipalAngleRad
End Property

Public Property Get RadiusX() As Double
    RadiusX = mRadiusX
End Property

Public Property Get RadiusY() As Double
    RadiusY = mRadiusY
End Property

Public Property Get PrincipalRadius1() As Double
    PrincipalRadius1 = mPrincipalRadius1
End Property

Public Property Get PrincipalRadius2() As Double
    PrincipalRadius2 = mPrincipalRadius2
End Property

Private Function Atn2(ByVal y As Double, ByVal x As Double) As Double
    If Abs(x) <= GEOM_TOLERANCE Then
        If y > 0# Then
            Atn2 = GEOM_PI / 2#
        ElseIf y < 0# Then
            Atn2 = -GEOM_PI / 2#
        Else
            Atn2 = 0#
        End If
    ElseIf x > 0# Then
        Atn2 = Atn(y / x)
    ElseIf y >= 0# Then
        Atn2 = Atn(y / x) + GEOM_PI
    Else
        Atn2 = Atn(y / x) - GEOM_PI
    End If
End Function

Public Property Get LastCalculationSeconds() As Double
    LastCalculationSeconds = mLastCalculationSeconds
End Property











