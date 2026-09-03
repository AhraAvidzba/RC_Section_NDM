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
    ' Проходим по бетонной сетке и накапливаем геометрические моменты. Если
    ' элемент пришел из AutoCAD Region или генератора уже с собственными
    ' инерциями вокруг своего центра, учитываем их здесь же: это единая точка
    ' расчета геометрических характеристик для всех последующих проверок.
    For i = 1 To section.ConcreteCount
        AddWeightedArea section.ConcreteArea(i), _
            section.ConcreteX(i), section.ConcreteY(i), _
            section.ConcreteLocalIx(i), section.ConcreteLocalIy(i), section.ConcreteLocalIxy(i)
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
        AddWeightedArea section.ConcreteArea(i), _
            section.ConcreteX(i), section.ConcreteY(i), _
            section.ConcreteLocalIx(i), section.ConcreteLocalIy(i), section.ConcreteLocalIxy(i)
    Next i

    If section.RebarCount > 0 Then
        Dim replacementAreaFactor As Double ' Добавочная площадь арматуры в бетонном приведении: n - 1.
        Dim rebarWeight As Double
        Dim rebarLocalI As Double
        replacementAreaFactor = steelModularRatio - 1#
        For i = 1 To section.RebarCount
            rebarWeight = replacementAreaFactor * section.RebarArea(i)
            rebarLocalI = replacementAreaFactor * RebarCircleInertia(section.RebarArea(i))
            AddWeightedArea rebarWeight, section.RebarX(i), section.RebarY(i), rebarLocalI, rebarLocalI, 0#
        Next i
    End If

    If mArea <= GEOM_TOLERANCE Then Err.Raise vbObjectError + 2207, "CSectionPropertiesCalculator", "Площадь приведенного сечения равна нулю."

    FinalizeProperties t0
End Sub

' Добавляет элемент с весовой площадью и собственными инерциями вокруг центра.
' Для бетонного сечения weight равен площади. Для приведенного сечения арматура
' добавляется как (n-1)*As, поэтому ее собственная инерция также умножается на
' тот же коэффициент приведения.
Private Sub AddWeightedArea(ByVal weight As Double, ByVal x As Double, ByVal y As Double, _
        Optional ByVal localIx As Double = 0#, Optional ByVal localIy As Double = 0#, _
        Optional ByVal localIxy As Double = 0#)
    mArea = mArea + weight
    mStaticMomentX = mStaticMomentX + weight * y
    mStaticMomentY = mStaticMomentY + weight * x
    mIx = mIx + localIx + weight * y * y
    mIy = mIy + localIy + weight * x * x
    mIxy = mIxy + localIxy + weight * x * y
End Sub

' Возвращает проекцию расчетного сечения на произвольную ось.
' Метод централизует высоты/габариты, которые нужны Summary, трещинам и
' устойчивости: бетон берется по сохраненной ширине/высоте элемента или по
' эквивалентному кругу, арматура - по эквивалентному радиусу из площади.
Public Sub CalculateProjection(ByVal section As CSectionModel, ByVal axisX As Double, ByVal axisY As Double, _
        ByVal includeRebar As Boolean, ByRef minProjection As Double, ByRef maxProjection As Double)
    If section Is Nothing Then Err.Raise vbObjectError + 2210, "CSectionPropertiesCalculator", "Модель сечения не передана."
    If section.ConcreteCount <= 0 Then Err.Raise vbObjectError + 2211, "CSectionPropertiesCalculator", "В модели сечения нет бетонных элементов."

    Dim normValue As Double
    normValue = Sqr(axisX * axisX + axisY * axisY)
    If normValue <= GEOM_TOLERANCE Then Err.Raise vbObjectError + 2212, "CSectionPropertiesCalculator", "Ось проекции имеет нулевую длину."

    Dim nx As Double
    Dim ny As Double
    nx = axisX / normValue
    ny = axisY / normValue

    Dim initialized As Boolean
    Dim i As Long
    For i = 1 To section.ConcreteCount
        ExtendProjection initialized, minProjection, maxProjection, _
            nx * section.ConcreteX(i) + ny * section.ConcreteY(i), _
            ConcreteHalfProjection(section, i, nx, ny)
    Next i

    If includeRebar And section.RebarCount > 0 Then
        For i = 1 To section.RebarCount
            ExtendProjection initialized, minProjection, maxProjection, _
                nx * section.RebarX(i) + ny * section.RebarY(i), _
                RebarEquivalentRadius(section.RebarArea(i))
        Next i
    End If
End Sub

' Возвращает полную высоту/габарит сечения по заданной оси проекции.
Public Function ProjectionDepth(ByVal section As CSectionModel, ByVal axisX As Double, ByVal axisY As Double, _
        Optional ByVal includeRebar As Boolean = True) As Double
    Dim minProjection As Double
    Dim maxProjection As Double
    CalculateProjection section, axisX, axisY, includeRebar, minProjection, maxProjection
    ProjectionDepth = maxProjection - minProjection
End Function

' Возвращает суммарную площадь продольной арматуры.
' Метод живет в геометрическом калькуляторе, потому что это чистая
' характеристика CSectionModel и она нужна разным нормативным проверкам.
Public Function RebarAreaTotal(ByVal section As CSectionModel) As Double
    If section Is Nothing Then Err.Raise vbObjectError + 2213, "CSectionPropertiesCalculator", "Модель сечения не передана."

    Dim i As Long
    For i = 1 To section.RebarCount
        RebarAreaTotal = RebarAreaTotal + section.RebarArea(i)
    Next i
End Function

' Считает момент инерции всей арматуры относительно произвольной оси.
' Ось задается направляющим вектором и точкой прохождения. Для каждого
' стержня учитывается перенос площади и собственная инерция круглого сечения,
' вычисленная через площадь: I = A^2/(4*pi).
Public Function RebarInertiaAboutAxis(ByVal section As CSectionModel, ByVal axisX As Double, _
        ByVal axisY As Double, ByVal originX As Double, ByVal originY As Double) As Double
    If section Is Nothing Then Err.Raise vbObjectError + 2214, "CSectionPropertiesCalculator", "Модель сечения не передана."

    Dim normValue As Double
    normValue = Sqr(axisX * axisX + axisY * axisY)
    If normValue <= GEOM_TOLERANCE Then Err.Raise vbObjectError + 2215, "CSectionPropertiesCalculator", "Ось инерции арматуры имеет нулевую длину."

    Dim nx As Double
    Dim ny As Double
    nx = axisX / normValue
    ny = axisY / normValue

    Dim i As Long
    Dim dx As Double
    Dim dy As Double
    Dim distanceToAxis As Double
    For i = 1 To section.RebarCount
        dx = section.RebarX(i) - originX
        dy = section.RebarY(i) - originY
        distanceToAxis = -ny * dx + nx * dy
        RebarInertiaAboutAxis = RebarInertiaAboutAxis + _
            section.RebarArea(i) * distanceToAxis * distanceToAxis + RebarCircleInertia(section.RebarArea(i))
    Next i
End Function

' Расширяет накопленный диапазон проекций одним элементом.
Private Sub ExtendProjection(ByRef initialized As Boolean, ByRef minProjection As Double, ByRef maxProjection As Double, _
        ByVal centerProjection As Double, ByVal halfProjection As Double)
    If Not initialized Then
        minProjection = centerProjection - halfProjection
        maxProjection = centerProjection + halfProjection
        initialized = True
    Else
        minProjection = GeomMin(minProjection, centerProjection - halfProjection)
        maxProjection = GeomMax(maxProjection, centerProjection + halfProjection)
    End If
End Sub

' Проекция бетонного элемента. Для обычной сетки используется сохраненный
' прямоугольник с поворотом. Если точных размеров нет, берется эквивалентный
' круг по площади: это устойчивый fallback для произвольных Region.
Private Function ConcreteHalfProjection(ByVal section As CSectionModel, ByVal index As Long, _
        ByVal axisX As Double, ByVal axisY As Double) As Double
    Dim widthValue As Double
    Dim heightValue As Double
    widthValue = section.ConcreteWidth(index)
    heightValue = section.ConcreteHeight(index)

    If widthValue > GEOM_TOLERANCE And heightValue > GEOM_TOLERANCE Then
        Dim angleValue As Double
        Dim ux As Double
        Dim uy As Double
        Dim vx As Double
        Dim vy As Double
        angleValue = section.ConcreteRotation(index)
        ux = Cos(angleValue)
        uy = Sin(angleValue)
        vx = -uy
        vy = ux
        ConcreteHalfProjection = 0.5 * (widthValue * Abs(axisX * ux + axisY * uy) + _
            heightValue * Abs(axisX * vx + axisY * vy))
    ElseIf section.ConcreteArea(index) > GEOM_TOLERANCE Then
        ConcreteHalfProjection = Sqr(section.ConcreteArea(index) / GEOM_PI)
    End If
End Function

' Собственная инерция круглого стержня через его площадь: I = A^2/(4*pi).
Private Function RebarCircleInertia(ByVal areaValue As Double) As Double
    If areaValue > GEOM_TOLERANCE Then RebarCircleInertia = areaValue * areaValue / (4# * GEOM_PI)
End Function

' Радиус круглого стержня через площадь. Используется для габарита по оси.
Private Function RebarEquivalentRadius(ByVal areaValue As Double) As Double
    If areaValue > GEOM_TOLERANCE Then RebarEquivalentRadius = Sqr(areaValue / GEOM_PI)
End Function

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











