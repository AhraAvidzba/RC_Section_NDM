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
' свойства нужны для справочного вывода, главных осей, ядровых расстояний и
' нормативных расчетов, где СП прямо требует центр приведенного сечения.
' Пользовательская точка Load.ReferenceOffset при этом остается привязанной к
' бетонному центру; перенос к приведенному центру выполняет вызывающий код.

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
Private mPrincipalAngleRad As Double ' Угол от +X к физической главной оси 1, для которой момент инерции равен I1.
Private mRadiusX As Double ' Радиус инерции относительно центральной оси X.
Private mRadiusY As Double ' Радиус инерции относительно центральной оси Y.
Private mPrincipalRadius1 As Double ' Радиус инерции по главной оси I1.
Private mPrincipalRadius2 As Double ' Радиус инерции по главной оси I2.
Private mLastCalculationSeconds As Double ' Время последнего расчета свойств для диагностики производительности.

Private Const PRINCIPAL_AXIS_ISOTROPIC_RELATIVE_TOLERANCE As Double = 0.000001 ' Допуск для почти круговых/изотропных сечений.
Private Const PROJECTION_INTERVAL_TOLERANCE As Double = 0.000001 ' Допуск склейки соседних КЭ на луче, мм.

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
        AddConcreteWeightedArea section, i, 1#
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

    Dim concreteModulus As Double ' Начальная жесткость бетона, базовый модуль для приведения.
    Dim steelModulus As Double ' Начальная жесткость арматуры, из нее получаем n = Es / Eb.
    concreteModulus = concrete.GetStress(probeStrain) / probeStrain
    steelModulus = steel.GetStress(probeStrain) / probeStrain
    CalculateTransformedByModuli section, concreteModulus, steelModulus
End Sub

' Считает свойства приведенного сечения по уже известным модулям Eb и Es.
' Этот вход используется нормативными проверками, которым нужны приведенный
' центр и инерции, но не нужна полноценная диаграмма материала.
Public Sub CalculateTransformedByModuli(ByVal section As CSectionModel, _
        ByVal concreteModulus As Double, ByVal steelModulus As Double)
    If section Is Nothing Then Err.Raise vbObjectError + 2203, "CSectionPropertiesCalculator", "Модель сечения не передана."
    If section.ConcreteCount <= 0 Then Err.Raise vbObjectError + 2204, "CSectionPropertiesCalculator", "В модели сечения нет бетонных элементов."
    If Abs(concreteModulus) <= GEOM_TOLERANCE Then Err.Raise vbObjectError + 2208, _
        "CSectionPropertiesCalculator", "Начальный модуль бетона для приведенного сечения равен нулю."

    Clear

    Dim t0 As Double
    t0 = Timer

    Dim steelModularRatio As Double
    steelModularRatio = steelModulus / concreteModulus

    Dim i As Long
    For i = 1 To section.ConcreteCount
        AddConcreteWeightedArea section, i, 1#
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

' Добавляет бетонный элемент в накопление A/I через CSectionModel.
' Модель сама возвращает реальные локальные инерции или последний fallback,
' поэтому калькулятор свойств не интерпретирует Width/Height как источник I.
Private Sub AddConcreteWeightedArea(ByVal section As CSectionModel, ByVal index As Long, _
        ByVal weightFactor As Double)
    Dim localIx As Double
    Dim localIy As Double
    Dim localIxy As Double
    section.ConcreteLocalInertiaComponents index, localIx, localIy, localIxy

    AddWeightedArea section.ConcreteArea(index) * weightFactor, _
        section.ConcreteX(index), section.ConcreteY(index), _
        localIx * weightFactor, localIy * weightFactor, localIxy * weightFactor
End Sub

' Возвращает внешнюю проекционную оболочку сечения на произвольную ось.
' includeRebar управляет только тем, расширяют ли стержни эту внешнюю оболочку
' своими эквивалентными радиусами. Он не отвечает за приведенные A/I: если
' свойства уже посчитаны через CalculateTransformedByModuli, арматура в них
' участвует независимо от includeRebar.
' Для габарита h сжатого бетонного элемента нужно includeRebar = False, потому
' что стержни находятся внутри бетонного контура. Для чисто графических задач,
' где важно поместить на схему и бетон, и арматуру, можно использовать True.
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

' Ищет расстояние от заданной точки до ближайшего выхода из бетонного контура
' вдоль выбранного направления. В отличие от общей ProjectionDepth, метод не
' берет глобальную опорную линию всего сечения: он проходит лучом через
' конкретную точку, поэтому подходит для a_s в расчете трещин у несимметричных
' сечений и косой нейтральной линии.
' Возвращает -1, если точка/луч не попали в дискретный бетонный контур.
Public Function ConcreteCoverFromPointAlongDirection(ByVal section As CSectionModel, _
        ByVal pointX As Double, ByVal pointY As Double, _
        ByVal directionX As Double, ByVal directionY As Double) As Double
    If section Is Nothing Then Err.Raise vbObjectError + 2220, "CSectionPropertiesCalculator", "Модель сечения не передана."
    If section.ConcreteCount <= 0 Then Err.Raise vbObjectError + 2221, "CSectionPropertiesCalculator", "В модели сечения нет бетонных элементов."

    Dim normValue As Double
    normValue = Sqr(directionX * directionX + directionY * directionY)
    If normValue <= GEOM_TOLERANCE Then Err.Raise vbObjectError + 2222, "CSectionPropertiesCalculator", "Направление до бетонной границы имеет нулевую длину."

    Dim nx As Double
    Dim ny As Double
    nx = directionX / normValue
    ny = directionY / normValue

    Dim intervalStart() As Double
    Dim intervalEnd() As Double
    ReDim intervalStart(1 To section.ConcreteCount)
    ReDim intervalEnd(1 To section.ConcreteCount)

    Dim intervalCount As Long
    Dim i As Long
    For i = 1 To section.ConcreteCount
        Dim tMin As Double
        Dim tMax As Double
        If ConcreteRayInterval(section, i, pointX, pointY, nx, ny, tMin, tMax) Then
            If tMax >= -PROJECTION_INTERVAL_TOLERANCE Then
                intervalCount = intervalCount + 1
                intervalStart(intervalCount) = tMin
                intervalEnd(intervalCount) = tMax
            End If
        End If
    Next i

    If intervalCount = 0 Then
        ConcreteCoverFromPointAlongDirection = -1#
        Exit Function
    End If

    Dim hasSeed As Boolean
    Dim currentEnd As Double
    For i = 1 To intervalCount
        If intervalStart(i) <= PROJECTION_INTERVAL_TOLERANCE And _
                intervalEnd(i) >= -PROJECTION_INTERVAL_TOLERANCE Then
            If Not hasSeed Or intervalEnd(i) > currentEnd Then currentEnd = intervalEnd(i)
            hasSeed = True
        End If
    Next i

    If Not hasSeed Then
        ConcreteCoverFromPointAlongDirection = -1#
        Exit Function
    End If

    ' Склеиваем соседние интервалы вдоль луча, пока не встретим первый реальный
    ' выход из бетонной области. Дальние несвязанные части Г-сечения дальше по
    ' лучу не должны увеличивать a_s.
    Dim changed As Boolean
    Do
        changed = False
        For i = 1 To intervalCount
            If intervalStart(i) <= currentEnd + PROJECTION_INTERVAL_TOLERANCE And _
                    intervalEnd(i) > currentEnd Then
                currentEnd = intervalEnd(i)
                changed = True
            End If
        Next i
    Loop While changed

    ConcreteCoverFromPointAlongDirection = GeomMax(currentEnd, 0#)
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

' Считает момент инерции бетонной части относительно произвольной оси.
' Ось задается направлением и точкой прохождения. Для каждого бетонного
' элемента учитываются его собственные локальные инерции и перенос площади.
Public Function ConcreteInertiaAboutAxis(ByVal section As CSectionModel, ByVal axisX As Double, _
        ByVal axisY As Double, ByVal originX As Double, ByVal originY As Double) As Double
    If section Is Nothing Then Err.Raise vbObjectError + 2216, "CSectionPropertiesCalculator", "Модель сечения не передана."

    Dim normValue As Double
    normValue = Sqr(axisX * axisX + axisY * axisY)
    If normValue <= GEOM_TOLERANCE Then Err.Raise vbObjectError + 2217, "CSectionPropertiesCalculator", "Ось инерции бетона имеет нулевую длину."

    Dim nx As Double
    Dim ny As Double
    nx = axisX / normValue
    ny = axisY / normValue

    Dim i As Long
    Dim dx As Double
    Dim dy As Double
    Dim distanceToAxis As Double
    Dim localInertia As Double
    For i = 1 To section.ConcreteCount
        dx = section.ConcreteX(i) - originX
        dy = section.ConcreteY(i) - originY
        distanceToAxis = -ny * dx + nx * dy
        Dim localIx As Double
        Dim localIy As Double
        Dim localIxy As Double
        section.ConcreteLocalInertiaComponents i, localIx, localIy, localIxy
        localInertia = ny * ny * localIy + nx * nx * localIx - 2# * nx * ny * localIxy
        ConcreteInertiaAboutAxis = ConcreteInertiaAboutAxis + _
            localInertia + section.ConcreteArea(i) * distanceToAxis * distanceToAxis
    Next i
End Function

' Возвращает центральный момент инерции относительно оси, перпендикулярной
' указанному направлению. Это величина I_n = integral(s^2 dA), где s - проекция
' точки на направление. Она нужна для ядрового расстояния r = I/(A*c).
Public Function ProjectedInertiaAboutCentroid(ByVal directionX As Double, ByVal directionY As Double) As Double
    Dim normValue As Double
    normValue = Sqr(directionX * directionX + directionY * directionY)
    If normValue <= GEOM_TOLERANCE Then Err.Raise vbObjectError + 2218, "CSectionPropertiesCalculator", "Направление для инерции имеет нулевую длину."

    Dim nx As Double
    Dim ny As Double
    nx = directionX / normValue
    ny = directionY / normValue
    ProjectedInertiaAboutCentroid = nx * nx * mIyc + ny * ny * mIxc + 2# * nx * ny * mIxyc
End Function

' Возвращает направление физической главной оси 1 или 2.
' PrincipalAngleRad отсчитывается от глобальной +X против часовой стрелки до
' оси 1, у которой момент инерции равен PrincipalI1. Ось 2 получается
' поворотом оси 1 на +90 градусов и имеет момент инерции PrincipalI2.
Public Sub PrincipalAxisDirection(ByVal axisIndex As Long, ByRef axisX As Double, ByRef axisY As Double)
    Dim axis1X As Double
    Dim axis1Y As Double
    axis1X = Cos(mPrincipalAngleRad)
    axis1Y = Sin(mPrincipalAngleRad)

    If axisIndex = 1 Then
        axisX = axis1X
        axisY = axis1Y
    ElseIf axisIndex = 2 Then
        axisX = -axis1Y
        axisY = axis1X
    Else
        Err.Raise vbObjectError + 2223, "CSectionPropertiesCalculator", "Номер главной оси должен быть 1 или 2."
    End If
End Sub

' Возвращает положительное направление эксцентриситета для главной плоскости.
' Важно: главная плоскость и главная ось - не одно и то же направление.
' Плоскость 1 означает изгиб вокруг главной оси 1; плечо, h и ядровое
' расстояние этой плоскости считаются по нормали к оси 1, то есть вдоль оси 2.
' Плоскость 2 аналогично связана с изгибом вокруг оси 2 и нормалью вдоль оси 1.
Public Sub PrincipalPlaneNormalDirection(ByVal planeIndex As Long, ByRef normalX As Double, ByRef normalY As Double)
    If planeIndex = 1 Then
        PrincipalAxisDirection 2, normalX, normalY
    ElseIf planeIndex = 2 Then
        PrincipalAxisDirection 1, normalX, normalY
    Else
        Err.Raise vbObjectError + 2224, "CSectionPropertiesCalculator", "Номер главной плоскости должен быть 1 или 2."
    End If
End Sub

' Считает габарит бетонного контура по нормали выбранной главной плоскости.
' Это единая точка для устойчивости и справочного вывода: арматура не
' расширяет внешний бетонный контур, если вызывающий код передает False.
Public Function PrincipalPlaneDepth(ByVal section As CSectionModel, ByVal planeIndex As Long, _
        Optional ByVal includeRebar As Boolean = False) As Double
    Dim normalX As Double
    Dim normalY As Double
    PrincipalPlaneNormalDirection planeIndex, normalX, normalY
    PrincipalPlaneDepth = ProjectionDepth(section, normalX, normalY, includeRebar)
End Function

' Считает ядровое расстояние для выбранной главной плоскости и стороны.
' plusSide=True означает положительное направление нормали плоскости.
Public Function PrincipalPlaneCoreDistance(ByVal section As CSectionModel, ByVal planeIndex As Long, _
        ByVal plusSide As Boolean, Optional ByVal includeRebar As Boolean = False) As Double
    Dim normalX As Double
    Dim normalY As Double
    PrincipalPlaneNormalDirection planeIndex, normalX, normalY
    If plusSide Then
        PrincipalPlaneCoreDistance = CoreDistanceAlong(section, normalX, normalY, includeRebar)
    Else
        PrincipalPlaneCoreDistance = CoreDistanceAlong(section, -normalX, -normalY, includeRebar)
    End If
End Function

' Считает ядровое расстояние для выбранной стороны сечения.
' direction задает сторону, куда направлен эксцентриситет. Само ядровое
' расстояние использует текущие накопленные A/I этого калькулятора: для
' приведенного сечения их надо заранее получить через CalculateTransformedByModuli.
' includeRebar влияет только на выбор крайней внешней точки c в формуле
' r = I/(A*c). Для устойчивости c берется до бетонного контура, поэтому туда
' передается False: арматура влияет на A/I, но не раздувает габарит h и плечо c.
Public Function CoreDistanceAlong(ByVal section As CSectionModel, ByVal directionX As Double, _
        ByVal directionY As Double, Optional ByVal includeRebar As Boolean = True) As Double
    If mArea <= GEOM_TOLERANCE Then Exit Function

    Dim normValue As Double
    normValue = Sqr(directionX * directionX + directionY * directionY)
    If normValue <= GEOM_TOLERANCE Then Err.Raise vbObjectError + 2219, "CSectionPropertiesCalculator", "Направление ядрового расстояния имеет нулевую длину."

    Dim nx As Double
    Dim ny As Double
    nx = directionX / normValue
    ny = directionY / normValue

    Dim minProjection As Double
    Dim maxProjection As Double
    CalculateProjection section, nx, ny, includeRebar, minProjection, maxProjection

    Dim centroidProjection As Double
    Dim extremeDistance As Double
    centroidProjection = nx * mCentroidX + ny * mCentroidY
    extremeDistance = maxProjection - centroidProjection
    If extremeDistance <= GEOM_TOLERANCE Then Exit Function

    CoreDistanceAlong = ProjectedInertiaAboutCentroid(nx, ny) / (mArea * extremeDistance)
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

' Проекция бетонного элемента для габаритов, ядровых расстояний и других
' геометрических проверок. CSectionModel является единственной точкой, где
' реальные A/I переводятся в прямоугольную оболочку элемента.
Private Function ConcreteHalfProjection(ByVal section As CSectionModel, ByVal index As Long, _
        ByVal axisX As Double, ByVal axisY As Double) As Double
    ConcreteHalfProjection = section.ConcreteBoundaryHalfProjection(index, axisX, axisY)
End Function

' Возвращает прямоугольник, которым бетонный КЭ участвует в пересечениях с
' лучом. Интерпретация Rectangle/Equivalent rectangle/Equivalent square
' находится в CSectionModel, чтобы схема, AutoCAD и расчетные габариты
' использовали одну и ту же оболочку.
Private Function ConcreteProjectionRectangle(ByVal section As CSectionModel, ByVal index As Long, _
        ByRef widthValue As Double, ByRef heightValue As Double, _
        ByRef angleValue As Double) As Boolean
    ConcreteProjectionRectangle = section.ConcreteBoundaryRectangle(index, widthValue, heightValue, angleValue)
End Function

' Находит интервал параметра t, на котором луч P + t*n проходит внутри
' расчетного прямоугольника бетонного элемента.
Private Function ConcreteRayInterval(ByVal section As CSectionModel, ByVal index As Long, _
        ByVal pointX As Double, ByVal pointY As Double, ByVal directionX As Double, _
        ByVal directionY As Double, ByRef tMin As Double, ByRef tMax As Double) As Boolean
    Dim widthValue As Double
    Dim heightValue As Double
    Dim angleValue As Double
    If Not ConcreteProjectionRectangle(section, index, widthValue, heightValue, angleValue) Then Exit Function

    ConcreteRayInterval = RectangleRayInterval(pointX, pointY, _
        section.ConcreteX(index), section.ConcreteY(index), _
        widthValue, heightValue, angleValue, directionX, directionY, tMin, tMax)
End Function

' Пересекает луч с повернутым прямоугольником через clipping по двум локальным
' координатам. t измеряется в мм, потому что direction нормирован.
Private Function RectangleRayInterval(ByVal pointX As Double, ByVal pointY As Double, _
        ByVal centerX As Double, ByVal centerY As Double, ByVal widthValue As Double, _
        ByVal heightValue As Double, ByVal angleValue As Double, _
        ByVal directionX As Double, ByVal directionY As Double, _
        ByRef tMin As Double, ByRef tMax As Double) As Boolean
    Dim ux As Double
    Dim uy As Double
    Dim vx As Double
    Dim vy As Double
    ux = Cos(angleValue)
    uy = Sin(angleValue)
    vx = -uy
    vy = ux

    Dim dx As Double
    Dim dy As Double
    dx = pointX - centerX
    dy = pointY - centerY

    tMin = -1E+100
    tMax = 1E+100
    If Not ClipRectangleRayCoordinate(dx * ux + dy * uy, _
            directionX * ux + directionY * uy, widthValue / 2#, tMin, tMax) Then Exit Function
    If Not ClipRectangleRayCoordinate(dx * vx + dy * vy, _
            directionX * vx + directionY * vy, heightValue / 2#, tMin, tMax) Then Exit Function

    RectangleRayInterval = (tMin <= tMax + PROJECTION_INTERVAL_TOLERANCE)
End Function

' Сужает интервал t одним условием |coord + t*dir| <= halfSize.
Private Function ClipRectangleRayCoordinate(ByVal coord As Double, ByVal dirCoord As Double, _
        ByVal halfSize As Double, ByRef tMin As Double, ByRef tMax As Double) As Boolean
    If halfSize <= GEOM_TOLERANCE Then Exit Function

    If Abs(dirCoord) <= GEOM_TOLERANCE Then
        ClipRectangleRayCoordinate = (Abs(coord) <= halfSize + PROJECTION_INTERVAL_TOLERANCE)
        Exit Function
    End If

    Dim t1 As Double
    Dim t2 As Double
    Dim tmp As Double
    t1 = (-halfSize - coord) / dirCoord
    t2 = (halfSize - coord) / dirCoord
    If t1 > t2 Then
        tmp = t1
        t1 = t2
        t2 = tmp
    End If

    tMin = GeomMax(tMin, t1)
    tMax = GeomMin(tMax, t2)
    ClipRectangleRayCoordinate = (tMin <= tMax + PROJECTION_INTERVAL_TOLERANCE)
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

    ' Для круга и почти изотропных сечений направление главных осей
    ' математически неопределено. Малый численный Ixy от дискретной сетки не
    ' должен поворачивать оси и раскладывать чистый Mx/My на две плоскости.
    If Abs(mIxyc) / inertiaRef <= PRINCIPAL_AXIS_ISOTROPIC_RELATIVE_TOLERANCE And _
            Abs(mIxc - mIyc) / inertiaRef <= PRINCIPAL_AXIS_ISOTROPIC_RELATIVE_TOLERANCE Then
        mPrincipalAngleRad = 0#
    Else
        mPrincipalAngleRad = 0.5 * Atn2(-2# * mIxyc, mIxc - mIyc)
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

' Возвращает atan2(y, x) в диапазоне VBA-арифметики. Нужен для главных осей:
' стандартной функции Atn хватает только на отношение y/x и она не знает
' квадрант исходного вектора.
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











