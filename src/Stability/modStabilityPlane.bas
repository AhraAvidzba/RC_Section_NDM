Attribute VB_Name = "modStabilityPlane"
Option Explicit

' ==========================================================================
' Общие данные и операции расчёта главных плоскостей устойчивости
' ==========================================================================

' Геометрию готовит CStabilityCalculator один раз; нормативные алгоритмы
' меняют только расчётные поля переданных плоскостей. Единицы: мм, Н, МПа.
Public Type TStabilityPlane
    Moment As Double ' Исходный момент в рассматриваемой главной плоскости, Н*мм.
    SustainedMoment As Double ' Постоянная и длительная часть момента, Н*мм.
    DesignMoment As Double ' Момент после учёта устойчивости, Н*мм.
    InertiaConcrete As Double ' Момент инерции бетона относительно расчётной оси, мм4.
    InertiaSteel As Double ' Момент инерции арматуры относительно расчётной оси, мм4.
    Radius As Double ' Радиус инерции приведённого сечения, мм.
    CoreDistancePositive As Double ' Ядровое расстояние в положительном направлении нормали, мм.
    CoreDistanceNegative As Double ' Ядровое расстояние в отрицательном направлении нормали, мм.
    Depth As Double ' Высота сечения в рассматриваемой плоскости, мм.
    GeometricLength As Double ' Геометрическая длина элемента, мм.
    EffectiveLength As Double ' Расчётная длина элемента, мм.
    Slenderness As Double ' Гибкость как отношение расчётной длины к радиусу инерции.
    NormalX As Double ' Компонента X положительной нормали к оси изгиба.
    NormalY As Double ' Компонента Y положительной нормали к оси изгиба.
    CentroidX As Double ' Координата X центра приведённого сечения, мм.
    CentroidY As Double ' Координата Y центра приведённого сечения, мм.
    ConcreteMinProjection As Double ' Наименьшая координата бетонной границы вдоль нормали, мм.
    ConcreteMaxProjection As Double ' Наибольшая координата бетонной границы вдоль нормали, мм.
    AccidentalEccentricity As Double ' Случайный эксцентриситет, мм.
    Eccentricity As Double ' Принятый расчётный эксцентриситет, мм.
    CoreDistance As Double ' Ядровое расстояние со стороны рассматриваемого эксцентриситета, мм.
    PhiL As Double ' Коэффициент влияния длительности нагрузки.
    PhiLTable As Double ' Табличный коэффициент phil по СП 35.
    PhiM As Double ' Табличный коэффициент phim по СП 35.
    PhiValue As Double ' Итоговый коэффициент phi табличной проверки СП 35.
    PhiP As Double ' Коэффициент phip по СП 35.
    DeltaValue As Double ' Принятое относительное значение эксцентриситета для выбранной нормативной ветви.
    Kb As Double ' Коэффициент вклада бетона в жёсткость по СП 63.
    Ks As Double ' Коэффициент вклада арматуры в жёсткость по СП 63.
    Eta As Double ' Коэффициент увеличения момента из-за прогиба.
    StiffnessD As Double ' Условная изгибная жёсткость D, Н*мм2.
    CriticalForce As Double ' Контрольная сила выбранной нормативной проверки, Н.
    NcrForce As Double ' Критическая сила Ncr для расчёта увеличения момента, Н.
    NultimateForce As Double ' Предельная сжимающая сила табличной проверки СП 35, Н.
    ReserveFactor As Double ' Коэффициент запаса в рассматриваемой плоскости.
    CheckPassed As Boolean ' Выполнена ли проверка в рассматриваемой плоскости.
    Applicable As Boolean ' Требуется ли проверка для рассматриваемой плоскости.
    BranchName As String ' Название применённой нормативной расчётной ветви.
End Type

Public Const STABILITY_CODE_SP63 As String = "SP63" ' Расчет по СП 63.
Public Const STABILITY_CODE_SP35 As String = "SP35" ' Расчет по СП 35.
Public Const STABILITY_SYSTEM_DETERMINATE As String = "Determinate" ' Статически определимая схема.
Public Const STABILITY_SYSTEM_INDETERMINATE As String = "Indeterminate" ' Статически неопределимая схема.
Public Const STABILITY_PHIL_AUTO As String = "Auto" ' phi_l считается по отношению длительного момента к полному.
Public Const STABILITY_PHIL_2 As String = "PhiL2" ' Консервативно принимается phi_l = 2.
Public Const STABILITY_ACC_AUTO_L As String = "AutoWithL" ' Случайный эксцентриситет по геометрической длине L.
Public Const STABILITY_ACC_AUTO_MUL As String = "AutoWithMuL" ' Случайный эксцентриситет по расчетной длине mu*L.
Public Const STABILITY_ACC_USER As String = "User" ' Пользователь задает модуль случайного эксцентриситета.
Public Const STABILITY_ACC_PLANES_ONLY_MOMENT As String = "OnlyMomentPlane" ' Случайный эксцентриситет только в плоскости исходного момента.
Public Const STABILITY_ACC_PLANES_BOTH As String = "BothPlanes" ' Случайный эксцентриситет в обеих главных плоскостях.
Public Const STABILITY_SMALL_FORCE As Double = 0.000001 ' Малый допуск по силе, Н.
Public Const STABILITY_SMALL_VALUE As Double = 0.000000001 ' Общий малый допуск для безразмерных величин.
Public Const STABILITY_PASS_TOLERANCE As Double = 0.000000001 ' Малый допуск для граничного PASS.

' Возвращает край бетонного сечения на стороне, противоположной направлению
' сжатой равнодействующей. Для СП 35 это грань, относительно которой считается
' момент в коэффициенте длительности.
Public Function StabilityConcreteFaceProjectionForMoment(ByRef plane As TStabilityPlane) As Double
    If plane.Eccentricity >= 0# Then
        StabilityConcreteFaceProjectionForMoment = plane.ConcreteMinProjection
    Else
        StabilityConcreteFaceProjectionForMoment = plane.ConcreteMaxProjection
    End If
End Function

' Пересчитывает момент из центра приведенного сечения к точке с проекцией p.
' Сжатая сила nCompression передается положительным числом.
Public Function StabilityMomentAboutProjection(ByVal momentValue As Double, ByVal nCompression As Double, _
        ByVal projectionFromCentroid As Double) As Double
    StabilityMomentAboutProjection = momentValue - nCompression * projectionFromCentroid
End Function

' Проверяет положительность обязательных геометрических величин главной плоскости.
' При ошибке записывает фактические значения, в том числе
' для размеров элементов, восстановленных при импорте AutoCAD.
Public Sub StabilityEnsurePositivePlaneGeometry(ByRef plane As TStabilityPlane, _
        ByVal codeText As String, ByVal planeLabel As String)
    If plane.EffectiveLength > 0# And plane.InertiaConcrete > 0# And _
            plane.Depth > 0# And plane.Radius > 0# Then Exit Sub

    StabilityRaiseInput "Для расчета устойчивости " & codeText & " в главной плоскости " & _
        planeLabel & " нужны положительные l0, h, I и i. Сейчас: l0=" & _
        StabilityStabilityNumber(plane.EffectiveLength) & " мм; h=" & _
        StabilityStabilityNumber(plane.Depth) & " мм; I=" & _
        StabilityStabilityNumber(plane.InertiaConcrete) & " мм4; i=" & _
        StabilityStabilityNumber(plane.Radius) & " мм."
End Sub

' Форматирует диагностические числа без длинных хвостов, чтобы popup с причиной
' InputErr оставался читаемым для пользователя.
Public Function StabilityStabilityNumber(ByVal value As Double) As String
    StabilityStabilityNumber = Replace$(Format$(value, "0.############"), ",", ".")
End Function

' Определяет знак случайного эксцентриситета: при ненулевом моменте добавляет
' его в ту же сторону, при нулевом моменте использует пользовательские +1/-1.
Public Function StabilitySignForEccentricity(ByVal momentValue As Double, ByVal userSign As Double) As Double
    If Abs(momentValue) > STABILITY_SMALL_VALUE Then
        StabilitySignForEccentricity = Sgn(momentValue)
    ElseIf userSign < 0# Then
        StabilitySignForEccentricity = -1#
    Else
        StabilitySignForEccentricity = 1#
    End If
End Function

' Начальный эксцентриситет по типу расчетной схемы. Для статически определимой
' схемы складываются расчетный момент и случайный эксцентриситет; для
' неопределимой берется не меньше случайного эксцентриситета.
Public Function StabilityInitialEccentricity(ByVal momentValue As Double, ByVal nCompression As Double, _
        ByVal accidentalEccentricity As Double, ByVal systemType As String) As Double
    Dim staticEccentricity As Double
    If nCompression > STABILITY_SMALL_FORCE Then staticEccentricity = momentValue / nCompression

    If UCase$(Trim$(systemType)) = UCase$(STABILITY_SYSTEM_INDETERMINATE) Then
        StabilityInitialEccentricity = StabilitySignForEccentricity(momentValue, accidentalEccentricity) * _
            GeomMax(Abs(staticEccentricity), Abs(accidentalEccentricity))
    Else
        StabilityInitialEccentricity = staticEccentricity + accidentalEccentricity
    End If
End Function

' Возвращает phi_l. В Auto используется отношение длительного момента к
' полному, но не больше 2. При PhiL2 сразу принимается консервативное значение 2.
Public Function StabilityPhiLByMode(ByVal phiLMode As String, ByVal totalMoment As Double, ByVal sustainedMoment As Double) As Double
    StabilityPhiLByMode = StabilityPhiLFromMomentRatio(phiLMode, totalMoment, sustainedMoment)
End Function

' Общая реализация формулы phi_l = 1 + Ml/M с верхним ограничением 2.
' Отношение берется со знаком. Нижняя граница 1 оставлена как защитное
' инженерное ограничение, чтобы длительная часть не повышала расчетную жесткость.
Public Function StabilityPhiLFromMomentRatio(ByVal phiLMode As String, ByVal totalMoment As Double, _
        ByVal sustainedMoment As Double) As Double
    Dim normalizedMode As String
    normalizedMode = UCase$(Trim$(phiLMode))
    If normalizedMode = UCase$(STABILITY_PHIL_2) Then
        StabilityPhiLFromMomentRatio = 2#
    ElseIf normalizedMode <> UCase$(STABILITY_PHIL_AUTO) Then
        StabilityRaiseInput "На листе Config в строке Stability.PhiLMode выберите Auto или PhiL2."
    ElseIf Abs(totalMoment) <= STABILITY_SMALL_VALUE Then
        If Abs(sustainedMoment) <= STABILITY_SMALL_VALUE Then StabilityPhiLFromMomentRatio = 1# Else StabilityPhiLFromMomentRatio = 2#
    Else
        StabilityPhiLFromMomentRatio = 1# + sustainedMoment / totalMoment
        If StabilityPhiLFromMomentRatio > 2# Then StabilityPhiLFromMomentRatio = 2#
        If StabilityPhiLFromMomentRatio < 1# Then StabilityPhiLFromMomentRatio = 1#
    End If
End Function

' Коэффициент eta по критической силе. Если N >= Ncr, eta не имеет физически
' корректного положительного значения, поэтому возвращается 0 и статус станет FAIL.
Public Function StabilityEtaByCriticalForce(ByVal criticalForce As Double, ByVal nCompression As Double) As Double
    If criticalForce <= nCompression + STABILITY_SMALL_FORCE Then Exit Function
    StabilityEtaByCriticalForce = 1# / (1# - nCompression / criticalForce)
End Function

' Запас как отношение критической/предельной силы к действующей сжимающей.
Public Function StabilityReserveByCriticalForce(ByVal criticalForce As Double, ByVal nCompression As Double) As Double
    If criticalForce > 0# And nCompression > STABILITY_SMALL_FORCE Then StabilityReserveByCriticalForce = criticalForce / nCompression
End Function

' Расчетный момент после учета eta.
Public Function StabilityDesignMomentWithEta(ByVal nCompression As Double, ByVal eccentricity As Double, ByVal etaValue As Double) As Double
    If etaValue > 0# Then
        StabilityDesignMomentWithEta = nCompression * eccentricity * etaValue
    Else
        StabilityDesignMomentWithEta = nCompression * eccentricity
    End If
End Function

' Передаёт общую причину фасаду: его обработчик сохраняет InputErr и текст.
Public Sub StabilityRaiseInput(ByVal message As String)
    Err.Raise vbObjectError + 7600, "CStabilityCalculator", message
End Sub
