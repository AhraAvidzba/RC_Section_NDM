VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CStabilityCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Расчет продольного изгиба и устойчивости
' ==========================================================================
' Класс является изолированным фильтром усилий перед существующим НДМ.
' Он не запускает CSectionSolver, не читает Excel и не меняет материалы:
' вызывающий код передает готовую CSectionModel, выбранный набор характеристик
' материалов профиля и усилия LC. На выходе возвращаются N/Mx/My с учетом
' продольного изгиба, а также расчетные величины для Summary.
'
' Геометрические характеристики класс берет через CSectionPropertiesCalculator.
' Поэтому расчет высоты по повернутой оси, локальных инерций и радиусов
' арматуры не дублируется внутри методики устойчивости.

Private Const CODE_SP63 As String = "SP63" ' Расчет по СП 63.
Private Const CODE_SP35 As String = "SP35" ' Расчет по СП 35.
Private Const SYSTEM_DETERMINATE As String = "Determinate" ' Статически определимая схема.
Private Const SYSTEM_INDETERMINATE As String = "Indeterminate" ' Статически неопределимая схема.
Private Const PHIL_AUTO As String = "Auto" ' phi_l считается по отношению длительного момента к полному.
Private Const PHIL_2 As String = "PhiL2" ' Консервативно принимается phi_l = 2.
Private Const SMALL_FORCE As Double = 0.000001 ' Малый допуск по силе, Н.
Private Const SMALL_VALUE As Double = 0.000000001 ' Общий малый допуск для безразмерных величин.

Private Type TPlaneResult
    Moment As Double
    SustainedMoment As Double
    DesignMoment As Double
    InertiaConcrete As Double
    InertiaSteel As Double
    Radius As Double
    Depth As Double
    EffectiveLength As Double
    AccidentalEccentricity As Double
    Eccentricity As Double
    PhiL As Double
    Eta As Double
    StiffnessD As Double
    CriticalForce As Double
    NcrForce As Double
    NultimateForce As Double
    ReserveFactor As Double
    CheckPassed As Boolean
    Applicable As Boolean
End Type

Private mCode As String ' Нормативная ветвь, выбранная настройкой Stability.Code.
Private mBranch As String ' Внутренняя трассировка расчетной ветви для execution report и диагностики.
Private mApplicable As Boolean ' False, если сжатия нет и устойчивость не требуется.
Private mInputValid As Boolean ' False при нехватке обязательных исходных данных.
Private mCheckPassed As Boolean ' Итог проверки по выбранной ветви.
Private mStopReason As String ' Подробная внутренняя причина, если расчет не выполнен.
Private mDesignN As Double ' Продольная сила после фильтра; знак сохраняется как во внутреннем ядре.
Private mDesignMx As Double ' Mx после учета случайного эксцентриситета и eta.
Private mDesignMy As Double ' My после учета случайного эксцентриситета и eta.
Private mCriticalForce As Double ' Минимальная контрольная сила: Ncr для eta-ветвей или Nult,stab для табличной ветви.
Private mReserveFactor As Double ' Минимальный запас по устойчивости.
Private mNcr1 As Double ' Критическая сила Ncr первой главной плоскости; 0, если ветвь считает Nult вместо Ncr.
Private mNcr2 As Double ' Критическая сила Ncr второй главной плоскости; 0, если ветвь считает Nult вместо Ncr.
Private mNultimate1 As Double ' Предельная сжимающая сила Nult,stab первой главной плоскости для табличной ветви СП 35.
Private mNultimate2 As Double ' Предельная сжимающая сила Nult,stab второй главной плоскости для табличной ветви СП 35.
Private mReserve1 As Double ' Запас по первой главной плоскости.
Private mReserve2 As Double ' Запас по второй главной плоскости.
Private mEta1 As Double ' Коэффициент eta первой плоскости.
Private mEta2 As Double ' Коэффициент eta второй плоскости.
Private mPhiL1 As Double ' Коэффициент длительности первой плоскости.
Private mPhiL2 As Double ' Коэффициент длительности второй плоскости.
Private mEccentricity1 As Double ' Расчетный эксцентриситет первой плоскости, мм.
Private mEccentricity2 As Double ' Расчетный эксцентриситет второй плоскости, мм.
Private mAccidentalEccentricity1 As Double ' Случайный эксцентриситет первой плоскости, мм.
Private mAccidentalEccentricity2 As Double ' Случайный эксцентриситет второй плоскости, мм.
Private mEffectiveLength1 As Double ' l0 первой плоскости, мм.
Private mEffectiveLength2 As Double ' l0 второй плоскости, мм.
Private mStiffnessD1 As Double ' Условная жесткость D первой плоскости, Н*мм2.
Private mStiffnessD2 As Double ' Условная жесткость D второй плоскости, Н*мм2.

' Выполняет расчет по выбранному нормативному блоку и возвращает расчетные
' усилия для дальнейшего НДМ. Если продольная сила не сжимает сечение, фильтр
' считается неприменимым и усилия передаются дальше без изменения.
Public Sub Calculate(ByVal section As CSectionModel, ByVal materialProvider As CMaterialModelProvider, _
        ByVal valueSet As EMaterialValueSet, ByVal codeText As String, _
        ByVal nValue As Double, ByVal mxValue As Double, ByVal myValue As Double, _
        ByVal sustainedN As Double, ByVal sustainedMx As Double, ByVal sustainedMy As Double, _
        ByVal elementLength As Double, ByVal mu1 As Double, ByVal mu2 As Double, _
        ByVal systemType As String, ByVal zeroMomentSign1 As Double, ByVal zeroMomentSign2 As Double, _
        ByVal phiLMode As String, ByVal sp63Ks As Double, ByVal sp63DeltaEMin As Double, _
        ByVal sp63DeltaEMax As Double, ByVal sp35PhiP As Double, _
        ByVal sp35NOverNcrLimit As Double, ByVal sp35Table721 As Variant)
    On Error GoTo Failed
    Clear
    mCode = NormalizeCode(codeText)
    mDesignN = nValue
    mDesignMx = mxValue
    mDesignMy = myValue

    If section Is Nothing Then RaiseInput "Модель сечения не передана."
    If materialProvider Is Nothing Then RaiseInput "Material provider не передан."
    If elementLength <= 0# Then RaiseInput "Stability.ElementLength должен быть положительным."
    If mCode <> CODE_SP63 And mCode <> CODE_SP35 Then RaiseInput "Stability.Code должен быть SP63 или SP35."

    Dim nCompression As Double
    nCompression = CompressionForce(nValue)
    If nCompression <= SMALL_FORCE Then
        mApplicable = False
        mCheckPassed = True
        mBranch = "NoCompression"
        Exit Sub
    End If

    Dim props As CSectionPropertiesCalculator
    Set props = New CSectionPropertiesCalculator
    props.CalculateConcrete section

    Dim plane1 As TPlaneResult
    Dim plane2 As TPlaneResult
    PreparePlanes section, props, mxValue, myValue, sustainedMx, sustainedMy, _
        elementLength, mu1, mu2, systemType, zeroMomentSign1, zeroMomentSign2, plane1, plane2

    If mCode = CODE_SP63 Then
        CalculateSP63 section, materialProvider, valueSet, nCompression, systemType, phiLMode, sp63Ks, _
            sp63DeltaEMin, sp63DeltaEMax, plane1, plane2
    Else
        CalculateSP35 section, materialProvider, valueSet, props.Area, props.RebarAreaTotal(section), _
            nCompression, CompressionForce(sustainedN), systemType, phiLMode, _
            sp35PhiP, sp35NOverNcrLimit, sp35Table721, plane1, plane2
    End If

    StorePlaneResults plane1, plane2
    TransformDesignMoments props.PrincipalAngleRad, plane1.DesignMoment, plane2.DesignMoment
    Exit Sub

Failed:
    mInputValid = False
    mApplicable = False
    mCheckPassed = False
    If Len(mStopReason) = 0 Then mStopReason = Err.Description
End Sub

' Сбрасывает результаты перед новым расчетом.
Private Sub Clear()
    mCode = vbNullString
    mBranch = vbNullString
    mApplicable = True
    mInputValid = True
    mCheckPassed = False
    mStopReason = vbNullString
    mDesignN = 0#: mDesignMx = 0#: mDesignMy = 0#
    mCriticalForce = 0#: mReserveFactor = 0#
    mNcr1 = 0#: mNcr2 = 0#
    mNultimate1 = 0#: mNultimate2 = 0#
    mReserve1 = 0#: mReserve2 = 0#
    mEta1 = 0#: mEta2 = 0#
    mPhiL1 = 0#: mPhiL2 = 0#
    mEccentricity1 = 0#: mEccentricity2 = 0#
    mAccidentalEccentricity1 = 0#: mAccidentalEccentricity2 = 0#
    mEffectiveLength1 = 0#: mEffectiveLength2 = 0#
    mStiffnessD1 = 0#: mStiffnessD2 = 0#
End Sub

' Подготавливает две главные расчетные плоскости: моменты переводятся в
' главные центральные оси бетонного сечения, а высоты берутся из общего
' геометрического сервиса с учетом эквивалентного радиуса арматуры.
Private Sub PreparePlanes(ByVal section As CSectionModel, ByVal props As CSectionPropertiesCalculator, _
        ByVal mxValue As Double, ByVal myValue As Double, ByVal sustainedMx As Double, _
        ByVal sustainedMy As Double, ByVal elementLength As Double, ByVal mu1 As Double, _
        ByVal mu2 As Double, ByVal systemType As String, ByVal zeroMomentSign1 As Double, _
        ByVal zeroMomentSign2 As Double, ByRef plane1 As TPlaneResult, ByRef plane2 As TPlaneResult)
    Dim axis1X As Double
    Dim axis1Y As Double
    Dim axis2X As Double
    Dim axis2Y As Double
    axis1X = Cos(props.PrincipalAngleRad)
    axis1Y = Sin(props.PrincipalAngleRad)
    axis2X = -axis1Y
    axis2Y = axis1X

    plane1.Moment = mxValue * axis1X + myValue * axis1Y
    plane2.Moment = mxValue * axis2X + myValue * axis2Y
    plane1.SustainedMoment = sustainedMx * axis1X + sustainedMy * axis1Y
    plane2.SustainedMoment = sustainedMx * axis2X + sustainedMy * axis2Y
    plane1.InertiaConcrete = props.PrincipalI1
    plane2.InertiaConcrete = props.PrincipalI2
    plane1.InertiaSteel = props.RebarInertiaAboutAxis(section, axis1X, axis1Y, props.CentroidX, props.CentroidY)
    plane2.InertiaSteel = props.RebarInertiaAboutAxis(section, axis2X, axis2Y, props.CentroidX, props.CentroidY)
    plane1.Radius = props.PrincipalRadius1
    plane2.Radius = props.PrincipalRadius2
    plane1.Depth = props.ProjectionDepth(section, axis2X, axis2Y, True)
    plane2.Depth = props.ProjectionDepth(section, axis1X, axis1Y, True)
    plane1.EffectiveLength = elementLength * PositiveOrOne(mu1)
    plane2.EffectiveLength = elementLength * PositiveOrOne(mu2)
    plane1.Applicable = True
    plane2.Applicable = True

    ' Если момент уже задан в одной из главных плоскостей, случайный
    ' эксцентриситет добавляется только к действующему направлению момента.
    ' Для чистого сжатия без моментов остаются пользовательские направления +1/-1.
    Dim hasMomentVector As Boolean
    hasMomentVector = (Abs(plane1.Moment) > SMALL_VALUE Or Abs(plane2.Moment) > SMALL_VALUE)
    FillAccidentalAndInitialEccentricity plane1, zeroMomentSign1, systemType, _
        (Not hasMomentVector Or Abs(plane1.Moment) > SMALL_VALUE)
    FillAccidentalAndInitialEccentricity plane2, zeroMomentSign2, systemType, _
        (Not hasMomentVector Or Abs(plane2.Moment) > SMALL_VALUE)
End Sub

' Назначает случайный и начальный эксцентриситет для одной главной плоскости.
Private Sub FillAccidentalAndInitialEccentricity(ByRef plane As TPlaneResult, _
        ByVal zeroMomentSign As Double, ByVal systemType As String, ByVal allowZeroMomentAccidental As Boolean)
    If Abs(plane.Moment) <= SMALL_VALUE And Not allowZeroMomentAccidental Then
        plane.AccidentalEccentricity = 0#
        plane.Eccentricity = 0#
        Exit Sub
    End If

    Dim signValue As Double
    signValue = SignForEccentricity(plane.Moment, zeroMomentSign)
    plane.AccidentalEccentricity = signValue * AccidentalEccentricity(plane.EffectiveLength, plane.Depth)
    plane.Eccentricity = InitialEccentricity(plane.Moment, 1#, plane.AccidentalEccentricity, systemType)
End Sub

' ----------------------------- СП 63 --------------------------------------

' Расчет по СП 63: eta = 1/(1-N/Ncr), Ncr = pi^2*D/l0^2,
' D = kb*Eb*I + ks*Es*Is. Ветвь работает для произвольного дискретного
' сечения, потому что I, Is и h уже получены из CSectionModel.
Private Sub CalculateSP63(ByVal section As CSectionModel, ByVal materialProvider As CMaterialModelProvider, _
        ByVal valueSet As EMaterialValueSet, ByVal nCompression As Double, _
        ByVal systemType As String, _
        ByVal phiLMode As String, ByVal ksValue As Double, ByVal deltaEMin As Double, _
        ByVal deltaEMax As Double, ByRef plane1 As TPlaneResult, ByRef plane2 As TPlaneResult)
    Dim eb As Double
    Dim es As Double
    eb = materialProvider.ConcreteElasticModulusFromValueSet(valueSet)
    es = materialProvider.SteelElasticModulusFromValueSet(valueSet)
    If eb <= 0# Or es <= 0# Then RaiseInput "Для расчета устойчивости СП 63 нужны положительные Eb и Es."
    If ksValue <= 0# Then ksValue = 0.7

    CalculateSP63Plane nCompression, eb, es, ksValue, systemType, phiLMode, deltaEMin, deltaEMax, plane1
    CalculateSP63Plane nCompression, eb, es, ksValue, systemType, phiLMode, deltaEMin, deltaEMax, plane2
    mBranch = CODE_SP63
End Sub

' Считает одну главную плоскость по формулам СП 63.
Private Sub CalculateSP63Plane(ByVal nCompression As Double, ByVal eb As Double, ByVal es As Double, _
        ByVal ksValue As Double, ByVal systemType As String, ByVal phiLMode As String, ByVal deltaEMin As Double, _
        ByVal deltaEMax As Double, ByRef plane As TPlaneResult)
    If plane.EffectiveLength <= 0# Or plane.InertiaConcrete <= 0# Or plane.Depth <= 0# Then
        plane.Applicable = False
        Exit Sub
    End If

    plane.Eccentricity = InitialEccentricity(plane.Moment, nCompression, _
        plane.AccidentalEccentricity, systemType)
    plane.PhiL = PhiLByMode(phiLMode, plane.Moment, plane.SustainedMoment)

    Dim deltaE As Double
    deltaE = Abs(plane.Eccentricity) / plane.Depth
    If deltaEMin > 0# Then deltaE = GeomMax(deltaE, deltaEMin)
    If deltaEMax > 0# Then deltaE = GeomMin(deltaE, deltaEMax)

    Dim kb As Double
    kb = 0.15 / (plane.PhiL * (0.3 + deltaE))
    plane.StiffnessD = kb * eb * plane.InertiaConcrete + ksValue * es * plane.InertiaSteel
    plane.NcrForce = GEOM_PI * GEOM_PI * plane.StiffnessD / (plane.EffectiveLength * plane.EffectiveLength)
    plane.CriticalForce = plane.NcrForce
    plane.ReserveFactor = ReserveByCriticalForce(plane.NcrForce, nCompression)
    plane.Eta = EtaByCriticalForce(plane.NcrForce, nCompression)
    plane.DesignMoment = DesignMomentWithEta(nCompression, plane.Eccentricity, plane.Eta)
    plane.CheckPassed = (plane.ReserveFactor > 1# And plane.Eta > 0#)
End Sub

' ----------------------------- СП 35 --------------------------------------

' Расчет по СП 35. Для малых эксцентриситетов используется табличная ветвь
' phi по таблице 7.21, для больших - ветвь eta/Ncr по п. 7.54.
Private Sub CalculateSP35(ByVal section As CSectionModel, ByVal materialProvider As CMaterialModelProvider, _
        ByVal valueSet As EMaterialValueSet, ByVal concreteArea As Double, ByVal totalSteelArea As Double, _
        ByVal nCompression As Double, ByVal sustainedCompression As Double, _
        ByVal systemType As String, ByVal phiLMode As String, ByVal phiP As Double, _
        ByVal nOverNcrLimit As Double, ByVal table721 As Variant, _
        ByRef plane1 As TPlaneResult, ByRef plane2 As TPlaneResult)
    Dim eb As Double
    Dim es As Double
    Dim rb As Double
    Dim rsc As Double
    eb = materialProvider.ConcreteElasticModulusFromValueSet(valueSet)
    es = materialProvider.SteelElasticModulusFromValueSet(valueSet)
    rb = materialProvider.ConcreteCompressionResistanceFromValueSet(valueSet)
    rsc = materialProvider.SteelCompressionResistanceFromValueSet(valueSet)
    If eb <= 0# Or es <= 0# Or rb <= 0# Or rsc <= 0# Then RaiseInput "Для расчета устойчивости СП 35 нужны положительные Eb, Es, Rb и Rsc."
    If phiP <= 0# Then phiP = 1#
    If nOverNcrLimit <= 0# Then nOverNcrLimit = 0.7

    If sustainedCompression < 0# Then sustainedCompression = 0#
    If sustainedCompression > nCompression Then sustainedCompression = nCompression

    CalculateSP35Plane concreteArea, totalSteelArea, nCompression, sustainedCompression, eb, es, rb, rsc, systemType, phiLMode, phiP, nOverNcrLimit, table721, plane1
    CalculateSP35Plane concreteArea, totalSteelArea, nCompression, sustainedCompression, eb, es, rb, rsc, systemType, phiLMode, phiP, nOverNcrLimit, table721, plane2
End Sub

' Считает одну главную плоскость по СП 35: табличная ветвь при e/r <= 1,
' иначе eta/Ncr. Таблица 7.21 передается с Config и не зашита в код.
Private Sub CalculateSP35Plane(ByVal concreteArea As Double, ByVal totalSteelArea As Double, _
        ByVal nCompression As Double, ByVal sustainedCompression As Double, ByVal eb As Double, ByVal es As Double, _
        ByVal rb As Double, ByVal rsc As Double, ByVal systemType As String, ByVal phiLMode As String, _
        ByVal phiP As Double, ByVal nOverNcrLimit As Double, ByVal table721 As Variant, _
        ByRef plane As TPlaneResult)
    If plane.EffectiveLength <= 0# Or plane.InertiaConcrete <= 0# Or plane.Depth <= 0# Or plane.Radius <= 0# Then
        plane.Applicable = False
        Exit Sub
    End If

    ' Нулевой accidental из PreparePlanes означает, что эта плоскость не должна
    ' получать дополнительный момент даже в ветке СП 35 с заменой на l0/400.
    If Abs(plane.Moment) <= SMALL_VALUE And Abs(plane.AccidentalEccentricity) <= SMALL_VALUE Then
        plane.AccidentalEccentricity = 0#
    Else
        plane.AccidentalEccentricity = SignForEccentricity(plane.Moment, plane.AccidentalEccentricity) * plane.EffectiveLength / 400#
    End If
    plane.Eccentricity = InitialEccentricity(plane.Moment, nCompression, _
        plane.AccidentalEccentricity, systemType)

    Dim q As Double
    q = Abs(plane.Eccentricity) / plane.Radius
    If q <= 1# Then
        CalculateSP35TablePlane concreteArea, totalSteelArea, nCompression, sustainedCompression, rb, rsc, phiLMode, table721, q, plane
    Else
        CalculateSP35EtaPlane nCompression, eb, es, rb, phiLMode, phiP, nOverNcrLimit, plane
    End If
End Sub

' Табличная ветвь СП 35 для e/r <= 1.
Private Sub CalculateSP35TablePlane(ByVal concreteArea As Double, ByVal totalSteelArea As Double, _
        ByVal nCompression As Double, ByVal sustainedCompression As Double, ByVal rb As Double, ByVal rsc As Double, _
        ByVal phiLMode As String, ByVal table721 As Variant, ByVal q As Double, _
        ByRef plane As TPlaneResult)
    Dim slenderness As Double
    slenderness = plane.EffectiveLength / plane.Radius

    Dim phiM As Double
    Dim phiLTable As Double
    InterpolateSP35Table table721, slenderness, q, phiM, phiLTable
    If UCase$(Trim$(phiLMode)) = UCase$(PHIL_2) Then
        plane.PhiL = 2#
    Else
        plane.PhiL = phiLTable
    End If

    Dim nUltimate As Double
    Dim phiValue As Double
    phiValue = PhiForSustainedCompression(phiM, plane.PhiL, nCompression, sustainedCompression)
    nUltimate = phiValue * (rb * concreteArea + rsc * totalSteelArea)

    plane.NultimateForce = nUltimate
    plane.CriticalForce = plane.NultimateForce
    plane.ReserveFactor = ReserveByCriticalForce(nUltimate, nCompression)
    plane.Eta = 1#
    plane.StiffnessD = 0#
    plane.DesignMoment = nCompression * plane.Eccentricity
    plane.CheckPassed = (plane.ReserveFactor > 1#)
    mBranch = "SP35-table"
End Sub

' Ветвь СП 35 через eta и Ncr для e/r > 1.
Private Sub CalculateSP35EtaPlane(ByVal nCompression As Double, ByVal eb As Double, ByVal es As Double, _
        ByVal rb As Double, ByVal phiLMode As String, ByVal phiP As Double, _
        ByVal nOverNcrLimit As Double, ByRef plane As TPlaneResult)
    plane.PhiL = PhiLSP35Eta(phiLMode, plane.Moment, plane.SustainedMoment, plane.Eccentricity, plane.Depth)

    Dim deltaValue As Double
    deltaValue = Abs(plane.Eccentricity) / plane.Depth
    Dim deltaMin As Double
    deltaMin = 0.5 - 0.01 * (plane.EffectiveLength / plane.Depth) - 0.01 * rb
    deltaValue = GeomMax(deltaValue, deltaMin)

    Dim phiPValue As Double
    phiPValue = phiP
    Dim stiffnessFactor As Double
    stiffnessFactor = 0.11 / (0.1 + deltaValue / phiPValue) + 0.1
    plane.StiffnessD = 6.4 * eb * (plane.InertiaConcrete / plane.PhiL * stiffnessFactor + _
        (es / eb) * plane.InertiaSteel)
    plane.NcrForce = plane.StiffnessD / (plane.EffectiveLength * plane.EffectiveLength)
    plane.CriticalForce = plane.NcrForce
    plane.ReserveFactor = ReserveByCriticalForce(plane.NcrForce, nCompression)
    plane.Eta = EtaByCriticalForce(plane.NcrForce, nCompression)
    plane.DesignMoment = DesignMomentWithEta(nCompression, plane.Eccentricity, plane.Eta)
    plane.CheckPassed = (plane.ReserveFactor > 1# And plane.Eta > 0# And _
        nCompression / plane.NcrForce <= nOverNcrLimit)
    mBranch = "SP35-eta"
End Sub

' Для табличной ветви СП 35 учитывает долю постоянной нагрузки в полном
' сжимающем усилии по формуле (7.15). Если длительная часть равна нулю,
' остается кратковременный коэффициент phi_m; если вся сила длительная,
' результат переходит к phi_l.
Private Function PhiForSustainedCompression(ByVal phiM As Double, ByVal phiL As Double, _
        ByVal nCompression As Double, ByVal sustainedCompression As Double) As Double
    If phiM <= 0# Then Exit Function
    If phiL <= 0# Then
        PhiForSustainedCompression = phiM
        Exit Function
    End If
    If nCompression <= SMALL_FORCE Then
        PhiForSustainedCompression = phiM
        Exit Function
    End If

    Dim sustainedRatio As Double
    sustainedRatio = sustainedCompression / nCompression
    If sustainedRatio < 0# Then sustainedRatio = 0#
    If sustainedRatio > 1# Then sustainedRatio = 1#

    Dim shortRatio As Double
    shortRatio = 1# - sustainedRatio

    Dim denominator As Double
    denominator = sustainedRatio * (phiM / phiL) + shortRatio
    If denominator > SMALL_VALUE Then
        PhiForSustainedCompression = phiM / denominator
    Else
        PhiForSustainedCompression = phiM
    End If
End Function

' Возвращает phi_l для ветви eta/Ncr по СП 35. В Auto обычно используется
' отношение длительного момента к полному, но при противоположных знаках
' полного и длительного моментов применяются специальные правила п. 7.54.
Private Function PhiLSP35Eta(ByVal phiLMode As String, ByVal totalMoment As Double, _
        ByVal sustainedMoment As Double, ByVal eccentricity As Double, ByVal depthValue As Double) As Double
    If UCase$(Trim$(phiLMode)) = UCase$(PHIL_2) Then
        PhiLSP35Eta = 2#
        Exit Function
    End If

    If Abs(totalMoment) > SMALL_VALUE And Abs(sustainedMoment) > SMALL_VALUE Then
        If Sgn(totalMoment) <> Sgn(sustainedMoment) And depthValue > SMALL_VALUE Then
            If Abs(eccentricity) >= 0.1 * depthValue Then
                PhiLSP35Eta = 1#
            Else
                PhiLSP35Eta = 1.05
            End If
            Exit Function
        End If
    End If

    PhiLSP35Eta = PhiLByMode(phiLMode, totalMoment, sustainedMoment)
End Function

' -------------------------- Общие методы ----------------------------------

' Собирает итог по двум главным плоскостям и выбирает минимальный запас.
Private Sub StorePlaneResults(ByRef plane1 As TPlaneResult, ByRef plane2 As TPlaneResult)
    mApplicable = (plane1.Applicable Or plane2.Applicable)
    If Not mApplicable Then
        mCheckPassed = True
        Exit Sub
    End If

    mNcr1 = plane1.NcrForce
    mNcr2 = plane2.NcrForce
    mNultimate1 = plane1.NultimateForce
    mNultimate2 = plane2.NultimateForce
    mReserve1 = plane1.ReserveFactor
    mReserve2 = plane2.ReserveFactor
    mEta1 = plane1.Eta
    mEta2 = plane2.Eta
    mPhiL1 = plane1.PhiL
    mPhiL2 = plane2.PhiL
    mEccentricity1 = plane1.Eccentricity
    mEccentricity2 = plane2.Eccentricity
    mAccidentalEccentricity1 = plane1.AccidentalEccentricity
    mAccidentalEccentricity2 = plane2.AccidentalEccentricity
    mEffectiveLength1 = plane1.EffectiveLength
    mEffectiveLength2 = plane2.EffectiveLength
    mStiffnessD1 = plane1.StiffnessD
    mStiffnessD2 = plane2.StiffnessD

    mCriticalForce = PositiveMinValue(plane1.CriticalForce, plane2.CriticalForce)
    mReserveFactor = PositiveMinValue(plane1.ReserveFactor, plane2.ReserveFactor)
    mCheckPassed = PlanePassedOrInactive(plane1) And PlanePassedOrInactive(plane2)
End Sub

' Возвращает моменты из главных осей обратно в глобальные Mx/My.
Private Sub TransformDesignMoments(ByVal angleValue As Double, ByVal moment1 As Double, ByVal moment2 As Double)
    Dim axis1X As Double
    Dim axis1Y As Double
    Dim axis2X As Double
    Dim axis2Y As Double
    axis1X = Cos(angleValue)
    axis1Y = Sin(angleValue)
    axis2X = -axis1Y
    axis2Y = axis1X
    mDesignMx = moment1 * axis1X + moment2 * axis2X
    mDesignMy = moment1 * axis1Y + moment2 * axis2Y
End Sub

' Возвращает сжимающую силу положительным числом. Во внутренней системе
' проекта сжатие имеет отрицательный знак.
Private Function CompressionForce(ByVal nValue As Double) As Double
    If nValue < -SMALL_FORCE Then CompressionForce = -nValue
End Function

' Случайный эксцентриситет для СП 63; для СП 35 метод позже заменяет его
' формулой l0/400 в соответствующей ветви.
Private Function AccidentalEccentricity(ByVal effectiveLength As Double, ByVal depthValue As Double) As Double
    AccidentalEccentricity = GeomMax(GeomMax(effectiveLength / 600#, depthValue / 30#), 10#)
End Function

' Определяет знак случайного эксцентриситета: при ненулевом моменте добавляет
' его в ту же сторону, при нулевом моменте использует пользовательские +1/-1.
Private Function SignForEccentricity(ByVal momentValue As Double, ByVal userSign As Double) As Double
    If Abs(momentValue) > SMALL_VALUE Then
        SignForEccentricity = Sgn(momentValue)
    ElseIf userSign < 0# Then
        SignForEccentricity = -1#
    Else
        SignForEccentricity = 1#
    End If
End Function

' Начальный эксцентриситет по типу расчетной схемы. Для статически определимой
' схемы складываются расчетный момент и случайный эксцентриситет; для
' неопределимой берется не меньше случайного эксцентриситета.
Private Function InitialEccentricity(ByVal momentValue As Double, ByVal nCompression As Double, _
        ByVal accidentalEccentricity As Double, ByVal systemType As String) As Double
    Dim staticEccentricity As Double
    If nCompression > SMALL_FORCE Then staticEccentricity = momentValue / nCompression

    If UCase$(Trim$(systemType)) = UCase$(SYSTEM_INDETERMINATE) Then
        InitialEccentricity = SignForEccentricity(momentValue, accidentalEccentricity) * _
            GeomMax(Abs(staticEccentricity), Abs(accidentalEccentricity))
    Else
        InitialEccentricity = staticEccentricity + accidentalEccentricity
    End If
End Function

' Возвращает phi_l. В Auto используется отношение длительного момента к
' полному, но не больше 2. При PhiL2 сразу принимается консервативное значение 2.
Private Function PhiLByMode(ByVal phiLMode As String, ByVal totalMoment As Double, ByVal sustainedMoment As Double) As Double
    If UCase$(Trim$(phiLMode)) = UCase$(PHIL_2) Then
        PhiLByMode = 2#
    ElseIf Abs(totalMoment) <= SMALL_VALUE Then
        If Abs(sustainedMoment) <= SMALL_VALUE Then PhiLByMode = 1# Else PhiLByMode = 2#
    Else
        PhiLByMode = 1# + Abs(sustainedMoment / totalMoment)
        If PhiLByMode > 2# Then PhiLByMode = 2#
        If PhiLByMode < 1# Then PhiLByMode = 1#
    End If
End Function

' Коэффициент eta по критической силе. Если N >= Ncr, eta не имеет физически
' корректного положительного значения, поэтому возвращается 0 и статус станет FAIL.
Private Function EtaByCriticalForce(ByVal criticalForce As Double, ByVal nCompression As Double) As Double
    If criticalForce <= nCompression + SMALL_FORCE Then Exit Function
    EtaByCriticalForce = 1# / (1# - nCompression / criticalForce)
End Function

' Запас как отношение критической/предельной силы к действующей сжимающей.
Private Function ReserveByCriticalForce(ByVal criticalForce As Double, ByVal nCompression As Double) As Double
    If criticalForce > 0# And nCompression > SMALL_FORCE Then ReserveByCriticalForce = criticalForce / nCompression
End Function

' Расчетный момент после учета eta.
Private Function DesignMomentWithEta(ByVal nCompression As Double, ByVal eccentricity As Double, ByVal etaValue As Double) As Double
    If etaValue > 0# Then
        DesignMomentWithEta = nCompression * eccentricity * etaValue
    Else
        DesignMomentWithEta = nCompression * eccentricity
    End If
End Function

' Интерполирует таблицу 7.21 СП 35, переданную с листа Config. В книге
' таблица сохранена в человекочитаемом виде: l0/b, l0/d, l0/i, далее phi_m
' для нескольких e/r и табличный phi_l. Расчет использует l0/i и phi-колонки.
Private Sub InterpolateSP35Table(ByVal table721 As Variant, ByVal slenderness As Double, ByVal q As Double, _
        ByRef phiM As Double, ByRef phiL As Double)
    If IsEmpty(table721) Then RaiseInput "Для СП 35 нужна таблица 7.21 на Config."

    Dim rowCount As Long
    Dim columnCount As Long
    On Error GoTo InvalidTable
    rowCount = UBound(table721, 1)
    columnCount = UBound(table721, 2)
    If rowCount < 2 Or columnCount < 8 Then GoTo InvalidTable

    Dim lowerRow As Long
    Dim upperRow As Long
    FindBoundingRows table721, rowCount, slenderness, lowerRow, upperRow

    Dim phiMLower As Double
    Dim phiMUpper As Double
    Dim phiLLower As Double
    Dim phiLUpper As Double
    phiMLower = InterpolatePhiMByQ(table721, lowerRow, q)
    phiMUpper = InterpolatePhiMByQ(table721, upperRow, q)
    phiLLower = CDbl(table721(lowerRow, 8))
    phiLUpper = CDbl(table721(upperRow, 8))

    Dim t As Double
    If lowerRow = upperRow Then
        t = 0#
    Else
        t = (slenderness - CDbl(table721(lowerRow, 3))) / _
            (CDbl(table721(upperRow, 3)) - CDbl(table721(lowerRow, 3)))
    End If
    phiM = phiMLower + t * (phiMUpper - phiMLower)
    phiL = phiLLower + t * (phiLUpper - phiLLower)
    Exit Sub

InvalidTable:
    RaiseInput "Таблица 7.21 СП 35 должна содержать числовые колонки l0/b, l0/d, l0/i, phi_m(0), phi_m(0.25), phi_m(0.5), phi_m(1), phi_l."
End Sub

' Находит соседние строки таблицы по гибкости.
Private Sub FindBoundingRows(ByVal tableData As Variant, ByVal rowCount As Long, ByVal slenderness As Double, _
        ByRef lowerRow As Long, ByRef upperRow As Long)
    lowerRow = 2
    upperRow = rowCount
    Dim r As Long
    For r = 2 To rowCount
        If CDbl(tableData(r, 3)) <= slenderness Then lowerRow = r
        If CDbl(tableData(r, 3)) >= slenderness Then
            upperRow = r
            Exit For
        End If
    Next r
    If slenderness <= CDbl(tableData(2, 3)) Then lowerRow = 2: upperRow = 2
    If slenderness >= CDbl(tableData(rowCount, 3)) Then lowerRow = rowCount: upperRow = rowCount
End Sub

' Интерполяция phi_m по отношению e/r.
Private Function InterpolatePhiMByQ(ByVal tableData As Variant, ByVal rowIndex As Long, ByVal q As Double) As Double
    Dim qValues(1 To 4) As Double
    qValues(1) = 0#: qValues(2) = 0.25: qValues(3) = 0.5: qValues(4) = 1#
    If q <= qValues(1) Then
        InterpolatePhiMByQ = CDbl(tableData(rowIndex, 4))
        Exit Function
    End If
    If q >= qValues(4) Then
        InterpolatePhiMByQ = CDbl(tableData(rowIndex, 7))
        Exit Function
    End If

    Dim i As Long
    For i = 1 To 3
        If q >= qValues(i) And q <= qValues(i + 1) Then
            Dim t As Double
            t = (q - qValues(i)) / (qValues(i + 1) - qValues(i))
            InterpolatePhiMByQ = CDbl(tableData(rowIndex, i + 3)) + _
                t * (CDbl(tableData(rowIndex, i + 4)) - CDbl(tableData(rowIndex, i + 3)))
            Exit Function
        End If
    Next i
End Function

' Минимальное положительное значение. Нули трактуются как отсутствие величины.
Private Function PositiveMinValue(ByVal firstValue As Double, ByVal secondValue As Double) As Double
    If firstValue > 0# And secondValue > 0# Then
        PositiveMinValue = GeomMin(firstValue, secondValue)
    ElseIf firstValue > 0# Then
        PositiveMinValue = firstValue
    ElseIf secondValue > 0# Then
        PositiveMinValue = secondValue
    End If
End Function

' Плоскость считается пройденной, если она применима и дала запас, либо если
' она была безопасно признана неактивной.
Private Function PlanePassedOrInactive(ByRef plane As TPlaneResult) As Boolean
    PlanePassedOrInactive = (Not plane.Applicable Or plane.CheckPassed)
End Function

' Положительный коэффициент длины; пустое/нулевое значение трактуется как 1.
Private Function PositiveOrOne(ByVal value As Double) As Double
    If value > 0# Then PositiveOrOne = value Else PositiveOrOne = 1#
End Function

' Нормализует имя нормативного блока.
Private Function NormalizeCode(ByVal value As String) As String
    Select Case UCase$(Trim$(value))
        Case "SP35", "СП35", "СП 35"
            NormalizeCode = CODE_SP35
        Case "SP63", "СП63", "СП 63"
            NormalizeCode = CODE_SP63
        Case Else
            NormalizeCode = Trim$(value)
    End Select
End Function

' Фиксирует ошибку исходных данных и прерывает расчет класса.
Private Sub RaiseInput(ByVal message As String)
    mInputValid = False
    mStopReason = message
    Err.Raise vbObjectError + 7600, "CStabilityCalculator", message
End Sub

Public Property Get Code() As String
    Code = mCode
End Property

Public Property Get Branch() As String
    Branch = mBranch
End Property

Public Property Get Applicable() As Boolean
    Applicable = mApplicable
End Property

Public Property Get InputValid() As Boolean
    InputValid = mInputValid
End Property

Public Property Get CheckPassed() As Boolean
    CheckPassed = mCheckPassed
End Property

Public Property Get StopReason() As String
    StopReason = mStopReason
End Property

Public Property Get DesignN() As Double
    DesignN = mDesignN
End Property

Public Property Get DesignMx() As Double
    DesignMx = mDesignMx
End Property

Public Property Get DesignMy() As Double
    DesignMy = mDesignMy
End Property

Public Property Get CriticalForce() As Double
    CriticalForce = mCriticalForce
End Property

Public Property Get ReserveFactor() As Double
    ReserveFactor = mReserveFactor
End Property

Public Property Get Ncr1() As Double
    Ncr1 = mNcr1
End Property

Public Property Get Ncr2() As Double
    Ncr2 = mNcr2
End Property

Public Property Get Nultimate1() As Double
    Nultimate1 = mNultimate1
End Property

Public Property Get Nultimate2() As Double
    Nultimate2 = mNultimate2
End Property

Public Property Get Reserve1() As Double
    Reserve1 = mReserve1
End Property

Public Property Get Reserve2() As Double
    Reserve2 = mReserve2
End Property

Public Property Get Eta1() As Double
    Eta1 = mEta1
End Property

Public Property Get Eta2() As Double
    Eta2 = mEta2
End Property

Public Property Get PhiL1() As Double
    PhiL1 = mPhiL1
End Property

Public Property Get PhiL2() As Double
    PhiL2 = mPhiL2
End Property

Public Property Get Eccentricity1() As Double
    Eccentricity1 = mEccentricity1
End Property

Public Property Get Eccentricity2() As Double
    Eccentricity2 = mEccentricity2
End Property

Public Property Get AccidentalEccentricity1() As Double
    AccidentalEccentricity1 = mAccidentalEccentricity1
End Property

Public Property Get AccidentalEccentricity2() As Double
    AccidentalEccentricity2 = mAccidentalEccentricity2
End Property

Public Property Get EffectiveLength1() As Double
    EffectiveLength1 = mEffectiveLength1
End Property

Public Property Get EffectiveLength2() As Double
    EffectiveLength2 = mEffectiveLength2
End Property

Public Property Get StiffnessD1() As Double
    StiffnessD1 = mStiffnessD1
End Property

Public Property Get StiffnessD2() As Double
    StiffnessD2 = mStiffnessD2
End Property
