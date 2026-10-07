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
' Класс является изолированным фильтром усилий перед НДМ.
' Он не запускает CSectionSolver, не читает Excel и не меняет материалы:
' вызывающий код передает готовую CSectionModel, выбранный набор характеристик
' материалов профиля и усилия LC. На выходе возвращаются N/Mx/My с учетом
' продольного изгиба, а также расчетные величины для Summary.
'
' Допустимость настроек проверяет тот же класс до расчетных ветвей.
' Уже прочитанный reader может добавить только адрес Config в ошибку;
' самого чтения листов, выбора профиля и преобразования единиц здесь нет.
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
Private Const ACC_AUTO_L As String = "AutoWithL" ' Случайный эксцентриситет по геометрической длине L.
Private Const ACC_AUTO_MUL As String = "AutoWithMuL" ' Случайный эксцентриситет по расчетной длине mu*L.
Private Const ACC_USER As String = "User" ' Пользователь задает модуль случайного эксцентриситета.
Private Const ACC_PLANES_ONLY_MOMENT As String = "OnlyMomentPlane" ' Случайный эксцентриситет только в плоскости исходного момента.
Private Const ACC_PLANES_BOTH As String = "BothPlanes" ' Случайный эксцентриситет в обеих главных плоскостях.
Private Const SMALL_FORCE As Double = 0.000001 ' Малый допуск по силе, Н.
Private Const SMALL_VALUE As Double = 0.000000001 ' Общий малый допуск для безразмерных величин.

Private Const PASS_TOLERANCE As Double = 0.000000001 ' Малый допуск для граничного PASS.
Private Const SP35_REBAR_AREA_LIMIT_RATIO As Double = 0.03 ' Порог 3% площади арматуры по СП 35.

Private Type TPlaneResult
    Moment As Double
    SustainedMoment As Double
    DesignMoment As Double
    InertiaConcrete As Double
    InertiaSteel As Double
    Radius As Double
    CoreDistancePositive As Double
    CoreDistanceNegative As Double
    Depth As Double
    GeometricLength As Double
    EffectiveLength As Double
    Slenderness As Double
    NormalX As Double
    NormalY As Double
    CentroidX As Double
    CentroidY As Double
    ConcreteMinProjection As Double
    ConcreteMaxProjection As Double
    AccidentalEccentricity As Double
    Eccentricity As Double
    CoreDistance As Double
    PhiL As Double
    PhiLTable As Double
    PhiM As Double
    PhiValue As Double
    PhiP As Double
    DeltaValue As Double
    Kb As Double
    Ks As Double
    Eta As Double
    StiffnessD As Double
    CriticalForce As Double
    NcrForce As Double
    NultimateForce As Double
    ReserveFactor As Double
    CheckPassed As Boolean
    Applicable As Boolean
    BranchName As String
End Type

Private mCode As String ' Нормативная ветвь, выбранная настройкой Stability.Code.
Private mBranch As String ' Внутренняя трассировка расчетной ветви для execution report и диагностики.
Private mApplicable As Boolean ' False, если сжатия нет и устойчивость не требуется.
Private mInputValid As Boolean ' False при нехватке обязательных исходных данных.
Private mResultInitialized As Boolean ' Был вызван расчет; новый объект не является готовым результатом.
Private mCalculationStarted As Boolean ' Выполнялась собственная инженерная проверка после валидации.
Private mInternalError As Boolean ' Нарушение программного контракта или неожиданный runtime-отказ, не ошибка пользователя.
Private mCheckPassed As Boolean ' Итог проверки по выбранной ветви.
Private mStopReason As String ' Подробная внутренняя причина, если расчет не выполнен.
Private mSP35TableInputError As String ' Адресная причина из Excel-слоя; не влияет на ветви, которым таблица не нужна.
Private mDesignN As Double ' Продольная сила после фильтра; знак сохраняется как во внутреннем ядре.
Private mDesignMx As Double ' Mx после учета случайного эксцентриситета и eta.
Private mDesignMy As Double ' My после учета случайного эксцентриситета и eta.
Private mPrincipalAngleRad As Double ' Угол от +X к главной оси 1 приведенного сечения.
Private mSustainedN As Double ' Сжимающая часть постоянной/длительной силы, переданная для учета длительности.
Private mMoment1 As Double ' Исходный расчетный момент в первой главной плоскости до eta.
Private mMoment2 As Double ' Исходный расчетный момент во второй главной плоскости до eta.
Private mSustainedMoment1 As Double ' Постоянный/длительный момент в первой главной плоскости.
Private mSustainedMoment2 As Double ' Постоянный/длительный момент во второй главной плоскости.
Private mDesignMoment1 As Double ' Итоговый момент первой главной плоскости после учета eta.
Private mDesignMoment2 As Double ' Итоговый момент второй главной плоскости после учета eta.
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
Private mPhiLTable1 As Double ' Табличный phi_l СП 35 первой плоскости для ветви ec/r <= 1.
Private mPhiLTable2 As Double ' Табличный phi_l СП 35 второй плоскости для ветви ec/r <= 1.
Private mPhiM1 As Double ' Табличный phi_m СП 35 первой плоскости.
Private mPhiM2 As Double ' Табличный phi_m СП 35 второй плоскости.
Private mPhiValue1 As Double ' Итоговый phi СП 35 первой плоскости в ветви ec <= r.
Private mPhiValue2 As Double ' Итоговый phi СП 35 второй плоскости в ветви ec <= r.
Private mPhiP1 As Double ' Коэффициент phi_p СП 35 первой плоскости в ветви ec > r.
Private mPhiP2 As Double ' Коэффициент phi_p СП 35 второй плоскости в ветви ec > r.
Private mEccentricity1 As Double ' Расчетный эксцентриситет первой плоскости, мм.
Private mEccentricity2 As Double ' Расчетный эксцентриситет второй плоскости, мм.
Private mAccidentalEccentricity1 As Double ' Случайный эксцентриситет первой плоскости, мм.
Private mAccidentalEccentricity2 As Double ' Случайный эксцентриситет второй плоскости, мм.
Private mEffectiveLength1 As Double ' l0 первой плоскости, мм.
Private mEffectiveLength2 As Double ' l0 второй плоскости, мм.
Private mSlenderness1 As Double ' Гибкость l0/i для плоскости 1: изгиб вокруг главной оси 1.
Private mSlenderness2 As Double ' Гибкость l0/i для плоскости 2: изгиб вокруг главной оси 2.
Private mDepth1 As Double ' Размер сечения h1 по нормали первой главной плоскости, мм.
Private mDepth2 As Double ' Размер сечения h2 по нормали второй главной плоскости, мм.
Private mRadius1 As Double ' Радиус инерции i1 приведенного сечения.
Private mRadius2 As Double ' Радиус инерции i2 приведенного сечения.
Private mCoreDistance1 As Double ' Ядровое расстояние r первой плоскости со стороны эксцентриситета.
Private mCoreDistance2 As Double ' Ядровое расстояние r второй плоскости со стороны эксцентриситета.
Private mDelta1 As Double ' delta или delta_e первой плоскости в выбранной расчетной ветви.
Private mDelta2 As Double ' delta или delta_e второй плоскости в выбранной расчетной ветви.
Private mKb1 As Double ' Коэффициент kb СП 63 первой плоскости.
Private mKb2 As Double ' Коэффициент kb СП 63 второй плоскости.
Private mKs1 As Double ' Коэффициент ks СП 63 первой плоскости.
Private mKs2 As Double ' Коэффициент ks СП 63 второй плоскости.
Private mStiffnessD1 As Double ' Условная жесткость D первой плоскости, Н*мм2.
Private mStiffnessD2 As Double ' Условная жесткость D второй плоскости, Н*мм2.
Private mPlaneBranch1 As String ' Фактическая ветвь расчета первой плоскости.
Private mPlaneBranch2 As String ' Фактическая ветвь расчета второй плоскости.
Private mPlaneApplicable1 As Boolean ' True, если первая плоскость участвовала в проверке.
Private mPlaneApplicable2 As Boolean ' True, если вторая плоскость участвовала в проверке.
Private mPlanePassed1 As Boolean ' True, если первая плоскость прошла свою проверку.
Private mPlanePassed2 As Boolean ' True, если вторая плоскость прошла свою проверку.
Private mPreparedTransformedProperties As CSectionPropertiesCalculator
Private mPreparedConcreteProperties As CSectionPropertiesCalculator ' Только механическая геометрия; LC-dependent Ncr/eta не переиспользуются.

' Подключает подготовку одного запуска. Реальные Eb/Es и геометрия проверяются
' владельцем properties лишь в применимой ветви после проверки входных данных.
Friend Sub UseMechanicalPreparation(ByVal transformed As CSectionPropertiesCalculator, ByVal concrete As CSectionPropertiesCalculator)
    Set mPreparedTransformedProperties = transformed
    Set mPreparedConcreteProperties = concrete
End Sub

' Проверяет допустимость настроек до геометрии и раннего выхода без сжатия.
' Значения уже нормализованы по единицам; reader нужен только для текущего
' адреса Config. Нулевые delta-границы отключают соответствующее ограничение,
' а не подменяются нормативным значением. Проверка не решает НДС.
Public Sub ValidateSettings(ByVal codeText As String, ByVal elementLength As Double, _
        ByVal mu1 As Double, ByVal mu2 As Double, ByVal systemType As String, _
        ByVal zeroMomentSign1 As Double, ByVal zeroMomentSign2 As Double, _
        ByVal phiLMode As String, ByVal accidentalMode As String, ByVal accidentalPlanes As String, _
        ByVal accidentalUserValue1 As Double, ByVal accidentalUserValue2 As Double, _
        ByVal sp63Ks As Double, ByVal sp63DeltaEMin As Double, ByVal sp63DeltaEMax As Double, _
        ByVal sp35PhiP As Double, ByVal sp35NOverNcrLimit As Double, _
        Optional ByVal settings As CSystemSettingsReader = Nothing)
    Dim code As String, mode As String
    code = NormalizeCode(codeText)
    If code <> CODE_SP63 And code <> CODE_SP35 Then RaiseSettingInput "Stability.Code", _
        "Неизвестная методика устойчивости.", "Выберите SP63 или SP35.", settings
    If elementLength <= 0# Then RaiseSettingInput "Stability.ElementLength", _
        "Длина элемента должна быть положительной.", "Введите число больше нуля в единицах INPUT Length.", settings
    If mu1 <= 0# Then RaiseSettingInput "Stability.Mu1", _
        "Коэффициент расчетной длины плоскости 1 должен быть положительным.", "Задайте число больше нуля.", settings
    If mu2 <= 0# Then RaiseSettingInput "Stability.Mu2", _
        "Коэффициент расчетной длины плоскости 2 должен быть положительным.", "Задайте число больше нуля.", settings
    ValidateEffectiveLengthRange elementLength, mu1, "Stability.Mu1", settings
    ValidateEffectiveLengthRange elementLength, mu2, "Stability.Mu2", settings
    If StrComp(Trim$(systemType), SYSTEM_DETERMINATE, vbTextCompare) <> 0 And _
            StrComp(Trim$(systemType), SYSTEM_INDETERMINATE, vbTextCompare) <> 0 Then _
        RaiseSettingInput "Stability.SystemType", "Неизвестная расчетная схема.", _
            "Выберите Determinate или Indeterminate.", settings
    If zeroMomentSign1 <> 1# And zeroMomentSign1 <> -1# Then RaiseSettingInput "Stability.ZeroMomentEccentricitySign1", _
        "Направление задается только знаком.", "Введите 1 или -1 для главной плоскости 1.", settings
    If zeroMomentSign2 <> 1# And zeroMomentSign2 <> -1# Then RaiseSettingInput "Stability.ZeroMomentEccentricitySign2", _
        "Направление задается только знаком.", "Введите 1 или -1 для главной плоскости 2.", settings
    If StrComp(Trim$(phiLMode), PHIL_AUTO, vbTextCompare) <> 0 And _
            StrComp(Trim$(phiLMode), PHIL_2, vbTextCompare) <> 0 Then _
        RaiseSettingInput "Stability.PhiLMode", "Неизвестный режим учета длительности.", "Выберите Auto или PhiL2.", settings
    mode = NormalizeAccidentalMode(accidentalMode, settings)
    Call NormalizeAccidentalPlanes(accidentalPlanes, settings)
    If accidentalUserValue1 < 0# Or (mode = ACC_USER And accidentalUserValue1 <= SMALL_VALUE) Then _
        RaiseSettingInput "Stability.AccidentalEccentricityUser1", "Требуется модуль, а не знаковый эксцентриситет.", _
            "Введите положительное значение для User; в автоматическом режиме допустим также 0.", settings
    If accidentalUserValue2 < 0# Or (mode = ACC_USER And accidentalUserValue2 <= SMALL_VALUE) Then _
        RaiseSettingInput "Stability.AccidentalEccentricityUser2", "Требуется модуль, а не знаковый эксцентриситет.", _
            "Введите положительное значение для User; в автоматическом режиме допустим также 0.", settings
    If sp63Ks <= 0# Then RaiseSettingInput "Stability.SP63.Ks", _
        "Коэффициент вклада арматуры должен быть положительным.", "Введите число больше нуля.", settings
    If sp63DeltaEMin < 0# Then RaiseSettingInput "Stability.SP63.DeltaEMin", _
        "Нижнее ограничение не может быть отрицательным.", "Введите число не меньше нуля; 0 отключает нижнюю границу.", settings
    If sp63DeltaEMax < 0# Then RaiseSettingInput "Stability.SP63.DeltaEMax", _
        "Верхнее ограничение не может быть отрицательным.", "Введите число не меньше нуля; 0 отключает верхнюю границу.", settings
    If sp63DeltaEMin > 0# And sp63DeltaEMax > 0# And sp63DeltaEMin > sp63DeltaEMax Then _
        RaiseSettingInput "Stability.SP63.DeltaEMin", "Нижняя граница больше верхней.", _
            "Исправьте DeltaEMin или DeltaEMax: при двух активных границах нижняя не должна превышать верхнюю.", settings
    If sp35PhiP <= 0# Then RaiseSettingInput "Stability.SP35.PhiP", _
        "Коэффициент должен быть положительным.", "Введите число больше нуля.", settings
    If sp35NOverNcrLimit <= 0# Then RaiseSettingInput "Stability.SP35.NOverNcrLimit", _
        "Предельное отношение N/Ncr должно быть положительным.", "Введите число больше нуля.", settings
End Sub

' Проверяет представимость расчетной длины и ее положительного квадрата.
' Это техническая граница арифметики, не нормативный предел длины или mu.
' При переполнении или обнулении сообщает обе ячейки до геометрии и расчета.
Private Sub ValidateEffectiveLengthRange(ByVal elementLength As Double, ByVal mu As Double, _
        ByVal muKey As String, ByVal settings As CSystemSettingsReader)
    On Error GoTo ArithmeticFailed
    Dim effectiveLength As Double, squaredLength As Double
    Dim reason As String, correction As String, message As String
    reason = "Произведение длины элемента L и коэффициента расчетной длины выходит за допустимый диапазон арифметики."
    effectiveLength = elementLength * mu
    reason = "Квадрат расчетной длины l0 = mu·L выходит за допустимый диапазон арифметики."
    squaredLength = effectiveLength * effectiveLength
    ' Excel может оставить бесконечность без VBA-ошибки; проверяем значение,
    ' а не полагаемся только на обработчик переполнения.
    If squaredLength <> squaredLength Or squaredLength > 1.797693134862315E+308 Then GoTo RangeFailed
    If squaredLength <= 0# Then
        reason = "Расчетная длина l0 = mu·L слишком мала: ее квадрат округляется до нуля в арифметике программы."
        correction = "Увеличьте длину L или коэффициент " & muKey & "; проверьте INPUT-единицу длины."
        GoTo RangeFailed
    End If
    Exit Sub
ArithmeticFailed:
    Dim number As Long, description As String
    number = Err.Number: description = Err.Description
    Err.Clear
    If number <> 6 Then Err.Raise number, "CStabilityCalculator", description
RangeFailed:
    If Len(correction) = 0 Then correction = "Уменьшите длину L или коэффициент " & muKey & "; проверьте INPUT-единицу длины."
    If settings Is Nothing Then
        message = "Настройки Stability.ElementLength и " & muKey & ": " & reason & " " & correction
    Else
        message = settings.InputErrorMessage("Stability.ElementLength", reason, correction) & " " & _
            settings.InputErrorMessage(muKey, "Этот коэффициент участвует в расчетной длине l0 = mu·L и ее квадрате.", vbNullString)
    End If
    RaiseInput message
End Sub

' Выполняет расчет по выбранному нормативному блоку и возвращает расчетные
' усилия для дальнейшего НДМ. Если продольная сила не сжимает сечение, фильтр
' считается неприменимым и усилия передаются дальше без изменения.
Public Sub Calculate(ByVal section As CSectionModel, ByVal materialProvider As CMaterialModelProvider, _
        ByVal valueSet As EMaterialValueSet, ByVal codeText As String, _
        ByVal nValue As Double, ByVal mxValue As Double, ByVal myValue As Double, _
        ByVal sustainedN As Double, ByVal sustainedMx As Double, ByVal sustainedMy As Double, _
        ByVal elementLength As Double, ByVal mu1 As Double, ByVal mu2 As Double, _
        ByVal systemType As String, ByVal zeroMomentSign1 As Double, ByVal zeroMomentSign2 As Double, _
        ByVal phiLMode As String, ByVal accidentalMode As String, _
        ByVal accidentalPlanes As String, _
        ByVal accidentalUserValue1 As Double, ByVal accidentalUserValue2 As Double, _
        ByVal sp63Ks As Double, ByVal sp63DeltaEMin As Double, _
        ByVal sp63DeltaEMax As Double, ByVal sp35PhiP As Double, _
        ByVal sp35NOverNcrLimit As Double, ByVal sp35Table721 As Variant, _
        Optional ByVal momentFilter As CMomentZeroFilter = Nothing, _
        Optional ByVal sp35TableInputError As String = vbNullString)
    On Error GoTo Failed
    Clear
    mSP35TableInputError = sp35TableInputError
    mCode = NormalizeCode(codeText)
    ValidateSettings codeText, elementLength, mu1, mu2, systemType, zeroMomentSign1, zeroMomentSign2, _
        phiLMode, accidentalMode, accidentalPlanes, accidentalUserValue1, accidentalUserValue2, _
        sp63Ks, sp63DeltaEMin, sp63DeltaEMax, sp35PhiP, sp35NOverNcrLimit

    Dim totalLoad As CSectionLoadState
    Set totalLoad = New CSectionLoadState
    totalLoad.Initialize nValue, mxValue, myValue, 0#, 0#

    Dim sustainedLoad As CSectionLoadState
    Set sustainedLoad = New CSectionLoadState
    sustainedLoad.Initialize sustainedN, sustainedMx, sustainedMy, 0#, 0#

    mDesignN = totalLoad.N
    mDesignMx = totalLoad.InternalMx
    mDesignMy = totalLoad.InternalMy
    mSustainedN = sustainedLoad.SignedCompression

    If section Is Nothing Then
        mInternalError = True
        mStopReason = "Для проверки устойчивости не передана модель сечения."
        GoTo Failed
    End If
    If materialProvider Is Nothing Then
        mInternalError = True
        mStopReason = "Для проверки устойчивости не передан поставщик материалов."
        GoTo Failed
    End If
    Dim nCompression As Double
    nCompression = totalLoad.CompressionMagnitude(SMALL_FORCE)
    If nCompression <= SMALL_FORCE Then
        mApplicable = False
        mCheckPassed = True
        mBranch = "NoCompression"
        Exit Sub
    End If

    Dim ebForGeometry As Double
    Dim esForGeometry As Double
    ebForGeometry = materialProvider.ConcreteElasticModulusFromValueSet(valueSet)
    esForGeometry = materialProvider.SteelElasticModulusFromValueSet(valueSet)
    If ebForGeometry <= 0# Or esForGeometry <= 0# Then RaiseInput "Для расчета устойчивости нужны положительные Eb и Es."

    Dim props As CSectionPropertiesCalculator
    mCalculationStarted = True
    If mPreparedTransformedProperties Is Nothing Then
        Set props = New CSectionPropertiesCalculator
    Else
        Set props = mPreparedTransformedProperties
    End If
    props.EnsureTransformed section, ebForGeometry, esForGeometry
    mPrincipalAngleRad = props.PrincipalAngleRad

    Dim concreteProps As CSectionPropertiesCalculator
    If mPreparedConcreteProperties Is Nothing Then
        Set concreteProps = New CSectionPropertiesCalculator
    Else
        Set concreteProps = mPreparedConcreteProperties
    End If
    concreteProps.EnsureConcrete section

    Dim plane1 As TPlaneResult
    Dim plane2 As TPlaneResult
    PreparePlanes section, props, mDesignMx, mDesignMy, sustainedLoad.InternalMx, sustainedLoad.InternalMy, _
        elementLength, mu1, mu2, systemType, zeroMomentSign1, zeroMomentSign2, _
        accidentalMode, accidentalPlanes, accidentalUserValue1, accidentalUserValue2, _
        momentFilter, plane1, plane2

    If mCode = CODE_SP63 Then
        CalculateSP63 section, materialProvider, valueSet, nCompression, sustainedLoad.SignedCompression, _
            systemType, phiLMode, sp63Ks, sp63DeltaEMin, sp63DeltaEMax, plane1, plane2
    Else
        CalculateSP35 section, materialProvider, valueSet, concreteProps.Area, props.RebarAreaTotal(section), _
            nCompression, sustainedLoad.SignedCompression, systemType, phiLMode, _
            sp35PhiP, sp35NOverNcrLimit, sp35Table721, plane1, plane2
    End If

    StorePlaneResults plane1, plane2
    TransformDesignMoments props, plane1.DesignMoment, plane2.DesignMoment
    Exit Sub

Failed:
    If Err.Number <> 0 And Err.Number <> vbObjectError + 7600 Then mInternalError = True
    mInputValid = False
    mApplicable = False
    mCheckPassed = False
    If Len(mStopReason) = 0 Then mStopReason = Err.Description
End Sub

' Сбрасывает результаты перед новым расчетом.
Private Sub Clear()
    mResultInitialized = True
    mCalculationStarted = False
    mInternalError = False
    mCode = vbNullString
    mBranch = vbNullString
    mApplicable = True
    mInputValid = True
    mCheckPassed = False
    mStopReason = vbNullString
    mSP35TableInputError = vbNullString
    mDesignN = 0#: mDesignMx = 0#: mDesignMy = 0#
    mPrincipalAngleRad = 0#
    mSustainedN = 0#
    mMoment1 = 0#: mMoment2 = 0#
    mSustainedMoment1 = 0#: mSustainedMoment2 = 0#
    mDesignMoment1 = 0#: mDesignMoment2 = 0#
    mCriticalForce = 0#: mReserveFactor = 0#
    mNcr1 = 0#: mNcr2 = 0#
    mNultimate1 = 0#: mNultimate2 = 0#
    mReserve1 = 0#: mReserve2 = 0#
    mEta1 = 0#: mEta2 = 0#
    mPhiL1 = 0#: mPhiL2 = 0#
    mPhiLTable1 = 0#: mPhiLTable2 = 0#
    mPhiM1 = 0#: mPhiM2 = 0#
    mPhiValue1 = 0#: mPhiValue2 = 0#
    mPhiP1 = 0#: mPhiP2 = 0#
    mEccentricity1 = 0#: mEccentricity2 = 0#
    mAccidentalEccentricity1 = 0#: mAccidentalEccentricity2 = 0#
    mEffectiveLength1 = 0#: mEffectiveLength2 = 0#
    mSlenderness1 = 0#: mSlenderness2 = 0#
    mDepth1 = 0#: mDepth2 = 0#
    mRadius1 = 0#: mRadius2 = 0#
    mCoreDistance1 = 0#: mCoreDistance2 = 0#
    mDelta1 = 0#: mDelta2 = 0#
    mKb1 = 0#: mKb2 = 0#
    mKs1 = 0#: mKs2 = 0#
    mStiffnessD1 = 0#: mStiffnessD2 = 0#
    mPlaneBranch1 = vbNullString: mPlaneBranch2 = vbNullString
    mPlaneApplicable1 = False: mPlaneApplicable2 = False
    mPlanePassed1 = False: mPlanePassed2 = False
End Sub

' Подготавливает две главные расчетные плоскости: моменты переводятся в
' главные центральные оси приведенного сечения. Это нормативная база
' эксцентриситетов устойчивости по СП 35 и СП 63; пользовательская точка
' нагрузки к этому моменту уже перенесена batch-ом в тот же центр. Плоскость 1
' связана с изгибом вокруг первой главной оси, плоскость 2 - вокруг второй.
' Если главные оси совпали с X/Y, плоскость 1 - это изгиб Mx вокруг X:
' расчетная плоскость YZ, эксцентриситет и габарит берутся по Y. Плоскость 2 -
' это изгиб My вокруг Y: расчетная плоскость XZ, эксцентриситет и габарит по X.
Private Sub PreparePlanes(ByVal section As CSectionModel, ByVal props As CSectionPropertiesCalculator, _
        ByVal mxValue As Double, ByVal myValue As Double, ByVal sustainedMx As Double, _
        ByVal sustainedMy As Double, ByVal elementLength As Double, ByVal mu1 As Double, _
        ByVal mu2 As Double, ByVal systemType As String, ByVal zeroMomentSign1 As Double, _
        ByVal zeroMomentSign2 As Double, ByVal accidentalMode As String, _
        ByVal accidentalPlanes As String, ByVal accidentalUserValue1 As Double, _
        ByVal accidentalUserValue2 As Double, ByVal momentFilter As CMomentZeroFilter, _
        ByRef plane1 As TPlaneResult, ByRef plane2 As TPlaneResult)
    Dim axis1X As Double
    Dim axis1Y As Double
    Dim axis2X As Double
    Dim axis2Y As Double
    props.PrincipalAxisDirection 1, axis1X, axis1Y
    props.PrincipalAxisDirection 2, axis2X, axis2Y

    plane1.Moment = mxValue * axis1X + myValue * axis1Y
    plane2.Moment = mxValue * axis2X + myValue * axis2Y
    plane1.SustainedMoment = sustainedMx * axis1X + sustainedMy * axis1Y
    plane2.SustainedMoment = sustainedMx * axis2X + sustainedMy * axis2Y
    plane1.InertiaConcrete = props.ConcreteInertiaAboutAxis(section, axis1X, axis1Y, props.CentroidX, props.CentroidY)
    plane2.InertiaConcrete = props.ConcreteInertiaAboutAxis(section, axis2X, axis2Y, props.CentroidX, props.CentroidY)
    plane1.InertiaSteel = props.RebarInertiaAboutAxis(section, axis1X, axis1Y, props.CentroidX, props.CentroidY)
    plane2.InertiaSteel = props.RebarInertiaAboutAxis(section, axis2X, axis2Y, props.CentroidX, props.CentroidY)
    plane1.Radius = props.PrincipalRadius1
    plane2.Radius = props.PrincipalRadius2
    props.PrincipalPlaneNormalDirection 1, plane1.NormalX, plane1.NormalY
    props.PrincipalPlaneNormalDirection 2, plane2.NormalX, plane2.NormalY
    plane1.CentroidX = props.CentroidX
    plane1.CentroidY = props.CentroidY
    plane2.CentroidX = props.CentroidX
    plane2.CentroidY = props.CentroidY
    StoreConcreteProjectionBounds section, props, plane1
    StoreConcreteProjectionBounds section, props, plane2
    plane1.CoreDistancePositive = props.PrincipalPlaneCoreDistance(section, 1, True, False)
    plane1.CoreDistanceNegative = props.PrincipalPlaneCoreDistance(section, 1, False, False)
    plane2.CoreDistancePositive = props.PrincipalPlaneCoreDistance(section, 2, True, False)
    plane2.CoreDistanceNegative = props.PrincipalPlaneCoreDistance(section, 2, False, False)
    plane1.Depth = props.PrincipalPlaneDepth(section, 1, False)
    plane2.Depth = props.PrincipalPlaneDepth(section, 2, False)
    NormalizePlaneMoments plane1, momentFilter
    NormalizePlaneMoments plane2, momentFilter
    plane1.GeometricLength = elementLength
    plane2.GeometricLength = elementLength
    plane1.EffectiveLength = elementLength * mu1
    plane2.EffectiveLength = elementLength * mu2
    If plane1.Radius > SMALL_VALUE Then plane1.Slenderness = plane1.EffectiveLength / plane1.Radius
    If plane2.Radius > SMALL_VALUE Then plane2.Slenderness = plane2.EffectiveLength / plane2.Radius
    plane1.Applicable = True
    plane2.Applicable = True

    ' Если момент уже задан в одной из главных плоскостей, случайный
    ' эксцентриситет добавляется только к действующему направлению момента.
    ' Для чистого сжатия без моментов остаются пользовательские направления +1/-1.
    Dim hasMomentVector As Boolean
    Dim applyBothPlanes As Boolean
    hasMomentVector = (Abs(plane1.Moment) > SMALL_VALUE Or Abs(plane2.Moment) > SMALL_VALUE)
    applyBothPlanes = (NormalizeAccidentalPlanes(accidentalPlanes) = ACC_PLANES_BOTH)
    FillAccidentalAndInitialEccentricity plane1, zeroMomentSign1, systemType, _
        accidentalMode, accidentalUserValue1, "Stability.AccidentalEccentricityUser1", _
        (applyBothPlanes Or Not hasMomentVector Or Abs(plane1.Moment) > SMALL_VALUE)
    FillAccidentalAndInitialEccentricity plane2, zeroMomentSign2, systemType, _
        accidentalMode, accidentalUserValue2, "Stability.AccidentalEccentricityUser2", _
        (applyBothPlanes Or Not hasMomentVector Or Abs(plane2.Moment) > SMALL_VALUE)
End Sub

' Применяет общий инженерный фильтр к моментам одной главной плоскости.
' Сравнение выполняется по бетонному габариту именно этой плоскости, поэтому
' после поворота осей случайный остаток M1/M2 не меняет ветвь устойчивости.
Private Sub NormalizePlaneMoments(ByRef plane As TPlaneResult, ByVal momentFilter As CMomentZeroFilter)
    If momentFilter Is Nothing Then Exit Sub
    plane.Moment = momentFilter.NormalizeMoment(plane.Moment, plane.Depth)
    plane.SustainedMoment = momentFilter.NormalizeMoment(plane.SustainedMoment, plane.Depth)
End Sub

' Назначает случайный и начальный эксцентриситет одной главной плоскости.
' Здесь сходятся пользовательский режим Auto/User, выбор L или mu*L и правило
' OnlyMomentPlane/BothPlanes для нулевых исходных моментов.
Private Sub FillAccidentalAndInitialEccentricity(ByRef plane As TPlaneResult, _
        ByVal zeroMomentSign As Double, ByVal systemType As String, _
        ByVal accidentalMode As String, ByVal accidentalUserValue As Double, _
        ByVal accidentalUserSettingName As String, _
        ByVal allowZeroMomentAccidental As Boolean)
    If Abs(plane.Moment) <= SMALL_VALUE And Not allowZeroMomentAccidental Then
        plane.AccidentalEccentricity = 0#
        plane.Eccentricity = 0#
        Exit Sub
    End If

    Dim signValue As Double
    signValue = SignForEccentricity(plane.Moment, zeroMomentSign)
    plane.AccidentalEccentricity = signValue * AccidentalEccentricity( _
        accidentalMode, accidentalUserValue, accidentalUserSettingName, _
        plane.GeometricLength, plane.EffectiveLength, plane.Depth)
    plane.Eccentricity = InitialEccentricity(plane.Moment, 1#, plane.AccidentalEccentricity, systemType)
End Sub

' Сохраняет крайние точки бетонного контура относительно центра приведенного
' сечения. Эти значения нужны для h и ядрового расстояния: арматура не должна
' расширять внешний габарит сжатого элемента.
Private Sub StoreConcreteProjectionBounds(ByVal section As CSectionModel, _
        ByVal props As CSectionPropertiesCalculator, ByRef plane As TPlaneResult)
    Dim minProjection As Double
    Dim maxProjection As Double
    props.CalculateProjection section, plane.NormalX, plane.NormalY, False, minProjection, maxProjection

    Dim centroidProjection As Double
    centroidProjection = plane.NormalX * props.CentroidX + plane.NormalY * props.CentroidY
    plane.ConcreteMinProjection = minProjection - centroidProjection
    plane.ConcreteMaxProjection = maxProjection - centroidProjection
End Sub

' Выбирает центр стержня, относительно которого СП 63 требует считать M1 и Ml1:
' наиболее растянутый или наименее сжатый стержень для текущего направления
' эксцентриситета. При отсутствии арматуры используется край бетонного контура,
' чтобы расчет не падал, но такая модель требует отдельной инженерной проверки.
Private Function RebarProjectionForMoment(ByVal section As CSectionModel, ByRef plane As TPlaneResult) As Double
    If section Is Nothing Then
        RebarProjectionForMoment = ConcreteFaceProjectionForMoment(plane)
        Exit Function
    End If
    If section.RebarCount <= 0 Then
        RebarProjectionForMoment = ConcreteFaceProjectionForMoment(plane)
        Exit Function
    End If

    Dim chooseNegativeSide As Boolean
    chooseNegativeSide = (plane.Eccentricity >= 0#)

    Dim i As Long
    Dim projectionValue As Double
    For i = 1 To section.RebarCount
        projectionValue = plane.NormalX * (section.RebarX(i) - plane.CentroidX) + _
            plane.NormalY * (section.RebarY(i) - plane.CentroidY)
        If i = 1 Then
            RebarProjectionForMoment = projectionValue
        ElseIf chooseNegativeSide Then
            If projectionValue < RebarProjectionForMoment Then RebarProjectionForMoment = projectionValue
        Else
            If projectionValue > RebarProjectionForMoment Then RebarProjectionForMoment = projectionValue
        End If
    Next i
End Function

' Возвращает край бетонного сечения на стороне, противоположной направлению
' сжатой равнодействующей. Для СП 35 это грань, относительно которой считается
' момент в коэффициенте длительности.
Private Function ConcreteFaceProjectionForMoment(ByRef plane As TPlaneResult) As Double
    If plane.Eccentricity >= 0# Then
        ConcreteFaceProjectionForMoment = plane.ConcreteMinProjection
    Else
        ConcreteFaceProjectionForMoment = plane.ConcreteMaxProjection
    End If
End Function

' Пересчитывает момент из центра приведенного сечения к точке с проекцией p.
' Сжатая сила nCompression передается положительным числом.
Private Function MomentAboutProjection(ByVal momentValue As Double, ByVal nCompression As Double, _
        ByVal projectionFromCentroid As Double) As Double
    MomentAboutProjection = momentValue - nCompression * projectionFromCentroid
End Function

' Коэффициент длительности СП 63 считается не по моменту в центре сечения, а по
' моментам M1 и Ml1 относительно центра выбранного стержня арматуры.
Private Function PhiLSP63(ByVal phiLMode As String, ByVal section As CSectionModel, _
        ByVal nCompression As Double, ByVal sustainedCompression As Double, _
        ByRef plane As TPlaneResult) As Double
    Dim rebarProjection As Double
    rebarProjection = RebarProjectionForMoment(section, plane)
    PhiLSP63 = PhiLFromMomentRatio(phiLMode, _
        MomentAboutProjection(plane.Moment, nCompression, rebarProjection), _
        MomentAboutProjection(plane.SustainedMoment, sustainedCompression, rebarProjection))
End Function

' ----------------------------- СП 63 --------------------------------------

' Расчет по СП 63: eta = 1/(1-N/Ncr), Ncr = pi^2*D/l0^2,
' D = kb*Eb*I + ks*Es*Is. Ветвь работает для произвольного дискретного
' сечения, потому что I, Is и h уже получены из CSectionModel.
Private Sub CalculateSP63(ByVal section As CSectionModel, ByVal materialProvider As CMaterialModelProvider, _
        ByVal valueSet As EMaterialValueSet, ByVal nCompression As Double, _
        ByVal sustainedCompression As Double, ByVal systemType As String, _
        ByVal phiLMode As String, ByVal ksValue As Double, ByVal deltaEMin As Double, _
        ByVal deltaEMax As Double, ByRef plane1 As TPlaneResult, ByRef plane2 As TPlaneResult)
    Dim eb As Double
    Dim es As Double
    eb = materialProvider.ConcreteElasticModulusFromValueSet(valueSet)
    es = materialProvider.SteelElasticModulusFromValueSet(valueSet)
    If eb <= 0# Or es <= 0# Then RaiseInput "Для расчета устойчивости СП 63 нужны положительные Eb и Es."
    CalculateSP63Plane section, nCompression, sustainedCompression, eb, es, ksValue, systemType, phiLMode, deltaEMin, deltaEMax, plane1, "1"
    CalculateSP63Plane section, nCompression, sustainedCompression, eb, es, ksValue, systemType, phiLMode, deltaEMin, deltaEMax, plane2, "2"
    mBranch = CODE_SP63
End Sub

' Считает одну главную плоскость по формулам СП 63.
Private Sub CalculateSP63Plane(ByVal section As CSectionModel, ByVal nCompression As Double, _
        ByVal sustainedCompression As Double, ByVal eb As Double, ByVal es As Double, _
        ByVal ksValue As Double, ByVal systemType As String, ByVal phiLMode As String, ByVal deltaEMin As Double, _
        ByVal deltaEMax As Double, ByRef plane As TPlaneResult, ByVal planeLabel As String)
    EnsurePositivePlaneGeometry plane, CODE_SP63, planeLabel

    plane.Eccentricity = InitialEccentricity(plane.Moment, nCompression, _
        plane.AccidentalEccentricity, systemType)
    plane.PhiL = PhiLSP63(phiLMode, section, nCompression, sustainedCompression, plane)

    Dim deltaE As Double
    deltaE = Abs(plane.Eccentricity) / plane.Depth
    If deltaEMin > 0# Then deltaE = GeomMax(deltaE, deltaEMin)
    If deltaEMax > 0# Then deltaE = GeomMin(deltaE, deltaEMax)

    Dim kb As Double
    kb = 0.15 / (plane.PhiL * (0.3 + deltaE))
    plane.DeltaValue = deltaE
    plane.Kb = kb
    plane.Ks = ksValue
    plane.StiffnessD = kb * eb * plane.InertiaConcrete + ksValue * es * plane.InertiaSteel
    plane.NcrForce = GEOM_PI * GEOM_PI * plane.StiffnessD / (plane.EffectiveLength * plane.EffectiveLength)
    plane.CriticalForce = plane.NcrForce
    plane.ReserveFactor = ReserveByCriticalForce(plane.NcrForce, nCompression)
    If plane.Slenderness <= 14# Then
        plane.Eta = 1#
    Else
        plane.Eta = EtaByCriticalForce(plane.NcrForce, nCompression)
    End If
    plane.DesignMoment = DesignMomentWithEta(nCompression, plane.Eccentricity, plane.Eta)
    If plane.Slenderness <= 14# Then
        plane.CheckPassed = True
    Else
        plane.CheckPassed = (plane.ReserveFactor > 1# And plane.Eta > 0#)
    End If
    plane.BranchName = CODE_SP63
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
    CalculateSP35Plane concreteArea, totalSteelArea, nCompression, sustainedCompression, eb, es, rb, rsc, systemType, phiLMode, phiP, nOverNcrLimit, table721, plane1, "1"
    CalculateSP35Plane concreteArea, totalSteelArea, nCompression, sustainedCompression, eb, es, rb, rsc, systemType, phiLMode, phiP, nOverNcrLimit, table721, plane2, "2"
    mBranch = CombinedSP35Branch(plane1, plane2)
End Sub

' Возвращает общий диагностический признак СП 35 без потери смысла, когда две
' главные плоскости попали в разные нормативные ветви расчета.
Private Function CombinedSP35Branch(ByRef plane1 As TPlaneResult, ByRef plane2 As TPlaneResult) As String
    Dim branch1 As String
    Dim branch2 As String
    branch1 = plane1.BranchName
    branch2 = plane2.BranchName

    If Not plane1.Applicable Then branch1 = vbNullString
    If Not plane2.Applicable Then branch2 = vbNullString

    If Len(branch1) = 0 And Len(branch2) = 0 Then
        CombinedSP35Branch = CODE_SP35
    ElseIf Len(branch1) = 0 Then
        CombinedSP35Branch = branch2
    ElseIf Len(branch2) = 0 Then
        CombinedSP35Branch = branch1
    ElseIf StrComp(branch1, branch2, vbTextCompare) = 0 Then
        CombinedSP35Branch = branch1
    Else
        CombinedSP35Branch = "SP35-mixed"
    End If
End Function

' Проверяет обязательные геометрические величины одной главной плоскости и
' пишет в StopReason фактические значения. Это особенно важно для импортированной
' AutoCAD-геометрии, где размеры элемента могут восстанавливаться по fallback-ам.
Private Sub EnsurePositivePlaneGeometry(ByRef plane As TPlaneResult, _
        ByVal codeText As String, ByVal planeLabel As String)
    If plane.EffectiveLength > 0# And plane.InertiaConcrete > 0# And _
            plane.Depth > 0# And plane.Radius > 0# Then Exit Sub

    RaiseInput "Для расчета устойчивости " & codeText & " в главной плоскости " & _
        planeLabel & " нужны положительные l0, h, I и i. Сейчас: l0=" & _
        StabilityNumber(plane.EffectiveLength) & " мм; h=" & _
        StabilityNumber(plane.Depth) & " мм; I=" & _
        StabilityNumber(plane.InertiaConcrete) & " мм4; i=" & _
        StabilityNumber(plane.Radius) & " мм."
End Sub

' Форматирует диагностические числа без длинных хвостов, чтобы popup с причиной
' InputErr оставался читаемым для пользователя.
Private Function StabilityNumber(ByVal value As Double) As String
    StabilityNumber = Replace$(Format$(value, "0.############"), ",", ".")
End Function

' Считает одну главную плоскость по СП 35: табличная ветвь при ec/r <= 1,
' иначе eta/Ncr. Таблица 7.21 передается с Config и не зашита в код.
Private Sub CalculateSP35Plane(ByVal concreteArea As Double, ByVal totalSteelArea As Double, _
        ByVal nCompression As Double, ByVal sustainedCompression As Double, ByVal eb As Double, ByVal es As Double, _
        ByVal rb As Double, ByVal rsc As Double, ByVal systemType As String, ByVal phiLMode As String, _
        ByVal phiP As Double, ByVal nOverNcrLimit As Double, ByVal table721 As Variant, _
        ByRef plane As TPlaneResult, ByVal planeLabel As String)
    EnsurePositivePlaneGeometry plane, CODE_SP35, planeLabel

    plane.Eccentricity = InitialEccentricity(plane.Moment, nCompression, _
        plane.AccidentalEccentricity, systemType)

    Dim coreDistance As Double
    coreDistance = CoreDistanceForEccentricity(plane)
    If coreDistance <= SMALL_VALUE Then RaiseInput "Для расчета устойчивости СП 35 нужно положительное ядровое расстояние r."
    plane.CoreDistance = coreDistance

    Dim q As Double
    q = Abs(plane.Eccentricity) / coreDistance
    If q <= 1# Then
        CalculateSP35TablePlane concreteArea, totalSteelArea, nCompression, sustainedCompression, rb, rsc, table721, q, plane
    Else
        CalculateSP35EtaPlane nCompression, sustainedCompression, eb, es, rb, phiLMode, phiP, nOverNcrLimit, plane
    End If
End Sub

' Табличная ветвь СП 35 для ec/r <= 1.
Private Sub CalculateSP35TablePlane(ByVal concreteArea As Double, ByVal totalSteelArea As Double, _
        ByVal nCompression As Double, ByVal sustainedCompression As Double, ByVal rb As Double, ByVal rsc As Double, _
        ByVal table721 As Variant, ByVal q As Double, _
        ByRef plane As TPlaneResult)
    Dim slenderness As Double
    slenderness = plane.Slenderness

    Dim phiM As Double
    Dim phiLTable As Double
    InterpolateSP35Table table721, slenderness, q, phiM, phiLTable
    plane.PhiL = phiLTable
    plane.PhiM = phiM
    plane.PhiLTable = phiLTable

    Dim nUltimate As Double
    Dim phiValue As Double
    phiValue = PhiForSustainedCompression(phiM, plane.PhiL, nCompression, sustainedCompression)
    plane.PhiValue = phiValue
    nUltimate = phiValue * (rb * SP35ConcreteAreaForTableNult(concreteArea, totalSteelArea) + _
        rsc * totalSteelArea)

    plane.NultimateForce = nUltimate
    plane.CriticalForce = plane.NultimateForce
    plane.ReserveFactor = ReserveByCriticalForce(nUltimate, nCompression)
    plane.Eta = 1#
    plane.StiffnessD = 0#
    plane.DesignMoment = nCompression * plane.Eccentricity
    plane.CheckPassed = (plane.ReserveFactor + PASS_TOLERANCE >= 1#)
    mBranch = "SP35-table"
    plane.BranchName = "SP35-table"
End Sub

' Возвращает бетонную площадь для формулы Nult табличной ветви СП 35.
' Приведенное сечение уже строится через (n - 1)As и этим определяет центр,
' оси и ядровое расстояние. Здесь работает другая нормативная поправка:
' если суммарная площадь стали больше 3% gross-площади бетона, бетонный
' вклад в Nult уменьшается на As, чтобы не завышать сжатую бетонную часть.
Private Function SP35ConcreteAreaForTableNult(ByVal concreteArea As Double, _
        ByVal totalSteelArea As Double) As Double
    If concreteArea <= SMALL_VALUE Then RaiseInput "Для табличной ветви СП 35 нужна положительная площадь бетона."

    SP35ConcreteAreaForTableNult = concreteArea
    If totalSteelArea <= SMALL_VALUE Then Exit Function

    If totalSteelArea / concreteArea > SP35_REBAR_AREA_LIMIT_RATIO Then
        SP35ConcreteAreaForTableNult = concreteArea - totalSteelArea
        If SP35ConcreteAreaForTableNult <= SMALL_VALUE Then _
            RaiseInput "Площадь стали больше или равна gross-площади бетона; формула Nult СП 35 неприменима."
    End If
End Function

' Рассчитывает ветвь СП 35 для ec/r > 1: phi_l по моментам относительно
' требуемой грани, delta, D, Ncr, eta и запас по условию N <= k*Ncr.
Private Sub CalculateSP35EtaPlane(ByVal nCompression As Double, ByVal sustainedCompression As Double, _
        ByVal eb As Double, ByVal es As Double, _
        ByVal rb As Double, ByVal phiLMode As String, ByVal phiP As Double, _
        ByVal nOverNcrLimit As Double, ByRef plane As TPlaneResult)
    Dim faceProjection As Double
    faceProjection = ConcreteFaceProjectionForMoment(plane)
    plane.PhiL = PhiLSP35Eta(phiLMode, MomentAboutProjection(plane.Moment, nCompression, faceProjection), _
        MomentAboutProjection(plane.SustainedMoment, sustainedCompression, faceProjection), _
        plane.Eccentricity, plane.Depth)

    Dim deltaValue As Double
    deltaValue = Abs(plane.Eccentricity) / plane.Depth
    Dim deltaMin As Double
    deltaMin = 0.5 - 0.01 * (plane.EffectiveLength / plane.Depth) - 0.01 * rb
    deltaValue = GeomMax(deltaValue, deltaMin)

    Dim phiPValue As Double
    phiPValue = phiP
    plane.PhiP = phiPValue
    Dim stiffnessFactor As Double
    stiffnessFactor = 0.11 / (0.1 + deltaValue / phiPValue) + 0.1
    plane.DeltaValue = deltaValue
    plane.StiffnessD = 6.4 * eb * (plane.InertiaConcrete / plane.PhiL * stiffnessFactor + _
        (es / eb) * plane.InertiaSteel)
    plane.NcrForce = plane.StiffnessD / (plane.EffectiveLength * plane.EffectiveLength)
    plane.CriticalForce = plane.NcrForce
    plane.ReserveFactor = SP35EtaReserveByLimit(plane.NcrForce, nCompression, nOverNcrLimit)
    plane.Eta = EtaByCriticalForce(plane.NcrForce, nCompression)
    plane.DesignMoment = DesignMomentWithEta(nCompression, plane.Eccentricity, plane.Eta)
    plane.CheckPassed = (plane.ReserveFactor + PASS_TOLERANCE >= 1# And plane.Eta > 0#)
    mBranch = "SP35-eta"
    plane.BranchName = "SP35-eta"
End Sub

' Возвращает запас для специального условия СП 35 eta-ветви:
' N/Ncr <= Stability.SP35.NOverNcrLimit. В Results и batch summary
' показывается именно проверочный запас limit*Ncr/N, а не сырое Ncr/N.
Private Function SP35EtaReserveByLimit(ByVal ncrForce As Double, ByVal nCompression As Double, _
        ByVal nOverNcrLimit As Double) As Double
    If nOverNcrLimit <= 0# Then Exit Function
    SP35EtaReserveByLimit = ReserveByCriticalForce(ncrForce, nCompression) * nOverNcrLimit
End Function

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

    Dim shortRatio As Double
    shortRatio = 1# - sustainedRatio

    Dim denominator As Double
    denominator = sustainedRatio * (phiM / phiL) + shortRatio
    If denominator > SMALL_VALUE Then
        PhiForSustainedCompression = phiM / denominator
    Else
        RaiseInput "Для СП 35 получен некорректный знаменатель при учете постоянной/длительной продольной силы."
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
    mMoment1 = plane1.Moment
    mMoment2 = plane2.Moment
    mSustainedMoment1 = plane1.SustainedMoment
    mSustainedMoment2 = plane2.SustainedMoment
    mDesignMoment1 = plane1.DesignMoment
    mDesignMoment2 = plane2.DesignMoment
    mReserve1 = plane1.ReserveFactor
    mReserve2 = plane2.ReserveFactor
    mEta1 = plane1.Eta
    mEta2 = plane2.Eta
    mPhiL1 = plane1.PhiL
    mPhiL2 = plane2.PhiL
    mPhiLTable1 = plane1.PhiLTable
    mPhiLTable2 = plane2.PhiLTable
    mPhiM1 = plane1.PhiM
    mPhiM2 = plane2.PhiM
    mPhiValue1 = plane1.PhiValue
    mPhiValue2 = plane2.PhiValue
    mPhiP1 = plane1.PhiP
    mPhiP2 = plane2.PhiP
    mEccentricity1 = plane1.Eccentricity
    mEccentricity2 = plane2.Eccentricity
    mAccidentalEccentricity1 = plane1.AccidentalEccentricity
    mAccidentalEccentricity2 = plane2.AccidentalEccentricity
    mEffectiveLength1 = plane1.EffectiveLength
    mEffectiveLength2 = plane2.EffectiveLength
    mSlenderness1 = plane1.Slenderness
    mSlenderness2 = plane2.Slenderness
    mDepth1 = plane1.Depth
    mDepth2 = plane2.Depth
    mRadius1 = plane1.Radius
    mRadius2 = plane2.Radius
    mCoreDistance1 = plane1.CoreDistance
    mCoreDistance2 = plane2.CoreDistance
    mDelta1 = plane1.DeltaValue
    mDelta2 = plane2.DeltaValue
    mKb1 = plane1.Kb
    mKb2 = plane2.Kb
    mKs1 = plane1.Ks
    mKs2 = plane2.Ks
    mStiffnessD1 = plane1.StiffnessD
    mStiffnessD2 = plane2.StiffnessD
    mPlaneBranch1 = plane1.BranchName
    mPlaneBranch2 = plane2.BranchName
    mPlaneApplicable1 = plane1.Applicable
    mPlaneApplicable2 = plane2.Applicable
    mPlanePassed1 = plane1.CheckPassed
    mPlanePassed2 = plane2.CheckPassed

    mCriticalForce = PositiveMinValue(plane1.CriticalForce, plane2.CriticalForce)
    mReserveFactor = PositiveMinValue(plane1.ReserveFactor, plane2.ReserveFactor)
    mCheckPassed = PlanePassedOrInactive(plane1) And PlanePassedOrInactive(plane2)
End Sub

' Возвращает моменты из главных осей обратно в глобальные Mx/My.
Private Sub TransformDesignMoments(ByVal props As CSectionPropertiesCalculator, ByVal moment1 As Double, ByVal moment2 As Double)
    Dim axis1X As Double
    Dim axis1Y As Double
    Dim axis2X As Double
    Dim axis2Y As Double
    props.PrincipalAxisDirection 1, axis1X, axis1Y
    props.PrincipalAxisDirection 2, axis2X, axis2Y
    mDesignMx = moment1 * axis1X + moment2 * axis2X
    mDesignMy = moment1 * axis1Y + moment2 * axis2Y
End Sub

' Считает модуль случайного эксцентриситета для выбранного СП и режима Config.
' В нормативном режиме AutoWithL используется геометрическая длина L:
' для СП 63 - ea = max(L/600; h/30; 10 мм), для СП 35 - ec,sl = L/400.
' Расчетная длина l0 = mu*L применяется только в режиме AutoWithMuL.
Private Function AccidentalEccentricity(ByVal modeText As String, ByVal userValue As Double, _
        ByVal userSettingName As String, ByVal geometricLength As Double, _
        ByVal effectiveLength As Double, ByVal depthValue As Double) As Double
    Dim normalizedMode As String
    normalizedMode = NormalizeAccidentalMode(modeText)
    If normalizedMode = ACC_USER Then
        AccidentalEccentricity = Abs(userValue)
        If AccidentalEccentricity <= SMALL_VALUE Then RaiseInput "На листе Config в строке " & userSettingName & " введите положительный случайный эксцентриситет в единицах INPUT Length для режима User."
        Exit Function
    End If

    Dim lengthForAccidental As Double
    If normalizedMode = ACC_AUTO_MUL Then
        lengthForAccidental = effectiveLength
    Else
        lengthForAccidental = geometricLength
    End If
    If lengthForAccidental <= SMALL_VALUE Then RaiseInput "На листе Config проверьте Stability.ElementLength и коэффициент расчетной длины: длина для случайного эксцентриситета должна быть больше нуля."

    If mCode = CODE_SP35 Then
        AccidentalEccentricity = lengthForAccidental / 400#
    Else
        AccidentalEccentricity = GeomMax(GeomMax(lengthForAccidental / 600#, depthValue / 30#), 10#)
    End If
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
    PhiLByMode = PhiLFromMomentRatio(phiLMode, totalMoment, sustainedMoment)
End Function

' Общая реализация формулы phi_l = 1 + Ml/M с верхним ограничением 2.
' Отношение берется со знаком. Нижняя граница 1 оставлена как защитное
' инженерное ограничение, чтобы длительная часть не повышала расчетную жесткость.
Private Function PhiLFromMomentRatio(ByVal phiLMode As String, ByVal totalMoment As Double, _
        ByVal sustainedMoment As Double) As Double
    Dim normalizedMode As String
    normalizedMode = UCase$(Trim$(phiLMode))
    If normalizedMode = UCase$(PHIL_2) Then
        PhiLFromMomentRatio = 2#
    ElseIf normalizedMode <> UCase$(PHIL_AUTO) Then
        RaiseInput "На листе Config в строке Stability.PhiLMode выберите Auto или PhiL2."
    ElseIf Abs(totalMoment) <= SMALL_VALUE Then
        If Abs(sustainedMoment) <= SMALL_VALUE Then PhiLFromMomentRatio = 1# Else PhiLFromMomentRatio = 2#
    Else
        PhiLFromMomentRatio = 1# + sustainedMoment / totalMoment
        If PhiLFromMomentRatio > 2# Then PhiLFromMomentRatio = 2#
        If PhiLFromMomentRatio < 1# Then PhiLFromMomentRatio = 1#
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

' Выбирает ядровое расстояние с той стороны сечения, куда направлен
' расчетный эксцентриситет. Для несимметричного сечения положительная и
' отрицательная стороны могут иметь разные r.
Private Function CoreDistanceForEccentricity(ByRef plane As TPlaneResult) As Double
    If plane.Eccentricity < 0# Then
        CoreDistanceForEccentricity = plane.CoreDistanceNegative
    Else
        CoreDistanceForEccentricity = plane.CoreDistancePositive
    End If
End Function

' Интерполирует таблицу 7.21 СП 35, переданную с листа Config. В книге
' таблица сохранена в человекочитаемом виде: l0/b, l0/d, l0/i, далее phi_m
' для нескольких ec/r и табличный phi_l. Расчет использует l0/i и phi-колонки.
Private Sub InterpolateSP35Table(ByVal table721 As Variant, ByVal slenderness As Double, ByVal q As Double, _
        ByRef phiM As Double, ByRef phiL As Double)
    If Len(mSP35TableInputError) > 0 Then RaiseInput mSP35TableInputError

    Dim invalidRow As Long, invalidColumn As Long, reason As String
    If Not TryValidateSP35Table721(table721, invalidRow, invalidColumn, reason) Then
        If invalidRow > 0 Then reason = "Строка " & CStr(invalidRow) & ", колонка " & _
            CStr(invalidColumn) & ": " & reason
        RaiseInput "Таблица 7.21 СП 35: " & reason
    End If

    Dim rowCount As Long
    rowCount = UBound(table721, 1)

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
End Sub

' Проверяет расчетные колонки таблицы без обращения к Excel и без скрытых
' нормативных defaults. Возвращает первую неверную строку/колонку для адресной
' диагностики вызывающего слоя. Две справочные колонки l0/b и l0/d не читаются.
' Положительность phi и возрастающие узлы защищают интерполяцию и предельную
' силу; дополнительные ограничения вроде phi <= 1 здесь не навязываются.
Public Function TryValidateSP35Table721(ByVal table721 As Variant, ByRef invalidRow As Long, _
        ByRef invalidColumn As Long, ByRef reason As String) As Boolean
    invalidRow = 0: invalidColumn = 0: reason = vbNullString
    If Not IsArray(table721) Then
        reason = "Не передана таблица. Восстановите именованный диапазон rngSP35Table721 на Config."
        Exit Function
    End If
    On Error GoTo InvalidStructure
    Dim rowCount As Long, columnCount As Long, firstValue As Variant
    rowCount = UBound(table721, 1): columnCount = UBound(table721, 2)
    If LBound(table721, 1) <> 1 Or LBound(table721, 2) <> 1 Or _
            rowCount < 2 Or columnCount < 8 Then GoTo InvalidStructure
    firstValue = table721(2, 3)
    On Error GoTo 0

    Dim row As Long, column As Long, value As Double, previousNode As Double
    For row = 2 To rowCount
        For column = 3 To 8
            invalidRow = row: invalidColumn = column
            If Not SP35TablePositiveNumber(table721(row, column), value) Then
                If column = 3 Then
                    reason = "Гибкость l0/i должна быть числом больше нуля. Исправьте значение в таблице."
                Else
                    reason = "Коэффициент " & SP35TableColumnName(column) & _
                        " должен быть числом больше нуля. Заполните ячейку корректным коэффициентом."
                End If
                Exit Function
            End If
            If column = 3 Then
                If row > 2 And value <= previousNode Then
                    reason = "Узлы l0/i должны строго возрастать сверху вниз. Исправьте эту строку или предыдущий узел."
                    Exit Function
                End If
                previousNode = value
            End If
        Next column
    Next row
    invalidRow = 0: invalidColumn = 0
    TryValidateSP35Table721 = True
    Exit Function
InvalidStructure:
    invalidRow = 0: invalidColumn = 0
    reason = "Нужен единый прямоугольный диапазон с заголовком, строками данных и восемью колонками: " & _
        "l0/b, l0/d, l0/i, phi_m(0), phi_m(0.25), phi_m(0.5), phi_m(1), phi_l. Восстановите структуру таблицы."
End Function

' Отличает пустоту, ошибку формулы и нечисловое значение от допустимого
' положительного коэффициента. Числовые строки остаются допустимыми;
' логические значения и даты не трактуются как инженерные числа.
Private Function SP35TablePositiveNumber(ByVal rawValue As Variant, ByRef value As Double) As Boolean
    On Error GoTo InvalidValue
    If IsError(rawValue) Or IsEmpty(rawValue) Or IsNull(rawValue) Then Exit Function
    If IsObject(rawValue) Or IsArray(rawValue) Then Exit Function
    If VarType(rawValue) = vbBoolean Or VarType(rawValue) = vbDate Then Exit Function
    If Not IsNumeric(rawValue) Then Exit Function
    value = CDbl(rawValue)
    SP35TablePositiveNumber = (value > 0#)
    Exit Function
InvalidValue:
End Function

' Называет неверный коэффициент по заголовку таблицы, чтобы сообщение
' оставалось понятным и при автономном расчете без Excel-адреса.
Private Function SP35TableColumnName(ByVal column As Long) As String
    Select Case column
        Case 4: SP35TableColumnName = "phi_m при ec/r = 0"
        Case 5: SP35TableColumnName = "phi_m при ec/r = 0.25"
        Case 6: SP35TableColumnName = "phi_m при ec/r = 0.5"
        Case 7: SP35TableColumnName = "phi_m при ec/r = 1"
        Case 8: SP35TableColumnName = "phi_l"
    End Select
End Function

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

' Интерполяция phi_m по отношению ec/r.
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

' Нормализует режим случайного эксцентриситета. В Excel показывается
' AutoWithμL, но код также принимает AutoWithMuL, чтобы не зависеть от способа
' ввода греческой буквы mu.
Private Function NormalizeAccidentalMode(ByVal value As String, _
        Optional ByVal settings As CSystemSettingsReader = Nothing) As String
    Dim normalized As String
    normalized = UCase$(Trim$(value))
    normalized = Replace$(normalized, ChrW$(&H3BC), "MU")
    normalized = Replace$(normalized, ChrW$(&H39C), "MU")

    Select Case normalized
        Case "USER"
            NormalizeAccidentalMode = ACC_USER
        Case "AUTOWITHMUL"
            NormalizeAccidentalMode = ACC_AUTO_MUL
        Case "AUTOWITHL"
            NormalizeAccidentalMode = ACC_AUTO_L
        Case Else
            RaiseSettingInput "Stability.AccidentalEccentricityMode", "Неизвестный режим случайного эксцентриситета.", _
                "Выберите AutoWithL, AutoWithMuL или User.", settings
    End Select
End Function

' Нормализует настройку выбора плоскостей для случайного эксцентриситета.
Private Function NormalizeAccidentalPlanes(ByVal value As String, _
        Optional ByVal settings As CSystemSettingsReader = Nothing) As String
    Select Case UCase$(Trim$(value))
        Case UCase$(ACC_PLANES_BOTH)
            NormalizeAccidentalPlanes = ACC_PLANES_BOTH
        Case UCase$(ACC_PLANES_ONLY_MOMENT)
            NormalizeAccidentalPlanes = ACC_PLANES_ONLY_MOMENT
        Case Else
            RaiseSettingInput "Stability.AccidentalEccentricityPlanes", "Неизвестный выбор главных плоскостей.", _
                "Выберите OnlyMomentPlane или BothPlanes.", settings
    End Select
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

' Добавляет к причине ключ и текущее место Config. При автономном вызове
' без reader сохраняет ключ и действие, не придумывая адрес ячейки.
Private Sub RaiseSettingInput(ByVal key As String, ByVal reason As String, _
        ByVal correction As String, ByVal settings As CSystemSettingsReader)
    Dim message As String
    If settings Is Nothing Then
        message = "Настройка " & key & ": " & reason & " " & correction
    Else
        message = settings.InputErrorMessage(key, reason, correction)
    End If
    RaiseInput message
End Sub

' Сохраняет конкретную причину InputErr для автономного вызова Calculator.
' Программные исключения остаются отдельными от этого контролируемого номера.
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

' Объясняет готовую инженерную проверку, не определяя статус по тексту.
' При отказе перечисляет только не прошедшие главные плоскости и их безразмерные
' запас/eta; физические условия остаются в CalculateSP63/CalculateSP35.
Private Function StabilityCheckComment() As String
    If mCheckPassed Then
        If mCode = CODE_SP35 Then
            StabilityCheckComment = "Проверка устойчивости по СП 35 выполнена."
        Else
            StabilityCheckComment = "Проверка устойчивости по СП 63 выполнена."
        End If
        Exit Function
    End If
    StabilityCheckComment = "Проверка устойчивости не проходит"
    Dim planeDetails As String
    If mPlaneApplicable1 And Not mPlanePassed1 Then
        planeDetails = "в главной плоскости 1: запас = " & CStr(mReserve1) & _
            "; коэффициент увеличения момента eta = " & CStr(mEta1)
    End If
    If mPlaneApplicable2 And Not mPlanePassed2 Then
        If Len(planeDetails) > 0 Then planeDetails = planeDetails & "; "
        planeDetails = planeDetails & "в главной плоскости 2: запас = " & CStr(mReserve2) & _
            "; коэффициент увеличения момента eta = " & CStr(mEta2)
    End If
    If Len(planeDetails) > 0 Then StabilityCheckComment = StabilityCheckComment & " " & planeDetails
    StabilityCheckComment = StabilityCheckComment & "."
End Function

' Формирует инженерную meta собственной проверки устойчивости. Ни writer,
' ни display-policy не повторяют эти условия и не придумывают причину по
' Branch. Отсутствующие программные зависимости отличаются от ошибочного Config.
Public Property Get ResultMeta() As CResultMeta
    Dim meta As CResultMeta
    Set meta = New CResultMeta
    If Not mResultInitialized Then
        meta.SetResult rsInternalError, rcInternalError, rkStability, "Результат проверки устойчивости не заполнен.", _
            vbNullString, True, False
    ElseIf mInternalError Then
        meta.SetResult rsInternalError, rcInternalError, rkStability, mStopReason, _
            vbNullString, True, mCalculationStarted
    ElseIf Not mInputValid Then
        meta.SetResult rsInvalidInput, rcInvalidInput, rkStability, mStopReason, _
            vbNullString, True, mCalculationStarted
    ElseIf Not mApplicable Then
        meta.SetNotApplicable rkStability, "Устойчивость не проверяется: в сочетании нет сжатия."
    ElseIf mCheckPassed Then
        meta.SetResult rsSuccess, rcCheckPassed, rkStability, StabilityCheckComment()
    Else
        meta.SetResult rsCheckFailed, rcCheckFailed, rkStability, StabilityCheckComment()
    End If
    Set ResultMeta = meta
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

' Возвращает угол главных осей приведенного сечения, использованный в расчете.
' Batch применяет его только для обновления итогового Results snapshot после
' общего фильтра практически нулевых моментов.
Public Property Get PrincipalAngleRad() As Double
    PrincipalAngleRad = mPrincipalAngleRad
End Property

Public Property Get SustainedN() As Double
    SustainedN = mSustainedN
End Property

Public Property Get Moment1() As Double
    Moment1 = mMoment1
End Property

Public Property Get Moment2() As Double
    Moment2 = mMoment2
End Property

Public Property Get SustainedMoment1() As Double
    SustainedMoment1 = mSustainedMoment1
End Property

Public Property Get SustainedMoment2() As Double
    SustainedMoment2 = mSustainedMoment2
End Property

Public Property Get DesignMoment1() As Double
    DesignMoment1 = mDesignMoment1
End Property

Public Property Get DesignMoment2() As Double
    DesignMoment2 = mDesignMoment2
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

Public Property Get PhiLTable1() As Double
    PhiLTable1 = mPhiLTable1
End Property

Public Property Get PhiLTable2() As Double
    PhiLTable2 = mPhiLTable2
End Property

Public Property Get PhiM1() As Double
    PhiM1 = mPhiM1
End Property

Public Property Get PhiM2() As Double
    PhiM2 = mPhiM2
End Property

Public Property Get PhiValue1() As Double
    PhiValue1 = mPhiValue1
End Property

Public Property Get PhiValue2() As Double
    PhiValue2 = mPhiValue2
End Property

Public Property Get PhiP1() As Double
    PhiP1 = mPhiP1
End Property

Public Property Get PhiP2() As Double
    PhiP2 = mPhiP2
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

Public Property Get Slenderness1() As Double
    Slenderness1 = mSlenderness1
End Property

Public Property Get Slenderness2() As Double
    Slenderness2 = mSlenderness2
End Property

Public Property Get Depth1() As Double
    Depth1 = mDepth1
End Property

Public Property Get Depth2() As Double
    Depth2 = mDepth2
End Property

Public Property Get Radius1() As Double
    Radius1 = mRadius1
End Property

Public Property Get Radius2() As Double
    Radius2 = mRadius2
End Property

Public Property Get CoreDistance1() As Double
    CoreDistance1 = mCoreDistance1
End Property

Public Property Get CoreDistance2() As Double
    CoreDistance2 = mCoreDistance2
End Property

Public Property Get Delta1() As Double
    Delta1 = mDelta1
End Property

Public Property Get Delta2() As Double
    Delta2 = mDelta2
End Property

Public Property Get Kb1() As Double
    Kb1 = mKb1
End Property

Public Property Get Kb2() As Double
    Kb2 = mKb2
End Property

Public Property Get Ks1() As Double
    Ks1 = mKs1
End Property

Public Property Get Ks2() As Double
    Ks2 = mKs2
End Property

Public Property Get StiffnessD1() As Double
    StiffnessD1 = mStiffnessD1
End Property

Public Property Get StiffnessD2() As Double
    StiffnessD2 = mStiffnessD2
End Property

Public Property Get PlaneBranch1() As String
    PlaneBranch1 = mPlaneBranch1
End Property

Public Property Get PlaneBranch2() As String
    PlaneBranch2 = mPlaneBranch2
End Property

Public Property Get PlaneApplicable1() As Boolean
    PlaneApplicable1 = mPlaneApplicable1
End Property

Public Property Get PlaneApplicable2() As Boolean
    PlaneApplicable2 = mPlaneApplicable2
End Property

Public Property Get PlanePassed1() As Boolean
    PlanePassed1 = mPlanePassed1
End Property

Public Property Get PlanePassed2() As Boolean
    PlanePassed2 = mPlanePassed2
End Property
