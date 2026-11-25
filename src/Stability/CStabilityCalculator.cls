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
' Расчёт продольного изгиба и устойчивости
' ==========================================================================
' По готовой модели, материалам профиля и нагрузкам сочетания определяет
' усилия N/Mx/My с учётом продольного изгиба и величины для сводной таблицы.
' Проверяет допустимость своих настроек до расчёта. Адрес ошибки Config
' может быть добавлен по уже прочитанным настройкам без нового чтения листов.
' Геометрические характеристики получает через CSectionPropertiesCalculator.
' Нормативные формулы выполняют CStabilitySP35Calculator и CStabilitySP63Calculator;
' подготовка главных плоскостей, итоговые статусы и перевод моментов остаются общими.

Private mSP35 As New CStabilitySP35Calculator ' Нормативные ветви СП 35.
Private mSP63 As New CStabilitySP63Calculator ' Нормативная ветвь СП 63.


Private mCode As String ' Нормативная ветвь, выбранная настройкой Stability.Code.
Private mBranch As String ' Название применённого способа проверки устойчивости для отчёта и диагностики.
Private mApplicable As Boolean ' False, если сжатия нет и устойчивость не требуется.
Private mInputValid As Boolean ' False при нехватке обязательных исходных данных.
Private mResultInitialized As Boolean ' Был вызван расчет; новый объект не является готовым результатом.
Private mCalculationStarted As Boolean ' Выполнялась собственная инженерная проверка после валидации.
Private mInternalError As Boolean ' Признак отсутствующих обязательных объектов или неожиданной ошибки программы; не относится к пользовательскому вводу.
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
Private mPhiLTable1 As Double ' Табличный phi_l СП 35 первой плоскости для ветви |ec|/r <= 1.
Private mPhiLTable2 As Double ' Табличный phi_l СП 35 второй плоскости для ветви |ec|/r <= 1.
Private mPhiM1 As Double ' Табличный phi_m СП 35 первой плоскости.
Private mPhiM2 As Double ' Табличный phi_m СП 35 второй плоскости.
Private mPhiValue1 As Double ' Итоговый phi СП 35 первой плоскости в ветви |ec| <= r.
Private mPhiValue2 As Double ' Итоговый phi СП 35 второй плоскости в ветви |ec| <= r.
Private mPhiP1 As Double ' Коэффициент phi_p СП 35 первой плоскости в ветви |ec| > r.
Private mPhiP2 As Double ' Коэффициент phi_p СП 35 второй плоскости в ветви |ec| > r.
Private mEccentricity1 As Double ' Расчетный эксцентриситет первой плоскости, мм.
Private mEccentricity2 As Double ' Расчетный эксцентриситет второй плоскости, мм.
Private mAccidentalEccentricity1 As Double ' Случайный эксцентриситет первой плоскости, мм.
Private mAccidentalEccentricity2 As Double ' Случайный эксцентриситет второй плоскости, мм.
Private mEffectiveLength1 As Double ' l0 первой плоскости, мм.
Private mEffectiveLength2 As Double ' l0 второй плоскости, мм.
Private mSlenderness1 As Double ' Гибкость l0/i для плоскости 1: изгиб вокруг главной оси 1.
Private mSlenderness2 As Double ' Гибкость l0/i для плоскости 2: изгиб вокруг главной оси 2.
Private mDepth1 As Double ' Бетонный габарит h1 вдоль оси 2 при изгибе вокруг оси 1, мм.
Private mDepth2 As Double ' Бетонный габарит h2 вдоль оси 1 при изгибе вокруг оси 2, мм.
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
Private mConcreteElasticModulus As Double ' Модуль упругости бетона для обеих плоскостей, МПа.
Private mSteelElasticModulus As Double ' Модуль упругости арматуры для обеих плоскостей, МПа.
Private mConcreteInertia1 As Double, mConcreteInertia2 As Double ' Инерции бетона относительно расчётных осей, мм4.
Private mSteelInertia1 As Double, mSteelInertia2 As Double ' Инерции арматуры относительно расчётных осей, мм4.
Private mPlaneBranch1 As String ' Фактическая ветвь расчета первой плоскости.
Private mPlaneBranch2 As String ' Фактическая ветвь расчета второй плоскости.
Private mPlaneApplicable1 As Boolean ' True, если первая плоскость участвовала в проверке.
Private mPlaneApplicable2 As Boolean ' True, если вторая плоскость участвовала в проверке.
Private mPlanePassed1 As Boolean ' True, если первая плоскость прошла свою проверку.
Private mPlanePassed2 As Boolean ' True, если вторая плоскость прошла свою проверку.
Private mPreparedTransformedProperties As CSectionPropertiesCalculator ' Заранее подготовленные характеристики приведённого сечения для всех проверок запуска.
Private mPreparedConcreteProperties As CSectionPropertiesCalculator ' Только механическая геометрия; LC-dependent Ncr/eta не переиспользуются.

' Подключает подготовку одного запуска. Реальные Eb/Es и геометрия проверяются
' владельцем properties лишь в применимой ветви после проверки входных данных.
Friend Sub UseMechanicalPreparation(ByVal transformed As CSectionPropertiesCalculator, ByVal concrete As CSectionPropertiesCalculator)
    Set mPreparedTransformedProperties = transformed
    Set mPreparedConcreteProperties = concrete
End Sub

' Проверяет настройки устойчивости до геометрических вычислений и раннего
' выхода при отсутствии сжатия. Единицы уже переведены; адрес ошибки берёт
' из прочитанного Config. Нулевая граница delta отключает соответствующее
' ограничение без подстановки другого значения. НДС не решается.
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
    If code <> STABILITY_CODE_SP63 And code <> STABILITY_CODE_SP35 Then RaiseSettingInput "Stability.Code", _
        "Неизвестная методика устойчивости.", "Выберите SP63 или SP35.", settings
    If elementLength <= 0# Then RaiseSettingInput "Stability.ElementLength", _
        "Длина элемента должна быть положительной.", "Введите число больше нуля в единицах INPUT Length.", settings
    If mu1 <= 0# Then RaiseSettingInput "Stability.Mu1", _
        "Коэффициент расчетной длины плоскости 1 должен быть положительным.", "Задайте число больше нуля.", settings
    If mu2 <= 0# Then RaiseSettingInput "Stability.Mu2", _
        "Коэффициент расчетной длины плоскости 2 должен быть положительным.", "Задайте число больше нуля.", settings
    ValidateEffectiveLengthRange elementLength, mu1, "Stability.Mu1", settings
    ValidateEffectiveLengthRange elementLength, mu2, "Stability.Mu2", settings
    If StrComp(Trim$(systemType), STABILITY_SYSTEM_DETERMINATE, vbTextCompare) <> 0 And _
            StrComp(Trim$(systemType), STABILITY_SYSTEM_INDETERMINATE, vbTextCompare) <> 0 Then _
        RaiseSettingInput "Stability.SystemType", "Неизвестная расчетная схема.", _
            "Выберите Determinate или Indeterminate.", settings
    If zeroMomentSign1 <> 1# And zeroMomentSign1 <> -1# Then RaiseSettingInput "Stability.ZeroMomentEccentricitySign1", _
        "Направление задается только знаком.", "Введите 1 или -1 для главной плоскости 1.", settings
    If zeroMomentSign2 <> 1# And zeroMomentSign2 <> -1# Then RaiseSettingInput "Stability.ZeroMomentEccentricitySign2", _
        "Направление задается только знаком.", "Введите 1 или -1 для главной плоскости 2.", settings
    If StrComp(Trim$(phiLMode), STABILITY_PHIL_AUTO, vbTextCompare) <> 0 And _
            StrComp(Trim$(phiLMode), STABILITY_PHIL_2, vbTextCompare) <> 0 Then _
        RaiseSettingInput "Stability.PhiLMode", "Неизвестный режим учета длительности.", "Выберите Auto или PhiL2.", settings
    mode = NormalizeAccidentalMode(accidentalMode, settings)
    Call NormalizeAccidentalPlanes(accidentalPlanes, settings)
    If accidentalUserValue1 < 0# Or (mode = STABILITY_ACC_USER And accidentalUserValue1 <= STABILITY_SMALL_VALUE) Then _
        RaiseSettingInput "Stability.AccidentalEccentricityUser1", "Требуется модуль, а не знаковый эксцентриситет.", _
            "Введите положительное значение для User; в автоматическом режиме допустим также 0.", settings
    If accidentalUserValue2 < 0# Or (mode = STABILITY_ACC_USER And accidentalUserValue2 <= STABILITY_SMALL_VALUE) Then _
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
    If squaredLength <> squaredLength Or squaredLength > 1.797693134862315E+308 Then GoTo RangeFailed ' Excel может оставить бесконечность без VBA-ошибки; проверяем значение, а не полагаемся только на обработчик переполнения.
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
        Optional ByVal momentFilter As CLoadZeroFilter = Nothing, _
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
    nCompression = totalLoad.CompressionMagnitude(STABILITY_SMALL_FORCE)
    If nCompression <= STABILITY_SMALL_FORCE Then
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

    Dim plane1 As TStabilityPlane
    Dim plane2 As TStabilityPlane
    PreparePlanes section, props, mDesignMx, mDesignMy, sustainedLoad.InternalMx, sustainedLoad.InternalMy, _
        elementLength, mu1, mu2, systemType, zeroMomentSign1, zeroMomentSign2, _
        accidentalMode, accidentalPlanes, accidentalUserValue1, accidentalUserValue2, _
        momentFilter, plane1, plane2

    If mCode = STABILITY_CODE_SP63 Then
        CalculateSP63 section, ebForGeometry, esForGeometry, nCompression, sustainedLoad.SignedCompression, _
            systemType, phiLMode, sp63Ks, sp63DeltaEMin, sp63DeltaEMax, plane1, plane2
    Else
        CalculateSP35 section, materialProvider, valueSet, concreteProps.Area, props.RebarAreaTotal(section), _
            nCompression, sustainedLoad.SignedCompression, ebForGeometry, esForGeometry, systemType, phiLMode, _
            sp35PhiP, sp35NOverNcrLimit, sp35Table721, plane1, plane2
    End If

    StorePlaneResults plane1, plane2, ebForGeometry, esForGeometry
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
    mConcreteElasticModulus = 0#: mSteelElasticModulus = 0#
    mConcreteInertia1 = 0#: mConcreteInertia2 = 0#
    mSteelInertia1 = 0#: mSteelInertia2 = 0#
    mPlaneBranch1 = vbNullString: mPlaneBranch2 = vbNullString
    mPlaneApplicable1 = False: mPlaneApplicable2 = False
    mPlanePassed1 = False: mPlanePassed2 = False
End Sub

' Подготавливает две главные расчетные плоскости: моменты переводятся в
' главные центральные оси приведённого сечения. Эксцентриситеты устойчивости
' по СП 35 и СП 63 вычисляются относительно этих осей. Моменты уже пересчитаны
' относительно того же центра при подготовке сочетания. Плоскость 1
' связана с изгибом вокруг первой главной оси, плоскость 2 - вокруг второй.
' Ось 1 ближе к X, ось 2 повёрнута от неё на +90°. При совпадении с X/Y
' плоскость 1 соответствует изгибу Mx и габариту по Y, а плоскость 2 —
' изгибу My и габариту по X. Знаки эксцентриситетов следуют нормалям.
Private Sub PreparePlanes(ByVal section As CSectionModel, ByVal props As CSectionPropertiesCalculator, _
        ByVal mxValue As Double, ByVal myValue As Double, ByVal sustainedMx As Double, _
        ByVal sustainedMy As Double, ByVal elementLength As Double, ByVal mu1 As Double, _
        ByVal mu2 As Double, ByVal systemType As String, ByVal zeroMomentSign1 As Double, _
        ByVal zeroMomentSign2 As Double, ByVal accidentalMode As String, _
        ByVal accidentalPlanes As String, ByVal accidentalUserValue1 As Double, _
        ByVal accidentalUserValue2 As Double, ByVal momentFilter As CLoadZeroFilter, _
        ByRef plane1 As TStabilityPlane, ByRef plane2 As TStabilityPlane)
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
    If plane1.Radius > STABILITY_SMALL_VALUE Then plane1.Slenderness = plane1.EffectiveLength / plane1.Radius
    If plane2.Radius > STABILITY_SMALL_VALUE Then plane2.Slenderness = plane2.EffectiveLength / plane2.Radius
    plane1.Applicable = True
    plane2.Applicable = True

    ' В режиме OnlyMomentPlane случайный эксцентриситет добавляется только
    ' к плоскостям с моментом; BothPlanes разрешает его в обеих плоскостях.
    ' При чистом сжатии без моментов в обоих режимах применяются направления +1/-1.
    Dim hasMomentVector As Boolean
    Dim applyBothPlanes As Boolean
    hasMomentVector = (Abs(plane1.Moment) > STABILITY_SMALL_VALUE Or Abs(plane2.Moment) > STABILITY_SMALL_VALUE)
    applyBothPlanes = (NormalizeAccidentalPlanes(accidentalPlanes) = STABILITY_ACC_PLANES_BOTH)
    FillAccidentalAndInitialEccentricity plane1, zeroMomentSign1, systemType, _
        accidentalMode, accidentalUserValue1, "Stability.AccidentalEccentricityUser1", _
        (applyBothPlanes Or Not hasMomentVector Or Abs(plane1.Moment) > STABILITY_SMALL_VALUE)
    FillAccidentalAndInitialEccentricity plane2, zeroMomentSign2, systemType, _
        accidentalMode, accidentalUserValue2, "Stability.AccidentalEccentricityUser2", _
        (applyBothPlanes Or Not hasMomentVector Or Abs(plane2.Moment) > STABILITY_SMALL_VALUE)
End Sub

' Применяет общий инженерный фильтр к моментам одной главной плоскости.
' Сравнение выполняется по бетонному габариту именно этой плоскости, поэтому
' после поворота осей случайный остаток M1/M2 не меняет ветвь устойчивости.
Private Sub NormalizePlaneMoments(ByRef plane As TStabilityPlane, ByVal momentFilter As CLoadZeroFilter)
    If momentFilter Is Nothing Then Exit Sub
    plane.Moment = momentFilter.NormalizeMoment(plane.Moment, plane.Depth)
    plane.SustainedMoment = momentFilter.NormalizeMoment(plane.SustainedMoment, plane.Depth)
End Sub

' Назначает случайный и начальный эксцентриситет одной главной плоскости.
' Здесь сходятся пользовательский режим Auto/User, выбор L или mu*L и правило
' OnlyMomentPlane/BothPlanes для нулевых исходных моментов.
Private Sub FillAccidentalAndInitialEccentricity(ByRef plane As TStabilityPlane, _
        ByVal zeroMomentSign As Double, ByVal systemType As String, _
        ByVal accidentalMode As String, ByVal accidentalUserValue As Double, _
        ByVal accidentalUserSettingName As String, _
        ByVal allowZeroMomentAccidental As Boolean)
    If Abs(plane.Moment) <= STABILITY_SMALL_VALUE And Not allowZeroMomentAccidental Then
        plane.AccidentalEccentricity = 0#
        plane.Eccentricity = 0#
        Exit Sub
    End If

    Dim signValue As Double
    signValue = StabilitySignForEccentricity(plane.Moment, zeroMomentSign)
    plane.AccidentalEccentricity = signValue * AccidentalEccentricity( _
        accidentalMode, accidentalUserValue, accidentalUserSettingName, _
        plane.GeometricLength, plane.EffectiveLength, plane.Depth)
    plane.Eccentricity = StabilityInitialEccentricity(plane.Moment, 1#, plane.AccidentalEccentricity, systemType)
End Sub

' Сохраняет крайние точки бетонного контура относительно центра приведенного
' сечения. Эти значения нужны для h и ядрового расстояния: арматура не должна
' расширять внешний габарит сжатого элемента.
Private Sub StoreConcreteProjectionBounds(ByVal section As CSectionModel, _
        ByVal props As CSectionPropertiesCalculator, ByRef plane As TStabilityPlane)
    Dim minProjection As Double
    Dim maxProjection As Double
    props.CalculateProjection section, plane.NormalX, plane.NormalY, False, minProjection, maxProjection

    Dim centroidProjection As Double
    centroidProjection = plane.NormalX * props.CentroidX + plane.NormalY * props.CentroidY
    plane.ConcreteMinProjection = minProjection - centroidProjection
    plane.ConcreteMaxProjection = maxProjection - centroidProjection
End Sub


' ----------------------------- СП 63 --------------------------------------

' Расчет по СП 63: eta = 1/(1-N/Ncr), Ncr = pi^2*D/l0^2,
' D = kb*Eb*I + ks*Es*Is. Ветвь работает для произвольного дискретного
' сечения, потому что I, Is и h уже получены из CSectionModel.
Private Sub CalculateSP63(ByVal section As CSectionModel, ByVal eb As Double, ByVal es As Double, _
        ByVal nCompression As Double, _
        ByVal sustainedCompression As Double, ByVal systemType As String, _
        ByVal phiLMode As String, ByVal ksValue As Double, ByVal deltaEMin As Double, _
        ByVal deltaEMax As Double, ByRef plane1 As TStabilityPlane, ByRef plane2 As TStabilityPlane)
    mSP63.CalculateSP63 section, eb, es, nCompression, sustainedCompression, systemType, phiLMode, ksValue, deltaEMin, deltaEMax, plane1, plane2
    mBranch = STABILITY_CODE_SP63
End Sub


' ----------------------------- СП 35 --------------------------------------

' Расчет по СП 35. Для малых эксцентриситетов используется табличная ветвь
' phi по таблице 7.21, для больших - ветвь eta/Ncr по п. 7.54.
Private Sub CalculateSP35(ByVal section As CSectionModel, ByVal materialProvider As CMaterialModelProvider, _
        ByVal valueSet As EMaterialValueSet, ByVal concreteArea As Double, ByVal totalSteelArea As Double, _
        ByVal nCompression As Double, ByVal sustainedCompression As Double, _
        ByVal eb As Double, ByVal es As Double, ByVal systemType As String, ByVal phiLMode As String, ByVal phiP As Double, _
        ByVal nOverNcrLimit As Double, ByVal table721 As Variant, _
        ByRef plane1 As TStabilityPlane, ByRef plane2 As TStabilityPlane)
    mSP35.CalculateSP35 mBranch, mSP35TableInputError, section, materialProvider, valueSet, concreteArea, totalSteelArea, nCompression, sustainedCompression, eb, es, systemType, phiLMode, phiP, nOverNcrLimit, table721, plane1, plane2
End Sub


' -------------------------- Общие методы ----------------------------------

' Собирает итог по двум главным плоскостям и выбирает минимальный запас.
Private Sub StorePlaneResults(ByRef plane1 As TStabilityPlane, ByRef plane2 As TStabilityPlane, ByVal eb As Double, ByVal es As Double)
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
    mConcreteElasticModulus = eb
    mSteelElasticModulus = es
    mConcreteInertia1 = plane1.InertiaConcrete
    mConcreteInertia2 = plane2.InertiaConcrete
    mSteelInertia1 = plane1.InertiaSteel
    mSteelInertia2 = plane2.InertiaSteel
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
' Выбирает пользовательский ввод либо длину L/l0 для автоматического режима.
' Нормативную формулу выполняет калькулятор выбранного СП.
Private Function AccidentalEccentricity(ByVal modeText As String, ByVal userValue As Double, _
        ByVal userSettingName As String, ByVal geometricLength As Double, _
        ByVal effectiveLength As Double, ByVal depthValue As Double) As Double
    Dim normalizedMode As String
    normalizedMode = NormalizeAccidentalMode(modeText)
    If normalizedMode = STABILITY_ACC_USER Then
        AccidentalEccentricity = Abs(userValue)
        If AccidentalEccentricity <= STABILITY_SMALL_VALUE Then RaiseInput "На листе Config в строке " & userSettingName & " введите положительный случайный эксцентриситет в единицах INPUT Length для режима User."
        Exit Function
    End If

    Dim lengthForAccidental As Double
    If normalizedMode = STABILITY_ACC_AUTO_MUL Then
        lengthForAccidental = effectiveLength
    Else
        lengthForAccidental = geometricLength
    End If
    If lengthForAccidental <= STABILITY_SMALL_VALUE Then RaiseInput "На листе Config проверьте Stability.ElementLength и коэффициент расчетной длины: длина для случайного эксцентриситета должна быть больше нуля."

    If mCode = STABILITY_CODE_SP35 Then
        AccidentalEccentricity = mSP35.AutomaticAccidentalEccentricity(lengthForAccidental)
    Else
        AccidentalEccentricity = mSP63.AutomaticAccidentalEccentricity(lengthForAccidental, depthValue)
    End If
End Function


' Проверяет расчетные колонки таблицы без обращения к Excel и без скрытых
' нормативных defaults. Возвращает первую неверную строку/колонку для адресной
' диагностики вызывающего слоя. Гибкость неотрицательна; коэффициенты положительны.
' Положительность phi и возрастающие узлы защищают интерполяцию и предельную
' силу; дополнительные ограничения вроде phi <= 1 здесь не навязываются.
Public Function TryValidateSP35Table721(ByVal table721 As Variant, ByRef invalidRow As Long, _
        ByRef invalidColumn As Long, ByRef reason As String) As Boolean
    TryValidateSP35Table721 = mSP35.TryValidateSP35Table721(table721, invalidRow, invalidColumn, reason)
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
Private Function PlanePassedOrInactive(ByRef plane As TStabilityPlane) As Boolean
    PlanePassedOrInactive = (Not plane.Applicable Or plane.CheckPassed)
End Function

' Нормализует режим случайного эксцентриситета. В Excel показывается
' AutoWithmuL, но код также принимает AutoWithMuL, чтобы не зависеть от способа
' ввода греческой буквы mu.
Private Function NormalizeAccidentalMode(ByVal value As String, _
        Optional ByVal settings As CSystemSettingsReader = Nothing) As String
    Dim normalized As String
    normalized = UCase$(Trim$(value))
    normalized = Replace$(normalized, ChrW$(&H3BC), "MU")
    normalized = Replace$(normalized, ChrW$(&H39C), "MU")

    Select Case normalized
        Case "USER"
            NormalizeAccidentalMode = STABILITY_ACC_USER
        Case "AUTOWITHMUL"
            NormalizeAccidentalMode = STABILITY_ACC_AUTO_MUL
        Case "AUTOWITHL"
            NormalizeAccidentalMode = STABILITY_ACC_AUTO_L
        Case Else
            RaiseSettingInput "Stability.AccidentalEccentricityMode", "Неизвестный режим случайного эксцентриситета.", _
                "Выберите AutoWithL, AutoWithMuL или User.", settings
    End Select
End Function

' Нормализует настройку выбора плоскостей для случайного эксцентриситета.
Private Function NormalizeAccidentalPlanes(ByVal value As String, _
        Optional ByVal settings As CSystemSettingsReader = Nothing) As String
    Select Case UCase$(Trim$(value))
        Case UCase$(STABILITY_ACC_PLANES_BOTH)
            NormalizeAccidentalPlanes = STABILITY_ACC_PLANES_BOTH
        Case UCase$(STABILITY_ACC_PLANES_ONLY_MOMENT)
            NormalizeAccidentalPlanes = STABILITY_ACC_PLANES_ONLY_MOMENT
        Case Else
            RaiseSettingInput "Stability.AccidentalEccentricityPlanes", "Неизвестный выбор главных плоскостей.", _
                "Выберите OnlyMomentPlane или BothPlanes.", settings
    End Select
End Function

' Нормализует имя нормативного блока.
Private Function NormalizeCode(ByVal value As String) As String
    Select Case UCase$(Trim$(value))
        Case "SP35", "СП35", "СП 35"
            NormalizeCode = STABILITY_CODE_SP35
        Case "SP63", "СП63", "СП 63"
            NormalizeCode = STABILITY_CODE_SP63
        Case Else
            NormalizeCode = Trim$(value)
    End Select
End Function

' Добавляет к ошибке имя настройки и фактический адрес Config.
' При отдельном программном вызове сохраняет имя и способ исправления
' без вымышленного адреса ячейки.
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

' Нормативная ветвь, выбранная настройкой Stability.Code.
Public Property Get Code() As String
    Code = mCode
End Property

' Название применённого способа проверки устойчивости для отчёта и диагностики.
Public Property Get Branch() As String
    Branch = mBranch
End Property

' False, если сжатия нет и устойчивость не требуется.
Public Property Get Applicable() As Boolean
    Applicable = mApplicable
End Property

' False при нехватке обязательных исходных данных.
Public Property Get InputValid() As Boolean
    InputValid = mInputValid
End Property

' Итог проверки по выбранной ветви.
Public Property Get CheckPassed() As Boolean
    CheckPassed = mCheckPassed
End Property

' Подробная внутренняя причина, если расчет не выполнен.
Public Property Get StopReason() As String
    StopReason = mStopReason
End Property

' Объясняет готовую инженерную проверку, не определяя статус по тексту.
' При отказе перечисляет только не прошедшие главные плоскости и их безразмерные
' запас/eta; физические условия остаются в CalculateSP63/CalculateSP35.
Private Function StabilityCheckComment() As String
    If mCheckPassed Then
        If mCode = STABILITY_CODE_SP35 Then
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

' Формирует статус и пояснение самой проверки устойчивости.
' При записи на лист они не определяются заново по названию метода Branch.
' Отсутствие программного объекта и ошибка Config имеют разные причины.
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

' Продольная сила после фильтра; знак сохраняется как во внутреннем ядре.
Public Property Get DesignN() As Double
    DesignN = mDesignN
End Property

' Mx после учета случайного эксцентриситета и eta.
Public Property Get DesignMx() As Double
    DesignMx = mDesignMx
End Property

' My после учета случайного эксцентриситета и eta.
Public Property Get DesignMy() As Double
    DesignMy = mDesignMy
End Property

' Возвращает угол главных осей приведённого сечения, использованный в устойчивости.
' Пакетный расчёт сохраняет его вместе с итоговыми данными
' после фильтра практически нулевых моментов.
Public Property Get PrincipalAngleRad() As Double
    PrincipalAngleRad = mPrincipalAngleRad
End Property

' Сжимающая часть постоянной/длительной силы, переданная для учета длительности.
Public Property Get SustainedN() As Double
    SustainedN = mSustainedN
End Property

' Исходный расчетный момент в первой главной плоскости до eta.
Public Property Get Moment1() As Double
    Moment1 = mMoment1
End Property

' Исходный расчетный момент во второй главной плоскости до eta.
Public Property Get Moment2() As Double
    Moment2 = mMoment2
End Property

' Постоянный/длительный момент в первой главной плоскости.
Public Property Get SustainedMoment1() As Double
    SustainedMoment1 = mSustainedMoment1
End Property

' Постоянный/длительный момент во второй главной плоскости.
Public Property Get SustainedMoment2() As Double
    SustainedMoment2 = mSustainedMoment2
End Property

' Итоговый момент первой главной плоскости после учета eta.
Public Property Get DesignMoment1() As Double
    DesignMoment1 = mDesignMoment1
End Property

' Итоговый момент второй главной плоскости после учета eta.
Public Property Get DesignMoment2() As Double
    DesignMoment2 = mDesignMoment2
End Property

' Минимальная контрольная сила: Ncr для eta-ветвей или Nult,stab для табличной ветви.
Public Property Get CriticalForce() As Double
    CriticalForce = mCriticalForce
End Property

' Минимальный запас по устойчивости.
Public Property Get ReserveFactor() As Double
    ReserveFactor = mReserveFactor
End Property

' Критическая сила Ncr первой главной плоскости; 0, если ветвь считает Nult вместо Ncr.
Public Property Get Ncr1() As Double
    Ncr1 = mNcr1
End Property

' Критическая сила Ncr второй главной плоскости; 0, если ветвь считает Nult вместо Ncr.
Public Property Get Ncr2() As Double
    Ncr2 = mNcr2
End Property

' Предельная сжимающая сила Nult,stab первой главной плоскости для табличной ветви СП 35.
Public Property Get Nultimate1() As Double
    Nultimate1 = mNultimate1
End Property

' Предельная сжимающая сила Nult,stab второй главной плоскости для табличной ветви СП 35.
Public Property Get Nultimate2() As Double
    Nultimate2 = mNultimate2
End Property

' Запас по первой главной плоскости.
Public Property Get Reserve1() As Double
    Reserve1 = mReserve1
End Property

' Запас по второй главной плоскости.
Public Property Get Reserve2() As Double
    Reserve2 = mReserve2
End Property

' Коэффициент eta первой плоскости.
Public Property Get Eta1() As Double
    Eta1 = mEta1
End Property

' Коэффициент eta второй плоскости.
Public Property Get Eta2() As Double
    Eta2 = mEta2
End Property

' Коэффициент длительности первой плоскости.
Public Property Get PhiL1() As Double
    PhiL1 = mPhiL1
End Property

' Коэффициент длительности второй плоскости.
Public Property Get PhiL2() As Double
    PhiL2 = mPhiL2
End Property

' Табличный phi_l СП 35 первой плоскости для ветви |ec|/r <= 1.
Public Property Get PhiLTable1() As Double
    PhiLTable1 = mPhiLTable1
End Property

' Табличный phi_l СП 35 второй плоскости для ветви |ec|/r <= 1.
Public Property Get PhiLTable2() As Double
    PhiLTable2 = mPhiLTable2
End Property

' Табличный phi_m СП 35 первой плоскости.
Public Property Get PhiM1() As Double
    PhiM1 = mPhiM1
End Property

' Табличный phi_m СП 35 второй плоскости.
Public Property Get PhiM2() As Double
    PhiM2 = mPhiM2
End Property

' Итоговый phi СП 35 первой плоскости в ветви |ec| <= r.
Public Property Get PhiValue1() As Double
    PhiValue1 = mPhiValue1
End Property

' Итоговый phi СП 35 второй плоскости в ветви |ec| <= r.
Public Property Get PhiValue2() As Double
    PhiValue2 = mPhiValue2
End Property

' Коэффициент phi_p СП 35 первой плоскости в ветви |ec| > r.
Public Property Get PhiP1() As Double
    PhiP1 = mPhiP1
End Property

' Коэффициент phi_p СП 35 второй плоскости в ветви |ec| > r.
Public Property Get PhiP2() As Double
    PhiP2 = mPhiP2
End Property

' Расчетный эксцентриситет первой плоскости, мм.
Public Property Get Eccentricity1() As Double
    Eccentricity1 = mEccentricity1
End Property

' Расчетный эксцентриситет второй плоскости, мм.
Public Property Get Eccentricity2() As Double
    Eccentricity2 = mEccentricity2
End Property

' Случайный эксцентриситет первой плоскости, мм.
Public Property Get AccidentalEccentricity1() As Double
    AccidentalEccentricity1 = mAccidentalEccentricity1
End Property

' Случайный эксцентриситет второй плоскости, мм.
Public Property Get AccidentalEccentricity2() As Double
    AccidentalEccentricity2 = mAccidentalEccentricity2
End Property

' l0 первой плоскости, мм.
Public Property Get EffectiveLength1() As Double
    EffectiveLength1 = mEffectiveLength1
End Property

' l0 второй плоскости, мм.
Public Property Get EffectiveLength2() As Double
    EffectiveLength2 = mEffectiveLength2
End Property

' Гибкость l0/i для плоскости 1: изгиб вокруг главной оси 1.
Public Property Get Slenderness1() As Double
    Slenderness1 = mSlenderness1
End Property

' Гибкость l0/i для плоскости 2: изгиб вокруг главной оси 2.
Public Property Get Slenderness2() As Double
    Slenderness2 = mSlenderness2
End Property

' Бетонный габарит h1 вдоль оси 2 при изгибе вокруг оси 1, мм.
Public Property Get Depth1() As Double
    Depth1 = mDepth1
End Property

' Бетонный габарит h2 вдоль оси 1 при изгибе вокруг оси 2, мм.
Public Property Get Depth2() As Double
    Depth2 = mDepth2
End Property

' Радиус инерции i1 приведенного сечения.
Public Property Get Radius1() As Double
    Radius1 = mRadius1
End Property

' Радиус инерции i2 приведенного сечения.
Public Property Get Radius2() As Double
    Radius2 = mRadius2
End Property

' Модуль упругости бетона для расчёта устойчивости, МПа.
Public Property Get ConcreteElasticModulus() As Double
    ConcreteElasticModulus = mConcreteElasticModulus
End Property

' Модуль упругости арматуры для расчёта устойчивости, МПа.
Public Property Get SteelElasticModulus() As Double
    SteelElasticModulus = mSteelElasticModulus
End Property

' Момент инерции бетона относительно главной оси 1 приведённого сечения, мм4.
Public Property Get ConcreteInertia1() As Double
    ConcreteInertia1 = mConcreteInertia1
End Property

' Момент инерции бетона относительно главной оси 2 приведённого сечения, мм4.
Public Property Get ConcreteInertia2() As Double
    ConcreteInertia2 = mConcreteInertia2
End Property

' Момент инерции арматуры относительно главной оси 1 приведённого сечения, мм4.
Public Property Get SteelInertia1() As Double
    SteelInertia1 = mSteelInertia1
End Property

' Момент инерции арматуры относительно главной оси 2 приведённого сечения, мм4.
Public Property Get SteelInertia2() As Double
    SteelInertia2 = mSteelInertia2
End Property

' Ядровое расстояние r первой плоскости со стороны эксцентриситета.
Public Property Get CoreDistance1() As Double
    CoreDistance1 = mCoreDistance1
End Property

' Ядровое расстояние r второй плоскости со стороны эксцентриситета.
Public Property Get CoreDistance2() As Double
    CoreDistance2 = mCoreDistance2
End Property

' delta или delta_e первой плоскости в выбранной расчетной ветви.
Public Property Get Delta1() As Double
    Delta1 = mDelta1
End Property

' delta или delta_e второй плоскости в выбранной расчетной ветви.
Public Property Get Delta2() As Double
    Delta2 = mDelta2
End Property

' Коэффициент kb СП 63 первой плоскости.
Public Property Get Kb1() As Double
    Kb1 = mKb1
End Property

' Коэффициент kb СП 63 второй плоскости.
Public Property Get Kb2() As Double
    Kb2 = mKb2
End Property

' Коэффициент ks СП 63 первой плоскости.
Public Property Get Ks1() As Double
    Ks1 = mKs1
End Property

' Коэффициент ks СП 63 второй плоскости.
Public Property Get Ks2() As Double
    Ks2 = mKs2
End Property

' Условная жесткость D первой плоскости, Н*мм2.
Public Property Get StiffnessD1() As Double
    StiffnessD1 = mStiffnessD1
End Property

' Условная жесткость D второй плоскости, Н*мм2.
Public Property Get StiffnessD2() As Double
    StiffnessD2 = mStiffnessD2
End Property

' Фактическая ветвь расчета первой плоскости.
Public Property Get PlaneBranch1() As String
    PlaneBranch1 = mPlaneBranch1
End Property

' Фактическая ветвь расчета второй плоскости.
Public Property Get PlaneBranch2() As String
    PlaneBranch2 = mPlaneBranch2
End Property

' True, если первая плоскость участвовала в проверке.
Public Property Get PlaneApplicable1() As Boolean
    PlaneApplicable1 = mPlaneApplicable1
End Property

' True, если вторая плоскость участвовала в проверке.
Public Property Get PlaneApplicable2() As Boolean
    PlaneApplicable2 = mPlaneApplicable2
End Property

' True, если первая плоскость прошла свою проверку.
Public Property Get PlanePassed1() As Boolean
    PlanePassed1 = mPlanePassed1
End Property

' True, если вторая плоскость прошла свою проверку.
Public Property Get PlanePassed2() As Boolean
    PlanePassed2 = mPlanePassed2
End Property
