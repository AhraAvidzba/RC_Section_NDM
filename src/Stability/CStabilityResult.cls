VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CStabilityResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Результат проверки устойчивости одного сочетания нагрузок
' ==========================================================================
' Хранит статус, пояснения и инженерные величины, полученные
' CStabilityCalculator. Входит в CCombinationResult и служит источником
' данных для таблицы Results без повторного расчёта устойчивости.

Private mFrozen As Boolean ' После передачи готового результата его заполнение и очистка запрещены; сохранённые числа доступны для чтения.
Private mMeta As CResultMeta ' Статус, код причины и пояснение проверки устойчивости.
Private mCode As String ' Нормативная ветка SP35/SP63.
Private mBranch As String ' Диагностическое имя расчетной ветви.
Private mReferenceX As Double ' X точки, относительно которой выполнен перенос нагрузок.
Private mReferenceY As Double ' Y точки, относительно которой выполнен перенос нагрузок.
Private mDesignN As Double ' Итоговая продольная сила для последующего НДМ.
Private mDesignMx As Double ' Итоговый Mx после учета устойчивости/фильтров.
Private mDesignMy As Double ' Итоговый My после учета устойчивости/фильтров.
Private mSustainedN As Double ' Длительная/постоянная часть N.
Private mMoment1 As Double ' Исходный расчетный момент в первой главной плоскости до eta.
Private mMoment2 As Double ' Исходный расчетный момент во второй главной плоскости до eta.
Private mSustainedMoment1 As Double ' Постоянный/длительный момент в первой главной плоскости.
Private mSustainedMoment2 As Double ' Постоянный/длительный момент во второй главной плоскости.
Private mDesignMoment1 As Double ' Итоговый момент первой главной плоскости после учета eta.
Private mDesignMoment2 As Double ' Итоговый момент второй главной плоскости после учета eta.
Private mCriticalForce As Double ' Минимальная контрольная сила: Ncr для eta-ветвей или Nult,stab для табличной ветви.
Private mReserveFactor As Double ' Минимальный запас по устойчивости.
Private mUtilization As Double ' Наибольший коэффициент использования при проверке устойчивости.
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
Private mPhiLTable1 As String ' Табличный phi_l СП 35 первой плоскости для ветви ec/r <= 1.
Private mPhiLTable2 As String ' Табличный phi_l СП 35 второй плоскости для ветви ec/r <= 1.
Private mPhiM1 As Double ' Табличный phi_m СП 35 первой плоскости.
Private mPhiM2 As Double ' Табличный phi_m СП 35 второй плоскости.
Private mPhiValue1 As Double ' Итоговый phi СП 35 первой плоскости в ветви ec <= r.
Private mPhiValue2 As Double ' Итоговый phi СП 35 второй плоскости в ветви ec <= r.
Private mPhiP1 As Double ' Коэффициент phi_p СП 35 первой плоскости в ветви ec > r.
Private mPhiP2 As Double ' Коэффициент phi_p СП 35 второй плоскости в ветви ec > r.
Private mEccentricity1 As Double ' Расчетный эксцентриситет первой плоскости, мм.
Private mEccentricity2 As Double ' Расчетный эксцентриситет второй плоскости, мм.
Private mAccidentalEcc1 As Double ' Случайный эксцентриситет в первой плоскости, мм.
Private mAccidentalEcc2 As Double ' Случайный эксцентриситет во второй плоскости, мм.
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

' Создаёт результат N/A для проверки устойчивости, не запрошенной профилем.
Private Sub Class_Initialize()
    Clear
End Sub

' Очищает численные результаты, статус и пояснение перед новым заполнением.
Public Sub Clear()
    AssertWritable
    Set mMeta = New CResultMeta
    mMeta.SetNotApplicable rkStability
    mCode = vbNullString
    mBranch = vbNullString
    mReferenceX = 0#
    mReferenceY = 0#
    mDesignN = 0#
    mDesignMx = 0#
    mDesignMy = 0#
    mSustainedN = 0#
    mMoment1 = 0#
    mMoment2 = 0#
    mSustainedMoment1 = 0#
    mSustainedMoment2 = 0#
    mDesignMoment1 = 0#
    mDesignMoment2 = 0#
    mCriticalForce = 0#
    mReserveFactor = 0#
    mUtilization = 0#
    mNcr1 = 0#
    mNcr2 = 0#
    mNultimate1 = 0#
    mNultimate2 = 0#
    mReserve1 = 0#
    mReserve2 = 0#
    mEta1 = 0#
    mEta2 = 0#
    mPhiL1 = 0#
    mPhiL2 = 0#
    mPhiLTable1 = vbNullString
    mPhiLTable2 = vbNullString
    mPhiM1 = 0#
    mPhiM2 = 0#
    mPhiValue1 = 0#
    mPhiValue2 = 0#
    mPhiP1 = 0#
    mPhiP2 = 0#
    mEccentricity1 = 0#
    mEccentricity2 = 0#
    mAccidentalEcc1 = 0#
    mAccidentalEcc2 = 0#
    mEffectiveLength1 = 0#
    mEffectiveLength2 = 0#
    mSlenderness1 = 0#
    mSlenderness2 = 0#
    mDepth1 = 0#
    mDepth2 = 0#
    mRadius1 = 0#
    mRadius2 = 0#
    mCoreDistance1 = 0#
    mCoreDistance2 = 0#
    mDelta1 = 0#
    mDelta2 = 0#
    mKb1 = 0#
    mKb2 = 0#
    mKs1 = 0#
    mKs2 = 0#
    mStiffnessD1 = 0#
    mStiffnessD2 = 0#
    mPlaneBranch1 = vbNullString
    mPlaneBranch2 = vbNullString
    mPlaneApplicable1 = False
    mPlaneApplicable2 = False
    mPlanePassed1 = False
    mPlanePassed2 = False
End Sub

' Сохраняет статус и пояснение устойчивости без отдельной независимо изменяемой строки статуса.
Public Sub SetMeta(ByVal meta As CResultMeta, Optional ByVal resultComment As String = vbNullString)
    AssertWritable
    If meta Is Nothing Then
        Set mMeta = New CResultMeta
        mMeta.SetResult rsInternalError, rcInternalError, rkStability, _
            "Результат устойчивости не сформирован.", vbNullString, True, False
    Else
        Set mMeta = meta.Clone
        If Len(Trim$(resultComment)) > 0 Then mMeta.ResultComment = StabilityBranchDisplayName(resultComment)
    End If
    mBranch = StabilityBranchDisplayName(resultComment)
End Sub

' Копирует готовые величины CStabilityCalculator и формирует статус для общей сводки.
Public Sub InitializeFromCalculator(ByVal meta As CResultMeta, ByVal stability As CStabilityCalculator, _
        ByVal referenceXValue As Double, ByVal referenceYValue As Double)
    AssertWritable
    Clear
    mReferenceX = referenceXValue
    mReferenceY = referenceYValue

    If meta Is Nothing Then
        Set mMeta = New CResultMeta
        mMeta.SetResult rsInternalError, rcInternalError, rkStability, _
            "Результат устойчивости не сформирован.", vbNullString, True, False
    Else
        Set mMeta = meta.Clone
    End If

    If stability Is Nothing Then
        Exit Sub
    End If

    mCode = stability.Code
    mBranch = stability.Branch ' Сохраняем Branch как код применённого метода. Причина отказа записывается отдельно в ResultComment и не подменяет этот код.
    mDesignN = stability.DesignN
    mDesignMx = stability.DesignMx
    mDesignMy = stability.DesignMy
    mSustainedN = stability.SustainedN
    mMoment1 = stability.Moment1
    mMoment2 = stability.Moment2
    mSustainedMoment1 = stability.SustainedMoment1
    mSustainedMoment2 = stability.SustainedMoment2
    mDesignMoment1 = stability.DesignMoment1
    mDesignMoment2 = stability.DesignMoment2
    mCriticalForce = stability.CriticalForce
    mReserveFactor = stability.ReserveFactor
    If mReserveFactor > 0# Then mUtilization = 1# / mReserveFactor
    mNcr1 = stability.Ncr1
    mNcr2 = stability.Ncr2
    mNultimate1 = stability.Nultimate1
    mNultimate2 = stability.Nultimate2
    mReserve1 = stability.Reserve1
    mReserve2 = stability.Reserve2
    mEta1 = stability.Eta1
    mEta2 = stability.Eta2
    mPhiL1 = stability.PhiL1
    mPhiL2 = stability.PhiL2
    mPhiLTable1 = stability.PhiLTable1
    mPhiLTable2 = stability.PhiLTable2
    mPhiM1 = stability.PhiM1
    mPhiM2 = stability.PhiM2
    mPhiValue1 = stability.PhiValue1
    mPhiValue2 = stability.PhiValue2
    mPhiP1 = stability.PhiP1
    mPhiP2 = stability.PhiP2
    mEccentricity1 = stability.Eccentricity1
    mEccentricity2 = stability.Eccentricity2
    mAccidentalEcc1 = stability.AccidentalEccentricity1
    mAccidentalEcc2 = stability.AccidentalEccentricity2
    mEffectiveLength1 = stability.EffectiveLength1
    mEffectiveLength2 = stability.EffectiveLength2
    mSlenderness1 = stability.Slenderness1
    mSlenderness2 = stability.Slenderness2
    mDepth1 = stability.Depth1
    mDepth2 = stability.Depth2
    mRadius1 = stability.Radius1
    mRadius2 = stability.Radius2
    mCoreDistance1 = stability.CoreDistance1
    mCoreDistance2 = stability.CoreDistance2
    mDelta1 = stability.Delta1
    mDelta2 = stability.Delta2
    mKb1 = stability.Kb1
    mKb2 = stability.Kb2
    mKs1 = stability.Ks1
    mKs2 = stability.Ks2
    mStiffnessD1 = stability.StiffnessD1
    mStiffnessD2 = stability.StiffnessD2
    mPlaneBranch1 = stability.PlaneBranch1
    mPlaneBranch2 = stability.PlaneBranch2
    mPlaneApplicable1 = stability.PlaneApplicable1
    mPlaneApplicable2 = stability.PlaneApplicable2
    mPlanePassed1 = stability.PlanePassed1
    mPlanePassed2 = stability.PlanePassed2
End Sub

' Возвращает русское название применённого метода для Results.
' Расчётные коды SP35-eta/SP35-table/SP35-mixed сохраняются отдельно.
Private Function StabilityBranchDisplayName(ByVal branchText As String) As String
    Select Case LCase$(Trim$(branchText))
        Case "nocompression"
            StabilityBranchDisplayName = "устойчивость не требуется: нет сжатия"
        Case "sp63"
            StabilityBranchDisplayName = "СП 63: расчет eta и Ncr"
        Case "sp35-eta"
            StabilityBranchDisplayName = "СП 35: ветвь ec > r через eta и Ncr"
        Case "sp35-table"
            StabilityBranchDisplayName = "СП 35: ветвь ec <= r через Nult,stab"
        Case "sp35-mixed"
            StabilityBranchDisplayName = "СП 35: плоскости попали в разные расчетные ветви"
        Case Else
            StabilityBranchDisplayName = branchText
    End Select
End Function

' Сохраняет расчётные N, Mx, My после обнуления малых моментов.
' Вычисляет моменты в главных осях для вывода в Results.
Public Sub UpdateDesignLoads(ByVal designNValue As Double, ByVal designMxValue As Double, _
        ByVal designMyValue As Double, ByVal principalAngleRad As Double)
    AssertWritable
    mDesignN = designNValue
    mDesignMx = designMxValue
    mDesignMy = designMyValue
    mDesignMoment1 = designMxValue * Cos(principalAngleRad) + designMyValue * Sin(principalAngleRad)
    mDesignMoment2 = -designMxValue * Sin(principalAngleRad) + designMyValue * Cos(principalAngleRad)
End Sub

' Возвращает копию статуса и пояснения, чтобы вызывающий код не мог изменить готовую проверку.
Public Property Get Meta() As CResultMeta
    If mMeta Is Nothing Then
        Set mMeta = New CResultMeta
        mMeta.SetResult rsInternalError, rcInternalError, rkStability, _
            "Результат устойчивости не сформирован.", vbNullString, True, False
    End If
    Set Meta = mMeta.Clone
End Property

' Возвращает видимое обозначение по сохранённому статусу; отдельного изменяемого значения нет.
Public Property Get Status() As String
    Dim policy As CResultStatusPolicy
    Set policy = New CResultStatusPolicy
    Status = policy.ExternalStatus(mMeta)
End Property

' Выдает минимальный доступный запас главных плоскостей для сводки и Worst.
' Числа уже учитывают примененный критерий устойчивости; здесь не пересчитываются СП.
Public Property Get SummaryReserve() As Double
    Dim policy As CResultStatusPolicy
    Set policy = New CResultStatusPolicy
    If Not policy.IsMetaFinished(mMeta) Then Exit Property
    If mPlaneApplicable1 And mReserve1 > 0# Then SummaryReserve = mReserve1
    If mPlaneApplicable2 And mReserve2 > 0# Then
        If SummaryReserve <= 0# Or mReserve2 < SummaryReserve Then SummaryReserve = mReserve2
    End If
    If SummaryReserve <= 0# And mReserveFactor > 0# Then SummaryReserve = mReserveFactor
End Property

' Ниже возвращаются готовые величины проверки устойчивости.
' Они доступны только для чтения и не являются новыми входными данными расчёта.
Public Property Get Code() As String
    Code = mCode
End Property

' Диагностическое имя расчетной ветви.
Public Property Get Branch() As String
    Branch = mBranch
End Property

' X точки, относительно которой выполнен перенос нагрузок.
Public Property Get ReferenceX() As Double
    ReferenceX = mReferenceX
End Property

' Y точки, относительно которой выполнен перенос нагрузок.
Public Property Get ReferenceY() As Double
    ReferenceY = mReferenceY
End Property

' Итоговая продольная сила для последующего НДМ.
Public Property Get DesignN() As Double
    DesignN = mDesignN
End Property

' Итоговый Mx после учета устойчивости/фильтров.
Public Property Get DesignMx() As Double
    DesignMx = mDesignMx
End Property

' Итоговый My после учета устойчивости/фильтров.
Public Property Get DesignMy() As Double
    DesignMy = mDesignMy
End Property

' Длительная/постоянная часть N.
Public Property Get SustainedN() As Double
    SustainedN = mSustainedN
End Property

' Сохранённый параметр устойчивости в плоскости 1: исходный момент. Значение определяет
' CStabilityCalculator по выбранным нормам и Config.
Public Property Get Moment1() As Double
    Moment1 = mMoment1
End Property

' Сохранённый параметр устойчивости в плоскости 2: исходный момент. Значение определяет
' CStabilityCalculator по выбранным нормам и Config.
Public Property Get Moment2() As Double
    Moment2 = mMoment2
End Property

' Сохранённый параметр устойчивости в плоскости 1: длительная часть момента. Значение определяет
' CStabilityCalculator по выбранным нормам и Config.
Public Property Get SustainedMoment1() As Double
    SustainedMoment1 = mSustainedMoment1
End Property

' Сохранённый параметр устойчивости в плоскости 2: длительная часть момента. Значение определяет
' CStabilityCalculator по выбранным нормам и Config.
Public Property Get SustainedMoment2() As Double
    SustainedMoment2 = mSustainedMoment2
End Property

' Сохранённый параметр устойчивости в плоскости 1: расчётный момент после учёта устойчивости.
' Значение определяет CStabilityCalculator по выбранным нормам и Config.
Public Property Get DesignMoment1() As Double
    DesignMoment1 = mDesignMoment1
End Property

' Сохранённый параметр устойчивости в плоскости 2: расчётный момент после учёта устойчивости.
' Значение определяет CStabilityCalculator по выбранным нормам и Config.
Public Property Get DesignMoment2() As Double
    DesignMoment2 = mDesignMoment2
End Property

' Критическая продольная сила по определяющей плоскости устойчивости, Н.
Public Property Get CriticalForce() As Double
    CriticalForce = mCriticalForce
End Property

' Определяющий коэффициент запаса устойчивости.
Public Property Get ReserveFactor() As Double
    ReserveFactor = mReserveFactor
End Property

' Сохранённый коэффициент использования проверяемого предела; смысл предела задаёт результат
' этого класса.
Public Property Get Utilization() As Double
    Utilization = mUtilization
End Property

' Сохранённый параметр устойчивости в плоскости 1: критическая сила. Значение определяет
' CStabilityCalculator по выбранным нормам и Config.
Public Property Get Ncr1() As Double
    Ncr1 = mNcr1
End Property

' Сохранённый параметр устойчивости в плоскости 2: критическая сила. Значение определяет
' CStabilityCalculator по выбранным нормам и Config.
Public Property Get Ncr2() As Double
    Ncr2 = mNcr2
End Property

' Сохранённый параметр устойчивости в плоскости 1: предельная сила. Значение определяет
' CStabilityCalculator по выбранным нормам и Config.
Public Property Get Nultimate1() As Double
    Nultimate1 = mNultimate1
End Property

' Сохранённый параметр устойчивости в плоскости 2: предельная сила. Значение определяет
' CStabilityCalculator по выбранным нормам и Config.
Public Property Get Nultimate2() As Double
    Nultimate2 = mNultimate2
End Property

' Сохранённый параметр устойчивости в плоскости 1: коэффициент запаса. Значение определяет
' CStabilityCalculator по выбранным нормам и Config.
Public Property Get Reserve1() As Double
    Reserve1 = mReserve1
End Property

' Сохранённый параметр устойчивости в плоскости 2: коэффициент запаса. Значение определяет
' CStabilityCalculator по выбранным нормам и Config.
Public Property Get Reserve2() As Double
    Reserve2 = mReserve2
End Property

' Сохранённый параметр устойчивости в плоскости 1: коэффициент увеличения эксцентриситета.
' Значение определяет CStabilityCalculator по выбранным нормам и Config.
Public Property Get Eta1() As Double
    Eta1 = mEta1
End Property

' Сохранённый параметр устойчивости в плоскости 2: коэффициент увеличения эксцентриситета.
' Значение определяет CStabilityCalculator по выбранным нормам и Config.
Public Property Get Eta2() As Double
    Eta2 = mEta2
End Property

' Сохранённый параметр устойчивости в плоскости 1: коэффициент длительности. Значение определяет
' CStabilityCalculator по выбранным нормам и Config.
Public Property Get PhiL1() As Double
    PhiL1 = mPhiL1
End Property

' Сохранённый параметр устойчивости в плоскости 2: коэффициент длительности. Значение определяет
' CStabilityCalculator по выбранным нормам и Config.
Public Property Get PhiL2() As Double
    PhiL2 = mPhiL2
End Property

' Сохранённый параметр устойчивости в плоскости 1: табличная ветвь коэффициента длительности.
' Значение определяет CStabilityCalculator по выбранным нормам и Config.
Public Property Get PhiLTable1() As String
    PhiLTable1 = mPhiLTable1
End Property

' Сохранённый параметр устойчивости в плоскости 2: табличная ветвь коэффициента длительности.
' Значение определяет CStabilityCalculator по выбранным нормам и Config.
Public Property Get PhiLTable2() As String
    PhiLTable2 = mPhiLTable2
End Property

' Сохранённый параметр устойчивости в плоскости 1: коэффициент доли длительного момента.
' Значение определяет CStabilityCalculator по выбранным нормам и Config.
Public Property Get PhiM1() As Double
    PhiM1 = mPhiM1
End Property

' Сохранённый параметр устойчивости в плоскости 2: коэффициент доли длительного момента.
' Значение определяет CStabilityCalculator по выбранным нормам и Config.
Public Property Get PhiM2() As Double
    PhiM2 = mPhiM2
End Property

' Сохранённый параметр устойчивости в плоскости 1: итоговый коэффициент для жёсткости. Значение
' определяет CStabilityCalculator по выбранным нормам и Config.
Public Property Get PhiValue1() As Double
    PhiValue1 = mPhiValue1
End Property

' Сохранённый параметр устойчивости в плоскости 2: итоговый коэффициент для жёсткости. Значение
' определяет CStabilityCalculator по выбранным нормам и Config.
Public Property Get PhiValue2() As Double
    PhiValue2 = mPhiValue2
End Property

' Сохранённый параметр устойчивости в плоскости 1: коэффициент влияния предварительного
' напряжения в выбранной ветви. Значение определяет CStabilityCalculator по выбранным нормам и
' Config.
Public Property Get PhiP1() As Double
    PhiP1 = mPhiP1
End Property

' Сохранённый параметр устойчивости в плоскости 2: коэффициент влияния предварительного
' напряжения в выбранной ветви. Значение определяет CStabilityCalculator по выбранным нормам и
' Config.
Public Property Get PhiP2() As Double
    PhiP2 = mPhiP2
End Property

' Сохранённый параметр устойчивости в плоскости 1: эксцентриситет силы. Значение определяет
' CStabilityCalculator по выбранным нормам и Config.
Public Property Get Eccentricity1() As Double
    Eccentricity1 = mEccentricity1
End Property

' Сохранённый параметр устойчивости в плоскости 2: эксцентриситет силы. Значение определяет
' CStabilityCalculator по выбранным нормам и Config.
Public Property Get Eccentricity2() As Double
    Eccentricity2 = mEccentricity2
End Property

' Сохранённый параметр устойчивости в плоскости 1: случайный эксцентриситет. Значение определяет
' CStabilityCalculator по выбранным нормам и Config.
Public Property Get AccidentalEcc1() As Double
    AccidentalEcc1 = mAccidentalEcc1
End Property

' Сохранённый параметр устойчивости в плоскости 2: случайный эксцентриситет. Значение определяет
' CStabilityCalculator по выбранным нормам и Config.
Public Property Get AccidentalEcc2() As Double
    AccidentalEcc2 = mAccidentalEcc2
End Property

' Сохранённый параметр устойчивости в плоскости 1: расчётная длина. Значение определяет
' CStabilityCalculator по выбранным нормам и Config.
Public Property Get EffectiveLength1() As Double
    EffectiveLength1 = mEffectiveLength1
End Property

' Сохранённый параметр устойчивости в плоскости 2: расчётная длина. Значение определяет
' CStabilityCalculator по выбранным нормам и Config.
Public Property Get EffectiveLength2() As Double
    EffectiveLength2 = mEffectiveLength2
End Property

' Сохранённый параметр устойчивости в плоскости 1: гибкость. Значение определяет
' CStabilityCalculator по выбранным нормам и Config.
Public Property Get Slenderness1() As Double
    Slenderness1 = mSlenderness1
End Property

' Сохранённый параметр устойчивости в плоскости 2: гибкость. Значение определяет
' CStabilityCalculator по выбранным нормам и Config.
Public Property Get Slenderness2() As Double
    Slenderness2 = mSlenderness2
End Property

' Сохранённый параметр устойчивости в плоскости 1: размер сечения в плоскости. Значение
' определяет CStabilityCalculator по выбранным нормам и Config.
Public Property Get Depth1() As Double
    Depth1 = mDepth1
End Property

' Сохранённый параметр устойчивости в плоскости 2: размер сечения в плоскости. Значение
' определяет CStabilityCalculator по выбранным нормам и Config.
Public Property Get Depth2() As Double
    Depth2 = mDepth2
End Property

' Сохранённый параметр устойчивости в плоскости 1: радиус инерции. Значение определяет
' CStabilityCalculator по выбранным нормам и Config.
Public Property Get Radius1() As Double
    Radius1 = mRadius1
End Property

' Сохранённый параметр устойчивости в плоскости 2: радиус инерции. Значение определяет
' CStabilityCalculator по выбранным нормам и Config.
Public Property Get Radius2() As Double
    Radius2 = mRadius2
End Property

' Сохранённый параметр устойчивости в плоскости 1: расстояние ядра сечения. Значение определяет
' CStabilityCalculator по выбранным нормам и Config.
Public Property Get CoreDistance1() As Double
    CoreDistance1 = mCoreDistance1
End Property

' Сохранённый параметр устойчивости в плоскости 2: расстояние ядра сечения. Значение определяет
' CStabilityCalculator по выбранным нормам и Config.
Public Property Get CoreDistance2() As Double
    CoreDistance2 = mCoreDistance2
End Property

' Сохранённый параметр устойчивости в плоскости 1: относительный эксцентриситет. Значение
' определяет CStabilityCalculator по выбранным нормам и Config.
Public Property Get Delta1() As Double
    Delta1 = mDelta1
End Property

' Сохранённый параметр устойчивости в плоскости 2: относительный эксцентриситет. Значение
' определяет CStabilityCalculator по выбранным нормам и Config.
Public Property Get Delta2() As Double
    Delta2 = mDelta2
End Property

' Сохранённый параметр устойчивости в плоскости 1: коэффициент бетонной части жёсткости.
' Значение определяет CStabilityCalculator по выбранным нормам и Config.
Public Property Get Kb1() As Double
    Kb1 = mKb1
End Property

' Сохранённый параметр устойчивости в плоскости 2: коэффициент бетонной части жёсткости.
' Значение определяет CStabilityCalculator по выбранным нормам и Config.
Public Property Get Kb2() As Double
    Kb2 = mKb2
End Property

' Сохранённый параметр устойчивости в плоскости 1: коэффициент арматурной части жёсткости.
' Значение определяет CStabilityCalculator по выбранным нормам и Config.
Public Property Get Ks1() As Double
    Ks1 = mKs1
End Property

' Сохранённый параметр устойчивости в плоскости 2: коэффициент арматурной части жёсткости.
' Значение определяет CStabilityCalculator по выбранным нормам и Config.
Public Property Get Ks2() As Double
    Ks2 = mKs2
End Property

' Сохранённый параметр устойчивости в плоскости 1: расчётная жёсткость. Значение определяет
' CStabilityCalculator по выбранным нормам и Config.
Public Property Get StiffnessD1() As Double
    StiffnessD1 = mStiffnessD1
End Property

' Сохранённый параметр устойчивости в плоскости 2: расчётная жёсткость. Значение определяет
' CStabilityCalculator по выбранным нормам и Config.
Public Property Get StiffnessD2() As Double
    StiffnessD2 = mStiffnessD2
End Property

' Сохранённый параметр устойчивости в плоскости 1: принятая расчётная ветвь. Значение определяет
' CStabilityCalculator по выбранным нормам и Config.
Public Property Get PlaneBranch1() As String
    PlaneBranch1 = mPlaneBranch1
End Property

' Сохранённый параметр устойчивости в плоскости 2: принятая расчётная ветвь. Значение определяет
' CStabilityCalculator по выбранным нормам и Config.
Public Property Get PlaneBranch2() As String
    PlaneBranch2 = mPlaneBranch2
End Property

' Сохранённый параметр устойчивости в плоскости 1: признак применимости проверки. Значение
' определяет CStabilityCalculator по выбранным нормам и Config.
Public Property Get PlaneApplicable1() As Boolean
    PlaneApplicable1 = mPlaneApplicable1
End Property

' Сохранённый параметр устойчивости в плоскости 2: признак применимости проверки. Значение
' определяет CStabilityCalculator по выбранным нормам и Config.
Public Property Get PlaneApplicable2() As Boolean
    PlaneApplicable2 = mPlaneApplicable2
End Property

' Сохранённый параметр устойчивости в плоскости 1: признак выполнения проверки. Значение
' определяет CStabilityCalculator по выбранным нормам и Config.
Public Property Get PlanePassed1() As Boolean
    PlanePassed1 = mPlanePassed1
End Property

' Сохранённый параметр устойчивости в плоскости 2: признак выполнения проверки. Значение
' определяет CStabilityCalculator по выбранным нормам и Config.
Public Property Get PlanePassed2() As Boolean
    PlanePassed2 = mPlanePassed2
End Property


' Завершает заполнение устойчивости и запрещает менять сохранённые числа,
' статус и пояснение. Данные не копируются и не рассчитываются заново;
' повторный вызов безопасен.
Friend Sub Freeze()
    If mFrozen Then Exit Sub
    mFrozen = True
End Sub

' Запрещает изменение результата после передачи потребителям.
' Для повторного расчёта создаётся новый объект, а прежний результат сохраняется.
Private Sub AssertWritable()
    If mFrozen Then Err.Raise vbObjectError + 4220, "CStabilityResult", _
        "Опубликованный результат нельзя изменить; для нового расчета нужен новый объект."
End Sub
