VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCombinationResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Результат одного сочетания нагрузок
' ==========================================================================
' CCombinationResult хранит все данные, которые относятся к одному LC после
' прохождения batch-pipeline: прямое НДС, несущая способность, трещины,
' пользовательские статусы и машинные признаки вроде ExtensionUsed.
'
' Класс не запускает solver-ы и не принимает инженерных решений. Его задача -
' держать результат одной строки в одном объекте, чтобы CBatchSectionCalculator
' не превращался в набор параллельных массивов и длинных блоков присваивания.

Public Status As String ' Итоговый пользовательский статус LC.
Public ProfileId As String ' Расчетный профиль, назначенный этому LC в rngLoadCombinations.
Public DirectStateStatus As String ' Статус прямого решения НДС фактического LC.
Public CapacityStatus As String ' Статус поиска предельной несущей способности.
Public CapacityLimitState As String ' Причина предельного состояния или текст ошибки capacity.
Public CapacitySolutionMethod As String ' Фактический метод, которым получена capacity-точка.
Public CrackStatus As String ' Статус расчета раскрытия нормальных трещин.
Public LongitudinalCrackStatus As String ' Статус проверки продольных трещин по сжатому бетону.
Public StabilityStatus As String ' Статус учета продольного изгиба/устойчивости перед НДМ.
Public ExtensionUsed As Boolean ' True, если прямое НДС использовало technical extension диаграммы.
Public DirectStateDiagramExtensionEnabled As Boolean ' Снимок настройки extension на момент расчета LC.
Public LambdaCapacity As Double ' Предельный множитель выбранной lambda-траектории.
Public UtilCapacity As Double ' Использование несущей способности, обычно 1/lambda.
Public CapacityPathResolved As String ' Нормализованный ключ выбранной lambda-траектории.
Public NUltimate As Double ' Полная предельная N CapacityState для вывода в пользовательской точке нагрузки.
Public MxUltimate As Double ' Полный предельный Mx CapacityState, перенесенный в точку приложения нагрузки.
Public MyUltimate As Double ' Полный предельный My CapacityState, перенесенный в точку приложения нагрузки.
Public CapacityStateN As Double ' Полная N в найденном предельном CapacityState независимо от выбранной траектории.
Public CapacityStateMx As Double ' Полный Mx в найденном предельном CapacityState независимо от выбранной траектории.
Public CapacityStateMy As Double ' Полный My в найденном предельном CapacityState независимо от выбранной траектории.
Public MomentUltimate As Double ' Полный предельный момент sqrt(Mx^2+My^2).
Public CapacityStateAvailable As Boolean ' True, если есть плоскость деформаций capacity-состояния.
Public CapacityEpsilon0 As Double ' epsilon0 в найденной capacity-точке.
Public CapacityKappaX As Double ' kappaX в найденной capacity-точке.
Public CapacityKappaY As Double ' kappaY в найденной capacity-точке.
Public MinConcreteStrain As Double ' Минимальная деформация бетонных волокон в прямом НДС.
Public MaxConcreteStrain As Double ' Максимальная деформация бетонных волокон в прямом НДС.
Public MaxConcreteCompressionStress As Double ' Максимальное по модулю сжимающее напряжение бетона в прямом НДС.
Public MinSteelStrain As Double ' Минимальная деформация арматуры в прямом НДС.
Public MaxSteelStrain As Double ' Максимальная деформация арматуры в прямом НДС.
Public CrackWidth As Double ' Расчетная ширина раскрытия нормальных трещин.
Public CrackAllowableValue As Double ' Допустимая ширина раскрытия, заданная в Config.
Public CrackUtil As Double ' Использование по раскрытию трещин a_crc/a_crc,ult.
Public CrackFormed As Boolean ' True, если по принятой методике трещина образована.
Public CrackPhi1 As Double ' Коэффициент длительности действия нагрузки.
Public CrackPhi2 As Double ' Пользовательский коэффициент phi2 из настроек трещин.
Public CrackPhi3 As Double ' Коэффициент характера продольной силы.
Public CrackPsiS As Double ' Коэффициент неравномерности деформаций арматуры между трещинами.
Public CrackSigmaS As Double ' Средневзвешенное напряжение выбранной растянутой арматуры.
Public CrackSigmaSCrc As Double ' Напряжение той же группы стержней при образовании трещины.
Public CrackAs As Double ' Площадь выбранной растянутой арматуры.
Public CrackAbt As Double ' Площадь бетона принятой зоны взаимодействия.
Public CrackDsEquivalent As Double ' Эквивалентный диаметр выбранной растянутой арматуры.
Public CrackLsRaw As Double ' Базовое расстояние между трещинами до нормативных ограничений.
Public CrackLs As Double ' Принятое расстояние между трещинами после ограничений.
Public CrackLambdaCrc As Double ' Множитель нагрузки при образовании трещины.
Public CrackFormationMethod As String ' Фактически использованный путь поиска стадии образования трещины.
Public CrackNcrc As Double ' Продольная сила образования трещины для центрального растяжения.
Public CrackFormationNcrc As Double ' Продольная сила в найденной точке образования трещины выбранной lambda-траектории.
Public CrackMcrc As Double ' Модуль изгибающего момента образования трещины для общей ветки N+M.
Public CrackAred As Double ' Площадь приведенного сечения для Ncrc при центральном растяжении.
Public CrackRbtSer As Double ' Расчетное сопротивление бетона растяжению для II ГПС, использованное в Ncrc.
Public CrackTensionRebarCount As Long ' Число стержней, попавших в расчет трещин.
Public CrackTensionRebarIds As String ' Имена стержней, по которым определялись As и sigma_s.
Public CrackSectionDepthH As Double ' Полная высота сечения по нормали к нейтральной линии.
Public CrackCoverA As Double ' Расстояние от растянутой поверхности до крайней растянутой арматуры.
Public CrackTensionDepth As Double ' Фактическая глубина растянутой зоны.
Public CrackEffectiveZoneDepth As Double ' Принятая эффективная глубина зоны взаимодействия.
Public CrackCentralBranch As Boolean ' True для отдельной ветки центрального растяжения.
Public LongitudinalCrackRbMc2 As Double ' Допустимое Rb,mc2, с которым сравнивается сжатие бетона.
Public LongitudinalCrackUtil As Double ' Использование по продольным трещинам: sigma_c,max / Rb,mc2.
Public StabilityCode As String ' Выбранная нормативная ветвь: SP63 или SP35.
Public StabilityBranch As String ' Внутренняя трассировка расчетной ветви устойчивости для diagnostics/report.
Public StabilityReferenceX As Double ' X центра приведенного сечения, относительно которого считалась устойчивость.
Public StabilityReferenceY As Double ' Y центра приведенного сечения, относительно которого считалась устойчивость.
Public StabilityDesignN As Double ' Продольная сила после фильтра устойчивости; в текущей схеме равна исходной N.
Public StabilityDesignMx As Double ' Mx после учета продольного изгиба относительно центра приведенного сечения.
Public StabilityDesignMy As Double ' My после учета продольного изгиба относительно центра приведенного сечения.
Public StabilitySustainedN As Double ' Сжимающая часть постоянной/длительной силы, использованная для коэффициентов длительности.
Public StabilityMoment1 As Double ' Исходный момент первой главной плоскости до умножения на eta.
Public StabilityMoment2 As Double ' Исходный момент второй главной плоскости до умножения на eta.
Public StabilitySustainedMoment1 As Double ' Постоянный/длительный момент первой главной плоскости.
Public StabilitySustainedMoment2 As Double ' Постоянный/длительный момент второй главной плоскости.
Public StabilityDesignMoment1 As Double ' Итоговый момент первой главной плоскости после учета eta.
Public StabilityDesignMoment2 As Double ' Итоговый момент второй главной плоскости после учета eta.
Public StabilityCriticalForce As Double ' Минимальная контрольная сила устойчивости: Ncr или Nult,stab, Н.
Public StabilityReserveFactor As Double ' Запас устойчивости: Ncr/N или Nult,stab/N по выбранной ветви.
Public StabilityUtil As Double ' Использование устойчивости: 1 / StabilityReserveFactor.
Public StabilityNcr1 As Double ' Критическая сила Ncr в первой главной плоскости, Н.
Public StabilityNcr2 As Double ' Критическая сила Ncr во второй главной плоскости, Н.
Public StabilityNultimate1 As Double ' Предельная сжимающая сила Nult,stab в первой главной плоскости, Н.
Public StabilityNultimate2 As Double ' Предельная сжимающая сила Nult,stab во второй главной плоскости, Н.
Public StabilityReserve1 As Double ' Запас по первой главной плоскости.
Public StabilityReserve2 As Double ' Запас по второй главной плоскости.
Public StabilityEta1 As Double ' Коэффициент увеличения момента eta в первой главной плоскости.
Public StabilityEta2 As Double ' Коэффициент увеличения момента eta во второй главной плоскости.
Public StabilityPhiL1 As Double ' Коэффициент длительности для Ncr в первой главной плоскости.
Public StabilityPhiL2 As Double ' Коэффициент длительности для Ncr во второй главной плоскости.
Public StabilityEccentricity1 As Double ' Начальный расчетный эксцентриситет в первой главной плоскости, мм.
Public StabilityEccentricity2 As Double ' Начальный расчетный эксцентриситет во второй главной плоскости, мм.
Public StabilityAccidentalEcc1 As Double ' Случайный эксцентриситет в первой главной плоскости, мм.
Public StabilityAccidentalEcc2 As Double ' Случайный эксцентриситет во второй главной плоскости, мм.
Public StabilityEffectiveLength1 As Double ' Расчетная длина l0 первой главной плоскости, мм.
Public StabilityEffectiveLength2 As Double ' Расчетная длина l0 второй главной плоскости, мм.
Public StabilitySlenderness1 As Double ' Гибкость l0/i первой главной плоскости.
Public StabilitySlenderness2 As Double ' Гибкость l0/i второй главной плоскости.
Public StabilityPhiLTable1 As Double ' Табличный phi_l СП 35 первой плоскости для ветви ec/r <= 1.
Public StabilityPhiLTable2 As Double ' Табличный phi_l СП 35 второй плоскости для ветви ec/r <= 1.
Public StabilityPhiM1 As Double ' Табличный phi_m СП 35 первой плоскости.
Public StabilityPhiM2 As Double ' Табличный phi_m СП 35 второй плоскости.
Public StabilityPhiValue1 As Double ' Итоговый phi СП 35 первой плоскости для ветви ec <= r.
Public StabilityPhiValue2 As Double ' Итоговый phi СП 35 второй плоскости для ветви ec <= r.
Public StabilityPhiP1 As Double ' Коэффициент phi_p СП 35 первой плоскости для ветви ec > r.
Public StabilityPhiP2 As Double ' Коэффициент phi_p СП 35 второй плоскости для ветви ec > r.
Public StabilityDepth1 As Double ' Размер сечения h1 по нормали первой главной плоскости, мм.
Public StabilityDepth2 As Double ' Размер сечения h2 по нормали второй главной плоскости, мм.
Public StabilityRadius1 As Double ' Радиус инерции i1 приведенного сечения, мм.
Public StabilityRadius2 As Double ' Радиус инерции i2 приведенного сечения, мм.
Public StabilityCoreDistance1 As Double ' Ядровое расстояние r первой главной плоскости со стороны эксцентриситета, мм.
Public StabilityCoreDistance2 As Double ' Ядровое расстояние r второй главной плоскости со стороны эксцентриситета, мм.
Public StabilityDelta1 As Double ' Коэффициент delta или delta_e первой плоскости в выбранной ветви.
Public StabilityDelta2 As Double ' Коэффициент delta или delta_e второй плоскости в выбранной ветви.
Public StabilityKb1 As Double ' Коэффициент kb СП 63 первой плоскости.
Public StabilityKb2 As Double ' Коэффициент kb СП 63 второй плоскости.
Public StabilityKs1 As Double ' Коэффициент ks СП 63 первой плоскости.
Public StabilityKs2 As Double ' Коэффициент ks СП 63 второй плоскости.
Public StabilityStiffnessD1 As Double ' Условная жесткость D первой главной плоскости, Н*мм2.
Public StabilityStiffnessD2 As Double ' Условная жесткость D второй главной плоскости, Н*мм2.
Public StabilityPlaneBranch1 As String ' Фактическая ветвь расчета первой главной плоскости.
Public StabilityPlaneBranch2 As String ' Фактическая ветвь расчета второй главной плоскости.
Public StabilityPlaneApplicable1 As Boolean ' True, если первая плоскость участвовала в проверке.
Public StabilityPlaneApplicable2 As Boolean ' True, если вторая плоскость участвовала в проверке.
Public StabilityPlanePassed1 As Boolean ' True, если проверка первой плоскости прошла.
Public StabilityPlanePassed2 As Boolean ' True, если проверка второй плоскости прошла.
Public GoverningUtil As Double ' Наибольшее использование LC по применимым проверкам.
Public StateAvailable As Boolean ' True, если есть сохраненное прямое состояние НДС.
Public StateConverged As Boolean ' True, если прямой CSectionSolver сошелся.
Public Epsilon0 As Double ' epsilon0 прямого состояния НДС.
Public KappaX As Double ' kappaX прямого состояния НДС.
Public KappaY As Double ' kappaY прямого состояния НДС.
Public Nint As Double ' Внутренняя N в найденном прямом НДС.
Public Mxint As Double ' Внутренний Mx в найденном прямом НДС.
Public Myint As Double ' Внутренний My в найденном прямом НДС.
Public ResidualN As Double ' Остаточная невязка по N.
Public ResidualMx As Double ' Остаточная невязка по Mx.
Public ResidualMy As Double ' Остаточная невязка по My.
Public StateIterations As Long ' Число итераций прямого StateSolution.

Private mStates As Collection ' Конечные именованные НДС, которые должны попасть в Results snapshot.

' Полностью очищает результат LC перед новым расчетом.
' Исходные нагрузки и комментарий сочетания здесь не хранятся, поэтому метод
' не может случайно стереть пользовательский ввод.
Public Sub Clear(ByVal defaultCapacityPath As String)
    Set mStates = New Collection
    Status = vbNullString
    ProfileId = vbNullString
    DirectStateStatus = vbNullString
    CapacityStatus = vbNullString
    CapacityLimitState = vbNullString
    CapacitySolutionMethod = vbNullString
    CrackStatus = vbNullString
    LongitudinalCrackStatus = vbNullString
    StabilityStatus = vbNullString
    ExtensionUsed = False
    DirectStateDiagramExtensionEnabled = False
    LambdaCapacity = 0#
    UtilCapacity = 0#
    CapacityPathResolved = defaultCapacityPath
    NUltimate = 0#
    MxUltimate = 0#
    MyUltimate = 0#
    CapacityStateN = 0#
    CapacityStateMx = 0#
    CapacityStateMy = 0#
    MomentUltimate = 0#
    CapacityStateAvailable = False
    CapacityEpsilon0 = 0#
    CapacityKappaX = 0#
    CapacityKappaY = 0#
    MinConcreteStrain = 0#
    MaxConcreteStrain = 0#
    MaxConcreteCompressionStress = 0#
    MinSteelStrain = 0#
    MaxSteelStrain = 0#
    ClearCrackResult
    ClearLongitudinalCrackResult
    ClearStabilityResult
    GoverningUtil = 0#
    StateAvailable = False
    StateConverged = False
    Epsilon0 = 0#
    KappaX = 0#
    KappaY = 0#
    Nint = 0#
    Mxint = 0#
    Myint = 0#
    ResidualN = 0#
    ResidualMx = 0#
    ResidualMy = 0#
    StateIterations = 0
End Sub

' Сохраняет результат изолированного фильтра продольного изгиба/устойчивости.
' Калькулятор не знает о batch-status policy, поэтому сюда копируются уже
' нормализованные численные величины и короткий пользовательский статус.
Public Sub StoreStabilityResult(ByVal statusText As String, ByVal stability As CStabilityCalculator, _
        Optional ByVal referenceX As Double = 0#, Optional ByVal referenceY As Double = 0#)
    StabilityStatus = statusText
    StabilityReferenceX = referenceX
    StabilityReferenceY = referenceY
    If stability Is Nothing Then Exit Sub
    StabilityCode = stability.Code
    StabilityBranch = stability.Branch
    If Len(Trim$(stability.StopReason)) > 0 Then StabilityBranch = stability.StopReason
    StabilityDesignN = stability.DesignN
    StabilityDesignMx = stability.DesignMx
    StabilityDesignMy = stability.DesignMy
    StabilitySustainedN = stability.SustainedN
    StabilityMoment1 = stability.Moment1
    StabilityMoment2 = stability.Moment2
    StabilitySustainedMoment1 = stability.SustainedMoment1
    StabilitySustainedMoment2 = stability.SustainedMoment2
    StabilityDesignMoment1 = stability.DesignMoment1
    StabilityDesignMoment2 = stability.DesignMoment2
    StabilityCriticalForce = stability.CriticalForce
    StabilityReserveFactor = stability.ReserveFactor
    If StabilityReserveFactor > 0# Then StabilityUtil = 1# / StabilityReserveFactor
    StabilityNcr1 = stability.Ncr1
    StabilityNcr2 = stability.Ncr2
    StabilityNultimate1 = stability.Nultimate1
    StabilityNultimate2 = stability.Nultimate2
    StabilityReserve1 = stability.Reserve1
    StabilityReserve2 = stability.Reserve2
    StabilityEta1 = stability.Eta1
    StabilityEta2 = stability.Eta2
    StabilityPhiL1 = stability.PhiL1
    StabilityPhiL2 = stability.PhiL2
    StabilityPhiLTable1 = stability.PhiLTable1
    StabilityPhiLTable2 = stability.PhiLTable2
    StabilityPhiM1 = stability.PhiM1
    StabilityPhiM2 = stability.PhiM2
    StabilityPhiValue1 = stability.PhiValue1
    StabilityPhiValue2 = stability.PhiValue2
    StabilityPhiP1 = stability.PhiP1
    StabilityPhiP2 = stability.PhiP2
    StabilityEccentricity1 = stability.Eccentricity1
    StabilityEccentricity2 = stability.Eccentricity2
    StabilityAccidentalEcc1 = stability.AccidentalEccentricity1
    StabilityAccidentalEcc2 = stability.AccidentalEccentricity2
    StabilityEffectiveLength1 = stability.EffectiveLength1
    StabilityEffectiveLength2 = stability.EffectiveLength2
    StabilitySlenderness1 = stability.Slenderness1
    StabilitySlenderness2 = stability.Slenderness2
    StabilityDepth1 = stability.Depth1
    StabilityDepth2 = stability.Depth2
    StabilityRadius1 = stability.Radius1
    StabilityRadius2 = stability.Radius2
    StabilityCoreDistance1 = stability.CoreDistance1
    StabilityCoreDistance2 = stability.CoreDistance2
    StabilityDelta1 = stability.Delta1
    StabilityDelta2 = stability.Delta2
    StabilityKb1 = stability.Kb1
    StabilityKb2 = stability.Kb2
    StabilityKs1 = stability.Ks1
    StabilityKs2 = stability.Ks2
    StabilityStiffnessD1 = stability.StiffnessD1
    StabilityStiffnessD2 = stability.StiffnessD2
    StabilityPlaneBranch1 = stability.PlaneBranch1
    StabilityPlaneBranch2 = stability.PlaneBranch2
    StabilityPlaneApplicable1 = stability.PlaneApplicable1
    StabilityPlaneApplicable2 = stability.PlaneApplicable2
    StabilityPlanePassed1 = stability.PlanePassed1
    StabilityPlanePassed2 = stability.PlanePassed2
End Sub

' Обновляет итоговые усилия после общего фильтра практически нулевых моментов.
' CStabilityCalculator возвращает предварительный DesignLoad, а batch затем
' применяет тот же zero-filter, через который идут DirectState/Capacity/Crack.
' Поэтому Results должен хранить именно этот финальный вектор.
Public Sub UpdateStabilityDesignLoads(ByVal designN As Double, ByVal designMx As Double, _
        ByVal designMy As Double, ByVal principalAngleRad As Double)
    StabilityDesignN = designN
    StabilityDesignMx = designMx
    StabilityDesignMy = designMy
    StabilityDesignMoment1 = designMx * Cos(principalAngleRad) + designMy * Sin(principalAngleRad)
    StabilityDesignMoment2 = -designMx * Sin(principalAngleRad) + designMy * Cos(principalAngleRad)
End Sub

' Очищает только блок продольного изгиба/устойчивости.
Public Sub ClearStabilityResult()
    StabilityStatus = vbNullString
    StabilityCode = vbNullString
    StabilityBranch = vbNullString
    StabilityReferenceX = 0#
    StabilityReferenceY = 0#
    StabilityDesignN = 0#
    StabilityDesignMx = 0#
    StabilityDesignMy = 0#
    StabilitySustainedN = 0#
    StabilityMoment1 = 0#
    StabilityMoment2 = 0#
    StabilitySustainedMoment1 = 0#
    StabilitySustainedMoment2 = 0#
    StabilityDesignMoment1 = 0#
    StabilityDesignMoment2 = 0#
    StabilityCriticalForce = 0#
    StabilityReserveFactor = 0#
    StabilityUtil = 0#
    StabilityNcr1 = 0#
    StabilityNcr2 = 0#
    StabilityNultimate1 = 0#
    StabilityNultimate2 = 0#
    StabilityReserve1 = 0#
    StabilityReserve2 = 0#
    StabilityEta1 = 0#
    StabilityEta2 = 0#
    StabilityPhiL1 = 0#
    StabilityPhiL2 = 0#
    StabilityPhiLTable1 = 0#
    StabilityPhiLTable2 = 0#
    StabilityPhiM1 = 0#
    StabilityPhiM2 = 0#
    StabilityPhiValue1 = 0#
    StabilityPhiValue2 = 0#
    StabilityPhiP1 = 0#
    StabilityPhiP2 = 0#
    StabilityEccentricity1 = 0#
    StabilityEccentricity2 = 0#
    StabilityAccidentalEcc1 = 0#
    StabilityAccidentalEcc2 = 0#
    StabilityEffectiveLength1 = 0#
    StabilityEffectiveLength2 = 0#
    StabilitySlenderness1 = 0#
    StabilitySlenderness2 = 0#
    StabilityDepth1 = 0#
    StabilityDepth2 = 0#
    StabilityRadius1 = 0#
    StabilityRadius2 = 0#
    StabilityCoreDistance1 = 0#
    StabilityCoreDistance2 = 0#
    StabilityDelta1 = 0#
    StabilityDelta2 = 0#
    StabilityKb1 = 0#
    StabilityKb2 = 0#
    StabilityKs1 = 0#
    StabilityKs2 = 0#
    StabilityStiffnessD1 = 0#
    StabilityStiffnessD2 = 0#
    StabilityPlaneBranch1 = vbNullString
    StabilityPlaneBranch2 = vbNullString
    StabilityPlaneApplicable1 = False
    StabilityPlaneApplicable2 = False
    StabilityPlanePassed1 = False
    StabilityPlanePassed2 = False
End Sub

' Добавляет конечное named-state LC. Промежуточные итерации solver-а и
' probe-точки сюда не пишутся; состояния BeforeMcrc/AfterMcrc сохраняются,
' если обязательная проверка образования трещины реально дошла до lambda_crc.
Public Sub AddState(ByVal stateResult As CSectionStateResult)
    If stateResult Is Nothing Then Exit Sub
    If mStates Is Nothing Then Set mStates = New Collection
    mStates.Add stateResult
End Sub

' Возвращает количество сохраненных named states.
Public Property Get StateCount() As Long
    If Not mStates Is Nothing Then StateCount = mStates.Count
End Property

' Возвращает named-state по порядковому номеру.
Public Property Get StateAt(ByVal index As Long) As CSectionStateResult
    If mStates Is Nothing Then Err.Raise vbObjectError + 3940, "CCombinationResult", "Коллекция states не инициализирована."
    If index < 1 Or index > mStates.Count Then Err.Raise vbObjectError + 3941, "CCombinationResult", "Индекс named-state вне диапазона."
    Set StateAt = mStates.Item(index)
End Property

' Ищет named-state по типу. Если состояние не рассчитывалось, возвращает Nothing.
Public Function FindState(ByVal stateType As ESectionStateType) As CSectionStateResult
    If mStates Is Nothing Then Exit Function
    Dim i As Long
    For i = 1 To mStates.Count
        Dim stateResult As CSectionStateResult
        Set stateResult = mStates.Item(i)
        If stateResult.StateType = stateType Then
            Set FindState = stateResult
            Exit Function
        End If
    Next i
End Function

' Сохраняет найденную плоскость деформаций и внутренние усилия прямого НДС.
' Эти данные используются Results snapshot, Excel-схемой и AutoCAD export.
Public Sub StoreSectionState(ByVal solver As CSectionSolver)
    StateAvailable = True
    StateConverged = solver.Converged
    Epsilon0 = solver.Epsilon0
    KappaX = solver.KappaX
    KappaY = solver.KappaY
    Nint = solver.Nint
    Mxint = solver.Mxint
    Myint = solver.Myint
    MinConcreteStrain = solver.MinConcreteStrain
    MaxConcreteStrain = solver.MaxConcreteStrain
    MaxConcreteCompressionStress = solver.MaxConcreteCompressionStress
    MinSteelStrain = solver.MinSteelStrain
    MaxSteelStrain = solver.MaxSteelStrain
    ResidualN = solver.ResidualN
    ResidualMx = solver.ResidualMx
    ResidualMy = solver.ResidualMy
    StateIterations = solver.Iterations
End Sub

' Копирует инженерные величины расчета трещин из специализированного
' calculator-а в результат LC. Пользовательский CrackStatus выставляет
' CBatchStatusPolicy, поэтому здесь он не определяется.
Public Sub StoreCrackResult(ByVal crack As CCrackWidthCalculator)
    CrackWidth = crack.CrackWidth
    CrackAllowableValue = crack.AllowableCrackWidth
    CrackUtil = crack.Utilization
    CrackFormed = crack.CrackFormed
    CrackPhi1 = crack.Phi1
    CrackPhi2 = crack.Phi2
    CrackPhi3 = crack.Phi3
    CrackPsiS = crack.PsiS
    CrackSigmaS = crack.SigmaS
    CrackSigmaSCrc = crack.SigmaSCrc
    CrackAs = crack.AsTension
    CrackAbt = crack.Abt
    CrackDsEquivalent = crack.DsEquivalent
    CrackLsRaw = crack.CrackSpacingRaw
    CrackLs = crack.CrackSpacing
    CrackLambdaCrc = crack.LambdaCrc
    CrackFormationMethod = crack.CrackFormationMethod
    CrackNcrc = crack.Ncrc
    CrackFormationNcrc = crack.FormationNcrc
    CrackMcrc = crack.Mcrc
    CrackAred = crack.Ared
    CrackRbtSer = crack.RbtSer
    CrackTensionRebarCount = crack.TensionRebarCount
    CrackTensionRebarIds = crack.TensionRebarIds
    CrackSectionDepthH = crack.SectionDepthH
    CrackCoverA = crack.CoverA
    CrackTensionDepth = crack.TensionDepth
    CrackEffectiveZoneDepth = crack.EffectiveZoneDepth
    CrackCentralBranch = crack.CentralTensionBranch
End Sub

' Сохраняет результат проверки продольных трещин по сжатому бетону.
' Проверка использует уже найденные напряжения прямого НДС и не запускает
' отдельный material purpose или повторный solve.
Public Sub StoreLongitudinalCrackResult(ByVal statusText As String, ByVal maxCompressionStress As Double, _
        ByVal rbMc2 As Double, ByVal utilization As Double)
    LongitudinalCrackStatus = statusText
    MaxConcreteCompressionStress = maxCompressionStress
    LongitudinalCrackRbMc2 = rbMc2
    LongitudinalCrackUtil = utilization
End Sub

' Очищает только проверку продольных трещин. N/A здесь означает, что для LC
' нет сжатого бетона или прямое НДС не было получено.
Public Sub ClearLongitudinalCrackResult()
    LongitudinalCrackStatus = vbNullString
    LongitudinalCrackRbMc2 = 0#
    LongitudinalCrackUtil = 0#
End Sub

' Очищает только блок трещин. Это удобно для режимов, где прямое НДС и capacity
' уже заполнены, а расчет раскрытия неприменим или отключен.
Public Sub ClearCrackResult()
    CrackWidth = 0#
    CrackAllowableValue = 0#
    CrackUtil = 0#
    CrackFormed = False
    CrackPhi1 = 0#
    CrackPhi2 = 0#
    CrackPhi3 = 0#
    CrackPsiS = 0#
    CrackSigmaS = 0#
    CrackSigmaSCrc = 0#
    CrackAs = 0#
    CrackAbt = 0#
    CrackDsEquivalent = 0#
    CrackLsRaw = 0#
    CrackLs = 0#
    CrackLambdaCrc = 0#
    CrackFormationMethod = vbNullString
    CrackNcrc = 0#
    CrackFormationNcrc = 0#
    CrackMcrc = 0#
    CrackAred = 0#
    CrackRbtSer = 0#
    CrackTensionRebarCount = 0
    CrackTensionRebarIds = vbNullString
    CrackSectionDepthH = 0#
    CrackCoverA = 0#
    CrackTensionDepth = 0#
    CrackEffectiveZoneDepth = 0#
    CrackCentralBranch = False
End Sub
