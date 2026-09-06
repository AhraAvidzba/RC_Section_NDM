VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCrackWidthCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Расчет ширины раскрытия нормальных трещин по II группе
' ==========================================================================
' Класс реализует принятую для проекта НДМ-интерпретацию формулы СП 63 для
' продолжительного раскрытия нормальных трещин при общем состоянии N + Mx + My.
' Он не знает тип сечения: вся геометрия приходит как CSectionModel, а НДС
' находится штатным CSectionSolver.
'
' Текущая версия считает только продолжительное раскрытие по формуле
' a_crc = a_crc1. Коэффициенты phi1, phi2, phi3 и psi_s задаются в Config:
' phi3 и psi_s могут быть приняты пользователем либо определены автоматически.
' Способ определения расстояния a_s задается отдельно: локально до ближайшего
' бетонного контура или глобально до дальней растянутой линии всего сечения.
'
' В режиме User ширина считается по переданному НДС текущего сочетания:
' повторный solve не выполняется, а psi_s берется из настройки.
' В режиме Auto расчет сначала выполняется с psi_s = 1. Только если
' первая проверка не проходит, класс ищет состояние BeforeMcrcState с
' растянутым бетоном, затем отдельно решает AfterMcrcState без растянутого
' бетона. Именно AfterMcrcState дает sigma_s,crc для уточнения psi_s.

Private Const PHI1_LONG_TERM As Double = 1.4 ' Коэффициент длительного действия нагрузки для a_crc1 по СП 63.
Private Const PHI2_DEFAULT As Double = 0.5 ' Типовой default для обычной арматуры периодического профиля.
Private Const PHI3_TENSION As Double = 1.2 ' Коэффициент для центрального/сквозного растяжения.
Private Const PHI3_BENDING As Double = 1# ' Коэффициент для изгиба и внецентренного сжатия.
Private Const CRACK_DIRECTION_TOLERANCE As Double = 0.001 ' Допуск сохранения направления вектора N/Mx/My в прямом поиске lambda_crc.
Private Const CRACK_STRAIN_RATIO_TOLERANCE As Double = 0.000001 ' Допуск попадания в eps_bt,ult как в безразмерное условие.
Private Const NEAR_ZERO_STRAIN As Double = 0.000000000001 ' Численный ноль для деформаций и кривизн.
Private Const NEAR_ZERO_FORCE As Double = 0.000001 ' Численный ноль для отличия центрального растяжения от нулевой нагрузки.
Private Const LARGE_LAMBDA As Double = 1E+30 ' Диагностическая отметка: трещина не достигнута при lambda = 1.
Private Const COVER_MODE_NEAREST_CONTOUR As String = "NearestContour" ' a_s до ближайшего выхода бетонного контура по нормали.
Private Const COVER_MODE_GLOBAL_EXTREME As String = "GlobalExtreme" ' a_s до дальней растянутой опорной линии всего сечения.

Private mConverged As Boolean ' True, если расчет трещины завершился без ошибки входных данных или несходимости.
Private mStopReason As String ' Короткое пояснение результата: расчет выполнен, трещина не образована или solver не сошелся.
Private mDiagnosticLog As String ' Текстовая диагностика выбора зоны, арматуры, ls и psi_s.
Private mDiagnosticsEnabled As Boolean ' Отключает подробный журнал, когда txt-отчет выполнения не нужен.

Private mAllowableCrackWidth As Double ' Пользовательский предел a_crc,ult во внутренних единицах длины, мм.
Private mCrackWidth As Double ' Расчетная ширина раскрытия a_crc, мм.
Private mUtilization As Double ' Коэффициент использования по трещине: a_crc / a_crc,ult.
Private mCrackFormed As Boolean ' True, если по выбранной ветке расчета трещина считается образованной.

Private mPsiMode As String ' User или Auto: брать psi_s из Config или уточнять после неудачной первой проверки.
Private mPhi3Mode As String ' User или Auto: брать phi3 из Config или выбирать по характеру продольной силы.
Private mTensionZoneMode As String ' Effective или FullTension: способ задания бетонной зоны Abt.
Private mCoverMode As String ' NearestContour или GlobalExtreme: способ измерения расстояния a_s.
Private mPhi1Input As Double ' Пользовательское значение phi1 для формулы раскрытия.
Private mPhi2Input As Double ' Пользовательское значение phi2 для формулы раскрытия.
Private mPhi3Input As Double ' Пользовательское значение phi3, если Phi3Mode = User.
Private mPsiSInput As Double ' Пользовательское значение psi_s, если PsiMode = User.
Private mConcreteRbtSer As Double ' Rbt,ser для центральной ветки Ncrc и справочных расчетов Mcrc, МПа.
Private mConcreteEb As Double ' Eb для Ared в центральной ветке Ncrc, МПа.
Private mConcreteEbt0 As Double ' eps_bt0 для формулы (8.54) СП 63 при однозначно растянутой эпюре.
Private mConcreteEbt2 As Double ' eps_bt2: предельная растягивающая деформация при двузначной эпюре Mcrc.
Private mSteelEs As Double ' Es для формулы раскрытия трещины, МПа.
Private mCrackInitiationSpec As CMaterialModelSpec ' Спецификация модели Mcrc, если Auto потребует lambda_crc.
Private mCrackedStateSpec As CMaterialModelSpec ' Спецификация модели раскрытой трещины текущего LC.
Private mBeforeMcrcState As CSectionStateResult ' Необязательное состояние до выключения растянутого бетона.
Private mAfterMcrcState As CSectionStateResult ' Необязательное состояние после выключения растянутого бетона для sigma_s,crc.

Private mSolverMethod As String ' Newton или Secant для внутренних прямых расчетов состояния сечения.
Private mSolverLoadSteps As Long ' Число ступеней приложения нагрузки во внутреннем CSectionSolver.
Private mSolverMaxIterations As Long ' Максимум итераций на ступень для внутренних solve.
Private mSolverToleranceN As Double ' Допуск равновесия по N, Н.
Private mSolverToleranceMx As Double ' Допуск равновесия по Mx, Н*мм.
Private mSolverToleranceMy As Double ' Допуск равновесия по My, Н*мм.
Private mSolverLineSearchEnabled As Boolean ' Разрешает line search во внутренних solve.
Private mSolverDampingInitial As Double ' Начальный alpha line search.
Private mSolverMinLineSearchAlpha As Double ' Минимальный alpha line search.
Private mSolverMaxDeltaEpsilon0 As Double ' Ограничение шага по epsilon0.
Private mSolverMaxDeltaKappa As Double ' Ограничение шага по кривизнам.
Private mSolverSecantMaxRestarts As Long ' Число перезапусков секущей матрицы.
Private mSolverSecantMinStepNorm As Double ' Минимальный шаг Secant до признания застоя.

Private mTensionRebarCount As Long ' Сколько стержней попало в расчетную растянутую арматуру As.
Private mMaxSteelStrain As Double ' Максимальная растягивающая деформация среди выбранных стержней As.
Private mMaxSteelStress As Double ' Напряжение в стержне с максимальной растягивающей деформацией, МПа.
Private mSigmaS As Double ' Средневзвешенное напряжение выбранной растянутой арматуры при текущей нагрузке, МПа.
Private mSigmaSCrc As Double ' То же напряжение в состоянии lambda_crc * LC после образования трещины, МПа.
Private mAs As Double ' Площадь выбранной растянутой арматуры, мм2.
Private mAbt As Double ' Площадь бетона, попавшая в расчетную растянутую зону, мм2.
Private mDsEquivalent As Double ' Эквивалентный диаметр выбранных стержней для разных d, мм.
Private mLsRaw As Double ' Нормативная длина ls до ограничений 10ds/100/40ds/400, мм.
Private mLs As Double ' Итоговое расстояние между трещинами после ограничений, мм.
Private mPsiS As Double ' Коэффициент неравномерности деформаций арматуры между трещинами.
Private mPhi1 As Double ' Коэффициент длительности действия нагрузки.
Private mPhi2 As Double ' Пользовательский коэффициент, учитывающий поверхность арматуры.
Private mPhi3 As Double ' Коэффициент характера нагружения.
Private mLambdaCrc As Double ' Множитель нагрузки на пороге трещины: eps_bt,ult для общей ветки или Ncrc/N для центрального растяжения.
Private mNcrc As Double ' Усилие образования трещины для ветки истинного центрального растяжения, Н.

Private mSectionDepthH As Double ' Габарит сечения по нормали к нейтральной линии по общему бетонному контуру КЭ.
Private mCoverA As Double ' Расстояние a_s по выбранному режиму CoverDistanceMode, мм.
Private mTensionDepth As Double ' Фактическая глубина растянутой бетонной зоны до нейтральной линии, мм.
Private mEffectiveZoneDepth As Double ' Глубина зоны Abt, реально использованная после ограничений, мм.
Private mCentralTensionBranch As Boolean ' True, если сработала отдельная ветка центрального растяжения.

' ------------------------------
' Настройки по умолчанию
' ------------------------------
' Значения соответствуют текущей принятой постановке: длительное раскрытие,
' эффективная зона СП и пользовательские коэффициенты с нормативными defaults.
Private Sub Class_Initialize()
    mAllowableCrackWidth = 0.3
    mPsiMode = "User"
    mPhi3Mode = "Auto"
    mTensionZoneMode = "Effective"
    mCoverMode = COVER_MODE_GLOBAL_EXTREME
    mPhi1Input = PHI1_LONG_TERM
    mPhi2Input = PHI2_DEFAULT
    mPhi3Input = PHI3_BENDING
    mPsiSInput = 1#
    mConcreteRbtSer = 1.8
    mConcreteEb = 32500#
    mSteelEs = 200000#
    mSolverMethod = "Newton"
    mSolverLoadSteps = 1
    mSolverMaxIterations = 40
    mSolverToleranceN = 1#
    mSolverToleranceMx = 1000#
    mSolverToleranceMy = 1000#
    mSolverLineSearchEnabled = True
    mSolverDampingInitial = 1#
    mSolverMinLineSearchAlpha = 0.03125
    mSolverMaxDeltaEpsilon0 = 0.0005
    mSolverMaxDeltaKappa = 0.00001
    mSolverSecantMaxRestarts = 2
    mSolverSecantMinStepNorm = 0.000000000001
    mDiagnosticsEnabled = True
End Sub

' ------------------------------
' Чтение настроек Config
' ------------------------------
' Загружает только те настройки, которые нужны расчету трещин и внутренним
' прямым solve. Все пользовательские единицы сразу переводятся во внутренние
' мм, Н, Н*мм и МПа; дальше класс работает как часть расчетного ядра.
Public Sub ApplySettings(ByVal settings As CSystemSettingsReader, Optional ByVal units As CUnitSystem = Nothing)
    If settings Is Nothing Then Exit Sub

    mAllowableCrackWidth = settings.GetDouble("SLS.Crack.Allowable", mAllowableCrackWidth)
    mPsiMode = settings.GetString("SLS.Crack.PsiMode", mPsiMode)
    mTensionZoneMode = settings.GetString("SLS.Crack.TensionZoneMode", mTensionZoneMode)
    If settings.HasKey("SLS.Crack.CoverDistanceMode") Then
        mCoverMode = settings.GetString("SLS.Crack.CoverDistanceMode", mCoverMode)
    Else
        mCoverMode = settings.GetString("SLS.Crack.CoverMode", mCoverMode)
    End If
    mPhi1Input = settings.GetDouble("SLS.Crack.Phi1", mPhi1Input)
    mPhi2Input = settings.GetDouble("SLS.Crack.Phi2", mPhi2Input)
    mPhi3Mode = settings.GetString("SLS.Crack.Phi3Mode", mPhi3Mode)
    mPhi3Input = settings.GetDouble("SLS.Crack.Phi3", mPhi3Input)
    mPsiSInput = settings.GetDouble("SLS.Crack.PsiS", mPsiSInput)

    mSolverMaxIterations = settings.GetLong("Solver.MaxIterations", mSolverMaxIterations)
    mSolverLoadSteps = settings.GetLong("Solver.LoadSteps", mSolverLoadSteps)
    mSolverToleranceN = settings.GetDouble("Solver.ToleranceN", mSolverToleranceN)
    mSolverToleranceMx = settings.GetDouble("Solver.ToleranceMx", mSolverToleranceMx)
    mSolverToleranceMy = settings.GetDouble("Solver.ToleranceMy", mSolverToleranceMy)
    mSolverLineSearchEnabled = settings.GetBoolean("Solver.LineSearchEnabled", mSolverLineSearchEnabled)
    mSolverDampingInitial = settings.GetDouble("Solver.DampingInitial", mSolverDampingInitial)
    mSolverMinLineSearchAlpha = settings.GetDouble("Solver.MinLineSearchAlpha", mSolverMinLineSearchAlpha)
    mSolverMaxDeltaEpsilon0 = settings.GetDouble("Solver.MaxDeltaEpsilon0", mSolverMaxDeltaEpsilon0)
    mSolverMaxDeltaKappa = settings.GetDouble("Solver.MaxDeltaKappa", mSolverMaxDeltaKappa)
    If settings.HasKey("Solver.Method") Then mSolverMethod = settings.GetRawString("Solver.Method", mSolverMethod)
    mSolverSecantMaxRestarts = settings.GetLong("Solver.SecantMaxRestarts", mSolverSecantMaxRestarts)
    mSolverSecantMinStepNorm = settings.GetDouble("Solver.SecantMinStepNorm", mSolverSecantMinStepNorm)

    If Not units Is Nothing Then
        mAllowableCrackWidth = units.InputLengthToInternal(mAllowableCrackWidth)
        mSolverToleranceN = Abs(units.InputForceToInternal(mSolverToleranceN))
        mSolverToleranceMx = Abs(units.InputMomentMxToInternal(mSolverToleranceMx))
        mSolverToleranceMy = Abs(units.InputMomentMyToInternal(mSolverToleranceMy))
        mSolverMaxDeltaKappa = Abs(units.InputCurvatureToInternal(mSolverMaxDeltaKappa))
    End If
End Sub

' ------------------------------
' Основной расчет
' ------------------------------
' Метод получает уже найденное НДС текущего сочетания и, в режиме User,
' использует его напрямую без повторного вызова CSectionSolver. Дополнительные
' solve-ы нужны только для Auto и только после неудачной первой проверки
' раскрытия с psi_s = 1.
Public Sub Calculate(ByVal currentSolver As CSectionSolver, ByVal section As CSectionModel, _
        ByVal materialProvider As CMaterialModelProvider, ByVal crackedStateSpec As CMaterialModelSpec, _
        ByVal crackInitiationSpec As CMaterialModelSpec, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double, _
        Optional ByVal centroidMxForCentralCheck As Variant, Optional ByVal centroidMyForCentralCheck As Variant)
    ClearResult
    ValidateInputs currentSolver, section, materialProvider
    If crackedStateSpec Is Nothing Or Not crackedStateSpec.IsComplete Then _
        Err.Raise vbObjectError + 3714, "CCrackWidthCalculator", "MaterialModel.CrackedState не передан."
    Set mCrackedStateSpec = crackedStateSpec.Clone
    If Not crackInitiationSpec Is Nothing Then
        If crackInitiationSpec.IsComplete Then Set mCrackInitiationSpec = crackInitiationSpec.Clone
    End If
    ApplyMaterialProviderProperties materialProvider
    NormalizeSettings

    If IsCentralTensionState(targetN, centroidMxForCentralCheck, centroidMyForCentralCheck) Then
        CalculateCentralTension currentSolver, section, materialProvider, targetN, targetMx, targetMy
    Else
        CalculateGeneralBending currentSolver, section, materialProvider, targetN, targetMx, targetMy
    End If
End Sub

' Берет Rbt,ser, Eb, eps_bt0, eps_bt2 и Es из material provider-а.
' Эти значения больше не являются отдельными настройками crack-класса:
' они должны совпадать с диаграммами, которыми реально решается НДС.
Private Sub ApplyMaterialProviderProperties(ByVal materialProvider As CMaterialModelProvider)
    If Not mCrackInitiationSpec Is Nothing Then
        mConcreteRbtSer = materialProvider.ConcreteTensionResistanceFromSpec(mCrackInitiationSpec)
        mConcreteEb = materialProvider.ConcreteElasticModulusFromSpec(mCrackInitiationSpec)
        mConcreteEbt0 = materialProvider.ConcreteTensionEbt0FromSpec(mCrackInitiationSpec)
        mConcreteEbt2 = materialProvider.ConcreteTensionLimitFromSpec(mCrackInitiationSpec)
    Else
        mConcreteRbtSer = materialProvider.ConcreteTensionResistanceFromSpec(mCrackedStateSpec)
        mConcreteEb = materialProvider.ConcreteElasticModulusFromSpec(mCrackedStateSpec)
        mConcreteEbt0 = 0#
        mConcreteEbt2 = 0#
    End If
    mSteelEs = materialProvider.SteelElasticModulusFromSpec(mCrackedStateSpec)
End Sub

' Выполняет отдельную ветку СП для истинного центрального растяжения.
' Здесь нет нейтральной линии, поэтому Abt принимается равной всей площади
' бетона, а As - всей продольной арматуре, которая растянута в текущем НДС.
Private Sub CalculateCentralTension(ByVal currentSolver As CSectionSolver, ByVal section As CSectionModel, _
        ByVal materialProvider As CMaterialModelProvider, ByVal targetN As Double, _
        ByVal targetMx As Double, ByVal targetMy As Double)
    mCentralTensionBranch = True
    mPhi3 = ResolvePhi3(targetN)

    If targetN <= NEAR_ZERO_FORCE Then
        FinishNoCrack "Центральная ветка: продольная сила не является растягивающей."
        Exit Sub
    End If

    Dim selected() As Boolean
    ReDim selected(1 To section.RebarCount)
    SelectAllTensionRebars currentSolver, section, selected
    If mTensionRebarCount = 0 Then
        FinishNoCrack "В центральном растяжении не найдено растянутой арматуры."
        Exit Sub
    End If

    mAbt = TotalConcreteArea(section)
    mAs = SelectedRebarArea(section, selected)
    mNcrc = CentralTensionCrackForce(section)
    If mNcrc <= 0# Then
        mStopReason = "InputError: для центрального растяжения не удалось определить Ncrc."
        Exit Sub
    End If

    mLambdaCrc = mNcrc / targetN
    If mLambdaCrc > 1# Then
        FinishNoCrack "Трещина не образована: Ncrc больше заданной растягивающей силы."
        Exit Sub
    End If

    CompleteCrackFormula currentSolver, section, materialProvider.SteelMaterialFromSpec(mCrackedStateSpec), selected
    If mStopReason <> vbNullString And Not mCrackFormed Then Exit Sub
    If Not FinalizeWithPsiMode(section, materialProvider, currentSolver, selected, _
            targetN, targetMx, targetMy, "Расчет выполнен для центрального растяжения; Abt принята по всей площади бетона.") Then Exit Sub
End Sub

' Выполняет общий расчет для N + Mx + My при существующей нейтральной линии.
' Все геометрические величины строятся в координате s вдоль нормали к
' нейтральной линии, направленной в сторону растяжения.
Private Sub CalculateGeneralBending(ByVal currentSolver As CSectionSolver, ByVal section As CSectionModel, _
        ByVal materialProvider As CMaterialModelProvider, ByVal targetN As Double, _
        ByVal targetMx As Double, ByVal targetMy As Double)
    mPhi3 = ResolvePhi3(targetN)

    Dim nx As Double
    Dim ny As Double
    Dim normK As Double
    normK = Sqr(currentSolver.KappaX * currentSolver.KappaX + currentSolver.KappaY * currentSolver.KappaY)
    nx = currentSolver.KappaY / normK
    ny = currentSolver.KappaX / normK

    Dim st As Double
    Dim sc As Double
    ProjectionBoundsByConcreteContour section, nx, ny, st, sc
    mSectionDepthH = st - sc
    If mSectionDepthH <= NEAR_ZERO_STRAIN Then
        mStopReason = "InputError: не удалось определить габарит сечения по нормали к нейтральной линии."
        Exit Sub
    End If

    Dim neutralS As Double
    neutralS = -currentSolver.Epsilon0 / normK
    mTensionDepth = st - MaxDouble(neutralS, sc)
    If mTensionDepth <= NEAR_ZERO_STRAIN Then
        FinishNoCrack "В текущем состоянии нет растянутой бетонной зоны."
        Exit Sub
    End If

    Dim sBarMax As Double
    Dim coverRebarIndex As Long
    If Not FindMostTensionRebarProjection(currentSolver, section, nx, ny, sBarMax, coverRebarIndex) Then
        FinishNoCrack "В расчетной растянутой зоне не найдено растянутой арматуры."
        Exit Sub
    End If
    mCoverA = CoverDistanceForRebar(section, coverRebarIndex, nx, ny, st, sBarMax)
    If mCoverA < 0# Then mCoverA = 0#

    Dim zoneLowerS As Double
    zoneLowerS = TensionZoneLowerBound(st, sc, neutralS)
    mEffectiveZoneDepth = st - zoneLowerS
    If mEffectiveZoneDepth <= NEAR_ZERO_STRAIN Then
        FinishNoCrack "Эффективная растянутая зона имеет нулевую глубину."
        Exit Sub
    End If

    mAbt = ConcreteAreaInZone(currentSolver, section, nx, ny, zoneLowerS, _
        StrComp(mTensionZoneMode, "FullTension", vbTextCompare) = 0)
    If mAbt <= 0# Then
        FinishNoCrack "В расчетную растянутую зону не попали бетонные элементы."
        Exit Sub
    End If

    Dim selected() As Boolean
    ReDim selected(1 To section.RebarCount)
    SelectRebarsInZone currentSolver, section, nx, ny, zoneLowerS, selected
    If mTensionRebarCount = 0 Then
        FinishNoCrack "В расчетную растянутую зону не попали растянутые стержни."
        Exit Sub
    End If

    CompleteCrackFormula currentSolver, section, materialProvider.SteelMaterialFromSpec(mCrackedStateSpec), selected
    If mStopReason <> vbNullString And Not mCrackFormed Then Exit Sub
    If Not FinalizeWithPsiMode(section, materialProvider, currentSolver, selected, _
            targetN, targetMx, targetMy, "Расчет выполнен для произвольного сечения по общей НДМ-ветке.") Then Exit Sub
End Sub

' Находит lambda_crc в режиме Auto. Нагрузка масштабируется целиком, чтобы
' направление вектора N/Mx/My оставалось постоянным, а критерием служит
' достижение eps_bt,ult по СП 63 п. 8.2.14 и п. 8.1.30:
' для двузначной эпюры eps_bt,ult = eps_bt2, для однозначно растянутой
' эпюры с ненулевой кривизной применяется формула (8.54).
Private Function CalculateLambdaCrc(ByVal section As CSectionModel, ByVal concreteMaterial As Object, _
        ByVal steelMaterial As Object, ByVal startSolver As CSectionSolver, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double) As Boolean
    On Error GoTo Failed

    If mConcreteEbt0 <= 0# Or mConcreteEbt2 <= 0# Then
        mStopReason = "InputError: для Mcrc в режиме Auto нужны положительные eps_bt0 и eps_bt2 в модели CrackInitiation."
        Exit Function
    End If

    Dim eps0 As Double
    Dim kx As Double
    Dim ky As Double
    InitializeCrcStrainGuess section, startSolver, targetN, targetMx, targetMy, mConcreteEbt2, eps0, kx, ky

    Dim iter As Long
    Dim rDir1 As Double
    Dim rDir2 As Double
    Dim rLimit As Double
    Dim lambdaValue As Double
    Dim solver As CSectionSolver

    ' Auto не подбирает lambda внешней бицекцией. Здесь ищется сама плоскость
    ' деформаций образования трещины: направление полного вектора N/Mx/My
    ' сохраняется, а максимальная растягивающая деформация бетона достигает
    ' предельного значения eps_bt,ult по СП 63 п. 8.1.30.
    For iter = 1 To mSolverMaxIterations
        If Not EvaluateCrcStrainResidual(section, concreteMaterial, steelMaterial, targetN, targetMx, targetMy, _
                eps0, kx, ky, solver, rDir1, rDir2, rLimit, lambdaValue) Then Exit Function

        If Abs(rDir1) <= CRACK_DIRECTION_TOLERANCE And _
                Abs(rDir2) <= CRACK_DIRECTION_TOLERANCE And _
                Abs(rLimit) <= CRACK_STRAIN_RATIO_TOLERANCE Then
            If lambdaValue <= 0# Then
                mStopReason = "NumericalFailure: найденное состояние образования трещины имеет неположительный lambda_crc."
                Exit Function
            End If
            mLambdaCrc = lambdaValue
            Set mBeforeMcrcState = New CSectionStateResult
            mBeforeMcrcState.InitializeFromSolver sstBeforeMcrcState, cpMcrc, _
                mCrackInitiationSpec, solver, False, "OK"
            AppendDiagnostic "lambda_crc=" & FormatNumberInvariant(mLambdaCrc) & _
                "; iterations=" & CStr(iter) & "; method=DirectCrackingStrain"
            CalculateLambdaCrc = True
            Exit Function
        End If

        If Not ApplyCrcNewtonStep(section, concreteMaterial, steelMaterial, targetN, targetMx, targetMy, _
                eps0, kx, ky, rDir1, rDir2, rLimit) Then Exit Function
    Next iter

    mStopReason = "NumericalFailure: прямой поиск lambda_crc по eps_bt,ult не сошелся за заданное число итераций."
    Exit Function

Failed:
    mStopReason = "NumericalFailure: расчет lambda_crc не выполнен: " & Err.Description
End Function


' Строит стартовое приближение для прямого состояния образования трещины.
' Если текущий service-solver уже содержит растяжение бетона, его плоскость
' просто масштабируется до eps_bt2 как устойчивой начальной оценки. Иначе используется грубая геометрическая
' оценка по направлению заданного вектора Mx/My.
Private Sub InitializeCrcStrainGuess(ByVal section As CSectionModel, ByVal startSolver As CSectionSolver, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double, _
        ByVal epsBt2Guess As Double, ByRef eps0 As Double, ByRef kx As Double, ByRef ky As Double)
    If Not startSolver Is Nothing Then
        Dim maxTension As Double
        maxTension = MaxConcreteStrain(startSolver, section)
        If maxTension > NEAR_ZERO_STRAIN Then
            Dim strainScale As Double
            strainScale = epsBt2Guess / maxTension
            eps0 = startSolver.Epsilon0 * strainScale
            kx = startSolver.KappaX * strainScale
            ky = startSolver.KappaY * strainScale
            Exit Sub
        End If
    End If

    Dim area As Double
    Dim maxAbsX As Double
    Dim maxAbsY As Double
    Dim i As Long
    For i = 1 To section.ConcreteCount
        area = area + section.ConcreteArea(i)
        If Abs(section.ConcreteX(i)) > maxAbsX Then maxAbsX = Abs(section.ConcreteX(i))
        If Abs(section.ConcreteY(i)) > maxAbsY Then maxAbsY = Abs(section.ConcreteY(i))
    Next i
    For i = 1 To section.RebarCount
        area = area + section.RebarArea(i)
        If Abs(section.RebarX(i)) > maxAbsX Then maxAbsX = Abs(section.RebarX(i))
        If Abs(section.RebarY(i)) > maxAbsY Then maxAbsY = Abs(section.RebarY(i))
    Next i

    If area > 0# And mConcreteEb > 0# Then eps0 = targetN / (area * mConcreteEb)
    Dim momentNorm As Double
    momentNorm = Sqr(targetMx * targetMx + targetMy * targetMy)
    If momentNorm > 0# Then
        If maxAbsY > 0# Then kx = Sgn(targetMx) * epsBt2Guess / maxAbsY * Abs(targetMx) / momentNorm
        If maxAbsX > 0# Then ky = Sgn(targetMy) * epsBt2Guess / maxAbsX * Abs(targetMy) / momentNorm
    Else
        eps0 = epsBt2Guess
    End If
    If Abs(kx) <= NEAR_ZERO_STRAIN And Abs(ky) <= NEAR_ZERO_STRAIN And maxAbsY > 0# Then kx = epsBt2Guess / maxAbsY
End Sub

' Считает три безразмерных условия для прямого поиска lambda_crc:
' два условия сохраняют направление полного вектора N/Mx/My, третье требует
' достижения eps_bt,ult в наиболее растянутом бетонном конечном элементе.
Private Function EvaluateCrcStrainResidual(ByVal section As CSectionModel, ByVal concreteMaterial As Object, _
        ByVal steelMaterial As Object, ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double, _
        ByVal eps0 As Double, ByVal kx As Double, ByVal ky As Double, _
        ByRef solver As CSectionSolver, ByRef rDir1 As Double, ByRef rDir2 As Double, _
        ByRef rLimit As Double, ByRef lambdaValue As Double) As Boolean
    On Error GoTo Failed

    Set solver = New CSectionSolver
    solver.EvaluateStrainPlane section, concreteMaterial, steelMaterial, eps0, kx, ky

    Dim baseN As Double
    Dim baseMx As Double
    Dim baseMy As Double
    Dim intN As Double
    Dim intMx As Double
    Dim intMy As Double
    baseN = targetN / MaxDouble(mSolverToleranceN, 1#)
    baseMx = targetMx / MaxDouble(mSolverToleranceMx, 1#)
    baseMy = targetMy / MaxDouble(mSolverToleranceMy, 1#)
    intN = solver.Nint / MaxDouble(mSolverToleranceN, 1#)
    intMx = solver.Mxint / MaxDouble(mSolverToleranceMx, 1#)
    intMy = solver.Myint / MaxDouble(mSolverToleranceMy, 1#)

    Dim baseNorm2 As Double
    baseNorm2 = baseN * baseN + baseMx * baseMx + baseMy * baseMy
    If baseNorm2 <= NEAR_ZERO_STRAIN Then
        mStopReason = "InputError: для Auto нужен ненулевой вектор N/Mx/My."
        Exit Function
    End If

    lambdaValue = (intN * baseN + intMx * baseMx + intMy * baseMy) / baseNorm2
    SetCrcDirectionResiduals baseN, baseMx, baseMy, intN, intMx, intMy, rDir1, rDir2

    Dim minConcreteStrain As Double
    Dim maxConcreteStrain As Double
    ConcreteStrainBounds solver, section, minConcreteStrain, maxConcreteStrain

    Dim epsBtUlt As Double
    epsBtUlt = ConcreteTensionUltimateStrain(minConcreteStrain, maxConcreteStrain)
    If epsBtUlt <= 0# Then
        mStopReason = "InputError: не удалось определить eps_bt,ult для состояния образования трещины."
        Exit Function
    End If
    rLimit = maxConcreteStrain / epsBtUlt - 1#

    AppendDiagnostic "lambda_crc_iter=" & FormatNumberInvariant(lambdaValue) & _
        "; rDir1=" & FormatNumberInvariant(rDir1) & _
        "; rDir2=" & FormatNumberInvariant(rDir2) & _
        "; rLimit=" & FormatNumberInvariant(rLimit) & _
        "; eps_bt_ult=" & FormatNumberInvariant(epsBtUlt)
    EvaluateCrcStrainResidual = True
    Exit Function

Failed:
    mStopReason = "NumericalFailure: оценка состояния образования трещины не выполнена: " & Err.Description
End Function

' Возвращает две независимые знаковые невязки направления. Они равны нулю,
' когда внутренний вектор усилий параллелен заданному вектору N/Mx/My.
Private Sub SetCrcDirectionResiduals(ByVal baseN As Double, ByVal baseMx As Double, ByVal baseMy As Double, _
        ByVal intN As Double, ByVal intMx As Double, ByVal intMy As Double, _
        ByRef rDir1 As Double, ByRef rDir2 As Double)
    Dim denominator As Double
    denominator = MaxDouble(Abs(baseN) + Abs(baseMx) + Abs(baseMy), 1#) ^ 2
    If Abs(baseN) >= Abs(baseMx) And Abs(baseN) >= Abs(baseMy) Then
        rDir1 = (intMx * baseN - intN * baseMx) / denominator
        rDir2 = (intMy * baseN - intN * baseMy) / denominator
    ElseIf Abs(baseMx) >= Abs(baseMy) Then
        rDir1 = (intN * baseMx - intMx * baseN) / denominator
        rDir2 = (intMy * baseMx - intMx * baseMy) / denominator
    Else
        rDir1 = (intN * baseMy - intMy * baseN) / denominator
        rDir2 = (intMx * baseMy - intMy * baseMx) / denominator
    End If
End Sub

' Выполняет один Newton-шаг для прямой системы образования трещины. Якобиан
' строится численно, как в UltimateStrain, но предельным условием является
' нормативная предельная растягивающая деформация бетона eps_bt,ult.
Private Function ApplyCrcNewtonStep(ByVal section As CSectionModel, ByVal concreteMaterial As Object, _
        ByVal steelMaterial As Object, ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double, _
        ByRef eps0 As Double, ByRef kx As Double, ByRef ky As Double, _
        ByVal rDir1 As Double, ByVal rDir2 As Double, ByVal rLimit As Double) As Boolean
    Dim hEps As Double
    Dim hKx As Double
    Dim hKy As Double
    hEps = MaxDouble(0.0000001, Abs(eps0) * 0.0001)
    hKx = MaxDouble(0.000000001, Abs(kx) * 0.0001)
    hKy = MaxDouble(0.000000001, Abs(ky) * 0.0001)

    Dim j11 As Double, j21 As Double, j31 As Double
    Dim j12 As Double, j22 As Double, j32 As Double
    Dim j13 As Double, j23 As Double, j33 As Double
    BuildCrcJacobianColumn section, concreteMaterial, steelMaterial, targetN, targetMx, targetMy, _
        eps0 + hEps, kx, ky, hEps, rDir1, rDir2, rLimit, j11, j21, j31
    BuildCrcJacobianColumn section, concreteMaterial, steelMaterial, targetN, targetMx, targetMy, _
        eps0, kx + hKx, ky, hKx, rDir1, rDir2, rLimit, j12, j22, j32
    BuildCrcJacobianColumn section, concreteMaterial, steelMaterial, targetN, targetMx, targetMy, _
        eps0, kx, ky + hKy, hKy, rDir1, rDir2, rLimit, j13, j23, j33

    Dim system As CLinearSystem3x3
    Set system = New CLinearSystem3x3
    If Not system.Solve(j11, j12, j13, j21, j22, j23, j31, j32, j33, -rDir1, -rDir2, -rLimit) Then
        mStopReason = "NumericalFailure: матрица прямого поиска lambda_crc вырождена: " & system.StopReason
        Exit Function
    End If

    Dim dEps As Double
    Dim dKx As Double
    Dim dKy As Double
    dEps = system.X1
    dKx = system.X2
    dKy = system.X3
    LimitCrcIncrement dEps, dKx, dKy

    Dim baseNorm As Double
    baseNorm = CrcResidualNorm(rDir1, rDir2, rLimit)
    Dim alpha As Double
    alpha = mSolverDampingInitial
    If alpha <= 0# Or alpha > 1# Then alpha = 1#
    Do While alpha >= mSolverMinLineSearchAlpha
        Dim trialSolver As CSectionSolver
        Dim trialR1 As Double, trialR2 As Double, trialRLimit As Double, trialLambda As Double
        If EvaluateCrcStrainResidual(section, concreteMaterial, steelMaterial, targetN, targetMx, targetMy, _
                eps0 + alpha * dEps, kx + alpha * dKx, ky + alpha * dKy, _
                trialSolver, trialR1, trialR2, trialRLimit, trialLambda) Then
            If CrcResidualNorm(trialR1, trialR2, trialRLimit) < baseNorm Then
                eps0 = eps0 + alpha * dEps
                kx = kx + alpha * dKx
                ky = ky + alpha * dKy
                ApplyCrcNewtonStep = True
                Exit Function
            End If
        End If
        If Not mSolverLineSearchEnabled Then Exit Do
        alpha = alpha / 2#
    Loop

    mStopReason = "NumericalFailure: line search прямого поиска lambda_crc не улучшил невязку."
End Function

' Один столбец численного Якобиана для системы образования трещины.
Private Sub BuildCrcJacobianColumn(ByVal section As CSectionModel, ByVal concreteMaterial As Object, _
        ByVal steelMaterial As Object, ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double, _
        ByVal eps0 As Double, ByVal kx As Double, ByVal ky As Double, _
        ByVal stepValue As Double, ByVal rDir1 As Double, ByVal rDir2 As Double, ByVal rLimit As Double, _
        ByRef j1 As Double, ByRef j2 As Double, ByRef j3 As Double)
    Dim solver As CSectionSolver
    Dim p1 As Double, p2 As Double, pLimit As Double, pLambda As Double
    If Not EvaluateCrcStrainResidual(section, concreteMaterial, steelMaterial, targetN, targetMx, targetMy, _
            eps0, kx, ky, solver, p1, p2, pLimit, pLambda) Then Exit Sub
    j1 = (p1 - rDir1) / stepValue
    j2 = (p2 - rDir2) / stepValue
    j3 = (pLimit - rLimit) / stepValue
End Sub

' Норма смешивает условие направления и условие предельной деформации только
' для line search. Она не имеет самостоятельного нормативного смысла.
Private Function CrcResidualNorm(ByVal rDir1 As Double, ByVal rDir2 As Double, ByVal rLimit As Double) As Double
    CrcResidualNorm = Abs(rDir1) * 100# + Abs(rDir2) * 100# + Abs(rLimit) * 10#
End Function

' Ограничивает приращение прямой плоскости образования трещины теми же
' численными пределами, которыми пользователь управляет для CSectionSolver.
Private Sub LimitCrcIncrement(ByRef dEps As Double, ByRef dKx As Double, ByRef dKy As Double)
    Dim scaleFactor As Double
    scaleFactor = 1#
    If mSolverMaxDeltaEpsilon0 > 0# And Abs(dEps) > mSolverMaxDeltaEpsilon0 Then
        scaleFactor = MinDouble(scaleFactor, mSolverMaxDeltaEpsilon0 / Abs(dEps))
    End If
    If mSolverMaxDeltaKappa > 0# And Abs(dKx) > mSolverMaxDeltaKappa Then
        scaleFactor = MinDouble(scaleFactor, mSolverMaxDeltaKappa / Abs(dKx))
    End If
    If mSolverMaxDeltaKappa > 0# And Abs(dKy) > mSolverMaxDeltaKappa Then
        scaleFactor = MinDouble(scaleFactor, mSolverMaxDeltaKappa / Abs(dKy))
    End If
    dEps = dEps * scaleFactor
    dKx = dKx * scaleFactor
    dKy = dKy * scaleFactor
End Sub

Private Sub CompleteCrackFormula(ByVal solver As CSectionSolver, ByVal section As CSectionModel, _
        ByVal steelMaterial As Object, ByRef selected() As Boolean)
    mPhi1 = mPhi1Input
    mPhi2 = mPhi2Input
    If mPhi3 <= 0# Then mPhi3 = mPhi3Input
    mPsiS = mPsiSInput

    mAs = SelectedRebarArea(section, selected)
    mDsEquivalent = EquivalentDiameter(section, selected)
    mSigmaS = AverageSelectedSteelStress(solver, section, steelMaterial, selected)
    If mAs <= 0# Or mAbt <= 0# Or mDsEquivalent <= 0# Or mSigmaS <= 0# Then
        FinishNoCrack "Недостаточно растянутой арматуры или бетона для расчета раскрытия трещины."
        Exit Sub
    End If

    mLsRaw = 0.5 * (mAbt / mAs) * mDsEquivalent
    mLs = LimitedCrackSpacing(mLsRaw, mDsEquivalent)
End Sub

' Завершает расчет раскрытия с учетом режима psi_s.
' User заканчивает расчет после первой формулы с пользовательским psi_s.
' Auto сначала делает проверку с psi_s = 1; если a_crc <= a_crc,ult,
' уточнение не запускается.
' Только при непрохождении проверки Auto уточняет lambda_crc. В общей ветке
' при этом появляется BeforeMcrcState, а AfterMcrcState решается для той же
' lambda_crc уже без растянутого бетона. В центральной ветке lambda_crc берется
' из Ncrc/N, поэтому создается только AfterMcrcState.
Private Function FinalizeWithPsiMode(ByVal section As CSectionModel, ByVal materialProvider As CMaterialModelProvider, _
        ByVal currentSolver As CSectionSolver, ByRef selected() As Boolean, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double, _
        ByVal finalReason As String) As Boolean
    If StrComp(mPsiMode, "Auto", vbTextCompare) <> 0 Then
        mPsiS = mPsiSInput
        FinalizeWidth finalReason & " Psi_s принят из настройки SLS.Crack.PsiS."
        FinalizeWithPsiMode = mConverged
        Exit Function
    End If

    mPsiS = 1#
    FinalizeWidth finalReason & " Auto psi_s: первая проверка выполнена с psi_s=1."
    If Not mConverged Or Not mCrackFormed Then Exit Function

    If mAllowableCrackWidth > 0# And mCrackWidth <= mAllowableCrackWidth Then
        AppendDiagnostic "Auto psi_s: первая проверка с psi_s=1 проходит; lambda_crc не считался."
        FinalizeWithPsiMode = True
        Exit Function
    End If

    AppendDiagnostic "Auto psi_s: первая проверка с psi_s=1 не проходит; запускается расчет lambda_crc."
    mConverged = False

    ' При истинном центральном растяжении СП задает отдельный порог Ncrc.
    ' В этой ветке lambda_crc уже получена как Ncrc/N, поэтому ее нельзя
    ' перетирать общим поиском по eps_bt,ult для изгибных эпюр.
    If mCentralTensionBranch Then
        If mLambdaCrc <= 0# Then
            mStopReason = "InputError: для центрального растяжения lambda_crc через Ncrc не определена."
            Exit Function
        End If
    Else
        If mCrackInitiationSpec Is Nothing Then
            mStopReason = "InputError: для SLS.Crack.PsiMode=Auto требуется MaterialModel.CrackInitiation."
            Exit Function
        End If
        If Not CalculateLambdaCrc(section, materialProvider.ConcreteMaterialFromSpec(mCrackInitiationSpec), _
                materialProvider.SteelMaterialFromSpec(mCrackInitiationSpec), currentSolver, _
                targetN, targetMx, targetMy) Then Exit Function
    End If

    If mLambdaCrc > 1# Then
        FinishNoCrack "Трещина не образована: предельная растянутая деформация бетона не достигнута при lambda = 1."
        FinalizeWithPsiMode = True
        Exit Function
    End If

    Dim crcSolver As CSectionSolver
    Set crcSolver = SolveState(section, materialProvider.ConcreteMaterialFromSpec(mCrackedStateSpec), _
        materialProvider.SteelMaterialFromSpec(mCrackedStateSpec), currentSolver, _
        targetN * mLambdaCrc, targetMx * mLambdaCrc, targetMy * mLambdaCrc)
    If Not crcSolver.Converged Then
        mStopReason = "NumericalFailure: состояние lambda_crc после образования трещины не сошлось: " & crcSolver.StopReason
        Exit Function
    End If

    Set mAfterMcrcState = New CSectionStateResult
    mAfterMcrcState.InitializeFromSolver sstAfterMcrcState, cpCrackedNDS, _
        mCrackedStateSpec, crcSolver, False, "OK"

    mSigmaSCrc = AverageSelectedSteelStress(crcSolver, section, _
        materialProvider.SteelMaterialFromSpec(mCrackedStateSpec), selected)
    mPsiS = AutoPsi(mSigmaS, mSigmaSCrc)
    FinalizeWidth finalReason & " Psi_s уточнен автоматически после непрохождения первой проверки."
    FinalizeWithPsiMode = mConverged
End Function

' Завершает расчет по формуле СП 63 п. 8.2.15 для продолжительного раскрытия.
Private Sub FinalizeWidth(ByVal reason As String)
    If mSigmaS <= 0# Or mLs <= 0# Then
        FinishNoCrack "Напряжение растянутой арматуры или ls равны нулю."
        Exit Sub
    End If

    mCrackWidth = mPhi1 * mPhi2 * mPhi3 * mPsiS * (mSigmaS / mSteelEs) * mLs
    If mAllowableCrackWidth > 0# Then mUtilization = mCrackWidth / mAllowableCrackWidth
    mCrackFormed = True
    mConverged = True
    mStopReason = reason
    AppendDiagnostic "a_crc=" & FormatNumberInvariant(mCrackWidth) & _
        "; phi1=" & FormatNumberInvariant(mPhi1) & _
        "; phi2=" & FormatNumberInvariant(mPhi2) & _
        "; phi3=" & FormatNumberInvariant(mPhi3) & _
        "; psi_s=" & FormatNumberInvariant(mPsiS) & _
        "; sigma_s=" & FormatNumberInvariant(mSigmaS) & _
        "; ls=" & FormatNumberInvariant(mLs)
End Sub

' Помечает расчет как корректно завершенный без раскрытой трещины.
' Это не численная ошибка: заданное состояние либо не растягивает арматуру,
' либо еще не достигло момента образования трещины.
Private Sub FinishNoCrack(ByVal reason As String)
    mConverged = True
    mCrackFormed = False
    mCrackWidth = 0#
    mUtilization = 0#
    mStopReason = reason
    AppendDiagnostic reason
End Sub

' Создает и настраивает внутренний CSectionSolver. Сам решатель не изменяется:
' класс трещин только выбирает режим бетона и передает готовые N, Mx, My.
Private Function SolveState(ByVal section As CSectionModel, ByVal concreteMaterial As Object, _
        ByVal steelMaterial As Object, ByVal startSolver As CSectionSolver, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double) As CSectionSolver
    Dim solver As CSectionSolver
    Set solver = New CSectionSolver
    ConfigureSolver solver
    If Not startSolver Is Nothing Then
        If startSolver.Converged Then
            solver.SetInitialLoadState startSolver.Epsilon0, startSolver.KappaX, startSolver.KappaY, _
                startSolver.Nint, startSolver.Mxint, startSolver.Myint
        Else
            solver.SetInitialState startSolver.Epsilon0, startSolver.KappaX, startSolver.KappaY
        End If
    End If
    solver.Solve section, concreteMaterial, steelMaterial, targetN, targetMx, targetMy
    Set SolveState = solver
End Function

' Передает внутренние настройки в CSectionSolver без чтения Config внутри solver-а.
Private Sub ConfigureSolver(ByVal solver As CSectionSolver)
    solver.LoadSteps = mSolverLoadSteps
    solver.MaxIterations = mSolverMaxIterations
    solver.ToleranceN = mSolverToleranceN
    solver.ToleranceMx = mSolverToleranceMx
    solver.ToleranceMy = mSolverToleranceMy
    solver.LineSearchEnabled = mSolverLineSearchEnabled
    solver.DampingInitial = mSolverDampingInitial
    solver.MinLineSearchAlpha = mSolverMinLineSearchAlpha
    solver.SolverMethod = mSolverMethod
    solver.SecantMaxRestarts = mSolverSecantMaxRestarts
    solver.SecantMinStepNorm = mSolverSecantMinStepNorm
    solver.MaxDeltaEpsilon0 = mSolverMaxDeltaEpsilon0
    solver.MaxDeltaKappa = mSolverMaxDeltaKappa
    solver.DiagnosticsEnabled = mDiagnosticsEnabled
End Sub

' Центральное растяжение для трещин определяется по моментному вектору
' относительно центра тяжести бетонного сечения. В этот вектор уже должен
' входить эксцентриситет точки приложения нагрузки: Mx_cg = Mx_user + N*dy,
' My_cg = My_user + N*dx. Поэтому случай N при Mx/My = 0, но с ненулевым
' Load.ReferenceOffset, не попадает в центральную ветку.
'
' Кривизны solver-а здесь сознательно не используются. Если при таком LC
' появляется заметная кривизна, это уже ошибка в переносе нагрузок, сетке или
' тестовом сценарии; она должна ловиться регрессионными проверками, а не
' менять расчетную ветку в рабочей книге.
Private Function IsCentralTensionState(ByVal targetN As Double, _
        Optional ByVal centroidMxForCentralCheck As Variant, Optional ByVal centroidMyForCentralCheck As Variant) As Boolean
    If targetN <= NEAR_ZERO_FORCE Then Exit Function

    If IsMissing(centroidMxForCentralCheck) Or IsMissing(centroidMyForCentralCheck) Then
        AppendDiagnostic "Центральная ветка не проверялась: не переданы моменты относительно центра тяжести бетонного сечения."
        Exit Function
    End If

    If Abs(CDbl(centroidMxForCentralCheck)) > MaxDouble(mSolverToleranceMx, 1#) Or _
            Abs(CDbl(centroidMyForCentralCheck)) > MaxDouble(mSolverToleranceMy, 1#) Then Exit Function

    IsCentralTensionState = True
End Function

' Выбирает все растянутые стержни для центрального растяжения.
Private Sub SelectAllTensionRebars(ByVal solver As CSectionSolver, ByVal section As CSectionModel, ByRef selected() As Boolean)
    Dim i As Long
    For i = 1 To section.RebarCount
        If RebarStrain(solver, section, i) > 0# Then
            selected(i) = True
            mTensionRebarCount = mTensionRebarCount + 1
        End If
    Next i
End Sub

' Находит крайнюю растянутую арматуру по нормали к нейтральной линии.
Private Function FindMostTensionRebarProjection(ByVal solver As CSectionSolver, ByVal section As CSectionModel, _
        ByVal nx As Double, ByVal ny As Double, ByRef sBarMax As Double, ByRef rebarIndex As Long) As Boolean
    Dim i As Long
    For i = 1 To section.RebarCount
        If RebarStrain(solver, section, i) > 0# Then
            Dim s As Double
            s = Projection(section.RebarX(i), section.RebarY(i), nx, ny)
            If Not FindMostTensionRebarProjection Or s > sBarMax Then sBarMax = s
            If s >= sBarMax - NEAR_ZERO_STRAIN Then rebarIndex = i
            FindMostTensionRebarProjection = True
        End If
    Next i
End Function

' Выбирает пользовательский смысл расстояния a_s. В обычном режиме берется
' локальный выход луча из выбранного стержня к ближайшему бетонному контуру.
' Режим GlobalExtreme оставлен для проверки старой глобальной трактовки, когда
' расстояние считается до дальней опорной линии всего сечения.
Private Function CoverDistanceForRebar(ByVal section As CSectionModel, ByVal rebarIndex As Long, _
        ByVal nx As Double, ByVal ny As Double, ByVal st As Double, ByVal sBarMax As Double) As Double
    If StrComp(mCoverMode, COVER_MODE_GLOBAL_EXTREME, vbTextCompare) = 0 Then
        CoverDistanceForRebar = st - sBarMax
        Exit Function
    End If

    CoverDistanceForRebar = LocalConcreteCoverForRebar(section, rebarIndex, nx, ny)
    If CoverDistanceForRebar < 0# Then CoverDistanceForRebar = st - sBarMax
End Function

' Для a_s нужна не глобальная крайняя проекция всего сечения, а расстояние от
' оси выбранного растянутого стержня до ближайшего бетонного контура по нормали
' к нейтральной линии. Это особенно важно для Г-сечений и импортированных
' повернутых сеток, где глобальная опорная линия может лежать на другом выступе.
Private Function LocalConcreteCoverForRebar(ByVal section As CSectionModel, ByVal rebarIndex As Long, _
        ByVal nx As Double, ByVal ny As Double) As Double
    If rebarIndex <= 0 Then
        LocalConcreteCoverForRebar = -1#
        Exit Function
    End If

    Dim props As CSectionPropertiesCalculator
    Set props = New CSectionPropertiesCalculator
    LocalConcreteCoverForRebar = props.ConcreteCoverFromPointAlongDirection(section, _
        section.RebarX(rebarIndex), section.RebarY(rebarIndex), nx, ny)
End Function

' Определяет нижнюю границу расчетной зоны в координате s.
' FullTension берет всю фактически растянутую часть до нейтральной линии.
' Effective задает нормативную глубину h_bt = min(max(x_t, 2a), 0.5h).
' Если фактическая растянутая зона x_t меньше 2a, полоса может геометрически
' зайти за нейтральную линию. Для Abt это намеренно учитывается: бетонные КЭ
' выбираются по попаданию центра в полосу h_bt, а не по знаку текущей деформации.
' Арматура As остается только растянутой, потому что sigma_s в формуле раскрытия
' относится к работающей растянутой арматуре между трещинами.
Private Function TensionZoneLowerBound(ByVal st As Double, ByVal sc As Double, ByVal neutralS As Double) As Double
    Dim lower As Double
    If StrComp(mTensionZoneMode, "FullTension", vbTextCompare) = 0 Then
        lower = MaxDouble(neutralS, sc)
    Else
        Dim effectiveDepth As Double
        effectiveDepth = MinDouble(MaxDouble(mTensionDepth, 2# * mCoverA), 0.5 * mSectionDepthH)
        If effectiveDepth <= NEAR_ZERO_STRAIN Then effectiveDepth = MinDouble(mTensionDepth, 0.5 * mSectionDepthH)
        lower = st - effectiveDepth
        lower = MaxDouble(lower, sc)
    End If
    TensionZoneLowerBound = lower
End Function

' Суммирует площадь бетонных КЭ, чьи центры попали в расчетную растянутую
' зону. Частичные площади намеренно не считаются: это принятое дискретное
' приближение текущей НДМ-модели.
Private Function ConcreteAreaInZone(ByVal solver As CSectionSolver, ByVal section As CSectionModel, _
        ByVal nx As Double, ByVal ny As Double, ByVal lowerS As Double, _
        ByVal requireTensionStrain As Boolean) As Double
    Dim i As Long
    For i = 1 To section.ConcreteCount
        If Not requireTensionStrain Or ConcreteStrain(solver, section, i) > 0# Then
            If Projection(section.ConcreteX(i), section.ConcreteY(i), nx, ny) >= lowerS - NEAR_ZERO_STRAIN Then
                ConcreteAreaInZone = ConcreteAreaInZone + section.ConcreteArea(i)
            End If
        End If
    Next i
End Function

' Формирует набор стержней As по той же полосе, которая принята для Abt,
' но с обязательным физическим фильтром по растяжению стержня. Поэтому при
' Effective и срабатывании ограничения 2a бетонная полоса может зайти за
' нейтральную линию, однако сжатые стержни в As, sigma_s и d_s,eq не попадают.
Private Sub SelectRebarsInZone(ByVal solver As CSectionSolver, ByVal section As CSectionModel, _
        ByVal nx As Double, ByVal ny As Double, ByVal lowerS As Double, ByRef selected() As Boolean)
    Dim i As Long
    For i = 1 To section.RebarCount
        If RebarStrain(solver, section, i) > 0# Then
            If Projection(section.RebarX(i), section.RebarY(i), nx, ny) >= lowerS - NEAR_ZERO_STRAIN Then
                selected(i) = True
                mTensionRebarCount = mTensionRebarCount + 1
            End If
        End If
    Next i
End Sub

' Внешняя растянутая поверхность берется через общий геометрический калькулятор.
' Для сгенерированной сетки используются реальные размеры бетонных КЭ, а для
' импортированных Region без габаритов сохраняется общий fallback: сначала
' эквивалентный прямоугольник по A/I, затем квадрат той же площади. Та же
' логика используется устойчивостью.
Private Sub ProjectionBoundsByConcreteContour(ByVal section As CSectionModel, _
        ByVal nx As Double, ByVal ny As Double, ByRef st As Double, ByRef sc As Double)
    Dim props As CSectionPropertiesCalculator
    Set props = New CSectionPropertiesCalculator
    props.CalculateProjection section, nx, ny, False, sc, st
End Sub

' Считает средневзвешенное напряжение выбранной арматуры. В формулу раскрытия
' идет одно глобальное sigma_s, а не локальные значения по отдельным зонам.
Private Function AverageSelectedSteelStress(ByVal solver As CSectionSolver, ByVal section As CSectionModel, _
        ByVal steelMaterial As Object, ByRef selected() As Boolean) As Double
    Dim weightedStress As Double
    Dim areaSum As Double
    Dim i As Long
    For i = 1 To section.RebarCount
        If selected(i) Then
            Dim strain As Double
            Dim stress As Double
            strain = RebarStrain(solver, section, i)
            stress = steelMaterial.GetStress(strain)
            If stress < 0# Then stress = 0#
            weightedStress = weightedStress + section.RebarArea(i) * stress
            areaSum = areaSum + section.RebarArea(i)
            TrackMaxSteel section, i, strain, stress
        End If
    Next i
    If areaSum > 0# Then AverageSelectedSteelStress = weightedStress / areaSum
End Function

' Считает площадь выбранной арматуры.
Private Function SelectedRebarArea(ByVal section As CSectionModel, ByRef selected() As Boolean) As Double
    Dim i As Long
    For i = 1 To section.RebarCount
        If selected(i) Then SelectedRebarArea = SelectedRebarArea + section.RebarArea(i)
    Next i
End Function

' Для разных диаметров используется эквивалентный диаметр Eurocode:
' ds,eq = sum(n_i*d_i^2) / sum(n_i*d_i). В дискретной модели каждый стержень
' считается отдельным i, поэтому n_i фактически равен количеству совпадающих d.
Private Function EquivalentDiameter(ByVal section As CSectionModel, ByRef selected() As Boolean) As Double
    Dim numerator As Double
    Dim denominator As Double
    Dim i As Long
    For i = 1 To section.RebarCount
        If selected(i) Then
            Dim d As Double
            d = section.RebarDiameter(i)
            If d <= 0# Then d = Sqr(4# * section.RebarArea(i) / GEOM_PI)
            numerator = numerator + d * d
            denominator = denominator + d
        End If
    Next i
    If denominator > 0# Then EquivalentDiameter = numerator / denominator
End Function

' Применяет ограничения СП 63 к расстоянию между трещинами.
Private Function LimitedCrackSpacing(ByVal rawLs As Double, ByVal ds As Double) As Double
    Dim lower As Double
    Dim upper As Double
    lower = MaxDouble(10# * ds, 100#)
    upper = MinDouble(40# * ds, 400#)
    If lower > upper Then
        AppendDiagnostic "Предупреждение: ограничения ls конфликтуют; использована верхняя граница " & FormatNumberInvariant(upper) & " мм."
        LimitedCrackSpacing = upper
    Else
        LimitedCrackSpacing = MinDouble(MaxDouble(rawLs, lower), upper)
    End If
End Function

' Формула psi_s = 1 - 0.8*sigma_s,crc/sigma_s. Ограничение 0...1 является
' численной защитой от отрицательного psi_s при почти равных напряжениях или
' случайных малых несоответствиях solve-ов.
Private Function AutoPsi(ByVal sigmaS As Double, ByVal sigmaSCrc As Double) As Double
    If sigmaS <= 0# Then
        AutoPsi = 1#
    Else
        AutoPsi = 1# - 0.8 * sigmaSCrc / sigmaS
        If AutoPsi < 0# Then AutoPsi = 0#
        If AutoPsi > 1# Then AutoPsi = 1#
    End If
End Function

' Возвращает усилие образования трещины при истинном центральном растяжении:
' Ncrc = Ared * Rbt,ser, где Ared включает бетон и арматуру через Es/Eb.
Private Function CentralTensionCrackForce(ByVal section As CSectionModel) As Double
    Dim aRed As Double
    aRed = TotalConcreteArea(section)
    If mConcreteEb > 0# Then aRed = aRed + TotalRebarArea(section) * mSteelEs / mConcreteEb
    CentralTensionCrackForce = aRed * mConcreteRbtSer
End Function

' Суммарная площадь бетона модели.
Private Function TotalConcreteArea(ByVal section As CSectionModel) As Double
    Dim i As Long
    For i = 1 To section.ConcreteCount
        TotalConcreteArea = TotalConcreteArea + section.ConcreteArea(i)
    Next i
End Function

' Суммарная площадь арматуры модели.
Private Function TotalRebarArea(ByVal section As CSectionModel) As Double
    Dim i As Long
    For i = 1 To section.RebarCount
        TotalRebarArea = TotalRebarArea + section.RebarArea(i)
    Next i
End Function

' Максимальная растягивающая деформация бетонных КЭ по сохраненной плоскости.
Private Function MaxConcreteStrain(ByVal solver As CSectionSolver, ByVal section As CSectionModel) As Double
    Dim i As Long
    For i = 1 To section.ConcreteCount
        MaxConcreteStrain = MaxDouble(MaxConcreteStrain, ConcreteStrain(solver, section, i))
    Next i
End Function

' Деформация бетонного КЭ.
Private Function ConcreteStrain(ByVal solver As CSectionSolver, ByVal section As CSectionModel, ByVal index As Long) As Double
    ConcreteStrain = solver.Epsilon0 + solver.KappaX * section.ConcreteY(index) + solver.KappaY * section.ConcreteX(index)
End Function

' Деформация стержня арматуры.
Private Function RebarStrain(ByVal solver As CSectionSolver, ByVal section As CSectionModel, ByVal index As Long) As Double
    RebarStrain = solver.Epsilon0 + solver.KappaX * section.RebarY(index) + solver.KappaY * section.RebarX(index)
End Function

' Проекция точки на нормаль к нейтральной линии.
Private Function Projection(ByVal xCoord As Double, ByVal yCoord As Double, ByVal nx As Double, ByVal ny As Double) As Double
    Projection = nx * xCoord + ny * yCoord
End Function

' Обновляет максимум деформации и напряжения арматуры для диагностики.
Private Sub TrackMaxSteel(ByVal section As CSectionModel, ByVal index As Long, ByVal strain As Double, ByVal stress As Double)
    AppendDiagnostic "bar=" & section.RebarID(index) & _
        "; strain=" & FormatNumberInvariant(strain) & _
        "; stress=" & FormatNumberInvariant(stress)
    If strain > mMaxSteelStrain Then
        mMaxSteelStrain = strain
        mMaxSteelStress = stress
    End If
End Sub

' Определяет минимальную и максимальную деформацию бетонных КЭ для пробной
' плоскости Mcrc. Эти две величины нужны, чтобы отличить двузначную эпюру от
' однозначно растянутой и выбрать нормативный eps_bt,ult по СП 63 п. 8.1.30.
Private Sub ConcreteStrainBounds(ByVal solver As CSectionSolver, ByVal section As CSectionModel, _
        ByRef minStrain As Double, ByRef maxStrain As Double)
    Dim i As Long
    For i = 1 To section.ConcreteCount
        Dim strain As Double
        strain = ConcreteStrain(solver, section, i)
        If i = 1 Then
            minStrain = strain
            maxStrain = strain
        Else
            If strain < minStrain Then minStrain = strain
            If strain > maxStrain Then maxStrain = strain
        End If
    Next i
End Sub

' Возвращает eps_bt,ult для критерия образования трещины по СП 63 п. 8.1.30.
' При двузначной эпюре используется eps_bt2. При однозначном растяжении с
' ненулевой кривизной применяется формула (8.54):
' eps_bt,ult = eps_bt2 - (eps_bt2 - eps_bt0) * eps1 / eps2.
' Истинное центральное растяжение сюда не попадает: для него выше работает
' отдельная ветка Ncrc по п. 8.2.13.
Private Function ConcreteTensionUltimateStrain(ByVal minStrain As Double, ByVal maxStrain As Double) As Double
    If mConcreteEbt2 <= 0# Then Exit Function

    If maxStrain <= NEAR_ZERO_STRAIN Then
        ConcreteTensionUltimateStrain = mConcreteEbt2
        Exit Function
    End If

    If minStrain <= NEAR_ZERO_STRAIN Then
        ConcreteTensionUltimateStrain = mConcreteEbt2
        Exit Function
    End If

    Dim ratio As Double
    ratio = minStrain / maxStrain
    If ratio < 0# Then ratio = 0#
    If ratio > 1# Then ratio = 1#
    ConcreteTensionUltimateStrain = mConcreteEbt2 - (mConcreteEbt2 - mConcreteEbt0) * ratio
End Function

' Возвращает phi3 для текущего LC. В Auto используется принятое правило СП:
' при растягивающей продольной силе принимается 1.2, иначе 1.0. В User
' программа не выводит характер нагрузки из НДС, а берет введенное значение.
Private Function ResolvePhi3(ByVal targetN As Double) As Double
    If StrComp(mPhi3Mode, "User", vbTextCompare) = 0 Then
        ResolvePhi3 = mPhi3Input
    ElseIf targetN > NEAR_ZERO_FORCE Then
        ResolvePhi3 = PHI3_TENSION
    Else
        ResolvePhi3 = PHI3_BENDING
    End If
End Function

' Нормализует текстовые режимы и проверяет численные параметры до запуска
' внутренних solve, чтобы ошибки Config были понятны пользователю.
Private Sub NormalizeSettings()
    Select Case LCase$(Trim$(mPsiMode))
        Case "user", "manual", "fixed1", "unity"
            mPsiMode = "User"
        Case "auto", "refined"
            mPsiMode = "Auto"
        Case Else
            Err.Raise vbObjectError + 3707, "CCrackWidthCalculator", "SLS.Crack.PsiMode должен быть User или Auto."
    End Select

    Select Case LCase$(Trim$(mPhi3Mode))
        Case "user", "manual"
            mPhi3Mode = "User"
        Case "auto"
            mPhi3Mode = "Auto"
        Case Else
            Err.Raise vbObjectError + 3709, "CCrackWidthCalculator", "SLS.Crack.Phi3Mode должен быть User или Auto."
    End Select

    Select Case LCase$(Trim$(mTensionZoneMode))
        Case "effective": mTensionZoneMode = "Effective"
        Case "fulltension": mTensionZoneMode = "FullTension"
        Case Else: Err.Raise vbObjectError + 3708, "CCrackWidthCalculator", "SLS.Crack.TensionZoneMode должен быть Effective или FullTension."
    End Select

    Select Case LCase$(Trim$(mCoverMode))
        Case "nearestcontour", "nearest", "local"
            mCoverMode = COVER_MODE_NEAREST_CONTOUR
        Case "globalextreme", "global", "extreme"
            mCoverMode = COVER_MODE_GLOBAL_EXTREME
        Case Else
            Err.Raise vbObjectError + 3721, "CCrackWidthCalculator", "SLS.Crack.CoverDistanceMode должен быть NearestContour или GlobalExtreme."
    End Select

    If mAllowableCrackWidth <= 0# Then Err.Raise vbObjectError + 3710, "CCrackWidthCalculator", "SLS.Crack.Allowable должен быть положительным."
    If mPhi1Input <= 0# Then Err.Raise vbObjectError + 3717, "CCrackWidthCalculator", "SLS.Crack.Phi1 должен быть положительным."
    If mPhi2Input <= 0# Then Err.Raise vbObjectError + 3718, "CCrackWidthCalculator", "SLS.Crack.Phi2 должен быть положительным."
    If mPhi3Input <= 0# Then Err.Raise vbObjectError + 3719, "CCrackWidthCalculator", "SLS.Crack.Phi3 должен быть положительным."
    If mPsiSInput <= 0# Then Err.Raise vbObjectError + 3720, "CCrackWidthCalculator", "SLS.Crack.PsiS должен быть положительным."
    If mConcreteRbtSer <= 0# Then Err.Raise vbObjectError + 3711, "CCrackWidthCalculator", "Concrete.Rbt.SLS должен быть положительным для расчета трещин."
    If mConcreteEb <= 0# Then Err.Raise vbObjectError + 3712, "CCrackWidthCalculator", "Concrete.Eb должен быть положительным для расчета трещин."
    If mSteelEs <= 0# Then Err.Raise vbObjectError + 3713, "CCrackWidthCalculator", "Steel.Es должен быть положительным для расчета трещин."
    If StrComp(mPsiMode, "Auto", vbTextCompare) = 0 And Not mCrackInitiationSpec Is Nothing Then
        If mConcreteEbt0 <= 0# Then Err.Raise vbObjectError + 3714, "CCrackWidthCalculator", "Concrete.Ebt0 должен быть положительным для расчета Mcrc."
        If mConcreteEbt2 <= 0# Then Err.Raise vbObjectError + 3715, "CCrackWidthCalculator", "Concrete.Ebt2 должен быть положительным для расчета Mcrc."
        If mConcreteEbt0 > mConcreteEbt2 Then Err.Raise vbObjectError + 3716, "CCrackWidthCalculator", "Concrete.Ebt0 не должен быть больше Concrete.Ebt2 для формулы (8.54)."
    End If
End Sub

' Проверяет расчетный контракт класса.
Private Sub ValidateInputs(ByVal currentSolver As CSectionSolver, ByVal section As CSectionModel, _
        ByVal materialProvider As CMaterialModelProvider)
    If currentSolver Is Nothing Then Err.Raise vbObjectError + 3700, "CCrackWidthCalculator", "Результат решателя сечения не передан."
    If Not currentSolver.Converged Then Err.Raise vbObjectError + 3701, "CCrackWidthCalculator", "Результат решателя сечения не сошелся."
    If section Is Nothing Then Err.Raise vbObjectError + 3702, "CCrackWidthCalculator", "Модель расчетного сечения не передана."
    If section.ConcreteCount <= 0 Then Err.Raise vbObjectError + 3703, "CCrackWidthCalculator", "В модели расчетного сечения нет бетонных элементов."
    If section.RebarCount <= 0 Then Err.Raise vbObjectError + 3704, "CCrackWidthCalculator", "В модели расчетного сечения нет арматурных элементов."
    If materialProvider Is Nothing Then Err.Raise vbObjectError + 3705, "CCrackWidthCalculator", "Material provider не передан."
End Sub

' Сбрасывает результаты предыдущего вызова.
Private Sub ClearResult()
    mConverged = False
    mStopReason = vbNullString
    mDiagnosticLog = vbNullString
    mCrackWidth = 0#
    mUtilization = 0#
    mCrackFormed = False
    mTensionRebarCount = 0
    mMaxSteelStrain = 0#
    mMaxSteelStress = 0#
    mSigmaS = 0#
    mSigmaSCrc = 0#
    mAs = 0#
    mAbt = 0#
    mDsEquivalent = 0#
    mLsRaw = 0#
    mLs = 0#
    mPsiS = mPsiSInput
    mPhi1 = mPhi1Input
    mPhi2 = mPhi2Input
    mPhi3 = 0#
    mLambdaCrc = 0#
    mNcrc = 0#
    Set mBeforeMcrcState = Nothing
    Set mAfterMcrcState = Nothing
    Set mCrackInitiationSpec = Nothing
    Set mCrackedStateSpec = Nothing
    mSectionDepthH = 0#
    mCoverA = 0#
    mTensionDepth = 0#
    mEffectiveZoneDepth = 0#
    mCentralTensionBranch = False
End Sub

' Добавляет строку в диагностический журнал.
Private Sub AppendDiagnostic(ByVal text As String)
    If Not mDiagnosticsEnabled Then Exit Sub
    mDiagnosticLog = mDiagnosticLog & text & vbCrLf
End Sub

' Форматирует числа диагностического журнала через точку.
Private Function FormatNumberInvariant(ByVal value As Double) As String
    FormatNumberInvariant = Replace$(Format$(value, "0.############"), ",", ".")
End Function

Private Function MinDouble(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then MinDouble = a Else MinDouble = b
End Function

Private Function MaxDouble(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then MaxDouble = a Else MaxDouble = b
End Function

' --------------------------------------------------------------------------
' Свойства результатов
' --------------------------------------------------------------------------
' Writer-ы и тесты читают эти значения после Calculate. Свойства ничего не
' пересчитывают и не обращаются к solver-у повторно.

Public Property Get Converged() As Boolean
    Converged = mConverged
End Property

Public Property Get StopReason() As String
    StopReason = mStopReason
End Property

Public Property Get CrackFormed() As Boolean
    CrackFormed = mCrackFormed
End Property

Public Property Get TensionRebarCount() As Long
    TensionRebarCount = mTensionRebarCount
End Property

Public Property Get MaxSteelStrain() As Double
    MaxSteelStrain = mMaxSteelStrain
End Property

Public Property Get MaxSteelStress() As Double
    MaxSteelStress = mMaxSteelStress
End Property

Public Property Get CrackWidth() As Double
    CrackWidth = mCrackWidth
End Property

Public Property Get AllowableCrackWidth() As Double
    AllowableCrackWidth = mAllowableCrackWidth
End Property

Public Property Get Utilization() As Double
    Utilization = mUtilization
End Property

Public Property Get DiagnosticLog() As String
    DiagnosticLog = mDiagnosticLog
End Property

Public Property Get PsiMode() As String
    PsiMode = mPsiMode
End Property

Public Property Get TensionZoneMode() As String
    TensionZoneMode = mTensionZoneMode
End Property

Public Property Get CoverMode() As String
    CoverMode = mCoverMode
End Property

Public Property Get Phi1() As Double
    Phi1 = mPhi1
End Property

Public Property Get Phi2() As Double
    Phi2 = mPhi2
End Property

Public Property Get Phi3() As Double
    Phi3 = mPhi3
End Property

Public Property Get PsiS() As Double
    PsiS = mPsiS
End Property

Public Property Get SigmaS() As Double
    SigmaS = mSigmaS
End Property

Public Property Get SigmaSCrc() As Double
    SigmaSCrc = mSigmaSCrc
End Property

Public Property Get AsTension() As Double
    AsTension = mAs
End Property

Public Property Get Abt() As Double
    Abt = mAbt
End Property

Public Property Get DsEquivalent() As Double
    DsEquivalent = mDsEquivalent
End Property

Public Property Get CrackSpacingRaw() As Double
    CrackSpacingRaw = mLsRaw
End Property

Public Property Get CrackSpacing() As Double
    CrackSpacing = mLs
End Property

Public Property Get LambdaCrc() As Double
    LambdaCrc = mLambdaCrc
End Property

Public Property Get Ncrc() As Double
    Ncrc = mNcrc
End Property

Public Property Get SectionDepthH() As Double
    SectionDepthH = mSectionDepthH
End Property

Public Property Get CoverA() As Double
    CoverA = mCoverA
End Property

Public Property Get TensionDepth() As Double
    TensionDepth = mTensionDepth
End Property

Public Property Get EffectiveZoneDepth() As Double
    EffectiveZoneDepth = mEffectiveZoneDepth
End Property

Public Property Get CentralTensionBranch() As Boolean
    CentralTensionBranch = mCentralTensionBranch
End Property

Public Property Get BeforeMcrcState() As CSectionStateResult
    Set BeforeMcrcState = mBeforeMcrcState
End Property

Public Property Get AfterMcrcState() As CSectionStateResult
    Set AfterMcrcState = mAfterMcrcState
End Property

Public Property Let AllowableCrackWidth(ByVal value As Double)
    mAllowableCrackWidth = value
End Property

Public Property Let PsiMode(ByVal value As String)
    mPsiMode = value
End Property

Public Property Let Phi1(ByVal value As Double)
    mPhi1Input = value
End Property

Public Property Let Phi2(ByVal value As Double)
    mPhi2Input = value
End Property

Public Property Let Phi3Mode(ByVal value As String)
    mPhi3Mode = value
End Property

Public Property Let Phi3(ByVal value As Double)
    mPhi3Input = value
End Property

Public Property Let PsiS(ByVal value As Double)
    mPsiSInput = value
End Property

Public Property Let TensionZoneMode(ByVal value As String)
    mTensionZoneMode = value
End Property

Public Property Let CoverMode(ByVal value As String)
    mCoverMode = value
End Property

Public Property Let SolverMethod(ByVal value As String)
    mSolverMethod = value
End Property

Public Property Let SolverLoadSteps(ByVal value As Long)
    mSolverLoadSteps = value
End Property

Public Property Let SolverMaxIterations(ByVal value As Long)
    mSolverMaxIterations = value
End Property

Public Property Let SolverToleranceN(ByVal value As Double)
    mSolverToleranceN = value
End Property

Public Property Let SolverToleranceMx(ByVal value As Double)
    mSolverToleranceMx = value
End Property

Public Property Let SolverToleranceMy(ByVal value As Double)
    mSolverToleranceMy = value
End Property

Public Property Let SolverLineSearchEnabled(ByVal value As Boolean)
    mSolverLineSearchEnabled = value
End Property

Public Property Let SolverDampingInitial(ByVal value As Double)
    mSolverDampingInitial = value
End Property

Public Property Let SolverMinLineSearchAlpha(ByVal value As Double)
    mSolverMinLineSearchAlpha = value
End Property

Public Property Let SolverMaxDeltaEpsilon0(ByVal value As Double)
    mSolverMaxDeltaEpsilon0 = value
End Property

Public Property Let SolverMaxDeltaKappa(ByVal value As Double)
    mSolverMaxDeltaKappa = value
End Property

Public Property Let SolverSecantMaxRestarts(ByVal value As Long)
    mSolverSecantMaxRestarts = value
End Property

Public Property Let SolverSecantMinStepNorm(ByVal value As Double)
    mSolverSecantMinStepNorm = value
End Property

Public Property Let DiagnosticsEnabled(ByVal value As Boolean)
    mDiagnosticsEnabled = value
End Property
