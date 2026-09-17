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
' Перед расчетом ширины всегда проверяется факт образования нормальной трещины:
' для центрального растяжения через Ncrc, для остальных случаев через точку
' достижения eps_bt,ult по модели CrackInitiation. Batch отдельно считает
' текущий CrackedState полного LC для продольных трещин и, если нормальная
' трещина образуется, передает это же состояние сюда для формулы раскрытия.
' Режим psi_s влияет только на коэффициент в этой формуле: User берет
' введенное значение, Auto при необходимости уточняет psi_s по уже найденному
' состоянию AfterMcrcState.

Private Const PHI1_LONG_TERM As Double = 1.4 ' Коэффициент длительного действия нагрузки для a_crc1 по СП 63.
Private Const PHI2_DEFAULT As Double = 0.5 ' Типовой default для обычной арматуры периодического профиля.
Private Const PHI3_TENSION As Double = 1.2 ' Коэффициент для центрального/сквозного растяжения.
Private Const PHI3_BENDING As Double = 1# ' Коэффициент для изгиба и внецентренного сжатия.
Private Const CRACK_DIRECTION_TOLERANCE As Double = 0.001 ' Допуск сохранения направления моментного вектора в прямом поиске lambda_crc.
Private Const CRACK_EQUILIBRIUM_TOLERANCE As Double = 1# ' Допуск равновесия N в долях пользовательского solver tolerance.
Private Const CRACK_STRAIN_RATIO_TOLERANCE As Double = 0.000001 ' Допуск попадания в eps_bt,ult как в безразмерное условие.
Private Const CRACK_LAMBDA_BRACKET_LIMIT As Double = 1024# ' Верхняя граница служебного поиска lambda_crc, если прямой Newton не сошелся.
Private Const CRACK_LAMBDA_BISECTION_TOLERANCE As Double = 0.000001 ' Относительная точность fallback-бисекции lambda_crc.
Private Const NEAR_ZERO_STRAIN As Double = 0.000000000001 ' Численный ноль для деформаций и кривизн.
Private Const NEAR_ZERO_FORCE As Double = 0.000001 ' Численный ноль для отличия центрального растяжения от нулевой нагрузки.
Private Const LARGE_LAMBDA As Double = 1E+30 ' Диагностическая отметка: трещина не достигнута при lambda = 1.
Private Const COVER_MODE_NEAREST_CONTOUR As String = "NearestContour" ' a_s до ближайшего выхода бетонного контура по нормали.
Private Const COVER_MODE_GLOBAL_EXTREME As String = "GlobalExtreme" ' a_s до дальней растянутой опорной линии всего сечения.
Private Const CRACK_PATH_AUTO As String = "Auto" ' Последовательно пробует lambda*Mxy, lambda*N и lambda*NMxy.
Private Const CRACK_PATH_MXY As String = "LambdaMxy" ' При поиске трещинообразования масштабируется только моментный вектор Mx/My.
Private Const CRACK_PATH_N As String = "LambdaN" ' При поиске трещинообразования масштабируется только продольная сила N.
Private Const CRACK_PATH_NMXY As String = "LambdaNMxy" ' При поиске трещинообразования масштабируется весь вектор N/Mx/My.
Private Const STRATEGY_AUTO As String = "Auto" ' Сначала UltimateStrain, при численной неудаче LoadMultiplier.
Private Const STRATEGY_ULTIMATE_STRAIN As String = "UltimateStrain" ' Прямой поиск плоскости по eps_bt,ult без одномерной скобки.
Private Const STRATEGY_LOAD_MULTIPLIER As String = "LoadMultiplier" ' Одномерный поиск lambda через серию обычных solve.

Private mConverged As Boolean ' True, если расчет трещины завершился без ошибки входных данных или несходимости.
Private mStopReason As String ' Короткое пояснение результата: расчет выполнен, трещина не образована или solver не сошелся.
Private mDiagnosticLog As String ' Текстовая диагностика выбора зоны, арматуры, ls и psi_s.
Private mDiagnosticsEnabled As Boolean ' Отключает подробный журнал, когда txt-отчет выполнения не нужен.
Private mSolverCallCount As Long ' Сколько внутренних CSectionSolver/CStateSolutionRunner-вызовов выполнил расчет трещин.

Private mAllowableCrackWidth As Double ' Пользовательский предел a_crc,ult во внутренних единицах длины, мм.
Private mCrackWidth As Double ' Расчетная ширина раскрытия a_crc, мм.
Private mUtilization As Double ' Коэффициент использования по трещине: a_crc / a_crc,ult.
Private mCrackFormed As Boolean ' True, если по выбранной ветке расчета трещина считается образованной.

Private mPsiMode As String ' User или Auto: брать psi_s из Config или уточнять после неудачной первой проверки.
Private mPhi3Mode As String ' User или Auto: брать phi3 из Config или выбирать по характеру продольной силы.
Private mTensionZoneMode As String ' Effective или FullTension: способ задания бетонной зоны Abt.
Private mCoverMode As String ' NearestContour или GlobalExtreme: способ измерения расстояния a_s.
Private mCrackFormationPath As String ' Auto, LambdaMxy, LambdaN или LambdaNMxy: какая часть LC масштабируется при поиске образования трещины.
Private mResolvedCrackFormationPath As String ' Фактический путь, который дал принятую точку образования трещины.
Private mCrackFormationSolutionStrategy As String ' Auto, UltimateStrain или LoadMultiplier для поиска стадии образования трещины.
Private mPhi1Input As Double ' Пользовательское значение phi1 для формулы раскрытия.
Private mPhi2Input As Double ' Пользовательское значение phi2 для формулы раскрытия.
Private mPhi3Input As Double ' Пользовательское значение phi3, если Phi3Mode = User.
Private mPsiSInput As Double ' Пользовательское значение psi_s, если PsiMode = User.
Private mConcreteRbtSer As Double ' Rbt,ser для центральной ветки Ncrc и модели образования трещины, МПа.
Private mConcreteEb As Double ' Eb для Ared в центральной ветке Ncrc, МПа.
Private mConcreteEbt0 As Double ' eps_bt0 для формулы (8.54) СП 63 при однозначно растянутой эпюре.
Private mConcreteEbt2 As Double ' eps_bt2: предельная растягивающая деформация при двузначной эпюре.
Private mSteelEs As Double ' Es для формулы раскрытия трещины, МПа.
Private mCrackInitiationSpec As CMaterialModelSpec ' Спецификация модели образования трещины для обязательной предварительной проверки.
Private mCrackedStateSpec As CMaterialModelSpec ' Спецификация модели раскрытой трещины текущего LC.
Private mBeforeMcrcState As CSectionStateResult ' Необязательное состояние до выключения растянутого бетона.
Private mBeforeMcrcSolver As CSectionSolver ' Solver состояния BeforeMcrcState для warm start состояния сразу после трещины.
Private mAfterMcrcState As CSectionStateResult ' Необязательное состояние после выключения растянутого бетона для sigma_s,crc.
Private mAfterMcrcSolver As CSectionSolver ' Solver состояния AfterMcrcState, чтобы не решать его повторно для sigma_s,crc.
Private mFormationPrechecked As Boolean ' True, если Ncrc/Mcrc уже проверены перед расчетом полной ширины.
Private mFormationLoadDefined As Boolean ' True, если сохранены внутренние усилия именно на уровне образования трещины.
Private mFormationLoadN As Double ' Продольная сила служебного состояния образования трещины, Н.
Private mFormationLoadMx As Double ' Внутренний Mx служебного состояния образования трещины, Н*мм.
Private mFormationLoadMy As Double ' Внутренний My служебного состояния образования трещины, Н*мм.
Private mFormationPathFallbackPsi1 As Boolean ' True, если выбранная lambda-траектория уже треснула в постоянной части и sigma_s,crc не имеет честной точки.

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
Private mFormationNcrc As Double ' Продольная сила в общей N+M ветке на найденной стадии образования трещины, Н.
Private mMcrc As Double ' Модуль изгибающего момента на стадии образования трещины, Н*мм.
Private mAred As Double ' Площадь приведенного сечения для Ncrc при центральном растяжении, мм2.
Private mTensionRebarIds As String ' Имена стержней, вошедших в расчет sigma_s и As, через запятую.

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
    mPsiMode = "Auto"
    mPhi3Mode = "Auto"
    mTensionZoneMode = "Effective"
    mCoverMode = COVER_MODE_GLOBAL_EXTREME
    mCrackFormationPath = CRACK_PATH_AUTO
    mCrackFormationSolutionStrategy = STRATEGY_AUTO
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
    mCrackFormationPath = settings.GetString("SLS.Crack.InitiationLoadPath", mCrackFormationPath)
    mCrackFormationSolutionStrategy = settings.GetString("SLS.Crack.InitiationSolutionStrategy", mCrackFormationSolutionStrategy)
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
' Метод оставлен как компактный сценарий для тестов: сначала выполняется
' обязательная проверка образования трещины, затем по переданному CrackedState
' считается ширина раскрытия. В рабочем batch эти два шага разделены явно:
' orchestrator всегда сохраняет текущее CrackedState для профиля трещин, а
' calculator считает a_crc только при положительной проверке образования.
Public Sub Calculate(ByVal currentSolver As CSectionSolver, ByVal section As CSectionModel, _
        ByVal materialProvider As CMaterialModelProvider, ByVal crackedStateSpec As CMaterialModelSpec, _
        ByVal crackInitiationSpec As CMaterialModelSpec, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double, _
        ByVal loadStateForClassification As CSectionLoadState, _
        ByVal centralReferenceX As Double, ByVal centralReferenceY As Double)
    If Not CheckFormationBeforeWidth(section, materialProvider, crackedStateSpec, crackInitiationSpec, _
            targetN, targetMx, targetMy, loadStateForClassification, centralReferenceX, centralReferenceY) Then Exit Sub
    CalculateAfterFormationCheck currentSolver, section, materialProvider, crackedStateSpec, crackInitiationSpec, _
        targetN, targetMx, targetMy, loadStateForClassification, centralReferenceX, centralReferenceY
End Sub

' Продолжает расчет ширины после обязательной предварительной проверки
' образования трещины. Метод сохраняет уже найденные Before/After-состояния
' порога трещинообразования и использует AfterMcrcState для sigma_s,crc,
' если Auto-подбор psi_s потребует уточнения коэффициента.
Public Sub CalculateAfterFormationCheck(ByVal currentSolver As CSectionSolver, ByVal section As CSectionModel, _
        ByVal materialProvider As CMaterialModelProvider, ByVal crackedStateSpec As CMaterialModelSpec, _
        ByVal crackInitiationSpec As CMaterialModelSpec, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double, _
        ByVal loadStateForClassification As CSectionLoadState, _
        ByVal centralReferenceX As Double, ByVal centralReferenceY As Double)
    If Not mFormationPrechecked Then _
        Err.Raise vbObjectError + 3724, "CCrackWidthCalculator", "Предварительная проверка образования трещины не выполнена."
    If Not mCrackFormed Then _
        Err.Raise vbObjectError + 3725, "CCrackWidthCalculator", "Нельзя считать ширину: проверка показала, что трещина не образована."

    CalculateCore currentSolver, section, materialProvider, crackedStateSpec, crackInitiationSpec, _
        targetN, targetMx, targetMy, loadStateForClassification, centralReferenceX, centralReferenceY
End Sub

' Выполняет только проверку факта образования трещины по Ncrc или по общей
' траектории N + M.
' Если уровень образования трещины найден, сохраняются состояния
' BeforeMcrcState/AfterMcrcState даже при lambda_crc > 1: они описывают
' пороговое состояние, а не обязательно фактическую трещину при текущем LC.
' Сама ширина a_crc еще не считается: для нее нужен полный CrackedState LC.
Public Function CheckFormationBeforeWidth(ByVal section As CSectionModel, _
        ByVal materialProvider As CMaterialModelProvider, ByVal crackedStateSpec As CMaterialModelSpec, _
        ByVal crackInitiationSpec As CMaterialModelSpec, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double, _
        ByVal loadStateForClassification As CSectionLoadState, _
        ByVal centralReferenceX As Double, ByVal centralReferenceY As Double) As Boolean
    ClearResult
    ValidateFormationInputs section, materialProvider, loadStateForClassification, crackedStateSpec, crackInitiationSpec
    Set mCrackedStateSpec = crackedStateSpec.Clone
    If crackInitiationSpec.IsComplete Then Set mCrackInitiationSpec = crackInitiationSpec.Clone
    ApplyMaterialProviderProperties materialProvider
    NormalizeSettings

    mFormationPrechecked = True
    If IsCentralTensionState(loadStateForClassification, centralReferenceX, centralReferenceY) Then
        CheckCentralTensionFormation section, materialProvider, targetN, targetMx, targetMy, loadStateForClassification
    Else
        CheckGeneralFormation section, materialProvider, targetN, targetMx, targetMy, _
            loadStateForClassification, centralReferenceX, centralReferenceY
    End If
    CheckFormationBeforeWidth = (mConverged And mCrackFormed)
End Function

' Общая часть расчета ширины раскрытия. В обычном режиме вызывается после
' ClearResult, а после обязательной проверки образования трещины работает
' поверх уже сохраненных пороговых состояний.
Private Sub CalculateCore(ByVal currentSolver As CSectionSolver, ByVal section As CSectionModel, _
        ByVal materialProvider As CMaterialModelProvider, ByVal crackedStateSpec As CMaterialModelSpec, _
        ByVal crackInitiationSpec As CMaterialModelSpec, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double, _
        ByVal loadStateForClassification As CSectionLoadState, _
        ByVal centralReferenceX As Double, ByVal centralReferenceY As Double)
    ValidateInputs currentSolver, section, materialProvider
    If loadStateForClassification Is Nothing Then _
        Err.Raise vbObjectError + 3715, "CCrackWidthCalculator", "CSectionLoadState не передан."
    If crackedStateSpec Is Nothing Or Not crackedStateSpec.IsComplete Then _
        Err.Raise vbObjectError + 3714, "CCrackWidthCalculator", "MaterialModel.CrackedState не передан."
    Set mCrackedStateSpec = crackedStateSpec.Clone
    If Not crackInitiationSpec Is Nothing Then
        If crackInitiationSpec.IsComplete Then Set mCrackInitiationSpec = crackInitiationSpec.Clone
    End If
    ApplyMaterialProviderProperties materialProvider
    NormalizeSettings
    mConverged = False
    mCrackFormed = False
    mStopReason = vbNullString
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
    mCrackWidth = 0#
    mUtilization = 0#

    If IsCentralTensionState(loadStateForClassification, centralReferenceX, centralReferenceY) Then
        CalculateCentralTension currentSolver, section, materialProvider, targetN, targetMx, targetMy, _
            loadStateForClassification
    Else
        CalculateGeneralBending currentSolver, section, materialProvider, targetN, targetMx, targetMy, _
            loadStateForClassification
    End If
End Sub

' Проверяет минимальный набор данных для предварительной проверки образования
' нормальной трещины.
' В отличие от основного Calculate здесь еще нет полного CrackedState solver-а,
' поэтому проверяется только геометрия, material provider и модель
' CrackInitiation.
Private Sub ValidateFormationInputs(ByVal section As CSectionModel, ByVal materialProvider As CMaterialModelProvider, _
        ByVal loadStateForClassification As CSectionLoadState, ByVal crackedStateSpec As CMaterialModelSpec, _
        ByVal crackInitiationSpec As CMaterialModelSpec)
    If section Is Nothing Then Err.Raise vbObjectError + 3726, "CCrackWidthCalculator", "Модель расчетного сечения не передана."
    If section.ConcreteCount <= 0 Then Err.Raise vbObjectError + 3727, "CCrackWidthCalculator", "В модели расчетного сечения нет бетонных элементов."
    If section.RebarCount <= 0 Then Err.Raise vbObjectError + 3728, "CCrackWidthCalculator", "В модели расчетного сечения нет арматурных элементов."
    If materialProvider Is Nothing Then Err.Raise vbObjectError + 3729, "CCrackWidthCalculator", "Material provider не передан."
    If loadStateForClassification Is Nothing Then Err.Raise vbObjectError + 3730, "CCrackWidthCalculator", "CSectionLoadState не передан."
    If crackedStateSpec Is Nothing Or Not crackedStateSpec.IsComplete Then _
        Err.Raise vbObjectError + 3731, "CCrackWidthCalculator", "MaterialModel.CrackedState не передан."
    If crackInitiationSpec Is Nothing Or Not crackInitiationSpec.IsComplete Then _
        Err.Raise vbObjectError + 3732, "CCrackWidthCalculator", "Для Calculation.Crack.Width нужна MaterialModel.CrackInitiation."
End Sub

' Проверяет центральное растяжение по Ncrc. Для единого snapshot-а состояния
' до и после трещины сохраняются так же, как в общей ветке образования
' трещины.
Private Sub CheckCentralTensionFormation(ByVal section As CSectionModel, _
        ByVal materialProvider As CMaterialModelProvider, ByVal targetN As Double, _
        ByVal targetMx As Double, ByVal targetMy As Double, _
        ByVal loadStateForClassification As CSectionLoadState)
    mCentralTensionBranch = True
    mPhi3 = ResolvePhi3(loadStateForClassification)

    If Not IsTensionLoad(loadStateForClassification) Then
        FinishNoCrack "Центральная ветка: продольная сила не является растягивающей."
        Exit Sub
    End If

    mNcrc = CentralTensionCrackForce(section)
    mFormationNcrc = mNcrc
    mMcrc = 0#
    If mNcrc <= 0# Or targetN <= 0# Then
        mStopReason = "InputError: для центрального растяжения не удалось определить Ncrc."
        Exit Sub
    End If

    Dim activePath As String
    activePath = mCrackFormationPath
    If StrComp(activePath, CRACK_PATH_AUTO, vbTextCompare) = 0 Then
        AppendDiagnostic "Auto-путь: центральное растяжение распознано до перебора путей; " & ChrW$(&H3BB) & _
            "*Mxy не масштабирует продольную силу N, поэтому эта попытка пропущена и программа переходит к " & _
            ChrW$(&H3BB) & "*N."
        activePath = CRACK_PATH_N
    End If

    If StrComp(activePath, CRACK_PATH_MXY, vbTextCompare) = 0 Then
        If targetN >= mNcrc Then
            MarkFormationPathFallback "Центральное растяжение уже образует трещину, но выбран путь lambda*Mxy без масштабируемого момента; psi_s остается равным 1."
            mResolvedCrackFormationPath = activePath
        Else
            FinishNoCrack "Трещина не образована: Ncrc больше заданной растягивающей силы, а выбранный путь lambda*Mxy не масштабирует N."
        End If
        Exit Sub
    End If

    mLambdaCrc = mNcrc / targetN
    If mLambdaCrc <= 0# Then
        mStopReason = "NumericalFailure: найденное lambda_crc по Ncrc неположительно."
        Exit Sub
    End If

    StoreFormationLoad targetN * mLambdaCrc, targetMx * mLambdaCrc, targetMy * mLambdaCrc
    mResolvedCrackFormationPath = activePath
    If Not BuildBeforeMcrcState(section, materialProvider, targetN, targetMx, targetMy) Then Exit Sub
    If Not EnsureAfterMcrcState(section, materialProvider, Nothing, targetN, targetMx, targetMy) Then Exit Sub

    If mLambdaCrc > 1# Then
        FinishNoCrack "Трещина не образована: Ncrc больше заданной растягивающей силы."
        Exit Sub
    End If

    mConverged = True
    mCrackFormed = True
    mStopReason = "Предварительная проверка Ncrc: трещина образуется при текущей нагрузке."
    AppendDiagnostic "formation=Ncrc; lambda_crc=" & FormatNumberInvariant(mLambdaCrc)
End Sub

' Проверяет общую изгибную/внецентренную ветку по предельной растягивающей
' деформации бетона. Состояние AfterMcrcState решается сразу, чтобы затем не
' пересчитывать его при уточнении psi_s.
Private Sub CheckGeneralFormation(ByVal section As CSectionModel, _
        ByVal materialProvider As CMaterialModelProvider, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double, _
        ByVal loadStateForClassification As CSectionLoadState, _
        ByVal centralReferenceX As Double, ByVal centralReferenceY As Double)
    If mCrackInitiationSpec Is Nothing Then
        mStopReason = "InputError: для Calculation.Crack.Width требуется MaterialModel.CrackInitiation."
        Exit Sub
    End If
    If Not HasNormalCrackDrivingAction(loadStateForClassification, centralReferenceX, centralReferenceY) Then
        FinishNoCrack "Трещина не образована: нет растягивающей силы или изгибающего момента, вызывающего нормальную трещину."
        Exit Sub
    End If

    Dim momentMxAtReference As Double
    Dim momentMyAtReference As Double
    momentMxAtReference = loadStateForClassification.MxAboutPoint(centralReferenceY)
    momentMyAtReference = loadStateForClassification.MyAboutPoint(centralReferenceX)

    If Not FindGeneralFormationLambda(section, materialProvider, targetN, targetMx, targetMy, _
            momentMxAtReference, momentMyAtReference, centralReferenceX, centralReferenceY) Then Exit Sub

    If mFormationLoadDefined Then
        mFormationNcrc = FormationNValue()
        mMcrc = CrackMomentFromFormationLoad(centralReferenceX, centralReferenceY)
    End If
    If mFormationPathFallbackPsi1 Then Exit Sub

    If Not EnsureAfterMcrcState(section, materialProvider, Nothing, targetN, targetMx, targetMy) Then Exit Sub

    If mLambdaCrc > 1# Then
        FinishNoCrack "Трещина не образована: предельная растянутая деформация бетона не достигнута при lambda = 1."
        Exit Sub
    End If

    mConverged = True
    mCrackFormed = True
    mStopReason = "Предварительная проверка Mcrc: трещина образуется при текущей нагрузке."
End Sub

' Выбирает численную стратегию поиска стадии образования трещины для общей
' ветки N+M. Путь нагрузки и метод решения разделены так же, как в capacity:
' UltimateStrain ищет плоскость по eps_bt,ult, LoadMultiplier перебирает
' lambda обычными solve, Auto сначала пробует UltimateStrain и только при
' численной неудаче переходит на LoadMultiplier.
Private Function FindGeneralFormationLambda(ByVal section As CSectionModel, _
        ByVal materialProvider As CMaterialModelProvider, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double, _
        ByVal momentMxAtReference As Double, ByVal momentMyAtReference As Double, _
        ByVal referenceX As Double, ByVal referenceY As Double) As Boolean
    If StrComp(mCrackFormationPath, CRACK_PATH_AUTO, vbTextCompare) = 0 Then
        FindGeneralFormationLambda = FindGeneralFormationLambdaAuto(section, materialProvider, _
            targetN, targetMx, targetMy, momentMxAtReference, momentMyAtReference, referenceX, referenceY)
        Exit Function
    End If

    FindGeneralFormationLambda = FindGeneralFormationLambdaForCurrentPath(section, materialProvider, _
        targetN, targetMx, targetMy, momentMxAtReference, momentMyAtReference, referenceX, referenceY)
    If FindGeneralFormationLambda Then mResolvedCrackFormationPath = mCrackFormationPath
End Function

' В Auto-режиме по очереди пробует все разрешенные lambda-пути. Если путь
' дал только резервный psi_s=1 без найденного состояния образования трещины,
' это считается неудачной попыткой Auto и программа идет дальше.
Private Function FindGeneralFormationLambdaAuto(ByVal section As CSectionModel, _
        ByVal materialProvider As CMaterialModelProvider, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double, _
        ByVal momentMxAtReference As Double, ByVal momentMyAtReference As Double, _
        ByVal referenceX As Double, ByVal referenceY As Double) As Boolean
    Dim savedPath As String
    savedPath = mCrackFormationPath

    Dim paths(1 To 3) As String
    paths(1) = CRACK_PATH_MXY
    paths(2) = CRACK_PATH_N
    paths(3) = CRACK_PATH_NMXY
    AppendDiagnostic "Auto-путь: начинаем поиск образования нормальной трещины. Порядок попыток: " & _
        ChrW$(&H3BB) & "*Mxy -> " & ChrW$(&H3BB) & "*N -> " & ChrW$(&H3BB) & "*NMxy."

    Dim lastStopReason As String
    Dim i As Long
    For i = LBound(paths) To UBound(paths)
        ClearFormationAttemptState
        mCrackFormationPath = paths(i)
        AppendDiagnostic "Auto-путь: попытка " & CStr(i) & " из " & CStr(UBound(paths)) & _
            " - пробуем " & CrackFormationPathDisplayName() & "."

        If FindGeneralFormationLambdaForCurrentPath(section, materialProvider, _
                targetN, targetMx, targetMy, momentMxAtReference, momentMyAtReference, referenceX, referenceY) Then
            If mFormationPathFallbackPsi1 And i < UBound(paths) Then
                lastStopReason = mStopReason
                AppendDiagnostic "Auto-путь: " & CrackFormationPathDisplayName() & _
                    " дал только резервный результат psi_s=1 без найденной точки образования трещины. Для более информативного результата переходим к следующему пути."
            Else
                mResolvedCrackFormationPath = mCrackFormationPath
                AppendDiagnostic "Auto-путь: выбран " & CrackFormationPathDisplayName() & "."
                mCrackFormationPath = savedPath
                FindGeneralFormationLambdaAuto = True
                Exit Function
            End If
        Else
            lastStopReason = mStopReason
            AppendDiagnostic "Auto-путь: " & CrackFormationPathDisplayName() & _
                " не дал точку образования трещины: " & mStopReason
        End If
    Next i

    mCrackFormationPath = savedPath
    If Len(mStopReason) = 0 Then mStopReason = lastStopReason
End Function

' Выполняет поиск по уже выбранному mCrackFormationPath без перебора путей.
Private Function FindGeneralFormationLambdaForCurrentPath(ByVal section As CSectionModel, _
        ByVal materialProvider As CMaterialModelProvider, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double, _
        ByVal momentMxAtReference As Double, ByVal momentMyAtReference As Double, _
        ByVal referenceX As Double, ByVal referenceY As Double) As Boolean
    Dim concreteMaterial As CMaterialDiagram
    Dim steelMaterial As CMaterialDiagram
    Set concreteMaterial = materialProvider.ConcreteMaterialFromSpec(mCrackInitiationSpec)
    Set steelMaterial = materialProvider.SteelMaterialFromSpec(mCrackInitiationSpec)

    Select Case mCrackFormationSolutionStrategy
        Case STRATEGY_LOAD_MULTIPLIER
            FindGeneralFormationLambdaForCurrentPath = CalculateLambdaCrcByLoadScaling(section, concreteMaterial, steelMaterial, _
                Nothing, targetN, momentMxAtReference, momentMyAtReference, referenceX, referenceY)
            Exit Function
        Case STRATEGY_ULTIMATE_STRAIN
            FindGeneralFormationLambdaForCurrentPath = CalculateLambdaCrcByUltimateStrain(section, concreteMaterial, steelMaterial, _
                targetN, targetMx, targetMy, momentMxAtReference, momentMyAtReference, referenceX, referenceY)
            Exit Function
    End Select

    Dim firstStopReason As String
    If CalculateLambdaCrcByUltimateStrain(section, concreteMaterial, steelMaterial, _
            targetN, targetMx, targetMy, momentMxAtReference, momentMyAtReference, referenceX, referenceY) Then
        FindGeneralFormationLambdaForCurrentPath = True
        Exit Function
    End If
    If FormationSearchStoppedTerminally() Then Exit Function

    firstStopReason = mStopReason
    AppendDiagnostic "UltimateStrain не дал точку образования трещины: " & firstStopReason
    AppendDiagnostic "Пробуем LoadMultiplier по тому же пути образования трещины."
    mStopReason = vbNullString

    If CalculateLambdaCrcByLoadScaling(section, concreteMaterial, steelMaterial, Nothing, _
            targetN, momentMxAtReference, momentMyAtReference, referenceX, referenceY) Then
        FindGeneralFormationLambdaForCurrentPath = True
    ElseIf Len(mStopReason) = 0 Then
        mStopReason = firstStopReason
    End If
End Function

' Пытается найти lambda_crc прямым решением по предельной растягивающей
' деформации бетона. Для lambda*Mxy используется специализированная система
' с фиксированным N и направлением момента; для остальных путей - общая
' load-path система Offset + lambda*Base. Обе ветки являются одним и тем же
' методом UltimateStrain с разной формой траектории нагрузки.
Private Function CalculateLambdaCrcByUltimateStrain(ByVal section As CSectionModel, _
        ByVal concreteMaterial As CMaterialDiagram, ByVal steelMaterial As CMaterialDiagram, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double, _
        ByVal momentMxAtReference As Double, ByVal momentMyAtReference As Double, _
        ByVal referenceX As Double, ByVal referenceY As Double) As Boolean
    AppendDiagnostic "UltimateStrain: начинаем прямой поиск точки образования трещины по eps_bt,ult; путь " & _
        CrackFormationPathDisplayName() & "."

    If TryMarkConstantPartFormationFallback(section, concreteMaterial, steelMaterial, _
            FormationLoadPath(targetN, momentMxAtReference, momentMyAtReference, referenceX, referenceY), 0#) Then
        CalculateLambdaCrcByUltimateStrain = True
        Exit Function
    End If

    If TryUltimateStrainFormation(section, concreteMaterial, steelMaterial, Nothing, _
            targetN, momentMxAtReference, momentMyAtReference, referenceX, referenceY) Then
        CalculateLambdaCrcByUltimateStrain = True
        Exit Function
    End If
    If FormationSearchStoppedTerminally() Then Exit Function

    Dim firstStopReason As String
    firstStopReason = mStopReason
    AppendDiagnostic "UltimateStrain: быстрый старт не дал устойчивую точку образования трещины: " & firstStopReason
    AppendDiagnostic "UltimateStrain: пробуем дополнительный старт от полного НДС CrackInitiation."
    mStopReason = vbNullString

    ' Страховочный старт от полного состояния нужен только если быстрый старт
    ' не сработал. Это сохраняет производительность и одновременно помогает
    ' сложным сочетаниям, где уже найденная плоскость НДС дает лучший initial.
    Dim formationStart As CSectionSolver
    Set formationStart = SolveState(section, concreteMaterial, steelMaterial, Nothing, targetN, targetMx, targetMy)
    If formationStart Is Nothing Or Not formationStart.Converged Then
        If Len(mStopReason) = 0 Then mStopReason = firstStopReason
        Exit Function
    End If

    If TryUltimateStrainFormation(section, concreteMaterial, steelMaterial, formationStart, _
            targetN, momentMxAtReference, momentMyAtReference, referenceX, referenceY) Then
        AppendDiagnostic "UltimateStrain: дополнительный старт от полного НДС успешно нашел точку образования трещины."
        CalculateLambdaCrcByUltimateStrain = True
    ElseIf Len(mStopReason) = 0 Then
        AppendDiagnostic "UltimateStrain: дополнительный старт от полного НДС также не дал точку образования трещины."
        mStopReason = firstStopReason
    End If
End Function

' Выполняет одну попытку UltimateStrain на выбранном пути нагрузки.
Private Function TryUltimateStrainFormation(ByVal section As CSectionModel, _
        ByVal concreteMaterial As CMaterialDiagram, ByVal steelMaterial As CMaterialDiagram, _
        ByVal startSolver As CSectionSolver, ByVal targetN As Double, _
        ByVal momentMxAtReference As Double, ByVal momentMyAtReference As Double, _
        ByVal referenceX As Double, ByVal referenceY As Double) As Boolean
    If StrComp(mCrackFormationPath, CRACK_PATH_MXY, vbTextCompare) = 0 Then
        TryUltimateStrainFormation = CalculateLambdaCrc(section, concreteMaterial, steelMaterial, _
            startSolver, targetN, momentMxAtReference, momentMyAtReference, referenceX, referenceY)
    Else
        Dim loadPath As CLoadPathVector
        Set loadPath = FormationLoadPath(targetN, momentMxAtReference, momentMyAtReference, referenceX, referenceY)
        TryUltimateStrainFormation = CalculateLambdaCrcByLoadPathNewton(section, concreteMaterial, steelMaterial, _
            loadPath, startSolver)
    End If
End Function

' Возвращает True, когда результат прямого поиска уже окончательный и Auto
' не должен переходить к LoadMultiplier: это физическое отсутствие трещины,
' fallback с psi_s=1 или ошибка входных данных.
Private Function FormationSearchStoppedTerminally() As Boolean
    FormationSearchStoppedTerminally = (mConverged And Not mCrackFormed) Or _
        mFormationPathFallbackPsi1 Or _
        (InStr(1, mStopReason, "InputError", vbTextCompare) > 0)
End Function

' Создает BeforeMcrcState для центрального растяжения, где lambda_crc задана
' через Ncrc и нет отдельного прямого поиска плоскости по eps_bt,ult.
Private Function BuildBeforeMcrcState(ByVal section As CSectionModel, ByVal materialProvider As CMaterialModelProvider, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double) As Boolean
    If mBeforeMcrcState Is Nothing Then
        Dim beforeSolver As CSectionSolver
        Dim formationN As Double
        Dim formationMx As Double
        Dim formationMy As Double
        FormationLoads formationN, formationMx, formationMy
        Set beforeSolver = SolveState(section, materialProvider.ConcreteMaterialFromSpec(mCrackInitiationSpec), _
            materialProvider.SteelMaterialFromSpec(mCrackInitiationSpec), Nothing, _
            formationN, formationMx, formationMy)
        If Not beforeSolver.Converged Then
            mStopReason = "NumericalFailure: состояние Ncrc до образования трещины не сошлось: " & beforeSolver.StopReason
            Exit Function
        End If
        Set mBeforeMcrcSolver = beforeSolver
        Set mBeforeMcrcState = New CSectionStateResult
        mBeforeMcrcState.InitializeFromSolver sstBeforeMcrcState, cpMcrc, _
            mCrackInitiationSpec, beforeSolver, False, "OK"
    End If
    BuildBeforeMcrcState = True
End Function

' Возвращает AfterMcrcState для уже найденного lambda_crc. Предварительная
' проверка создает это состояние заранее, а Auto-уточнение psi_s потом только
' читает sigma_s,crc из сохраненного solver-а без повторного решения.
Private Function EnsureAfterMcrcState(ByVal section As CSectionModel, ByVal materialProvider As CMaterialModelProvider, _
        ByVal startSolver As CSectionSolver, ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double) As Boolean
    If Not mAfterMcrcSolver Is Nothing Then
        If mAfterMcrcSolver.Converged Then
            EnsureAfterMcrcState = True
            Exit Function
        End If
    End If

    Dim effectiveStartSolver As CSectionSolver
    Set effectiveStartSolver = startSolver
    If effectiveStartSolver Is Nothing Then Set effectiveStartSolver = mBeforeMcrcSolver

    Dim formationN As Double
    Dim formationMx As Double
    Dim formationMy As Double
    FormationLoads formationN, formationMx, formationMy
    Set mAfterMcrcSolver = SolveState(section, materialProvider.ConcreteMaterialFromSpec(mCrackedStateSpec), _
        materialProvider.SteelMaterialFromSpec(mCrackedStateSpec), effectiveStartSolver, _
        formationN, formationMx, formationMy)
    If Not mAfterMcrcSolver.Converged Then
        mStopReason = "NumericalFailure: состояние lambda_crc после образования трещины не сошлось: " & mAfterMcrcSolver.StopReason
        Exit Function
    End If

    Set mAfterMcrcState = New CSectionStateResult
    mAfterMcrcState.InitializeFromSolver sstAfterMcrcState, cpCrackedNDS, _
        mCrackedStateSpec, mAfterMcrcSolver, False, "OK"
    EnsureAfterMcrcState = True
End Function

' Сохраняет внутренние усилия, соответствующие уже найденной стадии
' трещинообразования. Дальше Before/AfterMcrcState решаются именно по ним,
' чтобы повторные вызовы не пересчитывали lambda и не масштабировали не ту
' часть нагрузки.
Private Sub StoreFormationLoad(ByVal nValue As Double, ByVal mxValue As Double, ByVal myValue As Double)
    mFormationLoadN = nValue
    mFormationLoadMx = mxValue
    mFormationLoadMy = myValue
    mFormationLoadDefined = True
End Sub

' Для общей N+M ветки переводит моментный множитель lambda обратно во
' внутренние усилия solver-а: N остается постоянной, а масштабируется только
' момент относительно расчетного центра образования трещины.
Private Sub StoreMomentFormationLoad(ByVal targetN As Double, ByVal targetMomentMx As Double, _
        ByVal targetMomentMy As Double, ByVal referenceX As Double, ByVal referenceY As Double)
    Dim loadPath As CLoadPathVector
    Set loadPath = MomentFormationLoadPath(targetN, targetMomentMx, targetMomentMy, referenceX, referenceY)

    StoreFormationLoadFromPath loadPath, mLambdaCrc
End Sub

' Сохраняет точку выбранной lambda-траектории как внутренние усилия solver-а.
Private Sub StoreFormationLoadFromPath(ByVal loadPath As CLoadPathVector, ByVal lambdaValue As Double)
    Dim nValue As Double
    Dim mxValue As Double
    Dim myValue As Double
    loadPath.Target lambdaValue, nValue, mxValue, myValue
    StoreFormationLoad nValue, mxValue, myValue
End Sub

' Возвращает сохраненные усилия трещинообразования. Они должны быть
' определены в той же ветке, которая нашла lambda_crc.
Private Sub FormationLoads(ByRef nValue As Double, ByRef mxValue As Double, ByRef myValue As Double)
    If Not mFormationLoadDefined Then _
        Err.Raise vbObjectError + 3733, "CCrackWidthCalculator", "Внутренние усилия состояния образования трещины не сохранены."
    nValue = mFormationLoadN
    mxValue = mFormationLoadMx
    myValue = mFormationLoadMy
End Sub

' Создает общую lambda-траекторию для общей ветки образования трещины:
' продольная сила остается постоянной, а моментный вектор относительно
' расчетного центра масштабируется. Сама формула Offset + lambda*Base живет
' в CLoadPathVector, поэтому трещины и Capacity не держат две разные
' реализации load-path.
Private Function MomentFormationLoadPath(ByVal targetN As Double, _
        ByVal targetMomentMx As Double, ByVal targetMomentMy As Double, _
        ByVal referenceX As Double, ByVal referenceY As Double) As CLoadPathVector
    Dim loadPath As CLoadPathVector
    Set loadPath = New CLoadPathVector
    loadPath.InitializeFixedNAndMomentAboutPoint targetN, targetMomentMx, targetMomentMy, referenceX, referenceY
    Set MomentFormationLoadPath = loadPath
End Function

' Создает lambda-траекторию образования трещины по выбранной настройке
' SLS.Crack.InitiationLoadPath. Моменты targetMomentMx/My передаются
' относительно расчетного центра трещин, а CSectionSolver получает моменты
' относительно начала модели, поэтому вклад масштабируемой N переносится
' в Base, когда путь масштабирует продольную силу.
Private Function FormationLoadPath(ByVal targetN As Double, _
        ByVal targetMomentMx As Double, ByVal targetMomentMy As Double, _
        ByVal referenceX As Double, ByVal referenceY As Double) As CLoadPathVector
    Dim loadPath As CLoadPathVector
    Set loadPath = New CLoadPathVector

    Select Case mCrackFormationPath
        Case CRACK_PATH_N
            loadPath.Initialize 0#, targetN, _
                targetMomentMx, targetN * referenceY, _
                targetMomentMy, targetN * referenceX
        Case CRACK_PATH_NMXY
            loadPath.Initialize 0#, targetN, _
                0#, targetMomentMx + targetN * referenceY, _
                0#, targetMomentMy + targetN * referenceX
        Case Else
            loadPath.InitializeFixedNAndMomentAboutPoint targetN, targetMomentMx, targetMomentMy, referenceX, referenceY
    End Select
    Set FormationLoadPath = loadPath
End Function

' Проверяет, есть ли у выбранного пути ненулевая часть Base. Если Base нет,
' нижняя точка lambda=0 уже является единственной проверяемой нагрузкой.
Private Function FormationPathHasScaledLoad(ByVal targetN As Double, _
        ByVal targetMomentMx As Double, ByVal targetMomentMy As Double) As Boolean
    Dim momentNorm As Double
    momentNorm = Sqr(targetMomentMx * targetMomentMx + targetMomentMy * targetMomentMy)
    Select Case mCrackFormationPath
        Case CRACK_PATH_N
            FormationPathHasScaledLoad = (Abs(targetN) > NEAR_ZERO_FORCE)
        Case CRACK_PATH_NMXY
            FormationPathHasScaledLoad = (Abs(targetN) > NEAR_ZERO_FORCE Or _
                momentNorm > MaxDouble(mSolverToleranceMx, mSolverToleranceMy))
        Case Else
            FormationPathHasScaledLoad = (momentNorm > MaxDouble(mSolverToleranceMx, mSolverToleranceMy))
    End Select
End Function

' Возвращает пользовательское имя пути для диагностики.
Private Function CrackFormationPathDisplayName() As String
    Select Case mCrackFormationPath
        Case CRACK_PATH_AUTO
            CrackFormationPathDisplayName = "Auto"
        Case CRACK_PATH_N
            CrackFormationPathDisplayName = ChrW$(&H3BB) & "*N"
        Case CRACK_PATH_NMXY
            CrackFormationPathDisplayName = ChrW$(&H3BB) & "*NMxy"
        Case Else
            CrackFormationPathDisplayName = ChrW$(&H3BB) & "*Mxy"
    End Select
End Function

' Возвращает имя фактически использованного пути. Для Auto это уже не Auto,
' а конкретный lambda-путь, на котором была принята стадия образования.
Private Function CrackFormationPathDisplayNameForKey(ByVal pathKey As String) As String
    Select Case pathKey
        Case CRACK_PATH_N
            CrackFormationPathDisplayNameForKey = ChrW$(&H3BB) & "*N"
        Case CRACK_PATH_NMXY
            CrackFormationPathDisplayNameForKey = ChrW$(&H3BB) & "*NMxy"
        Case CRACK_PATH_MXY
            CrackFormationPathDisplayNameForKey = ChrW$(&H3BB) & "*Mxy"
        Case CRACK_PATH_AUTO
            CrackFormationPathDisplayNameForKey = "Auto"
        Case Else
            CrackFormationPathDisplayNameForKey = vbNullString
    End Select
End Function

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
        ByVal targetMx As Double, ByVal targetMy As Double, _
        ByVal loadStateForClassification As CSectionLoadState)
    mCentralTensionBranch = True
    mPhi3 = ResolvePhi3(loadStateForClassification)

    If Not IsTensionLoad(loadStateForClassification) Then
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

    ' При fallback выбранного пути lambda не является физически найденной точкой,
    ' поэтому сохраняем диагностический lambda = 0 и только выводим Ncrc.
    If Not mFormationPathFallbackPsi1 Then
        mLambdaCrc = mNcrc / targetN
        If mLambdaCrc > 1# Then
            FinishNoCrack "Трещина не образована: Ncrc больше заданной растягивающей силы."
            Exit Sub
        End If
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
        ByVal targetMx As Double, ByVal targetMy As Double, _
        ByVal loadStateForClassification As CSectionLoadState)
    mPhi3 = ResolvePhi3(loadStateForClassification)

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

' Находит lambda_crc прямым методом UltimateStrain для пути lambda*Mxy.
' Продольная сила
' остается равной заданной N, а lambda масштабирует только моментный вектор
' относительно расчетного центра образования трещины. Критерием служит достижение
' eps_bt,ult по СП 63 п. 8.2.14 и п. 8.1.30:
' для двузначной эпюры eps_bt,ult = eps_bt2, для однозначно растянутой
' эпюры с ненулевой кривизной применяется формула (8.54). Fallback на
' LoadMultiplier здесь не выполняется: его включает общий диспетчер стратегии.
Private Function CalculateLambdaCrc(ByVal section As CSectionModel, ByVal concreteMaterial As CMaterialDiagram, _
        ByVal steelMaterial As CMaterialDiagram, ByVal startSolver As CSectionSolver, _
        ByVal targetN As Double, ByVal targetMomentMx As Double, ByVal targetMomentMy As Double, _
        ByVal referenceX As Double, ByVal referenceY As Double) As Boolean
    On Error GoTo Failed

    If mConcreteEbt0 <= 0# Or mConcreteEbt2 <= 0# Then
        mStopReason = "InputError: для проверки образования трещины нужны положительные eps_bt0 и eps_bt2 в модели CrackInitiation."
        Exit Function
    End If
    If targetMomentMx * targetMomentMx + targetMomentMy * targetMomentMy <= NEAR_ZERO_STRAIN Then
        mStopReason = "InputError: для пути λ*Mxy нужен ненулевой моментный вектор относительно расчетного центра."
        Exit Function
    End If

    Dim eps0 As Double
    Dim kx As Double
    Dim ky As Double
    InitializeCrcStrainGuess section, startSolver, targetN, targetMomentMx, targetMomentMy, mConcreteEbt2, eps0, kx, ky

    Dim iter As Long
    Dim rForce As Double
    Dim rMomentDirection As Double
    Dim rLimit As Double
    Dim lambdaValue As Double
    Dim solver As CSectionSolver
    Dim directStopReason As String

    ' Метод ищет саму плоскость деформаций образования трещины:
    ' N сохраняется, направление Mx/My относительно расчетного центра
    ' сохраняется, а максимальная растягивающая деформация бетона достигает
    ' eps_bt,ult по СП 63 п. 8.1.30.
    For iter = 1 To mSolverMaxIterations
        If Not EvaluateCrcStrainResidual(section, concreteMaterial, steelMaterial, targetN, _
                targetMomentMx, targetMomentMy, referenceX, referenceY, _
                eps0, kx, ky, solver, rForce, rMomentDirection, rLimit, lambdaValue) Then Exit Function

        If Abs(rForce) <= CRACK_EQUILIBRIUM_TOLERANCE And _
                Abs(rMomentDirection) <= CRACK_DIRECTION_TOLERANCE And _
                Abs(rLimit) <= CRACK_STRAIN_RATIO_TOLERANCE Then
            If lambdaValue < 0# Then
                mStopReason = "NumericalFailure: найденное состояние образования трещины имеет отрицательный lambda_crc."
                Exit Function
            End If
            If TryMarkConstantPartFormationFallback(section, concreteMaterial, steelMaterial, _
                    MomentFormationLoadPath(targetN, targetMomentMx, targetMomentMy, referenceX, referenceY), _
                    lambdaValue) Then
                CalculateLambdaCrc = True
                Exit Function
            End If
            mLambdaCrc = lambdaValue
            StoreMomentFormationLoad targetN, targetMomentMx, targetMomentMy, referenceX, referenceY
            Set mBeforeMcrcSolver = solver
            Set mBeforeMcrcState = New CSectionStateResult
            mBeforeMcrcState.InitializeFromSolver sstBeforeMcrcState, cpMcrc, _
                mCrackInitiationSpec, solver, False, "OK"
            AppendDiagnostic "lambda_crc=" & FormatNumberInvariant(mLambdaCrc) & _
                "; iterations=" & CStr(iter) & "; method=DirectCrackingStrain"
            AppendDiagnostic "UltimateStrain успешно нашел точку образования трещины по направлению момента."
            CalculateLambdaCrc = True
            Exit Function
        End If

        If Not ApplyCrcNewtonStep(section, concreteMaterial, steelMaterial, targetN, _
                targetMomentMx, targetMomentMy, referenceX, referenceY, _
                eps0, kx, ky, rForce, rMomentDirection, rLimit) Then
            directStopReason = mStopReason
            Exit For
        End If
    Next iter

    If Len(directStopReason) = 0 Then _
        directStopReason = "NumericalFailure: прямой поиск lambda_crc по eps_bt,ult не сошелся за заданное число итераций."
    mStopReason = directStopReason
    Exit Function

Failed:
    mStopReason = "NumericalFailure: расчет lambda_crc не выполнен: " & Err.Description
End Function

' Поиск lambda_crc методом LoadMultiplier по одному параметру выбранной траектории
' SLS.Crack.InitiationLoadPath. Для каждого lambda общий CStateSolutionRunner
' решает НДС по модели CrackInitiation. Метод не выполняет прямой Newton:
' решение о fallback принимает FindGeneralFormationLambda.
Private Function CalculateLambdaCrcByLoadScaling(ByVal section As CSectionModel, _
        ByVal concreteMaterial As CMaterialDiagram, ByVal steelMaterial As CMaterialDiagram, _
        ByVal startSolver As CSectionSolver, ByVal targetN As Double, _
        ByVal targetMomentMx As Double, ByVal targetMomentMy As Double, _
        ByVal referenceX As Double, ByVal referenceY As Double) As Boolean
    On Error GoTo Failed
    AppendDiagnostic "LoadMultiplier: начинаем подбор lambda_crc пробными расчетами НДС; путь " & _
        CrackFormationPathDisplayName() & "."

    Dim lowerLambda As Double
    Dim upperLambda As Double
    Dim lowerSolver As CSectionSolver
    Dim upperSolver As CSectionSolver
    Dim lowerResidual As Double
    Dim upperResidual As Double
    lowerLambda = 0#
    upperLambda = 1#

    Dim loadPath As CLoadPathVector
    Set loadPath = FormationLoadPath(targetN, targetMomentMx, targetMomentMy, referenceX, referenceY)

    Dim hasScaledLoad As Boolean
    hasScaledLoad = FormationPathHasScaledLoad(targetN, targetMomentMx, targetMomentMy)

    Dim probeN As Double
    Dim probeMx As Double
    Dim probeMy As Double
    loadPath.Target lowerLambda, probeN, probeMx, probeMy
    If IsZeroLoadProbe(probeN, probeMx, probeMy) Then
        ' Для lambda*NMxy нижняя точка - полное отсутствие нагрузки. Ее не
        ' нужно решать НДС solver-ом: трещины заведомо нет, а нулевая плоскость
        ' как warm-start только мешает активному старту следующих пробных точек.
        lowerResidual = -1#
        Set lowerSolver = Nothing
    Else
        Set lowerSolver = SolveState(section, concreteMaterial, steelMaterial, Nothing, _
            probeN, probeMx, probeMy)
        If Not SolverFormationResidual(section, lowerSolver, lowerResidual) Then
            mStopReason = "NumericalFailure: fallback lambda_crc не получил НДС при lambda = 0."
            Exit Function
        End If
    End If
    If StrComp(mCrackFormationPath, CRACK_PATH_N, vbTextCompare) = 0 And targetN <= NEAR_ZERO_FORCE Then
        If lowerResidual >= 0# Then
            MarkFormationPathFallback "Трещина уже образуется от постоянного момента выбранного пути " & _
                CrackFormationPathDisplayName() & "; сжимающая или нулевая N не масштабируется для поиска Ncrc, psi_s остается равным 1."
            CalculateLambdaCrcByLoadScaling = True
        Else
            FinishNoCrack "Трещина не образована от постоянного момента, а выбранный путь " & _
                CrackFormationPathDisplayName() & " не масштабирует растягивающую продольную силу."
        End If
        Exit Function
    End If

    If lowerResidual >= 0# Then
        MarkFormationPathFallback "Трещина уже образуется в постоянной части выбранного пути " & _
            CrackFormationPathDisplayName() & "; точка sigma_s,crc по lambda не определяется, psi_s остается равным 1."
        CalculateLambdaCrcByLoadScaling = True
        Exit Function
    End If

    If Not hasScaledLoad Then
        FinishNoCrack "Трещина не образована в постоянной части, а выбранный путь " & _
            CrackFormationPathDisplayName() & " не содержит ненулевой масштабируемой нагрузки."
        Exit Function
    End If

    If Not FindLoadScalingFormationBracket(section, concreteMaterial, steelMaterial, loadPath, _
            startSolver, lowerLambda, lowerSolver, lowerResidual, upperLambda, upperSolver, upperResidual) Then Exit Function

    If upperResidual < 0# Then
        FinishNoCrack "Трещина не образована: eps_bt,ult не достигнута при пропорциональном увеличении нагрузки до lambda = " & _
            FormatNumberInvariant(CRACK_LAMBDA_BRACKET_LIMIT) & "."
        Exit Function
    End If

    Dim iter As Long
    For iter = 1 To mSolverMaxIterations
        If upperLambda - lowerLambda <= CRACK_LAMBDA_BISECTION_TOLERANCE * MaxDouble(1#, upperLambda) Then Exit For

        Dim midLambda As Double
        Dim midSolver As CSectionSolver
        Dim midResidual As Double
        midLambda = 0.5 * (lowerLambda + upperLambda)
        Dim midStartSolver As CSectionSolver
        If lowerSolver Is Nothing Then
            Set midStartSolver = upperSolver
        Else
            Set midStartSolver = lowerSolver
        End If
        If Not ProbeLoadScalingFormation(section, concreteMaterial, steelMaterial, loadPath, _
                midStartSolver, midLambda, midSolver, midResidual) Then
            mStopReason = "NumericalFailure: fallback lambda_crc не получил НДС при lambda = " & _
                FormatNumberInvariant(midLambda) & "."
            Exit Function
        End If

        If midResidual >= 0# Then
            upperLambda = midLambda
            Set upperSolver = midSolver
            upperResidual = midResidual
        Else
            lowerLambda = midLambda
            Set lowerSolver = midSolver
        End If
    Next iter

    mLambdaCrc = upperLambda
    StoreFormationLoadFromPath loadPath, mLambdaCrc
    Set mBeforeMcrcSolver = upperSolver
    Set mBeforeMcrcState = New CSectionStateResult
    mBeforeMcrcState.InitializeFromSolver sstBeforeMcrcState, cpMcrc, _
        mCrackInitiationSpec, upperSolver, False, "OK"
    AppendDiagnostic "lambda_crc=" & FormatNumberInvariant(mLambdaCrc) & _
        "; iterations=" & CStr(iter) & "; method=LoadMultiplier"
    AppendDiagnostic "LoadMultiplier успешно нашел точку образования трещины."
    CalculateLambdaCrcByLoadScaling = True
    Exit Function

Failed:
    mStopReason = "NumericalFailure: fallback lambda_crc не выполнен: " & Err.Description
End Function

' Прямой Newton-поиск стадии трещинообразования на общей траектории
' Offset + lambda*Base. Это аналог идеи UltimateLoadPath из capacity:
' две невязки удерживают внутренние N/Mx/My на выбранной прямой нагрузки,
' третья доводит бетон до eps_bt,ult. Если этот путь не сошелся, вызывающий
' код оставляет прежний bisection-fallback по отдельным StateSolution.
Private Function CalculateLambdaCrcByLoadPathNewton(ByVal section As CSectionModel, _
        ByVal concreteMaterial As CMaterialDiagram, ByVal steelMaterial As CMaterialDiagram, _
        ByVal loadPath As CLoadPathVector, ByVal startSolver As CSectionSolver) As Boolean
    On Error GoTo Failed

    Dim guessN As Double
    Dim guessMx As Double
    Dim guessMy As Double
    loadPath.Target 1#, guessN, guessMx, guessMy

    Dim eps0 As Double
    Dim kx As Double
    Dim ky As Double
    InitializeCrcLoadPathGuess section, concreteMaterial, steelMaterial, startSolver, _
        guessN, guessMx, guessMy, eps0, kx, ky

    Dim iter As Long
    Dim rPath1 As Double
    Dim rPath2 As Double
    Dim rLimit As Double
    Dim lambdaValue As Double
    Dim solver As CSectionSolver
    Dim directStopReason As String

    For iter = 1 To mSolverMaxIterations
        If Not EvaluateCrcLoadPathResidual(section, concreteMaterial, steelMaterial, loadPath, _
                eps0, kx, ky, solver, rPath1, rPath2, rLimit, lambdaValue) Then Exit Function

        If Abs(rPath1) <= CRACK_EQUILIBRIUM_TOLERANCE And _
                Abs(rPath2) <= CRACK_EQUILIBRIUM_TOLERANCE And _
                Abs(rLimit) <= CRACK_STRAIN_RATIO_TOLERANCE Then
            If lambdaValue < 0# Then
                mStopReason = "NumericalFailure: найденное состояние образования трещины имеет отрицательный lambda_crc."
                Exit Function
            End If
            If TryMarkConstantPartFormationFallback(section, concreteMaterial, steelMaterial, _
                    loadPath, lambdaValue) Then
                CalculateLambdaCrcByLoadPathNewton = True
                Exit Function
            End If
            mLambdaCrc = lambdaValue
            StoreFormationLoadFromPath loadPath, mLambdaCrc
            Set mBeforeMcrcSolver = solver
            Set mBeforeMcrcState = New CSectionStateResult
            mBeforeMcrcState.InitializeFromSolver sstBeforeMcrcState, cpMcrc, _
                mCrackInitiationSpec, solver, False, "OK"
            AppendDiagnostic "lambda_crc=" & FormatNumberInvariant(mLambdaCrc) & _
                "; iterations=" & CStr(iter) & "; method=LoadPathCrackingStrain"
            AppendDiagnostic "UltimateStrain успешно нашел точку образования трещины по общей lambda-траектории."
            CalculateLambdaCrcByLoadPathNewton = True
            Exit Function
        End If

        If Not ApplyCrcLoadPathNewtonStep(section, concreteMaterial, steelMaterial, loadPath, _
                eps0, kx, ky, rPath1, rPath2, rLimit) Then
            directStopReason = mStopReason
            Exit For
        End If
    Next iter

    If Len(directStopReason) = 0 Then _
        directStopReason = "NumericalFailure: прямой поиск lambda_crc по общей load-path траектории не сошелся за заданное число итераций."
    mStopReason = directStopReason
    Exit Function

Failed:
    mStopReason = "NumericalFailure: прямой поиск lambda_crc по load-path не выполнен: " & Err.Description
End Function

' Превращает нулевую точку прямого UltimateStrain-поиска в тот же fallback,
' который уже использует LoadMultiplier. Если трещина появляется при lambda=0,
' выбранный путь не управляет причиной трещины: для lambda*Mxy это трещина от
' постоянной N, для lambda*N - от постоянного Mxy. В Auto это не окончательный
' успех, а сигнал попробовать следующий путь.
Private Function TryMarkConstantPartFormationFallback(ByVal section As CSectionModel, _
        ByVal concreteMaterial As CMaterialDiagram, ByVal steelMaterial As CMaterialDiagram, _
        ByVal loadPath As CLoadPathVector, ByVal lambdaValue As Double) As Boolean
    If lambdaValue > CRACK_LAMBDA_BISECTION_TOLERANCE Then Exit Function
    If StrComp(mCrackFormationPath, CRACK_PATH_NMXY, vbTextCompare) = 0 Then Exit Function
    If loadPath Is Nothing Then Exit Function

    Dim probeN As Double
    Dim probeMx As Double
    Dim probeMy As Double
    loadPath.Target 0#, probeN, probeMx, probeMy
    If IsZeroLoadProbe(probeN, probeMx, probeMy) Then Exit Function

    Dim probeSolver As CSectionSolver
    Set probeSolver = SolveState(section, concreteMaterial, steelMaterial, Nothing, _
        probeN, probeMx, probeMy)
    If probeSolver Is Nothing Then Exit Function
    If Not probeSolver.Converged Then Exit Function

    Dim residual As Double
    If Not SolverFormationResidual(section, probeSolver, residual) Then Exit Function
    If residual < 0# Then Exit Function

    MarkFormationPathFallback "Трещина уже образуется в постоянной части выбранного пути " & _
        CrackFormationPathDisplayName() & "; точка sigma_s,crc по этому пути не определяется, psi_s остается равным 1."
    TryMarkConstantPartFormationFallback = True
End Function

' Оценивает невязки для прямого поиска lambda_crc на общей линии нагрузки.
Private Function EvaluateCrcLoadPathResidual(ByVal section As CSectionModel, _
        ByVal concreteMaterial As CMaterialDiagram, ByVal steelMaterial As CMaterialDiagram, _
        ByVal loadPath As CLoadPathVector, ByVal eps0 As Double, ByVal kx As Double, ByVal ky As Double, _
        ByRef solver As CSectionSolver, ByRef rPath1 As Double, ByRef rPath2 As Double, _
        ByRef rLimit As Double, ByRef lambdaValue As Double) As Boolean
    On Error GoTo Failed

    Set solver = New CSectionSolver
    mSolverCallCount = mSolverCallCount + 1
    solver.EvaluateStrainPlane section, concreteMaterial, steelMaterial, eps0, kx, ky

    Dim loadPathMath As CLoadPathMath
    Set loadPathMath = NewLoadPathMath()
    lambdaValue = loadPathMath.Lambda(solver.Nint, solver.Mxint, solver.Myint, loadPath)
    loadPathMath.BuildResiduals solver.Nint, solver.Mxint, solver.Myint, _
        loadPath, lambdaValue, rPath1, rPath2

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

    AppendDiagnostic "lambda_crc_path_iter=" & FormatNumberInvariant(lambdaValue) & _
        "; rPath1=" & FormatNumberInvariant(rPath1) & _
        "; rPath2=" & FormatNumberInvariant(rPath2) & _
        "; rLimit=" & FormatNumberInvariant(rLimit) & _
        "; eps_bt_ult=" & FormatNumberInvariant(epsBtUlt)
    EvaluateCrcLoadPathResidual = True
    Exit Function

Failed:
    mStopReason = "NumericalFailure: оценка lambda_crc по load-path не выполнена: " & Err.Description
End Function

' Выполняет один Newton-шаг для прямого поиска lambda_crc на общей
' траектории нагрузки.
Private Function ApplyCrcLoadPathNewtonStep(ByVal section As CSectionModel, _
        ByVal concreteMaterial As CMaterialDiagram, ByVal steelMaterial As CMaterialDiagram, _
        ByVal loadPath As CLoadPathVector, ByRef eps0 As Double, ByRef kx As Double, ByRef ky As Double, _
        ByVal rPath1 As Double, ByVal rPath2 As Double, ByVal rLimit As Double) As Boolean
    Dim hEps As Double
    Dim hKx As Double
    Dim hKy As Double
    hEps = MaxDouble(0.0000001, Abs(eps0) * 0.0001)
    hKx = MaxDouble(0.000000001, Abs(kx) * 0.0001)
    hKy = MaxDouble(0.000000001, Abs(ky) * 0.0001)

    Dim j11 As Double, j21 As Double, j31 As Double
    Dim j12 As Double, j22 As Double, j32 As Double
    Dim j13 As Double, j23 As Double, j33 As Double
    BuildCrcLoadPathJacobianColumn section, concreteMaterial, steelMaterial, loadPath, _
        eps0 + hEps, kx, ky, hEps, rPath1, rPath2, rLimit, j11, j21, j31
    BuildCrcLoadPathJacobianColumn section, concreteMaterial, steelMaterial, loadPath, _
        eps0, kx + hKx, ky, hKx, rPath1, rPath2, rLimit, j12, j22, j32
    BuildCrcLoadPathJacobianColumn section, concreteMaterial, steelMaterial, loadPath, _
        eps0, kx, ky + hKy, hKy, rPath1, rPath2, rLimit, j13, j23, j33

    Dim system As CLinearSystem3x3
    Set system = New CLinearSystem3x3
    If Not system.Solve(j11, j12, j13, j21, j22, j23, j31, j32, j33, _
            -rPath1, -rPath2, -rLimit) Then
        mStopReason = "NumericalFailure: матрица load-path поиска lambda_crc вырождена: " & system.StopReason
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
    baseNorm = CrcLoadPathResidualNorm(rPath1, rPath2, rLimit)
    Dim alpha As Double
    alpha = mSolverDampingInitial
    If alpha <= 0# Or alpha > 1# Then alpha = 1#
    Do While alpha >= mSolverMinLineSearchAlpha
        Dim trialSolver As CSectionSolver
        Dim trialRPath1 As Double, trialRPath2 As Double, trialRLimit As Double, trialLambda As Double
        If EvaluateCrcLoadPathResidual(section, concreteMaterial, steelMaterial, loadPath, _
                eps0 + alpha * dEps, kx + alpha * dKx, ky + alpha * dKy, _
                trialSolver, trialRPath1, trialRPath2, trialRLimit, trialLambda) Then
            If CrcLoadPathResidualNorm(trialRPath1, trialRPath2, trialRLimit) < baseNorm Then
                eps0 = eps0 + alpha * dEps
                kx = kx + alpha * dKx
                ky = ky + alpha * dKy
                ApplyCrcLoadPathNewtonStep = True
                Exit Function
            End If
        End If
        If Not mSolverLineSearchEnabled Then Exit Do
        alpha = alpha / 2#
    Loop

    mStopReason = "NumericalFailure: line search load-path поиска lambda_crc не улучшил невязку."
End Function

' Старт для прямого поиска по общей lambda-траектории. Сначала используем
' централизованный CStateGuessBuilder, как в опасных DirectState/Capacity
' случаях; если он не смог построить плоскость, остаемся на старой грубой
' оценке по направлению нагрузки.
Private Sub InitializeCrcLoadPathGuess(ByVal section As CSectionModel, _
        ByVal concreteMaterial As CMaterialDiagram, ByVal steelMaterial As CMaterialDiagram, _
        ByVal startSolver As CSectionSolver, ByVal targetN As Double, _
        ByVal targetMx As Double, ByVal targetMy As Double, _
        ByRef eps0 As Double, ByRef kx As Double, ByRef ky As Double)
    Dim guessBuilder As CStateGuessBuilder
    Set guessBuilder = New CStateGuessBuilder
    If guessBuilder.BuildForDirectState(section, concreteMaterial, steelMaterial, _
            targetN, targetMx, targetMy, eps0, kx, ky) Then
        Exit Sub
    End If

    InitializeCrcStrainGuess section, startSolver, targetN, targetMx, targetMy, _
        mConcreteEbt2, eps0, kx, ky
End Sub

' Норма невязки для прямого load-path поиска. В отличие от старой
' rMomentDirection здесь rPath1/rPath2 уже выражены в долях solver tolerance,
' поэтому дополнительный множитель 100 не нужен.
Private Function CrcLoadPathResidualNorm(ByVal rPath1 As Double, _
        ByVal rPath2 As Double, ByVal rLimit As Double) As Double
    Dim loadPathMath As CLoadPathMath
    Set loadPathMath = NewLoadPathMath()
    CrcLoadPathResidualNorm = loadPathMath.ResidualNorm(rPath1, rPath2, rLimit)
End Function

' Один столбец численного Якобиана для load-path системы образования трещины.
Private Sub BuildCrcLoadPathJacobianColumn(ByVal section As CSectionModel, _
        ByVal concreteMaterial As CMaterialDiagram, ByVal steelMaterial As CMaterialDiagram, _
        ByVal loadPath As CLoadPathVector, ByVal eps0 As Double, ByVal kx As Double, ByVal ky As Double, _
        ByVal stepValue As Double, ByVal rPath1 As Double, ByVal rPath2 As Double, ByVal rLimit As Double, _
        ByRef j1 As Double, ByRef j2 As Double, ByRef j3 As Double)
    Dim solver As CSectionSolver
    Dim pPath1 As Double, pPath2 As Double, pLimit As Double, pLambda As Double
    If Not EvaluateCrcLoadPathResidual(section, concreteMaterial, steelMaterial, loadPath, _
            eps0, kx, ky, solver, pPath1, pPath2, pLimit, pLambda) Then Exit Sub
    j1 = (pPath1 - rPath1) / stepValue
    j2 = (pPath2 - rPath2) / stepValue
    j3 = (pLimit - rLimit) / stepValue
End Sub

' Создает общий helper lambda-траектории с текущими допусками crack solver-а.
Private Function NewLoadPathMath() As CLoadPathMath
    Dim loadPathMath As CLoadPathMath
    Set loadPathMath = New CLoadPathMath
    loadPathMath.Configure mSolverToleranceN, mSolverToleranceMx, mSolverToleranceMy, NEAR_ZERO_FORCE
    Set NewLoadPathMath = loadPathMath
End Function

' Проверяет, что пробная точка lambda действительно не содержит усилий.
' Такой случай возникает у lambda*NMxy при lambda=0: это не расчетное НДС, а
' только нижняя граница bracket-а "трещины еще нет".
Private Function IsZeroLoadProbe(ByVal nValue As Double, ByVal mxValue As Double, _
        ByVal myValue As Double) As Boolean
    IsZeroLoadProbe = (Abs(nValue) <= NEAR_ZERO_FORCE And _
        Abs(mxValue) <= MaxDouble(mSolverToleranceMx, CRACK_DIRECTION_TOLERANCE) And _
        Abs(myValue) <= MaxDouble(mSolverToleranceMy, CRACK_DIRECTION_TOLERANCE))
End Function

' Ищет верхнюю точку bracket-а для fallback-пути lambda*N/lambda*NMxy.
' Если полный уровень lambda=1 уже не решается из-за чрезмерной нагрузки,
' пробуем найти стадию образования трещины снизу, не требуя сходимости
' заведомо перегруженного НДС.
Private Function FindLoadScalingFormationBracket(ByVal section As CSectionModel, _
        ByVal concreteMaterial As CMaterialDiagram, ByVal steelMaterial As CMaterialDiagram, _
        ByVal loadPath As CLoadPathVector, ByVal startSolver As CSectionSolver, _
        ByRef lowerLambda As Double, ByRef lowerSolver As CSectionSolver, ByRef lowerResidual As Double, _
        ByRef upperLambda As Double, ByRef upperSolver As CSectionSolver, ByRef upperResidual As Double) As Boolean
    upperLambda = 1#
    If Not startSolver Is Nothing Then
        Set upperSolver = startSolver
        If SolverFormationResidual(section, upperSolver, upperResidual) Then GoTo UpperSolved
    ElseIf ProbeLoadScalingFormation(section, concreteMaterial, steelMaterial, loadPath, _
            lowerSolver, upperLambda, upperSolver, upperResidual) Then
        GoTo UpperSolved
    End If

    Dim failedLambda As Double
    failedLambda = upperLambda
    Dim iter As Long
    For iter = 1 To mSolverMaxIterations
        If failedLambda - lowerLambda <= CRACK_LAMBDA_BISECTION_TOLERANCE * MaxDouble(1#, failedLambda) Then Exit For

        Dim midLambda As Double
        Dim midSolver As CSectionSolver
        Dim midResidual As Double
        midLambda = 0.5 * (lowerLambda + failedLambda)
        If ProbeLoadScalingFormation(section, concreteMaterial, steelMaterial, loadPath, _
                lowerSolver, midLambda, midSolver, midResidual) Then
            If midResidual >= 0# Then
                upperLambda = midLambda
                Set upperSolver = midSolver
                upperResidual = midResidual
                FindLoadScalingFormationBracket = True
                Exit Function
            End If
            lowerLambda = midLambda
            Set lowerSolver = midSolver
            lowerResidual = midResidual
        Else
            failedLambda = midLambda
        End If
    Next iter

    ' Если выше найденной границы НДС уже не сходится из-за выхода за
    ' физическую область CrackInitiation-диаграммы, принимаем последнюю
    ' сходящуюся точку снизу как состояние непосредственно перед трещиной.
    ' Это сохраняет смысл LoadMultiplier для lambda*N/lambda*NMxy: нам не
    ' нужно решать заведомо перегруженное состояние, достаточно подойти к
    ' eps_bt,ult с рабочей стороны диаграммы.
    If Not lowerSolver Is Nothing And lowerLambda > 0# Then
        upperLambda = lowerLambda
        Set upperSolver = lowerSolver
        upperResidual = lowerResidual
        AppendDiagnostic "lambda_crc bracket принят по последней сходящейся точке перед несходящейся верхней границей; residual=" & _
            FormatNumberInvariant(lowerResidual)
        FindLoadScalingFormationBracket = True
        Exit Function
    End If

    mStopReason = "NumericalFailure: fallback lambda_crc не нашел сходящееся НДС с eps_bt,ult до перегруженной точки lambda = 1."
    Exit Function

UpperSolved:
    Do While upperResidual < 0# And upperLambda < CRACK_LAMBDA_BRACKET_LIMIT
        lowerLambda = upperLambda
        Set lowerSolver = upperSolver
        lowerResidual = upperResidual
        upperLambda = MinDouble(upperLambda * 2#, CRACK_LAMBDA_BRACKET_LIMIT)
        If Not ProbeLoadScalingFormation(section, concreteMaterial, steelMaterial, loadPath, _
                lowerSolver, upperLambda, upperSolver, upperResidual) Then
            mStopReason = "NumericalFailure: fallback lambda_crc не получил НДС при lambda = " & _
                FormatNumberInvariant(upperLambda) & "."
            Exit Function
        End If
    Loop
    FindLoadScalingFormationBracket = True
End Function

' Решает одну пробную точку lambda выбранного пути и возвращает residual
' образования трещины. startSolver используется только как стартовая плоскость.
Private Function ProbeLoadScalingFormation(ByVal section As CSectionModel, _
        ByVal concreteMaterial As CMaterialDiagram, ByVal steelMaterial As CMaterialDiagram, _
        ByVal loadPath As CLoadPathVector, ByVal startSolver As CSectionSolver, _
        ByVal lambdaValue As Double, ByRef solver As CSectionSolver, ByRef residual As Double) As Boolean
    Dim probeN As Double
    Dim probeMx As Double
    Dim probeMy As Double
    loadPath.Target lambdaValue, probeN, probeMx, probeMy

    ' Нулевая нижняя точка lambda=0 не является полезной стартовой плоскостью:
    ' если передать ее как warm-start, общий CStateSolutionRunner не построит
    ' свое active-set приближение для сложного N+M. Ненулевые соседние точки
    ' оставляем как обычный warm-start, потому что они ускоряют бисекцию.
    Dim effectiveStart As CSectionSolver
    Set effectiveStart = EffectiveLoadScalingWarmStart(startSolver)
    If effectiveStart Is Nothing Then
        Set effectiveStart = LoadScalingProbeGuess(section, concreteMaterial, steelMaterial, _
            probeN, probeMx, probeMy)
    End If
    Set solver = SolveState(section, concreteMaterial, steelMaterial, effectiveStart, _
        probeN, probeMx, probeMy, True)
    If SolverFormationResidual(section, solver, residual) Then
        AppendLoadScalingProbeDiagnostic lambdaValue, residual, True
        ProbeLoadScalingFormation = True
        Exit Function
    End If

    ' Если явная стартовая плоскость оказалась неудачной, даем общему runner-у
    ' второй шанс с его собственным active-set стартом. Это особенно важно для
    ' lambda*NMxy: рядом с моментом трещинообразования плохой warm-start может
    ' быть хуже, чем централизованный подбор начальной плоскости.
    If Not effectiveStart Is Nothing Then
        Set solver = SolveState(section, concreteMaterial, steelMaterial, Nothing, _
            probeN, probeMx, probeMy, True)
        If SolverFormationResidual(section, solver, residual) Then
            AppendLoadScalingProbeDiagnostic lambdaValue, residual, True
            ProbeLoadScalingFormation = True
            Exit Function
        End If
    End If

    ' Если warm-start от соседней lambda-точки оказался неудачным, пробуем
    ' стартовую плоскость от самой probe-нагрузки. Это тот же общий механизм,
    ' которым пользуются опасные траектории DirectState/Capacity: меняется
    ' только начальная плоскость Newton, а не сама траектория lambda.
    If Not startSolver Is Nothing Then
        Dim probeGuess As CSectionSolver
        Set probeGuess = LoadScalingProbeGuess(section, concreteMaterial, steelMaterial, _
            probeN, probeMx, probeMy)
        If Not probeGuess Is Nothing Then
            Set solver = SolveState(section, concreteMaterial, steelMaterial, probeGuess, _
                probeN, probeMx, probeMy, True)
            ProbeLoadScalingFormation = SolverFormationResidual(section, solver, residual)
            If ProbeLoadScalingFormation Then _
                AppendLoadScalingProbeDiagnostic lambdaValue, residual, True
        End If
    End If
    If Not ProbeLoadScalingFormation Then _
        AppendLoadScalingProbeDiagnostic lambdaValue, 0#, False
End Function

' Записывает одну пробную точку fallback-метода LoadMultiplier. В отличие от
' строк UltimateStrain, здесь lambda проверяется обычным НДС CrackInitiation,
' поэтому отдельная подпись метода важна для читаемой хронологии отчета.
Private Sub AppendLoadScalingProbeDiagnostic(ByVal lambdaValue As Double, _
        ByVal residual As Double, ByVal converged As Boolean)
    If converged Then
        AppendDiagnostic "lambda_crc_load_probe=" & FormatNumberInvariant(lambdaValue) & _
            "; residual=" & FormatNumberInvariant(residual) & _
            "; result=converged"
    Else
        AppendDiagnostic "lambda_crc_load_probe=" & FormatNumberInvariant(lambdaValue) & _
            "; result=failed"
    End If
End Sub

' Возвращает warm-start для probe solve только если в нем действительно есть
' ненулевая плоскость деформаций. Это защищает lambda*NMxy от плохого старта
' из полностью нулевого НДС при lambda=0.
Private Function EffectiveLoadScalingWarmStart(ByVal startSolver As CSectionSolver) As CSectionSolver
    If startSolver Is Nothing Then Exit Function
    If Not startSolver.Converged Then Exit Function
    If Abs(startSolver.Epsilon0) <= NEAR_ZERO_STRAIN And _
            Abs(startSolver.KappaX) <= NEAR_ZERO_STRAIN And _
            Abs(startSolver.KappaY) <= NEAR_ZERO_STRAIN Then Exit Function
    Set EffectiveLoadScalingWarmStart = startSolver
End Function

' Строит активную стартовую плоскость для отдельной probe-точки lambda.
' Используется тем же CStateGuessBuilder, что и общий runner прямого НДС:
' это помогает lambda*N/lambda*NMxy начинать не с нулевой плоскости и не
' зависеть от плохой соседней точки bracket-а.
Private Function LoadScalingProbeGuess(ByVal section As CSectionModel, _
        ByVal concreteMaterial As CMaterialDiagram, ByVal steelMaterial As CMaterialDiagram, _
        ByVal probeN As Double, ByVal probeMx As Double, ByVal probeMy As Double) As CSectionSolver
    If Sqr(probeMx * probeMx + probeMy * probeMy) <= _
            MaxDouble(mSolverToleranceMx, mSolverToleranceMy) Then Exit Function

    Dim guessBuilder As CStateGuessBuilder
    Set guessBuilder = New CStateGuessBuilder

    Dim eps0 As Double
    Dim kx As Double
    Dim ky As Double
    If Not guessBuilder.BuildForDirectState(section, concreteMaterial, steelMaterial, _
            probeN, probeMx, probeMy, eps0, kx, ky) Then Exit Function

    Dim initialSolver As CSectionSolver
    Set initialSolver = New CSectionSolver
    initialSolver.SetInitialState eps0, kx, ky
    Set LoadScalingProbeGuess = initialSolver
End Function

' Возвращает безразмерную невязку образования трещины для уже решенного НДС:
' отрицательная - eps_bt,ult еще не достигнута, положительная - порог пройден.
Private Function SolverFormationResidual(ByVal section As CSectionModel, _
        ByVal solver As CSectionSolver, ByRef residual As Double) As Boolean
    If solver Is Nothing Then Exit Function
    If Not solver.Converged Then Exit Function

    Dim minConcreteStrain As Double
    Dim maxConcreteStrain As Double
    ConcreteStrainBounds solver, section, minConcreteStrain, maxConcreteStrain

    Dim epsBtUlt As Double
    epsBtUlt = ConcreteTensionUltimateStrain(minConcreteStrain, maxConcreteStrain)
    If epsBtUlt <= 0# Then Exit Function
    residual = maxConcreteStrain / epsBtUlt - 1#
    SolverFormationResidual = True
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
' первое держит заданную N, второе сохраняет направление Mx/My относительно
' расчетного центра, третье требует достижения eps_bt,ult в наиболее
' растянутом бетонном конечном элементе.
Private Function EvaluateCrcStrainResidual(ByVal section As CSectionModel, ByVal concreteMaterial As Object, _
        ByVal steelMaterial As Object, ByVal targetN As Double, _
        ByVal targetMomentMx As Double, ByVal targetMomentMy As Double, _
        ByVal referenceX As Double, ByVal referenceY As Double, _
        ByVal eps0 As Double, ByVal kx As Double, ByVal ky As Double, _
        ByRef solver As CSectionSolver, ByRef rForce As Double, ByRef rMomentDirection As Double, _
        ByRef rLimit As Double, ByRef lambdaValue As Double) As Boolean
    On Error GoTo Failed

    Set solver = New CSectionSolver
    mSolverCallCount = mSolverCallCount + 1
    solver.EvaluateStrainPlane section, concreteMaterial, steelMaterial, eps0, kx, ky

    Dim baseMx As Double
    Dim baseMy As Double
    Dim intMx As Double
    Dim intMy As Double
    baseMx = targetMomentMx / MaxDouble(mSolverToleranceMx, 1#)
    baseMy = targetMomentMy / MaxDouble(mSolverToleranceMy, 1#)

    Dim momentMxAtReference As Double
    Dim momentMyAtReference As Double
    momentMxAtReference = solver.Mxint - solver.Nint * referenceY
    momentMyAtReference = solver.Myint - solver.Nint * referenceX
    intMx = momentMxAtReference / MaxDouble(mSolverToleranceMx, 1#)
    intMy = momentMyAtReference / MaxDouble(mSolverToleranceMy, 1#)

    Dim baseMomentNorm2 As Double
    baseMomentNorm2 = baseMx * baseMx + baseMy * baseMy
    If baseMomentNorm2 <= NEAR_ZERO_STRAIN Then
        mStopReason = "InputError: для Auto нужен ненулевой моментный вектор Mx/My."
        Exit Function
    End If

    rForce = (solver.Nint - targetN) / MaxDouble(mSolverToleranceN, 1#)
    lambdaValue = (momentMxAtReference * targetMomentMx + momentMyAtReference * targetMomentMy) / _
        (targetMomentMx * targetMomentMx + targetMomentMy * targetMomentMy)
    SetCrcMomentDirectionResidual baseMx, baseMy, intMx, intMy, rMomentDirection

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
        "; rN=" & FormatNumberInvariant(rForce) & _
        "; rMomentDir=" & FormatNumberInvariant(rMomentDirection) & _
        "; rLimit=" & FormatNumberInvariant(rLimit) & _
        "; eps_bt_ult=" & FormatNumberInvariant(epsBtUlt)
    EvaluateCrcStrainResidual = True
    Exit Function

Failed:
    mStopReason = "NumericalFailure: оценка состояния образования трещины не выполнена: " & Err.Description
End Function

' Возвращает знаковую невязку направления момента. Она равна нулю, когда
' внутренний момент относительно расчетного центра параллелен заданному Mx/My.
Private Sub SetCrcMomentDirectionResidual(ByVal baseMx As Double, ByVal baseMy As Double, _
        ByVal intMx As Double, ByVal intMy As Double, ByRef rMomentDirection As Double)
    Dim denominator As Double
    denominator = MaxDouble(Abs(baseMx) + Abs(baseMy), 1#) ^ 2
    rMomentDirection = (intMy * baseMx - intMx * baseMy) / denominator
End Sub

' Выполняет один Newton-шаг для прямой системы образования трещины. Якобиан
' строится численно, как в UltimateStrain, но предельным условием является
' нормативная предельная растягивающая деформация бетона eps_bt,ult.
Private Function ApplyCrcNewtonStep(ByVal section As CSectionModel, ByVal concreteMaterial As Object, _
        ByVal steelMaterial As Object, ByVal targetN As Double, _
        ByVal targetMomentMx As Double, ByVal targetMomentMy As Double, _
        ByVal referenceX As Double, ByVal referenceY As Double, _
        ByRef eps0 As Double, ByRef kx As Double, ByRef ky As Double, _
        ByVal rForce As Double, ByVal rMomentDirection As Double, ByVal rLimit As Double) As Boolean
    Dim hEps As Double
    Dim hKx As Double
    Dim hKy As Double
    hEps = MaxDouble(0.0000001, Abs(eps0) * 0.0001)
    hKx = MaxDouble(0.000000001, Abs(kx) * 0.0001)
    hKy = MaxDouble(0.000000001, Abs(ky) * 0.0001)

    Dim j11 As Double, j21 As Double, j31 As Double
    Dim j12 As Double, j22 As Double, j32 As Double
    Dim j13 As Double, j23 As Double, j33 As Double
    BuildCrcJacobianColumn section, concreteMaterial, steelMaterial, targetN, _
        targetMomentMx, targetMomentMy, referenceX, referenceY, _
        eps0 + hEps, kx, ky, hEps, rForce, rMomentDirection, rLimit, j11, j21, j31
    BuildCrcJacobianColumn section, concreteMaterial, steelMaterial, targetN, _
        targetMomentMx, targetMomentMy, referenceX, referenceY, _
        eps0, kx + hKx, ky, hKx, rForce, rMomentDirection, rLimit, j12, j22, j32
    BuildCrcJacobianColumn section, concreteMaterial, steelMaterial, targetN, _
        targetMomentMx, targetMomentMy, referenceX, referenceY, _
        eps0, kx, ky + hKy, hKy, rForce, rMomentDirection, rLimit, j13, j23, j33

    Dim system As CLinearSystem3x3
    Set system = New CLinearSystem3x3
    If Not system.Solve(j11, j12, j13, j21, j22, j23, j31, j32, j33, -rForce, -rMomentDirection, -rLimit) Then
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
    baseNorm = CrcResidualNorm(rForce, rMomentDirection, rLimit)
    Dim alpha As Double
    alpha = mSolverDampingInitial
    If alpha <= 0# Or alpha > 1# Then alpha = 1#
    Do While alpha >= mSolverMinLineSearchAlpha
        Dim trialSolver As CSectionSolver
        Dim trialRForce As Double, trialRMoment As Double, trialRLimit As Double, trialLambda As Double
        If EvaluateCrcStrainResidual(section, concreteMaterial, steelMaterial, targetN, _
                targetMomentMx, targetMomentMy, referenceX, referenceY, _
                eps0 + alpha * dEps, kx + alpha * dKx, ky + alpha * dKy, _
                trialSolver, trialRForce, trialRMoment, trialRLimit, trialLambda) Then
            If CrcResidualNorm(trialRForce, trialRMoment, trialRLimit) < baseNorm Then
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
        ByVal steelMaterial As Object, ByVal targetN As Double, _
        ByVal targetMomentMx As Double, ByVal targetMomentMy As Double, _
        ByVal referenceX As Double, ByVal referenceY As Double, _
        ByVal eps0 As Double, ByVal kx As Double, ByVal ky As Double, _
        ByVal stepValue As Double, ByVal rForce As Double, ByVal rMomentDirection As Double, ByVal rLimit As Double, _
        ByRef j1 As Double, ByRef j2 As Double, ByRef j3 As Double)
    Dim solver As CSectionSolver
    Dim pForce As Double, pMoment As Double, pLimit As Double, pLambda As Double
    If Not EvaluateCrcStrainResidual(section, concreteMaterial, steelMaterial, targetN, _
            targetMomentMx, targetMomentMy, referenceX, referenceY, _
            eps0, kx, ky, solver, pForce, pMoment, pLimit, pLambda) Then Exit Sub
    j1 = (pForce - rForce) / stepValue
    j2 = (pMoment - rMomentDirection) / stepValue
    j3 = (pLimit - rLimit) / stepValue
End Sub

' Норма смешивает условие направления и условие предельной деформации только
' для line search. Она не имеет самостоятельного нормативного смысла.
Private Function CrcResidualNorm(ByVal rForce As Double, ByVal rMomentDirection As Double, ByVal rLimit As Double) As Double
    CrcResidualNorm = Abs(rForce) + Abs(rMomentDirection) * 100# + Abs(rLimit) * 10#
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
    mTensionRebarIds = SelectedRebarIds(section, selected)
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
' уточнение не запускается. Если уточнение нужно, lambda_crc и состояния
' Before/AfterMcrcState уже должны быть найдены обязательной предварительной
' проверкой образования трещины, поэтому повторный поиск точки образования
' здесь не выполняется.
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
        AppendDiagnostic "Auto psi_s: первая проверка с psi_s=1 проходит; sigma_s,crc не уточнялся."
        FinalizeWithPsiMode = True
        Exit Function
    End If

    If mFormationPathFallbackPsi1 Then
        AppendDiagnostic "Auto psi_s: выбранный путь образования трещины не дает отдельной точки sigma_s,crc; psi_s оставлен равным 1."
        FinalizeWithPsiMode = True
        Exit Function
    End If

    AppendDiagnostic "Auto psi_s: первая проверка с psi_s=1 не проходит; используется ранее найденный AfterMcrcState."
    mConverged = False

    If mLambdaCrc < 0# Then
        mStopReason = "InputError: lambda_crc предварительной проверки образования трещины не определена."
        Exit Function
    End If
    If Not EnsureAfterMcrcState(section, materialProvider, currentSolver, targetN, targetMx, targetMy) Then Exit Function

    mSigmaSCrc = AverageSelectedSteelStress(mAfterMcrcSolver, section, _
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

' Завершает проверку как "трещина есть", но без расчетной точки sigma_s,crc.
' Такой fallback нужен, когда выбранный путь lambda не управляет причиной,
' которая уже вызвала трещину: например lambda*Mxy при трещине от одной N или
' lambda*N при трещине от уже приложенного момента. В этих случаях Auto-режим
' не имеет честного состояния сразу после образования трещины по выбранному
' пути, поэтому дальше остается консервативное psi_s = 1.
Private Sub MarkFormationPathFallback(ByVal reason As String)
    mLambdaCrc = 0#
    mFormationPathFallbackPsi1 = True
    mConverged = True
    mCrackFormed = True
    mStopReason = reason
    AppendDiagnostic reason
End Sub

' Запускает служебное НДС через общий CStateSolutionRunner. Так состояния
' BeforeMcrc/AfterMcrc получают тот же централизованный подбор стартовой
' плоскости, настройки solver-а и диагностику, что и Strength/CrackedState.
Private Function SolveState(ByVal section As CSectionModel, ByVal concreteMaterial As CMaterialDiagram, _
        ByVal steelMaterial As CMaterialDiagram, ByVal startSolver As CSectionSolver, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double, _
        Optional ByVal allowRetryStarts As Boolean = False) As CSectionSolver
    Dim runner As CStateSolutionRunner
    Set runner = New CStateSolutionRunner
    ConfigureStateRunner runner
    If startSolver Is Nothing Then
        runner.Solve section, concreteMaterial, steelMaterial, targetN, targetMx, targetMy, allowRetryStarts
    Else
        runner.SolveWithInitialSolver section, concreteMaterial, steelMaterial, _
            targetN, targetMx, targetMy, startSolver, allowRetryStarts
    End If
    If runner.SolverCallCount > 0 Then
        mSolverCallCount = mSolverCallCount + runner.SolverCallCount
    Else
        mSolverCallCount = mSolverCallCount + 1
    End If
    If Len(runner.DiagnosticLog) > 0 Then AppendDiagnostic runner.DiagnosticLog
    Set SolveState = runner.ResultSolver
End Function

' Передает crack-настройки в общий runner прямого НДС.
Private Sub ConfigureStateRunner(ByVal runner As CStateSolutionRunner)
    runner.LoadSteps = mSolverLoadSteps
    runner.MaxIterations = mSolverMaxIterations
    runner.ToleranceN = mSolverToleranceN
    runner.ToleranceMx = mSolverToleranceMx
    runner.ToleranceMy = mSolverToleranceMy
    runner.LineSearchEnabled = mSolverLineSearchEnabled
    runner.DampingInitial = mSolverDampingInitial
    runner.MinLineSearchAlpha = mSolverMinLineSearchAlpha
    runner.SolverMethod = mSolverMethod
    runner.SecantMaxRestarts = mSolverSecantMaxRestarts
    runner.SecantMinStepNorm = mSolverSecantMinStepNorm
    runner.MaxDeltaEpsilon0 = mSolverMaxDeltaEpsilon0
    runner.MaxDeltaKappa = mSolverMaxDeltaKappa
    runner.DiagnosticsEnabled = mDiagnosticsEnabled
End Sub

' Центральное растяжение для трещин определяется по моментному вектору
' относительно центра тяжести приведенного сечения. Перенос от
' Load.ReferenceOffset и обнуление практически нулевых моментов выполняет
' CSectionLoadState; crack-калькулятор только спрашивает у него моменты
' относительно нужного расчетного центра и не повторяет геометрию приложения
' нагрузки.
'
' Кривизны solver-а здесь сознательно не используются. Если при таком LC
' появляется заметная кривизна, это уже ошибка в переносе нагрузок, сетке или
' тестовом сценарии; она должна ловиться регрессионными проверками, а не
' менять расчетную ветку в рабочей книге.
Private Function IsCentralTensionState(ByVal loadStateForClassification As CSectionLoadState, _
        ByVal centralReferenceX As Double, ByVal centralReferenceY As Double) As Boolean
    IsCentralTensionState = loadStateForClassification.IsCentralTensionAbout( _
        centralReferenceX, centralReferenceY, NEAR_ZERO_FORCE, _
        MaxDouble(mSolverToleranceMx, 1#), MaxDouble(mSolverToleranceMy, 1#))
End Function

' Проверяет, есть ли в LC растяжение или изгиб относительно расчетного центра.
' Если после централизованного переноса и ZeroMoment-фильтра нет ни
' растягивающей силы, ни момента, нормальная трещина не может образоваться и
' прямой поиск точки образования трещины запускать не нужно.
Private Function HasNormalCrackDrivingAction(ByVal loadStateForClassification As CSectionLoadState, _
        ByVal centralReferenceX As Double, ByVal centralReferenceY As Double) As Boolean
    If loadStateForClassification.IsTension(NEAR_ZERO_FORCE) Then
        HasNormalCrackDrivingAction = True
        Exit Function
    End If
    If Abs(loadStateForClassification.MxAboutPoint(centralReferenceY)) > MaxDouble(mSolverToleranceMx, 1#) Then
        HasNormalCrackDrivingAction = True
        Exit Function
    End If
    If Abs(loadStateForClassification.MyAboutPoint(centralReferenceX)) > MaxDouble(mSolverToleranceMy, 1#) Then
        HasNormalCrackDrivingAction = True
    End If
End Function

' Возвращает N в найденной общей N+M точке трещинообразования. Для пути
' lambda*Mxy это просто текущая N, а для lambda*N/lambda*NMxy - масштабированная
' продольная сила на пороге.
Private Function FormationNValue() As Double
    If Not mFormationLoadDefined Then Exit Function
    FormationNValue = mFormationLoadN
End Function

' Возвращает модуль моментного вектора в найденной общей N+M точке
' трещинообразования относительно расчетного центра трещин.
Private Function CrackMomentFromFormationLoad(ByVal referenceX As Double, ByVal referenceY As Double) As Double
    If Not mFormationLoadDefined Then Exit Function

    Dim mxRef As Double
    Dim myRef As Double
    mxRef = mFormationLoadMx - mFormationLoadN * referenceY
    myRef = mFormationLoadMy - mFormationLoadN * referenceX
    CrackMomentFromFormationLoad = Sqr(mxRef * mxRef + myRef * myRef)
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

' Перечисляет имена выбранных стержней для контрольной таблицы напряжений.
Private Function SelectedRebarIds(ByVal section As CSectionModel, ByRef selected() As Boolean) As String
    Dim result As String
    Dim i As Long
    For i = 1 To section.RebarCount
        If selected(i) Then
            If Len(result) > 0 Then result = result & ", "
            result = result & section.RebarID(i)
        End If
    Next i
    SelectedRebarIds = result
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

' Возвращает усилие образования трещины при истинном центральном растяжении.
' Ared считаем через общий калькулятор приведенного сечения, чтобы ветка
' трещин использовала ту же модель Ab + (Es/Eb - 1)As, что устойчивость,
' свойства Results и остальные проверки, работающие с приведенной геометрией.
Private Function CentralTensionCrackForce(ByVal section As CSectionModel) As Double
    If mConcreteEb <= 0# Or mSteelEs <= 0# Then Exit Function

    Dim props As CSectionPropertiesCalculator
    Set props = New CSectionPropertiesCalculator
    props.CalculateTransformedByModuli section, mConcreteEb, mSteelEs
    mAred = props.Area
    CentralTensionCrackForce = mAred * mConcreteRbtSer
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
' плоскости образования трещины. Эти две величины нужны, чтобы отличить
' двузначную эпюру от однозначно растянутой и выбрать нормативный eps_bt,ult
' по СП 63 п. 8.1.30.
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
Private Function ResolvePhi3(ByVal loadStateForClassification As CSectionLoadState) As Double
    If StrComp(mPhi3Mode, "User", vbTextCompare) = 0 Then
        ResolvePhi3 = mPhi3Input
    ElseIf IsTensionLoad(loadStateForClassification) Then
        ResolvePhi3 = PHI3_TENSION
    Else
        ResolvePhi3 = PHI3_BENDING
    End If
End Function

' Определяет растягивающий характер продольной силы только через единый
' объект нагрузки. Так central tension и phi3 используют ту же интерпретацию
' Load.ReferenceOffset и фильтра малых моментов, что batch/stability/capacity.
Private Function IsTensionLoad(ByVal loadStateForClassification As CSectionLoadState) As Boolean
    IsTensionLoad = loadStateForClassification.IsTension(NEAR_ZERO_FORCE)
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

    mCrackFormationPath = NormalizeCrackFormationPath(mCrackFormationPath)
    mCrackFormationSolutionStrategy = NormalizeCrackFormationSolutionStrategy(mCrackFormationSolutionStrategy)

    If mAllowableCrackWidth <= 0# Then Err.Raise vbObjectError + 3710, "CCrackWidthCalculator", "SLS.Crack.Allowable должен быть положительным."
    If mPhi1Input <= 0# Then Err.Raise vbObjectError + 3717, "CCrackWidthCalculator", "SLS.Crack.Phi1 должен быть положительным."
    If mPhi2Input <= 0# Then Err.Raise vbObjectError + 3718, "CCrackWidthCalculator", "SLS.Crack.Phi2 должен быть положительным."
    If mPhi3Input <= 0# Then Err.Raise vbObjectError + 3719, "CCrackWidthCalculator", "SLS.Crack.Phi3 должен быть положительным."
    If mPsiSInput <= 0# Then Err.Raise vbObjectError + 3720, "CCrackWidthCalculator", "SLS.Crack.PsiS должен быть положительным."
    If mConcreteRbtSer <= 0# Then Err.Raise vbObjectError + 3711, "CCrackWidthCalculator", "Concrete.Rbt.SLS должен быть положительным для расчета трещин."
    If mConcreteEb <= 0# Then Err.Raise vbObjectError + 3712, "CCrackWidthCalculator", "Concrete.Eb должен быть положительным для расчета трещин."
    If mSteelEs <= 0# Then Err.Raise vbObjectError + 3713, "CCrackWidthCalculator", "Steel.Es должен быть положительным для расчета трещин."
    If Not mCrackInitiationSpec Is Nothing Then
        If mConcreteEbt0 <= 0# Then Err.Raise vbObjectError + 3714, "CCrackWidthCalculator", "Concrete.Ebt0 должен быть положительным для проверки образования трещины."
        If mConcreteEbt2 <= 0# Then Err.Raise vbObjectError + 3715, "CCrackWidthCalculator", "Concrete.Ebt2 должен быть положительным для проверки образования трещины."
        If mConcreteEbt0 > mConcreteEbt2 Then Err.Raise vbObjectError + 3716, "CCrackWidthCalculator", "Concrete.Ebt0 не должен быть больше Concrete.Ebt2 для формулы (8.54)."
    End If
End Sub

' Приводит пользовательский выбор пути трещинообразования к внутреннему ключу.
' Поддерживаются те же обозначения lambda, что и в CapacityLoadPath, но без
' одноосных Mx/My: для трещин пользователь выбирает Mxy, N или весь N+Mxy.
Private Function NormalizeCrackFormationPath(ByVal rawPath As String) As String
    Dim value As String
    value = LCase$(Trim$(rawPath))
    value = Replace$(value, " ", vbNullString)
    value = Replace$(value, ChrW$(&H3BB), "lambda")

    Select Case value
        Case "auto"
            NormalizeCrackFormationPath = CRACK_PATH_AUTO
        Case vbNullString, "lambda*mxy", "lambdamxy", "mxy"
            NormalizeCrackFormationPath = CRACK_PATH_MXY
        Case "lambda*n", "lambdan", "n"
            NormalizeCrackFormationPath = CRACK_PATH_N
        Case "lambda*nmxy", "lambdanmxy", "nmxy", "all"
            NormalizeCrackFormationPath = CRACK_PATH_NMXY
        Case Else
            Err.Raise vbObjectError + 3734, "CCrackWidthCalculator", _
                "SLS.Crack.InitiationLoadPath должен быть Auto, lambda*Mxy, lambda*N или lambda*NMxy."
    End Select
End Function

' Приводит численную стратегию поиска трещинообразования к той же логике,
' что используется для capacity: Auto, UltimateStrain или LoadMultiplier.
Private Function NormalizeCrackFormationSolutionStrategy(ByVal rawStrategy As String) As String
    Select Case LCase$(Trim$(rawStrategy))
        Case vbNullString, "auto"
            NormalizeCrackFormationSolutionStrategy = STRATEGY_AUTO
        Case "ultimatestrain", "strain", "direct"
            NormalizeCrackFormationSolutionStrategy = STRATEGY_ULTIMATE_STRAIN
        Case "loadmultiplier", "multiplier", "lambda"
            NormalizeCrackFormationSolutionStrategy = STRATEGY_LOAD_MULTIPLIER
        Case Else
            Err.Raise vbObjectError + 3735, "CCrackWidthCalculator", _
                "SLS.Crack.InitiationSolutionStrategy должен быть Auto, UltimateStrain или LoadMultiplier."
    End Select
End Function

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
    mSolverCallCount = 0
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
    mFormationNcrc = 0#
    mMcrc = 0#
    mAred = 0#
    mTensionRebarIds = vbNullString
    Set mBeforeMcrcState = Nothing
    Set mBeforeMcrcSolver = Nothing
    Set mAfterMcrcState = Nothing
    Set mAfterMcrcSolver = Nothing
    Set mCrackInitiationSpec = Nothing
    Set mCrackedStateSpec = Nothing
    mSectionDepthH = 0#
    mCoverA = 0#
    mTensionDepth = 0#
    mEffectiveZoneDepth = 0#
    mCentralTensionBranch = False
    mFormationPrechecked = False
    mFormationLoadDefined = False
    mFormationLoadN = 0#
    mFormationLoadMx = 0#
    mFormationLoadMy = 0#
    mFormationPathFallbackPsi1 = False
    mResolvedCrackFormationPath = vbNullString
End Sub

' Сбрасывает только данные одной попытки поиска точки образования трещины. Диагностический
' журнал и выбранные материальные модели сохраняются, чтобы Auto мог показать
' всю цепочку переключения путей.
Private Sub ClearFormationAttemptState()
    mConverged = False
    mStopReason = vbNullString
    mCrackFormed = False
    mLambdaCrc = 0#
    mNcrc = 0#
    mFormationNcrc = 0#
    mMcrc = 0#
    Set mBeforeMcrcState = Nothing
    Set mBeforeMcrcSolver = Nothing
    Set mAfterMcrcState = Nothing
    Set mAfterMcrcSolver = Nothing
    mFormationLoadDefined = False
    mFormationLoadN = 0#
    mFormationLoadMx = 0#
    mFormationLoadMy = 0#
    mFormationPathFallbackPsi1 = False
    mResolvedCrackFormationPath = vbNullString
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

' Возвращает число внутренних запусков/оценок solver-а, выполненных именно
' этим расчетом трещин. Batch использует его для общей диагностической сводки,
' чтобы не терять служебные solve при поиске Ncrc/Mcrc.
Public Property Get SolverCallCount() As Long
    SolverCallCount = mSolverCallCount
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

Public Property Get CrackFormationPath() As String
    CrackFormationPath = CrackFormationPathDisplayName()
End Property

Public Property Get CrackFormationMethod() As String
    If Len(mResolvedCrackFormationPath) > 0 Then
        CrackFormationMethod = CrackFormationPathDisplayNameForKey(mResolvedCrackFormationPath)
    Else
        CrackFormationMethod = vbNullString
    End If
End Property

Public Property Get CrackFormationSolutionStrategy() As String
    CrackFormationSolutionStrategy = mCrackFormationSolutionStrategy
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

Public Property Get FormationNcrc() As Double
    FormationNcrc = mFormationNcrc
End Property

Public Property Get Mcrc() As Double
    Mcrc = mMcrc
End Property

Public Property Get Ared() As Double
    Ared = mAred
End Property

Public Property Get RbtSer() As Double
    RbtSer = mConcreteRbtSer
End Property

Public Property Get TensionRebarIds() As String
    TensionRebarIds = mTensionRebarIds
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

Public Property Let CrackFormationPath(ByVal value As String)
    mCrackFormationPath = value
End Property

Public Property Let CrackFormationSolutionStrategy(ByVal value As String)
    mCrackFormationSolutionStrategy = value
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
