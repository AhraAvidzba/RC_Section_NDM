VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCrackFormationCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Образование нормальной трещины
' ==========================================================================
' Владелец инженерной постановки трещинообразования: material spec, физический
' критерий бетона, результаты Pre/PostCrackState. Общий LoadPath и порядок
' Auto-попыток задает CLimitSearchCoordinator.
' Численные bracket/Newton выполняют общие CLimitSearch*-классы; обычное НДС
' решается общим State pipeline. Здесь нет зоны As/Abt, sigma_s, psi или формулы
' ширины. Width получает только готовый CCrackFormationResult.

Private Const CRACK_EQUILIBRIUM_TOLERANCE As Double = 1# ' Допуск равновесия N в долях пользовательского solver tolerance.
Private Const CRACK_DIRECTION_TOLERANCE As Double = 0.001 ' Допуск сохранения направления моментного вектора в прямом поиске lambda_crc.
Private Const CRACK_STRAIN_RATIO_TOLERANCE As Double = 0.000000001 ' Точность предельного критерия для приема State внутри физической границы диаграммы.
Private Const CRACK_LAMBDA_BRACKET_LIMIT As Double = 1024# ' Верхняя граница служебного поиска lambda_crc, если прямой Newton не сошелся.
Private Const CRACK_LAMBDA_BISECTION_TOLERANCE As Double = 0.000001 ' Относительная точность fallback-бисекции lambda_crc.
Private Const NEAR_ZERO_STRAIN As Double = 0.000000000001 ' Численный ноль для деформаций и кривизн.
Private Const NEAR_ZERO_FORCE As Double = 0.000001 ' Численный ноль для отличия центрального растяжения от нулевой нагрузки.
Private Const CRACK_PATH_AUTO As String = "Auto" ' Общий Search выбирает порядок по активным компонентам сочетания.
Private Const CRACK_PATH_N As String = "LambdaN" ' При поиске трещинообразования масштабируется только продольная сила N.
Private Const CRACK_PATH_NMXY As String = "LambdaNMxy" ' При поиске трещинообразования масштабируется весь вектор N/Mx/My.
Private Const STRATEGY_AUTO As String = "Auto" ' Сначала UltimateStrain, при численной неудаче LoadMultiplier.
Private Const STRATEGY_ULTIMATE_STRAIN As String = "UltimateStrain" ' Прямой поиск плоскости по eps_bt,ult без одномерной скобки.
Private Const STRATEGY_LOAD_MULTIPLIER As String = "LoadMultiplier" ' Одномерный поиск lambda через серию обычных solve.

Private mConverged As Boolean ' True, если расчет трещины завершился без ошибки входных данных или несходимости.
Private mStopReason As String ' Короткое пояснение результата: расчет выполнен, трещина не образована или solver не сошелся.
Private mDiagnosticLog As String ' Диагностика пути, критерия и состояний трещинообразования.
Private mDiagnosticsEnabled As Boolean ' Отключает подробный журнал, когда txt-отчет выполнения не нужен.
Private mSolverCallCount As Long ' Сколько внутренних CSectionSolver/CStateSolutionRunner-вызовов выполнил расчет трещин.
Private mSearchExecuted As Boolean ' Хотя бы одна численная попытка поиска точки в текущем LC; аналитический Ncrc не является Search.
Private mFormationInternalStatus As EResultInternalStatus ' Машинный статус этапа образования нормальной трещины.
Private mFormationResultCode As EResultCode ' Машинный код причины результата образования нормальной трещины.
Private mCrackFormed As Boolean ' True, если по выбранной ветке расчета трещина считается образованной.
Private mCrackFormationPath As String ' Auto либо LambdaMx/My/Mxy/N/NMxy: масштабируемая часть LC общего LoadPath.
Private mResolvedCrackFormationPath As String ' Фактический путь, который дал принятую точку образования трещины.
Private mCrackFormationSolutionStrategy As String ' Auto, UltimateStrain или LoadMultiplier для поиска стадии образования трещины.
Private mFormationSolutionMethod As String ' Фактически выполненный метод; Ncrc для аналитической центральной ветки.
Private mConcreteRbtSer As Double ' Rbt,ser для центральной ветки Ncrc и модели образования трещины, МПа.
Private mConcreteEb As Double ' Eb для Ared в центральной ветке Ncrc, МПа.
Private mConcreteEbt0 As Double ' eps_bt0 для формулы (8.54) СП 63 при однозначно растянутой эпюре.
Private mConcreteEbt2 As Double ' eps_bt2: предельная растягивающая деформация при двузначной эпюре.
Private mSteelEs As Double ' Es для приведенной площади центральной ветки Ncrc, МПа.
Private mCrackInitiationSpec As CMaterialModelSpec ' Спецификация модели образования трещины для обязательной предварительной проверки.
Private mCrackedStateSpec As CMaterialModelSpec ' Спецификация модели раскрытой трещины текущего LC.
Private mPreCrackState As CSectionStateResult ' Необязательное состояние до выключения растянутого бетона.
Private mPreCrackSolver As CSectionSolver ' Solver состояния PreCrackState для warm start состояния сразу после трещины.
Private mPostCrackState As CSectionStateResult ' Необязательное состояние после выключения растянутого бетона для sigma_s,crc.
Private mPostCrackSolver As CSectionSolver ' Solver состояния PostCrackState, чтобы не решать его повторно для sigma_s,crc.
Private mStateProvider As CStateProvider ' Внешний provider конечных named-state; search-probes трещин через него не кэшируются.
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
Private mLambdaCrc As Double ' Множитель нагрузки на пороге трещины: eps_bt,ult для общей ветки или Ncrc/N для центрального растяжения.
Private mNcrc As Double ' Усилие образования трещины для ветки истинного центрального растяжения, Н.
Private mFormationNcrc As Double ' Продольная сила в общей N+M ветке на найденной стадии образования трещины, Н.
Private mMcrc As Double ' Модуль изгибающего момента на стадии образования трещины, Н*мм.
Private mAred As Double ' Площадь приведенного сечения для Ncrc при центральном растяжении, мм2.
Private mCentralTensionBranch As Boolean ' True, если сработала отдельная ветка центрального растяжения.
Private mInputLoadState As CSectionLoadState ' Единый нормализованный LC с пользовательской точкой приложения.
Private mActiveSearchLoadPath As CLoadPathDescriptor ' Общий описатель конкретной траектории текущей search-session.
Private mRequestedLoadPath As String ' Исходный выбор LC; физическое отсутствие точки для Auto и фиксированного пути трактуется раздельно.
Private mSearchReferenceX As Double ' Центр, относительно которого проверяется осевая составляющая выбранного пути, мм.
Private mSearchReferenceY As Double ' Центр, относительно которого проверяется осевая составляющая выбранного пути, мм.


' ------------------------------
' Настройки по умолчанию
' ------------------------------
' Задает default пути, стратегии и внутренних solver-options. Параметры
' материалов перед расчетом берутся из фактического material provider/spec.
Private Sub Class_Initialize()
    mCrackFormationPath = CRACK_PATH_AUTO
    mCrackFormationSolutionStrategy = STRATEGY_AUTO
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
' Стратегия Formation обязательна. Путь передается из строки сочетания,
' а не из отдельной настройки Config. Пустой численный Solver-параметр ошибочен.
Public Sub ApplySettings(ByVal settings As CSystemSettingsReader, Optional ByVal units As CUnitSystem = Nothing)
    If settings Is Nothing Then Exit Sub

    mCrackFormationSolutionStrategy = settings.GetRequiredString("SLS.Crack.InitiationSolutionStrategy")
    mCrackFormationSolutionStrategy = NormalizeCrackFormationSolutionStrategy(mCrackFormationSolutionStrategy, settings)

    mSolverMaxIterations = settings.GetLong("Solver.MaxIterations", mSolverMaxIterations, False)
    mSolverLoadSteps = settings.GetLong("Solver.LoadSteps", mSolverLoadSteps, False)
    mSolverToleranceN = settings.GetDouble("Solver.ToleranceN", mSolverToleranceN, False)
    mSolverToleranceMx = settings.GetDouble("Solver.ToleranceMx", mSolverToleranceMx, False)
    mSolverToleranceMy = settings.GetDouble("Solver.ToleranceMy", mSolverToleranceMy, False)
    mSolverLineSearchEnabled = settings.GetBoolean("Solver.LineSearchEnabled", mSolverLineSearchEnabled)
    mSolverDampingInitial = settings.GetDouble("Solver.DampingInitial", mSolverDampingInitial, False)
    mSolverMinLineSearchAlpha = settings.GetDouble("Solver.MinLineSearchAlpha", mSolverMinLineSearchAlpha, False)
    mSolverMaxDeltaEpsilon0 = settings.GetDouble("Solver.MaxDeltaEpsilon0", mSolverMaxDeltaEpsilon0, False)
    mSolverMaxDeltaKappa = settings.GetDouble("Solver.MaxDeltaKappa", mSolverMaxDeltaKappa, False)
    If settings.HasKey("Solver.Method") Then mSolverMethod = settings.GetRawString("Solver.Method", mSolverMethod)
    mSolverSecantMaxRestarts = settings.GetLong("Solver.SecantMaxRestarts", mSolverSecantMaxRestarts, False)
    mSolverSecantMinStepNorm = settings.GetDouble("Solver.SecantMinStepNorm", mSolverSecantMinStepNorm, False)

    If Not units Is Nothing Then
        ' Отрицательный допуск остается ошибкой настройки и после перевода единиц.
        mSolverToleranceN = Sgn(mSolverToleranceN) * Abs(units.InputForceToInternal(mSolverToleranceN))
        mSolverToleranceMx = Sgn(mSolverToleranceMx) * Abs(units.InputMomentMxToInternal(mSolverToleranceMx))
        mSolverToleranceMy = Sgn(mSolverToleranceMy) * Abs(units.InputMomentMyToInternal(mSolverToleranceMy))
        mSolverMaxDeltaKappa = Sgn(mSolverMaxDeltaKappa) * Abs(units.InputCurvatureToInternal(mSolverMaxDeltaKappa))
    End If
End Sub


' Сбрасывает весь итог предыдущего LC, сохраняя только настройки.
Private Sub ClearResult()
    mConverged = False
    mStopReason = vbNullString
    mDiagnosticLog = vbNullString
    mSolverCallCount = 0
    mSearchExecuted = False
    mFormationInternalStatus = rsUnset
    mFormationResultCode = rcNone
    mCrackFormed = False
    mLambdaCrc = 0#
    mNcrc = 0#
    mFormationNcrc = 0#
    mMcrc = 0#
    mAred = 0#
    Set mPreCrackState = Nothing
    Set mPreCrackSolver = Nothing
    Set mPostCrackState = Nothing
    Set mPostCrackSolver = Nothing
    Set mCrackInitiationSpec = Nothing
    Set mCrackedStateSpec = Nothing
    mCentralTensionBranch = False
    mFormationLoadDefined = False
    mFormationLoadN = 0#
    mFormationLoadMx = 0#
    mFormationLoadMy = 0#
    mFormationPathFallbackPsi1 = False
    mResolvedCrackFormationPath = vbNullString
    mFormationSolutionMethod = vbNullString
    Set mInputLoadState = Nothing
    Set mActiveSearchLoadPath = Nothing
    mRequestedLoadPath = vbNullString
    mSearchReferenceX = 0#
    mSearchReferenceY = 0#
End Sub

' Выполняет самостоятельный formation-сценарий без Width-калькулятора.
' Выбирает физическую постановку, получает Pre/Post через общий State pipeline
' и возвращает неизменяемый результат для последующих формул и вывода.
Public Function CheckFormation(ByVal section As CSectionModel, _
        ByVal materialProvider As CMaterialModelProvider, _
        ByVal crackedStateSpec As CMaterialModelSpec, ByVal crackInitiationSpec As CMaterialModelSpec, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double, _
        ByVal loadStateForClassification As CSectionLoadState, _
        ByVal centralReferenceX As Double, ByVal centralReferenceY As Double) As CCrackFormationResult
    Dim requestedPath As String
    requestedPath = NormalizeCrackFormationPath(mCrackFormationPath)
    CheckFormationCore section, materialProvider, crackedStateSpec, crackInitiationSpec, _
        targetN, targetMx, targetMy, loadStateForClassification, centralReferenceX, centralReferenceY
    mCrackFormationPath = requestedPath
    Dim result As CCrackFormationResult
    Set result = New CCrackFormationResult
    result.InitializeFromCalculator Me, ResultMeta
    Set CheckFormation = result
End Function

' Возвращает фактический машинный исход formation, не выводя его из строк.
Public Property Get ResultMeta() As CResultMeta
    Dim meta As CResultMeta
    Set meta = New CResultMeta
    If mFormationInternalStatus = rsUnset Then
        meta.SetResult rsInternalError, rcInternalError, rkCrackFormation, "Результат образования трещины не заполнен.", _
            vbNullString, True, False
    Else
        meta.SetResult mFormationInternalStatus, mFormationResultCode, rkCrackFormation, mStopReason, _
            vbNullString, True, (mSearchExecuted Or mSolverCallCount > 0 Or mFormationInternalStatus = rsSuccess Or _
                mFormationInternalStatus = rsSuccessWithWarning)
    End If
    Set ResultMeta = meta
End Property

' Подтверждает физическую точку, а не только сохраненную нагрузку кандидата.
' Неуспешное PreCrackState остается диагностикой отказа и не дает limit point;
' аналитический Ncrc без State допустим только при успешном исходе проверки.
Public Property Get HasLimitPoint() As Boolean
    If Not mFormationLoadDefined Or mFormationPathFallbackPsi1 Then Exit Property
    If Not mPreCrackState Is Nothing Then
        HasLimitPoint = mPreCrackState.Converged And mPreCrackState.WithinPhysicalRange And _
            Not mPreCrackState.ExtensionUsed
    Else
        HasLimitPoint = (mFormationInternalStatus = rsSuccess Or mFormationInternalStatus = rsSuccessWithWarning)
    End If
End Property

' Передает исходный полный вектор найденной точки, до переноса моментов.
Public Sub PointLoads(ByRef nValue As Double, ByRef mxValue As Double, ByRef myValue As Double)
    FormationLoads nValue, mxValue, myValue
End Sub

' Fallback сохраняет факт трещины, но не придумывает PostState для psi.
Public Property Get FallbackPsi1() As Boolean
    FallbackPsi1 = mFormationPathFallbackPsi1
End Property

' Выполняет проверку и сохраняет рабочий контекст одной formation-задачи.
Private Sub CheckFormationCore(ByVal section As CSectionModel, _
        ByVal materialProvider As CMaterialModelProvider, ByVal crackedStateSpec As CMaterialModelSpec, _
        ByVal crackInitiationSpec As CMaterialModelSpec, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double, _
        ByVal loadStateForClassification As CSectionLoadState, _
        ByVal centralReferenceX As Double, ByVal centralReferenceY As Double)
    ClearResult
    ValidateFormationInputs section, materialProvider, loadStateForClassification, crackedStateSpec, crackInitiationSpec
    Set mInputLoadState = loadStateForClassification
    Set mCrackedStateSpec = crackedStateSpec.Clone
    If crackInitiationSpec.IsComplete Then Set mCrackInitiationSpec = crackInitiationSpec.Clone
    ApplyMaterialProviderProperties materialProvider
    NormalizeSettings
    mRequestedLoadPath = mCrackFormationPath
    mSearchReferenceX = centralReferenceX
    mSearchReferenceY = centralReferenceY

    ' Нулевой Auto-LC не задает физической траектории. Не запускаем Search
    ' ради отсутствующего действия; явный пустой scaled path проверяется
    ' общим механизмом отдельно как ошибка выбора пути.
    If mRequestedLoadPath = CRACK_PATH_AUTO And loadStateForClassification.N = 0# And _
            loadStateForClassification.LoadPointMx = 0# And loadStateForClassification.LoadPointMy = 0# Then
        mConverged = True
        mResolvedCrackFormationPath = "None"
        SetFormationResult rsNotApplicable, rcNotApplicable, _
            "Нагрузки сочетания равны нулю; поиск точки образования трещины не требуется."
        Exit Sub
    End If

    Dim paths As Variant, coordinator As CLimitSearchCoordinator
    Set coordinator = New CLimitSearchCoordinator
    paths = coordinator.LoadPathKeys(mCrackFormationPath, loadStateForClassification)
    If IsCentralTensionState(loadStateForClassification, centralReferenceX, centralReferenceY) And _
            UBound(paths) = LBound(paths) And _
            (CStr(paths(LBound(paths))) = CRACK_PATH_N Or CStr(paths(LBound(paths))) = CRACK_PATH_NMXY) Then
        mCrackFormationPath = CStr(paths(LBound(paths)))
        CheckCentralTensionFormation section, materialProvider, targetN, targetMx, targetMy, loadStateForClassification
    Else
        CheckGeneralFormation section, materialProvider, targetN, targetMx, targetMy, _
            loadStateForClassification, centralReferenceX, centralReferenceY
    End If
End Sub


' Проверяет только настройки образования трещины и физические параметры
' его модели; коэффициенты и предел ширины принадлежат Width.
Private Sub NormalizeSettings()
    mCrackFormationPath = NormalizeCrackFormationPath(mCrackFormationPath)
    mCrackFormationSolutionStrategy = NormalizeCrackFormationSolutionStrategy(mCrackFormationSolutionStrategy)
    If mConcreteRbtSer <= 0# Then Err.Raise vbObjectError + 3711, "CCrackFormationCalculator", "Concrete.Rbt.SLS должен быть положительным для расчета трещин."
    If mConcreteEb <= 0# Then Err.Raise vbObjectError + 3712, "CCrackFormationCalculator", "Concrete.Eb должен быть положительным для расчета трещин."
    If mSteelEs <= 0# Then Err.Raise vbObjectError + 3713, "CCrackFormationCalculator", "Steel.Es должен быть положительным для расчета трещин."
    If mConcreteEbt0 <= 0# Then Err.Raise vbObjectError + 3714, "CCrackFormationCalculator", "Concrete.Ebt0 должен быть положительным для проверки образования трещины."
    If mConcreteEbt2 <= 0# Then Err.Raise vbObjectError + 3715, "CCrackFormationCalculator", "Concrete.Ebt2 должен быть положительным для проверки образования трещины."
    If mConcreteEbt0 > mConcreteEbt2 Then Err.Raise vbObjectError + 3716, "CCrackFormationCalculator", "Concrete.Ebt0 не должен быть больше Concrete.Ebt2 для формулы (8.54)."
End Sub


' Проверяет минимальный набор данных для предварительной проверки образования
' нормальной трещины.
' В отличие от основного Calculate здесь еще нет полного CrackedState solver-а,
' поэтому проверяется только геометрия, поставщик материалов и модель
' CrackInitiation.
Private Sub ValidateFormationInputs(ByVal section As CSectionModel, ByVal materialProvider As CMaterialModelProvider, _
        ByVal loadStateForClassification As CSectionLoadState, ByVal crackedStateSpec As CMaterialModelSpec, _
        ByVal crackInitiationSpec As CMaterialModelSpec)
    If section Is Nothing Then Err.Raise vbObjectError + 3726, "CCrackFormationCalculator", "Модель расчетного сечения не передана."
    If section.ConcreteCount <= 0 Then Err.Raise vbObjectError + 3727, "CCrackFormationCalculator", "В модели расчетного сечения нет бетонных элементов."
    If section.RebarCount <= 0 Then Err.Raise vbObjectError + 3728, "CCrackFormationCalculator", "В модели расчетного сечения нет арматурных элементов."
    If materialProvider Is Nothing Then Err.Raise vbObjectError + 3729, "CCrackFormationCalculator", "Поставщик материалов не передан."
    If loadStateForClassification Is Nothing Then Err.Raise vbObjectError + 3730, "CCrackFormationCalculator", "CSectionLoadState не передан."
    If crackedStateSpec Is Nothing Then _
        Err.Raise vbObjectError + 3731, "CCrackFormationCalculator", "MaterialModel.CrackedState не передан."
    If Not crackedStateSpec.IsComplete Then _
        Err.Raise vbObjectError + 3731, "CCrackFormationCalculator", "MaterialModel.CrackedState не передан."
    If crackInitiationSpec Is Nothing Then _
        Err.Raise vbObjectError + 3732, "CCrackFormationCalculator", "Для Calculation.Crack.Width нужна MaterialModel.CrackInitiation."
    If Not crackInitiationSpec.IsComplete Then _
        Err.Raise vbObjectError + 3732, "CCrackFormationCalculator", "Для Calculation.Crack.Width нужна MaterialModel.CrackInitiation."
End Sub



' Проверяет центральное растяжение по Ncrc. Для единого snapshot-а состояния
' до и после трещины сохраняются так же, как в общей ветке образования
' трещины.
Private Sub CheckCentralTensionFormation(ByVal section As CSectionModel, _
        ByVal materialProvider As CMaterialModelProvider, ByVal targetN As Double, _
        ByVal targetMx As Double, ByVal targetMy As Double, _
        ByVal loadStateForClassification As CSectionLoadState)
    mCentralTensionBranch = True
    mFormationSolutionMethod = "Ncrc"

    If Not IsTensionLoad(loadStateForClassification) Then
        FinishNoCrack "Центральная ветка: продольная сила не является растягивающей."
        Exit Sub
    End If

    mNcrc = CentralTensionCrackForce(section)
    mFormationNcrc = mNcrc
    mMcrc = 0#
    If mNcrc <= 0# Or targetN <= 0# Then
        SetFormationInvalidInput "InputError: для центрального растяжения не удалось определить Ncrc."
        Exit Sub
    End If

    Dim activePath As String
    activePath = mCrackFormationPath

    mLambdaCrc = mNcrc / targetN
    If mLambdaCrc <= 0# Then
        SetFormationNumericalFailure "NumericalFailure: найденное lambda_crc по Ncrc неположительно."
        Exit Sub
    End If

    If mLambdaCrc > 1# Then
        StoreFormationLoad targetN * mLambdaCrc, targetMx * mLambdaCrc, targetMy * mLambdaCrc
        mResolvedCrackFormationPath = activePath
        FinishFormationNotReached "Трещина не образована: расчетная сила образования трещины больше заданной растягивающей силы."
        Exit Sub
    End If

    StoreFormationLoad targetN * mLambdaCrc, targetMx * mLambdaCrc, targetMy * mLambdaCrc
    mResolvedCrackFormationPath = activePath
    If Not BuildPreCrackState(section, materialProvider, targetN, targetMx, targetMy) Then Exit Sub
    If Not EnsurePostCrackState(section, materialProvider, Nothing, targetN, targetMx, targetMy) Then Exit Sub

    mConverged = True
    mCrackFormed = True
    SetFormationSuccess "Предварительная проверка Ncrc: трещина образуется при текущей нагрузке."
    AppendDiagnostic "formation=Ncrc; lambda_crc=" & FormatNumberInvariant(mLambdaCrc)
End Sub



' Проверяет общую изгибную/внецентренную ветку по предельной растягивающей
' деформации бетона. Состояние PostCrackState решается сразу, чтобы затем не
' пересчитывать его при уточнении psi_s.
Private Sub CheckGeneralFormation(ByVal section As CSectionModel, _
        ByVal materialProvider As CMaterialModelProvider, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double, _
        ByVal loadStateForClassification As CSectionLoadState, _
        ByVal centralReferenceX As Double, ByVal centralReferenceY As Double)
    If mCrackInitiationSpec Is Nothing Then
        SetFormationInvalidInput "InputError: для Calculation.Crack.Width требуется MaterialModel.CrackInitiation."
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

    If mLambdaCrc > 1# Then
        FinishFormationNotReached "Трещина не образована: пороговая точка трещинообразования находится за текущей нагрузкой."
        Exit Sub
    End If

    If mFormationPathFallbackPsi1 Then Exit Sub

    If Not EnsurePostCrackState(section, materialProvider, Nothing, targetN, targetMx, targetMy) Then Exit Sub

    mConverged = True
    mCrackFormed = True
    SetFormationSuccess "Предварительная проверка Mcrc: трещина образуется при текущей нагрузке."
End Sub



' Передает поиск стадии образования нормальной трещины в общий LimitSearch.
' CCrackFormationCalculator остается владельцем crack-критерия, диагностического
' журнала и сохранения PreCrackState/PostCrackState. Физический путь выбирается
' общим Search, численная стратегия выполняется тем же поисковым слоем.
Private Function FindGeneralFormationLambda(ByVal section As CSectionModel, _
        ByVal materialProvider As CMaterialModelProvider, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double, _
        ByVal momentMxAtReference As Double, ByVal momentMyAtReference As Double, _
        ByVal referenceX As Double, ByVal referenceY As Double) As Boolean
    Dim concreteMaterial As CMaterialDiagram
    Dim steelMaterial As CMaterialDiagram
    Set concreteMaterial = materialProvider.ConcreteMaterialForEquilibriumFromSpec(mCrackInitiationSpec)
    Set steelMaterial = materialProvider.SteelMaterialForEquilibriumFromSpec(mCrackInitiationSpec)

    Dim request As CLimitSearchRequest
    Set request = New CLimitSearchRequest
    request.InitializeCrackFormation Me, section, concreteMaterial, steelMaterial, _
        mCrackFormationSolutionStrategy, mCrackFormationPath, targetN, targetMx, targetMy, _
        momentMxAtReference, momentMyAtReference, referenceX, referenceY
    request.ConfigureLoadPathSelection mCrackFormationPath, mInputLoadState

    Dim searchResult As CLimitSearchResult
    Set searchResult = ExecuteFormationSearch(request)
    If searchResult Is Nothing Then Exit Function

    Select Case searchResult.Meta.ResultCode
        Case rcCriterionNotReached
            FindGeneralFormationLambda = mFormationLoadDefined
        Case rcCrackNotFormed
            Exit Function
        Case rcCheckPassed
            If mFormationPathFallbackPsi1 Or mFormationLoadDefined Then
                FindGeneralFormationLambda = True
            Else
                SetFormationNumericalFailure _
                    "NumericalFailure: поиск сообщил об образовании трещины, но не передал усилия точки трещинообразования."
            End If
        Case Else
            FindGeneralFormationLambda = searchResult.Succeeded
    End Select
End Function



' Готовит одну попытку поиска образования трещины для общего LimitSearch.
' Диагностика сохраняется общей для всего Auto-перебора, а расчетные поля
' конкретной попытки очищаются.
Public Sub LimitSearchBeginCrackFormationAttempt(ByVal pathKey As String, _
        Optional ByVal loadPath As CLoadPathDescriptor = Nothing)
    ClearFormationAttemptState
    mCrackFormationPath = pathKey
    Set mActiveSearchLoadPath = loadPath
End Sub

' Сохраняет фактический общий путь и причины Auto-переходов у Formation.
' Последующая проверка наличия трещины не теряет эту историю в ResultComment.
Public Sub LimitSearchFinishLoadPath(ByVal pathKey As String, ByVal history As String)
    mResolvedCrackFormationPath = pathKey
    If Len(Trim$(history)) = 0 Then Exit Sub
    If mFormationInternalStatus = rsSuccess Then mFormationInternalStatus = rsSuccessWithWarning
    mStopReason = Trim$(history) & " Итоговый путь " & CrackFormationPathDisplayNameForKey(pathKey) & ": " & mStopReason
End Sub












' Сообщает координатору, что текущий результат окончательный и переключаться
' с UltimateStrain на LoadMultiplier уже не нужно.
Public Function LimitSearchStoppedTerminally() As Boolean
    LimitSearchStoppedTerminally = FormationSearchStoppedTerminally()
End Function



' Возвращает отображаемое имя текущего пути lambda для диагностического отчета.
Public Function LimitSearchCrackFormationPathDisplayName() As String
    LimitSearchCrackFormationPathDisplayName = CrackFormationPathDisplayName()
End Function



' Возвращает отображаемое имя указанного пути lambda для Auto-диагностики.
Public Function LimitSearchCrackFormationPathDisplayNameForKey(ByVal pathKey As String) As String
    LimitSearchCrackFormationPathDisplayNameForKey = CrackFormationPathDisplayNameForKey(pathKey)
End Function



' Добавляет строку в crack-диагностику из общего LimitSearch.
Public Sub LimitSearchAppendDiagnostic(ByVal text As String)
    AppendDiagnostic text
End Sub



' Очищает текущую причину остановки перед fallback-веткой того же пути.
Public Sub LimitSearchClearStopReason()
    mStopReason = vbNullString
    mFormationInternalStatus = rsUnset
    mFormationResultCode = rcNone
End Sub



' Возвращает предыдущую причину, если fallback-ветка не сформировала своей.
Public Sub LimitSearchSetStopReasonIfEmpty(ByVal text As String)
    If Len(mStopReason) = 0 Then SetFormationNumericalFailure text
End Sub



' Запускает низкоуровневый UltimateStrain-критерий образования трещины.
' Метод публичен только для общего LimitSearch; инженерная логика crack-расчета
' вызывается через самостоятельный CheckFormation.
Public Function LimitSearchRunCrackFormationUltimate(ByVal section As CSectionModel, _
        ByVal concreteMaterial As CMaterialDiagram, ByVal steelMaterial As CMaterialDiagram, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double, _
        ByVal momentMxAtReference As Double, ByVal momentMyAtReference As Double, _
        ByVal referenceX As Double, ByVal referenceY As Double) As Boolean
    mFormationSolutionMethod = STRATEGY_ULTIMATE_STRAIN
    LimitSearchRunCrackFormationUltimate = CalculateLambdaCrcByUltimateStrain(section, _
        concreteMaterial, steelMaterial, targetN, targetMx, targetMy, _
        momentMxAtReference, momentMyAtReference, referenceX, referenceY)
End Function



' Создает load-path для общей ветки LoadMultiplier трещинообразования.
' Метод нужен общему CLoadMultiplierSearch, чтобы не восстанавливать в нем
' crack-специфическую связь моментов относительно расчетной точки.
Public Function LimitSearchCrackFormationLoadPath(ByVal targetN As Double, _
        ByVal targetMomentMx As Double, ByVal targetMomentMy As Double, _
        ByVal referenceX As Double, ByVal referenceY As Double) As CLoadPathVector
    Set LimitSearchCrackFormationLoadPath = FormationLoadPath(targetN, targetMomentMx, targetMomentMy, referenceX, referenceY)
End Function

' Оценивает физический смысл начальной точки LoadMultiplier после решения НДС.
' Постоянная трещина и отсутствие активной нагрузки являются исходами задачи,
' а не правилами общего расширения lambda-скобки.
Public Function LimitSearchAcceptFormationInitialPoint(ByVal targetN As Double, _
        ByVal momentMxAtReference As Double, ByVal momentMyAtReference As Double, _
        ByVal residual As Double) As Boolean
    If residual >= 0# Then
        MarkFormationPathFallback "Постоянная часть пути " & CrackFormationPathDisplayName() & _
            " уже образует нормальную трещину. Отдельная точка ее образования на этом пути не определяется."
        Exit Function
    End If
    If Not LimitSearchCrackFormationHasScaledLoad(targetN, momentMxAtReference, momentMyAtReference) Then
        SetFormationInvalidInput "Для пути " & CrackFormationPathDisplayName() & _
            " в колонке LoadPath сочетания соответствующая нагрузка должна быть ненулевой."
        Exit Function
    End If
    If IsNonTensileForcePath(targetN) Then
        FinishNoCrack "Путь " & CrackFormationPathDisplayName() & _
            " увеличивает только центральное сжатие, а постоянные моменты не образуют нормальную трещину. Критерий трещинообразования на этом пути не достигается."
        Exit Function
    End If
    LimitSearchAcceptFormationInitialPoint = True
End Function



' Проверяет, содержит ли выбранный crack-path ненулевую масштабируемую часть.
' Если Base отсутствует, общий LoadMultiplier-search не запускает бессмысленную
' скобку и оставляет физическое решение за crack-слоем.
Public Function LimitSearchCrackFormationHasScaledLoad(ByVal targetN As Double, _
        ByVal targetMomentMx As Double, ByVal targetMomentMy As Double) As Boolean
    If Not mActiveSearchLoadPath Is Nothing Then
        LimitSearchCrackFormationHasScaledLoad = mActiveSearchLoadPath.HasScaledLoad
        Exit Function
    End If
    Dim path As CLoadPathVector
    Set path = FormationLoadPath(targetN, targetMomentMx, targetMomentMy, mSearchReferenceX, mSearchReferenceY)
    LimitSearchCrackFormationHasScaledLoad = (path.NBase <> 0# Or path.MxBase <> 0# Or path.MyBase <> 0#)
End Function



' Выполняет одну lambda-точку LoadMultiplier для образования трещины.
' Возвращаемая функция отрицательна до eps_bt,ult и неотрицательна после
' достижения порога; это тот же контракт, который использует capacity-search.
Public Function LimitSearchEvaluateCrackFormationLoadMultiplier(ByVal section As CSectionModel, _
        ByVal concreteMaterial As CMaterialDiagram, ByVal steelMaterial As CMaterialDiagram, _
        ByVal loadPath As CLoadPathVector, ByVal startSolver As CSectionSolver, _
        ByVal lambdaValue As Double, ByRef functionValue As Double, _
        ByRef probeState As String, ByRef solver As CSectionSolver) As Boolean
    mFormationSolutionMethod = STRATEGY_LOAD_MULTIPLIER
    mSearchExecuted = True
    Dim probeN As Double
    Dim probeMx As Double
    Dim probeMy As Double
    loadPath.Target lambdaValue, probeN, probeMx, probeMy

    If IsZeroLoadProbe(probeN, probeMx, probeMy) Then
        functionValue = -1#
        probeState = "Acceptable"
        Set solver = Nothing
        LimitSearchEvaluateCrackFormationLoadMultiplier = True
        Exit Function
    End If

    If ProbeLoadScalingFormation(section, concreteMaterial, steelMaterial, loadPath, _
            startSolver, lambdaValue, solver, functionValue) Then
        If functionValue >= 0# Then
            probeState = "CrackFormationLimit"
        Else
            probeState = "Acceptable"
        End If
        LimitSearchEvaluateCrackFormationLoadMultiplier = True
    Else
        probeState = "NumericalFailure"
    End If
End Function



' Отличает нулевую начальную нагрузку от обычной пробы равновесия. Моменты
' сравниваются с фактическими размерными допусками решателя.
Private Function IsZeroLoadProbe(ByVal nValue As Double, ByVal mxValue As Double, _
        ByVal myValue As Double) As Boolean
    IsZeroLoadProbe = (Abs(nValue) <= NEAR_ZERO_FORCE And _
        Abs(mxValue) <= MaxDouble(mSolverToleranceMx, CRACK_DIRECTION_TOLERANCE) And _
        Abs(myValue) <= MaxDouble(mSolverToleranceMy, CRACK_DIRECTION_TOLERANCE))
End Function

' Финализирует найденную LoadMultiplier-точку образования трещины.
' В CCrackFormationCalculator остается только смысл результата: сохранение Ncrc/Mcrc
' и состояния PreCrackState; сам одномерный поиск уже завершен выше.
Public Function LimitSearchFinalizeCrackFormationLoadMultiplier(ByVal section As CSectionModel, _
        ByVal concreteMaterial As CMaterialDiagram, ByVal steelMaterial As CMaterialDiagram, _
        ByVal loadPath As CLoadPathVector, ByVal lambdaValue As Double, _
        ByVal solver As CSectionSolver, ByVal iterations As Long) As Boolean
    If solver Is Nothing Then
        SetFormationResult rsInternalError, rcInternalError, "Точка образования трещины не вернула состояние НДС."
        Exit Function
    End If
    If Not solver.Converged Then
        SetFormationNumericalFailure "NumericalFailure: точка образования трещины не имеет сходящегося НДС."
        Exit Function
    End If

    ' Верхняя сторона скобки доказывает достижение критерия, но не является
    ' автоматически физической предельной точкой. Уточняем плоскость и lambda
    ' общим Newton по тому же пути; сдвиг нагрузки без равновесия запрещен.
    Dim residual As Double
    If Not SolverFormationResidual(section, solver, residual) Then Exit Function
    Dim stateRunner As CStateSolutionRunner
    Set stateRunner = New CStateSolutionRunner
    If Abs(residual) > CRACK_STRAIN_RATIO_TOLERANCE Or _
            Not stateRunner.StateWithinPhysicalRange(section, solver, concreteMaterial, steelMaterial) Then
        Dim problem As CCrackLimitSearchProblem
        Set problem = New CCrackLimitSearchProblem
        problem.Initialize Me
        Dim search As CUltimateStrainSearch
        Set search = New CUltimateStrainSearch
        LimitSearchFinalizeCrackFormationLoadMultiplier = search.RunNewton(problem, section, _
            concreteMaterial, steelMaterial, loadPath, solver, _
            "Не удалось уточнить физическую точку образования трещины после поиска скобки.")
        Exit Function
    End If

    mLambdaCrc = lambdaValue
    StoreFormationLoadFromPath loadPath, mLambdaCrc
    Set mPreCrackSolver = solver
    Set mPreCrackState = New CSectionStateResult
    mPreCrackState.InitializeFromSolver sstPreCrackState, cpMcrc, _
        mCrackInitiationSpec, solver, _
        stateRunner.StateUsesExtension(section, solver, concreteMaterial, steelMaterial), _
        stateRunner.StateWithinPhysicalRange(section, solver, concreteMaterial, steelMaterial)
    If Not AcceptFormationState(mPreCrackState, "До образования трещины") Then Exit Function
    If mLambdaCrc > 1# Then
        FinishFormationNotReached "Трещина не образована: найденная точка трещинообразования находится за текущей нагрузкой."
        LimitSearchFinalizeCrackFormationLoadMultiplier = True
        Exit Function
    End If
    AppendDiagnostic "lambda_crc=" & FormatNumberInvariant(mLambdaCrc) & _
        "; iterations=" & CStr(iterations) & "; method=LoadMultiplier"
    AppendDiagnostic "LoadMultiplier успешно нашел точку образования трещины."
    SetFormationSuccess "Точка образования трещины найдена методом LoadMultiplier."
    LimitSearchFinalizeCrackFormationLoadMultiplier = True
End Function



' Фиксирует физический fallback выбранного crack-path: трещина уже есть в
' постоянной части траектории, поэтому отдельная sigma_s,crc-точка не определяется.
Public Sub LimitSearchMarkFormationPathFallback(ByVal reason As String)
    MarkFormationPathFallback reason
End Sub



' Фиксирует результат "трещина не образована" для выбранного crack-path.
Public Sub LimitSearchFinishNoCrack(ByVal reason As String)
    FinishNoCrack reason
End Sub



' Записывает причину численной остановки LoadMultiplier-поиска трещины.
Public Sub LimitSearchSetCrackFormationStopReason(ByVal reason As String)
    SetFormationNumericalFailure reason
End Sub



' Фиксирует техническое достижение границы crack-search без доказанного физического результата.
' Для Auto-пути это сигнал перейти к следующей траектории, для фиксированного пути - численная ошибка поиска.
Public Sub LimitSearchSetCrackFormationSearchBoundReached(ByVal reason As String)
    mConverged = False
    mCrackFormed = False
    SetFormationResult rsCheckFailed, rcSearchBoundReached, reason
    AppendDiagnostic reason
End Sub



' Возвращает допуск бисекции для точки образования трещины.
Public Property Get LimitSearchCrackLambdaTolerance() As Double
    LimitSearchCrackLambdaTolerance = CRACK_LAMBDA_BISECTION_TOLERANCE
End Property



' Возвращает верхнюю границу расширения скобки трещинообразования.
Public Property Get LimitSearchCrackBracketLimit() As Double
    LimitSearchCrackBracketLimit = CRACK_LAMBDA_BRACKET_LIMIT
End Property



' Возвращает максимум итераций для LoadMultiplier-ветки трещинообразования.
Public Property Get LimitSearchCrackMaxIterations() As Long
    LimitSearchCrackMaxIterations = mSolverMaxIterations
End Property



' Строит стартовую плоскость для общего UltimateStrain-поиска трещинообразования.
Public Sub LimitSearchInitializeUltimateGuess(ByVal section As CSectionModel, _
        ByVal concreteMaterial As CMaterialDiagram, ByVal steelMaterial As CMaterialDiagram, _
        ByVal loadPath As CLoadPathVector, ByVal startSolver As CSectionSolver, _
        ByRef eps0 As Double, ByRef kx As Double, ByRef ky As Double)
    ' Общий Newton уже проверил свои настройки; подготовка старта является
    ' частью реально начатой попытки, даже если ее арифметика завершится отказом.
    mSearchExecuted = True
    Dim guessN As Double
    Dim guessMx As Double
    Dim guessMy As Double
    loadPath.Target 1#, guessN, guessMx, guessMy
    InitializeCrcLoadPathGuess section, concreteMaterial, steelMaterial, startSolver, _
        guessN, guessMx, guessMy, eps0, kx, ky
End Sub



' Возвращает невязки UltimateStrain для критерия образования трещины.
Public Function LimitSearchEvaluateUltimateResidual(ByVal section As CSectionModel, _
        ByVal concreteMaterial As CMaterialDiagram, ByVal steelMaterial As CMaterialDiagram, _
        ByVal loadPath As CLoadPathVector, ByVal eps0 As Double, _
        ByVal kx As Double, ByVal ky As Double, ByRef solver As CSectionSolver, _
        ByRef rPath1 As Double, ByRef rPath2 As Double, ByRef rLimit As Double, _
        ByRef lambdaValue As Double, ByRef state As String) As Boolean
    LimitSearchEvaluateUltimateResidual = EvaluateCrcLoadPathResidual(section, _
        concreteMaterial, steelMaterial, loadPath, eps0, kx, ky, solver, _
        rPath1, rPath2, rLimit, lambdaValue)
    If LimitSearchEvaluateUltimateResidual Then
        If rLimit >= 0# Then state = "CrackFormationLimit" Else state = "Acceptable"
    End If
End Function



' Проверяет, что общий UltimateStrain дошел до точки eps_bt,ult.
Public Function LimitSearchUltimateResidualReached(ByVal rPath1 As Double, _
        ByVal rPath2 As Double, ByVal rLimit As Double) As Boolean
    LimitSearchUltimateResidualReached = (Abs(rPath1) <= CRACK_EQUILIBRIUM_TOLERANCE And _
        Abs(rPath2) <= CRACK_EQUILIBRIUM_TOLERANCE And _
        Abs(rLimit) <= CRACK_STRAIN_RATIO_TOLERANCE)
End Function



' Возвращает норму невязок для общего UltimateStrain line-search при поиске
' трещинообразования. Невязки пути уже выражены в долях допуска равновесия,
' а деформационный критерий приводится к своему допуску. Иначе шаг к точке
' трещины отвергается из-за временной невязки усилий, несмотря на уменьшение
' относительной ошибки критерия. Условия приема State здесь не меняются.
Public Function LimitSearchUltimateResidualNorm(ByVal rPath1 As Double, _
        ByVal rPath2 As Double, ByVal rLimit As Double) As Double
    Dim loadPathMath As CLoadPathMath
    Set loadPathMath = NewLoadPathMath()
    LimitSearchUltimateResidualNorm = loadPathMath.ResidualNorm(rPath1, rPath2, rLimit, _
        CRACK_EQUILIBRIUM_TOLERANCE / CRACK_STRAIN_RATIO_TOLERANCE)
End Function



' Сохраняет найденное состояние образования трещины после общего UltimateStrain.
Public Function LimitSearchFinalizeUltimateResult(ByVal section As CSectionModel, _
        ByVal concreteMaterial As CMaterialDiagram, ByVal steelMaterial As CMaterialDiagram, _
        ByVal loadPath As CLoadPathVector, _
        ByVal lambdaValue As Double, ByVal solver As CSectionSolver, _
        ByVal state As String, ByVal iterations As Long) As Boolean
    If solver Is Nothing Then
        SetFormationResult rsInternalError, rcInternalError, "Не получена предельная плоскость образования трещины."
        Exit Function
    End If
    solver.ToleranceN = mSolverToleranceN
    solver.ToleranceMx = mSolverToleranceMx
    solver.ToleranceMy = mSolverToleranceMy
    If Not solver.ConfirmEquilibrium(loadPath.TargetN(lambdaValue), _
            loadPath.TargetMx(lambdaValue), loadPath.TargetMy(lambdaValue)) Then
        SetFormationNumericalFailure "Предельная плоскость образования трещины не удовлетворяет равновесию."
        Exit Function
    End If
    If lambdaValue < 0# Then
        SetFormationNumericalFailure "NumericalFailure: найденное состояние образования трещины имеет отрицательный lambda_crc."
        Exit Function
    End If
    If TryMarkConstantPartFormationFallback(section, concreteMaterial, steelMaterial, _
            loadPath, lambdaValue) Then
        LimitSearchFinalizeUltimateResult = True
        Exit Function
    End If

    mLambdaCrc = lambdaValue
    StoreFormationLoadFromPath loadPath, mLambdaCrc
    Set mPreCrackSolver = solver
    Set mPreCrackState = New CSectionStateResult
    Dim stateRunner As CStateSolutionRunner
    Set stateRunner = New CStateSolutionRunner
    mPreCrackState.InitializeFromSolver sstPreCrackState, cpMcrc, _
        mCrackInitiationSpec, solver, _
        stateRunner.StateUsesExtension(section, solver, concreteMaterial, steelMaterial), _
        stateRunner.StateWithinPhysicalRange(section, solver, concreteMaterial, steelMaterial)
    If Not AcceptFormationState(mPreCrackState, "До образования трещины") Then Exit Function
    If mLambdaCrc > 1# Then
        FinishFormationNotReached "Трещина не образована: найденная точка трещинообразования находится за текущей нагрузкой."
        LimitSearchFinalizeUltimateResult = True
        Exit Function
    End If
    AppendDiagnostic "lambda_crc=" & FormatNumberInvariant(mLambdaCrc) & _
        "; iterations=" & CStr(iterations) & "; method=UltimateStrain"
    AppendDiagnostic "UltimateStrain успешно нашел точку образования трещины."
    SetFormationSuccess "Точка образования трещины найдена методом UltimateStrain."
    LimitSearchFinalizeUltimateResult = True
End Function



' Переводит UltimateStrain-поиск трещины в численную остановку.
Public Sub LimitSearchSetUltimateNumericalFailure(ByVal reason As String)
    SetFormationNumericalFailure reason
End Sub

' Переводит typed-причину общего Search в Formation meta единым механизмом.
' Отсутствие объекта и ошибка настройки остаются терминальными, не NumFail.
Public Sub LimitSearchSetFailure(ByVal failureCode As ESolverFailureCode, ByVal reason As String)
    Dim meta As CResultMeta
    Set meta = New CResultMeta
    meta.SetSolverFailure failureCode, rkCrackFormation, reason, True
    mConverged = False
    mCrackFormed = False
    SetFormationResult meta.InternalStatus, meta.ResultCode, meta.ResultComment
End Sub



' Настройки общего UltimateStrain-цикла для трещинообразования.
Public Property Get LimitSearchUltimateMaxIterations() As Long
    LimitSearchUltimateMaxIterations = mSolverMaxIterations
End Property



Public Property Get LimitSearchUltimateDampingInitial() As Double
    LimitSearchUltimateDampingInitial = mSolverDampingInitial
End Property



Public Property Get LimitSearchUltimateMinLineSearchAlpha() As Double
    LimitSearchUltimateMinLineSearchAlpha = mSolverMinLineSearchAlpha
End Property



Public Property Get LimitSearchUltimateLineSearchEnabled() As Boolean
    LimitSearchUltimateLineSearchEnabled = mSolverLineSearchEnabled
End Property



Public Property Get LimitSearchUltimateMaxDeltaEpsilon0() As Double
    LimitSearchUltimateMaxDeltaEpsilon0 = mSolverMaxDeltaEpsilon0
End Property



Public Property Get LimitSearchUltimateMaxDeltaKappa() As Double
    LimitSearchUltimateMaxDeltaKappa = mSolverMaxDeltaKappa
End Property



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
    If FormationSearchStoppedTerminally() Then Exit Function

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
    mFormationInternalStatus = rsUnset
    mFormationResultCode = rcNone

    ' Страховочный старт от полного состояния нужен только если быстрый старт
    ' не сработал. Это сохраняет производительность и одновременно помогает
    ' сложным сочетаниям, где уже найденная плоскость НДС дает лучший initial.
    Dim formationStart As CSectionSolver
    mSearchExecuted = True
    Set formationStart = SolveState(section, concreteMaterial, steelMaterial, Nothing, targetN, targetMx, targetMy)
    If formationStart Is Nothing Then
        SetFormationResult rsInternalError, rcInternalError, "Не получен рабочий решатель стартового состояния трещинообразования."
        Exit Function
    End If
    If Not formationStart.Converged Then
        If Len(mStopReason) = 0 Then SetFormationNumericalFailure firstStopReason
        Exit Function
    End If

    If TryUltimateStrainFormation(section, concreteMaterial, steelMaterial, formationStart, _
            targetN, momentMxAtReference, momentMyAtReference, referenceX, referenceY) Then
        AppendDiagnostic "UltimateStrain: дополнительный старт от полного НДС успешно нашел точку образования трещины."
        CalculateLambdaCrcByUltimateStrain = True
    ElseIf Len(mStopReason) = 0 Then
        AppendDiagnostic "UltimateStrain: дополнительный старт от полного НДС также не дал точку образования трещины."
        SetFormationNumericalFailure firstStopReason
    End If
End Function



' Выполняет одну попытку UltimateStrain на выбранном пути нагрузки.
Private Function TryUltimateStrainFormation(ByVal section As CSectionModel, _
        ByVal concreteMaterial As CMaterialDiagram, ByVal steelMaterial As CMaterialDiagram, _
        ByVal startSolver As CSectionSolver, ByVal targetN As Double, _
        ByVal momentMxAtReference As Double, ByVal momentMyAtReference As Double, _
        ByVal referenceX As Double, ByVal referenceY As Double) As Boolean
    Dim loadPath As CLoadPathVector
    Set loadPath = FormationLoadPath(targetN, momentMxAtReference, momentMyAtReference, referenceX, referenceY)

    Dim search As CUltimateStrainSearch
    Set search = New CUltimateStrainSearch
    Dim problem As CCrackLimitSearchProblem
    Set problem = New CCrackLimitSearchProblem
    problem.Initialize Me
    TryUltimateStrainFormation = search.RunCrackFormationNewton(problem, section, _
        concreteMaterial, steelMaterial, loadPath, startSolver)
End Function



' Возвращает True, когда результат прямого поиска уже окончательный и Auto
' не должен переходить к LoadMultiplier: это физическое отсутствие трещины,
' fallback с psi_s=1 или ошибка входных данных.
Private Function FormationSearchStoppedTerminally() As Boolean
    FormationSearchStoppedTerminally = (mConverged And Not mCrackFormed) Or _
        mFormationPathFallbackPsi1 Or _
        (mFormationInternalStatus = rsCheckFailed) Or _
        (mFormationInternalStatus = rsInvalidInput) Or _
        (mFormationInternalStatus = rsInvalidConfiguration) Or _
        (mFormationInternalStatus = rsInternalError)
End Function



' Старт для прямого поиска по общей lambda-траектории. Явно переданная плоскость
' имеет приоритет: это warm-start соседней точки или точки найденной скобки.
' Без нее используем общий CStateGuessBuilder, затем геометрическую оценку.
Private Sub InitializeCrcLoadPathGuess(ByVal section As CSectionModel, _
        ByVal concreteMaterial As CMaterialDiagram, ByVal steelMaterial As CMaterialDiagram, _
        ByVal startSolver As CSectionSolver, ByVal targetN As Double, _
        ByVal targetMx As Double, ByVal targetMy As Double, _
        ByRef eps0 As Double, ByRef kx As Double, ByRef ky As Double)
    If Not startSolver Is Nothing Then
        eps0 = startSolver.Epsilon0
        kx = startSolver.KappaX
        ky = startSolver.KappaY
        Exit Sub
    End If
    Dim guessBuilder As CStateGuessBuilder
    Set guessBuilder = New CStateGuessBuilder
    If guessBuilder.BuildForDirectState(section, concreteMaterial, steelMaterial, _
            targetN, targetMx, targetMy, eps0, kx, ky) Then
        Exit Sub
    End If

    InitializeCrcStrainGuess section, startSolver, targetN, targetMx, targetMy, _
        mConcreteEbt2, eps0, kx, ky
End Sub



' Проверяет найденную точку около lambda=0 отдельным равновесием постоянной
' части. Если она уже треснула, выбранный путь не дает отдельной sigma_s,crc:
' фиксируется доменный fallback, а Auto может попробовать следующий путь.
Private Function TryMarkConstantPartFormationFallback(ByVal section As CSectionModel, _
        ByVal concreteMaterial As CMaterialDiagram, ByVal steelMaterial As CMaterialDiagram, _
        ByVal loadPath As CLoadPathVector, ByVal lambdaValue As Double) As Boolean
    If lambdaValue > CRACK_LAMBDA_BISECTION_TOLERANCE Then Exit Function
    If loadPath Is Nothing Then Exit Function

    Dim probeN As Double
    Dim probeMx As Double
    Dim probeMy As Double
    loadPath.Target 0#, probeN, probeMx, probeMy
    If IsZeroLoadProbe(probeN, probeMx, probeMy) Then
        If IsNonTensileForcePath(loadPath.NBase) Then
            FinishNoCrack "Путь " & CrackFormationPathDisplayName() & _
                " увеличивает только центральное сжатие и не создает растяжения для образования нормальной трещины."
            TryMarkConstantPartFormationFallback = True
        End If
        Exit Function
    End If

    Dim probeSolver As CSectionSolver
    mSearchExecuted = True
    Set probeSolver = SolveState(section, concreteMaterial, steelMaterial, Nothing, _
        probeN, probeMx, probeMy)
    If probeSolver Is Nothing Then Exit Function
    If Not probeSolver.Converged Then Exit Function

    Dim residual As Double
    If Not SolverFormationResidual(section, probeSolver, residual) Then Exit Function
    If FormationProbeBlockedByPhysicalLimit(section, probeSolver, concreteMaterial, steelMaterial, residual) Then Exit Function
    If residual >= 0# Then
        MarkFormationPathFallback "Постоянная часть пути " & CrackFormationPathDisplayName() & _
            " уже образует нормальную трещину. Отдельная точка ее образования на этом пути не определяется."
        TryMarkConstantPartFormationFallback = True
    ElseIf IsNonTensileForcePath(loadPath.NBase) Then
        FinishNoCrack "Путь " & CrackFormationPathDisplayName() & _
            " увеличивает только центральное сжатие, а постоянные моменты не образуют нормальную трещину. Критерий трещинообразования на этом пути не достигается."
        TryMarkConstantPartFormationFallback = True
    ElseIf loadPath.NBase = 0# And loadPath.MxBase = 0# And loadPath.MyBase = 0# Then
        SetFormationInvalidInput "Для пути " & CrackFormationPathDisplayName() & _
            " в колонке LoadPath сочетания соответствующая нагрузка должна быть ненулевой."
        TryMarkConstantPartFormationFallback = True
    End If
End Function



' Оценивает невязки для прямого UltimateStrain-поиска lambda_crc
' на общей линии нагрузки Offset + lambda*Base.
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
        SetFormationInvalidInput "InputError: не удалось определить eps_bt,ult для состояния образования трещины."
        Exit Function
    End If
    rLimit = maxConcreteStrain / epsBtUlt - 1#

    AppendDiagnostic "lambda_crc_path_iter=" & FormatNumberInvariant(lambdaValue) & _
        "; rPath1=" & FormatNumberInvariant(rPath1) & _
        "; rPath2=" & FormatNumberInvariant(rPath2) & _
        "; rLimit=" & FormatNumberInvariant(rLimit) & _
        "; eps_bt_ult=" & FormatNumberInvariant(epsBtUlt) & _
        "; Nint=" & FormatNumberInvariant(solver.Nint) & _
        "; Mxint=" & FormatNumberInvariant(solver.Mxint) & _
        "; Myint=" & FormatNumberInvariant(solver.Myint) & _
        "; eps0=" & FormatNumberInvariant(eps0) & _
        "; kx=" & FormatNumberInvariant(kx) & "; ky=" & FormatNumberInvariant(ky)
    EvaluateCrcLoadPathResidual = True
    Exit Function

Failed:
    SetFormationNumericalFailure "NumericalFailure: оценка lambda_crc по load-path не выполнена: " & Err.Description
End Function



' Создает общий helper lambda-траектории с текущими допусками crack solver-а.
Private Function NewLoadPathMath() As CLoadPathMath
    Dim loadPathMath As CLoadPathMath
    Set loadPathMath = New CLoadPathMath
    loadPathMath.Configure mSolverToleranceN, mSolverToleranceMx, mSolverToleranceMy, NEAR_ZERO_FORCE
    Set NewLoadPathMath = loadPathMath
End Function



' Получает физическое состояние на найденном пороге через общий provider или
' runner. Готовый cache-hit восстанавливается без нового поиска равновесия.
Private Function BuildPreCrackState(ByVal section As CSectionModel, ByVal materialProvider As CMaterialModelProvider, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double) As Boolean
    If mPreCrackState Is Nothing Then
        Dim beforeSolver As CSectionSolver
        Dim formationN As Double
        Dim formationMx As Double
        Dim formationMy As Double
        FormationLoads formationN, formationMx, formationMy

        If Not mStateProvider Is Nothing Then
            Dim request As CStateRequest
            Set request = New CStateRequest
            request.Initialize sstPreCrackState, cpMcrc, mCrackInitiationSpec, _
                formationN, formationMx, formationMy, _
                materialProvider.DiagramExtensionEnabled, materialProvider.DiagramExtensionEnabled
            Set mPreCrackState = mStateProvider.GetOrSolve(request)
            If Not mStateProvider.LastRunner Is Nothing Then
                mSolverCallCount = mSolverCallCount + mStateProvider.LastRunner.SolverCallCount
                Set beforeSolver = mStateProvider.LastRunner.ResultSolver
            ElseIf Not mPreCrackState Is Nothing Then
                Set beforeSolver = mStateProvider.SolverSnapshot(mPreCrackState)
            End If
        Else
            Dim preConcreteMaterial As CMaterialDiagram
            Dim preSteelMaterial As CMaterialDiagram
            Set preConcreteMaterial = materialProvider.ConcreteMaterialForEquilibriumFromSpec(mCrackInitiationSpec)
            Set preSteelMaterial = materialProvider.SteelMaterialForEquilibriumFromSpec(mCrackInitiationSpec)
            Set beforeSolver = SolveState(section, preConcreteMaterial, preSteelMaterial, Nothing, _
                formationN, formationMx, formationMy, True)
            Set mPreCrackState = New CSectionStateResult
            Dim stateRunner As CStateSolutionRunner
            Set stateRunner = New CStateSolutionRunner
            mPreCrackState.InitializeFromSolver sstPreCrackState, cpMcrc, _
                mCrackInitiationSpec, beforeSolver, _
                stateRunner.StateUsesExtension(section, beforeSolver, preConcreteMaterial, preSteelMaterial), _
                stateRunner.StateWithinPhysicalRange(section, beforeSolver, preConcreteMaterial, preSteelMaterial)
        End If

        If Not AcceptFormationState(mPreCrackState, "До образования трещины") Then Exit Function
        If beforeSolver Is Nothing Then
            SetFormationResult rsInternalError, rcInternalError, "Состояние до образования трещины найдено, но его снимок недоступен для дальнейшего расчета."
            Exit Function
        End If
        Set mPreCrackSolver = beforeSolver
    End If
    If Not AcceptFormationState(mPreCrackState, "До образования трещины") Then Exit Function
    BuildPreCrackState = True
End Function



' Возвращает PostCrackState для уже найденного lambda_crc. Предварительная
' проверка создает это состояние заранее, а Width затем читает напряжения
' из неизменяемого State без повторного решения.
Private Function EnsurePostCrackState(ByVal section As CSectionModel, ByVal materialProvider As CMaterialModelProvider, _
        ByVal startSolver As CSectionSolver, ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double) As Boolean
    If Not mPostCrackState Is Nothing Then
        EnsurePostCrackState = AcceptFormationState(mPostCrackState, "После образования трещины")
        Exit Function
    End If

    Dim effectiveStartSolver As CSectionSolver
    Set effectiveStartSolver = startSolver
    If effectiveStartSolver Is Nothing Then Set effectiveStartSolver = mPreCrackSolver

    Dim formationN As Double
    Dim formationMx As Double
    Dim formationMy As Double
    FormationLoads formationN, formationMx, formationMy
    If Not mStateProvider Is Nothing Then
        Dim request As CStateRequest
        Set request = New CStateRequest
        request.Initialize sstPostCrackState, cpCrackedNDS, mCrackedStateSpec, _
            formationN, formationMx, formationMy, _
            materialProvider.DiagramExtensionEnabled, materialProvider.DiagramExtensionEnabled
        Set mPostCrackState = mStateProvider.GetOrSolve(request, effectiveStartSolver)
        If Not mStateProvider.LastRunner Is Nothing Then
            mSolverCallCount = mSolverCallCount + mStateProvider.LastRunner.SolverCallCount
            Set mPostCrackSolver = mStateProvider.LastRunner.ResultSolver
        ElseIf Not mPostCrackState Is Nothing Then
            Set mPostCrackSolver = mStateProvider.SolverSnapshot(mPostCrackState)
        End If
    Else
        Dim postConcreteMaterial As CMaterialDiagram
        Dim postSteelMaterial As CMaterialDiagram
        Set postConcreteMaterial = materialProvider.ConcreteMaterialForEquilibriumFromSpec(mCrackedStateSpec)
        Set postSteelMaterial = materialProvider.SteelMaterialForEquilibriumFromSpec(mCrackedStateSpec)
        Set mPostCrackSolver = SolveState(section, postConcreteMaterial, postSteelMaterial, effectiveStartSolver, _
            formationN, formationMx, formationMy, True)
        Set mPostCrackState = New CSectionStateResult
        Dim stateRunner As CStateSolutionRunner
        Set stateRunner = New CStateSolutionRunner
        mPostCrackState.InitializeFromSolver sstPostCrackState, cpCrackedNDS, _
            mCrackedStateSpec, mPostCrackSolver, _
            stateRunner.StateUsesExtension(section, mPostCrackSolver, postConcreteMaterial, postSteelMaterial), _
            stateRunner.StateWithinPhysicalRange(section, mPostCrackSolver, postConcreteMaterial, postSteelMaterial)
    End If

    If Not AcceptFormationState(mPostCrackState, "После образования трещины") Then Exit Function

    EnsurePostCrackState = True
End Function



' Сохраняет фактическую причину неуспеха State; численная несходимость не
' подменяет ошибку входа, конфигурации, внутреннюю ошибку или физический отказ.
Private Function AcceptFormationState(ByVal state As CSectionStateResult, ByVal stageName As String) As Boolean
    If state Is Nothing Then
        mConverged = False
        mCrackFormed = False
        SetFormationResult rsInternalError, rcInternalError, stageName & ": результат НДС отсутствует."
        Exit Function
    End If
    If Not state.Converged Or Not state.WithinPhysicalRange Then
        Dim meta As CResultMeta
        Set meta = state.ResultMeta
        mConverged = False
        mCrackFormed = False
        SetFormationResult meta.InternalStatus, meta.ResultCode, stageName & ": " & meta.ResultComment
        Exit Function
    End If
    AcceptFormationState = True
End Function


' Сохраняет внутренние усилия, соответствующие уже найденной стадии
' трещинообразования. Дальше PreCrackState/PostCrackState решаются именно по ним,
' чтобы повторные вызовы не пересчитывали lambda и не масштабировали не ту
' часть нагрузки.
Private Sub StoreFormationLoad(ByVal nValue As Double, ByVal mxValue As Double, ByVal myValue As Double)
    mFormationLoadN = nValue
    mFormationLoadMx = mxValue
    mFormationLoadMy = myValue
    mFormationLoadDefined = True
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
        Err.Raise vbObjectError + 3733, "CCrackFormationCalculator", "Внутренние усилия состояния образования трещины не сохранены."
    nValue = mFormationLoadN
    mxValue = mFormationLoadMx
    myValue = mFormationLoadMy
End Sub






' Возвращает общую траекторию текущей search-session. Для низкоуровневого
' запроса без LC восстанавливает нагрузку относительно переданной точки и
' использует тот же описатель пути, не дублируя правила Offset/Base.
Private Function FormationLoadPath(ByVal targetN As Double, _
        ByVal targetMomentMx As Double, ByVal targetMomentMy As Double, _
        ByVal referenceX As Double, ByVal referenceY As Double) As CLoadPathVector
    If Not mActiveSearchLoadPath Is Nothing Then
        Set FormationLoadPath = mActiveSearchLoadPath.Vector
        Exit Function
    End If
    Dim load As CSectionLoadState, descriptor As CLoadPathDescriptor
    Set load = New CSectionLoadState
    load.Initialize targetN, targetMomentMx, targetMomentMy, referenceX, referenceY
    Set descriptor = New CLoadPathDescriptor
    Dim pathKey As String, coordinator As CLimitSearchCoordinator, paths As Variant
    pathKey = mCrackFormationPath
    If pathKey = CRACK_PATH_AUTO Then
        Set coordinator = New CLimitSearchCoordinator
        paths = coordinator.LoadPathKeys(pathKey, load)
        pathKey = CStr(paths(LBound(paths)))
    End If
    descriptor.InitializeFromLoadState pathKey, load
    Set FormationLoadPath = descriptor.Vector
End Function






' Возвращает пользовательское имя пути для диагностики.
Private Function CrackFormationPathDisplayName() As String
    CrackFormationPathDisplayName = CrackFormationPathDisplayNameForKey(mCrackFormationPath)
End Function



' Возвращает имя фактически использованного пути. Для Auto это уже не Auto,
' а конкретный lambda-путь, на котором была принята стадия образования.
Private Function CrackFormationPathDisplayNameForKey(ByVal pathKey As String) As String
    Dim descriptor As CLoadPathDescriptor
    Set descriptor = New CLoadPathDescriptor
    CrackFormationPathDisplayNameForKey = descriptor.DisplayNameForKey(pathKey)
End Function



' Берет Rbt,ser, Eb, eps_bt0, eps_bt2 и Es из поставщика материалов.
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



' Решает одну lambda-точку общего LoadMultiplier через единый state pipeline.
' Подбирает полезный warm-start, при его неудаче пробует самостоятельный старт
' и возвращает formation residual только после подтвержденного равновесия.
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
    If FormationSearchStoppedTerminally() Then Exit Function
    If SolverFormationResidual(section, solver, residual) Then
        If FormationProbeBlockedByPhysicalLimit(section, solver, concreteMaterial, steelMaterial, residual) Then Exit Function
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
        If FormationSearchStoppedTerminally() Then Exit Function
        If SolverFormationResidual(section, solver, residual) Then
            If FormationProbeBlockedByPhysicalLimit(section, solver, concreteMaterial, steelMaterial, residual) Then Exit Function
            AppendLoadScalingProbeDiagnostic lambdaValue, residual, True
            ProbeLoadScalingFormation = True
            Exit Function
        End If
    End If

    ' Если warm-start от соседней lambda-точки оказался неудачным, пробуем
    ' стартовую плоскость от нагрузки в текущей lambda-точке. Это тот же общий механизм,
    ' которым пользуются опасные траектории DirectState/Capacity: меняется
    ' только начальная плоскость Newton, а не сама траектория lambda.
    If Not startSolver Is Nothing Then
        Dim probeGuess As CSectionSolver
        Set probeGuess = LoadScalingProbeGuess(section, concreteMaterial, steelMaterial, _
            probeN, probeMx, probeMy)
        If Not probeGuess Is Nothing Then
            Set solver = SolveState(section, concreteMaterial, steelMaterial, probeGuess, _
                probeN, probeMx, probeMy, True)
            If FormationSearchStoppedTerminally() Then Exit Function
            ProbeLoadScalingFormation = SolverFormationResidual(section, solver, residual)
            If ProbeLoadScalingFormation Then
                If FormationProbeBlockedByPhysicalLimit(section, solver, concreteMaterial, steelMaterial, residual) Then
                    ProbeLoadScalingFormation = False
                    Exit Function
                End If
                AppendLoadScalingProbeDiagnostic lambdaValue, residual, True
            End If
        End If
    End If
    If Not ProbeLoadScalingFormation Then
        CaptureFormationSolverFailure solver
        AppendLoadScalingProbeDiagnostic lambdaValue, 0#, False
    End If
End Function



' Записывает одну lambda-точку fallback-метода LoadMultiplier. В отличие от
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



' Возвращает warm-start для решения НДС только если в нем действительно есть
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



' Строит активную стартовую плоскость для отдельной lambda-точки.
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



' Записывает машинный результат этапа образования трещины. StopReason остается
' только пользовательским комментарием и больше не используется для классификации.
Private Sub SetFormationResult(ByVal internalStatus As EResultInternalStatus, _
        ByVal resultCode As EResultCode, ByVal reason As String)
    mFormationInternalStatus = internalStatus
    mFormationResultCode = resultCode
    mStopReason = UserStopReason(reason)
End Sub



' Помечает успешное нахождение точки образования нормальной трещины.
Private Sub SetFormationSuccess(ByVal reason As String)
    If mFormationInternalStatus = rsSuccessWithWarning Then
        SetFormationResult rsSuccessWithWarning, rcCheckPassed, mStopReason & " " & reason
    Else
        SetFormationResult rsSuccess, rcCheckPassed, reason
    End If
End Sub



' Помечает ошибку пользовательских данных или обязательных входов crack-этапа.
Private Sub SetFormationInvalidInput(ByVal reason As String)
    mConverged = False
    mCrackFormed = False
    SetFormationResult rsInvalidInput, rcInvalidInput, reason
End Sub



' Помечает именно численную неудачу поиска состояния/предела.
Private Sub SetFormationNumericalFailure(ByVal reason As String)
    ' Последующая неудачная проба не отменяет уже подтвержденный физический
    ' отказ на этом пути. Новая Auto-попытка очищает статус самостоятельно.
    If mFormationInternalStatus = rsCheckFailed Or mFormationInternalStatus = rsInvalidInput Or _
            mFormationInternalStatus = rsInvalidConfiguration Or mFormationInternalStatus = rsInternalError Then Exit Sub
    If mFormationInternalStatus = rsNumericalFailure And mFormationResultCode = rcSingularTangent Then
        AppendDiagnostic reason
        Exit Sub
    End If
    mConverged = False
    mCrackFormed = False
    SetFormationResult rsNumericalFailure, rcNumericalFailure, reason
End Sub

' Проба с равновесием, вышедшая за другой физический предел до критерия
' трещины, не является допустимой нижней стороной crack-скобки. Причина
' остается инженерным отказом, а не численной несходимостью следующей пробы.
Private Function FormationProbeBlockedByPhysicalLimit(ByVal section As CSectionModel, _
        ByVal solver As CSectionSolver, ByVal concreteMaterial As CMaterialDiagram, _
        ByVal steelMaterial As CMaterialDiagram, ByVal residual As Double) As Boolean
    If solver Is Nothing Then Exit Function
    If Not solver.Converged Then Exit Function
    If residual >= 0# Then Exit Function
    Dim runner As CStateSolutionRunner
    Set runner = New CStateSolutionRunner
    If runner.StateWithinPhysicalRange(section, solver, concreteMaterial, steelMaterial) Then Exit Function
    mConverged = False
    mCrackFormed = False
    SetFormationResult rsCheckFailed, rcPhysicalLimitExceeded, _
        "При поиске образования трещины физический предел бетона или арматуры достигнут раньше порога трещинообразования."
    FormationProbeBlockedByPhysicalLimit = True
End Function



' Готовит короткий пользовательский комментарий из внутренней причины.
' Машинная классификация уже сохранена в InternalStatus/ResultCode, поэтому
' технические префиксы NumericalFailure/InputError на лист выводить не нужно.
Private Function UserStopReason(ByVal reason As String) As String
    UserStopReason = Trim$(reason)
    If Len(UserStopReason) = 0 Then Exit Function

    UserStopReason = Replace$(UserStopReason, "NumericalFailure:", vbNullString, 1, 1, vbTextCompare)
    UserStopReason = Replace$(UserStopReason, "InputError:", vbNullString, 1, 1, vbTextCompare)
    UserStopReason = Replace$(UserStopReason, "lambda_crc", "lambda трещинообразования", 1, -1, vbTextCompare)
    UserStopReason = Replace$(UserStopReason, "eps_bt,ult", "предельная растягивающая деформация бетона", 1, -1, vbTextCompare)
    UserStopReason = Trim$(UserStopReason)
    Do While InStr(1, UserStopReason, "..", vbBinaryCompare) > 0
        UserStopReason = Replace$(UserStopReason, "..", ".")
    Loop
End Function



' Помечает расчет как корректно завершенный без раскрытой трещины.
' Это не численная ошибка: заданное состояние либо не растягивает арматуру,
' либо еще не достигло момента образования трещины.
Private Sub FinishNoCrack(ByVal reason As String)
    mConverged = True
    mCrackFormed = False
    mStopReason = reason
    If mRequestedLoadPath <> CRACK_PATH_AUTO And Len(mRequestedLoadPath) > 0 Then
        mConverged = False
        SetFormationResult rsCheckFailed, rcCriterionNotReached, reason
    ElseIf mFormationInternalStatus = rsUnset Then
        SetFormationResult rsSuccess, rcCriterionNotReached, reason
    End If
    AppendDiagnostic reason
End Sub



' Завершает именно этап проверки образования трещины как корректный результат
' "порог найден, но текущая нагрузка его еще не достигла". В отличие от общего
' FinishNoCrack здесь сохраняется успех нахождения порога. Критерий достигнут
' на выбранном пути, но текущий LC до него еще не дошел; это не доказанная
' недостижимость критерия. PostCrackState для текущей ширины не требуется.
Private Sub FinishFormationNotReached(ByVal reason As String)
    mConverged = True
    mCrackFormed = False
    SetFormationResult rsSuccess, rcCheckPassed, reason
    AppendDiagnostic reason
End Sub



' Сохраняет физический отказ постоянной части выбранной траектории.
' Положительной точки образования на ней нет; Auto может попробовать другой
' путь, а фиксированный путь возвращает BaseFail без фиктивных Pre/Post.
Private Sub MarkFormationPathFallback(ByVal reason As String)
    mLambdaCrc = 0#
    mFormationPathFallbackPsi1 = True
    mConverged = False
    mCrackFormed = False
    SetFormationResult rsCheckFailed, rcInitialStateBeyondLimit, reason
    AppendDiagnostic reason
End Sub

' Проверяет осевой путь, не увеличивающий растягивающее действие.
' Эксцентрическая N сюда не попадает: ее масштабируемый момент может вызвать
' трещину, поэтому такой путь обязан пройти обычный поиск равновесия.
Private Function IsNonTensileForcePath(ByVal targetN As Double) As Boolean
    If mCrackFormationPath <> CRACK_PATH_N And mCrackFormationPath <> CRACK_PATH_NMXY Then Exit Function
    If targetN >= 0# Then Exit Function
    If mActiveSearchLoadPath Is Nothing Then
        IsNonTensileForcePath = (mCrackFormationPath = CRACK_PATH_N)
    Else
        IsNonTensileForcePath = _
            Abs(mActiveSearchLoadPath.MxBase - targetN * mSearchReferenceY) <= MaxDouble(mSolverToleranceMx, 1#) And _
            Abs(mActiveSearchLoadPath.MyBase - targetN * mSearchReferenceX) <= MaxDouble(mSolverToleranceMy, 1#)
    End If
End Function



' Запускает служебное НДС через общий CStateSolutionRunner. Так состояния
' PreCrackState/PostCrackState получают тот же централизованный подбор стартовой
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
    If runner.ResultSolver Is Nothing Then
        CaptureFormationSolverFailure Nothing
    Else
        Select Case runner.ResultSolver.FailureCode
            Case sfcInvalidInput, sfcInvalidConfiguration, sfcInternalError
                CaptureFormationSolverFailure runner.ResultSolver
        End Select
    End If
End Function

' Сохраняет машинную причину неуспешной state-пробы. Вызов после исчерпания
' численных стартов сохраняет singular tangent отдельно; terminal ошибки
' фиксируются сразу и не заменяются последующими generic сообщениями Search.
Private Sub CaptureFormationSolverFailure(ByVal solver As CSectionSolver)
    mConverged = False
    mCrackFormed = False
    If solver Is Nothing Then
        SetFormationResult rsInternalError, rcSolverDidNotReturnState, "Поиск трещинообразования не получил результат решателя НДС."
        Exit Sub
    End If
    Dim meta As CResultMeta
    Set meta = New CResultMeta
    meta.SetSolverFailure solver.FailureCode, rkCrackFormation, solver.StopReason, (solver.Iterations > 0)
    SetFormationResult meta.InternalStatus, meta.ResultCode, meta.ResultComment
End Sub



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



' Определяет растягивающий характер продольной силы только через единый
' объект нагрузки. Так central tension и phi3 используют ту же интерпретацию
' Load.ReferenceOffset и фильтра малых моментов, что batch/stability/capacity.
Private Function IsTensionLoad(ByVal loadStateForClassification As CSectionLoadState) As Boolean
    IsTensionLoad = loadStateForClassification.IsTension(NEAR_ZERO_FORCE)
End Function



' Нормализует общий LoadPath из LC тем же описателем, что Capacity.
' Пустое значение означает Auto; отдельной настройки пути трещин в Config нет.
Private Function NormalizeCrackFormationPath(ByVal rawPath As String) As String
    Dim descriptor As CLoadPathDescriptor
    Set descriptor = New CLoadPathDescriptor
    NormalizeCrackFormationPath = descriptor.NormalizeKey(rawPath)
    If Len(NormalizeCrackFormationPath) = 0 Then Err.Raise vbObjectError + 3734, "CCrackFormationCalculator", _
        "В таблице сочетаний исправьте LoadPath: выберите Auto, lambda*Mx, lambda*My, lambda*Mxy, lambda*N или lambda*NMxy."
End Function



' Приводит численную стратегию поиска трещинообразования к той же логике,
' что используется для capacity: Auto, UltimateStrain или LoadMultiplier.
' Допустимые aliases проверяются здесь для всех consumers; reader при наличии
' указывает адрес Config, не влияя на выбор метода или расчетную постановку.
Public Function NormalizeCrackFormationSolutionStrategy(ByVal rawStrategy As String, _
        Optional ByVal settings As CSystemSettingsReader = Nothing) As String
    Select Case LCase$(Trim$(rawStrategy))
        Case vbNullString, "auto"
            NormalizeCrackFormationSolutionStrategy = STRATEGY_AUTO
        Case "ultimatestrain", "strain", "direct"
            NormalizeCrackFormationSolutionStrategy = STRATEGY_ULTIMATE_STRAIN
        Case "loadmultiplier", "multiplier", "lambda"
            NormalizeCrackFormationSolutionStrategy = STRATEGY_LOAD_MULTIPLIER
        Case Else
            Dim message As String
            message = "На листе Config в строке SLS.Crack.InitiationSolutionStrategy выберите Auto, UltimateStrain или LoadMultiplier."
            If Not settings Is Nothing Then message = settings.InputErrorMessage("SLS.Crack.InitiationSolutionStrategy", _
                "Неизвестная стратегия: """ & rawStrategy & """.", "Выберите Auto, UltimateStrain или LoadMultiplier.")
            Err.Raise vbObjectError + 3735, "CCrackFormationCalculator", message
    End Select
End Function



' Сбрасывает только данные одной попытки поиска точки образования трещины. Диагностический
' журнал и выбранные материальные модели сохраняются, чтобы Auto мог показать
' всю цепочку переключения путей.
Private Sub ClearFormationAttemptState()
    mConverged = False
    mStopReason = vbNullString
    mCrackFormed = False
    mFormationInternalStatus = rsUnset
    mFormationResultCode = rcNone
    mLambdaCrc = 0#
    mNcrc = 0#
    mFormationNcrc = 0#
    mMcrc = 0#
    Set mPreCrackState = Nothing
    Set mPreCrackSolver = Nothing
    Set mPostCrackState = Nothing
    Set mPostCrackSolver = Nothing
    mFormationLoadDefined = False
    mFormationLoadN = 0#
    mFormationLoadMx = 0#
    mFormationLoadMy = 0#
    mFormationPathFallbackPsi1 = False
    mResolvedCrackFormationPath = vbNullString
    mFormationSolutionMethod = vbNullString
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



' Утилитарный минимум без WorksheetFunction в расчетном ядре трещин.
Private Function MinDouble(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then MinDouble = a Else MinDouble = b
End Function



' Утилитарный максимум без WorksheetFunction в расчетном ядре трещин.
Private Function MaxDouble(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then MaxDouble = a Else MaxDouble = b
End Function



' --------------------------------------------------------------------------
' Свойства результатов
' --------------------------------------------------------------------------
' Сохраняет машинный исход и подтвержденное НДС в общий численный результат.
' Search-result не трактует отсутствие трещины при текущем LC и не читает
' рабочий Formation-контекст после создания этого снимка. Факт Search хранится
' отдельно: аналитический Ncrc может получить Pre/PostState без lambda-поиска.
Public Function BuildSearchSnapshot(ByVal requestedStrategy As String) As CLimitSearchResult
    Dim meta As CResultMeta
    Set meta = ResultMeta
    Dim nValue As Double, mxValue As Double, myValue As Double
    If HasLimitPoint Then PointLoads nValue, mxValue, myValue
    Dim result As CLimitSearchResult
    Set result = New CLimitSearchResult
    result.Initialize meta, requestedStrategy, mFormationSolutionMethod, HasLimitPoint, _
        mLambdaCrc, nValue, mxValue, myValue, mPreCrackState, vbNullString, mDiagnosticLog, _
        vbNullString, mSearchExecuted
    Set BuildSearchSnapshot = result
End Function

' Writer-ы и тесты читают эти значения после Calculate. Свойства ничего не
' пересчитывают и не обращаются к solver-у повторно.

Public Property Get Converged() As Boolean
    Converged = mConverged
End Property



Public Property Get StopReason() As String
    StopReason = mStopReason
End Property



Public Property Get CrackFormationInternalStatus() As EResultInternalStatus
    CrackFormationInternalStatus = mFormationInternalStatus
End Property

' Отличает выбранную Auto-стратегию от фактически выполненного метода поиска.
Public Property Get FormationSolutionMethod() As String
    FormationSolutionMethod = mFormationSolutionMethod
End Property



Public Property Get CrackFormationResultCode() As EResultCode
    CrackFormationResultCode = mFormationResultCode
End Property



Public Property Get CrackFormed() As Boolean
    CrackFormed = mCrackFormed
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



Public Property Get CrackFormationPath() As String
    CrackFormationPath = CrackFormationPathDisplayName()
End Property



' Возвращает обозначение фактически выбранного load path для таблицы трещин.
' Пустая строка означает, что путь еще не принят; это не имя численного метода.
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



Public Property Get CentralTensionBranch() As Boolean
    CentralTensionBranch = mCentralTensionBranch
End Property



' Возвращает состояние на пороге образования трещины с работающим растянутым бетоном.
Public Property Get PreCrackState() As CSectionStateResult
    Set PreCrackState = mPreCrackState
End Property



' Возвращает состояние на том же уровне нагрузки сразу после выключения растянутого бетона.
Public Property Get PostCrackState() As CSectionStateResult
    Set PostCrackState = mPostCrackState
End Property



' Подключает общий provider конечных named-state. Если свойство не задано,
' калькулятор остается пригодным для изолированных тестов и решает состояния
' собственным локальным CStateSolutionRunner.
Public Property Set StateProvider(ByVal value As CStateProvider)
    Set mStateProvider = value
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

' Передает единый выбор и перебор путей общему Search.
Private Function ExecuteFormationSearch(ByVal request As CLimitSearchRequest) As CLimitSearchResult
    Dim coordinator As CLimitSearchCoordinator
    Set coordinator = New CLimitSearchCoordinator
    If request Is Nothing Then Err.Raise vbObjectError + 4231, "CCrackFormationCalculator", "Не передан запрос limit-search."
    Set ExecuteFormationSearch = coordinator.ExecuteCrackFormation(request)
End Function
