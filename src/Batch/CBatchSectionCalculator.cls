VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CBatchSectionCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Пакетный расчет сочетаний нагрузок
' ==========================================================================
' CBatchSectionCalculator запускает один и тот же расчетный набор для всех LC:
' прямое НДС, несущую способность для 1 группы и трещины для 2 группы. Класс
' хранит результаты каждого сочетания в массивах, чтобы writer мог потом
' выгрузить сводку и Results без повторного обращения к solver-ам.
'
' Здесь нет чтения Excel-ячеек и нет генерации геометрии. На вход уже должны
' прийти CSectionModel, материалы и нормализованные усилия N + Mx + My.

Private Const MAX_COMBINATIONS As Long = 20 ' Максимум LC, поддерживаемый таблицей книги.
Private Const STATUS_OK As String = "OK" ' Расчетная ветка завершилась без ошибки.
Private Const STATUS_INVALID_INPUT As String = "InvalidInput" ' Ошибка входных данных, а не численная несходимость.
Private Const STATUS_SKIPPED As String = "SkippedZeroBaseMoment" ' Capacity не запускался из-за нулевого моментного вектора.
Private Const STATUS_NUMERICAL_FAILURE As String = "NumericalFailure" ' Solver не нашел равновесие в заданных ограничениях.
Private Const STATUS_NOT_CALCULATED As String = "NotCalculated" ' Расчетная проверка не относится к этому LC/режиму.
Private Const STATUS_STRENGTH_PASSED As String = "StrengthPassed" ' По lambda сечение несет заданный момент.
Private Const STATUS_STRENGTH_FAILED As String = "StrengthFailed" ' Lambda меньше 1: заданный момент выше найденной несущей.
Private Const CALC_GROUP1 As String = "Group1" ' Первая группа: проверка несущей способности.
Private Const CALC_GROUP2 As String = "Group2" ' Вторая группа: эксплуатационная проверка трещин.

Private mSection As CSectionModel ' Расчетная модель сечения, общая для всех LC.
Private mConcreteMaterial As Object ' Материал бетона с методами Stress/Tangent/Secant.
Private mSteelMaterial As Object ' Материал арматуры с тем же контрактом материала.
Private mCount As Long ' Сколько сочетаний реально добавлено в пакет.

Private mCombinationID(1 To MAX_COMBINATIONS) As String ' ID сочетания из таблицы нагрузок.
Private mN(1 To MAX_COMBINATIONS) As Double ' Продольная сила после переноса к координатам модели, Н.
Private mMx(1 To MAX_COMBINATIONS) As Double ' Mx после учета точки приложения нагрузки, Н*мм.
Private mMy(1 To MAX_COMBINATIONS) As Double ' My после учета точки приложения нагрузки, Н*мм.
Private mUserMx(1 To MAX_COMBINATIONS) As Double ' Mx как ввел пользователь относительно выбранной точки нагрузки.
Private mUserMy(1 To MAX_COMBINATIONS) As Double ' My как ввел пользователь относительно выбранной точки нагрузки.
Private mCalculationType(1 To MAX_COMBINATIONS) As String ' Group1/Group2: какая инженерная проверка нужна для LC.
Private mComment(1 To MAX_COMBINATIONS) As String ' Комментарий пользователя из таблицы сочетаний.
Private mInputValid(1 To MAX_COMBINATIONS) As Boolean ' False, если строка LC прочитана, но входные данные некорректны.
Private mInputError(1 To MAX_COMBINATIONS) As String ' Текст причины InvalidInput для конкретного LC.
Private mLoadReferenceX As Double ' X точки приложения нагрузки относительно расчетной модели, мм.
Private mLoadReferenceY As Double ' Y точки приложения нагрузки относительно расчетной модели, мм.

Private mStatus(1 To MAX_COMBINATIONS) As String ' Итоговый сводный статус LC.
Private mCapacityStatus(1 To MAX_COMBINATIONS) As String ' Статус поиска несущей способности.
Private mCapacityLimitState(1 To MAX_COMBINATIONS) As String ' Какой предел остановил capacity solver.
Private mStrengthCheckStatus(1 To MAX_COMBINATIONS) As String ' Простая проверка: держит/не держит по lambda.
Private mCrackStatus(1 To MAX_COMBINATIONS) As String ' Статус расчета трещин для второй группы.
Private mLambdaCapacity(1 To MAX_COMBINATIONS) As Double ' Найденный множитель несущей способности.
Private mUtilCapacity(1 To MAX_COMBINATIONS) As Double ' Использование по несущей способности, обычно 1/lambda.
Private mCrackWidth(1 To MAX_COMBINATIONS) As Double ' Расчетная ширина раскрытия трещины, мм.
Private mCrackAllowableValue(1 To MAX_COMBINATIONS) As Double ' Нормативный/пользовательский предел a_crc,ult, мм.
Private mCrackUtil(1 To MAX_COMBINATIONS) As Double ' Использование по ширине трещины.
Private mCrackFormed(1 To MAX_COMBINATIONS) As Boolean ' True, если при текущем LC трещина считается образованной.
Private mCrackPhi1(1 To MAX_COMBINATIONS) As Double ' Коэффициент длительности нагрузки в формуле a_crc.
Private mCrackPhi2(1 To MAX_COMBINATIONS) As Double ' Коэффициент профиля арматуры в формуле a_crc.
Private mCrackPhi3(1 To MAX_COMBINATIONS) As Double ' Коэффициент характера нагружения в формуле a_crc.
Private mCrackPsiS(1 To MAX_COMBINATIONS) As Double ' Коэффициент psi_s: 1 или уточненное значение Refined.
Private mCrackSigmaS(1 To MAX_COMBINATIONS) As Double ' Среднее напряжение выбранной растянутой арматуры, МПа.
Private mCrackSigmaSCrc(1 To MAX_COMBINATIONS) As Double ' Напряжение той же арматуры в состоянии lambda_crc, МПа.
Private mCrackAs(1 To MAX_COMBINATIONS) As Double ' Площадь выбранной растянутой арматуры As, мм2.
Private mCrackAbt(1 To MAX_COMBINATIONS) As Double ' Площадь растянутого бетона Abt, мм2.
Private mCrackDsEquivalent(1 To MAX_COMBINATIONS) As Double ' Эквивалентный диаметр стержней ds, мм.
Private mCrackLsRaw(1 To MAX_COMBINATIONS) As Double ' Неограниченное расстояние между трещинами ls, мм.
Private mCrackLs(1 To MAX_COMBINATIONS) As Double ' Расчетное ls после ограничений СП, мм.
Private mCrackLambdaCrc(1 To MAX_COMBINATIONS) As Double ' Множитель нагрузки в момент образования трещины.
Private mCrackNcrc(1 To MAX_COMBINATIONS) As Double ' Усилие образования трещины для центрального растяжения, Н.
Private mCrackTensionRebarCount(1 To MAX_COMBINATIONS) As Long ' Количество стержней, вошедших в As.
Private mCrackEffectiveZoneDepth(1 To MAX_COMBINATIONS) As Double ' Фактическая глубина зоны Abt, мм.
Private mCrackCentralBranch(1 To MAX_COMBINATIONS) As Boolean ' True, если использована ветка центрального растяжения.
Private mStrainSafetyFactor(1 To MAX_COMBINATIONS) As Double ' Ориентировочный запас по деформациям для DirectState.
Private mGoverningUtil(1 To MAX_COMBINATIONS) As Double ' Максимальное использование из доступных проверок LC.
Private mGoverningCombinationIndex As Long ' Индекс LC, который стал определяющим в batch summary.
Private mElapsedSeconds As Double ' Полное время пакетного расчета, сек.
Private mDiagnosticLog As String ' Короткий текстовый журнал по LC для отладки batch-потока.
Private mStateAvailable(1 To MAX_COMBINATIONS) As Boolean ' True, если для LC сохранена плоскость деформаций.
Private mStateConverged(1 To MAX_COMBINATIONS) As Boolean ' True, если прямой CSectionSolver сошелся.
Private mEpsilon0(1 To MAX_COMBINATIONS) As Double ' Свободный член плоскости деформаций LC.
Private mKappaX(1 To MAX_COMBINATIONS) As Double ' Кривизна при y в формуле epsilon = e0 + kx*y + ky*x.
Private mKappaY(1 To MAX_COMBINATIONS) As Double ' Кривизна при x в формуле epsilon = e0 + kx*y + ky*x.
Private mNint(1 To MAX_COMBINATIONS) As Double ' Внутренняя продольная сила из напряжений элементов, Н.
Private mMxint(1 To MAX_COMBINATIONS) As Double ' Внутренний Mx из напряжений элементов, Н*мм.
Private mMyint(1 To MAX_COMBINATIONS) As Double ' Внутренний My из напряжений элементов, Н*мм.
Private mResidualN(1 To MAX_COMBINATIONS) As Double ' Невязка Nint - N после прямого solve, Н.
Private mResidualMx(1 To MAX_COMBINATIONS) As Double ' Невязка Mxint - Mx после прямого solve, Н*мм.
Private mResidualMy(1 To MAX_COMBINATIONS) As Double ' Невязка Myint - My после прямого solve, Н*мм.
Private mStateIterations(1 To MAX_COMBINATIONS) As Long ' Число итераций прямого solver-а для LC.

Private mCrackAllowable As Double ' Допустимая ширина раскрытия трещины a_crc,ult, мм.
Private mCrackPsiMode As String ' Unity или Refined: как определять psi_s для нормальных трещин.
Private mCrackTensionZoneMode As String ' Effective или FullTension: какая бетонная зона входит в Abt.
Private mCrackRebarProfile As String ' Ribbed или Smooth: профиль арматуры для коэффициента phi2.
Private mCrackConcreteRbtSer As Double ' SLS-сопротивление бетона растяжению для eps_bt,crc и Ncrc, МПа.
Private mCrackConcreteEb As Double ' Модуль бетона для eps_bt,crc и приведенной площади Ared, МПа.
Private mCrackSteelEs As Double ' Модуль арматуры Es в формуле ширины раскрытия, МПа.
Private mCrackEnabled As Boolean ' Общий выключатель расчета трещин для сочетаний второй группы.
Private mCalculationMode As String ' DirectState или FullCapacity.
Private mSolverMethod As String ' Newton или Secant для прямого равновесия.
Private mSolverLoadSteps As Long ' Число ступеней приложения нагрузки в CSectionSolver.
Private mSolverMaxIterations As Long ' Лимит итераций прямого solver-а на одной ступени.
Private mSolverToleranceN As Double ' Допуск равновесия по N, Н.
Private mSolverToleranceMx As Double ' Допуск равновесия по Mx, Н*мм.
Private mSolverToleranceMy As Double ' Допуск равновесия по My, Н*мм.
Private mSolverLineSearchEnabled As Boolean ' Разрешает дробить итерационный шаг, если полный шаг ухудшает невязку.
Private mSolverDampingInitial As Double ' Начальный alpha для line search общего решателя НДС.
Private mSolverMinLineSearchAlpha As Double ' Минимальный alpha, ниже которого line search считается неудачным.
Private mSolverMaxDeltaEpsilon0 As Double ' Защита от слишком большого шага по epsilon0.
Private mSolverMaxDeltaKappa As Double ' Защита от слишком большого шага по кривизнам.
Private mSolverSecantMaxRestarts As Long ' Сколько раз Secant может перестроить приближенную матрицу.
Private mSolverSecantMinStepNorm As Double ' Минимальная норма шага Secant до признания застоя.
Private mCapacityInitialLambda As Double ' Начальный шаг/точка поиска lambda для LoadMultiplier.
Private mCapacityMethod As String ' LoadMultiplier или UltimateStrain.
Private mCapacitySearchMethod As String ' Bisection, Brent или Secant для LoadMultiplier.
Private mCapacityMaxLambda As Double ' Верхняя защитная граница lambda.
Private mCapacityToleranceLambda As Double ' Точность lambda в одномерном поиске LoadMultiplier.
Private mCapacityToleranceStrain As Double ' Точность попадания в предельную деформацию UltimateStrain.
Private mCapacityMaxRetries As Long ' Повторные попытки LoadMultiplier при несходимости пробной точки.
Private mCapacityBaseLoadSteps As Long ' Базовое число load steps для пробной точки capacity.
Private mCapacitySolverMaxIterations As Long ' Лимит итераций внутреннего solver-а при capacity-поиске.
Private mCapacityConcreteCompressionLimit As Double ' Предельная сжатая деформация бетона.
Private mCapacityConcreteTensionLimit As Double ' Предельная растянутая деформация бетона при UseDiagram.
Private mCapacitySteelStrainLimit As Double ' Предельная деформация арматуры по модулю.
Private mConcreteTensionMode As String ' Ignore или UseDiagram для растянутого бетона.
Private mReport As CExecutionReport ' Необязательный человекочитаемый отчет выполнения; расчетная логика от него не зависит.

' Инициализирует внутреннее состояние объекта значениями по умолчанию, безопасными для последующей настройки.
Private Sub Class_Initialize()
    mCrackAllowable = 0.3
    mCrackPsiMode = "Unity"
    mCrackTensionZoneMode = "Effective"
    mCrackRebarProfile = "Ribbed"
    mCrackConcreteRbtSer = 1.8
    mCrackConcreteEb = 32500#
    mCrackSteelEs = 200000#
    mCalculationMode = "FullCapacity"
    mSolverMethod = "Newton"
    mCrackEnabled = True
    mSolverLoadSteps = 8
    mSolverMaxIterations = 80
    mSolverToleranceN = 5#
    mSolverToleranceMx = 5000#
    mSolverToleranceMy = 5000#
    mSolverLineSearchEnabled = True
    mSolverDampingInitial = 1#
    mSolverMinLineSearchAlpha = 0.03125
    mSolverMaxDeltaEpsilon0 = 0.0005
    mSolverMaxDeltaKappa = 0.00001
    mSolverSecantMaxRestarts = 2
    mSolverSecantMinStepNorm = 0.000000000001
    mCapacityInitialLambda = 1#
    mCapacityMethod = "UltimateStrain"
    mCapacitySearchMethod = "Bisection"
    mCapacityMaxLambda = 64#
    mCapacityToleranceLambda = 0.01
    mCapacityToleranceStrain = 0.00001
    mCapacityMaxRetries = 0
    mCapacityBaseLoadSteps = 1
    mCapacitySolverMaxIterations = 60
    mCapacityConcreteCompressionLimit = -0.0035
    mCapacityConcreteTensionLimit = 0.00015
    mCapacitySteelStrainLimit = 0.025
    mConcreteTensionMode = "Ignore"
End Sub

' Принимает готовую модель и материалы перед добавлением/расчетом LC.
' После этого класс не обращается к геометрическим генераторам: все сочетания
' считаются на одной и той же CSectionModel.
Public Sub Initialize(ByVal section As CSectionModel, ByVal concreteMaterial As Object, ByVal steelMaterial As Object)
    If section Is Nothing Then Err.Raise vbObjectError + 3900, "CBatchSectionCalculator", "Модель сечения не передана."
    If section.ConcreteCount <= 0 Then Err.Raise vbObjectError + 3901, "CBatchSectionCalculator", "В модели сечения нет бетонных элементов."
    If concreteMaterial Is Nothing Then Err.Raise vbObjectError + 3902, "CBatchSectionCalculator", "Материал бетона не передан."
    If steelMaterial Is Nothing Then Err.Raise vbObjectError + 3903, "CBatchSectionCalculator", "Материал арматуры не передан."
    Set mSection = section
    Set mConcreteMaterial = concreteMaterial
    Set mSteelMaterial = steelMaterial
End Sub

' Подключает внешний отчет выполнения. Batch только добавляет события, но не
' открывает и не сохраняет файл: жизненным циклом отчета управляет Excel-макрос.
Public Property Set ExecutionReport(ByVal report As CExecutionReport)
    Set mReport = report
End Property

' Считывает численные настройки solver-ов и трещин из уже загруженного reader-а.
' Если передан CUnitSystem, пользовательские допуски в силе/моменте/длине
' сразу переводятся во внутренние единицы ядра.
Public Sub ApplySettings(ByVal settings As CSystemSettingsReader, Optional ByVal units As CUnitSystem = Nothing)
    If settings Is Nothing Then Exit Sub
    mCrackAllowable = settings.GetDouble("SLS.Crack.Allowable", mCrackAllowable)
    mCrackPsiMode = settings.GetString("SLS.Crack.PsiMode", mCrackPsiMode)
    mCrackTensionZoneMode = settings.GetString("SLS.Crack.TensionZoneMode", mCrackTensionZoneMode)
    mCrackRebarProfile = settings.GetString("Steel.RebarProfile", mCrackRebarProfile)
    mCrackConcreteRbtSer = settings.GetDouble("Concrete.Rbt.SLS", mCrackConcreteRbtSer)
    mCrackConcreteEb = settings.GetDouble("Concrete.Eb", mCrackConcreteEb)
    mCrackSteelEs = settings.GetDouble("Steel.Es", mCrackSteelEs)
    mCalculationMode = settings.GetString("Calculation.Mode", mCalculationMode)
    If settings.HasKey("Solver.Method") Then
        mSolverMethod = settings.GetRawString("Solver.Method", vbNullString)
    Else
        mSolverMethod = "Newton"
    End If
    mCrackEnabled = settings.GetBoolean("SLS.Crack.Enabled", mCrackEnabled)
    mSolverLoadSteps = settings.GetLong("Solver.LoadSteps", mSolverLoadSteps)
    mSolverMaxIterations = settings.GetLong("Solver.MaxIterations", mSolverMaxIterations)
    mSolverToleranceN = settings.GetDouble("Solver.ToleranceN", mSolverToleranceN)
    mSolverToleranceMx = settings.GetDouble("Solver.ToleranceMx", mSolverToleranceMx)
    mSolverToleranceMy = settings.GetDouble("Solver.ToleranceMy", mSolverToleranceMy)
    mSolverLineSearchEnabled = settings.GetBoolean("Solver.LineSearchEnabled", mSolverLineSearchEnabled)
    mSolverDampingInitial = settings.GetDouble("Solver.DampingInitial", mSolverDampingInitial)
    mSolverMinLineSearchAlpha = settings.GetDouble("Solver.MinLineSearchAlpha", mSolverMinLineSearchAlpha)
    mSolverMaxDeltaEpsilon0 = settings.GetDouble("Solver.MaxDeltaEpsilon0", mSolverMaxDeltaEpsilon0)
    mSolverMaxDeltaKappa = settings.GetDouble("Solver.MaxDeltaKappa", mSolverMaxDeltaKappa)
    mSolverSecantMaxRestarts = settings.GetLong("Solver.SecantMaxRestarts", mSolverSecantMaxRestarts)
    mSolverSecantMinStepNorm = settings.GetDouble("Solver.SecantMinStepNorm", mSolverSecantMinStepNorm)
    mCapacityInitialLambda = settings.GetDouble("Capacity.InitialLambda", mCapacityInitialLambda)
    mCapacityMethod = settings.GetString("Capacity.Method", mCapacityMethod)
    If settings.HasKey("Capacity.SearchMethod") Then
        mCapacitySearchMethod = settings.GetRawString("Capacity.SearchMethod", vbNullString)
    Else
        mCapacitySearchMethod = "Bisection"
    End If
    mCapacityMaxLambda = settings.GetDouble("Capacity.MaxLambda", mCapacityMaxLambda)
    mCapacityToleranceLambda = settings.GetDouble("Capacity.ToleranceLambda", mCapacityToleranceLambda)
    mCapacityToleranceStrain = settings.GetDouble("Capacity.ToleranceStrain", mCapacityToleranceStrain)
    mCapacityMaxRetries = settings.GetLong("Capacity.MaxRetries", mCapacityMaxRetries)
    mCapacityBaseLoadSteps = settings.GetLong("Capacity.BaseLoadSteps", mCapacityBaseLoadSteps)
    mCapacitySolverMaxIterations = settings.GetLong("Capacity.SolverMaxIterations", mCapacitySolverMaxIterations)
    mCapacityConcreteCompressionLimit = settings.GetDouble("Capacity.ConcreteCompressionLimit", mCapacityConcreteCompressionLimit)
    mCapacityConcreteTensionLimit = settings.GetDouble("Capacity.ConcreteTensionLimit", mCapacityConcreteTensionLimit)
    mCapacitySteelStrainLimit = settings.GetDouble("Capacity.SteelStrainLimit", mCapacitySteelStrainLimit)
    mConcreteTensionMode = settings.GetString("Concrete.TensionMode", mConcreteTensionMode)
    If Not units Is Nothing Then
        mCrackAllowable = units.InputLengthToInternal(mCrackAllowable)
        mCrackConcreteRbtSer = units.InputStressToInternal(mCrackConcreteRbtSer)
        mCrackConcreteEb = units.InputStressToInternal(mCrackConcreteEb)
        mCrackSteelEs = units.InputStressToInternal(mCrackSteelEs)
        mSolverToleranceN = Abs(units.InputForceToInternal(mSolverToleranceN))
        mSolverToleranceMx = Abs(units.InputMomentMxToInternal(mSolverToleranceMx))
        mSolverToleranceMy = Abs(units.InputMomentMyToInternal(mSolverToleranceMy))
        mSolverMaxDeltaKappa = Abs(units.InputCurvatureToInternal(mSolverMaxDeltaKappa))
    End If
End Sub

' Очищает список LC и накопленные результаты.
' Используется перед чтением новой таблицы сочетаний, чтобы старые статусы не
' попали в новый batch summary.
Public Sub ClearCombinations()
    mCount = 0
    mGoverningCombinationIndex = 0
    mElapsedSeconds = 0#
    mDiagnosticLog = vbNullString
    mLoadReferenceX = 0#
    mLoadReferenceY = 0#
End Sub

' Добавляет корректно прочитанное сочетание нагрузок.
' Усилия уже должны быть во внутренних единицах и в INTERNAL-системе знаков.
Public Sub AddCombination(ByVal combinationID As String, ByVal nValue As Double, ByVal mxValue As Double, _
        ByVal myValue As Double, ByVal calculationType As String, ByVal comment As String)
    AddCombinationInternal combinationID, nValue, mxValue, myValue, calculationType, comment, True, vbNullString
End Sub

' Сохраняет строку LC с ошибкой ввода, чтобы пользователь увидел ее в отчете.
' Такие сочетания не передаются в solver, но отображаются как InvalidInput.
Public Sub AddInvalidCombination(ByVal combinationID As String, ByVal calculationType As String, _
        ByVal comment As String, ByVal reason As String)
    AddCombinationInternal combinationID, 0#, 0#, 0#, calculationType, comment, False, reason
End Sub

Private Sub AddCombinationInternal(ByVal combinationID As String, ByVal nValue As Double, ByVal mxValue As Double, _
        ByVal myValue As Double, ByVal calculationType As String, ByVal comment As String, _
        ByVal inputValid As Boolean, ByVal inputError As String)
    If mCount >= MAX_COMBINATIONS Then Err.Raise vbObjectError + 3911, "CBatchSectionCalculator", "Поддерживается не более 20 сочетаний нагрузок."
    mCount = mCount + 1
    If Len(Trim$(combinationID)) = 0 Then combinationID = "ROW" & CStr(mCount)
    mCombinationID(mCount) = combinationID
    mN(mCount) = nValue
    mMx(mCount) = mxValue
    mMy(mCount) = myValue
    mCalculationType(mCount) = calculationType
    mComment(mCount) = comment
    mUserMx(mCount) = mxValue
    mUserMy(mCount) = myValue
    mInputValid(mCount) = inputValid
    mInputError(mCount) = inputError
End Sub

' Переносит пользовательские моменты из точки приложения нагрузки к координатам элементов.
' Пользователь задает Mx/My относительно выбранной точки нагрузки, а solver
' должен получить моменты относительно внутренней системы координат CSectionModel.
Public Sub ApplyLoadReference(ByVal referenceX As Double, ByVal referenceY As Double)
    mLoadReferenceX = referenceX
    mLoadReferenceY = referenceY

    Dim i As Long
    ' Эквивалентный перенос момента от продольной силы:
    ' Mx_model = Mx_user + N * y_ref, My_model = My_user + N * x_ref.
    For i = 1 To mCount
        If mInputValid(i) Then
            mMx(i) = mUserMx(i) + mN(i) * referenceY
            mMy(i) = mUserMy(i) + mN(i) * referenceX
        End If
    Next i
End Sub

' Последовательно считает все добавленные LC и выбирает определяющее сочетание.
' Порядок один для всех геометрий: capacity при необходимости, прямое НДС,
' затем трещины для сочетаний второй группы.
Public Sub Execute()
    ValidateReady
    Dim t0 As Double
    t0 = Timer
    mGoverningCombinationIndex = 0
    mDiagnosticLog = vbNullString
    ReportStep "Начат пакетный расчет. Количество сочетаний: " & CStr(mCount) & "."

    Dim i As Long
    ' Каждый LC изолирован: ошибка одного сочетания не должна стирать результаты
    ' соседних строк, но попадает в итоговый статус и batch summary.
    For i = 1 To mCount
        ClearResultRow i
        RunCombination i
        If mInputValid(i) Then
            If mGoverningCombinationIndex = 0 Then
                mGoverningCombinationIndex = i
            ElseIf IsBetterGoverningCandidate(i, mGoverningCombinationIndex) Then
                mGoverningCombinationIndex = i
            End If
        End If
    Next i
    mElapsedSeconds = Timer - t0
    ReportStep "Пакетный расчет завершен. Время: " & FormatNumberInvariant(mElapsedSeconds) & " с."
End Sub

' Считает одно сочетание нагрузок и собирает итоговый статус.
' Внутри сохраняется разделение: capacity отвечает за запас по несущей,
' прямой solve дает фактическое НДС, crack запускается только для 2 группы.
Private Sub RunCombination(ByVal index As Long)
    On Error GoTo Failed
    ReportSection "Сочетание " & CStr(index) & ": " & mCombinationID(index)
    ReportValue "Комментарий", mComment(index)
    If Not mInputValid(index) Then
        ReportStep "Сочетание не передано в solver из-за ошибки входных данных: " & mInputError(index)
        MarkInvalid index, mInputError(index)
        Exit Sub
    End If
    If Not IsSupportedCalculationMode() Then
        ReportStep "Сочетание не рассчитано: неизвестный Calculation.Mode."
        MarkInvalid index, "Calculation.Mode должен быть DirectState или FullCapacity."
        Exit Sub
    End If

    Dim calculationType As String
    calculationType = NormalizedCombinationCalculationType(mCalculationType(index))
    ReportValue "Тип проверки", IIf(Len(calculationType) > 0, calculationType, mCalculationType(index))
    ReportValue "N", FormatNumberInvariant(mN(index)) & " Н"
    ReportValue "Mx расчетный", FormatNumberInvariant(mMx(index)) & " Н*мм"
    ReportValue "My расчетный", FormatNumberInvariant(mMy(index)) & " Н*мм"
    If Len(calculationType) = 0 Then
        ReportStep "Сочетание не рассчитано: CalculationType должен быть Group1 или Group2."
        MarkInvalid index, "CalculationType must be Group1 or Group2."
        Exit Sub
    End If

    mStatus(index) = STATUS_OK
    If ShouldRunCapacity(calculationType) And IsFullCapacityMode Then
        ReportStep "Запускается проверка несущей способности для сочетания I группы."
        RunCapacity index
    Else
        ReportStep "Проверка несущей способности не требуется для этого сочетания или режима Calculation.Mode."
        StoreNotCalculatedCapacity index
    End If
    RunSectionStateAndCrack index, ShouldRunCrack(calculationType)

    mGoverningUtil(index) = MaxDouble(MaxDouble(mUtilCapacity(index), StrainUtilization(index)), mCrackUtil(index))
    mStatus(index) = AggregateStatus(index)
    AppendDiagnostic index
    ReportCombinationSummary index
    Exit Sub

Failed:
    ReportStep "Ошибка при расчете сочетания: " & Err.Description
    MarkInvalid index, Err.Description
End Sub

' Запускает поиск несущей способности для LC первой группы.
' Если моментный вектор Mx/My нулевой, множитель lambda не имеет смысла, и
' capacity-проверка помечается как пропущенная без ошибки solver-а.
Private Sub RunCapacity(ByVal index As Long)
    If Sqr(mUserMx(index) * mUserMx(index) + mUserMy(index) * mUserMy(index)) <= 0.000000001 Then
        ReportStep "Capacity не запускается: моментный вектор Mx/My равен нулю."
        StoreSkippedCapacity index
        Exit Sub
    End If

    Dim cap As CCapacitySolver
    Set cap = New CCapacitySolver
    ConfigureCapacity cap
    Select Case NormalizedCapacityMethod()
        Case "UltimateStrain"
            cap.SolveByUltimateStrain mSection, mConcreteMaterial, mSteelMaterial, _
                mN(index), mUserMx(index), mUserMy(index), mN(index) * mLoadReferenceY, mN(index) * mLoadReferenceX
        Case "LoadMultiplier"
            cap.SolveByLoadMultiplier mSection, mConcreteMaterial, mSteelMaterial, _
                mN(index), mUserMx(index), mUserMy(index), mN(index) * mLoadReferenceY, mN(index) * mLoadReferenceX
        Case Else
            mCapacityStatus(index) = STATUS_INVALID_INPUT & ": Capacity.Method должен быть LoadMultiplier или UltimateStrain."
            mCapacityLimitState(index) = STATUS_INVALID_INPUT
            mStrengthCheckStatus(index) = STATUS_INVALID_INPUT
            ReportStep "Capacity остановлен: неизвестный Capacity.Method."
            Exit Sub
    End Select
    StoreCapacity index, cap
    ReportValue "CapacityStatus", mCapacityStatus(index)
    ReportValue "StrengthCheckStatus", mStrengthCheckStatus(index)
    ReportValue "LambdaCapacity", FormatNumberInvariant(mLambdaCapacity(index))
    ReportBlock "Диагностика поиска несущей способности", cap.DiagnosticLog
    If Not cap.LastSolver Is Nothing Then
        ReportBlock "Итерации внутреннего CSectionSolver для последней capacity-точки", cap.LastSolver.DiagnosticLog
    End If
End Sub

' Выполняет прямой расчет НДС и, если нужно, расчет ширины трещин.
' Даже в FullCapacity этот solve нужен отдельно: он показывает состояние
' именно от заданного LC, а не только предельное состояние capacity solver-а.
Private Sub RunSectionStateAndCrack(ByVal index As Long, ByVal crackRequested As Boolean)
    Dim service As CSectionSolver
    Set service = New CSectionSolver
    ConfigureSectionSolver service
    ReportStep "Запускается прямой расчет НДС CSectionSolver."
    service.Solve mSection, mConcreteMaterial, mSteelMaterial, mN(index), mMx(index), mMy(index)
    StoreSectionState index, service
    mStrainSafetyFactor(index) = CalculateStrainSafetyFactor(service)
    ReportValue "НДС сошелся", IIf(service.Converged, "да", "нет")
    ReportValue "Итераций НДС", CStr(service.Iterations)
    ReportBlock "Итерации прямого CSectionSolver", service.DiagnosticLog
    If Not service.Converged Then
        mCrackStatus(index) = STATUS_NUMERICAL_FAILURE & ": " & service.StopReason
        ReportStep "Расчет трещин не запускается, потому что прямое НДС не сошлось: " & service.StopReason
        Exit Sub
    End If

    If Not crackRequested Or Not mCrackEnabled Then
        mCrackStatus(index) = STATUS_NOT_CALCULATED
        If crackRequested Then
            ReportStep "Расчет трещин отключен настройкой SLS.Crack.Enabled."
        Else
            ReportStep "Расчет трещин не требуется для этого типа сочетания."
        End If
        Exit Sub
    End If

    Dim crack As CCrackWidthCalculator
    Set crack = New CCrackWidthCalculator
    ConfigureCrackCalculator crack
    ReportStep "Запускается расчет ширины раскрытия нормальных трещин."
    crack.Calculate service, mSection, mConcreteMaterial, mSteelMaterial, mN(index), mMx(index), mMy(index), _
        mUserMx(index), mUserMy(index)
    mCrackStatus(index) = IIf(crack.Converged, STATUS_OK, crack.StopReason)
    mCrackWidth(index) = crack.CrackWidth
    mCrackAllowableValue(index) = crack.AllowableCrackWidth
    mCrackUtil(index) = crack.Utilization
    mCrackFormed(index) = crack.CrackFormed
    mCrackPhi1(index) = crack.Phi1
    mCrackPhi2(index) = crack.Phi2
    mCrackPhi3(index) = crack.Phi3
    mCrackPsiS(index) = crack.PsiS
    mCrackSigmaS(index) = crack.SigmaS
    mCrackSigmaSCrc(index) = crack.SigmaSCrc
    mCrackAs(index) = crack.AsTension
    mCrackAbt(index) = crack.Abt
    mCrackDsEquivalent(index) = crack.DsEquivalent
    mCrackLsRaw(index) = crack.CrackSpacingRaw
    mCrackLs(index) = crack.CrackSpacing
    mCrackLambdaCrc(index) = crack.LambdaCrc
    mCrackNcrc(index) = crack.Ncrc
    mCrackTensionRebarCount(index) = crack.TensionRebarCount
    mCrackEffectiveZoneDepth(index) = crack.EffectiveZoneDepth
    mCrackCentralBranch(index) = crack.CentralTensionBranch
    ReportValue "CrackStatus", mCrackStatus(index)
    ReportValue "Трещина образована", IIf(mCrackFormed(index), "да", "нет")
    ReportValue "a_crc", FormatNumberInvariant(mCrackWidth(index)) & " мм"
    ReportValue "psi_s", FormatNumberInvariant(mCrackPsiS(index))
    ReportValue "As", FormatNumberInvariant(mCrackAs(index)) & " мм2"
    ReportValue "Abt", FormatNumberInvariant(mCrackAbt(index)) & " мм2"
    ReportValue "ds_eq", FormatNumberInvariant(mCrackDsEquivalent(index)) & " мм"
    ReportValue "ls", FormatNumberInvariant(mCrackLs(index)) & " мм"
    ReportBlock "Диагностика расчета трещин", crack.DiagnosticLog
End Sub

Private Function NormalizedCombinationCalculationType(ByVal calculationType As String) As String
    Select Case LCase$(Trim$(calculationType))
        Case "group1", "1", "uls", "strength", "strengthonly"
            NormalizedCombinationCalculationType = CALC_GROUP1
        Case "group2", "2", "sls", "crack", "crackonly"
            NormalizedCombinationCalculationType = CALC_GROUP2
    End Select
End Function

' Показывает, должна ли для LC запускаться несущая способность.
' По принятой логике проекта это относится только к сочетаниям первой группы.
Private Function ShouldRunCapacity(ByVal calculationType As String) As Boolean
    ShouldRunCapacity = (calculationType = CALC_GROUP1)
End Function

' Показывает, должна ли для LC запускаться проверка трещин.
' Текущая реализация считает только продолжительное раскрытие нормальных
' трещин для второй группы; сочетания первой группы сюда не попадают.
Private Function ShouldRunCrack(ByVal calculationType As String) As Boolean
    ShouldRunCrack = (calculationType = CALC_GROUP2)
End Function

' True, когда выбран полный расчет с поиском несущей способности.
Private Function IsFullCapacityMode() As Boolean
    IsFullCapacityMode = (StrComp(mCalculationMode, "FullCapacity", vbTextCompare) = 0)
End Function

' True, когда пользователь запросил только фактическое НДС без capacity-поиска.
Private Function IsDirectStateMode() As Boolean
    IsDirectStateMode = (StrComp(mCalculationMode, "DirectState", vbTextCompare) = 0)
End Function

' Защищает RunCombination от неизвестных режимов Calculation.Mode.
Private Function IsSupportedCalculationMode() As Boolean
    IsSupportedCalculationMode = IsDirectStateMode() Or IsFullCapacityMode()
End Function

' Приводит пользовательскую строку Capacity.Method к одному из двух
' внутренних имен. Пустой результат дальше превращается в InvalidInput.
Private Function NormalizedCapacityMethod() As String
    Select Case LCase$(Trim$(mCapacityMethod))
        Case "loadmultiplier"
            NormalizedCapacityMethod = "LoadMultiplier"
        Case "ultimatestrain"
            NormalizedCapacityMethod = "UltimateStrain"
    End Select
End Function

' Заполняет capacity-поля нейтральным статусом, когда эта проверка не должна
' выполняться для выбранной группы/режима.
Private Sub StoreNotCalculatedCapacity(ByVal index As Long)
    mCapacityStatus(index) = STATUS_NOT_CALCULATED
    mCapacityLimitState(index) = STATUS_NOT_CALCULATED
    mStrengthCheckStatus(index) = STATUS_NOT_CALCULATED
End Sub

' Помечает capacity как пропущенную из-за нулевого моментного вектора.
' Это не ошибка исходных данных: для чистого N нет направления момента,
' которое можно масштабировать через lambda.
Private Sub StoreSkippedCapacity(ByVal index As Long)
    mCapacityStatus(index) = STATUS_SKIPPED
    mCapacityLimitState(index) = STATUS_SKIPPED
    mStrengthCheckStatus(index) = STATUS_NOT_CALCULATED
End Sub

' Переносит результат CCapacitySolver в массивы batch-отчета.
' Здесь разделяется математическая сходимость capacity solver-а и простая
' инженерная проверка: lambda меньше или больше единицы.
Private Sub StoreCapacity(ByVal index As Long, ByVal cap As CCapacitySolver)
    Dim statusText As String
    If cap.Converged Then
        statusText = STATUS_OK
    Else
        statusText = cap.LimitState
        If Len(cap.StopReason) > 0 Then statusText = statusText & ": " & cap.StopReason
    End If

    mCapacityStatus(index) = statusText
    mCapacityLimitState(index) = cap.LimitState
    mLambdaCapacity(index) = cap.LambdaUltimate
    If cap.LambdaUltimate > 0# Then mUtilCapacity(index) = 1# / cap.LambdaUltimate
    If cap.Converged Then mUtilCapacity(index) = cap.Utilization
    mStrengthCheckStatus(index) = StrengthStatusFromLambda(cap)
End Sub

' Фиксирует InvalidInput по конкретному LC и не дает этому сочетанию пройти
' в solver. Так пользователь видит причину в таблице, а остальные LC считаются.
Private Sub MarkInvalid(ByVal index As Long, ByVal reason As String)
    If Len(Trim$(reason)) = 0 Then reason = "Некорректное сочетание нагрузок."
    mStatus(index) = STATUS_INVALID_INPUT & ": " & reason
    mCapacityStatus(index) = STATUS_INVALID_INPUT
    mCapacityLimitState(index) = STATUS_INVALID_INPUT
    mStrengthCheckStatus(index) = STATUS_INVALID_INPUT
    mCrackStatus(index) = STATUS_INVALID_INPUT
    mGoverningUtil(index) = 0#
    AppendDiagnostic index
End Sub

' Передает в CCapacitySolver ровно те настройки, которые относятся к поиску
' несущей способности и внутренним пробным расчетам равновесия.
Private Sub ConfigureCapacity(ByVal cap As CCapacitySolver)
    cap.ConcreteCompressionLimit = mCapacityConcreteCompressionLimit
    cap.ConcreteTensionLimit = mCapacityConcreteTensionLimit
    cap.SteelStrainLimit = mCapacitySteelStrainLimit
    cap.LambdaTolerance = mCapacityToleranceLambda
    cap.StrainTolerance = mCapacityToleranceStrain
    cap.SearchMethod = mCapacitySearchMethod
    cap.MaxLambda = mCapacityMaxLambda
    cap.InitialLambdaStep = mCapacityInitialLambda
    cap.MaxRetries = mCapacityMaxRetries
    cap.SolverBaseLoadSteps = mCapacityBaseLoadSteps
    cap.SolverMaxIterations = mCapacitySolverMaxIterations
    cap.SolverToleranceN = mSolverToleranceN
    cap.SolverToleranceMx = mSolverToleranceMx
    cap.SolverToleranceMy = mSolverToleranceMy
    cap.SolverMethod = mSolverMethod
    cap.SolverLineSearchEnabled = mSolverLineSearchEnabled
    cap.SolverDampingInitial = mSolverDampingInitial
    cap.SolverMinLineSearchAlpha = mSolverMinLineSearchAlpha
    cap.SolverMaxDeltaEpsilon0 = mSolverMaxDeltaEpsilon0
    cap.SolverMaxDeltaKappa = mSolverMaxDeltaKappa
    cap.SolverSecantMaxRestarts = mSolverSecantMaxRestarts
    cap.SolverSecantMinStepNorm = mSolverSecantMinStepNorm
End Sub

' Настраивает прямой CSectionSolver для расчета фактического состояния LC.
' Capacity-настройки сюда не попадают, чтобы прямое НДС не зависело от
' выбранного метода поиска предельного состояния.
Private Sub ConfigureSectionSolver(ByVal solver As CSectionSolver)
    solver.LoadSteps = mSolverLoadSteps
    solver.SolverMethod = mSolverMethod
    solver.MaxIterations = mSolverMaxIterations
    solver.ToleranceN = mSolverToleranceN
    solver.ToleranceMx = mSolverToleranceMx
    solver.ToleranceMy = mSolverToleranceMy
    solver.LineSearchEnabled = mSolverLineSearchEnabled
    solver.DampingInitial = mSolverDampingInitial
    solver.MinLineSearchAlpha = mSolverMinLineSearchAlpha
    solver.MaxDeltaEpsilon0 = mSolverMaxDeltaEpsilon0
    solver.MaxDeltaKappa = mSolverMaxDeltaKappa
    solver.SecantMaxRestarts = mSolverSecantMaxRestarts
    solver.SecantMinStepNorm = mSolverSecantMinStepNorm
End Sub

' Передает в расчет трещин те же настройки прямого равновесия, что и в
' сервисный solve LC. Это важно для сопоставимости: состояние с раскрытой
' трещиной ищется тем же CSectionSolver, только с принудительно выключенным
' растянутым бетоном.
Private Sub ConfigureCrackCalculator(ByVal crack As CCrackWidthCalculator)
    crack.AllowableCrackWidth = mCrackAllowable
    crack.PsiMode = mCrackPsiMode
    crack.TensionZoneMode = mCrackTensionZoneMode
    crack.RebarProfile = mCrackRebarProfile
    crack.ConcreteTensionLimitStress = mCrackConcreteRbtSer
    crack.ConcreteElasticModulus = mCrackConcreteEb
    crack.SteelElasticModulus = mCrackSteelEs
    crack.SolverLoadSteps = mSolverLoadSteps
    crack.SolverMethod = mSolverMethod
    crack.SolverMaxIterations = mSolverMaxIterations
    crack.SolverToleranceN = mSolverToleranceN
    crack.SolverToleranceMx = mSolverToleranceMx
    crack.SolverToleranceMy = mSolverToleranceMy
    crack.SolverLineSearchEnabled = mSolverLineSearchEnabled
    crack.SolverDampingInitial = mSolverDampingInitial
    crack.SolverMinLineSearchAlpha = mSolverMinLineSearchAlpha
    crack.SolverMaxDeltaEpsilon0 = mSolverMaxDeltaEpsilon0
    crack.SolverMaxDeltaKappa = mSolverMaxDeltaKappa
    crack.SolverSecantMaxRestarts = mSolverSecantMaxRestarts
    crack.SolverSecantMinStepNorm = mSolverSecantMinStepNorm
End Sub

' Проверяет, что batch готов к расчету.
' Это последний защитный барьер перед запуском solver-ов: модель и хотя бы
' одно сочетание должны быть уже загружены.
Private Sub ValidateReady()
    If mSection Is Nothing Then Err.Raise vbObjectError + 3920, "CBatchSectionCalculator", "Пакетный расчетчик не инициализирован моделью сечения."
    If mCount <= 0 Then Err.Raise vbObjectError + 3921, "CBatchSectionCalculator", "Не задано ни одного сочетания нагрузок."
End Sub

' Сбрасывает все результаты одной строки перед повторным расчетом LC.
' Входные усилия и комментарий не трогаются.
Private Sub ClearResultRow(ByVal index As Long)
    mStatus(index) = vbNullString
    mCapacityStatus(index) = vbNullString
    mCapacityLimitState(index) = vbNullString
    mStrengthCheckStatus(index) = vbNullString
    mCrackStatus(index) = vbNullString
    mLambdaCapacity(index) = 0#
    mUtilCapacity(index) = 0#
    mCrackWidth(index) = 0#: mCrackAllowableValue(index) = 0#: mCrackUtil(index) = 0#
    mCrackFormed(index) = False
    mCrackPhi1(index) = 0#: mCrackPhi2(index) = 0#: mCrackPhi3(index) = 0#: mCrackPsiS(index) = 0#
    mCrackSigmaS(index) = 0#: mCrackSigmaSCrc(index) = 0#
    mCrackAs(index) = 0#: mCrackAbt(index) = 0#: mCrackDsEquivalent(index) = 0#
    mCrackLsRaw(index) = 0#: mCrackLs(index) = 0#: mCrackLambdaCrc(index) = 0#: mCrackNcrc(index) = 0#
    mCrackTensionRebarCount(index) = 0#: mCrackEffectiveZoneDepth(index) = 0#: mCrackCentralBranch(index) = False
    mStrainSafetyFactor(index) = 0#: mGoverningUtil(index) = 0#
    mStateAvailable(index) = False
    mStateConverged(index) = False
    mEpsilon0(index) = 0#: mKappaX(index) = 0#: mKappaY(index) = 0#
    mNint(index) = 0#: mMxint(index) = 0#: mMyint(index) = 0#
    mResidualN(index) = 0#: mResidualMx(index) = 0#: mResidualMy(index) = 0#
    mStateIterations(index) = 0
End Sub

' Сохраняет найденную плоскость деформаций и внутренние усилия прямого solve.
' Эти данные затем уходят в Results и используются для схемы/AutoCAD export.
Private Sub StoreSectionState(ByVal index As Long, ByVal solver As CSectionSolver)
    mStateAvailable(index) = True
    mStateConverged(index) = solver.Converged
    mEpsilon0(index) = solver.Epsilon0
    mKappaX(index) = solver.KappaX
    mKappaY(index) = solver.KappaY
    mNint(index) = solver.Nint
    mMxint(index) = solver.Mxint
    mMyint(index) = solver.Myint
    mResidualN(index) = solver.ResidualN
    mResidualMx(index) = solver.ResidualMx
    mResidualMy(index) = solver.ResidualMy
    mStateIterations(index) = solver.Iterations
End Sub

' Добавляет одну строку в текстовую диагностику batch-расчета.
' Это компактная сводка для отладки, а не подробный журнал итераций solver-а.
Private Sub AppendDiagnostic(ByVal index As Long)
    mDiagnosticLog = mDiagnosticLog & "combination=" & mCombinationID(index) & _
        "; status=" & mStatus(index) & _
        "; CapacityStatus=" & mCapacityStatus(index) & _
        "; StrengthCheckStatus=" & mStrengthCheckStatus(index) & _
        "; CrackStatus=" & mCrackStatus(index) & _
        "; governingUtil=" & FormatNumberInvariant(mGoverningUtil(index)) & vbCrLf
End Sub

' Добавляет в подробный txt-отчет краткий итог по текущему LC.
' Batch summary остается основным табличным выводом, а здесь пользователь
' видит ход выполнения без открытия Results и без чтения внутренних массивов.
Private Sub ReportCombinationSummary(ByVal index As Long)
    ReportValue "Итоговый Status", mStatus(index)
    ReportValue "CapacityStatus", mCapacityStatus(index)
    ReportValue "StrengthCheckStatus", mStrengthCheckStatus(index)
    ReportValue "CrackStatus", mCrackStatus(index)
    ReportValue "GoverningUtilization", FormatNumberInvariant(mGoverningUtil(index))
End Sub

Private Sub ReportSection(ByVal title As String)
    If mReport Is Nothing Then Exit Sub
    If Not mReport.Enabled Then Exit Sub
    mReport.AddSection title
End Sub

Private Sub ReportStep(ByVal text As String)
    If mReport Is Nothing Then Exit Sub
    If Not mReport.Enabled Then Exit Sub
    mReport.AddStep text
End Sub

Private Sub ReportValue(ByVal label As String, ByVal valueText As String)
    If mReport Is Nothing Then Exit Sub
    If Not mReport.Enabled Then Exit Sub
    mReport.AddValue label, valueText
End Sub

Private Sub ReportBlock(ByVal title As String, ByVal body As String)
    If mReport Is Nothing Then Exit Sub
    If Not mReport.Enabled Then Exit Sub
    mReport.AddBlock title, body
End Sub

' Собирает пользовательский Status из статусов отдельных проверок.
' Важное правило: статус capacity "OK: предел найден" сам по себе не значит,
' что заданная нагрузка безопасна; это отдельно отражает StrengthCheckStatus.
Private Function AggregateStatus(ByVal index As Long) As String
    If mStrengthCheckStatus(index) = STATUS_STRENGTH_FAILED Then
        AggregateStatus = STATUS_STRENGTH_FAILED
    ElseIf mStrengthCheckStatus(index) = STATUS_STRENGTH_PASSED Then
        If IsNumericalStatus(mCrackStatus(index)) Then
            AggregateStatus = "CrackNumericalFailure"
        Else
            AggregateStatus = STATUS_OK
        End If
    ElseIf IsInvalidStatus(mCapacityStatus(index)) Or IsInvalidStatus(mStrengthCheckStatus(index)) Then
        AggregateStatus = mCapacityStatus(index)
    ElseIf StrainUtilization(index) > 1# Then
        AggregateStatus = "StrainLimitExceeded"
    ElseIf IsNumericalStatus(mCapacityStatus(index)) Then
        AggregateStatus = mCapacityStatus(index)
    ElseIf IsNumericalStatus(mCrackStatus(index)) Then
        AggregateStatus = mCrackStatus(index)
    Else
        AggregateStatus = STATUS_OK
    End If
End Function

' Распознает статусы, означающие именно численную несходимость.
Private Function IsNumericalStatus(ByVal statusText As String) As Boolean
    IsNumericalStatus = (InStr(1, statusText, STATUS_NUMERICAL_FAILURE, vbTextCompare) > 0 Or _
        InStr(1, statusText, "SingularTangent", vbTextCompare) > 0)
End Function

' Распознает статусы ошибок исходных данных.
Private Function IsInvalidStatus(ByVal statusText As String) As Boolean
    IsInvalidStatus = (InStr(1, statusText, STATUS_INVALID_INPUT, vbTextCompare) > 0)
End Function

' Переводит lambdaUltimate в простую инженерную проверку.
' Lambda < 1 означает, что заданный моментный вектор превышает найденную
' несущую способность при принятой постановке.
Private Function StrengthStatusFromLambda(ByVal cap As CCapacitySolver) As String
    If Not cap.Converged Then
        StrengthStatusFromLambda = STATUS_NOT_CALCULATED
    ElseIf cap.LambdaUltimate <= 0# Then
        StrengthStatusFromLambda = STATUS_NOT_CALCULATED
    ElseIf cap.LambdaUltimate < 1# Then
        StrengthStatusFromLambda = STATUS_STRENGTH_FAILED
    Else
        StrengthStatusFromLambda = STATUS_STRENGTH_PASSED
    End If
End Function

' Утилитарный максимум без WorksheetFunction, чтобы не тащить Excel в hot path.
Private Function MaxDouble(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then MaxDouble = a Else MaxDouble = b
End Function

' Считает ориентировочный запас по деформациям для DirectState.
' Дополнительный solve не выполняется: используется уже найденное состояние LC.
Private Function CalculateStrainSafetyFactor(ByVal solver As CSectionSolver) As Double
    Dim maxUtilization As Double
    maxUtilization = 0#

    ' Сжатый бетон сравнивается с заданным предельным сжатием.
    If mCapacityConcreteCompressionLimit < 0# And solver.MinConcreteStrain < 0# Then
        maxUtilization = MaxDouble(maxUtilization, Abs(solver.MinConcreteStrain) / Abs(mCapacityConcreteCompressionLimit))
    End If

    ' Растянутый бетон проверяется только когда он действительно участвует
    ' в материалной диаграмме UseDiagram.
    If StrComp(mConcreteTensionMode, "UseDiagram", vbTextCompare) = 0 Then
        If mCapacityConcreteTensionLimit > 0# And solver.MaxConcreteStrain > 0# Then
            maxUtilization = MaxDouble(maxUtilization, solver.MaxConcreteStrain / mCapacityConcreteTensionLimit)
        End If
    End If

    ' Для арматуры используется максимальная деформация по модулю.
    If mCapacitySteelStrainLimit > 0# And solver.MaxAbsSteelStrain > 0# Then
        maxUtilization = MaxDouble(maxUtilization, solver.MaxAbsSteelStrain / mCapacitySteelStrainLimit)
    End If

    If maxUtilization > 0# Then CalculateStrainSafetyFactor = 1# / maxUtilization
End Function

' Преобразует запас по деформации в использование, чтобы его можно было
' сравнивать с capacity/crack utilization при выборе худшего LC.
Private Function StrainUtilization(ByVal index As Long) As Double
    If mStrainSafetyFactor(index) > 0# Then StrainUtilization = 1# / mStrainSafetyFactor(index)
End Function

' Выбирает худшее сочетание для batch summary.
' Сначала сравнивается lambda, если она есть у обоих LC; для DirectState или
' не рассчитанной capacity используется общий governing utilization.
Private Function IsBetterGoverningCandidate(ByVal candidateIndex As Long, ByVal currentIndex As Long) As Boolean
    Dim candidateLambda As Double
    Dim currentLambda As Double
    candidateLambda = mLambdaCapacity(candidateIndex)
    currentLambda = mLambdaCapacity(currentIndex)

    If candidateLambda > 0# And currentLambda > 0# Then
        IsBetterGoverningCandidate = (candidateLambda < currentLambda)
    ElseIf candidateLambda > 0# Then
        IsBetterGoverningCandidate = True
    ElseIf currentLambda > 0# Then
        IsBetterGoverningCandidate = False
    Else
        IsBetterGoverningCandidate = (mGoverningUtil(candidateIndex) > mGoverningUtil(currentIndex))
    End If
End Function

' Форматирует число с точкой как десятичным разделителем для диагностической строки.
Private Function FormatNumberInvariant(ByVal value As Double) As String
    FormatNumberInvariant = Replace$(Format$(value, "0.############"), ",", ".")
End Function

Public Property Get Count() As Long
    Count = mCount
End Property

Public Property Get CombinationID(ByVal index As Long) As String
    CombinationID = mCombinationID(index)
End Property

Public Property Get CombinationName(ByVal index As Long) As String
    CombinationName = mComment(index)
End Property

Public Property Get CalculationType(ByVal index As Long) As String
    CalculationType = mCalculationType(index)
End Property

Public Property Get N(ByVal index As Long) As Double
    N = mN(index)
End Property

Public Property Get Mx(ByVal index As Long) As Double
    Mx = mMx(index)
End Property

Public Property Get My(ByVal index As Long) As Double
    My = mMy(index)
End Property

Public Property Get UserMx(ByVal index As Long) As Double
    UserMx = mUserMx(index)
End Property

Public Property Get UserMy(ByVal index As Long) As Double
    UserMy = mUserMy(index)
End Property

Public Property Get LoadReferenceX() As Double
    LoadReferenceX = mLoadReferenceX
End Property

Public Property Get LoadReferenceY() As Double
    LoadReferenceY = mLoadReferenceY
End Property

Public Property Get Status(ByVal index As Long) As String
    Status = mStatus(index)
End Property

Public Property Get CapacityStatus(ByVal index As Long) As String
    CapacityStatus = mCapacityStatus(index)
End Property

Public Property Get CapacityLimitState(ByVal index As Long) As String
    CapacityLimitState = mCapacityLimitState(index)
End Property

Public Property Get StrengthCheckStatus(ByVal index As Long) As String
    StrengthCheckStatus = mStrengthCheckStatus(index)
End Property

Public Property Get CrackStatus(ByVal index As Long) As String
    CrackStatus = mCrackStatus(index)
End Property

Public Property Get LambdaCapacity(ByVal index As Long) As Double
    LambdaCapacity = mLambdaCapacity(index)
End Property

Public Property Get UtilizationCapacity(ByVal index As Long) As Double
    UtilizationCapacity = mUtilCapacity(index)
End Property

Public Property Get CrackWidth(ByVal index As Long) As Double
    CrackWidth = mCrackWidth(index)
End Property

Public Property Get CrackAllowable(ByVal index As Long) As Double
    CrackAllowable = mCrackAllowableValue(index)
End Property

Public Property Get CrackUtilization(ByVal index As Long) As Double
    CrackUtilization = mCrackUtil(index)
End Property

Public Property Get CrackFormed(ByVal index As Long) As Boolean
    CrackFormed = mCrackFormed(index)
End Property

Public Property Get CrackPhi1(ByVal index As Long) As Double
    CrackPhi1 = mCrackPhi1(index)
End Property

Public Property Get CrackPhi2(ByVal index As Long) As Double
    CrackPhi2 = mCrackPhi2(index)
End Property

Public Property Get CrackPhi3(ByVal index As Long) As Double
    CrackPhi3 = mCrackPhi3(index)
End Property

Public Property Get CrackPsiS(ByVal index As Long) As Double
    CrackPsiS = mCrackPsiS(index)
End Property

Public Property Get CrackSigmaS(ByVal index As Long) As Double
    CrackSigmaS = mCrackSigmaS(index)
End Property

Public Property Get CrackSigmaSCrc(ByVal index As Long) As Double
    CrackSigmaSCrc = mCrackSigmaSCrc(index)
End Property

Public Property Get CrackAs(ByVal index As Long) As Double
    CrackAs = mCrackAs(index)
End Property

Public Property Get CrackAbt(ByVal index As Long) As Double
    CrackAbt = mCrackAbt(index)
End Property

Public Property Get CrackDsEquivalent(ByVal index As Long) As Double
    CrackDsEquivalent = mCrackDsEquivalent(index)
End Property

Public Property Get CrackSpacingRaw(ByVal index As Long) As Double
    CrackSpacingRaw = mCrackLsRaw(index)
End Property

Public Property Get CrackSpacing(ByVal index As Long) As Double
    CrackSpacing = mCrackLs(index)
End Property

Public Property Get CrackLambdaCrc(ByVal index As Long) As Double
    CrackLambdaCrc = mCrackLambdaCrc(index)
End Property

Public Property Get CrackNcrc(ByVal index As Long) As Double
    CrackNcrc = mCrackNcrc(index)
End Property

Public Property Get CrackTensionRebarCount(ByVal index As Long) As Long
    CrackTensionRebarCount = mCrackTensionRebarCount(index)
End Property

Public Property Get CrackEffectiveZoneDepth(ByVal index As Long) As Double
    CrackEffectiveZoneDepth = mCrackEffectiveZoneDepth(index)
End Property

Public Property Get CrackCentralBranch(ByVal index As Long) As Boolean
    CrackCentralBranch = mCrackCentralBranch(index)
End Property

Public Property Get StrainSafetyFactor(ByVal index As Long) As Double
    StrainSafetyFactor = mStrainSafetyFactor(index)
End Property

Public Property Get GoverningUtilization(ByVal index As Long) As Double
    GoverningUtilization = mGoverningUtil(index)
End Property

Public Property Get GoverningCombinationIndex() As Long
    GoverningCombinationIndex = mGoverningCombinationIndex
End Property

Public Property Get GoverningCombinationID() As String
    If mGoverningCombinationIndex > 0 Then GoverningCombinationID = mCombinationID(mGoverningCombinationIndex)
End Property

Public Property Get GoverningCombinationName() As String
    If mGoverningCombinationIndex > 0 Then GoverningCombinationName = mComment(mGoverningCombinationIndex)
End Property

Public Property Get ElapsedSeconds() As Double
    ElapsedSeconds = mElapsedSeconds
End Property

Public Property Get DiagnosticLog() As String
    DiagnosticLog = mDiagnosticLog
End Property

Public Property Get InvalidInputCount() As Long
    Dim i As Long
    For i = 1 To mCount
        If IsInvalidStatus(mStatus(i)) Then InvalidInputCount = InvalidInputCount + 1
    Next i
End Property

Public Property Get FirstInvalidInputMessage() As String
    Dim i As Long
    For i = 1 To mCount
        If IsInvalidStatus(mStatus(i)) Then
            FirstInvalidInputMessage = mStatus(i)
            Exit Property
        End If
    Next i
End Property

Public Property Get StateAvailable(ByVal index As Long) As Boolean
    StateAvailable = mStateAvailable(index)
End Property

Public Property Get StateAvailableCount() As Long
    Dim i As Long
    For i = 1 To mCount
        If mStateAvailable(i) Then StateAvailableCount = StateAvailableCount + 1
    Next i
End Property

Public Property Get StateConverged(ByVal index As Long) As Boolean
    StateConverged = mStateConverged(index)
End Property

Public Property Get Epsilon0(ByVal index As Long) As Double
    Epsilon0 = mEpsilon0(index)
End Property

Public Property Get KappaX(ByVal index As Long) As Double
    KappaX = mKappaX(index)
End Property

Public Property Get KappaY(ByVal index As Long) As Double
    KappaY = mKappaY(index)
End Property

Public Property Get Nint(ByVal index As Long) As Double
    Nint = mNint(index)
End Property

Public Property Get Mxint(ByVal index As Long) As Double
    Mxint = mMxint(index)
End Property

Public Property Get Myint(ByVal index As Long) As Double
    Myint = mMyint(index)
End Property

Public Property Get ResidualN(ByVal index As Long) As Double
    ResidualN = mResidualN(index)
End Property

Public Property Get ResidualMx(ByVal index As Long) As Double
    ResidualMx = mResidualMx(index)
End Property

Public Property Get ResidualMy(ByVal index As Long) As Double
    ResidualMy = mResidualMy(index)
End Property

Public Property Get StateIterations(ByVal index As Long) As Long
    StateIterations = mStateIterations(index)
End Property


