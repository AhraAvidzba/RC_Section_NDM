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
Private Const CALC_GROUP1 As String = "Group1" ' Первая группа: проверка несущей способности.
Private Const CALC_GROUP2 As String = "Group2" ' Вторая группа: эксплуатационная проверка трещин.
Private Const CAPACITY_PATH_NONE As String = "None" ' Capacity не запускался для LC.

' Данные одной строки таблицы сочетаний. Batch хранит их отдельно от результатов,
' чтобы было видно, где исходные нагрузки пользователя, а где то, что вернул расчет.
Private Type TCombinationInput
    CombinationID As String
    N As Double
    Mx As Double
    My As Double
    UserMx As Double
    UserMy As Double
    CalculationType As String
    CapacityLoadPath As String
    Comment As String
    InputValid As Boolean
    InputError As String
End Type

' Результаты одного LC после прямого НДС, capacity и трещин. Поля собраны в
' одну запись, чтобы исходные данные сочетаний и расчетные результаты не
' смешивались между разными строками batch.
Private Type TCombinationResult
    Status As String
    DirectStateStatus As String
    CapacityStatus As String
    CapacityLimitState As String
    CapacitySolutionMethod As String
    CrackStatus As String
    ExtensionUsed As Boolean
    DirectStateDiagramExtensionEnabled As Boolean
    LambdaCapacity As Double
    UtilCapacity As Double
    CapacityPathResolved As String
    NUltimate As Double
    MxUltimate As Double
    MyUltimate As Double
    MomentUltimate As Double
    CapacityStateAvailable As Boolean
    CapacityEpsilon0 As Double
    CapacityKappaX As Double
    CapacityKappaY As Double
    MinConcreteStrain As Double
    MaxConcreteStrain As Double
    MinSteelStrain As Double
    MaxSteelStrain As Double
    CrackWidth As Double
    CrackAllowableValue As Double
    CrackUtil As Double
    CrackFormed As Boolean
    CrackPhi1 As Double
    CrackPhi2 As Double
    CrackPhi3 As Double
    CrackPsiS As Double
    CrackSigmaS As Double
    CrackSigmaSCrc As Double
    CrackAs As Double
    CrackAbt As Double
    CrackDsEquivalent As Double
    CrackLsRaw As Double
    CrackLs As Double
    CrackLambdaCrc As Double
    CrackNcrc As Double
    CrackTensionRebarCount As Long
    CrackSectionDepthH As Double
    CrackCoverA As Double
    CrackTensionDepth As Double
    CrackEffectiveZoneDepth As Double
    CrackCentralBranch As Boolean
    GoverningUtil As Double
    StateAvailable As Boolean
    StateConverged As Boolean
    Epsilon0 As Double
    KappaX As Double
    KappaY As Double
    Nint As Double
    Mxint As Double
    Myint As Double
    ResidualN As Double
    ResidualMx As Double
    ResidualMy As Double
    StateIterations As Long
End Type

Private mSection As CSectionModel ' Расчетная модель сечения, общая для всех LC.
Private mMaterialProvider As CMaterialModelProvider ' Единая точка выдачи диаграмм для Strength/Mcrc/CrackedNDS.
Private mStatusPolicy As CBatchStatusPolicy ' Единый словарь пользовательских статусов и правил их агрегирования.
Private mCount As Long ' Сколько сочетаний реально добавлено в пакет.

Private mInputs(1 To MAX_COMBINATIONS) As TCombinationInput ' Исходные данные LC, прочитанные из таблицы сочетаний.
Private mResults(1 To MAX_COMBINATIONS) As TCombinationResult ' Расчетные результаты LC для Summary/Results.
Private mLoadReferenceX As Double ' X точки приложения нагрузки относительно расчетной модели, мм.
Private mLoadReferenceY As Double ' Y точки приложения нагрузки относительно расчетной модели, мм.
Private mLoadReferenceOffsetX As Double ' Смещение точки нагрузки от центра тяжести бетонного сечения по X, мм.
Private mLoadReferenceOffsetY As Double ' Смещение точки нагрузки от центра тяжести бетонного сечения по Y, мм.
Private mGoverningCombinationIndex As Long ' Индекс LC, который стал определяющим в batch summary.
Private mCrackGoverningCombinationIndex As Long ' Индекс LC с минимальным запасом по раскрытию трещин.
Private mElapsedSeconds As Double ' Полное время пакетного расчета, сек.
Private mDiagnosticLog As String ' Короткий текстовый журнал по LC для отладки batch-потока.

Private mCrackAllowable As Double ' Допустимая ширина раскрытия трещины a_crc,ult, мм.
Private mCrackPsiMode As String ' Fixed1 или Auto: как определять psi_s для нормальных трещин.
Private mCrackTensionZoneMode As String ' Effective или FullTension: какая бетонная зона входит в Abt.
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
Private mCapacitySolutionStrategy As String ' Auto, LoadMultiplier или UltimateStrain.
Private mCapacityCalculationScope As String ' Какие LC отправлять в поиск несущего момента: только Group1 или Group1+Group2.
Private mCapacitySearchMethod As String ' Bisection, Brent или Secant для LoadMultiplier.
Private mCapacityMaxLambda As Double ' Верхняя защитная граница lambda.
Private mCapacityToleranceLambda As Double ' Точность lambda в одномерном поиске LoadMultiplier.
Private mCapacityToleranceStrain As Double ' Точность попадания в предельную деформацию UltimateStrain.
Private mCapacityMaxRetries As Long ' Повторные попытки LoadMultiplier при несходимости пробной точки.
Private mCapacityBaseLoadSteps As Long ' Базовое число load steps для пробной точки capacity.
Private mCapacitySolverMaxIterations As Long ' Лимит итераций внутреннего solver-а при capacity-поиске.
Private mReport As CExecutionReport ' Необязательный человекочитаемый отчет выполнения; расчетная логика от него не зависит.

' Инициализирует внутреннее состояние объекта значениями по умолчанию, безопасными для последующей настройки.
Private Sub Class_Initialize()
    Set mStatusPolicy = New CBatchStatusPolicy
    mCrackAllowable = 0.3
    mCrackPsiMode = "Fixed1"
    mCrackTensionZoneMode = "Effective"
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
    mCapacitySolutionStrategy = "Auto"
    mCapacityCalculationScope = "Group1+2"
    mCapacitySearchMethod = "Bisection"
    mCapacityMaxLambda = 64#
    mCapacityToleranceLambda = 0.01
    mCapacityToleranceStrain = 0.00001
    mCapacityMaxRetries = 0
    mCapacityBaseLoadSteps = 1
    mCapacitySolverMaxIterations = 60
End Sub

' Принимает готовую модель и материалы перед добавлением/расчетом LC.
' После этого класс не обращается к геометрическим генераторам: все сочетания
' считаются на одной и той же CSectionModel.
Public Sub Initialize(ByVal section As CSectionModel, ByVal materialProvider As CMaterialModelProvider)
    If section Is Nothing Then Err.Raise vbObjectError + 3900, "CBatchSectionCalculator", "Модель сечения не передана."
    If section.ConcreteCount <= 0 Then Err.Raise vbObjectError + 3901, "CBatchSectionCalculator", "В модели сечения нет бетонных элементов."
    If materialProvider Is Nothing Then Err.Raise vbObjectError + 3902, "CBatchSectionCalculator", "Material provider не передан."
    Set mSection = section
    Set mMaterialProvider = materialProvider
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
    mCapacitySolutionStrategy = settings.GetString("Capacity.SolutionStrategy", mCapacitySolutionStrategy)
    mCapacityCalculationScope = settings.GetString("Capacity.CalculationScope", mCapacityCalculationScope)
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
    If Not units Is Nothing Then
        mCrackAllowable = units.InputLengthToInternal(mCrackAllowable)
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
    mCrackGoverningCombinationIndex = 0
    mElapsedSeconds = 0#
    mDiagnosticLog = vbNullString
    mLoadReferenceX = 0#
    mLoadReferenceY = 0#
    mLoadReferenceOffsetX = 0#
    mLoadReferenceOffsetY = 0#
End Sub

' Добавляет корректно прочитанное сочетание нагрузок.
' Усилия уже должны быть во внутренних единицах и в INTERNAL-системе знаков.
Public Sub AddCombination(ByVal combinationID As String, ByVal nValue As Double, ByVal mxValue As Double, _
        ByVal myValue As Double, ByVal calculationType As String, ByVal comment As String, _
        Optional ByVal capacityLoadPath As String = vbNullString)
    AddCombinationInternal combinationID, nValue, mxValue, myValue, calculationType, comment, True, vbNullString, capacityLoadPath
End Sub

' Сохраняет строку LC с ошибкой ввода, чтобы пользователь увидел ее в отчете.
' Такие сочетания не передаются в solver, но отображаются как InvalidInput.
Public Sub AddInvalidCombination(ByVal combinationID As String, ByVal calculationType As String, _
        ByVal comment As String, ByVal reason As String, Optional ByVal capacityLoadPath As String = vbNullString)
    AddCombinationInternal combinationID, 0#, 0#, 0#, calculationType, comment, False, reason, capacityLoadPath
End Sub

Private Sub AddCombinationInternal(ByVal combinationID As String, ByVal nValue As Double, ByVal mxValue As Double, _
        ByVal myValue As Double, ByVal calculationType As String, ByVal comment As String, _
        ByVal inputValid As Boolean, ByVal inputError As String, ByVal capacityLoadPath As String)
    If mCount >= MAX_COMBINATIONS Then Err.Raise vbObjectError + 3911, "CBatchSectionCalculator", "Поддерживается не более 20 сочетаний нагрузок."
    mCount = mCount + 1
    If Len(Trim$(combinationID)) = 0 Then combinationID = "ROW" & CStr(mCount)
    mInputs(mCount).CombinationID = combinationID
    mInputs(mCount).N = nValue
    mInputs(mCount).Mx = mxValue
    mInputs(mCount).My = myValue
    mInputs(mCount).CalculationType = calculationType
    mInputs(mCount).CapacityLoadPath = Trim$(capacityLoadPath)
    mInputs(mCount).Comment = comment
    mInputs(mCount).UserMx = mxValue
    mInputs(mCount).UserMy = myValue
    mInputs(mCount).InputValid = inputValid
    mInputs(mCount).InputError = inputError
End Sub

' Переносит пользовательские моменты из точки приложения нагрузки к координатам элементов.
' Пользователь задает Mx/My относительно выбранной точки нагрузки, а solver
' должен получить моменты относительно внутренней системы координат CSectionModel.
Public Sub ApplyLoadReference(ByVal referenceX As Double, ByVal referenceY As Double, _
        Optional ByVal baseX As Double = 0#, Optional ByVal baseY As Double = 0#)
    mLoadReferenceX = referenceX
    mLoadReferenceY = referenceY
    mLoadReferenceOffsetX = referenceX - baseX
    mLoadReferenceOffsetY = referenceY - baseY

    Dim i As Long
    ' Эквивалентный перенос момента от продольной силы:
    ' Mx_model = Mx_user + N * y_ref, My_model = My_user + N * x_ref.
    For i = 1 To mCount
        If mInputs(i).InputValid Then
            mInputs(i).Mx = mInputs(i).UserMx + mInputs(i).N * referenceY
            mInputs(i).My = mInputs(i).UserMy + mInputs(i).N * referenceX
        End If
    Next i
End Sub

' Последовательно считает все добавленные LC и выбирает определяющее сочетание.
' Порядок один для всех геометрий: capacity при необходимости, прямое НДС,
' затем трещины для сочетаний второй группы.
Public Sub Execute()
    On Error GoTo Failed

    ValidateReady
    Dim t0 As Double
    Dim stage As String
    Dim currentIndex As Long

    stage = "Initialize"
    t0 = Timer
    mGoverningCombinationIndex = 0
    mCrackGoverningCombinationIndex = 0
    mDiagnosticLog = vbNullString
    ReportStep "Начат пакетный расчет. Количество сочетаний: " & CStr(mCount) & "."

    Dim i As Long
    ' Каждый LC изолирован: ошибка одного сочетания не должна стирать результаты
    ' соседних строк, но попадает в итоговый статус и batch summary.
    For i = 1 To mCount
        currentIndex = i
        stage = "ClearResultRow"
        ClearResultRow i
        stage = "RunCombination"
        RunCombination i
        If mInputs(i).InputValid Then
            stage = "SelectStrengthGoverning"
            If StrengthSafetyFactorForGoverning(i) > 0# Then
                If mGoverningCombinationIndex = 0 Then
                    mGoverningCombinationIndex = i
                ElseIf IsBetterGoverningCandidate(i, mGoverningCombinationIndex) Then
                    mGoverningCombinationIndex = i
                End If
            End If
            stage = "SelectCrackGoverning"
            If IsBetterCrackGoverningCandidate(i, mCrackGoverningCombinationIndex) Then
                mCrackGoverningCombinationIndex = i
            End If
        End If
    Next i
    stage = "Finalize"
    mElapsedSeconds = Timer - t0
    ReportStep "Пакетный расчет завершен. Время: " & FormatNumberInvariant(mElapsedSeconds) & " с."
    Exit Sub

Failed:
    Err.Raise Err.Number, "CBatchSectionCalculator.Execute", _
        stage & " LC=" & CStr(currentIndex) & ": " & Err.Description
End Sub

' Считает одно сочетание нагрузок и собирает итоговый статус.
' Внутри сохраняется разделение: capacity отвечает за запас по несущей,
' прямой solve дает фактическое НДС, crack запускается только для 2 группы.
Private Sub RunCombination(ByVal index As Long)
    On Error GoTo Failed
    ReportSection "Сочетание " & CStr(index) & ": " & mInputs(index).CombinationID
    ReportValue "Комментарий", mInputs(index).Comment
    If Not mInputs(index).InputValid Then
        ReportStep "Сочетание не передано в solver из-за ошибки входных данных: " & mInputs(index).InputError
        MarkInvalid index, mInputs(index).InputError
        Exit Sub
    End If
    If Not IsSupportedCalculationMode() Then
        ReportStep "Сочетание не рассчитано: неизвестный Calculation.Mode."
        MarkInvalid index, "Calculation.Mode должен быть DirectState или FullCapacity."
        Exit Sub
    End If

    Dim calculationType As String
    calculationType = NormalizedCombinationCalculationType(mInputs(index).CalculationType)
    ReportValue "Тип проверки", IIf(Len(calculationType) > 0, calculationType, mInputs(index).CalculationType)
    ReportValue "N", FormatNumberInvariant(mInputs(index).N) & " Н"
    ReportValue "Mx расчетный", FormatNumberInvariant(mInputs(index).Mx) & " Н*мм"
    ReportValue "My расчетный", FormatNumberInvariant(mInputs(index).My) & " Н*мм"
    If Len(calculationType) = 0 Then
        ReportStep "Сочетание не рассчитано: CalculationType должен быть Group1 или Group2."
        MarkInvalid index, "CalculationType должен быть Group1 или Group2."
        Exit Sub
    End If
    If IsFullCapacityMode And Len(NormalizedCapacityCalculationScope()) = 0 Then
        ReportStep "Сочетание не рассчитано: неизвестная область расчета Capacity."
        MarkInvalid index, "Capacity.CalculationScope должен быть Group1Only или Group1+2."
        Exit Sub
    End If

    mResults(index).Status = mStatusPolicy.NotApplicable
    If ShouldRunCapacity(calculationType) And IsFullCapacityMode Then
        ReportStep "Запускается проверка несущей способности для сочетания."
        RunCapacity index
    Else
        ReportStep "Проверка несущей способности не требуется для этого сочетания или режима Calculation.Mode."
        StoreNotCalculatedCapacity index
    End If
    RunSectionStateAndCrack index, ShouldRunCrack(calculationType)

    mResults(index).GoverningUtil = MaxDouble(mResults(index).UtilCapacity, mResults(index).CrackUtil)
    mResults(index).Status = mStatusPolicy.Aggregate(mResults(index).DirectStateStatus, _
        mResults(index).CapacityStatus, mResults(index).CrackStatus)
    AppendDiagnostic index
    ReportCombinationSummary index
    Exit Sub

Failed:
    ReportStep "Ошибка при расчете сочетания: " & Err.Description
    MarkInvalid index, Err.Description
End Sub

' Запускает поиск несущей способности для LC.
' Пользовательский столбец CapacityLoadPath задает, какие компоненты нагрузки
' умножаются на lambda. Все варианты дальше сводятся к одной математике:
' Target = Offset + lambda*Base для N, Mx и My.
Private Sub RunCapacity(ByVal index As Long)
    Dim loadPath As CCapacityLoadPath
    Dim cap As CCapacitySolver

    Set loadPath = BuildLoadPathForCombination(index)
    mResults(index).CapacityPathResolved = loadPath.Key
    ReportValue "CapacityLoadPath", IIf(loadPath.HasPath, loadPath.DisplayName, mInputs(index).CapacityLoadPath)

    If Not loadPath.Valid Then
        StoreInvalidCapacity index, loadPath.ErrorMessage
        ReportStep "Capacity остановлен: неизвестное значение CapacityLoadPath."
        Exit Sub
    End If

    If Not loadPath.HasPath Then
        ReportStep "Capacity не запускается: в сочетании нет ни момента, ни продольной силы."
        StoreSkippedCapacity index
        Exit Sub
    End If

    If Not loadPath.HasScaledLoad Then
        StoreInvalidCapacity index, "Для выбранного CapacityLoadPath соответствующая нагрузка должна быть ненулевой."
        ReportStep "Capacity остановлен: выбранная lambda-траектория не содержит ненулевой масштабируемой нагрузки."
        Exit Sub
    End If

    Set cap = New CCapacitySolver
    ConfigureCapacity cap
    Select Case NormalizedCapacitySolutionStrategy()
        Case "Auto"
            If loadPath.ForceOnly Then
                cap.SolveByLoadPathMultiplier mSection, mMaterialProvider.ConcreteMaterial(cpStrength), _
                    mMaterialProvider.SteelMaterial(cpStrength), loadPath.NOffset, loadPath.NBase, _
                    loadPath.MxOffset, loadPath.MxBase, loadPath.MyOffset, loadPath.MyBase, True
            Else
                cap.SolveByAutoLoadPath mSection, mMaterialProvider.ConcreteMaterial(cpStrength), _
                    mMaterialProvider.SteelMaterial(cpStrength), loadPath.NOffset, loadPath.NBase, _
                    loadPath.MxOffset, loadPath.MxBase, loadPath.MyOffset, loadPath.MyBase
            End If
        Case "UltimateStrain"
            If loadPath.ForceOnly Then
                cap.SolveByLoadPathMultiplier mSection, mMaterialProvider.ConcreteMaterial(cpStrength), _
                    mMaterialProvider.SteelMaterial(cpStrength), loadPath.NOffset, loadPath.NBase, _
                    loadPath.MxOffset, loadPath.MxBase, loadPath.MyOffset, loadPath.MyBase, True
            Else
                cap.SolveByUltimateLoadPath mSection, mMaterialProvider.ConcreteMaterial(cpStrength), _
                    mMaterialProvider.SteelMaterial(cpStrength), loadPath.NOffset, loadPath.NBase, _
                    loadPath.MxOffset, loadPath.MxBase, loadPath.MyOffset, loadPath.MyBase
            End If
        Case "LoadMultiplier"
            cap.SolveByLoadPathMultiplier mSection, mMaterialProvider.ConcreteMaterial(cpStrength), _
                mMaterialProvider.SteelMaterial(cpStrength), loadPath.NOffset, loadPath.NBase, _
                loadPath.MxOffset, loadPath.MxBase, loadPath.MyOffset, loadPath.MyBase, loadPath.ForceOnly
        Case Else
            mResults(index).CapacityStatus = mStatusPolicy.InputErr
            mResults(index).CapacityLimitState = "Capacity.SolutionStrategy должен быть Auto, LoadMultiplier или UltimateStrain."
            mResults(index).CapacitySolutionMethod = vbNullString
            ReportStep "Capacity остановлен: неизвестный Capacity.SolutionStrategy."
            Exit Sub
    End Select

    StoreCapacity index, cap, loadPath

    ReportValue "CapacityStatus", mResults(index).CapacityStatus
    ReportValue "LambdaCapacity", FormatNumberInvariant(mResults(index).LambdaCapacity)
    ReportValue "CapacityLoadPath", CapacityLoadPathDisplay(mResults(index).CapacityPathResolved)
    ReportBlock "Диагностика поиска несущей способности", cap.DiagnosticLog
    If Not cap.LastSolver Is Nothing Then
        ReportBlock "Итерации внутреннего CSectionSolver для последней capacity-точки", cap.LastSolver.DiagnosticLog
    End If
End Sub

' Выполняет прямой расчет НДС и, если нужно, расчет ширины трещин.
' Даже в FullCapacity этот solve нужен отдельно: он показывает состояние
' именно от заданного LC, а не только предельное состояние capacity solver-а.
Private Sub RunSectionStateAndCrack(ByVal index As Long, ByVal crackRequested As Boolean)
    Dim basePurpose As ECalculationPurpose
    basePurpose = MaterialPurposeForCombination(index)

    Dim concreteState As CMaterialDiagram
    Dim steelState As CMaterialDiagram
    Set concreteState = mMaterialProvider.ConcreteStateMaterial(basePurpose)
    Set steelState = mMaterialProvider.SteelStateMaterial(basePurpose)
    mResults(index).DirectStateDiagramExtensionEnabled = mMaterialProvider.DirectStateDiagramExtensionEnabled

    Dim runner As CStateSolutionRunner
    Set runner = New CStateSolutionRunner
    ConfigureStateRunner runner
    ReportStep "Запускается прямой расчет НДС CSectionSolver."
    runner.Solve mSection, concreteState, steelState, mInputs(index).N, mInputs(index).Mx, mInputs(index).My, _
        mResults(index).DirectStateDiagramExtensionEnabled

    Dim service As CSectionSolver
    Set service = runner.ResultSolver
    If service Is Nothing Then
        mResults(index).DirectStateStatus = mStatusPolicy.NumFail
        mResults(index).CrackStatus = mStatusPolicy.NotApplicable
        ReportStep "Расчет трещин не запускается: прямой НДС не вернул состояние solver-а."
        Exit Sub
    End If
    StoreSectionState index, service
    ReportValue "НДС сошелся", IIf(service.Converged, "да", "нет")
    ReportValue "Итераций НДС", CStr(runner.Iterations)
    ReportBlock "Сценарий прямого StateSolution", runner.DiagnosticLog
    ReportBlock "Итерации прямого CSectionSolver", service.DiagnosticLog
    If Not service.Converged Then
        mResults(index).DirectStateStatus = mStatusPolicy.NumFail
        mResults(index).CrackStatus = mStatusPolicy.NotApplicable
        ReportStep "Расчет трещин не запускается, потому что прямое НДС не сошлось: " & service.StopReason
        Exit Sub
    End If

    Dim stateWithinPhysical As Boolean
    stateWithinPhysical = runner.WithinPhysicalRange
    mResults(index).ExtensionUsed = runner.ExtensionUsed

    If mResults(index).ExtensionUsed Or Not stateWithinPhysical Then
        mResults(index).DirectStateStatus = mStatusPolicy.Fail
        ReportStep "Прямое НДС найдено, но вышло за физический диапазон диаграммы материала."
    Else
        mResults(index).DirectStateStatus = mStatusPolicy.OK
    End If
    ReportValue "DirectStateStatus", mResults(index).DirectStateStatus
    ReportValue "ExtensionUsed", IIf(mResults(index).ExtensionUsed, "да", "нет")

    If Not crackRequested Or Not mCrackEnabled Then
        mResults(index).CrackStatus = mStatusPolicy.NotApplicable
        If crackRequested Then
            ReportStep "Расчет трещин отключен настройкой SLS.Crack.Enabled."
        Else
            ReportStep "Расчет трещин не требуется для этого типа сочетания."
        End If
        Exit Sub
    End If

    If mResults(index).DirectStateStatus <> mStatusPolicy.OK Or mResults(index).ExtensionUsed Then
        mResults(index).CrackStatus = mStatusPolicy.NotApplicable
        ReportStep "Расчет трещин не запускается: исходное НДС Group2 не находится полностью в физическом диапазоне диаграммы."
        Exit Sub
    End If

    Dim crack As CCrackWidthCalculator
    Set crack = New CCrackWidthCalculator
    ConfigureCrackCalculator crack
    ReportStep "Запускается расчет ширины раскрытия нормальных трещин."

    Dim centroidMxForCrack As Double
    Dim centroidMyForCrack As Double
    ' Для отдельной ветки центрального растяжения важны моменты относительно
    ' центра тяжести бетонного сечения. Если точка приложения N смещена от
    ' бетонного центра, то даже при Mx_user = My_user = 0 появляется
    ' внецентренное растяжение.
    centroidMxForCrack = mInputs(index).UserMx + mInputs(index).N * mLoadReferenceOffsetY
    centroidMyForCrack = mInputs(index).UserMy + mInputs(index).N * mLoadReferenceOffsetX

    crack.Calculate service, mSection, mMaterialProvider, basePurpose, _
        mInputs(index).N, mInputs(index).Mx, mInputs(index).My, _
        centroidMxForCrack, centroidMyForCrack
    mResults(index).CrackStatus = mStatusPolicy.CrackStatusFromCalculator(crack)
    mResults(index).CrackWidth = crack.CrackWidth
    mResults(index).CrackAllowableValue = crack.AllowableCrackWidth
    mResults(index).CrackUtil = crack.Utilization
    mResults(index).CrackFormed = crack.CrackFormed
    mResults(index).CrackPhi1 = crack.Phi1
    mResults(index).CrackPhi2 = crack.Phi2
    mResults(index).CrackPhi3 = crack.Phi3
    mResults(index).CrackPsiS = crack.PsiS
    mResults(index).CrackSigmaS = crack.SigmaS
    mResults(index).CrackSigmaSCrc = crack.SigmaSCrc
    mResults(index).CrackAs = crack.AsTension
    mResults(index).CrackAbt = crack.Abt
    mResults(index).CrackDsEquivalent = crack.DsEquivalent
    mResults(index).CrackLsRaw = crack.CrackSpacingRaw
    mResults(index).CrackLs = crack.CrackSpacing
    mResults(index).CrackLambdaCrc = crack.LambdaCrc
    mResults(index).CrackNcrc = crack.Ncrc
    mResults(index).CrackTensionRebarCount = crack.TensionRebarCount
    mResults(index).CrackSectionDepthH = crack.SectionDepthH
    mResults(index).CrackCoverA = crack.CoverA
    mResults(index).CrackTensionDepth = crack.TensionDepth
    mResults(index).CrackEffectiveZoneDepth = crack.EffectiveZoneDepth
    mResults(index).CrackCentralBranch = crack.CentralTensionBranch
    ReportValue "CrackStatus", mResults(index).CrackStatus
    ReportValue "Трещина образована", IIf(mResults(index).CrackFormed, "да", "нет")
    ReportValue "a_crc", FormatNumberInvariant(mResults(index).CrackWidth) & " мм"
    ReportValue "psi_s", FormatNumberInvariant(mResults(index).CrackPsiS)
    ReportValue "As", FormatNumberInvariant(mResults(index).CrackAs) & " мм2"
    ReportValue "Abt", FormatNumberInvariant(mResults(index).CrackAbt) & " мм2"
    ReportValue "ds_eq", FormatNumberInvariant(mResults(index).CrackDsEquivalent) & " мм"
    ReportValue "ls", FormatNumberInvariant(mResults(index).CrackLs) & " мм"
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
' Пользовательская настройка Capacity.CalculationScope позволяет оставить
' поиск предельной lambda-траектории только для первой группы или временно считать
' его также для Group2 как диагностический прочностной запас.
Private Function ShouldRunCapacity(ByVal calculationType As String) As Boolean
    Select Case NormalizedCapacityCalculationScope()
        Case "Group1Only"
            ShouldRunCapacity = (calculationType = CALC_GROUP1)
        Case "Group1AndGroup2"
            ShouldRunCapacity = (calculationType = CALC_GROUP1 Or calculationType = CALC_GROUP2)
    End Select
End Function

' Показывает, должна ли для LC запускаться проверка трещин.
' Текущая реализация считает только продолжительное раскрытие нормальных
' трещин для второй группы; сочетания первой группы сюда не попадают.
Private Function ShouldRunCrack(ByVal calculationType As String) As Boolean
    ShouldRunCrack = (calculationType = CALC_GROUP2)
End Function

' Создает объект lambda-траектории для одной строки LC.
' Batch передает сюда только уже нормализованные усилия и точку приложения
' нагрузки, а все подробности интерпретации lambda*Mx/lambda*N остаются в CCapacityLoadPath.
Private Function BuildLoadPathForCombination(ByVal index As Long) As CCapacityLoadPath
    Dim loadPath As CCapacityLoadPath
    Set loadPath = New CCapacityLoadPath
    loadPath.Initialize mInputs(index).CapacityLoadPath, mInputs(index).N, mInputs(index).UserMx, mInputs(index).UserMy, _
        mLoadReferenceX, mLoadReferenceY
    Set BuildLoadPathForCombination = loadPath
End Function

' Формирует человекочитаемую подпись load-path для Results и отчета.
' Само соответствие ключа и подписи хранится в CCapacityLoadPath, чтобы batch
' не становился вторым местом интерпретации lambda-траекторий.
Private Function CapacityLoadPathDisplay(ByVal loadPathKey As String) As String
    Dim loadPath As CCapacityLoadPath
    Set loadPath = New CCapacityLoadPath
    CapacityLoadPathDisplay = loadPath.DisplayNameForKey(loadPathKey)
End Function

' Выбирает material purpose для фактического НДС конкретного LC.
' Group1 решается по Strength-модели I ГПС, Group2 - по CrackedNDS-модели
' II ГПС с неработающим растянутым бетоном, как принято для a_crc.
Private Function MaterialPurposeForCombination(ByVal index As Long) As ECalculationPurpose
    If NormalizedCombinationCalculationType(mInputs(index).CalculationType) = CALC_GROUP2 Then
        MaterialPurposeForCombination = cpCrackedNDS
    Else
        MaterialPurposeForCombination = cpStrength
    End If
End Function

' Возвращает бетонную StateSolution-диаграмму для writer-а Results.
' Если прямой НДС решался с extension, именно эта диаграмма воспроизводит
' сохраненные Stress, которые пользователь увидит на Results/схеме/AutoCAD.
Public Function ConcreteMaterialForCombination(ByVal index As Long) As Object
    Set ConcreteMaterialForCombination = mMaterialProvider.ConcreteStateMaterial(MaterialPurposeForCombination(index))
End Function

' Возвращает StateSolution-диаграмму арматуры для writer-а Results.
Public Function SteelMaterialForCombination(ByVal index As Long) As Object
    Set SteelMaterialForCombination = mMaterialProvider.SteelStateMaterial(MaterialPurposeForCombination(index))
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

' Приводит пользовательскую строку Capacity.SolutionStrategy к одному из трех
' внутренних имен. Пустой результат дальше превращается в InvalidInput.
Private Function NormalizedCapacitySolutionStrategy() As String
    Select Case LCase$(Trim$(mCapacitySolutionStrategy))
        Case "auto"
            NormalizedCapacitySolutionStrategy = "Auto"
        Case "loadmultiplier"
            NormalizedCapacitySolutionStrategy = "LoadMultiplier"
        Case "ultimatestrain"
            NormalizedCapacitySolutionStrategy = "UltimateStrain"
    End Select
End Function

' Заполняет capacity-поля нейтральным статусом, когда эта проверка не должна
' выполняться для выбранной группы/режима.
Private Sub StoreNotCalculatedCapacity(ByVal index As Long)
    mResults(index).CapacityStatus = mStatusPolicy.NotApplicable
    mResults(index).CapacityLimitState = mStatusPolicy.NotApplicable
    mResults(index).CapacitySolutionMethod = vbNullString
    mResults(index).CapacityPathResolved = CAPACITY_PATH_NONE
End Sub

' Помечает capacity как пропущенную из-за нулевой нагрузки.
' Это не ошибка исходных данных: для нулевых N, Mx и My нельзя построить
' направление внешнего поиска несущей способности.
Private Sub StoreSkippedCapacity(ByVal index As Long)
    mResults(index).CapacityStatus = mStatusPolicy.NotApplicable
    mResults(index).CapacityLimitState = "ZeroBaseLoad"
    mResults(index).CapacitySolutionMethod = vbNullString
    mResults(index).CapacityPathResolved = CAPACITY_PATH_NONE
End Sub

' Помечает только ветку capacity как ошибку исходных данных.
' Прямое НДС при этом все равно может быть посчитано: ошибка относится к
' выбору траектории несущей способности, а не к самим нагрузкам LC.
Private Sub StoreInvalidCapacity(ByVal index As Long, ByVal reason As String)
    mResults(index).CapacityStatus = mStatusPolicy.InputErr
    mResults(index).CapacityLimitState = reason
    mResults(index).CapacitySolutionMethod = vbNullString
    mResults(index).CapacityPathResolved = CAPACITY_PATH_NONE
    mResults(index).LambdaCapacity = 0#
    mResults(index).UtilCapacity = 0#
    mResults(index).NUltimate = 0#
    mResults(index).MxUltimate = 0#
    mResults(index).MyUltimate = 0#
    mResults(index).MomentUltimate = 0#
End Sub

' Переносит результат CCapacitySolver в массивы batch-отчета.
' Здесь разделяется математическая сходимость capacity solver-а и простая
' инженерная проверка: lambda меньше или больше единицы.
Private Sub StoreCapacity(ByVal index As Long, ByVal cap As CCapacitySolver, ByVal loadPath As CCapacityLoadPath)
    mResults(index).CapacityStatus = mStatusPolicy.CapacityStatusFromSolver(cap)
    mResults(index).CapacityLimitState = cap.LimitState
    mResults(index).CapacitySolutionMethod = cap.SolutionMethod
    mResults(index).CapacityPathResolved = loadPath.Key
    mResults(index).LambdaCapacity = cap.LambdaUltimate
    If cap.LambdaUltimate > 0# Then
        If loadPath.ScalesN Then mResults(index).NUltimate = cap.LambdaUltimate * mInputs(index).N
        If loadPath.ScalesMx Then mResults(index).MxUltimate = cap.LambdaUltimate * mInputs(index).UserMx
        If loadPath.ScalesMy Then mResults(index).MyUltimate = cap.LambdaUltimate * mInputs(index).UserMy
    End If
    mResults(index).MomentUltimate = Sqr(mResults(index).MxUltimate * mResults(index).MxUltimate + mResults(index).MyUltimate * mResults(index).MyUltimate)
    If cap.LambdaUltimate > 0# Then mResults(index).UtilCapacity = 1# / cap.LambdaUltimate
    If cap.Converged Then mResults(index).UtilCapacity = cap.Utilization

    ' h/x в блоке предельного момента должны относиться к найденному
    ' предельному состоянию, а не к фактическому НДС LC. Поэтому сохраняем
    ' отдельную плоскость деформаций только когда capacity-поиск действительно
    ' дал физически принятое предельное состояние.
    If cap.Converged And cap.LambdaUltimate > 0# Then
        If Not cap.LastSolver Is Nothing Then
            mResults(index).CapacityStateAvailable = True
            mResults(index).CapacityEpsilon0 = cap.LastSolver.Epsilon0
            mResults(index).CapacityKappaX = cap.LastSolver.KappaX
            mResults(index).CapacityKappaY = cap.LastSolver.KappaY
        End If
    End If
End Sub

' Фиксирует InvalidInput по конкретному LC и не дает этому сочетанию пройти
' в solver. Так пользователь видит причину в таблице, а остальные LC считаются.
Private Sub MarkInvalid(ByVal index As Long, ByVal reason As String)
    If Len(Trim$(reason)) = 0 Then reason = "Некорректное сочетание нагрузок."
    mInputs(index).InputError = reason
    mResults(index).Status = mStatusPolicy.InputErr
    mResults(index).DirectStateStatus = mStatusPolicy.InputErr
    mResults(index).CapacityStatus = mStatusPolicy.InputErr
    mResults(index).CapacityLimitState = reason
    mResults(index).CapacitySolutionMethod = vbNullString
    mResults(index).CapacityPathResolved = CAPACITY_PATH_NONE
    mResults(index).CrackStatus = mStatusPolicy.NotApplicable
    If Not mMaterialProvider Is Nothing Then
        mResults(index).DirectStateDiagramExtensionEnabled = mMaterialProvider.DirectStateDiagramExtensionEnabled
    End If
    mResults(index).GoverningUtil = 0#
    AppendDiagnostic index
End Sub

' Передает в CCapacitySolver ровно те настройки, которые относятся к поиску
' несущей способности и внутренним пробным расчетам равновесия.
Private Sub ConfigureCapacity(ByVal cap As CCapacitySolver)
    cap.ConcreteCompressionLimit = mMaterialProvider.ConcreteCompressionLimit(cpStrength)
    cap.ConcreteTensionLimit = mMaterialProvider.ConcreteTensionLimit(cpStrength)
    cap.ConcreteTensionLimitEnabled = mMaterialProvider.ConcreteTensionLimitEnabled(cpStrength)
    cap.SteelStrainLimit = MaxDouble(Abs(mMaterialProvider.SteelCompressionLimit(cpStrength)), _
        Abs(mMaterialProvider.SteelTensionLimit(cpStrength)))
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
    cap.DiagnosticsEnabled = DetailedDiagnosticsEnabled()
End Sub

' Настраивает runner прямого НДС общими solver-параметрами batch-расчета.
' Здесь остается только передача настроек; сама логика initial guess, retry
' и extension живет в CStateSolutionRunner.
Private Sub ConfigureStateRunner(ByVal runner As CStateSolutionRunner)
    runner.LoadSteps = mSolverLoadSteps
    runner.SolverMethod = mSolverMethod
    runner.MaxIterations = mSolverMaxIterations
    runner.ToleranceN = mSolverToleranceN
    runner.ToleranceMx = mSolverToleranceMx
    runner.ToleranceMy = mSolverToleranceMy
    runner.LineSearchEnabled = mSolverLineSearchEnabled
    runner.DampingInitial = mSolverDampingInitial
    runner.MinLineSearchAlpha = mSolverMinLineSearchAlpha
    runner.MaxDeltaEpsilon0 = mSolverMaxDeltaEpsilon0
    runner.MaxDeltaKappa = mSolverMaxDeltaKappa
    runner.SecantMaxRestarts = mSolverSecantMaxRestarts
    runner.SecantMinStepNorm = mSolverSecantMinStepNorm
    runner.DiagnosticsEnabled = DetailedDiagnosticsEnabled()
End Sub

' Передает в расчет трещин те же настройки прямого равновесия, что и в
' сервисный solve LC. Это важно для сопоставимости: состояние с раскрытой
' трещиной ищется тем же CSectionSolver, только с принудительно выключенным
' растянутым бетоном.
Private Sub ConfigureCrackCalculator(ByVal crack As CCrackWidthCalculator)
    crack.AllowableCrackWidth = mCrackAllowable
    crack.PsiMode = mCrackPsiMode
    crack.TensionZoneMode = mCrackTensionZoneMode
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
    crack.DiagnosticsEnabled = DetailedDiagnosticsEnabled()
End Sub

' Проверяет, нужен ли подробный журнал внутренних итераций.
' Короткая batch-сводка всегда остается в памяти для тестов и диагностики, но
' тяжелые строки CSectionSolver/CCapacitySolver/CCrackWidthCalculator имеют
' смысл только тогда, когда пользователь включил txt-отчет выполнения.
Private Function DetailedDiagnosticsEnabled() As Boolean
    If mReport Is Nothing Then Exit Function
    DetailedDiagnosticsEnabled = mReport.Enabled
End Function

' Проверяет, что batch готов к расчету.
' Это последний защитный барьер перед запуском solver-ов: модель и хотя бы
' одно сочетание должны быть уже загружены.
Private Sub ValidateReady()
    If mSection Is Nothing Then Err.Raise vbObjectError + 3920, "CBatchSectionCalculator", "Пакетный расчетчик не инициализирован моделью сечения."
    If mMaterialProvider Is Nothing Then Err.Raise vbObjectError + 3922, "CBatchSectionCalculator", "Пакетный расчетчик не инициализирован material provider-ом."
    If mCount <= 0 Then Err.Raise vbObjectError + 3921, "CBatchSectionCalculator", "Не задано ни одного сочетания нагрузок."
End Sub

' Сбрасывает все результаты одной строки перед повторным расчетом LC.
' Входные усилия и комментарий не трогаются.
Private Sub ClearResultRow(ByVal index As Long)
    mResults(index).Status = vbNullString
    mResults(index).DirectStateStatus = vbNullString
    mResults(index).CapacityStatus = vbNullString
    mResults(index).CapacityLimitState = vbNullString
    mResults(index).CapacitySolutionMethod = vbNullString
    mResults(index).CrackStatus = vbNullString
    mResults(index).LambdaCapacity = 0#
    mResults(index).UtilCapacity = 0#
    mResults(index).CapacityPathResolved = CAPACITY_PATH_NONE
    mResults(index).NUltimate = 0#
    mResults(index).MxUltimate = 0#
    mResults(index).MyUltimate = 0#
    mResults(index).CrackWidth = 0#: mResults(index).CrackAllowableValue = 0#: mResults(index).CrackUtil = 0#
    mResults(index).CrackFormed = False
    mResults(index).CrackPhi1 = 0#: mResults(index).CrackPhi2 = 0#: mResults(index).CrackPhi3 = 0#: mResults(index).CrackPsiS = 0#
    mResults(index).CrackSigmaS = 0#: mResults(index).CrackSigmaSCrc = 0#
    mResults(index).CrackAs = 0#: mResults(index).CrackAbt = 0#: mResults(index).CrackDsEquivalent = 0#
    mResults(index).MomentUltimate = 0#
    mResults(index).CapacityStateAvailable = False
    mResults(index).CapacityEpsilon0 = 0#: mResults(index).CapacityKappaX = 0#: mResults(index).CapacityKappaY = 0#
    mResults(index).MinConcreteStrain = 0#: mResults(index).MaxConcreteStrain = 0#
    mResults(index).MinSteelStrain = 0#: mResults(index).MaxSteelStrain = 0#
    mResults(index).CrackLsRaw = 0#: mResults(index).CrackLs = 0#: mResults(index).CrackLambdaCrc = 0#: mResults(index).CrackNcrc = 0#
    mResults(index).CrackTensionRebarCount = 0#
    mResults(index).CrackSectionDepthH = 0#: mResults(index).CrackCoverA = 0#: mResults(index).CrackTensionDepth = 0#
    mResults(index).CrackEffectiveZoneDepth = 0#: mResults(index).CrackCentralBranch = False
    mResults(index).GoverningUtil = 0#
    mResults(index).StateAvailable = False
    mResults(index).StateConverged = False
    mResults(index).Epsilon0 = 0#: mResults(index).KappaX = 0#: mResults(index).KappaY = 0#
    mResults(index).Nint = 0#: mResults(index).Mxint = 0#: mResults(index).Myint = 0#
    mResults(index).ResidualN = 0#: mResults(index).ResidualMx = 0#: mResults(index).ResidualMy = 0#
    mResults(index).StateIterations = 0
    mResults(index).ExtensionUsed = False
    mResults(index).DirectStateDiagramExtensionEnabled = False
End Sub

' Сохраняет найденную плоскость деформаций и внутренние усилия прямого solve.
' Эти данные затем уходят в Results и используются для схемы/AutoCAD export.
Private Sub StoreSectionState(ByVal index As Long, ByVal solver As CSectionSolver)
    mResults(index).StateAvailable = True
    mResults(index).StateConverged = solver.Converged
    mResults(index).Epsilon0 = solver.Epsilon0
    mResults(index).KappaX = solver.KappaX
    mResults(index).KappaY = solver.KappaY
    mResults(index).Nint = solver.Nint
    mResults(index).Mxint = solver.Mxint
    mResults(index).Myint = solver.Myint
    mResults(index).MinConcreteStrain = solver.MinConcreteStrain
    mResults(index).MaxConcreteStrain = solver.MaxConcreteStrain
    mResults(index).MinSteelStrain = solver.MinSteelStrain
    mResults(index).MaxSteelStrain = solver.MaxSteelStrain
    mResults(index).ResidualN = solver.ResidualN
    mResults(index).ResidualMx = solver.ResidualMx
    mResults(index).ResidualMy = solver.ResidualMy
    mResults(index).StateIterations = solver.Iterations
End Sub

' Добавляет одну строку в текстовую диагностику batch-расчета.
' Это компактная сводка для отладки, а не подробный журнал итераций solver-а.
Private Sub AppendDiagnostic(ByVal index As Long)
    mDiagnosticLog = mDiagnosticLog & "combination=" & mInputs(index).CombinationID & _
        "; OverallStatus=" & mResults(index).Status & _
        "; DirectStateStatus=" & mResults(index).DirectStateStatus & _
        "; ExtensionUsed=" & CStr(mResults(index).ExtensionUsed) & _
        "; CapacityStatus=" & mResults(index).CapacityStatus & _
        "; CrackStatus=" & mResults(index).CrackStatus & _
        "; governingUtil=" & FormatNumberInvariant(mResults(index).GoverningUtil) & vbCrLf
End Sub

' Добавляет в подробный txt-отчет краткий итог по текущему LC.
' Batch summary остается основным табличным выводом, а здесь пользователь
' видит ход выполнения без открытия Results и без чтения внутренних массивов.
Private Sub ReportCombinationSummary(ByVal index As Long)
    ReportValue "Итоговый Status", mResults(index).Status
    ReportValue "DirectStateStatus", mResults(index).DirectStateStatus
    ReportValue "ExtensionUsed", CStr(mResults(index).ExtensionUsed)
    ReportValue "CapacityStatus", mResults(index).CapacityStatus
    ReportValue "CrackStatus", mResults(index).CrackStatus
    ReportValue "GoverningUtilization", FormatNumberInvariant(mResults(index).GoverningUtil)
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

' Нормализует область запуска capacity-поиска. Алиасы оставлены на границе
' Excel-ввода, чтобы в расчетной логике дальше ходили только два значения.
Private Function NormalizedCapacityCalculationScope() As String
    Select Case LCase$(Trim$(mCapacityCalculationScope))
        Case "group1only", "group1", "uls"
            NormalizedCapacityCalculationScope = "Group1Only"
        Case "group1+2", "group1andgroup2", "group1+group2", "all", "both"
            NormalizedCapacityCalculationScope = "Group1AndGroup2"
    End Select
End Function

' Утилитарный максимум без WorksheetFunction, чтобы не тащить Excel в hot path.
Private Function MaxDouble(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then MaxDouble = a Else MaxDouble = b
End Function

Private Function MinDouble(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then MinDouble = a Else MinDouble = b
End Function

' Выбирает худшее сочетание для batch summary и режима Worst.
' По пользовательскому требованию Worst относится именно к прочности: после
' отказа от пользовательского StrainSafetyFactor выбор идет по запасу
' найденного предельного момента, а трещины остаются отдельной проверкой.
Private Function IsBetterGoverningCandidate(ByVal candidateIndex As Long, ByVal currentIndex As Long) As Boolean
    Dim candidateSafety As Double
    Dim currentSafety As Double
    candidateSafety = StrengthSafetyFactorForGoverning(candidateIndex)
    currentSafety = StrengthSafetyFactorForGoverning(currentIndex)

    If candidateSafety <= 0# Then
        IsBetterGoverningCandidate = False
    ElseIf currentSafety <= 0# Then
        IsBetterGoverningCandidate = True
    Else
        IsBetterGoverningCandidate = (candidateSafety < currentSafety)
    End If
End Function

' Возвращает прочностной запас LC как найденный множитель lambda.
' После перехода на CapacityLoadPath это единая величина для lambda*Mx,
' lambda*My, lambda*Mxy, lambda*N и lambda*NMxy; отдельные отношения Nult/N или Mult/Mxy больше
' не нужны для выбора Worst.
Private Function StrengthSafetyFactorForGoverning(ByVal index As Long) As Double
    If mStatusPolicy.IsFinished(mResults(index).CapacityStatus) And mResults(index).LambdaCapacity > 0# Then
        StrengthSafetyFactorForGoverning = mResults(index).LambdaCapacity
    End If
End Function

' Сравнивает только запасы по раскрытию трещин. Нулевой запас означает, что
' проверка трещин для LC не считалась или ширина раскрытия равна нулю.
Private Function IsBetterCrackGoverningCandidate(ByVal candidateIndex As Long, ByVal currentIndex As Long) As Boolean
    Dim candidateSafety As Double
    Dim currentSafety As Double
    candidateSafety = CrackSafetyFactorForGoverning(candidateIndex)
    If candidateSafety <= 0# Then Exit Function

    If currentIndex <= 0 Then
        IsBetterCrackGoverningCandidate = True
    Else
        currentSafety = CrackSafetyFactorForGoverning(currentIndex)
        IsBetterCrackGoverningCandidate = (currentSafety <= 0# Or candidateSafety < currentSafety)
    End If
End Function

' Возвращает запас по трещинам как величину, обратную использованию.
Private Function CrackSafetyFactorForGoverning(ByVal index As Long) As Double
    If mResults(index).CrackUtil > 0# Then CrackSafetyFactorForGoverning = 1# / mResults(index).CrackUtil
End Function

' Берет минимальное положительное значение из двух запасов.
' Нулевые значения означают, что соответствующая проверка не рассчитана.
Private Function PositiveMin(ByVal firstValue As Double, ByVal secondValue As Double) As Double
    If firstValue > 0# And secondValue > 0# Then
        PositiveMin = MinDouble(firstValue, secondValue)
    ElseIf firstValue > 0# Then
        PositiveMin = firstValue
    ElseIf secondValue > 0# Then
        PositiveMin = secondValue
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
    CombinationID = mInputs(index).CombinationID
End Property

Public Property Get CombinationName(ByVal index As Long) As String
    CombinationName = mInputs(index).Comment
End Property

Public Property Get CalculationType(ByVal index As Long) As String
    CalculationType = mInputs(index).CalculationType
End Property

Public Property Get CapacityLoadPath(ByVal index As Long) As String
    CapacityLoadPath = CapacityLoadPathDisplay(CapacityLoadPathKey(index))
End Property

Public Property Get CapacityLoadPathKey(ByVal index As Long) As String
    If Len(mResults(index).CapacityPathResolved) > 0 And mResults(index).CapacityPathResolved <> CAPACITY_PATH_NONE Then
        CapacityLoadPathKey = mResults(index).CapacityPathResolved
    Else
        CapacityLoadPathKey = BuildLoadPathForCombination(index).Key
    End If
End Property

Public Property Get N(ByVal index As Long) As Double
    N = mInputs(index).N
End Property

Public Property Get Mx(ByVal index As Long) As Double
    Mx = mInputs(index).Mx
End Property

Public Property Get My(ByVal index As Long) As Double
    My = mInputs(index).My
End Property

Public Property Get UserMx(ByVal index As Long) As Double
    UserMx = mInputs(index).UserMx
End Property

Public Property Get UserMy(ByVal index As Long) As Double
    UserMy = mInputs(index).UserMy
End Property

Public Property Get LoadReferenceX() As Double
    LoadReferenceX = mLoadReferenceX
End Property

Public Property Get LoadReferenceY() As Double
    LoadReferenceY = mLoadReferenceY
End Property

Public Property Get LoadReferenceOffsetX() As Double
    LoadReferenceOffsetX = mLoadReferenceOffsetX
End Property

Public Property Get LoadReferenceOffsetY() As Double
    LoadReferenceOffsetY = mLoadReferenceOffsetY
End Property

Public Property Get Status(ByVal index As Long) As String
    Status = mResults(index).Status
End Property

Public Property Get OverallStatus(ByVal index As Long) As String
    OverallStatus = mResults(index).Status
End Property

Public Property Get DirectStateStatus(ByVal index As Long) As String
    DirectStateStatus = mResults(index).DirectStateStatus
End Property

Public Property Get CapacityStatus(ByVal index As Long) As String
    CapacityStatus = mResults(index).CapacityStatus
End Property

Public Property Get CapacityLimitState(ByVal index As Long) As String
    CapacityLimitState = mResults(index).CapacityLimitState
End Property

Public Property Get CapacitySolutionMethod(ByVal index As Long) As String
    CapacitySolutionMethod = mResults(index).CapacitySolutionMethod
End Property

Public Property Get CrackStatus(ByVal index As Long) As String
    CrackStatus = mResults(index).CrackStatus
End Property

Public Property Get ExtensionUsed(ByVal index As Long) As Boolean
    ExtensionUsed = mResults(index).ExtensionUsed
End Property

Public Property Get DirectStateDiagramExtensionEnabled(ByVal index As Long) As Boolean
    DirectStateDiagramExtensionEnabled = mResults(index).DirectStateDiagramExtensionEnabled
End Property

Public Property Get UtilizationCapacity(ByVal index As Long) As Double
    UtilizationCapacity = mResults(index).UtilCapacity
End Property

Public Property Get NUltimate(ByVal index As Long) As Double
    NUltimate = mResults(index).NUltimate
End Property

Public Property Get MxUltimate(ByVal index As Long) As Double
    MxUltimate = mResults(index).MxUltimate
End Property

Public Property Get MyUltimate(ByVal index As Long) As Double
    MyUltimate = mResults(index).MyUltimate
End Property

Public Property Get MomentUltimate(ByVal index As Long) As Double
    MomentUltimate = mResults(index).MomentUltimate
End Property

Public Property Get CapacityStateAvailable(ByVal index As Long) As Boolean
    CapacityStateAvailable = mResults(index).CapacityStateAvailable
End Property

Public Property Get CapacityEpsilon0(ByVal index As Long) As Double
    CapacityEpsilon0 = mResults(index).CapacityEpsilon0
End Property

Public Property Get CapacityKappaX(ByVal index As Long) As Double
    CapacityKappaX = mResults(index).CapacityKappaX
End Property

Public Property Get CapacityKappaY(ByVal index As Long) As Double
    CapacityKappaY = mResults(index).CapacityKappaY
End Property

Public Property Get MinConcreteStrain(ByVal index As Long) As Double
    MinConcreteStrain = mResults(index).MinConcreteStrain
End Property

Public Property Get MaxConcreteStrain(ByVal index As Long) As Double
    MaxConcreteStrain = mResults(index).MaxConcreteStrain
End Property

Public Property Get MinSteelStrain(ByVal index As Long) As Double
    MinSteelStrain = mResults(index).MinSteelStrain
End Property

Public Property Get MaxSteelStrain(ByVal index As Long) As Double
    MaxSteelStrain = mResults(index).MaxSteelStrain
End Property

Public Property Get CrackWidth(ByVal index As Long) As Double
    CrackWidth = mResults(index).CrackWidth
End Property

Public Property Get CrackAllowable(ByVal index As Long) As Double
    CrackAllowable = mResults(index).CrackAllowableValue
End Property

Public Property Get CrackUtilization(ByVal index As Long) As Double
    CrackUtilization = mResults(index).CrackUtil
End Property

Public Property Get CrackFormed(ByVal index As Long) As Boolean
    CrackFormed = mResults(index).CrackFormed
End Property

Public Property Get CrackPhi1(ByVal index As Long) As Double
    CrackPhi1 = mResults(index).CrackPhi1
End Property

Public Property Get CrackPhi2(ByVal index As Long) As Double
    CrackPhi2 = mResults(index).CrackPhi2
End Property

Public Property Get CrackPhi3(ByVal index As Long) As Double
    CrackPhi3 = mResults(index).CrackPhi3
End Property

Public Property Get CrackPsiS(ByVal index As Long) As Double
    CrackPsiS = mResults(index).CrackPsiS
End Property

Public Property Get CrackSigmaS(ByVal index As Long) As Double
    CrackSigmaS = mResults(index).CrackSigmaS
End Property

Public Property Get CrackSigmaSCrc(ByVal index As Long) As Double
    CrackSigmaSCrc = mResults(index).CrackSigmaSCrc
End Property

Public Property Get CrackAs(ByVal index As Long) As Double
    CrackAs = mResults(index).CrackAs
End Property

Public Property Get CrackAbt(ByVal index As Long) As Double
    CrackAbt = mResults(index).CrackAbt
End Property

Public Property Get CrackDsEquivalent(ByVal index As Long) As Double
    CrackDsEquivalent = mResults(index).CrackDsEquivalent
End Property

Public Property Get CrackSpacingRaw(ByVal index As Long) As Double
    CrackSpacingRaw = mResults(index).CrackLsRaw
End Property

Public Property Get CrackSpacing(ByVal index As Long) As Double
    CrackSpacing = mResults(index).CrackLs
End Property

Public Property Get CrackLambdaCrc(ByVal index As Long) As Double
    CrackLambdaCrc = mResults(index).CrackLambdaCrc
End Property

Public Property Get CrackNcrc(ByVal index As Long) As Double
    CrackNcrc = mResults(index).CrackNcrc
End Property

Public Property Get CrackTensionRebarCount(ByVal index As Long) As Long
    CrackTensionRebarCount = mResults(index).CrackTensionRebarCount
End Property

Public Property Get CrackSectionDepthH(ByVal index As Long) As Double
    CrackSectionDepthH = mResults(index).CrackSectionDepthH
End Property

Public Property Get CrackCoverA(ByVal index As Long) As Double
    CrackCoverA = mResults(index).CrackCoverA
End Property

Public Property Get CrackTensionDepth(ByVal index As Long) As Double
    CrackTensionDepth = mResults(index).CrackTensionDepth
End Property

Public Property Get CrackEffectiveZoneDepth(ByVal index As Long) As Double
    CrackEffectiveZoneDepth = mResults(index).CrackEffectiveZoneDepth
End Property

Public Property Get CrackCentralBranch(ByVal index As Long) As Boolean
    CrackCentralBranch = mResults(index).CrackCentralBranch
End Property

Public Property Get GoverningUtilization(ByVal index As Long) As Double
    GoverningUtilization = mResults(index).GoverningUtil
End Property

Public Property Get CrackSteelElasticModulus() As Double
    CrackSteelElasticModulus = mMaterialProvider.SteelElasticModulus(cpCrackedNDS)
End Property

' Предельная сжатая деформация бетона для фактического НДС выбранного LC.
Public Function ConcreteCompressionLimitForCombination(ByVal index As Long) As Double
    ConcreteCompressionLimitForCombination = mMaterialProvider.ConcreteCompressionLimit(MaterialPurposeForCombination(index))
End Function

' Предельная растянутая деформация бетона для фактического НДС выбранного LC.
Public Function ConcreteTensionLimitForCombination(ByVal index As Long) As Double
    ConcreteTensionLimitForCombination = mMaterialProvider.ConcreteTensionLimit(MaterialPurposeForCombination(index))
End Function

' True, если текущая диаграмма LC учитывает растянутый бетон и его предел нужно выводить.
Public Function ConcreteTensionLimitEnabledForCombination(ByVal index As Long) As Boolean
    ConcreteTensionLimitEnabledForCombination = mMaterialProvider.ConcreteTensionLimitEnabled(MaterialPurposeForCombination(index))
End Function

' Предельная сжатая деформация арматуры для фактического НДС выбранного LC.
Public Function SteelCompressionLimitForCombination(ByVal index As Long) As Double
    SteelCompressionLimitForCombination = mMaterialProvider.SteelCompressionLimit(MaterialPurposeForCombination(index))
End Function

' Предельная растянутая деформация арматуры для фактического НДС выбранного LC.
Public Function SteelTensionLimitForCombination(ByVal index As Long) As Double
    SteelTensionLimitForCombination = mMaterialProvider.SteelTensionLimit(MaterialPurposeForCombination(index))
End Function

Public Property Get GoverningCombinationIndex() As Long
    GoverningCombinationIndex = mGoverningCombinationIndex
End Property

Public Property Get GoverningCombinationID() As String
    If mGoverningCombinationIndex > 0 Then GoverningCombinationID = mInputs(mGoverningCombinationIndex).CombinationID
End Property

Public Property Get GoverningCombinationName() As String
    If mGoverningCombinationIndex > 0 Then GoverningCombinationName = mInputs(mGoverningCombinationIndex).Comment
End Property

Public Property Get CrackGoverningCombinationIndex() As Long
    CrackGoverningCombinationIndex = mCrackGoverningCombinationIndex
End Property

Public Property Get CrackGoverningCombinationID() As String
    If mCrackGoverningCombinationIndex > 0 Then CrackGoverningCombinationID = mInputs(mCrackGoverningCombinationIndex).CombinationID
End Property

Public Property Get CrackGoverningCombinationName() As String
    If mCrackGoverningCombinationIndex > 0 Then CrackGoverningCombinationName = mInputs(mCrackGoverningCombinationIndex).Comment
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
        If mStatusPolicy.IsInvalid(mResults(i).Status) Then InvalidInputCount = InvalidInputCount + 1
    Next i
End Property

Public Property Get FirstInvalidInputMessage() As String
    Dim i As Long
    For i = 1 To mCount
        If mStatusPolicy.IsInvalid(mResults(i).Status) Then
            FirstInvalidInputMessage = mInputs(i).InputError
            If Len(Trim$(FirstInvalidInputMessage)) = 0 Then FirstInvalidInputMessage = mResults(i).CapacityLimitState
            If Len(Trim$(FirstInvalidInputMessage)) = 0 Then FirstInvalidInputMessage = mResults(i).Status
            Exit Property
        End If
    Next i
End Property

Public Property Get StateAvailable(ByVal index As Long) As Boolean
    StateAvailable = mResults(index).StateAvailable
End Property

Public Property Get StateAvailableCount() As Long
    Dim i As Long
    For i = 1 To mCount
        If mResults(i).StateAvailable Then StateAvailableCount = StateAvailableCount + 1
    Next i
End Property

Public Property Get Epsilon0(ByVal index As Long) As Double
    Epsilon0 = mResults(index).Epsilon0
End Property

Public Property Get KappaX(ByVal index As Long) As Double
    KappaX = mResults(index).KappaX
End Property

Public Property Get KappaY(ByVal index As Long) As Double
    KappaY = mResults(index).KappaY
End Property

Public Property Get Nint(ByVal index As Long) As Double
    Nint = mResults(index).Nint
End Property

Public Property Get Mxint(ByVal index As Long) As Double
    Mxint = mResults(index).Mxint
End Property

Public Property Get Myint(ByVal index As Long) As Double
    Myint = mResults(index).Myint
End Property

Public Property Get ResidualN(ByVal index As Long) As Double
    ResidualN = mResults(index).ResidualN
End Property

Public Property Get ResidualMx(ByVal index As Long) As Double
    ResidualMx = mResults(index).ResidualMx
End Property

Public Property Get ResidualMy(ByVal index As Long) As Double
    ResidualMy = mResults(index).ResidualMy
End Property

Public Property Get StateIterations(ByVal index As Long) As Long
    StateIterations = mResults(index).StateIterations
End Property

' ==========================================================================
' ДЛЯ ТЕСТОВ
' ==========================================================================
' Эти свойства раскрывают промежуточные численные данные batch-расчета,
' которые удобны regression-тестам: отдельно проверяются lambda capacity
' и факт сходимости прямого НДС.

Public Property Get LambdaCapacity(ByVal index As Long) As Double
    LambdaCapacity = mResults(index).LambdaCapacity
End Property

Public Property Get StateConverged(ByVal index As Long) As Boolean
    StateConverged = mResults(index).StateConverged
End Property



