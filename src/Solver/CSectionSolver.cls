VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSectionSolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Решатель равновесия N + Mx + My
' ==========================================================================
' Класс подбирает три параметра плоскости деформаций: epsilon0, kappaX и
' kappaY. По этой плоскости считаются деформации бетонных волокон и арматуры,
' затем материалы возвращают напряжения, а из них собираются внутренние
' усилия Nint, Mxint и Myint.
'
' Здесь нет поиска несущей способности и нет чтения Excel. Capacity solver
' только вызывает этот класс с разными усилиями, а пользовательские единицы
' заранее переводятся во внутренние мм, Н, Н*мм, МПа и 1/мм.

Private mConverged As Boolean ' True, если равновесие достигнуто в заданных допусках по N/Mx/My.
Private mStopReason As String ' Человекочитаемая причина остановки: сходимость, InputError или NumericalFailure.
Private mFailureCode As ESolverFailureCode ' Машинная причина неуспеха без парсинга текста StopReason.
Private mEpsilon0 As Double ' Равномерная составляющая деформации в формуле epsilon = epsilon0 + kappaX*y + kappaY*x.
Private mKappaX As Double ' Кривизна, которая умножается на координату y и дает вклад изгиба вокруг X.
Private mKappaY As Double ' Кривизна, которая умножается на координату x и дает вклад изгиба вокруг Y.
Private mNint As Double ' Внутренняя продольная сила от всех бетонных и арматурных элементов.
Private mMxint As Double ' Внутренний момент Mx, собранный интегрированием напряжений по элементам.
Private mMyint As Double ' Внутренний момент My, собранный интегрированием напряжений по элементам.
Private mResidualN As Double ' Невязка равновесия по силе: Nint - Ntarget.
Private mResidualMx As Double ' Невязка равновесия по моменту: Mxint - Mxtarget.
Private mResidualMy As Double ' Невязка равновесия по моменту: Myint - Mytarget.
Private mRelativeResidualN As Double ' Относительная невязка N для диагностики качества найденного состояния.
Private mRelativeResidualMx As Double ' Относительная невязка Mx для диагностики качества найденного состояния.
Private mRelativeResidualMy As Double ' Относительная невязка My для диагностики качества найденного состояния.
Private mMinConcreteStrain As Double ' Минимальная деформация бетона в последнем посчитанном состоянии.
Private mMaxConcreteStrain As Double ' Максимальная деформация бетона в последнем посчитанном состоянии.
Private mMaxConcreteCompressionStress As Double ' Максимальное по модулю сжимающее напряжение бетона в последнем состоянии.
Private mMinConcreteStress As Double ' Наиболее сжимающее напряжение бетона со знаком, МПа.
Private mMaxConcreteStress As Double ' Наиболее растягивающее напряжение бетона со знаком, МПа.
Private mHasConcreteCompression As Boolean ' True, если в последнем состоянии реально есть сжатый бетон.
Private mHasConcreteTension As Boolean ' True, если бетон реально работает на растяжение, а не выключен диаграммой.
Private mMinSteelStrain As Double ' Наиболее сжатая деформация арматуры в последнем посчитанном состоянии.
Private mMaxSteelStrain As Double ' Наиболее растянутая деформация арматуры в последнем посчитанном состоянии.
Private mMaxAbsSteelStrain As Double ' Максимальная по модулю деформация арматуры в последнем состоянии.
Private mMinSteelStress As Double ' Наиболее сжимающее напряжение стали со знаком, МПа.
Private mMaxSteelStress As Double ' Наиболее растягивающее напряжение стали со знаком, МПа.
Private mHasSteelCompression As Boolean ' True, если есть сжатая арматура.
Private mHasSteelTension As Boolean ' True, если есть растянутая арматура.
Private mIterations As Long ' Суммарное число итераций на всех ступенях текущего вызова Solve.
Private mLoadStepsCompleted As Long ' Сколько ступеней нагрузки успешно доведено до равновесия.
Private mK(1 To 3, 1 To 3) As Double ' Касательная матрица d(N,Mx,My)/d(epsilon0,kappaX,kappaY).
Private mMaxIterations As Long ' Максимум итераций на одну ступень нагрузки.
Private mLoadSteps As Long ' На сколько ступеней дробится заданный вектор N/Mx/My.
Private mToleranceN As Double ' Абсолютный допуск равновесия по N во внутренних единицах.
Private mToleranceMx As Double ' Абсолютный допуск равновесия по Mx во внутренних единицах.
Private mToleranceMy As Double ' Абсолютный допуск равновесия по My во внутренних единицах.
Private mLineSearchEnabled As Boolean ' Разрешает дробить шаг, если полный шаг ухудшает невязку.
Private mDampingInitial As Double ' Начальный множитель шага перед line search.
Private mMinLineSearchAlpha As Double ' Минимальный alpha, ниже которого шаг считается неудачным.
Private mMaxDeltaEpsilon0 As Double ' Ограничение приращения epsilon0 за одну итерацию.
Private mMaxDeltaKappa As Double ' Ограничение приращения kappaX/kappaY за одну итерацию.
Private mDiagnosticsEnabled As Boolean ' Управляет накоплением текстовой диагностики итераций.
Private mSolverMethod As String ' Выбранная численная ветка: Newton или Secant.
Private mSecantMaxRestarts As Long ' Сколько раз Secant может пересобрать приближенную матрицу чувствительности.
Private mSecantMinStepNorm As Double ' Минимальный размер secant-шага; меньше считается застоем.
Private mDiagnosticLog As String ' Текстовый журнал ключевых итераций для тестов и диагностики.
Private mLastAlpha As Double ' Последний принятый коэффициент line search.
Private mLastDeltaScale As Double ' Во сколько раз был уменьшен последний шаг из-за лимитов приращений.
Private mLastResidualNorm As Double ' Норма невязки после последней оценки состояния.
Private mLastStepNorm As Double ' Норма последней поправки [deltaEps, deltaKx, deltaKy].
Private mLastMatrixConditionEstimate As Double ' Грубая оценка вырожденности касательной матрицы.
Private mMatrixRestartCount As Long ' Сколько раз Secant перезапускал матрицу чувствительности.
Private mLineSearchUsed As Boolean ' True, если последний принятый шаг был меньше полного.
Private mLineSearchIterations As Long ' Сколько пробных alpha понадобилось последнему line search.
Private mInternalForceEvaluationCount As Long ' Сколько раз пересчитывались внутренние усилия.
Private mInternalNewtonCallCount As Long ' Сколько раз запускалась ветка Newton на ступени нагрузки.
Private mUseInitialState As Boolean ' Использовать заранее заданную стартовую плоскость деформаций.
Private mInitialEpsilon0 As Double ' Стартовое epsilon0 для продолжения расчета от известного состояния.
Private mInitialKappaX As Double ' Стартовая kappaX для продолжения расчета от известного состояния.
Private mInitialKappaY As Double ' Стартовая kappaY для продолжения расчета от известного состояния.
Private mUseInitialLoads As Boolean ' Использовать ненулевую начальную точку приложения ступеней нагрузки.
Private mInitialTargetN As Double ' Начальное N при продолжении расчета от уже уравновешенной точки.
Private mInitialTargetMx As Double ' Начальное Mx при продолжении расчета от уже уравновешенной точки.
Private mInitialTargetMy As Double ' Начальное My при продолжении расчета от уже уравновешенной точки.

' ------------------------------
' Настройки по умолчанию
' ------------------------------
' Значения достаточно устойчивы для тестовой книги. Excel-слой может заменить
' их через ApplySettings перед запуском расчета.
Private Sub Class_Initialize()
    mMaxIterations = 80
    mLoadSteps = 5
    mToleranceN = 1#
    mToleranceMx = 1000#
    mToleranceMy = 1000#
    mLineSearchEnabled = True
    mDampingInitial = 1#
    mMinLineSearchAlpha = 0.03125
    mMaxDeltaEpsilon0 = 0.0005
    mMaxDeltaKappa = 0.00001
    mDiagnosticsEnabled = True
    mSolverMethod = "Newton"
    mSecantMaxRestarts = 2
    mSecantMinStepNorm = 0.000000000001
End Sub

' ------------------------------
' Чтение настроек решателя
' ------------------------------
' Забирает численные параметры из CSystemSettingsReader. Если передан
' CUnitSystem, пользовательские допуски переводятся во внутренние единицы,
' чтобы дальше solver работал только в одной системе.
Public Sub ApplySettings(ByVal settings As CSystemSettingsReader, Optional ByVal units As CUnitSystem = Nothing)
    If settings Is Nothing Then Exit Sub
    mMaxIterations = settings.GetLong("Solver.MaxIterations", mMaxIterations)
    mLoadSteps = settings.GetLong("Solver.LoadSteps", mLoadSteps)
    mToleranceN = settings.GetDouble("Solver.ToleranceN", mToleranceN)
    mToleranceMx = settings.GetDouble("Solver.ToleranceMx", mToleranceMx)
    mToleranceMy = settings.GetDouble("Solver.ToleranceMy", mToleranceMy)
    mLineSearchEnabled = settings.GetBoolean("Solver.LineSearchEnabled", mLineSearchEnabled)
    mDampingInitial = settings.GetDouble("Solver.DampingInitial", mDampingInitial)
    mMinLineSearchAlpha = settings.GetDouble("Solver.MinLineSearchAlpha", mMinLineSearchAlpha)
    mMaxDeltaEpsilon0 = settings.GetDouble("Solver.MaxDeltaEpsilon0", mMaxDeltaEpsilon0)
    mMaxDeltaKappa = settings.GetDouble("Solver.MaxDeltaKappa", mMaxDeltaKappa)
    If settings.HasKey("Solver.Method") Then
        mSolverMethod = settings.GetRawString("Solver.Method", vbNullString)
    Else
        mSolverMethod = "Newton"
    End If
    mSecantMaxRestarts = settings.GetLong("Solver.SecantMaxRestarts", mSecantMaxRestarts)
    mSecantMinStepNorm = settings.GetDouble("Solver.SecantMinStepNorm", mSecantMinStepNorm)
    If Not units Is Nothing Then
        mToleranceN = Abs(units.InputForceToInternal(mToleranceN))
        mToleranceMx = Abs(units.InputMomentMxToInternal(mToleranceMx))
        mToleranceMy = Abs(units.InputMomentMyToInternal(mToleranceMy))
        mMaxDeltaKappa = Abs(units.InputCurvatureToInternal(mMaxDeltaKappa))
    End If
End Sub

' ------------------------------
' Основной расчет равновесия
' ------------------------------
' Подбирает epsilon0, kappaX и kappaY для заданных N, Mx, My. Нагрузка может
' прикладываться ступенями: каждая следующая ступень стартует с предыдущего
' состояния, что повышает устойчивость на нелинейных диаграммах материалов.
Public Sub Solve(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double)

    On Error GoTo Failed
    ClearResult
    If Not ValidateInputs(section, concreteMaterial, steelMaterial) Then Exit Sub
    If Not ValidateSolverMethod() Then Exit Sub
    RecordSectionEquilibriumSolve
    If mUseInitialState Then
        mEpsilon0 = mInitialEpsilon0
        mKappaX = mInitialKappaX
        mKappaY = mInitialKappaY
    End If

    Dim stepIndex As Long
    Dim stepN As Double
    Dim stepMx As Double
    Dim stepMy As Double
    Dim startN As Double
    Dim startMx As Double
    Dim startMy As Double

    If mUseInitialLoads Then
        startN = mInitialTargetN
        startMx = mInitialTargetMx
        startMy = mInitialTargetMy
    End If

    ' Нагрузка прикладывается ступенями, чтобы для нелинейной диаграммы материал не требовал
    ' от Newton/Secant сразу найти равновесие на полном векторе усилий из далекого начального состояния.
    For stepIndex = 1 To mLoadSteps
        stepN = startN + (targetN - startN) * stepIndex / mLoadSteps
        stepMx = startMx + (targetMx - startMx) * stepIndex / mLoadSteps
        stepMy = startMy + (targetMy - startMy) * stepIndex / mLoadSteps

        If Not SolveSelectedLoadStep(section, concreteMaterial, steelMaterial, stepIndex, stepN, stepMx, stepMy) Then
            Exit Sub
        End If
        mLoadStepsCompleted = stepIndex
    Next stepIndex

    ' Финальная проверка выполняется уже на полной нагрузке: промежуточная сходимость ступеней
    ' сама по себе еще не гарантирует, что итоговые невязки попали в пользовательские допуски.
    EvaluateState section, concreteMaterial, steelMaterial, mEpsilon0, mKappaX, mKappaY, True
    SetResiduals targetN, targetMx, targetMy
    mConverged = IsConverged(targetN, targetMx, targetMy)
    If mConverged Then
        mStopReason = "Равновесие найдено методом " & mSolverMethod & " со ступенчатым приложением нагрузки."
        mFailureCode = sfcNone
    Else
        mStopReason = "Финальная проверка невязок равновесия не пройдена."
        mFailureCode = sfcNumericalFailure
    End If
    Exit Sub
Failed:
    mConverged = False
    mFailureCode = sfcInternalError
    mStopReason = "Ошибка при вычислении равновесия: " & Err.Description
End Sub

' ------------------------------
' Оценка заданной плоскости деформаций
' ------------------------------
' Используется, когда epsilon0/kappaX/kappaY уже известны и нужно только
' посчитать внутренние усилия, напряжения и предельные деформации без
' итерационного подбора равновесия.
Public Sub EvaluateStrainPlane(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal epsilon0 As Double, ByVal kappaX As Double, ByVal kappaY As Double)

    ClearResult
    If Not ValidateInputs(section, concreteMaterial, steelMaterial) Then Exit Sub
    mEpsilon0 = epsilon0
    mKappaX = kappaX
    mKappaY = kappaY
    EvaluateState section, concreteMaterial, steelMaterial, mEpsilon0, mKappaX, mKappaY, True
    SetResiduals mNint, mMxint, mMyint
    mConverged = False
    mFailureCode = sfcNone
    mStopReason = "Плоскость деформаций оценена; равновесие с заданной нагрузкой еще не проверено."
End Sub

' Проверяет уже оцененную плоскость по заданным усилиям без нового solve.
' Только эта проверка превращает EvaluateStrainPlane в подтвержденное
' равновесие; напряжения и деформации при этом не меняются.
Public Function ConfirmEquilibrium(ByVal targetN As Double, ByVal targetMx As Double, _
        ByVal targetMy As Double) As Boolean
    SetResiduals targetN, targetMx, targetMy
    mLastResidualNorm = ResidualNorm(mResidualN, mResidualMx, mResidualMy, targetN, targetMx, targetMy)
    mConverged = IsConverged(targetN, targetMx, targetMy)
    If mConverged Then
        mFailureCode = sfcNone
        mStopReason = "Равновесие заданной плоскости подтверждено по N, Mx и My."
    Else
        mFailureCode = sfcNumericalFailure
        mStopReason = "Заданная плоскость не удовлетворяет равновесию по N, Mx и My."
    End If
    ConfirmEquilibrium = mConverged
End Function

' Выбирает Newton или Secant для одной ступени нагрузки. Неверное имя метода
' не подменяется запасным алгоритмом, а сразу дает InputError.
Private Function SolveSelectedLoadStep(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal stepIndex As Long, ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double) As Boolean

    If StrComp(mSolverMethod, "Newton", vbTextCompare) = 0 Then
        SolveSelectedLoadStep = SolveNewtonLoadStep(section, concreteMaterial, steelMaterial, stepIndex, targetN, targetMx, targetMy)
    ElseIf StrComp(mSolverMethod, "Secant", vbTextCompare) = 0 Then
        SolveSelectedLoadStep = SolveSecantLoadStep(section, concreteMaterial, steelMaterial, stepIndex, targetN, targetMx, targetMy)
    Else
        mConverged = False
        mFailureCode = sfcInvalidConfiguration
        mStopReason = "InputError: Solver.Method должен быть Newton или Secant."
        SolveSelectedLoadStep = False
    End If
End Function

' ------------------------------
' Шаг Newton
' ------------------------------
' На каждой итерации считаем текущую невязку равновесия, собираем касательную
' матрицу K и решаем K * delta = -R. Затем шаг ограничивается по величине и
' принимается через line search, чтобы не прыгнуть в состояние с худшей невязкой.
Private Function SolveNewtonLoadStep(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal stepIndex As Long, ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double) As Boolean

    mInternalNewtonCallCount = mInternalNewtonCallCount + 1

    Dim iter As Long
    For iter = 1 To mMaxIterations
        mIterations = mIterations + 1

        ' Считаем текущие внутренние усилия и касательную матрицу для этой
        ' плоскости деформаций. После этого можно проверить, достигнуто ли
        ' равновесие без дополнительного шага.
        EvaluateState section, concreteMaterial, steelMaterial, mEpsilon0, mKappaX, mKappaY, True
        SetResiduals targetN, targetMx, targetMy
        If IsConverged(targetN, targetMx, targetMy) Then
            AppendDiagnostics stepIndex, iter, 0#, 0#, targetN, targetMx, targetMy, "сошлось"
            SolveNewtonLoadStep = True
            Exit Function
        End If

        ' Решаем линейную поправку Newton. Правая часть равна минус невязке,
        ' потому что хотим сдвинуть состояние в сторону уменьшения R.
        Dim system As CLinearSystem3x3
        Set system = New CLinearSystem3x3
        If Not system.Solve(mK(1, 1), mK(1, 2), mK(1, 3), _
                mK(2, 1), mK(2, 2), mK(2, 3), _
                mK(3, 1), mK(3, 2), mK(3, 3), _
                -mResidualN, -mResidualMx, -mResidualMy) Then
            mConverged = False
            mFailureCode = sfcSingularTangent
            mStopReason = "Не удалось решить касательную систему: " & system.StopReason
            AppendDiagnostics stepIndex, iter, 0#, 0#, targetN, targetMx, targetMy, mStopReason
            SolveNewtonLoadStep = False
            Exit Function
        End If

        Dim deltaEps As Double
        Dim deltaKx As Double
        Dim deltaKy As Double
        deltaEps = system.X1
        deltaKx = system.X2
        deltaKy = system.X3

        ' Слишком большой шаг сначала режется по инженерным лимитам приращений,
        ' а затем line search подбирает alpha, который реально улучшает невязку.
        mLastDeltaScale = LimitIncrement(deltaEps, deltaKx, deltaKy)
        mLastAlpha = AcceptStep(section, concreteMaterial, steelMaterial, targetN, targetMx, targetMy, deltaEps, deltaKx, deltaKy)
        AppendDiagnostics stepIndex, iter, mLastAlpha, mLastDeltaScale, targetN, targetMx, targetMy, "итерация"
        If mLastAlpha <= 0# Then
            mConverged = False
            mFailureCode = sfcNumericalFailure
            mStopReason = "Line search не смог уменьшить невязку."
            SolveNewtonLoadStep = False
            Exit Function
        End If
    Next iter

    RefreshCurrentState section, concreteMaterial, steelMaterial, targetN, targetMx, targetMy
    mConverged = False
    mFailureCode = sfcNumericalFailure
    mStopReason = "Достигнут лимит итераций на ступени нагрузки " & CStr(stepIndex) & "."
    SolveNewtonLoadStep = False
End Function

' Решает одну ступень нагрузки методом Secant/Broyden без аналитической
' касательной матрицы. Метод использует численную начальную матрицу,
' обновляет ее по Бройдену и при необходимости ограниченно перезапускает.
Private Function SolveSecantLoadStep(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal stepIndex As Long, ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double) As Boolean

    Dim b(1 To 3, 1 To 3) As Double
    EvaluateSecantResidual section, concreteMaterial, steelMaterial, targetN, targetMx, targetMy, _
        mEpsilon0, mKappaX, mKappaY, mResidualN, mResidualMx, mResidualMy

    If Not BuildInitialSecantMatrix(section, concreteMaterial, steelMaterial, targetN, targetMx, targetMy, _
            mEpsilon0, mKappaX, mKappaY, mResidualN, mResidualMx, mResidualMy, b) Then
        SolveSecantLoadStep = False
        Exit Function
    End If

    Dim iter As Long
    For iter = 1 To mMaxIterations
        mIterations = mIterations + 1
        mLastResidualNorm = ResidualNorm(mResidualN, mResidualMx, mResidualMy, targetN, targetMx, targetMy)

        If IsConverged(targetN, targetMx, targetMy) Then
            AppendDiagnostics stepIndex, iter, 0#, 0#, targetN, targetMx, targetMy, "Secant сошелся"
            SolveSecantLoadStep = True
            Exit Function
        End If

        Dim deltaEps As Double
        Dim deltaKx As Double
        Dim deltaKy As Double
        If Not SolveSecantCorrection(b, deltaEps, deltaKx, deltaKy) Then
            If Not RestartSecantMatrix(section, concreteMaterial, steelMaterial, targetN, targetMx, targetMy, _
                    mEpsilon0, mKappaX, mKappaY, mResidualN, mResidualMx, mResidualMy, b, "линейная система не решена") Then
                SolveSecantLoadStep = False
                Exit Function
            End If
            If Not SolveSecantCorrection(b, deltaEps, deltaKx, deltaKy) Then
                mConverged = False
                mFailureCode = sfcSingularTangent
                mStopReason = "NumericalFailure: матрица Secant не решилась после перезапуска."
                SolveSecantLoadStep = False
                Exit Function
            End If
        End If

        Dim oldEps As Double
        Dim oldKx As Double
        Dim oldKy As Double
        Dim oldRN As Double
        Dim oldRMx As Double
        Dim oldRMy As Double
        oldEps = mEpsilon0: oldKx = mKappaX: oldKy = mKappaY
        oldRN = mResidualN: oldRMx = mResidualMx: oldRMy = mResidualMy

        mLastDeltaScale = LimitIncrement(deltaEps, deltaKx, deltaKy)
        mLastStepNorm = SecantStepNorm(deltaEps, deltaKx, deltaKy)
        If Not IsFiniteNumber(mLastStepNorm) Or mLastStepNorm <= mSecantMinStepNorm Then
            If Not RestartSecantMatrix(section, concreteMaterial, steelMaterial, targetN, targetMx, targetMy, _
                    mEpsilon0, mKappaX, mKappaY, mResidualN, mResidualMx, mResidualMy, b, "шаг слишком мал") Then
                SolveSecantLoadStep = False
                Exit Function
            End If
            AppendDiagnostics stepIndex, iter, 0#, mLastDeltaScale, targetN, targetMx, targetMy, "Secant перезапущен"
        Else
            mLastAlpha = AcceptStep(section, concreteMaterial, steelMaterial, targetN, targetMx, targetMy, deltaEps, deltaKx, deltaKy)
            If mLastAlpha <= 0# Then
                If Not RestartSecantMatrix(section, concreteMaterial, steelMaterial, targetN, targetMx, targetMy, _
                        oldEps, oldKx, oldKy, oldRN, oldRMx, oldRMy, b, "line search не дал улучшения") Then
                    SolveSecantLoadStep = False
                    Exit Function
                End If
            Else
                EvaluateSecantResidual section, concreteMaterial, steelMaterial, targetN, targetMx, targetMy, _
                    mEpsilon0, mKappaX, mKappaY, mResidualN, mResidualMx, mResidualMy
                If Not UpdateBroydenMatrix(b, mEpsilon0 - oldEps, mKappaX - oldKx, mKappaY - oldKy, _
                        mResidualN - oldRN, mResidualMx - oldRMx, mResidualMy - oldRMy) Then
                    If Not RestartSecantMatrix(section, concreteMaterial, steelMaterial, targetN, targetMx, targetMy, _
                            mEpsilon0, mKappaX, mKappaY, mResidualN, mResidualMx, mResidualMy, b, "обновление Broyden не выполнено") Then
                        SolveSecantLoadStep = False
                        Exit Function
                    End If
                End If
                AppendDiagnostics stepIndex, iter, mLastAlpha, mLastDeltaScale, targetN, targetMx, targetMy, "итерация Secant"
            End If
        End If
    Next iter

    RefreshCurrentState section, concreteMaterial, steelMaterial, targetN, targetMx, targetMy
    mConverged = False
    mFailureCode = sfcNumericalFailure
    mStopReason = "NumericalFailure: метод Secant достиг лимита итераций на ступени нагрузки " & CStr(stepIndex) & "."
    SolveSecantLoadStep = False
End Function

' Пересчитывает внутренние усилия для текущей плоскости деформаций и сразу
' обновляет невязки относительно целевого вектора нагрузки.
Private Sub RefreshCurrentState(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double)
    EvaluateState section, concreteMaterial, steelMaterial, mEpsilon0, mKappaX, mKappaY, True
    SetResiduals targetN, targetMx, targetMy
End Sub

' Служебно оценивает невязку Secant в произвольной плоскости деформаций.
' Основное состояние solver-а после вызова остается синхронизированным через
' SetResiduals, чтобы диагностика и последующий шаг видели актуальные R.
Private Sub EvaluateSecantResidual(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double, _
        ByVal eps0 As Double, ByVal kx As Double, ByVal ky As Double, _
        ByRef resN As Double, ByRef resMx As Double, ByRef resMy As Double)

    EvaluateStateInto section, concreteMaterial, steelMaterial, eps0, kx, ky, False, mNint, mMxint, mMyint
    resN = mNint - targetN
    resMx = mMxint - targetMx
    resMy = mMyint - targetMy
    SetResiduals targetN, targetMx, targetMy
End Sub

' Строит начальную численную матрицу чувствительности для Secant. Производные
' берутся центральными разностями по epsilon0, kappaX и kappaY, чтобы метод
' мог стартовать без касательной матрицы материала.
Private Function BuildInitialSecantMatrix(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double, _
        ByVal eps0 As Double, ByVal kx As Double, ByVal ky As Double, _
        ByVal baseRN As Double, ByVal baseRMx As Double, ByVal baseRMy As Double, _
        ByRef b() As Double) As Boolean

    Dim hEps As Double
    Dim hKx As Double
    Dim hKy As Double
    hEps = MaxDouble(0.0000001, Abs(eps0) * 0.0001)
    hKx = MaxDouble(0.000000001, Abs(kx) * 0.0001)
    hKy = MaxDouble(0.000000001, Abs(ky) * 0.0001)

    BuildSecantJacobianColumn section, concreteMaterial, steelMaterial, targetN, targetMx, targetMy, _
        eps0 + hEps, kx, ky, eps0 - hEps, kx, ky, 2# * hEps, b, 1
    BuildSecantJacobianColumn section, concreteMaterial, steelMaterial, targetN, targetMx, targetMy, _
        eps0, kx + hKx, ky, eps0, kx - hKx, ky, 2# * hKx, b, 2
    BuildSecantJacobianColumn section, concreteMaterial, steelMaterial, targetN, targetMx, targetMy, _
        eps0, kx, ky + hKy, eps0, kx, ky - hKy, 2# * hKy, b, 3
    mLastMatrixConditionEstimate = EstimateMatrixCondition(b)
    BuildInitialSecantMatrix = IsFiniteNumber(mLastMatrixConditionEstimate)
End Function

' Заполняет один столбец начальной Secant-матрицы центральной разностью.
' В колонку попадает изменение невязок N/Mx/My при малом изменении одного
' параметра плоскости деформаций.
Private Sub BuildSecantJacobianColumn(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double, _
        ByVal plusEps0 As Double, ByVal plusKx As Double, ByVal plusKy As Double, _
        ByVal minusEps0 As Double, ByVal minusKx As Double, ByVal minusKy As Double, _
        ByVal h As Double, _
        ByRef b() As Double, ByVal columnIndex As Long)

    Dim pN As Double
    Dim pMx As Double
    Dim pMy As Double
    Dim mN As Double
    Dim mMx As Double
    Dim mMy As Double
    Dim rN As Double
    Dim rMx As Double
    Dim rMy As Double
    Dim minusRN As Double
    Dim minusRMx As Double
    Dim minusRMy As Double
    EvaluateStateInto section, concreteMaterial, steelMaterial, plusEps0, plusKx, plusKy, False, pN, pMx, pMy
    EvaluateStateInto section, concreteMaterial, steelMaterial, minusEps0, minusKx, minusKy, False, mN, mMx, mMy
    rN = pN - targetN
    rMx = pMx - targetMx
    rMy = pMy - targetMy
    minusRN = mN - targetN
    minusRMx = mMx - targetMx
    minusRMy = mMy - targetMy
    b(1, columnIndex) = (rN - minusRN) / h
    b(2, columnIndex) = (rMx - minusRMx) / h
    b(3, columnIndex) = (rMy - minusRMy) / h
End Sub

' Решает линейную систему Secant для поправки плоскости деформаций.
' Возвращает False, если матрица вырождена или полученный шаг нечисловой.
Private Function SolveSecantCorrection(ByRef b() As Double, ByRef deltaEps As Double, _
        ByRef deltaKx As Double, ByRef deltaKy As Double) As Boolean

    Dim system As CLinearSystem3x3
    Set system = New CLinearSystem3x3
    If Not system.Solve(b(1, 1), b(1, 2), b(1, 3), _
            b(2, 1), b(2, 2), b(2, 3), _
            b(3, 1), b(3, 2), b(3, 3), _
            -mResidualN, -mResidualMx, -mResidualMy) Then
        SolveSecantCorrection = False
        Exit Function
    End If
    deltaEps = system.X1
    deltaKx = system.X2
    deltaKy = system.X3
    SolveSecantCorrection = IsFiniteNumber(deltaEps) And IsFiniteNumber(deltaKx) And IsFiniteNumber(deltaKy)
End Function

' Обновляет приближенную матрицу чувствительности по формуле Бройдена.
' Если шаг слишком мал или новая матрица выглядит нечисловой, вызывающий код
' должен пересобрать матрицу с нуля.
Private Function UpdateBroydenMatrix(ByRef b() As Double, ByVal s1 As Double, ByVal s2 As Double, _
        ByVal s3 As Double, ByVal y1 As Double, ByVal y2 As Double, ByVal y3 As Double) As Boolean

    Dim denom As Double
    denom = s1 * s1 + s2 * s2 + s3 * s3
    If denom <= mSecantMinStepNorm * mSecantMinStepNorm Then Exit Function

    Dim bs1 As Double
    Dim bs2 As Double
    Dim bs3 As Double
    bs1 = b(1, 1) * s1 + b(1, 2) * s2 + b(1, 3) * s3
    bs2 = b(2, 1) * s1 + b(2, 2) * s2 + b(2, 3) * s3
    bs3 = b(3, 1) * s1 + b(3, 2) * s2 + b(3, 3) * s3

    Dim u1 As Double
    Dim u2 As Double
    Dim u3 As Double
    u1 = y1 - bs1
    u2 = y2 - bs2
    u3 = y3 - bs3

    Dim i As Long
    Dim s(1 To 3) As Double
    Dim u(1 To 3) As Double
    s(1) = s1: s(2) = s2: s(3) = s3
    u(1) = u1: u(2) = u2: u(3) = u3
    For i = 1 To 3
        b(i, 1) = b(i, 1) + u(i) * s(1) / denom
        b(i, 2) = b(i, 2) + u(i) * s(2) / denom
        b(i, 3) = b(i, 3) + u(i) * s(3) / denom
    Next i

    mLastMatrixConditionEstimate = EstimateMatrixCondition(b)
    UpdateBroydenMatrix = IsFiniteNumber(mLastMatrixConditionEstimate)
End Function

' Пересобирает Secant-матрицу после застоя, плохого line search или неудачного
' обновления Бройдена. Количество перезапусков ограничено настройкой solver-а.
Private Function RestartSecantMatrix(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double, _
        ByVal eps0 As Double, ByVal kx As Double, ByVal ky As Double, _
        ByVal resN As Double, ByVal resMx As Double, ByVal resMy As Double, _
        ByRef b() As Double, ByVal reason As String) As Boolean

    If mMatrixRestartCount >= mSecantMaxRestarts Then
        mConverged = False
        mFailureCode = sfcNumericalFailure
        mStopReason = "NumericalFailure: превышен лимит перезапусков матрицы Secant; причина=" & reason
        Exit Function
    End If
    mMatrixRestartCount = mMatrixRestartCount + 1
    AppendDiagnostics mLoadStepsCompleted + 1, mIterations, 0#, 0#, targetN, targetMx, targetMy, "перезапуск матрицы Secant: " & reason
    RestartSecantMatrix = BuildInitialSecantMatrix(section, concreteMaterial, steelMaterial, targetN, targetMx, targetMy, _
        eps0, kx, ky, resN, resMx, resMy, b)
End Function

' Возвращает простую норму шага Secant по трем параметрам плоскости.
' Она нужна только для обнаружения численного застоя.
Private Function SecantStepNorm(ByVal deltaEps As Double, ByVal deltaKx As Double, ByVal deltaKy As Double) As Double
    SecantStepNorm = Abs(deltaEps) + Abs(deltaKx) + Abs(deltaKy)
End Function

' Дает грубую оценку разброса элементов матрицы чувствительности.
' Это диагностический индикатор вырожденности, а не строгий condition number.
Private Function EstimateMatrixCondition(ByRef b() As Double) As Double
    Dim maxAbs As Double
    Dim minAbs As Double
    Dim i As Long
    Dim j As Long
    minAbs = 1E+100
    For i = 1 To 3
        For j = 1 To 3
            If Abs(b(i, j)) > maxAbs Then maxAbs = Abs(b(i, j))
            If Abs(b(i, j)) > 0# And Abs(b(i, j)) < minAbs Then minAbs = Abs(b(i, j))
        Next j
    Next i
    If maxAbs <= 0# Or minAbs >= 1E+99 Then
        EstimateMatrixCondition = 1E+100
    Else
        EstimateMatrixCondition = maxAbs / minAbs
    End If
End Function

' Принимает или отклоняет найденную поправку Newton/Secant. При включенном
' line search метод дробит шаг, пока норма невязки не перестанет ухудшаться.
Private Function AcceptStep(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double, _
        ByVal deltaEps As Double, ByVal deltaKx As Double, ByVal deltaKy As Double) As Double

    Dim baseNorm As Double
    baseNorm = ResidualNorm(mResidualN, mResidualMx, mResidualMy, targetN, targetMx, targetMy)

    Dim alpha As Double
    alpha = mDampingInitial
    If alpha <= 0# Or alpha > 1# Then alpha = 1#
    mLineSearchIterations = 0
    mLineSearchUsed = False

    Do While alpha >= mMinLineSearchAlpha
        mLineSearchIterations = mLineSearchIterations + 1
        Dim trialEps As Double
        Dim trialKx As Double
        Dim trialKy As Double
        trialEps = mEpsilon0 + alpha * deltaEps
        trialKx = mKappaX + alpha * deltaKx
        trialKy = mKappaY + alpha * deltaKy

        Dim trialN As Double
        Dim trialMx As Double
        Dim trialMy As Double
        EvaluateStateInto section, concreteMaterial, steelMaterial, trialEps, trialKx, trialKy, False, trialN, trialMx, trialMy
        Dim trialNorm As Double
        trialNorm = ResidualNorm(trialN - targetN, trialMx - targetMx, trialMy - targetMy, targetN, targetMx, targetMy)

        If (Not mLineSearchEnabled) Or trialNorm <= baseNorm Or baseNorm <= 0.000000000001 Then
            If alpha < mDampingInitial Then mLineSearchUsed = True
            mEpsilon0 = trialEps
            mKappaX = trialKx
            mKappaY = trialKy
            AcceptStep = alpha
            Exit Function
        End If

        alpha = alpha / 2#
    Loop

    AcceptStep = 0#
End Function

' Ограничивает приращения epsilon0 и кривизн пользовательскими лимитами.
' Возвращает множитель, на который был уменьшен исходный расчетный шаг.
Private Function LimitIncrement(ByRef deltaEps As Double, ByRef deltaKx As Double, ByRef deltaKy As Double) As Double
    Dim scaleFactor As Double
    scaleFactor = 1#

    If mMaxDeltaEpsilon0 > 0# And Abs(deltaEps) > mMaxDeltaEpsilon0 Then
        scaleFactor = MinDouble(scaleFactor, mMaxDeltaEpsilon0 / Abs(deltaEps))
    End If
    If mMaxDeltaKappa > 0# And Abs(deltaKx) > mMaxDeltaKappa Then
        scaleFactor = MinDouble(scaleFactor, mMaxDeltaKappa / Abs(deltaKx))
    End If
    If mMaxDeltaKappa > 0# And Abs(deltaKy) > mMaxDeltaKappa Then
        scaleFactor = MinDouble(scaleFactor, mMaxDeltaKappa / Abs(deltaKy))
    End If

    deltaEps = deltaEps * scaleFactor
    deltaKx = deltaKx * scaleFactor
    deltaKy = deltaKy * scaleFactor
    LimitIncrement = scaleFactor
End Function

' Пересчитывает состояние и записывает внутренние усилия в поля solver-а.
' Если buildTangent=True, дополнительно собираются касательная матрица и
' экстремальные деформации/напряжения для Results.
Private Sub EvaluateState(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal eps0 As Double, ByVal kx As Double, ByVal ky As Double, ByVal buildTangent As Boolean)
    EvaluateStateInto section, concreteMaterial, steelMaterial, eps0, kx, ky, buildTangent, mNint, mMxint, mMyint
End Sub

' Интегрирует напряжения по бетонным и арматурным элементам для заданной
' плоскости деформаций. Это самая горячая операция solver-а; она работает
' через CSectionModel и не обращается к Excel.
Private Sub EvaluateStateInto(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal eps0 As Double, ByVal kx As Double, ByVal ky As Double, ByVal buildTangent As Boolean, _
        ByRef nInt As Double, ByRef mxInt As Double, ByRef myInt As Double)

    nInt = 0#: mxInt = 0#: myInt = 0#
    mInternalForceEvaluationCount = mInternalForceEvaluationCount + 1
    If buildTangent Then ClearTangent
    If buildTangent Then
        mMinConcreteStrain = 1E+100
        mMaxConcreteStrain = -1E+100
        mMaxConcreteCompressionStress = 0#
        mMinConcreteStress = 1E+100
        mMaxConcreteStress = -1E+100
        mHasConcreteCompression = False
        mHasConcreteTension = False
        mMinSteelStrain = 1E+100
        mMaxSteelStrain = -1E+100
        mMaxAbsSteelStrain = 0#
        mMinSteelStress = 1E+100
        mMaxSteelStress = -1E+100
        mHasSteelCompression = False
        mHasSteelTension = False
    End If

    Dim i As Long
    Dim strain As Double
    Dim stress As Double
    Dim tangent As Double
    Dim force As Double
    Dim area As Double
    Dim xCoord As Double
    Dim yCoord As Double
    Dim steelStress As Double
    Dim steelTangent As Double
    Dim concreteStress As Double
    Dim concreteTangent As Double
    Dim concreteDiagram As CMaterialDiagram ' Быстрый путь для штатной кусочно-линейной диаграммы.
    Dim steelDiagram As CMaterialDiagram ' Быстрый путь для штатной диаграммы арматуры.

    If TypeName(concreteMaterial) = "CMaterialDiagram" Then Set concreteDiagram = concreteMaterial
    If TypeName(steelMaterial) = "CMaterialDiagram" Then Set steelDiagram = steelMaterial

    For i = 1 To section.ConcreteCount
        section.ConcreteElementForSolver i, xCoord, yCoord, area
        strain = eps0 + kx * yCoord + ky * xCoord
        If buildTangent Then
            If strain < mMinConcreteStrain Then mMinConcreteStrain = strain
            If strain > mMaxConcreteStrain Then mMaxConcreteStrain = strain
        End If
        If concreteDiagram Is Nothing Then
            stress = concreteMaterial.GetStress(strain)
            If buildTangent Then tangent = concreteMaterial.GetTangentModulus(strain)
        ElseIf buildTangent Then
            concreteDiagram.EvaluateAtStrain strain, stress, tangent
        Else
            stress = concreteDiagram.StressAtStrain(strain)
        End If
        If buildTangent Then
            If stress < 0# Then
                mHasConcreteCompression = True
                If stress < mMinConcreteStress Then mMinConcreteStress = stress
            ElseIf stress > 0# Then
                mHasConcreteTension = True
                If stress > mMaxConcreteStress Then mMaxConcreteStress = stress
            End If
            If stress < 0# And Abs(stress) > mMaxConcreteCompressionStress Then
                mMaxConcreteCompressionStress = Abs(stress)
            End If
        End If
        force = stress * area
        nInt = nInt + force
        mxInt = mxInt + force * yCoord
        myInt = myInt + force * xCoord
        If buildTangent Then AddTangent tangent * area, xCoord, yCoord
    Next i

    If section.RebarCount > 0 Then
        For i = 1 To section.RebarCount
            section.RebarElementForSolver i, xCoord, yCoord, area
            strain = eps0 + kx * yCoord + ky * xCoord
            If buildTangent Then
                If strain < mMinSteelStrain Then mMinSteelStrain = strain
                If strain > mMaxSteelStrain Then mMaxSteelStrain = strain
                If Abs(strain) > mMaxAbsSteelStrain Then mMaxAbsSteelStrain = Abs(strain)
            End If
            If steelDiagram Is Nothing Then
                steelStress = steelMaterial.GetStress(strain)
                If buildTangent Then steelTangent = steelMaterial.GetTangentModulus(strain)
            ElseIf buildTangent Then
                steelDiagram.EvaluateAtStrain strain, steelStress, steelTangent
            Else
                steelStress = steelDiagram.StressAtStrain(strain)
            End If
            If buildTangent Then
                If steelStress < 0# Then
                    mHasSteelCompression = True
                    If steelStress < mMinSteelStress Then mMinSteelStress = steelStress
                ElseIf steelStress > 0# Then
                    mHasSteelTension = True
                    If steelStress > mMaxSteelStress Then mMaxSteelStress = steelStress
                End If
            End If
            If concreteDiagram Is Nothing Then
                concreteStress = concreteMaterial.GetStress(strain)
                If buildTangent Then concreteTangent = concreteMaterial.GetTangentModulus(strain)
            ElseIf buildTangent Then
                concreteDiagram.EvaluateAtStrain strain, concreteStress, concreteTangent
            Else
                concreteStress = concreteDiagram.StressAtStrain(strain)
            End If
            stress = steelStress - concreteStress
            force = stress * area
            nInt = nInt + force
            mxInt = mxInt + force * yCoord
            myInt = myInt + force * xCoord
            If buildTangent Then
                tangent = steelTangent - concreteTangent
                AddTangent tangent * area, xCoord, yCoord
            End If
        Next i
    End If
    If buildTangent Then
        ' У отсутствующего знака напряжения нет экстремума. Служебные
        ' границы поиска не должны попадать в физический snapshot.
        If Not mHasConcreteCompression Then mMinConcreteStress = 0#
        If Not mHasConcreteTension Then mMaxConcreteStress = 0#
        If Not mHasSteelCompression Then mMinSteelStress = 0#
        If Not mHasSteelTension Then mMaxSteelStress = 0#
    End If
End Sub

' Очищает касательную матрицу перед новой сборкой производных равновесия.
Private Sub ClearTangent()
    Dim i As Long
    For i = 1 To 3
        mK(i, 1) = 0#: mK(i, 2) = 0#: mK(i, 3) = 0#
    Next i
End Sub

' Добавляет вклад одного элемента в касательную матрицу равновесия.
' axialStiffness уже включает E_tangent * Area, поэтому здесь остается только
' умножение на координаты по правилу моментов.
Private Sub AddTangent(ByVal axialStiffness As Double, ByVal xCoord As Double, ByVal yCoord As Double)
    mK(1, 1) = mK(1, 1) + axialStiffness
    mK(1, 2) = mK(1, 2) + axialStiffness * yCoord
    mK(1, 3) = mK(1, 3) + axialStiffness * xCoord
    mK(2, 1) = mK(2, 1) + axialStiffness * yCoord
    mK(2, 2) = mK(2, 2) + axialStiffness * yCoord * yCoord
    mK(2, 3) = mK(2, 3) + axialStiffness * yCoord * xCoord
    mK(3, 1) = mK(3, 1) + axialStiffness * xCoord
    mK(3, 2) = mK(3, 2) + axialStiffness * xCoord * yCoord
    mK(3, 3) = mK(3, 3) + axialStiffness * xCoord * xCoord
End Sub

' Обновляет абсолютные и относительные невязки равновесия после очередной
' оценки внутренних усилий.
Private Sub SetResiduals(ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double)
    mResidualN = mNint - targetN
    mResidualMx = mMxint - targetMx
    mResidualMy = mMyint - targetMy
    mRelativeResidualN = RelativeResidual(mResidualN, targetN)
    mRelativeResidualMx = RelativeResidual(mResidualMx, targetMx)
    mRelativeResidualMy = RelativeResidual(mResidualMy, targetMy)
End Sub

' Проверяет выполнение трех независимых условий равновесия по абсолютным
' допускам solver-а.
Private Function IsConverged(ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double) As Boolean
    IsConverged = Abs(mResidualN) <= mToleranceN And _
        Abs(mResidualMx) <= mToleranceMx And _
        Abs(mResidualMy) <= mToleranceMy
End Function

' Собирает безразмерную норму невязки для line search и диагностики.
' Нормировка идет по максимуму из заданного solver tolerance и малой доли
' целевого усилия, чтобы крупные нагрузки не доминировали численно.
Private Function ResidualNorm(ByVal resN As Double, ByVal resMx As Double, ByVal resMy As Double, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double) As Double
    ResidualNorm = Abs(resN) / MaxDouble(mToleranceN, Abs(targetN) * 0.000000001) + _
        Abs(resMx) / MaxDouble(mToleranceMx, Abs(targetMx) * 0.000000001) + _
        Abs(resMy) / MaxDouble(mToleranceMy, Abs(targetMy) * 0.000000001)
End Function

' Возвращает относительную невязку, а для нулевого target - абсолютную.
Private Function RelativeResidual(ByVal residual As Double, ByVal target As Double) As Double
    If Abs(target) > 0.000000000001 Then
        RelativeResidual = Abs(residual / target)
    Else
        RelativeResidual = Abs(residual)
    End If
End Function

' Записывает одну строку подробной диагностики итераций solver-а.
' Диагностика включается только для отчета/тестов и не влияет на расчет.
Private Sub AppendDiagnostics(ByVal stepIndex As Long, ByVal iter As Long, ByVal alpha As Double, _
        ByVal deltaScale As Double, ByVal targetN As Double, ByVal targetMx As Double, _
        ByVal targetMy As Double, ByVal state As String)
    If Not mDiagnosticsEnabled Then Exit Sub
    mLastResidualNorm = ResidualNorm(mResidualN, mResidualMx, mResidualMy, targetN, targetMx, targetMy)
    mDiagnosticLog = mDiagnosticLog & "step=" & CStr(stepIndex) & _
        "; iter=" & CStr(iter) & _
        "; method=" & mSolverMethod & _
        "; state=" & state & _
        "; eps0=" & FormatNumberInvariant(mEpsilon0) & _
        "; kappaX=" & FormatNumberInvariant(mKappaX) & _
        "; kappaY=" & FormatNumberInvariant(mKappaY) & _
        "; RN=" & FormatNumberInvariant(mResidualN) & _
        "; RMx=" & FormatNumberInvariant(mResidualMx) & _
        "; RMy=" & FormatNumberInvariant(mResidualMy) & _
        "; residualNorm=" & FormatNumberInvariant(mLastResidualNorm) & _
        "; stepNorm=" & FormatNumberInvariant(mLastStepNorm) & _
        "; matrixConditionEstimate=" & FormatNumberInvariant(mLastMatrixConditionEstimate) & _
        "; matrixRestartCount=" & CStr(mMatrixRestartCount) & _
        "; lineSearchUsed=" & CStr(mLineSearchUsed) & _
        "; lineSearchIterations=" & CStr(mLineSearchIterations) & _
        "; internalForceEvaluationCount=" & CStr(mInternalForceEvaluationCount) & _
        "; internalNewtonCallCount=" & CStr(mInternalNewtonCallCount) & _
        "; alpha=" & FormatNumberInvariant(alpha) & _
        "; deltaScale=" & FormatNumberInvariant(deltaScale) & vbCrLf
End Sub

' Проверяет модель и настройки до численных итераций. Программная зависимость,
' входная геометрия и ошибочная конфигурация имеют разные машинные причины.
Private Function ValidateInputs(ByVal section As CSectionModel, ByVal concreteMaterial As Object, _
        ByVal steelMaterial As Object) As Boolean
    If section Is Nothing Then
        SetValidationFailure sfcInternalError, "Модель сечения не передана."
    ElseIf concreteMaterial Is Nothing Then
        SetValidationFailure sfcInternalError, "Материал бетона не передан."
    ElseIf steelMaterial Is Nothing Then
        SetValidationFailure sfcInternalError, "Материал арматуры не передан."
    ElseIf section.ConcreteCount <= 0 Then
        SetValidationFailure sfcInvalidInput, "В модели сечения нет бетонных элементов."
    ElseIf mMaxIterations <= 0 Then
        SetValidationFailure sfcInvalidConfiguration, "Solver.MaxIterations должен быть положительным."
    ElseIf mLoadSteps <= 0 Then
        SetValidationFailure sfcInvalidConfiguration, "Solver.LoadSteps должен быть положительным."
    ElseIf mSecantMaxRestarts < 0 Then
        SetValidationFailure sfcInvalidConfiguration, "Solver.SecantMaxRestarts не может быть отрицательным."
    ElseIf mSecantMinStepNorm <= 0# Then
        SetValidationFailure sfcInvalidConfiguration, "Solver.SecantMinStepNorm должен быть положительным."
    ElseIf mToleranceN <= 0# Or mToleranceMx <= 0# Or mToleranceMy <= 0# Then
        SetValidationFailure sfcInvalidConfiguration, "Допуски Solver.ToleranceN/Mx/My должны быть положительными."
    Else
        ValidateInputs = True
    End If
End Function

' Сохраняет терминальную причину проверки контракта без фиктивного solve.
Private Sub SetValidationFailure(ByVal failureCode As ESolverFailureCode, ByVal reason As String)
    mConverged = False
    mFailureCode = failureCode
    mStopReason = reason
End Sub

' Нормализует выбранный численный метод и формирует InputError при неизвестном
' значении Solver.Method.
Private Function ValidateSolverMethod() As Boolean
    If StrComp(mSolverMethod, "Newton", vbTextCompare) = 0 Then
        mSolverMethod = "Newton"
        ValidateSolverMethod = True
    ElseIf StrComp(mSolverMethod, "Secant", vbTextCompare) = 0 Then
        mSolverMethod = "Secant"
        ValidateSolverMethod = True
    Else
        mConverged = False
        mFailureCode = sfcInvalidConfiguration
        mStopReason = "InputError: Solver.Method должен быть Newton или Secant."
    End If
End Function

' Очищает результаты последнего solve, но не сбрасывает настройки solver-а и
' заданное начальное приближение.
Public Sub ClearResult()
    mConverged = False
    mStopReason = vbNullString
    mFailureCode = sfcNone
    mEpsilon0 = 0#: mKappaX = 0#: mKappaY = 0#
    mNint = 0#: mMxint = 0#: mMyint = 0#
    mResidualN = 0#: mResidualMx = 0#: mResidualMy = 0#
    mRelativeResidualN = 0#: mRelativeResidualMx = 0#: mRelativeResidualMy = 0#
    mMinConcreteStrain = 0#: mMaxConcreteStrain = 0#
    mMaxConcreteCompressionStress = 0#
    mMinConcreteStress = 0#: mMaxConcreteStress = 0#
    mHasConcreteCompression = False: mHasConcreteTension = False
    mMinSteelStrain = 0#: mMaxSteelStrain = 0#: mMaxAbsSteelStrain = 0#
    mMinSteelStress = 0#: mMaxSteelStress = 0#
    mHasSteelCompression = False: mHasSteelTension = False
    mIterations = 0
    mLoadStepsCompleted = 0
    mDiagnosticLog = vbNullString
    mLastAlpha = 0#
    mLastDeltaScale = 0#
    mLastResidualNorm = 0#
    mLastStepNorm = 0#
    mLastMatrixConditionEstimate = 0#
    mMatrixRestartCount = 0
    mLineSearchUsed = False
    mLineSearchIterations = 0
    mInternalForceEvaluationCount = 0
    mInternalNewtonCallCount = 0
    ClearTangent
End Sub

' Задает стартовую плоскость деформаций для следующего solve. Используется
' warm-start сценариями: трещины, несущая способность и continuation-поиски.
Public Sub SetInitialState(ByVal epsilon0 As Double, ByVal kappaX As Double, ByVal kappaY As Double)
    mUseInitialState = True
    mInitialEpsilon0 = epsilon0
    mInitialKappaX = kappaX
    mInitialKappaY = kappaY
End Sub

' Задает стартовую плоскость вместе с нагрузкой, относительно которой эта
' плоскость уже была уравновешена. Это позволяет прикладывать следующую
' нагрузку ступенями от известной точки, а не от нуля.
Public Sub SetInitialLoadState(ByVal epsilon0 As Double, ByVal kappaX As Double, ByVal kappaY As Double, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double)
    SetInitialState epsilon0, kappaX, kappaY
    mUseInitialLoads = True
    mInitialTargetN = targetN
    mInitialTargetMx = targetMx
    mInitialTargetMy = targetMy
End Sub

' Сбрасывает warm-start перед независимым расчетом нового состояния.
Public Sub ClearInitialState()
    mUseInitialState = False
    mInitialEpsilon0 = 0#: mInitialKappaX = 0#: mInitialKappaY = 0#
    mUseInitialLoads = False
    mInitialTargetN = 0#: mInitialTargetMx = 0#: mInitialTargetMy = 0#
End Sub

' Утилитарный минимум без WorksheetFunction в горячем коде solver-а.
Private Function MinDouble(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then MinDouble = a Else MinDouble = b
End Function

' Утилитарный максимум без WorksheetFunction в горячем коде solver-а.
Private Function MaxDouble(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then MaxDouble = a Else MaxDouble = b
End Function

' Форматирует число для диагностического журнала с точкой как разделителем,
' чтобы отчет был одинаковым при разных региональных настройках Excel.
Private Function FormatNumberInvariant(ByVal value As Double) As String
    FormatNumberInvariant = Replace$(Format$(value, "0.############"), ",", ".")
End Function

' Проверяет, что число конечно и не является NaN/overflow-маркером.
Private Function IsFiniteNumber(ByVal value As Double) As Boolean
    IsFiniteNumber = (value = value And Abs(value) < 1E+100)
End Function

Public Property Get Converged() As Boolean
    Converged = mConverged
End Property

Public Property Get StopReason() As String
    StopReason = mStopReason
End Property

' Возвращает машинную причину отказа последнего solve. Расчетные слои
' используют ее вместо анализа текста StopReason.
Public Property Get FailureCode() As ESolverFailureCode
    FailureCode = mFailureCode
End Property

Public Property Get Epsilon0() As Double
    Epsilon0 = mEpsilon0
End Property

Public Property Get KappaX() As Double
    KappaX = mKappaX
End Property

Public Property Get KappaY() As Double
    KappaY = mKappaY
End Property

Public Property Get Nint() As Double
    Nint = mNint
End Property

Public Property Get Mxint() As Double
    Mxint = mMxint
End Property

Public Property Get Myint() As Double
    Myint = mMyint
End Property

Public Property Get ResidualN() As Double
    ResidualN = mResidualN
End Property

Public Property Get ResidualMx() As Double
    ResidualMx = mResidualMx
End Property

Public Property Get ResidualMy() As Double
    ResidualMy = mResidualMy
End Property

Public Property Get RelativeResidualN() As Double
    RelativeResidualN = mRelativeResidualN
End Property

Public Property Get RelativeResidualMx() As Double
    RelativeResidualMx = mRelativeResidualMx
End Property

Public Property Get RelativeResidualMy() As Double
    RelativeResidualMy = mRelativeResidualMy
End Property

Public Property Get MinConcreteStrain() As Double
    MinConcreteStrain = mMinConcreteStrain
End Property

Public Property Get MaxConcreteStrain() As Double
    MaxConcreteStrain = mMaxConcreteStrain
End Property

Public Property Get MaxConcreteCompressionStress() As Double
    MaxConcreteCompressionStress = mMaxConcreteCompressionStress
End Property

Public Property Get MinConcreteStress() As Double
    MinConcreteStress = mMinConcreteStress
End Property

Public Property Get MaxConcreteStress() As Double
    MaxConcreteStress = mMaxConcreteStress
End Property

Public Property Get HasConcreteCompression() As Boolean
    HasConcreteCompression = mHasConcreteCompression
End Property

Public Property Get HasConcreteTension() As Boolean
    HasConcreteTension = mHasConcreteTension
End Property

Public Property Get MinSteelStrain() As Double
    MinSteelStrain = mMinSteelStrain
End Property

Public Property Get MaxSteelStrain() As Double
    MaxSteelStrain = mMaxSteelStrain
End Property

Public Property Get MaxAbsSteelStrain() As Double
    MaxAbsSteelStrain = mMaxAbsSteelStrain
End Property

Public Property Get MinSteelStress() As Double
    MinSteelStress = mMinSteelStress
End Property

Public Property Get MaxSteelStress() As Double
    MaxSteelStress = mMaxSteelStress
End Property

Public Property Get HasSteelCompression() As Boolean
    HasSteelCompression = mHasSteelCompression
End Property

Public Property Get HasSteelTension() As Boolean
    HasSteelTension = mHasSteelTension
End Property

Public Property Get Iterations() As Long
    Iterations = mIterations
End Property

Public Property Get Tangent(ByVal rowIndex As Long, ByVal columnIndex As Long) As Double
    Tangent = mK(rowIndex, columnIndex)
End Property

Public Property Get DiagnosticLog() As String
    DiagnosticLog = mDiagnosticLog
End Property

Public Property Get SolverMethod() As String
    SolverMethod = mSolverMethod
End Property

Public Property Get LastResidualNorm() As Double
    LastResidualNorm = mLastResidualNorm
End Property

Public Property Get LastStepNorm() As Double
    LastStepNorm = mLastStepNorm
End Property

Public Property Get LastMatrixConditionEstimate() As Double
    LastMatrixConditionEstimate = mLastMatrixConditionEstimate
End Property

Public Property Let MaxIterations(ByVal value As Long)
    mMaxIterations = value
End Property

Public Property Let LoadSteps(ByVal value As Long)
    mLoadSteps = value
End Property

Public Property Let ToleranceN(ByVal value As Double)
    mToleranceN = value
End Property

Public Property Let ToleranceMx(ByVal value As Double)
    mToleranceMx = value
End Property

Public Property Let ToleranceMy(ByVal value As Double)
    mToleranceMy = value
End Property

Public Property Let LineSearchEnabled(ByVal value As Boolean)
    mLineSearchEnabled = value
End Property

Public Property Let DampingInitial(ByVal value As Double)
    mDampingInitial = value
End Property

Public Property Let MinLineSearchAlpha(ByVal value As Double)
    mMinLineSearchAlpha = value
End Property

Public Property Let SolverMethod(ByVal value As String)
    mSolverMethod = value
End Property

Public Property Let SecantMaxRestarts(ByVal value As Long)
    mSecantMaxRestarts = value
End Property

Public Property Let SecantMinStepNorm(ByVal value As Double)
    mSecantMinStepNorm = value
End Property

Public Property Let MaxDeltaEpsilon0(ByVal value As Double)
    mMaxDeltaEpsilon0 = value
End Property

Public Property Let MaxDeltaKappa(ByVal value As Double)
    mMaxDeltaKappa = value
End Property

' ==========================================================================
' ДЛЯ ТЕСТОВ
' ==========================================================================
' Эти свойства и настройки раскрывают внутренний ход Newton/Secant. Они нужны
' unit- и regression-тестам, чтобы отличать ветки алгоритма, считать пробные
' пересчеты усилий и проверять рестарты матрицы. Пользовательский расчет
' использует сходимость, найденные деформации, внутренние усилия и DiagnosticLog.

Public Property Get LoadStepsCompleted() As Long
    LoadStepsCompleted = mLoadStepsCompleted
End Property

Public Property Get InternalForceEvaluationCount() As Long
    InternalForceEvaluationCount = mInternalForceEvaluationCount
End Property

Public Property Get InternalNewtonCallCount() As Long
    InternalNewtonCallCount = mInternalNewtonCallCount
End Property

Public Property Get MatrixRestartCount() As Long
    MatrixRestartCount = mMatrixRestartCount
End Property

Public Property Let DiagnosticsEnabled(ByVal value As Boolean)
    mDiagnosticsEnabled = value
End Property














