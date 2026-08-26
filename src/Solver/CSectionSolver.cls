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
Private mMinSteelStrain As Double ' Наиболее сжатая деформация арматуры в последнем посчитанном состоянии.
Private mMaxSteelStrain As Double ' Наиболее растянутая деформация арматуры в последнем посчитанном состоянии.
Private mMaxAbsSteelStrain As Double ' Максимальная по модулю деформация арматуры в последнем состоянии.
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
    mMaxIterations = 40
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

    ClearResult
    ValidateInputs section, concreteMaterial, steelMaterial
    If Not ValidateSolverMethod() Then Exit Sub
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
        mStopReason = "Converged by " & mSolverMethod & " with load stepping."
    Else
        mStopReason = "Final residual check failed."
    End If
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
    ValidateInputs section, concreteMaterial, steelMaterial
    mEpsilon0 = epsilon0
    mKappaX = kappaX
    mKappaY = kappaY
    EvaluateState section, concreteMaterial, steelMaterial, mEpsilon0, mKappaX, mKappaY, True
    SetResiduals mNint, mMxint, mMyint
    mConverged = True
    mStopReason = "Evaluated at prescribed strain plane."
End Sub

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
        mStopReason = "InputError: Solver.Method must be Newton or Secant."
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
            AppendDiagnostics stepIndex, iter, 0#, 0#, targetN, targetMx, targetMy, "converged"
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
            mStopReason = "Tangent matrix solve failed: " & system.StopReason
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
        AppendDiagnostics stepIndex, iter, mLastAlpha, mLastDeltaScale, targetN, targetMx, targetMy, "iteration"
        If mLastAlpha <= 0# Then
            mConverged = False
            mStopReason = "Line search failed to improve residual."
            SolveNewtonLoadStep = False
            Exit Function
        End If
    Next iter

    RefreshCurrentState section, concreteMaterial, steelMaterial, targetN, targetMx, targetMy
    mConverged = False
    mStopReason = "Maximum iterations reached at load step " & CStr(stepIndex) & "."
    SolveNewtonLoadStep = False
End Function

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
            AppendDiagnostics stepIndex, iter, 0#, 0#, targetN, targetMx, targetMy, "secant converged"
            SolveSecantLoadStep = True
            Exit Function
        End If

        Dim deltaEps As Double
        Dim deltaKx As Double
        Dim deltaKy As Double
        If Not SolveSecantCorrection(b, deltaEps, deltaKx, deltaKy) Then
            If Not RestartSecantMatrix(section, concreteMaterial, steelMaterial, targetN, targetMx, targetMy, _
                    mEpsilon0, mKappaX, mKappaY, mResidualN, mResidualMx, mResidualMy, b, "linear solve failed") Then
                SolveSecantLoadStep = False
                Exit Function
            End If
            If Not SolveSecantCorrection(b, deltaEps, deltaKx, deltaKy) Then
                mConverged = False
                mStopReason = "NumericalFailure: Secant matrix solve failed after restart."
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
                    mEpsilon0, mKappaX, mKappaY, mResidualN, mResidualMx, mResidualMy, b, "step is too small") Then
                SolveSecantLoadStep = False
                Exit Function
            End If
            AppendDiagnostics stepIndex, iter, 0#, mLastDeltaScale, targetN, targetMx, targetMy, "secant restarted"
        Else
            mLastAlpha = AcceptStep(section, concreteMaterial, steelMaterial, targetN, targetMx, targetMy, deltaEps, deltaKx, deltaKy)
            If mLastAlpha <= 0# Then
                If Not RestartSecantMatrix(section, concreteMaterial, steelMaterial, targetN, targetMx, targetMy, _
                        oldEps, oldKx, oldKy, oldRN, oldRMx, oldRMy, b, "line search failed") Then
                    SolveSecantLoadStep = False
                    Exit Function
                End If
            Else
                EvaluateSecantResidual section, concreteMaterial, steelMaterial, targetN, targetMx, targetMy, _
                    mEpsilon0, mKappaX, mKappaY, mResidualN, mResidualMx, mResidualMy
                If Not UpdateBroydenMatrix(b, mEpsilon0 - oldEps, mKappaX - oldKx, mKappaY - oldKy, _
                        mResidualN - oldRN, mResidualMx - oldRMx, mResidualMy - oldRMy) Then
                    If Not RestartSecantMatrix(section, concreteMaterial, steelMaterial, targetN, targetMx, targetMy, _
                            mEpsilon0, mKappaX, mKappaY, mResidualN, mResidualMx, mResidualMy, b, "Broyden update failed") Then
                        SolveSecantLoadStep = False
                        Exit Function
                    End If
                End If
                AppendDiagnostics stepIndex, iter, mLastAlpha, mLastDeltaScale, targetN, targetMx, targetMy, "secant iteration"
            End If
        End If
    Next iter

    RefreshCurrentState section, concreteMaterial, steelMaterial, targetN, targetMx, targetMy
    mConverged = False
    mStopReason = "NumericalFailure: Secant maximum iterations reached at load step " & CStr(stepIndex) & "."
    SolveSecantLoadStep = False
End Function

Private Sub RefreshCurrentState(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double)
    EvaluateState section, concreteMaterial, steelMaterial, mEpsilon0, mKappaX, mKappaY, True
    SetResiduals targetN, targetMx, targetMy
End Sub

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

' Создает расчетный или интерфейсный объект из нормализованных исходных данных и локальных настроек.
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

' Создает расчетный или интерфейсный объект из нормализованных исходных данных и локальных настроек.
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

Private Function RestartSecantMatrix(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double, _
        ByVal eps0 As Double, ByVal kx As Double, ByVal ky As Double, _
        ByVal resN As Double, ByVal resMx As Double, ByVal resMy As Double, _
        ByRef b() As Double, ByVal reason As String) As Boolean

    If mMatrixRestartCount >= mSecantMaxRestarts Then
        mConverged = False
        mStopReason = "NumericalFailure: Secant matrix restart limit exceeded; reason=" & reason
        Exit Function
    End If
    mMatrixRestartCount = mMatrixRestartCount + 1
    AppendDiagnostics mLoadStepsCompleted + 1, mIterations, 0#, 0#, targetN, targetMx, targetMy, "secant matrix restart: " & reason
    RestartSecantMatrix = BuildInitialSecantMatrix(section, concreteMaterial, steelMaterial, targetN, targetMx, targetMy, _
        eps0, kx, ky, resN, resMx, resMy, b)
End Function

Private Function SecantStepNorm(ByVal deltaEps As Double, ByVal deltaKx As Double, ByVal deltaKy As Double) As Double
    SecantStepNorm = Abs(deltaEps) + Abs(deltaKx) + Abs(deltaKy)
End Function

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

Private Sub EvaluateState(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal eps0 As Double, ByVal kx As Double, ByVal ky As Double, ByVal buildTangent As Boolean)
    EvaluateStateInto section, concreteMaterial, steelMaterial, eps0, kx, ky, buildTangent, mNint, mMxint, mMyint
End Sub

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
        mMinSteelStrain = 1E+100
        mMaxSteelStrain = -1E+100
        mMaxAbsSteelStrain = 0#
    End If

    Dim i As Long
    Dim strain As Double
    Dim stress As Double
    Dim tangent As Double
    Dim force As Double
    Dim area As Double

    For i = 1 To section.ConcreteCount
        strain = eps0 + kx * section.ConcreteY(i) + ky * section.ConcreteX(i)
        If buildTangent Then
            If strain < mMinConcreteStrain Then mMinConcreteStrain = strain
            If strain > mMaxConcreteStrain Then mMaxConcreteStrain = strain
        End If
        stress = concreteMaterial.GetStress(strain)
        tangent = concreteMaterial.GetTangentModulus(strain)
        area = section.ConcreteArea(i)
        force = stress * area
        nInt = nInt + force
        mxInt = mxInt + force * section.ConcreteY(i)
        myInt = myInt + force * section.ConcreteX(i)
        If buildTangent Then AddTangent tangent * area, section.ConcreteX(i), section.ConcreteY(i)
    Next i

    If section.RebarCount > 0 Then
        For i = 1 To section.RebarCount
            strain = eps0 + kx * section.RebarY(i) + ky * section.RebarX(i)
            If buildTangent Then
                If strain < mMinSteelStrain Then mMinSteelStrain = strain
                If strain > mMaxSteelStrain Then mMaxSteelStrain = strain
                If Abs(strain) > mMaxAbsSteelStrain Then mMaxAbsSteelStrain = Abs(strain)
            End If
            Dim steelStress As Double
            Dim steelTangent As Double
            Dim concreteStress As Double
            Dim concreteTangent As Double
            steelStress = steelMaterial.GetStress(strain)
            steelTangent = steelMaterial.GetTangentModulus(strain)
            concreteStress = concreteMaterial.GetStress(strain)
            concreteTangent = concreteMaterial.GetTangentModulus(strain)
            stress = steelStress - concreteStress
            tangent = steelTangent - concreteTangent
            force = stress * section.RebarArea(i)
            nInt = nInt + force
            mxInt = mxInt + force * section.RebarY(i)
            myInt = myInt + force * section.RebarX(i)
            If buildTangent Then AddTangent tangent * section.RebarArea(i), section.RebarX(i), section.RebarY(i)
        Next i
    End If
End Sub

' Очищает накопленное состояние перед новым расчетом или повторным формированием вывода.
Private Sub ClearTangent()
    Dim i As Long
    For i = 1 To 3
        mK(i, 1) = 0#: mK(i, 2) = 0#: mK(i, 3) = 0#
    Next i
End Sub

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

Private Sub SetResiduals(ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double)
    mResidualN = mNint - targetN
    mResidualMx = mMxint - targetMx
    mResidualMy = mMyint - targetMy
    mRelativeResidualN = RelativeResidual(mResidualN, targetN)
    mRelativeResidualMx = RelativeResidual(mResidualMx, targetMx)
    mRelativeResidualMy = RelativeResidual(mResidualMy, targetMy)
End Sub

Private Function IsConverged(ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double) As Boolean
    IsConverged = Abs(mResidualN) <= mToleranceN And _
        Abs(mResidualMx) <= mToleranceMx And _
        Abs(mResidualMy) <= mToleranceMy
End Function

Private Function ResidualNorm(ByVal resN As Double, ByVal resMx As Double, ByVal resMy As Double, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double) As Double
    ResidualNorm = Abs(resN) / MaxDouble(mToleranceN, Abs(targetN) * 0.000000001) + _
        Abs(resMx) / MaxDouble(mToleranceMx, Abs(targetMx) * 0.000000001) + _
        Abs(resMy) / MaxDouble(mToleranceMy, Abs(targetMy) * 0.000000001)
End Function

Private Function RelativeResidual(ByVal residual As Double, ByVal target As Double) As Double
    If Abs(target) > 0.000000000001 Then
        RelativeResidual = Abs(residual / target)
    Else
        RelativeResidual = Abs(residual)
    End If
End Function

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

' Проверяет входные данные и прерывает выполнение понятной ошибкой, если расчетный контракт нарушен.
Private Sub ValidateInputs(ByVal section As CSectionModel, ByVal concreteMaterial As Object, ByVal steelMaterial As Object)
    If section Is Nothing Then Err.Raise vbObjectError + 3400, "CSectionSolver", "Модель сечения не передана."
    If section.ConcreteCount <= 0 Then Err.Raise vbObjectError + 3401, "CSectionSolver", "В модели сечения нет бетонных элементов."
    If concreteMaterial Is Nothing Then Err.Raise vbObjectError + 3402, "CSectionSolver", "Материал бетона не передан."
    If steelMaterial Is Nothing Then Err.Raise vbObjectError + 3403, "CSectionSolver", "Материал арматуры не передан."
    If mMaxIterations <= 0 Then Err.Raise vbObjectError + 3404, "CSectionSolver", "Solver.MaxIterations должен быть положительным."
    If mLoadSteps <= 0 Then Err.Raise vbObjectError + 3405, "CSectionSolver", "Solver.LoadSteps должен быть положительным."
    If mSecantMaxRestarts < 0 Then Err.Raise vbObjectError + 3406, "CSectionSolver", "Solver.SecantMaxRestarts не может быть отрицательным."
    If mSecantMinStepNorm <= 0# Then Err.Raise vbObjectError + 3407, "CSectionSolver", "Solver.SecantMinStepNorm должен быть положительным."
End Sub

' Проверяет входные данные и прерывает выполнение понятной ошибкой, если расчетный контракт нарушен.
Private Function ValidateSolverMethod() As Boolean
    If StrComp(mSolverMethod, "Newton", vbTextCompare) = 0 Then
        mSolverMethod = "Newton"
        ValidateSolverMethod = True
    ElseIf StrComp(mSolverMethod, "Secant", vbTextCompare) = 0 Then
        mSolverMethod = "Secant"
        ValidateSolverMethod = True
    Else
        mConverged = False
        mStopReason = "InputError: Solver.Method must be Newton or Secant."
    End If
End Function

' Очищает накопленное состояние перед новым расчетом или повторным формированием вывода.
Public Sub ClearResult()
    mConverged = False
    mStopReason = vbNullString
    mEpsilon0 = 0#: mKappaX = 0#: mKappaY = 0#
    mNint = 0#: mMxint = 0#: mMyint = 0#
    mResidualN = 0#: mResidualMx = 0#: mResidualMy = 0#
    mRelativeResidualN = 0#: mRelativeResidualMx = 0#: mRelativeResidualMy = 0#
    mMinConcreteStrain = 0#: mMaxConcreteStrain = 0#
    mMinSteelStrain = 0#: mMaxSteelStrain = 0#: mMaxAbsSteelStrain = 0#
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

Public Sub SetInitialState(ByVal epsilon0 As Double, ByVal kappaX As Double, ByVal kappaY As Double)
    mUseInitialState = True
    mInitialEpsilon0 = epsilon0
    mInitialKappaX = kappaX
    mInitialKappaY = kappaY
End Sub

Public Sub SetInitialLoadState(ByVal epsilon0 As Double, ByVal kappaX As Double, ByVal kappaY As Double, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double)
    SetInitialState epsilon0, kappaX, kappaY
    mUseInitialLoads = True
    mInitialTargetN = targetN
    mInitialTargetMx = targetMx
    mInitialTargetMy = targetMy
End Sub

' Очищает накопленное состояние перед новым расчетом или повторным формированием вывода.
Public Sub ClearInitialState()
    mUseInitialState = False
    mInitialEpsilon0 = 0#: mInitialKappaX = 0#: mInitialKappaY = 0#
    mUseInitialLoads = False
    mInitialTargetN = 0#: mInitialTargetMx = 0#: mInitialTargetMy = 0#
End Sub

Private Function MinDouble(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then MinDouble = a Else MinDouble = b
End Function

Private Function MaxDouble(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then MaxDouble = a Else MaxDouble = b
End Function

Private Function FormatNumberInvariant(ByVal value As Double) As String
    FormatNumberInvariant = Replace$(Format$(value, "0.############"), ",", ".")
End Function

Private Function IsFiniteNumber(ByVal value As Double) As Boolean
    IsFiniteNumber = (value = value And Abs(value) < 1E+100)
End Function

Public Property Get Converged() As Boolean
    Converged = mConverged
End Property

Public Property Get StopReason() As String
    StopReason = mStopReason
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

Public Property Get MinSteelStrain() As Double
    MinSteelStrain = mMinSteelStrain
End Property

Public Property Get MaxSteelStrain() As Double
    MaxSteelStrain = mMaxSteelStrain
End Property

Public Property Get MaxAbsSteelStrain() As Double
    MaxAbsSteelStrain = mMaxAbsSteelStrain
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














