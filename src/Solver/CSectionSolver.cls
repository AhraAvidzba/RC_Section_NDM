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
' Равновесие сечения при N + Mx + My
' ==========================================================================
' Подбирает epsilon0, kappaX и kappaY линейного поля деформаций. По ним
' получает напряжения бетона и арматуры и внутренние Nint, Mxint, Myint.
' Решение подтверждается отдельными допусками равновесия всех трёх усилий.
' Материалы, нагрузки и единицы заранее подготовлены вызывающим расчётом;
' чтение Excel и поиск предельной нагрузки выполняются отдельно.

Private Const STRAIN_EXTREMA_ROUNDING_MARGIN As Double = 1.77635683940025E-15            ' Допуск в восемь машинных погрешностей Double для близких экстремумов деформаций; при неоднозначном сравнении используется полный обход.

' Рабочие результаты и настройки решения равновесия.

Private mConverged As Boolean ' True, если равновесие достигнуто в заданных допусках по N/Mx/My.
Private mStopReason As String ' Русское пояснение итога; машинная причина хранится отдельно в mFailureCode.
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
Private mFrame As CSectionNumericalFrame ' Общий центр и масштаб без изменения публичных координат.
Private mCenterEpsilon As Double, mCenterMx As Double, mCenterMy As Double ' Деформация в численном центре и внутренние моменты относительно него, Н*мм.
Private mCenterResidualMx As Double, mCenterResidualMy As Double ' Невязки моментов относительно численного центра, Н*мм.
Private mCenterK(1 To 3, 1 To 3) As Double ' Собирается напрямую по малым плечам, без вычитания больших K.
Private mMaxIterations As Long ' Максимум итераций на одну ступень нагрузки.
Private mLoadSteps As Long ' На сколько ступеней дробится заданный вектор N/Mx/My.
Private mToleranceN As Double ' Абсолютный допуск равновесия по N во внутренних единицах.
Private mToleranceMx As Double ' Абсолютный допуск равновесия по Mx во внутренних единицах.
Private mToleranceMy As Double ' Абсолютный допуск равновесия по My во внутренних единицах.
Private mDiagnosticsEnabled As Boolean ' Управляет накоплением текстовой диагностики итераций.
Private mSolverMethod As String ' Выбранная численная ветка: Newton или Secant.
Private mSecantMaxRestarts As Long ' Сколько раз Secant может пересобрать приближенную матрицу чувствительности.
Private mDiagnosticLog As String ' Текстовый журнал ключевых итераций для тестов и диагностики.
Private mLastAlpha As Double ' Коэффициент изменения последнего принятого шага при подборе его величины.
Private mLastResidualNorm As Double ' Норма невязки после последней оценки состояния.
Private mLastStepNorm As Double ' Норма последней поправки [deltaEps, deltaKx, deltaKy].
Private mLastMatrixConditionEstimate As Double ' Грубая оценка вырожденности касательной матрицы.
Private mMatrixRestartCount As Long ' Сколько раз Secant перезапускал матрицу чувствительности.
Private mLineSearchUsed As Boolean ' True, если последний принятый шаг был меньше полного.
Private mLineSearchIterations As Long ' Число пробных коэффициентов alpha при последнем подборе шага.
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
Private mSnapshotSection As CSectionModel ' Модель последнего Solve/Evaluate; общая ссылка без копирования сетки.
Private mSnapshotConcrete As CMaterialDiagram ' Фактически переданная диаграмма, не повторный выбор по Config.
Private mSnapshotSteel As CMaterialDiagram ' Фактически переданная диаграмма арматуры.
Private mEvaluationReuseEnabled As Boolean ' Только внутри одного Solve с неизменными штатными диаграммами.
Private mPhysicalEvaluationValid As Boolean, mPhysicalEvaluationHasK As Boolean ' Готовность последнего физического НДС и наличие его касательной матрицы.
Private mPhysicalEvaluationRevision As Long, mPhysicalEvaluationPlane(1 To 3) As Double ' Версия геометрии и три параметра поля последнего физического НДС.
Private mPhysicalEvaluationForces(1 To 3) As Double ' Экстремумы/K остаются в своих полях; forces-only пробы их не меняют.
Private mForceEvaluationValid As Boolean, mForceEvaluationRevision As Long ' Готовность последнего расчёта усилий и соответствующая версия геометрии.
Private mForceEvaluationPlane(1 To 3) As Double, mForceEvaluationForces(1 To 3) As Double ' Три параметра поля деформаций и усилия N, Mx, My последнего вычисления.
Private mEvaluationReuseCount As Long, mTangentBuildCount As Long ' Технические счетчики фактической работы, не инженерные результаты.
Private mConcreteMinStrainIndex As Long, mConcreteMaxStrainIndex As Long ' Номера бетонных элементов с наименьшей и наибольшей деформациями.
Private mSteelMinStrainIndex As Long, mSteelMaxStrainIndex As Long ' Номера стержней с наименьшей и наибольшей деформациями.
Private mStrainExtremaValid As Boolean, mStrainExtremaRevision As Long ' Готовность сохранённых крайних деформаций и соответствующая версия геометрии.
Private mStrainExtremaAmbiguous As Boolean ' Почти равные разные деформации могут дать одинаковое отношение к пределу после округления.
Private mStrainExtremaPlane(1 To 3) As Double ' Экстремумы принадлежат только этой центрированной плоскости и ревизии модели.

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
    mDiagnosticsEnabled = True
    mSolverMethod = "Newton"
    mSecantMaxRestarts = 2
End Sub

' ------------------------------
' Чтение настроек решателя
' ------------------------------
' Принимает численные параметры CSystemSettingsReader. При наличии
' CUnitSystem переводит пользовательские допуски во внутренние единицы.
' Явно пустые обязательные настройки Solver вызывают ошибку ввода
' без подстановки начальных значений.
Public Sub ApplySettings(ByVal settings As CSystemSettingsReader, Optional ByVal units As CUnitSystem = Nothing)
    If settings Is Nothing Then Exit Sub
    mMaxIterations = settings.GetLong("Solver.MaxIterations", mMaxIterations, False)
    mLoadSteps = settings.GetLong("Solver.LoadSteps", mLoadSteps, False)
    mToleranceN = settings.GetDouble("Solver.ToleranceN", mToleranceN, False)
    mToleranceMx = settings.GetDouble("Solver.ToleranceMx", mToleranceMx, False)
    mToleranceMy = settings.GetDouble("Solver.ToleranceMy", mToleranceMy, False)
    If settings.HasKey("Solver.Method") Then
        mSolverMethod = settings.GetRawString("Solver.Method", vbNullString)
    Else
        mSolverMethod = "Newton"
    End If
    mSecantMaxRestarts = settings.GetLong("Solver.SecantMaxRestarts", mSecantMaxRestarts, False)
    If Not units Is Nothing Then
        ' Знак нагрузки не относится к допуску; неверный отрицательный ввод
        ' сохраняем для адресной проверки конфигурации до начала solve.
        mToleranceN = Sgn(mToleranceN) * Abs(units.InputForceToInternal(mToleranceN))
        mToleranceMx = Sgn(mToleranceMx) * Abs(units.InputMomentMxToInternal(mToleranceMx))
        mToleranceMy = Sgn(mToleranceMy) * Abs(units.InputMomentMyToInternal(mToleranceMy))
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
    CaptureMaterialContext section, concreteMaterial, steelMaterial
    If Not mSnapshotConcrete Is Nothing And Not mSnapshotSteel Is Nothing Then
        mEvaluationReuseEnabled = mSnapshotConcrete.IsFrozen And mSnapshotSteel.IsFrozen
    End If
    If Not ValidateSolverMethod() Then GoTo Finished
    RecordSectionEquilibriumSolve
    If mUseInitialState Then
        mEpsilon0 = mInitialEpsilon0
        mKappaX = mInitialKappaX
        mKappaY = mInitialKappaY
    End If
    mCenterEpsilon = mFrame.CenterStrain(mEpsilon0, mKappaX, mKappaY)

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
            GoTo Finished
        End If
        mLoadStepsCompleted = stepIndex
    Next stepIndex

    ' Финальная проверка выполняется уже на полной нагрузке: промежуточная сходимость ступеней
    ' сама по себе еще не гарантирует, что итоговые невязки попали в пользовательские допуски.
    EvaluateState section, concreteMaterial, steelMaterial, mCenterEpsilon, mKappaX, mKappaY, True, False, True
    SetResiduals targetN, targetMx, targetMy
    mConverged = IsConverged(targetN, targetMx, targetMy)
    If mConverged Then
        mStopReason = "Равновесие найдено методом " & mSolverMethod & " со ступенчатым приложением нагрузки."
        mFailureCode = sfcNone
    Else
        mStopReason = "Финальная проверка невязок равновесия не пройдена."
        mFailureCode = sfcNumericalFailure
    End If
Finished:
    mEvaluationReuseEnabled = False
    mPhysicalEvaluationValid = False: mForceEvaluationValid = False
    Exit Sub
Failed:
    mConverged = False
    ' Численный предел арифметики не является ошибкой структуры программы.
    ' Индексы/объекты и другие неожиданные runtime-ошибки сохраняют InternalError.
    Select Case Err.Number
        Case 6
            mFailureCode = sfcNumericalFailure
            mStopReason = "Равновесие не найдено: при численном расчете произошло переполнение диапазона Double."
        Case 11
            mFailureCode = sfcNumericalFailure
            mStopReason = "Равновесие не найдено: при численном расчете возникло деление на ноль."
        Case Else
            mFailureCode = sfcInternalError
            mStopReason = "Ошибка при вычислении равновесия: " & Err.Description
    End Select
    Resume Finished
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

    EvaluatePlaneCore section, concreteMaterial, steelMaterial, epsilon0, kappaX, kappaY, True
End Sub

' Для проб UltimateStrain сохраняет тот же stress API, физические экстремумы
' и усилия, но не собирает неиспользуемую аналитическую матрицу K.
Friend Sub EvaluatePhysicalStrainPlane(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal epsilon0 As Double, ByVal kappaX As Double, ByVal kappaY As Double)
    Dim buildTangent As Boolean
    buildTangent = Not (TypeOf concreteMaterial Is CMaterialDiagram And TypeOf steelMaterial Is CMaterialDiagram)
    EvaluatePlaneCore section, concreteMaterial, steelMaterial, epsilon0, kappaX, kappaY, buildTangent
End Sub

' Внутренний численный поиск передает деформацию центра напрямую, без
' потери точности на переносе в удаленное начало координат и обратно.
Friend Sub EvaluatePhysicalCenterPlane(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal centerEpsilon As Double, ByVal kappaX As Double, ByVal kappaY As Double)
    Dim buildTangent As Boolean
    buildTangent = Not (TypeOf concreteMaterial Is CMaterialDiagram And TypeOf steelMaterial Is CMaterialDiagram)
    EvaluatePlaneCore section, concreteMaterial, steelMaterial, centerEpsilon, kappaX, kappaY, buildTangent, True
End Sub

' Общая оценка известной плоскости; запрос физических данных независим от K.
' Публичная оценка по-прежнему возвращает полную касательную матрицу.
Private Sub EvaluatePlaneCore(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal epsilon0 As Double, ByVal kappaX As Double, ByVal kappaY As Double, ByVal buildTangent As Boolean, _
        Optional ByVal centeredPlane As Boolean = False)

    ClearResult
    If Not ValidateInputs(section, concreteMaterial, steelMaterial) Then Exit Sub
    CaptureMaterialContext section, concreteMaterial, steelMaterial
    mKappaX = kappaX
    mKappaY = kappaY
    If centeredPlane Then
        mCenterEpsilon = epsilon0
        mEpsilon0 = mFrame.OriginalStrain(epsilon0, kappaX, kappaY)
    Else
        mEpsilon0 = epsilon0
        mCenterEpsilon = mFrame.CenterStrain(epsilon0, kappaX, kappaY)
    End If
    EvaluateState section, concreteMaterial, steelMaterial, mCenterEpsilon, mKappaX, mKappaY, buildTangent, True, True
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
    mLastResidualNorm = ResidualNorm(mResidualN, mCenterResidualMx, mCenterResidualMy, targetN, targetMx, targetMy)
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
        mStopReason = "Ошибка настройки Solver.Method: допустимы только Newton или Secant."
        SolveSelectedLoadStep = False
    End If
End Function

' ------------------------------
' Шаг Newton
' ------------------------------
' На каждой итерации считаем текущую невязку равновесия, собираем касательную
' матрицу K и решаем K * delta = -R. Длина шага подбирается по снижению
' невязки, без фиксированных ограничений приращения деформаций.
Private Function SolveNewtonLoadStep(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal stepIndex As Long, ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double) As Boolean

    mInternalNewtonCallCount = mInternalNewtonCallCount + 1

    Dim iter As Long
    For iter = 1 To mMaxIterations
        mIterations = mIterations + 1
        RecordSectionEquilibriumIteration

        ' Считаем текущие внутренние усилия и касательную матрицу для этой
        ' плоскости деформаций. После этого можно проверить, достигнуто ли
        ' равновесие без дополнительного шага.
        EvaluateState section, concreteMaterial, steelMaterial, mCenterEpsilon, mKappaX, mKappaY, True, False, True
        SetResiduals targetN, targetMx, targetMy
        If IsConverged(targetN, targetMx, targetMy) Then
            AppendDiagnostics stepIndex, iter, 0#, targetN, targetMx, targetMy, "сошлось"
            SolveNewtonLoadStep = True
            Exit Function
        End If

        ' Решаем линейную поправку Newton. Правая часть равна минус невязке,
        ' потому что хотим сдвинуть состояние в сторону уменьшения R.
        Dim system As CLinearSystem3x3
        Set system = New CLinearSystem3x3
        Dim length As Double
        length = mFrame.LengthScale
        If Not system.Solve(mCenterK(1, 1), mCenterK(1, 2) / length, mCenterK(1, 3) / length, _
                mCenterK(2, 1) / length, mCenterK(2, 2) / length / length, mCenterK(2, 3) / length / length, _
                mCenterK(3, 1) / length, mCenterK(3, 2) / length / length, mCenterK(3, 3) / length / length, _
                -mResidualN, -mCenterResidualMx / length, -mCenterResidualMy / length) Then
            If TryRegularizedStep(section, concreteMaterial, steelMaterial, mCenterK, targetN, targetMx, targetMy) Then
                AppendDiagnostics stepIndex, iter, mLastAlpha, targetN, targetMx, targetMy, "регуляризованная поправка Newton"
                GoTo NewtonIterationCompleted
            End If
            mConverged = False
            mFailureCode = system.FailureCode
            mStopReason = "Не удалось решить касательную систему: " & system.StopReason
            AppendDiagnostics stepIndex, iter, 0#, targetN, targetMx, targetMy, mStopReason
            SolveNewtonLoadStep = False
            Exit Function
        End If

        Dim deltaEps As Double
        Dim deltaKx As Double
        Dim deltaKy As Double
        deltaEps = system.X1
        deltaKx = system.X2 / length
        deltaKy = system.X3 / length

        ' Полный центрированный шаг допускается сразу. Обратный поиск ниже
        ' автоматически уменьшает его по фактическому снижению невязки.
        mLastAlpha = AcceptStep(section, concreteMaterial, steelMaterial, targetN, targetMx, targetMy, deltaEps, deltaKx, deltaKy)
        AppendDiagnostics stepIndex, iter, mLastAlpha, targetN, targetMx, targetMy, "итерация"
        If mLastAlpha <= 0# Then
            If TryRegularizedStep(section, concreteMaterial, steelMaterial, mCenterK, targetN, targetMx, targetMy) Then
                AppendDiagnostics stepIndex, iter, mLastAlpha, targetN, targetMx, targetMy, "регуляризованная поправка Newton"
                GoTo NewtonIterationCompleted
            End If
            mConverged = False
            mFailureCode = sfcNumericalFailure
            If Len(mStopReason) = 0 Then mStopReason = "Подбор шага решения не смог уменьшить невязку равновесия."
            SolveNewtonLoadStep = False
            Exit Function
        End If
NewtonIterationCompleted:
    Next iter

    RefreshCurrentState section, concreteMaterial, steelMaterial, targetN, targetMx, targetMy
    If IsConverged(targetN, targetMx, targetMy) Then
        AppendDiagnostics stepIndex, mMaxIterations, mLastAlpha, targetN, targetMx, targetMy, "сошлось на последнем шаге"
        SolveNewtonLoadStep = True
        Exit Function
    End If
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
    Dim scaleKx As Double, scaleKy As Double ' Плечи для приведения кривизн к деформациям.
    SecantCoordinateScales section, scaleKx, scaleKy
    EvaluateSecantResidual section, concreteMaterial, steelMaterial, targetN, targetMx, targetMy, _
        mCenterEpsilon, mKappaX, mKappaY, mResidualN, mCenterResidualMx, mCenterResidualMy

    If Not BuildInitialSecantMatrix(section, concreteMaterial, steelMaterial, targetN, targetMx, targetMy, _
            mCenterEpsilon, mKappaX, mKappaY, mResidualN, mCenterResidualMx, mCenterResidualMy, b) Then
        SolveSecantLoadStep = False
        Exit Function
    End If

    Dim iter As Long
    For iter = 1 To mMaxIterations
        mIterations = mIterations + 1
        RecordSectionEquilibriumIteration
        mLastResidualNorm = ResidualNorm(mResidualN, mCenterResidualMx, mCenterResidualMy, targetN, targetMx, targetMy)

        If IsConverged(targetN, targetMx, targetMy) Then
            AppendDiagnostics stepIndex, iter, 0#, targetN, targetMx, targetMy, "Secant сошелся"
            SolveSecantLoadStep = True
            Exit Function
        End If

        Dim deltaEps As Double
        Dim deltaKx As Double
        Dim deltaKy As Double
        Dim linearFailureCode As ESolverFailureCode, linearFailureReason As String
        If Not SolveSecantCorrection(b, deltaEps, deltaKx, deltaKy, linearFailureCode, linearFailureReason) Then
            If TryRegularizedStep(section, concreteMaterial, steelMaterial, b, targetN, targetMx, targetMy) Then GoTo SecantRegularizedStep
            If Not RestartSecantMatrix(section, concreteMaterial, steelMaterial, targetN, targetMx, targetMy, _
                    mCenterEpsilon, mKappaX, mKappaY, mResidualN, mCenterResidualMx, mCenterResidualMy, b, "линейная система не решена") Then
                SolveSecantLoadStep = False
                Exit Function
            End If
            If Not SolveSecantCorrection(b, deltaEps, deltaKx, deltaKy, linearFailureCode, linearFailureReason) Then
                mConverged = False
                mFailureCode = linearFailureCode
                mStopReason = "Равновесие не найдено: система Secant не решилась после перезапуска. " & linearFailureReason
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
        oldEps = mCenterEpsilon: oldKx = mKappaX: oldKy = mKappaY
        oldRN = mResidualN: oldRMx = mCenterResidualMx: oldRMy = mCenterResidualMy

        mLastStepNorm = SecantStepNorm(deltaEps, deltaKx, deltaKy)
        If Not IsFiniteNumber(mLastStepNorm) Or mLastStepNorm <= 0# Then
            If Not RestartSecantMatrix(section, concreteMaterial, steelMaterial, targetN, targetMx, targetMy, _
                    mCenterEpsilon, mKappaX, mKappaY, mResidualN, mCenterResidualMx, mCenterResidualMy, b, "шаг слишком мал") Then
                SolveSecantLoadStep = False
                Exit Function
            End If
            AppendDiagnostics stepIndex, iter, 0#, targetN, targetMx, targetMy, "Secant перезапущен"
        Else
            mLastAlpha = AcceptStep(section, concreteMaterial, steelMaterial, targetN, targetMx, targetMy, deltaEps, deltaKx, deltaKy)
            If mLastAlpha <= 0# Then
                If TryRegularizedStep(section, concreteMaterial, steelMaterial, b, targetN, targetMx, targetMy) Then GoTo SecantRegularizedStep
                If Not RestartSecantMatrix(section, concreteMaterial, steelMaterial, targetN, targetMx, targetMy, _
                        oldEps, oldKx, oldKy, oldRN, oldRMx, oldRMy, b, "подбор шага не уменьшил невязку") Then
                    SolveSecantLoadStep = False
                    Exit Function
                End If
            Else
                EvaluateSecantResidual section, concreteMaterial, steelMaterial, targetN, targetMx, targetMy, _
                    mCenterEpsilon, mKappaX, mKappaY, mResidualN, mCenterResidualMx, mCenterResidualMy
                If IsConverged(targetN, targetMx, targetMy) Then
                    AppendDiagnostics stepIndex, iter, mLastAlpha, targetN, targetMx, targetMy, "Secant сошелся"
                    SolveSecantLoadStep = True
                    Exit Function
                End If
                If Not UpdateBroydenMatrix(b, mCenterEpsilon - oldEps, mKappaX - oldKx, mKappaY - oldKy, _
                        mResidualN - oldRN, mCenterResidualMx - oldRMx, mCenterResidualMy - oldRMy, scaleKx, scaleKy) Then
                    If Not RestartSecantMatrix(section, concreteMaterial, steelMaterial, targetN, targetMx, targetMy, _
                            mCenterEpsilon, mKappaX, mKappaY, mResidualN, mCenterResidualMx, mCenterResidualMy, b, "обновление Broyden не выполнено") Then
                        SolveSecantLoadStep = False
                        Exit Function
                    End If
                End If
                AppendDiagnostics stepIndex, iter, mLastAlpha, targetN, targetMx, targetMy, "итерация Secant"
            End If
        End If
        GoTo SecantIterationCompleted
SecantRegularizedStep:
        EvaluateSecantResidual section, concreteMaterial, steelMaterial, targetN, targetMx, targetMy, _
            mCenterEpsilon, mKappaX, mKappaY, mResidualN, mCenterResidualMx, mCenterResidualMy
        If IsConverged(targetN, targetMx, targetMy) Then
            SolveSecantLoadStep = True
            Exit Function
        End If
        BuildInitialSecantMatrix section, concreteMaterial, steelMaterial, targetN, targetMx, targetMy, _
            mCenterEpsilon, mKappaX, mKappaY, mResidualN, mCenterResidualMx, mCenterResidualMy, b
        AppendDiagnostics stepIndex, iter, mLastAlpha, targetN, targetMx, targetMy, "регуляризованная поправка Secant"
SecantIterationCompleted:
    Next iter

    RefreshCurrentState section, concreteMaterial, steelMaterial, targetN, targetMx, targetMy
    If IsConverged(targetN, targetMx, targetMy) Then
        AppendDiagnostics stepIndex, mMaxIterations, mLastAlpha, targetN, targetMx, targetMy, "Secant сошелся на последнем шаге"
        SolveSecantLoadStep = True
        Exit Function
    End If
    mConverged = False
    mFailureCode = sfcNumericalFailure
    mStopReason = "Равновесие не найдено: метод Secant достиг лимита итераций на ступени нагрузки " & CStr(stepIndex) & "."
    SolveSecantLoadStep = False
End Function

' На плато пробует поправку по системе J + mu*I, чтобы получить рабочий шаг.
' Добавка меняет только вычисление поправки, не напряжения или физическую
' жёсткость материала. Шаг принимается обычным AcceptStep, только если
' уменьшает фактическую невязку; исходные допуски равновесия сохраняются.
' Основа метода: PETSc, TS Pseudo-Timestepping for Steady-State Problems.
Private Function TryRegularizedStep(ByVal section As CSectionModel, ByVal concreteMaterial As Object, _
        ByVal steelMaterial As Object, ByRef matrix() As Double, ByVal targetN As Double, _
        ByVal targetMx As Double, ByVal targetMy As Double) As Boolean
    If Not mFrame.HasFullPlaneSupport Then Exit Function
    Dim a(1 To 3, 1 To 3) As Double, i As Long, j As Long, matrixScale As Double, length As Double
    length = mFrame.LengthScale
    For i = 1 To 3
        For j = 1 To 3
            a(i, j) = matrix(i, j)
            If i > 1 Then a(i, j) = a(i, j) / length
            If j > 1 Then a(i, j) = a(i, j) / length
            If Abs(a(i, j)) > matrixScale Then matrixScale = Abs(a(i, j))
        Next j
    Next i
    If matrixScale <= 0# Or Not IsFiniteNumber(matrixScale) Then Exit Function
    For i = 1 To 3
        For j = 1 To 3
            a(i, j) = a(i, j) / matrixScale
        Next j
    Next i
    Dim system As CLinearSystem3x3, shift As Double, attempt As Long
    Set system = New CLinearSystem3x3
    shift = 0.000000000001
    For attempt = 1 To 15
        If system.Solve(a(1, 1) + shift, a(1, 2), a(1, 3), a(2, 1), a(2, 2) + shift, a(2, 3), _
                a(3, 1), a(3, 2), a(3, 3) + shift, -mResidualN / matrixScale, _
                -mCenterResidualMx / length / matrixScale, -mCenterResidualMy / length / matrixScale) Then
            mLastAlpha = AcceptStep(section, concreteMaterial, steelMaterial, targetN, targetMx, targetMy, _
                system.X1, system.X2 / length, system.X3 / length)
            If mLastAlpha > 0# Then
                mFailureCode = sfcNone: mStopReason = vbNullString
                TryRegularizedStep = True
                Exit Function
            End If
        End If
        shift = shift * 10#
    Next attempt
End Function

' Пересчитывает внутренние усилия для текущей плоскости деформаций и сразу
' обновляет невязки относительно целевого вектора нагрузки.
Private Sub RefreshCurrentState(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double)
    EvaluateState section, concreteMaterial, steelMaterial, mCenterEpsilon, mKappaX, mKappaY, True, False, True
    SetResiduals targetN, targetMx, targetMy
End Sub

' Оценивает невязки метода Secant для заданного поля деформаций.
' Через SetResiduals сохраняет актуальные усилия и невязки
' для диагностики и следующего шага.
Private Sub EvaluateSecantResidual(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double, _
        ByVal eps0 As Double, ByVal kx As Double, ByVal ky As Double, _
        ByRef resN As Double, ByRef resMx As Double, ByRef resMy As Double)

    EvaluateState section, concreteMaterial, steelMaterial, eps0, kx, ky, False, False, True
    SetResiduals targetN, targetMx, targetMy
    resN = mResidualN: resMx = mCenterResidualMx: resMy = mCenterResidualMy
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
    hKx = MaxDouble(0.0000001 / mFrame.LengthScale, Abs(kx) * 0.0001)
    hKy = MaxDouble(0.0000001 / mFrame.LengthScale, Abs(ky) * 0.0001)

    BuildSecantJacobianColumn section, concreteMaterial, steelMaterial, targetN, targetMx, targetMy, _
        eps0 + hEps, kx, ky, eps0 - hEps, kx, ky, 2# * hEps, b, 1
    BuildSecantJacobianColumn section, concreteMaterial, steelMaterial, targetN, targetMx, targetMy, _
        eps0, kx + hKx, ky, eps0, kx - hKx, ky, 2# * hKx, b, 2
    BuildSecantJacobianColumn section, concreteMaterial, steelMaterial, targetN, targetMx, targetMy, _
        eps0, kx, ky + hKy, eps0, kx, ky - hKy, 2# * hKy, b, 3
    mLastMatrixConditionEstimate = EstimateMatrixCondition(b)
    BuildInitialSecantMatrix = IsFiniteNumber(mLastMatrixConditionEstimate)
    If Not BuildInitialSecantMatrix Then
        mConverged = False
        mFailureCode = sfcSingularTangent
        mStopReason = "Равновесие не найдено: матрица чувствительности Secant вырождена или ее численная оценка не представима."
    End If
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
    EvaluateStateInto section, concreteMaterial, steelMaterial, plusEps0, plusKx, plusKy, False, pN, pMx, pMy, False, True, True
    EvaluateStateInto section, concreteMaterial, steelMaterial, minusEps0, minusKx, minusKy, False, mN, mMx, mMy, False, True, True
    b(1, columnIndex) = (pN - mN) / h
    b(2, columnIndex) = (pMx - mMx) / h
    b(3, columnIndex) = (pMy - mMy) / h
End Sub

' Решает линейную систему Secant для поправки плоскости деформаций.
' Возвращает False, если матрица вырождена или поправка непредставима.
' Сохраняет машинную причину для финального отказа после Secant-перезапуска.
Private Function SolveSecantCorrection(ByRef b() As Double, ByRef deltaEps As Double, _
        ByRef deltaKx As Double, ByRef deltaKy As Double, _
        ByRef failureCode As ESolverFailureCode, ByRef failureReason As String) As Boolean

    Dim system As CLinearSystem3x3
    Set system = New CLinearSystem3x3
    failureCode = sfcNone: failureReason = vbNullString
    Dim length As Double
    length = mFrame.LengthScale
    If Not system.Solve(b(1, 1), b(1, 2) / length, b(1, 3) / length, _
            b(2, 1) / length, b(2, 2) / length / length, b(2, 3) / length / length, _
            b(3, 1) / length, b(3, 2) / length / length, b(3, 3) / length / length, _
            -mResidualN, -mCenterResidualMx / length, -mCenterResidualMy / length) Then
        failureCode = system.FailureCode
        failureReason = system.StopReason
        SolveSecantCorrection = False
        Exit Function
    End If
    deltaEps = system.X1
    deltaKx = system.X2 / length
    deltaKy = system.X3 / length
    SolveSecantCorrection = IsFiniteNumber(deltaEps) And IsFiniteNumber(deltaKx) And IsFiniteNumber(deltaKy)
    If Not SolveSecantCorrection Then
        failureCode = sfcNumericalFailure
        failureReason = "Поправка Secant выходит за допустимый численный диапазон."
    End If
End Function

' Обновляет приближенную матрицу чувствительности по формуле Бройдена.
' Поправка оценивается в координатах [epsilon0, scaleKx*kappaX, scaleKy*kappaY],
' чтобы размерная кривизна не исчезала на фоне безразмерной epsilon0.
' Матрица остается в исходных координатах; условие B*s = deltaR сохраняется.
' Если шаг слишком мал или новая матрица выглядит нечисловой, вызывающий код
' должен пересобрать матрицу с нуля.
Private Function UpdateBroydenMatrix(ByRef b() As Double, ByVal s1 As Double, ByVal s2 As Double, _
        ByVal s3 As Double, ByVal y1 As Double, ByVal y2 As Double, ByVal y3 As Double, _
        ByVal scaleKx As Double, ByVal scaleKy As Double) As Boolean

    Dim denom As Double
    denom = s1 * s1 + (s2 * scaleKx) * (s2 * scaleKx) + (s3 * scaleKy) * (s3 * scaleKy)
    If denom <= 0# Then Exit Function

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
        b(i, 2) = b(i, 2) + u(i) * (s(2) * scaleKx) * scaleKx / denom
        b(i, 3) = b(i, 3) + u(i) * (s(3) * scaleKy) * scaleKy / denom
    Next i

    mLastMatrixConditionEstimate = EstimateMatrixCondition(b)
    UpdateBroydenMatrix = IsFiniteNumber(mLastMatrixConditionEstimate)
End Function

' Определяет плечи по размаху расчетных точек, независимо от переноса начала
' координат. Они нужны только для размерно согласованного обновления Secant,
' не меняют поля деформаций, усилия и пользовательские допуски равновесия.
Private Sub SecantCoordinateScales(ByVal section As CSectionModel, _
        ByRef scaleKx As Double, ByRef scaleKy As Double)
    scaleKx = mFrame.LengthScale: scaleKy = mFrame.LengthScale
End Sub

' Перестраивает матрицу Secant при застое, неудачном подборе шага
' или обновлении Бройдена. Число таких перезапусков ограничено настройкой.
Private Function RestartSecantMatrix(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double, _
        ByVal eps0 As Double, ByVal kx As Double, ByVal ky As Double, _
        ByVal resN As Double, ByVal resMx As Double, ByVal resMy As Double, _
        ByRef b() As Double, ByVal reason As String) As Boolean

    If mMatrixRestartCount >= mSecantMaxRestarts Then
        mConverged = False
        mFailureCode = sfcNumericalFailure
        mStopReason = "Равновесие не найдено: превышен лимит перезапусков матрицы Secant; причина: " & reason
        Exit Function
    End If
    mMatrixRestartCount = mMatrixRestartCount + 1
    AppendDiagnostics mLoadStepsCompleted + 1, mIterations, 0#, targetN, targetMx, targetMy, "перезапуск матрицы Secant: " & reason
    RestartSecantMatrix = BuildInitialSecantMatrix(section, concreteMaterial, steelMaterial, targetN, targetMx, targetMy, _
        eps0, kx, ky, resN, resMx, resMy, b)
End Function

' Возвращает простую норму шага Secant по трем параметрам плоскости.
' Она нужна только для обнаружения численного застоя.
Private Function SecantStepNorm(ByVal deltaEps As Double, ByVal deltaKx As Double, ByVal deltaKy As Double) As Double
    SecantStepNorm = Abs(deltaEps) + mFrame.LengthScale * (Abs(deltaKx) + Abs(deltaKy))
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

' Общая автоматическая глобализация Newton/Secant: полный шаг, затем обратный
' поиск с условием достаточного снижения Armijo. Малый шаг не означает успех;
' выход без равновесия остается численным отказом. См. SUNDIALS KINSOL, 7.2.5.
Private Function AcceptStep(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double, _
        ByVal deltaEps As Double, ByVal deltaKx As Double, ByVal deltaKy As Double) As Double

    Dim baseNorm As Double
    baseNorm = ResidualNorm(mResidualN, mCenterResidualMx, mCenterResidualMy, targetN, targetMx, targetMy)

    Dim alpha As Double
    alpha = 1#
    mLineSearchIterations = 0
    mLineSearchUsed = False

    ' 52 деления соответствуют точности мантиссы Double. Дополнительно
    ' проверяется фактическая представимость каждой пробной плоскости.
    Do While mLineSearchIterations < NONLINEAR_MAX_BACKTRACKS
        mLineSearchIterations = mLineSearchIterations + 1
        Dim trialEps As Double
        Dim trialKx As Double
        Dim trialKy As Double
        trialEps = mCenterEpsilon + alpha * deltaEps
        trialKx = mKappaX + alpha * deltaKx
        trialKy = mKappaY + alpha * deltaKy
        If trialEps = mCenterEpsilon And trialKx = mKappaX And trialKy = mKappaY Then
            mStopReason = "Невязка не достигла допуска, но приращение плоскости деформаций меньше представимого шага Double."
            Exit Function
        End If

        Dim trialN As Double
        Dim trialMx As Double
        Dim trialMy As Double
        EvaluateStateInto section, concreteMaterial, steelMaterial, trialEps, trialKx, trialKy, False, trialN, trialMx, trialMy, False, True, True
        Dim trialNorm As Double
        trialNorm = ResidualNorm(trialN - targetN, trialMx - mFrame.CenterMx(targetN, targetMx), _
            trialMy - mFrame.CenterMy(targetN, targetMy), targetN, targetMx, targetMy)

        If NonlinearStepImproves(baseNorm, trialNorm, alpha) Then
            If alpha < 1# Then mLineSearchUsed = True
            mCenterEpsilon = trialEps
            mEpsilon0 = mFrame.OriginalStrain(trialEps, trialKx, trialKy)
            mKappaX = trialKx
            mKappaY = trialKy
            AcceptStep = alpha
            Exit Function
        End If

        Dim nextAlpha As Double
        nextAlpha = alpha / 2#
        If nextAlpha <= 0# Or nextAlpha >= alpha Then
            mStopReason = "Невязка не достигла допуска; дальнейшее уменьшение шага невозможно в машинной арифметике."
            Exit Function
        End If
        alpha = nextAlpha
    Loop

    AcceptStep = 0#
End Function

' Оценивает НДС и сохраняет внутренние усилия. buildTangent запрашивает
' матрицу жёсткости K, а collectPhysical — напряжения и экстремумы;
' их можно получить и без ненужной в данном вызове сборки K.
Private Sub EvaluateState(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal eps0 As Double, ByVal kx As Double, ByVal ky As Double, ByVal buildTangent As Boolean, _
        Optional ByVal collectPhysical As Boolean = False, Optional ByVal centeredInput As Boolean = False)
    EvaluateStateInto section, concreteMaterial, steelMaterial, eps0, kx, ky, buildTangent, _
        mNint, mCenterMx, mCenterMy, collectPhysical, centeredInput, True
    mMxint = mFrame.OriginalMx(mNint, mCenterMx)
    mMyint = mFrame.OriginalMy(mNint, mCenterMy)
End Sub

' Вычисляет внутренние усилия интегрированием напряжений бетонных элементов
' и стержней для заданного поля деформаций. Это основной повторяемый цикл
' решения, работающий с CSectionModel без обращения к Excel.
Private Sub EvaluateStateInto(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal eps0 As Double, ByVal kx As Double, ByVal ky As Double, ByVal buildTangent As Boolean, _
        ByRef nInt As Double, ByRef mxInt As Double, ByRef myInt As Double, _
        Optional ByVal collectPhysical As Boolean = False, Optional ByVal centeredInput As Boolean = False, _
        Optional ByVal centeredOutput As Boolean = False)

    Dim physical As Boolean, reusable As Boolean, revision As Long
    mFrame.Initialize section
    If Not centeredInput Then eps0 = mFrame.CenterStrain(eps0, kx, ky)
    physical = buildTangent Or collectPhysical
    If mEvaluationReuseEnabled Then
        reusable = section Is mSnapshotSection And concreteMaterial Is mSnapshotConcrete And steelMaterial Is mSnapshotSteel
        If reusable Then
            revision = section.Revision
            If physical Then
                If mPhysicalEvaluationValid And (Not buildTangent Or mPhysicalEvaluationHasK) Then
                    If revision = mPhysicalEvaluationRevision And eps0 = mPhysicalEvaluationPlane(1) And _
                            kx = mPhysicalEvaluationPlane(2) And ky = mPhysicalEvaluationPlane(3) Then
                        nInt = mPhysicalEvaluationForces(1): mxInt = mPhysicalEvaluationForces(2): myInt = mPhysicalEvaluationForces(3)
                        mEvaluationReuseCount = mEvaluationReuseCount + 1
                        GoTo ConvertOutput
                    End If
                End If
            ElseIf mForceEvaluationValid Then
                If revision = mForceEvaluationRevision And eps0 = mForceEvaluationPlane(1) And _
                        kx = mForceEvaluationPlane(2) And ky = mForceEvaluationPlane(3) Then
                    nInt = mForceEvaluationForces(1): mxInt = mForceEvaluationForces(2): myInt = mForceEvaluationForces(3)
                    mEvaluationReuseCount = mEvaluationReuseCount + 1
                    GoTo ConvertOutput
                End If
            End If
        End If
    End If
    If physical Then mPhysicalEvaluationValid = False Else mForceEvaluationValid = False
    nInt = 0#: mxInt = 0#: myInt = 0#
    mInternalForceEvaluationCount = mInternalForceEvaluationCount + 1
    If buildTangent Then
        ClearTangent
        mTangentBuildCount = mTangentBuildCount + 1
    End If
    If physical Then
        mStrainExtremaValid = False
        mStrainExtremaAmbiguous = False
        mConcreteMinStrainIndex = 0: mConcreteMaxStrainIndex = 0
        mSteelMinStrainIndex = 0: mSteelMaxStrainIndex = 0
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
    Dim originX As Double, originY As Double
    originX = mFrame.OriginX: originY = mFrame.OriginY

    If TypeName(concreteMaterial) = "CMaterialDiagram" Then Set concreteDiagram = concreteMaterial
    If TypeName(steelMaterial) = "CMaterialDiagram" Then Set steelDiagram = steelMaterial

    For i = 1 To section.ConcreteCount
        section.ConcreteElementForSolver i, xCoord, yCoord, area
        xCoord = xCoord - originX: yCoord = yCoord - originY
        strain = eps0 + kx * yCoord + ky * xCoord
        If physical Then
            If strain < mMinConcreteStrain Then
                If Abs(strain - mMinConcreteStrain) <= Abs(strain) * STRAIN_EXTREMA_ROUNDING_MARGIN Then mStrainExtremaAmbiguous = True
                mMinConcreteStrain = strain: mConcreteMinStrainIndex = i
            End If
            If strain > mMaxConcreteStrain Then
                If Abs(strain - mMaxConcreteStrain) <= Abs(strain) * STRAIN_EXTREMA_ROUNDING_MARGIN Then mStrainExtremaAmbiguous = True
                mMaxConcreteStrain = strain: mConcreteMaxStrainIndex = i
            End If
        End If
        If concreteDiagram Is Nothing Then
            stress = concreteMaterial.GetStress(strain)
            If buildTangent Then tangent = concreteMaterial.GetTangentModulus(strain)
        ElseIf physical Then
            concreteDiagram.EvaluateAtStrain strain, stress, tangent
        Else
            stress = concreteDiagram.StressAtStrain(strain)
        End If
        If physical Then
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
            xCoord = xCoord - originX: yCoord = yCoord - originY
            strain = eps0 + kx * yCoord + ky * xCoord
            If physical Then
                If strain < mMinSteelStrain Then
                    If Abs(strain - mMinSteelStrain) <= Abs(strain) * STRAIN_EXTREMA_ROUNDING_MARGIN Then mStrainExtremaAmbiguous = True
                    mMinSteelStrain = strain: mSteelMinStrainIndex = i
                End If
                If strain > mMaxSteelStrain Then
                    If Abs(strain - mMaxSteelStrain) <= Abs(strain) * STRAIN_EXTREMA_ROUNDING_MARGIN Then mStrainExtremaAmbiguous = True
                    mMaxSteelStrain = strain: mSteelMaxStrainIndex = i
                End If
                If Abs(strain) > mMaxAbsSteelStrain Then mMaxAbsSteelStrain = Abs(strain)
            End If
            If steelDiagram Is Nothing Then
                steelStress = steelMaterial.GetStress(strain)
                If buildTangent Then steelTangent = steelMaterial.GetTangentModulus(strain)
            ElseIf physical Then
                steelDiagram.EvaluateAtStrain strain, steelStress, steelTangent
            Else
                steelStress = steelDiagram.StressAtStrain(strain)
            End If
            If physical Then
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
            ElseIf physical Then
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
    If physical Then
        ' Если сжатой или растянутой зоны нет, её экстремум не сохраняем.
        ' Служебные начальные границы поиска не должны выдаваться за физическое напряжение.
        If Not mHasConcreteCompression Then mMinConcreteStress = 0#
        If Not mHasConcreteTension Then mMaxConcreteStress = 0#
        If Not mHasSteelCompression Then mMinSteelStress = 0#
        If Not mHasSteelTension Then mMaxSteelStress = 0#
        mStrainExtremaRevision = section.Revision
        mStrainExtremaPlane(1) = eps0: mStrainExtremaPlane(2) = kx: mStrainExtremaPlane(3) = ky
        mStrainExtremaValid = True
    End If
    If buildTangent Then BuildOriginalTangent
    If reusable Then
        If physical Then
            mPhysicalEvaluationRevision = revision
            mPhysicalEvaluationPlane(1) = eps0: mPhysicalEvaluationPlane(2) = kx: mPhysicalEvaluationPlane(3) = ky
            mPhysicalEvaluationForces(1) = nInt: mPhysicalEvaluationForces(2) = mxInt: mPhysicalEvaluationForces(3) = myInt
            mPhysicalEvaluationHasK = buildTangent: mPhysicalEvaluationValid = True
        Else
            mForceEvaluationRevision = revision
            mForceEvaluationPlane(1) = eps0: mForceEvaluationPlane(2) = kx: mForceEvaluationPlane(3) = ky
            mForceEvaluationForces(1) = nInt: mForceEvaluationForces(2) = mxInt: mForceEvaluationForces(3) = myInt
            mForceEvaluationValid = True
        End If
    End If
ConvertOutput:
    If Not centeredOutput Then
        mxInt = mFrame.OriginalMx(nInt, mxInt)
        myInt = mFrame.OriginalMy(nInt, myInt)
    End If
End Sub

' Очищает касательную матрицу перед новой сборкой производных равновесия.
Private Sub ClearTangent()
    Dim i As Long
    For i = 1 To 3
        mK(i, 1) = 0#: mK(i, 2) = 0#: mK(i, 3) = 0#
        mCenterK(i, 1) = 0#: mCenterK(i, 2) = 0#: mCenterK(i, 3) = 0#
    Next i
End Sub

' Добавляет вклад одного элемента в касательную матрицу равновесия.
' axialStiffness уже включает E_tangent * Area, поэтому здесь остается только
' умножение на координаты по правилу моментов.
Private Sub AddTangent(ByVal axialStiffness As Double, ByVal xCoord As Double, ByVal yCoord As Double)
    mCenterK(1, 1) = mCenterK(1, 1) + axialStiffness
    mCenterK(1, 2) = mCenterK(1, 2) + axialStiffness * yCoord
    mCenterK(1, 3) = mCenterK(1, 3) + axialStiffness * xCoord
    mCenterK(2, 1) = mCenterK(2, 1) + axialStiffness * yCoord
    mCenterK(2, 2) = mCenterK(2, 2) + axialStiffness * yCoord * yCoord
    mCenterK(2, 3) = mCenterK(2, 3) + axialStiffness * yCoord * xCoord
    mCenterK(3, 1) = mCenterK(3, 1) + axialStiffness * xCoord
    mCenterK(3, 2) = mCenterK(3, 2) + axialStiffness * xCoord * yCoord
    mCenterK(3, 3) = mCenterK(3, 3) + axialStiffness * xCoord * xCoord
End Sub

' Публичная K относится к прежнему началу координат. Численные поправки
' используют только напрямую собранную центрированную K, а не ее обратный перенос.
Private Sub BuildOriginalTangent()
    Dim x As Double, y As Double
    x = mFrame.OriginX: y = mFrame.OriginY
    mK(1, 1) = mCenterK(1, 1)
    mK(1, 2) = mCenterK(1, 2) + y * mCenterK(1, 1)
    mK(1, 3) = mCenterK(1, 3) + x * mCenterK(1, 1)
    mK(2, 1) = mK(1, 2): mK(3, 1) = mK(1, 3)
    mK(2, 2) = mCenterK(2, 2) + 2# * y * mCenterK(1, 2) + y * y * mCenterK(1, 1)
    mK(2, 3) = mCenterK(2, 3) + y * mCenterK(1, 3) + x * mCenterK(1, 2) + x * y * mCenterK(1, 1)
    mK(3, 2) = mK(2, 3)
    mK(3, 3) = mCenterK(3, 3) + 2# * x * mCenterK(1, 3) + x * x * mCenterK(1, 1)
End Sub

' Обновляет абсолютные и относительные невязки равновесия после очередной
' оценки внутренних усилий.
Private Sub SetResiduals(ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double)
    mResidualN = mNint - targetN
    mResidualMx = mMxint - targetMx
    mResidualMy = mMyint - targetMy
    mCenterResidualMx = mCenterMx - mFrame.CenterMx(targetN, targetMx)
    mCenterResidualMy = mCenterMy - mFrame.CenterMy(targetN, targetMy)
    mRelativeResidualN = RelativeResidual(mResidualN, targetN)
    mRelativeResidualMx = RelativeResidual(mResidualMx, targetMx)
    mRelativeResidualMy = RelativeResidual(mResidualMy, targetMy)
End Sub

' Проверяет равновесие N, Mx и My отдельно по их абсолютным допускам.
Private Function IsConverged(ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double) As Boolean
    IsConverged = Abs(mResidualN) <= mToleranceN And _
        Abs(mResidualMx) <= mToleranceMx And _
        Abs(mResidualMy) <= mToleranceMy And _
        Abs(mCenterResidualMx) <= mToleranceMx And Abs(mCenterResidualMy) <= mToleranceMy
End Function

' Масштабированная евклидова норма центрированных невязок для поиска шага.
' Масштабы не меняют компонентные допуски окончательной приемки равновесия.
Private Function ResidualNorm(ByVal resN As Double, ByVal resMx As Double, ByVal resMy As Double, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double) As Double
    Dim n As Double, mx As Double, my As Double, normScale As Double
    n = Abs(resN) / MaxDouble(mToleranceN, Abs(targetN) * 0.000000001)
    mx = Abs(resMx) / MaxDouble(mToleranceMx, Abs(mFrame.CenterMx(targetN, targetMx)) * 0.000000001)
    my = Abs(resMy) / MaxDouble(mToleranceMy, Abs(mFrame.CenterMy(targetN, targetMy)) * 0.000000001)
    normScale = MaxDouble(n, MaxDouble(mx, my))
    If normScale > 0# Then ResidualNorm = normScale * Sqr((n / normScale) ^ 2 + (mx / normScale) ^ 2 + (my / normScale) ^ 2)
End Function

' Возвращает относительную невязку, а при нулевой целевой нагрузке — абсолютную.
Private Function RelativeResidual(ByVal residual As Double, ByVal target As Double) As Double
    If Abs(target) > 0.000000000001 Then
        RelativeResidual = Abs(residual / target)
    Else
        RelativeResidual = Abs(residual)
    End If
End Function

' Добавляет подробную строку итерации, если включена диагностика отчёта или тестов.
' Сами численные результаты от ведения журнала не меняются.
Private Sub AppendDiagnostics(ByVal stepIndex As Long, ByVal iter As Long, ByVal alpha As Double, _
        ByVal targetN As Double, ByVal targetMx As Double, _
        ByVal targetMy As Double, ByVal state As String)
    If Not mDiagnosticsEnabled Then Exit Sub
    mLastResidualNorm = ResidualNorm(mResidualN, mCenterResidualMx, mCenterResidualMy, targetN, targetMx, targetMy)
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
        "; alpha=" & FormatNumberInvariant(alpha) & vbCrLf
End Sub

' Проверяет модель и настройки до численных итераций. Программная зависимость,
' входная геометрия и ошибочная конфигурация имеют разные машинные причины.
Friend Function ValidateProblem(ByVal section As CSectionModel, ByVal concreteMaterial As Object, _
        ByVal steelMaterial As Object) As Boolean
    ClearResult
    mSolverMethod = Trim$(mSolverMethod)
    If Not ValidateInputs(section, concreteMaterial, steelMaterial) Then Exit Function
    ValidateProblem = ValidateSolverMethod()
End Function

' Проверяет модель и материалы до итераций равновесия; ошибки ввода не маскируются численным
' статусом.
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
    ElseIf mToleranceN <= 0# Or mToleranceMx <= 0# Or mToleranceMy <= 0# Then
        SetValidationFailure sfcInvalidConfiguration, "Допуски Solver.ToleranceN/Mx/My должны быть положительными."

    Else
        ValidateInputs = True
    End If
End Function

' Сохраняет окончательную причину ошибки обязательных программных данных
' без признака несуществующего численного запуска.
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
        mStopReason = "Ошибка настройки Solver.Method: допустимы только Newton или Secant."
    End If
End Function

' Очищает результат последней попытки, сохраняя численные настройки
' и явно заданное начальное приближение.
Public Sub ClearResult()
    mEvaluationReuseEnabled = False
    mPhysicalEvaluationValid = False: mForceEvaluationValid = False
    mStrainExtremaValid = False
    mEvaluationReuseCount = 0: mTangentBuildCount = 0
    Set mSnapshotSection = Nothing
    Set mSnapshotConcrete = Nothing
    Set mSnapshotSteel = Nothing
    mConverged = False
    mStopReason = vbNullString
    mFailureCode = sfcNone
    mEpsilon0 = 0#: mKappaX = 0#: mKappaY = 0#
    mCenterEpsilon = 0#: mCenterMx = 0#: mCenterMy = 0#
    mCenterResidualMx = 0#: mCenterResidualMy = 0#
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

' Общий критерий предельных деформаций использует уже вычисленные экстремумы.
' После изменения модели или плоскости этот путь запрещен; равные значения
' сохраняют первый индекс, как последовательный обход элементов.
Friend Function TryGetStrainExtrema(ByVal section As CSectionModel, _
        ByRef concreteMin As Double, ByRef concreteMax As Double, _
        ByRef concreteMinIndex As Long, ByRef concreteMaxIndex As Long, _
        ByRef steelMin As Double, ByRef steelMax As Double, _
        ByRef steelMinIndex As Long, ByRef steelMaxIndex As Long) As Boolean
    If mStrainExtremaAmbiguous Then Exit Function
    If Not StrainExtremaContextMatches(section) Then Exit Function
    concreteMin = mMinConcreteStrain: concreteMax = mMaxConcreteStrain
    concreteMinIndex = mConcreteMinStrainIndex: concreteMaxIndex = mConcreteMaxStrainIndex
    steelMin = mMinSteelStrain: steelMax = mMaxSteelStrain
    steelMinIndex = mSteelMinStrainIndex: steelMaxIndex = mSteelMaxStrainIndex
    TryGetStrainExtrema = True
End Function

' Для выбора растянутого критерия нужны только значения, без выбора элемента
' при равенстве отношений. Почти совпадающие значения здесь не неоднозначны.
Friend Function TryGetConcreteStrainBounds(ByVal section As CSectionModel, _
        ByRef minimum As Double, ByRef maximum As Double) As Boolean
    If Not StrainExtremaContextMatches(section) Then Exit Function
    minimum = mMinConcreteStrain: maximum = mMaxConcreteStrain
    TryGetConcreteStrainBounds = True
End Function

' Проверяет точное совпадение контекста для повторного использования экстремумов деформаций.
Private Function StrainExtremaContextMatches(ByVal section As CSectionModel) As Boolean
    If Not mStrainExtremaValid Then Exit Function
    If Not section Is mSnapshotSection Then Exit Function
    If section.Revision <> mStrainExtremaRevision Then Exit Function
    StrainExtremaContextMatches = (mCenterEpsilon = mStrainExtremaPlane(1) And _
        mKappaX = mStrainExtremaPlane(2) And mKappaY = mStrainExtremaPlane(3))
End Function

' Запоминает исходные данные последней оценки НДС без отдельного объекта на элемент.
' Сохраняет возможность работать с общим материалом; для CMaterialDiagram
' также позволяет получить использованные диаграммы.
Private Sub CaptureMaterialContext(ByVal section As CSectionModel, ByVal concrete As Object, ByVal steel As Object)
    Set mSnapshotSection = section
    If mFrame Is Nothing Then Set mFrame = New CSectionNumericalFrame
    mFrame.Initialize section
    If TypeOf concrete Is CMaterialDiagram Then Set mSnapshotConcrete = concrete
    If TypeOf steel Is CMaterialDiagram Then Set mSnapshotSteel = steel
End Sub

' Копирует использованные диаграммы для сохранённого НДС без ссылки на рабочий решатель.
' Сохраняется вся диаграмма: продолжение бетона может участвовать
' в вычитании вытесненного бетона на месте стержня.
Friend Sub CopyMaterialContext(ByRef section As CSectionModel, ByRef concrete As CMaterialDiagram, _
        ByRef steel As CMaterialDiagram)
    Set section = mSnapshotSection
    If Not mSnapshotConcrete Is Nothing Then Set concrete = mSnapshotConcrete.Clone
    If Not mSnapshotSteel Is Nothing Then Set steel = mSnapshotSteel.Clone
End Sub

' Задаёт начальное поле деформаций для следующего решения.
' Расчёт трещин и несущей способности может передать близкое ранее найденное НДС.
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

' Удаляет заданное начальное приближение перед независимым решением нового состояния.
Public Sub ClearInitialState()
    mUseInitialState = False
    mInitialEpsilon0 = 0#: mInitialKappaX = 0#: mInitialKappaY = 0#
    mUseInitialLoads = False
    mInitialTargetN = 0#: mInitialTargetMx = 0#: mInitialTargetMy = 0#
End Sub

' Возвращает минимум двух чисел без обращения к функциям листа Excel в расчётном цикле.
Private Function MinDouble(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then MinDouble = a Else MinDouble = b
End Function

' Возвращает максимум двух чисел без обращения к функциям листа Excel в расчётном цикле.
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

' True, если равновесие достигнуто в заданных допусках по N/Mx/My.
Public Property Get Converged() As Boolean
    Converged = mConverged
End Property

' Русское пояснение итога; машинная причина хранится отдельно в mFailureCode.
Public Property Get StopReason() As String
    StopReason = mStopReason
End Property

' Возвращает машинную причину отказа последнего solve. Расчетные слои
' используют ее вместо анализа текста StopReason.
Public Property Get FailureCode() As ESolverFailureCode
    FailureCode = mFailureCode
End Property

' Равномерная составляющая деформации в формуле epsilon = epsilon0 + kappaX*y + kappaY*x.
Public Property Get Epsilon0() As Double
    Epsilon0 = mEpsilon0
End Property

' Кривизна, которая умножается на координату y и дает вклад изгиба вокруг X.
Public Property Get KappaX() As Double
    KappaX = mKappaX
End Property

' Кривизна, которая умножается на координату x и дает вклад изгиба вокруг Y.
Public Property Get KappaY() As Double
    KappaY = mKappaY
End Property

' Внутренняя продольная сила от всех бетонных и арматурных элементов.
Public Property Get Nint() As Double
    Nint = mNint
End Property

' Усилия в общем численном центре интегрируются напрямую по малым плечам.
' Поиск предела не восстанавливает их вычитанием больших исходных моментов.
Friend Sub NumericalEfforts(ByRef n As Double, ByRef mx As Double, ByRef my As Double)
    n = mNint: mx = mCenterMx: my = mCenterMy
End Sub

' Возвращает параметры сохранённого поля деформаций отдельными числами
' для проверки пределов всех элементов. Вызывающий код использует тот же
' порядок вычисления, что StrainAt, без повторного чтения центра на каждом элементе.
Friend Sub NumericalStrainPlane(ByRef epsilon As Double, ByRef kx As Double, ByRef ky As Double, _
        ByRef originX As Double, ByRef originY As Double)
    epsilon = mCenterEpsilon: kx = mKappaX: ky = mKappaY
    originX = mFrame.OriginX: originY = mFrame.OriginY
End Sub

' Физические проверки используют ту же плоскость и малые плечи, что ядро.
Friend Function StrainAt(ByVal x As Double, ByVal y As Double) As Double
    StrainAt = mCenterEpsilon + mKappaX * (y - mFrame.OriginY) + mKappaY * (x - mFrame.OriginX)
End Function

' Внутренний момент Mx, собранный интегрированием напряжений по элементам.
Public Property Get Mxint() As Double
    Mxint = mMxint
End Property

' Внутренний момент My, собранный интегрированием напряжений по элементам.
Public Property Get Myint() As Double
    Myint = mMyint
End Property

' Невязка равновесия по силе: Nint - Ntarget.
Public Property Get ResidualN() As Double
    ResidualN = mResidualN
End Property

' Невязка равновесия по моменту: Mxint - Mxtarget.
Public Property Get ResidualMx() As Double
    ResidualMx = mResidualMx
End Property

' Невязка равновесия по моменту: Myint - Mytarget.
Public Property Get ResidualMy() As Double
    ResidualMy = mResidualMy
End Property

' Относительная невязка N для диагностики качества найденного состояния.
Public Property Get RelativeResidualN() As Double
    RelativeResidualN = mRelativeResidualN
End Property

' Относительная невязка Mx для диагностики качества найденного состояния.
Public Property Get RelativeResidualMx() As Double
    RelativeResidualMx = mRelativeResidualMx
End Property

' Относительная невязка My для диагностики качества найденного состояния.
Public Property Get RelativeResidualMy() As Double
    RelativeResidualMy = mRelativeResidualMy
End Property

' Минимальная деформация бетона в последнем посчитанном состоянии.
Public Property Get MinConcreteStrain() As Double
    MinConcreteStrain = mMinConcreteStrain
End Property

' Максимальная деформация бетона в последнем посчитанном состоянии.
Public Property Get MaxConcreteStrain() As Double
    MaxConcreteStrain = mMaxConcreteStrain
End Property

' Максимальное по модулю сжимающее напряжение бетона в последнем состоянии.
Public Property Get MaxConcreteCompressionStress() As Double
    MaxConcreteCompressionStress = mMaxConcreteCompressionStress
End Property

' Наиболее сжимающее напряжение бетона со знаком, МПа.
Public Property Get MinConcreteStress() As Double
    MinConcreteStress = mMinConcreteStress
End Property

' Наиболее растягивающее напряжение бетона со знаком, МПа.
Public Property Get MaxConcreteStress() As Double
    MaxConcreteStress = mMaxConcreteStress
End Property

' True, если в последнем состоянии реально есть сжатый бетон.
Public Property Get HasConcreteCompression() As Boolean
    HasConcreteCompression = mHasConcreteCompression
End Property

' True, если бетон реально работает на растяжение, а не выключен диаграммой.
Public Property Get HasConcreteTension() As Boolean
    HasConcreteTension = mHasConcreteTension
End Property

' Наиболее сжатая деформация арматуры в последнем посчитанном состоянии.
Public Property Get MinSteelStrain() As Double
    MinSteelStrain = mMinSteelStrain
End Property

' Наиболее растянутая деформация арматуры в последнем посчитанном состоянии.
Public Property Get MaxSteelStrain() As Double
    MaxSteelStrain = mMaxSteelStrain
End Property

' Максимальная по модулю деформация арматуры в последнем состоянии.
Public Property Get MaxAbsSteelStrain() As Double
    MaxAbsSteelStrain = mMaxAbsSteelStrain
End Property

' Наиболее сжимающее напряжение стали со знаком, МПа.
Public Property Get MinSteelStress() As Double
    MinSteelStress = mMinSteelStress
End Property

' Наиболее растягивающее напряжение стали со знаком, МПа.
Public Property Get MaxSteelStress() As Double
    MaxSteelStress = mMaxSteelStress
End Property

' True, если есть сжатая арматура.
Public Property Get HasSteelCompression() As Boolean
    HasSteelCompression = mHasSteelCompression
End Property

' True, если есть растянутая арматура.
Public Property Get HasSteelTension() As Boolean
    HasSteelTension = mHasSteelTension
End Property

' Суммарное число итераций на всех ступенях текущего вызова Solve.
Public Property Get Iterations() As Long
    Iterations = mIterations
End Property

' Касательная матрица d(N,Mx,My)/d(epsilon0,kappaX,kappaY).
Public Property Get Tangent(ByVal rowIndex As Long, ByVal columnIndex As Long) As Double
    Tangent = mK(rowIndex, columnIndex)
End Property

' Текстовый журнал ключевых итераций для тестов и диагностики.
Public Property Get DiagnosticLog() As String
    DiagnosticLog = mDiagnosticLog
End Property

' Выбранная численная ветка: Newton или Secant.
Public Property Get SolverMethod() As String
    SolverMethod = mSolverMethod
End Property

' Норма невязки после последней оценки состояния.
Public Property Get LastResidualNorm() As Double
    LastResidualNorm = mLastResidualNorm
End Property

' Норма последней поправки [deltaEps, deltaKx, deltaKy].
Public Property Get LastStepNorm() As Double
    LastStepNorm = mLastStepNorm
End Property

' Грубая оценка вырожденности касательной матрицы.
Public Property Get LastMatrixConditionEstimate() As Double
    LastMatrixConditionEstimate = mLastMatrixConditionEstimate
End Property

' Максимум итераций на одну ступень нагрузки.
Public Property Let MaxIterations(ByVal value As Long)
    mMaxIterations = value
End Property

' На сколько ступеней дробится заданный вектор N/Mx/My.
Public Property Let LoadSteps(ByVal value As Long)
    mLoadSteps = value
End Property

' Абсолютный допуск равновесия по N во внутренних единицах.
Public Property Let ToleranceN(ByVal value As Double)
    mToleranceN = value
End Property

' Абсолютный допуск равновесия по Mx во внутренних единицах.
Public Property Let ToleranceMx(ByVal value As Double)
    mToleranceMx = value
End Property

' Абсолютный допуск равновесия по My во внутренних единицах.
Public Property Let ToleranceMy(ByVal value As Double)
    mToleranceMy = value
End Property

' Выбранная численная ветка: Newton или Secant.
Public Property Let SolverMethod(ByVal value As String)
    mSolverMethod = value
End Property

' Сколько раз Secant может пересобрать приближенную матрицу чувствительности.
Public Property Let SecantMaxRestarts(ByVal value As Long)
    mSecantMaxRestarts = value
End Property

' Включает подробный журнал при пользовательском txt-отчете и диагностических
' тестах. Отключение журнала не меняет численные шаги и критерии равновесия.
Public Property Let DiagnosticsEnabled(ByVal value As Boolean)
    mDiagnosticsEnabled = value
End Property

' ==========================================================================
' ДЛЯ ТЕСТОВ
' ==========================================================================

' ДЛЯ ТЕСТОВ: эти свойства и настройки раскрывают внутренний ход Newton/Secant. Они нужны
' unit- и regression-тестам, чтобы отличать ветки алгоритма, считать пробные
' пересчеты усилий и проверять рестарты матрицы. Пользовательский расчет
' использует сходимость, найденные деформации, внутренние усилия и DiagnosticLog.
Public Property Get LoadStepsCompleted() As Long
    LoadStepsCompleted = mLoadStepsCompleted
End Property

' ДЛЯ ТЕСТОВ: сколько раз пересчитывались внутренние усилия.
Public Property Get InternalForceEvaluationCount() As Long
    InternalForceEvaluationCount = mInternalForceEvaluationCount
End Property

' ДЛЯ ТЕСТОВ: фактическая сборка K и устраненные точные повторы одного Solve.
Friend Property Get TangentBuildCount() As Long
    TangentBuildCount = mTangentBuildCount
End Property

' ДЛЯ ТЕСТОВ: технические счетчики фактической работы, не инженерные результаты.
Friend Property Get EvaluationReuseCount() As Long
    EvaluationReuseCount = mEvaluationReuseCount
End Property

' ДЛЯ ТЕСТОВ: сколько раз запускалась ветка Newton на ступени нагрузки.
Public Property Get InternalNewtonCallCount() As Long
    InternalNewtonCallCount = mInternalNewtonCallCount
End Property

' ДЛЯ ТЕСТОВ: сколько раз Secant перезапускал матрицу чувствительности.
Public Property Get MatrixRestartCount() As Long
    MatrixRestartCount = mMatrixRestartCount
End Property
