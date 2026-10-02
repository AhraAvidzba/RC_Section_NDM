Attribute VB_Name = "modTestRegressionBaseline"
Option Explicit

' ==========================================================================
' Регрессионная проверка численных результатов
' ==========================================================================
' Модуль сравнивает текущие результаты с сохраненным baseline, чтобы правки
' интерфейса, документации или архитектуры не меняли расчетную математику.

Private Type TRegressionStats
    Passed As Long
    Failed As Long
    Report As String
End Type

' Выполняет пять опорных задач круглого сечения и проверки повторяемости,
' изменения геометрии и переноса начала координат. Возвращает численный
' протокол с прежними компонентными допусками и отдельными счетчиками отказов.
Public Function RunRegressionBaselineTests() As String
    On Error GoTo Failed

    Dim stats As TRegressionStats
    Dim t0 As Double
    t0 = Timer

    AppendLine stats, "REGRESSION_BASELINE_VERSION: Stage01 TEMPORARY_BASELINE"
    RunBaselineCase stats, "pure_compression", 300#, 0#, 0#, 40#, 8, 20#, 20#, 1, -100000#, 0#, 0#, False
    RunBaselineCase stats, "n_plus_mx_my_zero", 300#, 0#, 0#, 40#, 8, 20#, 20#, 1, -100000#, -4000000#, 0#, True
    RunBaselineCase stats, "n_plus_my_mx_zero", 300#, 0#, 0#, 40#, 8, 20#, 20#, 1, -100000#, 0#, -3000000#, True
    RunBaselineCase stats, "full_n_mx_my", 300#, 0#, 0#, 40#, 8, 20#, 20#, 1, -100000#, -4000000#, -3000000#, True
    RunBaselineCase stats, "symmetric_circle", 300#, 0#, 0#, 40#, 8, 20#, 10#, 1, -220000#, -6000000#, 0#, True
    TestRepeatedRun stats
    TestGeometryChange stats
    TestOriginShift stats

    AppendLine stats, "TOTAL_REGRESSION_BASELINE: passed=" & CStr(stats.Passed) & _
        "; failed=" & CStr(stats.Failed) & "; elapsedSec=" & FormatNumberInvariant(Timer - t0)
    RunRegressionBaselineTests = stats.Report
    Exit Function

Failed:
    RunRegressionBaselineTests = stats.Report & "RUNTIME ERROR: " & CStr(Err.Number) & _
        "; source=" & Err.Source & "; description=" & Err.Description
End Function

' Строит заданную круговую модель и решает N+Mx+My опорными материалами.
' Проверяет равновесие, при запросе отдельно ищет lambda и записывает
' плоскость, размеры модели и время в воспроизводимый baseline-протокол.
Private Sub RunBaselineCase(ByRef stats As TRegressionStats, ByVal caseName As String, _
        ByVal diameter As Double, ByVal centerX As Double, ByVal centerY As Double, _
        ByVal axisDistance As Double, ByVal barCount As Long, ByVal barDiameter As Double, _
        ByVal meshStep As Double, ByVal boundarySubdivisions As Long, _
        ByVal nValue As Double, ByVal mxValue As Double, ByVal myValue As Double, _
        ByVal calculateLambda As Boolean)

    Dim geom As CGeometryCircle
    Set geom = New CGeometryCircle
    geom.InitializeByDiameter diameter, centerX, centerY

    Dim mesh As CFiberMeshBuilder
    Set mesh = New CFiberMeshBuilder
    mesh.BuildMesh geom, meshStep, meshStep, 1, boundarySubdivisions

    Dim rebarBuilder As CCircleRebarLayoutBuilder
    Set rebarBuilder = New CCircleRebarLayoutBuilder
    Dim rebars As CRebarLayout
    Set rebars = rebarBuilder.Build(diameter, centerX, centerY, axisDistance, barCount, barDiameter, "A400")

    Dim solver As CSectionSolver
    Set solver = New CSectionSolver
    ConfigureBaselineSolver solver

    Dim t0 As Double
    t0 = Timer
    solver.Solve BuildGeneratedSectionModel(mesh, rebars), ProvisionalConcrete(), ProvisionalSteel(), nValue, mxValue, myValue
    Dim elapsed As Double
    elapsed = Timer - t0

    AssertTrue stats, caseName & ".converged", solver.Converged
    AssertEquilibrium stats, caseName, solver, nValue, mxValue, myValue

    Dim lambdaText As String
    lambdaText = "NA"
    If calculateLambda Then lambdaText = CalculateLambdaText(mesh, rebars, nValue, mxValue, myValue)

    AppendBaseline stats, caseName, solver, lambdaText, elapsed, mesh.FiberCount, rebars.Count
End Sub

' Дважды решает один и тот же LC на одинаковой круговой модели отдельными
' solver-ами. Проверяет равновесие и совпадение epsilon0/кривизн, защищая
' повторный запуск от случайного сохраненного состояния предыдущей задачи.
Private Sub TestRepeatedRun(ByRef stats As TRegressionStats)
    Dim geom As CGeometryCircle
    Set geom = New CGeometryCircle
    geom.InitializeByDiameter 300#

    Dim mesh As CFiberMeshBuilder
    Set mesh = New CFiberMeshBuilder
    mesh.BuildMesh geom, 20#, 20#, 1, 1

    Dim rebarBuilder As CCircleRebarLayoutBuilder
    Set rebarBuilder = New CCircleRebarLayoutBuilder
    Dim rebars As CRebarLayout
    Set rebars = rebarBuilder.Build(300#, 0#, 0#, 40#, 8, 20#, "A400")

    Dim firstSolver As CSectionSolver
    Set firstSolver = New CSectionSolver
    ConfigureBaselineSolver firstSolver
    firstSolver.Solve BuildGeneratedSectionModel(mesh, rebars), ProvisionalConcrete(), ProvisionalSteel(), -100000#, -4000000#, -3000000#

    Dim secondSolver As CSectionSolver
    Set secondSolver = New CSectionSolver
    ConfigureBaselineSolver secondSolver
    secondSolver.Solve BuildGeneratedSectionModel(mesh, rebars), ProvisionalConcrete(), ProvisionalSteel(), -100000#, -4000000#, -3000000#

    AssertTrue stats, "repeated_run.first.converged", firstSolver.Converged
    AssertTrue stats, "repeated_run.second.converged", secondSolver.Converged
    AssertClose stats, "repeated_run.epsilon0", secondSolver.Epsilon0, firstSolver.Epsilon0, 0.000000000001
    AssertClose stats, "repeated_run.kappaX", secondSolver.KappaX, firstSolver.KappaX, 0.0000000000001
    AssertClose stats, "repeated_run.kappaY", secondSolver.KappaY, firstSolver.KappaY, 0.0000000000001
    AppendBaseline stats, "repeated_run_first", firstSolver, "NA", 0#, mesh.FiberCount, rebars.Count
    AppendBaseline stats, "repeated_run_second", secondSolver, "NA", 0#, mesh.FiberCount, rebars.Count
End Sub

' Сравнивает круги 300 и 360 мм под одинаковыми усилиями. Обе задачи должны
' найти равновесие, а плоскости отличаться: изменение геометрии должно
' действительно доходить до расчетного ядра, а не использовать старую модель.
Private Sub TestGeometryChange(ByRef stats As TRegressionStats)
    Dim smallSolver As CSectionSolver
    Dim largeSolver As CSectionSolver
    Set smallSolver = SolveCircleDirect(300#, 0#, 0#, 40#, 8, 20#, 20#, -100000#, -4000000#, -3000000#)
    Set largeSolver = SolveCircleDirect(360#, 0#, 0#, 45#, 8, 20#, 20#, -100000#, -4000000#, -3000000#)

    AssertTrue stats, "geometry_change.small.converged", smallSolver.Converged
    AssertTrue stats, "geometry_change.large.converged", largeSolver.Converged
    AssertTrue stats, "geometry_change.state_changed", Abs(smallSolver.KappaX - largeSolver.KappaX) > 0.000000000001 Or _
        Abs(smallSolver.KappaY - largeSolver.KappaY) > 0.000000000001 Or _
        Abs(smallSolver.Epsilon0 - largeSolver.Epsilon0) > 0.000000000001
    AppendBaseline stats, "geometry_change_d300", smallSolver, "NA", 0#, 0, 0
    AppendBaseline stats, "geometry_change_d360", largeSolver, "NA", 0#, 0, 0
End Sub

' Переносит модель и усилия к другому началу координат через общий load adapter.
' Кривизны должны сохраняться, epsilon0 и моменты преобразоваться по заданному
' сдвигу; равновесие подтверждается отдельно в смещенной системе.
Private Sub TestOriginShift(ByRef stats As TRegressionStats)
    Dim baseN As Double
    Dim baseMx As Double
    Dim baseMy As Double
    baseN = -100000#
    baseMx = -4000000#
    baseMy = -3000000#

    Dim baseSolver As CSectionSolver
    Set baseSolver = SolveCircleDirect(300#, 0#, 0#, 40#, 8, 20#, 20#, baseN, baseMx, baseMy)

    Dim shiftX As Double
    Dim shiftY As Double
    shiftX = 75#
    shiftY = -50#

    Dim shiftedMx As Double
    Dim shiftedMy As Double
    Dim shiftedLoad As CSectionLoadState
    Set shiftedLoad = New CSectionLoadState
    shiftedLoad.Initialize baseN, baseMx, baseMy, shiftX, shiftY
    shiftedMx = shiftedLoad.InternalMx
    shiftedMy = shiftedLoad.InternalMy

    Dim shiftedSolver As CSectionSolver
    Set shiftedSolver = SolveCircleDirect(300#, shiftX, shiftY, 40#, 8, 20#, 20#, baseN, shiftedMx, shiftedMy)

    AssertTrue stats, "origin_shift.base.converged", baseSolver.Converged
    AssertTrue stats, "origin_shift.shifted.converged", shiftedSolver.Converged
    AssertClose stats, "origin_shift.kappaX", shiftedSolver.KappaX, baseSolver.KappaX, 0.00000000001
    AssertClose stats, "origin_shift.kappaY", shiftedSolver.KappaY, baseSolver.KappaY, 0.00000000001
    AssertClose stats, "origin_shift.epsilon0", shiftedSolver.Epsilon0, _
        baseSolver.Epsilon0 - baseSolver.KappaX * shiftY - baseSolver.KappaY * shiftX, 0.00000001
    AssertClose stats, "origin_shift.N.transform", baseN, -100000#, 0.000000001
    AssertClose stats, "origin_shift.Mx.transform", shiftedMx, 1000000#, 0.000000001
    AssertClose stats, "origin_shift.My.transform", shiftedMy, -10500000#, 0.000000001
    AssertEquilibrium stats, "origin_shift.shifted", shiftedSolver, baseN, shiftedMx, shiftedMy
    AppendBaseline stats, "origin_shift_base", baseSolver, "NA", 0#, 0, 0
    AppendBaseline stats, "origin_shift_shifted", shiftedSolver, "NA", 0#, 0, 0
End Sub

' Создает самостоятельную круговую модель с автоматическими стержнями и
' возвращает решенное прямое НДС. Использует фиксированные baseline-настройки,
' не меняя Config и не переиспользуя solver соседнего теста.
Private Function SolveCircleDirect(ByVal diameter As Double, ByVal centerX As Double, ByVal centerY As Double, _
        ByVal axisDistance As Double, ByVal barCount As Long, ByVal barDiameter As Double, ByVal meshStep As Double, _
        ByVal nValue As Double, ByVal mxValue As Double, ByVal myValue As Double) As CSectionSolver

    Dim geom As CGeometryCircle
    Set geom = New CGeometryCircle
    geom.InitializeByDiameter diameter, centerX, centerY

    Dim mesh As CFiberMeshBuilder
    Set mesh = New CFiberMeshBuilder
    mesh.BuildMesh geom, meshStep, meshStep, 1, 1

    Dim rebarBuilder As CCircleRebarLayoutBuilder
    Set rebarBuilder = New CCircleRebarLayoutBuilder
    Dim rebars As CRebarLayout
    Set rebars = rebarBuilder.Build(diameter, centerX, centerY, axisDistance, barCount, barDiameter, "A400")

    Dim solver As CSectionSolver
    Set solver = New CSectionSolver
    ConfigureBaselineSolver solver
    solver.Solve BuildGeneratedSectionModel(mesh, rebars), ProvisionalConcrete(), ProvisionalSteel(), nValue, mxValue, myValue
    Set SolveCircleDirect = solver
End Function

' Выполняет отдельный поиск предельного множителя при ненулевом базовом моменте.
' Возвращает число либо технический маркер baseline; это не display-статус книги.
Private Function CalculateLambdaText(ByVal mesh As CFiberMeshBuilder, ByVal rebars As CRebarLayout, _
        ByVal nValue As Double, ByVal mxBase As Double, ByVal myBase As Double) As String
    If Sqr(mxBase * mxBase + myBase * myBase) <= 0.000000001 Then
        CalculateLambdaText = "NA"
        Exit Function
    End If

    Dim cap As CCapacitySolver
    Set cap = New CCapacitySolver
    ConfigureBaselineCapacity cap
    cap.SolveByLoadMultiplier BuildGeneratedSectionModel(mesh, rebars), ProvisionalConcrete(), ProvisionalSteel(), nValue, mxBase, myBase
    If cap.Converged Then
        CalculateLambdaText = FormatNumberInvariant(cap.LambdaUltimate)
    Else
        CalculateLambdaText = "FAIL:" & cap.LimitState
    End If
End Function

' Фиксирует численные ограничения baseline независимо от сохраненного Config.
Private Sub ConfigureBaselineSolver(ByVal solver As CSectionSolver)
    solver.LoadSteps = 1
    solver.MaxIterations = 60
    solver.ToleranceN = 1#
    solver.ToleranceMx = 1000#
    solver.ToleranceMy = 1000#
    solver.LineSearchEnabled = True
    solver.MaxDeltaEpsilon0 = 0.0005
    solver.MaxDeltaKappa = 0.00001
End Sub

' Применяет неизменные bounds/tolerances и физические пределы контрольной Capacity-задачи.
Private Sub ConfigureBaselineCapacity(ByVal cap As CCapacitySolver)
    cap.InitialLambdaStep = 1#
    cap.MaxLambda = 64#
    cap.LambdaTolerance = 0.01
    cap.MaxRetries = 0
    cap.SolverBaseLoadSteps = 1
    cap.SolverMaxIterations = 60
    cap.ConcreteCompressionLimit = -0.0035
    cap.SteelStrainLimit = 0.025
End Sub

' Материалы этого baseline построены по фиксированным физическим точкам, без extension.
Private Function ProvisionalConcrete() As CMaterialDiagram
    Dim concrete As CMaterialDiagram
    Set concrete = New CMaterialDiagram
    concrete.Initialize -0.0015, -15.5, -0.0035, -15.5
    Set ProvisionalConcrete = concrete
End Function

Private Function ProvisionalSteel() As CMaterialDiagram
    Dim steel As CMaterialDiagram
    Set steel = New CMaterialDiagram
    steel.Initialize 0.00175, 350#, 0.025
    Set ProvisionalSteel = steel
End Function

' Проверяет три компоненты усилий с раздельными силовым/моментными допусками.
Private Sub AssertEquilibrium(ByRef stats As TRegressionStats, ByVal prefix As String, _
        ByVal solver As CSectionSolver, ByVal nValue As Double, ByVal mxValue As Double, ByVal myValue As Double)
    AssertLoadComponent stats, prefix & ".N", solver.Nint, nValue, 2#, 0.00001
    AssertLoadComponent stats, prefix & ".Mx", solver.Mxint, mxValue, 2000#, 0.00001
    AssertLoadComponent stats, prefix & ".My", solver.Myint, myValue, 2000#, 0.00001
End Sub

' Около нуля использует абсолютный допуск компоненты, иначе - относительный.
Private Sub AssertLoadComponent(ByRef stats As TRegressionStats, ByVal name As String, _
        ByVal actual As Double, ByVal expected As Double, ByVal zeroTolerance As Double, _
        ByVal relTolerance As Double)
    If Abs(expected) <= zeroTolerance Then
        AssertClose stats, name, actual, expected, zeroTolerance
    Else
        AssertRelative stats, name, actual, expected, relTolerance
    End If
End Sub

' Assertions ниже сохраняют результат каждой проверки, не прерывая остальные кейсы.
Private Sub AssertTrue(ByRef stats As TRegressionStats, ByVal name As String, ByVal condition As Boolean)
    If condition Then
        stats.Passed = stats.Passed + 1
        AppendLine stats, "OK: " & name
    Else
        stats.Failed = stats.Failed + 1
        AppendLine stats, "FAIL: " & name
    End If
End Sub

' Записывает абсолютную разницу и исходные числа, чтобы отказ можно было перепроверить.
Private Sub AssertClose(ByRef stats As TRegressionStats, ByVal name As String, ByVal actual As Double, _
        ByVal expected As Double, ByVal tolerance As Double)
    Dim diff As Double
    diff = Abs(actual - expected)
    If diff <= tolerance Then
        stats.Passed = stats.Passed + 1
        AppendLine stats, "OK: " & name & "; actual=" & FormatNumberInvariant(actual) & _
            "; expected=" & FormatNumberInvariant(expected) & "; absDiff=" & FormatNumberInvariant(diff)
    Else
        stats.Failed = stats.Failed + 1
        AppendLine stats, "FAIL: " & name & "; actual=" & FormatNumberInvariant(actual) & _
            "; expected=" & FormatNumberInvariant(expected) & "; absDiff=" & FormatNumberInvariant(diff)
    End If
End Sub

' Сравнивает относительную разницу; при практически нулевом эталоне не делит на него.
Private Sub AssertRelative(ByRef stats As TRegressionStats, ByVal name As String, ByVal actual As Double, _
        ByVal expected As Double, ByVal relTolerance As Double)
    Dim relDiff As Double
    If Abs(expected) <= 0.000000001 Then
        relDiff = Abs(actual - expected)
    Else
        relDiff = Abs((actual - expected) / expected)
    End If

    If relDiff <= relTolerance Then
        stats.Passed = stats.Passed + 1
        AppendLine stats, "OK: " & name & "; actual=" & FormatNumberInvariant(actual) & _
            "; expected=" & FormatNumberInvariant(expected) & "; relDiff=" & FormatNumberInvariant(relDiff)
    Else
        stats.Failed = stats.Failed + 1
        AppendLine stats, "FAIL: " & name & "; actual=" & FormatNumberInvariant(actual) & _
            "; expected=" & FormatNumberInvariant(expected) & "; relDiff=" & FormatNumberInvariant(relDiff)
    End If
End Sub

' Сохраняет плоскость, усилия, невязки, счетчики и время одним именованным протоколом.
' Текст lambda передан отдельно: отсутствующая/неуспешная точка не заменяется нулем.
Private Sub AppendBaseline(ByRef stats As TRegressionStats, ByVal caseName As String, _
        ByVal solver As CSectionSolver, ByVal lambdaText As String, ByVal elapsedSeconds As Double, _
        ByVal fiberCount As Long, ByVal rebarCount As Long)
    AppendLine stats, "BASELINE|" & caseName & _
        "|converged=" & BoolText(solver.Converged) & _
        "|eps0=" & FormatNumberInvariant(solver.Epsilon0) & _
        "|kappaX=" & FormatNumberInvariant(solver.KappaX) & _
        "|kappaY=" & FormatNumberInvariant(solver.KappaY) & _
        "|Nint=" & FormatNumberInvariant(solver.Nint) & _
        "|Mxint=" & FormatNumberInvariant(solver.Mxint) & _
        "|Myint=" & FormatNumberInvariant(solver.Myint) & _
        "|resN=" & FormatNumberInvariant(solver.ResidualN) & _
        "|resMx=" & FormatNumberInvariant(solver.ResidualMx) & _
        "|resMy=" & FormatNumberInvariant(solver.ResidualMy) & _
        "|lambda=" & lambdaText & _
        "|iterations=" & CStr(solver.Iterations) & _
        "|elapsedSec=" & FormatNumberInvariant(elapsedSeconds) & _
        "|fibers=" & CStr(fiberCount) & _
        "|rebars=" & CStr(rebarCount)
End Sub

' Однотипные helpers протокола задают устойчивые Boolean/числовые подписи и перевод строки.
Private Function BoolText(ByVal value As Boolean) As String
    If value Then
        BoolText = "True"
    Else
        BoolText = "False"
    End If
End Function

Private Sub AppendLine(ByRef stats As TRegressionStats, ByVal text As String)
    stats.Report = stats.Report & text & vbCrLf
End Sub

Private Function FormatNumberInvariant(ByVal value As Double) As String
    FormatNumberInvariant = Replace$(Format$(value, "0.############"), ",", ".")
End Function


