Attribute VB_Name = "modTestRegressionBaseline"
Option Explicit

' ==========================================================================
' ДЛЯ ТЕСТОВ: Регрессионная проверка численных результатов
' ==========================================================================
' Модуль сравнивает текущие результаты с сохраненным baseline, чтобы правки
' интерфейса, документации или архитектуры не меняли расчетную математику.

Private Type TRegressionStats
    Passed As Long
    Failed As Long
    Report As String
End Type

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Выполняет пять опорных задач круглого сечения и проверки повторяемости,
' изменения геометрии и переноса начала координат. Возвращает численный
' протокол с прежними компонентными допусками и отдельными счетчиками отказов.
Public Function RunRegressionBaselineTests() As String
    On Error GoTo Failed

    Dim stats As TRegressionStats
    Dim t0 As Double
    t0 = Timer

    TestAppendLine stats, "REGRESSION_BASELINE_VERSION: Stage01 TEMPORARY_BASELINE"
    TestRunBaselineCase stats, "pure_compression", 300#, 0#, 0#, 40#, 8, 20#, 20#, 1, -100000#, 0#, 0#, False
    TestRunBaselineCase stats, "n_plus_mx_my_zero", 300#, 0#, 0#, 40#, 8, 20#, 20#, 1, -100000#, -4000000#, 0#, True
    TestRunBaselineCase stats, "n_plus_my_mx_zero", 300#, 0#, 0#, 40#, 8, 20#, 20#, 1, -100000#, 0#, -3000000#, True
    TestRunBaselineCase stats, "full_n_mx_my", 300#, 0#, 0#, 40#, 8, 20#, 20#, 1, -100000#, -4000000#, -3000000#, True
    TestRunBaselineCase stats, "symmetric_circle", 300#, 0#, 0#, 40#, 8, 20#, 10#, 1, -220000#, -6000000#, 0#, True
    TestRepeatedRun stats
    TestGeometryChange stats
    TestOriginShift stats

    TestAppendLine stats, "TOTAL_REGRESSION_BASELINE: passed=" & CStr(stats.Passed) & _
        "; failed=" & CStr(stats.Failed) & "; elapsedSec=" & TestFormatNumberInvariant(Timer - t0)
    RunRegressionBaselineTests = stats.Report
    Exit Function

Failed:
    RunRegressionBaselineTests = stats.Report & "RUNTIME ERROR: " & CStr(Err.Number) & _
        "; source=" & Err.Source & "; description=" & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Строит заданную круговую модель и решает N+Mx+My опорными материалами.
' Проверяет равновесие, при запросе отдельно ищет lambda и записывает
' плоскость, размеры модели и время в воспроизводимый baseline-протокол.
Private Sub TestRunBaselineCase(ByRef stats As TRegressionStats, ByVal caseName As String, _
        ByVal diameter As Double, ByVal centerX As Double, ByVal centerY As Double, _
        ByVal axisDistance As Double, ByVal barCount As Long, ByVal barDiameter As Double, _
        ByVal meshStep As Double, ByVal boundarySubdivisions As Long, _
        ByVal nValue As Double, ByVal mxValue As Double, ByVal myValue As Double, _
        ByVal calculateLambda As Boolean)

    Dim geom As CGeometryRegion
    Set geom = TestFixtureDisk(diameter, centerX, centerY)

    Dim mesh As CFiberMeshBuilder
    Set mesh = New CFiberMeshBuilder
    mesh.BuildMesh geom, meshStep, meshStep, 1, boundarySubdivisions

    Dim facade As CSectionDefinitionBuilder: Set facade = New CSectionDefinitionBuilder
    facade.Initialize "BaselineCirclePaths"
    Dim rebars As CRebarLayout
    Set rebars = TestFixtureClosedCountRow(facade.CirclePath("Boundary", centerX, centerY, diameter / 2#), _
        facade.CirclePath("Axis", centerX, centerY, diameter / 2# - axisDistance), barCount, barDiameter, "A400", "Circle", "Круг: основной ряд арматуры")

    Dim solver As CSectionSolver
    Set solver = New CSectionSolver
    TestConfigureBaselineSolver solver

    Dim t0 As Double
    t0 = Timer
    solver.Solve TestBuildGeneratedSectionModel(mesh, rebars), TestProvisionalConcrete(), TestProvisionalSteel(), nValue, mxValue, myValue
    Dim elapsed As Double
    elapsed = Timer - t0

    TestAssertTrue stats, caseName & ".converged", solver.Converged
    TestAssertEquilibrium stats, caseName, solver, nValue, mxValue, myValue

    Dim lambdaText As String
    lambdaText = "NA"
    If calculateLambda Then lambdaText = TestCalculateLambdaText(mesh, rebars, nValue, mxValue, myValue)

    TestAppendBaseline stats, caseName, solver, lambdaText, elapsed, mesh.FiberCount, rebars.Count
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Дважды решает одинаковое сочетание на круглом сечении разными экземплярами
' решателя. Проверяет равновесие и совпадение epsilon0 и кривизн: предыдущий запуск
' не должен влиять на следующий.
Private Sub TestRepeatedRun(ByRef stats As TRegressionStats)
    Dim geom As CGeometryRegion
    Set geom = TestFixtureDisk(300#)

    Dim mesh As CFiberMeshBuilder
    Set mesh = New CFiberMeshBuilder
    mesh.BuildMesh geom, 20#, 20#, 1, 1

    Dim facade As CSectionDefinitionBuilder: Set facade = New CSectionDefinitionBuilder
    facade.Initialize "BaselineCirclePaths"
    Dim rebars As CRebarLayout
    Set rebars = TestFixtureClosedCountRow(facade.CirclePath("Boundary", 0#, 0#, 150#), _
        facade.CirclePath("Axis", 0#, 0#, 110#), 8, 20#, "A400", "Circle", "Круг: основной ряд арматуры")

    Dim firstSolver As CSectionSolver
    Set firstSolver = New CSectionSolver
    TestConfigureBaselineSolver firstSolver
    firstSolver.Solve TestBuildGeneratedSectionModel(mesh, rebars), TestProvisionalConcrete(), TestProvisionalSteel(), -100000#, -4000000#, -3000000#

    Dim secondSolver As CSectionSolver
    Set secondSolver = New CSectionSolver
    TestConfigureBaselineSolver secondSolver
    secondSolver.Solve TestBuildGeneratedSectionModel(mesh, rebars), TestProvisionalConcrete(), TestProvisionalSteel(), -100000#, -4000000#, -3000000#

    TestAssertTrue stats, "repeated_run.first.converged", firstSolver.Converged
    TestAssertTrue stats, "repeated_run.second.converged", secondSolver.Converged
    TestAssertClose stats, "repeated_run.epsilon0", secondSolver.Epsilon0, firstSolver.Epsilon0, 0.000000000001
    TestAssertClose stats, "repeated_run.kappaX", secondSolver.KappaX, firstSolver.KappaX, 0.0000000000001
    TestAssertClose stats, "repeated_run.kappaY", secondSolver.KappaY, firstSolver.KappaY, 0.0000000000001
    TestAppendBaseline stats, "repeated_run_first", firstSolver, "NA", 0#, mesh.FiberCount, rebars.Count
    TestAppendBaseline stats, "repeated_run_second", secondSolver, "NA", 0#, mesh.FiberCount, rebars.Count
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает круги 300 и 360 мм под одинаковыми усилиями. Обе задачи должны
' найти равновесие, а плоскости отличаться: изменение геометрии должно
' действительно доходить до расчетного ядра, а не использовать старую модель.
Private Sub TestGeometryChange(ByRef stats As TRegressionStats)
    Dim smallSolver As CSectionSolver
    Dim largeSolver As CSectionSolver
    Set smallSolver = TestSolveCircleDirect(300#, 0#, 0#, 40#, 8, 20#, 20#, -100000#, -4000000#, -3000000#)
    Set largeSolver = TestSolveCircleDirect(360#, 0#, 0#, 45#, 8, 20#, 20#, -100000#, -4000000#, -3000000#)

    TestAssertTrue stats, "geometry_change.small.converged", smallSolver.Converged
    TestAssertTrue stats, "geometry_change.large.converged", largeSolver.Converged
    TestAssertTrue stats, "geometry_change.state_changed", Abs(smallSolver.KappaX - largeSolver.KappaX) > 0.000000000001 Or _
        Abs(smallSolver.KappaY - largeSolver.KappaY) > 0.000000000001 Or _
        Abs(smallSolver.Epsilon0 - largeSolver.Epsilon0) > 0.000000000001
    TestAppendBaseline stats, "geometry_change_d300", smallSolver, "NA", 0#, 0, 0
    TestAppendBaseline stats, "geometry_change_d360", largeSolver, "NA", 0#, 0, 0
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
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
    Set baseSolver = TestSolveCircleDirect(300#, 0#, 0#, 40#, 8, 20#, 20#, baseN, baseMx, baseMy)

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
    Set shiftedSolver = TestSolveCircleDirect(300#, shiftX, shiftY, 40#, 8, 20#, 20#, baseN, shiftedMx, shiftedMy)

    TestAssertTrue stats, "origin_shift.base.converged", baseSolver.Converged
    TestAssertTrue stats, "origin_shift.shifted.converged", shiftedSolver.Converged
    TestAssertClose stats, "origin_shift.kappaX", shiftedSolver.KappaX, baseSolver.KappaX, 0.00000000001
    TestAssertClose stats, "origin_shift.kappaY", shiftedSolver.KappaY, baseSolver.KappaY, 0.00000000001
    TestAssertClose stats, "origin_shift.epsilon0", shiftedSolver.Epsilon0, _
        baseSolver.Epsilon0 - baseSolver.KappaX * shiftY - baseSolver.KappaY * shiftX, 0.00000001
    TestAssertClose stats, "origin_shift.N.transform", baseN, -100000#, 0.000000001
    TestAssertClose stats, "origin_shift.Mx.transform", shiftedMx, 1000000#, 0.000000001
    TestAssertClose stats, "origin_shift.My.transform", shiftedMy, -10500000#, 0.000000001
    TestAssertEquilibrium stats, "origin_shift.shifted", shiftedSolver, baseN, shiftedMx, shiftedMy
    TestAppendBaseline stats, "origin_shift_base", baseSolver, "NA", 0#, 0, 0
    TestAppendBaseline stats, "origin_shift_shifted", shiftedSolver, "NA", 0#, 0, 0
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Создаёт отдельную круговую модель с автоматически размещённой арматурой
' и рассчитывает НДС при заданной нагрузке. Использует фиксированные эталонные
' настройки без изменения Config и повторного использования соседнего решателя.
Private Function TestSolveCircleDirect(ByVal diameter As Double, ByVal centerX As Double, ByVal centerY As Double, _
        ByVal axisDistance As Double, ByVal barCount As Long, ByVal barDiameter As Double, ByVal meshStep As Double, _
        ByVal nValue As Double, ByVal mxValue As Double, ByVal myValue As Double) As CSectionSolver

    Dim geom As CGeometryRegion
    Set geom = TestFixtureDisk(diameter, centerX, centerY)

    Dim mesh As CFiberMeshBuilder
    Set mesh = New CFiberMeshBuilder
    mesh.BuildMesh geom, meshStep, meshStep, 1, 1

    Dim facade As CSectionDefinitionBuilder: Set facade = New CSectionDefinitionBuilder
    facade.Initialize "BaselineCirclePaths"
    Dim rebars As CRebarLayout
    Set rebars = TestFixtureClosedCountRow(facade.CirclePath("Boundary", centerX, centerY, diameter / 2#), _
        facade.CirclePath("Axis", centerX, centerY, diameter / 2# - axisDistance), barCount, barDiameter, "A400", "Circle", "Круг: основной ряд арматуры")

    Dim solver As CSectionSolver
    Set solver = New CSectionSolver
    TestConfigureBaselineSolver solver
    solver.Solve TestBuildGeneratedSectionModel(mesh, rebars), TestProvisionalConcrete(), TestProvisionalSteel(), nValue, mxValue, myValue
    Set TestSolveCircleDirect = solver
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Выполняет отдельный поиск предельного множителя при ненулевом базовом моменте.
' Возвращает число или служебную отметку файла эталона; пользовательский статус не назначает.
Private Function TestCalculateLambdaText(ByVal mesh As CFiberMeshBuilder, ByVal rebars As CRebarLayout, _
        ByVal nValue As Double, ByVal mxBase As Double, ByVal myBase As Double) As String
    If Sqr(mxBase * mxBase + myBase * myBase) <= 0.000000001 Then
        TestCalculateLambdaText = "NA"
        Exit Function
    End If

    Dim cap As CCapacitySolver
    Set cap = New CCapacitySolver
    TestConfigureBaselineCapacity cap
    cap.SolveByLoadMultiplier TestBuildGeneratedSectionModel(mesh, rebars), TestProvisionalConcrete(), TestProvisionalSteel(), nValue, mxBase, myBase
    If cap.Converged Then
        TestCalculateLambdaText = TestFormatNumberInvariant(cap.LambdaUltimate)
    Else
        TestCalculateLambdaText = "FAIL:" & cap.LimitState
    End If
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Фиксирует численные ограничения baseline независимо от сохраненного Config.
Private Sub TestConfigureBaselineSolver(ByVal solver As CSectionSolver)
    solver.LoadSteps = 1
    solver.MaxIterations = 60
    solver.ToleranceN = 1#
    solver.ToleranceMx = 1000#
    solver.ToleranceMy = 1000#
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Задаёт фиксированные границы поиска, допуски и физические пределы контрольной задачи несущей
' способности.
Private Sub TestConfigureBaselineCapacity(ByVal cap As CCapacitySolver)
    cap.InitialLambdaStep = 1#
    cap.MaxLambda = 64#
    cap.LambdaTolerance = 0.01
    cap.MaxRetries = 0
    cap.SolverBaseLoadSteps = 1
    cap.SolverMaxIterations = 60
    cap.ConcreteCompressionLimit = -0.0035
    cap.SteelStrainLimit = 0.025
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Материалы эталона построены по фиксированным физическим точкам без численного продолжения
' диаграмм.
Private Function TestProvisionalConcrete() As CMaterialDiagram
    Dim concrete As CMaterialDiagram
    Set concrete = New CMaterialDiagram
    concrete.TestInitialize -0.0015, -15.5, -0.0035, -15.5
    Set TestProvisionalConcrete = concrete
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Создает физическую диаграмму арматуры baseline: 350 МПа, начало плато
' 0.00175 и предельная деформация 0.025; расширение не включается.
Private Function TestProvisionalSteel() As CMaterialDiagram
    Dim steel As CMaterialDiagram
    Set steel = New CMaterialDiagram
    steel.TestInitialize 0.00175, 350#, 0.025
    Set TestProvisionalSteel = steel
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет три компоненты усилий с раздельными силовым/моментными допусками.
Private Sub TestAssertEquilibrium(ByRef stats As TRegressionStats, ByVal prefix As String, _
        ByVal solver As CSectionSolver, ByVal nValue As Double, ByVal mxValue As Double, ByVal myValue As Double)
    TestAssertLoadComponent stats, prefix & ".N", solver.Nint, nValue, 2#, 0.00001
    TestAssertLoadComponent stats, prefix & ".Mx", solver.Mxint, mxValue, 2000#, 0.00001
    TestAssertLoadComponent stats, prefix & ".My", solver.Myint, myValue, 2000#, 0.00001
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Около нуля использует абсолютный допуск компоненты, иначе - относительный.
Private Sub TestAssertLoadComponent(ByRef stats As TRegressionStats, ByVal name As String, _
        ByVal actual As Double, ByVal expected As Double, ByVal zeroTolerance As Double, _
        ByVal relTolerance As Double)
    If Abs(expected) <= zeroTolerance Then
        TestAssertClose stats, name, actual, expected, zeroTolerance
    Else
        TestAssertRelative stats, name, actual, expected, relTolerance
    End If
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Assertions ниже сохраняют результат каждой проверки, не прерывая остальные кейсы.
Private Sub TestAssertTrue(ByRef stats As TRegressionStats, ByVal name As String, ByVal condition As Boolean)
    If condition Then
        stats.Passed = stats.Passed + 1
        TestAppendLine stats, "OK: " & name
    Else
        stats.Failed = stats.Failed + 1
        TestAppendLine stats, "FAIL: " & name
    End If
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Записывает абсолютную разницу и исходные числа, чтобы отказ можно было перепроверить.
Private Sub TestAssertClose(ByRef stats As TRegressionStats, ByVal name As String, ByVal actual As Double, _
        ByVal expected As Double, ByVal tolerance As Double)
    Dim diff As Double
    diff = Abs(actual - expected)
    If diff <= tolerance Then
        stats.Passed = stats.Passed + 1
        TestAppendLine stats, "OK: " & name & "; actual=" & TestFormatNumberInvariant(actual) & _
            "; expected=" & TestFormatNumberInvariant(expected) & "; absDiff=" & TestFormatNumberInvariant(diff)
    Else
        stats.Failed = stats.Failed + 1
        TestAppendLine stats, "FAIL: " & name & "; actual=" & TestFormatNumberInvariant(actual) & _
            "; expected=" & TestFormatNumberInvariant(expected) & "; absDiff=" & TestFormatNumberInvariant(diff)
    End If
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает относительную разницу; при практически нулевом эталоне не делит на него.
Private Sub TestAssertRelative(ByRef stats As TRegressionStats, ByVal name As String, ByVal actual As Double, _
        ByVal expected As Double, ByVal relTolerance As Double)
    Dim relDiff As Double
    If Abs(expected) <= 0.000000001 Then
        relDiff = Abs(actual - expected)
    Else
        relDiff = Abs((actual - expected) / expected)
    End If

    If relDiff <= relTolerance Then
        stats.Passed = stats.Passed + 1
        TestAppendLine stats, "OK: " & name & "; actual=" & TestFormatNumberInvariant(actual) & _
            "; expected=" & TestFormatNumberInvariant(expected) & "; relDiff=" & TestFormatNumberInvariant(relDiff)
    Else
        stats.Failed = stats.Failed + 1
        TestAppendLine stats, "FAIL: " & name & "; actual=" & TestFormatNumberInvariant(actual) & _
            "; expected=" & TestFormatNumberInvariant(expected) & "; relDiff=" & TestFormatNumberInvariant(relDiff)
    End If
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сохраняет плоскость, усилия, невязки, счетчики и время одним именованным протоколом.
' Текст lambda передан отдельно: отсутствующая/неуспешная точка не заменяется нулем.
Private Sub TestAppendBaseline(ByRef stats As TRegressionStats, ByVal caseName As String, _
        ByVal solver As CSectionSolver, ByVal lambdaText As String, ByVal elapsedSeconds As Double, _
        ByVal fiberCount As Long, ByVal rebarCount As Long)
    TestAppendLine stats, "BASELINE|" & caseName & _
        "|converged=" & TestBoolText(solver.Converged) & _
        "|eps0=" & TestFormatNumberInvariant(solver.Epsilon0) & _
        "|kappaX=" & TestFormatNumberInvariant(solver.KappaX) & _
        "|kappaY=" & TestFormatNumberInvariant(solver.KappaY) & _
        "|Nint=" & TestFormatNumberInvariant(solver.Nint) & _
        "|Mxint=" & TestFormatNumberInvariant(solver.Mxint) & _
        "|Myint=" & TestFormatNumberInvariant(solver.Myint) & _
        "|resN=" & TestFormatNumberInvariant(solver.ResidualN) & _
        "|resMx=" & TestFormatNumberInvariant(solver.ResidualMx) & _
        "|resMy=" & TestFormatNumberInvariant(solver.ResidualMy) & _
        "|lambda=" & lambdaText & _
        "|iterations=" & CStr(solver.Iterations) & _
        "|elapsedSec=" & TestFormatNumberInvariant(elapsedSeconds) & _
        "|fibers=" & CStr(fiberCount) & _
        "|rebars=" & CStr(rebarCount)
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Однотипные helpers протокола задают устойчивые Boolean/числовые подписи и перевод строки.
Private Function TestBoolText(ByVal value As Boolean) As String
    If value Then
        TestBoolText = "True"
    Else
        TestBoolText = "False"
    End If
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Добавляет строку к контрольному протоколу с явным разделителем строк.
Private Sub TestAppendLine(ByRef stats As TRegressionStats, ByVal text As String)
    stats.Report = stats.Report & text & vbCrLf
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Форматирует число для текстового протокола с десятичной точкой; исходные расчётные значения не меняются.
Private Function TestFormatNumberInvariant(ByVal value As Double) As String
    TestFormatNumberInvariant = Replace$(Format$(value, "0.############"), ",", ".")
End Function


