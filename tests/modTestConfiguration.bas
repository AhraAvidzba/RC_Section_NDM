Attribute VB_Name = "modTestConfiguration"
Option Explicit

' ==========================================================================
' ДЛЯ ТЕСТОВ: поведение расчетных настроек через настоящий Config
' ==========================================================================
' Модуль меняет ячейки изолированной книги, читает их штатным reader-ом и
' проверяет реальные результаты Capacity/Formation/Stability, а не только getters.
' Отдельный входной набор проверяет структуру именованных таблиц Config,
' диагностические адреса и ошибки ключей без запуска расчетного ядра.
' Нагрузки и геометрия заданы во внутренних Н, Н*мм и мм; параметры материалов,
' профили и solve-options проходят через Config. После набора все формулы
' входных таблиц восстанавливаются, а результаты остаются для проверки writer.

Private Type TConfigTestStats
    Passed As Long
    Failed As Long
    Report As String
End Type

Private Const FIXTURE_N As Double = 130000# ' Положительный внутренний знак - растяжение.
Private Const FIXTURE_MX As Double = 5200000# ' Оба момента активны для невырожденного Search.
Private Const FIXTURE_MY As Double = 2600000#
Private Const FORCE_TOLERANCE As Double = 0.01 ' Точный solve для независимого oracle предельной деформации.
Private Const MOMENT_TOLERANCE As Double = 1#

' ДЛЯ ТЕСТОВ: проверяет все стратегии, одномерные методы и пути Formation.
' Возвращает отчет и счетчики для отдельного runner-а и полного batch-suite.
Public Function RunAudit03SearchConfigTests(Optional ByRef passed As Long = 0, _
        Optional ByRef failed As Long = 0) As String
    Dim stats As TConfigTestStats, systemRange As Object, unitRange As Object, profileRange As Object
    Dim savedSystem As Variant, savedUnits As Variant, savedProfiles As Variant
    Dim section As CSectionModel, provider As CMaterialModelProvider, profiles As CCalculationProfileCatalog
    On Error GoTo FailedRun
    Set systemRange = ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange
    Set unitRange = ThisWorkbook.Names.Item("rngUnitSettings").RefersToRange
    Set profileRange = ThisWorkbook.Names.Item("rngCalculationProfiles").RefersToRange
    savedSystem = systemRange.Formula: savedUnits = unitRange.Formula: savedProfiles = profileRange.Formula
    ConfigureSearchFixture systemRange, unitRange, profileRange
    Dim settings As CSystemSettingsReader, units As CUnitSystem
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    Set provider = New CMaterialModelProvider: provider.Initialize settings, units
    Set profiles = New CCalculationProfileCatalog: profiles.LoadFromWorkbook ThisWorkbook
    Set section = SearchFixtureSection()
    CheckCapacitySelectors stats, systemRange, section, provider, profiles, units
    CheckFormationSelectors stats, systemRange, section, provider, profiles, units
    CheckSelectorIsolation stats, systemRange, section, provider, profiles, units
    CheckInvalidSelectors stats, systemRange, section, provider, profiles, units
    GoTo Restore
FailedRun:
    stats.Failed = stats.Failed + 1
    LogLine stats, "FAIL: searchConfig.runtime; " & CStr(Err.Number) & "; " & Err.Description
Restore:
    On Error Resume Next
    If Not systemRange Is Nothing Then systemRange.Formula = savedSystem
    If Not unitRange Is Nothing Then unitRange.Formula = savedUnits
    If Not profileRange Is Nothing Then profileRange.Formula = savedProfiles
    On Error GoTo 0
    LogLine stats, "TOTAL_AUDIT03_SEARCH_CONFIG: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    passed = stats.Passed: failed = stats.Failed
    RunAudit03SearchConfigTests = stats.Report
End Function

' Создает управляемые единицы и профили, чтобы сила/момент и устойчивость
' не меняли независимую проверку выбранной lambda-траектории.
Private Sub ConfigureSearchFixture(ByVal systemRange As Object, ByVal unitRange As Object, ByVal profileRange As Object)
    Dim row As Long, value As String
    For row = 2 To unitRange.Rows.Count
        Select Case CStr(unitRange.Cells(row, 1).Value2)
            Case "Length": value = "mm"
            Case "Area": value = "mm2"
            Case "Force": value = "N"
            Case "Moment": value = "N*mm"
            Case "Stress": value = "MPa"
            Case "Curvature": value = "1/mm"
            Case Else: Err.Raise vbObjectError + 4499, "ConfigureSearchFixture", "Неизвестная строка единиц."
        End Select
        unitRange.Cells(row, 2).Value2 = value: unitRange.Cells(row, 4).Value2 = value
    Next row
    SetValue systemRange, "Calculation.ZeroMomentPerDepth", 0#
    SetValue systemRange, "General.ExecutionReportEnabled", "Yes"
    SetValue systemRange, "Solver.ToleranceN", FORCE_TOLERANCE
    SetValue systemRange, "Solver.ToleranceMx", MOMENT_TOLERANCE
    SetValue systemRange, "Solver.ToleranceMy", MOMENT_TOLERANCE
    SetValue systemRange, "Capacity.InitialLambda", 1#
    SetValue systemRange, "Capacity.MaxLambda", 1024#
    SetValue systemRange, "Capacity.ToleranceLambda", 0.000001
    SetValue systemRange, "Capacity.ToleranceStrain", 0.0001
    SetValue systemRange, "Capacity.MaxRetries", 5
    SetValue systemRange, "Capacity.BaseLoadSteps", 1
    SetValue systemRange, "Capacity.SolverMaxIterations", 60
    SetValue systemRange, "Capacity.SolutionStrategy", "Auto"
    SetValue systemRange, "Capacity.SearchMethod", "Bisection"
    SetValue systemRange, "SLS.Crack.InitiationSolutionStrategy", "Auto"
    SetValue systemRange, "SLS.Crack.Allowable", 1#
    SetValue systemRange, "SLS.Crack.Code", "SP63"
    SetValue systemRange, "SLS.Crack.PsiMode", "User"
    SetValue systemRange, "SLS.Crack.PsiS", 1#
    SetProfileValue profileRange, "Calculation.Stability.Enabled", "No", "PR1"
    SetProfileValue profileRange, "Calculation.Stability.Enabled", "No", "PR2"
    SetProfileValue profileRange, "Calculation.Strength.DirectState", "Yes", "PR1"
    SetProfileValue profileRange, "Calculation.Strength.Capacity", "Yes", "PR1"
    SetProfileValue profileRange, "Calculation.Crack.Width", "No", "PR1"
    SetProfileValue profileRange, "Calculation.Strength.DirectState", "No", "PR2"
    SetProfileValue profileRange, "Calculation.Strength.Capacity", "No", "PR2"
    SetProfileValue profileRange, "Calculation.Crack.Width", "Yes", "PR2"
End Sub

' Меняет только селектор другого раздела на допустимые варианты. Неактивный
' потребитель не должен менять физическую точку, метод или число solve
' активного раздела; validation допустимости при этом остается обязательной.
Private Sub CheckSelectorIsolation(ByRef stats As TConfigTestStats, ByVal systemRange As Object, _
        ByVal section As CSectionModel, ByVal provider As CMaterialModelProvider, _
        ByVal profiles As CCalculationProfileCatalog, ByVal units As CUnitSystem)
    Dim capacityBase As CBatchSectionCalculator, formationBase As CBatchSectionCalculator
    Dim changed As CBatchSectionCalculator, key As Variant, value As Variant, values As Variant
    Dim saved As Variant, profileId As String, prefix As String, expectedLambda As Double
    saved = systemRange.Formula
    Set capacityBase = RunConfigBatch(stats, "searchConfig.isolation.capacityBase", "PR1", section, provider, profiles, units)
    Set formationBase = RunConfigBatch(stats, "searchConfig.isolation.formationBase", "PR2", section, provider, profiles, units)
    For Each key In Array("Capacity.SolutionStrategy", "Capacity.SearchMethod", "SLS.Crack.InitiationSolutionStrategy")
        Select Case CStr(key)
            Case "Capacity.SearchMethod": values = Array("Bisection", "Brent", "Secant")
            Case Else: values = Array("Auto", "UltimateStrain", "LoadMultiplier")
        End Select
        profileId = "PR1"
        If Left$(CStr(key), 9) = "Capacity." Then profileId = "PR2"
        For Each value In values
            systemRange.Formula = saved
            SetValue systemRange, CStr(key), CStr(value)
            prefix = "searchConfig.isolation." & CStr(key) & "." & CStr(value)
            Set changed = RunConfigBatch(stats, prefix, profileId, section, provider, profiles, units)
            If profileId = "PR1" Then
                expectedLambda = capacityBase.ResultAt(1).StrengthResult.Capacity.LambdaCapacity
                CheckClose stats, prefix & ".lambda", changed.ResultAt(1).StrengthResult.Capacity.LambdaCapacity, expectedLambda, 0.0000000001
                Check stats, prefix & ".method", changed.ResultAt(1).StrengthResult.Capacity.SolutionMethod = _
                    capacityBase.ResultAt(1).StrengthResult.Capacity.SolutionMethod
                Check stats, prefix & ".solveCount", changed.SolverCallCount = capacityBase.SolverCallCount
            Else
                expectedLambda = formationBase.ResultAt(1).CrackResult.Formation.LambdaCrc
                CheckClose stats, prefix & ".lambda", changed.ResultAt(1).CrackResult.Formation.LambdaCrc, expectedLambda, 0.0000000001
                Check stats, prefix & ".path", changed.ResultAt(1).CrackResult.Formation.FormationMethod = _
                    formationBase.ResultAt(1).CrackResult.Formation.FormationMethod
                Check stats, prefix & ".solveCount", changed.SolverCallCount = formationBase.SolverCallCount
            End If
        Next value
    Next key
    systemRange.Formula = saved
End Sub

' Собирает симметричный армированный круг общим builder-ом. Его физика
' обеспечивает достижимую точку для каждого активного пути без AutoCAD/Excel IO.
Private Function SearchFixtureSection() As CSectionModel
    Dim geometry As CGeometryCircle, mesh As CFiberMeshBuilder, rebars As CRebarLayout
    Set geometry = New CGeometryCircle: geometry.InitializeByDiameter 300#
    Set mesh = New CFiberMeshBuilder: mesh.BuildMesh geometry, 20#, 20#, 1, 2
    Set rebars = New CRebarLayout
    Dim i As Long, angle As Double
    For i = 0 To 11
        angle = 2# * GEOM_PI * CDbl(i) / 12#
        rebars.AddBar "R" & CStr(i + 1), 120# * Cos(angle), 120# * Sin(angle), 16#, 0#, "A400", vbNullString, geometry
    Next i
    Set SearchFixtureSection = BuildGeneratedSectionModel(mesh, rebars, "SearchConfigFixture")
End Function

' Проверяет фактически выполненный метод и конечную физическую точку всех
' девяти пар strategy/method; при UltimateStrain одномерный метод не активен.
Private Sub CheckCapacitySelectors(ByRef stats As TConfigTestStats, ByVal systemRange As Object, _
        ByVal section As CSectionModel, ByVal provider As CMaterialModelProvider, ByVal profiles As CCalculationProfileCatalog, ByVal units As CUnitSystem)
    Dim strategy As Variant, method As Variant, batch As CBatchSectionCalculator, cap As CCapacityResult
    Dim prefix As String, baselineUltimate As Double, point As CSectionStateResult, ratio As Double
    For Each strategy In Array("Auto", "UltimateStrain", "LoadMultiplier")
        baselineUltimate = 0#
        For Each method In Array("Bisection", "Brent", "Secant")
            SetValue systemRange, "Capacity.SolutionStrategy", CStr(strategy)
            SetValue systemRange, "Capacity.SearchMethod", CStr(method)
            prefix = "searchConfig.capacity." & CStr(strategy) & "." & CStr(method)
            Set batch = RunConfigBatch(stats, prefix, "PR1", section, provider, profiles, units)
            Set cap = batch.ResultAt(1).StrengthResult.Capacity
            Check stats, prefix & ".searchExists", Not cap.SearchResult Is Nothing
            If cap.SearchResult Is Nothing Then GoTo NextMethod
            Check stats, prefix & ".requested", cap.SearchResult.RequestedStrategy = CStr(strategy)
            If CStr(strategy) = "Auto" Then
                Check stats, prefix & ".actual", cap.SolutionMethod = "UltimateStrain" Or cap.SolutionMethod = "LoadMultiplier"
            Else
                Check stats, prefix & ".actual", cap.SolutionMethod = CStr(strategy)
            End If
            Check stats, prefix & ".point", cap.SearchResult.HasLimitPoint And cap.LambdaCapacity > 0#
            Set point = cap.StateResult
            Check stats, prefix & ".stateExists", Not point Is Nothing
            If Not point Is Nothing Then
                CheckState stats, prefix, point
                CheckClose stats, prefix & ".pathN", point.Nint, FIXTURE_N * cap.LambdaCapacity, FORCE_TOLERANCE
                CheckClose stats, prefix & ".pathMx", point.Mxint, FIXTURE_MX * cap.LambdaCapacity, MOMENT_TOLERANCE
                CheckClose stats, prefix & ".pathMy", point.Myint, FIXTURE_MY * cap.LambdaCapacity, MOMENT_TOLERANCE
                ratio = LargerValue(Abs(point.MinConcreteStrain / provider.ConcreteCompressionLimitFromSpec(point.MaterialSpec)), _
                    LargerValue(Abs(point.MinSteelStrain / provider.SteelCompressionLimitFromSpec(point.MaterialSpec)), _
                    Abs(point.MaxSteelStrain / provider.SteelTensionLimitFromSpec(point.MaterialSpec))))
                CheckClose stats, prefix & ".physicalCriterion", ratio, 1#, 0.0002
            End If
            If cap.SolutionMethod = "LoadMultiplier" Then
                Check stats, prefix & ".activeSearchMethod", InStr(1, cap.DiagnosticLog, "searchMethod=" & CStr(method), vbBinaryCompare) > 0
            ElseIf CStr(strategy) = "UltimateStrain" Then
                If baselineUltimate = 0# Then baselineUltimate = cap.LambdaCapacity
                CheckClose stats, prefix & ".inactiveSearchMethod", cap.LambdaCapacity, baselineUltimate, 0.0000000001
                Check stats, prefix & ".noLoadMultiplier", InStr(1, cap.DiagnosticLog, "searchMethod=", vbBinaryCompare) = 0
            End If
            LogLine stats, "SEARCH_CONFIG: " & prefix & "|lambda=" & FormatNumberInvariant(cap.LambdaCapacity) & _
                "|method=" & cap.SolutionMethod & "|status=" & cap.Status & "|ratio=" & FormatNumberInvariant(ratio)
NextMethod:
        Next method
    Next strategy
    ' Отдельно сохраняем точный negative-контрпример Brent с менее строгим
    ' lambda-допуском. Его физическая пригодность не подменяется близостью
    ' критерия к единице или дополнительным уменьшением допуска в fixture.
    SetValue systemRange, "Capacity.ToleranceLambda", 0.0001
    SetValue systemRange, "Capacity.SolutionStrategy", "LoadMultiplier"
    SetValue systemRange, "Capacity.SearchMethod", "Brent"
    SetValue systemRange, "Solver.ToleranceN", 5#
    SetValue systemRange, "Solver.ToleranceMx", 5000#
    SetValue systemRange, "Solver.ToleranceMy", 5000#
    Set batch = RunConfigBatch(stats, "searchConfig.brentPhysicalEndpoint", "PR1", section, provider, profiles, units, 10# / 13#)
    Set cap = batch.ResultAt(1).StrengthResult.Capacity
    Check stats, "searchConfig.brentPhysicalEndpoint.point", cap.SearchResult.HasLimitPoint
    CheckState stats, "searchConfig.brentPhysicalEndpoint", cap.StateResult, 5#, 5000#
    CheckClose stats, "searchConfig.brentPhysicalEndpoint.pathN", cap.StateResult.Nint, 100000# * cap.LambdaCapacity, 5#
    CheckClose stats, "searchConfig.brentPhysicalEndpoint.pathMx", cap.StateResult.Mxint, 4000000# * cap.LambdaCapacity, 5000#
    CheckClose stats, "searchConfig.brentPhysicalEndpoint.pathMy", cap.StateResult.Myint, 2000000# * cap.LambdaCapacity, 5000#
    SetValue systemRange, "Capacity.ToleranceLambda", 0.000001
    SetValue systemRange, "Solver.ToleranceN", FORCE_TOLERANCE
    SetValue systemRange, "Solver.ToleranceMx", MOMENT_TOLERANCE
    SetValue systemRange, "Solver.ToleranceMy", MOMENT_TOLERANCE
    SetValue systemRange, "Capacity.SolutionStrategy", "Auto"
    SetValue systemRange, "Capacity.SearchMethod", "Bisection"
End Sub

' Проверяет все пары Formation strategy/path и независимо сравнивает
' компоненты физической точки с выбранной постоянной и масштабируемой частью.
Private Sub CheckFormationSelectors(ByRef stats As TConfigTestStats, ByVal systemRange As Object, _
        ByVal section As CSectionModel, ByVal provider As CMaterialModelProvider, ByVal profiles As CCalculationProfileCatalog, ByVal units As CUnitSystem)
    Dim strategy As Variant, path As Variant, batch As CBatchSectionCalculator, formation As CCrackFormationResult
    Dim prefix As String, actualPath As String, point As CSectionStateResult, nExpected As Double, mxExpected As Double, myExpected As Double
    Dim loadFactor As Double, strainLimit As Double
    For Each path In Array("Auto", "Mx", "My", "Mxy", "N", "NMxy")
        loadFactor = 1#
        If CStr(path) = "Mx" Or CStr(path) = "My" Then loadFactor = 0.5
        For Each strategy In Array("Auto", "UltimateStrain", "LoadMultiplier")
            SetValue systemRange, "SLS.Crack.InitiationSolutionStrategy", CStr(strategy)
            prefix = "searchConfig.formation." & CStr(path) & "." & CStr(strategy)
            Set batch = RunConfigBatch(stats, prefix, "PR2", section, provider, profiles, units, loadFactor, LambdaPath(CStr(path)))
            Set formation = batch.ResultAt(1).CrackResult.Formation
            Check stats, prefix & ".point", formation.HasLimitPoint And formation.LambdaCrc > 0#
            Check stats, prefix & ".searchExists", Not formation.SearchResult Is Nothing
            If Not formation.SearchResult Is Nothing Then
                Check stats, prefix & ".requested", formation.SearchResult.RequestedStrategy = CStr(strategy)
                If CStr(strategy) <> "Auto" Then Check stats, prefix & ".actual", formation.SearchResult.ActualMethod = CStr(strategy)
            End If
            actualPath = formation.FormationMethod
            If CStr(path) = "Auto" Then
                Check stats, prefix & ".autoFirstPath", actualPath = LambdaPath("Mxy")
            Else
                Check stats, prefix & ".fixedPath", actualPath = LambdaPath(CStr(path))
            End If
            Set point = formation.PreCrackState
            Check stats, prefix & ".preExists", Not point Is Nothing
            If Not point Is Nothing Then
                CheckState stats, prefix & ".pre", point
                strainLimit = provider.ConcreteTensionLimitFromSpec(point.MaterialSpec)
                If point.MinConcreteStrain > 0# And point.MaxConcreteStrain > 0# Then _
                    strainLimit = strainLimit - (strainLimit - provider.ConcreteTensionEbt0FromSpec(point.MaterialSpec)) * _
                        point.MinConcreteStrain / point.MaxConcreteStrain
                CheckClose stats, prefix & ".tensionCriterion", point.MaxConcreteStrain, strainLimit, 0.00000001
                nExpected = FIXTURE_N * loadFactor: mxExpected = FIXTURE_MX * loadFactor: myExpected = FIXTURE_MY * loadFactor
                If actualPath = LambdaPath("N") Or actualPath = LambdaPath("NMxy") Then nExpected = nExpected * formation.LambdaCrc
                If actualPath = LambdaPath("Mx") Or actualPath = LambdaPath("Mxy") Or actualPath = LambdaPath("NMxy") Then _
                    mxExpected = mxExpected * formation.LambdaCrc
                If actualPath = LambdaPath("My") Or actualPath = LambdaPath("Mxy") Or actualPath = LambdaPath("NMxy") Then _
                    myExpected = myExpected * formation.LambdaCrc
                CheckClose stats, prefix & ".pathN", point.Nint, nExpected, FORCE_TOLERANCE
                CheckClose stats, prefix & ".pathMx", point.Mxint, mxExpected, MOMENT_TOLERANCE
                CheckClose stats, prefix & ".pathMy", point.Myint, myExpected, MOMENT_TOLERANCE
            End If
            If formation.CrackFormed Then
                Check stats, prefix & ".postExists", Not formation.PostCrackState Is Nothing
            Else
                Check stats, prefix & ".noPostAboveCurrent", formation.PostCrackState Is Nothing And formation.LambdaCrc > 1#
            End If
            If Not formation.PostCrackState Is Nothing Then CheckState stats, prefix & ".post", formation.PostCrackState
            LogLine stats, "SEARCH_CONFIG: " & prefix & "|lambda=" & FormatNumberInvariant(formation.LambdaCrc) & _
                "|path=" & actualPath & "|code=" & ResultCodeToText(formation.ResultMeta.ResultCode)
            If Not formation.HasLimitPoint Then LogLine stats, "SEARCH_DIAGNOSTIC: " & prefix & vbCrLf & formation.DiagnosticLog
        Next strategy
    Next path
    SetValue systemRange, "SLS.Crack.InitiationSolutionStrategy", "Auto"
End Sub

' Отсутствующая строка, пустая ячейка, TODO и неизвестный выбор должны
' давать InputErr с местом и действием, без solve и без подстановки дефолта.
' После возврата каждого ключа выполняется настоящий успешный расчет.
Private Sub CheckInvalidSelectors(ByRef stats As TConfigTestStats, ByVal systemRange As Object, _
        ByVal section As CSectionModel, ByVal provider As CMaterialModelProvider, ByVal profiles As CCalculationProfileCatalog, ByVal units As CUnitSystem)
    Dim saved As Variant, key As Variant, bad As Variant, cell As Object, batch As CBatchSectionCalculator
    Dim prefix As String, profileId As String, message As String, values As Variant
    saved = systemRange.Formula
    values = Array(vbNullString, "TODO", "Unknown", CVErr(2042), "MISSING_ROW")
    For Each key In Array("Capacity.SolutionStrategy", "Capacity.SearchMethod", "SLS.Crack.InitiationSolutionStrategy")
        profileId = "PR2"
        If Left$(CStr(key), 8) = "Capacity" Then profileId = "PR1"
        For Each bad In values
            systemRange.Formula = saved
            SetValue systemRange, "Capacity.SolutionStrategy", "LoadMultiplier"
            Set cell = ValueCell(systemRange, CStr(key), 2)
            prefix = "searchConfig.invalid." & CStr(key)
            If IsError(bad) Then
                prefix = prefix & ".excelError": cell.Value2 = bad
            ElseIf CStr(bad) = "MISSING_ROW" Then
                prefix = prefix & ".missing": cell.Offset(0, -1).Value2 = "Audit03.RemovedSelector"
            Else
                prefix = prefix & "." & CStr(bad): cell.Value2 = bad
            End If
            Set batch = RunConfigBatch(stats, prefix, profileId, section, provider, profiles, units)
            message = batch.ResultAt(1).OverallMeta.ResultComment
            Check stats, prefix & ".inputErr", batch.ResultAt(1).Status = "InputErr"
            Check stats, prefix & ".noSolve", batch.SolverCallCount = 0
            Check stats, prefix & ".key", InStr(1, message, CStr(key), vbBinaryCompare) > 0
            Check stats, prefix & ".location", InStr(1, message, "Config", vbBinaryCompare) > 0
            If Not IsError(bad) Then
                If CStr(bad) = "MISSING_ROW" Then
                    Check stats, prefix & ".section", InStr(1, message, "раздел", vbTextCompare) > 0
                Else
                    Check stats, prefix & ".cell", InStr(1, message, cell.Address(False, False), vbTextCompare) > 0
                End If
            Else
                Check stats, prefix & ".cell", InStr(1, message, cell.Address(False, False), vbTextCompare) > 0
            End If
            Check stats, prefix & ".action", InStr(1, message, "Введите", vbTextCompare) > 0 Or _
                InStr(1, message, "Выберите", vbTextCompare) > 0 Or InStr(1, message, "Восстановите", vbTextCompare) > 0
            LogLine stats, "INPUT_MESSAGE: " & prefix & "|" & message
        Next bad
        systemRange.Formula = saved
        Set batch = RunConfigBatch(stats, "searchConfig.recovery." & CStr(key), profileId, section, provider, profiles, units)
        Check stats, "searchConfig.recovery." & CStr(key), batch.ResultAt(1).Status <> "InputErr" And batch.SolverCallCount > 0
    Next key
    systemRange.Formula = saved
End Sub

' Создает новый batch на той же физической модели и читает текущие ячейки.
' Reader failure сохраняется как ошибка ввода, не как численная несходимость.
Private Function RunConfigBatch(ByRef stats As TConfigTestStats, ByVal prefix As String, ByVal profileId As String, _
        ByVal section As CSectionModel, ByVal provider As CMaterialModelProvider, ByVal profiles As CCalculationProfileCatalog, _
        ByVal units As CUnitSystem, Optional ByVal loadFactor As Double = 1#, _
        Optional ByVal loadPathText As String = "LambdaNMxy") As CBatchSectionCalculator
    Dim batch As CBatchSectionCalculator, settings As CSystemSettingsReader, reason As String, report As CExecutionReport
    Set batch = New CBatchSectionCalculator: batch.Initialize section, provider
    Set batch.ProfileCatalog = profiles
    Set settings = New CSystemSettingsReader
    LogLine stats, "RUN: " & prefix
    On Error GoTo InvalidRead
    settings.LoadFromWorkbook ThisWorkbook
    Set report = New CExecutionReport: report.Initialize ThisWorkbook, settings
    Set batch.ExecutionReport = report
    batch.ApplySettings settings, units
    batch.AddCombination prefix, FIXTURE_N * loadFactor, FIXTURE_MX * loadFactor, FIXTURE_MY * loadFactor, _
        profileId, "Проверка выбора Config", loadPathText
    batch.Execute
    GoTo Publish
InvalidRead:
    reason = Err.Description
    On Error GoTo 0
    batch.AddInvalidCombination prefix, profileId, "Ошибочный Config", reason
    batch.Execute
Publish:
    Dim passed As Long, failed As Long
    stats.Report = stats.Report & Audit03ValidateConfigBatchResults(batch, units, passed, failed)
    stats.Passed = stats.Passed + passed: stats.Failed = stats.Failed + failed
    Set RunConfigBatch = batch
End Function

' ДЛЯ ТЕСТОВ: проверяет единый порядок Auto и независимую арифметику пути,
' затем все конкретные пути через реальные профили, Search и result writers.
' Это направленная приемка механизма, не замена полной нагрузочной матрицы.
Public Function RunAudit03UniversalLoadPathTests(Optional ByRef passed As Long = 0, _
        Optional ByRef failed As Long = 0) As String
    Dim stats As TConfigTestStats, systemRange As Object, unitRange As Object, profileRange As Object
    Dim savedSystem As Variant, savedUnits As Variant, savedProfiles As Variant
    Dim settings As CSystemSettingsReader, units As CUnitSystem, provider As CMaterialModelProvider
    Dim profiles As CCalculationProfileCatalog, section As CSectionModel
    On Error GoTo FailedRun
    CheckUniversalPathMath stats
    Set systemRange = ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange
    Set unitRange = ThisWorkbook.Names.Item("rngUnitSettings").RefersToRange
    Set profileRange = ThisWorkbook.Names.Item("rngCalculationProfiles").RefersToRange
    savedSystem = systemRange.Formula: savedUnits = unitRange.Formula: savedProfiles = profileRange.Formula
    ConfigureSearchFixture systemRange, unitRange, profileRange
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    Set provider = New CMaterialModelProvider: provider.Initialize settings, units
    Set profiles = New CCalculationProfileCatalog: profiles.LoadFromWorkbook ThisWorkbook
    Set section = SearchFixtureSection()
    CheckUniversalPathReader stats, section, provider, profiles, units
    Dim path As Variant, profileId As Variant, batch As CBatchSectionCalculator
    Dim result As CCombinationResult, point As CSectionStateResult, descriptor As CLoadPathDescriptor
    Dim prefix As String, actualPath As String, lambdaValue As Double
    Dim nExpected As Double, mxExpected As Double, myExpected As Double, load As CSectionLoadState
    Set load = New CSectionLoadState: load.Initialize FIXTURE_N * 0.5, FIXTURE_MX * 0.5, FIXTURE_MY * 0.5, 0#, 0#
    For Each path In Array("LambdaMx", "LambdaMy", "LambdaMxy", "LambdaN", "LambdaNMxy", "Auto", vbNullString)
        For Each profileId In Array("PR1", "PR2")
            prefix = "universalPath." & CStr(profileId) & "." & CStr(path)
            Set batch = RunConfigBatch(stats, prefix, CStr(profileId), section, provider, profiles, units, 0.5, CStr(path))
            Set result = batch.ResultAt(1)
            Check stats, prefix & ".noInputError", result.Status <> "InputErr" And result.Status <> "CalcErr"
            Set point = Nothing
            If CStr(profileId) = "PR1" Then
                Check stats, prefix & ".point", result.StrengthResult.Capacity.SearchResult.HasLimitPoint
                actualPath = result.StrengthResult.Capacity.PathResolved
                lambdaValue = result.StrengthResult.Capacity.LambdaCapacity
                Set point = result.StrengthResult.Capacity.StateResult
            Else
                Check stats, prefix & ".noNumericalFailure", result.CrackResult.Formation.ResultMeta.InternalStatus <> rsNumericalFailure
                Check stats, prefix & ".point", result.CrackResult.Formation.HasLimitPoint
                actualPath = result.CrackResult.Formation.FormationMethod
                lambdaValue = result.CrackResult.Formation.LambdaCrc
                Set point = result.CrackResult.Formation.PreCrackState
            End If
            Set descriptor = New CLoadPathDescriptor
            If CStr(path) = "Auto" Or Len(CStr(path)) = 0 Then
                Check stats, prefix & ".autoFirst", actualPath = "LambdaMxy" Or actualPath = LambdaPath("Mxy")
            Else
                Check stats, prefix & ".fixed", descriptor.NormalizeKey(actualPath) = CStr(path)
            End If
            If Not point Is Nothing Then
                CheckState stats, prefix & ".state", point
                descriptor.InitializeFromLoadState actualPath, load
                descriptor.Target lambdaValue, nExpected, mxExpected, myExpected
                CheckClose stats, prefix & ".N", point.Nint, nExpected, FORCE_TOLERANCE
                CheckClose stats, prefix & ".Mx", point.Mxint, mxExpected, MOMENT_TOLERANCE
                CheckClose stats, prefix & ".My", point.Myint, myExpected, MOMENT_TOLERANCE
            End If
            LogLine stats, "UNIVERSAL_PATH: " & prefix & "|actual=" & actualPath & "|lambda=" & FormatNumberInvariant(lambdaValue)
        Next profileId
    Next path
    Set batch = RunConfigBatch(stats, "universalPath.constantAlreadyCracked", "PR2", section, provider, profiles, units, 1#, "LambdaMy")
    Set result = batch.ResultAt(1)
    Check stats, "universalPath.constantAlreadyCracked.typed", _
        result.CrackResult.Formation.ResultMeta.InternalStatus = rsCheckFailed And _
        result.CrackResult.Formation.ResultMeta.ResultCode = rcInitialStateBeyondLimit
    Dim policy As CResultStatusPolicy
    Set policy = New CResultStatusPolicy
    Check stats, "universalPath.constantAlreadyCracked.display", policy.ExternalStatus(result.CrackResult.Formation.ResultMeta) = "BaseFail"
    Check stats, "universalPath.constantAlreadyCracked.noPoint", Not result.CrackResult.Formation.HasLimitPoint
    Check stats, "universalPath.constantAlreadyCracked.widthCalculated", result.CrackWidthMeta.Calculated
    Check stats, "universalPath.constantAlreadyCracked.widthWarning", result.CrackWidthMeta.InternalStatus = rsSuccessWithWarning
    CheckClose stats, "universalPath.constantAlreadyCracked.psi1", result.CrackResult.Width.PsiS, 1#, 0#
    Check stats, "universalPath.constantAlreadyCracked.warningComment", InStr(1, result.CrackWidthMeta.ResultComment, "Предупреждение", vbTextCompare) > 0
    Check stats, "universalPath.constantAlreadyCracked.longitudinalCompleted", _
        result.CrackResult.Longitudinal.ResultMeta.Calculated Or _
        result.CrackResult.Longitudinal.ResultMeta.InternalStatus = rsNotApplicable
    Check stats, "universalPath.constantAlreadyCracked.longitudinalIndependent", _
        result.CrackResult.Longitudinal.ResultMeta.InternalStatus <> rsBlockedByDependency
    Dim writer As CBatchResultWriter, detailedComment As String, summaryComment As String
    Set writer = New CBatchResultWriter
    writer.WriteSummary ThisWorkbook, batch, units
    detailedComment = CStr(ThisWorkbook.Names.Item("rngCrackSummaryAnchor").RefersToRange.Offset(0, 1).Value2)
    summaryComment = CStr(ThisWorkbook.Names.Item("rngBatchSummary").RefersToRange.Offset(12, 2).Value2)
    Check stats, "universalPath.constantAlreadyCracked.detailComment", detailedComment = result.CrackResult.ResultMeta.ResultComment
    Check stats, "universalPath.constantAlreadyCracked.summaryComment", summaryComment = result.OverallMeta.ResultComment
    Check stats, "universalPath.constantAlreadyCracked.detailWarning", InStr(1, detailedComment, "Предупреждение", vbTextCompare) > 0
    Check stats, "universalPath.constantAlreadyCracked.summaryWarning", InStr(1, summaryComment, "Предупреждение", vbTextCompare) > 0
    Dim crackAnchor As Object
    Set crackAnchor = ThisWorkbook.Names.Item("rngCrackSummaryAnchor").RefersToRange
    CheckClose stats, "universalPath.constantAlreadyCracked.outputPsi1", CDbl(crackAnchor.Offset(0, 57).Value2), 1#, 0#
    CheckClose stats, "universalPath.constantAlreadyCracked.outputWidth", CDbl(crackAnchor.Offset(0, 63).Value2), _
        units.InternalLengthToOutput(result.CrackResult.Width.CrackWidth), 0.000000000001
    CheckClose stats, "universalPath.constantAlreadyCracked.outputEs", CDbl(crackAnchor.Offset(0, 62).Value2), _
        units.InternalStressToOutput(result.CrackResult.Width.SteelEs), 0.000000000001
    Check stats, "universalPath.constantAlreadyCracked.outputWidthStatus", CStr(crackAnchor.Offset(0, 66).Value2) = _
        policy.ExternalStatus(result.CrackWidthMeta)
    GoTo Restore
FailedRun:
    stats.Failed = stats.Failed + 1
    LogLine stats, "FAIL: universalPath.runtime; " & CStr(Err.Number) & "; " & Err.Description
Restore:
    On Error Resume Next
    If Not systemRange Is Nothing Then systemRange.Formula = savedSystem
    If Not unitRange Is Nothing Then unitRange.Formula = savedUnits
    If Not profileRange Is Nothing Then profileRange.Formula = savedProfiles
    On Error GoTo 0
    LogLine stats, "TOTAL_AUDIT03_UNIVERSAL_PATH: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    passed = stats.Passed: failed = stats.Failed
    RunAudit03UniversalLoadPathTests = stats.Report
End Function

' ДЛЯ ТЕСТОВ: реальные ячейки LoadPath проверяются штатным reader-ом.
' Пустое значение и формула с пустым результатом выбирают Auto; неизвестный
' путь и ошибка Excel сохраняют адресную InputErr-строку. Последующий
' корректный LC обязан рассчитываться. Временный лист удаляется; количество
' пользовательских строк rngLoadCombinations и их значения не меняются.
Private Sub CheckUniversalPathReader(ByRef stats As TConfigTestStats, ByVal section As CSectionModel, _
        ByVal provider As CMaterialModelProvider, ByVal profiles As CCalculationProfileCatalog, ByVal units As CUnitSystem)
    Dim sheet As Object, source As Object, data() As Variant, previousAlerts As Boolean
    Dim batch As CBatchSectionCalculator, settings As CSystemSettingsReader, reader As CLoadCombinationReader
    Dim profileId As Variant, item As Variant, row As Long, i As Long, prefix As String, message As String
    Dim passed As Long, failed As Long, failureReason As String
    previousAlerts = Application.DisplayAlerts
    On Error GoTo FailedRun
    Set sheet = ThisWorkbook.Worksheets.Add
    Set source = sheet.Range("A1:G11")
    ReDim data(1 To 11, 1 To 7)
    data(1, 1) = "CombinationID": data(1, 2) = "N": data(1, 3) = "Mx": data(1, 4) = "My"
    data(1, 5) = "ProfileId": data(1, 6) = "LoadPath": data(1, 7) = "Comment"
    row = 1
    For Each profileId In Array("PR1", "PR2")
        For Each item In Array("blank", "formulaEmpty", "unknown", "excelError", "recovery")
            row = row + 1
            data(row, 1) = "PATH_" & CStr(profileId) & "_" & CStr(item)
            data(row, 2) = FIXTURE_N * 0.5 / units.InputForceToInternal(1#)
            data(row, 3) = FIXTURE_MX * 0.5 / units.InputMomentMxToInternal(1#)
            data(row, 4) = FIXTURE_MY * 0.5 / units.InputMomentMyToInternal(1#)
            data(row, 5) = CStr(profileId): data(row, 7) = "Проверка ячейки универсального пути."
            Select Case CStr(item)
                Case "unknown": data(row, 6) = "UnknownPath"
                Case "excelError": data(row, 6) = CVErr(2042)
                Case "recovery": data(row, 6) = "Auto"
                Case Else: data(row, 6) = vbNullString
            End Select
        Next item
    Next profileId
    source.Value2 = data
    source.Cells(3, 6).Formula = "=" & Chr$(34) & Chr$(34)
    source.Cells(8, 6).Formula = "=" & Chr$(34) & Chr$(34)
    source.Calculate
    Set batch = New CBatchSectionCalculator: batch.Initialize section, provider
    Set batch.ProfileCatalog = profiles
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    batch.ApplySettings settings, units
    Set reader = New CLoadCombinationReader: reader.LoadFromRange source, batch, units
    batch.Execute
    Check stats, "universalPath.reader.keepsAllRows", batch.Count = 10
    For i = 1 To batch.Count
        prefix = "universalPath.reader." & batch.CombinationID(i)
        If (i - 1) Mod 5 = 2 Or (i - 1) Mod 5 = 3 Then
            Check stats, prefix & ".inputErr", batch.ResultAt(i).Status = "InputErr"
            message = batch.ResultAt(i).OverallMeta.ResultComment
            Check stats, prefix & ".address", InStr(1, message, source.Worksheet.Name & "!" & _
                source.Cells(i + 1, 6).Address(False, False), vbTextCompare) > 0
            Check stats, prefix & ".key", InStr(1, message, "LoadPath", vbTextCompare) > 0
        Else
            Check stats, prefix & ".valid", batch.ResultAt(i).Status <> "InputErr" And batch.ResultAt(i).Status <> "CalcErr"
            If i <= 5 Then
                Check stats, prefix & ".autoPath", batch.ResultAt(i).StrengthResult.Capacity.PathResolved = "LambdaMxy"
                Check stats, prefix & ".point", batch.ResultAt(i).StrengthResult.Capacity.SearchResult.HasLimitPoint
            Else
                Check stats, prefix & ".autoPath", batch.ResultAt(i).CrackResult.Formation.FormationMethod = LambdaPath("Mxy")
                Check stats, prefix & ".point", batch.ResultAt(i).CrackResult.Formation.HasLimitPoint
            End If
        End If
    Next i
    stats.Report = stats.Report & Audit03ValidateConfigBatchResults(batch, units, passed, failed)
    stats.Passed = stats.Passed + passed: stats.Failed = stats.Failed + failed
    GoTo Restore
FailedRun:
    failureReason = Err.Description
    stats.Failed = stats.Failed + 1: LogLine stats, "FAIL: universalPath.reader.runtime; " & failureReason
Restore:
    On Error GoTo RestoreFailed
    Application.DisplayAlerts = False
    If Not sheet Is Nothing Then sheet.Delete
    Application.DisplayAlerts = previousAlerts
    Exit Sub
RestoreFailed:
    failureReason = Err.Description
    Application.DisplayAlerts = previousAlerts
    stats.Failed = stats.Failed + 1: LogLine stats, "FAIL: universalPath.reader.restore; " & failureReason
End Sub

' ДЛЯ ТЕСТОВ: эталон задан вручную, а не через factory production-пути.
' Проверяет момент от эксцентриситета N и классификацию Auto по исходным
' компонентам в точке приложения, включая два момента как одно действие.
Private Sub CheckUniversalPathMath(ByRef stats As TConfigTestStats)
    Dim coordinator As CLimitSearchCoordinator, load As CSectionLoadState, descriptor As CLoadPathDescriptor
    Set coordinator = New CLimitSearchCoordinator: Set load = New CSectionLoadState
    Set descriptor = New CLoadPathDescriptor
    Dim item As Variant, keys As Variant, cases As Variant, path As Variant, lambdaValue As Variant
    Dim nActual As Double, mxActual As Double, myActual As Double
    Dim nExpected As Double, mxExpected As Double, myExpected As Double
    Dim scaleN As Boolean, scaleMx As Boolean, scaleMy As Boolean, prefix As String
    cases = Array(Array(20#, 0#, 0#, "LambdaN"), Array(0#, 80#, 0#, "LambdaMx"), _
        Array(0#, 0#, 40#, "LambdaMy"), Array(0#, 80#, 40#, "LambdaMxy"), _
        Array(20#, 80#, 0#, "LambdaMx,LambdaN,LambdaNMxy"), _
        Array(20#, 0#, 40#, "LambdaMy,LambdaN,LambdaNMxy"), _
        Array(20#, 80#, 40#, "LambdaMxy,LambdaN,LambdaNMxy"))
    For Each item In cases
        load.Initialize CDbl(item(0)), CDbl(item(1)), CDbl(item(2)), 3#, -2#
        keys = coordinator.LoadPathKeys("Auto", load)
        Check stats, "universalPath.math.auto." & CStr(item(3)), Join(keys, ",") = CStr(item(3))
        keys = coordinator.LoadPathKeys(vbNullString, load)
        Check stats, "universalPath.math.blank." & CStr(item(3)), Join(keys, ",") = CStr(item(3))
    Next item
    load.Initialize 20#, 80#, 40#, 3#, -2#
    For Each path In Array("LambdaMx", "LambdaMy", "LambdaMxy", "LambdaN", "LambdaNMxy")
        descriptor.InitializeFromLoadState CStr(path), load
        scaleN = (CStr(path) = "LambdaN" Or CStr(path) = "LambdaNMxy")
        scaleMx = (CStr(path) = "LambdaMx" Or CStr(path) = "LambdaMxy" Or CStr(path) = "LambdaNMxy")
        scaleMy = (CStr(path) = "LambdaMy" Or CStr(path) = "LambdaMxy" Or CStr(path) = "LambdaNMxy")
        keys = coordinator.LoadPathKeys(CStr(path), load)
        Check stats, "universalPath.math.fixed." & CStr(path), UBound(keys) = LBound(keys) And CStr(keys(0)) = CStr(path)
        For Each lambdaValue In Array(0#, 1#, 2#)
            nExpected = 20# * IIf(scaleN, CDbl(lambdaValue), 1#)
            mxExpected = 80# * IIf(scaleMx, CDbl(lambdaValue), 1#) - 2# * nExpected
            myExpected = 40# * IIf(scaleMy, CDbl(lambdaValue), 1#) + 3# * nExpected
            descriptor.Target CDbl(lambdaValue), nActual, mxActual, myActual
            prefix = "universalPath.math." & CStr(path) & "." & CStr(lambdaValue)
            CheckClose stats, prefix & ".N", nActual, nExpected, 0#
            CheckClose stats, prefix & ".Mx", mxActual, mxExpected, 0#
            CheckClose stats, prefix & ".My", myActual, myExpected, 0#
        Next lambdaValue
    Next path
    ' Уже нормализованные малые нагрузки не исчезают в описателе пути.
    For Each item In Array(-0.000000000001, 0.000000000001)
        load.Initialize CDbl(item), CDbl(item), CDbl(item), 0#, 0#
        For Each path In Array("LambdaMx", "LambdaMy", "LambdaMxy", "LambdaN", "LambdaNMxy")
            descriptor.InitializeFromLoadState CStr(path), load
            prefix = "universalPath.math.tiny." & CStr(path) & "." & CStr(item)
            Check stats, prefix & ".scaled", descriptor.HasScaledLoad
            descriptor.Target 1#, nActual, mxActual, myActual
            CheckClose stats, prefix & ".N", nActual, CDbl(item), 0#
            CheckClose stats, prefix & ".Mx", mxActual, CDbl(item), 0#
            CheckClose stats, prefix & ".My", myActual, CDbl(item), 0#
        Next path
    Next item
    Dim search As CLimitSearchResult, meta As CResultMeta
    Set search = New CLimitSearchResult: Set meta = New CResultMeta
    meta.SetResult rsNumericalFailure, rcSingularTangent, rkCapacity, "Точная причина последнего поиска."
    search.Initialize meta, "Auto", "LoadMultiplier", False, 0#, 0#, 0#, 0#, Nothing, _
        vbNullString, "Последняя численная попытка."
    search.AddLoadPathHistory "Предыдущий путь не дал точку.", "lambda*N", "Первая численная попытка."
    Check stats, "universalPath.history.keepsBothDiagnostics", search.DiagnosticLog = _
        "Первая численная попытка." & vbCrLf & "Последняя численная попытка."
    Check stats, "universalPath.history.keepsTypedCause", search.Meta.InternalStatus = rsNumericalFailure And _
        search.Meta.ResultCode = rcSingularTangent
    search.AddLoadPathHistory vbNullString, "lambda*N", "Первая численная попытка."
    Check stats, "universalPath.history.noDuplicatePrefix", search.DiagnosticLog = _
        "Первая численная попытка." & vbCrLf & "Последняя численная попытка."
    ' Инженерный FAIL ниже текущего LC не теряет историю успешного Auto-search.
    Dim capacity As CCapacityResult, historyComment As String
    historyComment = "Путь " & ChrW$(&H3BB) & "*Mxy: постоянная часть не проходит. Путь " & ChrW$(&H3BB) & "*N: точка не найдена."
    meta.SetResult rsSuccessWithWarning, rcCheckPassed, rkCapacity, historyComment, "Сохраненная техническая диагностика.", True, True
    search.Initialize meta, "Auto", "UltimateStrain", True, 0.5, 10#, 40#, 20#, Nothing, _
        "ConcreteStrainLimit", "Численный журнал."
    load.Initialize 20#, 80#, 40#, 0#, 0#
    descriptor.InitializeFromLoadState "LambdaNMxy", load
    Set capacity = New CCapacityResult
    capacity.InitializeFromSearch search, descriptor, False, 0#, 0#, Nothing, 200#, 300#
    Check stats, "universalPath.history.capacityFailed", capacity.ResultMeta.InternalStatus = rsCheckFailed And _
        capacity.ResultMeta.ResultCode = rcPhysicalLimitExceeded
    Check stats, "universalPath.history.capacityComment", InStr(1, capacity.ResultMeta.ResultComment, historyComment, vbBinaryCompare) = 1
    Check stats, "universalPath.history.capacityFailureReason", InStr(1, capacity.ResultMeta.ResultComment, "раньше заданной нагрузки", vbBinaryCompare) > 0
    Check stats, "universalPath.history.capacityDiagnostics", capacity.ResultMeta.DiagnosticDetails = "Сохраненная техническая диагностика."
    Check stats, "universalPath.history.searchUnchanged", search.Meta.InternalStatus = rsSuccessWithWarning And _
        search.Meta.ResultComment = historyComment
End Sub

' Выбирает наибольшую независимую долю предельной деформации в тестовом oracle.
' Не использует приватные helpers production-классов и не меняет их критерий.
Private Function LargerValue(ByVal first As Double, ByVal second As Double) As Double
    LargerValue = first
    If second > first Then LargerValue = second
End Function

' ==================== ДЛЯ ТЕСТОВ: СТРУКТУРА ИМЕНОВАННЫХ ТАБЛИЦ CONFIG ====================

' Проверяет поврежденные шапки и ключи на настоящих Excel.Range в отдельной
' временной книге. Возвращает счетчики для общего интерфейсного suite;
' расчетные результаты и исходные пользовательские данные не изменяются.
Public Function RunAudit03SettingsTableGuardTests(Optional ByRef passed As Long = 0, _
        Optional ByRef failed As Long = 0) As String
    Dim stats As TConfigTestStats
    TestAudit03SettingsTableGuards stats
    TestAudit03ProgressFileFailure stats
    LogLine stats, "TOTAL_AUDIT03_SETTINGS_TABLE_GUARDS: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    passed = stats.Passed: failed = stats.Failed
    RunAudit03SettingsTableGuardTests = stats.Report
End Function

' Сохраняет исходную ошибку reader-а отдельно от успешного recovery.
' Чтение обязательного значения проверяет и отложенные ошибки оформления,
' не требуя его применения к рисунку или запуска НДС.
Private Sub Audit03ReadSettingsTable(ByVal settings As CSystemSettingsReader, ByVal fixture As Object, _
        ByVal key As String, ByRef code As Long, ByRef reason As String)
    Dim value As String
    On Error Resume Next
    Err.Clear
    settings.LoadFromWorkbook fixture
    If Err.Number = 0 Then value = settings.GetRequiredString(key)
    code = Err.Number: reason = Err.Description
    Err.Clear
    On Error GoTo 0
End Sub

' Структурные заголовки Units/Signs/Materials определяют смысл колонок.
' Ошибка ключа должна называться до потери его строки, а Excel-ошибка в
' нерасчетном комментарии игнорируется так же, как в scalar Config. Каждый
' сценарий повторяется после перемещения имени; восстановление проверяется
' новым полным чтением, без подмены проверки уже подготовленным словарем.
Private Sub TestAudit03SettingsTableGuards(ByRef stats As TConfigTestStats)
    Dim fixture As Object, sheet As Object, source As Object, table As Object, target As Object, cell As Object
    Dim names As Variant, addresses As Variant, keys As Variant, headerCounts As Variant, comments As Variant
    Dim baselines(0 To 5) As Variant, baseline As Variant, bad As Variant, settings As CSystemSettingsReader
    Dim index As Long, position As Long, column As Long, row As Long, code As Long, cases As Long, solveCount As Long
    Dim reason As String, prefix As String, originalReference As String, expectedValue As String, currentReference As String
    Dim minimumColumns As Variant
    On Error GoTo FailedRun
    solveCount = SectionEquilibriumSolveCount()
    Set fixture = Application.Workbooks.Add(-4167): Set sheet = fixture.Worksheets(1): sheet.Name = "Config"
    names = Array("rngUnitSettings", "rngSignConventionSettings", "rngConcreteMaterialParameters", _
        "rngSteelMaterialParameters", "rngSystemSettings", "rngPlotAnnotationSettings")
    addresses = Array("A5", "A20", "A30", "A50", "H5", "A70")
    keys = Array("Units.Length.Input", "Sign.N.User", "Concrete.Rb.ULS", "Steel.Rsc.ULS", _
        "General.ExecutionReportEnabled", "Plot.RebarLabels.Enabled")
    headerCounts = Array(4, 3, 3, 3, 0, 0): comments = Array(0, 0, 5, 5, 4, 5)
    minimumColumns = Array(4, 3, 4, 4, 0, 0)
    For index = 0 To UBound(names)
        Set source = ThisWorkbook.Names.Item(CStr(names(index))).RefersToRange
        baselines(index) = source.Value2
        Set table = sheet.Range(CStr(addresses(index))).Resize(source.Rows.Count, source.Columns.Count)
        table.NumberFormat = "@": table.Value2 = baselines(index)
        fixture.Names.Add Name:=CStr(names(index)), RefersTo:="=Config!" & table.Address
    Next index
    Set settings = New CSystemSettingsReader
    For index = 0 To UBound(names)
        baseline = baselines(index)
        originalReference = fixture.Names.Item(CStr(names(index))).RefersTo
        For position = 0 To 1
            If position = 0 Then Set table = sheet.Range(CStr(addresses(index))).Resize(UBound(baseline, 1), UBound(baseline, 2)) Else Set table = sheet.Range("CH800").Resize(UBound(baseline, 1), UBound(baseline, 2))
            table.NumberFormat = "@": table.Value2 = baseline
            fixture.Names.Item(CStr(names(index))).RefersTo = "=Config!" & table.Address
            settings.LoadFromWorkbook fixture
            expectedValue = settings.GetRequiredString(CStr(keys(index)))
            prefix = "audit03.settingsTable.p" & CStr(position) & "." & CStr(names(index))
            If CLng(headerCounts(index)) > 0 Then
                cases = cases + 1
                For column = 1 To CLng(headerCounts(index))
                    table.Cells(1, column).Value2 = " " & LCase$(CStr(baseline(1, column))) & " "
                Next column
                Audit03ReadSettingsTable settings, fixture, CStr(keys(index)), code, reason
                Check stats, prefix & ".headerCase.accepted", code = 0
                If code = 0 Then Check stats, prefix & ".headerCase.value", settings.GetRequiredString(CStr(keys(index))) = expectedValue
            End If
            For column = 1 To CLng(headerCounts(index))
                Set cell = table.Cells(1, column)
                For Each bad In Array("", "TODO", "UNKNOWN_HEADER", CVErr(2015))
                    cases = cases + 1: table.Value2 = baseline: cell.Value2 = bad
                    Audit03ReadSettingsTable settings, fixture, CStr(keys(index)), code, reason
                    Check stats, prefix & ".header" & CStr(column) & ".bad" & CStr(VarType(bad)) & ".rejected", code <> 0
                    Check stats, prefix & ".header" & CStr(column) & ".bad" & CStr(VarType(bad)) & ".typed", _
                        code = vbObjectError + 4309 Or code = vbObjectError + 4316
                    Check stats, prefix & ".header" & CStr(column) & ".bad" & CStr(VarType(bad)) & ".address", _
                        InStr(1, reason, cell.Address(False, False), vbTextCompare) > 0 And InStr(1, reason, "Config", vbTextCompare) > 0
                    Check stats, prefix & ".header" & CStr(column) & ".bad" & CStr(VarType(bad)) & ".action", _
                        InStr(1, reason, "Восстановите", vbTextCompare) > 0 Or InStr(1, reason, "Исправьте", vbTextCompare) > 0
                    LogLine stats, "SETTINGS_TABLE_ERROR: " & prefix & ".header" & CStr(column) & "|code=" & CStr(code) & "|" & reason
                Next bad
            Next column
            table.Value2 = baseline: row = 2
            If index = 4 Then row = ValueCell(table, CStr(keys(index)), 2).Row - table.Row + 1
            Set cell = table.Cells(row, 1): cell.Value2 = CVErr(2015): cases = cases + 1
            Audit03ReadSettingsTable settings, fixture, CStr(keys(index)), code, reason
            Check stats, prefix & ".key.typed", code = vbObjectError + 4309
            Check stats, prefix & ".key.address", InStr(1, reason, cell.Address(False, False), vbTextCompare) > 0
            Check stats, prefix & ".key.action", InStr(1, reason, "Исправьте", vbTextCompare) > 0
            LogLine stats, "SETTINGS_TABLE_ERROR: " & prefix & ".key|code=" & CStr(code) & "|" & reason
            table.Value2 = baseline: Set cell = Nothing
            Set cell = table.Cells(row, 2): cell.Value2 = CVErr(2015): cases = cases + 1
            Audit03ReadSettingsTable settings, fixture, CStr(keys(index)), code, reason
            Check stats, prefix & ".value.typed", code = vbObjectError + 4309
            Check stats, prefix & ".value.address", InStr(1, reason, cell.Address(False, False), vbTextCompare) > 0
            LogLine stats, "SETTINGS_TABLE_ERROR: " & prefix & ".value|code=" & CStr(code) & "|" & reason
            If CLng(comments(index)) > 0 Then
                table.Value2 = baseline: Set cell = table.Cells(row, CLng(comments(index)))
                cell.Value2 = CVErr(2015): cases = cases + 1
                Audit03ReadSettingsTable settings, fixture, CStr(keys(index)), code, reason
                Check stats, prefix & ".comment.noncritical", code = 0
                If code = 0 Then
                    Check stats, prefix & ".comment.emptySource", settings.GetSource(CStr(keys(index))) = vbNullString
                    Check stats, prefix & ".comment.valuePreserved", settings.GetRequiredString(CStr(keys(index))) = expectedValue
                End If
            End If
            table.Value2 = baseline
            Audit03ReadSettingsTable settings, fixture, CStr(keys(index)), code, reason
            Check stats, prefix & ".recovery", code = 0
            If code = 0 Then Check stats, prefix & ".recovery.value", settings.GetRequiredString(CStr(keys(index))) = expectedValue
            cases = cases + 1
            settings.LoadFromRange table
            Check stats, prefix & ".rangeFormat.value", settings.GetRequiredString(CStr(keys(index))) = expectedValue
            If CLng(minimumColumns(index)) > 0 Then
                cases = cases + 1: currentReference = fixture.Names.Item(CStr(names(index))).RefersTo
                fixture.Names.Item(CStr(names(index))).RefersTo = "=Config!" & table.Resize(table.Rows.Count, CLng(minimumColumns(index)) - 1).Address
                Audit03ReadSettingsTable settings, fixture, CStr(keys(index)), code, reason
                Check stats, prefix & ".truncated.typed", code = vbObjectError + 4316
                Check stats, prefix & ".truncated.address", InStr(1, reason, table.Cells(1, 1).Address(False, False), vbTextCompare) > 0
                Check stats, prefix & ".truncated.action", InStr(1, reason, "Восстановите", vbTextCompare) > 0
                LogLine stats, "SETTINGS_TABLE_ERROR: " & prefix & ".truncated|code=" & CStr(code) & "|" & reason
                fixture.Names.Item(CStr(names(index))).RefersTo = currentReference
            End If
        Next position
        fixture.Names.Item(CStr(names(index))).RefersTo = originalReference
    Next index
    ' Reader не навязывает материал, пока он не нужен активному расчетному
    ' потребителю. Пропавшая optional-таблица остается отсутствующим ключом.
    For index = 2 To 3
        cases = cases + 1: originalReference = fixture.Names.Item(CStr(names(index))).RefersTo
        fixture.Names.Item(CStr(names(index))).Delete
        settings.LoadFromWorkbook fixture
        Check stats, "audit03.settingsTable.optional." & CStr(names(index)), Not settings.HasKey(CStr(keys(index)))
        fixture.Names.Add Name:=CStr(names(index)), RefersTo:=originalReference
        settings.LoadFromWorkbook fixture
        Check stats, "audit03.settingsTable.optionalRecovery." & CStr(names(index)), settings.HasKey(CStr(keys(index)))
    Next index
    Check stats, "audit03.settingsTable.noSolve", SectionEquilibriumSolveCount() = solveCount
    LogLine stats, "SETTINGS_TABLE_CASES: blocks=6; positions=2; variants=" & CStr(cases) & "; equilibriumCases=0"
    GoTo Cleanup
FailedRun:
    Check stats, "audit03.settingsTable.runtime." & CStr(Err.Number) & "." & Err.Description, False
Cleanup:
    On Error Resume Next
    If Not fixture Is Nothing Then fixture.Close False
    On Error GoTo 0
End Sub

' ==================== ДЛЯ ТЕСТОВ: НЕСМЕЖНЫЕ ОБЛАСТИ CONFIG ====================

' Проверяет, что чтение Config не пропускает часть несмежного Excel.Range.
' Настоящие таблицы копируются в собственную книгу; отдельный entrypoint
' позволяет сохранить отрицательный опыт до изменения production-reader-а.
Public Function RunAudit03MultiAreaInputTests(Optional ByRef passed As Long = 0, _
        Optional ByRef failed As Long = 0) As String
    Dim stats As TConfigTestStats
    CheckMultiAreaInputs stats
    LogLine stats, "TOTAL_AUDIT03_MULTI_AREA_INPUT: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    passed = stats.Passed: failed = stats.Failed
    RunAudit03MultiAreaInputTests = stats.Report
End Function

' Для каждого формата сохраняет полноценную первую область и добавляет
' вторую с контрольной строкой. Прямой Union.Range проверяет отказ до Value2
' и текущий адрес. Если Excel не разрешает созданное имя через RefersToRange,
' workbook-reader должен сообщить о недоступной таблице, а не игнорировать
' ее поврежденное имя. После восстановления проверяет оба entrypoint-а.
Private Sub CheckMultiAreaInputs(ByRef stats As TConfigTestStats)
    Dim fixture As Object, config As Object, source As Object, table As Object, tail As Object, combined As Object
    Dim names As Variant, addresses As Variant, keys As Variant, baselines(0 To 5) As Variant, baseline As Variant
    Dim settings As CSystemSettingsReader, index As Long, position As Long, code As Long, solveCount As Long
    Dim reason As String, prefix As String, originalReference As String, data As Variant
    On Error GoTo FailedRun
    solveCount = SectionEquilibriumSolveCount()
    Set fixture = Application.Workbooks.Add(-4167): Set config = fixture.Worksheets(1): config.Name = "InputConfig"
    names = Array("rngUnitSettings", "rngSignConventionSettings", "rngConcreteMaterialParameters", _
        "rngSteelMaterialParameters", "rngSystemSettings", "rngPlotAnnotationSettings")
    addresses = Array("A5", "A20", "A30", "A50", "H5", "A70")
    keys = Array("Units.Length.Input", "Sign.N.User", "Concrete.Rb.ULS", "Steel.Rsc.ULS", _
        "General.ExecutionReportEnabled", "Plot.RebarLabels.Enabled")
    For index = 0 To UBound(names)
        LogLine stats, "MULTI_AREA_SETUP: copy " & CStr(names(index))
        Set source = ThisWorkbook.Names.Item(CStr(names(index))).RefersToRange
        baselines(index) = source.Value2
        Set table = config.Range(CStr(addresses(index))).Resize(source.Rows.Count, source.Columns.Count)
        table.NumberFormat = "@": table.Value2 = baselines(index)
        fixture.Names.Add Name:=CStr(names(index)), RefersTo:="='" & config.Name & "'!" & table.Address
    Next index
    Set settings = New CSystemSettingsReader
    For index = 0 To UBound(names)
        baseline = baselines(index): originalReference = fixture.Names.Item(CStr(names(index))).RefersTo
        For position = 0 To 1
            If position = 0 Then Set table = config.Range(CStr(addresses(index))).Resize(UBound(baseline, 1), UBound(baseline, 2)) Else Set table = config.Range("CH800").Resize(UBound(baseline, 1), UBound(baseline, 2))
            table.NumberFormat = "@": table.Value2 = baseline
            Set tail = config.Range("DA1200").Resize(2, UBound(baseline, 2)): tail.ClearContents
            tail.Cells(1, 1).Value2 = "Параметр": tail.Cells(1, 2).Value2 = "Значение": tail.Cells(1, 3).Value2 = "Ед."
            tail.Cells(2, 1).Value2 = "Audit03.MultiArea.Marker": tail.Cells(2, 2).Value2 = "SECOND_AREA"
            LogLine stats, "MULTI_AREA_SETUP: union " & CStr(names(index)) & "|" & table.Address & "|" & tail.Address
            fixture.Names.Item(CStr(names(index))).Delete
            Set combined = Application.Union(table, tail)
            combined.Name = CStr(names(index))
            LogLine stats, "MULTI_AREA_SETUP: resolve " & fixture.Names.Item(CStr(names(index))).RefersTo
            prefix = "audit03.multiArea.p" & CStr(position) & "." & CStr(names(index))
            Check stats, prefix & ".fixture", combined.Areas.Count = 2
            data = combined.Value2
            LogLine stats, "MULTI_AREA_VALUE2: " & prefix & "|areas=" & CStr(combined.Areas.Count) & _
                "|returnedRows=" & CStr(UBound(data, 1)) & "|firstAreaRows=" & CStr(table.Rows.Count) & _
                "|secondAreaRows=" & CStr(tail.Rows.Count)
            Audit03ReadSettingsTable settings, fixture, CStr(keys(index)), code, reason
            Check stats, prefix & ".named.typed", code = vbObjectError + 4316
            Check stats, prefix & ".named.name", InStr(1, reason, CStr(names(index)), vbTextCompare) > 0
            Check stats, prefix & ".named.reason", InStr(1, reason, "област", vbTextCompare) > 0 Or _
                InStr(1, reason, "недоступ", vbTextCompare) > 0
            Check stats, prefix & ".named.action", InStr(1, reason, "Восстановите", vbTextCompare) > 0 And _
                InStr(1, reason, "диспетчере имен", vbTextCompare) > 0
            LogLine stats, "MULTI_AREA_ERROR: " & prefix & ".named|code=" & CStr(code) & _
                "|secondAreaRead=" & CStr(settings.HasKey("Audit03.MultiArea.Marker")) & "|" & reason
            Audit03ReadMultiAreaRange settings, combined, code, reason
            Check stats, prefix & ".range.typed", code = vbObjectError + 4316
            Check stats, prefix & ".range.address", InStr(1, reason, config.Name, vbTextCompare) > 0 And _
                InStr(1, reason, table.Cells(1, 1).Address(False, False), vbTextCompare) > 0
            Check stats, prefix & ".range.reason", InStr(1, reason, "област", vbTextCompare) > 0
            Check stats, prefix & ".range.action", InStr(1, reason, "один непрерывный прямоугольный", vbTextCompare) > 0
            LogLine stats, "MULTI_AREA_ERROR: " & prefix & ".range|code=" & CStr(code) & _
                "|secondAreaRead=" & CStr(settings.HasKey("Audit03.MultiArea.Marker")) & "|" & reason
            fixture.Names.Item(CStr(names(index))).RefersTo = "='" & config.Name & "'!" & table.Address
            Audit03ReadSettingsTable settings, fixture, CStr(keys(index)), code, reason
            Check stats, prefix & ".named.recovery", code = 0
            If code = 0 Then Check stats, prefix & ".named.recoveryValue", Len(settings.GetRequiredString(CStr(keys(index)))) > 0
            Audit03ReadMultiAreaRange settings, table, code, reason
            Check stats, prefix & ".range.recovery", code = 0
            If code = 0 Then Check stats, prefix & ".range.recoveryValue", Len(settings.GetRequiredString(CStr(keys(index)))) > 0
        Next position
        fixture.Names.Item(CStr(names(index))).RefersTo = originalReference
    Next index
    Check stats, "audit03.multiArea.noSolve", SectionEquilibriumSolveCount() = solveCount
    LogLine stats, "MULTI_AREA_CASES: blocks=6; positions=2; entrypoints=2; invalidCases=24; equilibriumCases=0"
    GoTo Cleanup
FailedRun:
    Check stats, "audit03.multiArea.runtime." & CStr(Err.Number) & "." & Err.Description, False
Cleanup:
    On Error Resume Next
    If Not fixture Is Nothing Then fixture.Close False
    On Error GoTo 0
End Sub

' Сохраняет исходную ошибку автономного reader-а, не применяя defaults и
' не вызывая расчетные consumers; повторный вызов проверяет очистку reader-а.
Private Sub Audit03ReadMultiAreaRange(ByVal settings As CSystemSettingsReader, ByVal table As Object, _
        ByRef code As Long, ByRef reason As String)
    On Error Resume Next
    Err.Clear
    settings.LoadFromRange table
    code = Err.Number: reason = Err.Description
    Err.Clear
    On Error GoTo 0
End Sub

' ==================== ДЛЯ ТЕСТОВ: ОБЩИЕ НАСТРОЙКИ ЗАПУСКА ====================

' Проверяет реальные потребители флагов отчета и информационных сообщений.
' Сохраненная собственная книга изолирует Config, файлы и пользовательский
' Excel. НДС не запускается; отчет проверяется по фактическому txt-файлу.
Public Function RunAudit03GeneralRunConfigTests(Optional ByRef passed As Long = 0, _
        Optional ByRef failed As Long = 0) As String
    Dim stats As TConfigTestStats
    CheckGeneralRunControls stats
    LogLine stats, "TOTAL_AUDIT03_GENERAL_RUN_CONFIG: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    passed = stats.Passed: failed = stats.Failed
    RunAudit03GeneralRunConfigTests = stats.Report
End Function

' Сравнивает Yes/No с реальным файлом отчета и решением показа сообщения.
' Ошибочные значения/удаленные ключи проверяются до и после переноса Name;
' recovery и повторный Initialize не должны сохранять прежнее состояние.
Private Sub CheckGeneralRunControls(ByRef stats As TConfigTestStats)
    Dim fixture As Object, config As Object, source As Object, table As Object, target As Object, cell As Object
    Dim settings As CSystemSettingsReader, report As CExecutionReport, fso As Object, stream As Object
    Dim name As Variant, addresses As Variant, baseline As Variant, bad As Variant, badValues As Variant, keys As Variant, key As Variant
    Dim index As Long, position As Long, mode As Long, caseIndex As Long, code As Long, cases As Long, solveCount As Long
    Dim reason As String, prefix As String, folder As String, path As String, originalText As String, currentText As String
    Dim value As Boolean
    On Error GoTo FailedRun
    solveCount = SectionEquilibriumSolveCount()
    Set fso = CreateObject("Scripting.FileSystemObject")
    folder = ThisWorkbook.Path & "\Audit03_RunControls_" & Format$(Now, "yyyymmdd_hhnnss")
    fso.CreateFolder folder
    Set fixture = Application.Workbooks.Add(-4167): Set config = fixture.Worksheets(1): config.Name = "RunConfig"
    addresses = Array("A5", "H5", "H20", "H30")
    For Each name In Array("rngSystemSettings", "rngUnitSettings", "rngSignConventionSettings", "rngPlotAnnotationSettings")
        Set source = ThisWorkbook.Names.Item(CStr(name)).RefersToRange
        Set target = config.Range(CStr(addresses(index))).Resize(source.Rows.Count, source.Columns.Count)
        target.NumberFormat = "@": target.Value2 = source.Value2
        fixture.Names.Add Name:=CStr(name), RefersTo:="='" & config.Name & "'!" & target.Address
        index = index + 1
    Next name
    fixture.SaveAs folder & "\RunControls.xlsx", 51
    Set table = fixture.Names.Item("rngSystemSettings").RefersToRange: baseline = table.Value2
    Set settings = New CSystemSettingsReader: Set report = New CExecutionReport
    keys = Array("General.ExecutionReportEnabled", "General.NonCriticalMessagesEnabled")
    badValues = Array("", "TODO", "INVALID", CVErr(2015))
    For position = 0 To 1
        If position = 1 Then Set table = config.Range("CH800").Resize(UBound(baseline, 1), UBound(baseline, 2))
        table.NumberFormat = "@": table.Value2 = baseline
        fixture.Names.Item("rngSystemSettings").RefersTo = "='" & config.Name & "'!" & table.Address
        SetValue table, "General.ExecutionReportEnabled", "Yes"
        settings.LoadFromWorkbook fixture: report.Initialize fixture, settings
        report.AddStep "AUDIT03_REPORT_MARKER": report.Save "AUDIT03_FINAL"
        path = report.FilePath
        Check stats, "audit03.generalRun.report.Yes.p" & CStr(position), report.Enabled And fso.FileExists(path)
        Set stream = fso.OpenTextFile(path, 1, False, -1): originalText = stream.ReadAll: stream.Close
        Check stats, "audit03.generalRun.report.content.p" & CStr(position), InStr(1, originalText, "AUDIT03_REPORT_MARKER", vbBinaryCompare) > 0 And InStr(1, originalText, "AUDIT03_FINAL", vbBinaryCompare) > 0
        SetValue table, "General.ExecutionReportEnabled", "No"
        settings.LoadFromWorkbook fixture: report.Initialize fixture, settings
        report.AddError "AUDIT03", "SHOULD_NOT_BE_WRITTEN": report.Save "SHOULD_NOT_BE_WRITTEN"
        Set stream = fso.OpenTextFile(path, 1, False, -1): currentText = stream.ReadAll: stream.Close
        Check stats, "audit03.generalRun.report.No.p" & CStr(position), Not report.Enabled And report.FilePath = vbNullString And currentText = originalText
        For Each bad In Array("Yes", "No", " TRUE ", " FALSE ", "1", "0", "Да", "Нет")
            SetValue table, "General.NonCriticalMessagesEnabled", bad
            CaptureGeneralRunControl fixture, report, False, code, reason, value
            Check stats, "audit03.generalRun.messages.valid.p" & CStr(position) & "." & CStr(bad), code = 0
            Check stats, "audit03.generalRun.messages.effect.p" & CStr(position) & "." & CStr(bad), value = (CStr(bad) = "Yes" Or CStr(bad) = " TRUE " Or CStr(bad) = "1" Or CStr(bad) = "Да")
        Next bad
        For mode = 0 To 1
            key = keys(mode): Set cell = ValueCell(table, CStr(key), 2)
            For caseIndex = 0 To 4
                table.Value2 = baseline: Set cell = ValueCell(table, CStr(key), 2)
                If caseIndex = 4 Then
                    cell.Offset(0, -1).Value2 = "AUDIT03_REMOVED_" & CStr(key)
                Else
                    bad = badValues(caseIndex)
                    cell.Value2 = bad
                End If
                cases = cases + 1
                CaptureGeneralRunControl fixture, report, mode = 0, code, reason, value
                prefix = "audit03.generalRun.invalid.p" & CStr(position) & "." & CStr(key) & ".v" & CStr(caseIndex)
                Check stats, prefix & ".rejected", code <> 0
                Check stats, prefix & ".typed", code >= vbObjectError + 4300 And code <= vbObjectError + 4399
                Check stats, prefix & ".key", InStr(1, reason, CStr(key), vbTextCompare) > 0
                Check stats, prefix & ".action", InStr(1, reason, "Выберите", vbTextCompare) > 0 Or InStr(1, reason, "Заполните", vbTextCompare) > 0 Or InStr(1, reason, "Введите", vbTextCompare) > 0 Or InStr(1, reason, "Исправьте", vbTextCompare) > 0
                If caseIndex < 4 Then Check stats, prefix & ".address", InStr(1, reason, config.Name, vbTextCompare) > 0 And InStr(1, reason, cell.Address(False, False), vbTextCompare) > 0
                If mode = 0 Then Check stats, prefix & ".reset", Not report.Enabled And report.FilePath = vbNullString
                LogLine stats, "GENERAL_RUN_ERROR: " & prefix & "|" & reason
            Next caseIndex
            table.Value2 = baseline: SetValue table, CStr(key), "Yes"
            CaptureGeneralRunControl fixture, report, mode = 0, code, reason, value
            Check stats, "audit03.generalRun.recovery.p" & CStr(position) & "." & CStr(key), code = 0 And value
        Next mode
        report.AddStep "AUDIT03_RECOVERY": report.Save "AUDIT03_RECOVERY_FINAL"
        Set stream = fso.OpenTextFile(report.FilePath, 1, False, -1): currentText = stream.ReadAll: stream.Close
        Check stats, "audit03.generalRun.report.reinitialize.p" & CStr(position), InStr(1, currentText, "AUDIT03_RECOVERY", vbBinaryCompare) > 0 And InStr(1, currentText, "AUDIT03_REPORT_MARKER", vbBinaryCompare) = 0
    Next position
    CaptureGeneralRunControl Nothing, report, False, code, reason, value
    Check stats, "audit03.generalRun.messages.noWorkbook", code <> 0 And InStr(1, reason, "Книга", vbTextCompare) > 0
    Check stats, "audit03.generalRun.noSolve", SectionEquilibriumSolveCount() = solveCount
    LogLine stats, "GENERAL_RUN_CASES: positions=2; invalidVariants=" & CStr(cases) & "; equilibriumCases=0"
    GoTo Cleanup
FailedRun:
    Check stats, "audit03.generalRun.runtime." & CStr(Err.Number) & "." & Err.Description, False
Cleanup:
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    If Not fixture Is Nothing Then fixture.Close False
    On Error GoTo 0
End Sub

' Вызывает настоящий потребитель без MsgBox и сохраняет исходную typed ошибку.
' Для отчета проверяется повторный Initialize; для сообщения используется
' тот же публичный selector, который вызывают пользовательские кнопки.
Private Sub CaptureGeneralRunControl(ByVal fixture As Object, ByVal report As CExecutionReport, _
        ByVal reportMode As Boolean, ByRef code As Long, ByRef reason As String, ByRef value As Boolean)
    On Error GoTo Failed
    code = 0: reason = vbNullString: value = False
    If reportMode Then
        Dim settings As CSystemSettingsReader
        Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook fixture
        report.Initialize fixture, settings: value = report.Enabled
    Else
        value = NonCriticalMessagesEnabled(fixture)
    End If
    Exit Sub
Failed:
    code = Err.Number: reason = Err.Description
End Sub

' ==================== ДЛЯ ТЕСТОВ: ЧИСЛЕННЫЕ НАСТРОЙКИ CAPACITY ====================

' Читает семь численных параметров из настоящей таблицы Config в двух местах.
' Проверяет наблюдаемые пробы Search, физические конечные точки, ошибки ввода
' и восстановление после них. Перенос таблицы не должен менять адресный контракт.
Public Function RunAudit03CapacityNumericConfigTests(Optional ByRef passed As Long = 0, _
        Optional ByRef failed As Long = 0) As String
    Dim stats As TConfigTestStats, originalRange As Object, systemRange As Object, movedRange As Object
    Dim unitRange As Object, profileRange As Object, savedName As String
    Dim savedSystem As Variant, savedMoved As Variant, savedUnits As Variant, savedProfiles As Variant
    Dim configured As Variant, position As Long, settings As CSystemSettingsReader
    Dim section As CSectionModel, provider As CMaterialModelProvider, profiles As CCalculationProfileCatalog, units As CUnitSystem
    On Error GoTo FailedRun
    Set originalRange = ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange
    savedName = ThisWorkbook.Names.Item("rngSystemSettings").RefersTo
    Set movedRange = originalRange.Worksheet.Range("CH800").Resize(originalRange.Rows.Count, originalRange.Columns.Count)
    Set unitRange = ThisWorkbook.Names.Item("rngUnitSettings").RefersToRange
    Set profileRange = ThisWorkbook.Names.Item("rngCalculationProfiles").RefersToRange
    savedSystem = originalRange.Formula: savedMoved = movedRange.Formula
    savedUnits = unitRange.Formula: savedProfiles = profileRange.Formula
    ConfigureSearchFixture originalRange, unitRange, profileRange
    SetProfileValue profileRange, "Calculation.Strength.DirectState", "No", "PR1"
    configured = originalRange.Formula
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    Set provider = New CMaterialModelProvider: provider.Initialize settings, units
    Set profiles = New CCalculationProfileCatalog: profiles.LoadFromWorkbook ThisWorkbook
    Set section = SearchFixtureSection()
    For position = 1 To 2
        If position = 1 Then
            Set systemRange = originalRange
        Else
            Set systemRange = movedRange
        End If
        systemRange.Formula = configured
        ThisWorkbook.Names.Item("rngSystemSettings").RefersTo = "=" & systemRange.Address(True, True, 1, True)
        CheckCapacityNumericEffects stats, systemRange, configured, position, section, provider, profiles, units
        CheckCapacityNumericIsolation stats, systemRange, configured, position, section, provider, profiles, units
        CheckCapacityNumericInput stats, systemRange, configured, position, section, provider, profiles, units
    Next position
    TestAudit03CapacityRetryRange stats
    GoTo Restore
FailedRun:
    Check stats, "audit03.capacityNumeric.runtime." & CStr(Err.Number) & "." & Err.Description, False
Restore:
    On Error Resume Next
    ThisWorkbook.Names.Item("rngSystemSettings").RefersTo = savedName
    If Not originalRange Is Nothing Then originalRange.Formula = savedSystem
    If Not movedRange Is Nothing Then movedRange.Formula = savedMoved
    If Not unitRange Is Nothing Then unitRange.Formula = savedUnits
    If Not profileRange Is Nothing Then profileRange.Formula = savedProfiles
    On Error GoTo 0
    LogLine stats, "TOTAL_AUDIT03_CAPACITY_NUMERIC_CONFIG: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    passed = stats.Passed: failed = stats.Failed
    RunAudit03CapacityNumericConfigTests = stats.Report
End Function

' Меняет по одному параметру на общей модели: реальный Search должен менять
' численный маршрут, но не выдавать диагностического кандидата за физический
' предел. Неактивные параметры одномерной скобки проверяются при UltimateStrain.
Private Sub CheckCapacityNumericEffects(ByRef stats As TConfigTestStats, ByVal table As Object, _
        ByRef configured As Variant, ByVal position As Long, ByVal section As CSectionModel, _
        ByVal provider As CMaterialModelProvider, ByVal profiles As CCalculationProfileCatalog, ByVal units As CUnitSystem)
    Dim baseline As CCapacityResult, changed As CCapacityResult, other As CCapacityResult
    Dim prefix As String, key As Variant, value As Variant, index As Long, countBase As Long, countChanged As Long
    prefix = "audit03.capacityNumeric.p" & CStr(position)
    table.Formula = configured
    SetValue table, "Capacity.SolutionStrategy", "LoadMultiplier"
    Set baseline = CapacityNumericRun(stats, prefix & ".baseline", section, provider, profiles, units)
    CheckCapacityNumericPoint stats, prefix & ".baseline", baseline

    SetValue table, "Capacity.MaxLambda", 0.125
    Set changed = CapacityNumericRun(stats, prefix & ".Capacity.MaxLambda.bounded", section, provider, profiles, units)
    Check stats, prefix & ".Capacity.MaxLambda.noPoint", Not changed.SearchResult.HasLimitPoint
    Check stats, prefix & ".Capacity.MaxLambda.code", changed.ResultMeta.ResultCode = rcSearchBoundReached
    CheckClose stats, prefix & ".Capacity.MaxLambda.lastProbe", CapacityNumericLastLambda(changed.DiagnosticLog), 0.125, 0#

    table.Formula = configured: SetValue table, "Capacity.SolutionStrategy", "LoadMultiplier"
    SetValue table, "Capacity.InitialLambda", 0.125
    Set changed = CapacityNumericRun(stats, prefix & ".Capacity.InitialLambda.small", section, provider, profiles, units)
    CheckCapacityNumericPoint stats, prefix & ".Capacity.InitialLambda.small", changed
    CheckClose stats, prefix & ".Capacity.InitialLambda.firstProbe", CapacityNumericFirstLambda(changed.DiagnosticLog), 0.125, 0#
    CheckClose stats, prefix & ".Capacity.InitialLambda.sameLimit", changed.LambdaCapacity, baseline.LambdaCapacity, 0.000002

    table.Formula = configured: SetValue table, "Capacity.SolutionStrategy", "LoadMultiplier"
    SetValue table, "Capacity.ToleranceLambda", 0.001
    Set changed = CapacityNumericRun(stats, prefix & ".Capacity.ToleranceLambda.coarse", section, provider, profiles, units)
    CheckCapacityNumericPoint stats, prefix & ".Capacity.ToleranceLambda.coarse", changed
    countBase = CapacityNumericRecordCount(baseline.DiagnosticLog, "probe=")
    countChanged = CapacityNumericRecordCount(changed.DiagnosticLog, "probe=")
    Check stats, prefix & ".Capacity.ToleranceLambda.fewerProbes", countChanged < countBase And countChanged > 0
    CheckClose stats, prefix & ".Capacity.ToleranceLambda.sameLimit", changed.LambdaCapacity, baseline.LambdaCapacity, 0.002

    table.Formula = configured: SetValue table, "Capacity.SolutionStrategy", "LoadMultiplier"
    SetValue table, "Capacity.MaxLambda", 0.125: SetValue table, "Capacity.InitialLambda", 0.125
    SetValue table, "Capacity.MaxRetries", 0: SetValue table, "Capacity.BaseLoadSteps", 1
    Set changed = CapacityNumericRun(stats, prefix & ".Capacity.BaseLoadSteps.one", section, provider, profiles, units)
    SetValue table, "Capacity.BaseLoadSteps", 3
    Set other = CapacityNumericRun(stats, prefix & ".Capacity.BaseLoadSteps.three", section, provider, profiles, units)
    Check stats, prefix & ".Capacity.BaseLoadSteps.oneStep", InStr(1, changed.LastSolverDiagnosticLog, "step=1;", vbBinaryCompare) > 0
    Check stats, prefix & ".Capacity.BaseLoadSteps.threeSteps", InStr(1, other.LastSolverDiagnosticLog, "step=3;", vbBinaryCompare) > 0
    Check stats, prefix & ".Capacity.BaseLoadSteps.sameBound", changed.ResultMeta.ResultCode = rcSearchBoundReached And _
        other.ResultMeta.ResultCode = rcSearchBoundReached

    table.Formula = configured: SetValue table, "Capacity.SolutionStrategy", "LoadMultiplier"
    SetValue table, "Capacity.SolverMaxIterations", 1: SetValue table, "Capacity.MaxRetries", 0
    Set changed = CapacityNumericRun(stats, prefix & ".Capacity.MaxRetries.zero", section, provider, profiles, units)
    SetValue table, "Capacity.MaxRetries", 2
    Set other = CapacityNumericRun(stats, prefix & ".Capacity.MaxRetries.two", section, provider, profiles, units)
    Check stats, prefix & ".Capacity.MaxRetries.noRepeat", CapacityNumericRecordCount(changed.DiagnosticLog, "; attempt=1;") = 0
    Check stats, prefix & ".Capacity.MaxRetries.repeatUsed", CapacityNumericRecordCount(other.DiagnosticLog, "; attempt=1;") > 0 And _
        CapacityNumericRecordCount(other.DiagnosticLog, "; attempt=2;") > 0
    Check stats, prefix & ".Capacity.MaxRetries.budget", CapacityNumericRecordCount(other.DiagnosticLog, "; attempt=3;") = 0
    Check stats, prefix & ".Capacity.MaxRetries.noFalsePoint", Not changed.SearchResult.HasLimitPoint And Not other.SearchResult.HasLimitPoint

    table.Formula = configured: SetValue table, "Capacity.SolutionStrategy", "UltimateStrain"
    Set baseline = CapacityNumericRun(stats, prefix & ".ultimateBaseline", section, provider, profiles, units)
    CheckCapacityNumericPoint stats, prefix & ".ultimateBaseline", baseline
    SetValue table, "Capacity.SolverMaxIterations", 1
    Set changed = CapacityNumericRun(stats, prefix & ".Capacity.SolverMaxIterations.one", section, provider, profiles, units)
    Check stats, prefix & ".Capacity.SolverMaxIterations.typedFailure", changed.ResultMeta.InternalStatus = rsNumericalFailure
    Check stats, prefix & ".Capacity.SolverMaxIterations.noPoint", Not changed.SearchResult.HasLimitPoint
    ' Один Newton-шаг включает оценки трех колонок Якобиана и line-search.
    ' Лимит итераций не равен числу EvaluateStrainPlane в диагностике.
    Check stats, prefix & ".Capacity.SolverMaxIterations.budget", _
        CapacityNumericRecordCount(changed.DiagnosticLog, "ultimatePath iter=") < _
        CapacityNumericRecordCount(baseline.DiagnosticLog, "ultimatePath iter=")

    table.Formula = configured: SetValue table, "Capacity.SolutionStrategy", "UltimateStrain"
    SetValue table, "Capacity.ToleranceStrain", 0.000000001
    Set changed = CapacityNumericRun(stats, prefix & ".Capacity.ToleranceStrain.tight", section, provider, profiles, units)
    CheckCapacityNumericPoint stats, prefix & ".Capacity.ToleranceStrain.tight", changed
    countBase = CapacityNumericRecordCount(baseline.DiagnosticLog, "ultimatePath iter=")
    countChanged = CapacityNumericRecordCount(changed.DiagnosticLog, "ultimatePath iter=")
    ' Финальный контроль равновесия здесь строже обоих деформационных допусков,
    ' поэтому точка вправе совпасть. Отдельно проверяем настоящий критерий
    ' остановки на той же физической плоскости с заданной невязкой деформации.
    CheckCapacityStrainTolerance stats, prefix, table, section, provider, units, baseline.StateResult
    CheckClose stats, prefix & ".Capacity.ToleranceStrain.sameLimit", changed.LambdaCapacity, baseline.LambdaCapacity, 0.0002
    CheckCapacityStrainToleranceBatch stats, prefix, table, configured, section, provider, profiles, units

    ' Эти настройки не участвуют в успешном прямом UltimateStrain.
    ' Валидность чтения сохраняется, но предел и выполненные пробы одинаковы.
    For Each key In Array("Capacity.MaxLambda", "Capacity.InitialLambda", "Capacity.ToleranceLambda", "Capacity.MaxRetries", "Capacity.BaseLoadSteps")
        table.Formula = configured: SetValue table, "Capacity.SolutionStrategy", "UltimateStrain"
        Select Case CStr(key)
            Case "Capacity.MaxLambda": value = 0.125
            Case "Capacity.InitialLambda": value = 0.25
            Case "Capacity.ToleranceLambda": value = 0.1
            Case "Capacity.MaxRetries": value = 0
            Case "Capacity.BaseLoadSteps": value = 3
        End Select
        SetValue table, CStr(key), value
        Set changed = CapacityNumericRun(stats, prefix & ".inactive." & CStr(key), section, provider, profiles, units)
        CheckClose stats, prefix & ".inactive." & CStr(key) & ".limit", changed.LambdaCapacity, baseline.LambdaCapacity, 0#
        Check stats, prefix & ".inactive." & CStr(key) & ".diagnostic", changed.DiagnosticLog = baseline.DiagnosticLog
    Next key
    table.Formula = configured
End Sub

' Изолирует деформационный допуск в настоящем batch-пути: демпфирование 0.5
' приближает критерий постепенно, а равновесие имеет свои явно заданные допуски.
' Изменение только ToleranceStrain должно менять остановку Newton; физические
' пределы и проверка конечных усилий не отключаются и не пересчитываются тестом.
Private Sub CheckCapacityStrainToleranceBatch(ByRef stats As TConfigTestStats, ByVal prefix As String, _
        ByVal table As Object, ByRef configured As Variant, ByVal section As CSectionModel, _
        ByVal provider As CMaterialModelProvider, ByVal profiles As CCalculationProfileCatalog, ByVal units As CUnitSystem)
    Dim coarse As CCapacityResult, tight As CCapacityResult
    table.Formula = configured
    SetValue table, "Capacity.SolutionStrategy", "UltimateStrain"
    SetValue table, "Capacity.SolverMaxIterations", 120
    SetValue table, "Solver.DampingInitial", 0.5
    SetValue table, "Solver.ToleranceN", 1000#
    SetValue table, "Solver.ToleranceMx", 100000#
    SetValue table, "Solver.ToleranceMy", 100000#
    SetValue table, "Capacity.ToleranceStrain", 0.0001
    Set coarse = CapacityNumericRun(stats, prefix & ".Capacity.ToleranceStrain.batchCoarse", section, provider, profiles, units)
    SetValue table, "Capacity.ToleranceStrain", 0.000000001
    Set tight = CapacityNumericRun(stats, prefix & ".Capacity.ToleranceStrain.batchTight", section, provider, profiles, units)
    Check stats, prefix & ".Capacity.ToleranceStrain.batchPoints", coarse.SearchResult.HasLimitPoint And tight.SearchResult.HasLimitPoint
    If Not coarse.SearchResult.HasLimitPoint Or Not tight.SearchResult.HasLimitPoint Then Exit Sub
    CheckState stats, prefix & ".Capacity.ToleranceStrain.batchCoarse", coarse.StateResult, 1000#, 100000#
    CheckState stats, prefix & ".Capacity.ToleranceStrain.batchTight", tight.StateResult, 1000#, 100000#
    Check stats, prefix & ".Capacity.ToleranceStrain.batchEffect", _
        CapacityNumericRecordCount(tight.DiagnosticLog, "ultimatePath iter=") > _
        CapacityNumericRecordCount(coarse.DiagnosticLog, "ultimatePath iter=") Or _
        Abs(tight.LambdaCapacity - coarse.LambdaCapacity) > 0.000000000001
    table.Formula = configured
End Sub

' Проверяет неактивную Capacity в реальном crack-профиле: изменение любого
' из семи допустимых параметров не меняет Formation/current/width и число solve.
' Ошибочные значения рассматриваются отдельным общим preflight, не скрываются.
Private Sub CheckCapacityNumericIsolation(ByRef stats As TConfigTestStats, ByVal table As Object, _
        ByRef configured As Variant, ByVal position As Long, ByVal section As CSectionModel, _
        ByVal provider As CMaterialModelProvider, ByVal profiles As CCalculationProfileCatalog, ByVal units As CUnitSystem)
    Dim baseline As CBatchSectionCalculator, changed As CBatchSectionCalculator, keys As Variant, values As Variant
    Dim index As Long, prefix As String, first As CCrackResult, second As CCrackResult
    keys = Array("Capacity.SolverMaxIterations", "Capacity.ToleranceStrain", "Capacity.MaxLambda", _
        "Capacity.InitialLambda", "Capacity.ToleranceLambda", "Capacity.MaxRetries", "Capacity.BaseLoadSteps")
    values = Array(1, 0.000000001, 0.125, 0.25, 0.1, 0, 3)
    prefix = "audit03.capacityNumeric.isolation.p" & CStr(position)
    table.Formula = configured
    Set baseline = RunConfigBatch(stats, prefix & ".baseline", "PR2", section, provider, profiles, units)
    Set first = baseline.ResultAt(1).CrackResult
    Check stats, prefix & ".formationPoint", first.Formation.HasLimitPoint
    For index = LBound(keys) To UBound(keys)
        table.Formula = configured: SetValue table, CStr(keys(index)), values(index)
        Set changed = RunConfigBatch(stats, prefix & "." & CStr(keys(index)), "PR2", section, provider, profiles, units)
        Set second = changed.ResultAt(1).CrackResult
        CheckClose stats, prefix & "." & CStr(keys(index)) & ".lambda", second.Formation.LambdaCrc, first.Formation.LambdaCrc, 0#
        Check stats, prefix & "." & CStr(keys(index)) & ".method", second.Formation.FormationMethod = first.Formation.FormationMethod
        Check stats, prefix & "." & CStr(keys(index)) & ".solveCount", changed.SolverCallCount = baseline.SolverCallCount
        Check stats, prefix & "." & CStr(keys(index)) & ".status", changed.ResultAt(1).Status = baseline.ResultAt(1).Status
        CheckClose stats, prefix & "." & CStr(keys(index)) & ".width", second.Width.CrackWidth, first.Width.CrackWidth, 0#
        CheckClose stats, prefix & "." & CStr(keys(index)) & ".sigmaS", second.Width.SigmaS, first.Width.SigmaS, 0#
    Next index
    table.Formula = configured
End Sub

' Передает Config в существующий Capacity-контекст и оценивает сохраненную
' физическую плоскость, чтобы критический материал/знак не были выдуманы.
' Для rLimit=0.001 грубый допуск разрешает остановку, строгий требует уточнения.
' Сам тест не финализирует эту контрольную невязку как физический результат.
Private Sub CheckCapacityStrainTolerance(ByRef stats As TConfigTestStats, ByVal prefix As String, _
        ByVal table As Object, ByVal section As CSectionModel, ByVal provider As CMaterialModelProvider, _
        ByVal units As CUnitSystem, ByVal point As CSectionStateResult)
    Dim settings As CSystemSettingsReader, context As CCapacitySolver, material As CMaterialModelSpec
    Dim concrete As CMaterialDiagram, steel As CMaterialDiagram, path As CLoadPathVector, solver As CSectionSolver
    Dim r1 As Double, r2 As Double, rLimit As Double, lambdaValue As Double, state As String
    Set material = point.MaterialSpec
    Set concrete = provider.ConcreteMaterialFromSpec(material): Set steel = provider.SteelMaterialFromSpec(material)
    Set context = New CCapacitySolver
    context.ConcreteCompressionLimit = concrete.UltimateCompressionStrain
    context.ConcreteTensionLimit = concrete.UltimateTensionStrain
    context.ConcreteTensionLimitEnabled = provider.ConcreteTensionLimitEnabledFromSpec(material)
    context.SteelCompressionLimit = Abs(steel.UltimateCompressionStrain)
    context.SteelTensionLimit = Abs(steel.UltimateTensionStrain)
    Set path = New CLoadPathVector: path.Initialize 0#, FIXTURE_N, 0#, FIXTURE_MX, 0#, FIXTURE_MY
    Set settings = New CSystemSettingsReader
    SetValue table, "Capacity.ToleranceStrain", 0.0001
    settings.LoadFromWorkbook ThisWorkbook: context.ApplySettings settings, units
    Check stats, prefix & ".Capacity.ToleranceStrain.evaluate", _
        context.LimitSearchEvaluateUltimateResidual(section, concrete, steel, path, point.Epsilon0, point.KappaX, point.KappaY, _
            solver, r1, r2, rLimit, lambdaValue, state)
    Check stats, prefix & ".Capacity.ToleranceStrain.coarseCriterion", context.LimitSearchUltimateResidualReached(0#, 0#, 0.001)
    SetValue table, "Capacity.ToleranceStrain", 0.000000001
    settings.LoadFromWorkbook ThisWorkbook: context.ApplySettings settings, units
    Check stats, prefix & ".Capacity.ToleranceStrain.tightCriterion", Not context.LimitSearchUltimateResidualReached(0#, 0#, 0.001)
End Sub

' Ошибки чтения и допустимого диапазона должны дать InputErr без solver-а.
' Сообщение принадлежит владельцу настройки и содержит реальную ячейку;
' после возврата значения тот же reader/batch путь должен снова найти предел.
Private Sub CheckCapacityNumericInput(ByRef stats As TConfigTestStats, ByVal table As Object, _
        ByRef configured As Variant, ByVal position As Long, ByVal section As CSectionModel, _
        ByVal provider As CMaterialModelProvider, ByVal profiles As CCalculationProfileCatalog, ByVal units As CUnitSystem)
    Dim keys As Variant, key As Variant, bad As Variant, invalid As Variant, index As Long
    Dim prefix As String, message As String, cell As Object, batch As CBatchSectionCalculator
    keys = Array("Capacity.SolverMaxIterations", "Capacity.ToleranceStrain", "Capacity.MaxLambda", _
        "Capacity.InitialLambda", "Capacity.ToleranceLambda", "Capacity.MaxRetries", "Capacity.BaseLoadSteps")
    For Each key In keys
        invalid = Array(vbNullString, "TODO", "abc", CVErr(2042), -1#, 0#, 1.5)
        For index = LBound(invalid) To UBound(invalid)
            If CStr(key) = "Capacity.MaxRetries" And index = 5 Then GoTo NextInvalid
            If index = 6 And CStr(key) <> "Capacity.MaxRetries" And CStr(key) <> "Capacity.BaseLoadSteps" And _
                CStr(key) <> "Capacity.SolverMaxIterations" Then GoTo NextInvalid
            table.Formula = configured
            If CStr(key) = "Capacity.ToleranceStrain" Then
                SetValue table, "Capacity.SolutionStrategy", "UltimateStrain"
            Else
                SetValue table, "Capacity.SolutionStrategy", "LoadMultiplier"
            End If
            Set cell = ValueCell(table, CStr(key), 2): cell.Value2 = invalid(index)
            prefix = "audit03.capacityNumeric.invalid.p" & CStr(position) & "." & CStr(key) & ".v" & CStr(index)
            Set batch = RunConfigBatch(stats, prefix, "PR1", section, provider, profiles, units)
            message = batch.ResultAt(1).OverallMeta.ResultComment
            Check stats, prefix & ".inputErr", batch.ResultAt(1).Status = "InputErr"
            Check stats, prefix & ".noSolve", batch.SolverCallCount = 0
            Check stats, prefix & ".key", InStr(1, message, CStr(key), vbBinaryCompare) > 0
            Check stats, prefix & ".cell", InStr(1, message, cell.Address(False, False), vbTextCompare) > 0
            Check stats, prefix & ".action", InStr(1, message, "Введите", vbTextCompare) > 0 Or _
                InStr(1, message, "задайте", vbTextCompare) > 0 Or InStr(1, message, "Исправьте", vbTextCompare) > 0
            LogLine stats, "CAPACITY_INPUT_MESSAGE: " & prefix & "|" & message
NextInvalid:
        Next index
        table.Formula = configured
        Set batch = RunConfigBatch(stats, "audit03.capacityNumeric.recovery.p" & CStr(position) & "." & CStr(key), _
            "PR1", section, provider, profiles, units)
        Check stats, "audit03.capacityNumeric.recovery.p" & CStr(position) & "." & CStr(key), _
            batch.ResultAt(1).StrengthResult.Capacity.SearchResult.HasLimitPoint
    Next key
    table.Formula = configured
End Sub

' Возвращает опубликованный результат реального batch и сохраняет счетчики
' фактических Search-проб для анализа причин, не назначая статус по журналу.
Private Function CapacityNumericRun(ByRef stats As TConfigTestStats, ByVal prefix As String, _
        ByVal section As CSectionModel, ByVal provider As CMaterialModelProvider, _
        ByVal profiles As CCalculationProfileCatalog, ByVal units As CUnitSystem) As CCapacityResult
    Dim batch As CBatchSectionCalculator, result As CCapacityResult
    Set batch = RunConfigBatch(stats, prefix, "PR1", section, provider, profiles, units)
    Set result = batch.ResultAt(1).StrengthResult.Capacity
    LogLine stats, "CAPACITY_NUMERIC: " & prefix & "|lambda=" & FormatNumberInvariant(result.LambdaCapacity) & _
        "|status=" & result.Status & "|code=" & CStr(result.ResultMeta.ResultCode) & _
        "|probes=" & CStr(CapacityNumericRecordCount(result.DiagnosticLog, "probe=")) & _
        "|ultimate=" & CStr(CapacityNumericRecordCount(result.DiagnosticLog, "ultimatePath iter="))
    LogLine stats, "CAPACITY_DIAGNOSTIC: " & prefix & vbCrLf & result.DiagnosticLog
    Set CapacityNumericRun = result
End Function

' Неуспешный Search не получает подставную точку. У принятой точки проверяются
' равновесие и соответствие полному заранее заданному lambda-вектору нагрузки.
Private Sub CheckCapacityNumericPoint(ByRef stats As TConfigTestStats, ByVal prefix As String, ByVal result As CCapacityResult)
    Check stats, prefix & ".point", result.SearchResult.HasLimitPoint
    If Not result.SearchResult.HasLimitPoint Then Exit Sub
    CheckState stats, prefix, result.StateResult
    CheckClose stats, prefix & ".pathN", result.StateResult.Nint, FIXTURE_N * result.LambdaCapacity, FORCE_TOLERANCE
    CheckClose stats, prefix & ".pathMx", result.StateResult.Mxint, FIXTURE_MX * result.LambdaCapacity, MOMENT_TOLERANCE
    CheckClose stats, prefix & ".pathMy", result.StateResult.Myint, FIXTURE_MY * result.LambdaCapacity, MOMENT_TOLERANCE
End Sub

' Считает реальные диагностические записи; это только test oracle маршрута,
' а инженерная классификация всегда проверяется по InternalStatus/ResultCode.
Private Function CapacityNumericRecordCount(ByVal diagnostic As String, ByVal marker As String) As Long
    Dim line As Variant
    For Each line In Split(diagnostic, vbCrLf)
        If InStr(1, CStr(line), marker, vbBinaryCompare) > 0 Then CapacityNumericRecordCount = CapacityNumericRecordCount + 1
    Next line
End Function

' Возвращает первый ненулевой lambda из подробной записи фактически решенной
' Capacity-пробы. Val читает invariant-формат журнала независимо от locale.
Private Function CapacityNumericFirstLambda(ByVal diagnostic As String) As Double
    Dim line As Variant, at As Long, value As Double
    For Each line In Split(diagnostic, vbCrLf)
        If Left$(CStr(line), 6) = "probe=" Then
            at = InStr(1, CStr(line), "; lambda=", vbBinaryCompare)
            If at > 0 Then
                value = Val(Mid$(CStr(line), at + 9))
                If value > 0# Then CapacityNumericFirstLambda = value: Exit Function
            End If
        End If
    Next line
End Function

' Возвращает последний фактически решенный lambda без подмены ненайденного
' физического предела: так проверяется достижение технической границы поиска.
Private Function CapacityNumericLastLambda(ByVal diagnostic As String) As Double
    Dim line As Variant, at As Long
    For Each line In Split(diagnostic, vbCrLf)
        If Left$(CStr(line), 6) = "probe=" Then
            at = InStr(1, CStr(line), "; lambda=", vbBinaryCompare)
            If at > 0 Then CapacityNumericLastLambda = Val(Mid$(CStr(line), at + 9))
        End If
    Next line
End Function

' ==================== ДЛЯ ТЕСТОВ: НАСТРОЙКИ УСТОЙЧИВОСТИ ====================

' Меняет все 17 полей устойчивости через настоящий Config. Проверяет обе
' главные плоскости, активные/неактивные ветви, длины в разных единицах,
' ошибки и восстановление. Именованная таблица испытывается также после переноса.
Public Function RunAudit03StabilityConfigTests(Optional ByRef passed As Long = 0, _
        Optional ByRef failed As Long = 0, Optional ByVal lengthRangeOnly As Boolean = False) As String
    Dim stats As TConfigTestStats, originalRange As Object, table As Object, movedRange As Object
    Dim unitRange As Object, profileRange As Object, savedName As String, position As Long
    Dim savedSystem As Variant, savedMoved As Variant, savedUnits As Variant, savedProfiles As Variant
    Dim configured As Variant, configuredUnits As Variant, configuredProfiles As Variant
    Dim settings As CSystemSettingsReader, units As CUnitSystem, provider As CMaterialModelProvider
    Dim section As CSectionModel
    On Error GoTo FailedRun
    Set originalRange = ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange
    savedName = ThisWorkbook.Names.Item("rngSystemSettings").RefersTo
    Set movedRange = originalRange.Worksheet.Range("CH800").Resize(originalRange.Rows.Count, originalRange.Columns.Count)
    Set unitRange = ThisWorkbook.Names.Item("rngUnitSettings").RefersToRange
    Set profileRange = ThisWorkbook.Names.Item("rngCalculationProfiles").RefersToRange
    savedSystem = originalRange.Formula: savedMoved = movedRange.Formula
    savedUnits = unitRange.Formula: savedProfiles = profileRange.Formula
    ConfigureSearchFixture originalRange, unitRange, profileRange
    ConfigureStabilityFixture originalRange, profileRange
    configured = originalRange.Formula: configuredUnits = unitRange.Formula
    configuredProfiles = profileRange.Formula
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    Set provider = New CMaterialModelProvider: provider.Initialize settings, units
    Set section = SearchFixtureSection()
    For position = 1 To 2
        If position = 1 Then Set table = originalRange Else Set table = movedRange
        table.Formula = configured
        ThisWorkbook.Names.Item("rngSystemSettings").RefersTo = "=" & table.Address(True, True, 1, True)
        If Not lengthRangeOnly Then
            CheckStabilityConfigEffects stats, table, configured, position, section, provider, units
            CheckStabilityConfigSP35 stats, table, configured, position, section, provider, units
            CheckStabilityConfigIsolation stats, table, configured, position, section, provider, units
            CheckStabilityConfigInput stats, table, configured, position, section, provider, units
        End If
        CheckStabilityLengthRange stats, table, configured, position, section, provider, units
        If Not lengthRangeOnly Then CheckStabilityConfigUnits stats, table, configured, unitRange, configuredUnits, position, section, provider, units
        profileRange.Formula = configuredProfiles
    Next position
    If Not lengthRangeOnly Then CheckStabilityStandaloneInput stats, section, provider
    GoTo Restore
FailedRun:
    Check stats, "audit03.stabilityConfig.runtime." & CStr(Err.Number) & "." & Err.Description, False
Restore:
    On Error Resume Next
    ThisWorkbook.Names.Item("rngSystemSettings").RefersTo = savedName
    If Not originalRange Is Nothing Then originalRange.Formula = savedSystem
    If Not movedRange Is Nothing Then movedRange.Formula = savedMoved
    If Not unitRange Is Nothing Then unitRange.Formula = savedUnits
    If Not profileRange Is Nothing Then profileRange.Formula = savedProfiles
    On Error GoTo 0
    LogLine stats, "TOTAL_AUDIT03_STABILITY_CONFIG: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    passed = stats.Passed: failed = stats.Failed
    RunAudit03StabilityConfigTests = stats.Report
End Function

' Запускает только крайние пары расчетной длины, сохраняя те же настоящие
' Config/reader/batch/writer и восстановление, что и полный Config-набор.
Public Function RunAudit03StabilityLengthRangeTests() As String
    Dim passed As Long, failed As Long
    RunAudit03StabilityLengthRangeTests = RunAudit03StabilityConfigTests(passed, failed, True)
End Function

' Фиксирует исходные значения до серии: длинный круг с обеими моментными
' компонентами дает наблюдаемые Ncr/eta; случайные эксцентриситеты независимы
' от длины. Второй профиль проверяет настоящий DirectState без устойчивости.
Private Sub ConfigureStabilityFixture(ByVal table As Object, ByVal profiles As Object)
    Dim keys As Variant, values As Variant, i As Long
    keys = StabilityConfigKeys()
    values = Array("SP63", 20000#, 1#, 1#, "Determinate", 1#, 1#, "User", 100#, 80#, _
        "BothPlanes", "Auto", 0.7, 0.15, 1.5, 1#, 0.7)
    For i = LBound(keys) To UBound(keys)
        SetValue table, CStr(keys(i)), values(i)
    Next i
    SetProfileValue profiles, "Calculation.Stability.Enabled", "Yes", "PR1"
    SetProfileValue profiles, "MaterialModel.Stability.ValueSet", "ULS(I)", "PR1"
    SetProfileValue profiles, "Calculation.Strength.DirectState", "No", "PR1"
    SetProfileValue profiles, "Calculation.Strength.Capacity", "No", "PR1"
    SetProfileValue profiles, "Calculation.Crack.Width", "No", "PR1"
    SetProfileValue profiles, "Calculation.Strength.DirectState", "Yes", "PR2"
    SetProfileValue profiles, "Calculation.Crack.Width", "No", "PR2"
End Sub

' Возвращает именно фактические поля Config в устойчивом порядке для эффектов,
' input-контрактов и реестра; заголовки таблицы и computed-поля сюда не входят.
Private Function StabilityConfigKeys() As Variant
    StabilityConfigKeys = Array("Stability.Code", "Stability.ElementLength", "Stability.Mu1", "Stability.Mu2", _
        "Stability.SystemType", "Stability.ZeroMomentEccentricitySign1", "Stability.ZeroMomentEccentricitySign2", _
        "Stability.AccidentalEccentricityMode", "Stability.AccidentalEccentricityUser1", "Stability.AccidentalEccentricityUser2", _
        "Stability.AccidentalEccentricityPlanes", "Stability.PhiLMode", "Stability.SP63.Ks", _
        "Stability.SP63.DeltaEMin", "Stability.SP63.DeltaEMax", "Stability.SP35.PhiP", "Stability.SP35.NOverNcrLimit")
End Function

' Проверяет наблюдаемые величины обеих плоскостей. Обратный квадрат длины
' испытывается при неизменной жесткости, знаки - при нулевых моментах,
' границы delta и эксцентриситеты - отдельно от смены нормативной ветви.
Private Sub CheckStabilityConfigEffects(ByRef stats As TConfigTestStats, ByVal table As Object, _
        ByRef configured As Variant, ByVal position As Long, ByVal section As CSectionModel, _
        ByVal provider As CMaterialModelProvider, ByVal units As CUnitSystem)
    Dim baseline As CStabilityResult, changed As CStabilityResult
    Dim prefix As String, batch As CBatchSectionCalculator
    prefix = "audit03.stabilityConfig.p" & CStr(position)
    table.Formula = configured
    Set batch = StabilityConfigRun(stats, prefix & ".base", section, provider, units)
    Set baseline = batch.ResultAt(1).StabilityResult
    Check stats, prefix & ".base.branch", baseline.Branch = "SP63"
    Check stats, prefix & ".base.positiveNcr", baseline.Ncr1 > 0# And baseline.Ncr2 > 0#
    CheckClose stats, prefix & ".base.l01", baseline.EffectiveLength1, 20000#, 0.0000001
    CheckClose stats, prefix & ".base.e1", baseline.Eccentricity1, 120#, 0.0000001
    CheckClose stats, prefix & ".base.e2", baseline.Eccentricity2, 90#, 0.0000001

    Dim key As Variant
    For Each key In Array("Stability.ElementLength", "Stability.Mu1", "Stability.Mu2")
        table.Formula = configured
        If CStr(key) = "Stability.ElementLength" Then SetValue table, CStr(key), 10000# Else SetValue table, CStr(key), 2#
        Set batch = StabilityConfigRun(stats, prefix & "." & CStr(key), section, provider, units)
        Set changed = batch.ResultAt(1).StabilityResult
        If CStr(key) = "Stability.ElementLength" Then
            CheckClose stats, prefix & ".L.l01", changed.EffectiveLength1, 10000#, 0.0000001
            CheckClose stats, prefix & ".L.Ncr1", changed.Ncr1, baseline.Ncr1 * 4#, 0.00001
            CheckClose stats, prefix & ".L.Ncr2", changed.Ncr2, baseline.Ncr2 * 4#, 0.00001
        ElseIf CStr(key) = "Stability.Mu1" Then
            CheckClose stats, prefix & ".mu1.l01", changed.EffectiveLength1, 40000#, 0.0000001
            CheckClose stats, prefix & ".mu1.Ncr1", changed.Ncr1, baseline.Ncr1 / 4#, 0.00001
            CheckClose stats, prefix & ".mu1.Ncr2.unchanged", changed.Ncr2, baseline.Ncr2, 0.00001
        Else
            CheckClose stats, prefix & ".mu2.l02", changed.EffectiveLength2, 40000#, 0.0000001
            CheckClose stats, prefix & ".mu2.Ncr2", changed.Ncr2, baseline.Ncr2 / 4#, 0.00001
            CheckClose stats, prefix & ".mu2.Ncr1.unchanged", changed.Ncr1, baseline.Ncr1, 0.00001
        End If
    Next key
    table.Formula = configured: SetValue table, "Stability.SystemType", "Indeterminate"
    Set batch = StabilityConfigRun(stats, prefix & ".SystemType", section, provider, units)
    Set changed = batch.ResultAt(1).StabilityResult
    CheckClose stats, prefix & ".system.e1", changed.Eccentricity1, 100#, 0.0000001
    CheckClose stats, prefix & ".system.e2", changed.Eccentricity2, 80#, 0.0000001
    Check stats, prefix & ".system.stiffnessEffect", changed.StiffnessD1 > baseline.StiffnessD1

    Dim axis As Long
    For axis = 1 To 2
        table.Formula = configured: SetValue table, "Stability.ZeroMomentEccentricitySign" & CStr(axis), -1#
        Set batch = StabilityConfigRun(stats, prefix & ".zeroSign" & CStr(axis), section, provider, units, -10000#, 0#, 0#)
        Set changed = batch.ResultAt(1).StabilityResult
        CheckClose stats, prefix & ".zeroSign.ea1." & CStr(axis), changed.AccidentalEcc1, IIf(axis = 1, -100#, 100#), 0.0000001
        CheckClose stats, prefix & ".zeroSign.ea2." & CStr(axis), changed.AccidentalEcc2, IIf(axis = 2, -80#, 80#), 0.0000001
        Set batch = StabilityConfigRun(stats, prefix & ".nonzeroSignInactive" & CStr(axis), section, provider, units)
        CheckClose stats, prefix & ".nonzeroSign.e1." & CStr(axis), batch.ResultAt(1).StabilityResult.Eccentricity1, baseline.Eccentricity1, 0.0000001
        CheckClose stats, prefix & ".nonzeroSign.e2." & CStr(axis), batch.ResultAt(1).StabilityResult.Eccentricity2, baseline.Eccentricity2, 0.0000001
    Next axis
    Dim mode As Variant
    For Each mode In Array("AutoWithL", "AutoWithMuL", "AutoWith" & ChrW$(&H3BC) & "L")
        table.Formula = configured
        SetValue table, "Stability.Mu1", 2#: SetValue table, "Stability.Mu2", 3#
        SetValue table, "Stability.AccidentalEccentricityMode", CStr(mode)
        Set batch = StabilityConfigRun(stats, prefix & ".mode." & CStr(mode), section, provider, units)
        Set changed = batch.ResultAt(1).StabilityResult
        If CStr(mode) = "AutoWithL" Then
            CheckClose stats, prefix & ".mode.L.ea1", changed.AccidentalEcc1, 20000# / 600#, 0.0000001
            CheckClose stats, prefix & ".mode.L.ea2", changed.AccidentalEcc2, 20000# / 600#, 0.0000001
        Else
            CheckClose stats, prefix & ".mode.muL.ea1." & CStr(mode), changed.AccidentalEcc1, 40000# / 600#, 0.0000001
            CheckClose stats, prefix & ".mode.muL.ea2." & CStr(mode), changed.AccidentalEcc2, 60000# / 600#, 0.0000001
        End If
    Next mode
    For axis = 1 To 2
        table.Formula = configured: SetValue table, "Stability.AccidentalEccentricityUser" & CStr(axis), 150#
        Set batch = StabilityConfigRun(stats, prefix & ".User" & CStr(axis), section, provider, units)
        Set changed = batch.ResultAt(1).StabilityResult
        If axis = 1 Then
            CheckClose stats, prefix & ".user1.ea", changed.AccidentalEcc1, 150#, 0.0000001
            CheckClose stats, prefix & ".user1.other", changed.AccidentalEcc2, 80#, 0.0000001
        Else
            CheckClose stats, prefix & ".user2.ea", changed.AccidentalEcc2, 150#, 0.0000001
            CheckClose stats, prefix & ".user2.other", changed.AccidentalEcc1, 100#, 0.0000001
        End If
    Next axis
    table.Formula = configured: SetValue table, "Stability.AccidentalEccentricityPlanes", "OnlyMomentPlane"
    Set batch = StabilityConfigRun(stats, prefix & ".onlyMoment", section, provider, units, -10000#, 200000#, 0#)
    Set changed = batch.ResultAt(1).StabilityResult
    CheckClose stats, prefix & ".planes.inactive", changed.AccidentalEcc2, 0#, 0.0000001
    CheckClose stats, prefix & ".planes.active", changed.AccidentalEcc1, 100#, 0.0000001
    SetValue table, "Stability.AccidentalEccentricityPlanes", "BothPlanes"
    Set batch = StabilityConfigRun(stats, prefix & ".bothPlanes", section, provider, units, -10000#, 200000#, 0#)
    CheckClose stats, prefix & ".planes.both", batch.ResultAt(1).StabilityResult.AccidentalEcc2, 80#, 0.0000001

    table.Formula = configured: SetValue table, "Stability.PhiLMode", "PhiL2"
    Set batch = StabilityConfigRun(stats, prefix & ".PhiL2", section, provider, units)
    Set changed = batch.ResultAt(1).StabilityResult
    CheckClose stats, prefix & ".phi.two1", changed.PhiL1, 2#, 0.0000001
    CheckClose stats, prefix & ".phi.two2", changed.PhiL2, 2#, 0.0000001
    Check stats, prefix & ".phi.NcrEffect", changed.Ncr1 < baseline.Ncr1 And changed.Ncr2 < baseline.Ncr2
    table.Formula = configured
    Set batch = StabilityConfigRun(stats, prefix & ".sustained", section, provider, units, -10000#, 200000#, 100000#, -5000#, 100000#, 50000#)
    CheckClose stats, prefix & ".phi.sustained1", batch.ResultAt(1).StabilityResult.PhiL1, 1.5, 0.0000001
    CheckClose stats, prefix & ".phi.sustained2", batch.ResultAt(1).StabilityResult.PhiL2, 1.5, 0.0000001
    table.Formula = configured: SetValue table, "Stability.SP63.Ks", 1.4
    Set batch = StabilityConfigRun(stats, prefix & ".Ks", section, provider, units)
    Set changed = batch.ResultAt(1).StabilityResult
    CheckClose stats, prefix & ".Ks.used", changed.Ks1, 1.4, 0.0000001
    Check stats, prefix & ".Ks.DEffect", changed.StiffnessD1 > baseline.StiffnessD1 And changed.StiffnessD2 > baseline.StiffnessD2
    table.Formula = configured: SetValue table, "Stability.SP63.DeltaEMin", 0.8
    Set batch = StabilityConfigRun(stats, prefix & ".deltaMin", section, provider, units)
    CheckClose stats, prefix & ".deltaMin.used1", batch.ResultAt(1).StabilityResult.Delta1, 0.8, 0.0000001
    CheckClose stats, prefix & ".deltaMin.used2", batch.ResultAt(1).StabilityResult.Delta2, 0.8, 0.0000001
    table.Formula = configured: SetValue table, "Stability.SP63.DeltaEMax", 0.2
    Set batch = StabilityConfigRun(stats, prefix & ".deltaMax", section, provider, units)
    CheckClose stats, prefix & ".deltaMax.used1", batch.ResultAt(1).StabilityResult.Delta1, 0.2, 0.0000001
    CheckClose stats, prefix & ".deltaMax.used2", batch.ResultAt(1).StabilityResult.Delta2, 0.2, 0.0000001
    table.Formula = configured: SetValue table, "Stability.SP63.DeltaEMin", 0#: SetValue table, "Stability.SP63.DeltaEMax", 0#
    Set batch = StabilityConfigRun(stats, prefix & ".deltaDisabled", section, provider, units)
    CheckClose stats, prefix & ".delta.disabled1", batch.ResultAt(1).StabilityResult.Delta1, 0.4, 0.0000001
    table.Formula = configured
End Sub

' Переключает на реальные eta/table ветви СП 35. PhiP меняет жесткость только
' eta-ветви, NOverNcrLimit - только проверочный запас. В табличной ветви
' допустимые параметры eta и СП 63 не должны подменять сохраненный результат.
Private Sub CheckStabilityConfigSP35(ByRef stats As TConfigTestStats, ByVal table As Object, _
        ByRef configured As Variant, ByVal position As Long, ByVal section As CSectionModel, _
        ByVal provider As CMaterialModelProvider, ByVal units As CUnitSystem)
    Dim batch As CBatchSectionCalculator, baseline As CStabilityResult, changed As CStabilityResult
    Dim prefix As String, nValue As Double, tableConfig As Variant, key As Variant, mode As Variant
    prefix = "audit03.stabilityConfig.p" & CStr(position) & ".SP35"
    table.Formula = configured: SetValue table, "Stability.Code", "SP35"
    Set batch = StabilityConfigRun(stats, prefix & ".etaBase", section, provider, units)
    Set baseline = batch.ResultAt(1).StabilityResult
    Check stats, prefix & ".code", baseline.Code = "SP35"
    Check stats, prefix & ".eta.branch", baseline.Branch = "SP35-eta"
    For Each mode In Array("AutoWithL", "AutoWithMuL")
        SetValue table, "Stability.AccidentalEccentricityMode", CStr(mode)
        SetValue table, "Stability.Mu1", 2#: SetValue table, "Stability.Mu2", 3#
        Set batch = StabilityConfigRun(stats, prefix & ".mode." & CStr(mode), section, provider, units)
        Set changed = batch.ResultAt(1).StabilityResult
        If CStr(mode) = "AutoWithL" Then
            CheckClose stats, prefix & ".mode.L.ea1", changed.AccidentalEcc1, 50#, 0.0000001
            CheckClose stats, prefix & ".mode.L.ea2", changed.AccidentalEcc2, 50#, 0.0000001
        Else
            CheckClose stats, prefix & ".mode.muL.ea1", changed.AccidentalEcc1, 100#, 0.0000001
            CheckClose stats, prefix & ".mode.muL.ea2", changed.AccidentalEcc2, 150#, 0.0000001
        End If
    Next mode
    table.Formula = configured: SetValue table, "Stability.Code", "SP35"
    SetValue table, "Stability.PhiLMode", "PhiL2"
    Set batch = StabilityConfigRun(stats, prefix & ".PhiL2", section, provider, units)
    CheckClose stats, prefix & ".phi.used", batch.ResultAt(1).StabilityResult.PhiL1, 2#, 0.0000001
    Check stats, prefix & ".phi.NcrEffect", batch.ResultAt(1).StabilityResult.Ncr1 < baseline.Ncr1
    SetValue table, "Stability.PhiLMode", "Auto"
    SetValue table, "Stability.SP35.PhiP", 2#
    Set batch = StabilityConfigRun(stats, prefix & ".PhiP", section, provider, units)
    Set changed = batch.ResultAt(1).StabilityResult
    CheckClose stats, prefix & ".PhiP.used", changed.PhiP1, 2#, 0.0000001
    Check stats, prefix & ".PhiP.DEffect", changed.StiffnessD1 > baseline.StiffnessD1 And changed.StiffnessD2 > baseline.StiffnessD2
    SetValue table, "Stability.SP35.PhiP", 1#: SetValue table, "Stability.SP35.NOverNcrLimit", 0.35
    Set batch = StabilityConfigRun(stats, prefix & ".limit", section, provider, units)
    Set changed = batch.ResultAt(1).StabilityResult
    CheckClose stats, prefix & ".limit.reserve", changed.Reserve1, baseline.Reserve1 / 2#, 0.0000001
    CheckClose stats, prefix & ".limit.NcrUnchanged", changed.Ncr1, baseline.Ncr1, 0.00001
    CheckClose stats, prefix & ".limit.etaUnchanged", changed.Eta1, baseline.Eta1, 0.0000001
    nValue = baseline.CriticalForce * 0.5
    SetValue table, "Stability.SP35.NOverNcrLimit", 0.7
    Set batch = StabilityConfigRun(stats, prefix & ".limitPass", section, provider, units, -nValue, nValue * 20#, nValue * 10#)
    Check stats, prefix & ".limit.pass", batch.ResultAt(1).StabilityResult.Status = "OK"
    SetValue table, "Stability.SP35.NOverNcrLimit", 0.3
    Set batch = StabilityConfigRun(stats, prefix & ".limitFail", section, provider, units, -nValue, nValue * 20#, nValue * 10#)
    Check stats, prefix & ".limit.fail", batch.ResultAt(1).StabilityResult.Status = "FAIL"

    table.Formula = configured: SetValue table, "Stability.Code", "SP35"
    SetValue table, "Stability.AccidentalEccentricityUser1", 5#: SetValue table, "Stability.AccidentalEccentricityUser2", 5#
    tableConfig = table.Formula
    Set batch = StabilityConfigRun(stats, prefix & ".tableBase", section, provider, units, -10000#, 0#, 0#)
    Set baseline = batch.ResultAt(1).StabilityResult
    Check stats, prefix & ".table.branch", baseline.Branch = "SP35-table"
    Check stats, prefix & ".table.capacity", baseline.Nultimate1 > 0# And baseline.Nultimate2 > 0#
    SetValue table, "Stability.ElementLength", 3000#
    Set batch = StabilityConfigRun(stats, prefix & ".tableShort", section, provider, units, -10000#, 0#, 0#)
    Set changed = batch.ResultAt(1).StabilityResult
    SetValue table, "Stability.ElementLength", 6000#
    Set batch = StabilityConfigRun(stats, prefix & ".tableLong", section, provider, units, -10000#, 0#, 0#)
    Check stats, prefix & ".table.lengthEffect", changed.Nultimate1 > batch.ResultAt(1).StabilityResult.Nultimate1
    For Each key In Array("Stability.SP35.PhiP", "Stability.SP35.NOverNcrLimit", "Stability.SP63.Ks", _
            "Stability.SP63.DeltaEMin", "Stability.SP63.DeltaEMax", "Stability.PhiLMode")
        table.Formula = tableConfig
        Select Case CStr(key)
            Case "Stability.PhiLMode": SetValue table, CStr(key), "PhiL2"
            Case "Stability.SP63.DeltaEMin": SetValue table, CStr(key), 0.3
            Case Else: SetValue table, CStr(key), 1.2
        End Select
        Set batch = StabilityConfigRun(stats, prefix & ".tableInactive." & CStr(key), section, provider, units, -10000#, 0#, 0#)
        Set changed = batch.ResultAt(1).StabilityResult
        CheckClose stats, prefix & ".tableInactive.Nult1." & CStr(key), changed.Nultimate1, baseline.Nultimate1, 0.0000001
        CheckClose stats, prefix & ".tableInactive.Nult2." & CStr(key), changed.Nultimate2, baseline.Nultimate2, 0.0000001
        CheckClose stats, prefix & ".tableInactive.phi." & CStr(key), changed.PhiValue1, baseline.PhiValue1, 0.0000001
    Next key
    table.Formula = configured
    Set batch = StabilityConfigRun(stats, prefix & ".SP63Base", section, provider, units)
    Set baseline = batch.ResultAt(1).StabilityResult
    SetValue table, "Stability.SP35.PhiP", 2#: SetValue table, "Stability.SP35.NOverNcrLimit", 0.3
    Set batch = StabilityConfigRun(stats, prefix & ".SP35Inactive", section, provider, units)
    CheckClose stats, prefix & ".SP35Inactive.Ncr1", batch.ResultAt(1).StabilityResult.Ncr1, baseline.Ncr1, 0.0000001
    CheckClose stats, prefix & ".SP35Inactive.Ncr2", batch.ResultAt(1).StabilityResult.Ncr2, baseline.Ncr2, 0.0000001
    table.Formula = configured
End Sub

' Допустимые настройки отключенного фильтра не меняют настоящее НДС прочности.
' Плоскость и число solve сравниваются с исходным запуском, не с пустым result.
Private Sub CheckStabilityConfigIsolation(ByRef stats As TConfigTestStats, ByVal table As Object, _
        ByRef configured As Variant, ByVal position As Long, ByVal section As CSectionModel, _
        ByVal provider As CMaterialModelProvider, ByVal units As CUnitSystem)
    Dim baseline As CBatchSectionCalculator, changed As CBatchSectionCalculator, keys As Variant, values As Variant
    Dim i As Long, prefix As String, first As CSectionStateResult, point As CSectionStateResult
    prefix = "audit03.stabilityConfig.p" & CStr(position) & ".inactive"
    table.Formula = configured
    Set baseline = StabilityConfigRun(stats, prefix & ".base", section, provider, units, -10000#, 200000#, 100000#, 0#, 0#, 0#, "PR2")
    Set first = baseline.ResultAt(1).StrengthResult.DirectState.StateResult
    CheckState stats, prefix & ".base.state", first
    keys = StabilityConfigKeys()
    values = Array("SP35", 10000#, 2#, 3#, "Indeterminate", -1#, -1#, "AutoWithL", 150#, 120#, _
        "OnlyMomentPlane", "PhiL2", 1.4, 0.3, 1.2, 2#, 0.3)
    For i = LBound(keys) To UBound(keys)
        table.Formula = configured: SetValue table, CStr(keys(i)), values(i)
        Set changed = StabilityConfigRun(stats, prefix & "." & CStr(keys(i)), section, provider, units, -10000#, 200000#, 100000#, 0#, 0#, 0#, "PR2")
        Set point = changed.ResultAt(1).StrengthResult.DirectState.StateResult
        Check stats, prefix & ".notRequested." & CStr(keys(i)), changed.ResultAt(1).StabilityMeta.InternalStatus = rsNotRequested
        Check stats, prefix & ".solveCount." & CStr(keys(i)), changed.SolverCallCount = baseline.SolverCallCount
        CheckClose stats, prefix & ".e0." & CStr(keys(i)), point.Epsilon0, first.Epsilon0, 0#
        CheckClose stats, prefix & ".kx." & CStr(keys(i)), point.KappaX, first.KappaX, 0#
        CheckClose stats, prefix & ".ky." & CStr(keys(i)), point.KappaY, first.KappaY, 0#
    Next i
    table.Formula = configured
End Sub

' Все обязательные ключи испытываются на blank/TODO/text/CVErr и потерянную
' строку. Численные диапазоны и selectors должны дать InputErr с действующим
' адресом и действием, а не CalcErr, FAIL или скрытое значение по умолчанию.
Private Sub CheckStabilityConfigInput(ByRef stats As TConfigTestStats, ByVal table As Object, _
        ByRef configured As Variant, ByVal position As Long, ByVal section As CSectionModel, _
        ByVal provider As CMaterialModelProvider, ByVal units As CUnitSystem)
    Dim keys As Variant, key As Variant, invalid As Variant, i As Long, prefix As String, cell As Object
    Dim batch As CBatchSectionCalculator, message As String
    keys = StabilityConfigKeys()
    For Each key In keys
        Select Case CStr(key)
            Case "Stability.Code", "Stability.SystemType", "Stability.AccidentalEccentricityMode", _
                    "Stability.AccidentalEccentricityPlanes", "Stability.PhiLMode"
                invalid = Array(vbNullString, "TODO", "abc", CVErr(2042))
            Case "Stability.ZeroMomentEccentricitySign1", "Stability.ZeroMomentEccentricitySign2"
                invalid = Array(vbNullString, "TODO", "abc", CVErr(2042), 0#, 2#, -2#, 0.5)
            Case "Stability.SP63.DeltaEMin", "Stability.SP63.DeltaEMax"
                invalid = Array(vbNullString, "TODO", "abc", CVErr(2042), -1#)
            Case Else
                invalid = Array(vbNullString, "TODO", "abc", CVErr(2042), -1#, 0#)
        End Select
        For i = LBound(invalid) To UBound(invalid)
            table.Formula = configured
            If Left$(CStr(key), 15) = "Stability.SP35." Then SetValue table, "Stability.Code", "SP35"
            Set cell = ValueCell(table, CStr(key), 2): cell.Value2 = invalid(i)
            prefix = "audit03.stabilityConfig.invalid.p" & CStr(position) & "." & CStr(key) & ".v" & CStr(i)
            Set batch = StabilityConfigRun(stats, prefix, section, provider, units, -10000#, 0#, 0#)
            message = batch.ResultAt(1).OverallMeta.ResultComment
            Check stats, prefix & ".inputErr", batch.ResultAt(1).Status = "InputErr"
            Check stats, prefix & ".noSolve", batch.SolverCallCount = 0
            Check stats, prefix & ".key", InStr(1, message, CStr(key), vbBinaryCompare) > 0
            Check stats, prefix & ".cell", InStr(1, message, cell.Address(False, False), vbTextCompare) > 0
            Check stats, prefix & ".action", InStr(1, message, "Введите", vbTextCompare) > 0 Or _
                InStr(1, message, "выберите", vbTextCompare) > 0 Or InStr(1, message, "Исправьте", vbTextCompare) > 0 Or _
                InStr(1, message, "задайте", vbTextCompare) > 0
            LogLine stats, "STABILITY_INPUT_MESSAGE: " & prefix & "|" & message
        Next i
        table.Formula = configured
        Set cell = ValueCell(table, CStr(key), 2): cell.Offset(0, -1).Value2 = "Removed." & CStr(key)
        prefix = "audit03.stabilityConfig.missing.p" & CStr(position) & "." & CStr(key)
        Set batch = StabilityConfigRun(stats, prefix, section, provider, units)
        Check stats, prefix & ".inputErr", batch.ResultAt(1).Status = "InputErr"
        Check stats, prefix & ".key", InStr(1, batch.ResultAt(1).OverallMeta.ResultComment, CStr(key), vbBinaryCompare) > 0
        table.Formula = configured
        Set batch = StabilityConfigRun(stats, "audit03.stabilityConfig.recovery.p" & CStr(position) & "." & CStr(key), section, provider, units)
        Check stats, "audit03.stabilityConfig.recovery.p" & CStr(position) & "." & CStr(key), batch.ResultAt(1).StabilityResult.Ncr1 > 0#
    Next key
    table.Formula = configured: SetValue table, "Stability.SP63.DeltaEMin", 2#
    prefix = "audit03.stabilityConfig.deltaOrder.p" & CStr(position)
    Set batch = StabilityConfigRun(stats, prefix, section, provider, units)
    Check stats, prefix & ".inputErr", batch.ResultAt(1).Status = "InputErr"
    Check stats, prefix & ".cell", InStr(1, batch.ResultAt(1).OverallMeta.ResultComment, _
        ValueCell(table, "Stability.SP63.DeltaEMin", 2).Address(False, False), vbTextCompare) > 0
    table.Formula = configured: SetValue table, "Stability.AccidentalEccentricityMode", "AutoWithL"
    SetValue table, "Stability.AccidentalEccentricityUser1", 0#: SetValue table, "Stability.AccidentalEccentricityUser2", 0#
    Set batch = StabilityConfigRun(stats, "audit03.stabilityConfig.autoUnusedZero.p" & CStr(position), section, provider, units)
    Check stats, "audit03.stabilityConfig.autoUnusedZero.p" & CStr(position), batch.ResultAt(1).StabilityResult.Ncr1 > 0#
    table.Formula = configured
    Set batch = StabilityConfigRun(stats, "audit03.stabilityConfig.tension.p" & CStr(position), section, provider, units, 10000#, 0#, 0#)
    Check stats, "audit03.stabilityConfig.tension.p" & CStr(position), batch.ResultAt(1).StabilityMeta.InternalStatus = rsNotApplicable
    For Each key In Array("Stability.SystemType", "Stability.AccidentalEccentricityMode", "Stability.PhiLMode")
        table.Formula = configured: SetValue table, CStr(key), "abc"
        Set batch = StabilityConfigRun(stats, "audit03.stabilityConfig.tensionInvalid.p" & CStr(position) & "." & CStr(key), _
            section, provider, units, 10000#, 0#, 0#)
        Check stats, "audit03.stabilityConfig.tensionInvalid.p" & CStr(position) & "." & CStr(key), batch.ResultAt(1).Status = "InputErr"
    Next key
    table.Formula = configured
End Sub

' Перевод INPUT Length мм -> м с эквивалентными числами не меняет фильтр.
' Геометрия остается внутренней в мм; OUTPUT Length не переключается.
Private Sub CheckStabilityConfigUnits(ByRef stats As TConfigTestStats, ByVal table As Object, _
        ByRef configured As Variant, ByVal unitRange As Object, ByRef configuredUnits As Variant, _
        ByVal position As Long, ByVal section As CSectionModel, ByVal provider As CMaterialModelProvider, ByVal units As CUnitSystem)
    Dim baseline As CBatchSectionCalculator, changed As CBatchSectionCalculator, prefix As String
    table.Formula = configured: unitRange.Formula = configuredUnits
    prefix = "audit03.stabilityConfig.units.p" & CStr(position)
    Set baseline = StabilityConfigRun(stats, prefix & ".mm", section, provider, units)
    SetValue unitRange, "Length", "m", 2
    SetValue table, "Stability.ElementLength", 20#
    SetValue table, "Stability.AccidentalEccentricityUser1", 0.1
    SetValue table, "Stability.AccidentalEccentricityUser2", 0.08
    Set changed = StabilityConfigRun(stats, prefix & ".m", section, provider, units)
    CheckClose stats, prefix & ".l01", changed.ResultAt(1).StabilityResult.EffectiveLength1, baseline.ResultAt(1).StabilityResult.EffectiveLength1, 0.0000001
    CheckClose stats, prefix & ".ea1", changed.ResultAt(1).StabilityResult.AccidentalEcc1, baseline.ResultAt(1).StabilityResult.AccidentalEcc1, 0.0000001
    CheckClose stats, prefix & ".ea2", changed.ResultAt(1).StabilityResult.AccidentalEcc2, baseline.ResultAt(1).StabilityResult.AccidentalEcc2, 0.0000001
    CheckClose stats, prefix & ".Ncr1", changed.ResultAt(1).StabilityResult.Ncr1, baseline.ResultAt(1).StabilityResult.Ncr1, 0.0000001
    CheckClose stats, prefix & ".Ncr2", changed.ResultAt(1).StabilityResult.Ncr2, baseline.ResultAt(1).StabilityResult.Ncr2, 0.0000001
    Check stats, prefix & ".status", changed.ResultAt(1).Status = baseline.ResultAt(1).Status
    table.Formula = configured: unitRange.Formula = configuredUnits
End Sub

' Проверяет автономный Calculator без Excel-адресов: ранняя ошибка настройки
' сохраняет InputErr и Calculated=False, а не успешный физический результат.
' Это отдельная проверка lifecycle, не повтор формул устойчивости в тесте.
Private Sub CheckStabilityStandaloneInput(ByRef stats As TConfigTestStats, ByVal section As CSectionModel, _
        ByVal provider As CMaterialModelProvider)
    Dim calculator As CStabilityCalculator, meta As CResultMeta, index As Long, systemType As String, length As Double
    For index = 1 To 2
        Set calculator = New CStabilityCalculator
        systemType = "Determinate": length = 20000#
        If index = 1 Then systemType = "abc" Else length = -1#
        calculator.Calculate section, provider, mvsULS, "SP63", -10000#, 200000#, 100000#, 0#, 0#, 0#, _
            length, 1#, 1#, systemType, 1#, 1#, "Auto", "User", "BothPlanes", 100#, 80#, _
            0.7, 0.15, 1.5, 1#, 0.7, Empty
        Set meta = calculator.ResultMeta
        Check stats, "audit03.stabilityConfig.standalone." & CStr(index) & ".inputErr", meta.InternalStatus = rsInvalidInput
        Check stats, "audit03.stabilityConfig.standalone." & CStr(index) & ".notCalculated", Not meta.Calculated
        Check stats, "audit03.stabilityConfig.standalone." & CStr(index) & ".comment", Len(meta.ResultComment) > 0
    Next index
End Sub

' Выполняет reader/units/material/profile/batch pipeline заново для каждого
' случая. После проверки всех writers сохраняет численные данные фильтра;
' устойчивость сама не должна вызывать state-solver. НДС PR2 считается отдельно.
Private Function StabilityConfigRun(ByRef stats As TConfigTestStats, ByVal prefix As String, _
        ByVal section As CSectionModel, ByVal initialProvider As CMaterialModelProvider, ByVal initialUnits As CUnitSystem, _
        Optional ByVal nValue As Double = -10000#, Optional ByVal mxValue As Double = 200000#, _
        Optional ByVal myValue As Double = 100000#, Optional ByVal sustainedN As Double = 0#, _
        Optional ByVal sustainedMx As Double = 0#, Optional ByVal sustainedMy As Double = 0#, _
        Optional ByVal profileId As String = "PR1") As CBatchSectionCalculator
    Dim batch As CBatchSectionCalculator, settings As CSystemSettingsReader, units As CUnitSystem
    Dim provider As CMaterialModelProvider, profiles As CCalculationProfileCatalog, report As CExecutionReport
    Dim reason As String, passed As Long, failed As Long, result As CStabilityResult
    Set batch = New CBatchSectionCalculator: batch.Initialize section, initialProvider
    Set profiles = New CCalculationProfileCatalog: profiles.LoadFromWorkbook ThisWorkbook
    Set batch.ProfileCatalog = profiles
    Set units = initialUnits
    LogLine stats, "RUN: " & prefix
    On Error GoTo InvalidRead
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    Set provider = New CMaterialModelProvider: provider.Initialize settings, units
    batch.Initialize section, provider: Set batch.ProfileCatalog = profiles
    Set report = New CExecutionReport: report.Initialize ThisWorkbook, settings
    Set batch.ExecutionReport = report
    batch.ApplySettings settings, units
    batch.SetSP35Table721 ThisWorkbook.Names.Item("rngSP35Table721").RefersToRange.Value2
    batch.AddCombination prefix, nValue, mxValue, myValue, profileId, "Проверка устойчивости через Config", "Auto"
    batch.AddStabilityDurationLoad prefix, sustainedN, sustainedMx, sustainedMy
    batch.Execute
    GoTo Publish
InvalidRead:
    reason = Err.Description
    On Error GoTo 0
    batch.AddInvalidCombination prefix, profileId, "Ошибочный Config", reason
    batch.Execute
Publish:
    stats.Report = stats.Report & Audit03ValidateConfigBatchResults(batch, units, passed, failed)
    stats.Passed = stats.Passed + passed: stats.Failed = stats.Failed + failed
    If profileId = "PR1" Then Check stats, prefix & ".noStateSolve", batch.SolverCallCount = 0
    Set result = batch.ResultAt(1).StabilityResult
    LogLine stats, "STABILITY_CONFIG: " & prefix & "|status=" & result.Status & "|code=" & result.Code & _
        "|branch=" & result.Branch & "|Ncr1=" & FormatNumberInvariant(result.Ncr1) & "|Ncr2=" & FormatNumberInvariant(result.Ncr2) & _
        "|e1=" & FormatNumberInvariant(result.Eccentricity1) & "|e2=" & FormatNumberInvariant(result.Eccentricity2) & _
        "|L01=" & FormatNumberInvariant(result.EffectiveLength1) & "|L02=" & FormatNumberInvariant(result.EffectiveLength2)
    Set StabilityConfigRun = batch
End Function

' Сохраняет численное значение в журнале без зависимости от десятичного
' разделителя Windows; расчетные значения и допуски при этом не округляются.
Private Function FormatNumberInvariant(ByVal value As Double) As String
    FormatNumberInvariant = Replace(CStr(value), ",", ".")
End Function

' Проверяет только корректно полученный физический State, не диагностического
' кандидата; сохраняет независимый контроль равновесия по трем компонентам.
Private Sub CheckState(ByRef stats As TConfigTestStats, ByVal prefix As String, ByVal point As CSectionStateResult, _
        Optional ByVal forceTolerance As Double = FORCE_TOLERANCE, Optional ByVal momentTolerance As Double = MOMENT_TOLERANCE)
    Check stats, prefix & ".converged", point.Converged
    Check stats, prefix & ".physical", point.WithinPhysicalRange
    Check stats, prefix & ".noExtension", Not point.ExtensionUsed
    CheckClose stats, prefix & ".residualN", point.ResidualN, 0#, forceTolerance
    CheckClose stats, prefix & ".residualMx", point.ResidualMx, 0#, momentTolerance
    CheckClose stats, prefix & ".residualMy", point.ResidualMy, 0#, momentTolerance
End Sub

' Находит реальную строку по ключу, сохраняя возможность проверки отсутствия.
' Не найденный ключ является ошибкой самого теста, а не молчаливой подстановкой.
Private Function ValueCell(ByVal table As Object, ByVal key As String, ByVal column As Long) As Object
    Dim row As Long
    For row = 2 To table.Rows.Count
        If CStr(table.Cells(row, 1).Value2) = key Then
            Set ValueCell = table.Cells(row, column)
            Exit Function
        End If
    Next row
    Err.Raise vbObjectError + 4499, "ValueCell", "Не найден ключ Config " & key
End Function

' Изменяет только значение ячейки; формулы всей таблицы возвращает владелец набора.
Private Sub SetValue(ByVal table As Object, ByVal key As String, ByVal value As Variant, Optional ByVal column As Long = 2)
    Dim cell As Object: Set cell = ValueCell(table, key, column)
    cell.Value2 = value
End Sub

' Меняет параметр профиля по его ID и ключу во втором столбце таблицы.
' Положение PR1/PR2 не предполагается заранее, чтобы перестановка профилей
' не превратила проверку Config в изменение другого потребителя.
Private Sub SetProfileValue(ByVal table As Object, ByVal key As String, ByVal value As String, ByVal profileId As String)
    Dim row As Long, column As Long, targetRow As Long, targetColumn As Long
    For row = 1 To table.Rows.Count
        If CStr(table.Cells(row, 2).Value2) = key Then targetRow = row
        For column = 3 To table.Columns.Count
            If CStr(table.Cells(row, column).Value2) = profileId Then targetColumn = column
        Next column
    Next row
    If targetRow = 0 Or targetColumn = 0 Then Err.Raise vbObjectError + 4499, "SetProfileValue", _
        "Не найден параметр " & key & " профиля " & profileId
    table.Cells(targetRow, targetColumn).Value2 = value
End Sub

' Возвращает точную пользовательскую запись пути с символом lambda.
Private Function LambdaPath(ByVal component As String) As String
    If component = "Auto" Then
        LambdaPath = component
    Else
        LambdaPath = ChrW$(&H3BB) & "*" & component
    End If
End Function

' Записывает assertion без исключения, чтобы отчет сохранил все отклонения.
Private Sub Check(ByRef stats As TConfigTestStats, ByVal name As String, ByVal condition As Boolean)
    If condition Then
        stats.Passed = stats.Passed + 1: LogLine stats, "OK: " & name
    Else
        stats.Failed = stats.Failed + 1: LogLine stats, "FAIL: " & name
    End If
End Sub

' Сравнивает численное значение с заранее заданным допуском и сохраняет числа.
Private Sub CheckClose(ByRef stats As TConfigTestStats, ByVal name As String, ByVal actual As Double, _
        ByVal expected As Double, ByVal tolerance As Double)
    Check stats, name & "|actual=" & FormatNumberInvariant(actual) & "|expected=" & FormatNumberInvariant(expected), _
        Abs(actual - expected) <= tolerance
End Sub

' Сохраняет промежуточный отчет рядом с тестовой книгой для watchdog-диагностики.
' Тестовый журнал не участвует в расчетах и не заменяет конечный gate.
' Ошибка необязательной файловой копии остается warning в возвращаемом
' отчете; она не обрывает assertions и не скрывает их настоящие отказы.
Private Sub LogLine(ByRef stats As TConfigTestStats, ByVal value As String)
    stats.Report = stats.Report & value & vbCrLf
    Dim fileNumber As Integer: fileNumber = FreeFile
    On Error GoTo ProgressUnavailable
    Open ThisWorkbook.Path & Application.PathSeparator & "Audit03_Search_Progress.txt" For Output As #fileNumber
    Print #fileNumber, stats.Report;
    Close #fileNumber
    Exit Sub
ProgressUnavailable:
    Dim failureNumber As Long, failureDescription As String
    failureNumber = Err.Number: failureDescription = Err.Description
    On Error Resume Next
    Close #fileNumber
    On Error GoTo 0
    stats.Report = stats.Report & "TEST_PROGRESS_WARNING: " & CStr(failureNumber) & "; " & failureDescription & vbCrLf
End Sub

' ==================== ДЛЯ ТЕСТОВ: ОТКАЗ ПРОМЕЖУТОЧНОГО ЖУРНАЛА ====================

' Блокирует только журнал собственной тестовой книги. Проверяет, что отказ
' файловой копии сохраняет assertion и диагностику в возвращаемом отчете;
' после освобождения файла запись возобновляется без повторения расчета.
Private Sub TestAudit03ProgressFileFailure(ByRef stats As TConfigTestStats)
    Dim lockedFile As Integer, probe As TConfigTestStats, warningPreserved As Boolean
    On Error GoTo Failed
    lockedFile = FreeFile
    Open ThisWorkbook.Path & Application.PathSeparator & "Audit03_Search_Progress.txt" For Binary Access Read Write Lock Read Write As #lockedFile
    Check probe, "progressFile.locked.assertion", True
    warningPreserved = probe.Passed = 1 And probe.Failed = 0 And _
        InStr(1, probe.Report, "TEST_PROGRESS_WARNING:", vbBinaryCompare) > 0 And _
        InStr(1, probe.Report, "OK: progressFile.locked.assertion", vbBinaryCompare) > 0
    Close #lockedFile
    lockedFile = 0
    Check stats, "audit03.progressFile.failureDoesNotAbort", warningPreserved
    Check probe, "progressFile.recovered.assertion", True
    Check stats, "audit03.progressFile.recovered", probe.Passed = 2 And probe.Failed = 0 And _
        InStr(1, probe.Report, "OK: progressFile.recovered.assertion", vbBinaryCompare) > 0
    Exit Sub
Failed:
    Dim failureNumber As Long, failureDescription As String
    failureNumber = Err.Number: failureDescription = Err.Description
    On Error Resume Next
    If lockedFile <> 0 Then Close #lockedFile
    On Error GoTo 0
    Check stats, "audit03.progressFile.runtime." & CStr(failureNumber) & "." & failureDescription, False
End Sub

' ==================== ДЛЯ ТЕСТОВ: ПРЕДСТАВИМОСТЬ СТУПЕНЕЙ RETRY ====================

' Проверяет совместный диапазон числа ступеней и повторов до любого solve.
' Удвоение ступеней каждой попытки должно оставаться представимым в Long;
' сообщение о неверной паре использует адрес фактически прочитанной таблицы.
Private Sub TestAudit03CapacityRetryRange(ByRef stats As TConfigTestStats)
    Dim sheet As Object, table As Object, cell As Object, settings As CSystemSettingsReader
    Dim calculator As CCapacityCalculator, position As Long, scenario As Long
    Dim baseSteps As Variant, retries As Variant, accepted As Variant
    Dim errorNumber As Long, reason As String, prefix As String, key As String
    Dim solveCount As Long, oldAlerts As Boolean
    On Error GoTo Failed
    solveCount = SectionEquilibriumSolveCount()
    oldAlerts = Application.DisplayAlerts
    Set sheet = ThisWorkbook.Worksheets.Add
    baseSteps = Array(1, 1, 1, 3, 3, 1073741823, 1073741824, 2147483647, 1)
    retries = Array(0, 30, 31, 29, 30, 1, 1, 0, 2147483647)
    accepted = Array(True, True, False, True, False, True, False, False, False)
    For position = 0 To 1
        If position = 0 Then Set table = sheet.Range("A1:C3") Else Set table = sheet.Range("CH800:CJ802")
        table.Cells(1, 1).Value2 = "Параметр": table.Cells(1, 2).Value2 = "Значение"
        table.Cells(1, 3).Value2 = "Комментарий"
        table.Cells(2, 1).Value2 = "Capacity.BaseLoadSteps"
        table.Cells(3, 1).Value2 = "Capacity.MaxRetries"
        For scenario = LBound(baseSteps) To UBound(baseSteps)
            table.Cells(2, 2).Value2 = baseSteps(scenario)
            table.Cells(3, 2).Value2 = retries(scenario)
            Set settings = New CSystemSettingsReader: settings.LoadFromRange table
            Set calculator = New CCapacityCalculator
            calculator.ConfigureSearch 1#, "LoadMultiplier", "Bisection", 1024#, 0.01, 0.00001, _
                settings.GetLong("Capacity.MaxRetries"), settings.GetLong("Capacity.BaseLoadSteps"), 60
            errorNumber = 0: reason = vbNullString
            Audit03ValidateCapacityRetryRange calculator, settings, errorNumber, reason
            prefix = "audit03.capacityRetryRange.p" & CStr(position) & ".v" & CStr(scenario)
            If CBool(accepted(scenario)) Then
                Check stats, prefix & ".accepted", errorNumber = 0
            Else
                key = "Capacity.MaxRetries"
                If CLng(baseSteps(scenario)) = 2147483647 Then key = "Capacity.BaseLoadSteps"
                Set cell = ValueCell(table, key, 2)
                Check stats, prefix & ".inputError", errorNumber = vbObjectError + 4190
                Check stats, prefix & ".key", InStr(1, reason, key, vbBinaryCompare) > 0
                Check stats, prefix & ".address", InStr(1, reason, cell.Address(False, False), vbTextCompare) > 0
                Check stats, prefix & ".action", InStr(1, reason, "Уменьшите", vbTextCompare) > 0
            End If
            LogLine stats, "CAPACITY_RETRY_RANGE: " & prefix & "; steps=" & CStr(baseSteps(scenario)) & _
                "; retries=" & CStr(retries(scenario)) & "; error=" & CStr(errorNumber) & "; " & reason
        Next scenario
    Next position
    Check stats, "audit03.capacityRetryRange.noSolve", SectionEquilibriumSolveCount() = solveCount
    GoTo Restore
Failed:
    Check stats, "audit03.capacityRetryRange.runtime." & CStr(Err.Number) & "." & Err.Description, False
Restore:
    On Error Resume Next
    Application.DisplayAlerts = False
    If Not sheet Is Nothing Then sheet.Delete
    Application.DisplayAlerts = oldAlerts
    On Error GoTo 0
End Sub

' Перехватывает только ожидаемую ошибку предварительной валидации, чтобы
' тест дошел до остальных пар значений. НДС и Search здесь не запускаются.
Private Sub Audit03ValidateCapacityRetryRange(ByVal calculator As CCapacityCalculator, _
        ByVal settings As CSystemSettingsReader, ByRef errorNumber As Long, ByRef reason As String)
    On Error GoTo Failed
    calculator.ValidateSettings settings
    Exit Sub
Failed:
    errorNumber = Err.Number: reason = Err.Description
End Sub

' Возвращает отдельный короткий протокол крайних пар Config без запуска
' полного набора численных расчетов и без изменения исходных таблиц книги.
Public Function RunAudit03CapacityRetryRangeTests() As String
    Dim stats As TConfigTestStats
    TestAudit03CapacityRetryRange stats
    LogLine stats, "TOTAL_CAPACITY_RETRY_RANGE: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunAudit03CapacityRetryRangeTests = stats.Report
End Function

' ==================== ДЛЯ ТЕСТОВ: ДИАПАЗОН РАСЧЕТНОЙ ДЛИНЫ ====================

' Конечные L и mu могут дать непредставимое произведение или его квадрат.
' Проверяет Config -> batch -> writer для обеих методик, сжатия/растяжения и переноса
' именованной таблицы; ошибка должна показывать оба адреса до запуска расчета.
Private Sub CheckStabilityLengthRange(ByRef stats As TConfigTestStats, ByVal table As Object, _
        ByRef configured As Variant, ByVal position As Long, ByVal section As CSectionModel, _
        ByVal provider As CMaterialModelProvider, ByVal units As CUnitSystem)
    Dim code As Variant, force As Variant, scenario As Long, muKey As String, prefix As String
    Dim batch As CBatchSectionCalculator, meta As CResultMeta, message As String
    For Each code In Array("SP63", "SP35")
        For Each force In Array(-10000#, 10000#)
            For scenario = 1 To 9
                table.Formula = configured
                SetValue table, "Stability.Code", CStr(code)
                muKey = "Stability.Mu1"
                If scenario = 3 Or scenario = 6 Or scenario = 9 Then muKey = "Stability.Mu2"
                If scenario = 1 Then
                    SetValue table, "Stability.ElementLength", 1E+308
                    SetValue table, muKey, 2#
                ElseIf scenario <= 3 Then
                    SetValue table, muKey, 1E+308
                ElseIf scenario = 4 Then
                    SetValue table, "Stability.ElementLength", 1E+200
                    SetValue table, muKey, 1#
                ElseIf scenario <= 6 Then
                    SetValue table, muKey, 1E+200
                ElseIf scenario = 7 Then
                    SetValue table, "Stability.ElementLength", 1E-200
                    SetValue table, muKey, 1#
                Else
                    SetValue table, muKey, 1E-200
                End If
                prefix = "audit03.stabilityLengthRange.p" & CStr(position) & "." & CStr(code) & _
                    ".N" & CStr(force) & ".s" & CStr(scenario)
                Set batch = StabilityConfigRun(stats, prefix, section, provider, units, CDbl(force), 0#, 0#)
                Set meta = batch.ResultAt(1).StabilityMeta
                message = batch.ResultAt(1).OverallMeta.ResultComment
                Check stats, prefix & ".inputErr", batch.ResultAt(1).Status = "InputErr"
                Check stats, prefix & ".notCalculated", Not meta.Calculated
                Check stats, prefix & ".noStateSolve", batch.SolverCallCount = 0
                Check stats, prefix & ".lengthKey", InStr(1, message, "Stability.ElementLength", vbBinaryCompare) > 0
                Check stats, prefix & ".muKey", InStr(1, message, muKey, vbBinaryCompare) > 0
                Check stats, prefix & ".lengthCell", InStr(1, message, _
                    ValueCell(table, "Stability.ElementLength", 2).Address(False, False), vbTextCompare) > 0
                Check stats, prefix & ".muCell", InStr(1, message, ValueCell(table, muKey, 2).Address(False, False), vbTextCompare) > 0
                If scenario <= 6 Then
                    Check stats, prefix & ".action", InStr(1, message, "Уменьшите", vbTextCompare) > 0
                Else
                    Check stats, prefix & ".action", InStr(1, message, "Увеличьте", vbTextCompare) > 0
                End If
                Check stats, prefix & ".localized", InStr(1, message, "Overflow", vbTextCompare) = 0
                LogLine stats, "STABILITY_LENGTH_RANGE: " & prefix & "|" & message
            Next scenario
            table.Formula = configured
            SetValue table, "Stability.Code", CStr(code)
            Set batch = StabilityConfigRun(stats, "audit03.stabilityLengthRange.recovery.p" & CStr(position) & _
                "." & CStr(code) & ".N" & CStr(force), section, provider, units)
            Check stats, "audit03.stabilityLengthRange.recovery.p" & CStr(position) & "." & CStr(code) & ".N" & CStr(force), _
                batch.ResultAt(1).StabilityMeta.Calculated And batch.ResultAt(1).Status <> "InputErr" And _
                batch.ResultAt(1).Status <> "CalcErr" And batch.ResultAt(1).Status <> "NumFail"
        Next force
    Next code
    table.Formula = configured
End Sub
