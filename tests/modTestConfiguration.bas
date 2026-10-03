Attribute VB_Name = "modTestConfiguration"
Option Explicit

' ==========================================================================
' ДЛЯ ТЕСТОВ: поведение численных селекторов через настоящий Config
' ==========================================================================
' Модуль меняет ячейки изолированной книги, читает их штатным reader-ом и
' проверяет реальные результаты Capacity/Formation, а не только getters.
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
    CheckClose stats, "universalPath.constantAlreadyCracked.outputPsi1", CDbl(crackAnchor.Offset(0, 35).Value2), 1#, 0#
    CheckClose stats, "universalPath.constantAlreadyCracked.outputWidth", CDbl(crackAnchor.Offset(0, 41).Value2), _
        units.InternalLengthToOutput(result.CrackResult.Width.CrackWidth), 0.000000000001
    CheckClose stats, "universalPath.constantAlreadyCracked.outputEs", CDbl(crackAnchor.Offset(0, 40).Value2), _
        units.InternalStressToOutput(result.CrackResult.Width.SteelEs), 0.000000000001
    Check stats, "universalPath.constantAlreadyCracked.outputWidthStatus", CStr(crackAnchor.Offset(0, 44).Value2) = _
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
Private Sub LogLine(ByRef stats As TConfigTestStats, ByVal value As String)
    stats.Report = stats.Report & value & vbCrLf
    Dim fileNumber As Integer: fileNumber = FreeFile
    Open ThisWorkbook.Path & Application.PathSeparator & "Audit03_Search_Progress.txt" For Output As #fileNumber
    Print #fileNumber, stats.Report;
    Close #fileNumber
End Sub
