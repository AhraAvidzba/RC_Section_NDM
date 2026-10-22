Attribute VB_Name = "modTestFinalGeometryChecks"
Option Explicit
' ==========================================================================
' ДЛЯ ТЕСТОВ: ограниченный итоговый набор проверок в отдельной копии книги.
' В рабочую книгу не включается. Использует обычные построители, расчёты и запись результатов.
' ==========================================================================
Private mReport As String
Private mPassed As Long
Private mFailed As Long

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет восемь нагрузок и два сочетания с устойчивостью для девяти форм; отдельно контролирует однократную подготовку механических свойств и освобождение временных данных.
Public Function TestRunFamilyComplement(Optional ByVal code As String = "SP35") As String
    Dim config As Object, saved As Variant, fixture As Variant, k As Long
    Dim settings As CSystemSettingsReader, units As CUnitSystem, model As CSectionModel
    Dim provider As CMaterialModelProvider, props As CSectionPropertiesCalculator
    Dim solver As CSectionSolver, batch As CBatchSectionCalculator, profiles As CCalculationProfileCatalog
    Dim n As Double, mx As Double, my As Double, dx As Double, dy As Double
    Dim lo As Double, hi As Double, label As String
    TestResetReport
    On Error GoTo Failed
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = config.Formula
    TestSetSetting "Stability.Code", code
    TestSetSetting "Stability.ElementLength", "1000"
    TestSetSetting "Stability.Mu1", "1": TestSetSetting "Stability.Mu2", "1"
    TestSetSetting "Stability.AccidentalEccentricityMode", "User"
    TestSetSetting "Stability.AccidentalEccentricityPlanes", "BothPlanes"
    TestSetSetting "Stability.AccidentalEccentricityUser1", "5"
    TestSetSetting "Stability.AccidentalEccentricityUser2", "5"
    TestSetSetting "Calculation.ZeroMomentPerDepth", "0"
    TestSetProfile "Calculation.Strength.DirectState", "Yes"
    TestSetProfile "Calculation.Strength.Capacity", "No"
    TestSetProfile "Calculation.Crack.Width", "No"
    TestSetProfile "Calculation.Stability.Enabled", "Yes"
    For Each fixture In TestFixtures()
        modTestShapeArchitecture.TestConfigureShapeFixture CStr(fixture)
        Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
        Set units = New CUnitSystem: units.LoadFromSettings settings
        Set model = BuildWorkbookSectionModel(ThisWorkbook, settings, units)
        Set provider = New CMaterialModelProvider: provider.Initialize settings, units
        Set props = New CSectionPropertiesCalculator: props.CalculateConcrete model
        props.CalculateProjection model, 1#, 0#, False, lo, hi: dx = hi - lo
        props.CalculateProjection model, 0#, 1#, False, lo, hi: dy = hi - lo
        ' Задаём моменты относительно начала координат с явной добавкой N*xc/yc.
        ' Эксцентриситеты и оба знака нагрузки проверяются независимо от симметрии формы.
        For k = 0 To 7
            n = -2# * props.Area
            If k = 1 Then n = 0.1 * props.Area
            If k = 7 Then n = -0.0001 * props.Area
            mx = n * props.CentroidY: my = n * props.CentroidX
            If k = 2 Then mx = mx + props.Area * dy * 0.1
            If k = 3 Then mx = mx - props.Area * dy * 0.1
            If k = 4 Then my = my + props.Area * dx * 0.1
            If k = 5 Then my = my - props.Area * dx * 0.1
            If k = 6 Then mx = mx - props.Area * dy * 0.1: my = my + props.Area * dx * 0.1
            Set solver = New CSectionSolver: solver.ApplySettings settings, units
            solver.Solve model, provider.ConcreteMaterial(cpStrength), provider.SteelMaterial(cpStrength), n, mx, my
            label = CStr(fixture) & ".load" & CStr(k)
            TestCheck label & ".converged", solver.Converged
            mReport = mReport & "FAMILY_LOAD: " & label & "; reference=" & CStr(props.CentroidX) & "," & CStr(props.CentroidY) & _
                "; N=" & CStr(n) & "; Mx=" & CStr(mx) & "; My=" & CStr(my) & "; residual=" & _
                CStr(solver.ResidualN) & "," & CStr(solver.ResidualMx) & "," & CStr(solver.ResidualMy) & vbCrLf
        Next k
        Set profiles = New CCalculationProfileCatalog: profiles.LoadFromWorkbook ThisWorkbook
        Set batch = New CBatchSectionCalculator: batch.Initialize model, provider
        Set batch.ProfileCatalog = profiles: batch.ApplySettings settings, units
        batch.SetSP35Table721 ThisWorkbook.Names("rngSP35Table721").RefersToRange.Value2
        batch.AddCombination "COMP", -120000#, 0#, 0#, "PR1", "Central compression", "Auto"
        batch.AddCombination "BEND", -120000#, 1500000#, -800000#, "PR1", "Biaxial", "Auto"
        batch.ApplyLoadReference props.CentroidX, props.CentroidY
        batch.Execute
        For k = 1 To batch.Count
            label = CStr(fixture) & "." & code & ".batch" & CStr(k)
            TestCheck label & ".stability", batch.ResultAt(k).StabilityResult.Status = "OK"
            TestCheck label & ".direct", batch.ResultAt(k).DirectStateMeta.Calculated
            mReport = mReport & "FAMILY_STABILITY: " & label & "; status=" & batch.ResultAt(k).StabilityResult.Status & _
                "; branch=" & batch.ResultAt(k).StabilityResult.Branch & "; comment=" & batch.ResultAt(k).StabilityMeta.ResultComment & vbCrLf
        Next k
        ' Бетонные свойства и приведённые свойства единственного профиля —
        ' два разных кэша. Каждый готовится один раз для двух сочетаний.
        TestCheck CStr(fixture) & ".buildOnce", batch.TestMechanicalPreparationBuildCount = 2
        mReport = mReport & "FAMILY_PREPARATION: " & CStr(fixture) & "; concrete+transformed=" & CStr(batch.TestMechanicalPreparationBuildCount) & "; combinations=" & CStr(batch.Count) & vbCrLf
        TestCheck CStr(fixture) & ".released", batch.TestRunPreparationReleased
        TestCheck CStr(fixture) & ".worst", Len(batch.WorstCombinationID) > 0
    Next fixture
    GoTo Restore
Failed:
    TestCheck "family.runtime." & CStr(Err.Number) & "." & Err.Description, False
Restore:
    If IsArray(saved) Then config.Formula = saved
    TestRunFamilyComplement = TestFinishReport("FAMILY_COMPLEMENT_" & code)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' одна нагрузка для каждой группы сечений. Замер штатного запуска
' включает подготовку, Results, отчёт и автоматическую схему, а не только решение НДС.
Public Function TestMeasureFamilyWorkflow(Optional ByVal outputDirectory As String = "") As String
    Dim fixture As Variant, result As String, label As String, before As Long
    Dim config As Object, saved As Variant, settings As CSystemSettingsReader
    Dim reader As CSectionPlotDataReader, plotter As CSectionPlotter, chart As Object
    Dim diagramMode As String
    On Error GoTo Failed
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = config.Formula
    diagramMode = TestReadDiagramMode()
    For Each fixture In TestFixtures()
        modTestShapeArchitecture.TestConfigureShapeFixture CStr(fixture)
        modTestPerformance.TestConfigurePerformanceFixture "DIRECT", 1, "Newton", "Saved", "Normal", diagramMode
        TestSetSetting "Plot.Enabled", "Yes": TestSetSetting "Plot.AutoUpdateAfterCalculation", "Yes"
        TestSetSetting "Plot.LoadCase", "Worst": TestSetSetting "General.ExecutionReportEnabled", "Yes"
        TestSetProfile "Visualization.State", "StrengthState": TestSetProfile "Visualization.Quantity", "Stress"
        label = CStr(fixture)
        result = result & "FAMILY_WORKFLOW: " & label & "; diagramExtension=" & TestReadDiagramMode() & "; " & Replace(modTestPerformance.TestMeasurePerformanceBatch("Full"), Chr$(30), "; ") & vbCrLf
        If Len(outputDirectory) > 0 Then
            Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
            Set reader = New CSectionPlotDataReader: reader.LoadFromWorkbook ThisWorkbook, settings
            Set plotter = New CSectionPlotter
            before = SectionEquilibriumSolveCount()
            plotter.Draw ThisWorkbook, reader, settings
            Set chart = ThisWorkbook.Worksheets("Расчет").ChartObjects("chtNDMSectionPlot")
            If Not chart.Chart.Export(outputDirectory & Application.PathSeparator & label & "_Stress.png", "PNG") Then Err.Raise 5, , "PNG export failed"
            TestSetProfile "Visualization.Quantity", "Strain"
            settings.LoadFromWorkbook ThisWorkbook: reader.LoadFromWorkbook ThisWorkbook, settings
            plotter.Draw ThisWorkbook, reader, settings
            If Not chart.Chart.Export(outputDirectory & Application.PathSeparator & label & "_Strain.png", "PNG") Then Err.Raise 5, , "PNG export failed"
            If SectionEquilibriumSolveCount() <> before Then Err.Raise 5, , "Passive plot added solve"
            result = result & "PASSIVE_FAMILY: " & label & "; stress/strain PNG; addedSolves=0; elements=" & CStr(reader.Count) & vbCrLf
        End If
    Next fixture
    GoTo Restore
Failed:
    result = result & "FAIL: workflow; " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Restore:
    If IsArray(saved) Then config.Formula = saved
    TestMeasureFamilyWorkflow = result
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' фазовый замер одного HollowOffset для объяснения общей
' разницы времени. Стратегии решения и число сочетаний прежние.
Public Function TestMeasureFamilyStages() As String
    Dim diagramMode As String
    diagramMode = TestReadDiagramMode()
    modTestShapeArchitecture.TestConfigureShapeFixture "HollowOffset"
    modTestPerformance.TestConfigurePerformanceFixture "DIRECT", 1, "Newton", "Saved", "Normal", diagramMode
    TestSetSetting "Plot.Enabled", "Yes": TestSetSetting "Plot.AutoUpdateAfterCalculation", "Yes"
    TestSetSetting "Plot.LoadCase", "Worst": TestSetSetting "General.ExecutionReportEnabled", "Yes"
    TestSetProfile "Visualization.State", "StrengthState": TestSetProfile "Visualization.Quantity", "Stress"
    Dim i As Long, result As String
    For i = 1 To 3
        result = result & "FAMILY_STAGES: HollowOffset.sample" & CStr(i) & "; diagramExtension=" & TestReadDiagramMode() & "; " & Replace(modTestPerformance.TestMeasurePerformanceBatch("Staged"), Chr$(30), "; ") & vbCrLf
    Next i
    TestMeasureFamilyStages = result
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' одно полное сочетание с несущей способностью, трещинами, устойчивостью,
' отчётом и схемой для направленного замера времени.
Public Function TestMeasureMixedWorkflow() As String
    Dim diagramMode As String
    diagramMode = TestReadDiagramMode()
    modTestShapeArchitecture.TestConfigureShapeFixture "HollowOffset"
    modTestPerformance.TestConfigurePerformanceFixture "MIXED", 1, "Newton", "Saved", "Normal", diagramMode
    TestSetSetting "Stability.ElementLength", "1000"
    TestSetSetting "Plot.Enabled", "Yes": TestSetSetting "Plot.AutoUpdateAfterCalculation", "Yes"
    TestSetSetting "Plot.LoadCase", "Worst": TestSetSetting "General.ExecutionReportEnabled", "Yes"
    TestSetProfile "Visualization.State", "StrengthState": TestSetProfile "Visualization.Quantity", "Stress"
    Dim settings As CSystemSettingsReader, units As CUnitSystem, loads As Object
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    Set loads = ThisWorkbook.Names("rngLoadCombinations").RefersToRange
    loads.Cells(2, 2).Value2 = -50000# / units.InputForceToInternal(1#)
    loads.Cells(2, 3).Value2 = 150000000# / units.InputMomentMxToInternal(1#)
    loads.Cells(2, 4).Value2 = -30000000# / units.InputMomentMyToInternal(1#)
    TestMeasureMixedWorkflow = "MIXED_WORKFLOW: diagramExtension=" & TestReadDiagramMode() & "; " & Replace(modTestPerformance.TestMeasurePerformanceBatch("Full"), Chr$(30), "; ")
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' проверка проекции текущего состояния требует применимого НДС для СП 35;
' пользовательская книга может содержать AutoCAD/иной профиль без этой ветви.
Public Function TestRunSP35ProjectionFixture() As String
    Dim diagramMode As String
    diagramMode = TestReadDiagramMode()
    modTestShapeArchitecture.TestConfigureShapeFixture "HollowOffset"
    modTestPerformance.TestConfigurePerformanceFixture "CRACK_SP35", 1, "Newton", "Saved", "Normal", diagramMode
    Dim settings As CSystemSettingsReader, units As CUnitSystem, loads As Object
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    Set loads = ThisWorkbook.Names("rngLoadCombinations").RefersToRange
    loads.Cells(2, 2).Value2 = -50000# / units.InputForceToInternal(1#)
    loads.Cells(2, 3).Value2 = 800000000# / units.InputMomentMxToInternal(1#)
    loads.Cells(2, 4).Value2 = -100000000# / units.InputMomentMyToInternal(1#)
    TestRunSP35ProjectionFixture = "PROJECTION_DIAGRAM_MODE: " & TestReadDiagramMode() & vbCrLf & modTestSP35CrackWidth.RunSP35CurrentProjectionTests()
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сохраняет фактический режим продолжения диаграмм до подготовки контрольного примера.
Private Function TestReadDiagramMode() As String
    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook
    If settings.GetRequiredBoolean("General.DiagramExtension") Then
        TestReadDiagramMode = "Yes"
    Else
        TestReadDiagramMode = "No"
    End If
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Перечисляет девять контрольных вариантов геометрии в постоянном порядке для одинакового охвата обоих режимов диаграмм.
Private Function TestFixtures() As Variant
    TestFixtures = Array("Circle", "Rectangle", "LSection", "TwoRectangles", "RoundedSimple", "RoundedTapered", "RoundedMixed", "HollowCentered", "HollowOffset")
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' проверяет отображение без полного ручного прогона интерфейса.
' Два знака изгиба дают смешанные поля; дальнейшие операции только читают Results
' при изменении Config, размера схемы и отсутствии выбранного НДС.
Public Function TestRunDirectedPlotModes(ByVal outputDirectory As String) As String
    Dim config As Object, saved As Variant, settings As CSystemSettingsReader, units As CUnitSystem
    Dim loads As Object, reader As CSectionPlotDataReader, plotter As CSectionPlotter
    Dim container As Object, marker As Object, state As Variant, quantity As Variant
    Dim before As Long, count As Long, minValue As Double, maxValue As Double, i As Long
    Dim initialCount As Long, originalWidth As Double, originalHeight As Double, label As String
    Dim diagramMode As String
    TestResetReport
    On Error GoTo Failed
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = config.Formula
    diagramMode = TestReadDiagramMode()
    modTestShapeArchitecture.TestConfigureShapeFixture "HollowOffset"
    modTestPerformance.TestConfigurePerformanceFixture "DIRECT", 2, "Newton", "Saved", "Normal", diagramMode
    TestCheck "plotModes.diagramMode", TestReadDiagramMode() = diagramMode
    mReport = mReport & "PLOT_DIAGRAM_MODE: " & TestReadDiagramMode() & vbCrLf
    TestSetSetting "Plot.Enabled", "Yes": TestSetSetting "Plot.AutoUpdateAfterCalculation", "Yes"
    TestSetSetting "Plot.LoadCase", "PERF_01": TestSetSetting "General.ExecutionReportEnabled", "No"
    TestSetProfile "Visualization.State", "StrengthState"
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    Set loads = ThisWorkbook.Names("rngLoadCombinations").RefersToRange
    For i = 2 To 3
        loads.Cells(i, 2).Value2 = -50000# / units.InputForceToInternal(1#)
        loads.Cells(i, 3).Value2 = ((-1#) ^ i) * 70000000# / units.InputMomentMxToInternal(1#) ' Чередование степеней меняет знак изгиба без изменения его модуля.
        loads.Cells(i, 4).Value2 = ((-1#) ^ (i + 1)) * 30000000# / units.InputMomentMyToInternal(1#)
    Next i
    mReport = mReport & modTestPerformance.TestMeasurePerformanceBatch("Full") & vbCrLf
    Set reader = New CSectionPlotDataReader: Set plotter = New CSectionPlotter
    before = SectionEquilibriumSolveCount()
    Set container = ThisWorkbook.Worksheets("Расчет").ChartObjects("chtNDMSectionPlot")
    originalWidth = container.Width: originalHeight = container.Height
    Set marker = container.Chart.Shapes.AddShape(1, 10#, 10#, 10#, 10#): marker.Name = "User_QA_shape"
    For Each state In Array("PERF_01", "PERF_02", "Worst")
        TestSetSetting "Plot.LoadCase", CStr(state)
        For Each quantity In Array("Stress", "Strain")
            TestSetProfile "Visualization.Quantity", CStr(quantity)
            settings.LoadFromWorkbook ThisWorkbook: reader.LoadFromWorkbook ThisWorkbook, settings
            TestCheck CStr(state) & "." & CStr(quantity) & ".state", Not reader.GeometryOnly
            minValue = 1E+100: maxValue = -1E+100
            For i = 1 To reader.Count
                If reader.ResultValue(i) < minValue Then minValue = reader.ResultValue(i)
                If reader.ResultValue(i) > maxValue Then maxValue = reader.ResultValue(i)
            Next i
            TestCheck CStr(state) & "." & CStr(quantity) & ".signed", minValue < 0# And maxValue > 0#
            plotter.Draw ThisWorkbook, reader, settings
            initialCount = container.Chart.Shapes.Count
            plotter.Draw ThisWorkbook, reader, settings
            TestCheck CStr(state) & "." & CStr(quantity) & ".repeat", container.Chart.Shapes.Count = initialCount
            TestCheck CStr(state) & "." & CStr(quantity) & ".foreign", container.Chart.Shapes("User_QA_shape").Name = "User_QA_shape"
            label = CStr(state) & "_" & CStr(quantity)
            If Not container.Chart.Export(outputDirectory & Application.PathSeparator & label & ".png", "PNG") Then Err.Raise 5, , "PNG export failed"
            mReport = mReport & "PLOT_MODE: " & label & "; LC=" & reader.LoadCase & "; elements=" & CStr(reader.Count) & _
                "; shapes=" & CStr(initialCount) & "; signed=" & CStr(minValue) & "," & CStr(maxValue) & vbCrLf
        Next quantity
    Next state
    count = reader.Count
    TestSetSetting "Geometry.Type", "Circle"
    settings.LoadFromWorkbook ThisWorkbook: reader.LoadFromWorkbook ThisWorkbook, settings
    TestCheck "saved.afterConfigChange", reader.Count = count
    TestSetAnnotationFlags "No"
    settings.LoadFromWorkbook ThisWorkbook
    container.Width = 650#: container.Height = 500#
    plotter.Draw ThisWorkbook, reader, settings
    If Not container.Chart.Export(outputDirectory & Application.PathSeparator & "Resize_NoLabels.png", "PNG") Then Err.Raise 5, , "PNG export failed"
    TestCheck "passive.noSolve", SectionEquilibriumSolveCount() = before
    ' Повторяем штатный запуск без прочности, несущей способности и трещин.
    ' Состояния StrengthState нет; схема должна показать только сохранённую геометрию.
    TestSetSetting "Geometry.Type", "HollowRectangle"
    TestSetProfile "Calculation.Strength.DirectState", "No": TestSetProfile "Calculation.Stability.Enabled", "Yes"
    mReport = mReport & modTestPerformance.TestMeasurePerformanceBatch("Full") & vbCrLf
    settings.LoadFromWorkbook ThisWorkbook: reader.LoadGeometryOnlyForMissingState ThisWorkbook, settings
    TestCheck "missingState.geometryOnly", reader.GeometryOnly
    before = SectionEquilibriumSolveCount()
    plotter.Draw ThisWorkbook, reader, settings
    TestCheck "missingState.noSolve", SectionEquilibriumSolveCount() = before
    If Not container.Chart.Export(outputDirectory & Application.PathSeparator & "MissingState.png", "PNG") Then Err.Raise 5, , "PNG export failed"
    GoTo Restore
Failed:
    TestCheck "plotModes.runtime." & CStr(Err.Number) & "." & Err.Description, False
Restore:
    If IsArray(saved) Then config.Formula = saved
    If originalWidth > 0# Then container.Width = originalWidth: container.Height = originalHeight
    TestRunDirectedPlotModes = TestFinishReport("DIRECTED_PLOT_MODES")
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Включает или отключает обе группы аннотаций по строке Enabled именованной таблицы; при отсутствии строки останавливает проверку.
Private Sub TestSetAnnotationFlags(ByVal value As String)
    Dim table As Object, k As Long
    Set table = ThisWorkbook.Names("rngPlotAnnotationSettings").RefersToRange
    For k = 1 To table.Rows.Count
        If CStr(table.Cells(k, 1).Value2) = "Enabled" Then
            table.Cells(k, 2).Value2 = value: table.Cells(k, 3).Value2 = value: Exit Sub
        End If
    Next k
    Err.Raise 5, , "Missing annotation Enabled row"
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Меняет системную настройку по ключу в текущем именованном диапазоне; отсутствие ключа не заменяется подстановкой.
Private Sub TestSetSetting(ByVal key As String, ByVal value As String)
    Dim table As Object, k As Long
    Set table = ThisWorkbook.Names("rngSystemSettings").RefersToRange
    For k = 1 To table.Rows.Count
        If CStr(table.Cells(k, 1).Value2) = key Then table.Cells(k, 2).Value2 = value: Exit Sub
    Next k
    Err.Raise 5, "modTestFinalGeometryChecks", "Missing setting: " & key
End Sub
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Меняет значение первого профиля по ключу строки в таблице профилей, подготовленной для отдельной проверочной книги.
Private Sub TestSetProfile(ByVal key As String, ByVal value As String)
    Dim table As Object, k As Long
    Set table = ThisWorkbook.Names("rngCalculationProfiles").RefersToRange
    For k = 1 To table.Rows.Count
        If CStr(table.Cells(k, 2).Value2) = key Then table.Cells(k, 3).Value2 = value: Exit Sub
    Next k
    Err.Raise 5, "modTestFinalGeometryChecks", "Missing profile key: " & key
End Sub
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Обнуляет протокол и оба счётчика перед самостоятельным набором проверок.
Private Sub TestResetReport()
    mReport = vbNullString: mPassed = 0: mFailed = 0
End Sub
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Учитывает отдельное условие и сохраняет его имя со статусом OK или FAIL, не прерывая остальные проверки.
Private Sub TestCheck(ByVal label As String, ByVal success As Boolean)
    If success Then mPassed = mPassed + 1: mReport = mReport & "OK: " & label & vbCrLf Else mFailed = mFailed + 1: mReport = mReport & "FAIL: " & label & vbCrLf
End Sub
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Дополняет накопленный протокол именем набора и итоговыми числами успешных и неуспешных проверок.
Private Function TestFinishReport(ByVal label As String) As String
    TestFinishReport = mReport & "TOTAL_" & label & ": passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function
