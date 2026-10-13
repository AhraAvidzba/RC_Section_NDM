Attribute VB_Name = "modFinalGeometryChecks"
Option Explicit
' ДЛЯ ТЕСТОВ: ограниченная итоговая матрица на изолированной копии книги.
' Не подключается в production. Использует штатные builders, solver и writers.
Private mReport As String
Private mPassed As Long
Private mFailed As Long

Public Function RunFamilyComplement(Optional ByVal code As String = "SP35") As String
    Dim config As Object, saved As Variant, fixture As Variant, k As Long
    Dim settings As CSystemSettingsReader, units As CUnitSystem, model As CSectionModel
    Dim provider As CMaterialModelProvider, props As CSectionPropertiesCalculator
    Dim solver As CSectionSolver, batch As CBatchSectionCalculator, profiles As CCalculationProfileCatalog
    Dim n As Double, mx As Double, my As Double, dx As Double, dy As Double
    Dim lo As Double, hi As Double, label As String
    ResetReport
    On Error GoTo Failed
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = config.Formula
    SetSetting "Stability.Code", code
    SetSetting "Stability.ElementLength", "1000"
    SetSetting "Stability.Mu1", "1": SetSetting "Stability.Mu2", "1"
    SetSetting "Stability.AccidentalEccentricityMode", "User"
    SetSetting "Stability.AccidentalEccentricityPlanes", "BothPlanes"
    SetSetting "Stability.AccidentalEccentricityUser1", "5"
    SetSetting "Stability.AccidentalEccentricityUser2", "5"
    SetSetting "Calculation.ZeroMomentPerDepth", "0"
    SetProfile "Calculation.Strength.DirectState", "Yes"
    SetProfile "Calculation.Strength.Capacity", "No"
    SetProfile "Calculation.Crack.Width", "No"
    SetProfile "Calculation.Stability.Enabled", "Yes"
    For Each fixture In Fixtures()
        modTestShapeArchitecture.ConfigureShapeFixture CStr(fixture)
        Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
        Set units = New CUnitSystem: units.LoadFromSettings settings
        Set model = BuildWorkbookSectionModel(ThisWorkbook, settings, units)
        Set provider = New CMaterialModelProvider: provider.Initialize settings, units
        Set props = New CSectionPropertiesCalculator: props.CalculateConcrete model
        props.CalculateProjection model, 1#, 0#, False, lo, hi: dx = hi - lo
        props.CalculateProjection model, 0#, 1#, False, lo, hi: dy = hi - lo
        ' Моменты относительно origin: центральная часть N*xc/yc добавляется
        ' явно. Эксцентриситеты и оба знака не зависят от симметрии формы.
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
            Check label & ".converged", solver.Converged
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
            Check label & ".stability", batch.ResultAt(k).StabilityResult.Status = "OK"
            Check label & ".direct", batch.ResultAt(k).DirectStateMeta.Calculated
            mReport = mReport & "FAMILY_STABILITY: " & label & "; status=" & batch.ResultAt(k).StabilityResult.Status & _
                "; branch=" & batch.ResultAt(k).StabilityResult.Branch & "; comment=" & batch.ResultAt(k).StabilityMeta.ResultComment & vbCrLf
        Next k
        ' Бетонные свойства и приведённые свойства единственного профиля —
        ' два разных кэша. Каждый готовится один раз для двух сочетаний.
        Check CStr(fixture) & ".buildOnce", batch.MechanicalPreparationBuildCount = 2
        mReport = mReport & "FAMILY_PREPARATION: " & CStr(fixture) & "; concrete+transformed=" & CStr(batch.MechanicalPreparationBuildCount) & "; combinations=" & CStr(batch.Count) & vbCrLf
        Check CStr(fixture) & ".released", batch.RunPreparationReleased
        Check CStr(fixture) & ".worst", Len(batch.WorstCombinationID) > 0
    Next fixture
    GoTo Restore
Failed:
    Check "family.runtime." & CStr(Err.Number) & "." & Err.Description, False
Restore:
    If IsArray(saved) Then config.Formula = saved
    RunFamilyComplement = FinishReport("FAMILY_COMPLEMENT_" & code)
End Function

' ДЛЯ ТЕСТОВ: один LC на семью; штатный запуск включает подготовку,
' Results, execution report и автоматическую схему, а не только solver.
Public Function MeasureFamilyWorkflow(Optional ByVal outputDirectory As String = "") As String
    Dim fixture As Variant, result As String, label As String, before As Long
    Dim config As Object, saved As Variant, settings As CSystemSettingsReader
    Dim reader As CSectionPlotDataReader, plotter As CSectionPlotter, chart As Object
    Dim diagramMode As String
    On Error GoTo Failed
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = config.Formula
    diagramMode = ReadDiagramMode()
    For Each fixture In Fixtures()
        modTestShapeArchitecture.ConfigureShapeFixture CStr(fixture)
        modTestPerformance.ConfigurePerformanceFixture "DIRECT", 1, "Newton", "Saved", "Normal", diagramMode
        SetSetting "Plot.Enabled", "Yes": SetSetting "Plot.AutoUpdateAfterCalculation", "Yes"
        SetSetting "Plot.LoadCase", "Worst": SetSetting "General.ExecutionReportEnabled", "Yes"
        SetProfile "Visualization.State", "StrengthState": SetProfile "Visualization.Quantity", "Stress"
        label = CStr(fixture)
        result = result & "FAMILY_WORKFLOW: " & label & "; diagramExtension=" & ReadDiagramMode() & "; " & Replace(modTestPerformance.MeasurePerformanceBatch("Full"), Chr$(30), "; ") & vbCrLf
        If Len(outputDirectory) > 0 Then
            Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
            Set reader = New CSectionPlotDataReader: reader.LoadFromWorkbook ThisWorkbook, settings
            Set plotter = New CSectionPlotter
            before = SectionEquilibriumSolveCount()
            plotter.Draw ThisWorkbook, reader, settings
            Set chart = ThisWorkbook.Worksheets("Расчет").ChartObjects("chtNDMSectionPlot")
            If Not chart.Chart.Export(outputDirectory & Application.PathSeparator & label & "_Stress.png", "PNG") Then Err.Raise 5, , "PNG export failed"
            SetProfile "Visualization.Quantity", "Strain"
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
    MeasureFamilyWorkflow = result
End Function

' ДЛЯ ТЕСТОВ: фазовый замер одного HollowOffset для объяснения общей
' разницы времени. Стратегии решения и число сочетаний прежние.
Public Function MeasureFamilyStages() As String
    Dim diagramMode As String
    diagramMode = ReadDiagramMode()
    modTestShapeArchitecture.ConfigureShapeFixture "HollowOffset"
    modTestPerformance.ConfigurePerformanceFixture "DIRECT", 1, "Newton", "Saved", "Normal", diagramMode
    SetSetting "Plot.Enabled", "Yes": SetSetting "Plot.AutoUpdateAfterCalculation", "Yes"
    SetSetting "Plot.LoadCase", "Worst": SetSetting "General.ExecutionReportEnabled", "Yes"
    SetProfile "Visualization.State", "StrengthState": SetProfile "Visualization.Quantity", "Stress"
    Dim i As Long, result As String
    For i = 1 To 3
        result = result & "FAMILY_STAGES: HollowOffset.sample" & CStr(i) & "; diagramExtension=" & ReadDiagramMode() & "; " & Replace(modTestPerformance.MeasurePerformanceBatch("Staged"), Chr$(30), "; ") & vbCrLf
    Next i
    MeasureFamilyStages = result
End Function

' ДЛЯ ТЕСТОВ: один реальный полный LC с capacity/crack/stability,
' report и схемой. Это направленный замер, не полный on/off набор.
Public Function MeasureMixedWorkflow() As String
    Dim diagramMode As String
    diagramMode = ReadDiagramMode()
    modTestShapeArchitecture.ConfigureShapeFixture "HollowOffset"
    modTestPerformance.ConfigurePerformanceFixture "MIXED", 1, "Newton", "Saved", "Normal", diagramMode
    SetSetting "Stability.ElementLength", "1000"
    SetSetting "Plot.Enabled", "Yes": SetSetting "Plot.AutoUpdateAfterCalculation", "Yes"
    SetSetting "Plot.LoadCase", "Worst": SetSetting "General.ExecutionReportEnabled", "Yes"
    SetProfile "Visualization.State", "StrengthState": SetProfile "Visualization.Quantity", "Stress"
    Dim settings As CSystemSettingsReader, units As CUnitSystem, loads As Object
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    Set loads = ThisWorkbook.Names("rngLoadCombinations").RefersToRange
    loads.Cells(2, 2).Value2 = -50000# / units.InputForceToInternal(1#)
    loads.Cells(2, 3).Value2 = 150000000# / units.InputMomentMxToInternal(1#)
    loads.Cells(2, 4).Value2 = -30000000# / units.InputMomentMyToInternal(1#)
    MeasureMixedWorkflow = "MIXED_WORKFLOW: diagramExtension=" & ReadDiagramMode() & "; " & Replace(modTestPerformance.MeasurePerformanceBatch("Full"), Chr$(30), "; ")
End Function

' ДЛЯ ТЕСТОВ: current-projection gate требует применимого СП35 НДС;
' пользовательская книга может содержать AutoCAD/иной профиль без этой ветви.
Public Function RunSP35ProjectionFixture() As String
    Dim diagramMode As String
    diagramMode = ReadDiagramMode()
    modTestShapeArchitecture.ConfigureShapeFixture "HollowOffset"
    modTestPerformance.ConfigurePerformanceFixture "CRACK_SP35", 1, "Newton", "Saved", "Normal", diagramMode
    Dim settings As CSystemSettingsReader, units As CUnitSystem, loads As Object
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    Set loads = ThisWorkbook.Names("rngLoadCombinations").RefersToRange
    loads.Cells(2, 2).Value2 = -50000# / units.InputForceToInternal(1#)
    loads.Cells(2, 3).Value2 = 800000000# / units.InputMomentMxToInternal(1#)
    loads.Cells(2, 4).Value2 = -100000000# / units.InputMomentMyToInternal(1#)
    RunSP35ProjectionFixture = "PROJECTION_DIAGRAM_MODE: " & ReadDiagramMode() & vbCrLf & modTestSP35CrackWidth.RunSP35CurrentProjectionTests()
End Function

' Сохраняет фактический режим runner-а до подготовки конкретной fixture.
Private Function ReadDiagramMode() As String
    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook
    If settings.GetRequiredBoolean("General.DiagramExtension") Then
        ReadDiagramMode = "Yes"
    Else
        ReadDiagramMode = "No"
    End If
End Function

Private Function Fixtures() As Variant
    Fixtures = Array("Circle", "Rectangle", "LSection", "TwoRectangles", "RoundedSimple", "RoundedTapered", "RoundedMixed", "HollowCentered", "HollowOffset")
End Function

' ДЛЯ ТЕСТОВ: специальное покрытие отображения без полного UI-прогона.
' Два знака изгиба дают реальные смешанные поля; последующие операции
' только читают Results, включая изменение Config, resize и отсутствие НДС.
Public Function RunDirectedPlotModes(ByVal outputDirectory As String) As String
    Dim config As Object, saved As Variant, settings As CSystemSettingsReader, units As CUnitSystem
    Dim loads As Object, reader As CSectionPlotDataReader, plotter As CSectionPlotter
    Dim container As Object, marker As Object, state As Variant, quantity As Variant
    Dim before As Long, count As Long, minValue As Double, maxValue As Double, i As Long
    Dim initialCount As Long, originalWidth As Double, originalHeight As Double, label As String
    Dim diagramMode As String
    ResetReport
    On Error GoTo Failed
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = config.Formula
    diagramMode = ReadDiagramMode()
    modTestShapeArchitecture.ConfigureShapeFixture "HollowOffset"
    modTestPerformance.ConfigurePerformanceFixture "DIRECT", 2, "Newton", "Saved", "Normal", diagramMode
    Check "plotModes.diagramMode", ReadDiagramMode() = diagramMode
    mReport = mReport & "PLOT_DIAGRAM_MODE: " & ReadDiagramMode() & vbCrLf
    SetSetting "Plot.Enabled", "Yes": SetSetting "Plot.AutoUpdateAfterCalculation", "Yes"
    SetSetting "Plot.LoadCase", "PERF_01": SetSetting "General.ExecutionReportEnabled", "No"
    SetProfile "Visualization.State", "StrengthState"
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    Set loads = ThisWorkbook.Names("rngLoadCombinations").RefersToRange
    For i = 2 To 3
        loads.Cells(i, 2).Value2 = -50000# / units.InputForceToInternal(1#)
        loads.Cells(i, 3).Value2 = ((-1#) ^ i) * 70000000# / units.InputMomentMxToInternal(1#)
        loads.Cells(i, 4).Value2 = ((-1#) ^ (i + 1)) * 30000000# / units.InputMomentMyToInternal(1#)
    Next i
    mReport = mReport & modTestPerformance.MeasurePerformanceBatch("Full") & vbCrLf
    Set reader = New CSectionPlotDataReader: Set plotter = New CSectionPlotter
    before = SectionEquilibriumSolveCount()
    Set container = ThisWorkbook.Worksheets("Расчет").ChartObjects("chtNDMSectionPlot")
    originalWidth = container.Width: originalHeight = container.Height
    Set marker = container.Chart.Shapes.AddShape(1, 10#, 10#, 10#, 10#): marker.Name = "User_QA_shape"
    For Each state In Array("PERF_01", "PERF_02", "Worst")
        SetSetting "Plot.LoadCase", CStr(state)
        For Each quantity In Array("Stress", "Strain")
            SetProfile "Visualization.Quantity", CStr(quantity)
            settings.LoadFromWorkbook ThisWorkbook: reader.LoadFromWorkbook ThisWorkbook, settings
            Check CStr(state) & "." & CStr(quantity) & ".state", Not reader.GeometryOnly
            minValue = 1E+100: maxValue = -1E+100
            For i = 1 To reader.Count
                If reader.ResultValue(i) < minValue Then minValue = reader.ResultValue(i)
                If reader.ResultValue(i) > maxValue Then maxValue = reader.ResultValue(i)
            Next i
            Check CStr(state) & "." & CStr(quantity) & ".signed", minValue < 0# And maxValue > 0#
            plotter.Draw ThisWorkbook, reader, settings
            initialCount = container.Chart.Shapes.Count
            plotter.Draw ThisWorkbook, reader, settings
            Check CStr(state) & "." & CStr(quantity) & ".repeat", container.Chart.Shapes.Count = initialCount
            Check CStr(state) & "." & CStr(quantity) & ".foreign", container.Chart.Shapes("User_QA_shape").Name = "User_QA_shape"
            label = CStr(state) & "_" & CStr(quantity)
            If Not container.Chart.Export(outputDirectory & Application.PathSeparator & label & ".png", "PNG") Then Err.Raise 5, , "PNG export failed"
            mReport = mReport & "PLOT_MODE: " & label & "; LC=" & reader.LoadCase & "; elements=" & CStr(reader.Count) & _
                "; shapes=" & CStr(initialCount) & "; signed=" & CStr(minValue) & "," & CStr(maxValue) & vbCrLf
        Next quantity
    Next state
    count = reader.Count
    SetSetting "Geometry.Type", "Circle"
    settings.LoadFromWorkbook ThisWorkbook: reader.LoadFromWorkbook ThisWorkbook, settings
    Check "saved.afterConfigChange", reader.Count = count
    SetAnnotationFlags "No"
    settings.LoadFromWorkbook ThisWorkbook
    container.Width = 650#: container.Height = 500#
    plotter.Draw ThisWorkbook, reader, settings
    If Not container.Chart.Export(outputDirectory & Application.PathSeparator & "Resize_NoLabels.png", "PNG") Then Err.Raise 5, , "PNG export failed"
    Check "passive.noSolve", SectionEquilibriumSolveCount() = before
    ' Второй штатный запуск без direct/capacity/crack действительно не имеет
    ' выбранного StrengthState; схема показывает только сохраненную геометрию.
    SetSetting "Geometry.Type", "HollowRectangle"
    SetProfile "Calculation.Strength.DirectState", "No": SetProfile "Calculation.Stability.Enabled", "Yes"
    mReport = mReport & modTestPerformance.MeasurePerformanceBatch("Full") & vbCrLf
    settings.LoadFromWorkbook ThisWorkbook: reader.LoadGeometryOnlyForMissingState ThisWorkbook, settings
    Check "missingState.geometryOnly", reader.GeometryOnly
    before = SectionEquilibriumSolveCount()
    plotter.Draw ThisWorkbook, reader, settings
    Check "missingState.noSolve", SectionEquilibriumSolveCount() = before
    If Not container.Chart.Export(outputDirectory & Application.PathSeparator & "MissingState.png", "PNG") Then Err.Raise 5, , "PNG export failed"
    GoTo Restore
Failed:
    Check "plotModes.runtime." & CStr(Err.Number) & "." & Err.Description, False
Restore:
    If IsArray(saved) Then config.Formula = saved
    If originalWidth > 0# Then container.Width = originalWidth: container.Height = originalHeight
    RunDirectedPlotModes = FinishReport("DIRECTED_PLOT_MODES")
End Function

Private Sub SetAnnotationFlags(ByVal value As String)
    Dim table As Object, k As Long
    Set table = ThisWorkbook.Names("rngPlotAnnotationSettings").RefersToRange
    For k = 1 To table.Rows.Count
        If CStr(table.Cells(k, 1).Value2) = "Enabled" Then
            table.Cells(k, 2).Value2 = value: table.Cells(k, 3).Value2 = value: Exit Sub
        End If
    Next k
    Err.Raise 5, , "Missing annotation Enabled row"
End Sub

Private Sub SetSetting(ByVal key As String, ByVal value As String)
    Dim table As Object, k As Long
    Set table = ThisWorkbook.Names("rngSystemSettings").RefersToRange
    For k = 1 To table.Rows.Count
        If CStr(table.Cells(k, 1).Value2) = key Then table.Cells(k, 2).Value2 = value: Exit Sub
    Next k
    Err.Raise 5, "modFinalGeometryChecks", "Missing setting: " & key
End Sub
Private Sub SetProfile(ByVal key As String, ByVal value As String)
    Dim table As Object, k As Long
    Set table = ThisWorkbook.Names("rngCalculationProfiles").RefersToRange
    For k = 1 To table.Rows.Count
        If CStr(table.Cells(k, 2).Value2) = key Then table.Cells(k, 3).Value2 = value: Exit Sub
    Next k
    Err.Raise 5, "modFinalGeometryChecks", "Missing profile key: " & key
End Sub
Private Sub ResetReport()
    mReport = vbNullString: mPassed = 0: mFailed = 0
End Sub
Private Sub Check(ByVal label As String, ByVal success As Boolean)
    If success Then mPassed = mPassed + 1: mReport = mReport & "OK: " & label & vbCrLf Else mFailed = mFailed + 1: mReport = mReport & "FAIL: " & label & vbCrLf
End Sub
Private Function FinishReport(ByVal label As String) As String
    FinishReport = mReport & "TOTAL_" & label & ": passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function
