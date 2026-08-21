Attribute VB_Name = "modAutoCADStressExport"
Option Explicit

Private Type TAutoCADExportSettings
    ConcreteLayer As String
    RebarLayer As String
    ConcreteTensionLayer As String
    ConcreteCompressionLayer As String
    RebarTensionLayer As String
    RebarCompressionLayer As String
    ConcreteTensionColor As Long
    ConcreteCompressionColor As Long
    RebarTensionColor As Long
    RebarCompressionColor As Long
    NeutralColor As Long
    IncludeElementNames As Boolean
    NeutralLineEnabled As Boolean
    PrincipalAxesEnabled As Boolean
    LoadPointEnabled As Boolean
End Type

Public Sub ExportSectionStressToAutoCAD()
    On Error GoTo Failed

    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook
    Dim units As CUnitSystem
    Set units = New CUnitSystem
    units.LoadFromSettings settings

    Dim section As CSectionModel
    Dim stressByID As Object
    Dim combinationID As String
    Dim epsilon0 As Double
    Dim kappaX As Double
    Dim kappaY As Double
    Dim loadReferenceX As Double
    Dim loadReferenceY As Double
    ReadResultsExportState ThisWorkbook, settings, units, section, stressByID, combinationID, _
        epsilon0, kappaX, kappaY, loadReferenceX, loadReferenceY

    Dim exportSettings As TAutoCADExportSettings
    exportSettings = ReadAutoCADExportSettings(settings)

    Dim concrete As CConcreteDiagramMaterial
    Set concrete = New CConcreteDiagramMaterial
    concrete.InitializeFromSettings settings, units

    Dim steel As CSteelDiagramMaterial
    Set steel = New CSteelDiagramMaterial
    steel.InitializeFromSettings settings, units

    DrawResultsStressExport section, concrete, steel, stressByID, epsilon0, kappaX, kappaY, _
        loadReferenceX, loadReferenceY, exportSettings
    MsgBox "Экспорт в AutoCAD завершен. Волокон бетона: " & CStr(section.ConcreteCount) & _
        "; стержней арматуры: " & CStr(section.RebarCount) & _
        "; сочетание: " & combinationID, vbInformation, "RC Section NDM"
    Exit Sub

Failed:
    MsgBox "Экспорт в AutoCAD не выполнен: " & Err.Description, vbExclamation, "RC Section NDM"
End Sub

Public Sub PrepareAutoCADExportState(ByVal workbook As Object, ByRef section As CSectionModel, _
        ByRef concrete As CConcreteDiagramMaterial, _
        ByRef steel As CSteelDiagramMaterial, ByRef solver As CSectionSolver, _
        ByRef loadReferenceX As Double, ByRef loadReferenceY As Double)
    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook workbook
    Dim units As CUnitSystem
    Set units = New CUnitSystem
    units.LoadFromSettings settings

    Set section = BuildWorkbookSectionModel(workbook, settings, units)

    Set concrete = New CConcreteDiagramMaterial
    concrete.InitializeFromSettings settings, units

    Set steel = New CSteelDiagramMaterial
    steel.InitializeFromSettings settings, units

    Dim nValue As Double
    Dim userMxValue As Double
    Dim userMyValue As Double
    ReadFirstExportLoad workbook, nValue, userMxValue, userMyValue
    nValue = units.InputForceToInternal(nValue)
    userMxValue = units.InputMomentMxToInternal(userMxValue)
    userMyValue = units.InputMomentMyToInternal(userMyValue)

    Dim referenceX As Double
    Dim referenceY As Double
    CalculateTransformedSectionCentroid section, concrete, steel, referenceX, referenceY
    loadReferenceX = referenceX + units.InputLengthToInternal(settings.GetDouble("Load.ReferenceOffsetX", 0#))
    loadReferenceY = referenceY + units.InputLengthToInternal(settings.GetDouble("Load.ReferenceOffsetY", 0#))

    Dim internalMxValue As Double
    Dim internalMyValue As Double
    internalMxValue = userMxValue + nValue * loadReferenceY
    internalMyValue = userMyValue + nValue * loadReferenceX

    Set solver = New CSectionSolver
    solver.ApplySettings settings, units
    solver.Solve section, concrete, steel, nValue, internalMxValue, internalMyValue
    If Not solver.Converged Then Err.Raise vbObjectError + 4300, "PrepareAutoCADExportState", solver.StopReason
End Sub

Private Sub ReadResultsExportState(ByVal workbook As Object, ByVal settings As CSystemSettingsReader, _
        ByVal units As CUnitSystem, _
        ByRef section As CSectionModel, ByRef stressByID As Object, ByRef combinationID As String, _
        ByRef epsilon0 As Double, ByRef kappaX As Double, ByRef kappaY As Double, _
        ByRef loadReferenceX As Double, ByRef loadReferenceY As Double)
    Set section = ReadSectionGeometryFromResults(workbook, units)

    combinationID = ResolveExportCombinationID(workbook, settings.GetRawString("AutoCAD.Export.CombinationID", "Worst"))

    Set stressByID = CreateObject("Scripting.Dictionary")
    stressByID.CompareMode = vbTextCompare
    ReadElementResultsForCombination workbook, settings.GetString("AutoCAD.Export.ResultType", "Stress"), combinationID, stressByID
    ReadSectionPropertiesForCombination workbook, units, combinationID, epsilon0, kappaX, kappaY, _
        loadReferenceX, loadReferenceY
End Sub

Private Function ReadSectionGeometryFromResults(ByVal workbook As Object, ByVal units As CUnitSystem) As CSectionModel
    Dim anchor As Object
    Set anchor = workbook.Names.Item("rngNDMSectionGeometry").RefersToRange

    Dim data As Variant
    data = anchor.CurrentRegion.Value2
    If Not HasResultTableRows(data) Then Err.Raise vbObjectError + 4350, "ReadSectionGeometryFromResults", _
        "На листе Results нет таблицы расчетной геометрии. Сначала выполните расчет."

    Dim colID As Long: colID = ResultColumn(data, "ElementID")
    Dim colType As Long: colType = ResultColumn(data, "MaterialType")
    Dim colClass As Long: colClass = ResultColumn(data, "MaterialClass")
    Dim colX As Long: colX = ResultColumn(data, "X")
    Dim colY As Long: colY = ResultColumn(data, "Y")
    Dim colArea As Long: colArea = ResultColumn(data, "Area")
    Dim colShape As Long: colShape = ResultColumn(data, "ShapeType")
    Dim colWidth As Long: colWidth = ResultColumn(data, "Width")
    Dim colHeight As Long: colHeight = ResultColumn(data, "Height")
    Dim colDiameter As Long: colDiameter = ResultColumn(data, "Diameter")
    Dim colRotation As Long: colRotation = ResultColumn(data, "Rotation")
    Dim colLocalIx As Long: colLocalIx = ResultColumn(data, "LocalIx")
    Dim colLocalIy As Long: colLocalIy = ResultColumn(data, "LocalIy")
    Dim colLocalIxy As Long: colLocalIxy = ResultColumn(data, "LocalIxy")
    Dim colComment As Long: colComment = ResultColumn(data, "Comment")

    Dim model As CSectionModel
    Set model = New CSectionModel
    model.SourceType = "Results"

    Dim rowIndex As Long
    For rowIndex = 2 To UBound(data, 1)
        If Len(Trim$(CStr(data(rowIndex, colID)))) > 0 Then
            If StrComp(CStr(data(rowIndex, colType)), "Concrete", vbTextCompare) = 0 Then
                model.AddConcreteElement OutputLengthToInternal(CDbl(data(rowIndex, colX)), units), OutputLengthToInternal(CDbl(data(rowIndex, colY)), units), _
                    OutputAreaToInternal(CDbl(data(rowIndex, colArea)), units), 1, vbNullString, vbNullString, _
                    CStr(data(rowIndex, colShape)), OutputLengthToInternal(CDbl(Val(CStr(data(rowIndex, colWidth)))), units), _
                    OutputLengthToInternal(CDbl(Val(CStr(data(rowIndex, colHeight)))), units), CDbl(Val(CStr(data(rowIndex, colRotation)))), _
                    CStr(data(rowIndex, colComment)), OutputFourthPowerLengthToInternal(CDbl(Val(CStr(data(rowIndex, colLocalIx)))), units), _
                    OutputFourthPowerLengthToInternal(CDbl(Val(CStr(data(rowIndex, colLocalIy)))), units), OutputFourthPowerLengthToInternal(CDbl(Val(CStr(data(rowIndex, colLocalIxy)))), units)
            ElseIf StrComp(CStr(data(rowIndex, colType)), "Rebar", vbTextCompare) = 0 Then
                model.AddRebarElement OutputLengthToInternal(CDbl(data(rowIndex, colX)), units), OutputLengthToInternal(CDbl(data(rowIndex, colY)), units), _
                    OutputLengthToInternal(CDbl(data(rowIndex, colDiameter)), units), OutputAreaToInternal(CDbl(data(rowIndex, colArea)), units), _
                    CStr(data(rowIndex, colClass)), 1, vbNullString, _
                    vbNullString, CStr(data(rowIndex, colComment))
            End If
        End If
    Next rowIndex

    If model.ConcreteCount <= 0 Then Err.Raise vbObjectError + 4351, "ReadSectionGeometryFromResults", _
        "В таблице Results нет бетонных элементов."
    Set ReadSectionGeometryFromResults = model
End Function

Private Sub ReadElementResultsForCombination(ByVal workbook As Object, ByVal resultType As String, ByRef combinationID As String, _
        ByVal stressByID As Object)
    Dim anchor As Object
    Set anchor = workbook.Names.Item("rngNDMElementResults").RefersToRange

    Dim data As Variant
    data = anchor.CurrentRegion.Value2
    If Not HasResultTableRows(data) Then Err.Raise vbObjectError + 4352, "ReadElementResultsForCombination", _
        "На листе Results нет таблицы результатов НДМ. Сначала выполните расчет."

    Dim colCombination As Long: colCombination = ResultColumn(data, "LoadCase")
    Dim colID As Long: colID = ResultColumn(data, "ElementID")
    Dim colValue As Long
    Select Case LCase$(Trim$(resultType))
        Case "stress"
            colValue = ResultColumn(data, "Stress")
        Case "strain"
            colValue = ResultColumn(data, "Strain")
        Case Else
            Err.Raise vbObjectError + 4359, "ReadElementResultsForCombination", _
                "AutoCAD.Export.ResultType должен быть Stress или Strain."
    End Select

    If Len(combinationID) = 0 Then Err.Raise vbObjectError + 4353, "ReadElementResultsForCombination", _
        "В Results нет рассчитанных сочетаний для экспорта."

    Dim found As Boolean
    Dim rowIndex As Long
    For rowIndex = 2 To UBound(data, 1)
        If StrComp(CStr(data(rowIndex, colCombination)), combinationID, vbTextCompare) = 0 Then
            found = True
            stressByID(CStr(data(rowIndex, colID))) = CDbl(data(rowIndex, colValue))
        End If
    Next rowIndex

    If Not found Then Err.Raise vbObjectError + 4354, "ReadElementResultsForCombination", _
        "В Results не найдено сочетание для экспорта: " & combinationID
End Sub

Private Sub ReadSectionPropertiesForCombination(ByVal workbook As Object, ByVal units As CUnitSystem, _
        ByVal combinationID As String, ByRef epsilon0 As Double, ByRef kappaX As Double, _
        ByRef kappaY As Double, ByRef loadReferenceX As Double, ByRef loadReferenceY As Double)
    Dim data As Variant
    data = workbook.Names.Item("rngNDMSectionProperties").RefersToRange.CurrentRegion.Value2
    If Not HasResultTableRows(data) Then Err.Raise vbObjectError + 4360, "ReadSectionPropertiesForCombination", _
        "На листе Results нет rngNDMSectionProperties. Сначала выполните расчет."

    Dim colLoadCase As Long: colLoadCase = ResultColumn(data, "LoadCase")
    Dim colParameter As Long: colParameter = ResultColumn(data, "Parameter")
    Dim colValue As Long: colValue = ResultColumn(data, "Value")

    Dim foundState As Boolean
    Dim rowIndex As Long
    For rowIndex = 2 To UBound(data, 1)
        If StrComp(CStr(data(rowIndex, colLoadCase)), combinationID, vbTextCompare) = 0 Then
            Select Case LCase$(Trim$(CStr(data(rowIndex, colParameter))))
                Case "epsilon0"
                    epsilon0 = CDbl(data(rowIndex, colValue))
                    foundState = True
                Case "kappax"
                    kappaX = OutputCurvatureToInternal(CDbl(data(rowIndex, colValue)), units)
                Case "kappay"
                    kappaY = OutputCurvatureToInternal(CDbl(data(rowIndex, colValue)), units)
                Case "loadreferencex"
                    loadReferenceX = OutputLengthToInternal(CDbl(data(rowIndex, colValue)), units)
                Case "loadreferencey"
                    loadReferenceY = OutputLengthToInternal(CDbl(data(rowIndex, colValue)), units)
            End Select
        End If
    Next rowIndex

    If Not foundState Then Err.Raise vbObjectError + 4361, "ReadSectionPropertiesForCombination", _
        "В rngNDMSectionProperties нет состояния для сочетания: " & combinationID
End Sub

Private Function ReadGoverningCombinationID(ByVal workbook As Object) As String
    On Error GoTo Failed
    Dim data As Variant
    data = workbook.Names.Item("rngBatchSummary").RefersToRange.Value2
    If UBound(data, 1) >= 2 And UBound(data, 2) >= 2 Then
        ReadGoverningCombinationID = Trim$(CStr(data(2, 2)))
    End If
Failed:
End Function

Private Function OutputLengthToInternal(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then OutputLengthToInternal = value Else OutputLengthToInternal = units.OutputLengthToInternal(value)
End Function

Private Function OutputAreaToInternal(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then OutputAreaToInternal = value Else OutputAreaToInternal = units.OutputAreaToInternal(value)
End Function

Private Function OutputFourthPowerLengthToInternal(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then OutputFourthPowerLengthToInternal = value Else OutputFourthPowerLengthToInternal = units.OutputFourthPowerLengthToInternal(value)
End Function

Private Function OutputCurvatureToInternal(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then OutputCurvatureToInternal = value Else OutputCurvatureToInternal = units.OutputCurvatureToInternal(value)
End Function

Private Function ResolveExportCombinationID(ByVal workbook As Object, ByVal settingValue As String) As String
    Dim valueText As String
    valueText = Trim$(settingValue)
    If Len(valueText) = 0 Then Err.Raise vbObjectError + 4357, "ResolveExportCombinationID", _
        "AutoCAD.Export.CombinationID должен быть Worst или точным CombinationID из Results."

    If StrComp(valueText, "Worst", vbTextCompare) = 0 Then
        ResolveExportCombinationID = ReadGoverningCombinationID(workbook)
        If Len(ResolveExportCombinationID) = 0 Then Err.Raise vbObjectError + 4358, "ResolveExportCombinationID", _
            "В rngBatchSummary не найдено определяющее сочетание для AutoCAD.Export.CombinationID = Worst."
    Else
        ResolveExportCombinationID = valueText
    End If
End Function

Private Function HasResultTableRows(ByVal data As Variant) As Boolean
    On Error GoTo Failed
    HasResultTableRows = (UBound(data, 1) >= 2 And UBound(data, 2) >= 1)
    Exit Function
Failed:
    HasResultTableRows = False
End Function

Private Function ResultColumn(ByRef data As Variant, ByVal headerName As String) As Long
    Dim colIndex As Long
    For colIndex = 1 To UBound(data, 2)
        If StrComp(ResultHeaderBase(CStr(data(1, colIndex))), headerName, vbTextCompare) = 0 Then
            ResultColumn = colIndex
            Exit Function
        End If
    Next colIndex
    Err.Raise vbObjectError + 4355, "ResultColumn", "В Results не найден столбец: " & headerName
End Function

Private Function ResultHeaderBase(ByVal headerText As String) As String
    Dim commaPos As Long
    commaPos = InStr(1, headerText, ",", vbTextCompare)
    If commaPos > 0 Then
        ResultHeaderBase = Trim$(Left$(headerText, commaPos - 1))
    Else
        ResultHeaderBase = Trim$(headerText)
    End If
End Function

Private Sub DrawResultsStressExport(ByVal section As CSectionModel, _
        ByVal concrete As Object, ByVal steel As Object, ByVal stressByID As Object, _
        ByVal epsilon0 As Double, ByVal kappaX As Double, ByVal kappaY As Double, _
        ByVal loadReferenceX As Double, ByVal loadReferenceY As Double, _
        ByRef exportSettings As TAutoCADExportSettings)
    Dim acad As Object
    Set acad = GetObject(, "AutoCAD.Application")
    If acad Is Nothing Then Err.Raise vbObjectError + 4310, "DrawResultsStressExport", "Откройте AutoCAD и активный чертеж."

    Dim doc As Object
    Set doc = acad.ActiveDocument
    If doc Is Nothing Then Err.Raise vbObjectError + 4311, "DrawResultsStressExport", "В AutoCAD нет активного чертежа."

    Dim ms As Object
    Set ms = doc.ModelSpace

    EnsureAcadLayer doc, exportSettings.ConcreteLayer, 8
    EnsureAcadLayer doc, exportSettings.RebarLayer, 1
    EnsureAcadLayer doc, exportSettings.ConcreteTensionLayer, exportSettings.ConcreteTensionColor
    EnsureAcadLayer doc, exportSettings.ConcreteCompressionLayer, exportSettings.ConcreteCompressionColor
    EnsureAcadLayer doc, exportSettings.RebarTensionLayer, exportSettings.RebarTensionColor
    EnsureAcadLayer doc, exportSettings.RebarCompressionLayer, exportSettings.RebarCompressionColor
    EnsureAcadLayer doc, "RC_NDM_Axes", 3
    EnsureAcadLayer doc, "RC_NDM_LoadPoint", 2
    EnsureAcadLayer doc, "RC_NDM_NeutralLine", exportSettings.NeutralColor

    Dim i As Long
    Dim stress As Double
    Dim textHeight As Double

    For i = 1 To section.ConcreteCount
        Dim concreteWidth As Double
        Dim concreteHeight As Double
        concreteWidth = ConcreteDrawWidth(section, i)
        concreteHeight = ConcreteDrawHeight(section, i)
        stress = ResultStress(stressByID, section.ConcreteID(i))
        textHeight = 0.22 * MinDouble(concreteWidth, concreteHeight)
        If textHeight <= 0# Then textHeight = 1#
        AddAcadRectangleRegion ms, section.ConcreteX(i), section.ConcreteY(i), concreteWidth, concreteHeight, _
            exportSettings.ConcreteLayer, StressColor(stress, exportSettings.ConcreteTensionColor, _
            exportSettings.ConcreteCompressionColor, exportSettings.NeutralColor)
        AddAcadText ms, StressLabelText(section.ConcreteID(i), stress, exportSettings.IncludeElementNames), _
            section.ConcreteX(i) - 0.45 * concreteWidth, _
            section.ConcreteY(i) - 0.1 * concreteHeight, textHeight, _
            StressAnnotationLayer(stress, exportSettings.ConcreteTensionLayer, exportSettings.ConcreteCompressionLayer), _
            StressColor(stress, exportSettings.ConcreteTensionColor, _
            exportSettings.ConcreteCompressionColor, exportSettings.NeutralColor)
    Next i

    For i = 1 To section.RebarCount
        stress = ResultStress(stressByID, section.RebarID(i))
        AddAcadCircleRegion ms, section.RebarX(i), section.RebarY(i), section.RebarDiameter(i) / 2#, _
            exportSettings.RebarLayer, StressColor(stress, exportSettings.RebarTensionColor, _
            exportSettings.RebarCompressionColor, exportSettings.NeutralColor)
        AddAcadText ms, StressLabelText(section.RebarID(i), stress, exportSettings.IncludeElementNames), _
            section.RebarX(i) + section.RebarDiameter(i) / 2#, section.RebarY(i) + section.RebarDiameter(i) / 2#, _
            MaxDouble(2.5, section.RebarDiameter(i) * 0.18), _
            StressAnnotationLayer(stress, exportSettings.RebarTensionLayer, exportSettings.RebarCompressionLayer), _
            StressColor(stress, exportSettings.RebarTensionColor, _
            exportSettings.RebarCompressionColor, exportSettings.NeutralColor)
    Next i

    DrawCentroidAxesAndLoadPoint ms, section, concrete, steel, loadReferenceX, loadReferenceY, _
        exportSettings.PrincipalAxesEnabled, exportSettings.LoadPointEnabled
    If exportSettings.NeutralLineEnabled Then
        DrawNeutralLineByState ms, section, epsilon0, kappaX, kappaY, exportSettings.NeutralColor
    End If

    doc.Regen 1
End Sub

Private Function ResultStress(ByVal stressByID As Object, ByVal elementID As String) As Double
    If stressByID.Exists(elementID) Then
        ResultStress = CDbl(stressByID.Item(elementID))
    Else
        Err.Raise vbObjectError + 4356, "ResultStress", "В Results нет напряжения для элемента: " & elementID
    End If
End Function

Private Sub DrawStressExport(ByVal section As CSectionModel, _
        ByVal concrete As Object, ByVal steel As Object, ByVal solver As CSectionSolver, _
        ByVal loadReferenceX As Double, ByVal loadReferenceY As Double, ByRef exportSettings As TAutoCADExportSettings)
    Dim acad As Object
    Set acad = GetObject(, "AutoCAD.Application")
    If acad Is Nothing Then Err.Raise vbObjectError + 4310, "DrawStressExport", "Откройте AutoCAD и активный чертеж."

    Dim doc As Object
    Set doc = acad.ActiveDocument
    If doc Is Nothing Then Err.Raise vbObjectError + 4311, "DrawStressExport", "В AutoCAD нет активного чертежа."

    Dim ms As Object
    Set ms = doc.ModelSpace

    EnsureAcadLayer doc, exportSettings.ConcreteLayer, 8
    EnsureAcadLayer doc, exportSettings.RebarLayer, 1
    EnsureAcadLayer doc, exportSettings.ConcreteTensionLayer, exportSettings.ConcreteTensionColor
    EnsureAcadLayer doc, exportSettings.ConcreteCompressionLayer, exportSettings.ConcreteCompressionColor
    EnsureAcadLayer doc, exportSettings.RebarTensionLayer, exportSettings.RebarTensionColor
    EnsureAcadLayer doc, exportSettings.RebarCompressionLayer, exportSettings.RebarCompressionColor
    EnsureAcadLayer doc, "RC_NDM_Axes", 3
    EnsureAcadLayer doc, "RC_NDM_LoadPoint", 2
    EnsureAcadLayer doc, "RC_NDM_NeutralLine", exportSettings.NeutralColor

    Dim i As Long
    Dim strain As Double
    Dim stress As Double
    Dim textHeight As Double

    For i = 1 To section.ConcreteCount
        Dim concreteWidth As Double
        Dim concreteHeight As Double
        concreteWidth = ConcreteDrawWidth(section, i)
        concreteHeight = ConcreteDrawHeight(section, i)
        strain = solver.Epsilon0 + solver.KappaX * section.ConcreteY(i) + solver.KappaY * section.ConcreteX(i)
        stress = concrete.GetStress(strain)
        textHeight = 0.22 * MinDouble(concreteWidth, concreteHeight)
        If textHeight <= 0# Then textHeight = 1#
        AddAcadRectangleRegion ms, section.ConcreteX(i), section.ConcreteY(i), concreteWidth, concreteHeight, _
            exportSettings.ConcreteLayer, StressColor(stress, exportSettings.ConcreteTensionColor, _
            exportSettings.ConcreteCompressionColor, exportSettings.NeutralColor)
        AddAcadText ms, StressLabelText(section.ConcreteID(i), stress, exportSettings.IncludeElementNames), _
            section.ConcreteX(i) - 0.45 * concreteWidth, _
            section.ConcreteY(i) - 0.1 * concreteHeight, textHeight, _
            StressAnnotationLayer(stress, exportSettings.ConcreteTensionLayer, exportSettings.ConcreteCompressionLayer), _
            StressColor(stress, exportSettings.ConcreteTensionColor, _
            exportSettings.ConcreteCompressionColor, exportSettings.NeutralColor)
    Next i

    For i = 1 To section.RebarCount
        strain = solver.Epsilon0 + solver.KappaX * section.RebarY(i) + solver.KappaY * section.RebarX(i)
        stress = steel.GetStress(strain)
        AddAcadCircleRegion ms, section.RebarX(i), section.RebarY(i), section.RebarDiameter(i) / 2#, _
            exportSettings.RebarLayer, StressColor(stress, exportSettings.RebarTensionColor, _
            exportSettings.RebarCompressionColor, exportSettings.NeutralColor)
        AddAcadText ms, StressLabelText(section.RebarID(i), stress, exportSettings.IncludeElementNames), _
            section.RebarX(i) + section.RebarDiameter(i) / 2#, section.RebarY(i) + section.RebarDiameter(i) / 2#, _
            MaxDouble(2.5, section.RebarDiameter(i) * 0.18), _
            StressAnnotationLayer(stress, exportSettings.RebarTensionLayer, exportSettings.RebarCompressionLayer), _
            StressColor(stress, exportSettings.RebarTensionColor, _
            exportSettings.RebarCompressionColor, exportSettings.NeutralColor)
    Next i

    DrawCentroidAxesAndLoadPoint ms, section, concrete, steel, loadReferenceX, loadReferenceY, _
        exportSettings.PrincipalAxesEnabled, exportSettings.LoadPointEnabled
    If exportSettings.NeutralLineEnabled Then DrawNeutralLine ms, section, solver, exportSettings.NeutralColor

    doc.Regen 1
End Sub

Private Function StressLabelText(ByVal elementID As String, ByVal stress As Double, _
        ByVal includeElementName As Boolean) As String
    Dim valueText As String
    valueText = Format$(stress, "0.0")
    If includeElementName And Len(Trim$(elementID)) > 0 Then
        StressLabelText = elementID & " " & valueText
    Else
        StressLabelText = valueText
    End If
End Function

Private Sub DrawCentroidAxesAndLoadPoint(ByVal ms As Object, ByVal section As CSectionModel, _
        ByVal concrete As Object, ByVal steel As Object, _
        ByVal loadReferenceX As Double, ByVal loadReferenceY As Double, _
        ByVal principalAxesEnabled As Boolean, ByVal loadPointEnabled As Boolean)
    If Not principalAxesEnabled And Not loadPointEnabled Then Exit Sub

    Dim minX As Double
    Dim maxX As Double
    Dim minY As Double
    Dim maxY As Double
    GetSectionBounds section, minX, maxX, minY, maxY

    Dim axisLength As Double
    axisLength = 0.65 * MaxDouble(maxX - minX, maxY - minY)
    If axisLength <= 0# Then axisLength = 100#

    Dim centroidX As Double
    Dim centroidY As Double
    Dim a As Double
    Dim props As CSectionPropertiesCalculator
    Set props = New CSectionPropertiesCalculator
    props.CalculateTransformed section, concrete, steel
    centroidX = props.CentroidX
    centroidY = props.CentroidY
    a = props.PrincipalAngleRad

    If principalAxesEnabled Then
        AddAcadLine ms, centroidX - axisLength * Cos(a), centroidY - axisLength * Sin(a), _
            centroidX + axisLength * Cos(a), centroidY + axisLength * Sin(a), "RC_NDM_Axes", 3
        AddAcadLine ms, centroidX - axisLength * Cos(a + GEOM_PI / 2#), centroidY - axisLength * Sin(a + GEOM_PI / 2#), _
            centroidX + axisLength * Cos(a + GEOM_PI / 2#), centroidY + axisLength * Sin(a + GEOM_PI / 2#), "RC_NDM_Axes", 3

        AddAcadCircle ms, centroidX, centroidY, MaxDouble(axisLength * 0.018, 5#), "RC_NDM_Axes", 3
    End If
    If loadPointEnabled Then DrawLoadPointMarker ms, loadReferenceX, loadReferenceY, MaxDouble(axisLength * 0.035, 8#)
End Sub

Private Function ReadAutoCADExportSettings(ByVal settings As CSystemSettingsReader) As TAutoCADExportSettings
    With ReadAutoCADExportSettings
        .ConcreteLayer = settings.GetString("AutoCAD.Layer.Concrete", "Concrete")
        .RebarLayer = settings.GetString("AutoCAD.Layer.Rebar", "Reinf")
        .ConcreteTensionLayer = settings.GetString("AutoCAD.Layer.ConcreteTension", "Anno_Concrete_Positive")
        .ConcreteCompressionLayer = settings.GetString("AutoCAD.Layer.ConcreteCompression", "Anno_Concrete_Negative")
        .RebarTensionLayer = settings.GetString("AutoCAD.Layer.RebarTension", "Anno_Rebar_Positive")
        .RebarCompressionLayer = settings.GetString("AutoCAD.Layer.RebarCompression", "Anno_Rebar_Negative")
        .ConcreteTensionColor = settings.GetLong("AutoCAD.Color.ConcreteTension", 9)
        .ConcreteCompressionColor = settings.GetLong("AutoCAD.Color.ConcreteCompression", 5)
        .RebarTensionColor = settings.GetLong("AutoCAD.Color.RebarTension", 1)
        .RebarCompressionColor = settings.GetLong("AutoCAD.Color.RebarCompression", 6)
        .NeutralColor = settings.GetLong("AutoCAD.Color.Neutral", 8)
        .IncludeElementNames = AutoCADLabelModeIncludesNames(settings.GetRawString("AutoCAD.Export.LabelMode", "NamesAndValues"))
        .NeutralLineEnabled = settings.GetBoolean("AutoCAD.Export.NeutralLineEnabled", True)
        .PrincipalAxesEnabled = settings.GetBoolean("AutoCAD.Export.PrincipalAxesEnabled", True)
        .LoadPointEnabled = settings.GetBoolean("AutoCAD.Export.LoadPointEnabled", True)
    End With
End Function

Private Function AutoCADLabelModeIncludesNames(ByVal labelMode As String) As Boolean
    Select Case LCase$(Trim$(labelMode))
        Case "namesandvalues"
            AutoCADLabelModeIncludesNames = True
        Case "valuesonly"
            AutoCADLabelModeIncludesNames = False
        Case Else
            Err.Raise vbObjectError + 4312, "ReadAutoCADExportSettings", _
                "AutoCAD.Export.LabelMode должен быть ValuesOnly или NamesAndValues."
    End Select
End Function

Private Sub GetSectionBounds(ByVal section As CSectionModel, ByRef minX As Double, ByRef maxX As Double, _
        ByRef minY As Double, ByRef maxY As Double)
    If section Is Nothing Then Err.Raise vbObjectError + 4340, "GetSectionBounds", "Модель сечения не передана."
    If section.ConcreteCount <= 0 Then Err.Raise vbObjectError + 4341, "GetSectionBounds", "В модели сечения нет бетонных элементов."

    Dim i As Long
    minX = section.ConcreteX(1) - ConcreteDrawWidth(section, 1) / 2#
    maxX = section.ConcreteX(1) + ConcreteDrawWidth(section, 1) / 2#
    minY = section.ConcreteY(1) - ConcreteDrawHeight(section, 1) / 2#
    maxY = section.ConcreteY(1) + ConcreteDrawHeight(section, 1) / 2#

    For i = 2 To section.ConcreteCount
        minX = MinDouble(minX, section.ConcreteX(i) - ConcreteDrawWidth(section, i) / 2#)
        maxX = MaxDouble(maxX, section.ConcreteX(i) + ConcreteDrawWidth(section, i) / 2#)
        minY = MinDouble(minY, section.ConcreteY(i) - ConcreteDrawHeight(section, i) / 2#)
        maxY = MaxDouble(maxY, section.ConcreteY(i) + ConcreteDrawHeight(section, i) / 2#)
    Next i
End Sub

Private Function ConcreteDrawWidth(ByVal section As CSectionModel, ByVal index As Long) As Double
    ConcreteDrawWidth = section.ConcreteWidth(index)
    If ConcreteDrawWidth <= 0# Then ConcreteDrawWidth = Sqr(section.ConcreteArea(index))
End Function

Private Function ConcreteDrawHeight(ByVal section As CSectionModel, ByVal index As Long) As Double
    ConcreteDrawHeight = section.ConcreteHeight(index)
    If ConcreteDrawHeight <= 0# Then ConcreteDrawHeight = Sqr(section.ConcreteArea(index))
End Function

Private Sub DrawLoadPointMarker(ByVal ms As Object, ByVal x As Double, ByVal y As Double, ByVal size As Double)
    AddAcadLine ms, x - size, y, x + size, y, "RC_NDM_LoadPoint", 2
    AddAcadLine ms, x, y - size, x, y + size, "RC_NDM_LoadPoint", 2
    AddAcadCircle ms, x, y, size * 0.65, "RC_NDM_LoadPoint", 2
End Sub

Private Sub DrawNeutralLine(ByVal ms As Object, ByVal section As CSectionModel, _
        ByVal solver As CSectionSolver, ByVal colorIndex As Long)
    If Abs(solver.KappaX) + Abs(solver.KappaY) <= 0.000000000000001 Then Exit Sub

    Dim minX As Double
    Dim maxX As Double
    Dim minY As Double
    Dim maxY As Double
    GetSectionBounds section, minX, maxX, minY, maxY

    Dim margin As Double
    margin = 0.1 * MaxDouble(maxX - minX, maxY - minY)
    minX = minX - margin
    maxX = maxX + margin
    minY = minY - margin
    maxY = maxY + margin
    ExpandNeutralBoundsByState solver.Epsilon0, solver.KappaX, solver.KappaY, minX, maxX, minY, maxY

    Dim pointX() As Double
    Dim pointY() As Double
    ReDim pointX(1 To 4)
    ReDim pointY(1 To 4)
    Dim pointCount As Long

    AppendNeutralIntersection pointX, pointY, pointCount, minX, _
        NeutralYAtX(solver, minX), minX, maxX, minY, maxY
    AppendNeutralIntersection pointX, pointY, pointCount, maxX, _
        NeutralYAtX(solver, maxX), minX, maxX, minY, maxY
    AppendNeutralIntersection pointX, pointY, pointCount, _
        NeutralXAtY(solver, minY), minY, minX, maxX, minY, maxY
    AppendNeutralIntersection pointX, pointY, pointCount, _
        NeutralXAtY(solver, maxY), maxY, minX, maxX, minY, maxY

    If pointCount < 2 Then Exit Sub

    Dim i As Long
    Dim j As Long
    Dim bestI As Long
    Dim bestJ As Long
    Dim bestDistance2 As Double
    For i = 1 To pointCount - 1
        For j = i + 1 To pointCount
            Dim distance2 As Double
            distance2 = (pointX(i) - pointX(j)) ^ 2 + (pointY(i) - pointY(j)) ^ 2
            If distance2 > bestDistance2 Then
                bestDistance2 = distance2
                bestI = i
                bestJ = j
            End If
        Next j
    Next i

    If bestDistance2 > 0# Then
        AddAcadLine ms, pointX(bestI), pointY(bestI), pointX(bestJ), pointY(bestJ), _
            "RC_NDM_NeutralLine", colorIndex
    End If
End Sub

Private Sub DrawNeutralLineByState(ByVal ms As Object, ByVal section As CSectionModel, _
        ByVal epsilon0 As Double, ByVal kappaX As Double, ByVal kappaY As Double, ByVal colorIndex As Long)
    If Abs(kappaX) + Abs(kappaY) <= 0.000000000000001 Then Exit Sub

    Dim minX As Double
    Dim maxX As Double
    Dim minY As Double
    Dim maxY As Double
    GetSectionBounds section, minX, maxX, minY, maxY

    Dim margin As Double
    margin = 0.1 * MaxDouble(maxX - minX, maxY - minY)
    minX = minX - margin
    maxX = maxX + margin
    minY = minY - margin
    maxY = maxY + margin
    ExpandNeutralBoundsByState epsilon0, kappaX, kappaY, minX, maxX, minY, maxY

    Dim pointX() As Double
    Dim pointY() As Double
    ReDim pointX(1 To 4)
    ReDim pointY(1 To 4)
    Dim pointCount As Long

    AppendNeutralIntersection pointX, pointY, pointCount, minX, _
        NeutralYAtXState(epsilon0, kappaX, kappaY, minX), minX, maxX, minY, maxY
    AppendNeutralIntersection pointX, pointY, pointCount, maxX, _
        NeutralYAtXState(epsilon0, kappaX, kappaY, maxX), minX, maxX, minY, maxY
    AppendNeutralIntersection pointX, pointY, pointCount, _
        NeutralXAtYState(epsilon0, kappaX, kappaY, minY), minY, minX, maxX, minY, maxY
    AppendNeutralIntersection pointX, pointY, pointCount, _
        NeutralXAtYState(epsilon0, kappaX, kappaY, maxY), maxY, minX, maxX, minY, maxY

    If pointCount < 2 Then Exit Sub

    Dim i As Long
    Dim j As Long
    Dim bestI As Long
    Dim bestJ As Long
    Dim bestDistance2 As Double
    For i = 1 To pointCount - 1
        For j = i + 1 To pointCount
            Dim distance2 As Double
            distance2 = (pointX(i) - pointX(j)) ^ 2 + (pointY(i) - pointY(j)) ^ 2
            If distance2 > bestDistance2 Then
                bestDistance2 = distance2
                bestI = i
                bestJ = j
            End If
        Next j
    Next i

    If bestDistance2 > 0# Then
        AddAcadLine ms, pointX(bestI), pointY(bestI), pointX(bestJ), pointY(bestJ), _
            "RC_NDM_NeutralLine", colorIndex
    End If
End Sub

Private Sub ExpandNeutralBoundsByState(ByVal epsilon0 As Double, ByVal kappaX As Double, _
        ByVal kappaY As Double, ByRef minX As Double, ByRef maxX As Double, _
        ByRef minY As Double, ByRef maxY As Double)
    Dim normKappa As Double
    normKappa = Sqr(kappaX * kappaX + kappaY * kappaY)
    If normKappa <= 0.000000000000001 Then Exit Sub

    Dim centerX As Double
    Dim centerY As Double
    centerX = 0.5 * (minX + maxX)
    centerY = 0.5 * (minY + maxY)

    Dim sectionSize As Double
    sectionSize = MaxDouble(maxX - minX, maxY - minY)
    If sectionSize <= 0# Then sectionSize = 100#

    Dim halfSize As Double
    halfSize = 0.5 * sectionSize

    Dim distanceToNeutral As Double
    distanceToNeutral = Abs(epsilon0 + kappaX * centerY + kappaY * centerX) / normKappa

    If distanceToNeutral <= halfSize Then Exit Sub

    Dim expandedHalfSize As Double
    expandedHalfSize = MinDouble(distanceToNeutral + 0.5 * sectionSize, 20# * sectionSize)
    If expandedHalfSize <= halfSize Then Exit Sub

    minX = centerX - expandedHalfSize
    maxX = centerX + expandedHalfSize
    minY = centerY - expandedHalfSize
    maxY = centerY + expandedHalfSize
End Sub

Private Function NeutralYAtX(ByVal solver As CSectionSolver, ByVal x As Double) As Double
    If Abs(solver.KappaX) <= 0.000000000000001 Then
        NeutralYAtX = 1E+99
    Else
        NeutralYAtX = -(solver.Epsilon0 + solver.KappaY * x) / solver.KappaX
    End If
End Function

Private Function NeutralXAtY(ByVal solver As CSectionSolver, ByVal y As Double) As Double
    If Abs(solver.KappaY) <= 0.000000000000001 Then
        NeutralXAtY = 1E+99
    Else
        NeutralXAtY = -(solver.Epsilon0 + solver.KappaX * y) / solver.KappaY
    End If
End Function

Private Function NeutralYAtXState(ByVal epsilon0 As Double, ByVal kappaX As Double, _
        ByVal kappaY As Double, ByVal x As Double) As Double
    If Abs(kappaX) <= 0.000000000000001 Then
        NeutralYAtXState = 1E+99
    Else
        NeutralYAtXState = -(epsilon0 + kappaY * x) / kappaX
    End If
End Function

Private Function NeutralXAtYState(ByVal epsilon0 As Double, ByVal kappaX As Double, _
        ByVal kappaY As Double, ByVal y As Double) As Double
    If Abs(kappaY) <= 0.000000000000001 Then
        NeutralXAtYState = 1E+99
    Else
        NeutralXAtYState = -(epsilon0 + kappaX * y) / kappaY
    End If
End Function

Private Sub AppendNeutralIntersection(ByRef pointX() As Double, ByRef pointY() As Double, _
        ByRef pointCount As Long, ByVal x As Double, ByVal y As Double, _
        ByVal minX As Double, ByVal maxX As Double, ByVal minY As Double, ByVal maxY As Double)
    If x < minX - 0.0000001 Or x > maxX + 0.0000001 Then Exit Sub
    If y < minY - 0.0000001 Or y > maxY + 0.0000001 Then Exit Sub

    Dim i As Long
    For i = 1 To pointCount
        If Abs(pointX(i) - x) < 0.0000001 And Abs(pointY(i) - y) < 0.0000001 Then Exit Sub
    Next i

    If pointCount >= UBound(pointX) Then Exit Sub
    pointCount = pointCount + 1
    pointX(pointCount) = x
    pointY(pointCount) = y
End Sub

Private Sub AddAcadRectangleRegion(ByVal ms As Object, ByVal x As Double, ByVal y As Double, _
        ByVal width As Double, ByVal height As Double, ByVal layerName As String, ByVal colorIndex As Long)
    Dim p(0 To 9) As Double
    p(0) = x - width / 2#: p(1) = y - height / 2#
    p(2) = x + width / 2#: p(3) = y - height / 2#
    p(4) = x + width / 2#: p(5) = y + height / 2#
    p(6) = x - width / 2#: p(7) = y + height / 2#
    p(8) = x - width / 2#: p(9) = y - height / 2#

    Dim source As Object
    Set source = ms.AddLightWeightPolyline(p)
    source.Closed = True
    AddAcadRegionFromCurve ms, source, layerName, colorIndex
End Sub

Private Sub AddAcadLine(ByVal ms As Object, ByVal x1 As Double, ByVal y1 As Double, _
        ByVal x2 As Double, ByVal y2 As Double, ByVal layerName As String, ByVal colorIndex As Long)
    Dim p1(0 To 2) As Double
    Dim p2(0 To 2) As Double
    p1(0) = x1: p1(1) = y1: p1(2) = 0#
    p2(0) = x2: p2(1) = y2: p2(2) = 0#
    Dim entity As Object
    Set entity = ms.AddLine(p1, p2)
    entity.Layer = layerName
    entity.Color = colorIndex
End Sub

Private Sub AddAcadCircleRegion(ByVal ms As Object, ByVal x As Double, ByVal y As Double, _
        ByVal radius As Double, ByVal layerName As String, ByVal colorIndex As Long)
    Dim p(0 To 2) As Double
    p(0) = x: p(1) = y: p(2) = 0#
    Dim source As Object
    Set source = ms.AddCircle(p, radius)
    AddAcadRegionFromCurve ms, source, layerName, colorIndex
End Sub

Private Sub AddAcadRegionFromCurve(ByVal ms As Object, ByVal source As Object, _
        ByVal layerName As String, ByVal colorIndex As Long)
    Dim sourceObjects(0 To 0) As Object
    Set sourceObjects(0) = source

    Dim regions As Variant
    regions = ms.AddRegion(sourceObjects)

    Dim entity As Object
    Set entity = regions(LBound(regions))
    entity.Layer = layerName
    entity.Color = colorIndex
    source.Delete
End Sub

Private Sub EnsureAcadLayer(ByVal doc As Object, ByVal layerName As String, ByVal colorIndex As Long)
    On Error Resume Next
    Dim layer As Object
    Set layer = doc.Layers.Item(layerName)
    If layer Is Nothing Then Set layer = doc.Layers.Add(layerName)
    layer.Color = colorIndex
    On Error GoTo 0
End Sub

Private Sub AddAcadCircle(ByVal ms As Object, ByVal x As Double, ByVal y As Double, _
        ByVal radius As Double, ByVal layerName As String, ByVal colorIndex As Long)
    Dim p(0 To 2) As Double
    p(0) = x: p(1) = y: p(2) = 0#
    Dim entity As Object
    Set entity = ms.AddCircle(p, radius)
    entity.Layer = layerName
    entity.Color = colorIndex
End Sub

Private Sub AddAcadText(ByVal ms As Object, ByVal value As String, ByVal x As Double, ByVal y As Double, _
        ByVal height As Double, ByVal layerName As String, ByVal colorIndex As Long)
    Dim p(0 To 2) As Double
    p(0) = x: p(1) = y: p(2) = 0#
    Dim entity As Object
    Set entity = ms.AddText(value, p, height)
    entity.Layer = layerName
    entity.Color = colorIndex
End Sub

Private Function StressColor(ByVal stress As Double, ByVal tensionColor As Long, _
        ByVal compressionColor As Long, ByVal neutralColor As Long) As Long
    If stress > 0.000000001 Then
        StressColor = tensionColor
    ElseIf stress < -0.000000001 Then
        StressColor = compressionColor
    Else
        StressColor = neutralColor
    End If
End Function

Private Function StressAnnotationLayer(ByVal stress As Double, ByVal tensionLayer As String, _
        ByVal compressionLayer As String) As String
    If stress > 0.000000001 Then
        StressAnnotationLayer = tensionLayer
    Else
        StressAnnotationLayer = compressionLayer
    End If
End Function

Private Sub ReadFirstExportLoad(ByVal workbook As Object, ByRef nValue As Double, ByRef mxValue As Double, ByRef myValue As Double)
    Dim source As Object
    Set source = workbook.Names.Item("rngLoadCombinations").RefersToRange
    Dim data As Variant
    data = source.Value2

    Dim rowIndex As Long
    For rowIndex = 2 To UBound(data, 1)
        If Not IsExportEmptyRow(data, rowIndex, 7) Then
            nValue = ReadExportRequiredDouble(data(rowIndex, 2), "N")
            mxValue = ReadExportRequiredDouble(data(rowIndex, 3), "Mx")
            myValue = ReadExportRequiredDouble(data(rowIndex, 4), "My")
            Exit Sub
        End If
    Next rowIndex
    Err.Raise vbObjectError + 4320, "ReadFirstExportLoad", "Не задано ни одного сочетания нагрузок."
End Sub

Private Function IsExportEmptyRow(ByRef data As Variant, ByVal rowIndex As Long, ByVal columnCount As Long) As Boolean
    Dim col As Long
    For col = 1 To columnCount
        If Len(Trim$(CStr(data(rowIndex, col)))) > 0 Then Exit Function
    Next col
    IsExportEmptyRow = True
End Function

Private Function ReadExportRequiredDouble(ByVal value As Variant, ByVal fieldName As String) As Double
    If Len(Trim$(CStr(value))) = 0 Then Err.Raise vbObjectError + 4330, "ReadExportRequiredDouble", fieldName & " не заполнен."
    If Not IsNumeric(value) Then Err.Raise vbObjectError + 4331, "ReadExportRequiredDouble", fieldName & " должен быть числом."
    ReadExportRequiredDouble = CDbl(value)
End Function

Private Function MaxDouble(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then MaxDouble = a Else MaxDouble = b
End Function

Private Function MinDouble(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then MinDouble = a Else MinDouble = b
End Function

