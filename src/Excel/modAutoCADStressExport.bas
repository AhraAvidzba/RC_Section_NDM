Attribute VB_Name = "modAutoCADStressExport"
Option Explicit

Public Sub ExportSectionStressToAutoCAD()
    On Error GoTo Failed

    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook

    Dim geometry As ISectionGeometry
    Set geometry = ReadExportGeometry(ThisWorkbook, settings)

    Dim stepX As Double
    Dim stepY As Double
    stepX = settings.GetDouble("Mesh.StepX", 10#)
    stepY = settings.GetDouble("Mesh.StepY", 10#)

    Dim mesh As CFiberMeshBuilder
    Set mesh = New CFiberMeshBuilder
    mesh.BuildMesh geometry, stepX, stepY, 1, ExportMeshBoundarySubdivisions(settings)

    Dim rebars As CRebarLayout
    Set rebars = ReadExportRebars(ThisWorkbook, geometry, settings)

    Dim concrete As CConcreteDiagramMaterial
    Set concrete = New CConcreteDiagramMaterial
    concrete.InitializeFromSettings settings

    Dim steel As CSteelDiagramMaterial
    Set steel = New CSteelDiagramMaterial
    steel.InitializeFromSettings settings


    Dim nValue As Double
    Dim mxValue As Double
    Dim myValue As Double
    ReadFirstExportLoad ThisWorkbook, nValue, mxValue, myValue

    Dim solver As CSectionSolver
    Set solver = New CSectionSolver
    solver.ApplySettings settings
    solver.Solve mesh, rebars, concrete, steel, nValue, mxValue, myValue
    If Not solver.Converged Then Err.Raise vbObjectError + 4300, "ExportSectionStressToAutoCAD", solver.StopReason

    DrawStressExport mesh, rebars, concrete, steel, solver
    MsgBox "Выгрузка в AutoCAD завершена. Волокон: " & CStr(mesh.FiberCount) & _
        "; стержней: " & CStr(rebars.Count), vbInformation, "RC Section NDM"
    Exit Sub

Failed:
    MsgBox "Выгрузка в AutoCAD не выполнена: " & Err.Description, vbExclamation, "RC Section NDM"
End Sub

Private Sub DrawStressExport(ByVal mesh As CFiberMeshBuilder, ByVal rebars As CRebarLayout, _
        ByVal concrete As Object, ByVal steel As Object, ByVal solver As CSectionSolver)
    Dim acad As Object
    Set acad = GetObject(, "AutoCAD.Application")
    If acad Is Nothing Then Err.Raise vbObjectError + 4310, "DrawStressExport", "Откройте AutoCAD и активный чертеж."

    Dim doc As Object
    Set doc = acad.ActiveDocument
    If doc Is Nothing Then Err.Raise vbObjectError + 4311, "DrawStressExport", "В AutoCAD нет активного чертежа."

    Dim ms As Object
    Set ms = doc.ModelSpace

    EnsureAcadLayer doc, "RC_NDM_Fibers", 8
    EnsureAcadLayer doc, "RC_NDM_Rebar", 1
    EnsureAcadLayer doc, "RC_NDM_StressText", 7

    Dim i As Long
    Dim strain As Double
    Dim stress As Double
    Dim textHeight As Double

    For i = 1 To mesh.FiberCount
        strain = solver.Epsilon0 + solver.KappaX * mesh.FiberY(i) + solver.KappaY * mesh.FiberX(i)
        stress = concrete.GetStress(strain)
        textHeight = 0.22 * MinDouble(mesh.FiberWidth(i), mesh.FiberHeight(i))
        If textHeight <= 0# Then textHeight = 1#
        AddAcadRectangle ms, mesh.FiberX(i), mesh.FiberY(i), mesh.FiberWidth(i), mesh.FiberHeight(i), "RC_NDM_Fibers", StressColor(stress)
        AddAcadText ms, Format$(stress, "0.0"), mesh.FiberX(i) - 0.45 * mesh.FiberWidth(i), _
            mesh.FiberY(i) - 0.1 * mesh.FiberHeight(i), textHeight, "RC_NDM_StressText", StressColor(stress)
    Next i

    For i = 1 To rebars.Count
        strain = solver.Epsilon0 + solver.KappaX * rebars.Y(i) + solver.KappaY * rebars.X(i)
        stress = steel.GetStress(strain)
        AddAcadCircle ms, rebars.X(i), rebars.Y(i), rebars.Diameter(i) / 2#, "RC_NDM_Rebar", StressColor(stress)
        AddAcadText ms, rebars.BarID(i) & " " & Format$(stress, "0.0"), _
            rebars.X(i) + rebars.Diameter(i) / 2#, rebars.Y(i) + rebars.Diameter(i) / 2#, _
            MaxDouble(2.5, rebars.Diameter(i) * 0.18), "RC_NDM_StressText", StressColor(stress)
    Next i

    doc.Regen 1
End Sub

Private Sub AddAcadRectangle(ByVal ms As Object, ByVal x As Double, ByVal y As Double, _
        ByVal width As Double, ByVal height As Double, ByVal layerName As String, ByVal colorIndex As Long)
    Dim p(0 To 9) As Double
    p(0) = x - width / 2#: p(1) = y - height / 2#
    p(2) = x + width / 2#: p(3) = y - height / 2#
    p(4) = x + width / 2#: p(5) = y + height / 2#
    p(6) = x - width / 2#: p(7) = y + height / 2#
    p(8) = x - width / 2#: p(9) = y - height / 2#

    Dim entity As Object
    Set entity = ms.AddLightWeightPolyline(p)
    entity.Closed = True
    entity.Layer = layerName
    entity.Color = colorIndex
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

Private Function StressColor(ByVal stress As Double) As Long
    If stress < -0.000000001 Then
        StressColor = 5
    ElseIf stress > 0.000000001 Then
        StressColor = 1
    Else
        StressColor = 8
    End If
End Function

Private Function ExportMeshBoundarySubdivisions(ByVal settings As CSystemSettingsReader) As Long
    ExportMeshBoundarySubdivisions = settings.GetLong("Mesh.BoundarySubdivisions", 1)
    If ExportMeshBoundarySubdivisions < 1 Then ExportMeshBoundarySubdivisions = 1
End Function

Private Function ReadExportGeometry(ByVal workbook As Object, ByVal settings As CSystemSettingsReader) As ISectionGeometry
    Dim inputRange As Object
    Set inputRange = workbook.Names.Item("rngMainInput").RefersToRange

    Dim geometryType As String
    geometryType = settings.GetString("Geometry.Type", "RoundedRectangle")
    If Len(geometryType) = 0 Then geometryType = "RoundedRectangle"

    If StrComp(geometryType, "Circle", vbTextCompare) = 0 Then
        Dim circleGeom As CGeometryCircle
        Set circleGeom = New CGeometryCircle
        circleGeom.InitializeByDiameter settings.GetDouble("Circle.Diameter", 300#), _
            settings.GetDouble("Circle.CenterX", 0#), _
            settings.GetDouble("Circle.CenterY", 0#)
        Set ReadExportGeometry = circleGeom
    Else
        Dim rect As CGeometryRoundedRectangle
        Set rect = New CGeometryRoundedRectangle
        rect.Initialize ReadExportRequiredDouble(inputRange.Cells.Item(3, 5).Value2, "Width"), _
            ReadExportRequiredDouble(inputRange.Cells.Item(4, 5).Value2, "Height"), _
            ReadExportOptionalDouble(inputRange.Cells.Item(5, 5).Value2, 0#), _
            ReadExportOptionalDouble(inputRange.Cells.Item(6, 5).Value2, 0#), _
            ReadExportOptionalDouble(inputRange.Cells.Item(7, 5).Value2, 0#), _
            ReadExportOptionalDouble(inputRange.Cells.Item(8, 5).Value2, 0#)
        Set ReadExportGeometry = rect
    End If
End Function

Private Function ReadExportRebars(ByVal workbook As Object, ByVal geometry As ISectionGeometry, ByVal settings As CSystemSettingsReader) As CRebarLayout
    Dim inputRange As Object
    Set inputRange = workbook.Names.Item("rngMainInput").RefersToRange

    Dim geometryType As String
    geometryType = settings.GetString("Geometry.Type", "RoundedRectangle")
    If StrComp(geometryType, "Circle", vbTextCompare) <> 0 Then
        Err.Raise vbObjectError + 4321, "ReadExportRebars", "Автоматическая расстановка арматуры сейчас поддерживает только Circle."
    End If

    Dim builder As CCircleRebarLayoutBuilder
    Set builder = New CCircleRebarLayoutBuilder
    Set ReadExportRebars = builder.Build( _
        settings.GetDouble("Circle.Diameter", 300#), _
        settings.GetDouble("Circle.CenterX", 0#), _
        settings.GetDouble("Circle.CenterY", 0#), _
        ReadExportRequiredDouble(inputRange.Cells.Item(14, 5).Value2, "Rebar.AxisDistance"), _
        CLng(ReadExportRequiredDouble(inputRange.Cells.Item(15, 5).Value2, "Rebar.Count")), _
        ReadExportRequiredDouble(inputRange.Cells.Item(16, 5).Value2, "Rebar.Diameter"), _
        Trim$(CStr(inputRange.Cells.Item(13, 5).Value2)))
End Function

Private Sub ReadFirstExportLoad(ByVal workbook As Object, ByRef nValue As Double, ByRef mxValue As Double, ByRef myValue As Double)
    Dim source As Object
    Set source = workbook.Names.Item("rngLoadCombinations").RefersToRange

    Dim rowIndex As Long
    For rowIndex = 2 To source.Rows.Count
        If Not IsExportEmptyRow(source, rowIndex, 7) Then
            nValue = ReadExportRequiredDouble(source.Cells.Item(rowIndex, 2).Value2, "N")
            mxValue = ReadExportRequiredDouble(source.Cells.Item(rowIndex, 3).Value2, "Mx")
            myValue = ReadExportRequiredDouble(source.Cells.Item(rowIndex, 4).Value2, "My")
            Exit Sub
        End If
    Next rowIndex
    Err.Raise vbObjectError + 4320, "ReadFirstExportLoad", "Не задано ни одного сочетания нагрузок."
End Sub

Private Function IsExportEmptyRow(ByVal source As Object, ByVal rowIndex As Long, ByVal columnCount As Long) As Boolean
    Dim col As Long
    For col = 1 To columnCount
        If Len(Trim$(CStr(source.Cells.Item(rowIndex, col).Value2))) > 0 Then Exit Function
    Next col
    IsExportEmptyRow = True
End Function

Private Function ReadExportRequiredDouble(ByVal value As Variant, ByVal fieldName As String) As Double
    If Len(Trim$(CStr(value))) = 0 Then Err.Raise vbObjectError + 4330, "ReadExportRequiredDouble", fieldName & " is empty."
    If Not IsNumeric(value) Then Err.Raise vbObjectError + 4331, "ReadExportRequiredDouble", fieldName & " is not numeric."
    ReadExportRequiredDouble = CDbl(value)
End Function

Private Function ReadExportOptionalDouble(ByVal value As Variant, ByVal defaultValue As Double) As Double
    If Len(Trim$(CStr(value))) = 0 Then
        ReadExportOptionalDouble = defaultValue
    ElseIf IsNumeric(value) Then
        ReadExportOptionalDouble = CDbl(value)
    Else
        Err.Raise vbObjectError + 4332, "ReadExportOptionalDouble", "Value is not numeric."
    End If
End Function

Private Function MaxDouble(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then MaxDouble = a Else MaxDouble = b
End Function

Private Function MinDouble(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then MinDouble = a Else MinDouble = b
End Function



