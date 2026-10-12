Attribute VB_Name = "modTestRecipeAnnotations"
Option Explicit

' ДЛЯ ТЕСТОВ: измерения общего API и сохраненные подписи без повторной сборки.
Private mPassed As Long, mFailed As Long, mReport As String

Public Function RunRecipeAnnotationTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    On Error GoTo Failed
    Dim facade As CSectionDefinitionBuilder, definition As CSectionDefinition, builder As CSectionBuilder
    Dim model As CSectionModel, annotations As CSectionAnnotations, x As Double, y As Double, i As Long
    Set facade = AnnotationFacade
    facade.Source "Diagonal", "ProgramInputs.PointB"
    Set definition = facade.Finish
    Check "source.scalar", definition.SourceReference("Diagonal") = "ProgramInputs.PointB"
    definition.AnchorCoordinates "B", x, y: x = 999#: y = 999#
    definition.AnchorCoordinates "B", x, y
    Near "anchor.immutable.x", x, 30#: Near "anchor.immutable.y", y, 40#
    Set builder = New CSectionBuilder: Set model = BuildDefinition(definition, builder): Set annotations = model.Annotations
    Check "annotations.count", annotations.Count = 11
    For i = 1 To 9: Check "annotation.id." & CStr(i), Len(annotations.AnnotationID(i)) > 0: Next i
    Near "horizontal.projection", annotations.Value(1), 30#
    Near "horizontal.placement", annotations.EndY(1), 0#
    Near "vertical.projection", annotations.Value(2), 40#
    Near "vertical.placement", annotations.EndX(2), 0#
    Near "diagonal.distance", annotations.Value(3), 50#
    Near "diagonal.normal.x", annotations.NormalX(3), -0.8
    Near "diagonal.normal.y", annotations.NormalY(3), 0.6
    Near "vertical.explicit", annotations.Value(4), 100#
    Near "extent.full.height", annotations.Value(5), 400#
    Near "extent.full.width", annotations.Value(6), 400#
    Near "radius", annotations.Value(7), 200#
    Near "diameter", annotations.Value(8), 400#
    Near "radius.endpoint.x", annotations.EndX(7), 200#
    Near "radius.endpoint.y", annotations.EndY(7), 0#
    Near "diameter.opposite.x", annotations.StartX(8), -200#
    Near "length.analytic", annotations.Value(9), 400# * GEOM_PI
    Check "length.distinct.extent", annotations.Value(9) > annotations.Value(5)
    Check "text.unicode", annotations.Text(10) = "Проверка привязки"
    Check "text.kind", annotations.AnnotationType(10) = "TEXT_ANNOTATION"
    Check "length.kind.unit", annotations.AnnotationType(9) = "MEASUREMENT_LABEL" And annotations.Unit(9) = "Length"
    Check "physical.labels", InStr(annotations.Text(11), "2") > 0 And model.RebarCount = 2
    For i = 0 To 4: Check "invalid." & CStr(i), InvalidAnnotation(i): Next i
    GoTo Finished
Failed:
    Check "annotations.runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    RunRecipeAnnotationTests = mReport & "TOTAL RECIPE ANNOTATIONS: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

Public Function RunRecipeAnnotationSnapshotTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    Dim config As Object, saved As Variant, events As Boolean, updating As Boolean, calculation As Long
    On Error GoTo Failed
    events = Application.EnableEvents: updating = Application.ScreenUpdating: calculation = Application.Calculation
    Application.EnableEvents = False: Application.ScreenUpdating = False
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = config.Formula
    Dim facade As CSectionDefinitionBuilder, builder As CSectionBuilder, definition As CSectionDefinition, model As CSectionModel
    Dim properties As CSectionPropertiesCalculator, units As CUnitSystem, settings As CSystemSettingsReader
    Dim writer As CNDMResultsWriter, reader As CSectionPlotDataReader, outputUnit As Variant, i As Long, found As Long
    Set facade = AnnotationFacade: Set definition = facade.Finish
    Set builder = New CSectionBuilder: Set model = BuildDefinition(definition, builder)
    Set properties = New CSectionPropertiesCalculator: properties.PrepareSnapshot model
    For Each outputUnit In Array("mm", "cm", "m")
        SetUnitSetting "Units.Length.Output", CStr(outputUnit)
        Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
        Set units = New CUnitSystem: units.LoadFromSettings settings
        Set writer = New CNDMResultsWriter: writer.WriteGeometryPreview ThisWorkbook, model, properties, units
        Dim geometryTable As Variant, propertyTable As Variant, row As Long, foundDiagnostics As Boolean
        geometryTable = ThisWorkbook.Names("rngNDMSectionGeometry").RefersToRange.Resize(1 + model.ConcreteCount + model.RebarCount, 15).Value2
        Check "snapshot.provenance." & CStr(outputUnit), InStr(CStr(geometryTable(UBound(geometryTable, 1), 15)), "RecipeTrace:RuleId=TopBars|PathId=Row|SourcePositionId=TopBars.2") > 0
        propertyTable = ThisWorkbook.Names("rngNDMSectionProperties").RefersToRange.Resize(46, 6).Value2: foundDiagnostics = False
        For row = 2 To UBound(propertyTable, 1)
            If CStr(propertyTable(row, 3)) = "Geometry.RecipeDiagnostics" Then
                foundDiagnostics = True
                Check "snapshot.diagnostics.counts." & CStr(outputUnit), InStr(CStr(propertyTable(row, 4)), "Requested=2; Created=2") > 0
            End If
        Next row
        Check "snapshot.diagnostics.present." & CStr(outputUnit), foundDiagnostics
        ' Изменение input и output после сохранения не меняет измерение Results.
        SetUnitSetting "Units.Length.Input", "m"
        SetUnitSetting "Units.Length.Output", "mm"
        settings.LoadFromWorkbook ThisWorkbook
        Set reader = New CSectionPlotDataReader: reader.LoadGeometryPreviewFromWorkbook ThisWorkbook, settings
        found = 0
        For i = 1 To reader.AnnotationCount
            If reader.AnnotationType(i) = "MEASUREMENT_LABEL" Then
                found = found + 1
                Near "snapshot.length.anchor.x." & CStr(outputUnit), reader.AnnotationStartX(i), 0#
                Near "snapshot.length.anchor.y." & CStr(outputUnit), reader.AnnotationStartY(i), -100#
                Check "snapshot.length.text." & CStr(outputUnit) & "; " & reader.AnnotationText(i), InStr(reader.AnnotationText(i), "L=") > 0 And InStr(reader.AnnotationText(i), CStr(outputUnit)) > 0
            ElseIf reader.AnnotationType(i) = "TEXT_ANNOTATION" Then
                found = found + 1
                Check "snapshot.text." & CStr(outputUnit), reader.AnnotationText(i) = "Проверка привязки"
            End If
        Next i
        Check "snapshot.labels." & CStr(outputUnit), found = 2
    Next outputUnit
    GoTo Finished
Failed:
    Check "snapshot.runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    On Error Resume Next
    If Not config Is Nothing And Not IsEmpty(saved) Then config.Formula = saved
    Application.EnableEvents = events: Application.ScreenUpdating = updating: Application.Calculation = calculation
    On Error GoTo 0
    RunRecipeAnnotationSnapshotTests = mReport & "TOTAL RECIPE ANNOTATION SNAPSHOT: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

Private Sub SetUnitSetting(ByVal key As String, ByVal value As String)
    Dim range As Object, table As CSectionInputTable
    Set range = ThisWorkbook.Names("rngUnitSettings").RefersToRange
    Set table = New CSectionInputTable: table.Initialize "rngUnitSettings", range.Value2
    If key = "Units.Length.Output" Then
        range.Cells(table.RowByKey("Length"), 4).Value2 = value
    Else
        range.Cells(table.RowByKey("Length"), 2).Value2 = value
    End If
End Sub

Public Function DumpRecipeAnnotationPreview(ByVal outputDirectory As String) As String
    Dim facade As CSectionDefinitionBuilder, builder As CSectionBuilder, model As CSectionModel, definition As CSectionDefinition
    Dim range As Object, saved As Variant, table As CSectionInputTable, row As Long, result As String, number As Long, reason As String
    On Error GoTo Failed
    Set range = ThisWorkbook.Names("rngPlotAnnotationSettings").RefersToRange: saved = range.Formula
    Set table = New CSectionInputTable: table.Initialize "rngPlotAnnotationSettings", range.Value2
    row = table.RowByKey("Enabled"): range.Cells(row, 2).Value2 = "Yes": range.Cells(row, 3).Value2 = "Yes"
    Set facade = AnnotationFacade(True): Set definition = facade.Finish: Set builder = New CSectionBuilder
    Set model = BuildDefinition(definition, builder)
    result = modTestShapeArchitecture.PreparedPreviewSnapshot("RecipeAnnotations", model, outputDirectory)
    GoTo Finished
Failed:
    number = Err.Number: reason = Err.Description
Finished:
    On Error Resume Next
    If Not range Is Nothing And Not IsEmpty(saved) Then range.Formula = saved
    On Error GoTo 0
    If number <> 0 Then Err.Raise number, "modTestRecipeAnnotations", reason
    DumpRecipeAnnotationPreview = result
End Function

Private Function AnnotationFacade(Optional ByVal previewPlacement As Boolean = False) As CSectionDefinitionBuilder
    Dim facade As CSectionDefinitionBuilder, path As CGeometryPath, options As CRebarRuleOptions
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "AnnotationFixture"
    Set path = facade.CirclePath("Boundary", 0#, 0#, 200#): facade.Outer "Concrete", path.ID
    Set path = facade.LinePath("Row", -50#, 100#, 50#, 100#)
    Set options = New CRebarRuleOptions: options.AnnotationGroup = "Top"
    facade.RebarByCount "TopBars", path.ID, 2, 10#, options
    If previewPlacement Then
        facade.Anchor "A", -100#, -75#: facade.Anchor "B", -70#, -35#
        facade.Anchor "LengthNote", -60#, -150#: facade.Anchor "TextNote", 50#, -50#
    Else
        facade.Anchor "A", 0#, 0#: facade.Anchor "B", 30#, 40#: facade.Anchor "LengthNote", 0#, -100#
        facade.Anchor "TextNote", 200#, 0#
    End If
    facade.PathAnchor "East", "Boundary", "Boundary", 0#
    facade.Dimension "Horizontal", "A", "B", "dx", "Horizontal"
    facade.Dimension "Vertical", "A", "B", "dy", "Vertical"
    facade.Dimension "Diagonal", "A", "B", "s"
    If previewPlacement Then facade.VerticalDistance "VerticalExplicit", "v", -100#, 0#, 150#, 1# Else facade.VerticalDistance "VerticalExplicit", "v", -100#, 0#, -50#
    facade.VerticalExtent "Height", "Boundary", "H", -200#
    If previewPlacement Then
        facade.HorizontalExtent "Width", "Boundary", "B", -200#, -1#
        facade.Radius "Radius", "Boundary", "Boundary", 0.75
        facade.Diameter "Diameter", "Boundary", "Boundary", 0.25
    Else
        facade.HorizontalExtent "Width", "Boundary", "B", -200#
        facade.Radius "Radius", "Boundary", "Boundary", 0#
        facade.Diameter "Diameter", "Boundary", "Boundary", 0#
    End If
    facade.PathLength "Length", "Boundary", "LengthNote"
    facade.Text "Note", "TextNote", "Проверка привязки"
    facade.RebarRowLabel "Top", "TopBars", 0#, 1#
    Set AnnotationFacade = facade
End Function

Private Function BuildDefinition(ByVal definition As CSectionDefinition, ByVal builder As CSectionBuilder) As CSectionModel
    Dim mesh As TSectionMeshOptions
    mesh.StepX = 100#: mesh.StepY = 100#: mesh.MaterialID = 1: mesh.BoundarySubdivisions = 2
    Set BuildDefinition = builder.Build(definition, mesh)
End Function

Private Function InvalidAnnotation(ByVal scenario As Long) As Boolean
    On Error GoTo Rejected
    Dim facade As CSectionDefinitionBuilder, definition As CSectionDefinition, builder As CSectionBuilder, model As CSectionModel
    Set facade = AnnotationFacade
    Select Case scenario
        Case 0: facade.Dimension "Missing", "Unknown", "B", "s"
        Case 1: facade.Dimension "UnknownMode", "A", "B", "s", "Unknown"
        Case 2: facade.Radius "LineRadius", "Row", "Row"
        Case 3: facade.Diameter "InvalidParameter", "Boundary", "Boundary", 2#
        Case 4: facade.PathAnchor "MissingSegment", "Boundary", "RemovedArc"
    End Select
    Set definition = facade.Finish: Set builder = New CSectionBuilder: Set model = BuildDefinition(definition, builder)
    Exit Function
Rejected:
    InvalidAnnotation = (Err.Number = vbObjectError + 5530 Or Err.Number = vbObjectError + 5520): Err.Clear
End Function

Private Sub Check(ByVal name As String, ByVal value As Boolean)
    If value Then mPassed = mPassed + 1 Else mFailed = mFailed + 1: mReport = mReport & "FAIL: " & name & vbCrLf
End Sub

Private Sub Near(ByVal name As String, ByVal actual As Double, ByVal expected As Double)
    Check name & "; actual=" & CStr(actual) & "; expected=" & CStr(expected), Abs(actual - expected) <= 0.000000001 * (1# + Abs(expected))
End Sub
