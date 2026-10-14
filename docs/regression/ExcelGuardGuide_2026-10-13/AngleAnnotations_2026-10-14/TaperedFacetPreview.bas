' ДЛЯ ТЕСТОВ: оригинальная модель и отдельная количественная подпись двух её стержней.
Public GoalFacetDefinition As CSectionDefinition
Public Function GoalTaperedFacetPreview(ByVal outputDirectory As String) As String
    Dim model As CSectionModel, builder As CSectionBuilder, subset As CRebarLayout
    Dim annotations As CSectionAnnotations, annotationBuilder As CSectionAnnotationBuilder, i As Long
    Dim before As String, after As String
    Set builder = New CSectionBuilder: Set model = FixtureModel("RoundedTapered", builder)
    before = ShapeSnapshot("RoundedTapered", model, builder.Mesh, builder.Rebars)
    Set subset = New CRebarLayout
    For i = 1 To builder.Rebars.Count
        If builder.Rebars.BarAnnotationGroupName(i) = "H.Right" And builder.Rebars.Y(i) > 0# Then
            subset.AddBar builder.Rebars.BarID(i), builder.Rebars.X(i), builder.Rebars.Y(i), builder.Rebars.Diameter(i), _
                builder.Rebars.Area(i), builder.Rebars.SteelClass(i), annotationGroupName:="UpperRightFacet"
        End If
    Next i
    If subset.Count <> 2 Then Err.Raise 5, "GoalTaperedFacetPreview", "Expected two existing bars on the upper facet."
    Set annotations = New CSectionAnnotations: Set annotationBuilder = New CSectionAnnotationBuilder
    annotationBuilder.Build GoalFacetDefinition, subset, annotations, builder.Rules
    For i = 1 To annotations.Count
        If annotations.AnnotationID(i) = "REBAR_UpperRightFacet" Then
            model.Annotations.AddRebarLabel annotations.AnnotationID(i), annotations.StartX(i), annotations.StartY(i), _
                annotations.EndX(i), annotations.EndY(i), annotations.NormalX(i), annotations.NormalY(i), annotations.Text(i)
        End If
    Next i
    after = ShapeSnapshot("RoundedTapered", model, builder.Mesh, builder.Rebars)
    Dim stream As Object, fs As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set stream = fs.CreateTextFile(outputDirectory & Application.PathSeparator & "BeforeExtraLabel.tsv", True, True): stream.Write before: stream.Close
    Set stream = fs.CreateTextFile(outputDirectory & Application.PathSeparator & "AfterExtraLabel.tsv", True, True): stream.Write after: stream.Close
    GoalTaperedFacetPreview = PreviewSnapshot("RoundedTapered", outputDirectory, model)
End Function
