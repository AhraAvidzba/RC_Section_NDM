' ДЛЯ ТЕСТОВ: добавляет декларации наклонной грани только в тестовую копию рецепта.
Private Sub GoalAppendSlopeAnnotations(ByVal facade As CSectionDefinitionBuilder, ByRef sizes As TRoundedInput)
    Dim facetAxis As CGeometryPath, normalX As Double, normalY As Double, length As Double
    length = Sqr((sizes.H / 2#) ^ 2 + sizes.Widths(2) ^ 2)
    normalX = sizes.H / 2# / length: normalY = sizes.Widths(2) / length
    Set facetAxis = facade.Chain("UpperRightFacetAxis", "Offset.H.Right", "Edge.TR", 0#, "Edge.TR", 1#)
    facade.RebarLabel "UpperRightFacet", facetAxis.ID, normalX, normalY
    facade.PathAnchor "UpperRightFacet.Start", "Boundary", "Edge.TR", 0#
    facade.PathAnchor "UpperRightFacet.End", "Boundary", "Edge.TR", 1#
    facade.Dimension "UpperRightFacet.Length", "UpperRightFacet.Start", "UpperRightFacet.End", "L", "Aligned", normalX, normalY
End Sub
