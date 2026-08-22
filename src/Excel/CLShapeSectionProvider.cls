VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLShapeSectionProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Provider Г-образного сечения
' ==========================================================================
' Provider связывает CGeometryLShape, CLShapeRebarLayoutBuilder и
' CLShapeAnnotationBuilder. Это единственное место, где универсальный поток
' получает набор компонентов именно для LShape.

Implements ISectionTypeProvider

Private Function ISectionTypeProvider_TypeName() As String
    ISectionTypeProvider_TypeName = "LShape"
End Function

Private Function ISectionTypeProvider_BuildGeometry(ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem) As ISectionGeometry
    Dim geom As CGeometryLShape
    Set geom = New CGeometryLShape
    geom.Initialize units.InputLengthToInternal(settings.GetDouble("LShape.B1", 250#)), _
        units.InputLengthToInternal(settings.GetDouble("LShape.H1", 550#)), _
        units.InputLengthToInternal(settings.GetDouble("LShape.B2", 600#)), _
        units.InputLengthToInternal(settings.GetDouble("LShape.H2", 250#)), 0#, 0#
    Set ISectionTypeProvider_BuildGeometry = geom
End Function

Private Function ISectionTypeProvider_BuildRebars(ByVal geometry As ISectionGeometry, ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem) As CRebarLayout
    Dim builder As CLShapeRebarLayoutBuilder
    Set builder = New CLShapeRebarLayoutBuilder
    Set ISectionTypeProvider_BuildRebars = builder.Build( _
        units.InputLengthToInternal(settings.GetDouble("LShape.B1", 250#)), _
        units.InputLengthToInternal(settings.GetDouble("LShape.H1", 550#)), _
        units.InputLengthToInternal(settings.GetDouble("LShape.B2", 600#)), _
        units.InputLengthToInternal(settings.GetDouble("LShape.H2", 250#)), _
        0#, _
        0#, _
        LShapeFaceSettings(settings, units, "H1"), _
        LShapeFaceSettings(settings, units, "H2"), _
        LShapeFaceSettings(settings, units, "B1"), _
        LShapeFaceSettings(settings, units, "B2"), _
        settings.GetString("Steel.RebarProfile", "Ribbed"))
End Function

Private Sub ISectionTypeProvider_BuildAnnotations(ByVal model As CSectionModel, ByVal geometry As ISectionGeometry, ByVal rebars As CRebarLayout)
    Dim builder As CLShapeAnnotationBuilder
    Set builder = New CLShapeAnnotationBuilder
    builder.Build model, geometry, rebars
End Sub

Private Function LShapeFaceSettings(ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem, ByVal faceName As String) As Variant
    LShapeFaceSettings = Array( _
        units.InputLengthToInternal(settings.GetDouble("LShape." & faceName & ".as_1", 40#)), _
        units.InputLengthToInternal(settings.GetDouble("LShape." & faceName & ".as_2", 40#)), _
        units.InputLengthToInternal(settings.GetDouble("LShape." & faceName & ".d_1", 32#)), _
        units.InputLengthToInternal(settings.GetDouble("LShape." & faceName & ".d_2", 32#)), _
        settings.GetLong("LShape." & faceName & ".n_1", 0), _
        settings.GetLong("LShape." & faceName & ".n_2", 0), _
        units.InputLengthToInternal(settings.GetDouble("LShape." & faceName & ".StartOffset1", 80#)), _
        units.InputLengthToInternal(settings.GetDouble("LShape." & faceName & ".EndOffset1", 80#)), _
        units.InputLengthToInternal(settings.GetDouble("LShape." & faceName & ".StartOffset2", 80#)), _
        units.InputLengthToInternal(settings.GetDouble("LShape." & faceName & ".EndOffset2", 80#)), _
        units.InputLengthToInternal(settings.GetDouble("LShape." & faceName & ".d_2row_1", 0#)), _
        units.InputLengthToInternal(settings.GetDouble("LShape." & faceName & ".d_2row_2", 0#)), _
        units.InputLengthToInternal(settings.GetDouble("LShape." & faceName & ".d_3row_1", 0#)), _
        units.InputLengthToInternal(settings.GetDouble("LShape." & faceName & ".d_3row_2", 0#)), _
        settings.GetString("LShape." & faceName & ".loc_2row", "Stacked"), _
        settings.GetString("LShape." & faceName & ".loc_3row", "Stacked"), _
        settings.GetString("LShape." & faceName & ".bind_2row", "EachBar"), _
        settings.GetString("LShape." & faceName & ".bind_3row", "EachBar"))
End Function

