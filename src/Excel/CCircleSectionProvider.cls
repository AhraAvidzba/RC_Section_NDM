VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCircleSectionProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Provider круглого сечения
' ==========================================================================
' Provider собирает специализированные компоненты для Circle: геометрию,
' раскладку арматуры и semantic-аннотации. Универсальный workbook-код
' обращается к нему через ISectionTypeProvider и не содержит условий вида
' If GeometryType = Circle.

Implements ISectionTypeProvider

Private Function ISectionTypeProvider_TypeName() As String
    ISectionTypeProvider_TypeName = "Circle"
End Function

Private Function ISectionTypeProvider_BuildGeometry(ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem) As ISectionGeometry
    Dim geom As CGeometryCircle
    Set geom = New CGeometryCircle
    geom.InitializeByDiameter units.InputLengthToInternal(settings.GetDouble("Circle.Diameter", 300#)), 0#, 0#
    Set ISectionTypeProvider_BuildGeometry = geom
End Function

Private Function ISectionTypeProvider_BuildRebars(ByVal geometry As ISectionGeometry, ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem) As CRebarLayout
    Dim builder As CCircleRebarLayoutBuilder
    Set builder = New CCircleRebarLayoutBuilder
    Set ISectionTypeProvider_BuildRebars = builder.Build( _
        units.InputLengthToInternal(settings.GetDouble("Circle.Diameter", 300#)), _
        0#, _
        0#, _
        units.InputLengthToInternal(settings.GetDouble("Rebar.AxisDistance", 40#)), _
        settings.GetLong("Rebar.Count", 8), _
        units.InputLengthToInternal(settings.GetDouble("Rebar.Diameter", 20#)), _
        settings.GetString("Steel.RebarProfile", "Ribbed"), _
        units.InputLengthToInternal(settings.GetDouble("Rebar.Diameter2", 0#)), _
        units.InputLengthToInternal(settings.GetDouble("Rebar.Diameter3", 0#)), _
        settings.GetString("Rebar.Loc2row", "Stacked"), _
        settings.GetString("Rebar.Loc3row", "Stacked"))
End Function

Private Sub ISectionTypeProvider_BuildAnnotations(ByVal model As CSectionModel, ByVal geometry As ISectionGeometry, ByVal rebars As CRebarLayout)
    Dim builder As CCircleAnnotationBuilder
    Set builder = New CCircleAnnotationBuilder
    builder.Build model, geometry, rebars
End Sub

