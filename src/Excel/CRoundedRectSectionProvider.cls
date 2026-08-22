VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CRoundedRectSectionProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Provider прямоугольного сечения со скруглениями
' ==========================================================================
' Provider создает компоненты RoundedRectangle и скрывает их от общей
' workbook-логики. Благодаря этому добавление новой формы должно сводиться
' к новому provider-у и регистрации в CSectionTypeRegistry.

Implements ISectionTypeProvider

Private Function ISectionTypeProvider_TypeName() As String
    ISectionTypeProvider_TypeName = "RoundedRectangle"
End Function

Private Function ISectionTypeProvider_BuildGeometry(ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem) As ISectionGeometry
    Dim geom As CGeometryRoundedRectangle
    Set geom = New CGeometryRoundedRectangle
    geom.Initialize units.InputLengthToInternal(settings.GetDouble("RoundedRectangle.Width", 300#)), _
        units.InputLengthToInternal(settings.GetDouble("RoundedRectangle.Height", 200#)), _
        units.InputLengthToInternal(settings.GetDouble("RoundedRectangle.RadiusTopLeft", 0#)), _
        units.InputLengthToInternal(settings.GetDouble("RoundedRectangle.RadiusTopRight", 0#)), _
        units.InputLengthToInternal(settings.GetDouble("RoundedRectangle.RadiusBottomRight", 0#)), _
        units.InputLengthToInternal(settings.GetDouble("RoundedRectangle.RadiusBottomLeft", 0#))
    Set ISectionTypeProvider_BuildGeometry = geom
End Function

Private Function ISectionTypeProvider_BuildRebars(ByVal geometry As ISectionGeometry, ByVal settings As CSystemSettingsReader, _
        ByVal units As CUnitSystem) As CRebarLayout
    Err.Raise vbObjectError + 4103, "CRoundedRectSectionProvider", _
        "Automatic rebar layout is supported only for Circle and LShape."
End Function

Private Sub ISectionTypeProvider_BuildAnnotations(ByVal model As CSectionModel, ByVal geometry As ISectionGeometry, _
        ByVal rebars As CRebarLayout)
    Dim builder As CRoundedRectAnnotationBuilder
    Set builder = New CRoundedRectAnnotationBuilder
    builder.Build model, geometry, rebars
End Sub

