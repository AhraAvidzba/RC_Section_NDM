VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSectionTypeRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Реестр поддерживаемых типов сечений
' ==========================================================================
' Registry - единственная точка, где строка Geometry.Type из Config
' сопоставляется с конкретным ISectionTypeProvider. Остальные слои не должны
' угадывать форму сечения и не должны содержать shape-specific Select Case.

Public Function CreateProvider(ByVal geometryType As String) As ISectionTypeProvider
    Select Case LCase$(Trim$(geometryType))
        Case "circle"
            Dim circleProvider As CCircleSectionProvider
            Set circleProvider = New CCircleSectionProvider
            Set CreateProvider = circleProvider
        Case "lshape"
            Dim lshapeProvider As CLShapeSectionProvider
            Set lshapeProvider = New CLShapeSectionProvider
            Set CreateProvider = lshapeProvider
        Case "roundedrectangle"
            Dim roundedProvider As CRoundedRectSectionProvider
            Set roundedProvider = New CRoundedRectSectionProvider
            Set CreateProvider = roundedProvider
        Case Else
            Err.Raise vbObjectError + 4134, "CSectionTypeRegistry", _
                "Unsupported section geometry type: " & geometryType
    End Select
End Function

Public Function ActiveProvider(ByVal settings As CSystemSettingsReader) As ISectionTypeProvider
    If settings Is Nothing Then Err.Raise vbObjectError + 4135, _
        "CSectionTypeRegistry", "Настройки Config не переданы."

    Dim geometryType As String
    geometryType = settings.GetString("Geometry.Type", "LShape")
    If Len(Trim$(geometryType)) = 0 Then geometryType = "LShape"
    Set ActiveProvider = CreateProvider(geometryType)
End Function

' Создает расчетный или интерфейсный объект из нормализованных исходных данных и локальных настроек.
Public Function CreateGeometry(ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem) As ISectionGeometry
    Dim provider As ISectionTypeProvider
    Set provider = ActiveProvider(settings)
    Set CreateGeometry = provider.BuildGeometry(settings, units)
End Function

' Создает расчетный или интерфейсный объект из нормализованных исходных данных и локальных настроек.
Public Function CreateRebars(ByVal geometry As ISectionGeometry, ByVal settings As CSystemSettingsReader, _
        ByVal units As CUnitSystem) As CRebarLayout
    Dim provider As ISectionTypeProvider
    Set provider = ActiveProvider(settings)
    Set CreateRebars = provider.BuildRebars(geometry, settings, units)
End Function

' Создает расчетный или интерфейсный объект из нормализованных исходных данных и локальных настроек.
Public Function BuildGeneratedModel(ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem) As CSectionModel
    Dim provider As ISectionTypeProvider
    Set provider = ActiveProvider(settings)

    Dim geometry As ISectionGeometry
    Set geometry = provider.BuildGeometry(settings, units)

    Dim mesh As CFiberMeshBuilder
    Set mesh = New CFiberMeshBuilder

    Dim meshStep As Double
    meshStep = units.InputLengthToInternal(settings.GetDouble("Mesh.Step", 10#))
    mesh.BuildMesh geometry, meshStep, meshStep, 1, MeshBoundarySubdivisions(settings)

    Dim rebars As CRebarLayout
    Set rebars = provider.BuildRebars(geometry, settings, units)

    Dim model As CSectionModel
    Set model = BuildGeneratedSectionModel(mesh, rebars, "Generated")
    provider.BuildAnnotations model, geometry, rebars
    Set BuildGeneratedModel = model
End Function

Public Function MeshBoundarySubdivisions(ByVal settings As CSystemSettingsReader) As Long
    MeshBoundarySubdivisions = settings.GetLong("Mesh.BoundarySubdivisions", 1)
    If MeshBoundarySubdivisions < 1 Then MeshBoundarySubdivisions = 1
End Function

