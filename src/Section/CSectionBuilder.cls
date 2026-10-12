VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSectionBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Единственная сборка definition -> область -> сетка/арматура -> модель.
' Тип сечения не анализируется. Источники mesh/layout доступны для проверок
' из этой же сборки, без повторного вызова рецепта или старых генераторов.
Private mMesh As CFiberMeshBuilder
Private mRebars As CRebarLayout
Private mRules As CRebarRuleExecutor

Public Function Build(ByVal definition As CSectionDefinition, ByRef meshOptions As TSectionMeshOptions) As CSectionModel
    Set mMesh = Nothing: Set mRebars = Nothing: Set mRules = Nothing
    If definition Is Nothing Then SectionRecipeError "CSectionBuilder", "Определение не передано."
    If Not definition.IsSealed Then SectionRecipeError "CSectionBuilder", "Нельзя собирать неопубликованное определение."
    If meshOptions.StepX <= 0# Or meshOptions.StepY <= 0# Or meshOptions.BoundarySubdivisions < 1 Or meshOptions.MaterialID < 1 Then _
        SectionRecipeError "CSectionBuilder", "Недопустимые внутренние настройки бетонной сетки."
    Dim geometry As CGeometryRegion, mesh As CFiberMeshBuilder, rebars As CRebarLayout, rules As CRebarRuleExecutor
    Dim modelBuilder As CSectionModelBuilder, annotationBuilder As CSectionAnnotationBuilder, model As CSectionModel
    Set geometry = New CGeometryRegion: geometry.Initialize definition.ConcreteDomain
    Set mesh = New CFiberMeshBuilder
    mesh.BuildMesh geometry, meshOptions.StepX, meshOptions.StepY, meshOptions.MaterialID, meshOptions.BoundarySubdivisions
    Set rules = New CRebarRuleExecutor: Set rebars = rules.Build(definition, geometry)
    If rebars.Count = 0 Then SectionRecipeError "CSectionBuilder", "Для параметрического сечения не задан ни один стержень арматуры."
    Set modelBuilder = New CSectionModelBuilder: Set model = modelBuilder.BuildFromGenerated(mesh, rebars)
    model.GeometryDiagnosticsText = rules.DiagnosticsText
    Dim i As Long
    For i = 1 To rules.PositionCount
        model.AppendRebarRecipeTrace rules.PositionPhysicalOwner(i), rules.PositionRuleID(i), rules.PositionPathID(i), rules.PositionSourceID(i)
    Next i
    definition.WriteContours model.Contours
    Set annotationBuilder = New CSectionAnnotationBuilder
    annotationBuilder.Build definition, rebars, model.Annotations, rules
    Set mMesh = mesh: Set mRebars = rebars: Set mRules = rules
    Set Build = model
End Function

Public Property Get Mesh() As CFiberMeshBuilder
    If mMesh Is Nothing Then SectionRecipeError "CSectionBuilder", "Нет успешно собранной сетки."
    Set Mesh = mMesh
End Property

Public Property Get Rebars() As CRebarLayout
    If mRebars Is Nothing Then SectionRecipeError "CSectionBuilder", "Нет успешно собранной арматуры."
    Set Rebars = mRebars
End Property

Public Property Get Rules() As CRebarRuleExecutor
    If mRules Is Nothing Then SectionRecipeError "CSectionBuilder", "Нет успешно исполненных правил."
    Set Rules = mRules
End Property
