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

' ==========================================================================
' Сборка параметрического сечения по описанию рецепта
' ==========================================================================
' По готовому CSectionDefinition строит бетонную сетку, размещает арматуру
' и добавляет размеры и подписи в единую модель CSectionModel.
' Для всех типов сечения использует одни общие правила построения.
' Сохраняет сетку и данные расстановки последней успешной сборки:
' их можно прочитать для проверки без повторного построения сечения.

Private mMesh As CFiberMeshBuilder ' Готовая сетка последней сборки, доступная только для чтения.
Private mRebars As CRebarLayout ' Готовые стержни последней сборки, доступные только для чтения.
Private mRules As CRebarRuleExecutor ' Позиции стержней и диагностика последней успешной сборки.
Private mModel As CSectionModel ' Модель, для которой подготовлены выдаваемые данные.
Private mModelRevision As Long ' Изменение модели запрещает чтение уже не соответствующей ей сборки.

' Создаёт единую модель: проверяет описание и арматуру, строит сетку бетона
' и добавляет контуры, размеры и подписи сечения.
Public Function Build(ByVal definition As CSectionDefinition, ByRef meshOptions As TSectionMeshOptions) As CSectionModel
    Set mMesh = Nothing: Set mRebars = Nothing: Set mRules = Nothing: Set mModel = Nothing
    If definition Is Nothing Then SectionRecipeError "CSectionBuilder", "Не передано описание геометрии и армирования сечения."
    If Not definition.IsSealed Then SectionRecipeError "CSectionBuilder", "Описание сечения ещё не завершено. Перед сборкой завершите его методом Finish."
    If meshOptions.StepX <= 0# Or meshOptions.StepY <= 0# Or meshOptions.BoundarySubdivisions < 1 Or meshOptions.MaterialID < 1 Then _
        SectionRecipeError "CSectionBuilder", "Недопустимые внутренние настройки бетонной сетки."
    Dim geometry As CGeometryRegion, mesh As CFiberMeshBuilder, rebars As CRebarLayout, rules As CRebarRuleExecutor
    Dim modelBuilder As CSectionModelBuilder, annotationBuilder As CSectionAnnotationBuilder, model As CSectionModel
    Set geometry = New CGeometryRegion: geometry.Initialize definition.ConcreteDomain
    Set rules = New CRebarRuleExecutor: Set rebars = rules.Build(definition, geometry) ' Сначала проверяем арматуру общим исполнителем правил, затем строим бетонную сетку.
    If rebars.Count = 0 Then SectionRecipeError "CSectionBuilder", "После построения сечения не получен ни один стержень. Проверьте количества и диаметры в исходной таблице выбранного сечения."
    Set mesh = New CFiberMeshBuilder
    mesh.BuildMesh geometry, meshOptions.StepX, meshOptions.StepY, meshOptions.MaterialID, meshOptions.BoundarySubdivisions, meshOptions.SourceReference
    Set modelBuilder = New CSectionModelBuilder: Set model = modelBuilder.BuildFromGenerated(mesh, rebars)
    model.GeometryDiagnosticsText = rules.DiagnosticsText
    Dim i As Long
    For i = 1 To rules.PositionCount
        model.AppendRebarRecipeTrace rules.PositionPhysicalOwner(i), rules.PositionRuleID(i), rules.PositionPathID(i), rules.PositionSourceID(i)
    Next i
    definition.WriteContours model.Contours
    Set annotationBuilder = New CSectionAnnotationBuilder
    annotationBuilder.Build definition, rebars, model.Annotations, rules
    mesh.Freeze: rebars.Freeze: rules.Freeze
    Set mModel = model: mModelRevision = model.Revision
    Set mMesh = mesh: Set mRebars = rebars: Set mRules = rules
    Set Build = model
End Function

' Возвращает готовую бетонную сетку последней успешной сборки без её повторного построения.
Public Property Get Mesh() As CFiberMeshBuilder
    RequireCurrentModel
    If mMesh Is Nothing Then SectionRecipeError "CSectionBuilder", "Бетонная сетка ещё не была успешно построена."
    Set Mesh = mMesh
End Property

' Возвращает стержни последней успешной сборки с уже обработанными совпадениями позиций.
Public Property Get Rebars() As CRebarLayout
    RequireCurrentModel
    If mRebars Is Nothing Then SectionRecipeError "CSectionBuilder", "Расстановка арматуры ещё не была успешно построена."
    Set Rebars = mRebars
End Property

' Возвращает исполненные правила расстановки: количество, шаги, связи стержней и диагностику.
Public Property Get Rules() As CRebarRuleExecutor
    RequireCurrentModel
    If mRules Is Nothing Then SectionRecipeError "CSectionBuilder", "Правила размещения арматуры ещё не были успешно выполнены."
    Set Rules = mRules
End Property

' Проверяет, что сохранённые сетка, арматура и правила ещё соответствуют модели.
' После изменения её геометрии требует новую сборку; большие массивы не копирует.
Private Sub RequireCurrentModel()
    If mModel Is Nothing Then SectionRecipeError "CSectionBuilder", "Нет успешно собранной модели."
    If mModel.Revision <> mModelRevision Then SectionRecipeError "CSectionBuilder", "После сборки геометрия модели изменилась; прежние сетка, арматура и правила больше не соответствуют модели. Выполните новую сборку."
End Sub
