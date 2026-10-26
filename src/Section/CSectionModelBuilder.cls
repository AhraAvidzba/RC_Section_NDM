VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSectionModelBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Перенос готовой сетки и арматуры в расчётную модель
' ==========================================================================
' Соединяет бетонные элементы CFiberMeshBuilder и стержни CRebarLayout
' в одну CSectionModel. Имена расчётных элементов назначает сама модель.
' Построение геометрии, аннотаций и решение равновесия выполняются отдельно.

' Переносит готовые бетонные элементы с коэффициентом заполнения и собственными
' инерциями в CSectionModel, затем добавляет стержни. Модель сама назначает
' расчётные имена. Исходная сетка и арматура не изменяются; НДС здесь не решается.
Public Function BuildFromGenerated(ByVal mesh As CFiberMeshBuilder, ByVal rebars As CRebarLayout, _
        Optional ByVal sourceType As String = "Generated") As CSectionModel
    If mesh Is Nothing Then Err.Raise vbObjectError + 5250, "CSectionModelBuilder", "Сетка бетонных элементов не передана."
    If mesh.FiberCount <= 0 Then Err.Raise vbObjectError + 5251, "CSectionModelBuilder", "Сетка бетонных элементов пуста."

    Dim model As CSectionModel
    Set model = New CSectionModel
    model.SourceType = sourceType
    model.ConcreteMeshSource = "Generated"
    Dim rebarCount As Long
    If Not rebars Is Nothing Then rebarCount = rebars.Count
    model.ReserveElements mesh.FiberCount, rebarCount

    AppendConcreteMesh model, mesh

    Dim i As Long
    If Not rebars Is Nothing Then
        For i = 1 To rebars.Count
            model.AddRebarElement rebars.X(i), rebars.Y(i), rebars.Diameter(i), rebars.Area(i), _
                rebars.SteelClass(i), 1, rebars.BarID(i), vbNullString, rebars.Comment(i)
        Next i
    End If

    Set BuildFromGenerated = model
End Function

' Один перенос волокон для параметрической формы и формы по CAD-контуру.
' Уже импортированные стержни и достоверные границы модели не меняются.
Public Sub AppendConcreteMesh(ByVal model As CSectionModel, ByVal mesh As CFiberMeshBuilder, _
        Optional ByVal comment As String = "Параметрическая бетонная сетка")
    If model Is Nothing Or mesh Is Nothing Then Err.Raise vbObjectError + 5250, "CSectionModelBuilder", "Модель и бетонная сетка не переданы."
    If model.ConcreteCount > 0 Then Err.Raise vbObjectError + 5252, "CSectionModelBuilder", "Нельзя добавлять сгенерированную сетку поверх существующих бетонных элементов."
    If mesh.FiberCount <= 0 Then Err.Raise vbObjectError + 5251, "CSectionModelBuilder", "Сетка бетонных элементов пуста."
    model.ReserveElements mesh.FiberCount, model.RebarCount
    Dim i As Long
    For i = 1 To mesh.FiberCount
        Dim localIx As Double, localIy As Double, localIxy As Double, region As CConcreteRegion
        mesh.FiberLocalInertia i, localIx, localIy, localIxy
        Set region = mesh.FiberGeometry(i)
        Dim shapeType As String
        shapeType = "Rectangle"
        If Not region Is Nothing Then shapeType = "Contour"
        model.AddConcreteElement mesh.FiberX(i), mesh.FiberY(i), _
            mesh.FiberArea(i) * mesh.FiberFillFactor(i), mesh.FiberMaterialID(i), _
            "Бетонный КЭ " & CStr(i), vbNullString, shapeType, _
            mesh.FiberWidth(i), mesh.FiberHeight(i), 0#, comment, _
            localIx, localIy, localIxy
        If Not region Is Nothing Then model.SetConcreteGeometry model.ConcreteCount, region
    Next i
End Sub
