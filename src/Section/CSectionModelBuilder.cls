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
' Сборка CSectionModel из сетки бетона и раскладки арматуры
' ==========================================================================
' Builder соединяет результат геометрической дискретизации и арматурный layout
' в один CSectionModel. Это тонкий адаптер: он не решает равновесие, не читает
' Excel и не создает semantic-аннотации конкретной формы сам по себе.

' Переносит сгенерированные волокна с коэффициентом заполнения и их местными
' инерциями в единый расчетный model, затем добавляет готовую арматуру.
' Новая модель получает собственные расчетные имена; builder не меняет сетку
' или layout и не выполняет расчет НДС.
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
        model.AddConcreteElement mesh.FiberX(i), mesh.FiberY(i), _
            mesh.FiberArea(i) * mesh.FiberFillFactor(i), mesh.FiberMaterialID(i), _
            "Бетонный КЭ " & CStr(i), vbNullString, "Rectangle", _
            mesh.FiberWidth(i), mesh.FiberHeight(i), 0#, comment, _
            mesh.FiberArea(i) * mesh.FiberFillFactor(i) * mesh.FiberHeight(i) * mesh.FiberHeight(i) / 12#, _
            mesh.FiberArea(i) * mesh.FiberFillFactor(i) * mesh.FiberWidth(i) * mesh.FiberWidth(i) / 12#, 0#
    Next i
End Sub

