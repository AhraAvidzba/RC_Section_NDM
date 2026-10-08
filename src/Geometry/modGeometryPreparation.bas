Attribute VB_Name = "modGeometryPreparation"
Option Explicit

' Подготовка единой бетонной области из параметрического описания формы.
' После этого сетка и проверки арматуры не зависят от конкретного типа.
Public Function BuildConcreteGeometry(ByVal shape As ISectionShape) As CGeometryRegion
    If shape Is Nothing Then Err.Raise vbObjectError + 5290, "modGeometryPreparation", "Описание формы сечения не передано."
    Dim message As String
    If Not shape.IsValid(message) Then Err.Raise vbObjectError + 5290, "modGeometryPreparation", message
    Dim contours As CSectionContours, query As CSectionGeometryQuery
    Set contours = New CSectionContours
    shape.BuildContours contours
    Set query = New CSectionGeometryQuery
    Set BuildConcreteGeometry = New CGeometryRegion
    BuildConcreteGeometry.Initialize query.RegionFromContours(contours)
End Function
