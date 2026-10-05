Attribute VB_Name = "modAutoCADContourOwnership"
Option Explicit

' ==========================================================================
' Владение экспортированными контурами на общих слоях AutoCAD
' ==========================================================================
' Общие слои содержат как исходные контуры пользователя, так и вывод NDM.
' Метка XData отличает собственный вывод для безопасной очистки и исключения
' дублей при импорте. Модуль не читает Excel и не интерпретирует геометрию.

Private Const CONTOUR_APP As String = "RC_SECTION_NDM"
Private Const CONTOUR_MARK As String = "MATERIAL_CONTOUR_OUTPUT"

' Помечает только что созданный контур. Без метки объект удаляется сразу:
' иначе следующий экспорт не сможет отличить его от исходника пользователя.
Public Sub MarkNDMContourOutput(ByVal entity As Object)
    On Error GoTo Failed
    Dim registered As Object
    On Error Resume Next
    Set registered = entity.Document.RegisteredApplications.Item(CONTOUR_APP)
    On Error GoTo Failed
    If registered Is Nothing Then Set registered = entity.Document.RegisteredApplications.Add(CONTOUR_APP)
    Dim codes(0 To 1) As Integer, values(0 To 1) As Variant
    codes(0) = 1001: values(0) = CONTOUR_APP
    codes(1) = 1000: values(1) = CONTOUR_MARK
    entity.SetXData codes, values
    Exit Sub
Failed:
    Dim number As Long, description As String
    number = Err.Number: description = Err.Description
    On Error Resume Next
    entity.Delete
    On Error GoTo 0
    Err.Raise number, "MarkNDMContourOutput", "Не удалось пометить экспортированный контур AutoCAD: " & description
End Sub

' Проверяет собственную XData, не используя имя слоя как доказательство
' владения. Исходные объекты без метки никогда не считаются выводом NDM.
Public Function IsNDMContourOutput(ByVal entity As Object) As Boolean
    On Error GoTo NotOwned
    Dim codes As Variant, values As Variant, first As Long
    entity.GetXData CONTOUR_APP, codes, values
    If Not IsArray(codes) Or Not IsArray(values) Then Exit Function
    first = LBound(values)
    If UBound(values) < first + 1 Then Exit Function
    IsNDMContourOutput = (CLng(codes(first)) = 1001 And CStr(values(first)) = CONTOUR_APP And _
        CLng(codes(first + 1)) = 1000 And CStr(values(first + 1)) = CONTOUR_MARK)
NotOwned:
End Function
