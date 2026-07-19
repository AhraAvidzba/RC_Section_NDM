Attribute VB_Name = "ExportAllClasses"
Sub ExportAllModules()
    Dim vbComp As VBIDE.VBComponent
    Dim exportPath As String

    ' Папка, куда будет сохранён экспорт
    exportPath = ThisWorkbook.Path & "\ExportedModules\"
    On Error Resume Next
    MkDir exportPath
    On Error GoTo 0

    Application.StatusBar = "Экспорт модулей..."

    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        Select Case vbComp.Type
            Case vbext_ct_ClassModule, vbext_ct_Document, vbext_ct_StdModule
                vbComp.Export exportPath & vbComp.Name & ".cls"
        End Select
    Next vbComp

    Application.StatusBar = False
    MsgBox "Экспорт завершён! Папка: " & exportPath, vbInformation
End Sub
