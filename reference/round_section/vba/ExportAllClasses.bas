Attribute VB_Name = "ExportAllClasses"
Sub ExportAllModules()
    Dim vbComp As VBIDE.VBComponent
    Dim folderPath As String
    Dim ext As String
    
    ' ѕуть дл€ сохранени€
    folderPath = ThisWorkbook.Path & "\VBA_Export\"
    
    ' —оздаЄм папку
    On Error Resume Next
    MkDir folderPath
    On Error GoTo 0
    
    ' ѕеребираем все компоненты
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        ' Ёкспортируем модули, классы и формы
        If vbComp.Type = vbext_ct_StdModule Or _
           vbComp.Type = vbext_ct_ClassModule Or _
           vbComp.Type = vbext_ct_MSForm Then
            
            ' ¬ыбираем расширение
            If vbComp.Type = vbext_ct_MSForm Then
                ext = ".frm"
            Else
                ext = ".bas"
            End If
            
            ' Ёкспорт
            vbComp.Export folderPath & vbComp.Name & ext
        End If
    Next vbComp
    
    MsgBox "Ёкспортировано! ѕапка: " & folderPath
End Sub

